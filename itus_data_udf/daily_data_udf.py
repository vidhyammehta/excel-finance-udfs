"""
daily_data_udf.py

Excel UDFs to fetch financial data from a local SQLite database.
Supports single-value lookups, series, matrices, and full company history.
"""

import xlwings as xw
import sqlite3
import pandas as pd
import configparser
import os
from datetime import datetime
from functools import lru_cache, wraps
import logging
from logging.handlers import RotatingFileHandler
import time

# -------------------------------------------------------------------
# CONFIGURATION
# -------------------------------------------------------------------
CONFIG_FILE = os.path.join(os.path.dirname(__file__), 'config.ini')
LOG_FILE = os.path.join(os.path.dirname(__file__), 'query_log.txt')

config = configparser.ConfigParser(interpolation=None)
config.read(CONFIG_FILE)

DB_PATH = config.get('DATABASE', 'db_path', fallback='valuations.db')
TABLE_NAME = config.get('DATABASE', 'table_name', fallback='valuations')
DATE_FORMAT = config.get('FORMAT', 'date_format', fallback='%Y-%m-%d')

# -------------------------------------------------------------------
# LOGGING
# -------------------------------------------------------------------
logger = logging.getLogger("DailyDataLogger")
logger.setLevel(logging.INFO)
if not logger.handlers:  # Prevents duplicate handlers in Excel context
    handler = RotatingFileHandler(LOG_FILE, maxBytes=1_000_000, backupCount=3, encoding='utf-8')
    formatter = logging.Formatter("%(asctime)s | %(levelname)s | %(message)s", "%Y-%m-%d %H:%M:%S")
    handler.setFormatter(formatter)
    logger.addHandler(handler)

# -------------------------------------------------------------------
# DATABASE CONNECTION
# -------------------------------------------------------------------
def _get_connection():
    if not os.path.exists(DB_PATH):
        raise FileNotFoundError(f"SQLite DB not found at path: {DB_PATH}")
    return sqlite3.connect(DB_PATH)

# -------------------------------------------------------------------
# CACHING AND QUERY EXECUTION
# -------------------------------------------------------------------
@lru_cache(maxsize=128)
def _cached_query(sql: str, params: tuple):
    conn = _get_connection()
    try:
        df = pd.read_sql_query(sql, conn, params=params)
        return df
    finally:
        conn.close()

def _run_query_df(sql: str, params: tuple = ()):
    start = time.perf_counter()
    df = _cached_query(sql, params)
    duration_ms = round((time.perf_counter() - start) * 1000, 3)
    logger.info(f"SQL executed | Duration={duration_ms} ms | Params={params}")
    return df

# -------------------------------------------------------------------
# INPUT VALIDATION & DATE FORMATTING
# -------------------------------------------------------------------
def _validate_inputs(**kwargs):
    for k, v in kwargs.items():
        if v is None or str(v).strip() == "":
            raise ValueError(f"Missing required input: {k}")

def _format_date_for_db(date_value: str) -> str:
    """Convert Excel input date to DB-compatible text (YYYY-MM-DD)."""
    try:
        dt = datetime.strptime(str(date_value).strip(), "%Y-%m-%d")
        return dt.strftime("%Y-%m-%d")
    except Exception:
        return str(date_value).strip()

def _format_date_for_excel(date_value: str) -> str:
    """Format date for Excel output according to config.ini."""
    try:
        dt = datetime.strptime(str(date_value).strip(), "%Y-%m-%d")
        return dt.strftime(DATE_FORMAT)
    except Exception:
        return str(date_value).strip()

# -------------------------------------------------------------------
# LOGGING DECORATOR
# -------------------------------------------------------------------
def log_call(func):
    @wraps(func)
    def wrapper(*args):
        start_time = time.perf_counter()
        status = "SUCCESS"
        error_msg = None
        try:
            result = func(*args)
            return result
        except Exception as e:
            status = "FAILED"
            error_msg = str(e)
            raise
        finally:
            duration_ms = round((time.perf_counter() - start_time) * 1000, 3)
            params_str = ", ".join([repr(a) for a in args])
            msg = f"Function={func.__name__} | Params=({params_str}) | Time={duration_ms} ms | Status={status}"
            if error_msg:
                msg += f" | Error='{error_msg}'"
            logger.info(msg)
    return wrapper

# -------------------------------------------------------------------
# UDFS
# -------------------------------------------------------------------
@xw.func(category="Finance UDFs")
@xw.ret(expand='table')
@log_call
def get_daily_data(accord_code, field: str, date_value: str):
    _validate_inputs(accord_code=accord_code, field=field, date_value=date_value)
    accord_code = int(float(accord_code))
    formatted_date = _format_date_for_db(date_value)
    sql = f"SELECT {field} FROM {TABLE_NAME} WHERE accord_code=? AND date=?"
    df = _run_query_df(sql, (accord_code, formatted_date))
    if df.empty:
        return [[f"No data found for {accord_code} on {formatted_date}"]]
    # Always return as a table, not just value, for Excel expand compatibility
    return [[df.iloc[0, 0]]]

@xw.func(category="Finance UDFs")
@xw.ret(expand='table')
@log_call
def get_series(accord_code, field: str, start_date: str, end_date: str):
    _validate_inputs(accord_code=accord_code, field=field, start_date=start_date, end_date=end_date)
    accord_code = int(float(accord_code))
    start_fmt = _format_date_for_db(start_date)
    end_fmt = _format_date_for_db(end_date)
    sql = f"SELECT date, {field} FROM {TABLE_NAME} WHERE accord_code=? AND date BETWEEN ? AND ? ORDER BY date"
    df = _run_query_df(sql, (accord_code, start_fmt, end_fmt))
    if df.empty:
        return [[f"No data found for {accord_code} between {start_fmt} and {end_fmt}"]]
    df['date'] = df['date'].apply(_format_date_for_excel)
    result = [df.columns.tolist()] + df.values.tolist()
    return result

@xw.func(category="Finance UDFs")
@xw.ret(expand='table')
@log_call
def get_daily_matrix(date_value: str, field: str):
    _validate_inputs(date_value=date_value, field=field)
    formatted_date = _format_date_for_db(date_value)
    sql = f"SELECT accord_code, company_name, sector, mcap_category, {field} FROM {TABLE_NAME} WHERE date=? ORDER BY accord_code"
    df = _run_query_df(sql, (formatted_date,))
    if df.empty:
        return [[f"No data found for {formatted_date}"]]
    # If date column is present add formatted date, else omit
    if 'date' in df.columns:
        df['date'] = df['date'].apply(_format_date_for_excel)
    result = [df.columns.tolist()] + df.values.tolist()
    return result

@xw.func(category="Finance UDFs")
@xw.ret(expand='table')
@log_call
def get_all_pe(accord_code, field: str):
    _validate_inputs(accord_code=accord_code, field=field)
    accord_code = int(float(accord_code))
    sql = f"SELECT date, {field} FROM {TABLE_NAME} WHERE accord_code=? ORDER BY date"
    df = _run_query_df(sql, (accord_code,))
    if df.empty:
        return [[f"No data found for {accord_code}"]]
    df['date'] = df['date'].apply(_format_date_for_excel)
    result = [df.columns.tolist()] + df.values.tolist()
    return result

@xw.func(category="Finance UDFs")
@xw.ret(expand='table')
@log_call
def get_mcap_matrix(mcap_category: str, date_value: str):
    _validate_inputs(mcap_category=mcap_category, date_value=date_value)
    formatted_date = _format_date_for_db(date_value)
    sql = f"SELECT accord_code, company_name, sector, pe FROM {TABLE_NAME} WHERE mcap_category=? AND date=? ORDER BY pe DESC"
    df = _run_query_df(sql, (mcap_category, formatted_date))
    if df.empty:
        return [[f"No data found for {mcap_category} on {formatted_date}"]]
    df['date'] = formatted_date
    result = [df.columns.tolist()] + df.values.tolist()
    return result

@xw.func(category="Finance UDFs")
@xw.ret(expand='table')
@log_call
def get_pe_for_sector(sector: str, date_value: str):
    _validate_inputs(sector=sector, date_value=date_value)
    formatted_date = _format_date_for_db(date_value)
    sql = f"SELECT accord_code, company_name, mcap_category, pe FROM {TABLE_NAME} WHERE sector=? AND date=? ORDER BY pe DESC"
    df = _run_query_df(sql, (sector, formatted_date))
    if df.empty:
        return [[f"No data found for sector {sector} on {formatted_date}"]]
    df['date'] = formatted_date
    result = [df.columns.tolist()] + df.values.tolist()
    return result

# -------------------------------------------------------------------
# CACHE CLEAR
# -------------------------------------------------------------------
@xw.func
@log_call
def clear_cache():
    _cached_query.cache_clear()
    return "Cache cleared successfully."

# -------------------------------------------------------------------
# TEST FUNCTION
# -------------------------------------------------------------------
@xw.func
def test_add(x, y):
    return x + y
