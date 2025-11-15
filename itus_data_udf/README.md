# Daily Data UDFs for Excel

[![Python](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/)  
[![License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)

Excel User Defined Functions (UDFs) to fetch financial data directly from a local SQLite database.  
This project allows Excel users to retrieve metrics such as **PE ratios, sector information, and market capitalization categories** for companies using Python and `xlwings`.

---

## 📌 Project Overview

At ITUS Capital, financial datasets are securely maintained in databases.  
This project bridges **Excel** and a **local SQLite database**, enabling users to:

- Query single or multiple company metrics
- Retrieve historical time series data
- Fetch full datasets in a matrix format
- Maintain logs and monitor query performance

All without exposing sensitive database credentials.

---

## ⚡ Features

| Function | Description | Example Excel Formula |
|----------|-------------|---------------------|
| `get_daily_data` | Fetch a single metric for a company on a specific date | `=get_daily_data(100186, "pe", "2025-11-08")` |
| `get_series` | Retrieve a time series of a metric for a company | `=get_series(100186, "pe", "2025-10-01", "2025-11-08")` |
| `get_daily_matrix` | Fetch all companies' data for a specific date | `=get_daily_matrix("2025-11-08", "pe")` |
| `get_all_pe` | Fetch all historical data for a company | `=get_all_pe(100186, "pe")` |
| `get_mcap_matrix` | Fetch all companies of a market cap on a date | `=get_mcap_matrix("Large Cap", "2025-11-08")` |
| `get_pe_for_sector` | Fetch all companies in a sector on a date | `=get_pe_for_sector("Technology", "2025-11-08")` |
| `clear_cache` | Clear cached queries for updated data | `=clear_cache()` |

**Additional Features:**

- Spillable outputs for Excel tables
- LRU caching for performance
- Detailed query logging with execution times
- Automatic error handling for missing data or invalid inputs

---

## 📂 Project Structure

itus_data_udf/
│
├─ daily_data_udf.py # Python code containing all UDFs
├─ valuations.db # Sample SQLite database
├─ config.ini # Configuration file for DB path, table, and date format
├─ schema.sql # Index creation script
├─ example.xlsx # Example Excel workbook demonstrating UDF usage
├─ query_log.txt # Generated log file
└─ README.md # Project documentation

---

## ⚙️ Setup Instructions

### 1. Install Dependencies

**Python Package Installation:**

```bash
pip install xlwings pandas

## Install xlwings Excel Add-in:

xlwings addin install

## Verify installation:

pip show xlwings

```

2. Enable Trust Access to VBA Project

To allow xlwings to control Excel macros:

Open Excel → File → Options → Trust Center → Trust Center Settings → Macro Settings

Check "Trust access to the VBA project object model"

Click OK to apply

3. Configure the Database

Edit config.ini with your SQLite path and table name:

[DATABASE]
db_path = valuations.db
table_name = valuations

[FORMAT]
date_format = %Y-%m-%d

## Create an index for performance:

```bash

sqlite3 valuations.db < schema.sql

```

4. Add xlwings Add-in to Excel

Open example.xlsm or any macro-enabled Excel file.

Go to File → Options → Add-ins

In Manage → Excel Add-ins → Go → Browse, select xlwings.xlam

Ensure the xlwings checkbox is checked

5. Enable xlwings Reference in VBA Editor

Open VBA Editor: Alt + F11

Go to Tools → References

Check xlwings (or browse to .xlam file if missing)

Click OK

6. Configure xlwings Settings in Excel

Go to the xlwings tab:

Set Python Interpreter: Verify Python 3.10+ path

Set Python Path: Add your project folder path

Set UDF Module: Enter daily_data_udf (without .py)

Click Import Functions to load UDFs into Excel

After this, the Excel functions listed above are available for direct use.

📝 Logging & Performance

Logs written to query_log.txt with timestamps, function names, parameters, execution time, and status

LRU caching ensures repeated queries execute efficiently

Indexed tables allow queries to respond in < 0.05s for large datasets

### Example Excel Formulas

=get_daily_data(100186, "pe", "2025-11-08")
=get_series(100186, "pe", "2025-10-01", "2025-11-08")
=get_daily_matrix("2025-11-08", "pe")
=get_all_pe(100186, "pe")
=get_mcap_matrix("Large Cap", "2025-11-08")
=get_pe_for_sector("Technology", "2025-11-08")
=clear_cache()
