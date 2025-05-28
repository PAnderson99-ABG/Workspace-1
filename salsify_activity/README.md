# Salsify Activity Reporting Tool

This Python tool consolidates quarterly Salsify export files, analyzes changes by property, user, and brand, and generates an Excel report with charts and breakdowns.

---

## 📦 Features

- Combines and normalizes multiple quarterly exports
- Validates and parses timestamps
- Generates summary tables:
  - Top 50 changed properties
  - Top users
  - Top 50 brands
- Creates breakdown tables by brand and property
- Inserts visual charts (bar, line)
- Monthly brand breakdown sheets

---

## 🛠 Setup

1. **Clone the repository**

2. **Create and activate a virtual environment**

```bash
python -m venv .venv
.venv\Scripts\activate  # On Windows
```

3. **Install dependencies**

```bash
pip install -r requirements.txt
```

4. **Create a `.env` file** with paths to your CSV exports:

```
WCWQ1P1=path/to/Q1P1.csv
WCWQ1P2=path/to/Q1P2.csv
WCWQ2P1=path/to/Q2P1.csv
WCWQ2P2=path/to/Q2P2.csv
```

---

## 🚀 Run It

```bash
python main.py
```

The script will output:  
✅ `report.xlsx` in the working directory

---

## 🗂 Directory Structure

```
salsify_activity/
├── main.py
├── .env
├── report.xlsx
├── modules/
│   ├── __init__.py
│   ├── load_data.py
│   ├── summarize.py
│   ├── write_helpers.py
│   ├── charts.py
│   ├── monthly_breakdown.py
│   └── report_builder.py
```

---

## 🧩 Module Overview

### `modules/load_data.py`
Handles loading, normalizing, and validating CSV data.
- `load_and_prepare()`
- `load_data()`

### `modules/summarize.py`
Creates summary tables and grouped user breakdowns.
- `generate_summaries()`
- `generate_breakdowns()`

### `modules/write_helpers.py`
Writes grouped DataFrames to Excel sheets.
- `write_breakdown_table()`

### `modules/charts.py`
Inserts visualizations into Excel (bar and line charts).
- `insert_summary_chart()`
- `insert_changes_over_time_sheet()`

### `modules/monthly_breakdown.py`
Generates monthly sheets with top brands and users.
- `generate_monthly_brand_breakdowns()`

### `modules/report_builder.py`
Coordinates Excel output.
- `build_excel_report()`

### `main.py`
Main entry point — calls each module in order:
1. Load data
2. Generate summaries & breakdowns
3. Build Excel report

---

## 📋 Requirements

- Python 3.9+
- `pandas`
- `xlsxwriter`
- `python-dotenv`

---

## 📎 License

MIT (or private/internal use)
