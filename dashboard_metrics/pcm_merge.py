import os
import smartsheet
import pandas as pd
from dotenv import load_dotenv
from datetime import datetime

# === Load credentials and sheet IDs from .env ===
load_dotenv()
ACCESS_TOKEN = os.getenv("SMARTSHEET_API_TOKEN")
SHEET_ID_1 = int(os.getenv("PCM_PROCESSING_ID"))  # Sheet
REPORT_ID_2 = 3488341374750596  # Report

# === Initialize Smartsheet client ===
smartsheet_client = smartsheet.Smartsheet(ACCESS_TOKEN)
smartsheet_client.errors_as_exceptions(True)

def fetch_data_to_df(item_id, is_report=False):
    if is_report:
        report = smartsheet_client.Reports.get_report(item_id)
        col_map = {col.virtual_id: col.title for col in report.columns}
        rows_data = []
        for row in report.rows:
            row_dict = {}
            for cell in row.cells:
                col_title = col_map.get(cell.virtual_column_id, f"Col_{cell.virtual_column_id}")
                row_dict[col_title] = cell.value if cell.value is not None else ""
            rows_data.append(row_dict)
    else:
        sheet = smartsheet_client.Sheets.get_sheet(item_id)
        columns = [col.title for col in sheet.columns]
        rows_data = []
        for row in sheet.rows:
            row_dict = {}
            for cell, col in zip(row.cells, columns):
                row_dict[col] = cell.value if cell.value is not None else ""
            rows_data.append(row_dict)

    return pd.DataFrame(rows_data)

# === Fetch both datasets ===
df1 = fetch_data_to_df(SHEET_ID_1, is_report=False)
df2 = fetch_data_to_df(REPORT_ID_2, is_report=True)

# === Create summary tab: current month only, unique Project IDs, all columns ===
now = datetime.now()
current_month = now.month
current_year = now.year

combined_df = pd.concat([df1, df2], ignore_index=True)

# Ensure necessary columns exist
if "Live with Dealers Date" not in combined_df.columns or "Project ID" not in combined_df.columns:
    raise Exception("Missing 'Project ID' or 'Live with Dealers Date' in data")

# Convert date column
combined_df["Live with Dealers Date"] = pd.to_datetime(combined_df["Live with Dealers Date"], errors="coerce")

# Filter to current month/year and drop duplicate Project IDs
summary_df = combined_df[
    (combined_df["Live with Dealers Date"].dt.month == current_month) &
    (combined_df["Live with Dealers Date"].dt.year == current_year)
].drop_duplicates(subset="Project ID", keep="first")

# === Save to Excel ===
output_file = "merged_sheets.xlsx"
with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    df1.to_excel(writer, index=False, sheet_name="Sheet1")
    df2.to_excel(writer, index=False, sheet_name="Sheet2")
    summary_df.to_excel(writer, index=False, sheet_name="Project Summary")

print(f"✅ Data saved to {output_file}")
