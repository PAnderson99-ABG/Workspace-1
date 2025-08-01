import os
import csv
from dotenv import load_dotenv
import smartsheet

# === LOAD API TOKEN FROM .env ===
load_dotenv()
ACCESS_TOKEN = os.getenv("SMARTSHEET_API_TOKEN")

# === CONFIGURATION ===
SMARTSHEET_ID = 3488341374750596  # Can be a Sheet ID or Report ID
OUTPUT_CSV = "C:/Users/panderson/OneDrive - American Bath Group/Documents/Paul_Anderson/Reports/output/sheet_data.csv"

# === INITIALIZE CLIENT ===
smartsheet_client = smartsheet.Smartsheet(ACCESS_TOKEN)
smartsheet_client.errors_as_exceptions(True)

from datetime import datetime

def format_cell(cell):
    """Format cell.value for output"""
    if cell.value is None:
        return ""
    elif isinstance(cell.value, bool):
        return "TRUE" if cell.value else "FALSE"
    elif isinstance(cell.value, datetime):
        return cell.value.strftime("%Y-%m-%d")
    return str(cell.value)

def export_sheet(sheet):
    visible_columns = [col for col in sheet.columns if not col.hidden]
    headers = [col.title for col in visible_columns]
    col_ids = [col.id for col in visible_columns]

    with open(OUTPUT_CSV, mode='w', newline='', encoding='utf-8-sig') as file:
        writer = csv.writer(file)
        writer.writerow(headers)

        for row in sheet.rows:
            if hasattr(row, 'filtered_out') and row.filtered_out:
                continue
            cell_map = {cell.column_id: cell for cell in row.cells}
            row_data = [
                format_cell(cell_map[col_id]) if col_id in cell_map else ""
                for col_id in col_ids
            ]
            writer.writerow(row_data)
    print(f"✅ Sheet data exported to: {OUTPUT_CSV}")

def export_report(report):
    headers = [col.title for col in report.columns]
    col_ids = [col.virtual_id for col in report.columns]

    with open(OUTPUT_CSV, mode='w', newline='', encoding='utf-8-sig') as file:
        writer = csv.writer(file)
        writer.writerow(headers)

        for row in report.rows:
            cell_map = {cell.virtual_column_id: cell for cell in row.cells}
            row_data = [
                format_cell(cell_map[col_id]) if col_id in cell_map else ""
                for col_id in col_ids
            ]
            writer.writerow(row_data)
    print(f"✅ Report data exported to: {OUTPUT_CSV}")

try:
    # Try to fetch as a sheet first
    try:
        sheet = smartsheet_client.Sheets.get_sheet(SMARTSHEET_ID)
        export_sheet(sheet)
    except smartsheet.exceptions.ApiError as sheet_err:
        # Try as report instead
        report = smartsheet_client.Reports.get_report(SMARTSHEET_ID)
        export_report(report)

except smartsheet.exceptions.ApiError as e:
    print(f"❌ Smartsheet API Error: {e.error.result_code} - {e.error.message}")
