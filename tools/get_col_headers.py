import os
import csv
from dotenv import load_dotenv
import smartsheet

# === LOAD API TOKEN FROM .env ===
load_dotenv()
ACCESS_TOKEN = os.getenv("SMARTSHEET_API_TOKEN")

# === SETUP ===
SHEET_ID = 3982583524183940
OUTPUT_CSV = "C:/Users/panderson/OneDrive - American Bath Group/Documents/Reports/column_headers.csv"

# === INITIALIZE CLIENT ===
smartsheet_client = smartsheet.Smartsheet(ACCESS_TOKEN)
smartsheet_client.errors_as_exceptions(True)

try:
    # Fetch sheet
    sheet = smartsheet_client.Sheets.get_sheet(SHEET_ID)

    # Get only visible (non-hidden) column titles
    headers = [col.title for col in sheet.columns if not col.hidden]
    # Get all column titles (including hidden)
    # headers = [col.title for col in sheet.columns]

    # Write to CSV
    with open(OUTPUT_CSV, mode='w', newline='', encoding='utf-8') as file:
        csv.writer(file).writerow(headers)

    print(f"✅ Column headers exported to: {OUTPUT_CSV}")

except smartsheet.exceptions.ApiError as e:
    print(f"❌ Smartsheet API Error: {e.error.result_code} - {e.error.message}")
