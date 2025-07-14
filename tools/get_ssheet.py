import os
import csv
from dotenv import load_dotenv
import smartsheet

# === LOAD API TOKEN FROM .env ===
load_dotenv()
ACCESS_TOKEN = os.getenv("SMARTSHEET_API_TOKEN")

# === SETUP ===
SHEET_ID = 3657436246265732
OUTPUT_CSV = "C:/Users/panderson/OneDrive - American Bath Group/Documents/Paul_Anderson/Reports/sheet_data.csv"

# === INITIALIZE CLIENT ===
smartsheet_client = smartsheet.Smartsheet(ACCESS_TOKEN)
smartsheet_client.errors_as_exceptions(True)

try:
    # Fetch sheet
    sheet = smartsheet_client.Sheets.get_sheet(SHEET_ID)

    # Get only visible (non-hidden) columns
    visible_columns = [col for col in sheet.columns if not col.hidden]
    headers = [col.title for col in visible_columns]
    col_ids = [col.id for col in visible_columns]

    # Write to CSV
    with open(OUTPUT_CSV, mode='w', newline='', encoding='utf-8-sig') as file:
        writer = csv.writer(file)
        writer.writerow(headers)

        for row in sheet.rows:
            # Skip rows filtered out by a sheet filter, if applicable
            if hasattr(row, 'filtered_out') and row.filtered_out:
                continue

            cell_map = {cell.column_id: cell for cell in row.cells}
            row_data = [
                cell_map[col_id].display_value if col_id in cell_map and cell_map[col_id].display_value is not None else ""
                for col_id in col_ids
            ]
            writer.writerow(row_data)

    print(f"✅ Visible sheet data exported to: {OUTPUT_CSV}")

except smartsheet.exceptions.ApiError as e:
    print(f"❌ Smartsheet API Error: {e.error.result_code} - {e.error.message}")
