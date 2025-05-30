import os
from dotenv import load_dotenv
import smartsheet
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook

# === Load API Token ===
load_dotenv()
ACCESS_TOKEN = os.getenv("SMARTSHEET_API_TOKEN")

# === Initialize Smartsheet Client ===
smartsheet_client = smartsheet.Smartsheet(ACCESS_TOKEN)
smartsheet_client.errors_as_exceptions(True)

# === Sheet Configuration ===
SHEET_ID = int(os.getenv("MDM_DC_TRACKER"))  # From .env
EXPORT_FOLDER = r"C:/Users/panderson/OneDrive - American Bath Group/Documents/Reports"

# === Date Setup ===
now = datetime.now()
current_year = now.year
current_month = now.month

# === Fetch Sheet and Column Map ===
sheet = smartsheet_client.Sheets.get_sheet(SHEET_ID)
column_id_map = {col.id: col.title for col in sheet.columns}

# === Extract Rows and Filter ===
all_data = []
filtered_live_dealers = []
filtered_request_attached = []

for row in sheet.rows:
    row_dict = {
        column_id_map.get(cell.column_id, f"Col_{cell.column_id}"): cell.value
        for cell in row.cells
    }
    all_data.append(row_dict)

    # === Filter by "Live with Dealers Date"
    live_date = row_dict.get("Live with Dealers Date")
    try:
        if isinstance(live_date, str):
            live_date = datetime.strptime(live_date.strip(), "%Y-%m-%d")
        elif hasattr(live_date, 'isoformat'):
            live_date = datetime.combine(live_date, datetime.min.time())
        if live_date.year == current_year and live_date.month == current_month:
            filtered_live_dealers.append(row_dict)
    except:
        pass

    # === Filter by "Request Attached Date"
    req_date = row_dict.get("Request Attached Date")
    try:
        if isinstance(req_date, str):
            req_date = datetime.strptime(req_date.strip(), "%Y-%m-%d")
        elif hasattr(req_date, 'isoformat'):
            req_date = datetime.combine(req_date, datetime.min.time())
        if req_date.year == current_year and req_date.month == current_month:
            filtered_request_attached.append(row_dict)
    except:
        pass

# === Export to Excel with Three Sheets ===
df_all = pd.DataFrame(all_data)
df_live = pd.DataFrame(filtered_live_dealers)
df_request = pd.DataFrame(filtered_request_attached)

today_str = now.strftime("%Y-%m-%d")
filename_excel = f"MDM_DC_Export_{today_str}.xlsx"
output_path_excel = os.path.join(EXPORT_FOLDER, filename_excel)

with pd.ExcelWriter(output_path_excel, engine="openpyxl") as writer:
    df_all.to_excel(writer, index=False, sheet_name="All Rows")
    df_live.to_excel(writer, index=False, sheet_name="Live with Dealers")
    df_request.to_excel(writer, index=False, sheet_name="Request Attached")

print(f"✅ Excel file with 3 sheets saved to: {output_path_excel}")
