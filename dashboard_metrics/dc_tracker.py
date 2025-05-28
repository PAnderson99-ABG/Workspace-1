import os
from dotenv import load_dotenv
import smartsheet
import pandas as pd
from datetime import datetime

# === Load API Token ===
load_dotenv()
ACCESS_TOKEN = os.getenv("SMARTSHEET_API_TOKEN")

# === Initialize Smartsheet Client ===
smartsheet_client = smartsheet.Smartsheet(ACCESS_TOKEN)
smartsheet_client.errors_as_exceptions(True)

# === Sheet Configuration ===
SHEET_ID = int(os.getenv("MDM_DC_TRACKER"))  # From .env
EXPORT_FOLDER = r"C:/Users/panderson/OneDrive - American Bath Group/Documents/Reports"
TARGET_COLUMN_NAME = "Completed Data Contract Attached Date"

# === Date Setup ===
now = datetime.now()
current_year = now.year
current_month = now.month

# === Fetch Sheet and Column Map ===
sheet = smartsheet_client.Sheets.get_sheet(SHEET_ID)
column_id_map = {col.id: col.title for col in sheet.columns}
column_title_map = {col.title: col.id for col in sheet.columns}
target_col_id = column_title_map.get(TARGET_COLUMN_NAME)

if not target_col_id:
    raise Exception(f"Column '{TARGET_COLUMN_NAME}' not found in the sheet.")

# === Filter Rows by Date ===
filtered_data = []

for row in sheet.rows:
    row_dict = {}
    target_date_raw = None

    for cell in row.cells:
        col_title = column_id_map.get(cell.column_id, f"Col_{cell.column_id}")
        row_dict[col_title] = cell.value

        if cell.column_id == target_col_id:
            target_date_raw = cell.value

    # Attempt to parse date and filter
    try:
        if isinstance(target_date_raw, str):
            parsed_date = datetime.strptime(target_date_raw.strip(), "%Y-%m-%d")
        elif isinstance(target_date_raw, datetime):
            parsed_date = target_date_raw
        elif hasattr(target_date_raw, 'isoformat'):
            parsed_date = datetime.combine(target_date_raw, datetime.min.time())
        else:
            continue

        if parsed_date.year == current_year and parsed_date.month == current_month:
            filtered_data.append(row_dict)

    except Exception as e:
        print(f"⚠️ Row {row.id} skipped — invalid date format: {e}")
        continue

# === Export to Excel ===
df = pd.DataFrame(filtered_data)
today_str = now.strftime("%Y-%m-%d")
filename_excel = f"Filtered_MDM_DC_Export_{today_str}.xlsx"
output_path_excel = os.path.join(EXPORT_FOLDER, filename_excel)
df.to_excel(output_path_excel, index=False)
print(f"✅ Excel file saved to: {output_path_excel}")

# === Compute Metrics ===
df["Completed Data Contract"] = pd.to_numeric(df.get("Completed Data Contract"), errors="coerce").fillna(0).astype(int)
df["Number of SKUs (Base + Sellable)"] = pd.to_numeric(df.get("Number of SKUs (Base + Sellable)"), errors="coerce").fillna(0)

validated_df = df[df["Completed Data Contract"] == 1]
num_contracts = len(validated_df)
total_skus = int(validated_df["Number of SKUs (Base + Sellable)"].sum())

# === Write Metrics to TXT File with UTF-8 ===
filename_txt = f"MDM_DC_Metrics.txt"
output_path_txt = os.path.join(EXPORT_FOLDER, filename_txt)

with open(output_path_txt, "w", encoding="utf-8") as f:
    f.write(f"📊 MDM Data Contract Tracker Metrics — {now.strftime('%B %Y')}\n")
    f.write("=" * 50 + "\n")
    f.write(f"✅ Completed Data Contracts This Month: {num_contracts}\n")
    f.write(f"📦 Total SKUs Onboarded This Month: {total_skus}\n")

print(f"✅ Metrics saved to: {output_path_txt}")
