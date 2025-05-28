import os
from dotenv import load_dotenv
import smartsheet
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook

# === Load API Token from .env ===
load_dotenv()
ACCESS_TOKEN = os.getenv("SMARTSHEET_API_TOKEN")

# === Initialize Smartsheet Client ===
smartsheet_client = smartsheet.Smartsheet(ACCESS_TOKEN)
smartsheet_client.errors_as_exceptions(True)

# === Configuration ===
SHEET_ID = int(os.getenv("PCM_REQUESTS_ID"))
EXPORT_FOLDER = r"C:/Users/panderson/OneDrive - American Bath Group/Documents/Reports"

# === Date Range for Current Month ===
now = datetime.now()
month_start = now.replace(day=1)
month_end = now

# === Fetch Sheet and Column Map ===
sheet = smartsheet_client.Sheets.get_sheet(SHEET_ID)
column_id_map = {col.id: col.title for col in sheet.columns}
column_title_map = {col.title: col.id for col in sheet.columns}

date_requested_col_id = column_title_map.get("Date Requested")
if not date_requested_col_id:
    raise Exception("Missing 'Date Requested' column.")

# === Filter Rows Within Current Month ===
included_rows = []

for row in sheet.rows:
    date_value = next((cell.value for cell in row.cells if cell.column_id == date_requested_col_id), None)
    if not date_value:
        continue

    try:
        if isinstance(date_value, str):
            try:
                request_date = datetime.strptime(date_value.strip(), "%m/%d/%Y")
            except ValueError:
                request_date = datetime.strptime(date_value.strip(), "%Y-%m-%d")
        elif isinstance(date_value, datetime):
            request_date = date_value
        elif hasattr(date_value, 'isoformat'):
            request_date = datetime.combine(date_value, datetime.min.time())
        else:
            continue

        if month_start <= request_date <= month_end:
            included_rows.append((row, request_date))

    except Exception as e:
        print(f"⏭️ Row {row.id} skipped — invalid date format: {e}")
        continue

print(f"✅ Found {len(included_rows)} rows from current month.")

# === Extract Row Data and Comments ===
data = []

for row, request_date in included_rows:
    row_dict = {column_id_map.get(cell.column_id, f"Col_{cell.column_id}"): cell.value for cell in row.cells}

    try:
        discussions = smartsheet_client.Discussions.get_row_discussions(
            sheet_id=SHEET_ID,
            row_id=row.id,
            include="comments"
        ).data

        all_comments = []
        latest_timestamp = None
        latest_author = ""

        for discussion in discussions:
            for comment in discussion.comments:
                author = comment.created_by.name if comment.created_by else "Unknown"
                comment_time = comment.created_at
                comment_text = comment.text.strip().replace('\n', ' ')

                all_comments.append(f"{author} [{comment_time.strftime('%Y-%m-%d %H:%M:%S')}]: {comment_text}")

                if not latest_timestamp or comment_time > latest_timestamp:
                    latest_timestamp = comment_time
                    latest_author = author

        row_dict["All Comments"] = "; ".join(all_comments)
        row_dict["Comment Author"] = latest_author
        row_dict["Comment Date"] = latest_timestamp.strftime("%Y-%m-%d %H:%M:%S") if latest_timestamp else ""

        print(f"✅ Row {row.id} added — {len(all_comments)} comments")

    except Exception as e:
        print(f"⚠️ Row {row.id} — error fetching comments: {e}")
        row_dict["All Comments"] = ""
        row_dict["Comment Author"] = ""
        row_dict["Comment Date"] = ""

    data.append(row_dict)

# === Export to Excel ===
df = pd.DataFrame(data)
today_str = datetime.now().strftime("%Y-%m-%d")
filename = f"Smartsheet_Export_{today_str}.xlsx"
output_path = os.path.join(EXPORT_FOLDER, filename)
df.to_excel(output_path, index=False)
print(f"\n✅ Excel file saved: {output_path}")

# === Compute Monthly Metrics ===
# Ensure correct data types
df["Status"] = df["Status"].fillna("— Unspecified —")
df["SKU Count"] = pd.to_numeric(df["SKU Count"], errors="coerce").fillna(0).astype(int)

# Total SKUs and Status Counts
sku_total = df["SKU Count"].sum()
status_counts = df["Status"].value_counts()
sku_per_status = df.groupby("Status")["SKU Count"].sum().sort_values(ascending=False)

# Output metrics to .txt file
metrics_filename = f"PCM_Metrics.txt"
metrics_path = os.path.join(EXPORT_FOLDER, metrics_filename)

with open(metrics_path, "w", encoding="utf-8") as f:
    f.write(f"📦 Total SKUs submitted this month: {sku_total:,}\n\n")
    f.write("📊 Status frequency breakdown:\n")
    for status, count in status_counts.items():
        f.write(f"- {status}: {count} rows\n")
    f.write("\n📦 Total SKUs per status:\n")
    for status, total in sku_per_status.items():
        f.write(f"- {status}: {total:,} SKUs\n")

print(f"\n📝 Metrics written to text file: {metrics_path}")

