import pandas as pd
from tkinter import Tk, filedialog
from marketing_cols import COLUMNS  # ← Column list is now imported

# === CONFIGURATION ===
OUTPUT_PATH = "C:/Users/panderson/OneDrive - American Bath Group/Documents/Paul_Anderson/Reports/marketing-output.xlsx"

# === FILE PICKER UTILITY ===
def select_file(prompt):
    root = Tk()
    root.withdraw()
    return filedialog.askopenfilename(title=prompt, filetypes=[("Excel files", "*.xlsx *.xls")])

# === MAIN ===
print("📁 Select your EXPORT (Marketing Data) file")
export_path = select_file("Select Export File")
print("📁 Select your IMAGE DATA file")
data_path = select_file("Select Image Data File")

# Load data
export_df = pd.read_excel(export_path, sheet_name=0)
data_df = pd.read_excel(data_path, sheet_name=0)

# Set Unique ID as index
export_df.set_index("Unique ID", inplace=True)
data_df.set_index("Unique ID", inplace=True)

# Create output with union of IDs
output = pd.DataFrame(index=export_df.index.union(data_df.index))

# Fill columns
for col in COLUMNS:
    if col == "Unique ID":
        continue
    col_export = export_df[col] if col in export_df.columns else pd.Series(dtype=object)
    col_data = data_df[col] if col in data_df.columns else pd.Series(dtype=object)
    output[col] = col_export.combine_first(col_data)

# Finalize
output.reset_index(inplace=True)
output = output[COLUMNS]
output.to_excel(OUTPUT_PATH, index=False)
print(f"✅ File saved to: {OUTPUT_PATH}")
