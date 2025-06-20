import pandas as pd
from tkinter import Tk, filedialog

# === FILE PICKER ===
def select_file(prompt):
    root = Tk()
    root.withdraw()
    return filedialog.askopenfilename(title=prompt, filetypes=[("Excel files", "*.xlsx *.xls")])

# === MAIN SCRIPT ===
print("📁 Select the marketing output file to fill...")
file_path = select_file("Select Marketing Output Excel File")

# Load Excel file
df = pd.read_excel(file_path, dtype=object)  # Load all as str/objects to preserve blanks

# Identify key columns
BASE_COL = "Base Part Number"
SELLABLE_COL = "Sellable Part Number"

# Track current base row
base_row = None

# Process each row
for i in range(len(df)):
    row = df.iloc[i]

    if pd.notna(row.get(BASE_COL)):  # It's a base part
        base_row = row
        continue

    if pd.notna(row.get(SELLABLE_COL)) and base_row is not None:  # It's a sellable
        for col in df.columns:
            if pd.isna(row[col]) and pd.notna(base_row[col]):
                df.at[i, col] = base_row[col]

# Save the updated file
output_path = file_path.replace(".xlsx", "-hierarchy-filled.xlsx")
df.to_excel(output_path, index=False)
print(f"✅ Filled sellables based on base part hierarchy.\n📄 Saved to: {output_path}")
