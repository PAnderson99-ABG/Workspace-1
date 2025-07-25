import pandas as pd
import tkinter as tk
from tkinter import filedialog
import os

def row4_has_data(sheet_df):
    """Check if row 4 has any non-blank/non-NaN data"""
    if sheet_df.shape[0] < 4:
        return False
    row4 = sheet_df.iloc[3]  # 0-based index, so row 4 = index 3
    return row4.notna().any() and (row4.astype(str).str.strip() != "").any()

def main():
    # GUI File Selection
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(title="Select Salsify Export", filetypes=[("Excel files", "*.xlsx")])

    if not file_path:
        print("No file selected. Exiting.")
        return

    # Load all sheets
    xls = pd.ExcelFile(file_path)
    valid_sheets = {}

    for sheet_name in xls.sheet_names:
        df = xls.parse(sheet_name, header=None)  # No header to ensure row indexing is accurate
        if row4_has_data(df):
            valid_sheets[sheet_name] = df
        else:
            print(f"Deleted tab: {sheet_name}")

    if not valid_sheets:
        print("No sheets with data in row 4. Nothing to save.")
        return

    # Save new workbook
    output_path = file_path.replace(".xlsx", "_trimmed.xlsx")
    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        for sheet_name, df in valid_sheets.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)

    print(f"Trimmed file saved to: {output_path}")

if __name__ == "__main__":
    main()
