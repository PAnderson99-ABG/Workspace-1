import pandas as pd
import os
from tkinter import Tk, filedialog

# === File Picker Dialog ===
def select_file(title, filetypes):
    root = Tk()
    root.withdraw()  # Hide the root window
    file_path = filedialog.askopenfilename(title=title, filetypes=filetypes)
    root.destroy()
    return file_path

def save_file(title, defaultextension, filetypes):
    root = Tk()
    root.withdraw()
    file_path = filedialog.asksaveasfilename(title=title, defaultextension=defaultextension, filetypes=filetypes)
    root.destroy()
    return file_path

# === Prompt for input and output files ===
INPUT_FILE_PATH = select_file("Select Excel File to Format", [("Excel files", "*.xlsx")])
if not INPUT_FILE_PATH:
    raise ValueError("❌ No input file selected.")

OUTPUT_FILE_PATH = save_file("Save Formatted Excel File As", ".xlsx", [("Excel files", "*.xlsx")])
if not OUTPUT_FILE_PATH:
    raise ValueError("❌ No output file selected.")

# === Load Excel Data ===
df = pd.read_excel(INPUT_FILE_PATH, dtype=str)
df.fillna("", inplace=True)

# === Identify base products (Unique ID == Base Part Number) ===
base_df = df[df["Unique ID"] == df["Base Part Number"]]
base_ids = sorted(base_df["Unique ID"].unique())  # Lexicographic sort

# === Build reordered DataFrame ===
reordered_rows = []

for base_id in base_ids:
    base_row = df[(df["Unique ID"] == base_id) & (df["Base Part Number"] == base_id)]
    child_rows = df[(df["Base Part Number"] == base_id) & (df["Unique ID"] != base_id)]
    if not base_row.empty:
        reordered_rows.append(base_row)
    if not child_rows.empty:
        reordered_rows.append(child_rows)

# === Concatenate all rows
output_df = pd.concat(reordered_rows, ignore_index=True)

# === Save to Excel
output_df.to_excel(OUTPUT_FILE_PATH, index=False)

print(f"✅ Restructured file saved to: {OUTPUT_FILE_PATH}")
