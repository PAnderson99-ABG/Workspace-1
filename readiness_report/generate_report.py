import pandas as pd
import os
import tkinter as tk
from tkinter import messagebox
import re

def is_filled(value):
    if pd.isna(value):
        return False
    value = str(value).strip().lower()
    return value != "" and value != "n/a"

def calculate_attribute_completion(df_data, headers, tiers):
    attr_data = []
    first_col = df_data.iloc[:, 0].astype(str)
    sellable_rows = df_data[first_col.str.strip().str.lower() == "sellable"]

    for i, attr in enumerate(headers):
        if not attr:
            continue
        tier_label = tiers[i] if i < len(tiers) else ""
        col = sellable_rows.iloc[:, i]
        non_blank = col.apply(is_filled).sum()
        total = len(col)
        completion_pct = round(non_blank / total, 4) if total else 0.0
        attr_data.append((attr, completion_pct, tier_label))
    return attr_data

def get_user_selected_sheets(sheet_names):
    selected_sheets = []

    def submit():
        for i, var in enumerate(sheet_vars):
            if var.get():
                selected_sheets.append(sheet_names[i])
        if not selected_sheets:
            messagebox.showwarning("No Selection", "Please select at least one sheet.")
            return
        root.destroy()

    root = tk.Tk()
    root.title("Select Product Categories")
    tk.Label(root, text="Select product categories to include in the report:").pack(pady=10)

    sheet_vars = []
    for name in sheet_names:
        var = tk.BooleanVar()
        cb = tk.Checkbutton(root, text=name, variable=var)
        cb.pack(anchor='w')
        sheet_vars.append(var)

    tk.Button(root, text="Generate Report", command=submit).pack(pady=10)
    root.mainloop()

    return selected_sheets

def process_workbook(input_path, selected_sheets):
    xl = pd.ExcelFile(input_path)
    static_blocks = []

    for sheet_name in selected_sheets:
        df = xl.parse(sheet_name, header=None)
        if df.shape[0] < 4 or df.iloc[0].isna().all():
            print(f"Skipping empty or malformed sheet: {sheet_name}")
            continue

        headers = df.iloc[0].tolist()
        tiers = df.iloc[2].tolist()
        data = df.iloc[3:].reset_index(drop=True)
        data.columns = headers

        data = data.replace("N/A", "").replace("n/a", "").replace("N/a", "")

        attr_completion = calculate_attribute_completion(data, headers, tiers)

        print(f"\n[DEBUG] Sheet: {sheet_name}")
        print(f"[DEBUG] Tiers Found: {list(set(t for _, _, t in attr_completion))}")

        static_blocks.append((sheet_name, attr_completion))

    return static_blocks

def build_static_df(static_blocks):
    max_rows = max(len(block) for _, block in static_blocks)

    padded = []
    for category, block in static_blocks:
        block_data = {
            "Attributes": [a for a, _, _ in block] + [""] * (max_rows - len(block)),
            category: [pct for _, pct, _ in block] + [""] * (max_rows - len(block)),
            f"Tier_{category}": [t for _, _, t in block] + [""] * (max_rows - len(block)),
        }
        padded.append(pd.DataFrame(block_data))

    from functools import reduce
    return reduce(lambda left, right: pd.concat([left, right], axis=1), padded)

def extract_tier_parts(tier_str):
    """Convert 'Tier - 1, 2, 3' to [1, 2, 3] for sorting."""
    match = re.findall(r'\d+', str(tier_str))
    return [int(m) for m in match]

def build_dashboard(static_df, static_blocks):
    # Extract and clean tier labels
    tier_labels = set()
    for col in static_df.columns:
        if col.startswith("Tier_"):
            tier_labels.update(static_df[col].dropna().astype(str))

    def is_valid_tier(t):
        return t.startswith("Tier -") and re.search(r'\d', t)

    # Filter and sort tiers logically
    valid_tiers = [t for t in tier_labels if is_valid_tier(t)]
    sorted_tiers = sorted(valid_tiers, key=extract_tier_parts)

    print(f"[DEBUG] Sorted Dashboard Tiers: {sorted_tiers}")

    # Build rows per category
    product_categories = [cat for cat, _ in static_blocks]
    dashboard_records = []

    for cat in product_categories:
        tier_col = f"Tier_{cat}"
        pct_col = cat

        row = {"Product Category": cat}
        for tier_label in sorted_tiers:
            mask = static_df[tier_col].astype(str).apply(
                lambda attr_tier: isinstance(attr_tier, str) and tier_label in attr_tier
            )
            relevant = static_df.loc[mask, pct_col]
            numeric_values = pd.to_numeric(relevant, errors='coerce').dropna()
            row[tier_label] = round(numeric_values.mean(), 4) if not numeric_values.empty else ""

        dashboard_records.append(row)

    return pd.DataFrame(dashboard_records)

def main(input_file, output_file):
    xl = pd.ExcelFile(input_file)
    sheet_names = xl.sheet_names
    selected_sheets = get_user_selected_sheets(sheet_names)

    static_blocks = process_workbook(input_file, selected_sheets)

    static_df = build_static_df(static_blocks)
    dashboard_df = build_dashboard(static_df, static_blocks)

    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        static_df.to_excel(writer, sheet_name="static", index=False)
        dashboard_df.to_excel(writer, sheet_name="dashboard", index=False)

    print(f"\n✅ Report generated: {os.path.abspath(output_file)}")

if __name__ == "__main__":
    input_file = "input.xlsx"
    output_file = "generated_output.xlsx"
    main(input_file, output_file)