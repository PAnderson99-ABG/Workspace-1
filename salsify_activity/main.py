# === main.py ===
from modules.load_data import load_data
from modules.summarize import generate_summaries, generate_breakdowns
from modules.report_builder import build_excel_report

if __name__ == "__main__":
    df = load_data()
    top_fields, all_users, top_brands = generate_summaries(df)
    brand_user_df = generate_breakdowns(df, top_brands, "brand")
    property_user_df = generate_breakdowns(df, top_fields, "property_name")
    build_excel_report(df, top_fields, all_users, top_brands, brand_user_df, property_user_df)
