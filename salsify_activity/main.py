# === main.py ===
from modules.load_data import load_data
from modules.summarize import generate_summaries, generate_breakdowns
from modules.report_builder import build_excel_report
from datetime import datetime

if __name__ == "__main__":
    df = load_data()

    # Normalize brand names
    df["brand"] = df["brand"].str.strip().str.lower()

    # === Individual brand reports ===
    for brand in ["dreamline", "maax"]:
        brand_df = df[df["brand"] == brand].copy()

        print(f"📊 Generating report for brand: {brand}")

        top_fields, all_users, top_brands = generate_summaries(brand_df)
        brand_user_df = generate_breakdowns(brand_df, top_brands, "brand")
        property_user_df = generate_breakdowns(brand_df, top_fields, "property_name")

        build_excel_report(
            df=brand_df,
            top_fields=top_fields,
            all_users=all_users,
            top_brands=top_brands,
            brand_user_df=brand_user_df,
            property_user_df=property_user_df,
            output_filename=f"{brand}_report.xlsx"
        )

    # === Combined report: Who Did What - YYYY-MM-DD.xlsx ===
    today_str = datetime.today().strftime("%Y-%m-%d")
    combined_filename = f"Who Did What - {today_str}.xlsx"

    print(f"📊 Generating combined report: {combined_filename}")

    top_fields, all_users, top_brands = generate_summaries(df)
    brand_user_df = generate_breakdowns(df, top_brands, "brand")
    property_user_df = generate_breakdowns(df, top_fields, "property_name")

    build_excel_report(
        df=df,
        top_fields=top_fields,
        all_users=all_users,
        top_brands=top_brands,
        brand_user_df=brand_user_df,
        property_user_df=property_user_df,
        output_filename=combined_filename,
    )

