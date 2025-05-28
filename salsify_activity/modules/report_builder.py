# === report_builder.py ===
import pandas as pd
from .charts import insert_summary_chart, insert_changes_over_time_sheet
from .write_helpers import write_breakdown_table, autofit_columns
from .monthly_breakdown import generate_monthly_brand_breakdowns

def build_excel_report(df, top_fields, all_users, top_brands, brand_user_df, property_user_df):
    output_file = "report.xlsx"
    with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:

        # 1. Changes Over Time
        insert_changes_over_time_sheet(writer, df)
        from_date = df["timestamp"].dt.date.min()
        to_date = df["timestamp"].dt.date.max()
        daily_counts = df["timestamp"].dt.date.value_counts().sort_index().reset_index(name="change_count")
        daily_counts.columns = ["date", "change_count"]
        autofit_columns(writer.sheets["Changes Over Time"], daily_counts)

        # 2. All Users
        all_users.to_excel(writer, sheet_name="All Users", index=False)
        autofit_columns(writer.sheets["All Users"], all_users)

        # 3. Top Brands
        top_brands.to_excel(writer, sheet_name="Top Brands", index=False)
        autofit_columns(writer.sheets["Top Brands"], top_brands)

        # 4. Brand_User_Breakdown
        brand_ws = writer.book.add_worksheet("Brand_User_Breakdown")
        write_breakdown_table(brand_ws, brand_user_df, "brand")
        # Optional: autofit skipped here since it's manually written

        # 5. Top Properties
        top_fields.to_excel(writer, sheet_name="Top Properties", index=False)
        autofit_columns(writer.sheets["Top Properties"], top_fields)

        # 6. Property_User_Breakdown
        property_ws = writer.book.add_worksheet("Property_User_Breakdown")
        write_breakdown_table(property_ws, property_user_df, "property_name")
        # Optional: autofit skipped here too

        # 7. Monthly breakdowns (adds sheets and charts)
        generate_monthly_brand_breakdowns(df, writer)

        # Insert summary charts
        wb = writer.book
        insert_summary_chart(wb, writer.sheets["Top Properties"], "Top 50 Properties Changed", 0, 1, len(top_fields))
        insert_summary_chart(wb, writer.sheets["All Users"], "Top 10 Users by Changes", 0, 1, min(10, len(all_users)))
        insert_summary_chart(wb, writer.sheets["Top Brands"], "Top 50 Brands Changed", 0, 1, len(top_brands))

        for sheet_name in [
            "All Users",
            "Top Brands",
            "Brand_User_Breakdown",
            "Top Properties",
            "Property_User_Breakdown"
        ]:
            if sheet_name in writer.sheets:
                writer.sheets[sheet_name].hidden = True


    print("✅ Excel report generated: report.xlsx")
