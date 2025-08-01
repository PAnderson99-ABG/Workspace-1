import pandas as pd
from .write_helpers import autofit_columns

# === DEFINE GROUPS ===
PIM_TEAM = {
    "justined@vintagetub.com",
    "olepagedumont@americanbathgroup.com",
    "panderson@americanbathgroup.com",
    "rclough@americanbathgroup.com",
    "lmasterson@americanbathgroup.com",
    "ktubbs@americanbathgroup.com",
    "mwaxman@americanbathgroup.com"
}

DREAMLINE_TEAM = {
    "vadimm@dreamline.com",
    "michael.snyder@dreamline.com",
    "michelle.sanders@dreamline.com",
    "alinak@dreamline.com",
    "janinec@vintagetub.com",
    "leo.volkov@dreamline.com",
    "aspeck@americanbathgroup.com"
}

MAAX_TEAM = {
    "jleboeuf@produitsneptune.com",
    "lina.napoli@maax.com",
    "cavellaneda@americanbathgroup.com",
    "hbernier@americanbathgroup.com",
    "luciana.disario@maax.com",
    "mhoumadi@americanbathgroup.com",
    "suzanne.jacques@maax.com",
    "mbesner@americanbathgroup.com",
    "kdevries-thomas@americanbathgroup.com",
    "adiaw@americanbathgroup.com",
    "amunozperez@americanbathgroup.com",
    "bmassimetti@americanbathgroup.com",
    "bparekh@americanbathgroup.com",
    "csexton@americanbathgroup.com",
    "mlongo@americanbathgroup.com",
    "oelbacha@americanbathgroup.com",
    "rdagenais@americanbathgroup.com",
    "lcieza@americanbathgroup.com",
    "gdeland@americanbathgroup.com",
    "kpeak@americanbathgroup.com",
    "lbourgeois@americanbathgroup.com",
    "s.e.bernard@americanbathgroup.com",
    "asimard@americanbathgroup.com",
    "vbehzadi@americanbathgroup.com",
    "csthilaire@americanbathgroup.com",
    "falvarado@americanbathgroup.com",
    "gdionne@americanbathgroup.com",
    "jonathan.thibault@maax.com",
    "sdoyle@americanbathgroup.com",
    "sbeauregard@neptuneb.com",
    "rdayyani@americanbathgroup.com",
    "tboulanger@americanbathgroup.com",
    "marie-helene.dumont@maax.com"
}

GROUP_ORDER = ["PIM Team", "Dreamline Team", "MAAX Team", "Other"]

# === DEFINE COLORS ===
GROUP_COLORS = {
    "PIM Team": "#4472C4",
    "Dreamline Team": "#ED7D31",
    "MAAX Team": "#FFC000",
    "Other": "#A9D18E"
}
DEFAULT_COLOR = "#D9D9D9"

def get_user_group(email):
    if email in PIM_TEAM:
        return "PIM Team"
    elif email in DREAMLINE_TEAM:
        return "Dreamline Team"
    elif email in MAAX_TEAM:
        return "MAAX Team"
    else:
        return "Other"

def generate_monthly_brand_breakdowns(df, writer, brand_name=None):
    df["month"] = df["timestamp"].dt.strftime("%b")
    month_order = sorted(df["month"].unique(), key=lambda m: pd.to_datetime(m, format="%b").month)

    # Infer brand from first value (assumes df is already filtered)
    if brand_name is None:
        chart_title_prefix = "Team"
    else:
        chart_title_prefix = f"{brand_name.capitalize()} Team"


    for month in month_order:
        month_df = df[df["month"] == month].copy()
        month_df["Group"] = month_df["user_email"].apply(get_user_group)

        # === GROUP SUMMARY TABLE ===
        group_counts = (
            month_df["Group"]
            .value_counts()
            .rename_axis("Group")
            .reset_index(name="Change Count")
        )
        group_counts["GroupOrder"] = group_counts["Group"].map({g: i for i, g in enumerate(GROUP_ORDER)})
        group_counts = group_counts.sort_values(by="GroupOrder").drop(columns="GroupOrder").reset_index(drop=True)

        if group_counts.empty:
            continue

        # === SHEET SETUP ===
        sheet = writer.book.add_worksheet(month)
        header_format = writer.book.add_format({'bold': True, 'bg_color': '#D9E1F2', 'border': 1})
        cell_format = writer.book.add_format({'border': 1})

        total_changes = group_counts["Change Count"].sum()
        group_counts["Percentage"] = group_counts["Change Count"] / total_changes

        percent_format = writer.book.add_format({'num_format': '0.00%', 'border': 1})

        sheet.write_row(0, 0, ["Group", "Change Count", "Percentage"], header_format)

        for i, row in group_counts.iterrows():
            sheet.write(i + 1, 0, row["Group"], cell_format)
            sheet.write(i + 1, 1, row["Change Count"], cell_format)
            sheet.write(i + 1, 2, row["Percentage"], percent_format)

        color_points = [
            {'fill': {'color': GROUP_COLORS.get(g, DEFAULT_COLOR)}}
            for g in group_counts["Group"]
        ]

# === CHART INSERTION (revised portion only) ===

        chart = writer.book.add_chart({'type': 'column'})
        chart.add_series({
            'name': f"{chart_title_prefix} Contributions - {month}",
            'categories': [month, 1, 0, len(group_counts), 0],
            'values':     [month, 1, 1, len(group_counts), 1],
            'data_labels': {'value': True},
            'points': color_points
        })
        chart.set_title({'name': f"{chart_title_prefix} Contributions - {month}"})
        chart.set_x_axis({'name': 'Group'})
        chart.set_y_axis({'name': 'Change Count'})
        chart.set_style(10)
        chart.set_size({'width': 600, 'height': 400})
        sheet.insert_chart("E2", chart)

        # === USER BREAKDOWN ===
        user_breakdown = (
            month_df.groupby(["Group", "user_email"])
            .size()
            .reset_index(name="Change Count")
        )
        user_breakdown["GroupOrder"] = user_breakdown["Group"].map({g: i for i, g in enumerate(GROUP_ORDER)})
        user_breakdown = user_breakdown.sort_values(
            by=["GroupOrder", "Change Count"],
            ascending=[True, False]
        ).drop(columns="GroupOrder")

        subtotals = user_breakdown.groupby("Group")["Change Count"].sum().to_dict()

        start_row = 6
        sheet.write_row(start_row, 0, ["Group", "Email", "Change Count", "Subtotal"], header_format)

        row_cursor = start_row + 1
        previous_group = None

        for _, row in user_breakdown.iterrows():
            group = row["Group"]
            email = row["user_email"]
            count = row["Change Count"]

            sheet.write(row_cursor, 0, group if group != previous_group else "", cell_format)
            sheet.write(row_cursor, 1, email, cell_format)
            sheet.write(row_cursor, 2, count, cell_format)

            if group != previous_group:
                sheet.write(row_cursor, 3, subtotals[group], cell_format)

            previous_group = group
            row_cursor += 1

        display_df = user_breakdown.rename(columns={"user_email": "Email"})
        display_df["Subtotal"] = display_df["Group"].map(subtotals)
        autofit_columns(sheet, display_df[["Group", "Email", "Change Count", "Subtotal"]])
