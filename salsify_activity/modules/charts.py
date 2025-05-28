# === charts.py ===

def insert_summary_chart(wb, ws, title, label_col, value_col, count):
    chart = wb.add_chart({"type": "column"})
    chart.add_series({
        "name": title,
        "categories": [ws.name, 1, label_col, count, label_col],
        "values":     [ws.name, 1, value_col, count, value_col],
    })
    chart.set_title({"name": title})
    chart.set_x_axis({"name": "", "label_position": "low"})
    chart.set_y_axis({"name": "Number of Changes"})
    chart.set_size({'width': 900, 'height': 600})
    ws.insert_chart("D2", chart)

def insert_changes_over_time_sheet(writer, df):
    df["date"] = df["timestamp"].dt.date
    daily_counts = df.groupby("date").size().reset_index(name="change_count")

    sheet_name = "Changes Over Time"
    daily_counts.to_excel(writer, sheet_name=sheet_name, index=False)

    if daily_counts.empty:
        print("⚠️ Skipping 'Changes Over Time' chart: No valid timestamp data.")
        return

    wb = writer.book
    ws = writer.sheets[sheet_name]

    chart = wb.add_chart({"type": "line"})
    chart.add_series({
        "name": "Changes Per Day",
        "categories": [sheet_name, 1, 0, len(daily_counts), 0],
        "values":     [sheet_name, 1, 1, len(daily_counts), 1],
    })
    chart.set_title({"name": "Change Volume Over Time"})
    chart.set_x_axis({"name": "Date"})
    chart.set_y_axis({"name": "Number of Changes"})
    chart.set_size({"width": 900, "height": 400})
    ws.insert_chart("D2", chart)
