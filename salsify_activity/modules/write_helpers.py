# === write_helpers.py ===

def write_breakdown_table(ws, grouped_df, group_col):
    row_cursor = 0
    headers = grouped_df.columns.tolist()

    for group_val in grouped_df[group_col].unique():
        subset = grouped_df[group_col] == group_val
        rows = grouped_df[subset].reset_index(drop=True)

        for col_idx, header in enumerate(headers):
            ws.write(row_cursor, col_idx, header)

        for r, (_, row) in enumerate(rows.iterrows(), start=1):
            for c, value in enumerate(row):
                ws.write(row_cursor + r, c, value)

        row_cursor += len(rows) + 3


def autofit_columns(ws, dataframe, start_row=0, start_col=0):
    for i, col in enumerate(dataframe.columns):
        max_len = max(
            [len(str(val)) for val in dataframe[col].astype(str)] + [len(col)]
        )
        ws.set_column(start_col + i, start_col + i, max_len + 2)
