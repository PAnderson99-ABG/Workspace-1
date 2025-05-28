# === summarize.py ===
import pandas as pd

def generate_summaries(df):
    top_fields = df["property_name"].value_counts().head(50).reset_index()
    top_fields.columns = ["property_name", "change_count"]

    all_users = df["user_email"].value_counts().reset_index()
    all_users.columns = ["user", "change_count"]

    top_brands = df["brand"].value_counts().head(50).reset_index()
    top_brands.columns = ["brand", "change_count"]

    return top_fields, all_users, top_brands

def generate_breakdowns(df, top_items, group_col):
    breakdowns = []
    for val in top_items[group_col]:
        subset = df[df[group_col] == val]
        top_users = (
            subset["user_email"]
            .value_counts()
            .head(5)
            .reset_index()
            .rename(columns={"index": "user", "user_email": "change_count"})
        )
        top_users.insert(0, group_col, val)
        breakdowns.append(top_users)
    return pd.concat(breakdowns, ignore_index=True)
