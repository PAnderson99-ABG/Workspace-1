# === load_data.py ===
import os
import pandas as pd
from dotenv import load_dotenv

def load_and_prepare(path, label, skip_header=False):
    print(f"\n📂 Loading file: {label} ({path})")
    df = pd.read_csv(path, skiprows=1 if skip_header else 0, low_memory=False)
    print(f"🔎 {label} original columns: {df.columns.tolist()}")
    df.columns = df.columns.str.strip().str.lower().str.replace(" ", "_")
    print(f"✅ {label} normalized columns: {df.columns.tolist()}")
    return df

def load_data():
    load_dotenv()
    paths = {k: os.getenv(k) for k in ["WCWQ1P1", "WCWQ1P2", "WCWQ2P1", "WCWQ2P2", "WCWQ3P1", "WCWQ3P2"]}
    
    for k, path in paths.items():
        if not path or not os.path.exists(path):
            raise FileNotFoundError(f"Missing or invalid path for {k}: {path}")

    dfs = {label: load_and_prepare(path, label) for label, path in zip(["Q1P1", "Q1P2", "Q2P1", "Q2P2", "Q3P1", "Q3P2"], paths.values())}

    for label, df in dfs.items():
        if "timestamp" not in df.columns:
            print(f"❌ DEBUG: Columns in {label}:", df.columns.tolist())
            raise ValueError(f"Missing 'timestamp' column in file: {label}")
        df["timestamp"] = pd.to_datetime(df["timestamp"], format="%y-%m-%d %H:%M:%S %z", errors="coerce")
        print(f"📅 {label}: Parsed timestamps - valid: {df['timestamp'].notna().sum()}, invalid: {df['timestamp'].isna().sum()}")

    df = pd.concat(dfs.values(), ignore_index=True)

    required_cols = ["property_name", "user_email", "brand", "timestamp"]
    missing = [col for col in required_cols if col not in df.columns]
    if missing:
        raise ValueError(f"Missing required columns in combined dataset: {', '.join(missing)}")

    print("\n🧪 Sample of parsed timestamps:")
    print(df["timestamp"].head(10))

    return df
