import pandas as pd
import os
from dotenv import load_dotenv

# === Load .env ===
load_dotenv()
file_path = os.getenv("WCWQ1P1")

if not file_path or not os.path.exists(file_path):
    raise FileNotFoundError(f"Missing or invalid path for WCWQ1P1: {file_path}")

# === Read CSV (first row only, suppress dtype warnings) ===
df = pd.read_csv(file_path, nrows=0, low_memory=False)

# === Print column headers ===
print("📄 Columns in file:")
for col in df.columns:
    print(f"- {col}")
