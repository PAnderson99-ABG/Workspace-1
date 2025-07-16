import subprocess
import os
import sys

# Paths
report_script = "report.py"
trim_script = "trim.py"
home_ss = "amazon_home.png"
dashboard_ss = "amazon_dashboard.png"
cropped_ss = "amazon_home_cropped.png"

# Step 1: Run report.py to generate screenshots
print("▶️ Running report.py to capture screenshots...")
subprocess.run([sys.executable, report_script], check=True)

# Check files
if not os.path.exists(home_ss) or not os.path.exists(dashboard_ss):
    raise FileNotFoundError("❌ Screenshot(s) missing. Ensure login was completed.")

print("✅ Screenshots captured.")

# Step 2: Run trim.py to crop the home screenshot
print("✂️ Running trim.py to crop the Global Snapshot section...")
subprocess.run([sys.executable, trim_script], check=True)

if not os.path.exists(cropped_ss):
    raise FileNotFoundError("❌ Cropped image not found. Check trim.py.")

print("✅ Cropped image saved:", cropped_ss)
