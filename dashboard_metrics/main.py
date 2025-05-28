import subprocess
import sys
from pathlib import Path

# === Configuration ===
scripts_dir = Path(__file__).resolve().parent
reports_dir = Path(r"C:/Users/panderson/OneDrive - American Bath Group/Documents/Reports")
html_template_path = scripts_dir / "email_body.html"
email_script_path = scripts_dir / "send_email.py"

# === Step 1: Run MDM DC Tracker Script ===
print("\n▶️ Running MDM Data Contract Tracker...")
subprocess.run([sys.executable, str(scripts_dir / "dc_tracker.py")], check=True)

# === Step 2: Run PCM Requests Export Script ===
print("\n▶️ Running PCM Requests Export...")
subprocess.run([sys.executable, str(scripts_dir / "pcm_reqs.py")], check=True)

# === Step 3: Run Email Sender Script (with optional --dry-run) ===
print("\n▶️ Preparing and sending summary email...")
subprocess.run([
    sys.executable,
    str(email_script_path),
    # Uncomment next line to preview instead of send
    # "--dry-run"
], check=True)

print("\n✅ All steps completed successfully.")
