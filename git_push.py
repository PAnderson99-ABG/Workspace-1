import subprocess
import sys
from datetime import datetime

# === Build dynamic default message ===
timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
default_message = f"Workspace-1 Updates [{timestamp}]"

# === Use passed message if available ===
commit_message = sys.argv[1] if len(sys.argv) > 1 else default_message

try:
    # Stage all changes
    subprocess.run(["git", "add", "."], check=True)

    # Commit
    subprocess.run(["git", "commit", "-m", commit_message], check=True)

    # Push to origin/main
    subprocess.run(["git", "push", "origin", "main"], check=True)

    print(f"✅ Commit and push successful.\n📝 Message: {commit_message}")
except subprocess.CalledProcessError as e:
    print(f"❌ Git error: {e}")
