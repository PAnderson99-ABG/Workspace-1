import os
import re
import json
import base64
import requests
import msal
import sys
from pathlib import Path
from dotenv import load_dotenv

# === Load environment variables ===
load_dotenv()

CLIENT_ID = os.getenv("CLIENT_ID", "").strip("'")
TENANT_ID = os.getenv("TENANT_ID", "").strip("'")
AUTHORITY = os.getenv("AUTHORITY", "").strip("'")
SCOPES = os.getenv("SCOPES", "Mail.Send").strip("'").split()
CACHE_FILE = os.getenv("CACHE_FILE", "").strip("'")


# === Parse metric values from .txt files ===
def parse_metrics_txt(file_paths):
    replacements = {}

    patterns = {
        "sku_total": r"Total SKUs submitted this month:\s*([\d,]+)",
        "status_complete_count": r"Complete:\s*(\d+)\s+rows",
        "status_unspecified_count": r"— Unspecified —:\s*(\d+)\s+rows",
        "status_not_started_count": r"Not Started:\s*(\d+)\s+rows",
        "status_in_progress_count": r"In Progress:\s*(\d+)\s+rows",
        "status_cancelled_count": r"Cancelled:\s*(\d+)\s+rows",
        "status_submitted_count": r"Submitted/Not Online:\s*(\d+)\s+rows",
        "status_on_hold_count": r"On Hold:\s*(\d+)\s+rows",
        "sku_complete": r"Complete:\s*([\d,]+)\s+SKUs",
        "sku_not_started": r"Not Started:\s*([\d,]+)\s+SKUs",
        "sku_in_progress": r"In Progress:\s*([\d,]+)\s+SKUs",
        "sku_cancelled": r"Cancelled:\s*([\d,]+)\s+SKUs",
        "sku_submitted": r"Submitted/Not Online:\s*([\d,]+)\s+SKUs",
        "sku_unspecified": r"— Unspecified —:\s*([\d,]+)\s+SKUs",
        "sku_on_hold": r"On Hold:\s*([\d,]+)\s+SKUs",
        "sku_30_day_complete": r"Completed Data Contracts This Month:\s*(\d+)",
        "sku_onboarded_this_month": r"Total SKUs Onboarded This Month:\s*([\d,]+)"
    }

    for file_path in file_paths:
        path = Path(file_path)
        if not path.exists():
            raise FileNotFoundError(f"❌ File not found: {path}")

        with path.open("r", encoding="utf-8") as f:
            content = f.read()

        for key, pattern in patterns.items():
            if key not in replacements:
                match = re.search(pattern, content)
                if match:
                    replacements[key] = match.group(1).replace(",", "")

    return replacements


# === Load HTML content from file and substitute variables ===
def load_html_body(path="email_body.html", replacements=None):
    if replacements is None:
        replacements = {}

    with open(path, "r", encoding="utf-8") as file:
        body = file.read()

    all_placeholders = set(re.findall(r"\{\{(.*?)\}\}", body))
    unresolved = []

    for key in all_placeholders:
        if key in replacements:
            body = body.replace(f"{{{{{key}}}}}", replacements[key])
        else:
            body = body.replace(f"{{{{{key}}}}}", "n/a")
            unresolved.append(key)

    if unresolved:
        print("⚠️  Missing variables substituted with 'n/a':", unresolved)

    return body


# === Get an access token using MSAL ===
def get_access_token():
    cache = msal.SerializableTokenCache()
    if Path(CACHE_FILE).exists():
        cache.deserialize(Path(CACHE_FILE).read_text(encoding="utf-8"))

    app = msal.PublicClientApplication(
        client_id=CLIENT_ID,
        authority=AUTHORITY,
        token_cache=cache
    )

    accounts = app.get_accounts()
    token = app.acquire_token_silent(SCOPES, account=accounts[0]) if accounts else None

    if not token:
        print("🔐 No cached token found — starting device code flow...")
        flow = app.initiate_device_flow(scopes=SCOPES)
        if "user_code" not in flow:
            raise Exception("Failed to initiate device flow")
        print(flow["message"])
        token = app.acquire_token_by_device_flow(flow)

    if cache.has_state_changed:
        Path(CACHE_FILE).write_text(cache.serialize(), encoding="utf-8")

    if "access_token" not in token:
        raise Exception(f"Token acquisition failed:\n{json.dumps(token, indent=2)}")

    return token["access_token"]


# === Send email via Microsoft Graph API ===
def send_email(subject, body_html, recipients, attachments=None):
    token = get_access_token()

    # Force recipient to yourself for test safety
    to_recipients = [{"emailAddress": {"address": "panderson@americanbathgroup.com"}}]

    attachment_objects = []
    if attachments:
        for path in attachments:
            with open(path, "rb") as file:
                encoded = base64.b64encode(file.read()).decode("utf-8")
            attachment_objects.append({
                "@odata.type": "#microsoft.graph.fileAttachment",
                "name": os.path.basename(path),
                "contentType": "application/octet-stream",
                "contentBytes": encoded
            })

    message = {
        "message": {
            "subject": subject,
            "body": {
                "contentType": "html",
                "content": body_html
            },
            "toRecipients": to_recipients,
            "attachments": attachment_objects
        },
        "saveToSentItems": "true"
    }

    response = requests.post(
        url="https://graph.microsoft.com/v1.0/me/sendMail",
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json"
        },
        data=json.dumps(message)
    )

    if response.status_code == 202:
        print("✅ Email sent successfully.")
    else:
        print(f"❌ Failed to send email: {response.status_code}")
        print("🧾 Response:", response.text[:500])


# === Run the email preparation process ===
if __name__ == "__main__":
    metric_files = [
        r"C:/Users/panderson/OneDrive - American Bath Group/Documents/Paul_Anderson/Reports/PCM_Metrics.txt",
        r"C:/Users/panderson/OneDrive - American Bath Group/Documents/Paul_Anderson/Reports/MDM_DC_Metrics.txt"
    ]

    metrics = parse_metrics_txt(metric_files)
    html_body = load_html_body("email_body.html", replacements=metrics)

    if "--dry-run" in sys.argv:
        print("\n=== DRY RUN: Email Body Preview ===\n")
        print(html_body)
        sys.exit()

    send_email(
        subject="📊 Weekly Product Onboarding Dashboard Summary",
        body_html=html_body,
        recipients=["panderson@americanbathgroup.com"],  # Overridden internally for safety
        attachments=[]  # Optional: Add file paths here
    )
