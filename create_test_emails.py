#!/usr/bin/env python3
“””
create_test_emails.py - Creates test emails in Outlook from CSV using Microsoft Graph API
Reads the test email CSV and creates draft messages in the specified folder
“””

import os
import csv
import requests
import msal
from datetime import datetime

# ========= CONFIG (use same as graph_demo.py) =========

CLIENT_ID = “0ade3d5c-b527-46ad-adac-af00003a111b”
AUTHORITIES = [
“https://login.microsoftonline.com/consumers”,
“https://login.microsoftonline.com/common”,
]
SCOPES = [“User.Read”, “Mail.ReadWrite”]
GRAPH = “https://graph.microsoft.com/v1.0”
DEMO_FOLDER_NAME = “DEMO for PNC”

# Path to your CSV file

CSV_FILE = “test_emails.csv”  # Change this to your CSV filename

# ==========================

def acquire_token_with_diagnostics():
“”“Same auth function from graph_demo.py”””
last_error_detail = None
for authority in AUTHORITIES:
print(f”Trying authority: {authority}”)
app = msal.PublicClientApplication(client_id=CLIENT_ID, authority=authority)

```
    # Try cached token first
    accounts = app.get_accounts()
    if accounts:
        try:
            result = app.acquire_token_silent(SCOPES, account=accounts[0])
            if result and "access_token" in result:
                print("Got cached token.")
                return result["access_token"]
        except Exception as e:
            print(f"Silent token attempt failed: {e}")

    # Device code flow
    try:
        flow = app.initiate_device_flow(scopes=SCOPES)
    except Exception as e:
        last_error_detail = f"initiate_device_flow exception: {e}"
        print(f"Device flow init error: {e}")
        continue

    if "user_code" not in flow:
        err = flow.get("error") or "unknown_error"
        desc = flow.get("error_description") or "No description"
        last_error_detail = f"{err}: {desc}"
        print(f"Device flow init response error — {err}: {desc}")
        continue

    print(flow["message"])
    result = app.acquire_token_by_device_flow(flow)

    if "access_token" in result:
        print("Token acquired.")
        return result["access_token"]

    err = result.get("error")
    desc = result.get("error_description")
    last_error_detail = f"{err}: {desc}"
    print(f"Device flow acquisition failed — {err}: {desc}")

raise RuntimeError(f"Failed to acquire token. Details: {last_error_detail}")
```

def find_mail_folder_id(headers, display_name):
“”“Find folder ID by name”””
safe = display_name.replace(”’”, “’’”)
r = requests.get(
f”{GRAPH}/me/mailFolders”,
headers=headers,
params={”$filter”: f”displayName eq ‘{safe}’”, “$top”: 10}
)
if not r.ok:
raise RuntimeError(f”Failed to find folder: {r.text}”)
vals = r.json().get(“value”, [])
return vals[0][“id”] if vals else None

def create_email_in_folder(headers, folder_id, email_data):
“”“Create a draft email in the specified folder”””
# Build the message payload
message = {
“subject”: email_data[‘subject’],
“body”: {
“contentType”: “Text”,
“content”: email_data[‘body’]
},
“from”: {
“emailAddress”: {
“address”: email_data[‘from_email’],
“name”: email_data[‘from_name’]
}
},
“receivedDateTime”: email_data[‘date’],
“isDraft”: False  # Make it look like a received message
}

```
# Create the message
r = requests.post(
    f"{GRAPH}/me/mailFolders/{folder_id}/messages",
    headers=headers,
    json=message
)

if r.ok:
    return r.json()
else:
    print(f"Failed to create email: {email_data['subject']}")
    print(f"Error: {r.status_code} - {r.text}")
    return None
```

def parse_csv_and_create_emails(csv_path, headers, folder_id):
“”“Read CSV and create all test emails”””
created_count = 0
failed_count = 0

```
with open(csv_path, 'r', encoding='utf-8') as f:
    reader = csv.DictReader(f)
    
    for row in reader:
        email_data = {
            'date': row['Date'] + 'T09:00:00Z',  # Add time component
            'from_email': row['From Email'],
            'from_name': row['From Name'],
            'subject': row['Subject'],
            'body': row['Email Body Preview']
        }
        
        print(f"Creating: {email_data['subject'][:50]}...")
        result = create_email_in_folder(headers, folder_id, email_data)
        
        if result:
            created_count += 1
            print(f"  ✓ Created successfully")
        else:
            failed_count += 1
            print(f"  ✗ Failed")

return created_count, failed_count
```

def main():
print(“Starting test email creation…\n”)

```
# Check if CSV exists
if not os.path.exists(CSV_FILE):
    print(f"ERROR: CSV file '{CSV_FILE}' not found!")
    print("Please save your CSV file and update the CSV_FILE path in the script.")
    return

# Authenticate
token = acquire_token_with_diagnostics()
headers = {"Authorization": f"Bearer {token}"}

# Find the folder
folder_id = find_mail_folder_id(headers, DEMO_FOLDER_NAME)
if not folder_id:
    print(f"\nERROR: Folder '{DEMO_FOLDER_NAME}' not found!")
    print("Please create this folder in Outlook first.")
    return

print(f"\nFound folder: {DEMO_FOLDER_NAME}")
print(f"Reading CSV: {CSV_FILE}\n")

# Create all emails
created, failed = parse_csv_and_create_emails(CSV_FILE, headers, folder_id)

print(f"\n{'='*50}")
print(f"COMPLETE!")
print(f"  Created: {created} emails")
print(f"  Failed:  {failed} emails")
print(f"{'='*50}")
```

if **name** == “**main**”:
main()
