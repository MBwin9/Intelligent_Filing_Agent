#!/usr/bin/env python3
"""
graph_demo.py — Personal Outlook.com demo using Microsoft Graph with delegated (device-code) auth.
- Signs you in as /me (device code)
- Finds "DEMO for PNC" folder and pulls messages
- Classifies into Filed / Triage / Skipped
- Renders Tailwind dashboard from dashboard_template.html and auto-opens it
"""

import os, sys, json, html, webbrowser
from datetime import datetime, timezone
from typing import Dict, List

import requests
import msal

# ========= CONFIG =========
CLIENT_ID = "0ade3d5c-b527-46ad-adac-af00003a111b"  # <-- your App Registration's Application (client) ID
AUTHORITIES = [
    "https://login.microsoftonline.com/consumers",  # personal MSA only
    "https://login.microsoftonline.com/common",     # any org + personal
]
SCOPES = ["User.Read", "Mail.ReadWrite"]            # Delegated Graph scopes (do NOT include 'offline_access')

GRAPH = "https://graph.microsoft.com/v1.0"
DEMO_FOLDER_NAME = "DEMO for PNC"
TOP = 50

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_FILE = os.path.join(BASE_DIR, "dashboard_template.html")
OUT_FILE = os.path.join(BASE_DIR, "dashboard.html")
# ==========================


# ---------- HTTP / Graph helpers ----------
def _graph_get(path: str, headers: Dict, params: Dict = None) -> Dict:
    r = requests.get(f"{GRAPH}{path}", headers=headers, params=params)
    if not r.ok:
        raise RuntimeError(f"GET {path} failed ({r.status_code}): {r.text}")
    return r.json()

def _get_me(headers: Dict) -> Dict:
    return _graph_get("/me", headers)

def _find_mail_folder_id(headers: Dict, display_name: str) -> str:
    safe = display_name.replace("'", "''")
    data = _graph_get("/me/mailFolders", headers, params={"$filter": f"displayName eq '{safe}'", "$top": 10})
    vals = data.get("value", [])
    return vals[0]["id"] if vals else None

def _list_messages(headers: Dict, folder_id: str, top: int) -> List[Dict]:
    params = {
        "$top": min(top, 100),
        "$select": "id,subject,from,receivedDateTime,hasAttachments,conversationId",
        "$orderby": "receivedDateTime desc",
    }
    data = _graph_get(f"/me/mailFolders/{folder_id}/messages", headers, params=params)
    return data.get("value", [])


# ---------- Classification & formatting ----------
def _classify(msg: Dict) -> str:
    """
    Demo rules:
      - Filed: subject contains 'quote', 'policy', 'binder', or 'endorsement'
      - Triage: has attachments OR subject contains 'claim'
      - Skipped: everything else
    """
    subject = (msg.get("subject") or "").lower()
    has_attach = bool(msg.get("hasAttachments"))
    if any(k in subject for k in ("quote", "policy", "binder", "endorsement")):
        return "filed"
    if has_attach or "claim" in subject:
        return "triage"
    return "skipped"

def _iso_to_display(s: str) -> str:
    if not s:
        return ""
    try:
        dt = datetime.fromisoformat(s.replace("Z", "+00:00")).astimezone(timezone.utc)
        return dt.strftime("%Y-%m-%d %H:%M UTC")
    except Exception:
        return s


# ---------- Tailwind dashboard renderer (uses your template) ----------
def render_tailwind_dashboard(messages: List[Dict], folder_display: str = DEMO_FOLDER_NAME) -> None:
    """
    messages: list of dicts:
      - status in {"filed","triage","skipped"}
      - subject (str), filed_dir (str), timestamp (str)
    """
    if not os.path.isfile(TEMPLATE_FILE):
        raise FileNotFoundError(f"Template not found: {TEMPLATE_FILE}")

    # Build rows expected by the template/JS counters (needs `.status-badge` with text: Filed/Triage/Skipped)
    rows_html = []
    for m in messages:
        status_label = {"filed": "Filed", "triage": "Triage", "skipped": "Skipped"}[m["status"]]
        rows_html.append(
            "<tr>"
            f"<td class='p-4'><span class='status-badge inline-block px-2 py-1 rounded-full text-xs bg-gray-100'>{html.escape(status_label)}</span></td>"
            f"<td class='p-4'>{html.escape(m.get('subject',''))}</td>"
            f"<td class='p-4'>{html.escape(m.get('filed_dir', folder_display))}</td>"
            f"<td class='p-4'>{html.escape(m.get('timestamp',''))}</td>"
            "</tr>"
        )
    if not rows_html:
        rows_block = "<tr><td colspan='4' class='text-center text-gray-500 py-4'>No emails were processed.</td></tr>"
    else:
        rows_block = "\n".join(rows_html)

    # Inject runtime + rows into template
    with open(TEMPLATE_FILE, "r", encoding="utf-8") as f:
        tpl = f.read()
    tpl = tpl.replace("{RUN_TIME}", datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC"))
    tpl = tpl.replace("<!-- REPORT_DATA -->", rows_block)

    with open(OUT_FILE, "w", encoding="utf-8") as f:
        f.write(tpl)

    webbrowser.open(OUT_FILE)


# ---------- Auth with diagnostics (device code) ----------
def acquire_token_with_diagnostics() -> str:
    """
    Try consumers -> common. If device flow init fails, print detailed hints.
    """
    last_error_detail = None
    for authority in AUTHORITIES:
        print(f"Trying authority: {authority}")
        app = msal.PublicClientApplication(client_id=CLIENT_ID, authority=authority)

        # Silent (cached token) first
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

        print(flow["message"])  # "Open https://www.microsoft.com/link and enter code ..."
        result = app.acquire_token_by_device_flow(flow)

        if "access_token" in result:
            print("Token acquired.")
            return result["access_token"]

        err = result.get("error")
        desc = result.get("error_description")
        last_error_detail = f"{err}: {desc}"
        print(f"Device flow acquisition failed — {err}: {desc}")

    raise RuntimeError(
        "Failed to create/complete device flow. "
        f"Details: {last_error_detail or 'no additional info'}\n\n"
        "Fix checklist:\n"
        "  1) Azure App Registration > Authentication: 'Allow public client flows' = Yes\n"
        "  2) Supported account types include Personal Microsoft accounts (MSA)\n"
        "  3) CLIENT_ID matches the configured app\n"
        "  4) SCOPES are delegated Graph scopes only: ['User.Read', 'Mail.ReadWrite']"
    )


# ---------- Main ----------
def main():
    print("Starting Microsoft Graph Email Filer Demo...")

    token = acquire_token_with_diagnostics()
    headers = {"Authorization": f"Bearer {token}"}

    me = _get_me(headers)
    upn = me.get("userPrincipalName") or me.get("mail")
    print("Signed in as:", upn)

    folder_id = _find_mail_folder_id(headers, DEMO_FOLDER_NAME)
    if not folder_id:
        print(f"ERROR: Folder '{DEMO_FOLDER_NAME}' not found in your mailbox.")
        sys.exit(1)

    msgs = _list_messages(headers, folder_id, TOP)
    print(f"Fetched {len(msgs)} message(s).")

    # Build rows for dashboard
    messages = []
    for m in msgs:
        status = _classify(m)
        messages.append({
            "status": status,
            "subject": m.get("subject") or "",
            "filed_dir": DEMO_FOLDER_NAME,  # for demo we show the folder name; adjust if you sub-route
            "timestamp": _iso_to_display(m.get("receivedDateTime")),
        })

    # Render Tailwind dashboard from your template
    render_tailwind_dashboard(messages, folder_display=DEMO_FOLDER_NAME)
    print("Dashboard saved to:", OUT_FILE)


if __name__ == "__main__":
    main()
