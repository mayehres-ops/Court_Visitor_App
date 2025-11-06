#!/usr/bin/env python3
# -*- coding: utf-8 -*-
r"""
Send follow-up emails to guardians using an interactive picker (last N rows) from Excel.
- Read-only access to the workbook (never writes back)
- Shows the last 15 rows in a GUI; you select which to send
- Uses recipient emails from columns: gemail, g2eamil (exact Excel headers)
- DRY_RUN=True saves .eml preview files locally (no Gmail required)

Run (Windows):
  py -3 "C:\GoogleSync\Automation\TX email to guardian\send_followups_picker.py"
"""

import sys
import base64  # kept for future 'send' mode
from email.message import EmailMessage
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd

# ----------------------- CONFIG ----------------------- #
# Dynamic path detection
_script_dir = Path(__file__).parent.parent  # Go up to app root
try:
    sys.path.insert(0, str(_script_dir / "Scripts"))
    from app_paths import get_app_paths
    _app_paths = get_app_paths(str(_script_dir))

    WORKBOOK_PATH = str(_app_paths.EXCEL_PATH)
    EML_OUTPUT_DIR = Path(str(_app_paths.APP_ROOT / "Automation" / "TX email to guardian"))
    _CONFIG_DIR = _app_paths.CONFIG_DIR

    # Gmail config
    GMAIL_API_DIR = Path(str(_CONFIG_DIR / "API"))
    GMAIL_CREDENTIALS = GMAIL_API_DIR / "gmail_oauth_client.json"
    GMAIL_TOKEN = GMAIL_API_DIR / "gmail_token.json"

    # Client folders
    GUARDIAN_BASE = Path(str(_app_paths.APP_ROOT))
    NEW_CLIENTS_DIR = GUARDIAN_BASE / "New Clients"
    CORRESPONDENCE_PENDING = GUARDIAN_BASE / "_Correspondence_Pending"

except Exception:
    # Fallback to hardcoded paths
    WORKBOOK_PATH = r"C:\GoogleSync\GuardianShip_App\App Data\ward_guardian_info.xlsx"
    EML_OUTPUT_DIR = Path(r"C:\GoogleSync\Automation\TX email to guardian")

    # Optional Gmail config (used ONLY if DRY_RUN=False)
    GMAIL_API_DIR = Path(r"C:\configlocal\API")
    GMAIL_CREDENTIALS = GMAIL_API_DIR / "gmail_oauth_client.json"
    GMAIL_TOKEN = GMAIL_API_DIR / "gmail_token.json"

    # Client folders for saving text copies
    GUARDIAN_BASE = Path(r"C:\GoogleSync\GuardianShip_App")
    NEW_CLIENTS_DIR = GUARDIAN_BASE / "New Clients"
    CORRESPONDENCE_PENDING = GUARDIAN_BASE / "_Correspondence_Pending"

EML_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

SHEET_NAME    = "Sheet1"
PICK_LAST_N   = 15
DATE_FORMAT   = "%B %d, %Y"  # e.g., August 14, 2025

# Exact recipient columns present in your workbook
RECIPIENT_COLUMNS = ["gemail", "g2eamil"]

# Display columns for the picker label
DISPLAY_COLUMNS = [
    "causeno", "wardfirst", "wardlast", "visitdate",
    "guardian1", "Guardian2", "gemail", "g2eamil"
]

# Email templates
SUBJECT_TEMPLATE = "Thank you for your time regarding {wardfirst}"
BODY_TEMPLATE = (
    "Dear Guardian,\n\n"
    "I wanted to thank you for your time and input during my visit with {wardfirst} {wardlast} on {visitdate}. "
    "I appreciate your cooperation and the valuable information you provided regarding care.\n\n"
    "Your input is a critical part of the required annual review, and I am grateful for your continued support. "
    "It was a pleasure talking to you and getting to know {wardfirst}.\n\n"
    "Wishing you and {wardfirst} all the best for the year ahead.\n\n"
    "Respectfully,\n"
    "May Ehresman\n"
    "Court Visitor\n"
)

# Preview only; does NOT send anything while True
DRY_RUN = False
# ------------------------------------------------------ #

# --- Gmail (loaded only if DRY_RUN=False) --- #
def build_gmail_service():
    """Build Gmail API client lazily so DRY_RUN requires no gmail packages."""
    from googleapiclient.discovery import build
    from google_auth_oauthlib.flow import InstalledAppFlow
    import google.oauth2.credentials
    from google.auth.transport.requests import Request as GRequest

    SCOPES = ["https://www.googleapis.com/auth/gmail.send"]

    creds = None
    if GMAIL_TOKEN.exists():
        creds = google.oauth2.credentials.Credentials.from_authorized_user_file(str(GMAIL_TOKEN), SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(GRequest())
            except Exception as e:
                # Token refresh failed (expired/revoked) - delete and re-authenticate
                print(f"Token refresh failed: {e}")
                print("Deleting expired token and starting fresh OAuth flow...")
                if GMAIL_TOKEN.exists():
                    GMAIL_TOKEN.unlink()
                creds = None  # Force re-auth below

        if not creds:
            flow = InstalledAppFlow.from_client_secrets_file(str(GMAIL_CREDENTIALS), SCOPES)
            creds = flow.run_local_server(port=0)
            GMAIL_API_DIR.mkdir(parents=True, exist_ok=True)
            with open(GMAIL_TOKEN, "w", encoding="utf-8") as token:
                token.write(creds.to_json())
    return build("gmail", "v1", credentials=creds)

def send_via_gmail_api(service, msg: EmailMessage):
    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    return service.users().messages().send(userId="me", body={"raw": raw}).execute()

# --- EML preview --- #
def save_eml_preview(msg: EmailMessage, basename: str):
    """Save an .eml preview file (works with Outlook/Thunderbird)."""
    safe = "".join(c if c.isalnum() or c in ("-", "_", ".") else "_" for c in basename)
    out_path = EML_OUTPUT_DIR / f"{safe}.eml"
    with open(out_path, "w", encoding="utf-8", newline="\n") as f:
        f.write(msg.as_string())
    return out_path

# --- Text copy saving --- #
def find_case_folder(base_dir, cause_no):
    """Find client folder by cause number"""
    try:
        for name in base_dir.iterdir():
            if name.is_dir() and str(cause_no) in name.name:
                return name
    except FileNotFoundError:
        pass
    return None

def save_text_copy(folder, filename, subject, body, to_emails):
    """Save text copy to client folder or _Correspondence_Pending"""
    folder.mkdir(parents=True, exist_ok=True)
    out_path = folder / filename
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(f"TO: {to_emails}\n")
        f.write(f"SUBJECT: {subject}\n\n")
        f.write(body)
    return out_path

# --- Excel helpers --- #
def load_last_rows(path: str, sheet: str, last_n: int) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=sheet)
    if "visitdate" in df.columns:
        df["visitdate"] = pd.to_datetime(df["visitdate"], errors="coerce")
    return df.tail(last_n).copy()

def row_to_display_str(row: pd.Series) -> str:
    parts = []
    for col in DISPLAY_COLUMNS:
        if col in row and pd.notna(row[col]):
            val = row[col]
            if isinstance(val, pd.Timestamp):
                val = val.date().isoformat()
            parts.append(f"{col}={val}")
    return " | ".join(parts)

def extract_recipients(row: pd.Series) -> list:
    emails = []
    for col in RECIPIENT_COLUMNS:
        if col in row and pd.notna(row[col]):
            e = str(row[col]).strip()
            if "@" in e:
                emails.append(e.lower())
    return sorted(set(emails))

def format_visit_date(val) -> str:
    if pd.isna(val):
        return ""
    if isinstance(val, pd.Timestamp):
        return val.strftime(DATE_FORMAT)
    try:
        dt = pd.to_datetime(val, errors="coerce")
        return dt.strftime(DATE_FORMAT) if pd.notna(dt) else str(val)
    except Exception:
        return str(val)

# --- GUI Picker --- #
class Picker(tk.Tk):
    def __init__(self, df_tail: pd.DataFrame):
        super().__init__()
        self.title("Select follow-ups to send")
        self.geometry("1100x500")
        self.df_tail = df_tail.reset_index(drop=True)
        self.vars = []

        info = ttk.Label(self, text=f"Showing last {len(self.df_tail)} rows from {Path(WORKBOOK_PATH).name} — check the ones to send.")
        info.pack(anchor="w", pady=(8, 4), padx=10)

        frame = ttk.Frame(self)
        frame.pack(fill="both", expand=True, padx=10, pady=6)

        canvas = tk.Canvas(frame)
        scroll_y = ttk.Scrollbar(frame, orient="vertical", command=canvas.yview)
        inner = ttk.Frame(canvas)
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.configure(yscrollcommand=scroll_y.set)

        for _, row in self.df_tail.iterrows():
            var = tk.BooleanVar(value=False)
            self.vars.append(var)
            label = row_to_display_str(row)
            ttk.Checkbutton(inner, text=label, variable=var).pack(anchor="w", padx=4, pady=2)

        canvas.pack(side="left", fill="both", expand=True)
        scroll_y.pack(side="right", fill="y")

        btns = ttk.Frame(self)
        btns.pack(fill="x", padx=10, pady=8)
        ttk.Button(btns, text="Select all", command=self.select_all).pack(side="left")
        ttk.Button(btns, text="Clear all", command=self.clear_all).pack(side="left", padx=(6,0))
        ttk.Button(btns, text="Send", command=self.on_send).pack(side="right")
        ttk.Button(btns, text="Cancel", command=self.destroy).pack(side="right", padx=(0,6))

        self.selected_indices = None

    def select_all(self):
        for v in self.vars:
            v.set(True)

    def clear_all(self):
        for v in self.vars:
            v.set(False)

    def on_send(self):
        self.selected_indices = [i for i, v in enumerate(self.vars) if v.get()]
        if not self.selected_indices:
            messagebox.showwarning("No selection", "Please select at least one row to proceed.")
            return
        self.destroy()

# --- Build & (Preview/Send) --- #
def create_email_message(to_addrs, subject, body_text):
    msg = EmailMessage()
    msg["To"] = ", ".join(sorted(set(to_addrs)))
    msg["Subject"] = subject
    msg.set_content(body_text)
    return msg

def build_and_send(rows: pd.DataFrame):
    sent = 0
    service = None
    if not DRY_RUN:
        service = build_gmail_service()

    for _, row in rows.iterrows():
        recipients = extract_recipients(row)
        if not recipients:
            print("[SKIP] No recipient emails found in gemail/g2eamil for:", row_to_display_str(row))
            continue

        visitdate_fmt = format_visit_date(row.get("visitdate", ""))
        subject = SUBJECT_TEMPLATE.format(
            wardfirst=row.get("wardfirst", ""),
            wardlast=row.get("wardlast", ""),
        )
        body = BODY_TEMPLATE.format(
            wardfirst=row.get("wardfirst", ""),
            wardlast=row.get("wardlast", ""),
            visitdate=visitdate_fmt,
        )

        msg = create_email_message(recipients, subject, body)
        if "causeno" in row and pd.notna(row["causeno"]):
            msg["X-Cause-Number"] = str(row["causeno"])  # traceability

        if DRY_RUN:
            base = f"followup_{row.get('causeno','NA')}_{row.get('wardlast','')}_{row.get('wardfirst','')}"
            out = save_eml_preview(msg, base)
            print(f"[DRY-RUN] Preview saved: {out}")
            sent += 1
        else:
            try:
                resp = send_via_gmail_api(service, msg)
                print(f"[SENT] to {msg['To']}: Gmail ID {resp.get('id')}")
                sent += 1

                # Save text copy to client folder
                cause_no = row.get("causeno", "")
                ward_last = row.get("wardlast", "Unknown")
                ward_first = row.get("wardfirst", "Unknown")

                case_folder = find_case_folder(NEW_CLIENTS_DIR, cause_no) if cause_no else None
                if case_folder:
                    save_folder = case_folder
                else:
                    save_folder = CORRESPONDENCE_PENDING

                from datetime import datetime
                txt_filename = f"Followup Email - {ward_last}, {ward_first} - {datetime.now().strftime('%Y-%m-%d')}.txt"
                txt_path = save_text_copy(save_folder, txt_filename, subject, body, msg['To'])
                print(f"[SAVED] Text copy: {txt_path}")

            except Exception as e:
                print(f"[ERROR] Sending failed for {msg['To']}: {e}")

    print(f"Done. {'Prepared' if DRY_RUN else 'Sent'} {sent} message(s).")

def main():
    if not Path(WORKBOOK_PATH).exists():
        print(f"Workbook not found: {WORKBOOK_PATH}")
        sys.exit(2)

    df_tail = load_last_rows(WORKBOOK_PATH, SHEET_NAME, PICK_LAST_N)
    app = Picker(df_tail)
    app.mainloop()

    if app.selected_indices is None:
        print("Cancelled.")
        return

    chosen = df_tail.iloc[app.selected_indices]
    print(f"Selected {len(chosen)} row(s). Proceeding… DRY_RUN={DRY_RUN}")
    build_and_send(chosen)

if __name__ == "__main__":
    import sys
    import traceback
    try:
        main()
        print("\n[OK] Follow-up Emails SUCCESS")
        sys.exit(0)
    except Exception as e:
        print(f"\n[FAIL] Follow-up Emails FAILED: {e}")
        traceback.print_exc()
        sys.exit(1)
