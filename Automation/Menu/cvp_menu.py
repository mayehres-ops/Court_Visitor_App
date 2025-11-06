# cvp_menu.py  — Court Visitor Program Control Panel (GUI)
# Windows-only Tkinter launcher for your existing scripts/BATs
# No external dependencies. Python 3.12+ recommended.

import os
import subprocess
import sys
import datetime
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

# ===================== CONFIG =====================
# Log file (create folders if missing)
LOG_PATH = r"C:\GoogleSync\Automation\GuardianAutomation\logs\menu_launcher.log"

# Common working dir for menu (affects relative BAT behavior)
DEFAULT_CWD = r"C:\GoogleSync\Automation"

# Button definitions: (label, tooltip, command, working_dir)
# Uses YOUR actual paths from the menu doc wherever provided.
BUTTONS = [
    # 1 OCR PDFs and place new line in excel
    ("🧾  OCR → Excel", 
     "Run OCR on PDFs and update Excel (guardian_extractor)",
     r'py -3 "C:\GuardianAutomation\guardian_extractor_bestYet_test.py"', 
     r"C:\GuardianAutomation"),

    # 2 Make folders & move files
    ("🗂️  Build Folders", 
     "Create Court Visitor folders and move files",
     r'py -3 "C:\GoogleSync\Automation\CV Report_Folders Script\scripts\cvr_folder_builder.py"',
     r"C:\GoogleSync\Automation\CV Report_Folders Script\scripts"),

    # 3 Create Court Visitor Word form
    ("📄  Build CV Report", 
     "Create Court Visitor Word form and move to folder",
     r'"C:\GoogleSync\Automation\Create CV report_move to folder\Scripts\Run CVR (Real).bat"',
     r"C:\GoogleSync\Automation\Create CV report_move to folder\Scripts"),

    # 4 Create / Print Map
     ("🗺️  Build Map Sheet",
     "Generate Ward Map Sheet (wrapper auto-detects Python)",
     r'"C:\GoogleSync\Automation\Build Map Sheet\Run Ward Map Sheet.bat"',
     r"C:\GoogleSync\Automation\Build Map Sheet"),

    ("🖨️  Print Map Only", 
     "Print the Ward Map Sheet only",
     r'"C:\GoogleSync\Automation\Build Map Sheet\Print Ward Map Sheet Only.bat"',
     r"C:\GoogleSync\Automation\Build Map Sheet"),

    # 5 Meeting request email
    ("📧  Send Meeting Emails", 
     "Initial guardian meeting request emails (Confirm run)",
     r'"C:\GoogleSync\Automation\Email Meeting Request\scripts\Send Meeting Emails (Confirm).bat"',
     r"C:\GoogleSync\Automation\Email Meeting Request\scripts"),

    # 6 Calendar after date entered
    ("📅  Create Calendar Appt", 
      "Create calendar events after date set in Excel",
      r'"C:\GoogleSync\Automation\Calendar appt send email conf\scripts\run create_calendar_event.bat"',
      r"C:\GoogleSync\Automation\Calendar appt send email conf\scripts"),

    # 7 Summary sheet picker
    ("📝  Summary (Save Only)", 
     "Pick rows → save summaries → open folder",
     r'"C:\GoogleSync\Automation\Court Visitor Summary\Run_Summary.bat"',
     r"C:\GoogleSync\Automation\Court Visitor Summary"),
    ("🖨️  Summary (Save + Print)", 
     "Pick rows → save & print summaries → open folder",
     r'"C:\GoogleSync\Automation\Court Visitor Summary\Run_Summary_Print.bat"',
     r"C:\GoogleSync\Automation\Court Visitor Summary"),

    # 8 Appointment confirmation emails
    ("✅  Send Appt Confirms (All)", 
     "One-click live: send confirmations for all eligible rows",
     r'"C:\GoogleSync\Automation\Appt Email Confirm\scripts\send_confirmation_email all.bat"',
     r"C:\GoogleSync\Automation\Appt Email Confirm\scripts"),
    ("☑️  Appt Confirm (Menu Pick)", 
     "Interactive menu to pick confirmation options",
     r'"C:\GoogleSync\Automation\Appt Email Confirm\scripts\send_confirmation_email_option.bat"',
     r"C:\GoogleSync\Automation\Appt Email Confirm\scripts"),

    # 9 Add guardians to contacts
    ("👥  Add to Contacts", 
     "Add/format guardians in Contacts (with safe upserts)",
     r'"C:\GoogleSync\Automation\Contacts - Guardians\scripts\contacts_menu.bat"',
     r"C:\GoogleSync\Automation\Contacts - Guardians\scripts"),

    # 10 Follow-ups / Thank you
    ("🙏  Send Follow-ups", 
     "Send thank-you / follow-up emails (picker)",
     r'"C:\GoogleSync\Automation\TX email to guardian\Send Followups (Live).bat"',
     r"C:\GoogleSync\Automation\TX email to guardian"),

    # 11 Mileage form  — per your note “put first choice Mileage form…”
    ("🚗  Mileage Form", 
     "Build mileage reimbursement forms from visits",
     r'"C:\GoogleSync\Automation\Mileage Reimbursement Script\scripts\run_mileage_forms.py"',
     r"C:\GoogleSync\Automation\Mileage Reimbursement Script\scripts"),

    # 12 Payment form
    ("💵  Payment Form", 
     "Build Court Visitor payment forms",
     r'"C:\GoogleSync\Automation\CV Payment Form Script\scripts\Run payment Forms.bat"',
     r"C:\GoogleSync\Automation\CV Payment Form Script\scripts"),

    # 13 Send to Probate Court (placeholder)
    ("📨  Submit to Court (TBD)", 
     "Placeholder for future automation to send forms to Probate Court",
     r'cmd /c echo "Submit-to-Court automation not yet implemented."',
     DEFAULT_CWD),
]

# ===================== UTILITIES =====================
def ensure_log_dir():
    os.makedirs(os.path.dirname(LOG_PATH), exist_ok=True)

def log(line: str):
    ensure_log_dir()
    stamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(LOG_PATH, "a", encoding="utf-8") as f:
        f.write(f"[{stamp}] {line}\n")

def run_command(cmd: str, cwd: str = None):
    """
    Run command via cmd.exe so BATs and py work consistently.
    Use 'call' so batch files return to us, and '/wait' semantics.
    """
    try:
        display = cmd
        # Normalize quoting; run through cmd to support .bat, .py, spaces, etc.
        full_cmd = f'cmd /c call {cmd}'
        proc = subprocess.run(full_cmd, cwd=cwd or DEFAULT_CWD, shell=True)
        rc = proc.returncode
        if rc == 0:
            msg = f"✅ Done: {display}"
            log(msg)
            return True, msg
        else:
            msg = f"❌ Exit {rc}: {display}"
            log(msg)
            return False, msg
    except Exception as e:
        msg = f"💥 Exception: {cmd} :: {e}"
        log(msg)
        return False, msg

# ===================== UI =====================
class CVPMenu(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Court Visitor Program — Control Panel")
        self.geometry("880x620")
        self.minsize(760, 520)

        # Tk/ttk styling
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except:
            pass
        style.configure("TButton", padding=10, font=("Segoe UI", 10, "bold"))
        style.configure("Header.TLabel", font=("Segoe UI", 16, "bold"))
        style.configure("Sub.TLabel", font=("Segoe UI", 10))

        # Header
        header = ttk.Label(self, text="Court Visitor Program — Control Panel", style="Header.TLabel")
        header.pack(pady=(14, 2))
        sub = ttk.Label(self, text="Click a task below. Status and errors are logged.", style="Sub.TLabel")
        sub.pack(pady=(0, 10))

        # Content frame with a responsive grid
        content = ttk.Frame(self)
        content.pack(fill="both", expand=True, padx=14, pady=8)
        for i in range(3):
            content.columnconfigure(i, weight=1)

        # Build buttons (3 columns)
        self.status_var = tk.StringVar(value="Ready.")
        for idx, (label, tip, cmd, cwd) in enumerate(BUTTONS):
            col = idx % 3
            row = idx // 3
            btn = ttk.Button(content, text=label, command=lambda c=cmd, d=cwd: self._on_click(c, d))
            btn.grid(row=row, column=col, sticky="nsew", padx=6, pady=6, ipady=6)
            # Hover tooltip (simple)
            self._attach_tooltip(btn, tip)

        # Bottom bar
        sep = ttk.Separator(self, orient="horizontal")
        sep.pack(fill="x", padx=10, pady=(4, 6))

        bottom = ttk.Frame(self)
        bottom.pack(fill="x", padx=12, pady=(0, 10))

        # Utility buttons
        ttk.Button(bottom, text="📂 Open Scripts Folder", command=self._open_scripts_root).pack(side="left", padx=(0, 6))
        ttk.Button(bottom, text="📜 View Logs", command=self._open_log).pack(side="left", padx=6)
        ttk.Button(bottom, text="Exit", command=self.destroy).pack(side="right")

        # Status
        self.status = ttk.Label(self, textvariable=self.status_var, anchor="w")
        self.status.pack(fill="x", padx=14, pady=(0, 10))

        log("=== App started ===")

    def _on_click(self, cmd, cwd):
        self.status_var.set("Running…")
        self.update_idletasks()
        ok, msg = run_command(cmd, cwd)
        self.status_var.set(msg)
        if not ok:
            messagebox.showerror("Task Error", msg)

    def _open_scripts_root(self):
        path = DEFAULT_CWD
        if os.path.isdir(path):
            os.startfile(path)
        else:
            messagebox.showwarning("Not Found", f"Folder not found:\n{path}")

    def _open_log(self):
        ensure_log_dir()
        # Create the log file if it doesn't exist yet
        if not os.path.exists(LOG_PATH):
            with open(LOG_PATH, "w", encoding="utf-8") as f:
                f.write("")
        os.startfile(LOG_PATH)

    # lightweight tooltip
    def _attach_tooltip(self, widget, text):
        tip = tk.Toplevel(widget)
        tip.withdraw()
        tip.overrideredirect(True)
        label = ttk.Label(tip, text=text, background="#FFFFE0", relief="solid", borderwidth=1)
        label.pack(ipadx=6, ipady=3)

        def enter(_):
            tip.deiconify()
            x = widget.winfo_rootx() + 20
            y = widget.winfo_rooty() + 40
            tip.geometry(f"+{x}+{y}")

        def leave(_):
            tip.withdraw()

        widget.bind("<Enter>", enter)
        widget.bind("<Leave>", leave)

if __name__ == "__main__":
    try:
        app = CVPMenu()
        app.mainloop()
    except Exception as e:
        ensure_log_dir()
        import traceback
        tb = traceback.format_exc()
        log(f"Startup failure: {e}\n{tb}")
        try:
            # Show a pop-up even if UI didn't initialize
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror(
                "Startup error",
                f"{e}\n\nSee detailed log at:\n{LOG_PATH}"
            )
        except:
            pass
        print(f"Startup failure: {e}\nSee log: {LOG_PATH}")
        sys.exit(1)

