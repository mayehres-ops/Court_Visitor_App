#!/usr/bin/env python3
"""
Court Visitor App
Copyright (c) 2024 GuardianShip Easy, LLC. All rights reserved.

This software and associated documentation files are proprietary and confidential.
Unauthorized copying, distribution, or modification is strictly prohibited.

GuardianShip Easy App - Clean Desktop Application
Calls existing scripts without modifying them
"""

# Application version
__version__ = "1.0.0"

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import os
import sys
import subprocess
import threading
import queue
from pathlib import Path
from datetime import datetime

# Import auto-updater
try:
    from auto_updater import AutoUpdater
    AUTO_UPDATE_ENABLED = True
except ImportError:
    AUTO_UPDATE_ENABLED = False
    print("Auto-updater not available")

# Dynamic path management - works from any installation location
sys.path.insert(0, str(Path(__file__).parent / "Scripts"))
try:
    from app_paths import get_app_paths
    APP_PATHS = get_app_paths()
    PATHS_AVAILABLE = True
except ImportError:
    PATHS_AVAILABLE = False
    print("Warning: Dynamic paths not available, using fallback")

# Windows-specific startup info to hide console windows completely
if sys.platform == 'win32':
    from subprocess import STARTUPINFO, STARTF_USESHOWWINDOW

    def get_subprocess_startupinfo():
        """Get Windows STARTUPINFO to hide console windows (prevents Tesseract/OCR blinking)"""
        startupinfo = STARTUPINFO()
        startupinfo.dwFlags |= STARTF_USESHOWWINDOW
        startupinfo.wShowWindow = 0  # SW_HIDE
        return startupinfo
else:
    def get_subprocess_startupinfo():
        """Return None on non-Windows systems"""
        return None

class GuardianShipApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Court Visitor Program")

        # Set window size and position
        window_width = 1400
        window_height = 850
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")
        self.root.minsize(1200, 700)

        # Paths to existing scripts (NO MODIFICATIONS TO SCRIPTS)
        # Use dynamic paths if available, otherwise fallback to hardcoded
        if PATHS_AVAILABLE:
            self.base_dir = APP_PATHS.APP_ROOT
            self.guardian_extractor = APP_PATHS.OCR_SCRIPT
            self.new_files_dir = APP_PATHS.NEW_FILES_DIR
            self.excel_path = APP_PATHS.EXCEL_PATH
        else:
            # Fallback to hardcoded paths (for development)
            self.base_dir = Path(r"C:\GoogleSync\GuardianShip_App")
            self.guardian_extractor = self.base_dir / "guardian_extractor_claudecode20251023_bestever_11pm.py"
            self.new_files_dir = self.base_dir / "New Files"
            self.excel_path = self.base_dir / "App Data" / "ward_guardian_info.xlsx"

        # Script paths for all automation steps
        if PATHS_AVAILABLE:
            self.scripts = {
                2: APP_PATHS.FOLDER_BUILDER_SCRIPT,
                3: APP_PATHS.MAP_SHEET_SCRIPT,
                4: APP_PATHS.EMAIL_REQUEST_SCRIPT,
                5: APP_PATHS.ADD_CONTACTS_SCRIPT,
                6: APP_PATHS.CONFIRMATION_EMAIL_SCRIPT,
                7: APP_PATHS.CALENDAR_EVENT_SCRIPT,
                8: APP_PATHS.CVR_BUILDER_SCRIPT,
                9: self.base_dir / "Automation" / "Court Visitor Summary" / "build_court_visitor_summary.py",
                10: APP_PATHS.GOOGLE_SHEETS_SCRIPT,
                11: APP_PATHS.FOLLOWUP_EMAIL_SCRIPT,
                12: APP_PATHS.EMAIL_CVR_SCRIPT,
                13: APP_PATHS.PAYMENT_FORM_SCRIPT,
                14: APP_PATHS.MILEAGE_LOG_SCRIPT
            }
        else:
            # Fallback to hardcoded paths
            self.scripts = {
                2: self.base_dir / "Automation" / "CV Report_Folders Script" / "scripts" / "cvr_folder_builder.py",
                3: self.base_dir / "Automation" / "Build Map Sheet" / "Scripts" / "build_map_sheet.py",
                4: self.base_dir / "Automation" / "Email Meeting Request" / "scripts" / "send_guardian_emails.py",
                5: self.base_dir / "Automation" / "Contacts - Guardians" / "scripts" / "add_guardians_to_contacts.py",
                6: self.base_dir / "Automation" / "Appt Email Confirm" / "scripts" / "send_confirmation_email.py",
                7: self.base_dir / "Automation" / "Calendar appt send email conf" / "scripts" / "create_calendar_event.py",
                8: self.base_dir / "Automation" / "Create CV report_move to folder" / "Scripts" / "build_cvr_from_excel_cc_working.py",
                9: self.base_dir / "Automation" / "Court Visitor Summary" / "build_court_visitor_summary.py",
                10: self.base_dir / "google_sheets_cvr_integration_fixed.py",
                11: self.base_dir / "Automation" / "TX email to guardian" / "send_followups_picker.py",
                12: self.base_dir / "email_cvr_to_supervisor.py",
                13: self.base_dir / "Automation" / "CV Payment Form Script" / "scripts" / "build_payment_forms_sdt.py",
                14: self.base_dir / "Automation" / "Mileage Reimbursement Script" / "scripts" / "build_mileage_forms.py"
            }

        # State
        self.is_processing = False
        self.is_opening_excel = False

        # API setup status mapping
        self.api_requirements = {
            1: 'vision',      # OCR
            3: 'maps',        # Route Map (optional)
            4: 'gmail',       # Meeting Requests
            5: 'people',      # Contacts
            6: 'gmail',       # Confirmations
            7: 'calendar',    # Calendar Events
            10: 'sheets',     # CVR with Form Data
            11: 'gmail',      # Follow-ups
            12: 'gmail',      # Email CVR
            # Steps 2, 8, 9, 13, 14 have no API requirements
        }

        # Ensure required folders exist
        (self.base_dir / "App Data" / "Inbox").mkdir(parents=True, exist_ok=True)
        (self.base_dir / "New Files").mkdir(parents=True, exist_ok=True)
        (self.base_dir / "Completed").mkdir(parents=True, exist_ok=True)
        (self.base_dir / "App Data" / "Backup").mkdir(parents=True, exist_ok=True)
        (self.base_dir / "App Data" / "Staging").mkdir(parents=True, exist_ok=True)
        (self.base_dir / "App Data" / "Templates").mkdir(parents=True, exist_ok=True)

        # Setup UI
        self.setup_styling()
        self.create_menu_bar()
        self.create_header()
        self.create_sidebar()
        self.create_main_content()
        self.create_status_bar()

        # Check if first run and show setup wizard
        self.root.after(500, self.check_first_run_setup)

    def setup_styling(self):
        """Configure UI styling"""
        style = ttk.Style()
        style.theme_use('clam')

        # Colors
        self.colors = {
            'bg': '#f0f4f8',
            'header_bg': '#1e3a8a',
            'header_fg': 'white',
            'card_bg': 'white',
            'success': '#10b981',
            'warning': '#f59e0b',
            'error': '#ef4444',
            'primary': '#3b82f6'
        }

        self.root.configure(bg=self.colors['bg'])

        # Button styles
        style.configure('Primary.TButton',
                       background=self.colors['primary'],
                       foreground='white',
                       padding=(20, 10),
                       font=('Segoe UI', 11, 'bold'))

        style.map('Primary.TButton',
                 background=[('active', '#2563eb')])

        # Frame styles
        style.configure('Card.TFrame',
                       background=self.colors['card_bg'],
                       relief='raised',
                       borderwidth=1)

        style.configure('TLabel', background=self.colors['card_bg'])
        style.configure('Header.TLabel',
                       font=('Segoe UI', 16, 'bold'),
                       background=self.colors['card_bg'])

    def create_menu_bar(self):
        """Create application menu bar"""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)

        # Help menu
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="About Court Visitor App", command=self.show_about)
        help_menu.add_command(label="View License Agreement", command=self.show_license)
        help_menu.add_separator()
        help_menu.add_command(label="Documentation", command=self.open_documentation)

    def create_header(self):
        """Create application header"""
        header = tk.Frame(self.root, bg=self.colors['header_bg'], height=80)
        header.pack(fill='x', side='top')
        header.pack_propagate(False)

        title = tk.Label(header,
                        text="Court Visitor App",
                        font=('Segoe UI', 24, 'bold'),
                        fg=self.colors['header_fg'],
                        bg=self.colors['header_bg'])
        title.pack(pady=20)

        # Settings button in top-right corner
        settings_btn = tk.Button(header,
                                text="⚙️ Settings",
                                font=('Segoe UI', 10),
                                bg=self.colors['primary'],
                                fg='white',
                                padx=15,
                                pady=5,
                                cursor='hand2',
                                command=self.open_settings)
        settings_btn.place(relx=1.0, rely=0.5, anchor='e', x=-20)

    def create_sidebar(self):
        """Create right sidebar with utility buttons in 2 columns"""
        # Wider sidebar container for 2-column layout (NO scrollbar!)
        sidebar = tk.Frame(self.root, bg='#f8fafc', width=400, relief='sunken', borderwidth=1)
        sidebar.pack(side='right', fill='y', padx=(0, 0), pady=0)
        sidebar.pack_propagate(False)

        # Sidebar title
        sidebar_title = tk.Label(sidebar,
                                text="Quick Access",
                                font=('Segoe UI', 12, 'bold'),
                                bg='#f8fafc',
                                fg='#1e293b')
        sidebar_title.pack(pady=(20, 15), padx=10)

        # Container for 2-column button layout
        buttons_container = tk.Frame(sidebar, bg='#f8fafc')
        buttons_container.pack(fill='both', expand=True, padx=5, pady=5)

        # Utility buttons - REORGANIZED: Most used at top, API setup at bottom
        buttons = [
            ('📖 Getting Started', self.show_getting_started),
            ('🤖 Ask Chatbot', self.show_chatbot),  # FEATURED: Fun & helpful!
            ('💾 Backup My Data', self.backup_my_data),  # USER DATA BACKUP
            ('', None),  # Separator
            ('📊 Excel File', self.open_excel),
            ('📁 New Clients', self.open_new_clients),
            ('📂 New Files', self.open_new_files),
            ('', None),  # Separator
            ('💬 HELP & SUPPORT', None),  # Section header
            ('❓ Quick Help', self.show_help),
            ('📖 Manual', self.show_manual),
            ('🆘 Live Tech Support', self.open_ai_help),
            ('', None),  # Separator
            ('👥 Contacts', self.open_contacts),
            ('📧 Email', self.open_email),
            ('🐛 Report Bug', self.report_bug),
            ('💡 Request Feature', self.request_feature),
            ('', None),  # Separator
            ('⚙️ API SETUP (One-time)', None),  # Section header - moved to bottom!
            ('🔧 Setup Vision API', self.setup_google_vision),
            ('🗺️ Setup Maps API', self.setup_google_maps),
            ('📧 Setup Gmail API', self.setup_gmail_api),
            ('👥 Setup People/Calendar', self.setup_people_calendar_api),
        ]

        # Create buttons in 2-column grid layout
        row = 0
        col = 0

        for text, command in buttons:
            if text == '':
                # Separator spans both columns - finish current row first
                if col != 0:
                    row += 1
                    col = 0
                separator = tk.Frame(buttons_container, bg='#cbd5e1', height=1)
                separator.grid(row=row, column=0, columnspan=2, sticky='ew', pady=5, padx=5)
                row += 1
                col = 0
            elif command is None:
                # Section header spans both columns - finish current row first
                if col != 0:
                    row += 1
                    col = 0
                header = tk.Label(buttons_container,
                                text=text,
                                font=('Segoe UI', 9, 'bold'),
                                bg='#f8fafc',
                                fg='#64748b',
                                anchor='w')
                header.grid(row=row, column=0, columnspan=2, sticky='ew', pady=(8, 2), padx=5)
                row += 1
                col = 0
            else:
                # Buttons alternate between columns
                if '🤖' in text:
                    # Special styling for chatbot button (spans both columns!)
                    btn = tk.Button(buttons_container,
                                  text=text,
                                  command=command,
                                  font=('Segoe UI', 10, 'bold'),
                                  bg='#7c3aed',
                                  fg='white',
                                  relief='raised',
                                  borderwidth=2,
                                  padx=10,
                                  pady=8,
                                  cursor='hand2',
                                  anchor='w')
                    btn.grid(row=row, column=0, columnspan=2, sticky='ew', padx=5, pady=3)

                    def on_enter(e, button=btn):
                        button['bg'] = '#6d28d9'
                    def on_leave(e, button=btn):
                        button['bg'] = '#7c3aed'
                    btn.bind('<Enter>', on_enter)
                    btn.bind('<Leave>', on_leave)
                    row += 1
                    col = 0
                else:
                    # Regular button in grid
                    btn = tk.Button(buttons_container,
                                  text=text,
                                  command=command,
                                  font=('Segoe UI', 9),
                                  bg='white',
                                  fg='#1e293b',
                                  relief='raised',
                                  borderwidth=1,
                                  padx=8,
                                  pady=6,
                                  cursor='hand2',
                                  anchor='w',
                                  wraplength=180)
                    btn.grid(row=row, column=col, sticky='ew', padx=3, pady=2)

                    btn.bind('<Enter>', lambda e, b=btn: b.config(bg='#e0f2fe'))
                    btn.bind('<Leave>', lambda e, b=btn: b.config(bg='white'))

                    # Move to next column or next row
                    col += 1
                    if col > 1:
                        col = 0
                        row += 1

        # Configure columns to expand equally
        buttons_container.grid_columnconfigure(0, weight=1)
        buttons_container.grid_columnconfigure(1, weight=1)

    def create_main_content(self):
        """Create main content area with 14 steps"""
        # Main container
        main_container = tk.Frame(self.root, bg=self.colors['bg'])
        main_container.pack(fill='both', expand=True, padx=20, pady=20)

        # Add scrollbar for 2-column layout
        canvas = tk.Canvas(main_container, bg=self.colors['bg'], highlightthickness=0)
        scrollbar = ttk.Scrollbar(main_container, orient='vertical', command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg=self.colors['bg'])

        scrollable_frame.bind(
            '<Configure>',
            lambda e: canvas.configure(scrollregion=canvas.bbox('all'))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')

        # Define 13-step workflow
        self.workflow_steps = [
            {
                'number': 1,
                'title': '📧 Process New Cases',
                'description': 'Extract data from ORDER.pdf and ARP.pdf files',
                'command': self.process_pdfs,
                'button_text': 'Process PDFs'
            },
            {
                'number': 2,
                'title': '🗂️ Organize Case Files',
                'description': 'Create folder structure and organize documents',
                'command': self.organize_files,
                'button_text': 'Organize Files'
            },
            {
                'number': 3,
                'title': '🗺️ Generate Route Map',
                'description': 'Create map showing ward locations',
                'command': self.generate_map,
                'button_text': 'Create Map'
            },
            {
                'number': 4,
                'title': '📧 Send Meeting Requests',
                'description': 'Send initial email to guardians',
                'command': self.send_meeting_requests,
                'button_text': 'Send Requests'
            },
            {
                'number': 5,
                'title': '👥 Add Contacts',
                'description': 'Add guardian and ward to contacts',
                'command': self.add_contacts,
                'button_text': 'Add Contacts'
            },
            {
                'number': 6,
                'title': '📅 Confirm Appointment',
                'description': 'Send appointment confirmation email',
                'command': self.send_confirmations,
                'button_text': 'Send Confirmation'
            },
            {
                'number': 7,
                'title': '📅 Schedule Calendar',
                'description': 'Add appointment to Google Calendar',
                'command': self.schedule_events,
                'button_text': 'Schedule'
            },
            {
                'number': 8,
                'title': '📄 Generate CVR',
                'description': 'Create Court Visitor Report documents',
                'command': self.generate_cvr,
                'button_text': 'Generate CVR'
            },
            {
                'number': 9,
                'title': '📊 Visit Summary',
                'description': 'Create visit summary sheet',
                'command': self.generate_summaries,
                'button_text': 'Generate Summary'
            },
            {
                'number': 10,
                'title': '📝 Complete CVR',
                'description': 'Fill CVR with Google Form data',
                'command': self.complete_cvr,
                'button_text': 'Complete CVR'
            },
            {
                'number': 11,
                'title': '📧 Send Follow-up',
                'description': 'Send follow-up emails to guardians',
                'command': self.send_followups,
                'button_text': 'Send Follow-up'
            },
            {
                'number': 12,
                'title': '📧 Email CVR',
                'description': 'Email CVR to supervisor',
                'command': self.email_cvr,
                'button_text': 'Email CVR'
            },
            {
                'number': 13,
                'title': '💵 Payment Form',
                'description': 'Generate payment reimbursement form',
                'command': self.generate_payment_forms,
                'button_text': 'Generate Form'
            },
            {
                'number': 14,
                'title': '🚗 Mileage Log',
                'description': 'Generate mileage reimbursement log',
                'command': self.generate_mileage_log,
                'button_text': 'Generate Mileage'
            }
        ]

        # Create 3-column layout using GRID (not pack) to avoid geometry manager conflicts
        scrollable_frame.columnconfigure(0, weight=1)
        scrollable_frame.columnconfigure(1, weight=1)
        scrollable_frame.columnconfigure(2, weight=1)

        # Column 1: Steps 1-5
        column1_frame = tk.Frame(scrollable_frame, bg=self.colors['bg'])
        column1_frame.grid(row=0, column=0, sticky='nsew', padx=(0, 10))

        col1_header = tk.Label(column1_frame,
                              text="Input Phase (Steps 1-5)",
                              font=('Segoe UI', 14, 'bold'),
                              bg=self.colors['bg'])
        col1_header.pack(pady=(0, 10))

        # Column 2: Steps 6-10
        column2_frame = tk.Frame(scrollable_frame, bg=self.colors['bg'])
        column2_frame.grid(row=0, column=1, sticky='nsew', padx=10)

        col2_header = tk.Label(column2_frame,
                               text="Communication (Steps 6-10)",
                               font=('Segoe UI', 14, 'bold'),
                               bg=self.colors['bg'])
        col2_header.pack(pady=(0, 10))

        # Column 3: Steps 11-15
        column3_frame = tk.Frame(scrollable_frame, bg=self.colors['bg'])
        column3_frame.grid(row=0, column=2, sticky='nsew', padx=(10, 0))

        col3_header = tk.Label(column3_frame,
                               text="Wrap-Up (Steps 11-15)",
                               font=('Segoe UI', 14, 'bold'),
                               bg=self.colors['bg'])
        col3_header.pack(pady=(0, 10))

        # Create step cards in 3 columns
        for i, step in enumerate(self.workflow_steps):
            if i < 5:  # Steps 1-5
                self.create_step_card(column1_frame, step)
            elif i < 10:  # Steps 6-10
                self.create_step_card(column2_frame, step)
            else:  # Steps 11-14
                self.create_step_card(column3_frame, step)

    def check_api_configured(self, api_name):
        """Check if an API is configured

        Returns:
            'ready': API is fully configured
            'partial': API is partially configured (e.g., Maps - works but limited)
            'missing': API is not configured
        """
        config_dir = self.base_dir / "Config"

        if api_name == 'vision':
            # Check for Vision API service account JSON
            service_account = config_dir / "API" / "google_service_account.json"
            return 'ready' if service_account.exists() else 'missing'

        elif api_name == 'maps':
            # Maps is optional - works without but limited
            maps_key = config_dir / "Keys" / "google_maps_api_key.txt"
            return 'ready' if maps_key.exists() else 'partial'

        elif api_name == 'gmail':
            # Check for Gmail OAuth credentials
            gmail_client = config_dir / "API" / "gmail_oauth_client.json"
            return 'ready' if gmail_client.exists() else 'missing'

        elif api_name == 'people':
            # Check for People API OAuth credentials (uses same as calendar)
            client_secret = config_dir / "API" / "client_secret_calendar.json"
            return 'ready' if client_secret.exists() else 'missing'

        elif api_name == 'calendar':
            # Check for Calendar API OAuth credentials
            client_secret = config_dir / "API" / "client_secret_calendar.json"
            return 'ready' if client_secret.exists() else 'missing'

        elif api_name == 'sheets':
            # Check for Sheets API service account JSON (same as Vision)
            service_account = config_dir / "API" / "google_service_account.json"
            return 'ready' if service_account.exists() else 'missing'

        return 'ready'  # No API required

    def create_step_card(self, parent, step):
        """Create a card for each workflow step"""
        card = tk.Frame(parent, bg=self.colors['card_bg'],
                       relief='raised', borderwidth=1)
        card.pack(fill='x', pady=8, padx=5)

        # Inner padding
        inner = tk.Frame(card, bg=self.colors['card_bg'])
        inner.pack(fill='both', expand=True, padx=15, pady=12)

        # Step number and title
        header = tk.Frame(inner, bg=self.colors['card_bg'])
        header.pack(fill='x')

        step_num = tk.Label(header,
                           text=f"Step {step['number']}",
                           font=('Segoe UI', 10, 'bold'),
                           fg=self.colors['primary'],
                           bg=self.colors['card_bg'])
        step_num.pack(side='left')

        title = tk.Label(header,
                        text=step['title'],
                        font=('Segoe UI', 12, 'bold'),
                        bg=self.colors['card_bg'])
        title.pack(side='left', padx=(10, 0))

        # Add status indicator if API is required
        step_number = step['number']
        if step_number in self.api_requirements:
            api_name = self.api_requirements[step_number]
            status = self.check_api_configured(api_name)

            if status == 'ready':
                status_icon = tk.Label(header,
                                     text="✅",
                                     font=('Segoe UI', 10),
                                     bg=self.colors['card_bg'])
                status_icon.pack(side='left', padx=(5, 0))
            elif status == 'partial':
                status_icon = tk.Label(header,
                                     text="⚠️",
                                     font=('Segoe UI', 10),
                                     bg=self.colors['card_bg'])
                status_icon.pack(side='left', padx=(5, 0))
            else:  # missing
                status_icon = tk.Label(header,
                                     text="🔒",
                                     font=('Segoe UI', 10),
                                     bg=self.colors['card_bg'])
                status_icon.pack(side='left', padx=(5, 0))
        else:
            # No API required - always ready
            status_icon = tk.Label(header,
                                 text="✅",
                                 font=('Segoe UI', 10),
                                 bg=self.colors['card_bg'])
            status_icon.pack(side='left', padx=(5, 0))

        # Description
        desc = tk.Label(inner,
                       text=step['description'],
                       font=('Segoe UI', 9),
                       fg='#64748b',
                       bg=self.colors['card_bg'],
                       wraplength=400,
                       justify='left')
        desc.pack(fill='x', pady=(5, 10))

        # Action button(s)
        btn_frame = tk.Frame(inner, bg=self.colors['card_bg'])
        btn_frame.pack(anchor='w', fill='x')

        btn = ttk.Button(btn_frame,
                        text=step['button_text'],
                        style='Primary.TButton',
                        command=step['command'])
        btn.pack(side='left')

        # Add Settings button for Step 14 (Mileage Log)
        if step['number'] == 14:
            settings_btn = ttk.Button(btn_frame,
                                     text="⚙️ Settings",
                                     command=self.mileage_settings)
            settings_btn.pack(side='left', padx=(10, 0))

    def create_status_bar(self):
        """Create status bar at bottom"""
        self.status_bar = tk.Frame(self.root, bg='#e5e7eb', height=40)
        self.status_bar.pack(fill='x', side='bottom')
        self.status_bar.pack_propagate(False)

        self.status_label = tk.Label(self.status_bar,
                                    text="Ready",
                                    font=('Segoe UI', 10),
                                    bg='#e5e7eb',
                                    fg='#374151')
        self.status_label.pack(side='left', padx=20, pady=10)

        # Copyright notice
        copyright_label = tk.Label(self.status_bar,
                                   text="© 2024 GuardianShip Easy, LLC",
                                   font=('Segoe UI', 9),
                                   bg='#e5e7eb',
                                   fg='#6b7280')
        copyright_label.pack(side='right', padx=20, pady=10)

    def update_status(self, message, color='#374151'):
        """Update status bar message"""
        self.status_label.config(text=message, fg=color)
        self.root.update_idletasks()

    # ==================== STEP 1: Process PDFs ====================
    def process_pdfs(self):
        """Step 1: Run guardian_extractor to process PDFs"""
        if self.is_processing:
            messagebox.showwarning("In Progress", "Processing is already running!")
            return

        # Check if script exists
        if not self.guardian_extractor.exists():
            messagebox.showerror("Error",
                               f"Guardian extractor script not found:\n{self.guardian_extractor}")
            return

        # Check if New Files directory exists
        if not self.new_files_dir.exists():
            messagebox.showerror("Error",
                               f"'New Files' directory not found:\n{self.new_files_dir}")
            return

        # Count PDF files
        pdf_count = len(list(self.new_files_dir.glob("*.pdf")))
        if pdf_count == 0:
            messagebox.showinfo("No Files",
                              f"No PDF files found in:\n{self.new_files_dir}\n\n"
                              "Please add ORDER.pdf and ARP.pdf files to process.")
            return

        # Confirm before processing
        response = messagebox.askyesno("Process PDFs",
                                      f"Found {pdf_count} PDF files in 'New Files' directory.\n\n"
                                      "This will extract data from ORDER and ARP files "
                                      "and update the Excel database.\n\n"
                                      "Continue?")
        if not response:
            return

        # Open processing window
        self.show_processing_window()
        self.process_window.title("Step 1: OCR Guardian Data")
        self.process_title_label.config(text="Processing: OCR Guardian Data from PDFs")

        # Create thread-safe queue for output
        self.output_queue = queue.Queue()
        self.is_processing = True
        self.update_status("Running OCR extraction...", self.colors['warning'])

        # Run extractor in background thread
        thread = threading.Thread(target=self.run_guardian_extractor, daemon=True)
        thread.start()

        # Start polling queue from main thread (THREAD SAFE!)
        self.poll_output_queue()

    def show_processing_window(self):
        """Show processing window with output"""
        # Destroy old window if it exists to prevent blinking/multiple windows
        if hasattr(self, 'process_window'):
            try:
                self.process_window.destroy()
            except:
                pass

        self.process_window = tk.Toplevel(self.root)
        self.process_window.title("Processing...")
        self.process_window.geometry("800x600")
        self.process_window.transient(self.root)

        # Title (will be updated by caller)
        self.process_title_label = tk.Label(self.process_window,
                        text="Please stand by, processing your request...",
                        font=('Segoe UI', 14, 'bold'),
                        bg='white',
                        pady=15)
        self.process_title_label.pack(fill='x')

        # Progress indicator
        self.progress = ttk.Progressbar(self.process_window,
                                       mode='indeterminate')
        self.progress.pack(fill='x', padx=20, pady=10)
        self.progress.start(10)

        # Output text
        output_frame = tk.Frame(self.process_window)
        output_frame.pack(fill='both', expand=True, padx=20, pady=10)

        self.output_text = scrolledtext.ScrolledText(output_frame,
                                                     wrap=tk.WORD,
                                                     font=('Consolas', 9),
                                                     bg='#1e1e1e',
                                                     fg='#d4d4d4')
        self.output_text.pack(fill='both', expand=True)

        # Close button (disabled during processing)
        self.close_btn = ttk.Button(self.process_window,
                                    text="Close",
                                    state='disabled',
                                    command=self.process_window.destroy)
        self.close_btn.pack(pady=10)

    def poll_output_queue(self):
        """Poll the output queue from main thread - THREAD SAFE GUI updates"""
        try:
            while True:
                msg_type, msg_data = self.output_queue.get_nowait()

                if msg_type == 'line':
                    # Update text widget (SAFE - running in main thread)
                    self.output_text.insert(tk.END, msg_data)
                    self.output_text.see(tk.END)

                elif msg_type == 'success':
                    # Processing completed successfully
                    self.output_text.insert(tk.END, "\n\n✅ Processing completed successfully!\n")
                    self.output_text.see(tk.END)
                    self.update_status("Processing complete!", self.colors['success'])
                    self.progress.stop()
                    self.close_btn.config(state='normal')
                    self.is_processing = False
                    messagebox.showinfo("Success",
                                      "PDF processing completed successfully!\n\n"
                                      f"Check Excel file:\n{self.excel_path}")
                    return  # Stop polling

                elif msg_type == 'error':
                    # Processing failed
                    self.output_text.insert(tk.END, f"\n\n❌ Process failed with exit code {msg_data}\n")
                    self.output_text.see(tk.END)
                    self.update_status("Processing failed!", self.colors['error'])
                    self.progress.stop()
                    self.close_btn.config(state='normal')
                    self.is_processing = False
                    messagebox.showerror("Error", f"Processing failed with exit code {msg_data}")
                    return  # Stop polling

                elif msg_type == 'exception':
                    # Exception occurred
                    self.output_text.insert(tk.END, f"\n\n❌ Error: {msg_data}\n")
                    self.output_text.see(tk.END)
                    self.update_status("Error occurred!", self.colors['error'])
                    self.progress.stop()
                    self.close_btn.config(state='normal')
                    self.is_processing = False
                    messagebox.showerror("Error", f"An error occurred:\n{msg_data}")
                    return  # Stop polling

        except queue.Empty:
            pass  # No messages yet

        # Schedule next poll in 100ms (SAFE - using after() from main thread)
        self.root.after(100, self.poll_output_queue)

    def run_guardian_extractor(self):
        """Run the guardian extractor script - THREAD SAFE (no direct GUI updates)"""
        try:
            # Run the extractor script (OCR script now patches subprocess globally!)
            process = subprocess.Popen(
                [sys.executable, str(self.guardian_extractor)],
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                bufsize=1,  # Line buffered
                cwd=str(self.base_dir),
                startupinfo=get_subprocess_startupinfo(),
                creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == 'win32' else 0,
                shell=False
            )

            # Stream output to queue (NO GUI updates in thread!)
            for line in process.stdout:
                self.output_queue.put(('line', line))

            # Wait for completion
            process.wait()

            # Send completion status to queue
            if process.returncode == 0:
                self.output_queue.put(('success', None))
            else:
                self.output_queue.put(('error', process.returncode))

        except Exception as e:
            self.output_queue.put(('exception', str(e)))

        finally:
            self.is_processing = False
            # Safely stop progress bar
            try:
                if hasattr(self, 'progress') and self.progress.winfo_exists():
                    self.progress.stop()
            except:
                pass
            if hasattr(self, 'close_btn') and self.close_btn.winfo_exists():
                self.close_btn.config(state='normal')
            if hasattr(self, 'process_window') and self.process_window.winfo_exists():
                self.update_status("Ready", '#374151')

    # ==================== GENERIC SCRIPT RUNNER ====================
    def run_automation_script(self, step_number, step_name, script_args=None):
        """Generic method to run any automation script

        Args:
            step_number: Step number (2-14)
            step_name: Display name for the step
            script_args: Optional list of command-line arguments to pass to script
        """
        if self.is_processing:
            messagebox.showwarning("In Progress", "Processing is already running!")
            return

        # Get script path
        script_path = self.scripts.get(step_number)
        if not script_path:
            messagebox.showerror("Error", f"No script configured for Step {step_number}")
            return

        # Check if script exists
        if not script_path.exists():
            messagebox.showerror("Script Not Found",
                               f"Script not found:\n{script_path}\n\n"
                               f"Please verify the script exists at this location.")
            return

        # Confirm before running
        response = messagebox.askyesno(f"Step {step_number}",
                                      f"Run {step_name}?\n\n"
                                      f"Script: {script_path.name}\n\n"
                                      "Continue?")
        if not response:
            return

        # Open processing window
        self.show_processing_window()
        self.process_window.title(f"Step {step_number}: {step_name}")
        self.process_title_label.config(text=f"Processing: {step_name}")

        # Run script in background thread
        thread = threading.Thread(
            target=self.run_script_thread,
            args=(script_path, step_name, script_args),
            daemon=True
        )
        thread.start()

    def run_script_thread(self, script_path, step_name, script_args=None):
        """Run script in background thread

        Args:
            script_path: Path to the script to run
            step_name: Display name for the step
            script_args: Optional list of command-line arguments
        """
        self.is_processing = True
        self.update_status(f"Running {step_name}...", self.colors['warning'])

        try:
            # Build command with optional arguments
            cmd = [sys.executable, str(script_path)]
            if script_args:
                cmd.extend(script_args)

            # Run the script (with STARTUPINFO to hide console windows)
            process = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                bufsize=0,  # Unbuffered to reduce console activity
                cwd=str(script_path.parent),
                startupinfo=get_subprocess_startupinfo(),
                creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == 'win32' else 0,
                shell=False
            )

            # Stream output WITHOUT update calls
            for line in process.stdout:
                self.output_text.insert(tk.END, line)
                self.output_text.see(tk.END)
                # No GUI update here - let Tkinter handle it naturally

            # Wait for completion
            process.wait()

            # Handle result
            if process.returncode == 0:
                self.output_text.insert(tk.END, f"\n\n[OK] {step_name} completed successfully!\n")
                self.output_text.see(tk.END)
                self.update_status(f"{step_name} complete!", self.colors['success'])
                messagebox.showinfo("Success", f"{step_name} completed successfully!")
            else:
                self.output_text.insert(tk.END, f"\n\n[FAIL] Process failed with exit code {process.returncode}\n")
                self.output_text.see(tk.END)
                self.update_status(f"{step_name} failed!", self.colors['error'])
                messagebox.showerror("Error", f"{step_name} failed with exit code {process.returncode}")

        except Exception as e:
            self.output_text.insert(tk.END, f"\n\n[ERROR] {str(e)}\n")
            self.output_text.see(tk.END)
            self.update_status("Error occurred!", self.colors['error'])
            messagebox.showerror("Error", f"An error occurred:\n{str(e)}")

        finally:
            self.is_processing = False
            # Safely stop progress bar
            try:
                if hasattr(self, 'progress') and self.progress.winfo_exists():
                    self.progress.stop()
            except:
                pass
            if hasattr(self, 'close_btn') and self.close_btn.winfo_exists():
                self.close_btn.config(state='normal')
            if hasattr(self, 'process_window') and self.process_window.winfo_exists():
                self.update_status("Ready", '#374151')

    # ==================== AUTOMATION METHODS FOR STEPS 2-13 ====================
    def organize_files(self):
        """Step 2: Organize Case Files"""
        self.run_automation_script(2, "Organize Case Files")

    def generate_map(self):
        """Step 3: Generate Route Map"""
        self.run_automation_script(3, "Generate Route Map")

    def send_meeting_requests(self):
        """Step 4: Send Meeting Requests"""
        try:
            # Show week picker dialog first
            week_data = self.show_week_picker_dialog()
            if week_data is None:
                return  # User cancelled

            # Pass week data as command-line arguments
            script_path = self.scripts[4]
            if not script_path.exists():
                messagebox.showerror("Script Not Found",
                                   f"Script not found:\n{script_path}\n\n"
                                   f"Please verify the script exists at this location.")
                return

            # Build command with arguments
            cmd = [
                sys.executable,
                str(script_path),
                "--week", week_data['week_range'],
                "--days", week_data['preferred_days'],
                "--mode", "send",  # Send emails directly via Gmail
                "--confirm-write",  # Mark Excel 'emailsent' column
            ]

            print(f"DEBUG: Week range: {week_data['week_range']}")
            print(f"DEBUG: Preferred days: {week_data['preferred_days']}")
            print(f"DEBUG: Running command: {' '.join(cmd)}")

            # Open processing window
            self.show_processing_window()
            self.process_window.title("Step 4: Send Meeting Requests")
            self.process_title_label.config(text="Processing: Send Meeting Requests")

            # Run script in background thread with custom command and success message
            thread = threading.Thread(
                target=self.run_step4_script,
                args=(cmd,),
                daemon=True
            )
            thread.start()
        except Exception as e:
            print(f"ERROR in send_meeting_requests: {e}")
            import traceback
            traceback.print_exc()
            messagebox.showerror("Error", f"An error occurred:\n{str(e)}")

    def add_contacts(self):
        """Step 5: Add Contacts"""
        self.run_automation_script(5, "Add Contacts", script_args=["--mode", "live"])

    def send_confirmations(self):
        """Step 6: Send Confirmations"""
        self.run_automation_script(6, "Send Confirmations", script_args=["--mode", "live"])

    def schedule_events(self):
        """Step 7: Schedule Calendar Events"""
        self.run_automation_script(7, "Schedule Calendar Events", script_args=["--mode", "live"])

    def generate_cvr(self):
        """Step 8: Generate CVR"""
        self.run_automation_script(8, "Generate CVR")

    def generate_summaries(self):
        """Step 9: Generate Visit Summaries"""
        self.run_automation_script(9, "Generate Visit Summaries", script_args=['--open'])

    def complete_cvr(self):
        """Step 10: Complete CVR with Form Data"""
        self.run_automation_script(10, "Complete CVR")

    def send_followups(self):
        """Step 11: Send Follow-up Emails"""
        self.run_automation_script(11, "Send Follow-up Emails")

    def email_cvr(self):
        """Step 12: Email CVR to Supervisor"""
        self.run_automation_script(12, "Email CVR to Supervisor")

    def generate_payment_forms(self):
        """Step 13: Generate Payment Forms"""
        self.run_automation_script(13, "Generate Payment Forms")

    def generate_mileage_log(self):
        """Step 14: Generate Mileage Log"""
        self.run_automation_script(14, "Generate Mileage Log")

    def mileage_settings(self):
        """Open mileage address settings dialog"""
        config_file = self.base_dir / "Config" / "mileage_settings.txt"

        # Load current settings
        default_starting = "Default Starting Address"
        default_ending = "Default Ending Address"

        if config_file.exists():
            try:
                with open(config_file, 'r', encoding='utf-8') as f:
                    lines = [line for line in f.read().strip().split('\n') if not line.startswith('#')]
                    if len(lines) >= 2:
                        default_starting = lines[0].strip()
                        default_ending = lines[1].strip()
            except Exception as e:
                print(f"Error reading config: {e}")

        # Show settings dialog
        dialog = tk.Toplevel(self.root)
        dialog.title("Mileage Log - Address Settings")
        dialog.geometry("550x350")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()

        # Center dialog
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (275)
        y = (dialog.winfo_screenheight() // 2) - (175)
        dialog.geometry(f'+{x}+{y}')

        # Header
        header = tk.Label(dialog,
                         text="Set Your Starting and Ending Addresses",
                         font=('Segoe UI', 14, 'bold'),
                         pady=15)
        header.pack()

        # Instructions
        instructions = tk.Label(dialog,
                               text="Enter your starting and ending addresses for mileage calculations.\n"
                                    "These will be used as departure and return points for all trips.",
                               font=('Segoe UI', 10),
                               pady=10,
                               justify='left')
        instructions.pack()

        # Form frame
        form_frame = tk.Frame(dialog, padx=20, pady=10)
        form_frame.pack(fill='x')

        # Starting address
        tk.Label(form_frame, text="Starting Address (departure point):",
                font=('Segoe UI', 10, 'bold')).pack(anchor='w', pady=(5, 2))
        start_entry = tk.Entry(form_frame, font=('Segoe UI', 10), width=60)
        start_entry.insert(0, default_starting)
        start_entry.pack(pady=(0, 15))

        # Ending address
        tk.Label(form_frame, text="Ending Address (return point):",
                font=('Segoe UI', 10, 'bold')).pack(anchor='w', pady=(5, 2))
        end_entry = tk.Entry(form_frame, font=('Segoe UI', 10), width=60)
        end_entry.insert(0, default_ending)
        end_entry.pack(pady=(0, 10))

        # Note
        note = tk.Label(form_frame,
                       text="Note: Use full address including city, state, and ZIP code\n"
                            "Example: 123 Main St, Austin, TX 78701",
                       font=('Segoe UI', 8),
                       fg='gray',
                       justify='left')
        note.pack(anchor='w')

        # Buttons
        def save_settings():
            start = start_entry.get().strip()
            end = end_entry.get().strip()

            if not start or not end:
                messagebox.showerror("Error", "Both addresses are required!", parent=dialog)
                return

            try:
                config_file.parent.mkdir(parents=True, exist_ok=True)
                with open(config_file, 'w', encoding='utf-8') as f:
                    f.write("# Mileage Log Address Settings\n")
                    f.write("# Line 1: Starting address (where you depart from each day)\n")
                    f.write("# Line 2: Ending address (where you return to each day)\n")
                    f.write(f"{start}\n")
                    f.write(f"{end}\n")
                messagebox.showinfo("Success", "Addresses saved successfully!", parent=dialog)
                dialog.destroy()
            except Exception as e:
                messagebox.showerror("Error", f"Could not save settings:\n{e}", parent=dialog)

        btn_frame = tk.Frame(dialog, pady=15)
        btn_frame.pack()

        save_btn = tk.Button(btn_frame, text="Save",
                            command=save_settings,
                            font=('Segoe UI', 10, 'bold'),
                            bg='#667eea', fg='white',
                            padx=20, pady=5)
        save_btn.pack(side='left', padx=5)

        cancel_btn = tk.Button(btn_frame, text="Cancel",
                              command=dialog.destroy,
                              font=('Segoe UI', 10),
                              padx=20, pady=5)
        cancel_btn.pack(side='left', padx=5)

    # ==================== SIDEBAR UTILITY BUTTONS ====================
    def open_excel(self):
        """Open the Excel database file"""
        # Prevent double-click from opening multiple instances
        if self.is_opening_excel:
            return

        if self.excel_path.exists():
            try:
                self.is_opening_excel = True
                self.update_status("Opening Excel file...", self.colors['primary'])
                os.startfile(str(self.excel_path))
                # Reset flag after 2 seconds to allow reopening if needed
                self.root.after(2000, lambda: setattr(self, 'is_opening_excel', False))
            except Exception as e:
                self.is_opening_excel = False
                messagebox.showerror("Error", f"Could not open Excel file:\n{str(e)}")
        else:
            messagebox.showerror("File Not Found",
                               f"Excel file not found:\n{self.excel_path}\n\n"
                               "Run Step 1 to create the database.")

    def open_new_clients(self):
        """Open the New Clients directory"""
        new_clients = self.base_dir / "New Clients"
        if new_clients.exists():
            try:
                os.startfile(str(new_clients))
            except Exception as e:
                messagebox.showerror("Error", f"Could not open New Clients folder:\n{str(e)}")
        else:
            messagebox.showwarning("Not Found", "New Clients folder not found")

    def open_new_files(self):
        """Open the New Files directory"""
        new_files = self.base_dir / "New Files"
        if new_files.exists():
            try:
                os.startfile(str(new_files))
                self.update_status("Opening New Files...", self.colors['primary'])
            except Exception as e:
                messagebox.showerror("Error", f"Could not open folder:\n{str(e)}")
        else:
            # Create the folder if it doesn't exist
            response = messagebox.askyesno("Folder Not Found",
                                         f"New Files directory not found.\n\n"
                                         "Would you like to create it?")
            if response:
                try:
                    new_files.mkdir(parents=True, exist_ok=True)
                    os.startfile(str(new_files))
                    self.update_status("Created New Files folder", self.colors['success'])
                except Exception as e:
                    messagebox.showerror("Error", f"Could not create folder:\n{str(e)}")

    def open_settings(self):
        """Open Court Visitor Settings dialog"""
        try:
            from cv_settings_dialog import show_cv_settings_dialog
            saved = show_cv_settings_dialog(parent=self.root, first_run=False)
            if saved:
                self.update_status("Settings saved successfully", self.colors['success'])
                messagebox.showinfo("Settings Saved",
                                  "Your Court Visitor information has been saved.\n\n"
                                  "This information will be automatically filled in your forms.",
                                  parent=self.root)
        except Exception as e:
            messagebox.showerror("Error", f"Could not open settings:\n{str(e)}")

    def check_first_run_setup(self):
        """Check if this is first run and show setup wizard if needed"""
        try:
            # First, check EULA acceptance
            from eula_dialog import check_eula_acceptance, show_eula_dialog

            if not check_eula_acceptance():
                accepted = show_eula_dialog(parent=self.root)
                if not accepted:
                    # User declined EULA - exit application
                    messagebox.showwarning(
                        "License Agreement Required",
                        "You must accept the End User License Agreement to use this software.\n\n"
                        "The application will now exit.",
                        parent=self.root
                    )
                    self.root.destroy()
                    sys.exit(0)
                    return

            # Then check CV settings
            from cv_info_manager import CVInfoManager
            cv_manager = CVInfoManager()

            # If not configured, show first-run setup
            if not cv_manager.is_configured():
                from cv_settings_dialog import show_cv_settings_dialog
                saved = show_cv_settings_dialog(parent=self.root, first_run=True)
                if saved:
                    self.update_status("Setup complete! Ready to use.", self.colors['success'])

            # Create desktop shortcut on first run (if it doesn't exist)
            from desktop_shortcut import check_desktop_shortcut_exists, create_desktop_shortcut
            if not check_desktop_shortcut_exists():
                try:
                    create_desktop_shortcut(app_name="Court Visitor App")
                except Exception as shortcut_error:
                    print(f"Note: Could not create desktop shortcut: {shortcut_error}")
                    # Non-critical error, continue anyway

        except Exception as e:
            print(f"First-run setup check failed: {e}")

    def show_about(self):
        """Show About dialog"""
        try:
            from about_dialog import show_about_dialog
            show_about_dialog(parent=self.root, version=__version__)
        except Exception as e:
            messagebox.showerror("Error", f"Could not open About dialog:\n{str(e)}")

    def show_license(self):
        """Show License Agreement dialog"""
        try:
            from eula_dialog import show_eula_dialog
            # Show in read-only mode (already accepted)
            show_eula_dialog(parent=self.root)
        except Exception as e:
            messagebox.showerror("Error", f"Could not open License Agreement:\n{str(e)}")

    def open_documentation(self):
        """Open documentation folder"""
        try:
            docs_dir = self.base_dir / "Documentation"
            if docs_dir.exists():
                os.startfile(docs_dir)
                self.update_status("Opening Documentation folder...", self.colors['primary'])
            else:
                messagebox.showinfo(
                    "Documentation",
                    "Documentation folder not found.\n\n"
                    f"Expected location: {docs_dir}",
                    parent=self.root
                )
        except Exception as e:
            messagebox.showerror("Error", f"Could not open documentation:\n{str(e)}")

    def backup_my_data(self):
        """Create backup of all user data"""
        try:
            from backup_manager import create_backup

            self.update_status("Creating backup...", self.colors['primary'])
            backup_path = create_backup(self.root, str(self.base_dir))

            if backup_path:
                self.update_status("Backup completed successfully!", self.colors['success'])

                # Ask if user wants to open backup folder
                response = messagebox.askyesno(
                    "Backup Complete",
                    f"Backup created successfully!\n\n"
                    f"File: {backup_path.name}\n"
                    f"Location: {backup_path.parent}\n\n"
                    f"Would you like to open the backup folder?",
                    parent=self.root
                )

                if response:
                    os.startfile(backup_path.parent)
            else:
                self.update_status("Backup failed - no data found", self.colors['warning'])
                messagebox.showwarning(
                    "Backup Failed",
                    "No data found to backup.\n\n"
                    "The backup includes:\n"
                    "• Excel data file\n"
                    "• Config folder (settings, API tokens)\n"
                    "• Output folders (generated forms)\n"
                    "• New Clients folder\n\n"
                    "If this is your first time using the app, you may not have any data yet.",
                    parent=self.root
                )
        except Exception as e:
            self.update_status("Backup failed", self.colors['error'])
            messagebox.showerror("Backup Error", f"Could not create backup:\n{str(e)}")

    def open_contacts(self):
        """Open contacts application"""
        try:
            # Try to open Windows Contacts/People app
            subprocess.Popen('explorer shell:contacts', shell=True)
            self.update_status("Opening Contacts...", self.colors['primary'])
        except Exception as e:
            messagebox.showinfo("Contacts",
                              "To open Contacts:\n\n"
                              "Windows 10: Search for 'People' app\n"
                              "Windows 11: Search for 'Contacts' in Settings\n\n"
                              "Or use Outlook/Google Contacts in your browser.")

    def open_email(self):
        """Open default email client"""
        try:
            # Open default email client
            subprocess.Popen(r'explorer shell:AppsFolder\microsoft.windowscommunicationsapps_8wekyb3d8bbwe!microsoft.windowslive.mail', shell=True)
            self.update_status("Opening Email...", self.colors['primary'])
        except:
            try:
                # Fallback to mailto protocol
                os.startfile('mailto:')
            except Exception as e:
                messagebox.showinfo("Email",
                                  "To open Email:\n\n"
                                  "- Open your email client (Outlook, Gmail, etc.)\n"
                                  "- Or use your web browser to access email")

    def show_getting_started(self):
        """Show Getting Started guide with API setup requirements"""
        gs_window = tk.Toplevel(self.root)
        gs_window.title("Getting Started - Court Visitor App (BETA)")
        gs_window.geometry("950x750")
        gs_window.transient(self.root)

        # Title
        title_frame = tk.Frame(gs_window, bg='#1e3a8a', pady=20)
        title_frame.pack(fill='x')

        title = tk.Label(title_frame,
                        text="📖 Getting Started - Court Visitor App",
                        font=('Segoe UI', 18, 'bold'),
                        bg='#1e3a8a',
                        fg='white')
        title.pack()

        subtitle = tk.Label(title_frame,
                           text="Court Visitor App © 2024-2025 • BETA Version • All Rights Reserved",
                           font=('Segoe UI', 9),
                           bg='#1e3a8a',
                           fg='#bfdbfe')
        subtitle.pack(pady=(5, 0))

        # Scrollable content
        canvas = tk.Canvas(gs_window, bg='white')
        scrollbar = ttk.Scrollbar(gs_window, orient="vertical", command=canvas.yview)
        scrollable = tk.Frame(canvas, bg='white')

        scrollable.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True, padx=10, pady=10)
        scrollbar.pack(side="right", fill="y")

        # Content sections
        content_frame = tk.Frame(scrollable, bg='white', padx=30, pady=20)
        content_frame.pack(fill='both', expand=True)

        # CLIFF NOTES - Quick Start Section
        cliff_notes_title = tk.Label(content_frame,
                                     text="⚡ QUICK START - Essential Information (For the Time-Pressed)",
                                     font=('Segoe UI', 14, 'bold'),
                                     bg='#fef3c7',
                                     fg='#92400e',
                                     anchor='w',
                                     padx=15,
                                     pady=12)
        cliff_notes_title.pack(fill='x', pady=(0, 10))

        cliff_notes_intro = tk.Label(content_frame,
                                     text="We understand you're busy! Here are the absolute must-know items.\n"
                                          "However, we strongly encourage reading the full guide - it will save you hours of troubleshooting!",
                                     font=('Segoe UI', 9, 'italic'),
                                     bg='white',
                                     fg='#92400e',
                                     justify='left',
                                     anchor='w')
        cliff_notes_intro.pack(fill='x', padx=15, pady=(0, 10))

        cliff_notes = tk.Label(content_frame,
                              text="THE 5 CRITICAL THINGS YOU MUST KNOW:\n\n"
                                   "1. 🔐 Google APIs Required - Gmail account needed. Click 'Setup Google Vision' BEFORE Step 1.\n\n"
                                   "2. 📊 Close Excel Before OCR - Step 1 CANNOT run if Excel is open. Save & close it first.\n\n"
                                   "3. ✅ Verify OCR Results - After Step 1, ALWAYS check Excel. OCR isn't perfect - verify names,\n"
                                   "    dates, and numbers match the PDFs. Report errors to help us improve!\n\n"
                                   "4. 📁 Check 'Unmatched' Folder - After Step 2, check New Files\\Unmatched folder. Should be\n"
                                   "    empty. If not, manually move files to correct ward folders.\n\n"
                                   "5. 🔒 Security Warnings = Normal - When you see 'Do you trust this document?' → Answer YES.\n"
                                   "    When Google asks to re-authorize → Click ALLOW. These are expected!\n\n"
                                   "QUICK TROUBLESHOOTING:\n"
                                   "• Step won't run? → Check Excel is closed, PDFs are closed, folders aren't open\n"
                                   "• Need to re-send email/recreate CVR? → Clear the status column in Excel\n"
                                   "• OCR wrong? → Manually fix Excel AND report the error\n\n"
                                   "⚠️ BETA SOFTWARE: Your feedback is invaluable! Use 🐛 Report Bug and 💡 Request Feature buttons.",
                              font=('Segoe UI', 10),
                              bg='#fffbeb',
                              fg='#78350f',
                              justify='left',
                              anchor='w',
                              padx=15,
                              pady=15)
        cliff_notes.pack(fill='x', padx=10, pady=(0, 20))

        # Separator
        separator = tk.Frame(content_frame, height=2, bg='#e5e7eb')
        separator.pack(fill='x', pady=20)

        # Section 1: Steps that work WITHOUT setup
        section1_title = tk.Label(content_frame,
                                 text="✅ STEPS AVAILABLE WITHOUT API SETUP",
                                 font=('Segoe UI', 13, 'bold'),
                                 bg='#dcfce7',
                                 fg='#166534',
                                 anchor='w',
                                 padx=15,
                                 pady=10)
        section1_title.pack(fill='x', pady=(0, 10))

        section1_text = tk.Label(content_frame,
                                text="These steps work immediately - no configuration needed:\n\n"
                                     "• Step 2: Organize Case Files - Moves PDF files into ward folders\n"
                                     "• Step 8: Generate Court Visitor Reports - Creates CVR Word documents\n"
                                     "• Step 9: Generate Visit Summaries - Creates summary Word documents\n"
                                     "• Step 13: Build Payment Forms - Creates payment request forms\n"
                                     "• Step 14: Build Mileage Forms - Creates mileage reimbursement forms\n\n"
                                     "⚠️ CRITICAL REQUIREMENT - Excel File:\n"
                                     "ALL steps (with or without API setup) depend on accurate data in:\n"
                                     "ward_guardian_info.xlsx\n\n"
                                     "Without OCR (Step 1): You MUST manually enter case data into Excel\n"
                                     "With OCR (Step 1): OCR extracts data automatically, BUT you must:\n"
                                     "  • Carefully review all extracted data for accuracy\n"
                                     "  • Verify names, addresses, phone numbers, emails\n"
                                     "  • Correct any OCR mistakes before running other steps\n\n"
                                     "The Excel file is the central database - errors here propagate everywhere!\n\n"
                                     "Other Requirements: Microsoft Word must be installed",
                                font=('Segoe UI', 10),
                                bg='white',
                                fg='#374151',
                                justify='left',
                                anchor='w')
        section1_text.pack(fill='x', padx=15, pady=(0, 20))

        # Section 2: Partial functionality
        section2_title = tk.Label(content_frame,
                                 text="⚠️ PARTIAL FUNCTIONALITY (Works But Limited)",
                                 font=('Segoe UI', 13, 'bold'),
                                 bg='#fef3c7',
                                 fg='#92400e',
                                 anchor='w',
                                 padx=15,
                                 pady=10)
        section2_title.pack(fill='x', pady=(0, 10))

        section2_text = tk.Label(content_frame,
                                text="• Step 3: Generate Route Map\n"
                                     "  Without setup: Shows numbered dots on blank background (still useful for planning)\n"
                                     "  With Google Maps API: Shows dots on actual street maps (much better!)\n\n"
                                     "  → Click '🗺️ Setup Google Maps' in the sidebar to enable street maps",
                                font=('Segoe UI', 10),
                                bg='white',
                                fg='#374151',
                                justify='left',
                                anchor='w')
        section2_text.pack(fill='x', padx=15, pady=(0, 20))

        # Section 3: Requires API setup
        section3_title = tk.Label(content_frame,
                                 text="🔒 STEPS REQUIRING API SETUP",
                                 font=('Segoe UI', 13, 'bold'),
                                 bg='#fee2e2',
                                 fg='#991b1b',
                                 anchor='w',
                                 padx=15,
                                 pady=10)
        section3_title.pack(fill='x', pady=(0, 10))

        section3_text = tk.Label(content_frame,
                                text="These steps require Google Cloud API setup (one-time configuration):\n\n"
                                     "📧 Gmail API (for email automation):\n"
                                     "   • Step 4: Send Meeting Requests\n"
                                     "   • Step 6: Send Confirmations\n"
                                     "   • Step 11: Send Follow-up Emails\n"
                                     "   • Step 12: Email CVR to Supervisor\n\n"
                                     "👁️ Vision API (for PDF text extraction):\n"
                                     "   • Step 1: OCR Guardian Data\n\n"
                                     "👥 People API (for contact management):\n"
                                     "   • Step 5: Add Contacts to Google Contacts\n\n"
                                     "📅 Calendar API (for scheduling):\n"
                                     "   • Step 7: Schedule Calendar Events\n\n"
                                     "📊 Sheets API (for form data integration):\n"
                                     "   • Step 10: Complete CVR with Form Data",
                                font=('Segoe UI', 10),
                                bg='white',
                                fg='#374151',
                                justify='left',
                                anchor='w')
        section3_text.pack(fill='x', padx=15, pady=(0, 20))

        # Section 4: Setup instructions
        section4_title = tk.Label(content_frame,
                                 text="⚙️ HOW TO SETUP GOOGLE CLOUD APIS",
                                 font=('Segoe UI', 13, 'bold'),
                                 bg='#dbeafe',
                                 fg='#1e40af',
                                 anchor='w',
                                 padx=15,
                                 pady=10)
        section4_title.pack(fill='x', pady=(0, 10))

        section4_text = tk.Label(content_frame,
                                text="Each Google API requires a one-time setup (~15-20 minutes total):\n\n"
                                     "1. Create a Google Cloud Project (free - $200/month credit included)\n"
                                     "2. Enable the APIs you need (Vision, Gmail, Maps, etc.)\n"
                                     "3. Create credentials (API keys or OAuth client secrets)\n"
                                     "4. Download credential files and save to app's Config folder\n\n"
                                     "Setup wizards available in sidebar:\n"
                                     "   • 🔧 Setup Google Vision - For OCR (Step 1)\n"
                                     "   • 🗺️ Setup Google Maps - For route maps (Step 3)\n\n"
                                     "More wizards coming soon for Gmail, Calendar, People, and Sheets APIs!\n\n"
                                     "💡 TIP: You can use the app immediately for Steps 2, 8, 9, 13, 14\n"
                                     "while you set up APIs for the other features.",
                                font=('Segoe UI', 10),
                                bg='white',
                                fg='#374151',
                                justify='left',
                                anchor='w')
        section4_text.pack(fill='x', padx=15, pady=(0, 20))

        # Important notice about full guide
        full_guide_notice = tk.Label(content_frame,
                                     text="⚠️ THIS IS JUST A QUICK START SUMMARY",
                                     font=('Segoe UI', 12, 'bold'),
                                     bg='#fef2f2',
                                     fg='#991b1b',
                                     anchor='w',
                                     padx=15,
                                     pady=10)
        full_guide_notice.pack(fill='x', pady=(10, 5))

        full_guide_details = tk.Label(content_frame,
                                      text="The complete guide includes:\n"
                                           "• Detailed step-by-step instructions for all 14 workflow steps\n"
                                           "• Excel status columns reference table\n"
                                           "• Comprehensive troubleshooting section\n"
                                           "• Best practices and tips\n"
                                           "• Security validation explanations\n\n"
                                           "👉 Click below to open the FULL GUIDE with all details:\n\n"
                                           "📚 ALSO AVAILABLE: Click '📖 Manual' in Quick Access sidebar for the\n"
                                           "    comprehensive Court Visitor App Manual with advanced topics.",
                                      font=('Segoe UI', 10),
                                      bg='#fef2f2',
                                      fg='#7f1d1d',
                                      justify='left',
                                      anchor='w',
                                      padx=15,
                                      pady=10)
        full_guide_details.pack(fill='x', pady=(0, 10))

        # Buttons frame
        button_frame = tk.Frame(content_frame, bg='white')
        button_frame.pack(pady=20)

        # Full guide button
        def open_full_guide():
            # Try PDF first (best for viewing), then MD fallback
            pdf_path = self.base_dir / "README_FIRST.pdf"
            md_path = self.base_dir / "README_FIRST.md"

            if pdf_path.exists():
                try:
                    os.startfile(str(pdf_path))
                except Exception as e:
                    messagebox.showerror("Error Opening PDF",
                                        f"Could not open PDF: {e}\n\nPlease open manually:\n{pdf_path}")
            elif md_path.exists():
                try:
                    os.startfile(str(md_path))
                except:
                    messagebox.showinfo("Full Guide",
                                      f"Please open this file for the complete guide:\n\n{md_path}")
            else:
                messagebox.showwarning("Not Found",
                                      "README_FIRST.pdf or README_FIRST.md not found in app directory.")

        full_guide_btn = ttk.Button(button_frame,
                                    text="📖 Open Full Guide (PDF)",
                                    command=open_full_guide)
        full_guide_btn.pack(side='left', padx=5)

        # Close button
        close_btn = ttk.Button(button_frame, text="Got It!", command=gs_window.destroy)
        close_btn.pack(side='left', padx=5)

    def show_help(self):
        """Show help dialog"""
        help_window = tk.Toplevel(self.root)
        help_window.title("Court Visitor App - Help")
        help_window.geometry("600x500")
        help_window.transient(self.root)

        # Title
        title = tk.Label(help_window,
                        text="Court Visitor App Help",
                        font=('Segoe UI', 16, 'bold'),
                        bg='white',
                        pady=15)
        title.pack(fill='x')

        # Help text
        help_text = scrolledtext.ScrolledText(help_window,
                                             wrap=tk.WORD,
                                             font=('Segoe UI', 10),
                                             padx=20,
                                             pady=20)
        help_text.pack(fill='both', expand=True, padx=10, pady=10)

        help_content = """
COURT VISITOR APP - QUICK HELP

=== STEP 1: PROCESS NEW CASES ===

1. Receive ORDER.pdf and ARP.pdf files via email
2. Save them to: C:\\GoogleSync\\GuardianShip_Easy_App\\New Files\\
3. Click "Process PDFs" button in the app
4. Wait for processing to complete
5. Check Excel file for extracted data

The app will:
- Extract ward and guardian information
- Parse dates and addresses
- Update the Excel database automatically

=== QUICK ACCESS BUTTONS ===

📊 Excel File - Opens the ward_guardian_info.xlsx database
📁 Guardian Folders - Opens the case folders directory
👥 Contacts - Opens Windows Contacts/People app
📧 Email - Opens your default email client
❓ Help - Shows this help screen
📖 Manual - Opens the full user manual

=== TROUBLESHOOTING ===

Problem: "No PDF files found"
Solution: Make sure PDFs are in the "New Files" folder

Problem: "Excel file not found"
Solution: Run Step 1 to create the database

Problem: Processing fails
Solution: Check the output window for error details

=== GETTING STARTED ===

1. Process new cases (Step 1)
2. Organize your case files (Step 2)
3. Follow the workflow steps in order
4. Use Quick Access buttons for common tasks

For detailed instructions, click the "Manual" button.
        """

        help_text.insert('1.0', help_content)
        help_text.config(state='disabled')

        # Close button
        close_btn = ttk.Button(help_window,
                              text="Close",
                              command=help_window.destroy)
        close_btn.pack(pady=10)

    def show_chatbot(self):
        """Show the interactive chatbot assistant"""
        try:
            # Lazy import to avoid startup errors
            from court_visitor_chatbot import CourtVisitorChatbot
            chatbot = CourtVisitorChatbot(parent=self.root)
            chatbot.show_chatbot()
        except ImportError as e:
            messagebox.showerror("Chatbot Not Found",
                               f"Could not find chatbot module.\n\n"
                               f"Error: {e}\n\n"
                               f"Please use 'Help' or 'Manual' buttons instead.")
        except Exception as e:
            messagebox.showerror("Chatbot Error",
                               f"Could not start chatbot: {e}\n\n"
                               f"Please use 'Help' or 'Manual' buttons instead.")

    def show_manual(self):
        """Show or open the user manual"""
        # Use dynamic path if available, otherwise fallback
        if PATHS_AVAILABLE:
            manual_pdf = APP_PATHS.APP_ROOT / "Documentation" / "Manual" / "MAIN_MANUAL_TOC.pdf"
        else:
            manual_pdf = Path(__file__).parent / "Documentation" / "Manual" / "MAIN_MANUAL_TOC.pdf"

        # Try to open the manual
        if manual_pdf.exists():
            try:
                os.startfile(str(manual_pdf))
                self.update_status("Opening Manual (PDF - 137 pages)...", self.colors['primary'])
                return
            except Exception as e:
                messagebox.showerror("Error Opening Manual",
                                   f"Could not open manual: {e}\n\n"
                                   f"Manual location:\n{manual_pdf}\n\n"
                                   "Please open it manually.")
                return

        # Manual not found
        messagebox.showerror("Manual Not Found",
                           f"User manual not found at:\n{manual_pdf}\n\n"
                           "Please ensure the manual exists in:\n"
                           "Documentation/Manual/MAIN_MANUAL_TOC.pdf")

    def report_bug(self):
        """Show bug report dialog"""
        bug_window = tk.Toplevel(self.root)
        bug_window.title("Report a Bug")
        bug_window.geometry("600x700")
        bug_window.transient(self.root)

        # Title
        title = tk.Label(bug_window,
                        text="Report a Bug",
                        font=('Segoe UI', 16, 'bold'),
                        bg='white',
                        pady=15)
        title.pack(fill='x')

        # Instructions
        instructions = tk.Label(bug_window,
                               text="Please describe the bug and copy this information to email:",
                               font=('Segoe UI', 10),
                               bg='white',
                               pady=10)
        instructions.pack(fill='x', padx=20)

        # Template text
        template_frame = tk.Frame(bug_window)
        template_frame.pack(fill='both', expand=True, padx=20, pady=10)

        template_text = scrolledtext.ScrolledText(template_frame,
                                                  wrap=tk.WORD,
                                                  font=('Consolas', 9),
                                                  height=25)
        template_text.pack(fill='both', expand=True)

        bug_template = f"""Send to: guardianshipeasy@gmail.com
Subject: Court Visitor App - Bug Report

Bug Description:
[Describe what happened]

Steps to Reproduce:
1.
2.
3.

Expected Behavior:
[What should have happened]

Actual Behavior:
[What actually happened]

Error Messages (if any):
[Copy any error messages here]

Which Step Were You Running:
[e.g., Step 1: Process PDFs]

Date/Time:
{datetime.now().strftime('%Y-%m-%d %H:%M')}

Additional Information:
[Any other relevant details]
"""
        template_text.insert('1.0', bug_template)

        # Button frame
        btn_frame = tk.Frame(bug_window, bg='white')
        btn_frame.pack(fill='x', padx=20, pady=10)

        def copy_to_clipboard():
            self.root.clipboard_clear()
            self.root.clipboard_append(template_text.get('1.0', 'end-1c'))
            messagebox.showinfo("Copied", "Bug report template copied to clipboard!\n\nPaste it into your email to guardianshipeasy@gmail.com")

        def open_email_client():
            try:
                os.startfile('mailto:guardianshipeasy@gmail.com?subject=Court Visitor App - Bug Report')
            except:
                messagebox.showinfo("Email", "Please open your email client and send to:\nguardianshipeasy@gmail.com")

        copy_btn = ttk.Button(btn_frame,
                             text="Copy to Clipboard",
                             command=copy_to_clipboard)
        copy_btn.pack(side='left', padx=5)

        email_btn = ttk.Button(btn_frame,
                              text="Open Email",
                              command=open_email_client)
        email_btn.pack(side='left', padx=5)

        close_btn = ttk.Button(btn_frame,
                              text="Close",
                              command=bug_window.destroy)
        close_btn.pack(side='right', padx=5)

    def request_feature(self):
        """Show feature request dialog"""
        feature_window = tk.Toplevel(self.root)
        feature_window.title("Request a Feature")
        feature_window.geometry("600x650")
        feature_window.transient(self.root)

        # Title
        title = tk.Label(feature_window,
                        text="Request a Feature",
                        font=('Segoe UI', 16, 'bold'),
                        bg='white',
                        pady=15)
        title.pack(fill='x')

        # Instructions
        instructions = tk.Label(feature_window,
                               text="Please describe the feature and copy this information to email:",
                               font=('Segoe UI', 10),
                               bg='white',
                               pady=10)
        instructions.pack(fill='x', padx=20)

        # Template text
        template_frame = tk.Frame(feature_window)
        template_frame.pack(fill='both', expand=True, padx=20, pady=10)

        template_text = scrolledtext.ScrolledText(template_frame,
                                                  wrap=tk.WORD,
                                                  font=('Consolas', 9),
                                                  height=22)
        template_text.pack(fill='both', expand=True)

        feature_template = f"""Send to: guardianshipeasy@gmail.com
Subject: Court Visitor App - Feature Request

Feature Description:
[Describe the feature you'd like to see]

Problem It Solves:
[What problem would this feature solve?]

How You Would Use It:
[Describe how you would use this feature in your workflow]

Priority:
[ ] High - Would save significant time
[ ] Medium - Would be helpful
[ ] Low - Nice to have

Additional Details:
[Any other information that would help]

Date:
{datetime.now().strftime('%Y-%m-%d')}
"""
        template_text.insert('1.0', feature_template)

        # Button frame
        btn_frame = tk.Frame(feature_window, bg='white')
        btn_frame.pack(fill='x', padx=20, pady=10)

        def copy_to_clipboard():
            self.root.clipboard_clear()
            self.root.clipboard_append(template_text.get('1.0', 'end-1c'))
            messagebox.showinfo("Copied", "Feature request template copied to clipboard!\n\nPaste it into your email to guardianshipeasy@gmail.com")

        def open_email_client():
            try:
                os.startfile('mailto:guardianshipeasy@gmail.com?subject=Court Visitor App - Feature Request')
            except:
                messagebox.showinfo("Email", "Please open your email client and send to:\nguardianshipeasy@gmail.com")

        copy_btn = ttk.Button(btn_frame,
                             text="Copy to Clipboard",
                             command=copy_to_clipboard)
        copy_btn.pack(side='left', padx=5)

        email_btn = ttk.Button(btn_frame,
                              text="Open Email",
                              command=open_email_client)
        email_btn.pack(side='left', padx=5)

        close_btn = ttk.Button(btn_frame,
                              text="Close",
                              command=feature_window.destroy)
        close_btn.pack(side='right', padx=5)

    def setup_google_vision(self):
        """Show Google Vision API setup wizard"""
        setup_window = tk.Toplevel(self.root)
        setup_window.title("Google Vision API Setup")
        setup_window.geometry("700x800")
        setup_window.transient(self.root)

        # Title
        title = tk.Label(setup_window,
                        text="Google Vision API Setup Wizard",
                        font=('Segoe UI', 16, 'bold'),
                        bg='#2563eb',
                        fg='white',
                        pady=15)
        title.pack(fill='x')

        # Main content with scrollbar
        content_frame = tk.Frame(setup_window)
        content_frame.pack(fill='both', expand=True, padx=20, pady=20)

        canvas = tk.Canvas(content_frame, highlightthickness=0)
        scrollbar = ttk.Scrollbar(content_frame, orient='vertical', command=canvas.yview)
        scrollable = tk.Frame(canvas)

        scrollable.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas.create_window((0, 0), window=scrollable, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')

        # Instructions
        instructions = tk.Label(scrollable,
                               text="The Court Visitor App uses Google Vision API for OCR (reading text from PDFs).\n"
                                    "Follow these steps to set up your Google Vision API credentials:",
                               font=('Segoe UI', 11),
                               wraplength=600,
                               justify='left')
        instructions.pack(anchor='w', pady=(0, 20))

        # Step 1
        step1_frame = tk.LabelFrame(scrollable, text="Step 1: Create Google Cloud Project", font=('Segoe UI', 11, 'bold'), padx=15, pady=10)
        step1_frame.pack(fill='x', pady=10)

        step1_text = tk.Label(step1_frame,
                             text="1. Go to: https://console.cloud.google.com/\n"
                                  "2. Click 'Select a Project' at the top\n"
                                  "3. Click 'NEW PROJECT'\n"
                                  "4. Enter project name: 'Court Visitor App'\n"
                                  "5. Click 'CREATE'",
                             font=('Segoe UI', 10),
                             justify='left',
                             wraplength=600)
        step1_text.pack(anchor='w')

        def open_google_cloud():
            try:
                import webbrowser
                webbrowser.open('https://console.cloud.google.com/')
            except:
                messagebox.showinfo("URL", "Open this URL in your browser:\nhttps://console.cloud.google.com/")

        step1_btn = ttk.Button(step1_frame, text="Open Google Cloud Console", command=open_google_cloud)
        step1_btn.pack(anchor='w', pady=(10, 0))

        # Step 2
        step2_frame = tk.LabelFrame(scrollable, text="Step 2: Enable Vision API", font=('Segoe UI', 11, 'bold'), padx=15, pady=10)
        step2_frame.pack(fill='x', pady=10)

        step2_text = tk.Label(step2_frame,
                             text="1. In your project, go to 'APIs & Services' > 'Library'\n"
                                  "2. Search for 'Vision API'\n"
                                  "3. Click on 'Cloud Vision API'\n"
                                  "4. Click 'ENABLE'\n"
                                  "5. Wait for it to enable (takes a few seconds)",
                             font=('Segoe UI', 10),
                             justify='left',
                             wraplength=600)
        step2_text.pack(anchor='w')

        def open_api_library():
            try:
                import webbrowser
                webbrowser.open('https://console.cloud.google.com/apis/library')
            except:
                messagebox.showinfo("URL", "Open this URL:\nhttps://console.cloud.google.com/apis/library")

        step2_btn = ttk.Button(step2_frame, text="Open API Library", command=open_api_library)
        step2_btn.pack(anchor='w', pady=(10, 0))

        # Step 3
        step3_frame = tk.LabelFrame(scrollable, text="Step 3: Create Service Account", font=('Segoe UI', 11, 'bold'), padx=15, pady=10)
        step3_frame.pack(fill='x', pady=10)

        step3_text = tk.Label(step3_frame,
                             text="1. Go to 'APIs & Services' > 'Credentials'\n"
                                  "2. Click 'CREATE CREDENTIALS' > 'Service Account'\n"
                                  "3. Enter name: 'court-visitor-app'\n"
                                  "4. Click 'CREATE AND CONTINUE'\n"
                                  "5. For role, select: 'Project' > 'Editor'\n"
                                  "6. Click 'CONTINUE' then 'DONE'",
                             font=('Segoe UI', 10),
                             justify='left',
                             wraplength=600)
        step3_text.pack(anchor='w')

        def open_credentials():
            try:
                import webbrowser
                webbrowser.open('https://console.cloud.google.com/apis/credentials')
            except:
                messagebox.showinfo("URL", "Open this URL:\nhttps://console.cloud.google.com/apis/credentials")

        step3_btn = ttk.Button(step3_frame, text="Open Credentials Page", command=open_credentials)
        step3_btn.pack(anchor='w', pady=(10, 0))

        # Step 4
        step4_frame = tk.LabelFrame(scrollable, text="Step 4: Download JSON Key", font=('Segoe UI', 11, 'bold'), padx=15, pady=10)
        step4_frame.pack(fill='x', pady=10)

        step4_text = tk.Label(step4_frame,
                             text="1. On the Credentials page, find your service account\n"
                                  "2. Click the email address (ends with @...iam.gserviceaccount.com)\n"
                                  "3. Go to 'KEYS' tab\n"
                                  "4. Click 'ADD KEY' > 'Create new key'\n"
                                  "5. Choose 'JSON' format\n"
                                  "6. Click 'CREATE'\n"
                                  "7. A JSON file will download to your computer",
                             font=('Segoe UI', 10),
                             justify='left',
                             wraplength=600)
        step4_text.pack(anchor='w')

        # Step 5
        step5_frame = tk.LabelFrame(scrollable, text="Step 5: Install JSON Key File", font=('Segoe UI', 11, 'bold'), padx=15, pady=10)
        step5_frame.pack(fill='x', pady=10)

        step5_text = tk.Label(step5_frame,
                             text="The JSON key file must be placed in your app directory with a specific name:",
                             font=('Segoe UI', 10),
                             justify='left',
                             wraplength=600)
        step5_text.pack(anchor='w')

        json_path = self.base_dir / "court-visitor-vision-api.json"
        path_label = tk.Label(step5_frame,
                             text=f"Required location:\n{json_path}",
                             font=('Consolas', 9),
                             bg='#f0f0f0',
                             fg='#000080',
                             padx=10,
                             pady=10,
                             justify='left')
        path_label.pack(fill='x', pady=(10, 10))

        def browse_json_file():
            from tkinter import filedialog
            import shutil

            filename = filedialog.askopenfilename(
                title="Select Google Vision API JSON Key",
                filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
            )

            if filename:
                try:
                    # Copy to correct location
                    shutil.copy(filename, str(json_path))
                    messagebox.showinfo("Success",
                                      f"JSON key file installed!\n\n"
                                      f"Location: {json_path}\n\n"
                                      "The app is now ready to use Google Vision API for OCR.")
                    check_setup_status()
                except Exception as e:
                    messagebox.showerror("Error", f"Could not copy file:\n{str(e)}")

        browse_btn = ttk.Button(step5_frame, text="Browse and Install JSON Key File", command=browse_json_file)
        browse_btn.pack(anchor='w', pady=(0, 10))

        # Status indicator
        status_frame = tk.Frame(scrollable, bg='#f8fafc', relief='solid', borderwidth=1)
        status_frame.pack(fill='x', pady=20)

        status_label = tk.Label(status_frame,
                               text="Setup Status: Checking...",
                               font=('Segoe UI', 11, 'bold'),
                               bg='#f8fafc',
                               pady=10)
        status_label.pack()

        def check_setup_status():
            if json_path.exists():
                status_label.config(text="Setup Status: COMPLETE ✓",
                                   fg='#10b981',
                                   font=('Segoe UI', 12, 'bold'))
            else:
                status_label.config(text="Setup Status: JSON Key Not Found",
                                   fg='#ef4444',
                                   font=('Segoe UI', 12, 'bold'))

        check_setup_status()

        # Important note
        note_frame = tk.Frame(scrollable, bg='#fef3c7', relief='solid', borderwidth=1)
        note_frame.pack(fill='x', pady=10)

        note_label = tk.Label(note_frame,
                             text="⚠️ IMPORTANT: Keep your JSON key file secure!\n"
                                  "• Do not share it with anyone\n"
                                  "• Do not commit it to version control\n"
                                  "• It contains credentials for your Google Cloud account",
                             font=('Segoe UI', 9),
                             bg='#fef3c7',
                             fg='#92400e',
                             justify='left',
                             padx=15,
                             pady=10)
        note_label.pack(fill='x')

        # Close button
        close_btn = ttk.Button(scrollable, text="Close", command=setup_window.destroy)
        close_btn.pack(pady=20)

    def setup_google_maps(self):
        """Show Google Maps API setup wizard"""
        setup_window = tk.Toplevel(self.root)
        setup_window.title("Google Maps API Setup Wizard")
        setup_window.geometry("700x900")
        setup_window.transient(self.root)

        # Title
        title = tk.Label(setup_window,
                        text="Google Maps API Setup Wizard",
                        font=('Segoe UI', 16, 'bold'),
                        bg='#10b981',
                        fg='white',
                        pady=15)
        title.pack(fill='x')

        # Main content with scrollbar
        content_frame = tk.Frame(setup_window)
        content_frame.pack(fill='both', expand=True, padx=20, pady=20)

        canvas = tk.Canvas(content_frame, highlightthickness=0)
        scrollbar = ttk.Scrollbar(content_frame, orient='vertical', command=canvas.yview)
        scrollable = tk.Frame(canvas)

        scrollable.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas.create_window((0, 0), window=scrollable, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')

        # Instructions
        instructions = tk.Label(scrollable,
                               text="Step 3 (Generate Route Map) requires Google Maps API to show actual street maps.\n"
                                    "Without it, you only get dots on a blank background.\n\n"
                                    "Follow these steps to set up your Google Maps API key:",
                               font=('Segoe UI', 11),
                               wraplength=600,
                               justify='left')
        instructions.pack(anchor='w', pady=(0, 20))

        # Step 1
        step1_frame = tk.LabelFrame(scrollable, text="Step 1: Create Google Cloud Project", font=('Segoe UI', 11, 'bold'), padx=15, pady=10)
        step1_frame.pack(fill='x', pady=10)

        step1_text = tk.Label(step1_frame,
                             text="1. Go to: https://console.cloud.google.com/\n"
                                  "2. Click 'Select a Project' at the top\n"
                                  "3. Click 'NEW PROJECT'\n"
                                  "4. Enter project name: 'Court Visitor App'\n"
                                  "5. Click 'CREATE'",
                             font=('Segoe UI', 10),
                             justify='left',
                             wraplength=600)
        step1_text.pack(anchor='w')

        def open_google_cloud():
            try:
                import webbrowser
                webbrowser.open('https://console.cloud.google.com/')
            except:
                messagebox.showinfo("URL", "Open this URL in your browser:\nhttps://console.cloud.google.com/")

        step1_btn = ttk.Button(step1_frame, text="Open Google Cloud Console", command=open_google_cloud)
        step1_btn.pack(anchor='w', pady=(10, 0))

        # Step 2
        step2_frame = tk.LabelFrame(scrollable, text="Step 2: Enable Maps APIs", font=('Segoe UI', 11, 'bold'), padx=15, pady=10)
        step2_frame.pack(fill='x', pady=10)

        step2_text = tk.Label(step2_frame,
                             text="1. In your project, go to 'APIs & Services' > 'Library'\n"
                                  "2. Search for and ENABLE these 3 APIs:\n"
                                  "   • Maps Static API (for map images)\n"
                                  "   • Geocoding API (for address lookup)\n"
                                  "   • Directions API (for route planning)\n"
                                  "3. Wait for each to enable (takes a few seconds each)",
                             font=('Segoe UI', 10),
                             justify='left',
                             wraplength=600)
        step2_text.pack(anchor='w')

        def open_api_library():
            try:
                import webbrowser
                webbrowser.open('https://console.cloud.google.com/apis/library')
            except:
                messagebox.showinfo("URL", "Open this URL:\nhttps://console.cloud.google.com/apis/library")

        step2_btn = ttk.Button(step2_frame, text="Open API Library", command=open_api_library)
        step2_btn.pack(anchor='w', pady=(10, 0))

        # Step 3
        step3_frame = tk.LabelFrame(scrollable, text="Step 3: Create API Key", font=('Segoe UI', 11, 'bold'), padx=15, pady=10)
        step3_frame.pack(fill='x', pady=10)

        step3_text = tk.Label(step3_frame,
                             text="1. Go to 'APIs & Services' > 'Credentials'\n"
                                  "2. Click '+ CREATE CREDENTIALS' > 'API key'\n"
                                  "3. Copy the API key (starts with 'AIzaSy')\n"
                                  "4. RECOMMENDED: Click 'EDIT API KEY' to restrict it:\n"
                                  "   • API restrictions: Select 'Restrict key'\n"
                                  "   • Select: Maps Static API, Geocoding API, Directions API\n"
                                  "   • Click 'SAVE'",
                             font=('Segoe UI', 10),
                             justify='left',
                             wraplength=600)
        step3_text.pack(anchor='w')

        def open_credentials():
            try:
                import webbrowser
                webbrowser.open('https://console.cloud.google.com/apis/credentials')
            except:
                messagebox.showinfo("URL", "Open this URL:\nhttps://console.cloud.google.com/apis/credentials")

        step3_btn = ttk.Button(step3_frame, text="Open Credentials Page", command=open_credentials)
        step3_btn.pack(anchor='w', pady=(10, 0))

        # Step 4: Save API Key
        step4_frame = tk.LabelFrame(scrollable, text="Step 4: Save Your API Key", font=('Segoe UI', 11, 'bold'), padx=15, pady=10)
        step4_frame.pack(fill='x', pady=10)

        step4_text = tk.Label(step4_frame,
                             text="Paste your Google Maps API key below and click 'Save to Config Folder'.\n"
                                  "It will be saved to: Config\\Keys\\google_maps_api_key.txt",
                             font=('Segoe UI', 10),
                             justify='left',
                             wraplength=600)
        step4_text.pack(anchor='w', pady=(0, 10))

        # API Key input
        api_key_label = tk.Label(step4_frame, text="Google Maps API Key:", font=('Segoe UI', 10, 'bold'))
        api_key_label.pack(anchor='w', pady=(0, 5))

        api_key_entry = tk.Entry(step4_frame, font=('Segoe UI', 11), width=50)
        api_key_entry.pack(anchor='w', pady=(0, 10))

        # Check if API key already exists and pre-fill
        config_path = self.base_dir / "Config" / "Keys" / "google_maps_api_key.txt"
        if config_path.exists():
            try:
                with open(config_path, 'r') as f:
                    existing_key = f.read().strip()
                    api_key_entry.insert(0, existing_key)
            except:
                pass

        def save_api_key():
            api_key = api_key_entry.get().strip()
            if not api_key:
                messagebox.showerror("Error", "Please enter an API key!")
                return

            if not api_key.startswith("AIzaSy"):
                response = messagebox.askyesno("Warning",
                    "API key doesn't look like a valid Google Maps API key (should start with 'AIzaSy').\n\n"
                    "Save anyway?")
                if not response:
                    return

            try:
                config_path.parent.mkdir(parents=True, exist_ok=True)
                with open(config_path, 'w') as f:
                    f.write(api_key)
                messagebox.showinfo("Success",
                    f"API key saved successfully!\n\n"
                    f"Location: {config_path}\n\n"
                    f"Step 3 (Generate Route Map) will now show actual street maps!")
                check_setup_status()
            except Exception as e:
                messagebox.showerror("Error", f"Could not save API key:\n{str(e)}")

        save_btn = ttk.Button(step4_frame, text="💾 Save to Config Folder", command=save_api_key)
        save_btn.pack(anchor='w', pady=(0, 10))

        # Setup status indicator
        status_label = tk.Label(step4_frame, text="", font=('Segoe UI', 12, 'bold'))
        status_label.pack(anchor='w', pady=(10, 0))

        def check_setup_status():
            if config_path.exists():
                try:
                    with open(config_path, 'r') as f:
                        key = f.read().strip()
                        if key:
                            status_label.config(text="✅ Setup Status: API Key Configured!",
                                               fg='#10b981',
                                               font=('Segoe UI', 12, 'bold'))
                        else:
                            status_label.config(text="⚠️ Setup Status: API Key File Empty",
                                               fg='#f59e0b',
                                               font=('Segoe UI', 12, 'bold'))
                except:
                    status_label.config(text="❌ Setup Status: Could Not Read File",
                                       fg='#ef4444',
                                       font=('Segoe UI', 12, 'bold'))
            else:
                status_label.config(text="❌ Setup Status: API Key Not Found",
                                   fg='#ef4444',
                                   font=('Segoe UI', 12, 'bold'))

        check_setup_status()

        # Pricing info
        pricing_frame = tk.Frame(scrollable, bg='#dbeafe', relief='solid', borderwidth=1)
        pricing_frame.pack(fill='x', pady=10)

        pricing_label = tk.Label(pricing_frame,
                                text="💰 Pricing Info:\n"
                                     "• Google provides $200 FREE credit per month\n"
                                     "• Maps Static API: $2 per 1,000 requests\n"
                                     "• Geocoding API: $5 per 1,000 requests\n"
                                     "• Directions API: $5 per 1,000 requests\n"
                                     "• Typical usage: 50 cases/month = ~$2/month = FREE!\n"
                                     "• You only pay if you exceed $200/month",
                                font=('Segoe UI', 9),
                                bg='#dbeafe',
                                justify='left',
                                wraplength=600,
                                padx=15,
                                pady=10)
        pricing_label.pack()

        # Important note
        note_frame = tk.Frame(scrollable, bg='#fef3c7', relief='solid', borderwidth=1)
        note_frame.pack(fill='x', pady=10)

        note_label = tk.Label(note_frame,
                             text="⚠️ IMPORTANT:\n"
                                  "• Without this API key, Step 3 shows dots on a blank background (useless!)\n"
                                  "• With this API key, Step 3 shows dots on actual street maps (essential!)\n"
                                  "• Each end user needs their own API key ($200/month free credit each)",
                             font=('Segoe UI', 9),
                             bg='#fef3c7',
                             justify='left',
                             wraplength=600,
                             padx=15,
                             pady=10)
        note_label.pack()

        # Close button
        close_btn = ttk.Button(scrollable, text="Close", command=setup_window.destroy)
        close_btn.pack(pady=20)

    def show_week_picker_dialog(self):
        """Show dialog to pick week and preferred days for meeting requests"""
        from datetime import date, timedelta
        from dateutil.relativedelta import relativedelta, MO

        # Calculate default week (next Monday - Sunday)
        today = date.today()
        next_mon = today + relativedelta(weekday=MO(+1))
        next_sun = next_mon + timedelta(days=6)

        result = {'cancelled': True}

        # Create dialog
        dialog = tk.Toplevel(self.root)
        dialog.title("Step 4: Choose Meeting Week")
        dialog.geometry("550x500")
        dialog.configure(bg='#f9fafb')
        dialog.transient(self.root)
        dialog.grab_set()

        # Center on screen
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (550 // 2)
        y = (dialog.winfo_screenheight() // 2) - (500 // 2)
        dialog.geometry(f"550x500+{x}+{y}")

        # Main frame
        main_frame = tk.Frame(dialog, bg='#f9fafb', padx=20, pady=20)
        main_frame.pack(fill='both', expand=True)

        # Title
        title_label = tk.Label(main_frame,
                              text="📅 Select Meeting Week & Days",
                              font=('Segoe UI', 16, 'bold'),
                              bg='#f9fafb',
                              fg='#1f2937')
        title_label.pack(pady=(0, 20))

        # Week selection frame
        week_frame = tk.LabelFrame(main_frame,
                                  text="Meeting Week (Monday - Sunday)",
                                  font=('Segoe UI', 11, 'bold'),
                                  bg='#ffffff',
                                  fg='#374151',
                                  padx=15,
                                  pady=15)
        week_frame.pack(fill='x', pady=10)

        # Week start date
        start_label = tk.Label(week_frame,
                              text="Week Start (Monday):",
                              font=('Segoe UI', 10),
                              bg='#ffffff')
        start_label.grid(row=0, column=0, sticky='w', pady=5)

        start_var = tk.StringVar(value=next_mon.strftime("%Y-%m-%d"))
        start_entry = tk.Entry(week_frame,
                              textvariable=start_var,
                              font=('Segoe UI', 10),
                              width=15)
        start_entry.grid(row=0, column=1, padx=10, pady=5)

        # Week end date
        end_label = tk.Label(week_frame,
                            text="Week End (Sunday):",
                            font=('Segoe UI', 10),
                            bg='#ffffff')
        end_label.grid(row=1, column=0, sticky='w', pady=5)

        end_var = tk.StringVar(value=next_sun.strftime("%Y-%m-%d"))
        end_entry = tk.Entry(week_frame,
                            textvariable=end_var,
                            font=('Segoe UI', 10),
                            width=15)
        end_entry.grid(row=1, column=1, padx=10, pady=5)

        # Help text
        help_label = tk.Label(week_frame,
                             text="Format: YYYY-MM-DD (e.g., 2025-11-03)",
                             font=('Segoe UI', 8),
                             bg='#ffffff',
                             fg='#6b7280')
        help_label.grid(row=2, column=0, columnspan=2, pady=(5, 0))

        # Preferred days frame
        days_frame = tk.LabelFrame(main_frame,
                                  text="Preferred Meeting Days",
                                  font=('Segoe UI', 11, 'bold'),
                                  bg='#ffffff',
                                  fg='#374151',
                                  padx=15,
                                  pady=15)
        days_frame.pack(fill='x', pady=10)

        days_label = tk.Label(days_frame,
                             text="Preferred days (comma-separated):",
                             font=('Segoe UI', 10),
                             bg='#ffffff')
        days_label.pack(anchor='w')

        days_var = tk.StringVar(value="Wed,Thu")
        days_entry = tk.Entry(days_frame,
                             textvariable=days_var,
                             font=('Segoe UI', 10),
                             width=30)
        days_entry.pack(pady=5)

        days_help = tk.Label(days_frame,
                            text="Examples: Mon,Tue,Wed  or  Thu,Fri  or  Wed",
                            font=('Segoe UI', 8),
                            bg='#ffffff',
                            fg='#6b7280')
        days_help.pack()

        # Button frame
        btn_frame = tk.Frame(main_frame, bg='#f9fafb')
        btn_frame.pack(pady=20)

        def on_ok():
            try:
                # Validate dates
                start_date = date.fromisoformat(start_var.get().strip())
                end_date = date.fromisoformat(end_var.get().strip())

                if end_date < start_date:
                    messagebox.showerror("Invalid Dates",
                                       "End date must be after start date!")
                    return

                # Validate days
                days_input = days_var.get().strip()
                if not days_input:
                    messagebox.showerror("Invalid Days",
                                       "Please enter at least one preferred day!")
                    return

                # Store result
                result['cancelled'] = False
                result['week_range'] = f"{start_date.isoformat()}..{end_date.isoformat()}"
                result['preferred_days'] = days_input
                dialog.destroy()

            except ValueError as e:
                messagebox.showerror("Invalid Date",
                                   f"Invalid date format!\n\nUse YYYY-MM-DD format (e.g., 2025-11-03)")

        def on_cancel():
            result['cancelled'] = True
            dialog.destroy()

        ok_btn = tk.Button(btn_frame,
                          text="OK",
                          command=on_ok,
                          bg='#3b82f6',
                          fg='white',
                          font=('Segoe UI', 10, 'bold'),
                          padx=30,
                          pady=8,
                          relief='flat',
                          cursor='hand2')
        ok_btn.pack(side='left', padx=10)

        cancel_btn = tk.Button(btn_frame,
                              text="Cancel",
                              command=on_cancel,
                              bg='#6b7280',
                              fg='white',
                              font=('Segoe UI', 10, 'bold'),
                              padx=30,
                              pady=8,
                              relief='flat',
                              cursor='hand2')
        cancel_btn.pack(side='left', padx=10)

        # Wait for dialog to close
        self.root.wait_window(dialog)

        # Return result or None if cancelled
        if result['cancelled']:
            return None
        return result

    def run_script_with_command(self, cmd, step_name):
        """Run script with custom command in background thread"""
        self.is_processing = True
        self.update_status(f"Running {step_name}...", self.colors['warning'])

        try:
            # Run the script with custom command (with STARTUPINFO to hide console windows)
            process = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                bufsize=0,  # Unbuffered to reduce console activity
                startupinfo=get_subprocess_startupinfo(),
                creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == 'win32' else 0,
                shell=False
            )

            # Stream output WITHOUT update calls
            for line in process.stdout:
                self.output_text.insert(tk.END, line)
                self.output_text.see(tk.END)
                # No GUI update here - let Tkinter handle it naturally

            # Wait for completion
            process.wait()

            # Handle result
            if process.returncode == 0:
                self.output_text.insert(tk.END, f"\n\n[OK] {step_name} completed successfully!\n")
                self.output_text.see(tk.END)
                self.update_status(f"{step_name} complete!", self.colors['success'])
                messagebox.showinfo("Success", f"{step_name} completed successfully!")
            else:
                self.output_text.insert(tk.END, f"\n\n[FAIL] Process failed with exit code {process.returncode}\n")
                self.output_text.see(tk.END)
                self.update_status(f"{step_name} failed!", self.colors['error'])
                messagebox.showerror("Error", f"{step_name} failed with exit code {process.returncode}")

        except Exception as e:
            self.output_text.insert(tk.END, f"\n\n[ERROR] {str(e)}\n")
            self.output_text.see(tk.END)
            self.update_status("Error occurred!", self.colors['error'])
            messagebox.showerror("Error", f"An error occurred:\n{str(e)}")

        finally:
            self.is_processing = False
            # Safely stop progress bar
            try:
                if hasattr(self, 'progress') and self.progress.winfo_exists():
                    self.progress.stop()
            except:
                pass
            if hasattr(self, 'close_btn') and self.close_btn.winfo_exists():
                self.close_btn.config(state='normal')
            if hasattr(self, 'process_window') and self.process_window.winfo_exists():
                self.update_status("Ready", '#374151')

    def setup_gmail_api(self):
        """Show Gmail API setup instructions"""
        messagebox.showinfo(
            "Gmail API Setup",
            "Gmail API Setup Guide\n\n"
            "The Gmail API is used for:\n"
            "• Step 4: Send Meeting Requests\n"
            "• Step 6: Send Confirmations\n"
            "• Step 11: Send Follow-ups\n"
            "• Step 12: Email CVR to Supervisor\n\n"
            "📋 QUICK SETUP STEPS:\n\n"
            "1. Go to: https://console.cloud.google.com\n"
            "2. Create a new project (or select existing)\n"
            "3. Enable 'Gmail API'\n"
            "4. Go to 'Credentials' → 'Create Credentials' → 'OAuth client ID'\n"
            "5. Application type: 'Desktop app'\n"
            "6. Download the JSON file\n"
            "7. Rename it to: gmail_oauth_client.json\n"
            "8. Save to: C:\\GoogleSync\\GuardianShip_App\\Config\\API\\\n\n"
            "⚡ FIRST RUN:\n"
            "When you first use Steps 4, 6, 11, or 12, a browser will open.\n"
            "Sign in with your Google account and click 'Allow'.\n"
            "The authorization will be saved automatically.\n\n"
            "💡 TIP: This is a one-time setup!"
        )

    def setup_people_calendar_api(self):
        """Show People/Calendar API setup instructions"""
        messagebox.showinfo(
            "People & Calendar API Setup",
            "People & Calendar API Setup Guide\n\n"
            "These APIs are used for:\n"
            "• Step 5: Add Contacts (People API)\n"
            "• Step 7: Schedule Calendar Events (Calendar + Drive API)\n\n"
            "📋 QUICK SETUP STEPS:\n\n"
            "1. Go to: https://console.cloud.google.com\n"
            "2. Select your project\n"
            "3. Enable these APIs:\n"
            "   - People API\n"
            "   - Google Calendar API\n"
            "   - Google Drive API\n"
            "4. Go to 'Credentials' → 'Create Credentials' → 'OAuth client ID'\n"
            "5. Application type: 'Desktop app'\n"
            "6. Download the JSON file\n"
            "7. Rename it to: client_secret_calendar.json\n"
            "8. Save to: C:\\GoogleSync\\GuardianShip_App\\Config\\API\\\n\n"
            "⚡ FIRST RUN:\n"
            "When you first use Steps 5 or 7, a browser will open.\n"
            "Sign in with your Google account and click 'Allow'.\n"
            "The authorization will be saved automatically.\n\n"
            "💡 TIP: One OAuth file works for both APIs!"
        )

    def run_step4_script(self, cmd):
        """Run Step 4 script with custom success message about Gmail drafts"""
        self.is_processing = True
        self.update_status("Running Send Meeting Requests...", self.colors['warning'])

        try:
            # Run the script with custom command (with STARTUPINFO to hide console windows)
            process = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                bufsize=0,  # Unbuffered to reduce console activity
                startupinfo=get_subprocess_startupinfo(),
                creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == 'win32' else 0,
                shell=False
            )

            # Stream output WITHOUT update calls
            for line in process.stdout:
                self.output_text.insert(tk.END, line)
                self.output_text.see(tk.END)
                # No GUI update here - let Tkinter handle it naturally

            # Wait for completion
            process.wait()

            # Handle result with custom message for Step 4
            if process.returncode == 0:
                self.output_text.insert(tk.END, f"\n\n[OK] Send Meeting Requests completed successfully!\n")
                self.output_text.see(tk.END)
                self.update_status("Send Meeting Requests complete!", self.colors['success'])

                # Custom success message with instructions
                messagebox.showinfo(
                    "Meeting Request Emails Sent!",
                    "Meeting request emails have been sent successfully!\n\n"
                    "📧 WHAT HAPPENED:\n"
                    "• Emails sent to all guardians via Gmail\n"
                    "• Excel 'emailsent' column marked with today's date\n"
                    "• Text copies automatically saved:\n"
                    "  - To client folders (if folder exists)\n"
                    "  - To _Correspondence_Pending (if folder not found)\n\n"
                    "📁 NEXT STEPS:\n"
                    "• Check _Correspondence_Pending folder\n"
                    "• Move any pending emails to correct client folders\n"
                    "• Wait for guardian responses\n\n"
                    "💡 TIP: Guardian responses will appear in your Gmail inbox"
                )
            else:
                self.output_text.insert(tk.END, f"\n\n[FAIL] Process failed with exit code {process.returncode}\n")
                self.output_text.see(tk.END)
                self.update_status("Send Meeting Requests failed!", self.colors['error'])
                messagebox.showerror("Error", f"Send Meeting Requests failed with exit code {process.returncode}")

        except Exception as e:
            self.output_text.insert(tk.END, f"\n\n[ERROR] {str(e)}\n")
            self.output_text.see(tk.END)
            self.update_status("Error occurred!", self.colors['error'])
            messagebox.showerror("Error", f"An error occurred:\n{str(e)}")

        finally:
            self.is_processing = False
            # Safely stop progress bar
            try:
                if hasattr(self, 'progress') and self.progress.winfo_exists():
                    self.progress.stop()
            except:
                pass
            if hasattr(self, 'close_btn') and self.close_btn.winfo_exists():
                self.close_btn.config(state='normal')
            if hasattr(self, 'process_window') and self.process_window.winfo_exists():
                self.update_status("Ready", '#374151')

    def open_ai_help(self):
        """Open AI Help Assistant with app context"""
        ai_window = tk.Toplevel(self.root)
        ai_window.title("Live Tech Support")
        ai_window.geometry("800x700")
        ai_window.transient(self.root)

        # Title
        title = tk.Label(ai_window,
                        text="🆘 Live Tech Support - AI Assistants",
                        font=('Segoe UI', 16, 'bold'),
                        bg='#2563eb',
                        fg='white',
                        pady=15)
        title.pack(fill='x')

        # Instructions
        instructions_frame = tk.Frame(ai_window, bg='#f0f4f8')
        instructions_frame.pack(fill='x', padx=20, pady=20)

        instructions = tk.Label(instructions_frame,
                               text="Get instant help from AI assistants that know everything about this app!\n\n"
                                    "The AI has been pre-loaded with:\n"
                                    "• Complete app documentation\n"
                                    "• All 14 workflow steps\n"
                                    "• Google Vision setup guide\n"
                                    "• Common errors and solutions\n"
                                    "• Troubleshooting tips\n\n"
                                    "Just copy the context, paste it into the AI, and ask your question!",
                               font=('Segoe UI', 10),
                               bg='#f0f4f8',
                               justify='left')
        instructions.pack(padx=15, pady=15)

        # Choose AI section
        choice_frame = tk.LabelFrame(ai_window, text="Choose Your AI Assistant", font=('Segoe UI', 12, 'bold'), padx=20, pady=15)
        choice_frame.pack(fill='both', expand=True, padx=20, pady=10)

        # Claude
        claude_frame = tk.Frame(choice_frame, bg='white', relief='raised', borderwidth=2)
        claude_frame.pack(fill='x', pady=10)

        claude_title = tk.Label(claude_frame,
                               text="Claude (Anthropic) - Recommended for Technical Issues",
                               font=('Segoe UI', 11, 'bold'),
                               bg='white')
        claude_title.pack(anchor='w', padx=15, pady=(10, 5))

        claude_desc = tk.Label(claude_frame,
                              text="• Best for step-by-step technical guidance\n"
                                   "• Excellent at troubleshooting errors\n"
                                   "• Can walk you through complex setups\n"
                                   "• Free: Claude.ai",
                              font=('Segoe UI', 9),
                              bg='white',
                              justify='left')
        claude_desc.pack(anchor='w', padx=15, pady=(0, 10))

        claude_btn = ttk.Button(claude_frame,
                               text="Open Claude AI",
                               command=lambda: self.open_claude_with_context())
        claude_btn.pack(anchor='w', padx=15, pady=(0, 10))

        # ChatGPT
        chatgpt_frame = tk.Frame(choice_frame, bg='white', relief='raised', borderwidth=2)
        chatgpt_frame.pack(fill='x', pady=10)

        chatgpt_title = tk.Label(chatgpt_frame,
                                text="ChatGPT (OpenAI) - Good for Code Questions",
                                font=('Segoe UI', 11, 'bold'),
                                bg='white')
        chatgpt_title.pack(anchor='w', padx=15, pady=(10, 5))

        chatgpt_desc = tk.Label(chatgpt_frame,
                               text="• Good for general questions\n"
                                    "• Can analyze and explain code\n"
                                    "• Helpful for Python errors\n"
                                    "• Free: ChatGPT.com",
                               font=('Segoe UI', 9),
                               bg='white',
                               justify='left')
        chatgpt_desc.pack(anchor='w', padx=15, pady=(0, 10))

        chatgpt_btn = ttk.Button(chatgpt_frame,
                                text="Open ChatGPT",
                                command=lambda: self.open_chatgpt_with_context())
        chatgpt_btn.pack(anchor='w', padx=15, pady=(0, 10))

        # Gemini
        gemini_frame = tk.Frame(choice_frame, bg='white', relief='raised', borderwidth=2)
        gemini_frame.pack(fill='x', pady=10)

        gemini_title = tk.Label(gemini_frame,
                               text="Gemini (Google) - Best for Google API Issues",
                               font=('Segoe UI', 11, 'bold'),
                               bg='white')
        gemini_title.pack(anchor='w', padx=15, pady=(10, 5))

        gemini_desc = tk.Label(gemini_frame,
                              text="• Specialist in Google Vision API setup\n"
                                   "• Integrates with Google services\n"
                                   "• Good for Gmail/Calendar issues\n"
                                   "• Free: Gemini.google.com",
                              font=('Segoe UI', 9),
                              bg='white',
                              justify='left')
        gemini_desc.pack(anchor='w', padx=15, pady=(0, 10))

        gemini_btn = ttk.Button(gemini_frame,
                               text="Open Gemini",
                               command=lambda: self.open_gemini_with_context())
        gemini_btn.pack(anchor='w', padx=15, pady=(0, 10))

        # How it works
        how_frame = tk.LabelFrame(ai_window, text="Example Questions You Can Ask", font=('Segoe UI', 10, 'bold'), padx=15, pady=10)
        how_frame.pack(fill='x', padx=20, pady=10)

        how_text = tk.Label(how_frame,
                           text="• 'I can't connect to Google Vision, walk me through setup step by step'\n"
                                "• 'Step 1 is failing with error XYZ, how do I fix it?'\n"
                                "• 'My emails to supervisor aren't working, what should I check?'\n"
                                "• 'How do I fix OCR errors in the Excel file?'\n"
                                "• 'The JSON key file won't install, what am I doing wrong?'\n"
                                "• 'Can you show me screenshots of the Google Cloud Console setup?'",
                           font=('Segoe UI', 9),
                           justify='left')
        how_text.pack(anchor='w')

        # Close button
        close_btn = ttk.Button(ai_window, text="Close", command=ai_window.destroy)
        close_btn.pack(pady=10)

    def open_claude_with_context(self):
        """Open Claude.ai with pre-loaded context"""
        import webbrowser

        # Read context file
        context_file = Path(__file__).parent / "AI_HELP_CONTEXT.txt"

        if context_file.exists():
            with open(context_file, 'r', encoding='utf-8') as f:
                context = f.read()

            # Create window with context
            context_window = tk.Toplevel(self.root)
            context_window.title("Claude AI - Copy This Context")
            context_window.geometry("800x600")

            title = tk.Label(context_window,
                           text="Step 1: Copy the context below (it's already selected for you)",
                           font=('Segoe UI', 12, 'bold'),
                           bg='#2563eb',
                           fg='white',
                           pady=10)
            title.pack(fill='x')

            instructions = tk.Label(context_window,
                                   text="Step 2: Click 'Copy & Open Claude' button below\n"
                                        "Step 3: When Claude opens, paste the context (Ctrl+V)\n"
                                        "Step 4: After pasting, type your question!",
                                   font=('Segoe UI', 10),
                                   pady=10)
            instructions.pack()

            text_frame = tk.Frame(context_window)
            text_frame.pack(fill='both', expand=True, padx=20, pady=10)

            context_text = scrolledtext.ScrolledText(text_frame,
                                                    wrap=tk.WORD,
                                                    font=('Consolas', 8))
            context_text.pack(fill='both', expand=True)
            context_text.insert('1.0', context)
            context_text.tag_add('sel', '1.0', 'end')  # Select all
            context_text.focus()  # Focus for easy Ctrl+C

            def copy_and_open():
                self.root.clipboard_clear()
                self.root.clipboard_append(context)
                messagebox.showinfo("Ready!", "✅ Context copied to clipboard!\n\nClaude.ai will open now.\n\n1. Paste the context (Ctrl+V)\n2. Ask your question!")
                webbrowser.open('https://claude.ai/new')

            btn_frame = tk.Frame(context_window)
            btn_frame.pack(pady=10)

            copy_btn = ttk.Button(btn_frame,
                                 text="Copy Context & Open Claude",
                                 command=copy_and_open)
            copy_btn.pack(side='left', padx=5)

            close_btn = ttk.Button(btn_frame,
                                  text="Close",
                                  command=context_window.destroy)
            close_btn.pack(side='left', padx=5)
        else:
            messagebox.showerror("Error", "AI context file not found.\n\nPlease reinstall the app.")

    def open_chatgpt_with_context(self):
        """Open ChatGPT with pre-loaded context"""
        import webbrowser

        context_file = Path(__file__).parent / "AI_HELP_CONTEXT.txt"

        if context_file.exists():
            with open(context_file, 'r', encoding='utf-8') as f:
                context = f.read()

            # Create window with context
            context_window = tk.Toplevel(self.root)
            context_window.title("ChatGPT - Copy This Context")
            context_window.geometry("800x600")

            title = tk.Label(context_window,
                           text="Step 1: Copy the context below (it's already selected for you)",
                           font=('Segoe UI', 12, 'bold'),
                           bg='#10a37f',
                           fg='white',
                           pady=10)
            title.pack(fill='x')

            instructions = tk.Label(context_window,
                                   text="Step 2: Click 'Copy & Open ChatGPT' button below\n"
                                        "Step 3: When ChatGPT opens, paste the context (Ctrl+V)\n"
                                        "Step 4: After pasting, type your question!",
                                   font=('Segoe UI', 10),
                                   pady=10)
            instructions.pack()

            text_frame = tk.Frame(context_window)
            text_frame.pack(fill='both', expand=True, padx=20, pady=10)

            context_text = scrolledtext.ScrolledText(text_frame,
                                                    wrap=tk.WORD,
                                                    font=('Consolas', 8))
            context_text.pack(fill='both', expand=True)
            context_text.insert('1.0', context)
            context_text.tag_add('sel', '1.0', 'end')
            context_text.focus()

            def copy_and_open():
                self.root.clipboard_clear()
                self.root.clipboard_append(context)
                messagebox.showinfo("Ready!", "✅ Context copied to clipboard!\n\nChatGPT will open now.\n\n1. Paste the context (Ctrl+V)\n2. Ask your question!")
                webbrowser.open('https://chatgpt.com/')

            btn_frame = tk.Frame(context_window)
            btn_frame.pack(pady=10)

            copy_btn = ttk.Button(btn_frame,
                                 text="Copy Context & Open ChatGPT",
                                 command=copy_and_open)
            copy_btn.pack(side='left', padx=5)

            close_btn = ttk.Button(btn_frame,
                                  text="Close",
                                  command=context_window.destroy)
            close_btn.pack(side='left', padx=5)
        else:
            messagebox.showerror("Error", "AI context file not found.\n\nPlease reinstall the app.")

    def open_gemini_with_context(self):
        """Open Gemini with pre-loaded context"""
        import webbrowser

        context_file = Path(__file__).parent / "AI_HELP_CONTEXT.txt"

        if context_file.exists():
            with open(context_file, 'r', encoding='utf-8') as f:
                context = f.read()

            # Create window with context
            context_window = tk.Toplevel(self.root)
            context_window.title("Gemini - Copy This Context")
            context_window.geometry("800x600")

            title = tk.Label(context_window,
                           text="Step 1: Copy the context below (it's already selected for you)",
                           font=('Segoe UI', 12, 'bold'),
                           bg='#4285f4',
                           fg='white',
                           pady=10)
            title.pack(fill='x')

            instructions = tk.Label(context_window,
                                   text="Step 2: Click 'Copy & Open Gemini' button below\n"
                                        "Step 3: When Gemini opens, paste the context (Ctrl+V)\n"
                                        "Step 4: After pasting, type your question!",
                                   font=('Segoe UI', 10),
                                   pady=10)
            instructions.pack()

            text_frame = tk.Frame(context_window)
            text_frame.pack(fill='both', expand=True, padx=20, pady=10)

            context_text = scrolledtext.ScrolledText(text_frame,
                                                    wrap=tk.WORD,
                                                    font=('Consolas', 8))
            context_text.pack(fill='both', expand=True)
            context_text.insert('1.0', context)
            context_text.tag_add('sel', '1.0', 'end')
            context_text.focus()

            def copy_and_open():
                self.root.clipboard_clear()
                self.root.clipboard_append(context)
                messagebox.showinfo("Ready!", "✅ Context copied to clipboard!\n\nGemini will open now.\n\n1. Paste the context (Ctrl+V)\n2. Ask your question!")
                webbrowser.open('https://gemini.google.com/')

            btn_frame = tk.Frame(context_window)
            btn_frame.pack(pady=10)

            copy_btn = ttk.Button(btn_frame,
                                 text="Copy Context & Open Gemini",
                                 command=copy_and_open)
            copy_btn.pack(side='left', padx=5)

            close_btn = ttk.Button(btn_frame,
                                  text="Close",
                                  command=context_window.destroy)
            close_btn.pack(side='left', padx=5)
        else:
            messagebox.showerror("Error", "AI context file not found.\n\nPlease reinstall the app.")


def main():
    root = tk.Tk()
    app = GuardianShipApp(root)

    # Bring window to foreground and maximize
    root.lift()  # Bring window above other windows
    root.attributes('-topmost', True)  # Temporarily make topmost
    root.after_idle(root.attributes, '-topmost', False)  # Remove topmost after display
    root.focus_force()  # Force focus on the window

    # Maximize window on Windows
    try:
        root.state('zoomed')  # Windows maximize
    except:
        pass  # If zoomed not supported, use geometry setting

    # Check for updates on startup (in background thread)
    if AUTO_UPDATE_ENABLED:
        try:
            updater = AutoUpdater(
                current_version=__version__,
                github_repo="your-username/court-visitor-app"  # TODO: Update with your GitHub repo
            )
            # Check for updates after window is displayed (500ms delay)
            root.after(500, lambda: updater.check_on_startup(parent=root, silent=True))
        except Exception as e:
            print(f"Auto-update check failed: {e}")

    root.mainloop()


if __name__ == "__main__":
    main()
