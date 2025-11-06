"""
Court Visitor App - Centralized Path Management
Copyright (c) 2024 GuardianShip Easy, LLC. All rights reserved.

This module provides a centralized way to manage all application paths.
Paths are dynamically determined based on where the app is installed,
NOT hardcoded to C:\\GoogleSync\\GuardianShip_App.

Usage:
    from app_paths import AppPaths

    paths = AppPaths()
    excel_path = paths.EXCEL_PATH
    template_path = paths.CVR_TEMPLATE_PATH
    new_files = paths.NEW_FILES_DIR
"""

import os
import sys
from pathlib import Path
import json

# Fix encoding for Windows console
if sys.platform == 'win32':
    import io
    try:
        if sys.stdout and hasattr(sys.stdout, 'buffer'):
            sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
        if sys.stderr and hasattr(sys.stderr, 'buffer'):
            sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
    except (AttributeError, TypeError):
        # GUI mode - stdout/stderr are None or don't have buffer
        pass


class AppPaths:
    """
    Centralized path management for the Court Visitor App.

    All paths are calculated relative to the application root directory,
    allowing the app to work from any installation location.
    """

    def __init__(self, app_root=None):
        """
        Initialize paths based on application root directory.

        Args:
            app_root: Optional custom root directory. If not provided,
                     auto-detects based on this file's location.
        """
        if app_root:
            self.APP_ROOT = Path(app_root)
        else:
            # Auto-detect: This file is in Scripts/, so go up one level
            self.APP_ROOT = Path(__file__).parent.parent.resolve()

        # Verify this looks like the app directory
        if not self._validate_app_root():
            print(f"WARNING: App root may be incorrect: {self.APP_ROOT}")

        # Initialize all paths
        self._setup_paths()

    def _validate_app_root(self):
        """Verify the app root looks correct."""
        # Check for key indicators
        indicators = [
            self.APP_ROOT / "guardianship_app.py",
            self.APP_ROOT / "Config",
            self.APP_ROOT / "App Data"
        ]
        return any(p.exists() for p in indicators)

    def _setup_paths(self):
        """Set up all application paths."""
        root = self.APP_ROOT

        # ==========================================
        # Main Application Files
        # ==========================================
        self.MAIN_APP = root / "guardianship_app.py"
        self.AUTO_UPDATER = root / "auto_updater.py"
        self.SETUP_WIZARD = root / "setup_wizard.py"

        # ==========================================
        # Configuration & Settings
        # ==========================================
        self.CONFIG_DIR = root / "Config"
        self.CONFIG_API_DIR = self.CONFIG_DIR / "API"
        self.APP_SETTINGS_FILE = self.CONFIG_DIR / "app_settings.json"

        # Google API credentials
        self.CREDENTIALS_FILE = self.CONFIG_API_DIR / "credentials.json"
        self.TOKEN_GMAIL_FILE = self.CONFIG_API_DIR / "token_gmail.json"
        self.TOKEN_CALENDAR_FILE = self.CONFIG_API_DIR / "token_calendar.json"
        self.TOKEN_SHEETS_FILE = self.CONFIG_API_DIR / "token_sheets.json"
        self.TOKEN_CONTACTS_FILE = self.CONFIG_API_DIR / "token_contacts.json"

        # ==========================================
        # Data Directories
        # ==========================================
        self.APP_DATA_DIR = root / "App Data"
        self.BACKUP_DIR = self.APP_DATA_DIR / "Backup"
        self.INBOX_DIR = self.APP_DATA_DIR / "Inbox"
        self.STAGING_DIR = self.APP_DATA_DIR / "Staging"
        self.TEMPLATES_DIR = self.APP_DATA_DIR / "Templates"

        # Excel database
        self.EXCEL_PATH = self.APP_DATA_DIR / "ward_guardian_info.xlsx"
        self.EXCEL_SHEET_NAME = "Sheet1"

        # ==========================================
        # Work Directories
        # ==========================================
        self.NEW_FILES_DIR = root / "New Files"
        self.NEW_CLIENTS_DIR = root / "New Clients"
        self.COMPLETED_DIR = root / "Completed"

        # ==========================================
        # Templates
        # ==========================================
        # Note: Templates can be in either App Data/Templates or Templates/
        templates_alt = root / "Templates"

        self.CVR_TEMPLATE_PATH = self._find_file([
            self.TEMPLATES_DIR / "Court Visitor Report fillable new.docx",
            templates_alt / "Court Visitor Report fillable new.docx"
        ])

        self.PAYMENT_FORM_TEMPLATE = self._find_file([
            self.TEMPLATES_DIR / "Court Visitor Payment Form TEMPLATE.docx",
            templates_alt / "Court Visitor Payment Form TEMPLATE.docx"
        ])

        self.MILEAGE_LOG_TEMPLATE = self._find_file([
            self.TEMPLATES_DIR / "MILEAGE LOG CV Visitors template.docx",
            templates_alt / "MILEAGE LOG CV Visitors template.docx"
        ])

        self.MAP_SHEET_TEMPLATE = self._find_file([
            self.TEMPLATES_DIR / "Ward Map Sheet.docx",
            templates_alt / "Ward Map Sheet.docx"
        ])

        # ==========================================
        # Scripts & Automation
        # ==========================================
        self.SCRIPTS_DIR = root / "Scripts"
        self.AUTOMATION_DIR = root / "Automation"

        # Main processing scripts
        self.OCR_SCRIPT = root / "guardian_extractor_claudecode20251023_bestever_11pm.py"
        self.GOOGLE_SHEETS_SCRIPT = root / "google_sheets_cvr_integration_fixed.py"
        self.EMAIL_CVR_SCRIPT = root / "email_cvr_to_supervisor.py"

        # Step scripts (Automation folder)
        self.CVR_BUILDER_SCRIPT = self.AUTOMATION_DIR / "Create CV report_move to folder" / "Scripts" / "build_cvr_from_excel_cc_working.py"
        self.FOLDER_BUILDER_SCRIPT = self.AUTOMATION_DIR / "CV Report_Folders Script" / "scripts" / "cvr_folder_builder.py"
        self.MAP_SHEET_SCRIPT = self.AUTOMATION_DIR / "Build Map Sheet" / "Scripts" / "build_map_sheet.py"
        self.PAYMENT_FORM_SCRIPT = self.AUTOMATION_DIR / "CV Payment Form Script" / "scripts" / "build_payment_forms_sdt.py"
        self.MILEAGE_LOG_SCRIPT = self.AUTOMATION_DIR / "Mileage Reimbursement Script" / "scripts" / "build_mileage_forms.py"
        self.EMAIL_REQUEST_SCRIPT = self.AUTOMATION_DIR / "Email Meeting Request" / "scripts" / "send_guardian_emails.py"
        self.CONFIRMATION_EMAIL_SCRIPT = self.AUTOMATION_DIR / "Appt Email Confirm" / "scripts" / "send_confirmation_email.py"
        self.CALENDAR_EVENT_SCRIPT = self.AUTOMATION_DIR / "Calendar appt send email conf" / "scripts" / "create_calendar_event.py"
        self.ADD_CONTACTS_SCRIPT = self.AUTOMATION_DIR / "Contacts - Guardians" / "scripts" / "add_guardians_to_contacts.py"
        self.FOLLOWUP_EMAIL_SCRIPT = self.AUTOMATION_DIR / "TX email to guardian" / "send_followups_picker.py"

        # Utility scripts
        self.CVR_CONTENT_CONTROL_UTILS = self.SCRIPTS_DIR / "cvr_content_control_utils.py"
        self.CONFIG_MANAGER = self.SCRIPTS_DIR / "app_config_manager.py"

        # ==========================================
        # Documentation
        # ==========================================
        self.DOCS_DIR = root / "Documentation"
        self.README_FILE = root / "README_FIRST.md"
        self.USER_MANUAL_FILE = root / "Court_Visitor_App_Manual_Updated.md"
        self.INSTALLATION_GUIDE = root / "END_USER_INSTALLATION_GUIDE.md"
        self.EULA_FILE = root / "EULA.txt"

    def _find_file(self, paths):
        """
        Find the first existing file from a list of possible paths.

        Args:
            paths: List of Path objects to check

        Returns:
            First existing path, or first path if none exist (for creation)
        """
        for path in paths:
            if path.exists():
                return path
        # Return first path if none exist
        return paths[0] if paths else None

    def create_directories(self):
        """Create all necessary directories if they don't exist."""
        dirs_to_create = [
            self.CONFIG_DIR,
            self.CONFIG_API_DIR,
            self.APP_DATA_DIR,
            self.BACKUP_DIR,
            self.INBOX_DIR,
            self.STAGING_DIR,
            self.TEMPLATES_DIR,
            self.NEW_FILES_DIR,
            self.NEW_CLIENTS_DIR,
            self.COMPLETED_DIR,
            self.SCRIPTS_DIR,
            self.AUTOMATION_DIR,
            self.DOCS_DIR,
        ]

        for directory in dirs_to_create:
            try:
                directory.mkdir(parents=True, exist_ok=True)
            except Exception as e:
                print(f"Warning: Could not create directory {directory}: {e}")

    def validate_critical_paths(self):
        """
        Validate that critical paths exist.

        Returns:
            dict: {path_name: exists_bool}
        """
        critical_paths = {
            "Excel Database": self.EXCEL_PATH,
            "Config Directory": self.CONFIG_DIR,
            "App Data Directory": self.APP_DATA_DIR,
            "New Files Directory": self.NEW_FILES_DIR,
            "Scripts Directory": self.SCRIPTS_DIR,
            "Automation Directory": self.AUTOMATION_DIR,
        }

        return {name: path.exists() for name, path in critical_paths.items()}

    def to_dict(self):
        """
        Export all paths as a dictionary (useful for debugging).

        Returns:
            dict: All paths as strings
        """
        return {
            attr: str(getattr(self, attr))
            for attr in dir(self)
            if not attr.startswith('_') and isinstance(getattr(self, attr), Path)
        }

    def print_summary(self):
        """Print a summary of all paths."""
        print("=" * 70)
        print("Court Visitor App - Path Configuration")
        print("=" * 70)
        print(f"\nApp Root: {self.APP_ROOT}")
        print(f"Valid: {self._validate_app_root()}")

        print("\n--- Critical Paths ---")
        validation = self.validate_critical_paths()
        for name, exists in validation.items():
            status = "✓" if exists else "✗"
            print(f"{status} {name}")

        print("\n--- Key Files ---")
        print(f"Excel: {self.EXCEL_PATH}")
        print(f"CVR Template: {self.CVR_TEMPLATE_PATH}")
        print(f"Config: {self.APP_SETTINGS_FILE}")


# Global singleton instance
_paths_instance = None

def get_app_paths(app_root=None):
    """
    Get the global AppPaths instance (singleton pattern).

    Args:
        app_root: Optional custom root directory (only used on first call)

    Returns:
        AppPaths instance
    """
    global _paths_instance
    if _paths_instance is None:
        _paths_instance = AppPaths(app_root)
    return _paths_instance


# Convenience functions for quick access
def get_excel_path():
    """Quick access to Excel database path."""
    return str(get_app_paths().EXCEL_PATH)

def get_templates_dir():
    """Quick access to templates directory."""
    return str(get_app_paths().TEMPLATES_DIR)

def get_new_files_dir():
    """Quick access to new files directory."""
    return str(get_app_paths().NEW_FILES_DIR)


# Test/debug mode
if __name__ == "__main__":
    paths = AppPaths()
    paths.print_summary()

    print("\n--- All Paths ---")
    for name, path in sorted(paths.to_dict().items()):
        print(f"{name}: {path}")
