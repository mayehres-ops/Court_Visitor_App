"""
EULA Acceptance Dialog
Shows End User License Agreement on first run and tracks acceptance.
"""

import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
from pathlib import Path
import json
from datetime import datetime


class EULADialog:
    """Dialog for EULA acceptance."""

    def __init__(self, parent=None):
        """
        Initialize EULA dialog.

        Args:
            parent: Parent window (if None, creates standalone window)
        """
        self.result = False

        # Create window
        if parent:
            self.root = tk.Toplevel(parent)
        else:
            self.root = tk.Tk()

        self.root.title("End User License Agreement")

        # Window size
        width = 700
        height = 600
        self.root.geometry(f"{width}x{height}")
        self.root.resizable(True, True)

        # Center window
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'+{x}+{y}')

        # Make modal
        self.root.transient(parent)
        self.root.grab_set()

        # Load EULA text
        self.eula_text = self._load_eula()

        # Build UI
        self._build_ui()

        # Handle window close
        self.root.protocol("WM_DELETE_WINDOW", self._on_decline)

    def _load_eula(self):
        """Load EULA text from file."""
        try:
            eula_file = Path(__file__).parent.parent / "EULA.txt"
            if eula_file.exists():
                with open(eula_file, 'r', encoding='utf-8') as f:
                    return f.read()
            else:
                return "EULA file not found. Please contact support."
        except Exception as e:
            return f"Error loading EULA: {e}"

    def _build_ui(self):
        """Build the user interface."""
        # Header
        header_frame = tk.Frame(self.root, bg='#dc2626', height=80)
        header_frame.pack(fill='x')
        header_frame.pack_propagate(False)

        header_text = "Court Visitor App\nEnd User License Agreement"
        header_label = tk.Label(
            header_frame,
            text=header_text,
            font=('Segoe UI', 14, 'bold'),
            bg='#dc2626',
            fg='white',
            justify='center'
        )
        header_label.pack(expand=True)

        # Main content frame
        content_frame = tk.Frame(self.root, padx=20, pady=20)
        content_frame.pack(fill='both', expand=True)

        # Instructions
        instructions = tk.Label(
            content_frame,
            text="Please read the following license agreement carefully.\n"
                 "You must accept the terms to use this software.",
            font=('Segoe UI', 10),
            justify='left',
            wraplength=650
        )
        instructions.pack(anchor='w', pady=(0, 10))

        # EULA text area (scrolled text)
        text_frame = tk.Frame(content_frame, relief='sunken', borderwidth=1)
        text_frame.pack(fill='both', expand=True)

        self.text_widget = scrolledtext.ScrolledText(
            text_frame,
            font=('Courier New', 9),
            wrap='word',
            padx=10,
            pady=10
        )
        self.text_widget.pack(fill='both', expand=True)
        self.text_widget.insert('1.0', self.eula_text)
        self.text_widget.config(state='disabled')  # Read-only

        # Acceptance checkbox
        accept_frame = tk.Frame(content_frame)
        accept_frame.pack(fill='x', pady=(15, 10))

        self.accept_var = tk.BooleanVar(value=False)
        self.accept_checkbox = tk.Checkbutton(
            accept_frame,
            text="I have read and agree to the terms of the End User License Agreement",
            variable=self.accept_var,
            font=('Segoe UI', 10, 'bold'),
            command=self._on_checkbox_change
        )
        self.accept_checkbox.pack(anchor='w')

        # Warning label
        self.warning_label = tk.Label(
            content_frame,
            text="⚠ You must accept the agreement to continue",
            font=('Segoe UI', 9),
            fg='#dc2626'
        )
        self.warning_label.pack(anchor='w', pady=(5, 10))

        # Button frame
        btn_frame = tk.Frame(self.root, pady=15, bg='#f3f4f6')
        btn_frame.pack(side='bottom', fill='x')

        # Accept button (disabled initially)
        self.accept_btn = tk.Button(
            btn_frame,
            text="Accept and Continue",
            command=self._on_accept,
            font=('Segoe UI', 10, 'bold'),
            bg='#16a34a',
            fg='white',
            padx=30,
            pady=8,
            cursor='hand2',
            state='disabled'
        )
        self.accept_btn.pack(side='right', padx=(5, 30))

        # Decline button
        decline_btn = tk.Button(
            btn_frame,
            text="Decline and Exit",
            command=self._on_decline,
            font=('Segoe UI', 10),
            padx=30,
            pady=8,
            cursor='hand2'
        )
        decline_btn.pack(side='right', padx=5)

    def _on_checkbox_change(self):
        """Handle checkbox state change."""
        if self.accept_var.get():
            self.accept_btn.config(state='normal', bg='#16a34a')
            self.warning_label.config(text="")
        else:
            self.accept_btn.config(state='disabled', bg='#9ca3af')
            self.warning_label.config(text="⚠ You must accept the agreement to continue")

    def _on_accept(self):
        """Handle accept button click."""
        if not self.accept_var.get():
            messagebox.showwarning(
                "Agreement Required",
                "You must check the box to accept the agreement.",
                parent=self.root
            )
            return

        self.result = True
        self._save_acceptance()
        self.root.destroy()

    def _on_decline(self):
        """Handle decline button click."""
        response = messagebox.askyesno(
            "Exit Application",
            "If you decline the license agreement, the application will exit.\n\n"
            "Are you sure you want to exit?",
            parent=self.root,
            icon='warning'
        )
        if response:
            self.result = False
            self.root.destroy()

    def _save_acceptance(self):
        """Save EULA acceptance to app settings."""
        try:
            from app_paths import get_app_paths
            app_paths = get_app_paths()
            settings_file = app_paths.CONFIG_DIR / "app_settings.json"

            # Load existing settings
            settings = {}
            if settings_file.exists():
                with open(settings_file, 'r', encoding='utf-8') as f:
                    settings = json.load(f)

            # Update EULA acceptance
            settings['eula_accepted'] = True
            settings['eula_accepted_date'] = datetime.now().isoformat()

            # Save
            settings_file.parent.mkdir(parents=True, exist_ok=True)
            with open(settings_file, 'w', encoding='utf-8') as f:
                json.dump(settings, f, indent=2)

        except Exception as e:
            print(f"Error saving EULA acceptance: {e}")

    def show(self):
        """Show the dialog and return the result."""
        self.root.mainloop()
        return self.result


def show_eula_dialog(parent=None) -> bool:
    """
    Show EULA dialog and return acceptance status.

    Args:
        parent: Parent window (optional)

    Returns:
        True if accepted, False if declined
    """
    dialog = EULADialog(parent)
    return dialog.show()


def check_eula_acceptance() -> bool:
    """
    Check if EULA has been accepted.

    Returns:
        True if EULA has been accepted, False otherwise
    """
    try:
        from app_paths import get_app_paths
        app_paths = get_app_paths()
        settings_file = app_paths.CONFIG_DIR / "app_settings.json"

        if settings_file.exists():
            with open(settings_file, 'r', encoding='utf-8') as f:
                settings = json.load(f)
            return settings.get('eula_accepted', False)
        return False
    except Exception:
        return False


if __name__ == "__main__":
    # Test the dialog
    accepted = show_eula_dialog()
    print(f"EULA accepted: {accepted}")

    if accepted:
        print("User accepted EULA - application can continue")
    else:
        print("User declined EULA - application should exit")
