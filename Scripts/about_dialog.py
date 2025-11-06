"""
About Dialog
Displays copyright, version, and legal information.
"""

import tkinter as tk
from tkinter import ttk


class AboutDialog:
    """About dialog with copyright and version info."""

    def __init__(self, parent=None, version="1.0.0"):
        """
        Initialize About dialog.

        Args:
            parent: Parent window (if None, creates standalone window)
            version: Application version string
        """
        # Create window
        if parent:
            self.root = tk.Toplevel(parent)
        else:
            self.root = tk.Tk()

        self.root.title("About Court Visitor App")

        # Window size
        width = 500
        height = 550
        self.root.geometry(f"{width}x{height}")
        self.root.resizable(False, False)

        # Center window
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'+{x}+{y}')

        # Make modal
        if parent:
            self.root.transient(parent)
            self.root.grab_set()

        self.version = version

        # Build UI
        self._build_ui()

    def _build_ui(self):
        """Build the user interface."""
        # Header with app icon/name
        header_frame = tk.Frame(self.root, bg='#667eea', height=120)
        header_frame.pack(fill='x')
        header_frame.pack_propagate(False)

        app_name_label = tk.Label(
            header_frame,
            text="Court Visitor App",
            font=('Segoe UI', 18, 'bold'),
            bg='#667eea',
            fg='white'
        )
        app_name_label.pack(pady=(30, 5))

        version_label = tk.Label(
            header_frame,
            text=f"Version {self.version}",
            font=('Segoe UI', 11),
            bg='#667eea',
            fg='white'
        )
        version_label.pack()

        # Main content
        content_frame = tk.Frame(self.root, padx=30, pady=20)
        content_frame.pack(fill='both', expand=True)

        # Copyright
        copyright_label = tk.Label(
            content_frame,
            text="Copyright © 2024-2025 GuardianShip Easy, LLC\nAll rights reserved.",
            font=('Segoe UI', 10, 'bold'),
            justify='center'
        )
        copyright_label.pack(pady=(10, 15))

        # Description
        description_text = (
            "Professional workflow automation software for\n"
            "Court-Appointed Visitors in Travis County, Texas."
        )
        description_label = tk.Label(
            content_frame,
            text=description_text,
            font=('Segoe UI', 10),
            justify='center',
            fg='#6b7280'
        )
        description_label.pack(pady=(0, 20))

        # Separator
        separator = ttk.Separator(content_frame, orient='horizontal')
        separator.pack(fill='x', pady=15)

        # Legal notices
        legal_frame = tk.Frame(content_frame)
        legal_frame.pack(fill='x', pady=10)

        legal_text = (
            "PROPRIETARY AND CONFIDENTIAL\n\n"
            "This software and associated documentation files are\n"
            "proprietary and confidential. Unauthorized copying,\n"
            "distribution, or modification is strictly prohibited.\n\n"
            "NOT FOR RESALE\n\n"
            "This software is licensed for use by authorized\n"
            "Court Visitors only and may not be sold, rented,\n"
            "leased, or otherwise transferred to third parties."
        )

        legal_label = tk.Label(
            legal_frame,
            text=legal_text,
            font=('Segoe UI', 9),
            justify='center',
            fg='#374151'
        )
        legal_label.pack()

        # Separator
        separator2 = ttk.Separator(content_frame, orient='horizontal')
        separator2.pack(fill='x', pady=15)

        # Contact info
        contact_text = (
            "For support or licensing information:\n"
            "Email: support@guardianshipeasy.com\n"
            "Web: www.GuardianshipEasy.com"
        )
        contact_label = tk.Label(
            content_frame,
            text=contact_text,
            font=('Segoe UI', 9),
            justify='center',
            fg='#6b7280'
        )
        contact_label.pack(pady=(0, 15))

        # Close button
        close_btn = tk.Button(
            content_frame,
            text="Close",
            command=self.root.destroy,
            font=('Segoe UI', 10),
            bg='#667eea',
            fg='white',
            padx=40,
            pady=8,
            cursor='hand2'
        )
        close_btn.pack()

    def show(self):
        """Show the dialog."""
        self.root.mainloop()


def show_about_dialog(parent=None, version="1.0.0"):
    """
    Show About dialog.

    Args:
        parent: Parent window (optional)
        version: Application version string
    """
    dialog = AboutDialog(parent, version)
    dialog.show()


if __name__ == "__main__":
    # Test the dialog
    show_about_dialog(version="1.0.0")
