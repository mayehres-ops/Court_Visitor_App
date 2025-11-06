"""
Court Visitor Settings Dialog
GUI for managing Court Visitor personal information.

Usage:
    from cv_settings_dialog import show_cv_settings_dialog

    # Returns True if settings were saved, False if cancelled
    saved = show_cv_settings_dialog()
"""

import tkinter as tk
from tkinter import ttk, messagebox
from typing import Optional
from cv_info_manager import CVInfoManager


class CVSettingsDialog:
    """Dialog for editing Court Visitor information."""

    def __init__(self, parent=None, first_run=False):
        """
        Initialize settings dialog.

        Args:
            parent: Parent window (if None, creates standalone window)
            first_run: If True, shows as first-run setup wizard
        """
        self.result = False
        self.first_run = first_run

        # Create window
        if parent:
            self.root = tk.Toplevel(parent)
        else:
            self.root = tk.Tk()

        # Set title based on mode
        title = "First-Time Setup - Court Visitor Information" if first_run else "Court Visitor Settings"
        self.root.title(title)

        # Window size
        width = 600
        height = 520 if first_run else 480
        self.root.geometry(f"{width}x{height}")
        self.root.resizable(False, False)

        # Center window
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'+{x}+{y}')

        # Load current info
        self.cv_manager = CVInfoManager()
        self.current_info = self.cv_manager.get_info()

        # Build UI
        self._build_ui()

        # Handle window close
        self.root.protocol("WM_DELETE_WINDOW", self._on_cancel)

    def _build_ui(self):
        """Build the user interface."""
        # Header
        header_frame = tk.Frame(self.root, bg='#667eea', height=80)
        header_frame.pack(fill='x')
        header_frame.pack_propagate(False)

        if self.first_run:
            header_text = "Welcome to Court Visitor App!\n\nPlease enter your information to personalize forms."
        else:
            header_text = "Court Visitor Information\n\nThis information is used to auto-fill your forms."

        header_label = tk.Label(
            header_frame,
            text=header_text,
            font=('Segoe UI', 12, 'bold'),
            bg='#667eea',
            fg='white',
            justify='center'
        )
        header_label.pack(expand=True)

        # Main content frame
        content_frame = tk.Frame(self.root, padx=30, pady=20)
        content_frame.pack(fill='both', expand=True)

        # Form fields
        self.entries = {}

        # Court Visitor Name
        self._add_field(content_frame, 0, "Court Visitor Name:", 'name', required=True)

        # Vendor Number
        self._add_field(content_frame, 1, "Vendor Number:", 'vendor_number')

        # GL Number
        self._add_field(content_frame, 2, "GL Number:", 'gl_number')

        # Cost Center
        self._add_field(content_frame, 3, "Cost Center Number:", 'cost_center')

        # Address Line 1
        self._add_field(content_frame, 4, "Address Line 1:", 'address_line1')

        # Address Line 2
        self._add_field(content_frame, 5, "Address Line 2:", 'address_line2')

        # Info text
        info_text = "These fields will be automatically filled in your mileage logs, payment forms, and court visitor reports."
        info_label = tk.Label(
            content_frame,
            text=info_text,
            font=('Segoe UI', 9),
            fg='gray',
            wraplength=500,
            justify='left'
        )
        info_label.grid(row=6, column=0, columnspan=2, pady=(15, 10), sticky='w')

        # Button frame
        btn_frame = tk.Frame(self.root, pady=15)
        btn_frame.pack(side='bottom', fill='x')

        # Save button
        save_text = "Save and Continue" if self.first_run else "Save"
        save_btn = tk.Button(
            btn_frame,
            text=save_text,
            command=self._on_save,
            font=('Segoe UI', 10, 'bold'),
            bg='#667eea',
            fg='white',
            padx=30,
            pady=8,
            cursor='hand2'
        )
        save_btn.pack(side='right', padx=(5, 30))

        # Cancel/Skip button
        if self.first_run:
            cancel_text = "Skip for Now"
        else:
            cancel_text = "Cancel"

        cancel_btn = tk.Button(
            btn_frame,
            text=cancel_text,
            command=self._on_cancel,
            font=('Segoe UI', 10),
            padx=30,
            pady=8,
            cursor='hand2'
        )
        cancel_btn.pack(side='right', padx=5)

        # Focus first field
        self.entries['name'].focus_set()

    def _add_field(self, parent, row, label_text, field_key, required=False):
        """Add a form field to the dialog."""
        # Label
        label = tk.Label(
            parent,
            text=label_text + (" *" if required else ""),
            font=('Segoe UI', 10, 'bold' if required else 'normal'),
            anchor='w'
        )
        label.grid(row=row, column=0, sticky='w', pady=(5, 2))

        # Entry
        entry = tk.Entry(
            parent,
            font=('Segoe UI', 10),
            width=50
        )
        entry.insert(0, self.current_info.get(field_key, ''))
        entry.grid(row=row, column=1, sticky='ew', pady=(5, 2))

        self.entries[field_key] = entry

        # Configure grid
        parent.grid_columnconfigure(1, weight=1)

    def _on_save(self):
        """Handle save button click."""
        # Collect data
        info = {}
        for key, entry in self.entries.items():
            info[key] = entry.get().strip()

        # Validate
        if not info['name']:
            messagebox.showerror(
                "Required Field",
                "Court Visitor Name is required.",
                parent=self.root
            )
            self.entries['name'].focus_set()
            return

        # Save
        if self.cv_manager.save_info(info):
            self.result = True
            self.root.destroy()
        else:
            messagebox.showerror(
                "Save Error",
                "Could not save settings. Please try again.",
                parent=self.root
            )

    def _on_cancel(self):
        """Handle cancel button click."""
        if self.first_run:
            # Confirm skip on first run
            response = messagebox.askyesno(
                "Skip Setup",
                "Are you sure you want to skip this setup?\n\n"
                "You can configure this later from Settings.",
                parent=self.root
            )
            if response:
                self.result = False
                self.root.destroy()
        else:
            self.result = False
            self.root.destroy()

    def show(self):
        """Show the dialog and return the result."""
        self.root.mainloop()
        return self.result


def show_cv_settings_dialog(parent=None, first_run=False) -> bool:
    """
    Show Court Visitor settings dialog.

    Args:
        parent: Parent window (optional)
        first_run: If True, shows as first-run wizard

    Returns:
        True if settings were saved, False if cancelled
    """
    dialog = CVSettingsDialog(parent, first_run)
    return dialog.show()


if __name__ == "__main__":
    # Test the dialog
    saved = show_cv_settings_dialog(first_run=True)
    print(f"Settings saved: {saved}")

    if saved:
        from cv_info_manager import CVInfoManager
        manager = CVInfoManager()
        info = manager.get_info()
        print("\nSaved information:")
        for key, value in info.items():
            print(f"  {key}: {value}")
