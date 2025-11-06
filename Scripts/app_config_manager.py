"""
Court Visitor App - Configuration Manager
Copyright (c) 2024 GuardianShip Easy, LLC. All rights reserved.

Manages app settings including Court Visitor name, EULA acceptance, and licensing.
"""

import json
import os
from pathlib import Path
from datetime import datetime
import tkinter as tk
from tkinter import simpledialog, messagebox

# Default configuration file path
CONFIG_FILE = Path(__file__).parent.parent / "Config" / "app_settings.json"

DEFAULT_SETTINGS = {
    "court_visitor_name": "",
    "first_run_complete": False,
    "eula_accepted": False,
    "eula_accepted_date": "",
    "license_key": "",
    "activation_date": "",
    "app_version": "1.0.0"
}


class AppConfigManager:
    """Manages application configuration and settings."""

    def __init__(self, config_path=None):
        """
        Initialize the configuration manager.

        Args:
            config_path: Optional custom path to config file
        """
        self.config_path = Path(config_path) if config_path else CONFIG_FILE
        self.settings = self.load_settings()

    def load_settings(self):
        """Load settings from JSON file, creating it if it doesn't exist."""
        try:
            if self.config_path.exists():
                with open(self.config_path, 'r') as f:
                    settings = json.load(f)
                    # Merge with defaults in case new settings were added
                    return {**DEFAULT_SETTINGS, **settings}
            else:
                # Create config directory if it doesn't exist
                self.config_path.parent.mkdir(parents=True, exist_ok=True)
                # Create with default settings
                self.save_settings(DEFAULT_SETTINGS)
                return DEFAULT_SETTINGS.copy()
        except Exception as e:
            print(f"Warning: Could not load settings file: {e}")
            return DEFAULT_SETTINGS.copy()

    def save_settings(self, settings=None):
        """Save settings to JSON file."""
        if settings is None:
            settings = self.settings

        try:
            self.config_path.parent.mkdir(parents=True, exist_ok=True)
            with open(self.config_path, 'w') as f:
                json.dump(settings, f, indent=2)
            return True
        except Exception as e:
            print(f"Error saving settings: {e}")
            return False

    def get(self, key, default=None):
        """Get a setting value."""
        return self.settings.get(key, default)

    def set(self, key, value):
        """Set a setting value and save."""
        self.settings[key] = value
        self.save_settings()

    def get_court_visitor_name(self):
        """Get the configured Court Visitor name."""
        return self.settings.get("court_visitor_name", "")

    def set_court_visitor_name(self, name):
        """Set the Court Visitor name."""
        self.set("court_visitor_name", name)

    def is_first_run(self):
        """Check if this is the first run of the application."""
        return not self.settings.get("first_run_complete", False)

    def mark_first_run_complete(self):
        """Mark that the first run is complete."""
        self.set("first_run_complete", True)

    def is_eula_accepted(self):
        """Check if EULA has been accepted."""
        return self.settings.get("eula_accepted", False)

    def accept_eula(self):
        """Mark EULA as accepted with timestamp."""
        self.settings["eula_accepted"] = True
        self.settings["eula_accepted_date"] = datetime.now().isoformat()
        self.save_settings()

    def get_license_key(self):
        """Get the license key."""
        return self.settings.get("license_key", "")

    def set_license_key(self, key):
        """Set the license key."""
        self.settings["license_key"] = key
        self.settings["activation_date"] = datetime.now().isoformat()
        self.save_settings()


def prompt_for_court_visitor_name(parent=None, current_name=""):
    """
    Show a dialog to get or update the Court Visitor name.

    Args:
        parent: Parent tkinter window (optional)
        current_name: Current name value (for editing)

    Returns:
        The entered name, or None if cancelled
    """
    # Create a custom dialog
    dialog = tk.Toplevel(parent) if parent else tk.Tk()
    dialog.title("Court Visitor Name")
    dialog.geometry("500x250")
    dialog.resizable(False, False)

    # Center the window
    dialog.update_idletasks()
    x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
    y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
    dialog.geometry(f"+{x}+{y}")

    result = {"name": None}

    # Header
    header_frame = tk.Frame(dialog, bg="#4a90e2", height=60)
    header_frame.pack(fill=tk.X)
    header_frame.pack_propagate(False)

    title_label = tk.Label(
        header_frame,
        text="⚖️ Court Visitor Name Setup",
        font=("Arial", 14, "bold"),
        bg="#4a90e2",
        fg="white"
    )
    title_label.pack(pady=15)

    # Body
    body_frame = tk.Frame(dialog, padx=30, pady=20)
    body_frame.pack(fill=tk.BOTH, expand=True)

    instruction_text = (
        "Please enter your name as it should appear on Court Visitor Reports.\n\n"
        "This name will be automatically filled in the CVR documents.\n"
        "You can change this later in the app settings."
    )

    instruction_label = tk.Label(
        body_frame,
        text=instruction_text,
        font=("Arial", 10),
        justify=tk.LEFT,
        wraplength=440
    )
    instruction_label.pack(pady=(0, 15))

    # Name entry
    entry_frame = tk.Frame(body_frame)
    entry_frame.pack(fill=tk.X, pady=10)

    label = tk.Label(entry_frame, text="Your Full Name:", font=("Arial", 10, "bold"))
    label.pack(anchor=tk.W)

    name_var = tk.StringVar(value=current_name)
    name_entry = tk.Entry(entry_frame, textvariable=name_var, font=("Arial", 11), width=40)
    name_entry.pack(fill=tk.X, pady=(5, 0))
    name_entry.focus()

    # Buttons
    button_frame = tk.Frame(body_frame)
    button_frame.pack(pady=(15, 0))

    def on_ok():
        name = name_var.get().strip()
        if name:
            result["name"] = name
            dialog.destroy()
        else:
            messagebox.showwarning(
                "Name Required",
                "Please enter your name to continue.",
                parent=dialog
            )

    def on_cancel():
        dialog.destroy()

    ok_button = tk.Button(
        button_frame,
        text="Save",
        command=on_ok,
        font=("Arial", 10, "bold"),
        bg="#4a90e2",
        fg="white",
        padx=30,
        pady=8,
        relief=tk.FLAT,
        cursor="hand2"
    )
    ok_button.pack(side=tk.LEFT, padx=5)

    cancel_button = tk.Button(
        button_frame,
        text="Cancel",
        command=on_cancel,
        font=("Arial", 10),
        padx=30,
        pady=8,
        relief=tk.FLAT,
        cursor="hand2"
    )
    cancel_button.pack(side=tk.LEFT, padx=5)

    # Bind Enter key to OK
    name_entry.bind('<Return>', lambda e: on_ok())

    # Make modal
    dialog.transient(parent)
    dialog.grab_set()

    # Wait for dialog to close
    if parent:
        dialog.wait_window()
    else:
        dialog.mainloop()

    return result["name"]


def ensure_court_visitor_name_set(config_manager, parent=None):
    """
    Ensure a Court Visitor name is set, prompting if needed.

    Args:
        config_manager: AppConfigManager instance
        parent: Parent tkinter window

    Returns:
        The Court Visitor name (guaranteed to be non-empty)
    """
    current_name = config_manager.get_court_visitor_name()

    if not current_name:
        name = prompt_for_court_visitor_name(parent, current_name)
        if name:
            config_manager.set_court_visitor_name(name)
            return name
        else:
            # User cancelled - ask again or use default
            messagebox.showwarning(
                "Name Required",
                "Court Visitor name is required. Using 'Court Visitor' as default.\n\n"
                "You can change this later in Settings.",
                parent=parent
            )
            default_name = "Court Visitor"
            config_manager.set_court_visitor_name(default_name)
            return default_name

    return current_name


# Standalone test
if __name__ == "__main__":
    config = AppConfigManager()
    print(f"Current Court Visitor Name: '{config.get_court_visitor_name()}'")

    # Test the prompt
    name = ensure_court_visitor_name_set(config)
    print(f"Court Visitor Name set to: '{name}'")

    # Show all settings
    print("\nAll settings:")
    for key, value in config.settings.items():
        print(f"  {key}: {value}")
