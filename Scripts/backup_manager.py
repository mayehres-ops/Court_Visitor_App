"""
User Data Backup Manager
Creates timestamped backups of all user data.
"""

import zipfile
import shutil
from pathlib import Path
from datetime import datetime
import tkinter as tk
from tkinter import messagebox, ttk
import threading


class BackupManager:
    """Manages user data backups."""

    def __init__(self, app_root):
        """
        Initialize backup manager.

        Args:
            app_root: Path to application root directory
        """
        self.app_root = Path(app_root)
        self.backup_dir = self.app_root / "App Data" / "Backups"
        self.backup_dir.mkdir(parents=True, exist_ok=True)

    def get_backup_items(self):
        """
        Get list of items to backup.

        Returns:
            List of (source_path, archive_path) tuples
        """
        items = []

        # Main data file
        excel_file = self.app_root / "App Data" / "ward_guardian_info.xlsx"
        if excel_file.exists():
            items.append((excel_file, "App Data/ward_guardian_info.xlsx"))

        # Config folder (all settings, API tokens)
        config_dir = self.app_root / "Config"
        if config_dir.exists():
            for file in config_dir.rglob("*"):
                if file.is_file():
                    rel_path = file.relative_to(self.app_root)
                    items.append((file, str(rel_path)))

        # Output folders
        output_dirs = [
            "App Data/Output/Mileage Logs",
            "App Data/Output/Payment Forms",
            "App Data/Output/CVR",
        ]
        for output_dir in output_dirs:
            output_path = self.app_root / output_dir
            if output_path.exists():
                for file in output_path.rglob("*"):
                    if file.is_file():
                        rel_path = file.relative_to(self.app_root)
                        items.append((file, str(rel_path)))

        # New Clients folder (case folders)
        new_clients = self.app_root / "New Clients"
        if new_clients.exists():
            for file in new_clients.rglob("*"):
                if file.is_file():
                    rel_path = file.relative_to(self.app_root)
                    items.append((file, str(rel_path)))

        return items

    def create_backup(self, progress_callback=None):
        """
        Create a timestamped backup ZIP file.

        Args:
            progress_callback: Optional callback function(current, total, message)

        Returns:
            Path to created backup file, or None if failed
        """
        try:
            # Generate backup filename
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_filename = f"CourtVisitorApp_Backup_{timestamp}.zip"
            backup_path = self.backup_dir / backup_filename

            # Get items to backup
            items = self.get_backup_items()

            if not items:
                return None

            # Create ZIP file
            with zipfile.ZipFile(backup_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for i, (source, archive_name) in enumerate(items, 1):
                    try:
                        zipf.write(source, archive_name)
                        if progress_callback:
                            progress_callback(i, len(items), f"Backing up: {archive_name}")
                    except Exception as e:
                        print(f"Warning: Could not backup {source}: {e}")

            return backup_path

        except Exception as e:
            print(f"Backup failed: {e}")
            return None

    def list_backups(self):
        """
        List all existing backups.

        Returns:
            List of (filename, filepath, size_mb, date) tuples, sorted by date (newest first)
        """
        backups = []
        for backup_file in self.backup_dir.glob("CourtVisitorApp_Backup_*.zip"):
            try:
                size_mb = backup_file.stat().st_size / (1024 * 1024)
                modified = datetime.fromtimestamp(backup_file.stat().st_mtime)
                backups.append((
                    backup_file.name,
                    backup_file,
                    size_mb,
                    modified
                ))
            except Exception:
                pass

        # Sort by date, newest first
        backups.sort(key=lambda x: x[3], reverse=True)
        return backups

    def delete_old_backups(self, keep_count=10):
        """
        Delete old backups, keeping only the most recent ones.

        Args:
            keep_count: Number of backups to keep
        """
        backups = self.list_backups()
        for filename, filepath, _, _ in backups[keep_count:]:
            try:
                filepath.unlink()
                print(f"Deleted old backup: {filename}")
            except Exception as e:
                print(f"Could not delete {filename}: {e}")


class BackupDialog:
    """Backup progress dialog."""

    def __init__(self, parent, backup_manager):
        """
        Initialize backup dialog.

        Args:
            parent: Parent window
            backup_manager: BackupManager instance
        """
        self.backup_manager = backup_manager
        self.backup_path = None

        # Create dialog
        self.root = tk.Toplevel(parent)
        self.root.title("Backup My Data")
        self.root.geometry("500x300")
        self.root.resizable(False, False)

        # Center window
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() // 2) - 250
        y = (self.root.winfo_screenheight() // 2) - 150
        self.root.geometry(f'+{x}+{y}')

        # Make modal
        self.root.transient(parent)
        self.root.grab_set()

        self._build_ui()

        # Start backup in background thread
        self.thread = threading.Thread(target=self._run_backup)
        self.thread.daemon = True
        self.thread.start()

    def _build_ui(self):
        """Build the user interface."""
        # Header
        header_frame = tk.Frame(self.root, bg='#667eea', height=80)
        header_frame.pack(fill='x')
        header_frame.pack_propagate(False)

        header_label = tk.Label(
            header_frame,
            text="💾 Backing Up Your Data",
            font=('Segoe UI', 14, 'bold'),
            bg='#667eea',
            fg='white'
        )
        header_label.pack(expand=True)

        # Content
        content_frame = tk.Frame(self.root, padx=30, pady=20)
        content_frame.pack(fill='both', expand=True)

        # Status label
        self.status_label = tk.Label(
            content_frame,
            text="Preparing backup...",
            font=('Segoe UI', 10),
            wraplength=440
        )
        self.status_label.pack(pady=(10, 20))

        # Progress bar
        self.progress = ttk.Progressbar(
            content_frame,
            length=440,
            mode='determinate'
        )
        self.progress.pack(pady=10)

        # Progress text
        self.progress_text = tk.Label(
            content_frame,
            text="0 / 0 files",
            font=('Segoe UI', 9),
            fg='#6b7280'
        )
        self.progress_text.pack()

        # Close button (disabled during backup)
        self.close_btn = tk.Button(
            content_frame,
            text="Close",
            command=self.root.destroy,
            font=('Segoe UI', 10),
            state='disabled',
            padx=30,
            pady=8
        )
        self.close_btn.pack(pady=(20, 0))

    def _update_progress(self, current, total, message):
        """Update progress UI."""
        def update():
            self.status_label.config(text=message)
            self.progress['maximum'] = total
            self.progress['value'] = current
            self.progress_text.config(text=f"{current} / {total} files")

        self.root.after(0, update)

    def _run_backup(self):
        """Run backup in background thread."""
        try:
            self.backup_path = self.backup_manager.create_backup(
                progress_callback=self._update_progress
            )

            if self.backup_path:
                # Clean up old backups
                self.backup_manager.delete_old_backups(keep_count=10)

                # Show success
                def show_success():
                    self.status_label.config(
                        text=f"✓ Backup completed successfully!\n\n"
                             f"Saved to:\n{self.backup_path.name}"
                    )
                    self.close_btn.config(state='normal', bg='#16a34a', fg='white')

                self.root.after(0, show_success)
            else:
                # Show error
                def show_error():
                    self.status_label.config(
                        text="⚠ Backup failed - No data found to backup"
                    )
                    self.close_btn.config(state='normal')

                self.root.after(0, show_error)

        except Exception as e:
            # Show error
            def show_error():
                self.status_label.config(
                    text=f"⚠ Backup failed:\n{str(e)}"
                )
                self.close_btn.config(state='normal')

            self.root.after(0, show_error)

    def show(self):
        """Show the dialog."""
        self.root.mainloop()
        return self.backup_path


def create_backup(parent, app_root):
    """
    Show backup dialog and create backup.

    Args:
        parent: Parent window
        app_root: Path to application root

    Returns:
        Path to backup file, or None if failed
    """
    manager = BackupManager(app_root)
    dialog = BackupDialog(parent, manager)
    return dialog.show()


if __name__ == "__main__":
    # Test
    import os
    os.chdir(r"C:\GoogleSync\GuardianShip_App")

    root = tk.Tk()
    root.withdraw()

    backup_path = create_backup(root, os.getcwd())
    if backup_path:
        print(f"Backup created: {backup_path}")
    else:
        print("Backup failed")
