#!/usr/bin/env python3
"""
Auto-updater for Court Visitor App
Checks GitHub for new releases and prompts user to update
"""

import requests
import json
import webbrowser
from pathlib import Path
from packaging import version
import tkinter as tk
from tkinter import messagebox

class AutoUpdater:
    def __init__(self, current_version, github_repo):
        """
        Args:
            current_version: Current version string (e.g., "1.0.0")
            github_repo: GitHub repo in format "username/repo"
        """
        self.current_version = current_version
        self.github_repo = github_repo
        self.api_url = f"https://api.github.com/repos/{github_repo}/releases/latest"

    def check_for_updates(self, silent=False):
        """
        Check for updates on GitHub

        Args:
            silent: If True, don't show "no updates" message

        Returns:
            tuple: (update_available, latest_version, download_url, release_notes)
        """
        try:
            response = requests.get(self.api_url, timeout=5)

            if response.status_code != 200:
                if not silent:
                    print(f"Could not check for updates (status {response.status_code})")
                return False, None, None, None

            data = response.json()
            latest_version = data['tag_name'].lstrip('v')

            # Compare versions
            if version.parse(latest_version) > version.parse(self.current_version):
                download_url = data.get('html_url')  # Link to release page
                release_notes = data.get('body', 'No release notes available')

                # Try to find Windows executable in assets
                for asset in data.get('assets', []):
                    if asset['name'].endswith('.exe'):
                        download_url = asset['browser_download_url']
                        break

                return True, latest_version, download_url, release_notes

            return False, latest_version, None, None

        except requests.exceptions.RequestException as e:
            if not silent:
                print(f"Network error checking for updates: {e}")
            return False, None, None, None
        except Exception as e:
            if not silent:
                print(f"Error checking for updates: {e}")
            return False, None, None, None

    def prompt_update(self, parent=None):
        """
        Check for updates and show dialog if available

        Args:
            parent: Tkinter parent window (optional)
        """
        update_available, latest_version, download_url, release_notes = self.check_for_updates(silent=True)

        if not update_available:
            return False

        # Create update dialog
        message = f"A new version is available!\n\n"
        message += f"Current version: {self.current_version}\n"
        message += f"Latest version: {latest_version}\n\n"
        message += f"Release notes:\n{release_notes[:200]}{'...' if len(release_notes) > 200 else ''}\n\n"
        message += "Would you like to download the update?"

        result = messagebox.askyesno(
            "Update Available",
            message,
            parent=parent
        )

        if result and download_url:
            webbrowser.open(download_url)
            return True

        return False

    def check_on_startup(self, parent=None, silent=True):
        """
        Check for updates on app startup

        Args:
            parent: Tkinter parent window
            silent: If True, only show message if update available
        """
        import threading

        def check():
            update_available, latest_version, download_url, release_notes = self.check_for_updates(silent=True)

            if update_available:
                # Schedule dialog on main thread
                if parent:
                    parent.after(100, lambda: self.prompt_update(parent))

        # Check in background thread to avoid blocking startup
        thread = threading.Thread(target=check, daemon=True)
        thread.start()


# Example usage:
if __name__ == "__main__":
    # Test the updater
    updater = AutoUpdater(
        current_version="1.0.0",
        github_repo="your-username/court-visitor-app"  # UPDATE THIS
    )

    update_available, latest_version, download_url, release_notes = updater.check_for_updates()

    if update_available:
        print(f"Update available: {latest_version}")
        print(f"Download: {download_url}")
        print(f"Notes: {release_notes}")
    else:
        print("No updates available")
