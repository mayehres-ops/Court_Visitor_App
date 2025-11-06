"""
Desktop Shortcut Creator
Creates a desktop shortcut for the application on first run.
"""

import os
import sys
from pathlib import Path
import winshell
from win32com.client import Dispatch


def create_desktop_shortcut(app_name="Court Visitor App", exe_path=None):
    """
    Create a desktop shortcut to the application.

    Args:
        app_name: Name of the application (for shortcut name)
        exe_path: Path to the .exe file (if None, uses current script)

    Returns:
        True if successful, False otherwise
    """
    try:
        # Get desktop path
        desktop = winshell.desktop()

        # Determine exe path
        if exe_path is None:
            if getattr(sys, 'frozen', False):
                # Running as compiled .exe
                exe_path = sys.executable
            else:
                # Running as Python script (development mode)
                # Point to guardianship_app.py for now
                exe_path = Path(__file__).parent.parent / "guardianship_app.py"
                if not exe_path.exists():
                    print(f"Warning: Could not find guardianship_app.py at {exe_path}")
                    return False

        exe_path = Path(exe_path).resolve()

        # Create shortcut path
        shortcut_path = Path(desktop) / f"{app_name}.lnk"

        # Check if shortcut already exists
        if shortcut_path.exists():
            print(f"Desktop shortcut already exists: {shortcut_path}")
            return True

        # Create shortcut
        shell = Dispatch('WScript.Shell')
        shortcut = shell.CreateShortCut(str(shortcut_path))
        shortcut.Targetpath = str(exe_path)
        shortcut.WorkingDirectory = str(exe_path.parent)
        shortcut.IconLocation = str(exe_path)
        shortcut.Description = f"Launch {app_name}"
        shortcut.save()

        print(f"Desktop shortcut created: {shortcut_path}")
        return True

    except Exception as e:
        print(f"Error creating desktop shortcut: {e}")
        return False


def check_desktop_shortcut_exists(app_name="Court Visitor App"):
    """
    Check if desktop shortcut already exists.

    Args:
        app_name: Name of the application

    Returns:
        True if shortcut exists, False otherwise
    """
    try:
        desktop = winshell.desktop()
        shortcut_path = Path(desktop) / f"{app_name}.lnk"
        return shortcut_path.exists()
    except Exception:
        return False


def remove_desktop_shortcut(app_name="Court Visitor App"):
    """
    Remove desktop shortcut (for uninstall).

    Args:
        app_name: Name of the application

    Returns:
        True if successful, False otherwise
    """
    try:
        desktop = winshell.desktop()
        shortcut_path = Path(desktop) / f"{app_name}.lnk"

        if shortcut_path.exists():
            shortcut_path.unlink()
            print(f"Desktop shortcut removed: {shortcut_path}")
            return True
        else:
            print("Desktop shortcut does not exist")
            return False

    except Exception as e:
        print(f"Error removing desktop shortcut: {e}")
        return False


if __name__ == "__main__":
    # Test creating shortcut
    print("Testing desktop shortcut creation...")
    success = create_desktop_shortcut()

    if success:
        print("\n✓ Desktop shortcut created successfully!")
        print(f"Check your Desktop for 'Court Visitor App.lnk'")
    else:
        print("\n⚠ Failed to create desktop shortcut")
