"""
Court Visitor Information Manager
Manages Court Visitor personal information used across all forms.

Usage:
    from cv_info_manager import CVInfoManager

    cv_info = CVInfoManager()
    info = cv_info.get_info()
    print(info['name'])  # Court Visitor Name
"""

import json
import os
from pathlib import Path
from typing import Dict, Optional


class CVInfoManager:
    """Manages Court Visitor information configuration."""

    # Default field structure
    DEFAULT_INFO = {
        'name': '',
        'vendor_number': '',
        'gl_number': '',
        'cost_center': '',
        'address_line1': '',
        'address_line2': ''
    }

    def __init__(self, config_dir: Optional[str] = None):
        """
        Initialize CV Info Manager.

        Args:
            config_dir: Path to config directory. If None, auto-detects from app_paths.
        """
        if config_dir:
            self.config_dir = Path(config_dir)
        else:
            # Auto-detect using app_paths
            try:
                from app_paths import get_app_paths
                app_paths = get_app_paths()
                self.config_dir = app_paths.CONFIG_DIR
            except Exception:
                # Fallback to relative path
                self.config_dir = Path(__file__).parent.parent / "Config"

        self.config_file = self.config_dir / "court_visitor_info.json"

        # Ensure config directory exists
        self.config_dir.mkdir(parents=True, exist_ok=True)

    def get_info(self) -> Dict[str, str]:
        """
        Get Court Visitor information.

        Returns:
            Dictionary with CV info fields. Returns empty strings if not configured.
        """
        if not self.config_file.exists():
            return self.DEFAULT_INFO.copy()

        try:
            with open(self.config_file, 'r', encoding='utf-8') as f:
                info = json.load(f)

            # Ensure all fields exist
            result = self.DEFAULT_INFO.copy()
            result.update(info)
            return result

        except Exception as e:
            print(f"Warning: Could not read CV info: {e}")
            return self.DEFAULT_INFO.copy()

    def save_info(self, info: Dict[str, str]) -> bool:
        """
        Save Court Visitor information.

        Args:
            info: Dictionary with CV info fields

        Returns:
            True if successful, False otherwise
        """
        try:
            # Validate required fields
            if not info.get('name'):
                raise ValueError("Court Visitor Name is required")

            # Save to file
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(info, f, indent=2)

            return True

        except Exception as e:
            print(f"Error saving CV info: {e}")
            return False

    def is_configured(self) -> bool:
        """
        Check if Court Visitor information has been configured.

        Returns:
            True if configured with at least a name, False otherwise
        """
        info = self.get_info()
        return bool(info.get('name'))

    def clear_info(self) -> bool:
        """
        Clear Court Visitor information.

        Returns:
            True if successful, False otherwise
        """
        try:
            if self.config_file.exists():
                self.config_file.unlink()
            return True
        except Exception as e:
            print(f"Error clearing CV info: {e}")
            return False


def get_cv_info(config_dir: Optional[str] = None) -> Dict[str, str]:
    """
    Convenience function to get Court Visitor information.

    Args:
        config_dir: Optional config directory path

    Returns:
        Dictionary with CV info
    """
    manager = CVInfoManager(config_dir)
    return manager.get_info()


if __name__ == "__main__":
    # Test the manager
    manager = CVInfoManager()

    print("Current CV Info:")
    info = manager.get_info()
    for key, value in info.items():
        print(f"  {key}: {value or '(not set)'}")

    print(f"\nConfigured: {manager.is_configured()}")
    print(f"Config file: {manager.config_file}")
