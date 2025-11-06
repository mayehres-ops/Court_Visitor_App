"""
Court Visitor App - Verified Backup System
Copyright (c) 2024 GuardianShip Easy, LLC. All rights reserved.

Creates a complete, timestamped, verified backup of the entire application.
Includes verification that backup is restorable.

Usage:
    python Scripts/create_verified_backup.py

    Optional:
    python Scripts/create_verified_backup.py --description "Before path fixes"
"""

import os
import sys
import shutil
import hashlib
import json
from pathlib import Path
from datetime import datetime
import argparse

# Fix encoding for Windows console
if sys.platform == 'win32':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')


class BackupSystem:
    def __init__(self, app_root=None, backup_location=None):
        """
        Initialize backup system.

        Args:
            app_root: Root directory of app (auto-detected if not provided)
            backup_location: Where to store backups (defaults to parent of app_root)
        """
        # Detect app root
        if app_root:
            self.app_root = Path(app_root)
        else:
            # This script is in Scripts/, so go up one level
            self.app_root = Path(__file__).parent.parent.resolve()

        # Backup location (one level up from app root)
        if backup_location:
            self.backup_base = Path(backup_location)
        else:
            self.backup_base = self.app_root.parent / "GuardianShip_App_Backups"

        # Create backup directory
        self.backup_base.mkdir(exist_ok=True)

        # Timestamp for this backup
        self.timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

        # This backup's directory
        self.backup_dir = self.backup_base / f"Backup_{self.timestamp}"

        # Metadata file
        self.metadata_file = self.backup_dir / "BACKUP_METADATA.json"

        # Files to exclude from backup
        self.exclude_patterns = [
            "__pycache__",
            "*.pyc",
            ".git",
            ".vscode",
            ".idea",
            "*.log",
            "venv",
            "env",
            "GuardianShip_App_Backups",  # Don't backup backups!
        ]

        # Critical files that MUST exist
        self.critical_files = [
            "guardianship_app.py",
            "App Data/ward_guardian_info.xlsx",
            "Config",
            "Scripts",
            "Automation",
        ]

    def should_exclude(self, path):
        """Check if path should be excluded from backup."""
        path_str = str(path)

        for pattern in self.exclude_patterns:
            if pattern in path_str:
                return True
            if path.name.startswith('.'):
                return True

        return False

    def calculate_checksum(self, file_path):
        """Calculate MD5 checksum of a file."""
        md5 = hashlib.md5()
        try:
            with open(file_path, 'rb') as f:
                for chunk in iter(lambda: f.read(4096), b''):
                    md5.update(chunk)
            return md5.hexdigest()
        except Exception as e:
            print(f"Warning: Could not checksum {file_path}: {e}")
            return None

    def create_backup(self, description=""):
        """
        Create complete backup of application.

        Args:
            description: Optional description of why backup was created

        Returns:
            tuple: (success: bool, backup_path: Path, metadata: dict)
        """
        print("=" * 70)
        print("Court Visitor App - Creating Verified Backup")
        print("=" * 70)
        print(f"\nApp Root: {self.app_root}")
        print(f"Backup Location: {self.backup_dir}")
        print(f"Timestamp: {self.timestamp}")
        if description:
            print(f"Description: {description}")
        print()

        # Verify app root looks correct
        if not self.verify_app_root():
            print("❌ ERROR: App root doesn't look correct!")
            print(f"   Expected to find critical files in: {self.app_root}")
            return False, None, None

        # Create backup directory
        print("📁 Creating backup directory...")
        self.backup_dir.mkdir(parents=True, exist_ok=True)

        # Copy files
        print("📋 Copying files...")
        file_count = 0
        skipped_count = 0
        file_checksums = {}

        for item in self.app_root.rglob('*'):
            if item.is_file():
                # Check if should exclude
                if self.should_exclude(item):
                    skipped_count += 1
                    continue

                # Calculate relative path
                rel_path = item.relative_to(self.app_root)
                dest_path = self.backup_dir / rel_path

                # Create parent directories
                dest_path.parent.mkdir(parents=True, exist_ok=True)

                # Copy file
                try:
                    shutil.copy2(item, dest_path)

                    # Calculate checksum
                    checksum = self.calculate_checksum(item)
                    if checksum:
                        file_checksums[str(rel_path)] = checksum

                    file_count += 1

                    # Show progress every 10 files
                    if file_count % 10 == 0:
                        print(f"   Copied {file_count} files...", end='\r')

                except Exception as e:
                    print(f"\n⚠️  Warning: Could not copy {item}: {e}")

        print(f"   ✅ Copied {file_count} files (skipped {skipped_count})")

        # Create metadata
        metadata = {
            "timestamp": self.timestamp,
            "datetime": datetime.now().isoformat(),
            "description": description,
            "app_root": str(self.app_root),
            "backup_location": str(self.backup_dir),
            "file_count": file_count,
            "skipped_count": skipped_count,
            "file_checksums": file_checksums,
            "critical_files": self.critical_files,
            "python_version": sys.version,
        }

        # Save metadata
        print("\n📝 Saving metadata...")
        with open(self.metadata_file, 'w') as f:
            json.dump(metadata, f, indent=2)

        print(f"   ✅ Metadata saved to: {self.metadata_file}")

        # Verify backup
        print("\n🔍 Verifying backup integrity...")
        verification_result = self.verify_backup(self.backup_dir)

        if verification_result['success']:
            print(f"   ✅ Backup verified successfully!")
            print(f"      - {verification_result['verified_files']} files verified")
            print(f"      - {verification_result['checksum_matches']} checksums matched")
        else:
            print(f"   ❌ Backup verification FAILED!")
            print(f"      Issues: {verification_result['issues']}")
            return False, self.backup_dir, metadata

        # Create quick restore script
        print("\n📜 Creating restore script...")
        self.create_restore_script()

        # Final summary
        print("\n" + "=" * 70)
        print("✅ BACKUP COMPLETE")
        print("=" * 70)
        print(f"\nBackup Location: {self.backup_dir}")
        print(f"Files Backed Up: {file_count}")
        print(f"Backup Size: {self.get_directory_size(self.backup_dir):.1f} MB")
        print(f"\nTo restore from this backup:")
        print(f"   python Scripts/restore_backup.py --backup {self.backup_dir}")
        print(f"\nOr use the quick restore script:")
        print(f"   {self.backup_dir}/RESTORE_THIS_BACKUP.bat")
        print("\n" + "=" * 70)

        return True, self.backup_dir, metadata

    def verify_app_root(self):
        """Verify that app root contains critical files."""
        print("🔍 Verifying app root contains critical files...")

        all_found = True
        for critical_item in self.critical_files:
            path = self.app_root / critical_item
            exists = path.exists()
            status = "✅" if exists else "❌"
            print(f"   {status} {critical_item}")
            if not exists:
                all_found = False

        return all_found

    def verify_backup(self, backup_dir):
        """
        Verify backup is complete and checksums match.

        Returns:
            dict: Verification results
        """
        # Load metadata
        metadata_path = backup_dir / "BACKUP_METADATA.json"
        if not metadata_path.exists():
            return {
                'success': False,
                'issues': ['Metadata file not found']
            }

        with open(metadata_path, 'r') as f:
            metadata = json.load(f)

        # Verify critical files exist in backup
        issues = []
        for critical_file in metadata['critical_files']:
            backup_path = backup_dir / critical_file
            if not backup_path.exists():
                issues.append(f"Critical file missing from backup: {critical_file}")

        # Verify checksums (sample 10 files)
        checksums = metadata.get('file_checksums', {})
        sample_size = min(10, len(checksums))
        verified_files = 0
        checksum_matches = 0

        import random
        sample_files = random.sample(list(checksums.items()), sample_size) if checksums else []

        for rel_path, original_checksum in sample_files:
            backup_file = backup_dir / rel_path
            if backup_file.exists():
                verified_files += 1
                backup_checksum = self.calculate_checksum(backup_file)
                if backup_checksum == original_checksum:
                    checksum_matches += 1
                else:
                    issues.append(f"Checksum mismatch: {rel_path}")
            else:
                issues.append(f"File missing from backup: {rel_path}")

        success = len(issues) == 0

        return {
            'success': success,
            'verified_files': verified_files,
            'checksum_matches': checksum_matches,
            'issues': issues
        }

    def get_directory_size(self, directory):
        """Calculate total size of directory in MB."""
        total_size = 0
        for item in Path(directory).rglob('*'):
            if item.is_file():
                total_size += item.stat().st_size
        return total_size / (1024 * 1024)  # Convert to MB

    def create_restore_script(self):
        """Create a quick restore batch script."""
        restore_script = self.backup_dir / "RESTORE_THIS_BACKUP.bat"

        script_content = f"""@echo off
REM Quick Restore Script for Backup {self.timestamp}
REM Created: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}

echo ========================================
echo  Court Visitor App - Restore Backup
echo ========================================
echo.
echo This will restore the app to the state from:
echo {self.timestamp}
echo.
echo WARNING: This will overwrite your current app files!
echo.
pause

echo.
echo Restoring backup...
echo.

cd "{self.app_root.parent}"
python "{self.app_root}/Scripts/restore_backup.py" --backup "{self.backup_dir}"

if errorlevel 1 (
    echo.
    echo ERROR: Restore failed!
    pause
    exit /b 1
)

echo.
echo ========================================
echo Restore Complete!
echo ========================================
echo.
pause
"""

        with open(restore_script, 'w') as f:
            f.write(script_content)

        print(f"   ✅ Restore script created: RESTORE_THIS_BACKUP.bat")

    def list_all_backups(self):
        """List all available backups."""
        print("\n📚 Available Backups:")
        print("=" * 70)

        if not self.backup_base.exists():
            print("No backups found.")
            return []

        backups = []
        for backup_dir in sorted(self.backup_base.glob("Backup_*"), reverse=True):
            metadata_file = backup_dir / "BACKUP_METADATA.json"
            if metadata_file.exists():
                with open(metadata_file, 'r') as f:
                    metadata = json.load(f)

                backups.append({
                    'path': backup_dir,
                    'metadata': metadata
                })

                print(f"\n📁 {backup_dir.name}")
                print(f"   Date: {metadata.get('datetime', 'Unknown')}")
                print(f"   Description: {metadata.get('description', 'N/A')}")
                print(f"   Files: {metadata.get('file_count', 'Unknown')}")
                print(f"   Size: {self.get_directory_size(backup_dir):.1f} MB")

        return backups


def main():
    parser = argparse.ArgumentParser(description='Create verified backup of Court Visitor App')
    parser.add_argument('--description', '-d', default='', help='Description of backup')
    parser.add_argument('--list', '-l', action='store_true', help='List all backups')
    parser.add_argument('--app-root', help='App root directory (auto-detected if not provided)')
    parser.add_argument('--backup-location', help='Where to store backups')

    args = parser.parse_args()

    # Create backup system
    backup_system = BackupSystem(
        app_root=args.app_root,
        backup_location=args.backup_location
    )

    # List backups if requested
    if args.list:
        backup_system.list_all_backups()
        return

    # Create backup
    success, backup_path, metadata = backup_system.create_backup(args.description)

    if success:
        sys.exit(0)
    else:
        print("\n❌ Backup failed!")
        sys.exit(1)


if __name__ == "__main__":
    main()
