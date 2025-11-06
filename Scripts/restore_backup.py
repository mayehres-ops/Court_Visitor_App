"""
Court Visitor App - Backup Restoration System
Copyright (c) 2024 GuardianShip Easy, LLC. All rights reserved.

Safely restores application from a verified backup.

Usage:
    python Scripts/restore_backup.py --backup path/to/backup
    python Scripts/restore_backup.py --latest
    python Scripts/restore_backup.py --date 20241105_120000
"""

import os
import sys
import shutil
import json
import argparse
from pathlib import Path
from datetime import datetime

# Fix encoding for Windows console
if sys.platform == 'win32':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')


class RestoreSystem:
    def __init__(self, app_root=None):
        """
        Initialize restore system.

        Args:
            app_root: Root directory of app (auto-detected if not provided)
        """
        # Detect app root
        if app_root:
            self.app_root = Path(app_root)
        else:
            # This script is in Scripts/, so go up one level
            self.app_root = Path(__file__).parent.parent.resolve()

        # Default backup location
        self.backup_base = self.app_root.parent / "GuardianShip_App_Backups"

    def find_backup(self, backup_identifier=None, use_latest=False):
        """
        Find a backup directory.

        Args:
            backup_identifier: Path to backup, date string, or None
            use_latest: If True, use the most recent backup

        Returns:
            Path to backup directory or None
        """
        # If given full path
        if backup_identifier and Path(backup_identifier).exists():
            return Path(backup_identifier)

        # If looking for latest
        if use_latest or backup_identifier == "latest":
            backups = sorted(self.backup_base.glob("Backup_*"), reverse=True)
            if backups:
                return backups[0]
            else:
                print("❌ No backups found")
                return None

        # If given date string
        if backup_identifier:
            backup_dir = self.backup_base / f"Backup_{backup_identifier}"
            if backup_dir.exists():
                return backup_dir
            else:
                print(f"❌ Backup not found: {backup_dir}")
                return None

        return None

    def verify_backup(self, backup_dir):
        """
        Verify backup before restoring.

        Returns:
            tuple: (is_valid: bool, metadata: dict or None)
        """
        metadata_file = backup_dir / "BACKUP_METADATA.json"

        if not metadata_file.exists():
            print("❌ Backup metadata not found - may be corrupted")
            return False, None

        with open(metadata_file, 'r') as f:
            metadata = json.load(f)

        # Verify critical files exist in backup
        critical_files = metadata.get('critical_files', [])
        missing_files = []

        for critical_file in critical_files:
            backup_path = backup_dir / critical_file
            if not backup_path.exists():
                missing_files.append(critical_file)

        if missing_files:
            print(f"❌ Critical files missing from backup:")
            for f in missing_files:
                print(f"   - {f}")
            return False, None

        return True, metadata

    def restore_backup(self, backup_dir, dry_run=False, skip_confirmation=False):
        """
        Restore from backup.

        Args:
            backup_dir: Path to backup directory
            dry_run: If True, show what would be done without doing it
            skip_confirmation: If True, don't ask for confirmation

        Returns:
            bool: Success or failure
        """
        print("=" * 70)
        print("Court Visitor App - Restore from Backup")
        print("=" * 70)
        print(f"\nBackup: {backup_dir}")
        print(f"App Root: {self.app_root}")

        # Verify backup
        print("\n🔍 Verifying backup...")
        is_valid, metadata = self.verify_backup(backup_dir)

        if not is_valid:
            print("\n❌ Backup verification failed - restore aborted")
            return False

        print("   ✅ Backup verified")

        # Show backup info
        print(f"\n📋 Backup Information:")
        print(f"   Created: {metadata.get('datetime', 'Unknown')}")
        print(f"   Description: {metadata.get('description', 'N/A')}")
        print(f"   Files: {metadata.get('file_count', 'Unknown')}")

        # Confirm restore
        if not skip_confirmation and not dry_run:
            print("\n⚠️  WARNING: This will OVERWRITE your current application files!")
            print("\nAre you sure you want to restore from this backup?")
            response = input("Type 'yes' to continue: ")

            if response.lower() != 'yes':
                print("\n❌ Restore cancelled by user")
                return False

        # Create backup of current state before restoring
        if not dry_run:
            print("\n📦 Creating safety backup of current state...")
            try:
                from create_verified_backup import BackupSystem
                safety_backup = BackupSystem()
                success, _, _ = safety_backup.create_backup("Before restore - safety backup")
                if success:
                    print("   ✅ Safety backup created")
                else:
                    print("   ⚠️  Warning: Could not create safety backup")
            except Exception as e:
                print(f"   ⚠️  Warning: Could not create safety backup: {e}")

        # Restore files
        print("\n📋 Restoring files...")
        file_count = 0
        error_count = 0

        for item in backup_dir.rglob('*'):
            if item.is_file():
                # Skip metadata file
                if item.name == "BACKUP_METADATA.json":
                    continue

                # Skip restore script
                if item.name == "RESTORE_THIS_BACKUP.bat":
                    continue

                # Calculate relative path
                rel_path = item.relative_to(backup_dir)
                dest_path = self.app_root / rel_path

                if dry_run:
                    print(f"   [DRY RUN] Would restore: {rel_path}")
                    file_count += 1
                else:
                    try:
                        # Create parent directories
                        dest_path.parent.mkdir(parents=True, exist_ok=True)

                        # Copy file
                        shutil.copy2(item, dest_path)
                        file_count += 1

                        # Show progress every 10 files
                        if file_count % 10 == 0:
                            print(f"   Restored {file_count} files...", end='\r')

                    except Exception as e:
                        print(f"\n   ⚠️  Warning: Could not restore {rel_path}: {e}")
                        error_count += 1

        if dry_run:
            print(f"\n   [DRY RUN] Would restore {file_count} files")
        else:
            print(f"\n   ✅ Restored {file_count} files")
            if error_count > 0:
                print(f"   ⚠️  {error_count} files had errors")

        # Final summary
        print("\n" + "=" * 70)
        if dry_run:
            print("✅ DRY RUN COMPLETE - No files were actually changed")
        else:
            print("✅ RESTORE COMPLETE")
            print("\nYour application has been restored to:")
            print(f"   {metadata.get('datetime', 'the backup state')}")
            if metadata.get('description'):
                print(f"   {metadata['description']}")
        print("=" * 70)

        return True

    def list_backups(self):
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

                # Calculate size
                total_size = sum(f.stat().st_size for f in backup_dir.rglob('*') if f.is_file())
                size_mb = total_size / (1024 * 1024)

                print(f"\n📁 {backup_dir.name}")
                print(f"   Date: {metadata.get('datetime', 'Unknown')}")
                print(f"   Description: {metadata.get('description', 'N/A')}")
                print(f"   Files: {metadata.get('file_count', 'Unknown')}")
                print(f"   Size: {size_mb:.1f} MB")
                print(f"   To restore: python Scripts/restore_backup.py --backup \"{backup_dir}\"")

        if not backups:
            print("No backups found.")

        return backups


def main():
    parser = argparse.ArgumentParser(description='Restore Court Visitor App from backup')
    parser.add_argument('--backup', '-b', help='Path to backup directory')
    parser.add_argument('--latest', '-l', action='store_true', help='Restore from latest backup')
    parser.add_argument('--date', '-d', help='Restore from backup with specific date (e.g., 20241105_120000)')
    parser.add_argument('--list', action='store_true', help='List all available backups')
    parser.add_argument('--dry-run', action='store_true', help='Show what would be restored without doing it')
    parser.add_argument('--yes', '-y', action='store_true', help='Skip confirmation prompt')
    parser.add_argument('--app-root', help='App root directory (auto-detected if not provided)')

    args = parser.parse_args()

    # Create restore system
    restore_system = RestoreSystem(app_root=args.app_root)

    # List backups if requested
    if args.list:
        restore_system.list_backups()
        return

    # Find backup
    backup_dir = None

    if args.backup:
        backup_dir = restore_system.find_backup(args.backup)
    elif args.latest:
        backup_dir = restore_system.find_backup(use_latest=True)
    elif args.date:
        backup_dir = restore_system.find_backup(args.date)
    else:
        print("❌ Error: Must specify --backup, --latest, or --date")
        print("\nUse --list to see available backups")
        sys.exit(1)

    if not backup_dir:
        print("\n❌ No backup found to restore")
        sys.exit(1)

    # Restore backup
    success = restore_system.restore_backup(
        backup_dir,
        dry_run=args.dry_run,
        skip_confirmation=args.yes
    )

    if success:
        sys.exit(0)
    else:
        print("\n❌ Restore failed!")
        sys.exit(1)


if __name__ == "__main__":
    main()
