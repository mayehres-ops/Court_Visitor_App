"""
Quick backup script for GuardianShip App
"""
import shutil
import os
from datetime import datetime
from pathlib import Path

def create_backup():
    source = Path("C:/GoogleSync/GuardianShip_App")
    backup_dir = Path("C:/GoogleSync/Backup")
    backup_dir.mkdir(exist_ok=True)

    backup_name = f"GuardianShip_App_CV_Config_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    dest = backup_dir / backup_name

    print(f"Creating backup: {dest}")

    def ignore_files(dir, files):
        """Ignore problematic files"""
        ignore_list = []
        for f in files:
            if f.lower() in ['nul', 'con', 'prn', 'aux', 'com1', 'lpt1']:
                ignore_list.append(f)
            # Skip pycache and other temp files
            if f in ['__pycache__', '.pyc', '.pyo']:
                ignore_list.append(f)
        return ignore_list

    shutil.copytree(source, dest, ignore=ignore_files)
    print(f"Backup created successfully at: {dest}")
    return dest

if __name__ == "__main__":
    create_backup()
