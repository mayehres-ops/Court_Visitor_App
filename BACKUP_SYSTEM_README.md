# Backup & Restore System - Quick Guide

**Created:** November 5, 2024
**Purpose:** Safe backup before making path fixes

---

## Quick Start

### Step 1: Create Your First Backup (DO THIS NOW!)

```bash
cd C:\GoogleSync\GuardianShip_App
python Scripts/create_verified_backup.py --description "Before any path fixes - working state"
```

**What this does:**
- ✅ Copies entire app to backup location
- ✅ Verifies all critical files are backed up
- ✅ Calculates checksums to ensure integrity
- ✅ Creates quick restore script
- ✅ Shows backup summary

**Where backups are stored:**
- Default: `C:\GoogleSync\GuardianShip_App_Backups\`
- Each backup in: `Backup_YYYYMMDD_HHMMSS\` folder

**How long it takes:**
- ~1-2 minutes for full backup

---

### Step 2: Verify Backup Works

```bash
# List all backups
python Scripts/restore_backup.py --list

# Test restore (dry run - doesn't actually change anything)
python Scripts/restore_backup.py --latest --dry-run
```

**What to check:**
- ✅ Backup appears in list
- ✅ File count looks correct (~100-200 files)
- ✅ Size is reasonable (50-200 MB)
- ✅ Dry run shows it would restore files

---

### Step 3: If You Ever Need to Restore

```bash
# Restore from latest backup
python Scripts/restore_backup.py --latest

# Or restore from specific backup
python Scripts/restore_backup.py --backup "C:\GoogleSync\GuardianShip_App_Backups\Backup_20241105_120000"

# Or use the quick restore script in backup folder
C:\GoogleSync\GuardianShip_App_Backups\Backup_20241105_120000\RESTORE_THIS_BACKUP.bat
```

**What happens:**
1. Shows backup info
2. Asks for confirmation
3. Creates safety backup of current state
4. Restores all files from backup
5. Done!

---

## When to Create Backups

### Before Making Changes:
```bash
# Before fixing paths
python Scripts/create_verified_backup.py --description "Before fixing guardianship_app.py paths"

# Before updating CVR script
python Scripts/create_verified_backup.py --description "Before integrating Court Visitor name"

# Before any risky change
python Scripts/create_verified_backup.py --description "Before [what you're about to do]"
```

### After Successful Changes:
```bash
# After paths work
python Scripts/create_verified_backup.py --description "After fixing paths - all tests passed"

# After CVR name integration
python Scripts/create_verified_backup.py --description "After CV name integration - working"
```

---

## Backup Strategy for Path Fixes

### Daily Backup Schedule:

**Before starting work:**
```bash
python Scripts/create_verified_backup.py --description "Morning backup - $(date)"
```

**After each successful file fix:**
```bash
python Scripts/create_verified_backup.py --description "After fixing [filename] - tested OK"
```

**Before attempting risky file:**
```bash
python Scripts/create_verified_backup.py --description "Before fixing [risky_filename] - last known good"
```

### Example Workflow:

```bash
# Day 1: Start work
python Scripts/create_verified_backup.py --description "Day 1 start - before any changes"

# Fix first utility script
# ... make changes ...
# ... test ...
python Scripts/create_verified_backup.py --description "After fixing check_extraction_results.py - OK"

# Fix second utility script
# ... make changes ...
# ... test ...
python Scripts/create_verified_backup.py --description "After fixing clear_sheet1_rows.py - OK"

# End of day
python Scripts/create_verified_backup.py --description "End of Day 1 - 2 files fixed, all working"
```

---

## Restore Scenarios

### Scenario 1: Single File Broke
**Problem:** You fixed a file and it doesn't work

**Solution:** Just restore that one file manually
```bash
# Copy from latest backup
copy "C:\GoogleSync\GuardianShip_App_Backups\Backup_20241105_120000\[filename]" "C:\GoogleSync\GuardianShip_App\[filename]"
```

### Scenario 2: Multiple Files Broken
**Problem:** Several changes and now things are broken

**Solution:** Restore from last known good backup
```bash
# Find last good backup
python Scripts/restore_backup.py --list

# Restore it
python Scripts/restore_backup.py --backup "[path to good backup]"
```

### Scenario 3: Complete Disaster
**Problem:** Everything is broken, can't even run Python

**Solution:** Use quick restore script
```
1. Open File Explorer
2. Navigate to: C:\GoogleSync\GuardianShip_App_Backups\
3. Find latest Backup_YYYYMMDD_HHMMSS folder
4. Double-click: RESTORE_THIS_BACKUP.bat
5. Confirm restoration
```

---

## Backup Files Explained

### Inside Each Backup Folder:

```
Backup_20241105_143000/
├── guardianship_app.py          ← Your app files
├── Scripts/                      ← All scripts
├── Automation/                   ← All automation
├── App Data/                     ← Database and data
├── Config/                       ← Configuration
├── BACKUP_METADATA.json          ← Backup information
└── RESTORE_THIS_BACKUP.bat       ← Quick restore script
```

### BACKUP_METADATA.json Contains:
- Timestamp of backup
- Description you provided
- File count
- Checksums of all files
- Python version
- What was backed up

---

## Safety Features

### What Gets Backed Up:
- ✅ All Python scripts (.py files)
- ✅ Excel database
- ✅ Configuration files
- ✅ Templates
- ✅ Documentation

### What Gets Skipped:
- ❌ `__pycache__` folders
- ❌ `.pyc` files
- ❌ `.git` folders
- ❌ `venv` folders
- ❌ Log files
- ❌ Previous backups (doesn't backup backups!)

### Verification:
- ✅ Checks critical files exist
- ✅ Calculates checksums
- ✅ Verifies backup integrity
- ✅ Warns if anything missing

### Safety Backup Before Restore:
- ✅ Before restoring, creates backup of current state
- ✅ So you can undo the restore if needed
- ✅ Double safety!

---

## Troubleshooting

### Backup Failed - Critical Files Missing
**Error:** "Critical files missing from backup"

**Solution:**
- Check you're in correct directory
- Verify `guardianship_app.py` exists
- Make sure you're not running from wrong location

### Restore Failed - Files In Use
**Error:** "Could not restore [file]: Permission denied"

**Solution:**
1. Close the Court Visitor App
2. Close any Python processes
3. Close any Excel files
4. Try restore again

### Backup Taking Too Long
**Normal:** 1-2 minutes for full app
**If longer:**
- Check if antivirus is scanning
- Check if backup location has enough space
- Check if backing up to slow drive (network, external)

### Can't Find Backups
**Check:**
```bash
# Default location
dir C:\GoogleSync\GuardianShip_App_Backups

# List backups via script
python Scripts/restore_backup.py --list
```

---

## Advanced Usage

### Backup to Custom Location
```bash
python Scripts/create_verified_backup.py --backup-location "D:\MyBackups"
```

### Restore Specific Date
```bash
python Scripts/restore_backup.py --date 20241105_120000
```

### Dry Run (Test Without Actually Restoring)
```bash
python Scripts/restore_backup.py --latest --dry-run
```

### Skip Confirmation (Auto-Restore)
```bash
python Scripts/restore_backup.py --latest --yes
```

---

## Best Practices

### 1. Always Backup Before Changes
**NEVER** make risky changes without a recent backup

### 2. Verify Backup After Creation
Run `--list` to check backup was created

### 3. Test Restore Occasionally
Do a dry-run restore to ensure system works

### 4. Keep Multiple Backups
Don't delete old backups immediately
Keep at least 3-5 recent backups

### 5. Name Your Backups Descriptively
Use `--description` to remember what state it represents

### 6. Backup Before AND After
- Before: Safety net
- After: Known good state

---

## Your First Backup Checklist

- [ ] Navigate to app directory
- [ ] Run: `python Scripts/create_verified_backup.py --description "Before path fixes - working state"`
- [ ] Wait for completion (1-2 minutes)
- [ ] Verify backup succeeded (✅ BACKUP COMPLETE)
- [ ] Note backup location shown in output
- [ ] Test list: `python Scripts/restore_backup.py --list`
- [ ] Test dry-run: `python Scripts/restore_backup.py --latest --dry-run`
- [ ] Confirm you see backup in list
- [ ] Save backup location for future reference

**Once checklist complete:** You're safe to start making changes!

---

## Emergency Contacts

If something goes catastrophically wrong:

1. **Stop making changes**
2. **Don't panic**
3. **Check backup location exists**
4. **Use RESTORE_THIS_BACKUP.bat in most recent backup**
5. **Restore will bring you back to working state**

Remember: The backup system creates a safety backup BEFORE restoring,
so even if restore fails, you can't lose data!

---

**Ready? Create your first backup now before proceeding with any changes!**

```bash
cd C:\GoogleSync\GuardianShip_App
python Scripts/create_verified_backup.py --description "Before any path fixes - working state"
```
