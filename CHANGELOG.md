# Changelog

All notable changes to Court Visitor App will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.0.0] - 2024-11-06

### Initial Release

#### Added
- **Complete Workflow Automation** for Court Visitors
  - 14 automated steps from PDF intake to final reporting
  - OCR for PDF data extraction
  - Automated form generation (Mileage, Payment, CVR)
  - Google API integration (Gmail, Calendar, Sheets, Drive, Maps)

- **Court Visitor Configuration System**
  - Settings dialog for personal information
  - Auto-fill CV name, vendor number, GL #, cost center, address
  - First-run setup wizard
  - Configuration persists across sessions

- **User Data Backup**
  - One-click backup button
  - Creates timestamped ZIP files
  - Includes Excel data, Config, Output folders, New Clients
  - Automatic cleanup (keeps 10 most recent backups)

- **Desktop Shortcut**
  - Automatic desktop shortcut creation on first run
  - Quick access to launch the application
  - Removes need to navigate to installation folder

- **Legal Protection**
  - EULA with NOT FOR RESALE clause
  - Data confidentiality/HIPAA compliance notice
  - EULA acceptance dialog on first launch
  - Copyright protection throughout

- **User Interface**
  - Modern, clean design
  - Quick Access sidebar with utility buttons
  - Help menu (About, License, Documentation)
  - Status bar with real-time feedback
  - Integrated chatbot helper

- **Documentation**
  - Installation guide with security warning education
  - User manual
  - Legal protections summary
  - API setup guides
  - Troubleshooting documentation

- **Dynamic Path Detection**
  - Works from any installation location
  - No hardcoded paths
  - Auto-detects all required directories

#### Technical Details
- Python 3.x
- Tkinter GUI
- Google Cloud APIs (Gmail, Calendar, Sheets, Drive, Maps, Vision)
- Win32 COM automation for Office
- OCR with Google Document AI
- Automated mileage calculations with Google Maps Directions API

---

## [Unreleased]

### Planned for v1.1
- Digital code signing (remove security warnings)
- Unified settings panel
- Automatic update checker
- Custom app icon (.ico file)
- Improved backup restore functionality
- "What's New" dialog on updates

### Under Consideration
- Mac version
- Automatic PDF monitoring for New Files folder
- Custom PDF drop zone widget
- Enhanced error reporting
- Usage analytics (opt-in)
- Multi-language support

---

## Version History

### Version Numbering
- **Major.Minor.Patch** (e.g., 1.2.3)
- **Major**: Breaking changes, major new features
- **Minor**: New features, backwards compatible
- **Patch**: Bug fixes, minor improvements

### Release Schedule
- Major releases: As needed (significant features/changes)
- Minor releases: Monthly (new features, enhancements)
- Patch releases: Weekly or as needed (bug fixes)

---

## How to Use This Changelog

**For Users:**
- Check this file to see what's new in each version
- Look for [Added], [Changed], [Fixed] sections
- Review [Breaking Changes] carefully before updating

**For Developers:**
- Update this file with EVERY release
- Use clear, user-friendly language
- Link to issues/PRs when relevant
- Date all entries

---

## Categories

### Added
New features or functionality

### Changed
Changes to existing functionality

### Deprecated
Features that will be removed in upcoming releases

### Removed
Features that have been removed

### Fixed
Bug fixes

### Security
Security fixes or improvements

---

**Current Version:** 1.0.0
**Last Updated:** November 6, 2024
**Next Release:** TBD
