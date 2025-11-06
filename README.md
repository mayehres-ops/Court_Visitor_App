# Court Visitor App

**Professional workflow automation software for Court-Appointed Visitors in Travis County, Texas.**

![Version](https://img.shields.io/badge/version-1.0.0-blue)
![License](https://img.shields.io/badge/license-Proprietary-red)
![Platform](https://img.shields.io/badge/platform-Windows-lightgrey)

---

## Overview

Court Visitor App is a comprehensive automation tool designed specifically for Court-Appointed Visitors in Travis County, Texas. It streamlines the entire workflow from PDF intake to final reporting, automating 14 manual steps and integrating with Google services.

### Key Features

- **Complete Workflow Automation** - 14 automated steps from PDF to report submission
- **OCR Data Extraction** - Automatically extracts ward/guardian information from PDFs
- **Form Generation** - Creates Mileage, Payment, and CVR forms automatically
- **Google Integration** - Gmail, Calendar, Sheets, Drive, and Maps APIs
- **Configuration System** - Customizable Court Visitor settings and preferences
- **User Data Backup** - One-click backup of all user data with automatic cleanup
- **Legal Protection** - Built-in EULA, copyright notices, and confidentiality safeguards

---

## System Requirements

- **Operating System**: Windows 10 or Windows 11
- **Memory**: 4 GB RAM minimum (8 GB recommended)
- **Storage**: 500 MB free space
- **Internet**: Required for Google API services
- **Software**: Microsoft Office (Word and Excel)

---

## Installation

### For End Users

1. **Download** the latest release from the [Releases](https://github.com/mayehres-ops/Court_Visitor_App/releases) page
2. **Extract** the ZIP file to your desired location (e.g., `C:\CourtVisitorApp`)
3. **Run** `CourtVisitorApp.exe`
4. **Follow** the first-run setup wizard:
   - Accept EULA
   - Configure your Court Visitor information
   - Authorize Google API access

**Important:** Windows SmartScreen may show a warning (unsigned application). Click "More info" → "Run anyway". See [Installation Guide](Documentation/INSTALLATION_GUIDE.md) for details.

### For Developers

See [Development Setup](#development-setup) below.

---

## Documentation

- **[Installation Guide](Documentation/INSTALLATION_GUIDE.md)** - Complete installation instructions with security warning explanations
- **[PyInstaller Build Guide](Documentation/PYINSTALLER_BUILD_GUIDE.md)** - Building the executable
- **[GitHub Setup Guide](Documentation/GITHUB_SETUP_GUIDE.md)** - Distribution and version control
- **[Legal Protections](Documentation/LEGAL_PROTECTIONS_SUMMARY.md)** - Copyright, EULA, confidentiality
- **[CHANGELOG](CHANGELOG.md)** - Version history and release notes

---

## Quick Start

### First Launch

1. **Accept EULA** - Read and accept the End User License Agreement
2. **Configure Settings** - Enter your Court Visitor information:
   - Name
   - Vendor Number
   - GL Number
   - Cost Center
   - Address
3. **Authorize Google** - Sign in to Google and authorize the app to access:
   - Gmail (send emails)
   - Calendar (create events)
   - Drive (store documents)
   - Sheets (track data)
   - Maps (calculate mileage)

### Basic Workflow

1. **Drop PDF** in New Files folder
2. **Run Step 1** - Extract data from PDF (OCR)
3. **Run Steps 2-14** - Complete automation workflow
4. **Review Output** - Check generated forms and reports

---

## Features in Detail

### Automation Steps

1. **OCR Extraction** - Extract ward/guardian data from court-issued PDFs
2. **Folder Creation** - Create organized case folders
3. **Map Sheet** - Generate address tracking spreadsheet
4. **Email Request** - Send initial contact email to guardian
5. **Add Contacts** - Add guardian to Gmail contacts
6. **Confirmation Email** - Send visit confirmation
7. **Calendar Event** - Create visit appointment in Google Calendar
8. **CVR Builder** - Generate Court Visitor Report
9. **Court Visitor Summary** - Create summary document
10. **Google Sheets** - Update tracking spreadsheet
11. **Follow-up Email** - Send post-visit follow-up
12. **Mileage Form** - Generate mileage reimbursement form
13. **Payment Form** - Generate payment invoice
14. **Final Email** - Send completion email to supervisor

### Configuration

- **Court Visitor Info** - Personal information auto-fills all forms
- **Email Addresses** - Configurable supervisor and default emails
- **Mileage Addresses** - Home/office addresses for calculations
- **Google API** - Secure OAuth2 authentication

### Data Management

- **Backup System** - One-click backup to timestamped ZIP files
- **Auto-Cleanup** - Keeps 10 most recent backups automatically
- **Secure Storage** - All data stored locally, no cloud upload

---

## Security & Privacy

### Data Confidentiality

This application processes sensitive information including:
- Ward personal information (names, addresses, health data)
- Guardian contact information
- Court-appointed visitor personal information

**Important:** All data remains on your local machine or your authorized Google account. GuardianShip Easy, LLC does NOT collect, store, or have access to any personal data.

### HIPAA Compliance

Users must ensure compliance with:
- HIPAA privacy rules
- Texas state confidentiality requirements
- Court-ordered confidentiality provisions

See [EULA.txt](EULA.txt) for complete confidentiality requirements.

### Security Features

- ✓ Encrypted OAuth2 tokens
- ✓ Local data storage only
- ✓ No telemetry or usage tracking
- ✓ Secure Google API communication (HTTPS)

---

## Google API Setup

### Required APIs

This application uses the following Google APIs:
- Gmail API (sending emails)
- Google Calendar API (creating events)
- Google Sheets API (tracking data)
- Google Drive API (storing documents)
- Google Maps Directions API (mileage calculations)
- Google Cloud Vision API (OCR for PDF extraction)

### OAuth Consent Screen

The application requires:
- **Publishing status**: Testing mode
- **Test users**: Authorized Court Visitor email addresses must be added manually

See [Publish Google API Guide](Documentation/PUBLISH_GOOGLE_API.md) for setup instructions.

---

## License

**Proprietary and Confidential**

Copyright © 2024-2025 GuardianShip Easy, LLC. All rights reserved.

This software is licensed for use by authorized Court Visitors only. Unauthorized copying, distribution, modification, or commercial use is strictly prohibited.

**NOT FOR RESALE** - This software may not be sold, rented, leased, or otherwise transferred to third parties.

See [EULA.txt](EULA.txt) for complete license terms.

---

## Support

### Contact Information

- **Email**: support@guardianshipeasy.com
- **Website**: www.GuardianshipEasy.com

### Reporting Issues

If you encounter bugs or issues, contact support with:
- Description of the issue
- Steps to reproduce
- Screenshots (if applicable)
- Error messages

### Feature Requests

We welcome feedback and feature suggestions! Email us at support@guardianshipeasy.com.

---

## Development Setup

### Prerequisites

- Python 3.8 or higher
- Git
- Microsoft Office (Word and Excel)
- Google Cloud Console account

### Installation

```bash
# Clone the repository
git clone https://github.com/mayehres-ops/Court_Visitor_App.git
cd Court_Visitor_App

# Install dependencies
pip install -r requirements.txt

# Run the application
python guardianship_app.py
```

### Project Structure

```
Court_Visitor_App/
├── guardianship_app.py          # Main application entry point
├── Scripts/                     # Python modules
│   ├── app_paths.py            # Dynamic path management
│   ├── backup_manager.py       # Backup system
│   ├── cv_info_manager.py      # Configuration management
│   ├── cv_settings_dialog.py   # Settings GUI
│   ├── eula_dialog.py          # EULA acceptance dialog
│   ├── about_dialog.py         # About dialog
│   └── desktop_shortcut.py     # Desktop shortcut creation
├── Automation/                  # Automation scripts
│   ├── Mileage Reimbursement Script/
│   ├── CV Payment Form Script/
│   ├── Create CV report_move to folder/
│   └── ...
├── Templates/                   # Form templates
│   ├── Mileage_Reimbursement_Form.xlsx
│   ├── Court_Visitor_Payment_Invoice.docx
│   └── Court Visitor Report fillable new.docx
├── Documentation/               # User guides and developer docs
└── Config/                      # Configuration files (not in repo)
```

### Building the Executable

See [PyInstaller Build Guide](Documentation/PYINSTALLER_BUILD_GUIDE.md) for detailed instructions.

```bash
# Install PyInstaller
pip install pyinstaller

# Build executable
pyinstaller --name "CourtVisitorApp" --onefile --windowed --add-data "Templates;Templates" --add-data "Documentation;Documentation" --add-data "EULA.txt;." guardianship_app.py
```

---

## Roadmap

### Version 1.0.0 (Current)
- ✓ Complete 14-step workflow automation
- ✓ Court Visitor configuration system
- ✓ User data backup
- ✓ Legal protections (EULA, copyright)
- ✓ Desktop shortcut creation

### Version 1.1 (Planned)
- [ ] Digital code signing (remove security warnings)
- [ ] Automatic update checker
- [ ] Custom app icon
- [ ] Unified settings panel
- [ ] Backup restore functionality
- [ ] "What's New" dialog

### Future Considerations
- [ ] Mac version
- [ ] Automatic PDF monitoring
- [ ] Enhanced error reporting
- [ ] Multi-language support
- [ ] Custom PDF drop zone widget

See [CHANGELOG.md](CHANGELOG.md) for complete version history.

---

## Contributing

This is a proprietary, closed-source application. Contributions are limited to authorized developers only.

---

## Acknowledgments

### Technologies Used

- **Python** - Core programming language
- **Tkinter** - GUI framework
- **Google Cloud APIs** - Gmail, Calendar, Sheets, Drive, Maps, Vision
- **OpenPyXL** - Excel file manipulation
- **python-docx** - Word document manipulation
- **PyInstaller** - Executable building
- **win32com** - Office automation

### Libraries

- openpyxl
- python-docx
- google-api-python-client
- google-auth
- Pillow (PIL)
- winshell
- pywin32

---

## Disclaimer

This software is provided "AS IS" without warranty of any kind. See [EULA.txt](EULA.txt) for complete warranty disclaimer and limitation of liability.

Users are solely responsible for:
- Data security and privacy compliance
- HIPAA compliance
- Accurate data entry and verification
- Following court procedures and requirements

---

## Version

**Current Version:** 1.0.0
**Release Date:** November 6, 2024
**Last Updated:** November 6, 2024

---

**Copyright © 2024-2025 GuardianShip Easy, LLC. All rights reserved.**
