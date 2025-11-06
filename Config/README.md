# GuardianShip App - Configuration Files

This folder contains all API keys and credentials needed for the app to work.

## Folder Structure

```
Config\
├── API\
│   ├── google_service_account.json     # Google Vision API credentials (Step 1)
│   ├── credentials.json                 # Google OAuth credentials (Steps 4-7, 10-12)
│   └── token_*.json                     # Auto-generated OAuth tokens
│
└── Keys\
    └── google_maps_api_key.txt          # Google Maps API key (Step 3)
```

## What You Need to Set Up

### 1. Google Maps API Key (CRITICAL for Step 3 - Route Maps)

**File:** `Config\Keys\google_maps_api_key.txt`

**What it's for:** Generates maps showing where ward addresses are located

**How to get it:**
1. Go to https://console.cloud.google.com/
2. Create a project
3. Enable: Maps Static API, Geocoding API, Directions API
4. Create API key
5. Paste the key into `google_maps_api_key.txt` (just the key, nothing else)

**Example file content:**
```
AIzaSyB...xyz123
```

### 2. Google Vision API (for Step 1 - OCR Extraction)

**File:** `Config\API\google_service_account.json`

**What it's for:** Extracts text from PDF documents (ARP, Order, Approval)

**How to get it:**
1. Use the app's built-in wizard: Click "Google API Setup" in sidebar
2. OR manually:
   - Go to https://console.cloud.google.com/
   - Enable Vision API
   - Create Service Account
   - Download JSON key
   - Save as `google_service_account.json` in Config\API\

### 3. Google OAuth Credentials (for Steps 4-7, 10-12)

**File:** `Config\API\credentials.json`

**What it's for:**
- Email (Gmail API) - Steps 4, 6, 11, 12
- Calendar (Calendar API) - Step 7
- Contacts (People API) - Step 5
- Google Sheets - Step 10

**How to get it:**
1. Go to https://console.cloud.google.com/
2. Enable: Gmail API, Calendar API, People API, Sheets API
3. Create OAuth 2.0 Client ID (Desktop app)
4. Download JSON
5. Rename to `credentials.json`
6. Save in Config\API\

**First time you use these features:**
- App will open browser for you to authorize access
- Token files (`token_gmail.json`, `token_calendar.json`, etc.) will be auto-created
- You only need to authorize once

## How the App Finds These Files

The app automatically looks for config files in this folder:
- `C:\GoogleSync\GuardianShip_App\Config\`

All scripts have been updated to check here first before falling back to:
- Environment variables
- C:\configlocal\ (legacy location)

## Template Files for Distribution

When you give this app to end users, they need to create their own:

1. **google_maps_api_key.txt** - Each user gets their own ($200/month free credit)
2. **google_service_account.json** - Each user creates their own
3. **credentials.json** - Each user creates their own

You can include this README and the folder structure, but users must add their own keys.

## Security Notes

**DO NOT share your personal API keys/credentials with others!**

Each end user should:
1. Create their own Google Cloud project
2. Enable the APIs they need
3. Generate their own keys/credentials
4. Place them in their Config folder

**Exception:** The app can ship with a shared Service Account for Vision API if you want to centralize OCR costs, but be aware this counts against your quota/billing.

## Troubleshooting

### "Google Maps API key not found"
- Check that `Config\Keys\google_maps_api_key.txt` exists
- File should contain ONLY the API key (no extra spaces or lines)
- Key should start with `AIzaSy`

### "Vision API credentials not found"
- Check that `Config\API\google_service_account.json` exists
- File should be valid JSON
- Contains `type`: `service_account`

### "OAuth authorization required"
- Check that `Config\API\credentials.json` exists
- Browser will open for first-time authorization
- After authorization, token files are created automatically
- If stuck, delete the token_*.json files and try again

## Quick Setup Script

Run `Config\setup_config.bat` to:
- Check which files are present
- Validate file formats
- Guide you through missing setup

---

**For detailed setup instructions, see:**
- `GOOGLE_MAPS_SETUP.md` - Maps API setup
- `GETTING_STARTED.md` - Complete setup guide
