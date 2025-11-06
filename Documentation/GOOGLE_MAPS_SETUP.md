# Google Maps API Setup for Step 3 (Generate Route Map)

## Why Do You Need This?

Step 3 creates a map showing where all your ward addresses are located, with numbered markers and a legend. This helps you plan efficient routes for visits.

**Without Google Maps API:** You get dots on a blank background (no streets, no map)
**With Google Maps API:** You get dots overlaid on an actual street map showing the area

## What You Get With Google Maps API

1. **Better Geocoding** - More accurate address-to-coordinates conversion
2. **Actual Map Background** - Real street maps showing neighborhoods, roads, landmarks
3. **Route Planning** - Optimal visit order based on driving directions
4. **Travel Time Estimates** - Estimated time between locations

## Setup Instructions

### Step 1: Get a Google Maps API Key

1. **Go to Google Cloud Console**
   - Visit: https://console.cloud.google.com/

2. **Create a New Project** (or use existing)
   - Click "Select a project" dropdown at top
   - Click "NEW PROJECT"
   - Name it: "GuardianShip App"
   - Click "CREATE"

3. **Enable Required APIs**
   - Go to "APIs & Services" > "Library"
   - Search for and ENABLE these APIs:
     - ✅ **Maps Static API** (for map images)
     - ✅ **Geocoding API** (for address lookup)
     - ✅ **Directions API** (for route planning)

4. **Create API Key**
   - Go to "APIs & Services" > "Credentials"
   - Click "+ CREATE CREDENTIALS" > "API key"
   - Copy the API key (looks like: `AIzaSyB...xyz123`)
   - Click "EDIT API KEY" to restrict it (recommended):
     - Application restrictions: None (or HTTP referrers)
     - API restrictions: Select "Restrict key"
     - Select: Maps Static API, Geocoding API, Directions API
   - Click "SAVE"

### Step 2: Set the API Key on Your Computer

**Option A: Use the Batch File (Easiest)**
1. Double-click `SET_GOOGLE_MAPS_KEY.bat` in the app folder
2. Paste your API key when prompted
3. **Restart your computer** (required!)

**Option B: Manual Setup (Permanent)**
1. Right-click "This PC" → Properties
2. Click "Advanced system settings"
3. Click "Environment Variables"
4. Under "System variables" click "New"
5. Variable name: `GOOGLE_MAPS_API_KEY`
6. Variable value: Paste your API key
7. Click OK on all dialogs
8. **Restart your computer** (required!)

### Step 3: Verify It Works

1. Launch the GuardianShip App
2. Click **Step 3: Generate Route Map**
3. Check the console output - you should see:
   ```
   [init] Google key detected: True; Directions: True
   [map] Google Static Maps basemap added successfully
   ```
4. Open the generated `Ward_Map_Sheet.docx`
5. You should see an actual street map with your ward locations!

## Pricing (As of 2024)

Google provides **$200 free credit per month** which covers:

- **Maps Static API:** $2 per 1,000 requests
- **Geocoding API:** $5 per 1,000 requests
- **Directions API:** $5 per 1,000 requests

**For typical use:**
- 50 cases/month = ~150 API calls
- Cost: ~$2/month
- **FREE** with the $200 credit!

You only pay if you exceed $200/month in usage.

## Troubleshooting

### "No basemap tiles" message
- API key not set as environment variable
- API key not valid
- Required APIs not enabled in Google Cloud Console
- Computer not restarted after setting environment variable

### "Google Static Maps returned status 403"
- API key restrictions too strict
- APIs not enabled
- Billing not enabled on Google Cloud account

### Still shows blank background
- Check that you restarted the computer after setting the API key
- Verify the environment variable is set:
  - Open Command Prompt
  - Type: `echo %GOOGLE_MAPS_API_KEY%`
  - Should show your API key, not blank

## For End Users (Distribution)

When you give the app to others, they need to:

1. Get their own Google Maps API key (free $200/month credit)
2. Run `SET_GOOGLE_MAPS_KEY.bat`
3. Restart their computer
4. App will work with maps!

**Include this file (GOOGLE_MAPS_SETUP.md) in your distribution package.**

---

**Note:** The app still works WITHOUT Google Maps API - you just get numbered dots on a plain background instead of on a real map. But for efficient route planning, the map background is essential!
