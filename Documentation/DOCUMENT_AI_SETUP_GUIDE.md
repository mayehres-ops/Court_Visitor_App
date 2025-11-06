# Document AI Setup Guide

## What You're Doing:
Upgrading from Google Vision API to Document AI for better handwriting recognition (90-95% accuracy vs 85-90%)

## Cost:
- First 1,000 pages/month: FREE
- After that: $0.005 per page
- **Your expected cost: ~$0.30-0.60/month** (60-120 ARPs/month)

---

## Step-by-Step Setup (15 minutes)

### Step 1: Go to Google Cloud Console
1. Open your browser
2. Go to: https://console.cloud.google.com
3. Sign in with your Google account (same one you used for Vision API)

### Step 2: Select Your Project
1. At the top of the page, click the project dropdown
2. Select your existing project (the one you created for Vision API)
   - It's probably named something like "guardianship-app" or similar

### Step 3: Enable Document AI API
1. In the search bar at the top, type: **Document AI API**
2. Click on "Cloud Document AI API" in the results
3. Click the blue **"ENABLE"** button
4. Wait 30-60 seconds for it to enable

### Step 4: Verify Your Service Account Has Access
Since you already have a service account for Vision API, it should automatically work for Document AI too!

**To verify:**
1. Go to: https://console.cloud.google.com/iam-admin/iam
2. Find your service account (ends in @...iam.gserviceaccount.com)
3. Make sure it has one of these roles:
   - **Owner** (full access) OR
   - **Editor** (can use APIs) OR
   - **Document AI API User** (specific permission)

**If it doesn't have the right role:**
1. Click the pencil icon next to your service account
2. Click "+ ADD ANOTHER ROLE"
3. Search for: **Document AI API User**
4. Click "Save"

### Step 5: Create a Processor (Document AI specific)
1. Go to: https://console.cloud.google.com/ai/document-ai/processors
2. Click **"CREATE PROCESSOR"**
3. Select: **"Document OCR"** (best for mixed handwritten/typed documents)
4. Click **"CREATE"**
5. Give it a name: **"ARP OCR Processor"**
6. Select region: **us** (United States)
7. Click **"CREATE"**

8. **IMPORTANT:** Copy the **Processor ID** that appears!
   - It looks like: `1234567890abcdef`
   - Save this - you'll need it in the Python code!

### Step 6: Get Your Project ID
1. Go to: https://console.cloud.google.com/home/dashboard
2. Look for "Project Info" widget on the left
3. Copy your **Project ID** (not Project Name!)
   - Example: `guardianship-app-123456`

---

## What You Need to Provide Me:

After completing the above steps, give me these 3 things:

1. **Project ID**: ________________________
2. **Processor ID**: ________________________
3. **Location**: (probably "us")

---

## Notes:

✅ You're using the same service account JSON file as Vision API
- Located at: `C:\configlocal\API\google_service_account.json`
- No need to download a new one!

✅ No new credit card setup needed
- Same billing account as Vision API

✅ First 1,000 pages/month are FREE
- You'll use ~60-120 pages/month
- Well within free tier!

---

## Troubleshooting:

**"I don't see my project"**
- Make sure you're signed in with the correct Google account
- The account that set up Vision API

**"Enable button is grayed out"**
- The API might already be enabled - check for a "Manage" button instead

**"Billing not enabled"**
- Go to: https://console.cloud.google.com/billing
- Make sure your project is linked to a billing account
- (Same account used for Vision API)

---

## After Setup - Configure the App:

Once you have your Project ID and Processor ID from Google Cloud Console:

1. Open the template file: `C:\GoogleSync\GuardianShip_App\DOCUMENT_AI_CONFIG_TEMPLATE.env`

2. Fill in your values:
   ```
   DOCUMENT_AI_PROJECT_ID=your-project-id-here
   DOCUMENT_AI_PROCESSOR_ID=your-processor-id-here
   DOCUMENT_AI_LOCATION=us
   ```

3. Save the file as: `C:\configlocal\API\document_ai_config.env`

4. Restart the GuardianShip app

**That's it!** The app will automatically:
- Load your Document AI configuration
- Use Document AI for handwritten ARP processing (90-95% accuracy)
- Fall back to Vision API if Document AI isn't configured or fails
- Still use Tesseract as a last resort

Your handwriting accuracy will jump from 85-90% to 90-95%!
