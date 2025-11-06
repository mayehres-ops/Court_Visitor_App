# Document AI Quick Start

## What is This?

Upgrade your app to recognize **handwritten ARPs** with 90-95% accuracy (vs 85-90% with Vision API).

## 5-Minute Setup

### 1. Enable API (2 minutes)
1. Go to: https://console.cloud.google.com
2. Select your existing project (same as Vision API)
3. Search for "Document AI API" and click **ENABLE**

### 2. Create Processor (2 minutes)
1. Go to: https://console.cloud.google.com/ai/document-ai/processors
2. Click **CREATE PROCESSOR**
3. Select **"Document OCR"**
4. Name it: **"ARP OCR Processor"**
5. Region: **us**
6. Click **CREATE**
7. **Copy the Processor ID** (looks like: `1234567890abcdef`)

### 3. Get Your Project ID (30 seconds)
1. Go to: https://console.cloud.google.com/home/dashboard
2. Look for **"Project Info"** on the left
3. **Copy your Project ID** (e.g., `guardianship-app-123456`)

### 4. Configure the App (30 seconds)
1. Open: `C:\GoogleSync\GuardianShip_App\DOCUMENT_AI_CONFIG_TEMPLATE.env`
2. Fill in these 2 values:
   ```
   DOCUMENT_AI_PROJECT_ID=your-project-id-here
   DOCUMENT_AI_PROCESSOR_ID=your-processor-id-here
   ```
3. Save as: `C:\configlocal\API\document_ai_config.env`
4. Restart the app

## Done!

The app will now use Document AI automatically for handwritten ARPs.

## Cost

- First 1,000 pages/month: **FREE**
- After that: $0.005 per page
- Your expected cost: **~$0.30-0.60/month** (well within free tier)

## How to Test

1. Run Step 1 (Guardian Extractor) with a handwritten ARP
2. Look for this in the output:
   ```
   Document AI extracted 2847 characters
   ```
3. Check if guardian names/addresses are more accurate!

## Troubleshooting

**"Document AI not configured"** → Follow steps 1-4 above

**"Document AI failed"** → App will automatically fall back to Vision API (still works!)

**"Billing not enabled"** → Go to https://console.cloud.google.com/billing and link your project

## Need More Help?

See the detailed guide: `DOCUMENT_AI_SETUP_GUIDE.md`
