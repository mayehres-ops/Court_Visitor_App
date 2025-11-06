# Publishing Google API OAuth Consent Screen

## Why Publish?

In "Testing" mode, OAuth tokens expire every 7 days, requiring users to re-authorize.
Publishing the app removes this limitation.

## Steps to Publish:

### 1. Go to Google Cloud Console
1. Visit: https://console.cloud.google.com
2. Select your project (the one with your Court Visitor App APIs)

### 2. Navigate to OAuth Consent Screen
1. In the left menu, click **"APIs & Services"** → **"OAuth consent screen"**
2. You should see your current configuration

### 3. Review Information
Make sure these are filled in:
- **App name**: Court Visitor App (or your preferred name)
- **User support email**: Your email
- **Developer contact information**: Your email
- **App logo**: (Optional, but recommended)
- **App domain**: (Can leave blank for internal use)
- **Authorized domains**: (Can leave blank for internal use)

### 4. Review Scopes
1. Click **"Edit App"**
2. Go to **"Scopes"** section
3. Verify you have the scopes you need:
   - Gmail API (for sending emails)
   - Google Calendar API (for creating events)
   - Google Sheets API (for reading data)
   - People API (for contacts)

### 5. Add Test Users (Important!)
1. In **"Test users"** section
2. Add emails of users who will use the app
3. You can add up to 100 test users

### 6. Publish Options

You have TWO options:

#### **Option A: Keep in Testing + Add All Users** (EASIEST)
- Stay in "Testing" mode
- Add ALL Court Visitor emails as "Test Users"
- Tokens won't expire for test users
- No verification needed
- **RECOMMENDED if you have <100 users**

#### **Option B: Publish to Production** (If you need more than 100 users)
1. Click **"Publish App"** button
2. **WARNING**: This triggers a Google verification process
3. Google will review your app (can take weeks)
4. They may ask for:
   - Privacy policy
   - Terms of service
   - Video demo
   - Justification for scopes

---

## RECOMMENDED APPROACH:

### For Internal/Limited Use (Best Option):

1. **Stay in Testing Mode**
2. **Add all Court Visitor email addresses as Test Users**
3. **Benefits**:
   - ✅ No expiration for test users
   - ✅ No Google verification needed
   - ✅ Can add up to 100 users
   - ✅ Quick and easy

### Steps:
1. Go to OAuth consent screen
2. Scroll to **"Test users"**
3. Click **"Add Users"**
4. Enter Court Visitor email addresses (one per line):
   ```
   courtvisitor1@example.com
   courtvisitor2@example.com
   courtvisitor3@example.com
   ```
5. Click **"Save"**

**That's it!** Test users won't need to re-authorize every week.

---

## If You Need More Than 100 Users:

Then you must publish to production:

1. Prepare required documents:
   - Privacy policy (what data you collect/use)
   - Terms of service
   - App homepage/documentation

2. Click **"Publish App"**

3. Submit for verification

4. Wait for Google approval (1-6 weeks)

---

## Current Status Check:

To see your current status:
1. Go to: https://console.cloud.google.com/apis/credentials/consent
2. Look at **"Publishing status"**:
   - **Testing** = Tokens expire in 7 days (unless user is added as test user)
   - **In production** = Tokens don't expire

---

## Notes:

- **Internal use**: Keep in Testing mode, add all users as test users
- **Public distribution**: Must publish to production
- **Hybrid**: Can start in Testing and publish later if needed
- Test users can be added/removed anytime
