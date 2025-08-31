# CellPilot Beta Installation Summary

## 🚀 Quick Start for Beta Testers

### Step 1: Sign Up & Activate Beta
1. Go to https://www.cellpilot.io
2. Sign up with your Google email
3. Go to Dashboard
4. Click "Activate Beta Access"

### Step 2: Install CellPilot
1. Click "Install CellPilot Now" on dashboard
   - OR visit directly: https://script.google.com/macros/s/AKfycbzvsOpM-GJvXvdyV18HvgsE13ts4JKwSMLDcMKtHk8HMBh7IAH8ZQ58zV2h6xqtj39W/exec
2. Sign in with same Google account
3. Copy installation code
4. Open Google Sheets → Extensions → Apps Script
5. Paste code and add library
6. Save and refresh sheet

## 📦 Component URLs

### Main Library (CellPilot Core)
- **Script ID**: `1EZDAGoLY8UEMdbfKTZO-AQ7pkiPe-n-zrz3Rw0ec6VBBH5MdC43Avx0O`
- **Version**: 10
- **Identifier**: CellPilot

### Beta Installer Web App (v1.1)
- **URL**: https://script.google.com/macros/s/AKfycbzvsOpM-GJvXvdyV18HvgsE13ts4JKwSMLDcMKtHk8HMBh7IAH8ZQ58zV2h6xqtj39W/exec
- **Script ID**: `1PAypZbjwbnMf22dyHXItxbyRIhCZgIME4JtR8REpzk_PfvH17sg-h33f`
- **Features**:
  - Email-based beta access verification
  - Step-by-step installation guide
  - Automatic code generation

### API Endpoint
- **URL**: https://www.cellpilot.io/api/check-beta
- **Method**: POST
- **Payload**: `{"email": "user@example.com"}`
- **Response**: `{"hasBeta": true/false, "subscriptionTier": "beta/free/none"}`

## 🔒 Security Features

1. **Email Verification**: Only registered users can access
2. **Beta Tier Check**: Must have beta tier activated
3. **Source Code Protection**: Library model hides implementation
4. **Session-based Auth**: Uses Google account for verification

## 🛠️ Troubleshooting

### "Beta Access Required" Message
- Ensure you've signed up at www.cellpilot.io
- Activate beta access on your dashboard
- Use the same email for signup and Google account

### Installation Code Not Working
- Make sure to add the library (Script ID above)
- Set identifier as "CellPilot"
- Use version 10 or latest
- Save project after adding library

### Menu Not Appearing
- Refresh the Google Sheet
- Check Extensions menu
- First time: authorize when prompted

## 📊 Beta Testing Flow

```
User Journey:
1. Sign up at cellpilot.io → Creates profile
2. Activate beta on dashboard → Updates subscription_tier to 'beta'
3. Click installer link → Checks email against database
4. Copy installation code → Gets BETA_INSTALLER.gs content
5. Add to Apps Script → Proxies all functions to library
6. Use CellPilot → Full access to 200+ features
```

## 🔄 Updates

- **2025-08-31**: Initial beta deployment system
- **2025-08-31 v1.1**: Fixed email detection, updated to www.cellpilot.io

## 📝 Notes for Developers

### To Update Library:
```bash
npx clasp push
npx clasp deploy --description "New version description"
```

### To Update Installer:
```bash
cd beta-installer
npx clasp push
npx clasp deploy --description "Installer update"
```

### To Check Beta Users:
```sql
SELECT email, subscription_tier 
FROM auth.users u 
JOIN profiles p ON p.id = u.id 
WHERE p.subscription_tier = 'beta';
```

## 🎯 Success Metrics

- User can self-activate beta ✅
- Source code remains hidden ✅
- Easy installation process ✅
- Secure access control ✅
- Professional experience ✅