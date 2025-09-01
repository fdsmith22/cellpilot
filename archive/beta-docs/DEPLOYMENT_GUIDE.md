# CellPilot Beta Deployment Guide

## 🚀 Quick Start (Launch Beta Today!)

### Step 1: Deploy Logo to Site
Your logo files are already at `site/public/`. The Google Add-on needs a 48x48 PNG accessible via HTTPS.

**Option A: Use existing android-chrome PNG (Quick fix)**
```bash
# Copy existing PNG as logo-48x48.png
cp site/public/android-chrome-192x192.png site/public/logo-48x48.png
```

**Option B: Generate proper 48x48 PNG from SVG (Recommended)**
```bash
# Install ImageMagick if not installed
sudo apt-get install imagemagick

# Convert SVG to 48x48 PNG
convert -background transparent -resize 48x48 \
  site/public/logo/icons/icon-64x64.svg \
  site/public/logo-48x48.png
```

### Step 2: Deploy Next.js Site
```bash
cd site
npm run build
npm run start  # For local testing

# For production deployment (Vercel recommended)
vercel --prod
```

### Step 3: Update Logo URL in appsscript.json
After deploying your site, update the logo URL:
```json
"logoUrl": "https://www.cellpilot.io/logo-48x48.png",
```

### Step 4: Deploy Apps Script
```bash
# From project root
clasp push
clasp deploy --description "Beta release v1.0.0"

# Get the deployment URL
clasp open
```

## 🔐 Environment Configuration

### For Local Development (.env.local)
```env
# Supabase Configuration (Optional for beta)
NEXT_PUBLIC_SUPABASE_URL=your_supabase_url
NEXT_PUBLIC_SUPABASE_ANON_KEY=your_supabase_anon_key

# API Configuration
NEXT_PUBLIC_API_URL=https://api.cellpilot.io
NEXT_PUBLIC_SITE_URL=http://localhost:3000

# Stripe (For future monetization)
STRIPE_SECRET_KEY=sk_test_...
NEXT_PUBLIC_STRIPE_PUBLISHABLE_KEY=pk_test_...

# Analytics (Recommended for beta)
NEXT_PUBLIC_GA_ID=G-XXXXXXXXXX
NEXT_PUBLIC_MIXPANEL_TOKEN=your_mixpanel_token
```

### For Production (Vercel Environment Variables)
Add these in Vercel Dashboard > Settings > Environment Variables:
- All the above variables with production values
- Remove `NEXT_PUBLIC_SITE_URL` (Vercel provides VERCEL_URL)

## 📦 Deployment Methods Comparison

### Method 1: Site-Based Installation (Use This for Beta!)
**How it works:**
1. Users visit cellpilot.io
2. Click "Install" button
3. Copy the installation code
4. Paste in their Google Sheets Script Editor
5. Add library with Script ID: `1EZDAGoLY8UEMdbfKTZO-AQ7pkiPe-n-zrz3Rw0ec6VBBH5MdC43Avx0O`

**Pros:**
- ✅ Immediate deployment (no review)
- ✅ Full control over updates
- ✅ Can push updates instantly via `clasp push`
- ✅ Perfect for beta testing

**Cons:**
- ❌ Manual installation process
- ❌ Users need to trust your library

### Method 2: Google Workspace Add-ons (Future)
**How it works:**
1. Users search "CellPilot" in Google Workspace Marketplace
2. Click "Install"
3. Google handles everything automatically

**Pros:**
- ✅ One-click installation
- ✅ Higher trust (Google verified)
- ✅ Automatic updates
- ✅ Better discovery

**Cons:**
- ❌ 3-7 day review process
- ❌ Strict compliance requirements
- ❌ Updates need review

## 🎯 Beta Launch Checklist

### Immediate Tasks (Before Launch)
- [ ] Deploy logo to site public folder
- [ ] Deploy Next.js site to Vercel/hosting
- [ ] Update logo URL in appsscript.json
- [ ] Run `clasp push` to deploy latest code
- [ ] Test installation on fresh Google account

### Site Setup
- [ ] Create installation page at cellpilot.io/install
- [ ] Add installation instructions with screenshots
- [ ] Include Script ID prominently
- [ ] Add beta disclaimer and feedback form

### Beta Features Configuration
Your `FeatureGate.js` is already configured perfectly:
```javascript
BETA_CONFIG: {
  enabled: true,
  endDate: '2025-12-31',
  unlockAllFeatures: true,
  showBetaBadge: true,
  collectFeedback: true
}
```

## 📊 Tracking & Analytics

### Google Analytics Setup
1. Create GA4 property at analytics.google.com
2. Add to site's `_app.tsx`:
```typescript
import Script from 'next/script'

<Script
  src={`https://www.googletagmanager.com/gtag/js?id=${process.env.NEXT_PUBLIC_GA_ID}`}
  strategy="afterInteractive"
/>
```

### Apps Script Analytics
Your `ApiIntegration.js` already tracks events to `api.cellpilot.io`

## 🚦 Go-Live Steps (In Order)

### Today (Soft Launch)
1. **Deploy Site**
   ```bash
   cd site
   vercel --prod  # or your hosting command
   ```

2. **Deploy Apps Script**
   ```bash
   clasp push
   clasp deploy --description "Beta v1.0.0"
   ```

3. **Test Installation**
   - Create new Google Sheet
   - Extensions > Apps Script
   - Add library with your Script ID
   - Test core features

### Tomorrow (Beta Announcement)
1. Share with 10-20 trusted users
2. Create feedback form at cellpilot.io/feedback
3. Monitor error logs
4. Iterate based on feedback

### Next Week (Open Beta)
1. Public announcement
2. Submit to Product Hunt (optional)
3. Share in Google Sheets communities

### In 2 Weeks (Marketplace Submission)
1. Prepare marketplace listing
2. Create demo video
3. Submit for Google review
4. Continue beta via site method

## 🔑 API Keys & Security

### Development (Already Handled)
Your `FeatureGate.js` handles beta period gracefully:
```javascript
if (FeatureGate.isBetaPeriod()) {
  // All features unlocked, no API key needed
}
```

### Production (Post-Beta)
1. Implement Stripe subscriptions
2. Generate API keys per user
3. Store in Supabase
4. Validate on each operation

## 📱 Installation Instructions for Users

### What to Include on cellpilot.io/install:

```markdown
## Quick Installation (2 minutes)

1. Open any Google Sheet
2. Click Extensions → Apps Script
3. Delete any existing code
4. Copy and paste this:

\```javascript
function onOpen() {
  CellPilot.onOpen();
}

function onInstall(e) {
  CellPilot.onInstall(e);
}
\```

5. Click "Libraries" → "+"
6. Enter Script ID: `1EZDAGoLY8UEMdbfKTZO-AQ7pkiPe-n-zrz3Rw0ec6VBBH5MdC43Avx0O`
7. Click "Look up" → Select latest version → "Add"
8. Save and reload your sheet
9. Find "CellPilot" in your Extensions menu!
```

## 🎉 You're Ready to Launch!

Your architecture is solid, code is clean, and beta features are configured. Just:
1. Deploy your site
2. Push to Apps Script
3. Share with beta users

The Google Workspace Marketplace can wait until after successful beta. Launch today with the site method!