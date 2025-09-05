# CellPilot Logo Setup for Google Sheets Sidebar

## 🎯 Quick Solution - Your Vercel Site is Already Running!

Since your Vercel site is already deployed at `cellpilot.io`, you just need to:

### Step 1: Add Logo to Your Site
The logo file needs to be in your `site/public` folder. I've created a simple SVG logo at:
- `/home/freddy/cellpilot/site/public/logo-48x48.svg`

### Step 2: Deploy to Vercel
```bash
cd /home/freddy/cellpilot/site
git add public/logo-48x48.svg
git commit -m "Add Google Sheets sidebar logo"
git push
```

Vercel will automatically deploy this. The logo will be available at:
- `https://www.cellpilot.io/logo-48x48.svg`

### Step 3: Verify Logo is Accessible
After deployment (usually takes 1-2 minutes), test the URL:
```bash
curl -I https://www.cellpilot.io/logo-48x48.svg
# Should return 200 OK
```

Or simply open in browser: https://www.cellpilot.io/logo-48x48.svg

### Step 4: Update Apps Script
Your `appsscript.json` is already configured correctly:
```json
"logoUrl": "https://www.cellpilot.io/logo-48x48.png",
```

Change this to `.svg` OR create a PNG version (see below).

### Step 5: Deploy Apps Script
```bash
clasp push
```

## 🖼️ Creating a PNG Version (If Required)

Google Sheets might require PNG format. Here are options:

### Option A: Use an Online Converter
1. Go to https://cloudconvert.com/svg-to-png
2. Upload the SVG file
3. Set dimensions to 48x48
4. Download and save as `logo-48x48.png` in `site/public/`

### Option B: Use Command Line (if ImageMagick installed)
```bash
# Install ImageMagick first if needed
sudo apt-get install imagemagick

# Convert SVG to PNG
convert -background transparent -resize 48x48 \
  /home/freddy/cellpilot/site/public/logo-48x48.svg \
  /home/freddy/cellpilot/site/public/logo-48x48.png
```

### Option C: Create PNG in Code (Node.js)
```bash
cd /home/freddy/cellpilot/site
npm install sharp --save-dev

# Create convert script
cat > convert-logo.js << 'EOF'
const sharp = require('sharp');

sharp('public/logo-48x48.svg')
  .resize(48, 48)
  .png()
  .toFile('public/logo-48x48.png')
  .then(() => console.log('Logo converted to PNG'))
  .catch(err => console.error('Error:', err));
EOF

node convert-logo.js
```

### Option D: Use Your Existing SVG Icons
You already have icon SVGs at:
- `/home/freddy/cellpilot/site/public/logo/icons/icon-64x64.svg` (869KB)

Copy one as your logo:
```bash
cp /home/freddy/cellpilot/site/public/logo/icons/icon-64x64.svg \
   /home/freddy/cellpilot/site/public/logo-48x48.svg
```

## 📝 What Gets Displayed Where

### In Google Sheets Sidebar
- **Icon**: The 48x48 logo appears at the top of the CellPilot sidebar
- **Location**: When users open CellPilot from Extensions menu
- **Requirements**: 
  - Must be HTTPS URL
  - Recommended 48x48 pixels
  - PNG or SVG format
  - Should be under 100KB

### In Google Workspace Marketplace (Future)
- **Different sizes needed**:
  - 32x32 - Small icon
  - 128x128 - Marketplace listing
  - 440x280 - Promotional tile

## ✅ Testing Your Logo

After deploying to Vercel:

1. **Test the URL directly**:
   ```
   https://www.cellpilot.io/logo-48x48.svg
   ```

2. **Test in Google Sheets**:
   - Open any Google Sheet
   - Extensions → Apps Script
   - Run your CellPilot code
   - Check if logo appears in sidebar

3. **If Logo Doesn't Appear**:
   - Check browser console for 404 errors
   - Verify URL is exactly as specified in appsscript.json
   - Try clearing browser cache
   - Use PNG instead of SVG

## 🚀 Complete Process (5 minutes)

```bash
# 1. Go to your site directory
cd /home/freddy/cellpilot/site

# 2. Ensure logo file exists
ls -la public/logo-48x48.*

# 3. Commit and push to Git (Vercel auto-deploys)
git add public/logo-48x48.svg
git commit -m "Add Google Sheets sidebar logo"
git push

# 4. Wait 2 minutes for Vercel deployment

# 5. Test the URL
curl https://www.cellpilot.io/logo-48x48.svg

# 6. Push to Apps Script
cd /home/freddy/cellpilot
clasp push

# Done! Logo will now appear in Google Sheets sidebar
```

## 💡 Pro Tips

1. **SVG vs PNG**: SVG scales better but PNG has better compatibility
2. **File Size**: Keep under 100KB for fast loading
3. **Transparent Background**: Use transparent BG for professional look
4. **High Contrast**: Ensure logo is visible on both light/dark themes
5. **Caching**: After updating logo, might need to clear browser cache

Your Vercel site at cellpilot.io will automatically serve any file you put in the `public` folder. So `public/logo-48x48.png` becomes `https://www.cellpilot.io/logo-48x48.png` - it's that simple!