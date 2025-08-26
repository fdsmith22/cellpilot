# CellPilot Development Workflow

## Direct Development (No Libraries, No Bridges)

### Your Current Setup
- **Script ID**: `1EZDAGoLY8UEMdbfKTZO-AQ7pkiPe-n-zrz3Rw0ec6VBBH5MdC43Avx0O`
- **Local Files**: `/home/freddy/cellpilot/`
- **Script URL**: https://script.google.com/d/1EZDAGoLY8UEMdbfKTZO-AQ7pkiPe-n-zrz3Rw0ec6VBBH5MdC43Avx0O/edit

## Simple Development Workflow

### 1. Make Changes Locally
Edit any files in your cellpilot folder:
- `src/Code.js` - Main entry point
- `src/Config.js` - Settings and configuration
- `src/DataCleaner.js` - Data cleaning functions
- `src/FormulaBuilder.js` - Formula generation
- `html/*.html` - UI templates

### 2. Push Changes to Google
```bash
cd /home/freddy/cellpilot
clasp push
```

### 3. Test in Google Sheets
Open any Google Sheet and:
1. Go to **Extensions > Apps Script**
2. Click the "Libraries +" button
3. Add your script ID (or if already added, update to "HEAD" version for development)
4. In the sheet's script, just have:

```javascript
function onOpen() {
  CellPilot.onOpen();
}
```

OR for truly direct testing:

1. Open your main script: https://script.google.com/d/1EZDAGoLY8UEMdbfKTZO-AQ7pkiPe-n-zrz3Rw0ec6VBBH5MdC43Avx0O/edit
2. Click "Test deployments" button (clock icon)
3. This opens a test sheet with your code running directly

## Even Simpler: Container-Bound Testing

For the simplest testing:
1. Create a Google Sheet for testing
2. Extensions > Apps Script
3. Delete everything
4. In terminal: `clasp push --force`
5. This makes your test sheet run CellPilot directly

## Making Changes

### Example: Update duplicate removal threshold
1. Edit `src/DataCleaner.js`
2. Run `clasp push`
3. Refresh your test sheet
4. Test the feature

### Example: Change menu text
1. Edit `src/Code.js`
2. Run `clasp push`
3. Refresh sheet to see new menu

## Quick Commands

```bash
# Push your changes
clasp push

# Check what will be pushed
clasp status

# See your code in browser
echo "https://script.google.com/d/1EZDAGoLY8UEMdbfKTZO-AQ7pkiPe-n-zrz3Rw0ec6VBBH5MdC43Avx0O/edit"
```

## No More Bridges!
You don't need:
- Bridge scripts
- Multiple projects
- Complex library setups

Just:
1. Edit locally
2. `clasp push`
3. Test in your sheet

That's it!