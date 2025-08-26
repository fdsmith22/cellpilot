# CellPilot Testing Guide

## Quick Test Setup

### Method 1: Test as Library (Recommended)

1. **Open a test Google Sheet**

2. **Go to Extensions > Apps Script**

3. **Clear any existing code**

4. **Add CellPilot as a library:**
   - Click Libraries "+"
   - Script ID: `1EZDAGoLY8UEMdbfKTZO-AQ7pkiPe-n-zrz3Rw0ec6VBBH5MdC43Avx0O`
   - Version: HEAD (for development)
   - Identifier: CellPilot

5. **Copy proxy functions from test-sheet-proxy.js:**
   - Open `/home/freddy/cellpilot/test-sheet-proxy.js`
   - Copy ALL the code
   - Paste into your test sheet's Apps Script editor

6. **Save and authorize:**
   - Save the script (Ctrl+S)
   - Run onOpen function once
   - Authorize when prompted

7. **Refresh your Google Sheet**
   - The CellPilot menu should appear

### Method 2: Direct Testing (Alternative)

1. **Open the main script directly:**
   ```
   https://script.google.com/d/1EZDAGoLY8UEMdbfKTZO-AQ7pkiPe-n-zrz3Rw0ec6VBBH5MdC43Avx0O/edit
   ```

2. **Click "Test deployments" (clock icon)**

3. **This opens a test sheet with your code running directly**

## Test Checklist

### Basic Functionality
- [ ] CellPilot menu appears
- [ ] "Open CellPilot" sidebar loads
- [ ] Settings dialog opens

### Data Cleaning Features
- [ ] Duplicate removal opens
- [ ] Preview duplicates works
- [ ] Remove duplicates processes correctly
- [ ] Text standardization works

### Formula Builder
- [ ] Formula builder sidebar opens
- [ ] Natural language input accepted
- [ ] Formula generation works
- [ ] Insert formula into cell works

### Error Checking
- [ ] Check Apps Script logs for errors
- [ ] Verify all proxy functions work
- [ ] Test with different data types

## Common Issues

### "Function not found" errors
- Make sure ALL proxy functions from test-sheet-proxy.js are copied
- Verify library is added with identifier "CellPilot"

### Menu doesn't appear
- Run onOpen function manually once
- Refresh the Google Sheet

### Changes not reflecting
- After `clasp push`, refresh test sheet
- For library testing, no version update needed with HEAD

## Development Workflow

1. Make changes locally in `/home/freddy/cellpilot/src/`
2. Run `clasp push`
3. Refresh test sheet
4. Test the changes
5. Check logs for any errors

## Logs and Debugging

View logs in Apps Script editor:
- View > Logs
- Or use Logger.log() in code

Check execution transcript:
- View > Executions