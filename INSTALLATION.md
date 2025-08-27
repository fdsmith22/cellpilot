# CellPilot Installation Instructions

## Quick Installation (Recommended)

### Step 1: Open your Google Sheet

### Step 2: Open Apps Script Editor
- Go to **Extensions > Apps Script**

### Step 3: Clear Any Existing Bridge Script
- Delete all existing code in the editor

### Step 4: Add CellPilot as a Library
1. Click the **"+"** button next to "Libraries" in the left sidebar
2. Enter this Script ID:
   ```
   1EZDAGoLY8UEMdbfKTZO-AQ7pkiPe-n-zrz3Rw0ec6VBBH5MdC43Avx0O
   ```
3. Click "Look up"
4. Select the latest version (should be version 3 or higher)
5. Keep the identifier as "CellPilot"
6. Click "Add"

### Step 5: Create the Connection Script
Replace all code in the editor with:

```javascript
/**
 * CellPilot Connection Script
 * This connects your sheet to the CellPilot library
 */

function onOpen() {
  CellPilot.onOpen();
}

function onInstall() {
  CellPilot.onInstall();
}

// Proxy functions for menu items
function showCellPilotSidebar() {
  CellPilot.showCellPilotSidebar();
}

function showDuplicateRemoval() {
  CellPilot.showDuplicateRemoval();
}

function showTextStandardization() {
  CellPilot.showTextStandardization();
}

function showDateFormatting() {
  CellPilot.showDateFormatting();
}

function showFormulaBuilder() {
  CellPilot.showFormulaBuilder();
}

function showFormulaTemplates() {
  CellPilot.showFormulaTemplates();
}

function showEmailAutomation() {
  CellPilot.showEmailAutomation();
}

function showCalendarIntegration() {
  CellPilot.showCalendarIntegration();
}

function showSettings() {
  CellPilot.showSettings();
}

function showHelp() {
  CellPilot.showHelp();
}

function showUpgradeDialog() {
  CellPilot.showUpgradeDialog();
}

function showUpgradeOptions() {
  CellPilot.showUpgradeOptions();
}

// Proxy functions for HTML callbacks
function removeDuplicatesProcess(options) {
  return CellPilot.removeDuplicatesProcess(options);
}

function previewDuplicates(options) {
  return CellPilot.previewDuplicates(options);
}

function generateFormulaFromDescription(description) {
  return CellPilot.generateFormulaFromDescription(description);
}

function checkFormulaBuilderAccess() {
  return CellPilot.checkFormulaBuilderAccess();
}

function insertFormulaIntoCell(formulaText) {
  return CellPilot.insertFormulaIntoCell(formulaText);
}

function standardizeText(options) {
  return CellPilot.standardizeText(options);
}
```

### Step 6: Save and Authorize
1. Save the script (Ctrl+S or Cmd+S)
2. Run the `onOpen` function once by:
   - Selecting `onOpen` from the function dropdown
   - Clicking the "Run" button
3. Authorize the script when prompted

### Step 7: Refresh Your Sheet
- Close and reopen your Google Sheet OR just refresh the page
- You should see the **CellPilot** menu appear in your menu bar

## Testing the Installation

### Check the Menu
You should see a "CellPilot" menu with these options:
- Open CellPilot (main sidebar)
- Data Cleaning submenu
- Formula Builder submenu
- Settings
- Help

### Test Basic Features
1. Select some data in your sheet
2. Go to **CellPilot > Open CellPilot**
3. The sidebar should open showing available tools

### Test Duplicate Removal
1. Select a range with some duplicate data
2. Go to **CellPilot > Data Cleaning > Remove Duplicates**
3. The duplicate removal sidebar should open

## Troubleshooting

### Menu doesn't appear
- Make sure you've run the `onOpen` function at least once
- Refresh your Google Sheet
- Check the Apps Script editor for any error messages

### "CellPilot is not defined" error
- Make sure the library was added correctly
- Check that the identifier is exactly "CellPilot"
- Verify you're using the latest version

### Functions not working
- Make sure all proxy functions are copied correctly
- Check that you've authorized the script
- Look for errors in Apps Script editor (View > Logs)

## Development Mode

If you want to modify CellPilot code directly:
1. Instead of using as a library, you can copy all source files
2. Use `clasp clone 1EZDAGoLY8UEMdbfKTZO-AQ7pkiPe-n-zrz3Rw0ec6VBBH5MdC43Avx0O`
3. Make your modifications
4. Push changes with `clasp push`

## Support

For issues or questions:
- Check the logs in Apps Script editor
- Visit https://www.cellpilot.io/docs
- Contact support@cellpilot.io