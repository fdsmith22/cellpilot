# CellPilot Testing Guide

## Quick Start Testing Instructions

### Setup (5 minutes)

1. **Create Test Spreadsheet**
   ```
   - Go to Google Sheets
   - Create new spreadsheet
   - Name it "CellPilot Test Sheet"
   - Add sample data (see below)
   ```

2. **Install CellPilot as Library**
   ```
   - Extensions → Apps Script
   - Click "Libraries" → Add Library
   - Script ID: 1EZDAGoLY8UEMdbfKTZO-AQ7pkiPe-n-zrz3Rw0ec6VBBH5MdC43Avx0O
   - Choose latest version
   - Click "Add"
   ```

3. **Copy Test Proxy Code**
   ```
   - Delete all code in Code.gs
   - Copy entire contents of test-sheet-proxy.js
   - Paste into Code.gs
   - Save (Ctrl+S)
   - Close Apps Script editor
   - Refresh the spreadsheet
   ```

### Sample Test Data

Add this data to your test sheet:

**Sheet 1: Sales Data (A1:E11)**
```
Product,Category,Date,Quantity,Revenue
Laptop,Electronics,2024-01-15,5,6000
Phone,Electronics,2024-01-16,12,9600
Desk,Furniture,2024-01-17,3,900
Chair,Furniture,2024-01-18,8,1600
Tablet,Electronics,2024-01-19,7,3500
Monitor,Electronics,2024-01-20,4,1600
Bookshelf,Furniture,2024-01-21,2,400
Mouse,Electronics,2024-01-22,20,600
Keyboard,Electronics,2024-01-23,15,1500
Lamp,Furniture,2024-01-24,10,500
```

**Sheet 2: Duplicate Test Data (A1:C6)**
```
Name,Email,Score
John Doe,john@example.com,85
Jane Smith,jane@example.com,92
John Doe,john@example.com,85
Bob Wilson,bob@example.com,78
Jane Smith,jane@example.com,92
```

## Phase 1: Critical Path Tests

### Test 1: Menu Creation ✅
1. After refresh, look for "CellPilot" menu
2. Click menu to see all options
3. **Expected**: Full menu with all submenus visible

### Test 2: Main Dashboard
1. Click `CellPilot → Open CellPilot`
2. **Expected**: 
   - Sidebar opens on right
   - Shows 6 quick action buttons
   - Shows selection info at top
   - All dropdowns present

### Test 3: Preview/Apply Separation

#### 3A. Pivot Table Assistant
1. Select sales data (A1:E11)
2. Click `CellPilot → Advanced Tools → Pivot Table Assistant`
3. Click "Analyze Current Selection"
4. **Expected**: 
   - Step indicator shows progress
   - Data analysis displayed
   - AI suggestions appear
5. Select a suggestion
6. Click "Preview Result"
7. **Expected**: 
   - Preview modal shows sample data
   - No changes to spreadsheet yet
8. Click "Create Pivot Table"
9. **Expected**: 
   - Confirmation dialog appears
   - After confirm, pivot table created

#### 3B. Data Pipeline Manager
1. Click `CellPilot → Advanced Tools → Data Pipeline Manager`
2. Select CSV import
3. Paste test data:
   ```
   Name,Age,City
   Alice,30,NYC
   Bob,25,LA
   ```
4. Click "Preview Data"
5. **Expected**: 
   - Preview modal shows data
   - Stats displayed (3 rows, 3 columns)
6. Click "Import to Sheet"
7. **Expected**: 
   - Confirmation dialog
   - Data imported after confirm

#### 3C. Data Cleaning
1. Select duplicate data (Sheet 2)
2. Click `CellPilot → Data Cleaning → Remove Duplicates`
3. Click "Preview Changes"
4. **Expected**: 
   - Shows duplicates to be removed
   - No changes to data yet
5. Click "Remove Duplicates"
6. **Expected**: 
   - Confirmation dialog
   - Duplicates removed after confirm

### Test 4: Formula Builder
1. Select a cell
2. Click `CellPilot → Formula Builder → Natural Language Builder`
3. Type: "sum all revenue where category is Electronics"
4. Click "Generate Formula"
5. **Expected**: 
   - Formula displayed in preview
   - "Insert Into Cell" button appears
6. Click "Insert Into Cell"
7. **Expected**: Formula inserted

### Test 5: Tableize
1. Add messy data in new sheet:
   ```
   name|age|city
   Alice|30|NYC
   Bob|25|LA
   ```
2. Select the data
3. Click `CellPilot → Open CellPilot`
4. Click "Tableize Data"
5. Click "Preview Results"
6. **Expected**: 
   - Preview shows cleaned structure
7. Click "Apply Structure"
8. **Expected**: 
   - Data properly formatted as table

## Phase 2: Feature Tests

### Test 6: ML Features
1. Click `CellPilot → Enable ML Features`
2. Go to Data Cleaning
3. **Expected**: 
   - ML confidence indicators visible
   - Adaptive threshold suggestions

### Test 7: Industry Templates
1. Click `CellPilot → Industry Templates`
2. Select "Construction" → "Project Tracker"
3. Click "Preview Template"
4. **Expected**: Preview shows template structure
5. Click "Create Template"
6. **Expected**: New sheet with template created

### Test 8: Automation
1. Click `CellPilot → Automation`
2. Create a simple rule
3. Test trigger
4. **Expected**: Action executes

### Test 9: Export Functions
1. Select data range
2. Open Data Pipeline Manager
3. Go to Export tab
4. Select JSON format
5. Click "Preview Export"
6. **Expected**: JSON preview shown
7. Click "Download File"
8. **Expected**: File downloads

### Test 10: Undo Operations
1. Perform any operation (e.g., remove duplicates)
2. Look for undo notification
3. Click "Undo"
4. **Expected**: Operation reversed

## Common Issues & Solutions

### Issue: Menu doesn't appear
**Solution**: 
- Refresh browser (F5)
- Check script is saved
- Verify library is added correctly
- Run `testCellPilotConnection()` in Apps Script editor

### Issue: "CellPilot is not defined" error
**Solution**:
1. Open Apps Script editor
2. Click "Libraries" in left sidebar
3. Remove CellPilot if present
4. Click "Add Library"
5. Enter Script ID: `1EZDAGoLY8UEMdbfKTZO-AQ7pkiPe-n-zrz3Rw0ec6VBBH5MdC43Avx0O`
6. Set Identifier: `CellPilot` (exact case)
7. Choose "HEAD (Development mode)"
8. Click "Add"
9. Save and refresh

### Issue: "getAdaptiveDuplicateThreshold is not a function"
**Solution**:
1. In Apps Script editor, run this test:
   ```javascript
   function testLibrary() {
     console.log(testCellPilotConnection());
   }
   ```
2. If it fails, the library isn't loaded
3. If it succeeds, try:
   - Clear browser cache
   - Use incognito mode
   - Re-authorize the script

### Issue: Functions not working
**Solution**:
- Open Apps Script editor
- View → Execution transcript
- Check for error messages
- Ensure all permissions granted
- Try running functions directly in editor

### Issue: Preview buttons not showing
**Solution**:
- Clear browser cache
- Use Chrome browser
- Check console for errors (F12)
- Verify HTML templates are included

## Success Criteria

### Phase 1 (Must Pass - 100%)
- [x] Menu loads
- [x] Sidebar opens
- [x] Preview/Apply separation works
- [x] Confirmation dialogs appear
- [x] Basic operations complete

### Phase 2 (Should Pass - 90%)
- [ ] All features accessible
- [ ] ML features work
- [ ] Templates create correctly
- [ ] Import/Export functional
- [ ] Undo works

### Phase 3 (Nice to Have - 80%)
- [ ] Performance acceptable
- [ ] No console errors
- [ ] Smooth animations
- [ ] Mobile responsive

## Testing Checklist

Print this checklist and mark items as you test:

**Core Features**
- [ ] Main dashboard loads
- [ ] Quick actions work
- [ ] Selection detection works
- [ ] All menus accessible

**Preview/Apply Pattern**
- [ ] Pivot Table: Preview → Apply
- [ ] Data Pipeline: Preview → Import
- [ ] Data Cleaning: Preview → Apply
- [ ] Tableize: Preview → Apply
- [ ] Formula: Generate → Insert

**Confirmation Dialogs**
- [ ] Remove duplicates confirms
- [ ] Import data confirms
- [ ] Create pivot confirms
- [ ] Apply template confirms

**Data Operations**
- [ ] CSV import works
- [ ] JSON import works
- [ ] Export to CSV works
- [ ] Export to JSON works

**Advanced Features**
- [ ] ML features enable
- [ ] Confidence scores show
- [ ] Templates preview/create
- [ ] Undo operations work

## Reporting Issues

When reporting issues, include:
1. What you were trying to do
2. What happened instead
3. Error messages (if any)
4. Browser and OS
5. Screenshot if possible

Report at: https://github.com/anthropics/claude-code/issues

## Next Steps After Testing

1. **If all Phase 1 tests pass**: Ready for production
2. **If Phase 2 tests pass**: Ready for beta users
3. **If issues found**: Document and fix before deployment

---

**Testing Time Estimate**: 
- Phase 1: 15 minutes
- Phase 2: 20 minutes
- Phase 3: 10 minutes
- Total: ~45 minutes