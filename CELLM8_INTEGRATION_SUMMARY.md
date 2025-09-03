# CellM8 Integration & Development Guide

## Project Overview
CellPilot is a Google Sheets add-on distributed as a Google Apps Script library. CellM8 is a presentation generator feature that transforms spreadsheet data into Google Slides presentations.

### Key Project Locations
- **Main Library**: `/home/freddy/cellpilot/apps-script/main-library/`
- **Beta Installer**: `/home/freddy/cellpilot/apps-script/beta-installer/`
- **Test Proxy**: `/home/freddy/cellpilot/test-sheet-proxy.js`
- **Deployment**: Using `clasp` CLI for Google Apps Script

## CellM8 Development Process - CRITICAL STEPS

### When Adding New CellM8 Features - Follow This Exact Process:

#### Step 1: Add Function to CellM8.js
```javascript
// In /apps-script/main-library/src/CellM8.js
const CellM8 = {
  // Add your new function here
  yourNewFunction: function(params) {
    try {
      // Implementation
      return { success: true, data: result };
    } catch (error) {
      Logger.error('Error in yourNewFunction:', error);
      return { success: false, error: error.toString() };
    }
  }
};
```

#### Step 2: Create Proxy in Code.js
```javascript
// In /apps-script/main-library/src/Code.js
// Add in the CellM8 Helper Functions section (around line 1761)
function yourNewCellM8Function(params) {
  return CellM8.yourNewFunction(params);
}
```

#### Step 3: Export in Library.js
```javascript
// In /apps-script/main-library/src/Library.js
// Add in CellM8 section (around line 160)
var yourNewCellM8Function = yourNewCellM8Function || function(params) { 
  return yourNewCellM8Function(params); 
};
```

#### Step 4: Add to Beta Installer
```javascript
// In /apps-script/beta-installer/Code.gs
// Add in CellM8 section (around line 712)
function yourNewCellM8Function(params) { 
  return CellPilot.yourNewCellM8Function(params); 
}
```

#### Step 5: Add to Test Sheet Proxy
```javascript
// In /test-sheet-proxy.js
// Add in CellM8 section (around line 2240)
function yourNewCellM8Function(params) {
  return CellPilot.yourNewCellM8Function(params);
}
```

#### Step 6: Update UI Template if Needed
```javascript
// In /apps-script/main-library/CellM8Template.html
// Call your function from JavaScript
google.script.run
  .withSuccessHandler(function(result) {
    // Handle success
  })
  .withFailureHandler(function(error) {
    // Handle error
  })
  .yourNewCellM8Function(params);
```

### Deployment Commands
```bash
# From /apps-script/main-library
clasp push

# From /apps-script/beta-installer  
cd ../beta-installer && clasp push

# Deploy new version
clasp deploy --description "CellM8: Added new feature"

# Commit changes
git add -A
git commit -m "CellM8: Description of changes"
```

## Architecture & Function Export Pattern

### Critical Understanding: How CellPilot Functions Work
CellPilot uses a **multi-layer proxy architecture** for function distribution:

1. **Core Implementation** (`/apps-script/main-library/src/`)
   - Actual function implementations live here
   - Example: `Code.js` contains `showCellM8()` implementation

2. **Library Export Layer** (`/apps-script/main-library/src/Library.js`)
   - EVERY function must be exported here for library access
   - Uses conditional variable assignment pattern:
   ```javascript
   var showCellM8 = showCellM8 || function() { return showCellM8(); };
   var createPresentation = createPresentation || function(config) { return CellM8.createPresentation(config); };
   ```

3. **Beta Installer Proxy** (`/apps-script/beta-installer/Code.gs`)
   - Simple proxy functions that call the library
   ```javascript
   function showCellM8() { return CellPilot.showCellM8(); }
   function createPresentation(config) { return CellPilot.createPresentation(config); }
   ```

4. **Test Sheet Proxy** (`/test-sheet-proxy.js`)
   - Development/testing proxy file
   - Must include ALL functions for local testing
   ```javascript
   function showCellM8() { 
     return CellPilot.showCellM8(); 
   }
   ```

### IMPORTANT: When Adding New Features
**EVERY new function must be added to ALL FOUR locations:**
1. Implementation in `/apps-script/main-library/src/` files
2. Export in `/apps-script/main-library/src/Library.js`
3. Proxy in `/apps-script/beta-installer/Code.gs`
4. Proxy in `/test-sheet-proxy.js`

Missing any location will cause "function not found" errors!

## CellM8 Integration Details

### What is CellM8?
CellM8 is a presentation helper that transforms spreadsheet data into Google Slides presentations. It was fully integrated on 2025-09-02.

### Integration Components

#### 1. Core Implementation Files
- **CellM8.js** (`/apps-script/main-library/src/CellM8.js`)
  - Main CellM8 class with all presentation generation logic
  - 34 total functions including:
    - `createPresentation(config)`
    - `extractSheetData(source)`
    - `generateSlides(presentation, data, config)`
    - Template application functions
    - Data visualization functions

- **CellM8SlideGenerator.js** (`/apps-script/main-library/src/CellM8SlideGenerator.js`) - THE ONLY SLIDE GENERATOR
  - Research-based implementation using best practices
  - Three-request method to eliminate placeholders
  - Build-from-scratch approach for consistency
  - Intelligent data analysis and visualization
  - Safe margins and error handling

**Note**: CellM8SlideGenerator.js is the ONLY slide generator. It's an internal helper module called by CellM8.js and does not require proxy functions in the 4-location pattern.

#### 2. UI Template
- **CellM8Template.html** (`/apps-script/main-library/CellM8Template.html`)
  - Uses CellPilot design system via `<?!= include('SharedStyles'); ?>`
  - Navigation header with back button
  - Card-based layout matching other panels
  - Custom template options for branding
  - NO EMOJIS (per user requirement)

#### 3. Menu Integration
- Added to main menu in `Code.js`:
```javascript
.addItem('CellM8 - Presentation Helper', 'showCellM8')
```

- Sidebar card in main panel (made compact):
```javascript
CardService.newTextParagraph()
  .setText('<b>CellM8 Presentation Helper</b>\nTransform data into presentations')
```

## Current CellM8 Implementation Status (2025-09-02)

### Core Functions (Implemented & Working)
```javascript
// Basic operations
showCellM8()                          // Opens the CellM8 sidebar
testCellM8Function()                  // Tests CellM8 connectivity
extractCellM8Data()                   // Extracts data from active sheet
getCurrentSelection()                 // Gets selected range info

// Presentation operations  
previewCellM8Presentation(config)    // Generates preview
createCellM8Presentation(config)     // Creates actual presentation
generateCellM8Slides(presentation, data, config) // Generates slide content

// Analysis functions
analyzeCellM8Data(data)              // Analyzes data for insights
detectColumnType(columnData)         // Detects column data types
calculateColumnStats(columnData)     // Calculates statistics

// Template and styling
applyCellM8Template(presentation, templateName) // Applies templates
createCellM8Chart(data, chartType)   // Creates charts

// Export and sharing
exportCellM8Presentation(presentationId, format) // Export as PDF/PPTX
shareCellM8Presentation(presentationId, emails, permission) // Share
```

### Functions to Re-implement (From Original)
```javascript
// Core functions
showCellM8()
createPresentation(config)
extractSheetData(source)
generateSlides(presentation, data, config)
applyTemplate(presentation, template)
createTitleSlide(presentation, config)
createDataSlide(presentation, data, config)
createChartSlide(presentation, chartData, config)
createTableSlide(presentation, tableData, config)
createSummarySlide(presentation, data, config)

// Helper functions
formatDataForPresentation(data)
generateChartFromData(data, chartType)
createTableFromData(data, maxRows)
applyCustomStyling(slide, styles)
addTransitions(presentation, transitionType)
exportPresentation(presentation, format)
sharePresentation(presentation, emails, permission)
getTemplates()
saveAsTemplate(presentation, name)
loadTemplate(templateId)

// Data processing functions
aggregateData(data, groupBy, aggregationType)
filterData(data, criteria)
sortData(data, column, direction)
pivotData(data, rowField, columnField, valueField)
calculateStatistics(data)

// Validation functions  
validateDataSource(source)
validateConfiguration(config)
checkPermissions()

// Error handling functions
handleCellM8Error(error)
logActivity(action, details)

// Integration functions
connectToSheets(spreadsheetId)
connectToSlides(presentationId)
refreshData()
schedulePresentation(config, cronExpression)
```

## UI Design System Patterns

### SharedStyles Include
**CRITICAL**: Must use `HtmlService.createTemplateFromFile()` not `createHtmlOutputFromFile()`
```javascript
// CORRECT - Processes server-side includes
const html = HtmlService.createTemplateFromFile('CellM8Template')
  .evaluate()
  
// WRONG - Shows <?!= include('SharedStyles'); ?> as text
const html = HtmlService.createHtmlOutputFromFile('CellM8Template')
```

### Standard UI Components
1. **Navigation Header**:
```html
<div class="nav-header">
  <button class="nav-back" onclick="backToMain()">←</button>
  <div class="nav-title">
    <h2>Feature Name</h2>
    <p>Feature description</p>
  </div>
</div>
```

2. **Card Sections**:
```html
<div class="card">
  <div class="card-header">
    <h3>Section Title</h3>
    <p>Section description</p>
  </div>
  <div class="card-body">
    <!-- Content -->
  </div>
</div>
```

3. **Form Controls**:
```html
<input type="text" class="form-control" placeholder="...">
<select class="form-control">
<button class="btn btn-primary">
```

### Design Requirements
- NO EMOJIS in UI (user preference)
- Consistent color scheme via CSS variables
- Card-based layouts with proper spacing
- Navigation consistency across all panels

## Deployment Process

### Steps for Deploying Changes
1. **Push to Apps Script**:
```bash
clasp push
```

2. **Create New Deployment** (if needed):
```bash
clasp deploy --description "Feature description"
```

3. **Check Deployments**:
```bash
clasp deployments
```

### Current Active Deployment
- **ID**: AKfycbyPA5zqXx3pQOOG_m5AdZP38XQ32mYJQF0IBxKpe14xnMG9Nb4RFAIpg7YQOqHc05wBHA
- **Version**: @21
- **Description**: CellM8 Advanced: 20 chart types, AI insights, master templates

## Project Context Checking Process

### Before Developing Any Feature
1. **Check existing patterns**:
```bash
# Find similar features
grep -r "function show" src/
# Check UI templates
ls -la *Template.html
```

2. **Verify design system**:
```bash
# Check SharedStyles usage
grep "include('SharedStyles')" *.html
# Check existing card layouts
grep "card-header" *.html
```

3. **Check function exports**:
```bash
# Verify Library.js exports
grep "var functionName" src/Library.js
# Check beta installer proxies
grep "function functionName" ../beta-installer/Code.gs
# Check test proxy
grep "function functionName" ../../test-sheet-proxy.js
```

## Common Issues & Solutions

### Issue 1: "google.script.run.functionName is not a function"
**Cause**: Function not exported in Library.js or missing from proxy files
**Solution**: Add to all 4 locations (implementation, Library.js, beta-installer, test-proxy)

### Issue 2: Template shows "<?!= include('SharedStyles'); ?>" as text
**Cause**: Using `createHtmlOutputFromFile` instead of `createTemplateFromFile`
**Solution**: Use `HtmlService.createTemplateFromFile('TemplateName').evaluate()`

### Issue 3: UI styling doesn't match
**Cause**: Not using standard design system classes
**Solution**: Copy patterns from DataPipelineTemplate.html or TableizeTemplate.html

### Issue 4: Function works in dev but not in production
**Cause**: Forgot to create new deployment after changes
**Solution**: Run `clasp push` then `clasp deploy`

## Testing Checklist for New Features

1. ✅ Function implemented in appropriate src/ file
2. ✅ Function exported in Library.js
3. ✅ Proxy added to beta-installer/Code.gs
4. ✅ Proxy added to test-sheet-proxy.js
5. ✅ UI template uses SharedStyles include
6. ✅ UI follows design system (cards, navigation, forms)
7. ✅ No emojis in UI
8. ✅ Menu item added if needed
9. ✅ Sidebar card added if needed
10. ✅ Changes pushed via clasp
11. ✅ New deployment created if needed

## Recent Changes Summary (2025-09-02)

### Files Modified for CellM8 Integration
1. **src/Code.js**
   - Added `showCellM8()` function with proper template processing
   - Added CellM8 menu item
   - Added compact CellM8 card to sidebar

2. **src/Library.js**
   - Added all 34 CellM8 function exports

3. **beta-installer/Code.gs**
   - Added all 34 CellM8 proxy functions

4. **test-sheet-proxy.js**
   - Added all 34 CellM8 proxy functions

5. **CellM8Template.html**
   - Complete redesign with CellPilot design system
   - Added navigation header
   - Card-based layout
   - Custom template options
   - Removed all emojis

6. **src/CellM8.js**
   - Full implementation of presentation generation
   - 34 functions for complete functionality

## Git Status
- Repository: Main branch, up to date
- Latest commit: CellM8 integration completed
- All changes pushed to Apps Script

## Next Steps for Continuation
When resuming development:
1. Run `git pull` to ensure latest changes
2. Check `clasp deployments` for current version
3. Verify all proxy files are in sync
4. Use this document as reference for adding new features

## Environment Details
- Platform: Linux
- Working directory: `/home/freddy/cellpilot/apps-script/main-library`
- Google Apps Script Library ID: `1EZDAGoLY8UEMdbfKTZO-AQ7pkiPe-n-zrz3Rw0ec6VBBH5MdC43Avx0O`
- Beta installer uses this library via `CellPilot` identifier

## Key Commands Reference
```bash
# Deploy changes
clasp push
clasp deploy --description "Description"

# Check status
clasp deployments
git status
git diff

# Search patterns
grep -r "pattern" src/
grep "functionName" src/Library.js

# File locations
cd /home/freddy/cellpilot/apps-script/main-library
cd /home/freddy/cellpilot/apps-script/beta-installer
```

## Critical Rules for Development
1. **ALWAYS** add new functions to all 4 locations
2. **ALWAYS** use `createTemplateFromFile` for templates with includes
3. **NEVER** use emojis in UI
4. **ALWAYS** follow existing design patterns
5. **ALWAYS** test in both development and after deployment
6. **ALWAYS** create new deployment for production changes

---
*This summary created on 2025-09-02 after successful CellM8 integration*
*Updated on 2025-09-03 - Consolidated to single generator:*
- *CellM8SlideGenerator.js: The ONLY implementation*
- *Uses research-based three-request method*
- *Properly handles all placeholder and formatting issues*
*Use this document to maintain consistency when adding new features*

## Current CellM8 Implementation (2025-09-03)

### CellM8SlideGenerator - The Single Implementation
- **Three-Request Method**: Properly eliminates default placeholders
- **Build From Scratch**: No dependency on templates, full control
- **Intelligent Data Analysis**: Automatic type detection and statistics
- **Dynamic Slide Planning**: Adapts to data characteristics
- **Safe Margins**: Prevents content overflow (50px sides, 40px vertical)
- **Professional Themes**: Clean light and dark modes
- **Smart Tables**: Limits to 8 rows × 5 columns with truncation notes
- **Chart Selection**: Automatically chooses best visualization type
- **Error Handling**: Graceful fallbacks if operations fail

This is the ONLY generator that consistently produces clean, professional presentations without placeholder issues or formatting problems.