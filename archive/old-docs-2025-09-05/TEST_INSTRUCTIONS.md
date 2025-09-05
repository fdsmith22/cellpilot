# CellM8 Testing Instructions - NO DEPLOYMENT NEEDED!

## Quick Test Setup (2 minutes)

### 1. Push the code (no deployment needed)
```bash
clasp push
```

### 2. Open Apps Script Editor
- Go to your Google Sheet
- Extensions > Apps Script
- You should see the new file: `CellM8Test.js`

### 3. Run Test Functions Directly

#### Option A: Test Full Presentation Creation
1. In Apps Script editor, open `CellM8Test.js`
2. Select function: `testCellM8Presentation`
3. Click the "Run" button (▶️)
4. Grant permissions if asked
5. Check results:
   - A popup will show the presentation URL
   - View > Execution transcript (for detailed logs)

#### Option B: Test Just Chart Creation (Faster)
1. Select function: `testChartCreation`
2. Click Run
3. Watch the logs to see if charts are unique

#### Option C: Test Data Analysis
1. Select function: `testDataAnalysis`
2. Click Run
3. See what columns and visualizations are detected

## View Logs Without Deployment

### In Apps Script Editor:
- **View > Execution transcript** - Shows console.log() output with timestamps
- **View > Logs** - Shows Logger.log() output
- Both update in real-time during execution

### What to Look For in Logs:

```
CHART CLEANUP: Found X charts to remove
CREATING CHART: vizIndex=0, type=column
CHART DATA SELECTION: Starting for vizIndex=0
SELECTED COLUMNS: indices=[0,1]
SELECTED HEADERS: Category, Value1
RANGE NOTATION: A1:B100 for chart 0
CHART CREATED: ID=xyz123, inserting into slide

CREATING CHART: vizIndex=1, type=line  
SELECTED COLUMNS: indices=[0,2]       <-- Different columns!
SELECTED HEADERS: Category, Value2     <-- Different data!
RANGE NOTATION: A1:C100 for chart 1    <-- Different range!
```

## Troubleshooting

### If charts are still duplicated:
1. Run `clearAllCharts()` first to clean slate
2. Run `testChartCreation()` to see each chart being created
3. Check the SELECTED COLUMNS in logs - they should be different

### If you get permission errors:
1. Make sure you're running from the Apps Script editor, not the UI
2. Grant all requested permissions

### To test with clean data:
1. Run `testWithMinimalData()`
2. This creates a temporary sheet with simple test data
3. Automatically cleans up after

## Benefits of This Approach

✅ **No deployment needed** - Test immediately after `clasp push`  
✅ **See detailed logs** - Track exactly what's happening  
✅ **Fast iteration** - Change code, push, test in < 30 seconds  
✅ **Isolate issues** - Test specific functions independently  
✅ **Real-time debugging** - See logs as code executes  

## Next Steps

1. Run `testCellM8Presentation()` 
2. Check if charts are unique in the logs
3. If still seeing duplicates, share the log output
4. We can then pinpoint exactly where the issue is

No more waiting for deployments or clicking through the UI!