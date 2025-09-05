# Safe Testing Approach for CellM8

## Option 1: Test Directly in Main Library (RECOMMENDED)

Since you have access to the main library code, you can test directly there without going through the library interface:

### Steps:
1. Open your main CellPilot Apps Script project (not the test sheet)
2. In the Apps Script editor, open `CellM8Test.js` 
3. Run `testCellM8Presentation()` directly
4. This tests the actual code without library complications

### Benefits:
- Direct access to all CellM8 functions
- No library interface issues
- See all console.log output immediately
- Can set breakpoints and debug

## Option 2: Simple HTTP Testing Endpoint

Add this to your main library's Code.js:

```javascript
function doGet(e) {
  if (e.parameter.test === 'cellm8') {
    // Run test and return results
    const result = testCellM8Presentation();
    return ContentService.createTextOutput(JSON.stringify(result, null, 2))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
```

Then test via URL without UI interaction.

## Option 3: Test from Sheet's Script Editor (Current Approach)

Keep using the test-sheet-functions.js but with these changes:

```javascript
// Instead of CellPilot.CellM8.function()
// Use CellPilot.function() directly
const result = CellPilot.createOptimalPresentation(config);
```

## Option 4: Create a Test Sidebar

Create a simple HTML sidebar in your test sheet that calls the functions and displays results:

```html
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    body { font-family: Arial; padding: 10px; }
    pre { background: #f0f0f0; padding: 10px; overflow: auto; }
    button { margin: 5px; padding: 8px; }
  </style>
</head>
<body>
  <h3>CellM8 Test Panel</h3>
  
  <button onclick="testPresentation()">Test Presentation</button>
  <button onclick="testCharts()">Test Charts Only</button>
  <button onclick="clearLog()">Clear Log</button>
  
  <h4>Results:</h4>
  <pre id="log"></pre>
  
  <script>
    function log(message) {
      document.getElementById('log').textContent += message + '\n';
    }
    
    function clearLog() {
      document.getElementById('log').textContent = '';
    }
    
    function testPresentation() {
      log('Starting presentation test...');
      google.script.run
        .withSuccessHandler(result => {
          log('Success: ' + JSON.stringify(result, null, 2));
        })
        .withFailureHandler(error => {
          log('Error: ' + error);
        })
        .testCellM8Presentation();
    }
    
    function testCharts() {
      log('Starting chart test...');
      google.script.run
        .withSuccessHandler(result => {
          log('Chart test complete: ' + JSON.stringify(result, null, 2));
        })
        .withFailureHandler(error => {
          log('Error: ' + error);
        })
        .testChartCreation();
    }
  </script>
</body>
</html>
```

## Debugging the Chart Duplication Issue

To specifically debug why charts are duplicating:

1. **Add Unique Identifiers**: In CellM8.js, add timestamps to chart titles:
```javascript
const uniqueId = Date.now() + '_' + vizIndex;
chartTitle = chartTitle + ' [' + uniqueId + ']';
```

2. **Log Chart Data Ranges**: Add logging to see exact data being used:
```javascript
console.log('Chart ' + vizIndex + ' using columns: ' + 
            dataColumnIndices.join(','));
console.log('Chart ' + vizIndex + ' range: ' + 
            dataRange.getA1Notation());
```

3. **Check Chart IDs**: Log chart IDs to verify they're different:
```javascript
const chartId = chart.getId();
console.log('Created chart with ID: ' + chartId);
```

## Quick Fix Test

To quickly test if the fix works without full deployment:

1. In your main library project
2. Create a simple test function:

```javascript
function quickTestChartUniqueness() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const data = {
    headers: ['Category', 'Value1', 'Value2'],
    data: [['A', 10, 20], ['B', 15, 25], ['C', 20, 30]]
  };
  
  // Clear existing charts
  sheet.getCharts().forEach(c => sheet.removeChart(c));
  
  // Create 3 charts with different vizIndex
  for (let i = 0; i < 3; i++) {
    console.log('Creating chart ' + i);
    const spec = {
      vizIndex: i,
      chartType: i === 0 ? 'column' : i === 1 ? 'line' : 'pie',
      title: 'Chart ' + i
    };
    
    const analysis = CellM8.analyzeDataIntelligently(data);
    const chart = CellM8.createSheetsChart(sheet, data, analysis, spec);
    
    if (chart) {
      console.log('Chart ' + i + ' ID: ' + chart.getId());
    }
  }
  
  const finalCharts = sheet.getCharts();
  console.log('Total charts: ' + finalCharts.length);
  
  // Check if charts are different
  const chartIds = finalCharts.map(c => c.getId());
  const uniqueIds = [...new Set(chartIds)];
  
  if (uniqueIds.length === finalCharts.length) {
    console.log('SUCCESS: All charts are unique!');
  } else {
    console.log('PROBLEM: Some charts are duplicated');
  }
}
```

Run this directly in the Apps Script editor of your main library project for immediate feedback!