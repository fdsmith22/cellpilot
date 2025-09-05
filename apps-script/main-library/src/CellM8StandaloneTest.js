/**
 * CellM8 Standalone Test Functions
 * =================================
 * These functions create their own test spreadsheet
 * Can be run from Apps Script editor without active sheet
 */

/**
 * Creates a test spreadsheet and runs chart uniqueness test
 * RUN THIS FROM APPS SCRIPT EDITOR IN LIBRARY PROJECT
 */
function standaloneTestChartUniqueness() {
  console.log('=== STANDALONE CHART UNIQUENESS TEST ===');
  console.log('Creating test spreadsheet...');
  
  // Create a new test spreadsheet
  const testSpreadsheet = SpreadsheetApp.create('CellM8 Test - ' + new Date().toLocaleTimeString());
  const sheet = testSpreadsheet.getActiveSheet();
  console.log('Created test spreadsheet: ' + testSpreadsheet.getUrl());
  
  // Add test data
  const testData = [
    ['Category', 'Sales Q1', 'Sales Q2', 'Sales Q3', 'Profit'],
    ['Product A', 100, 120, 140, 25],
    ['Product B', 150, 130, 170, 35],
    ['Product C', 200, 250, 220, 45],
    ['Product D', 180, 160, 200, 40]
  ];
  
  sheet.getRange(1, 1, testData.length, testData[0].length).setValues(testData);
  console.log('Added test data: ' + testData.length + ' rows, ' + testData[0].length + ' columns');
  
  // Prepare data structure for CellM8
  const dataStructure = {
    headers: testData[0],
    data: testData.slice(1)
  };
  
  // Analyze data
  console.log('\nAnalyzing data...');
  const analysis = CellM8.analyzeDataIntelligently(dataStructure);
  console.log('Analysis complete:');
  console.log('- Numeric columns: ' + analysis.numericColumns.length);
  console.log('- Category columns: ' + analysis.categoryColumns.length);
  console.log('- Suggested visualizations: ' + analysis.suggestedVisualizations.length);
  
  // Log column details
  analysis.numericColumns.forEach(col => {
    console.log('  Numeric: ' + col.name + ' (index ' + col.index + ')');
  });
  
  // Create 3 different charts
  console.log('\n=== CREATING CHARTS ===');
  const chartSpecs = [
    { vizIndex: 0, chartType: 'column', title: 'Test Chart 0', themeName: 'executive' },
    { vizIndex: 1, chartType: 'line', title: 'Test Chart 1', themeName: 'executive' },
    { vizIndex: 2, chartType: 'pie', title: 'Test Chart 2', themeName: 'executive' }
  ];
  
  const createdChartIds = [];
  
  chartSpecs.forEach((spec, index) => {
    console.log('\n--- Chart ' + index + ' (vizIndex=' + spec.vizIndex + ') ---');
    
    try {
      // Call CellM8.createSheetsChart
      const chart = CellM8.createSheetsChart(sheet, dataStructure, analysis, spec);
      
      if (chart) {
        const chartId = chart.getId();
        console.log('✓ Created chart with ID: ' + chartId);
        console.log('  Type: ' + spec.chartType);
        console.log('  Title: ' + spec.title);
        createdChartIds.push(chartId);
      } else {
        console.log('✗ Chart creation returned null');
      }
      
      // Small delay to ensure chart is fully created
      Utilities.sleep(1000);
      
    } catch (error) {
      console.log('✗ Error creating chart: ' + error.toString());
      console.log('  Stack: ' + error.stack);
    }
  });
  
  // Verify results
  console.log('\n=== VERIFICATION ===');
  const finalCharts = sheet.getCharts();
  console.log('Total charts in sheet: ' + finalCharts.length);
  
  const actualChartIds = finalCharts.map(c => c.getId());
  const uniqueIds = [...new Set(actualChartIds)];
  
  console.log('Created chart IDs: ' + createdChartIds.join(', '));
  console.log('Actual chart IDs: ' + actualChartIds.join(', '));
  console.log('Unique chart count: ' + uniqueIds.length);
  
  const allUnique = uniqueIds.length === finalCharts.length;
  
  if (allUnique) {
    console.log('\n✓✓✓ SUCCESS: All ' + finalCharts.length + ' charts are unique!');
  } else {
    console.log('\n✗✗✗ PROBLEM: Charts are duplicated!');
    console.log('Expected ' + finalCharts.length + ' unique charts, got ' + uniqueIds.length);
  }
  
  // Test presentation creation
  console.log('\n=== TESTING PRESENTATION CREATION ===');
  
  try {
    const config = {
      title: 'Test Presentation',
      slideCount: 5,
      template: 'executive'
    };
    
    // Make the test spreadsheet active for presentation creation
    SpreadsheetApp.setActiveSpreadsheet(testSpreadsheet);
    
    console.log('Creating presentation with config:', JSON.stringify(config));
    const result = CellM8.createOptimalPresentation(config);
    
    if (result.success) {
      console.log('✓ Presentation created successfully!');
      console.log('  URL: ' + result.url);
      console.log('  Slides: ' + result.slideCount);
    } else {
      console.log('✗ Presentation creation failed: ' + result.error);
    }
  } catch (error) {
    console.log('✗ Error creating presentation: ' + error.toString());
  }
  
  // Return results
  const result = {
    success: allUnique,
    spreadsheetUrl: testSpreadsheet.getUrl(),
    totalCharts: finalCharts.length,
    uniqueCharts: uniqueIds.length,
    chartIds: actualChartIds
  };
  
  console.log('\n=== TEST COMPLETE ===');
  console.log('Test spreadsheet: ' + testSpreadsheet.getUrl());
  console.log('Result:', JSON.stringify(result, null, 2));
  
  return result;
}

/**
 * Test just the chart creation logic without sheets
 */
function testChartLogicOnly() {
  console.log('=== TESTING CHART LOGIC ===');
  
  // Simulate data
  const testData = {
    headers: ['Category', 'Value1', 'Value2', 'Value3'],
    data: [
      ['A', 10, 20, 30],
      ['B', 15, 25, 35],
      ['C', 20, 30, 40]
    ]
  };
  
  // Analyze
  const analysis = CellM8.analyzeDataIntelligently(testData);
  
  console.log('Analysis Results:');
  console.log('- Headers: ' + testData.headers.join(', '));
  console.log('- Numeric columns: ' + analysis.numericColumns.map(c => c.name).join(', '));
  console.log('- Category columns: ' + analysis.categoryColumns.map(c => c.name).join(', '));
  
  // Test column selection for different vizIndex values
  console.log('\n=== TESTING COLUMN SELECTION LOGIC ===');
  
  for (let vizIndex = 0; vizIndex < 3; vizIndex++) {
    console.log('\nvizIndex ' + vizIndex + ':');
    
    let dataColumnIndices = [];
    
    if (vizIndex === 0) {
      // First chart logic
      if (analysis.categoryColumns.length > 0) {
        dataColumnIndices = [analysis.categoryColumns[0].index];
      }
      if (analysis.numericColumns.length > 0) {
        dataColumnIndices.push(analysis.numericColumns[0].index);
      }
    } else if (vizIndex === 1) {
      // Second chart logic
      if (analysis.categoryColumns.length > 0) {
        const catIndex = Math.min(1, analysis.categoryColumns.length - 1);
        dataColumnIndices = [analysis.categoryColumns[catIndex].index];
      }
      if (analysis.numericColumns.length > 1) {
        dataColumnIndices.push(analysis.numericColumns[1].index);
      } else if (analysis.numericColumns.length > 0) {
        dataColumnIndices.push(analysis.numericColumns[0].index);
      }
    } else if (vizIndex === 2) {
      // Third chart logic
      if (analysis.categoryColumns.length > 0) {
        dataColumnIndices = [analysis.categoryColumns[0].index];
      }
      if (analysis.numericColumns.length > 0) {
        const numIndex = Math.min(2, analysis.numericColumns.length - 1);
        dataColumnIndices.push(analysis.numericColumns[numIndex].index);
      }
    }
    
    console.log('  Selected column indices: ' + dataColumnIndices.join(', '));
    
    if (dataColumnIndices.length > 0) {
      const selectedHeaders = dataColumnIndices.map(idx => testData.headers[idx]);
      console.log('  Selected headers: ' + selectedHeaders.join(', '));
    }
  }
  
  return { success: true, analysis: analysis };
}

/**
 * Web endpoint for testing
 */
function doGet(e) {
  const param = e ? e.parameter : {};
  
  if (param.test === 'cellm8') {
    const result = standaloneTestChartUniqueness();
    return ContentService.createTextOutput(JSON.stringify(result, null, 2))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  if (param.test === 'logic') {
    const result = testChartLogicOnly();
    return ContentService.createTextOutput(JSON.stringify(result, null, 2))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  return ContentService.createTextOutput('CellM8 Test Endpoint\n\nAdd ?test=cellm8 or ?test=logic to URL')
    .setMimeType(ContentService.MimeType.TEXT);
}

/**
 * Simple test that just logs without creating anything
 */
function testCellM8Logging() {
  console.log('=== CELLM8 LOGGING TEST ===');
  console.log('This test just verifies logging works');
  
  // Test that CellM8 object exists
  console.log('CellM8 object exists: ' + (typeof CellM8 !== 'undefined'));
  
  // Test that key functions exist
  const functions = [
    'createOptimalPresentation',
    'analyzeDataIntelligently', 
    'createSheetsChart',
    'buildChartSlide'
  ];
  
  functions.forEach(func => {
    const exists = typeof CellM8[func] === 'function';
    console.log('CellM8.' + func + ': ' + (exists ? '✓' : '✗'));
  });
  
  console.log('\nTest complete - check logs above');
  return { success: true };
}