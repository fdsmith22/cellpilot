/**
 * CellM8 Test Harness
 * ====================
 * Run these functions directly from Apps Script editor to test without deployment
 * 
 * HOW TO USE:
 * 1. Open Apps Script Editor
 * 2. Select any test function below
 * 3. Click Run button
 * 4. Check Execution Log (View > Execution transcript)
 */

/**
 * Quick test - creates presentation with sample data
 * RUN THIS DIRECTLY FROM APPS SCRIPT EDITOR
 */
function testCellM8Presentation() {
  try {
    console.log('=== STARTING CELLM8 PRESENTATION TEST ===');
    
    // Get sample data from active sheet
    const sheet = SpreadsheetApp.getActiveSheet();
    const range = sheet.getDataRange();
    const values = range.getValues();
    
    if (values.length < 2) {
      console.log('ERROR: Need at least 2 rows of data (headers + 1 data row)');
      return;
    }
    
    // Create test config
    const config = {
      title: 'TEST PRESENTATION - ' + new Date().toLocaleTimeString(),
      slideCount: 5,
      template: 'executive',
      includeCharts: true,
      includeMetrics: true,
      includeInsights: true
    };
    
    console.log('Config:', JSON.stringify(config));
    console.log('Data rows:', values.length);
    console.log('Data columns:', values[0].length);
    
    // Call CellM8 directly
    const result = CellM8.createOptimalPresentation(config);
    
    console.log('=== RESULT ===');
    console.log(JSON.stringify(result, null, 2));
    
    if (result.success) {
      console.log('SUCCESS! Presentation URL:', result.url);
      console.log('Slides created:', result.slideCount);
      
      // Open the presentation
      const ui = SpreadsheetApp.getUi();
      ui.alert('Test Complete', 'Presentation created!\n\nURL: ' + result.url + '\n\nCheck the Execution Log for details.', ui.ButtonSet.OK);
    } else {
      console.log('FAILED:', result.error);
    }
    
    return result;
    
  } catch (error) {
    console.error('Test failed:', error.toString());
    console.error('Stack:', error.stack);
    return { success: false, error: error.toString() };
  }
}

/**
 * Test just the chart creation part
 */
function testChartCreation() {
  try {
    console.log('=== TESTING CHART CREATION ===');
    
    const sheet = SpreadsheetApp.getActiveSheet();
    const range = sheet.getDataRange();
    const values = range.getValues();
    
    // Prepare data structure
    const dataStructure = {
      headers: values[0],
      data: values.slice(1)
    };
    
    // Analyze the data
    console.log('Analyzing data...');
    const analysis = CellM8.analyzeDataIntelligently(dataStructure);
    console.log('Analysis complete. Found:');
    console.log('- Numeric columns:', analysis.numericColumns.length);
    console.log('- Category columns:', analysis.categoryColumns.length);
    console.log('- Suggested visualizations:', analysis.suggestedVisualizations.length);
    
    // Test creating 3 different charts
    const testSpecs = [
      { vizIndex: 0, chartType: 'column', title: 'Test Chart 1' },
      { vizIndex: 1, chartType: 'line', title: 'Test Chart 2' },
      { vizIndex: 2, chartType: 'pie', title: 'Test Chart 3' }
    ];
    
    // Clear any existing charts first
    const existingCharts = sheet.getCharts();
    console.log('Removing', existingCharts.length, 'existing charts');
    existingCharts.forEach(c => sheet.removeChart(c));
    
    // Create each chart
    testSpecs.forEach((spec, index) => {
      console.log('\n--- Creating Chart', index + 1, '---');
      console.log('Spec:', JSON.stringify(spec));
      
      try {
        const chart = CellM8.createSheetsChart(sheet, dataStructure, analysis, spec);
        if (chart) {
          console.log('✓ Chart created successfully');
          // Wait a bit before next chart
          Utilities.sleep(1000);
        } else {
          console.log('✗ Chart creation returned null');
        }
      } catch (e) {
        console.log('✗ Error creating chart:', e.toString());
      }
    });
    
    // Check final chart count
    const finalCharts = sheet.getCharts();
    console.log('\n=== FINAL RESULT ===');
    console.log('Charts in sheet:', finalCharts.length);
    
    return { success: true, chartCount: finalCharts.length };
    
  } catch (error) {
    console.error('Chart test failed:', error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Test data analysis
 */
function testDataAnalysis() {
  try {
    console.log('=== TESTING DATA ANALYSIS ===');
    
    const sheet = SpreadsheetApp.getActiveSheet();
    const range = sheet.getDataRange();
    const values = range.getValues();
    
    const dataStructure = {
      headers: values[0],
      data: values.slice(1)
    };
    
    console.log('Headers:', dataStructure.headers);
    console.log('Data rows:', dataStructure.data.length);
    
    const analysis = CellM8.analyzeDataIntelligently(dataStructure);
    
    console.log('\n=== ANALYSIS RESULTS ===');
    console.log('Total rows:', analysis.totalRows);
    console.log('Total columns:', analysis.totalColumns);
    
    console.log('\nNumeric columns:');
    analysis.numericColumns.forEach(col => {
      console.log('  -', col.name, '(index:', col.index + ')');
      if (col.stats) {
        console.log('    Min:', col.stats.min, 'Max:', col.stats.max, 'Avg:', col.stats.avg);
      }
    });
    
    console.log('\nCategory columns:');
    analysis.categoryColumns.forEach(col => {
      console.log('  -', col.name, '(index:', col.index + ')', 
                  'Unique values:', col.uniqueValues || col.uniqueCount);
    });
    
    console.log('\nSuggested visualizations:');
    analysis.suggestedVisualizations.forEach((viz, i) => {
      console.log('  ' + (i+1) + '.', viz.type, '-', viz.title);
      if (viz.reason) console.log('     Reason:', viz.reason);
    });
    
    console.log('\nKey metrics:', analysis.keyMetrics.length);
    analysis.keyMetrics.forEach(metric => {
      console.log('  -', metric.label + ':', metric.value);
    });
    
    return analysis;
    
  } catch (error) {
    console.error('Analysis test failed:', error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Test slide building without full presentation
 */
function testSlideGeneration() {
  try {
    console.log('=== TESTING SLIDE GENERATION ===');
    
    // Create a test presentation
    const presentation = SlidesApp.create('TEST - Slide Generation - ' + new Date().toLocaleTimeString());
    console.log('Created test presentation:', presentation.getUrl());
    
    // Get data
    const sheet = SpreadsheetApp.getActiveSheet();
    const range = sheet.getDataRange();
    const values = range.getValues();
    
    const dataStructure = {
      headers: values[0],
      data: values.slice(1)
    };
    
    const analysis = CellM8.analyzeDataIntelligently(dataStructure);
    const theme = CellM8.PROFESSIONAL_THEMES.executive;
    
    // Test creating different slide types
    const slides = presentation.getSlides();
    const slide = slides[0];
    
    console.log('\nTesting chart slide creation...');
    const chartSpec = {
      type: 'chart',
      vizIndex: 0,
      chartType: 'column',
      title: 'Test Chart Slide'
    };
    
    try {
      CellM8.buildChartSlide(slide, chartSpec, dataStructure, analysis, theme);
      console.log('✓ Chart slide created');
    } catch (e) {
      console.log('✗ Chart slide failed:', e.toString());
    }
    
    // Add more slide types
    const slide2 = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
    console.log('\nTesting summary slide...');
    
    try {
      const summarySpec = {
        type: 'summary',
        metrics: analysis.keyMetrics.slice(0, 4)
      };
      CellM8.buildSummarySlide(slide2, summarySpec, theme);
      console.log('✓ Summary slide created');
    } catch (e) {
      console.log('✗ Summary slide failed:', e.toString());
    }
    
    console.log('\n=== SLIDE TEST COMPLETE ===');
    console.log('Presentation URL:', presentation.getUrl());
    
    return { 
      success: true, 
      url: presentation.getUrl(),
      slideCount: presentation.getSlides().length
    };
    
  } catch (error) {
    console.error('Slide test failed:', error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * View recent logs from Logger
 */
function viewRecentLogs() {
  // This will show in the Apps Script editor's log viewer
  console.log('=== RECENT CELLM8 LOGS ===');
  console.log('To view detailed logs:');
  console.log('1. Go to Apps Script Editor');
  console.log('2. View > Execution transcript');
  console.log('3. Or use View > Logs for simple logging');
  console.log('\nNote: Logger.log() output appears in View > Logs');
  console.log('Note: console.log() output appears in View > Execution transcript');
}

/**
 * Clear all charts from the active sheet
 */
function clearAllCharts() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const charts = sheet.getCharts();
  console.log('Removing', charts.length, 'charts...');
  charts.forEach(chart => sheet.removeChart(chart));
  console.log('All charts cleared');
}

/**
 * QUICK TEST - Run this in main library to test chart uniqueness
 */
function quickTestChartUniqueness() {
  console.log('=== QUICK CHART UNIQUENESS TEST ===');
  
  const sheet = SpreadsheetApp.getActiveSheet();
  const testData = {
    headers: ['Category', 'Value1', 'Value2', 'Value3'],
    data: [
      ['Product A', 10, 20, 30],
      ['Product B', 15, 25, 35],
      ['Product C', 20, 30, 40],
      ['Product D', 25, 35, 45]
    ]
  };
  
  // Clear existing charts
  console.log('Clearing existing charts...');
  sheet.getCharts().forEach(c => sheet.removeChart(c));
  
  // Analyze data once
  console.log('Analyzing data...');
  const analysis = CellM8.analyzeDataIntelligently(testData);
  console.log('Found ' + analysis.numericColumns.length + ' numeric columns');
  console.log('Found ' + analysis.categoryColumns.length + ' category columns');
  
  // Create 3 different charts
  const chartSpecs = [
    { vizIndex: 0, chartType: 'column', title: 'Chart 0 - Column' },
    { vizIndex: 1, chartType: 'line', title: 'Chart 1 - Line' },
    { vizIndex: 2, chartType: 'pie', title: 'Chart 2 - Pie' }
  ];
  
  const createdCharts = [];
  
  chartSpecs.forEach(spec => {
    console.log('\n--- Creating chart with vizIndex=' + spec.vizIndex + ' ---');
    
    try {
      const chart = CellM8.createSheetsChart(sheet, testData, analysis, spec);
      
      if (chart) {
        const chartId = chart.getId();
        console.log('✓ Created chart with ID: ' + chartId);
        createdCharts.push({ spec: spec, id: chartId });
      } else {
        console.log('✗ Chart creation returned null');
      }
      
      // Small delay between charts
      Utilities.sleep(500);
      
    } catch (e) {
      console.log('✗ Error: ' + e.toString());
    }
  });
  
  // Verify uniqueness
  console.log('\n=== VERIFICATION ===');
  const finalCharts = sheet.getCharts();
  console.log('Total charts in sheet: ' + finalCharts.length);
  
  const chartIds = finalCharts.map(c => c.getId());
  const uniqueIds = [...new Set(chartIds)];
  
  console.log('Chart IDs: ' + chartIds.join(', '));
  console.log('Unique IDs: ' + uniqueIds.length);
  
  if (uniqueIds.length === finalCharts.length) {
    console.log('✓✓✓ SUCCESS: All charts are unique!');
  } else {
    console.log('✗✗✗ PROBLEM: Some charts are duplicated');
    console.log('Expected ' + finalCharts.length + ' unique charts, got ' + uniqueIds.length);
  }
  
  return {
    success: uniqueIds.length === finalCharts.length,
    totalCharts: finalCharts.length,
    uniqueCharts: uniqueIds.length,
    chartIds: chartIds
  };
}

/**
 * Test with minimal data to isolate issues
 */
function testWithMinimalData() {
  try {
    console.log('=== TESTING WITH MINIMAL DATA ===');
    
    // Create a simple test sheet with known data
    const testSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('Test Data ' + Date.now());
    
    // Add simple test data
    const testData = [
      ['Category', 'Value 1', 'Value 2'],
      ['A', 100, 200],
      ['B', 150, 250],
      ['C', 200, 300],
      ['D', 250, 350]
    ];
    
    testSheet.getRange(1, 1, testData.length, testData[0].length).setValues(testData);
    console.log('Created test sheet with simple data');
    
    // Now test presentation creation
    SpreadsheetApp.setActiveSheet(testSheet);
    const result = testCellM8Presentation();
    
    // Clean up test sheet
    SpreadsheetApp.getActiveSpreadsheet().deleteSheet(testSheet);
    
    return result;
    
  } catch (error) {
    console.error('Minimal data test failed:', error.toString());
    return { success: false, error: error.toString() };
  }
}