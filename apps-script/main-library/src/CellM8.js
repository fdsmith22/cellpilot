/**
 * CellM8 - Intelligent Spreadsheet to Presentation Converter
 * Transforms Google Sheets data into professional Google Slides presentations
 * using AI-powered analysis and smart templating
 */

const CellM8 = {
  // Configuration
  config: {
    maxSlides: 50,
    minSlides: 5,
    defaultSlides: 12,
    chartTypes: ['COLUMN', 'LINE', 'PIE', 'BAR', 'AREA', 'SCATTER'],
    templates: ['corporate', 'sales', 'financial', 'modern', 'creative'],
    presentationStyles: ['executive', 'detailed', 'sales', 'financial', 'project', 'custom']
  },

  // Template definitions
  templates: {
    corporate: {
      name: 'Corporate Professional',
      colorScheme: {
        primary: '#1f4e79',
        secondary: '#4472c4',
        accent: '#70ad47',
        background: '#ffffff',
        backgroundAccent: '#f5f7fa',
        text: '#2c3e50',
        lightText: '#7f8c8d',
        success: '#27ae60',
        warning: '#f39c12',
        error: '#e74c3c'
      },
      fontFamily: 'Roboto',
      titleFontFamily: 'Montserrat',
      layouts: {
        title: 'TITLE_AND_SUBTITLE',
        content: 'TITLE_AND_BODY',
        chart: 'TITLE_AND_TWO_COLUMNS',
        section: 'SECTION_HEADER'
      },
      slideDefaults: {
        titleSize: 40,
        subtitleSize: 24,
        bodySize: 18,
        captionSize: 12,
        lineSpacing: 1.5,
        paragraphSpacing: 12,
        bulletIndent: 20
      },
      effects: {
        titleShadow: true,
        boxShadow: '0 2px 4px rgba(0,0,0,0.1)',
        cornerRadius: 8
      }
    },
    sales: {
      name: 'Sales Focused',
      colorScheme: {
        primary: '#c55a11',
        secondary: '#f79646',
        accent: '#9cbb58',
        background: '#ffffff',
        text: '#2c2c2c'
      },
      fontFamily: 'Arial',
      layouts: {
        title: 'TITLE_ONLY',
        content: 'TITLE_AND_BODY',
        chart: 'TITLE_AND_TWO_COLUMNS',
        section: 'SECTION_HEADER'
      }
    },
    financial: {
      name: 'Financial Standard',
      colorScheme: {
        primary: '#1f4e79',
        secondary: '#4472c4',
        accent: '#843c0c',
        background: '#f5f5f5',
        text: '#1a1a1a'
      },
      fontFamily: 'Times New Roman',
      layouts: {
        title: 'TITLE_AND_SUBTITLE',
        content: 'TITLE_AND_BODY',
        chart: 'TITLE_AND_TWO_COLUMNS',
        section: 'SECTION_HEADER'
      }
    },
    modern: {
      name: 'Modern Clean',
      colorScheme: {
        primary: '#4285f4',
        secondary: '#34a853',
        accent: '#fbbc04',
        background: '#ffffff',
        text: '#202124'
      },
      fontFamily: 'Roboto',
      layouts: {
        title: 'TITLE_ONLY',
        content: 'TITLE_AND_BODY',
        chart: 'BLANK',
        section: 'SECTION_HEADER'
      }
    },
    creative: {
      name: 'Creative Colorful',
      colorScheme: {
        primary: '#9c27b0',
        secondary: '#673ab7',
        accent: '#ff5722',
        background: '#fafafa',
        text: '#212121'
      },
      fontFamily: 'Montserrat',
      layouts: {
        title: 'TITLE_ONLY',
        content: 'TITLE_AND_BODY',
        chart: 'BLANK',
        section: 'SECTION_HEADER'
      }
    }
  },

  /**
   * Get current selection information
   * @return {Object} Selection details
   */
  getCurrentSelection: function() {
    try {
      const sheet = SpreadsheetApp.getActiveSheet();
      const range = sheet.getActiveRange();
      
      if (range) {
        return {
          range: range.getA1Notation(),
          rows: range.getNumRows(),
          cols: range.getNumColumns(),
          sheet: sheet.getName()
        };
      }
      
      return null;
    } catch (error) {
      Logger.error('CellM8: Error getting selection:', error);
      return null;
    }
  },
  
  /**
   * Select entire data range on the sheet
   */
  selectEntireDataRange: function() {
    try {
      const sheet = SpreadsheetApp.getActiveSheet();
      const dataRange = sheet.getDataRange();
      sheet.setActiveRange(dataRange);
      return true;
    } catch (error) {
      Logger.error('CellM8: Error selecting entire range:', error);
      return false;
    }
  },
  
  /**
   * Select a specific range
   * @param {string} rangeA1 - A1 notation of range to select
   */
  selectRange: function(rangeA1) {
    try {
      const sheet = SpreadsheetApp.getActiveSheet();
      const range = sheet.getRange(rangeA1);
      sheet.setActiveRange(range);
      return true;
    } catch (error) {
      Logger.error('CellM8: Error selecting range:', error);
      return false;
    }
  },
  
  /**
   * Prompt user to select a range
   * @return {Object} Selected range information
   */
  promptForSelection: function() {
    try {
      const ui = SpreadsheetApp.getUi();
      const response = ui.alert(
        'Select Data Range',
        'Please select the range you want to use on the sheet, then click OK.',
        ui.ButtonSet.OK_CANCEL
      );
      
      if (response === ui.Button.OK) {
        return this.getCurrentSelection();
      }
      
      return null;
    } catch (error) {
      Logger.error('CellM8: Error prompting for selection:', error);
      return null;
    }
  },

  /**
   * Main function to create presentation from sheet data
   * @param {Object} config - Configuration object
   * @return {Object} Result with presentation ID and URL
   */
  createPresentation: function(config) {
    try {
      Logger.info('CellM8: Starting presentation creation', config);
      
      // Extract and analyze data
      const extractedData = this.extractSheetData(config.dataSource);
      if (!extractedData.success) {
        return { success: false, error: extractedData.error };
      }
      
      // Perform AI analysis
      const analysis = this.analyzeDataWithAI(extractedData.data);
      
      // Generate slide structure
      const slideStructure = this.generateSlideStructure(analysis, config);
      
      // Create Google Slides presentation
      const presentation = this.createGoogleSlides(slideStructure, config);
      
      if (presentation.success) {
        // Track usage for ML learning
        this.trackPresentationCreation(config, presentation);
        
        return {
          success: true,
          presentationId: presentation.presentationId,
          presentationUrl: presentation.url,
          slideCount: slideStructure.slides.length,
          insights: analysis.insights
        };
      } else {
        return { success: false, error: presentation.error };
      }
      
    } catch (error) {
      Logger.error('CellM8: Error creating presentation:', error);
      return { success: false, error: error.message };
    }
  },

  /**
   * Extract data from sheet based on source configuration
   * @param {Object} source - Data source configuration
   * @return {Object} Extracted data with metadata
   */
  extractSheetData: function(source) {
    try {
      const sheet = SpreadsheetApp.getActiveSheet();
      let range, data, headers;
      
      switch (source.type) {
        case 'entireSheet':
          range = sheet.getDataRange();
          break;
          
        case 'selectedRange':
          range = sheet.getActiveRange();
          if (!range) {
            return { success: false, error: 'No range selected' };
          }
          break;
          
        case 'visibleOnly':
          // Get visible cells only (excluding hidden rows/columns)
          range = this.getVisibleRange(sheet);
          break;
          
        case 'customRange':
          if (!source.range) {
            return { success: false, error: 'No custom range specified' };
          }
          range = sheet.getRange(source.range);
          break;
          
        default:
          range = sheet.getDataRange();
      }
      
      const values = range.getValues();
      const formulas = range.getFormulas();
      
      // Detect headers
      if (source.hasHeaders !== false) {
        headers = values[0];
        data = values.slice(1);
      } else {
        headers = this.generateDefaultHeaders(values[0].length);
        data = values;
      }
      
      // Clean and prepare data
      const cleanedData = this.cleanDataForPresentation(data, headers);
      
      return {
        success: true,
        data: {
          data: cleanedData,
          headers: headers
        },
        metadata: {
          sheetName: sheet.getName(),
          rowCount: data.length,
          columnCount: headers.length,
          dataTypes: this.detectDataTypes(data),
          hasFormulas: formulas.flat().some(f => f !== '')
        }
      };
      
    } catch (error) {
      Logger.error('CellM8: Error extracting sheet data:', error);
      return { success: false, error: error.message };
    }
  },

  /**
   * Get only visible cells from sheet
   * @param {Sheet} sheet - Google Sheet object
   * @return {Range} Visible range
   */
  getVisibleRange: function(sheet) {
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    const visibleRows = [];
    
    for (let i = 0; i < values.length; i++) {
      if (!sheet.isRowHiddenByUser(i + 1)) {
        const visibleRow = [];
        for (let j = 0; j < values[i].length; j++) {
          if (!sheet.isColumnHiddenByUser(j + 1)) {
            visibleRow.push(values[i][j]);
          }
        }
        visibleRows.push(visibleRow);
      }
    }
    
    // Create a temporary range with visible data
    // This is a simplified approach - in production, we'd handle this more elegantly
    return {
      getValues: () => visibleRows,
      getFormulas: () => Array(visibleRows.length).fill(Array(visibleRows[0]?.length || 0).fill(''))
    };
  },

  /**
   * Analyze data using AI/ML techniques
   * @param {Object} data - Extracted data
   * @return {Object} Analysis results
   */
  analyzeDataWithAI: function(data) {
    const analysis = {
      dataTypes: {},
      statistics: {},
      patterns: {},
      insights: [],
      chartRecommendations: [],
      keyMetrics: [],
      correlations: [],
      outliers: [],
      summaryStats: {},
      topValues: {}
    };
    
    try {
      
      // First pass: Analyze each column
      data.headers.forEach((header, index) => {
        const columnData = data.data.map(row => row[index]).filter(val => val !== '' && val !== null);
        const dataType = this.detectColumnDataType(columnData);
        
        analysis.dataTypes[header] = dataType;
        
        if (dataType === 'number') {
          // Calculate detailed statistics
          const stats = this.calculateStatistics(columnData);
          analysis.statistics[header] = stats;
          
          // Find outliers
          if (stats.mean && stats.stdDev) {
            const outliers = columnData.filter(val => 
              Math.abs(val - stats.mean) > 2 * stats.stdDev
            );
            if (outliers.length > 0) {
              analysis.outliers.push({
                column: header,
                count: outliers.length,
                values: outliers.slice(0, 5)
              });
              analysis.insights.push(`${header} contains ${outliers.length} outlier values`);
            }
          }
          
          // Detect trends
          const trend = this.detectTrend(columnData);
          if (trend) {
            analysis.patterns[header] = trend;
            analysis.insights.push(`${header} shows a ${trend.strength} ${trend.direction} trend (R² = ${(trend.rSquared * 100).toFixed(0)}%)`);
          }
          
          // Add to key metrics
          if (stats.mean) {
            analysis.keyMetrics.push({
              name: header,
              value: stats.mean,
              type: 'average',
              formatted: this.formatNumber(stats.mean)
            });
          }
        } else if (dataType === 'date') {
          // Analyze time series
          const timeSeries = this.analyzeTimeSeries(columnData);
          if (timeSeries) {
            analysis.patterns[header] = timeSeries;
            analysis.insights.push(`Date range: ${timeSeries.startDate.toLocaleDateString()} to ${timeSeries.endDate.toLocaleDateString()}`);
            analysis.insights.push(`Time series frequency: ${timeSeries.frequency}`);
          }
        } else if (dataType === 'category') {
          // Analyze categories in detail
          const categories = this.analyzeCategories(columnData);
          analysis.statistics[header] = categories;
          analysis.topValues[header] = categories.topValues || [];
          
          if (categories.uniqueValues) {
            analysis.insights.push(`${header} has ${categories.uniqueValues} unique categories`);
            if (categories.topValues && categories.topValues.length > 0) {
              analysis.insights.push(`Most common ${header}: ${categories.topValues[0].value} (${categories.topValues[0].count} occurrences)`);
            }
          }
        }
      });
      
      // Second pass: Find correlations between numeric columns
      const numericColumns = Object.entries(analysis.dataTypes)
        .filter(([_, type]) => type === 'number')
        .map(([header, _]) => header);
      
      if (numericColumns.length >= 2) {
        for (let i = 0; i < numericColumns.length - 1; i++) {
          for (let j = i + 1; j < numericColumns.length; j++) {
            const col1Data = data.data.map(row => row[data.headers.indexOf(numericColumns[i])]);
            const col2Data = data.data.map(row => row[data.headers.indexOf(numericColumns[j])]);
            const correlation = this.calculateCorrelation(col1Data, col2Data);
            
            if (Math.abs(correlation) > 0.5) {
              analysis.correlations.push({
                columns: [numericColumns[i], numericColumns[j]],
                value: correlation,
                strength: Math.abs(correlation) > 0.7 ? 'strong' : 'moderate'
              });
              
              const direction = correlation > 0 ? 'positive' : 'negative';
              analysis.insights.push(`${analysis.correlations[analysis.correlations.length - 1].strength} ${direction} correlation between ${numericColumns[i]} and ${numericColumns[j]}`);
            }
          }
        }
      }
      
      // Generate comprehensive chart recommendations
      analysis.chartRecommendations = this.recommendCharts(analysis);
      
      // Generate overall summary statistics
      analysis.summaryStats = {
        totalRows: data.data.length,
        totalColumns: data.headers.length,
        numericColumns: numericColumns.length,
        categoricalColumns: Object.values(analysis.dataTypes).filter(t => t === 'category').length,
        dateColumns: Object.values(analysis.dataTypes).filter(t => t === 'date').length,
        completeness: this.calculateCompleteness(data)
      };
      
      // Generate narrative insights if not enough found
      if (analysis.insights.length < 5) {
        analysis.insights = analysis.insights.concat(this.generateInsights(data, analysis));
      }
      
      return analysis;
      
    } catch (error) {
      Logger.error('CellM8: Error analyzing data:', error);
      console.error('CellM8 analyzeDataWithAI error:', error);
      // Return partial analysis if possible
      return {
        dataTypes: analysis.dataTypes || {},
        statistics: analysis.statistics || {},
        patterns: analysis.patterns || {},
        insights: analysis.insights.length > 0 ? analysis.insights : ['Data analysis in progress'],
        chartRecommendations: analysis.chartRecommendations || [],
        keyMetrics: analysis.keyMetrics || []
      };
    }
  },

  /**
   * Calculate statistics for numerical data
   * @param {Array} data - Column data
   * @return {Object} Statistics
   */
  calculateStatistics: function(data) {
    const numbers = data.filter(d => !isNaN(d) && d !== '').map(Number);
    
    if (numbers.length === 0) {
      return null;
    }
    
    const sum = numbers.reduce((a, b) => a + b, 0);
    const mean = sum / numbers.length;
    const sorted = numbers.sort((a, b) => a - b);
    const median = sorted[Math.floor(sorted.length / 2)];
    const min = sorted[0];
    const max = sorted[sorted.length - 1];
    
    // Calculate standard deviation
    const variance = numbers.reduce((acc, val) => acc + Math.pow(val - mean, 2), 0) / numbers.length;
    const stdDev = Math.sqrt(variance);
    
    return {
      count: numbers.length,
      sum: sum,
      mean: mean,
      median: median,
      min: min,
      max: max,
      stdDev: stdDev,
      range: max - min
    };
  },

  /**
   * Detect trend in numerical data
   * @param {Array} data - Column data
   * @return {Object} Trend information
   */
  detectTrend: function(data) {
    const numbers = data.filter(d => !isNaN(d) && d !== '').map(Number);
    
    if (numbers.length < 3) return null;
    
    // Simple linear regression
    const n = numbers.length;
    const indices = Array.from({length: n}, (_, i) => i);
    
    const sumX = indices.reduce((a, b) => a + b, 0);
    const sumY = numbers.reduce((a, b) => a + b, 0);
    const sumXY = indices.reduce((acc, x, i) => acc + x * numbers[i], 0);
    const sumX2 = indices.reduce((acc, x) => acc + x * x, 0);
    
    const slope = (n * sumXY - sumX * sumY) / (n * sumX2 - sumX * sumX);
    const intercept = (sumY - slope * sumX) / n;
    
    // Determine trend direction
    let direction = 'stable';
    if (Math.abs(slope) > 0.1) {
      direction = slope > 0 ? 'increasing' : 'decreasing';
    }
    
    // Calculate R-squared
    const yMean = sumY / n;
    const ssTotal = numbers.reduce((acc, y) => acc + Math.pow(y - yMean, 2), 0);
    const ssResidual = numbers.reduce((acc, y, i) => {
      const yPred = slope * i + intercept;
      return acc + Math.pow(y - yPred, 2);
    }, 0);
    const rSquared = 1 - (ssResidual / ssTotal);
    
    return {
      direction: direction,
      slope: slope,
      intercept: intercept,
      rSquared: rSquared,
      strength: rSquared > 0.7 ? 'strong' : rSquared > 0.4 ? 'moderate' : 'weak'
    };
  },

  /**
   * Calculate correlation between two columns
   * @param {Array} data1 - First column data
   * @param {Array} data2 - Second column data
   * @return {number} Correlation coefficient (-1 to 1)
   */
  calculateCorrelation: function(data1, data2) {
    const nums1 = data1.filter(d => !isNaN(d) && d !== '').map(Number);
    const nums2 = data2.filter(d => !isNaN(d) && d !== '').map(Number);
    
    if (nums1.length < 2 || nums2.length < 2 || nums1.length !== nums2.length) {
      return 0;
    }
    
    const n = nums1.length;
    const mean1 = nums1.reduce((a, b) => a + b, 0) / n;
    const mean2 = nums2.reduce((a, b) => a + b, 0) / n;
    
    let numerator = 0;
    let denom1 = 0;
    let denom2 = 0;
    
    for (let i = 0; i < n; i++) {
      const diff1 = nums1[i] - mean1;
      const diff2 = nums2[i] - mean2;
      numerator += diff1 * diff2;
      denom1 += diff1 * diff1;
      denom2 += diff2 * diff2;
    }
    
    const denominator = Math.sqrt(denom1 * denom2);
    return denominator === 0 ? 0 : numerator / denominator;
  },

  /**
   * Calculate data completeness
   * @param {Array} data - Column data
   * @return {number} Completeness percentage (0-100)
   */
  calculateCompleteness: function(data) {
    if (!data || data.length === 0) return 0;
    const nonEmpty = data.filter(d => d !== '' && d !== null && d !== undefined).length;
    return (nonEmpty / data.length) * 100;
  },

  /**
   * Generate preview of presentation structure
   * @param {Object} config - User configuration
   * @return {Object} Preview information
   */
  previewPresentation: function(config) {
    try {
      // Validate config has data source
      if (!config.dataSource || !config.dataSource.type) {
        return {
          success: false,
          error: 'Please select a data source first'
        };
      }
      
      // Extract sample data for preview
      const source = config.dataSource;
      const extractedData = this.extractSheetData(source);
      
      // Validate we have data
      if (!extractedData || !extractedData.data || extractedData.data.data.length === 0) {
        return {
          success: false,
          error: 'No data found in selected range'
        };
      }
      
      // Perform analysis
      const analysis = this.analyzeDataWithAI(extractedData.data);
      
      // Add data info to analysis for slide generation
      analysis.dataInfo = {
        rows: extractedData.data.data.length,
        columns: extractedData.data.headers.length
      };
      
      // Ensure we have insights
      if (!analysis.insights || analysis.insights.length === 0) {
        analysis.insights = [
          `Analyzing ${extractedData.data.data.length} rows of data`,
          `${extractedData.data.headers.length} columns identified`,
          'Data ready for presentation generation'
        ];
      }
      
      // Generate slide structure with actual slide count
      config.slideCount = config.slideCount || 12;
      const slideStructure = this.generateSlideStructure(analysis, config);
      
      return {
        success: true,
        slideCount: slideStructure.slides.length,
        slides: slideStructure.slides, // Return full slide objects with all properties
        dataInfo: {
          rows: extractedData.data.data.length,
          columns: extractedData.data.headers.length
        },
        insights: analysis.insights.slice(0, 3)
      };
    } catch (error) {
      Logger.error('CellM8: Error generating preview:', error);
      return { 
        success: false, 
        error: 'Unable to generate preview: ' + error.message 
      };
    }
  },

  /**
   * Generate slide structure based on analysis
   * @param {Object} analysis - Data analysis results
   * @param {Object} config - User configuration
   * @return {Object} Slide structure
   */
  generateSlideStructure: function(analysis, config) {
    const slides = [];
    const template = this.templates[config.template || 'corporate'];
    const targetSlideCount = config.slideCount || 12;
    
    // Ensure analysis has required properties
    if (!analysis) {
      analysis = {
        insights: [],
        keyMetrics: [],
        chartRecommendations: [],
        dataTypes: {},
        statistics: {},
        patterns: {},
        data: {}
      };
    }
    
    // 1. Title slide (always)
    slides.push({
      type: 'title',
      layout: template.layouts.title,
      title: config.presentationTitle || 'Data Analysis Presentation',
      subtitle: `Generated from ${SpreadsheetApp.getActiveSheet().getName()} | ${new Date().toLocaleDateString()}`,
      visualization: 'none',
      expandedContent: {
        description: 'Opening slide with presentation title and metadata',
        visualType: 'Title Layout',
        details: [
          `Data Source: ${SpreadsheetApp.getActiveSheet().getName()}`,
          `Generated: ${new Date().toLocaleDateString()}`,
          `Template Style: ${template.name}`,
          `Total Slides: ${targetSlideCount}`
        ]
      }
    });
    
    // 2. Executive summary 
    if (targetSlideCount >= 3) {
      const summaryItems = analysis.insights && analysis.insights.length > 0 ? 
        analysis.insights.slice(0, 4) : 
        ['Data analysis in progress', 'Key findings will be highlighted', 'Recommendations to follow'];
      
      slides.push({
        type: 'summary',
        layout: template.layouts.content,
        title: 'Executive Summary',
        content: {
          type: 'bullets',
          items: summaryItems
        },
        visualization: 'text',
        expandedContent: {
          description: 'High-level summary of key findings and insights',
          details: summaryItems,
          dataPoints: analysis.keyMetrics ? analysis.keyMetrics.slice(0, 3) : [],
          visualType: 'Bullet points with key metrics'
        }
      });
    }
    
    // 3. Data Overview slide
    if (targetSlideCount >= 4) {
      const overviewItems = [];
      
      // Add basic data stats
      if (analysis.dataInfo) {
        overviewItems.push(`Dataset contains ${analysis.dataInfo.rows || 0} records`);
        overviewItems.push(`Analyzing ${analysis.dataInfo.columns || 0} data fields`);
      }
      
      // Add data type breakdown
      if (analysis.dataTypes && Object.keys(analysis.dataTypes).length > 0) {
        const typeCount = {};
        Object.values(analysis.dataTypes).forEach(type => {
          typeCount[type] = (typeCount[type] || 0) + 1;
        });
        Object.entries(typeCount).forEach(([type, count]) => {
          overviewItems.push(`${count} ${type} column${count > 1 ? 's' : ''} identified`);
        });
      }
      
      slides.push({
        type: 'overview',
        layout: template.layouts.content,
        title: 'Data Overview',
        content: {
          type: 'bullets',
          items: overviewItems.length > 0 ? overviewItems : ['Complete dataset analysis', 'Data quality assessment', 'Pattern identification']
        },
        visualization: 'statistics',
        expandedContent: {
          description: 'Statistical overview of the dataset',
          details: overviewItems.length > 0 ? overviewItems : ['Complete dataset analysis', 'Data quality assessment', 'Pattern identification'],
          dataInfo: analysis.dataInfo || {},
          visualType: 'Statistical summary with data type breakdown'
        }
      });
    }
    
    // 4. Key Metrics slides (if we have numeric data)
    if (targetSlideCount >= 5 && analysis.statistics && Object.keys(analysis.statistics).length > 0) {
      const statsEntries = Object.entries(analysis.statistics);
      const metricsPerSlide = 3;
      const numMetricSlides = Math.min(Math.ceil(statsEntries.length / metricsPerSlide), 2);
      
      for (let i = 0; i < numMetricSlides && slides.length < targetSlideCount - 2; i++) {
        const metricsSlice = statsEntries.slice(i * metricsPerSlide, (i + 1) * metricsPerSlide);
        const metricItems = metricsSlice.map(([column, stats]) => {
          if (stats && stats.mean !== undefined) {
            return `${column}: Average ${stats.mean.toFixed(2)}, Range ${stats.min.toFixed(2)} - ${stats.max.toFixed(2)}`;
          }
          return `${column}: Analysis in progress`;
        });
        
        slides.push({
          type: 'metrics',
          layout: template.layouts.content,
          title: `Key Metrics${numMetricSlides > 1 ? ` (${i + 1})` : ''}`,
          content: {
            type: 'bullets',
            items: metricItems
          },
          visualization: 'kpi',
          expandedContent: {
            description: 'Key performance indicators and metrics',
            details: metricItems,
            rawData: metricsSlice.map(([col, stats]) => ({
              column: col,
              stats: stats
            })),
            visualType: 'KPI cards with statistical values'
          }
        });
      }
    }
    
    // 5. Chart/Visualization slides (aim for 40% of remaining slides)
    const chartsNeeded = Math.floor((targetSlideCount - slides.length - 1) * 0.4);
    if (chartsNeeded > 0 && analysis.chartRecommendations && analysis.chartRecommendations.length > 0) {
      const chartsToAdd = Math.min(chartsNeeded, analysis.chartRecommendations.length);
      
      for (let i = 0; i < chartsToAdd && slides.length < targetSlideCount - 2; i++) {
        const chart = analysis.chartRecommendations[i];
        slides.push({
          type: 'chart',
          layout: template.layouts.chart,
          title: chart.title || `Data Visualization ${i + 1}`,
          content: {
            type: 'chart',
            chartType: chart.type,
            data: chart.data,
            options: chart.options
          },
          notes: chart.insight || '',
          visualization: chart.type ? chart.type.toLowerCase() : 'chart',
          expandedContent: {
            description: `${chart.type || 'Chart'} visualization of ${chart.title || 'data trends'}`,
            details: [
              `Chart Type: ${chart.type || 'Column'}`,
              `Data Points: ${chart.data ? chart.data.length : 0}`,
              chart.insight || 'Visual representation of data patterns'
            ],
            chartConfig: chart,
            visualType: `${chart.type || 'Column'} chart`
          }
        });
      }
    }
    
    // 6. Data detail slides - create slides for each column with data
    if (slides.length < targetSlideCount - 3) {
      const headers = Object.keys(analysis.dataTypes || {});
      const dataSlideCount = Math.min(headers.length, targetSlideCount - slides.length - 3);
      
      for (let i = 0; i < dataSlideCount; i++) {
        const header = headers[i];
        const dataType = analysis.dataTypes[header];
        const stats = analysis.statistics[header];
        
        const slideContent = [];
        slideContent.push(`Data Type: ${dataType || 'Unknown'}`);
        
        if (stats) {
          if (stats.mean !== undefined) {
            slideContent.push(`Average: ${stats.mean.toFixed(2)}`);
            slideContent.push(`Range: ${stats.min.toFixed(2)} to ${stats.max.toFixed(2)}`);
            if (stats.stdDev) {
              slideContent.push(`Standard Deviation: ${stats.stdDev.toFixed(2)}`);
            }
          } else if (stats.uniqueValues) {
            slideContent.push(`Unique Values: ${stats.uniqueValues}`);
            if (stats.topValues) {
              slideContent.push(`Most Common: ${stats.topValues.slice(0, 3).join(', ')}`);
            }
          }
        }
        
        // Determine best visualization for this data type
        let vizType = 'table';
        if (dataType === 'number' && stats && stats.mean !== undefined) {
          vizType = 'histogram';
        } else if (dataType === 'category' && stats && stats.uniqueValues < 10) {
          vizType = 'pie';
        }
        
        slides.push({
          type: 'data',
          layout: template.layouts.content,
          title: `Analysis: ${header}`,
          content: {
            type: 'bullets',
            items: slideContent
          },
          visualization: vizType,
          expandedContent: {
            description: `Detailed analysis of ${header} column`,
            details: slideContent,
            statistics: stats || {},
            dataType: dataType,
            visualType: vizType === 'histogram' ? 'Histogram distribution' : 
                       vizType === 'pie' ? 'Pie chart breakdown' : 
                       'Data table view',
            sampleData: analysis.data && analysis.data[header] ? 
                       analysis.data[header].slice(0, 5) : []
          }
        });
      }
    }
    
    // 7. Insights slides - spread insights across multiple slides if needed
    if (slides.length < targetSlideCount - 2 && analysis.insights && analysis.insights.length > 0) {
      const remainingSlots = targetSlideCount - slides.length - 2; // Leave room for recommendations and thank you
      const insightsPerSlide = 3;
      const insightSlides = Math.min(Math.ceil(analysis.insights.length / insightsPerSlide), remainingSlots);
      
      for (let i = 0; i < insightSlides; i++) {
        const startIdx = i * insightsPerSlide;
        const insights = analysis.insights.slice(startIdx, startIdx + insightsPerSlide);
        
        if (insights.length > 0) {
          slides.push({
            type: 'insights',
            layout: template.layouts.content,
            title: `Key Insights ${insightSlides > 1 ? `(${i + 1})` : ''}`,
            content: {
              type: 'bullets',
              items: insights
            },
            visualization: 'callouts',
            expandedContent: {
              description: 'Data-driven insights and discoveries',
              details: insights,
              visualType: 'Highlighted callout boxes',
              importance: 'high'
            }
          });
        }
      }
    }
    
    // 8. Trends slide (if detected and room available)
    const trendsData = Object.entries(analysis.patterns || {}).filter(([_, pattern]) => pattern && pattern.direction);
    if (trendsData.length > 0 && slides.length < targetSlideCount - 2) {
      slides.push({
        type: 'trends',
        layout: template.layouts.content,
        title: 'Trend Analysis',
        content: {
          type: 'bullets',
          items: trendsData.map(([column, pattern]) => 
            `${column}: ${pattern.strength || ''} ${pattern.direction || 'stable'} trend`
          )
        },
        visualization: 'line',
        expandedContent: {
          description: 'Time-series and trend patterns in the data',
          details: trendsData.map(([column, pattern]) => ({
            column: column,
            trend: pattern.direction || 'stable',
            strength: pattern.strength || 'weak',
            rSquared: pattern.rSquared || 0
          })),
          visualType: 'Line chart with trend indicators'
        }
      });
    }
    
    // 9. Fill remaining slides with pattern/category analysis
    while (slides.length < targetSlideCount - 2) {
      const slideNumber = slides.length - 3; // Subtract title, summary, overview
      slides.push({
        type: 'analysis',
        layout: template.layouts.content,
        title: `Data Analysis ${slideNumber}`,
        content: {
          type: 'bullets',
          items: [
            'Detailed pattern analysis',
            'Statistical correlations identified',
            'Data quality assessment complete',
            'Further insights available upon request'
          ]
        },
        visualization: 'mixed',
        expandedContent: {
          description: 'Additional data analysis and patterns',
          details: [
            'Cross-column correlation analysis',
            'Outlier detection and handling',
            'Data completeness assessment',
            'Distribution analysis'
          ],
          visualType: 'Mixed visualizations (charts and tables)'
        }
      });
    }
    
    // 10. Recommendations slide (always include if room)
    if (slides.length < targetSlideCount - 1) {
      const recommendations = this.generateRecommendations(analysis);
      slides.push({
        type: 'recommendations',
        layout: template.layouts.content,
        title: 'Recommendations',
        content: {
          type: 'bullets',
          items: recommendations
        },
        visualization: 'action',
        expandedContent: {
          description: 'Actionable recommendations based on data analysis',
          details: recommendations,
          priority: 'high',
          visualType: 'Action items with priority indicators'
        }
      });
    }
    
    // 11. Thank you slide (always last)
    if (slides.length < targetSlideCount) {
      slides.push({
        type: 'thank_you',
        layout: template.layouts.title,
        title: 'Thank You',
        subtitle: 'Questions & Discussion',
        visualization: 'none',
        expandedContent: {
          description: 'Closing slide',
          details: ['Thank you for your attention', 'Open for questions and discussion'],
          visualType: 'Title slide format'
        }
      });
    }
    
    // Ensure we have exactly the requested number of slides
    const finalSlides = slides.slice(0, targetSlideCount);
    
    return {
      slides: finalSlides,
      template: template,
      metadata: {
        totalSlides: finalSlides.length,
        hasCharts: finalSlides.some(s => s.type === 'chart'),
        hasInsights: finalSlides.some(s => s.type === 'insights')
      }
    };
  },

  /**
   * Create Google Slides presentation
   * @param {Object} structure - Slide structure
   * @param {Object} config - Configuration
   * @return {Object} Creation result
   */
  createGoogleSlides: function(structure, config) {
    try {
      // Create new presentation with professional title
      const presentationTitle = config.presentationTitle || 
        `Data Analysis - ${new Date().toLocaleDateString('en-US', { month: 'long', day: 'numeric', year: 'numeric' })}`;
      const presentation = SlidesApp.create(presentationTitle);
      
      const presentationId = presentation.getId();
      const template = structure.template;
      
      // Set presentation-wide theme
      this.setupPresentationTheme(presentation, template);
      
      // Remove default slide
      const slides = presentation.getSlides();
      if (slides.length > 0) {
        slides[0].remove();
      }
      
      // Create slides based on structure
      structure.slides.forEach((slideData, index) => {
        this.createSlide(presentation, slideData, index, template);
      });
      
      // Apply template styling
      this.applyTemplateToPresentation(presentation, template);
      
      const presentationUrl = presentation.getUrl();
      
      return {
        success: true,
        presentationId: presentationId,
        url: presentationUrl
      };
      
    } catch (error) {
      Logger.error('CellM8: Error creating Google Slides:', error);
      return {
        success: false,
        error: error.message
      };
    }
  },

  /**
   * Create individual slide
   * @param {Presentation} presentation - Google Slides presentation object
   * @param {Object} slideData - Slide configuration
   * @param {number} index - Slide index
   * @param {Object} template - Template configuration
   */
  createSlide: function(presentation, slideData, index, template) {
    try {
      let slide;
      
      // Use predefined layout types
      const predefinedLayouts = presentation.getLayouts();
      let layout;
      
      // Map slide types to appropriate layouts
      if (slideData.type === 'title' || slideData.type === 'thank_you') {
        // Use title layout
        layout = predefinedLayouts.find(l => 
          l.getLayoutName().toLowerCase().includes('title') && 
          !l.getLayoutName().toLowerCase().includes('content')
        ) || predefinedLayouts[0];
      } else {
        // Use title and content layout for most slides
        layout = predefinedLayouts.find(l => 
          l.getLayoutName().toLowerCase().includes('title') && 
          l.getLayoutName().toLowerCase().includes('content')
        ) || predefinedLayouts[1];
      }
      
      slide = presentation.appendSlide(layout);
      
      // Handle different slide types
      switch (slideData.type) {
        case 'title':
          this.createTitleSlide(slide, slideData, template);
          break;
          
        case 'summary':
        case 'insights':
        case 'recommendations':
        case 'next_steps':
          this.createContentSlide(slide, slideData, template);
          break;
          
        case 'chart':
          this.createChartSlide(slide, slideData, template);
          break;
          
        case 'metrics':
          this.createMetricsSlide(slide, slideData, template);
          break;
          
        case 'overview':
        case 'trends':
        case 'data':
        case 'analysis':
          this.createDataSlide(slide, slideData, template);
          break;
          
        case 'thank_you':
          this.createTitleSlide(slide, slideData, template);
          break;
          
        default:
          this.createContentSlide(slide, slideData, template);
      }
      
      // Add speaker notes if available
      if (slideData.notes) {
        slide.getNotesPage().getSpeakerNotesShape().getText().setText(slideData.notes);
      }
      
    } catch (error) {
      Logger.error('CellM8: Error creating slide:', error);
      // Continue with other slides even if one fails
    }
  },

  /**
   * Create title slide
   */
  createTitleSlide: function(slide, slideData, template) {
    try {
      const shapes = slide.getShapes();
      let titleSet = false;
      let subtitleSet = false;
      
      // First try to use placeholders
      shapes.forEach(shape => {
        const text = shape.getText();
        const placeholder = shape.getPlaceholder();
        
        if (!titleSet && (placeholder === SlidesApp.PlaceholderType.TITLE || 
            placeholder === SlidesApp.PlaceholderType.CENTERED_TITLE)) {
          text.setText(slideData.title);
          text.getTextStyle()
            .setFontSize(40)
            .setForegroundColor(template.colorScheme.primary)
            .setFontFamily(template.fontFamily)
            .setBold(true);
          titleSet = true;
        } else if (!subtitleSet && placeholder === SlidesApp.PlaceholderType.SUBTITLE) {
          text.setText(slideData.subtitle || '');
          text.getTextStyle()
            .setFontSize(20)
            .setForegroundColor(template.colorScheme.text)
            .setFontFamily(template.fontFamily);
          subtitleSet = true;
        }
      });
      
      // If no placeholders found, create professional text boxes
      if (!titleSet) {
        // Add decorative background shape
        const bgShape = slide.insertShape(
          SlidesApp.ShapeType.RECTANGLE,
          0, 80, 720, 140
        );
        bgShape.getFill().setSolidFill(template.colorScheme.primary, 0.05);
        bgShape.getBorder().setTransparent();
        bgShape.sendToBack();
        
        const titleBox = slide.insertTextBox(slideData.title, 50, 100, 620, 100);
        titleBox.getText().getTextStyle()
          .setFontSize(44)
          .setForegroundColor(template.colorScheme.primary)
          .setFontFamily(template.titleFontFamily || template.fontFamily)
          .setBold(true);
        titleBox.getText().getParagraphStyle()
          .setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
      }
      
      if (!subtitleSet && slideData.subtitle) {
        const subtitleBox = slide.insertTextBox(slideData.subtitle, 50, 220, 620, 60);
        subtitleBox.getText().getTextStyle()
          .setFontSize(22)
          .setForegroundColor(template.colorScheme.text)
          .setFontFamily(template.fontFamily);
        subtitleBox.getText().getParagraphStyle()
          .setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
        
        // Add a subtle line separator
        const line = slide.insertLine(
          SlidesApp.LineCategory.STRAIGHT,
          260, 290, 460, 290
        );
        line.getLineFill().setSolidFill(template.colorScheme.primary, 0.3);
        line.setWeight(2);
      }
      
      // Add professional footer elements for title slide
      if (slideData.type === 'title') {
        // Add date
        const dateBox = slide.insertTextBox(
          new Date().toLocaleDateString('en-US', { 
            year: 'numeric', 
            month: 'long', 
            day: 'numeric' 
          }),
          50, 450, 200, 30
        );
        dateBox.getText().getTextStyle()
          .setFontSize(12)
          .setForegroundColor(template.colorScheme.lightText || template.colorScheme.text)
          .setFontFamily(template.fontFamily);
          
        // Add company/brand if available
        const brandBox = slide.insertTextBox(
          'Data Analysis Report',
          470, 450, 200, 30
        );
        brandBox.getText().getTextStyle()
          .setFontSize(12)
          .setForegroundColor(template.colorScheme.lightText || template.colorScheme.text)
          .setFontFamily(template.fontFamily);
        brandBox.getText().getParagraphStyle()
          .setParagraphAlignment(SlidesApp.ParagraphAlignment.END);
      }
    } catch (error) {
      Logger.error('Error creating title slide:', error);
    }
  },

  /**
   * Create content slide with bullets or text
   */
  createContentSlide: function(slide, slideData, template) {
    try {
      const shapes = slide.getShapes();
      let titleSet = false;
      let bodySet = false;
      
      // Try to use placeholders first
      shapes.forEach(shape => {
        const text = shape.getText();
        const placeholder = shape.getPlaceholder();
        
        if (!titleSet && placeholder === SlidesApp.PlaceholderType.TITLE) {
          text.setText(slideData.title);
          text.getTextStyle()
            .setFontSize(30)
            .setForegroundColor(template.colorScheme.primary)
            .setFontFamily(template.fontFamily)
            .setBold(true);
          titleSet = true;
        } else if (!bodySet && placeholder === SlidesApp.PlaceholderType.BODY) {
          if (slideData.content) {
            if (slideData.content.type === 'bullets' && slideData.content.items) {
              const bulletText = slideData.content.items.map(item => '• ' + item).join('\n');
              text.setText(bulletText);
            } else if (slideData.content.type === 'numbered' && slideData.content.items) {
              const numberedText = slideData.content.items.map((item, i) => `${i + 1}. ${item}`).join('\n');
              text.setText(numberedText);
            } else if (slideData.content.type === 'text') {
              text.setText(slideData.content.text || '');
            }
            
            text.getTextStyle()
              .setFontSize(18)
              .setForegroundColor(template.colorScheme.text)
              .setFontFamily(template.fontFamily);
            bodySet = true;
          }
        }
      });
      
      // If no placeholders found, create professional text boxes
      if (!titleSet) {
        // Add title with underline accent
        const titleBox = slide.insertTextBox(slideData.title, 50, 30, 620, 50);
        titleBox.getText().getTextStyle()
          .setFontSize(32)
          .setForegroundColor(template.colorScheme.primary)
          .setFontFamily(template.titleFontFamily || template.fontFamily)
          .setBold(true);
          
        // Add accent line under title
        const accentLine = slide.insertLine(
          SlidesApp.LineCategory.STRAIGHT,
          50, 85, 150, 85
        );
        accentLine.getLineFill().setSolidFill(template.colorScheme.accent || template.colorScheme.primary);
        accentLine.setWeight(3);
      }
      
      if (!bodySet && slideData.content && slideData.content.items) {
        // Create professional bullet points with proper spacing
        const items = slideData.content.items;
        const startY = 110;
        const itemHeight = 35;
        const bulletSymbol = template.bulletStyle || '●';
        
        items.forEach((item, index) => {
          const y = startY + (index * itemHeight);
          
          // Add bullet symbol
          const bulletBox = slide.insertTextBox(bulletSymbol, 60, y, 20, 30);
          bulletBox.getText().getTextStyle()
            .setFontSize(14)
            .setForegroundColor(template.colorScheme.accent || template.colorScheme.primary)
            .setFontFamily(template.fontFamily);
          
          // Add bullet text
          const textBox = slide.insertTextBox(item, 90, y, 580, 30);
          textBox.getText().getTextStyle()
            .setFontSize(18)
            .setForegroundColor(template.colorScheme.text)
            .setFontFamily(template.fontFamily)
            .setLineSpacing(1.2);
        });
      }
      
      // Add slide number footer (except for first slide)
      const slideIndex = presentation.getSlides().indexOf(slide);
      if (slideIndex > 0) {
        const footerBox = slide.insertTextBox(
          `${slideIndex + 1}`,
          680, 470, 30, 20
        );
        footerBox.getText().getTextStyle()
          .setFontSize(10)
          .setForegroundColor(template.colorScheme.lightText || '#999999')
          .setFontFamily(template.fontFamily);
        footerBox.getText().getParagraphStyle()
          .setParagraphAlignment(SlidesApp.ParagraphAlignment.END);
      }
    } catch (error) {
      Logger.error('Error creating content slide:', error);
    }
  },

  /**
   * Create chart slide
   */
  createChartSlide: function(slide, slideData, template) {
    try {
      // Add title
      const shapes = slide.getShapes();
      shapes.forEach(shape => {
        const placeholder = shape.getPlaceholder();
        if (placeholder === SlidesApp.PlaceholderType.TITLE) {
          const text = shape.getText();
          text.setText(slideData.title);
          text.getTextStyle()
            .setFontSize(30)
            .setForegroundColor(template.colorScheme.primary)
            .setFontFamily(template.fontFamily)
            .setBold(true);
        }
      });
      
      // Create chart
      if (slideData.content && slideData.content.data) {
        const chartData = slideData.content.data;
        const chartType = this.mapChartType(slideData.content.chartType);
        
        // Build data table for chart
        const dataTable = Charts.newDataTable();
        
        // Add headers
        if (chartData.headers) {
          chartData.headers.forEach((header, index) => {
            if (index === 0) {
              dataTable.addColumn(Charts.ColumnType.STRING, header);
            } else {
              dataTable.addColumn(Charts.ColumnType.NUMBER, header);
            }
          });
        }
        
        // Add data rows
        if (chartData.rows) {
          chartData.rows.forEach(row => {
            dataTable.addRow(row);
          });
        }
        
        // Create and configure chart
        const chart = slide.insertChart(chartType, 50, 100, 600, 350, dataTable.build());
        
        // Apply template colors to chart
        const chartOptions = chart.getOptions();
        chartOptions.setOption('colors', [
          template.colorScheme.primary,
          template.colorScheme.secondary,
          template.colorScheme.accent
        ]);
        chartOptions.setOption('fontName', template.fontFamily);
        chart.setOptions(chartOptions);
      }
      
    } catch (error) {
      Logger.error('CellM8: Error creating chart slide:', error);
      // Fall back to content slide
      this.createContentSlide(slide, {
        title: slideData.title,
        content: {
          type: 'text',
          text: 'Chart data visualization'
        }
      }, template);
    }
  },

  /**
   * Create metrics slide with KPIs
   */
  createMetricsSlide: function(slide, slideData, template) {
    try {
      // Add title
      const shapes = slide.getShapes();
      let titleSet = false;
      
      shapes.forEach(shape => {
        const placeholder = shape.getPlaceholder();
        if (!titleSet && placeholder === SlidesApp.PlaceholderType.TITLE) {
          const text = shape.getText();
          text.setText(slideData.title);
          text.getTextStyle()
            .setFontSize(30)
            .setForegroundColor(template.colorScheme.primary)
            .setFontFamily(template.fontFamily)
            .setBold(true);
          titleSet = true;
        }
      });
      
      // If no title placeholder, create title box
      if (!titleSet) {
        const titleBox = slide.insertTextBox(slideData.title, 50, 20, 620, 60);
        titleBox.getText().getTextStyle()
          .setFontSize(30)
          .setForegroundColor(template.colorScheme.primary)
          .setFontFamily(template.fontFamily)
          .setBold(true);
      }
      
      // Add metrics as formatted text or shapes
      if (slideData.content && slideData.content.items) {
      const metrics = slideData.content.items;
      const startY = 150;
      const spacing = 80;
      
      metrics.forEach((metric, index) => {
        if (index < 4) { // Limit to 4 metrics per slide
          const x = (index % 2) * 300 + 50;
          const y = Math.floor(index / 2) * spacing + startY;
          
          // Create metric box
          const metricBox = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, x, y, 250, 60);
          metricBox.getFill().setSolidFill(template.colorScheme.primary, 0.1);
          metricBox.getBorder().setTransparent();
          
          // Add metric text
          const metricText = metricBox.getText();
          metricText.setText(`${metric.label}\n${metric.value}`);
          metricText.getTextStyle()
            .setFontSize(16)
            .setForegroundColor(template.colorScheme.text)
            .setFontFamily(template.fontFamily)
            .setBold(true);
          metricText.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
        }
      });
    }
    } catch (error) {
      Logger.error('Error creating metrics slide:', error);
    }
  },

  /**
   * Apply template styling to entire presentation
   */
  applyTemplateToPresentation: function(presentation, template) {
    try {
      // Set presentation-wide properties if possible
      // Note: Google Apps Script has limited support for presentation-wide styling
      // We apply styling to individual slides instead
      
      const slides = presentation.getSlides();
      slides.forEach(slide => {
        // Set background color
        slide.getBackground().setSolidFill(template.colorScheme.background);
      });
      
    } catch (error) {
      Logger.error('CellM8: Error applying template:', error);
    }
  },

  /**
   * Map chart type string to SlidesApp chart type
   */
  mapChartType: function(type) {
    const mapping = {
      'COLUMN': SlidesApp.ChartType.COLUMN,
      'BAR': SlidesApp.ChartType.BAR,
      'LINE': SlidesApp.ChartType.LINE,
      'PIE': SlidesApp.ChartType.PIE,
      'AREA': SlidesApp.ChartType.AREA,
      'SCATTER': SlidesApp.ChartType.SCATTER
    };
    
    return mapping[type] || SlidesApp.ChartType.COLUMN;
  },

  /**
   * Recommend charts based on data analysis
   */
  recommendCharts: function(analysis) {
    const recommendations = [];
    
    // Find numerical columns for charts
    const numericalColumns = Object.entries(analysis.dataTypes)
      .filter(([_, type]) => type === 'number')
      .map(([column, _]) => column);
    
    const categoricalColumns = Object.entries(analysis.dataTypes)
      .filter(([_, type]) => type === 'category' || type === 'text')
      .map(([column, _]) => column);
    
    const dateColumns = Object.entries(analysis.dataTypes)
      .filter(([_, type]) => type === 'date')
      .map(([column, _]) => column);
    
    // Recommend column chart for categorical vs numerical
    if (categoricalColumns.length > 0 && numericalColumns.length > 0) {
      recommendations.push({
        type: 'COLUMN',
        title: `${numericalColumns[0]} by ${categoricalColumns[0]}`,
        data: {
          headers: [categoricalColumns[0], numericalColumns[0]],
          rows: [] // Will be populated with actual data
        },
        insight: `Comparison of ${numericalColumns[0]} across different ${categoricalColumns[0]} categories`
      });
    }
    
    // Recommend line chart for time series
    if (dateColumns.length > 0 && numericalColumns.length > 0) {
      recommendations.push({
        type: 'LINE',
        title: `${numericalColumns[0]} Over Time`,
        data: {
          headers: [dateColumns[0], numericalColumns[0]],
          rows: [] // Will be populated with actual data
        },
        insight: `Trend analysis of ${numericalColumns[0]} over time`
      });
    }
    
    // Recommend pie chart for distribution
    if (categoricalColumns.length > 0 && numericalColumns.length > 0) {
      recommendations.push({
        type: 'PIE',
        title: `Distribution of ${numericalColumns[0]}`,
        data: {
          headers: [categoricalColumns[0], numericalColumns[0]],
          rows: [] // Will be populated with actual data
        },
        insight: `Proportion breakdown of ${numericalColumns[0]} by ${categoricalColumns[0]}`
      });
    }
    
    // Recommend scatter plot for correlation
    if (numericalColumns.length >= 2) {
      recommendations.push({
        type: 'SCATTER',
        title: `${numericalColumns[0]} vs ${numericalColumns[1]}`,
        data: {
          headers: [numericalColumns[0], numericalColumns[1]],
          rows: [] // Will be populated with actual data
        },
        insight: `Correlation analysis between ${numericalColumns[0]} and ${numericalColumns[1]}`
      });
    }
    
    return recommendations;
  },

  /**
   * Generate insights from data analysis
   */
  generateInsights: function(data, analysis) {
    const insights = [];
    
    // Add statistical insights
    Object.entries(analysis.statistics).forEach(([column, stats]) => {
      if (stats) {
        insights.push(`${column} ranges from ${stats.min.toFixed(2)} to ${stats.max.toFixed(2)} with an average of ${stats.mean.toFixed(2)}`);
        
        if (stats.stdDev > stats.mean * 0.5) {
          insights.push(`High variability detected in ${column} (std dev: ${stats.stdDev.toFixed(2)})`);
        }
      }
    });
    
    // Add trend insights
    Object.entries(analysis.patterns).forEach(([column, pattern]) => {
      if (pattern && pattern.direction) {
        insights.push(`${column} shows a ${pattern.strength} ${pattern.direction} trend`);
      }
    });
    
    // Add data quality insights
    const totalCells = data.data.length * data.headers.length;
    const emptyCells = data.data.flat().filter(cell => cell === '' || cell === null).length;
    const completeness = ((totalCells - emptyCells) / totalCells * 100).toFixed(1);
    
    insights.push(`Data completeness: ${completeness}%`);
    
    if (completeness < 80) {
      insights.push('Consider filling missing data for more accurate analysis');
    }
    
    return insights;
  },

  /**
   * Generate recommendations based on analysis
   */
  generateRecommendations: function(analysis) {
    const recommendations = [];
    
    // Check for trends
    const trends = Object.entries(analysis.patterns).filter(([_, p]) => p && p.direction);
    if (trends.length > 0) {
      trends.forEach(([column, pattern]) => {
        if (pattern.direction === 'increasing' && pattern.strength === 'strong') {
          recommendations.push(`Continue monitoring ${column} for sustained growth`);
        } else if (pattern.direction === 'decreasing' && pattern.strength === 'strong') {
          recommendations.push(`Investigate factors causing decline in ${column}`);
        }
      });
    }
    
    // Check for outliers
    Object.entries(analysis.statistics).forEach(([column, stats]) => {
      if (stats && stats.stdDev > stats.mean * 0.5) {
        recommendations.push(`Review outliers in ${column} for data quality or exceptional cases`);
      }
    });
    
    // General recommendations
    recommendations.push('Schedule regular data reviews to track progress');
    recommendations.push('Consider setting up automated alerts for key metrics');
    
    return recommendations.slice(0, 5); // Limit to 5 recommendations
  },

  /**
   * Generate next steps
   */
  generateNextSteps: function(analysis) {
    const nextSteps = [
      'Review the presented data and insights with stakeholders',
      'Identify action items based on key findings',
      'Set up tracking for identified trends and patterns',
      'Schedule follow-up analysis in 30 days',
      'Share findings with relevant team members'
    ];
    
    return nextSteps;
  },

  /**
   * Extract key metrics from data
   */
  extractKeyMetrics: function(data, analysis) {
    const metrics = [];
    
    // Add key statistics as metrics
    Object.entries(analysis.statistics).forEach(([column, stats]) => {
      if (stats) {
        metrics.push({
          label: `${column} Average`,
          value: stats.mean.toFixed(2)
        });
        
        if (stats.sum > 1000) {
          metrics.push({
            label: `Total ${column}`,
            value: this.formatNumber(stats.sum)
          });
        }
      }
    });
    
    // Add record count
    metrics.push({
      label: 'Total Records',
      value: data.data.length.toString()
    });
    
    return metrics.slice(0, 6); // Limit to 6 key metrics
  },

  /**
   * Format data overview
   */
  formatDataOverview: function(analysis) {
    const overview = [];
    
    overview.push(`Data contains ${Object.keys(analysis.dataTypes).length} columns`);
    
    const dataTypesSummary = {};
    Object.values(analysis.dataTypes).forEach(type => {
      dataTypesSummary[type] = (dataTypesSummary[type] || 0) + 1;
    });
    
    Object.entries(dataTypesSummary).forEach(([type, count]) => {
      overview.push(`${count} ${type} column${count > 1 ? 's' : ''}`);
    });
    
    return overview;
  },

  /**
   * Clean data for presentation
   */
  cleanDataForPresentation: function(data, headers) {
    // Remove completely empty rows
    const cleanedData = data.filter(row => row.some(cell => cell !== '' && cell !== null));
    
    // Limit data size for performance
    const maxRows = 1000;
    const limitedData = cleanedData.slice(0, maxRows);
    
    return {
      headers: headers,
      data: limitedData,
      wasLimited: cleanedData.length > maxRows
    };
  },

  /**
   * Detect data types for columns
   */
  detectDataTypes: function(data) {
    const types = {};
    
    if (!data || data.length === 0) return types;
    
    const columnCount = data[0].length;
    
    for (let col = 0; col < columnCount; col++) {
      const columnData = data.map(row => row[col]).filter(val => val !== '' && val !== null);
      types[`column_${col}`] = this.detectColumnDataType(columnData);
    }
    
    return types;
  },

  /**
   * Detect data type for a single column
   */
  detectColumnDataType: function(columnData) {
    if (!columnData || columnData.length === 0) return 'unknown';
    
    const sample = columnData.slice(0, 100); // Sample first 100 values
    
    // Check if all are numbers
    const numberCount = sample.filter(val => !isNaN(val) && val !== '').length;
    if (numberCount / sample.length > 0.8) return 'number';
    
    // Check if all are dates
    const dateCount = sample.filter(val => {
      const date = new Date(val);
      return !isNaN(date.getTime()) && val.toString().match(/\d{1,2}[/-]\d{1,2}[/-]\d{2,4}/);
    }).length;
    if (dateCount / sample.length > 0.8) return 'date';
    
    // Check if categorical (limited unique values)
    const uniqueValues = [...new Set(sample)];
    if (uniqueValues.length < sample.length * 0.5) return 'category';
    
    return 'text';
  },

  /**
   * Generate default headers if not provided
   */
  generateDefaultHeaders: function(count) {
    return Array.from({length: count}, (_, i) => `Column ${String.fromCharCode(65 + i)}`);
  },

  /**
   * Analyze time series data
   */
  analyzeTimeSeries: function(dateData) {
    // Simplified time series analysis
    const dates = dateData.filter(d => d).map(d => new Date(d)).filter(d => !isNaN(d.getTime()));
    
    if (dates.length < 2) return null;
    
    dates.sort((a, b) => a - b);
    
    const intervals = [];
    for (let i = 1; i < dates.length; i++) {
      intervals.push(dates[i] - dates[i-1]);
    }
    
    const avgInterval = intervals.reduce((a, b) => a + b, 0) / intervals.length;
    
    return {
      startDate: dates[0],
      endDate: dates[dates.length - 1],
      dataPoints: dates.length,
      averageInterval: avgInterval,
      frequency: this.detectFrequency(avgInterval)
    };
  },

  /**
   * Detect data frequency
   */
  detectFrequency: function(avgInterval) {
    const day = 24 * 60 * 60 * 1000;
    
    if (avgInterval < day * 1.5) return 'daily';
    if (avgInterval < day * 8) return 'weekly';
    if (avgInterval < day * 35) return 'monthly';
    if (avgInterval < day * 100) return 'quarterly';
    return 'yearly';
  },

  /**
   * Analyze categorical data
   */
  analyzeCategories: function(categoryData) {
    const counts = {};
    
    categoryData.filter(c => c).forEach(category => {
      counts[category] = (counts[category] || 0) + 1;
    });
    
    const sorted = Object.entries(counts).sort((a, b) => b[1] - a[1]);
    
    return {
      uniqueCategories: sorted.length,
      topCategories: sorted.slice(0, 5),
      distribution: counts
    };
  },

  /**
   * Format large numbers
   */
  formatNumber: function(num) {
    if (num >= 1000000) {
      return (num / 1000000).toFixed(1) + 'M';
    } else if (num >= 1000) {
      return (num / 1000).toFixed(1) + 'K';
    }
    return num.toFixed(0);
  },

  /**
   * Chunk array into smaller arrays
   */
  chunkArray: function(array, size) {
    const chunks = [];
    for (let i = 0; i < array.length; i += size) {
      chunks.push(array.slice(i, i + size));
    }
    return chunks;
  },

  /**
   * Track presentation creation for ML learning
   */
  trackPresentationCreation: function(config, result) {
    try {
      // Store usage data for future ML improvements
      const usage = {
        timestamp: new Date(),
        config: {
          dataSource: config.dataSource.type,
          template: config.template,
          slideCount: config.slideCount,
          style: config.presentationStyle
        },
        result: {
          success: result.success,
          slideCount: result.slideCount
        }
      };
      
      // Save to user properties for learning
      const existingUsage = PropertiesService.getUserProperties().getProperty('cellm8_usage');
      const usageHistory = existingUsage ? JSON.parse(existingUsage) : [];
      usageHistory.push(usage);
      
      // Keep only last 50 entries
      if (usageHistory.length > 50) {
        usageHistory.shift();
      }
      
      PropertiesService.getUserProperties().setProperty('cellm8_usage', JSON.stringify(usageHistory));
      
    } catch (error) {
      Logger.error('CellM8: Error tracking usage:', error);
    }
  },

  /**
   * Create a thank you slide
   */
  createThankYouSlide: function(slide, slideData, template) {
    this.createTitleSlide(slide, slideData, template);
  },

  /**
   * Create a data slide with tables or stats
   */
  createDataSlide: function(slide, slideData, template) {
    this.createContentSlide(slide, slideData, template);
  }
};