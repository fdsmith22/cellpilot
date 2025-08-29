// Pivot Table Automation Assistant with ML Intelligence
const PivotTableAssistant = {
  // Core properties
  suggestions: [],
  currentAnalysis: null,
  pivotConfigurations: [],
  
  // ML properties
  mlEnabled: false,
  patternHistory: [],
  successfulPivots: [],
  userPreferences: {},
  
  /**
   * Initialize Pivot Table Assistant
   */
  initialize: function() {
    try {
      // Check ML status
      const mlStatus = PropertiesService.getUserProperties().getProperty('cellpilot_ml_enabled');
      this.mlEnabled = mlStatus === 'true';
      
      if (this.mlEnabled) {
        // Load pattern history
        const savedHistory = PropertiesService.getUserProperties().getProperty('pivot_patterns');
        if (savedHistory) {
          this.patternHistory = JSON.parse(savedHistory);
        }
        
        // Load successful pivots
        const savedPivots = PropertiesService.getUserProperties().getProperty('successful_pivots');
        if (savedPivots) {
          this.successfulPivots = JSON.parse(savedPivots);
        }
      }
      
      return true;
    } catch (error) {
      Logger.error('Error initializing Pivot Table Assistant:', error);
      return false;
    }
  },
  
  /**
   * Analyze data for pivot table opportunities
   */
  analyzeDataForPivot: function(range) {
    try {
      const sheet = SpreadsheetApp.getActiveSheet();
      const dataRange = range || sheet.getDataRange();
      const values = dataRange.getValues();
      
      if (values.length < 2) {
        return {
          success: false,
          error: 'Not enough data for pivot analysis'
        };
      }
      
      // Assume first row is headers
      const headers = values[0];
      const data = values.slice(1);
      
      // Analyze columns
      const columnAnalysis = this.analyzeColumns(headers, data);
      
      // Detect patterns
      const patterns = this.detectDataPatterns(headers, data, columnAnalysis);
      
      // Generate suggestions
      const suggestions = this.generatePivotSuggestions(columnAnalysis, patterns);
      
      // If ML enabled, enhance with learned preferences
      if (this.mlEnabled) {
        this.enhanceSuggestionsWithML(suggestions, columnAnalysis);
      }
      
      this.currentAnalysis = {
        headers: headers,
        rowCount: data.length,
        columnCount: headers.length,
        columnAnalysis: columnAnalysis,
        patterns: patterns,
        suggestions: suggestions
      };
      
      return {
        success: true,
        analysis: this.currentAnalysis
      };
      
    } catch (error) {
      Logger.error('Error analyzing data for pivot:', error);
      return {
        success: false,
        error: error.toString()
      };
    }
  },
  
  /**
   * Analyze columns for pivot suitability
   */
  analyzeColumns: function(headers, data) {
    const analysis = [];
    
    headers.forEach((header, index) => {
      const columnData = data.map(row => row[index]).filter(val => val !== '' && val !== null);
      
      const columnInfo = {
        name: header,
        index: index,
        dataType: this.detectColumnType(columnData),
        uniqueValues: [...new Set(columnData)].length,
        nullCount: data.length - columnData.length,
        suitability: {}
      };
      
      // Determine suitability for different pivot roles
      columnInfo.suitability.rowField = this.calculateRowFieldSuitability(columnInfo, columnData);
      columnInfo.suitability.columnField = this.calculateColumnFieldSuitability(columnInfo, columnData);
      columnInfo.suitability.valueField = this.calculateValueFieldSuitability(columnInfo, columnData);
      columnInfo.suitability.filterField = this.calculateFilterFieldSuitability(columnInfo, columnData);
      
      // Calculate statistics for numeric columns
      if (columnInfo.dataType === 'number' || columnInfo.dataType === 'currency') {
        const numericData = columnData.map(v => parseFloat(v)).filter(v => !isNaN(v));
        if (numericData.length > 0) {
          columnInfo.stats = {
            sum: numericData.reduce((a, b) => a + b, 0),
            avg: numericData.reduce((a, b) => a + b, 0) / numericData.length,
            min: Math.min(...numericData),
            max: Math.max(...numericData),
            count: numericData.length
          };
        }
      }
      
      analysis.push(columnInfo);
    });
    
    return analysis;
  },
  
  /**
   * Detect column data type
   */
  detectColumnType: function(columnData) {
    if (columnData.length === 0) return 'empty';
    
    const sample = columnData.slice(0, Math.min(100, columnData.length));
    
    // Check for dates
    const dateCount = sample.filter(v => {
      const date = new Date(v);
      return date instanceof Date && !isNaN(date) && date.getFullYear() > 1900;
    }).length;
    
    if (dateCount / sample.length > 0.8) return 'date';
    
    // Check for numbers/currency
    const numericCount = sample.filter(v => {
      const str = String(v);
      return /^[$£€¥]?[-]?\d+([.,]\d+)*%?$/.test(str.replace(/,/g, ''));
    }).length;
    
    if (numericCount / sample.length > 0.8) {
      const hasCurrency = sample.some(v => /^[$£€¥]/.test(String(v)));
      return hasCurrency ? 'currency' : 'number';
    }
    
    // Check for boolean
    const booleanValues = ['true', 'false', 'yes', 'no', 'y', 'n', '1', '0'];
    const booleanCount = sample.filter(v => 
      booleanValues.includes(String(v).toLowerCase())
    ).length;
    
    if (booleanCount / sample.length > 0.8) return 'boolean';
    
    // Check for categories (limited unique values)
    const uniqueRatio = [...new Set(columnData)].length / columnData.length;
    if (uniqueRatio < 0.1) return 'category';
    
    return 'text';
  },
  
  /**
   * Calculate row field suitability
   */
  calculateRowFieldSuitability: function(columnInfo, columnData) {
    // Good row fields have moderate cardinality
    const uniqueRatio = columnInfo.uniqueValues / columnData.length;
    
    if (columnInfo.dataType === 'category') return 0.95;
    if (columnInfo.dataType === 'text' && uniqueRatio < 0.3) return 0.85;
    if (columnInfo.dataType === 'date') return 0.80;
    if (columnInfo.dataType === 'boolean') return 0.75;
    if (uniqueRatio > 0.01 && uniqueRatio < 0.5) return 0.70;
    
    return 0.30;
  },
  
  /**
   * Calculate column field suitability
   */
  calculateColumnFieldSuitability: function(columnInfo, columnData) {
    // Good column fields have low cardinality
    const uniqueRatio = columnInfo.uniqueValues / columnData.length;
    
    if (columnInfo.uniqueValues <= 10 && columnInfo.uniqueValues > 1) return 0.95;
    if (columnInfo.dataType === 'boolean') return 0.90;
    if (columnInfo.dataType === 'category' && columnInfo.uniqueValues <= 20) return 0.85;
    if (columnInfo.dataType === 'date') return 0.60; // Dates can work but need grouping
    if (uniqueRatio < 0.1) return 0.70;
    
    return 0.20;
  },
  
  /**
   * Calculate value field suitability
   */
  calculateValueFieldSuitability: function(columnInfo, columnData) {
    // Good value fields are numeric
    if (columnInfo.dataType === 'number') return 0.95;
    if (columnInfo.dataType === 'currency') return 0.98;
    if (columnInfo.dataType === 'date') return 0.10; // Can count dates
    if (columnInfo.dataType === 'text') return 0.15; // Can count text
    
    return 0.05;
  },
  
  /**
   * Calculate filter field suitability
   */
  calculateFilterFieldSuitability: function(columnInfo, columnData) {
    // Good filter fields have low to moderate cardinality
    const uniqueRatio = columnInfo.uniqueValues / columnData.length;
    
    if (columnInfo.dataType === 'boolean') return 0.95;
    if (columnInfo.dataType === 'category') return 0.90;
    if (columnInfo.uniqueValues <= 50 && columnInfo.uniqueValues > 1) return 0.85;
    if (columnInfo.dataType === 'date') return 0.80;
    if (uniqueRatio < 0.2) return 0.70;
    
    return 0.40;
  },
  
  /**
   * Detect data patterns
   */
  detectDataPatterns: function(headers, data, columnAnalysis) {
    const patterns = {
      temporal: false,
      hierarchical: false,
      categorical: false,
      numerical: false,
      mixed: false
    };
    
    // Check for temporal data
    patterns.temporal = columnAnalysis.some(col => col.dataType === 'date');
    
    // Check for hierarchical patterns (e.g., Country > State > City)
    const textColumns = columnAnalysis.filter(col => col.dataType === 'text' || col.dataType === 'category');
    if (textColumns.length >= 2) {
      // Simple check: if unique values increase progressively, might be hierarchical
      const uniqueCounts = textColumns.map(col => col.uniqueValues).sort((a, b) => a - b);
      patterns.hierarchical = uniqueCounts[uniqueCounts.length - 1] > uniqueCounts[0] * 3;
    }
    
    // Check for categorical data
    patterns.categorical = columnAnalysis.some(col => col.dataType === 'category' || col.dataType === 'boolean');
    
    // Check for numerical data
    patterns.numerical = columnAnalysis.some(col => col.dataType === 'number' || col.dataType === 'currency');
    
    // Check for mixed data
    const types = [...new Set(columnAnalysis.map(col => col.dataType))];
    patterns.mixed = types.length >= 3;
    
    return patterns;
  },
  
  /**
   * Generate pivot table suggestions
   */
  generatePivotSuggestions: function(columnAnalysis, patterns) {
    const suggestions = [];
    
    // Find best columns for each role
    const bestRowFields = columnAnalysis
      .filter(col => col.suitability.rowField > 0.6)
      .sort((a, b) => b.suitability.rowField - a.suitability.rowField)
      .slice(0, 3);
    
    const bestColumnFields = columnAnalysis
      .filter(col => col.suitability.columnField > 0.6)
      .sort((a, b) => b.suitability.columnField - a.suitability.columnField)
      .slice(0, 2);
    
    const bestValueFields = columnAnalysis
      .filter(col => col.suitability.valueField > 0.6)
      .sort((a, b) => b.suitability.valueField - a.suitability.valueField)
      .slice(0, 3);
    
    // Suggestion 1: Basic summary
    if (bestRowFields.length > 0 && bestValueFields.length > 0) {
      suggestions.push({
        name: 'Basic Summary',
        description: `Summarize ${bestValueFields[0].name} by ${bestRowFields[0].name}`,
        confidence: 0.90,
        configuration: {
          rows: [bestRowFields[0].name],
          columns: [],
          values: [{
            field: bestValueFields[0].name,
            summarizeFunction: bestValueFields[0].dataType === 'currency' ? 'SUM' : 'SUM'
          }],
          filters: []
        }
      });
    }
    
    // Suggestion 2: Cross-tabulation
    if (bestRowFields.length > 0 && bestColumnFields.length > 0 && bestValueFields.length > 0) {
      suggestions.push({
        name: 'Cross-Tabulation',
        description: `Compare ${bestValueFields[0].name} across ${bestRowFields[0].name} and ${bestColumnFields[0].name}`,
        confidence: 0.85,
        configuration: {
          rows: [bestRowFields[0].name],
          columns: [bestColumnFields[0].name],
          values: [{
            field: bestValueFields[0].name,
            summarizeFunction: 'SUM'
          }],
          filters: []
        }
      });
    }
    
    // Suggestion 3: Time-based analysis
    if (patterns.temporal) {
      const dateColumn = columnAnalysis.find(col => col.dataType === 'date');
      const valueColumn = bestValueFields[0] || columnAnalysis.find(col => col.dataType === 'number');
      
      if (dateColumn && valueColumn) {
        suggestions.push({
          name: 'Time Series Analysis',
          description: `Track ${valueColumn.name} over time`,
          confidence: 0.88,
          configuration: {
            rows: [dateColumn.name],
            columns: [],
            values: [{
              field: valueColumn.name,
              summarizeFunction: valueColumn.dataType === 'number' ? 'SUM' : 'COUNT'
            }],
            filters: []
          }
        });
      }
    }
    
    // Suggestion 4: Category breakdown
    if (patterns.categorical && bestValueFields.length > 0) {
      const categoryColumns = columnAnalysis
        .filter(col => col.dataType === 'category' || col.dataType === 'boolean')
        .slice(0, 2);
      
      if (categoryColumns.length > 0) {
        suggestions.push({
          name: 'Category Breakdown',
          description: `Analyze ${bestValueFields[0].name} by categories`,
          confidence: 0.82,
          configuration: {
            rows: categoryColumns.map(col => col.name),
            columns: [],
            values: [{
              field: bestValueFields[0].name,
              summarizeFunction: 'SUM'
            }, {
              field: bestValueFields[0].name,
              summarizeFunction: 'AVERAGE'
            }],
            filters: []
          }
        });
      }
    }
    
    // Suggestion 5: Multi-metric dashboard
    if (bestValueFields.length >= 2) {
      suggestions.push({
        name: 'Multi-Metric Dashboard',
        description: 'Compare multiple metrics side by side',
        confidence: 0.78,
        configuration: {
          rows: bestRowFields.length > 0 ? [bestRowFields[0].name] : [],
          columns: [],
          values: bestValueFields.slice(0, 3).map(field => ({
            field: field.name,
            summarizeFunction: field.dataType === 'currency' ? 'SUM' : 'AVERAGE'
          })),
          filters: []
        }
      });
    }
    
    // Sort by confidence
    suggestions.sort((a, b) => b.confidence - a.confidence);
    
    return suggestions;
  },
  
  /**
   * Enhance suggestions with ML learning
   */
  enhanceSuggestionsWithML: function(suggestions, columnAnalysis) {
    // Check for similar successful pivots
    this.successfulPivots.forEach(pivot => {
      suggestions.forEach(suggestion => {
        const similarity = this.calculatePivotSimilarity(suggestion.configuration, pivot.configuration);
        if (similarity > 0.7) {
          suggestion.confidence = Math.min(1.0, suggestion.confidence * 1.1);
          suggestion.mlEnhanced = true;
          suggestion.similarTo = pivot.name;
        }
      });
    });
    
    // Apply user preferences
    if (this.userPreferences.preferredSummarize) {
      suggestions.forEach(suggestion => {
        suggestion.configuration.values.forEach(value => {
          if (this.userPreferences.preferredSummarize[value.field]) {
            value.summarizeFunction = this.userPreferences.preferredSummarize[value.field];
          }
        });
      });
    }
  },
  
  /**
   * Calculate similarity between pivot configurations
   */
  calculatePivotSimilarity: function(config1, config2) {
    let similarity = 0;
    let factors = 0;
    
    // Compare rows
    const rowOverlap = config1.rows.filter(r => config2.rows.includes(r)).length;
    similarity += rowOverlap / Math.max(config1.rows.length, config2.rows.length);
    factors++;
    
    // Compare columns
    const colOverlap = config1.columns.filter(c => config2.columns.includes(c)).length;
    similarity += colOverlap / Math.max(config1.columns.length, config2.columns.length, 1);
    factors++;
    
    // Compare values
    const val1Fields = config1.values.map(v => v.field);
    const val2Fields = config2.values.map(v => v.field);
    const valOverlap = val1Fields.filter(v => val2Fields.includes(v)).length;
    similarity += valOverlap / Math.max(val1Fields.length, val2Fields.length);
    factors++;
    
    return similarity / factors;
  },
  
  /**
   * Create pivot table from suggestion
   */
  createPivotTable: function(suggestion, destinationSheet) {
    try {
      const sourceSheet = SpreadsheetApp.getActiveSheet();
      const sourceRange = sourceSheet.getDataRange();
      
      // Get or create destination sheet
      let pivotSheet;
      if (destinationSheet) {
        pivotSheet = destinationSheet;
      } else {
        const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        pivotSheet = spreadsheet.insertSheet(`Pivot_${new Date().getTime()}`);
      }
      
      // Create pivot table
      const pivotTable = pivotSheet.newPivotTable()
        .setSourceDataRange(sourceRange);
      
      // Add row fields
      suggestion.configuration.rows.forEach((rowField, index) => {
        const columnIndex = this.getColumnIndex(sourceSheet, rowField);
        if (columnIndex !== -1) {
          const rowGroup = pivotTable.addRowGroup(columnIndex + 1);
          rowGroup.showTotals(index === suggestion.configuration.rows.length - 1);
        }
      });
      
      // Add column fields
      suggestion.configuration.columns.forEach(colField => {
        const columnIndex = this.getColumnIndex(sourceSheet, colField);
        if (columnIndex !== -1) {
          pivotTable.addColumnGroup(columnIndex + 1);
        }
      });
      
      // Add value fields
      suggestion.configuration.values.forEach(valueConfig => {
        const columnIndex = this.getColumnIndex(sourceSheet, valueConfig.field);
        if (columnIndex !== -1) {
          const pivotValue = pivotTable.addPivotValue(
            columnIndex + 1,
            this.getSummarizeFunction(valueConfig.summarizeFunction)
          );
          
          // Set display name
          pivotValue.setDisplayName(`${valueConfig.summarizeFunction} of ${valueConfig.field}`);
        }
      });
      
      // Add filters
      suggestion.configuration.filters.forEach(filterField => {
        const columnIndex = this.getColumnIndex(sourceSheet, filterField);
        if (columnIndex !== -1) {
          pivotTable.addFilter(columnIndex + 1);
        }
      });
      
      // Build the pivot table
      const anchor = pivotSheet.getRange('A1');
      pivotTable.build(anchor);
      
      // Track successful pivot if ML enabled
      if (this.mlEnabled) {
        this.trackSuccessfulPivot(suggestion);
      }
      
      return {
        success: true,
        sheetName: pivotSheet.getName(),
        message: `Pivot table "${suggestion.name}" created successfully`
      };
      
    } catch (error) {
      Logger.error('Error creating pivot table:', error);
      return {
        success: false,
        error: error.toString()
      };
    }
  },
  
  /**
   * Get column index by name
   */
  getColumnIndex: function(sheet, columnName) {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    return headers.indexOf(columnName);
  },
  
  /**
   * Get summarize function enum
   */
  getSummarizeFunction: function(functionName) {
    const functions = {
      'SUM': SpreadsheetApp.PivotTableSummarizeFunction.SUM,
      'COUNT': SpreadsheetApp.PivotTableSummarizeFunction.COUNT,
      'AVERAGE': SpreadsheetApp.PivotTableSummarizeFunction.AVERAGE,
      'MAX': SpreadsheetApp.PivotTableSummarizeFunction.MAX,
      'MIN': SpreadsheetApp.PivotTableSummarizeFunction.MIN,
      'MEDIAN': SpreadsheetApp.PivotTableSummarizeFunction.MEDIAN,
      'PRODUCT': SpreadsheetApp.PivotTableSummarizeFunction.PRODUCT,
      'STDEV': SpreadsheetApp.PivotTableSummarizeFunction.STDEV,
      'STDEVP': SpreadsheetApp.PivotTableSummarizeFunction.STDEVP,
      'VAR': SpreadsheetApp.PivotTableSummarizeFunction.VAR,
      'VARP': SpreadsheetApp.PivotTableSummarizeFunction.VARP
    };
    
    return functions[functionName] || SpreadsheetApp.PivotTableSummarizeFunction.SUM;
  },
  
  /**
   * Track successful pivot creation
   */
  trackSuccessfulPivot: function(suggestion) {
    try {
      this.successfulPivots.push({
        name: suggestion.name,
        configuration: suggestion.configuration,
        timestamp: new Date().toISOString(),
        confidence: suggestion.confidence
      });
      
      // Keep only recent successful pivots
      if (this.successfulPivots.length > 50) {
        this.successfulPivots = this.successfulPivots.slice(-50);
      }
      
      // Update user preferences
      suggestion.configuration.values.forEach(value => {
        if (!this.userPreferences.preferredSummarize) {
          this.userPreferences.preferredSummarize = {};
        }
        this.userPreferences.preferredSummarize[value.field] = value.summarizeFunction;
      });
      
      // Save to properties
      PropertiesService.getUserProperties().setProperty(
        'successful_pivots',
        JSON.stringify(this.successfulPivots)
      );
      
    } catch (error) {
      Logger.error('Error tracking successful pivot:', error);
    }
  },
  
  /**
   * Get pivot templates
   */
  getPivotTemplates: function() {
    return [
      {
        name: 'Sales Analysis',
        description: 'Analyze sales by product and region',
        requirements: ['Product', 'Region', 'Sales Amount', 'Date'],
        configuration: {
          rows: ['Product'],
          columns: ['Region'],
          values: [{
            field: 'Sales Amount',
            summarizeFunction: 'SUM'
          }],
          filters: ['Date']
        }
      },
      {
        name: 'Time Series',
        description: 'Track metrics over time periods',
        requirements: ['Date', 'Metric'],
        configuration: {
          rows: ['Date'],
          columns: [],
          values: [{
            field: 'Metric',
            summarizeFunction: 'SUM'
          }],
          filters: []
        }
      },
      {
        name: 'Category Summary',
        description: 'Summarize data by categories',
        requirements: ['Category', 'Value'],
        configuration: {
          rows: ['Category'],
          columns: [],
          values: [{
            field: 'Value',
            summarizeFunction: 'SUM'
          }, {
            field: 'Value',
            summarizeFunction: 'COUNT'
          }],
          filters: []
        }
      },
      {
        name: 'Performance Matrix',
        description: 'Compare performance across dimensions',
        requirements: ['Dimension1', 'Dimension2', 'Performance Metric'],
        configuration: {
          rows: ['Dimension1'],
          columns: ['Dimension2'],
          values: [{
            field: 'Performance Metric',
            summarizeFunction: 'AVERAGE'
          }],
          filters: []
        }
      },
      {
        name: 'Statistical Summary',
        description: 'Get statistical measures of your data',
        requirements: ['Category', 'Value'],
        configuration: {
          rows: ['Category'],
          columns: [],
          values: [{
            field: 'Value',
            summarizeFunction: 'AVERAGE'
          }, {
            field: 'Value',
            summarizeFunction: 'STDEV'
          }, {
            field: 'Value',
            summarizeFunction: 'MIN'
          }, {
            field: 'Value',
            summarizeFunction: 'MAX'
          }],
          filters: []
        }
      }
    ];
  },
  
  /**
   * Apply template to current data
   */
  applyTemplate: function(templateName) {
    const templates = this.getPivotTemplates();
    const template = templates.find(t => t.name === templateName);
    
    if (!template) {
      return {
        success: false,
        error: 'Template not found'
      };
    }
    
    // Analyze current data
    const analysis = this.analyzeDataForPivot();
    if (!analysis.success) {
      return analysis;
    }
    
    // Map template fields to actual columns
    const mapping = this.mapTemplateToColumns(template, analysis.analysis.columnAnalysis);
    
    if (!mapping.success) {
      return mapping;
    }
    
    // Create suggestion from template
    const suggestion = {
      name: template.name,
      description: template.description,
      confidence: 0.85,
      configuration: {
        rows: mapping.rows,
        columns: mapping.columns,
        values: mapping.values,
        filters: mapping.filters
      }
    };
    
    // Create the pivot table
    return this.createPivotTable(suggestion);
  },
  
  /**
   * Map template requirements to actual columns
   */
  mapTemplateToColumns: function(template, columnAnalysis) {
    const mapping = {
      rows: [],
      columns: [],
      values: [],
      filters: [],
      success: false
    };
    
    // Try to find best matches for each requirement
    template.requirements.forEach(req => {
      const bestMatch = this.findBestColumnMatch(req, columnAnalysis);
      if (bestMatch) {
        // Map to appropriate configuration
        if (template.configuration.rows.includes(req)) {
          mapping.rows.push(bestMatch.name);
        }
        if (template.configuration.columns.includes(req)) {
          mapping.columns.push(bestMatch.name);
        }
        template.configuration.values.forEach(val => {
          if (val.field === req) {
            mapping.values.push({
              field: bestMatch.name,
              summarizeFunction: val.summarizeFunction
            });
          }
        });
        if (template.configuration.filters.includes(req)) {
          mapping.filters.push(bestMatch.name);
        }
      }
    });
    
    // Check if we have minimum requirements
    if (mapping.rows.length > 0 || mapping.columns.length > 0) {
      mapping.success = true;
    }
    
    return mapping;
  },
  
  /**
   * Find best column match for requirement
   */
  findBestColumnMatch: function(requirement, columnAnalysis) {
    const reqLower = requirement.toLowerCase();
    
    // First try exact match
    let match = columnAnalysis.find(col => 
      col.name.toLowerCase() === reqLower
    );
    
    if (match) return match;
    
    // Try partial match
    match = columnAnalysis.find(col => 
      col.name.toLowerCase().includes(reqLower) ||
      reqLower.includes(col.name.toLowerCase())
    );
    
    if (match) return match;
    
    // Try type-based match
    if (reqLower.includes('date') || reqLower.includes('time')) {
      match = columnAnalysis.find(col => col.dataType === 'date');
    } else if (reqLower.includes('amount') || reqLower.includes('value') || 
               reqLower.includes('metric') || reqLower.includes('sales')) {
      match = columnAnalysis.find(col => 
        col.dataType === 'number' || col.dataType === 'currency'
      );
    } else if (reqLower.includes('category') || reqLower.includes('dimension') ||
               reqLower.includes('product') || reqLower.includes('region')) {
      match = columnAnalysis.find(col => 
        col.dataType === 'category' || col.dataType === 'text'
      );
    }
    
    return match;
  },
  
  /**
   * Get pivot statistics
   */
  getPivotStats: function() {
    return {
      totalSuggestions: this.suggestions.length,
      successfulPivots: this.successfulPivots.length,
      mlEnabled: this.mlEnabled,
      userPreferences: this.userPreferences,
      lastAnalysis: this.currentAnalysis ? {
        rowCount: this.currentAnalysis.rowCount,
        columnCount: this.currentAnalysis.columnCount,
        timestamp: new Date().toISOString()
      } : null
    };
  }
};