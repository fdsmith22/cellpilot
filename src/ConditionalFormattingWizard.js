// Conditional Formatting Wizard Module with ML Anomaly Detection
const ConditionalFormattingWizard = {
  // ML Enhancement Properties
  mlEnabled: false,
  anomalyDetector: null,
  anomalyHistory: [],
  confidenceThresholds: {
    outlier: 0.85,
    trend: 0.75,
    pattern: 0.80
  },
  detectedAnomalies: {},
  
  // Initialize ML anomaly detection
  initializeML: function() {
    try {
      const mlStatus = PropertiesService.getUserProperties().getProperty('cellpilot_ml_enabled');
      this.mlEnabled = mlStatus === 'true';
      
      if (this.mlEnabled) {
        // Load anomaly history from user properties
        const savedHistory = PropertiesService.getUserProperties().getProperty('anomaly_patterns');
        if (savedHistory) {
          this.anomalyHistory = JSON.parse(savedHistory);
        }
      }
      
      return this.mlEnabled;
    } catch (error) {
      console.error('Error initializing ML for anomaly detection:', error);
      return false;
    }
  },
  
  // Get current selection for formatting
  getCurrentSelectionForFormatting: function() {
    try {
      const sheet = SpreadsheetApp.getActiveSheet();
      const range = sheet.getActiveRange();
      
      if (!range) {
        return null;
      }
      
      return {
        range: range.getA1Notation(),
        sheet: sheet.getName(),
        numRows: range.getNumRows(),
        numCols: range.getNumColumns()
      };
    } catch (error) {
      console.error('Error getting selection:', error);
      return null;
    }
  },
  
  // Apply conditional formatting rule
  applyConditionalFormatting: function(rule) {
    try {
      const sheet = SpreadsheetApp.getActiveSheet();
      const range = sheet.getRange(rule.range);
      
      let formatRule = null;
      
      switch(rule.type) {
        case 'value':
          formatRule = this.createValueRule(rule, range);
          break;
          
        case 'text':
          formatRule = this.createTextRule(rule, range);
          break;
          
        case 'date':
          formatRule = this.createDateRule(rule, range);
          break;
          
        case 'duplicate':
          formatRule = this.createDuplicateRule(rule, range);
          break;
          
        case 'scale':
          formatRule = this.createColorScaleRule(rule, range);
          break;
          
        case 'databar':
          // Data bars are not directly supported in Apps Script
          // We'll simulate with gradient formatting
          formatRule = this.createDataBarRule(rule, range);
          break;
          
        case 'formula':
          formatRule = this.createFormulaRule(rule, range);
          break;
          
        case 'top':
          formatRule = this.createTopBottomRule(rule, range);
          break;
          
        default:
          return {success: false, error: 'Unknown rule type: ' + rule.type};
      }
      
      if (formatRule) {
        // Get existing rules
        const existingRules = sheet.getConditionalFormatRules();
        
        // Add new rule
        existingRules.push(formatRule);
        
        // Apply all rules
        sheet.setConditionalFormatRules(existingRules);
        
        return {
          success: true,
          message: 'Conditional formatting applied successfully',
          rulesCount: existingRules.length
        };
      } else {
        return {success: false, error: 'Failed to create formatting rule'};
      }
      
    } catch (error) {
      console.error('Error applying conditional formatting:', error);
      return {
        success: false,
        error: error.toString()
      };
    }
  },
  
  // Create value-based rule
  createValueRule: function(rule, range) {
    const builder = SpreadsheetApp.newConditionalFormatRule();
    const conditions = rule.conditions;
    
    switch(conditions.operator) {
      case 'equal':
        builder.whenNumberEqualTo(parseFloat(conditions.value1));
        break;
      case 'notEqual':
        builder.whenNumberNotEqualTo(parseFloat(conditions.value1));
        break;
      case 'greater':
        builder.whenNumberGreaterThan(parseFloat(conditions.value1));
        break;
      case 'greaterEqual':
        builder.whenNumberGreaterThanOrEqualTo(parseFloat(conditions.value1));
        break;
      case 'less':
        builder.whenNumberLessThan(parseFloat(conditions.value1));
        break;
      case 'lessEqual':
        builder.whenNumberLessThanOrEqualTo(parseFloat(conditions.value1));
        break;
      case 'between':
        builder.whenNumberBetween(parseFloat(conditions.value1), parseFloat(conditions.value2));
        break;
      case 'notBetween':
        builder.whenNumberNotBetween(parseFloat(conditions.value1), parseFloat(conditions.value2));
        break;
    }
    
    this.applyFormat(builder, rule.format);
    builder.setRanges([range]);
    
    return builder.build();
  },
  
  // Create text-based rule
  createTextRule: function(rule, range) {
    const builder = SpreadsheetApp.newConditionalFormatRule();
    const conditions = rule.conditions;
    
    switch(conditions.operator) {
      case 'contains':
        builder.whenTextContains(conditions.value);
        break;
      case 'notContains':
        builder.whenTextDoesNotContain(conditions.value);
        break;
      case 'startsWith':
        builder.whenTextStartsWith(conditions.value);
        break;
      case 'endsWith':
        builder.whenTextEndsWith(conditions.value);
        break;
      case 'exact':
        builder.whenTextEqualTo(conditions.value);
        break;
    }
    
    this.applyFormat(builder, rule.format);
    builder.setRanges([range]);
    
    return builder.build();
  },
  
  // Create date-based rule
  createDateRule: function(rule, range) {
    const builder = SpreadsheetApp.newConditionalFormatRule();
    const conditions = rule.conditions;
    
    switch(conditions.operator) {
      case 'today':
        builder.whenDateEqualTo(SpreadsheetApp.RelativeDate.TODAY);
        break;
      case 'yesterday':
        builder.whenDateEqualTo(SpreadsheetApp.RelativeDate.YESTERDAY);
        break;
      case 'tomorrow':
        builder.whenDateEqualTo(SpreadsheetApp.RelativeDate.TOMORROW);
        break;
      case 'thisWeek':
        // Use formula for this week
        const weekFormula = '=AND(A1>=TODAY()-WEEKDAY(TODAY())+1, A1<=TODAY()-WEEKDAY(TODAY())+7)';
        builder.whenFormulaSatisfied(weekFormula);
        break;
      case 'lastWeek':
        // Use formula for last week
        const lastWeekFormula = '=AND(A1>=TODAY()-WEEKDAY(TODAY())-6, A1<=TODAY()-WEEKDAY(TODAY()))';
        builder.whenFormulaSatisfied(lastWeekFormula);
        break;
      case 'thisMonth':
        // Use formula for this month
        const monthFormula = '=AND(YEAR(A1)=YEAR(TODAY()), MONTH(A1)=MONTH(TODAY()))';
        builder.whenFormulaSatisfied(monthFormula);
        break;
      case 'lastMonth':
        // Use formula for last month
        const lastMonthFormula = '=AND(YEAR(A1)=YEAR(TODAY()-DAY(TODAY())), MONTH(A1)=MONTH(TODAY()-DAY(TODAY())))';
        builder.whenFormulaSatisfied(lastMonthFormula);
        break;
      case 'between':
        if (conditions.date1 && conditions.date2) {
          builder.whenDateBetween(new Date(conditions.date1), new Date(conditions.date2));
        }
        break;
      case 'before':
        if (conditions.date1) {
          builder.whenDateBefore(new Date(conditions.date1));
        }
        break;
      case 'after':
        if (conditions.date1) {
          builder.whenDateAfter(new Date(conditions.date1));
        }
        break;
    }
    
    this.applyFormat(builder, rule.format);
    builder.setRanges([range]);
    
    return builder.build();
  },
  
  // Create duplicate rule
  createDuplicateRule: function(rule, range) {
    const builder = SpreadsheetApp.newConditionalFormatRule();
    
    // Use COUNTIF formula to detect duplicates
    const startRow = range.getRow();
    const startCol = range.getColumn();
    const colLetter = this.columnToLetter(startCol);
    
    // Formula to check for duplicates
    const formula = `=COUNTIF($${colLetter}$${startRow}:$${colLetter}$${startRow + range.getNumRows() - 1}, ${colLetter}${startRow})>1`;
    
    builder.whenFormulaSatisfied(formula);
    this.applyFormat(builder, rule.format);
    builder.setRanges([range]);
    
    return builder.build();
  },
  
  // Create color scale rule
  createColorScaleRule: function(rule, range) {
    const builder = SpreadsheetApp.newConditionalFormatRule();
    const conditions = rule.conditions;
    
    if (conditions.type === '3color') {
      // Create 3-color gradient
      builder.setGradientMinpoint(conditions.minColor)
             .setGradientMidpointWithValue(conditions.midColor, SpreadsheetApp.InterpolationType.PERCENTILE, '50')
             .setGradientMaxpoint(conditions.maxColor);
    } else {
      // Create 2-color gradient
      builder.setGradientMinpoint(conditions.minColor)
             .setGradientMaxpoint(conditions.maxColor);
    }
    
    builder.setRanges([range]);
    
    return builder.build();
  },
  
  // Create data bar rule (simulated with gradient)
  createDataBarRule: function(rule, range) {
    const builder = SpreadsheetApp.newConditionalFormatRule();
    const conditions = rule.conditions;
    
    // Simulate data bars with gradient
    const lightColor = this.lightenColor(conditions.color, 0.8);
    
    builder.setGradientMinpoint(lightColor)
           .setGradientMaxpoint(conditions.color);
    
    builder.setRanges([range]);
    
    return builder.build();
  },
  
  // Create formula-based rule
  createFormulaRule: function(rule, range) {
    const builder = SpreadsheetApp.newConditionalFormatRule();
    
    builder.whenFormulaSatisfied(rule.conditions.formula);
    this.applyFormat(builder, rule.format);
    builder.setRanges([range]);
    
    return builder.build();
  },
  
  // Create top/bottom rule
  createTopBottomRule: function(rule, range) {
    const builder = SpreadsheetApp.newConditionalFormatRule();
    const conditions = rule.conditions;
    
    // Get all values in range
    const values = range.getValues().flat().filter(v => !isNaN(v) && v !== '');
    
    if (values.length === 0) {
      throw new Error('No numeric values in range');
    }
    
    // Sort values
    values.sort((a, b) => b - a); // Descending order
    
    let threshold;
    if (conditions.type === 'percent') {
      const percentIndex = Math.floor(values.length * parseFloat(conditions.count) / 100);
      threshold = values[Math.min(percentIndex, values.length - 1)];
    } else {
      threshold = values[Math.min(parseInt(conditions.count) - 1, values.length - 1)];
    }
    
    if (conditions.operator === 'top') {
      builder.whenNumberGreaterThanOrEqualTo(threshold);
    } else {
      // For bottom values, reverse the logic
      values.sort((a, b) => a - b); // Ascending order
      if (conditions.type === 'percent') {
        const percentIndex = Math.floor(values.length * parseFloat(conditions.count) / 100);
        threshold = values[Math.min(percentIndex, values.length - 1)];
      } else {
        threshold = values[Math.min(parseInt(conditions.count) - 1, values.length - 1)];
      }
      builder.whenNumberLessThanOrEqualTo(threshold);
    }
    
    this.applyFormat(builder, rule.format);
    builder.setRanges([range]);
    
    return builder.build();
  },
  
  // Apply format to builder
  applyFormat: function(builder, format) {
    if (format.backgroundColor) {
      builder.setBackground(format.backgroundColor);
    }
    
    if (format.textColor) {
      builder.setFontColor(format.textColor);
    }
    
    if (format.bold) {
      builder.setBold(true);
    }
    
    if (format.italic) {
      builder.setItalic(true);
    }
    
    if (format.underline) {
      builder.setUnderline(true);
    }
    
    if (format.strikethrough) {
      builder.setStrikethrough(true);
    }
  },
  
  // Get existing formatting rules
  getExistingFormattingRules: function(rangeA1) {
    try {
      const sheet = SpreadsheetApp.getActiveSheet();
      const range = sheet.getRange(rangeA1);
      const allRules = sheet.getConditionalFormatRules();
      
      const relevantRules = [];
      
      allRules.forEach(rule => {
        const ruleRanges = rule.getRanges();
        for (let ruleRange of ruleRanges) {
          if (this.rangesOverlap(range, ruleRange)) {
            relevantRules.push({
              type: this.getRuleType(rule),
              description: this.getRuleDescription(rule),
              ranges: ruleRanges.map(r => r.getA1Notation())
            });
            break;
          }
        }
      });
      
      return relevantRules;
      
    } catch (error) {
      console.error('Error getting existing rules:', error);
      return [];
    }
  },
  
  // Delete a formatting rule
  deleteFormattingRule: function(rangeA1, index) {
    try {
      const sheet = SpreadsheetApp.getActiveSheet();
      const allRules = sheet.getConditionalFormatRules();
      
      // Remove the rule at index
      allRules.splice(index, 1);
      
      // Apply remaining rules
      sheet.setConditionalFormatRules(allRules);
      
      return {success: true};
      
    } catch (error) {
      console.error('Error deleting rule:', error);
      return {success: false, error: error.toString()};
    }
  },
  
  // Clear all formatting rules in range
  clearFormattingRules: function(rangeA1) {
    try {
      const sheet = SpreadsheetApp.getActiveSheet();
      const range = sheet.getRange(rangeA1);
      const allRules = sheet.getConditionalFormatRules();
      
      // Filter out rules that overlap with the range
      const remainingRules = allRules.filter(rule => {
        const ruleRanges = rule.getRanges();
        return !ruleRanges.some(ruleRange => this.rangesOverlap(range, ruleRange));
      });
      
      sheet.setConditionalFormatRules(remainingRules);
      
      return {
        success: true,
        removed: allRules.length - remainingRules.length
      };
      
    } catch (error) {
      console.error('Error clearing rules:', error);
      return {success: false, error: error.toString()};
    }
  },
  
  // Helper: Check if ranges overlap
  rangesOverlap: function(range1, range2) {
    const r1StartRow = range1.getRow();
    const r1EndRow = r1StartRow + range1.getNumRows() - 1;
    const r1StartCol = range1.getColumn();
    const r1EndCol = r1StartCol + range1.getNumColumns() - 1;
    
    const r2StartRow = range2.getRow();
    const r2EndRow = r2StartRow + range2.getNumRows() - 1;
    const r2StartCol = range2.getColumn();
    const r2EndCol = r2StartCol + range2.getNumColumns() - 1;
    
    return !(r1EndRow < r2StartRow || r2EndRow < r1StartRow ||
             r1EndCol < r2StartCol || r2EndCol < r1StartCol);
  },
  
  // Helper: Get rule type
  getRuleType: function(rule) {
    // This is simplified - Apps Script doesn't expose the exact condition type
    const booleanCondition = rule.getBooleanCondition();
    if (booleanCondition) {
      const type = booleanCondition.getCriteriaType();
      return type.toString().replace('WHEN_', '').toLowerCase();
    }
    
    const gradientCondition = rule.getGradientCondition();
    if (gradientCondition) {
      return 'gradient';
    }
    
    return 'unknown';
  },
  
  // Helper: Get rule description
  getRuleDescription: function(rule) {
    const booleanCondition = rule.getBooleanCondition();
    if (booleanCondition) {
      const type = booleanCondition.getCriteriaType();
      const values = booleanCondition.getCriteriaValues();
      
      if (values && values.length > 0) {
        return `${type.toString()}: ${values.join(', ')}`;
      }
      return type.toString();
    }
    
    const gradientCondition = rule.getGradientCondition();
    if (gradientCondition) {
      return 'Color gradient';
    }
    
    return 'Custom rule';
  },
  
  // Helper: Convert column number to letter
  columnToLetter: function(column) {
    let temp;
    let letter = '';
    while (column > 0) {
      temp = (column - 1) % 26;
      letter = String.fromCharCode(temp + 65) + letter;
      column = (column - temp - 1) / 26;
    }
    return letter;
  },
  
  // Helper: Lighten a color
  lightenColor: function(color, percent) {
    const num = parseInt(color.replace('#', ''), 16);
    const r = Math.min(255, Math.floor((num >> 16) + (255 - (num >> 16)) * percent));
    const g = Math.min(255, Math.floor(((num >> 8) & 0x00FF) + (255 - ((num >> 8) & 0x00FF)) * percent));
    const b = Math.min(255, Math.floor((num & 0x0000FF) + (255 - (num & 0x0000FF)) * percent));
    
    return '#' + ((r << 16) | (g << 8) | b).toString(16).padStart(6, '0');
  },
  
  // Create preset formatting templates
  applyPresetFormatting: function(preset, rangeA1) {
    const presets = {
      'heatmap': {
        type: 'scale',
        conditions: {
          type: '3color',
          minColor: '#FFFFFF',
          midColor: '#FFFF00',
          maxColor: '#FF0000'
        }
      },
      'progress': {
        type: 'scale',
        conditions: {
          type: '2color',
          minColor: '#FFE5E5',
          maxColor: '#00FF00'
        }
      },
      'highlight_negative': {
        type: 'value',
        conditions: {
          operator: 'less',
          value1: '0'
        },
        format: {
          backgroundColor: '#FFE5E5',
          textColor: '#FF0000'
        }
      },
      'highlight_positive': {
        type: 'value',
        conditions: {
          operator: 'greater',
          value1: '0'
        },
        format: {
          backgroundColor: '#E5FFE5',
          textColor: '#008000'
        }
      },
      'duplicate_red': {
        type: 'duplicate',
        conditions: {
          includeFirst: true
        },
        format: {
          backgroundColor: '#FFE5E5',
          textColor: '#FF0000',
          bold: true
        }
      },
      'anomaly_detection': {
        type: 'formula',
        conditions: {
          formula: '=CUSTOM_ANOMALY_CHECK(A1)'  // Will be replaced with actual anomaly detection
        },
        format: {
          backgroundColor: '#FFE5CC',
          textColor: '#FF6600',
          bold: true
        }
      }
    };
    
    const rule = presets[preset];
    if (!rule) {
      return {success: false, error: 'Preset not found'};
    }
    
    // If ML anomaly detection is requested, detect anomalies first
    if (preset === 'anomaly_detection' && this.mlEnabled) {
      const anomalyResult = this.detectAnomaliesInRange(rangeA1);
      if (anomalyResult.success) {
        return this.applyAnomalyFormatting(anomalyResult.anomalies, rangeA1);
      }
    }
    
    rule.range = rangeA1;
    return this.applyConditionalFormatting(rule);
  },
  
  // ML-powered anomaly detection
  detectAnomaliesInRange: function(rangeA1) {
    try {
      const sheet = SpreadsheetApp.getActiveSheet();
      const range = sheet.getRange(rangeA1);
      const values = range.getValues();
      const numRows = values.length;
      const numCols = values[0].length;
      
      const anomalies = [];
      
      // Process each column for anomalies
      for (let col = 0; col < numCols; col++) {
        const columnData = values.map(row => row[col]).filter(v => v !== '' && !isNaN(v));
        
        if (columnData.length > 3) {  // Need at least 4 values for meaningful detection
          const columnAnomalies = this.detectColumnAnomalies(columnData, col);
          anomalies.push(...columnAnomalies);
        }
      }
      
      // Detect cross-column pattern anomalies
      if (numCols > 1) {
        const patternAnomalies = this.detectPatternAnomalies(values);
        anomalies.push(...patternAnomalies);
      }
      
      // Store detection results
      this.detectedAnomalies[rangeA1] = {
        timestamp: new Date().toISOString(),
        anomalies: anomalies,
        confidence: this.calculateOverallConfidence(anomalies)
      };
      
      return {
        success: true,
        anomalies: anomalies,
        totalFound: anomalies.length
      };
      
    } catch (error) {
      console.error('Error detecting anomalies:', error);
      return {
        success: false,
        error: error.toString()
      };
    }
  },
  
  // Detect anomalies in a single column
  detectColumnAnomalies: function(columnData, colIndex) {
    const anomalies = [];
    
    // Calculate statistics
    const stats = this.calculateStatistics(columnData);
    
    // Z-score based outlier detection
    columnData.forEach((value, rowIndex) => {
      const zScore = Math.abs((value - stats.mean) / stats.stdDev);
      
      if (zScore > 3) {  // Standard threshold for outliers
        anomalies.push({
          type: 'outlier',
          row: rowIndex,
          col: colIndex,
          value: value,
          zScore: zScore,
          confidence: Math.min(0.99, 0.7 + (zScore - 3) * 0.1),
          reason: `Value is ${zScore.toFixed(1)} standard deviations from mean`
        });
      }
    });
    
    // IQR-based outlier detection
    const iqrOutliers = this.detectIQROutliers(columnData, colIndex);
    anomalies.push(...iqrOutliers);
    
    // Trend break detection
    const trendBreaks = this.detectTrendBreaks(columnData, colIndex);
    anomalies.push(...trendBreaks);
    
    return anomalies;
  },
  
  // Calculate basic statistics
  calculateStatistics: function(data) {
    const n = data.length;
    const mean = data.reduce((sum, val) => sum + val, 0) / n;
    const variance = data.reduce((sum, val) => sum + Math.pow(val - mean, 2), 0) / n;
    const stdDev = Math.sqrt(variance);
    
    // Calculate median
    const sorted = [...data].sort((a, b) => a - b);
    const median = n % 2 === 0 
      ? (sorted[n/2 - 1] + sorted[n/2]) / 2
      : sorted[Math.floor(n/2)];
    
    // Calculate quartiles
    const q1Index = Math.floor(n * 0.25);
    const q3Index = Math.floor(n * 0.75);
    const q1 = sorted[q1Index];
    const q3 = sorted[q3Index];
    const iqr = q3 - q1;
    
    return {
      mean: mean,
      median: median,
      stdDev: stdDev,
      min: sorted[0],
      max: sorted[n - 1],
      q1: q1,
      q3: q3,
      iqr: iqr
    };
  },
  
  // Detect outliers using IQR method
  detectIQROutliers: function(data, colIndex) {
    const anomalies = [];
    const stats = this.calculateStatistics(data);
    
    const lowerBound = stats.q1 - 1.5 * stats.iqr;
    const upperBound = stats.q3 + 1.5 * stats.iqr;
    
    data.forEach((value, rowIndex) => {
      if (value < lowerBound || value > upperBound) {
        const deviation = value < lowerBound 
          ? lowerBound - value 
          : value - upperBound;
        
        anomalies.push({
          type: 'iqr_outlier',
          row: rowIndex,
          col: colIndex,
          value: value,
          confidence: Math.min(0.95, 0.75 + deviation / stats.iqr * 0.1),
          reason: `Value outside IQR bounds [${lowerBound.toFixed(2)}, ${upperBound.toFixed(2)}]`
        });
      }
    });
    
    return anomalies;
  },
  
  // Detect trend breaks in sequential data
  detectTrendBreaks: function(data, colIndex) {
    const anomalies = [];
    
    if (data.length < 5) return anomalies;
    
    // Calculate moving averages
    const windowSize = Math.min(5, Math.floor(data.length / 3));
    const movingAvg = [];
    
    for (let i = windowSize - 1; i < data.length; i++) {
      const window = data.slice(i - windowSize + 1, i + 1);
      const avg = window.reduce((sum, val) => sum + val, 0) / windowSize;
      movingAvg.push(avg);
    }
    
    // Detect significant deviations from moving average
    for (let i = 0; i < movingAvg.length; i++) {
      const actualIndex = i + windowSize - 1;
      const actual = data[actualIndex];
      const expected = movingAvg[i];
      const deviation = Math.abs(actual - expected);
      const relativeDeviation = deviation / Math.abs(expected);
      
      if (relativeDeviation > 0.5) {  // 50% deviation from moving average
        anomalies.push({
          type: 'trend_break',
          row: actualIndex,
          col: colIndex,
          value: actual,
          expectedValue: expected,
          confidence: Math.min(0.9, 0.6 + relativeDeviation * 0.2),
          reason: `Value deviates ${(relativeDeviation * 100).toFixed(0)}% from trend`
        });
      }
    }
    
    return anomalies;
  },
  
  // Detect pattern anomalies across multiple columns
  detectPatternAnomalies: function(values) {
    const anomalies = [];
    const numRows = values.length;
    const numCols = values[0].length;
    
    // Check for row-wise anomalies
    for (let row = 0; row < numRows; row++) {
      const rowData = values[row].filter(v => v !== '' && !isNaN(v));
      
      if (rowData.length > 2) {
        // Check if row sum/pattern is anomalous compared to other rows
        const rowSum = rowData.reduce((sum, val) => sum + val, 0);
        const allRowSums = [];
        
        for (let r = 0; r < numRows; r++) {
          const rData = values[r].filter(v => v !== '' && !isNaN(v));
          if (rData.length > 0) {
            allRowSums.push(rData.reduce((sum, val) => sum + val, 0));
          }
        }
        
        if (allRowSums.length > 3) {
          const stats = this.calculateStatistics(allRowSums);
          const zScore = Math.abs((rowSum - stats.mean) / stats.stdDev);
          
          if (zScore > 2.5) {
            anomalies.push({
              type: 'row_pattern',
              row: row,
              col: -1,  // Entire row
              value: rowSum,
              confidence: Math.min(0.85, 0.65 + zScore * 0.05),
              reason: `Row sum is anomalous (z-score: ${zScore.toFixed(2)})`
            });
          }
        }
      }
    }
    
    return anomalies;
  },
  
  // Calculate overall confidence for detected anomalies
  calculateOverallConfidence: function(anomalies) {
    if (anomalies.length === 0) return 0;
    
    const totalConfidence = anomalies.reduce((sum, a) => sum + a.confidence, 0);
    return totalConfidence / anomalies.length;
  },
  
  // Apply formatting to detected anomalies
  applyAnomalyFormatting: function(anomalies, rangeA1) {
    try {
      const sheet = SpreadsheetApp.getActiveSheet();
      const range = sheet.getRange(rangeA1);
      const startRow = range.getRow();
      const startCol = range.getColumn();
      
      // Group anomalies by confidence level
      const highConfidence = anomalies.filter(a => a.confidence >= 0.85);
      const mediumConfidence = anomalies.filter(a => a.confidence >= 0.70 && a.confidence < 0.85);
      const lowConfidence = anomalies.filter(a => a.confidence < 0.70);
      
      // Apply different formatting based on confidence
      const rules = [];
      
      // High confidence - strong highlighting
      if (highConfidence.length > 0) {
        highConfidence.forEach(anomaly => {
          const cellRow = startRow + anomaly.row;
          const cellCol = anomaly.col === -1 ? startCol : startCol + anomaly.col;
          const cellRange = sheet.getRange(cellRow, cellCol);
          
          const rule = SpreadsheetApp.newConditionalFormatRule()
            .whenNumberEqualTo(anomaly.value)
            .setBackground('#FF9999')
            .setFontColor('#CC0000')
            .setBold(true)
            .setRanges([cellRange])
            .build();
          
          rules.push(rule);
        });
      }
      
      // Medium confidence - moderate highlighting
      if (mediumConfidence.length > 0) {
        mediumConfidence.forEach(anomaly => {
          const cellRow = startRow + anomaly.row;
          const cellCol = anomaly.col === -1 ? startCol : startCol + anomaly.col;
          const cellRange = sheet.getRange(cellRow, cellCol);
          
          const rule = SpreadsheetApp.newConditionalFormatRule()
            .whenNumberEqualTo(anomaly.value)
            .setBackground('#FFCC99')
            .setFontColor('#FF6600')
            .setRanges([cellRange])
            .build();
          
          rules.push(rule);
        });
      }
      
      // Add rules to sheet
      if (rules.length > 0) {
        const existingRules = sheet.getConditionalFormatRules();
        existingRules.push(...rules);
        sheet.setConditionalFormatRules(existingRules);
      }
      
      // Save anomaly detection to history
      this.saveAnomalyHistory(rangeA1, anomalies);
      
      return {
        success: true,
        message: `Applied anomaly formatting to ${anomalies.length} cells`,
        highConfidence: highConfidence.length,
        mediumConfidence: mediumConfidence.length,
        lowConfidence: lowConfidence.length
      };
      
    } catch (error) {
      console.error('Error applying anomaly formatting:', error);
      return {
        success: false,
        error: error.toString()
      };
    }
  },
  
  // Save anomaly detection history for learning
  saveAnomalyHistory: function(rangeA1, anomalies) {
    try {
      if (!this.mlEnabled) return;
      
      this.anomalyHistory.push({
        timestamp: new Date().toISOString(),
        range: rangeA1,
        anomaliesFound: anomalies.length,
        types: [...new Set(anomalies.map(a => a.type))],
        avgConfidence: this.calculateOverallConfidence(anomalies)
      });
      
      // Keep only recent history (last 100 detections)
      if (this.anomalyHistory.length > 100) {
        this.anomalyHistory = this.anomalyHistory.slice(-100);
      }
      
      // Save to user properties
      PropertiesService.getUserProperties().setProperty(
        'anomaly_patterns',
        JSON.stringify(this.anomalyHistory)
      );
      
    } catch (error) {
      console.error('Error saving anomaly history:', error);
    }
  },
  
  // Get anomaly detection statistics
  getAnomalyStats: function() {
    try {
      const stats = {
        totalDetections: this.anomalyHistory.length,
        totalAnomalies: 0,
        avgAnomaliesPerDetection: 0,
        mostCommonTypes: {},
        avgConfidence: 0
      };
      
      if (this.anomalyHistory.length === 0) {
        return stats;
      }
      
      // Calculate statistics
      let totalConfidence = 0;
      const typeCounts = {};
      
      this.anomalyHistory.forEach(detection => {
        stats.totalAnomalies += detection.anomaliesFound;
        totalConfidence += detection.avgConfidence;
        
        detection.types.forEach(type => {
          typeCounts[type] = (typeCounts[type] || 0) + 1;
        });
      });
      
      stats.avgAnomaliesPerDetection = stats.totalAnomalies / stats.totalDetections;
      stats.avgConfidence = totalConfidence / stats.totalDetections;
      
      // Sort types by frequency
      const sortedTypes = Object.entries(typeCounts)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 3);
      
      sortedTypes.forEach(([type, count]) => {
        stats.mostCommonTypes[type] = count;
      });
      
      return stats;
      
    } catch (error) {
      console.error('Error getting anomaly stats:', error);
      return null;
    }
  },
  
  // Learn from user feedback on anomaly detection
  learnFromAnomalyFeedback: function(rangeA1, wasAccurate, falsePositives, missedAnomalies) {
    try {
      if (!this.mlEnabled || !this.detectedAnomalies[rangeA1]) {
        return false;
      }
      
      const detection = this.detectedAnomalies[rangeA1];
      detection.feedback = {
        wasAccurate: wasAccurate,
        falsePositives: falsePositives || 0,
        missedAnomalies: missedAnomalies || 0,
        feedbackTime: new Date().toISOString()
      };
      
      // Adjust confidence thresholds based on feedback
      if (!wasAccurate) {
        if (falsePositives > missedAnomalies) {
          // Too many false positives - increase thresholds
          this.confidenceThresholds.outlier = Math.min(0.95, this.confidenceThresholds.outlier + 0.02);
          this.confidenceThresholds.trend = Math.min(0.90, this.confidenceThresholds.trend + 0.02);
          this.confidenceThresholds.pattern = Math.min(0.90, this.confidenceThresholds.pattern + 0.02);
        } else {
          // Missed too many anomalies - decrease thresholds
          this.confidenceThresholds.outlier = Math.max(0.70, this.confidenceThresholds.outlier - 0.02);
          this.confidenceThresholds.trend = Math.max(0.60, this.confidenceThresholds.trend - 0.02);
          this.confidenceThresholds.pattern = Math.max(0.65, this.confidenceThresholds.pattern - 0.02);
        }
        
        // Save updated thresholds
        PropertiesService.getUserProperties().setProperty(
          'anomaly_thresholds',
          JSON.stringify(this.confidenceThresholds)
        );
      }
      
      return true;
      
    } catch (error) {
      console.error('Error learning from anomaly feedback:', error);
      return false;
    }
  }
};