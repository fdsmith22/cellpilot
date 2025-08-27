// Conditional Formatting Wizard Module
const ConditionalFormattingWizard = {
  
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
      }
    };
    
    const rule = presets[preset];
    if (!rule) {
      return {success: false, error: 'Preset not found'};
    }
    
    rule.range = rangeA1;
    return this.applyConditionalFormatting(rule);
  }
};