/**
 * Visual Formula Builder Backend
 * Handles drag-and-drop formula creation with intelligent suggestions
 */

const VisualFormulaBuilder = {
  
  /**
   * Show the visual formula builder interface
   */
  showVisualFormulaBuilder: function() {
    try {
      const html = HtmlService.createTemplateFromFile('VisualFormulaBuilderTemplate')
        .evaluate()
        .setTitle('Visual Formula Builder')
        .setWidth(400);
      
      SpreadsheetApp.getUi().showSidebar(html);
    } catch (error) {
      console.error('Error showing visual formula builder:', error);
      SpreadsheetApp.getUi().alert('Failed to open Visual Formula Builder: ' + error.message);
    }
  },
  
  /**
   * Suggest formula based on current data selection
   */
  suggestFormulaBasedOnData: function() {
    try {
      const range = SpreadsheetApp.getActiveRange();
      const values = range.getValues();
      
      if (!values || values.length === 0) {
        return null;
      }
      
      // Analyze data patterns
      const analysis = this.analyzeDataPattern(values);
      
      // Generate multiple suggestions based on patterns
      const suggestions = this.generateMultipleSuggestions(analysis, range);
      
      return {
        primary: suggestions[0] || null,
        alternatives: suggestions.slice(1),
        analysis: analysis,
        context: this.getDataContext(range, values)
      };
      
    } catch (error) {
      console.error('Error suggesting formula:', error);
      return null;
    }
  },
  
  /**
   * Generate multiple formula suggestions based on data patterns
   */
  generateMultipleSuggestions: function(analysis, range) {
    const suggestions = [];
    const notation = range.getA1Notation();
    
    // Numeric data suggestions
    if (analysis.hasNumbers && analysis.rows > 1) {
      suggestions.push({
        formula: '=SUM(' + notation + ')',
        description: 'Sum all numbers',
        confidence: 0.9,
        category: 'Aggregation'
      });
      
      suggestions.push({
        formula: '=AVERAGE(' + notation + ')',
        description: 'Calculate average',
        confidence: 0.8,
        category: 'Aggregation'
      });
      
      if (analysis.rows > 2) {
        suggestions.push({
          formula: '=MAX(' + notation + ')',
          description: 'Find maximum value',
          confidence: 0.7,
          category: 'Analysis'
        });
        
        suggestions.push({
          formula: '=MIN(' + notation + ')',
          description: 'Find minimum value',
          confidence: 0.7,
          category: 'Analysis'
        });
      }
    }
    
    // Currency data suggestions
    if (analysis.hasCurrency) {
      suggestions.push({
        formula: '=SUM(' + notation + ')',
        description: 'Total monetary value',
        confidence: 0.95,
        category: 'Financial'
      });
      
      if (analysis.rows > 1) {
        suggestions.push({
          formula: '=AVERAGE(' + notation + ')',
          description: 'Average amount',
          confidence: 0.8,
          category: 'Financial'
        });
      }
    }
    
    // Percentage data suggestions
    if (analysis.hasPercentage) {
      suggestions.push({
        formula: '=AVERAGE(' + notation + ')',
        description: 'Average percentage',
        confidence: 0.85,
        category: 'Statistical'
      });
    }
    
    // Date data suggestions
    if (analysis.hasDates && analysis.rows > 1) {
      suggestions.push({
        formula: '=DAYS(MAX(' + notation + '), MIN(' + notation + '))',
        description: 'Days between earliest and latest',
        confidence: 0.8,
        category: 'Date/Time'
      });
      
      suggestions.push({
        formula: '=MAX(' + notation + ')',
        description: 'Latest date',
        confidence: 0.7,
        category: 'Date/Time'
      });
      
      suggestions.push({
        formula: '=MIN(' + notation + ')',
        description: 'Earliest date',
        confidence: 0.7,
        category: 'Date/Time'
      });
    }
    
    // Text data suggestions
    if (analysis.hasText) {
      if (analysis.uniqueValues < analysis.totalValues * 0.5) {
        suggestions.push({
          formula: '=COUNTIF(' + notation + ', "criteria")',
          description: 'Count matching text',
          confidence: 0.75,
          category: 'Text Analysis'
        });
      }
      
      suggestions.push({
        formula: '=COUNTA(' + notation + ')',
        description: 'Count non-empty cells',
        confidence: 0.6,
        category: 'Counting'
      });
    }
    
    // Time series patterns
    if (analysis.patterns.includes('time_series')) {
      suggestions.push({
        formula: '=TREND(' + notation + ')',
        description: 'Calculate trend line',
        confidence: 0.8,
        category: 'Forecasting'
      });
    }
    
    // Categorical data patterns
    if (analysis.patterns.includes('categorical_data')) {
      suggestions.push({
        formula: '=UNIQUE(' + notation + ')',
        description: 'Get unique values',
        confidence: 0.75,
        category: 'Data Processing'
      });
      
      suggestions.push({
        formula: '=MODE(' + notation + ')',
        description: 'Most common value',
        confidence: 0.7,
        category: 'Statistical'
      });
    }
    
    // Conditional formulas for numeric columns
    if (analysis.patterns.includes('numeric_column') && analysis.rows > 3) {
      suggestions.push({
        formula: '=COUNTIF(' + notation + ', ">0")',
        description: 'Count positive values',
        confidence: 0.65,
        category: 'Conditional'
      });
      
      suggestions.push({
        formula: '=SUMIF(' + notation + ', ">0")',
        description: 'Sum positive values only',
        confidence: 0.65,
        category: 'Conditional'
      });
    }
    
    // Sort by confidence and return top suggestions
    return suggestions
      .sort((a, b) => b.confidence - a.confidence)
      .slice(0, 5);
  },
  
  /**
   * Get contextual information about the data
   */
  getDataContext: function(range, values) {
    const sheet = range.getSheet();
    const startRow = range.getRow();
    const startCol = range.getColumn();
    const numRows = range.getNumRows();
    const numCols = range.getNumColumns();
    
    // Check for headers
    const hasHeaders = this.detectHeaders(values, sheet, startRow);
    
    // Get column headers if they exist
    let headers = [];
    if (hasHeaders && startRow > 1) {
      const headerRange = sheet.getRange(startRow - 1, startCol, 1, numCols);
      headers = headerRange.getValues()[0];
    }
    
    // Detect data types for each column
    const columnTypes = [];
    for (let col = 0; col < numCols; col++) {
      const columnData = values.map(row => row[col]).filter(cell => cell !== null && cell !== '');
      columnTypes.push(this.detectColumnDataType(columnData));
    }
    
    return {
      hasHeaders: hasHeaders,
      headers: headers,
      columnTypes: columnTypes,
      rowCount: numRows,
      columnCount: numCols,
      sheetName: sheet.getName(),
      range: range.getA1Notation()
    };
  },
  
  /**
   * Detect if the data has headers
   */
  detectHeaders: function(values, sheet, startRow) {
    if (values.length === 0 || startRow === 1) return false;
    
    // Check if the row above contains text while current selection starts with numbers
    try {
      const headerRow = sheet.getRange(startRow - 1, 1, 1, values[0].length).getValues()[0];
      const firstDataRow = values[0];
      
      const headersAreText = headerRow.every(cell => 
        typeof cell === 'string' && cell.trim() !== ''
      );
      
      const dataHasNumbers = firstDataRow.some(cell => 
        typeof cell === 'number'
      );
      
      return headersAreText && dataHasNumbers;
    } catch (error) {
      return false;
    }
  },
  
  /**
   * Detect the primary data type of a column
   */
  detectColumnDataType: function(columnData) {
    if (columnData.length === 0) return 'empty';
    
    let numberCount = 0;
    let dateCount = 0;
    let currencyCount = 0;
    let percentageCount = 0;
    let textCount = 0;
    
    columnData.forEach(cell => {
      if (typeof cell === 'number') {
        numberCount++;
      } else if (cell instanceof Date) {
        dateCount++;
      } else if (typeof cell === 'string') {
        if (/^[$£€¥]/.test(cell)) {
          currencyCount++;
        } else if (/%$/.test(cell)) {
          percentageCount++;
        } else {
          textCount++;
        }
      }
    });
    
    const total = columnData.length;
    
    if (dateCount / total > 0.7) return 'date';
    if (currencyCount / total > 0.7) return 'currency';
    if (percentageCount / total > 0.7) return 'percentage';
    if (numberCount / total > 0.7) return 'number';
    if (textCount / total > 0.7) return 'text';
    
    return 'mixed';
  },
  
  /**
   * Get smart formula suggestions with context
   */
  getSmartFormulaSuggestions: function() {
    try {
      const suggestionData = this.suggestFormulaBasedOnData();
      
      if (!suggestionData) {
        return {
          success: false,
          message: 'No data selected. Please select a range with data to get suggestions.'
        };
      }
      
      return {
        success: true,
        suggestions: [suggestionData.primary, ...suggestionData.alternatives].filter(Boolean),
        analysis: suggestionData.analysis,
        context: suggestionData.context
      };
      
    } catch (error) {
      console.error('Error getting smart suggestions:', error);
      return {
        success: false,
        error: error.message
      };
    }
  },
  
  /**
   * Analyze data pattern in selection
   */
  analyzeDataPattern: function(values) {
    const analysis = {
      rows: values.length,
      cols: values[0] ? values[0].length : 0,
      hasNumbers: false,
      hasText: false,
      hasDates: false,
      hasCurrency: false,
      hasPercentage: false,
      totalValues: 0,
      uniqueValues: 0,
      isEmpty: true,
      patterns: []
    };
    
    const uniqueSet = new Set();
    
    values.forEach(row => {
      row.forEach(cell => {
        if (cell !== null && cell !== '') {
          analysis.isEmpty = false;
          analysis.totalValues++;
          uniqueSet.add(String(cell));
          
          if (typeof cell === 'number') {
            analysis.hasNumbers = true;
          } else if (cell instanceof Date) {
            analysis.hasDates = true;
          } else if (typeof cell === 'string') {
            analysis.hasText = true;
            
            // Check for currency
            if (/^[$£€¥]/.test(cell)) {
              analysis.hasCurrency = true;
            }
            
            // Check for percentage
            if (/%$/.test(cell)) {
              analysis.hasPercentage = true;
            }
          }
        }
      });
    });
    
    analysis.uniqueValues = uniqueSet.size;
    
    // Detect specific patterns
    if (analysis.hasNumbers && analysis.rows > 1) {
      analysis.patterns.push('numeric_column');
    }
    
    if (analysis.uniqueValues < analysis.totalValues * 0.3) {
      analysis.patterns.push('categorical_data');
    }
    
    if (analysis.hasDates && analysis.rows > 1) {
      analysis.patterns.push('time_series');
    }
    
    return analysis;
  },
  
  /**
   * Validate the visual formula
   */
  validateVisualFormula: function(formula) {
    try {
      if (!formula || !formula.startsWith('=')) {
        return {
          valid: false,
          issues: ['Formula must start with =']
        };
      }
      
      const issues = [];
      
      // Check for balanced parentheses
      let parenCount = 0;
      for (let char of formula) {
        if (char === '(') parenCount++;
        if (char === ')') parenCount--;
        if (parenCount < 0) {
          issues.push('Unmatched closing parenthesis');
          break;
        }
      }
      if (parenCount > 0) {
        issues.push('Unmatched opening parenthesis');
      }
      
      // Check for valid function names
      const functionPattern = /\b([A-Z]+)\(/g;
      const validFunctions = [
        'SUM', 'AVERAGE', 'COUNT', 'MAX', 'MIN', 'IF', 'AND', 'OR', 'NOT',
        'VLOOKUP', 'INDEX', 'MATCH', 'FILTER', 'SUMIF', 'COUNTIF', 'AVERAGEIF',
        'IFS', 'ROUND', 'ABS', 'SQRT', 'POWER', 'CONCATENATE', 'LEFT', 'RIGHT',
        'MID', 'LEN', 'FIND', 'SUBSTITUTE', 'TODAY', 'NOW', 'DATE', 'DAYS'
      ];
      
      let match;
      while ((match = functionPattern.exec(formula)) !== null) {
        if (!validFunctions.includes(match[1])) {
          issues.push(`Unknown function: ${match[1]}`);
        }
      }
      
      // Check for valid range references
      const rangePattern = /[A-Z]+[0-9]+(?::[A-Z]+[0-9]+)?/g;
      const ranges = formula.match(rangePattern) || [];
      
      ranges.forEach(range => {
        // Validate range format
        if (range.includes(':')) {
          const [start, end] = range.split(':');
          if (!this.isValidCellReference(start) || !this.isValidCellReference(end)) {
            issues.push(`Invalid range: ${range}`);
          }
        } else {
          if (!this.isValidCellReference(range)) {
            issues.push(`Invalid cell reference: ${range}`);
          }
        }
      });
      
      // Check for common syntax errors
      if (/[+\-*/]{2,}/.test(formula)) {
        issues.push('Double operators detected');
      }
      
      if (/\(\s*\)/.test(formula)) {
        issues.push('Empty parentheses detected');
      }
      
      return {
        valid: issues.length === 0,
        issues: issues
      };
      
    } catch (error) {
      console.error('Error validating formula:', error);
      return {
        valid: false,
        issues: ['Validation error: ' + error.message]
      };
    }
  },
  
  /**
   * Check if a cell reference is valid
   */
  isValidCellReference: function(ref) {
    // Check format: Column letters (A-ZZ) followed by row numbers (1-999999)
    const pattern = /^[A-Z]{1,3}[0-9]{1,6}$/;
    return pattern.test(ref);
  },
  
  /**
   * Optimize the visual formula
   */
  optimizeVisualFormula: function(formula) {
    try {
      let optimized = formula;
      
      // Replace VLOOKUP with INDEX/MATCH for better performance
      if (formula.includes('VLOOKUP')) {
        // This is a simplified example - real implementation would parse parameters
        optimized = optimized.replace(
          /VLOOKUP\(([^,]+),\s*([^,]+),\s*([^,]+),\s*([^)]+)\)/g,
          'INDEX($2, MATCH($1, INDEX($2, 0, 1), $4), $3)'
        );
      }
      
      // Replace multiple IFs with IFS
      const ifCount = (formula.match(/\bIF\(/g) || []).length;
      if (ifCount > 2) {
        // Suggest using IFS for multiple conditions
        console.log('Consider using IFS function for multiple conditions');
      }
      
      // Replace SUM(A1, A2, A3...) with SUM(A1:A3)
      optimized = this.optimizeConsecutiveRanges(optimized);
      
      // Remove unnecessary parentheses
      optimized = this.removeUnnecessaryParentheses(optimized);
      
      // Suggest SUMPRODUCT for array calculations
      if (formula.includes('SUM') && formula.includes('*')) {
        console.log('Consider using SUMPRODUCT for array calculations');
      }
      
      return optimized;
      
    } catch (error) {
      console.error('Error optimizing formula:', error);
      return formula;
    }
  },
  
  /**
   * Optimize consecutive cell ranges
   */
  optimizeConsecutiveRanges: function(formula) {
    // Look for patterns like A1, A2, A3 and replace with A1:A3
    const cellPattern = /([A-Z]+)(\d+)(?:,\s*\1(\d+))+/g;
    
    return formula.replace(cellPattern, function(match, col, startRow) {
      const cells = match.split(',').map(c => c.trim());
      const rows = cells.map(c => parseInt(c.substring(col.length)));
      
      // Check if rows are consecutive
      let isConsecutive = true;
      for (let i = 1; i < rows.length; i++) {
        if (rows[i] !== rows[i-1] + 1) {
          isConsecutive = false;
          break;
        }
      }
      
      if (isConsecutive) {
        return `${col}${rows[0]}:${col}${rows[rows.length-1]}`;
      }
      
      return match;
    });
  },
  
  /**
   * Remove unnecessary parentheses
   */
  removeUnnecessaryParentheses: function(formula) {
    // Remove double parentheses
    let optimized = formula;
    while (optimized.includes('((') && optimized.includes('))')) {
      optimized = optimized.replace(/\(\(([^()]+)\)\)/g, '($1)');
    }
    return optimized;
  },
  
  /**
   * Test the visual formula with sample data
   */
  testVisualFormula: function(formula) {
    try {
      // Create a temporary cell to test the formula
      const sheet = SpreadsheetApp.getActiveSheet();
      const testRange = sheet.getRange(sheet.getMaxRows(), sheet.getMaxColumns());
      
      // Store original value
      const originalValue = testRange.getValue();
      
      try {
        // Test the formula
        testRange.setFormula(formula);
        const result = testRange.getValue();
        
        // Clean up
        if (originalValue) {
          testRange.setValue(originalValue);
        } else {
          testRange.clear();
        }
        
        return 'Result: ' + result;
        
      } catch (formulaError) {
        // Clean up on error
        if (originalValue) {
          testRange.setValue(originalValue);
        } else {
          testRange.clear();
        }
        
        return 'Formula error: ' + formulaError.message;
      }
      
    } catch (error) {
      console.error('Error testing formula:', error);
      return 'Test failed: ' + error.message;
    }
  },
  
  /**
   * Insert the visual formula into the active cell
   */
  insertVisualFormula: function(formula) {
    try {
      const cell = SpreadsheetApp.getActiveCell();
      cell.setFormula(formula);
      
      return {
        success: true,
        cell: cell.getA1Notation(),
        formula: formula
      };
      
    } catch (error) {
      console.error('Error inserting formula:', error);
      throw new Error('Failed to insert formula: ' + error.message);
    }
  },
  
  /**
   * Get formula templates based on data type
   */
  getFormulaTemplates: function(dataType) {
    const templates = {
      'numeric': [
        { name: 'Sum', formula: '=SUM(range)', description: 'Add all numbers' },
        { name: 'Average', formula: '=AVERAGE(range)', description: 'Calculate mean' },
        { name: 'Running Total', formula: '=SUM($A$2:A2)', description: 'Cumulative sum' },
        { name: 'Percentage of Total', formula: '=A2/SUM($A$2:$A$100)*100', description: 'Calculate percentage' },
        { name: 'Year-over-Year Growth', formula: '=(B2-A2)/A2*100', description: 'Calculate growth rate' }
      ],
      'text': [
        { name: 'Concatenate', formula: '=CONCATENATE(A2, " ", B2)', description: 'Combine text' },
        { name: 'Extract First Name', formula: '=LEFT(A2, FIND(" ", A2)-1)', description: 'Get first name' },
        { name: 'Extract Last Name', formula: '=RIGHT(A2, LEN(A2)-FIND(" ", A2))', description: 'Get last name' },
        { name: 'Count Occurrences', formula: '=COUNTIF(range, "text")', description: 'Count specific text' }
      ],
      'date': [
        { name: 'Days Between', formula: '=DAYS(end_date, start_date)', description: 'Calculate days' },
        { name: 'Month Name', formula: '=TEXT(A2, "MMMM")', description: 'Get month name' },
        { name: 'Quarter', formula: '=ROUNDUP(MONTH(A2)/3, 0)', description: 'Get quarter number' },
        { name: 'Age in Years', formula: '=DATEDIF(A2, TODAY(), "Y")', description: 'Calculate age' }
      ],
      'financial': [
        { name: 'Profit Margin', formula: '=(Revenue-Cost)/Revenue*100', description: 'Calculate margin %' },
        { name: 'Compound Interest', formula: '=FV(rate, periods, payment, present_value)', description: 'Future value' },
        { name: 'Loan Payment', formula: '=PMT(rate/12, months, -principal)', description: 'Monthly payment' },
        { name: 'ROI', formula: '=(Gain-Cost)/Cost*100', description: 'Return on investment' }
      ]
    };
    
    return templates[dataType] || templates['numeric'];
  },
  
  /**
   * Generate smart formula based on column headers
   */
  generateSmartFormula: function(headers, operation) {
    const formulas = {
      'total_sales': () => {
        const salesCol = this.findColumn(headers, ['sales', 'revenue', 'amount']);
        return salesCol ? `=SUM(${salesCol}:${salesCol})` : null;
      },
      'average_price': () => {
        const priceCol = this.findColumn(headers, ['price', 'cost', 'rate']);
        return priceCol ? `=AVERAGE(${priceCol}:${priceCol})` : null;
      },
      'count_items': () => {
        const idCol = this.findColumn(headers, ['id', 'item', 'product', 'name']);
        return idCol ? `=COUNTA(${idCol}:${idCol})-1` : null;
      },
      'lookup_value': () => {
        const idCol = this.findColumn(headers, ['id', 'code', 'key']);
        const nameCol = this.findColumn(headers, ['name', 'description', 'title']);
        if (idCol && nameCol) {
          return `=VLOOKUP(lookup_value, ${idCol}:${nameCol}, 2, FALSE)`;
        }
        return null;
      }
    };
    
    return formulas[operation] ? formulas[operation]() : null;
  },
  
  /**
   * Find column letter by header name
   */
  findColumn: function(headers, keywords) {
    for (let i = 0; i < headers.length; i++) {
      const header = headers[i].toLowerCase();
      for (let keyword of keywords) {
        if (header.includes(keyword)) {
          return String.fromCharCode(65 + i); // Convert to column letter
        }
      }
    }
    return null;
  },
  
  /**
   * Get cross-sheet information for formula builder
   */
  getCrossSheetInfo: function() {
    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const sheets = spreadsheet.getSheets();
      const currentSheet = SpreadsheetApp.getActiveSheet();
      
      const sheetInfo = sheets.map(sheet => {
        const lastRow = sheet.getLastRow();
        const lastColumn = sheet.getLastColumn();
        
        return {
          name: sheet.getName(),
          rowCount: lastRow,
          columnCount: lastColumn,
          lastColumn: this.columnNumberToLetter(lastColumn),
          dataRange: lastRow > 0 && lastColumn > 0 ? 
            'A1:' + this.columnNumberToLetter(lastColumn) + lastRow : 'A1',
          isCurrent: sheet.getName() === currentSheet.getName()
        };
      }).filter(sheet => !sheet.isCurrent); // Exclude current sheet
      
      return sheetInfo;
      
    } catch (error) {
      console.error('Error getting cross-sheet info:', error);
      throw new Error('Failed to get sheet information: ' + error.message);
    }
  },
  
  /**
   * Preview data from a sheet range
   */
  previewSheetRange: function(sheetName, range) {
    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = spreadsheet.getSheetByName(sheetName);
      
      if (!sheet) {
        return {
          success: false,
          error: 'Sheet "' + sheetName + '" not found'
        };
      }
      
      // Validate and adjust range
      const adjustedRange = this.validateAndAdjustRange(sheet, range);
      if (!adjustedRange) {
        return {
          success: false,
          error: 'Invalid range: ' + range
        };
      }
      
      const rangeObject = sheet.getRange(adjustedRange);
      const values = rangeObject.getValues();
      
      return {
        success: true,
        preview: values,
        actualRange: adjustedRange,
        rowCount: values.length,
        columnCount: values.length > 0 ? values[0].length : 0
      };
      
    } catch (error) {
      console.error('Error previewing sheet range:', error);
      return {
        success: false,
        error: error.message
      };
    }
  },
  
  /**
   * Validate and adjust a range string for a sheet
   */
  validateAndAdjustRange: function(sheet, range) {
    try {
      const lastRow = sheet.getLastRow();
      const lastColumn = sheet.getLastColumn();
      
      // Handle full column references (e.g., "A:A")
      if (range.match(/^[A-Z]+:[A-Z]+$/)) {
        return range;
      }
      
      // Handle row references (e.g., "1:1")
      if (range.match(/^\d+:\d+$/)) {
        return range;
      }
      
      // Handle standard range (e.g., "A1:C10")
      if (range.match(/^[A-Z]+\d+:[A-Z]+\d+$/)) {
        // Try to get the range to validate it exists
        try {
          sheet.getRange(range);
          return range;
        } catch (e) {
          // If range is invalid, try to adjust it
          const parts = range.split(':');
          const startCell = parts[0];
          const endCell = parts[1];
          
          const startCol = this.extractColumnFromCell(startCell);
          const startRow = this.extractRowFromCell(startCell);
          const endCol = this.extractColumnFromCell(endCell);
          const endRow = this.extractRowFromCell(endCell);
          
          // Adjust to sheet bounds
          const maxRow = Math.min(endRow, lastRow);
          const maxCol = Math.min(this.columnLetterToNumber(endCol), lastColumn);
          
          if (maxRow > 0 && maxCol > 0) {
            return startCell + ':' + this.columnNumberToLetter(maxCol) + maxRow;
          }
        }
      }
      
      // Handle single cell reference
      if (range.match(/^[A-Z]+\d+$/)) {
        try {
          sheet.getRange(range);
          return range;
        } catch (e) {
          return null;
        }
      }
      
      return null;
      
    } catch (error) {
      console.error('Error validating range:', error);
      return null;
    }
  },
  
  /**
   * Extract column letter from cell reference
   */
  extractColumnFromCell: function(cell) {
    return cell.replace(/\d+/g, '');
  },
  
  /**
   * Extract row number from cell reference
   */
  extractRowFromCell: function(cell) {
    const match = cell.match(/\d+/);
    return match ? parseInt(match[0]) : 1;
  },
  
  /**
   * Convert column number to letter (1 -> A, 26 -> Z, 27 -> AA)
   */
  columnNumberToLetter: function(column) {
    let temp, letter = '';
    while (column > 0) {
      temp = (column - 1) % 26;
      letter = String.fromCharCode(temp + 65) + letter;
      column = (column - temp - 1) / 26;
    }
    return letter;
  },
  
  /**
   * Convert column letter to number (A -> 1, Z -> 26, AA -> 27)
   */
  columnLetterToNumber: function(column) {
    let result = 0;
    for (let i = 0; i < column.length; i++) {
      result = result * 26 + (column.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
    }
    return result;
  },
  
  /**
   * Build cross-sheet formula with references
   */
  buildCrossSheetFormula: function(components) {
    try {
      let formula = '=';
      let hasSheetReference = false;
      
      components.forEach((component, index) => {
        if (component.type === 'sheet-reference') {
          // Handle sheet reference (e.g., "Sheet2!A1:C10")
          formula += component.value;
          hasSheetReference = true;
        } else if (component.type === 'function') {
          // Handle function with sheet references
          formula += component.value + '(';
        } else if (component.type === 'operator') {
          formula += component.value;
        } else {
          formula += component.value;
        }
      });
      
      // Close any open parentheses
      const openParens = (formula.match(/\(/g) || []).length;
      const closeParens = (formula.match(/\)/g) || []).length;
      formula += ')'.repeat(openParens - closeParens);
      
      return {
        success: true,
        formula: formula,
        hasSheetReference: hasSheetReference
      };
      
    } catch (error) {
      console.error('Error building cross-sheet formula:', error);
      return {
        success: false,
        error: error.message
      };
    }
  },
  
  /**
   * Analyze multi-tab relationships across all sheets
   */
  analyzeMultiTabRelationships: function() {
    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const sheets = spreadsheet.getSheets();
      
      if (sheets.length < 2) {
        return {
          success: false,
          error: 'Spreadsheet must have at least 2 sheets to analyze relationships'
        };
      }
      
      const sheetData = [];
      const relationships = [];
      
      // Analyze each sheet
      sheets.forEach(sheet => {
        const analysis = this.analyzeSheetForReferences(sheet, sheets);
        sheetData.push(analysis.sheetInfo);
        relationships.push(...analysis.relationships);
      });
      
      // Calculate relationship strengths
      const processedRelationships = this.calculateRelationshipStrengths(relationships);
      
      // Add incoming reference counts
      this.calculateIncomingReferences(sheetData, processedRelationships);
      
      return {
        success: true,
        data: {
          sheets: sheetData,
          relationships: processedRelationships,
          totalSheets: sheets.length,
          totalRelationships: processedRelationships.length,
          analysisTimestamp: new Date().toISOString()
        }
      };
      
    } catch (error) {
      console.error('Error analyzing multi-tab relationships:', error);
      return {
        success: false,
        error: error.message
      };
    }
  },
  
  /**
   * Analyze a single sheet for cross-sheet references
   */
  analyzeSheetForReferences: function(sheet, allSheets) {
    const sheetName = sheet.getName();
    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();
    
    const sheetInfo = {
      name: sheetName,
      rowCount: lastRow,
      columnCount: lastColumn,
      outgoingRefs: 0,
      incomingRefs: 0, // Will be calculated later
      lastModified: this.getSheetLastModified(sheet)
    };
    
    const relationships = [];
    
    if (lastRow === 0 || lastColumn === 0) {
      return { sheetInfo, relationships };
    }
    
    try {
      // Get all formulas in the sheet
      const formulas = sheet.getRange(1, 1, lastRow, lastColumn).getFormulas();
      const crossSheetRefs = new Map(); // Track references by target sheet
      
      // Analyze formulas for cross-sheet references
      formulas.forEach((row, rowIndex) => {
        row.forEach((formula, colIndex) => {
          if (formula && formula.startsWith('=')) {
            const references = this.extractCrossSheetReferences(formula, sheetName);
            
            references.forEach(ref => {
              if (!crossSheetRefs.has(ref.targetSheet)) {
                crossSheetRefs.set(ref.targetSheet, []);
              }
              crossSheetRefs.get(ref.targetSheet).push({
                sourceCell: this.columnNumberToLetter(colIndex + 1) + (rowIndex + 1),
                targetRange: ref.range,
                formula: formula
              });
            });
          }
        });
      });
      
      // Convert to relationship objects
      crossSheetRefs.forEach((refs, targetSheet) => {
        // Verify target sheet exists
        if (allSheets.some(s => s.getName() === targetSheet)) {
          relationships.push({
            fromSheet: sheetName,
            toSheet: targetSheet,
            references: refs,
            referenceCount: refs.length
          });
          sheetInfo.outgoingRefs += refs.length;
        }
      });
      
    } catch (error) {
      console.error('Error analyzing sheet ' + sheetName + ':', error);
    }
    
    return { sheetInfo, relationships };
  },
  
  /**
   * Extract cross-sheet references from a formula
   */
  extractCrossSheetReferences: function(formula, currentSheetName) {
    const references = [];
    
    // Pattern to match cross-sheet references: 'SheetName'!Range or SheetName!Range
    const crossSheetPattern = /(?:'([^']+)'|([A-Za-z0-9_\s]+))!([A-Z]+[0-9]+(?::[A-Z]+[0-9]+)?|[A-Z]+:[A-Z]+|[0-9]+:[0-9]+)/g;
    
    let match;
    while ((match = crossSheetPattern.exec(formula)) !== null) {
      const targetSheet = match[1] || match[2]; // Quoted or unquoted sheet name
      const range = match[3];
      
      // Skip if it's referencing the current sheet
      if (targetSheet !== currentSheetName) {
        references.push({
          targetSheet: targetSheet,
          range: range,
          fullReference: match[0]
        });
      }
    }
    
    return references;
  },
  
  /**
   * Calculate relationship strengths based on reference count
   */
  calculateRelationshipStrengths: function(relationships) {
    return relationships.map(rel => {
      let strength;
      if (rel.referenceCount >= 5) {
        strength = 'strong';
      } else if (rel.referenceCount >= 2) {
        strength = 'medium';
      } else {
        strength = 'weak';
      }
      
      return {
        ...rel,
        strength: strength
      };
    });
  },
  
  /**
   * Calculate incoming reference counts for each sheet
   */
  calculateIncomingReferences: function(sheetData, relationships) {
    const incomingCounts = new Map();
    
    // Initialize counts
    sheetData.forEach(sheet => {
      incomingCounts.set(sheet.name, 0);
    });
    
    // Count incoming references
    relationships.forEach(rel => {
      const currentCount = incomingCounts.get(rel.toSheet) || 0;
      incomingCounts.set(rel.toSheet, currentCount + rel.referenceCount);
    });
    
    // Update sheet data
    sheetData.forEach(sheet => {
      sheet.incomingRefs = incomingCounts.get(sheet.name) || 0;
    });
  },
  
  /**
   * Get approximate last modified time for a sheet
   */
  getSheetLastModified: function(sheet) {
    try {
      // This is a simplified approach - in reality, Google Sheets doesn't 
      // provide direct last modified times for individual sheets
      const lastRow = sheet.getLastRow();
      if (lastRow > 0) {
        // Try to get the last cell with data as a proxy for last modified
        const lastRange = sheet.getRange(lastRow, sheet.getLastColumn());
        return 'Recent'; // Placeholder
      }
      return 'Unknown';
    } catch (error) {
      return 'Unknown';
    }
  },
  
  /**
   * Generate a detailed relationship report
   */
  generateRelationshipReport: function(relationshipData) {
    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const reportSheetName = 'Relationship Analysis Report';
      
      // Delete existing report sheet if it exists
      const existingSheet = spreadsheet.getSheetByName(reportSheetName);
      if (existingSheet) {
        spreadsheet.deleteSheet(existingSheet);
      }
      
      // Create new report sheet
      const reportSheet = spreadsheet.insertSheet(reportSheetName);
      
      // Set up the report
      this.setupRelationshipReport(reportSheet, relationshipData);
      
      return {
        success: true,
        sheetName: reportSheetName
      };
      
    } catch (error) {
      console.error('Error generating relationship report:', error);
      return {
        success: false,
        error: error.message
      };
    }
  },
  
  /**
   * Set up the relationship report sheet
   */
  setupRelationshipReport: function(sheet, data) {
    // Title
    sheet.getRange(1, 1).setValue('Multi-Tab Relationship Analysis Report');
    sheet.getRange(1, 1).setFontSize(16).setFontWeight('bold');
    
    // Summary section
    let row = 3;
    sheet.getRange(row, 1).setValue('Summary');
    sheet.getRange(row, 1).setFontWeight('bold');
    row++;
    
    sheet.getRange(row, 1).setValue('Total Sheets:');
    sheet.getRange(row, 2).setValue(data.totalSheets);
    row++;
    
    sheet.getRange(row, 1).setValue('Total Relationships:');
    sheet.getRange(row, 2).setValue(data.totalRelationships);
    row++;
    
    sheet.getRange(row, 1).setValue('Analysis Date:');
    sheet.getRange(row, 2).setValue(new Date(data.analysisTimestamp));
    row += 2;
    
    // Sheet Details section
    sheet.getRange(row, 1).setValue('Sheet Details');
    sheet.getRange(row, 1).setFontWeight('bold');
    row++;
    
    // Headers
    const headers = ['Sheet Name', 'Rows', 'Columns', 'Outgoing Refs', 'Incoming Refs'];
    headers.forEach((header, index) => {
      sheet.getRange(row, index + 1).setValue(header);
      sheet.getRange(row, index + 1).setFontWeight('bold');
    });
    row++;
    
    // Sheet data
    data.sheets.forEach(sheetData => {
      sheet.getRange(row, 1).setValue(sheetData.name);
      sheet.getRange(row, 2).setValue(sheetData.rowCount);
      sheet.getRange(row, 3).setValue(sheetData.columnCount);
      sheet.getRange(row, 4).setValue(sheetData.outgoingRefs);
      sheet.getRange(row, 5).setValue(sheetData.incomingRefs);
      row++;
    });
    
    row += 2;
    
    // Relationships section
    sheet.getRange(row, 1).setValue('Detailed Relationships');
    sheet.getRange(row, 1).setFontWeight('bold');
    row++;
    
    // Relationship headers
    const relHeaders = ['From Sheet', 'To Sheet', 'References', 'Strength'];
    relHeaders.forEach((header, index) => {
      sheet.getRange(row, index + 1).setValue(header);
      sheet.getRange(row, index + 1).setFontWeight('bold');
    });
    row++;
    
    // Relationship data
    data.relationships.forEach(rel => {
      sheet.getRange(row, 1).setValue(rel.fromSheet);
      sheet.getRange(row, 2).setValue(rel.toSheet);
      sheet.getRange(row, 3).setValue(rel.referenceCount);
      sheet.getRange(row, 4).setValue(rel.strength);
      row++;
    });
    
    // Auto-resize columns
    sheet.autoResizeColumns(1, 5);
    
    // Activate the report sheet
    sheet.activate();
  }
};