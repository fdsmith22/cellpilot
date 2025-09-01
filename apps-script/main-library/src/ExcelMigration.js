/**
 * Excel to Google Sheets Migration Assistant
 * Handles formula conversion, reference fixes, and optimization
 */

const ExcelMigration = {
  
  /**
   * Scan sheet for Excel-related issues
   * @return {Object} Analysis results with issues found
   */
  scanForExcelIssues: function() {
    try {
      const sheet = SpreadsheetApp.getActiveSheet();
      const dataRange = sheet.getDataRange();
      const formulas = dataRange.getFormulas();
      const values = dataRange.getValues();
      
      let issues = {
        formulaErrors: [],
        refErrors: [],
        volatileFunctions: [],
        localeIssues: [],
        xlookupFormulas: [],
        structuredRefs: [],
        performanceIssues: []
      };
      
      // Scan all cells
      for (let row = 0; row < formulas.length; row++) {
        for (let col = 0; col < formulas[row].length; col++) {
          const formula = formulas[row][col];
          const value = values[row][col];
          const cellA1 = sheet.getRange(row + 1, col + 1).getA1Notation();
          
          if (formula) {
            // Check for Excel-specific functions
            if (formula.includes('XLOOKUP')) {
              issues.xlookupFormulas.push({
                cell: cellA1,
                formula: formula,
                suggestion: 'Convert to VLOOKUP or INDEX/MATCH'
              });
            }
            
            // Check for structured references (Excel tables)
            if (formula.match(/\[@?\[.*?\]\]/)) {
              issues.structuredRefs.push({
                cell: cellA1,
                formula: formula,
                suggestion: 'Convert to A1 notation'
              });
            }
            
            // Check for volatile functions
            const volatilePattern = /\b(NOW|TODAY|RAND|RANDBETWEEN|INDIRECT|OFFSET)\s*\(/gi;
            const volatileMatches = formula.match(volatilePattern);
            if (volatileMatches) {
              issues.volatileFunctions.push({
                cell: cellA1,
                formula: formula,
                functions: volatileMatches,
                suggestion: 'Consider using static values or less frequent updates'
              });
            }
            
            // Check for locale issues (semicolon vs comma)
            if (formula.includes(';') && formula.match(/\w+\s*\(/)) {
              issues.localeIssues.push({
                cell: cellA1,
                formula: formula,
                suggestion: 'Replace semicolons with commas for Google Sheets'
              });
            }
          }
          
          // Check for error values
          if (typeof value === 'string') {
            if (value === '#REF!') {
              issues.refErrors.push({
                cell: cellA1,
                error: '#REF!',
                suggestion: 'Fix broken reference'
              });
            } else if (value === '#NAME?') {
              issues.formulaErrors.push({
                cell: cellA1,
                error: '#NAME?',
                suggestion: 'Function not recognized - check syntax'
              });
            } else if (value.startsWith('#')) {
              issues.formulaErrors.push({
                cell: cellA1,
                error: value,
                suggestion: 'Formula error - check syntax'
              });
            }
          }
        }
      }
      
      // Check for performance issues
      const formulaCount = formulas.flat().filter(f => f !== '').length;
      if (formulaCount > 1000) {
        issues.performanceIssues.push({
          type: 'High formula count',
          count: formulaCount,
          suggestion: 'Consider using ArrayFormulas or reducing formula complexity'
        });
      }
      
      if (issues.volatileFunctions.length > 50) {
        issues.performanceIssues.push({
          type: 'Excessive volatile functions',
          count: issues.volatileFunctions.length,
          suggestion: 'Replace volatile functions with static values where possible'
        });
      }
      
      return {
        success: true,
        issues: issues,
        summary: {
          totalIssues: Object.values(issues).reduce((sum, arr) => sum + arr.length, 0),
          formulaErrors: issues.formulaErrors.length,
          refErrors: issues.refErrors.length,
          volatileFunctions: issues.volatileFunctions.length,
          localeIssues: issues.localeIssues.length,
          xlookupFormulas: issues.xlookupFormulas.length,
          structuredRefs: issues.structuredRefs.length,
          performanceIssues: issues.performanceIssues.length
        }
      };
      
    } catch (error) {
      Logger.error('Error scanning for Excel issues:', error);
      return { success: false, error: error.message };
    }
  },
  
  /**
   * Fix all detected Excel migration issues
   * @param {Object} options - Fix options
   * @return {Object} Result with fixes applied
   */
  fixExcelIssues: function(options) {
    try {
      const sheet = SpreadsheetApp.getActiveSheet();
      const scan = this.scanForExcelIssues();
      
      if (!scan.success) {
        return scan;
      }
      
      let fixCount = 0;
      let fixes = [];
      
      // Store original state for undo
      if (options.createBackup) {
        UndoManager.saveState('excel_migration', sheet.getDataRange());
      }
      
      // Fix locale issues (semicolons to commas)
      if (options.fixLocale && scan.issues.localeIssues.length > 0) {
        scan.issues.localeIssues.forEach(issue => {
          try {
            const range = sheet.getRange(issue.cell);
            const fixed = issue.formula.replace(/;/g, ',');
            range.setFormula(fixed);
            fixes.push(`Fixed locale in ${issue.cell}`);
            fixCount++;
          } catch (e) {
            Logger.error(`Failed to fix locale in ${issue.cell}:`, e);
          }
        });
      }
      
      // Convert XLOOKUP to VLOOKUP/INDEX-MATCH
      if (options.convertXlookup && scan.issues.xlookupFormulas.length > 0) {
        scan.issues.xlookupFormulas.forEach(issue => {
          try {
            const range = sheet.getRange(issue.cell);
            const converted = this.convertXlookupToVlookup(issue.formula);
            if (converted !== issue.formula) {
              range.setFormula(converted);
              fixes.push(`Converted XLOOKUP in ${issue.cell}`);
              fixCount++;
            }
          } catch (e) {
            Logger.error(`Failed to convert XLOOKUP in ${issue.cell}:`, e);
          }
        });
      }
      
      // Fix structured references
      if (options.fixStructuredRefs && scan.issues.structuredRefs.length > 0) {
        scan.issues.structuredRefs.forEach(issue => {
          try {
            const range = sheet.getRange(issue.cell);
            const fixed = this.convertStructuredRefs(issue.formula);
            if (fixed !== issue.formula) {
              range.setFormula(fixed);
              fixes.push(`Fixed structured reference in ${issue.cell}`);
              fixCount++;
            }
          } catch (e) {
            Logger.error(`Failed to fix structured ref in ${issue.cell}:`, e);
          }
        });
      }
      
      // Optimize volatile functions
      if (options.optimizeVolatile && scan.issues.volatileFunctions.length > 0) {
        const toOptimize = scan.issues.volatileFunctions.slice(0, 20); // Limit to prevent timeout
        toOptimize.forEach(issue => {
          try {
            const range = sheet.getRange(issue.cell);
            const optimized = this.optimizeVolatileFunction(issue.formula);
            if (optimized.changed) {
              if (optimized.useValue) {
                range.setValue(optimized.value);
              } else {
                range.setFormula(optimized.formula);
              }
              fixes.push(`Optimized volatile function in ${issue.cell}`);
              fixCount++;
            }
          } catch (e) {
            Logger.error(`Failed to optimize volatile in ${issue.cell}:`, e);
          }
        });
      }
      
      return {
        success: true,
        message: `Fixed ${fixCount} Excel migration issues`,
        fixCount: fixCount,
        fixes: fixes.slice(0, 50), // Limit returned fixes for UI
        remaining: Math.max(0, scan.summary.totalIssues - fixCount)
      };
      
    } catch (error) {
      Logger.error('Error fixing Excel issues:', error);
      return { success: false, error: error.message };
    }
  },
  
  /**
   * Convert XLOOKUP to VLOOKUP or INDEX/MATCH
   * @param {string} formula - Original XLOOKUP formula
   * @return {string} Converted formula
   */
  convertXlookupToVlookup: function(formula) {
    // Basic XLOOKUP to VLOOKUP conversion
    // XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])
    // to VLOOKUP(lookup_value, range, col_index, [is_sorted])
    
    const xlookupMatch = formula.match(/XLOOKUP\s*\((.*?)\)/i);
    if (!xlookupMatch) return formula;
    
    try {
      // Parse arguments (simplified - doesn't handle nested functions)
      const args = xlookupMatch[1].split(',').map(arg => arg.trim());
      
      if (args.length >= 3) {
        const lookupValue = args[0];
        const lookupArray = args[1];
        const returnArray = args[2];
        
        // Try to create an INDEX/MATCH formula (more flexible than VLOOKUP)
        const indexMatch = `INDEX(${returnArray}, MATCH(${lookupValue}, ${lookupArray}, 0))`;
        return formula.replace(xlookupMatch[0], indexMatch);
      }
    } catch (e) {
      Logger.error('Error converting XLOOKUP:', e);
    }
    
    return formula;
  },
  
  /**
   * Convert Excel structured references to A1 notation
   * @param {string} formula - Formula with structured references
   * @return {string} Formula with A1 notation
   */
  convertStructuredRefs: function(formula) {
    // Convert Table[Column] to column range
    // This is simplified - real implementation would need table mapping
    
    let converted = formula;
    
    // Replace [@Column] with relative reference
    converted = converted.replace(/\[@\[([^\]]+)\]\]/g, function(match, columnName) {
      // This would need actual column mapping
      return 'A:A'; // Placeholder
    });
    
    // Replace Table[Column] with absolute reference
    converted = converted.replace(/(\w+)\[\[([^\]]+)\]\]/g, function(match, tableName, columnName) {
      // This would need actual table and column mapping
      return `${tableName}!A:A`; // Placeholder
    });
    
    return converted;
  },
  
  /**
   * Optimize volatile functions
   * @param {string} formula - Formula with volatile functions
   * @return {Object} Optimization result
   */
  optimizeVolatileFunction: function(formula) {
    let result = {
      changed: false,
      useValue: false,
      formula: formula,
      value: null
    };
    
    // Replace TODAY() with current date value if it's not critical to update
    if (formula === '=TODAY()') {
      result.changed = true;
      result.useValue = true;
      result.value = new Date();
      result.value.setHours(0, 0, 0, 0);
      return result;
    }
    
    // Replace NOW() with current timestamp if it's not critical to update
    if (formula === '=NOW()') {
      result.changed = true;
      result.useValue = true;
      result.value = new Date();
      return result;
    }
    
    // For RAND() and RANDBETWEEN(), suggest using static values
    if (formula.match(/^=RAND\(\)$/)) {
      result.changed = true;
      result.useValue = true;
      result.value = Math.random();
      return result;
    }
    
    // For complex formulas with volatile functions, try to cache results
    // This would need more sophisticated analysis
    
    return result;
  },
  
  /**
   * Fix #REF! errors by finding and suggesting replacements
   * @param {string} cellA1 - Cell with #REF! error
   * @return {Object} Fix suggestion
   */
  fixRefError: function(cellA1) {
    try {
      const sheet = SpreadsheetApp.getActiveSheet();
      const range = sheet.getRange(cellA1);
      const formula = range.getFormula();
      
      if (!formula) {
        return { success: false, error: 'No formula found' };
      }
      
      // Try to identify what the reference might have been
      // Look for patterns like Sheet1!#REF!
      const refPattern = /#REF!/g;
      
      // Check nearby cells for similar formulas
      const row = range.getRow();
      const col = range.getColumn();
      
      // Look at cell above
      if (row > 1) {
        const aboveFormula = sheet.getRange(row - 1, col).getFormula();
        if (aboveFormula && !aboveFormula.includes('#REF!')) {
          // Try to infer the pattern
          return {
            success: true,
            suggestion: 'Based on cell above, try: ' + this.inferReference(formula, aboveFormula)
          };
        }
      }
      
      // Look at cell to the left
      if (col > 1) {
        const leftFormula = sheet.getRange(row, col - 1).getFormula();
        if (leftFormula && !leftFormula.includes('#REF!')) {
          return {
            success: true,
            suggestion: 'Based on cell to left, try: ' + this.inferReference(formula, leftFormula)
          };
        }
      }
      
      return {
        success: false,
        error: 'Could not determine correct reference'
      };
      
    } catch (error) {
      Logger.error('Error fixing #REF!:', error);
      return { success: false, error: error.message };
    }
  },
  
  /**
   * Infer correct reference based on pattern
   * @param {string} brokenFormula - Formula with #REF!
   * @param {string} workingFormula - Similar working formula
   * @return {string} Suggested fix
   */
  inferReference: function(brokenFormula, workingFormula) {
    // This is a simplified pattern matching
    // Real implementation would need more sophisticated analysis
    
    // Extract cell references from working formula
    const refPattern = /[A-Z]+\d+/g;
    const refs = workingFormula.match(refPattern);
    
    if (refs && refs.length > 0) {
      // Replace first #REF! with first found reference
      return brokenFormula.replace('#REF!', refs[0]);
    }
    
    return brokenFormula;
  },

  /**
   * Import Excel data from pasted content
   * @param {string} pastedData - Data pasted from Excel
   * @return {Object} Parsed data with formulas and values
   */
  importExcelData: function(pastedData) {
    try {
      const lines = pastedData.split('\n').filter(line => line.trim());
      const data = [];
      const formulas = [];
      
      lines.forEach((line, rowIndex) => {
        const cells = line.split('\t');
        const rowData = [];
        const rowFormulas = [];
        
        cells.forEach((cell, colIndex) => {
          // Check if it's a formula (starts with =)
          if (cell.startsWith('=')) {
            rowFormulas.push(cell);
            // Try to evaluate or keep as formula
            rowData.push(cell);
          } else {
            rowFormulas.push('');
            // Parse value (number, date, or text)
            const parsed = this.parseExcelValue(cell);
            rowData.push(parsed);
          }
        });
        
        data.push(rowData);
        formulas.push(rowFormulas);
      });
      
      return {
        success: true,
        data: data,
        formulas: formulas,
        rows: data.length,
        cols: data[0] ? data[0].length : 0
      };
      
    } catch (error) {
      Logger.error('Error importing Excel data:', error);
      return { success: false, error: error.message };
    }
  },

  /**
   * Parse Excel value to appropriate type
   * @param {string} value - Cell value from Excel
   * @return {*} Parsed value
   */
  parseExcelValue: function(value) {
    // Check for Excel date serial number (days since 1900)
    const numValue = Number(value);
    if (!isNaN(numValue)) {
      // Check if it might be an Excel date (between 1 and ~45000)
      if (numValue > 0 && numValue < 50000 && value.indexOf('.') === -1) {
        // Might be a date - Excel dates start from 1900-01-01
        const excelEpoch = new Date(1899, 11, 30); // Excel's epoch
        const date = new Date(excelEpoch.getTime() + numValue * 24 * 60 * 60 * 1000);
        
        // Only convert if it results in a reasonable date
        if (date.getFullYear() > 1900 && date.getFullYear() < 2100) {
          return date;
        }
      }
      return numValue;
    }
    
    // Check for percentage (Excel exports as decimal)
    if (value.endsWith('%')) {
      const percent = parseFloat(value.slice(0, -1));
      if (!isNaN(percent)) {
        return percent / 100;
      }
    }
    
    // Return as string
    return value;
  },

  /**
   * Convert Excel formulas to Google Sheets formulas
   * @param {string} formula - Excel formula
   * @return {string} Google Sheets compatible formula
   */
  convertExcelFormula: function(formula) {
    let converted = formula;
    
    // Convert XLOOKUP to INDEX/MATCH
    if (converted.includes('XLOOKUP')) {
      converted = this.convertXlookupToVlookup(converted);
    }
    
    // Convert structured references
    if (converted.match(/\[@?\[.*?\]\]/)) {
      converted = this.convertStructuredRefs(converted);
    }
    
    // Fix locale issues (semicolon to comma)
    converted = converted.replace(/;/g, ',');
    
    // Convert IFERROR syntax if different
    converted = this.convertIferror(converted);
    
    // Convert CONCAT to CONCATENATE
    converted = converted.replace(/\bCONCAT\(/gi, 'CONCATENATE(');
    
    // Convert TEXTJOIN if not supported
    converted = this.convertTextjoin(converted);
    
    // Fix date functions
    converted = this.convertDateFunctions(converted);
    
    // Convert dynamic arrays (FILTER, UNIQUE, SORT, etc.)
    converted = this.convertDynamicArrays(converted);
    
    return converted;
  },

  /**
   * Convert IFERROR for compatibility
   * @param {string} formula - Formula containing IFERROR
   * @return {string} Converted formula
   */
  convertIferror: function(formula) {
    // IFERROR is supported in Google Sheets, but syntax might differ
    // Just ensure proper formatting
    return formula.replace(/IFERROR\s*\(/gi, 'IFERROR(');
  },

  /**
   * Convert TEXTJOIN function
   * @param {string} formula - Formula containing TEXTJOIN
   * @return {string} Converted formula
   */
  convertTextjoin: function(formula) {
    // TEXTJOIN(delimiter, ignore_empty, text1, [text2], ...)
    // Google Sheets supports TEXTJOIN, but older versions might not
    // For now, keep as is but could convert to JOIN or CONCATENATE if needed
    return formula;
  },

  /**
   * Convert Excel date functions to Google Sheets
   * @param {string} formula - Formula with date functions
   * @return {string} Converted formula
   */
  convertDateFunctions: function(formula) {
    // NETWORKDAYS.INTL to NETWORKDAYS
    formula = formula.replace(/NETWORKDAYS\.INTL/gi, 'NETWORKDAYS');
    
    // WORKDAY.INTL to WORKDAY
    formula = formula.replace(/WORKDAY\.INTL/gi, 'WORKDAY');
    
    // Fix DATEDIF (sometimes has issues)
    formula = formula.replace(/DATEDIF\(/gi, 'DATEDIF(');
    
    return formula;
  },

  /**
   * Convert dynamic array formulas
   * @param {string} formula - Formula with dynamic arrays
   * @return {string} Converted formula
   */
  convertDynamicArrays: function(formula) {
    // UNIQUE is supported in Google Sheets
    // FILTER is supported
    // SORT is supported
    // SEQUENCE needs conversion
    formula = formula.replace(/SEQUENCE\(/gi, 'SEQUENCE(');
    
    // RANDARRAY might need conversion
    if (formula.includes('RANDARRAY')) {
      // Convert RANDARRAY to ARRAYFORMULA with RAND
      // This is complex and would need parsing
      // For now, flag for manual review
      formula = '/* NEEDS REVIEW: RANDARRAY */ ' + formula;
    }
    
    return formula;
  },

  /**
   * Detect and convert pivot tables
   * @param {Range} range - Range that might contain pivot table
   * @return {Object} Pivot table configuration
   */
  detectPivotTable: function(range) {
    try {
      // Check if range looks like a pivot table
      const values = range.getValues();
      const formulas = range.getFormulas();
      
      // Pivot tables typically have:
      // - Headers in first row/column
      // - Aggregated values
      // - GETPIVOTDATA formulas referring to them
      
      let isPivot = false;
      let pivotConfig = {
        rows: [],
        columns: [],
        values: [],
        filters: []
      };
      
      // Check for GETPIVOTDATA formulas
      for (let row of formulas) {
        for (let formula of row) {
          if (formula.includes('GETPIVOTDATA')) {
            isPivot = true;
            // Parse GETPIVOTDATA to understand structure
            const match = formula.match(/GETPIVOTDATA\("([^"]+)"/);
            if (match) {
              pivotConfig.values.push(match[1]);
            }
          }
        }
      }
      
      if (isPivot) {
        // Analyze structure to find row/column fields
        // This is simplified - real implementation would be more complex
        if (values.length > 0 && values[0].length > 0) {
          // Assume first column contains row labels
          pivotConfig.rows.push('Field1');
          // Assume first row contains column labels
          pivotConfig.columns.push('Field2');
        }
      }
      
      return {
        isPivotTable: isPivot,
        config: pivotConfig,
        suggestion: isPivot ? 'Create Google Sheets Pivot Table with similar configuration' : null
      };
      
    } catch (error) {
      Logger.error('Error detecting pivot table:', error);
      return { isPivotTable: false };
    }
  },

  /**
   * Convert conditional formatting rules
   * @param {Sheet} sheet - Sheet to analyze
   * @return {Object} Conversion results
   */
  convertConditionalFormatting: function(sheet) {
    try {
      // Note: Google Apps Script doesn't have direct access to Excel's conditional formatting
      // This would work on Google Sheets conditional formatting that was imported
      
      const rules = sheet.getConditionalFormatRules();
      const converted = [];
      
      rules.forEach(rule => {
        const ranges = rule.getRanges();
        const condition = rule.getBooleanCondition();
        
        if (condition) {
          converted.push({
            ranges: ranges.map(r => r.getA1Notation()),
            type: condition.getCriteriaType(),
            values: condition.getCriteriaValues(),
            format: {
              bold: rule.getBold(),
              italic: rule.getItalic(),
              strike: rule.getStrikethrough(),
              foreground: rule.getFontColor(),
              background: rule.getBackground()
            }
          });
        }
      });
      
      return {
        success: true,
        rules: converted,
        count: converted.length
      };
      
    } catch (error) {
      Logger.error('Error converting conditional formatting:', error);
      return { success: false, error: error.message };
    }
  },

  /**
   * Detect VBA/Macros in imported content
   * @param {string} content - Content to check
   * @return {Object} Detection results
   */
  detectVBAMacros: function(content) {
    const vbaIndicators = [
      'Sub ', 'Function ', 'End Sub', 'End Function',
      'Dim ', 'Set ', 'Private ', 'Public ',
      'ActiveSheet', 'ActiveWorkbook', 'Range(',
      'Cells(', 'Worksheets(', '.Select', '.Activate'
    ];
    
    const detected = [];
    
    vbaIndicators.forEach(indicator => {
      if (content.includes(indicator)) {
        detected.push(indicator);
      }
    });
    
    return {
      hasVBA: detected.length > 0,
      indicators: detected,
      warning: detected.length > 0 ? 
        'VBA macros detected. These need to be converted to Google Apps Script.' : null,
      suggestion: detected.length > 0 ?
        'Consider using Google Apps Script for automation instead of VBA macros.' : null
    };
  },

  /**
   * Create migration report
   * @param {Object} migrationResults - Results from migration
   * @return {Object} Formatted report
   */
  createMigrationReport: function(migrationResults) {
    const report = {
      timestamp: new Date(),
      summary: {
        totalCells: 0,
        formulasConverted: 0,
        errorsFixed: 0,
        warningsGenerated: 0
      },
      details: [],
      recommendations: []
    };
    
    // Compile results
    if (migrationResults.scan) {
      const scan = migrationResults.scan;
      report.summary.totalCells = scan.totalCells || 0;
      report.summary.formulasConverted = scan.issues.xlookupFormulas.length + 
                                          scan.issues.structuredRefs.length;
      report.summary.errorsFixed = scan.issues.formulaErrors.length + 
                                   scan.issues.refErrors.length;
      report.summary.warningsGenerated = scan.issues.volatileFunctions.length;
      
      // Add details
      if (scan.issues.xlookupFormulas.length > 0) {
        report.details.push({
          type: 'XLOOKUP Conversions',
          count: scan.issues.xlookupFormulas.length,
          items: scan.issues.xlookupFormulas.slice(0, 5)
        });
      }
      
      if (scan.issues.volatileFunctions.length > 0) {
        report.details.push({
          type: 'Volatile Functions',
          count: scan.issues.volatileFunctions.length,
          items: scan.issues.volatileFunctions.slice(0, 5)
        });
      }
    }
    
    // Add recommendations
    if (report.summary.warningsGenerated > 50) {
      report.recommendations.push('Consider replacing volatile functions with static values for better performance');
    }
    
    if (report.summary.formulasConverted > 100) {
      report.recommendations.push('Review converted formulas to ensure they work as expected');
    }
    
    return report;
  }
};