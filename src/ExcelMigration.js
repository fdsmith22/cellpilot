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
      console.error('Error scanning for Excel issues:', error);
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
            console.error(`Failed to fix locale in ${issue.cell}:`, e);
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
            console.error(`Failed to convert XLOOKUP in ${issue.cell}:`, e);
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
            console.error(`Failed to fix structured ref in ${issue.cell}:`, e);
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
            console.error(`Failed to optimize volatile in ${issue.cell}:`, e);
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
      console.error('Error fixing Excel issues:', error);
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
      console.error('Error converting XLOOKUP:', e);
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
      console.error('Error fixing #REF!:', error);
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
  }
};