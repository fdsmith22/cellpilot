/**
 * Smart Formula Debugger
 * Advanced formula analysis and debugging capabilities
 */

const FormulaDebugger = {
  
  /**
   * Show the formula debugger interface
   */
  showSmartFormulaDebugger: function() {
    try {
      const html = HtmlService.createTemplateFromFile('SmartFormulaDebuggerTemplate')
        .evaluate()
        .setTitle('Smart Formula Debugger')
        .setWidth(450);
      
      SpreadsheetApp.getUi().showSidebar(html);
    } catch (error) {
      Logger.error('Error showing formula debugger:', error);
      SpreadsheetApp.getUi().alert('Failed to open Formula Debugger: ' + error.message);
    }
  },
  
  /**
   * Debug the active cell's formula
   */
  debugActiveFormula: function() {
    try {
      const cell = SpreadsheetApp.getActiveCell();
      const formula = cell.getFormula();
      
      if (!formula || !formula.startsWith('=')) {
        return {
          success: false,
          error: 'No formula found in active cell'
        };
      }
      
      const debugData = this.analyzeFormula(formula, cell.getA1Notation());
      
      return {
        success: true,
        data: debugData
      };
      
    } catch (error) {
      Logger.error('Error debugging active formula:', error);
      return {
        success: false,
        error: error.message
      };
    }
  },
  
  /**
   * Debug all formulas in selection
   */
  debugSelectionFormulas: function() {
    try {
      const range = SpreadsheetApp.getActiveRange();
      const formulas = range.getFormulas();
      const startRow = range.getRow();
      const startCol = range.getColumn();
      
      let allIssues = [];
      let errorCount = 0;
      let warningCount = 0;
      let validCount = 0;
      let totalFormulas = 0;
      
      formulas.forEach((row, rowIndex) => {
        row.forEach((formula, colIndex) => {
          if (formula && formula.startsWith('=')) {
            totalFormulas++;
            const cell = this.getCellNotation(startRow + rowIndex, startCol + colIndex);
            const analysis = this.analyzeFormula(formula, cell);
            
            if (analysis.issues && analysis.issues.length > 0) {
              allIssues.push(...analysis.issues);
              errorCount += analysis.errorCount || 0;
              warningCount += analysis.warningCount || 0;
            } else {
              validCount++;
            }
          }
        });
      });
      
      return {
        success: true,
        data: {
          issues: allIssues,
          errorCount: errorCount,
          warningCount: warningCount,
          validCount: validCount,
          totalFormulas: totalFormulas
        }
      };
      
    } catch (error) {
      Logger.error('Error debugging selection:', error);
      return {
        success: false,
        error: error.message
      };
    }
  },
  
  /**
   * Debug all formulas in the active sheet
   */
  debugAllSheetFormulas: function() {
    try {
      const sheet = SpreadsheetApp.getActiveSheet();
      const lastRow = sheet.getLastRow();
      const lastColumn = sheet.getLastColumn();
      
      if (lastRow === 0 || lastColumn === 0) {
        return {
          success: false,
          error: 'Sheet is empty'
        };
      }
      
      const formulas = sheet.getRange(1, 1, lastRow, lastColumn).getFormulas();
      
      let allIssues = [];
      let errorCount = 0;
      let warningCount = 0;
      let validCount = 0;
      let totalFormulas = 0;
      
      formulas.forEach((row, rowIndex) => {
        row.forEach((formula, colIndex) => {
          if (formula && formula.startsWith('=')) {
            totalFormulas++;
            const cell = this.getCellNotation(rowIndex + 1, colIndex + 1);
            const analysis = this.analyzeFormula(formula, cell);
            
            if (analysis.issues && analysis.issues.length > 0) {
              allIssues.push(...analysis.issues);
              errorCount += analysis.errorCount || 0;
              warningCount += analysis.warningCount || 0;
            } else {
              validCount++;
            }
          }
        });
      });
      
      return {
        success: true,
        data: {
          issues: allIssues,
          errorCount: errorCount,
          warningCount: warningCount,
          validCount: validCount,
          totalFormulas: totalFormulas
        }
      };
      
    } catch (error) {
      Logger.error('Error debugging all formulas:', error);
      return {
        success: false,
        error: error.message
      };
    }
  },
  
  /**
   * Analyze a single formula for issues
   */
  analyzeFormula: function(formula, cellRef) {
    const issues = [];
    let errorCount = 0;
    let warningCount = 0;
    
    // Check for common errors
    const errorChecks = [
      this.checkFormulaErrors(formula, cellRef),
      this.checkCircularReference(formula, cellRef),
      this.checkInvalidReferences(formula, cellRef),
      this.checkDeprecatedFunctions(formula, cellRef),
      this.checkPerformanceIssues(formula, cellRef),
      this.checkSyntaxErrors(formula, cellRef),
      this.checkVolatileFunctions(formula, cellRef)
    ];
    
    errorChecks.forEach(check => {
      if (check && check.issue) {
        issues.push(check.issue);
        if (check.issue.severity === 'error') errorCount++;
        if (check.issue.severity === 'warning') warningCount++;
      }
    });
    
    return {
      issues: issues,
      errorCount: errorCount,
      warningCount: warningCount,
      formula: formula,
      cell: cellRef
    };
  },
  
  /**
   * Check for formula errors
   */
  checkFormulaErrors: function(formula, cellRef) {
    try {
      // Common Excel to Sheets incompatibilities
      const incompatibilities = [
        { pattern: /XLOOKUP/i, message: 'XLOOKUP is not available in Google Sheets. Use VLOOKUP or INDEX/MATCH instead.' },
        { pattern: /CONCATENATEX/i, message: 'CONCATENATEX is not available. Use CONCATENATE or & operator.' },
        { pattern: /TEXTJOIN/i, version: 'warning', message: 'TEXTJOIN may not work as expected. Consider using JOIN or CONCATENATE.' }
      ];
      
      for (let check of incompatibilities) {
        if (check.pattern.test(formula)) {
          return {
            issue: {
              type: 'Compatibility Error',
              severity: check.version || 'error',
              message: check.message,
              cell: cellRef,
              formula: formula,
              suggestions: this.generateCompatibilitySuggestions(formula, check.pattern)
            }
          };
        }
      }
      
      // Check for #REF!, #VALUE!, #NAME? errors
      if (formula.includes('#REF!')) {
        return {
          issue: {
            type: 'Reference Error',
            severity: 'error',
            message: 'Formula contains invalid cell reference (#REF!)',
            cell: cellRef,
            formula: formula,
            errorPosition: { start: formula.indexOf('#REF!'), end: formula.indexOf('#REF!') + 5 },
            suggestions: [{
              description: 'Check if referenced cells were deleted',
              formula: formula.replace(/#REF!/g, 'A1')
            }]
          }
        };
      }
      
      if (formula.includes('#VALUE!')) {
        return {
          issue: {
            type: 'Value Error',
            severity: 'error',
            message: 'Formula has wrong type of argument (#VALUE!)',
            cell: cellRef,
            formula: formula,
            suggestions: [{
              description: 'Check data types in referenced cells',
              formula: formula
            }]
          }
        };
      }
      
      if (formula.includes('#NAME?')) {
        return {
          issue: {
            type: 'Name Error',
            severity: 'error',
            message: 'Formula contains unrecognized function or name (#NAME?)',
            cell: cellRef,
            formula: formula,
            suggestions: this.suggestFunctionCorrections(formula)
          }
        };
      }
      
    } catch (error) {
      Logger.error('Error checking formula errors:', error);
    }
    
    return null;
  },
  
  /**
   * Check for circular references
   */
  checkCircularReference: function(formula, cellRef) {
    try {
      // Extract all cell references from formula
      const references = this.extractCellReferences(formula);
      
      // Check if any reference points back to the current cell
      for (let ref of references) {
        if (ref === cellRef) {
          return {
            issue: {
              type: 'Circular Reference',
              severity: 'error',
              message: 'Formula creates a circular reference by referencing its own cell',
              cell: cellRef,
              formula: formula,
              errorPosition: { start: formula.indexOf(ref), end: formula.indexOf(ref) + ref.length },
              suggestions: [{
                description: 'Remove self-reference',
                formula: formula.replace(new RegExp(ref, 'g'), '')
              }]
            }
          };
        }
      }
      
      // Check for indirect circular references (limited to 2 levels deep)
      const indirectRefs = this.checkIndirectCircularReferences(formula, cellRef, 2);
      if (indirectRefs) {
        return indirectRefs;
      }
      
    } catch (error) {
      Logger.error('Error checking circular reference:', error);
    }
    
    return null;
  },
  
  /**
   * Check for indirect circular references
   */
  checkIndirectCircularReferences: function(formula, cellRef, maxDepth = 2) {
    try {
      const sheet = SpreadsheetApp.getActiveSheet();
      const refs = this.extractReferences(formula);
      const visited = new Set([cellRef]);
      
      const checkDepth = (ref, depth) => {
        if (depth > maxDepth) return null;
        if (visited.has(ref)) {
          return {
            type: 'error',
            category: 'circular_reference',
            message: `Indirect circular reference detected: ${cellRef} -> ... -> ${ref}`,
            severity: 'error',
            fixes: [{
              description: 'Remove circular dependency',
              formula: formula
            }]
          };
        }
        
        visited.add(ref);
        
        try {
          const targetCell = sheet.getRange(ref);
          const targetFormula = targetCell.getFormula();
          if (targetFormula) {
            const targetRefs = this.extractReferences(targetFormula);
            for (const targetRef of targetRefs) {
              const result = checkDepth(targetRef, depth + 1);
              if (result) return result;
            }
          }
        } catch (e) {
          // Invalid reference, skip
        }
        
        return null;
      };
      
      for (const ref of refs) {
        const result = checkDepth(ref, 1);
        if (result) return result;
      }
      
    } catch (error) {
      Logger.error('Error checking indirect circular references:', error);
    }
    
    return null;
  },
  
  /**
   * Check for invalid cell references
   */
  checkInvalidReferences: function(formula, cellRef) {
    try {
      // Check for references outside sheet bounds
      const references = this.extractCellReferences(formula);
      const sheet = SpreadsheetApp.getActiveSheet();
      const maxRows = sheet.getMaxRows();
      const maxCols = sheet.getMaxColumns();
      
      for (let ref of references) {
        const parsed = this.parseCellReference(ref);
        if (parsed) {
          if (parsed.row > maxRows || parsed.col > maxCols) {
            return {
              issue: {
                type: 'Invalid Reference',
                severity: 'warning',
                message: `Reference ${ref} is outside sheet bounds (max: ${maxRows} rows, ${maxCols} cols)`,
                cell: cellRef,
                formula: formula,
                errorPosition: { start: formula.indexOf(ref), end: formula.indexOf(ref) + ref.length }
              }
            };
          }
        }
      }
      
    } catch (error) {
      Logger.error('Error checking invalid references:', error);
    }
    
    return null;
  },
  
  /**
   * Check for deprecated functions
   */
  checkDeprecatedFunctions: function(formula, cellRef) {
    const deprecated = [
      { func: 'DATEDIF', replacement: 'DAYS', message: 'DATEDIF is deprecated. Use DAYS or other date functions.' },
      { func: 'ROMAN', replacement: '', message: 'ROMAN function has limited use. Consider alternatives.' }
    ];
    
    for (let dep of deprecated) {
      if (formula.includes(dep.func)) {
        return {
          issue: {
            type: 'Deprecated Function',
            severity: 'warning',
            message: dep.message,
            cell: cellRef,
            formula: formula,
            suggestions: dep.replacement ? [{
              description: `Replace with ${dep.replacement}`,
              formula: formula.replace(new RegExp(dep.func, 'g'), dep.replacement)
            }] : []
          }
        };
      }
    }
    
    return null;
  },
  
  /**
   * Check for performance issues
   */
  checkPerformanceIssues: function(formula, cellRef) {
    try {
      // Check for array formulas on large ranges
      if (formula.includes('ARRAYFORMULA')) {
        const rangePattern = /([A-Z]+):([A-Z]+)/;
        const match = formula.match(rangePattern);
        if (match) {
          return {
            issue: {
              type: 'Performance',
              severity: 'warning',
              message: 'ARRAYFORMULA on entire columns can slow down sheet',
              cell: cellRef,
              formula: formula,
              suggestions: [{
                description: 'Limit range to necessary rows',
                formula: formula.replace(rangePattern, '$1$1:$2$1000')
              }]
            }
          };
        }
      }
      
      // Check for nested IF statements
      const ifCount = (formula.match(/\bIF\(/gi) || []).length;
      if (ifCount > 3) {
        return {
          issue: {
            type: 'Performance',
            severity: 'warning',
            message: `Formula has ${ifCount} nested IF statements. Consider using IFS or SWITCH`,
            cell: cellRef,
            formula: formula,
            suggestions: [{
              description: 'Use IFS function for multiple conditions',
              formula: this.convertToIFS(formula)
            }]
          }
        };
      }
      
      // Check for volatile functions in large formulas
      const volatileFuncs = ['NOW', 'TODAY', 'RAND', 'RANDBETWEEN'];
      for (let func of volatileFuncs) {
        if (formula.includes(func) && formula.length > 100) {
          return {
            issue: {
              type: 'Performance',
              severity: 'warning',
              message: `Volatile function ${func} recalculates on every change`,
              cell: cellRef,
              formula: formula,
              suggestions: [{
                description: 'Consider using static values or reducing recalculation frequency',
                formula: formula
              }]
            }
          };
        }
      }
      
    } catch (error) {
      Logger.error('Error checking performance:', error);
    }
    
    return null;
  },
  
  /**
   * Check for syntax errors
   */
  checkSyntaxErrors: function(formula, cellRef) {
    try {
      // Check for unbalanced parentheses
      let parenCount = 0;
      for (let char of formula) {
        if (char === '(') parenCount++;
        if (char === ')') parenCount--;
        if (parenCount < 0) {
          return {
            issue: {
              type: 'Syntax Error',
              severity: 'error',
              message: 'Unmatched closing parenthesis',
              cell: cellRef,
              formula: formula,
              suggestions: [{
                description: 'Add opening parenthesis',
                formula: '(' + formula
              }]
            }
          };
        }
      }
      
      if (parenCount > 0) {
        return {
          issue: {
            type: 'Syntax Error',
            severity: 'error',
            message: `Missing ${parenCount} closing parenthesis`,
            cell: cellRef,
            formula: formula,
            suggestions: [{
              description: 'Add closing parentheses',
              formula: formula + ')'.repeat(parenCount)
            }]
          }
        };
      }
      
      // Check for missing commas
      if (/\)\s+[A-Z]/.test(formula)) {
        return {
          issue: {
            type: 'Syntax Error',
            severity: 'warning',
            message: 'Possible missing comma between arguments',
            cell: cellRef,
            formula: formula
          }
        };
      }
      
    } catch (error) {
      Logger.error('Error checking syntax:', error);
    }
    
    return null;
  },
  
  /**
   * Check for volatile functions
   */
  checkVolatileFunctions: function(formula, cellRef) {
    const volatileFuncs = [
      { name: 'NOW', impact: 'high' },
      { name: 'TODAY', impact: 'medium' },
      { name: 'RAND', impact: 'high' },
      { name: 'RANDBETWEEN', impact: 'high' },
      { name: 'OFFSET', impact: 'medium' },
      { name: 'INDIRECT', impact: 'medium' }
    ];
    
    for (let func of volatileFuncs) {
      if (formula.includes(func.name)) {
        return {
          issue: {
            type: 'Volatile Function',
            severity: 'info',
            message: `${func.name} recalculates on every change (${func.impact} impact)`,
            cell: cellRef,
            formula: formula
          }
        };
      }
    }
    
    return null;
  },
  
  /**
   * Apply a formula fix to a cell
   */
  applyFormulaFix: function(cellRef, newFormula) {
    try {
      const sheet = SpreadsheetApp.getActiveSheet();
      const range = sheet.getRange(cellRef);
      range.setFormula(newFormula);
      
      return {
        success: true,
        cell: cellRef,
        formula: newFormula
      };
      
    } catch (error) {
      Logger.error('Error applying fix:', error);
      return {
        success: false,
        error: error.message
      };
    }
  },
  
  /**
   * Analyze formula dependencies
   */
  analyzeDependencies: function() {
    try {
      const cell = SpreadsheetApp.getActiveCell();
      const formula = cell.getFormula();
      
      if (!formula || !formula.startsWith('=')) {
        return {
          success: false,
          error: 'No formula in active cell'
        };
      }
      
      const dependencies = this.extractDependencies(formula, cell.getA1Notation());
      
      return {
        success: true,
        dependencies: dependencies
      };
      
    } catch (error) {
      Logger.error('Error analyzing dependencies:', error);
      return {
        success: false,
        error: error.message
      };
    }
  },
  
  /**
   * Extract dependencies from formula
   */
  extractDependencies: function(formula, cellRef) {
    const references = this.extractCellReferences(formula);
    const dependencies = [];
    
    references.forEach(ref => {
      if (ref !== cellRef) { // Avoid self-reference
        dependencies.push({
          cell: cellRef,
          reference: ref,
          type: this.getReferenceType(ref)
        });
      }
    });
    
    return dependencies;
  },
  
  /**
   * Analyze formula performance
   */
  analyzeFormulaPerformance: function() {
    try {
      const cell = SpreadsheetApp.getActiveCell();
      const formula = cell.getFormula();
      
      if (!formula || !formula.startsWith('=')) {
        return {
          success: false,
          error: 'No formula in active cell'
        };
      }
      
      const analysis = this.performPerformanceAnalysis(formula);
      
      return {
        success: true,
        analysis: analysis
      };
      
    } catch (error) {
      Logger.error('Error analyzing performance:', error);
      return {
        success: false,
        error: error.message
      };
    }
  },
  
  /**
   * Perform detailed performance analysis
   */
  performPerformanceAnalysis: function(formula) {
    let score = 100;
    const suggestions = [];
    
    // Check formula complexity
    if (formula.length > 200) {
      score -= 20;
      suggestions.push({
        description: 'Formula is very long. Consider breaking into helper columns',
        improvement: 20
      });
    }
    
    // Check for volatile functions
    const volatileCount = (formula.match(/\b(NOW|TODAY|RAND|RANDBETWEEN|OFFSET|INDIRECT)\b/gi) || []).length;
    if (volatileCount > 0) {
      score -= volatileCount * 10;
      suggestions.push({
        description: `Remove ${volatileCount} volatile function(s) for better performance`,
        improvement: volatileCount * 10
      });
    }
    
    // Check for full column references
    if (/[A-Z]+:[A-Z]+/.test(formula)) {
      score -= 15;
      suggestions.push({
        description: 'Avoid full column references. Use specific ranges',
        improvement: 15
      });
    }
    
    // Check for array formulas
    if (formula.includes('ARRAYFORMULA')) {
      score -= 10;
      suggestions.push({
        description: 'ARRAYFORMULA can be slow on large ranges',
        improvement: 10
      });
    }
    
    // Check for nested functions
    const nestedLevel = this.countNestingLevel(formula);
    if (nestedLevel > 5) {
      score -= 15;
      suggestions.push({
        description: 'Deeply nested formula. Consider simplifying',
        improvement: 15
      });
    }
    
    return {
      score: Math.max(0, score),
      suggestions: suggestions
    };
  },
  
  /**
   * Helper function to extract cell references
   */
  extractCellReferences: function(formula) {
    const references = [];
    const pattern = /[A-Z]+[0-9]+(?::[A-Z]+[0-9]+)?/g;
    let match;
    
    while ((match = pattern.exec(formula)) !== null) {
      references.push(match[0]);
    }
    
    return references;
  },
  
  /**
   * Helper function to parse cell reference
   */
  parseCellReference: function(ref) {
    const match = ref.match(/^([A-Z]+)([0-9]+)$/);
    if (match) {
      const col = this.columnLetterToNumber(match[1]);
      const row = parseInt(match[2]);
      return { col: col, row: row };
    }
    return null;
  },
  
  /**
   * Convert column letter to number
   */
  columnLetterToNumber: function(letter) {
    let result = 0;
    for (let i = 0; i < letter.length; i++) {
      result = result * 26 + (letter.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
    }
    return result;
  },
  
  /**
   * Get cell notation from row and column
   */
  getCellNotation: function(row, col) {
    let letter = '';
    while (col > 0) {
      const temp = (col - 1) % 26;
      letter = String.fromCharCode(temp + 65) + letter;
      col = (col - temp - 1) / 26;
    }
    return letter + row;
  },
  
  /**
   * Generate compatibility suggestions
   */
  generateCompatibilitySuggestions: function(formula, pattern) {
    const suggestions = [];
    
    if (pattern.test('XLOOKUP')) {
      suggestions.push({
        description: 'Replace with VLOOKUP',
        formula: formula.replace(/XLOOKUP/gi, 'VLOOKUP')
      });
      suggestions.push({
        description: 'Replace with INDEX/MATCH',
        formula: formula.replace(/XLOOKUP\([^,]+,[^,]+,[^,]+\)/gi, 'INDEX(range,MATCH(lookup,range,0))')
      });
    }
    
    return suggestions;
  },
  
  /**
   * Suggest function corrections
   */
  suggestFunctionCorrections: function(formula) {
    const suggestions = [];
    
    // Common typos
    const corrections = {
      'VLOKUP': 'VLOOKUP',
      'SUMIF': 'SUMIF',
      'COUNTIFF': 'COUNTIF',
      'AVERGE': 'AVERAGE'
    };
    
    for (let [wrong, correct] of Object.entries(corrections)) {
      if (formula.includes(wrong)) {
        suggestions.push({
          description: `Correct ${wrong} to ${correct}`,
          formula: formula.replace(new RegExp(wrong, 'gi'), correct)
        });
      }
    }
    
    return suggestions;
  },
  
  /**
   * Convert nested IFs to IFS function
   */
  convertToIFS: function(formula) {
    // Simplified conversion - would need more complex parsing for real implementation
    return formula.replace(/IF\(/g, 'IFS(');
  },
  
  /**
   * Get reference type
   */
  getReferenceType: function(ref) {
    if (ref.includes(':')) return 'range';
    if (ref.includes('!')) return 'cross-sheet';
    return 'cell';
  },
  
  /**
   * Count nesting level in formula
   */
  countNestingLevel: function(formula) {
    let maxLevel = 0;
    let currentLevel = 0;
    
    for (let char of formula) {
      if (char === '(') {
        currentLevel++;
        maxLevel = Math.max(maxLevel, currentLevel);
      } else if (char === ')') {
        currentLevel--;
      }
    }
    
    return maxLevel;
  }
};