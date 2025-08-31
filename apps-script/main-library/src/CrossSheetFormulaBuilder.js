/**
 * Cross-Sheet Formula Builder
 * Advanced formula creation for multi-tab spreadsheets with relationship awareness
 */

const CrossSheetFormulaBuilder = {
  
  /**
   * Analyze spreadsheet structure and relationships
   */
  analyzeSpreadsheetStructure: function() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = spreadsheet.getSheets();
    const structure = {
      sheets: [],
      relationships: [],
      commonColumns: {},
      suggestedJoins: []
    };
    
    // Analyze each sheet
    sheets.forEach(sheet => {
      const sheetInfo = {
        name: sheet.getName(),
        columns: [],
        dataTypes: {},
        rowCount: sheet.getLastRow(),
        columnCount: sheet.getLastColumn(),
        namedRanges: []
      };
      
      // Get headers and analyze data types
      if (sheetInfo.rowCount > 0 && sheetInfo.columnCount > 0) {
        const headers = sheet.getRange(1, 1, 1, sheetInfo.columnCount).getValues()[0];
        const sampleData = sheet.getRange(2, 1, Math.min(10, sheetInfo.rowCount - 1), sheetInfo.columnCount).getValues();
        
        headers.forEach((header, index) => {
          if (header) {
            sheetInfo.columns.push(header);
            sheetInfo.dataTypes[header] = this.detectDataType(sampleData.map(row => row[index]));
          }
        });
      }
      
      // Get named ranges for this sheet
      const namedRanges = spreadsheet.getNamedRanges();
      namedRanges.forEach(namedRange => {
        if (namedRange.getRange().getSheet().getName() === sheet.getName()) {
          sheetInfo.namedRanges.push({
            name: namedRange.getName(),
            range: namedRange.getRange().getA1Notation()
          });
        }
      });
      
      structure.sheets.push(sheetInfo);
    });
    
    // Detect relationships between sheets
    structure.relationships = this.detectRelationships(structure.sheets);
    
    // Find common columns across sheets
    structure.commonColumns = this.findCommonColumns(structure.sheets);
    
    // Suggest potential joins
    structure.suggestedJoins = this.suggestJoins(structure.sheets, structure.commonColumns);
    
    return structure;
  },
  
  /**
   * Detect data type from sample values
   */
  detectDataType: function(samples) {
    const types = {
      number: 0,
      date: 0,
      text: 0,
      currency: 0,
      percentage: 0,
      email: 0,
      id: 0
    };
    
    samples.forEach(value => {
      if (value === null || value === '') return;
      
      const strValue = String(value);
      
      // Check for specific patterns
      if (/^[A-Z]{2,}-\d+$/.test(strValue)) {
        types.id++;
      } else if (/^[\w\.-]+@[\w\.-]+\.\w+$/.test(strValue)) {
        types.email++;
      } else if (/^\$?[\d,]+\.?\d*$/.test(strValue)) {
        types.currency++;
      } else if (/^\d+\.?\d*%$/.test(strValue)) {
        types.percentage++;
      } else if (value instanceof Date) {
        types.date++;
      } else if (!isNaN(value)) {
        types.number++;
      } else {
        types.text++;
      }
    });
    
    // Return the most common type
    return Object.keys(types).reduce((a, b) => types[a] > types[b] ? a : b);
  },
  
  /**
   * Detect relationships between sheets based on column names and data
   */
  detectRelationships: function(sheets) {
    const relationships = [];
    
    for (let i = 0; i < sheets.length; i++) {
      for (let j = i + 1; j < sheets.length; j++) {
        const sheet1 = sheets[i];
        const sheet2 = sheets[j];
        
        // Check for matching ID columns
        const idColumns1 = sheet1.columns.filter(col => 
          col.toLowerCase().includes('id') || 
          col.toLowerCase().includes('key') ||
          col.toLowerCase().includes('code')
        );
        
        const idColumns2 = sheet2.columns.filter(col => 
          col.toLowerCase().includes('id') || 
          col.toLowerCase().includes('key') ||
          col.toLowerCase().includes('code')
        );
        
        // Check for exact matches
        idColumns1.forEach(col1 => {
          idColumns2.forEach(col2 => {
            if (col1.toLowerCase() === col2.toLowerCase()) {
              relationships.push({
                type: 'one-to-many',
                sheet1: sheet1.name,
                column1: col1,
                sheet2: sheet2.name,
                column2: col2,
                confidence: 'high'
              });
            }
          });
        });
        
        // Check for partial matches (e.g., "Customer ID" and "ID")
        sheet1.columns.forEach(col1 => {
          sheet2.columns.forEach(col2 => {
            const similarity = this.calculateSimilarity(col1, col2);
            if (similarity > 0.7 && !relationships.some(r => 
              r.column1 === col1 && r.column2 === col2)) {
              relationships.push({
                type: 'possible',
                sheet1: sheet1.name,
                column1: col1,
                sheet2: sheet2.name,
                column2: col2,
                confidence: similarity > 0.85 ? 'medium' : 'low',
                similarity: similarity
              });
            }
          });
        });
      }
    }
    
    return relationships;
  },
  
  /**
   * Calculate similarity between two column names
   */
  calculateSimilarity: function(str1, str2) {
    str1 = str1.toLowerCase().replace(/[_\s-]/g, '');
    str2 = str2.toLowerCase().replace(/[_\s-]/g, '');
    
    if (str1 === str2) return 1;
    if (str1.includes(str2) || str2.includes(str1)) return 0.8;
    
    // Levenshtein distance
    const matrix = [];
    for (let i = 0; i <= str2.length; i++) {
      matrix[i] = [i];
    }
    for (let j = 0; j <= str1.length; j++) {
      matrix[0][j] = j;
    }
    
    for (let i = 1; i <= str2.length; i++) {
      for (let j = 1; j <= str1.length; j++) {
        if (str2.charAt(i - 1) === str1.charAt(j - 1)) {
          matrix[i][j] = matrix[i - 1][j - 1];
        } else {
          matrix[i][j] = Math.min(
            matrix[i - 1][j - 1] + 1,
            matrix[i][j - 1] + 1,
            matrix[i - 1][j] + 1
          );
        }
      }
    }
    
    const distance = matrix[str2.length][str1.length];
    const maxLength = Math.max(str1.length, str2.length);
    return 1 - (distance / maxLength);
  },
  
  /**
   * Find common columns across sheets
   */
  findCommonColumns: function(sheets) {
    const columnCounts = {};
    
    sheets.forEach(sheet => {
      sheet.columns.forEach(col => {
        const normalizedCol = col.toLowerCase().trim();
        if (!columnCounts[normalizedCol]) {
          columnCounts[normalizedCol] = {
            originalNames: [],
            sheets: []
          };
        }
        columnCounts[normalizedCol].originalNames.push(col);
        columnCounts[normalizedCol].sheets.push(sheet.name);
      });
    });
    
    // Filter to only columns that appear in multiple sheets
    const commonColumns = {};
    Object.keys(columnCounts).forEach(col => {
      if (columnCounts[col].sheets.length > 1) {
        commonColumns[col] = columnCounts[col];
      }
    });
    
    return commonColumns;
  },
  
  /**
   * Suggest potential joins between sheets
   */
  suggestJoins: function(sheets, commonColumns) {
    const suggestions = [];
    
    Object.keys(commonColumns).forEach(col => {
      const info = commonColumns[col];
      if (info.sheets.length === 2) {
        suggestions.push({
          type: 'VLOOKUP',
          description: `Lookup ${info.originalNames[0]} from ${info.sheets[0]} in ${info.sheets[1]}`,
          formula: `=VLOOKUP(A2, '${info.sheets[1]}'!A:Z, COLUMN_INDEX, FALSE)`,
          confidence: 'high'
        });
      } else if (info.sheets.length > 2) {
        suggestions.push({
          type: 'QUERY',
          description: `Join data from ${info.sheets.join(', ')} using ${info.originalNames[0]}`,
          formula: `=QUERY({${info.sheets.map(s => `'${s}'!A:Z`).join(';')}}, "SELECT * WHERE Col1 IS NOT NULL")`,
          confidence: 'medium'
        });
      }
    });
    
    return suggestions;
  },
  
  /**
   * Generate complex cross-sheet formula
   */
  generateCrossSheetFormula: function(params) {
    const {
      operation,
      sourceSheet,
      sourceColumn,
      targetSheet,
      targetColumn,
      criteria,
      aggregation,
      joinType
    } = params;
    
    let formula = '';
    
    switch (operation) {
      case 'lookup':
        formula = this.generateLookupFormula(params);
        break;
      case 'aggregate':
        formula = this.generateAggregateFormula(params);
        break;
      case 'filter':
        formula = this.generateFilterFormula(params);
        break;
      case 'join':
        formula = this.generateJoinFormula(params);
        break;
      case 'pivot':
        formula = this.generatePivotFormula(params);
        break;
      default:
        formula = '=ERROR("Unknown operation")';
    }
    
    return {
      formula: formula,
      explanation: this.explainFormula(formula, params),
      validation: this.validateFormula(formula, params)
    };
  },
  
  /**
   * Generate VLOOKUP or INDEX/MATCH formula
   */
  generateLookupFormula: function(params) {
    const { sourceSheet, sourceColumn, targetSheet, targetColumn, returnColumn, useIndexMatch } = params;
    
    if (useIndexMatch) {
      // INDEX/MATCH is more flexible and faster for large datasets
      return `=INDEX('${targetSheet}'!${returnColumn}:${returnColumn}, MATCH(${sourceColumn}2, '${targetSheet}'!${targetColumn}:${targetColumn}, 0))`;
    } else {
      // Standard VLOOKUP
      const targetRange = `'${targetSheet}'!A:Z`; // Adjust based on actual range
      const columnIndex = 2; // This should be calculated based on actual column positions
      return `=VLOOKUP(${sourceColumn}2, ${targetRange}, ${columnIndex}, FALSE)`;
    }
  },
  
  /**
   * Generate aggregation formula (SUMIFS, COUNTIFS, AVERAGEIFS)
   */
  generateAggregateFormula: function(params) {
    const { aggregation, targetSheet, sumColumn, criteriaRanges, criteriaValues } = params;
    
    switch (aggregation) {
      case 'sum':
        return `=SUMIFS('${targetSheet}'!${sumColumn}:${sumColumn}, ${criteriaRanges.map((range, i) => 
          `'${targetSheet}'!${range}:${range}, "${criteriaValues[i]}"`).join(', ')})`;
      case 'count':
        return `=COUNTIFS(${criteriaRanges.map((range, i) => 
          `'${targetSheet}'!${range}:${range}, "${criteriaValues[i]}"`).join(', ')})`;
      case 'average':
        return `=AVERAGEIFS('${targetSheet}'!${sumColumn}:${sumColumn}, ${criteriaRanges.map((range, i) => 
          `'${targetSheet}'!${range}:${range}, "${criteriaValues[i]}"`).join(', ')})`;
      default:
        return '=ERROR("Unknown aggregation")';
    }
  },
  
  /**
   * Generate FILTER formula
   */
  generateFilterFormula: function(params) {
    const { sourceSheet, sourceRange, conditions } = params;
    
    const filterConditions = conditions.map(cond => 
      `'${sourceSheet}'!${cond.column}:${cond.column}${cond.operator}"${cond.value}"`
    ).join(', ');
    
    return `=FILTER('${sourceSheet}'!${sourceRange}, ${filterConditions})`;
  },
  
  /**
   * Generate JOIN using QUERY
   */
  generateJoinFormula: function(params) {
    const { sheet1, sheet2, joinColumn, selectColumns, joinType } = params;
    
    // Google Sheets QUERY function for SQL-like operations
    const query = `SELECT ${selectColumns.join(', ')} WHERE ${joinColumn} IS NOT NULL`;
    
    return `=QUERY({'${sheet1}'!A:Z; '${sheet2}'!A:Z}, "${query}")`;
  },
  
  /**
   * Generate PIVOT formula using QUERY
   */
  generatePivotFormula: function(params) {
    const { sourceSheet, sourceRange, rowHeaders, columnHeaders, values, aggregation } = params;
    
    const aggFunction = aggregation.toUpperCase();
    const query = `SELECT ${rowHeaders} PIVOT ${columnHeaders}`;
    
    return `=QUERY('${sourceSheet}'!${sourceRange}, "${query}")`;
  },
  
  /**
   * Explain formula in plain language
   */
  explainFormula: function(formula, params) {
    const explanations = {
      'VLOOKUP': `This formula looks up values from ${params.sourceColumn} in ${params.targetSheet} and returns corresponding values.`,
      'INDEX': `This formula finds matching rows using INDEX and MATCH for better performance and flexibility.`,
      'SUMIFS': `This formula sums values from ${params.targetSheet} where specified conditions are met.`,
      'FILTER': `This formula filters data from ${params.sourceSheet} based on the specified criteria.`,
      'QUERY': `This formula performs SQL-like operations to join or transform data from multiple sheets.`
    };
    
    for (const [key, explanation] of Object.entries(explanations)) {
      if (formula.includes(key)) {
        return explanation;
      }
    }
    
    return 'This formula performs complex data operations across multiple sheets.';
  },
  
  /**
   * Validate formula for common errors
   */
  validateFormula: function(formula, params) {
    const issues = [];
    
    // Check for circular references
    if (params.sourceSheet === params.targetSheet && params.sourceColumn === params.targetColumn) {
      issues.push('Warning: Possible circular reference detected');
    }
    
    // Check for valid sheet names
    const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    const sheetNames = sheets.map(s => s.getName());
    
    [params.sourceSheet, params.targetSheet].forEach(sheetName => {
      if (sheetName && !sheetNames.includes(sheetName)) {
        issues.push(`Error: Sheet "${sheetName}" not found`);
      }
    });
    
    // Check for #REF! errors
    if (formula.includes('#REF!')) {
      issues.push('Error: Invalid cell reference');
    }
    
    // Check formula length (Google Sheets has a limit)
    if (formula.length > 50000) {
      issues.push('Warning: Formula may be too long for Google Sheets');
    }
    
    return {
      valid: issues.length === 0,
      issues: issues,
      suggestions: this.getSuggestions(formula, params)
    };
  },
  
  /**
   * Get optimization suggestions
   */
  getSuggestions: function(formula, params) {
    const suggestions = [];
    
    // Suggest INDEX/MATCH over VLOOKUP
    if (formula.includes('VLOOKUP')) {
      suggestions.push('Consider using INDEX/MATCH for better performance and flexibility');
    }
    
    // Suggest FILTER over complex QUERY for simple filtering
    if (formula.includes('QUERY') && !params.joinType) {
      suggestions.push('For simple filtering, FILTER function may be faster than QUERY');
    }
    
    // Suggest named ranges for frequently used ranges
    if ((formula.match(/'/g) || []).length > 4) {
      suggestions.push('Consider using Named Ranges for frequently referenced ranges');
    }
    
    // Suggest ARRAYFORMULA for row-wise operations
    if (!formula.includes('ARRAYFORMULA') && params.applyToAllRows) {
      suggestions.push('Use ARRAYFORMULA to apply this formula to all rows at once');
    }
    
    return suggestions;
  }
};