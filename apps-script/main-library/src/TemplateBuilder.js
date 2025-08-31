/**
 * TemplateBuilder - Core template building utilities
 * Provides basic infrastructure for creating templates
 */

var TemplateBuilder = {
  /**
   * Build a template with comprehensive error handling
   */
  buildTemplate: function(spreadsheet, config) {
    const builder = {
      sheets: {},
      errors: [],
      warnings: [],
      
      /**
       * Create and store a sheet reference
       */
      createSheet: function(name, isPreview = false) {
        try {
          const sheetName = isPreview ? `[PREVIEW] ${name}` : name;
          const sheet = spreadsheet.insertSheet(sheetName);
          this.sheets[name] = {
            sheet: sheet,
            name: sheetName,
            lastRow: 1,
            lastCol: 1
          };
          return sheet;
        } catch (error) {
          this.errors.push(`Failed to create sheet ${name}: ${error.toString()}`);
          return null;
        }
      },
      
      /**
       * Safely set cell value with validation
       */
      setCellValue: function(sheetKey, row, col, value) {
        try {
          const sheetInfo = this.sheets[sheetKey];
          if (!sheetInfo) {
            this.errors.push(`Sheet ${sheetKey} not found`);
            return false;
          }
          
          sheetInfo.sheet.getRange(row, col).setValue(value);
          sheetInfo.lastRow = Math.max(sheetInfo.lastRow, row);
          sheetInfo.lastCol = Math.max(sheetInfo.lastCol, col);
          return true;
        } catch (error) {
          this.errors.push(`Failed to set cell value at ${sheetKey}[${row},${col}]: ${error.toString()}`);
          return false;
        }
      },
      
      /**
       * Safely set a range of values
       */
      setRangeValues: function(sheetKey, startRow, startCol, values) {
        try {
          const sheetInfo = this.sheets[sheetKey];
          if (!sheetInfo) {
            this.errors.push(`Sheet ${sheetKey} not found`);
            return false;
          }
          
          if (!values || values.length === 0) return false;
          
          const numRows = values.length;
          const numCols = values[0].length;
          
          sheetInfo.sheet.getRange(startRow, startCol, numRows, numCols).setValues(values);
          sheetInfo.lastRow = Math.max(sheetInfo.lastRow, startRow + numRows - 1);
          sheetInfo.lastCol = Math.max(sheetInfo.lastCol, startCol + numCols - 1);
          return true;
        } catch (error) {
          this.errors.push(`Failed to set range values at ${sheetKey}[${startRow},${startCol}]: ${error.toString()}`);
          return false;
        }
      },
      
      /**
       * Safely set a formula with sheet name resolution
       */
      setFormula: function(sheetKey, row, col, formula) {
        try {
          const sheetInfo = this.sheets[sheetKey];
          if (!sheetInfo) {
            this.errors.push(`Sheet ${sheetKey} not found`);
            return false;
          }
          
          // Replace sheet references with actual names
          let resolvedFormula = formula;
          for (const [key, info] of Object.entries(this.sheets)) {
            const placeholder = `{{${key}}}`;
            resolvedFormula = resolvedFormula.replace(new RegExp(placeholder, 'g'), `'${info.name}'`);
          }
          
          sheetInfo.sheet.getRange(row, col).setFormula(resolvedFormula);
          sheetInfo.lastRow = Math.max(sheetInfo.lastRow, row);
          sheetInfo.lastCol = Math.max(sheetInfo.lastCol, col);
          return true;
        } catch (error) {
          this.errors.push(`Failed to set formula at ${sheetKey}[${row},${col}]: ${error.toString()}`);
          return false;
        }
      },
      
      /**
       * Apply formatting to a range
       */
      formatRange: function(sheetKey, startRow, startCol, endRow, endCol, formatting) {
        try {
          const sheetInfo = this.sheets[sheetKey];
          if (!sheetInfo) {
            this.errors.push(`Sheet ${sheetKey} not found`);
            return false;
          }
          
          const range = sheetInfo.sheet.getRange(startRow, startCol, endRow - startRow + 1, endCol - startCol + 1);
          
          if (formatting.background) range.setBackground(formatting.background);
          if (formatting.fontColor) range.setFontColor(formatting.fontColor);
          if (formatting.fontSize) range.setFontSize(formatting.fontSize);
          if (formatting.fontWeight) range.setFontWeight(formatting.fontWeight);
          if (formatting.horizontalAlignment) range.setHorizontalAlignment(formatting.horizontalAlignment);
          if (formatting.numberFormat) range.setNumberFormat(formatting.numberFormat);
          if (formatting.border) {
            range.setBorder(true, true, true, true, true, true, '#D1D5DB', SpreadsheetApp.BorderStyle.SOLID);
          }
          
          return true;
        } catch (error) {
          this.errors.push(`Failed to format range at ${sheetKey}: ${error.toString()}`);
          return false;
        }
      },
      
      /**
       * Add data validation to a range
       */
      addValidation: function(sheetKey, rangeA1, values) {
        try {
          const sheetInfo = this.sheets[sheetKey];
          if (!sheetInfo) {
            this.errors.push(`Sheet ${sheetKey} not found`);
            return false;
          }
          
          const rule = SpreadsheetApp.newDataValidation()
            .requireValueInList(values, true)
            .build();
          
          sheetInfo.sheet.getRange(rangeA1).setDataValidation(rule);
          return true;
        } catch (error) {
          this.errors.push(`Failed to add validation at ${sheetKey}[${rangeA1}]: ${error.toString()}`);
          return false;
        }
      },
      
      /**
       * Get build results
       */
      getResults: function() {
        return {
          success: this.errors.length === 0,
          sheets: Object.keys(this.sheets),
          errors: this.errors,
          warnings: this.warnings
        };
      }
    };
    
    // Execute the build configuration
    config(builder);
    
    return builder.getResults();
  }
};

/**
 * Advanced Template Builder Pro
 * Shared utility class for creating complex template sheets
 */
class TemplateBuilderPro {
  constructor(spreadsheet, isPreview) {
    this.spreadsheet = spreadsheet;
    this.isPreview = isPreview;
    this.sheets = {};
    this.errors = [];
  }
  
  createSheets(names) {
    names.forEach(name => {
      const sheetName = this.isPreview ? `[PREVIEW] ${name}` : name;
      try {
        const sheet = this.spreadsheet.insertSheet(sheetName);
        this.sheets[name] = new SheetHelper(sheet, name, this);
      } catch (error) {
        this.errors.push(`Failed to create sheet ${name}: ${error.toString()}`);
      }
    });
    return this.sheets;
  }
  
  sheet(name) {
    return this.sheets[name];
  }
  
  complete() {
    return {
      success: this.errors.length === 0,
      sheets: Object.keys(this.sheets).map(name => 
        this.isPreview ? `[PREVIEW] ${name}` : name
      ),
      errors: this.errors
    };
  }
  
  resolveSheetReference(formula) {
    let resolved = formula;
    for (const [name, helper] of Object.entries(this.sheets)) {
      const placeholder = `{{${name}}}`;
      const actualName = this.isPreview ? `'[PREVIEW] ${name}'` : `'${name}'`;
      resolved = resolved.replace(new RegExp(placeholder, 'g'), actualName);
    }
    return resolved;
  }
}

/**
 * Sheet Helper Class
 * Provides utility methods for working with individual sheets
 */
class SheetHelper {
  constructor(sheet, name, builder) {
    this.sheet = sheet;
    this.name = name;
    this.builder = builder;
  }
  
  setValue(row, col, value) {
    try {
      this.sheet.getRange(row, col).setValue(value);
      return true;
    } catch (error) {
      this.builder.errors.push(`Error setting value at ${this.name}[${row},${col}]: ${error.toString()}`);
      return false;
    }
  }
  
  setFormula(row, col, formula) {
    try {
      const resolved = this.builder.resolveSheetReference(formula);
      this.sheet.getRange(row, col).setFormula(resolved);
      return true;
    } catch (error) {
      this.builder.errors.push(`Error setting formula at ${this.name}[${row},${col}]: ${error.toString()}`);
      return false;
    }
  }
  
  setRangeValues(startRow, startCol, values) {
    try {
      if (!values || values.length === 0) return false;
      const numRows = values.length;
      const numCols = values[0].length;
      this.sheet.getRange(startRow, startCol, numRows, numCols).setValues(values);
      return true;
    } catch (error) {
      this.builder.errors.push(`Error setting range values at ${this.name}[${startRow},${startCol}]: ${error.toString()}`);
      return false;
    }
  }
  
  format(startRow, startCol, endRow, endCol, formatting) {
    try {
      const range = this.sheet.getRange(startRow, startCol, endRow - startRow + 1, endCol - startCol + 1);
      
      if (formatting.background) range.setBackground(formatting.background);
      if (formatting.fontColor) range.setFontColor(formatting.fontColor);
      if (formatting.fontSize) range.setFontSize(formatting.fontSize);
      if (formatting.fontWeight) range.setFontWeight(formatting.fontWeight);
      if (formatting.horizontalAlignment) range.setHorizontalAlignment(formatting.horizontalAlignment);
      if (formatting.verticalAlignment) range.setVerticalAlignment(formatting.verticalAlignment);
      if (formatting.numberFormat) range.setNumberFormat(formatting.numberFormat);
      if (formatting.border) {
        range.setBorder(true, true, true, true, true, true, '#D1D5DB', SpreadsheetApp.BorderStyle.SOLID);
      }
      
      return true;
    } catch (error) {
      this.builder.errors.push(`Error formatting range at ${this.name}: ${error.toString()}`);
      return false;
    }
  }
  
  merge(startRow, startCol, endRow, endCol) {
    try {
      const range = this.sheet.getRange(startRow, startCol, endRow - startRow + 1, endCol - startCol + 1);
      range.merge();
      return true;
    } catch (error) {
      this.builder.errors.push(`Error merging cells at ${this.name}: ${error.toString()}`);
      return false;
    }
  }
  
  addValidation(rangeA1, values) {
    try {
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(values, true)
        .build();
      this.sheet.getRange(rangeA1).setDataValidation(rule);
      return true;
    } catch (error) {
      this.builder.errors.push(`Error adding validation at ${this.name}[${rangeA1}]: ${error.toString()}`);
      return false;
    }
  }
  
  addConditionalFormat(startRow, startCol, endRow, endCol, rule) {
    try {
      const range = this.sheet.getRange(startRow, startCol, endRow - startRow + 1, endCol - startCol + 1);
      const rules = this.sheet.getConditionalFormatRules();
      
      let newRule;
      const builder = SpreadsheetApp.newConditionalFormatRule()
        .setRanges([range]);
      
      if (rule.type === 'cell') {
        switch(rule.condition) {
          case 'NUMBER_GREATER_THAN':
            builder.whenNumberGreaterThan(rule.value);
            break;
          case 'NUMBER_LESS_THAN':
            builder.whenNumberLessThan(rule.value);
            break;
          case 'TEXT_CONTAINS':
            builder.whenTextContains(rule.value);
            break;
          case 'TEXT_EQUALS':
            builder.whenTextEqualTo(rule.value);
            break;
        }
        
        if (rule.format.fontColor) {
          builder.setFontColor(rule.format.fontColor);
        }
        if (rule.format.background) {
          builder.setBackground(rule.format.background);
        }
      } else if (rule.type === 'gradient') {
        builder.setGradientMaxpointWithValue(
          rule.maxColor,
          SpreadsheetApp.InterpolationType.NUMBER,
          rule.maxValue || '100'
        );
        builder.setGradientMidpointWithValue(
          rule.midColor,
          SpreadsheetApp.InterpolationType.NUMBER,
          rule.midValue || '50'
        );
        builder.setGradientMinpointWithValue(
          rule.minColor,
          SpreadsheetApp.InterpolationType.NUMBER,
          rule.minValue || '0'
        );
      }
      
      newRule = builder.build();
      rules.push(newRule);
      this.sheet.setConditionalFormatRules(rules);
      
      return true;
    } catch (error) {
      this.builder.errors.push(`Error adding conditional format at ${this.name}: ${error.toString()}`);
      return false;
    }
  }
  
  setColumnWidth(column, width) {
    try {
      this.sheet.setColumnWidth(column, width);
      return true;
    } catch (error) {
      this.builder.errors.push(`Error setting column width at ${this.name}[${column}]: ${error.toString()}`);
      return false;
    }
  }
  
  setRowHeight(row, height) {
    try {
      this.sheet.setRowHeight(row, height);
      return true;
    } catch (error) {
      this.builder.errors.push(`Error setting row height at ${this.name}[${row}]: ${error.toString()}`);
      return false;
    }
  }
}