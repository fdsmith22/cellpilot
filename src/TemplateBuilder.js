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