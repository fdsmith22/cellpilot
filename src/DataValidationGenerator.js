// Data Validation Generator Module
const DataValidationGenerator = {
  
  // Get current selection in the sheet
  getCurrentSelection: function() {
    try {
      const sheet = SpreadsheetApp.getActiveSheet();
      const range = sheet.getActiveRange();
      
      if (!range) {
        return null;
      }
      
      return {
        startCell: range.getA1Notation().split(':')[0],
        endCell: range.getA1Notation().split(':')[1] || range.getA1Notation().split(':')[0],
        sheet: sheet.getName(),
        numRows: range.getNumRows(),
        numCols: range.getNumColumns()
      };
    } catch (error) {
      console.error('Error getting current selection:', error);
      return null;
    }
  },
  
  // Apply data validation rule
  applyDataValidation: function(rule) {
    try {
      const sheet = SpreadsheetApp.getActiveSheet();
      const rangeNotation = `${rule.range.start}:${rule.range.end}`;
      const range = sheet.getRange(rangeNotation);
      
      let validation = null;
      
      switch(rule.type) {
        case 'list':
          if (!rule.items || rule.items.length === 0) {
            return {success: false, error: 'No items provided for list validation'};
          }
          validation = SpreadsheetApp.newDataValidation()
            .requireValueInList(rule.items, rule.showDropdown)
            .setAllowInvalid(!rule.rejectInvalid);
          break;
        
        case 'number':
          validation = SpreadsheetApp.newDataValidation();
          
          if (rule.min !== '' && rule.max !== '') {
            validation.requireNumberBetween(parseFloat(rule.min), parseFloat(rule.max));
          } else if (rule.min !== '') {
            validation.requireNumberGreaterThanOrEqualTo(parseFloat(rule.min));
          } else if (rule.max !== '') {
            validation.requireNumberLessThanOrEqualTo(parseFloat(rule.max));
          } else {
            validation.requireNumber();
          }
          
          validation.setAllowInvalid(!rule.rejectInvalid);
          break;
        
        case 'text':
          validation = SpreadsheetApp.newDataValidation();
          
          if (rule.pattern) {
            validation.requireTextContains(rule.pattern);
          } else if (rule.minLength !== '' || rule.maxLength !== '') {
            // Use custom formula for text length validation
            const cellRef = rule.range.start;
            let formula = '';
            
            if (rule.minLength !== '' && rule.maxLength !== '') {
              formula = `=AND(LEN(${cellRef})>=${rule.minLength}, LEN(${cellRef})<=${rule.maxLength})`;
            } else if (rule.minLength !== '') {
              formula = `=LEN(${cellRef})>=${rule.minLength}`;
            } else {
              formula = `=LEN(${cellRef})<=${rule.maxLength}`;
            }
            
            validation.requireFormulaSatisfied(formula);
          } else {
            validation.requireText();
          }
          
          validation.setAllowInvalid(!rule.rejectInvalid);
          break;
        
        case 'date':
          validation = SpreadsheetApp.newDataValidation();
          
          if (rule.startDate && rule.endDate) {
            validation.requireDateBetween(new Date(rule.startDate), new Date(rule.endDate));
          } else if (rule.startDate) {
            validation.requireDateOnOrAfter(new Date(rule.startDate));
          } else if (rule.endDate) {
            validation.requireDateOnOrBefore(new Date(rule.endDate));
          } else if (rule.allowPastDates && !rule.allowFutureDates) {
            validation.requireDateOnOrBefore(new Date());
          } else if (!rule.allowPastDates && rule.allowFutureDates) {
            validation.requireDateOnOrAfter(new Date());
          } else {
            validation.requireDate();
          }
          
          validation.setAllowInvalid(!rule.rejectInvalid);
          break;
        
        case 'custom':
          if (!rule.formula) {
            return {success: false, error: 'No formula provided for custom validation'};
          }
          
          validation = SpreadsheetApp.newDataValidation()
            .requireFormulaSatisfied(rule.formula)
            .setAllowInvalid(!rule.rejectInvalid);
          
          if (rule.errorMessage) {
            validation.setHelpText(rule.errorMessage);
          }
          break;
        
        case 'checkbox':
          validation = SpreadsheetApp.newDataValidation()
            .requireCheckbox();
          
          if (rule.defaultChecked) {
            // Set default values to TRUE
            range.setValue(true);
          }
          break;
        
        case 'email':
          // Use custom formula for email validation
          const emailCellRef = rule.range.start;
          let emailFormula = '';
          
          if (rule.allowMultiple) {
            // Complex formula for multiple emails
            emailFormula = `=REGEXMATCH(${emailCellRef}, "^([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\\.[a-zA-Z]{2,})(,\\s*[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\\.[a-zA-Z]{2,})*$")`;
          } else {
            // Simple email validation
            emailFormula = `=REGEXMATCH(${emailCellRef}, "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\\.[a-zA-Z]{2,}$")`;
          }
          
          validation = SpreadsheetApp.newDataValidation()
            .requireFormulaSatisfied(emailFormula)
            .setAllowInvalid(!rule.rejectInvalid)
            .setHelpText(rule.allowMultiple ? 'Enter email addresses separated by commas' : 'Enter a valid email address');
          break;
        
        case 'range':
          if (!rule.sourceRange) {
            return {success: false, error: 'No source range provided'};
          }
          
          try {
            const sourceRange = sheet.getRange(rule.sourceRange);
            validation = SpreadsheetApp.newDataValidation()
              .requireValueInRange(sourceRange, rule.showDropdown)
              .setAllowInvalid(!rule.rejectInvalid);
          } catch (e) {
            // Try with different sheet
            try {
              const sourceRange = SpreadsheetApp.getActiveSpreadsheet().getRange(rule.sourceRange);
              validation = SpreadsheetApp.newDataValidation()
                .requireValueInRange(sourceRange, rule.showDropdown)
                .setAllowInvalid(!rule.rejectInvalid);
            } catch (e2) {
              return {success: false, error: 'Invalid source range: ' + rule.sourceRange};
            }
          }
          break;
        
        default:
          return {success: false, error: 'Unknown validation type: ' + rule.type};
      }
      
      // Add help text if specified
      if (rule.showHelp && rule.helpText) {
        validation.setHelpText(rule.helpText);
      }
      
      // Build and apply the validation
      const builtValidation = validation.build();
      range.setDataValidation(builtValidation);
      
      return {
        success: true,
        message: `Data validation applied to ${rangeNotation}`,
        cellsAffected: range.getNumRows() * range.getNumColumns()
      };
      
    } catch (error) {
      console.error('Error applying data validation:', error);
      return {
        success: false,
        error: error.toString()
      };
    }
  },
  
  // Test a validation rule with a value
  testDataValidation: function(rule, testValue) {
    try {
      let isValid = false;
      let reason = '';
      
      switch(rule.type) {
        case 'list':
          isValid = rule.items.includes(testValue);
          if (!isValid) {
            reason = `Value must be one of: ${rule.items.join(', ')}`;
          }
          break;
        
        case 'number':
          const numValue = parseFloat(testValue);
          if (isNaN(numValue)) {
            isValid = false;
            reason = 'Value must be a number';
          } else {
            isValid = true;
            if (rule.min !== '' && numValue < parseFloat(rule.min)) {
              isValid = false;
              reason = `Value must be >= ${rule.min}`;
            }
            if (rule.max !== '' && numValue > parseFloat(rule.max)) {
              isValid = false;
              reason = `Value must be <= ${rule.max}`;
            }
            if (!rule.allowDecimals && numValue % 1 !== 0) {
              isValid = false;
              reason = 'Decimal values are not allowed';
            }
          }
          break;
        
        case 'text':
          const textLength = testValue.length;
          isValid = true;
          
          if (rule.minLength !== '' && textLength < parseInt(rule.minLength)) {
            isValid = false;
            reason = `Text must be at least ${rule.minLength} characters`;
          }
          if (rule.maxLength !== '' && textLength > parseInt(rule.maxLength)) {
            isValid = false;
            reason = `Text must be no more than ${rule.maxLength} characters`;
          }
          if (rule.pattern) {
            const regex = new RegExp(rule.pattern);
            if (!regex.test(testValue)) {
              isValid = false;
              reason = `Text must match pattern: ${rule.pattern}`;
            }
          }
          break;
        
        case 'date':
          const dateValue = new Date(testValue);
          if (isNaN(dateValue.getTime())) {
            isValid = false;
            reason = 'Invalid date format';
          } else {
            isValid = true;
            
            if (rule.startDate && dateValue < new Date(rule.startDate)) {
              isValid = false;
              reason = `Date must be on or after ${rule.startDate}`;
            }
            if (rule.endDate && dateValue > new Date(rule.endDate)) {
              isValid = false;
              reason = `Date must be on or before ${rule.endDate}`;
            }
          }
          break;
        
        case 'checkbox':
          isValid = testValue === 'true' || testValue === 'false' || 
                   testValue === true || testValue === false ||
                   testValue === 'TRUE' || testValue === 'FALSE';
          if (!isValid) {
            reason = 'Value must be TRUE or FALSE';
          }
          break;
        
        case 'email':
          const emailRegex = rule.allowMultiple
            ? /^([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})(,\s*[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})*$/
            : /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/;
          
          isValid = emailRegex.test(testValue);
          if (!isValid) {
            reason = rule.allowMultiple 
              ? 'Invalid email format (use commas for multiple)'
              : 'Invalid email format';
          }
          break;
        
        case 'custom':
          // For custom formulas, we can't easily test without applying to a cell
          isValid = true;
          reason = 'Custom formula validation requires cell context';
          break;
        
        case 'range':
          // For range validation, we'd need to check the actual range values
          isValid = true;
          reason = 'Range validation requires checking source range values';
          break;
      }
      
      return {
        isValid: isValid,
        reason: reason,
        value: testValue
      };
      
    } catch (error) {
      console.error('Error testing validation:', error);
      return {
        isValid: false,
        reason: 'Error testing validation: ' + error.toString()
      };
    }
  },
  
  // Clear data validation from a range
  clearDataValidation: function(startCell, endCell) {
    try {
      const sheet = SpreadsheetApp.getActiveSheet();
      const range = sheet.getRange(`${startCell}:${endCell}`);
      
      // Remove all data validation
      range.clearDataValidations();
      
      return {
        success: true,
        message: `Cleared validation from ${startCell}:${endCell}`
      };
      
    } catch (error) {
      console.error('Error clearing data validation:', error);
      return {
        success: false,
        error: error.toString()
      };
    }
  },
  
  // Get existing validation rules from a range
  getExistingValidation: function(rangeA1) {
    try {
      const sheet = SpreadsheetApp.getActiveSheet();
      const range = sheet.getRange(rangeA1);
      const validation = range.getDataValidation();
      
      if (!validation) {
        return null;
      }
      
      // Try to determine the validation type and settings
      const criteria = validation.getCriteriaType();
      const criteriaValues = validation.getCriteriaValues();
      
      return {
        type: criteria.toString(),
        values: criteriaValues,
        helpText: validation.getHelpText(),
        allowInvalid: validation.getAllowInvalid()
      };
      
    } catch (error) {
      console.error('Error getting existing validation:', error);
      return null;
    }
  },
  
  // Create validation from template
  createFromTemplate: function(templateName, range) {
    const templates = {
      'project_status': {
        type: 'list',
        items: ['Not Started', 'Planning', 'In Progress', 'Testing', 'Review', 'Complete', 'On Hold', 'Cancelled'],
        showDropdown: true,
        rejectInvalid: true,
        helpText: 'Select the current project status'
      },
      'priority': {
        type: 'list',
        items: ['Low', 'Medium', 'High', 'Critical', 'Urgent'],
        showDropdown: true,
        rejectInvalid: true,
        helpText: 'Select priority level'
      },
      'approval_status': {
        type: 'list',
        items: ['Pending', 'Approved', 'Rejected', 'Requires Changes', 'Under Review'],
        showDropdown: true,
        rejectInvalid: true,
        helpText: 'Select approval status'
      },
      'percentage': {
        type: 'number',
        min: '0',
        max: '100',
        allowDecimals: true,
        rejectInvalid: true,
        helpText: 'Enter a percentage between 0 and 100'
      },
      'rating': {
        type: 'number',
        min: '1',
        max: '5',
        allowDecimals: false,
        rejectInvalid: true,
        helpText: 'Enter a rating from 1 to 5'
      },
      'phone': {
        type: 'custom',
        formula: '=REGEXMATCH(A1, "^\\+?[1-9]\\d{1,14}$")',
        errorMessage: 'Enter a valid phone number',
        rejectInvalid: false
      },
      'url': {
        type: 'custom',
        formula: '=REGEXMATCH(A1, "^https?://[\\w\\-]+(\\.[\\w\\-]+)+[/#?]?.*$")',
        errorMessage: 'Enter a valid URL starting with http:// or https://',
        rejectInvalid: false
      }
    };
    
    const template = templates[templateName];
    if (!template) {
      return {success: false, error: 'Template not found: ' + templateName};
    }
    
    template.range = {start: range.start, end: range.end};
    return this.applyDataValidation(template);
  },
  
  // Analyze data to suggest validation rules
  suggestValidation: function(rangeA1) {
    try {
      const sheet = SpreadsheetApp.getActiveSheet();
      const range = sheet.getRange(rangeA1);
      const values = range.getValues().flat().filter(v => v !== '');
      
      if (values.length === 0) {
        return {type: 'none', reason: 'No data found in range'};
      }
      
      // Analyze the data
      const uniqueValues = [...new Set(values)];
      const isAllNumbers = values.every(v => !isNaN(v) && v !== '');
      const isAllDates = values.every(v => v instanceof Date || !isNaN(Date.parse(v)));
      const isAllEmails = values.every(v => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(v));
      const isAllBoolean = values.every(v => 
        v === true || v === false || 
        v === 'TRUE' || v === 'FALSE' || 
        v === 'Yes' || v === 'No'
      );
      
      // Suggest based on analysis
      if (isAllBoolean) {
        return {
          type: 'checkbox',
          reason: 'Data appears to be boolean values',
          confidence: 0.9
        };
      } else if (isAllEmails) {
        return {
          type: 'email',
          reason: 'Data appears to be email addresses',
          confidence: 0.95
        };
      } else if (isAllNumbers) {
        const min = Math.min(...values);
        const max = Math.max(...values);
        return {
          type: 'number',
          min: min,
          max: max,
          reason: `Data is numeric, ranging from ${min} to ${max}`,
          confidence: 0.85
        };
      } else if (isAllDates) {
        return {
          type: 'date',
          reason: 'Data appears to be dates',
          confidence: 0.8
        };
      } else if (uniqueValues.length <= 10 && values.length > uniqueValues.length * 2) {
        return {
          type: 'list',
          items: uniqueValues,
          reason: `Data has ${uniqueValues.length} unique values, suitable for a dropdown`,
          confidence: 0.75
        };
      } else {
        return {
          type: 'text',
          reason: 'Data appears to be free-form text',
          confidence: 0.5
        };
      }
      
    } catch (error) {
      console.error('Error suggesting validation:', error);
      return {
        type: 'error',
        reason: error.toString()
      };
    }
  }
};