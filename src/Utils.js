/**
 * CellPilot Utility Functions
 * Shared functions used across all modules
 */

const Utils = {
  
  /**
   * Safe execution wrapper with error handling
   */
  safeExecute: function(operation, fallback = null, context = 'Unknown operation') {
    try {
      return operation();
    } catch (error) {
      console.error(`Error in ${context}:`, error.message);
      this.logError(error, context);
      this.showUserError(error, context);
      return fallback;
    }
  },
  
  /**
   * Handle and display errors to users
   */
  handleError: function(error, userMessage = 'An error occurred') {
    console.error('CellPilot Error:', error);
    
    // Create error card for user
    return CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader()
        .setTitle('Error')
        .setImageUrl('https://fonts.gstatic.com/s/i/materialicons/error/v6/24px.svg'))
      .addSection(CardService.newCardSection()
        .addWidget(CardService.newTextParagraph()
          .setText(`<b>${userMessage}</b><br><br>Please try again or contact support if the problem persists.`)))
      .build();
  },
  
  /**
   * Show user-friendly error message
   */
  showUserError: function(error, context) {
    const userFriendlyMessages = {
      'insufficient_permissions': 'Please grant necessary permissions for this feature to work.',
      'quota_exceeded': 'You have reached your usage limit. Please upgrade your plan.',
      'invalid_data': 'The selected data format is not supported for this operation.',
      'network_error': 'Network error. Please check your connection and try again.'
    };
    
    const errorType = this.classifyError(error);
    const message = userFriendlyMessages[errorType] || `Error in ${context}: ${error.message}`;
    
    console.log('User error message:', message);
  },
  
  /**
   * Classify error types for better user messages
   */
  classifyError: function(error) {
    const message = error.message.toLowerCase();
    
    if (message.includes('permission') || message.includes('authorization')) {
      return 'insufficient_permissions';
    } else if (message.includes('quota') || message.includes('limit')) {
      return 'quota_exceeded';
    } else if (message.includes('invalid') || message.includes('format')) {
      return 'invalid_data';
    } else if (message.includes('network') || message.includes('timeout')) {
      return 'network_error';
    }
    
    return 'general_error';
  },
  
  /**
   * Log errors for debugging
   */
  logError: function(error, context, additionalData = {}) {
    const logEntry = {
      timestamp: new Date().toISOString(),
      context: context,
      error: error.message,
      stack: error.stack,
      data: additionalData
    };
    
    console.error('Error Log:', JSON.stringify(logEntry, null, 2));
  },
  
  /**
   * Detect data type in a range
   */
  detectDataType: function(data) {
    if (!data || data.length === 0) return 'empty';
    
    const firstRow = data[0];
    const sampleSize = Math.min(10, data.length);
    const sample = data.slice(0, sampleSize);
    
    let hasNumbers = false;
    let hasDates = false;
    let hasText = false;
    let hasEmails = false;
    let hasPhones = false;
    
    sample.forEach(row => {
      row.forEach(cell => {
        const cellStr = String(cell).trim();
        
        if (cellStr === '') return;
        
        if (this.isNumber(cellStr)) hasNumbers = true;
        if (this.isDate(cellStr)) hasDates = true;
        if (this.isEmail(cellStr)) hasEmails = true;
        if (this.isPhoneNumber(cellStr)) hasPhones = true;
        if (isNaN(cellStr) && !this.isDate(cellStr)) hasText = true;
      });
    });
    
    // Return most specific type
    if (hasEmails) return 'email';
    if (hasPhones) return 'phone';
    if (hasDates) return 'date';
    if (hasNumbers && hasText) return 'mixed';
    if (hasNumbers) return 'number';
    if (hasText) return 'text';
    
    return 'unknown';
  },
  
  /**
   * Check if value is a number
   */
  isNumber: function(value) {
    return !isNaN(value) && !isNaN(parseFloat(value));
  },
  
  /**
   * Check if value is a date
   */
  isDate: function(value) {
    const date = new Date(value);
    return date instanceof Date && !isNaN(date);
  },
  
  /**
   * Check if value is an email
   */
  isEmail: function(value) {
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    return emailRegex.test(value);
  },
  
  /**
   * Check if value is a phone number
   */
  isPhoneNumber: function(value) {
    const phoneRegex = /^[\+]?[\d\s\-\(\)]{10,}$/;
    return phoneRegex.test(String(value).replace(/\s/g, ''));
  },
  
  /**
   * Fuzzy string matching for duplicate detection
   */
  fuzzyMatch: function(str1, str2, threshold = 0.8) {
    if (str1 === str2) return 1;
    
    str1 = String(str1).toLowerCase().trim();
    str2 = String(str2).toLowerCase().trim();
    
    if (str1 === str2) return 1;
    
    const distance = this.levenshteinDistance(str1, str2);
    const maxLength = Math.max(str1.length, str2.length);
    
    if (maxLength === 0) return 1;
    
    const similarity = 1 - (distance / maxLength);
    return similarity >= threshold ? similarity : 0;
  },
  
  /**
   * Calculate Levenshtein distance between two strings
   */
  levenshteinDistance: function(str1, str2) {
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
    
    return matrix[str2.length][str1.length];
  },
  
  /**
   * Clean and standardize text
   */
  cleanText: function(text, options = {}) {
    let cleaned = String(text);
    
    if (options.trim !== false) {
      cleaned = cleaned.trim();
    }
    
    if (options.removeExtraSpaces !== false) {
      cleaned = cleaned.replace(/\s+/g, ' ');
    }
    
    if (options.case) {
      switch (options.case) {
        case 'upper':
          cleaned = cleaned.toUpperCase();
          break;
        case 'lower':
          cleaned = cleaned.toLowerCase();
          break;
        case 'title':
          cleaned = cleaned.replace(/\w\S*/g, (txt) => 
            txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase()
          );
          break;
        case 'sentence':
          cleaned = cleaned.charAt(0).toUpperCase() + cleaned.slice(1).toLowerCase();
          break;
      }
    }
    
    if (options.removeSpecialChars) {
      cleaned = cleaned.replace(/[^a-zA-Z0-9\s]/g, '');
    }
    
    return cleaned;
  },
  
  /**
   * Format numbers consistently
   */
  formatNumber: function(value, options = {}) {
    const num = parseFloat(value);
    if (isNaN(num)) return value;
    
    const formatter = new Intl.NumberFormat('en-US', {
      minimumFractionDigits: options.decimals || 0,
      maximumFractionDigits: options.decimals || 2,
      style: options.style || 'decimal',
      currency: options.currency || 'USD'
    });
    
    return formatter.format(num);
  },
  
  /**
   * Process data in batches to avoid timeout
   */
  processBatches: function(data, processor, batchSize = 100) {
    const results = [];
    
    for (let i = 0; i < data.length; i += batchSize) {
      const batch = data.slice(i, i + batchSize);
      const batchResults = processor(batch);
      results.push(...batchResults);
      
      // Allow other processes to run
      if (i % (batchSize * 10) === 0) {
        Utilities.sleep(10);
      }
    }
    
    return results;
  },
  
  /**
   * Create backup of range data
   */
  createBackup: function(range, label = 'Auto Backup') {
    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const backupSheet = spreadsheet.insertSheet(`Backup_${Date.now()}`);
      
      const data = range.getValues();
      const formulas = range.getFormulas();
      
      // Copy data
      if (data.length > 0) {
        backupSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
      }
      
      // Copy formulas where they exist
      formulas.forEach((row, i) => {
        row.forEach((formula, j) => {
          if (formula) {
            backupSheet.getRange(i + 1, j + 1).setFormula(formula);
          }
        });
      });
      
      // Add metadata
      backupSheet.getRange('A1').setNote(`Backup created: ${new Date().toISOString()}\nLabel: ${label}`);
      
      return backupSheet.getName();
    } catch (error) {
      console.error('Failed to create backup:', error);
      return null;
    }
  },
  
  /**
   * Show progress indicator
   */
  showProgress: function(current, total, operation = 'Processing') {
    const percentage = Math.round((current / total) * 100);
    console.log(`${operation}: ${percentage}% (${current}/${total})`);
  }
};