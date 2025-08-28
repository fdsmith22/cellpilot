/**
 * CellPilot Utility Functions with ML Adaptive Threshold Learning
 * Shared functions used across all modules with intelligent thresholds
 */

const Utils = {
  // ML Adaptive Threshold Properties
  mlEnabled: false,
  adaptiveThresholds: {
    fuzzyMatch: 0.8,
    duplicateDetection: 0.85,
    outlierDetection: 3.0,
    confidenceMinimum: 0.7,
    anomalyDetection: 0.75,
    patternRecognition: 0.8
  },
  thresholdHistory: [],
  userFeedback: [],
  performanceMetrics: {},
  
  /**
   * Initialize ML adaptive learning
   */
  initializeAdaptiveLearning: function() {
    try {
      const mlStatus = PropertiesService.getUserProperties().getProperty('cellpilot_ml_enabled');
      this.mlEnabled = mlStatus === 'true';
      
      if (this.mlEnabled) {
        // Load saved thresholds
        const savedThresholds = PropertiesService.getUserProperties().getProperty('adaptive_thresholds');
        if (savedThresholds) {
          this.adaptiveThresholds = JSON.parse(savedThresholds);
        }
        
        // Load threshold history
        const savedHistory = PropertiesService.getUserProperties().getProperty('threshold_history');
        if (savedHistory) {
          this.thresholdHistory = JSON.parse(savedHistory);
        }
        
        // Load performance metrics
        const savedMetrics = PropertiesService.getUserProperties().getProperty('threshold_metrics');
        if (savedMetrics) {
          this.performanceMetrics = JSON.parse(savedMetrics);
        }
      }
      
      return this.mlEnabled;
    } catch (error) {
      console.error('Error initializing adaptive learning:', error);
      return false;
    }
  },
  
  /**
   * Get adaptive threshold for specific operation
   */
  getAdaptiveThreshold: function(operation, defaultValue = null) {
    if (!this.mlEnabled) {
      return defaultValue || this.adaptiveThresholds[operation] || 0.8;
    }
    
    // Check if we have learned a better threshold
    const learnedThreshold = this.calculateOptimalThreshold(operation);
    if (learnedThreshold !== null) {
      return learnedThreshold;
    }
    
    return this.adaptiveThresholds[operation] || defaultValue || 0.8;
  },
  
  /**
   * Calculate optimal threshold based on feedback history
   */
  calculateOptimalThreshold: function(operation) {
    const relevantFeedback = this.userFeedback.filter(f => f.operation === operation);
    
    if (relevantFeedback.length < 5) {
      // Not enough data to calculate optimal threshold
      return null;
    }
    
    // Group feedback by threshold ranges
    const thresholdGroups = {};
    relevantFeedback.forEach(feedback => {
      const bucket = Math.floor(feedback.threshold * 10) / 10; // Round to nearest 0.1
      if (!thresholdGroups[bucket]) {
        thresholdGroups[bucket] = { correct: 0, incorrect: 0 };
      }
      
      if (feedback.wasCorrect) {
        thresholdGroups[bucket].correct++;
      } else {
        thresholdGroups[bucket].incorrect++;
      }
    });
    
    // Find threshold with best accuracy
    let bestThreshold = null;
    let bestAccuracy = 0;
    
    Object.entries(thresholdGroups).forEach(([threshold, counts]) => {
      const accuracy = counts.correct / (counts.correct + counts.incorrect);
      if (accuracy > bestAccuracy) {
        bestAccuracy = accuracy;
        bestThreshold = parseFloat(threshold);
      }
    });
    
    // Only use if significantly better than current
    if (bestAccuracy > 0.8 && bestThreshold !== null) {
      return bestThreshold;
    }
    
    return null;
  },
  
  /**
   * Record feedback for threshold learning
   */
  recordThresholdFeedback: function(operation, threshold, wasCorrect, context = {}) {
    if (!this.mlEnabled) return;
    
    const feedback = {
      timestamp: new Date().toISOString(),
      operation: operation,
      threshold: threshold,
      wasCorrect: wasCorrect,
      context: context
    };
    
    this.userFeedback.push(feedback);
    
    // Keep only recent feedback (last 1000 entries)
    if (this.userFeedback.length > 1000) {
      this.userFeedback = this.userFeedback.slice(-1000);
    }
    
    // Update threshold if we have enough positive feedback
    this.updateAdaptiveThreshold(operation);
    
    // Save feedback
    this.saveAdaptiveData();
  },
  
  /**
   * Update adaptive threshold based on feedback
   */
  updateAdaptiveThreshold: function(operation) {
    const relevantFeedback = this.userFeedback.filter(f => f.operation === operation);
    const recentFeedback = relevantFeedback.slice(-20); // Last 20 feedbacks
    
    if (recentFeedback.length < 10) return; // Need enough data
    
    const correctCount = recentFeedback.filter(f => f.wasCorrect).length;
    const accuracy = correctCount / recentFeedback.length;
    
    const currentThreshold = this.adaptiveThresholds[operation];
    let newThreshold = currentThreshold;
    
    if (accuracy < 0.7) {
      // Too many false positives, increase threshold
      newThreshold = Math.min(0.95, currentThreshold + 0.02);
    } else if (accuracy > 0.9 && recentFeedback.some(f => !f.wasCorrect)) {
      // Very high accuracy but some misses, might decrease threshold slightly
      newThreshold = Math.max(0.5, currentThreshold - 0.01);
    }
    
    // Record threshold change
    if (newThreshold !== currentThreshold) {
      this.thresholdHistory.push({
        timestamp: new Date().toISOString(),
        operation: operation,
        oldThreshold: currentThreshold,
        newThreshold: newThreshold,
        accuracy: accuracy,
        sampleSize: recentFeedback.length
      });
      
      this.adaptiveThresholds[operation] = newThreshold;
      
      // Keep history manageable
      if (this.thresholdHistory.length > 500) {
        this.thresholdHistory = this.thresholdHistory.slice(-500);
      }
    }
  },
  
  /**
   * Save adaptive learning data
   */
  saveAdaptiveData: function() {
    try {
      PropertiesService.getUserProperties().setProperty(
        'adaptive_thresholds',
        JSON.stringify(this.adaptiveThresholds)
      );
      
      PropertiesService.getUserProperties().setProperty(
        'threshold_history',
        JSON.stringify(this.thresholdHistory)
      );
      
      PropertiesService.getUserProperties().setProperty(
        'threshold_metrics',
        JSON.stringify(this.performanceMetrics)
      );
      
    } catch (error) {
      console.error('Error saving adaptive data:', error);
    }
  },
  
  /**
   * Track performance metrics for operations
   */
  trackPerformance: function(operation, duration, success, dataSize = 0) {
    if (!this.performanceMetrics[operation]) {
      this.performanceMetrics[operation] = {
        totalRuns: 0,
        successCount: 0,
        totalDuration: 0,
        avgDuration: 0,
        totalDataProcessed: 0,
        lastRun: null
      };
    }
    
    const metrics = this.performanceMetrics[operation];
    metrics.totalRuns++;
    if (success) metrics.successCount++;
    metrics.totalDuration += duration;
    metrics.avgDuration = metrics.totalDuration / metrics.totalRuns;
    metrics.totalDataProcessed += dataSize;
    metrics.lastRun = new Date().toISOString();
    metrics.successRate = (metrics.successCount / metrics.totalRuns) * 100;
    
    // Auto-adjust thresholds based on performance
    if (this.mlEnabled && metrics.totalRuns % 10 === 0) {
      this.autoAdjustThresholds(operation, metrics);
    }
  },
  
  /**
   * Auto-adjust thresholds based on performance metrics
   */
  autoAdjustThresholds: function(operation, metrics) {
    // If success rate is low, we might need to adjust thresholds
    if (metrics.successRate < 70 && this.adaptiveThresholds[operation]) {
      // Consider making threshold more lenient
      const adjustment = (70 - metrics.successRate) / 1000; // Small adjustment
      const currentThreshold = this.adaptiveThresholds[operation];
      const newThreshold = Math.max(0.5, currentThreshold - adjustment);
      
      if (Math.abs(newThreshold - currentThreshold) > 0.01) {
        this.adaptiveThresholds[operation] = newThreshold;
        
        console.log(`Auto-adjusted ${operation} threshold from ${currentThreshold} to ${newThreshold} based on ${metrics.successRate.toFixed(1)}% success rate`);
        
        this.saveAdaptiveData();
      }
    }
  },
  
  /**
   * Get threshold statistics
   */
  getThresholdStats: function() {
    const stats = {
      currentThresholds: this.adaptiveThresholds,
      totalFeedback: this.userFeedback.length,
      thresholdChanges: this.thresholdHistory.length,
      performanceMetrics: {},
      recentAdjustments: []
    };
    
    // Summarize performance metrics
    Object.entries(this.performanceMetrics).forEach(([operation, metrics]) => {
      stats.performanceMetrics[operation] = {
        runs: metrics.totalRuns,
        successRate: metrics.successRate?.toFixed(1) + '%',
        avgDuration: metrics.avgDuration?.toFixed(2) + 'ms'
      };
    });
    
    // Get recent threshold adjustments
    stats.recentAdjustments = this.thresholdHistory.slice(-5).map(h => ({
      operation: h.operation,
      change: `${h.oldThreshold.toFixed(2)} → ${h.newThreshold.toFixed(2)}`,
      accuracy: `${(h.accuracy * 100).toFixed(1)}%`,
      when: h.timestamp
    }));
    
    return stats;
  },
  
  /**
   * Reset adaptive thresholds to defaults
   */
  resetAdaptiveThresholds: function() {
    this.adaptiveThresholds = {
      fuzzyMatch: 0.8,
      duplicateDetection: 0.85,
      outlierDetection: 3.0,
      confidenceMinimum: 0.7,
      anomalyDetection: 0.75,
      patternRecognition: 0.8
    };
    
    this.thresholdHistory = [];
    this.userFeedback = [];
    this.performanceMetrics = {};
    
    this.saveAdaptiveData();
    
    return true;
  },
  
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
   * Detect data type in a range with smart header detection
   */
  detectDataType: function(data) {
    if (!data || data.length === 0) return 'empty';
    
    // Check if first row might be headers
    const hasHeaders = this.detectHeaders(data);
    const dataRows = hasHeaders && data.length > 1 ? data.slice(1) : data;
    
    // If we detected headers, return header info
    if (hasHeaders && data.length > 1) {
      const headerName = String(data[0][0]).trim();
      const dataType = this.analyzeColumnData(dataRows.map(row => row[0]));
      return `${headerName} (${dataType})`;
    }
    
    // Otherwise analyze all the data
    const sampleSize = Math.min(10, dataRows.length);
    const sample = dataRows.slice(0, sampleSize);
    
    // Count different data types
    let numbers = 0;
    let dates = 0;
    let text = 0;
    let emails = 0;
    let phones = 0;
    let currency = 0;
    let percentages = 0;
    let empty = 0;
    let total = 0;
    
    sample.forEach(row => {
      row.forEach(cell => {
        total++;
        const cellStr = String(cell).trim();
        
        if (cellStr === '') {
          empty++;
          return;
        }
        
        // Check for specific formats first
        if (this.isCurrency(cellStr)) {
          currency++;
        } else if (this.isPercentage(cellStr)) {
          percentages++;
        } else if (this.isEmail(cellStr)) {
          emails++;
        } else if (this.isPhoneNumber(cellStr)) {
          phones++;
        } else if (this.isDate(cellStr)) {
          dates++;
        } else if (this.isNumber(cellStr)) {
          numbers++;
        } else {
          text++;
        }
      });
    });
    
    // Calculate percentages
    const nonEmpty = total - empty;
    if (nonEmpty === 0) return 'empty';
    
    // Return most dominant type (at least 60% of non-empty cells)
    const threshold = nonEmpty * 0.6;
    
    if (currency > threshold) return 'currency';
    if (percentages > threshold) return 'percentage';
    if (emails > threshold) return 'email';
    if (phones > threshold) return 'phone';
    if (dates > threshold) return 'date';
    if (numbers > threshold) return 'number';
    if (text > threshold) return 'text';
    
    // If mixed, describe the mix
    if ((numbers + currency + percentages) > 0 && text > 0) return 'mixed';
    
    return 'unknown';
  },
  
  /**
   * Detect if first row contains headers
   */
  detectHeaders: function(data) {
    if (!data || data.length < 2) return false;
    
    const firstRow = data[0];
    const secondRow = data[1];
    
    // Common header patterns
    const headerPatterns = [
      /^(name|title|description|label|category|type|status|date|time|email|phone|address|city|state|country|zip|amount|price|cost|total|quantity|count|id|code|number|percent|rate|value|score|rating|notes|comments|salary|wage|income|revenue|expense|budget|tax|discount|profit|loss)$/i,
      /^(first|last|full|user|customer|client|product|item|order|invoice|payment|transaction|account|balance|credit|debit|employee|staff|department|location|branch|store)[\s_-]?(name|id|number|date|amount|type|status|code|salary|rate)?$/i,
      /^(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec|january|february|march|april|may|june|july|august|september|october|november|december|q[1-4]|quarter|year|month|week|day)[\s_-]?\d{0,4}$/i,
      /^(gross|net|base|total|annual|monthly|weekly|daily|hourly)[\s_-]?(salary|wage|pay|income|revenue|amount|value|cost|price)?$/i
    ];
    
    let headerLikeCount = 0;
    let firstRowIsText = 0;
    let secondRowIsNotText = 0;
    
    for (let i = 0; i < firstRow.length; i++) {
      const firstCell = String(firstRow[i]).trim();
      const secondCell = secondRow[i] ? String(secondRow[i]).trim() : '';
      
      // Check if first row cell matches header patterns
      if (headerPatterns.some(pattern => pattern.test(firstCell))) {
        headerLikeCount++;
      }
      
      // Check if first row is text and second row is data
      if (firstCell && isNaN(firstCell) && !this.isDate(firstCell)) {
        firstRowIsText++;
      }
      
      if (secondCell && (this.isNumber(secondCell) || this.isDate(secondCell) || this.isCurrency(secondCell))) {
        secondRowIsNotText++;
      }
    }
    
    // If most cells in first row look like headers
    if (headerLikeCount >= firstRow.length * 0.5) return true;
    
    // If first row is mostly text and second row is mostly data
    if (firstRowIsText >= firstRow.length * 0.7 && secondRowIsNotText >= secondRow.length * 0.5) return true;
    
    return false;
  },
  
  /**
   * Analyze a single column of data
   */
  analyzeColumnData: function(columnData) {
    if (!columnData || columnData.length === 0) return 'empty';
    
    const nonEmpty = columnData.filter(cell => String(cell).trim() !== '');
    if (nonEmpty.length === 0) return 'empty';
    
    let numbers = 0;
    let dates = 0;
    let currency = 0;
    let percentages = 0;
    let emails = 0;
    let phones = 0;
    let text = 0;
    
    nonEmpty.forEach(cell => {
      const cellStr = String(cell).trim();
      
      if (this.isCurrency(cellStr)) {
        currency++;
      } else if (this.isPercentage(cellStr)) {
        percentages++;
      } else if (this.isEmail(cellStr)) {
        emails++;
      } else if (this.isPhoneNumber(cellStr)) {
        phones++;
      } else if (this.isDate(cellStr)) {
        dates++;
      } else if (this.isNumber(cellStr)) {
        numbers++;
      } else {
        text++;
      }
    });
    
    const total = nonEmpty.length;
    const threshold = total * 0.6;
    
    if (currency > threshold) return 'currency';
    if (percentages > threshold) return 'percentage';
    if (emails > threshold) return 'email';
    if (phones > threshold) return 'phone';
    if (dates > threshold) return 'date';
    if (numbers > threshold) return 'number';
    if (text > threshold) return 'text';
    
    return 'mixed';
  },
  
  /**
   * Check if value is currency
   */
  isCurrency: function(value) {
    const currencyPattern = /^[$£€¥₹¢]\s?[\d,]+\.?\d*$|^[\d,]+\.?\d*\s?[$£€¥₹¢]$/;
    return currencyPattern.test(String(value).trim());
  },
  
  /**
   * Check if value is percentage
   */
  isPercentage: function(value) {
    const percentPattern = /^\d+\.?\d*\s?%$|^%\s?\d+\.?\d*$/;
    return percentPattern.test(String(value).trim());
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
   * Fuzzy string matching for duplicate detection with adaptive threshold
   */
  fuzzyMatch: function(str1, str2, threshold = null) {
    const startTime = Date.now();
    
    // Use adaptive threshold if ML is enabled
    if (threshold === null) {
      threshold = this.getAdaptiveThreshold('fuzzyMatch', 0.8);
    }
    
    if (str1 === str2) return 1;
    
    str1 = String(str1).toLowerCase().trim();
    str2 = String(str2).toLowerCase().trim();
    
    if (str1 === str2) return 1;
    
    const distance = this.levenshteinDistance(str1, str2);
    const maxLength = Math.max(str1.length, str2.length);
    
    if (maxLength === 0) return 1;
    
    const similarity = 1 - (distance / maxLength);
    const result = similarity >= threshold ? similarity : 0;
    
    // Track performance if ML is enabled
    if (this.mlEnabled) {
      const duration = Date.now() - startTime;
      this.trackPerformance('fuzzyMatch', duration, result > 0, str1.length + str2.length);
    }
    
    return result;
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
    
    // Remove currency symbols if requested
    if (options.removeCurrencySymbols) {
      // Common currency symbols and their names/codes
      cleaned = cleaned.replace(/[$£€¥₹¢₽₨₩₪₦₱₴₵₸₹]/g, '');
      cleaned = cleaned.replace(/\b(USD|GBP|EUR|JPY|INR|CAD|AUD)\b/gi, '');
      // Clean up any remaining spaces
      cleaned = cleaned.trim();
    }
    
    // Remove line breaks if requested
    if (options.removeLineBreaks) {
      cleaned = cleaned.replace(/[\r\n]+/g, ' ');
    }
    
    // Fix punctuation spacing if requested
    if (options.fixPunctuation) {
      // Add space after punctuation if missing
      cleaned = cleaned.replace(/([.,!?;:])([A-Za-z])/g, '$1 $2');
      // Remove space before punctuation
      cleaned = cleaned.replace(/\s+([.,!?;:])/g, '$1');
    }
    
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
        case 'proper':
          // Proper names - capitalize first letter of each word, keep rest as-is
          // Common name prefixes that should stay lowercase
          const lowercasePrefixes = ['de', 'van', 'von', 'la', 'le', 'di', 'da', 'del', 'della'];
          const words = cleaned.split(/\s+/);
          cleaned = words.map((word, index) => {
            // Keep lowercase prefixes lowercase (unless first word)
            if (index > 0 && lowercasePrefixes.includes(word.toLowerCase())) {
              return word.toLowerCase();
            }
            // Handle hyphenated names (Mary-Jane)
            if (word.includes('-')) {
              return word.split('-').map(part => 
                part.charAt(0).toUpperCase() + part.slice(1).toLowerCase()
              ).join('-');
            }
            // Handle apostrophes (O'Brien)
            if (word.includes("'")) {
              const parts = word.split("'");
              return parts.map(part => 
                part.charAt(0).toUpperCase() + part.slice(1).toLowerCase()
              ).join("'");
            }
            // Standard capitalization
            return word.charAt(0).toUpperCase() + word.slice(1).toLowerCase();
          }).join(' ');
          break;
      }
    }
    
    if (options.removeSpecialChars) {
      cleaned = cleaned.replace(/[^a-zA-Z0-9\s]/g, '');
    }
    
    return cleaned;
  },
  
  /**
   * Parse and standardize dates
   */
  parseDate: function(dateString, sourceFormat = 'auto') {
    if (!dateString) return null;
    
    const str = String(dateString).trim();
    
    // Try to parse as Date object first
    if (dateString instanceof Date) {
      return dateString;
    }
    
    // Common date formats to try
    const patterns = [
      // US formats (MM/DD/YYYY)
      { regex: /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/, format: 'US', parse: (m) => new Date(m[3], m[1] - 1, m[2]) },
      { regex: /^(\d{1,2})-(\d{1,2})-(\d{4})$/, format: 'US', parse: (m) => new Date(m[3], m[1] - 1, m[2]) },
      
      // European formats (DD/MM/YYYY)
      { regex: /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/, format: 'EU', parse: (m) => new Date(m[3], m[2] - 1, m[1]) },
      { regex: /^(\d{1,2})-(\d{1,2})-(\d{4})$/, format: 'EU', parse: (m) => new Date(m[3], m[2] - 1, m[1]) },
      
      // ISO format (YYYY-MM-DD)
      { regex: /^(\d{4})-(\d{1,2})-(\d{1,2})$/, format: 'ISO', parse: (m) => new Date(m[1], m[2] - 1, m[3]) },
      
      // Written dates
      { regex: /^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\w*\s+(\d{1,2}),?\s+(\d{4})$/i, format: 'written', 
        parse: (m) => new Date(Date.parse(m[0])) },
      { regex: /^(\d{1,2})\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\w*\s+(\d{4})$/i, format: 'written',
        parse: (m) => new Date(Date.parse(m[0])) }
    ];
    
    // If source format is specified, prioritize it
    if (sourceFormat === 'US') {
      // Try US format first (MM/DD/YYYY)
      const match = str.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
      if (match) {
        const year = match[3].length === 2 ? '20' + match[3] : match[3];
        return new Date(year, match[1] - 1, match[2]);
      }
    } else if (sourceFormat === 'EU') {
      // Try European format first (DD/MM/YYYY)
      const match = str.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
      if (match) {
        const year = match[3].length === 2 ? '20' + match[3] : match[3];
        return new Date(year, match[2] - 1, match[1]);
      }
    }
    
    // Try all patterns
    for (const pattern of patterns) {
      const match = str.match(pattern.regex);
      if (match) {
        try {
          const date = pattern.parse(match);
          // Validate the date is reasonable (between 1900 and 2100)
          if (date.getFullYear() >= 1900 && date.getFullYear() <= 2100) {
            return date;
          }
        } catch (e) {
          continue;
        }
      }
    }
    
    // Try native Date.parse as last resort
    const parsed = Date.parse(str);
    if (!isNaN(parsed)) {
      const date = new Date(parsed);
      if (date.getFullYear() >= 1900 && date.getFullYear() <= 2100) {
        return date;
      }
    }
    
    return null;
  },
  
  /**
   * Format date to specified format
   */
  formatDate: function(date, targetFormat = 'ISO') {
    if (!date) return '';
    
    // Ensure we have a Date object
    const dateObj = date instanceof Date ? date : new Date(date);
    if (isNaN(dateObj.getTime())) return '';
    
    const pad = (n) => String(n).padStart(2, '0');
    const day = pad(dateObj.getDate());
    const month = pad(dateObj.getMonth() + 1);
    const year = dateObj.getFullYear();
    
    const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 
                       'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    const monthName = monthNames[dateObj.getMonth()];
    
    switch (targetFormat) {
      case 'US':
        return `${month}/${day}/${year}`;
      case 'EU':
        return `${day}/${month}/${year}`;
      case 'ISO':
        return `${year}-${month}-${day}`;
      case 'written':
        return `${monthName} ${dateObj.getDate()}, ${year}`;
      case 'written-UK':
        return `${dateObj.getDate()} ${monthName} ${year}`;
      default:
        return `${year}-${month}-${day}`;
    }
  },
  
  /**
   * Format numbers consistently
   */
  formatNumber: function(value, options = {}) {
    // First, clean and extract the numeric value
    let cleanedValue = String(value).trim();
    
    // Remove currency symbols
    cleanedValue = cleanedValue.replace(/[$£€¥₹¢₽₨₩₪₦₱₴₵₸₹]/g, '');
    
    // Remove any other non-numeric characters except commas, periods, and minus
    cleanedValue = cleanedValue.replace(/[^\d,.-]/g, '');
    
    // Handle different decimal/thousand separator formats
    // Check if it's European format (. for thousands, , for decimal)
    const europeanPattern = /^\d{1,3}(\.\d{3})*,\d+$/;
    if (europeanPattern.test(cleanedValue)) {
      // Convert European to US format
      cleanedValue = cleanedValue.replace(/\./g, '').replace(',', '.');
    } else {
      // US format or no formatting - just remove commas
      cleanedValue = cleanedValue.replace(/,/g, '');
    }
    
    const num = parseFloat(cleanedValue);
    if (isNaN(num)) return value;
    
    // Format based on options
    if (options.format === 'plain') {
      // No formatting, just the number
      if (options.decimals !== undefined) {
        return num.toFixed(options.decimals);
      }
      return num.toString();
    } else if (options.format === 'currency') {
      // Format as currency
      return new Intl.NumberFormat('en-US', {
        style: 'currency',
        currency: options.currency || 'USD',
        minimumFractionDigits: options.decimals !== undefined ? options.decimals : 2,
        maximumFractionDigits: options.decimals !== undefined ? options.decimals : 2
      }).format(num);
    } else {
      // Default: format with thousand separators
      return new Intl.NumberFormat('en-US', {
        minimumFractionDigits: options.decimals !== undefined ? options.decimals : 0,
        maximumFractionDigits: options.decimals !== undefined ? options.decimals : 2,
        useGrouping: options.useCommas !== false
      }).format(num);
    }
  },
  
  /**
   * Clean and standardize numeric values
   */
  standardizeNumber: function(value, options = {}) {
    // Extract numeric value
    let cleanedValue = String(value).trim();
    
    // Remove currency symbols
    if (options.removeCurrencySymbols !== false) {
      cleanedValue = cleanedValue.replace(/[$£€¥₹¢₽₨₩₪₦₱₴₵₸₹]/g, '');
    }
    
    // Remove any other non-numeric characters except commas, periods, and minus
    cleanedValue = cleanedValue.replace(/[^\d,.-]/g, '').trim();
    
    // Handle different decimal/thousand separator formats
    const europeanPattern = /^\d{1,3}(\.\d{3})*,\d+$/;
    if (europeanPattern.test(cleanedValue)) {
      // Convert European to US format
      cleanedValue = cleanedValue.replace(/\./g, '').replace(',', '.');
    } else {
      // US format - just remove commas
      cleanedValue = cleanedValue.replace(/,/g, '');
    }
    
    const num = parseFloat(cleanedValue);
    if (isNaN(num)) return value;
    
    // Apply formatting based on options
    if (options.keepOriginalFormat) {
      // Try to maintain the original format (with/without decimals, with/without commas)
      const hasDecimals = String(value).includes('.');
      const hasCommas = String(value).includes(',');
      
      if (hasCommas && hasDecimals) {
        return new Intl.NumberFormat('en-US', {
          minimumFractionDigits: 2,
          maximumFractionDigits: 2
        }).format(num);
      } else if (hasCommas && !hasDecimals) {
        return new Intl.NumberFormat('en-US', {
          minimumFractionDigits: 0,
          maximumFractionDigits: 0
        }).format(num);
      } else if (!hasCommas && hasDecimals) {
        return num.toFixed(2);
      } else {
        return num.toString();
      }
    } else {
      // Apply consistent formatting
      const decimals = options.decimals !== undefined ? options.decimals : 2;
      const useCommas = options.useCommas !== false;
      
      if (useCommas) {
        return new Intl.NumberFormat('en-US', {
          minimumFractionDigits: decimals,
          maximumFractionDigits: decimals
        }).format(num);
      } else {
        return decimals > 0 ? num.toFixed(decimals) : Math.round(num).toString();
      }
    }
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
  },
  
  /**
   * Clean messy data with advanced handling
   */
  cleanMessyData: function(value, options = {}) {
    // Handle null, undefined, and empty values
    if (value === null || value === undefined) {
      return options.defaultValue || '';
    }
    
    let cleaned = String(value).trim();
    
    // Handle common data issues
    if (options.handleMissingData) {
      // Replace common missing data indicators
      const missingIndicators = [
        'N/A', 'n/a', 'NA', 'na', 'NULL', 'null', 'None', 'none',
        '#N/A', '#NA', '#NULL!', '#VALUE!', '#REF!', '#DIV/0!',
        '-', '--', '---', '...', '?', '??', 'unknown', 'Unknown',
        'not available', 'Not Available', 'TBD', 'tbd', 'pending'
      ];
      
      if (missingIndicators.includes(cleaned)) {
        return options.missingDataReplacement || '';
      }
    }
    
    // Remove invisible characters and normalize whitespace
    if (options.removeInvisibleChars !== false) {
      // Remove zero-width spaces, non-breaking spaces, etc.
      cleaned = cleaned.replace(/[\u200B\u200C\u200D\uFEFF]/g, '');
      // Replace non-breaking spaces with regular spaces
      cleaned = cleaned.replace(/\u00A0/g, ' ');
      // Remove other control characters
      cleaned = cleaned.replace(/[\x00-\x1F\x7F]/g, '');
    }
    
    // Handle encoding issues
    if (options.fixEncoding) {
      // Fix common encoding problems
      cleaned = cleaned
        .replace(/â€™/g, "'")
        .replace(/â€œ/g, '"')
        .replace(/â€/g, '"')
        .replace(/â€"/g, '–')
        .replace(/â€"/g, '—')
        .replace(/Â°/g, '°')
        .replace(/Â£/g, '£')
        .replace(/â‚¬/g, '€')
        .replace(/Ã©/g, 'é')
        .replace(/Ã¨/g, 'è')
        .replace(/Ã¢/g, 'â')
        .replace(/Ã´/g, 'ô')
        .replace(/Ã§/g, 'ç');
    }
    
    // Handle mixed data types
    if (options.normalizeDataType) {
      const dataType = this.detectValueType(cleaned);
      
      switch (dataType) {
        case 'number':
          // Extract number from text like "$1,234.56" or "123 items"
          const numMatch = cleaned.match(/[+-]?\d+(?:,\d{3})*(?:\.\d+)?/);
          if (numMatch) {
            cleaned = numMatch[0].replace(/,/g, '');
          }
          break;
          
        case 'percentage':
          // Extract percentage value
          const pctMatch = cleaned.match(/([+-]?\d+(?:\.\d+)?)\s*%/);
          if (pctMatch) {
            cleaned = (parseFloat(pctMatch[1]) / 100).toString();
          }
          break;
          
        case 'boolean':
          // Normalize boolean values
          const boolValue = this.parseBoolean(cleaned);
          cleaned = boolValue !== null ? boolValue.toString() : cleaned;
          break;
      }
    }
    
    // Handle HTML entities
    if (options.decodeHtml) {
      cleaned = cleaned
        .replace(/&amp;/g, '&')
        .replace(/&lt;/g, '<')
        .replace(/&gt;/g, '>')
        .replace(/&quot;/g, '"')
        .replace(/&#39;/g, "'")
        .replace(/&nbsp;/g, ' ');
    }
    
    // Fix common OCR errors
    if (options.fixOcrErrors) {
      cleaned = cleaned
        .replace(/[0O]([A-Z])/g, 'O$1')  // 0 instead of O
        .replace(/([A-Z])[0O]/g, '$1O')
        .replace(/[l1I]([A-Z])/g, 'I$1')  // 1 or l instead of I
        .replace(/([a-z])[1I]/g, '$1l')   // 1 or I instead of l
        .replace(/rn/g, 'm')              // rn instead of m
        .replace(/vv/g, 'w');             // vv instead of w
    }
    
    return cleaned;
  },
  
  /**
   * Detect the type of a value
   */
  detectValueType: function(value) {
    const str = String(value).trim();
    
    // Check for percentage
    if (/^\s*[+-]?\d+(?:\.\d+)?\s*%\s*$/.test(str)) {
      return 'percentage';
    }
    
    // Check for currency
    if (/^[\$£€¥₹]\s*[+-]?\d+(?:,\d{3})*(?:\.\d+)?$/.test(str)) {
      return 'currency';
    }
    
    // Check for number
    if (/^[+-]?\d+(?:,\d{3})*(?:\.\d+)?$/.test(str)) {
      return 'number';
    }
    
    // Check for date
    if (this.parseDate(str, 'auto')) {
      return 'date';
    }
    
    // Check for boolean
    const booleanValues = ['true', 'false', 'yes', 'no', 'y', 'n', '1', '0'];
    if (booleanValues.includes(str.toLowerCase())) {
      return 'boolean';
    }
    
    // Check for email
    if (/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(str)) {
      return 'email';
    }
    
    // Check for URL
    if (/^https?:\/\/[^\s]+/.test(str)) {
      return 'url';
    }
    
    return 'text';
  },
  
  /**
   * Parse boolean values from various formats
   */
  parseBoolean: function(value) {
    const str = String(value).trim().toLowerCase();
    
    const trueValues = ['true', 'yes', 'y', '1', 'on', 'enabled', 'active'];
    const falseValues = ['false', 'no', 'n', '0', 'off', 'disabled', 'inactive'];
    
    if (trueValues.includes(str)) return true;
    if (falseValues.includes(str)) return false;
    
    return null;
  },
  
  /**
   * Validate and fix phone numbers
   */
  cleanPhoneNumber: function(phone, defaultCountryCode = '+1') {
    let cleaned = String(phone).replace(/[^\d+]/g, '');
    
    // Remove leading zeros
    cleaned = cleaned.replace(/^0+/, '');
    
    // Add country code if missing
    if (!cleaned.startsWith('+')) {
      if (cleaned.length === 10) {
        // Assume US/Canada number
        cleaned = defaultCountryCode + cleaned;
      }
    }
    
    // Format consistently
    if (cleaned.length >= 10) {
      const match = cleaned.match(/^(\+?\d{1,3})(\d{3})(\d{3})(\d{4})$/);
      if (match) {
        return `${match[1]} (${match[2]}) ${match[3]}-${match[4]}`;
      }
    }
    
    return cleaned;
  },
  
  /**
   * Extract and clean email addresses
   */
  cleanEmail: function(email) {
    const cleaned = String(email).trim().toLowerCase();
    
    // Extract email from text like "John Doe <john@example.com>"
    const match = cleaned.match(/<?([^\s@<>]+@[^\s@<>]+\.[^\s@<>]+)>?/);
    
    if (match) {
      return match[1];
    }
    
    return cleaned;
  },
  
  /**
   * Handle large datasets efficiently
   */
  processLargeDataset: function(range, processor, options = {}) {
    const batchSize = options.batchSize || 500;
    const totalRows = range.getNumRows();
    const totalCols = range.getNumColumns();
    const results = [];
    
    for (let startRow = 1; startRow <= totalRows; startRow += batchSize) {
      const numRows = Math.min(batchSize, totalRows - startRow + 1);
      const batch = range.offset(startRow - 1, 0, numRows, totalCols);
      const batchData = batch.getValues();
      
      // Process batch
      const processedBatch = processor(batchData, {
        startRow: startRow,
        batchIndex: Math.floor(startRow / batchSize)
      });
      
      results.push(...processedBatch);
      
      // Show progress
      this.showProgress(startRow + numRows - 1, totalRows, 'Processing rows');
      
      // Prevent timeout
      if (startRow % (batchSize * 5) === 1) {
        Utilities.sleep(100);
      }
    }
    
    return results;
  },
  
  /**
   * Smart data type detection for columns
   */
  detectColumnTypes: function(data, sampleSize = 100) {
    if (!data || data.length === 0) return [];
    
    const numCols = data[0].length;
    const columnTypes = [];
    const sample = data.slice(0, Math.min(sampleSize, data.length));
    
    for (let col = 0; col < numCols; col++) {
      const columnValues = sample.map(row => row[col]).filter(val => val !== '' && val !== null);
      const typeCounts = {};
      
      columnValues.forEach(value => {
        const type = this.detectValueType(value);
        typeCounts[type] = (typeCounts[type] || 0) + 1;
      });
      
      // Find dominant type (>60% of non-empty values)
      let dominantType = 'text';
      const threshold = columnValues.length * 0.6;
      
      for (const [type, count] of Object.entries(typeCounts)) {
        if (count >= threshold) {
          dominantType = type;
          break;
        }
      }
      
      columnTypes.push({
        index: col,
        type: dominantType,
        distribution: typeCounts,
        sampleSize: columnValues.length
      });
    }
    
    return columnTypes;
  }
};