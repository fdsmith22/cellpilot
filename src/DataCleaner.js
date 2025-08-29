/**
 * CellPilot Data Cleaning Module
 * Handles all data cleaning and standardization operations
 */

const DataCleaner = {
  
  // ML-enhanced properties
  mlEnabled: false,
  adaptiveThreshold: 0.85,
  userFeedbackHistory: [],
  
  /**
   * Show duplicate removal options interface
   */
  showDuplicateRemovalOptions: function() {
    return Utils.safeExecute(() => {
      return UIComponents.buildDuplicateRemovalCard();
    }, UIComponents.buildLoadingCard('Loading duplicate removal options...'), 'showDuplicateRemovalOptions');
  },
  
  /**
   * Preview duplicate removal without applying changes
   */
  previewDuplicateRemoval: function(formInputs) {
    return Utils.safeExecute(() => {
      const range = SpreadsheetApp.getActiveRange();
      const data = range.getValues();
      
      if (data.length === 0) {
        throw new Error('No data selected');
      }
      
      // Check usage limits
      const usageCheck = UsageTracker.checkLimits('operations');
      if (!usageCheck.allowed) {
        return UIComponents.buildUpgradeCard();
      }
      
      // Get settings from form or defaults
      const threshold = (formInputs && formInputs.similarityThreshold) || 
                       UserSettings.load('duplicateThreshold', 0.85);
      const caseSensitive = (formInputs && formInputs.caseSensitive) || 
                           UserSettings.load('caseSensitiveDuplicates', false);
      
      const duplicates = this.findDuplicates(data, { threshold, caseSensitive });
      const previewData = this.preparePreviewData(data, duplicates);
      
      return UIComponents.buildDuplicateRemovalCard(previewData);
    }, Utils.handleError(new Error('Preview failed'), 'Failed to preview duplicates'), 'previewDuplicateRemoval');
  },
  
  /**
   * Apply duplicate removal to the selected range
   */
  applyDuplicateRemoval: function(formInputs) {
    return Utils.safeExecute(() => {
      const range = SpreadsheetApp.getActiveRange();
      const data = range.getValues();
      
      if (data.length === 0) {
        throw new Error('No data selected');
      }
      
      // Check and track usage
      const usageCheck = UsageTracker.track('operations');
      if (!usageCheck.allowed) {
        return UIComponents.buildUpgradeCard();
      }
      
      // Create backup if enabled
      if (UserSettings.load('autoBackup', true)) {
        const backupName = Utils.createBackup(range, 'Before Duplicate Removal');
        Logger.info('Backup created:', backupName);
      }
      
      // Get settings
      const threshold = (formInputs && formInputs.similarityThreshold) || 
                       UserSettings.load('duplicateThreshold', 0.85);
      const caseSensitive = (formInputs && formInputs.caseSensitive) || 
                           UserSettings.load('caseSensitiveDuplicates', false);
      
      const duplicates = this.findDuplicates(data, { threshold, caseSensitive });
      const cleanedData = this.removeDuplicates(data, duplicates);
      
      // Apply changes
      range.clearContent();
      if (cleanedData.length > 0) {
        range.getSheet().getRange(range.getRow(), range.getColumn(), 
          cleanedData.length, cleanedData[0].length).setValues(cleanedData);
      }
      
      const removedCount = data.length - cleanedData.length;
      return UIComponents.buildSuccessCard(
        `Removed ${removedCount} duplicate rows`,
        `Processed ${data.length} rows, kept ${cleanedData.length} unique rows`
      );
      
    }, Utils.handleError(new Error('Duplicate removal failed'), 'Failed to remove duplicates'), 'applyDuplicateRemoval');
  },
  
  /**
   * Find duplicate rows in data
   * Enhanced with ML-powered adaptive detection
   */
  findDuplicates: function(data, options = {}) {
    // Load adaptive threshold from user preferences
    const userProfile = UserSettings.load('mlProfile', {});
    const adaptiveThreshold = userProfile.adaptiveThresholds?.duplicate || this.adaptiveThreshold;
    
    const threshold = options.threshold || adaptiveThreshold;
    const caseSensitive = options.caseSensitive || false;
    const useML = options.useML !== false && this.mlEnabled;
    
    const duplicateGroups = [];
    const processed = new Set();
    
    for (let i = 0; i < data.length; i++) {
      if (processed.has(i)) continue;
      
      const currentGroup = [i];
      const currentRow = data[i];
      
      for (let j = i + 1; j < data.length; j++) {
        if (processed.has(j)) continue;
        
        const compareRow = data[j];
        const similarity = this.calculateRowSimilarity(currentRow, compareRow, caseSensitive);
        
        if (similarity >= threshold) {
          currentGroup.push(j);
          processed.add(j);
        }
      }
      
      if (currentGroup.length > 1) {
        duplicateGroups.push(currentGroup);
      }
      
      processed.add(i);
    }
    
    return duplicateGroups;
  },
  
  /**
   * Calculate similarity between two rows
   */
  calculateRowSimilarity: function(row1, row2, caseSensitive = false) {
    if (row1.length !== row2.length) return 0;
    
    let totalSimilarity = 0;
    let validComparisons = 0;
    
    for (let i = 0; i < row1.length; i++) {
      const val1 = caseSensitive ? String(row1[i]) : String(row1[i]).toLowerCase();
      const val2 = caseSensitive ? String(row2[i]) : String(row2[i]).toLowerCase();
      
      if (val1.trim() === '' && val2.trim() === '') continue;
      
      if (val1 === val2) {
        totalSimilarity += 1;
      } else {
        totalSimilarity += Utils.fuzzyMatch(val1, val2, 0);
      }
      
      validComparisons++;
    }
    
    return validComparisons > 0 ? totalSimilarity / validComparisons : 0;
  },
  
  /**
   * Remove duplicate rows, keeping the first occurrence
   */
  removeDuplicates: function(data, duplicateGroups) {
    const toRemove = new Set();
    
    duplicateGroups.forEach(group => {
      // Keep first row, mark others for removal
      for (let i = 1; i < group.length; i++) {
        toRemove.add(group[i]);
      }
    });
    
    return data.filter((row, index) => !toRemove.has(index));
  },
  
  /**
   * Prepare preview data for display
   */
  preparePreviewData: function(data, duplicateGroups) {
    const previewData = [];
    
    duplicateGroups.slice(0, 5).forEach(group => {
      for (let i = 1; i < group.length; i++) {
        const duplicateRow = data[group[i]];
        const originalRow = data[group[0]];
        
        previewData.push({
          before: duplicateRow.join(' | '),
          after: '[REMOVED - Duplicate of row ' + (group[0] + 1) + ']'
        });
      }
    });
    
    return previewData;
  },
  
  /**
   * Show text standardization options
   */
  showTextStandardizationOptions: function() {
    return Utils.safeExecute(() => {
      return UIComponents.buildTextStandardizationCard();
    }, UIComponents.buildLoadingCard('Loading text standardization options...'), 'showTextStandardizationOptions');
  },
  
  /**
   * Preview text standardization changes
   */
  previewTextStandardization: function(formInputs) {
    return Utils.safeExecute(() => {
      const range = SpreadsheetApp.getActiveRange();
      const data = range.getValues();
      
      if (data.length === 0) {
        throw new Error('No data selected');
      }
      
      const options = this.parseTextStandardizationOptions(formInputs);
      const previewData = this.prepareTextPreviewData(data, options);
      
      return UIComponents.buildTextStandardizationCard().addSection(
        UIComponents.buildPreviewSection(previewData)
      );
    }, Utils.handleError(new Error('Preview failed'), 'Failed to preview text changes'), 'previewTextStandardization');
  },
  
  /**
   * Apply text standardization
   */
  applyTextStandardization: function(formInputs) {
    return Utils.safeExecute(() => {
      const range = SpreadsheetApp.getActiveRange();
      const data = range.getValues();
      
      if (data.length === 0) {
        throw new Error('No data selected');
      }
      
      // Check and track usage
      const usageCheck = UsageTracker.track('operations');
      if (!usageCheck.allowed) {
        return UIComponents.buildUpgradeCard();
      }
      
      // Create backup
      if (UserSettings.load('autoBackup', true)) {
        Utils.createBackup(range, 'Before Text Standardization');
      }
      
      const options = this.parseTextStandardizationOptions(formInputs);
      const standardizedData = this.standardizeTextData(data, options);
      
      // Apply changes
      range.setValues(standardizedData);
      
      return UIComponents.buildSuccessCard(
        'Text standardization complete',
        `Processed ${data.length} rows with selected formatting options`
      );
      
    }, Utils.handleError(new Error('Text standardization failed'), 'Failed to standardize text'), 'applyTextStandardization');
  },
  
  /**
   * Parse text standardization options from form
   */
  parseTextStandardizationOptions: function(formInputs) {
    return {
      case: (formInputs && formInputs.textCase) || 'original',
      trim: (formInputs && formInputs.trimSpaces) !== false,
      removeExtraSpaces: (formInputs && formInputs.removeExtraSpaces) !== false,
      removeSpecialChars: (formInputs && formInputs.removeSpecialChars) || false
    };
  },
  
  /**
   * Standardize text data according to options
   */
  standardizeTextData: function(data, options) {
    return data.map(row => {
      return row.map(cell => {
        if (typeof cell === 'string') {
          return Utils.cleanText(cell, options);
        }
        return cell;
      });
    });
  },
  
  /**
   * Prepare text standardization preview data
   */
  prepareTextPreviewData: function(data, options) {
    const previewData = [];
    const maxSamples = 5;
    let sampleCount = 0;
    
    for (let i = 0; i < data.length && sampleCount < maxSamples; i++) {
      for (let j = 0; j < data[i].length && sampleCount < maxSamples; j++) {
        const original = data[i][j];
        if (typeof original === 'string' && original.trim() !== '') {
          const cleaned = Utils.cleanText(original, options);
          if (original !== cleaned) {
            previewData.push({
              before: original,
              after: cleaned
            });
            sampleCount++;
          }
        }
      }
    }
    
    return previewData;
  },
  
  /**
   * Show date formatting options
   */
  showDateFormattingOptions: function() {
    return Utils.safeExecute(() => {
      // TODO: Implement date formatting UI
      return UIComponents.buildLoadingCard('Date formatting coming soon...');
    }, Utils.handleError(new Error('Feature unavailable'), 'Date formatting not yet available'), 'showDateFormattingOptions');
  },
  
  /**
   * ML Methods - Learn from user feedback
   */
  learnFromDuplicateFeedback: function(acceptedDuplicates, rejectedDuplicates, threshold) {
    // Track user feedback
    this.userFeedbackHistory.push({
      accepted: acceptedDuplicates.length,
      rejected: rejectedDuplicates.length,
      threshold: threshold,
      timestamp: new Date().toISOString()
    });
    
    // Calculate acceptance rate
    const totalSuggestions = acceptedDuplicates.length + rejectedDuplicates.length;
    if (totalSuggestions === 0) return;
    
    const acceptanceRate = acceptedDuplicates.length / totalSuggestions;
    
    // Adjust adaptive threshold based on user behavior
    if (acceptanceRate < 0.7) {
      // Too many false positives, increase threshold (be more strict)
      this.adaptiveThreshold = Math.min(0.95, this.adaptiveThreshold + 0.02);
      Logger.debug('ML: Increasing duplicate threshold to', this.adaptiveThreshold);
    } else if (acceptanceRate > 0.9) {
      // Very accurate, can be slightly more aggressive
      this.adaptiveThreshold = Math.max(0.75, this.adaptiveThreshold - 0.01);
      Logger.debug('ML: Decreasing duplicate threshold to', this.adaptiveThreshold);
    }
    
    // Save updated threshold to user preferences
    const userProfile = UserSettings.load('mlProfile', {});
    userProfile.adaptiveThresholds = userProfile.adaptiveThresholds || {};
    userProfile.adaptiveThresholds.duplicate = this.adaptiveThreshold;
    userProfile.feedbackHistory = this.userFeedbackHistory.slice(-100); // Keep last 100 feedback entries
    UserSettings.save('mlProfile', userProfile);
    
    // Return learning summary
    return {
      newThreshold: this.adaptiveThreshold,
      acceptanceRate: acceptanceRate,
      totalFeedback: this.userFeedbackHistory.length
    };
  },
  
  /**
   * Enhanced similarity calculation with ML features
   */
  calculateRowSimilarityML: function(row1, row2, options = {}) {
    const caseSensitive = options.caseSensitive || false;
    
    // Extract features for ML comparison
    const features = {
      exactMatches: 0,
      fuzzyMatches: 0,
      typeMatches: 0,
      lengthSimilarity: 0,
      numericSimilarity: 0,
      patternMatch: 0
    };
    
    const minLength = Math.min(row1.length, row2.length);
    const maxLength = Math.max(row1.length, row2.length);
    
    for (let i = 0; i < minLength; i++) {
      const val1 = caseSensitive ? String(row1[i]) : String(row1[i]).toLowerCase();
      const val2 = caseSensitive ? String(row2[i]) : String(row2[i]).toLowerCase();
      
      // Exact match
      if (val1 === val2) {
        features.exactMatches++;
      }
      
      // Fuzzy match
      const fuzzyScore = Utils.fuzzyMatch(val1, val2, 0.7);
      if (fuzzyScore > 0) {
        features.fuzzyMatches += fuzzyScore;
      }
      
      // Type match
      const type1 = this.getValueType(row1[i]);
      const type2 = this.getValueType(row2[i]);
      if (type1 === type2) {
        features.typeMatches++;
      }
      
      // Length similarity
      features.lengthSimilarity += 1 - Math.abs(val1.length - val2.length) / Math.max(val1.length, val2.length, 1);
      
      // Numeric similarity
      if (!isNaN(row1[i]) && !isNaN(row2[i])) {
        const num1 = Number(row1[i]);
        const num2 = Number(row2[i]);
        features.numericSimilarity += 1 - Math.abs(num1 - num2) / Math.max(Math.abs(num1), Math.abs(num2), 1);
      }
    }
    
    // Pattern matching
    const pattern1 = row1.map(v => this.getValueType(v)).join('-');
    const pattern2 = row2.map(v => this.getValueType(v)).join('-');
    features.patternMatch = pattern1 === pattern2 ? 1 : 0;
    
    // Combine features with learned weights
    const userProfile = UserSettings.load('mlProfile', {});
    const weights = userProfile.featureWeights || {
      exactMatches: 0.4,
      fuzzyMatches: 0.2,
      typeMatches: 0.1,
      lengthSimilarity: 0.1,
      numericSimilarity: 0.1,
      patternMatch: 0.1
    };
    
    let totalScore = 0;
    totalScore += (features.exactMatches / maxLength) * weights.exactMatches;
    totalScore += (features.fuzzyMatches / maxLength) * weights.fuzzyMatches;
    totalScore += (features.typeMatches / maxLength) * weights.typeMatches;
    totalScore += (features.lengthSimilarity / maxLength) * weights.lengthSimilarity;
    totalScore += (features.numericSimilarity / Math.max(1, minLength)) * weights.numericSimilarity;
    totalScore += features.patternMatch * weights.patternMatch;
    
    return {
      score: totalScore,
      features: features,
      confidence: this.calculateConfidence(features, maxLength)
    };
  },
  
  /**
   * Get value type for ML analysis
   */
  getValueType: function(value) {
    if (value === null || value === undefined || value === '') return 'empty';
    if (!isNaN(Number(value))) return 'number';
    if (/^\d{1,2}\/\d{1,2}\/\d{2,4}$/.test(value)) return 'date';
    if (/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value)) return 'email';
    if (/^[\+]?[\d\s\-\(\)]{10,}$/.test(value)) return 'phone';
    if (/^[$£€¥]?[\d,]+\.?\d*$/.test(value)) return 'currency';
    return 'text';
  },
  
  /**
   * Calculate confidence score for duplicate detection
   */
  calculateConfidence: function(features, totalCells) {
    // Base confidence on multiple factors
    const exactMatchRatio = features.exactMatches / totalCells;
    const fuzzyMatchRatio = features.fuzzyMatches / totalCells;
    const typeMatchRatio = features.typeMatches / totalCells;
    
    // Weighted confidence calculation
    const confidence = (exactMatchRatio * 0.5) + 
                      (fuzzyMatchRatio * 0.3) + 
                      (typeMatchRatio * 0.2);
    
    return Math.min(1, confidence);
  },
  
  /**
   * Enable ML features for data cleaning
   */
  enableML: function() {
    this.mlEnabled = true;
    Logger.info('ML features enabled for DataCleaner');
    
    // Load user profile
    const userProfile = UserSettings.load('mlProfile', {});
    if (userProfile.adaptiveThresholds?.duplicate) {
      this.adaptiveThreshold = userProfile.adaptiveThresholds.duplicate;
    }
    if (userProfile.feedbackHistory) {
      this.userFeedbackHistory = userProfile.feedbackHistory;
    }
    
    return true;
  },
  
  /**
   * Get ML status and metrics
   */
  getMLStatus: function() {
    return {
      enabled: this.mlEnabled,
      adaptiveThreshold: this.adaptiveThreshold,
      feedbackCount: this.userFeedbackHistory.length,
      lastFeedback: this.userFeedbackHistory[this.userFeedbackHistory.length - 1] || null,
      averageAcceptanceRate: this.calculateAverageAcceptanceRate()
    };
  },
  
  /**
   * Calculate average acceptance rate from feedback history
   */
  calculateAverageAcceptanceRate: function() {
    if (this.userFeedbackHistory.length === 0) return null;
    
    const recentFeedback = this.userFeedbackHistory.slice(-20);
    let totalAccepted = 0;
    let totalSuggestions = 0;
    
    recentFeedback.forEach(feedback => {
      totalAccepted += feedback.accepted || 0;
      totalSuggestions += (feedback.accepted || 0) + (feedback.rejected || 0);
    });
    
    return totalSuggestions > 0 ? totalAccepted / totalSuggestions : null;
  }
};