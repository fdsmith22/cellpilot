/**
 * CellPilot Data Cleaning Module
 * Handles all data cleaning and standardization operations
 */

const DataCleaner = {
  
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
        console.log('Backup created:', backupName);
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
   */
  findDuplicates: function(data, options = {}) {
    const threshold = options.threshold || 0.85;
    const caseSensitive = options.caseSensitive || false;
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
  }
};