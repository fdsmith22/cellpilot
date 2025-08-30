/**
 * ML Backend Module
 * Handles server-side ML operations and settings management
 */

const MLBackend = {
  /**
   * Enable ML features for the current user
   */
  enableMLFeatures: function() {
    try {
      const settings = UserSettings.load('mlSettings', {});
      settings.enabled = true;
      settings.enabledDate = new Date().toISOString();
      settings.modelVersion = '1.0.0';
      
      UserSettings.save('mlSettings', settings);
      
      // Initialize ML profile if not exists
      if (!UserSettings.load('mlProfile')) {
        this.initializeMLProfile();
      }
      
      // Track ML activation
      if (typeof ApiIntegration !== 'undefined') {
        ApiIntegration.trackEvent('ml_enabled', {
          timestamp: new Date().toISOString()
        });
      }
      
      return {
        success: true,
        message: 'ML features enabled successfully'
      };
    } catch (error) {
      Logger.error('Failed to enable ML features:', error);
      return {
        success: false,
        error: error.toString()
      };
    }
  },
  
  /**
   * Disable ML features
   */
  disableMLFeatures: function() {
    try {
      const settings = UserSettings.load('mlSettings', {});
      settings.enabled = false;
      settings.disabledDate = new Date().toISOString();
      
      UserSettings.save('mlSettings', settings);
      
      // Track ML deactivation
      if (typeof ApiIntegration !== 'undefined') {
        ApiIntegration.trackEvent('ml_disabled', {
          timestamp: new Date().toISOString()
        });
      }
      
      return {
        success: true,
        message: 'ML features disabled'
      };
    } catch (error) {
      Logger.error('Failed to disable ML features:', error);
      return {
        success: false,
        error: error.toString()
      };
    }
  },
  
  /**
   * Get ML status
   */
  getMLStatus: function() {
    try {
      const settings = UserSettings.load('mlSettings', {});
      return {
        enabled: settings.enabled || false,
        modelVersion: settings.modelVersion || '1.0.0',
        lastUpdated: settings.enabledDate || null
      };
    } catch (error) {
      Logger.error('Failed to get ML status:', error);
      return {
        enabled: false
      };
    }
  },
  
  /**
   * Initialize ML profile with default values
   */
  initializeMLProfile: function() {
    const defaultProfile = {
      duplicateThreshold: 0.85,
      confidenceThreshold: 0.7,
      preferredFormulas: [],
      columnTypeHistory: {},
      feedbackHistory: [],
      adaptiveThresholds: {
        duplicate: 0.85,
        similarity: 0.8,
        anomaly: 0.95,
        textCleaning: 0.75,
        formulaConfidence: 0.7,
        validationStrictness: 0.8
      },
      usagePatterns: {
        mostUsedFeatures: {},
        timeOfDayPreferences: {},
        dataTypeFrequency: {}
      }
    };
    
    UserSettings.save('mlProfile', defaultProfile);
    return defaultProfile;
  },
  
  /**
   * Get user's ML learning profile
   */
  getUserLearningProfile: function() {
    try {
      let profile = UserSettings.load('mlProfile');
      if (!profile) {
        profile = this.initializeMLProfile();
      }
      return profile;
    } catch (error) {
      Logger.error('Failed to get ML profile:', error);
      return this.initializeMLProfile();
    }
  },
  
  /**
   * Save user's ML learning profile
   */
  saveUserLearningProfile: function(profile) {
    try {
      UserSettings.save('mlProfile', profile);
      return {
        success: true
      };
    } catch (error) {
      Logger.error('Failed to save ML profile:', error);
      return {
        success: false,
        error: error.toString()
      };
    }
  },
  
  /**
   * Get adaptive duplicate threshold based on ML learning
   */
  getAdaptiveDuplicateThreshold: function() {
    try {
      const profile = this.getUserLearningProfile();
      return profile.adaptiveThresholds.duplicate || 0.85;
    } catch (error) {
      Logger.error('Failed to get adaptive threshold:', error);
      return 0.85; // Default threshold
    }
  },
  
  /**
   * Track ML feedback for continuous learning
   */
  trackMLFeedback: function(operation, value, action, metadata) {
    try {
      const profile = this.getUserLearningProfile();
      
      if (!profile.feedbackHistory) {
        profile.feedbackHistory = [];
      }
      
      // Add feedback entry
      profile.feedbackHistory.push({
        operation: operation,
        value: value,
        action: action,
        metadata: metadata || {},
        timestamp: new Date().toISOString()
      });
      
      // Keep only last 100 feedback entries
      if (profile.feedbackHistory.length > 100) {
        profile.feedbackHistory = profile.feedbackHistory.slice(-100);
      }
      
      // Update adaptive thresholds based on feedback
      this.updateAdaptiveThresholds(profile, operation, value, action);
      
      // Save updated profile
      this.saveUserLearningProfile(profile);
      
      return {
        success: true,
        updatedThreshold: profile.adaptiveThresholds[operation] || null
      };
    } catch (error) {
      Logger.error('Failed to track ML feedback:', error);
      return {
        success: false,
        error: error.toString()
      };
    }
  },
  
  /**
   * Update adaptive thresholds based on user feedback
   */
  updateAdaptiveThresholds: function(profile, operation, value, action) {
    const learningRate = 0.02; // How quickly to adjust
    
    switch(operation) {
      case 'duplicateRemoval':
        if (action === 'accepted') {
          // User accepted the results, threshold was good
          // Slightly move threshold towards the used value
          profile.adaptiveThresholds.duplicate = 
            profile.adaptiveThresholds.duplicate * 0.95 + value * 0.05;
        } else if (action === 'rejected') {
          // User rejected, adjust threshold away from used value
          if (value < profile.adaptiveThresholds.duplicate) {
            // Threshold was too low, increase it
            profile.adaptiveThresholds.duplicate = 
              Math.min(0.95, profile.adaptiveThresholds.duplicate + learningRate);
          } else {
            // Threshold was too high, decrease it
            profile.adaptiveThresholds.duplicate = 
              Math.max(0.65, profile.adaptiveThresholds.duplicate - learningRate);
          }
        }
        break;
        
      case 'formulaSuggestion':
        if (action === 'used') {
          // User used the suggestion, increase confidence
          profile.adaptiveThresholds.formulaConfidence = 
            Math.min(0.95, profile.adaptiveThresholds.formulaConfidence + learningRate/2);
        } else if (action === 'ignored') {
          // User ignored suggestion, decrease confidence threshold
          profile.adaptiveThresholds.formulaConfidence = 
            Math.max(0.5, profile.adaptiveThresholds.formulaConfidence - learningRate/2);
        }
        break;
    }
  },
  
  /**
   * Get ML storage information
   */
  getMLStorageInfo: function() {
    try {
      const profile = UserSettings.load('mlProfile', {});
      const settings = UserSettings.load('mlSettings', {});
      
      // Estimate storage size (rough calculation)
      const profileSize = JSON.stringify(profile).length;
      const settingsSize = JSON.stringify(settings).length;
      const totalSize = profileSize + settingsSize;
      
      // Add any cached model data size
      const cacheSize = this.getMLCacheSize();
      
      return {
        used: totalSize + cacheSize,
        total: 5 * 1024 * 1024, // 5MB limit for user properties
        breakdown: {
          profile: profileSize,
          settings: settingsSize,
          cache: cacheSize
        }
      };
    } catch (error) {
      Logger.error('Failed to get ML storage info:', error);
      return {
        used: 0,
        total: 5 * 1024 * 1024
      };
    }
  },
  
  /**
   * Get ML cache size
   */
  getMLCacheSize: function() {
    try {
      // Check for any cached ML data
      const cache = CacheService.getUserCache();
      const keys = ['ml_models', 'ml_predictions', 'ml_features'];
      let totalSize = 0;
      
      keys.forEach(key => {
        const data = cache.get(key);
        if (data) {
          totalSize += data.length;
        }
      });
      
      return totalSize;
    } catch (error) {
      return 0;
    }
  },
  
  /**
   * Clear ML data
   */
  clearMLData: function() {
    try {
      // Clear ML profile
      UserSettings.remove('mlProfile');
      
      // Clear ML settings but keep enabled status
      const settings = UserSettings.load('mlSettings', {});
      const wasEnabled = settings.enabled;
      UserSettings.save('mlSettings', {
        enabled: wasEnabled,
        clearedDate: new Date().toISOString()
      });
      
      // Clear ML cache
      const cache = CacheService.getUserCache();
      cache.remove('ml_models');
      cache.remove('ml_predictions');
      cache.remove('ml_features');
      
      // Reinitialize profile if ML is enabled
      if (wasEnabled) {
        this.initializeMLProfile();
      }
      
      return {
        success: true,
        message: 'ML data cleared successfully'
      };
    } catch (error) {
      Logger.error('Failed to clear ML data:', error);
      return {
        success: false,
        error: error.toString()
      };
    }
  },
  
  /**
   * Export ML profile for backup
   */
  exportMLProfile: function() {
    try {
      const profile = this.getUserLearningProfile();
      const settings = UserSettings.load('mlSettings', {});
      
      return {
        profile: profile,
        settings: settings,
        exportDate: new Date().toISOString(),
        version: '1.0.0'
      };
    } catch (error) {
      Logger.error('Failed to export ML profile:', error);
      throw error;
    }
  },
  
  /**
   * Import ML profile from backup
   */
  importMLProfile: function(data) {
    try {
      if (data.profile) {
        UserSettings.save('mlProfile', data.profile);
      }
      
      if (data.settings) {
        UserSettings.save('mlSettings', data.settings);
      }
      
      return {
        success: true,
        message: 'ML profile imported successfully'
      };
    } catch (error) {
      Logger.error('Failed to import ML profile:', error);
      return {
        success: false,
        error: error.toString()
      };
    }
  }
};

// Export functions for Google Apps Script
function enableMLFeatures() {
  return MLBackend.enableMLFeatures();
}

function disableMLFeatures() {
  return MLBackend.disableMLFeatures();
}

function getMLStatus() {
  return MLBackend.getMLStatus();
}

function getUserLearningProfile() {
  return MLBackend.getUserLearningProfile();
}

function saveUserLearningProfile(profile) {
  return MLBackend.saveUserLearningProfile(profile);
}

function getAdaptiveDuplicateThreshold() {
  return MLBackend.getAdaptiveDuplicateThreshold();
}

function trackMLFeedback(operation, value, action, metadata) {
  return MLBackend.trackMLFeedback(operation, value, action, metadata);
}

function getMLStorageInfo() {
  return MLBackend.getMLStorageInfo();
}

function clearMLData() {
  return MLBackend.clearMLData();
}

function exportMLProfile() {
  return MLBackend.exportMLProfile();
}

function importMLProfile(data) {
  return MLBackend.importMLProfile(data);
}