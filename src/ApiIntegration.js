/**
 * API Integration for CellPilot
 * Handles communication with the CellPilot backend
 */

const ApiIntegration = {
  // API configuration
  API_URL: 'https://api.cellpilot.io',
  
  /**
   * Get API key from script properties
   * Returns null if not configured (beta mode allows null)
   */
  getApiKey: function() {
    try {
      const key = PropertiesService.getScriptProperties().getProperty('CELLPILOT_API_KEY');
      if (!key && typeof FeatureGate !== 'undefined' && FeatureGate.isBetaPeriod()) {
        // In beta, API key is optional
        return null;
      }
      return key;
    } catch (e) {
      return null;
    }
  },
  
  /**
   * Track installation event
   */
  trackInstallation: function() {
    try {
      const userId = this.getUserId();
      const data = {
        event: 'installation',
        userId: userId,
        data: {
          method: 'library', // or 'marketplace'
          version: '1.0.0',
          source: 'direct',
          timestamp: new Date().toISOString()
        }
      };
      
      this.sendTrackingEvent(data);
    } catch (error) {
      Logger.error('Failed to track installation:', error);
    }
  },
  
  /**
   * Track feature usage
   */
  trackUsage: function(feature, operation, count = 1) {
    try {
      const userId = this.getUserId();
      const data = {
        event: 'usage',
        userId: userId,
        data: {
          feature: feature,
          operation: operation,
          count: count,
          timestamp: new Date().toISOString()
        }
      };
      
      this.sendTrackingEvent(data);
    } catch (error) {
      Logger.error('Failed to track usage:', error);
    }
  },
  
  /**
   * Check subscription status
   */
  checkSubscription: function() {
    try {
      const apiKey = this.getApiKey();
      
      // In beta mode without API key, return free tier
      if (!apiKey) {
        if (typeof FeatureGate !== 'undefined' && FeatureGate.isBetaPeriod()) {
          return { plan: 'beta', operationsLimit: -1 };
        }
        return { plan: 'free', operationsLimit: 25 };
      }
      
      const userId = this.getUserId();
      const response = UrlFetchApp.fetch(
        `${this.API_URL}/api/subscription?userId=${userId}`,
        {
          method: 'GET',
          headers: {
            'x-api-key': apiKey,
            'Content-Type': 'application/json'
          },
          muteHttpExceptions: true
        }
      );
      
      if (response.getResponseCode() === 200) {
        const subscription = JSON.parse(response.getContentText());
        
        // Update local settings
        UserSettings.save('userTier', subscription.plan);
        UserSettings.save('operationsLimit', subscription.operationsLimit);
        
        return subscription;
      }
      
      return null;
    } catch (error) {
      Logger.error('Failed to check subscription:', error);
      return null;
    }
  },
  
  /**
   * Send tracking event to backend
   */
  sendTrackingEvent: function(data) {
    try {
      const apiKey = this.getApiKey();
      
      // Skip tracking if no API key (beta mode)
      if (!apiKey) {
        Logger.info('Skipping tracking - no API key configured');
        return true;
      }
      
      const response = UrlFetchApp.fetch(
        `${this.API_URL}/api/track`,
        {
          method: 'POST',
          headers: {
            'x-api-key': apiKey,
            'Content-Type': 'application/json'
          },
          payload: JSON.stringify(data),
          muteHttpExceptions: true
        }
      );
      
      return response.getResponseCode() === 200;
    } catch (error) {
      Logger.error('Failed to send tracking event:', error);
      return false;
    }
  },
  
  /**
   * Get or create unique user ID
   */
  getUserId: function() {
    let userId = PropertiesService.getUserProperties().getProperty('cellpilot_user_id');
    
    if (!userId) {
      // Generate unique ID based on email
      const email = Session.getActiveUser().getEmail();
      userId = Utilities.computeDigest(
        Utilities.DigestAlgorithm.SHA_256,
        email + 'cellpilot'
      ).reduce((str, byte) => str + byte.toString(16).padStart(2, '0'), '');
      
      PropertiesService.getUserProperties().setProperty('cellpilot_user_id', userId);
    }
    
    return userId;
  },
  
  /**
   * Report error to backend for monitoring
   */
  reportError: function(error, context) {
    try {
      const data = {
        event: 'error',
        userId: this.getUserId(),
        data: {
          error: error.toString(),
          stack: error.stack,
          context: context,
          timestamp: new Date().toISOString()
        }
      };
      
      this.sendTrackingEvent(data);
    } catch (e) {
      Logger.error('Failed to report error:', e);
    }
  }
};