/**
 * CellPilot Configuration and User Settings
 * Handles all persistent storage and user preferences
 */

const UserSettings = {
  
  /**
   * Initialize user settings with defaults
   */
  initialize: function() {
    const defaults = this.getDefaults();
    
    // Only set defaults if user hasn't configured anything yet
    Object.keys(defaults).forEach(key => {
      if (this.load(key) === null) {
        this.save(key, defaults[key]);
      }
    });
    
    console.log('User settings initialized');
  },
  
  /**
   * Save a setting value
   */
  save: function(key, value) {
    try {
      PropertiesService.getUserProperties().setProperty(
        `cellpilot_${key}`, 
        JSON.stringify(value)
      );
      return true;
    } catch (error) {
      console.error(`Failed to save setting ${key}:`, error);
      return false;
    }
  },
  
  /**
   * Load a setting value
   */
  load: function(key, defaultValue = null) {
    try {
      const value = PropertiesService.getUserProperties().getProperty(`cellpilot_${key}`);
      return value ? JSON.parse(value) : defaultValue;
    } catch (error) {
      console.error(`Failed to load setting ${key}:`, error);
      return defaultValue;
    }
  },
  
  /**
   * Remove a setting
   */
  remove: function(key) {
    try {
      PropertiesService.getUserProperties().deleteProperty(`cellpilot_${key}`);
      return true;
    } catch (error) {
      console.error(`Failed to remove setting ${key}:`, error);
      return false;
    }
  },
  
  /**
   * Get all user settings
   */
  getAll: function() {
    try {
      const allProps = PropertiesService.getUserProperties().getProperties();
      const cellpilotProps = {};
      
      Object.keys(allProps).forEach(key => {
        if (key.startsWith('cellpilot_')) {
          const cleanKey = key.replace('cellpilot_', '');
          cellpilotProps[cleanKey] = JSON.parse(allProps[key]);
        }
      });
      
      return cellpilotProps;
    } catch (error) {
      console.error('Failed to get all settings:', error);
      return {};
    }
  },
  
  /**
   * Default configuration values
   */
  getDefaults: function() {
    return {
      // Data cleaning preferences
      duplicateThreshold: 0.85,
      caseSensitiveDuplicates: false,
      previewEnabled: true,
      autoBackup: false, // Disabled in favor of undo system
      
      // Formula builder preferences
      formulaComplexity: 'simple', // simple, intermediate, advanced
      showFormulaExplanations: true,
      suggestOptimizations: true,
      
      // Automation preferences
      emailTemplate: 'standard',
      dateFormat: 'auto',
      confirmBeforeActions: true,
      
      // UI preferences
      showTips: true,
      compactView: false,
      theme: 'default',
      
      // Usage tracking (for limits)
      monthlyUsage: {},
      userTier: 'free', // free, starter, professional, business
      
      // Performance settings
      batchSize: 1000,
      enableCache: true,
      
      // Feature flags
      betaFeatures: false
    };
  }
};

/**
 * Application constants
 */
const Config = {
  
  // Version info
  VERSION: '1.0.0',
  BUILD: 'alpha',
  
  // Usage limits by tier
  USAGE_LIMITS: {
    free: {
      operations: 25,
      emails: 0,
      automations: 0
    },
    starter: {
      operations: 500,
      emails: 10,
      automations: 5
    },
    professional: {
      operations: -1, // unlimited
      emails: -1,
      automations: -1
    },
    business: {
      operations: -1,
      emails: -1,
      automations: -1
    }
  },
  
  // Feature access by tier
  FEATURE_ACCESS: {
    data_cleaning: ['free', 'starter', 'professional', 'business'],
    formula_builder: ['starter', 'professional', 'business'],
    automation: ['professional', 'business'],
    industry_tools: ['professional', 'business'],
    team_features: ['business']
  },
  
  // Performance settings
  PERFORMANCE: {
    MAX_BATCH_SIZE: 10000,
    CACHE_DURATION: 300000, // 5 minutes
    API_TIMEOUT: 30000 // 30 seconds
  },
  
  // UI constants
  UI: {
    COLORS: {
      primary: '#1a73e8',
      success: '#34a853',
      warning: '#fbbc04',
      error: '#ea4335'
    },
    ICONS: {
      clean: 'https://fonts.gstatic.com/s/i/materialicons/cleaning_services/v6/24px.svg',
      formula: 'https://fonts.gstatic.com/s/i/materialicons/functions/v8/24px.svg',
      automation: 'https://fonts.gstatic.com/s/i/materialicons/smart_toy/v7/24px.svg'
    }
  }
};

/**
 * Usage tracking for tier limitations
 */
const UsageTracker = {
  
  /**
   * Track an operation for usage limits
   */
  track: function(operation) {
    const currentMonth = new Date().getMonth();
    const currentYear = new Date().getFullYear();
    const key = `${currentYear}_${currentMonth}`;
    
    const usage = UserSettings.load('monthlyUsage', {});
    
    if (!usage[key]) {
      usage[key] = {};
    }
    
    usage[key][operation] = (usage[key][operation] || 0) + 1;
    
    UserSettings.save('monthlyUsage', usage);
    
    return this.checkLimits(operation);
  },
  
  /**
   * Check if user has hit limits for operation
   */
  checkLimits: function(operation) {
    const userTier = UserSettings.load('userTier', 'free');
    const limits = Config.USAGE_LIMITS[userTier];
    
    if (!limits || limits[operation] === -1) {
      return { allowed: true };
    }
    
    const currentUsage = this.getCurrentUsage();
    const used = currentUsage[operation] || 0;
    
    return {
      allowed: used < limits[operation],
      used: used,
      limit: limits[operation],
      remaining: Math.max(0, limits[operation] - used)
    };
  },
  
  /**
   * Get current month usage
   */
  getCurrentUsage: function() {
    const currentMonth = new Date().getMonth();
    const currentYear = new Date().getFullYear();
    const key = `${currentYear}_${currentMonth}`;
    
    const allUsage = UserSettings.load('monthlyUsage', {});
    return allUsage[key] || {};
  },
  
  /**
   * Reset usage (for testing or new billing cycle)
   */
  reset: function() {
    UserSettings.save('monthlyUsage', {});
  }
};