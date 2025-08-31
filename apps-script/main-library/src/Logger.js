/**
 * CellPilot Logger
 * Centralized logging system with environment-aware output
 */

const Logger = {
  // Configuration
  config: {
    enabled: false, // Set to true only in development
    logLevel: 'error', // 'debug', 'info', 'warn', 'error'
    sendToAnalytics: true,
    maxLogSize: 100 // Maximum number of logs to keep in memory
  },
  
  // In-memory log storage for debugging
  logs: [],
  
  // Log levels
  levels: {
    debug: 0,
    info: 1,
    warn: 2,
    error: 3
  },
  
  /**
   * Check if we're in development mode
   */
  isDevelopment: function() {
    // Check for development indicators
    const isDev = UserSettings.load('developerMode', false);
    const betaMode = typeof FeatureGate !== 'undefined' && FeatureGate.BETA_CONFIG.enabled;
    
    // Only enable console logging if explicitly enabled (removed email check for beta)
    return isDev;
  },
  
  /**
   * Main logging function
   */
  log: function(level, message, data) {
    // Store in memory for debugging
    const logEntry = {
      timestamp: new Date().toISOString(),
      level: level,
      message: message,
      data: data
    };
    
    this.logs.push(logEntry);
    if (this.logs.length > this.config.maxLogSize) {
      this.logs.shift();
    }
    
    // Check if we should output this log level
    const currentLevel = this.levels[this.config.logLevel] || this.levels.error;
    const messageLevel = this.levels[level] || this.levels.info;
    
    if (messageLevel < currentLevel) {
      return;
    }
    
    // Only output to console in development
    if (this.isDevelopment() && this.config.enabled) {
      const consoleMethod = console[level] || console.log;
      if (data) {
        consoleMethod(`[CellPilot ${level.toUpperCase()}] ${message}`, data);
      } else {
        consoleMethod(`[CellPilot ${level.toUpperCase()}] ${message}`);
      }
    }
    
    // Send errors to analytics
    if (level === 'error' && this.config.sendToAnalytics) {
      try {
        if (typeof ApiIntegration !== 'undefined' && ApiIntegration.reportError) {
          ApiIntegration.reportError({
            level: level,
            message: message,
            data: data,
            stack: new Error().stack
          });
        }
      } catch (e) {
        // Fail silently to avoid recursive errors
      }
    }
  },
  
  /**
   * Convenience methods
   */
  debug: function(message, data) {
    this.log('debug', message, data);
  },
  
  info: function(message, data) {
    this.log('info', message, data);
  },
  
  warn: function(message, data) {
    this.log('warn', message, data);
  },
  
  error: function(message, data) {
    this.log('error', message, data);
  },
  
  /**
   * Get recent logs for debugging
   */
  getRecentLogs: function(count) {
    count = count || 10;
    return this.logs.slice(-count);
  },
  
  /**
   * Clear all logs
   */
  clearLogs: function() {
    this.logs = [];
  },
  
  /**
   * Export logs for debugging
   */
  exportLogs: function() {
    return JSON.stringify(this.logs, null, 2);
  },
  
  /**
   * Enable/disable logging (for development)
   */
  setEnabled: function(enabled) {
    this.config.enabled = enabled;
    UserSettings.save('developerMode', enabled);
  },
  
  /**
   * Set log level
   */
  setLogLevel: function(level) {
    if (this.levels.hasOwnProperty(level)) {
      this.config.logLevel = level;
    }
  }
};

// Export for use in other modules
if (typeof module !== 'undefined' && module.exports) {
  module.exports = Logger;
}