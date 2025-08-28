// User Preference Learning System with ML
const UserPreferenceLearning = {
  // Core preference properties
  preferences: {
    features: {},
    formatting: {},
    automation: {},
    dataPatterns: {},
    workflowSequences: []
  },
  
  // Behavioral tracking
  behaviorHistory: [],
  sessionData: {},
  featureUsage: {},
  
  // ML learning parameters
  mlEnabled: false,
  confidenceThreshold: 0.75,
  learningRate: 0.1,
  minSampleSize: 5,
  
  // Prediction models
  nextActionPredictions: {},
  featureRecommendations: [],
  workflowSuggestions: [],
  
  /**
   * Initialize preference learning system
   */
  initialize: function() {
    try {
      // Check ML status
      const mlStatus = PropertiesService.getUserProperties().getProperty('cellpilot_ml_enabled');
      this.mlEnabled = mlStatus === 'true';
      
      // Load saved preferences
      const savedPreferences = PropertiesService.getUserProperties().getProperty('user_preferences');
      if (savedPreferences) {
        this.preferences = JSON.parse(savedPreferences);
      }
      
      // Load behavior history
      const savedHistory = PropertiesService.getUserProperties().getProperty('behavior_history');
      if (savedHistory) {
        this.behaviorHistory = JSON.parse(savedHistory);
      }
      
      // Load feature usage stats
      const savedUsage = PropertiesService.getUserProperties().getProperty('feature_usage');
      if (savedUsage) {
        this.featureUsage = JSON.parse(savedUsage);
      }
      
      // Start new session
      this.startSession();
      
      return true;
    } catch (error) {
      console.error('Error initializing preference learning:', error);
      return false;
    }
  },
  
  /**
   * Start a new user session
   */
  startSession: function() {
    this.sessionData = {
      sessionId: Utilities.getUuid(),
      startTime: new Date().toISOString(),
      actions: [],
      features: [],
      errors: [],
      completions: []
    };
  },
  
  /**
   * Track user action
   */
  trackAction: function(action, context = {}) {
    try {
      const timestamp = new Date().toISOString();
      
      const actionData = {
        timestamp: timestamp,
        action: action,
        context: context,
        sessionId: this.sessionData.sessionId
      };
      
      // Add to session
      this.sessionData.actions.push(actionData);
      
      // Add to history
      this.behaviorHistory.push(actionData);
      
      // Update feature usage
      this.updateFeatureUsage(action, context);
      
      // Learn patterns if ML enabled
      if (this.mlEnabled) {
        this.learnFromAction(actionData);
        this.predictNextAction(action);
      }
      
      // Keep history manageable
      if (this.behaviorHistory.length > 5000) {
        this.behaviorHistory = this.behaviorHistory.slice(-5000);
      }
      
      // Save periodically
      if (this.sessionData.actions.length % 10 === 0) {
        this.savePreferences();
      }
      
    } catch (error) {
      console.error('Error tracking action:', error);
    }
  },
  
  /**
   * Update feature usage statistics
   */
  updateFeatureUsage: function(action, context) {
    if (!this.featureUsage[action]) {
      this.featureUsage[action] = {
        count: 0,
        lastUsed: null,
        contexts: {},
        successRate: 100,
        avgDuration: 0,
        preferences: {}
      };
    }
    
    const usage = this.featureUsage[action];
    usage.count++;
    usage.lastUsed = new Date().toISOString();
    
    // Track context patterns
    const contextKey = JSON.stringify(context);
    usage.contexts[contextKey] = (usage.contexts[contextKey] || 0) + 1;
    
    // Identify preferred settings
    if (context.settings) {
      Object.entries(context.settings).forEach(([key, value]) => {
        if (!usage.preferences[key]) {
          usage.preferences[key] = {};
        }
        usage.preferences[key][value] = (usage.preferences[key][value] || 0) + 1;
      });
    }
  },
  
  /**
   * Learn from user action using ML
   */
  learnFromAction: function(actionData) {
    try {
      // Extract features from action
      const features = this.extractActionFeatures(actionData);
      
      // Update preference models
      this.updatePreferenceModels(features);
      
      // Detect workflow patterns
      this.detectWorkflowPatterns();
      
      // Update formatting preferences
      if (actionData.context.formatting) {
        this.learnFormattingPreference(actionData.context.formatting);
      }
      
      // Update data pattern preferences
      if (actionData.context.dataType) {
        this.learnDataPatternPreference(actionData.context.dataType, actionData.action);
      }
      
    } catch (error) {
      console.error('Error learning from action:', error);
    }
  },
  
  /**
   * Extract features from action for ML
   */
  extractActionFeatures: function(actionData) {
    const features = {
      action: actionData.action,
      time: new Date(actionData.timestamp).getHours(),
      dayOfWeek: new Date(actionData.timestamp).getDay(),
      previousAction: this.sessionData.actions.length > 1 
        ? this.sessionData.actions[this.sessionData.actions.length - 2].action 
        : null,
      contextType: actionData.context.type || 'unknown',
      dataSize: actionData.context.dataSize || 0,
      hasFormulas: actionData.context.hasFormulas || false,
      sheetType: actionData.context.sheetType || 'general'
    };
    
    return features;
  },
  
  /**
   * Update preference models based on features
   */
  updatePreferenceModels: function(features) {
    // Update feature preferences
    if (!this.preferences.features[features.action]) {
      this.preferences.features[features.action] = {
        frequency: 0,
        timePreferences: Array(24).fill(0),
        dayPreferences: Array(7).fill(0),
        followsActions: {},
        contexts: {}
      };
    }
    
    const pref = this.preferences.features[features.action];
    pref.frequency++;
    pref.timePreferences[features.time]++;
    pref.dayPreferences[features.dayOfWeek]++;
    
    if (features.previousAction) {
      pref.followsActions[features.previousAction] = 
        (pref.followsActions[features.previousAction] || 0) + 1;
    }
    
    pref.contexts[features.contextType] = 
      (pref.contexts[features.contextType] || 0) + 1;
  },
  
  /**
   * Detect workflow patterns
   */
  detectWorkflowPatterns: function() {
    const recentActions = this.sessionData.actions.slice(-10);
    
    if (recentActions.length < 3) return;
    
    // Look for repeated sequences
    for (let len = 2; len <= Math.min(5, recentActions.length - 1); len++) {
      const sequence = recentActions.slice(-len).map(a => a.action);
      const sequenceStr = sequence.join(' -> ');
      
      // Check if this sequence has occurred before
      let occurrences = 0;
      for (let i = 0; i <= this.behaviorHistory.length - len; i++) {
        const histSeq = this.behaviorHistory.slice(i, i + len).map(a => a.action).join(' -> ');
        if (histSeq === sequenceStr) {
          occurrences++;
        }
      }
      
      // If sequence is common, add to workflows
      if (occurrences >= 3) {
        const existing = this.preferences.workflowSequences.find(w => w.sequence === sequenceStr);
        if (existing) {
          existing.count = occurrences;
          existing.lastSeen = new Date().toISOString();
        } else {
          this.preferences.workflowSequences.push({
            sequence: sequenceStr,
            actions: sequence,
            count: occurrences,
            confidence: occurrences / this.behaviorHistory.length,
            lastSeen: new Date().toISOString()
          });
        }
      }
    }
    
    // Keep only top workflows
    this.preferences.workflowSequences.sort((a, b) => b.count - a.count);
    this.preferences.workflowSequences = this.preferences.workflowSequences.slice(0, 20);
  },
  
  /**
   * Learn formatting preferences
   */
  learnFormattingPreference: function(formatting) {
    Object.entries(formatting).forEach(([key, value]) => {
      if (!this.preferences.formatting[key]) {
        this.preferences.formatting[key] = {};
      }
      
      this.preferences.formatting[key][value] = 
        (this.preferences.formatting[key][value] || 0) + 1;
    });
  },
  
  /**
   * Learn data pattern preferences
   */
  learnDataPatternPreference: function(dataType, action) {
    if (!this.preferences.dataPatterns[dataType]) {
      this.preferences.dataPatterns[dataType] = {
        preferredActions: {},
        totalOccurrences: 0
      };
    }
    
    const pattern = this.preferences.dataPatterns[dataType];
    pattern.totalOccurrences++;
    pattern.preferredActions[action] = 
      (pattern.preferredActions[action] || 0) + 1;
  },
  
  /**
   * Predict next action based on current context
   */
  predictNextAction: function(currentAction) {
    try {
      if (!this.mlEnabled) return null;
      
      const predictions = [];
      
      // Get action preferences
      const actionPref = this.preferences.features[currentAction];
      if (actionPref && actionPref.followsActions) {
        // Find most likely next actions
        Object.entries(actionPref.followsActions).forEach(([nextAction, count]) => {
          const confidence = count / actionPref.frequency;
          if (confidence >= this.confidenceThreshold) {
            predictions.push({
              action: nextAction,
              confidence: confidence,
              reason: 'Frequently follows current action'
            });
          }
        });
      }
      
      // Check workflow patterns
      const currentSequence = this.sessionData.actions.slice(-4).map(a => a.action).join(' -> ');
      this.preferences.workflowSequences.forEach(workflow => {
        if (workflow.sequence.startsWith(currentSequence) && workflow.confidence >= this.confidenceThreshold) {
          const remainingActions = workflow.sequence.split(' -> ').slice(currentSequence.split(' -> ').length);
          if (remainingActions.length > 0) {
            predictions.push({
              action: remainingActions[0],
              confidence: workflow.confidence,
              reason: 'Part of common workflow',
              workflow: workflow.sequence
            });
          }
        }
      });
      
      // Sort by confidence
      predictions.sort((a, b) => b.confidence - a.confidence);
      
      // Store predictions
      this.nextActionPredictions = {
        timestamp: new Date().toISOString(),
        currentAction: currentAction,
        predictions: predictions.slice(0, 5)
      };
      
      return this.nextActionPredictions;
      
    } catch (error) {
      console.error('Error predicting next action:', error);
      return null;
    }
  },
  
  /**
   * Get feature recommendations based on usage patterns
   */
  getFeatureRecommendations: function() {
    try {
      const recommendations = [];
      const currentHour = new Date().getHours();
      const currentDay = new Date().getDay();
      
      // Analyze feature usage patterns
      Object.entries(this.preferences.features).forEach(([feature, data]) => {
        // Check time-based preferences
        const timeScore = data.timePreferences[currentHour] / Math.max(1, data.frequency);
        const dayScore = data.dayPreferences[currentDay] / Math.max(1, data.frequency);
        
        // Calculate overall recommendation score
        const score = (timeScore * 0.3 + dayScore * 0.2 + (data.frequency / 100) * 0.5);
        
        if (score >= 0.1) {
          recommendations.push({
            feature: feature,
            score: score,
            lastUsed: this.featureUsage[feature]?.lastUsed,
            usageCount: data.frequency,
            reason: this.getRecommendationReason(feature, timeScore, dayScore)
          });
        }
      });
      
      // Add rarely used but potentially useful features
      const allFeatures = ['cleanData', 'removeDuplicates', 'validateData', 'formatData', 'analyzeData'];
      allFeatures.forEach(feature => {
        if (!this.featureUsage[feature] || this.featureUsage[feature].count < 3) {
          recommendations.push({
            feature: feature,
            score: 0.5,
            reason: 'Useful feature you haven\'t tried much',
            isNew: true
          });
        }
      });
      
      // Sort by score
      recommendations.sort((a, b) => b.score - a.score);
      
      this.featureRecommendations = recommendations.slice(0, 10);
      return this.featureRecommendations;
      
    } catch (error) {
      console.error('Error getting recommendations:', error);
      return [];
    }
  },
  
  /**
   * Get recommendation reason
   */
  getRecommendationReason: function(feature, timeScore, dayScore) {
    const reasons = [];
    
    if (timeScore > 0.3) {
      reasons.push('Often used at this time');
    }
    
    if (dayScore > 0.3) {
      reasons.push('Frequently used on this day');
    }
    
    const usage = this.featureUsage[feature];
    if (usage && usage.count > 10) {
      reasons.push('One of your most used features');
    }
    
    if (usage && usage.lastUsed) {
      const daysSinceUse = Math.floor((Date.now() - new Date(usage.lastUsed)) / (1000 * 60 * 60 * 24));
      if (daysSinceUse > 7) {
        reasons.push('Haven\'t used in a while');
      }
    }
    
    return reasons.join('; ') || 'Recommended based on your patterns';
  },
  
  /**
   * Get preferred settings for a feature
   */
  getPreferredSettings: function(feature) {
    const usage = this.featureUsage[feature];
    if (!usage || !usage.preferences) return {};
    
    const preferred = {};
    
    Object.entries(usage.preferences).forEach(([setting, values]) => {
      // Find most common value
      let maxCount = 0;
      let preferredValue = null;
      
      Object.entries(values).forEach(([value, count]) => {
        if (count > maxCount) {
          maxCount = count;
          preferredValue = value;
        }
      });
      
      if (preferredValue !== null && maxCount >= this.minSampleSize) {
        preferred[setting] = preferredValue;
      }
    });
    
    return preferred;
  },
  
  /**
   * Get workflow suggestions
   */
  getWorkflowSuggestions: function() {
    const currentActions = this.sessionData.actions.slice(-3).map(a => a.action);
    const suggestions = [];
    
    if (currentActions.length === 0) return suggestions;
    
    // Find matching workflows
    this.preferences.workflowSequences.forEach(workflow => {
      const workflowActions = workflow.actions;
      
      // Check if current actions match the beginning of a workflow
      let matches = true;
      for (let i = 0; i < currentActions.length && i < workflowActions.length; i++) {
        if (currentActions[i] !== workflowActions[i]) {
          matches = false;
          break;
        }
      }
      
      if (matches && workflowActions.length > currentActions.length) {
        suggestions.push({
          workflow: workflow.sequence,
          nextSteps: workflowActions.slice(currentActions.length),
          confidence: workflow.confidence,
          completionRate: (currentActions.length / workflowActions.length) * 100
        });
      }
    });
    
    // Sort by confidence
    suggestions.sort((a, b) => b.confidence - a.confidence);
    
    this.workflowSuggestions = suggestions.slice(0, 3);
    return this.workflowSuggestions;
  },
  
  /**
   * Learn from error
   */
  learnFromError: function(error, action, context) {
    try {
      this.sessionData.errors.push({
        timestamp: new Date().toISOString(),
        error: error.message,
        action: action,
        context: context
      });
      
      // Reduce confidence in this action for this context
      const contextKey = JSON.stringify(context);
      if (!this.preferences.automation[action]) {
        this.preferences.automation[action] = {
          contexts: {}
        };
      }
      
      if (!this.preferences.automation[action].contexts[contextKey]) {
        this.preferences.automation[action].contexts[contextKey] = {
          successRate: 100,
          attempts: 0
        };
      }
      
      const automation = this.preferences.automation[action].contexts[contextKey];
      automation.attempts++;
      automation.successRate = Math.max(0, automation.successRate - 10);
      
    } catch (err) {
      console.error('Error learning from error:', err);
    }
  },
  
  /**
   * Learn from successful completion
   */
  learnFromSuccess: function(action, context, duration) {
    try {
      this.sessionData.completions.push({
        timestamp: new Date().toISOString(),
        action: action,
        context: context,
        duration: duration
      });
      
      // Increase confidence in this action for this context
      const contextKey = JSON.stringify(context);
      if (!this.preferences.automation[action]) {
        this.preferences.automation[action] = {
          contexts: {}
        };
      }
      
      if (!this.preferences.automation[action].contexts[contextKey]) {
        this.preferences.automation[action].contexts[contextKey] = {
          successRate: 100,
          attempts: 0,
          avgDuration: duration
        };
      }
      
      const automation = this.preferences.automation[action].contexts[contextKey];
      automation.attempts++;
      automation.successRate = Math.min(100, automation.successRate + 5);
      automation.avgDuration = (automation.avgDuration * (automation.attempts - 1) + duration) / automation.attempts;
      
    } catch (error) {
      console.error('Error learning from success:', error);
    }
  },
  
  /**
   * Should automate action based on preferences
   */
  shouldAutomate: function(action, context) {
    if (!this.mlEnabled) return false;
    
    const contextKey = JSON.stringify(context);
    const automation = this.preferences.automation[action]?.contexts[contextKey];
    
    if (!automation) return false;
    
    // Automate if high success rate and enough attempts
    return automation.successRate >= 90 && automation.attempts >= this.minSampleSize;
  },
  
  /**
   * Get user preference summary
   */
  getPreferenceSummary: function() {
    const summary = {
      mostUsedFeatures: [],
      preferredWorkflows: [],
      formattingPreferences: {},
      dataPatternPreferences: {},
      automationCandidates: [],
      timePatterns: {},
      totalActions: this.behaviorHistory.length
    };
    
    // Most used features
    const sortedFeatures = Object.entries(this.featureUsage)
      .sort((a, b) => b[1].count - a[1].count)
      .slice(0, 5);
    
    summary.mostUsedFeatures = sortedFeatures.map(([feature, data]) => ({
      feature: feature,
      count: data.count,
      lastUsed: data.lastUsed
    }));
    
    // Preferred workflows
    summary.preferredWorkflows = this.preferences.workflowSequences.slice(0, 3);
    
    // Formatting preferences
    Object.entries(this.preferences.formatting).forEach(([key, values]) => {
      let maxCount = 0;
      let preferred = null;
      
      Object.entries(values).forEach(([value, count]) => {
        if (count > maxCount) {
          maxCount = count;
          preferred = value;
        }
      });
      
      if (preferred) {
        summary.formattingPreferences[key] = preferred;
      }
    });
    
    // Data pattern preferences
    Object.entries(this.preferences.dataPatterns).forEach(([dataType, pattern]) => {
      const sortedActions = Object.entries(pattern.preferredActions)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 3);
      
      summary.dataPatternPreferences[dataType] = sortedActions.map(([action, count]) => ({
        action: action,
        frequency: (count / pattern.totalOccurrences * 100).toFixed(1) + '%'
      }));
    });
    
    // Automation candidates
    Object.entries(this.preferences.automation).forEach(([action, data]) => {
      Object.entries(data.contexts).forEach(([context, stats]) => {
        if (stats.successRate >= 90 && stats.attempts >= this.minSampleSize) {
          summary.automationCandidates.push({
            action: action,
            successRate: stats.successRate,
            attempts: stats.attempts
          });
        }
      });
    });
    
    // Time patterns
    const hourCounts = Array(24).fill(0);
    this.behaviorHistory.forEach(action => {
      const hour = new Date(action.timestamp).getHours();
      hourCounts[hour]++;
    });
    
    const maxHour = hourCounts.indexOf(Math.max(...hourCounts));
    summary.timePatterns = {
      mostActiveHour: maxHour,
      activityDistribution: hourCounts
    };
    
    return summary;
  },
  
  /**
   * Save preferences
   */
  savePreferences: function() {
    try {
      PropertiesService.getUserProperties().setProperty(
        'user_preferences',
        JSON.stringify(this.preferences)
      );
      
      PropertiesService.getUserProperties().setProperty(
        'behavior_history',
        JSON.stringify(this.behaviorHistory.slice(-5000))
      );
      
      PropertiesService.getUserProperties().setProperty(
        'feature_usage',
        JSON.stringify(this.featureUsage)
      );
      
      return true;
    } catch (error) {
      console.error('Error saving preferences:', error);
      return false;
    }
  },
  
  /**
   * Export preferences for analysis
   */
  exportPreferences: function() {
    return {
      preferences: this.preferences,
      featureUsage: this.featureUsage,
      behaviorHistory: this.behaviorHistory.slice(-100),
      summary: this.getPreferenceSummary(),
      exportDate: new Date().toISOString()
    };
  },
  
  /**
   * Reset preferences
   */
  resetPreferences: function() {
    this.preferences = {
      features: {},
      formatting: {},
      automation: {},
      dataPatterns: {},
      workflowSequences: []
    };
    
    this.behaviorHistory = [];
    this.featureUsage = {};
    this.sessionData = {};
    
    this.savePreferences();
    this.startSession();
    
    return true;
  }
};