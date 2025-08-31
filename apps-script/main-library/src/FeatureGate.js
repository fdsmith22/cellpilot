/**
 * Feature Gate System for CellPilot
 * Controls feature access based on user tier and beta status
 */

const FeatureGate = {
  
  /**
   * Beta configuration - Easy to toggle on/off
   */
  BETA_CONFIG: {
    enabled: true, // Set to false when beta ends
    endDate: '2025-10-01',
    unlockAllFeatures: true,
    showBetaBadge: true,
    collectFeedback: true,
    betaUserBenefits: {
      extendedTrial: 60, // days after beta ends
      discountPercentage: 50, // lifetime discount for beta users
      grandFatherFeatures: ['formula_builder', 'automation'] // Keep these free for beta users
    }
  },
  
  /**
   * Check if we're currently in beta period
   */
  isBetaPeriod: function() {
    if (!this.BETA_CONFIG.enabled) return false;
    
    const betaEndDate = new Date(this.BETA_CONFIG.endDate);
    const now = new Date();
    return now < betaEndDate;
  },
  
  /**
   * Check if user is a beta user (for grandfather benefits)
   */
  isBetaUser: function() {
    const betaJoinDate = UserSettings.load('betaJoinDate');
    if (!betaJoinDate) return false;
    
    const joinDate = new Date(betaJoinDate);
    const betaEndDate = new Date(this.BETA_CONFIG.endDate);
    return joinDate < betaEndDate;
  },
  
  /**
   * Main access control function
   */
  canAccess: function(featureName) {
    // During beta with unlockAllFeatures, grant access to everything
    if (this.isBetaPeriod() && this.BETA_CONFIG.unlockAllFeatures) {
      // Track that user joined during beta
      if (!UserSettings.load('betaJoinDate')) {
        UserSettings.save('betaJoinDate', new Date().toISOString());
      }
      
      return {
        allowed: true,
        reason: 'beta_unlimited_access',
        message: 'All features unlocked during beta!'
      };
    }
    
    // Check if beta user has grandfather access
    if (this.isBetaUser()) {
      const grandFathered = this.BETA_CONFIG.betaUserBenefits.grandFatherFeatures;
      if (grandFathered.includes(featureName)) {
        return {
          allowed: true,
          reason: 'beta_user_benefit',
          message: 'Feature permanently unlocked as beta user benefit!'
        };
      }
      
      // Check if still in extended trial period
      const betaEndDate = new Date(this.BETA_CONFIG.endDate);
      const extendedEndDate = new Date(betaEndDate);
      extendedEndDate.setDate(extendedEndDate.getDate() + this.BETA_CONFIG.betaUserBenefits.extendedTrial);
      
      if (new Date() < extendedEndDate) {
        return {
          allowed: true,
          reason: 'beta_extended_trial',
          message: `Extended trial until ${extendedEndDate.toLocaleDateString()}`
        };
      }
    }
    
    // Normal tier-based access control
    const userTier = UserSettings.load('userTier', 'free');
    const featureAccess = Config.FEATURE_ACCESS[featureName];
    
    // If feature not defined, allow access (fail open for undefined features)
    if (!featureAccess) {
      return {
        allowed: true,
        reason: 'feature_not_gated',
        message: 'Feature available to all users'
      };
    }
    
    // Check if user's tier has access
    const allowed = featureAccess.includes(userTier);
    
    if (allowed) {
      return {
        allowed: true,
        reason: 'tier_access',
        message: `Feature available in ${userTier} tier`
      };
    }
    
    // Access denied - need upgrade
    return {
      allowed: false,
      reason: 'upgrade_required',
      requiredTiers: featureAccess,
      message: `Upgrade to ${featureAccess.join(' or ')} to unlock this feature`,
      showUpgrade: true
    };
  },
  
  /**
   * Check if a feature should show upgrade prompt
   */
  shouldShowUpgrade: function(featureName) {
    const access = this.canAccess(featureName);
    return !access.allowed && access.showUpgrade;
  },
  
  /**
   * Get upgrade prompt configuration
   */
  getUpgradePrompt: function(featureName) {
    const access = this.canAccess(featureName);
    
    if (!access.allowed && access.reason === 'upgrade_required') {
      // Special messaging for beta users
      const isBetaUser = this.isBetaUser();
      const discount = isBetaUser ? this.BETA_CONFIG.betaUserBenefits.discountPercentage : 0;
      
      return {
        title: 'Premium Feature',
        subtitle: access.message,
        description: this.getFeatureDescription(featureName),
        benefits: this.getTierBenefits(access.requiredTiers[0]),
        cta: {
          text: isBetaUser ? 'Upgrade Now (' + discount + '% Beta Discount)' : 'Upgrade Now',
          action: 'showUpgradeDialog',
          params: {
            feature: featureName,
            requiredTiers: access.requiredTiers,
            discount: discount
          }
        },
        alternativeCta: {
          text: 'Learn More',
          action: 'showFeatureInfo',
          params: { feature: featureName }
        }
      };
    }
    
    return null;
  },
  
  /**
   * Get feature description for upgrade prompts
   */
  getFeatureDescription: function(featureName) {
    const descriptions = {
      'formula_builder': 'Build complex formulas with natural language descriptions',
      'automation': 'Automate repetitive tasks and schedule operations',
      'industry_tools': 'Access professional templates and industry-specific tools',
      'team_features': 'Collaborate with your team and share configurations',
      'advanced_cleaning': 'Advanced data cleaning with ML-powered detection',
      'api_integration': 'Connect to external services and APIs',
      'unlimited_operations': 'Remove all usage limits and restrictions'
    };
    
    return descriptions[featureName] || 'Unlock advanced functionality';
  },
  
  /**
   * Get tier benefits for upgrade prompts
   */
  getTierBenefits: function(tierName) {
    const benefits = {
      'starter': [
        '500 operations per month',
        'Basic automation features',
        'Email support',
        'Formula builder access'
      ],
      'professional': [
        'Unlimited operations',
        'All automation features',
        'Industry templates',
        'Priority support',
        'Advanced ML features'
      ],
      'business': [
        'Everything in Professional',
        'Team collaboration',
        'API access',
        'Custom integrations',
        'Dedicated support'
      ]
    };
    
    return benefits[tierName] || [];
  },
  
  /**
   * Track feature access attempts for analytics
   */
  trackAccess: function(featureName, allowed) {
    try {
      const eventData = {
        feature: featureName,
        allowed: allowed,
        userTier: UserSettings.load('userTier', 'free'),
        isBetaUser: this.isBetaUser(),
        timestamp: new Date().toISOString()
      };
      
      // Send to analytics
      if (typeof ApiIntegration !== 'undefined' && ApiIntegration.trackEvent) {
        ApiIntegration.trackEvent('feature_access', eventData);
      }
      
      // Store locally for usage patterns
      const accessLog = UserSettings.load('featureAccessLog', []);
      accessLog.push(eventData);
      
      // Keep only last 100 entries
      if (accessLog.length > 100) {
        accessLog.shift();
      }
      
      UserSettings.save('featureAccessLog', accessLog);
    } catch (error) {
      // Fail silently - don't break feature access for analytics
    }
  },
  
  /**
   * Check and enforce feature access with UI feedback
   */
  enforceAccess: function(featureName, callback) {
    const access = this.canAccess(featureName);
    
    // Track the access attempt
    this.trackAccess(featureName, access.allowed);
    
    if (access.allowed) {
      // Execute the callback for allowed access
      if (callback && typeof callback === 'function') {
        callback();
      }
      return true;
    }
    
    // Show upgrade prompt for denied access
    const upgradePrompt = this.getUpgradePrompt(featureName);
    if (upgradePrompt && typeof UIComponents !== 'undefined') {
      UIComponents.showUpgradeDialog(upgradePrompt);
    } else {
      // Fallback alert if UI components not available
      const message = `This feature requires a ${access.requiredTiers.join(' or ')} subscription.\n\nUpgrade at cellpilot.io/pricing`;
      SpreadsheetApp.getUi().alert('Premium Feature', message, SpreadsheetApp.getUi().ButtonSet.OK);
    }
    
    return false;
  },
  
  /**
   * Get beta status information
   */
  getBetaStatus: function() {
    if (!this.isBetaPeriod()) {
      return {
        active: false,
        message: 'Beta period has ended'
      };
    }
    
    const endDate = new Date(this.BETA_CONFIG.endDate);
    const now = new Date();
    const daysRemaining = Math.ceil((endDate - now) / (1000 * 60 * 60 * 24));
    
    // Format date as DD/MM/YYYY
    const day = String(endDate.getDate()).padStart(2, '0');
    const month = String(endDate.getMonth() + 1).padStart(2, '0');
    const year = endDate.getFullYear();
    const formattedDate = day + '/' + month + '/' + year;
    
    return {
      active: true,
      daysRemaining: daysRemaining,
      endDate: formattedDate,
      message: 'Beta period: ' + daysRemaining + ' days remaining. All features currently unlocked.',
      benefits: this.BETA_CONFIG.betaUserBenefits
    };
  },
  
  /**
   * Initialize feature gate system
   */
  initialize: function() {
    // Check and set beta status
    if (this.isBetaPeriod() && !UserSettings.load('betaNotificationShown')) {
      // Show beta welcome message
      const betaStatus = this.getBetaStatus();
      const ui = SpreadsheetApp.getUi();
      ui.alert(
        'Welcome to CellPilot Beta',
        'Thank you for joining our beta program.\n\n' +
        'Beta Program Benefits:\n' +
        '- Full access to all premium features\n' +
        '- Beta period ends: ' + betaStatus.endDate + '\n\n' +
        'Exclusive Beta User Rewards:\n' +
        '- ' + this.BETA_CONFIG.betaUserBenefits.extendedTrial + '-day extended trial after beta\n' +
        '- ' + this.BETA_CONFIG.betaUserBenefits.discountPercentage + '% lifetime discount on all plans\n' +
        '- Permanent access to select premium features\n\n' +
        'Important Notice:\n' +
        'CellPilot is actively being developed and improved. While we strive for excellence,\n' +
        'some features may still be under refinement. Please report any issues you encounter.\n\n' +
        'Your feedback helps us improve CellPilot. Thank you for your support.',
        ui.ButtonSet.OK
      );
      
      UserSettings.save('betaNotificationShown', true);
      UserSettings.save('betaJoinDate', new Date().toISOString());
    }
  }
};

// Export for use in other modules
if (typeof module !== 'undefined' && module.exports) {
  module.exports = FeatureGate;
}