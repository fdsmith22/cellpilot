/**
 * CellPilot UI Components
 * Handles all user interface elements and card building
 */

const UIComponents = {
  
  /**
   * Build the data cleaning tools section
   */
  buildDataCleaningSection: function() {
    const section = CardService.newCardSection()
      .setHeader('Data Cleaning')
      .setCollapsible(true);
    
    // Remove Duplicates
    const duplicatesButton = CardService.newTextButton()
      .setText('Remove Duplicates')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('showDuplicateRemovalOptions'));
    
    // Standardize Text
    const textButton = CardService.newTextButton()
      .setText('Standardize Text')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('showTextStandardizationOptions'));
    
    // Fix Dates
    const dateButton = CardService.newTextButton()
      .setText('Fix Date Formats')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('showDateFormattingOptions'));
    
    section.addWidget(duplicatesButton);
    section.addWidget(textButton);
    section.addWidget(dateButton);
    
    return section;
  },
  
  /**
   * Build the formula builder section
   */
  buildFormulaSection: function() {
    const userTier = UserSettings.load('userTier', 'free');
    const hasAccess = Config.FEATURE_ACCESS.formula_builder.includes(userTier);
    
    const section = CardService.newCardSection()
      .setHeader('Formula Builder')
      .setCollapsible(true);
    
    if (!hasAccess) {
      section.addWidget(CardService.newTextParagraph()
        .setText('<i>Upgrade to Starter plan to access formula builder</i>'));
      
      const upgradeButton = CardService.newTextButton()
        .setText('Upgrade Plan')
        .setOnClickAction(CardService.newAction()
          .setFunctionName('showUpgradeOptions'));
      
      section.addWidget(upgradeButton);
      return section;
    }
    
    // Natural Language Formula
    const nlFormulaButton = CardService.newTextButton()
      .setText('Build Formula from Description')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('showNaturalLanguageFormulaBuilder'));
    
    // Formula Templates
    const templatesButton = CardService.newTextButton()
      .setText('Formula Templates')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('showFormulaTemplates'));
    
    section.addWidget(nlFormulaButton);
    section.addWidget(templatesButton);
    
    return section;
  },
  
  /**
   * Build the automation section
   */
  buildAutomationSection: function() {
    const userTier = UserSettings.load('userTier', 'free');
    const hasAccess = Config.FEATURE_ACCESS.automation.includes(userTier);
    
    const section = CardService.newCardSection()
      .setHeader('Automation')
      .setCollapsible(true);
    
    if (!hasAccess) {
      section.addWidget(CardService.newTextParagraph()
        .setText('<i>Upgrade to Professional plan to access automation</i>'));
      
      const upgradeButton = CardService.newTextButton()
        .setText('Upgrade Plan')
        .setOnClickAction(CardService.newAction()
          .setFunctionName('showUpgradeOptions'));
      
      section.addWidget(upgradeButton);
      return section;
    }
    
    // Email Automation
    const emailButton = CardService.newTextButton()
      .setText('Set Up Email Alerts')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('showEmailAutomationSetup'));
    
    // Calendar Integration
    const calendarButton = CardService.newTextButton()
      .setText('Create Calendar Events')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('showCalendarIntegration'));
    
    section.addWidget(emailButton);
    section.addWidget(calendarButton);
    
    return section;
  },
  
  /**
   * Build duplicate removal options interface
   */
  buildDuplicateRemovalCard: function(previewData = null) {
    const card = CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader()
        .setTitle('Remove Duplicates')
        .setSubtitle('Configure duplicate detection settings'));
    
    // Settings section
    const settingsSection = CardService.newCardSection()
      .setHeader('Detection Settings');
    
    // Similarity threshold
    const thresholdSlider = CardService.newSlider()
      .setFieldName('similarityThreshold')
      .setTitle('Similarity Threshold')
      .setMin(0.5)
      .setMax(1.0)
      .setValue(UserSettings.load('duplicateThreshold', 0.85));
    
    // Case sensitive toggle
    const caseSensitiveSwitch = CardService.newSwitch()
      .setFieldName('caseSensitive')
      .setValue(UserSettings.load('caseSensitiveDuplicates', false))
      .setControlType(CardService.SwitchControlType.CHECKBOX);
    
    settingsSection.addWidget(thresholdSlider);
    settingsSection.addWidget(CardService.newDecoratedText()
      .setText('Case Sensitive')
      .setSwitchControl(caseSensitiveSwitch));
    
    card.addSection(settingsSection);
    
    // Preview section if data provided
    if (previewData) {
      const previewSection = this.buildPreviewSection(previewData);
      card.addSection(previewSection);
    }
    
    // Action buttons
    const actionSection = CardService.newCardSection();
    
    const previewButton = CardService.newTextButton()
      .setText('Preview Changes')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('previewDuplicateRemoval'));
    
    const applyButton = CardService.newTextButton()
      .setText('Apply Changes')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('applyDuplicateRemoval'));
    
    actionSection.addWidget(previewButton);
    actionSection.addWidget(applyButton);
    
    card.addSection(actionSection);
    
    return card.build();
  },
  
  /**
   * Build text standardization interface
   */
  buildTextStandardizationCard: function() {
    const card = CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader()
        .setTitle('Standardize Text')
        .setSubtitle('Clean and format text data'));
    
    const optionsSection = CardService.newCardSection()
      .setHeader('Standardization Options');
    
    // Case options
    const caseSelection = CardService.newSelectionInput()
      .setType(CardService.SelectionInputType.DROPDOWN)
      .setTitle('Text Case')
      .setFieldName('textCase')
      .addItem('Keep Original', 'original', true)
      .addItem('Title Case', 'title', false)
      .addItem('UPPER CASE', 'upper', false)
      .addItem('lower case', 'lower', false)
      .addItem('Sentence case', 'sentence', false);
    
    // Cleanup options
    const trimSwitch = CardService.newSwitch()
      .setFieldName('trimSpaces')
      .setValue(true)
      .setControlType(CardService.SwitchControlType.CHECKBOX);
    
    const extraSpacesSwitch = CardService.newSwitch()
      .setFieldName('removeExtraSpaces')
      .setValue(true)
      .setControlType(CardService.SwitchControlType.CHECKBOX);
    
    const specialCharsSwitch = CardService.newSwitch()
      .setFieldName('removeSpecialChars')
      .setValue(false)
      .setControlType(CardService.SwitchControlType.CHECKBOX);
    
    optionsSection.addWidget(caseSelection);
    optionsSection.addWidget(CardService.newDecoratedText()
      .setText('Trim whitespace')
      .setSwitchControl(trimSwitch));
    optionsSection.addWidget(CardService.newDecoratedText()
      .setText('Remove extra spaces')
      .setSwitchControl(extraSpacesSwitch));
    optionsSection.addWidget(CardService.newDecoratedText()
      .setText('Remove special characters')
      .setSwitchControl(specialCharsSwitch));
    
    card.addSection(optionsSection);
    
    // Action buttons
    const actionSection = CardService.newCardSection();
    
    const previewButton = CardService.newTextButton()
      .setText('Preview Changes')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('previewTextStandardization'));
    
    const applyButton = CardService.newTextButton()
      .setText('Apply Changes')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('applyTextStandardization'));
    
    actionSection.addWidget(previewButton);
    actionSection.addWidget(applyButton);
    
    card.addSection(actionSection);
    
    return card.build();
  },
  
  /**
   * Build preview section for changes
   */
  buildPreviewSection: function(previewData) {
    const section = CardService.newCardSection()
      .setHeader('Preview Changes');
    
    if (!previewData || previewData.length === 0) {
      section.addWidget(CardService.newTextParagraph()
        .setText('<i>No changes to preview</i>'));
      return section;
    }
    
    // Show sample of changes
    const maxPreviewRows = 5;
    const sampleData = previewData.slice(0, maxPreviewRows);
    
    let previewText = '<font color="#d93025"><b>Before → After:</b></font><br>';
    
    sampleData.forEach((item, index) => {
      const before = String(item.before).substring(0, 30);
      const after = String(item.after).substring(0, 30);
      previewText += `${index + 1}. ${before} → <font color="#137333">${after}</font><br>`;
    });
    
    if (previewData.length > maxPreviewRows) {
      previewText += `<i>... and ${previewData.length - maxPreviewRows} more changes</i>`;
    }
    
    section.addWidget(CardService.newTextParagraph()
      .setText(previewText));
    
    return section;
  },
  
  /**
   * Build success notification card
   */
  buildSuccessCard: function(message, details = null) {
    const card = CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader()
        .setTitle('Success')
        .setImageUrl('https://fonts.gstatic.com/s/i/materialicons/check_circle/v6/24px.svg'));
    
    const section = CardService.newCardSection()
      .addWidget(CardService.newTextParagraph()
        .setText(`<font color="#137333"><b>${message}</b></font>`));
    
    if (details) {
      section.addWidget(CardService.newTextParagraph()
        .setText(details));
    }
    
    // Return to main button
    const returnButton = CardService.newTextButton()
      .setText('Back to CellPilot')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('buildHomepage'));
    
    section.addWidget(returnButton);
    card.addSection(section);
    
    return card.build();
  },
  
  /**
   * Build upgrade/pricing card
   */
  buildUpgradeCard: function() {
    const card = CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader()
        .setTitle('Upgrade CellPilot')
        .setSubtitle('Unlock advanced features'));
    
    // Current tier info
    const currentTier = UserSettings.load('userTier', 'free');
    const currentSection = CardService.newCardSection()
      .setHeader('Current Plan')
      .addWidget(CardService.newTextParagraph()
        .setText(`You are on the <b>${currentTier.toUpperCase()}</b> plan`));
    
    card.addSection(currentSection);
    
    // Pricing tiers
    const pricingSection = CardService.newCardSection()
      .setHeader('Available Plans');
    
    const plans = [
      { name: 'Starter', price: '£5.99/month', features: ['500 operations/month', 'Email automation', 'Formula builder'] },
      { name: 'Professional', price: '£11.99/month', features: ['Unlimited operations', 'Industry tools', 'Calendar integration'] },
      { name: 'Business', price: '£19.99/month', features: ['Team features', 'Priority support', 'Custom integrations'] }
    ];
    
    plans.forEach(plan => {
      let planText = `<b>${plan.name}</b> - ${plan.price}<br>`;
      plan.features.forEach(feature => planText += `• ${feature}<br>`);
      
      pricingSection.addWidget(CardService.newTextParagraph()
        .setText(planText));
    });
    
    card.addSection(pricingSection);
    
    // Contact section
    const contactSection = CardService.newCardSection()
      .addWidget(CardService.newTextParagraph()
        .setText('Visit <b>cellpilot.co.uk</b> to upgrade your plan'));
    
    card.addSection(contactSection);
    
    return card.build();
  },
  
  /**
   * Build loading/progress indicator
   */
  buildLoadingCard: function(message = 'Processing...') {
    return CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader()
        .setTitle('CellPilot')
        .setImageUrl('https://developers.google.com/apps-script/images/logo.png'))
      .addSection(CardService.newCardSection()
        .addWidget(CardService.newTextParagraph()
          .setText(`<i>${message}</i>`)))
      .build();
  }
};