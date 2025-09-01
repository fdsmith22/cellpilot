/**
 * CellPilot - Hybrid Google Sheets Add-on
 * Supports both CardService (add-on) and Menu+Sidebar (independent) modes
 */

/**
 * Standard add-on entry points
 */
function onInstall(e) {
  onOpen(e);
}

function onOpen(e) {
  Logger.info('CellPilot initializing...');
  UserSettings.initialize();
  
  // Initialize Feature Gate system
  if (typeof FeatureGate !== 'undefined') {
    FeatureGate.initialize();
  }
  
  // Create fallback menu for development/independent installation
  createCellPilotMenu();
  
  // Track installation if first time
  if (!UserSettings.load('hasTrackedInstall', false)) {
    ApiIntegration.trackInstallation();
    UserSettings.save('hasTrackedInstall', true);
  }
  
  // Check subscription status (async)
  try {
    const subscription = ApiIntegration.checkSubscription();
    if (subscription && subscription.plan) {
      UserSettings.save('userTier', subscription.plan);
    }
  } catch (error) {
    Logger.info('Could not check subscription:', error);
  }
  
  Logger.info('CellPilot ready');
}

/**
 * Reset beta notification for testing
 * This allows the beta welcome message to show again
 */
function resetBetaNotification() {
  UserSettings.remove('betaNotificationShown');
  UserSettings.remove('betaJoinDate');
  SpreadsheetApp.getUi().alert('Beta notification reset. Reload the spreadsheet to see the welcome message.');
}

/**
 * CardService homepage for Google Workspace Add-on
 * Called automatically when add-on is opened from sidebar
 */
function buildHomepage(e) {
  try {
    Logger.info('Building add-on homepage...');
    
    // Create an improved card interface that matches our HTML sidebar styling
    const card = CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader()
        .setTitle('CellPilot')
        .setSubtitle('Smart Spreadsheet Assistant')
        .setImageStyle(CardService.ImageStyle.SQUARE)
        .setImageUrl('https://fonts.gstatic.com/s/i/materialicons/table_chart/v6/24px.svg'))
      .addSection(CardService.newCardSection()
        .addWidget(CardService.newTextButton()
          .setText('Launch CellPilot Dashboard')
          .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
          .setBackgroundColor('#6366f1')
          .setOnClickAction(CardService.newAction()
            .setFunctionName('launchHtmlSidebarFromAddon')))
        .addWidget(CardService.newDivider()))
      .addSection(CardService.newCardSection()
        .setHeader('Quick Actions')
        .addWidget(CardService.newDecoratedText()
          .setTopLabel('Data Cleaning')
          .setText('Tableize Data')
          .setBottomLabel('Structure messy data into columns')
          .setOnClickAction(CardService.newAction()
            .setFunctionName('launchTableizeFromAddon')))
        .addWidget(CardService.newDecoratedText()
          .setTopLabel('Data Cleaning')
          .setText('Clean & Standardize')
          .setBottomLabel('Remove duplicates, fix text, format dates')
          .setOnClickAction(CardService.newAction()
            .setFunctionName('launchDataCleaningFromAddon')))
        .addWidget(CardService.newDecoratedText()
          .setTopLabel('Formulas')
          .setText('Formula Builder')
          .setBottomLabel('Generate formulas from descriptions')
          .setOnClickAction(CardService.newAction()
            .setFunctionName('launchFormulasFromAddon'))))
      .addSection(CardService.newCardSection()
        .addWidget(CardService.newTextParagraph()
          .setText('For full features, click "Launch CellPilot Dashboard" above')))
      .build();
    
    return card;
    
  } catch (error) {
    Logger.error('Error building homepage:', error);
    return buildErrorCard('Failed to load CellPilot', error.message);
  }
}

/**
 * Launch HTML sidebar from add-on card
 */
function launchHtmlSidebarFromAddon() {
  showCellPilotSidebar();
  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification()
      .setText('Opening CellPilot...'))
    .build();
}

function launchTableizeFromAddon() {
  showTableize();
  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification()
      .setText('Opening Tableize...'))
    .build();
}

function launchDataCleaningFromAddon() {
  showDataCleaning();
  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification()
      .setText('Opening Data Cleaning...'))
    .build();
}

// Legacy support
function launchDuplicatesFromAddon() {
  return launchDataCleaningFromAddon();
}

function launchFormulasFromAddon() {
  showFormulaBuilder();
  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification()
      .setText('Opening Formula Builder...'))
    .build();
}

/**
 * Create menu for independent/development installation
 */
function createCellPilotMenu() {
  try {
    const ui = SpreadsheetApp.getUi();
    // When running as a library, we need to use the library prefix
    const prefix = typeof CellPilot !== 'undefined' ? 'CellPilot.' : '';
    
    ui.createMenu('CellPilot')
      .addItem('Open CellPilot', prefix + 'showCellPilotSidebar')
      .addSeparator()
      .addItem('Data Cleaning', prefix + 'showDataCleaning')
      .addItem('Tableize Data', prefix + 'showTableize')
      .addItem('Advanced Data Restructuring', prefix + 'showAdvancedRestructuring')
      .addSubMenu(ui.createMenu('Formula Builder')
        .addItem('Smart Formula Assistant', prefix + 'showSmartFormulaAssistant')
        .addItem('Natural Language Builder', prefix + 'showFormulaBuilder')
        .addItem('Cross-Sheet Formula Builder', prefix + 'showCrossSheetFormulaBuilder')
        .addItem('Formula Performance Optimizer', prefix + 'showFormulaPerformanceOptimizer')
        .addItem('Formula Templates', prefix + 'showFormulaTemplates'))
      .addSubMenu(ui.createMenu('Advanced Tools')
        .addItem('Pivot Table Assistant', prefix + 'showPivotTableAssistant')
        .addItem('Data Pipeline Manager', prefix + 'showDataPipelineManager')
        .addItem('Data Validation Generator', prefix + 'showDataValidationGenerator')
        .addItem('Conditional Formatting Wizard', prefix + 'showConditionalFormattingWizard')
        .addItem('Multi-Tab Relationship Mapper', prefix + 'showMultiTabRelationshipMapper')
        .addItem('Smart Formula Debugger', prefix + 'showSmartFormulaDebugger'))
      .addSeparator()
      .addItem('Settings', prefix + 'showSettings')
      .addItem('Help', prefix + 'showHelp')
      .addSeparator()
      .addItem('Reset Beta Welcome (Testing)', prefix + 'resetBetaNotification')
      .addToUi();
  } catch (error) {
    Logger.error('Error creating menu:', error);
  }
}

/**
 * Build main CardService interface
 */
function buildMainCard(context) {
  const card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
      .setTitle('CellPilot')
      .setSubtitle('Spreadsheet Automation Made Simple')
      .setImageUrl('https://fonts.gstatic.com/s/i/materialicons/table_chart/v6/24px.svg'));

  // Context info section
  if (context && context.hasSelection) {
    const contextSection = CardService.newCardSection()
      .setHeader('Current Selection')
      .addWidget(CardService.newTextParagraph()
        .setText(`${context.rowCount}×${context.colCount} ${context.dataType}`));
    card.addSection(contextSection);
  } else {
    const infoSection = CardService.newCardSection()
      .addWidget(CardService.newTextParagraph()
        .setText('Select data in your sheet to unlock context-aware tools'));
    card.addSection(infoSection);
  }

  // Data Cleaning section
  const cleaningSection = CardService.newCardSection()
    .setHeader('Data Cleaning')
    .addWidget(CardService.newTextButton()
      .setText('Remove Duplicates')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('showDuplicateRemovalCard')))
    .addWidget(CardService.newTextButton()
      .setText('Standardize Text')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('showTextStandardizationCard')))
    .addWidget(CardService.newTextButton()
      .setText('Fix Date Formats')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('showDateFormattingCard')));

  card.addSection(cleaningSection);

  // Formula Builder section
  const formulaSection = CardService.newCardSection()
    .setHeader('Formula Builder')
    .addWidget(CardService.newTextButton()
      .setText('Natural Language Builder')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('showFormulaBuilderCard')))
    .addWidget(CardService.newTextButton()
      .setText('Formula Templates')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('showFormulaTemplatesCard')));

  card.addSection(formulaSection);

  // Upgrade section
  const upgradeSection = CardService.newCardSection()
    .setHeader('Upgrade')
    .addWidget(CardService.newTextParagraph()
      .setText('Unlock unlimited operations and advanced features'))
    .addWidget(CardService.newTextButton()
      .setText('View Plans')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('showUpgradeCard')));

  card.addSection(upgradeSection);

  return card.build();
}

/**
 * CardService handlers for each feature
 */
function showDuplicateRemovalCard() {
  try {
    const range = SpreadsheetApp.getActiveRange();
    const data = range.getValues();
    
    if (data.length === 0) {
      return buildErrorCard('No Data Selected', 'Please select a range with data to remove duplicates.');
    }

    const card = CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader()
        .setTitle('Remove Duplicates')
        .setSubtitle('Configure duplicate detection'))
      .addSection(CardService.newCardSection()
        .setHeader('Settings')
        .addWidget(CardService.newTextParagraph()
          .setText('Similarity Threshold: 85%'))
        .addWidget(CardService.newDecoratedText()
          .setText('Case Sensitive')
          .setSwitchControl(CardService.newSwitch()
            .setFieldName('caseSensitive')
            .setValue(false))))
      .addSection(CardService.newCardSection()
        .addWidget(CardService.newTextButton()
          .setText('Remove Duplicates')
          .setOnClickAction(CardService.newAction()
            .setFunctionName('processDuplicateRemoval'))))
      .build();

    return card;
  } catch (error) {
    return buildErrorCard('Error', 'Failed to load duplicate removal: ' + error.message);
  }
}

function showTextStandardizationCard() {
  const card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
      .setTitle('Standardize Text')
      .setSubtitle('Clean and format text data'))
    .addSection(CardService.newCardSection()
      .setHeader('Text Case')
      .addWidget(CardService.newSelectionInput()
        .setType(CardService.SelectionInputType.DROPDOWN)
        .setTitle('Case Format')
        .setFieldName('caseFormat')
        .addItem('Keep Original', 'original', true)
        .addItem('Title Case', 'title', false)
        .addItem('UPPER CASE', 'upper', false)
        .addItem('lower case', 'lower', false)))
    .addSection(CardService.newCardSection()
      .setHeader('Cleanup Options')
      .addWidget(CardService.newDecoratedText()
        .setText('Trim whitespace')
        .setSwitchControl(CardService.newSwitch()
          .setFieldName('trimSpaces')
          .setValue(true)))
      .addWidget(CardService.newDecoratedText()
        .setText('Remove extra spaces')
        .setSwitchControl(CardService.newSwitch()
          .setFieldName('removeExtraSpaces')
          .setValue(true)))
      .addWidget(CardService.newDecoratedText()
        .setText('Remove special characters')
        .setSwitchControl(CardService.newSwitch()
          .setFieldName('removeSpecialChars')
          .setValue(false))))
    .addSection(CardService.newCardSection()
      .addWidget(CardService.newTextButton()
        .setText('Apply Standardization')
        .setOnClickAction(CardService.newAction()
          .setFunctionName('processTextStandardization'))))
    .build();

  return card;
}

function showFormulaBuilderCard() {
  // Check tier access
  const userTier = UserSettings.load('userTier', 'free');
  if (!Config.FEATURE_ACCESS.formula_builder.includes(userTier)) {
    return showUpgradeCard();
  }

  const card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
      .setTitle('Formula Builder')
      .setSubtitle('Describe what you want to calculate'))
    .addSection(CardService.newCardSection()
      .setHeader('Natural Language Input')
      .addWidget(CardService.newTextInput()
        .setFieldName('formulaDescription')
        .setTitle('Describe your formula')
        .setHint('e.g., "Sum sales where region is North"')
        .setMultiline(true)))
    .addSection(CardService.newCardSection()
      .addWidget(CardService.newTextButton()
        .setText('Generate Formula')
        .setOnClickAction(CardService.newAction()
          .setFunctionName('processFormulaGeneration'))))
    .build();

  return card;
}

function showUpgradeCard() {
  const card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
      .setTitle('Upgrade CellPilot')
      .setSubtitle('Unlock advanced features'))
    .addSection(CardService.newCardSection()
      .setHeader('Starter Plan - £5.99/month')
      .addWidget(CardService.newTextParagraph()
        .setText('• 500 operations per month\n• Formula builder\n• Email automation')))
    .addSection(CardService.newCardSection()
      .setHeader('Professional Plan - £11.99/month')
      .addWidget(CardService.newTextParagraph()
        .setText('• Unlimited operations\n• All automation features\n• Priority support')))
    .addSection(CardService.newCardSection()
      .addWidget(CardService.newTextButton()
        .setText('Visit CellPilot.co.uk')
        .setOpenLink(CardService.newOpenLink()
          .setUrl('https://www.cellpilot.io')
          .setOpenAs(CardService.OpenAs.FULL_SIZE))))
    .build();

  return card;
}

/**
 * Process functions for CardService actions
 */
function processDuplicateRemoval(e) {
  try {
    const formInput = e.formInput || {};
    const threshold = 0.85; // Fixed for now since slider not available
    const caseSensitive = formInput.caseSensitive || false;

    const result = removeDuplicatesProcess({ threshold, caseSensitive });
    
    if (result.success) {
      return buildSuccessCard('Duplicates Removed', result.message);
    } else {
      return buildErrorCard('Error', result.error);
    }
  } catch (error) {
    return buildErrorCard('Error', 'Failed to remove duplicates: ' + error.message);
  }
}

function processTextStandardization(e) {
  try {
    const formInput = e.formInput || {};
    const options = {
      case: formInput.caseFormat || 'original',
      trim: formInput.trimSpaces !== false,
      removeExtraSpaces: formInput.removeExtraSpaces !== false,
      removeSpecialChars: formInput.removeSpecialChars || false
    };

    const result = standardizeText(options);
    
    if (result.success) {
      return buildSuccessCard('Text Standardized', result.message);
    } else {
      return buildErrorCard('Error', result.error);
    }
  } catch (error) {
    return buildErrorCard('Error', 'Failed to standardize text: ' + error.message);
  }
}

function processFormulaGeneration(e) {
  try {
    const formInput = e.formInput || {};
    const description = formInput.formulaDescription;

    if (!description) {
      return buildErrorCard('Input Required', 'Please describe what you want to calculate');
    }

    const result = generateFormulaFromDescription(description);
    
    if (result.success) {
      return buildFormulaResultCard(result, description);
    } else {
      return buildErrorCard('Formula Generation Failed', result.error);
    }
  } catch (error) {
    return buildErrorCard('Error', 'Failed to generate formula: ' + error.message);
  }
}

/**
 * Helper functions for CardService
 */
function buildErrorCard(title, message) {
  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
      .setTitle(title))
    .addSection(CardService.newCardSection()
      .addWidget(CardService.newTextParagraph()
        .setText(message))
      .addWidget(CardService.newTextButton()
        .setText('Back to CellPilot')
        .setOnClickAction(CardService.newAction()
          .setFunctionName('buildHomepage'))))
    .build();
}

function buildSuccessCard(title, message) {
  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
      .setTitle(title))
    .addSection(CardService.newCardSection()
      .addWidget(CardService.newTextParagraph()
        .setText(message))
      .addWidget(CardService.newTextButton()
        .setText('Back to CellPilot')
        .setOnClickAction(CardService.newAction()
          .setFunctionName('buildHomepage'))))
    .build();
}

function buildFormulaResultCard(result, description) {
  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
      .setTitle('Generated Formula'))
    .addSection(CardService.newCardSection()
      .setHeader('Your Request')
      .addWidget(CardService.newTextParagraph()
        .setText(description)))
    .addSection(CardService.newCardSection()
      .setHeader('Generated Formula')
      .addWidget(CardService.newTextParagraph()
        .setText(result.formula))
      .addWidget(CardService.newTextButton()
        .setText('Insert into Active Cell')
        .setOnClickAction(CardService.newAction()
          .setFunctionName('insertGeneratedFormula')
          .setParameters({ formula: result.formula }))))
    .build();
}

function insertGeneratedFormula(e) {
  try {
    const formula = e.parameters.formula;
    const result = insertFormulaIntoCell(formula);
    
    if (result.success) {
      return buildSuccessCard('Formula Inserted', result.message);
    } else {
      return buildErrorCard('Error', result.error);
    }
  } catch (error) {
    return buildErrorCard('Error', 'Failed to insert formula: ' + error.message);
  }
}

/**
 * Sidebar functions for menu-based approach
 * These maintain the existing functionality for independent installation
 */
function showCellPilotSidebar() {
  try {
    const context = getCurrentUserContext();
    const html = createMainSidebarHtml(context);
    
    SpreadsheetApp.getUi()
      .showSidebar(html.setTitle('CellPilot').setWidth(320));
      
  } catch (error) {
    showErrorDialog('Failed to load CellPilot', error.message);
  }
}

/**
 * Include function for HTML templates
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Create the main sidebar HTML
 */
function createMainSidebarHtml(context) {
  const template = HtmlService.createTemplate(`
    <?!= include('SharedStyles') ?>
    <style>
      /* Quick action button grid */
      .quick-actions-grid {
        display: grid;
        grid-template-columns: 1fr 1fr;
        gap: 10px;
        margin-bottom: 16px;
      }
      
      .quick-action-btn {
        background: white;
        border: 1px solid var(--gray-200);
        border-radius: 8px;
        padding: 12px;
        cursor: pointer;
        transition: all 0.2s ease;
        text-align: center;
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
        gap: 0;
        position: relative;
        overflow: visible;
        min-height: 54px;
      }
      
      /* Very light, subtle theme colors for different tab types */
      .quick-action-btn[data-theme="data"] {
        background: linear-gradient(135deg, #fafbfc 0%, #f0f7ff 100%);
        border-color: #d1dae6;
      }
      
      .quick-action-btn[data-theme="formula"] {
        background: linear-gradient(135deg, #fafdf9 0%, #f0fdf4 100%);
        border-color: #d1e6d8;
      }
      
      .quick-action-btn[data-theme="advanced"] {
        background: linear-gradient(135deg, #fffdf8 0%, #fef9f0 100%);
        border-color: #e6dcc8;
      }
      
      .quick-action-btn[data-theme="pipeline"] {
        background: linear-gradient(135deg, #fdfafd 0%, #fdf0f8 100%);
        border-color: #e6d1dc;
      }
      
      .quick-action-btn:hover {
        transform: translateY(-1px);
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.1);
      }
      
      .quick-action-btn[data-theme="data"]:hover {
        border-color: #c7d2e7;
        background: linear-gradient(135deg, #e8f2ff 0%, #dae8ff 100%);
      }
      
      .quick-action-btn[data-theme="formula"]:hover {
        border-color: #c7e7d0;
        background: linear-gradient(135deg, #e8f9ec 0%, #daf5e3 100%);
      }
      
      .quick-action-btn[data-theme="advanced"]:hover {
        border-color: #e5d5b8;
        background: linear-gradient(135deg, #fef8ec 0%, #fdf3e3 100%);
      }
      
      .quick-action-btn[data-theme="pipeline"]:hover {
        border-color: #e7c7dd;
        background: linear-gradient(135deg, #fae8f5 0%, #f5daea 100%);
      }
      
      .quick-action-icon {
        width: 32px;
        height: 32px;
        border-radius: 8px;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 13px;
        font-weight: 700;
      }
      
      .quick-action-label {
        font-size: 11px;
        font-weight: 500;
        color: var(--gray-700);
        line-height: 1.3;
        letter-spacing: 0.01em;
      }
      
      .quick-badge {
        position: absolute;
        top: -6px;
        right: -6px;
        font-size: 8px;
        padding: 2px 5px;
        border-radius: 6px;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.3px;
        border: 2px solid white;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.1);
      }
      
      .quick-badge.ml {
        background: var(--primary-500);
        color: white;
      }
      
      .quick-badge.new {
        background: var(--success-500);
        color: white;
      }
      
      /* Dropdown sections */
      .dropdown-section {
        margin-bottom: 12px;
      }
      
      .dropdown-section[data-theme="industry"] .dropdown-header {
        background: linear-gradient(135deg, #fcfaff 0%, #f5f0ff 100%);
        border-color: #dcd1e6;
      }
      
      .dropdown-section[data-theme="industry"] .dropdown-header:hover {
        border-color: #c8b8e0;
        background: linear-gradient(135deg, #f2e8ff 0%, #e8daff 100%);
      }
      
      .dropdown-header {
        background: white;
        border: 1px solid var(--gray-200);
        border-radius: 8px;
        padding: 12px 14px;
        cursor: pointer;
        transition: all 0.2s ease;
        display: flex;
        align-items: center;
        justify-content: space-between;
      }
      
      .dropdown-header:hover {
        border-color: var(--primary-400);
        background: var(--primary-50);
        box-shadow: 0 2px 8px rgba(99, 102, 241, 0.1);
      }
      
      .dropdown-header.active {
        border-color: var(--primary-500);
        background: linear-gradient(135deg, rgba(99, 102, 241, 0.05), rgba(99, 102, 241, 0.1));
      }
      
      .dropdown-title {
        display: flex;
        align-items: center;
        gap: 10px;
      }
      
      .dropdown-icon {
        width: 28px;
        height: 28px;
        border-radius: 6px;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 12px;
        font-weight: 700;
      }
      
      .dropdown-text {
        flex: 1;
      }
      
      .dropdown-label {
        font-size: 13px;
        font-weight: 600;
        color: var(--gray-800);
      }
      
      .dropdown-sublabel {
        font-size: 11px;
        color: var(--gray-500);
        margin-top: 1px;
      }
      
      .dropdown-arrow {
        width: 20px;
        height: 20px;
        display: flex;
        align-items: center;
        justify-content: center;
        color: var(--gray-400);
        transition: transform 0.2s ease;
      }
      
      .dropdown-header.active .dropdown-arrow {
        transform: rotate(90deg);
      }
      
      .dropdown-content {
        display: none;
        margin-top: 8px;
        padding: 8px;
        background: var(--gray-50);
        border-radius: 8px;
      }
      
      .dropdown-content.show {
        display: block;
      }
      
      .dropdown-item {
        padding: 10px 12px;
        background: white;
        border-radius: 6px;
        margin-bottom: 6px;
        cursor: pointer;
        transition: var(--transition-base);
        font-size: 12px;
        color: var(--gray-700);
        display: flex;
        align-items: center;
        justify-content: space-between;
      }
      
      .dropdown-item:last-child {
        margin-bottom: 0;
      }
      
      .dropdown-item:hover {
        background: var(--primary-50);
        color: var(--primary-700);
        transform: translateX(2px);
      }
      
      .dropdown-item-badge {
        font-size: 10px;
        padding: 2px 6px;
        border-radius: 10px;
        background: var(--gray-200);
        color: var(--gray-600);
        font-weight: 600;
      }
      
      .dropdown-item:hover .dropdown-item-badge {
        background: var(--primary-200);
        color: var(--primary-700);
      }
      
      /* Footer Styles */
      .footer {
        margin-top: 24px;
        padding: 12px 16px;
        background: var(--gray-50);
        border-top: 1px solid var(--gray-200);
        text-align: center;
        font-size: 12px;
      }
      
      .footer-links {
        display: flex;
        justify-content: center;
        align-items: center;
        gap: 16px;
      }
      
      .footer-link {
        color: var(--gray-600);
        text-decoration: none;
        padding: 4px 8px;
        border-radius: 4px;
        transition: all 0.2s ease;
        font-weight: 500;
      }
      
      .footer-link:hover {
        color: var(--primary-600);
        background: var(--primary-50);
      }
      
      .footer-divider {
        color: var(--gray-400);
        font-size: 10px;
      }
    </style>
    
    <div class="nav-header" style="padding: 12px 16px; border-bottom: 1px solid var(--gray-200);">
      <div style="display: flex; align-items: center; justify-content: center;">
        <a href="https://www.cellpilot.io" target="_blank" style="text-decoration: none;">
          <svg viewBox="0 0 160 50" width="140" height="44" xmlns="http://www.w3.org/2000/svg" style="cursor: pointer;">
            <g transform="translate(6, 6)">
              <path d="M4 4 Q4 0 8 0 L30 0 Q34 0 34 4 L34 30 Q34 34 30 34 L8 34 Q4 34 4 30 Z" 
                    fill="none" stroke="#2563eb" stroke-width="2"/>
              <line x1="19" y1="0" x2="19" y2="34" stroke="#60a5fa" stroke-width="0.8"/>
              <line x1="4" y1="17" x2="34" y2="17" stroke="#60a5fa" stroke-width="0.8"/>
              <g transform="translate(7, 7)">
                <rect x="0" y="0" width="5" height="5" fill="#e5e7eb" opacity="0.8"/>
                <rect x="5" y="0" width="5" height="5" fill="#3b82f6" opacity="0.3"/>
                <rect x="10" y="0" width="5" height="5" fill="#60a5fa" opacity="0.2"/>
                <rect x="15" y="0" width="5" height="5" fill="#e5e7eb" opacity="0.8"/>
                <rect x="0" y="5" width="5" height="5" fill="#60a5fa" opacity="0.2"/>
                <rect x="5" y="5" width="5" height="5" fill="#3b82f6"/>
                <rect x="10" y="5" width="5" height="5" fill="#e5e7eb" opacity="0.8"/>
                <rect x="15" y="5" width="5" height="5" fill="#60a5fa" opacity="0.3"/>
              </g>
              <g transform="translate(19, 17)">
                <circle r="2.5" fill="none" stroke="#2563eb" stroke-width="1.2"/>
                <line x1="-5" y1="0" x2="5" y2="0" stroke="#2563eb" stroke-width="0.8"/>
                <line x1="0" y1="-5" x2="0" y2="5" stroke="#2563eb" stroke-width="0.8"/>
              </g>
            </g>
            <text x="48" y="32" font-family="system-ui, -apple-system, sans-serif" font-size="20" font-weight="600" fill="#1f2937">
              Cell<tspan fill="#2563eb">Pilot</tspan>
            </text>
          </svg>
        </a>
        <!-- Tier Badge -->
        <div style="position: absolute; right: 16px; top: 50%; transform: translateY(-50%);">
          <span style="display: inline-block; padding: 4px 10px; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; font-size: 10px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px; border-radius: 12px;">
            <?= context.tier || 'BETA' ?>
          </span>
        </div>
      </div>
    </div>
    
    <div id="undoBar" style="display: none; background: #fef3c7; padding: 10px; margin: 0 10px 10px; border-radius: 8px;">
      <div style="display: flex; align-items: center; justify-content: space-between;">
        <span style="font-size: 12px; color: #92400e;">
          <strong id="undoOperation"></strong> can be undone
        </span>
        <button onclick="performUndo()" style="background: #f59e0b; color: white; border: none; padding: 4px 12px; border-radius: 4px; font-size: 12px; cursor: pointer;">
          Undo
        </button>
      </div>
    </div>
    
    <div class="container">
      <div id="selectionInfo">
        <? if (!context.hasSelection) { ?>
          <div class="alert alert-warning" id="noSelectionAlert">
            Select cells in your spreadsheet to enable all features
          </div>
        <? } else { ?>
          <div class="card" style="padding: 12px; margin-bottom: 16px;" id="selectionCard">
            <div style="font-size: 11px; font-weight: 600; color: var(--gray-600); text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 6px;">Current Selection</div>
            <div style="font-size: 13px; color: var(--gray-800);" id="selectionDetails">
              <span id="currentRange"><?= context.range ?></span> • <span id="currentRows"><?= context.rowCount ?></span>×<span id="currentCols"><?= context.colCount ?></span> <span id="currentType"><?= context.dataType ?></span>
            </div>
          </div>
        <? } ?>
      </div>
      
      <!-- CellM8 Presentation Helper -->
      <div style="margin-bottom: 20px;">
        <div class="card" style="padding: 16px; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); position: relative; overflow: hidden;">
          <div style="position: absolute; top: 8px; right: 8px;">
            <span style="display: inline-block; padding: 3px 8px; background: rgba(255, 255, 255, 0.9); color: #764ba2; font-size: 9px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px; border-radius: 10px;">
              NEW
            </span>
          </div>
          <h3 style="font-size: 16px; font-weight: 700; color: white; margin-bottom: 6px;">
            CellM8 Presentation Helper
          </h3>
          <p style="font-size: 12px; color: rgba(255, 255, 255, 0.9); margin-bottom: 12px; line-height: 1.4;">
            Transform your spreadsheet data into stunning Google Slides presentations with AI-powered intelligence
          </p>
          <button onclick="google.script.run.showCellM8()" style="width: 100%; padding: 10px 16px; background: white; color: #667eea; border: none; border-radius: 6px; font-size: 13px; font-weight: 600; cursor: pointer; transition: all 0.2s ease;">
            Create Presentation →
          </button>
        </div>
      </div>
      
      <!-- Quick Actions Grid -->
      <div style="margin-bottom: 16px; position: relative;">
        <h3 style="font-size: 12px; font-weight: 600; color: var(--gray-600); text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 10px;">Quick Actions</h3>
        <div class="quick-actions-grid" style="position: relative;">
          <div class="quick-action-btn" data-theme="data" onclick="google.script.run.showTableize()">
            <div class="quick-action-label">Tableize Data</div>
            <span class="quick-badge ml">ML</span>
          </div>
          
          <div class="quick-action-btn" data-theme="data" onclick="google.script.run.showDataCleaning()">
            <div class="quick-action-label">Data Cleaning</div>
            <span class="quick-badge ml">ML</span>
          </div>
          
          <div class="quick-action-btn" data-theme="data" onclick="google.script.run.showAdvancedRestructuring()">
            <div class="quick-action-label">Restructure Data</div>
            <span class="quick-badge new">New</span>
          </div>
          
          <div class="quick-action-btn" data-theme="formula" onclick="google.script.run.showSmartFormulaAssistant()">
            <div class="quick-action-label">Smart Formulas</div>
            <span class="quick-badge ml">ML</span>
          </div>
          
          <div class="quick-action-btn" data-theme="advanced" onclick="google.script.run.showAutomation()">
            <div class="quick-action-label">Automation</div>
          </div>
          
          <div class="quick-action-btn" data-theme="formula" onclick="google.script.run.showPivotTableAssistant()">
            <div class="quick-action-label">Pivot Tables</div>
            <span class="quick-badge ml">ML</span>
          </div>
          
          <div class="quick-action-btn" data-theme="pipeline" onclick="google.script.run.showDataPipelineManager()">
            <div class="quick-action-label">Data Pipeline</div>
            <span class="quick-badge new">NEW</span>
          </div>
          
          <div class="quick-action-btn" data-theme="advanced" onclick="toggleDropdown('advanced')">
            <div class="quick-action-label">Advanced Tools</div>
            <span class="quick-badge">6</span>
          </div>
        </div>
        
        <!-- Advanced Tools Dropdown (positioned within grid container) -->
        <div id="advanced-dropdown" class="dropdown-content" style="display: none; position: absolute; background: white; border: 1px solid var(--gray-200); border-radius: 8px; padding: 8px; box-shadow: 0 4px 12px rgba(0,0,0,0.1); z-index: 1000; width: 280px; max-height: 300px; overflow-y: auto;">
          <div class="dropdown-item" onclick="google.script.run.showExcelMigration()">
            Excel Migration Assistant
          </div>
          <div class="dropdown-item" onclick="google.script.run.showApiIntegration()">
            API Integration
          </div>
          <div class="dropdown-item" onclick="google.script.run.showDataValidation()">
            Data Validation Rules
          </div>
          <div class="dropdown-item" onclick="google.script.run.showBatchOperations()">
            Batch Operations
          </div>
          <div class="dropdown-item" onclick="google.script.run.showUpgradeDialog()">
            Premium Features
            <span class="dropdown-item-badge">PRO</span>
          </div>
          <div class="dropdown-item" onclick="google.script.run.showErrorDialog('Feature Coming Soon', 'This feature is under development')">
            Advanced Analytics
            <span class="dropdown-item-badge">SOON</span>
          </div>
        </div>
      </div>
      
      <!-- Industry Templates Dropdown -->
      <div class="dropdown-section" data-theme="industry">
        <div class="dropdown-header" onclick="toggleDropdown('industry')">
          <div class="dropdown-title">
            <div class="dropdown-text">
              <div class="dropdown-label">Industry Templates</div>
              <div class="dropdown-sublabel">Pre-built solutions for your business</div>
            </div>
          </div>
          <div class="dropdown-arrow">→</div>
        </div>
        <div id="industry-dropdown" class="dropdown-content">
          <div class="dropdown-item" onclick="google.script.run.showIndustryTemplate('real-estate')">
            Real Estate
            <span class="dropdown-item-badge">4 templates</span>
          </div>
          <div class="dropdown-item" onclick="google.script.run.showIndustryTemplate('construction')">
            Construction
            <span class="dropdown-item-badge">4 templates</span>
          </div>
          <div class="dropdown-item" onclick="google.script.run.showIndustryTemplate('healthcare')">
            Healthcare
            <span class="dropdown-item-badge">4 templates</span>
          </div>
          <div class="dropdown-item" onclick="google.script.run.showIndustryTemplate('marketing')">
            Marketing Agency
            <span class="dropdown-item-badge">5 templates</span>
          </div>
          <div class="dropdown-item" onclick="google.script.run.showIndustryTemplate('ecommerce')">
            E-Commerce
            <span class="dropdown-item-badge">5 templates</span>
          </div>
          <div class="dropdown-item" onclick="google.script.run.showIndustryTemplate('professional')">
            Professional Services
            <span class="dropdown-item-badge">5 templates</span>
          </div>
        </div>
      </div>
      
      
      <!-- Upgrade Card -->
      <div class="card" style="background: linear-gradient(135deg, rgba(99, 102, 241, 0.05), rgba(99, 102, 241, 0.1)); border: 1px solid var(--primary-200); margin-top: 16px; padding: 10px 12px; overflow: hidden;">
        <div style="text-align: center;">
          <div style="font-size: 13px; font-weight: 600; color: var(--primary-700); margin-bottom: 4px;">Upgrade to Pro</div>
          <div style="font-size: 10px; color: var(--gray-600); margin-bottom: 8px;">
            Unlimited operations • Advanced features
          </div>
          <button class="btn btn-primary btn-no-ripple" style="width: 100%; padding: 5px; font-size: 11px; font-weight: 600;" onclick="event.stopPropagation(); google.script.run.showUpgradeOptions()">
            View Plans
          </button>
        </div>
      </div>
      
      <div class="footer">
        <div class="footer-links">
          <a href="#" class="footer-link" onclick="google.script.run.showSettings(); return false;">Settings</a>
          <span class="footer-divider">•</span>
          <a href="#" class="footer-link" onclick="google.script.run.showHelp(); return false;">Help</a>
          <span class="footer-divider">•</span>
          <a href="https://www.cellpilot.io" target="_blank" class="footer-link">Website</a>
        </div>
      </div>
    </div>
    
    <script>
      let currentContext = null;
      
      // Initialize real-time updates
      window.onload = function() {
        // Start periodic update for real-time selection changes
        setInterval(updateSelectionInfo, 2000); // Update every 2 seconds
      };
      
      function updateSelectionInfo() {
        google.script.run
          .withSuccessHandler(function(context) {
            // Only update if context has changed
            if (JSON.stringify(context) !== JSON.stringify(currentContext)) {
              currentContext = context;
              
              if (context && context.hasSelection) {
                // Show selection card, hide no selection alert
                const selectionCard = document.getElementById('selectionCard');
                const noSelectionAlert = document.getElementById('noSelectionAlert');
                const selectionInfo = document.getElementById('selectionInfo');
                
                if (!selectionCard && noSelectionAlert) {
                  // Create selection card from no selection state
                  selectionInfo.innerHTML = \`
                    <div class="card" style="padding: 12px; margin-bottom: 16px;" id="selectionCard">
                      <div style="font-size: 11px; font-weight: 600; color: var(--gray-600); text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 6px;">Current Selection</div>
                      <div style="font-size: 13px; color: var(--gray-800);" id="selectionDetails">
                        <span id="currentRange"></span> • <span id="currentRows"></span>×<span id="currentCols"></span> <span id="currentType"></span>
                      </div>
                    </div>
                  \`;
                }
                
                // Update the selection details
                document.getElementById('currentRange').textContent = context.range;
                document.getElementById('currentRows').textContent = context.rowCount;
                document.getElementById('currentCols').textContent = context.colCount;
                document.getElementById('currentType').textContent = context.dataType;
                
              } else {
                // Show no selection alert, hide selection card
                const selectionCard = document.getElementById('selectionCard');
                const selectionInfo = document.getElementById('selectionInfo');
                
                if (selectionCard) {
                  selectionInfo.innerHTML = \`
                    <div class="alert alert-warning" id="noSelectionAlert">
                      Select cells in your spreadsheet to enable all features
                    </div>
                  \`;
                }
              }
            }
          })
          .withFailureHandler(function(error) {
            Logger.error('Error updating selection info:', error);
          })
          .getCurrentUserContext();
      }
      
      // Check for available undo on load
      google.script.run
        .withSuccessHandler(function(info) {
          if (info) {
            document.getElementById('undoBar').style.display = 'block';
            document.getElementById('undoOperation').textContent = info.operationType;
          }
        })
        .getUndoInfo();
      
      function performUndo() {
        google.script.run
          .withSuccessHandler(function(result) {
            if (result.success) {
              document.getElementById('undoBar').style.display = 'none';
              // Show success message
              const alert = document.createElement('div');
              alert.className = 'alert alert-success';
              alert.textContent = result.message;
              alert.style.margin = '10px';
              document.querySelector('.container').insertBefore(alert, document.querySelector('.container').firstChild);
              setTimeout(function() { alert.remove(); }, 3000);
            } else {
              alert(result.error || 'Undo failed');
            }
          })
          .withFailureHandler(function(error) {
            alert('Undo failed: ' + error.message);
          })
          .performUndo();
      }
      
      function toggleDropdown(type) {
        const dropdownId = type + '-dropdown';
        const dropdown = document.getElementById(dropdownId);
        
        if (type === 'advanced') {
          // Special handling for advanced tools - position relative to button
          const button = event.currentTarget;
          const buttonRect = button.getBoundingClientRect();
          const container = button.closest('.quick-actions-grid');
          const containerRect = container.getBoundingClientRect();
          
          if (dropdown.style.display === 'block') {
            dropdown.style.display = 'none';
          } else {
            // Position dropdown relative to the container
            dropdown.style.display = 'block';
            dropdown.style.top = (buttonRect.bottom - containerRect.top + 5) + 'px';
            dropdown.style.left = (buttonRect.left - containerRect.left) + 'px';
            
            // Make sure dropdown doesn't go off screen
            const dropdownRect = dropdown.getBoundingClientRect();
            if (dropdownRect.right > window.innerWidth) {
              dropdown.style.left = (buttonRect.right - containerRect.left - 280) + 'px';
            }
          }
        } else {
          const header = dropdown.previousElementSibling;
          
          // Toggle current dropdown
          const isOpen = dropdown.classList.contains('show');
          
          // Close all dropdowns first
          document.querySelectorAll('.dropdown-content').forEach(function(dd) {
            dd.classList.remove('show');
          });
          document.querySelectorAll('.dropdown-header').forEach(function(hdr) {
            hdr.classList.remove('active');
          });
          
          // Open clicked dropdown if it was closed
          if (!isOpen) {
            dropdown.classList.add('show');
            header.classList.add('active');
          }
        }
      }
      
      // Close advanced dropdown when clicking outside
      document.addEventListener('click', function(e) {
        const advDropdown = document.getElementById('advanced-dropdown');
        if (advDropdown && advDropdown.style.display === 'block') {
          const isAdvButton = e.target.closest('[data-theme="advanced"]');
          if (!isAdvButton && !advDropdown.contains(e.target)) {
            advDropdown.style.display = 'none';
          }
        }
      });
    </script>
  `);
  
  template.context = context;
  return template.evaluate().setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Get undo information for display
 */
function getUndoInfo() {
  return UndoManager.getUndoInfo();
}

/**
 * Perform undo operation
 */
function performUndo() {
  return UndoManager.undo();
}

/**
 * Show data cleaning interface (main entry point)
 */
function showDataCleaning() {
  try {
    const range = SpreadsheetApp.getActiveRange();
    const data = range.getValues();

    if (data.length === 0) {
      showErrorDialog('No Data Selected', 'Please select a range with data to clean.');
      return;
    }

    const html = HtmlService.createTemplateFromFile('DuplicateRemovalTemplate')
      .evaluate()
      .setTitle('Data Cleaning')
      .setWidth(400)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

    SpreadsheetApp.getUi().showSidebar(html);

  } catch (error) {
    showErrorDialog('Error', 'Failed to load data cleaning: ' + error.message);
  }
}

/**
 * Show duplicate removal interface (legacy - redirects to data cleaning)
 */
function showDuplicateRemoval() {
  showDataCleaning();
}

/**
 * Actually remove duplicates (called from the HTML interface)
 */
function removeDuplicatesProcess(options) {
  try {
    const range = SpreadsheetApp.getActiveRange();
    const totalRows = range.getNumRows();
    
    // For large datasets, process in batches
    const isLargeDataset = totalRows > 1000;
    
    if (isLargeDataset) {
      // Process large datasets efficiently
      return processLargeDuplicateRemoval(range, options);
    }
    
    // Standard processing for smaller datasets
    const data = range.getValues();
    if (data.length === 0) {
      return { success: false, error: 'No data selected' };
    }
    
    // Check usage limits
    const usageCheck = UsageTracker.track('operations');
    if (!usageCheck.allowed) {
      return {
        success: false,
        error: 'Usage limit reached. Please upgrade your plan.',
        showUpgrade: true
      };
    }
    
    // Save state for undo
    UndoManager.saveState(range, 'Duplicate Removal');
    
    // Enhanced duplicate detection with messy data handling
    const processedData = data.map(row => row.map(cell => {
      // Clean messy data before comparison
      return Utils.cleanMessyData(cell, {
        handleMissingData: true,
        missingDataReplacement: '',
        removeInvisibleChars: true,
        fixEncoding: true
      });
    }));
    
    const threshold = options.threshold || 0.85;
    const caseSensitive = options.caseSensitive || false;
    const duplicates = DataCleaner.findDuplicates(processedData, { threshold, caseSensitive });
    const cleanedData = DataCleaner.removeDuplicates(data, duplicates);
    
    // Apply changes
    range.clearContent();
    if (cleanedData.length > 0) {
      range.getSheet().getRange(range.getRow(), range.getColumn(),
        cleanedData.length, cleanedData[0].length).setValues(cleanedData);
    }
    
    const removedCount = data.length - cleanedData.length;
    return {
      success: true,
      message: `Successfully removed ${removedCount} duplicate rows`,
      details: `Processed ${data.length} rows, kept ${cleanedData.length} unique rows`,
      removedCount: removedCount
    };
  } catch (error) {
    Logger.error('Error removing duplicates:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Process large duplicate removal in batches
 */
function processLargeDuplicateRemoval(range, options) {
  try {
    const sheet = range.getSheet();
    const startRow = range.getRow();
    const startCol = range.getColumn();
    const totalRows = range.getNumRows();
    const totalCols = range.getNumColumns();
    
    // Check usage limits
    const usageCheck = UsageTracker.track('operations');
    if (!usageCheck.allowed) {
      return {
        success: false,
        error: 'Usage limit reached. Please upgrade your plan.',
        showUpgrade: true
      };
    }
    
    // Save state for undo
    UndoManager.saveState(range, 'Large Duplicate Removal');
    
    const uniqueRows = new Map();
    const batchSize = 500;
    let processedRows = 0;
    
    // Process in batches to avoid timeout
    for (let batchStart = 0; batchStart < totalRows; batchStart += batchSize) {
      const numRows = Math.min(batchSize, totalRows - batchStart);
      const batchRange = sheet.getRange(startRow + batchStart, startCol, numRows, totalCols);
      const batchData = batchRange.getValues();
      
      // Process each row in the batch
      batchData.forEach((row, index) => {
        const cleanedRow = row.map(cell => 
          Utils.cleanMessyData(cell, {
            handleMissingData: true,
            removeInvisibleChars: true
          })
        );
        
        // Create key for duplicate detection
        const key = options.caseSensitive ? 
          cleanedRow.join('|') : 
          cleanedRow.join('|').toLowerCase();
        
        // Store only unique rows
        if (!uniqueRows.has(key)) {
          uniqueRows.set(key, row);
        }
      });
      
      processedRows += numRows;
      Utils.showProgress(processedRows, totalRows, 'Analyzing duplicates');
      
      // Prevent timeout
      if (batchStart % (batchSize * 5) === 0 && batchStart > 0) {
        Utilities.sleep(100);
      }
    }
    
    // Convert unique rows back to array
    const cleanedData = Array.from(uniqueRows.values());
    
    // Clear original range and write cleaned data
    range.clearContent();
    if (cleanedData.length > 0) {
      const outputRange = sheet.getRange(startRow, startCol, 
        cleanedData.length, cleanedData[0].length);
      
      // Write in batches to avoid memory issues
      for (let i = 0; i < cleanedData.length; i += batchSize) {
        const batch = cleanedData.slice(i, Math.min(i + batchSize, cleanedData.length));
        const batchRange = sheet.getRange(startRow + i, startCol, batch.length, batch[0].length);
        batchRange.setValues(batch);
        
        Utils.showProgress(i + batch.length, cleanedData.length, 'Writing cleaned data');
      }
    }
    
    const removedCount = totalRows - cleanedData.length;
    return {
      success: true,
      message: `Successfully removed ${removedCount} duplicate rows from large dataset`,
      details: `Processed ${totalRows} rows, kept ${cleanedData.length} unique rows`,
      removedCount: removedCount,
      isLargeDataset: true
    };
    
  } catch (error) {
    Logger.error('Error in large duplicate removal:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Preview duplicates before removal (called from HTML)
 */
function previewDuplicates(options) {
  try {
    const range = SpreadsheetApp.getActiveRange();
    const data = range.getValues();
    
    if (data.length === 0) {
      return { success: false, error: 'No data selected' };
    }
    
    const threshold = options.threshold || 0.85;
    const caseSensitive = options.caseSensitive || false;
    const trimWhitespace = options.trimWhitespace !== false;
    const ignoreSpecialChars = options.ignoreSpecialChars || false;
    
    // Process data based on options
    let processedData = data;
    if (trimWhitespace) {
      processedData = data.map(row => row.map(cell => 
        typeof cell === 'string' ? cell.trim() : cell
      ));
    }
    
    // Find duplicates
    const duplicateMap = {};
    const duplicates = [];
    
    for (let i = 0; i < processedData.length; i++) {
      const rowStr = processedData[i].join('|');
      const key = caseSensitive ? rowStr : rowStr.toLowerCase();
      
      if (duplicateMap[key]) {
        duplicateMap[key].count++;
      } else {
        duplicateMap[key] = { value: data[i].join(', '), count: 1, index: i };
      }
    }
    
    // Filter to only show actual duplicates
    Object.values(duplicateMap).forEach(item => {
      if (item.count > 1) {
        duplicates.push({
          value: item.value.substring(0, 100), // Truncate for display
          count: item.count
        });
      }
    });
    
    return {
      success: true,
      duplicates: duplicates.slice(0, 20), // Limit preview to 20 items
      duplicateCount: duplicates.length,
      totalRows: data.length
    };
    
  } catch (error) {
    Logger.error('Error previewing duplicates:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Show formula builder interface
 */
function showFormulaBuilder() {
  try {
    // Check feature access with FeatureGate
    if (!FeatureGate.enforceAccess('formula_builder')) {
      return; // FeatureGate will show upgrade dialog
    }
    
    // Track feature usage
    ApiIntegration.trackUsage('formula_builder', 'open');

    const html = HtmlService.createTemplateFromFile('FormulaBuilderTemplate')
      .evaluate()
      .setTitle('Formula Builder')
      .setWidth(400)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

    SpreadsheetApp.getUi().showSidebar(html);

  } catch (error) {
    showErrorDialog('Error', 'Failed to load formula builder: ' + error.message);
  }
}

/**
 * Generate formula from description (called from HTML)
 */
function generateFormulaFromDescription(description) {
  try {
    if (!description || description.trim() === '') {
      return { success: false, error: 'Please provide a formula description' };
    }
    
    // Check tier access
    const userTier = UserSettings.load('userTier', 'free');
    if (!Config.FEATURE_ACCESS.formula_builder.includes(userTier)) {
      return {
        success: false,
        error: 'Formula Builder requires Starter plan or above',
        requiresUpgrade: true
      };
    }

    // Check usage
    const usageCheck = UsageTracker.track('operations');
    if (!usageCheck.allowed) {
      return {
        success: false,
        error: 'Usage limit reached',
        showUpgrade: true
      };
    }

    const formula = FormulaBuilder.parseNaturalLanguageFormula(description.trim());

    if (!formula) {
      return {
        success: false,
        error: 'Could not understand the description. Try using simpler language.',
        suggestions: [
          'Sum sales where region is North',
          'Count completed tasks',
          'Average revenue by month'
        ]
      };
    }

    return {
      success: true,
      formula: formula
    };

  } catch (error) {
    Logger.error('Error generating formula:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Check if user has access to formula builder (called from HTML)
 */
function checkFormulaBuilderAccess() {
  const userTier = UserSettings.load('userTier', 'free');
  return Config.FEATURE_ACCESS.formula_builder.includes(userTier);
}

/**
 * Insert formula into active cell (called from HTML)
 */
function insertFormulaIntoCell(formulaText) {
  try {
    const activeCell = SpreadsheetApp.getActiveCell();
    activeCell.setFormula(formulaText);
    
    return {
      success: true,
      message: `Formula inserted into cell ${activeCell.getA1Notation()}`
    };
    
  } catch (error) {
    Logger.error('Error inserting formula:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Show automation interface
 */
function showAutomation() {
  try {
    // Check feature access with FeatureGate
    if (!FeatureGate.enforceAccess('automation')) {
      return; // FeatureGate will show upgrade dialog
    }
    
    // Track feature usage
    ApiIntegration.trackUsage('automation', 'open');
    
    const html = HtmlService.createTemplateFromFile('AutomationTemplate')
      .evaluate()
      .setTitle('Smart Automation')
      .setWidth(450)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (error) {
    Logger.error('Error showing automation interface:', error);
    showErrorDialog('Error', 'Failed to open automation panel: ' + error.message);
  }
}

/**
 * Scan for Excel migration issues
 */
function scanForExcelIssues() {
  return ExcelMigration.scanForExcelIssues();
}

/**
 * Fix Excel migration issues
 */
function fixExcelIssues(options) {
  return ExcelMigration.fixExcelIssues(options);
}

/**
 * Analyze formats in the current selection
 */
function analyzeFormats(options) {
  try {
    const range = SpreadsheetApp.getActiveRange();
    const formats = range.getNumberFormats();
    const values = range.getValues();
    
    let numberFormats = 0;
    let dateFormats = 0;
    let textFormats = 0;
    const issues = [];
    
    for (let i = 0; i < formats.length; i++) {
      for (let j = 0; j < formats[i].length; j++) {
        const format = formats[i][j];
        const value = values[i][j];
        
        if (format.includes('#') || format.includes('0')) numberFormats++;
        if (format.includes('d') || format.includes('m') || format.includes('y')) dateFormats++;
        if (format === '@') textFormats++;
        
        // Check for potential issues
        if (typeof value === 'number' && value > 999999999) {
          issues.push(`Cell ${range.getCell(i+1, j+1).getA1Notation()}: Large number may convert to scientific notation`);
        }
        if (typeof value === 'string' && !isNaN(value) && value.length > 10) {
          issues.push(`Cell ${range.getCell(i+1, j+1).getA1Notation()}: ID-like text may be converted to number`);
        }
      }
    }
    
    return {
      success: true,
      numberFormats: numberFormats,
      dateFormats: dateFormats,
      textFormats: textFormats,
      issues: issues.slice(0, 10) // Limit to first 10 issues
    };
    
  } catch (error) {
    Logger.error('Error analyzing formats:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Show text standardization interface
 */
function showTableize() {
  try {
    const range = SpreadsheetApp.getActiveRange();
    const data = range.getValues();
    
    if (data.length === 0) {
      showErrorDialog('No Data Selected', 'Please select data to structure into columns.');
      return;
    }
    
    const html = HtmlService.createTemplateFromFile('TableizeTemplate')
      .evaluate()
      .setTitle('Tableize Data')
      .setWidth(450)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    
    SpreadsheetApp.getUi().showSidebar(html);
    
  } catch (error) {
    showErrorDialog('Error', 'Failed to load Tableize: ' + error.message);
  }
}

/**
 * Show Advanced Data Restructuring interface
 */
function showAdvancedRestructuring() {
  try {
    const range = SpreadsheetApp.getActiveRange();
    
    if (!range) {
      showErrorDialog('No Data Selected', 'Please select data to restructure.');
      return;
    }
    
    const html = HtmlService.createTemplateFromFile('AdvancedRestructuringTemplate')
      .evaluate()
      .setTitle('Advanced Data Restructuring')
      .setWidth(480)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    
    SpreadsheetApp.getUi().showSidebar(html);
    
  } catch (error) {
    showErrorDialog('Error', 'Failed to load Advanced Restructuring: ' + error.message);
  }
}

/**
 * Show CellM8 Presentation Helper
 */
function showCellM8() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('CellM8Template')
      .setTitle('CellM8 - Presentation Helper')
      .setWidth(320);
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (error) {
    showErrorDialog('Failed to load CellM8', error.message);
  }
}

/**
 * Show a toast notification
 * @param {string} message - The message to display
 * @param {string} title - The title of the toast (optional)
 * @param {number} timeout - Timeout in seconds (optional)
 */
function showToast(message, title, timeout) {
  const ui = SpreadsheetApp.getUi();
  const actualTitle = title || 'CellPilot';
  const actualTimeout = timeout || 3;
  
  SpreadsheetApp.getActiveSpreadsheet().toast(message, actualTitle, actualTimeout);
}

function hasDataSelection() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const range = sheet.getActiveRange();
    
    if (!range) {
      return false;
    }
    
    // Check if the selection has actual content
    const values = range.getValues();
    if (values.length === 0 || values[0].length === 0) {
      return false;
    }
    
    // Check if at least one cell has content
    for (let row of values) {
      for (let cell of row) {
        if (cell !== null && cell !== undefined && String(cell).trim() !== '') {
          return true;
        }
      }
    }
    
    return false;
  } catch (error) {
    Logger.error('Error checking selection:', error);
    return false;
  }
}

function analyzeDataForTableize() {
  Logger.info('analyzeDataForTableize called');
  
  try {
    const range = SpreadsheetApp.getActiveRange();
    const data = range.getValues();
    
    if (!data || data.length === 0) {
      return { hasData: false, error: 'No data selected' };
    }
    
    // Check if data is already in multiple columns
    const currentColumns = data[0].length;
    const hasMultipleColumns = currentColumns > 1 && 
      data.some(row => row.filter(cell => cell !== '' && cell !== null).length > 1);
    
    // Analyze delimiters in first column (for single column data)
    const delimiters = {
      space: { char: ' ', count: 0, consistent: true },
      comma: { char: ',', count: 0, consistent: true },
      tab: { char: '\t', count: 0, consistent: true },
      pipe: { char: '|', count: 0, consistent: true },
      semicolon: { char: ';', count: 0, consistent: true },
      doubleSpace: { char: '  ', count: 0, consistent: true }
    };
    
    let suggestedDelimiter = 'none';
    let needsTableize = !hasMultipleColumns;
    
    if (!hasMultipleColumns) {
      // Analyze first column for delimiters
      data.forEach(row => {
        const cellValue = String(row[0] || '');
        if (cellValue) {
          if (cellValue.includes(',')) delimiters.comma.count++;
          if (cellValue.includes('\t')) delimiters.tab.count++;
          if (cellValue.includes('|')) delimiters.pipe.count++;
          if (cellValue.includes(';')) delimiters.semicolon.count++;
          if (cellValue.includes('  ')) delimiters.doubleSpace.count++;
          if (cellValue.split(' ').length > 1) delimiters.space.count++;
        }
      });
      
      // Find most common delimiter
      let maxCount = 0;
      Object.entries(delimiters).forEach(([key, value]) => {
        if (value.count > maxCount) {
          maxCount = value.count;
          suggestedDelimiter = key;
        }
      });
    }
    
    // Get first row sample
    const firstRowSample = data[0].map(cell => String(cell || '')).join(' | ').substring(0, 100);
    
    return {
      hasData: true,
      rowCount: data.length,
      columnCount: currentColumns,
      hasMultipleColumns: hasMultipleColumns,
      needsTableize: needsTableize,
      suggestedDelimiter: suggestedDelimiter,
      delimiters: delimiters,
      patterns: { 
        csvLike: delimiters.comma.count > data.length * 0.5,
        keyValue: false 
      },
      hasHeaders: true,
      firstRowSample: firstRowSample,
      estimatedColumns: hasMultipleColumns ? currentColumns : 
        (suggestedDelimiter !== 'none' ? 3 : 1)
    };
    
  } catch (error) {
    Logger.error('Error in analyzeDataForTableize:', error);
    return { hasData: false, error: 'Function error: ' + error.toString() };
  }
}

function previewTableize(options) {
  try {
    const range = SpreadsheetApp.getActiveRange();
    const data = range.getValues();
    
    Logger.info('Preview tableize with options:', options);
    Logger.info('Data dimensions:', data.length, 'rows x', data[0].length, 'columns');
    
    let parsedData = [];
    
    // Check if data is already in multiple columns
    const currentColumns = data[0].length;
    const hasMultipleColumns = currentColumns > 1 && 
      data.some(row => row.filter(cell => cell !== '' && cell !== null).length > 1);
    
    if (hasMultipleColumns && options.preserveColumns !== false) {
      // Data is already in columns - preserve it
      Logger.info('Data already in columns, preserving structure');
      
      parsedData = data.map(row => 
        row.map(cell => String(cell || '').trim())
      );
      
      // Remove empty trailing columns
      if (parsedData.length > 0) {
        const maxNonEmptyCol = parsedData.reduce((max, row) => {
          for (let i = row.length - 1; i >= 0; i--) {
            if (row[i]) return Math.max(max, i);
          }
          return max;
        }, 0);
        
        parsedData = parsedData.map(row => row.slice(0, maxNonEmptyCol + 1));
      }
      
    } else if (options.method === 'smart') {
      // Use smart parser for single column data
      Logger.info('Using smart parser for single column data');
      parsedData = SmartTableParser.smartParse(data, options);
      
      // Add suggested headers if first row isn't headers
      if (!options.hasHeaders && parsedData.length > 0) {
        const firstRow = parsedData[0];
        const headers = [];
        
        if (firstRow.length >= 3) {
          headers.push('Name', 'Region', 'Sales');
        } else if (firstRow.length === 2) {
          headers.push('Name', 'Value');
        } else {
          for (let i = 0; i < firstRow.length; i++) {
            headers.push('Column ' + (i + 1));
          }
        }
        
        parsedData.unshift(headers);
      }
      
    } else {
      // Simple delimiter splitting for single column data
      Logger.info('Using delimiter splitting');
      
      for (let i = 0; i < data.length; i++) {
        if (hasMultipleColumns) {
          // Join all columns then split by delimiter
          const rowText = data[i].map(cell => String(cell || '')).join(' ').trim();
          if (rowText) {
            let columns;
            
            if (options.delimiter === 'space') {
              columns = rowText.split(/\s+/).filter(c => c);
            } else if (options.delimiter === 'comma') {
              columns = rowText.split(',').map(c => c.trim());
            } else if (options.delimiter === 'tab') {
              columns = rowText.split('\t').map(c => c.trim());
            } else if (options.delimiter === 'pipe') {
              columns = rowText.split('|').map(c => c.trim());
            } else if (options.delimiter === 'semicolon') {
              columns = rowText.split(';').map(c => c.trim());
            } else {
              // No delimiter - keep as is
              columns = data[i].map(cell => String(cell || '').trim()).filter(c => c);
            }
            
            if (columns.length > 0) {
              parsedData.push(columns);
            }
          }
        } else {
          // Single column data - split first column only
          const cellValue = String(data[i][0] || '').trim();
          
          if (cellValue) {
            let columns;
            
            if (options.delimiter === 'space') {
              columns = cellValue.split(/\s+/).filter(c => c);
            } else if (options.delimiter === 'comma') {
              columns = cellValue.split(',').map(c => c.trim());
            } else if (options.delimiter === 'tab') {
              columns = cellValue.split('\t').map(c => c.trim());
            } else if (options.delimiter === 'pipe') {
              columns = cellValue.split('|').map(c => c.trim());
            } else if (options.delimiter === 'semicolon') {
              columns = cellValue.split(';').map(c => c.trim());
            } else {
              columns = [cellValue];
            }
            
            if (columns.length > 0) {
              parsedData.push(columns);
            }
          }
        }
      }
    }
    
    const maxColumns = Math.max(...parsedData.map(row => row.length));
    
    Logger.info('Parsed data:', parsedData.length, 'rows,', maxColumns, 'columns');
    
    return {
      success: true,
      data: parsedData,
      columns: maxColumns,
      rows: parsedData.length,
      method: options.method
    };
    
  } catch (error) {
    Logger.error('Error previewing tableize:', error);
    return { success: false, error: error.message };
  }
}

function applyTableize(parsedData) {
  try {
    const range = SpreadsheetApp.getActiveRange();
    
    Logger.info('Applying tableize, data rows:', parsedData.length);
    
    // Save state for undo
    UndoManager.saveState(range, 'Tableize Data');
    
    // Simple apply for now
    if (parsedData && parsedData.length > 0) {
      // Find max columns
      const maxCols = Math.max(...parsedData.map(row => row.length));
      
      // Pad rows to have same columns
      const paddedData = parsedData.map(row => {
        while (row.length < maxCols) {
          row.push('');
        }
        return row;
      });
      
      // Clear current range
      range.clear();
      
      // Get sheet and apply new data
      const sheet = range.getSheet();
      const startRow = range.getRow();
      const startCol = range.getColumn();
      
      // Set the new values
      const targetRange = sheet.getRange(startRow, startCol, paddedData.length, maxCols);
      targetRange.setValues(paddedData);
      
      // Format headers if first row
      const headerRange = sheet.getRange(startRow, startCol, 1, maxCols);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#f3f3f3');
      
      return {
        success: true,
        rows: paddedData.length,
        columns: maxCols
      };
    }
    
    return { success: false, error: 'No data to apply' };
    
  } catch (error) {
    Logger.error('Error applying tableize:', error);
    return { success: false, error: error.message };
  }
}

function showTextStandardization() {
  const html = HtmlService.createHtmlOutput(`
    <div style="font-family: Arial, sans-serif; padding: 16px;">
      <h3>Text Standardization</h3>
      <p>Select your formatting options:</p>

      <label style="display: block; margin: 8px 0;">
        <input type="radio" name="case" value="original" checked> Keep Original Case
      </label>
      <label style="display: block; margin: 8px 0;">
        <input type="radio" name="case" value="title"> Title Case
      </label>
      <label style="display: block; margin: 8px 0;">
        <input type="radio" name="case" value="upper"> UPPER CASE
      </label>
      <label style="display: block; margin: 8px 0;">
        <input type="radio" name="case" value="lower"> lower case
      </label>

      <hr style="margin: 16px 0;">

      <label style="display: block; margin: 8px 0;">
        <input type="checkbox" checked> Trim whitespace
      </label>
      <label style="display: block; margin: 8px 0;">
        <input type="checkbox" checked> Remove extra spaces
      </label>
      <label style="display: block; margin: 8px 0;">
        <input type="checkbox"> Remove special characters
      </label>

      <button onclick="applyTextStandardization()"
              style="background: #1a73e8; color: white; border: none; padding: 8px 16px; border-radius: 4px; margin-top: 16px; cursor: pointer;">
        Apply Changes
      </button>
    </div>

    <script>
      function applyTextStandardization() {
        // Get selected options
        const caseOption = document.querySelector('input[name="case"]:checked').value;
        const options = {
          case: caseOption,
          trim: document.querySelectorAll('input[type="checkbox"]')[0].checked,
          removeExtraSpaces: document.querySelectorAll('input[type="checkbox"]')[1].checked,
          removeSpecialChars: document.querySelectorAll('input[type="checkbox"]')[2].checked
        };

        google.script.run
          .withSuccessHandler(showResult)
          .withFailureHandler(showError)
          .standardizeText(options);
      }

      function showResult(result) {
        if (result.success) {
          alert('Success: ' + result.message);
        } else {
          alert('Error: ' + result.error);
        }
      }

      function showError(error) {
        alert('Error: ' + error.message);
      }
    </script>
  `).setTitle('Standardize Text').setWidth(350);

  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Standardize text in selected range
 */
function standardizeText(options) {
  try {
    const range = SpreadsheetApp.getActiveRange();
    const data = range.getValues();

    if (data.length === 0) {
      return { success: false, error: 'No data selected' };
    }

    // Check usage
    const usageCheck = UsageTracker.track('operations');
    if (!usageCheck.allowed) {
      return { success: false, error: 'Usage limit reached. Please upgrade.' };
    }

    // Save state for undo
    UndoManager.saveState(range, 'Text Standardization');

    // Map UI options to cleanText options
    const cleanOptions = {
      case: options.textCase !== 'none' ? options.textCase : null,
      removeExtraSpaces: options.removeExtraSpaces,
      trim: options.trimCells !== false,
      removeLineBreaks: options.removeLineBreaks,
      fixPunctuation: options.fixPunctuation,
      removeCurrencySymbols: options.removeCurrencySymbols,
      removeSpecialChars: options.removeSpecialChars
    };

    // Process the data with enhanced messy data handling
    const standardizedData = data.map(row => {
      return row.map(cell => {
        // First handle messy data issues
        let cleaned = Utils.cleanMessyData(cell, {
          handleMissingData: options.handleMissingData,
          missingDataReplacement: options.missingDataReplacement || '',
          fixEncoding: options.fixEncoding,
          decodeHtml: options.decodeHtml,
          fixOcrErrors: options.fixOcrErrors,
          normalizeDataType: options.normalizeDataType
        });
        
        // Handle numbers separately if standardizeNumbers is checked
        if (options.standardizeNumbers) {
          // Use the new standardizeNumber function with appropriate options
          const standardized = Utils.standardizeNumber(cleaned, {
            removeCurrencySymbols: options.removeCurrencySymbols,
            decimals: options.numberDecimals !== undefined ? options.numberDecimals : 2,
            useCommas: options.useThousandSeparators !== false,
            keepOriginalFormat: options.keepOriginalNumberFormat
          });
          
          // If it was successfully standardized as a number, return it
          if (standardized !== cleaned) {
            return standardized;
          }
        }
        
        // Apply text cleaning to string values
        if (cleaned !== null && cleaned !== '') {
          return Utils.cleanText(cleaned, cleanOptions);
        }
        
        return cleaned;
      });
    });

    range.setValues(standardizedData);

    return {
      success: true,
      message: `Text standardization complete! Processed ${data.length} rows.`
    };

  } catch (error) {
    Logger.error('Error standardizing text:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Preview text standardization changes
 */
function previewStandardization(options) {
  try {
    const range = SpreadsheetApp.getActiveRange();
    const data = range.getValues();
    
    if (data.length === 0) {
      return { success: false, error: 'No data selected' };
    }
    
    // Map UI options to cleanText options
    const cleanOptions = {
      case: options.textCase !== 'none' ? options.textCase : null,
      removeExtraSpaces: options.removeExtraSpaces,
      trim: options.trimCells !== false,
      removeLineBreaks: options.removeLineBreaks,
      fixPunctuation: options.fixPunctuation,
      removeCurrencySymbols: options.removeCurrencySymbols,
      removeSpecialChars: options.removeSpecialChars
    };
    
    // Collect preview samples (max 10 changes)
    const changes = [];
    const maxSamples = 10;
    
    for (let i = 0; i < data.length && changes.length < maxSamples; i++) {
      for (let j = 0; j < data[i].length && changes.length < maxSamples; j++) {
        const original = data[i][j];
        if (original === null || original === '') continue;
        
        let processed;
        
        // Handle numbers if standardizeNumbers is checked
        if (options.standardizeNumbers) {
          // Try to standardize as number first
          const standardized = Utils.standardizeNumber(original, {
            removeCurrencySymbols: options.removeCurrencySymbols,
            decimals: options.numberDecimals !== undefined ? options.numberDecimals : 2,
            useCommas: options.useThousandSeparators !== false,
            keepOriginalFormat: options.keepOriginalNumberFormat
          });
          
          // If it was successfully standardized as a number, use that
          if (standardized !== original) {
            processed = standardized;
          } else {
            processed = Utils.cleanText(String(original), cleanOptions);
          }
        } else {
          processed = Utils.cleanText(String(original), cleanOptions);
        }
        
        if (String(original) !== processed) {
          changes.push({
            row: i + 1,
            col: j + 1,
            before: String(original),
            after: processed
          });
        }
      }
    }
    
    return {
      success: true,
      changes: changes,
      totalCells: data.length * (data[0] ? data[0].length : 0),
      affectedCells: changes.length
    };
    
  } catch (error) {
    Logger.error('Error previewing standardization:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Preview date formatting changes
 */
function previewDateFormatting(options) {
  try {
    const range = SpreadsheetApp.getActiveRange();
    const data = range.getValues();
    
    if (data.length === 0) {
      return { success: false, error: 'No data selected' };
    }
    
    const sourceFormat = options.sourceFormat || 'auto';
    const targetFormat = options.targetFormat || 'ISO';
    
    // Collect preview samples (max 10 changes)
    const changes = [];
    const maxSamples = 10;
    let totalDates = 0;
    
    for (let i = 0; i < data.length && changes.length < maxSamples; i++) {
      for (let j = 0; j < data[i].length && changes.length < maxSamples; j++) {
        const original = data[i][j];
        if (original === null || original === '') continue;
        
        // Try to parse as date
        const parsedDate = Utils.parseDate(String(original), sourceFormat);
        if (parsedDate) {
          totalDates++;
          const formatted = Utils.formatDate(parsedDate, targetFormat);
          
          if (String(original) !== formatted) {
            changes.push({
              row: i + 1,
              col: j + 1,
              cell: String.fromCharCode(65 + j) + (i + 1),
              before: String(original),
              after: formatted
            });
          }
        }
      }
    }
    
    return {
      success: true,
      changes: changes,
      totalDates: totalDates,
      totalCells: data.length * data[0].length
    };
    
  } catch (error) {
    Logger.error('Error previewing date formatting:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Apply date formatting to selected range
 */
function formatDates(options) {
  try {
    const range = SpreadsheetApp.getActiveRange();
    const data = range.getValues();
    
    if (data.length === 0) {
      return { success: false, error: 'No data selected' };
    }
    
    // Check usage
    const usageCheck = UsageTracker.track('operations');
    if (!usageCheck.allowed) {
      return { success: false, error: 'Usage limit reached. Please upgrade.' };
    }
    
    // Save state for undo
    UndoManager.saveState(range, 'Date Formatting');
    
    const sourceFormat = options.sourceFormat || 'auto';
    const targetFormat = options.targetFormat || 'ISO';
    let formattedCount = 0;
    
    // Process the data
    const formattedData = data.map(row => {
      return row.map(cell => {
        if (cell === null || cell === '') return cell;
        
        // Try to parse as date
        const parsedDate = Utils.parseDate(String(cell), sourceFormat);
        if (parsedDate) {
          formattedCount++;
          return Utils.formatDate(parsedDate, targetFormat);
        }
        
        return cell;
      });
    });
    
    range.setValues(formattedData);
    
    return {
      success: true,
      message: `Successfully formatted ${formattedCount} date values`,
      formattedCount: formattedCount
    };
    
  } catch (error) {
    Logger.error('Error formatting dates:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Show upgrade dialog
 */
function showUpgradeDialog() {
  try {
    const html = HtmlService.createTemplateFromFile('UpgradeTemplate')
      .evaluate()
      .setTitle('Upgrade CellPilot')
      .setWidth(450)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (error) {
    Logger.error('Error showing upgrade dialog:', error);
    showErrorDialog('Error', 'Failed to open upgrade panel: ' + error.message);
  }
}

/**
 * Track plan selection for analytics
 */
function trackPlanSelection(plan) {
  try {
    // Track the plan selection for analytics
    Logger.info('User selected plan:', plan);
    return { success: true };
  } catch (error) {
    Logger.error('Error tracking plan selection:', error);
    return { success: false };
  }
}

/**
 * Show error dialog
 */
function showErrorDialog(title, message) {
  SpreadsheetApp.getUi().alert(title, message, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Get current user context
 */
function getCurrentUserContext() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const range = sheet.getActiveRange();
    const values = range.getValues();
    const numCols = range.getNumColumns();
    const numRows = range.getNumRows();
    
    let displayType = '';
    let headerName = null;
    
    // Handle different selection scenarios
    if (numCols === 1) {
      // Single column selected
      const dataType = Utils.detectDataType(values);
      displayType = dataType;
      
      if (dataType.includes('(')) {
        const parts = dataType.match(/^(.+?)\s*\((.+?)\)$/);
        if (parts) {
          headerName = parts[1];
          displayType = `${parts[1]} column (${parts[2]})`;
        }
      }
    } else if (numCols <= 3) {
      // 2-3 columns selected - show each column type
      const columnTypes = [];
      for (let col = 0; col < numCols; col++) {
        const colData = [];
        for (let row = 0; row < numRows; row++) {
          colData.push([values[row][col]]);
        }
        const colType = Utils.detectDataType(colData);
        
        // Extract column header if detected
        if (colType.includes('(')) {
          const parts = colType.match(/^(.+?)\s*\((.+?)\)$/);
          if (parts) {
            columnTypes.push(`${parts[1]} (${parts[2]})`);
          } else {
            columnTypes.push(colType);
          }
        } else {
          // Add column letter for non-header columns
          const colLetter = String.fromCharCode(65 + range.getColumn() - 1 + col);
          columnTypes.push(`Col ${colLetter}: ${colType}`);
        }
      }
      displayType = columnTypes.join(', ');
    } else {
      // More than 3 columns - show range summary
      const startCol = String.fromCharCode(65 + range.getColumn() - 1);
      const endCol = String.fromCharCode(65 + range.getColumn() - 1 + numCols - 1);
      
      // Detect predominant data types
      const types = new Set();
      for (let col = 0; col < Math.min(numCols, 5); col++) {
        const colData = [];
        for (let row = 0; row < Math.min(numRows, 10); row++) {
          colData.push([values[row][col]]);
        }
        const colType = Utils.detectDataType(colData);
        const baseType = colType.includes('(') ? colType.split('(')[1].replace(')', '') : colType;
        types.add(baseType);
      }
      
      const typesList = Array.from(types).slice(0, 3).join('/');
      displayType = `Columns ${startCol}-${endCol} (${numCols} cols, mixed: ${typesList})`;
    }

    return {
      sheet: sheet,
      range: range.getA1Notation(),
      hasSelection: numRows > 1 || numCols > 1,
      dataType: displayType,
      rawDataType: displayType,
      headerName: headerName,
      rowCount: numRows,
      colCount: numCols
    };
  } catch (error) {
    Logger.error('Error getting user context:', error);
    return {
      hasSelection: false,
      dataType: 'unknown',
      rawDataType: 'unknown',
      headerName: null,
      rowCount: 0,
      colCount: 0
    };
  }
}

/**
 * Show industry template with specific category filter
 */
function showIndustryTemplate(category) {
  try {
    // Check feature access with FeatureGate
    if (!FeatureGate.enforceAccess('industry_tools')) {
      return; // FeatureGate will show upgrade dialog
    }
    
    // Track feature usage
    ApiIntegration.trackUsage('industry_templates', 'open', 1);
    
    const template = HtmlService.createTemplateFromFile('IndustryTemplatesTemplate');
    template.defaultCategory = category || 'real-estate';
    
    const html = template.evaluate()
      .setTitle(`Industry Templates - ${formatCategoryName(category)}`)
      .setWidth(320);
    
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (error) {
    showErrorDialog('Failed to load templates', error.message);
  }
}

/**
 * Format category name for display
 */
function formatCategoryName(category) {
  const names = {
    'real-estate': 'Real Estate',
    'construction': 'Construction',
    'healthcare': 'Healthcare',
    'marketing': 'Marketing Agency',
    'ecommerce': 'E-Commerce',
    'professional': 'Professional Services'
  };
  return names[category] || category;
}

// Placeholder functions for missing features  
function showDateFormatting() { 
  SpreadsheetApp.getUi().alert('Coming Soon', 'Date formatting feature will be available in the next update!', SpreadsheetApp.getUi().ButtonSet.OK); 
}

function showDateFormattingCard() { 
  return buildErrorCard('Coming Soon', 'Date formatting feature will be available in the next update!'); 
}

function showExcelMigration() {
  try {
    const template = HtmlService.createTemplateFromFile('ExcelMigrationTemplate');
    const html = template.evaluate()
      .setTitle('Excel Migration Assistant')
      .setWidth(450);
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (error) {
    Logger.error('Error showing Excel Migration:', error);
    SpreadsheetApp.getUi().alert('Error', 'Failed to open Excel Migration Assistant. Please try again.', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function showApiIntegration() {
  showUpgradeDialog();
}

function showDataValidation() {
  SpreadsheetApp.getUi().alert('Coming Soon', 'Advanced data validation feature will be available in the next update!', SpreadsheetApp.getUi().ButtonSet.OK);
}

function showBatchOperations() {
  SpreadsheetApp.getUi().alert('Coming Soon', 'Batch operations feature will be available in the next update!', SpreadsheetApp.getUi().ButtonSet.OK);
}

function showEmailAutomation() { showUpgradeDialog(); }
function showCalendarIntegration() { showUpgradeDialog(); }
function showFormulaTemplates() { showUpgradeDialog(); }
function showFormulaTemplatesCard() { return showUpgradeCard(); }
function showSettings() { 
  try {
    const html = HtmlService.createTemplateFromFile('SettingsTemplate')
      .evaluate()
      .setTitle('Settings')
      .setWidth(350);
    
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (error) {
    Logger.error('Error showing settings:', error);
    SpreadsheetApp.getUi().alert('Settings', 'Failed to open settings panel.', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}
function showHelp() { 
  try {
    const html = HtmlService.createTemplateFromFile('HelpFeedbackTemplate')
      .evaluate()
      .setTitle('Help & Documentation')
      .setWidth(450)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (error) {
    SpreadsheetApp.getUi().alert('Help', 'Visit www.cellpilot.io for documentation and support!', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Show feedback form for bug reports or feature requests
 * @param {string} type - Type of feedback ('bug' or 'feature')
 */
function showFeedback(type) {
  try {
    // For now, redirect to submitFeedback with type
    const ui = SpreadsheetApp.getUi();
    const title = type === 'bug' ? 'Report a Bug' : 'Request a Feature';
    const prompt = type === 'bug' ? 
      'Please describe the issue you encountered:' : 
      'Please describe the feature you\'d like to see:';
    
    const response = ui.prompt(title, prompt, ui.ButtonSet.OK_CANCEL);
    
    if (response.getSelectedButton() === ui.Button.OK) {
      const feedbackData = {
        type: type,
        message: response.getResponseText(),
        timestamp: new Date(),
        userEmail: 'beta-user' // Removed Session.getActiveUser() for beta testing
      };
      
      submitFeedback(feedbackData);
      ui.alert('Thank you!', 'Your feedback has been submitted.', ui.ButtonSet.OK);
    }
  } catch (error) {
    Logger.error('Error showing feedback form:', error);
    SpreadsheetApp.getUi().alert('Failed to submit feedback: ' + error.toString());
  }
}

/**
 * Submit user feedback to backend
 */
function submitFeedback(data) {
  try {
    // Add user info
    data.userEmail = 'beta-user'; // Removed Session.getActiveUser() for beta testing
    data.spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
    data.spreadsheetName = SpreadsheetApp.getActiveSpreadsheet().getName();
    
    // Store in Properties for now (in production, send to API)
    const feedbackStore = PropertiesService.getScriptProperties();
    const feedbackKey = 'feedback_' + new Date().getTime();
    feedbackStore.setProperty(feedbackKey, JSON.stringify(data));
    
    // Log for debugging
    Logger.info('Feedback submitted:', data);
    
    // In production, send to API endpoint
    // UrlFetchApp.fetch('https://api.cellpilot.io/feedback', {
    //   method: 'post',
    //   contentType: 'application/json',
    //   payload: JSON.stringify(data)
    // });
    
    // Track in analytics
    if (data.type === 'bug') {
      // Track bug report
      Logger.info('Bug report:', data.title);
    } else if (data.type === 'feature') {
      // Track feature request
      Logger.info('Feature request:', data.title);
    }
    
    return {
      success: true,
      message: 'Feedback submitted successfully'
    };
    
  } catch (error) {
    Logger.error('Error submitting feedback:', error);
    throw new Error('Failed to submit feedback');
  }
}
function showUpgradeOptions() { showUpgradeDialog(); }



/**
 * Smart Formula Assistant - outcome-focused formula discovery
 */
function showSmartFormulaAssistant() {
  try {
    const html = HtmlService.createTemplateFromFile('SmartFormulaAssistantTemplate')
      .evaluate()
      .setTitle('Smart Formula Assistant')
      .setWidth(400);
    
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (error) {
    Logger.error('Error showing smart formula assistant:', error);
    showErrorDialog('Error', 'Failed to load Smart Formula Assistant');
  }
}

function getSmartFormulaContext() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const range = sheet.getActiveRange();
    
    if (!range) {
      return { hasSelection: false };
    }
    
    const values = range.getValues();
    const context = {
      hasSelection: true,
      range: range.getA1Notation(),
      rows: range.getNumRows(),
      columns: range.getNumColumns(),
      dataType: detectDataType(values),
      sheet: sheet.getName()
    };
    
    return context;
  } catch (error) {
    Logger.error('Error getting context:', error);
    return { hasSelection: false };
  }
}

function detectDataType(values) {
  let hasNumbers = false;
  let hasText = false;
  let hasDates = false;
  
  for (let row of values) {
    for (let cell of row) {
      if (cell !== null && cell !== '') {
        if (typeof cell === 'number') hasNumbers = true;
        else if (cell instanceof Date) hasDates = true;
        else if (typeof cell === 'string') hasText = true;
      }
    }
  }
  
  if (hasDates) return 'Dates';
  if (hasNumbers && !hasText) return 'Numbers';
  if (hasText && !hasNumbers) return 'Text';
  if (hasNumbers && hasText) return 'Mixed';
  return 'Empty';
}

function insertSmartFormula(formula) {
  try {
    const cell = SpreadsheetApp.getActiveCell();
    cell.setFormula(formula);
    
    return {
      success: true,
      cell: cell.getA1Notation()
    };
  } catch (error) {
    Logger.error('Error inserting formula:', error);
    throw new Error('Failed to insert formula');
  }
}

/**
 * Apply an industry template to the current spreadsheet
 * @param {string} templateType - The type of template to apply
 * @return {Object} Result object with success status and details
 */
function applyIndustryTemplate(templateType) {
  try {
    // Use the IndustryTemplates.applyTemplate method which properly handles individual templates
    if (typeof IndustryTemplates === 'undefined') {
      return {
        success: false,
        error: 'IndustryTemplates module not available'
      };
    }
    
    // Call the applyTemplate method which handles template selection properly
    const result = IndustryTemplates.applyTemplate(templateType);
    
    return result;
    
  } catch (error) {
    Logger.error('Error applying industry template:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Preview an industry template without applying it
 * @param {string} templateType - The type of template to preview
 * @returns {Object} Template preview information
 */
function previewTemplate(templateType) {
  try {
    // Check if IndustryTemplates module is available
    if (typeof IndustryTemplates === 'undefined') {
      return {
        success: false,
        error: 'IndustryTemplates module not available'
      };
    }
    
    // Call the IndustryTemplates.previewTemplate method to create actual preview sheets
    const result = IndustryTemplates.previewTemplate(templateType);
    return result;
    
  } catch (error) {
    Logger.error('Error creating template preview:', error);
    return {
      success: false,
      error: error.message || error.toString()
    };
  }
}

/**
 * Clean up industry template preview sheets
 */
function cleanupIndustryPreviews() {
  try {
    if (typeof IndustryTemplates !== 'undefined' && IndustryTemplates.cleanupPreviewSheets) {
      return IndustryTemplates.cleanupPreviewSheets();
    }
    return { success: true };
  } catch (error) {
    Logger.error('Error cleaning up previews:', error);
    return { success: false, error: error.toString() };
  }
}

// Keep the old function for reference (can be removed later)
function previewTemplateOld(templateType) {
  try {
    // Template preview data with descriptions and sheet structures
    const templatePreviews = {
      // Real Estate Templates
      'commission-tracker': {
        name: 'Commission Tracker',
        description: 'Track real estate commissions, splits, and payments with automated calculations',
        sheets: ['Transactions', 'Agents', 'Commission Calculator', 'Dashboard'],
        keyFeatures: [
          'Automatic commission calculations based on split percentages',
          'Agent performance tracking and ranking',
          'Monthly and quarterly commission reports',
          'Outstanding payments and collection tracking'
        ],
        sampleData: 'Includes 15 sample transactions across different property types'
      },
      
      'property-manager': {
        name: 'Property Management Suite',
        description: 'Complete property management with rent tracking, maintenance, and tenant management',
        sheets: ['Properties', 'Tenants', 'Rent Roll', 'Maintenance', 'Financial Summary'],
        keyFeatures: [
          'Automated rent collection tracking',
          'Maintenance request management',
          'Tenant screening and lease tracking',
          'Property performance analytics'
        ],
        sampleData: 'Demo portfolio with 25 properties and 40 tenants'
      },
      
      'investment-analyzer': {
        name: 'Investment Property Analyzer',
        description: 'Comprehensive property investment analysis with cash flow projections',
        sheets: ['Investment Analysis', 'Cash Flow', 'Comparables', 'Reports'],
        keyFeatures: [
          'Cap rate and ROI calculations',
          '10-year cash flow projections',
          'Comparable property analysis',
          'Investment performance tracking'
        ],
        sampleData: 'Sample analysis for residential and commercial properties'
      },
      
      'lead-pipeline': {
        name: 'Lead Pipeline Manager',
        description: 'Real estate lead management with conversion tracking and follow-up automation',
        sheets: ['Leads', 'Pipeline Stages', 'Activities', 'Conversion Analytics'],
        keyFeatures: [
          'Lead source tracking and attribution',
          'Automated follow-up scheduling',
          'Conversion rate analysis by source',
          'Sales activity tracking'
        ],
        sampleData: '50+ sample leads across different sources and stages'
      },
      
      // Construction Templates
      'cost-estimator': {
        name: 'Construction Cost Estimator',
        description: 'Detailed construction cost estimation with material waste tracking',
        sheets: ['Project Overview', 'Materials', 'Labor', 'Equipment', 'Cost Summary'],
        keyFeatures: [
          'Material quantity calculations with waste factors',
          'Labor cost estimation by trade',
          'Equipment rental cost tracking',
          'Change order management'
        ],
        sampleData: 'Sample residential construction project breakdown'
      },
      
      // Healthcare Templates  
      'insurance-verifier': {
        name: 'Insurance Verification System',
        description: 'Streamline insurance verification and billing processes',
        sheets: ['Patients', 'Insurance Plans', 'Verification Log', 'Billing Dashboard'],
        keyFeatures: [
          'Insurance eligibility tracking',
          'Prior authorization management',
          'Denial tracking and appeals',
          'Revenue cycle analytics'
        ],
        sampleData: 'Sample patient data and insurance scenarios'
      },
      
      // Marketing Templates
      'campaign-dashboard': {
        name: 'Marketing Attribution Dashboard',
        description: 'Multi-channel marketing attribution with ROI tracking',
        sheets: ['Campaigns', 'Channels', 'Attribution', 'ROI Analysis'],
        keyFeatures: [
          'Cross-channel attribution modeling',
          'Campaign performance tracking',
          'Customer acquisition cost analysis',
          'Lifetime value calculations'
        ],
        sampleData: 'Sample marketing campaigns across digital channels'
      },
      
      // E-commerce Templates
      'ecommerce-inventory': {
        name: 'Multi-Channel Inventory Manager',
        description: 'Inventory synchronization across multiple sales channels',
        sheets: ['Products', 'Inventory Levels', 'Sales Channels', 'Reorder Alerts'],
        keyFeatures: [
          'Multi-marketplace inventory sync',
          'Automated reorder point calculations',
          'Sales velocity tracking',
          'Profitability analysis by product'
        ],
        sampleData: 'Sample product catalog with inventory levels'
      },
      
      // Consulting Templates
      'time-billing-tracker': {
        name: 'Time & Billing Tracker',
        description: 'Comprehensive time tracking and billing for consultants',
        sheets: ['Time Entries', 'Projects', 'Clients', 'Invoicing', 'Reports'],
        keyFeatures: [
          'Detailed time tracking by project',
          'Automated invoice generation',
          'Project profitability analysis',
          'Utilization rate tracking'
        ],
        sampleData: 'Sample professional services projects with time entries'
      }
    };
    
    const preview = templatePreviews[templateType];
    if (!preview) {
      return {
        success: false,
        error: `Preview not available for template: ${templateType}`
      };
    }
    
    return {
      success: true,
      preview: preview
    };
    
  } catch (error) {
    Logger.error('Error generating template preview:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Navigate to a specific sheet by name
 * @param {string} sheetName - Name of the sheet to navigate to
 */
/**
 * Get cross-sheet information for visual formula builder
 */
function getCrossSheetInfo() {
  return VisualFormulaBuilder.getCrossSheetInfo();
}

/**
 * Preview data from a sheet range for cross-sheet formulas
 */
function previewSheetRange(sheetName, range) {
  return VisualFormulaBuilder.previewSheetRange(sheetName, range);
}

/**
 * Get smart formula suggestions based on selected data patterns
 */
function getSmartFormulaSuggestions() {
  return VisualFormulaBuilder.getSmartFormulaSuggestions();
}

/**
 * Analyze multi-tab relationships across all sheets
 */
function analyzeMultiTabRelationships() {
  return VisualFormulaBuilder.analyzeMultiTabRelationships();
}

/**
 * Generate a detailed relationship report
 */
function generateRelationshipReport(relationshipData) {
  return VisualFormulaBuilder.generateRelationshipReport(relationshipData);
}

/**
 * Show multi-tab relationship mapper
 */
function showMultiTabRelationshipMapper() {
  try {
    const html = HtmlService.createTemplateFromFile('MultiTabRelationshipMapperTemplate')
      .evaluate()
      .setTitle('Multi-Tab Relationship Mapper')
      .setWidth(500);
    
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (error) {
    Logger.error('Error showing relationship mapper:', error);
    SpreadsheetApp.getUi().alert('Failed to open Relationship Mapper: ' + error.message);
  }
}

/**
 * Show smart formula debugger
 */
function showSmartFormulaDebugger() {
  return FormulaDebugger.showSmartFormulaDebugger();
}

/**
 * Debug the active cell formula
 */
function debugActiveFormula() {
  return FormulaDebugger.debugActiveFormula();
}

/**
 * Debug all formulas in selection
 */
function debugSelectionFormulas() {
  return FormulaDebugger.debugSelectionFormulas();
}

/**
 * Debug all formulas in the sheet
 */
function debugAllSheetFormulas() {
  return FormulaDebugger.debugAllSheetFormulas();
}

/**
 * Apply a formula fix
 */
function applyFormulaFix(cellRef, formula) {
  return FormulaDebugger.applyFormulaFix(cellRef, formula);
}

/**
 * Analyze formula dependencies
 */
function analyzeDependencies() {
  return FormulaDebugger.analyzeDependencies();
}

/**
 * Analyze formula performance
 */
function analyzeFormulaPerformance() {
  return FormulaDebugger.analyzeFormulaPerformance();
}

function navigateToSheet(sheetName) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (sheet) {
      sheet.activate();
      return true;
    }
    return false;
  } catch (error) {
    Logger.error('Error navigating to sheet:', error);
    return false;
  }
}

// Data Validation Generator proxy functions
function getCurrentSelection() {
  return DataValidationGenerator.getCurrentSelection();
}

function applyDataValidation(rule) {
  return DataValidationGenerator.applyDataValidation(rule);
}

function testDataValidation(rule, testValue) {
  return DataValidationGenerator.testDataValidation(rule, testValue);
}

function clearDataValidation(startCell, endCell) {
  return DataValidationGenerator.clearDataValidation(startCell, endCell);
}

function getExistingValidation(rangeA1) {
  return DataValidationGenerator.getExistingValidation(rangeA1);
}

function createFromTemplate(templateName, range) {
  return DataValidationGenerator.createFromTemplate(templateName, range);
}

function suggestValidation(rangeA1) {
  return DataValidationGenerator.suggestValidation(rangeA1);
}

// Conditional Formatting Wizard proxy functions
function getCurrentSelectionForFormatting() {
  return ConditionalFormattingWizard.getCurrentSelectionForFormatting();
}

function applyConditionalFormatting(rule) {
  return ConditionalFormattingWizard.applyConditionalFormatting(rule);
}

function getExistingFormattingRules(rangeA1) {
  return ConditionalFormattingWizard.getExistingFormattingRules(rangeA1);
}

function deleteFormattingRule(rangeA1, index) {
  return ConditionalFormattingWizard.deleteFormattingRule(rangeA1, index);
}

function clearFormattingRules(rangeA1) {
  return ConditionalFormattingWizard.clearFormattingRules(rangeA1);
}

function applyPresetFormatting(preset, rangeA1) {
  return ConditionalFormattingWizard.applyPresetFormatting(preset, rangeA1);
}

// ML Support Functions
function getUserLearningProfile() {
  try {
    const profile = PropertiesService.getUserProperties().getProperty('cellpilot_ml_profile');
    return profile ? JSON.parse(profile) : null;
  } catch (error) {
    Logger.error('Error loading user ML profile:', error);
    return null;
  }
}

function saveUserLearningProfile(profile) {
  try {
    const serialized = JSON.stringify(profile);
    PropertiesService.getUserProperties().setProperty('cellpilot_ml_profile', serialized);
    return { success: true };
  } catch (error) {
    Logger.error('Error saving user ML profile:', error);
    return { success: false, error: error.message };
  }
}

function trackMLFeedback(operation, prediction, userAction, metadata) {
  try {
    // Get existing feedback history
    const historyStr = PropertiesService.getUserProperties().getProperty('cellpilot_ml_feedback');
    const history = historyStr ? JSON.parse(historyStr) : [];
    
    // Add new feedback
    history.push({
      operation: operation,
      prediction: prediction,
      userAction: userAction,
      metadata: metadata,
      timestamp: new Date().toISOString()
    });
    
    // Keep only recent feedback (last 500 entries)
    if (history.length > 500) {
      history.splice(0, history.length - 500);
    }
    
    // Save back to properties
    PropertiesService.getUserProperties().setProperty('cellpilot_ml_feedback', JSON.stringify(history));
    
    // Update DataCleaner if this is duplicate detection feedback
    if (operation === 'duplicate_detection' && userAction) {
      DataCleaner.learnFromDuplicateFeedback(
        userAction === 'accept' ? [prediction] : [],
        userAction === 'reject' ? [prediction] : [],
        metadata?.threshold || 0.85
      );
    }
    
    return { success: true };
  } catch (error) {
    Logger.error('Error tracking ML feedback:', error);
    return { success: false, error: error.message };
  }
}

function getMLFeedbackHistory() {
  try {
    const historyStr = PropertiesService.getUserProperties().getProperty('cellpilot_ml_feedback');
    return historyStr ? JSON.parse(historyStr) : [];
  } catch (error) {
    Logger.error('Error getting ML feedback history:', error);
    return [];
  }
}

function enableMLFeatures() {
  try {
    // Enable ML in DataCleaner
    DataCleaner.enableML();
    
    // Set flag in user properties
    PropertiesService.getUserProperties().setProperty('cellpilot_ml_enabled', 'true');
    
    // Show success message
    SpreadsheetApp.getUi().alert('ML Features Enabled', 
      'Machine learning features have been activated. CellPilot will now learn from your actions to provide better suggestions.', 
      SpreadsheetApp.getUi().ButtonSet.OK);
    
    Logger.info('ML features enabled');
    return { success: true, message: 'ML features enabled successfully' };
  } catch (error) {
    Logger.error('Error enabling ML features:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Disable ML features
 */
function disableMLFeatures() {
  try {
    PropertiesService.getUserProperties().setProperty('cellpilot_ml_enabled', 'false');
    Logger.info('ML features disabled');
    return { success: true, message: 'ML features disabled' };
  } catch (error) {
    Logger.error('Error disabling ML features:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Load user settings from properties
 */
function loadUserSettings() {
  try {
    const props = PropertiesService.getUserProperties();
    const settings = {
      mlEnabled: props.getProperty('cellpilot_ml_enabled') === 'true',
      personalizedRecs: props.getProperty('cellpilot_personalized_recs') === 'true',
      predictiveActions: props.getProperty('cellpilot_predictive_actions') === 'true',
      defaultView: props.getProperty('cellpilot_default_view') || 'dashboard',
      autoSave: props.getProperty('cellpilot_auto_save') !== 'false',
      showTips: props.getProperty('cellpilot_show_tips') !== 'false',
      analytics: props.getProperty('cellpilot_analytics') === 'true'
    };
    return settings;
  } catch (error) {
    Logger.error('Error loading settings:', error);
    return {};
  }
}

/**
 * Save user settings to properties
 */
function saveUserSettings(settings) {
  try {
    const props = PropertiesService.getUserProperties();
    
    if (settings.mlEnabled !== undefined) {
      props.setProperty('cellpilot_ml_enabled', settings.mlEnabled.toString());
    }
    if (settings.personalizedRecs !== undefined) {
      props.setProperty('cellpilot_personalized_recs', settings.personalizedRecs.toString());
    }
    if (settings.predictiveActions !== undefined) {
      props.setProperty('cellpilot_predictive_actions', settings.predictiveActions.toString());
    }
    if (settings.defaultView !== undefined) {
      props.setProperty('cellpilot_default_view', settings.defaultView);
    }
    if (settings.autoSave !== undefined) {
      props.setProperty('cellpilot_auto_save', settings.autoSave.toString());
    }
    if (settings.showTips !== undefined) {
      props.setProperty('cellpilot_show_tips', settings.showTips.toString());
    }
    if (settings.analytics !== undefined) {
      props.setProperty('cellpilot_analytics', settings.analytics.toString());
    }
    
    return { success: true };
  } catch (error) {
    Logger.error('Error saving settings:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Get ML storage information
 */
function getMLStorageInfo() {
  try {
    const props = PropertiesService.getUserProperties();
    const allProps = props.getProperties();
    
    let mlDataSize = 0;
    for (const key in allProps) {
      if (key.startsWith('cellpilot_ml_') || key.startsWith('ml_')) {
        mlDataSize += key.length + allProps[key].length;
      }
    }
    
    return {
      used: mlDataSize,
      total: 500000 // 500KB limit for properties
    };
  } catch (error) {
    Logger.error('Error getting ML storage info:', error);
    return null;
  }
}

/**
 * Clear ML data
 */
function clearMLData() {
  try {
    const props = PropertiesService.getUserProperties();
    const allProps = props.getProperties();
    
    for (const key in allProps) {
      if (key.startsWith('cellpilot_ml_') || key.startsWith('ml_') || key.includes('_history') || key.includes('_patterns')) {
        props.deleteProperty(key);
      }
    }
    
    return { success: true };
  } catch (error) {
    Logger.error('Error clearing ML data:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Reset settings to defaults
 */
function resetSettings() {
  try {
    const props = PropertiesService.getUserProperties();
    
    // Clear all cellpilot settings
    const allProps = props.getProperties();
    for (const key in allProps) {
      if (key.startsWith('cellpilot_')) {
        props.deleteProperty(key);
      }
    }
    
    // Set defaults
    props.setProperty('cellpilot_auto_save', 'true');
    props.setProperty('cellpilot_show_tips', 'true');
    props.setProperty('cellpilot_default_view', 'dashboard');
    
    return { success: true };
  } catch (error) {
    Logger.error('Error resetting settings:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Export user settings
 */
function exportUserSettings() {
  try {
    const settings = loadUserSettings();
    const mlInfo = getMLStorageInfo();
    
    return {
      settings: settings,
      mlStorage: mlInfo,
      exportDate: new Date().toISOString(),
      version: '1.0'
    };
  } catch (error) {
    Logger.error('Error exporting settings:', error);
    return null;
  }
}

// User Preference Learning Functions
function initializeUserPreferences() {
  try {
    UserPreferenceLearning.initialize();
    return { success: true };
  } catch (error) {
    Logger.error('Error initializing user preferences:', error);
    return { success: false, error: error.message };
  }
}

function trackUserAction(action, context) {
  try {
    UserPreferenceLearning.trackAction(action, context);
    return { success: true };
  } catch (error) {
    Logger.error('Error tracking user action:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Get adaptive duplicate detection threshold
 */
function getAdaptiveDuplicateThreshold() {
  try {
    const userProperties = PropertiesService.getUserProperties();
    const thresholdHistory = JSON.parse(userProperties.getProperty('thresholdHistory') || '[]');
    
    if (thresholdHistory.length > 0) {
      const recent = thresholdHistory.slice(-10);
      const avg = recent.reduce((sum, val) => sum + val, 0) / recent.length;
      return {
        success: true,
        threshold: Math.round(avg * 100) / 100,
        confidence: Math.min(0.95, 0.5 + (thresholdHistory.length * 0.05)),
        basedOn: recent.length + ' previous selections'
      };
    }
    
    return {
      success: true,
      threshold: 0.85,
      confidence: 0.5,
      basedOn: 'default'
    };
  } catch (error) {
    Logger.error('Error getting adaptive threshold:', error);
    return {
      success: false,
      threshold: 0.85,
      error: error.message
    };
  }
}

function getUserPreferenceSummary() {
  try {
    return UserPreferenceLearning.getPreferenceSummary();
  } catch (error) {
    Logger.error('Error getting preference summary:', error);
    return null;
  }
}

function getFeatureRecommendations() {
  try {
    return UserPreferenceLearning.getFeatureRecommendations();
  } catch (error) {
    Logger.error('Error getting feature recommendations:', error);
    return [];
  }
}

function getWorkflowSuggestions() {
  try {
    return UserPreferenceLearning.getWorkflowSuggestions();
  } catch (error) {
    Logger.error('Error getting workflow suggestions:', error);
    return [];
  }
}

function getPreferredSettings(feature) {
  try {
    return UserPreferenceLearning.getPreferredSettings(feature);
  } catch (error) {
    Logger.error('Error getting preferred settings:', error);
    return {};
  }
}

function predictNextAction(currentAction) {
  try {
    return UserPreferenceLearning.predictNextAction(currentAction);
  } catch (error) {
    Logger.error('Error predicting next action:', error);
    return null;
  }
}

function shouldAutomate(action, context) {
  try {
    return UserPreferenceLearning.shouldAutomate(action, context);
  } catch (error) {
    Logger.error('Error checking automation preference:', error);
    return false;
  }
}

function learnFromError(error, action, context) {
  try {
    UserPreferenceLearning.learnFromError(error, action, context);
    return { success: true };
  } catch (error) {
    Logger.error('Error learning from error:', error);
    return { success: false, error: error.message };
  }
}

function learnFromSuccess(action, context, duration) {
  try {
    UserPreferenceLearning.learnFromSuccess(action, context, duration);
    return { success: true };
  } catch (error) {
    Logger.error('Error learning from success:', error);
    return { success: false, error: error.message };
  }
}

function exportUserPreferences() {
  try {
    return UserPreferenceLearning.exportPreferences();
  } catch (error) {
    Logger.error('Error exporting preferences:', error);
    return null;
  }
}

function resetUserPreferences() {
  try {
    UserPreferenceLearning.resetPreferences();
    return { success: true, message: 'User preferences reset successfully' };
  } catch (error) {
    Logger.error('Error resetting preferences:', error);
    return { success: false, error: error.message };
  }
}

// ML Performance Monitoring and Optimization
function getMLPerformanceStats() {
  try {
    const stats = {
      userPreferences: UserPreferenceLearning.getPreferenceSummary(),
      adaptiveThresholds: Utils.getThresholdStats(),
      validationPatterns: DataValidationGenerator.getValidationStats(),
      anomalyDetection: ConditionalFormattingWizard.getAnomalyStats(),
      totalActions: UserPreferenceLearning.behaviorHistory.length,
      mlEnabled: PropertiesService.getUserProperties().getProperty('cellpilot_ml_enabled') === 'true'
    };
    
    return stats;
  } catch (error) {
    Logger.error('Error getting ML performance stats:', error);
    return null;
  }
}

function optimizeMLPerformance() {
  try {
    // Initialize adaptive learning in Utils
    Utils.initializeAdaptiveLearning();
    
    // Initialize user preferences
    UserPreferenceLearning.initialize();
    
    // Initialize ML in various modules
    DataValidationGenerator.initializeML();
    ConditionalFormattingWizard.initializeML();
    
    return { 
      success: true, 
      message: 'ML performance optimization completed',
      stats: getMLPerformanceStats()
    };
  } catch (error) {
    Logger.error('Error optimizing ML performance:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Show Data Validation Generator interface
 */
function showDataValidationGenerator() {
  try {
    const html = HtmlService.createTemplateFromFile('DataValidationGeneratorTemplate');
    const ui = HtmlService.createHtmlOutput(html.evaluate())
      .setTitle('Data Validation Generator')
      .setWidth(350);
    
    SpreadsheetApp.getUi().showSidebar(ui);
  } catch (error) {
    showErrorDialog('Failed to load Data Validation Generator', error.message);
  }
}

/**
 * Show Conditional Formatting Wizard interface
 */
function showConditionalFormattingWizard() {
  try {
    const html = HtmlService.createTemplateFromFile('ConditionalFormattingWizardTemplate');
    const ui = HtmlService.createHtmlOutput(html.evaluate())
      .setTitle('Conditional Formatting Wizard')
      .setWidth(350);
    
    SpreadsheetApp.getUi().showSidebar(ui);
  } catch (error) {
    showErrorDialog('Failed to load Conditional Formatting Wizard', error.message);
  }
}

/**
 * Show Pivot Table Assistant interface
 */
function showPivotTableAssistant() {
  try {
    // Check feature access with FeatureGate
    if (!FeatureGate.enforceAccess('advanced_analysis')) {
      return; // FeatureGate will show upgrade dialog
    }
    
    // Track feature usage
    ApiIntegration.trackUsage('pivot_table_assistant', 'open');
    
    const html = HtmlService.createTemplateFromFile('PivotTableAssistantTemplate');
    const ui = HtmlService.createHtmlOutput(html.evaluate())
      .setTitle('Pivot Table Assistant')
      .setWidth(350);
    
    SpreadsheetApp.getUi().showSidebar(ui);
  } catch (error) {
    showErrorDialog('Failed to load Pivot Table Assistant', error.message);
  }
}

// Pivot Table Assistant Functions
function analyzePivotData(range) {
  try {
    PivotTableAssistant.initialize();
    return PivotTableAssistant.analyzeDataForPivot(range);
  } catch (error) {
    Logger.error('Error analyzing pivot data:', error);
    return { success: false, error: error.message };
  }
}

function createPivotFromSuggestion(suggestion) {
  try {
    const result = PivotTableAssistant.createPivotTable(suggestion);
    
    // Track user action for ML
    if (result.success) {
      trackUserAction('createPivotTable', {
        type: suggestion.name,
        confidence: suggestion.confidence
      });
    }
    
    return result;
  } catch (error) {
    Logger.error('Error creating pivot from suggestion:', error);
    return { success: false, error: error.message };
  }
}

function applyPivotTemplate(templateName) {
  try {
    PivotTableAssistant.initialize();
    const result = PivotTableAssistant.applyTemplate(templateName);
    
    // Track user action for ML
    if (result.success) {
      trackUserAction('applyPivotTemplate', {
        template: templateName
      });
    }
    
    return result;
  } catch (error) {
    Logger.error('Error applying pivot template:', error);
    return { success: false, error: error.message };
  }
}

function getPivotTemplates() {
  try {
    return PivotTableAssistant.getPivotTemplates();
  } catch (error) {
    Logger.error('Error getting pivot templates:', error);
    return [];
  }
}

function getPivotStats() {
  try {
    return PivotTableAssistant.getPivotStats();
  } catch (error) {
    Logger.error('Error getting pivot stats:', error);
    return null;
  }
}

/**
 * Show Data Pipeline Manager interface
 */
function showDataPipelineManager() {
  try {
    const html = HtmlService.createTemplateFromFile('DataPipelineTemplate');
    const ui = HtmlService.createHtmlOutput(html.evaluate())
      .setTitle('Data Pipeline Manager')
      .setWidth(350);
    
    SpreadsheetApp.getUi().showSidebar(ui);
  } catch (error) {
    showErrorDialog('Failed to load Data Pipeline Manager', error.message);
  }
}

// Data Pipeline Manager Functions
function importPipelineData(config) {
  try {
    DataPipelineManager.initialize();
    const result = DataPipelineManager.importData(config);
    
    // Track user action for ML
    if (result.success) {
      trackUserAction('importData', {
        type: config.type,
        rowsImported: result.rowsImported,
        source: config.url ? 'url' : 'manual'
      });
    }
    
    return result;
  } catch (error) {
    Logger.error('Error importing pipeline data:', error);
    return { success: false, error: error.message };
  }
}

function exportPipelineData(options) {
  try {
    DataPipelineManager.initialize();
    const result = DataPipelineManager.exportData(options);
    
    // Track user action for ML
    if (result.success) {
      trackUserAction('exportData', {
        format: options.format,
        rowsExported: result.rowsExported
      });
    }
    
    return result;
  } catch (error) {
    Logger.error('Error exporting pipeline data:', error);
    return { success: false, error: error.message };
  }
}

function getPipelineHistory() {
  try {
    DataPipelineManager.initialize();
    return DataPipelineManager.importHistory || [];
  } catch (error) {
    Logger.error('Error getting pipeline history:', error);
    return [];
  }
}

function getPipelineStats() {
  try {
    DataPipelineManager.initialize();
    return DataPipelineManager.getPipelineStats();
  } catch (error) {
    Logger.error('Error getting pipeline stats:', error);
    return null;
  }
}

function clearPipelineHistory() {
  try {
    // Clear import history
    PropertiesService.getUserProperties().deleteProperty('import_history');
    
    // Reinitialize
    DataPipelineManager.initialize();
    
    return { success: true };
  } catch (error) {
    Logger.error('Error clearing pipeline history:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Show Cross-Sheet Formula Builder interface
 */
function showCrossSheetFormulaBuilder() {
  try {
    const html = HtmlService.createTemplateFromFile('CrossSheetFormulaBuilderTemplate');
    const ui = HtmlService.createHtmlOutput(html.evaluate())
      .setTitle('Cross-Sheet Formula Builder')
      .setWidth(350);
    
    SpreadsheetApp.getUi().showSidebar(ui);
  } catch (error) {
    showErrorDialog('Failed to load Cross-Sheet Formula Builder', error.message);
  }
}

/**
 * Get names of all sheets in the current spreadsheet
 * Used by Cross-Sheet Formula Builder
 */
function getSheetNames() {
  try {
    const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    return sheets.map(sheet => sheet.getName());
  } catch (error) {
    Logger.error('Error getting sheet names:', error);
    return [];
  }
}

/**
 * Show Formula Performance Optimizer interface
 */
function showFormulaPerformanceOptimizer() {
  try {
    // Use Smart Formula Debugger template with performance focus
    const html = HtmlService.createTemplateFromFile('SmartFormulaDebuggerTemplate');
    const ui = HtmlService.createHtmlOutput(html.evaluate())
      .setTitle('Formula Performance Optimizer')
      .setWidth(350);
    
    SpreadsheetApp.getUi().showSidebar(ui);
  } catch (error) {
    showErrorDialog('Failed to load Formula Performance Optimizer', error.message);
  }
}

function getMLStatus() {
  try {
    const enabled = PropertiesService.getUserProperties().getProperty('cellpilot_ml_enabled') === 'true';
    const profile = getUserLearningProfile();
    const feedbackHistory = getMLFeedbackHistory();
    
    return {
      enabled: enabled,
      profile: profile,
      feedbackCount: feedbackHistory.length,
      dataCleanerStatus: DataCleaner.getMLStatus()
    };
  } catch (error) {
    Logger.error('Error getting ML status:', error);
    return { error: error.message };
  }
}