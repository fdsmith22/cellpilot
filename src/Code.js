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
  console.log('CellPilot initializing...');
  UserSettings.initialize();
  
  // Create fallback menu for development/independent installation
  createCellPilotMenu();
  
  // Track installation if first time
  if (!UserSettings.load('hasTrackedInstall', false)) {
    ApiIntegration.trackInstallation();
    UserSettings.save('hasTrackedInstall', true);
  }
  
  // Check subscription status (async)
  try {
    ApiIntegration.checkSubscription();
  } catch (error) {
    console.log('Could not check subscription:', error);
  }
  
  console.log('CellPilot ready');
}

/**
 * CardService homepage for Google Workspace Add-on
 * Called automatically when add-on is opened from sidebar
 */
function buildHomepage(e) {
  try {
    console.log('Building add-on homepage...');
    
    // Create an improved card interface that matches our HTML sidebar styling
    const card = CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader()
        .setTitle('CellPilot')
        .setSubtitle('Smart Spreadsheet Assistant')
        .setImageStyle(CardService.ImageStyle.SQUARE)
        .setImageUrl('https://www.gstatic.com/images/branding/product/2x/sheets_48dp.png'))
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
          .setText('Remove Duplicates')
          .setBottomLabel('Clean duplicate entries from your data')
          .setOnClickAction(CardService.newAction()
            .setFunctionName('launchDuplicatesFromAddon')))
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
    console.error('Error building homepage:', error);
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

function launchDuplicatesFromAddon() {
  showDuplicateRemoval();
  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification()
      .setText('Opening Data Cleaning...'))
    .build();
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
      .addSubMenu(ui.createMenu('Data Cleaning')
        .addItem('Remove Duplicates', prefix + 'showDuplicateRemoval')
        .addItem('Standardize Text', prefix + 'showTextStandardization')
        .addItem('Fix Dates', prefix + 'showDateFormatting'))
      .addSubMenu(ui.createMenu('Formula Builder')
        .addItem('Natural Language Builder', prefix + 'showFormulaBuilder')
        .addItem('Formula Templates', prefix + 'showFormulaTemplates'))
      .addSeparator()
      .addItem('Settings', prefix + 'showSettings')
      .addItem('Help', prefix + 'showHelp')
      .addToUi();
  } catch (error) {
    console.error('Error creating menu:', error);
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
      .setImageUrl('https://developers.google.com/apps-script/images/logo.png'));

  // Context info section
  if (context && context.hasSelection) {
    const contextSection = CardService.newCardSection()
      .setHeader('Current Selection')
      .addWidget(CardService.newTextParagraph()
        .setText(`Selected: ${context.rowCount} rows × ${context.colCount} columns\nData type: ${context.dataType}`));
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
          .setUrl('https://cellpilot.co.uk')
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
        border-radius: 10px;
        padding: 14px 10px;
        cursor: pointer;
        transition: var(--transition-base);
        text-align: center;
        display: flex;
        flex-direction: column;
        align-items: center;
        gap: 8px;
      }
      
      .quick-action-btn:hover {
        border-color: var(--primary-500);
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.06);
        transform: translateY(-1px);
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
        font-size: 12px;
        font-weight: 600;
        color: var(--gray-700);
        line-height: 1.2;
      }
      
      /* Dropdown sections */
      .dropdown-section {
        margin-bottom: 12px;
      }
      
      .dropdown-header {
        background: white;
        border: 1px solid var(--gray-200);
        border-radius: 10px;
        padding: 14px 16px;
        cursor: pointer;
        transition: var(--transition-base);
        display: flex;
        align-items: center;
        justify-content: space-between;
      }
      
      .dropdown-header:hover {
        border-color: var(--primary-500);
        background: var(--gray-50);
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
    </style>
    
    <div class="nav-header">
      <div class="nav-title">
        <h2>CellPilot</h2>
        <p>Smart Spreadsheet Assistant</p>
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
      <? if (!context.hasSelection) { ?>
        <div class="alert alert-warning">
          Select cells in your spreadsheet to enable all features
        </div>
      <? } else { ?>
        <div class="stats-grid">
          <div class="stat-card">
            <div class="stat-value"><?= context.rowCount ?></div>
            <div class="stat-label">Rows</div>
          </div>
          <div class="stat-card">
            <div class="stat-value"><?= context.colCount ?></div>
            <div class="stat-label">Columns</div>
          </div>
          <div class="stat-card">
            <div class="stat-value"><?= context.dataType ?></div>
            <div class="stat-label">Type</div>
          </div>
        </div>
      <? } ?>
      
      <!-- Quick Actions Grid -->
      <div style="margin-bottom: 16px;">
        <h3 style="font-size: 12px; font-weight: 600; color: var(--gray-600); text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 10px;">Quick Actions</h3>
        <div class="quick-actions-grid">
          <div class="quick-action-btn" onclick="google.script.run.showTableize()">
            <div class="quick-action-icon" style="background: linear-gradient(135deg, #fef3c7, #fde68a); color: #92400e;">TZ</div>
            <div class="quick-action-label">Tableize Data</div>
          </div>
          
          <div class="quick-action-btn" onclick="google.script.run.showDuplicateRemoval()">
            <div class="quick-action-icon" style="background: linear-gradient(135deg, #eff6ff, #dbeafe); color: #2563eb;">DC</div>
            <div class="quick-action-label">Remove Duplicates</div>
          </div>
          
          <div class="quick-action-btn" onclick="google.script.run.showFormulaBuilder()">
            <div class="quick-action-icon" style="background: linear-gradient(135deg, #f0fdf4, #d1fae5); color: #16a34a;">FB</div>
            <div class="quick-action-label">Formula Builder</div>
          </div>
          
          <div class="quick-action-btn" onclick="google.script.run.showAutomation()">
            <div class="quick-action-icon" style="background: linear-gradient(135deg, #faf5ff, #e9d5ff); color: #7c3aed;">AI</div>
            <div class="quick-action-label">Automation</div>
          </div>
        </div>
      </div>
      
      <!-- Industry Templates Dropdown -->
      <div class="dropdown-section">
        <div class="dropdown-header" onclick="toggleDropdown('industry')">
          <div class="dropdown-title">
            <div class="dropdown-icon" style="background: linear-gradient(135deg, #fef3c7, #fed7aa); color: #92400e;">IT</div>
            <div class="dropdown-text">
              <div class="dropdown-label">Industry Templates</div>
              <div class="dropdown-sublabel">Pre-built solutions saving 20+ hrs/week</div>
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
            <span class="dropdown-item-badge">4 templates</span>
          </div>
          <div class="dropdown-item" onclick="google.script.run.showIndustryTemplate('ecommerce')">
            E-Commerce
            <span class="dropdown-item-badge">3 templates</span>
          </div>
          <div class="dropdown-item" onclick="google.script.run.showIndustryTemplate('consulting')">
            Consulting
            <span class="dropdown-item-badge">3 templates</span>
          </div>
        </div>
      </div>
      
      <!-- Advanced Tools Dropdown -->
      <div class="dropdown-section">
        <div class="dropdown-header" onclick="toggleDropdown('advanced')">
          <div class="dropdown-title">
            <div class="dropdown-icon" style="background: linear-gradient(135deg, #e0e7ff, #c7d2fe); color: #4f46e5;">AT</div>
            <div class="dropdown-text">
              <div class="dropdown-label">Advanced Tools</div>
              <div class="dropdown-sublabel">Power features for complex tasks</div>
            </div>
          </div>
          <div class="dropdown-arrow">→</div>
        </div>
        <div id="advanced-dropdown" class="dropdown-content">
          <div class="dropdown-item" onclick="google.script.run.showExcelMigration()">
            Excel Migration Assistant
          </div>
          <div class="dropdown-item" onclick="google.script.run.showApiIntegration()">
            API Integration
          </div>
          <div class="dropdown-item" onclick="google.script.run.showDataValidation()">
            Advanced Data Validation
          </div>
          <div class="dropdown-item" onclick="google.script.run.showBatchOperations()">
            Batch Operations
          </div>
        </div>
      </div>
      
      <!-- Upgrade Card -->
      <div class="card" style="background: linear-gradient(135deg, #faf5ff, #f3e8ff); border-color: #e9d5ff; margin-top: 16px;">
        <div class="card-title">Upgrade to Pro</div>
        <div class="card-description">Unlock unlimited operations and advanced features</div>
        <button class="btn btn-primary btn-block" onclick="google.script.run.showUpgradeOptions()">
          View Premium Plans
        </button>
      </div>
      
      <div class="footer">
        <a href="#" onclick="google.script.run.showSettings()">Settings</a> • 
        <a href="#" onclick="google.script.run.showHelp()">Help</a> • 
        <a href="https://cellpilot.co.uk" target="_blank">Website</a>
      </div>
    </div>
    
    <script>
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
 * Show duplicate removal interface
 */
function showDuplicateRemoval() {
  try {
    const range = SpreadsheetApp.getActiveRange();
    const data = range.getValues();

    if (data.length === 0) {
      showErrorDialog('No Data Selected', 'Please select a range with data to remove duplicates.');
      return;
    }

    const html = HtmlService.createTemplateFromFile('DuplicateRemovalTemplate')
      .evaluate()
      .setTitle('Remove Duplicates')
      .setWidth(400)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

    SpreadsheetApp.getUi().showSidebar(html);

  } catch (error) {
    showErrorDialog('Error', 'Failed to load duplicate removal: ' + error.message);
  }
}

/**
 * Actually remove duplicates (called from the HTML interface)
 */
function removeDuplicatesProcess(options) {
  try {
    const range = SpreadsheetApp.getActiveRange();
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

    const threshold = options.threshold || 0.85;
    const caseSensitive = options.caseSensitive || false;

    const duplicates = DataCleaner.findDuplicates(data, { threshold, caseSensitive });
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
    console.error('Error removing duplicates:', error);
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
    console.error('Error previewing duplicates:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Show formula builder interface
 */
function showFormulaBuilder() {
  try {
    // Check tier access
    const userTier = UserSettings.load('userTier', 'free');
    if (!Config.FEATURE_ACCESS.formula_builder.includes(userTier)) {
      showUpgradeDialog();
      return;
    }

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
    console.error('Error generating formula:', error);
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
    console.error('Error inserting formula:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Show automation interface
 */
function showAutomation() {
  try {
    const html = HtmlService.createTemplateFromFile('AutomationTemplate')
      .evaluate()
      .setTitle('Smart Automation')
      .setWidth(450)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (error) {
    console.error('Error showing automation interface:', error);
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
    console.error('Error analyzing formats:', error);
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

function analyzeDataForTableize() {
  // Super simple version to test if function is even being called
  console.log('analyzeDataForTableize called');
  
  try {
    // Just return a basic response immediately
    return {
      hasData: true,
      rowCount: 3,
      suggestedDelimiter: 'space',
      delimiters: {
        space: { char: ' ', count: 2, consistent: true },
        comma: { char: ',', count: 0, consistent: true },
        tab: { char: '\t', count: 0, consistent: true },
        pipe: { char: '|', count: 0, consistent: true },
        semicolon: { char: ';', count: 0, consistent: true },
        doubleSpace: { char: '  ', count: 0, consistent: true }
      },
      patterns: { csvLike: false, keyValue: false },
      hasHeaders: true,
      firstRowSample: 'Test data',
      estimatedColumns: 3
    };
    
  } catch (error) {
    console.error('Error in analyzeDataForTableize:', error);
    return { hasData: false, error: 'Function error: ' + error.toString() };
  }
}

function previewTableize(options) {
  try {
    const range = SpreadsheetApp.getActiveRange();
    const data = range.getValues();
    
    console.log('Preview tableize with options:', options);
    
    let parsedData = [];
    
    if (options.method === 'smart') {
      // Use smart parser for intelligent grouping
      console.log('Using smart parser');
      parsedData = SmartTableParser.smartParse(data, options);
      
      // Add suggested headers if first row isn't headers
      if (!options.hasHeaders && parsedData.length > 0) {
        // Detect likely column types and add headers
        const firstRow = parsedData[0];
        const headers = [];
        
        if (firstRow.length >= 3) {
          // Likely pattern: Name, Region, Sales
          headers.push('Name', 'Region', 'Sales');
        } else if (firstRow.length === 2) {
          headers.push('Name', 'Value');
        } else {
          // Generic headers
          for (let i = 0; i < firstRow.length; i++) {
            headers.push('Column ' + (i + 1));
          }
        }
        
        // Insert headers at beginning
        parsedData.unshift(headers);
      }
      
    } else {
      // Original simple parsing
      for (let i = 0; i < data.length; i++) {
        const cellValue = String(data[i][0] || '');
        
        if (cellValue.trim()) {
          let columns;
          
          // Simple delimiter splitting
          if (options.delimiter === 'space') {
            columns = cellValue.split(/\s+/).filter(c => c);
          } else if (options.delimiter === 'comma') {
            columns = cellValue.split(',').map(c => c.trim());
          } else if (options.delimiter === 'tab') {
            columns = cellValue.split('\t').map(c => c.trim());
          } else {
            columns = cellValue.split(/\s+/).filter(c => c);
          }
          
          parsedData.push(columns);
        }
      }
    }
    
    const maxColumns = Math.max(...parsedData.map(row => row.length));
    
    console.log('Parsed data:', parsedData.length, 'rows,', maxColumns, 'columns');
    
    return {
      success: true,
      data: parsedData,
      columns: maxColumns,
      rows: parsedData.length,
      method: options.method
    };
    
  } catch (error) {
    console.error('Error previewing tableize:', error);
    return { success: false, error: error.message };
  }
}

function applyTableize(parsedData) {
  try {
    const range = SpreadsheetApp.getActiveRange();
    
    console.log('Applying tableize, data rows:', parsedData.length);
    
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
    console.error('Error applying tableize:', error);
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

    const standardizedData = DataCleaner.standardizeTextData(data, options);
    range.setValues(standardizedData);

    return {
      success: true,
      message: `Text standardization complete! Processed ${data.length} rows.`
    };

  } catch (error) {
    console.error('Error standardizing text:', error);
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
    console.error('Error showing upgrade dialog:', error);
    showErrorDialog('Error', 'Failed to open upgrade panel: ' + error.message);
  }
}

/**
 * Track plan selection for analytics
 */
function trackPlanSelection(plan) {
  try {
    // Track the plan selection for analytics
    console.log('User selected plan:', plan);
    return { success: true };
  } catch (error) {
    console.error('Error tracking plan selection:', error);
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

    return {
      sheet: sheet,
      range: range,
      hasSelection: range.getNumRows() > 1 || range.getNumColumns() > 1,
      dataType: Utils.detectDataType(range.getValues()),
      rowCount: range.getNumRows(),
      colCount: range.getNumColumns()
    };
  } catch (error) {
    console.error('Error getting user context:', error);
    return {
      hasSelection: false,
      dataType: 'unknown',
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
    'consulting': 'Consulting'
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
  showAutomation(); // Redirect to automation tab which includes Excel Migration
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
  SpreadsheetApp.getUi().alert('Settings', 'Settings panel coming soon!', SpreadsheetApp.getUi().ButtonSet.OK); 
}
function showHelp() { 
  SpreadsheetApp.getUi().alert('Help', 'Visit cellpilot.co.uk for documentation and support!', SpreadsheetApp.getUi().ButtonSet.OK); 
}
function showUpgradeOptions() { showUpgradeDialog(); }