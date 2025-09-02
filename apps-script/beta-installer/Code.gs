/**
 * CellPilot Beta Installer Web App
 * =================================
 * This web app helps beta users install CellPilot
 */

/**
 * Handle GET requests - show installer page
 */
function doGet(e) {
  // Try multiple methods to get the user email
  let userEmail = Session.getActiveUser().getEmail();
  
  // If email is empty, try effective user
  if (!userEmail) {
    userEmail = Session.getEffectiveUser().getEmail();
  }
  
  console.log('User email from session:', userEmail);
  console.log('Active user:', Session.getActiveUser().getEmail());
  console.log('Effective user:', Session.getEffectiveUser().getEmail());
  
  // Check if user has beta access
  const hasBeta = checkBetaAccess(userEmail);
  console.log('Beta access check result:', hasBeta);
  
  if (!hasBeta) {
    return HtmlService.createHtmlOutput(`
      <div style="font-family: Arial, sans-serif; padding: 40px; max-width: 600px; margin: 0 auto;">
        <div style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 30px; border-radius: 16px 16px 0 0; text-align: center;">
          <h1 style="margin: 0;">CellPilot Beta</h1>
        </div>
        <div style="background: white; padding: 30px; border: 1px solid #e5e7eb; border-radius: 0 0 16px 16px;">
          <h2 style="color: #1a202c; margin-top: 0;">Beta Access Required</h2>
          <p style="color: #4a5568; line-height: 1.6;">
            To install CellPilot, you need to:
          </p>
          <ol style="color: #4a5568; line-height: 2;">
            <li>Sign up at <a href="https://www.cellpilot.io" target="_blank" style="color: #667eea;">www.cellpilot.io</a></li>
            <li>Go to your dashboard</li>
            <li>Click "Activate Beta Access"</li>
            <li>Return here to install</li>
          </ol>
          <p style="color: #718096; font-size: 14px; margin-top: 20px;">
            <strong>Your email:</strong> ${userEmail || 'Not signed in'}<br>
            <small>Make sure you're signed in with the same Google account you used to sign up.</small>
          </p>
          <div style="text-align: center; margin-top: 30px;">
            <a href="https://www.cellpilot.io" target="_blank" style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 12px 24px; text-decoration: none; border-radius: 8px; display: inline-block;">
              Sign Up for Beta Access
            </a>
          </div>
        </div>
      </div>
    `).setTitle('CellPilot Beta Access');
  }
  
  // Return the installer page
  return HtmlService.createTemplateFromFile('installer')
    .evaluate()
    .setTitle('Install CellPilot Beta')
    .setWidth(800)
    .setHeight(600);
}

/**
 * Check if user has beta access
 */
function checkBetaAccess(email) {
  if (!email) return false;
  
  try {
    // Check against CellPilot API
    const response = UrlFetchApp.fetch('https://www.cellpilot.io/api/check-beta', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({ email: email }),
      muteHttpExceptions: true // Get response even on error
    });
    
    const responseText = response.getContentText();
    console.log('API Response:', responseText);
    console.log('Response Code:', response.getResponseCode());
    
    if (response.getResponseCode() !== 200) {
      console.error('API returned error code:', response.getResponseCode());
      return false;
    }
    
    const result = JSON.parse(responseText);
    console.log('Parsed result:', result);
    return result.hasBeta === true;
    
  } catch (e) {
    console.error('Error checking beta access:', e.toString());
    // If API is down, deny access for security
    return false;
  }
}

/**
 * Get the installation code
 */
function getInstallationCode() {
  const scriptId = '1EZDAGoLY8UEMdbfKTZO-AQ7pkiPe-n-zrz3Rw0ec6VBBH5MdC43Avx0O';
  const version = '11'; // Updated to v11 with proper exports
  
  // Return the full installation code
  const code = getFullBetaInstallerCode(scriptId, version);
  
  return {
    scriptId: scriptId,
    version: version,
    code: code,
    success: true
  };
}

/**
 * Get the full beta installer code
 */
function getFullBetaInstallerCode(scriptId, version) {
  // Return the complete proxy code that users will paste
  return `/**
 * CellPilot Beta Installation Script
 * ===================================
 * Version: 1.0.0-beta
 * 
 * IMPORTANT: After pasting this code, you MUST add the CellPilot library:
 * 1. Click Libraries (+) in the left sidebar
 * 2. Script ID: ${scriptId}
 * 3. Version: ${version}
 * 4. Identifier: CellPilot
 * 5. Click Add
 * 6. SAVE THE PROJECT (Ctrl+S or Cmd+S)
 * 7. CLOSE AND REOPEN THE SPREADSHEET
 */

// ============================================
// INITIALIZATION & MENU FUNCTIONS
// ============================================

function onOpen(e) {
  try {
    // Create a simple menu first
    const ui = SpreadsheetApp.getUi();
    
    // Try to access CellPilot
    if (typeof CellPilot !== 'undefined' && CellPilot && CellPilot.onOpen) {
      // Library is loaded, call it
      CellPilot.onOpen(e);
    } else {
      // Library not found, create setup menu
      ui.createMenu('⚠️ CellPilot Setup')
        .addItem('Initialize CellPilot', 'initializeCellPilot')
        .addItem('Test Library Connection', 'testLibrary')
        .addSeparator()
        .addItem('Show Instructions', 'showInstructions')
        .addToUi();
        
      // Don't show alert on every open, just create the menu
    }
  } catch (error) {
    // If there's an error, create a setup menu
    try {
      SpreadsheetApp.getUi()
        .createMenu('⚠️ CellPilot Setup')
        .addItem('Initialize CellPilot', 'initializeCellPilot')
        .addItem('Test Library Connection', 'testLibrary')
        .addItem('Show Error Details', 'showError')
        .addToUi();
    } catch (e) {
      // Even menu creation failed, do nothing
    }
  }
}

function onInstall(e) {
  onOpen(e);
}

// Initialize CellPilot - this helps trigger authorization
function initializeCellPilot() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    if (typeof CellPilot === 'undefined') {
      ui.alert(
        'Library Not Found',
        'Please add the CellPilot library first:\\n\\n' +
        '1. Click Libraries (+)\\n' +
        '2. Add Script ID: 1EZDAGoLY8UEMdbfKTZO-AQ7pkiPe-n-zrz3Rw0ec6VBBH5MdC43Avx0O\\n' +
        '3. Set Identifier to: CellPilot\\n' +
        '4. Save project\\n' +
        '5. Close and reopen this sheet',
        ui.ButtonSet.OK
      );
      return;
    }
    
    // Try to call a library function to trigger auth
    CellPilot.getCurrentUserContext();
    
    ui.alert(
      'Success!',
      'CellPilot is initialized!\\n\\nPlease close and reopen this spreadsheet to see the CellPilot menu.',
      ui.ButtonSet.OK
    );
  } catch (e) {
    ui.alert(
      'Authorization Needed',
      'Please authorize CellPilot to access your spreadsheets.\\n\\nAfter authorizing, close and reopen this sheet.',
      ui.ButtonSet.OK
    );
  }
}

// Show instructions
function showInstructions() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'CellPilot Installation Instructions',
    '1. Click Libraries (+) in Apps Script\\n' +
    '2. Script ID: 1EZDAGoLY8UEMdbfKTZO-AQ7pkiPe-n-zrz3Rw0ec6VBBH5MdC43Avx0O\\n' +
    '3. Identifier MUST be: CellPilot\\n' +
    '4. Version: 11 or HEAD\\n' +
    '5. Click Add\\n' +
    '6. Save project (Ctrl+S)\\n' +
    '7. CLOSE and REOPEN this spreadsheet\\n\\n' +
    'If still not working:\\n' +
    '- Click "Initialize CellPilot" in the Setup menu\\n' +
    '- Authorize when prompted\\n' +
    '- Close and reopen sheet',
    ui.ButtonSet.OK
  );
}

// Test function to verify library is installed
function testLibrary() {
  const ui = SpreadsheetApp.getUi();
  try {
    if (typeof CellPilot === 'undefined') {
      ui.alert('Test Result', 'Library not found. Please add it with identifier "CellPilot"', ui.ButtonSet.OK);
      return;
    }
    const version = CellPilot.getVersion ? CellPilot.getVersion() : 'Unknown';
    ui.alert('Test Result', 'CellPilot library is installed! Version: ' + version, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Test Result', 'Error: ' + e.toString(), ui.ButtonSet.OK);
  }
}

// Show error details
function showError() {
  const ui = SpreadsheetApp.getUi();
  try {
    if (typeof CellPilot === 'undefined') {
      ui.alert('Error Details', 'CellPilot is undefined. Library not properly linked.', ui.ButtonSet.OK);
    } else {
      ui.alert('Error Details', 'CellPilot is defined but there may be an authorization issue.', ui.ButtonSet.OK);
    }
  } catch (e) {
    ui.alert('Error Details', e.toString(), ui.ButtonSet.OK);
  }
}

// Test CellPilot connection
function testCellPilotConnection() { return CellPilot.testCellPilotConnection(); }

// Reset beta notification
function resetBetaNotification() { return CellPilot.resetBetaNotification(); }

// ============================================
// MAIN SIDEBAR & NAVIGATION
// ============================================

function showCellPilotSidebar() { return CellPilot.showCellPilotSidebar(); }
function include(filename) { return CellPilot.include(filename); }
function createMainSidebarHtml(context) { return CellPilot.createMainSidebarHtml(context); }
function getCurrentUserContext() { return CellPilot.getCurrentUserContext(); }

// ============================================
// DATA CLEANING FEATURES
// ============================================

function showDataCleaning() { return CellPilot.showDataCleaning(); }
function showDuplicateRemoval() { return CellPilot.showDuplicateRemoval(); }
function removeDuplicatesProcess(options) { return CellPilot.removeDuplicatesProcess(options); }
function processLargeDuplicateRemoval(range, options) { return CellPilot.processLargeDuplicateRemoval(range, options); }
function previewDuplicates(options) { return CellPilot.previewDuplicates(options); }
function showTextStandardization() { return CellPilot.showTextStandardization(); }
function standardizeText(options) { return CellPilot.standardizeText(options); }
function previewStandardization(options) { return CellPilot.previewStandardization(options); }

// ============================================
// TABLEIZE FEATURE (DATA STRUCTURING)
// ============================================

function showTableize() { return CellPilot.showTableize(); }
function hasDataSelection() { return CellPilot.hasDataSelection(); }
function analyzeDataForTableize() { return CellPilot.analyzeDataForTableize(); }
function previewTableize(options) { return CellPilot.previewTableize(options); }
function applyTableize(parsedData) { return CellPilot.applyTableize(parsedData); }

// ============================================
// SMART AUTOMATION FEATURES
// ============================================

function showAutomation() { return CellPilot.showAutomation(); }
function analyzeFormats(options) { return CellPilot.analyzeFormats(options); }
function scanForExcelIssues() { return CellPilot.scanForExcelIssues(); }
function fixExcelIssues(options) { return CellPilot.fixExcelIssues(options); }

// ============================================
// FORMULA BUILDER FEATURES
// ============================================

function showFormulaBuilder() { return CellPilot.showFormulaBuilder(); }
function generateFormulaFromDescription(description) { return CellPilot.generateFormulaFromDescription(description); }
function checkFormulaBuilderAccess() { return CellPilot.checkFormulaBuilderAccess(); }
function insertFormulaIntoCell(formulaText) { return CellPilot.insertFormulaIntoCell(formulaText); }
function showFormulaTemplates() { return CellPilot.showFormulaTemplates(); }

// ============================================
// UNDO/REDO SYSTEM
// ============================================

function getUndoInfo() { return CellPilot.getUndoInfo(); }
function performUndo() { return CellPilot.performUndo(); }

// ============================================
// PREMIUM FEATURES (SHOW UPGRADE)
// ============================================

function showDateFormatting() { return CellPilot.showDateFormatting(); }
function previewDateFormatting(options) { return CellPilot.previewDateFormatting(options); }
function formatDates(options) { return CellPilot.formatDates(options); }
function showEmailAutomation() { return CellPilot.showEmailAutomation(); }
function showCalendarIntegration() { return CellPilot.showCalendarIntegration(); }

// ============================================
// ADVANCED DATA RESTRUCTURING
// ============================================

function showAdvancedRestructuring() { return CellPilot.showAdvancedRestructuring(); }
function analyzeDataStructure() { return CellPilot.analyzeDataStructure(); }
function processSectionConfiguration(config) { return CellPilot.processSectionConfiguration(config); }
function processColumnConfiguration(config) { return CellPilot.processColumnConfiguration(config); }
function applyRestructuredData(config) { return CellPilot.applyRestructuredData(config); }
function clearRestructuringSession() { return CellPilot.clearRestructuringSession(); }
function showToast(message, title, timeout) { return CellPilot.showToast(message, title, timeout); }

// ============================================
// SETTINGS & HELP
// ============================================

function showSettings() { return CellPilot.showSettings(); }
function loadUserSettings() { return CellPilot.loadUserSettings(); }
function saveUserSettings(settings) { return CellPilot.saveUserSettings(settings); }
function getMLStorageInfo() { return CellPilot.getMLStorageInfo(); }
function clearMLData() { return CellPilot.clearMLData(); }
function resetSettings() { return CellPilot.resetSettings(); }
function exportUserSettings() { return CellPilot.exportUserSettings(); }
function disableMLFeatures() { return CellPilot.disableMLFeatures(); }
function showHelp() { return CellPilot.showHelp(); }
function showSmartFormulaAssistant() { return CellPilot.showSmartFormulaAssistant(); }
function showMultiTabRelationshipMapper() { return CellPilot.showMultiTabRelationshipMapper(); }
function showSmartFormulaDebugger() { return CellPilot.showSmartFormulaDebugger(); }
function showDataValidationGenerator() { return CellPilot.showDataValidationGenerator(); }
function showConditionalFormattingWizard() { return CellPilot.showConditionalFormattingWizard(); }
function showPivotTableAssistant() { return CellPilot.showPivotTableAssistant(); }
function showDataPipelineManager() { return CellPilot.showDataPipelineManager(); }
function showCrossSheetFormulaBuilder() { return CellPilot.showCrossSheetFormulaBuilder(); }
function showFormulaPerformanceOptimizer() { return CellPilot.showFormulaPerformanceOptimizer(); }
function showIndustryTemplate(category) { return CellPilot.showIndustryTemplate(category); }
function showExcelMigration() { return CellPilot.showExcelMigration(); }
function showApiIntegration() { return CellPilot.showApiIntegration(); }
function showDataValidation() { return CellPilot.showDataValidation(); }
function showBatchOperations() { return CellPilot.showBatchOperations(); }
function showUpgradeDialog() { return CellPilot.showUpgradeDialog(); }
function showErrorDialog(title, message) { return CellPilot.showErrorDialog(title, message); }
function showFeedback(type) { return CellPilot.showFeedback(type); }
function showUpgradeOptions() { return CellPilot.showUpgradeOptions(); }
function trackPlanSelection(plan) { return CellPilot.trackPlanSelection(plan); }

// ============================================
// GOOGLE WORKSPACE ADD-ON FUNCTIONS
// ============================================

function buildHomepage(e) { return CellPilot.buildHomepage(e); }
function launchHtmlSidebarFromAddon() { return CellPilot.launchHtmlSidebarFromAddon(); }
function launchTableizeFromAddon() { return CellPilot.launchTableizeFromAddon(); }
function launchDuplicatesFromAddon() { return CellPilot.launchDuplicatesFromAddon(); }
function launchFormulasFromAddon() { return CellPilot.launchFormulasFromAddon(); }
function formatCategoryName(category) { return CellPilot.formatCategoryName(category); }

// ============================================
// VISUAL FORMULA BUILDER
// ============================================

function showVisualFormulaBuilder() { return CellPilot.showVisualFormulaBuilder(); }
function suggestFormulaBasedOnData() { return CellPilot.suggestFormulaBasedOnData(); }
function validateVisualFormula(formula) { return CellPilot.validateVisualFormula(formula); }
function optimizeVisualFormula(formula) { return CellPilot.optimizeVisualFormula(formula); }
function testVisualFormula(formula) { return CellPilot.testVisualFormula(formula); }
function insertVisualFormula(formula) { return CellPilot.insertVisualFormula(formula); }
function submitFeedback(data) { return CellPilot.submitFeedback(data); }

// ============================================
// SMART FORMULA ASSISTANT
// ============================================

function getSmartFormulaContext() { return CellPilot.getSmartFormulaContext(); }
function insertSmartFormula(formula) { return CellPilot.insertSmartFormula(formula); }

// ============================================
// CARD SERVICE FUNCTIONS (FOR ADD-ON)
// ============================================

function createCellPilotMenu() { return CellPilot.createCellPilotMenu(); }
function buildMainCard(context) { return CellPilot.buildMainCard(context); }
function showDuplicateRemovalCard() { return CellPilot.showDuplicateRemovalCard(); }
function showTextStandardizationCard() { return CellPilot.showTextStandardizationCard(); }
function showFormulaBuilderCard() { return CellPilot.showFormulaBuilderCard(); }
function showDateFormattingCard() { return CellPilot.showDateFormattingCard(); }
function showFormulaTemplatesCard() { return CellPilot.showFormulaTemplatesCard(); }
function showUpgradeCard() { return CellPilot.showUpgradeCard(); }

// ============================================
// CARD ACTION HANDLERS
// ============================================

function processDuplicateRemoval(e) { return CellPilot.processDuplicateRemoval(e); }
function processTextStandardization(e) { return CellPilot.processTextStandardization(e); }
function processFormulaGeneration(e) { return CellPilot.processFormulaGeneration(e); }
function insertGeneratedFormula(e) { return CellPilot.insertGeneratedFormula(e); }

// ============================================
// CARD BUILDERS FOR MESSAGES
// ============================================

function buildErrorCard(title, message) { return CellPilot.buildErrorCard(title, message); }
function buildSuccessCard(title, message) { return CellPilot.buildSuccessCard(title, message); }
function buildFormulaResultCard(result, description) { return CellPilot.buildFormulaResultCard(result, description); }

// ============================================
// INDUSTRY TEMPLATES
// ============================================

function previewIndustryTemplate(templateType) { return CellPilot.previewIndustryTemplate(templateType); }
function applyIndustryTemplate(templateType) { return CellPilot.applyIndustryTemplate(templateType); }
function cleanupPreviewSheets() { return CellPilot.cleanupPreviewSheets(); }
function navigateToSheet(sheetName) { return CellPilot.navigateToSheet(sheetName); }
function previewTemplate(templateType) { return CellPilot.previewTemplate(templateType); }
function cleanupIndustryPreviews() { return CellPilot.cleanupIndustryPreviews(); }

// ============================================
// CROSS-SHEET & ANALYSIS FUNCTIONS
// ============================================

function getCrossSheetInfo() { return CellPilot.getCrossSheetInfo(); }
function previewSheetRange(sheetName, range) { return CellPilot.previewSheetRange(sheetName, range); }
function getSmartFormulaSuggestions() { return CellPilot.getSmartFormulaSuggestions(); }
function analyzeMultiTabRelationships() { return CellPilot.analyzeMultiTabRelationships(); }
function generateRelationshipReport(relationshipData) { return CellPilot.generateRelationshipReport(relationshipData); }

// ============================================
// FORMULA DEBUGGING
// ============================================

function debugActiveFormula() { return CellPilot.debugActiveFormula(); }
function debugSelectionFormulas() { return CellPilot.debugSelectionFormulas(); }
function debugAllSheetFormulas() { return CellPilot.debugAllSheetFormulas(); }
function applyFormulaFix(cellRef, formula) { return CellPilot.applyFormulaFix(cellRef, formula); }
function analyzeDependencies() { return CellPilot.analyzeDependencies(); }
function analyzeFormulaPerformance() { return CellPilot.analyzeFormulaPerformance(); }

// ============================================
// ML-POWERED FEATURES
// ============================================

function getUserLearningProfile() { return CellPilot.getUserLearningProfile(); }
function saveUserLearningProfile(profile) { return CellPilot.saveUserLearningProfile(profile); }
function trackMLFeedback(operation, prediction, userAction, metadata) { return CellPilot.trackMLFeedback(operation, prediction, userAction, metadata); }
function enableMLFeatures() { return CellPilot.enableMLFeatures(); }
function getMLStatus() { return CellPilot.getMLStatus(); }

// ============================================
// DATA VALIDATION & CONDITIONAL FORMATTING
// ============================================

function applyValidationRule(rule) { return CellPilot.applyValidationRule(rule); }
function generateValidationRules() { return CellPilot.generateValidationRules(); }
function applyConditionalFormat(rule) { return CellPilot.applyConditionalFormat(rule); }
function generateFormatSuggestions() { return CellPilot.generateFormatSuggestions(); }

// ============================================
// FORMULA DISCOVERY & PATTERN DETECTION
// ============================================

function discoverFormulaPatterns() { return CellPilot.discoverFormulaPatterns(); }
function getCommonFormulas() { return CellPilot.getCommonFormulas(); }
function suggestFormulasML(description) { return CellPilot.suggestFormulasML(description); }

// ============================================
// CROSS-SHEET FORMULA BUILDER
// ============================================

function buildCrossSheetFormula(config) { return CellPilot.buildCrossSheetFormula(config); }
function validateCrossSheetFormula(formula) { return CellPilot.validateCrossSheetFormula(formula); }
function getSheetNames() {
  try {
    const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    return sheets.map(sheet => sheet.getName());
  } catch (error) {
    console.error('Error getting sheet names:', error);
    return [];
  }
}

// ============================================
// FORMULA PERFORMANCE OPTIMIZER
// ============================================

function optimizeFormulaPerformance(formula) { return CellPilot.optimizeFormulaPerformance(formula); }
function getFormulaPerformanceMetrics(formula) { return CellPilot.getFormulaPerformanceMetrics(formula); }

// ============================================
// TEMPLATE PREVIEW & APPLICATION
// ============================================

function getTemplateCategories() { return CellPilot.getTemplateCategories(); }
function getTemplatesByCategory(category) { return CellPilot.getTemplatesByCategory(category); }
function applyTemplate(templateId, options) { return CellPilot.applyTemplate(templateId, options); }

// ============================================
// ADVANCED PARSING & SMART TABLE
// ============================================

function smartParseML(data, options) { return CellPilot.smartParseML(data, options); }
function detectColumnTypesML(data, headers) { return CellPilot.detectColumnTypesML(data, headers); }
function learnFromParsingCorrection(columnIndex, correctType, context) { return CellPilot.learnFromParsingCorrection(columnIndex, correctType, context); }

// ============================================
// ADAPTIVE DUPLICATE DETECTION
// ============================================

function detectDuplicatesML(options) { return CellPilot.detectDuplicatesML(options); }
function learnFromDuplicateFeedback(acceptedDuplicates, rejectedDuplicates, threshold) { return CellPilot.learnFromDuplicateFeedback(acceptedDuplicates, rejectedDuplicates, threshold); }
function getAdaptiveDuplicateThreshold() { return CellPilot.getAdaptiveDuplicateThreshold(); }

// ============================================
// FORMULA LEARNING & SUGGESTIONS
// ============================================

function learnFromFormulaSelection(selectedFormula, description, wasSuccessful) { return CellPilot.learnFromFormulaSelection(selectedFormula, description, wasSuccessful); }
function getPersonalizedFormulas() { return CellPilot.getPersonalizedFormulas(); }
function predictNextFormula() { return CellPilot.predictNextFormula(); }
function recordFormulaFeedback(params) { return CellPilot.recordFormulaFeedback(params); }
function insertFormulaWithTracking(params) { return CellPilot.insertFormulaWithTracking(params); }

// ============================================
// PIVOT TABLE ASSISTANT FUNCTIONS
// ============================================

function analyzePivotData(range) { return CellPilot.analyzePivotData(range); }
function createPivotFromSuggestion(suggestion) { return CellPilot.createPivotFromSuggestion(suggestion); }
function applyPivotTemplate(templateName) { return CellPilot.applyPivotTemplate(templateName); }
function getPivotTemplates() { return CellPilot.getPivotTemplates(); }
function getPivotStats() { return CellPilot.getPivotStats(); }

// ============================================
// DATA PIPELINE MANAGER FUNCTIONS
// ============================================

function importPipelineData(config) { return CellPilot.importPipelineData(config); }
function exportPipelineData(options) { return CellPilot.exportPipelineData(options); }
function getPipelineHistory() { return CellPilot.getPipelineHistory(); }
function getPipelineStats() { return CellPilot.getPipelineStats(); }
function clearPipelineHistory() { return CellPilot.clearPipelineHistory(); }
function addToPipelineHistory(item) { return CellPilot.addToPipelineHistory(item); }

// ============================================
// USER PREFERENCE LEARNING FUNCTIONS
// ============================================

function initializeUserPreferences() { return CellPilot.initializeUserPreferences(); }
function trackUserAction(action, context) { return CellPilot.trackUserAction(action, context); }
function getUserPreferenceSummary() { return CellPilot.getUserPreferenceSummary(); }
function getFeatureRecommendations() { return CellPilot.getFeatureRecommendations(); }
function getWorkflowSuggestions() { return CellPilot.getWorkflowSuggestions(); }
function getPreferredSettings(feature) { return CellPilot.getPreferredSettings(feature); }
function predictNextAction(currentAction) { return CellPilot.predictNextAction(currentAction); }
function shouldAutomate(action, context) { return CellPilot.shouldAutomate(action, context); }
function learnFromError(error, action, context) { return CellPilot.learnFromError(error, action, context); }
function learnFromSuccess(action, context, duration) { return CellPilot.learnFromSuccess(action, context, duration); }
function exportUserPreferences() { return CellPilot.exportUserPreferences(); }
function resetUserPreferences() { return CellPilot.resetUserPreferences(); }

// ============================================
// ML PERFORMANCE FUNCTIONS
// ============================================

function getMLPerformanceStats() { return CellPilot.getMLPerformanceStats(); }
function optimizeMLPerformance() { return CellPilot.optimizeMLPerformance(); }
function getMLFeedbackHistory() { return CellPilot.getMLFeedbackHistory(); }
function exportMLProfile() { return CellPilot.exportMLProfile(); }
function importMLProfile(data) { return CellPilot.importMLProfile(data); }

// ============================================
// PREVIEW & VALIDATION FUNCTIONS
// ============================================

function generatePivotPreview(config) { return CellPilot.generatePivotPreview(config); }
function generateImportPreview(config) { return CellPilot.generateImportPreview(config); }
function generateExportPreview(config) { return CellPilot.generateExportPreview(config); }

// ============================================
// ENHANCED DASHBOARD FUNCTIONS
// ============================================

function createDashboard(sheet, config) { return CellPilot.createDashboard(sheet, config); }
function setupClientsSheet(sheet) { return CellPilot.setupClientsSheet(sheet); }
function setupFollowUpSheet(sheet) { return CellPilot.setupFollowUpSheet(sheet); }
function setupRentalIncomeSheet(sheet) { return CellPilot.setupRentalIncomeSheet(sheet); }
function setupPropertyComparisonSheet(sheet) { return CellPilot.setupPropertyComparisonSheet(sheet); }
function setupSuppliersSheet(sheet) { return CellPilot.setupSuppliersSheet(sheet); }
function setupPurchaseOrdersSheet(sheet) { return CellPilot.setupPurchaseOrdersSheet(sheet); }
function setupCrewScheduleSheet(sheet) { return CellPilot.setupCrewScheduleSheet(sheet); }
function setupPayrollSheet(sheet) { return CellPilot.setupPayrollSheet(sheet); }
function setupCostImpactSheet(sheet) { return CellPilot.setupCostImpactSheet(sheet); }
function setupApprovalsSheet(sheet) { return CellPilot.setupApprovalsSheet(sheet); }
function setupCostBreakdownSheet(sheet) { return CellPilot.setupCostBreakdownSheet(sheet); }
function setupContingencySheet(sheet) { return CellPilot.setupContingencySheet(sheet); }
function setupAuthStatusSheet(sheet) { return CellPilot.setupAuthStatusSheet(sheet); }
function setupProvidersSheet(sheet) { return CellPilot.setupProvidersSheet(sheet); }
function setupClaimsSheet(sheet) { return CellPilot.setupClaimsSheet(sheet); }
function setupARAgingSheet(sheet) { return CellPilot.setupARAgingSheet(sheet); }
function setupAppealsSheet(sheet) { return CellPilot.setupAppealsSheet(sheet); }
function setupDenialTrendsSheet(sheet) { return CellPilot.setupDenialTrendsSheet(sheet); }
function setupEligibilitySheet(sheet) { return CellPilot.setupEligibilitySheet(sheet); }
function setupBenefitsSheet(sheet) { return CellPilot.setupBenefitsSheet(sheet); }
function setupSegmentsSheet(sheet) { return CellPilot.setupSegmentsSheet(sheet); }
function setupEngagementSheet(sheet) { return CellPilot.setupEngagementSheet(sheet); }
function setupMetricsSheet(sheet) { return CellPilot.setupMetricsSheet(sheet); }
function setupChannelsSheet(sheet) { return CellPilot.setupChannelsSheet(sheet); }
function setupTouchpointsSheet(sheet) { return CellPilot.setupTouchpointsSheet(sheet); }
function setupConversionSheet(sheet) { return CellPilot.setupConversionSheet(sheet); }
function setupPerformanceSheet(sheet) { return CellPilot.setupPerformanceSheet(sheet); }
function setupBudgetSheet(sheet) { return CellPilot.setupBudgetSheet(sheet); }
function setupProductsSheet(sheet) { return CellPilot.setupProductsSheet(sheet); }
function setupCostsSheet(sheet) { return CellPilot.setupCostsSheet(sheet); }
function setupTrendsSheet(sheet) { return CellPilot.setupTrendsSheet(sheet); }
function setupSeasonalitySheet(sheet) { return CellPilot.setupSeasonalitySheet(sheet); }
function setupStockLevelsSheet(sheet) { return CellPilot.setupStockLevelsSheet(sheet); }
function setupReorderPointsSheet(sheet) { return CellPilot.setupReorderPointsSheet(sheet); }
function setupResourceSheet(sheet) { return CellPilot.setupResourceSheet(sheet); }
function setupExpensesSheet(sheet) { return CellPilot.setupExpensesSheet(sheet); }
function setupRevenueSheet(sheet) { return CellPilot.setupRevenueSheet(sheet); }
function setupTimesheetSheet(sheet) { return CellPilot.setupTimesheetSheet(sheet); }
function setupInvoicesSheet(sheet) { return CellPilot.setupInvoicesSheet(sheet); }

// ============================================
// TEMPLATE MODULE PROXY FUNCTIONS
// ============================================

function buildTemplate(spreadsheet, config) { return CellPilot.TemplateBuilder.buildTemplate(spreadsheet, config); }
function buildRealEstateTemplate(spreadsheet, templateType, isPreview) { return CellPilot.RealEstateTemplate.build(spreadsheet, templateType, isPreview); }
function buildConstructionTemplate(spreadsheet, templateType, isPreview) { return CellPilot.ConstructionTemplate.build(spreadsheet, templateType, isPreview); }
function buildHealthcareTemplate(spreadsheet, templateType, isPreview) { return CellPilot.HealthcareTemplate.build(spreadsheet, templateType, isPreview); }
function buildMarketingTemplate(spreadsheet, templateType, isPreview) { return CellPilot.MarketingTemplate.build(spreadsheet, templateType, isPreview); }
function buildEcommerceTemplate(spreadsheet, templateType, isPreview) { return CellPilot.EcommerceTemplate.build(spreadsheet, templateType, isPreview); }
function buildConsultingTemplate(spreadsheet, templateType, isPreview) { return CellPilot.ConsultingTemplate.build(spreadsheet, templateType, isPreview); }

// Excel Migration Functions
function scanForExcelIssues() { return CellPilot.scanForExcelIssues(); }
function fixExcelIssues(options) { return CellPilot.fixExcelIssues(options); }
function convertXlookupToVlookup(formula) { return CellPilot.convertXlookupToVlookup(formula); }
function convertStructuredRefs(formula) { return CellPilot.convertStructuredRefs(formula); }
function optimizeVolatileFunction(formula) { return CellPilot.optimizeVolatileFunction(formula); }
function fixRefError(cellA1) { return CellPilot.fixRefError(cellA1); }
function importExcelData(pastedData) { return CellPilot.importExcelData(pastedData); }
function parseExcelValue(value) { return CellPilot.parseExcelValue(value); }
function convertExcelFormula(formula) { return CellPilot.convertExcelFormula(formula); }
function convertDateFunctions(formula) { return CellPilot.convertDateFunctions(formula); }
function convertDynamicArrays(formula) { return CellPilot.convertDynamicArrays(formula); }
function detectPivotTable(range) { return CellPilot.detectPivotTable(range); }
function convertConditionalFormatting(sheet) { return CellPilot.convertConditionalFormatting(sheet); }
function detectVBAMacros(content) { return CellPilot.detectVBAMacros(content); }
function createMigrationReport(migrationResults) { return CellPilot.createMigrationReport(migrationResults); }
function showExcelMigration() { return CellPilot.showExcelMigration(); }

// Data Pipeline Functions
function createDataPipeline(config) { return CellPilot.createDataPipeline(config); }
function runDataPipeline(pipelineId) { return CellPilot.runDataPipeline(pipelineId); }
function getDataPipelines() { return CellPilot.getDataPipelines(); }
function deleteDataPipeline(pipelineId) { return CellPilot.deleteDataPipeline(pipelineId); }
function importData(source) { return CellPilot.importData(source); }
function exportData(options) { return CellPilot.exportData(options); }
function exportToCSV(dataToExport, options) { return CellPilot.exportToCSV(dataToExport, options); }
function exportToJSON(dataToExport, options) { return CellPilot.exportToJSON(dataToExport, options); }
function exportToXML(dataToExport, options) { return CellPilot.exportToXML(dataToExport, options); }
function exportToHTML(dataToExport, options) { return CellPilot.exportToHTML(dataToExport, options); }
function exportToAPI(dataToExport, options) { return CellPilot.exportToAPI(dataToExport, options); }
function exportToEmail(dataToExport, options) { return CellPilot.exportToEmail(dataToExport, options); }
function saveToGoogleDrive(content, filename, mimeType) { return CellPilot.saveToGoogleDrive(content, filename, mimeType); }
function prepareDataForExport(options) { return CellPilot.prepareDataForExport(options); }
function showDataPipelineManager() { return CellPilot.showDataPipelineManager(); }

// CellM8 Presentation Helper Functions
function showCellM8() { return CellPilot.showCellM8(); }
function createPresentation(config) { return CellPilot.createPresentation(config); }
function getCurrentSelection() { return CellPilot.getCurrentSelection(); }
function selectEntireDataRange() { return CellPilot.selectEntireDataRange(); }
function selectRange(rangeA1) { return CellPilot.selectRange(rangeA1); }
function promptForSelection() { return CellPilot.promptForSelection(); }
function extractSheetData(source) { return CellPilot.extractSheetData(source); }
function getVisibleRange(sheet) { return CellPilot.getVisibleRange(sheet); }
function analyzeDataWithAI(data) { return CellPilot.analyzeDataWithAI(data); }
function calculateStatistics(data) { return CellPilot.calculateStatistics(data); }
function detectTrend(data) { return CellPilot.detectTrend(data); }
function calculateCorrelation(data1, data2) { return CellPilot.calculateCorrelation(data1, data2); }
function calculateCompleteness(data) { return CellPilot.calculateCompleteness(data); }
function previewPresentation(config) { return CellPilot.previewPresentation(config); }
function generateSlideStructure(analysis, config) { return CellPilot.generateSlideStructure(analysis, config); }
function createGoogleSlides(structure, config) { return CellPilot.createGoogleSlides(structure, config); }
function createSlide(presentation, slideData, index, template) { return CellPilot.createSlide(presentation, slideData, index, template); }
function createTitleSlide(slide, slideData, template) { return CellPilot.createTitleSlide(slide, slideData, template); }
function createContentSlide(slide, slideData, template) { return CellPilot.createContentSlide(slide, slideData, template); }
function createChartSlide(slide, slideData, template) { return CellPilot.createChartSlide(slide, slideData, template); }
function createMetricsSlide(slide, slideData, template) { return CellPilot.createMetricsSlide(slide, slideData, template); }
function createThankYouSlide(slide, slideData, template) { return CellPilot.createThankYouSlide(slide, slideData, template); }
function createDataSlide(slide, slideData, template) { return CellPilot.createDataSlide(slide, slideData, template); }
function applyTemplateToPresentation(presentation, template) { return CellPilot.applyTemplateToPresentation(presentation, template); }
function mapChartType(type) { return CellPilot.mapChartType(type); }
function recommendCharts(analysis) { return CellPilot.recommendCharts(analysis); }
function generateInsights(data, analysis) { return CellPilot.generateInsights(data, analysis); }
function generateRecommendations(analysis) { return CellPilot.generateRecommendations(analysis); }
function generateNextSteps(analysis) { return CellPilot.generateNextSteps(analysis); }
function extractKeyMetrics(data, analysis) { return CellPilot.extractKeyMetrics(data, analysis); }
function formatDataOverview(analysis) { return CellPilot.formatDataOverview(analysis); }
function cleanDataForPresentation(data, headers) { return CellPilot.cleanDataForPresentation(data, headers); }
function detectDataTypes(data) { return CellPilot.detectDataTypes(data); }
function detectColumnDataType(columnData) { return CellPilot.detectColumnDataType(columnData); }
function generateDefaultHeaders(count) { return CellPilot.generateDefaultHeaders(count); }
function analyzeTimeSeries(dateData) { return CellPilot.analyzeTimeSeries(dateData); }
function detectFrequency(avgInterval) { return CellPilot.detectFrequency(avgInterval); }
function analyzeCategories(categoryData) { return CellPilot.analyzeCategories(categoryData); }
function formatNumber(num) { return CellPilot.formatNumber(num); }
function chunkArray(array, size) { return CellPilot.chunkArray(array, size); }
function trackPresentationCreation(config, result) { return CellPilot.trackPresentationCreation(config, result); }

// ============================================
// END OF PROXY FUNCTIONS
// ============================================`;
}

/**
 * Include HTML files
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Test function to debug API access
 */
function testBetaAccess() {
  const testEmail = 'freddyfiesta@gmail.com';
  console.log('Testing beta access for:', testEmail);
  
  const result = checkBetaAccess(testEmail);
  console.log('Result:', result);
  
  // Also try a direct API call
  try {
    const response = UrlFetchApp.fetch('https://www.cellpilot.io/api/check-beta', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({ email: testEmail }),
      muteHttpExceptions: true
    });
    
    console.log('Direct API test:');
    console.log('Status:', response.getResponseCode());
    console.log('Headers:', response.getAllHeaders());
    console.log('Content:', response.getContentText());
  } catch (e) {
    console.log('Direct API error:', e.toString());
  }
  
  return result;
}