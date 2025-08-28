/**
* CellPilot Test Proxy Functions
* ================================
* This file contains proxy functions for testing CellPilot as a library.
*
* SETUP INSTRUCTIONS:
* 1. Copy this entire code to your test sheet's Apps Script editor
* 2. Add CellPilot library with script ID: 1EZDAGoLY8UEMdbfKTZO-AQ7pkiPe-n-zrz3Rw0ec6VBBH5MdC43Avx0O
* 3. Set identifier as: CellPilot
* 4. Select the latest version
* 5. Save and reload the spreadsheet
*
* The functions below proxy all CellPilot features to your test sheet.
*/

// ============================================
// INITIALIZATION & MENU FUNCTIONS
// ============================================

/**
* Test function to verify CellPilot library is loaded
*/
function testCellPilotConnection() {
  try {
    // Test if CellPilot is available
    if (typeof CellPilot === 'undefined') {
      return { success: false, error: 'CellPilot library not found. Please add the library.' };
    }
    
    // Test a simple function
    const context = CellPilot.getCurrentUserContext();
    return { success: true, message: 'CellPilot library connected successfully!', context: context };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
* Triggered when the spreadsheet is opened
* Creates the CellPilot menu and initializes features
*/
function onOpen() {
  try {
    CellPilot.onOpen();
  } catch (e) {
    SpreadsheetApp.getUi().alert('CellPilot library not loaded properly. Please check library setup.');
    console.error('CellPilot error:', e);
  }
}

/**
* Triggered when add-on is installed
* Calls onOpen to initialize menus
*/
function onInstall(e) {
CellPilot.onInstall(e);
}

// ============================================
// MAIN SIDEBAR & NAVIGATION
// ============================================

/**
* Opens the main CellPilot sidebar dashboard
* This is the primary entry point for all features
*/
function showCellPilotSidebar() {
CellPilot.showCellPilotSidebar();
}

/**
* Includes HTML template files (for server-side includes)
* @param {string} filename - Name of the HTML file to include
*/
function include(filename) {
return CellPilot.include(filename);
}

/**
* Creates the main sidebar HTML with context
* @param {Object} context - User context data
*/
function createMainSidebarHtml(context) {
return CellPilot.createMainSidebarHtml(context);
}

/**
* Gets current user context (selection, sheet info)
* @return {Object} Context object with sheet and range data
*/
function getCurrentUserContext() {
return CellPilot.getCurrentUserContext();
}

// ============================================
// DATA CLEANING FEATURES
// ============================================

/**
* Opens the main data cleaning interface
* Includes duplicate removal, text standardization, date formatting, and validation
*/
function showDataCleaning() {
CellPilot.showDataCleaning();
}

/**
* Opens the duplicate removal interface (legacy - redirects to data cleaning)
* Allows removing duplicate rows with various matching options
*/
function showDuplicateRemoval() {
CellPilot.showDuplicateRemoval();
}

/**
* Process duplicate removal with specified options
* @param {Object} options - Duplicate detection settings
* @return {Object} Result with success status and message
*/
function removeDuplicatesProcess(options) {
return CellPilot.removeDuplicatesProcess(options);
}

/**
* Process large duplicate removal in batches
* @param {Range} range - The range to process
* @param {Object} options - Duplicate detection settings
* @return {Object} Result with success status
*/
function processLargeDuplicateRemoval(range, options) {
return CellPilot.processLargeDuplicateRemoval(range, options);
}

/**
* Preview duplicates before removal
* @param {Object} options - Duplicate detection settings
* @return {Object} Preview data with duplicate list
*/
function previewDuplicates(options) {
return CellPilot.previewDuplicates(options);
}

/**
* Opens the text standardization interface
* For cleaning and formatting text data consistently
*/
function showTextStandardization() {
CellPilot.showTextStandardization();
}

/**
* Standardize text with specified options
* @param {Object} options - Text formatting settings
* @return {Object} Result with success status
*/
function standardizeText(options) {
return CellPilot.standardizeText(options);
}

/**
* Preview text standardization changes
* @param {Object} options - Text formatting settings
* @return {Object} Preview of changes
*/
function previewStandardization(options) {
return CellPilot.previewStandardization(options);
}

// ============================================
// TABLEIZE FEATURE (DATA STRUCTURING)
// ============================================

/**
* Opens the Tableize interface
* Converts unstructured data into proper columns
*/
function showTableize() {
CellPilot.showTableize();
}

/**
* Check if data is selected in the sheet
* @return {boolean} True if data is selected
*/
function hasDataSelection() {
return CellPilot.hasDataSelection();
}

/**
* Analyze data for tableizing patterns
* @return {Object} Analysis results with detected patterns
*/
function analyzeDataForTableize() {
return CellPilot.analyzeDataForTableize();
}

/**
* Preview tableize results before applying
* @param {Object} options - Tableize settings
* @return {Object} Preview of structured data
*/
function previewTableize(options) {
return CellPilot.previewTableize(options);
}

/**
* Apply tableize to restructure data
* @param {Array} parsedData - Parsed data to apply
* @return {Object} Result with success status
*/
function applyTableize(parsedData) {
return CellPilot.applyTableize(parsedData);
}

// ============================================
// SMART AUTOMATION FEATURES
// ============================================

/**
* Opens the Smart Automation interface
* Includes format persistence, auto-categorization, migration tools
*/
function showAutomation() {
CellPilot.showAutomation();
}

/**
* Analyze cell formats for potential issues
* @param {Object} options - Analysis settings
* @return {Object} Format analysis with issues found
*/
function analyzeFormats(options) {
return CellPilot.analyzeFormats(options);
}

/**
* Scan for Excel migration issues
* @return {Object} Analysis of Excel-related problems
*/
function scanForExcelIssues() {
return CellPilot.scanForExcelIssues();
}

/**
* Fix detected Excel migration issues
* @param {Object} options - Fix options
* @return {Object} Result with fixes applied
*/
function fixExcelIssues(options) {
return CellPilot.fixExcelIssues(options);
}

// ============================================
// FORMULA BUILDER FEATURES
// ============================================

/**
* Opens the Formula Builder interface
* Create formulas using natural language
*/
function showFormulaBuilder() {
CellPilot.showFormulaBuilder();
}

/**
* Generate formula from natural language description
* @param {string} description - Natural language formula description
* @return {Object} Generated formula or error
*/
function generateFormulaFromDescription(description) {
return CellPilot.generateFormulaFromDescription(description);
}

/**
* Check if user has access to formula builder
* @return {boolean} True if user has access
*/
function checkFormulaBuilderAccess() {
return CellPilot.checkFormulaBuilderAccess();
}

/**
* Insert formula into currently selected cell
* @param {string} formulaText - Formula to insert
* @return {Object} Result with success status
*/
function insertFormulaIntoCell(formulaText) {
return CellPilot.insertFormulaIntoCell(formulaText);
}

/**
* Opens formula templates (currently shows upgrade)
*/
function showFormulaTemplates() {
CellPilot.showFormulaTemplates();
}

// ============================================
// UNDO/REDO SYSTEM
// ============================================

/**
* Get information about available undo operations
* @return {Object} Undo operation details
*/
function getUndoInfo() {
return CellPilot.getUndoInfo();
}

/**
* Perform undo of last operation
* @return {Object} Result with success status
*/
function performUndo() {
return CellPilot.performUndo();
}

// ============================================
// PREMIUM FEATURES (SHOW UPGRADE)
// ============================================

/**
* Show date formatting options (premium)
*/
function showDateFormatting() {
CellPilot.showDateFormatting();
}

/**
* Preview date formatting changes
* @param {Object} options - Date formatting options
* @return {Object} Preview of date changes
*/
function previewDateFormatting(options) {
return CellPilot.previewDateFormatting(options);
}

/**
* Apply date formatting to selected range
* @param {Object} options - Date formatting options
* @return {Object} Result with success status
*/
function formatDates(options) {
return CellPilot.formatDates(options);
}

/**
* Show email automation options (premium)
*/
function showEmailAutomation() {
CellPilot.showEmailAutomation();
}

/**
* Show calendar integration options (premium)
*/
function showCalendarIntegration() {
CellPilot.showCalendarIntegration();
}

// ============================================
// SETTINGS & HELP
// ============================================

/**
* Show settings panel
*/
function showSettings() {
CellPilot.showSettings();
}

/**
* Show help documentation
*/
function showHelp() {
CellPilot.showHelp();
}

/**
* Show upgrade dialog with pricing options
*/
function showUpgradeDialog() {
CellPilot.showUpgradeDialog();
}

/**
* Show upgrade options
*/
function showUpgradeOptions() {
CellPilot.showUpgradeOptions();
}

/**
* Track plan selection for analytics
* @param {string} plan - Selected plan name
* @return {Object} Success status
*/
function trackPlanSelection(plan) {
return CellPilot.trackPlanSelection(plan);
}

// ============================================
// UTILITY FUNCTIONS
// ============================================

/**
* Show error dialog to user
* @param {string} title - Error title
* @param {string} message - Error message
*/
function showErrorDialog(title, message) {
CellPilot.showErrorDialog(title, message);
}

// ============================================
// GOOGLE WORKSPACE ADD-ON FUNCTIONS
// ============================================

/**
* Build homepage for Google Workspace Add-on
* @param {Object} e - Event object from Google
* @return {Card} CardService card for add-on homepage
*/
function buildHomepage(e) {
return CellPilot.buildHomepage(e);
}

/**
* Launch HTML sidebar from add-on card
* @return {ActionResponse} Card navigation response
*/
function launchHtmlSidebarFromAddon() {
return CellPilot.launchHtmlSidebarFromAddon();
}

/**
* Launch Tableize from add-on card
* @return {ActionResponse} Card navigation response
*/
function launchTableizeFromAddon() {
return CellPilot.launchTableizeFromAddon();
}

/**
* Launch duplicate removal from add-on card
* @return {ActionResponse} Card navigation response
*/
function launchDuplicatesFromAddon() {
return CellPilot.launchDuplicatesFromAddon();
}

/**
* Launch formula builder from add-on card
* @return {ActionResponse} Card navigation response
*/
function launchFormulasFromAddon() {
return CellPilot.launchFormulasFromAddon();
}

/**
* Show industry template with specific category
* @param {string} category - The category to show (real-estate, construction, etc.)
* @public
*/
function showIndustryTemplate(category) {
CellPilot.showIndustryTemplate(category);
}

/**
* Format category name for display
* @param {string} category - Category identifier
* @return {string} Formatted category name
*/
function formatCategoryName(category) {
return CellPilot.formatCategoryName(category);
}

/**
* Show Excel migration tool
* @public
*/
function showExcelMigration() {
CellPilot.showExcelMigration();
}

/**
* Show Visual Formula Builder with drag-and-drop interface
* @public
*/
function showVisualFormulaBuilder() {
CellPilot.showVisualFormulaBuilder();
}

/**
* Suggest formula based on selected data patterns
* @return {string} Suggested formula or null
*/
function suggestFormulaBasedOnData() {
return CellPilot.suggestFormulaBasedOnData();
}

/**
* Validate a formula created in visual builder
* @param {string} formula - Formula to validate
* @return {Object} Validation result with issues array
*/
function validateVisualFormula(formula) {
return CellPilot.validateVisualFormula(formula);
}

/**
* Optimize a formula for better performance
* @param {string} formula - Formula to optimize
* @return {string} Optimized formula
*/
function optimizeVisualFormula(formula) {
return CellPilot.optimizeVisualFormula(formula);
}

/**
* Test a formula with sample data
* @param {string} formula - Formula to test
* @return {string} Test result or error message
*/
function testVisualFormula(formula) {
return CellPilot.testVisualFormula(formula);
}

/**
* Insert visual formula into active cell
* @param {string} formula - Formula to insert
* @return {Object} Result with cell location
*/
function insertVisualFormula(formula) {
return CellPilot.insertVisualFormula(formula);
}

/**
* Submit user feedback (bug reports, feature requests, questions)
* @param {Object} data - Feedback data including type, title, description
* @return {Object} Submission result
*/
function submitFeedback(data) {
return CellPilot.submitFeedback(data);
}

/**
* Show Smart Formula Assistant - outcome-focused formula discovery
* @public
*/
function showSmartFormulaAssistant() {
CellPilot.showSmartFormulaAssistant();
}

/**
* Get context for smart formula suggestions
* @return {Object} Current selection context with data type
*/
function getSmartFormulaContext() {
return CellPilot.getSmartFormulaContext();
}

/**
* Insert formula from Smart Formula Assistant
* @param {string} formula - Formula to insert
* @return {Object} Result with cell location
*/
function insertSmartFormula(formula) {
return CellPilot.insertSmartFormula(formula);
}

/**
* Show API integration panel
* @public
*/
function showApiIntegration() {
CellPilot.showApiIntegration();
}

/**
* Show data validation tools
* @public
*/
function showDataValidation() {
CellPilot.showDataValidation();
}

/**
* Show batch operations panel
* @public
*/
function showBatchOperations() {
CellPilot.showBatchOperations();
}

// ============================================
// CARD SERVICE FUNCTIONS (FOR ADD-ON)
// ============================================

/**
* Create menu for Google Sheets
*/
function createCellPilotMenu() {
return CellPilot.createCellPilotMenu();
}

/**
* Build main card for add-on interface
* @param {Object} context - User context
* @return {Card} Main card with all features
*/
function buildMainCard(context) {
return CellPilot.buildMainCard(context);
}

/**
* Show duplicate removal card
* @return {Card} Duplicate removal options card
*/
function showDuplicateRemovalCard() {
return CellPilot.showDuplicateRemovalCard();
}

/**
* Show text standardization card
* @return {Card} Text standardization options card
*/
function showTextStandardizationCard() {
return CellPilot.showTextStandardizationCard();
}

/**
* Show formula builder card
* @return {Card} Formula builder card
*/
function showFormulaBuilderCard() {
return CellPilot.showFormulaBuilderCard();
}

/**
* Show date formatting card
* @return {Card} Date formatting card
*/
function showDateFormattingCard() {
return CellPilot.showDateFormattingCard();
}

/**
* Show formula templates card
* @return {Card} Formula templates card
*/
function showFormulaTemplatesCard() {
return CellPilot.showFormulaTemplatesCard();
}

/**
* Show upgrade card
* @return {Card} Upgrade options card
*/
function showUpgradeCard() {
return CellPilot.showUpgradeCard();
}

// ============================================
// CARD ACTION HANDLERS
// ============================================

/**
* Process duplicate removal from card interface
* @param {Object} e - Event object with form inputs
* @return {ActionResponse} Card response with results
*/
function processDuplicateRemoval(e) {
return CellPilot.processDuplicateRemoval(e);
}

/**
* Process text standardization from card interface
* @param {Object} e - Event object with form inputs
* @return {ActionResponse} Card response with results
*/
function processTextStandardization(e) {
return CellPilot.processTextStandardization(e);
}

/**
* Process formula generation from card interface
* @param {Object} e - Event object with form inputs
* @return {ActionResponse} Card response with results
*/
function processFormulaGeneration(e) {
return CellPilot.processFormulaGeneration(e);
}

/**
* Insert generated formula from card interface
* @param {Object} e - Event object with formula
* @return {ActionResponse} Card response
*/
function insertGeneratedFormula(e) {
return CellPilot.insertGeneratedFormula(e);
}

// ============================================
// CARD BUILDERS FOR MESSAGES
// ============================================

/**
* Build error card for displaying errors
* @param {string} title - Error title
* @param {string} message - Error message
* @return {Card} Error card
*/
function buildErrorCard(title, message) {
return CellPilot.buildErrorCard(title, message);
}

/**
* Build success card for displaying success
* @param {string} title - Success title
* @param {string} message - Success message
* @return {Card} Success card
*/
function buildSuccessCard(title, message) {
return CellPilot.buildSuccessCard(title, message);
}

/**
* Build formula result card
* @param {string} result - Generated formula
* @param {string} description - Original description
* @return {Card} Formula result card
*/
function buildFormulaResultCard(result, description) {
return CellPilot.buildFormulaResultCard(result, description);
}

/**
* Apply an industry template to the current spreadsheet
* @param {string} templateType - The type of template to apply
* @return {Object} Result object with success status and details
*/
function applyIndustryTemplate(templateType) {
return CellPilot.applyIndustryTemplate(templateType);
}

/**
* Navigate to a specific sheet by name
* @param {string} sheetName - Name of the sheet to navigate to
* @return {boolean} Success status
*/
function navigateToSheet(sheetName) {
return CellPilot.navigateToSheet(sheetName);
}

/**
* Preview an industry template without applying it
* @param {string} templateType - The type of template to preview
* @return {Object} Template preview information
*/
function previewTemplate(templateType) {
return CellPilot.previewTemplate(templateType);
}

/**
* Get cross-sheet information for visual formula builder
* @return {Array} Array of sheet information objects
*/
function getCrossSheetInfo() {
return CellPilot.getCrossSheetInfo();
}

/**
* Preview data from a sheet range for cross-sheet formulas
* @param {string} sheetName - Name of the sheet
* @param {string} range - Range to preview (e.g., "A1:C10")
* @return {Object} Preview result with data
*/
function previewSheetRange(sheetName, range) {
return CellPilot.previewSheetRange(sheetName, range);
}

/**
* Get smart formula suggestions based on selected data patterns
* @return {Object} Suggestions with analysis and context
*/
function getSmartFormulaSuggestions() {
return CellPilot.getSmartFormulaSuggestions();
}

/**
* Analyze multi-tab relationships across all sheets
* @return {Object} Analysis result with relationships and sheet data
*/
function analyzeMultiTabRelationships() {
return CellPilot.analyzeMultiTabRelationships();
}

/**
* Generate a detailed relationship report
* @param {Object} relationshipData - The relationship data to create report from
* @return {Object} Result with success status and sheet name
*/
function generateRelationshipReport(relationshipData) {
return CellPilot.generateRelationshipReport(relationshipData);
}

/**
* Show multi-tab relationship mapper interface
*/
function showMultiTabRelationshipMapper() {
return CellPilot.showMultiTabRelationshipMapper();
}

/**
* Show smart formula debugger interface
*/
function showSmartFormulaDebugger() {
return CellPilot.showSmartFormulaDebugger();
}

/**
* Debug the active cell formula
* @return {Object} Debug analysis results
*/
function debugActiveFormula() {
return CellPilot.debugActiveFormula();
}

/**
* Debug all formulas in selection
* @return {Object} Debug analysis results for selection
*/
function debugSelectionFormulas() {
return CellPilot.debugSelectionFormulas();
}

/**
* Debug all formulas in the sheet
* @return {Object} Debug analysis results for entire sheet
*/
function debugAllSheetFormulas() {
return CellPilot.debugAllSheetFormulas();
}

/**
* Apply a formula fix to a cell
* @param {string} cellRef - Cell reference (e.g., "A1")
* @param {string} formula - New formula to apply
* @return {Object} Result of fix application
*/
function applyFormulaFix(cellRef, formula) {
return CellPilot.applyFormulaFix(cellRef, formula);
}

/**
* Analyze formula dependencies
* @return {Object} Dependency analysis results
*/
function analyzeDependencies() {
return CellPilot.analyzeDependencies();
}

/**
* Analyze formula performance
* @return {Object} Performance analysis results
*/
function analyzeFormulaPerformance() {
return CellPilot.analyzeFormulaPerformance();
}

// ============================================
// ML-POWERED FEATURES
// ============================================

/**
* Get user learning profile for ML features
* @return {Object} User's ML learning profile with preferences and history
*/
function getUserLearningProfile() {
return CellPilot.getUserLearningProfile();
}

/**
* Save user learning profile for ML features
* @param {Object} profile - ML profile to save
* @return {boolean} Success status
*/
function saveUserLearningProfile(profile) {
return CellPilot.saveUserLearningProfile(profile);
}

/**
* Track ML feedback for continuous learning
* @param {string} operation - Type of operation (formulaSelection, columnType, etc.)
* @param {*} prediction - What ML predicted
* @param {*} userAction - What the user actually chose
* @param {Object} metadata - Additional context
* @return {boolean} Success status
*/
function trackMLFeedback(operation, prediction, userAction, metadata) {
return CellPilot.trackMLFeedback(operation, prediction, userAction, metadata);
}

/**
* Enable ML features for the current user
* @return {Object} ML status and capabilities
*/
function enableMLFeatures() {
return CellPilot.enableMLFeatures();
}

/**
* Get ML feature status
* @return {Object} Current ML status and model information
*/
function getMLStatus() {
return CellPilot.getMLStatus();
}

// ============================================
// DATA VALIDATION & CONDITIONAL FORMATTING
// ============================================

/**
* Show data validation generator interface
*/
function showDataValidationGenerator() {
return CellPilot.showDataValidationGenerator();
}

/**
* Apply data validation rule
* @param {Object} rule - Validation rule configuration
* @return {Object} Result with success status
*/
function applyValidationRule(rule) {
return CellPilot.applyValidationRule(rule);
}

/**
* Generate data validation rules from data patterns
* @return {Object} Suggested validation rules
*/
function generateValidationRules() {
return CellPilot.generateValidationRules();
}

/**
* Show conditional formatting wizard interface
*/
function showConditionalFormattingWizard() {
return CellPilot.showConditionalFormattingWizard();
}

/**
* Apply conditional formatting rule
* @param {Object} rule - Formatting rule configuration
* @return {Object} Result with success status
*/
function applyConditionalFormat(rule) {
return CellPilot.applyConditionalFormat(rule);
}

/**
* Generate conditional formatting suggestions
* @return {Object} Suggested formatting rules
*/
function generateFormatSuggestions() {
return CellPilot.generateFormatSuggestions();
}

// ============================================
// FORMULA DISCOVERY & PATTERN DETECTION
// ============================================

/**
* Discover formula patterns in current sheet
* @return {Object} Discovered patterns and suggestions
*/
function discoverFormulaPatterns() {
return CellPilot.discoverFormulaPatterns();
}

/**
* Get commonly used formulas in sheet
* @return {Array} List of common formulas with usage stats
*/
function getCommonFormulas() {
return CellPilot.getCommonFormulas();
}

/**
* Suggest formulas based on current context
* @param {string} description - Natural language description
* @return {Array} ML-powered formula suggestions
*/
function suggestFormulasML(description) {
return CellPilot.suggestFormulasML(description);
}

// ============================================
// CROSS-SHEET FORMULA BUILDER
// ============================================

/**
* Show cross-sheet formula builder interface
*/
function showCrossSheetFormulaBuilder() {
return CellPilot.showCrossSheetFormulaBuilder();
}

/**
* Build cross-sheet reference formula
* @param {Object} config - Cross-sheet reference configuration
* @return {string} Generated formula
*/
function buildCrossSheetFormula(config) {
return CellPilot.buildCrossSheetFormula(config);
}

/**
* Validate cross-sheet references
* @param {string} formula - Formula with cross-sheet references
* @return {Object} Validation result
*/
function validateCrossSheetFormula(formula) {
return CellPilot.validateCrossSheetFormula(formula);
}

// ============================================
// FORMULA PERFORMANCE OPTIMIZER
// ============================================

/**
* Show formula performance optimizer interface
*/
function showFormulaPerformanceOptimizer() {
return CellPilot.showFormulaPerformanceOptimizer();
}

/**
* Optimize formula for performance
* @param {string} formula - Formula to optimize
* @return {Object} Optimized formula with performance metrics
*/
function optimizeFormulaPerformance(formula) {
return CellPilot.optimizeFormulaPerformance(formula);
}

/**
* Get formula performance metrics
* @param {string} formula - Formula to analyze
* @return {Object} Performance metrics
*/
function getFormulaPerformanceMetrics(formula) {
return CellPilot.getFormulaPerformanceMetrics(formula);
}

// ============================================
// TEMPLATE PREVIEW & APPLICATION
// ============================================

/**
* Get available template categories
* @return {Array} List of template categories
*/
function getTemplateCategories() {
return CellPilot.getTemplateCategories();
}

/**
* Get templates for a specific category
* @param {string} category - Template category
* @return {Array} List of templates in category
*/
function getTemplatesByCategory(category) {
return CellPilot.getTemplatesByCategory(category);
}

/**
* Apply selected template to spreadsheet
* @param {string} templateId - ID of template to apply
* @param {Object} options - Application options
* @return {Object} Result with created sheets
*/
function applyTemplate(templateId, options) {
return CellPilot.applyTemplate(templateId, options);
}

// ============================================
// ADVANCED PARSING & SMART TABLE
// ============================================

/**
* Parse data using ML-enhanced smart parser
* @param {Array} data - Data to parse
* @param {Object} options - Parsing options
* @return {Array} Parsed data with ML confidence scores
*/
function smartParseML(data, options) {
return CellPilot.smartParseML(data, options);
}

/**
* Detect column types using ML
* @param {Array} data - Data to analyze
* @param {Array} headers - Column headers
* @return {Array} Detected column types with confidence
*/
function detectColumnTypesML(data, headers) {
return CellPilot.detectColumnTypesML(data, headers);
}

/**
* Learn from parsing corrections
* @param {number} columnIndex - Column that was corrected
* @param {string} correctType - Correct column type
* @param {Object} context - Additional context
* @return {boolean} Success status
*/
function learnFromParsingCorrection(columnIndex, correctType, context) {
return CellPilot.learnFromParsingCorrection(columnIndex, correctType, context);
}

// ============================================
// ADAPTIVE DUPLICATE DETECTION
// ============================================

/**
* Detect duplicates using ML-enhanced detection
* @param {Object} options - Detection options
* @return {Object} Detected duplicates with ML confidence
*/
function detectDuplicatesML(options) {
return CellPilot.detectDuplicatesML(options);
}

/**
* Learn from duplicate feedback
* @param {Array} acceptedDuplicates - User-confirmed duplicates
* @param {Array} rejectedDuplicates - User-rejected duplicates
* @param {number} threshold - Current threshold
* @return {Object} Updated threshold and learning status
*/
function learnFromDuplicateFeedback(acceptedDuplicates, rejectedDuplicates, threshold) {
return CellPilot.learnFromDuplicateFeedback(acceptedDuplicates, rejectedDuplicates, threshold);
}

/**
* Get adaptive duplicate threshold
* @return {number} Current adaptive threshold
*/
function getAdaptiveDuplicateThreshold() {
  return CellPilot.getAdaptiveDuplicateThreshold();
}

// ============================================
// FORMULA LEARNING & SUGGESTIONS
// ============================================

/**
* Learn from formula selection
* @param {string} selectedFormula - Formula user selected
* @param {string} description - Original description
* @param {boolean} wasSuccessful - Whether formula worked
* @return {boolean} Success status
*/
function learnFromFormulaSelection(selectedFormula, description, wasSuccessful) {
return CellPilot.learnFromFormulaSelection(selectedFormula, description, wasSuccessful);
}

/**
* Get personalized formula recommendations
* @return {Array} Personalized formula suggestions
*/
function getPersonalizedFormulas() {
return CellPilot.getPersonalizedFormulas();
}

/**
* Predict next formula based on context
* @return {Object} Predicted formula with confidence
*/
function predictNextFormula() {
return CellPilot.predictNextFormula();
}

/**
* Record formula feedback for ML learning
* @param {Object} params - Feedback parameters
* @return {Card} Success confirmation
*/
function recordFormulaFeedback(params) {
return CellPilot.recordFormulaFeedback(params);
}

/**
* Insert formula with ML tracking
* @param {Object} params - Formula and tracking parameters
* @return {Card} Success confirmation with tracking
*/
function insertFormulaWithTracking(params) {
return CellPilot.insertFormulaWithTracking(params);
}

/**
* ================================
* PIVOT TABLE ASSISTANT FUNCTIONS
* ================================
*/

/**
* Show Pivot Table Assistant interface
*/
function showPivotTableAssistant() {
  return CellPilot.showPivotTableAssistant();
}

/**
* Analyze data for pivot table opportunities
*/
function analyzePivotData(range) {
  return CellPilot.analyzePivotData(range);
}

/**
* Create pivot table from suggestion
*/
function createPivotFromSuggestion(suggestion) {
  return CellPilot.createPivotFromSuggestion(suggestion);
}

/**
* Apply pivot table template
*/
function applyPivotTemplate(templateName) {
  return CellPilot.applyPivotTemplate(templateName);
}

/**
* Get pivot table templates
*/
function getPivotTemplates() {
  return CellPilot.getPivotTemplates();
}

/**
* Get pivot table statistics
*/
function getPivotStats() {
  return CellPilot.getPivotStats();
}

/**
* ================================
* DATA PIPELINE MANAGER FUNCTIONS
* ================================
*/

/**
* Show Data Pipeline Manager interface
*/
function showDataPipelineManager() {
  return CellPilot.showDataPipelineManager();
}

/**
* Import data through pipeline
*/
function importPipelineData(config) {
  return CellPilot.importPipelineData(config);
}

/**
* Export data through pipeline
*/
function exportPipelineData(options) {
  return CellPilot.exportPipelineData(options);
}

/**
* Get pipeline import history
*/
function getPipelineHistory() {
  return CellPilot.getPipelineHistory();
}

/**
* Get pipeline statistics
*/
function getPipelineStats() {
  return CellPilot.getPipelineStats();
}

/**
* Clear pipeline history
*/
function clearPipelineHistory() {
  return CellPilot.clearPipelineHistory();
}

/**
* ================================
* USER PREFERENCE LEARNING FUNCTIONS
* ================================
*/

/**
* Initialize user preferences
*/
function initializeUserPreferences() {
  return CellPilot.initializeUserPreferences();
}

/**
* Track user action for ML learning
*/
function trackUserAction(action, context) {
  return CellPilot.trackUserAction(action, context);
}

/**
* Get user preference summary
*/
function getUserPreferenceSummary() {
  return CellPilot.getUserPreferenceSummary();
}

/**
* Get feature recommendations based on usage
*/
function getFeatureRecommendations() {
  return CellPilot.getFeatureRecommendations();
}

/**
* Get workflow suggestions
*/
function getWorkflowSuggestions() {
  return CellPilot.getWorkflowSuggestions();
}

/**
* Get preferred settings for a feature
*/
function getPreferredSettings(feature) {
  return CellPilot.getPreferredSettings(feature);
}

/**
* Predict next user action
*/
function predictNextAction(currentAction) {
  return CellPilot.predictNextAction(currentAction);
}

/**
* Check if action should be automated
*/
function shouldAutomate(action, context) {
  return CellPilot.shouldAutomate(action, context);
}

/**
* Learn from error occurrence
*/
function learnFromError(error, action, context) {
  return CellPilot.learnFromError(error, action, context);
}

/**
* Learn from successful completion
*/
function learnFromSuccess(action, context, duration) {
  return CellPilot.learnFromSuccess(action, context, duration);
}

/**
* Export user preferences
*/
function exportUserPreferences() {
  return CellPilot.exportUserPreferences();
}

/**
* Reset user preferences
*/
function resetUserPreferences() {
  return CellPilot.resetUserPreferences();
}

/**
* ================================
* ML PERFORMANCE FUNCTIONS
* ================================
*/

/**
* Get ML performance statistics
*/
function getMLPerformanceStats() {
  return CellPilot.getMLPerformanceStats();
}

/**
* Optimize ML performance
*/
function optimizeMLPerformance() {
  return CellPilot.optimizeMLPerformance();
}

/**
* Enable ML features
*/
function enableMLFeatures() {
  return CellPilot.enableMLFeatures();
}

/**
* Get ML feedback history
*/
function getMLFeedbackHistory() {
  return CellPilot.getMLFeedbackHistory();
}

/**
* Track ML feedback
*/
function trackMLFeedback(operation, prediction, userAction, metadata) {
  return CellPilot.trackMLFeedback(operation, prediction, userAction, metadata);
}

/**
* ================================
* CONDITIONAL FORMATTING & VALIDATION
* ================================
*/

/**
* Show Data Validation Generator
*/
function showDataValidationGenerator() {
  return CellPilot.showDataValidationGenerator();
}

/**
* Show Conditional Formatting Wizard
*/
function showConditionalFormattingWizard() {
  return CellPilot.showConditionalFormattingWizard();
}

/**
* ================================
* PREVIEW & VALIDATION FUNCTIONS
* ================================
*/

/**
* Generate pivot table preview
*/
function generatePivotPreview(config) {
  return CellPilot.generatePivotPreview(config);
}

/**
* Generate import preview for data pipeline
*/
function generateImportPreview(config) {
  return CellPilot.generateImportPreview(config);
}

/**
* Generate export preview for data pipeline
*/
function generateExportPreview(config) {
  return CellPilot.generateExportPreview(config);
}

/**
* Get pipeline history
*/
function getPipelineHistory() {
  return CellPilot.getPipelineHistory();
}

/**
* Add to pipeline history
*/
function addToPipelineHistory(item) {
  return CellPilot.addToPipelineHistory(item);
}

/**
* Clear pipeline history
*/
function clearPipelineHistory() {
  return CellPilot.clearPipelineHistory();
}

/**
* ================================
* END OF PROXY FUNCTIONS
* ================================
*
* All functions above proxy to the CellPilot library.
* To add new features:
* 1. Add the function to CellPilot library
* 2. Add a proxy function here
* 3. Test in your spreadsheet
* 4. Push changes to production
*/