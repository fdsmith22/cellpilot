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
* Triggered when the spreadsheet is opened
* Creates the CellPilot menu and initializes features
*/
function onOpen() {
CellPilot.onOpen();
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