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
 * Opens the duplicate removal interface
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
 * Show Excel migration tool
 * @public
 */
function showExcelMigration() {
  CellPilot.showExcelMigration();
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