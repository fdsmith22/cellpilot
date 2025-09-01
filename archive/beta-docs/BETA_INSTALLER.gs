/**
 * CellPilot Beta Installation Script
 * ===================================
 * Version: 1.0.0-beta
 * 
 * INSTALLATION INSTRUCTIONS:
 * 1. Copy this entire code to your Google Sheet's Apps Script editor
 *    (Extensions → Apps Script)
 * 2. Delete any existing code in the editor
 * 3. Paste this code
 * 4. Click the "Libraries" button (+) in the left sidebar
 * 5. Add Script ID: 1EZDAGoLY8UEMdbfKTZO-AQ7pkiPe-n-zrz3Rw0ec6VBBH5MdC43Avx0O
 * 6. Select Version: 10 (or latest available)
 * 7. Set Identifier: CellPilot
 * 8. Click "Add"
 * 9. Click "Save" (💾)
 * 10. Refresh your Google Sheet
 * 11. Find CellPilot in the Extensions menu!
 * 
 * IMPORTANT: All functions below proxy to the CellPilot library.
 * You have full access to all CellPilot features during beta.
 */

// ============================================
// INITIALIZATION
// ============================================

function onOpen() {
  try {
    CellPilot.onOpen();
  } catch (e) {
    SpreadsheetApp.getUi().alert('Please complete CellPilot library setup. Check installation instructions.');
  }
}

function onInstall(e) {
  CellPilot.onInstall(e);
}

// ============================================
// MAIN FEATURES
// ============================================

function showCellPilotSidebar() {
  CellPilot.showCellPilotSidebar();
}

function tableize() {
  return CellPilot.tableize();
}

function removeDuplicates() {
  return CellPilot.removeDuplicates();
}

function cleanData() {
  return CellPilot.cleanData();
}

function showSmartFormulaAssistant() {
  CellPilot.showSmartFormulaAssistant();
}

function showFormulaBuilder() {
  CellPilot.showFormulaBuilder();
}

function showAdvancedRestructuring() {
  CellPilot.showAdvancedRestructuring();
}

function showIndustryTemplates() {
  CellPilot.showIndustryTemplates();
}

function showConditionalFormattingWizard() {
  CellPilot.showConditionalFormattingWizard();
}

function showCrossSheetFormulaBuilder() {
  CellPilot.showCrossSheetFormulaBuilder();
}

function showDataPipeline() {
  CellPilot.showDataPipeline();
}

function showDataValidationGenerator() {
  CellPilot.showDataValidationGenerator();
}

function showSmartFormulaDebugger() {
  CellPilot.showSmartFormulaDebugger();
}

function showPivotTableAssistant() {
  CellPilot.showPivotTableAssistant();
}

function showSettings() {
  CellPilot.showSettings();
}

function showHelpFeedback() {
  CellPilot.showHelpFeedback();
}

function showMLEngine() {
  CellPilot.showMLEngine();
}

function showUpgrade() {
  CellPilot.showUpgrade();
}

// ============================================
// SUPPORTING FUNCTIONS
// ============================================

function include(filename) {
  return CellPilot.include(filename);
}

function createMainSidebarHtml(context) {
  return CellPilot.createMainSidebarHtml(context);
}

function getCurrentUserContext() {
  return CellPilot.getCurrentUserContext();
}

function getExtendedUserContext() {
  return CellPilot.getExtendedUserContext();
}

function analyzeDataStructure() {
  return CellPilot.analyzeDataStructure();
}

function detectDataTypes() {
  return CellPilot.detectDataTypes();
}

function suggestDataImprovements() {
  return CellPilot.suggestDataImprovements();
}

// ============================================
// TABLEIZE FUNCTIONS
// ============================================

function tableizeSelection() {
  return CellPilot.tableizeSelection();
}

function detectTables() {
  return CellPilot.detectTables();
}

function applyTableFormatting(options) {
  return CellPilot.applyTableFormatting(options);
}

function convertToTable(data, headers) {
  return CellPilot.convertToTable(data, headers);
}

function formatAsTable(range, options) {
  return CellPilot.formatAsTable(range, options);
}

function createTableFromData(data, options) {
  return CellPilot.createTableFromData(data, options);
}

// ============================================
// DUPLICATE REMOVAL FUNCTIONS
// ============================================

function findDuplicates() {
  return CellPilot.findDuplicates();
}

function removeDuplicatesInSelection() {
  return CellPilot.removeDuplicatesInSelection();
}

function highlightDuplicates() {
  return CellPilot.highlightDuplicates();
}

function removeDuplicatesWithOptions(options) {
  return CellPilot.removeDuplicatesWithOptions(options);
}

function getDuplicateStats() {
  return CellPilot.getDuplicateStats();
}

// ============================================
// DATA CLEANING FUNCTIONS
// ============================================

function trimWhitespace() {
  return CellPilot.trimWhitespace();
}

function removeEmptyRows() {
  return CellPilot.removeEmptyRows();
}

function fixDateFormats() {
  return CellPilot.fixDateFormats();
}

function standardizePhoneNumbers() {
  return CellPilot.standardizePhoneNumbers();
}

function cleanEmails() {
  return CellPilot.cleanEmails();
}

function cleanDataWithOptions(options) {
  return CellPilot.cleanDataWithOptions(options);
}

function normalizeData() {
  return CellPilot.normalizeData();
}

function standardizeCapitalization() {
  return CellPilot.standardizeCapitalization();
}

function removeSpecialCharacters() {
  return CellPilot.removeSpecialCharacters();
}

function fillEmptyCells(fillValue) {
  return CellPilot.fillEmptyCells(fillValue);
}

// ============================================
// FORMULA FUNCTIONS
// ============================================

function generateFormula(prompt) {
  return CellPilot.generateFormula(prompt);
}

function suggestFormulas() {
  return CellPilot.suggestFormulas();
}

function explainFormula(formula) {
  return CellPilot.explainFormula(formula);
}

function fixFormula(formula) {
  return CellPilot.fixFormula(formula);
}

function optimizeFormula(formula) {
  return CellPilot.optimizeFormula(formula);
}

function buildComplexFormula(components) {
  return CellPilot.buildComplexFormula(components);
}

function convertExcelFormula(excelFormula) {
  return CellPilot.convertExcelFormula(excelFormula);
}

function validateFormula(formula) {
  return CellPilot.validateFormula(formula);
}

function getFormulaDependencies(formula) {
  return CellPilot.getFormulaDependencies(formula);
}

function suggestNextFormula() {
  return CellPilot.suggestNextFormula();
}

// ============================================
// ADVANCED RESTRUCTURING FUNCTIONS
// ============================================

function pivotData(options) {
  return CellPilot.pivotData(options);
}

function unpivotData(options) {
  return CellPilot.unpivotData(options);
}

function transposeData() {
  return CellPilot.transposeData();
}

function mergeSheets(options) {
  return CellPilot.mergeSheets(options);
}

function splitData(options) {
  return CellPilot.splitData(options);
}

function reshapeData(options) {
  return CellPilot.reshapeData(options);
}

function consolidateData(options) {
  return CellPilot.consolidateData(options);
}

function normalizeStructure() {
  return CellPilot.normalizeStructure();
}

function createRelationalStructure() {
  return CellPilot.createRelationalStructure();
}

function flattenNestedData() {
  return CellPilot.flattenNestedData();
}

// ============================================
// INDUSTRY TEMPLATE FUNCTIONS
// ============================================

function applyIndustryTemplate(industry, templateType) {
  return CellPilot.applyIndustryTemplate(industry, templateType);
}

function getAvailableTemplates() {
  return CellPilot.getAvailableTemplates();
}

function customizeTemplate(templateId, options) {
  return CellPilot.customizeTemplate(templateId, options);
}

function createCustomTemplate(config) {
  return CellPilot.createCustomTemplate(config);
}

function saveAsTemplate(name, description) {
  return CellPilot.saveAsTemplate(name, description);
}

function loadTemplate(templateId) {
  return CellPilot.loadTemplate(templateId);
}

function previewTemplate(templateId) {
  return CellPilot.previewTemplate(templateId);
}

function getTemplateInfo(templateId) {
  return CellPilot.getTemplateInfo(templateId);
}

function exportTemplate(templateId) {
  return CellPilot.exportTemplate(templateId);
}

function importTemplate(templateData) {
  return CellPilot.importTemplate(templateData);
}

// ============================================
// CONDITIONAL FORMATTING FUNCTIONS
// ============================================

function applySmartFormatting() {
  return CellPilot.applySmartFormatting();
}

function createHeatmap(options) {
  return CellPilot.createHeatmap(options);
}

function highlightOutliers() {
  return CellPilot.highlightOutliers();
}

function applyDataBars() {
  return CellPilot.applyDataBars();
}

function createColorScales(options) {
  return CellPilot.createColorScales(options);
}

function applyIconSets(options) {
  return CellPilot.applyIconSets(options);
}

function createCustomRule(rule) {
  return CellPilot.createCustomRule(rule);
}

function removeFormatting() {
  return CellPilot.removeFormatting();
}

function copyFormatting() {
  return CellPilot.copyFormatting();
}

function suggestFormatting() {
  return CellPilot.suggestFormatting();
}

// ============================================
// CROSS-SHEET FORMULA FUNCTIONS
// ============================================

function buildCrossSheetFormula(config) {
  return CellPilot.buildCrossSheetFormula(config);
}

function linkSheets(sourceSheet, targetSheet, key) {
  return CellPilot.linkSheets(sourceSheet, targetSheet, key);
}

function createLookupFormula(options) {
  return CellPilot.createLookupFormula(options);
}

function buildImportRange(url, range) {
  return CellPilot.buildImportRange(url, range);
}

function createConsolidationFormula(sheets) {
  return CellPilot.createConsolidationFormula(sheets);
}

function validateCrossSheetReferences() {
  return CellPilot.validateCrossSheetReferences();
}

function updateLinkedData() {
  return CellPilot.updateLinkedData();
}

function breakLinks() {
  return CellPilot.breakLinks();
}

function auditExternalReferences() {
  return CellPilot.auditExternalReferences();
}

function optimizeCrossSheetFormulas() {
  return CellPilot.optimizeCrossSheetFormulas();
}

// ============================================
// DATA PIPELINE FUNCTIONS
// ============================================

function createDataPipeline(config) {
  return CellPilot.createDataPipeline(config);
}

function runPipeline(pipelineId) {
  return CellPilot.runPipeline(pipelineId);
}

function schedulePipeline(pipelineId, schedule) {
  return CellPilot.schedulePipeline(pipelineId, schedule);
}

function getPipelineStatus(pipelineId) {
  return CellPilot.getPipelineStatus(pipelineId);
}

function stopPipeline(pipelineId) {
  return CellPilot.stopPipeline(pipelineId);
}

function editPipeline(pipelineId, updates) {
  return CellPilot.editPipeline(pipelineId, updates);
}

function deletePipeline(pipelineId) {
  return CellPilot.deletePipeline(pipelineId);
}

function exportPipeline(pipelineId) {
  return CellPilot.exportPipeline(pipelineId);
}

function importPipeline(pipelineData) {
  return CellPilot.importPipeline(pipelineData);
}

function getPipelineHistory(pipelineId) {
  return CellPilot.getPipelineHistory(pipelineId);
}

// ============================================
// DATA VALIDATION FUNCTIONS
// ============================================

function generateValidationRules() {
  return CellPilot.generateValidationRules();
}

function applyValidation(rules) {
  return CellPilot.applyValidation(rules);
}

function validateData() {
  return CellPilot.validateData();
}

function createDropdownList(options) {
  return CellPilot.createDropdownList(options);
}

function setNumberValidation(min, max) {
  return CellPilot.setNumberValidation(min, max);
}

function setDateValidation(startDate, endDate) {
  return CellPilot.setDateValidation(startDate, endDate);
}

function setTextValidation(pattern) {
  return CellPilot.setTextValidation(pattern);
}

function createCustomValidation(formula) {
  return CellPilot.createCustomValidation(formula);
}

function removeValidation() {
  return CellPilot.removeValidation();
}

function getValidationReport() {
  return CellPilot.getValidationReport();
}

// ============================================
// FORMULA DEBUGGER FUNCTIONS
// ============================================

function debugFormula(formula) {
  return CellPilot.debugFormula(formula);
}

function traceFormulaPrecedents() {
  return CellPilot.traceFormulaPrecedents();
}

function traceDependents() {
  return CellPilot.traceDependents();
}

function evaluateFormula(formula) {
  return CellPilot.evaluateFormula(formula);
}

function findErrors() {
  return CellPilot.findErrors();
}

function fixErrors() {
  return CellPilot.fixErrors();
}

function explainError(error) {
  return CellPilot.explainError(error);
}

function auditFormulas() {
  return CellPilot.auditFormulas();
}

function compareFormulas(formula1, formula2) {
  return CellPilot.compareFormulas(formula1, formula2);
}

function simplifyFormula(formula) {
  return CellPilot.simplifyFormula(formula);
}

// ============================================
// PIVOT TABLE FUNCTIONS
// ============================================

function createPivotTable(options) {
  return CellPilot.createPivotTable(options);
}

function suggestPivotTables() {
  return CellPilot.suggestPivotTables();
}

function updatePivotTable(pivotId, updates) {
  return CellPilot.updatePivotTable(pivotId, updates);
}

function refreshPivotTable(pivotId) {
  return CellPilot.refreshPivotTable(pivotId);
}

function analyzePivotData() {
  return CellPilot.analyzePivotData();
}

function exportPivotTable(pivotId, format) {
  return CellPilot.exportPivotTable(pivotId, format);
}

function createPivotChart(pivotId, chartType) {
  return CellPilot.createPivotChart(pivotId, chartType);
}

function drillDown(pivotId, cell) {
  return CellPilot.drillDown(pivotId, cell);
}

function groupPivotData(pivotId, field, grouping) {
  return CellPilot.groupPivotData(pivotId, field, grouping);
}

function addCalculatedField(pivotId, name, formula) {
  return CellPilot.addCalculatedField(pivotId, name, formula);
}

// ============================================
// ML/AI FUNCTIONS
// ============================================

function predictValues(options) {
  return CellPilot.predictValues(options);
}

function classifyData(options) {
  return CellPilot.classifyData(options);
}

function clusterData(options) {
  return CellPilot.clusterData(options);
}

function detectAnomalies() {
  return CellPilot.detectAnomalies();
}

function forecastTrend(options) {
  return CellPilot.forecastTrend(options);
}

function analyzeCorrelations() {
  return CellPilot.analyzeCorrelations();
}

function suggestInsights() {
  return CellPilot.suggestInsights();
}

function trainModel(config) {
  return CellPilot.trainModel(config);
}

function applyModel(modelId, data) {
  return CellPilot.applyModel(modelId, data);
}

function evaluateModel(modelId) {
  return CellPilot.evaluateModel(modelId);
}

// ============================================
// USER PREFERENCE & SETTINGS
// ============================================

function saveUserPreferences(preferences) {
  return CellPilot.saveUserPreferences(preferences);
}

function getUserPreferences() {
  return CellPilot.getUserPreferences();
}

function resetPreferences() {
  return CellPilot.resetPreferences();
}

function exportSettings() {
  return CellPilot.exportSettings();
}

function importSettings(settings) {
  return CellPilot.importSettings(settings);
}

function getUserUsageStats() {
  return CellPilot.getUserUsageStats();
}

function clearCache() {
  return CellPilot.clearCache();
}

function optimizePerformance() {
  return CellPilot.optimizePerformance();
}

function checkForUpdates() {
  return CellPilot.checkForUpdates();
}

function reportIssue(issue) {
  return CellPilot.reportIssue(issue);
}

// ============================================
// UNDO/REDO MANAGEMENT
// ============================================

function undo() {
  return CellPilot.undo();
}

function redo() {
  return CellPilot.redo();
}

function getUndoHistory() {
  return CellPilot.getUndoHistory();
}

function clearUndoHistory() {
  return CellPilot.clearUndoHistory();
}

function saveCheckpoint(name) {
  return CellPilot.saveCheckpoint(name);
}

function restoreCheckpoint(name) {
  return CellPilot.restoreCheckpoint(name);
}

// ============================================
// UTILITY FUNCTIONS
// ============================================

function getVersion() {
  return CellPilot.getVersion();
}

function getFeatureStatus(feature) {
  return CellPilot.getFeatureStatus(feature);
}

function logActivity(action, details) {
  return CellPilot.logActivity(action, details);
}

function getHelp(topic) {
  return CellPilot.getHelp(topic);
}

function runDiagnostics() {
  return CellPilot.runDiagnostics();
}

function exportData(format) {
  return CellPilot.exportData(format);
}

function importData(data, format) {
  return CellPilot.importData(data, format);
}

function shareConfiguration(email) {
  return CellPilot.shareConfiguration(email);
}

function getShortcuts() {
  return CellPilot.getShortcuts();
}

function setShortcut(action, keys) {
  return CellPilot.setShortcut(action, keys);
}

// ============================================
// TEMPLATE SPECIFIC SETUP FUNCTIONS
// ============================================

function setupPropertiesSheet(sheet) {
  return CellPilot.setupPropertiesSheet(sheet);
}

function setupCommissionsSheet(sheet) {
  return CellPilot.setupCommissionsSheet(sheet);
}

function setupAnalyticsSheet(sheet) {
  return CellPilot.setupAnalyticsSheet(sheet);
}

function setupMaterialsSheet(sheet) {
  return CellPilot.setupMaterialsSheet(sheet);
}

function setupSubcontractorsSheet(sheet) {
  return CellPilot.setupSubcontractorsSheet(sheet);
}

function setupTimelineSheet(sheet) {
  return CellPilot.setupTimelineSheet(sheet);
}

function setupPaymentsSheet(sheet) {
  return CellPilot.setupPaymentsSheet(sheet);
}

function setupVerificationSheet(sheet) {
  return CellPilot.setupVerificationSheet(sheet);
}

function setupPriorAuthSheet(sheet) {
  return CellPilot.setupPriorAuthSheet(sheet);
}

function setupClaimsSheet(sheet) {
  return CellPilot.setupClaimsSheet(sheet);
}

function setupMetricsSheet(sheet) {
  return CellPilot.setupMetricsSheet(sheet);
}

function setupCampaignsSheet(sheet) {
  return CellPilot.setupCampaignsSheet(sheet);
}

function setupLeadsSheet(sheet) {
  return CellPilot.setupLeadsSheet(sheet);
}

function setupPerformanceSheet(sheet) {
  return CellPilot.setupPerformanceSheet(sheet);
}

function setupBudgetSheet(sheet) {
  return CellPilot.setupBudgetSheet(sheet);
}

function setupProductsSheet(sheet) {
  return CellPilot.setupProductsSheet(sheet);
}

function setupCostsSheet(sheet) {
  return CellPilot.setupCostsSheet(sheet);
}

function setupTrendsSheet(sheet) {
  return CellPilot.setupTrendsSheet(sheet);
}

function setupSeasonalitySheet(sheet) {
  return CellPilot.setupSeasonalitySheet(sheet);
}

function setupStockLevelsSheet(sheet) {
  return CellPilot.setupStockLevelsSheet(sheet);
}

function setupReorderPointsSheet(sheet) {
  return CellPilot.setupReorderPointsSheet(sheet);
}

function setupResourceSheet(sheet) {
  return CellPilot.setupResourceSheet(sheet);
}

function setupExpensesSheet(sheet) {
  return CellPilot.setupExpensesSheet(sheet);
}

function setupEngagementSheet(sheet) {
  return CellPilot.setupEngagementSheet(sheet);
}

function setupRevenueSheet(sheet) {
  return CellPilot.setupRevenueSheet(sheet);
}

function setupTimesheetSheet(sheet) {
  return CellPilot.setupTimesheetSheet(sheet);
}

function setupInvoicesSheet(sheet) {
  return CellPilot.setupInvoicesSheet(sheet);
}

// ============================================
// TEMPLATE BUILDER FUNCTIONS
// ============================================

function buildTemplate(spreadsheet, config) {
  return CellPilot.TemplateBuilder.buildTemplate(spreadsheet, config);
}

function buildRealEstateTemplate(spreadsheet, templateType, isPreview) {
  return CellPilot.RealEstateTemplate.build(spreadsheet, templateType, isPreview);
}

function buildConstructionTemplate(spreadsheet, templateType, isPreview) {
  return CellPilot.ConstructionTemplate.build(spreadsheet, templateType, isPreview);
}

function buildHealthcareTemplate(spreadsheet, templateType, isPreview) {
  return CellPilot.HealthcareTemplate.build(spreadsheet, templateType, isPreview);
}

function buildMarketingTemplate(spreadsheet, templateType, isPreview) {
  return CellPilot.MarketingTemplate.build(spreadsheet, templateType, isPreview);
}

function buildEcommerceTemplate(spreadsheet, templateType, isPreview) {
  return CellPilot.EcommerceTemplate.build(spreadsheet, templateType, isPreview);
}

function buildConsultingTemplate(spreadsheet, templateType, isPreview) {
  return CellPilot.ConsultingTemplate.build(spreadsheet, templateType, isPreview);
}

/**
 * ================================
 * END OF CELLPILOT BETA INSTALLER
 * ================================
 * 
 * You now have access to all CellPilot features!
 * Remember to add the library as described in the instructions above.
 * 
 * For support: Visit cellpilot.io/support
 * For feedback: Email beta@cellpilot.io
 */