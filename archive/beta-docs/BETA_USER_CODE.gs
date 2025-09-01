/**
 * CellPilot Beta User Installation Code
 * 
 * Instructions:
 * 1. Add CellPilot library with Script ID: 1EZDAGoLY8UEMdbfKTZO-AQ7pkiPe-n-zrz3Rw0ec6VBBH5MdC43Avx0O
 * 2. Set identifier as "CellPilot"
 * 3. Replace all code in your Apps Script project with this code
 * 4. Run onInstall() function to set up
 */

function onInstall(e) {
  CellPilot.onInstall(e);
}

function onOpen(e) {
  CellPilot.onOpen(e);
}

// Menu proxy functions - required because menu items can't call library functions directly
function showCellPilotSidebar() { return CellPilot.showCellPilotSidebar(); }
function showSettings() { return CellPilot.showSettings(); }
function showHelpFeedback() { return CellPilot.showHelpFeedback(); }
function showOnboarding() { return CellPilot.showOnboarding(); }
function showUpgrade() { return CellPilot.showUpgrade(); }
function showAdvancedRestructuring() { return CellPilot.showAdvancedRestructuring(); }
function showIndustryTemplates() { return CellPilot.showIndustryTemplates(); }
function showConditionalFormattingWizard() { return CellPilot.showConditionalFormattingWizard(); }
function showCrossSheetFormulaBuilder() { return CellPilot.showCrossSheetFormulaBuilder(); }
function showDataPipelineManager() { return CellPilot.showDataPipelineManager(); }
function showDataValidationGenerator() { return CellPilot.showDataValidationGenerator(); }
function showFormulaDebugger() { return CellPilot.showFormulaDebugger(); }
function showPivotTableAssistant() { return CellPilot.showPivotTableAssistant(); }
function showAutomation() { return CellPilot.showAutomation(); }
function showSmartFormulaAssistant() { return CellPilot.showSmartFormulaAssistant(); }
function showSmartFormulaDebugger() { return CellPilot.showSmartFormulaDebugger(); }
function showMultiTabRelationshipMapper() { return CellPilot.showMultiTabRelationshipMapper(); }
function showFormulaPerformanceOptimizer() { return CellPilot.showFormulaPerformanceOptimizer(); }

// Core functionality proxies - called from sidebars and dialogs
function tableize() { return CellPilot.tableize(); }
function smartTableize() { return CellPilot.smartTableize(); }
function removeDuplicates() { return CellPilot.removeDuplicates(); }
function cleanData() { return CellPilot.cleanData(); }
function buildFormula(formulaConfig) { return CellPilot.buildFormula(formulaConfig); }
function migrateFromExcel() { return CellPilot.migrateFromExcel(); }
function generateTemplate(templateType) { return CellPilot.generateTemplate(templateType); }
function analyzeTable() { return CellPilot.analyzeTable(); }
function parseTable(config) { return CellPilot.parseTable(config); }
function smartParseTable(config) { return CellPilot.smartParseTable(config); }
function cleanColumn(column, cleaningOptions) { return CellPilot.cleanColumn(column, cleaningOptions); }
function previewTableize(options) { return CellPilot.previewTableize(options); }
function applyTableize(parsedData) { return CellPilot.applyTableize(parsedData); }
function findDuplicates(options) { return CellPilot.findDuplicates(options); }
function applyDuplicateRemoval(duplicates, action) { return CellPilot.applyDuplicateRemoval(duplicates, action); }
function analyzeDataQuality(options) { return CellPilot.analyzeDataQuality(options); }
function applyDataCleaning(issues) { return CellPilot.applyDataCleaning(issues); }
function cleanColumnData(column, cleaningOptions) { return CellPilot.cleanColumnData(column, cleaningOptions); }
function applyIndustryTemplate(templateType, options) { return CellPilot.applyIndustryTemplate(templateType, options); }
function loadConstructionTemplate(options) { return CellPilot.loadConstructionTemplate(options); }
function loadRealEstateTemplate(options) { return CellPilot.loadRealEstateTemplate(options); }
function loadHealthcareTemplate(options) { return CellPilot.loadHealthcareTemplate(options); }
function loadMarketingTemplate(options) { return CellPilot.loadMarketingTemplate(options); }
function loadEcommerceTemplate(options) { return CellPilot.loadEcommerceTemplate(options); }
function loadProfessionalServicesTemplate(options) { return CellPilot.loadProfessionalServicesTemplate(options); }
function getActiveRangeData() { return CellPilot.getActiveRangeData(); }
function getSheetNames() { return CellPilot.getSheetNames(); }
function getCellValue(a1Notation) { return CellPilot.getCellValue(a1Notation); }
function setCellValue(a1Notation, value) { return CellPilot.setCellValue(a1Notation, value); }
function getRangeValues(a1Notation) { return CellPilot.getRangeValues(a1Notation); }
function setRangeValues(a1Notation, values) { return CellPilot.setRangeValues(a1Notation, values); }
function saveSettings(settings) { return CellPilot.saveSettings(settings); }
function loadSettings() { return CellPilot.loadSettings(); }
function toggleFeature(featureName, enabled) { return CellPilot.toggleFeature(featureName, enabled); }
function updateApiKey(service, apiKey) { return CellPilot.updateApiKey(service, apiKey); }
function trainUserPreferences(action, context, result) { return CellPilot.trainUserPreferences(action, context, result); }
function getUserLearningProfile() { return CellPilot.getUserLearningProfile(); }
function clearUserLearningProfile() { return CellPilot.clearUserLearningProfile(); }
function exportMLProfile() { return CellPilot.exportMLProfile(); }
function importMLProfile(data) { return CellPilot.importMLProfile(data); }
function getMLStatus() { return CellPilot.getMLStatus(); }
function getMLFeedbackHistory() { return CellPilot.getMLFeedbackHistory(); }
function submitMLFeedback(feedback) { return CellPilot.submitMLFeedback(feedback); }
function toggleMLFeature(enabled) { return CellPilot.toggleMLFeature(enabled); }
function analyzeFormula(formula) { return CellPilot.analyzeFormula(formula); }
function generateFormula(config) { return CellPilot.generateFormula(config); }
function suggestFormula(context) { return CellPilot.suggestFormula(context); }
function debugFormula(formula) { return CellPilot.debugFormula(formula); }
function optimizeFormula(formula) { return CellPilot.optimizeFormula(formula); }
function analyzeDataForConditionalFormatting() { return CellPilot.analyzeDataForConditionalFormatting(); }
function applyConditionalFormattingRule(rule) { return CellPilot.applyConditionalFormattingRule(rule); }
function suggestConditionalFormattingRules() { return CellPilot.suggestConditionalFormattingRules(); }
function analyzeCrossSheetReferences() { return CellPilot.analyzeCrossSheetReferences(); }
function buildCrossSheetFormula(config) { return CellPilot.buildCrossSheetFormula(config); }
function validateCrossSheetFormula(formula) { return CellPilot.validateCrossSheetFormula(formula); }
function createDataPipeline(config) { return CellPilot.createDataPipeline(config); }
function runDataPipeline(pipelineId) { return CellPilot.runDataPipeline(pipelineId); }
function getDataPipelines() { return CellPilot.getDataPipelines(); }
function deleteDataPipeline(pipelineId) { return CellPilot.deleteDataPipeline(pipelineId); }
function generateDataValidation(config) { return CellPilot.generateDataValidation(config); }
function applyDataValidation(rules) { return CellPilot.applyDataValidation(rules); }
function suggestDataValidationRules() { return CellPilot.suggestDataValidationRules(); }
function analyzePivotTableData() { return CellPilot.analyzePivotTableData(); }
function createPivotTable(config) { return CellPilot.createPivotTable(config); }
function suggestPivotTableConfiguration() { return CellPilot.suggestPivotTableConfiguration(); }
function getAutomationRules() { return CellPilot.getAutomationRules(); }
function createAutomationRule(rule) { return CellPilot.createAutomationRule(rule); }
function deleteAutomationRule(ruleId) { return CellPilot.deleteAutomationRule(ruleId); }
function runAutomationRule(ruleId) { return CellPilot.runAutomationRule(ruleId); }
function getSmartSuggestions() { return CellPilot.getSmartSuggestions(); }
function applySmartSuggestion(suggestionId) { return CellPilot.applySmartSuggestion(suggestionId); }
function analyzeTabRelationships() { return CellPilot.analyzeTabRelationships(); }
function createTabRelationship(config) { return CellPilot.createTabRelationship(config); }
function visualizeTabRelationships() { return CellPilot.visualizeTabRelationships(); }
function analyzeStructure() { return CellPilot.analyzeStructure(); }
function suggestRestructuring() { return CellPilot.suggestRestructuring(); }
function applyRestructuring(plan) { return CellPilot.applyRestructuring(plan); }
function previewRestructuring(plan) { return CellPilot.previewRestructuring(plan); }
function sendFeedback(feedback) { return CellPilot.sendFeedback(feedback); }
function reportBug(bugReport) { return CellPilot.reportBug(bugReport); }
function requestFeature(featureRequest) { return CellPilot.requestFeature(featureRequest); }
function getUserProfile() { return CellPilot.getUserProfile(); }
function updateUserProfile(profile) { return CellPilot.updateUserProfile(profile); }
function getUserUsageStats() { return CellPilot.getUserUsageStats(); }
function trackEvent(eventName, eventData) { return CellPilot.trackEvent(eventName, eventData); }
function trackError(error, context) { return CellPilot.trackError(error, context); }

// Web App endpoints
function doGet(e) { return CellPilot.doGet(e); }
function doPost(e) { return CellPilot.doPost(e); }
function processWebhook(e) { return CellPilot.processWebhook(e); }