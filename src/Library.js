/**
 * Library Export Module
 * This file ensures all CellPilot functions are properly exposed when used as a library
 * Auto-generated from test-sheet-proxy.js
 */

// Core Menu Functions
var showCellPilotSidebar = showCellPilotSidebar || function() { return showCellPilotSidebar(); };
var showSettings = showSettings || function() { return showSettings(); };
var showHelpFeedback = showHelpFeedback || function() { return showHelpFeedback(); };
var showOnboarding = showOnboarding || function() { return showOnboarding(); };
var showUpgrade = showUpgrade || function() { return showUpgrade(); };

// Main Feature Functions  
var tableize = tableize || function() { return tableize(); };
var smartTableize = smartTableize || function() { return smartTableize(); };
var removeDuplicates = removeDuplicates || function() { return removeDuplicates(); };
var cleanData = cleanData || function() { return cleanData(); };
var buildFormula = buildFormula || function(formulaConfig) { return buildFormula(formulaConfig); };
var migrateFromExcel = migrateFromExcel || function() { return migrateFromExcel(); };
var generateTemplate = generateTemplate || function(templateType) { return generateTemplate(templateType); };
var analyzeTable = analyzeTable || function() { return analyzeTable(); };
var parseTable = parseTable || function(config) { return parseTable(config); };
var smartParseTable = smartParseTable || function(config) { return smartParseTable(config); };
var cleanColumn = cleanColumn || function(column, cleaningOptions) { return cleanColumn(column, cleaningOptions); };

// UI Functions
var previewTableize = previewTableize || function(options) { return previewTableize(options); };
var applyTableize = applyTableize || function(parsedData) { return applyTableize(parsedData); };
var findDuplicates = findDuplicates || function(options) { return findDuplicates(options); };
var applyDuplicateRemoval = applyDuplicateRemoval || function(duplicates, action) { return applyDuplicateRemoval(duplicates, action); };
var analyzeDataQuality = analyzeDataQuality || function(options) { return analyzeDataQuality(options); };
var applyDataCleaning = applyDataCleaning || function(issues) { return applyDataCleaning(issues); };
var cleanColumnData = cleanColumnData || function(column, cleaningOptions) { return cleanColumnData(column, cleaningOptions); };

// Advanced Features
var showAdvancedRestructuring = showAdvancedRestructuring || function() { return showAdvancedRestructuring(); };
var showIndustryTemplates = showIndustryTemplates || function() { return showIndustryTemplates(); };
var showConditionalFormattingWizard = showConditionalFormattingWizard || function() { return showConditionalFormattingWizard(); };
var showCrossSheetFormulaBuilder = showCrossSheetFormulaBuilder || function() { return showCrossSheetFormulaBuilder(); };
var showDataPipelineManager = showDataPipelineManager || function() { return showDataPipelineManager(); };
var showDataValidationGenerator = showDataValidationGenerator || function() { return showDataValidationGenerator(); };
var showFormulaDebugger = showFormulaDebugger || function() { return showFormulaDebugger(); };
var showPivotTableAssistant = showPivotTableAssistant || function() { return showPivotTableAssistant(); };
var showAutomation = showAutomation || function() { return showAutomation(); };
var showSmartFormulaAssistant = showSmartFormulaAssistant || function() { return showSmartFormulaAssistant(); };
var showSmartFormulaDebugger = showSmartFormulaDebugger || function() { return showSmartFormulaDebugger(); };
var showMultiTabRelationshipMapper = showMultiTabRelationshipMapper || function() { return showMultiTabRelationshipMapper(); };
var showFormulaPerformanceOptimizer = showFormulaPerformanceOptimizer || function() { return showFormulaPerformanceOptimizer(); };

// Template Functions
var applyIndustryTemplate = applyIndustryTemplate || function(templateType, options) { return applyIndustryTemplate(templateType, options); };
var loadConstructionTemplate = loadConstructionTemplate || function(options) { return loadConstructionTemplate(options); };
var loadRealEstateTemplate = loadRealEstateTemplate || function(options) { return loadRealEstateTemplate(options); };
var loadHealthcareTemplate = loadHealthcareTemplate || function(options) { return loadHealthcareTemplate(options); };
var loadMarketingTemplate = loadMarketingTemplate || function(options) { return loadMarketingTemplate(options); };
var loadEcommerceTemplate = loadEcommerceTemplate || function(options) { return loadEcommerceTemplate(options); };
var loadProfessionalServicesTemplate = loadProfessionalServicesTemplate || function(options) { return loadProfessionalServicesTemplate(options); };

// Utility Functions
var getActiveRangeData = getActiveRangeData || function() { return getActiveRangeData(); };
var getSheetNames = getSheetNames || function() { return getSheetNames(); };
var getCellValue = getCellValue || function(a1Notation) { return getCellValue(a1Notation); };
var setCellValue = setCellValue || function(a1Notation, value) { return setCellValue(a1Notation, value); };
var getRangeValues = getRangeValues || function(a1Notation) { return getRangeValues(a1Notation); };
var setRangeValues = setRangeValues || function(a1Notation, values) { return setRangeValues(a1Notation, values); };

// Settings Functions
var saveSettings = saveSettings || function(settings) { return saveSettings(settings); };
var loadSettings = loadSettings || function() { return loadSettings(); };
var toggleFeature = toggleFeature || function(featureName, enabled) { return toggleFeature(featureName, enabled); };
var updateApiKey = updateApiKey || function(service, apiKey) { return updateApiKey(service, apiKey); };

// ML Functions
var trainUserPreferences = trainUserPreferences || function(action, context, result) { return trainUserPreferences(action, context, result); };
var getUserLearningProfile = getUserLearningProfile || function() { return getUserLearningProfile(); };
var clearUserLearningProfile = clearUserLearningProfile || function() { return clearUserLearningProfile(); };
var exportMLProfile = exportMLProfile || function() { return exportMLProfile(); };
var importMLProfile = importMLProfile || function(data) { return importMLProfile(data); };
var getMLStatus = getMLStatus || function() { return getMLStatus(); };
var getMLFeedbackHistory = getMLFeedbackHistory || function() { return getMLFeedbackHistory(); };
var submitMLFeedback = submitMLFeedback || function(feedback) { return submitMLFeedback(feedback); };
var toggleMLFeature = toggleMLFeature || function(enabled) { return toggleMLFeature(enabled); };

// Formula Functions
var analyzeFormula = analyzeFormula || function(formula) { return analyzeFormula(formula); };
var generateFormula = generateFormula || function(config) { return generateFormula(config); };
var suggestFormula = suggestFormula || function(context) { return suggestFormula(context); };
var debugFormula = debugFormula || function(formula) { return debugFormula(formula); };
var optimizeFormula = optimizeFormula || function(formula) { return optimizeFormula(formula); };

// Conditional Formatting Functions
var analyzeDataForConditionalFormatting = analyzeDataForConditionalFormatting || function() { return analyzeDataForConditionalFormatting(); };
var applyConditionalFormattingRule = applyConditionalFormattingRule || function(rule) { return applyConditionalFormattingRule(rule); };
var suggestConditionalFormattingRules = suggestConditionalFormattingRules || function() { return suggestConditionalFormattingRules(); };

// Cross-Sheet Functions
var analyzeCrossSheetReferences = analyzeCrossSheetReferences || function() { return analyzeCrossSheetReferences(); };
var buildCrossSheetFormula = buildCrossSheetFormula || function(config) { return buildCrossSheetFormula(config); };
var validateCrossSheetFormula = validateCrossSheetFormula || function(formula) { return validateCrossSheetFormula(formula); };

// Data Pipeline Functions
var createDataPipeline = createDataPipeline || function(config) { return createDataPipeline(config); };
var runDataPipeline = runDataPipeline || function(pipelineId) { return runDataPipeline(pipelineId); };
var getDataPipelines = getDataPipelines || function() { return getDataPipelines(); };
var deleteDataPipeline = deleteDataPipeline || function(pipelineId) { return deleteDataPipeline(pipelineId); };

// Data Validation Functions
var generateDataValidation = generateDataValidation || function(config) { return generateDataValidation(config); };
var applyDataValidation = applyDataValidation || function(rules) { return applyDataValidation(rules); };
var suggestDataValidationRules = suggestDataValidationRules || function() { return suggestDataValidationRules(); };

// Pivot Table Functions
var analyzePivotTableData = analyzePivotTableData || function() { return analyzePivotTableData(); };
var createPivotTable = createPivotTable || function(config) { return createPivotTable(config); };
var suggestPivotTableConfiguration = suggestPivotTableConfiguration || function() { return suggestPivotTableConfiguration(); };

// Automation Functions
var getAutomationRules = getAutomationRules || function() { return getAutomationRules(); };
var createAutomationRule = createAutomationRule || function(rule) { return createAutomationRule(rule); };
var deleteAutomationRule = deleteAutomationRule || function(ruleId) { return deleteAutomationRule(ruleId); };
var runAutomationRule = runAutomationRule || function(ruleId) { return runAutomationRule(ruleId); };

// Smart Assistant Functions
var getSmartSuggestions = getSmartSuggestions || function() { return getSmartSuggestions(); };
var applySmartSuggestion = applySmartSuggestion || function(suggestionId) { return applySmartSuggestion(suggestionId); };

// Multi-Tab Functions
var analyzeTabRelationships = analyzeTabRelationships || function() { return analyzeTabRelationships(); };
var createTabRelationship = createTabRelationship || function(config) { return createTabRelationship(config); };
var visualizeTabRelationships = visualizeTabRelationships || function() { return visualizeTabRelationships(); };

// Advanced Restructuring Functions
var analyzeStructure = analyzeStructure || function() { return analyzeStructure(); };
var suggestRestructuring = suggestRestructuring || function() { return suggestRestructuring(); };
var applyRestructuring = applyRestructuring || function(plan) { return applyRestructuring(plan); };
var previewRestructuring = previewRestructuring || function(plan) { return previewRestructuring(plan); };

// API Functions
var processWebhook = processWebhook || function(e) { return processWebhook(e); };
var doPost = doPost || function(e) { return doPost(e); };
var doGet = doGet || function(e) { return doGet(e); };

// Help Functions
var sendFeedback = sendFeedback || function(feedback) { return sendFeedback(feedback); };
var reportBug = reportBug || function(bugReport) { return reportBug(bugReport); };
var requestFeature = requestFeature || function(featureRequest) { return requestFeature(featureRequest); };

// User Management Functions
var getUserProfile = getUserProfile || function() { return getUserProfile(); };
var updateUserProfile = updateUserProfile || function(profile) { return updateUserProfile(profile); };
var getUserUsageStats = getUserUsageStats || function() { return getUserUsageStats(); };

// Tracking Functions
var trackEvent = trackEvent || function(eventName, eventData) { return trackEvent(eventName, eventData); };
var trackError = trackError || function(error, context) { return trackError(error, context); };

/**
 * This ensures all functions are available when CellPilot is used as a library
 */