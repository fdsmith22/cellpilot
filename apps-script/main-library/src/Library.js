/**
 * Library Export Module
 * This file ensures all CellPilot functions are properly exposed when used as a library
 * Auto-generated from test-sheet-proxy.js
 */

// Version information
var getVersion = function() { return '1.0.0-beta'; };

// Core initialization functions  
var onOpen = onOpen || function(e) { return onOpen(e); };
var onInstall = onInstall || function(e) { return onInstall(e); };
var getCurrentUserContext = getCurrentUserContext || function() { return getCurrentUserContext(); };

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
var importData = importData || function(source) { return DataPipelineManager.importData(source); };
var exportData = exportData || function(options) { return DataPipelineManager.exportData(options); };
var exportToCSV = exportToCSV || function(dataToExport, options) { return DataPipelineManager.exportToCSV(dataToExport, options); };
var exportToJSON = exportToJSON || function(dataToExport, options) { return DataPipelineManager.exportToJSON(dataToExport, options); };
var exportToXML = exportToXML || function(dataToExport, options) { return DataPipelineManager.exportToXML(dataToExport, options); };
var exportToHTML = exportToHTML || function(dataToExport, options) { return DataPipelineManager.exportToHTML(dataToExport, options); };
var exportToAPI = exportToAPI || function(dataToExport, options) { return DataPipelineManager.exportToAPI(dataToExport, options); };
var exportToEmail = exportToEmail || function(dataToExport, options) { return DataPipelineManager.exportToEmail(dataToExport, options); };
var saveToGoogleDrive = saveToGoogleDrive || function(content, filename, mimeType) { return DataPipelineManager.saveToGoogleDrive(content, filename, mimeType); };
var prepareDataForExport = prepareDataForExport || function(options) { return DataPipelineManager.prepareDataForExport(options); };

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

// Excel Migration Functions
var scanForExcelIssues = scanForExcelIssues || function() { return ExcelMigration.scanForExcelIssues(); };
var fixExcelIssues = fixExcelIssues || function(options) { return ExcelMigration.fixExcelIssues(options); };
var convertXlookupToVlookup = convertXlookupToVlookup || function(formula) { return ExcelMigration.convertXlookupToVlookup(formula); };
var convertStructuredRefs = convertStructuredRefs || function(formula) { return ExcelMigration.convertStructuredRefs(formula); };
var optimizeVolatileFunction = optimizeVolatileFunction || function(formula) { return ExcelMigration.optimizeVolatileFunction(formula); };
var fixRefError = fixRefError || function(cellA1) { return ExcelMigration.fixRefError(cellA1); };
var importExcelData = importExcelData || function(pastedData) { return ExcelMigration.importExcelData(pastedData); };
var parseExcelValue = parseExcelValue || function(value) { return ExcelMigration.parseExcelValue(value); };
var convertExcelFormula = convertExcelFormula || function(formula) { return ExcelMigration.convertExcelFormula(formula); };
var convertDateFunctions = convertDateFunctions || function(formula) { return ExcelMigration.convertDateFunctions(formula); };
var convertDynamicArrays = convertDynamicArrays || function(formula) { return ExcelMigration.convertDynamicArrays(formula); };
var detectPivotTable = detectPivotTable || function(range) { return ExcelMigration.detectPivotTable(range); };
var convertConditionalFormatting = convertConditionalFormatting || function(sheet) { return ExcelMigration.convertConditionalFormatting(sheet); };
var detectVBAMacros = detectVBAMacros || function(content) { return ExcelMigration.detectVBAMacros(content); };
var createMigrationReport = createMigrationReport || function(migrationResults) { return ExcelMigration.createMigrationReport(migrationResults); };
var showExcelMigration = showExcelMigration || function() { return showExcelMigration(); };

// CellM8 Presentation Functions
var showCellM8 = showCellM8 || function() { return showCellM8(); };
var createPresentation = createPresentation || function(config) { return createPresentation(config); };
var getCellM8Selection = getCellM8Selection || function() { return getCellM8Selection(); };
var getCurrentSelection = getCurrentSelection || function() { return getCurrentSelection(); };
var selectEntireDataRange = selectEntireDataRange || function() { return selectEntireDataRange(); };
var selectRange = selectRange || function(rangeA1) { return selectRange(rangeA1); };
var promptForSelection = promptForSelection || function() { return promptForSelection(); };
var extractSheetData = extractSheetData || function(source) { return CellM8.extractSheetData(source); };
var getVisibleRange = getVisibleRange || function(sheet) { return CellM8.getVisibleRange(sheet); };
var analyzeDataWithAI = analyzeDataWithAI || function(data) { return CellM8.analyzeDataWithAI(data); };
var calculateStatistics = calculateStatistics || function(data) { return CellM8.calculateStatistics(data); };
var detectTrend = detectTrend || function(data) { return CellM8.detectTrend(data); };
var calculateCorrelation = calculateCorrelation || function(data1, data2) { return CellM8.calculateCorrelation(data1, data2); };
var calculateCompleteness = calculateCompleteness || function(data) { return CellM8.calculateCompleteness(data); };
var previewPresentation = previewPresentation || function(config) { return previewPresentation(config); };
var generateSlideStructure = generateSlideStructure || function(analysis, config) { return CellM8.generateSlideStructure(analysis, config); };
var createGoogleSlides = createGoogleSlides || function(structure, config) { return CellM8.createGoogleSlides(structure, config); };
var createSlide = createSlide || function(presentation, slideData, index, template) { return CellM8.createSlide(presentation, slideData, index, template); };
var createTitleSlide = createTitleSlide || function(slide, slideData, template) { return CellM8.createTitleSlide(slide, slideData, template); };
var createContentSlide = createContentSlide || function(slide, slideData, template) { return CellM8.createContentSlide(slide, slideData, template); };
var createChartSlide = createChartSlide || function(slide, slideData, template) { return CellM8.createChartSlide(slide, slideData, template); };
var createMetricsSlide = createMetricsSlide || function(slide, slideData, template) { return CellM8.createMetricsSlide(slide, slideData, template); };
var createThankYouSlide = createThankYouSlide || function(slide, slideData, template) { return CellM8.createThankYouSlide(slide, slideData, template); };
var createDataSlide = createDataSlide || function(slide, slideData, template) { return CellM8.createDataSlide(slide, slideData, template); };
var applyTemplateToPresentation = applyTemplateToPresentation || function(presentation, template) { return CellM8.applyTemplateToPresentation(presentation, template); };
var mapChartType = mapChartType || function(type) { return CellM8.mapChartType(type); };
var recommendCharts = recommendCharts || function(analysis) { return CellM8.recommendCharts(analysis); };
var generateInsights = generateInsights || function(data, analysis) { return CellM8.generateInsights(data, analysis); };
var generateRecommendations = generateRecommendations || function(analysis) { return CellM8.generateRecommendations(analysis); };
var generateNextSteps = generateNextSteps || function(analysis) { return CellM8.generateNextSteps(analysis); };
var extractKeyMetrics = extractKeyMetrics || function(data, analysis) { return CellM8.extractKeyMetrics(data, analysis); };
var formatDataOverview = formatDataOverview || function(analysis) { return CellM8.formatDataOverview(analysis); };
var cleanDataForPresentation = cleanDataForPresentation || function(data, headers) { return CellM8.cleanDataForPresentation(data, headers); };
var detectDataTypes = detectDataTypes || function(data) { return CellM8.detectDataTypes(data); };
var detectColumnDataType = detectColumnDataType || function(columnData) { return CellM8.detectColumnDataType(columnData); };
var generateDefaultHeaders = generateDefaultHeaders || function(count) { return CellM8.generateDefaultHeaders(count); };
var analyzeTimeSeries = analyzeTimeSeries || function(dateData) { return CellM8.analyzeTimeSeries(dateData); };
var detectFrequency = detectFrequency || function(avgInterval) { return CellM8.detectFrequency(avgInterval); };
var analyzeCategories = analyzeCategories || function(categoryData) { return CellM8.analyzeCategories(categoryData); };
var formatNumber = formatNumber || function(num) { return CellM8.formatNumber(num); };
var chunkArray = chunkArray || function(array, size) { return CellM8.chunkArray(array, size); };
var trackPresentationCreation = trackPresentationCreation || function(config, result) { return CellM8.trackPresentationCreation(config, result); };

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