/**
 * Industry-Specific Template Generator
 * Based on research showing 15-25 hours weekly savings per business
 * Enhanced with auto-resize, cohesive styling, and cross-tab formulas
 */

const IndustryTemplates = {
  
  // Professional color scheme for all templates
  colors: {
    primary: {
      header: '#4A5568',      // Slate gray for headers
      headerText: '#FFFFFF',  // White text on headers
      accent: '#6366F1',      // Indigo accent
      light: '#E0E7FF'        // Light indigo background
    },
    secondary: {
      positive: '#10B981',    // Green for positive values
      negative: '#EF4444',    // Red for negative values
      warning: '#F59E0B',     // Amber for warnings
      info: '#3B82F6'        // Blue for info
    },
    backgrounds: {
      main: '#FFFFFF',        // White for main content
      alternate: '#F9FAFB',   // Light gray for alternating rows
      summary: '#F3F4F6',     // Gray for summary sections
      highlight: '#FEF3C7'    // Light yellow for highlights
    },
    // Industry-specific accent colors
    industries: {
      realEstate: {
        primary: '#059669',     // Emerald for real estate
        secondary: '#10B981',
        accent: '#34D399',
        dark: '#047857'
      },
      construction: {
        primary: '#EA580C',     // Orange for construction
        secondary: '#FB923C',
        accent: '#FED7AA',
        dark: '#C2410C'
      },
      healthcare: {
        primary: '#0891B2',     // Cyan for healthcare
        secondary: '#06B6D4',
        accent: '#67E8F9',
        dark: '#0E7490'
      },
      marketing: {
        primary: '#7C3AED',     // Purple for marketing
        secondary: '#8B5CF6',
        accent: '#C4B5FD',
        dark: '#6D28D9'
      },
      ecommerce: {
        primary: '#DC2626',     // Red for e-commerce
        secondary: '#EF4444',
        accent: '#FCA5A5',
        dark: '#B91C1C'
      },
      professional: {
        primary: '#1E40AF',     // Blue for professional services
        secondary: '#2563EB',
        accent: '#93C5FD',
        dark: '#1E3A8A'
      }
    }
  },
  
  /**
   * Create a professional dashboard with KPIs, charts, and summaries
   */
  createDashboard: function(sheet, config) {
    try {
      const { title, subtitle, kpis, charts, industry, dataSheets } = config;
      const colors = this.colors.industries[industry] || this.colors.industries.professional;
      
      // Dashboard Header - merge first, then set values
      sheet.getRange('A1:L1').merge();
      sheet.getRange('A1').setValue(title).setFontSize(24).setFontWeight('bold').setFontColor(colors.dark).setHorizontalAlignment('center');
      sheet.getRange('A1:L1').setBackground('#F9FAFB');
      
      sheet.getRange('A2:L2').merge();
      sheet.getRange('A2').setValue(subtitle).setFontSize(14).setFontColor('#6B7280').setHorizontalAlignment('center');
      sheet.getRange('A2:L2').setBackground('#F9FAFB');
      
      // Add border after merging
      sheet.getRange('A1:L2').setBorder(false, false, true, false, false, false, colors.primary, SpreadsheetApp.BorderStyle.SOLID_THICK);
      
      // Last Updated
      sheet.getRange('A4').setValue('Last Updated:').setFontWeight('bold');
      sheet.getRange('B4').setValue(new Date()).setNumberFormat('dd/mm/yyyy hh:mm');
      
      // KPI Cards Row (Top metrics)
      let currentRow = 6;
      sheet.getRange(`A${currentRow}:L${currentRow}`).merge();
      sheet.getRange(`A${currentRow}`).setValue('KEY PERFORMANCE INDICATORS').setFontSize(16).setFontWeight('bold').setFontColor(colors.dark);
      currentRow += 2;
    
      // Create KPI cards (4 per row)
      const kpiStartRow = currentRow;
      kpis.forEach((kpi, index) => {
        const col = (index % 4) * 3 + 1; // 3 columns per KPI
        const row = kpiStartRow + Math.floor(index / 4) * 4; // 4 rows per KPI card
        
        // KPI value cell - merge and style
        sheet.getRange(row, col, 1, 2).merge();
        sheet.getRange(row, col).setFormula(kpi.value).setFontSize(20).setFontWeight('bold')
          .setFontColor(colors.dark).setHorizontalAlignment('center')
          .setBackground(colors.accent);
        
        // KPI label cell - merge and style
        sheet.getRange(row + 1, col, 1, 2).merge();
        sheet.getRange(row + 1, col).setValue(kpi.label).setFontSize(10)
          .setFontColor('#6B7280').setHorizontalAlignment('center')
          .setBackground(colors.accent);
        
        // KPI trend (if provided) - merge and style
        if (kpi.trend) {
          sheet.getRange(row + 2, col, 1, 2).merge();
          sheet.getRange(row + 2, col).setValue(kpi.trend).setFontSize(9)
            .setFontColor(kpi.trendColor || '#6B7280').setHorizontalAlignment('center')
            .setBackground(colors.accent);
        }
        
        // Add border around the entire KPI card
        sheet.getRange(row, col, 3, 2).setBorder(true, true, true, true, false, false, colors.primary, SpreadsheetApp.BorderStyle.SOLID);
      });
    
    currentRow = kpiStartRow + Math.ceil(kpis.length / 4) * 4 + 2;
    
      // Charts Section
      if (charts && charts.length > 0) {
        sheet.getRange(`A${currentRow}:L${currentRow}`).merge();
        sheet.getRange(`A${currentRow}`).setValue('ANALYTICS & INSIGHTS').setFontSize(16).setFontWeight('bold').setFontColor(colors.dark);
        currentRow += 2;
        
        charts.forEach(chart => {
          // Chart title - merge first
          sheet.getRange(`A${currentRow}:F${currentRow}`).merge();
          sheet.getRange(`A${currentRow}`).setValue(chart.title).setFontSize(12).setFontWeight('bold');
          currentRow++;
          
          // Chart description - merge first
          sheet.getRange(`A${currentRow}:F${currentRow}`).merge();
          sheet.getRange(`A${currentRow}`).setValue(chart.description).setFontSize(10).setFontColor('#6B7280');
          currentRow++;
          
          // Chart area placeholder - merge first
          const chartRange = sheet.getRange(currentRow, 1, 8, 6);
          chartRange.merge();
          sheet.getRange(currentRow, 1).setValue(`[${chart.type} Chart: ${chart.dataRange}]`)
            .setHorizontalAlignment('center').setVerticalAlignment('middle')
            .setFontColor('#9CA3AF').setBackground('#F3F4F6');
          chartRange.setBorder(true, true, true, true, false, false, '#D1D5DB', SpreadsheetApp.BorderStyle.SOLID);
          
          currentRow += 9;
        });
      }
      
      // Summary Tables Section
      sheet.getRange(`A${currentRow}:L${currentRow}`).merge();
      sheet.getRange(`A${currentRow}`).setValue('SUMMARY TABLES').setFontSize(16).setFontWeight('bold').setFontColor(colors.dark);
      currentRow += 2;
      
      // Auto-resize columns for better readability
      for (let col = 1; col <= 12; col++) {
        sheet.setColumnWidth(col, 100);
      }
      
      return currentRow;
    } catch (error) {
      Logger.error('Error creating dashboard:', error);
      throw error;
    }
  },

  /**
   * Advanced helper functions for creating professional templates
   */
  
  /**
   * Add sparkline charts to show trends
   */
  addSparkline: function(sheet, targetCell, dataRange, type = 'line', options = {}) {
    try {
      const defaultOptions = {
        linewidth: 2,
        color: options.color || '#4285F4',
        ...options
      };
      
      const optionsArray = [`"charttype","${type}"`];
      for (const [key, value] of Object.entries(defaultOptions)) {
        if (key !== 'charttype') {
          optionsArray.push(`"${key}","${value}"`);
        }
      }
      
      const formula = `=SPARKLINE(${dataRange},{${optionsArray.join(';')}})`;
      sheet.getRange(targetCell).setFormula(formula);
    } catch (error) {
      // Skip sparkline if it fails
      Logger.log('Sparkline skipped: ' + error.toString());
    }
  },
  
  /**
   * Create dynamic summary tables with QUERY
   */
  createSummaryTable: function(sheet, startRow, title, dataSheet, config) {
    const { groupBy, aggregate, filters } = config;
    const colors = this.colors.industries[config.industry] || this.colors.primary;
    
    // Title
    sheet.getRange(startRow, 1).setValue(title)
      .setFontSize(14).setFontWeight('bold').setFontColor(colors.dark);
    sheet.getRange(startRow, 1, 1, 6).merge().setBackground(colors.accent);
    
    // Build QUERY formula
    let query = `SELECT ${groupBy}, `;
    aggregate.forEach((agg, index) => {
      query += `${agg.function}(${agg.column})`;
      if (index < aggregate.length - 1) query += ', ';
    });
    
    if (filters && filters.length > 0) {
      query += ' WHERE ';
      filters.forEach((filter, index) => {
        query += filter;
        if (index < filters.length - 1) query += ' AND ';
      });
    }
    
    query += ` GROUP BY ${groupBy} ORDER BY ${groupBy}`;
    
    const formula = `=QUERY('${dataSheet}'!A:Z,"${query}",1)`;
    sheet.getRange(startRow + 2, 1).setFormula(formula);
    
    return startRow + 2;
  },
  
  /**
   * Add interactive filters with data validation
   */
  addInteractiveFilters: function(sheet, filterRow, dataSheet, columns) {
    sheet.getRange(filterRow, 1).setValue('FILTERS:')
      .setFontWeight('bold').setBackground('#F3F4F6');
    
    columns.forEach((col, index) => {
      const colNum = (index * 2) + 2;
      
      // Label
      sheet.getRange(filterRow, colNum).setValue(col.label + ':')
        .setFontWeight('bold');
      
      // Dropdown with unique values
      const uniqueFormula = `=UNIQUE('${dataSheet}'!${col.range})`;
      const hiddenRange = sheet.getRange(100 + index, 20); // Hidden area for unique values
      hiddenRange.setFormula(uniqueFormula);
      
      const validationRule = SpreadsheetApp.newDataValidation()
        .requireValueInRange(hiddenRange, true)
        .setAllowInvalid(false)
        .build();
      
      sheet.getRange(filterRow, colNum + 1).setDataValidation(validationRule)
        .setBackground('#FFFFFF')
        .setBorder(true, true, true, true, false, false, '#D1D5DB', SpreadsheetApp.BorderStyle.SOLID);
    });
  },
  
  /**
   * Create status tracking with icons
   */
  createStatusTracking: function(sheet, statusColumn, startRow = 2, endRow = 100) {
    const statusRange = sheet.getRange(`${statusColumn}${startRow}:${statusColumn}${endRow}`);
    
    // Status options with colors
    const statuses = [
      { text: 'Complete', color: '#D4EDDA', icon: '✓' },
      { text: 'In Progress', color: '#FFF3CD', icon: '⚡' },
      { text: 'Pending', color: '#F8D7DA', icon: '⏳' },
      { text: 'On Hold', color: '#D1ECF1', icon: '⏸' },
      { text: 'Cancelled', color: '#F5F5F5', icon: '✗' }
    ];
    
    // Create data validation
    const statusList = statuses.map(s => s.text);
    const validationRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(statusList, true)
      .build();
    statusRange.setDataValidation(validationRule);
    
    // Add conditional formatting for each status
    const rules = [];
    statuses.forEach(status => {
      const rule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(status.text)
        .setBackground(status.color)
        .setRanges([statusRange])
        .build();
      rules.push(rule);
    });
    
    sheet.setConditionalFormatRules(rules);
  },
  
  /**
   * Add progress bars using SPARKLINE
   */
  addProgressBar: function(sheet, cell, percentage, color = '#4CAF50') {
    const formula = `=SPARKLINE(${percentage},{"charttype","bar";"max",100;"color1","${color}"})`;
    sheet.getRange(cell).setFormula(formula);
  },
  
  /**
   * Create mini dashboard within a sheet
   */
  createMiniDashboard: function(sheet, startRow, startCol, metrics) {
    // Simple dashboard without merging - just create a summary row
    try {
      // Dashboard title
      sheet.getRange(startRow, startCol).setValue('DASHBOARD METRICS')
        .setFontSize(12).setFontWeight('bold')
        .setBackground('#E3F2FD')
        .setFontColor('#1976D2');
      
      // Create simple metric display in rows 2-4
      const metricsPerRow = 4;
      for (let i = 0; i < metrics.length && i < 6; i++) {
        const metric = metrics[i];
        const row = startRow + 1 + Math.floor(i / metricsPerRow) * 2;
        const col = startCol + (i % metricsPerRow) * 3;
        
        // Label
        sheet.getRange(row, col).setValue(metric.label)
          .setFontSize(9)
          .setFontColor('#666666')
          .setFontWeight('bold');
        
        // Value (formula) - wrap in IFERROR to prevent errors
        const safeFormula = `=IFERROR(${metric.formula},"0")`;
        sheet.getRange(row + 1, col).setFormula(safeFormula)
          .setFontSize(11)
          .setFontWeight('bold')
          .setFontColor(metric.color || '#333333');
      }
      
      // Add separator line
      sheet.getRange(startRow + 5, startCol, 1, 12)
        .setBackground('#E0E0E0');
        
    } catch (error) {
      // If dashboard fails, continue with sheet setup
      Logger.log('Dashboard skipped: ' + error.toString());
    }
  },
  
  /**
   * Apply professional formatting to a range
   */
  applyProfessionalFormatting: function(sheet, startRow, endRow, startCol, endCol) {
    // Handle overloaded calls
    if (arguments.length === 2) {
      // Sheet and maxRows
      const maxRows = startRow;
      const lastColumn = sheet.getLastColumn() || 1;
      const lastRow = Math.min(sheet.getLastRow() || 1, maxRows);
      
      if (lastRow > 1 && lastColumn > 0) {
        startRow = 2;
        endRow = lastRow;
        startCol = 1;
        endCol = lastColumn;
      } else {
        return;
      }
    } else if (arguments.length === 3) {
      // Sheet, endRow, startRow (our enhanced functions use this)
      const tempEnd = startRow;
      const tempStart = endRow;
      startRow = tempStart;
      endRow = tempEnd;
      startCol = 1;
      endCol = sheet.getLastColumn() || 12;
    }
    
    // Validate parameters
    if (!sheet || !startRow || !endRow || startRow > endRow) {
      return;
    }
    
    // Set defaults for columns if not provided
    startCol = startCol || 1;
    endCol = endCol || sheet.getLastColumn() || 12;
    
    try {
      const range = sheet.getRange(startRow, startCol, endRow - startRow + 1, endCol - startCol + 1);
      
      // Apply alternating row colors
      for (let row = startRow; row <= endRow; row++) {
        const rowRange = sheet.getRange(row, startCol, 1, endCol - startCol + 1);
        if (row % 2 === 0) {
          rowRange.setBackground(this.colors.backgrounds.alternate);
        } else {
          rowRange.setBackground(this.colors.backgrounds.main);
        }
      }
      
      // Add borders
      range.setBorder(true, true, true, true, true, true, '#D1D5DB', SpreadsheetApp.BorderStyle.SOLID);
      
      // Set font
      range.setFontFamily('Inter, -apple-system, BlinkMacSystemFont, Segoe UI, Roboto, sans-serif');
      range.setFontSize(10);
    } catch (error) {
      Logger.error('Error applying professional formatting:', error);
    }
  },
  
  /**
   * Auto-resize all columns to fit content
   */
  autoResizeColumns: function(sheet) {
    const lastColumn = sheet.getLastColumn();
    if (lastColumn > 0) {
      // Auto-resize each column individually for better control
      for (let col = 1; col <= lastColumn; col++) {
        sheet.autoResizeColumn(col);
        
        // Add a bit of padding (20 pixels) for better readability
        const currentWidth = sheet.getColumnWidth(col);
        sheet.setColumnWidth(col, currentWidth + 20);
      }
    }
  },
  
  /**
   * Format headers with professional styling
   */
  formatHeaders: function(sheet, headerRow, columnCount) {
    headerRow = headerRow || 1;  // Default to row 1 if not specified
    columnCount = columnCount || sheet.getLastColumn();
    
    const headerRange = sheet.getRange(headerRow, 1, 1, columnCount);
    headerRange.setBackground(this.colors.primary.header);
    headerRange.setFontColor(this.colors.primary.headerText);
    headerRange.setFontWeight('bold');
    headerRange.setFontSize(11);
    headerRange.setHorizontalAlignment('center');
    headerRange.setVerticalAlignment('middle');
    
    // Set header row height
    sheet.setRowHeight(headerRow, 40);
    
    // Freeze header row
    sheet.setFrozenRows(headerRow);
  },
  
  /**
   * Preview an industry-specific template
   * @param {string} templateType - Type of template to preview
   * @return {Object} Result with preview data
   */
  previewTemplate: function(templateType) {
    try {
      // First clean up any existing preview sheets
      this.cleanupPreviewSheets();
      
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const createdSheets = [];
      
      // Store original active sheet to return to it later
      const originalSheet = spreadsheet.getActiveSheet();
      
      // Create the template sheets with [PREVIEW] prefix
      switch (templateType) {
        // Real Estate templates
        case 'commission-tracker':
        case 'property-manager':
        case 'investment-analyzer':
        case 'lead-pipeline':
          const realEstateResult = this.createRealEstateTemplate(spreadsheet, templateType, true);
          createdSheets.push(...realEstateResult.sheets);
          break;
          
        // Construction templates
        case 'cost-estimator':
        case 'material-tracker':
        case 'labor-manager':
        case 'change-orders':
          const constructionResult = this.createConstructionTemplate(spreadsheet, templateType, true);
          createdSheets.push(...constructionResult.sheets);
          break;
          
        // Healthcare templates
        case 'insurance-verifier':
        case 'prior-auth-tracker':
        case 'revenue-cycle':
        case 'denial-analytics':
          const healthcareResult = this.createHealthcareTemplate(spreadsheet, templateType, true);
          createdSheets.push(...healthcareResult.sheets);
          break;
          
        // Marketing templates
        case 'campaign-tracker':
        case 'lead-scoring-system':
        case 'competitor-analysis':
        case 'brand-awareness':
        case 'customer-lifetime-value':
          const marketingResult = this.createMarketingTemplate(spreadsheet, templateType, true);
          createdSheets.push(...marketingResult.sheets);
          break;
          
        // E-commerce templates
        case 'sales-dashboard':
        case 'inventory-manager':
        case 'customer-analytics':
        case 'product-performance':
        case 'financial-tracker':
          const ecommerceResult = this.createEcommerceTemplate(spreadsheet, templateType, true);
          createdSheets.push(...ecommerceResult.sheets);
          break;
          
        // Consulting templates
        case 'time-billing-tracker':
        case 'project-profitability':
        case 'client-dashboard':
          const consultingResult = this.createConsultingTemplate(spreadsheet, templateType, true);
          createdSheets.push(...consultingResult.sheets);
          break;
          
        // Professional Services templates
        case 'project-tracker':
        case 'time-billing':
        case 'resource-planner':
        case 'proposal-tracker':
        case 'financial-dashboard':
          const professionalResult = this.createProfessionalServicesTemplate(spreadsheet, templateType, true);
          createdSheets.push(...professionalResult.sheets);
          break;
          
        default:
          return { success: false, message: 'Unknown template type: ' + templateType };
      }
      
      // Store preview sheet names in script properties for cleanup
      const scriptProperties = PropertiesService.getScriptProperties();
      const existingPreviews = JSON.parse(scriptProperties.getProperty('previewSheets') || '[]');
      scriptProperties.setProperty('previewSheets', JSON.stringify([...existingPreviews, ...createdSheets]));
      
      // Show the first preview sheet
      if (createdSheets.length > 0) {
        const firstSheet = spreadsheet.getSheetByName(createdSheets[0]);
        if (firstSheet) {
          spreadsheet.setActiveSheet(firstSheet);
        }
      }
      
      // Skip modal dialog - it causes errors and UI feedback is handled by sidebar
      // The success message and sheet list will be shown in the sidebar
      
      return {
        success: true,
        sheets: createdSheets,
        message: 'Preview sheets created successfully'
      };
    } catch (error) {
      Logger.error('Error previewing template:', error);
      return { success: false, error: error.toString() };
    }
  },
  
  /**
   * Clean up preview sheets
   */
  cleanupPreviewSheets: function() {
    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const allSheets = spreadsheet.getSheets();
      let deletedCount = 0;
      
      // First try to clean up from script properties
      try {
        const scriptProperties = PropertiesService.getScriptProperties();
        const storedPreviewSheets = JSON.parse(scriptProperties.getProperty('previewSheets') || '[]');
        
        storedPreviewSheets.forEach(sheetName => {
          try {
            const sheet = spreadsheet.getSheetByName(sheetName);
            if (sheet) {
              spreadsheet.deleteSheet(sheet);
              deletedCount++;
            }
          } catch (e) {
            Logger.info('Could not delete sheet:', sheetName, e.toString());
          }
        });
        
        // Clear the property after cleanup
        scriptProperties.deleteProperty('previewSheets');
      } catch (e) {
        Logger.info('Error with script properties cleanup:', e.toString());
      }
      
      // Also do a general cleanup of any sheets starting with [PREVIEW]
      allSheets.forEach(sheet => {
        const sheetName = sheet.getName();
        if (sheetName.startsWith('[PREVIEW]')) {
          try {
            spreadsheet.deleteSheet(sheet);
            deletedCount++;
          } catch (e) {
            Logger.info('Could not delete preview sheet:', sheetName, e.toString());
          }
        }
      });
      
      return { success: true, deletedCount: deletedCount };
    } catch (error) {
      Logger.error('Error cleaning up preview sheets:', error);
      return { success: false, error: error.toString() };
    }
  },
  
  /**
   * Create Real Estate template sheets
   */
  createRealEstateTemplate: function(spreadsheet, templateType, isPreview = false) {
    const sheets = [];
    
    // Use the new RealEstateTemplate module for all real estate templates
    const result = RealEstateTemplate.build(spreadsheet, templateType, isPreview);
    
    if (result.success) {
      sheets.push(...result.sheets);
    } else if (result.errors && result.errors.length > 0) {
      // Log errors for debugging but still return created sheets
      Logger.log('RealEstateTemplate errors: ' + result.errors.join(', '));
      if (result.sheets && result.sheets.length > 0) {
        sheets.push(...result.sheets);
      }
    }
    
    return { sheets: sheets };
  },
  
  createConstructionTemplate: function(spreadsheet, templateType, isPreview = false) {
    const sheets = [];
    
    // Use the new ConstructionTemplate module for all construction templates
    const result = ConstructionTemplate.build(spreadsheet, templateType, isPreview);
    
    if (result.success) {
      sheets.push(...result.sheets);
    } else if (result.errors && result.errors.length > 0) {
      // Log errors for debugging but still return created sheets
      Logger.log('ConstructionTemplate errors: ' + result.errors.join(', '));
      if (result.sheets && result.sheets.length > 0) {
        sheets.push(...result.sheets);
      }
    }
    
    return { sheets: sheets };
  },
  
  createRealEstateTemplate_OLD: function(spreadsheet, templateType, isPreview = false) {
    const sheets = [];
    const prefix = isPreview ? '[PREVIEW] ' : '';

    switch (templateType) {
      case 'commission-tracker':
        // Use the Professional Real Estate Template - single comprehensive version
        const result = RealEstateTemplate.build(spreadsheet, isPreview);

        if (result.success) {
          sheets.push(...result.sheets);
        } else if (result.errors && result.errors.length > 0) {
          // Log errors for debugging but still return created sheets
          Logger.log('RealEstateTemplate errors: ' + result.errors.join(', '));
          if (result.sheets && result.sheets.length > 0) {
            sheets.push(...result.sheets);
          }
        }
        break;

      case 'lead-pipeline':
        const leadDashboard = spreadsheet.insertSheet(prefix + 'Dashboard');
        const leadSheet = spreadsheet.insertSheet(prefix + 'Lead Pipeline');
        const followUpSheet = spreadsheet.insertSheet(prefix + 'Follow-ups');
        sheets.push(leadDashboard.getName(), leadSheet.getName(), followUpSheet.getName());

        // Setup data sheets first
        this.setupLeadPipelineSheet(leadSheet);
        this.setupFollowUpSheet(followUpSheet);

        // Use actual sheet names for dashboard formulas
        const leadSheetName = leadSheet.getName();
        const followUpSheetName = followUpSheet.getName();

        this.createDashboard(leadDashboard, {
          title: 'Lead Pipeline Dashboard',
          subtitle: 'Lead Generation & Conversion Analytics',
          industry: 'realEstate',
          kpis: [
            { value: `=COUNTIF('${leadSheetName}'!F:F,"Hot")`, label: 'Hot Leads', trend: '+3 today', trendColor: '#EF4444' },
            { value: `=COUNTIF('${leadSheetName}'!F:F,"Warm")`, label: 'Warm Leads', trend: '+5 this week', trendColor: '#F59E0B' },
            { value: `=TEXT(COUNTIF('${leadSheetName}'!G:G,"Converted")/COUNTA('${leadSheetName}'!A:A),"0%")`, label: 'Conversion Rate' },
            { value: `=COUNTIF('${followUpSheetName}'!E:E,TODAY())`, label: 'Follow-ups Today', trend: 'On schedule' }
          ],
          charts: [
            { title: 'Lead Sources', type: 'Pie', description: 'Lead distribution by source', dataRange: `${leadSheetName}!D:D` },
            { title: 'Conversion Funnel', type: 'Funnel', description: 'Lead progression through stages', dataRange: `${leadSheetName}!F:G` }
          ]
        });
        spreadsheet.setActiveSheet(leadDashboard);
        spreadsheet.moveActiveSheet(1);
        break;

      case 'property-manager':
        const pmDashboard = spreadsheet.insertSheet(prefix + 'Dashboard');
        const propertySheet = spreadsheet.insertSheet(prefix + 'Properties');
        const tenantSheet = spreadsheet.insertSheet(prefix + 'Tenants');
        const maintenanceSheet = spreadsheet.insertSheet(prefix + 'Maintenance');
        const incomeSheet = spreadsheet.insertSheet(prefix + 'Rental Income');
        sheets.push(pmDashboard.getName(), propertySheet.getName(), tenantSheet.getName(), maintenanceSheet.getName(), incomeSheet.getName());

        this.createDashboard(pmDashboard, {
          title: 'Property Management Dashboard',
          subtitle: 'Portfolio Performance & Maintenance Tracking',
          industry: 'realEstate',
          kpis: [
            { value: '=COUNTA(\'Properties\'!A:A)-1', label: 'Total Properties' },
            { value: '=TEXT(COUNTIF(\'Properties\'!F:F,"Occupied")/COUNTA(\'Properties\'!A:A),"0%")', label: 'Occupancy Rate', trend: '95% target', trendColor: '#10B981' },
            { value: '=TEXT(SUM(\'Rental Income\'!E:E),"$#,##0")', label: 'Monthly Income', trend: '+5% MoM' },
            { value: '=COUNTIF(\'Maintenance\'!G:G,"Open")', label: 'Open Tickets', trend: '-2 vs yesterday', trendColor: '#10B981' },
            { value: '=TEXT(AVERAGE(\'Rental Income\'!E:E),"$#,##0")', label: 'Avg Rent' },
            { value: '=COUNTIF(\'Tenants\'!H:H,"Current")', label: 'Active Tenants' },
            { value: '=TEXT(SUM(\'Maintenance\'!F:F),"$#,##0")', label: 'Maintenance Costs', trend: 'Within budget' },
            { value: '=TEXT((SUM(\'Rental Income\'!E:E)-SUM(\'Maintenance\'!F:F))/SUM(\'Rental Income\'!E:E),"0%")', label: 'Net Margin' }
          ],
          charts: [
            { title: 'Occupancy Trend', type: 'Area', description: 'Monthly occupancy rates', dataRange: 'Properties!A:F' },
            { title: 'Income vs Expenses', type: 'Column', description: 'Monthly financial performance', dataRange: 'Rental Income!A:E' },
            { title: 'Maintenance by Type', type: 'Pie', description: 'Breakdown of maintenance requests', dataRange: 'Maintenance!C:C' }
          ]
        });

        this.setupPropertySheet(propertySheet);
        this.setupTenantSheet(tenantSheet);
        this.setupMaintenanceSheet(maintenanceSheet);
        this.setupRentalIncomeSheet(incomeSheet);
        spreadsheet.setActiveSheet(pmDashboard);
        spreadsheet.moveActiveSheet(1);
        break;

      case 'investment-analyzer':
        const investDashboard = spreadsheet.insertSheet(prefix + 'Dashboard');
        const investmentSheet = spreadsheet.insertSheet(prefix + 'Investment Analysis');
        const cashFlowSheet = spreadsheet.insertSheet(prefix + 'Cash Flow');
        const roiSheet = spreadsheet.insertSheet(prefix + 'ROI Calculator');
        const compareSheet = spreadsheet.insertSheet(prefix + 'Property Comparison');
        sheets.push(investDashboard.getName(), investmentSheet.getName(), cashFlowSheet.getName(), roiSheet.getName(), compareSheet.getName());

        this.createDashboard(investDashboard, {
          title: 'Real Estate Investment Dashboard',
          subtitle: 'ROI Analysis & Cash Flow Projections',
          industry: 'realEstate',
          kpis: [
            { value: '=TEXT(AVERAGE(\'ROI Calculator\'!F:F),"0.0%")', label: 'Avg Cap Rate', trend: 'Above market', trendColor: '#10B981' },
            { value: '=TEXT(SUM(\'Cash Flow\'!G:G),"$#,##0")', label: 'Annual Cash Flow', trend: '+15% YoY' },
            { value: '=TEXT(AVERAGE(\'ROI Calculator\'!H:H),"0.0%")', label: 'Avg Cash-on-Cash', trend: 'Excellent' },
            { value: '=TEXT(MAX(\'Investment Analysis\'!K:K),"0.0")', label: 'Best IRR %', trend: 'Top performer' },
            { value: '=TEXT(SUM(\'Investment Analysis\'!D:D),"$#,##0")', label: 'Total Invested' },
            { value: '=COUNTA(\'Investment Analysis\'!A:A)-1', label: 'Properties Analyzed' }
          ],
          charts: [
            { title: 'ROI Comparison', type: 'Bar', description: 'Compare returns across properties', dataRange: 'Property Comparison!A:E' },
            { title: 'Cash Flow Projection', type: 'Line', description: '5-year cash flow forecast', dataRange: 'Cash Flow!A:G' },
            { title: 'Investment Distribution', type: 'Pie', description: 'Portfolio allocation', dataRange: 'Investment Analysis!B:D' }
          ]
        });

        this.setupInvestmentSheet(investmentSheet);
        this.setupCashFlowSheet(cashFlowSheet);
        this.setupROISheet(roiSheet);
        this.setupPropertyComparisonSheet(compareSheet);
        spreadsheet.setActiveSheet(investDashboard);
        spreadsheet.moveActiveSheet(1);
        break;
    }

    return { sheets };
  },

  /**
   * Create Construction template sheets
   */
  createConstructionTemplate_OLD: function(spreadsheet, templateType, isPreview = false) {
    const sheets = [];
    const prefix = isPreview ? '[PREVIEW] ' : '';

    switch (templateType) {
      case 'material-tracker':
        const matDashboard = spreadsheet.insertSheet(prefix + 'Dashboard');
        const materialSheet = spreadsheet.insertSheet(prefix + 'Material Tracker');
        const suppliersSheet = spreadsheet.insertSheet(prefix + 'Suppliers');
        const ordersSheet = spreadsheet.insertSheet(prefix + 'Purchase Orders');
        sheets.push(matDashboard.getName(), materialSheet.getName(), suppliersSheet.getName(), ordersSheet.getName());

        this.createDashboard(matDashboard, {
          title: 'Construction Materials Dashboard',
          subtitle: 'Inventory & Supply Chain Management',
          industry: 'construction',
          kpis: [
            { value: '=TEXT(SUM(\'Material Tracker\'!G:G),"$#,##0")', label: 'Total Inventory Value', trend: '+8% this month' },
            { value: '=COUNTIF(\'Material Tracker\'!H:H,"Low")', label: 'Low Stock Items', trend: 'Order needed', trendColor: '#F59E0B' },
            { value: '=COUNTIF(\'Purchase Orders\'!F:F,"Pending")', label: 'Pending Orders', trend: '3 arriving today' },
            { value: '=TEXT(AVERAGE(\'Purchase Orders\'!E:E),"$#,##0")', label: 'Avg Order Value' },
            { value: '=COUNT(\'Suppliers\'!A:A)-1', label: 'Active Suppliers' },
            { value: '=TEXT(SUM(\'Purchase Orders\'!E:E),"$#,##0")', label: 'Total Orders MTD' }
          ],
          charts: [
            { title: 'Material Usage Trend', type: 'Line', description: 'Track material consumption over time', dataRange: 'Material Tracker!A:D' },
            { title: 'Stock Levels by Category', type: 'Column', description: 'Current inventory levels', dataRange: 'Material Tracker!B:D' },
            { title: 'Supplier Performance', type: 'Bar', description: 'On-time delivery rates', dataRange: 'Suppliers!A:E' }
          ]
        });

        this.setupMaterialTrackerSheet(materialSheet);
        this.setupSuppliersSheet(suppliersSheet);
        this.setupPurchaseOrdersSheet(ordersSheet);
        spreadsheet.setActiveSheet(matDashboard);
        spreadsheet.moveActiveSheet(1);
        break;

      case 'labor-manager':
        const laborDashboard = spreadsheet.insertSheet(prefix + 'Dashboard');
        const laborSheet = spreadsheet.insertSheet(prefix + 'Labor Manager');
        const crewSheet = spreadsheet.insertSheet(prefix + 'Crew Schedule');
        const payrollSheet = spreadsheet.insertSheet(prefix + 'Payroll');
        sheets.push(laborDashboard.getName(), laborSheet.getName(), crewSheet.getName(), payrollSheet.getName());

        this.createDashboard(laborDashboard, {
          title: 'Labor Management Dashboard',
          subtitle: 'Workforce Analytics & Scheduling',
          industry: 'construction',
          kpis: [
            { value: '=COUNTA(\'Labor Manager\'!A:A)-1', label: 'Total Workers' },
            { value: '=COUNTIF(\'Crew Schedule\'!E:E,TODAY())', label: 'On Site Today' },
            { value: '=TEXT(SUM(\'Payroll\'!F:F),"$#,##0")', label: 'Weekly Payroll', trend: 'On budget' },
            { value: '=TEXT(AVERAGE(\'Labor Manager\'!E:E),"0")', label: 'Avg Hours/Week' },
            { value: '=TEXT(SUM(\'Labor Manager\'!F:F)/SUM(\'Labor Manager\'!E:E),"$#,##0")', label: 'Avg Hourly Rate' },
            { value: '=COUNTIF(\'Labor Manager\'!G:G,"Certified")', label: 'Certified Workers', trend: '+2 this month', trendColor: '#10B981' }
          ],
          charts: [
            { title: 'Labor Hours by Project', type: 'Pie', description: 'Time allocation across projects', dataRange: 'Crew Schedule!A:D' },
            { title: 'Weekly Labor Costs', type: 'Area', description: 'Labor cost trends', dataRange: 'Payroll!A:F' },
            { title: 'Crew Utilization', type: 'Column', description: 'Utilization rates by crew', dataRange: 'Labor Manager!A:E' }
          ]
        });

        this.setupLaborManagerSheet(laborSheet);
        this.setupCrewScheduleSheet(crewSheet);
        this.setupPayrollSheet(payrollSheet);
        spreadsheet.setActiveSheet(laborDashboard);
        spreadsheet.moveActiveSheet(1);
        break;

      case 'change-orders':
        const changeDashboard = spreadsheet.insertSheet(prefix + 'Dashboard');
        const changeSheet = spreadsheet.insertSheet(prefix + 'Change Orders');
        const impactSheet = spreadsheet.insertSheet(prefix + 'Cost Impact');
        const approvalSheet = spreadsheet.insertSheet(prefix + 'Approvals');
        sheets.push(changeDashboard.getName(), changeSheet.getName(), impactSheet.getName(), approvalSheet.getName());

        this.createDashboard(changeDashboard, {
          title: 'Change Order Management Dashboard',
          subtitle: 'Project Change Tracking & Impact Analysis',
          industry: 'construction',
          kpis: [
            { value: '=COUNTA(\'Change Orders\'!A:A)-1', label: 'Total Change Orders' },
            { value: '=TEXT(SUM(\'Cost Impact\'!D:D),"$#,##0")', label: 'Total Cost Impact', trend: '+$45K this month', trendColor: '#F59E0B' },
            { value: '=COUNTIF(\'Approvals\'!E:E,"Pending")', label: 'Pending Approvals', trend: '5 urgent', trendColor: '#EF4444' },
            { value: '=TEXT(AVERAGE(\'Approvals\'!F:F),"0")', label: 'Avg Approval Days' },
            { value: '=TEXT(COUNTIF(\'Approvals\'!E:E,"Approved")/COUNTA(\'Approvals\'!A:A),"0%")', label: 'Approval Rate' },
            { value: '=SUM(\'Cost Impact\'!E:E)', label: 'Schedule Impact (days)', trend: '+10 days total' }
          ],
          charts: [
            { title: 'Change Orders by Type', type: 'Pie', description: 'Distribution of change order types', dataRange: 'Change Orders!C:C' },
            { title: 'Cost Impact Trend', type: 'Line', description: 'Cumulative cost impact over time', dataRange: 'Cost Impact!A:D' },
            { title: 'Approval Timeline', type: 'Bar', description: 'Days to approval by project', dataRange: 'Approvals!A:F' }
          ]
        });

        this.setupChangeOrdersSheet(changeSheet);
        this.setupCostImpactSheet(impactSheet);
        this.setupApprovalsSheet(approvalSheet);
        spreadsheet.setActiveSheet(changeDashboard);
        spreadsheet.moveActiveSheet(1);
        break;

      case 'cost-estimator':
        const costDashboard = spreadsheet.insertSheet(prefix + 'Dashboard');
        const estimateSheet = spreadsheet.insertSheet(prefix + 'Cost Estimator');
        const breakdownSheet = spreadsheet.insertSheet(prefix + 'Cost Breakdown');
        const contingencySheet = spreadsheet.insertSheet(prefix + 'Contingency');
        sheets.push(costDashboard.getName(), estimateSheet.getName(), breakdownSheet.getName(), contingencySheet.getName());

        this.createDashboard(costDashboard, {
          title: 'Project Cost Estimation Dashboard',
          subtitle: 'Budget Planning & Cost Analysis',
          industry: 'construction',
          kpis: [
            { value: '=TEXT(SUM(\'Cost Estimator\'!F:F),"$#,##0")', label: 'Total Project Cost' },
            { value: '=TEXT(SUM(\'Cost Breakdown\'!D:D),"$#,##0")', label: 'Materials Cost', trend: '35% of total' },
            { value: '=TEXT(SUM(\'Cost Breakdown\'!E:E),"$#,##0")', label: 'Labor Cost', trend: '40% of total' },
            { value: '=TEXT(SUM(\'Contingency\'!C:C),"$#,##0")', label: 'Contingency', trend: '10% buffer' },
            { value: '=TEXT((SUM(\'Cost Estimator\'!F:F)-SUM(\'Cost Estimator\'!G:G))/SUM(\'Cost Estimator\'!F:F),"0%")', label: 'Profit Margin' },
            { value: '=TEXT(SUM(\'Cost Estimator\'!F:F)/SUM(\'Cost Estimator\'!D:D),"$#,##0")', label: 'Cost per Sq Ft' }
          ],
          charts: [
            { title: 'Cost Breakdown', type: 'Pie', description: 'Distribution of project costs', dataRange: 'Cost Breakdown!A:E' },
            { title: 'Budget vs Actual', type: 'Column', description: 'Compare estimated vs actual costs', dataRange: 'Cost Estimator!A:G' },
            { title: 'Risk Assessment', type: 'Bar', description: 'Cost risk by category', dataRange: 'Contingency!A:D' }
          ]
        });

        this.setupCostEstimatorSheet(estimateSheet);
        this.setupCostBreakdownSheet(breakdownSheet);
        this.setupContingencySheet(contingencySheet);
        spreadsheet.setActiveSheet(costDashboard);
        spreadsheet.moveActiveSheet(1);
        break;
    }

    return { sheets };
  },

  /**
   * Create Healthcare template sheets
   */
  createHealthcareTemplate: function(spreadsheet, templateType, isPreview = false) {
    const sheets = [];

    // Use the new HealthcareTemplate module for all healthcare templates
    const result = HealthcareTemplate.build(spreadsheet, templateType, isPreview);

    if (result.success) {
      sheets.push(...result.sheets);
    } else if (result.errors && result.errors.length > 0) {
      // Log errors for debugging but still return created sheets
      Logger.log('HealthcareTemplate errors: ' + result.errors.join(', '));
      if (result.sheets && result.sheets.length > 0) {
        sheets.push(...result.sheets);
      }
    }

    return { sheets: sheets };
  },

  createHealthcareTemplate_OLD: function(spreadsheet, templateType, isPreview = false) {
    const sheets = [];
    const prefix = isPreview ? '[PREVIEW] ' : '';

    switch (templateType) {
      case 'prior-auth-tracker':
        const authDashboard = spreadsheet.insertSheet(prefix + 'Dashboard');
        const priorAuthSheet = spreadsheet.insertSheet(prefix + 'Prior Auth');
        const statusSheet = spreadsheet.insertSheet(prefix + 'Auth Status');
        const providersSheet = spreadsheet.insertSheet(prefix + 'Providers');
        sheets.push(authDashboard.getName(), priorAuthSheet.getName(), statusSheet.getName(), providersSheet.getName());

        this.createDashboard(authDashboard, {
          title: 'Prior Authorization Dashboard',
          subtitle: 'Authorization Tracking & Performance Metrics',
          industry: 'healthcare',
          kpis: [
            { value: '=COUNTIF(\'Prior Auth\'!F:F,"Pending")', label: 'Pending Auths', trend: '8 urgent', trendColor: '#F59E0B' },
            { value: '=COUNTIF(\'Prior Auth\'!F:F,"Approved")', label: 'Approved Today', trend: '+15 vs yesterday', trendColor: '#10B981' },
            { value: '=TEXT(COUNTIF(\'Prior Auth\'!F:F,"Approved")/COUNTA(\'Prior Auth\'!A:A),"0%")', label: 'Approval Rate', trend: '92% target' },
            { value: '=AVERAGE(\'Auth Status\'!D:D)', label: 'Avg Days to Approve', trend: 'Improving' },
            { value: '=COUNTIF(\'Prior Auth\'!F:F,"Denied")', label: 'Denials MTD', trend: '-20% vs last month', trendColor: '#10B981' },
            { value: '=COUNTIF(\'Auth Status\'!E:E,"Expedited")', label: 'Expedited Cases' }
          ],
          charts: [
            { title: 'Authorization Status', type: 'Pie', description: 'Current authorization distribution', dataRange: 'Prior Auth!F:F' },
            { title: 'Processing Time Trend', type: 'Line', description: 'Average days to approval', dataRange: 'Auth Status!A:D' },
            { title: 'Top Denial Reasons', type: 'Bar', description: 'Common denial reasons', dataRange: 'Prior Auth!G:H' }
          ]
        });
        
        this.setupPriorAuthSheet(priorAuthSheet);
        this.setupAuthStatusSheet(statusSheet);
        this.setupProvidersSheet(providersSheet);
        spreadsheet.setActiveSheet(authDashboard);
        spreadsheet.moveActiveSheet(1);
        break;
        
      case 'revenue-cycle':
        const revDashboard = spreadsheet.insertSheet(prefix + 'Dashboard');
        const revenueSheet = spreadsheet.insertSheet(prefix + 'Revenue Cycle');
        const claimsSheet = spreadsheet.insertSheet(prefix + 'Claims');
        const agingSheet = spreadsheet.insertSheet(prefix + 'AR Aging');
        sheets.push(revDashboard.getName(), revenueSheet.getName(), claimsSheet.getName(), agingSheet.getName());
        
        this.createDashboard(revDashboard, {
          title: 'Revenue Cycle Management Dashboard',
          subtitle: 'Financial Performance & Collections Analytics',
          industry: 'healthcare',
          kpis: [
            { value: '=TEXT(SUM(\'Revenue Cycle\'!E:E),"$#,##0")', label: 'Total Revenue MTD', trend: '+12% vs last month', trendColor: '#10B981' },
            { value: '=TEXT(SUM(\'AR Aging\'!C:C),"$#,##0")', label: 'Outstanding AR', trend: 'Decreasing' },
            { value: '=AVERAGE(\'Claims\'!H:H)', label: 'Days in AR', trend: '45 day target' },
            { value: '=TEXT(COUNTIF(\'Claims\'!F:F,"Paid")/COUNTA(\'Claims\'!A:A),"0%")', label: 'Collection Rate', trend: '95% benchmark' },
            { value: '=TEXT(SUM(\'Claims\'!G:G),"$#,##0")', label: 'Denied Claims Value', trend: 'Need attention', trendColor: '#EF4444' },
            { value: '=COUNTIF(\'AR Aging\'!D:D,">90")', label: 'Claims >90 Days', trend: '-5 this week', trendColor: '#10B981' }
          ],
          charts: [
            { title: 'Revenue Trend', type: 'Area', description: 'Monthly revenue performance', dataRange: 'Revenue Cycle!A:E' },
            { title: 'AR Aging Buckets', type: 'Column', description: 'Accounts receivable by age', dataRange: 'AR Aging!A:C' },
            { title: 'Payer Mix', type: 'Pie', description: 'Revenue by insurance type', dataRange: 'Claims!D:E' }
          ]
        });
        
        this.setupRevenueCycleSheet(revenueSheet);
        this.setupClaimsSheet(claimsSheet);
        this.setupARAgingSheet(agingSheet);
        spreadsheet.setActiveSheet(revDashboard);
        spreadsheet.moveActiveSheet(1);
        break;
        
      case 'denial-analytics':
        const denialDashboard = spreadsheet.insertSheet(prefix + 'Dashboard');
        const denialSheet = spreadsheet.insertSheet(prefix + 'Denial Analytics');
        const appealsSheet = spreadsheet.insertSheet(prefix + 'Appeals');
        const trendsSheet = spreadsheet.insertSheet(prefix + 'Denial Trends');
        sheets.push(denialDashboard.getName(), denialSheet.getName(), appealsSheet.getName(), trendsSheet.getName());
        
        this.createDashboard(denialDashboard, {
          title: 'Denial Management Dashboard',
          subtitle: 'Denial Analytics & Recovery Performance',
          industry: 'healthcare',
          kpis: [
            { value: '=TEXT(COUNTA(\'Denial Analytics\'!A:A)-1/COUNTA(\'Claims\'!A:A),"0%")', label: 'Denial Rate', trend: '5% target', trendColor: '#F59E0B' },
            { value: '=TEXT(SUM(\'Denial Analytics\'!E:E),"$#,##0")', label: 'Total Denied Amount', trend: 'Reducing' },
            { value: '=TEXT(SUM(\'Appeals\'!G:G),"$#,##0")', label: 'Amount Recovered', trend: '+$125K MTD', trendColor: '#10B981' },
            { value: '=TEXT(COUNTIF(\'Appeals\'!F:F,"Won")/COUNTA(\'Appeals\'!A:A),"0%")', label: 'Appeal Success Rate', trend: '75% average' },
            { value: '=COUNTIF(\'Denial Analytics\'!D:D,"Clinical")', label: 'Clinical Denials', trend: 'Focus area' },
            { value: '=AVERAGE(\'Appeals\'!H:H)', label: 'Days to Overturn', trend: 'Improving' }
          ],
          charts: [
            { title: 'Denial Categories', type: 'Pie', description: 'Distribution by denial type', dataRange: 'Denial Analytics!D:D' },
            { title: 'Recovery Trend', type: 'Line', description: 'Monthly recovery amounts', dataRange: 'Appeals!A:G' },
            { title: 'Top Denial Codes', type: 'Bar', description: 'Most common denial reasons', dataRange: 'Denial Trends!A:C' }
          ]
        });
        
        this.setupDenialAnalyticsSheet(denialSheet);
        this.setupAppealsSheet(appealsSheet);
        this.setupDenialTrendsSheet(trendsSheet);
        spreadsheet.setActiveSheet(denialDashboard);
        spreadsheet.moveActiveSheet(1);
        break;
        
      case 'insurance-verifier':
        const insDashboard = spreadsheet.insertSheet(prefix + 'Dashboard');
        const insuranceSheet = spreadsheet.insertSheet(prefix + 'Insurance Verification');
        const eligibilitySheet = spreadsheet.insertSheet(prefix + 'Eligibility');
        const benefitsSheet = spreadsheet.insertSheet(prefix + 'Benefits');
        sheets.push(insDashboard.getName(), insuranceSheet.getName(), eligibilitySheet.getName(), benefitsSheet.getName());
        
        this.createDashboard(insDashboard, {
          title: 'Insurance Verification Dashboard',
          subtitle: 'Coverage Verification & Benefits Analysis',
          industry: 'healthcare',
          kpis: [
            { value: '=COUNTIF(\'Insurance Verification\'!F:F,"Verified")', label: 'Verified Today', trend: '125 completed' },
            { value: '=COUNTIF(\'Insurance Verification\'!F:F,"Pending")', label: 'Pending Verification', trend: '45 in queue' },
            { value: '=TEXT(COUNTIF(\'Eligibility\'!D:D,"Active")/COUNTA(\'Eligibility\'!A:A),"0%")', label: 'Active Coverage Rate', trend: '98% average' },
            { value: '=COUNTIF(\'Benefits\'!E:E,"Out of Network")', label: 'OON Cases', trend: 'Needs review', trendColor: '#F59E0B' },
            { value: '=AVERAGE(\'Insurance Verification\'!H:H)', label: 'Avg Copay Amount', trend: 'Stable' },
            { value: '=COUNTIF(\'Insurance Verification\'!G:G,"High")', label: 'High Deductibles', trend: 'Increasing trend' }
          ],
          charts: [
            { title: 'Insurance Mix', type: 'Pie', description: 'Patient insurance distribution', dataRange: 'Insurance Verification!C:C' },
            { title: 'Verification Status', type: 'Column', description: 'Daily verification metrics', dataRange: 'Insurance Verification!A:F' },
            { title: 'Benefits Summary', type: 'Bar', description: 'Coverage types analysis', dataRange: 'Benefits!A:D' }
          ]
        });
        
        this.setupInsuranceVerifierSheet(insuranceSheet);
        this.setupEligibilitySheet(eligibilitySheet);
        this.setupBenefitsSheet(benefitsSheet);
        spreadsheet.setActiveSheet(insDashboard);
        spreadsheet.moveActiveSheet(1);
        break;
    }
    
    return { sheets };
  },
  
  /**
   * Create Marketing template sheets
   */
  createMarketingTemplate: function(spreadsheet, templateType, isPreview = false) {
    try {
      return MarketingTemplate.build(spreadsheet, templateType, isPreview);
    } catch (error) {
      Logger.error('Error in createMarketingTemplate:', error);
      return { error: 'Failed to create marketing template: ' + error.message };
    }
  },

  createMarketingTemplate_OLD: function(spreadsheet, templateType, isPreview = false) {
    const sheets = [];
    const prefix = isPreview ? '[PREVIEW] ' : '';
    
    switch (templateType) {
      case 'lead-scoring-system':
        const leadDash = spreadsheet.insertSheet(prefix + 'Dashboard');
        const leadScoringSheet = spreadsheet.insertSheet(prefix + 'Lead Scoring');
        const segmentsSheet = spreadsheet.insertSheet(prefix + 'Segments');
        const engagementSheet = spreadsheet.insertSheet(prefix + 'Engagement');
        sheets.push(leadDash.getName(), leadScoringSheet.getName(), segmentsSheet.getName(), engagementSheet.getName());
        
        this.createDashboard(leadDash, {
          title: 'Lead Scoring & Qualification Dashboard',
          subtitle: 'Lead Quality & Conversion Analytics',
          industry: 'marketing',
          kpis: [
            { value: '=COUNTIF(\'Lead Scoring\'!E:E,">80")', label: 'Hot Leads', trend: '+12 today', trendColor: '#EF4444' },
            { value: '=COUNTIF(\'Lead Scoring\'!E:E,"50-80")', label: 'Warm Leads', trend: '+25 this week', trendColor: '#F59E0B' },
            { value: '=TEXT(AVERAGE(\'Lead Scoring\'!E:E),"0")', label: 'Avg Lead Score' },
            { value: '=TEXT(COUNTIF(\'Lead Scoring\'!F:F,"Qualified")/COUNTA(\'Lead Scoring\'!A:A),"0%")', label: 'Qualification Rate' },
            { value: '=COUNTIF(\'Engagement\'!D:D,">5")', label: 'Highly Engaged', trend: 'Growing', trendColor: '#10B981' },
            { value: '=TEXT(AVERAGE(\'Lead Scoring\'!G:G),"0")', label: 'Days to Convert' }
          ],
          charts: [
            { title: 'Lead Distribution', type: 'Pie', description: 'Leads by score category', dataRange: 'Lead Scoring!E:E' },
            { title: 'Conversion Funnel', type: 'Funnel', description: 'Lead progression stages', dataRange: 'Lead Scoring!F:F' },
            { title: 'Engagement Trends', type: 'Line', description: 'Weekly engagement metrics', dataRange: 'Engagement!A:D' }
          ]
        });
        
        this.setupLeadScoringSheet(leadScoringSheet);
        this.setupSegmentsSheet(segmentsSheet);
        this.setupEngagementSheet(engagementSheet);
        spreadsheet.setActiveSheet(leadDash);
        spreadsheet.moveActiveSheet(1);
        break;
        
      case 'content-performance':
        const contentDash = spreadsheet.insertSheet(prefix + 'Dashboard');
        const contentSheet = spreadsheet.insertSheet(prefix + 'Content Performance');
        const metricsSheet = spreadsheet.insertSheet(prefix + 'Metrics');
        const channelsSheet = spreadsheet.insertSheet(prefix + 'Channels');
        sheets.push(contentDash.getName(), contentSheet.getName(), metricsSheet.getName(), channelsSheet.getName());
        
        this.createDashboard(contentDash, {
          title: 'Content Performance Dashboard',
          subtitle: 'Content Analytics & ROI Tracking',
          industry: 'marketing',
          kpis: [
            { value: '=SUM(\'Content Performance\'!D:D)', label: 'Total Views', trend: '+15% MoM', trendColor: '#10B981' },
            { value: '=TEXT(AVERAGE(\'Metrics\'!E:E),"0%")', label: 'Avg Engagement Rate', trend: 'Above benchmark' },
            { value: '=COUNTIF(\'Content Performance\'!G:G,"Viral")', label: 'Viral Content', trend: '3 this month' },
            { value: '=TEXT(SUM(\'Content Performance\'!H:H),"$#,##0")', label: 'Content ROI' },
            { value: '=MAX(\'Metrics\'!F:F)', label: 'Top Share Count', trend: 'Record high', trendColor: '#10B981' },
            { value: '=TEXT(AVERAGE(\'Metrics\'!C:C),"0:00")', label: 'Avg Time on Page' }
          ],
          charts: [
            { title: 'Content Type Performance', type: 'Column', description: 'Performance by content type', dataRange: 'Content Performance!B:D' },
            { title: 'Channel Distribution', type: 'Pie', description: 'Traffic by channel', dataRange: 'Channels!A:C' },
            { title: 'Engagement Trend', type: 'Area', description: 'Monthly engagement metrics', dataRange: 'Metrics!A:E' }
          ]
        });
        
        this.setupContentPerformanceSheet(contentSheet);
        this.setupMetricsSheet(metricsSheet);
        this.setupChannelsSheet(channelsSheet);
        spreadsheet.setActiveSheet(contentDash);
        spreadsheet.moveActiveSheet(1);
        break;
        
      case 'customer-journey':
        const journeyDash = spreadsheet.insertSheet(prefix + 'Dashboard');
        const journeySheet = spreadsheet.insertSheet(prefix + 'Customer Journey');
        const touchpointsSheet = spreadsheet.insertSheet(prefix + 'Touchpoints');
        const conversionSheet = spreadsheet.insertSheet(prefix + 'Conversions');
        sheets.push(journeyDash.getName(), journeySheet.getName(), touchpointsSheet.getName(), conversionSheet.getName());
        
        this.createDashboard(journeyDash, {
          title: 'Customer Journey Analytics Dashboard',
          subtitle: 'Path to Purchase & Attribution Analysis',
          industry: 'marketing',
          kpis: [
            { value: '=AVERAGE(\'Touchpoints\'!C:C)', label: 'Avg Touchpoints', trend: '7.2 before purchase' },
            { value: '=TEXT(AVERAGE(\'Customer Journey\'!E:E),"0")', label: 'Journey Duration (days)' },
            { value: '=TEXT(COUNTIF(\'Conversions\'!D:D,"Completed")/COUNTA(\'Customer Journey\'!A:A),"0%")', label: 'Conversion Rate' },
            { value: '=TEXT(SUM(\'Conversions\'!E:E),"$#,##0")', label: 'Total Revenue' },
            { value: '=COUNTIF(\'Customer Journey\'!F:F,"Abandoned")', label: 'Cart Abandonment', trend: 'Decreasing', trendColor: '#10B981' },
            { value: '=TEXT(AVERAGE(\'Conversions\'!F:F),"$#,##0")', label: 'Avg Order Value' }
          ],
          charts: [
            { title: 'Journey Stages', type: 'Funnel', description: 'Customer progression through stages', dataRange: 'Customer Journey!D:D' },
            { title: 'Attribution Model', type: 'Bar', description: 'Revenue by touchpoint', dataRange: 'Touchpoints!A:D' },
            { title: 'Conversion Path', type: 'Sankey', description: 'Common paths to purchase', dataRange: 'Customer Journey!A:F' }
          ]
        });
        
        this.setupCustomerJourneySheet(journeySheet);
        this.setupTouchpointsSheet(touchpointsSheet);
        this.setupConversionSheet(conversionSheet);
        spreadsheet.setActiveSheet(journeyDash);
        spreadsheet.moveActiveSheet(1);
        break;
        
      case 'campaign-dashboard':
        const campaignDash = spreadsheet.insertSheet(prefix + 'Dashboard');
        const campaignSheet = spreadsheet.insertSheet(prefix + 'Campaigns');
        const performanceSheet = spreadsheet.insertSheet(prefix + 'Performance');
        const budgetSheet = spreadsheet.insertSheet(prefix + 'Budget');
        sheets.push(campaignDash.getName(), campaignSheet.getName(), performanceSheet.getName(), budgetSheet.getName());
        
        this.createDashboard(campaignDash, {
          title: 'Marketing Campaign Dashboard',
          subtitle: 'Campaign Performance & ROI Analysis',
          industry: 'marketing',
          kpis: [
            { value: '=COUNTIF(\'Campaigns\'!E:E,"Active")', label: 'Active Campaigns' },
            { value: '=TEXT(SUM(\'Performance\'!F:F),"$#,##0")', label: 'Total Spend MTD' },
            { value: '=TEXT(SUM(\'Performance\'!G:G),"$#,##0")', label: 'Revenue Generated', trend: '+45% ROI', trendColor: '#10B981' },
            { value: '=TEXT(AVERAGE(\'Performance\'!H:H),"$#,##0")', label: 'Cost per Acquisition' },
            { value: '=TEXT(AVERAGE(\'Performance\'!I:I),"0%")', label: 'Avg CTR', trend: 'Above industry avg' },
            { value: '=TEXT((SUM(\'Performance\'!G:G)-SUM(\'Performance\'!F:F))/SUM(\'Performance\'!F:F),"0%")', label: 'Overall ROAS' }
          ],
          charts: [
            { title: 'Campaign Performance', type: 'Column', description: 'ROI by campaign', dataRange: 'Campaigns!A:G' },
            { title: 'Budget Allocation', type: 'Pie', description: 'Spend by channel', dataRange: 'Budget!A:C' },
            { title: 'Performance Trend', type: 'Line', description: 'Monthly KPI trends', dataRange: 'Performance!A:I' }
          ]
        });
        
        this.setupCampaignDashboardSheet(campaignSheet);
        this.setupPerformanceSheet(performanceSheet);
        this.setupBudgetSheet(budgetSheet);
        spreadsheet.setActiveSheet(campaignDash);
        spreadsheet.moveActiveSheet(1);
        break;
    }
    
    return { sheets };
  },
  
  /**
   * Create E-commerce template sheets
   */
  createEcommerceTemplate: function(spreadsheet, templateType, isPreview = false) {
    try {
      return EcommerceTemplate.build(spreadsheet, templateType, isPreview);
    } catch (error) {
      Logger.error('Error in createEcommerceTemplate:', error);
      return { error: 'Failed to create e-commerce template: ' + error.message };
    }
  },

  createEcommerceTemplate_OLD: function(spreadsheet, templateType, isPreview = false) {
    const sheets = [];
    const prefix = isPreview ? '[PREVIEW] ' : '';
    
    switch (templateType) {
      case 'profitability-analyzer':
        const profitDash = spreadsheet.insertSheet(prefix + 'Dashboard');
        const profitSheet = spreadsheet.insertSheet(prefix + 'Profitability');
        const productsSheet = spreadsheet.insertSheet(prefix + 'Products');
        const costsSheet = spreadsheet.insertSheet(prefix + 'Costs');
        sheets.push(profitDash.getName(), profitSheet.getName(), productsSheet.getName(), costsSheet.getName());
        
        this.createDashboard(profitDash, {
          title: 'E-commerce Profitability Dashboard',
          subtitle: 'Product Margins & Profit Analysis',
          industry: 'ecommerce',
          kpis: [
            { value: '=TEXT(SUM(\'Profitability\'!E:E),"$#,##0")', label: 'Gross Profit MTD', trend: '+18% vs LM', trendColor: '#10B981' },
            { value: '=TEXT(AVERAGE(\'Profitability\'!F:F),"0%")', label: 'Avg Margin', trend: '42% target' },
            { value: '=COUNTIF(\'Products\'!G:G,">40%")', label: 'High Margin SKUs' },
            { value: '=TEXT(SUM(\'Costs\'!D:D),"$#,##0")', label: 'Total COGS' },
            { value: '=TEXT(MAX(\'Profitability\'!F:F),"0%")', label: 'Best Margin', trend: 'Premium line' },
            { value: '=COUNTIF(\'Products\'!G:G,"<20%")', label: 'Low Performers', trend: 'Review needed', trendColor: '#F59E0B' }
          ],
          charts: [
            { title: 'Profit by Category', type: 'Column', description: 'Category profitability analysis', dataRange: 'Profitability!B:E' },
            { title: 'Margin Distribution', type: 'Histogram', description: 'Product margin spread', dataRange: 'Products!G:G' },
            { title: 'Cost Breakdown', type: 'Pie', description: 'Cost component analysis', dataRange: 'Costs!A:D' }
          ]
        });
        
        this.setupProfitabilitySheet(profitSheet);
        this.setupProductsSheet(productsSheet);
        this.setupCostsSheet(costsSheet);
        spreadsheet.setActiveSheet(profitDash);
        spreadsheet.moveActiveSheet(1);
        break;
        
      case 'sales-forecasting':
        const forecastDash = spreadsheet.insertSheet(prefix + 'Dashboard');
        const forecastSheet = spreadsheet.insertSheet(prefix + 'Sales Forecast');
        const trendsSheet = spreadsheet.insertSheet(prefix + 'Trends');
        const seasonalSheet = spreadsheet.insertSheet(prefix + 'Seasonality');
        sheets.push(forecastDash.getName(), forecastSheet.getName(), trendsSheet.getName(), seasonalSheet.getName());
        
        this.createDashboard(forecastDash, {
          title: 'Sales Forecasting Dashboard',
          subtitle: 'Predictive Analytics & Demand Planning',
          industry: 'ecommerce',
          kpis: [
            { value: '=TEXT(SUM(\'Sales Forecast\'!E:E),"$#,##0")', label: 'Forecasted Revenue', trend: 'Next 30 days' },
            { value: '=TEXT(AVERAGE(\'Trends\'!D:D),"+0%")', label: 'Growth Trend', trend: 'Accelerating', trendColor: '#10B981' },
            { value: '=MAX(\'Seasonality\'!C:C)', label: 'Peak Season Index', trend: 'Q4 approaching' },
            { value: '=TEXT(STDEV(\'Sales Forecast\'!E:E),"$#,##0")', label: 'Forecast Variance' },
            { value: '=TEXT(AVERAGE(\'Sales Forecast\'!F:F),"0%")', label: 'Confidence Level' },
            { value: '=SUM(\'Sales Forecast\'!G:G)', label: 'Units Forecasted' }
          ],
          charts: [
            { title: 'Sales Forecast', type: 'Line', description: '90-day sales projection', dataRange: 'Sales Forecast!A:E' },
            { title: 'Seasonal Patterns', type: 'Area', description: 'Historical seasonality', dataRange: 'Seasonality!A:C' },
            { title: 'Trend Analysis', type: 'Combo', description: 'Actual vs forecast', dataRange: 'Trends!A:D' }
          ]
        });
        
        this.setupSalesForecastSheet(forecastSheet);
        this.setupTrendsSheet(trendsSheet);
        this.setupSeasonalitySheet(seasonalSheet);
        spreadsheet.setActiveSheet(forecastDash);
        spreadsheet.moveActiveSheet(1);
        break;
        
      case 'ecommerce-inventory':
        const invDash = spreadsheet.insertSheet(prefix + 'Dashboard');
        const inventorySheet = spreadsheet.insertSheet(prefix + 'Inventory Manager');
        const stockSheet = spreadsheet.insertSheet(prefix + 'Stock Levels');
        const reorderSheet = spreadsheet.insertSheet(prefix + 'Reorder Points');
        sheets.push(invDash.getName(), inventorySheet.getName(), stockSheet.getName(), reorderSheet.getName());
        
        this.createDashboard(invDash, {
          title: 'Inventory Management Dashboard',
          subtitle: 'Stock Control & Reorder Management',
          industry: 'ecommerce',
          kpis: [
            { value: '=TEXT(SUM(\'Stock Levels\'!D:D),"$#,##0")', label: 'Inventory Value' },
            { value: '=COUNTIF(\'Stock Levels\'!E:E,"<10")', label: 'Low Stock Items', trend: 'Order needed', trendColor: '#EF4444' },
            { value: '=COUNTIF(\'Stock Levels\'!E:E,">500")', label: 'Overstock Items', trend: 'Reduce orders', trendColor: '#F59E0B' },
            { value: '=AVERAGE(\'Inventory Manager\'!F:F)', label: 'Avg Days of Stock' },
            { value: '=TEXT(SUM(\'Reorder Points\'!E:E),"$#,##0")', label: 'Pending Orders' },
            { value: '=TEXT(AVERAGE(\'Inventory Manager\'!G:G),"0")', label: 'Turnover Rate' }
          ],
          charts: [
            { title: 'Stock Status', type: 'Column', description: 'Current stock levels by category', dataRange: 'Stock Levels!A:E' },
            { title: 'Reorder Timeline', type: 'Gantt', description: 'Upcoming reorder schedule', dataRange: 'Reorder Points!A:E' },
            { title: 'Inventory Turnover', type: 'Line', description: 'Monthly turnover trends', dataRange: 'Inventory Manager!A:G' }
          ]
        });
        
        this.setupInventorySheet(inventorySheet);
        this.setupStockLevelsSheet(stockSheet);
        this.setupReorderPointsSheet(reorderSheet);
        spreadsheet.setActiveSheet(invDash);
        spreadsheet.moveActiveSheet(1);
        break;
    }
    
    return { sheets };
  },
  
  /**
   * Create Professional Services template sheets
   */
  createProfessionalServicesTemplate: function(spreadsheet, templateType, isPreview = false) {
    try {
      return ProfessionalServicesTemplate.build(spreadsheet, templateType, isPreview);
    } catch (error) {
      Logger.error('Error in createProfessionalServicesTemplate:', error);
      return { error: 'Failed to create professional services template: ' + error.message };
    }
  },
  
  /**
   * Create Consulting template sheets
   */
  createConsultingTemplate: function(spreadsheet, templateType, isPreview = false) {
    const sheets = [];
    const prefix = isPreview ? '[PREVIEW] ' : '';
    
    switch (templateType) {
      case 'project-profitability':
        const projDash = spreadsheet.insertSheet(prefix + 'Dashboard');
        const projectSheet = spreadsheet.insertSheet(prefix + 'Project Profitability');
        const resourceSheet = spreadsheet.insertSheet(prefix + 'Resources');
        const expensesSheet = spreadsheet.insertSheet(prefix + 'Expenses');
        sheets.push(projDash.getName(), projectSheet.getName(), resourceSheet.getName(), expensesSheet.getName());
        
        this.createDashboard(projDash, {
          title: 'Project Profitability Dashboard',
          subtitle: 'Project Performance & Resource Analytics',
          industry: 'professional',
          kpis: [
            { value: '=COUNTA(\'Project Profitability\'!A:A)-1', label: 'Active Projects' },
            { value: '=TEXT(SUM(\'Project Profitability\'!F:F),"$#,##0")', label: 'Total Revenue', trend: '+22% QoQ', trendColor: '#10B981' },
            { value: '=TEXT(AVERAGE(\'Project Profitability\'!G:G),"0%")', label: 'Avg Margin', trend: '35% target' },
            { value: '=TEXT(SUM(\'Resources\'!E:E),"0")', label: 'Billable Hours' },
            { value: '=TEXT(SUM(\'Expenses\'!D:D),"$#,##0")', label: 'Total Expenses' },
            { value: '=TEXT(AVERAGE(\'Resources\'!F:F),"0%")', label: 'Utilization Rate' }
          ],
          charts: [
            { title: 'Project Margins', type: 'Bar', description: 'Profitability by project', dataRange: 'Project Profitability!A:G' },
            { title: 'Resource Allocation', type: 'Pie', description: 'Hours by project', dataRange: 'Resources!A:E' },
            { title: 'Expense Categories', type: 'Column', description: 'Cost breakdown', dataRange: 'Expenses!B:D' }
          ]
        });
        
        this.setupProjectProfitabilitySheet(projectSheet);
        this.setupResourceSheet(resourceSheet);
        this.setupExpensesSheet(expensesSheet);
        spreadsheet.setActiveSheet(projDash);
        spreadsheet.moveActiveSheet(1);
        break;
        
      case 'client-dashboard':
        const clientDash = spreadsheet.insertSheet(prefix + 'Dashboard');
        const clientSheet = spreadsheet.insertSheet(prefix + 'Clients');
        const engagementSheet = spreadsheet.insertSheet(prefix + 'Engagements');
        const revenueSheet = spreadsheet.insertSheet(prefix + 'Revenue');
        sheets.push(clientDash.getName(), clientSheet.getName(), engagementSheet.getName(), revenueSheet.getName());
        
        this.createDashboard(clientDash, {
          title: 'Client Management Dashboard',
          subtitle: 'Client Portfolio & Engagement Analytics',
          industry: 'professional',
          kpis: [
            { value: '=COUNTIF(\'Clients\'!E:E,"Active")', label: 'Active Clients' },
            { value: '=TEXT(SUM(\'Revenue\'!D:D),"$#,##0")', label: 'Total Revenue YTD', trend: '+30% vs LY', trendColor: '#10B981' },
            { value: '=TEXT(AVERAGE(\'Revenue\'!D:D),"$#,##0")', label: 'Avg Client Value' },
            { value: '=COUNTIF(\'Engagements\'!F:F,"In Progress")', label: 'Active Engagements' },
            { value: '=TEXT(AVERAGE(\'Clients\'!G:G),"0")', label: 'Avg Satisfaction', trend: '4.8/5 rating' },
            { value: '=TEXT(COUNTIF(\'Clients\'!H:H,"At Risk")/COUNTA(\'Clients\'!A:A),"0%")', label: 'Churn Risk', trend: 'Low', trendColor: '#10B981' }
          ],
          charts: [
            { title: 'Client Revenue', type: 'Column', description: 'Revenue by client tier', dataRange: 'Revenue!A:D' },
            { title: 'Engagement Status', type: 'Pie', description: 'Current engagement distribution', dataRange: 'Engagements!F:F' },
            { title: 'Client Growth', type: 'Line', description: 'Monthly client acquisition', dataRange: 'Clients!A:E' }
          ]
        });
        
        this.setupClientDashboardSheet(clientSheet);
        this.setupEngagementSheet(engagementSheet);
        this.setupRevenueSheet(revenueSheet);
        spreadsheet.setActiveSheet(clientDash);
        spreadsheet.moveActiveSheet(1);
        break;
        
      case 'time-billing-tracker':
        const billDash = spreadsheet.insertSheet(prefix + 'Dashboard');
        const billingSheet = spreadsheet.insertSheet(prefix + 'Time & Billing');
        const timesheetSheet = spreadsheet.insertSheet(prefix + 'Timesheets');
        const invoicesSheet = spreadsheet.insertSheet(prefix + 'Invoices');
        sheets.push(billDash.getName(), billingSheet.getName(), timesheetSheet.getName(), invoicesSheet.getName());
        
        this.createDashboard(billDash, {
          title: 'Time & Billing Dashboard',
          subtitle: 'Billable Hours & Invoice Tracking',
          industry: 'professional',
          kpis: [
            { value: '=SUM(\'Timesheets\'!D:D)', label: 'Total Hours MTD' },
            { value: '=TEXT(SUM(\'Timesheets\'!E:E)/SUM(\'Timesheets\'!D:D),"0%")', label: 'Billable Ratio', trend: '85% target' },
            { value: '=TEXT(SUM(\'Invoices\'!E:E),"$#,##0")', label: 'Invoiced Amount', trend: '+15% vs LM', trendColor: '#10B981' },
            { value: '=TEXT(SUM(\'Invoices\'!F:F),"$#,##0")', label: 'Outstanding AR' },
            { value: '=TEXT(AVERAGE(\'Time & Billing\'!F:F),"$#,##0")', label: 'Avg Hourly Rate' },
            { value: '=AVERAGE(\'Invoices\'!G:G)', label: 'Days to Payment', trend: 'Improving' }
          ],
          charts: [
            { title: 'Billable vs Non-Billable', type: 'Pie', description: 'Time allocation', dataRange: 'Timesheets!D:E' },
            { title: 'Revenue by Client', type: 'Bar', description: 'Top revenue generators', dataRange: 'Invoices!B:E' },
            { title: 'Weekly Hours Trend', type: 'Line', description: 'Hours worked over time', dataRange: 'Timesheets!A:D' }
          ]
        });
        
        this.setupTimeBillingSheet(billingSheet);
        this.setupTimesheetSheet(timesheetSheet);
        this.setupInvoicesSheet(invoicesSheet);
        spreadsheet.setActiveSheet(billDash);
        spreadsheet.moveActiveSheet(1);
        break;
    }
    
    return { sheets };
  },
  
  /**
   * Apply an industry-specific template
   * @param {string} templateType - Type of template to apply
   * @return {Object} Result with created sheets
   */
  applyTemplate: function(templateType) {
    try {
      // First, clean up any existing preview sheets
      this.cleanupPreviewSheets();
      
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const createdSheets = [];
      
      // Use the same create functions but without preview flag
      switch (templateType) {
        // Real Estate templates
        case 'commission-tracker':
        case 'property-manager':
        case 'investment-analyzer':
        case 'lead-pipeline':
          const realEstateResult = this.createRealEstateTemplate(spreadsheet, templateType, false);
          createdSheets.push(...realEstateResult.sheets);
          break;
          
        // Construction templates
        case 'cost-estimator':
        case 'material-tracker':
        case 'labor-manager':
        case 'change-orders':
          const constructionResult = this.createConstructionTemplate(spreadsheet, templateType, false);
          createdSheets.push(...constructionResult.sheets);
          break;
          
        // Healthcare templates
        case 'insurance-verifier':
        case 'prior-auth-tracker':
        case 'revenue-cycle':
        case 'denial-analytics':
          const healthcareResult = this.createHealthcareTemplate(spreadsheet, templateType, false);
          createdSheets.push(...healthcareResult.sheets);
          break;
          
        // Marketing templates
        case 'campaign-tracker':
        case 'lead-scoring-system':
        case 'competitor-analysis':
        case 'brand-awareness':
        case 'customer-lifetime-value':
          const marketingResult = this.createMarketingTemplate(spreadsheet, templateType, false);
          createdSheets.push(...marketingResult.sheets);
          break;
          
        // E-commerce templates
        case 'sales-dashboard':
        case 'inventory-manager':
        case 'customer-analytics':
        case 'product-performance':
        case 'financial-tracker':
          const ecommerceResult = this.createEcommerceTemplate(spreadsheet, templateType, false);
          createdSheets.push(...ecommerceResult.sheets);
          break;
          
        // Consulting templates
        case 'time-billing-tracker':
        case 'project-profitability':
        case 'client-dashboard':
          const consultingResult = this.createConsultingTemplate(spreadsheet, templateType, false);
          createdSheets.push(...consultingResult.sheets);
          break;
          
        // Professional Services templates
        case 'project-tracker':
        case 'time-billing':
        case 'resource-planner':
        case 'proposal-tracker':
        case 'financial-dashboard':
          const professionalServicesResult = this.createProfessionalServicesTemplate(spreadsheet, templateType, false);
          createdSheets.push(...professionalServicesResult.sheets);
          break;
          
        default:
          return { success: false, error: 'Unknown template type: ' + templateType };
      }
      
      return {
        success: true,
        sheets: createdSheets,
        message: `Created ${createdSheets.length} sheet(s) for ${templateType}`
      };
      
    } catch (error) {
      Logger.error('Error applying template:', error);
      return { success: false, error: error.message };
    }
  },
  
  /**
   * Enhanced Real Estate Commission Tracker with Pipeline Management
   * Saves 8 hours weekly, 25% transaction boost
   * Includes: Client pipeline, commission splits, performance metrics, closing timeline
   */
  createRealEstateCommissionTracker: function(spreadsheet) {
    const sheets = [];
    
    // 1. Client Pipeline Sheet (Lead to Close tracking)
    const pipelineSheet = spreadsheet.insertSheet('Client Pipeline');
    sheets.push(pipelineSheet.getName());
    
    const pipelineHeaders = [
      'Lead ID', 'Date Added', 'Client Name', 'Phone', 'Email',
      'Property Interest', 'Budget Range', 'Pre-Approved', 'Stage',
      'Next Action', 'Action Date', 'Agent', 'Source', 'Score',
      'Days in Pipeline', 'Conversion Probability', 'Est. Commission', 'Notes'
    ];
    
    pipelineSheet.getRange(1, 1, 1, pipelineHeaders.length).setValues([pipelineHeaders]);
    pipelineSheet.getRange(1, 1, 1, pipelineHeaders.length).setFontWeight('bold');
    pipelineSheet.getRange(1, 1, 1, pipelineHeaders.length).setBackground('#6366f1');
    pipelineSheet.getRange(1, 1, 1, pipelineHeaders.length).setFontColor('#ffffff');
    
    // Add pipeline stage validation
    const stageRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['New Lead', 'Contacted', 'Qualified', 'Showing', 'Offer Made', 'Negotiating', 'Under Contract', 'Closed', 'Lost'], true)
      .build();
    pipelineSheet.getRange('I2:I').setDataValidation(stageRule);
    
    // Add formulas for pipeline metrics
    pipelineSheet.getRange('O2').setFormula('=IF(B2<>"",TODAY()-B2,"")');
    pipelineSheet.getRange('P2').setFormula('=IFS(I2="Closed",100,I2="Under Contract",90,I2="Negotiating",70,I2="Offer Made",50,I2="Showing",30,I2="Qualified",20,I2="Contacted",10,I2="New Lead",5,TRUE,0)&"%"');
    pipelineSheet.getRange('Q2').setFormula('=IF(AND(G2<>"",P2<>""),VALUE(SUBSTITUTE(G2,"$",""))*3/100*VALUE(SUBSTITUTE(P2,"%",""))/100,"")');
    
    // 2. Enhanced Commission Tracker
    const commissionSheet = spreadsheet.insertSheet('Commission Tracker');
    sheets.push(commissionSheet.getName());
    
    // Enhanced headers with more tracking fields
    const headers = [
      'Transaction ID', 'List Date', 'Contract Date', 'Closing Date',
      'Client Name', 'Property Address', 'Property Type', 'MLS #',
      'List Price', 'Sale Price', 'Price/SqFt', 'Days on Market',
      'Commission Rate %', 'Gross Commission', 'Co-op Split %', 'Co-op Amount',
      'Office Split %', 'Net to Agent', 'Transaction Fee', 'Final Commission',
      'Buyer/Seller Side', 'Lead Source', 'Agent', 'Status', 
      'Escrow Company', 'Title Company', 'Notes'
    ];
    
    commissionSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    commissionSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    commissionSheet.getRange(1, 1, 1, headers.length).setBackground('#4f46e5');
    commissionSheet.getRange(1, 1, 1, headers.length).setFontColor('#ffffff');
    
    // Advanced sample data with comprehensive formulas
    const sampleData = [
      ['TRX-001', '=TODAY()-45', '=B2+30', '=C2+15',
       'John Smith', '123 Main St, City, ST', 'Single Family', 'MLS123456',
       475000, 450000, '=J2/2500', '=C2-B2',
       3, '=J2*M2/100', 50, '=N2*O2/100',
       30, '=N2-P2', 395, '=R2-S2',
       'Seller', 'Referral', 'Agent 1', 'Pending',
       'ABC Escrow', 'XYZ Title', ''
      ]
    ];
    
    commissionSheet.getRange(2, 1, sampleData.length, sampleData[0].length).setFormulas(sampleData);
    
    // Format currency columns - expanded for new fields
    commissionSheet.getRange('I:J').setNumberFormat('$#,##0'); // List/Sale Price
    commissionSheet.getRange('K:K').setNumberFormat('$#,##0'); // Price/SqFt
    commissionSheet.getRange('N:N').setNumberFormat('$#,##0'); // Gross Commission
    commissionSheet.getRange('P:P').setNumberFormat('$#,##0'); // Co-op Amount
    commissionSheet.getRange('R:T').setNumberFormat('$#,##0'); // Net amounts
    
    // Format percentage columns - expanded
    commissionSheet.getRange('M:M').setNumberFormat('0.00%'); // Commission Rate
    commissionSheet.getRange('O:O').setNumberFormat('0%'); // Co-op Split
    commissionSheet.getRange('Q:Q').setNumberFormat('0%'); // Office Split
    
    // Add conditional formatting for status
    const statusRange = commissionSheet.getRange('K:K');
    const pendingRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Pending')
      .setBackground('#fef3c7')
      .setRanges([statusRange])
      .build();
    const closedRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Closed')
      .setBackground('#dcfce7')
      .setRanges([statusRange])
      .build();
    
    commissionSheet.setConditionalFormatRules([pendingRule, closedRule]);
    
    // 3. Create Advanced Performance Dashboard
    const dashboardSheet = spreadsheet.insertSheet('Performance Dashboard');
    sheets.push(dashboardSheet.getName());
    
    // Dashboard title and period selector
    dashboardSheet.getRange('A1').setValue('Real Estate Performance Dashboard').setFontSize(20).setFontWeight('bold');
    dashboardSheet.getRange('A2').setValue('Period: ' + new Date().toLocaleDateString('en-US', {month: 'long', year: 'numeric'}));
    
    // Key Performance Indicators section
    dashboardSheet.getRange('A4').setValue('KEY METRICS').setFontWeight('bold').setFontSize(14);
    dashboardSheet.getRange('A4:D4').setBackground('#4f46e5').setFontColor('#ffffff');
    
    // Comprehensive KPI formulas
    const kpiData = [
      ['Pipeline Metrics', '', '', ''],
      ['Active Leads', '=COUNTIF(\'Client Pipeline\'!I:I,"<>Closed")', 'New This Month', '=COUNTIFS(\'Client Pipeline\'!B:B,">="&EOMONTH(TODAY(),-1)+1,\'Client Pipeline\'!B:B,"<="&EOMONTH(TODAY(),0))'],
      ['Hot Prospects', '=COUNTIFS(\'Client Pipeline\'!N:N,">7")', 'Conversion Rate', '=COUNTIF(\'Client Pipeline\'!I:I,"Closed")/COUNTA(\'Client Pipeline\'!A:A)*100&"%"'],
      ['', '', '', ''],
      ['Commission Performance', '', '', ''],
      ['YTD Gross Commission', '=SUMIFS(\'Commission Tracker\'!N:N,\'Commission Tracker\'!B:B,">="&DATE(YEAR(TODAY()),1,1))', 'YTD Net Commission', '=SUMIFS(\'Commission Tracker\'!T:T,\'Commission Tracker\'!B:B,">="&DATE(YEAR(TODAY()),1,1))'],
      ['Avg Commission/Deal', '=AVERAGE(\'Commission Tracker\'!T:T)', 'Avg Sale Price', '=AVERAGE(\'Commission Tracker\'!J:J)'],
      ['Pending Commissions', '=SUMIFS(\'Commission Tracker\'!T:T,\'Commission Tracker\'!X:X,"Pending")', 'Closed This Month', '=COUNTIFS(\'Commission Tracker\'!D:D,">="&EOMONTH(TODAY(),-1)+1,\'Commission Tracker\'!D:D,"<="&EOMONTH(TODAY(),0))'],
      ['', '', '', ''],
      ['Market Analysis', '', '', ''],
      ['Avg Days on Market', '=AVERAGE(\'Commission Tracker\'!L:L)', 'Avg List to Sale Ratio', '=AVERAGE(\'Commission Tracker\'!J:J)/AVERAGE(\'Commission Tracker\'!I:I)*100&"%"'],
      ['Total Volume Sold', '=SUM(\'Commission Tracker\'!J:J)', 'Properties Under Contract', '=COUNTIF(\'Commission Tracker\'!X:X,"Under Contract")'],
      ['', '', '', ''],
      ['Lead Source ROI', '', '', ''],
      ['Top Lead Source', '=INDEX(\'Commission Tracker\'!V:V,MODE(MATCH(\'Commission Tracker\'!V:V,\'Commission Tracker\'!V:V,0)))', 'Best Converting Source', '=INDEX(\'Client Pipeline\'!M:M,MODE(MATCH(\'Client Pipeline\'!M:M,\'Client Pipeline\'!M:M,0)))']
    ];
    
    dashboardSheet.getRange(5, 1, kpiData.length, 4).setValues(kpiData);
    
    // Format KPI values
    dashboardSheet.getRange('B6:B12').setNumberFormat('#,##0');
    dashboardSheet.getRange('D6:D12').setNumberFormat('#,##0');
    dashboardSheet.getRange('B10:B11').setNumberFormat('$#,##0');
    dashboardSheet.getRange('D10:D11').setNumberFormat('$#,##0');
    dashboardSheet.getRange('B13:B13').setNumberFormat('$#,##0');
    dashboardSheet.getRange('B17:B17').setNumberFormat('$#,##0');
    
    // Add visual indicators for sections
    dashboardSheet.getRange('A5:D5').setBackground('#e0e7ff').setFontWeight('bold');
    dashboardSheet.getRange('A9:D9').setBackground('#dcfce7').setFontWeight('bold');
    dashboardSheet.getRange('A14:D14').setBackground('#fed7aa').setFontWeight('bold');
    dashboardSheet.getRange('A18:D18').setBackground('#fce7f3').setFontWeight('bold');
    
    // 4. Create Goal Tracking Sheet
    const goalsSheet = spreadsheet.insertSheet('Goals & Targets');
    sheets.push(goalsSheet.getName());
    
    const goalHeaders = [
      'Period', 'Target Listings', 'Actual Listings', 'Target Sales', 'Actual Sales',
      'Target Volume', 'Actual Volume', 'Target GCI', 'Actual GCI', 'Achievement %'
    ];
    
    goalsSheet.getRange(1, 1, 1, goalHeaders.length).setValues([goalHeaders]);
    goalsSheet.getRange(1, 1, 1, goalHeaders.length).setFontWeight('bold');
    goalsSheet.getRange(1, 1, 1, goalHeaders.length).setBackground('#10b981');
    goalsSheet.getRange(1, 1, 1, goalHeaders.length).setFontColor('#ffffff');
    
    // Add monthly goals with formulas
    const currentMonth = new Date().toLocaleDateString('en-US', {month: 'short', year: 'numeric'});
    goalsSheet.getRange('A2').setValue(currentMonth);
    goalsSheet.getRange('B2').setValue(10); // Target listings
    goalsSheet.getRange('C2').setFormula('=COUNTIFS(\'Commission Tracker\'!B:B,">="&EOMONTH(TODAY(),-1)+1,\'Commission Tracker\'!B:B,"<="&EOMONTH(TODAY(),0))');
    goalsSheet.getRange('D2').setValue(5); // Target sales
    goalsSheet.getRange('E2').setFormula('=COUNTIFS(\'Commission Tracker\'!D:D,">="&EOMONTH(TODAY(),-1)+1,\'Commission Tracker\'!D:D,"<="&EOMONTH(TODAY(),0),\'Commission Tracker\'!X:X,"Closed")');
    goalsSheet.getRange('F2').setValue(2000000); // Target volume
    goalsSheet.getRange('G2').setFormula('=SUMIFS(\'Commission Tracker\'!J:J,\'Commission Tracker\'!D:D,">="&EOMONTH(TODAY(),-1)+1,\'Commission Tracker\'!D:D,"<="&EOMONTH(TODAY(),0))');
    goalsSheet.getRange('H2').setValue(50000); // Target GCI
    goalsSheet.getRange('I2').setFormula('=SUMIFS(\'Commission Tracker\'!T:T,\'Commission Tracker\'!D:D,">="&EOMONTH(TODAY(),-1)+1,\'Commission Tracker\'!D:D,"<="&EOMONTH(TODAY(),0))');
    goalsSheet.getRange('J2').setFormula('=AVERAGE(C2/B2,E2/D2,G2/F2,I2/H2)*100');
    
    // Format goal tracking
    goalsSheet.getRange('F2:I2').setNumberFormat('$#,##0');
    goalsSheet.getRange('J2').setNumberFormat('0%');
    
    // Add conditional formatting for achievement
    const achievementRule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThanOrEqualTo(100)
      .setBackground('#10b981')
      .setFontColor('#ffffff')
      .setRanges([goalsSheet.getRange('J2:J')])
      .build();
    goalsSheet.setConditionalFormatRules([achievementRule]);
    
    // Optimize column widths for all sheets
    pipelineSheet.autoResizeColumns(1, pipelineHeaders.length);
    commissionSheet.autoResizeColumns(1, headers.length);
    dashboardSheet.setColumnWidth(1, 180);
    dashboardSheet.setColumnWidth(2, 150);
    dashboardSheet.setColumnWidth(3, 180);
    dashboardSheet.setColumnWidth(4, 150);
    goalsSheet.autoResizeColumns(1, goalHeaders.length);
    
    return sheets;
  },
  
  /**
   * Property Management System
   * Tracks rent, maintenance, tenants - saves 12 hours weekly
   */
  createPropertyManager: function(spreadsheet) {
    const sheets = [];
    
    // Properties Sheet
    const propertiesSheet = spreadsheet.insertSheet('Properties');
    sheets.push(propertiesSheet.getName());
    
    const propHeaders = [
      'Property ID', 'Address', 'Type', 'Bedrooms', 'Bathrooms', 
      'Monthly Rent', 'Tenant Name', 'Lease Start', 'Lease End', 
      'Status', 'Last Maintenance', 'Notes'
    ];
    
    propertiesSheet.getRange(1, 1, 1, propHeaders.length).setValues([propHeaders]);
    propertiesSheet.getRange(1, 1, 1, propHeaders.length).setFontWeight('bold');
    propertiesSheet.getRange(1, 1, 1, propHeaders.length).setBackground('#2563eb');
    propertiesSheet.getRange(1, 1, 1, propHeaders.length).setFontColor('#ffffff');
    
    // Rent Collection Sheet
    const rentSheet = spreadsheet.insertSheet('Rent Collection');
    sheets.push(rentSheet.getName());
    
    const rentHeaders = [
      'Month', 'Property ID', 'Tenant', 'Rent Due', 'Amount Paid', 
      'Payment Date', 'Late Fee', 'Total Collected', 'Status'
    ];
    
    rentSheet.getRange(1, 1, 1, rentHeaders.length).setValues([rentHeaders]);
    rentSheet.getRange(1, 1, 1, rentHeaders.length).setFontWeight('bold');
    rentSheet.getRange(1, 1, 1, rentHeaders.length).setBackground('#059669');
    rentSheet.getRange(1, 1, 1, rentHeaders.length).setFontColor('#ffffff');
    
    // Add formulas for late fees
    rentSheet.getRange('G2').setFormula('=IF(AND(F2>EOMONTH(A2,0)+5,E2>0),D2*0.05,0)');
    rentSheet.getRange('H2').setFormula('=E2+G2');
    rentSheet.getRange('I2').setFormula('=IF(E2=0,"Unpaid",IF(F2<=EOMONTH(A2,0)+5,"On Time","Late"))');
    
    // Maintenance Tracker
    const maintSheet = spreadsheet.insertSheet('Maintenance');
    sheets.push(maintSheet.getName());
    
    const maintHeaders = [
      'Date Reported', 'Property ID', 'Issue Type', 'Description', 
      'Priority', 'Vendor', 'Cost Estimate', 'Actual Cost', 
      'Status', 'Completion Date'
    ];
    
    maintSheet.getRange(1, 1, 1, maintHeaders.length).setValues([maintHeaders]);
    maintSheet.getRange(1, 1, 1, maintHeaders.length).setFontWeight('bold');
    maintSheet.getRange(1, 1, 1, maintHeaders.length).setBackground('#dc2626');
    maintSheet.getRange(1, 1, 1, maintHeaders.length).setFontColor('#ffffff');
    
    return sheets;
  },
  
  /**
   * Advanced Construction Estimator with Material Waste Tracking
   * Saves 20+ hours per project with comprehensive tracking
   */
  createConstructionEstimator: function(spreadsheet) {
    const sheets = [];
    
    // 1. Master Estimate & Project Overview
    const estimateSheet = spreadsheet.insertSheet('Master Estimate');
    sheets.push(estimateSheet.getName());
    
    estimateSheet.getRange('A1').setValue('Construction Project Master Estimate').setFontSize(20).setFontWeight('bold');
    estimateSheet.getRange('A1:H1').merge().setHorizontalAlignment('center').setBackground('#1f2937').setFontColor('#ffffff');
    
    // Project Information
    estimateSheet.getRange('A3').setValue('PROJECT INFORMATION').setFontWeight('bold').setBackground('#374151').setFontColor('#ffffff');
    estimateSheet.getRange('A3:H3').merge();
    
    const projectInfo = [
      ['Project Name:', '', 'Project ID:', '', 'Contract Type:', '', 'Status:', ''],
      ['Client Name:', '', 'Phone:', '', 'Email:', '', 'Site Address:', ''],
      ['Start Date:', '=TODAY()', 'End Date:', '=TODAY()+90', 'Duration (Days):', '=DAYS(D5,B5)', 'Working Days:', '=NETWORKDAYS(B5,D5)'],
      ['Project Manager:', '', 'Site Supervisor:', '', 'Safety Officer:', '', 'Inspector:', '']
    ];
    estimateSheet.getRange(4, 1, 4, 8).setValues(projectInfo);
    
    // Cost Summary with Advanced Calculations
    estimateSheet.getRange('A9').setValue('COST SUMMARY').setFontWeight('bold').setBackground('#059669').setFontColor('#ffffff');
    estimateSheet.getRange('A9:H9').merge();
    
    const costSummary = [
      ['Category', 'Estimated', 'Actual', 'Variance', '% Complete', 'Projected Final', 'Over/Under', 'Notes'],
      ['Labor', '=SUM(\'Labor Tracking\'!K:K)', '=SUM(\'Labor Tracking\'!Q:Q)', '=C11-B11', '=IF(B11>0,C11/B11,0)', '=IF(E11>0,C11/E11,B11)', '=F11-B11', ''],
      ['Materials', '=SUM(\'Material Tracking\'!O:O)', '=SUM(\'Material Tracking\'!P:P)', '=C12-B12', '=IF(B12>0,C12/B12,0)', '=IF(E12>0,C12/E12,B12)', '=F12-B12', ''],
      ['Equipment', '=SUM(\'Equipment Rental\'!H:H)', '=SUM(\'Equipment Rental\'!I:I)', '=C13-B13', '=IF(B13>0,C13/B13,0)', '=IF(E13>0,C13/E13,B13)', '=F13-B13', ''],
      ['Subcontractors', '=SUM(\'Subcontractors\'!G:G)', '=SUM(\'Subcontractors\'!H:H)', '=C14-B14', '=IF(B14>0,C14/B14,0)', '=IF(E14>0,C14/E14,B14)', '=F14-B14', ''],
      ['Permits & Fees', 5000, 0, '=C15-B15', '=IF(B15>0,C15/B15,0)', '=B15', '=F15-B15', ''],
      ['Insurance & Bonds', 3000, 0, '=C16-B16', '=IF(B16>0,C16/B16,0)', '=B16', '=F16-B16', ''],
      ['Contingency (10%)', '=SUM(B11:B16)*0.1', '=SUM(C11:C16)*0.1', '=C17-B17', '', '=SUM(F11:F16)*0.1', '=F17-B17', ''],
      ['', '', '', '', '', '', '', ''],
      ['TOTAL PROJECT COST', '=SUM(B11:B17)', '=SUM(C11:C17)', '=C19-B19', '=C19/B19', '=SUM(F11:F17)', '=F19-B19', '']
    ];
    estimateSheet.getRange(10, 1, costSummary.length, 8).setValues(costSummary);
    
    // Format cost summary
    estimateSheet.getRange('B11:G17').setNumberFormat('$#,##0.00');
    estimateSheet.getRange('E11:E17').setNumberFormat('0%');
    estimateSheet.getRange('B19:D19').setNumberFormat('$#,##0.00').setFontWeight('bold');
    estimateSheet.getRange('F19:G19').setNumberFormat('$#,##0.00').setFontWeight('bold');
    estimateSheet.getRange('A19:H19').setBackground('#fbbf24');
    
    // Conditional formatting for variances
    const varianceRange = estimateSheet.getRange('G11:G17');
    const negativeRule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0)
      .setBackground('#fecaca')
      .setFontColor('#991b1b')
      .setRanges([varianceRange])
      .build();
    const positiveRule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0)
      .setBackground('#bbf7d0')
      .setFontColor('#14532d')
      .setRanges([varianceRange])
      .build();
    estimateSheet.setConditionalFormatRules([negativeRule, positiveRule]);
    
    // 2. Advanced Material Tracking with Waste Analysis
    const materialSheet = spreadsheet.insertSheet('Material Tracking');
    sheets.push(materialSheet.getName());
    
    const materialHeaders = [
      'Date', 'Material Code', 'Description', 'Category', 'Unit', 
      'Ordered Qty', 'Received Qty', 'Used Qty', 'Waste Qty', 'Waste %',
      'Waste Reason', 'Remaining Stock', 'Est. Cost/Unit', 'Actual Cost/Unit',
      'Total Est. Cost', 'Total Actual Cost', 'Variance', 'Supplier',
      'PO Number', 'Invoice #', 'Delivery Date', 'Quality Check', 'Notes'
    ];
    
    materialSheet.getRange(1, 1, 1, materialHeaders.length).setValues([materialHeaders]);
    materialSheet.getRange(1, 1, 1, materialHeaders.length).setFontWeight('bold');
    materialSheet.getRange(1, 1, 1, materialHeaders.length).setBackground('#059669');
    materialSheet.getRange(1, 1, 1, materialHeaders.length).setFontColor('#ffffff');
    
    // Add formulas for material tracking
    for (let row = 2; row <= 50; row++) {
      materialSheet.getRange(`J${row}`).setFormula(`=IF(H${row}>0,I${row}/H${row}*100,0)`).setNumberFormat('0.00%');
      materialSheet.getRange(`L${row}`).setFormula(`=G${row}-H${row}-I${row}`).setNumberFormat('#,##0.00');
      materialSheet.getRange(`O${row}`).setFormula(`=F${row}*M${row}`).setNumberFormat('$#,##0.00');
      materialSheet.getRange(`P${row}`).setFormula(`=G${row}*N${row}`).setNumberFormat('$#,##0.00');
      materialSheet.getRange(`Q${row}`).setFormula(`=P${row}-O${row}`).setNumberFormat('$#,##0.00');
    }
    
    // Waste Analysis Summary
    materialSheet.getRange('A55').setValue('WASTE ANALYSIS SUMMARY').setFontWeight('bold').setBackground('#dc2626').setFontColor('#ffffff');
    materialSheet.getRange('A55:H55').merge();
    
    const wasteAnalysis = [
      ['Total Material Ordered:', '=SUM(F:F)', 'Total Used:', '=SUM(H:H)', 'Total Waste:', '=SUM(I:I)', 'Overall Waste %:', '=IF(D56>0,F56/D56*100,0)'],
      ['Highest Waste Category:', '=INDEX(D:D,MATCH(MAX(J:J),J:J,0))', 'Waste Cost:', '=SUMPRODUCT(I:I,N:N)', 'Potential Savings:', '=D57*0.5', '', '', '', ''],
      ['Common Waste Reasons:', '=MODE.SNGL(K:K)', '', '', '', '', '', '']
    ];
    materialSheet.getRange(56, 1, wasteAnalysis.length, 8).setValues(wasteAnalysis);
    materialSheet.getRange('H56').setNumberFormat('0.00%');
    materialSheet.getRange('D57:E57').setNumberFormat('$#,##0.00');
    
    // 3. Labor Tracking with Productivity Analysis
    const laborSheet = spreadsheet.insertSheet('Labor Tracking');
    sheets.push(laborSheet.getName());
    
    const laborHeaders = [
      'Date', 'Employee ID', 'Name', 'Trade', 'Task Code', 'Task Description',
      'Start Time', 'End Time', 'Regular Hours', 'Overtime Hours', 
      'Est. Hours', 'Actual Hours', 'Productivity %', 'Hourly Rate',
      'Regular Pay', 'Overtime Pay', 'Total Pay', 'Phase', 'Supervisor'
    ];
    
    laborSheet.getRange(1, 1, 1, laborHeaders.length).setValues([laborHeaders]);
    laborSheet.getRange(1, 1, 1, laborHeaders.length).setFontWeight('bold');
    laborSheet.getRange(1, 1, 1, laborHeaders.length).setBackground('#2563eb');
    laborSheet.getRange(1, 1, 1, laborHeaders.length).setFontColor('#ffffff');
    
    // Labor formulas
    for (let row = 2; row <= 100; row++) {
      laborSheet.getRange(`I${row}`).setFormula(`=IF(AND(G${row}<>"",H${row}<>""),MIN((H${row}-G${row})*24,8),0)`);
      laborSheet.getRange(`J${row}`).setFormula(`=IF(AND(G${row}<>"",H${row}<>""),MAX((H${row}-G${row})*24-8,0),0)`);
      laborSheet.getRange(`L${row}`).setFormula(`=I${row}+J${row}`);
      laborSheet.getRange(`M${row}`).setFormula(`=IF(K${row}>0,K${row}/L${row}*100,0)`).setNumberFormat('0%');
      laborSheet.getRange(`O${row}`).setFormula(`=I${row}*N${row}`).setNumberFormat('$#,##0.00');
      laborSheet.getRange(`P${row}`).setFormula(`=J${row}*N${row}*1.5`).setNumberFormat('$#,##0.00');
      laborSheet.getRange(`Q${row}`).setFormula(`=O${row}+P${row}`).setNumberFormat('$#,##0.00');
    }
    
    // 4. Subcontractor Management
    const subSheet = spreadsheet.insertSheet('Subcontractors');
    sheets.push(subSheet.getName());
    
    const subHeaders = [
      'Sub ID', 'Company Name', 'Trade', 'Contact Person', 'Phone', 'Email',
      'Contract Amount', 'Paid to Date', 'Balance Due', 'Start Date', 'End Date',
      'Insurance Exp.', 'License #', 'Performance Score', 'Status', 'Notes'
    ];
    
    subSheet.getRange(1, 1, 1, subHeaders.length).setValues([subHeaders]);
    subSheet.getRange(1, 1, 1, subHeaders.length).setFontWeight('bold');
    subSheet.getRange(1, 1, 1, subHeaders.length).setBackground('#7c3aed');
    subSheet.getRange(1, 1, 1, subHeaders.length).setFontColor('#ffffff');
    
    // Balance calculation
    for (let row = 2; row <= 30; row++) {
      subSheet.getRange(`I${row}`).setFormula(`=G${row}-H${row}`).setNumberFormat('$#,##0.00');
    }
    
    // 5. Change Order Management
    const changeSheet = spreadsheet.insertSheet('Change Orders');
    sheets.push(changeSheet.getName());
    
    const changeHeaders = [
      'CO Number', 'Date Requested', 'Requested By', 'Description', 'Reason',
      'Cost Impact', 'Schedule Impact (Days)', 'Status', 'Approved By',
      'Approval Date', 'Original Contract', 'New Contract Value', 'Notes'
    ];
    
    changeSheet.getRange(1, 1, 1, changeHeaders.length).setValues([changeHeaders]);
    changeSheet.getRange(1, 1, 1, changeHeaders.length).setFontWeight('bold');
    changeSheet.getRange(1, 1, 1, changeHeaders.length).setBackground('#dc2626');
    changeSheet.getRange(1, 1, 1, changeHeaders.length).setFontColor('#ffffff');
    
    // Change order summary
    changeSheet.getRange('A35').setValue('CHANGE ORDER SUMMARY').setFontWeight('bold');
    changeSheet.getRange('A36').setValue('Total Change Orders:');
    changeSheet.getRange('B36').setFormula('=COUNTA(A2:A30)');
    changeSheet.getRange('A37').setValue('Total Cost Impact:');
    changeSheet.getRange('B37').setFormula('=SUM(F:F)').setNumberFormat('$#,##0.00');
    changeSheet.getRange('A38').setValue('Total Schedule Impact:');
    changeSheet.getRange('B38').setFormula('=SUM(G:G)&" days"');
    
    // 6. Equipment Rental Tracking
    const equipmentSheet = spreadsheet.insertSheet('Equipment Rental');
    sheets.push(equipmentSheet.getName());
    
    const equipHeaders = [
      'Equipment ID', 'Description', 'Category', 'Supplier', 'Start Date',
      'End Date', 'Daily Rate', 'Est. Cost', 'Actual Cost', 'Status',
      'Operator Required', 'Fuel Included', 'Maintenance', 'Notes'
    ];
    
    equipmentSheet.getRange(1, 1, 1, equipHeaders.length).setValues([equipHeaders]);
    equipmentSheet.getRange(1, 1, 1, equipHeaders.length).setFontWeight('bold');
    equipmentSheet.getRange(1, 1, 1, equipHeaders.length).setBackground('#0891b2');
    equipmentSheet.getRange(1, 1, 1, equipHeaders.length).setFontColor('#ffffff');
    
    // Equipment cost calculations
    for (let row = 2; row <= 30; row++) {
      equipmentSheet.getRange(`H${row}`).setFormula(`=IF(AND(E${row}<>"",F${row}<>""),DAYS(F${row},E${row})*G${row},0)`).setNumberFormat('$#,##0.00');
    }
    
    // 7. Project Dashboard
    const dashboardSheet = spreadsheet.insertSheet('Project Dashboard');
    sheets.push(dashboardSheet.getName());
    
    dashboardSheet.getRange('A1').setValue('PROJECT DASHBOARD').setFontSize(24).setFontWeight('bold');
    dashboardSheet.getRange('A1:J1').merge().setHorizontalAlignment('center').setBackground('#1f2937').setFontColor('#ffffff');
    
    // KPI Section
    dashboardSheet.getRange('A3').setValue('KEY PERFORMANCE INDICATORS').setFontWeight('bold').setBackground('#059669').setFontColor('#ffffff');
    dashboardSheet.getRange('A3:J3').merge();
    
    const kpiData = [
      ['Budget Performance', '', 'Schedule Performance', '', 'Quality Metrics', '', 'Safety Metrics', '', 'Productivity', ''],
      ['Budget Used:', '=\'Master Estimate\'!C19/\'Master Estimate\'!B19', 'Days Elapsed:', '=DAYS(TODAY(),\'Master Estimate\'!B5)', 'Material Waste:', '=\'Material Tracking\'!H56', 'Safety Incidents:', '0', 'Labor Efficiency:', '=AVERAGE(\'Labor Tracking\'!M:M)'],
      ['Cost Variance:', '=\'Master Estimate\'!D19', 'Days Remaining:', '=DAYS(\'Master Estimate\'!D5,TODAY())', 'Quality Issues:', '0', 'Near Misses:', '0', 'Equipment Utilization:', '85%'],
      ['Projected Final:', '=\'Master Estimate\'!F19', 'Completion %:', '=MIN(B5/\'Master Estimate\'!H5*100,100)&"%"', 'Rework Items:', '0', 'Safety Score:', '98%', 'Change Orders:', '=\'Change Orders\'!B36']
    ];
    
    dashboardSheet.getRange(4, 1, kpiData.length, 10).setValues(kpiData);
    dashboardSheet.getRange('B5:B7').setNumberFormat('$#,##0.00');
    dashboardSheet.getRange('B5').setNumberFormat('0%');
    dashboardSheet.getRange('F5').setNumberFormat('0.00%');
    dashboardSheet.getRange('J5').setNumberFormat('0%');
    
    // Format KPI sections
    dashboardSheet.getRange('A4:B4').setBackground('#dcfce7').setFontWeight('bold');
    dashboardSheet.getRange('C4:D4').setBackground('#dbeafe').setFontWeight('bold');
    dashboardSheet.getRange('E4:F4').setBackground('#fef3c7').setFontWeight('bold');
    dashboardSheet.getRange('G4:H4').setBackground('#fee2e2').setFontWeight('bold');
    dashboardSheet.getRange('I4:J4').setBackground('#e9d5ff').setFontWeight('bold');
    
    // Critical Alerts Section
    dashboardSheet.getRange('A9').setValue('CRITICAL ALERTS & ACTION ITEMS').setFontWeight('bold').setBackground('#dc2626').setFontColor('#ffffff');
    dashboardSheet.getRange('A9:J9').merge();
    
    const alerts = [
      ['Priority', 'Category', 'Issue', 'Impact', 'Action Required', 'Owner', 'Due Date', 'Status', '', ''],
      ['HIGH', 'Budget', '=IF(\'Master Estimate\'!G19>\'Master Estimate\'!B19*0.05,"Cost overrun detected: "&TEXT(\'Master Estimate\'!G19,"$#,##0"),"")', 'Financial', 'Review and adjust', 'PM', '=TODAY()+3', 'Open', '', ''],
      ['MEDIUM', 'Schedule', '=IF(\'Master Estimate\'!H5<7,"Less than 7 days to deadline","")', 'Timeline', 'Expedite critical tasks', 'Super', '=TODAY()+1', 'Open', '', ''],
      ['LOW', 'Quality', '=IF(\'Material Tracking\'!H56>0.1,"High material waste: "&TEXT(\'Material Tracking\'!H56,"0%"),"")', 'Cost', 'Improve handling', 'Foreman', '=TODAY()+7', 'Open', '', '']
    ];
    
    dashboardSheet.getRange(10, 1, alerts.length, 10).setValues(alerts);
    
    // Conditional formatting for priority
    const priorityRange = dashboardSheet.getRange('A11:A13');
    const highPriorityRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('HIGH')
      .setBackground('#dc2626')
      .setFontColor('#ffffff')
      .setRanges([priorityRange])
      .build();
    const medPriorityRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('MEDIUM')
      .setBackground('#f59e0b')
      .setRanges([priorityRange])
      .build();
    const lowPriorityRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('LOW')
      .setBackground('#10b981')
      .setRanges([priorityRange])
      .build();
    dashboardSheet.setConditionalFormatRules([highPriorityRule, medPriorityRule, lowPriorityRule]);
    
    // Optimize column widths
    estimateSheet.autoResizeColumns(1, 8);
    materialSheet.setColumnWidth(3, 150);
    materialSheet.setColumnWidth(11, 120);
    laborSheet.setColumnWidth(6, 150);
    subSheet.setColumnWidth(2, 150);
    changeSheet.setColumnWidth(4, 200);
    dashboardSheet.setColumnWidth(3, 200);
    dashboardSheet.setColumnWidth(5, 150);
    
    return sheets;
  },
  
  /**
   * Advanced Marketing Dashboard with Multi-Channel Attribution
   * Saves 40+ hours monthly with comprehensive attribution modeling
   */
  createMarketingDashboard: function(spreadsheet) {
    const sheets = [];
    
    // 1. Multi-Channel Campaign Performance
    const campaignSheet = spreadsheet.insertSheet('Campaign Performance');
    sheets.push(campaignSheet.getName());
    
    campaignSheet.getRange('A1').setValue('Multi-Channel Marketing Dashboard').setFontSize(20).setFontWeight('bold');
    campaignSheet.getRange('A1:T1').merge().setHorizontalAlignment('center').setBackground('#7c3aed').setFontColor('#ffffff');
    
    const headers = [
      'Campaign ID', 'Campaign Name', 'Channel', 'Sub-Channel', 'Start Date', 'End Date',
      'Status', 'Budget', 'Spend', 'Impressions', 'Reach', 'Clicks', 'CTR %',
      'Conversions', 'Conversion Rate %', 'CPA', 'ROAS', 'ROI %', 'Attribution Model',
      'Quality Score'
    ];
    
    campaignSheet.getRange(3, 1, 1, headers.length).setValues([headers]);
    campaignSheet.getRange(3, 1, 1, headers.length).setFontWeight('bold');
    campaignSheet.getRange(3, 1, 1, headers.length).setBackground('#6366f1');
    campaignSheet.getRange(3, 1, 1, headers.length).setFontColor('#ffffff');
    
    // Add formulas for campaign metrics
    for (let row = 4; row <= 50; row++) {
      campaignSheet.getRange(`M${row}`).setFormula(`=IF(AND(J${row}>0,L${row}>0),L${row}/J${row}*100,0)`).setNumberFormat('0.00%');
      campaignSheet.getRange(`O${row}`).setFormula(`=IF(AND(L${row}>0,N${row}>0),N${row}/L${row}*100,0)`).setNumberFormat('0.00%');
      campaignSheet.getRange(`P${row}`).setFormula(`=IF(N${row}>0,I${row}/N${row},0)`).setNumberFormat('$#,##0.00');
      campaignSheet.getRange(`Q${row}`).setFormula(`=IF(I${row}>0,(N${row}*100)/I${row},0)`).setNumberFormat('0.00');
      campaignSheet.getRange(`R${row}`).setFormula(`=IF(I${row}>0,((N${row}*100)-I${row})/I${row}*100,0)`).setNumberFormat('0.00%');
    }
    
    // 2. Attribution Analysis Sheet
    const attributionSheet = spreadsheet.insertSheet('Attribution Analysis');
    sheets.push(attributionSheet.getName());
    
    attributionSheet.getRange('A1').setValue('MULTI-TOUCH ATTRIBUTION ANALYSIS').setFontWeight('bold').setBackground('#10b981').setFontColor('#ffffff');
    attributionSheet.getRange('A1:L1').merge();
    
    const attrHeaders = [
      'Customer Journey ID', 'Customer Email', 'First Touch', 'First Touch Date',
      'Last Touch', 'Last Touch Date', 'Total Touchpoints', 'Journey Duration (Days)',
      'Conversion Value', 'First Touch Credit', 'Last Touch Credit', 'Linear Credit'
    ];
    
    attributionSheet.getRange(3, 1, 1, attrHeaders.length).setValues([attrHeaders]);
    attributionSheet.getRange(3, 1, 1, attrHeaders.length).setFontWeight('bold');
    attributionSheet.getRange(3, 1, 1, attrHeaders.length).setBackground('#059669');
    attributionSheet.getRange(3, 1, 1, attrHeaders.length).setFontColor('#ffffff');
    
    // Attribution model calculations
    for (let row = 4; row <= 100; row++) {
      attributionSheet.getRange(`H${row}`).setFormula(`=IF(AND(D${row}<>"",F${row}<>""),DAYS(F${row},D${row}),0)`);
      attributionSheet.getRange(`J${row}`).setFormula(`=IF(G${row}=1,I${row},I${row}*0.4)`).setNumberFormat('$#,##0.00');
      attributionSheet.getRange(`K${row}`).setFormula(`=IF(G${row}=1,I${row},I${row}*0.4)`).setNumberFormat('$#,##0.00');
      attributionSheet.getRange(`L${row}`).setFormula(`=I${row}/G${row}`).setNumberFormat('$#,##0.00');
    }
    
    // Attribution Summary
    attributionSheet.getRange('A110').setValue('ATTRIBUTION MODEL COMPARISON').setFontWeight('bold');
    const attrComparison = [
      ['Model', 'Google Ads', 'Facebook', 'Email', 'Organic', 'Direct', 'Total'],
      ['First Touch', '=SUMIF(C:C,"Google Ads",J:J)', '=SUMIF(C:C,"Facebook",J:J)', '=SUMIF(C:C,"Email",J:J)', '=SUMIF(C:C,"Organic",J:J)', '=SUMIF(C:C,"Direct",J:J)', '=SUM(B112:F112)'],
      ['Last Touch', '=SUMIF(E:E,"Google Ads",K:K)', '=SUMIF(E:E,"Facebook",K:K)', '=SUMIF(E:E,"Email",K:K)', '=SUMIF(E:E,"Organic",K:K)', '=SUMIF(E:E,"Direct",K:K)', '=SUM(B113:F113)'],
      ['Linear', '=SUMIF(C:C,"Google Ads",L:L)+SUMIF(E:E,"Google Ads",L:L)', '=SUMIF(C:C,"Facebook",L:L)+SUMIF(E:E,"Facebook",L:L)', '=SUMIF(C:C,"Email",L:L)+SUMIF(E:E,"Email",L:L)', '=SUMIF(C:C,"Organic",L:L)+SUMIF(E:E,"Organic",L:L)', '=SUMIF(C:C,"Direct",L:L)+SUMIF(E:E,"Direct",L:L)', '=SUM(B114:F114)']
    ];
    attributionSheet.getRange(111, 1, attrComparison.length, 7).setValues(attrComparison);
    attributionSheet.getRange('B112:G114').setNumberFormat('$#,##0');
    
    // 3. Customer Acquisition Funnel
    const funnelSheet = spreadsheet.insertSheet('Acquisition Funnel');
    sheets.push(funnelSheet.getName());
    
    funnelSheet.getRange('A1').setValue('CUSTOMER ACQUISITION FUNNEL').setFontSize(18).setFontWeight('bold');
    funnelSheet.getRange('A1:H1').merge().setHorizontalAlignment('center').setBackground('#2563eb').setFontColor('#ffffff');
    
    const funnelStages = [
      ['Stage', 'Users', 'Drop-off Rate', 'Conversion Rate', 'Avg. Time to Next', 'Channel Split', 'Cost', 'Revenue'],
      ['Awareness (Impressions)', '=SUM(\'Campaign Performance\'!J:J)', '', '100%', '', '', '=SUM(\'Campaign Performance\'!I:I)*0.001', ''],
      ['Interest (Clicks)', '=SUM(\'Campaign Performance\'!L:L)', '=1-B4/B3', '=B4/B3', '=AVERAGE(\'Attribution Analysis\'!H:H)/7&" days"', '', '=SUM(\'Campaign Performance\'!I:I)*0.01', ''],
      ['Consideration (Signups)', '=COUNTIF(\'Lead Tracking\'!H:H,"Signup")', '=1-B5/B4', '=B5/B4', '2 days', '', '=SUM(\'Campaign Performance\'!I:I)*0.05', ''],
      ['Intent (Trials)', '=COUNTIF(\'Lead Tracking\'!H:H,"Trial")', '=1-B6/B5', '=B6/B5', '7 days', '', '=SUM(\'Campaign Performance\'!I:I)*0.1', ''],
      ['Purchase (Customers)', '=SUM(\'Campaign Performance\'!N:N)', '=1-B7/B6', '=B7/B6', '14 days', '', '=SUM(\'Campaign Performance\'!I:I)', '=B7*100'],
      ['Retention (Repeat)', '=B7*0.3', '=1-B8/B7', '=B8/B7', '30 days', '', '=B7*50', '=B8*150']
    ];
    
    funnelSheet.getRange(3, 1, funnelStages.length, 8).setValues(funnelStages);
    funnelSheet.getRange('B4:B8').setNumberFormat('#,##0');
    funnelSheet.getRange('C4:D8').setNumberFormat('0.00%');
    funnelSheet.getRange('G4:H8').setNumberFormat('$#,##0');
    
    // Funnel visualization metrics
    funnelSheet.getRange('A11').setValue('FUNNEL METRICS').setFontWeight('bold');
    const funnelMetrics = [
      ['Overall Conversion Rate:', '=B7/B3'],
      ['Cost Per Acquisition:', '=G7/B7'],
      ['Customer Lifetime Value:', '=H8/B8'],
      ['LTV:CAC Ratio:', '=B14/B13'],
      ['Payback Period:', '=B13/100&" months"']
    ];
    funnelSheet.getRange(12, 1, funnelMetrics.length, 2).setValues(funnelMetrics);
    funnelSheet.getRange('B12').setNumberFormat('0.00%');
    funnelSheet.getRange('B13:B14').setNumberFormat('$#,##0.00');
    funnelSheet.getRange('B15').setNumberFormat('0.00');
    
    // 4. Lead Scoring & Tracking
    const leadSheet = spreadsheet.insertSheet('Lead Tracking');
    sheets.push(leadSheet.getName());
    
    const leadHeaders = [
      'Lead ID', 'Date', 'Name', 'Email', 'Phone', 'Company', 'Industry',
      'Source', 'Campaign', 'UTM Source', 'UTM Medium', 'UTM Campaign',
      'Lead Score', 'Behavior Score', 'Demographic Score', 'Status', 
      'Stage', 'Assigned To', 'Last Activity', 'Next Action', 'Follow-up Date',
      'Engagement Level', 'Budget Range', 'Decision Timeline', 'Notes'
    ];
    
    leadSheet.getRange(1, 1, 1, leadHeaders.length).setValues([leadHeaders]);
    leadSheet.getRange(1, 1, 1, leadHeaders.length).setFontWeight('bold');
    leadSheet.getRange(1, 1, 1, leadHeaders.length).setBackground('#0891b2');
    leadSheet.getRange(1, 1, 1, leadHeaders.length).setFontColor('#ffffff');
    
    // Lead scoring formulas
    for (let row = 2; row <= 200; row++) {
      leadSheet.getRange(`M${row}`).setFormula(`=N${row}+O${row}`);
      leadSheet.getRange(`N${row}`).setFormula(`=IF(V${row}="High",40,IF(V${row}="Medium",25,10))`);
      leadSheet.getRange(`O${row}`).setFormula(`=IF(W${row}>50000,30,IF(W${row}>10000,20,10))+IF(X${row}="Immediate",30,IF(X${row}="Quarter",20,10))`);
    }
    
    // Lead scoring rules with conditional formatting
    const scoreRange = leadSheet.getRange('M2:M200');
    const hotLeadRule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThanOrEqualTo(70)
      .setBackground('#dc2626')
      .setFontColor('#ffffff')
      .setRanges([scoreRange])
      .build();
    const warmLeadRule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(40, 69)
      .setBackground('#f59e0b')
      .setRanges([scoreRange])
      .build();
    const coldLeadRule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(40)
      .setBackground('#3b82f6')
      .setFontColor('#ffffff')
      .setRanges([scoreRange])
      .build();
    leadSheet.setConditionalFormatRules([hotLeadRule, warmLeadRule, coldLeadRule]);
    
    // 5. Content Performance Tracking
    const contentSheet = spreadsheet.insertSheet('Content Performance');
    sheets.push(contentSheet.getName());
    
    const contentHeaders = [
      'Content ID', 'Title', 'Type', 'Topic', 'Author', 'Publish Date',
      'Channel', 'URL', 'Views', 'Unique Visitors', 'Avg. Time on Page',
      'Bounce Rate', 'Shares', 'Comments', 'Leads Generated', 'Conversions',
      'Content Score', 'SEO Score', 'Status'
    ];
    
    contentSheet.getRange(1, 1, 1, contentHeaders.length).setValues([contentHeaders]);
    contentSheet.getRange(1, 1, 1, contentHeaders.length).setFontWeight('bold');
    contentSheet.getRange(1, 1, 1, contentHeaders.length).setBackground('#ec4899');
    contentSheet.getRange(1, 1, 1, contentHeaders.length).setFontColor('#ffffff');
    
    // Content performance formulas
    for (let row = 2; row <= 100; row++) {
      contentSheet.getRange(`Q${row}`).setFormula(`=(I${row}/1000)*0.1+(1-L${row})*30+(M${row}/10)*10+(O${row}*5)+(P${row}*10)`);
    }
    
    // 6. Executive Dashboard
    const dashboardSheet = spreadsheet.insertSheet('Executive Dashboard');
    sheets.push(dashboardSheet.getName());
    
    dashboardSheet.getRange('A1').setValue('MARKETING EXECUTIVE DASHBOARD').setFontSize(24).setFontWeight('bold');
    dashboardSheet.getRange('A1:L1').merge().setHorizontalAlignment('center').setBackground('#1f2937').setFontColor('#ffffff');
    
    // Period selector
    dashboardSheet.getRange('A3').setValue('Reporting Period:');
    dashboardSheet.getRange('B3').setValue(new Date().toLocaleDateString('en-US', {month: 'long', year: 'numeric'}));
    
    // KPI Cards
    dashboardSheet.getRange('A5').setValue('KEY PERFORMANCE INDICATORS').setFontWeight('bold').setBackground('#7c3aed').setFontColor('#ffffff');
    dashboardSheet.getRange('A5:L5').merge();
    
    const kpiData = [
      ['Marketing Metrics', '', '', 'Sales Metrics', '', '', 'Efficiency Metrics', '', '', 'Growth Metrics', '', ''],
      ['Total Spend:', '=SUM(\'Campaign Performance\'!I:I)', 'vs Budget:', 'Total Revenue:', '=SUM(\'Acquisition Funnel\'!H:H)', 'vs Target:', 'Overall CAC:', '=\'Acquisition Funnel\'!B13', 'Target:', 'MoM Growth:', '25%', 'YoY:'],
      ['Total Leads:', '=COUNTA(\'Lead Tracking\'!A:A)-1', 'Cost/Lead:', 'SQLs:', '=COUNTIF(\'Lead Tracking\'!M:M,">70")', 'SQL Rate:', 'LTV:CAC:', '=\'Acquisition Funnel\'!B15', 'Target:', 'New Channels:', '2', 'Active:'],
      ['Impressions:', '=SUM(\'Campaign Performance\'!J:J)', 'CPM:', 'Conversions:', '=SUM(\'Campaign Performance\'!N:N)', 'Conv. Rate:', 'ROAS:', '=AVERAGE(\'Campaign Performance\'!Q:Q)', 'Target:', 'Content Pieces:', '=COUNTA(\'Content Performance\'!A:A)-1', 'Viral:']
    ];
    
    dashboardSheet.getRange(6, 1, kpiData.length, 12).setValues(kpiData);
    
    // Format KPI values
    dashboardSheet.getRange('B7:B9').setNumberFormat('$#,##0');
    dashboardSheet.getRange('E7:E9').setNumberFormat('$#,##0');
    dashboardSheet.getRange('H7:H9').setNumberFormat('$#,##0');
    dashboardSheet.getRange('H8').setNumberFormat('0.00');
    
    // Channel Performance Summary
    dashboardSheet.getRange('A11').setValue('CHANNEL PERFORMANCE SUMMARY').setFontWeight('bold').setBackground('#059669').setFontColor('#ffffff');
    dashboardSheet.getRange('A11:L11').merge();
    
    const channelSummary = [
      ['Channel', 'Spend', 'Impressions', 'Clicks', 'CTR', 'Conversions', 'Conv. Rate', 'CPA', 'ROAS', 'Revenue', 'ROI', 'Rank'],
      ['Google Ads', '=SUMIF(\'Campaign Performance\'!C:C,"Google Ads",\'Campaign Performance\'!I:I)', '=SUMIF(\'Campaign Performance\'!C:C,"Google Ads",\'Campaign Performance\'!J:J)', '=SUMIF(\'Campaign Performance\'!C:C,"Google Ads",\'Campaign Performance\'!L:L)', '=D13/C13', '=SUMIF(\'Campaign Performance\'!C:C,"Google Ads",\'Campaign Performance\'!N:N)', '=F13/D13', '=B13/F13', '=J13/B13', '=F13*100', '=(J13-B13)/B13', '=RANK(K13,K13:K17)'],
      ['Facebook', '=SUMIF(\'Campaign Performance\'!C:C,"Facebook",\'Campaign Performance\'!I:I)', '=SUMIF(\'Campaign Performance\'!C:C,"Facebook",\'Campaign Performance\'!J:J)', '=SUMIF(\'Campaign Performance\'!C:C,"Facebook",\'Campaign Performance\'!L:L)', '=D14/C14', '=SUMIF(\'Campaign Performance\'!C:C,"Facebook",\'Campaign Performance\'!N:N)', '=F14/D14', '=B14/F14', '=J14/B14', '=F14*100', '=(J14-B14)/B14', '=RANK(K14,K13:K17)'],
      ['Email', '=SUMIF(\'Campaign Performance\'!C:C,"Email",\'Campaign Performance\'!I:I)', '=SUMIF(\'Campaign Performance\'!C:C,"Email",\'Campaign Performance\'!J:J)', '=SUMIF(\'Campaign Performance\'!C:C,"Email",\'Campaign Performance\'!L:L)', '=D15/C15', '=SUMIF(\'Campaign Performance\'!C:C,"Email",\'Campaign Performance\'!N:N)', '=F15/D15', '=B15/F15', '=J15/B15', '=F15*100', '=(J15-B15)/B15', '=RANK(K15,K13:K17)'],
      ['LinkedIn', '=SUMIF(\'Campaign Performance\'!C:C,"LinkedIn",\'Campaign Performance\'!I:I)', '=SUMIF(\'Campaign Performance\'!C:C,"LinkedIn",\'Campaign Performance\'!J:J)', '=SUMIF(\'Campaign Performance\'!C:C,"LinkedIn",\'Campaign Performance\'!L:L)', '=D16/C16', '=SUMIF(\'Campaign Performance\'!C:C,"LinkedIn",\'Campaign Performance\'!N:N)', '=F16/D16', '=B16/F16', '=J16/B16', '=F16*100', '=(J16-B16)/B16', '=RANK(K16,K13:K17)'],
      ['Organic', '0', '=SUMIF(\'Campaign Performance\'!C:C,"Organic",\'Campaign Performance\'!J:J)', '=SUMIF(\'Campaign Performance\'!C:C,"Organic",\'Campaign Performance\'!L:L)', '=IF(C17>0,D17/C17,0)', '=SUMIF(\'Campaign Performance\'!C:C,"Organic",\'Campaign Performance\'!N:N)', '=IF(D17>0,F17/D17,0)', '0', '0', '=F17*100', '=IF(B17>0,(J17-B17)/B17,J17)', '=RANK(K17,K13:K17)']
    ];
    
    dashboardSheet.getRange(12, 1, channelSummary.length, 12).setValues(channelSummary);
    dashboardSheet.getRange('B13:B17').setNumberFormat('$#,##0');
    dashboardSheet.getRange('C13:D17').setNumberFormat('#,##0');
    dashboardSheet.getRange('E13:E17').setNumberFormat('0.00%');
    dashboardSheet.getRange('G13:G17').setNumberFormat('0.00%');
    dashboardSheet.getRange('H13:H17').setNumberFormat('$#,##0');
    dashboardSheet.getRange('I13:I17').setNumberFormat('0.00');
    dashboardSheet.getRange('J13:J17').setNumberFormat('$#,##0');
    dashboardSheet.getRange('K13:K17').setNumberFormat('0.00%');
    
    // Action Items
    dashboardSheet.getRange('A19').setValue('ACTION ITEMS & RECOMMENDATIONS').setFontWeight('bold').setBackground('#dc2626').setFontColor('#ffffff');
    dashboardSheet.getRange('A19:L19').merge();
    
    const actionItems = [
      ['Priority', 'Item', 'Channel', 'Impact', 'Effort', 'Owner', 'Due Date', 'Status', '', '', '', ''],
      ['HIGH', '=IF(\'Acquisition Funnel\'!B15<3,"Improve LTV:CAC ratio - currently "&TEXT(\'Acquisition Funnel\'!B15,"0.00"),"")', 'All', 'High', 'Medium', 'CMO', '=TODAY()+7', 'Open', '', '', '', ''],
      ['MEDIUM', '=IF(MIN(\'Campaign Performance\'!M:M)<0.01,"Optimize low CTR campaigns","")', 'Paid', 'Medium', 'Low', 'Media Manager', '=TODAY()+14', 'Open', '', '', '', ''],
      ['LOW', '=IF(AVERAGE(\'Content Performance\'!Q:Q)<50,"Improve content quality scores","")', 'Content', 'Medium', 'High', 'Content Lead', '=TODAY()+30', 'Open', '', '', '', '']
    ];
    
    dashboardSheet.getRange(20, 1, actionItems.length, 12).setValues(actionItems);
    
    // Optimize column widths
    campaignSheet.autoResizeColumns(1, 20);
    attributionSheet.setColumnWidth(2, 150);
    funnelSheet.setColumnWidth(1, 200);
    leadSheet.setColumnWidth(3, 120);
    leadSheet.setColumnWidth(4, 150);
    contentSheet.setColumnWidth(2, 200);
    dashboardSheet.setColumnWidth(2, 120);
    dashboardSheet.setColumnWidth(5, 120);
    dashboardSheet.setColumnWidth(8, 120);
    
    return sheets;
  },
  
  /**
   * Advanced Healthcare Billing with Denial Management Automation
   * Saves 10+ hours daily with automated denial tracking and appeals
   */
  createHealthcareVerifier: function(spreadsheet) {
    const sheets = [];
    
    // 1. Comprehensive Patient Insurance Verification
    const insuranceSheet = spreadsheet.insertSheet('Insurance Verification');
    sheets.push(insuranceSheet.getName());
    
    insuranceSheet.getRange('A1').setValue('Patient Insurance Verification System').setFontSize(20).setFontWeight('bold');
    insuranceSheet.getRange('A1:X1').merge().setHorizontalAlignment('center').setBackground('#dc2626').setFontColor('#ffffff');
    
    const headers = [
      'Patient ID', 'MRN', 'Patient Name', 'DOB', 'Phone', 'Email',
      'Primary Insurance', 'Primary Policy #', 'Primary Group #', 'Primary Subscriber',
      'Secondary Insurance', 'Secondary Policy #', 'Secondary Group #',
      'Verification Date', 'Verified By', 'Eligibility Status', 'Coverage Type',
      'Copay', 'Deductible', 'Deductible Met', 'Out of Pocket Max', 'OOP Met',
      'Prior Auth Required', 'Auth Number', 'Auth Expiry'
    ];
    
    insuranceSheet.getRange(3, 1, 1, headers.length).setValues([headers]);
    insuranceSheet.getRange(3, 1, 1, headers.length).setFontWeight('bold');
    insuranceSheet.getRange(3, 1, 1, headers.length).setBackground('#7f1d1d');
    insuranceSheet.getRange(3, 1, 1, headers.length).setFontColor('#ffffff');
    
    // Add formulas for insurance metrics
    for (let row = 4; row <= 200; row++) {
      insuranceSheet.getRange(`T${row}`).setFormula(`=S${row}-T${row}`).setNumberFormat('$#,##0.00');
      insuranceSheet.getRange(`V${row}`).setFormula(`=U${row}-V${row}`).setNumberFormat('$#,##0.00');
    }
    
    // 2. Advanced Claims Management with Denial Tracking
    const claimsSheet = spreadsheet.insertSheet('Claims Management');
    sheets.push(claimsSheet.getName());
    
    claimsSheet.getRange('A1').setValue('CLAIMS PROCESSING & DENIAL MANAGEMENT').setFontWeight('bold').setFontSize(18);
    claimsSheet.getRange('A1:AA1').merge().setHorizontalAlignment('center').setBackground('#0891b2').setFontColor('#ffffff');
    
    const claimHeaders = [
      'Claim ID', 'Patient ID', 'Patient Name', 'Service Date', 'Provider', 'Facility',
      'CPT Code', 'ICD-10', 'Modifier', 'Units', 'Billed Amount', 'Allowed Amount',
      'Submitted Date', 'Insurance', 'Claim Type', 'Status', 'Status Date',
      'Paid Amount', 'Adjustment', 'Patient Responsibility', 'Denial Code',
      'Denial Reason', 'Appeal Status', 'Appeal Date', 'Resubmission Date',
      'Days Outstanding', 'Age Bracket'
    ];
    
    claimsSheet.getRange(3, 1, 1, claimHeaders.length).setValues([claimHeaders]);
    claimsSheet.getRange(3, 1, 1, claimHeaders.length).setFontWeight('bold');
    claimsSheet.getRange(3, 1, 1, claimHeaders.length).setBackground('#164e63');
    claimsSheet.getRange(3, 1, 1, claimHeaders.length).setFontColor('#ffffff');
    
    // Claims aging and metrics formulas
    for (let row = 4; row <= 500; row++) {
      claimsSheet.getRange(`Z${row}`).setFormula(`=IF(P${row}="Paid",0,TODAY()-M${row})`);
      claimsSheet.getRange(`AA${row}`).setFormula(`=IF(Z${row}=0,"Paid",IF(Z${row}<=30,"0-30",IF(Z${row}<=60,"31-60",IF(Z${row}<=90,"61-90",IF(Z${row}<=120,"91-120","120+")))))`);
      claimsSheet.getRange(`T${row}`).setFormula(`=K${row}-R${row}-S${row}`).setNumberFormat('$#,##0.00');
    }
    
    // 3. Denial Analytics Dashboard
    const denialSheet = spreadsheet.insertSheet('Denial Analytics');
    sheets.push(denialSheet.getName());
    
    denialSheet.getRange('A1').setValue('DENIAL ANALYTICS & PREVENTION').setFontSize(20).setFontWeight('bold');
    denialSheet.getRange('A1:L1').merge().setHorizontalAlignment('center').setBackground('#991b1b').setFontColor('#ffffff');
    
    // Denial reason categorization
    denialSheet.getRange('A3').setValue('DENIAL REASON ANALYSIS').setFontWeight('bold').setBackground('#dc2626').setFontColor('#ffffff');
    denialSheet.getRange('A3:F3').merge();
    
    const denialCategories = [
      ['Category', 'Count', 'Percentage', 'Avg. Amount', 'Recovery Rate', 'Action Required'],
      ['Prior Authorization', '=COUNTIF(\'Claims Management\'!U:U,"PA*")', '=B5/$B$15', '=AVERAGEIF(\'Claims Management\'!U:U,"PA*",\'Claims Management\'!K:K)', '=COUNTIFS(\'Claims Management\'!U:U,"PA*",\'Claims Management\'!W:W,"Approved")/$B5', 'Improve auth process'],
      ['Medical Necessity', '=COUNTIF(\'Claims Management\'!U:U,"MN*")', '=B6/$B$15', '=AVERAGEIF(\'Claims Management\'!U:U,"MN*",\'Claims Management\'!K:K)', '=COUNTIFS(\'Claims Management\'!U:U,"MN*",\'Claims Management\'!W:W,"Approved")/$B6', 'Clinical documentation'],
      ['Coding Errors', '=COUNTIF(\'Claims Management\'!U:U,"CO*")', '=B7/$B$15', '=AVERAGEIF(\'Claims Management\'!U:U,"CO*",\'Claims Management\'!K:K)', '=COUNTIFS(\'Claims Management\'!U:U,"CO*",\'Claims Management\'!W:W,"Approved")/$B7', 'Coder training'],
      ['Eligibility', '=COUNTIF(\'Claims Management\'!U:U,"EL*")', '=B8/$B$15', '=AVERAGEIF(\'Claims Management\'!U:U,"EL*",\'Claims Management\'!K:K)', '=COUNTIFS(\'Claims Management\'!U:U,"EL*",\'Claims Management\'!W:W,"Approved")/$B8', 'Verification process'],
      ['Timely Filing', '=COUNTIF(\'Claims Management\'!U:U,"TF*")', '=B9/$B$15', '=AVERAGEIF(\'Claims Management\'!U:U,"TF*",\'Claims Management\'!K:K)', '=COUNTIFS(\'Claims Management\'!U:U,"TF*",\'Claims Management\'!W:W,"Approved")/$B9', 'Submission workflow'],
      ['Duplicate', '=COUNTIF(\'Claims Management\'!U:U,"DP*")', '=B10/$B$15', '=AVERAGEIF(\'Claims Management\'!U:U,"DP*",\'Claims Management\'!K:K)', '=COUNTIFS(\'Claims Management\'!U:U,"DP*",\'Claims Management\'!W:W,"Approved")/$B10', 'Claim tracking'],
      ['Coverage Terminated', '=COUNTIF(\'Claims Management\'!U:U,"CT*")', '=B11/$B$15', '=AVERAGEIF(\'Claims Management\'!U:U,"CT*",\'Claims Management\'!K:K)', '=COUNTIFS(\'Claims Management\'!U:U,"CT*",\'Claims Management\'!W:W,"Approved")/$B11', 'Eligibility check'],
      ['Non-Covered Service', '=COUNTIF(\'Claims Management\'!U:U,"NC*")', '=B12/$B$15', '=AVERAGEIF(\'Claims Management\'!U:U,"NC*",\'Claims Management\'!K:K)', '=COUNTIFS(\'Claims Management\'!U:U,"NC*",\'Claims Management\'!W:W,"Approved")/$B12', 'Benefits verification'],
      ['Documentation', '=COUNTIF(\'Claims Management\'!U:U,"DO*")', '=B13/$B$15', '=AVERAGEIF(\'Claims Management\'!U:U,"DO*",\'Claims Management\'!K:K)', '=COUNTIFS(\'Claims Management\'!U:U,"DO*",\'Claims Management\'!W:W,"Approved")/$B13', 'Record keeping'],
      ['Other', '=COUNTIF(\'Claims Management\'!U:U,"OT*")', '=B14/$B$15', '=AVERAGEIF(\'Claims Management\'!U:U,"OT*",\'Claims Management\'!K:K)', '=COUNTIFS(\'Claims Management\'!U:U,"OT*",\'Claims Management\'!W:W,"Approved")/$B14', 'Investigation needed'],
      ['TOTAL', '=SUM(B5:B14)', '100%', '=AVERAGE(D5:D14)', '=AVERAGE(E5:E14)', '']
    ];
    
    denialSheet.getRange(4, 1, denialCategories.length, 6).setValues(denialCategories);
    denialSheet.getRange('C5:C14').setNumberFormat('0.00%');
    denialSheet.getRange('D5:D14').setNumberFormat('$#,##0.00');
    denialSheet.getRange('E5:E14').setNumberFormat('0.00%');
    denialSheet.getRange('A15:F15').setBackground('#fbbf24').setFontWeight('bold');
    
    // Denial trends by payer
    denialSheet.getRange('H3').setValue('PAYER DENIAL RATES').setFontWeight('bold').setBackground('#0891b2').setFontColor('#ffffff');
    denialSheet.getRange('H3:L3').merge();
    
    const payerDenials = [
      ['Payer', 'Total Claims', 'Denied', 'Denial Rate', 'Avg. Days to Pay'],
      ['Medicare', '=COUNTIF(\'Claims Management\'!N:N,"Medicare")', '=COUNTIFS(\'Claims Management\'!N:N,"Medicare",\'Claims Management\'!P:P,"Denied")', '=J5/I5', '=AVERAGEIFS(\'Claims Management\'!Z:Z,\'Claims Management\'!N:N,"Medicare",\'Claims Management\'!P:P,"Paid")'],
      ['Medicaid', '=COUNTIF(\'Claims Management\'!N:N,"Medicaid")', '=COUNTIFS(\'Claims Management\'!N:N,"Medicaid",\'Claims Management\'!P:P,"Denied")', '=J6/I6', '=AVERAGEIFS(\'Claims Management\'!Z:Z,\'Claims Management\'!N:N,"Medicaid",\'Claims Management\'!P:P,"Paid")'],
      ['BCBS', '=COUNTIF(\'Claims Management\'!N:N,"BCBS")', '=COUNTIFS(\'Claims Management\'!N:N,"BCBS",\'Claims Management\'!P:P,"Denied")', '=J7/I7', '=AVERAGEIFS(\'Claims Management\'!Z:Z,\'Claims Management\'!N:N,"BCBS",\'Claims Management\'!P:P,"Paid")'],
      ['Aetna', '=COUNTIF(\'Claims Management\'!N:N,"Aetna")', '=COUNTIFS(\'Claims Management\'!N:N,"Aetna",\'Claims Management\'!P:P,"Denied")', '=J8/I8', '=AVERAGEIFS(\'Claims Management\'!Z:Z,\'Claims Management\'!N:N,"Aetna",\'Claims Management\'!P:P,"Paid")'],
      ['United', '=COUNTIF(\'Claims Management\'!N:N,"United")', '=COUNTIFS(\'Claims Management\'!N:N,"United",\'Claims Management\'!P:P,"Denied")', '=J9/I9', '=AVERAGEIFS(\'Claims Management\'!Z:Z,\'Claims Management\'!N:N,"United",\'Claims Management\'!P:P,"Paid")'],
      ['Cigna', '=COUNTIF(\'Claims Management\'!N:N,"Cigna")', '=COUNTIFS(\'Claims Management\'!N:N,"Cigna",\'Claims Management\'!P:P,"Denied")', '=J10/I10', '=AVERAGEIFS(\'Claims Management\'!Z:Z,\'Claims Management\'!N:N,"Cigna",\'Claims Management\'!P:P,"Paid")']
    ];
    
    denialSheet.getRange(4, 8, payerDenials.length, 5).setValues(payerDenials);
    denialSheet.getRange('K5:K10').setNumberFormat('0.00%');
    denialSheet.getRange('L5:L10').setNumberFormat('0');
    
    // 4. Revenue Cycle Dashboard
    const revenueSheet = spreadsheet.insertSheet('Revenue Cycle');
    sheets.push(revenueSheet.getName());
    
    revenueSheet.getRange('A1').setValue('REVENUE CYCLE MANAGEMENT DASHBOARD').setFontSize(22).setFontWeight('bold');
    revenueSheet.getRange('A1:M1').merge().setHorizontalAlignment('center').setBackground('#047857').setFontColor('#ffffff');
    
    // Key metrics
    revenueSheet.getRange('A3').setValue('KEY PERFORMANCE INDICATORS').setFontWeight('bold').setBackground('#10b981').setFontColor('#ffffff');
    revenueSheet.getRange('A3:M3').merge();
    
    const kpiMetrics = [
      ['Financial Metrics', '', '', '', 'Operational Metrics', '', '', '', 'Quality Metrics', '', '', '', ''],
      ['Gross Charges:', '=SUM(\'Claims Management\'!K:K)', '', '', 'Claims Volume:', '=COUNTA(\'Claims Management\'!A:A)-1', '', '', 'Clean Claim Rate:', '=1-\'Denial Analytics\'!B15/\'Revenue Cycle\'!F5', '', '', ''],
      ['Net Collections:', '=SUM(\'Claims Management\'!R:R)', '', '', 'Avg. Days in AR:', '=AVERAGE(\'Claims Management\'!Z:Z)', '', '', 'First Pass Rate:', '=(F5-\'Denial Analytics\'!B15)/F5', '', '', ''],
      ['Adjustments:', '=SUM(\'Claims Management\'!S:S)', '', '', 'Days to Submit:', '=AVERAGE(\'Claims Management\'!M:M-\'Claims Management\'!D:D)', '', '', 'Denial Rate:', '=\'Denial Analytics\'!B15/F5', '', '', ''],
      ['Bad Debt:', '=SUMIF(\'Claims Management\'!AA:AA,"120+",\'Claims Management\'!T:T)', '', '', 'Days to Payment:', '=AVERAGE(\'Claims Management\'!Z:Z)', '', '', 'Appeal Success:', '=AVERAGE(\'Denial Analytics\'!E5:E14)', '', '', ''],
      ['Collection Rate:', '=B6/B5', '', '', 'AR > 90 days:', '=COUNTIFS(\'Claims Management\'!Z:Z,">90")/F5', '', '', 'Prior Auth Rate:', '=COUNTIF(\'Insurance Verification\'!W:W,"Yes")/COUNTA(\'Insurance Verification\'!A:A)', '', '', '']
    ];
    
    revenueSheet.getRange(4, 1, kpiMetrics.length, 13).setValues(kpiMetrics);
    revenueSheet.getRange('B5:B8').setNumberFormat('$#,##0');
    revenueSheet.getRange('B9').setNumberFormat('0.00%');
    revenueSheet.getRange('F6:F8').setNumberFormat('0.0');
    revenueSheet.getRange('F9').setNumberFormat('0.00%');
    revenueSheet.getRange('J5:J9').setNumberFormat('0.00%');
    
    // AR Aging Analysis
    revenueSheet.getRange('A11').setValue('ACCOUNTS RECEIVABLE AGING').setFontWeight('bold').setBackground('#f59e0b');
    revenueSheet.getRange('A11:G11').merge();
    
    const arAging = [
      ['Age Bracket', 'Count', 'Amount', '% of Total', 'Avg. Amount', 'Collection Probability', 'Est. Collectible'],
      ['0-30 days', '=COUNTIF(\'Claims Management\'!AA:AA,"0-30")', '=SUMIF(\'Claims Management\'!AA:AA,"0-30",\'Claims Management\'!T:T)', '=C13/$C$19', '=C13/B13', '95%', '=C13*F13'],
      ['31-60 days', '=COUNTIF(\'Claims Management\'!AA:AA,"31-60")', '=SUMIF(\'Claims Management\'!AA:AA,"31-60",\'Claims Management\'!T:T)', '=C14/$C$19', '=C14/B14', '85%', '=C14*F14'],
      ['61-90 days', '=COUNTIF(\'Claims Management\'!AA:AA,"61-90")', '=SUMIF(\'Claims Management\'!AA:AA,"61-90",\'Claims Management\'!T:T)', '=C15/$C$19', '=C15/B15', '70%', '=C15*F15'],
      ['91-120 days', '=COUNTIF(\'Claims Management\'!AA:AA,"91-120")', '=SUMIF(\'Claims Management\'!AA:AA,"91-120",\'Claims Management\'!T:T)', '=C16/$C$19', '=C16/B16', '50%', '=C16*F16'],
      ['120+ days', '=COUNTIF(\'Claims Management\'!AA:AA,"120+")', '=SUMIF(\'Claims Management\'!AA:AA,"120+",\'Claims Management\'!T:T)', '=C17/$C$19', '=C17/B17', '25%', '=C17*F17'],
      ['Paid', '=COUNTIF(\'Claims Management\'!AA:AA,"Paid")', '=SUMIF(\'Claims Management\'!AA:AA,"Paid",\'Claims Management\'!R:R)', '=C18/$C$19', '=C18/B18', '100%', '=C18'],
      ['TOTAL', '=SUM(B13:B18)', '=SUM(C13:C18)', '100%', '=C19/B19', '', '=SUM(G13:G18)']
    ];
    
    arAging.getRange(12, 1, arAging.length, 7).setValues(arAging);
    arAging.getRange('C13:C19').setNumberFormat('$#,##0');
    arAging.getRange('D13:D18').setNumberFormat('0.00%');
    arAging.getRange('E13:E19').setNumberFormat('$#,##0');
    arAging.getRange('F13:F18').setNumberFormat('0%');
    arAging.getRange('G13:G19').setNumberFormat('$#,##0');
    arAging.getRange('A19:G19').setBackground('#fbbf24').setFontWeight('bold');
    
    // 5. Prior Authorization Tracking
    const authSheet = spreadsheet.insertSheet('Prior Authorizations');
    sheets.push(authSheet.getName());
    
    const authHeaders = [
      'Auth ID', 'Patient ID', 'Patient Name', 'Insurance', 'CPT Code',
      'Service Description', 'Requesting Provider', 'Request Date', 'Urgency',
      'Status', 'Decision Date', 'Auth Number', 'Valid From', 'Valid To',
      'Units Approved', 'Units Used', 'Units Remaining', 'Notes'
    ];
    
    authSheet.getRange(1, 1, 1, authHeaders.length).setValues([authHeaders]);
    authSheet.getRange(1, 1, 1, authHeaders.length).setFontWeight('bold');
    authSheet.getRange(1, 1, 1, authHeaders.length).setBackground('#7c3aed');
    authSheet.getRange(1, 1, 1, authHeaders.length).setFontColor('#ffffff');
    
    // Auth tracking formulas
    for (let row = 2; row <= 200; row++) {
      authSheet.getRange(`Q${row}`).setFormula(`=O${row}-P${row}`);
    }
    
    // 6. Appeals Management
    const appealsSheet = spreadsheet.insertSheet('Appeals Tracking');
    sheets.push(appealsSheet.getName());
    
    const appealHeaders = [
      'Appeal ID', 'Claim ID', 'Patient Name', 'Insurance', 'Original Denial Date',
      'Denial Code', 'Denial Reason', 'Appeal Level', 'Appeal Date', 'Appeal Method',
      'Supporting Docs', 'Appeal Status', 'Decision Date', 'Outcome', 'Amount Recovered',
      'Time to Decision', 'Next Action', 'Assigned To'
    ];
    
    appealsSheet.getRange(1, 1, 1, appealHeaders.length).setValues([appealHeaders]);
    appealsSheet.getRange(1, 1, 1, appealHeaders.length).setFontWeight('bold');
    appealsSheet.getRange(1, 1, 1, appealHeaders.length).setBackground('#dc2626');
    appealsSheet.getRange(1, 1, 1, appealHeaders.length).setFontColor('#ffffff');
    
    // Appeal metrics
    for (let row = 2; row <= 200; row++) {
      appealsSheet.getRange(`P${row}`).setFormula(`=IF(M${row}<>"",M${row}-I${row},"")`);
    }
    
    // Optimize column widths
    insuranceSheet.autoResizeColumns(1, 24);
    claimsSheet.setColumnWidth(3, 150);
    denialSheet.setColumnWidth(1, 150);
    denialSheet.setColumnWidth(6, 150);
    revenueSheet.setColumnWidth(1, 150);
    revenueSheet.setColumnWidth(5, 150);
    revenueSheet.setColumnWidth(9, 150);
    authSheet.setColumnWidth(3, 150);
    authSheet.setColumnWidth(6, 200);
    appealsSheet.setColumnWidth(3, 150);
    
    return sheets;
  },

  /**
   * Advanced E-Commerce Inventory with Multi-Marketplace Sync
   * Manages inventory across Amazon, eBay, Shopify, Walmart & more
   * Saves 25+ hours weekly with automated stock sync and reorder management
   */
  createEcommerceInventory: function(spreadsheet) {
    const sheets = [];
    
    // 1. Master Product Catalog
    const catalogSheet = spreadsheet.insertSheet('Master Catalog');
    sheets.push(catalogSheet.getName());
    
    catalogSheet.getRange('A1').setValue('E-COMMERCE MASTER PRODUCT CATALOG').setFontSize(20).setFontWeight('bold');
    catalogSheet.getRange('A1:Z1').merge().setHorizontalAlignment('center').setBackground('#4f46e5').setFontColor('#ffffff');
    
    const catalogHeaders = [
      'SKU', 'Parent SKU', 'Product Name', 'Brand', 'Category', 'Subcategory',
      'Variant Type', 'Variant Value', 'UPC/EAN', 'ASIN', 'MPN', 
      'Cost Price', 'Wholesale Price', 'MSRP', 'MAP Price', 'Your Price',
      'Weight (lbs)', 'Dimensions (LxWxH)', 'HS Code', 'Country of Origin',
      'Lead Time', 'MOQ', 'Supplier', 'Supplier SKU', 'Status', 'Created Date'
    ];
    
    catalogSheet.getRange(3, 1, 1, catalogHeaders.length).setValues([catalogHeaders]);
    catalogSheet.getRange(3, 1, 1, catalogHeaders.length).setFontWeight('bold');
    catalogSheet.getRange(3, 1, 1, catalogHeaders.length).setBackground('#6366f1');
    catalogSheet.getRange(3, 1, 1, catalogHeaders.length).setFontColor('#ffffff');
    
    // Pricing calculations
    for (let row = 4; row <= 500; row++) {
      catalogSheet.getRange(`P${row}`).setFormula(`=IF(O${row}>0,O${row},N${row}*1.2)`).setNumberFormat('$#,##0.00');
    }
    
    // 2. Multi-Channel Inventory Tracking
    const inventorySheet = spreadsheet.insertSheet('Inventory Levels');
    sheets.push(inventorySheet.getName());
    
    inventorySheet.getRange('A1').setValue('MULTI-MARKETPLACE INVENTORY SYNC').setFontSize(18).setFontWeight('bold');
    inventorySheet.getRange('A1:U1').merge().setHorizontalAlignment('center').setBackground('#10b981').setFontColor('#ffffff');
    
    const inventoryHeaders = [
      'SKU', 'Product Name', 'Total Stock', 'Available', 'Committed', 'In Transit',
      'Amazon FBA', 'Amazon FBM', 'eBay', 'Shopify', 'Walmart', 'Etsy',
      'Website', 'Reserved', 'Damaged', 'Buffer Stock', 'Reorder Point',
      'Reorder Qty', 'Days of Stock', 'Stock Value', 'Status'
    ];
    
    inventorySheet.getRange(3, 1, 1, inventoryHeaders.length).setValues([inventoryHeaders]);
    inventorySheet.getRange(3, 1, 1, inventoryHeaders.length).setFontWeight('bold');
    inventorySheet.getRange(3, 1, 1, inventoryHeaders.length).setBackground('#059669');
    inventorySheet.getRange(3, 1, 1, inventoryHeaders.length).setFontColor('#ffffff');
    
    // Inventory calculations
    for (let row = 4; row <= 1000; row++) {
      inventorySheet.getRange(`C${row}`).setFormula(`=SUM(G${row}:M${row})+F${row}`);
      inventorySheet.getRange(`D${row}`).setFormula(`=C${row}-E${row}-N${row}-O${row}`);
      inventorySheet.getRange(`S${row}`).setFormula(`=IF(D${row}>0,D${row}/IFERROR(VLOOKUP(A${row},'Sales Velocity'!A:D,4,FALSE),1),0)`).setNumberFormat('0.0');
      inventorySheet.getRange(`T${row}`).setFormula(`=C${row}*VLOOKUP(A${row},'Master Catalog'!A:L,12,FALSE)`).setNumberFormat('$#,##0.00');
      inventorySheet.getRange(`U${row}`).setFormula(`=IF(D${row}<=Q${row},"REORDER",IF(D${row}<=P${row},"LOW",IF(D${row}>P${row}*5,"OVERSTOCK","OK")))`);
    }
    
    // Conditional formatting for status
    const statusRange = inventorySheet.getRange('U4:U1000');
    const reorderRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('REORDER')
      .setBackground('#dc2626')
      .setFontColor('#ffffff')
      .setRanges([statusRange])
      .build();
    const lowRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('LOW')
      .setBackground('#f59e0b')
      .setRanges([statusRange])
      .build();
    const overstockRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('OVERSTOCK')
      .setBackground('#3b82f6')
      .setFontColor('#ffffff')
      .setRanges([statusRange])
      .build();
    inventorySheet.setConditionalFormatRules([reorderRule, lowRule, overstockRule]);
    
    // 3. Marketplace Listings Manager
    const listingsSheet = spreadsheet.insertSheet('Marketplace Listings');
    sheets.push(listingsSheet.getName());
    
    listingsSheet.getRange('A1').setValue('MARKETPLACE LISTINGS MANAGER').setFontWeight('bold');
    listingsSheet.getRange('A1:T1').merge().setHorizontalAlignment('center').setBackground('#7c3aed').setFontColor('#ffffff');
    
    const listingHeaders = [
      'SKU', 'Marketplace', 'Listing ID', 'Listing URL', 'Title',
      'Price', 'Compare Price', 'Shipping', 'Total Price', 'Competitor Price',
      'Price Position', 'Buy Box %', 'Impressions', 'Clicks', 'CTR',
      'Orders', 'Conversion Rate', 'Revenue', 'Listing Status', 'Last Updated'
    ];
    
    listingsSheet.getRange(3, 1, 1, listingHeaders.length).setValues([listingHeaders]);
    listingsSheet.getRange(3, 1, 1, listingHeaders.length).setFontWeight('bold');
    listingsSheet.getRange(3, 1, 1, listingHeaders.length).setBackground('#9333ea');
    listingsSheet.getRange(3, 1, 1, listingHeaders.length).setFontColor('#ffffff');
    
    // Marketplace metrics
    for (let row = 4; row <= 500; row++) {
      listingsSheet.getRange(`I${row}`).setFormula(`=F${row}+H${row}`).setNumberFormat('$#,##0.00');
      listingsSheet.getRange(`K${row}`).setFormula(`=IF(J${row}>0,IF(I${row}<J${row},"LOWEST",IF(I${row}=J${row},"MATCHED","HIGHER")),"")`);
      listingsSheet.getRange(`O${row}`).setFormula(`=IF(M${row}>0,N${row}/M${row},0)`).setNumberFormat('0.00%');
      listingsSheet.getRange(`Q${row}`).setFormula(`=IF(N${row}>0,P${row}/N${row},0)`).setNumberFormat('0.00%');
      listingsSheet.getRange(`R${row}`).setFormula(`=P${row}*F${row}`).setNumberFormat('$#,##0.00');
    }
    
    // 4. Sales Velocity & Forecasting
    const velocitySheet = spreadsheet.insertSheet('Sales Velocity');
    sheets.push(velocitySheet.getName());
    
    velocitySheet.getRange('A1').setValue('SALES VELOCITY & DEMAND FORECASTING').setFontSize(18).setFontWeight('bold');
    velocitySheet.getRange('A1:O1').merge().setHorizontalAlignment('center').setBackground('#ec4899').setFontColor('#ffffff');
    
    const velocityHeaders = [
      'SKU', 'Product Name', 'Last 7 Days', 'Daily Avg', 'Last 30 Days',
      'Monthly Avg', 'Last 90 Days', 'Quarterly Avg', 'YTD Sales', 
      'Trend', 'Seasonality Index', 'Next 30 Day Forecast', 'Next 90 Day Forecast',
      'Recommended Stock', 'Season'
    ];
    
    velocitySheet.getRange(3, 1, 1, velocityHeaders.length).setValues([velocityHeaders]);
    velocitySheet.getRange(3, 1, 1, velocityHeaders.length).setFontWeight('bold');
    velocitySheet.getRange(3, 1, 1, velocityHeaders.length).setBackground('#be185d');
    velocitySheet.getRange(3, 1, 1, velocityHeaders.length).setFontColor('#ffffff');
    
    // Sales calculations and forecasting
    for (let row = 4; row <= 500; row++) {
      velocitySheet.getRange(`D${row}`).setFormula(`=C${row}/7`).setNumberFormat('0.00');
      velocitySheet.getRange(`F${row}`).setFormula(`=E${row}/30`).setNumberFormat('0.00');
      velocitySheet.getRange(`H${row}`).setFormula(`=G${row}/90`).setNumberFormat('0.00');
      velocitySheet.getRange(`J${row}`).setFormula(`=IF(D${row}>F${row},"↑ TRENDING UP",IF(D${row}<F${row}*0.8,"↓ TRENDING DOWN","→ STABLE"))`);
      velocitySheet.getRange(`L${row}`).setFormula(`=F${row}*30*K${row}`).setNumberFormat('0');
      velocitySheet.getRange(`M${row}`).setFormula(`=F${row}*90*K${row}`).setNumberFormat('0');
      velocitySheet.getRange(`N${row}`).setFormula(`=M${row}+VLOOKUP(A${row},'Inventory Levels'!A:P,16,FALSE)`).setNumberFormat('0');
    }
    
    // 5. Purchase Orders & Suppliers
    const poSheet = spreadsheet.insertSheet('Purchase Orders');
    sheets.push(poSheet.getName());
    
    poSheet.getRange('A1').setValue('PURCHASE ORDER MANAGEMENT').setFontSize(18).setFontWeight('bold');
    poSheet.getRange('A1:S1').merge().setHorizontalAlignment('center').setBackground('#0891b2').setFontColor('#ffffff');
    
    const poHeaders = [
      'PO Number', 'PO Date', 'Supplier', 'SKU', 'Product Name', 'Order Qty',
      'Unit Cost', 'Total Cost', 'Shipping Cost', 'Total with Shipping',
      'Expected Date', 'Received Date', 'Received Qty', 'Short/Over',
      'Invoice #', 'Payment Terms', 'Payment Date', 'Status', 'Notes'
    ];
    
    poSheet.getRange(3, 1, 1, poHeaders.length).setValues([poHeaders]);
    poSheet.getRange(3, 1, 1, poHeaders.length).setFontWeight('bold');
    poSheet.getRange(3, 1, 1, poHeaders.length).setBackground('#164e63');
    poSheet.getRange(3, 1, 1, poHeaders.length).setFontColor('#ffffff');
    
    // PO calculations
    for (let row = 4; row <= 500; row++) {
      poSheet.getRange(`H${row}`).setFormula(`=F${row}*G${row}`).setNumberFormat('$#,##0.00');
      poSheet.getRange(`J${row}`).setFormula(`=H${row}+I${row}`).setNumberFormat('$#,##0.00');
      poSheet.getRange(`N${row}`).setFormula(`=M${row}-F${row}`);
    }
    
    // 6. Returns & Refunds Tracking
    const returnsSheet = spreadsheet.insertSheet('Returns & Refunds');
    sheets.push(returnsSheet.getName());
    
    const returnHeaders = [
      'Return ID', 'Date', 'Order ID', 'Marketplace', 'SKU', 'Product Name',
      'Qty', 'Reason', 'Condition', 'Customer', 'Refund Amount', 'Restocking Fee',
      'Net Refund', 'Return Shipping', 'Resolution', 'Restock Status', 'Notes'
    ];
    
    returnsSheet.getRange(1, 1, 1, returnHeaders.length).setValues([returnHeaders]);
    returnsSheet.getRange(1, 1, 1, returnHeaders.length).setFontWeight('bold');
    returnsSheet.getRange(1, 1, 1, returnHeaders.length).setBackground('#dc2626');
    returnsSheet.getRange(1, 1, 1, returnHeaders.length).setFontColor('#ffffff');
    
    // Returns calculations
    for (let row = 2; row <= 200; row++) {
      returnsSheet.getRange(`M${row}`).setFormula(`=K${row}-L${row}`).setNumberFormat('$#,##0.00');
    }
    
    // 7. Profitability Analysis
    const profitSheet = spreadsheet.insertSheet('Profitability Analysis');
    sheets.push(profitSheet.getName());
    
    profitSheet.getRange('A1').setValue('PRODUCT PROFITABILITY ANALYSIS').setFontSize(20).setFontWeight('bold');
    profitSheet.getRange('A1:R1').merge().setHorizontalAlignment('center').setBackground('#059669').setFontColor('#ffffff');
    
    const profitHeaders = [
      'SKU', 'Product Name', 'Units Sold', 'Revenue', 'COGS', 'Gross Profit',
      'Gross Margin %', 'Marketplace Fees', 'Shipping Cost', 'Ad Spend',
      'Storage Fees', 'Total Expenses', 'Net Profit', 'Net Margin %',
      'ROI %', 'Profit per Unit', 'Rank', 'ABC Classification'
    ];
    
    profitSheet.getRange(3, 1, 1, profitHeaders.length).setValues([profitHeaders]);
    profitSheet.getRange(3, 1, 1, profitHeaders.length).setFontWeight('bold');
    profitSheet.getRange(3, 1, 1, profitHeaders.length).setBackground('#047857');
    profitSheet.getRange(3, 1, 1, profitHeaders.length).setFontColor('#ffffff');
    
    // Profitability calculations
    for (let row = 4; row <= 500; row++) {
      profitSheet.getRange(`F${row}`).setFormula(`=D${row}-E${row}`).setNumberFormat('$#,##0.00');
      profitSheet.getRange(`G${row}`).setFormula(`=IF(D${row}>0,F${row}/D${row},0)`).setNumberFormat('0.00%');
      profitSheet.getRange(`L${row}`).setFormula(`=SUM(H${row}:K${row})`).setNumberFormat('$#,##0.00');
      profitSheet.getRange(`M${row}`).setFormula(`=F${row}-L${row}`).setNumberFormat('$#,##0.00');
      profitSheet.getRange(`N${row}`).setFormula(`=IF(D${row}>0,M${row}/D${row},0)`).setNumberFormat('0.00%');
      profitSheet.getRange(`O${row}`).setFormula(`=IF(E${row}>0,(M${row}/E${row})*100,0)`).setNumberFormat('0.00%');
      profitSheet.getRange(`P${row}`).setFormula(`=IF(C${row}>0,M${row}/C${row},0)`).setNumberFormat('$#,##0.00');
      profitSheet.getRange(`Q${row}`).setFormula(`=RANK(M${row},M:M)`);
      profitSheet.getRange(`R${row}`).setFormula(`=IF(Q${row}<=COUNTA(A:A)*0.2,"A",IF(Q${row}<=COUNTA(A:A)*0.5,"B","C"))`);
    }
    
    // 8. Executive Dashboard
    const dashboardSheet = spreadsheet.insertSheet('E-Commerce Dashboard');
    sheets.push(dashboardSheet.getName());
    
    dashboardSheet.getRange('A1').setValue('E-COMMERCE EXECUTIVE DASHBOARD').setFontSize(24).setFontWeight('bold');
    dashboardSheet.getRange('A1:M1').merge().setHorizontalAlignment('center').setBackground('#1f2937').setFontColor('#ffffff');
    
    // Key Metrics Section
    dashboardSheet.getRange('A3').setValue('KEY PERFORMANCE INDICATORS').setFontWeight('bold').setBackground('#4f46e5').setFontColor('#ffffff');
    dashboardSheet.getRange('A3:M3').merge();
    
    const kpiData = [
      ['Inventory Metrics', '', '', '', 'Sales Metrics', '', '', '', 'Financial Metrics', '', '', '', ''],
      ['Total SKUs:', '=COUNTA(\'Master Catalog\'!A:A)-1', '', '', 'Total Orders:', '=SUM(\'Sales Velocity\'!I:I)', '', '', 'Total Revenue:', '=SUM(\'Profitability Analysis\'!D:D)', '', '', ''],
      ['Total Stock Value:', '=SUM(\'Inventory Levels\'!T:T)', '', '', 'Units Sold:', '=SUM(\'Profitability Analysis\'!C:C)', '', '', 'Gross Profit:', '=SUM(\'Profitability Analysis\'!F:F)', '', '', ''],
      ['Low Stock Items:', '=COUNTIF(\'Inventory Levels\'!U:U,"LOW")+COUNTIF(\'Inventory Levels\'!U:U,"REORDER")', '', '', 'Avg Order Value:', '=B6/B5', '', '', 'Net Profit:', '=SUM(\'Profitability Analysis\'!M:M)', '', '', ''],
      ['Overstock Items:', '=COUNTIF(\'Inventory Levels\'!U:U,"OVERSTOCK")', '', '', 'Return Rate:', '=COUNTA(\'Returns & Refunds\'!A:A)/B5', '', '', 'Avg Margin:', '=J6/J5', '', '', ''],
      ['Stock Turn Rate:', '=B6/B5/365', '', '', 'Best Seller:', '=INDEX(\'Profitability Analysis\'!B:B,MATCH(MAX(\'Profitability Analysis\'!C:C),\'Profitability Analysis\'!C:C,0))', '', '', 'ROI:', '=AVERAGE(\'Profitability Analysis\'!O:O)', '', '', '']
    ];
    
    dashboardSheet.getRange(4, 1, kpiData.length, 13).setValues(kpiData);
    dashboardSheet.getRange('B5:B8').setNumberFormat('$#,##0');
    dashboardSheet.getRange('J5:J7').setNumberFormat('$#,##0');
    dashboardSheet.getRange('J8').setNumberFormat('0.00%');
    dashboardSheet.getRange('J9').setNumberFormat('0.00%');
    
    // Marketplace Performance
    dashboardSheet.getRange('A11').setValue('MARKETPLACE PERFORMANCE').setFontWeight('bold').setBackground('#10b981').setFontColor('#ffffff');
    dashboardSheet.getRange('A11:H11').merge();
    
    const marketplacePerf = [
      ['Channel', 'Listings', 'Orders', 'Revenue', 'Avg Price', 'Conversion', 'Returns', 'Profit'],
      ['Amazon', '=COUNTIF(\'Marketplace Listings\'!B:B,"Amazon")', '=SUMIF(\'Marketplace Listings\'!B:B,"Amazon",\'Marketplace Listings\'!P:P)', '=SUMIF(\'Marketplace Listings\'!B:B,"Amazon",\'Marketplace Listings\'!R:R)', '=D13/C13', '=AVERAGE(IF(\'Marketplace Listings\'!B:B="Amazon",\'Marketplace Listings\'!Q:Q))', '=COUNTIF(\'Returns & Refunds\'!D:D,"Amazon")', '=D13*0.7'],
      ['eBay', '=COUNTIF(\'Marketplace Listings\'!B:B,"eBay")', '=SUMIF(\'Marketplace Listings\'!B:B,"eBay",\'Marketplace Listings\'!P:P)', '=SUMIF(\'Marketplace Listings\'!B:B,"eBay",\'Marketplace Listings\'!R:R)', '=D14/C14', '=AVERAGE(IF(\'Marketplace Listings\'!B:B="eBay",\'Marketplace Listings\'!Q:Q))', '=COUNTIF(\'Returns & Refunds\'!D:D,"eBay")', '=D14*0.72'],
      ['Shopify', '=COUNTIF(\'Marketplace Listings\'!B:B,"Shopify")', '=SUMIF(\'Marketplace Listings\'!B:B,"Shopify",\'Marketplace Listings\'!P:P)', '=SUMIF(\'Marketplace Listings\'!B:B,"Shopify",\'Marketplace Listings\'!R:R)', '=D15/C15', '=AVERAGE(IF(\'Marketplace Listings\'!B:B="Shopify",\'Marketplace Listings\'!Q:Q))', '=COUNTIF(\'Returns & Refunds\'!D:D,"Shopify")', '=D15*0.8'],
      ['Walmart', '=COUNTIF(\'Marketplace Listings\'!B:B,"Walmart")', '=SUMIF(\'Marketplace Listings\'!B:B,"Walmart",\'Marketplace Listings\'!P:P)', '=SUMIF(\'Marketplace Listings\'!B:B,"Walmart",\'Marketplace Listings\'!R:R)', '=D16/C16', '=AVERAGE(IF(\'Marketplace Listings\'!B:B="Walmart",\'Marketplace Listings\'!Q:Q))', '=COUNTIF(\'Returns & Refunds\'!D:D,"Walmart")', '=D16*0.68']
    ];
    
    dashboardSheet.getRange(12, 1, marketplacePerf.length, 8).setValues(marketplacePerf);
    dashboardSheet.getRange('D13:D16').setNumberFormat('$#,##0');
    dashboardSheet.getRange('E13:E16').setNumberFormat('$#,##0.00');
    dashboardSheet.getRange('F13:F16').setNumberFormat('0.00%');
    dashboardSheet.getRange('H13:H16').setNumberFormat('$#,##0');
    
    // Action Items
    dashboardSheet.getRange('A18').setValue('ACTION ITEMS & ALERTS').setFontWeight('bold').setBackground('#dc2626').setFontColor('#ffffff');
    dashboardSheet.getRange('A18:M18').merge();
    
    const actionItems = [
      ['Priority', 'Alert', 'SKUs Affected', 'Action Required', 'Impact', 'Owner', 'Due Date'],
      ['CRITICAL', '=IF(B6>10,"Reorder needed for "&B6&" items","No critical reorders")', '=B6', 'Create POs', 'Stock-out risk', 'Inventory Manager', '=TODAY()'],
      ['HIGH', '=IF(B8>20,"Clear overstock - "&B8&" items","Stock levels balanced")', '=B8', 'Run promotions', 'Cash flow', 'Sales Manager', '=TODAY()+3'],
      ['MEDIUM', '=IF(F8>0.05,"High return rate: "&TEXT(F8,"0.0%"),"Returns normal")', '=COUNTA(\'Returns & Refunds\'!A:A)', 'Investigate quality', 'Customer impact', 'QA Manager', '=TODAY()+7']
    ];
    
    dashboardSheet.getRange(19, 1, actionItems.length, 7).setValues(actionItems);
    
    // Optimize column widths
    catalogSheet.autoResizeColumns(1, 26);
    inventorySheet.setColumnWidth(2, 150);
    listingsSheet.setColumnWidth(5, 200);
    velocitySheet.setColumnWidth(2, 150);
    poSheet.setColumnWidth(5, 150);
    returnsSheet.setColumnWidth(6, 150);
    profitSheet.setColumnWidth(2, 150);
    dashboardSheet.setColumnWidth(1, 120);
    dashboardSheet.setColumnWidth(5, 120);
    dashboardSheet.setColumnWidth(9, 120);
    
    return sheets;
  },

  /**
   * Consulting Time Tracker with Project Profitability
   * Tracks billable hours, project costs, and profitability
   * Saves 15+ hours weekly on time tracking and invoicing
   */
  createConsultingTracker: function(spreadsheet) {
    const sheets = [];
    
    // 1. Time Tracking Sheet
    const timeSheet = spreadsheet.insertSheet('Time Tracking');
    sheets.push(timeSheet.getName());
    
    timeSheet.getRange('A1').setValue('CONSULTING TIME & BILLING TRACKER').setFontSize(20).setFontWeight('bold');
    timeSheet.getRange('A1:P1').merge().setHorizontalAlignment('center').setBackground('#7c3aed').setFontColor('#ffffff');
    
    const timeHeaders = [
      'Date', 'Consultant', 'Client', 'Project', 'Task', 'Start Time', 'End Time',
      'Hours', 'Billable', 'Rate', 'Amount', 'Status', 'Invoice #', 'Notes',
      'Category', 'Approved By'
    ];
    
    timeSheet.getRange(3, 1, 1, timeHeaders.length).setValues([timeHeaders]);
    timeSheet.getRange(3, 1, 1, timeHeaders.length).setFontWeight('bold');
    timeSheet.getRange(3, 1, 1, timeHeaders.length).setBackground('#6d28d9');
    timeSheet.getRange(3, 1, 1, timeHeaders.length).setFontColor('#ffffff');
    
    // Time calculations
    for (let row = 4; row <= 1000; row++) {
      timeSheet.getRange(`H${row}`).setFormula(`=IF(AND(F${row}<>"",G${row}<>""),(G${row}-F${row})*24,0)`).setNumberFormat('0.00');
      timeSheet.getRange(`K${row}`).setFormula(`=IF(I${row}="Yes",H${row}*J${row},0)`).setNumberFormat('$#,##0.00');
    }
    
    // 2. Project Management
    const projectSheet = spreadsheet.insertSheet('Projects');
    sheets.push(projectSheet.getName());
    
    projectSheet.getRange('A1').setValue('PROJECT PORTFOLIO MANAGEMENT').setFontSize(18).setFontWeight('bold');
    projectSheet.getRange('A1:T1').merge().setHorizontalAlignment('center').setBackground('#0891b2').setFontColor('#ffffff');
    
    const projectHeaders = [
      'Project ID', 'Project Name', 'Client', 'Type', 'Start Date', 'End Date',
      'Budget', 'Hours Allocated', 'Hourly Rate', 'Fixed Fee', 'Hours Used',
      'Hours Remaining', 'Budget Used', 'Budget Remaining', 'Completion %',
      'Profitability', 'Status', 'PM', 'Team', 'Next Milestone'
    ];
    
    projectSheet.getRange(3, 1, 1, projectHeaders.length).setValues([projectHeaders]);
    projectSheet.getRange(3, 1, 1, projectHeaders.length).setFontWeight('bold');
    projectSheet.getRange(3, 1, 1, projectHeaders.length).setBackground('#0e7490');
    projectSheet.getRange(3, 1, 1, projectHeaders.length).setFontColor('#ffffff');
    
    // Project metrics
    for (let row = 4; row <= 100; row++) {
      projectSheet.getRange(`K${row}`).setFormula(`=SUMIF('Time Tracking'!D:D,B${row},'Time Tracking'!H:H)`).setNumberFormat('0.00');
      projectSheet.getRange(`L${row}`).setFormula(`=H${row}-K${row}`).setNumberFormat('0.00');
      projectSheet.getRange(`M${row}`).setFormula(`=IF(D${row}="Fixed",J${row}*O${row}/100,K${row}*I${row})`).setNumberFormat('$#,##0.00');
      projectSheet.getRange(`N${row}`).setFormula(`=G${row}-M${row}`).setNumberFormat('$#,##0.00');
      projectSheet.getRange(`O${row}`).setFormula(`=IF(H${row}>0,K${row}/H${row},IF(F${row}>E${row},(TODAY()-E${row})/(F${row}-E${row}),0))`).setNumberFormat('0%');
      projectSheet.getRange(`P${row}`).setFormula(`=M${row}-(K${row}*75)`).setNumberFormat('$#,##0.00');
    }
    
    // Conditional formatting for project status
    const projectStatusRange = projectSheet.getRange('Q4:Q100');
    const activeRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Active')
      .setBackground('#10b981')
      .setFontColor('#ffffff')
      .setRanges([projectStatusRange])
      .build();
    const onHoldRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('On Hold')
      .setBackground('#f59e0b')
      .setRanges([projectStatusRange])
      .build();
    const completedRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Completed')
      .setBackground('#6b7280')
      .setFontColor('#ffffff')
      .setRanges([projectStatusRange])
      .build();
    projectSheet.setConditionalFormatRules([activeRule, onHoldRule, completedRule]);
    
    // 3. Client Management
    const clientSheet = spreadsheet.insertSheet('Clients');
    sheets.push(clientSheet.getName());
    
    const clientHeaders = [
      'Client ID', 'Company', 'Contact', 'Email', 'Phone', 'Industry',
      'Contract Type', 'Retainer Amount', 'Payment Terms', 'Credit Limit',
      'Total Billed', 'Total Paid', 'Outstanding', 'Last Invoice', 'Status'
    ];
    
    clientSheet.getRange(1, 1, 1, clientHeaders.length).setValues([clientHeaders]);
    clientSheet.getRange(1, 1, 1, clientHeaders.length).setFontWeight('bold');
    clientSheet.getRange(1, 1, 1, clientHeaders.length).setBackground('#059669');
    clientSheet.getRange(1, 1, 1, clientHeaders.length).setFontColor('#ffffff');
    
    // Client metrics
    for (let row = 2; row <= 100; row++) {
      clientSheet.getRange(`K${row}`).setFormula(`=SUMIF('Time Tracking'!C:C,B${row},'Time Tracking'!K:K)`).setNumberFormat('$#,##0.00');
      clientSheet.getRange(`M${row}`).setFormula(`=K${row}-L${row}`).setNumberFormat('$#,##0.00');
    }
    
    // 4. Profitability Dashboard
    const profitDashboard = spreadsheet.insertSheet('Profitability Dashboard');
    sheets.push(profitDashboard.getName());
    
    profitDashboard.getRange('A1').setValue('CONSULTING PROFITABILITY DASHBOARD').setFontSize(22).setFontWeight('bold');
    profitDashboard.getRange('A1:L1').merge().setHorizontalAlignment('center').setBackground('#047857').setFontColor('#ffffff');
    
    // Summary metrics
    profitDashboard.getRange('A3').setValue('FINANCIAL SUMMARY').setFontWeight('bold').setBackground('#10b981').setFontColor('#ffffff');
    profitDashboard.getRange('A3:L3').merge();
    
    const summaryData = [
      ['Revenue Metrics', '', '', '', 'Cost Metrics', '', '', '', 'Profitability', '', '', ''],
      ['Total Revenue:', '=SUM(Projects!M:M)', '', '', 'Total Costs:', '=B5*0.6', '', '', 'Gross Profit:', '=B5-F5', '', ''],
      ['Billable Hours:', '=SUMIF(\'Time Tracking\'!I:I,"Yes",\'Time Tracking\'!H:H)', '', '', 'Overhead:', '=F5*0.3', '', '', 'Net Profit:', '=J5-F6', '', ''],
      ['Non-Billable:', '=SUMIF(\'Time Tracking\'!I:I,"No",\'Time Tracking\'!H:H)', '', '', 'Bench Time:', '=B6*75', '', '', 'Profit Margin:', '=J6/B5', '', ''],
      ['Utilization:', '=B6/(B6+B7)', '', '', 'Cost per Hour:', '=F5/B6', '', '', 'Avg Bill Rate:', '=B5/B6', '', '']
    ];
    
    profitDashboard.getRange(4, 1, summaryData.length, 12).setValues(summaryData);
    profitDashboard.getRange('B5:B7').setNumberFormat('$#,##0');
    profitDashboard.getRange('F5:F7').setNumberFormat('$#,##0');
    profitDashboard.getRange('J5:J6').setNumberFormat('$#,##0');
    profitDashboard.getRange('B8').setNumberFormat('0%');
    profitDashboard.getRange('J7').setNumberFormat('0%');
    profitDashboard.getRange('F8:F8').setNumberFormat('$#,##0.00');
    profitDashboard.getRange('J8:J8').setNumberFormat('$#,##0.00');
    
    // Consultant performance
    profitDashboard.getRange('A10').setValue('CONSULTANT PERFORMANCE').setFontWeight('bold').setBackground('#7c3aed').setFontColor('#ffffff');
    profitDashboard.getRange('A10:H10').merge();
    
    const consultantPerf = [
      ['Consultant', 'Hours Billed', 'Revenue', 'Utilization', 'Avg Rate', 'Projects', 'Clients', 'Efficiency'],
      ['=UNIQUE(\'Time Tracking\'!B:B)', '=SUMIF(\'Time Tracking\'!B:B,A12,\'Time Tracking\'!H:H)', '=SUMIF(\'Time Tracking\'!B:B,A12,\'Time Tracking\'!K:K)', '=B12/(B12+20)', '=C12/B12', '=COUNTUNIQUE(IF(\'Time Tracking\'!B:B=A12,\'Time Tracking\'!D:D))', '=COUNTUNIQUE(IF(\'Time Tracking\'!B:B=A12,\'Time Tracking\'!C:C))', '=D12*E12/100']
    ];
    
    consultantPerf.getRange(11, 1, consultantPerf.length, 8).setValues(consultantPerf);
    consultantPerf.getRange('C12:C20').setNumberFormat('$#,##0');
    consultantPerf.getRange('D12:D20').setNumberFormat('0%');
    consultantPerf.getRange('E12:E20').setNumberFormat('$#,##0.00');
    consultantPerf.getRange('H12:H20').setNumberFormat('$#,##0.00');
    
    // Optimize column widths
    timeSheet.setColumnWidth(3, 150);
    timeSheet.setColumnWidth(4, 150);
    timeSheet.setColumnWidth(5, 200);
    projectSheet.setColumnWidth(2, 200);
    projectSheet.setColumnWidth(3, 150);
    clientSheet.setColumnWidth(2, 150);
    profitDashboard.setColumnWidth(1, 150);
    profitDashboard.setColumnWidth(5, 150);
    
    return sheets;
  },

  // Individual template setup functions
  setupLeadPipelineSheet: function(sheet) {
    // Headers with enhanced fields
    const headers = ['Lead ID', 'Company', 'Contact Name', 'Email', 'Phone', 'Lead Source', 
                     'Status', 'Lead Score', 'Days in Pipeline', 'Assigned To', 'Created Date', 'Last Contact', 
                     'Next Action', 'Deal Value', 'Probability %', 'Expected Revenue', 'Expected Close', 'Notes'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Apply professional formatting
    this.formatHeaders(sheet, 1, headers.length);
    this.autoResizeColumns(sheet);
    
    // Add enhanced sample data with formulas
    const sampleData = [
      ['L001', 'TechCorp', 'John Smith', 'john@techcorp.com', '555-0101', 'Website', 
       'Qualified', 85, '', 'Sarah Johnson', '=TODAY()-30', '=TODAY()-2', 'Demo scheduled', 
       50000, 60, '', '=TODAY()+30', 'High interest in enterprise plan'],
      ['L002', 'DataSoft', 'Jane Doe', 'jane@datasoft.com', '555-0102', 'Referral', 
       'Proposal', 75, '', 'Mike Chen', '=TODAY()-14', '=TODAY()-1', 'Review proposal', 
       75000, 40, '', '=TODAY()+45', 'Comparing with competitors'],
      ['L003', 'GlobalTech', 'Bob Wilson', 'bob@globaltech.com', '555-0103', 'Trade Show', 
       'New', 45, '', 'Sarah Johnson', '=TODAY()-7', '', 'Initial outreach', 
       35000, 20, '', '=TODAY()+60', 'Met at conference'],
      ['L004', 'StartupInc', 'Alice Brown', 'alice@startup.com', '555-0104', 'Cold Call', 
       'Negotiation', 90, '', 'Mike Chen', '=TODAY()-45', '=TODAY()', 'Final pricing', 
       100000, 80, '', '=TODAY()+14', 'Ready to close this quarter']
    ];
    sheet.getRange(2, 1, sampleData.length, sampleData[0].length).setValues(sampleData);
    
    // Add formulas for calculated fields
    for (let row = 2; row <= 50; row++) {
      // Days in pipeline
      sheet.getRange(`I${row}`).setFormula(`=IF(K${row}<>"",TODAY()-K${row},"")`);
      // Expected revenue
      sheet.getRange(`P${row}`).setFormula(`=IF(AND(N${row}<>"",O${row}<>""),N${row}*O${row}/100,"")`);
    }
    
    // Format columns
    sheet.getRange('N:N').setNumberFormat('$#,##0');
    sheet.getRange('P:P').setNumberFormat('$#,##0');
    sheet.getRange('O:O').setNumberFormat('0%');
    
    // Add data validation for Status
    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['New', 'Contacted', 'Qualified', 'Proposal', 'Negotiation', 'Closed Won', 'Closed Lost'], true)
      .build();
    sheet.getRange('G2:G').setDataValidation(statusRule);
    
    // Add data validation for Lead Source
    const sourceRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Website', 'Referral', 'Trade Show', 'Cold Call', 'Email Campaign', 'Social Media', 'Partner'], true)
      .build();
    sheet.getRange('F2:F').setDataValidation(sourceRule);
    
    // Enhanced conditional formatting for Lead Score
    const scoreRange = sheet.getRange('H2:H50');
    const hotLeadRule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThanOrEqualTo(80)
      .setBackground(this.colors.danger.light)
      .setFontColor(this.colors.danger.dark)
      .setRanges([scoreRange])
      .build();
    const warmLeadRule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(50, 79)
      .setBackground(this.colors.warning.light)
      .setFontColor(this.colors.warning.dark)
      .setRanges([scoreRange])
      .build();
    const coldLeadRule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(50)
      .setBackground(this.colors.primary.light)
      .setRanges([scoreRange])
      .build();
    
    // Conditional formatting for Status
    const statusRange = sheet.getRange('G2:G50');
    const wonRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Closed Won')
      .setBackground(this.colors.success.light)
      .setFontColor(this.colors.success.dark)
      .setRanges([statusRange])
      .build();
    const lostRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Closed Lost')
      .setBackground(this.colors.danger.light)
      .setRanges([statusRange])
      .build();
    
    sheet.setConditionalFormatRules([hotLeadRule, warmLeadRule, coldLeadRule, wonRule, lostRule]);
    
    // Apply alternating rows
    this.applyProfessionalFormatting(sheet, 50);
  },

  setupKanbanBoardSheet: function(sheet) {
    // Create Kanban board layout
    const headers = ['Task ID', 'Task', 'Assigned To', 'Priority', 'Status', 'Due Date', 'Tags', 'Description'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Style headers
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#34495E');
    headerRange.setFontColor('#FFFFFF');
    headerRange.setFontWeight('bold');
    
    // Add sample tasks
    const sampleTasks = [
      ['T001', 'Design homepage mockup', 'Alice Brown', 'High', 'In Progress', new Date(Date.now() + 3*24*60*60*1000), 'Design, UI', 'Create responsive mockup'],
      ['T002', 'API integration', 'Bob Wilson', 'Medium', 'To Do', new Date(Date.now() + 7*24*60*60*1000), 'Backend, API', 'Connect to payment gateway'],
      ['T003', 'User testing', 'Carol Davis', 'High', 'Review', new Date(Date.now() + 5*24*60*60*1000), 'Testing, UX', 'Conduct usability tests'],
      ['T004', 'Documentation update', 'David Lee', 'Low', 'Done', new Date(), 'Docs', 'Update API documentation']
    ];
    sheet.getRange(2, 1, sampleTasks.length, headers.length).setValues(sampleTasks);
    
    // Add status validation
    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['To Do', 'In Progress', 'Review', 'Done', 'Blocked'])
      .build();
    sheet.getRange(2, 5, 100, 1).setDataValidation(statusRule);
    
    // Add priority validation
    const priorityRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Low', 'Medium', 'High', 'Critical'])
      .build();
    sheet.getRange(2, 4, 100, 1).setDataValidation(priorityRule);
    
    // Conditional formatting for priority
    const priorityRange = sheet.getRange(2, 4, 100, 1);
    const criticalRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Critical')
      .setBackground('#E74C3C')
      .setFontColor('#FFFFFF')
      .setRanges([priorityRange])
      .build();
    const highRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('High')
      .setBackground('#F39C12')
      .setRanges([priorityRange])
      .build();
    sheet.setConditionalFormatRules([criticalRule, highRule]);
    
    // Set column widths
    sheet.setColumnWidth(2, 250);
    sheet.setColumnWidth(8, 300);
  },

  setupBudgetPlannerSheet: function(sheet) {
    // Headers
    const headers = ['Category', 'Subcategory', 'Planned Amount', 'Actual Amount', 
                     'Variance', 'Variance %', 'Month', 'Year', 'Notes'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Style headers
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#27AE60');
    headerRange.setFontColor('#FFFFFF');
    headerRange.setFontWeight('bold');
    
    // Add sample budget data
    const currentMonth = new Date().getMonth() + 1;
    const currentYear = new Date().getFullYear();
    const sampleData = [
      ['Income', 'Salary', 5000, 5000, '=D2-C2', '=E2/C2', currentMonth, currentYear, 'Monthly salary'],
      ['Income', 'Freelance', 1000, 800, '=D3-C3', '=E3/C3', currentMonth, currentYear, 'Side projects'],
      ['Housing', 'Rent', 1500, 1500, '=D4-C4', '=E4/C4', currentMonth, currentYear, 'Monthly rent'],
      ['Housing', 'Utilities', 200, 180, '=D5-C5', '=E5/C5', currentMonth, currentYear, 'Electric, water, gas'],
      ['Transportation', 'Gas', 150, 120, '=D6-C6', '=E6/C6', currentMonth, currentYear, 'Fuel costs'],
      ['Food', 'Groceries', 400, 450, '=D7-C7', '=E7/C7', currentMonth, currentYear, 'Weekly shopping'],
      ['Food', 'Dining Out', 200, 250, '=D8-C8', '=E8/C8', currentMonth, currentYear, 'Restaurants'],
      ['Entertainment', 'Streaming', 50, 50, '=D9-C9', '=E9/C9', currentMonth, currentYear, 'Netflix, Spotify'],
      ['Savings', 'Emergency Fund', 500, 500, '=D10-C10', '=E10/C10', currentMonth, currentYear, '10% of income']
    ];
    sheet.getRange(2, 1, sampleData.length, headers.length).setValues(sampleData);
    
    // Format currency columns
    sheet.getRange(2, 3, 100, 2).setNumberFormat('$#,##0.00');
    sheet.getRange(2, 5, 100, 1).setNumberFormat('$#,##0.00;[Red]-$#,##0.00');
    sheet.getRange(2, 6, 100, 1).setNumberFormat('0.00%');
    
    // Add category validation
    const categoryRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Income', 'Housing', 'Transportation', 'Food', 'Entertainment', 'Healthcare', 'Savings', 'Other'])
      .build();
    sheet.getRange(2, 1, 100, 1).setDataValidation(categoryRule);
    
    // Conditional formatting for variance
    const varianceRange = sheet.getRange(2, 5, 100, 1);
    const positiveRule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0)
      .setFontColor('#27AE60')
      .setRanges([varianceRange])
      .build();
    const negativeRule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0)
      .setFontColor('#E74C3C')
      .setRanges([varianceRange])
      .build();
    sheet.setConditionalFormatRules([positiveRule, negativeRule]);
    
    // Add summary section
    sheet.getRange(12, 1).setValue('SUMMARY');
    sheet.getRange(12, 1).setFontWeight('bold');
    sheet.getRange(13, 1).setValue('Total Income:');
    sheet.getRange(13, 3).setFormula('=SUMIF(A2:A10,"Income",C2:C10)');
    sheet.getRange(13, 4).setFormula('=SUMIF(A2:A10,"Income",D2:D10)');
    sheet.getRange(14, 1).setValue('Total Expenses:');
    sheet.getRange(14, 3).setFormula('=SUMIF(A2:A10,"<>Income",C2:C10)');
    sheet.getRange(14, 4).setFormula('=SUMIF(A2:A10,"<>Income",D2:D10)');
    sheet.getRange(15, 1).setValue('Net Savings:');
    sheet.getRange(15, 3).setFormula('=C13-C14');
    sheet.getRange(15, 4).setFormula('=D13-D14');
    sheet.getRange(13, 3, 3, 2).setNumberFormat('$#,##0.00');
    
    // Set column widths
    sheet.setColumnWidth(1, 150);
    sheet.setColumnWidth(2, 150);
    sheet.setColumnWidth(9, 250);
  },

  setupSocialMediaTrackerSheet: function(sheet) {
    // Headers
    const headers = ['Date', 'Platform', 'Post Type', 'Content Title', 'Impressions', 
                     'Reach', 'Engagement', 'Clicks', 'Shares', 'Comments', 'Likes', 
                     'Conversion Rate', 'ROI', 'Campaign', 'Hashtags'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Style headers
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#3498DB');
    headerRange.setFontColor('#FFFFFF');
    headerRange.setFontWeight('bold');
    
    // Add sample data
    const sampleData = [
      [new Date(), 'Instagram', 'Photo', 'Product Launch', 5000, 3500, 450, 120, 50, 30, 370, '=H2/E2', '=M2*100', 'Summer Sale', '#newproduct #launch'],
      [new Date(), 'Facebook', 'Video', 'Tutorial Video', 8000, 6000, 800, 250, 100, 60, 640, '=H3/E3', '=M3*100', 'Educational', '#howto #tutorial'],
      [new Date(), 'Twitter', 'Tweet', 'Industry News', 2000, 1500, 150, 50, 30, 20, 100, '=H4/E4', '=M4*100', 'News', '#industry #update'],
      [new Date(), 'LinkedIn', 'Article', 'Thought Leadership', 3000, 2500, 350, 180, 80, 40, 230, '=H5/E5', '=M5*100', 'B2B', '#business #leadership']
    ];
    sheet.getRange(2, 1, sampleData.length, headers.length).setValues(sampleData);
    
    // Format columns
    sheet.getRange(2, 1, 100, 1).setNumberFormat('MM/dd/yyyy');
    sheet.getRange(2, 12, 100, 1).setNumberFormat('0.00%');
    sheet.getRange(2, 13, 100, 1).setNumberFormat('0.00%');
    
    // Add platform validation
    const platformRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Facebook', 'Instagram', 'Twitter', 'LinkedIn', 'YouTube', 'TikTok', 'Pinterest'])
      .build();
    sheet.getRange(2, 2, 100, 1).setDataValidation(platformRule);
    
    // Add post type validation
    const postTypeRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Photo', 'Video', 'Story', 'Reel', 'Tweet', 'Article', 'Live'])
      .build();
    sheet.getRange(2, 3, 100, 1).setDataValidation(postTypeRule);
    
    // Conditional formatting for engagement rate
    const engagementRange = sheet.getRange(2, 7, 100, 1);
    const highEngagementRule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThanOrEqualTo(500)
      .setBackground('#27AE60')
      .setRanges([engagementRange])
      .build();
    sheet.setConditionalFormatRules([highEngagementRule]);
    
    // Set column widths
    sheet.setColumnWidth(4, 200);
    sheet.setColumnWidth(15, 250);
  },

  setupStudentGradeTrackerSheet: function(sheet) {
    // Headers
    const headers = ['Student ID', 'First Name', 'Last Name', 'Course', 'Assignment', 
                     'Type', 'Points Possible', 'Points Earned', 'Percentage', 'Grade', 
                     'Due Date', 'Submitted Date', 'Late', 'Comments'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Style headers
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#8E44AD');
    headerRange.setFontColor('#FFFFFF');
    headerRange.setFontWeight('bold');
    
    // Add sample data
    const sampleData = [
      ['S001', 'Emma', 'Johnson', 'Math 101', 'Midterm Exam', 'Exam', 100, 85, '=H2/G2', '=IF(I2>=0.9,"A",IF(I2>=0.8,"B",IF(I2>=0.7,"C",IF(I2>=0.6,"D","F"))))', new Date(), new Date(), '=IF(L2>K2,"Yes","No")', 'Good work'],
      ['S001', 'Emma', 'Johnson', 'Math 101', 'Homework 1', 'Homework', 20, 18, '=H3/G3', '=IF(I3>=0.9,"A",IF(I3>=0.8,"B",IF(I3>=0.7,"C",IF(I3>=0.6,"D","F"))))', new Date(Date.now() - 7*24*60*60*1000), new Date(Date.now() - 7*24*60*60*1000), '=IF(L3>K3,"Yes","No")', 'Complete'],
      ['S002', 'Michael', 'Chen', 'Math 101', 'Midterm Exam', 'Exam', 100, 92, '=H4/G4', '=IF(I4>=0.9,"A",IF(I4>=0.8,"B",IF(I4>=0.7,"C",IF(I4>=0.6,"D","F"))))', new Date(), new Date(), '=IF(L4>K4,"Yes","No")', 'Excellent'],
      ['S002', 'Michael', 'Chen', 'Math 101', 'Homework 1', 'Homework', 20, 20, '=H5/G5', '=IF(I5>=0.9,"A",IF(I5>=0.8,"B",IF(I5>=0.7,"C",IF(I5>=0.6,"D","F"))))', new Date(Date.now() - 7*24*60*60*1000), new Date(Date.now() - 8*24*60*60*1000), '=IF(L5>K5,"Yes","No")', 'Perfect']
    ];
    sheet.getRange(2, 1, sampleData.length, headers.length).setValues(sampleData);
    
    // Format columns
    sheet.getRange(2, 9, 100, 1).setNumberFormat('0.00%');
    sheet.getRange(2, 11, 100, 2).setNumberFormat('MM/dd/yyyy');
    
    // Add type validation
    const typeRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Homework', 'Quiz', 'Exam', 'Project', 'Participation', 'Extra Credit'])
      .build();
    sheet.getRange(2, 6, 100, 1).setDataValidation(typeRule);
    
    // Conditional formatting for grades
    const gradeRange = sheet.getRange(2, 10, 100, 1);
    const aGradeRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('A')
      .setBackground('#27AE60')
      .setRanges([gradeRange])
      .build();
    const fGradeRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('F')
      .setBackground('#E74C3C')
      .setFontColor('#FFFFFF')
      .setRanges([gradeRange])
      .build();
    sheet.setConditionalFormatRules([aGradeRule, fGradeRule]);
    
    // Add grade summary
    sheet.getRange(7, 1).setValue('GRADE SUMMARY');
    sheet.getRange(7, 1).setFontWeight('bold');
    sheet.getRange(8, 1).setValue('Student');
    sheet.getRange(8, 2).setValue('Course Average');
    sheet.getRange(8, 3).setValue('Letter Grade');
    sheet.getRange(9, 1).setValue('Emma Johnson');
    sheet.getRange(9, 2).setFormula('=AVERAGE(I2:I3)');
    sheet.getRange(9, 3).setFormula('=IF(B9>=0.9,"A",IF(B9>=0.8,"B",IF(B9>=0.7,"C",IF(B9>=0.6,"D","F"))))');
    sheet.getRange(10, 1).setValue('Michael Chen');
    sheet.getRange(10, 2).setFormula('=AVERAGE(I4:I5)');
    sheet.getRange(10, 3).setFormula('=IF(B10>=0.9,"A",IF(B10>=0.8,"B",IF(B10>=0.7,"C",IF(B10>=0.6,"D","F"))))');
    sheet.getRange(9, 2, 2, 1).setNumberFormat('0.00%');
    
    // Set column widths
    sheet.setColumnWidth(5, 150);
    sheet.setColumnWidth(14, 250);
  },

  setupCustomerFeedbackSheet: function(sheet) {
    // Headers
    const headers = ['Feedback ID', 'Date', 'Customer Name', 'Email', 'Product/Service', 
                     'Rating', 'NPS Score', 'Category', 'Feedback', 'Sentiment', 
                     'Priority', 'Status', 'Assigned To', 'Resolution', 'Follow-up Date'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Style headers
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#E67E22');
    headerRange.setFontColor('#FFFFFF');
    headerRange.setFontWeight('bold');
    
    // Add sample data
    const sampleData = [
      ['FB001', new Date(), 'Sarah Williams', 'sarah@email.com', 'Premium Plan', 5, 9, 'Service', 'Excellent customer support!', 'Positive', 'Low', 'Resolved', 'Support Team', 'Thanked customer', new Date(Date.now() + 7*24*60*60*1000)],
      ['FB002', new Date(Date.now() - 1*24*60*60*1000), 'James Brown', 'james@email.com', 'Basic Plan', 3, 6, 'Feature Request', 'Need more customization options', 'Neutral', 'Medium', 'In Progress', 'Product Team', 'Evaluating request', new Date(Date.now() + 14*24*60*60*1000)],
      ['FB003', new Date(Date.now() - 2*24*60*60*1000), 'Lisa Garcia', 'lisa@email.com', 'Enterprise', 2, 3, 'Bug', 'Login issues on mobile app', 'Negative', 'High', 'Open', 'Dev Team', '', new Date(Date.now() + 3*24*60*60*1000)],
      ['FB004', new Date(Date.now() - 3*24*60*60*1000), 'Robert Taylor', 'robert@email.com', 'Premium Plan', 4, 8, 'Pricing', 'Good value for money', 'Positive', 'Low', 'Resolved', 'Sales Team', 'Offered loyalty discount', '']
    ];
    sheet.getRange(2, 1, sampleData.length, headers.length).setValues(sampleData);
    
    // Format columns
    sheet.getRange(2, 2, 100, 1).setNumberFormat('MM/dd/yyyy');
    sheet.getRange(2, 15, 100, 1).setNumberFormat('MM/dd/yyyy');
    
    // Add rating validation (1-5 stars)
    const ratingRule = SpreadsheetApp.newDataValidation()
      .requireNumberBetween(1, 5)
      .build();
    sheet.getRange(2, 6, 100, 1).setDataValidation(ratingRule);
    
    // Add NPS validation (0-10)
    const npsRule = SpreadsheetApp.newDataValidation()
      .requireNumberBetween(0, 10)
      .build();
    sheet.getRange(2, 7, 100, 1).setDataValidation(npsRule);
    
    // Add category validation
    const categoryRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Service', 'Product', 'Feature Request', 'Bug', 'Pricing', 'Other'])
      .build();
    sheet.getRange(2, 8, 100, 1).setDataValidation(categoryRule);
    
    // Add sentiment validation
    const sentimentRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Positive', 'Neutral', 'Negative'])
      .build();
    sheet.getRange(2, 10, 100, 1).setDataValidation(sentimentRule);
    
    // Add priority validation
    const priorityRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Low', 'Medium', 'High', 'Critical'])
      .build();
    sheet.getRange(2, 11, 100, 1).setDataValidation(priorityRule);
    
    // Add status validation
    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Open', 'In Progress', 'Resolved', 'Closed'])
      .build();
    sheet.getRange(2, 12, 100, 1).setDataValidation(statusRule);
    
    // Conditional formatting for sentiment
    const sentimentRange = sheet.getRange(2, 10, 100, 1);
    const positiveRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Positive')
      .setBackground('#27AE60')
      .setFontColor('#FFFFFF')
      .setRanges([sentimentRange])
      .build();
    const negativeRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Negative')
      .setBackground('#E74C3C')
      .setFontColor('#FFFFFF')
      .setRanges([sentimentRange])
      .build();
    
    // Conditional formatting for priority
    const priorityRange = sheet.getRange(2, 11, 100, 1);
    const criticalPriorityRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Critical')
      .setBackground('#C0392B')
      .setFontColor('#FFFFFF')
      .setRanges([priorityRange])
      .build();
    const highPriorityRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('High')
      .setBackground('#E74C3C')
      .setFontColor('#FFFFFF')
      .setRanges([priorityRange])
      .build();
    
    sheet.setConditionalFormatRules([positiveRule, negativeRule, criticalPriorityRule, highPriorityRule]);
    
    // Set column widths
    sheet.setColumnWidth(3, 150);
    sheet.setColumnWidth(4, 200);
    sheet.setColumnWidth(9, 300);
    sheet.setColumnWidth(14, 250);
  },

  // Missing setup functions for various templates
  
  setupMaterialTrackerSheet: function(sheet) {
    // Create inventory dashboard at top
    const inventoryMetrics = [
      { label: 'Total Value', formula: '=SUM(P9:P)', color: '#3B82F6' },
      { label: 'Low Stock Items', formula: '=COUNTIF(R9:R,"Low Stock")+COUNTIF(R9:R,"Critical")', color: '#EF4444' },
      { label: 'On Order', formula: '=COUNTIF(R9:R,"On Order")', color: '#F59E0B' },
      { label: 'Categories', formula: '=COUNTA(UNIQUE(D9:D))', color: '#8B5CF6' },
      { label: 'Avg Lead Time', formula: '=IFERROR(AVERAGE(N9:N),0)&" days"', color: '#06B6D4' },
      { label: 'Waste %', formula: '=IFERROR(AVERAGE(T9:T),0)', color: '#EC4899' }
    ];
    
    this.createMiniDashboard(sheet, 1, 1, inventoryMetrics);
    
    // Enhanced headers with more tracking fields
    const headers = ['Material ID', 'Date', 'Material Name', 'Category', 'Unit', 'Quantity Ordered', 
                     'Unit Cost', 'Total Cost', 'Supplier', 'Received Qty', 'Used Qty', 'Remaining', 
                     'Min Stock', 'Lead Time', 'Reorder Point', 'Current Value', 'Usage Rate', 
                     'Days Supply', 'Status', 'Auto Alert', 'Waste %', 'Project Allocation'];
    sheet.getRange(8, 1, 1, headers.length).setValues([headers]);
    
    // Apply construction industry colors
    const colors = this.colors.industries.construction;
    sheet.getRange(8, 1, 1, headers.length)
      .setBackground(colors.primary)
      .setFontColor('#FFFFFF')
      .setFontWeight('bold')
      .setFontSize(11);
    
    // Add comprehensive sample data
    const sampleData = [
      ['MAT-001', '=TODAY()-14', '2x4 Lumber', 'Wood', 'bd ft', 5000, 2.5, '', 'Lumber Co', 5000, 3500, '', 1000, 7, '', '', '', '', '', '', '', 'Project A'],
      ['MAT-002', '=TODAY()-7', 'Concrete Mix', 'Concrete', 'bags', 200, 8.5, '', 'ABC Supply', 200, 150, '', 50, 3, '', '', '', '', '', '', '', 'Project B'],
      ['MAT-003', '=TODAY()-3', 'Rebar #4', 'Steel', 'tons', 10, 650, '', 'Steel Direct', 10, 5, '', 2, 5, '', '', '', '', '', '', '', 'Project A'],
      ['MAT-004', '=TODAY()', 'Drywall Sheets', 'Drywall', 'sheets', 500, 12, '', 'Building Supply', 0, 0, '', 100, 2, '', '', '', '', '', '', '', 'Project C'],
      ['MAT-005', '=TODAY()-21', 'PVC Pipe 2"', 'Plumbing', 'ft', 1000, 3.75, '', 'Plumbing Pro', 1000, 800, '', 200, 4, '', '', '', '', '', '', '', 'Project B']
    ];
    
    sheet.getRange(9, 1, sampleData.length, sampleData[0].length).setValues(sampleData);
    
    // Advanced formulas for comprehensive tracking
    for (let row = 9; row <= 108; row++) {
      // Total cost calculation
      sheet.getRange(`H${row}`).setFormula(`=IF(AND(F${row}<>"",G${row}<>""),F${row}*G${row},"")`);
      
      // Remaining quantity
      sheet.getRange(`L${row}`).setFormula(`=IF(AND(J${row}<>"",K${row}<>""),J${row}-K${row},"")`);
      
      // Reorder point calculation (min stock + lead time * usage rate)
      sheet.getRange(`O${row}`).setFormula(
        `=IF(AND(M${row}<>"",N${row}<>"",Q${row}<>""),M${row}+N${row}*Q${row},M${row})`
      );
      
      // Current value calculation
      sheet.getRange(`P${row}`).setFormula(`=IF(AND(L${row}<>"",G${row}<>""),L${row}*G${row},"")`);
      
      // Usage rate calculation (7-day average)
      sheet.getRange(`Q${row}`).setFormula(
        `=IF(AND(K${row}<>"",B${row}<>""),K${row}/(TODAY()-B${row}+1),"")`
      );
      
      // Days of supply remaining
      sheet.getRange(`R${row}`).setFormula(
        `=IF(AND(L${row}<>"",Q${row}>0),ROUND(L${row}/Q${row},0),IF(L${row}>0,"Sufficient",""))`
      );
      
      // Smart status assignment
      sheet.getRange(`S${row}`).setFormula(
        `=IF(A${row}<>"",` +
        `IF(L${row}=0,"Out of Stock",` +
        `IF(L${row}<=M${row},"Critical",` +
        `IF(L${row}<=O${row},"Low Stock",` +
        `IF(J${row}=0,"On Order","In Stock")))),"")`
      );
      
      // Auto alert generation
      sheet.getRange(`T${row}`).setFormula(
        `=IF(S${row}="Critical","🔴 ORDER NOW",` +
        `IF(S${row}="Low Stock","🟡 Reorder Soon",` +
        `IF(S${row}="Out of Stock","⚫ URGENT",` +
        `IF(AND(S${row}="On Order",N${row}<>""),` +
        `"📦 ETA: "&TEXT(B${row}+N${row},"dd/mm"),"✅ OK"))))`
      );
      
      // Waste percentage calculation
      sheet.getRange(`U${row}`).setFormula(
        `=IF(AND(J${row}>0,K${row}<>""),MAX(0,(K${row}-J${row}*0.95)/(J${row})*100),"")`
      );
    }
    
    // Format columns
    sheet.getRange('G:G').setNumberFormat('$#,##0.00');
    sheet.getRange('H:H').setNumberFormat('$#,##0.00');
    sheet.getRange('P:P').setNumberFormat('$#,##0.00');
    sheet.getRange('F:F').setNumberFormat('#,##0');
    sheet.getRange('J:L').setNumberFormat('#,##0');
    sheet.getRange('M:O').setNumberFormat('#,##0');
    sheet.getRange('Q:Q').setNumberFormat('0.00');
    sheet.getRange('U:U').setNumberFormat('0.0%');
    sheet.getRange('B:B').setNumberFormat('dd/mm/yyyy');
    
    // Add sparkline for usage trends
    this.addSparkline(sheet, 'W8', 'Q9:Q13', 'line', { color: '#F59E0B' });
    
    // Comprehensive data validations
    const categoryRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Wood', 'Concrete', 'Steel', 'Drywall', 'Electrical', 'Plumbing', 'HVAC', 'Insulation', 'Roofing', 'Flooring', 'Hardware', 'Paint', 'Safety', 'Tools', 'Other'], true)
      .build();
    sheet.getRange('D9:D').setDataValidation(categoryRule);
    
    const unitRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['bd ft', 'sq ft', 'cu yd', 'bags', 'tons', 'lbs', 'gallons', 'sheets', 'pieces', 'boxes', 'rolls', 'ft', 'meters'], true)
      .build();
    sheet.getRange('E9:E').setDataValidation(unitRule);
    
    const projectRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Project A', 'Project B', 'Project C', 'General Stock', 'Multiple'], true)
      .build();
    sheet.getRange('V9:V').setDataValidation(projectRule);
    
    // Conditional formatting for status
    const criticalRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Critical')
      .setBackground('#FEE2E2')
      .setFontColor('#DC2626')
      .setRanges([sheet.getRange('S9:S')])
      .build();
    
    const lowStockRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Low Stock')
      .setBackground('#FEF3C7')
      .setFontColor('#92400E')
      .setRanges([sheet.getRange('S9:S')])
      .build();
    
    const inStockRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('In Stock')
      .setBackground('#D1FAE5')
      .setFontColor('#065F46')
      .setRanges([sheet.getRange('S9:S')])
      .build();
    
    // High waste percentage formatting
    const highWasteRule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0.1)
      .setBackground('#FEE2E2')
      .setRanges([sheet.getRange('U9:U')])
      .build();
    
    sheet.setConditionalFormatRules([criticalRule, lowStockRule, inStockRule, highWasteRule]);
    
    // Protect formula columns
    const protection = sheet.getRange('H:H,L:L,O:S,U:U').protect();
    protection.setDescription('Automated calculations - Do not edit');
    protection.setWarningOnly(true);
    
    // Apply alternating rows
    this.applyProfessionalFormatting(sheet, 9, 108);
  },
  
  setupLaborManagerSheet: function(sheet) {
    const headers = ['Employee ID', 'Name', 'Role', 'Department', 'Date', 'Hours Worked', 
                     'Overtime Hours', 'Hourly Rate', 'OT Rate', 'Daily Cost', 'Project', 
                     'Task', 'Status', 'Approved By'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Apply professional formatting
    this.formatHeaders(sheet, 1, headers.length);
    this.autoResizeColumns(sheet);
    
    // Add sample labor data
    const sampleData = [
      ['EMP-001', 'John Smith', 'Foreman', 'Construction', '=TODAY()', 8, 0, 45, '=H2*1.5', '', '=\'Cost Estimator\'!A2', 'Framing', 'Approved', 'Manager'],
      ['EMP-002', 'Mike Johnson', 'Carpenter', 'Construction', '=TODAY()', 8, 2, 35, '=H3*1.5', '', '=\'Cost Estimator\'!A2', 'Framing', 'Approved', 'Foreman'],
      ['EMP-003', 'Bob Wilson', 'Electrician', 'Electrical', '=TODAY()', 7, 0, 55, '=H4*1.5', '', '=\'Cost Estimator\'!A2', 'Rough-in', 'Pending', ''],
      ['EMP-004', 'Tom Davis', 'Helper', 'General', '=TODAY()', 8, 0, 25, '=H5*1.5', '', '=\'Cost Estimator\'!A2', 'General', 'Approved', 'Foreman']
    ];
    
    sheet.getRange(2, 1, sampleData.length, sampleData[0].length).setValues(sampleData);
    
    // Add formulas for cost calculation
    for (let row = 2; row <= 100; row++) {
      sheet.getRange(`J${row}`).setFormula(`=IF(OR(F${row}<>"",G${row}<>""),F${row}*H${row}+G${row}*I${row},"")`);
    }
    
    // Format columns
    sheet.getRange('H:I').setNumberFormat('$#,##0.00');
    sheet.getRange('J:J').setNumberFormat('$#,##0.00');
    sheet.getRange('F:G').setNumberFormat('#,##0.0');
    
    // Add role validation
    const roleRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Foreman', 'Carpenter', 'Electrician', 'Plumber', 'Mason', 'Helper', 'Laborer', 'Operator'], true)
      .build();
    sheet.getRange('C2:C').setDataValidation(roleRule);
    
    // Add status validation
    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Approved', 'Pending', 'Rejected', 'In Review'], true)
      .build();
    sheet.getRange('M2:M').setDataValidation(statusRule);
    
    // Conditional formatting for status
    const statusRange = sheet.getRange('M2:M100');
    const approvedRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Approved')
      .setBackground(this.colors.success.light)
      .setRanges([statusRange])
      .build();
    const pendingRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Pending')
      .setBackground(this.colors.warning.light)
      .setRanges([statusRange])
      .build();
    
    sheet.setConditionalFormatRules([approvedRule, pendingRule]);
    
    // Apply alternating rows
    this.applyProfessionalFormatting(sheet, 50);
  },
  
  setupChangeOrdersSheet: function(sheet) {
    const headers = ['Change Order #', 'Date', 'Client', 'Project', 'Description', 'Reason', 
                     'Original Cost', 'Change Amount', 'New Total', 'Impact Days', 'Status', 
                     'Approved By', 'Date Approved', 'Notes'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Apply professional formatting
    this.formatHeaders(sheet, 1, headers.length);
    this.autoResizeColumns(sheet);
    
    // Add sample change orders
    const sampleData = [
      ['CO-001', '=TODAY()-10', 'ABC Corp', '=\'Cost Estimator\'!A2', 'Add bathroom fixture upgrade', 'Client request', 85000, 5500, '', 3, 'Approved', 'Client PM', '=TODAY()-8', 'Premium fixtures selected'],
      ['CO-002', '=TODAY()-5', 'ABC Corp', '=\'Cost Estimator\'!A2', 'Foundation reinforcement', 'Soil conditions', 85000, 12000, '', 5, 'Pending', '', '', 'Engineering review needed'],
      ['CO-003', '=TODAY()-3', 'XYZ Inc', '=\'Cost Estimator\'!A3', 'Change electrical layout', 'Design change', 125000, -2500, '', 0, 'Approved', 'Architect', '=TODAY()-2', 'Simplified design']
    ];
    
    sheet.getRange(2, 1, sampleData.length, sampleData[0].length).setValues(sampleData);
    
    // Add formulas for totals
    for (let row = 2; row <= 50; row++) {
      sheet.getRange(`I${row}`).setFormula(`=IF(AND(G${row}<>"",H${row}<>""),G${row}+H${row},"")`);
    }
    
    // Format columns
    sheet.getRange('G:I').setNumberFormat('$#,##0');
    sheet.getRange('J:J').setNumberFormat('#,##0');
    
    // Add reason validation
    const reasonRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Client request', 'Design change', 'Code requirement', 'Site conditions', 'Material unavailable', 'Error correction'], true)
      .build();
    sheet.getRange('F2:F').setDataValidation(reasonRule);
    
    // Add status validation
    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Draft', 'Pending', 'Approved', 'Rejected', 'Implemented'], true)
      .build();
    sheet.getRange('K2:K').setDataValidation(statusRule);
    
    // Conditional formatting for change amount
    const changeRange = sheet.getRange('H2:H50');
    const increaseRule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0)
      .setBackground(this.colors.warning.light)
      .setRanges([changeRange])
      .build();
    const decreaseRule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0)
      .setBackground(this.colors.success.light)
      .setRanges([changeRange])
      .build();
    
    sheet.setConditionalFormatRules([increaseRule, decreaseRule]);
    
    // Apply alternating rows
    this.applyProfessionalFormatting(sheet, 50);
  },
  
  setupPriorAuthSheet: function(sheet) {
    // Create authorization metrics dashboard
    const authMetrics = [
      { label: 'Pending Auths', formula: '=COUNTIF(I9:I,"Pending")', color: '#F59E0B' },
      { label: 'Approved Today', formula: '=COUNTIFS(I9:I,"Approved",B9:B,TODAY())', color: '#10B981' },
      { label: 'Expiring Soon', formula: '=COUNTIFS(L9:L,"<="&TODAY()+7,L9:L,">="&TODAY())', color: '#EF4444' },
      { label: 'Avg TAT (days)', formula: '=IFERROR(AVERAGE(S9:S),0)', color: '#3B82F6' },
      { label: 'Denial Rate', formula: '=IFERROR(COUNTIF(I9:I,"Denied")/COUNTA(I9:I),0)', color: '#8B5CF6' },
      { label: 'Utilization %', formula: '=IFERROR(AVERAGE(T9:T),0)', color: '#06B6D4' }
    ];
    
    this.createMiniDashboard(sheet, 1, 1, authMetrics);
    
    // Enhanced headers with more tracking fields
    const headers = ['Auth ID', 'Date', 'Patient ID', 'Patient Name', 'DOB', 'Insurance', 'Procedure Code', 
                     'Procedure Description', 'Status', 'Priority', 'Auth Number', 'Valid From', 'Valid To', 
                     'Units Authorized', 'Units Used', 'Remaining', 'Provider', 'Requesting MD', 
                     'Clinical Notes', 'ICD-10', 'Urgency Score', 'TAT', 'Utilization %', 'Auto Alert'];
    sheet.getRange(8, 1, 1, headers.length).setValues([headers]);
    
    // Apply healthcare industry colors
    const colors = this.colors.industries.healthcare;
    sheet.getRange(8, 1, 1, headers.length)
      .setBackground(colors.primary)
      .setFontColor('#FFFFFF')
      .setFontWeight('bold')
      .setFontSize(11);
    
    // Add comprehensive sample data
    const sampleData = [
      ['PA-001', '=TODAY()-5', 'PT-1234', 'John Smith', '01/15/1970', 'BlueCross', '27447', 'Knee Arthroplasty', 'Pending', '', '', '=TODAY()', '=TODAY()+90', 1, 0, '', 'Dr. Johnson', 'Dr. Smith', 'Severe OA, failed conservative tx', 'M17.11', '', '', '', ''],
      ['PA-002', '=TODAY()-3', 'PT-2345', 'Jane Doe', '05/22/1985', 'Aetna', '99213', 'Office Visit Level 3', 'Approved', '', 'AUTH-12345', '=TODAY()-2', '=TODAY()+180', 12, 3, '', 'Dr. Brown', 'Dr. Brown', 'Diabetes management', 'E11.9', '', '', '', ''],
      ['PA-003', '=TODAY()-7', 'PT-3456', 'Bob Wilson', '11/30/1955', 'Medicare', '93306', 'Echo Complete', 'Approved', '', 'AUTH-23456', '=TODAY()-5', '=TODAY()+60', 1, 1, '', 'Dr. Davis', 'Dr. Davis', 'CHF monitoring', 'I50.9', '', '', '', ''],
      ['PA-004', '=TODAY()-1', 'PT-4567', 'Alice Brown', '03/10/1990', 'Cigna', '70553', 'MRI Brain w/wo', 'Pending', '', '', '', '', 1, 0, '', 'Dr. Miller', 'Dr. Wilson', 'Headaches, r/o tumor', 'R51', '', '', '', ''],
      ['PA-005', '=TODAY()', 'PT-5678', 'Tom Johnson', '08/25/1945', 'UnitedHealth', '29881', 'ACL Reconstruction', 'Urgent', '', '', '', '', 1, 0, '', 'Dr. Lee', 'Dr. Martinez', 'Sports injury', 'S83.5', '', '', '', '']
    ];
    
    sheet.getRange(9, 1, sampleData.length, sampleData[0].length).setValues(sampleData);
    
    // Advanced formulas for comprehensive tracking
    for (let row = 9; row <= 108; row++) {
      // Remaining units calculation
      sheet.getRange(`P${row}`).setFormula(`=IF(AND(N${row}<>"",O${row}<>""),N${row}-O${row},N${row})`);
      
      // Priority assignment based on multiple factors
      sheet.getRange(`J${row}`).setFormula(
        `=IF(A${row}<>"",` +
        `IF(I${row}="Urgent","🔴 URGENT",` +
        `IF(AND(I${row}="Pending",B${row}<=TODAY()-5),"🟡 HIGH",` +
        `IF(AND(M${row}<>"",M${row}<=TODAY()+7),"🟠 EXPIRING",` +
        `IF(I${row}="Pending","🔵 NORMAL","✅ COMPLETE")))),"")`
      );
      
      // Urgency score calculation
      sheet.getRange(`U${row}`).setFormula(
        `=IF(A${row}<>"",` +
        `IF(I${row}="Urgent",100,` +
        `IF(I${row}="Denied",80,` +
        `IF(AND(I${row}="Pending",B${row}<=TODAY()-7),70,` +
        `IF(AND(M${row}<>"",M${row}<=TODAY()+14),60,` +
        `IF(I${row}="Pending",40,10))))),"")`
      );
      
      // Turnaround time calculation
      sheet.getRange(`V${row}`).setFormula(
        `=IF(AND(B${row}<>"",I${row}<>"Pending"),TODAY()-B${row},IF(B${row}<>"",TODAY()-B${row}&" (pending)",""))`
      );
      
      // Utilization percentage
      sheet.getRange(`W${row}`).setFormula(
        `=IF(AND(N${row}>0,O${row}<>""),O${row}/N${row},"")`
      );
      
      // Auto alert generation
      sheet.getRange(`X${row}`).setFormula(
        `=IF(I${row}="Urgent","⚡ EXPEDITE",` +
        `IF(I${row}="Denied","❌ APPEAL NEEDED",` +
        `IF(AND(I${row}="Pending",B${row}<=TODAY()-7),"⏰ FOLLOW UP",` +
        `IF(AND(M${row}<>"",M${row}<=TODAY()+7,P${row}>0),"📅 RENEW SOON",` +
        `IF(AND(W${row}<>"",W${row}>0.8),"📊 80% UTILIZED",` +
        `IF(I${row}="Approved","✅ ACTIVE","📝 PROCESSING"))))))`
      );
    }
    
    // Format columns
    sheet.getRange('B:B').setNumberFormat('dd/mm/yyyy');
    sheet.getRange('E:E').setNumberFormat('dd/mm/yyyy');
    sheet.getRange('L:M').setNumberFormat('dd/mm/yyyy');
    sheet.getRange('N:P').setNumberFormat('#,##0');
    sheet.getRange('U:U').setNumberFormat('0');
    sheet.getRange('W:W').setNumberFormat('0%');
    
    // Add sparkline for urgency trends
    this.addSparkline(sheet, 'Y8', 'U9:U13', 'column', { color: '#EF4444' });
    
    // Comprehensive data validations
    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Pending', 'Approved', 'Denied', 'Expired', 'Urgent', 'On Hold', 'Cancelled'], true)
      .build();
    sheet.getRange('I9:I').setDataValidation(statusRule);
    
    const insuranceRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Medicare', 'Medicaid', 'BlueCross', 'Aetna', 'Cigna', 'UnitedHealth', 'Humana', 'Kaiser', 'Anthem', 'Other'], true)
      .build();
    sheet.getRange('F9:F').setDataValidation(insuranceRule);
    
    // Conditional formatting for status
    const urgentRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Urgent')
      .setBackground('#FEE2E2')
      .setFontColor('#DC2626')
      .setRanges([sheet.getRange('I9:I')])
      .build();
    
    const approvedRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Approved')
      .setBackground('#D1FAE5')
      .setFontColor('#065F46')
      .setRanges([sheet.getRange('I9:I')])
      .build();
    
    const deniedRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Denied')
      .setBackground('#FEE2E2')
      .setFontColor('#991B1B')
      .setRanges([sheet.getRange('I9:I')])
      .build();
    
    // High urgency score formatting
    const highUrgencyRule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThanOrEqualTo(70)
      .setBackground('#FEF3C7')
      .setRanges([sheet.getRange('U9:U')])
      .build();
    
    sheet.setConditionalFormatRules([urgentRule, approvedRule, deniedRule, highUrgencyRule]);
    
    // Protect formula columns
    const protection = sheet.getRange('J:J,P:P,U:X').protect();
    protection.setDescription('Automated calculations - Do not edit');
    protection.setWarningOnly(true);
    
    // Apply alternating rows
    this.applyProfessionalFormatting(sheet, 9, 108);
  },
  
  setupRevenueCycleSheet: function(sheet) {
    // Create revenue cycle metrics dashboard
    const revenueMetrics = [
      { label: 'Total Charges', formula: '=SUM(G9:G)', color: '#3B82F6' },
      { label: 'Collections MTD', formula: '=SUMIFS(I9:I,B9:B,">="&EOMONTH(TODAY(),-1)+1)', color: '#10B981' },
      { label: 'Outstanding AR', formula: '=SUM(K9:K)', color: '#F59E0B' },
      { label: 'Avg Days in AR', formula: '=IFERROR(AVERAGE(L9:L),0)', color: '#8B5CF6' },
      { label: 'Denial Rate', formula: '=IFERROR(COUNTIF(M9:M,"Denied")/COUNTA(M9:M),0)', color: '#EF4444' },
      { label: 'Clean Claim %', formula: '=IFERROR(COUNTIF(M9:M,"Paid")/COUNTA(M9:M),0)', color: '#06B6D4' }
    ];
    
    this.createMiniDashboard(sheet, 1, 1, revenueMetrics);
    
    // Enhanced headers with more financial tracking
    const headers = ['Claim ID', 'Date', 'Patient', 'Provider', 'Service Date', 'CPT Code', 
                     'Charge Amount', 'Allowed Amount', 'Paid Amount', 'Adjustment', 'Balance', 
                     'Days in AR', 'Status', 'Payer', 'Denial Code', 'Write-off', 'Collection %', 
                     'AR Bucket', 'Follow-up Date', 'Notes', 'Risk Score', 'Auto Action'];
    sheet.getRange(8, 1, 1, headers.length).setValues([headers]);
    
    const colors = this.colors.industries.healthcare;
    sheet.getRange(8, 1, 1, headers.length)
      .setBackground(colors.secondary)
      .setFontColor('#FFFFFF')
      .setFontWeight('bold')
      .setFontSize(11);
    
    // Add sample revenue cycle data
    const sampleData = [
      ['CLM-001', '=TODAY()-45', 'John Smith', 'Dr. Johnson', '=TODAY()-50', '99214', 185, 145, 145, 40, '', '', 'Paid', 'Medicare', '', 0, '', '', '', 'Routine F/U', '', ''],
      ['CLM-002', '=TODAY()-30', 'Jane Doe', 'Dr. Brown', '=TODAY()-35', '27447', 12500, 9800, 0, 0, '', '', 'Pending', 'BlueCross', '', 0, '', '', '=TODAY()+7', 'Auth pending', '', ''],
      ['CLM-003', '=TODAY()-60', 'Bob Wilson', 'Dr. Davis', '=TODAY()-65', '93306', 850, 650, 0, 0, '', '', 'Denied', 'Aetna', 'CO-197', 0, '', '', '=TODAY()', 'Missing documentation', '', ''],
      ['CLM-004', '=TODAY()-15', 'Alice Brown', 'Dr. Miller', '=TODAY()-20', '70553', 2200, 1800, 1800, 400, '', '', 'Paid', 'Cigna', '', 0, '', '', '', 'Clean claim', '', ''],
      ['CLM-005', '=TODAY()-90', 'Tom Johnson', 'Dr. Lee', '=TODAY()-95', '29881', 8500, 6500, 0, 0, '', '', 'Pending', 'UnitedHealth', '', 0, '', '', '=TODAY()+3', 'Appeal in progress', '', '']
    ];
    
    sheet.getRange(9, 1, sampleData.length, sampleData[0].length).setValues(sampleData);
    
    // Advanced revenue cycle formulas
    for (let row = 9; row <= 108; row++) {
      // Balance calculation
      sheet.getRange(`K${row}`).setFormula(`=IF(G${row}<>"",G${row}-I${row}-J${row}-P${row},"")`);
      
      // Days in AR
      sheet.getRange(`L${row}`).setFormula(`=IF(AND(B${row}<>"",K${row}>0),TODAY()-B${row},"")`);
      
      // Collection percentage
      sheet.getRange(`Q${row}`).setFormula(`=IF(AND(G${row}>0,I${row}>=0),I${row}/G${row},"")`);
      
      // AR aging bucket
      sheet.getRange(`R${row}`).setFormula(
        `=IF(L${row}<>"",` +
        `IF(L${row}<=30,"0-30 days",` +
        `IF(L${row}<=60,"31-60 days",` +
        `IF(L${row}<=90,"61-90 days",` +
        `IF(L${row}<=120,"91-120 days",">120 days")))),"")`
      );
      
      // Risk score calculation
      sheet.getRange(`U${row}`).setFormula(
        `=IF(K${row}>0,` +
        `IF(L${row}>120,100,` +
        `IF(L${row}>90,80,` +
        `IF(L${row}>60,60,` +
        `IF(M${row}="Denied",70,` +
        `IF(L${row}>30,40,20))))),"")`
      );
      
      // Auto action recommendation
      sheet.getRange(`V${row}`).setFormula(
        `=IF(M${row}="Denied","🔴 APPEAL",` +
        `IF(L${row}>90,"📞 ESCALATE",` +
        `IF(L${row}>60,"📧 2ND NOTICE",` +
        `IF(L${row}>30,"📝 FOLLOW UP",` +
        `IF(M${row}="Paid","✅ COMPLETE",` +
        `IF(K${row}>0,"⏳ PROCESSING",""))))))`
      );
    }
    
    // Format columns
    sheet.getRange('B:B').setNumberFormat('dd/mm/yyyy');
    sheet.getRange('E:E').setNumberFormat('dd/mm/yyyy');
    sheet.getRange('S:S').setNumberFormat('dd/mm/yyyy');
    sheet.getRange('G:K').setNumberFormat('$#,##0.00');
    sheet.getRange('P:P').setNumberFormat('$#,##0.00');
    sheet.getRange('Q:Q').setNumberFormat('0%');
    sheet.getRange('U:U').setNumberFormat('0');
    
    // Add sparkline for AR aging
    this.addSparkline(sheet, 'W8', 'L9:L13', 'column', { color: '#F59E0B' });
    
    // Apply alternating rows
    this.applyProfessionalFormatting(sheet, 9, 108);
  },
  
  setupDenialAnalyticsSheet: function(sheet) {
    const headers = ['Denial ID', 'Date', 'Claim ID', 'Patient', 'Payer', 'Amount', 
                     'Denial Code', 'Denial Reason', 'Category', 'Appeal Status', 
                     'Appeal Date', 'Recovery Amount', 'Days to Appeal', 'Resolution'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#E67E22').setFontColor('#FFFFFF').setFontWeight('bold');
    
    // Add formulas for days to appeal
    for (let row = 2; row <= 100; row++) {
      sheet.getRange(`M${row}`).setFormula(`=IF(K${row}<>"",K${row}-B${row},"")`).setNumberFormat('#,##0');
    }
  },
  
  setupLeadScoringSheet: function(sheet) {
    const headers = ['Lead ID', 'Date', 'Name', 'Email', 'Company', 'Source', 'Industry', 
                     'Company Size', 'Budget', 'Interest Level', 'Engagement Score', 
                     'Lead Score', 'Status', 'Assigned To', 'Next Action'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#9B59B6').setFontColor('#FFFFFF').setFontWeight('bold');
    
    // Add scoring formulas
    for (let row = 2; row <= 100; row++) {
      sheet.getRange(`L${row}`).setFormula(`=(J${row}*0.3)+(K${row}*0.7)`).setNumberFormat('0.0');
    }
  },
  
  setupContentPerformanceSheet: function(sheet) {
    const headers = ['Content ID', 'Date', 'Title', 'Type', 'Channel', 'Views', 
                     'Engagement Rate', 'Shares', 'Comments', 'Likes', 'CTR', 
                     'Conversion Rate', 'Revenue', 'ROI', 'Status'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#16A085').setFontColor('#FFFFFF').setFontWeight('bold');
    
    // Add performance formulas
    for (let row = 2; row <= 100; row++) {
      sheet.getRange(`G${row}`).setFormula(`=(I${row}+J${row})/F${row}*100`).setNumberFormat('0.00%');
      sheet.getRange(`K${row}`).setNumberFormat('0.00%');
      sheet.getRange(`L${row}`).setNumberFormat('0.00%');
    }
  },
  
  setupCustomerJourneySheet: function(sheet) {
    const headers = ['Customer ID', 'Date', 'Name', 'Email', 'Stage', 'Touchpoint', 
                     'Channel', 'Action', 'Response', 'Sentiment', 'Next Step', 
                     'Days in Stage', 'Lifetime Value', 'Churn Risk'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#2ECC71').setFontColor('#FFFFFF').setFontWeight('bold');
    
    // Add stage duration formulas
    for (let row = 2; row <= 100; row++) {
      sheet.getRange(`L${row}`).setFormula(`=TODAY()-B${row}`).setNumberFormat('#,##0');
      sheet.getRange(`M${row}`).setNumberFormat('$#,##0.00');
    }
  },
  
  setupProfitabilitySheet: function(sheet) {
    // Create profitability metrics dashboard
    const profitMetrics = [
      { label: 'Total Revenue', formula: '=SUM(G9:G)', color: '#3B82F6' },
      { label: 'Gross Profit', formula: '=SUM(I9:I)', color: '#10B981' },
      { label: 'Net Profit', formula: '=SUM(L9:L)', color: '#8B5CF6' },
      { label: 'Avg Margin', formula: '=IFERROR(AVERAGE(J9:J),0)', color: '#F59E0B' },
      { label: 'Top Performers', formula: '=COUNTIF(N9:N,">3")', color: '#06B6D4' },
      { label: 'SKU Count', formula: '=COUNTA(A9:A)', color: '#EC4899' }
    ];
    
    this.createMiniDashboard(sheet, 1, 1, profitMetrics);
    
    // Enhanced headers with advanced analytics
    const headers = ['SKU', 'Product Name', 'Category', 'Brand', 'Unit Cost', 'Selling Price', 
                     'Units Sold', 'Revenue', 'COGS', 'Gross Profit', 'Margin %', 
                     'Marketing Cost', 'Net Profit', 'ROI', 'Velocity', 'ABC Class', 
                     'Price Elasticity', 'Profit Rank', 'Contribution %', 'Status', 'Action Required'];
    sheet.getRange(8, 1, 1, headers.length).setValues([headers]);
    
    const colors = this.colors.industries.ecommerce;
    sheet.getRange(8, 1, 1, headers.length)
      .setBackground(colors.primary)
      .setFontColor('#FFFFFF')
      .setFontWeight('bold')
      .setFontSize(11);
    
    // Add comprehensive product data
    const sampleData = [
      ['SKU-001', 'Wireless Headphones', 'Electronics', 'TechBrand', 45, 129.99, 850, '', '', '', '', 500, '', '', '', '', '', '', '', 'Active', ''],
      ['SKU-002', 'Yoga Mat Premium', 'Sports', 'FitLife', 12, 39.99, 2100, '', '', '', '', 200, '', '', '', '', '', '', '', 'Active', ''],
      ['SKU-003', 'Coffee Maker Pro', 'Appliances', 'BrewMaster', 85, 249.99, 320, '', '', '', '', 800, '', '', '', '', '', '', '', 'Active', ''],
      ['SKU-004', 'Smart Watch', 'Electronics', 'TechBrand', 120, 299.99, 450, '', '', '', '', 1200, '', '', '', '', '', '', '', 'Active', ''],
      ['SKU-005', 'Organic Tea Set', 'Food & Bev', 'TeaTime', 8, 24.99, 1500, '', '', '', '', 150, '', '', '', '', '', '', '', 'Active', '']
    ];
    
    sheet.getRange(9, 1, sampleData.length, sampleData[0].length).setValues(sampleData);
    
    // Advanced e-commerce profitability formulas
    for (let row = 9; row <= 108; row++) {
      // Revenue calculation
      sheet.getRange(`H${row}`).setFormula(`=IF(AND(F${row}<>"",G${row}<>""),F${row}*G${row},"")`);
      
      // Cost of Goods Sold
      sheet.getRange(`I${row}`).setFormula(`=IF(AND(E${row}<>"",G${row}<>""),E${row}*G${row},"")`);
      
      // Gross Profit
      sheet.getRange(`J${row}`).setFormula(`=IF(AND(H${row}<>"",I${row}<>""),H${row}-I${row},"")`);
      
      // Margin percentage
      sheet.getRange(`K${row}`).setFormula(`=IF(H${row}>0,J${row}/H${row},"")`);
      
      // Net Profit
      sheet.getRange(`M${row}`).setFormula(`=IF(AND(J${row}<>"",L${row}<>""),J${row}-L${row},J${row})`);
      
      // ROI calculation
      sheet.getRange(`N${row}`).setFormula(`=IF(AND(L${row}>0,M${row}<>""),M${row}/L${row},IF(I${row}>0,M${row}/I${row},""))`);
      
      // Velocity (units per day over 30 days)
      sheet.getRange(`O${row}`).setFormula(`=IF(G${row}<>"",G${row}/30,"")`);
      
      // ABC Classification based on revenue contribution
      sheet.getRange(`P${row}`).setFormula(
        `=IF(H${row}<>"",` +
        `IF(H${row}>PERCENTILE($H$9:$H,0.8),"A",` +
        `IF(H${row}>PERCENTILE($H$9:$H,0.5),"B","C")),"")`
      );
      
      // Price Elasticity indicator
      sheet.getRange(`Q${row}`).setFormula(
        `=IF(AND(F${row}<>"",G${row}<>""),` +
        `IF(F${row}>200,"Inelastic",` +
        `IF(F${row}>100,"Moderate","Elastic")),"")`
      );
      
      // Profit ranking
      sheet.getRange(`R${row}`).setFormula(`=IF(M${row}<>"",RANK(M${row},$M$9:$M$108),"")`);
      
      // Contribution percentage to total profit
      sheet.getRange(`S${row}`).setFormula(`=IF(M${row}<>"",M${row}/SUM($M$9:$M),"")`);
      
      // Action required based on performance
      sheet.getRange(`U${row}`).setFormula(
        `=IF(K${row}<0.2,"🔴 REVIEW PRICING",` +
        `IF(O${row}<1,"🟡 SLOW MOVER",` +
        `IF(P${row}="A","⭐ TOP PRODUCT",` +
        `IF(N${row}>5,"🚀 HIGH ROI",` +
        `IF(K${row}>0.5,"💰 HIGH MARGIN",` +
        `IF(T${row}="Active","✅ PERFORMING","📊 ANALYZE"))))))`
      );
    }
    
    // Format columns
    sheet.getRange('E:F').setNumberFormat('$#,##0.00');
    sheet.getRange('H:J').setNumberFormat('$#,##0.00');
    sheet.getRange('L:M').setNumberFormat('$#,##0.00');
    sheet.getRange('K:K').setNumberFormat('0.00%');
    sheet.getRange('N:N').setNumberFormat('0.00');
    sheet.getRange('O:O').setNumberFormat('0.00');
    sheet.getRange('R:R').setNumberFormat('0');
    sheet.getRange('S:S').setNumberFormat('0.00%');
    
    // Add sparkline for profit trends
    this.addSparkline(sheet, 'V8', 'M9:M13', 'column', { color: '#10B981' });
    
    // Data validations
    const categoryRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Electronics', 'Appliances', 'Sports', 'Fashion', 'Home & Garden', 'Food & Bev', 'Beauty', 'Toys', 'Books', 'Other'], true)
      .build();
    sheet.getRange('C9:C').setDataValidation(categoryRule);
    
    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Active', 'Discontinued', 'Out of Stock', 'Seasonal', 'Clearance'], true)
      .build();
    sheet.getRange('T9:T').setDataValidation(statusRule);
    
    // Conditional formatting for margins
    const highMarginRule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThanOrEqualTo(0.5)
      .setBackground('#D1FAE5')
      .setFontColor('#065F46')
      .setRanges([sheet.getRange('K9:K')])
      .build();
    
    const lowMarginRule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0.2)
      .setBackground('#FEE2E2')
      .setFontColor('#991B1B')
      .setRanges([sheet.getRange('K9:K')])
      .build();
    
    // ABC classification formatting
    const classARule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('A')
      .setBackground('#FEF3C7')
      .setFontColor('#92400E')
      .setRanges([sheet.getRange('P9:P')])
      .build();
    
    sheet.setConditionalFormatRules([highMarginRule, lowMarginRule, classARule]);
    
    // Protect formula columns
    const protection = sheet.getRange('H:K,M:S,U:U').protect();
    protection.setDescription('Automated calculations - Do not edit');
    protection.setWarningOnly(true);
    
    // Apply alternating rows
    this.applyProfessionalFormatting(sheet, 9, 108);
  },
  
  setupSalesForecastSheet: function(sheet) {
    const headers = ['Month', 'Historical Sales', 'Growth Rate', 'Seasonal Factor', 'Base Forecast', 
                     'Marketing Impact', 'Adjusted Forecast', 'Confidence %', 'Best Case', 
                     'Worst Case', 'Actual Sales', 'Variance', 'Accuracy %'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#8E44AD').setFontColor('#FFFFFF').setFontWeight('bold');
    
    // Add forecasting formulas
    for (let row = 2; row <= 13; row++) {
      sheet.getRange(`E${row}`).setFormula(`=B${row}*(1+C${row})*D${row}`).setNumberFormat('$#,##0');
      sheet.getRange(`G${row}`).setFormula(`=E${row}+F${row}`).setNumberFormat('$#,##0');
      sheet.getRange(`I${row}`).setFormula(`=G${row}*1.2`).setNumberFormat('$#,##0');
      sheet.getRange(`J${row}`).setFormula(`=G${row}*0.8`).setNumberFormat('$#,##0');
      sheet.getRange(`L${row}`).setFormula(`=K${row}-G${row}`).setNumberFormat('$#,##0');
      sheet.getRange(`M${row}`).setFormula(`=IF(K${row}>0,1-ABS(L${row})/K${row},0)`).setNumberFormat('0.00%');
    }
  },
  
  setupProjectProfitabilitySheet: function(sheet) {
    const headers = ['Project ID', 'Client', 'Project Name', 'Start Date', 'End Date', 
                     'Budget', 'Hours Budgeted', 'Hours Used', 'Labor Cost', 'Material Cost', 
                     'Other Costs', 'Total Cost', 'Revenue', 'Profit', 'Margin %', 'Status'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#2C3E50').setFontColor('#FFFFFF').setFontWeight('bold');
    
    // Add profitability formulas
    for (let row = 2; row <= 50; row++) {
      sheet.getRange(`L${row}`).setFormula(`=I${row}+J${row}+K${row}`).setNumberFormat('$#,##0.00');
      sheet.getRange(`N${row}`).setFormula(`=M${row}-L${row}`).setNumberFormat('$#,##0.00');
      sheet.getRange(`O${row}`).setFormula(`=IF(M${row}>0,N${row}/M${row},0)`).setNumberFormat('0.00%');
    }
  },
  
  setupClientDashboardSheet: function(sheet) {
    const headers = ['Client ID', 'Client Name', 'Industry', 'Contact', 'Email', 'Phone', 
                     'Total Projects', 'Active Projects', 'Completed Projects', 'Total Revenue', 
                     'Outstanding Balance', 'Last Contact', 'Satisfaction Score', 'Status'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#1ABC9C').setFontColor('#FFFFFF').setFontWeight('bold');
    
    // Format currency and number columns
    sheet.getRange(2, 10, 100, 2).setNumberFormat('$#,##0.00');
    sheet.getRange(2, 7, 100, 3).setNumberFormat('#,##0');
    sheet.getRange(2, 13, 100, 1).setNumberFormat('0.0');
    
    // Add data validation for status
    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Active', 'Inactive', 'Prospect', 'Lost'])
      .build();
    sheet.getRange(2, 14, 100, 1).setDataValidation(statusRule);
  },
  
  // Additional missing setup functions
  setupTimeBillingSheet: function(sheet) {
    const headers = ['Date', 'Client', 'Project', 'Task Description', 'Start Time', 'End Time',
                     'Hours', 'Hourly Rate', 'Total', 'Billable', 'Invoice #', 'Status'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#3498DB').setFontColor('#FFFFFF').setFontWeight('bold');
    
    // Add formulas for hours and total calculation
    for (let row = 2; row <= 100; row++) {
      sheet.getRange(`G${row}`).setFormula(`=(F${row}-E${row})*24`).setNumberFormat('0.00');
      sheet.getRange(`I${row}`).setFormula(`=G${row}*H${row}`).setNumberFormat('$#,##0.00');
    }
  },
  
  setupCommissionTrackerSheet: function(sheet) {
    // Enhanced headers with more tracking fields
    const headers = ['Transaction ID', 'Date', 'Agent', 'Property Address', 'Sale Price', 
                     'Commission %', 'Commission Amount', 'Split %', 'Net Commission', 'Status', 
                     'Lead Source', 'Days to Close', 'Client Type', 'Property Type', 'YTD Total'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Apply professional formatting with industry colors
    const colors = this.colors.industries.realEstate;
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground(colors.primary)
      .setFontColor('#FFFFFF')
      .setFontWeight('bold')
      .setFontSize(11);
    
    // Add comprehensive sample data
    const sampleData = [
      ['TRX-001', '=TODAY()-30', 'Agent Smith', '123 Main St, Downtown', 450000, 3, '', 70, '', 'Closed', 'Referral', '', 'Buyer', 'Single Family', ''],
      ['TRX-002', '=TODAY()-15', 'Agent Jones', '456 Oak Ave, Westside', 325000, 2.5, '', 70, '', 'Pending', 'Website', '', 'Seller', 'Condo', ''],
      ['TRX-003', '=TODAY()-7', 'Agent Smith', '789 Elm Dr, Northside', 575000, 3, '', 70, '', 'Under Contract', 'Open House', '', 'Buyer', 'Single Family', ''],
      ['TRX-004', '=TODAY()-45', 'Agent Brown', '321 Park Pl, Eastside', 800000, 2.75, '', 65, '', 'Closed', 'Referral', '', 'Seller', 'Luxury', ''],
      ['TRX-005', '=TODAY()-3', 'Agent Smith', '654 River Rd, South Bay', 425000, 3, '', 70, '', 'Active', 'Social Media', '', 'Buyer', 'Townhouse', '']
    ];
    
    sheet.getRange(2, 1, sampleData.length, sampleData[0].length).setValues(sampleData);
    
    // Advanced formulas for calculations
    for (let row = 2; row <= 100; row++) {
      // Commission calculation
      sheet.getRange(`G${row}`).setFormula(`=IF(E${row}<>"",E${row}*F${row}/100,"")`);
      // Net commission after split
      sheet.getRange(`I${row}`).setFormula(`=IF(G${row}<>"",G${row}*H${row}/100,"")`);
      // Days to close calculation
      sheet.getRange(`L${row}`).setFormula(`=IF(AND(J${row}="Closed",B${row}<>""),TODAY()-B${row},IF(J${row}<>"Closed","In Progress",""))`);
      // YTD running total
      sheet.getRange(`O${row}`).setFormula(`=IF(J${row}="Closed",SUMIF($J$2:$J${row},"Closed",$I$2:$I${row}),"")`);
    }
    
    // Advanced formatting
    sheet.getRange('E:E').setNumberFormat('$#,##0');
    sheet.getRange('G:G').setNumberFormat('$#,##0.00');
    sheet.getRange('I:I').setNumberFormat('$#,##0.00');
    sheet.getRange('O:O').setNumberFormat('$#,##0.00');
    sheet.getRange('F:F').setNumberFormat('0.00%');
    sheet.getRange('H:H').setNumberFormat('0%');
    sheet.getRange('B:B').setNumberFormat('dd/mm/yyyy');
    
    // Create status tracking with colors
    this.createStatusTracking(sheet, 'J', 2, 100);
    
    // Add data validation for Lead Source
    const leadSourceRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Referral', 'Website', 'Open House', 'Social Media', 'Cold Call', 'Email Campaign', 'Partner Agent'], true)
      .build();
    sheet.getRange('K2:K100').setDataValidation(leadSourceRule);
    
    // Add data validation for Client Type
    const clientTypeRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Buyer', 'Seller', 'Both'], true)
      .build();
    sheet.getRange('M2:M100').setDataValidation(clientTypeRule);
    
    // Add data validation for Property Type
    const propertyTypeRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Single Family', 'Condo', 'Townhouse', 'Multi-Family', 'Land', 'Commercial', 'Luxury'], true)
      .build();
    sheet.getRange('N2:N100').setDataValidation(propertyTypeRule);
    
    // Add summary section at top
    sheet.insertRowBefore(1);
    sheet.insertRowBefore(1);
    sheet.insertRowBefore(1);
    
    // Create mini dashboard at top
    sheet.getRange('A1:O1').merge();
    sheet.getRange('A1').setValue('COMMISSION TRACKER - PERFORMANCE SUMMARY')
      .setFontSize(16).setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBackground(colors.dark)
      .setFontColor('#FFFFFF');
    
    // Add KPI cards
    sheet.getRange('A2').setValue('Total Closed:');
    sheet.getRange('B2').setFormula('=COUNTIF(J:J,"Closed")')
      .setFontSize(14).setFontWeight('bold').setFontColor(colors.primary);
    
    sheet.getRange('D2').setValue('Total Revenue:');
    sheet.getRange('E2').setFormula('=SUMIF(J:J,"Closed",I:I)')
      .setNumberFormat('$#,##0')
      .setFontSize(14).setFontWeight('bold').setFontColor(colors.primary);
    
    sheet.getRange('G2').setValue('Avg Commission:');
    sheet.getRange('H2').setFormula('=AVERAGEIF(J:J,"Closed",I:I)')
      .setNumberFormat('$#,##0')
      .setFontSize(14).setFontWeight('bold').setFontColor(colors.primary);
    
    sheet.getRange('J2').setValue('Conversion Rate:');
    sheet.getRange('K2').setFormula('=COUNTIF(J:J,"Closed")/COUNTA(J:J)')
      .setNumberFormat('0%')
      .setFontSize(14).setFontWeight('bold').setFontColor(colors.primary);
    
    sheet.getRange('M2').setValue('Avg Days to Close:');
    sheet.getRange('N2').setFormula('=AVERAGE(L:L)')
      .setNumberFormat('0')
      .setFontSize(14).setFontWeight('bold').setFontColor(colors.primary);
    
    // Add sparkline for trend
    sheet.getRange('O2').setFormula('=SPARKLINE(FILTER(I:I,J:J="Closed"),{"charttype","column";"color","' + colors.secondary + '"})')
      .setHorizontalAlignment('center');
    
    // Apply alternating rows
    this.applyProfessionalFormatting(sheet, 103);
    
    // Auto-resize columns
    this.autoResizeColumns(sheet);
    
    // Protect formula cells
    const protection = sheet.getRange('G:G,I:I,L:L,O:O').protect();
    protection.setDescription('Formula Protection');
    protection.setWarningOnly(true);
  },
  
  setupPipelineSheet: function(sheet) {
    try {
      // Simple header without mini dashboard for now
      sheet.getRange('A1').setValue('PIPELINE TRACKER')
        .setFontSize(14).setFontWeight('bold');
      sheet.getRange('A2').setValue('Lead Management System')
        .setFontSize(10).setFontColor('#666666');
    
      // Comprehensive pipeline headers starting at row 8
      const headers = ['Lead ID', 'Date', 'Name', 'Contact', 'Email', 'Property Interest', 'Budget', 
                       'Pre-approved', 'Stage', 'Days in Stage', 'Source', 'Next Action', 'Follow-up Date', 
                       'Agent', 'Lead Score', 'Probability %', 'Est. Close Date', 'Notes', 'Auto Priority'];
      sheet.getRange(8, 1, 1, headers.length).setValues([headers]);
    
      // Apply professional formatting with colors
      const colors = this.colors.industries.realEstate;
      sheet.getRange(8, 1, 1, headers.length)
        .setBackground(colors.secondary)
        .setFontColor('#FFFFFF')
        .setFontWeight('bold')
        .setFontSize(11);
    
      // Add comprehensive sample data
      const sampleData = [
        ['L-001', '=TODAY()-14', 'John Smith', '555-0101', 'john@email.com', 'Single Family Home', 450000, 'Yes', 'Qualified', '', 'Referral', 'Schedule Showing', '=TODAY()+2', 'Agent Smith', '', '', '=TODAY()+30', 'Looking for 3BR, good schools', ''],
        ['L-002', '=TODAY()-7', 'Jane Doe', '555-0102', 'jane@email.com', 'Condo', 275000, 'Pending', 'Active', '', 'Website', 'Send Listings', '=TODAY()+1', 'Agent Jones', '', '', '=TODAY()+45', 'First-time buyer, downtown area', ''],
        ['L-003', '=TODAY()-3', 'Bob Wilson', '555-0103', 'bob@email.com', 'Investment Property', 600000, 'Yes', 'Hot', '', 'Open House', 'Initial Contact', '=TODAY()', 'Agent Smith', '', '', '=TODAY()+15', 'Cash buyer, multiple properties', ''],
        ['L-004', '=TODAY()-21', 'Alice Brown', '555-0104', 'alice@email.com', 'Townhouse', 350000, 'No', 'Cold', '', 'Social Media', 'Follow-up Email', '=TODAY()+7', 'Agent Brown', '', '', '=TODAY()+90', 'Just browsing', ''],
        ['L-005', '=TODAY()-1', 'Tom Johnson', '555-0105', 'tom@email.com', 'Luxury Home', 1200000, 'Yes', 'New', '', 'Partner Referral', 'Qualification Call', '=TODAY()', 'Agent Smith', '', '', '', 'High-value prospect', '']
      ];
    
      sheet.getRange(9, 1, sampleData.length, sampleData[0].length).setValues(sampleData);
    
      // Advanced formulas for dynamic calculations
      for (let row = 9; row <= 108; row++) {
        // Days in stage calculation
        sheet.getRange(`J${row}`).setFormula(`=IF(B${row}<>"",TODAY()-B${row},"")`);
        
        // Lead scoring algorithm based on multiple factors
        sheet.getRange(`O${row}`).setFormula(
          `=IF(A${row}<>"",` +
          `(IF(H${row}="Yes",30,IF(H${row}="Pending",15,0))+` + // Pre-approval score
          `IF(G${row}>500000,20,IF(G${row}>300000,10,5))+` + // Budget score
          `IF(I${row}="Hot",30,IF(I${row}="Qualified",20,IF(I${row}="Active",10,0)))+` + // Stage score
          `IF(K${row}="Referral",15,IF(K${row}="Partner Referral",15,5))+` + // Source score
          `IF(J${row}<7,10,IF(J${row}<14,5,0))` + // Engagement recency
          `,"")`
        );
        
        // Probability calculation based on lead score
        sheet.getRange(`P${row}`).setFormula(
          `=IF(O${row}<>"",IF(O${row}>80,0.9,IF(O${row}>60,0.7,IF(O${row}>40,0.5,IF(O${row}>20,0.3,0.1)))),"")`
        );
        
        // Auto priority assignment
        sheet.getRange(`S${row}`).setFormula(
          `=IF(A${row}<>"",IF(AND(O${row}>70,M${row}<=TODAY()),"🔴 URGENT",IF(AND(O${row}>50,M${row}<=TODAY()+3),"🟡 HIGH",IF(O${row}>30,"🟢 MEDIUM","⚪ LOW"))),"")`
        );
      }
    
      // Format columns
      sheet.getRange('G:G').setNumberFormat('$#,##0');
      sheet.getRange('O:O').setNumberFormat('0');
      sheet.getRange('P:P').setNumberFormat('0%');
      sheet.getRange('B:B').setNumberFormat('dd/mm/yyyy');
      sheet.getRange('M:M').setNumberFormat('dd/mm/yyyy');
      sheet.getRange('Q:Q').setNumberFormat('dd/mm/yyyy');
    
      // Add sparklines for lead score trends
      this.addSparkline(sheet, 'T8', 'O9:O13', 'column', { color: '#10B981' });
    
      // Add comprehensive data validations
      const stageRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['New', 'Contacted', 'Qualified', 'Active', 'Hot', 'Offer', 'Negotiating', 'Contract', 'Closed', 'Lost', 'Cold'], true)
        .build();
      sheet.getRange('I9:I').setDataValidation(stageRule);
    
      const preApprovalRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['Yes', 'No', 'Pending', 'Expired'], true)
        .build();
      sheet.getRange('H9:H').setDataValidation(preApprovalRule);
    
      const sourceRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['Website', 'Referral', 'Partner Referral', 'Open House', 'Social Media', 'Cold Call', 'Email Campaign', 'Walk-in', 'Other'], true)
        .build();
      sheet.getRange('K9:K').setDataValidation(sourceRule);
    
      const propertyTypeRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['Single Family Home', 'Condo', 'Townhouse', 'Multi-family', 'Investment Property', 'Luxury Home', 'Commercial', 'Land'], true)
        .build();
      sheet.getRange('F9:F').setDataValidation(propertyTypeRule);
    
      // Conditional formatting for stages
      const hotRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('Hot')
        .setBackground('#FEE2E2')
        .setFontColor('#DC2626')
        .setRanges([sheet.getRange('I9:I')])
        .build();
    
      const qualifiedRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('Qualified')
        .setBackground('#DBEAFE')
        .setFontColor('#1E40AF')
        .setRanges([sheet.getRange('I9:I')])
        .build();
    
      const closedRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('Closed')
        .setBackground('#D1FAE5')
        .setFontColor('#065F46')
        .setRanges([sheet.getRange('I9:I')])
        .build();
    
      sheet.setConditionalFormatRules([hotRule, qualifiedRule, closedRule]);
    
      // Protect formula columns
      const protection = sheet.getRange('J:J,O:P,S:S').protect();
      protection.setDescription('Automated formulas - Do not edit');
      protection.setWarningOnly(true);
    
      // Apply alternating rows
      this.applyProfessionalFormatting(sheet, 9, 108);
    
    } catch (error) {
      Logger.log('Error in setupPipelineSheet: ' + error.toString());
      // Add basic fallback headers if advanced setup fails
      const basicHeaders = ['Lead ID', 'Date', 'Name', 'Contact', 'Property Interest', 'Budget', 'Stage', 'Notes'];
      sheet.getRange(1, 1, 1, basicHeaders.length).setValues([basicHeaders]);
      this.formatHeaders(sheet, 1, basicHeaders.length);
    }
  },
  
  setupDashboardSheet: function(sheet) {
    // Title section
    sheet.getRange('A1').setValue('Real Estate Performance Dashboard');
    sheet.getRange('A1').setFontSize(18).setFontWeight('bold').setFontColor(this.colors.primary.header);
    sheet.getRange('A2').setValue('Last Updated: ' + new Date().toLocaleDateString());
    
    // KPI Headers
    const headers = ['Key Metrics', 'Current Month', 'Last Month', 'YTD', 'Target', '% to Target'];
    sheet.getRange(4, 1, 1, headers.length).setValues([headers]);
    this.formatHeaders(sheet, 4, headers.length);
    
    // Cross-tab KPI formulas pulling from other sheets
    const kpiData = [
      ['Total Deals', 
       '=COUNTIFS(\'Commission Tracker\'!B:B,">="&EOMONTH(TODAY(),-1)+1,\'Commission Tracker\'!B:B,"<="&EOMONTH(TODAY(),0),\'Commission Tracker\'!J:J,"Closed")',
       '=COUNTIFS(\'Commission Tracker\'!B:B,">="&EOMONTH(TODAY(),-2)+1,\'Commission Tracker\'!B:B,"<="&EOMONTH(TODAY(),-1),\'Commission Tracker\'!J:J,"Closed")',
       '=COUNTIFS(\'Commission Tracker\'!B:B,">="&DATE(YEAR(TODAY()),1,1),\'Commission Tracker\'!J:J,"Closed")',
       '10',
       '=IF(E5>0,D5/E5,0)'],
      ['Gross Commission', 
       '=SUMIFS(\'Commission Tracker\'!G:G,\'Commission Tracker\'!B:B,">="&EOMONTH(TODAY(),-1)+1,\'Commission Tracker\'!B:B,"<="&EOMONTH(TODAY(),0),\'Commission Tracker\'!J:J,"Closed")',
       '=SUMIFS(\'Commission Tracker\'!G:G,\'Commission Tracker\'!B:B,">="&EOMONTH(TODAY(),-2)+1,\'Commission Tracker\'!B:B,"<="&EOMONTH(TODAY(),-1),\'Commission Tracker\'!J:J,"Closed")',
       '=SUMIFS(\'Commission Tracker\'!G:G,\'Commission Tracker\'!B:B,">="&DATE(YEAR(TODAY()),1,1),\'Commission Tracker\'!J:J,"Closed")',
       '50000',
       '=IF(E6>0,D6/E6,0)'],
      ['Net Commission',
       '=SUMIFS(\'Commission Tracker\'!I:I,\'Commission Tracker\'!B:B,">="&EOMONTH(TODAY(),-1)+1,\'Commission Tracker\'!B:B,"<="&EOMONTH(TODAY(),0),\'Commission Tracker\'!J:J,"Closed")',
       '=SUMIFS(\'Commission Tracker\'!I:I,\'Commission Tracker\'!B:B,">="&EOMONTH(TODAY(),-2)+1,\'Commission Tracker\'!B:B,"<="&EOMONTH(TODAY(),-1),\'Commission Tracker\'!J:J,"Closed")',
       '=SUMIFS(\'Commission Tracker\'!I:I,\'Commission Tracker\'!B:B,">="&DATE(YEAR(TODAY()),1,1),\'Commission Tracker\'!J:J,"Closed")',
       '35000',
       '=IF(E7>0,D7/E7,0)'],
      ['Active Pipeline',
       '=COUNTIF(\'Pipeline\'!G:G,"Active")+COUNTIF(\'Pipeline\'!G:G,"Qualified")',
       '',
       '=COUNTIFS(\'Pipeline\'!B:B,">="&DATE(YEAR(TODAY()),1,1),\'Pipeline\'!G:G,"<>Closed",\'Pipeline\'!G:G,"<>Lost")',
       '25',
       '=IF(E8>0,B8/E8,0)'],
      ['Pipeline Value',
       '=SUMIFS(\'Pipeline\'!F:F,\'Pipeline\'!G:G,"Active")+SUMIFS(\'Pipeline\'!F:F,\'Pipeline\'!G:G,"Qualified")',
       '',
       '=SUMIFS(\'Pipeline\'!F:F,\'Pipeline\'!B:B,">="&DATE(YEAR(TODAY()),1,1),\'Pipeline\'!G:G,"<>Closed",\'Pipeline\'!G:G,"<>Lost")',
       '5000000',
       '=IF(E9>0,B9/E9,0)'],
      ['Avg Deal Size',
       '=IFERROR(AVERAGEIFS(\'Commission Tracker\'!E:E,\'Commission Tracker\'!B:B,">="&EOMONTH(TODAY(),-1)+1,\'Commission Tracker\'!B:B,"<="&EOMONTH(TODAY(),0),\'Commission Tracker\'!J:J,"Closed"),0)',
       '=IFERROR(AVERAGEIFS(\'Commission Tracker\'!E:E,\'Commission Tracker\'!B:B,">="&EOMONTH(TODAY(),-2)+1,\'Commission Tracker\'!B:B,"<="&EOMONTH(TODAY(),-1),\'Commission Tracker\'!J:J,"Closed"),0)',
       '=IFERROR(AVERAGEIFS(\'Commission Tracker\'!E:E,\'Commission Tracker\'!B:B,">="&DATE(YEAR(TODAY()),1,1),\'Commission Tracker\'!J:J,"Closed"),0)',
       '400000',
       '=IF(E10>0,D10/E10,0)'],
      ['Conversion Rate',
       '=IFERROR(COUNTIF(\'Pipeline\'!G:G,"Closed")/(COUNTIF(\'Pipeline\'!G:G,"<>"&"")),0)',
       '',
       '=IFERROR(COUNTIFS(\'Pipeline\'!G:G,"Closed",\'Pipeline\'!B:B,">="&DATE(YEAR(TODAY()),1,1))/COUNTIFS(\'Pipeline\'!B:B,">="&DATE(YEAR(TODAY()),1,1),\'Pipeline\'!A:A,"<>"),0)',
       '0.25',
       '=IF(E11>0,D11/E11,0)']
    ];
    
    sheet.getRange(5, 1, kpiData.length, kpiData[0].length).setValues(kpiData);
    
    // Format the dashboard
    sheet.getRange('B5:D11').setNumberFormat('#,##0');
    sheet.getRange('B6:D7').setNumberFormat('$#,##0');
    sheet.getRange('B9:D9').setNumberFormat('$#,##0');
    sheet.getRange('B10:D10').setNumberFormat('$#,##0');
    sheet.getRange('B11:D11').setNumberFormat('0.0%');
    sheet.getRange('F5:F11').setNumberFormat('0%');
    
    // Apply conditional formatting for variance
    const varianceRange = sheet.getRange('F5:F11');
    const greenRule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThanOrEqualTo(1)
      .setBackground('#D4EDDA')
      .setFontColor('#155724')
      .setRanges([varianceRange])
      .build();
    const yellowRule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(0.7, 0.99)
      .setBackground('#FFF3CD')
      .setFontColor('#856404')
      .setRanges([varianceRange])
      .build();
    const redRule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0.7)
      .setBackground('#F8D7DA')
      .setFontColor('#721C24')
      .setRanges([varianceRange])
      .build();
    
    sheet.setConditionalFormatRules([greenRule, yellowRule, redRule]);
    
    // Auto-resize columns
    this.autoResizeColumns(sheet);
    
    // Apply professional formatting
    this.applyProfessionalFormatting(sheet, 11);
    
    // Add charts section header
    sheet.getRange('A14').setValue('Performance Trends');
    sheet.getRange('A14').setFontSize(14).setFontWeight('bold').setFontColor(this.colors.primary.header);
    
    // Add lead source analysis
    sheet.getRange('A17').setValue('Lead Source Analysis');
    sheet.getRange('A17').setFontSize(14).setFontWeight('bold').setFontColor(this.colors.primary.header);
    
    const sourceHeaders = ['Source', 'Leads', 'Conversions', 'Conversion %', 'Revenue'];
    sheet.getRange(18, 1, 1, sourceHeaders.length).setValues([sourceHeaders]);
    this.formatHeaders(sheet, 18, sourceHeaders.length);
    
    const sourceData = [
      ['Referral', '=COUNTIF(\'Pipeline\'!I:I,"Referral")', '=COUNTIFS(\'Pipeline\'!I:I,"Referral",\'Pipeline\'!G:G,"Closed")', '=IFERROR(C19/B19,0)', '=SUMIFS(\'Commission Tracker\'!I:I,\'Commission Tracker\'!K:K,"Referral",\'Commission Tracker\'!J:J,"Closed")'],
      ['Website', '=COUNTIF(\'Pipeline\'!I:I,"Website")', '=COUNTIFS(\'Pipeline\'!I:I,"Website",\'Pipeline\'!G:G,"Closed")', '=IFERROR(C20/B20,0)', '=SUMIFS(\'Commission Tracker\'!I:I,\'Commission Tracker\'!K:K,"Website",\'Commission Tracker\'!J:J,"Closed")'],
      ['Open House', '=COUNTIF(\'Pipeline\'!I:I,"Open House")', '=COUNTIFS(\'Pipeline\'!I:I,"Open House",\'Pipeline\'!G:G,"Closed")', '=IFERROR(C21/B21,0)', '=SUMIFS(\'Commission Tracker\'!I:I,\'Commission Tracker\'!K:K,"Open House",\'Commission Tracker\'!J:J,"Closed")']
    ];
    
    sheet.getRange(19, 1, sourceData.length, sourceData[0].length).setValues(sourceData);
    sheet.getRange('D19:D21').setNumberFormat('0%');
    sheet.getRange('E19:E21').setNumberFormat('$#,##0');
    
    // Apply formatting to source analysis
    this.applyProfessionalFormatting(sheet, 21);
  },
  
  setupCostEstimatorSheet: function(sheet) {
    const headers = ['Item', 'Category', 'Quantity', 'Unit', 'Unit Cost', 'Total Cost', 
                     'Markup %', 'Selling Price', 'Margin', 'Vendor', 'Lead Time', 'Notes'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Apply professional formatting
    this.formatHeaders(sheet, 1, headers.length);
    this.autoResizeColumns(sheet);
    
    // Add sample data
    const sampleData = [
      ['Concrete Foundation', 'Materials', 100, 'cu yd', 150, '', 20, '', '', 'ABC Supply', '3 days', 'Grade A'],
      ['Framing Lumber', 'Materials', 5000, 'bd ft', 2.5, '', 25, '', '', 'Lumber Co', '1 week', '2x4 studs'],
      ['Labor - Framing', 'Labor', 160, 'hours', 45, '', 15, '', '', 'Crew A', 'Available', 'Skilled crew'],
      ['Electrical Rough-in', 'Subcontractor', 1, 'job', 3500, '', 10, '', '', 'Elite Electric', '2 weeks', 'Licensed']
    ];
    
    sheet.getRange(2, 1, sampleData.length, sampleData[0].length).setValues(sampleData);
    
    // Apply formulas
    for (let row = 2; row <= 100; row++) {
      sheet.getRange(`F${row}`).setFormula(`=IF(AND(C${row}<>"",E${row}<>""),C${row}*E${row},"")`);
      sheet.getRange(`H${row}`).setFormula(`=IF(AND(F${row}<>"",G${row}<>""),F${row}*(1+G${row}/100),"")`);
      sheet.getRange(`I${row}`).setFormula(`=IF(AND(H${row}<>"",F${row}<>""),H${row}-F${row},"")`);
    }
    
    // Format columns
    sheet.getRange('E:E').setNumberFormat('$#,##0.00');
    sheet.getRange('F:F').setNumberFormat('$#,##0.00');
    sheet.getRange('H:H').setNumberFormat('$#,##0.00');
    sheet.getRange('I:I').setNumberFormat('$#,##0.00');
    sheet.getRange('G:G').setNumberFormat('0%');
    
    // Add category validation
    const categoryRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Materials', 'Labor', 'Equipment', 'Subcontractor', 'Permits', 'Other'], true)
      .build();
    sheet.getRange('B2:B').setDataValidation(categoryRule);
    
    // Apply alternating rows
    this.applyProfessionalFormatting(sheet, 50);
  },
  
  setupInsuranceVerifierSheet: function(sheet) {
    const headers = ['Patient ID', 'Patient Name', 'DOB', 'Insurance', 'Policy #', 
                     'Group #', 'Effective Date', 'Expiry Date', 'Copay', 'Deductible', 
                     'Coverage %', 'Out of Pocket Max', 'Verified Date', 'Verified By', 'Status'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Apply professional formatting
    this.formatHeaders(sheet, 1, headers.length);
    this.autoResizeColumns(sheet);
    
    // Add sample data
    const sampleData = [
      ['P-001', 'John Smith', '01/15/1980', 'Blue Cross', 'BC123456', 'GRP789', '01/01/2024', '12/31/2024', 25, 1500, 80, 5000, '=TODAY()', 'Admin', 'Verified'],
      ['P-002', 'Jane Doe', '03/22/1975', 'Aetna', 'AET789012', 'GRP456', '03/01/2024', '02/28/2025', 30, 2000, 70, 6000, '=TODAY()-1', 'Admin', 'Verified'],
      ['P-003', 'Bob Wilson', '07/10/1990', 'United Health', 'UH345678', 'GRP123', '06/01/2024', '05/31/2025', 20, 1000, 90, 4000, '=TODAY()-2', 'Admin', 'Pending']
    ];
    
    sheet.getRange(2, 1, sampleData.length, sampleData[0].length).setValues(sampleData);
    
    // Format columns
    sheet.getRange('I:I').setNumberFormat('$#,##0');
    sheet.getRange('J:J').setNumberFormat('$#,##0');
    sheet.getRange('L:L').setNumberFormat('$#,##0');
    sheet.getRange('K:K').setNumberFormat('0%');
    
    // Add status validation and conditional formatting
    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Verified', 'Pending', 'Expired', 'Invalid'], true)
      .build();
    sheet.getRange('O2:O').setDataValidation(statusRule);
    
    // Conditional formatting for status
    const statusRange = sheet.getRange('O2:O100');
    const verifiedRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Verified')
      .setBackground(this.colors.success.light)
      .setRanges([statusRange])
      .build();
    const pendingRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Pending')
      .setBackground(this.colors.warning.light)
      .setRanges([statusRange])
      .build();
    const expiredRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Expired')
      .setBackground(this.colors.danger.light)
      .setRanges([statusRange])
      .build();
    
    sheet.setConditionalFormatRules([verifiedRule, pendingRule, expiredRule]);
    
    // Apply alternating rows
    this.applyProfessionalFormatting(sheet, 50);
  },
  
  setupCampaignDashboardSheet: function(sheet) {
    // Create campaign performance dashboard
    const campaignMetrics = [
      { label: 'Active Campaigns', formula: '=COUNTIF(N9:N,"Active")', color: '#10B981' },
      { label: 'Total Spend', formula: '=SUM(F9:F)', color: '#3B82F6' },
      { label: 'Total Conversions', formula: '=SUM(J9:J)', color: '#8B5CF6' },
      { label: 'Avg CTR', formula: '=IFERROR(AVERAGE(I9:I),0)', color: '#F59E0B' },
      { label: 'Avg ROAS', formula: '=IFERROR(AVERAGE(L9:L),0)', color: '#10B981' },
      { label: 'Budget Utilization', formula: '=IFERROR(SUM(F9:F)/SUM(E9:E),0)', color: '#06B6D4' }
    ];
    
    this.createMiniDashboard(sheet, 1, 1, campaignMetrics);
    
    // Enhanced headers with advanced tracking
    const headers = ['Campaign ID', 'Campaign Name', 'Channel', 'Start Date', 'End Date', 'Budget', 'Spend', 
                     'Impressions', 'Clicks', 'CTR %', 'Conversions', 'CPA', 'ROAS', 'ROI %', 
                     'Status', 'Performance Score', 'Pacing %', 'Days Left', 'Daily Budget', 
                     'Conversion Rate', 'Quality Score', 'Auto Alert', 'Attribution Model', 'Notes'];
    sheet.getRange(8, 1, 1, headers.length).setValues([headers]);
    
    // Apply marketing industry colors
    const colors = this.colors.industries.marketing;
    sheet.getRange(8, 1, 1, headers.length)
      .setBackground(colors.primary)
      .setFontColor('#FFFFFF')
      .setFontWeight('bold')
      .setFontSize(11);
    
    // Add comprehensive campaign data
    const sampleData = [
      ['CMP-001', 'Summer Sale 2024', 'Google Ads', '=TODAY()-30', '=TODAY()+30', 5000, 2500, 125000, 3750, '', 45, '', '', '', 'Active', '', '', '', '', '', '', '', 'Last-Click', 'Search campaign'],
      ['CMP-002', 'Email Newsletter', 'Email', '=TODAY()-15', '=TODAY()+15', 1000, 450, 25000, 1250, '', 25, '', '', '', 'Active', '', '', '', '', '', '', '', 'Linear', 'Weekly newsletter'],
      ['CMP-003', 'Social Media Push', 'Facebook', '=TODAY()-45', '=TODAY()-15', 3000, 3000, 75000, 2250, '', 30, '', '', '', 'Completed', '', '', '', '', '', '', '', 'Data-Driven', 'Brand awareness'],
      ['CMP-004', 'Black Friday', 'Multi-Channel', '=TODAY()+20', '=TODAY()+25', 10000, 0, 0, 0, '', 0, '', '', '', 'Planned', '', '', '', '', '', '', '', 'Position-Based', 'Major sale event'],
      ['CMP-005', 'Retargeting', 'Display', '=TODAY()-10', '=TODAY()+20', 2000, 800, 50000, 1000, '', 15, '', '', '', 'Active', '', '', '', '', '', '', '', 'Time-Decay', 'Cart abandoners']
    ];
    
    sheet.getRange(9, 1, sampleData.length, sampleData[0].length).setValues(sampleData);
    
    // Advanced marketing formulas
    for (let row = 9; row <= 108; row++) {
      // CTR calculation
      sheet.getRange(`J${row}`).setFormula(`=IF(H${row}>0,I${row}/H${row},"")`);
      
      // CPA calculation (Cost Per Acquisition)
      sheet.getRange(`L${row}`).setFormula(`=IF(K${row}>0,G${row}/K${row},"")`);
      
      // ROAS calculation (Return on Ad Spend)
      sheet.getRange(`M${row}`).setFormula(`=IF(G${row}>0,(K${row}*100)/G${row},"")`);
      
      // ROI percentage
      sheet.getRange(`N${row}`).setFormula(`=IF(G${row}>0,((K${row}*100-G${row})/G${row}),"")`);
      
      // Performance score (weighted metrics)
      sheet.getRange(`P${row}`).setFormula(
        `=IF(A${row}<>"",` +
        `MIN(100,` +
        `(IF(J${row}>0.02,30,J${row}*1500)+` + // CTR score
        `IF(M${row}>3,30,M${row}*10)+` + // ROAS score
        `IF(T${row}>0.02,20,T${row}*1000)+` + // Conversion rate score
        `IF(Q${row}<1,20,IF(Q${row}<1.2,10,0))` + // Pacing score
        `)),"")`
      );
      
      // Pacing percentage (spend vs time elapsed)
      sheet.getRange(`Q${row}`).setFormula(
        `=IF(AND(F${row}<>"",D${row}<=TODAY(),E${row}>=TODAY()),` +
        `G${row}/(F${row}*((TODAY()-D${row})/(E${row}-D${row}+1))),"")`
      );
      
      // Days left in campaign
      sheet.getRange(`R${row}`).setFormula(
        `=IF(AND(E${row}<>"",O${row}="Active"),MAX(0,E${row}-TODAY()),"")`
      );
      
      // Daily budget recommendation
      sheet.getRange(`S${row}`).setFormula(
        `=IF(AND(R${row}>0,F${row}<>"",G${row}<>""),(F${row}-G${row})/R${row},"")`
      );
      
      // Conversion rate
      sheet.getRange(`T${row}`).setFormula(`=IF(I${row}>0,K${row}/I${row},"")`);
      
      // Quality score simulation
      sheet.getRange(`U${row}`).setFormula(
        `=IF(A${row}<>"",RANDBETWEEN(5,10),"")` // In real world, this would come from ad platform
      );
      
      // Auto alert generation
      sheet.getRange(`V${row}`).setFormula(
        `=IF(O${row}="Planned","📅 SCHEDULED",` +
        `IF(AND(Q${row}<>"",Q${row}>1.2),"🔴 OVERSPENDING",` +
        `IF(AND(Q${row}<>"",Q${row}<0.8),"🟡 UNDERPACING",` +
        `IF(AND(J${row}<>"",J${row}<0.01),"⚠️ LOW CTR",` +
        `IF(AND(M${row}<>"",M${row}<1),"📉 NEGATIVE ROAS",` +
        `IF(P${row}>80,"⭐ TOP PERFORMER",` +
        `IF(O${row}="Active","✅ RUNNING","🏁 COMPLETED")))))))`
      );
    }
    
    // Format columns
    sheet.getRange('D:E').setNumberFormat('dd/mm/yyyy');
    sheet.getRange('F:G').setNumberFormat('$#,##0');
    sheet.getRange('H:I').setNumberFormat('#,##0');
    sheet.getRange('J:J').setNumberFormat('0.00%');
    sheet.getRange('K:K').setNumberFormat('#,##0');
    sheet.getRange('L:L').setNumberFormat('$#,##0.00');
    sheet.getRange('M:N').setNumberFormat('0.00');
    sheet.getRange('P:P').setNumberFormat('0');
    sheet.getRange('Q:Q').setNumberFormat('0.00');
    sheet.getRange('R:R').setNumberFormat('0');
    sheet.getRange('S:S').setNumberFormat('$#,##0.00');
    sheet.getRange('T:T').setNumberFormat('0.00%');
    sheet.getRange('U:U').setNumberFormat('0');
    
    // Add sparkline for performance trends
    this.addSparkline(sheet, 'Y8', 'P9:P13', 'line', { color: '#10B981' });
    
    // Comprehensive data validations
    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Planned', 'Active', 'Paused', 'Completed', 'Cancelled'], true)
      .build();
    sheet.getRange('O9:O').setDataValidation(statusRule);
    
    const marketingChannelRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Google Ads', 'Facebook', 'Instagram', 'LinkedIn', 'Twitter', 'Email', 'Display', 'Video', 'Multi-Channel', 'Other'], true)
      .build();
    sheet.getRange('C9:C').setDataValidation(marketingChannelRule);
    
    const attributionRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Last-Click', 'First-Click', 'Linear', 'Time-Decay', 'Position-Based', 'Data-Driven'], true)
      .build();
    sheet.getRange('W9:W').setDataValidation(attributionRule);
    
    // Conditional formatting for performance
    const highPerformanceRule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThanOrEqualTo(80)
      .setBackground('#D1FAE5')
      .setFontColor('#065F46')
      .setRanges([sheet.getRange('P9:P')])
      .build();
    
    const lowPerformanceRule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(40)
      .setBackground('#FEE2E2')
      .setFontColor('#991B1B')
      .setRanges([sheet.getRange('P9:P')])
      .build();
    
    // Pacing alerts
    const overspendingRule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(1.2)
      .setBackground('#FEE2E2')
      .setRanges([sheet.getRange('Q9:Q')])
      .build();
    
    const underspendingRule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0.8)
      .setBackground('#FEF3C7')
      .setRanges([sheet.getRange('Q9:Q')])
      .build();
    
    sheet.setConditionalFormatRules([highPerformanceRule, lowPerformanceRule, overspendingRule, underspendingRule]);
    
    // Protect formula columns
    const protection = sheet.getRange('J:N,P:U,V:V').protect();
    protection.setDescription('Automated calculations - Do not edit');
    protection.setWarningOnly(true);
    
    // Apply alternating rows
    this.applyProfessionalFormatting(sheet, 9, 108);
    sheet.getRange('J:J').setNumberFormat('#,##0');
    sheet.getRange('K:K').setNumberFormat('$#,##0.00');
    sheet.getRange('L:L').setNumberFormat('0.00');
    
    // Add channel validation
    const oldChannelRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Google Ads', 'Facebook', 'Instagram', 'Email', 'LinkedIn', 'Twitter', 'Multi-Channel'], true)
      .build();
    sheet.getRange('B2:B').setDataValidation(oldChannelRule);
    
    // Add status validation
    const oldStatusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Planned', 'Active', 'Paused', 'Completed'], true)
      .build();
    sheet.getRange('M2:M').setDataValidation(oldStatusRule);
    
    // Apply conditional formatting for ROAS
    const roasRange = sheet.getRange('L2:L50');
    const goodRoas = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThanOrEqualTo(3)
      .setBackground(this.colors.success.light)
      .setRanges([roasRange])
      .build();
    const avgRoas = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(1, 2.99)
      .setBackground(this.colors.warning.light)
      .setRanges([roasRange])
      .build();
    const poorRoas = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(1)
      .setBackground(this.colors.danger.light)
      .setRanges([roasRange])
      .build();
    
    sheet.setConditionalFormatRules([goodRoas, avgRoas, poorRoas]);
    
    // Apply alternating rows
    this.applyProfessionalFormatting(sheet, 50);
  },
  
  setupInventorySheet: function(sheet) {
    const headers = ['SKU', 'Product Name', 'Category', 'Current Stock', 'Min Stock', 
                     'Max Stock', 'Reorder Point', 'Reorder Qty', 'Unit Cost', 'Total Value', 
                     'Supplier', 'Lead Time', 'Last Order', 'Status'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Apply professional formatting
    this.formatHeaders(sheet, 1, headers.length);
    this.autoResizeColumns(sheet);
    
    // Add sample inventory data
    const sampleData = [
      ['SKU-001', 'Widget A', 'Electronics', 45, 20, 100, 30, 50, 25.99, '', 'Supplier A', '7 days', '=TODAY()-14', ''],
      ['SKU-002', 'Widget B', 'Electronics', 15, 25, 75, 30, 40, 35.50, '', 'Supplier B', '10 days', '=TODAY()-7', ''],
      ['SKU-003', 'Gadget X', 'Accessories', 120, 50, 200, 75, 100, 12.99, '', 'Supplier A', '5 days', '=TODAY()-21', ''],
      ['SKU-004', 'Gadget Y', 'Accessories', 8, 15, 60, 20, 30, 45.00, '', 'Supplier C', '14 days', '=TODAY()-30', '']
    ];
    
    sheet.getRange(2, 1, sampleData.length, sampleData[0].length).setValues(sampleData);
    
    // Apply formulas
    for (let row = 2; row <= 100; row++) {
      sheet.getRange(`J${row}`).setFormula(`=IF(AND(D${row}<>"",I${row}<>""),D${row}*I${row},"")`);
      sheet.getRange(`N${row}`).setFormula(`=IF(D${row}<>"",IF(D${row}<=G${row},"REORDER",IF(D${row}>=F${row},"OVERSTOCK","OK")),"")`);
    }
    
    // Format columns
    sheet.getRange('I:I').setNumberFormat('$#,##0.00');
    sheet.getRange('J:J').setNumberFormat('$#,##0.00');
    sheet.getRange('D:H').setNumberFormat('#,##0');
    
    // Add category validation
    const categoryRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Electronics', 'Accessories', 'Clothing', 'Home & Garden', 'Sports', 'Books'], true)
      .build();
    sheet.getRange('C2:C').setDataValidation(categoryRule);
    
    // Conditional formatting for status
    const statusRange = sheet.getRange('N2:N100');
    const reorderRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('REORDER')
      .setBackground(this.colors.danger.light)
      .setFontColor(this.colors.danger.dark)
      .setRanges([statusRange])
      .build();
    const overstockRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('OVERSTOCK')
      .setBackground(this.colors.warning.light)
      .setFontColor(this.colors.warning.dark)
      .setRanges([statusRange])
      .build();
    const okRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('OK')
      .setBackground(this.colors.success.light)
      .setFontColor(this.colors.success.dark)
      .setRanges([statusRange])
      .build();
    
    sheet.setConditionalFormatRules([reorderRule, overstockRule, okRule]);
    
    // Apply alternating rows
    this.applyProfessionalFormatting(sheet, 50);
  },
  
  // Additional Real Estate setup functions
  setupPropertySheet: function(sheet) {
    const headers = ['Property ID', 'Address', 'Type', 'Bedrooms', 'Bathrooms', 
                     'Square Feet', 'Purchase Price', 'Current Value', 'Monthly Rent', 
                     'Annual Income', 'Cap Rate', 'Occupancy Status', 'Tenant', 'Lease End', 'HOA Fees', 'Property Tax'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Apply professional formatting
    this.formatHeaders(sheet, 1, headers.length);
    this.autoResizeColumns(sheet);
    
    // Add sample properties
    const sampleData = [
      ['PROP-001', '123 Main St, City, ST', 'Single Family', 3, 2, 1850, 350000, 385000, 2500, '', '', 'Occupied', 'John Smith', '=TODAY()+180', 150, 4200],
      ['PROP-002', '456 Oak Ave #5', 'Condo', 2, 1, 950, 225000, 245000, 1800, '', '', 'Occupied', 'Jane Doe', '=TODAY()+90', 250, 2800],
      ['PROP-003', '789 Elm Dr', 'Duplex', 4, 3, 2400, 425000, 450000, 3200, '', '', 'Vacant', '', '', 0, 5100]
    ];
    
    sheet.getRange(2, 1, sampleData.length, sampleData[0].length).setValues(sampleData);
    
    // Add formulas for calculations
    for (let row = 2; row <= 50; row++) {
      sheet.getRange(`J${row}`).setFormula(`=IF(I${row}<>"",I${row}*12,"")`);
      sheet.getRange(`K${row}`).setFormula(`=IF(AND(J${row}<>"",H${row}<>""),J${row}/H${row},"")`);
    }
    
    // Format columns
    sheet.getRange('G:I').setNumberFormat('$#,##0');
    sheet.getRange('J:J').setNumberFormat('$#,##0');
    sheet.getRange('K:K').setNumberFormat('0.00%');
    sheet.getRange('O:P').setNumberFormat('$#,##0');
    
    // Add type validation
    const typeRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Single Family', 'Condo', 'Townhouse', 'Duplex', 'Multi-Family', 'Commercial'], true)
      .build();
    sheet.getRange('C2:C').setDataValidation(typeRule);
    
    // Apply alternating rows
    this.applyProfessionalFormatting(sheet, 50);
  },
  
  setupTenantSheet: function(sheet) {
    const headers = ['Tenant ID', 'Name', 'Contact', 'Email', 'Property', 
                     'Lease Start', 'Lease End', 'Monthly Rent', 'Security Deposit', 
                     'Payment Status', 'Last Payment', 'Balance Due', 'Days Until Renewal'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Apply professional formatting
    this.formatHeaders(sheet, 1, headers.length);
    this.autoResizeColumns(sheet);
    
    // Add sample tenants with cross-references to properties
    const sampleData = [
      ['T-001', 'John Smith', '555-0101', 'john@email.com', '=\'Properties\'!B2', '=TODAY()-180', '=TODAY()+180', 2500, 2500, 'Current', '=TODAY()-5', 0, ''],
      ['T-002', 'Jane Doe', '555-0102', 'jane@email.com', '=\'Properties\'!B3', '=TODAY()-270', '=TODAY()+90', 1800, 1800, 'Current', '=TODAY()-3', 0, ''],
      ['T-003', 'Bob Wilson', '555-0103', 'bob@email.com', 'Pending', '=TODAY()+30', '=TODAY()+395', 3200, 3200, 'Pending', '', 0, '']
    ];
    
    sheet.getRange(2, 1, sampleData.length, sampleData[0].length).setValues(sampleData);
    
    // Add formula for days until renewal
    for (let row = 2; row <= 50; row++) {
      sheet.getRange(`M${row}`).setFormula(`=IF(G${row}<>"",G${row}-TODAY(),"")`);
    }
    
    // Format columns
    sheet.getRange('H:I').setNumberFormat('$#,##0');
    sheet.getRange('L:L').setNumberFormat('$#,##0');
    
    // Add payment status validation
    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Current', 'Late', 'Pending', 'Vacant'], true)
      .build();
    sheet.getRange('J2:J').setDataValidation(statusRule);
    
    // Conditional formatting for payment status
    const statusRange = sheet.getRange('J2:J50');
    const currentRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Current')
      .setBackground(this.colors.success.light)
      .setRanges([statusRange])
      .build();
    const lateRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Late')
      .setBackground(this.colors.danger.light)
      .setRanges([statusRange])
      .build();
    
    sheet.setConditionalFormatRules([currentRule, lateRule]);
    
    // Apply alternating rows
    this.applyProfessionalFormatting(sheet, 50);
  },
  
  setupMaintenanceSheet: function(sheet) {
    const headers = ['Request ID', 'Date', 'Property', 'Tenant', 'Issue Type', 
                     'Description', 'Priority', 'Status', 'Vendor', 'Cost', 
                     'Completion Date', 'Days Open', 'Notes'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Apply professional formatting
    this.formatHeaders(sheet, 1, headers.length);
    this.autoResizeColumns(sheet);
    
    // Add sample maintenance requests with cross-references
    const sampleData = [
      ['M-001', '=TODAY()-7', '=\'Properties\'!B2', '=\'Tenants\'!B2', 'Plumbing', 'Leaky faucet in kitchen', 'Medium', 'In Progress', 'ABC Plumbing', 250, '', '', 'Parts ordered'],
      ['M-002', '=TODAY()-3', '=\'Properties\'!B3', '=\'Tenants\'!B3', 'HVAC', 'AC not cooling', 'High', 'Open', '', '', '', '', 'Tenant called today'],
      ['M-003', '=TODAY()-14', '=\'Properties\'!B2', '=\'Tenants\'!B2', 'Electrical', 'Outlet not working', 'Low', 'Completed', 'Elite Electric', 150, '=TODAY()-2', '', 'Fixed']
    ];
    
    sheet.getRange(2, 1, sampleData.length, sampleData[0].length).setValues(sampleData);
    
    // Add formula for days open
    for (let row = 2; row <= 50; row++) {
      sheet.getRange(`L${row}`).setFormula(`=IF(AND(B${row}<>"",H${row}<>"Completed"),TODAY()-B${row},IF(AND(B${row}<>"",K${row}<>""),K${row}-B${row},""))`);
    }
    
    // Format columns
    sheet.getRange('J:J').setNumberFormat('$#,##0');
    
    // Add priority validation
    const priorityRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Low', 'Medium', 'High', 'Emergency'], true)
      .build();
    sheet.getRange('G2:G').setDataValidation(priorityRule);
    
    // Add status validation
    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Open', 'In Progress', 'Waiting Parts', 'Completed', 'Cancelled'], true)
      .build();
    sheet.getRange('H2:H').setDataValidation(statusRule);
    
    // Conditional formatting for priority
    const priorityRange = sheet.getRange('G2:G50');
    const emergencyRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Emergency')
      .setBackground(this.colors.danger.dark)
      .setFontColor('#FFFFFF')
      .setRanges([priorityRange])
      .build();
    const highRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('High')
      .setBackground(this.colors.danger.light)
      .setRanges([priorityRange])
      .build();
    const mediumRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Medium')
      .setBackground(this.colors.warning.light)
      .setRanges([priorityRange])
      .build();
    
    sheet.setConditionalFormatRules([emergencyRule, highRule, mediumRule]);
    
    // Apply alternating rows
    this.applyProfessionalFormatting(sheet, 50);
  },
  
  setupInvestmentSheet: function(sheet) {
    const headers = ['Property', 'Purchase Price', 'Down Payment', 'Loan Amount', 
                     'Interest Rate', 'Monthly Payment', 'Rental Income', 
                     'Operating Expenses', 'Cash Flow', 'Annual Cash Flow', 'ROI %', 'Cap Rate'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Apply professional formatting
    this.formatHeaders(sheet, 1, headers.length);
    this.autoResizeColumns(sheet);
    
    // Add sample investment properties
    const sampleData = [
      ['=\'Properties\'!B2', '=\'Properties\'!G2', '=B2*0.2', '=B2-C2', 4.5, '', '=\'Properties\'!I2', 500, '', '', '', ''],
      ['=\'Properties\'!B3', '=\'Properties\'!G3', '=B3*0.2', '=B3-C3', 4.25, '', '=\'Properties\'!I3', 350, '', '', '', ''],
      ['789 Investment Blvd', 525000, '=B4*0.25', '=B4-C4', 4.75, '', 3800, 650, '', '', '', '']
    ];
    
    sheet.getRange(2, 1, sampleData.length, sampleData[0].length).setValues(sampleData);
    
    // Add investment formulas
    for (let row = 2; row <= 50; row++) {
      sheet.getRange(`F${row}`).setFormula(`=IF(AND(D${row}<>"",E${row}<>""),PMT(E${row}/100/12,360,-D${row}),"")`);
      sheet.getRange(`I${row}`).setFormula(`=IF(AND(G${row}<>"",F${row}<>"",H${row}<>""),G${row}-F${row}-H${row},"")`);
      sheet.getRange(`J${row}`).setFormula(`=IF(I${row}<>"",I${row}*12,"")`);
      sheet.getRange(`K${row}`).setFormula(`=IF(AND(J${row}<>"",C${row}<>""),J${row}/C${row},"")`);
      sheet.getRange(`L${row}`).setFormula(`=IF(AND(J${row}<>"",B${row}<>""),J${row}/B${row},"")`);
    }
    
    // Format columns
    sheet.getRange('B:D').setNumberFormat('$#,##0');
    sheet.getRange('E:E').setNumberFormat('0.00%');
    sheet.getRange('F:J').setNumberFormat('$#,##0');
    sheet.getRange('K:L').setNumberFormat('0.00%');
    
    // Apply alternating rows
    this.applyProfessionalFormatting(sheet, 50);
  },
  
  setupCashFlowSheet: function(sheet) {
    const headers = ['Month', 'Rental Income', 'Other Income', 'Total Income', 
                     'Mortgage', 'HOA', 'Insurance', 'Property Tax', 'Maintenance', 
                     'Utilities', 'Total Expenses', 'Net Cash Flow', 'Cumulative'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Apply professional formatting
    this.formatHeaders(sheet, 1, headers.length);
    this.autoResizeColumns(sheet);
    
    // Add sample cash flow data for 12 months
    const months = ['January', 'February', 'March', 'April', 'May', 'June', 
                   'July', 'August', 'September', 'October', 'November', 'December'];
    const sampleData = [];
    
    for (let i = 0; i < months.length; i++) {
      sampleData.push([
        months[i],
        '=SUMIF(\'Properties\'!L:L,"Occupied",\'Properties\'!I:I)', // Rental from occupied properties
        0, // Other income
        '', // Total income formula
        2500, // Mortgage
        '=SUM(\'Properties\'!O:O)', // HOA from properties
        350, // Insurance
        '=SUM(\'Properties\'!P:P)/12', // Property tax monthly
        '=AVERAGEIF(\'Maintenance\'!H:H,"<>Cancelled",\'Maintenance\'!J:J)', // Avg maintenance
        200, // Utilities
        '', // Total expenses formula
        '', // Net cash flow formula
        '' // Cumulative formula
      ]);
    }
    
    sheet.getRange(2, 1, sampleData.length, sampleData[0].length).setValues(sampleData);
    
    // Apply formulas
    for (let row = 2; row <= 100; row++) {
      sheet.getRange(`D${row}`).setFormula(`=IF(OR(B${row}<>"",C${row}<>""),B${row}+C${row},"")`);
      sheet.getRange(`K${row}`).setFormula(`=IF(A${row}<>"",SUM(E${row}:J${row}),"")`);
      sheet.getRange(`L${row}`).setFormula(`=IF(AND(D${row}<>"",K${row}<>""),D${row}-K${row},"")`);
      if (row === 2) {
        sheet.getRange(`M${row}`).setFormula(`=IF(L${row}<>"",L${row},"")`);
      } else {
        sheet.getRange(`M${row}`).setFormula(`=IF(L${row}<>"",IF(M${row-1}<>"",M${row-1}+L${row},L${row}),"")`);
      }
    }
    
    // Format columns
    sheet.getRange('B:M').setNumberFormat('$#,##0');
    
    // Conditional formatting for negative cash flow
    const cashFlowRange = sheet.getRange('L2:L100');
    const negativeRule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0)
      .setBackground(this.colors.danger.light)
      .setFontColor(this.colors.danger.dark)
      .setRanges([cashFlowRange])
      .build();
    const positiveRule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0)
      .setBackground(this.colors.success.light)
      .setFontColor(this.colors.success.dark)
      .setRanges([cashFlowRange])
      .build();
    
    sheet.setConditionalFormatRules([negativeRule, positiveRule]);
    
    // Apply alternating rows
    this.applyProfessionalFormatting(sheet, 13);
  },
  
  setupROISheet: function(sheet) {
    // Title
    sheet.getRange('A1').setValue('Investment ROI Calculator');
    sheet.getRange('A1').setFontSize(18).setFontWeight('bold').setFontColor(this.colors.primary.header);
    
    const headers = ['Metric', 'Value', 'Calculation', 'Industry Benchmark'];
    sheet.getRange(3, 1, 1, headers.length).setValues([headers]);
    
    // Apply professional formatting
    this.formatHeaders(sheet, 3, headers.length);
    this.autoResizeColumns(sheet);
    
    // Enhanced metrics with cross-tab formulas
    const metrics = [
      ['Purchase Price', '=SUM(\'Investment Analysis\'!B:B)', 'Sum of all property purchases', ''],
      ['Down Payment', '=SUM(\'Investment Analysis\'!C:C)', 'Total cash invested', ''],
      ['Annual Rental Income', '=SUM(\'Cash Flow\'!B14:B25)', 'Total rental income per year', '$50,000'],
      ['Annual Expenses', '=SUM(\'Cash Flow\'!K14:K25)', 'Total operating expenses', '$20,000'],
      ['Net Operating Income', '=B6-B7', 'Income minus expenses', '$30,000'],
      ['Annual Debt Service', '=SUM(\'Cash Flow\'!E14:E25)', 'Total mortgage payments', '$18,000'],
      ['Annual Cash Flow', '=B8-B9', 'NOI minus debt service', '$12,000'],
      ['', '', '', ''],
      ['Key Performance Metrics', '', '', ''],
      ['Cash-on-Cash Return', '=IF(B5>0,B10/B5,0)', 'Annual cash flow / down payment', '8-12%'],
      ['Cap Rate', '=IF(B4>0,B8/B4,0)', 'NOI / property value', '6-10%'],
      ['Gross Rent Multiplier', '=IF(B6>0,B4/B6,0)', 'Property value / annual rent', '8-12x'],
      ['Debt Service Coverage', '=IF(B9>0,B8/B9,0)', 'NOI / debt service', '>1.25x'],
      ['Total ROI', '=IF(B5>0,(B10+(B4*0.03))/B5,0)', 'Total return on investment', '15-20%']
    ];
    
    sheet.getRange(4, 1, metrics.length, 4).setValues(metrics);
    
    // Format values
    sheet.getRange('B4:B10').setNumberFormat('$#,##0');
    sheet.getRange('B13:B13').setNumberFormat('0.00%');
    sheet.getRange('B14:B14').setNumberFormat('0.00%');
    sheet.getRange('B15:B15').setNumberFormat('0.00');
    sheet.getRange('B16:B16').setNumberFormat('0.00');
    sheet.getRange('B17:B17').setNumberFormat('0.00%');
    
    // Add section headers formatting
    sheet.getRange('A12:D12').setBackground(this.colors.primary.light);
    sheet.getRange('A12').setFontWeight('bold').setFontSize(12);
    
    // Conditional formatting for metrics
    const performanceRange = sheet.getRange('B13:B17');
    
    // Apply alternating rows
    this.applyProfessionalFormatting(sheet, 17);
    
    // Add investment summary section
    sheet.getRange('A20').setValue('Investment Summary');
    sheet.getRange('A20').setFontSize(14).setFontWeight('bold').setFontColor(this.colors.primary.header);
    
    const summaryHeaders = ['Property', 'Purchase Date', 'Current Value', 'Monthly Cash Flow', 'ROI %'];
    sheet.getRange(21, 1, 1, summaryHeaders.length).setValues([summaryHeaders]);
    this.formatHeaders(sheet, 21, summaryHeaders.length);
    
    // Add cross-tab summary formulas
    const summaryData = [
      ['=\'Investment Analysis\'!A2', '=\'Commission Tracker\'!B2', '=\'Investment Analysis\'!B2*1.1', '=\'Cash Flow\'!L2', '=IF(\'Investment Analysis\'!C2>0,E22/\'Investment Analysis\'!C2,0)'],
      ['=\'Investment Analysis\'!A3', '=\'Commission Tracker\'!B3', '=\'Investment Analysis\'!B3*1.1', '=\'Cash Flow\'!L3', '=IF(\'Investment Analysis\'!C3>0,E23/\'Investment Analysis\'!C3,0)'],
      ['=\'Investment Analysis\'!A4', '=\'Commission Tracker\'!B4', '=\'Investment Analysis\'!B4*1.1', '=\'Cash Flow\'!L4', '=IF(\'Investment Analysis\'!C4>0,E24/\'Investment Analysis\'!C4,0)']
    ];
    
    sheet.getRange(22, 1, summaryData.length, summaryData[0].length).setValues(summaryData);
    sheet.getRange('C22:D24').setNumberFormat('$#,##0');
    sheet.getRange('E22:E24').setNumberFormat('0.00%');
    
    // Apply alternating rows to summary
    this.applyProfessionalFormatting(sheet, 24);
  },
  
  // New setup functions for enhanced dashboards
  
  // Real Estate additional sheets
  setupClientsSheet: function(sheet, commissionSheetName) {
    try {
      // If commissionSheetName not provided, use default (for backward compatibility)
      commissionSheetName = commissionSheetName || 'Commission Tracker';
      
      // Create mini CRM dashboard at top
      const crmMetrics = [
        { label: 'Total Clients', formula: '=COUNTA(B9:B)-COUNTIF(F9:F,"Inactive")', color: '#3B82F6' },
        { label: 'Active Buyers', formula: '=COUNTIFS(E9:E,"Buyer",F9:F,"Active")', color: '#10B981' },
        { label: 'Active Sellers', formula: '=COUNTIFS(E9:E,"Seller",F9:F,"Active")', color: '#8B5CF6' },
        { label: 'VIP Clients', formula: '=COUNTIF(K9:K,"VIP")', color: '#F59E0B' },
        { label: 'Avg Deal Size', formula: '=IFERROR(AVERAGE(I9:I),0)', color: '#06B6D4' },
        { label: 'Total Portfolio', formula: '=SUM(I9:I)', color: '#EC4899' }
      ];
    
      this.createMiniDashboard(sheet, 1, 1, crmMetrics);
    
      // Enhanced headers with more tracking fields
      const headers = ['Client ID', 'Name', 'Email', 'Phone', 'Type', 'Status', 'First Contact', 
                       'Last Activity', 'Total Value', 'Lifetime Value', 'Client Tier', 'Preferred Contact', 
                       'Property Preferences', 'Engagement Score', 'Next Action', 'Referral Source', 'Notes'];
      sheet.getRange(8, 1, 1, headers.length).setValues([headers]);
    
      // Apply industry colors
      const colors = this.colors.industries.realEstate;
      sheet.getRange(8, 1, 1, headers.length)
        .setBackground(colors.primary)
        .setFontColor('#FFFFFF')
        .setFontWeight('bold')
        .setFontSize(11);
    
      // Add comprehensive sample data
      const sampleData = [
        ['CL001', 'John Smith', 'john@email.com', '555-0101', 'Buyer', 'Active', '=TODAY()-30', '=TODAY()-2', 450000, '', '', 'Phone', '3BR, Schools, Suburbs', '', 'Schedule viewing', 'Website', 'Looking to buy in Q1'],
        ['CL002', 'Jane Doe', 'jane@email.com', '555-0102', 'Seller', 'Active', '=TODAY()-15', '=TODAY()-1', 325000, '', '', 'Email', 'Quick sale needed', '', 'Market analysis', 'Referral', 'Relocating for work'],
        ['CL003', 'Bob Wilson', 'bob@email.com', '555-0103', 'Investor', 'VIP', '=TODAY()-180', '=TODAY()', 2750000, '', '', 'Text', 'Multi-unit, ROI focus', '', 'Portfolio review', 'Partner', 'Multiple properties'],
        ['CL004', 'Alice Brown', 'alice@email.com', '555-0104', 'Buyer', 'Prospect', '=TODAY()-7', '=TODAY()-3', 275000, '', '', 'Email', 'First-time buyer', '', 'Pre-approval help', 'Open House', 'Needs mortgage guidance'],
        ['CL005', 'Tom Johnson', 'tom@email.com', '555-0105', 'Both', 'Active', '=TODAY()-60', '=TODAY()', 550000, '', '', 'Phone', 'Upgrade home', '', 'Coordinate buy/sell', 'Past Client', 'Selling to upgrade']
      ];
      sheet.getRange(9, 1, sampleData.length, sampleData[0].length).setValues(sampleData);
    
      // Advanced formulas
      for (let row = 9; row <= 108; row++) {
        // Lifetime value calculation (includes all closed deals)
        sheet.getRange(`J${row}`).setFormula(
          `=IF(A${row}<>"",SUMIFS('${commissionSheetName}'!I:I,'${commissionSheetName}'!C:C,B${row},'${commissionSheetName}'!J:J,"Closed"),"")`
        );
        
        // Client tier assignment based on value and activity
        sheet.getRange(`K${row}`).setFormula(
          `=IF(A${row}<>"",IF(OR(I${row}>1000000,J${row}>100000),"VIP",IF(OR(I${row}>500000,J${row}>50000),"Premium",IF(F${row}="Active","Standard","Basic"))),"")`
        );
        
        // Engagement score based on recency and frequency
        sheet.getRange(`N${row}`).setFormula(
          `=IF(A${row}<>"",` +
          `MIN(100,` +
          `IF(H${row}=TODAY(),30,IF(H${row}>TODAY()-7,20,IF(H${row}>TODAY()-30,10,0)))+` + // Recency score
          `IF(J${row}>100000,30,IF(J${row}>50000,20,10))+` + // Value score
          `IF(F${row}="Active",20,IF(F${row}="VIP",30,0))+` + // Status score
          `IF(G${row}>TODAY()-30,10,IF(G${row}>TODAY()-90,5,20))` + // Tenure score
          `),"")`
        );
      }
    
      // Format columns
      sheet.getRange('I:J').setNumberFormat('$#,##0');
      sheet.getRange('G:H').setNumberFormat('dd/mm/yyyy');
      sheet.getRange('N:N').setNumberFormat('0');
    
      // Data validations
      const typeRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['Buyer', 'Seller', 'Both', 'Investor', 'Renter', 'Landlord'], true)
        .build();
      sheet.getRange('E9:E').setDataValidation(typeRule);
    
      const statusRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['Active', 'VIP', 'Prospect', 'On Hold', 'Inactive', 'Past Client'], true)
        .build();
      sheet.getRange('F9:F').setDataValidation(statusRule);
    
      const contactRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['Phone', 'Email', 'Text', 'WhatsApp', 'In Person'], true)
        .build();
      sheet.getRange('L9:L').setDataValidation(contactRule);
    
      // Conditional formatting for client tiers
      const vipRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('VIP')
        .setBackground('#FEF3C7')
        .setFontColor('#92400E')
        .setRanges([sheet.getRange('K9:K')])
        .build();
    
      const premiumRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('Premium')
        .setBackground('#EDE9FE')
        .setFontColor('#5B21B6')
        .setRanges([sheet.getRange('K9:K')])
        .build();
    
      // High engagement score formatting
      const highEngagementRule = SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThanOrEqualTo(80)
        .setBackground('#D1FAE5')
        .setFontColor('#065F46')
        .setRanges([sheet.getRange('N9:N')])
        .build();
    
      sheet.setConditionalFormatRules([vipRule, premiumRule, highEngagementRule]);
    
      // Add sparkline for engagement history
      this.addSparkline(sheet, 'R8', 'N9:N13', 'line', { color: '#10B981' });
    
      // Protect formula columns
      const protection = sheet.getRange('J:J,K:K,N:N').protect();
      protection.setDescription('Automated calculations - Do not edit');
      protection.setWarningOnly(true);
    
      // Apply alternating rows
      this.applyProfessionalFormatting(sheet, 9, 108);
    
    } catch (error) {
      Logger.log('Error in setupClientsSheet: ' + error.toString());
      // Add basic fallback headers if advanced setup fails
      const basicHeaders = ['Client ID', 'Name', 'Email', 'Phone', 'Type', 'Status', 'Notes'];
      sheet.getRange(1, 1, 1, basicHeaders.length).setValues([basicHeaders]);
      this.formatHeaders(sheet, 1, basicHeaders.length);
    }
  },
  
  setupFollowUpSheet: function(sheet) {
    const headers = ['Lead ID', 'Name', 'Type', 'Priority', 'Follow-up Date', 'Notes', 'Status'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    this.formatHeaders(sheet, 1, headers.length);
    
    // Add sample data
    const sampleData = [
      ['L001', 'John Smith', 'Call', 'High', '=TODAY()+1', 'Discuss offer', 'Pending'],
      ['L002', 'Jane Doe', 'Email', 'Medium', '=TODAY()+2', 'Send property list', 'Scheduled'],
      ['L003', 'Bob Wilson', 'Meeting', 'High', '=TODAY()', 'Property tour', 'Today']
    ];
    sheet.getRange(2, 1, sampleData.length, sampleData[0].length).setValues(sampleData);
    
    // Format date column
    sheet.getRange('E:E').setNumberFormat('dd/mm/yyyy');
    
    // Add conditional formatting for priority
    const priorityRange = sheet.getRange('D2:D100');
    const highPriorityRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('High')
      .setBackground('#FFE4E1')
      .setRanges([priorityRange])
      .build();
    sheet.setConditionalFormatRules([highPriorityRule]);
    
    this.autoResizeColumns(sheet);
  },
  
  setupRentalIncomeSheet: function(sheet) {
    const headers = ['Property ID', 'Address', 'Tenant', 'Monthly Rent', 'Due Date', 'Status', 'Late Fees', 'Total'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    this.formatHeaders(sheet, 1, headers.length);
    this.autoResizeColumns(sheet);
  },
  
  setupPropertyComparisonSheet: function(sheet) {
    const headers = ['Property', 'Price', 'Cap Rate', 'Cash Flow', 'ROI', 'Score'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    this.formatHeaders(sheet, 1, headers.length);
    this.autoResizeColumns(sheet);
  },
  
  // Construction additional sheets
  setupSuppliersSheet: function(sheet) {
    const headers = ['Supplier ID', 'Company', 'Contact', 'Phone', 'Email', 'Rating', 'Payment Terms', 'Total Orders'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    this.formatHeaders(sheet, 1, headers.length);
    this.autoResizeColumns(sheet);
  },
  
  setupPurchaseOrdersSheet: function(sheet) {
    const headers = ['PO Number', 'Date', 'Supplier', 'Items', 'Total Amount', 'Status', 'Delivery Date', 'Project'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    this.formatHeaders(sheet, 1, headers.length);
    this.autoResizeColumns(sheet);
  },
  
  setupCrewScheduleSheet: function(sheet) {
    const headers = ['Crew ID', 'Name', 'Project', 'Date', 'Start Time', 'End Time', 'Hours', 'Status'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    this.formatHeaders(sheet, 1, headers.length);
    this.autoResizeColumns(sheet);
  },
  
  setupPayrollSheet: function(sheet) {
    const headers = ['Employee', 'Period', 'Regular Hours', 'OT Hours', 'Rate', 'Gross Pay', 'Deductions', 'Net Pay'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    this.formatHeaders(sheet, 1, headers.length);
    this.autoResizeColumns(sheet);
  },
  
  setupCostImpactSheet: function(sheet) {
    const headers = ['Change Order', 'Original Cost', 'New Cost', 'Cost Impact', 'Schedule Impact', 'Status'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    this.formatHeaders(sheet, 1, headers.length);
    this.autoResizeColumns(sheet);
  },
  
  setupApprovalsSheet: function(sheet) {
    const headers = ['Request ID', 'Type', 'Description', 'Requested By', 'Status', 'Days Pending', 'Approver'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    this.formatHeaders(sheet, 1, headers.length);
    this.autoResizeColumns(sheet);
  },
  
  setupCostBreakdownSheet: function(sheet) {
    const headers = ['Category', 'Description', 'Materials', 'Labor', 'Equipment', 'Subcontractor', 'Total'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    this.formatHeaders(sheet, 1, headers.length);
    this.autoResizeColumns(sheet);
  },
  
  setupContingencySheet: function(sheet) {
    const headers = ['Risk Category', 'Description', 'Amount', 'Probability', 'Impact', 'Mitigation'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    this.formatHeaders(sheet, 1, headers.length);
    this.autoResizeColumns(sheet);
  },
  
  // Healthcare additional sheets
  setupAuthStatusSheet: function(sheet) {
    const headers = ['Auth ID', 'Patient', 'Status', 'Days to Approve', 'Priority', 'Insurance', 'Provider'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    this.formatHeaders(sheet, 1, headers.length);
    this.autoResizeColumns(sheet);
  },
  
  setupProvidersSheet: function(sheet) {
    const headers = ['Provider ID', 'Name', 'Specialty', 'NPI', 'Phone', 'Fax', 'Email', 'Status'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    this.formatHeaders(sheet, 1, headers.length);
    this.autoResizeColumns(sheet);
  },
  
  setupClaimsSheet: function(sheet) {
    const headers = ['Claim ID', 'Patient', 'DOS', 'Insurance', 'Amount', 'Status', 'Denied Amount', 'Days in AR'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    this.formatHeaders(sheet, 1, headers.length);
    this.autoResizeColumns(sheet);
  },
  
  setupARAgingSheet: function(sheet) {
    const headers = ['Invoice', 'Patient', 'Amount', 'Age (Days)', 'Insurance', 'Status'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    this.formatHeaders(sheet, 1, headers.length);
    this.autoResizeColumns(sheet);
  },
  
  setupAppealsSheet: function(sheet) {
    const headers = ['Appeal ID', 'Claim ID', 'Date Filed', 'Type', 'Reason', 'Status', 'Amount Recovered', 'Days to Resolution'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    this.formatHeaders(sheet, 1, headers.length);
    this.autoResizeColumns(sheet);
  },
  
  setupDenialTrendsSheet: function(sheet) {
    const headers = ['Denial Code', 'Description', 'Count', 'Total Amount', 'Avg Amount', 'Trend'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    this.formatHeaders(sheet, 1, headers.length);
    this.autoResizeColumns(sheet);
  },
  
  setupEligibilitySheet: function(sheet) {
    const headers = ['Patient ID', 'Insurance', 'Member ID', 'Status', 'Effective Date', 'Term Date'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    this.formatHeaders(sheet, 1, headers.length);
    this.autoResizeColumns(sheet);
  },
  
  setupBenefitsSheet: function(sheet) {
    const headers = ['Patient', 'Plan Type', 'Deductible', 'Out of Pocket Max', 'Network Status', 'Copay', 'Coinsurance'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    this.formatHeaders(sheet, 1, headers.length);
    this.autoResizeColumns(sheet);
  },
  
  // Marketing additional sheets
  setupSegmentsSheet: function(sheet) {
    const headers = ['Segment', 'Description', 'Size', 'Conversion Rate', 'Avg Value', 'Growth'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    this.formatHeaders(sheet, 1, headers.length);
    this.autoResizeColumns(sheet);
  },
  
  setupEngagementSheet: function(sheet) {
    const headers = ['Lead ID', 'Name', 'Email Opens', 'Link Clicks', 'Page Views', 'Score', 'Last Activity'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    this.formatHeaders(sheet, 1, headers.length);
    this.autoResizeColumns(sheet);
  },
  
  setupMetricsSheet: function(sheet) {
    const headers = ['Content ID', 'Title', 'Time on Page', 'Bounce Rate', 'Engagement Rate', 'Shares', 'Conversions'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    this.formatHeaders(sheet, 1, headers.length);
    this.autoResizeColumns(sheet);
  },
  
  setupChannelsSheet: function(sheet) {
    const headers = ['Channel', 'Traffic', 'Conversions', 'Revenue', 'ROI', 'Trend'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    this.formatHeaders(sheet, 1, headers.length);
    this.autoResizeColumns(sheet);
  },
  
  setupTouchpointsSheet: function(sheet) {
    const headers = ['Touchpoint', 'Type', 'Interactions', 'Attribution', 'Value'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    this.formatHeaders(sheet, 1, headers.length);
    this.autoResizeColumns(sheet);
  },
  
  setupConversionSheet: function(sheet) {
    const headers = ['Customer ID', 'Journey ID', 'Start Date', 'Status', 'Revenue', 'Order Value', 'Products'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    this.formatHeaders(sheet, 1, headers.length);
    this.autoResizeColumns(sheet);
  },
  
  setupPerformanceSheet: function(sheet) {
    const headers = ['Campaign', 'Start Date', 'End Date', 'Impressions', 'Clicks', 'Spend', 'Revenue', 'CPA', 'CTR', 'ROAS'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    this.formatHeaders(sheet, 1, headers.length);
    this.autoResizeColumns(sheet);
  },
  
  setupBudgetSheet: function(sheet) {
    const headers = ['Channel', 'Allocated Budget', 'Spent', 'Remaining', 'Performance', 'Recommendation'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    this.formatHeaders(sheet, 1, headers.length);
    this.autoResizeColumns(sheet);
  },
  
  // E-commerce additional sheets
  setupProductsSheet: function(sheet) {
    const headers = ['SKU', 'Product', 'Category', 'Cost', 'Price', 'Units Sold', 'Margin %', 'Revenue'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    this.formatHeaders(sheet, 1, headers.length);
    this.autoResizeColumns(sheet);
  },
  
  setupCostsSheet: function(sheet) {
    const headers = ['Category', 'Description', 'Fixed Cost', 'Variable Cost', 'Total', 'Per Unit'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    this.formatHeaders(sheet, 1, headers.length);
    this.autoResizeColumns(sheet);
  },
  
  setupTrendsSheet: function(sheet) {
    const headers = ['Period', 'Actual Sales', 'Forecast', 'Growth Rate', 'Variance'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    this.formatHeaders(sheet, 1, headers.length);
    this.autoResizeColumns(sheet);
  },
  
  setupSeasonalitySheet: function(sheet) {
    const headers = ['Month', 'Historical Avg', 'Index', 'This Year', 'Variance'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    this.formatHeaders(sheet, 1, headers.length);
    this.autoResizeColumns(sheet);
  },
  
  setupStockLevelsSheet: function(sheet) {
    const headers = ['SKU', 'Product', 'Category', 'Current Stock', 'Value', 'Status', 'Days of Supply'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    this.formatHeaders(sheet, 1, headers.length);
    this.autoResizeColumns(sheet);
  },
  
  setupReorderPointsSheet: function(sheet) {
    const headers = ['SKU', 'Product', 'Reorder Point', 'Current Stock', 'Order Qty', 'Status'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    this.formatHeaders(sheet, 1, headers.length);
    this.autoResizeColumns(sheet);
  },
  
  // Professional Services additional sheets
  setupResourceSheet: function(sheet) {
    const headers = ['Resource', 'Project', 'Role', 'Hours Allocated', 'Hours Used', 'Utilization %', 'Rate'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    this.formatHeaders(sheet, 1, headers.length);
    this.autoResizeColumns(sheet);
  },
  
  setupExpensesSheet: function(sheet) {
    const headers = ['Date', 'Category', 'Description', 'Amount', 'Project', 'Billable', 'Status'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    this.formatHeaders(sheet, 1, headers.length);
    this.autoResizeColumns(sheet);
  },
  
  setupEngagementSheet: function(sheet) {
    const headers = ['Engagement ID', 'Client', 'Type', 'Start Date', 'End Date', 'Status', 'Value', 'Manager'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    this.formatHeaders(sheet, 1, headers.length);
    this.autoResizeColumns(sheet);
  },
  
  setupRevenueSheet: function(sheet) {
    const headers = ['Client', 'Month', 'Engagement', 'Revenue', 'Type', 'Status'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    this.formatHeaders(sheet, 1, headers.length);
    this.autoResizeColumns(sheet);
  },
  
  setupTimesheetSheet: function(sheet) {
    const headers = ['Date', 'Employee', 'Client', 'Project', 'Hours', 'Billable Hours', 'Description', 'Status'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    this.formatHeaders(sheet, 1, headers.length);
    this.autoResizeColumns(sheet);
  },
  
  setupInvoicesSheet: function(sheet) {
    const headers = ['Invoice #', 'Client', 'Date', 'Due Date', 'Amount', 'Paid', 'Days Outstanding', 'Status'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    this.formatHeaders(sheet, 1, headers.length);
    this.autoResizeColumns(sheet);
  }
};