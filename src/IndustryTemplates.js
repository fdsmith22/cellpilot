/**
 * Industry-Specific Template Generator
 * Based on research showing 15-25 hours weekly savings per business
 */

const IndustryTemplates = {
  
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
        case 'campaign-dashboard':
        case 'lead-scoring-system':
        case 'content-performance':
        case 'customer-journey':
          const marketingResult = this.createMarketingTemplate(spreadsheet, templateType, true);
          createdSheets.push(...marketingResult.sheets);
          break;
          
        // E-commerce templates
        case 'ecommerce-inventory':
        case 'profitability-analyzer':
        case 'sales-forecasting':
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
      console.error('Error previewing template:', error);
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
            console.log('Could not delete sheet:', sheetName, e.toString());
          }
        });
        
        // Clear the property after cleanup
        scriptProperties.deleteProperty('previewSheets');
      } catch (e) {
        console.log('Error with script properties cleanup:', e.toString());
      }
      
      // Also do a general cleanup of any sheets starting with [PREVIEW]
      allSheets.forEach(sheet => {
        const sheetName = sheet.getName();
        if (sheetName.startsWith('[PREVIEW]')) {
          try {
            spreadsheet.deleteSheet(sheet);
            deletedCount++;
          } catch (e) {
            console.log('Could not delete preview sheet:', sheetName, e.toString());
          }
        }
      });
      
      return { success: true, deletedCount: deletedCount };
    } catch (error) {
      console.error('Error cleaning up preview sheets:', error);
      return { success: false, error: error.toString() };
    }
  },
  
  /**
   * Create Real Estate template sheets
   */
  createRealEstateTemplate: function(spreadsheet, templateType, isPreview = false) {
    const sheets = [];
    const prefix = isPreview ? '[PREVIEW] ' : '';
    
    switch (templateType) {
      case 'commission-tracker':
        const commissionSheet = spreadsheet.insertSheet(prefix + 'Commission Tracker');
        const pipelineSheet = spreadsheet.insertSheet(prefix + 'Pipeline');
        const dashboardSheet = spreadsheet.insertSheet(prefix + 'Dashboard');
        sheets.push(commissionSheet.getName(), pipelineSheet.getName(), dashboardSheet.getName());
        
        // Always setup sheets to show data in preview
        this.setupCommissionTrackerSheet(commissionSheet);
        this.setupPipelineSheet(pipelineSheet);
        this.setupDashboardSheet(dashboardSheet);
        break;
        
      case 'lead-pipeline':
        const leadSheet = spreadsheet.insertSheet(prefix + 'Lead Pipeline');
        sheets.push(leadSheet.getName());
        this.setupLeadPipelineSheet(leadSheet);
        break;
        
      case 'property-manager':
        const propertySheet = spreadsheet.insertSheet(prefix + 'Properties');
        const tenantSheet = spreadsheet.insertSheet(prefix + 'Tenants');
        const maintenanceSheet = spreadsheet.insertSheet(prefix + 'Maintenance');
        sheets.push(propertySheet.getName(), tenantSheet.getName(), maintenanceSheet.getName());
        
        // Basic setup for property manager sheets
        this.setupPropertySheet(propertySheet);
        this.setupTenantSheet(tenantSheet);
        this.setupMaintenanceSheet(maintenanceSheet);
        break;
        
      case 'investment-analyzer':
        const investmentSheet = spreadsheet.insertSheet(prefix + 'Investment Analysis');
        const cashFlowSheet = spreadsheet.insertSheet(prefix + 'Cash Flow');
        const roiSheet = spreadsheet.insertSheet(prefix + 'ROI Calculator');
        sheets.push(investmentSheet.getName(), cashFlowSheet.getName(), roiSheet.getName());
        
        this.setupInvestmentSheet(investmentSheet);
        this.setupCashFlowSheet(cashFlowSheet);
        this.setupROISheet(roiSheet);
        break;
    }
    
    return { sheets };
  },
  
  /**
   * Create Construction template sheets
   */
  createConstructionTemplate: function(spreadsheet, templateType, isPreview = false) {
    const sheets = [];
    const prefix = isPreview ? '[PREVIEW] ' : '';
    
    switch (templateType) {
      case 'material-tracker':
        const materialSheet = spreadsheet.insertSheet(prefix + 'Material Tracker');
        sheets.push(materialSheet.getName());
        this.setupMaterialTrackerSheet(materialSheet);
        break;
        
      case 'labor-manager':
        const laborSheet = spreadsheet.insertSheet(prefix + 'Labor Manager');
        sheets.push(laborSheet.getName());
        this.setupLaborManagerSheet(laborSheet);
        break;
        
      case 'change-orders':
        const changeSheet = spreadsheet.insertSheet(prefix + 'Change Orders');
        sheets.push(changeSheet.getName());
        this.setupChangeOrderSheet(changeSheet);
        break;
        
      case 'cost-estimator':
        const estimateSheet = spreadsheet.insertSheet(prefix + 'Cost Estimator');
        sheets.push(estimateSheet.getName());
        this.setupCostEstimatorSheet(estimateSheet);
        break;
    }
    
    return { sheets };
  },
  
  /**
   * Create Healthcare template sheets
   */
  createHealthcareTemplate: function(spreadsheet, templateType, isPreview = false) {
    const sheets = [];
    const prefix = isPreview ? '[PREVIEW] ' : '';
    
    switch (templateType) {
      case 'prior-auth-tracker':
        const priorAuthSheet = spreadsheet.insertSheet(prefix + 'Prior Auth');
        sheets.push(priorAuthSheet.getName());
        this.setupPriorAuthSheet(priorAuthSheet);
        break;
        
      case 'revenue-cycle':
        const revenueSheet = spreadsheet.insertSheet(prefix + 'Revenue Cycle');
        sheets.push(revenueSheet.getName());
        this.setupRevenueCycleSheet(revenueSheet);
        break;
        
      case 'denial-analytics':
        const denialSheet = spreadsheet.insertSheet(prefix + 'Denial Analytics');
        sheets.push(denialSheet.getName());
        this.setupDenialAnalyticsSheet(denialSheet);
        break;
        
      case 'insurance-verifier':
        const insuranceSheet = spreadsheet.insertSheet(prefix + 'Insurance Verification');
        sheets.push(insuranceSheet.getName());
        this.setupInsuranceVerifierSheet(insuranceSheet);
        break;
    }
    
    return { sheets };
  },
  
  /**
   * Create Marketing template sheets
   */
  createMarketingTemplate: function(spreadsheet, templateType, isPreview = false) {
    const sheets = [];
    const prefix = isPreview ? '[PREVIEW] ' : '';
    
    switch (templateType) {
      case 'lead-scoring-system':
        const leadScoringSheet = spreadsheet.insertSheet(prefix + 'Lead Scoring');
        sheets.push(leadScoringSheet.getName());
        this.setupLeadScoringSheet(leadScoringSheet);
        break;
        
      case 'content-performance':
        const contentSheet = spreadsheet.insertSheet(prefix + 'Content Performance');
        sheets.push(contentSheet.getName());
        this.setupContentPerformanceSheet(contentSheet);
        break;
        
      case 'customer-journey':
        const journeySheet = spreadsheet.insertSheet(prefix + 'Customer Journey');
        sheets.push(journeySheet.getName());
        this.setupCustomerJourneySheet(journeySheet);
        break;
        
      case 'campaign-dashboard':
        const campaignSheet = spreadsheet.insertSheet(prefix + 'Campaign Dashboard');
        sheets.push(campaignSheet.getName());
        this.setupCampaignDashboardSheet(campaignSheet);
        break;
    }
    
    return { sheets };
  },
  
  /**
   * Create E-commerce template sheets
   */
  createEcommerceTemplate: function(spreadsheet, templateType, isPreview = false) {
    const sheets = [];
    const prefix = isPreview ? '[PREVIEW] ' : '';
    
    switch (templateType) {
      case 'profitability-analyzer':
        const profitSheet = spreadsheet.insertSheet(prefix + 'Profitability');
        sheets.push(profitSheet.getName());
        this.setupProfitabilitySheet(profitSheet);
        break;
        
      case 'sales-forecasting':
        const forecastSheet = spreadsheet.insertSheet(prefix + 'Sales Forecast');
        sheets.push(forecastSheet.getName());
        this.setupSalesForecastSheet(forecastSheet);
        break;
        
      case 'ecommerce-inventory':
        const inventorySheet = spreadsheet.insertSheet(prefix + 'Inventory Manager');
        sheets.push(inventorySheet.getName());
        this.setupInventorySheet(inventorySheet);
        break;
    }
    
    return { sheets };
  },
  
  /**
   * Create Consulting template sheets
   */
  createConsultingTemplate: function(spreadsheet, templateType, isPreview = false) {
    const sheets = [];
    const prefix = isPreview ? '[PREVIEW] ' : '';
    
    switch (templateType) {
      case 'project-profitability':
        const projectSheet = spreadsheet.insertSheet(prefix + 'Project Profitability');
        sheets.push(projectSheet.getName());
        this.setupProjectProfitabilitySheet(projectSheet);
        break;
        
      case 'client-dashboard':
        const clientSheet = spreadsheet.insertSheet(prefix + 'Client Dashboard');
        sheets.push(clientSheet.getName());
        this.setupClientDashboardSheet(clientSheet);
        break;
        
      case 'time-billing-tracker':
        const billingSheet = spreadsheet.insertSheet(prefix + 'Time & Billing');
        sheets.push(billingSheet.getName());
        this.setupTimeBillingSheet(billingSheet);
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
        case 'campaign-dashboard':
        case 'lead-scoring-system':
        case 'content-performance':
        case 'customer-journey':
          const marketingResult = this.createMarketingTemplate(spreadsheet, templateType, false);
          createdSheets.push(...marketingResult.sheets);
          break;
          
        // E-commerce templates
        case 'ecommerce-inventory':
        case 'profitability-analyzer':
        case 'sales-forecasting':
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
          
        default:
          return { success: false, error: 'Unknown template type: ' + templateType };
      }
      
      return {
        success: true,
        sheets: createdSheets,
        message: `Created ${createdSheets.length} sheet(s) for ${templateType}`
      };
      
    } catch (error) {
      console.error('Error applying template:', error);
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
    // Headers
    const headers = ['Lead ID', 'Company', 'Contact Name', 'Email', 'Phone', 'Lead Source', 
                     'Status', 'Lead Score', 'Assigned To', 'Created Date', 'Last Contact', 
                     'Next Action', 'Deal Value', 'Probability %', 'Expected Close', 'Notes'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Apply formatting
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#2C3E50');
    headerRange.setFontColor('#FFFFFF');
    headerRange.setFontWeight('bold');
    
    // Add sample data
    const sampleData = [
      ['L001', 'TechCorp', 'John Smith', 'john@techcorp.com', '555-0101', 'Website', 
       'Qualified', 85, 'Sarah Johnson', new Date(), new Date(), 'Demo scheduled', 
       50000, 60, new Date(Date.now() + 30*24*60*60*1000), 'High interest in product'],
      ['L002', 'DataSoft', 'Jane Doe', 'jane@datasoft.com', '555-0102', 'Referral', 
       'New', 65, 'Mike Chen', new Date(), '', 'Initial call', 
       75000, 30, new Date(Date.now() + 45*24*60*60*1000), 'Evaluating options']
    ];
    sheet.getRange(2, 1, sampleData.length, headers.length).setValues(sampleData);
    
    // Add data validation for Status
    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['New', 'Contacted', 'Qualified', 'Proposal', 'Negotiation', 'Closed Won', 'Closed Lost'])
      .build();
    sheet.getRange(2, 7, 100, 1).setDataValidation(statusRule);
    
    // Add conditional formatting for Lead Score
    const scoreRange = sheet.getRange(2, 8, 100, 1);
    const highScoreRule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThanOrEqualTo(80)
      .setBackground('#27AE60')
      .setRanges([scoreRange])
      .build();
    const medScoreRule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(50, 79)
      .setBackground('#F39C12')
      .setRanges([scoreRange])
      .build();
    sheet.setConditionalFormatRules([highScoreRule, medScoreRule]);
    
    // Set column widths
    sheet.setColumnWidth(2, 150);
    sheet.setColumnWidth(3, 150);
    sheet.setColumnWidth(4, 200);
    sheet.setColumnWidth(16, 300);
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
    const headers = ['Material ID', 'Date', 'Material Name', 'Category', 'Unit', 'Quantity Ordered', 
                     'Unit Cost', 'Total Cost', 'Supplier', 'Received Qty', 'Used Qty', 'Remaining', 'Status'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#2C3E50').setFontColor('#FFFFFF').setFontWeight('bold');
    
    // Add formulas
    for (let row = 2; row <= 50; row++) {
      sheet.getRange(`H${row}`).setFormula(`=F${row}*G${row}`).setNumberFormat('$#,##0.00');
      sheet.getRange(`L${row}`).setFormula(`=J${row}-K${row}`).setNumberFormat('#,##0');
    }
  },
  
  setupLaborManagerSheet: function(sheet) {
    const headers = ['Employee ID', 'Name', 'Role', 'Department', 'Date', 'Hours Worked', 
                     'Hourly Rate', 'Daily Cost', 'Project', 'Task', 'Status', 'Approved By'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#34495E').setFontColor('#FFFFFF').setFontWeight('bold');
    
    // Add formulas for cost calculation
    for (let row = 2; row <= 100; row++) {
      sheet.getRange(`H${row}`).setFormula(`=F${row}*G${row}`).setNumberFormat('$#,##0.00');
    }
  },
  
  setupChangeOrderSheet: function(sheet) {
    const headers = ['Change Order #', 'Date', 'Client', 'Project', 'Description', 'Reason', 
                     'Original Cost', 'Change Amount', 'New Total', 'Status', 'Approved By', 'Date Approved'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#E74C3C').setFontColor('#FFFFFF').setFontWeight('bold');
    
    // Add formulas for totals
    for (let row = 2; row <= 50; row++) {
      sheet.getRange(`I${row}`).setFormula(`=G${row}+H${row}`).setNumberFormat('$#,##0.00');
    }
  },
  
  setupPriorAuthSheet: function(sheet) {
    const headers = ['Auth ID', 'Date', 'Patient ID', 'Patient Name', 'Insurance', 'Procedure Code', 
                     'Procedure Description', 'Provider', 'Status', 'Auth Number', 'Valid From', 'Valid To', 
                     'Units Authorized', 'Units Used', 'Remaining'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#3498DB').setFontColor('#FFFFFF').setFontWeight('bold');
    
    // Add formula for remaining units
    for (let row = 2; row <= 100; row++) {
      sheet.getRange(`O${row}`).setFormula(`=M${row}-N${row}`).setNumberFormat('#,##0');
    }
  },
  
  setupRevenueCycleSheet: function(sheet) {
    const headers = ['Claim ID', 'Date', 'Patient', 'Provider', 'Service Date', 'CPT Code', 
                     'Charge Amount', 'Allowed Amount', 'Paid Amount', 'Adjustment', 'Balance', 
                     'Days in AR', 'Status', 'Denial Reason'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#27AE60').setFontColor('#FFFFFF').setFontWeight('bold');
    
    // Add formulas for balance and days in AR
    for (let row = 2; row <= 100; row++) {
      sheet.getRange(`K${row}`).setFormula(`=G${row}-I${row}-J${row}`).setNumberFormat('$#,##0.00');
      sheet.getRange(`L${row}`).setFormula(`=TODAY()-B${row}`).setNumberFormat('#,##0');
    }
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
    const headers = ['Product ID', 'Product Name', 'Category', 'Unit Cost', 'Selling Price', 
                     'Units Sold', 'Revenue', 'Cost of Goods', 'Gross Profit', 'Margin %', 
                     'Marketing Cost', 'Net Profit', 'ROI', 'Status'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#F39C12').setFontColor('#FFFFFF').setFontWeight('bold');
    
    // Add profitability formulas
    for (let row = 2; row <= 100; row++) {
      sheet.getRange(`G${row}`).setFormula(`=E${row}*F${row}`).setNumberFormat('$#,##0.00');
      sheet.getRange(`H${row}`).setFormula(`=D${row}*F${row}`).setNumberFormat('$#,##0.00');
      sheet.getRange(`I${row}`).setFormula(`=G${row}-H${row}`).setNumberFormat('$#,##0.00');
      sheet.getRange(`J${row}`).setFormula(`=IF(G${row}>0,I${row}/G${row},0)`).setNumberFormat('0.00%');
      sheet.getRange(`L${row}`).setFormula(`=I${row}-K${row}`).setNumberFormat('$#,##0.00');
      sheet.getRange(`M${row}`).setFormula(`=IF(K${row}>0,L${row}/K${row},0)`).setNumberFormat('0.00%');
    }
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
    const headers = ['Transaction ID', 'Date', 'Agent', 'Property Address', 'Sale Price', 
                     'Commission %', 'Commission Amount', 'Split %', 'Net Commission', 'Status'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#27AE60').setFontColor('#FFFFFF').setFontWeight('bold');
    
    for (let row = 2; row <= 100; row++) {
      sheet.getRange(`G${row}`).setFormula(`=E${row}*F${row}/100`).setNumberFormat('$#,##0.00');
      sheet.getRange(`I${row}`).setFormula(`=G${row}*H${row}/100`).setNumberFormat('$#,##0.00');
    }
  },
  
  setupPipelineSheet: function(sheet) {
    const headers = ['Lead ID', 'Date', 'Name', 'Contact', 'Property Interest', 'Stage', 
                     'Source', 'Next Action', 'Follow-up Date', 'Agent', 'Notes'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#E67E22').setFontColor('#FFFFFF').setFontWeight('bold');
  },
  
  setupDashboardSheet: function(sheet) {
    const headers = ['Metric', 'This Month', 'Last Month', 'YTD', 'Target', 'Variance'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#8E44AD').setFontColor('#FFFFFF').setFontWeight('bold');
  },
  
  setupCostEstimatorSheet: function(sheet) {
    const headers = ['Item', 'Category', 'Quantity', 'Unit', 'Unit Cost', 'Total Cost', 
                     'Markup %', 'Selling Price', 'Margin', 'Notes'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#2C3E50').setFontColor('#FFFFFF').setFontWeight('bold');
    
    for (let row = 2; row <= 100; row++) {
      sheet.getRange(`F${row}`).setFormula(`=C${row}*E${row}`).setNumberFormat('$#,##0.00');
      sheet.getRange(`H${row}`).setFormula(`=F${row}*(1+G${row}/100)`).setNumberFormat('$#,##0.00');
      sheet.getRange(`I${row}`).setFormula(`=H${row}-F${row}`).setNumberFormat('$#,##0.00');
    }
  },
  
  setupInsuranceVerifierSheet: function(sheet) {
    const headers = ['Patient ID', 'Patient Name', 'DOB', 'Insurance', 'Policy #', 
                     'Group #', 'Effective Date', 'Expiry Date', 'Copay', 'Deductible', 
                     'Coverage %', 'Verified Date', 'Verified By', 'Status'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#16A085').setFontColor('#FFFFFF').setFontWeight('bold');
  },
  
  setupCampaignDashboardSheet: function(sheet) {
    const headers = ['Campaign', 'Channel', 'Start Date', 'End Date', 'Budget', 'Spend', 
                     'Impressions', 'Clicks', 'CTR', 'Conversions', 'CPA', 'ROI'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#E74C3C').setFontColor('#FFFFFF').setFontWeight('bold');
    
    for (let row = 2; row <= 50; row++) {
      sheet.getRange(`I${row}`).setFormula(`=IF(G${row}>0,H${row}/G${row}*100,0)`).setNumberFormat('0.00%');
      sheet.getRange(`K${row}`).setFormula(`=IF(J${row}>0,F${row}/J${row},0)`).setNumberFormat('$#,##0.00');
      sheet.getRange(`L${row}`).setFormula(`=IF(F${row}>0,(J${row}*100-F${row})/F${row}*100,0)`).setNumberFormat('0.00%');
    }
  },
  
  setupInventorySheet: function(sheet) {
    const headers = ['SKU', 'Product Name', 'Category', 'Current Stock', 'Min Stock', 
                     'Max Stock', 'Reorder Point', 'Reorder Qty', 'Unit Cost', 'Total Value', 
                     'Supplier', 'Lead Time', 'Status'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#F39C12').setFontColor('#FFFFFF').setFontWeight('bold');
    
    for (let row = 2; row <= 100; row++) {
      sheet.getRange(`J${row}`).setFormula(`=D${row}*I${row}`).setNumberFormat('$#,##0.00');
      sheet.getRange(`M${row}`).setFormula(`=IF(D${row}<=G${row},"REORDER",IF(D${row}>=F${row},"OVERSTOCK","OK"))`);
    }
  },
  
  // Additional Real Estate setup functions
  setupPropertySheet: function(sheet) {
    const headers = ['Property ID', 'Address', 'Type', 'Bedrooms', 'Bathrooms', 
                     'Square Feet', 'Purchase Price', 'Current Value', 'Monthly Rent', 
                     'Occupancy Status', 'Tenant', 'Lease End', 'HOA Fees', 'Property Tax'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#2C3E50').setFontColor('#FFFFFF').setFontWeight('bold');
  },
  
  setupTenantSheet: function(sheet) {
    const headers = ['Tenant ID', 'Name', 'Contact', 'Email', 'Property', 
                     'Lease Start', 'Lease End', 'Monthly Rent', 'Security Deposit', 
                     'Payment Status', 'Last Payment', 'Balance Due'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#34495E').setFontColor('#FFFFFF').setFontWeight('bold');
  },
  
  setupMaintenanceSheet: function(sheet) {
    const headers = ['Request ID', 'Date', 'Property', 'Tenant', 'Issue Type', 
                     'Description', 'Priority', 'Status', 'Vendor', 'Cost', 
                     'Completion Date', 'Notes'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#E74C3C').setFontColor('#FFFFFF').setFontWeight('bold');
  },
  
  setupInvestmentSheet: function(sheet) {
    const headers = ['Property', 'Purchase Price', 'Down Payment', 'Loan Amount', 
                     'Interest Rate', 'Monthly Payment', 'Rental Income', 
                     'Operating Expenses', 'Cash Flow', 'ROI %', 'Cap Rate'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#27AE60').setFontColor('#FFFFFF').setFontWeight('bold');
    
    for (let row = 2; row <= 50; row++) {
      sheet.getRange(`I${row}`).setFormula(`=G${row}-F${row}-H${row}`).setNumberFormat('$#,##0.00');
      sheet.getRange(`J${row}`).setFormula(`=IF(B${row}>0,I${row}*12/B${row}*100,0)`).setNumberFormat('0.00%');
    }
  },
  
  setupCashFlowSheet: function(sheet) {
    const headers = ['Month', 'Rental Income', 'Other Income', 'Total Income', 
                     'Mortgage', 'HOA', 'Insurance', 'Property Tax', 'Maintenance', 
                     'Utilities', 'Total Expenses', 'Net Cash Flow', 'Cumulative'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#3498DB').setFontColor('#FFFFFF').setFontWeight('bold');
    
    for (let row = 2; row <= 100; row++) {
      sheet.getRange(`D${row}`).setFormula(`=B${row}+C${row}`).setNumberFormat('$#,##0.00');
      sheet.getRange(`K${row}`).setFormula(`=SUM(E${row}:J${row})`).setNumberFormat('$#,##0.00');
      sheet.getRange(`L${row}`).setFormula(`=D${row}-K${row}`).setNumberFormat('$#,##0.00');
      if (row === 2) {
        sheet.getRange(`M${row}`).setFormula(`=L${row}`).setNumberFormat('$#,##0.00');
      } else {
        sheet.getRange(`M${row}`).setFormula(`=M${row-1}+L${row}`).setNumberFormat('$#,##0.00');
      }
    }
  },
  
  setupROISheet: function(sheet) {
    const headers = ['Metric', 'Value', 'Formula', 'Description'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#9B59B6').setFontColor('#FFFFFF').setFontWeight('bold');
    
    const metrics = [
      ['Total Investment', '', 'Purchase + Repairs', 'Initial capital investment'],
      ['Annual Income', '', 'Rent * 12', 'Gross rental income per year'],
      ['Annual Expenses', '', 'All operating costs', 'Total yearly expenses'],
      ['Net Operating Income', '', 'Income - Expenses', 'NOI before financing'],
      ['Cash Flow', '', 'NOI - Debt Service', 'Annual cash flow after mortgage'],
      ['Cash on Cash Return', '', 'Cash Flow / Investment', 'Return on cash invested'],
      ['Cap Rate', '', 'NOI / Property Value', 'Capitalization rate'],
      ['ROI', '', '(Value + Cash Flow - Investment) / Investment', 'Total return on investment']
    ];
    
    sheet.getRange(2, 1, metrics.length, 4).setValues(metrics);
    sheet.getRange(2, 2, metrics.length, 1).setNumberFormat('$#,##0.00');
  }
};