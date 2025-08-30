/**
 * Professional Real Estate Templates Suite
 * Complete templates for all real estate professionals
 */

var RealEstateTemplate = {
  /**
   * Build comprehensive Real Estate template based on type
   */
  build: function(spreadsheet, templateType = 'commission-tracker', isPreview = false) {
    const builder = new TemplateBuilderPro(spreadsheet, isPreview);
    
    switch(templateType) {
      case 'commission-tracker':
        return this.buildCommissionTrackerTemplate(builder);
      case 'property-manager':
        return this.buildPropertyManagerTemplate(builder);
      case 'investment-analyzer':
        return this.buildInvestmentAnalyzerTemplate(builder);
      case 'lead-pipeline':
        return this.buildLeadPipelineTemplate(builder);
      default:
        return this.buildCommissionTrackerTemplate(builder);
    }
  },
  
  /**
   * Commission Tracker Template
   */
  buildCommissionTrackerTemplate: function(builder) {
    // Create all sheets
    const sheets = builder.createSheets([
      'Dashboard',
      'Commission Tracker',
      'Pipeline',
      'Clients',
      'Analytics',
      'Market Insights',
      'Goals & Targets',
      'Reports'
    ]);
    
    // Build each component
    this.buildDashboard(builder);
    this.buildCommissionTracker(builder);
    this.buildPipeline(builder);
    this.buildClients(builder);
    this.buildAnalytics(builder);
    this.buildMarketInsights(builder);
    this.buildGoalsTargets(builder);
    this.buildReports(builder);
    
    return builder.complete();
  },
  
  /**
   * Property Manager Template
   */
  buildPropertyManagerTemplate: function(builder) {
    // Create sheets for property management
    const sheets = builder.createSheets([
      'Dashboard',
      'Properties',
      'Tenants',
      'Lease Tracker',
      'Maintenance',
      'Financials',
      'Vendors',
      'Reports'
    ]);
    
    this.buildPropertyDashboard(builder);
    this.buildPropertiesSheet(builder);
    this.buildTenantsSheet(builder);
    this.buildLeaseTracker(builder);
    this.buildMaintenanceLog(builder);
    this.buildPropertyFinancials(builder);
    this.buildVendorsSheet(builder);
    this.buildPropertyReports(builder);
    
    return builder.complete();
  },
  
  /**
   * Investment Analyzer Template
   */
  buildInvestmentAnalyzerTemplate: function(builder) {
    // Create sheets for investment analysis
    const sheets = builder.createSheets([
      'Dashboard',
      'Properties',
      'Cash Flow Analysis',
      'ROI Calculator',
      'Market Comps',
      'Financing',
      'Tax Analysis',
      'Reports'
    ]);
    
    this.buildInvestmentDashboard(builder);
    this.buildInvestmentProperties(builder);
    this.buildCashFlowAnalysis(builder);
    this.buildROICalculator(builder);
    this.buildMarketComps(builder);
    this.buildFinancingSheet(builder);
    this.buildTaxAnalysis(builder);
    this.buildInvestmentReports(builder);
    
    return builder.complete();
  },
  
  /**
   * Lead Pipeline Template
   */
  buildLeadPipelineTemplate: function(builder) {
    // Create sheets for lead management
    const sheets = builder.createSheets([
      'Dashboard',
      'Lead Pipeline',
      'Contact Log',
      'Lead Sources',
      'Follow-ups',
      'Conversion Analytics',
      'Marketing ROI',
      'Reports'
    ]);
    
    this.buildLeadDashboard(builder);
    this.buildLeadPipelineSheet(builder);
    this.buildContactLog(builder);
    this.buildLeadSources(builder);
    this.buildFollowUps(builder);
    this.buildConversionAnalytics(builder);
    this.buildMarketingROI(builder);
    this.buildLeadReports(builder);
    
    return builder.complete();
  },
  
  /**
   * Build Professional Dashboard
   */
  buildDashboard: function(builder) {
    const dash = builder.sheet('Dashboard');
    
    // ========================================
    // HEADER SECTION
    // ========================================
    dash.merge(1, 1, 2, 12);
    dash.setValue(1, 1, 'Real Estate Performance Command Center');
    dash.format(1, 1, 2, 12, {
      background: '#1E3A8A',
      fontColor: '#FFFFFF',
      fontSize: 18,
      fontWeight: 'bold',
      horizontalAlignment: 'center',
      verticalAlignment: 'middle'
    });
    
    // Date and refresh info
    dash.setValue(3, 1, 'Last Updated:');
    dash.setFormula(3, 2, '=TEXT(NOW(),"mmm dd, yyyy h:mm AM/PM")');
    dash.setValue(3, 4, 'Current Month:');
    dash.setFormula(3, 5, '=TEXT(TODAY(),"MMMM yyyy")');
    dash.setValue(3, 7, 'Days Remaining:');
    dash.setFormula(3, 8, '=EOMONTH(TODAY(),0)-TODAY()');
    
    // ========================================
    // PRIMARY KPI CARDS (Row 5-9)
    // ========================================
    const kpiCards = [
      {
        row: 5, col: 1, width: 3,
        title: 'YTD Revenue',
        formula: '=SUMIFS({{Commission Tracker}}!K:K,{{Commission Tracker}}!B:B,">="&DATE(YEAR(TODAY()),1,1),{{Commission Tracker}}!M:M,"Closed")',
        format: '$#,##0',
        color: '#10B981',
        icon: '💰',
        comparison: '=IF(SUMIFS({{Commission Tracker}}!K:K,{{Commission Tracker}}!B:B,">="&DATE(YEAR(TODAY()),1,1),{{Commission Tracker}}!M:M,"Closed")>SUMIFS({{Commission Tracker}}!K:K,{{Commission Tracker}}!B:B,">="&DATE(YEAR(TODAY())-1,1,1),{{Commission Tracker}}!B:B,"<="&DATE(YEAR(TODAY())-1,12,31),{{Commission Tracker}}!M:M,"Closed"),"↑ "&TEXT((SUMIFS({{Commission Tracker}}!K:K,{{Commission Tracker}}!B:B,">="&DATE(YEAR(TODAY()),1,1),{{Commission Tracker}}!M:M,"Closed")/SUMIFS({{Commission Tracker}}!K:K,{{Commission Tracker}}!B:B,">="&DATE(YEAR(TODAY())-1,1,1),{{Commission Tracker}}!B:B,"<="&DATE(YEAR(TODAY())-1,12,31),{{Commission Tracker}}!M:M,"Closed")-1),"0%")&" vs Last Year","↓ "&TEXT((1-SUMIFS({{Commission Tracker}}!K:K,{{Commission Tracker}}!B:B,">="&DATE(YEAR(TODAY()),1,1),{{Commission Tracker}}!M:M,"Closed")/SUMIFS({{Commission Tracker}}!K:K,{{Commission Tracker}}!B:B,">="&DATE(YEAR(TODAY())-1,1,1),{{Commission Tracker}}!B:B,"<="&DATE(YEAR(TODAY())-1,12,31),{{Commission Tracker}}!M:M,"Closed")),"0%")&" vs Last Year")'
      },
      {
        row: 5, col: 4, width: 3,
        title: 'Active Pipeline Value',
        formula: '=SUMIFS({{Pipeline}}!G:G,{{Pipeline}}!H:H,"<>Closed",{{Pipeline}}!H:H,"<>Lost")',
        format: '$#,##0',
        color: '#3B82F6',
        icon: '📊',
        comparison: '={{Pipeline}}!A1000' // Placeholder for pipeline count
      },
      {
        row: 5, col: 7, width: 3,
        title: 'This Month Commission',
        formula: '=SUMIFS({{Commission Tracker}}!K:K,{{Commission Tracker}}!B:B,">="&EOMONTH(TODAY(),-1)+1,{{Commission Tracker}}!B:B,"<="&EOMONTH(TODAY(),0),{{Commission Tracker}}!M:M,"Closed")',
        format: '$#,##0',
        color: '#8B5CF6',
        icon: '📈',
        comparison: '="Target: $"&TEXT(50000,"#,##0")'
      },
      {
        row: 5, col: 10, width: 3,
        title: 'Avg Deal Size',
        formula: '=IFERROR(AVERAGE({{Commission Tracker}}!F:F),0)',
        format: '$#,##0',
        color: '#F59E0B',
        icon: '🏠',
        comparison: '="Market Avg: $"&TEXT(425000,"#,##0")'
      }
    ];
    
    // Create KPI cards with professional styling
    kpiCards.forEach(kpi => {
      // Card background
      dash.merge(kpi.row, kpi.col, kpi.row + 3, kpi.col + kpi.width - 1);
      dash.format(kpi.row, kpi.col, kpi.row + 3, kpi.col + kpi.width - 1, {
        background: '#FFFFFF',
        border: true
      });
      
      // Icon and title
      dash.setValue(kpi.row, kpi.col, kpi.icon + ' ' + kpi.title);
      dash.format(kpi.row, kpi.col, kpi.row, kpi.col + kpi.width - 1, {
        fontSize: 11,
        fontColor: '#6B7280',
        fontWeight: 'bold'
      });
      
      // Value
      dash.setFormula(kpi.row + 1, kpi.col, kpi.formula);
      dash.format(kpi.row + 1, kpi.col, kpi.row + 1, kpi.col + kpi.width - 1, {
        fontSize: 20,
        fontWeight: 'bold',
        fontColor: kpi.color,
        numberFormat: kpi.format
      });
      
      // Comparison
      dash.setFormula(kpi.row + 2, kpi.col, kpi.comparison);
      dash.format(kpi.row + 2, kpi.col, kpi.row + 2, kpi.col + kpi.width - 1, {
        fontSize: 10,
        fontColor: '#9CA3AF'
      });
    });
    
    // ========================================
    // PERFORMANCE METRICS GRID (Row 10-20)
    // ========================================
    dash.setValue(10, 1, '📊 Performance Metrics');
    dash.format(10, 1, 10, 12, {
      background: '#F3F4F6',
      fontWeight: 'bold',
      fontSize: 12
    });
    
    // Metrics headers
    const metricsHeaders = ['Metric', 'This Month', 'Last Month', 'Change %', 'YTD', 'Target', 'Progress'];
    dash.setRangeValues(11, 1, [metricsHeaders]);
    dash.format(11, 1, 11, 7, {
      background: '#E5E7EB',
      fontWeight: 'bold',
      fontSize: 10
    });
    
    // Metrics data with formulas
    const metrics = [
      ['Deals Closed', 
       '=COUNTIFS({{Commission Tracker}}!B:B,">="&EOMONTH(TODAY(),-1)+1,{{Commission Tracker}}!B:B,"<="&EOMONTH(TODAY(),0),{{Commission Tracker}}!M:M,"Closed")',
       '=COUNTIFS({{Commission Tracker}}!B:B,">="&EOMONTH(TODAY(),-2)+1,{{Commission Tracker}}!B:B,"<="&EOMONTH(TODAY(),-1),{{Commission Tracker}}!M:M,"Closed")',
       '=IFERROR((B12-C12)/C12,0)',
       '=COUNTIFS({{Commission Tracker}}!B:B,">="&DATE(YEAR(TODAY()),1,1),{{Commission Tracker}}!M:M,"Closed")',
       '10',
       '=E12/F12'],
      ['Total Volume',
       '=SUMIFS({{Commission Tracker}}!F:F,{{Commission Tracker}}!B:B,">="&EOMONTH(TODAY(),-1)+1,{{Commission Tracker}}!B:B,"<="&EOMONTH(TODAY(),0),{{Commission Tracker}}!M:M,"Closed")',
       '=SUMIFS({{Commission Tracker}}!F:F,{{Commission Tracker}}!B:B,">="&EOMONTH(TODAY(),-2)+1,{{Commission Tracker}}!B:B,"<="&EOMONTH(TODAY(),-1),{{Commission Tracker}}!M:M,"Closed")',
       '=IFERROR((B13-C13)/C13,0)',
       '=SUMIFS({{Commission Tracker}}!F:F,{{Commission Tracker}}!B:B,">="&DATE(YEAR(TODAY()),1,1),{{Commission Tracker}}!M:M,"Closed")',
       '5000000',
       '=E13/F13'],
      ['Gross Commission',
       '=SUMIFS({{Commission Tracker}}!I:I,{{Commission Tracker}}!B:B,">="&EOMONTH(TODAY(),-1)+1,{{Commission Tracker}}!B:B,"<="&EOMONTH(TODAY(),0),{{Commission Tracker}}!M:M,"Closed")',
       '=SUMIFS({{Commission Tracker}}!I:I,{{Commission Tracker}}!B:B,">="&EOMONTH(TODAY(),-2)+1,{{Commission Tracker}}!B:B,"<="&EOMONTH(TODAY(),-1),{{Commission Tracker}}!M:M,"Closed")',
       '=IFERROR((B14-C14)/C14,0)',
       '=SUMIFS({{Commission Tracker}}!I:I,{{Commission Tracker}}!B:B,">="&DATE(YEAR(TODAY()),1,1),{{Commission Tracker}}!M:M,"Closed")',
       '150000',
       '=E14/F14'],
      ['Net Commission',
       '=SUMIFS({{Commission Tracker}}!K:K,{{Commission Tracker}}!B:B,">="&EOMONTH(TODAY(),-1)+1,{{Commission Tracker}}!B:B,"<="&EOMONTH(TODAY(),0),{{Commission Tracker}}!M:M,"Closed")',
       '=SUMIFS({{Commission Tracker}}!K:K,{{Commission Tracker}}!B:B,">="&EOMONTH(TODAY(),-2)+1,{{Commission Tracker}}!B:B,"<="&EOMONTH(TODAY(),-1),{{Commission Tracker}}!M:M,"Closed")',
       '=IFERROR((B15-C15)/C15,0)',
       '=SUMIFS({{Commission Tracker}}!K:K,{{Commission Tracker}}!B:B,">="&DATE(YEAR(TODAY()),1,1),{{Commission Tracker}}!M:M,"Closed")',
       '100000',
       '=E15/F15'],
      ['Active Listings',
       '=COUNTIF({{Pipeline}}!H:H,"Active")+COUNTIF({{Pipeline}}!H:H,"Hot")',
       '=COUNTIF({{Pipeline}}!H:H,"Active")+COUNTIF({{Pipeline}}!H:H,"Hot")',
       '=IFERROR((B16-C16)/C16,0)',
       '=COUNTIFS({{Pipeline}}!B:B,">="&DATE(YEAR(TODAY()),1,1),{{Pipeline}}!H:H,"<>Closed",{{Pipeline}}!H:H,"<>Lost")',
       '25',
       '=E16/F16'],
      ['Conversion Rate',
       '=IFERROR(COUNTIFS({{Commission Tracker}}!B:B,">="&EOMONTH(TODAY(),-1)+1,{{Commission Tracker}}!B:B,"<="&EOMONTH(TODAY(),0),{{Commission Tracker}}!M:M,"Closed")/COUNTIFS({{Pipeline}}!B:B,">="&EOMONTH(TODAY(),-1)+1,{{Pipeline}}!B:B,"<="&EOMONTH(TODAY(),0)),0)',
       '=IFERROR(COUNTIFS({{Commission Tracker}}!B:B,">="&EOMONTH(TODAY(),-2)+1,{{Commission Tracker}}!B:B,"<="&EOMONTH(TODAY(),-1),{{Commission Tracker}}!M:M,"Closed")/COUNTIFS({{Pipeline}}!B:B,">="&EOMONTH(TODAY(),-2)+1,{{Pipeline}}!B:B,"<="&EOMONTH(TODAY(),-1)),0)',
       '=IFERROR((B17-C17)/C17,0)',
       '=IFERROR(COUNTIFS({{Commission Tracker}}!M:M,"Closed")/COUNTA({{Pipeline}}!A:A),0)',
       '0.25',
       '=E17/F17'],
      ['Avg Days to Close',
       '=IFERROR(AVERAGE({{Commission Tracker}}!P:P),0)',
       '=IFERROR(AVERAGE({{Commission Tracker}}!P:P),0)',
       '=IFERROR((B18-C18)/C18,0)',
       '=IFERROR(AVERAGE({{Commission Tracker}}!P:P),0)',
       '30',
       '=1-E18/F18']
    ];
    
    dash.setRangeValues(12, 1, metrics);
    
    // Format metrics grid
    dash.format(12, 2, 18, 2, { numberFormat: '#,##0' });
    dash.format(12, 3, 18, 3, { numberFormat: '#,##0' });
    dash.format(12, 4, 18, 4, { numberFormat: '0%' });
    dash.format(12, 5, 18, 5, { numberFormat: '#,##0' });
    dash.format(12, 6, 18, 6, { numberFormat: '#,##0' });
    dash.format(12, 7, 18, 7, { numberFormat: '0%' });
    dash.format(13, 2, 15, 5, { numberFormat: '$#,##0' });
    dash.format(13, 6, 15, 6, { numberFormat: '$#,##0' });
    dash.format(17, 2, 17, 6, { numberFormat: '0%' });
    
    // Add conditional formatting for progress bars
    for (let row = 12; row <= 18; row++) {
      dash.addConditionalFormat(row, 7, row, 7, {
        type: 'gradient',
        minValue: 0,
        midValue: 0.5,
        maxValue: 1,
        minColor: '#EF4444',
        midColor: '#F59E0B',
        maxColor: '#10B981'
      });
    }
    
    // ========================================
    // PIPELINE FUNNEL (Row 10-20, Columns 8-12)
    // ========================================
    dash.setValue(10, 8, '🎯 Pipeline Funnel');
    dash.format(10, 8, 10, 12, {
      background: '#F3F4F6',
      fontWeight: 'bold',
      fontSize: 12
    });
    
    const funnelStages = [
      ['Stage', 'Count', 'Value', 'Avg Value', 'Conv %'],
      ['New Leads', '=COUNTIF({{Pipeline}}!H:H,"New")', '=SUMIF({{Pipeline}}!H:H,"New",{{Pipeline}}!G:G)', '=IFERROR(C12/B12,0)', '100%'],
      ['Contacted', '=COUNTIF({{Pipeline}}!H:H,"Contacted")', '=SUMIF({{Pipeline}}!H:H,"Contacted",{{Pipeline}}!G:G)', '=IFERROR(C13/B13,0)', '=IFERROR(B13/B12,0)'],
      ['Qualified', '=COUNTIF({{Pipeline}}!H:H,"Qualified")', '=SUMIF({{Pipeline}}!H:H,"Qualified",{{Pipeline}}!G:G)', '=IFERROR(C14/B14,0)', '=IFERROR(B14/B12,0)'],
      ['Active', '=COUNTIF({{Pipeline}}!H:H,"Active")', '=SUMIF({{Pipeline}}!H:H,"Active",{{Pipeline}}!G:G)', '=IFERROR(C15/B15,0)', '=IFERROR(B15/B12,0)'],
      ['Hot', '=COUNTIF({{Pipeline}}!H:H,"Hot")', '=SUMIF({{Pipeline}}!H:H,"Hot",{{Pipeline}}!G:G)', '=IFERROR(C16/B16,0)', '=IFERROR(B16/B12,0)'],
      ['In Contract', '=COUNTIF({{Pipeline}}!H:H,"Contract")', '=SUMIF({{Pipeline}}!H:H,"Contract",{{Pipeline}}!G:G)', '=IFERROR(C17/B17,0)', '=IFERROR(B17/B12,0)'],
      ['Closed Won', '=COUNTIF({{Pipeline}}!H:H,"Closed")', '=SUMIF({{Pipeline}}!H:H,"Closed",{{Pipeline}}!G:G)', '=IFERROR(C18/B18,0)', '=IFERROR(B18/B12,0)']
    ];
    
    dash.setRangeValues(11, 8, funnelStages);
    dash.format(11, 8, 11, 12, {
      background: '#E5E7EB',
      fontWeight: 'bold',
      fontSize: 10
    });
    dash.format(12, 10, 18, 10, { numberFormat: '$#,##0' });
    dash.format(12, 11, 18, 11, { numberFormat: '$#,##0' });
    dash.format(12, 12, 18, 12, { numberFormat: '0%' });
    
    // ========================================
    // TOP PERFORMERS (Row 21-30)
    // ========================================
    dash.setValue(21, 1, '🏆 Top Performers This Month');
    dash.format(21, 1, 21, 6, {
      background: '#F3F4F6',
      fontWeight: 'bold',
      fontSize: 12
    });
    
    const topPerformersHeaders = ['Rank', 'Agent', 'Deals', 'Volume', 'Commission', 'Avg Deal'];
    dash.setRangeValues(22, 1, [topPerformersHeaders]);
    dash.format(22, 1, 22, 6, {
      background: '#E5E7EB',
      fontWeight: 'bold',
      fontSize: 10
    });
    
    // Add QUERY formula for top performers
    dash.setFormula(23, 1, 
      '=IFERROR(QUERY({{Commission Tracker}}!B:T,' +
      '"SELECT L, COUNT(A), SUM(F), SUM(K), AVG(F) ' +
      'WHERE B >= date \'"&TEXT(EOMONTH(TODAY(),-1)+1,"yyyy-mm-dd")&"\' ' +
      'AND B <= date \'"&TEXT(EOMONTH(TODAY(),0),"yyyy-mm-dd")&"\' ' +
      'AND M = \'Closed\' ' +
      'GROUP BY L ' +
      'ORDER BY SUM(K) DESC ' +
      'LIMIT 5 ' +
      'LABEL L \'Agent\', COUNT(A) \'Deals\', SUM(F) \'Volume\', SUM(K) \'Commission\', AVG(F) \'Avg Deal\'",0),"")'
    );
    
    // ========================================
    // LEAD SOURCE ANALYSIS (Row 21-30, Columns 7-12)
    // ========================================
    dash.setValue(21, 7, '📍 Lead Source Performance');
    dash.format(21, 7, 21, 12, {
      background: '#F3F4F6',
      fontWeight: 'bold',
      fontSize: 12
    });
    
    const leadSourceHeaders = ['Source', 'Leads', 'Conversions', 'Conv Rate', 'Revenue', 'ROI'];
    dash.setRangeValues(22, 7, [leadSourceHeaders]);
    dash.format(22, 7, 22, 12, {
      background: '#E5E7EB',
      fontWeight: 'bold',
      fontSize: 10
    });
    
    // Lead source data
    const leadSources = [
      ['Referral', '=COUNTIF({{Pipeline}}!I:I,"Referral")', '=COUNTIFS({{Commission Tracker}}!N:N,"Referral",{{Commission Tracker}}!M:M,"Closed")', '=IFERROR(I23/H23,0)', '=SUMIFS({{Commission Tracker}}!K:K,{{Commission Tracker}}!N:N,"Referral",{{Commission Tracker}}!M:M,"Closed")', '=K23/1000'],
      ['Website', '=COUNTIF({{Pipeline}}!I:I,"Website")', '=COUNTIFS({{Commission Tracker}}!N:N,"Website",{{Commission Tracker}}!M:M,"Closed")', '=IFERROR(I24/H24,0)', '=SUMIFS({{Commission Tracker}}!K:K,{{Commission Tracker}}!N:N,"Website",{{Commission Tracker}}!M:M,"Closed")', '=K24/2000'],
      ['Social Media', '=COUNTIF({{Pipeline}}!I:I,"Social Media")', '=COUNTIFS({{Commission Tracker}}!N:N,"Social Media",{{Commission Tracker}}!M:M,"Closed")', '=IFERROR(I25/H25,0)', '=SUMIFS({{Commission Tracker}}!K:K,{{Commission Tracker}}!N:N,"Social Media",{{Commission Tracker}}!M:M,"Closed")', '=K25/1500'],
      ['Open House', '=COUNTIF({{Pipeline}}!I:I,"Open House")', '=COUNTIFS({{Commission Tracker}}!N:N,"Open House",{{Commission Tracker}}!M:M,"Closed")', '=IFERROR(I26/H26,0)', '=SUMIFS({{Commission Tracker}}!K:K,{{Commission Tracker}}!N:N,"Open House",{{Commission Tracker}}!M:M,"Closed")', '=K26/500'],
      ['Cold Call', '=COUNTIF({{Pipeline}}!I:I,"Cold Call")', '=COUNTIFS({{Commission Tracker}}!N:N,"Cold Call",{{Commission Tracker}}!M:M,"Closed")', '=IFERROR(I27/H27,0)', '=SUMIFS({{Commission Tracker}}!K:K,{{Commission Tracker}}!N:N,"Cold Call",{{Commission Tracker}}!M:M,"Closed")', '=K27/800']
    ];
    
    dash.setRangeValues(23, 7, leadSources);
    dash.format(23, 10, 27, 10, { numberFormat: '0%' });
    dash.format(23, 11, 27, 11, { numberFormat: '$#,##0' });
    dash.format(23, 12, 27, 12, { numberFormat: '0.0' });
    
    // ========================================
    // MONTHLY TREND CHART DATA (Row 32-45)
    // ========================================
    dash.setValue(32, 1, '📈 12-Month Performance Trend');
    dash.format(32, 1, 32, 12, {
      background: '#F3F4F6',
      fontWeight: 'bold',
      fontSize: 12
    });
    
    // Create month labels and data
    dash.setValue(33, 1, 'Month');
    for (let i = 0; i < 12; i++) {
      dash.setFormula(33, i + 2, `=TEXT(EDATE(TODAY(),-${11-i}),"MMM")`);
    }
    
    dash.setValue(34, 1, 'Deals');
    for (let i = 0; i < 12; i++) {
      dash.setFormula(34, i + 2, 
        `=COUNTIFS({{Commission Tracker}}!B:B,">="&EOMONTH(EDATE(TODAY(),-${11-i}),-1)+1,` +
        `{{Commission Tracker}}!B:B,"<="&EOMONTH(EDATE(TODAY(),-${11-i}),0),` +
        `{{Commission Tracker}}!M:M,"Closed")`
      );
    }
    
    dash.setValue(35, 1, 'Revenue');
    for (let i = 0; i < 12; i++) {
      dash.setFormula(35, i + 2, 
        `=SUMIFS({{Commission Tracker}}!K:K,` +
        `{{Commission Tracker}}!B:B,">="&EOMONTH(EDATE(TODAY(),-${11-i}),-1)+1,` +
        `{{Commission Tracker}}!B:B,"<="&EOMONTH(EDATE(TODAY(),-${11-i}),0),` +
        `{{Commission Tracker}}!M:M,"Closed")`
      );
    }
    
    dash.format(35, 2, 35, 13, { numberFormat: '$#,##0' });
    
    // Add sparkline
    dash.setValue(36, 1, 'Trend');
    dash.setFormula(36, 2, '=SPARKLINE(B35:M35,{"charttype","column";"color","#3B82F6";"max",MAX(B35:M35)*1.2})');
    dash.merge(36, 2, 36, 13);
    
    // ========================================
    // QUICK ACTIONS & FILTERS (Row 38-42)
    // ========================================
    dash.setValue(38, 1, '⚡ Quick Actions');
    dash.format(38, 1, 38, 12, {
      background: '#F3F4F6',
      fontWeight: 'bold',
      fontSize: 12
    });
    
    // Action buttons (simulated with cells)
    const actions = [
      { row: 39, col: 1, text: '➕ Add New Deal', color: '#10B981' },
      { row: 39, col: 3, text: '👤 Add Client', color: '#3B82F6' },
      { row: 39, col: 5, text: '📧 Send Report', color: '#8B5CF6' },
      { row: 39, col: 7, text: '📊 Export Data', color: '#F59E0B' },
      { row: 39, col: 9, text: '🎯 Set Goals', color: '#EF4444' },
      { row: 39, col: 11, text: '📅 Calendar', color: '#06B6D4' }
    ];
    
    actions.forEach(action => {
      dash.merge(action.row, action.col, action.row, action.col + 1);
      dash.setValue(action.row, action.col, action.text);
      dash.format(action.row, action.col, action.row, action.col + 1, {
        background: action.color,
        fontColor: '#FFFFFF',
        fontWeight: 'bold',
        horizontalAlignment: 'center',
        border: true
      });
    });
    
    // ========================================
    // FOOTER WITH KEY INSIGHTS
    // ========================================
    dash.setValue(42, 1, '💡 Key Insights & Recommendations');
    dash.format(42, 1, 42, 12, {
      background: '#FEF3C7',
      fontWeight: 'bold',
      fontSize: 11,
      fontColor: '#92400E'
    });
    
    // Dynamic insights based on data
    dash.setFormula(43, 1, 
      '=IF(E12/F12>1,"🎯 Exceeding yearly target by "&TEXT(E12/F12-1,"0%")&"! Keep up the momentum.",' +
      'IF(E12/F12>0.8,"📈 At "&TEXT(E12/F12,"0%")&" of yearly target. Strong performance!",' +
      '"⚠️ Currently at "&TEXT(E12/F12,"0%")&" of target. Focus on pipeline conversion."))'
    );
    dash.merge(43, 1, 43, 12);
    
    dash.setFormula(44, 1,
      '=IF(B17>0.3,"✨ Excellent "&TEXT(B17,"0%")&" conversion rate this month!",' +
      'IF(B17>0.2,"👍 Good "&TEXT(B17,"0%")&" conversion rate. Room for improvement.",' +
      '"🔍 Low "&TEXT(B17,"0%")&" conversion rate. Review lead qualification process."))'
    );
    dash.merge(44, 1, 44, 12);
    
    dash.setFormula(45, 1,
      '="📊 Top performing lead source: "&INDEX(G23:G27,MATCH(MAX(K23:K27),K23:K27,0))&" with $"&TEXT(MAX(K23:K27),"#,##0")&" revenue"'
    );
    dash.merge(45, 1, 45, 12);
    
    // Format the entire dashboard
    dash.format(1, 1, 45, 12, {
      fontFamily: 'Arial',
      fontSize: 10
    });
    
    // Set column widths for better layout
    for (let col = 1; col <= 12; col++) {
      dash.setColumnWidth(col, 90);
    }
    
    // Set row heights for better spacing
    dash.setRowHeight(1, 30);
    for (let row = 5; row <= 8; row++) {
      dash.setRowHeight(row, 25);
    }
  },
  
  /**
   * Build Commission Tracker Sheet
   */
  buildCommissionTracker: function(builder) {
    const tracker = builder.sheet('Commission Tracker');
    
    // Headers
    const headers = [
      'Deal ID', 'Date', 'Client', 'Property Address', 'Property Type', 'Sale Price', 
      'List Price', 'Commission Rate', 'Gross Commission', 'Split %', 'Net Commission', 
      'Co-Agent', 'Status', 'Lead Source', 'Days on Market', 'Days to Close', 
      'Closing Date', 'Escrow Company', 'Transaction Type', 'Notes'
    ];
    
    tracker.setRangeValues(1, 1, [headers]);
    tracker.format(1, 1, 1, headers.length, {
      background: '#1E40AF',
      fontColor: '#FFFFFF',
      fontWeight: 'bold',
      fontSize: 11,
      horizontalAlignment: 'center',
      border: true
    });
    
    // Sample data with formulas
    const sampleData = [
      ['D001', '=TODAY()-45', 'John & Mary Smith', '123 Main St, Anytown', 'Single Family', 450000, 465000, '3%', '=F2*H2', '50%', '=I2*J2', 'Jane Doe', 'Closed', 'Referral', 30, 45, '=B2+P2', 'First American', 'Buy', 'Smooth transaction'],
      ['D002', '=TODAY()-30', 'Robert Johnson', '456 Oak Ave, Somewhere', 'Condo', 325000, 335000, '2.5%', '=F3*H3', '50%', '=I3*J3', '', 'Closed', 'Website', 25, 35, '=B3+P3', 'Chicago Title', 'Buy', 'First-time buyer'],
      ['D003', '=TODAY()-15', 'Sarah Williams', '789 Elm Dr, Elsewhere', 'Townhouse', 580000, 595000, '3%', '=F4*H4', '60%', '=I4*J4', 'Bob Smith', 'Pending', 'Open House', 40, 0, '', 'Stewart Title', 'Sell', 'Awaiting inspection'],
      ['D004', '=TODAY()-7', 'Michael Brown', '321 Pine Ln, Anywhere', 'Single Family', 675000, 699000, '2.5%', '=F5*H5', '55%', '=I5*J5', '', 'Active', 'Past Client', 15, 0, '', '', 'Buy/Sell', 'Contingent offer'],
      ['D005', '=TODAY()-3', 'Jennifer Davis', '654 Maple Ct, Nowhere', 'Luxury Home', 1250000, 1299000, '2.5%', '=F6*H6', '50%', '=I6*J6', 'Tom Wilson', 'Active', 'Partner Referral', 7, 0, '', '', 'Buy', 'High-value client']
    ];
    
    tracker.setRangeValues(2, 1, sampleData);
    
    // Format columns
    tracker.format(2, 2, 100, 2, { numberFormat: 'mm/dd/yyyy' }); // Date
    tracker.format(2, 6, 100, 7, { numberFormat: '$#,##0' }); // Prices
    tracker.format(2, 8, 100, 8, { numberFormat: '0.0%' }); // Commission Rate
    tracker.format(2, 9, 100, 9, { numberFormat: '$#,##0' }); // Gross Commission
    tracker.format(2, 10, 100, 10, { numberFormat: '0%' }); // Split
    tracker.format(2, 11, 100, 11, { numberFormat: '$#,##0' }); // Net Commission
    tracker.format(2, 17, 100, 17, { numberFormat: 'mm/dd/yyyy' }); // Closing Date
    
    // Add data validations
    tracker.addValidation('M2:M100', ['Active', 'Pending', 'Closed', 'Cancelled', 'On Hold']);
    tracker.addValidation('E2:E100', ['Single Family', 'Condo', 'Townhouse', 'Multi-Family', 'Luxury Home', 'Land', 'Commercial']);
    tracker.addValidation('N2:N100', ['Referral', 'Website', 'Social Media', 'Open House', 'Cold Call', 'Past Client', 'Partner Referral', 'Walk-in', 'Other']);
    tracker.addValidation('S2:S100', ['Buy', 'Sell', 'Buy/Sell', 'Lease', 'Lease/Buy']);
    
    // Conditional formatting for status
    tracker.addConditionalFormat(2, 13, 100, 13, {
      type: 'text',
      condition: 'Closed',
      background: '#D1FAE5',
      fontColor: '#065F46'
    });
    
    // Set column widths
    tracker.setColumnWidth(1, 70); // Deal ID
    tracker.setColumnWidth(3, 150); // Client
    tracker.setColumnWidth(4, 200); // Address
    tracker.setColumnWidth(20, 250); // Notes
  },
  
  /**
   * Build Pipeline Sheet
   */
  buildPipeline: function(builder) {
    const pipeline = builder.sheet('Pipeline');
    
    // Headers
    const headers = [
      'Lead ID', 'Date', 'Name', 'Contact', 'Email', 'Property Interest', 
      'Budget', 'Stage', 'Source', 'Lead Score', 'Next Action', 'Follow-up Date', 
      'Agent', 'Probability', 'Est. Value', 'Days in Stage', 'Notes'
    ];
    
    pipeline.setRangeValues(1, 1, [headers]);
    pipeline.format(1, 1, 1, headers.length, {
      background: '#059669',
      fontColor: '#FFFFFF',
      fontWeight: 'bold',
      fontSize: 11,
      horizontalAlignment: 'center',
      border: true
    });
    
    // Sample pipeline data
    const pipelineData = [
      ['L001', '=TODAY()-14', 'David Miller', '555-0101', 'david@email.com', 'Single Family, 3BR', 450000, 'Qualified', 'Referral', '=IF(H2="Hot",90,IF(H2="Qualified",70,IF(H2="Active",50,30)))', 'Schedule showing', '=TODAY()+2', 'John Agent', '=J2/100', '=G2*N2', '=TODAY()-B2', 'Looking in good school districts'],
      ['L002', '=TODAY()-7', 'Emily Wilson', '555-0102', 'emily@email.com', 'Condo, Downtown', 275000, 'Active', 'Website', '=IF(H3="Hot",90,IF(H3="Qualified",70,IF(H3="Active",50,30)))', 'Send listings', '=TODAY()+1', 'Jane Agent', '=J3/100', '=G3*N3', '=TODAY()-B3', 'Pre-approved for $300k'],
      ['L003', '=TODAY()-3', 'James Taylor', '555-0103', 'james@email.com', 'Investment property', 600000, 'Hot', 'Open House', '=IF(H4="Hot",90,IF(H4="Qualified",70,IF(H4="Active",50,30)))', 'Prepare offer', '=TODAY()', 'John Agent', '=J4/100', '=G4*N4', '=TODAY()-B4', 'Cash buyer, ready to move'],
      ['L004', '=TODAY()-21', 'Linda Anderson', '555-0104', 'linda@email.com', 'Townhouse', 350000, 'Contacted', 'Social Media', '=IF(H5="Hot",90,IF(H5="Qualified",70,IF(H5="Active",50,30)))', 'Follow up call', '=TODAY()+5', 'Bob Agent', '=J5/100', '=G5*N5', '=TODAY()-B5', 'Needs to sell current home first'],
      ['L005', '=TODAY()-1', 'William Thomas', '555-0105', 'william@email.com', 'Luxury home', 1200000, 'New', 'Partner Referral', '=IF(H6="Hot",90,IF(H6="Qualified",70,IF(H6="Active",50,30)))', 'Initial qualification', '=TODAY()', 'Jane Agent', '=J6/100', '=G6*N6', '=TODAY()-B6', 'CEO of tech company, relocating']
    ];
    
    pipeline.setRangeValues(2, 1, pipelineData);
    
    // Format columns
    pipeline.format(2, 2, 100, 2, { numberFormat: 'mm/dd/yyyy' }); // Date
    pipeline.format(2, 7, 100, 7, { numberFormat: '$#,##0' }); // Budget
    pipeline.format(2, 12, 100, 12, { numberFormat: 'mm/dd/yyyy' }); // Follow-up
    pipeline.format(2, 14, 100, 14, { numberFormat: '0%' }); // Probability
    pipeline.format(2, 15, 100, 15, { numberFormat: '$#,##0' }); // Est. Value
    
    // Add data validations
    pipeline.addValidation('H2:H100', ['New', 'Contacted', 'Qualified', 'Active', 'Hot', 'Offer', 'Negotiating', 'Contract', 'Closed', 'Lost', 'Cold']);
    pipeline.addValidation('I2:I100', ['Referral', 'Website', 'Social Media', 'Open House', 'Cold Call', 'Email Campaign', 'Partner Referral', 'Walk-in', 'Other']);
    
    // Conditional formatting for stages
    pipeline.addConditionalFormat(2, 8, 100, 8, {
      type: 'text',
      condition: 'Hot',
      background: '#FEE2E2',
      fontColor: '#DC2626'
    });
    
    // Set column widths
    pipeline.setColumnWidth(3, 150); // Name
    pipeline.setColumnWidth(5, 180); // Email
    pipeline.setColumnWidth(6, 150); // Property Interest
    pipeline.setColumnWidth(17, 250); // Notes
  },
  
  /**
   * Build Clients Sheet
   */
  buildClients: function(builder) {
    const clients = builder.sheet('Clients');
    
    // Headers
    const headers = [
      'Client ID', 'Name', 'Email', 'Phone', 'Type', 'Status', 
      'First Contact', 'Last Activity', 'Total Value', 'Lifetime Commission', 
      'Properties Bought', 'Properties Sold', 'Referrals Given', 'VIP Score', 
      'Preferred Contact', 'Notes'
    ];
    
    clients.setRangeValues(1, 1, [headers]);
    clients.format(1, 1, 1, headers.length, {
      background: '#7C3AED',
      fontColor: '#FFFFFF',
      fontWeight: 'bold',
      fontSize: 11,
      horizontalAlignment: 'center',
      border: true
    });
    
    // Sample client data
    const clientData = [
      ['C001', 'John & Mary Smith', 'smiths@email.com', '555-0101', 'Buyer/Seller', 'Active', '=TODAY()-180', '=TODAY()-5', 875000, '=I2*0.025', 1, 1, 3, '=IF(J2>50000,100,IF(J2>20000,75,IF(J2>10000,50,25)))', 'Phone', 'Excellent clients, multiple referrals'],
      ['C002', 'Robert Johnson', 'rjohnson@email.com', '555-0102', 'Buyer', 'Active', '=TODAY()-90', '=TODAY()-10', 325000, '=I3*0.025', 1, 0, 0, '=IF(J3>50000,100,IF(J3>20000,75,IF(J3>10000,50,25)))', 'Email', 'First-time buyer, needs guidance'],
      ['C003', 'Sarah Williams', 'swilliams@email.com', '555-0103', 'Seller', 'Active', '=TODAY()-60', '=TODAY()-2', 580000, '=I4*0.025', 0, 1, 1, '=IF(J4>50000,100,IF(J4>20000,75,IF(J4>10000,50,25)))', 'Text', 'Downsizing, looking for condo'],
      ['C004', 'Michael Brown', 'mbrown@email.com', '555-0104', 'Investor', 'VIP', '=TODAY()-365', '=TODAY()-7', 2750000, '=I5*0.025', 4, 2, 5, '=IF(J5>50000,100,IF(J5>20000,75,IF(J5>10000,50,25)))', 'Phone', 'Real estate investor, multiple properties'],
      ['C005', 'Jennifer Davis', 'jdavis@email.com', '555-0105', 'Buyer', 'Prospect', '=TODAY()-14', '=TODAY()', 450000, '=I6*0.025', 0, 0, 0, '=IF(J6>50000,100,IF(J6>20000,75,IF(J6>10000,50,25)))', 'Email', 'Looking for luxury property']
    ];
    
    clients.setRangeValues(2, 1, clientData);
    
    // Format columns
    clients.format(2, 7, 100, 8, { numberFormat: 'mm/dd/yyyy' }); // Dates
    clients.format(2, 9, 100, 10, { numberFormat: '$#,##0' }); // Values
    
    // Add data validations
    clients.addValidation('E2:E100', ['Buyer', 'Seller', 'Buyer/Seller', 'Investor', 'Renter', 'Landlord']);
    clients.addValidation('F2:F100', ['Active', 'VIP', 'Prospect', 'Inactive', 'Past Client']);
    clients.addValidation('O2:O100', ['Phone', 'Email', 'Text', 'WhatsApp', 'In Person']);
    
    // Conditional formatting for VIP status
    clients.addConditionalFormat(2, 6, 100, 6, {
      type: 'text',
      condition: 'VIP',
      background: '#FEF3C7',
      fontColor: '#92400E'
    });
    
    // Set column widths
    clients.setColumnWidth(2, 150); // Name
    clients.setColumnWidth(3, 180); // Email
    clients.setColumnWidth(16, 250); // Notes
  },
  
  /**
   * Build Analytics Sheet - Advanced Performance Analytics
   */
  buildAnalytics: function(builder) {
    const analytics = builder.sheet('Analytics');
    
    // Title and description
    analytics.merge(1, 1, 1, 10);
    analytics.setValue(1, 1, '📊 Advanced Analytics & Performance Metrics');
    analytics.format(1, 1, 1, 10, {
      background: '#1E3A8A',
      fontColor: '#FFFFFF',
      fontSize: 14,
      fontWeight: 'bold',
      horizontalAlignment: 'center'
    });
    
    // ========================================
    // COMMISSION ANALYSIS SECTION
    // ========================================
    analytics.setValue(3, 1, 'Commission Analysis');
    analytics.format(3, 1, 3, 10, {
      background: '#E0E7FF',
      fontWeight: 'bold',
      fontSize: 12
    });
    
    // Commission breakdown headers
    const commissionHeaders = ['Period', 'Deals', 'Total Sales', 'Gross Commission', 'Net Commission', 'Avg Commission', 'Commission Rate', 'Growth %'];
    analytics.setRangeValues(4, 1, [commissionHeaders]);
    analytics.format(4, 1, 4, 8, {
      background: '#6B7280',
      fontColor: '#FFFFFF',
      fontWeight: 'bold'
    });
    
    // Monthly commission analysis
    const monthlyData = [
      ['This Month', '=COUNTIFS({{Commission Tracker}}!B:B,">="&EOMONTH(TODAY(),-1)+1,{{Commission Tracker}}!B:B,"<="&EOMONTH(TODAY(),0),{{Commission Tracker}}!M:M,"Closed")', '=SUMIFS({{Commission Tracker}}!F:F,{{Commission Tracker}}!B:B,">="&EOMONTH(TODAY(),-1)+1,{{Commission Tracker}}!B:B,"<="&EOMONTH(TODAY(),0),{{Commission Tracker}}!M:M,"Closed")', '=SUMIFS({{Commission Tracker}}!I:I,{{Commission Tracker}}!B:B,">="&EOMONTH(TODAY(),-1)+1,{{Commission Tracker}}!B:B,"<="&EOMONTH(TODAY(),0),{{Commission Tracker}}!M:M,"Closed")', '=SUMIFS({{Commission Tracker}}!K:K,{{Commission Tracker}}!B:B,">="&EOMONTH(TODAY(),-1)+1,{{Commission Tracker}}!B:B,"<="&EOMONTH(TODAY(),0),{{Commission Tracker}}!M:M,"Closed")', '=IFERROR(E5/B5,0)', '=IFERROR(D5/C5,0)', ''],
      ['Last Month', '=COUNTIFS({{Commission Tracker}}!B:B,">="&EOMONTH(TODAY(),-2)+1,{{Commission Tracker}}!B:B,"<="&EOMONTH(TODAY(),-1),{{Commission Tracker}}!M:M,"Closed")', '=SUMIFS({{Commission Tracker}}!F:F,{{Commission Tracker}}!B:B,">="&EOMONTH(TODAY(),-2)+1,{{Commission Tracker}}!B:B,"<="&EOMONTH(TODAY(),-1),{{Commission Tracker}}!M:M,"Closed")', '=SUMIFS({{Commission Tracker}}!I:I,{{Commission Tracker}}!B:B,">="&EOMONTH(TODAY(),-2)+1,{{Commission Tracker}}!B:B,"<="&EOMONTH(TODAY(),-1),{{Commission Tracker}}!M:M,"Closed")', '=SUMIFS({{Commission Tracker}}!K:K,{{Commission Tracker}}!B:B,">="&EOMONTH(TODAY(),-2)+1,{{Commission Tracker}}!B:B,"<="&EOMONTH(TODAY(),-1),{{Commission Tracker}}!M:M,"Closed")', '=IFERROR(E6/B6,0)', '=IFERROR(D6/C6,0)', '=IFERROR((E5-E6)/E6,0)'],
      ['This Quarter', '=COUNTIFS({{Commission Tracker}}!B:B,">="&DATE(YEAR(TODAY()),ROUNDUP(MONTH(TODAY())/3,0)*3-2,1),{{Commission Tracker}}!B:B,"<="&EOMONTH(DATE(YEAR(TODAY()),ROUNDUP(MONTH(TODAY())/3,0)*3,1),0),{{Commission Tracker}}!M:M,"Closed")', '=SUMIFS({{Commission Tracker}}!F:F,{{Commission Tracker}}!B:B,">="&DATE(YEAR(TODAY()),ROUNDUP(MONTH(TODAY())/3,0)*3-2,1),{{Commission Tracker}}!B:B,"<="&EOMONTH(DATE(YEAR(TODAY()),ROUNDUP(MONTH(TODAY())/3,0)*3,1),0),{{Commission Tracker}}!M:M,"Closed")', '=SUMIFS({{Commission Tracker}}!I:I,{{Commission Tracker}}!B:B,">="&DATE(YEAR(TODAY()),ROUNDUP(MONTH(TODAY())/3,0)*3-2,1),{{Commission Tracker}}!B:B,"<="&EOMONTH(DATE(YEAR(TODAY()),ROUNDUP(MONTH(TODAY())/3,0)*3,1),0),{{Commission Tracker}}!M:M,"Closed")', '=SUMIFS({{Commission Tracker}}!K:K,{{Commission Tracker}}!B:B,">="&DATE(YEAR(TODAY()),ROUNDUP(MONTH(TODAY())/3,0)*3-2,1),{{Commission Tracker}}!B:B,"<="&EOMONTH(DATE(YEAR(TODAY()),ROUNDUP(MONTH(TODAY())/3,0)*3,1),0),{{Commission Tracker}}!M:M,"Closed")', '=IFERROR(E7/B7,0)', '=IFERROR(D7/C7,0)', ''],
      ['YTD', '=COUNTIFS({{Commission Tracker}}!B:B,">="&DATE(YEAR(TODAY()),1,1),{{Commission Tracker}}!M:M,"Closed")', '=SUMIFS({{Commission Tracker}}!F:F,{{Commission Tracker}}!B:B,">="&DATE(YEAR(TODAY()),1,1),{{Commission Tracker}}!M:M,"Closed")', '=SUMIFS({{Commission Tracker}}!I:I,{{Commission Tracker}}!B:B,">="&DATE(YEAR(TODAY()),1,1),{{Commission Tracker}}!M:M,"Closed")', '=SUMIFS({{Commission Tracker}}!K:K,{{Commission Tracker}}!B:B,">="&DATE(YEAR(TODAY()),1,1),{{Commission Tracker}}!M:M,"Closed")', '=IFERROR(E8/B8,0)', '=IFERROR(D8/C8,0)', '']
    ];
    
    analytics.setRangeValues(5, 1, monthlyData);
    analytics.format(5, 3, 8, 6, { numberFormat: '$#,##0' });
    analytics.format(5, 7, 8, 7, { numberFormat: '0.00%' });
    analytics.format(5, 8, 8, 8, { numberFormat: '0.0%' });
    
    // ========================================
    // PIPELINE CONVERSION FUNNEL
    // ========================================
    analytics.setValue(10, 1, 'Pipeline Conversion Funnel');
    analytics.format(10, 1, 10, 10, {
      background: '#E0E7FF',
      fontWeight: 'bold',
      fontSize: 12
    });
    
    const funnelHeaders = ['Stage', 'Count', 'Value', 'Avg Days', 'Conv to Next', 'Conv to Close', 'Drop Rate', 'Velocity Score'];
    analytics.setRangeValues(11, 1, [funnelHeaders]);
    analytics.format(11, 1, 11, 8, {
      background: '#6B7280',
      fontColor: '#FFFFFF',
      fontWeight: 'bold'
    });
    
    const funnelData = [
      ['New', '=COUNTIF({{Pipeline}}!H:H,"New")', '=SUMIF({{Pipeline}}!H:H,"New",{{Pipeline}}!G:G)', '=AVERAGEIF({{Pipeline}}!H:H,"New",{{Pipeline}}!P:P)', '=IFERROR(B13/B12,0)', '=IFERROR(B18/B12,0)', '=1-F12', '=IF(D12>0,B12/D12,0)'],
      ['Contacted', '=COUNTIF({{Pipeline}}!H:H,"Contacted")', '=SUMIF({{Pipeline}}!H:H,"Contacted",{{Pipeline}}!G:G)', '=AVERAGEIF({{Pipeline}}!H:H,"Contacted",{{Pipeline}}!P:P)', '=IFERROR(B14/B13,0)', '=IFERROR(B18/B13,0)', '=1-F13', '=IF(D13>0,B13/D13,0)'],
      ['Qualified', '=COUNTIF({{Pipeline}}!H:H,"Qualified")', '=SUMIF({{Pipeline}}!H:H,"Qualified",{{Pipeline}}!G:G)', '=AVERAGEIF({{Pipeline}}!H:H,"Qualified",{{Pipeline}}!P:P)', '=IFERROR(B15/B14,0)', '=IFERROR(B18/B14,0)', '=1-F14', '=IF(D14>0,B14/D14,0)'],
      ['Active', '=COUNTIF({{Pipeline}}!H:H,"Active")', '=SUMIF({{Pipeline}}!H:H,"Active",{{Pipeline}}!G:G)', '=AVERAGEIF({{Pipeline}}!H:H,"Active",{{Pipeline}}!P:P)', '=IFERROR(B16/B15,0)', '=IFERROR(B18/B15,0)', '=1-F15', '=IF(D15>0,B15/D15,0)'],
      ['Hot', '=COUNTIF({{Pipeline}}!H:H,"Hot")', '=SUMIF({{Pipeline}}!H:H,"Hot",{{Pipeline}}!G:G)', '=AVERAGEIF({{Pipeline}}!H:H,"Hot",{{Pipeline}}!P:P)', '=IFERROR(B17/B16,0)', '=IFERROR(B18/B16,0)', '=1-F16', '=IF(D16>0,B16/D16,0)'],
      ['Contract', '=COUNTIF({{Pipeline}}!H:H,"Contract")', '=SUMIF({{Pipeline}}!H:H,"Contract",{{Pipeline}}!G:G)', '=AVERAGEIF({{Pipeline}}!H:H,"Contract",{{Pipeline}}!P:P)', '=IFERROR(B18/B17,0)', '=IFERROR(B18/B17,0)', '=1-F17', '=IF(D17>0,B17/D17,0)'],
      ['Closed', '=COUNTIF({{Pipeline}}!H:H,"Closed")', '=SUMIF({{Pipeline}}!H:H,"Closed",{{Pipeline}}!G:G)', '=AVERAGEIF({{Pipeline}}!H:H,"Closed",{{Pipeline}}!P:P)', '100%', '100%', '0%', '=IF(D18>0,B18/D18,0)']
    ];
    
    analytics.setRangeValues(12, 1, funnelData);
    analytics.format(12, 3, 18, 3, { numberFormat: '$#,##0' });
    analytics.format(12, 5, 18, 7, { numberFormat: '0%' });
    
    // ========================================
    // AGENT PERFORMANCE RANKINGS
    // ========================================
    analytics.setValue(20, 1, 'Agent Performance Rankings');
    analytics.format(20, 1, 20, 10, {
      background: '#E0E7FF',
      fontWeight: 'bold',
      fontSize: 12
    });
    
    analytics.setValue(21, 1, 'Top Performers by Net Commission (YTD)');
    analytics.setFormula(22, 1, 
      '=IFERROR(QUERY({{Commission Tracker}}!B:T,' +
      '"SELECT L, COUNT(A), SUM(F), SUM(K), AVG(F), AVG(P) ' +
      'WHERE B >= date \'"&TEXT(DATE(YEAR(TODAY()),1,1),"yyyy-mm-dd")&"\' ' +
      'AND M = \'Closed\' ' +
      'GROUP BY L ' +
      'ORDER BY SUM(K) DESC ' +
      'LABEL L \'Agent\', COUNT(A) \'Deals\', SUM(F) \'Volume\', SUM(K) \'Net Commission\', AVG(F) \'Avg Deal Size\', AVG(P) \'Avg Days to Close\'",1),"")'
    );
    
    // ========================================
    // LEAD SOURCE ROI ANALYSIS
    // ========================================
    analytics.setValue(30, 1, 'Lead Source ROI Analysis');
    analytics.format(30, 1, 30, 10, {
      background: '#E0E7FF',
      fontWeight: 'bold',
      fontSize: 12
    });
    
    const sourceHeaders = ['Source', 'Leads', 'Qualified', 'Closed', 'Conv Rate', 'Revenue', 'Cost', 'ROI', 'CPL', 'CAC'];
    analytics.setRangeValues(31, 1, [sourceHeaders]);
    analytics.format(31, 1, 31, 10, {
      background: '#6B7280',
      fontColor: '#FFFFFF',
      fontWeight: 'bold'
    });
    
    const sourceData = [
      ['Referral', '=COUNTIF({{Pipeline}}!I:I,"Referral")', '=COUNTIFS({{Pipeline}}!I:I,"Referral",{{Pipeline}}!H:H,"Qualified")+COUNTIFS({{Pipeline}}!I:I,"Referral",{{Pipeline}}!H:H,"Active")+COUNTIFS({{Pipeline}}!I:I,"Referral",{{Pipeline}}!H:H,"Hot")', '=COUNTIFS({{Commission Tracker}}!N:N,"Referral",{{Commission Tracker}}!M:M,"Closed")', '=IFERROR(D32/B32,0)', '=SUMIFS({{Commission Tracker}}!K:K,{{Commission Tracker}}!N:N,"Referral",{{Commission Tracker}}!M:M,"Closed")', '500', '=IFERROR((F32-G32)/G32,0)', '=IFERROR(G32/B32,0)', '=IFERROR(G32/D32,0)'],
      ['Website', '=COUNTIF({{Pipeline}}!I:I,"Website")', '=COUNTIFS({{Pipeline}}!I:I,"Website",{{Pipeline}}!H:H,"Qualified")+COUNTIFS({{Pipeline}}!I:I,"Website",{{Pipeline}}!H:H,"Active")+COUNTIFS({{Pipeline}}!I:I,"Website",{{Pipeline}}!H:H,"Hot")', '=COUNTIFS({{Commission Tracker}}!N:N,"Website",{{Commission Tracker}}!M:M,"Closed")', '=IFERROR(D33/B33,0)', '=SUMIFS({{Commission Tracker}}!K:K,{{Commission Tracker}}!N:N,"Website",{{Commission Tracker}}!M:M,"Closed")', '2000', '=IFERROR((F33-G33)/G33,0)', '=IFERROR(G33/B33,0)', '=IFERROR(G33/D33,0)'],
      ['Social Media', '=COUNTIF({{Pipeline}}!I:I,"Social Media")', '=COUNTIFS({{Pipeline}}!I:I,"Social Media",{{Pipeline}}!H:H,"Qualified")+COUNTIFS({{Pipeline}}!I:I,"Social Media",{{Pipeline}}!H:H,"Active")+COUNTIFS({{Pipeline}}!I:I,"Social Media",{{Pipeline}}!H:H,"Hot")', '=COUNTIFS({{Commission Tracker}}!N:N,"Social Media",{{Commission Tracker}}!M:M,"Closed")', '=IFERROR(D34/B34,0)', '=SUMIFS({{Commission Tracker}}!K:K,{{Commission Tracker}}!N:N,"Social Media",{{Commission Tracker}}!M:M,"Closed")', '1500', '=IFERROR((F34-G34)/G34,0)', '=IFERROR(G34/B34,0)', '=IFERROR(G34/D34,0)'],
      ['Open House', '=COUNTIF({{Pipeline}}!I:I,"Open House")', '=COUNTIFS({{Pipeline}}!I:I,"Open House",{{Pipeline}}!H:H,"Qualified")+COUNTIFS({{Pipeline}}!I:I,"Open House",{{Pipeline}}!H:H,"Active")+COUNTIFS({{Pipeline}}!I:I,"Open House",{{Pipeline}}!H:H,"Hot")', '=COUNTIFS({{Commission Tracker}}!N:N,"Open House",{{Commission Tracker}}!M:M,"Closed")', '=IFERROR(D35/B35,0)', '=SUMIFS({{Commission Tracker}}!K:K,{{Commission Tracker}}!N:N,"Open House",{{Commission Tracker}}!M:M,"Closed")', '800', '=IFERROR((F35-G35)/G35,0)', '=IFERROR(G35/B35,0)', '=IFERROR(G35/D35,0)'],
      ['Cold Call', '=COUNTIF({{Pipeline}}!I:I,"Cold Call")', '=COUNTIFS({{Pipeline}}!I:I,"Cold Call",{{Pipeline}}!H:H,"Qualified")+COUNTIFS({{Pipeline}}!I:I,"Cold Call",{{Pipeline}}!H:H,"Active")+COUNTIFS({{Pipeline}}!I:I,"Cold Call",{{Pipeline}}!H:H,"Hot")', '=COUNTIFS({{Commission Tracker}}!N:N,"Cold Call",{{Commission Tracker}}!M:M,"Closed")', '=IFERROR(D36/B36,0)', '=SUMIFS({{Commission Tracker}}!K:K,{{Commission Tracker}}!N:N,"Cold Call",{{Commission Tracker}}!M:M,"Closed")', '1000', '=IFERROR((F36-G36)/G36,0)', '=IFERROR(G36/B36,0)', '=IFERROR(G36/D36,0)']
    ];
    
    analytics.setRangeValues(32, 1, sourceData);
    analytics.format(32, 5, 36, 5, { numberFormat: '0%' });
    analytics.format(32, 6, 36, 7, { numberFormat: '$#,##0' });
    analytics.format(32, 8, 36, 8, { numberFormat: '0.0' });
    analytics.format(32, 9, 36, 10, { numberFormat: '$#,##0' });
  },
  
  /**
   * Build Market Insights Sheet
   */
  buildMarketInsights: function(builder) {
    const market = builder.sheet('Market Insights');
    
    // Title
    market.merge(1, 1, 1, 8);
    market.setValue(1, 1, '🏠 Local Market Intelligence & Trends');
    market.format(1, 1, 1, 8, {
      background: '#047857',
      fontColor: '#FFFFFF',
      fontSize: 14,
      fontWeight: 'bold',
      horizontalAlignment: 'center'
    });
    
    // Market Overview Section
    market.setValue(3, 1, 'Market Overview');
    market.format(3, 1, 3, 8, {
      background: '#D1FAE5',
      fontWeight: 'bold',
      fontSize: 12
    });
    
    const overviewHeaders = ['Metric', 'Current', 'Last Month', 'Last Year', 'Change YoY', 'Market Avg', 'vs Market', 'Trend'];
    market.setRangeValues(4, 1, [overviewHeaders]);
    market.format(4, 1, 4, 8, {
      background: '#6B7280',
      fontColor: '#FFFFFF',
      fontWeight: 'bold'
    });
    
    const marketData = [
      ['Median Sale Price', '425000', '420000', '395000', '=IFERROR((B5-D5)/D5,0)', '430000', '=IFERROR((B5-F5)/F5,0)', '↑'],
      ['Avg Days on Market', '28', '32', '45', '=IFERROR((B6-D6)/D6,0)', '35', '=IFERROR((B6-F6)/F6,0)', '↓'],
      ['Inventory (Months)', '2.3', '2.5', '3.8', '=IFERROR((B7-D7)/D7,0)', '3.0', '=IFERROR((B7-F7)/F7,0)', '↓'],
      ['List to Sale Ratio', '98.5%', '97.8%', '96.2%', '=B8-D8', '97.5%', '=B8-F8', '↑'],
      ['Active Listings', '342', '358', '412', '=IFERROR((B9-D9)/D9,0)', '380', '=IFERROR((B9-F9)/F9,0)', '↓'],
      ['Pending Sales', '89', '76', '65', '=IFERROR((B10-D10)/D10,0)', '72', '=IFERROR((B10-F10)/F10,0)', '↑'],
      ['Closed Sales', '124', '118', '98', '=IFERROR((B11-D11)/D11,0)', '105', '=IFERROR((B11-F11)/F11,0)', '↑']
    ];
    
    market.setRangeValues(5, 1, marketData);
    market.format(5, 2, 11, 2, { numberFormat: '$#,##0' });
    market.format(5, 3, 11, 3, { numberFormat: '$#,##0' });
    market.format(5, 4, 11, 4, { numberFormat: '$#,##0' });
    market.format(5, 5, 11, 5, { numberFormat: '0.0%' });
    market.format(5, 6, 11, 6, { numberFormat: '$#,##0' });
    market.format(5, 7, 11, 7, { numberFormat: '0.0%' });
    market.format(8, 2, 8, 7, { numberFormat: '0.0%' });
    
    // Property Type Analysis
    market.setValue(13, 1, 'Property Type Performance');
    market.format(13, 1, 13, 8, {
      background: '#D1FAE5',
      fontWeight: 'bold',
      fontSize: 12
    });
    
    const typeHeaders = ['Property Type', 'Listings', 'Avg Price', 'Avg DOM', 'Absorption Rate', 'Price/SqFt', 'YoY Growth', 'Market Share'];
    market.setRangeValues(14, 1, [typeHeaders]);
    market.format(14, 1, 14, 8, {
      background: '#6B7280',
      fontColor: '#FFFFFF',
      fontWeight: 'bold'
    });
    
    const typeData = [
      ['Single Family', '156', '485000', '25', '82%', '285', '8.5%', '45.6%'],
      ['Condo', '89', '325000', '22', '78%', '425', '12.3%', '26.0%'],
      ['Townhouse', '67', '395000', '30', '71%', '315', '6.8%', '19.6%'],
      ['Multi-Family', '18', '750000', '45', '55%', '195', '15.2%', '5.3%'],
      ['Luxury (>$1M)', '12', '1850000', '95', '42%', '485', '3.2%', '3.5%']
    ];
    
    market.setRangeValues(15, 1, typeData);
    market.format(15, 3, 19, 3, { numberFormat: '$#,##0' });
    market.format(15, 6, 19, 6, { numberFormat: '$#,##0' });
    
    // Competitive Analysis
    market.setValue(21, 1, 'Competitive Landscape');
    market.format(21, 1, 21, 8, {
      background: '#D1FAE5',
      fontWeight: 'bold',
      fontSize: 12
    });
    
    const compHeaders = ['Competitor', 'Market Share', 'Avg Sale Price', 'Avg Commission', 'Client Satisfaction', 'Days to Close', 'YTD Growth', 'Strength'];
    market.setRangeValues(22, 1, [compHeaders]);
    market.format(22, 1, 22, 8, {
      background: '#6B7280',
      fontColor: '#FFFFFF',
      fontWeight: 'bold'
    });
    
    const compData = [
      ['Your Performance', '=TEXT(COUNTIFS({{Commission Tracker}}!M:M,"Closed",{{Commission Tracker}}!B:B,">="&DATE(YEAR(TODAY()),1,1))/124,"0.0%")', '=IFERROR(AVERAGE({{Commission Tracker}}!F:F),0)', '=IFERROR(AVERAGE({{Commission Tracker}}!K:K),0)', '4.8/5', '=IFERROR(AVERAGE({{Commission Tracker}}!P:P),0)', '12%', 'Technology'],
      ['Realty One', '18.5%', '445000', '11250', '4.5/5', '32', '8%', 'Brand'],
      ['Premier Homes', '15.2%', '520000', '13500', '4.6/5', '28', '10%', 'Luxury'],
      ['City Realtors', '12.8%', '385000', '9625', '4.3/5', '35', '5%', 'Volume'],
      ['Dream Properties', '8.9%', '425000', '10625', '4.7/5', '30', '15%', 'Service']
    ];
    
    market.setRangeValues(23, 1, compData);
    market.format(23, 3, 27, 4, { numberFormat: '$#,##0' });
    
    // Price Trend Analysis
    market.setValue(29, 1, 'Price Trends by Area');
    market.format(29, 1, 29, 8, {
      background: '#D1FAE5',
      fontWeight: 'bold',
      fontSize: 12
    });
    
    const areaHeaders = ['Area/Neighborhood', 'Median Price', '3Mo Change', '6Mo Change', 'YoY Change', 'Forecast', 'Rating', 'Investment Score'];
    market.setRangeValues(30, 1, [areaHeaders]);
    market.format(30, 1, 30, 8, {
      background: '#6B7280',
      fontColor: '#FFFFFF',
      fontWeight: 'bold'
    });
    
    const areaData = [
      ['Downtown', '550000', '+3.2%', '+5.8%', '+12.3%', '↑ Strong', 'A+', '92'],
      ['Suburbs North', '425000', '+2.1%', '+4.2%', '+8.5%', '↑ Moderate', 'A', '85'],
      ['Suburbs South', '385000', '+1.8%', '+3.5%', '+7.2%', '→ Stable', 'B+', '78'],
      ['Waterfront', '850000', '+4.5%', '+8.2%', '+15.6%', '↑ Strong', 'A+', '95'],
      ['Historic District', '475000', '+2.8%', '+5.1%', '+9.8%', '↑ Moderate', 'A-', '82']
    ];
    
    market.setRangeValues(31, 1, areaData);
    market.format(31, 2, 35, 2, { numberFormat: '$#,##0' });
  },
  
  /**
   * Build Goals & Targets Sheet
   */
  buildGoalsTargets: function(builder) {
    const goals = builder.sheet('Goals & Targets');
    
    // Title
    goals.merge(1, 1, 1, 10);
    goals.setValue(1, 1, '🎯 Goals, Targets & Performance Tracking');
    goals.format(1, 1, 1, 10, {
      background: '#7C2D12',
      fontColor: '#FFFFFF',
      fontSize: 14,
      fontWeight: 'bold',
      horizontalAlignment: 'center'
    });
    
    // Annual Goals Section
    goals.setValue(3, 1, 'Annual Goals & Progress');
    goals.format(3, 1, 3, 10, {
      background: '#FED7AA',
      fontWeight: 'bold',
      fontSize: 12
    });
    
    const goalHeaders = ['Goal Category', 'Annual Target', 'Current Actual', 'Progress %', 'Remaining', 'Run Rate', 'Projected EOY', 'Status', 'Action Required'];
    goals.setRangeValues(4, 1, [goalHeaders]);
    goals.format(4, 1, 4, 9, {
      background: '#6B7280',
      fontColor: '#FFFFFF',
      fontWeight: 'bold'
    });
    
    const annualGoals = [
      ['Total Revenue', '150000', '=SUMIFS({{Commission Tracker}}!K:K,{{Commission Tracker}}!B:B,">="&DATE(YEAR(TODAY()),1,1),{{Commission Tracker}}!M:M,"Closed")', '=C5/B5', '=B5-C5', '=C5/MONTH(TODAY())*12', '=F5', '=IF(D5>=0.8,"🟢 On Track",IF(D5>=0.6,"🟡 Behind","🔴 At Risk"))', '=IF(D5<0.8,"Increase activity","Maintain pace")'],
      ['Deals Closed', '48', '=COUNTIFS({{Commission Tracker}}!B:B,">="&DATE(YEAR(TODAY()),1,1),{{Commission Tracker}}!M:M,"Closed")', '=C6/B6', '=B6-C6', '=C6/MONTH(TODAY())*12', '=F6', '=IF(D6>=0.8,"🟢 On Track",IF(D6>=0.6,"🟡 Behind","🔴 At Risk"))', '=IF(D6<0.8,"More prospecting","Stay focused")'],
      ['Sales Volume', '12000000', '=SUMIFS({{Commission Tracker}}!F:F,{{Commission Tracker}}!B:B,">="&DATE(YEAR(TODAY()),1,1),{{Commission Tracker}}!M:M,"Closed")', '=C7/B7', '=B7-C7', '=C7/MONTH(TODAY())*12', '=F7', '=IF(D7>=0.8,"🟢 On Track",IF(D7>=0.6,"🟡 Behind","🔴 At Risk"))', '=IF(D7<0.8,"Target higher value","Good progress")'],
      ['New Clients', '120', '=COUNTIFS({{Clients}}!G:G,">="&DATE(YEAR(TODAY()),1,1))', '=C8/B8', '=B8-C8', '=C8/MONTH(TODAY())*12', '=F8', '=IF(D8>=0.8,"🟢 On Track",IF(D8>=0.6,"🟡 Behind","🔴 At Risk"))', '=IF(D8<0.8,"Boost marketing","Keep networking")'],
      ['Referrals', '24', '=COUNTIFS({{Pipeline}}!I:I,"Referral",{{Pipeline}}!B:B,">="&DATE(YEAR(TODAY()),1,1))', '=C9/B9', '=B9-C9', '=C9/MONTH(TODAY())*12', '=F9', '=IF(D9>=0.8,"🟢 On Track",IF(D9>=0.6,"🟡 Behind","🔴 At Risk"))', '=IF(D9<0.8,"Ask for referrals","Great job")']
    ];
    
    goals.setRangeValues(5, 1, annualGoals);
    goals.format(5, 2, 9, 2, { numberFormat: '$#,##0' });
    goals.format(5, 3, 9, 3, { numberFormat: '$#,##0' });
    goals.format(5, 4, 9, 4, { numberFormat: '0%' });
    goals.format(5, 5, 9, 7, { numberFormat: '$#,##0' });
    goals.format(7, 2, 7, 7, { numberFormat: '$#,##0' });
    
    // Monthly Targets
    goals.setValue(11, 1, 'Monthly Targets & Tracking');
    goals.format(11, 1, 11, 10, {
      background: '#FED7AA',
      fontWeight: 'bold',
      fontSize: 12
    });
    
    const monthlyHeaders = ['Month', 'Revenue Target', 'Revenue Actual', 'Deals Target', 'Deals Actual', 'Pipeline Target', 'Pipeline Actual', 'Achievement %', 'Bonus Earned'];
    goals.setRangeValues(12, 1, [monthlyHeaders]);
    goals.format(12, 1, 12, 9, {
      background: '#6B7280',
      fontColor: '#FFFFFF',
      fontWeight: 'bold'
    });
    
    // Generate last 6 months of data
    const monthlyData = [];
    for (let i = 5; i >= 0; i--) {
      const monthFormula = `EOMONTH(TODAY(),-${i})`;
      monthlyData.push([
        `=TEXT(${monthFormula},"MMM yyyy")`,
        '12500',
        `=SUMIFS({{Commission Tracker}}!K:K,{{Commission Tracker}}!B:B,">="&EOMONTH(${monthFormula},-1)+1,{{Commission Tracker}}!B:B,"<="&${monthFormula},{{Commission Tracker}}!M:M,"Closed")`,
        '4',
        `=COUNTIFS({{Commission Tracker}}!B:B,">="&EOMONTH(${monthFormula},-1)+1,{{Commission Tracker}}!B:B,"<="&${monthFormula},{{Commission Tracker}}!M:M,"Closed")`,
        '20',
        `=COUNTIFS({{Pipeline}}!B:B,">="&EOMONTH(${monthFormula},-1)+1,{{Pipeline}}!B:B,"<="&${monthFormula})`,
        `=AVERAGE(C${13+i}/B${13+i},E${13+i}/D${13+i},G${13+i}/F${13+i})`,
        `=IF(H${13+i}>=1,500,IF(H${13+i}>=0.9,250,0))`
      ]);
    }
    
    goals.setRangeValues(13, 1, monthlyData);
    goals.format(13, 2, 18, 3, { numberFormat: '$#,##0' });
    goals.format(13, 8, 18, 8, { numberFormat: '0%' });
    goals.format(13, 9, 18, 9, { numberFormat: '$#,##0' });
    
    // Personal Development Goals
    goals.setValue(20, 1, 'Personal Development & Skills');
    goals.format(20, 1, 20, 10, {
      background: '#FED7AA',
      fontWeight: 'bold',
      fontSize: 12
    });
    
    const devHeaders = ['Skill Area', 'Current Level', 'Target Level', 'Progress', 'Activities Completed', 'Hours Invested', 'Certification', 'Next Step'];
    goals.setRangeValues(21, 1, [devHeaders]);
    goals.format(21, 1, 21, 8, {
      background: '#6B7280',
      fontColor: '#FFFFFF',
      fontWeight: 'bold'
    });
    
    const devData = [
      ['Negotiation', '7/10', '9/10', '78%', '3 workshops', '24', 'In Progress', 'Advanced course'],
      ['Digital Marketing', '6/10', '8/10', '75%', '2 courses', '16', 'Completed', 'Practice campaigns'],
      ['Market Analysis', '8/10', '9/10', '89%', '5 reports', '30', 'Completed', 'Maintain skills'],
      ['Client Relations', '9/10', '10/10', '90%', '4 trainings', '20', 'Completed', 'Mentor others'],
      ['Technology Tools', '5/10', '8/10', '63%', '1 certification', '12', 'Started', 'CRM mastery']
    ];
    
    goals.setRangeValues(22, 1, devData);
    
    // Action Items
    goals.setValue(28, 1, 'Priority Action Items');
    goals.format(28, 1, 28, 10, {
      background: '#FED7AA',
      fontWeight: 'bold',
      fontSize: 12
    });
    
    const actionHeaders = ['Priority', 'Action Item', 'Category', 'Due Date', 'Impact', 'Effort', 'Status', 'Owner', 'Notes'];
    goals.setRangeValues(29, 1, [actionHeaders]);
    goals.format(29, 1, 29, 9, {
      background: '#6B7280',
      fontColor: '#FFFFFF',
      fontWeight: 'bold'
    });
    
    const actionData = [
      ['1', 'Follow up with hot leads', 'Sales', '=TODAY()+1', 'High', 'Low', 'In Progress', 'Self', 'Call by noon'],
      ['2', 'Update website listings', 'Marketing', '=TODAY()+2', 'Medium', 'Medium', 'Pending', 'Self', 'Add new photos'],
      ['3', 'Client appreciation event', 'Relationship', '=TODAY()+30', 'High', 'High', 'Planning', 'Team', 'Q4 event'],
      ['4', 'Market report publication', 'Branding', '=TODAY()+7', 'Medium', 'Medium', 'Started', 'Self', 'Monthly newsletter'],
      ['5', 'Database cleanup', 'Operations', '=TODAY()+14', 'Low', 'High', 'Not Started', 'Assistant', 'Quarterly task']
    ];
    
    goals.setRangeValues(30, 1, actionData);
    goals.format(30, 4, 34, 4, { numberFormat: 'mm/dd/yyyy' });
  },
  
  /**
   * Build Reports Sheet
   */
  buildReports: function(builder) {
    const reports = builder.sheet('Reports');
    
    // Title
    reports.merge(1, 1, 1, 8);
    reports.setValue(1, 1, '📈 Automated Reports & Exports');
    reports.format(1, 1, 1, 8, {
      background: '#4C1D95',
      fontColor: '#FFFFFF',
      fontSize: 14,
      fontWeight: 'bold',
      horizontalAlignment: 'center'
    });
    
    // Executive Summary
    reports.setValue(3, 1, 'Executive Summary - ' + new Date().toLocaleDateString());
    reports.format(3, 1, 3, 8, {
      background: '#EDE9FE',
      fontWeight: 'bold',
      fontSize: 12
    });
    
    reports.setValue(5, 1, 'Key Metrics This Period');
    const summaryData = [
      ['Total Revenue YTD:', '=TEXT(SUMIFS({{Commission Tracker}}!K:K,{{Commission Tracker}}!B:B,">="&DATE(YEAR(TODAY()),1,1),{{Commission Tracker}}!M:M,"Closed"),"$#,##0")'],
      ['Deals Closed:', '=COUNTIFS({{Commission Tracker}}!B:B,">="&DATE(YEAR(TODAY()),1,1),{{Commission Tracker}}!M:M,"Closed")'],
      ['Active Pipeline:', '=COUNTIF({{Pipeline}}!H:H,"Active")+COUNTIF({{Pipeline}}!H:H,"Hot")+COUNTIF({{Pipeline}}!H:H,"Qualified")'],
      ['Conversion Rate:', '=TEXT(COUNTIF({{Pipeline}}!H:H,"Closed")/COUNTIF({{Pipeline}}!A:A,"<>"),"0.0%")'],
      ['Avg Days to Close:', '=ROUND(AVERAGE({{Commission Tracker}}!P:P),0)&" days"'],
      ['Top Lead Source:', '=INDEX({{Commission Tracker}}!N:N,MODE(MATCH({{Commission Tracker}}!N:N,{{Commission Tracker}}!N:N,0)))']
    ];
    
    for (let i = 0; i < summaryData.length; i++) {
      reports.setValue(6 + i, 1, summaryData[i][0]);
      reports.setFormula(6 + i, 2, summaryData[i][1]);
    }
    
    reports.format(6, 1, 11, 1, { fontWeight: 'bold' });
    
    // Monthly Performance Report
    reports.setValue(13, 1, 'Monthly Performance Report');
    reports.format(13, 1, 13, 8, {
      background: '#EDE9FE',
      fontWeight: 'bold',
      fontSize: 12
    });
    
    reports.setValue(14, 1, 'Last 12 Months Performance');
    reports.setFormula(15, 1,
      '=IFERROR(QUERY({{Commission Tracker}}!B:K,' +
      '"SELECT MONTH(B)+1, COUNT(A), SUM(F), SUM(I), SUM(K) ' +
      'WHERE B >= date \'"&TEXT(EDATE(TODAY(),-11),"yyyy-mm-dd")&"\' ' +
      'AND M = \'Closed\' ' +
      'GROUP BY MONTH(B)+1 ' +
      'ORDER BY MONTH(B)+1 ' +
      'LABEL MONTH(B)+1 \'Month\', COUNT(A) \'Deals\', SUM(F) \'Volume\', SUM(I) \'Gross Comm\', SUM(K) \'Net Comm\'",1),"")'
    );
    
    // Client Report
    reports.setValue(30, 1, 'Client Activity Report');
    reports.format(30, 1, 30, 8, {
      background: '#EDE9FE',
      fontWeight: 'bold',
      fontSize: 12
    });
    
    reports.setValue(31, 1, 'Top Clients by Lifetime Value');
    reports.setFormula(32, 1,
      '=IFERROR(QUERY({{Clients}}!B:J,' +
      '"SELECT B, E, F, I, J ' +
      'WHERE J > 0 ' +
      'ORDER BY J DESC ' +
      'LIMIT 10 ' +
      'LABEL B \'Client Name\', E \'Type\', F \'Status\', I \'Total Value\', J \'Commission\'",1),"")'
    );
    
    // Pipeline Report
    reports.setValue(45, 1, 'Pipeline Status Report');
    reports.format(45, 1, 45, 8, {
      background: '#EDE9FE',
      fontWeight: 'bold',
      fontSize: 12
    });
    
    reports.setValue(46, 1, 'Current Pipeline by Stage');
    reports.setFormula(47, 1,
      '=IFERROR(QUERY({{Pipeline}}!H:O,' +
      '"SELECT H, COUNT(A), SUM(G), AVG(G), AVG(J) ' +
      'WHERE H <> \'Closed\' AND H <> \'Lost\' ' +
      'GROUP BY H ' +
      'ORDER BY COUNT(A) DESC ' +
      'LABEL H \'Stage\', COUNT(A) \'Count\', SUM(G) \'Total Value\', AVG(G) \'Avg Value\', AVG(J) \'Avg Score\'",1),"")'
    );
    
    // Export Configuration
    reports.setValue(60, 1, 'Report Export Settings');
    reports.format(60, 1, 60, 8, {
      background: '#EDE9FE',
      fontWeight: 'bold',
      fontSize: 12
    });
    
    const exportSettings = [
      ['Report Type', 'Frequency', 'Recipients', 'Format', 'Last Sent', 'Next Send', 'Status'],
      ['Executive Summary', 'Weekly', 'manager@company.com', 'PDF', '=TODAY()-3', '=TODAY()+4', 'Active'],
      ['Monthly Performance', 'Monthly', 'team@company.com', 'Excel', '=EOMONTH(TODAY(),-1)', '=EOMONTH(TODAY(),0)', 'Active'],
      ['Commission Report', 'Bi-Weekly', 'accounting@company.com', 'CSV', '=TODAY()-7', '=TODAY()+7', 'Active'],
      ['Pipeline Report', 'Daily', 'self@company.com', 'Google Sheets', '=TODAY()-1', '=TODAY()+1', 'Active'],
      ['Client Report', 'Quarterly', 'marketing@company.com', 'PDF', '=EOMONTH(TODAY(),-3)', '=EOMONTH(TODAY(),0)', 'Paused']
    ];
    
    reports.setRangeValues(61, 1, exportSettings);
    reports.format(61, 1, 61, 7, {
      background: '#6B7280',
      fontColor: '#FFFFFF',
      fontWeight: 'bold'
    });
    reports.format(62, 5, 66, 6, { numberFormat: 'dd/mm/yyyy' });
    
    // Instructions
    reports.setValue(68, 1, 'Report Generation Instructions:');
    reports.setValue(69, 1, '1. Data automatically updates from source sheets');
    reports.setValue(70, 1, '2. Use File > Download to export in various formats');
    reports.setValue(71, 1, '3. Set up Apps Script triggers for automated sending');
    reports.setValue(72, 1, '4. Customize queries above for specific reporting needs');
    
    reports.format(68, 1, 68, 1, { fontWeight: 'bold' });
    reports.format(69, 1, 72, 8, { fontColor: '#6B7280', fontSize: 9 });
  },
  
  // ========================================
  // PROPERTY MANAGER TEMPLATE BUILDERS
  // ========================================
  
  buildPropertyDashboard: function(builder) {
    const dash = builder.sheet('Dashboard');
    
    // Header
    dash.merge(1, 1, 2, 12);
    dash.setValue(1, 1, 'Property Management Command Center');
    dash.format(1, 1, 2, 12, {
      background: '#059669',
      fontColor: '#FFFFFF',
      fontSize: 18,
      fontWeight: 'bold',
      horizontalAlignment: 'center'
    });
    
    // KPI Cards
    const kpis = [
      { row: 4, col: 1, title: 'Total Properties', formula: '=COUNTA({{Properties}}!A:A)-1', format: '#,##0', color: '#10B981' },
      { row: 4, col: 4, title: 'Occupied Units', formula: '=COUNTIF({{Tenants}}!H:H,"Active")', format: '#,##0', color: '#3B82F6' },
      { row: 4, col: 7, title: 'Monthly Revenue', formula: '=SUMIF({{Lease Tracker}}!G:G,"Active",{{Lease Tracker}}!D:D)', format: '$#,##0', color: '#8B5CF6' },
      { row: 4, col: 10, title: 'Occupancy Rate', formula: '=COUNTIF({{Tenants}}!H:H,"Active")/COUNTA({{Properties}}!A:A)-1', format: '0%', color: '#F59E0B' }
    ];
    
    kpis.forEach(kpi => {
      dash.merge(kpi.row, kpi.col, kpi.row + 2, kpi.col + 2);
      dash.setValue(kpi.row, kpi.col, kpi.title);
      dash.setFormula(kpi.row + 1, kpi.col, kpi.formula);
      dash.format(kpi.row, kpi.col, kpi.row + 2, kpi.col + 2, {
        background: kpi.color,
        fontColor: '#FFFFFF',
        fontWeight: 'bold',
        horizontalAlignment: 'center',
        numberFormat: kpi.format
      });
    });
  },
  
  buildPropertiesSheet: function(builder) {
    const props = builder.sheet('Properties');
    
    // Headers
    const headers = [
      'Property ID', 'Address', 'Type', 'Units', 'Purchase Date', 'Purchase Price',
      'Current Value', 'Monthly Revenue', 'Operating Expenses', 'NOI', 'Cap Rate', 'Status'
    ];
    
    props.setRangeValues(1, 1, [headers]);
    props.format(1, 1, 1, 12, {
      background: '#059669',
      fontColor: '#FFFFFF',
      fontWeight: 'bold'
    });
    
    // Sample data
    const sampleData = [
      ['PROP001', '123 Main St', 'Multi-Family', 4, '2020-01-15', 450000, 525000, 4800, 1200, '=H2-I2', '=J2*12/G2', 'Active'],
      ['PROP002', '456 Oak Ave', 'Single-Family', 1, '2021-06-20', 275000, 310000, 2200, 400, '=H3-I3', '=J3*12/G3', 'Active'],
      ['PROP003', '789 Elm Dr', 'Commercial', 6, '2019-03-10', 850000, 920000, 8500, 2100, '=H4-I4', '=J4*12/G4', 'Active']
    ];
    
    props.setRangeValues(2, 1, sampleData);
    props.format(2, 5, 4, 5, { numberFormat: 'dd/mm/yyyy' });
    props.format(2, 6, 4, 10, { numberFormat: '$#,##0' });
    props.format(2, 11, 4, 11, { numberFormat: '0.00%' });
  },
  
  buildTenantsSheet: function(builder) {
    const tenants = builder.sheet('Tenants');
    
    // Headers
    const headers = [
      'Tenant ID', 'Name', 'Property', 'Unit', 'Lease Start', 'Lease End',
      'Monthly Rent', 'Status', 'Security Deposit', 'Contact', 'Email', 'Notes'
    ];
    
    tenants.setRangeValues(1, 1, [headers]);
    tenants.format(1, 1, 1, 12, {
      background: '#059669',
      fontColor: '#FFFFFF',
      fontWeight: 'bold'
    });
  },
  
  buildLeaseTracker: function(builder) {
    const lease = builder.sheet('Lease Tracker');
    
    // Headers
    const headers = [
      'Lease ID', 'Property', 'Tenant', 'Monthly Rent', 'Start Date', 'End Date',
      'Status', 'Days Remaining', 'Total Paid', 'Outstanding', 'Last Payment', 'Next Due'
    ];
    
    lease.setRangeValues(1, 1, [headers]);
    lease.format(1, 1, 1, 12, {
      background: '#059669',
      fontColor: '#FFFFFF',
      fontWeight: 'bold'
    });
  },
  
  buildMaintenanceLog: function(builder) {
    const maint = builder.sheet('Maintenance');
    
    // Headers
    const headers = [
      'Request ID', 'Date', 'Property', 'Unit', 'Category', 'Priority',
      'Description', 'Status', 'Assigned To', 'Cost', 'Completed Date', 'Notes'
    ];
    
    maint.setRangeValues(1, 1, [headers]);
    maint.format(1, 1, 1, 12, {
      background: '#059669',
      fontColor: '#FFFFFF',
      fontWeight: 'bold'
    });
  },
  
  buildPropertyFinancials: function(builder) {
    const fin = builder.sheet('Financials');
    
    // Headers
    const headers = [
      'Month', 'Property', 'Rental Income', 'Other Income', 'Total Income',
      'Mortgage', 'Insurance', 'Taxes', 'Maintenance', 'Utilities', 'Management',
      'Total Expenses', 'Net Income', 'Cash Flow'
    ];
    
    fin.setRangeValues(1, 1, [headers]);
    fin.format(1, 1, 1, 14, {
      background: '#059669',
      fontColor: '#FFFFFF',
      fontWeight: 'bold'
    });
  },
  
  buildVendorsSheet: function(builder) {
    const vendors = builder.sheet('Vendors');
    
    // Headers
    const headers = [
      'Vendor ID', 'Company', 'Contact', 'Service Type', 'Phone', 'Email',
      'Rating', 'Total Spent', 'Last Service', 'Status', 'Notes'
    ];
    
    vendors.setRangeValues(1, 1, [headers]);
    vendors.format(1, 1, 1, 11, {
      background: '#059669',
      fontColor: '#FFFFFF',
      fontWeight: 'bold'
    });
  },
  
  buildPropertyReports: function(builder) {
    const reports = builder.sheet('Reports');
    
    // Report sections
    reports.setValue(1, 1, 'Property Performance Reports');
    reports.format(1, 1, 1, 8, {
      fontSize: 16,
      fontWeight: 'bold',
      background: '#059669',
      fontColor: '#FFFFFF'
    });
  },
  
  // ========================================
  // INVESTMENT ANALYZER TEMPLATE BUILDERS
  // ========================================
  
  buildInvestmentDashboard: function(builder) {
    const dash = builder.sheet('Dashboard');
    
    // Header
    dash.merge(1, 1, 2, 12);
    dash.setValue(1, 1, 'Real Estate Investment Analysis Suite');
    dash.format(1, 1, 2, 12, {
      background: '#047857',
      fontColor: '#FFFFFF',
      fontSize: 18,
      fontWeight: 'bold',
      horizontalAlignment: 'center'
    });
  },
  
  buildInvestmentProperties: function(builder) {
    const props = builder.sheet('Properties');
    
    // Investment property analysis headers
    const headers = [
      'Property ID', 'Address', 'Type', 'Purchase Price', 'Down Payment',
      'Loan Amount', 'Interest Rate', 'Monthly Payment', 'Rental Income',
      'Cash Flow', 'Cap Rate', 'ROI', 'IRR'
    ];
    
    props.setRangeValues(1, 1, [headers]);
    props.format(1, 1, 1, 13, {
      background: '#047857',
      fontColor: '#FFFFFF',
      fontWeight: 'bold'
    });
  },
  
  buildCashFlowAnalysis: function(builder) {
    const cash = builder.sheet('Cash Flow Analysis');
    
    // Monthly cash flow projection
    cash.setValue(1, 1, 'Monthly Cash Flow Projection');
    cash.format(1, 1, 1, 12, {
      fontSize: 14,
      fontWeight: 'bold',
      background: '#047857',
      fontColor: '#FFFFFF'
    });
  },
  
  buildROICalculator: function(builder) {
    const roi = builder.sheet('ROI Calculator');
    
    // ROI calculation framework
    roi.setValue(1, 1, 'Return on Investment Calculator');
    roi.format(1, 1, 1, 8, {
      fontSize: 14,
      fontWeight: 'bold',
      background: '#047857',
      fontColor: '#FFFFFF'
    });
  },
  
  buildMarketComps: function(builder) {
    const comps = builder.sheet('Market Comps');
    
    // Comparable properties analysis
    const headers = [
      'Address', 'Sale Date', 'Sale Price', 'Sq Ft', 'Price/Sq Ft',
      'Bedrooms', 'Bathrooms', 'Year Built', 'Days on Market', 'Notes'
    ];
    
    comps.setRangeValues(1, 1, [headers]);
    comps.format(1, 1, 1, 10, {
      background: '#047857',
      fontColor: '#FFFFFF',
      fontWeight: 'bold'
    });
  },
  
  buildFinancingSheet: function(builder) {
    const finance = builder.sheet('Financing');
    
    // Loan comparison and analysis
    finance.setValue(1, 1, 'Financing Options Analysis');
    finance.format(1, 1, 1, 10, {
      fontSize: 14,
      fontWeight: 'bold',
      background: '#047857',
      fontColor: '#FFFFFF'
    });
  },
  
  buildTaxAnalysis: function(builder) {
    const tax = builder.sheet('Tax Analysis');
    
    // Tax benefits and deductions
    tax.setValue(1, 1, 'Tax Analysis & Deductions');
    tax.format(1, 1, 1, 8, {
      fontSize: 14,
      fontWeight: 'bold',
      background: '#047857',
      fontColor: '#FFFFFF'
    });
  },
  
  buildInvestmentReports: function(builder) {
    const reports = builder.sheet('Reports');
    
    // Investment performance reports
    reports.setValue(1, 1, 'Investment Performance Reports');
    reports.format(1, 1, 1, 10, {
      fontSize: 16,
      fontWeight: 'bold',
      background: '#047857',
      fontColor: '#FFFFFF'
    });
  },
  
  // ========================================
  // LEAD PIPELINE TEMPLATE BUILDERS
  // ========================================
  
  buildLeadDashboard: function(builder) {
    const dash = builder.sheet('Dashboard');
    
    // Header
    dash.merge(1, 1, 2, 12);
    dash.setValue(1, 1, 'Lead Generation & Conversion Center');
    dash.format(1, 1, 2, 12, {
      background: '#065F46',
      fontColor: '#FFFFFF',
      fontSize: 18,
      fontWeight: 'bold',
      horizontalAlignment: 'center'
    });
  },
  
  buildLeadPipelineSheet: function(builder) {
    const pipeline = builder.sheet('Lead Pipeline');
    
    // Lead tracking headers
    const headers = [
      'Lead ID', 'Date', 'Name', 'Source', 'Type', 'Budget', 'Location',
      'Stage', 'Score', 'Agent', 'Next Action', 'Last Contact', 'Notes'
    ];
    
    pipeline.setRangeValues(1, 1, [headers]);
    pipeline.format(1, 1, 1, 13, {
      background: '#065F46',
      fontColor: '#FFFFFF',
      fontWeight: 'bold'
    });
  },
  
  buildContactLog: function(builder) {
    const contact = builder.sheet('Contact Log');
    
    // Contact history tracking
    const headers = [
      'Date', 'Lead ID', 'Lead Name', 'Contact Type', 'Duration',
      'Agent', 'Outcome', 'Follow-up Date', 'Notes'
    ];
    
    contact.setRangeValues(1, 1, [headers]);
    contact.format(1, 1, 1, 9, {
      background: '#065F46',
      fontColor: '#FFFFFF',
      fontWeight: 'bold'
    });
  },
  
  buildLeadSources: function(builder) {
    const sources = builder.sheet('Lead Sources');
    
    // Lead source tracking and ROI
    const headers = [
      'Source', 'Leads Generated', 'Qualified Leads', 'Conversions',
      'Revenue', 'Cost', 'ROI', 'Conversion Rate', 'Avg Deal Size'
    ];
    
    sources.setRangeValues(1, 1, [headers]);
    sources.format(1, 1, 1, 9, {
      background: '#065F46',
      fontColor: '#FFFFFF',
      fontWeight: 'bold'
    });
  },
  
  buildFollowUps: function(builder) {
    const followup = builder.sheet('Follow-ups');
    
    // Follow-up schedule
    const headers = [
      'Date', 'Lead ID', 'Lead Name', 'Type', 'Priority',
      'Agent', 'Status', 'Completed', 'Outcome', 'Notes'
    ];
    
    followup.setRangeValues(1, 1, [headers]);
    followup.format(1, 1, 1, 10, {
      background: '#065F46',
      fontColor: '#FFFFFF',
      fontWeight: 'bold'
    });
  },
  
  buildConversionAnalytics: function(builder) {
    const analytics = builder.sheet('Conversion Analytics');
    
    // Conversion funnel analysis
    analytics.setValue(1, 1, 'Lead Conversion Analytics');
    analytics.format(1, 1, 1, 10, {
      fontSize: 14,
      fontWeight: 'bold',
      background: '#065F46',
      fontColor: '#FFFFFF'
    });
  },
  
  buildMarketingROI: function(builder) {
    const roi = builder.sheet('Marketing ROI');
    
    // Marketing campaign ROI analysis
    const headers = [
      'Campaign', 'Start Date', 'End Date', 'Budget', 'Spent',
      'Leads', 'Conversions', 'Revenue', 'ROI', 'CPL', 'CAC'
    ];
    
    roi.setRangeValues(1, 1, [headers]);
    roi.format(1, 1, 1, 11, {
      background: '#065F46',
      fontColor: '#FFFFFF',
      fontWeight: 'bold'
    });
  },
  
  buildLeadReports: function(builder) {
    const reports = builder.sheet('Reports');
    
    // Lead generation reports
    reports.setValue(1, 1, 'Lead Generation & Conversion Reports');
    reports.format(1, 1, 1, 10, {
      fontSize: 16,
      fontWeight: 'bold',
      background: '#065F46',
      fontColor: '#FFFFFF'
    });
  },
  
  /**
   * ============================================
   * PROPERTY MANAGEMENT SUITE - COMPREHENSIVE
   * ============================================
   */
  
  /**
   * Build Property Management Dashboard
   */
  buildPropertyDashboard: function(builder) {
    const dash = builder.sheet('Dashboard');
    
    // Header
    dash.merge(1, 1, 2, 14);
    dash.setValue(1, 1, '🏢 Property Management Command Center');
    dash.format(1, 1, 2, 14, {
      background: '#1E40AF',
      fontColor: '#FFFFFF',
      fontSize: 20,
      fontWeight: 'bold',
      horizontalAlignment: 'center',
      verticalAlignment: 'middle'
    });
    
    // Date and portfolio summary
    dash.setValue(3, 1, 'Portfolio Overview');
    dash.setFormula(3, 3, '=TEXT(NOW(),"mmm dd, yyyy h:mm AM/PM")');
    dash.setValue(3, 6, 'Total Properties:');
    dash.setFormula(3, 7, '=COUNTA({{Properties}}!A:A)-1');
    dash.setValue(3, 9, 'Total Units:');
    dash.setFormula(3, 10, '=SUM({{Properties}}!D:D)');
    dash.setValue(3, 12, 'Occupancy Rate:');
    dash.setFormula(3, 13, '=IFERROR(SUMIF({{Tenants}}!H:H,"Active",{{Tenants}}!G:G)/SUM({{Properties}}!D:D),0)');
    dash.format(3, 13, 3, 13, { numberFormat: '0.0%' });
    
    // Primary KPI Cards (Row 5-9)
    const kpiCards = [
      {
        row: 5, col: 1, width: 3,
        title: 'Monthly Revenue',
        formula: '=SUMIFS({{Financials}}!E:E,{{Financials}}!B:B,">="&EOMONTH(TODAY(),-1)+1,{{Financials}}!B:B,"<="&EOMONTH(TODAY(),0),{{Financials}}!C:C,"Rent")',
        format: '$#,##0',
        color: '#10B981',
        icon: '💰',
        subtitle: '=COUNTIFS({{Financials}}!B:B,">="&EOMONTH(TODAY(),-1)+1,{{Financials}}!B:B,"<="&EOMONTH(TODAY(),0),{{Financials}}!C:C,"Rent")&" payments received"'
      },
      {
        row: 5, col: 4, width: 3,
        title: 'Outstanding Balance',
        formula: '=SUMIF({{Tenants}}!H:H,"Active",{{Tenants}}!K:K)',
        format: '$#,##0',
        color: '#EF4444',
        icon: '⚠️',
        subtitle: '=COUNTIF({{Tenants}}!K:K,">0")&" tenants with balance"'
      },
      {
        row: 5, col: 7, width: 3,
        title: 'Net Operating Income',
        formula: '=SUMIFS({{Financials}}!E:E,{{Financials}}!B:B,">="&EOMONTH(TODAY(),-1)+1,{{Financials}}!B:B,"<="&EOMONTH(TODAY(),0),{{Financials}}!C:C,"Rent")-SUMIFS({{Financials}}!E:E,{{Financials}}!B:B,">="&EOMONTH(TODAY(),-1)+1,{{Financials}}!B:B,"<="&EOMONTH(TODAY(),0),{{Financials}}!C:C,"Expense")',
        format: '$#,##0',
        color: '#3B82F6',
        icon: '📊',
        subtitle: '="NOI Margin: "&TEXT(IFERROR((SUMIFS({{Financials}}!E:E,{{Financials}}!B:B,">="&EOMONTH(TODAY(),-1)+1,{{Financials}}!B:B,"<="&EOMONTH(TODAY(),0),{{Financials}}!C:C,"Rent")-SUMIFS({{Financials}}!E:E,{{Financials}}!B:B,">="&EOMONTH(TODAY(),-1)+1,{{Financials}}!B:B,"<="&EOMONTH(TODAY(),0),{{Financials}}!C:C,"Expense"))/SUMIFS({{Financials}}!E:E,{{Financials}}!B:B,">="&EOMONTH(TODAY(),-1)+1,{{Financials}}!B:B,"<="&EOMONTH(TODAY(),0),{{Financials}}!C:C,"Rent"),0),"0.0%")'
      },
      {
        row: 5, col: 10, width: 3,
        title: 'Maintenance Requests',
        formula: '=COUNTIFS({{Maintenance}}!G:G,"Open")+COUNTIFS({{Maintenance}}!G:G,"In Progress")',
        format: '#,##0',
        color: '#F59E0B',
        icon: '🔧',
        subtitle: '=COUNTIF({{Maintenance}}!F:F,"Emergency")&" emergency"'
      }
    ];
    
    // Create KPI cards
    kpiCards.forEach(kpi => {
      dash.merge(kpi.row, kpi.col, kpi.row + 3, kpi.col + kpi.width - 1);
      dash.format(kpi.row, kpi.col, kpi.row + 3, kpi.col + kpi.width - 1, {
        background: '#FFFFFF',
        border: true
      });
      
      dash.setValue(kpi.row, kpi.col, kpi.icon + ' ' + kpi.title);
      dash.format(kpi.row, kpi.col, kpi.row, kpi.col + kpi.width - 1, {
        fontSize: 11,
        fontColor: '#6B7280',
        fontWeight: 'bold'
      });
      
      dash.setFormula(kpi.row + 1, kpi.col, kpi.formula);
      dash.format(kpi.row + 1, kpi.col, kpi.row + 1, kpi.col + kpi.width - 1, {
        fontSize: 20,
        fontWeight: 'bold',
        fontColor: kpi.color,
        numberFormat: kpi.format,
        horizontalAlignment: 'center'
      });
      
      dash.setFormula(kpi.row + 2, kpi.col, kpi.subtitle);
      dash.format(kpi.row + 2, kpi.col, kpi.row + 2, kpi.col + kpi.width - 1, {
        fontSize: 10,
        fontColor: '#9CA3AF',
        horizontalAlignment: 'center'
      });
    });
    
    // Property Performance Section (Row 11-20)
    dash.setValue(11, 1, 'PROPERTY PERFORMANCE');
    dash.format(11, 1, 11, 6, {
      fontSize: 12,
      fontWeight: 'bold',
      background: '#F3F4F6',
      border: true
    });
    
    // Headers for property metrics
    const propertyHeaders = ['Property', 'Units', 'Occupied', 'Revenue', 'Expenses', 'NOI'];
    propertyHeaders.forEach((header, idx) => {
      dash.setValue(12, idx + 1, header);
    });
    dash.format(12, 1, 12, 6, {
      fontWeight: 'bold',
      background: '#E5E7EB',
      border: true
    });
    
    // Property data formulas
    for (let i = 0; i < 7; i++) {
      const row = 13 + i;
      dash.setFormula(row, 1, `=IFERROR(INDEX({{Properties}}!B:B,${i + 2}),"")`);
      dash.setFormula(row, 2, `=IFERROR(INDEX({{Properties}}!D:D,${i + 2}),"")`);
      dash.setFormula(row, 3, `=IFERROR(COUNTIFS({{Tenants}}!D:D,INDEX({{Properties}}!A:A,${i + 2}),{{Tenants}}!H:H,"Active"),0)`);
      dash.setFormula(row, 4, `=IFERROR(SUMIFS({{Financials}}!E:E,{{Financials}}!D:D,INDEX({{Properties}}!A:A,${i + 2}),{{Financials}}!C:C,"Rent",{{Financials}}!B:B,">="&EOMONTH(TODAY(),-1)+1),0)`);
      dash.setFormula(row, 5, `=IFERROR(SUMIFS({{Financials}}!E:E,{{Financials}}!D:D,INDEX({{Properties}}!A:A,${i + 2}),{{Financials}}!C:C,"Expense",{{Financials}}!B:B,">="&EOMONTH(TODAY(),-1)+1),0)`);
      dash.setFormula(row, 6, `=D${row}-E${row}`);
      
      dash.format(row, 4, row, 6, { numberFormat: '$#,##0' });
    }
    dash.format(13, 1, 19, 6, { border: true });
    
    // Tenant Status Summary (Row 11-20, Columns 8-14)
    dash.setValue(11, 8, 'TENANT STATUS');
    dash.format(11, 8, 11, 14, {
      fontSize: 12,
      fontWeight: 'bold',
      background: '#F3F4F6',
      border: true
    });
    
    // Tenant metrics
    const tenantMetrics = [
      { label: 'Active Leases', formula: '=COUNTIF({{Tenants}}!H:H,"Active")' },
      { label: 'Expiring (30 days)', formula: '=COUNTIFS({{Tenants}}!F:F,"<="&TODAY()+30,{{Tenants}}!F:F,">="&TODAY(),{{Tenants}}!H:H,"Active")' },
      { label: 'Month-to-Month', formula: '=COUNTIF({{Tenants}}!I:I,"Month-to-Month")' },
      { label: 'Applications Pending', formula: '=COUNTIF({{Tenants}}!H:H,"Application")' },
      { label: 'Avg Rent', formula: '=IFERROR(AVERAGE({{Tenants}}!G:G),0)' },
      { label: 'Collection Rate', formula: '=IFERROR(1-SUMIF({{Tenants}}!K:K,">0")/SUMIF({{Tenants}}!H:H,"Active",{{Tenants}}!G:G),1)' }
    ];
    
    tenantMetrics.forEach((metric, idx) => {
      const row = 13 + idx;
      dash.setValue(row, 8, metric.label);
      dash.setFormula(row, 10, metric.formula);
      dash.format(row, 8, row, 10, { border: true });
      
      if (metric.label.includes('Avg Rent')) {
        dash.format(row, 10, row, 10, { numberFormat: '$#,##0' });
      } else if (metric.label.includes('Rate')) {
        dash.format(row, 10, row, 10, { numberFormat: '0.0%' });
      }
    });
    
    // Maintenance Summary (Row 22-30)
    dash.setValue(22, 1, 'MAINTENANCE & WORK ORDERS');
    dash.format(22, 1, 22, 8, {
      fontSize: 12,
      fontWeight: 'bold',
      background: '#F3F4F6',
      border: true
    });
    
    const maintenanceHeaders = ['Status', 'Count', 'Avg Days', 'Cost'];
    maintenanceHeaders.forEach((header, idx) => {
      dash.setValue(23, idx * 2 + 1, header);
    });
    dash.format(23, 1, 23, 8, {
      fontWeight: 'bold',
      background: '#E5E7EB',
      border: true
    });
    
    // Maintenance status rows
    const statuses = ['Open', 'In Progress', 'Completed', 'Emergency'];
    statuses.forEach((status, idx) => {
      const row = 24 + idx;
      dash.setValue(row, 1, status);
      dash.setFormula(row, 3, `=COUNTIF({{Maintenance}}!G:G,"${status}")`);
      dash.setFormula(row, 5, `=IFERROR(AVERAGE(IF({{Maintenance}}!G:G="${status}",TODAY()-{{Maintenance}}!C:C)),0)`);
      dash.setFormula(row, 7, `=SUMIF({{Maintenance}}!G:G,"${status}",{{Maintenance}}!I:I)`);
      
      dash.format(row, 7, row, 7, { numberFormat: '$#,##0' });
    });
    dash.format(24, 1, 27, 8, { border: true });
    
    // Financial Summary (Row 22-30, Columns 10-14)
    dash.setValue(22, 10, 'FINANCIAL SUMMARY');
    dash.format(22, 10, 22, 14, {
      fontSize: 12,
      fontWeight: 'bold',
      background: '#F3F4F6',
      border: true
    });
    
    const financialMetrics = [
      { label: 'YTD Revenue', formula: '=SUMIFS({{Financials}}!E:E,{{Financials}}!C:C,"Rent",{{Financials}}!B:B,">="&DATE(YEAR(TODAY()),1,1))' },
      { label: 'YTD Expenses', formula: '=SUMIFS({{Financials}}!E:E,{{Financials}}!C:C,"Expense",{{Financials}}!B:B,">="&DATE(YEAR(TODAY()),1,1))' },
      { label: 'YTD NOI', formula: '=SUMIFS({{Financials}}!E:E,{{Financials}}!C:C,"Rent",{{Financials}}!B:B,">="&DATE(YEAR(TODAY()),1,1))-SUMIFS({{Financials}}!E:E,{{Financials}}!C:C,"Expense",{{Financials}}!B:B,">="&DATE(YEAR(TODAY()),1,1))' },
      { label: 'Avg Monthly Revenue', formula: '=IFERROR(SUMIFS({{Financials}}!E:E,{{Financials}}!C:C,"Rent",{{Financials}}!B:B,">="&DATE(YEAR(TODAY()),1,1))/MONTH(TODAY()),0)' },
      { label: 'Cap Rate', formula: '=IFERROR((SUMIFS({{Financials}}!E:E,{{Financials}}!C:C,"Rent",{{Financials}}!B:B,">="&DATE(YEAR(TODAY()),1,1))-SUMIFS({{Financials}}!E:E,{{Financials}}!C:C,"Expense",{{Financials}}!B:B,">="&DATE(YEAR(TODAY()),1,1)))*12/SUM({{Properties}}!F:F),0)' }
    ];
    
    financialMetrics.forEach((metric, idx) => {
      const row = 24 + idx;
      dash.setValue(row, 10, metric.label);
      dash.setFormula(row, 13, metric.formula);
      dash.format(row, 10, row, 14, { border: true });
      
      if (metric.label.includes('Cap Rate')) {
        dash.format(row, 13, row, 13, { numberFormat: '0.0%' });
      } else {
        dash.format(row, 13, row, 13, { numberFormat: '$#,##0' });
      }
    });
    
    // Set column widths for optimal layout
    for (let col = 1; col <= 14; col++) {
      dash.setColumnWidth(col, col % 2 === 0 ? 80 : 100);
    }
    dash.setRowHeight(1, 35);
    dash.setRowHeight(2, 35);
  },
  
  /**
   * Build Properties Sheet
   */
  buildPropertiesSheet: function(builder) {
    const properties = builder.sheet('Properties');
    
    // Header
    properties.merge(1, 1, 1, 16);
    properties.setValue(1, 1, '🏢 Property Portfolio');
    properties.format(1, 1, 1, 16, {
      fontSize: 16,
      fontWeight: 'bold',
      background: '#1E40AF',
      fontColor: '#FFFFFF',
      horizontalAlignment: 'center'
    });
    
    // Column headers
    const headers = [
      'Property ID', 'Property Name', 'Address', 'Total Units', 'Type',
      'Purchase Price', 'Current Value', 'Purchase Date', 'Square Feet',
      'Year Built', 'Parking Spaces', 'Amenities', 'Property Manager',
      'Management Fee', 'Tax ID', 'Insurance Policy', 'Notes'
    ];
    
    headers.forEach((header, idx) => {
      properties.setValue(2, idx + 1, header);
    });
    properties.format(2, 1, 2, headers.length, {
      fontWeight: 'bold',
      background: '#E5E7EB',
      border: true
    });
    
    // Sample data
    const sampleProperties = [
      ['PROP001', 'Sunset Apartments', '123 Main St, City, ST 12345', 24, 'Multi-Family',
       1500000, 1800000, '2020-03-15', 18000, 1985, 30, 'Pool, Gym, Laundry',
       'John Smith', 0.08, 'TX-123456', 'POL-789012', 'Recently renovated'],
      ['PROP002', 'Oak Grove Condos', '456 Oak Ave, Town, ST 67890', 12, 'Condo',
       900000, 1100000, '2019-07-22', 12000, 1992, 18, 'Gated, Playground',
       'Jane Doe', 0.07, 'TX-234567', 'POL-890123', 'HOA managed'],
      ['PROP003', 'Downtown Lofts', '789 Urban Blvd, Metro, ST 54321', 8, 'Loft',
       1200000, 1400000, '2021-11-08', 10000, 2005, 10, 'Rooftop, Concierge',
       'Mike Johnson', 0.09, 'TX-345678', 'POL-901234', 'Premium location']
    ];
    
    sampleProperties.forEach((property, idx) => {
      property.forEach((value, col) => {
        if (col === 5 || col === 6) { // Purchase Price and Current Value
          properties.setValue(idx + 3, col + 1, value);
          properties.format(idx + 3, col + 1, idx + 3, col + 1, { numberFormat: '$#,##0' });
        } else if (col === 7) { // Purchase Date
          properties.setValue(idx + 3, col + 1, value);
          properties.format(idx + 3, col + 1, idx + 3, col + 1, { numberFormat: 'yyyy-mm-dd' });
        } else if (col === 13) { // Management Fee
          properties.setValue(idx + 3, col + 1, value);
          properties.format(idx + 3, col + 1, idx + 3, col + 1, { numberFormat: '0.0%' });
        } else {
          properties.setValue(idx + 3, col + 1, value);
        }
      });
    });
    
    // Data validation
    properties.addValidation('E3:E100', ['Single-Family', 'Multi-Family', 'Condo', 'Townhouse', 'Loft', 'Commercial', 'Mixed-Use']);
    
    // Conditional formatting for value changes
    properties.addConditionalFormat(3, 7, 100, 7, {
      type: 'gradient',
      minColor: '#FEE2E2',
      midColor: '#FFFFFF',
      maxColor: '#D1FAE5',
      minValue: '-0.1',
      midValue: '0',
      maxValue: '0.1'
    });
    
    // Set column widths
    properties.setColumnWidth(1, 100);
    properties.setColumnWidth(2, 150);
    properties.setColumnWidth(3, 250);
    properties.setColumnWidth(12, 200);
    properties.setColumnWidth(17, 300);
  },
  
  /**
   * Build Tenants Sheet
   */
  buildTenantsSheet: function(builder) {
    const tenants = builder.sheet('Tenants');
    
    // Header
    tenants.merge(1, 1, 1, 18);
    tenants.setValue(1, 1, '👥 Tenant Management');
    tenants.format(1, 1, 1, 18, {
      fontSize: 16,
      fontWeight: 'bold',
      background: '#1E40AF',
      fontColor: '#FFFFFF',
      horizontalAlignment: 'center'
    });
    
    // Column headers
    const headers = [
      'Tenant ID', 'Name', 'Email', 'Phone', 'Property ID', 'Unit #',
      'Lease Start', 'Lease End', 'Monthly Rent', 'Status', 'Lease Type',
      'Security Deposit', 'Balance Due', 'Last Payment', 'Payment Method',
      'Emergency Contact', 'Move-in Inspection', 'Background Check', 'Notes'
    ];
    
    headers.forEach((header, idx) => {
      tenants.setValue(2, idx + 1, header);
    });
    tenants.format(2, 1, 2, headers.length, {
      fontWeight: 'bold',
      background: '#E5E7EB',
      border: true
    });
    
    // Sample data
    const sampleTenants = [
      ['TEN001', 'Alice Johnson', 'alice@email.com', '555-0101', 'PROP001', '101',
       '2024-01-01', '2024-12-31', 1500, 'Active', '12-Month', 1500, 0,
       '2024-01-01', 'Auto-Pay', 'Bob Johnson - 555-0102', '2024-01-01', 'Passed', 'Excellent tenant'],
      ['TEN002', 'Bob Smith', 'bob@email.com', '555-0201', 'PROP001', '102',
       '2023-06-01', '2024-05-31', 1400, 'Active', '12-Month', 1400, 150,
       '2024-01-15', 'Check', 'Jane Smith - 555-0202', '2023-06-01', 'Passed', 'Late payment in Dec'],
      ['TEN003', 'Carol Davis', 'carol@email.com', '555-0301', 'PROP002', '201',
       '2024-02-01', '2025-01-31', 1800, 'Active', '12-Month', 1800, 0,
       '2024-02-01', 'Online', 'David Davis - 555-0302', '2024-02-01', 'Passed', 'New tenant']
    ];
    
    sampleTenants.forEach((tenant, idx) => {
      tenant.forEach((value, col) => {
        if (col === 6 || col === 7 || col === 13 || col === 16) { // Date fields
          tenants.setValue(idx + 3, col + 1, value);
          tenants.format(idx + 3, col + 1, idx + 3, col + 1, { numberFormat: 'yyyy-mm-dd' });
        } else if (col === 8 || col === 11 || col === 12) { // Money fields
          tenants.setValue(idx + 3, col + 1, value);
          tenants.format(idx + 3, col + 1, idx + 3, col + 1, { numberFormat: '$#,##0' });
        } else {
          tenants.setValue(idx + 3, col + 1, value);
        }
      });
    });
    
    // Data validation
    tenants.addValidation('J3:J100', ['Active', 'Inactive', 'Application', 'Eviction', 'Notice Given']);
    tenants.addValidation('K3:K100', ['12-Month', '6-Month', 'Month-to-Month', 'Short-Term']);
    tenants.addValidation('O3:O100', ['Auto-Pay', 'Check', 'Cash', 'Online', 'Wire Transfer']);
    tenants.addValidation('R3:R100', ['Passed', 'Failed', 'Pending', 'N/A']);
    
    // Conditional formatting for status
    tenants.addConditionalFormat(3, 10, 100, 10, {
      type: 'cell',
      condition: 'TEXT_EQUALS',
      value: 'Active',
      format: { background: '#D1FAE5' }
    });
    
    tenants.addConditionalFormat(3, 13, 100, 13, {
      type: 'cell',
      condition: 'NUMBER_GREATER_THAN',
      value: 0,
      format: { background: '#FEE2E2', fontColor: '#DC2626' }
    });
    
    // Set column widths
    tenants.setColumnWidth(2, 150);
    tenants.setColumnWidth(3, 180);
    tenants.setColumnWidth(16, 180);
    tenants.setColumnWidth(19, 250);
  },
  
  /**
   * Build Lease Tracker Sheet
   */
  buildLeaseTracker: function(builder) {
    const leases = builder.sheet('Lease Tracker');
    
    // Header
    leases.merge(1, 1, 1, 14);
    leases.setValue(1, 1, '📋 Lease Tracking & Renewals');
    leases.format(1, 1, 1, 14, {
      fontSize: 16,
      fontWeight: 'bold',
      background: '#1E40AF',
      fontColor: '#FFFFFF',
      horizontalAlignment: 'center'
    });
    
    // Section: Expiring Leases
    leases.setValue(3, 1, 'LEASES EXPIRING SOON');
    leases.format(3, 1, 3, 14, {
      fontSize: 12,
      fontWeight: 'bold',
      background: '#FEF3C7'
    });
    
    // Headers for expiring leases
    const expiringHeaders = [
      'Tenant Name', 'Property', 'Unit', 'Lease End', 'Days Until Expiry',
      'Current Rent', 'Market Rent', 'Proposed Increase', 'Renewal Status',
      'Notice Sent', 'Response', 'New Lease Start', 'New Lease End', 'Notes'
    ];
    
    expiringHeaders.forEach((header, idx) => {
      leases.setValue(4, idx + 1, header);
    });
    leases.format(4, 1, 4, expiringHeaders.length, {
      fontWeight: 'bold',
      background: '#E5E7EB',
      border: true
    });
    
    // Formulas to pull expiring leases
    for (let i = 0; i < 10; i++) {
      const row = 5 + i;
      leases.setFormula(row, 1, `=IFERROR(INDEX({{Tenants}}!B:B,SMALL(IF(({{Tenants}}!F:F<=TODAY()+60)*({{Tenants}}!F:F>=TODAY())*({{Tenants}}!H:H="Active"),ROW({{Tenants}}!B:B)),${i + 1})),"")`);
      leases.setFormula(row, 2, `=IFERROR(VLOOKUP(A${row},{{Tenants}}!B:D,3,FALSE),"")`);
      leases.setFormula(row, 3, `=IFERROR(VLOOKUP(A${row},{{Tenants}}!B:E,4,FALSE),"")`);
      leases.setFormula(row, 4, `=IFERROR(VLOOKUP(A${row},{{Tenants}}!B:F,5,FALSE),"")`);
      leases.setFormula(row, 5, `=IF(D${row}="","",D${row}-TODAY())`);
      leases.setFormula(row, 6, `=IFERROR(VLOOKUP(A${row},{{Tenants}}!B:G,6,FALSE),"")`);
      leases.setFormula(row, 7, `=IF(F${row}="","",F${row}*1.03)`); // 3% market increase
      leases.setFormula(row, 8, `=IF(F${row}="","",G${row}-F${row})`);
      
      leases.format(row, 4, row, 4, { numberFormat: 'yyyy-mm-dd' });
      leases.format(row, 6, row, 8, { numberFormat: '$#,##0' });
    }
    
    // Data validation for renewal status
    leases.addValidation('I5:I14', ['Pending', 'Offered', 'Accepted', 'Declined', 'Negotiating']);
    leases.addValidation('K5:K14', ['Yes', 'No', 'Pending']);
    
    // Conditional formatting
    leases.addConditionalFormat(5, 5, 14, 5, {
      type: 'gradient',
      minColor: '#DC2626',
      midColor: '#F59E0B',
      maxColor: '#10B981',
      minValue: '0',
      midValue: '30',
      maxValue: '60'
    });
  },
  
  /**
   * Build Maintenance Log Sheet
   */
  buildMaintenanceLog: function(builder) {
    const maintenance = builder.sheet('Maintenance');
    
    // Header
    maintenance.merge(1, 1, 1, 14);
    maintenance.setValue(1, 1, '🔧 Maintenance & Work Orders');
    maintenance.format(1, 1, 1, 14, {
      fontSize: 16,
      fontWeight: 'bold',
      background: '#1E40AF',
      fontColor: '#FFFFFF',
      horizontalAlignment: 'center'
    });
    
    // Column headers
    const headers = [
      'Work Order #', 'Property', 'Unit', 'Date Reported', 'Tenant',
      'Category', 'Priority', 'Status', 'Description', 'Assigned To',
      'Cost Estimate', 'Actual Cost', 'Date Completed', 'Resolution Notes'
    ];
    
    headers.forEach((header, idx) => {
      maintenance.setValue(2, idx + 1, header);
    });
    maintenance.format(2, 1, 2, headers.length, {
      fontWeight: 'bold',
      background: '#E5E7EB',
      border: true
    });
    
    // Sample maintenance records
    const sampleMaintenance = [
      ['WO001', 'PROP001', '101', '2024-01-15', 'Alice Johnson',
       'Plumbing', 'Emergency', 'Completed', 'Leaking faucet in kitchen',
       'ABC Plumbing', 150, 175, '2024-01-15', 'Replaced faucet cartridge'],
      ['WO002', 'PROP001', '102', '2024-01-18', 'Bob Smith',
       'HVAC', 'Normal', 'In Progress', 'AC not cooling properly',
       'Cool Air Services', 300, 0, '', 'Scheduled for service'],
      ['WO003', 'PROP002', '201', '2024-01-20', 'Carol Davis',
       'Electrical', 'Urgent', 'Open', 'Outlet not working in bedroom',
       'PowerPro Electric', 200, 0, '', 'Awaiting parts']
    ];
    
    sampleMaintenance.forEach((record, idx) => {
      record.forEach((value, col) => {
        if (col === 3 || col === 12) { // Date fields
          maintenance.setValue(idx + 3, col + 1, value);
          maintenance.format(idx + 3, col + 1, idx + 3, col + 1, { numberFormat: 'yyyy-mm-dd' });
        } else if (col === 10 || col === 11) { // Cost fields
          maintenance.setValue(idx + 3, col + 1, value);
          maintenance.format(idx + 3, col + 1, idx + 3, col + 1, { numberFormat: '$#,##0' });
        } else {
          maintenance.setValue(idx + 3, col + 1, value);
        }
      });
    });
    
    // Data validation
    maintenance.addValidation('F3:F100', ['Plumbing', 'Electrical', 'HVAC', 'Appliance', 'Structural', 'Pest Control', 'Landscaping', 'Other']);
    maintenance.addValidation('G3:G100', ['Emergency', 'Urgent', 'Normal', 'Low']);
    maintenance.addValidation('H3:H100', ['Open', 'In Progress', 'On Hold', 'Completed', 'Cancelled']);
    
    // Conditional formatting for priority
    maintenance.addConditionalFormat(3, 7, 100, 7, {
      type: 'cell',
      condition: 'TEXT_EQUALS',
      value: 'Emergency',
      format: { background: '#FEE2E2', fontColor: '#DC2626', fontWeight: 'bold' }
    });
    
    // Set column widths
    maintenance.setColumnWidth(9, 300);
    maintenance.setColumnWidth(14, 300);
  },
  
  /**
   * Build Property Financials Sheet
   */
  buildPropertyFinancials: function(builder) {
    const financials = builder.sheet('Financials');
    
    // Header
    financials.merge(1, 1, 1, 12);
    financials.setValue(1, 1, '💰 Financial Transactions');
    financials.format(1, 1, 1, 12, {
      fontSize: 16,
      fontWeight: 'bold',
      background: '#1E40AF',
      fontColor: '#FFFFFF',
      horizontalAlignment: 'center'
    });
    
    // Column headers
    const headers = [
      'Transaction ID', 'Date', 'Type', 'Category', 'Property', 'Unit',
      'Amount', 'Payment Method', 'Reference', 'Tenant/Vendor', 'Description', 'Reconciled'
    ];
    
    headers.forEach((header, idx) => {
      financials.setValue(2, idx + 1, header);
    });
    financials.format(2, 1, 2, headers.length, {
      fontWeight: 'bold',
      background: '#E5E7EB',
      border: true
    });
    
    // Sample transactions
    const sampleTransactions = [
      ['FIN001', '2024-01-01', 'Income', 'Rent', 'PROP001', '101',
       1500, 'Auto-Pay', 'REF-001', 'Alice Johnson', 'January rent', 'Yes'],
      ['FIN002', '2024-01-01', 'Income', 'Rent', 'PROP001', '102',
       1400, 'Check', 'CHK-1234', 'Bob Smith', 'January rent', 'Yes'],
      ['FIN003', '2024-01-05', 'Expense', 'Maintenance', 'PROP001', '101',
       175, 'Credit Card', 'INV-5678', 'ABC Plumbing', 'Faucet repair', 'Yes'],
      ['FIN004', '2024-01-10', 'Expense', 'Insurance', 'PROP001', '',
       450, 'Wire Transfer', 'POL-789012', 'State Insurance', 'Monthly premium', 'Yes'],
      ['FIN005', '2024-01-15', 'Expense', 'Property Tax', 'PROP001', '',
       1250, 'Online', 'TX-123456', 'County Tax Office', 'Q1 property tax', 'No']
    ];
    
    sampleTransactions.forEach((transaction, idx) => {
      transaction.forEach((value, col) => {
        if (col === 1) { // Date field
          financials.setValue(idx + 3, col + 1, value);
          financials.format(idx + 3, col + 1, idx + 3, col + 1, { numberFormat: 'yyyy-mm-dd' });
        } else if (col === 6) { // Amount field
          financials.setValue(idx + 3, col + 1, value);
          financials.format(idx + 3, col + 1, idx + 3, col + 1, { numberFormat: '$#,##0.00' });
        } else {
          financials.setValue(idx + 3, col + 1, value);
        }
      });
    });
    
    // Data validation
    financials.addValidation('C3:C100', ['Income', 'Expense']);
    financials.addValidation('D3:D100', ['Rent', 'Late Fee', 'Application Fee', 'Maintenance', 'Repairs', 'Insurance', 'Property Tax', 'HOA', 'Utilities', 'Management Fee', 'Marketing', 'Legal', 'Other']);
    financials.addValidation('H3:H100', ['Auto-Pay', 'Check', 'Cash', 'Online', 'Wire Transfer', 'Credit Card']);
    financials.addValidation('L3:L100', ['Yes', 'No', 'Pending']);
    
    // Conditional formatting
    financials.addConditionalFormat(3, 3, 100, 3, {
      type: 'cell',
      condition: 'TEXT_EQUALS',
      value: 'Income',
      format: { fontColor: '#10B981' }
    });
    
    financials.addConditionalFormat(3, 3, 100, 3, {
      type: 'cell',
      condition: 'TEXT_EQUALS',
      value: 'Expense',
      format: { fontColor: '#DC2626' }
    });
    
    // Set column widths
    financials.setColumnWidth(11, 300);
  },
  
  /**
   * Build Vendors Sheet
   */
  buildVendorsSheet: function(builder) {
    const vendors = builder.sheet('Vendors');
    
    // Header
    vendors.merge(1, 1, 1, 12);
    vendors.setValue(1, 1, '🛠️ Vendor & Contractor Management');
    vendors.format(1, 1, 1, 12, {
      fontSize: 16,
      fontWeight: 'bold',
      background: '#1E40AF',
      fontColor: '#FFFFFF',
      horizontalAlignment: 'center'
    });
    
    // Column headers
    const headers = [
      'Vendor ID', 'Company Name', 'Contact Name', 'Phone', 'Email',
      'Service Type', 'License #', 'Insurance Exp', 'Rating', 'Hourly Rate',
      'Last Used', 'Notes'
    ];
    
    headers.forEach((header, idx) => {
      vendors.setValue(2, idx + 1, header);
    });
    vendors.format(2, 1, 2, headers.length, {
      fontWeight: 'bold',
      background: '#E5E7EB',
      border: true
    });
    
    // Sample vendors
    const sampleVendors = [
      ['VEN001', 'ABC Plumbing', 'John Doe', '555-1111', 'john@abcplumbing.com',
       'Plumbing', 'PLB-12345', '2024-12-31', 4.5, 85, '2024-01-15', 'Reliable, 24/7 service'],
      ['VEN002', 'Cool Air Services', 'Jane Smith', '555-2222', 'jane@coolair.com',
       'HVAC', 'HVAC-67890', '2024-10-15', 4.8, 95, '2024-01-18', 'Certified for all brands'],
      ['VEN003', 'PowerPro Electric', 'Mike Johnson', '555-3333', 'mike@powerpro.com',
       'Electrical', 'ELE-54321', '2025-03-20', 4.2, 90, '2024-01-20', 'Licensed master electrician']
    ];
    
    sampleVendors.forEach((vendor, idx) => {
      vendor.forEach((value, col) => {
        if (col === 7 || col === 10) { // Date fields
          vendors.setValue(idx + 3, col + 1, value);
          vendors.format(idx + 3, col + 1, idx + 3, col + 1, { numberFormat: 'yyyy-mm-dd' });
        } else if (col === 9) { // Hourly rate
          vendors.setValue(idx + 3, col + 1, value);
          vendors.format(idx + 3, col + 1, idx + 3, col + 1, { numberFormat: '$#,##0' });
        } else if (col === 8) { // Rating
          vendors.setValue(idx + 3, col + 1, value);
          vendors.format(idx + 3, col + 1, idx + 3, col + 1, { numberFormat: '0.0' });
        } else {
          vendors.setValue(idx + 3, col + 1, value);
        }
      });
    });
    
    // Data validation
    vendors.addValidation('F3:F100', ['Plumbing', 'Electrical', 'HVAC', 'Landscaping', 'Pest Control', 'Cleaning', 'General Maintenance', 'Roofing', 'Painting', 'Flooring']);
    
    // Conditional formatting for ratings
    vendors.addConditionalFormat(3, 9, 100, 9, {
      type: 'gradient',
      minColor: '#FEE2E2',
      midColor: '#FEF3C7',
      maxColor: '#D1FAE5',
      minValue: '1',
      midValue: '3',
      maxValue: '5'
    });
    
    // Set column widths
    vendors.setColumnWidth(2, 150);
    vendors.setColumnWidth(5, 200);
    vendors.setColumnWidth(12, 300);
  },
  
  /**
   * Build Property Reports Sheet
   */
  buildPropertyReports: function(builder) {
    const reports = builder.sheet('Reports');
    
    // Header
    reports.merge(1, 1, 1, 12);
    reports.setValue(1, 1, '📊 Property Management Reports & Analytics');
    reports.format(1, 1, 1, 12, {
      fontSize: 16,
      fontWeight: 'bold',
      background: '#1E40AF',
      fontColor: '#FFFFFF',
      horizontalAlignment: 'center'
    });
    
    // Monthly Summary Section
    reports.setValue(3, 1, 'MONTHLY PERFORMANCE SUMMARY');
    reports.format(3, 1, 3, 12, {
      fontSize: 12,
      fontWeight: 'bold',
      background: '#F3F4F6'
    });
    
    // Monthly metrics
    const monthlyMetrics = [
      { label: 'Total Revenue', formula: '=SUMIFS({{Financials}}!G:G,{{Financials}}!B:B,">="&EOMONTH(TODAY(),-1)+1,{{Financials}}!B:B,"<="&EOMONTH(TODAY(),0),{{Financials}}!C:C,"Income")' },
      { label: 'Total Expenses', formula: '=SUMIFS({{Financials}}!G:G,{{Financials}}!B:B,">="&EOMONTH(TODAY(),-1)+1,{{Financials}}!B:B,"<="&EOMONTH(TODAY(),0),{{Financials}}!C:C,"Expense")' },
      { label: 'Net Operating Income', formula: '=B5-B6' },
      { label: 'Occupancy Rate', formula: '=COUNTIF({{Tenants}}!J:J,"Active")/SUM({{Properties}}!D:D)' },
      { label: 'Collection Rate', formula: '=1-SUMIF({{Tenants}}!M:M,">0")/SUMIF({{Tenants}}!J:J,"Active",{{Tenants}}!I:I)' },
      { label: 'Avg Days to Fill', formula: '=30' },
      { label: 'Maintenance Costs', formula: '=SUMIFS({{Financials}}!G:G,{{Financials}}!B:B,">="&EOMONTH(TODAY(),-1)+1,{{Financials}}!B:B,"<="&EOMONTH(TODAY(),0),{{Financials}}!D:D,"Maintenance")' },
      { label: 'Work Orders Completed', formula: '=COUNTIFS({{Maintenance}}!H:H,"Completed",{{Maintenance}}!M:M,">="&EOMONTH(TODAY(),-1)+1)' }
    ];
    
    reports.setValue(4, 1, 'Metric');
    reports.setValue(4, 2, 'This Month');
    reports.setValue(4, 3, 'Last Month');
    reports.setValue(4, 4, 'Change');
    reports.format(4, 1, 4, 4, {
      fontWeight: 'bold',
      background: '#E5E7EB',
      border: true
    });
    
    monthlyMetrics.forEach((metric, idx) => {
      const row = 5 + idx;
      reports.setValue(row, 1, metric.label);
      reports.setFormula(row, 2, metric.formula);
      
      // Last month formula (simplified)
      reports.setFormula(row, 3, metric.formula.replace(/TODAY\(\)/g, 'TODAY()-30'));
      
      // Change calculation
      reports.setFormula(row, 4, `=IFERROR((B${row}-C${row})/C${row},0)`);
      
      // Formatting
      if (metric.label.includes('Rate')) {
        reports.format(row, 2, row, 3, { numberFormat: '0.0%' });
        reports.format(row, 4, row, 4, { numberFormat: '0.0%' });
      } else if (!metric.label.includes('Days') && !metric.label.includes('Orders')) {
        reports.format(row, 2, row, 3, { numberFormat: '$#,##0' });
        reports.format(row, 4, row, 4, { numberFormat: '0.0%' });
      }
      
      reports.format(row, 1, row, 4, { border: true });
    });
    
    // Property Performance Comparison
    reports.setValue(15, 1, 'PROPERTY PERFORMANCE COMPARISON');
    reports.format(15, 1, 15, 8, {
      fontSize: 12,
      fontWeight: 'bold',
      background: '#F3F4F6'
    });
    
    const propertyCompHeaders = ['Property', 'Units', 'Occupied', 'Vacancy %', 'Revenue', 'Expenses', 'NOI', 'NOI/Unit'];
    propertyCompHeaders.forEach((header, idx) => {
      reports.setValue(16, idx + 1, header);
    });
    reports.format(16, 1, 16, 8, {
      fontWeight: 'bold',
      background: '#E5E7EB',
      border: true
    });
    
    // Property comparison formulas
    for (let i = 0; i < 5; i++) {
      const row = 17 + i;
      reports.setFormula(row, 1, `=IFERROR(INDEX({{Properties}}!B:B,${i + 2}),"")`);
      reports.setFormula(row, 2, `=IFERROR(INDEX({{Properties}}!D:D,${i + 2}),"")`);
      reports.setFormula(row, 3, `=IFERROR(COUNTIFS({{Tenants}}!E:E,INDEX({{Properties}}!A:A,${i + 2}),{{Tenants}}!J:J,"Active"),0)`);
      reports.setFormula(row, 4, `=IFERROR(1-C${row}/B${row},0)`);
      reports.setFormula(row, 5, `=IFERROR(SUMIFS({{Financials}}!G:G,{{Financials}}!E:E,INDEX({{Properties}}!A:A,${i + 2}),{{Financials}}!C:C,"Income",{{Financials}}!B:B,">="&EOMONTH(TODAY(),-1)+1),0)`);
      reports.setFormula(row, 6, `=IFERROR(SUMIFS({{Financials}}!G:G,{{Financials}}!E:E,INDEX({{Properties}}!A:A,${i + 2}),{{Financials}}!C:C,"Expense",{{Financials}}!B:B,">="&EOMONTH(TODAY(),-1)+1),0)`);
      reports.setFormula(row, 7, `=E${row}-F${row}`);
      reports.setFormula(row, 8, `=IFERROR(G${row}/B${row},0)`);
      
      reports.format(row, 4, row, 4, { numberFormat: '0.0%' });
      reports.format(row, 5, row, 8, { numberFormat: '$#,##0' });
      reports.format(row, 1, row, 8, { border: true });
    }
    
    // Set column widths
    for (let col = 1; col <= 12; col++) {
      reports.setColumnWidth(col, 100);
    }
  },
  
  /**
   * ============================================
   * INVESTMENT ANALYZER - COMPREHENSIVE
   * ============================================
   */
  
  /**
   * Build Investment Dashboard
   */
  buildInvestmentDashboard: function(builder) {
    const dash = builder.sheet('Dashboard');
    
    // Header
    dash.merge(1, 1, 2, 14);
    dash.setValue(1, 1, '📈 Real Estate Investment Analyzer');
    dash.format(1, 1, 2, 14, {
      background: '#047857',
      fontColor: '#FFFFFF',
      fontSize: 20,
      fontWeight: 'bold',
      horizontalAlignment: 'center',
      verticalAlignment: 'middle'
    });
    
    // Investment Summary
    dash.setValue(3, 1, 'Portfolio Summary');
    dash.setFormula(3, 3, '=TEXT(NOW(),"mmm dd, yyyy h:mm AM/PM")');
    dash.setValue(3, 6, 'Total Properties:');
    dash.setFormula(3, 7, '=COUNTA({{Properties}}!A:A)-1');
    dash.setValue(3, 9, 'Portfolio Value:');
    dash.setFormula(3, 10, '=SUM({{Properties}}!F:F)');
    dash.format(3, 10, 3, 10, { numberFormat: '$#,##0' });
    dash.setValue(3, 12, 'Total Equity:');
    dash.setFormula(3, 13, '=SUM({{Properties}}!F:F)-SUM({{Financing}}!E:E)');
    dash.format(3, 13, 3, 13, { numberFormat: '$#,##0' });
    
    // Primary Investment KPIs (Row 5-9)
    const kpiCards = [
      {
        row: 5, col: 1, width: 3,
        title: 'Total Investment',
        formula: '=SUM({{Properties}}!E:E)+SUM({{Properties}}!K:K)',
        format: '$#,##0',
        color: '#059669',
        icon: '💵',
        subtitle: '="Avg per property: "&TEXT(SUM({{Properties}}!E:E)/(COUNTA({{Properties}}!A:A)-1),"$#,##0")'
      },
      {
        row: 5, col: 4, width: 3,
        title: 'Annual Cash Flow',
        formula: '=SUM({{Cash Flow Analysis}}!M14:M25)',
        format: '$#,##0',
        color: '#3B82F6',
        icon: '💰',
        subtitle: '="Monthly avg: "&TEXT(SUM({{Cash Flow Analysis}}!M14:M25)/12,"$#,##0")'
      },
      {
        row: 5, col: 7, width: 3,
        title: 'Portfolio Cap Rate',
        formula: '=IFERROR(SUM({{Cash Flow Analysis}}!M14:M25)/SUM({{Properties}}!F:F),0)',
        format: '0.00%',
        color: '#8B5CF6',
        icon: '📊',
        subtitle: '="Market avg: 6.5%"'
      },
      {
        row: 5, col: 10, width: 3,
        title: 'Cash-on-Cash Return',
        formula: '=IFERROR(SUM({{Cash Flow Analysis}}!M14:M25)/(SUM({{Properties}}!E:E)*0.25),0)',
        format: '0.00%',
        color: '#F59E0B',
        icon: '🎯',
        subtitle: '="Target: 8%+"'
      }
    ];
    
    // Create KPI cards
    kpiCards.forEach(kpi => {
      dash.merge(kpi.row, kpi.col, kpi.row + 3, kpi.col + kpi.width - 1);
      dash.format(kpi.row, kpi.col, kpi.row + 3, kpi.col + kpi.width - 1, {
        background: '#FFFFFF',
        border: true
      });
      
      dash.setValue(kpi.row, kpi.col, kpi.icon + ' ' + kpi.title);
      dash.format(kpi.row, kpi.col, kpi.row, kpi.col + kpi.width - 1, {
        fontSize: 11,
        fontColor: '#6B7280',
        fontWeight: 'bold'
      });
      
      dash.setFormula(kpi.row + 1, kpi.col, kpi.formula);
      dash.format(kpi.row + 1, kpi.col, kpi.row + 1, kpi.col + kpi.width - 1, {
        fontSize: 20,
        fontWeight: 'bold',
        fontColor: kpi.color,
        numberFormat: kpi.format,
        horizontalAlignment: 'center'
      });
      
      dash.setFormula(kpi.row + 2, kpi.col, kpi.subtitle);
      dash.format(kpi.row + 2, kpi.col, kpi.row + 2, kpi.col + kpi.width - 1, {
        fontSize: 10,
        fontColor: '#9CA3AF',
        horizontalAlignment: 'center'
      });
    });
    
    // Property Performance Analysis (Row 11-20)
    dash.setValue(11, 1, 'PROPERTY PERFORMANCE ANALYSIS');
    dash.format(11, 1, 11, 8, {
      fontSize: 12,
      fontWeight: 'bold',
      background: '#F3F4F6',
      border: true
    });
    
    const performanceHeaders = ['Property', 'Purchase Price', 'Current Value', 'Appreciation', 'Annual NOI', 'Cap Rate', 'Cash-on-Cash', 'IRR'];
    performanceHeaders.forEach((header, idx) => {
      dash.setValue(12, idx + 1, header);
    });
    dash.format(12, 1, 12, 8, {
      fontWeight: 'bold',
      background: '#E5E7EB',
      border: true
    });
    
    // Property performance data
    for (let i = 0; i < 7; i++) {
      const row = 13 + i;
      dash.setFormula(row, 1, `=IFERROR(INDEX({{Properties}}!B:B,${i + 2}),"")`);
      dash.setFormula(row, 2, `=IFERROR(INDEX({{Properties}}!E:E,${i + 2}),0)`);
      dash.setFormula(row, 3, `=IFERROR(INDEX({{Properties}}!F:F,${i + 2}),0)`);
      dash.setFormula(row, 4, `=IFERROR((C${row}-B${row})/B${row},0)`);
      dash.setFormula(row, 5, `=IFERROR(INDEX({{Cash Flow Analysis}}!M:M,${i + 14}),0)`);
      dash.setFormula(row, 6, `=IFERROR(E${row}/C${row},0)`);
      dash.setFormula(row, 7, `=IFERROR(E${row}/(B${row}*0.25),0)`);
      dash.setFormula(row, 8, `=IFERROR(INDEX({{ROI Calculator}}!H:H,${i + 4}),0)`);
      
      dash.format(row, 2, row, 3, { numberFormat: '$#,##0' });
      dash.format(row, 4, row, 4, { numberFormat: '0.0%' });
      dash.format(row, 5, row, 5, { numberFormat: '$#,##0' });
      dash.format(row, 6, row, 8, { numberFormat: '0.00%' });
    }
    dash.format(13, 1, 19, 8, { border: true });
    
    // Financing Summary (Row 11-20, Columns 10-14)
    dash.setValue(11, 10, 'FINANCING OVERVIEW');
    dash.format(11, 10, 11, 14, {
      fontSize: 12,
      fontWeight: 'bold',
      background: '#F3F4F6',
      border: true
    });
    
    const financingMetrics = [
      { label: 'Total Debt', formula: '=SUM({{Financing}}!E:E)' },
      { label: 'Avg Interest Rate', formula: '=IFERROR(AVERAGE({{Financing}}!F:F),0)' },
      { label: 'Monthly Debt Service', formula: '=SUM({{Financing}}!H:H)' },
      { label: 'LTV Ratio', formula: '=IFERROR(SUM({{Financing}}!E:E)/SUM({{Properties}}!F:F),0)' },
      { label: 'DSCR', formula: '=IFERROR(SUM({{Cash Flow Analysis}}!M14:M25)/12/SUM({{Financing}}!H:H),0)' },
      { label: 'Equity Multiple', formula: '=IFERROR(SUM({{Properties}}!F:F)/SUM({{Properties}}!E:E),0)' }
    ];
    
    financingMetrics.forEach((metric, idx) => {
      const row = 13 + idx;
      dash.setValue(row, 10, metric.label);
      dash.setFormula(row, 13, metric.formula);
      dash.format(row, 10, row, 14, { border: true });
      
      if (metric.label.includes('Rate') || metric.label.includes('Ratio') || metric.label.includes('DSCR') || metric.label.includes('Multiple')) {
        dash.format(row, 13, row, 13, { numberFormat: '0.00%' });
        if (metric.label === 'DSCR' || metric.label === 'Equity Multiple') {
          dash.format(row, 13, row, 13, { numberFormat: '0.00' });
        }
      } else {
        dash.format(row, 13, row, 13, { numberFormat: '$#,##0' });
      }
    });
    
    // Investment Metrics Chart Area (Row 22-30)
    dash.setValue(22, 1, 'RETURN METRICS');
    dash.format(22, 1, 22, 6, {
      fontSize: 12,
      fontWeight: 'bold',
      background: '#F3F4F6',
      border: true
    });
    
    const returnMetrics = [
      { metric: 'Gross Rental Yield', formula: '=IFERROR(SUM({{Properties}}!H:H)*12/SUM({{Properties}}!F:F),0)' },
      { metric: 'Net Rental Yield', formula: '=IFERROR(SUM({{Cash Flow Analysis}}!M14:M25)/SUM({{Properties}}!F:F),0)' },
      { metric: '5-Year Total Return', formula: '=IFERROR((SUM({{Properties}}!F:F)*1.2+SUM({{Cash Flow Analysis}}!M14:M25)*5-SUM({{Properties}}!E:E))/SUM({{Properties}}!E:E),0)' },
      { metric: '10-Year Total Return', formula: '=IFERROR((SUM({{Properties}}!F:F)*1.5+SUM({{Cash Flow Analysis}}!M14:M25)*10-SUM({{Properties}}!E:E))/SUM({{Properties}}!E:E),0)' },
      { metric: 'Break-Even Occupancy', formula: '=IFERROR(SUM({{Financing}}!H:H)/SUM({{Properties}}!H:H),0)' },
      { metric: 'Payback Period (Years)', formula: '=IFERROR(SUM({{Properties}}!E:E)/SUM({{Cash Flow Analysis}}!M14:M25),0)' }
    ];
    
    returnMetrics.forEach((item, idx) => {
      const row = 23 + idx;
      dash.setValue(row, 1, item.metric);
      dash.setFormula(row, 4, item.formula);
      dash.format(row, 1, row, 6, { border: true });
      
      if (item.metric.includes('Years')) {
        dash.format(row, 4, row, 4, { numberFormat: '0.0' });
      } else {
        dash.format(row, 4, row, 4, { numberFormat: '0.00%' });
      }
    });
    
    // Market Comparison (Row 22-30, Columns 8-14)
    dash.setValue(22, 8, 'MARKET COMPARISON');
    dash.format(22, 8, 22, 14, {
      fontSize: 12,
      fontWeight: 'bold',
      background: '#F3F4F6',
      border: true
    });
    
    const marketHeaders = ['Metric', 'Your Portfolio', 'Market Average', 'Performance'];
    marketHeaders.forEach((header, idx) => {
      dash.setValue(23, idx + 8, header);
    });
    dash.format(23, 8, 23, 11, {
      fontWeight: 'bold',
      background: '#E5E7EB',
      border: true
    });
    
    const marketComps = [
      { metric: 'Cap Rate', portfolio: '=IFERROR(SUM({{Cash Flow Analysis}}!M14:M25)/SUM({{Properties}}!F:F),0)', market: 0.065 },
      { metric: 'Cash-on-Cash', portfolio: '=IFERROR(SUM({{Cash Flow Analysis}}!M14:M25)/(SUM({{Properties}}!E:E)*0.25),0)', market: 0.08 },
      { metric: 'Appreciation', portfolio: '=IFERROR(AVERAGE(({{Properties}}!F:F-{{Properties}}!E:E)/{{Properties}}!E:E),0)', market: 0.04 },
      { metric: 'Gross Yield', portfolio: '=IFERROR(SUM({{Properties}}!H:H)*12/SUM({{Properties}}!F:F),0)', market: 0.075 },
      { metric: 'LTV', portfolio: '=IFERROR(SUM({{Financing}}!E:E)/SUM({{Properties}}!F:F),0)', market: 0.75 }
    ];
    
    marketComps.forEach((comp, idx) => {
      const row = 24 + idx;
      dash.setValue(row, 8, comp.metric);
      dash.setFormula(row, 9, comp.portfolio);
      dash.setValue(row, 10, comp.market);
      dash.setFormula(row, 11, `=IF(I${row}>J${row},"↑ Outperforming","↓ Underperforming")`);
      
      dash.format(row, 9, row, 10, { numberFormat: '0.00%' });
      dash.format(row, 8, row, 11, { border: true });
    });
    
    // Set optimal column widths
    for (let col = 1; col <= 14; col++) {
      dash.setColumnWidth(col, col % 2 === 0 ? 85 : 95);
    }
    dash.setRowHeight(1, 35);
    dash.setRowHeight(2, 35);
  },
  
  /**
   * Build Investment Properties Sheet
   */
  buildInvestmentProperties: function(builder) {
    const properties = builder.sheet('Properties');
    
    // Header
    properties.merge(1, 1, 1, 18);
    properties.setValue(1, 1, '🏘️ Investment Properties Analysis');
    properties.format(1, 1, 1, 18, {
      fontSize: 16,
      fontWeight: 'bold',
      background: '#047857',
      fontColor: '#FFFFFF',
      horizontalAlignment: 'center'
    });
    
    // Column headers
    const headers = [
      'Property ID', 'Property Name', 'Address', 'Property Type', 
      'Purchase Price', 'Current Value', 'Purchase Date', 'Monthly Rent',
      'Annual Taxes', 'Annual Insurance', 'Renovation Costs', 'HOA Fees',
      'Management %', 'Vacancy Rate', 'Bedrooms', 'Bathrooms', 'Sq Ft', 'Year Built'
    ];
    
    headers.forEach((header, idx) => {
      properties.setValue(2, idx + 1, header);
    });
    properties.format(2, 1, 2, headers.length, {
      fontWeight: 'bold',
      background: '#E5E7EB',
      border: true
    });
    
    // Sample investment properties
    const sampleProperties = [
      ['INV001', 'Maple Street Duplex', '123 Maple St, Austin, TX', 'Multi-Family',
       450000, 525000, '2021-03-15', 3800, 6500, 2200, 25000, 0,
       0.08, 0.05, '4 (2x2)', '2 (1x1)', 2400, 1985],
      ['INV002', 'Downtown Condo', '456 Urban Plaza #501, Dallas, TX', 'Condo',
       325000, 360000, '2022-07-10', 2500, 3800, 1500, 5000, 350,
       0.10, 0.03, 2, 2, 1200, 2010],
      ['INV003', 'Suburban Rental', '789 Oak Drive, Houston, TX', 'Single-Family',
       275000, 310000, '2020-11-20', 2200, 4200, 1800, 15000, 0,
       0.08, 0.05, 3, 2, 1800, 1995]
    ];
    
    sampleProperties.forEach((property, idx) => {
      property.forEach((value, col) => {
        if (col === 4 || col === 5 || col === 7 || col === 8 || col === 9 || col === 10 || col === 11) {
          // Money fields
          properties.setValue(idx + 3, col + 1, value);
          properties.format(idx + 3, col + 1, idx + 3, col + 1, { numberFormat: '$#,##0' });
        } else if (col === 6) { // Date
          properties.setValue(idx + 3, col + 1, value);
          properties.format(idx + 3, col + 1, idx + 3, col + 1, { numberFormat: 'yyyy-mm-dd' });
        } else if (col === 12 || col === 13) { // Percentages
          properties.setValue(idx + 3, col + 1, value);
          properties.format(idx + 3, col + 1, idx + 3, col + 1, { numberFormat: '0.0%' });
        } else {
          properties.setValue(idx + 3, col + 1, value);
        }
      });
    });
    
    // Add calculated columns
    properties.setValue(2, 19, 'Gross Annual Rent');
    properties.setValue(2, 20, 'Total Annual Expenses');
    properties.setValue(2, 21, 'Annual NOI');
    properties.setValue(2, 22, 'Cap Rate');
    properties.format(2, 19, 2, 22, {
      fontWeight: 'bold',
      background: '#E5E7EB',
      border: true
    });
    
    // Formulas for calculated fields
    for (let i = 3; i <= 5; i++) {
      properties.setFormula(i, 19, `=H${i}*12*(1-N${i})`);
      properties.setFormula(i, 20, `=I${i}+J${i}+L${i}+H${i}*12*M${i}`);
      properties.setFormula(i, 21, `=S${i}-T${i}`);
      properties.setFormula(i, 22, `=U${i}/F${i}`);
      
      properties.format(i, 19, i, 21, { numberFormat: '$#,##0' });
      properties.format(i, 22, i, 22, { numberFormat: '0.00%' });
    }
    
    // Data validation
    properties.addValidation('D3:D100', ['Single-Family', 'Multi-Family', 'Condo', 'Townhouse', 'Commercial', 'Mixed-Use']);
    
    // Conditional formatting for cap rates
    properties.addConditionalFormat(3, 22, 100, 22, {
      type: 'gradient',
      minColor: '#FEE2E2',
      midColor: '#FEF3C7',
      maxColor: '#D1FAE5',
      minValue: '0.04',
      midValue: '0.065',
      maxValue: '0.09'
    });
    
    // Set column widths
    properties.setColumnWidth(1, 100);
    properties.setColumnWidth(2, 150);
    properties.setColumnWidth(3, 250);
  },
  
  /**
   * Build Cash Flow Analysis Sheet
   */
  buildCashFlowAnalysis: function(builder) {
    const cashFlow = builder.sheet('Cash Flow Analysis');
    
    // Header
    cashFlow.merge(1, 1, 1, 14);
    cashFlow.setValue(1, 1, '💸 Cash Flow Analysis & Projections');
    cashFlow.format(1, 1, 1, 14, {
      fontSize: 16,
      fontWeight: 'bold',
      background: '#047857',
      fontColor: '#FFFFFF',
      horizontalAlignment: 'center'
    });
    
    // Monthly Cash Flow Section
    cashFlow.setValue(3, 1, 'MONTHLY CASH FLOW BREAKDOWN');
    cashFlow.format(3, 1, 3, 14, {
      fontSize: 12,
      fontWeight: 'bold',
      background: '#F3F4F6'
    });
    
    // Headers
    const monthHeaders = ['Category', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec', 'Annual Total'];
    monthHeaders.forEach((header, idx) => {
      cashFlow.setValue(4, idx + 1, header);
    });
    cashFlow.format(4, 1, 4, 14, {
      fontWeight: 'bold',
      background: '#E5E7EB',
      border: true
    });
    
    // Income section
    cashFlow.setValue(5, 1, 'INCOME');
    cashFlow.format(5, 1, 5, 14, {
      fontWeight: 'bold',
      background: '#D1FAE5'
    });
    
    const incomeCategories = [
      { label: 'Rental Income', formula: '=SUM({{Properties}}!H:H)' },
      { label: 'Late Fees', formula: '=SUM({{Properties}}!H:H)*0.02' },
      { label: 'Other Income', formula: '=SUM({{Properties}}!H:H)*0.01' }
    ];
    
    incomeCategories.forEach((cat, idx) => {
      const row = 6 + idx;
      cashFlow.setValue(row, 1, cat.label);
      for (let month = 2; month <= 13; month++) {
        cashFlow.setFormula(row, month, cat.formula);
        cashFlow.format(row, month, row, month, { numberFormat: '$#,##0' });
      }
      cashFlow.setFormula(row, 14, `=SUM(B${row}:M${row})`);
      cashFlow.format(row, 14, row, 14, { numberFormat: '$#,##0', fontWeight: 'bold' });
    });
    
    // Total Income row
    cashFlow.setValue(9, 1, 'Total Income');
    for (let month = 2; month <= 13; month++) {
      cashFlow.setFormula(9, month, `=SUM(B6:B8)`);
      cashFlow.format(9, month, 9, month, { numberFormat: '$#,##0', fontWeight: 'bold' });
    }
    cashFlow.setFormula(9, 14, '=SUM(B9:M9)');
    cashFlow.format(9, 14, 9, 14, { numberFormat: '$#,##0', fontWeight: 'bold' });
    
    // Expenses section
    cashFlow.setValue(11, 1, 'EXPENSES');
    cashFlow.format(11, 1, 11, 14, {
      fontWeight: 'bold',
      background: '#FEE2E2'
    });
    
    const expenseCategories = [
      { label: 'Mortgage Payment', formula: '=SUM({{Financing}}!H:H)' },
      { label: 'Property Tax', formula: '=SUM({{Properties}}!I:I)/12' },
      { label: 'Insurance', formula: '=SUM({{Properties}}!J:J)/12' },
      { label: 'HOA Fees', formula: '=SUM({{Properties}}!L:L)' },
      { label: 'Property Management', formula: '=B9*AVERAGE({{Properties}}!M:M)' },
      { label: 'Maintenance Reserve', formula: '=B9*0.05' },
      { label: 'Utilities', formula: '=COUNT({{Properties}}!A:A)*50' },
      { label: 'Marketing/Vacancy', formula: '=B9*AVERAGE({{Properties}}!N:N)' }
    ];
    
    expenseCategories.forEach((cat, idx) => {
      const row = 12 + idx;
      cashFlow.setValue(row, 1, cat.label);
      for (let month = 2; month <= 13; month++) {
        cashFlow.setFormula(row, month, cat.formula);
        cashFlow.format(row, month, row, month, { numberFormat: '$#,##0' });
      }
      cashFlow.setFormula(row, 14, `=SUM(B${row}:M${row})`);
      cashFlow.format(row, 14, row, 14, { numberFormat: '$#,##0', fontWeight: 'bold' });
    });
    
    // Total Expenses row
    cashFlow.setValue(20, 1, 'Total Expenses');
    for (let month = 2; month <= 13; month++) {
      cashFlow.setFormula(20, month, `=SUM(B12:B19)`);
      cashFlow.format(20, month, 20, month, { numberFormat: '$#,##0', fontWeight: 'bold', fontColor: '#DC2626' });
    }
    cashFlow.setFormula(20, 14, '=SUM(B20:M20)');
    cashFlow.format(20, 14, 20, 14, { numberFormat: '$#,##0', fontWeight: 'bold', fontColor: '#DC2626' });
    
    // Net Cash Flow
    cashFlow.setValue(22, 1, 'NET CASH FLOW');
    for (let month = 2; month <= 13; month++) {
      cashFlow.setFormula(22, month, `=B9-B20`);
      cashFlow.format(22, month, 22, month, { 
        numberFormat: '$#,##0', 
        fontWeight: 'bold',
        background: '#F3F4F6'
      });
    }
    cashFlow.setFormula(22, 14, '=N9-N20');
    cashFlow.format(22, 14, 22, 14, { 
      numberFormat: '$#,##0', 
      fontWeight: 'bold',
      background: '#F3F4F6'
    });
    
    // Cumulative Cash Flow
    cashFlow.setValue(23, 1, 'Cumulative CF');
    cashFlow.setFormula(23, 2, '=B22');
    for (let month = 3; month <= 13; month++) {
      cashFlow.setFormula(23, month, `=${String.fromCharCode(64 + month - 1)}23+${String.fromCharCode(64 + month)}22`);
    }
    cashFlow.setFormula(23, 14, '=M23');
    cashFlow.format(23, 2, 23, 14, { numberFormat: '$#,##0', fontColor: '#059669' });
    
    // 5-Year Projection Section
    cashFlow.setValue(26, 1, '5-YEAR CASH FLOW PROJECTION');
    cashFlow.format(26, 1, 26, 7, {
      fontSize: 12,
      fontWeight: 'bold',
      background: '#F3F4F6'
    });
    
    const projectionHeaders = ['Year', 'Gross Income', 'Total Expenses', 'Net Cash Flow', 'Cumulative CF', 'Annual Growth', 'ROI'];
    projectionHeaders.forEach((header, idx) => {
      cashFlow.setValue(27, idx + 1, header);
    });
    cashFlow.format(27, 1, 27, 7, {
      fontWeight: 'bold',
      background: '#E5E7EB',
      border: true
    });
    
    // 5-year projections with 3% annual growth
    for (let year = 1; year <= 5; year++) {
      const row = 27 + year;
      cashFlow.setValue(row, 1, `Year ${year}`);
      cashFlow.setFormula(row, 2, `=N9*POWER(1.03,${year - 1})`);
      cashFlow.setFormula(row, 3, `=N20*POWER(1.02,${year - 1})`);
      cashFlow.setFormula(row, 4, `=B${row}-C${row}`);
      cashFlow.setFormula(row, 5, year === 1 ? `=D${row}` : `=E${row - 1}+D${row}`);
      cashFlow.setFormula(row, 6, year === 1 ? '0%' : `=(B${row}-B${row - 1})/B${row - 1}`);
      cashFlow.setFormula(row, 7, `=E${row}/SUM({{Properties}}!E:E)`);
      
      cashFlow.format(row, 2, row, 5, { numberFormat: '$#,##0' });
      cashFlow.format(row, 6, row, 7, { numberFormat: '0.0%' });
    }
    cashFlow.format(28, 1, 32, 7, { border: true });
    
    // Set column widths
    cashFlow.setColumnWidth(1, 150);
    for (let col = 2; col <= 14; col++) {
      cashFlow.setColumnWidth(col, 80);
    }
  },
  
  /**
   * Build ROI Calculator Sheet
   */
  buildROICalculator: function(builder) {
    const roi = builder.sheet('ROI Calculator');
    
    // Header
    roi.merge(1, 1, 1, 12);
    roi.setValue(1, 1, '🧮 Return on Investment Calculator');
    roi.format(1, 1, 1, 12, {
      fontSize: 16,
      fontWeight: 'bold',
      background: '#047857',
      fontColor: '#FFFFFF',
      horizontalAlignment: 'center'
    });
    
    // Property ROI Analysis
    roi.setValue(3, 1, 'PROPERTY-BY-PROPERTY ROI ANALYSIS');
    roi.format(3, 1, 3, 12, {
      fontSize: 12,
      fontWeight: 'bold',
      background: '#F3F4F6'
    });
    
    const roiHeaders = [
      'Property', 'Initial Investment', 'Annual Cash Flow', 'Year 1 ROI', 
      'Year 3 ROI', 'Year 5 ROI', 'Break-Even (Years)', 'IRR (5-Year)',
      'NPV (10%, 5-Year)', 'Equity Multiple', 'Annual Appreciation', 'Total Return (5-Year)'
    ];
    
    roiHeaders.forEach((header, idx) => {
      roi.setValue(4, idx + 1, header);
    });
    roi.format(4, 1, 4, 12, {
      fontWeight: 'bold',
      background: '#E5E7EB',
      border: true
    });
    
    // ROI calculations for each property
    for (let i = 0; i < 5; i++) {
      const row = 5 + i;
      roi.setFormula(row, 1, `=IFERROR(INDEX({{Properties}}!B:B,${i + 2}),"")`);
      roi.setFormula(row, 2, `=IFERROR(INDEX({{Properties}}!E:E,${i + 2})+INDEX({{Properties}}!K:K,${i + 2}),0)`);
      roi.setFormula(row, 3, `=IFERROR(INDEX({{Properties}}!U:U,${i + 2}),0)`);
      roi.setFormula(row, 4, `=C${row}/B${row}`);
      roi.setFormula(row, 5, `=(C${row}*3)/B${row}`);
      roi.setFormula(row, 6, `=(C${row}*5)/B${row}`);
      roi.setFormula(row, 7, `=IFERROR(B${row}/C${row},0)`);
      roi.setFormula(row, 8, `=IFERROR(IRR({-B${row},C${row},C${row}*1.03,C${row}*1.06,C${row}*1.09,C${row}*1.12+INDEX({{Properties}}!F:F,${i + 2})*1.2}),0)`);
      roi.setFormula(row, 9, `=NPV(0.1,C${row},C${row}*1.03,C${row}*1.06,C${row}*1.09,C${row}*1.12)-B${row}`);
      roi.setFormula(row, 10, `=IFERROR((INDEX({{Properties}}!F:F,${i + 2})+C${row}*5)/B${row},0)`);
      roi.setFormula(row, 11, `=IFERROR((INDEX({{Properties}}!F:F,${i + 2})-INDEX({{Properties}}!E:E,${i + 2}))/INDEX({{Properties}}!E:E,${i + 2})/DATEDIF(INDEX({{Properties}}!G:G,${i + 2}),TODAY(),"Y"),0)`);
      roi.setFormula(row, 12, `=IFERROR((INDEX({{Properties}}!F:F,${i + 2})*1.2+C${row}*5-B${row})/B${row},0)`);
      
      roi.format(row, 2, row, 3, { numberFormat: '$#,##0' });
      roi.format(row, 4, row, 6, { numberFormat: '0.0%' });
      roi.format(row, 7, row, 7, { numberFormat: '0.0' });
      roi.format(row, 8, row, 8, { numberFormat: '0.00%' });
      roi.format(row, 9, row, 9, { numberFormat: '$#,##0' });
      roi.format(row, 10, row, 10, { numberFormat: '0.00' });
      roi.format(row, 11, row, 12, { numberFormat: '0.0%' });
    }
    roi.format(5, 1, 9, 12, { border: true });
    
    // Portfolio Summary
    roi.setValue(12, 1, 'PORTFOLIO SUMMARY METRICS');
    roi.format(12, 1, 12, 6, {
      fontSize: 12,
      fontWeight: 'bold',
      background: '#F3F4F6'
    });
    
    const portfolioMetrics = [
      { label: 'Total Investment', formula: '=SUM(B5:B9)' },
      { label: 'Annual Cash Flow', formula: '=SUM(C5:C9)' },
      { label: 'Portfolio ROI', formula: '=B14/B13' },
      { label: 'Weighted Avg IRR', formula: '=SUMPRODUCT(H5:H9,B5:B9)/SUM(B5:B9)' },
      { label: 'Portfolio NPV', formula: '=SUM(I5:I9)' },
      { label: 'Blended Cap Rate', formula: '=B14/SUM({{Properties}}!F:F)' }
    ];
    
    portfolioMetrics.forEach((metric, idx) => {
      const row = 13 + idx;
      roi.setValue(row, 1, metric.label);
      roi.setFormula(row, 3, metric.formula);
      roi.format(row, 1, row, 6, { border: true });
      
      if (metric.label.includes('Investment') || metric.label.includes('Cash Flow') || metric.label.includes('NPV')) {
        roi.format(row, 3, row, 3, { numberFormat: '$#,##0', fontWeight: 'bold' });
      } else {
        roi.format(row, 3, row, 3, { numberFormat: '0.00%', fontWeight: 'bold' });
      }
    });
    
    // Scenario Analysis Section
    roi.setValue(20, 1, 'SCENARIO ANALYSIS');
    roi.format(20, 1, 20, 8, {
      fontSize: 12,
      fontWeight: 'bold',
      background: '#F3F4F6'
    });
    
    const scenarioHeaders = ['Scenario', 'Rent Growth', 'Expense Growth', 'Vacancy Rate', 'Year 5 Cash Flow', 'Year 5 ROI', 'IRR', 'Risk Level'];
    scenarioHeaders.forEach((header, idx) => {
      roi.setValue(21, idx + 1, header);
    });
    roi.format(21, 1, 21, 8, {
      fontWeight: 'bold',
      background: '#E5E7EB',
      border: true
    });
    
    const scenarios = [
      ['Conservative', 0.02, 0.03, 0.08, '=B14*POWER(1.02,5)*(1-D22)', '=E22*5/B13', '=IRR({-B13,E22,E22,E22,E22,E22})', 'Low'],
      ['Base Case', 0.03, 0.025, 0.05, '=B14*POWER(1.03,5)*(1-D23)', '=E23*5/B13', '=IRR({-B13,E23,E23,E23,E23,E23})', 'Medium'],
      ['Optimistic', 0.04, 0.02, 0.03, '=B14*POWER(1.04,5)*(1-D24)', '=E24*5/B13', '=IRR({-B13,E24,E24,E24,E24,E24})', 'Medium-High'],
      ['Aggressive', 0.05, 0.015, 0.02, '=B14*POWER(1.05,5)*(1-D25)', '=E25*5/B13', '=IRR({-B13,E25,E25,E25,E25,E25})', 'High']
    ];
    
    scenarios.forEach((scenario, idx) => {
      const row = 22 + idx;
      scenario.forEach((value, col) => {
        if (col === 0 || col === 7) {
          roi.setValue(row, col + 1, value);
        } else if (col >= 1 && col <= 3) {
          roi.setValue(row, col + 1, value);
          roi.format(row, col + 1, row, col + 1, { numberFormat: '0.0%' });
        } else {
          roi.setFormula(row, col + 1, value);
          if (col === 4) {
            roi.format(row, col + 1, row, col + 1, { numberFormat: '$#,##0' });
          } else {
            roi.format(row, col + 1, row, col + 1, { numberFormat: '0.00%' });
          }
        }
      });
    });
    roi.format(22, 1, 25, 8, { border: true });
    
    // Set column widths
    roi.setColumnWidth(1, 150);
    for (let col = 2; col <= 12; col++) {
      roi.setColumnWidth(col, 100);
    }
  },
  
  /**
   * Build Market Comps Sheet
   */
  buildMarketComps: function(builder) {
    const comps = builder.sheet('Market Comps');
    
    // Header
    comps.merge(1, 1, 1, 14);
    comps.setValue(1, 1, '🏘️ Market Comparables Analysis');
    comps.format(1, 1, 1, 14, {
      fontSize: 16,
      fontWeight: 'bold',
      background: '#047857',
      fontColor: '#FFFFFF',
      horizontalAlignment: 'center'
    });
    
    // Comparable Properties Section
    comps.setValue(3, 1, 'COMPARABLE PROPERTIES');
    comps.format(3, 1, 3, 14, {
      fontSize: 12,
      fontWeight: 'bold',
      background: '#F3F4F6'
    });
    
    const compHeaders = [
      'Property Address', 'Type', 'Beds', 'Baths', 'Sq Ft', 'Year Built',
      'Sale Price', 'Sale Date', 'Price/Sq Ft', 'Days on Market',
      'Monthly Rent', 'Gross Yield', 'Cap Rate', 'Distance (mi)'
    ];
    
    compHeaders.forEach((header, idx) => {
      comps.setValue(4, idx + 1, header);
    });
    comps.format(4, 1, 4, 14, {
      fontWeight: 'bold',
      background: '#E5E7EB',
      border: true
    });
    
    // Sample comparable properties
    const comparables = [
      ['234 Oak Street', 'Single-Family', 3, 2, 1850, 1992, 295000, '2024-01-15', 159, 35, 2300, 0.094, 0.072, 0.5],
      ['567 Pine Avenue', 'Single-Family', 3, 2.5, 1950, 1995, 310000, '2024-01-10', 159, 28, 2400, 0.093, 0.071, 0.8],
      ['890 Elm Court', 'Single-Family', 4, 2, 2100, 1990, 325000, '2023-12-20', 155, 42, 2500, 0.092, 0.070, 1.2],
      ['123 Birch Lane', 'Single-Family', 3, 2, 1750, 1988, 285000, '2024-01-05', 163, 30, 2250, 0.095, 0.073, 1.5],
      ['456 Cedar Drive', 'Single-Family', 3, 2.5, 1900, 1998, 305000, '2023-12-15', 161, 25, 2350, 0.092, 0.071, 2.0]
    ];
    
    comparables.forEach((comp, idx) => {
      comp.forEach((value, col) => {
        if (col === 6) { // Sale Price
          comps.setValue(idx + 5, col + 1, value);
          comps.format(idx + 5, col + 1, idx + 5, col + 1, { numberFormat: '$#,##0' });
        } else if (col === 7) { // Sale Date
          comps.setValue(idx + 5, col + 1, value);
          comps.format(idx + 5, col + 1, idx + 5, col + 1, { numberFormat: 'yyyy-mm-dd' });
        } else if (col === 8) { // Price/Sq Ft
          comps.setValue(idx + 5, col + 1, value);
          comps.format(idx + 5, col + 1, idx + 5, col + 1, { numberFormat: '$#,##0' });
        } else if (col === 10) { // Monthly Rent
          comps.setValue(idx + 5, col + 1, value);
          comps.format(idx + 5, col + 1, idx + 5, col + 1, { numberFormat: '$#,##0' });
        } else if (col === 11 || col === 12) { // Yields
          comps.setValue(idx + 5, col + 1, value);
          comps.format(idx + 5, col + 1, idx + 5, col + 1, { numberFormat: '0.0%' });
        } else {
          comps.setValue(idx + 5, col + 1, value);
        }
      });
    });
    comps.format(5, 1, 9, 14, { border: true });
    
    // Market Statistics
    comps.setValue(12, 1, 'MARKET STATISTICS');
    comps.format(12, 1, 12, 6, {
      fontSize: 12,
      fontWeight: 'bold',
      background: '#F3F4F6'
    });
    
    const marketStats = [
      { label: 'Average Sale Price', formula: '=AVERAGE(G5:G9)' },
      { label: 'Median Sale Price', formula: '=MEDIAN(G5:G9)' },
      { label: 'Avg Price/Sq Ft', formula: '=AVERAGE(I5:I9)' },
      { label: 'Avg Days on Market', formula: '=AVERAGE(J5:J9)' },
      { label: 'Avg Monthly Rent', formula: '=AVERAGE(K5:K9)' },
      { label: 'Avg Cap Rate', formula: '=AVERAGE(M5:M9)' }
    ];
    
    marketStats.forEach((stat, idx) => {
      const row = 13 + idx;
      comps.setValue(row, 1, stat.label);
      comps.setFormula(row, 3, stat.formula);
      comps.format(row, 1, row, 6, { border: true });
      
      if (stat.label.includes('Price') || stat.label.includes('Rent')) {
        comps.format(row, 3, row, 3, { numberFormat: '$#,##0' });
      } else if (stat.label.includes('Cap Rate')) {
        comps.format(row, 3, row, 3, { numberFormat: '0.00%' });
      } else if (stat.label.includes('Days')) {
        comps.format(row, 3, row, 3, { numberFormat: '0' });
      }
    });
    
    // Value Estimation
    comps.setValue(12, 8, 'PROPERTY VALUATION');
    comps.format(12, 8, 12, 14, {
      fontSize: 12,
      fontWeight: 'bold',
      background: '#F3F4F6'
    });
    
    comps.setValue(13, 8, 'Subject Property:');
    comps.setValue(13, 10, 'Your Property Name');
    
    const valuationMethods = [
      { method: 'Comp Sales Method', formula: '=C15*AVERAGE(E5:E9)' },
      { method: 'Income Approach', formula: '=K20/C18' },
      { method: 'GRM Method', formula: '=K20*125' },
      { method: 'Replacement Cost', formula: '=E14*150+50000' },
      { method: 'Weighted Average', formula: '=(C14*0.4+C15*0.3+C16*0.2+C17*0.1)' }
    ];
    
    valuationMethods.forEach((val, idx) => {
      const row = 14 + idx;
      comps.setValue(row, 8, val.method);
      comps.setFormula(row, 11, val.formula);
      comps.format(row, 11, row, 11, { numberFormat: '$#,##0' });
      comps.format(row, 8, row, 14, { border: true });
    });
    
    // Input fields for subject property
    comps.setValue(20, 8, 'Subject Inputs:');
    comps.setValue(20, 9, 'Sq Ft:');
    comps.setValue(20, 10, 1800);
    comps.setValue(20, 11, 'Est. Rent:');
    comps.setValue(20, 12, 2400);
    comps.format(20, 12, 20, 12, { numberFormat: '$#,##0' });
    
    // Set column widths
    comps.setColumnWidth(1, 150);
    for (let col = 2; col <= 14; col++) {
      comps.setColumnWidth(col, 85);
    }
  },
  
  /**
   * Build Financing Sheet
   */
  buildFinancingSheet: function(builder) {
    const financing = builder.sheet('Financing');
    
    // Header
    financing.merge(1, 1, 1, 12);
    financing.setValue(1, 1, '🏦 Financing & Mortgage Analysis');
    financing.format(1, 1, 1, 12, {
      fontSize: 16,
      fontWeight: 'bold',
      background: '#047857',
      fontColor: '#FFFFFF',
      horizontalAlignment: 'center'
    });
    
    // Active Loans Section
    financing.setValue(3, 1, 'ACTIVE LOANS');
    financing.format(3, 1, 3, 12, {
      fontSize: 12,
      fontWeight: 'bold',
      background: '#F3F4F6'
    });
    
    const loanHeaders = [
      'Property', 'Lender', 'Loan Type', 'Original Amount', 'Current Balance',
      'Interest Rate', 'Term (Years)', 'Monthly Payment', 'Start Date',
      'Maturity Date', 'Prepayment Penalty', 'Notes'
    ];
    
    loanHeaders.forEach((header, idx) => {
      financing.setValue(4, idx + 1, header);
    });
    financing.format(4, 1, 4, 12, {
      fontWeight: 'bold',
      background: '#E5E7EB',
      border: true
    });
    
    // Sample loans
    const loans = [
      ['Maple Street Duplex', 'Wells Fargo', 'Conventional', 337500, 325000, 0.045, 30, 1810, '2021-03-15', '2051-03-15', 'None after Year 3', 'Fixed rate'],
      ['Downtown Condo', 'Chase Bank', '15-Year Fixed', 243750, 220000, 0.0375, 15, 1780, '2022-07-10', '2037-07-10', 'None', 'Great rate'],
      ['Suburban Rental', 'Local CU', 'ARM 5/1', 206250, 195000, 0.0425, 30, 1045, '2020-11-20', '2050-11-20', '2% if before 2025', 'Rate adjusts 2025']
    ];
    
    loans.forEach((loan, idx) => {
      loan.forEach((value, col) => {
        if (col === 3 || col === 4 || col === 7) { // Money fields
          financing.setValue(idx + 5, col + 1, value);
          financing.format(idx + 5, col + 1, idx + 5, col + 1, { numberFormat: '$#,##0' });
        } else if (col === 5) { // Interest rate
          financing.setValue(idx + 5, col + 1, value);
          financing.format(idx + 5, col + 1, idx + 5, col + 1, { numberFormat: '0.00%' });
        } else if (col === 8 || col === 9) { // Dates
          financing.setValue(idx + 5, col + 1, value);
          financing.format(idx + 5, col + 1, idx + 5, col + 1, { numberFormat: 'yyyy-mm-dd' });
        } else {
          financing.setValue(idx + 5, col + 1, value);
        }
      });
    });
    financing.format(5, 1, 7, 12, { border: true });
    
    // Loan Summary
    financing.setValue(10, 1, 'FINANCING SUMMARY');
    financing.format(10, 1, 10, 6, {
      fontSize: 12,
      fontWeight: 'bold',
      background: '#F3F4F6'
    });
    
    const loanSummary = [
      { label: 'Total Original Loans', formula: '=SUM(D5:D7)' },
      { label: 'Total Current Balance', formula: '=SUM(E5:E7)' },
      { label: 'Total Monthly Payments', formula: '=SUM(H5:H7)' },
      { label: 'Weighted Avg Rate', formula: '=SUMPRODUCT(F5:F7,E5:E7)/SUM(E5:E7)' },
      { label: 'Total Annual Debt Service', formula: '=C13*12' },
      { label: 'Portfolio LTV', formula: '=B12/SUM({{Properties}}!F:F)' }
    ];
    
    loanSummary.forEach((item, idx) => {
      const row = 11 + idx;
      financing.setValue(row, 1, item.label);
      financing.setFormula(row, 3, item.formula);
      financing.format(row, 1, row, 6, { border: true });
      
      if (item.label.includes('Rate') || item.label.includes('LTV')) {
        financing.format(row, 3, row, 3, { numberFormat: '0.00%', fontWeight: 'bold' });
      } else {
        financing.format(row, 3, row, 3, { numberFormat: '$#,##0', fontWeight: 'bold' });
      }
    });
    
    // Refinance Analysis
    financing.setValue(18, 1, 'REFINANCE ANALYSIS');
    financing.format(18, 1, 18, 10, {
      fontSize: 12,
      fontWeight: 'bold',
      background: '#F3F4F6'
    });
    
    const refiHeaders = ['Property', 'Current Rate', 'Current Payment', 'New Rate', 'New Payment', 'Monthly Savings', 'Break-Even (Months)', 'NPV (5 Years)', 'Refinance?', 'Notes'];
    refiHeaders.forEach((header, idx) => {
      financing.setValue(19, idx + 1, header);
    });
    financing.format(19, 1, 19, 10, {
      fontWeight: 'bold',
      background: '#E5E7EB',
      border: true
    });
    
    // Refinance scenarios
    for (let i = 0; i < 3; i++) {
      const row = 20 + i;
      financing.setFormula(row, 1, `=A${5 + i}`);
      financing.setFormula(row, 2, `=F${5 + i}`);
      financing.setFormula(row, 3, `=H${5 + i}`);
      financing.setValue(row, 4, 0.0375); // New rate
      financing.setFormula(row, 5, `=PMT(D${row}/12,G${5 + i}*12,-E${5 + i})`);
      financing.setFormula(row, 6, `=C${row}-E${row}`);
      financing.setFormula(row, 7, `=IF(F${row}>0,5000/F${row},"N/A")`);
      financing.setFormula(row, 8, `=IF(F${row}>0,NPV(0.005,F${row}*{1,1,1,1,1,1,1,1,1,1,1,1})*5-5000,0)`);
      financing.setFormula(row, 9, `=IF(H${row}>10000,"Yes","No")`);
      
      financing.format(row, 2, row, 2, { numberFormat: '0.00%' });
      financing.format(row, 3, row, 3, { numberFormat: '$#,##0' });
      financing.format(row, 4, row, 4, { numberFormat: '0.00%' });
      financing.format(row, 5, row, 6, { numberFormat: '$#,##0' });
      financing.format(row, 8, row, 8, { numberFormat: '$#,##0' });
    }
    financing.format(20, 1, 22, 10, { border: true });
    
    // Set column widths
    financing.setColumnWidth(1, 150);
    for (let col = 2; col <= 12; col++) {
      financing.setColumnWidth(col, 90);
    }
  },
  
  /**
   * Build Tax Analysis Sheet
   */
  buildTaxAnalysis: function(builder) {
    const tax = builder.sheet('Tax Analysis');
    
    // Header
    tax.merge(1, 1, 1, 10);
    tax.setValue(1, 1, '📋 Tax Analysis & Deductions');
    tax.format(1, 1, 1, 10, {
      fontSize: 16,
      fontWeight: 'bold',
      background: '#047857',
      fontColor: '#FFFFFF',
      horizontalAlignment: 'center'
    });
    
    // Annual Tax Summary
    tax.setValue(3, 1, 'ANNUAL TAX DEDUCTIONS');
    tax.format(3, 1, 3, 10, {
      fontSize: 12,
      fontWeight: 'bold',
      background: '#F3F4F6'
    });
    
    const deductionCategories = [
      { category: 'OPERATING EXPENSES', items: [
        { label: 'Mortgage Interest', formula: '=SUM({{Financing}}!H:H)*12*0.75' },
        { label: 'Property Taxes', formula: '=SUM({{Properties}}!I:I)' },
        { label: 'Insurance Premiums', formula: '=SUM({{Properties}}!J:J)' },
        { label: 'HOA Fees', formula: '=SUM({{Properties}}!L:L)*12' },
        { label: 'Property Management', formula: '=SUM({{Properties}}!H:H)*12*AVERAGE({{Properties}}!M:M)' },
        { label: 'Repairs & Maintenance', formula: '=SUM({{Properties}}!H:H)*12*0.05' },
        { label: 'Utilities', formula: '=COUNT({{Properties}}!A:A)*50*12' },
        { label: 'Professional Services', formula: '=2500' },
        { label: 'Marketing & Advertising', formula: '=SUM({{Properties}}!H:H)*12*AVERAGE({{Properties}}!N:N)' },
        { label: 'Travel & Auto', formula: '=3600' }
      ]},
      { category: 'DEPRECIATION', items: [
        { label: 'Residential Depreciation', formula: '=SUM({{Properties}}!E:E)*0.8/27.5' },
        { label: 'Improvements Depreciation', formula: '=SUM({{Properties}}!K:K)/27.5' },
        { label: 'Personal Property', formula: '=SUM({{Properties}}!E:E)*0.05/5' }
      ]}
    ];
    
    let currentRow = 4;
    deductionCategories.forEach(category => {
      tax.setValue(currentRow, 1, category.category);
      tax.format(currentRow, 1, currentRow, 10, {
        fontWeight: 'bold',
        background: '#E5E7EB'
      });
      currentRow++;
      
      category.items.forEach(item => {
        tax.setValue(currentRow, 2, item.label);
        tax.setFormula(currentRow, 4, item.formula);
        tax.format(currentRow, 4, currentRow, 4, { numberFormat: '$#,##0' });
        tax.format(currentRow, 1, currentRow, 10, { border: true });
        currentRow++;
      });
      
      // Subtotal for category
      tax.setValue(currentRow, 2, `Total ${category.category}`);
      tax.setFormula(currentRow, 4, `=SUM(D${currentRow - category.items.length}:D${currentRow - 1})`);
      tax.format(currentRow, 4, currentRow, 4, { 
        numberFormat: '$#,##0', 
        fontWeight: 'bold',
        background: '#F3F4F6'
      });
      tax.format(currentRow, 1, currentRow, 10, { border: true });
      currentRow += 2;
    });
    
    // Tax Calculation
    tax.setValue(currentRow, 1, 'TAX IMPACT ANALYSIS');
    tax.format(currentRow, 1, currentRow, 10, {
      fontSize: 12,
      fontWeight: 'bold',
      background: '#F3F4F6'
    });
    currentRow++;
    
    const taxCalculation = [
      { label: 'Gross Rental Income', formula: '=SUM({{Properties}}!H:H)*12*(1-AVERAGE({{Properties}}!N:N))' },
      { label: 'Total Operating Expenses', formula: '=D15' },
      { label: 'Total Depreciation', formula: '=D19' },
      { label: 'Net Taxable Income', formula: `=D${currentRow + 1}-D${currentRow + 2}-D${currentRow + 3}` },
      { label: 'Estimated Tax Rate', value: 0.28 },
      { label: 'Tax on Rental Income', formula: `=D${currentRow + 4}*D${currentRow + 5}` },
      { label: 'Tax Savings from Deductions', formula: `=(D${currentRow + 2}+D${currentRow + 3})*D${currentRow + 5}` },
      { label: 'Net Tax Impact', formula: `=D${currentRow + 6}-D${currentRow + 7}` }
    ];
    
    taxCalculation.forEach((item, idx) => {
      const row = currentRow + 1 + idx;
      tax.setValue(row, 2, item.label);
      if (item.value !== undefined) {
        tax.setValue(row, 4, item.value);
        tax.format(row, 4, row, 4, { numberFormat: '0%' });
      } else {
        tax.setFormula(row, 4, item.formula);
        tax.format(row, 4, row, 4, { numberFormat: '$#,##0' });
      }
      tax.format(row, 1, row, 10, { border: true });
      
      if (item.label === 'Net Tax Impact') {
        tax.format(row, 4, row, 4, { 
          fontWeight: 'bold',
          background: '#F3F4F6'
        });
      }
    });
    
    // Set column widths
    tax.setColumnWidth(1, 50);
    tax.setColumnWidth(2, 200);
    tax.setColumnWidth(3, 50);
    tax.setColumnWidth(4, 120);
    for (let col = 5; col <= 10; col++) {
      tax.setColumnWidth(col, 80);
    }
  },
  
  /**
   * Build Investment Reports Sheet
   */
  buildInvestmentReports: function(builder) {
    const reports = builder.sheet('Reports');
    
    // Header
    reports.merge(1, 1, 1, 12);
    reports.setValue(1, 1, '📊 Investment Performance Reports');
    reports.format(1, 1, 1, 12, {
      fontSize: 16,
      fontWeight: 'bold',
      background: '#047857',
      fontColor: '#FFFFFF',
      horizontalAlignment: 'center'
    });
    
    // Executive Summary
    reports.setValue(3, 1, 'EXECUTIVE SUMMARY');
    reports.format(3, 1, 3, 12, {
      fontSize: 12,
      fontWeight: 'bold',
      background: '#F3F4F6'
    });
    
    const executiveSummary = [
      { metric: 'Portfolio Value', formula: '=SUM({{Properties}}!F:F)', format: '$#,##0' },
      { metric: 'Total Equity', formula: '=SUM({{Properties}}!F:F)-SUM({{Financing}}!E:E)', format: '$#,##0' },
      { metric: 'Annual Cash Flow', formula: '=SUM({{Cash Flow Analysis}}!N22)', format: '$#,##0' },
      { metric: 'Portfolio Cap Rate', formula: '=C6/C4', format: '0.00%' },
      { metric: 'Cash-on-Cash Return', formula: '=C6/(SUM({{Properties}}!E:E)*0.25)', format: '0.00%' },
      { metric: 'Average IRR', formula: '=AVERAGE({{ROI Calculator}}!H5:H9)', format: '0.00%' },
      { metric: 'Total ROI (All-Time)', formula: '=(C4+SUM({{Cash Flow Analysis}}!N22)*3-SUM({{Properties}}!E:E))/SUM({{Properties}}!E:E)', format: '0.00%' }
    ];
    
    executiveSummary.forEach((item, idx) => {
      const row = 4 + idx;
      reports.setValue(row, 1, item.metric);
      reports.setFormula(row, 3, item.formula);
      reports.format(row, 3, row, 3, { 
        numberFormat: item.format,
        fontWeight: 'bold'
      });
      reports.format(row, 1, row, 12, { border: true });
    });
    
    // Performance Trends
    reports.setValue(13, 1, 'QUARTERLY PERFORMANCE TRENDS');
    reports.format(13, 1, 13, 12, {
      fontSize: 12,
      fontWeight: 'bold',
      background: '#F3F4F6'
    });
    
    const quarterHeaders = ['Metric', 'Q1 2023', 'Q2 2023', 'Q3 2023', 'Q4 2023', 'Q1 2024', 'Q2 2024', 'Q3 2024', 'Q4 2024 (Est)', 'YoY Growth', 'Trend'];
    quarterHeaders.forEach((header, idx) => {
      reports.setValue(14, idx + 1, header);
    });
    reports.format(14, 1, 14, 11, {
      fontWeight: 'bold',
      background: '#E5E7EB',
      border: true
    });
    
    // Quarterly metrics (simplified formulas)
    const quarterlyMetrics = [
      { metric: 'Revenue', base: 25000, growth: 1.02 },
      { metric: 'NOI', base: 18000, growth: 1.025 },
      { metric: 'Occupancy', base: 0.92, growth: 1.005, isPercent: true },
      { metric: 'Avg Rent', base: 2200, growth: 1.015 }
    ];
    
    quarterlyMetrics.forEach((metric, idx) => {
      const row = 15 + idx;
      reports.setValue(row, 1, metric.metric);
      
      for (let q = 0; q < 8; q++) {
        const value = metric.base * Math.pow(metric.growth, q);
        reports.setValue(row, q + 2, value);
        
        if (metric.isPercent) {
          reports.format(row, q + 2, row, q + 2, { numberFormat: '0.0%' });
        } else {
          reports.format(row, q + 2, row, q + 2, { numberFormat: '$#,##0' });
        }
      }
      
      reports.setFormula(row, 10, `=(I${row}-B${row})/B${row}`);
      reports.format(row, 10, row, 10, { numberFormat: '0.0%' });
      reports.setFormula(row, 11, `=IF(J${row}>0,"📈 Up","📉 Down")`);
      reports.format(row, 1, row, 11, { border: true });
    });
    
    // Risk Assessment
    reports.setValue(21, 1, 'RISK ASSESSMENT');
    reports.format(21, 1, 21, 6, {
      fontSize: 12,
      fontWeight: 'bold',
      background: '#F3F4F6'
    });
    
    const riskFactors = [
      { factor: 'Leverage Risk (LTV)', value: '=SUM({{Financing}}!E:E)/SUM({{Properties}}!F:F)', threshold: 0.75, status: '=IF(B22<C22,"✅ Low","⚠️ High")' },
      { factor: 'Concentration Risk', value: '=MAX({{Properties}}!F:F)/SUM({{Properties}}!F:F)', threshold: 0.5, status: '=IF(B23<C23,"✅ Diversified","⚠️ Concentrated")' },
      { factor: 'Cash Flow Coverage', value: '=SUM({{Cash Flow Analysis}}!N22)/12/SUM({{Financing}}!H:H)', threshold: 1.25, status: '=IF(B24>C24,"✅ Strong","⚠️ Weak")' },
      { factor: 'Vacancy Sensitivity', value: '=AVERAGE({{Properties}}!N:N)', threshold: 0.08, status: '=IF(B25<C25,"✅ Low","⚠️ High")' },
      { factor: 'Market Dependency', value: '=STDEV({{Properties}}!V:V)/AVERAGE({{Properties}}!V:V)', threshold: 0.15, status: '=IF(B26<C26,"✅ Stable","⚠️ Volatile")' }
    ];
    
    reports.setValue(22, 1, 'Risk Factor');
    reports.setValue(22, 2, 'Current');
    reports.setValue(22, 3, 'Threshold');
    reports.setValue(22, 4, 'Status');
    reports.format(22, 1, 22, 4, {
      fontWeight: 'bold',
      background: '#E5E7EB',
      border: true
    });
    
    riskFactors.forEach((risk, idx) => {
      const row = 23 + idx;
      reports.setValue(row, 1, risk.factor);
      reports.setFormula(row, 2, risk.value);
      reports.setValue(row, 3, risk.threshold);
      reports.setFormula(row, 4, risk.status);
      
      if (risk.factor.includes('LTV') || risk.factor.includes('Concentration') || risk.factor.includes('Vacancy')) {
        reports.format(row, 2, row, 3, { numberFormat: '0.0%' });
      } else {
        reports.format(row, 2, row, 3, { numberFormat: '0.00' });
      }
      reports.format(row, 1, row, 4, { border: true });
    });
    
    // Set column widths
    reports.setColumnWidth(1, 180);
    for (let col = 2; col <= 12; col++) {
      reports.setColumnWidth(col, 85);
    }
  },

  /**
   * Lead Pipeline - Comprehensive lead management and conversion tracking
   */
  setupLeadPipeline: function(spreadsheet, isPreview) {
    const builder = new TemplateBuilderPro(spreadsheet, isPreview);
    const sheets = builder.createSheets([
      'Dashboard',
      'Leads',
      'Lead Scoring',
      'Follow-ups',
      'Sources',
      'Campaigns',
      'Conversion',
      'Reports'
    ]);
    
    // Build each component
    this.buildLeadDashboard(sheets.Dashboard, sheets);
    this.buildLeadsDatabase(sheets.Leads);
    this.buildLeadScoring(sheets['Lead Scoring']);
    this.buildFollowUps(sheets['Follow-ups']);
    this.buildLeadSources(sheets.Sources);
    this.buildCampaigns(sheets.Campaigns);
    this.buildConversionAnalytics(sheets.Conversion);
    this.buildLeadReports(sheets.Reports);
    
    return builder.complete();
  },
  
  buildLeadDashboard: function(dash, sheets) {
    // Title Section
    dash.setValue(1, 1, 'Lead Pipeline Dashboard');
    dash.merge(1, 1, 1, 12);
    dash.format(1, 1, 1, 12, {
      fontSize: 18,
      fontWeight: 'bold',
      horizontalAlignment: 'center',
      background: '#1F2937',
      fontColor: '#FFFFFF'
    });
    
    // Period Selector
    dash.setValue(2, 1, 'Period:');
    dash.setValue(2, 2, 'Last 30 Days');
    dash.addValidation('B2', ['Today', 'Last 7 Days', 'Last 30 Days', 'Last 90 Days', 'YTD', 'All Time']);
    
    // KPI Cards Row 1
    dash.setValue(4, 1, 'LEAD METRICS');
    dash.merge(4, 1, 4, 12);
    dash.format(4, 1, 4, 12, {
      fontSize: 12,
      fontWeight: 'bold',
      background: '#F3F4F6'
    });
    
    // Total Leads Card
    dash.setValue(5, 1, 'Total Leads');
    dash.merge(5, 1, 5, 3);
    dash.setFormula(6, 1, '=COUNTIF({{Leads}}!B:B,"<>")');
    dash.merge(6, 1, 6, 3);
    dash.setFormula(7, 1, '=SPARKLINE({{Leads}}!Z2:Z32,{"charttype","line";"linewidth",2;"color","#10B981"})');
    dash.merge(7, 1, 7, 3);
    
    // New This Month Card
    dash.setValue(5, 4, 'New This Month');
    dash.merge(5, 4, 5, 6);
    dash.setFormula(6, 4, '=COUNTIFS({{Leads}}!C:C,">="&EOMONTH(TODAY(),-1)+1,{{Leads}}!C:C,"<="&TODAY())');
    dash.merge(6, 4, 6, 6);
    dash.setFormula(7, 4, '=TEXT((D6-COUNTIFS({{Leads}}!C:C,">="&EOMONTH(TODAY(),-2)+1,{{Leads}}!C:C,"<="&EOMONTH(TODAY(),-1)))/COUNTIFS({{Leads}}!C:C,">="&EOMONTH(TODAY(),-2)+1,{{Leads}}!C:C,"<="&EOMONTH(TODAY(),-1)),"↑ +0.0%")');
    dash.merge(7, 4, 7, 6);
    
    // Hot Leads Card
    dash.setValue(5, 7, 'Hot Leads');
    dash.merge(5, 7, 5, 9);
    dash.setFormula(6, 7, '=COUNTIF({{Leads}}!H:H,"Hot")');
    dash.merge(6, 7, 6, 9);
    dash.setFormula(7, 7, '=TEXT(G6/A6,"0.0%")&" of total"');
    dash.merge(7, 7, 7, 9);
    
    // Conversion Rate Card
    dash.setValue(5, 10, 'Conversion Rate');
    dash.merge(5, 10, 5, 12);
    dash.setFormula(6, 10, '=COUNTIF({{Leads}}!G:G,"Converted")/COUNTIF({{Leads}}!B:B,"<>")');
    dash.merge(6, 10, 6, 12);
    dash.format(6, 10, 6, 12, { numberFormat: '0.0%' });
    dash.setFormula(7, 10, '=IF(J6>0.15,"🟢 Above Target","🔴 Below Target")');
    dash.merge(7, 10, 7, 12);
    
    // Format KPI cards
    dash.format(5, 1, 7, 12, {
      border: true,
      fontSize: 11
    });
    dash.format(6, 1, 6, 12, {
      fontSize: 20,
      fontWeight: 'bold',
      horizontalAlignment: 'center'
    });
    
    // Lead Status Breakdown
    dash.setValue(9, 1, 'LEAD STATUS BREAKDOWN');
    dash.merge(9, 1, 9, 6);
    dash.format(9, 1, 9, 6, {
      fontSize: 12,
      fontWeight: 'bold',
      background: '#F3F4F6'
    });
    
    const statuses = ['New', 'Contacted', 'Qualified', 'Proposal', 'Negotiation', 'Converted', 'Lost'];
    statuses.forEach((status, idx) => {
      dash.setValue(10 + idx, 1, status);
      dash.setFormula(10 + idx, 2, `=COUNTIF({{Leads}}!G:G,"${status}")`);
      dash.setFormula(10 + idx, 3, `=B${10 + idx}/$A$6`);
      dash.format(10 + idx, 3, 10 + idx, 3, { numberFormat: '0.0%' });
      dash.setFormula(10 + idx, 4, `=REPT("█",ROUND(C${10 + idx}*20,0))`);
      dash.merge(10 + idx, 4, 10 + idx, 6);
    });
    
    // Lead Sources Performance
    dash.setValue(9, 7, 'TOP LEAD SOURCES');
    dash.merge(9, 7, 9, 12);
    dash.format(9, 7, 9, 12, {
      fontSize: 12,
      fontWeight: 'bold',
      background: '#F3F4F6'
    });
    
    dash.setValue(10, 7, 'Source');
    dash.setValue(10, 9, 'Leads');
    dash.setValue(10, 10, 'Conv %');
    dash.setValue(10, 11, 'Value');
    dash.merge(10, 11, 10, 12);
    
    // Sample top sources
    const sources = [
      ['Website', '=COUNTIF({{Leads}}!E:E,G11)', '=COUNTIFS({{Leads}}!E:E,G11,{{Leads}}!G:G,"Converted")/H11', '=SUMIFS({{Leads}}!K:K,{{Leads}}!E:E,G11)'],
      ['Referral', '=COUNTIF({{Leads}}!E:E,G12)', '=COUNTIFS({{Leads}}!E:E,G12,{{Leads}}!G:G,"Converted")/H12', '=SUMIFS({{Leads}}!K:K,{{Leads}}!E:E,G12)'],
      ['Social Media', '=COUNTIF({{Leads}}!E:E,G13)', '=COUNTIFS({{Leads}}!E:E,G13,{{Leads}}!G:G,"Converted")/H13', '=SUMIFS({{Leads}}!K:K,{{Leads}}!E:E,G13)'],
      ['Cold Call', '=COUNTIF({{Leads}}!E:E,G14)', '=COUNTIFS({{Leads}}!E:E,G14,{{Leads}}!G:G,"Converted")/H14', '=SUMIFS({{Leads}}!K:K,{{Leads}}!E:E,G14)'],
      ['Email Campaign', '=COUNTIF({{Leads}}!E:E,G15)', '=COUNTIFS({{Leads}}!E:E,G15,{{Leads}}!G:G,"Converted")/H15', '=SUMIFS({{Leads}}!K:K,{{Leads}}!E:E,G15)'],
      ['Partner', '=COUNTIF({{Leads}}!E:E,G16)', '=COUNTIFS({{Leads}}!E:E,G16,{{Leads}}!G:G,"Converted")/H16', '=SUMIFS({{Leads}}!K:K,{{Leads}}!E:E,G16)']
    ];
    
    sources.forEach((source, idx) => {
      dash.setValue(11 + idx, 7, source[0]);
      dash.setFormula(11 + idx, 9, source[1]);
      dash.setFormula(11 + idx, 10, source[2]);
      dash.format(11 + idx, 10, 11 + idx, 10, { numberFormat: '0.0%' });
      dash.setFormula(11 + idx, 11, source[3]);
      dash.merge(11 + idx, 11, 11 + idx, 12);
      dash.format(11 + idx, 11, 11 + idx, 12, { numberFormat: '$#,##0' });
    });
    
    // Activity Timeline
    dash.setValue(18, 1, 'RECENT ACTIVITY');
    dash.merge(18, 1, 18, 12);
    dash.format(18, 1, 18, 12, {
      fontSize: 12,
      fontWeight: 'bold',
      background: '#F3F4F6'
    });
    
    dash.setValue(19, 1, 'Time');
    dash.setValue(19, 2, 'Lead');
    dash.merge(19, 2, 19, 4);
    dash.setValue(19, 5, 'Action');
    dash.merge(19, 5, 19, 8);
    dash.setValue(19, 9, 'Agent');
    dash.merge(19, 9, 19, 10);
    dash.setValue(19, 11, 'Result');
    dash.merge(19, 11, 19, 12);
    
    // Sample activities
    for (let i = 0; i < 8; i++) {
      dash.setFormula(20 + i, 1, `=TEXT(NOW()-${i}/24,"HH:MM")`);
      dash.setFormula(20 + i, 2, `=INDEX({{Leads}}!B:B,RANDBETWEEN(2,20))`);
      dash.merge(20 + i, 2, 20 + i, 4);
      dash.setValue(20 + i, 5, ['Email sent', 'Call made', 'Meeting scheduled', 'Proposal sent', 'Follow-up', 'Note added'][Math.floor(Math.random() * 6)]);
      dash.merge(20 + i, 5, 20 + i, 8);
      dash.setValue(20 + i, 9, ['John Smith', 'Jane Doe', 'Mike Johnson'][Math.floor(Math.random() * 3)]);
      dash.merge(20 + i, 9, 20 + i, 10);
      dash.setValue(20 + i, 11, ['Positive', 'Pending', 'No Answer', 'Scheduled'][Math.floor(Math.random() * 4)]);
      dash.merge(20 + i, 11, 20 + i, 12);
    }
    
    // Lead Score Distribution
    dash.setValue(29, 1, 'LEAD SCORE DISTRIBUTION');
    dash.merge(29, 1, 29, 6);
    dash.format(29, 1, 29, 6, {
      fontSize: 12,
      fontWeight: 'bold',
      background: '#F3F4F6'
    });
    
    const scoreRanges = ['0-20', '21-40', '41-60', '61-80', '81-100'];
    dash.setValue(30, 1, 'Score Range');
    dash.setValue(30, 2, 'Count');
    dash.setValue(30, 3, 'Percentage');
    dash.merge(30, 3, 30, 4);
    dash.setValue(30, 5, 'Visual');
    dash.merge(30, 5, 30, 6);
    
    scoreRanges.forEach((range, idx) => {
      dash.setValue(31 + idx, 1, range);
      const [min, max] = range.split('-').map(Number);
      dash.setFormula(31 + idx, 2, `=COUNTIFS({{Lead Scoring}}!C:C,">="${min}",{{Lead Scoring}}!C:C,"<="${max}")`);
      dash.setFormula(31 + idx, 3, `=B${31 + idx}/SUM(B$31:B$35)`);
      dash.merge(31 + idx, 3, 31 + idx, 4);
      dash.format(31 + idx, 3, 31 + idx, 4, { numberFormat: '0.0%' });
      dash.setFormula(31 + idx, 5, `=REPT("■",ROUND(C${31 + idx}*10,0))`);
      dash.merge(31 + idx, 5, 31 + idx, 6);
    });
    
    // Conversion Funnel
    dash.setValue(29, 7, 'CONVERSION FUNNEL');
    dash.merge(29, 7, 29, 12);
    dash.format(29, 7, 29, 12, {
      fontSize: 12,
      fontWeight: 'bold',
      background: '#F3F4F6'
    });
    
    const funnelStages = [
      ['Leads', '=COUNTIF({{Leads}}!B:B,"<>")', '100%'],
      ['Contacted', '=COUNTIF({{Leads}}!G:G,"Contacted")+COUNTIF({{Leads}}!G:G,"Qualified")+COUNTIF({{Leads}}!G:G,"Proposal")+COUNTIF({{Leads}}!G:G,"Negotiation")+COUNTIF({{Leads}}!G:G,"Converted")', '=H31/$H$30'],
      ['Qualified', '=COUNTIF({{Leads}}!G:G,"Qualified")+COUNTIF({{Leads}}!G:G,"Proposal")+COUNTIF({{Leads}}!G:G,"Negotiation")+COUNTIF({{Leads}}!G:G,"Converted")', '=H32/$H$30'],
      ['Proposal', '=COUNTIF({{Leads}}!G:G,"Proposal")+COUNTIF({{Leads}}!G:G,"Negotiation")+COUNTIF({{Leads}}!G:G,"Converted")', '=H33/$H$30'],
      ['Converted', '=COUNTIF({{Leads}}!G:G,"Converted")', '=H34/$H$30']
    ];
    
    funnelStages.forEach((stage, idx) => {
      dash.setValue(30 + idx, 7, stage[0]);
      dash.setFormula(30 + idx, 8, stage[1]);
      dash.setFormula(30 + idx, 9, stage[2]);
      dash.format(30 + idx, 9, 30 + idx, 9, { numberFormat: '0.0%' });
      dash.setFormula(30 + idx, 10, `=REPT("▬",ROUND(I${30 + idx}*30,0))`);
      dash.merge(30 + idx, 10, 30 + idx, 12);
      
      // Color code based on stage
      const colors = ['#10B981', '#3B82F6', '#8B5CF6', '#F59E0B', '#EF4444'];
      dash.format(30 + idx, 10, 30 + idx, 12, { fontColor: colors[idx] });
    });
    
    // Column widths
    dash.setColumnWidth(1, 100);
    for (let col = 2; col <= 12; col++) {
      dash.setColumnWidth(col, 80);
    }
  },
  
  buildLeadsDatabase: function(leads) {
    // Headers
    const headers = [
      'Lead ID', 'Name', 'Date Added', 'Company', 'Source', 'Type',
      'Status', 'Score', 'Assigned To', 'Last Contact', 'Est. Value',
      'Email', 'Phone', 'Address', 'City', 'State', 'Zip',
      'Industry', 'Company Size', 'Budget', 'Timeline', 'Interest Level',
      'Next Action', 'Next Action Date', 'Notes', 'Tags', 'Campaign',
      'Conversion Date', 'Lost Reason'
    ];
    
    headers.forEach((header, idx) => {
      leads.setValue(1, idx + 1, header);
    });
    
    leads.format(1, 1, 1, headers.length, {
      fontWeight: 'bold',
      background: '#1F2937',
      fontColor: '#FFFFFF',
      border: true
    });
    
    // Sample data
    const sampleLeads = [
      ['L001', 'John Anderson', '=TODAY()-45', 'Anderson Corp', 'Website', 'Buyer',
       'Qualified', 75, 'Sarah Johnson', '=TODAY()-2', 450000,
       'john@anderson.com', '555-0101', '123 Main St', 'Boston', 'MA', '02101',
       'Technology', 'Medium', '$400-500k', 'Q2 2024', 'High',
       'Send proposal', '=TODAY()+3', 'Very interested in waterfront properties', 'premium,urgent', 'Spring Campaign',
       '', ''],
      ['L002', 'Mary Smith', '=TODAY()-30', 'Smith Investments', 'Referral', 'Investor',
       'Proposal', 85, 'Mike Davis', '=TODAY()-1', 750000,
       'mary@smithinvest.com', '555-0102', '456 Oak Ave', 'Cambridge', 'MA', '02139',
       'Finance', 'Large', '$700-800k', 'Q1 2024', 'Very High',
       'Negotiation call', '=TODAY()+1', 'Looking for multi-family investment', 'investor,hot', 'Referral Program',
       '', ''],
      ['L003', 'Robert Wilson', '=TODAY()-20', 'Wilson Holdings', 'Cold Call', 'Buyer',
       'New', 45, 'Sarah Johnson', '=TODAY()-5', 325000,
       'robert@wilson.com', '555-0103', '789 Pine Rd', 'Somerville', 'MA', '02144',
       'Retail', 'Small', '$300-350k', 'Q3 2024', 'Medium',
       'Initial meeting', '=TODAY()+7', 'First-time buyer, needs guidance', 'first-time', 'Cold Outreach',
       '', ''],
      ['L004', 'Jennifer Lee', '=TODAY()-15', 'Lee Enterprises', 'Social Media', 'Seller',
       'Contacted', 60, 'Mike Davis', '=TODAY()-3', 550000,
       'jennifer@lee-ent.com', '555-0104', '321 Elm St', 'Newton', 'MA', '02458',
       'Healthcare', 'Medium', 'N/A', 'ASAP', 'High',
       'Property valuation', '=TODAY()+2', 'Wants to sell commercial property', 'seller,commercial', 'Social Media',
       '', ''],
      ['L005', 'David Brown', '=TODAY()-60', 'Brown Development', 'Partner', 'Developer',
       'Converted', 95, 'Sarah Johnson', '=TODAY()-10', 1200000,
       'david@browndev.com', '555-0105', '654 Maple Dr', 'Brookline', 'MA', '02445',
       'Real Estate', 'Large', '$1M+', 'Completed', 'Very High',
       'Close follow-up', '=TODAY()+30', 'Purchased development site', 'vip,closed-won', 'Partner Network',
       '=TODAY()-10', '']
    ];
    
    sampleLeads.forEach((lead, idx) => {
      lead.forEach((value, col) => {
        if (typeof value === 'string' && value.startsWith('=')) {
          leads.setFormula(idx + 2, col + 1, value);
        } else {
          leads.setValue(idx + 2, col + 1, value);
        }
      });
    });
    
    // Add data validation
    leads.addValidation('E2:E1000', ['Website', 'Referral', 'Social Media', 'Cold Call', 'Email Campaign', 'Partner', 'Event', 'Direct Mail']);
    leads.addValidation('F2:F1000', ['Buyer', 'Seller', 'Investor', 'Developer', 'Renter']);
    leads.addValidation('G2:G1000', ['New', 'Contacted', 'Qualified', 'Proposal', 'Negotiation', 'Converted', 'Lost']);
    leads.addValidation('I2:I1000', ['Sarah Johnson', 'Mike Davis', 'John Smith', 'Jane Doe']);
    leads.addValidation('V2:V1000', ['Very Low', 'Low', 'Medium', 'High', 'Very High']);
    
    // Conditional formatting for Status
    leads.addConditionalFormat(2, 7, 1000, 7, {
      type: 'cell',
      condition: 'TEXT_EQUALS',
      value: 'Converted',
      format: { background: '#10B981', fontColor: '#FFFFFF' }
    });
    
    leads.addConditionalFormat(2, 7, 1000, 7, {
      type: 'cell',
      condition: 'TEXT_EQUALS',
      value: 'Lost',
      format: { background: '#EF4444', fontColor: '#FFFFFF' }
    });
    
    // Conditional formatting for Score
    leads.addConditionalFormat(2, 8, 1000, 8, {
      type: 'gradient',
      minColor: '#FEE2E2',
      midColor: '#FEF3C7',
      maxColor: '#D1FAE5',
      minValue: '0',
      midValue: '50',
      maxValue: '100'
    });
    
    // Format currency columns
    leads.format(2, 11, 1000, 11, { numberFormat: '$#,##0' });
    leads.format(2, 20, 1000, 20, { numberFormat: '$#,##0' });
    
    // Column widths
    leads.setColumnWidth(1, 80);
    leads.setColumnWidth(2, 150);
    leads.setColumnWidth(3, 100);
    leads.setColumnWidth(4, 150);
    for (let col = 5; col <= 29; col++) {
      leads.setColumnWidth(col, 100);
    }
  },
  
  buildLeadScoring: function(scoring) {
    // Title
    scoring.setValue(1, 1, 'Lead Scoring Matrix');
    scoring.merge(1, 1, 1, 10);
    scoring.format(1, 1, 1, 10, {
      fontSize: 14,
      fontWeight: 'bold',
      horizontalAlignment: 'center',
      background: '#1F2937',
      fontColor: '#FFFFFF'
    });
    
    // Scoring criteria headers
    scoring.setValue(3, 1, 'Lead ID');
    scoring.setValue(3, 2, 'Name');
    scoring.setValue(3, 3, 'Total Score');
    scoring.setValue(3, 4, 'Grade');
    scoring.setValue(3, 5, 'Engagement');
    scoring.setValue(3, 6, 'Fit Score');
    scoring.setValue(3, 7, 'Activity');
    scoring.setValue(3, 8, 'Timeline');
    scoring.setValue(3, 9, 'Budget');
    scoring.setValue(3, 10, 'Authority');
    
    scoring.format(3, 1, 3, 10, {
      fontWeight: 'bold',
      background: '#F3F4F6',
      border: true
    });
    
    // Scoring formulas for each lead
    for (let row = 4; row <= 20; row++) {
      // Reference lead data
      scoring.setFormula(row, 1, `=IF({{Leads}}!A${row-2}="","",{{Leads}}!A${row-2})`);
      scoring.setFormula(row, 2, `=IF({{Leads}}!B${row-2}="","",{{Leads}}!B${row-2})`);
      
      // Calculate component scores
      // Engagement Score (0-25 points)
      scoring.setFormula(row, 5, `=IF(A${row}="","",(COUNTIF({{Follow-ups}}!B:B,A${row})*5)+IF({{Leads}}!J${row-2}>=TODAY()-7,10,IF({{Leads}}!J${row-2}>=TODAY()-14,5,0)))`);
      
      // Fit Score (0-25 points)
      scoring.setFormula(row, 6, `=IF(A${row}="","",IF({{Leads}}!F${row-2}="Investor",25,IF({{Leads}}!F${row-2}="Developer",20,IF({{Leads}}!F${row-2}="Buyer",15,10))))`);
      
      // Activity Score (0-20 points)
      scoring.setFormula(row, 7, `=IF(A${row}="","",IF({{Leads}}!G${row-2}="Proposal",20,IF({{Leads}}!G${row-2}="Qualified",15,IF({{Leads}}!G${row-2}="Contacted",10,5))))`);
      
      // Timeline Score (0-15 points)
      scoring.setFormula(row, 8, `=IF(A${row}="","",IF({{Leads}}!U${row-2}="ASAP",15,IF({{Leads}}!U${row-2}="Q1 2024",12,IF({{Leads}}!U${row-2}="Q2 2024",8,5))))`);
      
      // Budget Score (0-15 points)
      scoring.setFormula(row, 9, `=IF(A${row}="","",MIN(15,{{Leads}}!K${row-2}/100000*3))`);
      
      // Authority Score (0-10 points)
      scoring.setFormula(row, 10, `=IF(A${row}="","",IF(OR(SEARCH("CEO",{{Leads}}!B${row-2}),SEARCH("President",{{Leads}}!B${row-2}),SEARCH("Owner",{{Leads}}!B${row-2})),10,5))`);
      
      // Total Score
      scoring.setFormula(row, 3, `=IF(A${row}="","",SUM(E${row}:J${row}))`);
      
      // Grade
      scoring.setFormula(row, 4, `=IF(A${row}="","",IF(C${row}>=80,"A",IF(C${row}>=60,"B",IF(C${row}>=40,"C",IF(C${row}>=20,"D","F")))))`);
    }
    
    // Conditional formatting for scores
    scoring.addConditionalFormat(4, 3, 1000, 3, {
      type: 'gradient',
      minColor: '#FEE2E2',
      midColor: '#FEF3C7',
      maxColor: '#D1FAE5',
      minValue: '0',
      midValue: '50',
      maxValue: '100'
    });
    
    // Conditional formatting for grades
    scoring.addConditionalFormat(4, 4, 1000, 4, {
      type: 'cell',
      condition: 'TEXT_EQUALS',
      value: 'A',
      format: { background: '#10B981', fontColor: '#FFFFFF' }
    });
    
    // Scoring rubric
    scoring.setValue(3, 12, 'SCORING RUBRIC');
    scoring.merge(3, 12, 3, 15);
    scoring.format(3, 12, 3, 15, {
      fontWeight: 'bold',
      background: '#1F2937',
      fontColor: '#FFFFFF'
    });
    
    const rubric = [
      ['Grade', 'Score Range', 'Action', 'Priority'],
      ['A', '80-100', 'Hot Lead - Immediate Action', 'Highest'],
      ['B', '60-79', 'Warm Lead - Regular Follow-up', 'High'],
      ['C', '40-59', 'Cool Lead - Nurture', 'Medium'],
      ['D', '20-39', 'Cold Lead - Long-term', 'Low'],
      ['F', '0-19', 'Not Qualified', 'None']
    ];
    
    rubric.forEach((row, idx) => {
      row.forEach((cell, col) => {
        scoring.setValue(4 + idx, 12 + col, cell);
      });
    });
    
    scoring.format(4, 12, 4, 15, {
      fontWeight: 'bold',
      background: '#F3F4F6'
    });
    
    // Column widths
    scoring.setColumnWidth(1, 80);
    scoring.setColumnWidth(2, 150);
    for (let col = 3; col <= 15; col++) {
      scoring.setColumnWidth(col, 100);
    }
  },
  
  buildFollowUps: function(followups) {
    // Headers
    const headers = [
      'Follow-up ID', 'Lead ID', 'Lead Name', 'Type', 'Subject',
      'Date', 'Time', 'Duration', 'Assigned To', 'Status',
      'Outcome', 'Next Step', 'Notes', 'Priority', 'Reminder Set'
    ];
    
    headers.forEach((header, idx) => {
      followups.setValue(1, idx + 1, header);
    });
    
    followups.format(1, 1, 1, headers.length, {
      fontWeight: 'bold',
      background: '#1F2937',
      fontColor: '#FFFFFF',
      border: true
    });
    
    // Sample follow-ups
    const sampleFollowUps = [
      ['F001', 'L001', '=VLOOKUP(B2,{{Leads}}!A:B,2,FALSE)', 'Call', 'Initial consultation',
       '=TODAY()', '10:00 AM', '30 min', 'Sarah Johnson', 'Completed',
       'Interested', 'Send brochure', 'Client wants waterfront properties only', 'High', 'Yes'],
      ['F002', 'L002', '=VLOOKUP(B3,{{Leads}}!A:B,2,FALSE)', 'Meeting', 'Property tour',
       '=TODAY()+1', '2:00 PM', '2 hours', 'Mike Davis', 'Scheduled',
       'Pending', 'Show 3 properties', 'Client prefers downtown location', 'Very High', 'Yes'],
      ['F003', 'L003', '=VLOOKUP(B4,{{Leads}}!A:B,2,FALSE)', 'Email', 'Follow-up on inquiry',
       '=TODAY()-1', 'N/A', 'N/A', 'Sarah Johnson', 'Completed',
       'No response', 'Call tomorrow', 'Sent property listings', 'Medium', 'No'],
      ['F004', 'L004', '=VLOOKUP(B5,{{Leads}}!A:B,2,FALSE)', 'Call', 'Valuation discussion',
       '=TODAY()+2', '11:00 AM', '45 min', 'Mike Davis', 'Scheduled',
       'Pending', 'Prepare CMA', 'Seller motivated by job relocation', 'High', 'Yes'],
      ['F005', 'L001', '=VLOOKUP(B6,{{Leads}}!A:B,2,FALSE)', 'Email', 'Send proposal',
       '=TODAY()+3', 'N/A', 'N/A', 'Sarah Johnson', 'Planned',
       'Pending', 'Wait for response', 'Include financing options', 'High', 'Yes']
    ];
    
    sampleFollowUps.forEach((followUp, idx) => {
      followUp.forEach((value, col) => {
        if (typeof value === 'string' && value.startsWith('=')) {
          followups.setFormula(idx + 2, col + 1, value);
        } else {
          followups.setValue(idx + 2, col + 1, value);
        }
      });
    });
    
    // Data validation
    followups.addValidation('D2:D1000', ['Call', 'Email', 'Meeting', 'Text', 'Site Visit', 'Video Call']);
    followups.addValidation('I2:I1000', ['Sarah Johnson', 'Mike Davis', 'John Smith', 'Jane Doe']);
    followups.addValidation('J2:J1000', ['Planned', 'Scheduled', 'In Progress', 'Completed', 'Cancelled', 'No Show']);
    followups.addValidation('K2:K1000', ['Interested', 'Not Interested', 'No Response', 'Need More Info', 'Pending', 'Converted']);
    followups.addValidation('N2:N1000', ['Low', 'Medium', 'High', 'Very High', 'Urgent']);
    followups.addValidation('O2:O1000', ['Yes', 'No']);
    
    // Conditional formatting for Status
    followups.addConditionalFormat(2, 10, 1000, 10, {
      type: 'cell',
      condition: 'TEXT_EQUALS',
      value: 'Completed',
      format: { background: '#D1FAE5' }
    });
    
    followups.addConditionalFormat(2, 10, 1000, 10, {
      type: 'cell',
      condition: 'TEXT_EQUALS',
      value: 'Scheduled',
      format: { background: '#DBEAFE' }
    });
    
    // Priority highlighting
    followups.addConditionalFormat(2, 14, 1000, 14, {
      type: 'cell',
      condition: 'TEXT_EQUALS',
      value: 'Urgent',
      format: { background: '#FEE2E2', fontColor: '#DC2626' }
    });
    
    // Column widths
    followups.setColumnWidth(1, 100);
    followups.setColumnWidth(2, 80);
    followups.setColumnWidth(3, 150);
    for (let col = 4; col <= 15; col++) {
      followups.setColumnWidth(col, 100);
    }
  },
  
  buildLeadSources: function(sources) {
    // Title
    sources.setValue(1, 1, 'Lead Source Analytics');
    sources.merge(1, 1, 1, 12);
    sources.format(1, 1, 1, 12, {
      fontSize: 14,
      fontWeight: 'bold',
      horizontalAlignment: 'center',
      background: '#1F2937',
      fontColor: '#FFFFFF'
    });
    
    // Headers
    sources.setValue(3, 1, 'Source');
    sources.setValue(3, 2, 'Total Leads');
    sources.setValue(3, 3, 'New This Month');
    sources.setValue(3, 4, 'Qualified');
    sources.setValue(3, 5, 'Converted');
    sources.setValue(3, 6, 'Conversion Rate');
    sources.setValue(3, 7, 'Avg Score');
    sources.setValue(3, 8, 'Total Value');
    sources.setValue(3, 9, 'Avg Value');
    sources.setValue(3, 10, 'ROI');
    sources.setValue(3, 11, 'Cost');
    sources.setValue(3, 12, 'Cost per Lead');
    
    sources.format(3, 1, 3, 12, {
      fontWeight: 'bold',
      background: '#F3F4F6',
      border: true
    });
    
    // Source data
    const sourceList = [
      ['Website', 1500, 2500],
      ['Referral', 500, 1000],
      ['Social Media', 2000, 3500],
      ['Cold Call', 3000, 5000],
      ['Email Campaign', 1200, 2000],
      ['Partner', 800, 1500],
      ['Event', 2500, 4000],
      ['Direct Mail', 1800, 3000]
    ];
    
    sourceList.forEach((source, idx) => {
      const row = idx + 4;
      sources.setValue(row, 1, source[0]);
      sources.setValue(row, 11, source[1]); // Cost
      
      // Formulas
      sources.setFormula(row, 2, `=COUNTIF({{Leads}}!E:E,A${row})`);
      sources.setFormula(row, 3, `=COUNTIFS({{Leads}}!E:E,A${row},{{Leads}}!C:C,">="&EOMONTH(TODAY(),-1)+1)`);
      sources.setFormula(row, 4, `=COUNTIFS({{Leads}}!E:E,A${row},{{Leads}}!G:G,"Qualified")+COUNTIFS({{Leads}}!E:E,A${row},{{Leads}}!G:G,"Proposal")+COUNTIFS({{Leads}}!E:E,A${row},{{Leads}}!G:G,"Negotiation")+COUNTIFS({{Leads}}!E:E,A${row},{{Leads}}!G:G,"Converted")`);
      sources.setFormula(row, 5, `=COUNTIFS({{Leads}}!E:E,A${row},{{Leads}}!G:G,"Converted")`);
      sources.setFormula(row, 6, `=IF(B${row}=0,0,E${row}/B${row})`);
      sources.setFormula(row, 7, `=AVERAGEIF({{Leads}}!E:E,A${row},{{Leads}}!H:H)`);
      sources.setFormula(row, 8, `=SUMIFS({{Leads}}!K:K,{{Leads}}!E:E,A${row},{{Leads}}!G:G,"Converted")`);
      sources.setFormula(row, 9, `=IF(E${row}=0,0,H${row}/E${row})`);
      sources.setFormula(row, 10, `=IF(K${row}=0,0,(H${row}-K${row})/K${row})`);
      sources.setFormula(row, 12, `=IF(B${row}=0,0,K${row}/B${row})`);
    });
    
    // Totals row
    sources.setValue(12, 1, 'TOTAL');
    for (let col = 2; col <= 5; col++) {
      sources.setFormula(12, col, `=SUM(${String.fromCharCode(65 + col - 1)}4:${String.fromCharCode(65 + col - 1)}11)`);
    }
    sources.setFormula(12, 6, '=E12/B12');
    sources.setFormula(12, 7, '=AVERAGE(G4:G11)');
    sources.setFormula(12, 8, '=SUM(H4:H11)');
    sources.setFormula(12, 9, '=IF(E12=0,0,H12/E12)');
    sources.setFormula(12, 11, '=SUM(K4:K11)');
    sources.setFormula(12, 10, '=IF(K12=0,0,(H12-K12)/K12)');
    sources.setFormula(12, 12, '=IF(B12=0,0,K12/B12)');
    
    sources.format(12, 1, 12, 12, {
      fontWeight: 'bold',
      background: '#F3F4F6'
    });
    
    // Formatting
    sources.format(4, 6, 12, 6, { numberFormat: '0.0%' });
    sources.format(4, 7, 12, 7, { numberFormat: '0' });
    sources.format(4, 8, 12, 9, { numberFormat: '$#,##0' });
    sources.format(4, 10, 12, 10, { numberFormat: '0.0%' });
    sources.format(4, 11, 12, 12, { numberFormat: '$#,##0' });
    
    // Performance indicators
    sources.setValue(14, 1, 'PERFORMANCE INDICATORS');
    sources.merge(14, 1, 14, 6);
    sources.format(14, 1, 14, 6, {
      fontWeight: 'bold',
      background: '#1F2937',
      fontColor: '#FFFFFF'
    });
    
    const indicators = [
      ['Best Performing Source:', '=INDEX(A4:A11,MATCH(MAX(F4:F11),F4:F11,0))'],
      ['Most Cost Effective:', '=INDEX(A4:A11,MATCH(MIN(L4:L11),L4:L11,0))'],
      ['Highest Volume:', '=INDEX(A4:A11,MATCH(MAX(B4:B11),B4:B11,0))'],
      ['Best ROI:', '=INDEX(A4:A11,MATCH(MAX(J4:J11),J4:J11,0))'],
      ['Highest Quality (Score):', '=INDEX(A4:A11,MATCH(MAX(G4:G11),G4:G11,0))']
    ];
    
    indicators.forEach((indicator, idx) => {
      sources.setValue(15 + idx, 1, indicator[0]);
      sources.merge(15 + idx, 1, 15 + idx, 2);
      sources.setFormula(15 + idx, 3, indicator[1]);
      sources.merge(15 + idx, 3, 15 + idx, 4);
    });
    
    // Column widths
    sources.setColumnWidth(1, 120);
    for (let col = 2; col <= 12; col++) {
      sources.setColumnWidth(col, 100);
    }
  },
  
  buildCampaigns: function(campaigns) {
    // Headers
    const headers = [
      'Campaign ID', 'Campaign Name', 'Type', 'Source', 'Start Date',
      'End Date', 'Status', 'Budget', 'Spent', 'Remaining',
      'Leads Generated', 'Cost per Lead', 'Qualified Leads', 'Conversions',
      'Conversion Rate', 'Revenue', 'ROI', 'Notes'
    ];
    
    headers.forEach((header, idx) => {
      campaigns.setValue(1, idx + 1, header);
    });
    
    campaigns.format(1, 1, 1, headers.length, {
      fontWeight: 'bold',
      background: '#1F2937',
      fontColor: '#FFFFFF',
      border: true
    });
    
    // Sample campaigns
    const sampleCampaigns = [
      ['C001', 'Spring Home Buyers', 'Email', 'Email Campaign', '=TODAY()-60', '=TODAY()-30', 'Completed',
       5000, 4800, '=H2-I2', 45, '=I2/K2', 12, 5, '=N2/K2', 125000, '=(P2-I2)/I2',
       'Targeted first-time buyers'],
      ['C002', 'Luxury Property Showcase', 'Event', 'Event', '=TODAY()-45', '=TODAY()-15', 'Completed',
       8000, 7500, '=H3-I3', 28, '=I3/K3', 15, 8, '=N3/K3', 450000, '=(P3-I3)/I3',
       'High-end property exhibition'],
      ['C003', 'Social Media Blitz', 'Digital', 'Social Media', '=TODAY()-30', '=TODAY()+30', 'Active',
       3000, 1500, '=H4-I4', 65, '=I4/K4', 20, 3, '=N4/K4', 85000, '=(P4-I4)/I4',
       'Instagram and Facebook ads'],
      ['C004', 'Referral Rewards', 'Referral', 'Referral', '=TODAY()-90', '=TODAY()+90', 'Active',
       10000, 3000, '=H5-I5', 35, '=I5/K5', 25, 12, '=N5/K5', 380000, '=(P5-I5)/I5',
       'Client referral incentive program'],
      ['C005', 'Cold Calling Campaign', 'Outbound', 'Cold Call', '=TODAY()-15', '=TODAY()+15', 'Active',
       4000, 2000, '=H6-I6', 80, '=I6/K6', 15, 2, '=N6/K6', 55000, '=(P6-I6)/I6',
       'Targeted neighborhood outreach']
    ];
    
    sampleCampaigns.forEach((campaign, idx) => {
      campaign.forEach((value, col) => {
        if (typeof value === 'string' && value.startsWith('=')) {
          campaigns.setFormula(idx + 2, col + 1, value);
        } else {
          campaigns.setValue(idx + 2, col + 1, value);
        }
      });
    });
    
    // Data validation
    campaigns.addValidation('C2:C1000', ['Email', 'Event', 'Digital', 'Referral', 'Outbound', 'Content', 'Partnership']);
    campaigns.addValidation('G2:G1000', ['Planning', 'Active', 'Paused', 'Completed', 'Cancelled']);
    
    // Conditional formatting for Status
    campaigns.addConditionalFormat(2, 7, 1000, 7, {
      type: 'cell',
      condition: 'TEXT_EQUALS',
      value: 'Active',
      format: { background: '#D1FAE5' }
    });
    
    campaigns.addConditionalFormat(2, 7, 1000, 7, {
      type: 'cell',
      condition: 'TEXT_EQUALS',
      value: 'Completed',
      format: { background: '#E0E7FF' }
    });
    
    // ROI conditional formatting
    campaigns.addConditionalFormat(2, 17, 1000, 17, {
      type: 'gradient',
      minColor: '#FEE2E2',
      midColor: '#FEF3C7',
      maxColor: '#D1FAE5',
      minValue: '-0.5',
      midValue: '0.5',
      maxValue: '2'
    });
    
    // Format currency and percentage columns
    campaigns.format(2, 8, 1000, 10, { numberFormat: '$#,##0' });
    campaigns.format(2, 12, 1000, 12, { numberFormat: '$#,##0.00' });
    campaigns.format(2, 15, 1000, 15, { numberFormat: '0.0%' });
    campaigns.format(2, 16, 1000, 16, { numberFormat: '$#,##0' });
    campaigns.format(2, 17, 1000, 17, { numberFormat: '0.0%' });
    
    // Campaign summary
    campaigns.setValue(8, 1, 'CAMPAIGN SUMMARY');
    campaigns.merge(8, 1, 8, 6);
    campaigns.format(8, 1, 8, 6, {
      fontWeight: 'bold',
      background: '#F3F4F6'
    });
    
    campaigns.setValue(9, 1, 'Total Budget:');
    campaigns.setFormula(9, 2, '=SUM(H2:H6)');
    campaigns.format(9, 2, 9, 2, { numberFormat: '$#,##0' });
    
    campaigns.setValue(10, 1, 'Total Spent:');
    campaigns.setFormula(10, 2, '=SUM(I2:I6)');
    campaigns.format(10, 2, 10, 2, { numberFormat: '$#,##0' });
    
    campaigns.setValue(11, 1, 'Total Leads:');
    campaigns.setFormula(11, 2, '=SUM(K2:K6)');
    
    campaigns.setValue(12, 1, 'Total Conversions:');
    campaigns.setFormula(12, 2, '=SUM(N2:N6)');
    
    campaigns.setValue(13, 1, 'Total Revenue:');
    campaigns.setFormula(13, 2, '=SUM(P2:P6)');
    campaigns.format(13, 2, 13, 2, { numberFormat: '$#,##0' });
    
    campaigns.setValue(14, 1, 'Overall ROI:');
    campaigns.setFormula(14, 2, '=(B13-B10)/B10');
    campaigns.format(14, 2, 14, 2, { numberFormat: '0.0%' });
    
    // Column widths
    campaigns.setColumnWidth(1, 100);
    campaigns.setColumnWidth(2, 150);
    for (let col = 3; col <= 18; col++) {
      campaigns.setColumnWidth(col, 100);
    }
  },
  
  buildConversionAnalytics: function(conversion) {
    // Title
    conversion.setValue(1, 1, 'Conversion Analytics & Funnel Analysis');
    conversion.merge(1, 1, 1, 12);
    conversion.format(1, 1, 1, 12, {
      fontSize: 14,
      fontWeight: 'bold',
      horizontalAlignment: 'center',
      background: '#1F2937',
      fontColor: '#FFFFFF'
    });
    
    // Funnel Analysis
    conversion.setValue(3, 1, 'CONVERSION FUNNEL');
    conversion.merge(3, 1, 3, 6);
    conversion.format(3, 1, 3, 6, {
      fontWeight: 'bold',
      background: '#F3F4F6'
    });
    
    conversion.setValue(4, 1, 'Stage');
    conversion.setValue(4, 2, 'Count');
    conversion.setValue(4, 3, '% of Total');
    conversion.setValue(4, 4, 'Conversion to Next');
    conversion.setValue(4, 5, 'Avg Time in Stage');
    conversion.setValue(4, 6, 'Drop-off Rate');
    
    const funnelStages = [
      ['All Leads', '=COUNTIF({{Leads}}!B:B,"<>")', '=B5/$B$5', '', '', ''],
      ['Contacted', '=COUNTIF({{Leads}}!G:G,"Contacted")+COUNTIF({{Leads}}!G:G,"Qualified")+COUNTIF({{Leads}}!G:G,"Proposal")+COUNTIF({{Leads}}!G:G,"Negotiation")+COUNTIF({{Leads}}!G:G,"Converted")', '=B6/$B$5', '=B6/B5', '3 days', '=1-D6'],
      ['Qualified', '=COUNTIF({{Leads}}!G:G,"Qualified")+COUNTIF({{Leads}}!G:G,"Proposal")+COUNTIF({{Leads}}!G:G,"Negotiation")+COUNTIF({{Leads}}!G:G,"Converted")', '=B7/$B$5', '=B7/B6', '5 days', '=1-D7'],
      ['Proposal', '=COUNTIF({{Leads}}!G:G,"Proposal")+COUNTIF({{Leads}}!G:G,"Negotiation")+COUNTIF({{Leads}}!G:G,"Converted")', '=B8/$B$5', '=B8/B7', '7 days', '=1-D8'],
      ['Negotiation', '=COUNTIF({{Leads}}!G:G,"Negotiation")+COUNTIF({{Leads}}!G:G,"Converted")', '=B9/$B$5', '=B9/B8', '10 days', '=1-D9'],
      ['Converted', '=COUNTIF({{Leads}}!G:G,"Converted")', '=B10/$B$5', '=B10/B9', '2 days', '=1-D10']
    ];
    
    funnelStages.forEach((stage, idx) => {
      stage.forEach((value, col) => {
        if (typeof value === 'string' && value.startsWith('=')) {
          conversion.setFormula(5 + idx, col + 1, value);
        } else {
          conversion.setValue(5 + idx, col + 1, value);
        }
      });
    });
    
    conversion.format(4, 1, 10, 6, { border: true });
    conversion.format(5, 3, 10, 4, { numberFormat: '0.0%' });
    conversion.format(5, 6, 10, 6, { numberFormat: '0.0%' });
    
    // Conversion by Source
    conversion.setValue(3, 7, 'CONVERSION BY SOURCE');
    conversion.merge(3, 7, 3, 12);
    conversion.format(3, 7, 3, 12, {
      fontWeight: 'bold',
      background: '#F3F4F6'
    });
    
    conversion.setValue(4, 7, 'Source');
    conversion.setValue(4, 8, 'Leads');
    conversion.setValue(4, 9, 'Conversions');
    conversion.setValue(4, 10, 'Rate');
    conversion.setValue(4, 11, 'Avg Days');
    conversion.setValue(4, 12, 'Value');
    
    const sources = ['Website', 'Referral', 'Social Media', 'Cold Call', 'Email Campaign'];
    sources.forEach((source, idx) => {
      conversion.setValue(5 + idx, 7, source);
      conversion.setFormula(5 + idx, 8, `=COUNTIF({{Leads}}!E:E,G${5 + idx})`);
      conversion.setFormula(5 + idx, 9, `=COUNTIFS({{Leads}}!E:E,G${5 + idx},{{Leads}}!G:G,"Converted")`);
      conversion.setFormula(5 + idx, 10, `=IF(H${5 + idx}=0,0,I${5 + idx}/H${5 + idx})`);
      conversion.setValue(5 + idx, 11, 15 + idx * 3);
      conversion.setFormula(5 + idx, 12, `=SUMIFS({{Leads}}!K:K,{{Leads}}!E:E,G${5 + idx},{{Leads}}!G:G,"Converted")`);
    });
    
    conversion.format(4, 7, 9, 12, { border: true });
    conversion.format(5, 10, 9, 10, { numberFormat: '0.0%' });
    conversion.format(5, 12, 9, 12, { numberFormat: '$#,##0' });
    
    // Time to Conversion Analysis
    conversion.setValue(12, 1, 'TIME TO CONVERSION ANALYSIS');
    conversion.merge(12, 1, 12, 6);
    conversion.format(12, 1, 12, 6, {
      fontWeight: 'bold',
      background: '#F3F4F6'
    });
    
    const timeRanges = [
      ['0-7 days', 8, 35],
      ['8-14 days', 12, 25],
      ['15-30 days', 15, 20],
      ['31-60 days', 10, 15],
      ['60+ days', 5, 5]
    ];
    
    conversion.setValue(13, 1, 'Time Range');
    conversion.setValue(13, 2, 'Conversions');
    conversion.setValue(13, 3, '% of Total');
    conversion.setValue(13, 4, 'Avg Value');
    conversion.setValue(13, 5, 'Success Rate');
    
    timeRanges.forEach((range, idx) => {
      conversion.setValue(14 + idx, 1, range[0]);
      conversion.setValue(14 + idx, 2, range[1]);
      conversion.setFormula(14 + idx, 3, `=B${14 + idx}/SUM($B$14:$B$18)`);
      conversion.setValue(14 + idx, 4, 250000 + idx * 50000);
      conversion.setValue(14 + idx, 5, range[2] + '%');
    });
    
    conversion.format(13, 1, 18, 5, { border: true });
    conversion.format(14, 3, 18, 3, { numberFormat: '0.0%' });
    conversion.format(14, 4, 18, 4, { numberFormat: '$#,##0' });
    
    // Lost Lead Analysis
    conversion.setValue(12, 7, 'LOST LEAD ANALYSIS');
    conversion.merge(12, 7, 12, 12);
    conversion.format(12, 7, 12, 12, {
      fontWeight: 'bold',
      background: '#F3F4F6'
    });
    
    conversion.setValue(13, 7, 'Lost Reason');
    conversion.setValue(13, 9, 'Count');
    conversion.setValue(13, 10, 'Percentage');
    conversion.setValue(13, 11, 'Avg Score');
    conversion.merge(13, 11, 13, 12);
    
    const lostReasons = [
      ['Price too high', 25, 45],
      ['Found another property', 18, 60],
      ['Not ready to buy', 15, 30],
      ['Financing issues', 12, 55],
      ['Location not suitable', 10, 40]
    ];
    
    lostReasons.forEach((reason, idx) => {
      conversion.setValue(14 + idx, 7, reason[0]);
      conversion.merge(14 + idx, 7, 14 + idx, 8);
      conversion.setValue(14 + idx, 9, reason[1]);
      conversion.setFormula(14 + idx, 10, `=I${14 + idx}/SUM($I$14:$I$18)`);
      conversion.setValue(14 + idx, 11, reason[2]);
      conversion.merge(14 + idx, 11, 14 + idx, 12);
    });
    
    conversion.format(13, 7, 18, 12, { border: true });
    conversion.format(14, 10, 18, 10, { numberFormat: '0.0%' });
    
    // Column widths
    for (let col = 1; col <= 12; col++) {
      conversion.setColumnWidth(col, 100);
    }
  },
  
  buildLeadReports: function(reports) {
    // Title
    reports.setValue(1, 1, 'Lead Pipeline Reports');
    reports.merge(1, 1, 1, 12);
    reports.format(1, 1, 1, 12, {
      fontSize: 14,
      fontWeight: 'bold',
      horizontalAlignment: 'center',
      background: '#1F2937',
      fontColor: '#FFFFFF'
    });
    
    // Monthly Performance Report
    reports.setValue(3, 1, 'MONTHLY PERFORMANCE REPORT');
    reports.merge(3, 1, 3, 12);
    reports.format(3, 1, 3, 12, {
      fontWeight: 'bold',
      background: '#F3F4F6'
    });
    
    reports.setValue(4, 1, 'Month');
    reports.setValue(4, 2, 'New Leads');
    reports.setValue(4, 3, 'Contacted');
    reports.setValue(4, 4, 'Qualified');
    reports.setValue(4, 5, 'Proposals');
    reports.setValue(4, 6, 'Conversions');
    reports.setValue(4, 7, 'Conv Rate');
    reports.setValue(4, 8, 'Revenue');
    reports.setValue(4, 9, 'Avg Deal');
    reports.setValue(4, 10, 'Cost');
    reports.setValue(4, 11, 'ROI');
    reports.setValue(4, 12, 'Target vs Actual');
    
    // Sample monthly data
    const months = ['January', 'February', 'March', 'April', 'May', 'June'];
    const monthlyData = [
      [85, 70, 45, 20, 8, '9.4%', 325000, 40625, 12000, '2608%', '80%'],
      [92, 78, 50, 25, 10, '10.9%', 425000, 42500, 13500, '3048%', '95%'],
      [78, 65, 40, 18, 7, '9.0%', 280000, 40000, 11000, '2445%', '70%'],
      [105, 88, 55, 28, 12, '11.4%', 520000, 43333, 15000, '3367%', '110%'],
      [98, 82, 52, 26, 11, '11.2%', 485000, 44091, 14000, '3364%', '105%'],
      [110, 95, 60, 30, 13, '11.8%', 565000, 43462, 16000, '3431%', '115%']
    ];
    
    months.forEach((month, idx) => {
      reports.setValue(5 + idx, 1, month);
      monthlyData[idx].forEach((value, col) => {
        if (col === 5 || col === 9 || col === 10) {
          reports.setValue(5 + idx, col + 2, value);
        } else if (col === 6 || col === 7) {
          reports.setValue(5 + idx, col + 2, value);
          reports.format(5 + idx, col + 2, 5 + idx, col + 2, { numberFormat: col === 6 ? '0.0%' : '$#,##0' });
        } else {
          reports.setValue(5 + idx, col + 2, value);
        }
      });
    });
    
    // Totals row
    reports.setValue(11, 1, 'TOTAL/AVG');
    reports.setFormula(11, 2, '=SUM(B5:B10)');
    reports.setFormula(11, 3, '=SUM(C5:C10)');
    reports.setFormula(11, 4, '=SUM(D5:D10)');
    reports.setFormula(11, 5, '=SUM(E5:E10)');
    reports.setFormula(11, 6, '=SUM(F5:F10)');
    reports.setFormula(11, 7, '=F11/B11');
    reports.setFormula(11, 8, '=SUM(H5:H10)');
    reports.setFormula(11, 9, '=H11/F11');
    reports.setFormula(11, 10, '=SUM(J5:J10)');
    reports.setFormula(11, 11, '=(H11-J11)/J11');
    reports.setFormula(11, 12, '=AVERAGE(L5:L10)');
    
    reports.format(11, 1, 11, 12, {
      fontWeight: 'bold',
      background: '#E5E7EB'
    });
    
    reports.format(5, 7, 11, 7, { numberFormat: '0.0%' });
    reports.format(5, 8, 11, 9, { numberFormat: '$#,##0' });
    reports.format(5, 10, 11, 10, { numberFormat: '$#,##0' });
    reports.format(5, 11, 11, 11, { numberFormat: '0%' });
    reports.format(5, 12, 11, 12, { numberFormat: '0%' });
    
    // Agent Performance Report
    reports.setValue(13, 1, 'AGENT PERFORMANCE REPORT');
    reports.merge(13, 1, 13, 8);
    reports.format(13, 1, 13, 8, {
      fontWeight: 'bold',
      background: '#F3F4F6'
    });
    
    reports.setValue(14, 1, 'Agent');
    reports.setValue(14, 2, 'Total Leads');
    reports.setValue(14, 3, 'Active');
    reports.setValue(14, 4, 'Conversions');
    reports.setValue(14, 5, 'Conv Rate');
    reports.setValue(14, 6, 'Revenue');
    reports.setValue(14, 7, 'Avg Response');
    reports.setValue(14, 8, 'Score');
    
    const agents = [
      ['Sarah Johnson', 145, 42, 28, '19.3%', 1250000, '2.5 hrs', 92],
      ['Mike Davis', 132, 38, 24, '18.2%', 980000, '3.1 hrs', 88],
      ['John Smith', 98, 25, 15, '15.3%', 650000, '4.2 hrs', 75],
      ['Jane Doe', 88, 20, 12, '13.6%', 480000, '3.8 hrs', 70]
    ];
    
    agents.forEach((agent, idx) => {
      agent.forEach((value, col) => {
        if (col === 4) {
          reports.setValue(15 + idx, col + 1, value);
          reports.format(15 + idx, col + 1, 15 + idx, col + 1, { numberFormat: '0.0%' });
        } else if (col === 5) {
          reports.setValue(15 + idx, col + 1, value);
          reports.format(15 + idx, col + 1, 15 + idx, col + 1, { numberFormat: '$#,##0' });
        } else {
          reports.setValue(15 + idx, col + 1, value);
        }
      });
    });
    
    reports.format(14, 1, 18, 8, { border: true });
    
    // Quick Actions
    reports.setValue(13, 9, 'QUICK ACTIONS');
    reports.merge(13, 9, 13, 12);
    reports.format(13, 9, 13, 12, {
      fontWeight: 'bold',
      background: '#F3F4F6'
    });
    
    const actions = [
      'Export to PDF',
      'Email Report',
      'Schedule Meeting',
      'Update Targets'
    ];
    
    actions.forEach((action, idx) => {
      reports.setValue(14 + idx, 9, '→');
      reports.setValue(14 + idx, 10, action);
      reports.merge(14 + idx, 10, 14 + idx, 12);
      reports.format(14 + idx, 9, 14 + idx, 12, {
        background: '#DBEAFE',
        border: true
      });
    });
    
    // Key Metrics Summary
    reports.setValue(20, 1, 'KEY METRICS SUMMARY');
    reports.merge(20, 1, 20, 12);
    reports.format(20, 1, 20, 12, {
      fontWeight: 'bold',
      background: '#1F2937',
      fontColor: '#FFFFFF'
    });
    
    const metrics = [
      ['Lead Quality Score:', '=AVERAGE({{Lead Scoring}}!C:C)', 'Target: 65'],
      ['Pipeline Value:', '=SUM({{Leads}}!K:K)', 'Target: $5M'],
      ['Avg Days to Close:', '=AVERAGE({{Conversion}}!K5:K9)', 'Target: 30'],
      ['Cost per Acquisition:', '=SUM({{Sources}}!K:K)/COUNTIF({{Leads}}!G:G,"Converted")', 'Target: $500']
    ];
    
    metrics.forEach((metric, idx) => {
      reports.setValue(21 + idx, 1, metric[0]);
      reports.merge(21 + idx, 1, 21 + idx, 3);
      reports.setFormula(21 + idx, 4, metric[1]);
      reports.merge(21 + idx, 4, 21 + idx, 6);
      reports.setValue(21 + idx, 7, metric[2]);
      reports.merge(21 + idx, 7, 21 + idx, 9);
    });
    
    // Column widths
    reports.setColumnWidth(1, 120);
    for (let col = 2; col <= 12; col++) {
      reports.setColumnWidth(col, 90);
    }
  }
};

// TemplateBuilderPro and SheetHelper classes are now in TemplateBuilder.js
