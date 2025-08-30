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
  }
};

/**
 * Enhanced Template Builder with Professional Features
 */
class TemplateBuilderPro {
  constructor(spreadsheet, isPreview) {
    this.spreadsheet = spreadsheet;
    this.isPreview = isPreview;
    this.sheets = {};
    this.errors = [];
  }
  
  createSheets(names) {
    names.forEach(name => {
      const sheetName = this.isPreview ? `[PREVIEW] ${name}` : name;
      try {
        const sheet = this.spreadsheet.insertSheet(sheetName);
        this.sheets[name] = sheet;
      } catch (e) {
        this.errors.push(`Failed to create sheet ${name}: ${e.toString()}`);
      }
    });
    return this.sheets;
  }
  
  sheet(name) {
    const sheet = this.sheets[name];
    const self = this;
    
    return {
      setValue: (row, col, value) => {
        try {
          sheet.getRange(row, col).setValue(value);
        } catch (e) {
          self.errors.push(`Error setting value at ${name}[${row},${col}]: ${e.toString()}`);
        }
      },
      
      setFormula: (row, col, formula) => {
        try {
          // Replace sheet name placeholders
          let resolvedFormula = formula;
          for (const [key, s] of Object.entries(self.sheets)) {
            const actualName = self.isPreview ? `[PREVIEW] ${key}` : key;
            resolvedFormula = resolvedFormula.replace(new RegExp(`{{${key}}}`, 'g'), `'${actualName}'`);
          }
          sheet.getRange(row, col).setFormula(resolvedFormula);
        } catch (e) {
          self.errors.push(`Error setting formula at ${name}[${row},${col}]: ${e.toString()}`);
        }
      },
      
      setRangeValues: (row, col, values) => {
        try {
          const numRows = values.length;
          const numCols = values[0].length;
          sheet.getRange(row, col, numRows, numCols).setValues(values);
        } catch (e) {
          self.errors.push(`Error setting range at ${name}[${row},${col}]: ${e.toString()}`);
        }
      },
      
      format: (startRow, startCol, endRow, endCol, formatting) => {
        try {
          const range = sheet.getRange(startRow, startCol, endRow - startRow + 1, endCol - startCol + 1);
          if (formatting.background) range.setBackground(formatting.background);
          if (formatting.fontColor) range.setFontColor(formatting.fontColor);
          if (formatting.fontSize) range.setFontSize(formatting.fontSize);
          if (formatting.fontWeight) range.setFontWeight(formatting.fontWeight);
          if (formatting.fontFamily) range.setFontFamily(formatting.fontFamily);
          if (formatting.horizontalAlignment) range.setHorizontalAlignment(formatting.horizontalAlignment);
          if (formatting.verticalAlignment) range.setVerticalAlignment(formatting.verticalAlignment);
          if (formatting.numberFormat) range.setNumberFormat(formatting.numberFormat);
          if (formatting.border) {
            range.setBorder(true, true, true, true, true, true, '#D1D5DB', SpreadsheetApp.BorderStyle.SOLID);
          }
        } catch (e) {
          self.errors.push(`Error formatting ${name}[${startRow},${startCol}]: ${e.toString()}`);
        }
      },
      
      merge: (startRow, startCol, endRow, endCol) => {
        try {
          sheet.getRange(startRow, startCol, endRow - startRow + 1, endCol - startCol + 1).merge();
        } catch (e) {
          self.errors.push(`Error merging ${name}[${startRow},${startCol}]: ${e.toString()}`);
        }
      },
      
      setColumnWidth: (col, width) => {
        try {
          sheet.setColumnWidth(col, width);
        } catch (e) {
          self.errors.push(`Error setting column width ${name}[${col}]: ${e.toString()}`);
        }
      },
      
      setRowHeight: (row, height) => {
        try {
          sheet.setRowHeight(row, height);
        } catch (e) {
          self.errors.push(`Error setting row height ${name}[${row}]: ${e.toString()}`);
        }
      },
      
      addValidation: (rangeA1, values) => {
        try {
          const rule = SpreadsheetApp.newDataValidation()
            .requireValueInList(values, true)
            .build();
          sheet.getRange(rangeA1).setDataValidation(rule);
        } catch (e) {
          self.errors.push(`Error adding validation ${name}[${rangeA1}]: ${e.toString()}`);
        }
      },
      
      addConditionalFormat: (startRow, startCol, endRow, endCol, rule) => {
        try {
          const range = sheet.getRange(startRow, startCol, endRow - startRow + 1, endCol - startCol + 1);
          let formatRule;
          
          if (rule.type === 'gradient') {
            formatRule = SpreadsheetApp.newConditionalFormatRule()
              .setGradientMinpoint(rule.minColor, SpreadsheetApp.InterpolationType.NUMBER, rule.minValue.toString())
              .setGradientMidpoint(rule.midColor, SpreadsheetApp.InterpolationType.NUMBER, rule.midValue.toString())
              .setGradientMaxpoint(rule.maxColor, SpreadsheetApp.InterpolationType.NUMBER, rule.maxValue.toString())
              .setRanges([range])
              .build();
          } else if (rule.type === 'text') {
            formatRule = SpreadsheetApp.newConditionalFormatRule()
              .whenTextEqualTo(rule.condition)
              .setBackground(rule.background)
              .setFontColor(rule.fontColor)
              .setRanges([range])
              .build();
          }
          
          if (formatRule) {
            const rules = sheet.getConditionalFormatRules();
            rules.push(formatRule);
            sheet.setConditionalFormatRules(rules);
          }
        } catch (e) {
          self.errors.push(`Error adding conditional format ${name}[${startRow},${startCol}]: ${e.toString()}`);
        }
      }
    };
  }
  
  complete() {
    // Move Dashboard to first position
    if (this.sheets['Dashboard']) {
      this.spreadsheet.setActiveSheet(this.sheets['Dashboard']);
      this.spreadsheet.moveActiveSheet(1);
    }
    
    return {
      success: this.errors.length === 0,
      sheets: Object.keys(this.sheets),
      errors: this.errors
    };
  }
}