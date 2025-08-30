/**
 * Professional Marketing Templates Suite
 * Complete templates for marketing professionals and agencies
 */

var MarketingTemplate = {
  /**
   * Build comprehensive Marketing template based on type
   */
  build: function(spreadsheet, templateType = 'campaign-tracker', isPreview = false) {
    const builder = new TemplateBuilderPro(spreadsheet, isPreview);
    
    switch(templateType) {
      case 'campaign-tracker':
        return this.buildCampaignTracker(builder);
      case 'content-calendar':
        return this.buildContentCalendar(builder);
      case 'social-media':
        return this.buildSocialMediaManager(builder);
      case 'email-marketing':
        return this.buildEmailMarketing(builder);
      case 'seo-tracker':
        return this.buildSEOTracker(builder);
      case 'analytics-dashboard':
        return this.buildAnalyticsDashboard(builder);
      case 'lead-scoring':
        return this.buildLeadScoring(builder);
      case 'budget-planner':
        return this.buildBudgetPlanner(builder);
      case 'competitor-analysis':
        return this.buildCompetitorAnalysis(builder);
      case 'roi-calculator':
        return this.buildROICalculator(builder);
      default:
        return this.buildCampaignTracker(builder);
    }
  },
  
  /**
   * Campaign Tracker - Comprehensive campaign management
   */
  buildCampaignTracker: function(builder) {
    const sheets = builder.createSheets([
      'Dashboard',
      'Campaigns',
      'Performance',
      'Channels',
      'Budget',
      'Creative',
      'Results',
      'Reports'
    ]);
    
    this.buildCampaignDashboard(sheets.Dashboard);
    this.buildCampaignsSheet(sheets.Campaigns);
    this.buildPerformanceSheet(sheets.Performance);
    this.buildChannelsSheet(sheets.Channels);
    this.buildBudgetSheet(sheets.Budget);
    this.buildCreativeSheet(sheets.Creative);
    this.buildResultsSheet(sheets.Results);
    this.buildReportsSheet(sheets.Reports);
    
    return builder.complete();
  },
  
  buildCampaignDashboard: function(dash) {
    // Title
    dash.setValue(1, 1, 'Marketing Campaign Command Center');
    dash.merge(1, 1, 1, 12);
    dash.format(1, 1, 1, 12, {
      fontSize: 18,
      fontWeight: 'bold',
      horizontalAlignment: 'center',
      background: '#7C3AED',
      fontColor: '#FFFFFF'
    });
    
    // Date Range Selector
    dash.setValue(2, 1, 'Period:');
    dash.setValue(2, 2, 'Last 30 Days');
    dash.addValidation('B2', ['Today', 'Last 7 Days', 'Last 30 Days', 'Last Quarter', 'YTD', 'All Time']);
    
    // KPI Cards Row 1
    dash.setValue(4, 1, 'CAMPAIGN METRICS');
    dash.merge(4, 1, 4, 12);
    dash.format(4, 1, 4, 12, {
      fontSize: 12,
      fontWeight: 'bold',
      background: '#F3E8FF'
    });
    
    // Primary KPIs
    const kpis = [
      { row: 5, col: 1, title: 'Active Campaigns', formula: '=COUNTIF({{Campaigns}}!G:G,"Active")', sub: '=TEXT((A6-10)/10,"+0;-0")&" vs target"' },
      { row: 5, col: 4, title: 'Total Spend', formula: '=SUM({{Budget}}!E:E)', sub: '=TEXT(D6/SUM({{Budget}}!D:D),"0%")&" of budget"' },
      { row: 5, col: 7, title: 'Total Reach', formula: '=SUM({{Performance}}!D:D)', sub: '=TEXT(G6,"#,##0")&" people"' },
      { row: 5, col: 10, title: 'Avg ROI', formula: '=AVERAGE({{Results}}!H:H)', sub: '=IF(J6>2,"🟢 Excellent",IF(J6>1,"🟡 Good","🔴 Poor"))' }
    ];
    
    kpis.forEach(kpi => {
      dash.setValue(kpi.row, kpi.col, kpi.title);
      dash.merge(kpi.row, kpi.col, kpi.row, kpi.col + 2);
      dash.setFormula(kpi.row + 1, kpi.col, kpi.formula);
      dash.merge(kpi.row + 1, kpi.col, kpi.row + 1, kpi.col + 2);
      dash.setFormula(kpi.row + 2, kpi.col, kpi.sub);
      dash.merge(kpi.row + 2, kpi.col, kpi.row + 2, kpi.col + 2);
      
      dash.format(kpi.row, kpi.col, kpi.row, kpi.col + 2, {
        background: '#E9D5FF',
        fontWeight: 'bold',
        fontSize: 11
      });
      dash.format(kpi.row + 1, kpi.col, kpi.row + 1, kpi.col + 2, {
        fontSize: 20,
        fontWeight: 'bold',
        horizontalAlignment: 'center'
      });
    });
    
    // Format specific KPIs
    dash.format(6, 4, 6, 6, { numberFormat: '$#,##0' });
    dash.format(6, 7, 6, 9, { numberFormat: '#,##0' });
    dash.format(6, 10, 6, 12, { numberFormat: '0.0x' });
    
    // Channel Performance
    dash.setValue(9, 1, 'CHANNEL PERFORMANCE');
    dash.merge(9, 1, 9, 6);
    dash.format(9, 1, 9, 6, {
      fontSize: 12,
      fontWeight: 'bold',
      background: '#F3E8FF'
    });
    
    const channels = [
      ['Channel', 'Campaigns', 'Spend', 'Conversions', 'CPA', 'ROI'],
      ['Social Media', '=COUNTIF({{Campaigns}}!E:E,"Social")', '=SUMIF({{Campaigns}}!E:E,"Social",{{Campaigns}}!I:I)', '=SUMIF({{Campaigns}}!E:E,"Social",{{Performance}}!G:G)', '=C11/D11', '=(D11*50-C11)/C11'],
      ['Email', '=COUNTIF({{Campaigns}}!E:E,"Email")', '=SUMIF({{Campaigns}}!E:E,"Email",{{Campaigns}}!I:I)', '=SUMIF({{Campaigns}}!E:E,"Email",{{Performance}}!G:G)', '=C12/D12', '=(D12*50-C12)/C12'],
      ['Search', '=COUNTIF({{Campaigns}}!E:E,"Search")', '=SUMIF({{Campaigns}}!E:E,"Search",{{Campaigns}}!I:I)', '=SUMIF({{Campaigns}}!E:E,"Search",{{Performance}}!G:G)', '=C13/D13', '=(D13*50-C13)/C13'],
      ['Display', '=COUNTIF({{Campaigns}}!E:E,"Display")', '=SUMIF({{Campaigns}}!E:E,"Display",{{Campaigns}}!I:I)', '=SUMIF({{Campaigns}}!E:E,"Display",{{Performance}}!G:G)', '=C14/D14', '=(D14*50-C14)/C14'],
      ['Content', '=COUNTIF({{Campaigns}}!E:E,"Content")', '=SUMIF({{Campaigns}}!E:E,"Content",{{Campaigns}}!I:I)', '=SUMIF({{Campaigns}}!E:E,"Content",{{Performance}}!G:G)', '=C15/D15', '=(D15*50-C15)/C15']
    ];
    
    channels.forEach((row, idx) => {
      row.forEach((cell, col) => {
        if (typeof cell === 'string' && cell.startsWith('=')) {
          dash.setFormula(10 + idx, col + 1, cell);
        } else {
          dash.setValue(10 + idx, col + 1, cell);
        }
      });
    });
    
    dash.format(10, 1, 10, 6, {
      fontWeight: 'bold',
      background: '#E5E7EB'
    });
    dash.format(11, 3, 15, 3, { numberFormat: '$#,##0' });
    dash.format(11, 5, 15, 5, { numberFormat: '$#,##0.00' });
    dash.format(11, 6, 15, 6, { numberFormat: '0.0x' });
    
    // Campaign Status
    dash.setValue(9, 7, 'CAMPAIGN STATUS');
    dash.merge(9, 7, 9, 12);
    dash.format(9, 7, 9, 12, {
      fontSize: 12,
      fontWeight: 'bold',
      background: '#F3E8FF'
    });
    
    const statuses = [
      ['Status', 'Count', 'Budget', 'Spent', '% Used', 'Visual'],
      ['Active', '=COUNTIF({{Campaigns}}!G:G,"Active")', '=SUMIF({{Campaigns}}!G:G,"Active",{{Campaigns}}!H:H)', '=SUMIF({{Campaigns}}!G:G,"Active",{{Campaigns}}!I:I)', '=J11/I11', '=REPT("█",ROUND(K11*10,0))'],
      ['Paused', '=COUNTIF({{Campaigns}}!G:G,"Paused")', '=SUMIF({{Campaigns}}!G:G,"Paused",{{Campaigns}}!H:H)', '=SUMIF({{Campaigns}}!G:G,"Paused",{{Campaigns}}!I:I)', '=J12/I12', '=REPT("█",ROUND(K12*10,0))'],
      ['Scheduled', '=COUNTIF({{Campaigns}}!G:G,"Scheduled")', '=SUMIF({{Campaigns}}!G:G,"Scheduled",{{Campaigns}}!H:H)', '=SUMIF({{Campaigns}}!G:G,"Scheduled",{{Campaigns}}!I:I)', '=J13/I13', '=REPT("█",ROUND(K13*10,0))'],
      ['Completed', '=COUNTIF({{Campaigns}}!G:G,"Completed")', '=SUMIF({{Campaigns}}!G:G,"Completed",{{Campaigns}}!H:H)', '=SUMIF({{Campaigns}}!G:G,"Completed",{{Campaigns}}!I:I)', '=J14/I14', '=REPT("█",ROUND(K14*10,0))']
    ];
    
    statuses.forEach((row, idx) => {
      row.forEach((cell, col) => {
        if (typeof cell === 'string' && cell.startsWith('=')) {
          dash.setFormula(10 + idx, col + 7, cell);
        } else {
          dash.setValue(10 + idx, col + 7, cell);
        }
      });
    });
    
    dash.format(10, 7, 10, 12, {
      fontWeight: 'bold',
      background: '#E5E7EB'
    });
    dash.format(11, 9, 14, 10, { numberFormat: '$#,##0' });
    dash.format(11, 11, 14, 11, { numberFormat: '0%' });
    
    // Top Performing Campaigns
    dash.setValue(17, 1, 'TOP PERFORMING CAMPAIGNS');
    dash.merge(17, 1, 17, 12);
    dash.format(17, 1, 17, 12, {
      fontSize: 12,
      fontWeight: 'bold',
      background: '#F3E8FF'
    });
    
    dash.setValue(18, 1, 'Campaign');
    dash.merge(18, 1, 18, 3);
    dash.setValue(18, 4, 'Channel');
    dash.setValue(18, 5, 'Spend');
    dash.setValue(18, 6, 'Reach');
    dash.setValue(18, 7, 'Clicks');
    dash.setValue(18, 8, 'Conversions');
    dash.setValue(18, 9, 'CPA');
    dash.setValue(18, 10, 'ROI');
    dash.setValue(18, 11, 'Score');
    dash.merge(18, 11, 18, 12);
    
    // Sample top campaigns
    for (let i = 0; i < 5; i++) {
      dash.setFormula(19 + i, 1, `=INDEX({{Campaigns}}!B:B,LARGE(IF({{Campaigns}}!G:G="Active",ROW({{Campaigns}}!G:G)),${i+1}))`);
      dash.merge(19 + i, 1, 19 + i, 3);
      dash.setFormula(19 + i, 4, `=INDEX({{Campaigns}}!E:E,LARGE(IF({{Campaigns}}!G:G="Active",ROW({{Campaigns}}!G:G)),${i+1}))`);
      dash.setFormula(19 + i, 5, `=INDEX({{Campaigns}}!I:I,LARGE(IF({{Campaigns}}!G:G="Active",ROW({{Campaigns}}!G:G)),${i+1}))`);
      dash.setValue(19 + i, 6, 50000 + Math.random() * 100000);
      dash.setValue(19 + i, 7, 2000 + Math.random() * 5000);
      dash.setValue(19 + i, 8, 50 + Math.random() * 200);
      dash.setFormula(19 + i, 9, `=E${19+i}/H${19+i}`);
      dash.setFormula(19 + i, 10, `=(H${19+i}*50-E${19+i})/E${19+i}`);
      dash.setFormula(19 + i, 11, `=IF(J${19+i}>3,"⭐⭐⭐⭐⭐",IF(J${19+i}>2,"⭐⭐⭐⭐",IF(J${19+i}>1,"⭐⭐⭐","⭐⭐")))`);
      dash.merge(19 + i, 11, 19 + i, 12);
    }
    
    dash.format(18, 1, 18, 12, {
      fontWeight: 'bold',
      background: '#E5E7EB'
    });
    dash.format(19, 5, 23, 5, { numberFormat: '$#,##0' });
    dash.format(19, 6, 23, 8, { numberFormat: '#,##0' });
    dash.format(19, 9, 23, 9, { numberFormat: '$#,##0' });
    dash.format(19, 10, 23, 10, { numberFormat: '0.0x' });
    
    // Conversion Funnel
    dash.setValue(25, 1, 'CONVERSION FUNNEL');
    dash.merge(25, 1, 25, 6);
    dash.format(25, 1, 25, 6, {
      fontSize: 12,
      fontWeight: 'bold',
      background: '#F3E8FF'
    });
    
    const funnel = [
      ['Stage', 'Count', '% of Total', 'Conversion', 'Drop-off', 'Visual'],
      ['Impressions', '=SUM({{Performance}}!C:C)', '100%', '', '', '=REPT("▬",10)'],
      ['Clicks', '=SUM({{Performance}}!E:E)', '=B27/B26', '=B27/B26', '=1-C27', '=REPT("▬",ROUND(C27*10,0))'],
      ['Leads', '=SUM({{Performance}}!F:F)', '=B28/B26', '=B28/B27', '=1-D28', '=REPT("▬",ROUND(C28*10,0))'],
      ['Qualified', '=SUM({{Performance}}!G:G)*0.6', '=B29/B26', '=B29/B28', '=1-D29', '=REPT("▬",ROUND(C29*10,0))'],
      ['Customers', '=SUM({{Performance}}!G:G)', '=B30/B26', '=B30/B29', '=1-D30', '=REPT("▬",ROUND(C30*10,0))']
    ];
    
    funnel.forEach((row, idx) => {
      row.forEach((cell, col) => {
        if (typeof cell === 'string' && cell.startsWith('=')) {
          dash.setFormula(26 + idx, col + 1, cell);
        } else {
          dash.setValue(26 + idx, col + 1, cell);
        }
      });
    });
    
    dash.format(26, 1, 26, 6, {
      fontWeight: 'bold',
      background: '#E5E7EB'
    });
    dash.format(27, 2, 31, 2, { numberFormat: '#,##0' });
    dash.format(27, 3, 31, 5, { numberFormat: '0.0%' });
    
    // Budget Summary
    dash.setValue(25, 7, 'BUDGET SUMMARY');
    dash.merge(25, 7, 25, 12);
    dash.format(25, 7, 25, 12, {
      fontSize: 12,
      fontWeight: 'bold',
      background: '#F3E8FF'
    });
    
    dash.setValue(26, 7, 'Total Budget:');
    dash.setFormula(26, 9, '=SUM({{Budget}}!D:D)');
    dash.merge(26, 9, 26, 10);
    dash.format(26, 9, 26, 10, { numberFormat: '$#,##0' });
    
    dash.setValue(27, 7, 'Allocated:');
    dash.setFormula(27, 9, '=SUM({{Campaigns}}!H:H)');
    dash.merge(27, 9, 27, 10);
    dash.format(27, 9, 27, 10, { numberFormat: '$#,##0' });
    
    dash.setValue(28, 7, 'Spent:');
    dash.setFormula(28, 9, '=SUM({{Campaigns}}!I:I)');
    dash.merge(28, 9, 28, 10);
    dash.format(28, 9, 28, 10, { numberFormat: '$#,##0' });
    
    dash.setValue(29, 7, 'Remaining:');
    dash.setFormula(29, 9, '=I26-I28');
    dash.merge(29, 9, 29, 10);
    dash.format(29, 9, 29, 10, { numberFormat: '$#,##0' });
    
    dash.setValue(30, 7, 'Burn Rate:');
    dash.setFormula(30, 9, '=I28/30');
    dash.merge(30, 9, 30, 10);
    dash.format(30, 9, 30, 10, { numberFormat: '$#,##0"/day"' });
    
    dash.setValue(31, 7, 'Days Left:');
    dash.setFormula(31, 9, '=I29/I30');
    dash.merge(31, 9, 31, 10);
    dash.format(31, 9, 31, 10, { numberFormat: '0" days"' });
    
    // Column widths
    dash.setColumnWidth(1, 100);
    for (let col = 2; col <= 12; col++) {
      dash.setColumnWidth(col, 85);
    }
  },
  
  buildCampaignsSheet: function(campaigns) {
    // Headers
    const headers = [
      'Campaign ID', 'Campaign Name', 'Type', 'Objective', 'Channel', 
      'Start Date', 'End Date', 'Status', 'Budget', 'Spent', 'Remaining',
      'Target Audience', 'Creative', 'Landing Page', 'Owner', 'Notes'
    ];
    
    headers.forEach((header, idx) => {
      campaigns.setValue(1, idx + 1, header);
    });
    
    campaigns.format(1, 1, 1, headers.length, {
      fontWeight: 'bold',
      background: '#7C3AED',
      fontColor: '#FFFFFF',
      border: true
    });
    
    // Sample campaigns
    const sampleCampaigns = [
      ['CAM001', 'Spring Product Launch', 'Product', 'Awareness', 'Social', '=TODAY()-30', '=TODAY()+30', 'Active', 50000, 25000, '=I2-J2', 'Young Adults 18-34', 'Video + Images', 'spring-launch.com', 'Sarah Marketing', 'High priority campaign'],
      ['CAM002', 'Email Newsletter Q2', 'Newsletter', 'Engagement', 'Email', '=TODAY()-15', '=TODAY()+45', 'Active', 10000, 3500, '=I3-J3', 'Existing Customers', 'HTML Template', 'newsletter.com', 'Mike Email', 'Weekly sends'],
      ['CAM003', 'Summer Sale PPC', 'Promotional', 'Conversion', 'Search', '=TODAY()', '=TODAY()+60', 'Active', 75000, 0, '=I4-J4', 'Intent Buyers', 'Text Ads', 'summer-sale.com', 'John PPC', 'Google Ads + Bing'],
      ['CAM004', 'Brand Awareness Display', 'Brand', 'Awareness', 'Display', '=TODAY()-45', '=TODAY()-5', 'Completed', 30000, 30000, '=I5-J5', 'Broad Market', 'Banner Ads', 'brand.com', 'Lisa Brand', 'Q1 campaign completed'],
      ['CAM005', 'Content Marketing Blog', 'Content', 'Engagement', 'Content', '=TODAY()-60', '=TODAY()+90', 'Active', 20000, 8000, '=I6-J6', 'B2B Decision Makers', 'Blog Posts', 'blog.com', 'Tom Content', 'SEO focused']
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
    campaigns.addValidation('C2:C100', ['Product', 'Brand', 'Promotional', 'Content', 'Event', 'Newsletter']);
    campaigns.addValidation('D2:D100', ['Awareness', 'Consideration', 'Conversion', 'Retention', 'Engagement']);
    campaigns.addValidation('E2:E100', ['Social', 'Email', 'Search', 'Display', 'Content', 'Video', 'Affiliate']);
    campaigns.addValidation('H2:H100', ['Planning', 'Scheduled', 'Active', 'Paused', 'Completed', 'Cancelled']);
    
    // Formatting
    campaigns.format(2, 6, 100, 7, { numberFormat: 'dd/mm/yyyy' });
    campaigns.format(2, 9, 100, 11, { numberFormat: '$#,##0' });
    
    // Conditional formatting for status
    campaigns.addConditionalFormat(2, 8, 100, 8, {
      type: 'cell',
      condition: 'TEXT_EQUALS',
      value: 'Active',
      format: { background: '#10B981', fontColor: '#FFFFFF' }
    });
    
    campaigns.addConditionalFormat(2, 8, 100, 8, {
      type: 'cell',
      condition: 'TEXT_EQUALS',
      value: 'Completed',
      format: { background: '#6B7280', fontColor: '#FFFFFF' }
    });
    
    // Column widths
    campaigns.setColumnWidth(1, 100);
    campaigns.setColumnWidth(2, 180);
    for (let col = 3; col <= headers.length; col++) {
      campaigns.setColumnWidth(col, 100);
    }
  },
  
  buildPerformanceSheet: function(performance) {
    // Headers
    const headers = [
      'Campaign ID', 'Date', 'Impressions', 'Reach', 'Clicks', 'CTR',
      'Leads', 'Conversions', 'Conversion Rate', 'CPC', 'CPL', 'CPA',
      'Revenue', 'ROI', 'Quality Score'
    ];
    
    headers.forEach((header, idx) => {
      performance.setValue(1, idx + 1, header);
    });
    
    performance.format(1, 1, 1, headers.length, {
      fontWeight: 'bold',
      background: '#7C3AED',
      fontColor: '#FFFFFF',
      border: true
    });
    
    // Sample performance data
    const campaigns = ['CAM001', 'CAM002', 'CAM003', 'CAM004', 'CAM005'];
    let row = 2;
    
    campaigns.forEach(campaign => {
      for (let day = 0; day < 7; day++) {
        performance.setValue(row, 1, campaign);
        performance.setFormula(row, 2, `=TODAY()-${6-day}`);
        
        const impressions = Math.floor(10000 + Math.random() * 50000);
        const reach = Math.floor(impressions * 0.8);
        const clicks = Math.floor(impressions * (0.01 + Math.random() * 0.04));
        const leads = Math.floor(clicks * (0.05 + Math.random() * 0.15));
        const conversions = Math.floor(leads * (0.1 + Math.random() * 0.3));
        
        performance.setValue(row, 3, impressions);
        performance.setValue(row, 4, reach);
        performance.setValue(row, 5, clicks);
        performance.setFormula(row, 6, `=E${row}/C${row}`);
        performance.setValue(row, 7, leads);
        performance.setValue(row, 8, conversions);
        performance.setFormula(row, 9, `=H${row}/E${row}`);
        performance.setFormula(row, 10, `=VLOOKUP(A${row},{{Campaigns}}!A:J,10,FALSE)/E${row}`);
        performance.setFormula(row, 11, `=VLOOKUP(A${row},{{Campaigns}}!A:J,10,FALSE)/G${row}`);
        performance.setFormula(row, 12, `=VLOOKUP(A${row},{{Campaigns}}!A:J,10,FALSE)/H${row}`);
        performance.setFormula(row, 13, `=H${row}*RANDBETWEEN(50,200)`);
        performance.setFormula(row, 14, `=(M${row}-VLOOKUP(A${row},{{Campaigns}}!A:J,10,FALSE))/VLOOKUP(A${row},{{Campaigns}}!A:J,10,FALSE)`);
        performance.setValue(row, 15, 7 + Math.random() * 3);
        
        row++;
      }
    });
    
    // Formatting
    performance.format(2, 2, row, 2, { numberFormat: 'dd/mm/yyyy' });
    performance.format(2, 3, row, 5, { numberFormat: '#,##0' });
    performance.format(2, 6, row, 6, { numberFormat: '0.00%' });
    performance.format(2, 7, row, 8, { numberFormat: '#,##0' });
    performance.format(2, 9, row, 9, { numberFormat: '0.00%' });
    performance.format(2, 10, row, 12, { numberFormat: '$#,##0.00' });
    performance.format(2, 13, row, 13, { numberFormat: '$#,##0' });
    performance.format(2, 14, row, 14, { numberFormat: '0.0x' });
    performance.format(2, 15, row, 15, { numberFormat: '0.0' });
    
    // Column widths
    for (let col = 1; col <= headers.length; col++) {
      performance.setColumnWidth(col, 90);
    }
  },
  
  buildChannelsSheet: function(channels) {
    // Channel performance tracking
    const headers = [
      'Channel', 'Active Campaigns', 'Total Budget', 'Total Spend', 'Impressions',
      'Clicks', 'CTR', 'Conversions', 'CPA', 'Revenue', 'ROI', 'Efficiency Score'
    ];
    
    headers.forEach((header, idx) => {
      channels.setValue(1, idx + 1, header);
    });
    
    channels.format(1, 1, 1, headers.length, {
      fontWeight: 'bold',
      background: '#7C3AED',
      fontColor: '#FFFFFF',
      border: true
    });
    
    const channelList = ['Social', 'Email', 'Search', 'Display', 'Content', 'Video', 'Affiliate'];
    
    channelList.forEach((channel, idx) => {
      channels.setValue(idx + 2, 1, channel);
      channels.setFormula(idx + 2, 2, `=COUNTIF({{Campaigns}}!E:E,A${idx + 2})`);
      channels.setFormula(idx + 2, 3, `=SUMIF({{Campaigns}}!E:E,A${idx + 2},{{Campaigns}}!I:I)`);
      channels.setFormula(idx + 2, 4, `=SUMIF({{Campaigns}}!E:E,A${idx + 2},{{Campaigns}}!J:J)`);
      channels.setFormula(idx + 2, 5, `=SUMIFS({{Performance}}!C:C,{{Performance}}!A:A,{{Campaigns}}!A:A,{{Campaigns}}!E:E,A${idx + 2})`);
      channels.setFormula(idx + 2, 6, `=SUMIFS({{Performance}}!E:E,{{Performance}}!A:A,{{Campaigns}}!A:A,{{Campaigns}}!E:E,A${idx + 2})`);
      channels.setFormula(idx + 2, 7, `=F${idx + 2}/E${idx + 2}`);
      channels.setFormula(idx + 2, 8, `=SUMIFS({{Performance}}!H:H,{{Performance}}!A:A,{{Campaigns}}!A:A,{{Campaigns}}!E:E,A${idx + 2})`);
      channels.setFormula(idx + 2, 9, `=D${idx + 2}/H${idx + 2}`);
      channels.setFormula(idx + 2, 10, `=H${idx + 2}*RANDBETWEEN(50,150)`);
      channels.setFormula(idx + 2, 11, `=(J${idx + 2}-D${idx + 2})/D${idx + 2}`);
      channels.setFormula(idx + 2, 12, `=IF(K${idx + 2}>2,"Excellent",IF(K${idx + 2}>1,"Good",IF(K${idx + 2}>0.5,"Fair","Poor")))`);
    });
    
    // Formatting
    channels.format(2, 3, 8, 4, { numberFormat: '$#,##0' });
    channels.format(2, 5, 8, 6, { numberFormat: '#,##0' });
    channels.format(2, 7, 8, 7, { numberFormat: '0.00%' });
    channels.format(2, 8, 8, 8, { numberFormat: '#,##0' });
    channels.format(2, 9, 8, 9, { numberFormat: '$#,##0.00' });
    channels.format(2, 10, 8, 10, { numberFormat: '$#,##0' });
    channels.format(2, 11, 8, 11, { numberFormat: '0.0x' });
    
    // Column widths
    channels.setColumnWidth(1, 100);
    for (let col = 2; col <= headers.length; col++) {
      channels.setColumnWidth(col, 100);
    }
  },
  
  buildBudgetSheet: function(budget) {
    // Budget allocation and tracking
    const headers = [
      'Category', 'Subcategory', 'Q1 Budget', 'Q2 Budget', 'Q3 Budget', 'Q4 Budget',
      'Total Budget', 'Allocated', 'Spent', 'Remaining', '% Used', 'Status'
    ];
    
    headers.forEach((header, idx) => {
      budget.setValue(1, idx + 1, header);
    });
    
    budget.format(1, 1, 1, headers.length, {
      fontWeight: 'bold',
      background: '#7C3AED',
      fontColor: '#FFFFFF',
      border: true
    });
    
    const budgetItems = [
      ['Digital', 'Social Media', 25000, 30000, 35000, 30000],
      ['Digital', 'Search Marketing', 40000, 45000, 45000, 40000],
      ['Digital', 'Display Ads', 15000, 15000, 20000, 20000],
      ['Digital', 'Email Marketing', 5000, 5000, 5000, 5000],
      ['Content', 'Blog & Articles', 10000, 10000, 10000, 10000],
      ['Content', 'Video Production', 20000, 25000, 30000, 25000],
      ['Traditional', 'Print Ads', 10000, 8000, 8000, 10000],
      ['Traditional', 'Events', 15000, 20000, 25000, 20000],
      ['Tools', 'Software', 5000, 5000, 5000, 5000],
      ['Tools', 'Analytics', 3000, 3000, 3000, 3000]
    ];
    
    budgetItems.forEach((item, idx) => {
      const row = idx + 2;
      budget.setValue(row, 1, item[0]);
      budget.setValue(row, 2, item[1]);
      budget.setValue(row, 3, item[2]);
      budget.setValue(row, 4, item[3]);
      budget.setValue(row, 5, item[4]);
      budget.setValue(row, 6, item[5]);
      budget.setFormula(row, 7, `=SUM(C${row}:F${row})`);
      budget.setFormula(row, 8, `=G${row}*RANDBETWEEN(60,90)/100`);
      budget.setFormula(row, 9, `=H${row}*RANDBETWEEN(30,80)/100`);
      budget.setFormula(row, 10, `=H${row}-I${row}`);
      budget.setFormula(row, 11, `=I${row}/H${row}`);
      budget.setFormula(row, 12, `=IF(K${row}>0.9,"Over Budget",IF(K${row}>0.7,"On Track",IF(K${row}>0.5,"Good","Under Utilized")))`);
    });
    
    // Totals
    budget.setValue(12, 1, 'TOTAL');
    budget.merge(12, 1, 12, 2);
    for (let col = 3; col <= 10; col++) {
      budget.setFormula(12, col, `=SUM(${String.fromCharCode(64 + col)}2:${String.fromCharCode(64 + col)}11)`);
    }
    budget.setFormula(12, 11, '=I12/H12');
    budget.setValue(12, 12, '');
    
    budget.format(12, 1, 12, headers.length, {
      fontWeight: 'bold',
      background: '#E9D5FF'
    });
    
    // Formatting
    budget.format(2, 3, 12, 10, { numberFormat: '$#,##0' });
    budget.format(2, 11, 12, 11, { numberFormat: '0%' });
    
    // Conditional formatting for status
    budget.addConditionalFormat(2, 12, 11, 12, {
      type: 'cell',
      condition: 'TEXT_EQUALS',
      value: 'Over Budget',
      format: { background: '#FEE2E2', fontColor: '#DC2626' }
    });
    
    // Column widths
    budget.setColumnWidth(1, 100);
    budget.setColumnWidth(2, 120);
    for (let col = 3; col <= headers.length; col++) {
      budget.setColumnWidth(col, 90);
    }
  },
  
  buildCreativeSheet: function(creative) {
    // Creative assets tracking
    const headers = [
      'Asset ID', 'Campaign', 'Type', 'Name', 'Format', 'Size',
      'Created Date', 'Status', 'Performance', 'A/B Test', 'Notes'
    ];
    
    headers.forEach((header, idx) => {
      creative.setValue(1, idx + 1, header);
    });
    
    creative.format(1, 1, 1, headers.length, {
      fontWeight: 'bold',
      background: '#7C3AED',
      fontColor: '#FFFFFF',
      border: true
    });
    
    // Sample creative assets
    const assets = [
      ['CR001', 'CAM001', 'Video', 'Spring Launch Video', 'MP4', '1920x1080', '=TODAY()-20', 'Active', 'High', 'Version A', 'Main hero video'],
      ['CR002', 'CAM001', 'Image', 'Spring Banner Set', 'JPG', '728x90', '=TODAY()-20', 'Active', 'Medium', 'Version B', 'Display banner ads'],
      ['CR003', 'CAM002', 'Email', 'Newsletter Template', 'HTML', '600px', '=TODAY()-15', 'Active', 'High', 'Version A', 'Responsive template'],
      ['CR004', 'CAM003', 'Text', 'PPC Ad Copy Set', 'Text', 'N/A', '=TODAY()-10', 'Testing', 'Testing', 'Version A/B', '5 variations'],
      ['CR005', 'CAM005', 'Article', 'SEO Blog Post', 'HTML', '2000 words', '=TODAY()-5', 'Review', 'N/A', 'N/A', 'Pending approval']
    ];
    
    assets.forEach((asset, idx) => {
      asset.forEach((value, col) => {
        if (typeof value === 'string' && value.startsWith('=')) {
          creative.setFormula(idx + 2, col + 1, value);
        } else {
          creative.setValue(idx + 2, col + 1, value);
        }
      });
    });
    
    // Data validation
    creative.addValidation('C2:C100', ['Video', 'Image', 'Email', 'Text', 'Article', 'Infographic', 'Audio']);
    creative.addValidation('H2:H100', ['Draft', 'Review', 'Testing', 'Active', 'Paused', 'Archived']);
    creative.addValidation('I2:I100', ['High', 'Medium', 'Low', 'Testing', 'N/A']);
    
    // Formatting
    creative.format(2, 7, 100, 7, { numberFormat: 'dd/mm/yyyy' });
    
    // Column widths
    creative.setColumnWidth(1, 80);
    creative.setColumnWidth(2, 100);
    creative.setColumnWidth(4, 150);
    for (let col = 3; col <= headers.length; col++) {
      if (col !== 4) creative.setColumnWidth(col, 90);
    }
  },
  
  buildResultsSheet: function(results) {
    // Campaign results and analysis
    const headers = [
      'Campaign', 'Start Date', 'End Date', 'Duration', 'Total Spend',
      'Total Revenue', 'Profit', 'ROI', 'Conversions', 'CPA',
      'Lead Quality', 'Customer LTV', 'Success Rating'
    ];
    
    headers.forEach((header, idx) => {
      results.setValue(1, idx + 1, header);
    });
    
    results.format(1, 1, 1, headers.length, {
      fontWeight: 'bold',
      background: '#7C3AED',
      fontColor: '#FFFFFF',
      border: true
    });
    
    // Link to campaigns data with results
    for (let i = 0; i < 5; i++) {
      const row = i + 2;
      results.setFormula(row, 1, `=INDEX({{Campaigns}}!B:B,${row})`);
      results.setFormula(row, 2, `=INDEX({{Campaigns}}!F:F,${row})`);
      results.setFormula(row, 3, `=INDEX({{Campaigns}}!G:G,${row})`);
      results.setFormula(row, 4, `=C${row}-B${row}`);
      results.setFormula(row, 5, `=INDEX({{Campaigns}}!J:J,${row})`);
      results.setFormula(row, 6, `=SUMIF({{Performance}}!A:A,INDEX({{Campaigns}}!A:A,${row}),{{Performance}}!M:M)`);
      results.setFormula(row, 7, `=F${row}-E${row}`);
      results.setFormula(row, 8, `=G${row}/E${row}`);
      results.setFormula(row, 9, `=SUMIF({{Performance}}!A:A,INDEX({{Campaigns}}!A:A,${row}),{{Performance}}!H:H)`);
      results.setFormula(row, 10, `=E${row}/I${row}`);
      results.setFormula(row, 11, `=IF(H${row}>2,"High",IF(H${row}>1,"Medium","Low"))`);
      results.setValue(row, 12, 500 + Math.random() * 2000);
      results.setFormula(row, 13, `=IF(H${row}>3,"⭐⭐⭐⭐⭐",IF(H${row}>2,"⭐⭐⭐⭐",IF(H${row}>1,"⭐⭐⭐",IF(H${row}>0,"⭐⭐","⭐"))))`);
    }
    
    // Formatting
    results.format(2, 2, 10, 3, { numberFormat: 'dd/mm/yyyy' });
    results.format(2, 5, 10, 7, { numberFormat: '$#,##0' });
    results.format(2, 8, 10, 8, { numberFormat: '0.0x' });
    results.format(2, 9, 10, 9, { numberFormat: '#,##0' });
    results.format(2, 10, 10, 10, { numberFormat: '$#,##0.00' });
    results.format(2, 12, 10, 12, { numberFormat: '$#,##0' });
    
    // Column widths
    results.setColumnWidth(1, 150);
    for (let col = 2; col <= headers.length; col++) {
      results.setColumnWidth(col, 90);
    }
  },
  
  buildReportsSheet: function(reports) {
    // Executive reporting dashboard
    reports.setValue(1, 1, 'Marketing Executive Report');
    reports.merge(1, 1, 1, 10);
    reports.format(1, 1, 1, 10, {
      fontSize: 16,
      fontWeight: 'bold',
      horizontalAlignment: 'center',
      background: '#7C3AED',
      fontColor: '#FFFFFF'
    });
    
    // Report Period
    reports.setValue(3, 1, 'Report Period:');
    reports.setFormula(3, 2, '=TEXT(TODAY()-30,"MMM DD")&" - "&TEXT(TODAY(),"MMM DD, YYYY")');
    reports.merge(3, 2, 3, 4);
    
    // Executive Summary
    reports.setValue(5, 1, 'EXECUTIVE SUMMARY');
    reports.merge(5, 1, 5, 10);
    reports.format(5, 1, 5, 10, {
      fontWeight: 'bold',
      background: '#F3E8FF'
    });
    
    const summary = [
      ['Total Campaigns Run:', '=COUNTIF({{Campaigns}}!H:H,"<>")'],
      ['Total Marketing Spend:', '=SUM({{Campaigns}}!J:J)'],
      ['Total Revenue Generated:', '=SUM({{Results}}!F:F)'],
      ['Overall ROI:', '=AVERAGE({{Results}}!H:H)'],
      ['Total Conversions:', '=SUM({{Results}}!I:I)'],
      ['Average CPA:', '=AVERAGE({{Results}}!J:J)']
    ];
    
    summary.forEach((item, idx) => {
      reports.setValue(6 + idx, 1, item[0]);
      reports.setFormula(6 + idx, 2, item[1]);
      reports.merge(6 + idx, 2, 6 + idx, 3);
    });
    
    reports.format(7, 2, 7, 3, { numberFormat: '$#,##0' });
    reports.format(8, 2, 8, 3, { numberFormat: '$#,##0' });
    reports.format(9, 2, 9, 3, { numberFormat: '0.0x' });
    reports.format(11, 2, 11, 3, { numberFormat: '$#,##0.00' });
    
    // Key Achievements
    reports.setValue(5, 5, 'KEY ACHIEVEMENTS');
    reports.merge(5, 5, 5, 10);
    reports.format(5, 5, 5, 10, {
      fontWeight: 'bold',
      background: '#F3E8FF'
    });
    
    const achievements = [
      '✓ Exceeded quarterly revenue target by 15%',
      '✓ Reduced CPA by 22% through optimization',
      '✓ Increased email open rates to 35%',
      '✓ Launched 3 successful viral campaigns',
      '✓ Improved conversion rate by 2.5%',
      '✓ Achieved 4.2x average ROI'
    ];
    
    achievements.forEach((achievement, idx) => {
      reports.setValue(6 + idx, 5, achievement);
      reports.merge(6 + idx, 5, 6 + idx, 10);
    });
    
    // Channel Performance Table
    reports.setValue(13, 1, 'CHANNEL PERFORMANCE');
    reports.merge(13, 1, 13, 10);
    reports.format(13, 1, 13, 10, {
      fontWeight: 'bold',
      background: '#F3E8FF'
    });
    
    const channelHeaders = ['Channel', 'Campaigns', 'Spend', 'Revenue', 'ROI', 'Conversions', 'Trend'];
    channelHeaders.forEach((header, idx) => {
      reports.setValue(14, idx + 1, header);
    });
    reports.format(14, 1, 14, 7, {
      fontWeight: 'bold',
      background: '#E5E7EB'
    });
    
    // Pull channel data
    for (let i = 0; i < 5; i++) {
      reports.setFormula(15 + i, 1, `=INDEX({{Channels}}!A:A,${i + 2})`);
      reports.setFormula(15 + i, 2, `=INDEX({{Channels}}!B:B,${i + 2})`);
      reports.setFormula(15 + i, 3, `=INDEX({{Channels}}!D:D,${i + 2})`);
      reports.setFormula(15 + i, 4, `=INDEX({{Channels}}!J:J,${i + 2})`);
      reports.setFormula(15 + i, 5, `=INDEX({{Channels}}!K:K,${i + 2})`);
      reports.setFormula(15 + i, 6, `=INDEX({{Channels}}!H:H,${i + 2})`);
      reports.setFormula(15 + i, 7, `=IF(RANDBETWEEN(0,1)=1,"↑","↓")`);
    }
    
    reports.format(15, 3, 19, 4, { numberFormat: '$#,##0' });
    reports.format(15, 5, 19, 5, { numberFormat: '0.0x' });
    
    // Recommendations
    reports.setValue(21, 1, 'RECOMMENDATIONS');
    reports.merge(21, 1, 21, 10);
    reports.format(21, 1, 21, 10, {
      fontWeight: 'bold',
      background: '#F3E8FF'
    });
    
    const recommendations = [
      '1. Increase budget allocation to Social Media channel (highest ROI)',
      '2. Pause underperforming Display campaigns and reallocate budget',
      '3. Expand successful email campaigns to new segments',
      '4. Test new creative formats for video content',
      '5. Implement marketing automation for lead nurturing'
    ];
    
    recommendations.forEach((rec, idx) => {
      reports.setValue(22 + idx, 1, rec);
      reports.merge(22 + idx, 1, 22 + idx, 10);
    });
    
    // Column widths
    reports.setColumnWidth(1, 150);
    for (let col = 2; col <= 10; col++) {
      reports.setColumnWidth(col, 90);
    }
  },
  
  /**
   * Content Calendar - Editorial planning and scheduling
   */
  buildContentCalendar: function(builder) {
    const sheets = builder.createSheets([
      'Calendar View',
      'Content Pipeline',
      'Topics',
      'Authors',
      'Distribution',
      'Performance'
    ]);
    
    // Build calendar view and content management sheets
    // Implementation would follow similar pattern to above
    
    return builder.complete();
  },
  
  /**
   * Additional template stubs for other marketing templates
   * These would be implemented following the same comprehensive pattern
   */
  buildSocialMediaManager: function(builder) {
    // Implementation for social media management
    return builder.complete();
  },
  
  buildEmailMarketing: function(builder) {
    // Implementation for email marketing campaigns
    return builder.complete();
  },
  
  buildSEOTracker: function(builder) {
    // Implementation for SEO tracking and optimization
    return builder.complete();
  },
  
  buildAnalyticsDashboard: function(builder) {
    // Implementation for analytics dashboard
    return builder.complete();
  },
  
  buildLeadScoring: function(builder) {
    // Implementation for lead scoring system
    return builder.complete();
  },
  
  buildBudgetPlanner: function(builder) {
    // Implementation for marketing budget planning
    return builder.complete();
  },
  
  buildCompetitorAnalysis: function(builder) {
    // Implementation for competitor analysis
    return builder.complete();
  },
  
  buildROICalculator: function(builder) {
    // Implementation for ROI calculation
    return builder.complete();
  }
};