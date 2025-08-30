/**
 * Professional Healthcare Templates Suite
 * Complete templates for healthcare industry professionals
 */

var HealthcareTemplate = {
  /**
   * Build comprehensive Healthcare template based on type
   */
  build: function(spreadsheet, templateType = 'insurance-verifier', isPreview = false) {
    const builder = new TemplateBuilderPro(spreadsheet, isPreview);
    
    switch(templateType) {
      case 'insurance-verifier':
        return this.buildInsuranceVerifierTemplate(builder);
      case 'prior-auth-tracker':
        return this.buildPriorAuthTemplate(builder);
      case 'revenue-cycle':
        return this.buildRevenueCycleTemplate(builder);
      case 'denial-analytics':
        return this.buildDenialAnalyticsTemplate(builder);
      default:
        return this.buildInsuranceVerifierTemplate(builder);
    }
  },
  
  /**
   * Insurance Verifier Template
   */
  buildInsuranceVerifierTemplate: function(builder) {
    const sheets = builder.createSheets([
      'Dashboard',
      'Verifications',
      'Eligibility',
      'Benefits',
      'Authorizations',
      'Denials',
      'Appeals',
      'Reports'
    ]);
    
    this.buildInsuranceDashboard(builder);
    this.buildVerifications(builder);
    this.buildEligibility(builder);
    this.buildBenefits(builder);
    this.buildAuthorizations(builder);
    this.buildDenials(builder);
    this.buildAppeals(builder);
    this.buildInsuranceReports(builder);
    
    return builder.complete();
  },
  
  /**
   * Prior Authorization Tracker Template
   */
  buildPriorAuthTemplate: function(builder) {
    const sheets = builder.createSheets([
      'Dashboard',
      'Auth Requests',
      'Status Tracking',
      'Provider Notes',
      'Payer Responses',
      'Follow-ups',
      'Analytics',
      'Reports'
    ]);
    
    this.buildPriorAuthDashboard(builder);
    this.buildAuthRequests(builder);
    this.buildStatusTracking(builder);
    this.buildProviderNotes(builder);
    this.buildPayerResponses(builder);
    this.buildFollowUps(builder);
    this.buildAuthAnalytics(builder);
    this.buildPriorAuthReports(builder);
    
    return builder.complete();
  },
  
  /**
   * Revenue Cycle Template
   */
  buildRevenueCycleTemplate: function(builder) {
    const sheets = builder.createSheets([
      'Dashboard',
      'Claims',
      'Payments',
      'Aging',
      'Denials',
      'Write-offs',
      'Collections',
      'Reports'
    ]);
    
    this.buildRevenueDashboard(builder);
    this.buildClaims(builder);
    this.buildPayments(builder);
    this.buildAging(builder);
    this.buildRevenueDenials(builder);
    this.buildWriteOffs(builder);
    this.buildCollections(builder);
    this.buildRevenueReports(builder);
    
    return builder.complete();
  },
  
  /**
   * Denial Analytics Template
   */
  buildDenialAnalyticsTemplate: function(builder) {
    const sheets = builder.createSheets([
      'Dashboard',
      'Denial Trends',
      'Root Causes',
      'Appeal Success',
      'Payer Analysis',
      'Prevention Strategies',
      'Recovery',
      'Reports'
    ]);
    
    this.buildDenialDashboard(builder);
    this.buildDenialTrends(builder);
    this.buildRootCauses(builder);
    this.buildAppealSuccess(builder);
    this.buildPayerAnalysis(builder);
    this.buildPreventionStrategies(builder);
    this.buildRecovery(builder);
    this.buildDenialReports(builder);
    
    return builder.complete();
  },
  
  // ========================================
  // INSURANCE VERIFIER TEMPLATE BUILDERS
  // ========================================
  
  buildInsuranceDashboard: function(builder) {
    const dash = builder.sheet('Dashboard');
    
    // Header with healthcare cyan theme
    dash.merge(1, 1, 2, 12);
    dash.setValue(1, 1, 'Insurance Verification Command Center');
    dash.format(1, 1, 2, 12, {
      background: '#0891B2',
      fontColor: '#FFFFFF',
      fontSize: 18,
      fontWeight: 'bold',
      horizontalAlignment: 'center'
    });
    
    // Date and stats
    dash.setValue(3, 1, 'Last Updated:');
    dash.setFormula(3, 2, '=NOW()');
    dash.format(3, 2, 3, 2, { numberFormat: 'dd/mm/yyyy hh:mm' });
    dash.setValue(3, 4, 'Today\'s Date:');
    dash.setFormula(3, 5, '=TODAY()');
    dash.format(3, 5, 3, 5, { numberFormat: 'dd/mm/yyyy' });
    
    // KPI Cards
    const kpis = [
      { row: 5, col: 1, title: 'Pending Verifications', formula: '=COUNTIF({{Verifications}}!G:G,"Pending")', format: '#,##0', color: '#F59E0B' },
      { row: 5, col: 4, title: 'Completed Today', formula: '=COUNTIFS({{Verifications}}!G:G,"Completed",{{Verifications}}!B:B,TODAY())', format: '#,##0', color: '#10B981' },
      { row: 5, col: 7, title: 'Active Patients', formula: '=COUNTIF({{Eligibility}}!F:F,"Active")', format: '#,##0', color: '#3B82F6' },
      { row: 5, col: 10, title: 'Success Rate', formula: '=COUNTIF({{Verifications}}!G:G,"Approved")/COUNTA({{Verifications}}!A:A)-1', format: '0.0%', color: '#10B981' }
    ];
    
    kpis.forEach(kpi => {
      dash.merge(kpi.row, kpi.col, kpi.row + 2, kpi.col + 2);
      dash.setValue(kpi.row, kpi.col, kpi.title);
      dash.setFormula(kpi.row + 1, kpi.col, kpi.formula);
      dash.format(kpi.row, kpi.col, kpi.row, kpi.col + 2, {
        background: '#E0F2FE',
        fontWeight: 'bold'
      });
      dash.format(kpi.row + 1, kpi.col, kpi.row + 2, kpi.col + 2, {
        fontSize: 20,
        fontWeight: 'bold',
        horizontalAlignment: 'center',
        numberFormat: kpi.format,
        fontColor: kpi.color
      });
    });
    
    // Summary Table
    dash.setValue(9, 1, 'Insurance Verification Summary');
    dash.format(9, 1, 9, 8, {
      fontSize: 14,
      fontWeight: 'bold',
      background: '#06B6D4',
      fontColor: '#FFFFFF'
    });
    
    const summaryHeaders = [
      'Payer', 'Total Verifications', 'Approved', 'Denied', 'Pending', 'Avg Response Time', 'Success Rate', 'Notes'
    ];
    
    dash.setRangeValues(10, 1, [summaryHeaders]);
    dash.format(10, 1, 10, 8, {
      background: '#67E8F9',
      fontWeight: 'bold'
    });
    
    // Sample data
    const sampleData = [
      ['Blue Cross', 45, 38, 5, 2, '2.3 days', '84%', 'Fast response'],
      ['Aetna', 32, 28, 3, 1, '1.8 days', '88%', 'Good coverage'],
      ['United Health', 28, 22, 4, 2, '3.1 days', '79%', 'Slow approvals'],
      ['Medicare', 56, 51, 3, 2, '1.2 days', '91%', 'Excellent']
    ];
    
    dash.setRangeValues(11, 1, sampleData);
    dash.format(11, 7, 14, 7, { numberFormat: '0%' });
  },
  
  buildVerifications: function(builder) {
    const verif = builder.sheet('Verifications');
    
    // Headers
    const headers = [
      'Verification ID', 'Date', 'Patient Name', 'DOB', 'Insurance', 'Policy Number',
      'Status', 'Verified By', 'Effective Date', 'Term Date', 'Copay', 'Deductible', 'Notes'
    ];
    
    verif.setRangeValues(1, 1, [headers]);
    verif.format(1, 1, 1, 13, {
      background: '#0891B2',
      fontColor: '#FFFFFF',
      fontWeight: 'bold'
    });
    
    // Add validation for status
    verif.addValidation('G2:G1000', ['Pending', 'In Progress', 'Approved', 'Denied', 'Completed']);
  },
  
  buildEligibility: function(builder) {
    const elig = builder.sheet('Eligibility');
    
    // Headers
    const headers = [
      'Patient ID', 'Name', 'Insurance', 'Member ID', 'Group Number',
      'Status', 'Coverage Start', 'Coverage End', 'Plan Type', 'Network', 'Notes'
    ];
    
    elig.setRangeValues(1, 1, [headers]);
    elig.format(1, 1, 1, 11, {
      background: '#0891B2',
      fontColor: '#FFFFFF',
      fontWeight: 'bold'
    });
  },
  
  buildBenefits: function(builder) {
    const benefits = builder.sheet('Benefits');
    
    // Headers
    const headers = [
      'Patient', 'Insurance', 'Service Type', 'Coverage %', 'Copay',
      'Deductible', 'Out of Pocket Max', 'Used YTD', 'Remaining', 'Prior Auth Required', 'Notes'
    ];
    
    benefits.setRangeValues(1, 1, [headers]);
    benefits.format(1, 1, 1, 11, {
      background: '#0891B2',
      fontColor: '#FFFFFF',
      fontWeight: 'bold'
    });
  },
  
  buildAuthorizations: function(builder) {
    const auth = builder.sheet('Authorizations');
    
    // Headers
    const headers = [
      'Auth Number', 'Date', 'Patient', 'Service', 'Provider',
      'Units Approved', 'Start Date', 'End Date', 'Status', 'Notes'
    ];
    
    auth.setRangeValues(1, 1, [headers]);
    auth.format(1, 1, 1, 10, {
      background: '#0891B2',
      fontColor: '#FFFFFF',
      fontWeight: 'bold'
    });
  },
  
  buildDenials: function(builder) {
    const denials = builder.sheet('Denials');
    
    // Headers
    const headers = [
      'Denial ID', 'Date', 'Patient', 'Service', 'Insurance',
      'Reason Code', 'Reason Description', 'Amount', 'Appeal Status', 'Resolution', 'Notes'
    ];
    
    denials.setRangeValues(1, 1, [headers]);
    denials.format(1, 1, 1, 11, {
      background: '#0891B2',
      fontColor: '#FFFFFF',
      fontWeight: 'bold'
    });
  },
  
  buildAppeals: function(builder) {
    const appeals = builder.sheet('Appeals');
    
    // Headers
    const headers = [
      'Appeal ID', 'Denial ID', 'Date Filed', 'Patient', 'Insurance',
      'Appeal Level', 'Status', 'Decision Date', 'Outcome', 'Recovery Amount', 'Notes'
    ];
    
    appeals.setRangeValues(1, 1, [headers]);
    appeals.format(1, 1, 1, 11, {
      background: '#0891B2',
      fontColor: '#FFFFFF',
      fontWeight: 'bold'
    });
  },
  
  buildInsuranceReports: function(builder) {
    const reports = builder.sheet('Reports');
    
    // Report header
    reports.merge(1, 1, 1, 10);
    reports.setValue(1, 1, 'Insurance Verification Reports');
    reports.format(1, 1, 1, 10, {
      background: '#0891B2',
      fontColor: '#FFFFFF',
      fontSize: 16,
      fontWeight: 'bold',
      horizontalAlignment: 'center'
    });
  },
  
  // ========================================
  // PRIOR AUTH TRACKER TEMPLATE BUILDERS
  // ========================================
  
  buildPriorAuthDashboard: function(builder) {
    const dash = builder.sheet('Dashboard');
    
    // Header
    dash.merge(1, 1, 2, 12);
    dash.setValue(1, 1, 'Prior Authorization Management System');
    dash.format(1, 1, 2, 12, {
      background: '#0E7490',
      fontColor: '#FFFFFF',
      fontSize: 18,
      fontWeight: 'bold',
      horizontalAlignment: 'center'
    });
    
    // KPIs
    const kpis = [
      { row: 4, col: 1, title: 'Pending Auths', formula: '=COUNTIF({{Auth Requests}}!H:H,"Pending")', format: '#,##0' },
      { row: 4, col: 4, title: 'Urgent Cases', formula: '=COUNTIF({{Auth Requests}}!I:I,"Urgent")', format: '#,##0' },
      { row: 4, col: 7, title: 'Approved Today', formula: '=COUNTIFS({{Auth Requests}}!H:H,"Approved",{{Auth Requests}}!B:B,TODAY())', format: '#,##0' },
      { row: 4, col: 10, title: 'Avg TAT (days)', formula: '=AVERAGE({{Status Tracking}}!F:F)', format: '0.0' }
    ];
    
    kpis.forEach(kpi => {
      dash.merge(kpi.row, kpi.col, kpi.row + 2, kpi.col + 2);
      dash.setValue(kpi.row, kpi.col, kpi.title);
      dash.setFormula(kpi.row + 1, kpi.col, kpi.formula);
      dash.format(kpi.row, kpi.col, kpi.row + 2, kpi.col + 2, {
        background: '#CFFAFE',
        fontWeight: 'bold',
        horizontalAlignment: 'center',
        numberFormat: kpi.format
      });
    });
  },
  
  buildAuthRequests: function(builder) {
    const requests = builder.sheet('Auth Requests');
    
    // Headers
    const headers = [
      'Request ID', 'Date', 'Patient', 'Provider', 'Service/Procedure',
      'CPT Code', 'ICD-10', 'Status', 'Priority', 'Payer', 'Auth Number', 'Notes'
    ];
    
    requests.setRangeValues(1, 1, [headers]);
    requests.format(1, 1, 1, 12, {
      background: '#0E7490',
      fontColor: '#FFFFFF',
      fontWeight: 'bold'
    });
  },
  
  buildStatusTracking: function(builder) {
    const status = builder.sheet('Status Tracking');
    
    // Headers
    const headers = [
      'Request ID', 'Submitted Date', 'Last Update', 'Current Status',
      'Days Pending', 'TAT', 'Reviewer', 'Next Action', 'Due Date', 'Notes'
    ];
    
    status.setRangeValues(1, 1, [headers]);
    status.format(1, 1, 1, 10, {
      background: '#0E7490',
      fontColor: '#FFFFFF',
      fontWeight: 'bold'
    });
  },
  
  buildProviderNotes: function(builder) {
    const notes = builder.sheet('Provider Notes');
    
    // Headers
    const headers = [
      'Date', 'Request ID', 'Provider', 'Note Type', 'Clinical Notes',
      'Medical Necessity', 'Supporting Docs', 'Entered By', 'Status'
    ];
    
    notes.setRangeValues(1, 1, [headers]);
    notes.format(1, 1, 1, 9, {
      background: '#0E7490',
      fontColor: '#FFFFFF',
      fontWeight: 'bold'
    });
  },
  
  buildPayerResponses: function(builder) {
    const responses = builder.sheet('Payer Responses');
    
    // Headers
    const headers = [
      'Response Date', 'Request ID', 'Payer', 'Decision', 'Auth Number',
      'Units Approved', 'Effective Date', 'Expiry Date', 'Denial Reason', 'Notes'
    ];
    
    responses.setRangeValues(1, 1, [headers]);
    responses.format(1, 1, 1, 10, {
      background: '#0E7490',
      fontColor: '#FFFFFF',
      fontWeight: 'bold'
    });
  },
  
  buildFollowUps: function(builder) {
    const followups = builder.sheet('Follow-ups');
    
    // Headers
    const headers = [
      'Follow-up Date', 'Request ID', 'Type', 'Contact Method', 'Contact Person',
      'Outcome', 'Next Step', 'Scheduled By', 'Completed', 'Notes'
    ];
    
    followups.setRangeValues(1, 1, [headers]);
    followups.format(1, 1, 1, 10, {
      background: '#0E7490',
      fontColor: '#FFFFFF',
      fontWeight: 'bold'
    });
  },
  
  buildAuthAnalytics: function(builder) {
    const analytics = builder.sheet('Analytics');
    
    // Title
    analytics.merge(1, 1, 1, 10);
    analytics.setValue(1, 1, 'Prior Authorization Analytics');
    analytics.format(1, 1, 1, 10, {
      background: '#0E7490',
      fontColor: '#FFFFFF',
      fontSize: 14,
      fontWeight: 'bold',
      horizontalAlignment: 'center'
    });
    
    // Metrics sections
    analytics.setValue(3, 1, 'Performance Metrics');
    analytics.format(3, 1, 3, 4, {
      fontSize: 12,
      fontWeight: 'bold',
      background: '#E0F2FE'
    });
    
    const metrics = [
      ['Metric', 'Current Month', 'Last Month', 'Change'],
      ['Total Requests', '=COUNTIF({{Auth Requests}}!B:B,">="&EOMONTH(TODAY(),-1)+1)', '=COUNTIFS({{Auth Requests}}!B:B,">="&EOMONTH(TODAY(),-2)+1,{{Auth Requests}}!B:B,"<="&EOMONTH(TODAY(),-1))', '=B5-C5'],
      ['Approval Rate', '=COUNTIF({{Auth Requests}}!H:H,"Approved")/COUNTA({{Auth Requests}}!A:A)', '85%', '=B6-C6'],
      ['Avg TAT (days)', '=AVERAGE({{Status Tracking}}!F:F)', '3.2', '=B7-C7'],
      ['Urgent Cases', '=COUNTIF({{Auth Requests}}!I:I,"Urgent")', '12', '=B8-C8']
    ];
    
    analytics.setRangeValues(4, 1, metrics);
    analytics.format(5, 2, 8, 3, { numberFormat: '#,##0' });
    analytics.format(6, 2, 6, 4, { numberFormat: '0%' });
  },
  
  buildPriorAuthReports: function(builder) {
    const reports = builder.sheet('Reports');
    
    // Report header
    reports.setValue(1, 1, 'Prior Authorization Reports');
    reports.format(1, 1, 1, 10, {
      background: '#0E7490',
      fontColor: '#FFFFFF',
      fontSize: 16,
      fontWeight: 'bold'
    });
  },
  
  // ========================================
  // REVENUE CYCLE TEMPLATE BUILDERS
  // ========================================
  
  buildRevenueDashboard: function(builder) {
    const dash = builder.sheet('Dashboard');
    
    // Header
    dash.merge(1, 1, 2, 12);
    dash.setValue(1, 1, 'Revenue Cycle Management Dashboard');
    dash.format(1, 1, 2, 12, {
      background: '#164E63',
      fontColor: '#FFFFFF',
      fontSize: 18,
      fontWeight: 'bold',
      horizontalAlignment: 'center'
    });
    
    // KPIs
    const kpis = [
      { row: 4, col: 1, title: 'Total AR', formula: '=SUM({{Claims}}!G:G)-SUM({{Payments}}!D:D)', format: '$#,##0' },
      { row: 4, col: 4, title: 'Days in AR', formula: '=AVERAGE({{Aging}}!E:E)', format: '0' },
      { row: 4, col: 7, title: 'Collection Rate', formula: '=SUM({{Payments}}!D:D)/SUM({{Claims}}!G:G)', format: '0.0%' },
      { row: 4, col: 10, title: 'Denial Rate', formula: '=COUNTIF({{Denials}}!F:F,"Denied")/COUNTA({{Claims}}!A:A)', format: '0.0%' }
    ];
    
    kpis.forEach(kpi => {
      dash.merge(kpi.row, kpi.col, kpi.row + 2, kpi.col + 2);
      dash.setValue(kpi.row, kpi.col, kpi.title);
      dash.setFormula(kpi.row + 1, kpi.col, kpi.formula);
      dash.format(kpi.row, kpi.col, kpi.row + 2, kpi.col + 2, {
        background: '#ECFEFF',
        fontWeight: 'bold',
        horizontalAlignment: 'center',
        numberFormat: kpi.format
      });
    });
  },
  
  buildClaims: function(builder) {
    const claims = builder.sheet('Claims');
    
    // Headers
    const headers = [
      'Claim ID', 'Date', 'Patient', 'Provider', 'Service Date',
      'CPT Code', 'Billed Amount', 'Allowed Amount', 'Status', 'Payer', 'Days Outstanding', 'Notes'
    ];
    
    claims.setRangeValues(1, 1, [headers]);
    claims.format(1, 1, 1, 12, {
      background: '#164E63',
      fontColor: '#FFFFFF',
      fontWeight: 'bold'
    });
  },
  
  buildPayments: function(builder) {
    const payments = builder.sheet('Payments');
    
    // Headers
    const headers = [
      'Payment Date', 'Claim ID', 'Payer', 'Payment Amount', 'Patient Responsibility',
      'Adjustment', 'Check Number', 'Posting Date', 'Posted By', 'Notes'
    ];
    
    payments.setRangeValues(1, 1, [headers]);
    payments.format(1, 1, 1, 10, {
      background: '#164E63',
      fontColor: '#FFFFFF',
      fontWeight: 'bold'
    });
  },
  
  buildAging: function(builder) {
    const aging = builder.sheet('Aging');
    
    // Headers
    const headers = [
      'Claim ID', 'Patient', 'Payer', 'Balance', 'Days Outstanding',
      '0-30 Days', '31-60 Days', '61-90 Days', '91-120 Days', '>120 Days', 'Status'
    ];
    
    aging.setRangeValues(1, 1, [headers]);
    aging.format(1, 1, 1, 11, {
      background: '#164E63',
      fontColor: '#FFFFFF',
      fontWeight: 'bold'
    });
  },
  
  buildRevenueDenials: function(builder) {
    const denials = builder.sheet('Denials');
    
    // Headers
    const headers = [
      'Denial Date', 'Claim ID', 'Patient', 'Payer', 'Amount',
      'Status', 'Reason Code', 'Reason', 'Preventable', 'Action Taken', 'Recovery Amount'
    ];
    
    denials.setRangeValues(1, 1, [headers]);
    denials.format(1, 1, 1, 11, {
      background: '#164E63',
      fontColor: '#FFFFFF',
      fontWeight: 'bold'
    });
  },
  
  buildWriteOffs: function(builder) {
    const writeoffs = builder.sheet('Write-offs');
    
    // Headers
    const headers = [
      'Date', 'Claim ID', 'Patient', 'Original Amount', 'Write-off Amount',
      'Reason', 'Category', 'Approved By', 'Documentation', 'Notes'
    ];
    
    writeoffs.setRangeValues(1, 1, [headers]);
    writeoffs.format(1, 1, 1, 10, {
      background: '#164E63',
      fontColor: '#FFFFFF',
      fontWeight: 'bold'
    });
  },
  
  buildCollections: function(builder) {
    const collections = builder.sheet('Collections');
    
    // Headers
    const headers = [
      'Account', 'Patient', 'Balance', 'Days Past Due', 'Last Contact',
      'Next Action', 'Collection Agency', 'Status', 'Payment Plan', 'Notes'
    ];
    
    collections.setRangeValues(1, 1, [headers]);
    collections.format(1, 1, 1, 10, {
      background: '#164E63',
      fontColor: '#FFFFFF',
      fontWeight: 'bold'
    });
  },
  
  buildRevenueReports: function(builder) {
    const reports = builder.sheet('Reports');
    
    // Report header
    reports.setValue(1, 1, 'Revenue Cycle Performance Reports');
    reports.format(1, 1, 1, 10, {
      background: '#164E63',
      fontColor: '#FFFFFF',
      fontSize: 16,
      fontWeight: 'bold'
    });
  },
  
  // ========================================
  // DENIAL ANALYTICS TEMPLATE BUILDERS
  // ========================================
  
  buildDenialDashboard: function(builder) {
    const dash = builder.sheet('Dashboard');
    
    // Header
    dash.merge(1, 1, 2, 12);
    dash.setValue(1, 1, 'Denial Analytics & Prevention Center');
    dash.format(1, 1, 2, 12, {
      background: '#155E75',
      fontColor: '#FFFFFF',
      fontSize: 18,
      fontWeight: 'bold',
      horizontalAlignment: 'center'
    });
  },
  
  buildDenialTrends: function(builder) {
    const trends = builder.sheet('Denial Trends');
    
    // Headers
    const headers = [
      'Month', 'Total Claims', 'Denials', 'Denial Rate', 'Amount Denied',
      'Recovered', 'Recovery Rate', 'Top Reason', 'Top Payer', 'Preventable %'
    ];
    
    trends.setRangeValues(1, 1, [headers]);
    trends.format(1, 1, 1, 10, {
      background: '#155E75',
      fontColor: '#FFFFFF',
      fontWeight: 'bold'
    });
  },
  
  buildRootCauses: function(builder) {
    const causes = builder.sheet('Root Causes');
    
    // Headers
    const headers = [
      'Reason Code', 'Description', 'Frequency', 'Total Amount', 'Avg Amount',
      'Department', 'Preventable', 'Action Plan', 'Owner', 'Status'
    ];
    
    causes.setRangeValues(1, 1, [headers]);
    causes.format(1, 1, 1, 10, {
      background: '#155E75',
      fontColor: '#FFFFFF',
      fontWeight: 'bold'
    });
  },
  
  buildAppealSuccess: function(builder) {
    const appeals = builder.sheet('Appeal Success');
    
    // Headers
    const headers = [
      'Appeal Level', 'Total Appeals', 'Successful', 'Success Rate', 'Avg Days',
      'Total Recovered', 'ROI', 'Top Success Factor', 'Notes'
    ];
    
    appeals.setRangeValues(1, 1, [headers]);
    appeals.format(1, 1, 1, 9, {
      background: '#155E75',
      fontColor: '#FFFFFF',
      fontWeight: 'bold'
    });
  },
  
  buildPayerAnalysis: function(builder) {
    const payer = builder.sheet('Payer Analysis');
    
    // Headers
    const headers = [
      'Payer', 'Total Claims', 'Denials', 'Denial Rate', 'Top Denial Reason',
      'Avg TAT', 'Appeal Success Rate', 'Contract Terms', 'Relationship Score', 'Notes'
    ];
    
    payer.setRangeValues(1, 1, [headers]);
    payer.format(1, 1, 1, 10, {
      background: '#155E75',
      fontColor: '#FFFFFF',
      fontWeight: 'bold'
    });
  },
  
  buildPreventionStrategies: function(builder) {
    const prevention = builder.sheet('Prevention Strategies');
    
    // Headers
    const headers = [
      'Strategy ID', 'Category', 'Description', 'Implementation Date', 'Owner',
      'Expected Impact', 'Actual Impact', 'Status', 'Cost', 'ROI'
    ];
    
    prevention.setRangeValues(1, 1, [headers]);
    prevention.format(1, 1, 1, 10, {
      background: '#155E75',
      fontColor: '#FFFFFF',
      fontWeight: 'bold'
    });
  },
  
  buildRecovery: function(builder) {
    const recovery = builder.sheet('Recovery');
    
    // Headers
    const headers = [
      'Claim ID', 'Denial Date', 'Amount', 'Recovery Action', 'Action Date',
      'Recovered Amount', 'Days to Recover', 'Success', 'Staff', 'Notes'
    ];
    
    recovery.setRangeValues(1, 1, [headers]);
    recovery.format(1, 1, 1, 10, {
      background: '#155E75',
      fontColor: '#FFFFFF',
      fontWeight: 'bold'
    });
  },
  
  buildDenialReports: function(builder) {
    const reports = builder.sheet('Reports');
    
    // Report header
    reports.setValue(1, 1, 'Denial Analytics & Recovery Reports');
    reports.format(1, 1, 1, 10, {
      background: '#155E75',
      fontColor: '#FFFFFF',
      fontSize: 16,
      fontWeight: 'bold'
    });
  }
};

/**
 * Enhanced Template Builder with Professional Features
 * (Reusing from RealEstateTemplate for consistency)
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
        this.sheets[name] = new SheetHelper(sheet, name, this);
      } catch (error) {
        this.errors.push(`Failed to create sheet ${name}: ${error.toString()}`);
      }
    });
    return this.sheets;
  }
  
  sheet(name) {
    return this.sheets[name];
  }
  
  complete() {
    return {
      success: this.errors.length === 0,
      sheets: Object.keys(this.sheets).map(name => 
        this.isPreview ? `[PREVIEW] ${name}` : name
      ),
      errors: this.errors
    };
  }
  
  resolveSheetReference(formula) {
    let resolved = formula;
    for (const [name, helper] of Object.entries(this.sheets)) {
      const placeholder = `{{${name}}}`;
      const actualName = this.isPreview ? `'[PREVIEW] ${name}'` : `'${name}'`;
      resolved = resolved.replace(new RegExp(placeholder, 'g'), actualName);
    }
    return resolved;
  }
}

/**
 * Sheet Helper Class
 */
class SheetHelper {
  constructor(sheet, name, builder) {
    this.sheet = sheet;
    this.name = name;
    this.builder = builder;
  }
  
  setValue(row, col, value) {
    try {
      this.sheet.getRange(row, col).setValue(value);
      return true;
    } catch (error) {
      this.builder.errors.push(`Error setting value at ${this.name}[${row},${col}]: ${error.toString()}`);
      return false;
    }
  }
  
  setFormula(row, col, formula) {
    try {
      const resolved = this.builder.resolveSheetReference(formula);
      this.sheet.getRange(row, col).setFormula(resolved);
      return true;
    } catch (error) {
      this.builder.errors.push(`Error setting formula at ${this.name}[${row},${col}]: ${error.toString()}`);
      return false;
    }
  }
  
  setRangeValues(startRow, startCol, values) {
    try {
      if (!values || values.length === 0) return false;
      const numRows = values.length;
      const numCols = values[0].length;
      this.sheet.getRange(startRow, startCol, numRows, numCols).setValues(values);
      return true;
    } catch (error) {
      this.builder.errors.push(`Error setting range values at ${this.name}[${startRow},${startCol}]: ${error.toString()}`);
      return false;
    }
  }
  
  format(startRow, startCol, endRow, endCol, formatting) {
    try {
      const range = this.sheet.getRange(startRow, startCol, endRow - startRow + 1, endCol - startCol + 1);
      
      if (formatting.background) range.setBackground(formatting.background);
      if (formatting.fontColor) range.setFontColor(formatting.fontColor);
      if (formatting.fontSize) range.setFontSize(formatting.fontSize);
      if (formatting.fontWeight) range.setFontWeight(formatting.fontWeight);
      if (formatting.horizontalAlignment) range.setHorizontalAlignment(formatting.horizontalAlignment);
      if (formatting.verticalAlignment) range.setVerticalAlignment(formatting.verticalAlignment);
      if (formatting.numberFormat) range.setNumberFormat(formatting.numberFormat);
      if (formatting.border) {
        range.setBorder(true, true, true, true, true, true, '#D1D5DB', SpreadsheetApp.BorderStyle.SOLID);
      }
      
      return true;
    } catch (error) {
      this.builder.errors.push(`Error formatting range at ${this.name}: ${error.toString()}`);
      return false;
    }
  }
  
  merge(startRow, startCol, endRow, endCol) {
    try {
      const range = this.sheet.getRange(startRow, startCol, endRow - startRow + 1, endCol - startCol + 1);
      range.merge();
      return true;
    } catch (error) {
      this.builder.errors.push(`Error merging cells at ${this.name}: ${error.toString()}`);
      return false;
    }
  }
  
  addValidation(rangeA1, values) {
    try {
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(values, true)
        .build();
      this.sheet.getRange(rangeA1).setDataValidation(rule);
      return true;
    } catch (error) {
      this.builder.errors.push(`Error adding validation at ${this.name}[${rangeA1}]: ${error.toString()}`);
      return false;
    }
  }
  
  addConditionalFormat(startRow, startCol, endRow, endCol, rule) {
    try {
      const range = this.sheet.getRange(startRow, startCol, endRow - startRow + 1, endCol - startCol + 1);
      const rules = this.sheet.getConditionalFormatRules();
      
      let newRule;
      const builder = SpreadsheetApp.newConditionalFormatRule()
        .setRanges([range]);
      
      if (rule.type === 'cell') {
        switch(rule.condition) {
          case 'NUMBER_GREATER_THAN':
            builder.whenNumberGreaterThan(rule.value);
            break;
          case 'NUMBER_LESS_THAN':
            builder.whenNumberLessThan(rule.value);
            break;
          case 'TEXT_CONTAINS':
            builder.whenTextContains(rule.value);
            break;
          case 'TEXT_EQUALS':
            builder.whenTextEqualTo(rule.value);
            break;
        }
        
        if (rule.format.fontColor) {
          builder.setFontColor(rule.format.fontColor);
        }
        if (rule.format.background) {
          builder.setBackground(rule.format.background);
        }
      } else if (rule.type === 'gradient') {
        builder.setGradientMaxpointWithValue(
          rule.maxColor,
          SpreadsheetApp.InterpolationType.NUMBER,
          rule.maxValue || '100'
        );
        builder.setGradientMidpointWithValue(
          rule.midColor,
          SpreadsheetApp.InterpolationType.NUMBER,
          rule.midValue || '50'
        );
        builder.setGradientMinpointWithValue(
          rule.minColor,
          SpreadsheetApp.InterpolationType.NUMBER,
          rule.minValue || '0'
        );
      }
      
      newRule = builder.build();
      rules.push(newRule);
      this.sheet.setConditionalFormatRules(rules);
      
      return true;
    } catch (error) {
      this.builder.errors.push(`Error adding conditional format at ${this.name}: ${error.toString()}`);
      return false;
    }
  }
}