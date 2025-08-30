/**
 * Professional Construction Templates Suite
 * Complete templates for construction industry professionals
 */

var ConstructionTemplate = {
  /**
   * Build comprehensive Construction template based on type
   */
  build: function(spreadsheet, templateType = 'cost-estimator', isPreview = false) {
    const builder = new TemplateBuilderPro(spreadsheet, isPreview);
    
    switch(templateType) {
      case 'cost-estimator':
        return this.buildCostEstimatorTemplate(builder);
      case 'material-tracker':
        return this.buildMaterialTrackerTemplate(builder);
      case 'labor-manager':
        return this.buildLaborManagerTemplate(builder);
      case 'change-orders':
        return this.buildChangeOrdersTemplate(builder);
      default:
        return this.buildCostEstimatorTemplate(builder);
    }
  },
  
  /**
   * Cost Estimator Template
   */
  buildCostEstimatorTemplate: function(builder) {
    // Create all sheets
    const sheets = builder.createSheets([
      'Dashboard',
      'Estimates',
      'Materials',
      'Labor',
      'Overhead',
      'Markup Calculator',
      'Bid Comparison',
      'Reports'
    ]);
    
    // Build each component
    this.buildCostDashboard(builder);
    this.buildEstimates(builder);
    this.buildMaterialsCost(builder);
    this.buildLaborCost(builder);
    this.buildOverhead(builder);
    this.buildMarkupCalculator(builder);
    this.buildBidComparison(builder);
    this.buildCostReports(builder);
    
    return builder.complete();
  },
  
  /**
   * Material Tracker Template
   */
  buildMaterialTrackerTemplate: function(builder) {
    const sheets = builder.createSheets([
      'Dashboard',
      'Inventory',
      'Orders',
      'Suppliers',
      'Deliveries',
      'Usage Tracking',
      'Waste Analysis',
      'Reports'
    ]);
    
    this.buildMaterialDashboard(builder);
    this.buildInventory(builder);
    this.buildOrders(builder);
    this.buildSuppliers(builder);
    this.buildDeliveries(builder);
    this.buildUsageTracking(builder);
    this.buildWasteAnalysis(builder);
    this.buildMaterialReports(builder);
    
    return builder.complete();
  },
  
  /**
   * Labor Manager Template
   */
  buildLaborManagerTemplate: function(builder) {
    const sheets = builder.createSheets([
      'Dashboard',
      'Crews',
      'Timesheets',
      'Payroll',
      'Productivity',
      'Certifications',
      'Schedule',
      'Reports'
    ]);
    
    this.buildLaborDashboard(builder);
    this.buildCrews(builder);
    this.buildTimesheets(builder);
    this.buildPayroll(builder);
    this.buildProductivity(builder);
    this.buildCertifications(builder);
    this.buildSchedule(builder);
    this.buildLaborReports(builder);
    
    return builder.complete();
  },
  
  /**
   * Change Orders Template
   */
  buildChangeOrdersTemplate: function(builder) {
    const sheets = builder.createSheets([
      'Dashboard',
      'Change Log',
      'Cost Impact',
      'Schedule Impact',
      'Approvals',
      'Documentation',
      'Client Portal',
      'Reports'
    ]);
    
    this.buildChangeDashboard(builder);
    this.buildChangeLog(builder);
    this.buildCostImpact(builder);
    this.buildScheduleImpact(builder);
    this.buildApprovals(builder);
    this.buildDocumentation(builder);
    this.buildClientPortal(builder);
    this.buildChangeReports(builder);
    
    return builder.complete();
  },
  
  // ========================================
  // COST ESTIMATOR TEMPLATE BUILDERS
  // ========================================
  
  buildCostDashboard: function(builder) {
    const dash = builder.sheet('Dashboard');
    
    // Header with construction orange theme
    dash.merge(1, 1, 2, 12);
    dash.setValue(1, 1, 'Construction Cost Estimator Command Center');
    dash.format(1, 1, 2, 12, {
      background: '#EA580C',
      fontColor: '#FFFFFF',
      fontSize: 18,
      fontWeight: 'bold',
      horizontalAlignment: 'center'
    });
    
    // Date and project info
    dash.setValue(3, 1, 'Project:');
    dash.setValue(3, 2, 'Main Street Development');
    dash.setValue(3, 4, 'Estimator:');
    dash.setValue(3, 5, 'John Smith');
    dash.setValue(3, 7, 'Date:');
    dash.setFormula(3, 8, '=TODAY()');
    dash.format(3, 8, 3, 8, { numberFormat: 'dd/mm/yyyy' });
    
    // KPI Cards
    const kpis = [
      { row: 5, col: 1, title: 'Total Estimate', formula: '=SUM({{Materials}}!G:G)+SUM({{Labor}}!H:H)+SUM({{Overhead}}!E:E)', format: '$#,##0', color: '#DC2626' },
      { row: 5, col: 4, title: 'Material Cost', formula: '=SUM({{Materials}}!G:G)', format: '$#,##0', color: '#EA580C' },
      { row: 5, col: 7, title: 'Labor Cost', formula: '=SUM({{Labor}}!H:H)', format: '$#,##0', color: '#F59E0B' },
      { row: 5, col: 10, title: 'Profit Margin', formula: '=IFERROR(({{Markup Calculator}}!E10-{{Estimates}}!D2)/{{Markup Calculator}}!E10,0)', format: '0.0%', color: '#10B981' }
    ];
    
    kpis.forEach(kpi => {
      dash.merge(kpi.row, kpi.col, kpi.row + 2, kpi.col + 2);
      dash.setValue(kpi.row, kpi.col, kpi.title);
      dash.setFormula(kpi.row + 1, kpi.col, kpi.formula);
      dash.format(kpi.row, kpi.col, kpi.row, kpi.col + 2, {
        background: '#FED7AA',
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
    
    // Cost Breakdown Chart Area
    dash.setValue(9, 1, 'Cost Breakdown by Category');
    dash.format(9, 1, 9, 6, {
      fontSize: 14,
      fontWeight: 'bold',
      background: '#FEF3C7'
    });
    
    // Breakdown table
    const breakdown = [
      ['Category', 'Budgeted', 'Estimated', 'Variance', '% of Total', 'Status'],
      ['Materials', '=50000', '=SUM({{Materials}}!G:G)', '=B11-C11', '=C11/C$15', '=IF(D11<0,"Over","Under")'],
      ['Labor', '=35000', '=SUM({{Labor}}!H:H)', '=B12-C12', '=C12/C$15', '=IF(D12<0,"Over","Under")'],
      ['Equipment', '=15000', '=SUM({{Overhead}}!E:E)*0.3', '=B13-C13', '=C13/C$15', '=IF(D13<0,"Over","Under")'],
      ['Overhead', '=10000', '=SUM({{Overhead}}!E:E)*0.7', '=B14-C14', '=C14/C$15', '=IF(D14<0,"Over","Under")'],
      ['Total', '=SUM(B11:B14)', '=SUM(C11:C14)', '=SUM(D11:D14)', '=SUM(E11:E14)', '']
    ];
    
    dash.setRangeValues(10, 1, breakdown);
    dash.format(10, 1, 10, 6, {
      background: '#FB923C',
      fontColor: '#FFFFFF',
      fontWeight: 'bold'
    });
    dash.format(11, 2, 15, 4, { numberFormat: '$#,##0' });
    dash.format(11, 5, 14, 5, { numberFormat: '0%' });
    
    // Conditional formatting for variance
    dash.addConditionalFormat(11, 4, 14, 4, {
      type: 'cell',
      condition: 'NUMBER_LESS_THAN',
      value: 0,
      format: { fontColor: '#DC2626' }
    });
    
    dash.addConditionalFormat(11, 4, 14, 4, {
      type: 'cell',
      condition: 'NUMBER_GREATER_THAN',
      value: 0,
      format: { fontColor: '#10B981' }
    });
  },
  
  buildEstimates: function(builder) {
    const estimates = builder.sheet('Estimates');
    
    // Header
    estimates.merge(1, 1, 1, 10);
    estimates.setValue(1, 1, 'Project Cost Estimates');
    estimates.format(1, 1, 1, 10, {
      background: '#EA580C',
      fontColor: '#FFFFFF',
      fontSize: 14,
      fontWeight: 'bold',
      horizontalAlignment: 'center'
    });
    
    // Estimate details
    const headers = [
      'Estimate ID', 'Date', 'Project Name', 'Client', 'Scope',
      'Total Cost', 'Markup %', 'Total Price', 'Status', 'Notes'
    ];
    
    estimates.setRangeValues(2, 1, [headers]);
    estimates.format(2, 1, 2, 10, {
      background: '#FB923C',
      fontColor: '#FFFFFF',
      fontWeight: 'bold'
    });
    
    // Sample data
    const sampleData = [
      ['EST001', '=TODAY()-10', 'Main Street Development', 'ABC Corp', 'Full construction', 125000, 0.25, '=F3*(1+G3)', 'Pending', 'Awaiting approval'],
      ['EST002', '=TODAY()-5', 'Office Renovation', 'XYZ Ltd', 'Interior renovation', 75000, 0.20, '=F4*(1+G4)', 'Approved', 'Start date confirmed'],
      ['EST003', '=TODAY()-2', 'Warehouse Extension', 'LogiCo', 'Building extension', 200000, 0.22, '=F5*(1+G5)', 'Draft', 'Finalizing scope']
    ];
    
    estimates.setRangeValues(3, 1, sampleData);
    estimates.format(3, 2, 5, 2, { numberFormat: 'dd/mm/yyyy' });
    estimates.format(3, 6, 5, 6, { numberFormat: '$#,##0' });
    estimates.format(3, 7, 5, 7, { numberFormat: '0%' });
    estimates.format(3, 8, 5, 8, { numberFormat: '$#,##0' });
    
    // Add validation for status
    estimates.addValidation('I3:I100', ['Draft', 'Pending', 'Approved', 'Rejected', 'Expired']);
  },
  
  buildMaterialsCost: function(builder) {
    const materials = builder.sheet('Materials');
    
    // Headers
    const headers = [
      'Item Code', 'Description', 'Category', 'Quantity', 'Unit',
      'Unit Cost', 'Total Cost', 'Supplier', 'Lead Time', 'Notes'
    ];
    
    materials.setRangeValues(1, 1, [headers]);
    materials.format(1, 1, 1, 10, {
      background: '#EA580C',
      fontColor: '#FFFFFF',
      fontWeight: 'bold'
    });
    
    // Sample material data
    const sampleData = [
      ['MAT001', 'Concrete Mix', 'Foundation', 100, 'cu.yd', 120, '=D2*F2', 'BuildSupply Co', '2 days', 'Grade A'],
      ['MAT002', 'Steel Beams', 'Structure', 50, 'tons', 800, '=D3*F3', 'SteelWorks Inc', '1 week', 'I-beams'],
      ['MAT003', 'Lumber 2x4', 'Framing', 500, 'pieces', 8, '=D4*F4', 'WoodMart', '3 days', 'Treated'],
      ['MAT004', 'Drywall Sheets', 'Interior', 200, 'sheets', 12, '=D5*F5', 'BuildSupply Co', '2 days', '4x8 sheets']
    ];
    
    materials.setRangeValues(2, 1, sampleData);
    materials.format(2, 6, 5, 7, { numberFormat: '$#,##0' });
    
    // Category validation
    materials.addValidation('C2:C100', ['Foundation', 'Structure', 'Framing', 'Exterior', 'Interior', 'Mechanical', 'Electrical', 'Plumbing']);
  },
  
  buildLaborCost: function(builder) {
    const labor = builder.sheet('Labor');
    
    // Headers
    const headers = [
      'Trade', 'Workers', 'Hours/Day', 'Days', 'Total Hours',
      'Rate/Hour', 'Base Cost', 'Total Cost', 'Overtime', 'Notes'
    ];
    
    labor.setRangeValues(1, 1, [headers]);
    labor.format(1, 1, 1, 10, {
      background: '#EA580C',
      fontColor: '#FFFFFF',
      fontWeight: 'bold'
    });
    
    // Sample labor data
    const sampleData = [
      ['Excavation', 4, 8, 5, '=B2*C2*D2', 45, '=E2*F2', '=G2+I2', 0, 'Site preparation'],
      ['Foundation', 6, 8, 10, '=B3*C3*D3', 55, '=E3*F3', '=G3+I3', 500, 'Concrete work'],
      ['Framing', 8, 8, 15, '=B4*C4*D4', 50, '=E4*F4', '=G4+I4', 1000, 'Wood framing'],
      ['Electrical', 3, 8, 8, '=B5*C5*D5', 65, '=E5*F5', '=G5+I5', 300, 'Wiring and fixtures']
    ];
    
    labor.setRangeValues(2, 1, sampleData);
    labor.format(2, 6, 5, 9, { numberFormat: '$#,##0' });
  },
  
  buildOverhead: function(builder) {
    const overhead = builder.sheet('Overhead');
    
    // Headers
    const headers = [
      'Category', 'Description', 'Monthly Cost', 'Project Duration (months)', 'Total Cost', 'Allocation %', 'Notes'
    ];
    
    overhead.setRangeValues(1, 1, [headers]);
    overhead.format(1, 1, 1, 7, {
      background: '#EA580C',
      fontColor: '#FFFFFF',
      fontWeight: 'bold'
    });
    
    // Overhead items
    const sampleData = [
      ['Insurance', 'General Liability', 2000, 3, '=C2*D2', 0.15, 'Project insurance'],
      ['Equipment', 'Rental - Crane', 5000, 2, '=C3*D3', 0.30, 'Tower crane rental'],
      ['Office', 'Site office trailer', 800, 3, '=C4*D4', 0.10, 'On-site office'],
      ['Permits', 'Building permits', 3000, 1, '=C5*D5', 0.05, 'City permits'],
      ['Utilities', 'Temporary power/water', 500, 3, '=C6*D6', 0.08, 'Site utilities']
    ];
    
    overhead.setRangeValues(2, 1, sampleData);
    overhead.format(2, 3, 6, 5, { numberFormat: '$#,##0' });
    overhead.format(2, 6, 6, 6, { numberFormat: '0%' });
  },
  
  buildMarkupCalculator: function(builder) {
    const markup = builder.sheet('Markup Calculator');
    
    // Title
    markup.merge(1, 1, 1, 6);
    markup.setValue(1, 1, 'Markup & Profit Calculator');
    markup.format(1, 1, 1, 6, {
      background: '#EA580C',
      fontColor: '#FFFFFF',
      fontSize: 14,
      fontWeight: 'bold',
      horizontalAlignment: 'center'
    });
    
    // Calculator
    const calc = [
      ['', '', '', '', '', ''],
      ['Cost Component', 'Amount', '', 'Markup Calculation', 'Value', ''],
      ['Direct Costs:', '', '', '', '', ''],
      ['  Materials', '=SUM({{Materials}}!G:G)', '', 'Subtotal Costs', '=B4+B5+B6', ''],
      ['  Labor', '=SUM({{Labor}}!H:H)', '', 'Overhead %', '15%', ''],
      ['  Subcontractors', 0, '', 'Overhead Amount', '=E4*E5', ''],
      ['Indirect Costs:', '', '', 'Total Cost', '=E4+E6', ''],
      ['  Overhead', '=SUM({{Overhead}}!E:E)', '', 'Desired Profit %', '20%', ''],
      ['Total Cost', '=SUM(B4:B6,B8)', '', 'Markup Amount', '=E7*E8', ''],
      ['', '', '', 'Final Price', '=E7+E9', ''],
      ['', '', '', 'Actual Markup %', '=E9/B9', ''],
      ['', '', '', 'Actual Profit Margin', '=E9/E10', '']
    ];
    
    markup.setRangeValues(2, 1, calc);
    markup.format(4, 2, 9, 2, { numberFormat: '$#,##0' });
    markup.format(4, 5, 10, 5, { numberFormat: '$#,##0' });
    markup.format(5, 5, 5, 5, { numberFormat: '0%' });
    markup.format(8, 5, 8, 5, { numberFormat: '0%' });
    markup.format(11, 5, 12, 5, { numberFormat: '0.0%' });
    
    // Formatting
    markup.format(3, 1, 3, 2, {
      fontWeight: 'bold',
      background: '#FEF3C7'
    });
    markup.format(7, 1, 7, 2, {
      fontWeight: 'bold',
      background: '#FEF3C7'
    });
    markup.format(9, 1, 9, 2, {
      fontWeight: 'bold',
      fontSize: 12,
      background: '#FED7AA'
    });
    markup.format(10, 4, 10, 5, {
      fontWeight: 'bold',
      fontSize: 14,
      background: '#FBBF24',
      fontColor: '#FFFFFF'
    });
  },
  
  buildBidComparison: function(builder) {
    const bids = builder.sheet('Bid Comparison');
    
    // Headers
    const headers = [
      'Bid ID', 'Project', 'Our Bid', 'Competitor 1', 'Competitor 2',
      'Competitor 3', 'Winning Bid', 'Our Rank', 'Variance', 'Notes'
    ];
    
    bids.setRangeValues(1, 1, [headers]);
    bids.format(1, 1, 1, 10, {
      background: '#EA580C',
      fontColor: '#FFFFFF',
      fontWeight: 'bold'
    });
    
    // Sample comparison data
    const sampleData = [
      ['BID001', 'Office Complex', 450000, 475000, 440000, 465000, 440000, 2, '=C2-G2', 'Lost to lower bid'],
      ['BID002', 'Retail Center', 325000, 350000, 360000, 340000, 325000, 1, '=C3-G3', 'Won - best price'],
      ['BID003', 'Warehouse', 280000, 275000, 290000, 285000, 275000, 2, '=C4-G4', 'Close second']
    ];
    
    bids.setRangeValues(2, 1, sampleData);
    bids.format(2, 3, 4, 7, { numberFormat: '$#,##0' });
    bids.format(2, 9, 4, 9, { numberFormat: '$#,##0' });
  },
  
  buildCostReports: function(builder) {
    const reports = builder.sheet('Reports');
    
    // Report header
    reports.merge(1, 1, 1, 10);
    reports.setValue(1, 1, 'Cost Analysis Reports');
    reports.format(1, 1, 1, 10, {
      background: '#EA580C',
      fontColor: '#FFFFFF',
      fontSize: 16,
      fontWeight: 'bold',
      horizontalAlignment: 'center'
    });
    
    // Executive Summary
    reports.setValue(3, 1, 'Executive Summary');
    reports.format(3, 1, 3, 4, {
      fontSize: 14,
      fontWeight: 'bold',
      background: '#FED7AA'
    });
    
    // Summary metrics
    const summary = [
      ['Total Estimated Cost:', '=SUM({{Materials}}!G:G)+SUM({{Labor}}!H:H)+SUM({{Overhead}}!E:E)'],
      ['Target Selling Price:', '={{Markup Calculator}}!E10'],
      ['Expected Profit:', '=B5-B4'],
      ['Profit Margin:', '=B6/B5'],
      ['Materials as % of Total:', '=SUM({{Materials}}!G:G)/B4'],
      ['Labor as % of Total:', '=SUM({{Labor}}!H:H)/B4']
    ];
    
    reports.setRangeValues(4, 1, summary);
    reports.format(4, 2, 6, 2, { numberFormat: '$#,##0' });
    reports.format(7, 2, 9, 2, { numberFormat: '0.0%' });
  },
  
  // ========================================
  // MATERIAL TRACKER TEMPLATE BUILDERS
  // ========================================
  
  buildMaterialDashboard: function(builder) {
    const dash = builder.sheet('Dashboard');
    
    // Header
    dash.merge(1, 1, 2, 12);
    dash.setValue(1, 1, 'Material Tracking & Inventory Control');
    dash.format(1, 1, 2, 12, {
      background: '#DC2626',
      fontColor: '#FFFFFF',
      fontSize: 18,
      fontWeight: 'bold',
      horizontalAlignment: 'center'
    });
    
    // KPIs
    const kpis = [
      { row: 4, col: 1, title: 'Total Inventory Value', formula: '=SUMPRODUCT({{Inventory}}!D:D,{{Inventory}}!F:F)', format: '$#,##0' },
      { row: 4, col: 4, title: 'Items Below Reorder', formula: '=COUNTIFS({{Inventory}}!D:D,"<"&{{Inventory}}!E:E,{{Inventory}}!D:D,">0")', format: '#,##0' },
      { row: 4, col: 7, title: 'Pending Orders', formula: '=COUNTIF({{Orders}}!G:G,"Pending")', format: '#,##0' },
      { row: 4, col: 10, title: 'Waste Rate', formula: '=SUM({{Waste Analysis}}!E:E)/SUM({{Usage Tracking}}!D:D)', format: '0.0%' }
    ];
    
    kpis.forEach(kpi => {
      dash.merge(kpi.row, kpi.col, kpi.row + 2, kpi.col + 2);
      dash.setValue(kpi.row, kpi.col, kpi.title);
      dash.setFormula(kpi.row + 1, kpi.col, kpi.formula);
      dash.format(kpi.row, kpi.col, kpi.row + 2, kpi.col + 2, {
        background: '#FCA5A5',
        fontWeight: 'bold',
        horizontalAlignment: 'center',
        numberFormat: kpi.format
      });
    });
  },
  
  buildInventory: function(builder) {
    const inventory = builder.sheet('Inventory');
    
    // Headers
    const headers = [
      'Item Code', 'Description', 'Category', 'Current Stock', 'Reorder Point',
      'Unit Cost', 'Total Value', 'Location', 'Last Updated', 'Status'
    ];
    
    inventory.setRangeValues(1, 1, [headers]);
    inventory.format(1, 1, 1, 10, {
      background: '#DC2626',
      fontColor: '#FFFFFF',
      fontWeight: 'bold'
    });
  },
  
  buildOrders: function(builder) {
    const orders = builder.sheet('Orders');
    
    // Headers
    const headers = [
      'Order ID', 'Date', 'Supplier', 'Items', 'Total Amount',
      'Expected Delivery', 'Status', 'PO Number', 'Approved By', 'Notes'
    ];
    
    orders.setRangeValues(1, 1, [headers]);
    orders.format(1, 1, 1, 10, {
      background: '#DC2626',
      fontColor: '#FFFFFF',
      fontWeight: 'bold'
    });
  },
  
  buildSuppliers: function(builder) {
    const suppliers = builder.sheet('Suppliers');
    
    // Headers
    const headers = [
      'Supplier ID', 'Company Name', 'Contact', 'Phone', 'Email',
      'Address', 'Payment Terms', 'Lead Time', 'Rating', 'Total Purchases'
    ];
    
    suppliers.setRangeValues(1, 1, [headers]);
    suppliers.format(1, 1, 1, 10, {
      background: '#DC2626',
      fontColor: '#FFFFFF',
      fontWeight: 'bold'
    });
  },
  
  buildDeliveries: function(builder) {
    const deliveries = builder.sheet('Deliveries');
    
    // Headers
    const headers = [
      'Delivery ID', 'Date', 'Order ID', 'Supplier', 'Items Received',
      'Quantity', 'Condition', 'Received By', 'Location', 'Notes'
    ];
    
    deliveries.setRangeValues(1, 1, [headers]);
    deliveries.format(1, 1, 1, 10, {
      background: '#DC2626',
      fontColor: '#FFFFFF',
      fontWeight: 'bold'
    });
  },
  
  buildUsageTracking: function(builder) {
    const usage = builder.sheet('Usage Tracking');
    
    // Headers
    const headers = [
      'Date', 'Project', 'Item Code', 'Quantity Used', 'Purpose',
      'Used By', 'Remaining Stock', 'Cost', 'Notes'
    ];
    
    usage.setRangeValues(1, 1, [headers]);
    usage.format(1, 1, 1, 9, {
      background: '#DC2626',
      fontColor: '#FFFFFF',
      fontWeight: 'bold'
    });
  },
  
  buildWasteAnalysis: function(builder) {
    const waste = builder.sheet('Waste Analysis');
    
    // Headers
    const headers = [
      'Date', 'Item', 'Quantity Wasted', 'Value Lost', 'Waste %',
      'Reason', 'Project', 'Preventable', 'Action Taken', 'Notes'
    ];
    
    waste.setRangeValues(1, 1, [headers]);
    waste.format(1, 1, 1, 10, {
      background: '#DC2626',
      fontColor: '#FFFFFF',
      fontWeight: 'bold'
    });
  },
  
  buildMaterialReports: function(builder) {
    const reports = builder.sheet('Reports');
    
    // Report header
    reports.setValue(1, 1, 'Material Management Reports');
    reports.format(1, 1, 1, 10, {
      background: '#DC2626',
      fontColor: '#FFFFFF',
      fontSize: 16,
      fontWeight: 'bold'
    });
  },
  
  // ========================================
  // LABOR MANAGER TEMPLATE BUILDERS
  // ========================================
  
  buildLaborDashboard: function(builder) {
    const dash = builder.sheet('Dashboard');
    
    // Header
    dash.merge(1, 1, 2, 12);
    dash.setValue(1, 1, 'Labor Management & Productivity Center');
    dash.format(1, 1, 2, 12, {
      background: '#F59E0B',
      fontColor: '#FFFFFF',
      fontSize: 18,
      fontWeight: 'bold',
      horizontalAlignment: 'center'
    });
  },
  
  buildCrews: function(builder) {
    const crews = builder.sheet('Crews');
    
    // Headers
    const headers = [
      'Crew ID', 'Crew Name', 'Foreman', 'Members', 'Trade',
      'Current Project', 'Availability', 'Productivity Score', 'Total Hours YTD', 'Notes'
    ];
    
    crews.setRangeValues(1, 1, [headers]);
    crews.format(1, 1, 1, 10, {
      background: '#F59E0B',
      fontColor: '#FFFFFF',
      fontWeight: 'bold'
    });
  },
  
  buildTimesheets: function(builder) {
    const timesheets = builder.sheet('Timesheets');
    
    // Headers
    const headers = [
      'Date', 'Employee', 'Crew', 'Project', 'Task',
      'Start Time', 'End Time', 'Regular Hours', 'Overtime', 'Status'
    ];
    
    timesheets.setRangeValues(1, 1, [headers]);
    timesheets.format(1, 1, 1, 10, {
      background: '#F59E0B',
      fontColor: '#FFFFFF',
      fontWeight: 'bold'
    });
  },
  
  buildPayroll: function(builder) {
    const payroll = builder.sheet('Payroll');
    
    // Headers
    const headers = [
      'Pay Period', 'Employee', 'Regular Hours', 'OT Hours', 'Rate',
      'Regular Pay', 'OT Pay', 'Total Gross', 'Deductions', 'Net Pay'
    ];
    
    payroll.setRangeValues(1, 1, [headers]);
    payroll.format(1, 1, 1, 10, {
      background: '#F59E0B',
      fontColor: '#FFFFFF',
      fontWeight: 'bold'
    });
  },
  
  buildProductivity: function(builder) {
    const prod = builder.sheet('Productivity');
    
    // Headers
    const headers = [
      'Week', 'Crew', 'Project', 'Planned Hours', 'Actual Hours',
      'Units Completed', 'Productivity Rate', 'Efficiency %', 'Cost/Unit', 'Notes'
    ];
    
    prod.setRangeValues(1, 1, [headers]);
    prod.format(1, 1, 1, 10, {
      background: '#F59E0B',
      fontColor: '#FFFFFF',
      fontWeight: 'bold'
    });
  },
  
  buildCertifications: function(builder) {
    const certs = builder.sheet('Certifications');
    
    // Headers
    const headers = [
      'Employee', 'Certification', 'Type', 'Issue Date', 'Expiry Date',
      'Days Until Expiry', 'Status', 'Renewal Cost', 'Training Required', 'Notes'
    ];
    
    certs.setRangeValues(1, 1, [headers]);
    certs.format(1, 1, 1, 10, {
      background: '#F59E0B',
      fontColor: '#FFFFFF',
      fontWeight: 'bold'
    });
  },
  
  buildSchedule: function(builder) {
    const schedule = builder.sheet('Schedule');
    
    // Headers
    const headers = [
      'Week Starting', 'Project', 'Crew', 'Monday', 'Tuesday',
      'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Total Hours'
    ];
    
    schedule.setRangeValues(1, 1, [headers]);
    schedule.format(1, 1, 1, 10, {
      background: '#F59E0B',
      fontColor: '#FFFFFF',
      fontWeight: 'bold'
    });
  },
  
  buildLaborReports: function(builder) {
    const reports = builder.sheet('Reports');
    
    // Report header
    reports.setValue(1, 1, 'Labor Analytics & Performance Reports');
    reports.format(1, 1, 1, 10, {
      background: '#F59E0B',
      fontColor: '#FFFFFF',
      fontSize: 16,
      fontWeight: 'bold'
    });
  },
  
  // ========================================
  // CHANGE ORDERS TEMPLATE BUILDERS
  // ========================================
  
  buildChangeDashboard: function(builder) {
    const dash = builder.sheet('Dashboard');
    
    // Header
    dash.merge(1, 1, 2, 12);
    dash.setValue(1, 1, 'Change Order Management System');
    dash.format(1, 1, 2, 12, {
      background: '#B91C1C',
      fontColor: '#FFFFFF',
      fontSize: 18,
      fontWeight: 'bold',
      horizontalAlignment: 'center'
    });
  },
  
  buildChangeLog: function(builder) {
    const log = builder.sheet('Change Log');
    
    // Headers
    const headers = [
      'CO Number', 'Date', 'Project', 'Client', 'Description',
      'Reason', 'Cost Impact', 'Schedule Impact', 'Status', 'Notes'
    ];
    
    log.setRangeValues(1, 1, [headers]);
    log.format(1, 1, 1, 10, {
      background: '#B91C1C',
      fontColor: '#FFFFFF',
      fontWeight: 'bold'
    });
  },
  
  buildCostImpact: function(builder) {
    const cost = builder.sheet('Cost Impact');
    
    // Headers
    const headers = [
      'CO Number', 'Original Contract', 'Change Amount', 'New Contract',
      'Material Cost', 'Labor Cost', 'Other Cost', 'Markup', 'Total Impact'
    ];
    
    cost.setRangeValues(1, 1, [headers]);
    cost.format(1, 1, 1, 9, {
      background: '#B91C1C',
      fontColor: '#FFFFFF',
      fontWeight: 'bold'
    });
  },
  
  buildScheduleImpact: function(builder) {
    const schedule = builder.sheet('Schedule Impact');
    
    // Headers
    const headers = [
      'CO Number', 'Original Completion', 'Days Added', 'New Completion',
      'Critical Path Impact', 'Resource Impact', 'Risk Level', 'Mitigation', 'Notes'
    ];
    
    schedule.setRangeValues(1, 1, [headers]);
    schedule.format(1, 1, 1, 9, {
      background: '#B91C1C',
      fontColor: '#FFFFFF',
      fontWeight: 'bold'
    });
  },
  
  buildApprovals: function(builder) {
    const approvals = builder.sheet('Approvals');
    
    // Headers
    const headers = [
      'CO Number', 'Submitted Date', 'Client Approval', 'Approval Date',
      'Architect Approval', 'PM Approval', 'Contract Amended', 'Documents', 'Notes'
    ];
    
    approvals.setRangeValues(1, 1, [headers]);
    approvals.format(1, 1, 1, 9, {
      background: '#B91C1C',
      fontColor: '#FFFFFF',
      fontWeight: 'bold'
    });
  },
  
  buildDocumentation: function(builder) {
    const docs = builder.sheet('Documentation');
    
    // Headers
    const headers = [
      'CO Number', 'Document Type', 'Description', 'File Name',
      'Upload Date', 'Uploaded By', 'Version', 'Status', 'Notes'
    ];
    
    docs.setRangeValues(1, 1, [headers]);
    docs.format(1, 1, 1, 9, {
      background: '#B91C1C',
      fontColor: '#FFFFFF',
      fontWeight: 'bold'
    });
  },
  
  buildClientPortal: function(builder) {
    const portal = builder.sheet('Client Portal');
    
    // Headers
    portal.setValue(1, 1, 'Client Change Order Summary');
    portal.format(1, 1, 1, 8, {
      background: '#B91C1C',
      fontColor: '#FFFFFF',
      fontSize: 14,
      fontWeight: 'bold'
    });
    
    // Summary table headers
    const headers = [
      'CO Number', 'Date', 'Description', 'Cost Impact',
      'Schedule Impact', 'Status', 'Your Approval', 'Comments'
    ];
    
    portal.setRangeValues(3, 1, [headers]);
    portal.format(3, 1, 3, 8, {
      background: '#FCA5A5',
      fontWeight: 'bold'
    });
  },
  
  buildChangeReports: function(builder) {
    const reports = builder.sheet('Reports');
    
    // Report header
    reports.setValue(1, 1, 'Change Order Analysis & Reports');
    reports.format(1, 1, 1, 10, {
      background: '#B91C1C',
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