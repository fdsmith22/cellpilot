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
    
    // Project Information Bar
    dash.setValue(3, 1, 'Project:');
    dash.setValue(3, 2, 'Main Street Development');
    dash.merge(3, 2, 3, 3);
    dash.setValue(3, 4, 'Client:');
    dash.setValue(3, 5, 'ABC Corp');
    dash.merge(3, 5, 3, 6);
    dash.setValue(3, 7, 'Estimator:');
    dash.setValue(3, 8, 'John Smith');
    dash.setValue(3, 10, 'Date:');
    dash.setFormula(3, 11, '=TODAY()');
    dash.format(3, 11, 3, 11, { numberFormat: 'dd/mm/yyyy' });
    
    // Enhanced KPI Cards Row 1
    dash.setValue(5, 1, 'FINANCIAL OVERVIEW');
    dash.merge(5, 1, 5, 12);
    dash.format(5, 1, 5, 12, {
      fontSize: 12,
      fontWeight: 'bold',
      background: '#FEF3C7'
    });
    
    // KPI Cards with more detail
    const kpis = [
      { row: 6, col: 1, title: 'Total Estimate', formula: '=SUM({{Estimates}}!E:E)', subformula: '=TEXT((A7-{{Estimates}}!D2)/{{Estimates}}!D2,"↑ +0.0%")', format: '$#,##0', color: '#DC2626' },
      { row: 6, col: 4, title: 'Material Cost', formula: '=SUM({{Materials}}!G:G)', subformula: '=TEXT(D7/A7,"0.0%")&" of total"', format: '$#,##0', color: '#EA580C' },
      { row: 6, col: 7, title: 'Labor Cost', formula: '=SUM({{Labor}}!H:H)', subformula: '=TEXT(G7/A7,"0.0%")&" of total"', format: '$#,##0', color: '#F59E0B' },
      { row: 6, col: 10, title: 'Profit Margin', formula: '=IFERROR(({{Markup Calculator}}!E10-A7)/{{Markup Calculator}}!E10,0)', subformula: '=IF(J7>0.15,"🟢 Target Met","🔴 Below Target")', format: '0.0%', color: '#10B981' }
    ];
    
    kpis.forEach(kpi => {
      dash.setValue(kpi.row, kpi.col, kpi.title);
      dash.merge(kpi.row, kpi.col, kpi.row, kpi.col + 2);
      dash.setFormula(kpi.row + 1, kpi.col, kpi.formula);
      dash.merge(kpi.row + 1, kpi.col, kpi.row + 1, kpi.col + 2);
      dash.setFormula(kpi.row + 2, kpi.col, kpi.subformula);
      dash.merge(kpi.row + 2, kpi.col, kpi.row + 2, kpi.col + 2);
      
      dash.format(kpi.row, kpi.col, kpi.row, kpi.col + 2, {
        background: '#FED7AA',
        fontWeight: 'bold',
        fontSize: 11
      });
      dash.format(kpi.row + 1, kpi.col, kpi.row + 1, kpi.col + 2, {
        fontSize: 20,
        fontWeight: 'bold',
        horizontalAlignment: 'center',
        numberFormat: kpi.format,
        fontColor: kpi.color
      });
      dash.format(kpi.row + 2, kpi.col, kpi.row + 2, kpi.col + 2, {
        fontSize: 10,
        horizontalAlignment: 'center',
        fontColor: '#6B7280'
      });
    });
    
    // Cost Breakdown by Phase
    dash.setValue(10, 1, 'COST BREAKDOWN BY PHASE');
    dash.merge(10, 1, 10, 6);
    dash.format(10, 1, 10, 6, {
      fontSize: 12,
      fontWeight: 'bold',
      background: '#FEF3C7'
    });
    
    // Phase breakdown table with visual bars
    const phases = [
      ['Phase', 'Budgeted', 'Estimated', 'Variance', '% Complete', 'Visual'],
      ['Site Prep', 45000, '=SUMIF({{Estimates}}!B:B,A12,{{Estimates}}!E:E)', '=C12-B12', '100%', '=REPT("■",10)'],
      ['Foundation', 85000, '=SUMIF({{Estimates}}!B:B,A13,{{Estimates}}!E:E)', '=C13-B13', '75%', '=REPT("■",7)&REPT("□",3)'],
      ['Framing', 120000, '=SUMIF({{Estimates}}!B:B,A14,{{Estimates}}!E:E)', '=C14-B14', '50%', '=REPT("■",5)&REPT("□",5)'],
      ['Exterior', 95000, '=SUMIF({{Estimates}}!B:B,A15,{{Estimates}}!E:E)', '=C15-B15', '25%', '=REPT("■",2)&REPT("□",8)'],
      ['Interior', 110000, '=SUMIF({{Estimates}}!B:B,A16,{{Estimates}}!E:E)', '=C16-B16', '0%', '=REPT("□",10)'],
      ['Mechanical', 75000, '=SUMIF({{Estimates}}!B:B,A17,{{Estimates}}!E:E)', '=C17-B17', '0%', '=REPT("□",10)'],
      ['Finishing', 65000, '=SUMIF({{Estimates}}!B:B,A18,{{Estimates}}!E:E)', '=C18-B18', '0%', '=REPT("□",10)']
    ];
    
    phases.forEach((phase, idx) => {
      phase.forEach((cell, col) => {
        if (typeof cell === 'string' && cell.startsWith('=')) {
          dash.setFormula(11 + idx, col + 1, cell);
        } else {
          dash.setValue(11 + idx, col + 1, cell);
        }
      });
    });
    
    dash.format(11, 1, 11, 6, {
      fontWeight: 'bold',
      background: '#E5E7EB'
    });
    dash.format(12, 2, 18, 4, { numberFormat: '$#,##0' });
    dash.format(12, 5, 18, 5, { numberFormat: '0%' });
    
    // Conditional formatting for variance
    dash.addConditionalFormat(12, 4, 18, 4, {
      type: 'gradient',
      minColor: '#FEE2E2',
      midColor: '#FFFFFF',
      maxColor: '#D1FAE5',
      minValue: '-10000',
      midValue: '0',
      maxValue: '10000'
    });
    
    // Risk Analysis Section
    dash.setValue(10, 7, 'RISK ANALYSIS');
    dash.merge(10, 7, 10, 12);
    dash.format(10, 7, 10, 12, {
      fontSize: 12,
      fontWeight: 'bold',
      background: '#FEF3C7'
    });
    
    const risks = [
      ['Risk Factor', 'Impact', 'Probability', 'Score'],
      ['Material Price Increase', 'High', '30%', '=IF(B12="High",3,IF(B12="Medium",2,1))*VALUE(SUBSTITUTE(C12,"%",""))/100*10'],
      ['Labor Shortage', 'Medium', '40%', '=IF(B13="High",3,IF(B13="Medium",2,1))*VALUE(SUBSTITUTE(C13,"%",""))/100*10'],
      ['Weather Delays', 'Low', '60%', '=IF(B14="High",3,IF(B14="Medium",2,1))*VALUE(SUBSTITUTE(C14,"%",""))/100*10'],
      ['Permit Issues', 'High', '20%', '=IF(B15="High",3,IF(B15="Medium",2,1))*VALUE(SUBSTITUTE(C15,"%",""))/100*10'],
      ['Scope Creep', 'Medium', '50%', '=IF(B16="High",3,IF(B16="Medium",2,1))*VALUE(SUBSTITUTE(C16,"%",""))/100*10']
    ];
    
    risks.forEach((risk, idx) => {
      risk.forEach((cell, col) => {
        if (typeof cell === 'string' && cell.startsWith('=')) {
          dash.setFormula(11 + idx, col + 7, cell);
        } else {
          dash.setValue(11 + idx, col + 7, cell);
        }
      });
    });
    
    dash.format(11, 7, 11, 10, {
      fontWeight: 'bold',
      background: '#E5E7EB'
    });
    
    // Timeline Overview
    dash.setValue(20, 1, 'PROJECT TIMELINE');
    dash.merge(20, 1, 20, 12);
    dash.format(20, 1, 20, 12, {
      fontSize: 12,
      fontWeight: 'bold',
      background: '#FEF3C7'
    });
    
    // Gantt-style timeline
    const timeline = [
      ['Phase', 'Start', 'End', 'Duration', 'W1', 'W2', 'W3', 'W4', 'W5', 'W6', 'W7', 'W8'],
      ['Site Prep', '=TODAY()', '=B22+7', '7 days', '■', '■', '', '', '', '', '', ''],
      ['Foundation', '=C22+1', '=B23+14', '14 days', '', '', '■', '■', '■', '■', '', ''],
      ['Framing', '=C23+1', '=B24+21', '21 days', '', '', '', '', '', '', '■', '■']
    ];
    
    timeline.forEach((row, idx) => {
      row.forEach((cell, col) => {
        if (typeof cell === 'string' && cell.startsWith('=')) {
          dash.setFormula(21 + idx, col + 1, cell);
        } else {
          dash.setValue(21 + idx, col + 1, cell);
        }
      });
    });
    
    dash.format(21, 1, 21, 12, {
      fontWeight: 'bold',
      background: '#E5E7EB'
    });
    dash.format(22, 2, 24, 3, { numberFormat: 'dd/mm/yyyy' });
    
    // Quick Actions
    dash.setValue(26, 1, 'QUICK ACTIONS');
    dash.merge(26, 1, 26, 12);
    dash.format(26, 1, 26, 12, {
      fontSize: 12,
      fontWeight: 'bold',
      background: '#FEF3C7'
    });
    
    const actions = [
      '→ Generate Detailed Report',
      '→ Export to PDF',
      '→ Send to Client',
      '→ Update Prices'
    ];
    
    actions.forEach((action, idx) => {
      dash.setValue(27 + idx, 1, action);
      dash.merge(27 + idx, 1, 27 + idx, 3);
      dash.format(27 + idx, 1, 27 + idx, 3, {
        background: '#DBEAFE',
        border: true
      });
    });
    
    // Column widths
    dash.setColumnWidth(1, 100);
    for (let col = 2; col <= 12; col++) {
      dash.setColumnWidth(col, 80);
    }
    
    // Cost Breakdown Chart Area (simplified from original)
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
    
    // Title
    materials.setValue(1, 1, 'Materials Cost Tracker');
    materials.merge(1, 1, 1, 15);
    materials.format(1, 1, 1, 15, {
      fontSize: 16,
      fontWeight: 'bold',
      horizontalAlignment: 'center',
      background: '#EA580C',
      fontColor: '#FFFFFF'
    });
    
    // Headers with more detail
    const headers = [
      'Item Code', 'Description', 'Category', 'Phase', 'Quantity', 'Unit',
      'Unit Cost', 'Total Cost', 'Markup %', 'Sell Price', 'Supplier', 
      'Lead Time', 'Order Date', 'Delivery Date', 'Status', 'Notes'
    ];
    
    materials.setRangeValues(3, 1, [headers]);
    materials.format(3, 1, 3, headers.length, {
      background: '#FED7AA',
      fontWeight: 'bold',
      border: true
    });
    
    // Enhanced material data with formulas
    const sampleData = [
      ['MAT001', 'Concrete Mix (3000 PSI)', 'Foundation', 'Site Prep', 100, 'cu.yd', 120, '=E4*G4', 25, '=H4*(1+I4/100)', 'BuildSupply Co', '2 days', '=TODAY()', '=M4+2', 'Ordered', 'Grade A, pumping included'],
      ['MAT002', 'Steel Beams W12x26', 'Structure', 'Foundation', 50, 'tons', 800, '=E5*G5', 20, '=H5*(1+I5/100)', 'SteelWorks Inc', '1 week', '=TODAY()+3', '=M5+7', 'Pending', 'I-beams, galvanized'],
      ['MAT003', 'Lumber 2x4x8 PT', 'Framing', 'Framing', 500, 'pieces', 8.5, '=E6*G6', 30, '=H6*(1+I6/100)', 'WoodMart', '3 days', '=TODAY()+7', '=M6+3', 'Not Ordered', 'Pressure treated'],
      ['MAT004', 'Lumber 2x6x10 PT', 'Framing', 'Framing', 300, 'pieces', 12.75, '=E7*G7', 30, '=H7*(1+I7/100)', 'WoodMart', '3 days', '=TODAY()+7', '=M7+3', 'Not Ordered', 'Pressure treated'],
      ['MAT005', 'Plywood 3/4" T&G', 'Framing', 'Framing', 150, 'sheets', 45, '=E8*G8', 25, '=H8*(1+I8/100)', 'WoodMart', '3 days', '=TODAY()+10', '=M8+3', 'Not Ordered', '4x8 sheets, exterior grade'],
      ['MAT006', 'Drywall 1/2" Regular', 'Interior', 'Interior', 200, 'sheets', 12, '=E9*G9', 35, '=H9*(1+I9/100)', 'BuildSupply Co', '2 days', '=TODAY()+30', '=M9+2', 'Not Ordered', '4x8 sheets'],
      ['MAT007', 'Drywall 5/8" Fire Rated', 'Interior', 'Interior', 100, 'sheets', 15, '=E10*G10', 35, '=H10*(1+I10/100)', 'BuildSupply Co', '2 days', '=TODAY()+30', '=M10+2', 'Not Ordered', '4x8 sheets, Type X'],
      ['MAT008', 'Insulation R-19', 'Interior', 'Interior', 5000, 'sq.ft', 0.85, '=E11*G11', 40, '=H11*(1+I11/100)', 'InsulPro', '1 week', '=TODAY()+25', '=M11+7', 'Not Ordered', 'Fiberglass batts'],
      ['MAT009', 'Roofing Shingles', 'Exterior', 'Exterior', 35, 'squares', 125, '=E12*G12', 30, '=H12*(1+I12/100)', 'RoofMaster', '4 days', '=TODAY()+20', '=M12+4', 'Not Ordered', '30-year architectural'],
      ['MAT010', 'Vinyl Siding', 'Exterior', 'Exterior', 2500, 'sq.ft', 3.5, '=E13*G13', 35, '=H13*(1+I13/100)', 'SidingPro', '5 days', '=TODAY()+35', '=M13+5', 'Not Ordered', 'Dutch lap style']
    ];
    
    sampleData.forEach((row, idx) => {
      row.forEach((cell, col) => {
        if (typeof cell === 'string' && cell.startsWith('=')) {
          materials.setFormula(4 + idx, col + 1, cell);
        } else {
          materials.setValue(4 + idx, col + 1, cell);
        }
      });
    });
    
    // Formatting
    materials.format(4, 7, 13, 8, { numberFormat: '$#,##0.00' });
    materials.format(4, 9, 13, 9, { numberFormat: '0%' });
    materials.format(4, 10, 13, 10, { numberFormat: '$#,##0.00' });
    materials.format(4, 13, 13, 14, { numberFormat: 'dd/mm/yyyy' });
    
    // Data validation
    materials.addValidation('C4:C100', ['Foundation', 'Structure', 'Framing', 'Exterior', 'Interior', 'Mechanical', 'Electrical', 'Plumbing', 'Finishing']);
    materials.addValidation('D4:D100', ['Site Prep', 'Foundation', 'Framing', 'Exterior', 'Interior', 'Mechanical', 'Finishing']);
    materials.addValidation('O4:O100', ['Not Ordered', 'Pending', 'Ordered', 'In Transit', 'Delivered', 'Installed']);
    
    // Conditional formatting for status
    materials.addConditionalFormat(4, 15, 100, 15, {
      type: 'cell',
      condition: 'TEXT_EQUALS',
      value: 'Delivered',
      format: { background: '#10B981', fontColor: '#FFFFFF' }
    });
    
    materials.addConditionalFormat(4, 15, 100, 15, {
      type: 'cell',
      condition: 'TEXT_EQUALS',
      value: 'Ordered',
      format: { background: '#3B82F6', fontColor: '#FFFFFF' }
    });
    
    // Summary section
    materials.setValue(16, 1, 'SUMMARY BY CATEGORY');
    materials.merge(16, 1, 16, 6);
    materials.format(16, 1, 16, 6, {
      fontWeight: 'bold',
      background: '#FEF3C7'
    });
    
    const categories = ['Foundation', 'Structure', 'Framing', 'Exterior', 'Interior', 'Mechanical', 'Electrical', 'Plumbing'];
    materials.setValue(17, 1, 'Category');
    materials.setValue(17, 2, 'Items');
    materials.setValue(17, 3, 'Total Cost');
    materials.setValue(17, 4, 'With Markup');
    materials.setValue(17, 5, 'Status');
    materials.setValue(17, 6, 'Visual');
    
    categories.forEach((cat, idx) => {
      materials.setValue(18 + idx, 1, cat);
      materials.setFormula(18 + idx, 2, `=COUNTIF(C:C,"${cat}")`);
      materials.setFormula(18 + idx, 3, `=SUMIF(C:C,"${cat}",H:H)`);
      materials.setFormula(18 + idx, 4, `=SUMIF(C:C,"${cat}",J:J)`);
      materials.setFormula(18 + idx, 5, `=IF(COUNTIFS(C:C,"${cat}",O:O,"Delivered")=B${18+idx},"Complete",COUNTIFS(C:C,"${cat}",O:O,"Delivered")&"/"&B${18+idx})`);
      materials.setFormula(18 + idx, 6, `=REPT("■",ROUND(C${18+idx}/MAX(C$18:C$25)*10,0))`);
    });
    
    materials.format(17, 1, 17, 6, {
      fontWeight: 'bold',
      background: '#E5E7EB'
    });
    materials.format(18, 3, 25, 4, { numberFormat: '$#,##0' });
    
    // Totals row
    materials.setValue(26, 1, 'TOTAL');
    materials.setFormula(26, 2, '=SUM(B18:B25)');
    materials.setFormula(26, 3, '=SUM(C18:C25)');
    materials.setFormula(26, 4, '=SUM(D18:D25)');
    materials.format(26, 1, 26, 6, {
      fontWeight: 'bold',
      background: '#FED7AA',
      border: true
    });
    
    // Column widths
    materials.setColumnWidth(1, 80);
    materials.setColumnWidth(2, 180);
    materials.setColumnWidth(3, 100);
    for (let col = 4; col <= 16; col++) {
      materials.setColumnWidth(col, 90);
    }
  },
  
  buildLaborCost: function(builder) {
    const labor = builder.sheet('Labor');
    
    // Title
    labor.setValue(1, 1, 'Labor Cost Analysis');
    labor.merge(1, 1, 1, 12);
    labor.format(1, 1, 1, 12, {
      fontSize: 16,
      fontWeight: 'bold',
      horizontalAlignment: 'center',
      background: '#EA580C',
      fontColor: '#FFFFFF'
    });
    
    // Summary cards
    labor.setValue(3, 1, 'Total Labor Hours');
    labor.setFormula(3, 2, '=SUM(E5:E30)');
    labor.format(3, 2, 3, 2, { fontSize: 14, fontWeight: 'bold' });
    
    labor.setValue(3, 4, 'Total Labor Cost');
    labor.setFormula(3, 5, '=SUM(H5:H30)');
    labor.format(3, 5, 3, 5, { fontSize: 14, fontWeight: 'bold', numberFormat: '$#,##0' });
    
    labor.setValue(3, 7, 'Avg Hourly Rate');
    labor.setFormula(3, 8, '=AVERAGE(F5:F30)');
    labor.format(3, 8, 3, 8, { fontSize: 14, fontWeight: 'bold', numberFormat: '$#,##0.00' });
    
    labor.setValue(3, 10, 'Crews Active');
    labor.setFormula(3, 11, '=COUNTUNIQUE(B5:B30)');
    labor.format(3, 11, 3, 11, { fontSize: 14, fontWeight: 'bold' });
    
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

// TemplateBuilderPro and SheetHelper classes are now in TemplateBuilder.js
