/**
 * Professional Services Templates Suite
 * Complete templates for consultancies, agencies, and service businesses
 */

var ProfessionalServicesTemplate = {
  /**
   * Build comprehensive Professional Services template based on type
   */
  build: function(spreadsheet, templateType = 'project-tracker', isPreview = false) {
    const builder = new TemplateBuilderPro(spreadsheet, isPreview);
    
    switch(templateType) {
      case 'project-tracker':
        return this.buildProjectTracker(builder);
      case 'time-billing':
        return this.buildTimeBilling(builder);
      case 'client-dashboard':
        return this.buildClientDashboard(builder);
      case 'resource-planner':
        return this.buildResourcePlanner(builder);
      case 'proposal-tracker':
        return this.buildProposalTracker(builder);
      case 'invoice-manager':
        return this.buildInvoiceManager(builder);
      case 'capacity-planning':
        return this.buildCapacityPlanning(builder);
      case 'financial-dashboard':
        return this.buildFinancialDashboard(builder);
      case 'team-performance':
        return this.buildTeamPerformance(builder);
      case 'client-satisfaction':
        return this.buildClientSatisfaction(builder);
      default:
        return this.buildProjectTracker(builder);
    }
  },
  
  /**
   * Project Tracker - Comprehensive project and engagement management
   */
  buildProjectTracker: function(builder) {
    const sheets = builder.createSheets([
      'Dashboard',
      'Projects',
      'Tasks',
      'Resources',
      'Timeline',
      'Budget',
      'Deliverables',
      'Reports'
    ]);
    
    this.buildProjectDashboard(sheets.Dashboard);
    this.buildProjectsSheet(sheets.Projects);
    this.buildTasksSheet(sheets.Tasks);
    this.buildResourcesSheet(sheets.Resources);
    this.buildTimelineSheet(sheets.Timeline);
    this.buildBudgetSheet(sheets.Budget);
    this.buildDeliverablesSheet(sheets.Deliverables);
    this.buildProjectReportsSheet(sheets.Reports);
    
    return builder.complete();
  },
  
  buildProjectDashboard: function(dash) {
    // Title
    dash.setValue(1, 1, 'Professional Services Command Center');
    dash.merge(1, 1, 1, 12);
    dash.format(1, 1, 1, 12, {
      fontSize: 18,
      fontWeight: 'bold',
      horizontalAlignment: 'center',
      background: '#1E40AF',
      fontColor: '#FFFFFF'
    });
    
    // Quick Filters
    dash.setValue(2, 1, 'View:');
    dash.setValue(2, 2, 'All Projects');
    dash.addValidation('B2', ['All Projects', 'Active Only', 'At Risk', 'Completed', 'My Projects']);
    dash.setValue(2, 5, 'Period:');
    dash.setValue(2, 6, 'Current Quarter');
    dash.addValidation('F2', ['This Month', 'Current Quarter', 'YTD', 'All Time']);
    
    // Project Health KPIs
    dash.setValue(4, 1, 'PROJECT HEALTH');
    dash.merge(4, 1, 4, 12);
    dash.format(4, 1, 4, 12, {
      fontSize: 12,
      fontWeight: 'bold',
      background: '#DBEAFE'
    });
    
    const healthKpis = [
      { row: 5, col: 1, title: 'Active Projects', formula: '=COUNTIF({{Projects}}!F:F,"Active")', sub: '=TEXT((A6-15)/15,"+0;-0")&" vs capacity"' },
      { row: 5, col: 4, title: 'On Time %', formula: '=COUNTIFS({{Projects}}!F:F,"Active",{{Projects}}!L:L,"On Track")/COUNTIF({{Projects}}!F:F,"Active")', sub: '=IF(D6>0.8,"🟢 Excellent","⚠️ Review Needed")' },
      { row: 5, col: 7, title: 'On Budget %', formula: '=COUNTIFS({{Projects}}!F:F,"Active",{{Projects}}!M:M,"Under Budget")/COUNTIF({{Projects}}!F:F,"Active")', sub: '=TEXT(G6,"0%")&" projects"' },
      { row: 5, col: 10, title: 'Client Satisfaction', formula: '=AVERAGE({{Projects}}!N:N)', sub: '=REPT("⭐",ROUND(J6,0))&" avg rating"' }
    ];
    
    healthKpis.forEach(kpi => {
      dash.setValue(kpi.row, kpi.col, kpi.title);
      dash.merge(kpi.row, kpi.col, kpi.row, kpi.col + 2);
      dash.setFormula(kpi.row + 1, kpi.col, kpi.formula);
      dash.merge(kpi.row + 1, kpi.col, kpi.row + 1, kpi.col + 2);
      dash.setFormula(kpi.row + 2, kpi.col, kpi.sub);
      dash.merge(kpi.row + 2, kpi.col, kpi.row + 2, kpi.col + 2);
      
      dash.format(kpi.row, kpi.col, kpi.row, kpi.col + 2, {
        background: '#93C5FD',
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
    dash.format(6, 4, 6, 6, { numberFormat: '0%' });
    dash.format(6, 7, 6, 9, { numberFormat: '0%' });
    dash.format(6, 10, 6, 12, { numberFormat: '0.0' });
    
    // Resource Utilization
    dash.setValue(9, 1, 'RESOURCE UTILIZATION');
    dash.merge(9, 1, 9, 6);
    dash.format(9, 1, 9, 6, {
      fontSize: 12,
      fontWeight: 'bold',
      background: '#DBEAFE'
    });
    
    const resources = [
      ['Resource', 'Capacity', 'Allocated', 'Utilization', 'Efficiency', 'Status'],
      ['Senior Consultants', 160, '=SUMIF({{Resources}}!C:C,"Senior",{{Resources}}!E:E)', '=C11/B11', '=SUMIF({{Resources}}!C:C,"Senior",{{Resources}}!F:F)/COUNTIF({{Resources}}!C:C,"Senior")', '=IF(D11>0.9,"Overloaded",IF(D11>0.7,"Optimal","Available"))'],
      ['Project Managers', 120, '=SUMIF({{Resources}}!C:C,"PM",{{Resources}}!E:E)', '=C12/B12', '=SUMIF({{Resources}}!C:C,"PM",{{Resources}}!F:F)/COUNTIF({{Resources}}!C:C,"PM")', '=IF(D12>0.9,"Overloaded",IF(D12>0.7,"Optimal","Available"))'],
      ['Junior Consultants', 200, '=SUMIF({{Resources}}!C:C,"Junior",{{Resources}}!E:E)', '=C13/B13', '=SUMIF({{Resources}}!C:C,"Junior",{{Resources}}!F:F)/COUNTIF({{Resources}}!C:C,"Junior")', '=IF(D13>0.9,"Overloaded",IF(D13>0.7,"Optimal","Available"))'],
      ['Specialists', 80, '=SUMIF({{Resources}}!C:C,"Specialist",{{Resources}}!E:E)', '=C14/B14', '=SUMIF({{Resources}}!C:C,"Specialist",{{Resources}}!F:F)/COUNTIF({{Resources}}!C:C,"Specialist")', '=IF(D14>0.9,"Overloaded",IF(D14>0.7,"Optimal","Available"))'],
      ['Support Staff', 60, '=SUMIF({{Resources}}!C:C,"Support",{{Resources}}!E:E)', '=C15/B15', '=SUMIF({{Resources}}!C:C,"Support",{{Resources}}!F:F)/COUNTIF({{Resources}}!C:C,"Support")', '=IF(D15>0.9,"Overloaded",IF(D15>0.7,"Optimal","Available"))']
    ];
    
    resources.forEach((row, idx) => {
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
    dash.format(11, 2, 15, 3, { numberFormat: '#,##0"h"' });
    dash.format(11, 4, 15, 5, { numberFormat: '0%' });
    
    // Project Pipeline
    dash.setValue(9, 7, 'PROJECT PIPELINE');
    dash.merge(9, 7, 9, 12);
    dash.format(9, 7, 9, 12, {
      fontSize: 12,
      fontWeight: 'bold',
      background: '#DBEAFE'
    });
    
    const pipeline = [
      ['Stage', 'Projects', 'Value', 'Probability', 'Weighted', 'Avg Days'],
      ['Proposal', '=COUNTIF({{Projects}}!F:F,"Proposal")', '=SUMIF({{Projects}}!F:F,"Proposal",{{Projects}}!I:I)', '60%', '=I11*J11', '15'],
      ['Negotiation', '=COUNTIF({{Projects}}!F:F,"Negotiation")', '=SUMIF({{Projects}}!F:F,"Negotiation",{{Projects}}!I:I)', '80%', '=I12*J12', '8'],
      ['Signed', '=COUNTIF({{Projects}}!F:F,"Signed")', '=SUMIF({{Projects}}!F:F,"Signed",{{Projects}}!I:I)', '95%', '=I13*J13', '3'],
      ['Active', '=COUNTIF({{Projects}}!F:F,"Active")', '=SUMIF({{Projects}}!F:F,"Active",{{Projects}}!I:I)', '100%', '=I14*J14', 'Ongoing'],
      ['At Risk', '=COUNTIF({{Projects}}!L:L,"At Risk")', '=SUMIFS({{Projects}}!I:I,{{Projects}}!L:L,"At Risk")', '70%', '=I15*J15', 'Various']
    ];
    
    pipeline.forEach((row, idx) => {
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
    dash.format(11, 9, 15, 9, { numberFormat: '$#,##0' });
    dash.format(11, 10, 15, 10, { numberFormat: '0%' });
    dash.format(11, 11, 15, 11, { numberFormat: '$#,##0' });
    
    // Current Projects At Risk
    dash.setValue(17, 1, 'PROJECTS REQUIRING ATTENTION');
    dash.merge(17, 1, 17, 12);
    dash.format(17, 1, 17, 12, {
      fontSize: 12,
      fontWeight: 'bold',
      background: '#DBEAFE'
    });
    
    dash.setValue(18, 1, 'Project');
    dash.merge(18, 1, 18, 3);
    dash.setValue(18, 4, 'Client');
    dash.setValue(18, 5, 'Issue');
    dash.merge(18, 5, 18, 7);
    dash.setValue(18, 8, 'Due Date');
    dash.setValue(18, 9, 'Status');
    dash.setValue(18, 10, 'Owner');
    dash.setValue(18, 11, 'Action');
    dash.merge(18, 11, 18, 12);
    
    // Sample at-risk projects
    const riskProjects = [
      ['Website Redesign Project', 'TechCorp Inc', 'Behind schedule by 5 days', '=TODAY()+7', 'At Risk', 'Sarah PM', 'Resource reallocation needed'],
      ['Digital Transformation', 'Global Manufacturing', 'Budget variance +15%', '=TODAY()+14', 'Over Budget', 'Mike Lead', 'Client approval for additional scope'],
      ['Process Optimization', 'Healthcare Systems', 'Client feedback pending', '=TODAY()+3', 'Delayed', 'Lisa Consultant', 'Escalate to client sponsor'],
      ['Data Analytics Platform', 'RetailChain Co', 'Technical issues discovered', '=TODAY()+21', 'At Risk', 'David Tech', 'Additional specialist required']
    ];
    
    riskProjects.forEach((project, idx) => {
      project.forEach((cell, col) => {
        if (typeof cell === 'string' && cell.startsWith('=')) {
          dash.setFormula(19 + idx, col + 1, cell);
        } else {
          dash.setValue(19 + idx, col + 1, cell);
        }
      });
      dash.merge(19 + idx, 1, 19 + idx, 3);
      dash.merge(19 + idx, 5, 19 + idx, 7);
      dash.merge(19 + idx, 11, 19 + idx, 12);
    });
    
    dash.format(18, 1, 18, 12, {
      fontWeight: 'bold',
      background: '#E5E7EB'
    });
    dash.format(19, 8, 22, 8, { numberFormat: 'dd/mm/yyyy' });
    
    // Financial Overview
    dash.setValue(24, 1, 'FINANCIAL OVERVIEW');
    dash.merge(24, 1, 24, 6);
    dash.format(24, 1, 24, 6, {
      fontSize: 12,
      fontWeight: 'bold',
      background: '#DBEAFE'
    });
    
    const financials = [
      ['Revenue (YTD)', '=SUM({{Budget}}!H:H)'],
      ['Pipeline Value', '=SUM(K11:K15)'],
      ['Utilization Rate', '=SUM(C11:C15)/SUM(B11:B15)'],
      ['Avg Project Margin', '=AVERAGE({{Projects}}!O:O)'],
      ['Outstanding Invoices', '=COUNTIF({{Budget}}!K:K,"Outstanding")'],
      ['Cash Flow (30d)', '=SUMIFS({{Budget}}!H:H,{{Budget}}!G:G,">="&TODAY()-30)']
    ];
    
    financials.forEach((item, idx) => {
      dash.setValue(25 + idx, 1, item[0]);
      dash.setFormula(25 + idx, 2, item[1]);
      dash.merge(25 + idx, 2, 25 + idx, 3);
    });
    
    dash.format(25, 2, 26, 3, { numberFormat: '$#,##0' });
    dash.format(27, 2, 27, 3, { numberFormat: '0%' });
    dash.format(28, 2, 28, 3, { numberFormat: '0%' });
    dash.format(30, 2, 30, 3, { numberFormat: '$#,##0' });
    
    // Team Performance
    dash.setValue(24, 7, 'TOP PERFORMERS');
    dash.merge(24, 7, 24, 12);
    dash.format(24, 7, 24, 12, {
      fontSize: 12,
      fontWeight: 'bold',
      background: '#DBEAFE'
    });
    
    dash.setValue(25, 7, 'Team Member');
    dash.merge(25, 7, 25, 8);
    dash.setValue(25, 9, 'Utilization');
    dash.setValue(25, 10, 'Efficiency');
    dash.setValue(25, 11, 'Rating');
    dash.setValue(25, 12, 'Projects');
    
    // Sample top performers
    for (let i = 0; i < 5; i++) {
      dash.setFormula(26 + i, 7, `=INDEX({{Resources}}!B:B,MATCH(LARGE({{Resources}}!F:F,${i+1}),{{Resources}}!F:F,0))`);
      dash.merge(26 + i, 7, 26 + i, 8);
      dash.setFormula(26 + i, 9, `=LARGE({{Resources}}!E:E,${i+1})/INDEX({{Resources}}!D:D,MATCH(LARGE({{Resources}}!F:F,${i+1}),{{Resources}}!F:F,0))`);
      dash.setFormula(26 + i, 10, `=LARGE({{Resources}}!F:F,${i+1})`);
      dash.setValue(26 + i, 11, 4.2 + Math.random() * 0.7);
      dash.setValue(26 + i, 12, Math.floor(Math.random() * 5 + 2));
    }
    
    dash.format(25, 7, 25, 12, {
      fontWeight: 'bold',
      background: '#E5E7EB'
    });
    dash.format(26, 9, 30, 10, { numberFormat: '0%' });
    dash.format(26, 11, 30, 11, { numberFormat: '0.0' });
    
    // Column widths
    dash.setColumnWidth(1, 100);
    for (let col = 2; col <= 12; col++) {
      dash.setColumnWidth(col, 85);
    }
  },
  
  buildProjectsSheet: function(projects) {
    // Project master list with comprehensive tracking
    const headers = [
      'Project ID', 'Project Name', 'Client', 'Type', 'Priority', 'Status',
      'Start Date', 'End Date', 'Budget', 'Spent', 'Remaining', 'Health',
      'Budget Status', 'Client Rating', 'Margin %', 'PM', 'Team Size', 'Notes'
    ];
    
    headers.forEach((header, idx) => {
      projects.setValue(1, idx + 1, header);
    });
    
    projects.format(1, 1, 1, headers.length, {
      fontWeight: 'bold',
      background: '#1E40AF',
      fontColor: '#FFFFFF',
      border: true
    });
    
    // Sample projects
    const sampleProjects = [
      ['PRJ001', 'Digital Transformation Strategy', 'Global Manufacturing', 'Strategy', 'High', 'Active', '=TODAY()-45', '=TODAY()+30', 250000, 180000, '=I2-J2', 'On Track', 'Under Budget', 4.5, 28, 'Sarah Mitchell', 8, 'Large enterprise transformation'],
      ['PRJ002', 'Website Redesign & Development', 'TechCorp Inc', 'Development', 'Medium', 'At Risk', '=TODAY()-30', '=TODAY()+7', 85000, 72000, '=I3-J3', 'At Risk', 'On Budget', 4.0, 15, 'Mike Johnson', 5, 'Behind schedule due to client feedback delays'],
      ['PRJ003', 'Process Optimization Consulting', 'Healthcare Systems', 'Process', 'High', 'Active', '=TODAY()-60', '=TODAY()+45', 180000, 95000, '=I4-J4', 'On Track', 'Under Budget', 4.8, 35, 'Lisa Chen', 6, 'Excellent client collaboration'],
      ['PRJ004', 'Data Analytics Platform', 'RetailChain Co', 'Technology', 'Critical', 'At Risk', '=TODAY()-35', '=TODAY()+21', 320000, 245000, '=I5-J5', 'At Risk', 'Over Budget', 3.8, 22, 'David Park', 7, 'Technical challenges with legacy systems'],
      ['PRJ005', 'Marketing Automation Setup', 'StartupXYZ', 'Marketing', 'Low', 'Active', '=TODAY()-20', '=TODAY()+40', 45000, 18000, '=I6-J6', 'On Track', 'Under Budget', 4.2, 40, 'Jennifer Wu', 3, 'Small but strategic client']
    ];
    
    sampleProjects.forEach((project, idx) => {
      project.forEach((value, col) => {
        if (typeof value === 'string' && value.startsWith('=')) {
          projects.setFormula(idx + 2, col + 1, value);
        } else {
          projects.setValue(idx + 2, col + 1, value);
        }
      });
    });
    
    // Data validation
    projects.addValidation('D2:D100', ['Strategy', 'Development', 'Process', 'Technology', 'Marketing', 'Training', 'Research']);
    projects.addValidation('E2:E100', ['Low', 'Medium', 'High', 'Critical']);
    projects.addValidation('F2:F100', ['Proposal', 'Negotiation', 'Signed', 'Active', 'At Risk', 'Completed', 'On Hold', 'Cancelled']);
    projects.addValidation('L2:L100', ['On Track', 'At Risk', 'Delayed', 'Blocked']);
    projects.addValidation('M2:M100', ['Under Budget', 'On Budget', 'Over Budget']);
    
    // Conditional formatting
    projects.addConditionalFormat(2, 6, 100, 6, {
      type: 'cell',
      condition: 'TEXT_EQUALS',
      value: 'Active',
      format: { background: '#10B981', fontColor: '#FFFFFF' }
    });
    
    projects.addConditionalFormat(2, 6, 100, 6, {
      type: 'cell',
      condition: 'TEXT_EQUALS',
      value: 'At Risk',
      format: { background: '#F59E0B', fontColor: '#FFFFFF' }
    });
    
    projects.addConditionalFormat(2, 12, 100, 12, {
      type: 'cell',
      condition: 'TEXT_EQUALS',
      value: 'At Risk',
      format: { background: '#FEE2E2', fontColor: '#DC2626' }
    });
    
    // Formatting
    projects.format(2, 7, 100, 8, { numberFormat: 'dd/mm/yyyy' });
    projects.format(2, 9, 100, 11, { numberFormat: '$#,##0' });
    projects.format(2, 14, 100, 14, { numberFormat: '0.0' });
    projects.format(2, 15, 100, 15, { numberFormat: '0%' });
    
    // Column widths
    projects.setColumnWidth(1, 80);
    projects.setColumnWidth(2, 200);
    projects.setColumnWidth(3, 150);
    for (let col = 4; col <= headers.length; col++) {
      projects.setColumnWidth(col, 100);
    }
  },
  
  buildTasksSheet: function(tasks) {
    // Task management and tracking
    const headers = [
      'Task ID', 'Project ID', 'Task Name', 'Assigned To', 'Priority',
      'Status', 'Start Date', 'Due Date', 'Estimated Hours', 'Actual Hours',
      'Progress %', 'Dependencies', 'Notes', 'Deliverable', 'Client Facing'
    ];
    
    headers.forEach((header, idx) => {
      tasks.setValue(1, idx + 1, header);
    });
    
    tasks.format(1, 1, 1, headers.length, {
      fontWeight: 'bold',
      background: '#1E40AF',
      fontColor: '#FFFFFF',
      border: true
    });
    
    // Sample tasks
    const sampleTasks = [
      ['T001', 'PRJ001', 'Stakeholder Interviews', 'Sarah Mitchell', 'High', 'Completed', '=TODAY()-45', '=TODAY()-40', 24, 22, 100, 'None', 'Completed ahead of schedule', 'Interview Report', 'Yes'],
      ['T002', 'PRJ001', 'Current State Analysis', 'Mike Johnson', 'High', 'Completed', '=TODAY()-38', '=TODAY()-30', 40, 38, 100, 'T001', 'Good analysis depth', 'Current State Document', 'No'],
      ['T003', 'PRJ001', 'Future State Design', 'Lisa Chen', 'High', 'In Progress', '=TODAY()-25', '=TODAY()+5', 32, 20, 65, 'T002', 'On track for completion', 'Design Blueprints', 'Yes'],
      ['T004', 'PRJ002', 'UI/UX Design', 'Jennifer Wu', 'Medium', 'In Progress', '=TODAY()-30', '=TODAY()-5', 60, 58, 95, 'None', 'Waiting for client feedback', 'Design Mockups', 'Yes'],
      ['T005', 'PRJ002', 'Frontend Development', 'David Park', 'High', 'Not Started', '=TODAY()-5', '=TODAY()+14', 80, 0, 0, 'T004', 'Waiting for design approval', 'Website Frontend', 'No'],
      ['T006', 'PRJ003', 'Process Mapping', 'Lisa Chen', 'High', 'Completed', '=TODAY()-60', '=TODAY()-50', 32, 35, 100, 'None', 'Thorough analysis completed', 'Process Maps', 'No'],
      ['T007', 'PRJ003', 'Gap Analysis', 'Mike Johnson', 'Medium', 'In Progress', '=TODAY()-45', '=TODAY()+10', 24, 18, 75, 'T006', 'Good progress made', 'Gap Analysis Report', 'Yes'],
      ['T008', 'PRJ004', 'Data Architecture Review', 'David Park', 'Critical', 'Blocked', '=TODAY()-35', '=TODAY()-10', 40, 32, 80, 'None', 'Blocked by legacy system access', 'Architecture Document', 'No'],
      ['T009', 'PRJ005', 'Marketing Platform Setup', 'Jennifer Wu', 'Low', 'In Progress', '=TODAY()-20', '=TODAY()+20', 16, 8, 50, 'None', 'Steady progress', 'Platform Configuration', 'No']
    ];
    
    sampleTasks.forEach((task, idx) => {
      task.forEach((value, col) => {
        if (typeof value === 'string' && value.startsWith('=')) {
          tasks.setFormula(idx + 2, col + 1, value);
        } else {
          tasks.setValue(idx + 2, col + 1, value);
        }
      });
    });
    
    // Data validation
    tasks.addValidation('E2:E1000', ['Low', 'Medium', 'High', 'Critical']);
    tasks.addValidation('F2:F1000', ['Not Started', 'In Progress', 'Blocked', 'Under Review', 'Completed']);
    tasks.addValidation('O2:O1000', ['Yes', 'No']);
    
    // Conditional formatting
    tasks.addConditionalFormat(2, 6, 1000, 6, {
      type: 'cell',
      condition: 'TEXT_EQUALS',
      value: 'Completed',
      format: { background: '#D1FAE5' }
    });
    
    tasks.addConditionalFormat(2, 6, 1000, 6, {
      type: 'cell',
      condition: 'TEXT_EQUALS',
      value: 'Blocked',
      format: { background: '#FEE2E2', fontColor: '#DC2626' }
    });
    
    // Progress bar formatting
    tasks.addConditionalFormat(2, 11, 1000, 11, {
      type: 'gradient',
      minColor: '#FEE2E2',
      midColor: '#FEF3C7',
      maxColor: '#D1FAE5',
      minValue: '0',
      midValue: '50',
      maxValue: '100'
    });
    
    // Formatting
    tasks.format(2, 7, 1000, 8, { numberFormat: 'dd/mm/yyyy' });
    tasks.format(2, 11, 1000, 11, { numberFormat: '0%' });
    
    // Column widths
    tasks.setColumnWidth(1, 80);
    tasks.setColumnWidth(2, 80);
    tasks.setColumnWidth(3, 200);
    tasks.setColumnWidth(4, 120);
    for (let col = 5; col <= headers.length; col++) {
      tasks.setColumnWidth(col, 100);
    }
  },
  
  buildResourcesSheet: function(resources) {
    // Resource allocation and capacity tracking
    const headers = [
      'Resource ID', 'Name', 'Role', 'Capacity (h/week)', 'Allocated Hours',
      'Efficiency %', 'Hourly Rate', 'Current Projects', 'Skills', 'Availability',
      'Location', 'Manager', 'Performance Rating', 'Certifications', 'Notes'
    ];
    
    headers.forEach((header, idx) => {
      resources.setValue(1, idx + 1, header);
    });
    
    resources.format(1, 1, 1, headers.length, {
      fontWeight: 'bold',
      background: '#1E40AF',
      fontColor: '#FFFFFF',
      border: true
    });
    
    // Sample resources
    const sampleResources = [
      ['RES001', 'Sarah Mitchell', 'Senior', 40, 38, 92, 150, 'PRJ001, PRJ003', 'Strategy, Leadership', 'Full-time', 'New York', 'Executive Team', 4.8, 'PMP, MBA', 'Top performer, excellent client skills'],
      ['RES002', 'Mike Johnson', 'PM', 40, 35, 88, 120, 'PRJ002, PRJ007', 'Project Management, Agile', 'Full-time', 'San Francisco', 'Sarah Mitchell', 4.5, 'PMP, CSM', 'Strong technical background'],
      ['RES003', 'Lisa Chen', 'Senior', 40, 42, 95, 140, 'PRJ001, PRJ003', 'Process Optimization, Change Mgmt', 'Full-time', 'Chicago', 'Sarah Mitchell', 4.7, 'Six Sigma Black Belt', 'Excellent analytical skills'],
      ['RES004', 'David Park', 'Specialist', 40, 36, 85, 130, 'PRJ004, PRJ002', 'Data Analytics, Technology', 'Full-time', 'Austin', 'Mike Johnson', 4.2, 'AWS Certified, Tableau', 'Strong technical expertise'],
      ['RES005', 'Jennifer Wu', 'Junior', 40, 32, 82, 85, 'PRJ005, PRJ002', 'Marketing, Design, Development', 'Full-time', 'Seattle', 'Lisa Chen', 4.0, 'Google Ads, HubSpot', 'Fast learner, great potential'],
      ['RES006', 'Robert Kim', 'Junior', 40, 28, 78, 80, 'PRJ006', 'Research, Analysis', 'Part-time', 'Remote', 'David Park', 3.8, 'None', 'New hire, developing skills'],
      ['RES007', 'Amanda Foster', 'Support', 30, 25, 88, 65, 'Admin Tasks', 'Admin, Coordination', 'Full-time', 'New York', 'Sarah Mitchell', 4.3, 'Project Coordinator Cert', 'Excellent organizational skills'],
      ['RES008', 'Carlos Rodriguez', 'Specialist', 40, 40, 90, 135, 'PRJ008', 'Finance, Compliance', 'Full-time', 'Miami', 'Executive Team', 4.6, 'CPA, CFA', 'Financial analysis expert']
    ];
    
    sampleResources.forEach((resource, idx) => {
      resource.forEach((value, col) => {
        resources.setValue(idx + 2, col + 1, value);
      });
    });
    
    // Data validation
    resources.addValidation('C2:C1000', ['Senior', 'PM', 'Junior', 'Specialist', 'Support']);
    resources.addValidation('J2:J1000', ['Full-time', 'Part-time', 'Contract', 'On Leave']);
    resources.addValidation('K2:K1000', ['New York', 'San Francisco', 'Chicago', 'Austin', 'Seattle', 'Miami', 'Remote']);
    
    // Conditional formatting for utilization
    resources.addConditionalFormat(2, 6, 1000, 6, {
      type: 'gradient',
      minColor: '#FEE2E2',
      midColor: '#FEF3C7',
      maxColor: '#D1FAE5',
      minValue: '60',
      midValue: '80',
      maxValue: '100'
    });
    
    // Formatting
    resources.format(2, 6, 1000, 6, { numberFormat: '0%' });
    resources.format(2, 7, 1000, 7, { numberFormat: '$#,##0' });
    resources.format(2, 13, 1000, 13, { numberFormat: '0.0' });
    
    // Summary section
    resources.setValue(11, 1, 'RESOURCE SUMMARY');
    resources.merge(11, 1, 11, 6);
    resources.format(11, 1, 11, 6, {
      fontWeight: 'bold',
      background: '#DBEAFE'
    });
    
    const summaryMetrics = [
      ['Total Capacity:', '=SUM(D2:D9)'],
      ['Total Allocated:', '=SUM(E2:E9)'],
      ['Overall Utilization:', '=E12/D12'],
      ['Avg Efficiency:', '=AVERAGE(F2:F9)'],
      ['Avg Hourly Rate:', '=AVERAGE(G2:G9)'],
      ['Team Performance:', '=AVERAGE(M2:M9)']
    ];
    
    summaryMetrics.forEach((metric, idx) => {
      resources.setValue(12 + idx, 1, metric[0]);
      resources.setFormula(12 + idx, 2, metric[1]);
    });
    
    resources.format(12, 2, 13, 2, { numberFormat: '#,##0"h"' });
    resources.format(14, 2, 16, 2, { numberFormat: '0%' });
    resources.format(16, 2, 16, 2, { numberFormat: '$#,##0' });
    resources.format(17, 2, 17, 2, { numberFormat: '0.0' });
    
    // Column widths
    resources.setColumnWidth(1, 80);
    resources.setColumnWidth(2, 120);
    resources.setColumnWidth(8, 150);
    resources.setColumnWidth(9, 200);
    resources.setColumnWidth(15, 200);
    for (let col = 3; col <= 7; col++) {
      resources.setColumnWidth(col, 90);
    }
    for (let col = 10; col <= 14; col++) {
      resources.setColumnWidth(col, 100);
    }
  },
  
  buildTimelineSheet: function(timeline) {
    // Project timeline and Gantt-style view
    timeline.setValue(1, 1, 'Project Timeline & Milestones');
    timeline.merge(1, 1, 1, 20);
    timeline.format(1, 1, 1, 20, {
      fontSize: 16,
      fontWeight: 'bold',
      horizontalAlignment: 'center',
      background: '#1E40AF',
      fontColor: '#FFFFFF'
    });
    
    // Timeline headers
    const headers = [
      'Project', 'Phase', 'Start Date', 'End Date', 'Duration',
      'Status', 'Progress', 'Dependencies', 'W1', 'W2', 'W3', 'W4',
      'W5', 'W6', 'W7', 'W8', 'W9', 'W10', 'W11', 'W12'
    ];
    
    headers.forEach((header, idx) => {
      timeline.setValue(3, idx + 1, header);
    });
    
    timeline.format(3, 1, 3, headers.length, {
      fontWeight: 'bold',
      background: '#93C5FD',
      border: true
    });
    
    // Sample timeline data
    const timelineData = [
      ['Digital Transformation', 'Discovery', '=TODAY()-45', '=TODAY()-30', '=D4-C4', 'Completed', 100, 'None', '■', '■', '■', '', '', '', '', '', '', '', '', ''],
      ['Digital Transformation', 'Analysis', '=TODAY()-30', '=TODAY()-15', '=D5-C5', 'Completed', 100, 'Discovery', '', '', '■', '■', '■', '', '', '', '', '', '', ''],
      ['Digital Transformation', 'Design', '=TODAY()-15', '=TODAY()+15', '=D6-C6', 'In Progress', 65, 'Analysis', '', '', '', '', '■', '■', '■', '■', '', '', '', ''],
      ['Website Redesign', 'Planning', '=TODAY()-30', '=TODAY()-20', '=D7-C7', 'Completed', 100, 'None', '■', '■', '', '', '', '', '', '', '', '', '', ''],
      ['Website Redesign', 'Design', '=TODAY()-20', '=TODAY()-5', '=D8-C8', 'Completed', 100, 'Planning', '', '■', '■', '■', '', '', '', '', '', '', '', ''],
      ['Website Redesign', 'Development', '=TODAY()-5', '=TODAY()+20', '=D9-C9', 'In Progress', 30, 'Design', '', '', '', '', '■', '■', '■', '■', '■', '', '', ''],
      ['Process Optimization', 'Assessment', '=TODAY()-60', '=TODAY()-45', '=D10-C10', 'Completed', 100, 'None', '■', '■', '■', '', '', '', '', '', '', '', '', ''],
      ['Process Optimization', 'Optimization', '=TODAY()-45', '=TODAY()+30', '=D11-C11', 'In Progress', 40, 'Assessment', '', '', '■', '■', '■', '■', '■', '■', '■', '■', '', ''],
      ['Data Analytics', 'Infrastructure', '=TODAY()-35', '=TODAY()+10', '=D12-C12', 'At Risk', 75, 'None', '', '■', '■', '■', '■', '■', '■', '', '', '', '', ''],
      ['Data Analytics', 'Implementation', '=TODAY()+10', '=TODAY()+50', '=D13-C13', 'Not Started', 0, 'Infrastructure', '', '', '', '', '', '', '■', '■', '■', '■', '■', '■']
    ];
    
    timelineData.forEach((row, idx) => {
      row.forEach((cell, col) => {
        if (typeof cell === 'string' && cell.startsWith('=')) {
          timeline.setFormula(4 + idx, col + 1, cell);
        } else {
          timeline.setValue(4 + idx, col + 1, cell);
        }
      });
    });
    
    // Data validation
    timeline.addValidation('F4:F1000', ['Not Started', 'In Progress', 'At Risk', 'Completed', 'On Hold']);
    
    // Conditional formatting
    timeline.addConditionalFormat(4, 6, 1000, 6, {
      type: 'cell',
      condition: 'TEXT_EQUALS',
      value: 'Completed',
      format: { background: '#D1FAE5' }
    });
    
    timeline.addConditionalFormat(4, 6, 1000, 6, {
      type: 'cell',
      condition: 'TEXT_EQUALS',
      value: 'At Risk',
      format: { background: '#FEE2E2', fontColor: '#DC2626' }
    });
    
    // Formatting
    timeline.format(4, 3, 13, 4, { numberFormat: 'dd/mm/yyyy' });
    timeline.format(4, 5, 13, 5, { numberFormat: '0" days"' });
    timeline.format(4, 7, 13, 7, { numberFormat: '0%' });
    
    // Column widths
    timeline.setColumnWidth(1, 150);
    timeline.setColumnWidth(2, 120);
    for (let col = 3; col <= 8; col++) {
      timeline.setColumnWidth(col, 90);
    }
    for (let col = 9; col <= 20; col++) {
      timeline.setColumnWidth(col, 30);
    }
  },
  
  buildBudgetSheet: function(budget) {
    // Project budget and financial tracking
    const headers = [
      'Project ID', 'Project Name', 'Budget Category', 'Budgeted Amount',
      'Spent to Date', 'Remaining', 'Invoice Date', 'Revenue', 'Margin',
      'Payment Terms', 'Invoice Status', 'Expected Payment', 'Notes'
    ];
    
    headers.forEach((header, idx) => {
      budget.setValue(1, idx + 1, header);
    });
    
    budget.format(1, 1, 1, headers.length, {
      fontWeight: 'bold',
      background: '#1E40AF',
      fontColor: '#FFFFFF',
      border: true
    });
    
    // Sample budget data
    const budgetData = [
      ['PRJ001', 'Digital Transformation', 'Consulting', 200000, 150000, '=D2-E2', '=TODAY()-15', 250000, '=H2-E2', 'Net 30', 'Paid', '=G2+30', 'First milestone payment received'],
      ['PRJ001', 'Digital Transformation', 'Travel', 15000, 12000, '=D3-E3', '=TODAY()-15', 0, '=H3-E3', 'Net 30', 'Included', '', 'Included in consulting fee'],
      ['PRJ002', 'Website Redesign', 'Development', 70000, 55000, '=D4-E4', '=TODAY()-10', 85000, '=H4-E4', 'Net 15', 'Outstanding', '=G4+15', 'Awaiting client approval'],
      ['PRJ002', 'Website Redesign', 'Design', 15000, 15000, '=D5-E5', '=TODAY()-20', 0, '=H5-E5', 'Net 15', 'Included', '', 'Design phase complete'],
      ['PRJ003', 'Process Optimization', 'Consulting', 150000, 75000, '=D6-E6', '=TODAY()-25', 180000, '=H6-E6', 'Net 30', 'Paid', '=G6+30', 'Progress payment received'],
      ['PRJ003', 'Process Optimization', 'Training', 30000, 8000, '=D7-E7', '', 0, '=H7-E7', 'Net 30', 'Not Invoiced', '', 'Training phase upcoming'],
      ['PRJ004', 'Data Analytics', 'Technology', 250000, 200000, '=D8-E8', '=TODAY()-5', 320000, '=H8-E8', 'Net 30', 'Outstanding', '=G8+30', 'Major technology investment'],
      ['PRJ004', 'Data Analytics', 'Integration', 70000, 45000, '=D9-E9', '', 0, '=H9-E9', 'Net 30', 'Not Invoiced', '', 'Integration work in progress'],
      ['PRJ005', 'Marketing Automation', 'Setup', 35000, 15000, '=D10-E10', '', 45000, '=H10-E10', 'Net 15', 'Not Invoiced', '', 'Small project, quick turnaround'],
      ['PRJ005', 'Marketing Automation', 'Training', 10000, 3000, '=D11-E11', '', 0, '=H11-E11', 'Net 15', 'Not Invoiced', '', 'Training scheduled next month']
    ];
    
    budgetData.forEach((row, idx) => {
      row.forEach((cell, col) => {
        if (typeof cell === 'string' && cell.startsWith('=')) {
          budget.setFormula(2 + idx, col + 1, cell);
        } else {
          budget.setValue(2 + idx, col + 1, cell);
        }
      });
    });
    
    // Data validation
    budget.addValidation('C2:C1000', ['Consulting', 'Development', 'Design', 'Travel', 'Training', 'Technology', 'Integration', 'Setup']);
    budget.addValidation('J2:J1000', ['Net 15', 'Net 30', 'Net 45', 'Due on Receipt']);
    budget.addValidation('K2:K1000', ['Not Invoiced', 'Outstanding', 'Paid', 'Overdue', 'Included']);
    
    // Conditional formatting
    budget.addConditionalFormat(2, 11, 1000, 11, {
      type: 'cell',
      condition: 'TEXT_EQUALS',
      value: 'Paid',
      format: { background: '#D1FAE5' }
    });
    
    budget.addConditionalFormat(2, 11, 1000, 11, {
      type: 'cell',
      condition: 'TEXT_EQUALS',
      value: 'Overdue',
      format: { background: '#FEE2E2', fontColor: '#DC2626' }
    });
    
    // Formatting
    budget.format(2, 4, 11, 6, { numberFormat: '$#,##0' });
    budget.format(2, 7, 11, 9, { numberFormat: 'dd/mm/yyyy' });
    budget.format(2, 8, 11, 9, { numberFormat: '$#,##0' });
    budget.format(2, 12, 11, 12, { numberFormat: 'dd/mm/yyyy' });
    
    // Budget summary
    budget.setValue(13, 1, 'BUDGET SUMMARY');
    budget.merge(13, 1, 13, 8);
    budget.format(13, 1, 13, 8, {
      fontWeight: 'bold',
      background: '#DBEAFE'
    });
    
    const summaryMetrics = [
      ['Total Project Budgets:', '=SUM(D2:D11)'],
      ['Total Spent to Date:', '=SUM(E2:E11)'],
      ['Total Remaining Budget:', '=SUM(F2:F11)'],
      ['Total Revenue (Invoiced):', '=SUM(H2:H11)'],
      ['Total Margin:', '=SUM(I2:I11)'],
      ['Outstanding Invoices:', '=SUMIFS(H2:H11,K2:K11,"Outstanding")'],
      ['Overdue Amount:', '=SUMIFS(H2:H11,K2:K11,"Overdue")'],
      ['Average Margin %:', '=AVERAGE(I2:I11)/AVERAGE(H2:H11)']
    ];
    
    summaryMetrics.forEach((metric, idx) => {
      budget.setValue(14 + idx, 1, metric[0]);
      budget.setFormula(14 + idx, 2, metric[1]);
      budget.merge(14 + idx, 2, 14 + idx, 3);
    });
    
    budget.format(14, 2, 20, 3, { numberFormat: '$#,##0' });
    budget.format(21, 2, 21, 3, { numberFormat: '0%' });
    
    // Column widths
    budget.setColumnWidth(1, 80);
    budget.setColumnWidth(2, 150);
    budget.setColumnWidth(3, 100);
    for (let col = 4; col <= headers.length; col++) {
      budget.setColumnWidth(col, 100);
    }
  },
  
  buildDeliverablesSheet: function(deliverables) {
    // Deliverable tracking and client communication
    const headers = [
      'Deliverable ID', 'Project ID', 'Name', 'Type', 'Owner',
      'Due Date', 'Status', 'Progress %', 'Client Review', 'Approval Status',
      'Delivery Method', 'Version', 'File Location', 'Client Feedback', 'Notes'
    ];
    
    headers.forEach((header, idx) => {
      deliverables.setValue(1, idx + 1, header);
    });
    
    deliverables.format(1, 1, 1, headers.length, {
      fontWeight: 'bold',
      background: '#1E40AF',
      fontColor: '#FFFFFF',
      border: true
    });
    
    // Sample deliverables
    const sampleDeliverables = [
      ['DEL001', 'PRJ001', 'Stakeholder Interview Report', 'Document', 'Sarah Mitchell', '=TODAY()-35', 'Delivered', 100, 'Completed', 'Approved', 'Email', '1.0', '/projects/PRJ001/reports/', 'Excellent insights captured', 'Client very satisfied'],
      ['DEL002', 'PRJ001', 'Current State Analysis', 'Analysis', 'Mike Johnson', '=TODAY()-25', 'Delivered', 100, 'Completed', 'Approved', 'Portal', '2.1', '/projects/PRJ001/analysis/', 'Minor revisions requested', 'Revised based on feedback'],
      ['DEL003', 'PRJ001', 'Future State Blueprint', 'Design', 'Lisa Chen', '=TODAY()+5', 'In Progress', 75, 'Pending', 'Pending', 'Portal', '0.8', '/projects/PRJ001/design/', '', 'On track for delivery'],
      ['DEL004', 'PRJ002', 'Website Mockups', 'Design', 'Jennifer Wu', '=TODAY()-10', 'Under Review', 100, 'In Review', 'Pending', 'Email', '3.2', '/projects/PRJ002/design/', 'Requested color changes', 'Awaiting final feedback'],
      ['DEL005', 'PRJ002', 'Technical Specifications', 'Document', 'David Park', '=TODAY()-5', 'Delivered', 100, 'Completed', 'Approved', 'Portal', '1.0', '/projects/PRJ002/specs/', 'Clear and comprehensive', 'Well received by client'],
      ['DEL006', 'PRJ003', 'Process Maps', 'Diagram', 'Lisa Chen', '=TODAY()-45', 'Delivered', 100, 'Completed', 'Approved', 'Portal', '1.5', '/projects/PRJ003/maps/', 'Very detailed analysis', 'Helped identify key issues'],
      ['DEL007', 'PRJ003', 'Optimization Recommendations', 'Report', 'Mike Johnson', '=TODAY()+10', 'In Progress', 60, 'Pending', 'Pending', 'Portal', '0.6', '/projects/PRJ003/reports/', '', 'Draft in progress'],
      ['DEL008', 'PRJ004', 'Data Architecture Document', 'Technical', 'David Park', '=TODAY()-5', 'Blocked', 80, 'Pending', 'Pending', 'Portal', '0.8', '/projects/PRJ004/architecture/', '', 'Blocked by system access issues'],
      ['DEL009', 'PRJ005', 'Platform Setup Guide', 'Guide', 'Jennifer Wu', '=TODAY()+15', 'Not Started', 0, 'Pending', 'Pending', 'Email', '0.0', '', '', 'Scheduled to start next week']
    ];
    
    sampleDeliverables.forEach((deliverable, idx) => {
      deliverable.forEach((value, col) => {
        if (typeof value === 'string' && value.startsWith('=')) {
          deliverables.setFormula(idx + 2, col + 1, value);
        } else {
          deliverables.setValue(idx + 2, col + 1, value);
        }
      });
    });
    
    // Data validation
    deliverables.addValidation('D2:D1000', ['Document', 'Analysis', 'Design', 'Report', 'Diagram', 'Technical', 'Guide', 'Presentation']);
    deliverables.addValidation('G2:G1000', ['Not Started', 'In Progress', 'Under Review', 'Delivered', 'Blocked']);
    deliverables.addValidation('I2:I1000', ['Pending', 'In Review', 'Completed']);
    deliverables.addValidation('J2:J1000', ['Pending', 'Approved', 'Rejected', 'Needs Revision']);
    deliverables.addValidation('K2:K1000', ['Email', 'Portal', 'Meeting', 'Physical']);
    
    // Conditional formatting
    deliverables.addConditionalFormat(2, 7, 1000, 7, {
      type: 'cell',
      condition: 'TEXT_EQUALS',
      value: 'Delivered',
      format: { background: '#D1FAE5' }
    });
    
    deliverables.addConditionalFormat(2, 7, 1000, 7, {
      type: 'cell',
      condition: 'TEXT_EQUALS',
      value: 'Blocked',
      format: { background: '#FEE2E2', fontColor: '#DC2626' }
    });
    
    deliverables.addConditionalFormat(2, 10, 1000, 10, {
      type: 'cell',
      condition: 'TEXT_EQUALS',
      value: 'Approved',
      format: { background: '#D1FAE5' }
    });
    
    // Formatting
    deliverables.format(2, 6, 1000, 6, { numberFormat: 'dd/mm/yyyy' });
    deliverables.format(2, 8, 1000, 8, { numberFormat: '0%' });
    
    // Column widths
    deliverables.setColumnWidth(1, 100);
    deliverables.setColumnWidth(2, 80);
    deliverables.setColumnWidth(3, 200);
    deliverables.setColumnWidth(13, 200);
    deliverables.setColumnWidth(14, 200);
    deliverables.setColumnWidth(15, 200);
    for (let col = 4; col <= 12; col++) {
      deliverables.setColumnWidth(col, 100);
    }
  },
  
  buildProjectReportsSheet: function(reports) {
    // Executive project reporting
    reports.setValue(1, 1, 'Professional Services Executive Report');
    reports.merge(1, 1, 1, 12);
    reports.format(1, 1, 1, 12, {
      fontSize: 16,
      fontWeight: 'bold',
      horizontalAlignment: 'center',
      background: '#1E40AF',
      fontColor: '#FFFFFF'
    });
    
    // Executive Summary
    reports.setValue(3, 1, 'EXECUTIVE SUMMARY');
    reports.merge(3, 1, 3, 6);
    reports.format(3, 1, 3, 6, {
      fontWeight: 'bold',
      background: '#DBEAFE'
    });
    
    const summaryItems = [
      ['Active Projects:', '=COUNTIF({{Projects}}!F:F,"Active")'],
      ['Total Revenue (YTD):', '=SUM({{Budget}}!H:H)'],
      ['Average Project Margin:', '=AVERAGE({{Projects}}!O:O)'],
      ['Team Utilization:', '=SUM({{Resources}}!E:E)/SUM({{Resources}}!D:D)'],
      ['Client Satisfaction:', '=AVERAGE({{Projects}}!N:N)'],
      ['On-Time Delivery Rate:', '=COUNTIFS({{Projects}}!L:L,"On Track")/COUNTIF({{Projects}}!F:F,"Active")']
    ];
    
    summaryItems.forEach((item, idx) => {
      reports.setValue(4 + idx, 1, item[0]);
      reports.setFormula(4 + idx, 2, item[1]);
      reports.merge(4 + idx, 2, 4 + idx, 3);
    });
    
    reports.format(5, 2, 5, 3, { numberFormat: '$#,##0' });
    reports.format(6, 2, 6, 3, { numberFormat: '0%' });
    reports.format(7, 2, 7, 3, { numberFormat: '0%' });
    reports.format(8, 2, 8, 3, { numberFormat: '0.0' });
    reports.format(9, 2, 9, 3, { numberFormat: '0%' });
    
    // Key Achievements
    reports.setValue(3, 7, 'KEY ACHIEVEMENTS');
    reports.merge(3, 7, 3, 12);
    reports.format(3, 7, 3, 12, {
      fontWeight: 'bold',
      background: '#DBEAFE'
    });
    
    const achievements = [
      '✅ Completed 3 major client engagements',
      '📈 Achieved 92% team utilization rate',
      '🎯 Maintained 4.5/5 average client satisfaction',
      '💰 Exceeded quarterly revenue target by 8%',
      '🚀 Launched new data analytics service line',
      '👥 Added 2 senior consultants to the team'
    ];
    
    achievements.forEach((achievement, idx) => {
      reports.setValue(4 + idx, 7, achievement);
      reports.merge(4 + idx, 7, 4 + idx, 12);
    });
    
    // Project Status Overview
    reports.setValue(11, 1, 'PROJECT STATUS OVERVIEW');
    reports.merge(11, 1, 11, 12);
    reports.format(11, 1, 11, 12, {
      fontWeight: 'bold',
      background: '#DBEAFE'
    });
    
    const projectHeaders = ['Project', 'Client', 'Status', 'Progress', 'Budget Health', 'Due Date', 'Risk Level'];
    projectHeaders.forEach((header, idx) => {
      reports.setValue(12, idx + 1, header);
    });
    reports.format(12, 1, 12, 7, {
      fontWeight: 'bold',
      background: '#E5E7EB'
    });
    
    // Pull active projects
    for (let i = 0; i < 5; i++) {
      reports.setFormula(13 + i, 1, `=INDEX({{Projects}}!B:B,${i + 2})`);
      reports.setFormula(13 + i, 2, `=INDEX({{Projects}}!C:C,${i + 2})`);
      reports.setFormula(13 + i, 3, `=INDEX({{Projects}}!F:F,${i + 2})`);
      reports.setValue(13 + i, 4, Math.floor(Math.random() * 40 + 60) + '%');
      reports.setFormula(13 + i, 5, `=INDEX({{Projects}}!M:M,${i + 2})`);
      reports.setFormula(13 + i, 6, `=INDEX({{Projects}}!H:H,${i + 2})`);
      reports.setFormula(13 + i, 7, `=INDEX({{Projects}}!L:L,${i + 2})`);
    }
    
    reports.format(13, 6, 17, 6, { numberFormat: 'dd/mm/yyyy' });
    
    // Financial Performance
    reports.setValue(19, 1, 'FINANCIAL PERFORMANCE');
    reports.merge(19, 1, 19, 6);
    reports.format(19, 1, 19, 6, {
      fontWeight: 'bold',
      background: '#DBEAFE'
    });
    
    const financialMetrics = [
      ['Revenue This Quarter:', '=SUM({{Budget}}!H:H)*0.33'],
      ['Revenue YTD:', '=SUM({{Budget}}!H:H)'],
      ['Outstanding Invoices:', '=SUMIFS({{Budget}}!H:H,{{Budget}}!K:K,"Outstanding")'],
      ['Average Project Size:', '=AVERAGE({{Projects}}!I:I)'],
      ['Billable Utilization:', '=SUM({{Resources}}!E:E)/SUM({{Resources}}!D:D)'],
      ['Revenue per Employee:', '=SUM({{Budget}}!H:H)/COUNTIF({{Resources}}!J:J,"Full-time")']
    ];
    
    financialMetrics.forEach((metric, idx) => {
      reports.setValue(20 + idx, 1, metric[0]);
      reports.setFormula(20 + idx, 2, metric[1]);
      reports.merge(20 + idx, 2, 20 + idx, 3);
    });
    
    reports.format(20, 2, 24, 3, { numberFormat: '$#,##0' });
    reports.format(24, 2, 25, 3, { numberFormat: '0%' });
    
    // Action Items
    reports.setValue(19, 7, 'ACTION ITEMS');
    reports.merge(19, 7, 19, 12);
    reports.format(19, 7, 19, 12, {
      fontWeight: 'bold',
      background: '#DBEAFE'
    });
    
    const actions = [
      '1. Address resource constraints in Data Analytics project',
      '2. Follow up on outstanding invoices (total: $85k)',
      '3. Schedule client satisfaction surveys for Q4',
      '4. Review and optimize team utilization rates',
      '5. Prepare proposals for 3 new opportunities',
      '6. Plan Q1 capacity and resource allocation'
    ];
    
    actions.forEach((action, idx) => {
      reports.setValue(20 + idx, 7, action);
      reports.merge(20 + idx, 7, 20 + idx, 12);
    });
    
    // Column widths
    reports.setColumnWidth(1, 150);
    for (let col = 2; col <= 12; col++) {
      reports.setColumnWidth(col, 90);
    }
  },
  
  /**
   * Additional template stubs - would be implemented with similar comprehensive patterns
   */
  buildTimeBilling: function(builder) {
    // Time tracking, billing rates, invoice generation
    return builder.complete();
  },
  
  buildClientDashboard: function(builder) {
    // Client-specific project views and reporting
    return builder.complete();
  },
  
  buildResourcePlanner: function(builder) {
    // Resource capacity planning and allocation
    return builder.complete();
  },
  
  buildProposalTracker: function(builder) {
    // Proposal pipeline and win/loss tracking
    return builder.complete();
  },
  
  buildInvoiceManager: function(builder) {
    // Invoice generation and accounts receivable
    return builder.complete();
  },
  
  buildCapacityPlanning: function(builder) {
    // Team capacity and workload planning
    return builder.complete();
  },
  
  buildFinancialDashboard: function(builder) {
    // Financial metrics and P&L analysis
    return builder.complete();
  },
  
  buildTeamPerformance: function(builder) {
    // Team productivity and performance metrics
    return builder.complete();
  },
  
  buildClientSatisfaction: function(builder) {
    // Client satisfaction tracking and surveys
    return builder.complete();
  }
};