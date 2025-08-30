/**
 * Healthcare Template Module
 * Provides specialized templates for healthcare and medical practices
 */

var HealthcareTemplate = {
  /**
   * Main builder function
   */
  build: function(spreadsheet, templateType, isPreview = false) {
    const builder = new TemplateBuilderPro(spreadsheet, isPreview);
    
    try {
      switch (templateType) {
        case 'patient-tracker':
          this.buildPatientTracker(builder);
          break;
        case 'appointment-scheduler':
          this.buildAppointmentScheduler(builder);
          break;
        case 'medication-manager':
          this.buildMedicationManager(builder);
          break;
        case 'billing-tracker':
          this.buildBillingTracker(builder);
          break;
        default:
          builder.errors.push(`Unknown template type: ${templateType}`);
      }
    } catch (error) {
      builder.errors.push(`Build error: ${error.toString()}`);
    }
    
    return builder.complete();
  },
  
  /**
   * Patient Tracker Template
   */
  buildPatientTracker: function(builder) {
    const sheets = builder.createSheets([
      'Dashboard',
      'Patients',
      'Visits',
      'Conditions',
      'Lab Results'
    ]);
    
    // Dashboard Setup
    const dashboard = sheets.Dashboard;
    
    // Title
    dashboard.merge(1, 1, 1, 12);
    dashboard.setValue(1, 1, 'Patient Management System');
    dashboard.format(1, 1, 1, 12, {
      fontSize: 24,
      fontWeight: 'bold',
      horizontalAlignment: 'center',
      background: '#0891B2'
    });
    
    // KPIs Section
    dashboard.setValue(3, 1, 'KEY METRICS');
    dashboard.format(3, 1, 3, 12, {
      fontSize: 14,
      fontWeight: 'bold',
      background: '#E0F2FE'
    });
    
    // KPI Cards
    const kpis = [
      { label: 'Total Patients', formula: '=COUNTA({{Patients}}!A:A)-1' },
      { label: 'Active Cases', formula: '=COUNTIF({{Patients}}!E:E,"Active")' },
      { label: 'Appointments Today', formula: '=COUNTIF({{Visits}}!C:C,TODAY())' },
      { label: 'Critical Conditions', formula: '=COUNTIF({{Conditions}}!D:D,"Critical")' }
    ];
    
    let col = 1;
    kpis.forEach(kpi => {
      dashboard.merge(5, col, 5, col + 2);
      dashboard.setValue(5, col, kpi.label);
      dashboard.merge(6, col, 7, col + 2);
      dashboard.setFormula(6, col, kpi.formula);
      dashboard.format(5, col, 7, col + 2, {
        border: true,
        horizontalAlignment: 'center'
      });
      col += 3;
    });
    
    // Patients Sheet
    const patients = sheets.Patients;
    const patientHeaders = [
      'Patient ID', 'Name', 'DOB', 'Gender', 'Status', 
      'Primary Condition', 'Phone', 'Email', 'Emergency Contact'
    ];
    
    patients.setRangeValues(1, 1, [patientHeaders]);
    patients.format(1, 1, 1, patientHeaders.length, {
      fontWeight: 'bold',
      background: '#E0F2FE',
      border: true
    });
    
    // Add sample data
    const samplePatients = [
      ['P001', 'John Doe', '1985-03-15', 'Male', 'Active', 'Hypertension', '555-0101', 'john@email.com', 'Jane Doe'],
      ['P002', 'Mary Smith', '1990-07-22', 'Female', 'Active', 'Diabetes', '555-0102', 'mary@email.com', 'Bob Smith'],
      ['P003', 'Robert Johnson', '1975-11-08', 'Male', 'Inactive', 'Asthma', '555-0103', 'robert@email.com', 'Lisa Johnson']
    ];
    
    patients.setRangeValues(2, 1, samplePatients);
    
    // Data validation for Status
    patients.addValidation('E2:E100', ['Active', 'Inactive', 'Discharged']);
    
    // Conditional formatting for Status
    patients.addConditionalFormat(2, 5, 100, 5, {
      type: 'cell',
      condition: 'TEXT_EQUALS',
      value: 'Active',
      format: { background: '#D1FAE5' }
    });
    
    // Visits Sheet
    const visits = sheets.Visits;
    const visitHeaders = [
      'Visit ID', 'Patient ID', 'Date', 'Type', 'Provider',
      'Reason', 'Diagnosis', 'Follow-up Required', 'Notes'
    ];
    
    visits.setRangeValues(1, 1, [visitHeaders]);
    visits.format(1, 1, 1, visitHeaders.length, {
      fontWeight: 'bold',
      background: '#E0F2FE',
      border: true
    });
    
    // Conditions Sheet
    const conditions = sheets.Conditions;
    const conditionHeaders = [
      'Condition ID', 'Patient ID', 'Condition', 'Severity',
      'Diagnosed Date', 'Treatment Plan', 'Status'
    ];
    
    conditions.setRangeValues(1, 1, [conditionHeaders]);
    conditions.format(1, 1, 1, conditionHeaders.length, {
      fontWeight: 'bold',
      background: '#E0F2FE',
      border: true
    });
    
    // Add severity validation
    conditions.addValidation('D2:D100', ['Mild', 'Moderate', 'Severe', 'Critical']);
    
    // Lab Results Sheet
    const labs = sheets['Lab Results'];
    const labHeaders = [
      'Test ID', 'Patient ID', 'Test Date', 'Test Type',
      'Result', 'Normal Range', 'Flag', 'Ordered By'
    ];
    
    labs.setRangeValues(1, 1, [labHeaders]);
    labs.format(1, 1, 1, labHeaders.length, {
      fontWeight: 'bold',
      background: '#E0F2FE',
      border: true
    });
  },
  
  /**
   * Appointment Scheduler Template
   */
  buildAppointmentScheduler: function(builder) {
    const sheets = builder.createSheets([
      'Schedule Dashboard',
      'Appointments',
      'Providers',
      'Patients',
      'Availability'
    ]);
    
    // Schedule Dashboard
    const dashboard = sheets['Schedule Dashboard'];
    
    dashboard.merge(1, 1, 1, 10);
    dashboard.setValue(1, 1, 'Appointment Scheduling System');
    dashboard.format(1, 1, 1, 10, {
      fontSize: 20,
      fontWeight: 'bold',
      horizontalAlignment: 'center',
      background: '#06B6D4'
    });
    
    // Today's Schedule Section
    dashboard.setValue(3, 1, "TODAY'S APPOINTMENTS");
    dashboard.format(3, 1, 3, 10, {
      fontSize: 14,
      fontWeight: 'bold',
      background: '#E0F2FE'
    });
    
    // Appointments Sheet
    const appointments = sheets.Appointments;
    const appointmentHeaders = [
      'Appointment ID', 'Patient Name', 'Provider', 'Date', 'Time',
      'Duration (min)', 'Type', 'Status', 'Room', 'Notes'
    ];
    
    appointments.setRangeValues(1, 1, [appointmentHeaders]);
    appointments.format(1, 1, 1, appointmentHeaders.length, {
      fontWeight: 'bold',
      background: '#E0F2FE',
      border: true
    });
    
    // Add validation for appointment types
    appointments.addValidation('G2:G500', [
      'Consultation', 'Follow-up', 'Physical Exam', 
      'Lab Work', 'Procedure', 'Emergency'
    ]);
    
    // Add validation for status
    appointments.addValidation('H2:H500', [
      'Scheduled', 'Confirmed', 'In Progress', 
      'Completed', 'Cancelled', 'No Show'
    ]);
    
    // Providers Sheet
    const providers = sheets.Providers;
    const providerHeaders = [
      'Provider ID', 'Name', 'Specialty', 'License #',
      'Phone', 'Email', 'Available Days', 'Hours'
    ];
    
    providers.setRangeValues(1, 1, [providerHeaders]);
    providers.format(1, 1, 1, providerHeaders.length, {
      fontWeight: 'bold',
      background: '#E0F2FE',
      border: true
    });
    
    // Availability Sheet
    const availability = sheets.Availability;
    const availHeaders = [
      'Provider', 'Day', 'Start Time', 'End Time',
      'Lunch Start', 'Lunch End', 'Appointment Slots'
    ];
    
    availability.setRangeValues(1, 1, [availHeaders]);
    availability.format(1, 1, 1, availHeaders.length, {
      fontWeight: 'bold',
      background: '#E0F2FE',
      border: true
    });
  },
  
  /**
   * Medication Manager Template
   */
  buildMedicationManager: function(builder) {
    const sheets = builder.createSheets([
      'Medication Dashboard',
      'Prescriptions',
      'Inventory',
      'Patients',
      'Refill Requests'
    ]);
    
    // Dashboard
    const dashboard = sheets['Medication Dashboard'];
    
    dashboard.merge(1, 1, 1, 12);
    dashboard.setValue(1, 1, 'Medication Management System');
    dashboard.format(1, 1, 1, 12, {
      fontSize: 22,
      fontWeight: 'bold',
      horizontalAlignment: 'center',
      background: '#0891B2',
      fontColor: '#FFFFFF'
    });
    
    // Alert Section
    dashboard.setValue(3, 1, 'ALERTS & NOTIFICATIONS');
    dashboard.format(3, 1, 3, 12, {
      fontSize: 14,
      fontWeight: 'bold',
      background: '#FEE2E2'
    });
    
    // Alerts
    const alerts = [
      { label: 'Expiring Soon', formula: '=COUNTIF({{Inventory}}!F:F,"<30")' },
      { label: 'Low Stock', formula: '=COUNTIF({{Inventory}}!E:E,"<100")' },
      { label: 'Pending Refills', formula: '=COUNTIF({{Refill Requests}}!F:F,"Pending")' },
      { label: 'Drug Interactions', formula: '=COUNTIF({{Prescriptions}}!H:H,"Yes")' }
    ];
    
    let row = 5;
    alerts.forEach(alert => {
      dashboard.setValue(row, 1, alert.label);
      dashboard.setFormula(row, 3, alert.formula);
      dashboard.format(row, 1, row, 4, {
        border: true
      });
      row++;
    });
    
    // Prescriptions Sheet
    const prescriptions = sheets.Prescriptions;
    const rxHeaders = [
      'Rx Number', 'Patient ID', 'Patient Name', 'Medication',
      'Dosage', 'Frequency', 'Start Date', 'End Date',
      'Prescriber', 'Refills Remaining', 'Interaction Check'
    ];
    
    prescriptions.setRangeValues(1, 1, [rxHeaders]);
    prescriptions.format(1, 1, 1, rxHeaders.length, {
      fontWeight: 'bold',
      background: '#E0F2FE',
      border: true
    });
    
    // Inventory Sheet
    const inventory = sheets.Inventory;
    const invHeaders = [
      'Medication ID', 'Name', 'Generic Name', 'Category',
      'Current Stock', 'Days Supply', 'Reorder Point',
      'Unit Cost', 'Supplier', 'Expiration Date'
    ];
    
    inventory.setRangeValues(1, 1, [invHeaders]);
    inventory.format(1, 1, 1, invHeaders.length, {
      fontWeight: 'bold',
      background: '#E0F2FE',
      border: true
    });
    
    // Conditional formatting for low stock
    inventory.addConditionalFormat(2, 5, 500, 5, {
      type: 'cell',
      condition: 'NUMBER_LESS_THAN',
      value: 100,
      format: { background: '#FEE2E2' }
    });
    
    // Refill Requests Sheet
    const refills = sheets['Refill Requests'];
    const refillHeaders = [
      'Request ID', 'Date', 'Patient Name', 'Medication',
      'Rx Number', 'Status', 'Approved By', 'Filled Date'
    ];
    
    refills.setRangeValues(1, 1, [refillHeaders]);
    refills.format(1, 1, 1, refillHeaders.length, {
      fontWeight: 'bold',
      background: '#E0F2FE',
      border: true
    });
    
    // Add status validation
    refills.addValidation('F2:F500', [
      'Pending', 'Approved', 'Denied', 'Filled', 'Ready for Pickup'
    ]);
  },
  
  /**
   * Billing Tracker Template
   */
  buildBillingTracker: function(builder) {
    const sheets = builder.createSheets([
      'Billing Dashboard',
      'Claims',
      'Payments',
      'Denials',
      'Patient Balances'
    ]);
    
    // Dashboard
    const dashboard = sheets['Billing Dashboard'];
    
    dashboard.merge(1, 1, 1, 14);
    dashboard.setValue(1, 1, 'Medical Billing & Revenue Cycle');
    dashboard.format(1, 1, 1, 14, {
      fontSize: 24,
      fontWeight: 'bold',
      horizontalAlignment: 'center',
      background: '#059669',
      fontColor: '#FFFFFF'
    });
    
    // Revenue Metrics
    dashboard.setValue(3, 1, 'REVENUE METRICS');
    dashboard.format(3, 1, 3, 14, {
      fontSize: 14,
      fontWeight: 'bold',
      background: '#D1FAE5'
    });
    
    // Metric Cards
    const metrics = [
      { label: 'Total Billed', formula: '=SUM({{Claims}}!E:E)' },
      { label: 'Total Collected', formula: '=SUM({{Payments}}!D:D)' },
      { label: 'Outstanding', formula: '=SUM({{Patient Balances}}!E:E)' },
      { label: 'Collection Rate', formula: '=SUM({{Payments}}!D:D)/SUM({{Claims}}!E:E)' }
    ];
    
    let col = 1;
    metrics.forEach(metric => {
      dashboard.merge(5, col, 5, col + 2);
      dashboard.setValue(5, col, metric.label);
      dashboard.merge(6, col, 7, col + 2);
      dashboard.setFormula(6, col, metric.formula);
      dashboard.format(5, col, 7, col + 2, {
        border: true,
        horizontalAlignment: 'center'
      });
      
      // Format percentage for Collection Rate
      if (metric.label === 'Collection Rate') {
        dashboard.format(6, col, 7, col + 2, {
          numberFormat: '0.0%'
        });
      } else {
        dashboard.format(6, col, 7, col + 2, {
          numberFormat: '$#,##0'
        });
      }
      
      col += 3;
    });
    
    // Claims Sheet
    const claims = sheets.Claims;
    const claimHeaders = [
      'Claim ID', 'Patient Name', 'Service Date', 'CPT Code',
      'Billed Amount', 'Insurance', 'Status', 'Submitted Date',
      'Expected Payment', 'Days Outstanding'
    ];
    
    claims.setRangeValues(1, 1, [claimHeaders]);
    claims.format(1, 1, 1, claimHeaders.length, {
      fontWeight: 'bold',
      background: '#D1FAE5',
      border: true
    });
    
    // Status validation
    claims.addValidation('G2:G1000', [
      'Submitted', 'In Review', 'Approved', 'Partially Paid',
      'Paid', 'Denied', 'Appealed'
    ]);
    
    // Payments Sheet
    const payments = sheets.Payments;
    const paymentHeaders = [
      'Payment ID', 'Claim ID', 'Date Received', 'Amount',
      'Payment Method', 'Insurance/Patient', 'Check Number'
    ];
    
    payments.setRangeValues(1, 1, [paymentHeaders]);
    payments.format(1, 1, 1, paymentHeaders.length, {
      fontWeight: 'bold',
      background: '#D1FAE5',
      border: true
    });
    
    // Denials Sheet
    const denials = sheets.Denials;
    const denialHeaders = [
      'Denial ID', 'Claim ID', 'Denial Date', 'Reason Code',
      'Reason Description', 'Amount', 'Appeal Status', 'Resolution'
    ];
    
    denials.setRangeValues(1, 1, [denialHeaders]);
    denials.format(1, 1, 1, denialHeaders.length, {
      fontWeight: 'bold',
      background: '#FEE2E2',
      border: true
    });
    
    // Patient Balances Sheet
    const balances = sheets['Patient Balances'];
    const balanceHeaders = [
      'Patient ID', 'Patient Name', 'Total Charges', 'Insurance Paid',
      'Patient Balance', 'Last Payment Date', 'Days Overdue', 'Status'
    ];
    
    balances.setRangeValues(1, 1, [balanceHeaders]);
    balances.format(1, 1, 1, balanceHeaders.length, {
      fontWeight: 'bold',
      background: '#D1FAE5',
      border: true
    });
    
    // Conditional formatting for overdue accounts
    balances.addConditionalFormat(2, 7, 1000, 7, {
      type: 'gradient',
      minColor: '#FFFFFF',
      midColor: '#FEF3C7',
      maxColor: '#FEE2E2',
      minValue: '0',
      midValue: '30',
      maxValue: '90'
    });
  },
  
  /**
   * Helper function to create consistent headers
   */
  createHeaderRow: function(sheet, headers, row = 1) {
    sheet.setRangeValues(row, 1, [headers]);
    sheet.format(row, 1, row, headers.length, {
      background: '#E0F2FE',
      fontWeight: 'bold',
      fontSize: 11,
      border: true
    });
  },
  
  /**
   * Helper function to create dashboard title
   */
  createDashboardTitle: function(sheet, title, subtitle = '') {
    sheet.merge(1, 1, 1, 12);
    sheet.setValue(1, 1, title);
    sheet.format(1, 1, 1, 12, {
      background: '#0891B2',
      fontColor: '#FFFFFF',
      fontSize: 20,
      fontWeight: 'bold',
      horizontalAlignment: 'center'
    });
    
    if (subtitle) {
      sheet.merge(2, 1, 2, 12);
      sheet.setValue(2, 1, subtitle);
      sheet.format(2, 1, 2, 12, {
        background: '#E0F2FE',
        fontSize: 12,
        horizontalAlignment: 'center'
      });
    }
  },
  
  /**
   * Create KPI card
   */
  createKPICard: function(sheet, row, col, label, formula, format = {}) {
    sheet.merge(row, col, row, col + 2);
    sheet.setValue(row, col, label);
    sheet.format(row, col, row, col + 2, {
      background: '#F3F4F6',
      fontSize: 10,
      horizontalAlignment: 'center'
    });
    
    sheet.merge(row + 1, col, row + 2, col + 2);
    sheet.setFormula(row + 1, col, formula);
    sheet.format(row + 1, col, row + 2, col + 2, Object.assign({
      fontSize: 16,
      fontWeight: 'bold'
    }, format));
  }
};

// TemplateBuilderPro and SheetHelper classes are now in TemplateBuilder.js