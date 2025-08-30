/**
 * Professional E-commerce Templates Suite
 * Complete templates for online retailers and e-commerce businesses
 */

var EcommerceTemplate = {
  /**
   * Build comprehensive E-commerce template based on type
   */
  build: function(spreadsheet, templateType = 'sales-dashboard', isPreview = false) {
    const builder = new TemplateBuilderPro(spreadsheet, isPreview);
    
    switch(templateType) {
      case 'sales-dashboard':
        return this.buildSalesDashboard(builder);
      case 'inventory-manager':
        return this.buildInventoryManager(builder);
      case 'customer-analytics':
        return this.buildCustomerAnalytics(builder);
      case 'product-performance':
        return this.buildProductPerformance(builder);
      case 'financial-tracker':
        return this.buildFinancialTracker(builder);
      case 'marketing-roi':
        return this.buildMarketingROI(builder);
      case 'supplier-management':
        return this.buildSupplierManagement(builder);
      case 'shipping-logistics':
        return this.buildShippingLogistics(builder);
      case 'price-optimizer':
        return this.buildPriceOptimizer(builder);
      case 'forecasting':
        return this.buildForecasting(builder);
      default:
        return this.buildSalesDashboard(builder);
    }
  },
  
  /**
   * Sales Dashboard - Comprehensive sales tracking and analytics
   */
  buildSalesDashboard: function(builder) {
    const sheets = builder.createSheets([
      'Dashboard',
      'Sales Data',
      'Products',
      'Customers',
      'Orders',
      'Channels',
      'Trends',
      'Reports'
    ]);
    
    this.buildEcommerceDashboard(sheets.Dashboard);
    this.buildSalesDataSheet(sheets['Sales Data']);
    this.buildProductsSheet(sheets.Products);
    this.buildCustomersSheet(sheets.Customers);
    this.buildOrdersSheet(sheets.Orders);
    this.buildChannelsSheet(sheets.Channels);
    this.buildTrendsSheet(sheets.Trends);
    this.buildSalesReportsSheet(sheets.Reports);
    
    return builder.complete();
  },
  
  buildEcommerceDashboard: function(dash) {
    // Title
    dash.setValue(1, 1, 'E-commerce Sales Command Center');
    dash.merge(1, 1, 1, 12);
    dash.format(1, 1, 1, 12, {
      fontSize: 18,
      fontWeight: 'bold',
      horizontalAlignment: 'center',
      background: '#DC2626',
      fontColor: '#FFFFFF'
    });
    
    // Date Range Selector
    dash.setValue(2, 1, 'Period:');
    dash.setValue(2, 2, 'Last 30 Days');
    dash.addValidation('B2', ['Today', 'Last 7 Days', 'Last 30 Days', 'Last Quarter', 'YTD', 'All Time']);
    dash.setValue(2, 5, 'Store:');
    dash.setValue(2, 6, 'All Stores');
    dash.addValidation('F2', ['All Stores', 'Online', 'Amazon', 'eBay', 'Physical Store']);
    
    // Revenue KPIs Row
    dash.setValue(4, 1, 'REVENUE METRICS');
    dash.merge(4, 1, 4, 12);
    dash.format(4, 1, 4, 12, {
      fontSize: 12,
      fontWeight: 'bold',
      background: '#FEE2E2'
    });
    
    const revenueKpis = [
      { row: 5, col: 1, title: 'Total Revenue', formula: '=SUM({{Sales Data}}!F:F)', sub: '=TEXT((A6-500000)/500000,"+0.0%;-0.0%")&" vs target"' },
      { row: 5, col: 4, title: 'Orders', formula: '=COUNTIF({{Orders}}!B:B,"<>")', sub: '=TEXT(D6,"#,##0")&" orders"' },
      { row: 5, col: 7, title: 'AOV', formula: '=A6/D6', sub: '=IF(G6>75,"🟢 Above Average","🟡 Below Average")' },
      { row: 5, col: 10, title: 'Conversion Rate', formula: '=D6/SUMIF({{Channels}}!A:A,"Website",{{Channels}}!C:C)', sub: '=TEXT(J6,"0.0%")&" conversion"' }
    ];
    
    revenueKpis.forEach(kpi => {
      dash.setValue(kpi.row, kpi.col, kpi.title);
      dash.merge(kpi.row, kpi.col, kpi.row, kpi.col + 2);
      dash.setFormula(kpi.row + 1, kpi.col, kpi.formula);
      dash.merge(kpi.row + 1, kpi.col, kpi.row + 1, kpi.col + 2);
      dash.setFormula(kpi.row + 2, kpi.col, kpi.sub);
      dash.merge(kpi.row + 2, kpi.col, kpi.row + 2, kpi.col + 2);
      
      dash.format(kpi.row, kpi.col, kpi.row, kpi.col + 2, {
        background: '#FECACA',
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
    dash.format(6, 1, 6, 3, { numberFormat: '$#,##0' });
    dash.format(6, 7, 6, 9, { numberFormat: '$#,##0.00' });
    dash.format(6, 10, 6, 12, { numberFormat: '0.0%' });
    
    // Top Products
    dash.setValue(9, 1, 'TOP PERFORMING PRODUCTS');
    dash.merge(9, 1, 9, 6);
    dash.format(9, 1, 9, 6, {
      fontSize: 12,
      fontWeight: 'bold',
      background: '#FEE2E2'
    });
    
    const productHeaders = ['Product', 'Sales', 'Units', 'Revenue', 'Margin', 'Rank'];
    productHeaders.forEach((header, idx) => {
      dash.setValue(10, idx + 1, header);
    });
    dash.format(10, 1, 10, 6, {
      fontWeight: 'bold',
      background: '#E5E7EB'
    });
    
    // Top 5 products with dynamic ranking
    for (let i = 0; i < 5; i++) {
      dash.setFormula(11 + i, 1, `=INDEX({{Products}}!B:B,MATCH(LARGE({{Products}}!G:G,${i+1}),{{Products}}!G:G,0))`);
      dash.setFormula(11 + i, 2, `=LARGE({{Products}}!E:E,${i+1})`);
      dash.setFormula(11 + i, 3, `=LARGE({{Products}}!F:F,${i+1})`);
      dash.setFormula(11 + i, 4, `=LARGE({{Products}}!G:G,${i+1})`);
      dash.setFormula(11 + i, 5, `=INDEX({{Products}}!H:H,MATCH(LARGE({{Products}}!G:G,${i+1}),{{Products}}!G:G,0))`);
      dash.setValue(11 + i, 6, `#${i+1}`);
    }
    
    dash.format(11, 2, 15, 2, { numberFormat: '#,##0' });
    dash.format(11, 3, 15, 4, { numberFormat: '#,##0' });
    dash.format(11, 5, 15, 5, { numberFormat: '0.0%' });
    
    // Sales by Channel
    dash.setValue(9, 7, 'SALES BY CHANNEL');
    dash.merge(9, 7, 9, 12);
    dash.format(9, 7, 9, 12, {
      fontSize: 12,
      fontWeight: 'bold',
      background: '#FEE2E2'
    });
    
    const channels = [
      ['Channel', 'Revenue', '% of Total', 'Orders', 'AOV', 'Growth'],
      ['Website', '=SUMIF({{Orders}}!F:F,"Website",{{Orders}}!E:E)', '=H11/SUM(H:H)', '=COUNTIF({{Orders}}!F:F,"Website")', '=H11/I11', '+15%'],
      ['Amazon', '=SUMIF({{Orders}}!F:F,"Amazon",{{Orders}}!E:E)', '=H12/SUM(H:H)', '=COUNTIF({{Orders}}!F:F,"Amazon")', '=H12/I12', '+8%'],
      ['eBay', '=SUMIF({{Orders}}!F:F,"eBay",{{Orders}}!E:E)', '=H13/SUM(H:H)', '=COUNTIF({{Orders}}!F:F,"eBay")', '=H13/I13', '+3%'],
      ['Social', '=SUMIF({{Orders}}!F:F,"Social",{{Orders}}!E:E)', '=H14/SUM(H:H)', '=COUNTIF({{Orders}}!F:F,"Social")', '=H14/I14', '+25%'],
      ['Retail', '=SUMIF({{Orders}}!F:F,"Retail",{{Orders}}!E:E)', '=H15/SUM(H:H)', '=COUNTIF({{Orders}}!F:F,"Retail")', '=H15/I15', '-5%']
    ];
    
    channels.forEach((row, idx) => {
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
    dash.format(11, 8, 15, 8, { numberFormat: '$#,##0' });
    dash.format(11, 9, 15, 9, { numberFormat: '0.0%' });
    dash.format(11, 11, 15, 11, { numberFormat: '$#,##0' });
    
    // Customer Metrics
    dash.setValue(17, 1, 'CUSTOMER INSIGHTS');
    dash.merge(17, 1, 17, 6);
    dash.format(17, 1, 17, 6, {
      fontSize: 12,
      fontWeight: 'bold',
      background: '#FEE2E2'
    });
    
    const customerMetrics = [
      ['Metric', 'Value', 'Target', 'Status', 'Trend', 'Action'],
      ['New Customers', '=COUNTIFS({{Customers}}!D:D,">="&TODAY()-30)', '500', '=IF(B19>=C19,"✓","⚠")', '+12%', 'Good'],
      ['Returning Rate', '=COUNTIFS({{Orders}}!H:H,"Returning")/COUNTIF({{Orders}}!B:B,"<>")', '35%', '=IF(B20>=C20,"✓","⚠")', '+2%', 'Improve'],
      ['Customer LTV', '=AVERAGE({{Customers}}!F:F)', '$250', '=IF(B21>=C21,"✓","⚠")', '+5%', 'Good'],
      ['Churn Rate', '=COUNTIFS({{Customers}}!G:G,"<"&TODAY()-90)/COUNTIF({{Customers}}!B:B,"<>")', '15%', '=IF(B22<=C22,"✓","⚠")', '-3%', 'Watch'],
      ['Avg Order Freq', '=COUNTIF({{Orders}}!B:B,"<>")/COUNTIF({{Customers}}!B:B,"<>")', '3.5', '=IF(B23>=C23,"✓","⚠")', '+8%', 'Excellent']
    ];
    
    customerMetrics.forEach((row, idx) => {
      row.forEach((cell, col) => {
        if (typeof cell === 'string' && cell.startsWith('=')) {
          dash.setFormula(18 + idx, col + 1, cell);
        } else {
          dash.setValue(18 + idx, col + 1, cell);
        }
      });
    });
    
    dash.format(18, 1, 18, 6, {
      fontWeight: 'bold',
      background: '#E5E7EB'
    });
    dash.format(19, 2, 23, 2, { numberFormat: '#,##0' });
    dash.format(20, 2, 20, 3, { numberFormat: '0.0%' });
    dash.format(21, 2, 21, 3, { numberFormat: '$#,##0' });
    
    // Inventory Alerts
    dash.setValue(17, 7, 'INVENTORY ALERTS');
    dash.merge(17, 7, 17, 12);
    dash.format(17, 7, 17, 12, {
      fontSize: 12,
      fontWeight: 'bold',
      background: '#FEE2E2'
    });
    
    const alerts = [
      ['Alert Type', 'Count', 'Priority', 'Action Needed'],
      ['Low Stock', '=COUNTIF({{Products}}!J:J,"<10")', 'High', 'Reorder ASAP'],
      ['Out of Stock', '=COUNTIF({{Products}}!J:J,"0")', 'Critical', 'Immediate Action'],
      ['Slow Moving', '=COUNTIFS({{Products}}!E:E,"<5",{{Products}}!J:J,">50")', 'Medium', 'Review Pricing'],
      ['Overstock', '=COUNTIF({{Products}}!J:J,">100")', 'Low', 'Promote/Discount'],
      ['Expired Items', '=COUNTIFS({{Products}}!K:K,"<"&TODAY())', 'High', 'Remove/Discount']
    ];
    
    alerts.forEach((row, idx) => {
      row.forEach((cell, col) => {
        if (typeof cell === 'string' && cell.startsWith('=')) {
          dash.setFormula(18 + idx, col + 7, cell);
        } else {
          dash.setValue(18 + idx, col + 7, cell);
        }
      });
    });
    
    dash.format(18, 7, 18, 10, {
      fontWeight: 'bold',
      background: '#E5E7EB'
    });
    
    // Sales Trend Chart Area
    dash.setValue(25, 1, 'SALES TREND (Last 30 Days)');
    dash.merge(25, 1, 25, 12);
    dash.format(25, 1, 25, 12, {
      fontSize: 12,
      fontWeight: 'bold',
      background: '#FEE2E2'
    });
    
    // Weekly sales trend
    const trendHeaders = ['Week', 'Revenue', 'Orders', 'AOV', 'Growth %', 'Target', 'Performance'];
    trendHeaders.forEach((header, idx) => {
      dash.setValue(26, idx + 1, header);
    });
    dash.format(26, 1, 26, 7, {
      fontWeight: 'bold',
      background: '#E5E7EB'
    });
    
    // Last 4 weeks data
    for (let week = 0; week < 4; week++) {
      dash.setValue(27 + week, 1, `Week ${4-week}`);
      dash.setFormula(27 + week, 2, `=SUMIFS({{Sales Data}}!F:F,{{Sales Data}}!A:A,">="&TODAY()-${(week+1)*7},{{Sales Data}}!A:A,"<"&TODAY()-${week*7})`);
      dash.setFormula(27 + week, 3, `=COUNTIFS({{Orders}}!C:C,">="&TODAY()-${(week+1)*7},{{Orders}}!C:C,"<"&TODAY()-${week*7})`);
      dash.setFormula(27 + week, 4, `=B${27+week}/C${27+week}`);
      dash.setFormula(27 + week, 5, `=IF(ROW()=27,0,(B${27+week}-B${28+week})/B${28+week})`);
      dash.setValue(27 + week, 6, 125000 - week * 5000);
      dash.setFormula(27 + week, 7, `=IF(B${27+week}>F${27+week},"Above Target","Below Target")`);
    }
    
    dash.format(27, 2, 30, 2, { numberFormat: '$#,##0' });
    dash.format(27, 4, 30, 4, { numberFormat: '$#,##0' });
    dash.format(27, 5, 30, 5, { numberFormat: '+0.0%;-0.0%' });
    dash.format(27, 6, 30, 6, { numberFormat: '$#,##0' });
    
    // Column widths
    dash.setColumnWidth(1, 100);
    for (let col = 2; col <= 12; col++) {
      dash.setColumnWidth(col, 85);
    }
  },
  
  buildSalesDataSheet: function(salesData) {
    // Transaction-level sales data
    const headers = [
      'Date', 'Transaction ID', 'Customer ID', 'Product SKU', 'Product Name',
      'Revenue', 'Quantity', 'Unit Price', 'Cost', 'Profit', 'Margin %',
      'Channel', 'Payment Method', 'Shipping', 'Tax', 'Discount', 'Total'
    ];
    
    headers.forEach((header, idx) => {
      salesData.setValue(1, idx + 1, header);
    });
    
    salesData.format(1, 1, 1, headers.length, {
      fontWeight: 'bold',
      background: '#DC2626',
      fontColor: '#FFFFFF',
      border: true
    });
    
    // Sample sales transactions
    const sampleTransactions = [
      ['=TODAY()-RANDBETWEEN(1,30)', 'TXN001', 'CUST001', 'SKU001', 'Wireless Headphones', 199.99, 2, 99.99, 45.00, '=(F2*G2)-(H2*G2)', '=J2/(F2*G2)', 'Website', 'Credit Card', 9.99, '=F2*G2*0.08', 0, '=(F2*G2)+N2+O2-P2'],
      ['=TODAY()-RANDBETWEEN(1,30)', 'TXN002', 'CUST002', 'SKU002', 'Smart Watch', 299.99, 1, 299.99, 125.00, '=(F3*G3)-(H3*G3)', '=J3/(F3*G3)', 'Amazon', 'PayPal', 0, '=F3*G3*0.08', 20, '=(F3*G3)+N3+O3-P3'],
      ['=TODAY()-RANDBETWEEN(1,30)', 'TXN003', 'CUST003', 'SKU003', 'Bluetooth Speaker', 89.99, 1, 89.99, 32.00, '=(F4*G4)-(H4*G4)', '=J4/(F4*G4)', 'eBay', 'Credit Card', 7.99, '=F4*G4*0.08', 10, '=(F4*G4)+N4+O4-P4'],
      ['=TODAY()-RANDBETWEEN(1,30)', 'TXN004', 'CUST001', 'SKU004', 'Phone Case', 24.99, 3, 24.99, 8.50, '=(F5*G5)-(H5*G5)', '=J5/(F5*G5)', 'Website', 'Credit Card', 4.99, '=F5*G5*0.08', 0, '=(F5*G5)+N5+O5-P5'],
      ['=TODAY()-RANDBETWEEN(1,30)', 'TXN005', 'CUST004', 'SKU005', 'Laptop Stand', 49.99, 1, 49.99, 18.00, '=(F6*G6)-(H6*G6)', '=J6/(F6*G6)', 'Social', 'PayPal', 8.99, '=F6*G6*0.08', 5, '=(F6*G6)+N6+O6-P6']
    ];
    
    sampleTransactions.forEach((transaction, idx) => {
      transaction.forEach((value, col) => {
        if (typeof value === 'string' && value.startsWith('=')) {
          salesData.setFormula(idx + 2, col + 1, value);
        } else {
          salesData.setValue(idx + 2, col + 1, value);
        }
      });
    });
    
    // Data validation
    salesData.addValidation('L2:L1000', ['Website', 'Amazon', 'eBay', 'Social', 'Retail', 'Marketplace']);
    salesData.addValidation('M2:M1000', ['Credit Card', 'PayPal', 'Bank Transfer', 'Cash', 'Gift Card']);
    
    // Formatting
    salesData.format(2, 1, 1000, 1, { numberFormat: 'dd/mm/yyyy' });
    salesData.format(2, 6, 1000, 6, { numberFormat: '$#,##0.00' });
    salesData.format(2, 8, 1000, 10, { numberFormat: '$#,##0.00' });
    salesData.format(2, 11, 1000, 11, { numberFormat: '0.0%' });
    salesData.format(2, 14, 1000, 17, { numberFormat: '$#,##0.00' });
    
    // Column widths
    salesData.setColumnWidth(1, 100);
    salesData.setColumnWidth(2, 100);
    salesData.setColumnWidth(5, 150);
    for (let col = 3; col <= headers.length; col++) {
      if (col !== 5) salesData.setColumnWidth(col, 90);
    }
  },
  
  buildProductsSheet: function(products) {
    // Product catalog with performance metrics
    const headers = [
      'SKU', 'Product Name', 'Category', 'Brand', 'Sales (30d)', 'Units Sold',
      'Revenue', 'Margin %', 'Cost', 'Current Stock', 'Reorder Point',
      'Supplier', 'Last Ordered', 'Rating', 'Reviews', 'Status'
    ];
    
    headers.forEach((header, idx) => {
      products.setValue(1, idx + 1, header);
    });
    
    products.format(1, 1, 1, headers.length, {
      fontWeight: 'bold',
      background: '#DC2626',
      fontColor: '#FFFFFF',
      border: true
    });
    
    // Sample products with comprehensive data
    const sampleProducts = [
      ['SKU001', 'Wireless Headphones Pro', 'Electronics', 'AudioTech', 45, 85, 16999, 65, 6500, 24, 15, 'TechSupplier Co', '=TODAY()-15', 4.2, 128, 'Active'],
      ['SKU002', 'Smart Fitness Watch', 'Wearables', 'FitTrack', 32, 32, 9600, 58, 4800, 8, 10, 'WearablePlus', '=TODAY()-8', 4.5, 89, 'Low Stock'],
      ['SKU003', 'Bluetooth Speaker Mini', 'Electronics', 'SoundWave', 78, 78, 7020, 72, 2340, 45, 20, 'AudioSource', '=TODAY()-22', 3.9, 156, 'Active'],
      ['SKU004', 'Premium Phone Case', 'Accessories', 'ProtectPlus', 156, 234, 5850, 78, 1755, 67, 25, 'AccessoryWorld', '=TODAY()-5', 4.1, 203, 'Active'],
      ['SKU005', 'Adjustable Laptop Stand', 'Office', 'WorkSpace', 28, 28, 1400, 64, 1120, 12, 15, 'OfficeGear Inc', '=TODAY()-12', 4.3, 67, 'Active'],
      ['SKU006', 'Wireless Mouse Elite', 'Accessories', 'TechFlow', 65, 65, 3250, 69, 1625, 3, 20, 'TechSupplier Co', '=TODAY()-3', 4.0, 145, 'Critical Stock'],
      ['SKU007', 'USB-C Hub 7-in-1', 'Accessories', 'ConnectPro', 42, 42, 2940, 55, 1764, 23, 15, 'TechFlow Ltd', '=TODAY()-18', 4.4, 78, 'Active'],
      ['SKU008', 'Portable Charger 20000mAh', 'Electronics', 'PowerMax', 89, 89, 4450, 62, 2225, 41, 25, 'PowerTech', '=TODAY()-10', 4.2, 167, 'Active'],
      ['SKU009', 'Gaming Keyboard RGB', 'Gaming', 'GamePro', 34, 34, 5100, 66, 2550, 18, 12, 'GameGear Co', '=TODAY()-20', 4.6, 94, 'Active'],
      ['SKU010', 'Webcam HD 1080p', 'Electronics', 'ViewTech', 52, 52, 3900, 58, 2340, 29, 18, 'CameraTech', '=TODAY()-14', 3.8, 112, 'Active']
    ];
    
    sampleProducts.forEach((product, idx) => {
      product.forEach((value, col) => {
        if (typeof value === 'string' && value.startsWith('=')) {
          products.setFormula(idx + 2, col + 1, value);
        } else {
          products.setValue(idx + 2, col + 1, value);
        }
      });
    });
    
    // Data validation
    products.addValidation('C2:C1000', ['Electronics', 'Wearables', 'Accessories', 'Office', 'Gaming', 'Home', 'Fitness']);
    products.addValidation('P2:P1000', ['Active', 'Low Stock', 'Out of Stock', 'Discontinued', 'Critical Stock']);
    
    // Conditional formatting for stock status
    products.addConditionalFormat(2, 16, 1000, 16, {
      type: 'cell',
      condition: 'TEXT_EQUALS',
      value: 'Critical Stock',
      format: { background: '#FEE2E2', fontColor: '#DC2626' }
    });
    
    products.addConditionalFormat(2, 16, 1000, 16, {
      type: 'cell',
      condition: 'TEXT_EQUALS',
      value: 'Low Stock',
      format: { background: '#FEF3C7', fontColor: '#D97706' }
    });
    
    // Formatting
    products.format(2, 7, 1000, 7, { numberFormat: '$#,##0' });
    products.format(2, 8, 1000, 8, { numberFormat: '0%' });
    products.format(2, 9, 1000, 9, { numberFormat: '$#,##0' });
    products.format(2, 13, 1000, 13, { numberFormat: 'dd/mm/yyyy' });
    products.format(2, 14, 1000, 14, { numberFormat: '0.0' });
    
    // Column widths
    products.setColumnWidth(1, 80);
    products.setColumnWidth(2, 180);
    products.setColumnWidth(3, 100);
    for (let col = 4; col <= headers.length; col++) {
      products.setColumnWidth(col, 90);
    }
  },
  
  buildCustomersSheet: function(customers) {
    // Customer database with analytics
    const headers = [
      'Customer ID', 'Name', 'Email', 'Join Date', 'Status', 'LTV',
      'Last Order', 'Total Orders', 'Total Spent', 'AOV', 'Segment',
      'Location', 'Age', 'Gender', 'Preferred Channel', 'Notes'
    ];
    
    headers.forEach((header, idx) => {
      customers.setValue(1, idx + 1, header);
    });
    
    customers.format(1, 1, 1, headers.length, {
      fontWeight: 'bold',
      background: '#DC2626',
      fontColor: '#FFFFFF',
      border: true
    });
    
    // Sample customers
    const sampleCustomers = [
      ['CUST001', 'John Smith', 'john@email.com', '=TODAY()-365', 'VIP', 1250, '=TODAY()-5', 12, 1450, '=I2/H2', 'High Value', 'New York, NY', 34, 'M', 'Website', 'Frequent electronics buyer'],
      ['CUST002', 'Sarah Johnson', 'sarah@email.com', '=TODAY()-180', 'Active', 680, '=TODAY()-12', 6, 750, '=I3/H3', 'Regular', 'Los Angeles, CA', 28, 'F', 'Amazon', 'Prefers premium products'],
      ['CUST003', 'Mike Wilson', 'mike@email.com', '=TODAY()-90', 'New', 299, '=TODAY()-30', 3, 320, '=I4/H4', 'New Customer', 'Chicago, IL', 45, 'M', 'Social', 'Price-sensitive buyer'],
      ['CUST004', 'Emily Davis', 'emily@email.com', '=TODAY()-450', 'Loyal', 2100, '=TODAY()-8', 18, 2250, '=I5/H5', 'VIP', 'Miami, FL', 31, 'F', 'Website', 'Tech enthusiast, early adopter'],
      ['CUST005', 'David Brown', 'david@email.com', '=TODAY()-220', 'Active', 890, '=TODAY()-20', 8, 945, '=I6/H6', 'Regular', 'Seattle, WA', 39, 'M', 'eBay', 'Gaming accessories focus']
    ];
    
    sampleCustomers.forEach((customer, idx) => {
      customer.forEach((value, col) => {
        if (typeof value === 'string' && value.startsWith('=')) {
          customers.setFormula(idx + 2, col + 1, value);
        } else {
          customers.setValue(idx + 2, col + 1, value);
        }
      });
    });
    
    // Data validation
    customers.addValidation('E2:E1000', ['New', 'Active', 'VIP', 'Loyal', 'Inactive', 'Churned']);
    customers.addValidation('K2:K1000', ['New Customer', 'Regular', 'High Value', 'VIP', 'At Risk']);
    customers.addValidation('N2:N1000', ['M', 'F', 'Other', 'Prefer not to say']);
    customers.addValidation('O2:O1000', ['Website', 'Amazon', 'eBay', 'Social', 'Retail', 'Mobile App']);
    
    // Formatting
    customers.format(2, 4, 1000, 4, { numberFormat: 'dd/mm/yyyy' });
    customers.format(2, 6, 1000, 6, { numberFormat: '$#,##0' });
    customers.format(2, 7, 1000, 7, { numberFormat: 'dd/mm/yyyy' });
    customers.format(2, 9, 1000, 10, { numberFormat: '$#,##0' });
    
    // Column widths
    customers.setColumnWidth(1, 100);
    customers.setColumnWidth(2, 120);
    customers.setColumnWidth(3, 150);
    for (let col = 4; col <= headers.length; col++) {
      customers.setColumnWidth(col, 90);
    }
  },
  
  buildOrdersSheet: function(orders) {
    // Order management and fulfillment
    const headers = [
      'Order ID', 'Customer ID', 'Order Date', 'Items', 'Total Amount',
      'Channel', 'Payment Status', 'Customer Type', 'Shipping Method', 'Tracking',
      'Fulfillment Status', 'Shipped Date', 'Delivery Date', 'Return', 'Notes'
    ];
    
    headers.forEach((header, idx) => {
      orders.setValue(1, idx + 1, header);
    });
    
    orders.format(1, 1, 1, headers.length, {
      fontWeight: 'bold',
      background: '#DC2626',
      fontColor: '#FFFFFF',
      border: true
    });
    
    // Sample orders
    const sampleOrders = [
      ['ORD001', 'CUST001', '=TODAY()-5', 2, 399.98, 'Website', 'Paid', 'Returning', 'Standard', 'TRK12345', 'Delivered', '=TODAY()-3', '=TODAY()-1', 'No', 'Customer satisfied'],
      ['ORD002', 'CUST002', '=TODAY()-8', 1, 279.99, 'Amazon', 'Paid', 'Returning', 'Prime', 'TRK12346', 'Delivered', '=TODAY()-6', '=TODAY()-5', 'No', 'Prime delivery'],
      ['ORD003', 'CUST003', '=TODAY()-12', 3, 179.97, 'Website', 'Paid', 'New', 'Express', 'TRK12347', 'Shipped', '=TODAY()-10', '', 'No', 'First order, express'],
      ['ORD004', 'CUST001', '=TODAY()-15', 1, 49.99, 'eBay', 'Paid', 'Returning', 'Standard', 'TRK12348', 'Processing', '', '', 'No', 'Processing order'],
      ['ORD005', 'CUST004', '=TODAY()-18', 4, 589.96, 'Website', 'Paid', 'VIP', 'Free', 'TRK12349', 'Delivered', '=TODAY()-16', '=TODAY()-14', 'No', 'VIP customer, free shipping']
    ];
    
    sampleOrders.forEach((order, idx) => {
      order.forEach((value, col) => {
        if (typeof value === 'string' && value.startsWith('=')) {
          orders.setFormula(idx + 2, col + 1, value);
        } else {
          orders.setValue(idx + 2, col + 1, value);
        }
      });
    });
    
    // Data validation
    orders.addValidation('F2:F1000', ['Website', 'Amazon', 'eBay', 'Social', 'Retail', 'Mobile App']);
    orders.addValidation('G2:G1000', ['Pending', 'Paid', 'Failed', 'Refunded']);
    orders.addValidation('H2:H1000', ['New', 'Returning', 'VIP', 'Loyal']);
    orders.addValidation('I2:I1000', ['Standard', 'Express', 'Prime', 'Free', 'Next Day']);
    orders.addValidation('K2:K1000', ['Processing', 'Shipped', 'Delivered', 'Returned', 'Cancelled']);
    orders.addValidation('N2:N1000', ['No', 'Requested', 'Approved', 'Completed']);
    
    // Conditional formatting
    orders.addConditionalFormat(2, 11, 1000, 11, {
      type: 'cell',
      condition: 'TEXT_EQUALS',
      value: 'Delivered',
      format: { background: '#D1FAE5' }
    });
    
    orders.addConditionalFormat(2, 11, 1000, 11, {
      type: 'cell',
      condition: 'TEXT_EQUALS',
      value: 'Processing',
      format: { background: '#FEF3C7' }
    });
    
    // Formatting
    orders.format(2, 3, 1000, 3, { numberFormat: 'dd/mm/yyyy' });
    orders.format(2, 5, 1000, 5, { numberFormat: '$#,##0.00' });
    orders.format(2, 12, 1000, 13, { numberFormat: 'dd/mm/yyyy' });
    
    // Column widths
    orders.setColumnWidth(1, 100);
    orders.setColumnWidth(2, 100);
    for (let col = 3; col <= headers.length; col++) {
      orders.setColumnWidth(col, 90);
    }
  },
  
  buildChannelsSheet: function(channels) {
    // Channel performance analysis
    const headers = [
      'Channel', 'Traffic', 'Visitors', 'Conversion Rate', 'Orders', 'Revenue',
      'AOV', 'Cost', 'Profit', 'ROI', 'Trend', 'Status'
    ];
    
    headers.forEach((header, idx) => {
      channels.setValue(1, idx + 1, header);
    });
    
    channels.format(1, 1, 1, headers.length, {
      fontWeight: 'bold',
      background: '#DC2626',
      fontColor: '#FFFFFF',
      border: true
    });
    
    // Sample channel data
    const channelData = [
      ['Website', 125000, 45000, '3.2%', '=B2*D2', '=E2*RANDBETWEEN(75,125)', '=F2/E2', 15000, '=F2-H2', '=I2/H2', '+15%', 'Growing'],
      ['Amazon', 89000, 32000, '4.1%', '=B3*D3', '=E3*RANDBETWEEN(85,135)', '=F3/E3', 12000, '=F3-H3', '=I3/H3', '+8%', 'Stable'],
      ['eBay', 45000, 18000, '2.8%', '=B4*D4', '=E4*RANDBETWEEN(65,95)', '=F4/E4', 8000, '=F4-H4', '=I4/H4', '+3%', 'Declining'],
      ['Social Media', 78000, 28000, '1.9%', '=B5*D5', '=E5*RANDBETWEEN(55,85)', '=F5/E5', 18000, '=F5-H5', '=I5/H5', '+25%', 'Emerging'],
      ['Mobile App', 32000, 15000, '5.2%', '=B6*D6', '=E6*RANDBETWEEN(95,145)', '=F6/E6', 5000, '=F6-H6', '=I6/H6', '+45%', 'Hot'],
      ['Retail Partners', 25000, 12000, '3.8%', '=B7*D7', '=E7*RANDBETWEEN(105,155)', '=F7/E7', 10000, '=F7-H7', '=I7/H7', '-5%', 'Watch']
    ];
    
    channelData.forEach((channel, idx) => {
      channel.forEach((value, col) => {
        if (typeof value === 'string' && value.startsWith('=')) {
          channels.setFormula(idx + 2, col + 1, value);
        } else {
          channels.setValue(idx + 2, col + 1, value);
        }
      });
    });
    
    // Totals row
    channels.setValue(8, 1, 'TOTAL');
    for (let col = 2; col <= 6; col++) {
      channels.setFormula(8, col, `=SUM(${String.fromCharCode(64 + col)}2:${String.fromCharCode(64 + col)}7)`);
    }
    channels.setFormula(8, 7, '=F8/E8');
    for (let col = 8; col <= 10; col++) {
      channels.setFormula(8, col, `=SUM(${String.fromCharCode(64 + col)}2:${String.fromCharCode(64 + col)}7)`);
    }
    
    channels.format(8, 1, 8, headers.length, {
      fontWeight: 'bold',
      background: '#FECACA'
    });
    
    // Formatting
    channels.format(2, 2, 8, 3, { numberFormat: '#,##0' });
    channels.format(2, 4, 8, 4, { numberFormat: '0.0%' });
    channels.format(2, 5, 8, 5, { numberFormat: '#,##0' });
    channels.format(2, 6, 8, 9, { numberFormat: '$#,##0' });
    channels.format(2, 10, 8, 10, { numberFormat: '0.0x' });
    
    // Column widths
    channels.setColumnWidth(1, 120);
    for (let col = 2; col <= headers.length; col++) {
      channels.setColumnWidth(col, 90);
    }
  },
  
  buildTrendsSheet: function(trends) {
    // Sales trends and forecasting
    trends.setValue(1, 1, 'Sales Trends & Forecasting');
    trends.merge(1, 1, 1, 10);
    trends.format(1, 1, 1, 10, {
      fontSize: 16,
      fontWeight: 'bold',
      horizontalAlignment: 'center',
      background: '#DC2626',
      fontColor: '#FFFFFF'
    });
    
    // Daily trends
    const headers = [
      'Date', 'Revenue', 'Orders', 'AOV', 'New Customers', 'Returning %',
      'Conversion Rate', 'Traffic', 'Forecast', 'Variance'
    ];
    
    headers.forEach((header, idx) => {
      trends.setValue(3, idx + 1, header);
    });
    
    trends.format(3, 1, 3, headers.length, {
      fontWeight: 'bold',
      background: '#FECACA',
      border: true
    });
    
    // Generate 30 days of trend data
    for (let day = 29; day >= 0; day--) {
      const row = 4 + (29 - day);
      trends.setFormula(row, 1, `=TODAY()-${day}`);
      
      // Base revenue with trend and seasonality
      const baseRevenue = 25000;
      const trend = day * 100; // Growing trend
      const seasonality = Math.sin(day * 0.2) * 3000; // Weekly pattern
      const randomness = Math.random() * 4000 - 2000;
      
      trends.setValue(row, 2, Math.max(0, baseRevenue + trend + seasonality + randomness));
      trends.setValue(row, 3, Math.floor(Math.random() * 200 + 150));
      trends.setFormula(row, 4, `=B${row}/C${row}`);
      trends.setValue(row, 5, Math.floor(Math.random() * 50 + 20));
      trends.setValue(row, 6, 0.3 + Math.random() * 0.4);
      trends.setValue(row, 7, 0.02 + Math.random() * 0.03);
      trends.setValue(row, 8, Math.floor(Math.random() * 5000 + 8000));
      
      // Simple linear forecast
      if (row <= 25) {
        trends.setFormula(row, 9, `=TREND(B4:B${row},ROW(B4:B${row}),ROW(B${row+1}))`); 
      }
      
      trends.setFormula(row, 10, `=IF(I${row}<>"",B${row}-I${row},"")`);
    }
    
    // Formatting
    trends.format(4, 1, 33, 1, { numberFormat: 'dd/mm/yyyy' });
    trends.format(4, 2, 33, 2, { numberFormat: '$#,##0' });
    trends.format(4, 4, 33, 4, { numberFormat: '$#,##0' });
    trends.format(4, 6, 33, 7, { numberFormat: '0.0%' });
    trends.format(4, 9, 33, 10, { numberFormat: '$#,##0' });
    
    // Trend analysis
    trends.setValue(35, 1, 'TREND ANALYSIS');
    trends.merge(35, 1, 35, 6);
    trends.format(35, 1, 35, 6, {
      fontWeight: 'bold',
      background: '#FEE2E2'
    });
    
    const analysisMetrics = [
      ['30-Day Avg Revenue:', '=AVERAGE(B4:B33)'],
      ['Revenue Growth Rate:', '=(B33-B4)/B4'],
      ['Best Day Revenue:', '=MAX(B4:B33)'],
      ['Worst Day Revenue:', '=MIN(B4:B33)'],
      ['Volatility:', '=STDEV(B4:B33)/AVERAGE(B4:B33)'],
      ['Forecast Accuracy:', '=1-AVERAGE(ABS(J4:J28))/AVERAGE(B4:B28)']
    ];
    
    analysisMetrics.forEach((metric, idx) => {
      trends.setValue(36 + idx, 1, metric[0]);
      trends.setFormula(36 + idx, 2, metric[1]);
    });
    
    trends.format(37, 2, 37, 2, { numberFormat: '0.0%' });
    trends.format(38, 2, 40, 2, { numberFormat: '$#,##0' });
    trends.format(41, 2, 42, 2, { numberFormat: '0.0%' });
    
    // Column widths
    trends.setColumnWidth(1, 100);
    for (let col = 2; col <= headers.length; col++) {
      trends.setColumnWidth(col, 90);
    }
  },
  
  buildSalesReportsSheet: function(reports) {
    // Executive sales reporting
    reports.setValue(1, 1, 'E-commerce Executive Sales Report');
    reports.merge(1, 1, 1, 10);
    reports.format(1, 1, 1, 10, {
      fontSize: 16,
      fontWeight: 'bold',
      horizontalAlignment: 'center',
      background: '#DC2626',
      fontColor: '#FFFFFF'
    });
    
    // Executive summary section
    reports.setValue(3, 1, 'EXECUTIVE SUMMARY');
    reports.merge(3, 1, 3, 10);
    reports.format(3, 1, 3, 10, {
      fontWeight: 'bold',
      background: '#FEE2E2'
    });
    
    const summaryItems = [
      ['Total Revenue (30d):', '=SUM({{Sales Data}}!F:F)'],
      ['Total Orders:', '=COUNTIF({{Orders}}!A:A,"<>")'],
      ['Average Order Value:', '=A5/B5'],
      ['New Customers:', '=COUNTIFS({{Customers}}!D:D,">="&TODAY()-30)'],
      ['Customer Retention:', '=COUNTIFS({{Orders}}!H:H,"Returning")/B5'],
      ['Top Product Category:', '=INDEX({{Products}}!C:C,MATCH(MAX({{Products}}!G:G),{{Products}}!G:G,0))'],
      ['Best Sales Channel:', '=INDEX({{Channels}}!A:A,MATCH(MAX({{Channels}}!F:F),{{Channels}}!F:F,0))'],
      ['Inventory Turnover:', '=SUM({{Products}}!F:F)/AVERAGE({{Products}}!J:J)*12']
    ];
    
    summaryItems.forEach((item, idx) => {
      reports.setValue(4 + idx, 1, item[0]);
      reports.setFormula(4 + idx, 2, item[1]);
      reports.merge(4 + idx, 2, 4 + idx, 3);
    });
    
    reports.format(4, 2, 4, 3, { numberFormat: '$#,##0' });
    reports.format(6, 2, 6, 3, { numberFormat: '$#,##0.00' });
    reports.format(8, 2, 8, 3, { numberFormat: '0.0%' });
    reports.format(11, 2, 11, 3, { numberFormat: '0.0x' });
    
    // Performance highlights
    reports.setValue(3, 5, 'PERFORMANCE HIGHLIGHTS');
    reports.merge(3, 5, 3, 10);
    reports.format(3, 5, 3, 10, {
      fontWeight: 'bold',
      background: '#FEE2E2'
    });
    
    const highlights = [
      '✅ Revenue growth of 18% vs last month',
      '✅ Conversion rate improved to 3.5%',
      '📈 Mobile app sales up 45%',
      '🔥 Top product: Wireless Headphones Pro',
      '⚠️ 6 products below reorder point',
      '📊 Amazon channel performing strongly'
    ];
    
    highlights.forEach((highlight, idx) => {
      reports.setValue(4 + idx, 5, highlight);
      reports.merge(4 + idx, 5, 4 + idx, 10);
    });
    
    // Channel breakdown
    reports.setValue(13, 1, 'CHANNEL PERFORMANCE BREAKDOWN');
    reports.merge(13, 1, 13, 10);
    reports.format(13, 1, 13, 10, {
      fontWeight: 'bold',
      background: '#FEE2E2'
    });
    
    const channelHeaders = ['Channel', 'Revenue', '% of Total', 'Orders', 'AOV', 'Growth'];
    channelHeaders.forEach((header, idx) => {
      reports.setValue(14, idx + 1, header);
    });
    reports.format(14, 1, 14, 6, {
      fontWeight: 'bold',
      background: '#E5E7EB'
    });
    
    // Pull top 5 channels
    for (let i = 0; i < 5; i++) {
      reports.setFormula(15 + i, 1, `=INDEX({{Channels}}!A:A,${i + 2})`);
      reports.setFormula(15 + i, 2, `=INDEX({{Channels}}!F:F,${i + 2})`);
      reports.setFormula(15 + i, 3, `=B${15+i}/SUM(B$15:B$19)`);
      reports.setFormula(15 + i, 4, `=INDEX({{Channels}}!E:E,${i + 2})`);
      reports.setFormula(15 + i, 5, `=INDEX({{Channels}}!G:G,${i + 2})`);
      reports.setFormula(15 + i, 6, `=INDEX({{Channels}}!K:K,${i + 2})`);
    }
    
    reports.format(15, 2, 19, 2, { numberFormat: '$#,##0' });
    reports.format(15, 3, 19, 3, { numberFormat: '0.0%' });
    reports.format(15, 5, 19, 5, { numberFormat: '$#,##0' });
    
    // Action items
    reports.setValue(21, 1, 'ACTION ITEMS & RECOMMENDATIONS');
    reports.merge(21, 1, 21, 10);
    reports.format(21, 1, 21, 10, {
      fontWeight: 'bold',
      background: '#FEE2E2'
    });
    
    const actions = [
      '1. Restock 6 products currently below reorder points',
      '2. Increase marketing spend on Mobile App channel (+45% growth)',
      '3. Investigate declining eBay performance (-5% trend)',
      '4. Launch retention campaign for customers not ordered in 30+ days',
      '5. Analyze top-performing product categories for expansion',
      '6. Optimize website conversion rate (currently 3.2%)',
      '7. Review pricing strategy for slow-moving inventory'
    ];
    
    actions.forEach((action, idx) => {
      reports.setValue(22 + idx, 1, action);
      reports.merge(22 + idx, 1, 22 + idx, 10);
    });
    
    // Column widths
    reports.setColumnWidth(1, 200);
    for (let col = 2; col <= 10; col++) {
      reports.setColumnWidth(col, 90);
    }
  },
  
  /**
   * Additional template stubs - would be implemented with similar comprehensive patterns
   */
  buildInventoryManager: function(builder) {
    // Inventory tracking, reorder points, supplier management
    return builder.complete();
  },
  
  buildCustomerAnalytics: function(builder) {
    // Customer segmentation, LTV, churn analysis
    return builder.complete();
  },
  
  buildProductPerformance: function(builder) {
    // Product analysis, category performance, pricing optimization
    return builder.complete();
  },
  
  buildFinancialTracker: function(builder) {
    // P&L, cash flow, profitability analysis
    return builder.complete();
  },
  
  buildMarketingROI: function(builder) {
    // Marketing attribution, campaign ROI, customer acquisition costs
    return builder.complete();
  },
  
  buildSupplierManagement: function(builder) {
    // Supplier performance, lead times, cost analysis
    return builder.complete();
  },
  
  buildShippingLogistics: function(builder) {
    // Shipping costs, delivery performance, carrier analysis
    return builder.complete();
  },
  
  buildPriceOptimizer: function(builder) {
    // Dynamic pricing, competitive analysis, margin optimization
    return builder.complete();
  },
  
  buildForecasting: function(builder) {
    // Demand forecasting, inventory planning, sales predictions
    return builder.complete();
  }
};