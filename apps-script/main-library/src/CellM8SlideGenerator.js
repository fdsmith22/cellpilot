/**
 * CellM8 Intelligent Slide Generator
 * Professional slide generation with smart data visualization
 */

const CellM8SlideGenerator = {
  
  // Slide dimension constants (in points)
  SLIDE_WIDTH: 720,
  SLIDE_HEIGHT: 405,
  
  // Safe margins for content
  MARGINS: {
    top: 60,
    bottom: 40,
    left: 40,
    right: 40
  },
  
  // Typography settings
  FONTS: {
    title: { size: 32, family: 'Arial', weight: 'bold' },
    subtitle: { size: 20, family: 'Arial', weight: 'normal' },
    heading: { size: 24, family: 'Arial', weight: 'bold' },
    body: { size: 14, family: 'Arial', weight: 'normal' },
    caption: { size: 11, family: 'Arial', weight: 'normal' },
    data: { size: 10, family: 'Arial', weight: 'normal' }
  },
  
  // Color schemes for professional templates
  THEMES: {
    professional: {
      primary: '#1A73E8',
      secondary: '#5F6368',
      accent: '#34A853',
      background: '#FFFFFF',
      text: '#202124',
      lightGray: '#F8F9FA',
      border: '#DADCE0'
    },
    modern: {
      primary: '#6366F1',
      secondary: '#64748B',
      accent: '#8B5CF6',
      background: '#FFFFFF',
      text: '#1E293B',
      lightGray: '#F1F5F9',
      border: '#E2E8F0'
    },
    elegant: {
      primary: '#2563EB',
      secondary: '#475569',
      accent: '#0EA5E9',
      background: '#FFFFFF',
      text: '#111827',
      lightGray: '#F9FAFB',
      border: '#E5E7EB'
    }
  },
  
  /**
   * Analyze data to determine best visualization
   */
  analyzeDataForVisualization: function(data) {
    if (!data || !data.headers || !data.data) return null;
    
    const analysis = {
      dataTypes: [],
      numericColumns: [],
      categoricalColumns: [],
      dateColumns: [],
      recommendations: []
    };
    
    // Analyze each column
    for (let i = 0; i < data.headers.length; i++) {
      const columnData = data.data.map(row => row[i]);
      const dataType = this.detectColumnDataType(columnData);
      
      analysis.dataTypes.push(dataType);
      
      if (dataType === 'numeric') {
        analysis.numericColumns.push({
          index: i,
          name: data.headers[i],
          stats: this.calculateStats(columnData)
        });
      } else if (dataType === 'date') {
        analysis.dateColumns.push({
          index: i,
          name: data.headers[i]
        });
      } else {
        analysis.categoricalColumns.push({
          index: i,
          name: data.headers[i],
          uniqueValues: [...new Set(columnData)].length
        });
      }
    }
    
    // Generate visualization recommendations
    if (analysis.dateColumns.length > 0 && analysis.numericColumns.length > 0) {
      analysis.recommendations.push('LINE_CHART'); // Time series
    }
    if (analysis.categoricalColumns.length > 0 && analysis.numericColumns.length > 0) {
      analysis.recommendations.push('COLUMN_CHART'); // Category comparison
    }
    if (analysis.numericColumns.length >= 2) {
      analysis.recommendations.push('SCATTER_CHART'); // Correlation
    }
    if (analysis.categoricalColumns.length === 1 && analysis.numericColumns.length === 1) {
      analysis.recommendations.push('PIE_CHART'); // Distribution
    }
    
    return analysis;
  },
  
  /**
   * Detect column data type
   */
  detectColumnDataType: function(columnData) {
    let numericCount = 0;
    let dateCount = 0;
    let nonEmpty = 0;
    
    for (let value of columnData) {
      if (value === '' || value === null) continue;
      nonEmpty++;
      
      // Check if numeric
      if (!isNaN(value) && typeof value === 'number') {
        numericCount++;
      }
      // Check if date
      else if (value instanceof Date || this.isDateString(value)) {
        dateCount++;
      }
    }
    
    if (numericCount > nonEmpty * 0.8) return 'numeric';
    if (dateCount > nonEmpty * 0.8) return 'date';
    return 'categorical';
  },
  
  /**
   * Check if string is a date
   */
  isDateString: function(str) {
    if (typeof str !== 'string') return false;
    const datePatterns = [
      /^\d{4}-\d{2}-\d{2}$/,
      /^\d{2}\/\d{2}\/\d{4}$/,
      /^\d{2}-\d{2}-\d{4}$/
    ];
    return datePatterns.some(pattern => pattern.test(str));
  },
  
  /**
   * Calculate statistics for numeric column
   */
  calculateStats: function(data) {
    const numbers = data.filter(v => !isNaN(v) && v !== '' && v !== null).map(Number);
    if (numbers.length === 0) return null;
    
    numbers.sort((a, b) => a - b);
    const sum = numbers.reduce((a, b) => a + b, 0);
    
    return {
      min: numbers[0],
      max: numbers[numbers.length - 1],
      mean: sum / numbers.length,
      median: numbers[Math.floor(numbers.length / 2)],
      count: numbers.length
    };
  },
  
  /**
   * Generate professional title slide
   */
  createTitleSlide: function(presentation, config, theme) {
    try {
      const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
      const colors = this.THEMES[theme] || this.THEMES.professional;
      
      // Add gradient background
      const background = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 0, 0, this.SLIDE_WIDTH, this.SLIDE_HEIGHT);
      background.getFill().setSolidFill(colors.background);
      background.getBorder().setTransparent();
      
      // Add accent bar
      const accentBar = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 0, 0, this.SLIDE_WIDTH, 4);
      accentBar.getFill().setSolidFill(colors.primary);
      accentBar.getBorder().setTransparent();
      
      // Add title with proper positioning
      const titleBox = slide.insertTextBox(
        config.title,
        this.MARGINS.left,
        this.SLIDE_HEIGHT * 0.35,
        this.SLIDE_WIDTH - this.MARGINS.left - this.MARGINS.right,
        60
      );
      titleBox.getText().getTextStyle()
        .setFontSize(this.FONTS.title.size)
        .setFontFamily(this.FONTS.title.family)
        .setBold(true)
        .setForegroundColor(colors.text);
      titleBox.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
      
      // Add subtitle if provided
      if (config.subtitle) {
        const subtitleBox = slide.insertTextBox(
          config.subtitle,
          this.MARGINS.left,
          this.SLIDE_HEIGHT * 0.5,
          this.SLIDE_WIDTH - this.MARGINS.left - this.MARGINS.right,
          40
        );
        subtitleBox.getText().getTextStyle()
          .setFontSize(this.FONTS.subtitle.size)
          .setFontFamily(this.FONTS.subtitle.family)
          .setForegroundColor(colors.secondary);
        subtitleBox.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
      }
      
      // Add date/time stamp
      const dateBox = slide.insertTextBox(
        new Date().toLocaleDateString(),
        this.MARGINS.left,
        this.SLIDE_HEIGHT - this.MARGINS.bottom - 20,
        200,
        20
      );
      dateBox.getText().getTextStyle()
        .setFontSize(this.FONTS.caption.size)
        .setForegroundColor(colors.secondary);
      
      return slide;
    } catch (error) {
      Logger.error('Error creating title slide:', error);
      return null;
    }
  },
  
  /**
   * Create data visualization slide with chart
   */
  createChartSlide: function(presentation, data, analysis, title, theme) {
    try {
      const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
      const colors = this.THEMES[theme] || this.THEMES.professional;
      
      // Add title
      const titleBox = slide.insertTextBox(
        title,
        this.MARGINS.left,
        this.MARGINS.top - 40,
        this.SLIDE_WIDTH - this.MARGINS.left - this.MARGINS.right,
        40
      );
      titleBox.getText().getTextStyle()
        .setFontSize(this.FONTS.heading.size)
        .setFontFamily(this.FONTS.heading.family)
        .setBold(true)
        .setForegroundColor(colors.text);
      
      // Determine best chart type
      const chartType = analysis.recommendations[0] || 'COLUMN_CHART';
      
      // Prepare data for chart
      const chartData = this.prepareChartData(data, analysis, chartType);
      
      if (chartData) {
        // Create data table for chart
        const dataTable = Charts.newDataTable();
        
        // Add headers
        for (let header of chartData.headers) {
          dataTable.addColumn(Charts.ColumnType.STRING, header);
        }
        
        // Add data rows
        for (let row of chartData.rows) {
          dataTable.addRow(row);
        }
        
        // Build chart
        const chartBuilder = this.getChartBuilder(chartType, dataTable, colors);
        const chart = chartBuilder
          .setDimensions(
            this.SLIDE_WIDTH - this.MARGINS.left - this.MARGINS.right,
            this.SLIDE_HEIGHT - this.MARGINS.top - this.MARGINS.bottom - 40
          )
          .build();
        
        // Add chart to slide
        slide.insertImage(
          chart.getBlob(),
          this.MARGINS.left,
          this.MARGINS.top,
          this.SLIDE_WIDTH - this.MARGINS.left - this.MARGINS.right,
          this.SLIDE_HEIGHT - this.MARGINS.top - this.MARGINS.bottom - 40
        );
      }
      
      return slide;
    } catch (error) {
      Logger.error('Error creating chart slide:', error);
      return null;
    }
  },
  
  /**
   * Get appropriate chart builder
   */
  getChartBuilder: function(chartType, dataTable, colors) {
    let builder;
    
    switch(chartType) {
      case 'LINE_CHART':
        builder = Charts.newLineChart();
        break;
      case 'COLUMN_CHART':
        builder = Charts.newColumnChart();
        break;
      case 'PIE_CHART':
        builder = Charts.newPieChart();
        break;
      case 'SCATTER_CHART':
        builder = Charts.newScatterChart();
        break;
      default:
        builder = Charts.newColumnChart();
    }
    
    return builder
      .setDataTable(dataTable)
      .setOption('backgroundColor', colors.background)
      .setOption('colors', [colors.primary, colors.accent, colors.secondary])
      .setOption('fontName', this.FONTS.body.family)
      .setOption('fontSize', this.FONTS.body.size);
  },
  
  /**
   * Prepare data for chart
   */
  prepareChartData: function(data, analysis, chartType) {
    const result = {
      headers: [],
      rows: []
    };
    
    // Get first categorical and numeric columns for simple chart
    if (analysis.categoricalColumns.length > 0 && analysis.numericColumns.length > 0) {
      const catCol = analysis.categoricalColumns[0];
      const numCol = analysis.numericColumns[0];
      
      result.headers = [catCol.name, numCol.name];
      
      // Aggregate data by category
      const aggregated = {};
      for (let row of data.data) {
        const category = String(row[catCol.index]);
        const value = Number(row[numCol.index]) || 0;
        
        if (aggregated[category]) {
          aggregated[category] += value;
        } else {
          aggregated[category] = value;
        }
      }
      
      // Convert to rows
      for (let [category, value] of Object.entries(aggregated)) {
        result.rows.push([category, value]);
      }
      
      // Limit to top 10 for readability
      result.rows = result.rows.slice(0, 10);
    }
    
    return result;
  },
  
  /**
   * Create professional data table slide
   */
  createDataTableSlide: function(presentation, data, title, theme) {
    try {
      const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
      const colors = this.THEMES[theme] || this.THEMES.professional;
      
      // Add title
      const titleBox = slide.insertTextBox(
        title,
        this.MARGINS.left,
        20,
        this.SLIDE_WIDTH - this.MARGINS.left - this.MARGINS.right,
        40
      );
      titleBox.getText().getTextStyle()
        .setFontSize(this.FONTS.heading.size)
        .setFontFamily(this.FONTS.heading.family)
        .setBold(true)
        .setForegroundColor(colors.text);
      
      // Calculate table dimensions
      const tableWidth = this.SLIDE_WIDTH - this.MARGINS.left - this.MARGINS.right;
      const tableHeight = this.SLIDE_HEIGHT - this.MARGINS.top - this.MARGINS.bottom - 20;
      
      // Determine optimal table size
      const maxRows = Math.min(8, data.data.length + 1); // +1 for header
      const maxCols = Math.min(6, data.headers.length);
      
      // Create table with proper sizing
      const table = slide.insertTable(
        maxRows,
        maxCols,
        this.MARGINS.left,
        this.MARGINS.top,
        tableWidth,
        Math.min(tableHeight, maxRows * 35) // 35 points per row
      );
      
      // Style header row
      for (let col = 0; col < maxCols; col++) {
        const headerCell = table.getCell(0, col);
        headerCell.getText().setText(this.truncateText(data.headers[col], 20));
        headerCell.getFill().setSolidFill(colors.primary);
        headerCell.getText().getTextStyle()
          .setFontSize(this.FONTS.data.size)
          .setBold(true)
          .setForegroundColor('#FFFFFF');
      }
      
      // Add data rows with alternating colors
      for (let row = 0; row < maxRows - 1; row++) {
        for (let col = 0; col < maxCols; col++) {
          const cell = table.getCell(row + 1, col);
          const value = data.data[row] && data.data[row][col];
          cell.getText().setText(this.truncateText(String(value || ''), 25));
          
          // Alternate row colors
          if (row % 2 === 0) {
            cell.getFill().setSolidFill(colors.lightGray);
          }
          
          cell.getText().getTextStyle()
            .setFontSize(this.FONTS.data.size)
            .setForegroundColor(colors.text);
        }
      }
      
      // Add data summary below table if space permits
      if (maxRows < 6) {
        const summaryBox = slide.insertTextBox(
          `Showing ${maxRows - 1} of ${data.data.length} records`,
          this.MARGINS.left,
          this.MARGINS.top + (maxRows * 35) + 20,
          300,
          20
        );
        summaryBox.getText().getTextStyle()
          .setFontSize(this.FONTS.caption.size)
          .setForegroundColor(colors.secondary)
          .setItalic(true);
      }
      
      return slide;
    } catch (error) {
      Logger.error('Error creating data table slide:', error);
      return null;
    }
  },
  
  /**
   * Create insights slide with key metrics
   */
  createInsightsSlide: function(presentation, data, analysis, title, theme) {
    try {
      const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
      const colors = this.THEMES[theme] || this.THEMES.professional;
      
      // Add title
      const titleBox = slide.insertTextBox(
        title,
        this.MARGINS.left,
        20,
        this.SLIDE_WIDTH - this.MARGINS.left - this.MARGINS.right,
        40
      );
      titleBox.getText().getTextStyle()
        .setFontSize(this.FONTS.heading.size)
        .setFontFamily(this.FONTS.heading.family)
        .setBold(true)
        .setForegroundColor(colors.text);
      
      // Create metric cards
      const metrics = this.extractKeyMetrics(data, analysis);
      const cardWidth = (this.SLIDE_WIDTH - this.MARGINS.left - this.MARGINS.right - 30) / 3;
      const cardHeight = 100;
      
      for (let i = 0; i < Math.min(6, metrics.length); i++) {
        const row = Math.floor(i / 3);
        const col = i % 3;
        const x = this.MARGINS.left + (col * (cardWidth + 15));
        const y = this.MARGINS.top + 20 + (row * (cardHeight + 15));
        
        // Create card background
        const card = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, x, y, cardWidth, cardHeight);
        card.getFill().setSolidFill(colors.lightGray);
        card.getBorder().getLineFill().setSolidFill(colors.border);
        card.setCornerRadius(8);
        
        // Add metric value
        const valueBox = slide.insertTextBox(
          metrics[i].value,
          x + 10,
          y + 15,
          cardWidth - 20,
          40
        );
        valueBox.getText().getTextStyle()
          .setFontSize(28)
          .setBold(true)
          .setForegroundColor(colors.primary);
        valueBox.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
        
        // Add metric label
        const labelBox = slide.insertTextBox(
          metrics[i].label,
          x + 10,
          y + 55,
          cardWidth - 20,
          30
        );
        labelBox.getText().getTextStyle()
          .setFontSize(this.FONTS.caption.size)
          .setForegroundColor(colors.secondary);
        labelBox.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
      }
      
      return slide;
    } catch (error) {
      Logger.error('Error creating insights slide:', error);
      return null;
    }
  },
  
  /**
   * Extract key metrics from data
   */
  extractKeyMetrics: function(data, analysis) {
    const metrics = [];
    
    // Total records
    metrics.push({
      label: 'Total Records',
      value: String(data.data.length)
    });
    
    // Data completeness
    let filledCells = 0;
    let totalCells = data.data.length * data.headers.length;
    for (let row of data.data) {
      for (let cell of row) {
        if (cell !== '' && cell !== null) filledCells++;
      }
    }
    metrics.push({
      label: 'Data Completeness',
      value: Math.round((filledCells / totalCells) * 100) + '%'
    });
    
    // Numeric column metrics
    if (analysis && analysis.numericColumns.length > 0) {
      for (let numCol of analysis.numericColumns.slice(0, 4)) {
        if (numCol.stats) {
          metrics.push({
            label: `${numCol.name} Avg`,
            value: this.formatNumber(numCol.stats.mean)
          });
        }
      }
    }
    
    return metrics;
  },
  
  /**
   * Format number for display
   */
  formatNumber: function(num) {
    if (num >= 1000000) {
      return (num / 1000000).toFixed(1) + 'M';
    } else if (num >= 1000) {
      return (num / 1000).toFixed(1) + 'K';
    } else if (num < 1 && num > 0) {
      return num.toFixed(2);
    } else {
      return Math.round(num).toString();
    }
  },
  
  /**
   * Truncate text to fit
   */
  truncateText: function(text, maxLength) {
    if (text.length <= maxLength) return text;
    return text.substring(0, maxLength - 3) + '...';
  },
  
  /**
   * Generate intelligent slide deck
   */
  generateProfessionalPresentation: function(presentation, data, config) {
    try {
      const theme = config.template || 'professional';
      const analysis = this.analyzeDataForVisualization(data);
      
      // 1. Title slide
      this.createTitleSlide(presentation, config, theme);
      
      // 2. Key metrics slide
      if (analysis) {
        this.createInsightsSlide(presentation, data, analysis, 'Key Metrics', theme);
      }
      
      // 3. Data table slide
      this.createDataTableSlide(presentation, data, 'Data Overview', theme);
      
      // 4. Visualization slide if applicable
      if (analysis && analysis.recommendations.length > 0) {
        this.createChartSlide(presentation, data, analysis, 'Data Visualization', theme);
      }
      
      // 5. Summary slide
      this.createSummarySlide(presentation, data, analysis, theme);
      
      return true;
    } catch (error) {
      Logger.error('Error generating professional presentation:', error);
      return false;
    }
  },
  
  /**
   * Create summary slide
   */
  createSummarySlide: function(presentation, data, analysis, theme) {
    try {
      const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
      const colors = this.THEMES[theme] || this.THEMES.professional;
      
      // Add title
      const titleBox = slide.insertTextBox(
        'Summary & Next Steps',
        this.MARGINS.left,
        20,
        this.SLIDE_WIDTH - this.MARGINS.left - this.MARGINS.right,
        40
      );
      titleBox.getText().getTextStyle()
        .setFontSize(this.FONTS.heading.size)
        .setFontFamily(this.FONTS.heading.family)
        .setBold(true)
        .setForegroundColor(colors.text);
      
      // Add summary points
      const summaryPoints = [
        `✓ Analyzed ${data.data.length} records across ${data.headers.length} fields`,
        `✓ Data completeness: ${this.calculateCompleteness(data)}%`,
        `✓ ${analysis ? analysis.numericColumns.length : 0} numeric fields identified`,
        `✓ ${analysis ? analysis.categoricalColumns.length : 0} categorical fields identified`,
        `✓ Visualization recommendations: ${analysis ? analysis.recommendations.join(', ') : 'None'}`
      ];
      
      const summaryBox = slide.insertTextBox(
        summaryPoints.join('\n\n'),
        this.MARGINS.left,
        this.MARGINS.top + 20,
        this.SLIDE_WIDTH - this.MARGINS.left - this.MARGINS.right,
        200
      );
      summaryBox.getText().getTextStyle()
        .setFontSize(this.FONTS.body.size)
        .setFontFamily(this.FONTS.body.family)
        .setForegroundColor(colors.text)
        .setLineSpacing(1.5);
      
      // Add call to action
      const ctaBox = slide.insertTextBox(
        'Next Steps: Review insights • Share findings • Take action',
        this.MARGINS.left,
        this.SLIDE_HEIGHT - this.MARGINS.bottom - 60,
        this.SLIDE_WIDTH - this.MARGINS.left - this.MARGINS.right,
        40
      );
      ctaBox.getText().getTextStyle()
        .setFontSize(this.FONTS.subtitle.size)
        .setFontFamily(this.FONTS.subtitle.family)
        .setForegroundColor(colors.primary)
        .setBold(true);
      ctaBox.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
      
      return slide;
    } catch (error) {
      Logger.error('Error creating summary slide:', error);
      return null;
    }
  },
  
  /**
   * Calculate data completeness percentage
   */
  calculateCompleteness: function(data) {
    let filled = 0;
    let total = data.data.length * data.headers.length;
    
    for (let row of data.data) {
      for (let cell of row) {
        if (cell !== '' && cell !== null && cell !== undefined) {
          filled++;
        }
      }
    }
    
    return Math.round((filled / total) * 100);
  }
};