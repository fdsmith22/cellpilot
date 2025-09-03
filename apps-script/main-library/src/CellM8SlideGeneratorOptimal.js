/**
 * CellM8 Optimal Slide Generator
 * Complete rewrite using research-based best practices
 * Avoids all common pitfalls and creates truly clean presentations
 */

const CellM8SlideGeneratorOptimal = {
  
  // Constants
  SLIDE_WIDTH: 720,
  SLIDE_HEIGHT: 405,
  
  // Safe margins to prevent overflow
  SAFE_MARGINS: {
    top: 50,
    bottom: 40,
    left: 50,
    right: 50
  },
  
  // Professional color schemes
  THEMES: {
    professional: {
      primary: '#1a73e8',
      secondary: '#5f6368',
      accent: '#34a853',
      background: '#ffffff',
      surface: '#f8f9fa',
      text: '#202124',
      textLight: '#5f6368',
      error: '#ea4335',
      warning: '#fbbc04'
    },
    dark: {
      primary: '#4285f4',
      secondary: '#aecbfa',
      accent: '#81c995',
      background: '#202124',
      surface: '#303134',
      text: '#e8eaed',
      textLight: '#9aa0a6',
      error: '#f28b82',
      warning: '#fdd663'
    }
  },
  
  /**
   * Main entry point - Create presentation using optimal approach
   */
  createPresentation: function(title, data, config) {
    try {
      Logger.log('Starting optimal presentation creation');
      
      // Step 1: Create new presentation
      const presentation = this.createCleanPresentation(title);
      
      // Step 2: Analyze data intelligently
      const analysis = this.analyzeDataIntelligently(data);
      Logger.log('Data analysis complete:', analysis);
      
      // Step 3: Plan slide structure based on data
      const slidePlan = this.planSlideStructure(analysis, config);
      Logger.log('Slide plan created:', slidePlan);
      
      // Step 4: Build slides according to plan
      this.buildSlidesFromPlan(presentation, slidePlan, data, analysis, config);
      
      // Step 5: Apply finishing touches
      this.applyFinishingTouches(presentation, config);
      
      // Return presentation URL
      const url = presentation.getUrl();
      const slideCount = presentation.getSlides().length;
      
      return {
        success: true,
        url: url,
        presentationId: presentation.getId(),
        slideCount: slideCount,
        message: 'Presentation created successfully using optimal approach'
      };
      
    } catch (error) {
      Logger.error('Error in optimal presentation creation:', error);
      return {
        success: false,
        error: error.toString()
      };
    }
  },
  
  /**
   * Create truly clean presentation using three-request method
   */
  createCleanPresentation: function(title) {
    Logger.log('Creating clean presentation: ' + title);
    
    // Create presentation (will have default first slide)
    const presentation = SlidesApp.create(title);
    
    // Add a new blank slide
    const blankSlide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
    
    // Remove the original slide with placeholders
    const slides = presentation.getSlides();
    if (slides.length > 1) {
      slides[0].remove(); // Remove first slide with default content
      Logger.log('Removed default first slide');
    }
    
    // Clear any remaining content on the blank slide
    this.ensureSlideIsClean(blankSlide);
    
    return presentation;
  },
  
  /**
   * Ensure slide is completely clean
   */
  ensureSlideIsClean: function(slide) {
    const elements = slide.getPageElements();
    
    // Remove elements in reverse to avoid index issues
    for (let i = elements.length - 1; i >= 0; i--) {
      try {
        const element = elements[i];
        
        // Check if it should be removed
        if (this.shouldRemoveElement(element)) {
          element.remove();
          Logger.log('Removed element from slide');
        }
      } catch (e) {
        // Safe to continue if element can't be removed
        Logger.log('Could not remove element: ' + e.toString());
      }
    }
  },
  
  /**
   * Determine if element should be removed
   */
  shouldRemoveElement: function(element) {
    try {
      if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
        const shape = element.asShape();
        
        // Try to detect placeholder
        try {
          const placeholderType = shape.getPlaceholderType();
          // If it has a placeholder type, remove it
          return true;
        } catch (e) {
          // Not a placeholder, check text
          const text = shape.getText().asString().toLowerCase();
          
          // Remove if it contains default text or is empty
          return text.includes('click to add') || 
                 text.includes('title') ||
                 text.includes('subtitle') ||
                 text.includes('body') ||
                 text.trim() === '';
        }
      }
    } catch (e) {
      // If we can't check it, leave it
      return false;
    }
    
    return false;
  },
  
  /**
   * Analyze data intelligently to determine best presentation approach
   */
  analyzeDataIntelligently: function(data) {
    const analysis = {
      totalRows: data.data.length,
      totalColumns: data.headers.length,
      dataTypes: {},
      numericColumns: [],
      textColumns: [],
      dateColumns: [],
      keyMetrics: [],
      bestChartType: null,
      dataSummary: null,
      insights: []
    };
    
    // Analyze each column
    data.headers.forEach((header, colIndex) => {
      const columnData = data.data.map(row => row[colIndex]);
      const dataType = this.detectColumnType(columnData);
      
      analysis.dataTypes[header] = dataType;
      
      if (dataType === 'numeric') {
        const stats = this.calculateStatistics(columnData);
        analysis.numericColumns.push({
          name: header,
          index: colIndex,
          stats: stats
        });
        
        // Add as key metric
        if (stats.average !== null) {
          analysis.keyMetrics.push({
            label: header,
            value: Math.round(stats.average),
            type: 'average'
          });
        }
      } else if (dataType === 'date') {
        analysis.dateColumns.push({
          name: header,
          index: colIndex
        });
      } else {
        analysis.textColumns.push({
          name: header,
          index: colIndex,
          uniqueValues: [...new Set(columnData)].length
        });
      }
    });
    
    // Determine best chart type
    analysis.bestChartType = this.determineBestChartType(analysis);
    
    // Generate data summary
    analysis.dataSummary = this.generateDataSummary(data, analysis);
    
    // Generate insights
    analysis.insights = this.generateDataInsights(data, analysis);
    
    return analysis;
  },
  
  /**
   * Detect column data type
   */
  detectColumnType: function(columnData) {
    const nonEmpty = columnData.filter(val => val != null && val !== '');
    if (nonEmpty.length === 0) return 'empty';
    
    let numericCount = 0;
    let dateCount = 0;
    
    nonEmpty.forEach(val => {
      // Check if numeric
      if (!isNaN(parseFloat(val)) && isFinite(val)) {
        numericCount++;
      }
      // Check if date
      else if (this.isValidDate(val)) {
        dateCount++;
      }
    });
    
    // Determine type based on majority
    const threshold = nonEmpty.length * 0.8;
    
    if (numericCount >= threshold) return 'numeric';
    if (dateCount >= threshold) return 'date';
    return 'text';
  },
  
  /**
   * Calculate statistics for numeric column
   */
  calculateStatistics: function(columnData) {
    const numbers = columnData
      .filter(val => val != null && val !== '')
      .map(val => parseFloat(val))
      .filter(val => !isNaN(val) && isFinite(val));
    
    if (numbers.length === 0) {
      return { average: null, min: null, max: null, sum: null };
    }
    
    return {
      average: numbers.reduce((a, b) => a + b, 0) / numbers.length,
      min: Math.min(...numbers),
      max: Math.max(...numbers),
      sum: numbers.reduce((a, b) => a + b, 0),
      count: numbers.length
    };
  },
  
  /**
   * Check if value is a valid date
   */
  isValidDate: function(value) {
    if (!value) return false;
    const date = new Date(value);
    return date instanceof Date && !isNaN(date) && 
           (value.toString().includes('/') || 
            value.toString().includes('-') ||
            value.toString().match(/\d{4}/));
  },
  
  /**
   * Determine best chart type based on data
   */
  determineBestChartType: function(analysis) {
    // Time series data
    if (analysis.dateColumns.length > 0 && analysis.numericColumns.length > 0) {
      return Charts.ChartType.LINE;
    }
    
    // Categorical with values
    if (analysis.textColumns.length > 0 && analysis.numericColumns.length > 0) {
      if (analysis.totalRows <= 10) {
        return Charts.ChartType.COLUMN;
      } else {
        return Charts.ChartType.BAR;
      }
    }
    
    // Multiple numeric columns
    if (analysis.numericColumns.length >= 2) {
      return Charts.ChartType.SCATTER;
    }
    
    // Single numeric column
    if (analysis.numericColumns.length === 1) {
      return Charts.ChartType.COLUMN;
    }
    
    // Default to table if no numeric data
    return null;
  },
  
  /**
   * Generate data summary
   */
  generateDataSummary: function(data, analysis) {
    const summary = [];
    
    summary.push(`${data.data.length} records with ${data.headers.length} fields`);
    
    if (analysis.numericColumns.length > 0) {
      summary.push(`${analysis.numericColumns.length} numeric columns for analysis`);
    }
    
    if (analysis.dateColumns.length > 0) {
      summary.push('Time-series data available');
    }
    
    const completeness = this.calculateCompleteness(data);
    summary.push(`${completeness}% data completeness`);
    
    return summary.join(' | ');
  },
  
  /**
   * Generate data insights
   */
  generateDataInsights: function(data, analysis) {
    const insights = [];
    
    // Range insights for numeric data
    analysis.numericColumns.forEach(col => {
      if (col.stats.min !== null && col.stats.max !== null) {
        insights.push({
          type: 'range',
          text: `${col.name} ranges from ${Math.round(col.stats.min)} to ${Math.round(col.stats.max)}`
        });
      }
    });
    
    // Unique value insights for text data
    analysis.textColumns.forEach(col => {
      if (col.uniqueValues > 1) {
        insights.push({
          type: 'categories',
          text: `${col.uniqueValues} unique ${col.name} values`
        });
      }
    });
    
    // Data quality insight
    const completeness = this.calculateCompleteness(data);
    if (completeness < 100) {
      insights.push({
        type: 'quality',
        text: `Data is ${completeness}% complete with some missing values`
      });
    }
    
    return insights.slice(0, 5); // Limit to 5 insights
  },
  
  /**
   * Calculate data completeness percentage
   */
  calculateCompleteness: function(data) {
    const totalCells = data.data.length * data.headers.length;
    if (totalCells === 0) return 100;
    
    let filledCells = 0;
    data.data.forEach(row => {
      row.forEach(cell => {
        if (cell != null && cell !== '') {
          filledCells++;
        }
      });
    });
    
    return Math.round((filledCells / totalCells) * 100);
  },
  
  /**
   * Plan slide structure based on analysis
   */
  planSlideStructure: function(analysis, config) {
    const plan = [];
    const requestedSlides = config.slideCount || 5;
    
    // Always start with title slide
    plan.push({
      type: 'title',
      title: config.title || 'Data Presentation',
      subtitle: config.subtitle || analysis.dataSummary
    });
    
    // Add executive summary if we have enough slides
    if (requestedSlides >= 4 && analysis.keyMetrics.length > 0) {
      plan.push({
        type: 'summary',
        metrics: analysis.keyMetrics.slice(0, 4)
      });
    }
    
    // Add data visualization if we have numeric data and chart type
    if (analysis.bestChartType && analysis.numericColumns.length > 0 && requestedSlides >= 3) {
      plan.push({
        type: 'chart',
        chartType: analysis.bestChartType,
        title: 'Data Visualization'
      });
    }
    
    // Add data table for overview (if not too much data)
    if (analysis.totalRows <= 20 && requestedSlides >= plan.length + 2) {
      plan.push({
        type: 'table',
        title: 'Data Overview',
        maxRows: 10,
        maxCols: 6
      });
    }
    
    // Add insights slide if we have insights
    if (analysis.insights.length > 0 && requestedSlides >= plan.length + 2) {
      plan.push({
        type: 'insights',
        title: 'Key Insights',
        insights: analysis.insights
      });
    }
    
    // Always end with conclusion/next steps
    if (requestedSlides > plan.length) {
      plan.push({
        type: 'conclusion',
        title: 'Conclusion',
        nextSteps: this.generateNextSteps(analysis)
      });
    }
    
    // Trim to requested number of slides
    return plan.slice(0, requestedSlides);
  },
  
  /**
   * Generate next steps based on analysis
   */
  generateNextSteps: function(analysis) {
    const steps = [];
    
    if (analysis.numericColumns.length > 1) {
      steps.push('Analyze correlations between numeric variables');
    }
    
    if (analysis.dateColumns.length > 0) {
      steps.push('Perform detailed time-series analysis');
    }
    
    steps.push('Share findings with stakeholders');
    steps.push('Schedule follow-up review');
    
    return steps.slice(0, 3);
  },
  
  /**
   * Build slides from plan
   */
  buildSlidesFromPlan: function(presentation, plan, data, analysis, config) {
    const theme = this.THEMES[config.theme] || this.THEMES.professional;
    
    plan.forEach((slideSpec, index) => {
      Logger.log('Building slide ' + (index + 1) + ': ' + slideSpec.type);
      
      // Get or create slide
      const slides = presentation.getSlides();
      let slide;
      
      if (index < slides.length) {
        slide = slides[index];
        this.ensureSlideIsClean(slide); // Clean existing slide
      } else {
        slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
      }
      
      // Build slide based on type
      switch(slideSpec.type) {
        case 'title':
          this.buildTitleSlide(slide, slideSpec, theme);
          break;
        case 'summary':
          this.buildSummarySlide(slide, slideSpec, theme);
          break;
        case 'chart':
          this.buildChartSlide(slide, slideSpec, data, analysis, theme);
          break;
        case 'table':
          this.buildTableSlide(slide, slideSpec, data, theme);
          break;
        case 'insights':
          this.buildInsightsSlide(slide, slideSpec, theme);
          break;
        case 'conclusion':
          this.buildConclusionSlide(slide, slideSpec, theme);
          break;
        default:
          this.buildDefaultSlide(slide, slideSpec, theme);
      }
    });
    
    // Remove any extra slides
    const slides = presentation.getSlides();
    while (slides.length > plan.length) {
      slides[slides.length - 1].remove();
    }
  },
  
  /**
   * Build title slide
   */
  buildTitleSlide: function(slide, spec, theme) {
    // Add background color
    this.addBackground(slide, theme.primary);
    
    // Add title
    const titleBox = slide.insertTextBox(
      spec.title,
      this.SAFE_MARGINS.left,
      this.SLIDE_HEIGHT * 0.35,
      this.SLIDE_WIDTH - this.SAFE_MARGINS.left - this.SAFE_MARGINS.right,
      60
    );
    
    titleBox.getText().getTextStyle()
      .setFontSize(44)
      .setFontFamily('Arial')
      .setBold(true)
      .setForegroundColor('#FFFFFF');
    
    titleBox.getText().getParagraphStyle()
      .setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    
    // Add subtitle if provided
    if (spec.subtitle) {
      const subtitleBox = slide.insertTextBox(
        spec.subtitle,
        this.SAFE_MARGINS.left,
        this.SLIDE_HEIGHT * 0.5,
        this.SLIDE_WIDTH - this.SAFE_MARGINS.left - this.SAFE_MARGINS.right,
        40
      );
      
      subtitleBox.getText().getTextStyle()
        .setFontSize(18)
        .setFontFamily('Arial')
        .setForegroundColor('#FFFFFF');
      
      subtitleBox.getText().getParagraphStyle()
        .setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    }
    
    // Add decorative element
    const decorLine = slide.insertShape(
      SlidesApp.ShapeType.RECTANGLE,
      this.SLIDE_WIDTH * 0.25,
      this.SLIDE_HEIGHT * 0.6,
      this.SLIDE_WIDTH * 0.5,
      2
    );
    decorLine.getFill().setSolidFill('#FFFFFF');
    decorLine.getBorder().setTransparent();
  },
  
  /**
   * Build summary slide with metrics
   */
  buildSummarySlide: function(slide, spec, theme) {
    // White background
    this.addBackground(slide, theme.background);
    
    // Add title
    this.addSlideTitle(slide, 'Executive Summary', theme);
    
    // Add metric cards
    const metrics = spec.metrics || [];
    const cardWidth = 140;
    const cardHeight = 80;
    const spacing = 20;
    const startX = (this.SLIDE_WIDTH - (metrics.length * cardWidth + (metrics.length - 1) * spacing)) / 2;
    const startY = 120;
    
    metrics.forEach((metric, index) => {
      const x = startX + index * (cardWidth + spacing);
      
      // Card background
      const card = slide.insertShape(
        SlidesApp.ShapeType.RECTANGLE,
        x, startY, cardWidth, cardHeight
      );
      card.getFill().setSolidFill(theme.surface);
      card.getBorder().getLineFill().setSolidFill(theme.primary);
      card.getBorder().setWeight(1);
      
      // Metric value
      const valueBox = slide.insertTextBox(
        metric.value.toString(),
        x, startY + 15, cardWidth, 30
      );
      valueBox.getText().getTextStyle()
        .setFontSize(24)
        .setFontFamily('Arial')
        .setBold(true)
        .setForegroundColor(theme.primary);
      valueBox.getText().getParagraphStyle()
        .setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
      
      // Metric label
      const labelBox = slide.insertTextBox(
        metric.label,
        x, startY + 45, cardWidth, 20
      );
      labelBox.getText().getTextStyle()
        .setFontSize(12)
        .setFontFamily('Arial')
        .setForegroundColor(theme.textLight);
      labelBox.getText().getParagraphStyle()
        .setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    });
  },
  
  /**
   * Build chart slide
   */
  buildChartSlide: function(slide, spec, data, analysis, theme) {
    // White background
    this.addBackground(slide, theme.background);
    
    // Add title
    this.addSlideTitle(slide, spec.title || 'Data Visualization', theme);
    
    try {
      // Prepare data for chart
      const chartData = this.prepareChartData(data, analysis, spec.chartType);
      
      // Build chart
      const chart = this.buildChart(chartData, spec.chartType, theme);
      
      // Insert chart
      slide.insertChart(
        chart,
        this.SAFE_MARGINS.left,
        100,
        this.SLIDE_WIDTH - this.SAFE_MARGINS.left - this.SAFE_MARGINS.right,
        250
      );
      
    } catch (error) {
      Logger.error('Chart creation failed:', error);
      // Add error message
      this.addErrorMessage(slide, 'Unable to create chart from data', theme);
    }
  },
  
  /**
   * Build table slide
   */
  buildTableSlide: function(slide, spec, data, theme) {
    // White background
    this.addBackground(slide, theme.background);
    
    // Add title
    this.addSlideTitle(slide, spec.title || 'Data Overview', theme);
    
    // Determine table size
    const maxRows = spec.maxRows || 8;
    const maxCols = spec.maxCols || 5;
    const actualRows = Math.min(data.data.length, maxRows);
    const actualCols = Math.min(data.headers.length, maxCols);
    
    // Create table
    const tableWidth = this.SLIDE_WIDTH - this.SAFE_MARGINS.left - this.SAFE_MARGINS.right;
    const table = slide.insertTable(
      actualRows + 1, // +1 for header
      actualCols,
      this.SAFE_MARGINS.left,
      100,
      tableWidth,
      Math.min(250, (actualRows + 1) * 30)
    );
    
    // Style header row
    for (let col = 0; col < actualCols; col++) {
      const cell = table.getCell(0, col);
      cell.getFill().setSolidFill(theme.primary);
      cell.getText().setText(data.headers[col]);
      cell.getText().getTextStyle()
        .setFontSize(11)
        .setFontFamily('Arial')
        .setBold(true)
        .setForegroundColor('#FFFFFF');
    }
    
    // Add data rows
    for (let row = 0; row < actualRows; row++) {
      for (let col = 0; col < actualCols; col++) {
        const cell = table.getCell(row + 1, col);
        
        // Alternate row colors
        if (row % 2 === 1) {
          cell.getFill().setSolidFill(theme.surface);
        }
        
        // Set cell value
        const value = data.data[row][col];
        cell.getText().setText(this.formatCellValue(value));
        cell.getText().getTextStyle()
          .setFontSize(10)
          .setFontFamily('Arial')
          .setForegroundColor(theme.text);
      }
    }
    
    // Add note if truncated
    if (data.data.length > actualRows || data.headers.length > actualCols) {
      const noteBox = slide.insertTextBox(
        `Showing ${actualRows} of ${data.data.length} rows, ${actualCols} of ${data.headers.length} columns`,
        this.SAFE_MARGINS.left,
        360,
        tableWidth,
        20
      );
      noteBox.getText().getTextStyle()
        .setFontSize(10)
        .setFontFamily('Arial')
        .setForegroundColor(theme.textLight)
        .setItalic(true);
    }
  },
  
  /**
   * Build insights slide
   */
  buildInsightsSlide: function(slide, spec, theme) {
    // White background
    this.addBackground(slide, theme.background);
    
    // Add title
    this.addSlideTitle(slide, spec.title || 'Key Insights', theme);
    
    // Add insights
    const insights = spec.insights || [];
    const startY = 100;
    
    insights.forEach((insight, index) => {
      const y = startY + index * 50;
      
      // Bullet point
      const bullet = slide.insertShape(
        SlidesApp.ShapeType.ELLIPSE,
        this.SAFE_MARGINS.left,
        y + 5,
        8, 8
      );
      bullet.getFill().setSolidFill(theme.accent);
      bullet.getBorder().setTransparent();
      
      // Insight text
      const textBox = slide.insertTextBox(
        insight.text,
        this.SAFE_MARGINS.left + 20,
        y,
        this.SLIDE_WIDTH - this.SAFE_MARGINS.left - this.SAFE_MARGINS.right - 20,
        30
      );
      textBox.getText().getTextStyle()
        .setFontSize(14)
        .setFontFamily('Arial')
        .setForegroundColor(theme.text);
    });
  },
  
  /**
   * Build conclusion slide
   */
  buildConclusionSlide: function(slide, spec, theme) {
    // Primary color background
    this.addBackground(slide, theme.primary);
    
    // Add title
    const titleBox = slide.insertTextBox(
      spec.title || 'Next Steps',
      this.SAFE_MARGINS.left,
      this.SAFE_MARGINS.top,
      this.SLIDE_WIDTH - this.SAFE_MARGINS.left - this.SAFE_MARGINS.right,
      50
    );
    titleBox.getText().getTextStyle()
      .setFontSize(32)
      .setFontFamily('Arial')
      .setBold(true)
      .setForegroundColor('#FFFFFF');
    titleBox.getText().getParagraphStyle()
      .setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    
    // Add next steps
    const steps = spec.nextSteps || [];
    const startY = 140;
    
    steps.forEach((step, index) => {
      const y = startY + index * 60;
      
      // Step number
      const numberBox = slide.insertTextBox(
        (index + 1).toString(),
        this.SAFE_MARGINS.left + 100,
        y,
        30, 30
      );
      numberBox.getText().getTextStyle()
        .setFontSize(20)
        .setFontFamily('Arial')
        .setBold(true)
        .setForegroundColor('#FFFFFF');
      
      // Step text
      const stepBox = slide.insertTextBox(
        step,
        this.SAFE_MARGINS.left + 140,
        y + 5,
        400, 30
      );
      stepBox.getText().getTextStyle()
        .setFontSize(16)
        .setFontFamily('Arial')
        .setForegroundColor('#FFFFFF');
    });
  },
  
  /**
   * Build default slide (fallback)
   */
  buildDefaultSlide: function(slide, spec, theme) {
    this.addBackground(slide, theme.background);
    this.addSlideTitle(slide, spec.title || 'Slide', theme);
  },
  
  /**
   * Add background to slide
   */
  addBackground: function(slide, color) {
    const bg = slide.insertShape(
      SlidesApp.ShapeType.RECTANGLE,
      0, 0,
      this.SLIDE_WIDTH, this.SLIDE_HEIGHT
    );
    bg.getFill().setSolidFill(color);
    bg.getBorder().setTransparent();
    bg.sendToBack();
  },
  
  /**
   * Add title to slide
   */
  addSlideTitle: function(slide, title, theme) {
    const titleBox = slide.insertTextBox(
      title,
      this.SAFE_MARGINS.left,
      this.SAFE_MARGINS.top,
      this.SLIDE_WIDTH - this.SAFE_MARGINS.left - this.SAFE_MARGINS.right,
      40
    );
    titleBox.getText().getTextStyle()
      .setFontSize(24)
      .setFontFamily('Arial')
      .setBold(true)
      .setForegroundColor(theme.primary);
    
    // Add underline
    const underline = slide.insertShape(
      SlidesApp.ShapeType.RECTANGLE,
      this.SAFE_MARGINS.left,
      this.SAFE_MARGINS.top + 35,
      100, 3
    );
    underline.getFill().setSolidFill(theme.accent);
    underline.getBorder().setTransparent();
  },
  
  /**
   * Add error message to slide
   */
  addErrorMessage: function(slide, message, theme) {
    const errorBox = slide.insertTextBox(
      message,
      this.SAFE_MARGINS.left,
      200,
      this.SLIDE_WIDTH - this.SAFE_MARGINS.left - this.SAFE_MARGINS.right,
      30
    );
    errorBox.getText().getTextStyle()
      .setFontSize(14)
      .setFontFamily('Arial')
      .setForegroundColor(theme.error)
      .setItalic(true);
    errorBox.getText().getParagraphStyle()
      .setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
  },
  
  /**
   * Prepare data for chart
   */
  prepareChartData: function(data, analysis, chartType) {
    // Limit data points for clarity
    const maxPoints = 15;
    
    // Select appropriate columns
    let xColumn = 0;
    let yColumn = 1;
    
    // If we have date column, use it as X
    if (analysis.dateColumns.length > 0) {
      xColumn = analysis.dateColumns[0].index;
    } else if (analysis.textColumns.length > 0) {
      xColumn = analysis.textColumns[0].index;
    }
    
    // Use first numeric column as Y
    if (analysis.numericColumns.length > 0) {
      yColumn = analysis.numericColumns[0].index;
    }
    
    // Prepare limited data
    const limitedData = {
      headers: [data.headers[xColumn], data.headers[yColumn]],
      data: data.data.slice(0, maxPoints).map(row => [
        row[xColumn],
        row[yColumn]
      ])
    };
    
    return limitedData;
  },
  
  /**
   * Build chart
   */
  buildChart: function(chartData, chartType, theme) {
    const dataTable = Charts.newDataTable();
    
    // Add columns
    chartData.headers.forEach((header, index) => {
      // First column is usually labels
      if (index === 0) {
        dataTable.addColumn(Charts.ColumnType.STRING, header);
      } else {
        dataTable.addColumn(Charts.ColumnType.NUMBER, header);
      }
    });
    
    // Add rows
    chartData.data.forEach(row => {
      const rowData = row.map((val, index) => {
        if (index === 0) return val ? val.toString() : '';
        return parseFloat(val) || 0;
      });
      dataTable.addRow(rowData);
    });
    
    // Build appropriate chart
    let chartBuilder;
    
    switch(chartType) {
      case Charts.ChartType.LINE:
        chartBuilder = Charts.newLineChart();
        break;
      case Charts.ChartType.BAR:
        chartBuilder = Charts.newBarChart();
        break;
      case Charts.ChartType.SCATTER:
        chartBuilder = Charts.newScatterChart();
        break;
      default:
        chartBuilder = Charts.newColumnChart();
    }
    
    // Configure chart
    chartBuilder
      .setDataTable(dataTable.build())
      .setOption('backgroundColor', theme.background)
      .setOption('colors', [theme.primary, theme.accent])
      .setOption('legend', { position: 'bottom' })
      .setDimensions(600, 400);
    
    return chartBuilder.build();
  },
  
  /**
   * Format cell value for display
   */
  formatCellValue: function(value) {
    if (value == null || value === '') return '';
    
    const str = value.toString();
    
    // Truncate long strings
    if (str.length > 20) {
      return str.substring(0, 17) + '...';
    }
    
    // Format numbers
    if (!isNaN(parseFloat(value)) && isFinite(value)) {
      const num = parseFloat(value);
      if (num % 1 === 0) {
        return num.toString();
      }
      return num.toFixed(2);
    }
    
    return str;
  },
  
  /**
   * Apply finishing touches
   */
  applyFinishingTouches: function(presentation, config) {
    const slides = presentation.getSlides();
    const theme = this.THEMES[config.theme] || this.THEMES.professional;
    
    slides.forEach((slide, index) => {
      // Add slide numbers (except first and last)
      if (index > 0 && index < slides.length - 1) {
        const pageNum = slide.insertTextBox(
          (index + 1).toString(),
          this.SLIDE_WIDTH - 40,
          this.SLIDE_HEIGHT - 30,
          30, 20
        );
        pageNum.getText().getTextStyle()
          .setFontSize(10)
          .setFontFamily('Arial')
          .setForegroundColor(theme.textLight);
        pageNum.getText().getParagraphStyle()
          .setParagraphAlignment(SlidesApp.ParagraphAlignment.END);
      }
    });
  }
};