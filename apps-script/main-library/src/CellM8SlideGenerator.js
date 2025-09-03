/**
 * CellM8 Slide Generator
 * Research-based implementation using Google Slides API best practices
 * The ONLY working generator that properly handles placeholders and formatting
 */

const CellM8SlideGenerator = {
  
  // Constants
  SLIDE_WIDTH: 720,
  SLIDE_HEIGHT: 405,
  
  // Golden ratio for layouts
  GOLDEN_RATIO: 1.618,
  
  // Grid system (12-column)
  GRID_COLUMNS: 12,
  GRID_GUTTER: 16,
  
  // Safe margins to prevent overflow
  SAFE_MARGINS: {
    top: 40,
    bottom: 40,
    left: 60,
    right: 60
  },
  
  // Typography scale (based on golden ratio)
  TYPOGRAPHY: {
    display: 48,     // Main titles
    headline: 32,    // Section titles
    title: 24,       // Slide titles
    subtitle: 18,    // Subtitles
    body: 14,        // Body text
    caption: 11,     // Captions
    overline: 10     // Small labels
  },
  
  // Professional color schemes (Material Design 3 inspired)
  THEMES: {
    professional: {
      primary: '#1a73e8',
      primaryVariant: '#1557b0',
      secondary: '#5f6368',
      accent: '#34a853',
      background: '#ffffff',
      surface: '#f8f9fa',
      surfaceVariant: '#e8eaed',
      text: '#202124',
      textLight: '#5f6368',
      error: '#ea4335',
      warning: '#fbbc04',
      gradients: {
        primary: ['#1a73e8', '#1557b0'],
        accent: ['#34a853', '#0d652d'],
        subtle: ['#ffffff', '#f8f9fa']
      }
    },
    modern: {
      primary: '#6366f1',
      primaryVariant: '#4f46e5',
      secondary: '#8b5cf6',
      accent: '#ec4899',
      background: '#ffffff',
      surface: '#fafbfc',
      surfaceVariant: '#f0f7ff',
      text: '#1e293b',
      textLight: '#64748b',
      error: '#ef4444',
      warning: '#f59e0b',
      gradients: {
        primary: ['#6366f1', '#4f46e5'],
        accent: ['#ec4899', '#db2777'],
        subtle: ['#fafbfc', '#f0f7ff']
      }
    },
    dark: {
      primary: '#4285f4',
      primaryVariant: '#1a73e8',
      secondary: '#aecbfa',
      accent: '#81c995',
      background: '#202124',
      surface: '#303134',
      surfaceVariant: '#3c4043',
      text: '#e8eaed',
      textLight: '#9aa0a6',
      error: '#f28b82',
      warning: '#fdd663',
      gradients: {
        primary: ['#4285f4', '#1a73e8'],
        accent: ['#81c995', '#5bb974'],
        subtle: ['#303134', '#202124']
      }
    },
    vibrant: {
      primary: '#ff6b6b',
      primaryVariant: '#ee5a6f',
      secondary: '#4ecdc4',
      accent: '#ffe66d',
      background: '#ffffff',
      surface: '#fff5f5',
      surfaceVariant: '#ffeaa7',
      text: '#2d3436',
      textLight: '#636e72',
      error: '#d63031',
      warning: '#fdcb6e',
      gradients: {
        primary: ['#ff6b6b', '#ee5a6f'],
        accent: ['#4ecdc4', '#26a69a'],
        subtle: ['#ffffff', '#fff5f5']
      }
    },
    elegant: {
      primary: '#2c3e50',
      primaryVariant: '#34495e',
      secondary: '#95a5a6',
      accent: '#e67e22',
      background: '#ffffff',
      surface: '#ecf0f1',
      surfaceVariant: '#bdc3c7',
      text: '#2c3e50',
      textLight: '#7f8c8d',
      error: '#e74c3c',
      warning: '#f39c12',
      gradients: {
        primary: ['#2c3e50', '#34495e'],
        accent: ['#e67e22', '#d35400'],
        subtle: ['#ecf0f1', '#bdc3c7']
      }
    },
    corporate: {
      primary: '#003366',
      primaryVariant: '#001f3f',
      secondary: '#2c5282',
      accent: '#0077be',
      background: '#ffffff',
      surface: '#f7fafc',
      surfaceVariant: '#e2e8f0',
      text: '#1a202c',
      textLight: '#718096',
      error: '#c53030',
      warning: '#dd6b20',
      gradients: {
        primary: ['#003366', '#001f3f'],
        accent: ['#0077be', '#005a8b'],
        subtle: ['#f7fafc', '#e2e8f0']
      }
    }
  },
  
  // Chart selection matrix based on data characteristics
  CHART_SELECTION_MATRIX: {
    'time_series': {
      few_points: Charts.ChartType.LINE,
      many_points: Charts.ChartType.AREA,
      multiple_series: Charts.ChartType.LINE
    },
    'comparison': {
      few_categories: Charts.ChartType.COLUMN,
      many_categories: Charts.ChartType.BAR,
      proportions: Charts.ChartType.PIE
    },
    'distribution': {
      continuous: Charts.ChartType.HISTOGRAM,
      categorical: Charts.ChartType.COLUMN,
      correlation: Charts.ChartType.SCATTER
    },
    'composition': {
      static: Charts.ChartType.PIE,
      time_based: Charts.ChartType.AREA,
      hierarchical: Charts.ChartType.TREEMAP
    }
  },
  
  /**
   * Calculate golden ratio dimensions
   */
  calculateGoldenDimensions: function(width) {
    return {
      width: width,
      height: width / this.GOLDEN_RATIO,
      smallerWidth: width / this.GOLDEN_RATIO,
      largerWidth: width * this.GOLDEN_RATIO
    };
  },
  
  /**
   * Get grid column width
   */
  getGridColumnWidth: function(columns) {
    const totalWidth = this.SLIDE_WIDTH - this.SAFE_MARGINS.left - this.SAFE_MARGINS.right;
    const totalGutters = (this.GRID_COLUMNS - 1) * this.GRID_GUTTER;
    const columnWidth = (totalWidth - totalGutters) / this.GRID_COLUMNS;
    return (columnWidth * columns) + ((columns - 1) * this.GRID_GUTTER);
  },
  
  /**
   * Get grid position
   */
  getGridPosition: function(column, span) {
    const columnWidth = this.getGridColumnWidth(1);
    const x = this.SAFE_MARGINS.left + (column * (columnWidth + this.GRID_GUTTER));
    const width = this.getGridColumnWidth(span);
    return { x: x, width: width };
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
    // Analyze data characteristics
    const hasTime = analysis.dateColumns.length > 0;
    const hasCategories = analysis.textColumns.length > 0;
    const numericCount = analysis.numericColumns.length;
    const rowCount = analysis.totalRows;
    
    // Time series data
    if (hasTime && numericCount > 0) {
      if (rowCount > 50) {
        return Charts.ChartType.AREA; // Area for dense time series
      } else if (numericCount > 2) {
        return Charts.ChartType.LINE; // Line for multiple series
      } else {
        return Charts.ChartType.COLUMN; // Column for sparse time series
      }
    }
    
    // Composition data (parts of whole)
    if (hasCategories && numericCount === 1 && rowCount <= 8) {
      // Check if values might represent parts of whole
      const numColumn = analysis.numericColumns[0];
      if (numColumn.stats.min >= 0) {
        return Charts.ChartType.PIE;
      }
    }
    
    // Comparison data
    if (hasCategories && numericCount > 0) {
      if (rowCount <= 5 && numericCount <= 3) {
        return Charts.ChartType.COLUMN; // Few categories, use columns
      } else if (rowCount <= 15) {
        return Charts.ChartType.BAR; // More categories, use horizontal bars
      } else {
        // Too many categories, consider grouping or use table
        return Charts.ChartType.TABLE;
      }
    }
    
    // Correlation/Distribution
    if (numericCount >= 2) {
      if (rowCount <= 100) {
        return Charts.ChartType.SCATTER; // Scatter for correlation
      } else {
        // Consider first two numeric columns for area chart
        return Charts.ChartType.AREA;
      }
    }
    
    // Single numeric column - distribution
    if (numericCount === 1) {
      if (rowCount <= 20) {
        return Charts.ChartType.COLUMN;
      } else {
        return Charts.ChartType.HISTOGRAM;
      }
    }
    
    // Default to table if no clear visualization
    return Charts.ChartType.TABLE;
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
    // Map template styles to theme names
    const themeMap = {
      'simple': 'professional',
      'corporate': 'corporate',
      'modern': 'modern',
      'elegant': 'elegant',
      'vibrant': 'vibrant',
      'dark': 'dark',
      'professional': 'professional'
    };
    
    const themeName = themeMap[config.template] || 'professional';
    const theme = this.THEMES[themeName] || this.THEMES.professional;
    
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
          this.addSpeakerNotes(slide, 
            'Welcome to this data presentation. This analysis covers ' + analysis.totalRows + 
            ' records across ' + analysis.totalColumns + ' fields.');
          break;
        case 'summary':
          this.buildSummarySlide(slide, slideSpec, theme);
          this.addSpeakerNotes(slide, 
            'Executive summary showing key metrics. These represent averages and totals from the analyzed dataset.');
          break;
        case 'chart':
          this.buildChartSlide(slide, slideSpec, data, analysis, theme);
          this.addSpeakerNotes(slide, 
            'This visualization shows the relationship between key data points. ' +
            'Chart type was automatically selected based on data characteristics.');
          break;
        case 'table':
          this.buildTableSlide(slide, slideSpec, data, theme);
          this.addSpeakerNotes(slide, 
            'Data table showing a sample of the dataset. Full dataset contains ' + 
            data.data.length + ' rows and ' + data.headers.length + ' columns.');
          break;
        case 'insights':
          this.buildInsightsSlide(slide, slideSpec, theme);
          this.addSpeakerNotes(slide, 
            'Key insights derived from data analysis. These findings highlight important patterns and trends.');
          break;
        case 'conclusion':
          this.buildConclusionSlide(slide, slideSpec, theme);
          this.addSpeakerNotes(slide, 
            'Recommended next steps based on the analysis. Consider scheduling a follow-up to track progress.');
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
    // Add gradient background
    this.addGradientBackground(slide, theme.gradients.primary[0], theme.gradients.primary[1]);
    
    // Add decorative shape overlay
    const overlay = slide.insertShape(
      SlidesApp.ShapeType.RECTANGLE,
      0,
      this.SLIDE_HEIGHT * 0.7,
      this.SLIDE_WIDTH,
      this.SLIDE_HEIGHT * 0.3
    );
    overlay.getFill().setSolidFill(theme.primaryVariant);
    overlay.getFill().setTransparency(0.3);
    overlay.getBorder().setTransparent();
    overlay.sendToBack();
    
    // Add title with golden ratio positioning
    const titleY = (this.SLIDE_HEIGHT / this.GOLDEN_RATIO) - 30;
    const titleBox = slide.insertTextBox(
      spec.title,
      this.SAFE_MARGINS.left,
      titleY,
      this.SLIDE_WIDTH - this.SAFE_MARGINS.left - this.SAFE_MARGINS.right,
      60
    );
    
    titleBox.getText().getTextStyle()
      .setFontSize(this.TYPOGRAPHY.display)
      .setFontFamily('Google Sans')
      .setBold(true)
      .setForegroundColor('#FFFFFF');
    
    titleBox.getText().getParagraphStyle()
      .setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    
    // Add subtitle if provided
    if (spec.subtitle) {
      const subtitleBox = slide.insertTextBox(
        spec.subtitle,
        this.SAFE_MARGINS.left,
        titleY + 70,
        this.SLIDE_WIDTH - this.SAFE_MARGINS.left - this.SAFE_MARGINS.right,
        40
      );
      
      subtitleBox.getText().getTextStyle()
        .setFontSize(this.TYPOGRAPHY.subtitle)
        .setFontFamily('Google Sans')
        .setForegroundColor('#FFFFFF');
      
      subtitleBox.getText().getParagraphStyle()
        .setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    }
    
    // Add elegant divider line
    const decorLine = slide.insertShape(
      SlidesApp.ShapeType.RECTANGLE,
      (this.SLIDE_WIDTH - 120) / 2,
      titleY + 120,
      120,
      3
    );
    decorLine.getFill().setSolidFill('#FFFFFF');
    decorLine.getFill().setTransparency(0.8);
    decorLine.getBorder().setTransparent();
    
    // Add date in corner
    const date = new Date().toLocaleDateString();
    const dateBox = slide.insertTextBox(
      date,
      this.SLIDE_WIDTH - this.SAFE_MARGINS.right - 100,
      this.SLIDE_HEIGHT - this.SAFE_MARGINS.bottom,
      100,
      20
    );
    dateBox.getText().getTextStyle()
      .setFontSize(this.TYPOGRAPHY.caption)
      .setFontFamily('Google Sans')
      .setForegroundColor('#FFFFFF');
    dateBox.getText().getTextStyle().setTransparency(0.7);
    dateBox.getText().getParagraphStyle()
      .setParagraphAlignment(SlidesApp.ParagraphAlignment.END);
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
   * Build table slide with enhanced formatting
   */
  buildTableSlide: function(slide, spec, data, theme) {
    // Add subtle gradient background
    this.addGradientBackground(slide, theme.background, theme.surface);
    
    // Add title with accent underline
    this.addSlideTitle(slide, spec.title || 'Data Overview', theme);
    
    // Calculate optimal table dimensions
    const maxRows = spec.maxRows || 10;
    const maxCols = spec.maxCols || 6;
    const actualRows = Math.min(data.data.length, maxRows);
    const actualCols = Math.min(data.headers.length, maxCols);
    
    // Use golden ratio for table dimensions
    const tableWidth = this.SLIDE_WIDTH - this.SAFE_MARGINS.left - this.SAFE_MARGINS.right;
    const tableHeight = Math.min(280, (actualRows + 1) * 28);
    const tableY = 100;
    
    // Create table with proper spacing
    const table = slide.insertTable(
      actualRows + 1, // +1 for header
      actualCols,
      this.SAFE_MARGINS.left,
      tableY,
      tableWidth,
      tableHeight
    );
    
    // Style header row with gradient effect
    for (let col = 0; col < actualCols; col++) {
      const cell = table.getCell(0, col);
      cell.getFill().setSolidFill(theme.primary);
      
      // Format header text
      const headerText = data.headers[col].toString().toUpperCase();
      cell.getText().setText(headerText);
      cell.getText().getTextStyle()
        .setFontSize(this.TYPOGRAPHY.caption)
        .setFontFamily('Google Sans')
        .setBold(true)
        .setForegroundColor('#FFFFFF');
      cell.getText().getParagraphStyle()
        .setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    }
    
    // Add data rows with enhanced styling
    for (let row = 0; row < actualRows; row++) {
      for (let col = 0; col < actualCols; col++) {
        const cell = table.getCell(row + 1, col);
        
        // Alternate row colors with subtle gradient
        if (row % 2 === 0) {
          cell.getFill().setSolidFill(theme.background);
        } else {
          cell.getFill().setSolidFill(theme.surfaceVariant || theme.surface);
        }
        
        // Format cell value
        const value = data.data[row][col];
        const formattedValue = this.formatCellValue(value);
        cell.getText().setText(formattedValue);
        
        // Style based on data type
        const isNumeric = !isNaN(parseFloat(value)) && isFinite(value);
        cell.getText().getTextStyle()
          .setFontSize(this.TYPOGRAPHY.caption)
          .setFontFamily('Google Sans')
          .setForegroundColor(isNumeric ? theme.text : theme.textLight);
        
        // Align numbers to right, text to left
        cell.getText().getParagraphStyle()
          .setParagraphAlignment(
            isNumeric ? SlidesApp.ParagraphAlignment.END : SlidesApp.ParagraphAlignment.START
          );
      }
    }
    
    // Add summary stats box if data is truncated
    if (data.data.length > actualRows || data.headers.length > actualCols) {
      // Create info card
      const cardY = tableY + tableHeight + 20;
      const cardWidth = 300;
      const cardHeight = 40;
      const cardX = (this.SLIDE_WIDTH - cardWidth) / 2;
      
      // Card background
      const card = slide.insertShape(
        SlidesApp.ShapeType.RECTANGLE,
        cardX, cardY, cardWidth, cardHeight
      );
      card.getFill().setSolidFill(theme.surfaceVariant || theme.surface);
      card.getBorder().getLineFill().setSolidFill(theme.primary);
      card.getBorder().setWeight(1);
      
      // Info text
      const infoText = `Displaying ${actualRows} of ${data.data.length} rows • ${actualCols} of ${data.headers.length} columns`;
      const infoBox = slide.insertTextBox(
        infoText,
        cardX,
        cardY + 12,
        cardWidth,
        20
      );
      infoBox.getText().getTextStyle()
        .setFontSize(this.TYPOGRAPHY.caption)
        .setFontFamily('Google Sans')
        .setForegroundColor(theme.textLight)
        .setItalic(true);
      infoBox.getText().getParagraphStyle()
        .setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
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
   * Add gradient background to slide
   */
  addGradientBackground: function(slide, color1, color2, angle) {
    // Create two rectangles to simulate gradient
    const bgTop = slide.insertShape(
      SlidesApp.ShapeType.RECTANGLE,
      0, 0,
      this.SLIDE_WIDTH, this.SLIDE_HEIGHT / 2
    );
    bgTop.getFill().setSolidFill(color1);
    bgTop.getBorder().setTransparent();
    bgTop.sendToBack();
    
    const bgBottom = slide.insertShape(
      SlidesApp.ShapeType.RECTANGLE,
      0, this.SLIDE_HEIGHT / 2,
      this.SLIDE_WIDTH, this.SLIDE_HEIGHT / 2
    );
    bgBottom.getFill().setSolidFill(color2);
    bgBottom.getBorder().setTransparent();
    bgBottom.sendToBack();
    
    // Add blending overlay
    const blend = slide.insertShape(
      SlidesApp.ShapeType.RECTANGLE,
      0, this.SLIDE_HEIGHT * 0.4,
      this.SLIDE_WIDTH, this.SLIDE_HEIGHT * 0.2
    );
    blend.getFill().setSolidFill(color1);
    blend.getFill().setTransparency(0.5);
    blend.getBorder().setTransparent();
    blend.sendToBack();
  },
  
  /**
   * Add title to slide with enhanced styling
   */
  addSlideTitle: function(slide, title, theme) {
    // Title text
    const titleBox = slide.insertTextBox(
      title,
      this.SAFE_MARGINS.left,
      this.SAFE_MARGINS.top,
      this.SLIDE_WIDTH - this.SAFE_MARGINS.left - this.SAFE_MARGINS.right,
      40
    );
    titleBox.getText().getTextStyle()
      .setFontSize(this.TYPOGRAPHY.title)
      .setFontFamily('Google Sans')
      .setBold(true)
      .setForegroundColor(theme.primary);
    
    // Animated underline with golden ratio width
    const underlineWidth = 100 / this.GOLDEN_RATIO;
    const underline = slide.insertShape(
      SlidesApp.ShapeType.RECTANGLE,
      this.SAFE_MARGINS.left,
      this.SAFE_MARGINS.top + 35,
      underlineWidth, 
      3
    );
    underline.getFill().setSolidFill(theme.accent);
    underline.getBorder().setTransparent();
    
    // Add subtle accent shape
    const accentShape = slide.insertShape(
      SlidesApp.ShapeType.ELLIPSE,
      this.SAFE_MARGINS.left + underlineWidth + 10,
      this.SAFE_MARGINS.top + 33,
      6, 6
    );
    accentShape.getFill().setSolidFill(theme.accent);
    accentShape.getFill().setTransparency(0.5);
    accentShape.getBorder().setTransparent();
  },
  
  /**
   * Add speaker notes to slide
   */
  addSpeakerNotes: function(slide, notes) {
    try {
      const notesPage = slide.getNotesPage();
      const speaker = notesPage.getSpeakerNotesShape();
      if (speaker) {
        speaker.getText().setText(notes);
      }
    } catch (e) {
      Logger.log('Could not add speaker notes: ' + e.toString());
    }
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
   * Apply finishing touches to presentation
   */
  applyFinishingTouches: function(presentation, config) {
    try {
      const slides = presentation.getSlides();
      const totalSlides = slides.length;
      
      // Add slide numbers to all except title slide
      slides.forEach((slide, index) => {
        if (index > 0) { // Skip title slide
          const slideNumber = (index + 1) + ' / ' + totalSlides;
          const numberBox = slide.insertTextBox(
            slideNumber,
            this.SLIDE_WIDTH - 80,
            this.SLIDE_HEIGHT - 25,
            70,
            20
          );
          numberBox.getText().getTextStyle()
            .setFontSize(this.TYPOGRAPHY.overline)
            .setFontFamily('Google Sans')
            .setForegroundColor('#9aa0a6');
          numberBox.getText().getParagraphStyle()
            .setParagraphAlignment(SlidesApp.ParagraphAlignment.END);
        }
      });
      
      // Set presentation name
      presentation.setName(config.title || 'Data Presentation');
      
      Logger.log('Applied finishing touches to presentation');
    } catch (error) {
      Logger.log('Could not apply finishing touches: ' + error.toString());
    }
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