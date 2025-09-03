/**
 * CellM8 Professional Slide Generator
 * Complete redesign with proper slide management and beautiful templates
 */

const CellM8SlideGeneratorPro = {
  
  // Slide dimensions (Google Slides uses points: 1 point = 1/72 inch)
  DIMENSIONS: {
    width: 720,
    height: 405
  },
  
  // Professional margins
  MARGINS: {
    top: 40,
    bottom: 40,
    left: 50,
    right: 50,
    contentTop: 80  // Below title
  },
  
  // Beautiful, cohesive color schemes
  THEMES: {
    corporate: {
      name: 'Corporate Blue',
      background: '#FFFFFF',
      titleSlideGradient: ['#1E3A8A', '#2563EB'],
      primary: '#1E40AF',
      secondary: '#3B82F6', 
      accent: '#60A5FA',
      text: '#1F2937',
      textLight: '#6B7280',
      textOnDark: '#FFFFFF',
      surface: '#F3F4F6',
      border: '#E5E7EB',
      success: '#10B981',
      warning: '#F59E0B',
      error: '#EF4444',
      chartColors: ['#3B82F6', '#10B981', '#F59E0B', '#EF4444', '#8B5CF6', '#EC4899']
    },
    modern: {
      name: 'Modern Purple',
      background: '#FFFFFF',
      titleSlideGradient: ['#6366F1', '#8B5CF6'],
      primary: '#6366F1',
      secondary: '#8B5CF6',
      accent: '#A78BFA',
      text: '#111827',
      textLight: '#6B7280',
      textOnDark: '#FFFFFF',
      surface: '#FAF5FF',
      border: '#E9D5FF',
      success: '#22C55E',
      warning: '#FACC15',
      error: '#DC2626',
      chartColors: ['#8B5CF6', '#6366F1', '#A78BFA', '#C084FC', '#E879F9', '#F472B6']
    },
    elegant: {
      name: 'Elegant Teal',
      background: '#FFFFFF',
      titleSlideGradient: ['#0F766E', '#14B8A6'],
      primary: '#0F766E',
      secondary: '#14B8A6',
      accent: '#2DD4BF',
      text: '#134E4A',
      textLight: '#5EEAD4',
      textOnDark: '#FFFFFF',
      surface: '#F0FDFA',
      border: '#CCFBF1',
      success: '#14B8A6',
      warning: '#FCD34D',
      error: '#F87171',
      chartColors: ['#14B8A6', '#10B981', '#34D399', '#6EE7B7', '#A7F3D0', '#D1FAE5']
    },
    vibrant: {
      name: 'Vibrant Energy',
      background: '#FFFFFF',
      titleSlideGradient: ['#DC2626', '#F97316'],
      primary: '#DC2626',
      secondary: '#F97316',
      accent: '#FB923C',
      text: '#1F2937',
      textLight: '#6B7280',
      textOnDark: '#FFFFFF',
      surface: '#FFF7ED',
      border: '#FED7AA',
      success: '#16A34A',
      warning: '#FCD34D',
      error: '#DC2626',
      chartColors: ['#F97316', '#DC2626', '#FACC15', '#FB923C', '#FCD34D', '#FDE68A']
    }
  },
  
  // Typography system
  TYPOGRAPHY: {
    hero: { size: 44, family: 'Montserrat' },
    title: { size: 32, family: 'Montserrat' },
    subtitle: { size: 20, family: 'Open Sans' },
    heading: { size: 24, family: 'Montserrat' },
    subheading: { size: 18, family: 'Open Sans' },
    body: { size: 14, family: 'Open Sans' },
    caption: { size: 12, family: 'Open Sans' },
    label: { size: 11, family: 'Open Sans' }
  },
  
  /**
   * Main entry point - Generate professional presentation
   */
  generatePresentation: function(presentation, data, config) {
    try {
      // Set up configuration
      const theme = this.THEMES[config.theme] || this.THEMES.corporate;
      const slideCount = config.slideCount || 5;
      
      // STEP 1: Clean and prepare presentation
      this.cleanPresentation(presentation);
      
      // STEP 2: Create required number of slides
      this.ensureSlideCount(presentation, slideCount);
      
      // STEP 3: Analyze data for intelligent content
      const analysis = this.analyzeData(data);
      
      // STEP 4: Generate slides with proper order
      const slides = presentation.getSlides();
      let currentIndex = 0;
      
      // Title Slide (always first)
      if (currentIndex < slides.length) {
        this.createTitleSlide(slides[currentIndex], config, theme, data, analysis);
        currentIndex++;
      }
      
      // Executive Summary (if we have enough slides)
      if (currentIndex < slides.length && slideCount > 3) {
        this.createExecutiveSummary(slides[currentIndex], data, analysis, theme);
        currentIndex++;
      }
      
      // Data Visualization Slides (dynamic count)
      const visualizationCount = Math.min(
        slideCount - currentIndex - 1, // Leave room for closing
        Math.max(1, Math.floor((slideCount - 2) * 0.6)) // 60% for data
      );
      
      for (let i = 0; i < visualizationCount && currentIndex < slides.length - 1; i++) {
        this.createDataSlide(slides[currentIndex], data, analysis, theme, i);
        currentIndex++;
      }
      
      // Closing/Next Steps Slide (always last)
      if (currentIndex < slides.length) {
        this.createClosingSlide(slides[currentIndex], data, analysis, theme, config);
        currentIndex++;
      }
      
      // STEP 5: Apply finishing touches
      this.applyFinishingTouches(presentation, theme);
      
      return { success: true, slideCount: currentIndex };
      
    } catch (error) {
      Logger.error('Error in generatePresentation:', error);
      return { success: false, error: error.toString() };
    }
  },
  
  /**
   * Clean existing presentation
   */
  cleanPresentation: function(presentation) {
    const slides = presentation.getSlides();
    
    // Clean each slide
    slides.forEach((slide, index) => {
      const elements = slide.getPageElements();
      
      // Remove all placeholder text boxes and shapes
      elements.forEach(element => {
        try {
          // Check if it's a shape with placeholder text
          if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
            const shape = element.asShape();
            
            // Try to get placeholder type - if it exists, remove default text
            try {
              const placeholderType = shape.getPlaceholderType();
              // Clear default placeholder text
              if (shape.getText().asString().includes('Click to add') ||
                  shape.getText().asString().includes('Title') ||
                  shape.getText().asString() === '') {
                shape.remove();
              }
            } catch (e) {
              // Not a placeholder, check if it's default content
              const text = shape.getText().asString();
              if (text.includes('Click to add') || text.includes('Title') || text === '') {
                shape.remove();
              }
            }
          }
        } catch (e) {
          // Safe to ignore - element might not support these operations
        }
      });
    });
  },
  
  /**
   * Ensure we have exactly the right number of slides
   */
  ensureSlideCount: function(presentation, targetCount) {
    const slides = presentation.getSlides();
    const currentCount = slides.length;
    
    if (currentCount < targetCount) {
      // Add more slides
      for (let i = currentCount; i < targetCount; i++) {
        presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
      }
    } else if (currentCount > targetCount) {
      // Remove extra slides
      for (let i = currentCount - 1; i >= targetCount; i--) {
        slides[i].remove();
      }
    }
  },
  
  /**
   * Create title slide with beautiful gradient
   */
  createTitleSlide: function(slide, config, theme, data, analysis) {
    // Create gradient background
    this.addGradientBackground(slide, theme.titleSlideGradient[0], theme.titleSlideGradient[1]);
    
    // Add subtle pattern overlay
    this.addPatternOverlay(slide, theme);
    
    // Title
    const title = config.title || this.generateSmartTitle(data, analysis);
    const titleBox = slide.insertTextBox(title,
      this.MARGINS.left,
      this.DIMENSIONS.height * 0.35,
      this.DIMENSIONS.width - this.MARGINS.left - this.MARGINS.right,
      60
    );
    titleBox.getText().getTextStyle()
      .setFontSize(this.TYPOGRAPHY.hero.size)
      .setFontFamily(this.TYPOGRAPHY.hero.family)
      .setBold(true)
      .setForegroundColor(theme.textOnDark);
    titleBox.getText().getParagraphStyle()
      .setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    
    // Subtitle
    if (config.subtitle) {
      const subtitleBox = slide.insertTextBox(config.subtitle,
        this.MARGINS.left,
        this.DIMENSIONS.height * 0.5,
        this.DIMENSIONS.width - this.MARGINS.left - this.MARGINS.right,
        40
      );
      subtitleBox.getText().getTextStyle()
        .setFontSize(this.TYPOGRAPHY.subtitle.size)
        .setFontFamily(this.TYPOGRAPHY.subtitle.family)
        .setForegroundColor(theme.textOnDark + 'CC'); // Slightly transparent
      subtitleBox.getText().getParagraphStyle()
        .setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    }
    
    // Add data summary at bottom
    const summaryText = `${data.data.length} records | ${data.headers.length} fields`;
    const summaryBox = slide.insertTextBox(summaryText,
      this.MARGINS.left,
      this.DIMENSIONS.height - 60,
      this.DIMENSIONS.width - this.MARGINS.left - this.MARGINS.right,
      20
    );
    summaryBox.getText().getTextStyle()
      .setFontSize(this.TYPOGRAPHY.caption.size)
      .setFontFamily(this.TYPOGRAPHY.caption.family)
      .setForegroundColor(theme.textOnDark + '99'); // More transparent
    summaryBox.getText().getParagraphStyle()
      .setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    
    // Add decorative element
    const accent = slide.insertShape(SlidesApp.ShapeType.RECTANGLE,
      0, this.DIMENSIONS.height - 5, this.DIMENSIONS.width, 5);
    accent.getFill().setSolidFill(theme.accent);
    accent.getBorder().setTransparent();
  },
  
  /**
   * Create executive summary with key metrics
   */
  createExecutiveSummary: function(slide, data, analysis, theme) {
    // White background
    this.setSlideBackground(slide, theme.background);
    
    // Title
    const titleBox = slide.insertTextBox('Executive Summary',
      this.MARGINS.left,
      this.MARGINS.top,
      this.DIMENSIONS.width - this.MARGINS.left - this.MARGINS.right,
      40
    );
    titleBox.getText().getTextStyle()
      .setFontSize(this.TYPOGRAPHY.title.size)
      .setFontFamily(this.TYPOGRAPHY.title.family)
      .setBold(true)
      .setForegroundColor(theme.primary);
    
    // Key metrics in cards
    const metrics = this.extractKeyMetrics(data, analysis);
    const cardWidth = 150;
    const cardHeight = 80;
    const cardSpacing = 20;
    const startX = (this.DIMENSIONS.width - (metrics.length * cardWidth + (metrics.length - 1) * cardSpacing)) / 2;
    
    metrics.slice(0, 4).forEach((metric, index) => {
      const x = startX + index * (cardWidth + cardSpacing);
      const y = 120;
      
      // Card background
      const card = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE,
        x, y, cardWidth, cardHeight);
      card.getFill().setSolidFill(theme.surface);
      card.getBorder().getLineFill().setSolidFill(theme.border);
      card.getBorder().setWeight(1);
      
      // Metric value
      const valueBox = slide.insertTextBox(metric.value,
        x, y + 15, cardWidth, 30);
      valueBox.getText().getTextStyle()
        .setFontSize(24)
        .setFontFamily('Montserrat')
        .setBold(true)
        .setForegroundColor(theme.primary);
      valueBox.getText().getParagraphStyle()
        .setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
      
      // Metric label
      const labelBox = slide.insertTextBox(metric.label,
        x, y + 45, cardWidth, 20);
      labelBox.getText().getTextStyle()
        .setFontSize(this.TYPOGRAPHY.caption.size)
        .setFontFamily(this.TYPOGRAPHY.caption.family)
        .setForegroundColor(theme.textLight);
      labelBox.getText().getParagraphStyle()
        .setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    });
    
    // Insights section
    const insights = this.generateInsights(data, analysis);
    let insightY = 220;
    
    insights.slice(0, 3).forEach((insight, index) => {
      // Bullet point
      const bullet = slide.insertShape(SlidesApp.ShapeType.ELLIPSE,
        this.MARGINS.left, insightY + index * 40, 8, 8);
      bullet.getFill().setSolidFill(theme.accent);
      bullet.getBorder().setTransparent();
      
      // Insight text
      const insightBox = slide.insertTextBox(insight,
        this.MARGINS.left + 20, insightY + index * 40 - 5,
        this.DIMENSIONS.width - this.MARGINS.left - this.MARGINS.right - 20,
        30
      );
      insightBox.getText().getTextStyle()
        .setFontSize(this.TYPOGRAPHY.body.size)
        .setFontFamily(this.TYPOGRAPHY.body.family)
        .setForegroundColor(theme.text);
    });
  },
  
  /**
   * Create data visualization slide
   */
  createDataSlide: function(slide, data, analysis, theme, slideIndex) {
    // White background
    this.setSlideBackground(slide, theme.background);
    
    // Determine what type of visualization to create
    const vizType = this.determineVisualizationType(data, analysis, slideIndex);
    
    // Title based on visualization
    const titles = {
      chart: 'Data Visualization',
      table: 'Data Overview',
      metrics: 'Key Metrics',
      trend: 'Trend Analysis',
      distribution: 'Distribution'
    };
    
    const titleBox = slide.insertTextBox(titles[vizType] || 'Data Analysis',
      this.MARGINS.left,
      this.MARGINS.top,
      this.DIMENSIONS.width - this.MARGINS.left - this.MARGINS.right,
      40
    );
    titleBox.getText().getTextStyle()
      .setFontSize(this.TYPOGRAPHY.heading.size)
      .setFontFamily(this.TYPOGRAPHY.heading.family)
      .setBold(true)
      .setForegroundColor(theme.primary);
    
    // Create appropriate visualization
    const contentY = this.MARGINS.contentTop;
    const contentHeight = this.DIMENSIONS.height - this.MARGINS.contentTop - this.MARGINS.bottom;
    const contentWidth = this.DIMENSIONS.width - this.MARGINS.left - this.MARGINS.right;
    
    switch(vizType) {
      case 'chart':
        this.addChartToSlide(slide, data, analysis, theme, 
          this.MARGINS.left, contentY, contentWidth, contentHeight);
        break;
        
      case 'table':
        this.addSmartTableToSlide(slide, data, theme,
          this.MARGINS.left, contentY, contentWidth, contentHeight);
        break;
        
      case 'metrics':
        this.addMetricsGridToSlide(slide, data, analysis, theme,
          this.MARGINS.left, contentY, contentWidth, contentHeight);
        break;
        
      default:
        // Default to chart
        this.addChartToSlide(slide, data, analysis, theme,
          this.MARGINS.left, contentY, contentWidth, contentHeight);
    }
  },
  
  /**
   * Create closing slide with next steps
   */
  createClosingSlide: function(slide, data, analysis, theme, config) {
    // Gradient background (reversed from title)
    this.addGradientBackground(slide, theme.titleSlideGradient[1], theme.titleSlideGradient[0]);
    
    // Title
    const titleBox = slide.insertTextBox('Next Steps',
      this.MARGINS.left,
      this.MARGINS.top,
      this.DIMENSIONS.width - this.MARGINS.left - this.MARGINS.right,
      50
    );
    titleBox.getText().getTextStyle()
      .setFontSize(this.TYPOGRAPHY.title.size)
      .setFontFamily(this.TYPOGRAPHY.title.family)
      .setBold(true)
      .setForegroundColor(theme.textOnDark);
    titleBox.getText().getParagraphStyle()
      .setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    
    // Generate next steps
    const nextSteps = this.generateNextSteps(data, analysis);
    const stepStartY = 140;
    
    nextSteps.slice(0, 4).forEach((step, index) => {
      const y = stepStartY + index * 50;
      
      // Step number in circle
      const circle = slide.insertShape(SlidesApp.ShapeType.ELLIPSE,
        this.DIMENSIONS.width / 2 - 200, y, 30, 30);
      circle.getFill().setSolidFill(theme.textOnDark);
      circle.getBorder().setTransparent();
      
      const numberBox = slide.insertTextBox((index + 1).toString(),
        this.DIMENSIONS.width / 2 - 200, y + 5, 30, 20);
      numberBox.getText().getTextStyle()
        .setFontSize(14)
        .setFontFamily('Montserrat')
        .setBold(true)
        .setForegroundColor(theme.primary);
      numberBox.getText().getParagraphStyle()
        .setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
      
      // Step text
      const stepBox = slide.insertTextBox(step,
        this.DIMENSIONS.width / 2 - 160, y + 5,
        320, 30
      );
      stepBox.getText().getTextStyle()
        .setFontSize(this.TYPOGRAPHY.body.size)
        .setFontFamily(this.TYPOGRAPHY.body.family)
        .setForegroundColor(theme.textOnDark);
    });
    
    // Call to action
    if (config.callToAction) {
      const ctaButton = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE,
        this.DIMENSIONS.width / 2 - 100,
        this.DIMENSIONS.height - 80,
        200, 45
      );
      ctaButton.getFill().setSolidFill(theme.textOnDark);
      ctaButton.getBorder().setTransparent();
      
      const ctaText = slide.insertTextBox('Get Started',
        this.DIMENSIONS.width / 2 - 100,
        this.DIMENSIONS.height - 70,
        200, 25
      );
      ctaText.getText().getTextStyle()
        .setFontSize(16)
        .setFontFamily('Montserrat')
        .setBold(true)
        .setForegroundColor(theme.primary);
      ctaText.getText().getParagraphStyle()
        .setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    }
  },
  
  /**
   * Add gradient background to slide
   */
  addGradientBackground: function(slide, color1, color2) {
    // Since Google Slides doesn't support gradients directly, 
    // we create multiple rectangles with gradually changing colors
    const steps = 10;
    const stepHeight = this.DIMENSIONS.height / steps;
    
    for (let i = 0; i < steps; i++) {
      const ratio = i / steps;
      const color = this.blendColors(color1, color2, ratio);
      
      const rect = slide.insertShape(SlidesApp.ShapeType.RECTANGLE,
        0, i * stepHeight, this.DIMENSIONS.width, stepHeight + 1);
      rect.getFill().setSolidFill(color);
      rect.getBorder().setTransparent();
      rect.sendToBack();
    }
  },
  
  /**
   * Add subtle pattern overlay for visual interest
   */
  addPatternOverlay: function(slide, theme) {
    // Add subtle geometric shapes
    for (let i = 0; i < 3; i++) {
      const size = 100 + i * 50;
      const opacity = 10 - i * 3; // Decrease opacity for larger shapes
      
      const circle = slide.insertShape(SlidesApp.ShapeType.ELLIPSE,
        Math.random() * this.DIMENSIONS.width,
        Math.random() * this.DIMENSIONS.height,
        size, size
      );
      
      const hexOpacity = opacity.toString(16).padStart(2, '0');
      circle.getFill().setSolidFill('#FFFFFF' + hexOpacity);
      circle.getBorder().setTransparent();
      circle.sendToBack();
    }
  },
  
  /**
   * Set slide background color
   */
  setSlideBackground: function(slide, color) {
    const bg = slide.insertShape(SlidesApp.ShapeType.RECTANGLE,
      0, 0, this.DIMENSIONS.width, this.DIMENSIONS.height);
    bg.getFill().setSolidFill(color);
    bg.getBorder().setTransparent();
    bg.sendToBack();
  },
  
  /**
   * Add chart to slide
   */
  addChartToSlide: function(slide, data, analysis, theme, x, y, width, height) {
    try {
      // Determine best chart type
      const chartType = this.determineBestChartType(data, analysis);
      
      // Prepare data for chart
      const chartData = this.prepareChartData(data, chartType);
      
      // Build chart
      const chart = this.buildChart(chartData, chartType, theme);
      
      // Insert chart
      slide.insertChart(chart, x, y, width, height - 40);
      
      // Add insight caption
      const insight = this.generateChartInsight(data, analysis, chartType);
      if (insight) {
        const insightBox = slide.insertTextBox(insight,
          x, y + height - 30, width, 25);
        insightBox.getText().getTextStyle()
          .setFontSize(this.TYPOGRAPHY.caption.size)
          .setFontFamily(this.TYPOGRAPHY.caption.family)
          .setForegroundColor(theme.textLight)
          .setItalic(true);
        insightBox.getText().getParagraphStyle()
          .setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
      }
    } catch (error) {
      // If chart fails, fall back to table
      Logger.warn('Chart creation failed, using table instead:', error);
      this.addSmartTableToSlide(slide, data, theme, x, y, width, height);
    }
  },
  
  /**
   * Add smart table (not just raw data dump)
   */
  addSmartTableToSlide: function(slide, data, theme, x, y, width, height) {
    // Determine optimal table size
    const maxRows = 8; // Including header
    const maxCols = 5;
    
    // Select most important columns
    const importantColumns = this.selectImportantColumns(data, maxCols);
    const displayRows = Math.min(data.data.length, maxRows - 1);
    
    // Create table
    const table = slide.insertTable(displayRows + 1, importantColumns.length, x, y, width, Math.min(height, 250));
    
    // Style header row
    for (let col = 0; col < importantColumns.length; col++) {
      const cell = table.getCell(0, col);
      cell.getFill().setSolidFill(theme.primary);
      const cellText = cell.getText();
      cellText.setText(data.headers[importantColumns[col]]);
      cellText.getTextStyle()
        .setFontSize(this.TYPOGRAPHY.caption.size)
        .setFontFamily(this.TYPOGRAPHY.caption.family)
        .setBold(true)
        .setForegroundColor(theme.textOnDark);
    }
    
    // Add data with alternating row colors
    for (let row = 0; row < displayRows; row++) {
      for (let col = 0; col < importantColumns.length; col++) {
        const cell = table.getCell(row + 1, col);
        
        // Alternating row colors
        if (row % 2 === 1) {
          cell.getFill().setSolidFill(theme.surface);
        }
        
        // Set data
        const value = data.data[row][importantColumns[col]];
        const cellText = cell.getText();
        cellText.setText(this.formatCellValue(value));
        cellText.getTextStyle()
          .setFontSize(this.TYPOGRAPHY.caption.size)
          .setFontFamily(this.TYPOGRAPHY.caption.family)
          .setForegroundColor(theme.text);
      }
    }
    
    // Add note if data is truncated
    if (data.data.length > displayRows || data.headers.length > importantColumns.length) {
      const noteText = `Showing ${displayRows} of ${data.data.length} rows, ${importantColumns.length} of ${data.headers.length} columns`;
      const noteBox = slide.insertTextBox(noteText,
        x, y + Math.min(height, 250) + 10, width, 20);
      noteBox.getText().getTextStyle()
        .setFontSize(10)
        .setFontFamily(this.TYPOGRAPHY.caption.family)
        .setForegroundColor(theme.textLight)
        .setItalic(true);
    }
  },
  
  /**
   * Add metrics grid to slide
   */
  addMetricsGridToSlide: function(slide, data, analysis, theme, x, y, width, height) {
    const metrics = this.extractDetailedMetrics(data, analysis);
    const gridCols = 3;
    const gridRows = Math.ceil(metrics.length / gridCols);
    const cardWidth = (width - 20 * (gridCols - 1)) / gridCols;
    const cardHeight = 70;
    const spacing = 20;
    
    metrics.slice(0, 9).forEach((metric, index) => {
      const row = Math.floor(index / gridCols);
      const col = index % gridCols;
      const cardX = x + col * (cardWidth + spacing);
      const cardY = y + row * (cardHeight + spacing);
      
      // Card background
      const card = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE,
        cardX, cardY, cardWidth, cardHeight);
      card.getFill().setSolidFill(theme.surface);
      card.getBorder().getLineFill().setSolidFill(theme.border);
      card.getBorder().setWeight(1);
      
      // Metric value
      const valueBox = slide.insertTextBox(metric.value,
        cardX + 10, cardY + 10, cardWidth - 20, 25);
      valueBox.getText().getTextStyle()
        .setFontSize(20)
        .setFontFamily('Montserrat')
        .setBold(true)
        .setForegroundColor(theme.primary);
      
      // Metric label
      const labelBox = slide.insertTextBox(metric.label,
        cardX + 10, cardY + 35, cardWidth - 20, 20);
      labelBox.getText().getTextStyle()
        .setFontSize(this.TYPOGRAPHY.caption.size)
        .setFontFamily(this.TYPOGRAPHY.caption.family)
        .setForegroundColor(theme.textLight);
    });
  },
  
  /**
   * Apply finishing touches to presentation
   */
  applyFinishingTouches: function(presentation, theme) {
    const slides = presentation.getSlides();
    
    slides.forEach((slide, index) => {
      // Add slide number (except title slide)
      if (index > 0) {
        const pageNum = slide.insertTextBox(`${index + 1}`,
          this.DIMENSIONS.width - 40, this.DIMENSIONS.height - 30, 30, 20);
        pageNum.getText().getTextStyle()
          .setFontSize(10)
          .setFontFamily(this.TYPOGRAPHY.caption.family)
          .setForegroundColor(theme.textLight);
        pageNum.getText().getParagraphStyle()
          .setParagraphAlignment(SlidesApp.ParagraphAlignment.END);
      }
      
      // Add subtle footer line (except title and closing slides)
      if (index > 0 && index < slides.length - 1) {
        const footerLine = slide.insertShape(SlidesApp.ShapeType.RECTANGLE,
          this.MARGINS.left, this.DIMENSIONS.height - 35,
          this.DIMENSIONS.width - this.MARGINS.left - this.MARGINS.right, 1);
        footerLine.getFill().setSolidFill(theme.border);
        footerLine.getBorder().setTransparent();
      }
    });
  },
  
  // =====================================
  // HELPER FUNCTIONS
  // =====================================
  
  /**
   * Analyze data for intelligent content generation
   */
  analyzeData: function(data) {
    const analysis = {
      numericColumns: [],
      textColumns: [],
      dateColumns: [],
      totalRows: data.data.length,
      totalColumns: data.headers.length,
      hasNullValues: false,
      correlations: [],
      trends: []
    };
    
    // Analyze column types
    data.headers.forEach((header, index) => {
      const columnData = data.data.map(row => row[index]).filter(v => v != null && v !== '');
      if (columnData.length === 0) return;
      
      const sample = columnData[0];
      
      if (!isNaN(parseFloat(sample))) {
        analysis.numericColumns.push({
          index: index,
          name: header,
          min: Math.min(...columnData.map(v => parseFloat(v))),
          max: Math.max(...columnData.map(v => parseFloat(v))),
          avg: columnData.reduce((sum, v) => sum + parseFloat(v), 0) / columnData.length
        });
      } else if (this.isDate(sample)) {
        analysis.dateColumns.push({ index: index, name: header });
      } else {
        analysis.textColumns.push({ index: index, name: header });
      }
      
      // Check for null values
      if (columnData.length < data.data.length) {
        analysis.hasNullValues = true;
      }
    });
    
    return analysis;
  },
  
  /**
   * Generate smart title from data
   */
  generateSmartTitle: function(data, analysis) {
    if (analysis.dateColumns.length > 0 && analysis.numericColumns.length > 0) {
      return `${analysis.numericColumns[0].name} Analysis`;
    } else if (analysis.numericColumns.length > 0) {
      return `${analysis.numericColumns[0].name} Report`;
    } else if (data.headers.length > 0) {
      return `${data.headers[0]} Overview`;
    }
    return 'Data Presentation';
  },
  
  /**
   * Extract key metrics from data
   */
  extractKeyMetrics: function(data, analysis) {
    const metrics = [];
    
    // Add total records
    metrics.push({
      label: 'Total Records',
      value: data.data.length.toLocaleString()
    });
    
    // Add numeric column metrics
    analysis.numericColumns.slice(0, 3).forEach(col => {
      metrics.push({
        label: `Avg ${col.name}`,
        value: Math.round(col.avg).toLocaleString()
      });
    });
    
    // Add completeness if there are null values
    if (analysis.hasNullValues) {
      const totalCells = data.data.length * data.headers.length;
      const nullCells = data.data.flat().filter(v => v == null || v === '').length;
      const completeness = Math.round(((totalCells - nullCells) / totalCells) * 100);
      metrics.push({
        label: 'Data Completeness',
        value: completeness + '%'
      });
    }
    
    return metrics;
  },
  
  /**
   * Extract detailed metrics for grid display
   */
  extractDetailedMetrics: function(data, analysis) {
    const metrics = [];
    
    // Basic metrics
    metrics.push(
      { label: 'Rows', value: data.data.length.toLocaleString() },
      { label: 'Columns', value: data.headers.length.toLocaleString() }
    );
    
    // Numeric metrics
    analysis.numericColumns.forEach(col => {
      metrics.push(
        { label: `Min ${col.name}`, value: Math.round(col.min).toLocaleString() },
        { label: `Max ${col.name}`, value: Math.round(col.max).toLocaleString() },
        { label: `Avg ${col.name}`, value: Math.round(col.avg).toLocaleString() }
      );
    });
    
    return metrics.slice(0, 9); // Limit to 9 for 3x3 grid
  },
  
  /**
   * Generate insights from data
   */
  generateInsights: function(data, analysis) {
    const insights = [];
    
    // Data size insight
    insights.push(`Dataset contains ${data.data.length} records across ${data.headers.length} fields`);
    
    // Numeric insights
    if (analysis.numericColumns.length > 0) {
      const col = analysis.numericColumns[0];
      const range = col.max - col.min;
      insights.push(`${col.name} ranges from ${Math.round(col.min)} to ${Math.round(col.max)} (spread: ${Math.round(range)})`);
    }
    
    // Data quality insight
    if (analysis.hasNullValues) {
      insights.push('Some data fields contain missing values that may need attention');
    } else {
      insights.push('Dataset is complete with no missing values');
    }
    
    // Column type distribution
    const distribution = [];
    if (analysis.numericColumns.length > 0) distribution.push(`${analysis.numericColumns.length} numeric`);
    if (analysis.textColumns.length > 0) distribution.push(`${analysis.textColumns.length} text`);
    if (analysis.dateColumns.length > 0) distribution.push(`${analysis.dateColumns.length} date`);
    if (distribution.length > 0) {
      insights.push(`Data includes ${distribution.join(', ')} columns`);
    }
    
    return insights;
  },
  
  /**
   * Generate next steps
   */
  generateNextSteps: function(data, analysis) {
    const steps = [];
    
    // Data quality steps
    if (analysis.hasNullValues) {
      steps.push('Review and address missing data points');
    }
    
    // Analysis steps
    if (analysis.numericColumns.length > 1) {
      steps.push('Analyze correlations between numeric variables');
    }
    
    if (analysis.dateColumns.length > 0) {
      steps.push('Perform time-series analysis for trends');
    }
    
    // General steps
    steps.push('Share findings with stakeholders');
    steps.push('Schedule follow-up analysis');
    steps.push('Implement data-driven recommendations');
    
    return steps.slice(0, 4); // Limit to 4 steps
  },
  
  /**
   * Determine visualization type for slide
   */
  determineVisualizationType: function(data, analysis, slideIndex) {
    // First data slide - show chart if we have numeric data
    if (slideIndex === 0 && analysis.numericColumns.length > 0) {
      return 'chart';
    }
    
    // Second slide - show metrics grid
    if (slideIndex === 1 && analysis.numericColumns.length > 2) {
      return 'metrics';
    }
    
    // Otherwise show smart table
    return 'table';
  },
  
  /**
   * Determine best chart type
   */
  determineBestChartType: function(data, analysis) {
    // If we have time series data
    if (analysis.dateColumns.length > 0 && analysis.numericColumns.length > 0) {
      return Charts.ChartType.LINE;
    }
    
    // If we have categories and values
    if (analysis.textColumns.length > 0 && analysis.numericColumns.length > 0) {
      // Use bar for many categories, column for few
      return data.data.length > 8 ? Charts.ChartType.BAR : Charts.ChartType.COLUMN;
    }
    
    // If we have multiple numeric columns
    if (analysis.numericColumns.length > 1) {
      return Charts.ChartType.SCATTER;
    }
    
    // Default to column chart
    return Charts.ChartType.COLUMN;
  },
  
  /**
   * Prepare data for charting
   */
  prepareChartData: function(data, chartType) {
    // Limit data for readability
    const maxDataPoints = 20;
    const limitedData = {
      headers: data.headers,
      data: data.data.slice(0, maxDataPoints)
    };
    
    return limitedData;
  },
  
  /**
   * Build chart with theming
   */
  buildChart: function(chartData, chartType, theme) {
    const dataTable = Charts.newDataTable();
    
    // Add headers
    chartData.headers.forEach(header => {
      dataTable.addColumn(Charts.ColumnType.STRING, header);
    });
    
    // Add data
    chartData.data.forEach(row => {
      dataTable.addRow(row.map(val => val != null ? val.toString() : ''));
    });
    
    // Build chart
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
    
    chartBuilder
      .setDataTable(dataTable.build())
      .setOption('colors', theme.chartColors)
      .setOption('backgroundColor', theme.background)
      .setOption('titleTextStyle', { color: theme.text })
      .setOption('legendTextStyle', { color: theme.textLight })
      .setOption('hAxis.textStyle', { color: theme.textLight })
      .setOption('vAxis.textStyle', { color: theme.textLight })
      .setDimensions(600, 400);
    
    return chartBuilder.build();
  },
  
  /**
   * Generate chart insight
   */
  generateChartInsight: function(data, analysis, chartType) {
    if (analysis.numericColumns.length > 0) {
      const col = analysis.numericColumns[0];
      
      switch(chartType) {
        case Charts.ChartType.LINE:
          return `Showing trend over time for ${col.name}`;
        case Charts.ChartType.BAR:
        case Charts.ChartType.COLUMN:
          return `Comparing ${col.name} across categories`;
        case Charts.ChartType.SCATTER:
          return `Relationship between variables`;
        default:
          return `Visual representation of ${col.name}`;
      }
    }
    return null;
  },
  
  /**
   * Select most important columns for table display
   */
  selectImportantColumns: function(data, maxCols) {
    const columns = [];
    
    // First, add any ID or name columns
    data.headers.forEach((header, index) => {
      if (header.toLowerCase().includes('id') || 
          header.toLowerCase().includes('name')) {
        columns.push(index);
      }
    });
    
    // Then add numeric columns
    const numericCols = [];
    data.headers.forEach((header, index) => {
      if (!columns.includes(index)) {
        const sample = data.data[0][index];
        if (!isNaN(parseFloat(sample))) {
          numericCols.push(index);
        }
      }
    });
    columns.push(...numericCols);
    
    // Finally, add other columns
    data.headers.forEach((header, index) => {
      if (!columns.includes(index)) {
        columns.push(index);
      }
    });
    
    return columns.slice(0, maxCols);
  },
  
  /**
   * Format cell value for display
   */
  formatCellValue: function(value) {
    if (value == null || value === '') return '-';
    
    // Truncate long text
    const str = value.toString();
    if (str.length > 30) {
      return str.substring(0, 27) + '...';
    }
    
    // Format numbers
    if (!isNaN(parseFloat(value))) {
      const num = parseFloat(value);
      if (num % 1 === 0) {
        return num.toLocaleString();
      } else {
        return num.toFixed(2);
      }
    }
    
    return str;
  },
  
  /**
   * Check if value is a date
   */
  isDate: function(value) {
    if (!value) return false;
    const date = new Date(value);
    return date instanceof Date && !isNaN(date) && value.toString().match(/\d{4}|\d{2}\/\d{2}/);
  },
  
  /**
   * Blend two colors
   */
  blendColors: function(color1, color2, ratio) {
    // Remove # if present
    color1 = color1.replace('#', '');
    color2 = color2.replace('#', '');
    
    // Parse colors
    const r1 = parseInt(color1.substring(0, 2), 16);
    const g1 = parseInt(color1.substring(2, 4), 16);
    const b1 = parseInt(color1.substring(4, 6), 16);
    
    const r2 = parseInt(color2.substring(0, 2), 16);
    const g2 = parseInt(color2.substring(2, 4), 16);
    const b2 = parseInt(color2.substring(4, 6), 16);
    
    // Blend
    const r = Math.round(r1 * (1 - ratio) + r2 * ratio);
    const g = Math.round(g1 * (1 - ratio) + g2 * ratio);
    const b = Math.round(b1 * (1 - ratio) + b2 * ratio);
    
    // Convert back to hex
    return '#' + 
           r.toString(16).padStart(2, '0') +
           g.toString(16).padStart(2, '0') +
           b.toString(16).padStart(2, '0');
  }
};