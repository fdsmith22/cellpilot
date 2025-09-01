/**
 * CellM8 - Intelligent Spreadsheet to Presentation Converter
 * Transforms Google Sheets data into professional Google Slides presentations
 * using AI-powered analysis and smart templating
 */

const CellM8 = {
  // Configuration
  config: {
    maxSlides: 50,
    minSlides: 5,
    defaultSlides: 12,
    chartTypes: ['COLUMN', 'LINE', 'PIE', 'BAR', 'AREA', 'SCATTER'],
    templates: ['corporate', 'sales', 'financial', 'modern', 'creative'],
    presentationStyles: ['executive', 'detailed', 'sales', 'financial', 'project', 'custom']
  },

  // Template definitions
  templates: {
    corporate: {
      name: 'Corporate Professional',
      colorScheme: {
        primary: '#1f4e79',
        secondary: '#4472c4',
        accent: '#70ad47',
        background: '#ffffff',
        text: '#333333'
      },
      fontFamily: 'Calibri',
      layouts: {
        title: 'TITLE_AND_SUBTITLE',
        content: 'TITLE_AND_BODY',
        chart: 'TITLE_AND_TWO_COLUMNS',
        section: 'SECTION_HEADER'
      }
    },
    sales: {
      name: 'Sales Focused',
      colorScheme: {
        primary: '#c55a11',
        secondary: '#f79646',
        accent: '#9cbb58',
        background: '#ffffff',
        text: '#2c2c2c'
      },
      fontFamily: 'Arial',
      layouts: {
        title: 'TITLE_ONLY',
        content: 'TITLE_AND_BODY',
        chart: 'TITLE_AND_TWO_COLUMNS',
        section: 'SECTION_HEADER'
      }
    },
    financial: {
      name: 'Financial Standard',
      colorScheme: {
        primary: '#1f4e79',
        secondary: '#4472c4',
        accent: '#843c0c',
        background: '#f5f5f5',
        text: '#1a1a1a'
      },
      fontFamily: 'Times New Roman',
      layouts: {
        title: 'TITLE_AND_SUBTITLE',
        content: 'TITLE_AND_BODY',
        chart: 'TITLE_AND_TWO_COLUMNS',
        section: 'SECTION_HEADER'
      }
    },
    modern: {
      name: 'Modern Clean',
      colorScheme: {
        primary: '#4285f4',
        secondary: '#34a853',
        accent: '#fbbc04',
        background: '#ffffff',
        text: '#202124'
      },
      fontFamily: 'Roboto',
      layouts: {
        title: 'TITLE_ONLY',
        content: 'TITLE_AND_BODY',
        chart: 'BLANK',
        section: 'SECTION_HEADER'
      }
    },
    creative: {
      name: 'Creative Colorful',
      colorScheme: {
        primary: '#9c27b0',
        secondary: '#673ab7',
        accent: '#ff5722',
        background: '#fafafa',
        text: '#212121'
      },
      fontFamily: 'Montserrat',
      layouts: {
        title: 'TITLE_ONLY',
        content: 'TITLE_AND_BODY',
        chart: 'BLANK',
        section: 'SECTION_HEADER'
      }
    }
  },

  /**
   * Main function to create presentation from sheet data
   * @param {Object} config - Configuration object
   * @return {Object} Result with presentation ID and URL
   */
  createPresentation: function(config) {
    try {
      Logger.info('CellM8: Starting presentation creation', config);
      
      // Extract and analyze data
      const extractedData = this.extractSheetData(config.dataSource);
      if (!extractedData.success) {
        return { success: false, error: extractedData.error };
      }
      
      // Perform AI analysis
      const analysis = this.analyzeDataWithAI(extractedData.data);
      
      // Generate slide structure
      const slideStructure = this.generateSlideStructure(analysis, config);
      
      // Create Google Slides presentation
      const presentation = this.createGoogleSlides(slideStructure, config);
      
      if (presentation.success) {
        // Track usage for ML learning
        this.trackPresentationCreation(config, presentation);
        
        return {
          success: true,
          presentationId: presentation.presentationId,
          presentationUrl: presentation.url,
          slideCount: slideStructure.slides.length,
          insights: analysis.insights
        };
      } else {
        return { success: false, error: presentation.error };
      }
      
    } catch (error) {
      Logger.error('CellM8: Error creating presentation:', error);
      return { success: false, error: error.message };
    }
  },

  /**
   * Extract data from sheet based on source configuration
   * @param {Object} source - Data source configuration
   * @return {Object} Extracted data with metadata
   */
  extractSheetData: function(source) {
    try {
      const sheet = SpreadsheetApp.getActiveSheet();
      let range, data, headers;
      
      switch (source.type) {
        case 'entireSheet':
          range = sheet.getDataRange();
          break;
          
        case 'selectedRange':
          range = sheet.getActiveRange();
          if (!range) {
            return { success: false, error: 'No range selected' };
          }
          break;
          
        case 'visibleOnly':
          // Get visible cells only (excluding hidden rows/columns)
          range = this.getVisibleRange(sheet);
          break;
          
        case 'customRange':
          if (!source.range) {
            return { success: false, error: 'No custom range specified' };
          }
          range = sheet.getRange(source.range);
          break;
          
        default:
          range = sheet.getDataRange();
      }
      
      const values = range.getValues();
      const formulas = range.getFormulas();
      
      // Detect headers
      if (source.hasHeaders !== false) {
        headers = values[0];
        data = values.slice(1);
      } else {
        headers = this.generateDefaultHeaders(values[0].length);
        data = values;
      }
      
      // Clean and prepare data
      const cleanedData = this.cleanDataForPresentation(data, headers);
      
      return {
        success: true,
        data: cleanedData,
        headers: headers,
        metadata: {
          sheetName: sheet.getName(),
          rowCount: data.length,
          columnCount: headers.length,
          dataTypes: this.detectDataTypes(data),
          hasFormulas: formulas.flat().some(f => f !== '')
        }
      };
      
    } catch (error) {
      Logger.error('CellM8: Error extracting sheet data:', error);
      return { success: false, error: error.message };
    }
  },

  /**
   * Get only visible cells from sheet
   * @param {Sheet} sheet - Google Sheet object
   * @return {Range} Visible range
   */
  getVisibleRange: function(sheet) {
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    const visibleRows = [];
    
    for (let i = 0; i < values.length; i++) {
      if (!sheet.isRowHiddenByUser(i + 1)) {
        const visibleRow = [];
        for (let j = 0; j < values[i].length; j++) {
          if (!sheet.isColumnHiddenByUser(j + 1)) {
            visibleRow.push(values[i][j]);
          }
        }
        visibleRows.push(visibleRow);
      }
    }
    
    // Create a temporary range with visible data
    // This is a simplified approach - in production, we'd handle this more elegantly
    return {
      getValues: () => visibleRows,
      getFormulas: () => Array(visibleRows.length).fill(Array(visibleRows[0]?.length || 0).fill(''))
    };
  },

  /**
   * Analyze data using AI/ML techniques
   * @param {Object} data - Extracted data
   * @return {Object} Analysis results
   */
  analyzeDataWithAI: function(data) {
    try {
      const analysis = {
        dataTypes: {},
        statistics: {},
        patterns: {},
        insights: [],
        chartRecommendations: [],
        keyMetrics: []
      };
      
      // Analyze each column
      data.headers.forEach((header, index) => {
        const columnData = data.data.map(row => row[index]);
        const dataType = this.detectColumnDataType(columnData);
        
        analysis.dataTypes[header] = dataType;
        
        if (dataType === 'number') {
          // Calculate statistics
          analysis.statistics[header] = this.calculateStatistics(columnData);
          
          // Detect trends
          const trend = this.detectTrend(columnData);
          if (trend) {
            analysis.patterns[header] = trend;
            analysis.insights.push(`${header} shows a ${trend.direction} trend`);
          }
        } else if (dataType === 'date') {
          // Analyze time series
          const timeSeries = this.analyzeTimeSeries(columnData);
          if (timeSeries) {
            analysis.patterns[header] = timeSeries;
          }
        } else if (dataType === 'category') {
          // Analyze categories
          const categories = this.analyzeCategories(columnData);
          analysis.statistics[header] = categories;
        }
      });
      
      // Generate chart recommendations
      analysis.chartRecommendations = this.recommendCharts(analysis);
      
      // Extract key metrics
      analysis.keyMetrics = this.extractKeyMetrics(data, analysis);
      
      // Generate narrative insights
      analysis.insights = this.generateInsights(data, analysis);
      
      return analysis;
      
    } catch (error) {
      Logger.error('CellM8: Error analyzing data:', error);
      return {
        dataTypes: {},
        statistics: {},
        patterns: {},
        insights: ['Unable to fully analyze data'],
        chartRecommendations: [],
        keyMetrics: []
      };
    }
  },

  /**
   * Calculate statistics for numerical data
   * @param {Array} data - Column data
   * @return {Object} Statistics
   */
  calculateStatistics: function(data) {
    const numbers = data.filter(d => !isNaN(d) && d !== '').map(Number);
    
    if (numbers.length === 0) {
      return null;
    }
    
    const sum = numbers.reduce((a, b) => a + b, 0);
    const mean = sum / numbers.length;
    const sorted = numbers.sort((a, b) => a - b);
    const median = sorted[Math.floor(sorted.length / 2)];
    const min = sorted[0];
    const max = sorted[sorted.length - 1];
    
    // Calculate standard deviation
    const variance = numbers.reduce((acc, val) => acc + Math.pow(val - mean, 2), 0) / numbers.length;
    const stdDev = Math.sqrt(variance);
    
    return {
      count: numbers.length,
      sum: sum,
      mean: mean,
      median: median,
      min: min,
      max: max,
      stdDev: stdDev,
      range: max - min
    };
  },

  /**
   * Detect trend in numerical data
   * @param {Array} data - Column data
   * @return {Object} Trend information
   */
  detectTrend: function(data) {
    const numbers = data.filter(d => !isNaN(d) && d !== '').map(Number);
    
    if (numbers.length < 3) return null;
    
    // Simple linear regression
    const n = numbers.length;
    const indices = Array.from({length: n}, (_, i) => i);
    
    const sumX = indices.reduce((a, b) => a + b, 0);
    const sumY = numbers.reduce((a, b) => a + b, 0);
    const sumXY = indices.reduce((acc, x, i) => acc + x * numbers[i], 0);
    const sumX2 = indices.reduce((acc, x) => acc + x * x, 0);
    
    const slope = (n * sumXY - sumX * sumY) / (n * sumX2 - sumX * sumX);
    const intercept = (sumY - slope * sumX) / n;
    
    // Determine trend direction
    let direction = 'stable';
    if (Math.abs(slope) > 0.1) {
      direction = slope > 0 ? 'increasing' : 'decreasing';
    }
    
    // Calculate R-squared
    const yMean = sumY / n;
    const ssTotal = numbers.reduce((acc, y) => acc + Math.pow(y - yMean, 2), 0);
    const ssResidual = numbers.reduce((acc, y, i) => {
      const yPred = slope * i + intercept;
      return acc + Math.pow(y - yPred, 2);
    }, 0);
    const rSquared = 1 - (ssResidual / ssTotal);
    
    return {
      direction: direction,
      slope: slope,
      intercept: intercept,
      rSquared: rSquared,
      strength: rSquared > 0.7 ? 'strong' : rSquared > 0.4 ? 'moderate' : 'weak'
    };
  },

  /**
   * Generate slide structure based on analysis
   * @param {Object} analysis - Data analysis results
   * @param {Object} config - User configuration
   * @return {Object} Slide structure
   */
  generateSlideStructure: function(analysis, config) {
    const slides = [];
    const template = this.templates[config.template || 'corporate'];
    
    // Title slide
    slides.push({
      type: 'title',
      layout: template.layouts.title,
      title: config.presentationTitle || 'Data Analysis Presentation',
      subtitle: `Generated from ${SpreadsheetApp.getActiveSheet().getName()} | ${new Date().toLocaleDateString()}`
    });
    
    // Executive summary (if style is executive)
    if (config.presentationStyle === 'executive' || config.includeExecutiveSummary) {
      slides.push({
        type: 'summary',
        layout: template.layouts.content,
        title: 'Executive Summary',
        content: {
          type: 'bullets',
          items: analysis.insights.slice(0, 4)
        }
      });
    }
    
    // Key metrics slide
    if (analysis.keyMetrics.length > 0) {
      slides.push({
        type: 'metrics',
        layout: template.layouts.content,
        title: 'Key Metrics',
        content: {
          type: 'metrics',
          items: analysis.keyMetrics
        }
      });
    }
    
    // Data overview slide
    slides.push({
      type: 'overview',
      layout: template.layouts.content,
      title: 'Data Overview',
      content: {
        type: 'stats',
        items: this.formatDataOverview(analysis)
      }
    });
    
    // Chart slides based on recommendations
    const maxCharts = Math.min(analysis.chartRecommendations.length, Math.floor(config.slideCount * 0.5));
    for (let i = 0; i < maxCharts; i++) {
      const chart = analysis.chartRecommendations[i];
      slides.push({
        type: 'chart',
        layout: template.layouts.chart,
        title: chart.title,
        content: {
          type: 'chart',
          chartType: chart.type,
          data: chart.data,
          options: chart.options
        },
        notes: chart.insight || ''
      });
    }
    
    // Insights slides
    if (analysis.insights.length > 4) {
      const remainingInsights = analysis.insights.slice(4);
      const insightChunks = this.chunkArray(remainingInsights, 4);
      
      insightChunks.forEach((chunk, index) => {
        slides.push({
          type: 'insights',
          layout: template.layouts.content,
          title: `Key Findings ${insightChunks.length > 1 ? `(${index + 1})` : ''}`,
          content: {
            type: 'bullets',
            items: chunk
          }
        });
      });
    }
    
    // Trends slide (if detected)
    const trendsData = Object.entries(analysis.patterns).filter(([_, pattern]) => pattern.direction);
    if (trendsData.length > 0) {
      slides.push({
        type: 'trends',
        layout: template.layouts.content,
        title: 'Trend Analysis',
        content: {
          type: 'trends',
          items: trendsData.map(([column, pattern]) => ({
            column: column,
            trend: pattern.direction,
            strength: pattern.strength
          }))
        }
      });
    }
    
    // Recommendations slide
    slides.push({
      type: 'recommendations',
      layout: template.layouts.content,
      title: 'Recommendations',
      content: {
        type: 'bullets',
        items: this.generateRecommendations(analysis)
      }
    });
    
    // Next steps slide
    slides.push({
      type: 'next_steps',
      layout: template.layouts.content,
      title: 'Next Steps',
      content: {
        type: 'numbered',
        items: this.generateNextSteps(analysis)
      }
    });
    
    // Thank you slide
    slides.push({
      type: 'thank_you',
      layout: template.layouts.title,
      title: 'Thank You',
      subtitle: 'Questions?'
    });
    
    // Trim to requested slide count
    const finalSlides = slides.slice(0, config.slideCount || this.config.defaultSlides);
    
    return {
      slides: finalSlides,
      template: template,
      metadata: {
        totalSlides: finalSlides.length,
        hasCharts: finalSlides.some(s => s.type === 'chart'),
        hasInsights: finalSlides.some(s => s.type === 'insights')
      }
    };
  },

  /**
   * Create Google Slides presentation
   * @param {Object} structure - Slide structure
   * @param {Object} config - Configuration
   * @return {Object} Creation result
   */
  createGoogleSlides: function(structure, config) {
    try {
      // Create new presentation
      const presentation = SlidesApp.create(
        config.presentationTitle || 'CellPilot Analysis Presentation'
      );
      
      const presentationId = presentation.getId();
      const template = structure.template;
      
      // Remove default slide
      const slides = presentation.getSlides();
      if (slides.length > 0) {
        slides[0].remove();
      }
      
      // Create slides based on structure
      structure.slides.forEach((slideData, index) => {
        this.createSlide(presentation, slideData, index, template);
      });
      
      // Apply template styling
      this.applyTemplateToPresentation(presentation, template);
      
      const presentationUrl = presentation.getUrl();
      
      return {
        success: true,
        presentationId: presentationId,
        url: presentationUrl
      };
      
    } catch (error) {
      Logger.error('CellM8: Error creating Google Slides:', error);
      return {
        success: false,
        error: error.message
      };
    }
  },

  /**
   * Create individual slide
   * @param {Presentation} presentation - Google Slides presentation object
   * @param {Object} slideData - Slide configuration
   * @param {number} index - Slide index
   * @param {Object} template - Template configuration
   */
  createSlide: function(presentation, slideData, index, template) {
    try {
      let slide;
      
      // Create slide with appropriate layout
      const layoutName = slideData.layout || 'BLANK';
      const predefinedLayouts = presentation.getLayouts();
      const layout = predefinedLayouts.find(l => l.getLayoutName() === layoutName) || predefinedLayouts[0];
      
      slide = presentation.appendSlide(layout);
      
      // Handle different slide types
      switch (slideData.type) {
        case 'title':
          this.createTitleSlide(slide, slideData, template);
          break;
          
        case 'summary':
        case 'insights':
        case 'recommendations':
        case 'next_steps':
          this.createContentSlide(slide, slideData, template);
          break;
          
        case 'chart':
          this.createChartSlide(slide, slideData, template);
          break;
          
        case 'metrics':
          this.createMetricsSlide(slide, slideData, template);
          break;
          
        case 'overview':
        case 'trends':
          this.createDataSlide(slide, slideData, template);
          break;
          
        case 'thank_you':
          this.createThankYouSlide(slide, slideData, template);
          break;
          
        default:
          this.createContentSlide(slide, slideData, template);
      }
      
      // Add speaker notes if available
      if (slideData.notes) {
        slide.getNotesPage().getSpeakerNotesShape().getText().setText(slideData.notes);
      }
      
    } catch (error) {
      Logger.error('CellM8: Error creating slide:', error);
      // Continue with other slides even if one fails
    }
  },

  /**
   * Create title slide
   */
  createTitleSlide: function(slide, slideData, template) {
    const shapes = slide.getShapes();
    
    shapes.forEach(shape => {
      const text = shape.getText();
      const placeholder = shape.getPlaceholder();
      
      if (placeholder === SlidesApp.PlaceholderType.TITLE || 
          placeholder === SlidesApp.PlaceholderType.CENTERED_TITLE) {
        text.setText(slideData.title);
        text.getTextStyle()
          .setFontSize(40)
          .setForegroundColor(template.colorScheme.primary)
          .setFontFamily(template.fontFamily)
          .setBold(true);
      } else if (placeholder === SlidesApp.PlaceholderType.SUBTITLE) {
        text.setText(slideData.subtitle || '');
        text.getTextStyle()
          .setFontSize(20)
          .setForegroundColor(template.colorScheme.text)
          .setFontFamily(template.fontFamily);
      }
    });
  },

  /**
   * Create content slide with bullets or text
   */
  createContentSlide: function(slide, slideData, template) {
    const shapes = slide.getShapes();
    
    shapes.forEach(shape => {
      const text = shape.getText();
      const placeholder = shape.getPlaceholder();
      
      if (placeholder === SlidesApp.PlaceholderType.TITLE) {
        text.setText(slideData.title);
        text.getTextStyle()
          .setFontSize(30)
          .setForegroundColor(template.colorScheme.primary)
          .setFontFamily(template.fontFamily)
          .setBold(true);
      } else if (placeholder === SlidesApp.PlaceholderType.BODY) {
        if (slideData.content) {
          if (slideData.content.type === 'bullets') {
            const bulletText = slideData.content.items.map(item => '• ' + item).join('\n');
            text.setText(bulletText);
          } else if (slideData.content.type === 'numbered') {
            const numberedText = slideData.content.items.map((item, i) => `${i + 1}. ${item}`).join('\n');
            text.setText(numberedText);
          } else if (slideData.content.type === 'text') {
            text.setText(slideData.content.text);
          }
          
          text.getTextStyle()
            .setFontSize(18)
            .setForegroundColor(template.colorScheme.text)
            .setFontFamily(template.fontFamily);
        }
      }
    });
  },

  /**
   * Create chart slide
   */
  createChartSlide: function(slide, slideData, template) {
    try {
      // Add title
      const shapes = slide.getShapes();
      shapes.forEach(shape => {
        const placeholder = shape.getPlaceholder();
        if (placeholder === SlidesApp.PlaceholderType.TITLE) {
          const text = shape.getText();
          text.setText(slideData.title);
          text.getTextStyle()
            .setFontSize(30)
            .setForegroundColor(template.colorScheme.primary)
            .setFontFamily(template.fontFamily)
            .setBold(true);
        }
      });
      
      // Create chart
      if (slideData.content && slideData.content.data) {
        const chartData = slideData.content.data;
        const chartType = this.mapChartType(slideData.content.chartType);
        
        // Build data table for chart
        const dataTable = Charts.newDataTable();
        
        // Add headers
        if (chartData.headers) {
          chartData.headers.forEach((header, index) => {
            if (index === 0) {
              dataTable.addColumn(Charts.ColumnType.STRING, header);
            } else {
              dataTable.addColumn(Charts.ColumnType.NUMBER, header);
            }
          });
        }
        
        // Add data rows
        if (chartData.rows) {
          chartData.rows.forEach(row => {
            dataTable.addRow(row);
          });
        }
        
        // Create and configure chart
        const chart = slide.insertChart(chartType, 50, 100, 600, 350, dataTable.build());
        
        // Apply template colors to chart
        const chartOptions = chart.getOptions();
        chartOptions.setOption('colors', [
          template.colorScheme.primary,
          template.colorScheme.secondary,
          template.colorScheme.accent
        ]);
        chartOptions.setOption('fontName', template.fontFamily);
        chart.setOptions(chartOptions);
      }
      
    } catch (error) {
      Logger.error('CellM8: Error creating chart slide:', error);
      // Fall back to content slide
      this.createContentSlide(slide, {
        title: slideData.title,
        content: {
          type: 'text',
          text: 'Chart data visualization'
        }
      }, template);
    }
  },

  /**
   * Create metrics slide with KPIs
   */
  createMetricsSlide: function(slide, slideData, template) {
    // Add title
    const shapes = slide.getShapes();
    let titleSet = false;
    
    shapes.forEach(shape => {
      const placeholder = shape.getPlaceholder();
      if (!titleSet && placeholder === SlidesApp.PlaceholderType.TITLE) {
        const text = shape.getText();
        text.setText(slideData.title);
        text.getTextStyle()
          .setFontSize(30)
          .setForegroundColor(template.colorScheme.primary)
          .setFontFamily(template.fontFamily)
          .setBold(true);
        titleSet = true;
      }
    });
    
    // Add metrics as formatted text or shapes
    if (slideData.content && slideData.content.items) {
      const metrics = slideData.content.items;
      const startY = 150;
      const spacing = 80;
      
      metrics.forEach((metric, index) => {
        if (index < 4) { // Limit to 4 metrics per slide
          const x = (index % 2) * 300 + 50;
          const y = Math.floor(index / 2) * spacing + startY;
          
          // Create metric box
          const metricBox = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, x, y, 250, 60);
          metricBox.getFill().setSolidFill(template.colorScheme.primary, 0.1);
          metricBox.getBorder().setTransparent();
          
          // Add metric text
          const metricText = metricBox.getText();
          metricText.setText(`${metric.label}\n${metric.value}`);
          metricText.getTextStyle()
            .setFontSize(16)
            .setForegroundColor(template.colorScheme.text)
            .setFontFamily(template.fontFamily)
            .setBold(true);
          metricText.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
        }
      });
    }
  },

  /**
   * Apply template styling to entire presentation
   */
  applyTemplateToPresentation: function(presentation, template) {
    try {
      // Set presentation-wide properties if possible
      // Note: Google Apps Script has limited support for presentation-wide styling
      // We apply styling to individual slides instead
      
      const slides = presentation.getSlides();
      slides.forEach(slide => {
        // Set background color
        slide.getBackground().setSolidFill(template.colorScheme.background);
      });
      
    } catch (error) {
      Logger.error('CellM8: Error applying template:', error);
    }
  },

  /**
   * Map chart type string to SlidesApp chart type
   */
  mapChartType: function(type) {
    const mapping = {
      'COLUMN': SlidesApp.ChartType.COLUMN,
      'BAR': SlidesApp.ChartType.BAR,
      'LINE': SlidesApp.ChartType.LINE,
      'PIE': SlidesApp.ChartType.PIE,
      'AREA': SlidesApp.ChartType.AREA,
      'SCATTER': SlidesApp.ChartType.SCATTER
    };
    
    return mapping[type] || SlidesApp.ChartType.COLUMN;
  },

  /**
   * Recommend charts based on data analysis
   */
  recommendCharts: function(analysis) {
    const recommendations = [];
    
    // Find numerical columns for charts
    const numericalColumns = Object.entries(analysis.dataTypes)
      .filter(([_, type]) => type === 'number')
      .map(([column, _]) => column);
    
    const categoricalColumns = Object.entries(analysis.dataTypes)
      .filter(([_, type]) => type === 'category' || type === 'text')
      .map(([column, _]) => column);
    
    const dateColumns = Object.entries(analysis.dataTypes)
      .filter(([_, type]) => type === 'date')
      .map(([column, _]) => column);
    
    // Recommend column chart for categorical vs numerical
    if (categoricalColumns.length > 0 && numericalColumns.length > 0) {
      recommendations.push({
        type: 'COLUMN',
        title: `${numericalColumns[0]} by ${categoricalColumns[0]}`,
        data: {
          headers: [categoricalColumns[0], numericalColumns[0]],
          rows: [] // Will be populated with actual data
        },
        insight: `Comparison of ${numericalColumns[0]} across different ${categoricalColumns[0]} categories`
      });
    }
    
    // Recommend line chart for time series
    if (dateColumns.length > 0 && numericalColumns.length > 0) {
      recommendations.push({
        type: 'LINE',
        title: `${numericalColumns[0]} Over Time`,
        data: {
          headers: [dateColumns[0], numericalColumns[0]],
          rows: [] // Will be populated with actual data
        },
        insight: `Trend analysis of ${numericalColumns[0]} over time`
      });
    }
    
    // Recommend pie chart for distribution
    if (categoricalColumns.length > 0 && numericalColumns.length > 0) {
      recommendations.push({
        type: 'PIE',
        title: `Distribution of ${numericalColumns[0]}`,
        data: {
          headers: [categoricalColumns[0], numericalColumns[0]],
          rows: [] // Will be populated with actual data
        },
        insight: `Proportion breakdown of ${numericalColumns[0]} by ${categoricalColumns[0]}`
      });
    }
    
    // Recommend scatter plot for correlation
    if (numericalColumns.length >= 2) {
      recommendations.push({
        type: 'SCATTER',
        title: `${numericalColumns[0]} vs ${numericalColumns[1]}`,
        data: {
          headers: [numericalColumns[0], numericalColumns[1]],
          rows: [] // Will be populated with actual data
        },
        insight: `Correlation analysis between ${numericalColumns[0]} and ${numericalColumns[1]}`
      });
    }
    
    return recommendations;
  },

  /**
   * Generate insights from data analysis
   */
  generateInsights: function(data, analysis) {
    const insights = [];
    
    // Add statistical insights
    Object.entries(analysis.statistics).forEach(([column, stats]) => {
      if (stats) {
        insights.push(`${column} ranges from ${stats.min.toFixed(2)} to ${stats.max.toFixed(2)} with an average of ${stats.mean.toFixed(2)}`);
        
        if (stats.stdDev > stats.mean * 0.5) {
          insights.push(`High variability detected in ${column} (std dev: ${stats.stdDev.toFixed(2)})`);
        }
      }
    });
    
    // Add trend insights
    Object.entries(analysis.patterns).forEach(([column, pattern]) => {
      if (pattern && pattern.direction) {
        insights.push(`${column} shows a ${pattern.strength} ${pattern.direction} trend`);
      }
    });
    
    // Add data quality insights
    const totalCells = data.data.length * data.headers.length;
    const emptyCells = data.data.flat().filter(cell => cell === '' || cell === null).length;
    const completeness = ((totalCells - emptyCells) / totalCells * 100).toFixed(1);
    
    insights.push(`Data completeness: ${completeness}%`);
    
    if (completeness < 80) {
      insights.push('Consider filling missing data for more accurate analysis');
    }
    
    return insights;
  },

  /**
   * Generate recommendations based on analysis
   */
  generateRecommendations: function(analysis) {
    const recommendations = [];
    
    // Check for trends
    const trends = Object.entries(analysis.patterns).filter(([_, p]) => p && p.direction);
    if (trends.length > 0) {
      trends.forEach(([column, pattern]) => {
        if (pattern.direction === 'increasing' && pattern.strength === 'strong') {
          recommendations.push(`Continue monitoring ${column} for sustained growth`);
        } else if (pattern.direction === 'decreasing' && pattern.strength === 'strong') {
          recommendations.push(`Investigate factors causing decline in ${column}`);
        }
      });
    }
    
    // Check for outliers
    Object.entries(analysis.statistics).forEach(([column, stats]) => {
      if (stats && stats.stdDev > stats.mean * 0.5) {
        recommendations.push(`Review outliers in ${column} for data quality or exceptional cases`);
      }
    });
    
    // General recommendations
    recommendations.push('Schedule regular data reviews to track progress');
    recommendations.push('Consider setting up automated alerts for key metrics');
    
    return recommendations.slice(0, 5); // Limit to 5 recommendations
  },

  /**
   * Generate next steps
   */
  generateNextSteps: function(analysis) {
    const nextSteps = [
      'Review the presented data and insights with stakeholders',
      'Identify action items based on key findings',
      'Set up tracking for identified trends and patterns',
      'Schedule follow-up analysis in 30 days',
      'Share findings with relevant team members'
    ];
    
    return nextSteps;
  },

  /**
   * Extract key metrics from data
   */
  extractKeyMetrics: function(data, analysis) {
    const metrics = [];
    
    // Add key statistics as metrics
    Object.entries(analysis.statistics).forEach(([column, stats]) => {
      if (stats) {
        metrics.push({
          label: `${column} Average`,
          value: stats.mean.toFixed(2)
        });
        
        if (stats.sum > 1000) {
          metrics.push({
            label: `Total ${column}`,
            value: this.formatNumber(stats.sum)
          });
        }
      }
    });
    
    // Add record count
    metrics.push({
      label: 'Total Records',
      value: data.data.length.toString()
    });
    
    return metrics.slice(0, 6); // Limit to 6 key metrics
  },

  /**
   * Format data overview
   */
  formatDataOverview: function(analysis) {
    const overview = [];
    
    overview.push(`Data contains ${Object.keys(analysis.dataTypes).length} columns`);
    
    const dataTypesSummary = {};
    Object.values(analysis.dataTypes).forEach(type => {
      dataTypesSummary[type] = (dataTypesSummary[type] || 0) + 1;
    });
    
    Object.entries(dataTypesSummary).forEach(([type, count]) => {
      overview.push(`${count} ${type} column${count > 1 ? 's' : ''}`);
    });
    
    return overview;
  },

  /**
   * Clean data for presentation
   */
  cleanDataForPresentation: function(data, headers) {
    // Remove completely empty rows
    const cleanedData = data.filter(row => row.some(cell => cell !== '' && cell !== null));
    
    // Limit data size for performance
    const maxRows = 1000;
    const limitedData = cleanedData.slice(0, maxRows);
    
    return {
      headers: headers,
      data: limitedData,
      wasLimited: cleanedData.length > maxRows
    };
  },

  /**
   * Detect data types for columns
   */
  detectDataTypes: function(data) {
    const types = {};
    
    if (!data || data.length === 0) return types;
    
    const columnCount = data[0].length;
    
    for (let col = 0; col < columnCount; col++) {
      const columnData = data.map(row => row[col]).filter(val => val !== '' && val !== null);
      types[`column_${col}`] = this.detectColumnDataType(columnData);
    }
    
    return types;
  },

  /**
   * Detect data type for a single column
   */
  detectColumnDataType: function(columnData) {
    if (!columnData || columnData.length === 0) return 'unknown';
    
    const sample = columnData.slice(0, 100); // Sample first 100 values
    
    // Check if all are numbers
    const numberCount = sample.filter(val => !isNaN(val) && val !== '').length;
    if (numberCount / sample.length > 0.8) return 'number';
    
    // Check if all are dates
    const dateCount = sample.filter(val => {
      const date = new Date(val);
      return !isNaN(date.getTime()) && val.toString().match(/\d{1,2}[/-]\d{1,2}[/-]\d{2,4}/);
    }).length;
    if (dateCount / sample.length > 0.8) return 'date';
    
    // Check if categorical (limited unique values)
    const uniqueValues = [...new Set(sample)];
    if (uniqueValues.length < sample.length * 0.5) return 'category';
    
    return 'text';
  },

  /**
   * Generate default headers if not provided
   */
  generateDefaultHeaders: function(count) {
    return Array.from({length: count}, (_, i) => `Column ${String.fromCharCode(65 + i)}`);
  },

  /**
   * Analyze time series data
   */
  analyzeTimeSeries: function(dateData) {
    // Simplified time series analysis
    const dates = dateData.filter(d => d).map(d => new Date(d)).filter(d => !isNaN(d.getTime()));
    
    if (dates.length < 2) return null;
    
    dates.sort((a, b) => a - b);
    
    const intervals = [];
    for (let i = 1; i < dates.length; i++) {
      intervals.push(dates[i] - dates[i-1]);
    }
    
    const avgInterval = intervals.reduce((a, b) => a + b, 0) / intervals.length;
    
    return {
      startDate: dates[0],
      endDate: dates[dates.length - 1],
      dataPoints: dates.length,
      averageInterval: avgInterval,
      frequency: this.detectFrequency(avgInterval)
    };
  },

  /**
   * Detect data frequency
   */
  detectFrequency: function(avgInterval) {
    const day = 24 * 60 * 60 * 1000;
    
    if (avgInterval < day * 1.5) return 'daily';
    if (avgInterval < day * 8) return 'weekly';
    if (avgInterval < day * 35) return 'monthly';
    if (avgInterval < day * 100) return 'quarterly';
    return 'yearly';
  },

  /**
   * Analyze categorical data
   */
  analyzeCategories: function(categoryData) {
    const counts = {};
    
    categoryData.filter(c => c).forEach(category => {
      counts[category] = (counts[category] || 0) + 1;
    });
    
    const sorted = Object.entries(counts).sort((a, b) => b[1] - a[1]);
    
    return {
      uniqueCategories: sorted.length,
      topCategories: sorted.slice(0, 5),
      distribution: counts
    };
  },

  /**
   * Format large numbers
   */
  formatNumber: function(num) {
    if (num >= 1000000) {
      return (num / 1000000).toFixed(1) + 'M';
    } else if (num >= 1000) {
      return (num / 1000).toFixed(1) + 'K';
    }
    return num.toFixed(0);
  },

  /**
   * Chunk array into smaller arrays
   */
  chunkArray: function(array, size) {
    const chunks = [];
    for (let i = 0; i < array.length; i += size) {
      chunks.push(array.slice(i, i + size));
    }
    return chunks;
  },

  /**
   * Track presentation creation for ML learning
   */
  trackPresentationCreation: function(config, result) {
    try {
      // Store usage data for future ML improvements
      const usage = {
        timestamp: new Date(),
        config: {
          dataSource: config.dataSource.type,
          template: config.template,
          slideCount: config.slideCount,
          style: config.presentationStyle
        },
        result: {
          success: result.success,
          slideCount: result.slideCount
        }
      };
      
      // Save to user properties for learning
      const existingUsage = PropertiesService.getUserProperties().getProperty('cellm8_usage');
      const usageHistory = existingUsage ? JSON.parse(existingUsage) : [];
      usageHistory.push(usage);
      
      // Keep only last 50 entries
      if (usageHistory.length > 50) {
        usageHistory.shift();
      }
      
      PropertiesService.getUserProperties().setProperty('cellm8_usage', JSON.stringify(usageHistory));
      
    } catch (error) {
      Logger.error('CellM8: Error tracking usage:', error);
    }
  },

  /**
   * Create a thank you slide
   */
  createThankYouSlide: function(slide, slideData, template) {
    this.createTitleSlide(slide, slideData, template);
  },

  /**
   * Create a data slide with tables or stats
   */
  createDataSlide: function(slide, slideData, template) {
    this.createContentSlide(slide, slideData, template);
  }
};