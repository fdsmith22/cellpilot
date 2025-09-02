/**
 * CellM8 - Spreadsheet to Presentation Converter
 * Simplified and rebuilt following CellPilot patterns
 */

const CellM8 = {
  // Configuration
  config: {
    maxSlides: 20,
    defaultSlides: 10,
    templates: ['simple', 'professional', 'modern']
  },

  /**
   * Initialize CellM8
   */
  initialize: function() {
    try {
      Logger.info('CellM8 initialized');
      return true;
    } catch (error) {
      Logger.error('Error initializing CellM8:', error);
      return false;
    }
  },

  /**
   * Get current sheet selection
   */
  getCurrentSelection: function() {
    try {
      const sheet = SpreadsheetApp.getActiveSheet();
      const range = sheet.getActiveRange();
      
      if (range) {
        return {
          success: true,
          range: range.getA1Notation(),
          rows: range.getNumRows(),
          cols: range.getNumColumns(),
          sheetName: sheet.getName()
        };
      }
      
      return {
        success: false,
        message: 'No range selected'
      };
    } catch (error) {
      Logger.error('Error getting selection:', error);
      return {
        success: false,
        error: error.toString()
      };
    }
  },

  /**
   * Create a simple presentation
   */
  createPresentation: function(config) {
    try {
      Logger.info('Creating presentation with config:', config);
      
      // Validate config
      if (!config || !config.title) {
        return {
          success: false,
          error: 'Title is required'
        };
      }

      // Create presentation using Slides API
      const presentation = SlidesApp.create(config.title);
      const presentationId = presentation.getId();
      
      // Get the presentation for editing
      const pres = SlidesApp.openById(presentationId);
      
      // Add a simple title slide
      const slides = pres.getSlides();
      const titleSlide = slides[0];
      
      // Get title and subtitle placeholders
      const titleShape = titleSlide.getPlaceholder(SlidesApp.PlaceholderType.TITLE);
      const subtitleShape = titleSlide.getPlaceholder(SlidesApp.PlaceholderType.SUBTITLE);
      
      if (titleShape) {
        titleShape.getText().setText(config.title);
      }
      
      if (subtitleShape && config.subtitle) {
        subtitleShape.getText().setText(config.subtitle);
      }
      
      // Add content slides based on config
      if (config.content && config.content.length > 0) {
        for (let i = 0; i < Math.min(config.content.length, 5); i++) {
          const slide = pres.appendSlide(SlidesApp.PredefinedLayout.TITLE_AND_BODY);
          const title = slide.getPlaceholder(SlidesApp.PlaceholderType.TITLE);
          const body = slide.getPlaceholder(SlidesApp.PlaceholderType.BODY);
          
          if (title) {
            title.getText().setText(config.content[i].title || `Slide ${i + 2}`);
          }
          
          if (body) {
            body.getText().setText(config.content[i].body || 'Content goes here');
          }
        }
      }
      
      // Get the URL
      const url = pres.getUrl();
      
      return {
        success: true,
        presentationId: presentationId,
        url: url,
        slideCount: pres.getSlides().length,
        message: 'Presentation created successfully'
      };
      
    } catch (error) {
      Logger.error('Error creating presentation:', error);
      return {
        success: false,
        error: error.toString()
      };
    }
  },

  /**
   * Extract data from the current sheet
   */
  extractSheetData: function() {
    try {
      const sheet = SpreadsheetApp.getActiveSheet();
      const range = sheet.getDataRange();
      const values = range.getValues();
      
      if (values.length === 0) {
        return {
          success: false,
          error: 'No data found in sheet'
        };
      }
      
      // Assume first row is headers
      const headers = values[0];
      const data = values.slice(1);
      
      return {
        success: true,
        headers: headers,
        data: data,
        rowCount: data.length,
        columnCount: headers.length
      };
      
    } catch (error) {
      Logger.error('Error extracting sheet data:', error);
      return {
        success: false,
        error: error.toString()
      };
    }
  },

  /**
   * Generate a simple preview
   */
  previewPresentation: function(config) {
    try {
      // Extract current data
      const dataResult = this.extractSheetData();
      
      if (!dataResult.success) {
        return dataResult;
      }
      
      // Create preview structure
      const preview = {
        success: true,
        title: config.title || 'Data Presentation',
        slideCount: Math.min(config.slideCount || 5, 10),
        dataInfo: {
          rows: dataResult.rowCount,
          columns: dataResult.columnCount
        },
        slides: []
      };
      
      // Add title slide
      preview.slides.push({
        type: 'title',
        title: preview.title,
        subtitle: `Analysis of ${dataResult.rowCount} rows of data`
      });
      
      // Add overview slide
      preview.slides.push({
        type: 'overview',
        title: 'Data Overview',
        content: `Dataset contains ${dataResult.columnCount} columns and ${dataResult.rowCount} rows`
      });
      
      // Add sample data slides
      for (let i = 0; i < Math.min(3, preview.slideCount - 2); i++) {
        preview.slides.push({
          type: 'data',
          title: `Data Analysis ${i + 1}`,
          content: `Analysis of ${dataResult.headers[i] || 'Column ' + (i + 1)}`
        });
      }
      
      return preview;
      
    } catch (error) {
      Logger.error('Error generating preview:', error);
      return {
        success: false,
        error: error.toString()
      };
    }
  },

  /**
   * Analyze data for presentation insights
   */
  analyzeData: function(data) {
    try {
      if (!data || !data.headers || !data.data) {
        return {
          success: false,
          error: 'Invalid data structure'
        };
      }

      const analysis = {
        totalRows: data.rowCount,
        totalColumns: data.columnCount,
        headers: data.headers,
        dataTypes: [],
        statistics: {},
        completeness: 0
      };

      // Detect data types for each column
      for (let col = 0; col < data.headers.length; col++) {
        const columnData = data.data.map(row => row[col]);
        const dataType = this.detectColumnType(columnData);
        analysis.dataTypes.push(dataType);
        
        // Calculate statistics for numeric columns
        if (dataType === 'number') {
          analysis.statistics[data.headers[col]] = this.calculateColumnStats(columnData);
        }
      }

      // Calculate completeness
      let totalCells = data.rowCount * data.columnCount;
      let filledCells = 0;
      for (let row of data.data) {
        for (let cell of row) {
          if (cell !== '' && cell !== null) filledCells++;
        }
      }
      analysis.completeness = Math.round((filledCells / totalCells) * 100);

      return {
        success: true,
        analysis: analysis
      };
    } catch (error) {
      Logger.error('Error analyzing data:', error);
      return {
        success: false,
        error: error.toString()
      };
    }
  },

  /**
   * Detect column data type
   */
  detectColumnType: function(columnData) {
    let numericCount = 0;
    let dateCount = 0;
    let textCount = 0;
    
    for (let value of columnData) {
      if (value === '' || value === null) continue;
      
      if (!isNaN(value) && typeof value === 'number') {
        numericCount++;
      } else if (value instanceof Date) {
        dateCount++;
      } else {
        textCount++;
      }
    }
    
    const total = numericCount + dateCount + textCount;
    if (numericCount > total * 0.7) return 'number';
    if (dateCount > total * 0.7) return 'date';
    return 'text';
  },

  /**
   * Calculate statistics for numeric column
   */
  calculateColumnStats: function(columnData) {
    const numbers = columnData.filter(v => v !== '' && v !== null && !isNaN(v));
    if (numbers.length === 0) return null;
    
    const sum = numbers.reduce((a, b) => a + Number(b), 0);
    const avg = sum / numbers.length;
    const min = Math.min(...numbers);
    const max = Math.max(...numbers);
    
    return {
      count: numbers.length,
      sum: sum,
      average: avg,
      min: min,
      max: max
    };
  },

  /**
   * Generate slides from data
   */
  generateSlides: function(presentation, data, config) {
    try {
      const slides = [];
      
      // Title slide
      slides.push({
        type: 'title',
        title: config.title || 'Data Presentation',
        subtitle: config.subtitle || `Generated from ${data.rowCount} rows of data`
      });
      
      // Overview slide
      slides.push({
        type: 'overview',
        title: 'Data Overview',
        bullets: [
          `Total Records: ${data.rowCount}`,
          `Data Fields: ${data.columnCount}`,
          `Data Completeness: ${Math.round((data.filledCells || 0) / (data.rowCount * data.columnCount) * 100)}%`
        ]
      });
      
      // Key metrics slide
      if (data.statistics) {
        const metrics = [];
        for (let key in data.statistics) {
          if (data.statistics[key]) {
            metrics.push(`${key}: Avg ${Math.round(data.statistics[key].average)}`);
          }
        }
        if (metrics.length > 0) {
          slides.push({
            type: 'metrics',
            title: 'Key Metrics',
            bullets: metrics.slice(0, 5)
          });
        }
      }
      
      // Data table slide (first 5 rows)
      if (data.data && data.data.length > 0) {
        slides.push({
          type: 'table',
          title: 'Sample Data',
          headers: data.headers,
          rows: data.data.slice(0, 5)
        });
      }
      
      // Summary slide
      slides.push({
        type: 'summary',
        title: 'Summary',
        bullets: [
          'Data successfully analyzed',
          `${slides.length - 1} insights generated`,
          'Ready for further analysis'
        ]
      });
      
      return {
        success: true,
        slides: slides
      };
    } catch (error) {
      Logger.error('Error generating slides:', error);
      return {
        success: false,
        error: error.toString()
      };
    }
  },

  /**
   * Apply template to presentation
   */
  applyTemplate: function(presentation, templateName) {
    try {
      const templates = {
        simple: {
          backgroundColor: '#FFFFFF',
          titleColor: '#000000',
          textColor: '#333333',
          accentColor: '#4285F4'
        },
        professional: {
          backgroundColor: '#F8F9FA',
          titleColor: '#1A73E8',
          textColor: '#202124',
          accentColor: '#1A73E8'
        },
        modern: {
          backgroundColor: '#1F1F1F',
          titleColor: '#FFFFFF',
          textColor: '#E8EAED',
          accentColor: '#8AB4F8'
        }
      };
      
      const template = templates[templateName] || templates.simple;
      
      // Apply template would modify presentation here
      // For now, just return the template config
      return {
        success: true,
        template: template,
        appliedTo: presentation.getId()
      };
    } catch (error) {
      Logger.error('Error applying template:', error);
      return {
        success: false,
        error: error.toString()
      };
    }
  },

  /**
   * Create chart from data
   */
  createChart: function(data, chartType) {
    try {
      // This would create actual charts in the presentation
      // For now, return chart configuration
      return {
        success: true,
        chartType: chartType || 'column',
        dataPoints: data.length,
        message: 'Chart configuration created'
      };
    } catch (error) {
      Logger.error('Error creating chart:', error);
      return {
        success: false,
        error: error.toString()
      };
    }
  },

  /**
   * Export presentation
   */
  exportPresentation: function(presentationId, format) {
    try {
      const presentation = SlidesApp.openById(presentationId);
      const url = presentation.getUrl();
      
      // Format options: pdf, pptx
      const exportUrl = url.replace(/\/edit.*$/, `/export/${format || 'pdf'}`);
      
      return {
        success: true,
        url: exportUrl,
        format: format || 'pdf'
      };
    } catch (error) {
      Logger.error('Error exporting presentation:', error);
      return {
        success: false,
        error: error.toString()
      };
    }
  },

  /**
   * Share presentation
   */
  sharePresentation: function(presentationId, emails, permission) {
    try {
      const presentation = DriveApp.getFileById(presentationId);
      
      // Permission can be 'view', 'comment', or 'edit'
      const access = permission === 'edit' ? DriveApp.Access.ANYONE_WITH_LINK : DriveApp.Access.ANYONE_WITH_LINK;
      const perm = permission === 'edit' ? DriveApp.Permission.EDIT : DriveApp.Permission.VIEW;
      
      presentation.setSharing(access, perm);
      
      // Add specific users if emails provided
      if (emails && emails.length > 0) {
        for (let email of emails) {
          if (permission === 'edit') {
            presentation.addEditor(email);
          } else {
            presentation.addViewer(email);
          }
        }
      }
      
      return {
        success: true,
        url: presentation.getUrl(),
        sharedWith: emails || ['Anyone with link']
      };
    } catch (error) {
      Logger.error('Error sharing presentation:', error);
      return {
        success: false,
        error: error.toString()
      };
    }
  }
};