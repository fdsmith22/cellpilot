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
  }
};