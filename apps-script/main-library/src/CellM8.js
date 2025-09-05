/**
 * CellM8 - Presentation Generator Module
 * Creates professional presentations from spreadsheet data
 */

const CellM8 = {
  /**
   * Initialize CellM8
   */
  init: function () {
    Logger.log('CellM8 initialized');
    return { success: true };
  },

  /**
   * Show CellM8 sidebar
   */
  showSidebar: function () {
    try {
      const html = HtmlService.createTemplateFromFile('CellM8Template')
        .evaluate()
        .setTitle('CellM8 - Presentation Helper')
        .setWidth(400);

      SpreadsheetApp.getUi().showSidebar(html);

      return { success: true };
    } catch (error) {
      Logger.error('Error showing CellM8 sidebar:', error);
      return { success: false, error: error.toString() };
    }
  },

  /**
   * Test function to verify CellM8 is working
   */
  testFunction: function () {
    return {
      success: true,
      message: 'CellM8 is working properly',
      timestamp: new Date().toISOString(),
    };
  },

  /**
   * Get current selection info
   */
  getCurrentSelection: function () {
    try {
      const sheet = SpreadsheetApp.getActiveSheet();
      const range = sheet.getActiveRange();

      if (range) {
        return {
          success: true,
          hasSelection: true,
          range: range.getA1Notation(),
          rows: range.getNumRows(),
          cols: range.getNumColumns(),
          sheet: sheet.getName(),
        };
      } else {
        const dataRange = sheet.getDataRange();
        return {
          success: true,
          hasSelection: false,
          range: dataRange.getA1Notation(),
          rows: dataRange.getNumRows(),
          cols: dataRange.getNumColumns(),
          sheet: sheet.getName(),
        };
      }
    } catch (error) {
      Logger.error('Error getting selection:', error);
      return {
        success: false,
        error: error.toString(),
      };
    }
  },

  /**
   * Select entire data range
   */
  selectEntireDataRange: function () {
    try {
      const sheet = SpreadsheetApp.getActiveSheet();
      const dataRange = sheet.getDataRange();
      sheet.setActiveRange(dataRange);

      return {
        success: true,
        range: dataRange.getA1Notation(),
        rows: dataRange.getNumRows(),
        cols: dataRange.getNumColumns(),
      };
    } catch (error) {
      Logger.error('Error selecting data range:', error);
      return {
        success: false,
        error: error.toString(),
      };
    }
  },

  /**
   * Select specific range
   */
  selectRange: function (rangeA1) {
    try {
      const sheet = SpreadsheetApp.getActiveSheet();
      const range = sheet.getRange(rangeA1);
      sheet.setActiveRange(range);

      return {
        success: true,
        range: range.getA1Notation(),
        rows: range.getNumRows(),
        cols: range.getNumColumns(),
      };
    } catch (error) {
      Logger.error('Error selecting range:', error);
      return {
        success: false,
        error: error.toString(),
      };
    }
  },

  /**
   * Generate preview of presentation
   */
  previewPresentation: function (config) {
    try {
      // Extract data for preview
      const dataResult = this.extractSheetData();

      if (!dataResult.success) {
        return dataResult;
      }

      // Generate preview data
      const preview = {
        success: true,
        slideCount: config.slideCount || 5,
        dataInfo: {
          rows: dataResult.rowCount,
          columns: dataResult.columnCount,
          headers: dataResult.headers,
        },
        slides: [],
      };

      // Add preview slides
      preview.slides.push({
        type: 'title',
        title: config.title || 'Data Presentation',
        subtitle: config.subtitle || `Analysis of ${dataResult.rowCount} records`,
      });

      if (config.slideCount >= 3) {
        preview.slides.push({
          type: 'overview',
          title: 'Data Overview',
          content: `Dataset contains ${dataResult.rowCount} rows and ${dataResult.columnCount} columns`,
        });
      }

      return preview;
    } catch (error) {
      Logger.error('Error generating preview:', error);
      return {
        success: false,
        error: error.toString(),
      };
    }
  },

  /**
   * Create a presentation with data
   */
  createPresentation: function (config) {
    try {
      Logger.info('Creating presentation with config:', config);

      // Validate config
      if (!config || !config.title) {
        return {
          success: false,
          error: 'Title is required',
        };
      }

      // Extract data if not provided
      let dataResult = null;
      if (!config.data) {
        dataResult = this.extractSheetData();
        if (!dataResult.success) {
          return {
            success: false,
            error: 'Failed to extract data: ' + dataResult.error,
          };
        }
      } else {
        dataResult = config.data;
      }

      // Use the embedded SlideGenerator
      Logger.log('Using embedded CellM8.SlideGenerator');
      Logger.log('Config template selected: ' + config.template);

      // Validate dataResult structure
      if (!dataResult.success) {
        return {
          success: false,
          error: 'Failed to extract data from spreadsheet',
        };
      }

      if (!dataResult.data || !Array.isArray(dataResult.data)) {
        Logger.error('Invalid data structure from extractSheetData:', dataResult);
        return {
          success: false,
          error: 'Invalid data format: data must be an array',
        };
      }

      if (!dataResult.headers || dataResult.headers.length === 0) {
        Logger.error('No headers found in data:', dataResult);
        return {
          success: false,
          error: 'No headers found. First row must contain column headers.',
        };
      }

      Logger.log('Data validated. Headers:', dataResult.headers.join(', '));
      Logger.log('Data rows:', dataResult.data.length);

      const generatorResult = this.SlideGenerator.createPresentation(
        config.title || 'Data Presentation',
        dataResult.data, // Pass only the data array, headers go in config
        {
          title: config.title,
          subtitle: config.subtitle,
          headers: dataResult.headers, // Headers passed here
          slideCount: config.slideCount || 5,
          template: config.template || 'professional',
          masterTemplate: config.masterTemplate,
        }
      );

      Logger.log('Generator returned: ' + JSON.stringify(generatorResult));

      if (generatorResult.success) {
        Logger.log('CellM8.SlideGenerator SUCCEEDED');
        return generatorResult;
      } else {
        return {
          success: false,
          error: generatorResult.error || 'Failed to create presentation',
        };
      }
    } catch (error) {
      Logger.error('Error creating presentation:', error);
      return {
        success: false,
        error: error.toString(),
      };
    }
  },

  /**
   * Get current selection information for UI
   */
  getCellM8Selection: function () {
    try {
      const sheet = SpreadsheetApp.getActiveSheet();
      const selection = sheet.getActiveRange();

      if (selection) {
        return {
          success: true,
          hasSelection: true,
          range: selection.getA1Notation(),
          rows: selection.getNumRows(),
          cols: selection.getNumColumns(),
          sheet: sheet.getName(),
        };
      } else {
        const dataRange = sheet.getDataRange();
        return {
          success: true,
          hasSelection: false,
          range: dataRange.getA1Notation(),
          rows: dataRange.getNumRows(),
          cols: dataRange.getNumColumns(),
          sheet: sheet.getName(),
        };
      }
    } catch (error) {
      Logger.error('Failed to get selection:', error);
      return {
        success: false,
        error: error.toString(),
      };
    }
  },

  /**
   * Select entire data range
   */
  selectCellM8EntireDataRange: function () {
    try {
      const sheet = SpreadsheetApp.getActiveSheet();
      const dataRange = sheet.getDataRange();
      sheet.setActiveRange(dataRange);

      return {
        success: true,
        range: dataRange.getA1Notation(),
        rows: dataRange.getNumRows(),
        cols: dataRange.getNumColumns(),
      };
    } catch (error) {
      Logger.error('Failed to select entire data:', error);
      return {
        success: false,
        error: error.toString(),
      };
    }
  },

  /**
   * Select a specific range
   */
  selectCellM8Range: function (rangeA1) {
    try {
      const sheet = SpreadsheetApp.getActiveSheet();
      const range = sheet.getRange(rangeA1);
      sheet.setActiveRange(range);

      return {
        success: true,
        range: range.getA1Notation(),
        rows: range.getNumRows(),
        cols: range.getNumColumns(),
      };
    } catch (error) {
      Logger.error('Failed to select range:', error);
      return {
        success: false,
        error: error.toString(),
      };
    }
  },

  /**
   * Extract data from the current sheet or selection
   */
  extractSheetData: function (options) {
    try {
      const sheet = SpreadsheetApp.getActiveSheet();
      let range;

      // Check if specific range is requested
      if (options && options.rangeA1) {
        range = sheet.getRange(options.rangeA1);
      } else {
        // Try to get active range first, fall back to data range
        range = sheet.getActiveRange();
        if (!range || (range.getNumRows() === 1 && range.getNumColumns() === 1)) {
          range = sheet.getDataRange();
        }
      }

      const values = range.getValues();

      if (values.length === 0) {
        return {
          success: false,
          error: 'No data found in selected range',
        };
      }

      // Assume first row is headers
      const headers = values[0];
      const data = values.slice(1);

      // Filter out empty rows
      const filteredData = data.filter((row) => row.some((cell) => cell !== '' && cell !== null));

      return {
        success: true,
        headers: headers,
        data: filteredData,
        rowCount: filteredData.length,
        columnCount: headers.length,
        range: range.getA1Notation(),
        sheetName: sheet.getName(),
      };
    } catch (error) {
      Logger.error('Error extracting sheet data:', error);
      return {
        success: false,
        error: error.toString(),
      };
    }
  },

  /**
   * Analyze data for insights
   */
  analyzeData: function (dataResult) {
    try {
      if (!dataResult || !dataResult.data) {
        return {
          success: false,
          error: 'No data to analyze',
        };
      }

      const analysis = {
        totalRows: dataResult.rowCount,
        totalColumns: dataResult.columnCount,
        columnTypes: {},
        statistics: {},
        completeness: 0,
      };

      // Analyze each column
      dataResult.headers.forEach((header, index) => {
        const columnData = dataResult.data.map((row) => row[index]);
        const type = this.detectColumnType(columnData);

        analysis.columnTypes[header] = type;

        if (type === 'number') {
          analysis.statistics[header] = this.calculateColumnStats(columnData);
        }
      });

      // Calculate completeness
      const totalCells = dataResult.rowCount * dataResult.columnCount;
      let filledCells = 0;

      dataResult.data.forEach((row) => {
        row.forEach((cell) => {
          if (cell !== '' && cell !== null) {
            filledCells++;
          }
        });
      });

      analysis.completeness = Math.round((filledCells / totalCells) * 100);

      return {
        success: true,
        analysis: analysis,
      };
    } catch (error) {
      Logger.error('Error analyzing data:', error);
      return {
        success: false,
        error: error.toString(),
      };
    }
  },

  /**
   * Detect column data type
   */
  detectColumnType: function (columnData) {
    let numericCount = 0;
    let dateCount = 0;
    let textCount = 0;

    for (const value of columnData) {
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
  calculateColumnStats: function (columnData) {
    const numbers = columnData.filter((v) => v !== '' && v !== null && !isNaN(v));
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
      max: max,
    };
  },

  /**
   * Generate slides from data
   */
  generateSlides: function (presentation, data, config) {
    try {
      // This function would be implemented with actual slide generation logic
      // For now, it's a placeholder
      return {
        success: true,
        slideCount: config.slideCount || 5,
      };
    } catch (error) {
      Logger.error('Error generating slides:', error);
      return {
        success: false,
        error: error.toString(),
      };
    }
  },

  /**
   * Apply template to presentation (stub for now)
   */
  applyTemplate: function (presentation, templateName) {
    try {
      const templates = {
        simple: {
          backgroundColor: '#FFFFFF',
          titleColor: '#000000',
          textColor: '#333333',
          accentColor: '#4285F4',
        },
        professional: {
          backgroundColor: '#F8F9FA',
          titleColor: '#1A73E8',
          textColor: '#202124',
          accentColor: '#1A73E8',
        },
        modern: {
          backgroundColor: '#1F1F1F',
          titleColor: '#FFFFFF',
          textColor: '#E8EAED',
          accentColor: '#8AB4F8',
        },
      };

      const template = templates[templateName] || templates.simple;

      // Apply template would modify presentation here
      // For now, just return the template config
      return {
        success: true,
        template: template,
        message: 'Template functionality coming soon',
      };
    } catch (error) {
      Logger.error('Error applying template:', error);
      return {
        success: false,
        error: error.toString(),
      };
    }
  },

  /**
   * Create chart from data
   */
  createChart: function (data, chartType) {
    try {
      // This would create actual charts in the presentation
      // For now, return chart configuration
      return {
        success: true,
        chartType: chartType,
        dataPoints: data.length,
      };
    } catch (error) {
      Logger.error('Error creating chart:', error);
      return {
        success: false,
        error: error.toString(),
      };
    }
  },

  /**
   * Export presentation
   */
  exportPresentation: function (presentationId, format) {
    try {
      // This would export the presentation to different formats
      return {
        success: true,
        format: format,
        message: 'Export functionality coming soon',
      };
    } catch (error) {
      Logger.error('Error exporting presentation:', error);
      return {
        success: false,
        error: error.toString(),
      };
    }
  },

  /**
   * Share presentation
   */
  sharePresentation: function (presentationId, emails, permission) {
    try {
      // This would share the presentation
      return {
        success: true,
        sharedWith: emails,
        permission: permission,
        message: 'Share functionality coming soon',
      };
    } catch (error) {
      Logger.error('Error sharing presentation:', error);
      return {
        success: false,
        error: error.toString(),
      };
    }
  },

  /**
   * Embedded SlideGenerator - Professional presentation generator with themes and effects
   */
  SlideGenerator: {
    // Version stamp to force cache refresh
    VERSION: 'v3.0.1-no-transparency',
    LAST_UPDATED: new Date().toISOString(),

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
      right: 60,
    },

    // Typography scale (based on golden ratio)
    TYPOGRAPHY: {
      display: 48, // Main titles
      headline: 32, // Section titles
      title: 24, // Slide titles
      subtitle: 18, // Subtitles
      body: 14, // Body text
      caption: 11, // Captions
      overline: 10, // Small labels
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
          subtle: ['#ffffff', '#f8f9fa'],
        },
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
          subtle: ['#fafbfc', '#f0f7ff'],
        },
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
          subtle: ['#303134', '#202124'],
        },
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
          subtle: ['#ffffff', '#fff5f5'],
        },
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
          subtle: ['#ecf0f1', '#bdc3c7'],
        },
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
          subtle: ['#f7fafc', '#e2e8f0'],
        },
      },
    },

    // Enhanced Professional Theme Presets
    PROFESSIONAL_THEMES: {
      executive: {
        name: 'Executive Suite',
        primary: '#1B2951', // Deep professional blue
        primaryVariant: '#2A3F6F',
        secondary: '#4A6FA5', // Lighter blue
        accent: '#4285F4', // Google blue accent
        background: '#FFFFFF',
        surface: '#F7F9FC', // Very subtle blue-gray
        text: '#2C3E50',
        textLight: '#5F6368', // Google gray
        borderColor: '#DFE3E8',
        chartColors: ['#4285F4', '#34A853', '#FBBC04', '#EA4335', '#9B51E0', '#00ACC1'],
        fontTitle: 'Georgia',
        fontBody: 'Arial',
        chartStyle: 'clean',
        gridlines: '#E8EAED',
        gradients: {
          primary: ['#F7F9FC', '#E3E9F3'], // Subtle blue gradient
          overlay: ['rgba(66, 133, 244, 0.03)', 'rgba(27, 41, 81, 0.02)'],
        },
        darkMode: false,
      },
      modern: {
        name: 'Modern Tech',
        primary: '#5E4FDB', // Softer purple
        primaryVariant: '#7C6FE8',
        secondary: '#9B8DF1', // Lighter purple
        accent: '#00BCD4', // Cyan accent
        background: '#FFFFFF',
        surface: '#F9F7FF', // Very subtle purple tint
        text: '#2D3436',
        textLight: '#636E72',
        borderColor: '#E4E0F7',
        chartColors: ['#5E4FDB', '#00BCD4', '#9B8DF1', '#FFB74D', '#FF7979', '#81C784'],
        fontTitle: 'Roboto',
        fontBody: 'Roboto',
        chartStyle: 'modern',
        gridlines: '#F0F3F7',
        gradients: {
          primary: ['#F9F7FF', '#EDE9FF'], // Subtle purple gradient
          overlay: ['rgba(94, 79, 219, 0.04)', 'rgba(0, 188, 212, 0.03)'],
        },
        darkMode: false,
      },
      elegant: {
        name: 'Elegant Minimal',
        primary: '#2E3440', // Nord Polar Night
        primaryVariant: '#3B4252',
        secondary: '#88C0D0', // Nord Frost
        accent: '#BF616A', // Nord Aurora Red
        background: '#FFFFFF',
        surface: '#FAFBFC', // Very light gray
        text: '#2E3440',
        textLight: '#4C566A',
        borderColor: '#E5E9EC',
        chartColors: ['#88C0D0', '#BF616A', '#A3BE8C', '#D08770', '#B48EAD', '#5E81AC'],
        fontTitle: 'Helvetica',
        fontBody: 'Helvetica',
        chartStyle: 'minimal',
        gridlines: '#ECEFF4',
        gradients: {
          primary: ['#FAFBFC', '#F0F2F5'], // Subtle gray gradient
          overlay: ['rgba(46, 52, 64, 0.02)', 'rgba(136, 192, 208, 0.02)'],
        },
        darkMode: false,
      },
      professional: {
        name: 'Professional Blue',
        primary: '#003D82', // Deep corporate blue
        primaryVariant: '#005AAA',
        secondary: '#0078D4', // Microsoft blue
        accent: '#40E0D0', // Turquoise accent
        background: '#FFFFFF',
        surface: '#F3F6FB',
        text: '#323130',
        textLight: '#605E5C',
        borderColor: '#D9DCE0',
        chartColors: ['#0078D4', '#40E0D0', '#003D82', '#107C10', '#FFB900', '#E74856'],
        fontTitle: 'Georgia',
        fontBody: 'Segoe UI',
        chartStyle: 'professional',
        gridlines: '#EDEBE9',
        gradients: {
          primary: ['#F3F6FB', '#E1E9F4'], // Professional blue gradient
          overlay: ['rgba(0, 120, 212, 0.03)', 'rgba(0, 61, 130, 0.02)'],
        },
        darkMode: false,
      },
      dark: {
        name: 'Dark Professional',
        primary: '#BB86FC', // Material purple
        primaryVariant: '#9965F4',
        secondary: '#03DAC6', // Material teal
        accent: '#CF6679', // Material pink
        background: '#121212',
        surface: '#1E1E1E',
        text: '#FFFFFF',
        textLight: '#B0B0B0',
        borderColor: '#2C2C2C',
        chartColors: ['#BB86FC', '#03DAC6', '#CF6679', '#FFB86C', '#50FA7B', '#FF79C6'],
        fontTitle: 'Roboto',
        fontBody: 'Roboto',
        chartStyle: 'modern',
        darkMode: true,
        gridlines: '#2C2C2C',
        gradients: {
          primary: ['#1A1A1A', '#242424'], // Dark gradient
          overlay: ['rgba(187, 134, 252, 0.08)', 'rgba(3, 218, 198, 0.06)'],
        },
      },
    },

    /**
     * Apply professional theme to entire presentation
     */
    applyProfessionalTheme: function (presentation, themeName) {
      const theme = this.PROFESSIONAL_THEMES[themeName] || this.PROFESSIONAL_THEMES.executive;
      const slides = presentation.getSlides();

      Logger.log('Applying professional theme: ' + theme.name);

      slides.forEach((slide, index) => {
        try {
          // Apply background
          if (theme.darkMode) {
            this.applyDarkModeBackground(slide, theme);
          } else {
            this.applyLightBackground(slide, theme);
          }

          // Update all text elements
          this.updateSlideTextStyles(slide, theme);

          // Update charts with theme colors
          this.updateSlideCharts(slide, theme);

          // Update tables
          this.updateSlideTables(slide, theme);
        } catch (error) {
          Logger.error('Error applying theme to slide ' + index + ': ' + error);
        }
      });

      return {
        success: true,
        theme: theme.name,
        slidesUpdated: slides.length,
      };
    },

    /**
     * Apply dark mode background
     */
    applyDarkModeBackground: function (slide, theme) {
      const bg = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 0, 0, this.SLIDE_WIDTH, this.SLIDE_HEIGHT);
      bg.getFill().setSolidFill(theme.background);
      bg.getBorder().setTransparent();
      bg.sendToBack();
    },

    /**
     * Apply light background with subtle gradient
     */
    applyLightBackground: function (slide, theme) {
      const bg = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 0, 0, this.SLIDE_WIDTH, this.SLIDE_HEIGHT);
      bg.getFill().setSolidFill(theme.background);
      bg.getBorder().setTransparent();
      bg.sendToBack();

      // Add subtle gradient overlay if not minimal theme
      if (theme.chartStyle !== 'minimal') {
        const overlay = slide.insertShape(
          SlidesApp.ShapeType.RECTANGLE,
          0,
          this.SLIDE_HEIGHT * 0.7,
          this.SLIDE_WIDTH,
          this.SLIDE_HEIGHT * 0.3
        );
        overlay.getFill().setSolidFill(theme.surface);
        overlay.getBorder().setTransparent();
        overlay.sendToBack();
      }
    },

    /**
     * Update all text styles on a slide
     */
    updateSlideTextStyles: function (slide, theme) {
      const elements = slide.getPageElements();

      elements.forEach((element) => {
        if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
          const shape = element.asShape();
          const text = shape.getText();

          if (text && text.asString().length > 0) {
            const isTitle = this.isLikelyTitle(text.asString());

            text
              .getTextStyle()
              .setFontFamily(isTitle ? theme.fontTitle : theme.fontBody)
              .setForegroundColor(theme.text);

            if (isTitle) {
              text.getTextStyle().setBold(true);
            }
          }
        }
      });
    },

    /**
     * Check if text is likely a title
     */
    isLikelyTitle: function (text) {
      return text.length < 100 && (text === text.toUpperCase() || text.split(' ').length < 8);
    },

    /**
     * Update chart colors on slide
     */
    updateSlideCharts: function (slide, theme) {
      const charts = slide.getSheetsCharts();

      // Note: We can't directly modify embedded chart colors
      // But we can track them for refresh with new colors
      charts.forEach((chart) => {
        try {
          // Store theme preference for this chart
          Logger.log('Chart found - would apply colors: ' + theme.chartColors.join(', '));
        } catch (e) {
          Logger.log('Could not update chart: ' + e);
        }
      });
    },

    /**
     * Update table styling
     */
    updateSlideTables: function (slide, theme) {
      const tables = slide.getTables();

      tables.forEach((table) => {
        const numRows = table.getNumRows();
        const numCols = table.getNumColumns();

        // Style header row
        for (let col = 0; col < numCols; col++) {
          const headerCell = table.getCell(0, col);
          headerCell.getFill().setSolidFill(theme.primary);
          headerCell
            .getText()
            .getTextStyle()
            .setForegroundColor(theme.darkMode ? '#FFFFFF' : theme.text || '#2C3E50')
            .setFontFamily(theme.fontTitle)
            .setBold(true);
        }

        // Style data rows
        for (let row = 1; row < numRows; row++) {
          for (let col = 0; col < numCols; col++) {
            const cell = table.getCell(row, col);

            // Alternate row colors
            const bgColor = row % 2 === 0 ? theme.background : theme.surface;
            cell.getFill().setSolidFill(bgColor);

            cell.getText().getTextStyle().setForegroundColor(theme.text).setFontFamily(theme.fontBody);
          }
        }
      });
    },

    /**
     * Main entry point - Create presentation using optimal approach
     * REQUIRES: data must be an array of arrays (2D array from spreadsheet)
     * config.headers must contain the header row if not included in data
     */
    createPresentation: function (title, data, config) {
      try {
        Logger.log('Starting optimal presentation creation');
        Logger.log('Config received:', JSON.stringify(config));
        Logger.log('Data type received:', typeof data);
        Logger.log('Data is array:', Array.isArray(data));

        // Step 1: Create new presentation
        const presentation = this.createCleanPresentation(title);

        // Step 2: Validate and prepare data - SINGLE FORMAT ONLY
        if (!Array.isArray(data)) {
          throw new Error('Data must be an array. Received: ' + typeof data);
        }

        if (data.length === 0) {
          throw new Error('Data array is empty. Cannot create presentation.');
        }

        // Data structure MUST have headers and data arrays
        const dataStructure = {
          headers: config.headers || [],
          data: data,
        };

        // Validate we have headers
        if (dataStructure.headers.length === 0) {
          throw new Error('Headers are required in config.headers');
        }

        Logger.log(
          'Data structure prepared. Headers:',
          dataStructure.headers.length,
          'Rows:',
          dataStructure.data.length
        );

        // Step 3: Analyze data intelligently
        const analysis = this.analyzeDataIntelligently(dataStructure);
        Logger.log('Data analysis complete');

        // Step 3: Plan slide structure based on data
        const slidePlan = this.planSlideStructure(analysis, config);
        Logger.log('Slide plan created');

        // Step 4: Build slides according to plan
        this.buildSlidesFromPlan(presentation, slidePlan, dataStructure, analysis, config);

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
          message: 'Presentation created successfully with themes and effects',
        };
      } catch (error) {
        Logger.error('Error in optimal presentation creation:', error);
        return {
          success: false,
          error: error.toString(),
        };
      }
    },

    /**
     * Create truly clean presentation using three-request method
     */
    createCleanPresentation: function (title) {
      Logger.log('Creating clean presentation: ' + title);

      // Create presentation
      const presentation = SlidesApp.create(title);

      // Add a new blank slide
      const blankSlide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);

      // Remove the original slide with placeholders
      const slides = presentation.getSlides();
      if (slides.length > 1) {
        slides[0].remove();
        Logger.log('Removed default first slide');
      }

      // Clear any remaining content
      this.ensureSlideIsClean(blankSlide);

      return presentation;
    },

    /**
     * Ensure slide is completely clean
     */
    ensureSlideIsClean: function (slide) {
      const elements = slide.getPageElements();

      // Remove elements in reverse
      for (let i = elements.length - 1; i >= 0; i--) {
        try {
          const element = elements[i];
          if (this.shouldRemoveElement(element)) {
            element.remove();
          }
        } catch (e) {
          // Continue if element can't be removed
        }
      }
    },

    /**
     * Determine if element should be removed
     */
    shouldRemoveElement: function (element) {
      try {
        if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
          const shape = element.asShape();
          try {
            shape.getPlaceholderType();
            return true; // Has placeholder
          } catch (e) {
            const text = shape.getText().asString().toLowerCase();
            return (
              text.includes('click to add') ||
              text.includes('title') ||
              text.includes('subtitle') ||
              text.includes('body') ||
              text.trim() === ''
            );
          }
        }
      } catch (e) {
        return false;
      }
      return false;
    },

    /**
     * Analyze data intelligently
     */
    analyzeDataIntelligently: function (data) {
      const analysis = {
        totalRows: data.data.length,
        totalColumns: data.headers.length,
        dataTypes: {},
        numericColumns: [],
        textColumns: [],
        dateColumns: [],
        percentageColumns: [],
        categoryColumns: [],
        idColumns: [],
        relationships: [],
        keyMetrics: [],
        bestChartType: null,
        dataSummary: null,
        insights: [],
        trends: [],
        headers: data.headers || [],
        suggestedVisualizations: [],
      };

      // Analyze each column with semantic understanding
      data.headers.forEach((header, colIndex) => {
        const columnData = data.data.map((row) => row[colIndex]);
        const dataType = this.detectColumnType(columnData);
        const semanticType = this.detectSemanticType(header, columnData);

        analysis.dataTypes[header] = dataType;

        // Store semantic understanding
        if (semanticType) {
          if (semanticType.isCategory) {
            analysis.categoryColumns.push({
              name: header,
              index: colIndex,
              uniqueValues: [...new Set(columnData)].length,
            });
          }
          if (semanticType.isIdentifier) {
            analysis.idColumns.push(colIndex);
          }
          if (semanticType.isMoney) {
            analysis.keyMetrics.push({
              label: header,
              type: 'monetary',
              index: colIndex,
            });
          }
          if (semanticType.isPerformance) {
            analysis.keyMetrics.push({
              label: header,
              type: 'performance',
              index: colIndex,
            });
          }
        }

        if (dataType === 'numeric') {
          const stats = this.calculateStatistics(columnData);
          analysis.numericColumns.push({
            name: header,
            index: colIndex,
            stats: stats,
          });

          // Check if it's a percentage column
          const isPercentage = columnData.some((val) => {
            if (typeof val === 'string' && val.includes('%')) return true;
            const num = parseFloat(val);
            return !isNaN(num) && num >= 0 && num <= 100;
          });

          if (isPercentage && stats.max <= 100) {
            analysis.percentageColumns.push(colIndex);
          }

          if (stats.average !== null) {
            analysis.keyMetrics.push({
              label: header,
              value: Math.round(stats.average),
              type: 'average',
            });
          }
        } else if (dataType === 'date') {
          analysis.dateColumns.push({
            name: header,
            index: colIndex,
          });
        } else {
          analysis.textColumns.push({
            name: header,
            index: colIndex,
            uniqueValues: [...new Set(columnData)].length,
          });
        }
      });

      // Detect relationships between columns
      analysis.relationships = this.detectRelationships(data.headers, data.data, analysis);

      // Suggest visualizations based on relationships and data patterns
      analysis.suggestedVisualizations = this.suggestVisualizations(analysis, analysis.relationships);

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
    detectColumnType: function (columnData) {
      const nonEmpty = columnData.filter((val) => val != null && val !== '');
      if (nonEmpty.length === 0) return 'empty';

      let numericCount = 0;
      let dateCount = 0;

      nonEmpty.forEach((val) => {
        if (!isNaN(parseFloat(val)) && isFinite(val)) {
          numericCount++;
        } else if (this.isValidDate(val)) {
          dateCount++;
        }
      });

      const threshold = nonEmpty.length * 0.8;

      if (numericCount >= threshold) return 'numeric';
      if (dateCount >= threshold) return 'date';
      return 'text';
    },

    /**
     * Calculate statistics for numeric column
     */
    calculateStatistics: function (columnData) {
      const numbers = columnData
        .filter((val) => val != null && val !== '')
        .map((val) => parseFloat(val))
        .filter((val) => !isNaN(val) && isFinite(val));

      if (numbers.length === 0) {
        return { average: null, min: null, max: null, sum: null };
      }

      return {
        average: numbers.reduce((a, b) => a + b, 0) / numbers.length,
        min: Math.min(...numbers),
        max: Math.max(...numbers),
        sum: numbers.reduce((a, b) => a + b, 0),
        count: numbers.length,
      };
    },

    /**
     * Check if value is a valid date
     */
    isValidDate: function (value) {
      if (!value) return false;
      const date = new Date(value);
      return (
        date instanceof Date &&
        !isNaN(date) &&
        value &&
        value.toString &&
        (value.toString().includes('/') || value.toString().includes('-') || value.toString().match(/\d{4}/))
      );
    },

    /**
     * Detect semantic type of a column based on header and data
     */
    detectSemanticType: function (header, columnData) {
      const headerLower = header.toLowerCase();
      const result = {
        isCategory: false,
        isIdentifier: false,
        isMoney: false,
        isPerformance: false,
        isLocation: false,
        isPerson: false,
        isStatus: false,
        isDate: false,
        semanticLabel: null,
      };

      // Detect money columns
      if (
        headerLower.match(
          /salary|revenue|cost|price|amount|payment|budget|expense|income|profit|loss|fee|charge|total|subtotal|value|worth/
        )
      ) {
        result.isMoney = true;
        result.semanticLabel = 'monetary';
      }

      // Detect performance/score columns
      if (
        headerLower.match(
          /score|rating|performance|grade|rank|percentile|efficiency|accuracy|quality|satisfaction|metric|index|kpi/
        )
      ) {
        result.isPerformance = true;
        result.semanticLabel = 'performance';
      }

      // Detect person/employee columns
      if (
        headerLower.match(
          /employee|person|name|user|manager|supervisor|staff|worker|member|owner|customer|client|contact/
        )
      ) {
        result.isPerson = true;
        result.semanticLabel = 'person';
      }

      // Detect location columns
      if (
        headerLower.match(
          /location|city|state|country|region|area|zone|district|address|site|office|branch|store|department|division/
        )
      ) {
        result.isLocation = true;
        result.isCategory = true;
        result.semanticLabel = 'location';
      }

      // Detect status columns
      if (headerLower.match(/status|stage|phase|state|condition|flag|active|enabled|completed|approved|pending/)) {
        result.isStatus = true;
        result.isCategory = true;
        result.semanticLabel = 'status';
      }

      // Detect identifier columns
      if (headerLower.match(/\bid\b|code|number|ref|reference|key|identifier|\bno\b|serial/)) {
        result.isIdentifier = true;
        result.semanticLabel = 'identifier';
      }

      // Detect date columns
      if (
        headerLower.match(/date|time|year|month|day|created|updated|modified|due|deadline|start|end|period|quarter/)
      ) {
        result.isDate = true;
        result.semanticLabel = 'temporal';
      }

      // Check for categorical data by uniqueness
      if (!result.isIdentifier && !result.isMoney && !result.isPerformance) {
        const uniqueValues = [...new Set(columnData.filter((v) => v !== null && v !== ''))];
        const uniqueRatio = uniqueValues.length / columnData.length;

        // If less than 30% unique values, likely categorical
        if (uniqueRatio < 0.3 && uniqueValues.length < 20) {
          result.isCategory = true;
          if (!result.semanticLabel) {
            result.semanticLabel = 'category';
          }
        }
      }

      return result;
    },

    /**
     * Detect relationships between columns
     */
    detectRelationships: function (headers, data, analysis) {
      const relationships = [];

      // Find ALL meaningful category-numeric relationships
      if (analysis.categoryColumns.length > 0 && analysis.numericColumns.length > 0) {
        // Create relationships for multiple combinations, not just the first
        for (let catIdx = 0; catIdx < Math.min(analysis.categoryColumns.length, 3); catIdx++) {
          for (let numIdx = 0; numIdx < Math.min(analysis.numericColumns.length, 3); numIdx++) {
            const category = analysis.categoryColumns[catIdx];
            const numeric = analysis.numericColumns[numIdx];

            relationships.push({
              type: 'aggregation',
              groupBy: category.name,
              measure: numeric.name,
              description: `${numeric.name} by ${category.name}`,
              priority: catIdx === 0 && numIdx === 0 ? 1 : 2,
            });
          }
        }
      }

      // Add scatter plot relationships for numeric-numeric comparisons
      if (analysis.numericColumns.length >= 2) {
        for (let i = 0; i < Math.min(analysis.numericColumns.length - 1, 2); i++) {
          for (let j = i + 1; j < Math.min(analysis.numericColumns.length, 3); j++) {
            relationships.push({
              type: 'scatter',
              xColumn: analysis.numericColumns[i].name,
              yColumn: analysis.numericColumns[j].name,
              description: `${analysis.numericColumns[i].name} vs ${analysis.numericColumns[j].name}`,
            });
          }
        }
      }

      // Look for person-location-metric patterns
      const personCol = headers.find((h) => this.detectSemanticType(h, []).isPerson);
      const locationCol = headers.find((h) => this.detectSemanticType(h, []).isLocation);
      const moneyCol = headers.find((h) => this.detectSemanticType(h, []).isMoney);
      const performanceCol = headers.find((h) => this.detectSemanticType(h, []).isPerformance);

      if (personCol && locationCol && (moneyCol || performanceCol)) {
        relationships.push({
          type: 'hierarchy',
          levels: [personCol, locationCol],
          metric: moneyCol || performanceCol,
          description: `${moneyCol || performanceCol} by ${personCol} and ${locationCol}`,
        });
      }

      // Look for time series patterns
      if (analysis.dateColumns.length > 0 && analysis.numericColumns.length > 0) {
        relationships.push({
          type: 'timeseries',
          timeColumn: analysis.dateColumns[0].name,
          valueColumns: analysis.numericColumns.map((c) => c.name),
          description: 'Trends over time',
        });
      }

      // Look for distribution patterns
      if (analysis.categoryColumns.length > 0 && analysis.percentageColumns.length > 0) {
        relationships.push({
          type: 'distribution',
          category: analysis.categoryColumns[0].name,
          percentage: headers[analysis.percentageColumns[0]],
          description: 'Distribution analysis',
        });
      }

      return relationships;
    },

    /**
     * Suggest visualizations based on data patterns
     */
    suggestVisualizations: function (analysis, relationships) {
      const suggestions = [];

      // Check relationships first for more intelligent suggestions
      relationships.forEach((rel) => {
        switch (rel.type) {
        case 'aggregation':
          suggestions.push({
            type: 'bar_chart',
            title: rel.description,
            priority: 1,
            groupBy: rel.groupBy,
            measure: rel.measure,
            reason: `Show ${rel.measure} comparison across different ${rel.groupBy}`,
          });
          break;

        case 'hierarchy':
          suggestions.push({
            type: 'grouped_bar',
            title: rel.description,
            priority: 1,
            levels: rel.levels,
            metric: rel.metric,
            reason: `Analyze ${rel.metric} across multiple dimensions`,
          });
          break;

        case 'timeseries':
          suggestions.push({
            type: 'line_chart',
            title: 'Trend Analysis',
            priority: 1,
            timeColumn: rel.timeColumn,
            valueColumns: rel.valueColumns,
            reason: 'Show changes and patterns over time',
          });
          break;

        case 'distribution':
          suggestions.push({
            type: 'pie_chart',
            title: 'Distribution Breakdown',
            priority: 2,
            category: rel.category,
            value: rel.percentage,
            reason: 'Visualize proportions and percentages',
          });
          break;
        }
      });

      // Add fallback suggestions based on data types
      if (suggestions.length === 0) {
        if (analysis.numericColumns.length > 0 && analysis.categoryColumns.length > 0) {
          suggestions.push({
            type: 'bar_chart',
            title: 'Category Comparison',
            priority: 2,
            reason: 'Compare values across categories',
          });
        }

        if (analysis.numericColumns.length >= 2) {
          suggestions.push({
            type: 'scatter_plot',
            title: 'Correlation Analysis',
            priority: 3,
            reason: 'Explore relationships between numeric variables',
          });
        }

        if (analysis.keyMetrics.length > 0) {
          suggestions.push({
            type: 'metric_cards',
            title: 'Key Performance Indicators',
            priority: 1,
            metrics: analysis.keyMetrics,
            reason: 'Highlight important metrics and KPIs',
          });
        }
      }

      // Sort by priority
      return suggestions.sort((a, b) => a.priority - b.priority);
    },

    /**
     * Determine best chart type based on data
     */
    determineBestChartType: function (analysis) {
      const hasTime = analysis.dateColumns.length > 0;
      const hasCategories = analysis.textColumns.length > 0;
      const numericCount = analysis.numericColumns.length;
      const rowCount = analysis.totalRows;

      if (hasTime && numericCount > 0) {
        if (rowCount > 50) {
          return Charts.ChartType.AREA;
        } else if (numericCount > 2) {
          return Charts.ChartType.LINE;
        } else {
          return Charts.ChartType.COLUMN;
        }
      }

      if (hasCategories && numericCount === 1 && rowCount <= 8) {
        const numColumn = analysis.numericColumns[0];
        if (numColumn.stats.min >= 0) {
          return Charts.ChartType.PIE;
        }
      }

      if (hasCategories && numericCount > 0) {
        if (rowCount <= 5 && numericCount <= 3) {
          return Charts.ChartType.COLUMN;
        } else if (rowCount <= 15) {
          return Charts.ChartType.BAR;
        } else {
          return Charts.ChartType.TABLE;
        }
      }

      if (numericCount >= 2) {
        if (rowCount <= 100) {
          return Charts.ChartType.SCATTER;
        } else {
          return Charts.ChartType.AREA;
        }
      }

      if (numericCount === 1) {
        if (rowCount <= 20) {
          return Charts.ChartType.COLUMN;
        } else {
          return Charts.ChartType.HISTOGRAM;
        }
      }

      return Charts.ChartType.TABLE;
    },

    /**
     * Generate data summary
     */
    generateDataSummary: function (data, analysis) {
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
    generateDataInsights: function (data, analysis) {
      const insights = [];

      analysis.numericColumns.forEach((col) => {
        if (col.stats.min !== null && col.stats.max !== null) {
          insights.push({
            type: 'range',
            text: `${col.name} ranges from ${Math.round(col.stats.min)} to ${Math.round(col.stats.max)}`,
          });
        }
      });

      analysis.textColumns.forEach((col) => {
        if (col.uniqueValues > 1) {
          insights.push({
            type: 'categories',
            text: `${col.uniqueValues} unique ${col.name} values`,
          });
        }
      });

      const completeness = this.calculateCompleteness(data);
      if (completeness < 100) {
        insights.push({
          type: 'quality',
          text: `Data is ${completeness}% complete with some missing values`,
        });
      }

      return insights.slice(0, 5);
    },

    /**
     * Calculate data completeness percentage
     */
    calculateCompleteness: function (data) {
      const totalCells = data.data.length * data.headers.length;
      if (totalCells === 0) return 100;

      let filledCells = 0;
      data.data.forEach((row) => {
        row.forEach((cell) => {
          if (cell != null && cell !== '') {
            filledCells++;
          }
        });
      });

      return Math.round((filledCells / totalCells) * 100);
    },

    /**
     * Create metric cards for KPIs
     */
    createMetricCards: function (slide, metrics, theme, x, y, width, height) {
      if (!metrics || metrics.length === 0) return;

      const cardCount = Math.min(metrics.length, 4);
      const cardWidth = (width - (cardCount - 1) * 15) / cardCount;
      const cardHeight = height * 0.6;

      metrics.slice(0, cardCount).forEach((metric, index) => {
        const cardX = x + index * (cardWidth + 15);
        const cardY = y + 20;

        // Card shadow
        const shadow = slide.insertShape(
          SlidesApp.ShapeType.ROUND_RECTANGLE,
          cardX + 3,
          cardY + 3,
          cardWidth,
          cardHeight
        );
        shadow.getFill().setSolidFill('#e0e0e0');
        shadow.getBorder().setTransparent();
        shadow.sendToBack();

        // Main card
        const card = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, cardX, cardY, cardWidth, cardHeight);
        card.getFill().setSolidFill(theme.surface);
        card.getBorder().getLineFill().setSolidFill(theme.primary);
        card.getBorder().setWeight(1);

        // Metric icon/emoji based on type
        const icon =
          metric.type === 'monetary' ?
            '💰' :
            metric.type === 'performance' ?
              '📈' :
              metric.type === 'average' ?
                '📊' :
                '🎯';

        const iconBox = slide.insertTextBox(icon, cardX + cardWidth / 2 - 15, cardY + 15, 30, 30);
        iconBox.getText().getTextStyle().setFontSize(24);
        iconBox.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

        // Metric value
        let displayValue;
        if (metric.type === 'monetary') {
          displayValue = '$' + (metric.value > 1000 ? (metric.value / 1000).toFixed(0) + 'K' : metric.value);
        } else if (metric.value) {
          displayValue = metric.value.toString();
        } else {
          displayValue = '—';
        }

        const valueBox = slide.insertTextBox(displayValue, cardX + 10, cardY + 50, cardWidth - 20, 40);
        valueBox
          .getText()
          .getTextStyle()
          .setFontSize(28)
          .setFontFamily('Arial')
          .setBold(true)
          .setForegroundColor(theme.primary);
        valueBox.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

        // Metric label
        const labelBox = slide.insertTextBox(metric.label, cardX + 10, cardY + 95, cardWidth - 20, 30);
        labelBox.getText().getTextStyle().setFontSize(11).setFontFamily('Arial').setForegroundColor(theme.textLight);
        labelBox.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
      });
    },

    /**
     * Create time series chart
     */
    createTimeSeriesChart: function (slide, data, analysis, suggestion, theme, x, y, width, height) {
      Logger.log('Creating time series chart');

      // Find time and value columns
      let timeCol = 0;
      let valueCol = 1;

      if (suggestion.timeColumn) {
        const timeIndex = analysis.headers.indexOf(suggestion.timeColumn);
        if (timeIndex !== -1) timeCol = timeIndex;
      }

      if (suggestion.valueColumns && suggestion.valueColumns.length > 0) {
        const valueIndex = analysis.headers.indexOf(suggestion.valueColumns[0]);
        if (valueIndex !== -1) valueCol = valueIndex;
      }

      // For now, create a simple line visualization
      // In production, you'd create actual line charts with data points
      this.createBarChartVisualization(slide, data, analysis, theme, x, y, width, height);
    },

    /**
     * Create distribution chart (pie/donut)
     */
    createDistributionChart: function (slide, data, analysis, suggestion, theme, x, y, width, height) {
      Logger.log('Creating distribution chart');

      // For now, create percentage visualization
      // In production, you'd create actual pie charts
      this.createPercentageVisualization(slide, data, analysis, theme, x, y, width, height);
    },

    /**
     * Plan slide structure based on analysis
     */
    planSlideStructure: function (analysis, config) {
      const plan = [];
      const requestedSlides = parseInt(config.slideCount) || 5;

      Logger.log('Planning ' + requestedSlides + ' slides');

      // Always start with title slide
      plan.push({
        type: 'title',
        title: config.title || 'Data Presentation',
        subtitle: config.subtitle || analysis.dataSummary,
      });

      // Calculate remaining slides after title and conclusion
      const middleSlides = requestedSlides - 2; // -1 for title, -1 for conclusion

      // Build middle slides based on available data and requested count
      if (middleSlides > 0) {
        // Priority 1: Executive summary (if metrics available)
        if (analysis.keyMetrics && analysis.keyMetrics.length > 0) {
          plan.push({
            type: 'summary',
            metrics: analysis.keyMetrics.slice(0, 4),
            keyInsight: analysis.insights ? analysis.insights[0] : null,
          });
        }

        // Priority 2: Primary chart using first suggested visualization
        if (
          plan.length < requestedSlides - 1 &&
          analysis.suggestedVisualizations &&
          analysis.suggestedVisualizations.length > 0
        ) {
          const firstViz = analysis.suggestedVisualizations[0];
          plan.push({
            type: 'chart',
            chartType: firstViz.type || 'column',
            title: firstViz.title || 'Data Analysis',
            createNew: true,
            vizIndex: 0, // Use first visualization
            columns: firstViz.groupBy ? [firstViz.groupBy, firstViz.measure] : null,
          });
        }

        // Priority 3: Secondary chart using different visualization
        if (
          plan.length < requestedSlides - 1 &&
          analysis.suggestedVisualizations &&
          analysis.suggestedVisualizations.length > 1
        ) {
          const secondViz = analysis.suggestedVisualizations[1];
          plan.push({
            type: 'chart',
            chartType: secondViz.type || 'line',
            title: secondViz.title || 'Trend Analysis',
            createNew: true,
            vizIndex: 1, // Use second visualization
            columns: secondViz.groupBy ? [secondViz.groupBy, secondViz.measure] : null,
          });
        } else if (plan.length < requestedSlides - 1 && analysis.numericColumns && analysis.numericColumns.length > 1) {
          // Create a line chart with different numeric column
          plan.push({
            type: 'chart',
            chartType: 'line',
            title: 'Trend: ' + analysis.numericColumns[1].name,
            createNew: true,
            vizIndex: 1,
            useAlternateColumns: true, // Flag to use different columns
          });
        } else if (
          plan.length < requestedSlides - 1 &&
          analysis.categoryColumns &&
          analysis.categoryColumns.length > 0
        ) {
          // Create a pie chart for distribution
          plan.push({
            type: 'chart',
            chartType: 'pie',
            title: 'Distribution: ' + analysis.categoryColumns[0].name,
            createNew: true,
            vizIndex: 2,
          });
        }

        // Priority 4: Data table
        if (plan.length < requestedSlides - 1) {
          plan.push({
            type: 'table',
            title: 'Data Overview',
            maxRows: Math.min(10, analysis.totalRows || 10),
            maxCols: Math.min(6, analysis.totalColumns || 6),
          });
        }

        // Priority 5: Insights
        if (plan.length < requestedSlides - 1 && analysis.insights && analysis.insights.length > 0) {
          plan.push({
            type: 'insights',
            title: 'Key Insights',
            insights: analysis.insights.slice(0, 5),
          });
        }

        // Fill remaining slots with additional content
        while (plan.length < requestedSlides - 1) {
          const chartCount = plan.filter((s) => s.type === 'chart').length;
          if (chartCount < 3) {
            // Add another chart with proper vizIndex
            const nextChartSpec = {
              type: 'chart',
              createNew: true,
              vizIndex: chartCount, // Use the count as vizIndex
            };

            // Vary chart type based on what charts we already have
            if (chartCount === 1) {
              nextChartSpec.chartType = 'line';
              nextChartSpec.title = 'Trend Analysis';
            } else if (chartCount === 2) {
              nextChartSpec.chartType = 'pie';
              nextChartSpec.title = 'Distribution Analysis';
            } else {
              nextChartSpec.chartType = 'scatter';
              nextChartSpec.title = 'Correlation Analysis';
            }

            plan.push(nextChartSpec);
          } else {
            // Add a generic content slide
            plan.push({
              type: 'insights',
              title: 'Additional Analysis',
              insights: [
                'Data contains ' + (analysis.totalRows || 0) + ' records',
                'Analyzed across ' + (analysis.totalColumns || 0) + ' dimensions',
              ],
            });
          }
        }
      }

      // Always end with conclusion
      plan.push({
        type: 'conclusion',
        title: 'Conclusion',
        nextSteps: this.generateNextSteps(analysis),
      });

      // Ensure we have exactly the requested number of slides
      const finalPlan = plan.slice(0, requestedSlides);
      Logger.log('Final plan has ' + finalPlan.length + ' slides');

      return finalPlan;
    },

    /**
     * Generate next steps based on analysis
     */
    generateNextSteps: function (analysis) {
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
    buildSlidesFromPlan: function (presentation, plan, data, analysis, config) {
      // Map old template names to new professional themes
      const themeMap = {
        simple: 'executive',
        corporate: 'executive',
        modern: 'tech',
        elegant: 'minimal',
        vibrant: 'nature',
        dark: 'tech',
        professional: 'executive',
      };

      const themeName = themeMap[config.template] || 'executive';

      // Apply the professional theme to the presentation
      this.applyProfessionalTheme(presentation, themeName);

      // Get the theme object for slide building
      const theme = this.PROFESSIONAL_THEMES[themeName] || this.PROFESSIONAL_THEMES.executive;

      Logger.log('Building slides with professional theme: ' + theme.name);

      plan.forEach((slideSpec, index) => {
        Logger.log('Building slide ' + (index + 1) + ': ' + slideSpec.type);

        // Get or create slide
        const slides = presentation.getSlides();
        let slide;

        if (index < slides.length) {
          slide = slides[index];
          this.ensureSlideIsClean(slide);
        } else {
          slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
        }

        // Build slide based on type
        switch (slideSpec.type) {
        case 'title':
          this.buildTitleSlide(slide, slideSpec, theme);
          this.addSpeakerNotes(
            slide,
            'Welcome to this data presentation. This analysis covers ' +
                analysis.totalRows +
                ' records across ' +
                analysis.totalColumns +
                ' fields.'
          );
          break;
        case 'summary':
          this.buildSummarySlide(slide, slideSpec, theme);
          this.addSpeakerNotes(slide, 'Executive summary showing key metrics.');
          break;
        case 'chart':
          this.buildChartSlide(slide, slideSpec, data, analysis, theme);
          this.addSpeakerNotes(slide, 'This visualization shows the relationship between key data points.');
          break;
        case 'table':
          this.buildTableSlide(slide, slideSpec, data, theme);
          this.addSpeakerNotes(slide, 'Data table showing a sample of the dataset.');
          break;
        case 'insights':
          this.buildInsightsSlide(slide, slideSpec, theme);
          this.addSpeakerNotes(slide, 'Key insights derived from data analysis.');
          break;
        case 'conclusion':
          this.buildConclusionSlide(slide, slideSpec, theme);
          this.addSpeakerNotes(slide, 'Recommended next steps based on the analysis.');
          break;
        default:
          this.buildDefaultSlide(slide, slideSpec, theme);
        }
      });

      // Remove any extra slides
      let slides = presentation.getSlides();
      while (slides.length > plan.length) {
        slides[slides.length - 1].remove();
        slides = presentation.getSlides(); // Re-fetch after removal
      }
    },

    /**
     * Build title slide
     */
    buildTitleSlide: function (slide, spec, theme) {
      try {
        Logger.log('Building title slide with theme');

        // Add gradient background
        if (theme.gradients && theme.gradients.primary) {
          this.addGradientBackground(slide, theme.gradients.primary[0], theme.gradients.primary[1]);
        } else {
          this.addBackground(slide, theme.primary || '#1a73e8');
        }

        // Add decorative overlay
        const overlay = slide.insertShape(
          SlidesApp.ShapeType.RECTANGLE,
          0,
          this.SLIDE_HEIGHT * 0.7,
          this.SLIDE_WIDTH,
          this.SLIDE_HEIGHT * 0.3
        );
        overlay.getFill().setSolidFill(theme.primaryVariant || theme.primary || '#1557b0');
        // Note: Transparency on fill not supported in Slides API
        overlay.getBorder().setTransparent();
        overlay.sendToBack();
      } catch (error) {
        Logger.error('Error in buildTitleSlide background:', error);
      }

      // Add title with golden ratio positioning
      const titleY = this.SLIDE_HEIGHT / this.GOLDEN_RATIO - 30;
      const titleBox = slide.insertTextBox(
        spec.title || 'Presentation',
        this.SAFE_MARGINS.left,
        titleY,
        this.SLIDE_WIDTH - this.SAFE_MARGINS.left - this.SAFE_MARGINS.right,
        60
      );

      titleBox
        .getText()
        .getTextStyle()
        .setFontSize(this.TYPOGRAPHY.display)
        .setFontFamily('Arial')
        .setBold(true)
        .setForegroundColor(theme.darkMode ? '#FFFFFF' : theme.text || '#2C3E50');

      titleBox.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

      // Add subtitle
      if (spec.subtitle) {
        const subtitleBox = slide.insertTextBox(
          spec.subtitle || '',
          this.SAFE_MARGINS.left,
          titleY + 70,
          this.SLIDE_WIDTH - this.SAFE_MARGINS.left - this.SAFE_MARGINS.right,
          40
        );

        subtitleBox
          .getText()
          .getTextStyle()
          .setFontSize(this.TYPOGRAPHY.subtitle)
          .setFontFamily('Arial')
          .setForegroundColor(theme.darkMode ? '#FFFFFF' : theme.text || '#2C3E50');

        subtitleBox.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
      }

      // Add divider line
      console.log('CellM8 SlideGenerator VERSION:', this.VERSION);
      const decorLine = slide.insertShape(
        SlidesApp.ShapeType.RECTANGLE,
        (this.SLIDE_WIDTH - 120) / 2,
        titleY + 120,
        120,
        3
      );
      decorLine.getFill().setSolidFill(theme.darkMode ? '#FFFFFF' : theme.accent || '#4285F4');
      // Note: transparency on fill removed - not supported
      decorLine.getBorder().setTransparent();

      // Add date
      const date = new Date().toLocaleDateString();
      const dateBox = slide.insertTextBox(
        date,
        this.SLIDE_WIDTH - this.SAFE_MARGINS.right - 100,
        this.SLIDE_HEIGHT - this.SAFE_MARGINS.bottom,
        100,
        20
      );
      dateBox
        .getText()
        .getTextStyle()
        .setFontSize(this.TYPOGRAPHY.caption)
        .setFontFamily('Arial')
        .setForegroundColor(theme.darkMode ? '#FFFFFF' : theme.textLight || '#666666');
      // Note: text transparency would be: dateBox.getText().getTextStyle().setTransparency(0.7);
      dateBox.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.END);
    },

    /**
     * Build summary slide with metrics
     */
    buildSummarySlide: function (slide, spec, theme) {
      // Use the new dashboard layout system for executive summary
      const layoutConfig = {
        type: 'executive',
        theme: theme,
        data: {
          title: 'Executive Summary',
          metrics: spec.metrics || [],
          keyInsight: spec.keyInsight || 'Data analysis overview',
          highlights: spec.highlights || [],
        },
      };

      this.createDashboardLayout(slide, layoutConfig);
    },

    /**
     * Build chart slide with native linked charts
     */
    buildChartSlide: function (slide, spec, data, analysis, theme) {
      // Apply gradient background if available
      if (theme.gradients && theme.gradients.primary) {
        this.addGradientBackground(slide, theme.gradients.primary[0], theme.gradients.primary[1]);
      } else {
        this.addBackground(slide, theme.background);
      }
      this.addSlideTitle(slide, spec.title || 'Data Visualization', theme);

      try {
        const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = spreadsheet.getActiveSheet();

        // Ensure vizIndex is set
        if (spec.vizIndex === undefined) {
          // Count existing charts in presentation to determine vizIndex
          const presentation = SlidesApp.getActivePresentation();
          const existingSlides = presentation.getSlides();
          let chartCount = 0;
          for (let i = 0; i < existingSlides.length; i++) {
            if (existingSlides[i].getImages().length > 0 || existingSlides[i].getSheetsCharts().length > 0) {
              chartCount++;
            }
          }
          spec.vizIndex = chartCount;
        }

        Logger.log('=== BUILDING CHART SLIDE ===');
        Logger.log('Spec: ' + JSON.stringify(spec));
        Logger.log('Chart #' + spec.vizIndex + ': ' + spec.title + ', type: ' + spec.chartType);

        // Remove ALL existing charts from the sheet first to ensure clean slate
        const existingCharts = sheet.getCharts();
        Logger.log('Found ' + existingCharts.length + ' existing charts in sheet, removing them');
        console.log('CHART CLEANUP: Found ' + existingCharts.length + ' charts to remove');
        existingCharts.forEach((c, idx) => {
          try {
            const chartId = c.getId();
            console.log('  Removing chart ' + idx + ' (ID: ' + chartId + ')');
            sheet.removeChart(c);
          } catch (e) {
            Logger.log('Could not remove chart: ' + e);
            console.log('  ERROR removing chart ' + idx + ': ' + e);
          }
        });

        // Force flush to ensure charts are removed
        SpreadsheetApp.flush();
        console.log('CHART CLEANUP: Complete, flushed');

        spec.themeName = theme.name || 'executive';

        // Create unique chart based on vizIndex
        console.log('CREATING CHART: vizIndex=' + spec.vizIndex + ', type=' + spec.chartType);
        const chart = this.createSheetsChart(sheet, data, analysis, spec);

        if (chart) {
          Logger.log('Inserting linked Sheets chart with unique data');
          const chartId = chart.getId();
          console.log('CHART CREATED: ID=' + chartId + ', inserting into slide');

          // Calculate optimal position for chart
          const chartWidth = this.SLIDE_WIDTH - this.SAFE_MARGINS.left - this.SAFE_MARGINS.right;
          const chartHeight = 300;
          const chartY = 90;

          // Insert as linked chart (will auto-update when data changes)
          const linkedChart = slide.insertSheetsChart(chart, this.SAFE_MARGINS.left, chartY, chartWidth, chartHeight);

          // DON'T remove the chart - Google Slides needs it to exist in the sheet
          // The chart will be cleaned up when we create the next one
          Logger.log('Chart inserted successfully, keeping in sheet for link to work');
          console.log('CHART INSERTED: Successfully linked to slide');

          // Add visual indicator that chart is linked
          const linkIndicator = slide.insertTextBox(
            '🔗 Live data - Auto-updates with spreadsheet',
            this.SAFE_MARGINS.left,
            chartY + chartHeight + 10,
            200,
            20
          );
          linkIndicator.getText().getTextStyle().setFontSize(9).setForegroundColor(theme.textLight).setItalic(true);

          // Add data insights
          if (analysis.insights && analysis.insights.length > 0) {
            const insightBox = slide.insertTextBox(
              '💡 ' + (analysis.insights[0] || 'Data analysis complete'),
              this.SAFE_MARGINS.left,
              chartY + chartHeight + 35,
              chartWidth,
              30
            );
            insightBox.getText().getTextStyle().setFontSize(11).setForegroundColor(theme.accent);
          }
        } else {
          // Fallback to shape-based visualization
          Logger.log('Falling back to shape-based visualization');
          this.createVisualDataChart(slide, data, analysis, theme);
        }
      } catch (error) {
        Logger.error('Error in buildChartSlide:', error);
        // Fallback to shape-based visualization
        this.createVisualDataChart(slide, data, analysis, theme);
      }
    },

    /**
     * Create a chart in Google Sheets based on data analysis
     */
    createSheetsChart: function (sheet, dataStructure, analysis, spec) {
      try {
        // Use the correct visualization based on spec.vizIndex or create unique visualizations
        const vizIndex = spec.vizIndex !== undefined ? spec.vizIndex : 0;
        Logger.log('createSheetsChart - vizIndex: ' + vizIndex + ', chartType: ' + spec.chartType);

        // Create different data ranges and visualizations for each chart
        let visualization = null;
        let dataColumnIndices = [];

        console.log('CHART DATA SELECTION: Starting for vizIndex=' + vizIndex);

        // Create different visualizations based on vizIndex with different data columns
        if (vizIndex === 0) {
          // First chart: Primary analysis with first numeric column
          if (analysis.suggestedVisualizations && analysis.suggestedVisualizations.length > 0) {
            visualization = analysis.suggestedVisualizations[0];
          } else {
            visualization = {
              type: spec.chartType || 'column_chart',
              title: spec.title || 'Primary Analysis',
            };
          }
          // Use first category and first numeric column
          if (analysis.categoryColumns && analysis.categoryColumns.length > 0) {
            dataColumnIndices = [analysis.categoryColumns[0].index];
          }
          if (analysis.numericColumns && analysis.numericColumns.length > 0) {
            dataColumnIndices.push(analysis.numericColumns[0].index);
          }
        } else if (vizIndex === 1) {
          // Second chart: Different numeric column or different category grouping
          visualization = {
            type: 'line_chart',
            title: spec.title || 'Trend Analysis',
          };
          // Use different columns for second chart
          if (analysis.categoryColumns && analysis.categoryColumns.length > 0) {
            const catIndex = Math.min(1, analysis.categoryColumns.length - 1);
            dataColumnIndices = [analysis.categoryColumns[catIndex].index];
          }
          if (analysis.numericColumns && analysis.numericColumns.length > 1) {
            dataColumnIndices.push(analysis.numericColumns[1].index);
            visualization.measure = analysis.numericColumns[1].name;
          } else if (analysis.numericColumns && analysis.numericColumns.length > 0) {
            dataColumnIndices.push(analysis.numericColumns[0].index);
          }
        } else if (vizIndex === 2) {
          // Third chart: Pie chart with aggregated data
          visualization = {
            type: 'pie_chart',
            title: spec.title || 'Distribution',
          };
          // Use category and sum/count for pie chart
          if (analysis.categoryColumns && analysis.categoryColumns.length > 0) {
            dataColumnIndices = [analysis.categoryColumns[0].index];
          }
          if (analysis.numericColumns && analysis.numericColumns.length > 0) {
            const numIndex = Math.min(2, analysis.numericColumns.length - 1);
            dataColumnIndices.push(analysis.numericColumns[numIndex].index);
          }
        } else {
          // Additional charts: Use remaining columns
          visualization = {
            type: vizIndex % 2 === 0 ? 'bar_chart' : 'scatter_chart',
            title: spec.title || 'Analysis ' + vizIndex,
          };
          // Try to use different column combinations
          const catIndex = vizIndex % Math.max(1, analysis.categoryColumns.length);
          const numIndex = vizIndex % Math.max(1, analysis.numericColumns.length);
          if (analysis.categoryColumns && analysis.categoryColumns.length > catIndex) {
            dataColumnIndices = [analysis.categoryColumns[catIndex].index];
          }
          if (analysis.numericColumns && analysis.numericColumns.length > numIndex) {
            dataColumnIndices.push(analysis.numericColumns[numIndex].index);
          }
        }

        // Determine chart type - ensure variety based on vizIndex
        let chartTypeStr = spec.chartType;
        if (!chartTypeStr) {
          if (vizIndex === 0) chartTypeStr = 'column_chart';
          else if (vizIndex === 1) chartTypeStr = 'line_chart';
          else if (vizIndex === 2) chartTypeStr = 'pie_chart';
          else chartTypeStr = 'bar_chart';
        }
        if (visualization && visualization.type) {
          chartTypeStr = visualization.type;
        }

        const chartType = this.mapToSheetsChartType(chartTypeStr);

        // Log selected columns
        console.log('SELECTED COLUMNS: indices=' + JSON.stringify(dataColumnIndices));
        if (dataColumnIndices.length > 0 && analysis.headers) {
          const selectedHeaders = dataColumnIndices.map((idx) => analysis.headers[idx] || 'Col' + idx);
          console.log('SELECTED HEADERS: ' + selectedHeaders.join(', '));
        }

        // Select appropriate data range based on the visualization and column indices
        let dataRange;
        if (dataColumnIndices.length > 0) {
          // Create range with specific columns to ensure different data for each chart
          const lastRow = sheet.getLastRow();
          const rangeNotation = this.createRangeNotation(dataColumnIndices, lastRow);
          Logger.log('Using specific columns for chart ' + vizIndex + ': ' + rangeNotation);
          console.log('RANGE NOTATION: ' + rangeNotation + ' for chart ' + vizIndex);
          dataRange = sheet.getRange(rangeNotation);
        } else {
          // Fallback to default range selection
          dataRange = this.selectChartDataRange(
            sheet,
            dataStructure,
            {
              ...analysis,
              suggestedVisualizations: visualization ? [visualization] : analysis.suggestedVisualizations,
            },
            spec
          );
        }

        if (!dataRange) {
          Logger.log('Could not determine data range for chart');
          return null;
        }

        // Build the chart with unique title based on data
        let chartTitle = spec.title || 'Data Analysis';
        if (vizIndex === 1 && !spec.title) {
          chartTitle = 'Trend Analysis';
        } else if (vizIndex === 2 && !spec.title) {
          chartTitle = 'Distribution Overview';
        }

        Logger.log('Building chart: ' + chartTitle + ' with type: ' + chartTypeStr);

        // Use significantly different positions for different charts to ensure unique IDs
        const chartPosition = {
          row: 1 + vizIndex * 5, // Larger row spacing
          column: 10 + vizIndex * 2, // Different column positions
          offsetX: vizIndex * 50, // Larger offset differences
          offsetY: vizIndex * 50,
        };

        // Add unique ID to ensure charts are different
        const uniqueId = 'chart_' + vizIndex + '_' + Date.now();

        const chartBuilder = sheet
          .newChart()
          .setChartType(chartType)
          .addRange(dataRange)
          .setPosition(chartPosition.row, chartPosition.column, chartPosition.offsetX, chartPosition.offsetY)
          .setOption('title', chartTitle + ' (#' + (vizIndex + 1) + ')')
          .setOption('width', 600)
          .setOption('height', 400);

        // Apply professional theme colors
        const currentTheme = this.PROFESSIONAL_THEMES[spec.themeName] || this.PROFESSIONAL_THEMES.executive;
        chartBuilder.setOption('backgroundColor', currentTheme.background).setOption('titleTextStyle', {
          color: currentTheme.text,
          fontSize: 16,
          bold: true,
        });

        // Apply advanced chart configuration for professional look
        const advancedConfig = {
          style: 'gradient',
          animation: true,
          legend: { position: 'right' },
          colors: currentTheme.chartColors || [currentTheme.primary, currentTheme.secondary],
        };
        this.configureAdvancedChart(chartBuilder, advancedConfig);

        // Set chart-specific options based on type
        this.configureChartOptions(chartBuilder, chartType, analysis);

        const chart = chartBuilder.build();
        sheet.insertChart(chart);

        Logger.log('Successfully created Sheets chart');
        return chart;
      } catch (error) {
        Logger.error('Error creating Sheets chart:', error);
        return null;
      }
    },

    /**
     * Map analysis chart type to Google Sheets chart type
     */
    mapToSheetsChartType: function (analysisType) {
      const mapping = {
        bar_chart: Charts.ChartType.BAR,
        grouped_bar: Charts.ChartType.BAR,
        line_chart: Charts.ChartType.LINE,
        pie_chart: Charts.ChartType.PIE,
        scatter_plot: Charts.ChartType.SCATTER,
        area_chart: Charts.ChartType.AREA,
        column_chart: Charts.ChartType.COLUMN,
        metric_cards: Charts.ChartType.TABLE,
        distribution: Charts.ChartType.PIE,
      };
      return mapping[analysisType] || Charts.ChartType.COLUMN;
    },

    /**
     * Create range notation from column indices
     */
    createRangeNotation: function (columnIndices, lastRow) {
      if (!columnIndices || columnIndices.length === 0) return null;

      // Convert indices to column letters
      const columns = columnIndices.map((idx) => {
        let column = '';
        let num = idx + 1; // Convert to 1-based index
        while (num > 0) {
          num--;
          column = String.fromCharCode((num % 26) + 65) + column;
          num = Math.floor(num / 26);
        }
        return column;
      });

      // Create range notation like "A1:A100,C1:C100"
      const ranges = columns.map((col) => col + '1:' + col + lastRow);

      // If consecutive columns, use single range notation
      if (columns.length === 2 && columnIndices[1] === columnIndices[0] + 1) {
        return columns[0] + '1:' + columns[1] + lastRow;
      }

      // For non-consecutive columns, return the first range (Sheets can only handle one range)
      // We'll need to reorganize the data for multi-column ranges
      return columns[0] + '1:' + columns[columns.length - 1] + lastRow;
    },

    /**
     * Select appropriate data range for chart
     */
    selectChartDataRange: function (sheet, dataStructure, analysis, spec) {
      try {
        // Get the data range that includes headers
        const numRows = Math.min(dataStructure.data.length + 1, 100); // Limit to 100 rows
        const numCols = dataStructure.headers.length;

        // Find the best columns for the chart based on visualization
        const vizIndex = spec && spec.vizIndex !== undefined ? spec.vizIndex : 0;

        Logger.log('=== SELECTING DATA RANGE ===');
        Logger.log('vizIndex: ' + vizIndex + ', chartType: ' + spec.chartType);
        Logger.log('Available columns: ' + dataStructure.headers.join(', '));
        Logger.log(
          'Numeric columns: ' +
            (analysis.numericColumns ? analysis.numericColumns.map((c) => c.name).join(', ') : 'none')
        );
        Logger.log(
          'Category columns: ' +
            (analysis.categoryColumns ? analysis.categoryColumns.map((c) => c.name).join(', ') : 'none')
        );

        // For different chart types and indices, use different column combinations
        if (spec.chartType === 'pie' || vizIndex === 2) {
          // Pie chart: Use first category column and first numeric column
          if (
            analysis.categoryColumns &&
            analysis.categoryColumns.length > 0 &&
            analysis.numericColumns &&
            analysis.numericColumns.length > 0
          ) {
            const catIndex = dataStructure.headers.indexOf(analysis.categoryColumns[0].name) + 1;
            const numIndex = dataStructure.headers.indexOf(analysis.numericColumns[0].name) + 1;
            if (catIndex > 0 && numIndex > 0) {
              return sheet.getRange(1, Math.min(catIndex, numIndex), numRows, 2);
            }
          }
        } else if (spec.chartType === 'line' || vizIndex === 1) {
          // Line chart: Use different numeric columns if available
          if (analysis.numericColumns && analysis.numericColumns.length > 1) {
            const col1 = analysis.numericColumns[0];
            const col2 = analysis.numericColumns[1];
            const idx1 = dataStructure.headers.indexOf(col1.name) + 1;
            const idx2 = dataStructure.headers.indexOf(col2.name) + 1;
            if (idx1 > 0 && idx2 > 0) {
              return sheet.getRange(1, Math.min(idx1, idx2), numRows, Math.abs(idx2 - idx1) + 1);
            }
          }
        } else if (spec.chartType === 'scatter' || vizIndex === 3) {
          // Scatter: Use two numeric columns
          if (analysis.numericColumns && analysis.numericColumns.length >= 2) {
            const xCol = analysis.numericColumns[vizIndex % analysis.numericColumns.length];
            const yCol = analysis.numericColumns[(vizIndex + 1) % analysis.numericColumns.length];
            const xIdx = dataStructure.headers.indexOf(xCol.name) + 1;
            const yIdx = dataStructure.headers.indexOf(yCol.name) + 1;
            if (xIdx > 0 && yIdx > 0) {
              const range1 = sheet.getRange(1, xIdx, numRows, 1);
              const range2 = sheet.getRange(1, yIdx, numRows, 1);
              return [range1, range2];
            }
          }
        }

        // Try to use suggested visualizations if available
        if (analysis.suggestedVisualizations && analysis.suggestedVisualizations[vizIndex]) {
          const suggestion = analysis.suggestedVisualizations[vizIndex];

          // For aggregation charts, focus on category and measure columns
          if (suggestion.groupBy && suggestion.measure) {
            const groupByIndex = dataStructure.headers.indexOf(suggestion.groupBy) + 1;
            const measureIndex = dataStructure.headers.indexOf(suggestion.measure) + 1;

            if (groupByIndex > 0 && measureIndex > 0) {
              const range1 = sheet.getRange(1, groupByIndex, numRows, 1);
              const range2 = sheet.getRange(1, measureIndex, numRows, 1);
              return [range1, range2];
            }
          }
        }

        // For different chart indices, try to select different column combinations
        if (vizIndex > 0 && analysis.numericColumns && analysis.numericColumns.length > vizIndex) {
          // Use different numeric columns for different charts
          const numericCol = analysis.numericColumns[Math.min(vizIndex, analysis.numericColumns.length - 1)];
          const numericIndex = dataStructure.headers.indexOf(numericCol.name) + 1;

          if (analysis.categoryColumns && analysis.categoryColumns.length > 0) {
            const categoryIndex = dataStructure.headers.indexOf(analysis.categoryColumns[0].name) + 1;
            if (categoryIndex > 0 && numericIndex > 0) {
              const range1 = sheet.getRange(1, categoryIndex, numRows, 1);
              const range2 = sheet.getRange(1, numericIndex, numRows, 1);
              return [range1, range2];
            }
          }
        }

        // Default: use all data
        return sheet.getRange(1, 1, numRows, Math.min(numCols, 10));
      } catch (error) {
        Logger.error('Error selecting chart data range:', error);
        return null;
      }
    },

    /**
     * Configure chart-specific options
     */
    configureChartOptions: function (chartBuilder, chartType, analysis) {
      // Common options for all charts with focus on horizontal labels
      chartBuilder
        .setOption('legend', { position: 'bottom' })
        .setOption('fontSize', 12)
        .setOption('vAxis', {
          title: '',
          textStyle: {
            fontSize: 11,
          },
          format: 'short',
          textPosition: 'out',
          slantedText: false,
          slantedTextAngle: 0,
          viewWindow: { min: 0 },
        })
        .setOption('hAxis', {
          title: '',
          textStyle: {
            fontSize: 11,
          },
          format: 'short',
          slantedText: false,
          slantedTextAngle: 0,
          textPosition: 'out',
          maxAlternation: 1,
          showTextEvery: 1,
          maxTextLines: 1,
        })
        .setOption('chartArea', {
          left: 80, // More space for vertical axis labels
          right: 50,
          top: 30,
          bottom: 80,
        });

      // Type-specific options
      switch (chartType) {
      case Charts.ChartType.PIE:
        chartBuilder
          .setOption('is3D', true)
          .setOption('pieSliceText', 'percentage')
          .setOption('sliceVisibilityThreshold', 0.02);
        break;

      case Charts.ChartType.BAR:
      case Charts.ChartType.COLUMN:
        chartBuilder
          .setOption('animation', { startup: true, duration: 1000 })
          .setOption('bar', { groupWidth: '75%' });

        // Add trendline for numeric data
        if (analysis.trends && analysis.trends.length > 0) {
          chartBuilder.setOption('trendlines', { 0: {} });
        }
        break;

      case Charts.ChartType.LINE:
      case Charts.ChartType.AREA:
        chartBuilder.setOption('curveType', 'function').setOption('lineWidth', 2).setOption('pointSize', 5);
        break;

      case Charts.ChartType.SCATTER:
        chartBuilder.setOption('pointSize', 7).setOption('trendlines', { 0: { type: 'linear' } });
        break;
      }

      return chartBuilder;
    },

    // Chart Style Templates for professional visualizations
    CHART_STYLE_TEMPLATES: {
      minimal: {
        gridlines: { color: '#F5F5F5', count: 4 },
        backgroundColor: 'transparent',
        chartArea: { left: 50, top: 20, width: '85%', height: '75%' },
        legend: { position: 'none' },
        titlePosition: 'none',
        axisTitlesPosition: 'none',
        hAxis: { textStyle: { fontSize: 10, color: '#999999' } },
        vAxis: { textStyle: { fontSize: 10, color: '#999999' } },
      },
      detailed: {
        gridlines: { color: '#E0E0E0', count: 8 },
        backgroundColor: '#FAFAFA',
        chartArea: { left: 80, top: 50, width: '70%', height: '65%' },
        legend: { position: 'right', textStyle: { fontSize: 11 } },
        titleTextStyle: { fontSize: 16, bold: true },
        hAxis: {
          title: 'Categories',
          titleTextStyle: { italic: false, fontSize: 12 },
          textStyle: { fontSize: 11 },
        },
        vAxis: {
          title: 'Values',
          titleTextStyle: { italic: false, fontSize: 12 },
          textStyle: { fontSize: 11 },
        },
      },
      dashboard: {
        gridlines: { color: '#F0F0F0', count: 5 },
        backgroundColor: 'white',
        chartArea: { left: 40, top: 20, width: '90%', height: '80%' },
        legend: { position: 'top', maxLines: 1, textStyle: { fontSize: 10 } },
        titleTextStyle: { fontSize: 14, bold: true },
        hAxis: { textStyle: { fontSize: 9 } },
        vAxis: { textStyle: { fontSize: 9 } },
      },
      executive: {
        gridlines: { color: '#E8E8E8', count: 6 },
        backgroundColor: '#FFFFFF',
        chartArea: { left: 70, top: 40, width: '75%', height: '70%' },
        legend: { position: 'bottom', textStyle: { fontSize: 12, color: '#333333' } },
        titleTextStyle: { fontSize: 18, bold: true, color: '#1B2951' },
        hAxis: {
          titleTextStyle: { fontSize: 12, color: '#666666' },
          textStyle: { fontSize: 11, color: '#666666' },
        },
        vAxis: {
          titleTextStyle: { fontSize: 12, color: '#666666' },
          textStyle: { fontSize: 11, color: '#666666' },
        },
      },
    },

    // Animation presets for charts
    ANIMATION_PRESETS: {
      none: { startup: false },
      subtle: { startup: true, duration: 500, easing: 'linear' },
      smooth: { startup: true, duration: 1000, easing: 'inAndOut' },
      dramatic: { startup: true, duration: 2000, easing: 'out' },
      bounce: { startup: true, duration: 1500, easing: 'inAndOut' },
    },

    /**
     * Configure advanced chart with professional styling
     */
    configureAdvancedChart: function (chartBuilder, config) {
      // Apply style template
      const style = this.CHART_STYLE_TEMPLATES[config.style || 'detailed'];
      Object.keys(style).forEach((key) => {
        chartBuilder.setOption(key, style[key]);
      });

      // Apply theme colors if specified
      if (config.theme && this.PROFESSIONAL_THEMES[config.theme]) {
        const theme = this.PROFESSIONAL_THEMES[config.theme];
        chartBuilder.setOption('colors', theme.chartColors);

        if (theme.darkMode) {
          chartBuilder.setOption('backgroundColor', theme.background);
          chartBuilder.setOption('legendTextStyle', { color: theme.text });
          chartBuilder.setOption('titleTextStyle', { color: theme.text });
        }
      }

      // Apply animation
      const animation = this.ANIMATION_PRESETS[config.animation || 'smooth'];
      if (animation.startup) {
        chartBuilder.setOption('animation', animation);
      }

      // Add data labels if requested
      if (config.showDataLabels) {
        chartBuilder.setOption('dataLabels', 'value');

        // For pie charts, show percentage
        if (config.chartType === Charts.ChartType.PIE) {
          chartBuilder.setOption('pieSliceText', 'value-and-percentage');
        }
      }

      // Add trendline if applicable
      if (config.showTrendline) {
        chartBuilder.setOption('trendlines', {
          0: {
            type: config.trendlineType || 'linear',
            color: config.trendlineColor || '#999999',
            lineWidth: 2,
            opacity: 0.7,
            showR2: true,
            visibleInLegend: true,
            labelInLegend: 'Trend',
          },
        });
      }

      // Set specific formatting for financial data
      if (config.isFinancial) {
        chartBuilder.setOption('vAxis', {
          format: 'currency',
          gridlines: { count: 10 },
        });
      }

      // Set percentage format if needed
      if (config.isPercentage) {
        chartBuilder.setOption('vAxis', {
          format: 'percent',
          minValue: 0,
          maxValue: 1,
        });
      }

      return chartBuilder;
    },

    /**
     * Apply slide transitions to presentation
     */
    applySlideTransitions: function (presentation, transitionType) {
      try {
        const transitions = {
          fade: { transitionType: 'FADE', duration: 0.5 },
          slide: { transitionType: 'SLIDE_FROM_RIGHT', duration: 0.4 },
          flip: { transitionType: 'FLIP', duration: 0.6 },
          cube: { transitionType: 'CUBE', duration: 0.7 },
          dissolve: { transitionType: 'DISSOLVE', duration: 0.5 },
        };

        const transition = transitions[transitionType] || transitions.fade;
        const slides = presentation.getSlides();

        // Note: Google Slides API has limited transition support via Apps Script
        // This is a placeholder for future enhancement
        Logger.log('Applied transition type: ' + transitionType);

        return true;
      } catch (error) {
        Logger.log('Error applying slide transitions: ' + error);
        return false;
      }
    },

    /**
     * Apply chart animations
     */
    applyChartAnimation: function (chartBuilder, animationType) {
      try {
        const animations = {
          subtle: { startup: true, duration: 500, easing: 'linear' },
          smooth: { startup: true, duration: 1000, easing: 'inAndOut' },
          dramatic: { startup: true, duration: 2000, easing: 'out' },
          bounce: { startup: true, duration: 1500, easing: 'inAndOut' },
        };

        const animation = animations[animationType] || animations.smooth;
        chartBuilder.setOption('animation', animation);
        return chartBuilder;
      } catch (error) {
        Logger.log('Error applying chart animation: ' + error);
        return chartBuilder;
      }
    },

    /**
     * Add annotations to chart
     */
    addChartAnnotations: function (chartBuilder, annotations) {
      try {
        if (annotations && annotations.length > 0) {
          chartBuilder.setOption('annotations', {
            textStyle: {
              fontSize: 12,
              color: '#666666',
              bold: true,
            },
            boxStyle: {
              stroke: '#888',
              strokeWidth: 1,
              gradient: {
                color1: '#fbf6a7',
                color2: '#33b679',
                x1: '0%',
                y1: '0%',
                x2: '100%',
                y2: '100%',
              },
            },
          });
        }
        return chartBuilder;
      } catch (error) {
        Logger.log('Error adding chart annotations: ' + error);
        return chartBuilder;
      }
    },

    /**
     * Apply conditional formatting to chart based on data values
     */
    applyConditionalChartFormatting: function (chartBuilder, data, rules) {
      const colors = [];
      const dataLabels = [];

      // Evaluate each data point against rules
      data.forEach((row, index) => {
        let color = '#3498DB'; // default
        let label = '';

        rules.forEach((rule) => {
          const value = typeof row === 'object' ? row[rule.column] : row;

          if (this.evaluateCondition(value, rule)) {
            color = rule.color;
            if (rule.label) {
              label = rule.label;
            }
          }
        });

        colors.push(color);
        if (label) {
          dataLabels.push({ index: index, text: label });
        }
      });

      // Apply colors
      if (colors.length > 0) {
        chartBuilder.setOption('colors', colors);
      }

      // Apply data labels
      if (dataLabels.length > 0) {
        chartBuilder.setOption('annotations', {
          textStyle: { fontSize: 10 },
          stem: { color: 'transparent' },
          style: 'point',
        });
      }

      return chartBuilder;
    },

    /**
     * Evaluate condition for conditional formatting
     */
    evaluateCondition: function (value, rule) {
      const numValue = parseFloat(value);

      switch (rule.operator) {
      case '>':
        return numValue > rule.value;
      case '<':
        return numValue < rule.value;
      case '>=':
        return numValue >= rule.value;
      case '<=':
        return numValue <= rule.value;
      case '==':
        return value == rule.value;
      case '!=':
        return value != rule.value;
      case 'between':
        return numValue >= rule.min && numValue <= rule.max;
      case 'contains':
        return value.toString().includes(rule.value);
      default:
        return false;
      }
    },

    /**
     * Build table slide with data overview metrics instead of raw table
     */
    buildTableSlide: function (slide, spec, data, theme) {
      // Apply gradient background
      if (theme.gradients && theme.gradients.primary) {
        this.addGradientBackground(slide, theme.gradients.primary[0], theme.gradients.primary[1]);
      } else {
        this.addGradientBackground(slide, theme.background, theme.surface);
      }
      this.addSlideTitle(slide, spec.title || 'Data Overview', theme);

      // Ensure data structure is valid
      if (!data || !data.headers || !data.data) {
        this.addErrorMessage(slide, 'No data available for overview', theme);
        return;
      }

      // Analyze data for overview metrics
      const analysis = this.analyzeDataIntelligently(data);
      const startY = 100;

      // Create metric cards layout
      const cardWidth = 155;
      const cardHeight = 85;
      const spacing = 15;
      const cardsPerRow = 4;

      // Create overview metrics
      const metrics = [];

      // Core data metrics
      metrics.push({
        value: data.data.length.toLocaleString(),
        label: 'Total Records',
        icon: '📊',
        color: theme.accent,
      });

      metrics.push({
        value: ((data.headers && data.headers.length) || 0).toString(),
        label: 'Data Columns',
        icon: '📈',
        color: theme.primary,
      });

      // Add numeric statistics
      if (analysis.numericColumns && analysis.numericColumns.length > 0) {
        const numericCol = analysis.numericColumns[0];
        if (numericCol.stats) {
          metrics.push({
            value: this.formatMetricValue(numericCol.stats.average),
            label: 'Avg ' + this.truncateLabel(numericCol.name, 12),
            icon: '📉',
            color: theme.secondary,
          });

          metrics.push({
            value: this.formatMetricValue(numericCol.stats.max),
            label: 'Max ' + this.truncateLabel(numericCol.name, 12),
            icon: '⬆️',
            color: '#27AE60',
          });
        }
      }

      // Add category insights
      if (analysis.categoryColumns && analysis.categoryColumns.length > 0) {
        const catCol = analysis.categoryColumns[0];
        metrics.push({
          value: (catCol.uniqueValues || catCol.uniqueCount || 0).toString(),
          label: this.truncateLabel(catCol.name || 'Category', 12) + ' Types',
          icon: '🏷️',
          color: theme.accent,
        });
      }

      // Data completeness
      const completeness = this.calculateDataCompleteness(data);
      metrics.push({
        value: completeness + '%',
        label: 'Completeness',
        icon: '✓',
        color: completeness > 90 ? '#27AE60' : '#F39C12',
      });

      // If we have date columns, add date range
      if (analysis.dateColumns && analysis.dateColumns.length > 0) {
        const dateRange = this.getDateRange(data, analysis.dateColumns[0]);
        if (dateRange) {
          metrics.push({
            value: dateRange,
            label: 'Date Range',
            icon: '📅',
            color: theme.primary,
          });
        }
      }

      // Calculate layout
      const totalWidth = cardWidth * cardsPerRow + spacing * (cardsPerRow - 1);
      const startX = (this.SLIDE_WIDTH - totalWidth) / 2;

      // Create metric cards
      metrics.slice(0, 8).forEach((metric, index) => {
        const row = Math.floor(index / cardsPerRow);
        const col = index % cardsPerRow;
        const x = startX + col * (cardWidth + spacing);
        const y = startY + row * (cardHeight + spacing);

        // Card background
        const card = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, x, y, cardWidth, cardHeight);
        card.getFill().setSolidFill(theme.surface);
        card
          .getBorder()
          .getLineFill()
          .setSolidFill(metric.color || theme.borderColor);
        card.getBorder().setWeight(1);

        // Icon
        if (metric.icon) {
          const iconBox = slide.insertTextBox(metric.icon, x + 10, y + 8, 30, 30);
          iconBox.getText().getTextStyle().setFontSize(20);
        }

        // Metric value
        const valueBox = slide.insertTextBox(metric.value || '0', x + 45, y + 12, cardWidth - 50, 30);
        valueBox
          .getText()
          .getTextStyle()
          .setFontSize(22)
          .setFontFamily('Arial')
          .setBold(true)
          .setForegroundColor(metric.color || theme.primary);
        valueBox.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.END);

        // Metric label
        const labelBox = slide.insertTextBox(metric.label || 'Data', x + 10, y + 45, cardWidth - 20, 30);
        labelBox.getText().getTextStyle().setFontSize(11).setFontFamily('Arial').setForegroundColor(theme.textLight);
        labelBox.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
      });

      // Add data structure summary
      const summaryY = startY + Math.ceil(metrics.length / cardsPerRow) * (cardHeight + spacing) + 30;
      const summaryWidth = 650;
      const summaryX = (this.SLIDE_WIDTH - summaryWidth) / 2;

      // Create summary text
      const summaryParts = [];
      const headerCount = (data.headers && data.headers.length) || 0;
      summaryParts.push(`Dataset: ${data.data.length} rows × ${headerCount} columns`);

      if (analysis.numericColumns && analysis.numericColumns.length > 0) {
        summaryParts.push(`${analysis.numericColumns.length} numeric fields`);
      }
      if (analysis.categoryColumns && analysis.categoryColumns.length > 0) {
        summaryParts.push(`${analysis.categoryColumns.length} categorical fields`);
      }
      if (analysis.dateColumns && analysis.dateColumns.length > 0) {
        summaryParts.push(`${analysis.dateColumns.length} date fields`);
      }

      const summaryText = summaryParts.join(' • ');

      const summaryBox = slide.insertTextBox(
        '📋 ' + (summaryText || 'Data Summary'),
        summaryX,
        summaryY,
        summaryWidth,
        30
      );
      summaryBox
        .getText()
        .getTextStyle()
        .setFontSize(12)
        .setFontFamily('Arial')
        .setForegroundColor(theme.textLight)
        .setItalic(true);
      summaryBox.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

      // Add column names preview
      const columnsPreview =
        data.headers && data.headers.length > 0 ?
          'Columns: ' +
            data.headers
              .slice(0, 5)
              .map((h) => this.truncateLabel(h || 'Column', 15))
              .join(', ') +
            (data.headers.length > 5 ? ` +${data.headers.length - 5} more` : '') :
          'No column headers available';

      const columnsBox = slide.insertTextBox(columnsPreview, summaryX, summaryY + 35, summaryWidth, 20);
      columnsBox.getText().getTextStyle().setFontSize(10).setFontFamily('Arial').setForegroundColor(theme.textLight);
      columnsBox.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    },

    /**
     * Calculate data completeness percentage
     */
    calculateDataCompleteness: function (data) {
      let totalCells = 0;
      let filledCells = 0;

      data.data.forEach((row) => {
        row.forEach((cell) => {
          totalCells++;
          if (cell !== null && cell !== '' && cell !== undefined) {
            filledCells++;
          }
        });
      });

      return totalCells > 0 ? Math.round((filledCells / totalCells) * 100) : 0;
    },

    /**
     * Get date range from data
     */
    getDateRange: function (data, dateColumn) {
      if (!dateColumn || !dateColumn.index) return null;

      const dates = data.data
        .map((row) => row[dateColumn.index])
        .filter((val) => val)
        .map((val) => new Date(val))
        .filter((d) => !isNaN(d));

      if (dates.length === 0) return null;

      const minDate = new Date(Math.min(...dates));
      const maxDate = new Date(Math.max(...dates));

      const formatDate = (d) => {
        const month = (d.getMonth() + 1).toString().padStart(2, '0');
        const day = d.getDate().toString().padStart(2, '0');
        return `${month}/${day}/${d.getFullYear()}`;
      };

      if (minDate.getTime() === maxDate.getTime()) {
        return formatDate(minDate);
      } else {
        return formatDate(minDate) + ' - ' + formatDate(maxDate);
      }
    },

    /**
     * Format metric value for display
     */
    formatMetricValue: function (value) {
      if (value === null || value === undefined) return '0';
      if (typeof value === 'number') {
        if (value >= 1000000) {
          return (value / 1000000).toFixed(1) + 'M';
        } else if (value >= 1000) {
          return (value / 1000).toFixed(1) + 'K';
        } else if (value < 1 && value > 0) {
          return value.toFixed(2);
        } else {
          return Math.round(value).toString();
        }
      }
      return String(value).substring(0, 10);
    },

    /**
     * Truncate label to fit
     */
    truncateLabel: function (label, maxLength) {
      if (label === null || label === undefined) return '';
      const str = String(label);
      if (str.length <= maxLength) return str;
      return str.substring(0, maxLength - 3) + '...';
    },

    /**
     * Build insights slide
     */
    buildInsightsSlide: function (slide, spec, theme) {
      // Apply gradient background
      if (theme.gradients && theme.gradients.primary) {
        this.addGradientBackground(slide, theme.gradients.primary[0], theme.gradients.primary[1]);
      } else {
        this.addBackground(slide, theme.background);
      }
      this.addSlideTitle(slide, spec.title || 'Key Insights', theme);

      const insights = spec.insights || [];
      const startY = 100;

      insights.forEach((insight, index) => {
        const y = startY + index * 50;

        // Bullet point
        const bullet = slide.insertShape(SlidesApp.ShapeType.ELLIPSE, this.SAFE_MARGINS.left, y + 5, 8, 8);
        bullet.getFill().setSolidFill(theme.accent);
        bullet.getBorder().setTransparent();

        // Insight text
        const textBox = slide.insertTextBox(
          insight.text || 'No insight text available',
          this.SAFE_MARGINS.left + 20,
          y,
          this.SLIDE_WIDTH - this.SAFE_MARGINS.left - this.SAFE_MARGINS.right - 20,
          30
        );
        textBox.getText().getTextStyle().setFontSize(14).setFontFamily('Arial').setForegroundColor(theme.text);
      });
    },

    /**
     * Build conclusion slide
     */
    buildConclusionSlide: function (slide, spec, theme) {
      // Apply gradient background with overlay
      if (theme.gradients && theme.gradients.overlay) {
        this.addGradientBackground(slide, theme.gradients.primary[0], theme.gradients.primary[1]);
        // Add subtle overlay
        const overlay = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 0, 0, this.SLIDE_WIDTH, this.SLIDE_HEIGHT);
        overlay.getFill().setSolidFill(theme.primary);
        overlay.getBorder().setTransparent();
        overlay.sendToBack();
        // Note: Can't set transparency in Slides API
      } else {
        this.addBackground(slide, theme.primary);
      }

      const titleBox = slide.insertTextBox(
        spec.title || 'Next Steps',
        this.SAFE_MARGINS.left,
        this.SAFE_MARGINS.top,
        this.SLIDE_WIDTH - this.SAFE_MARGINS.left - this.SAFE_MARGINS.right,
        50
      );
      titleBox
        .getText()
        .getTextStyle()
        .setFontSize(32)
        .setFontFamily('Arial')
        .setBold(true)
        .setForegroundColor(theme.darkMode ? '#FFFFFF' : theme.text || '#2C3E50');
      titleBox.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

      const steps = spec.nextSteps || [];
      const startY = 140;

      steps.forEach((step, index) => {
        const y = startY + index * 60;

        // Step number
        const numberBox = slide.insertTextBox((index + 1).toString(), this.SAFE_MARGINS.left + 100, y, 30, 30);
        numberBox
          .getText()
          .getTextStyle()
          .setFontSize(20)
          .setFontFamily('Arial')
          .setBold(true)
          .setForegroundColor(theme.darkMode ? '#FFFFFF' : theme.text || '#2C3E50');

        // Step text
        const stepBox = slide.insertTextBox(step || 'Next step', this.SAFE_MARGINS.left + 140, y + 5, 400, 30);
        stepBox
          .getText()
          .getTextStyle()
          .setFontSize(16)
          .setFontFamily('Arial')
          .setForegroundColor(theme.darkMode ? '#FFFFFF' : theme.textLight || '#666666');
      });
    },

    /**
     * Build default slide
     */
    buildDefaultSlide: function (slide, spec, theme) {
      // Apply gradient background
      if (theme.gradients && theme.gradients.primary) {
        this.addGradientBackground(slide, theme.gradients.primary[0], theme.gradients.primary[1]);
      } else {
        this.addBackground(slide, theme.background);
      }
      this.addSlideTitle(slide, spec.title || 'Slide', theme);
    },

    /**
     * Helper methods
     */
    addBackground: function (slide, color) {
      const bg = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 0, 0, this.SLIDE_WIDTH, this.SLIDE_HEIGHT);
      bg.getFill().setSolidFill(color);
      bg.getBorder().setTransparent();
      bg.sendToBack();
    },

    addGradientBackground: function (slide, color1, color2) {
      try {
        Logger.log('Adding gradient: ' + color1 + ' to ' + color2);

        // Create gradient effect with rectangles
        const bgTop = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 0, 0, this.SLIDE_WIDTH, this.SLIDE_HEIGHT / 2);
        bgTop.getFill().setSolidFill(color1);
        bgTop.getBorder().setTransparent();
        bgTop.sendToBack();

        const bgBottom = slide.insertShape(
          SlidesApp.ShapeType.RECTANGLE,
          0,
          this.SLIDE_HEIGHT / 2,
          this.SLIDE_WIDTH,
          this.SLIDE_HEIGHT / 2
        );
        bgBottom.getFill().setSolidFill(color2);
        bgBottom.getBorder().setTransparent();
        bgBottom.sendToBack();

        // Add blending overlay
        const blend = slide.insertShape(
          SlidesApp.ShapeType.RECTANGLE,
          0,
          this.SLIDE_HEIGHT * 0.4,
          this.SLIDE_WIDTH,
          this.SLIDE_HEIGHT * 0.2
        );
        blend.getFill().setSolidFill(color1);
        // Note: Transparency on fill not supported in Slides API
        blend.getBorder().setTransparent();
        blend.sendToBack();
      } catch (error) {
        Logger.error('Gradient failed, using solid:', error);
        this.addBackground(slide, color1);
      }
    },

    addSlideTitle: function (slide, title, theme) {
      const titleBox = slide.insertTextBox(
        title || 'Slide Title',
        this.SAFE_MARGINS.left,
        this.SAFE_MARGINS.top,
        this.SLIDE_WIDTH - this.SAFE_MARGINS.left - this.SAFE_MARGINS.right,
        40
      );
      titleBox
        .getText()
        .getTextStyle()
        .setFontSize(this.TYPOGRAPHY.title)
        .setFontFamily('Arial')
        .setBold(true)
        .setForegroundColor(theme.primary);

      // Add underline accent
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

      // Add accent dot
      const accentDot = slide.insertShape(
        SlidesApp.ShapeType.ELLIPSE,
        this.SAFE_MARGINS.left + underlineWidth + 10,
        this.SAFE_MARGINS.top + 33,
        6,
        6
      );
      accentDot.getFill().setSolidFill(theme.accent);
      // Note: Transparency on fill not supported in Slides API
      accentDot.getBorder().setTransparent();
    },

    addSpeakerNotes: function (slide, notes) {
      try {
        const notesPage = slide.getNotesPage();
        const speaker = notesPage.getSpeakerNotesShape();
        if (speaker) {
          speaker.getText().setText(notes);
        }
      } catch (e) {
        // Continue without notes
      }
    },

    addErrorMessage: function (slide, message, theme) {
      const errorBox = slide.insertTextBox(
        message || 'An error occurred',
        this.SAFE_MARGINS.left,
        200,
        this.SLIDE_WIDTH - this.SAFE_MARGINS.left - this.SAFE_MARGINS.right,
        30
      );
      errorBox
        .getText()
        .getTextStyle()
        .setFontSize(14)
        .setFontFamily('Arial')
        .setForegroundColor(theme.error)
        .setItalic(true);
      errorBox.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    },

    prepareChartData: function (data, analysis, chartType) {
      const maxPoints = 15;
      let xColumn = 0;
      let yColumn = 1;

      if (analysis.dateColumns.length > 0) {
        xColumn = analysis.dateColumns[0].index;
      } else if (analysis.textColumns.length > 0) {
        xColumn = analysis.textColumns[0].index;
      }

      if (analysis.numericColumns.length > 0) {
        yColumn = analysis.numericColumns[0].index;
      }

      const limitedData = {
        headers: [data.headers[xColumn], data.headers[yColumn]],
        data: data.data.slice(0, maxPoints).map((row) => [row[xColumn], row[yColumn]]),
      };

      return limitedData;
    },

    /**
     * LEGACY: Create a beautiful visual data chart using shapes
     * NOTE: This is the old shape-based system - prefer createSheetsChart for linked charts
     * REQUIRES: data must be an array (2D array rows from spreadsheet)
     */
    createVisualDataChart: function (slide, dataStructure, analysis, theme) {
      const startY = 120;
      const chartWidth = this.SLIDE_WIDTH - this.SAFE_MARGINS.left - this.SAFE_MARGINS.right;
      const chartHeight = 240;
      const chartX = this.SAFE_MARGINS.left;

      // Create glass-morphism background card
      const chartBg = slide.insertShape(
        SlidesApp.ShapeType.ROUND_RECTANGLE,
        chartX - 10,
        startY - 10,
        chartWidth + 20,
        chartHeight + 20
      );
      chartBg.getFill().setSolidFill(theme.surface);
      chartBg.getBorder().getLineFill().setSolidFill(theme.primary);
      chartBg.getBorder().setWeight(0.5);
      chartBg.getBorder().setDashStyle(SlidesApp.DashStyle.SOLID);

      // Extract data array from structure
      const dataArray =
        dataStructure && dataStructure.data ? dataStructure.data : Array.isArray(dataStructure) ? dataStructure : [];

      if (!dataArray || dataArray.length === 0) {
        Logger.error('createVisualDataChart: No data available');
        const noDataMsg = slide.insertTextBox(
          'No data available for visualization',
          chartX,
          startY + chartHeight / 2 - 20,
          chartWidth,
          40
        );
        noDataMsg.getText().getTextStyle().setFontSize(14).setForegroundColor(theme.textLight).setItalic(true);
        noDataMsg.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
        return;
      }

      // Use intelligent visualization suggestions
      const maxRows = 7;
      const displayData = dataArray.slice(0, maxRows);

      if (analysis.suggestedVisualizations && analysis.suggestedVisualizations.length > 0) {
        // Use the top suggested visualization
        const topSuggestion = analysis.suggestedVisualizations[0];
        Logger.log('Using suggested visualization: ' + topSuggestion.type);

        switch (topSuggestion.type) {
        case 'bar_chart':
        case 'grouped_bar':
          this.createIntelligentBarChart(
            slide,
            displayData,
            analysis,
            topSuggestion,
            theme,
            chartX,
            startY,
            chartWidth,
            chartHeight
          );
          break;
        case 'line_chart':
          this.createTimeSeriesChart(
            slide,
            displayData,
            analysis,
            topSuggestion,
            theme,
            chartX,
            startY,
            chartWidth,
            chartHeight
          );
          break;
        case 'pie_chart':
          this.createDistributionChart(
            slide,
            displayData,
            analysis,
            topSuggestion,
            theme,
            chartX,
            startY,
            chartWidth,
            chartHeight
          );
          break;
        case 'metric_cards':
          this.createMetricCards(slide, analysis.keyMetrics, theme, chartX, startY, chartWidth, chartHeight);
          break;
        default:
          // Fallback to standard bar chart
          this.createBarChartVisualization(
            slide,
            displayData,
            analysis,
            theme,
            chartX,
            startY,
            chartWidth,
            chartHeight
          );
        }

        // Add reason for visualization choice
        if (topSuggestion.reason) {
          const reasonBox = slide.insertTextBox(
            '💡 ' + (topSuggestion.reason || 'Optimized for your data'),
            chartX,
            startY + chartHeight + 10,
            chartWidth,
            20
          );
          reasonBox.getText().getTextStyle().setFontSize(10).setForegroundColor(theme.textLight).setItalic(true);
        }
      } else {
        // Fallback to original logic
        const numericColumns = analysis.numericColumns || [];
        const hasNumeric = numericColumns.length > 0;

        if (hasNumeric && displayData.length > 0) {
          this.createBarChartVisualization(
            slide,
            displayData,
            analysis,
            theme,
            chartX,
            startY,
            chartWidth,
            chartHeight
          );
        } else if (analysis.percentageColumns && analysis.percentageColumns.length > 0) {
          this.createPercentageVisualization(
            slide,
            displayData,
            analysis,
            theme,
            chartX,
            startY,
            chartWidth,
            chartHeight
          );
        } else {
          this.createDataCards(slide, displayData, analysis, theme, chartX, startY, chartWidth, chartHeight);
        }
      }

      // Add trend indicator if applicable
      if (analysis.trends && analysis.trends.length > 0) {
        const trendBox = slide.insertTextBox(
          '📈 Trend: ' + (analysis.trends[0] || 'No trends detected'),
          chartX,
          startY + chartHeight + 30,
          chartWidth,
          30
        );
        trendBox.getText().getTextStyle().setFontSize(11).setForegroundColor(theme.accent).setItalic(true);
      }
    },

    /**
     * Create intelligent bar chart based on data relationships
     */
    createIntelligentBarChart: function (slide, data, analysis, suggestion, theme, x, y, width, height) {
      Logger.log('Creating intelligent bar chart for: ' + suggestion.title);

      // Find the correct columns based on suggestion
      let categoryCol = 0;
      let valueCol = 1;

      if (suggestion.groupBy) {
        // Find column index by name
        const groupByIndex = analysis.headers.indexOf(suggestion.groupBy);
        if (groupByIndex !== -1) categoryCol = groupByIndex;
      } else if (analysis.categoryColumns.length > 0) {
        categoryCol = analysis.categoryColumns[0].index;
      }

      if (suggestion.measure) {
        // Find column index by name
        const measureIndex = analysis.headers.indexOf(suggestion.measure);
        if (measureIndex !== -1) valueCol = measureIndex;
      } else if (analysis.numericColumns.length > 0) {
        valueCol = analysis.numericColumns[0].index;
      }

      // Aggregate data by category if needed
      const aggregatedData = this.aggregateDataByCategory(data, categoryCol, valueCol);
      const sortedData = aggregatedData.sort((a, b) => b.value - a.value).slice(0, 7);

      if (sortedData.length === 0) {
        Logger.log('No data to display after aggregation');
        return;
      }

      const maxValue = Math.max(...sortedData.map((d) => d.value), 1);
      const barCount = sortedData.length;
      const barWidth = Math.max(30, (width - 100) / (barCount * 1.5));
      const barSpacing = barWidth * 0.5;
      const maxBarHeight = Math.max(50, height - 100);

      // Create axis
      const yAxis = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, x + 40, y + 30, 1, maxBarHeight);
      yAxis.getFill().setSolidFill(theme.textLight);
      yAxis.getBorder().setTransparent();

      const xAxis = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, x + 40, y + 30 + maxBarHeight, width - 80, 1);
      xAxis.getFill().setSolidFill(theme.textLight);
      xAxis.getBorder().setTransparent();

      // Create bars
      sortedData.forEach((item, index) => {
        const barHeight = Math.max(1, (item.value / maxValue) * maxBarHeight);
        const barX = x + 60 + index * (barWidth + barSpacing);
        const barY = y + 30 + maxBarHeight - barHeight;

        // Gradient colors for visual appeal
        const colors = [theme.primary, theme.accent, theme.secondary, '#34a853', '#fbbc04', '#ea4335', '#673ab7'];
        const barColor = colors[index % colors.length];

        // Bar with shadow effect
        if (barHeight > 2) {
          const shadow = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, barX + 2, barY + 2, barWidth, barHeight);
          shadow.getFill().setSolidFill('#e0e0e0');
          shadow.getBorder().setTransparent();
          shadow.sendToBack();
        }

        const bar = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, barX, barY, barWidth, barHeight);
        bar.getFill().setSolidFill(barColor);
        bar.getBorder().setTransparent();

        // Value label with intelligent formatting
        let valueText;
        if (suggestion.measure && suggestion.measure.toLowerCase().includes('salary')) {
          valueText = '$' + (item.value > 1000 ? (item.value / 1000).toFixed(0) + 'K' : item.value.toFixed(0));
        } else if (item.value > 1000000) {
          valueText = (item.value / 1000000).toFixed(1) + 'M';
        } else if (item.value > 1000) {
          valueText = (item.value / 1000).toFixed(1) + 'K';
        } else {
          valueText = item.value.toFixed(0);
        }

        const valueLabel = slide.insertTextBox(valueText || '0', barX, barY - 25, barWidth, 20);
        valueLabel.getText().getTextStyle().setFontSize(10).setForegroundColor(barColor).setBold(true);
        valueLabel.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

        // Category label
        const categoryLabel = slide.insertTextBox(
          item.category ? item.category.toString().substring(0, 10) : 'Unknown',
          barX - 5,
          y + 30 + maxBarHeight + 5,
          barWidth + 10,
          30
        );
        categoryLabel.getText().getTextStyle().setFontSize(9).setForegroundColor(theme.textLight);
        categoryLabel.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

        // Add count if aggregated
        if (item.count > 1) {
          const countLabel = slide.insertTextBox(
            '(n=' + item.count + ')',
            barX,
            y + 30 + maxBarHeight + 30,
            barWidth,
            15
          );
          countLabel.getText().getTextStyle().setFontSize(8).setForegroundColor(theme.textLight).setItalic(true);
          countLabel.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
        }
      });

      // Add axis labels
      const yLabel = slide.insertTextBox(
        suggestion.measure || analysis.headers[valueCol] || 'Values',
        x + 5,
        y + height / 2 - 20,
        30,
        40
      );
      yLabel.getText().getTextStyle().setFontSize(10).setForegroundColor(theme.textLight);
      yLabel.setRotation(-90);

      const xLabel = slide.insertTextBox(
        suggestion.groupBy || analysis.headers[categoryCol] || 'Categories',
        x + width / 2 - 50,
        y + height - 20,
        100,
        20
      );
      xLabel.getText().getTextStyle().setFontSize(10).setForegroundColor(theme.textLight);
      xLabel.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    },

    /**
     * Aggregate data by category
     */
    aggregateDataByCategory: function (data, categoryCol, valueCol) {
      const aggregated = {};

      data.forEach((row) => {
        const category = row[categoryCol] || 'Unknown';
        const value = parseFloat(row[valueCol]) || 0;

        if (!aggregated[category]) {
          aggregated[category] = {
            category: category,
            value: 0,
            count: 0,
          };
        }

        aggregated[category].value += value;
        aggregated[category].count += 1;
      });

      // Calculate averages for better representation
      return Object.values(aggregated).map((item) => ({
        category: item.category,
        value: item.value / item.count, // Use average instead of sum
        count: item.count,
      }));
    },

    /**
     * Create modern bar chart visualization
     */
    createBarChartVisualization: function (slide, data, analysis, theme, x, y, width, height) {
      // Safely get numeric column
      if (!analysis.numericColumns || analysis.numericColumns.length === 0) {
        Logger.log('No numeric columns found for bar chart visualization');
        return;
      }
      const numericCol = analysis.numericColumns[0];
      const values = data.map((row) => parseFloat(row[numericCol]) || 0);
      const maxValue = Math.max(...values, 1);
      const barCount = Math.min(values.length, 7);
      const barWidth = Math.max(20, (width - 80) / (barCount * 1.5)); // Minimum bar width
      const barSpacing = barWidth * 0.5;
      const maxBarHeight = Math.max(50, height - 80); // Minimum bar chart height

      // Y-axis line
      const yAxis = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, x + 30, y + 30, 1, maxBarHeight);
      yAxis.getFill().setSolidFill(theme.textLight);
      yAxis.getBorder().setTransparent();

      // X-axis line
      const xAxis = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, x + 30, y + 30 + maxBarHeight, width - 60, 1);
      xAxis.getFill().setSolidFill(theme.textLight);
      xAxis.getBorder().setTransparent();

      // Create bars with gradient effect
      values.slice(0, barCount).forEach((value, index) => {
        const barHeight = Math.max(1, (value / maxValue) * maxBarHeight); // Ensure minimum height of 1
        const barX = x + 50 + index * (barWidth + barSpacing);
        const barY = y + 30 + maxBarHeight - barHeight;

        // Only create shadow if bar has meaningful height
        if (barHeight > 2) {
          const shadow = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, barX + 2, barY + 2, barWidth, barHeight);
          shadow.getFill().setSolidFill('#e0e0e0');
          shadow.getBorder().setTransparent();
          shadow.sendToBack();
        }

        // Main bar with gradient colors
        const colors = [theme.primary, theme.accent, theme.secondary, '#34a853', '#fbbc04', '#ea4335', '#673ab7'];
        const barColor = colors[index % colors.length];

        const bar = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, barX, barY, barWidth, barHeight);
        bar.getFill().setSolidFill(barColor);
        bar.getBorder().setTransparent();

        // Rounded top for bars
        if (barHeight > 5) {
          const barTop = slide.insertShape(
            SlidesApp.ShapeType.ROUND_RECTANGLE,
            barX,
            barY,
            barWidth,
            Math.min(10, barHeight) // Don't make rounded top larger than bar
          );
          barTop.getFill().setSolidFill(barColor);
          barTop.getBorder().setTransparent();
        }

        // Value label on top
        const valueText = value > 1000 ? (value / 1000).toFixed(1) + 'K' : value.toFixed(0);
        const valueLabel = slide.insertTextBox(valueText, barX, barY - 25, barWidth, 20);
        valueLabel.getText().getTextStyle().setFontSize(10).setForegroundColor(barColor).setBold(true);
        valueLabel.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

        // Category label
        const categoryText = data[index][0] ? data[index][0].toString().substring(0, 8) : 'Item';
        const label = slide.insertTextBox(categoryText, barX - 5, y + 30 + maxBarHeight + 5, barWidth + 10, 25);
        label.getText().getTextStyle().setFontSize(9).setForegroundColor(theme.textLight);
        label.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
      });

      // Add axis labels
      const yLabel = slide.insertTextBox(analysis.headers[numericCol] || 'Values', x, y + height / 2 - 20, 25, 40);
      yLabel.getText().getTextStyle().setFontSize(10).setForegroundColor(theme.textLight);
      yLabel.setRotation(-90);
    },

    /**
     * Create percentage visualization with progress bars
     */
    createPercentageVisualization: function (slide, data, analysis, theme, x, y, width, height) {
      // Ensure data is an array
      const dataArray = Array.isArray(data) ? data : data && data.rows ? data.rows : [];
      if (!dataArray || dataArray.length === 0) return;

      // Safely get percentage column
      if (!analysis.percentageColumns || analysis.percentageColumns.length === 0) {
        Logger.log('No percentage columns found for percentage visualization');
        return;
      }
      const percentCol = analysis.percentageColumns[0];
      const maxRows = Math.min(dataArray.length, 5);
      const rowHeight = Math.max(20, (height - 60) / maxRows); // Ensure minimum row height

      dataArray.slice(0, maxRows).forEach((row, index) => {
        const rowY = y + 40 + index * rowHeight;
        const percentage = Math.max(0, Math.min(100, parseFloat(row[percentCol]) || 0)); // Clamp 0-100
        const barWidth = Math.max(1, (width - 200) * (percentage / 100)); // Minimum width of 1

        // Label
        const label = slide.insertTextBox(row[0] ? row[0].toString() : '', x + 10, rowY, 150, 25);
        label.getText().getTextStyle().setFontSize(11).setForegroundColor(theme.text);

        // Progress bar background
        const barBg = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, x + 170, rowY + 5, width - 200, 15);
        barBg.getFill().setSolidFill(theme.surfaceVariant || theme.surface);
        barBg.getBorder().setTransparent();

        // Progress bar fill
        if (barWidth > 0) {
          const barFill = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, x + 170, rowY + 5, barWidth, 15);
          const colors = [theme.primary, theme.accent, '#34a853', '#fbbc04', '#ea4335'];
          barFill.getFill().setSolidFill(colors[index % colors.length]);
          barFill.getBorder().setTransparent();
        }

        // Percentage label
        const percentLabel = slide.insertTextBox(percentage.toFixed(1) + '%', x + width - 40, rowY, 35, 25);
        percentLabel.getText().getTextStyle().setFontSize(10).setForegroundColor(theme.textLight).setBold(true);
        percentLabel.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.END);
      });
    },

    /**
     * Create stylish data cards for non-numeric data
     */
    createDataCards: function (slide, data, analysis, theme, x, y, width, height) {
      const cardCount = Math.min(data.length, 4);
      const cardWidth = Math.max(80, (width - (cardCount - 1) * 10) / cardCount); // Minimum card width
      const cardHeight = Math.max(100, height - 40); // Minimum card height

      data.slice(0, cardCount).forEach((row, index) => {
        const cardX = x + index * (cardWidth + 10);
        const cardY = y + 20;

        // Card background with shadow
        const shadow = slide.insertShape(
          SlidesApp.ShapeType.ROUND_RECTANGLE,
          cardX + 3,
          cardY + 3,
          cardWidth,
          cardHeight
        );
        shadow.getFill().setSolidFill('#f5f5f5');
        shadow.getBorder().setTransparent();
        shadow.sendToBack();

        // Main card
        const card = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, cardX, cardY, cardWidth, cardHeight);
        card.getFill().setSolidFill(theme.surface);
        card.getBorder().getLineFill().setSolidFill(theme.primary);
        card.getBorder().setWeight(1);

        // Card accent bar
        const accentBar = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, cardX, cardY, cardWidth, 4);
        const colors = [theme.primary, theme.accent, theme.secondary, '#34a853'];
        accentBar.getFill().setSolidFill(colors[index % colors.length]);
        accentBar.getBorder().setTransparent();

        // Card content
        let contentY = cardY + 20;
        row.slice(0, 3).forEach((cell, cellIndex) => {
          if (cell) {
            const isTitle = cellIndex === 0;
            const textBox = slide.insertTextBox(
              this.formatCellValue(cell) || '',
              cardX + 10,
              contentY,
              cardWidth - 20,
              isTitle ? 30 : 25
            );
            textBox
              .getText()
              .getTextStyle()
              .setFontSize(isTitle ? 12 : 10)
              .setForegroundColor(isTitle ? theme.text : theme.textLight)
              .setBold(isTitle);
            textBox.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

            contentY += isTitle ? 35 : 30;
          }
        });
      });
    },

    /**
     * Keep the original buildChart for compatibility
     */
    buildChart: function (chartData, chartType, theme) {
      const dataTable = Charts.newDataTable();

      chartData.headers.forEach((header, index) => {
        if (index === 0) {
          dataTable.addColumn(Charts.ColumnType.STRING, header);
        } else {
          dataTable.addColumn(Charts.ColumnType.NUMBER, header);
        }
      });

      chartData.data.forEach((row) => {
        const rowData = row.map((val, index) => {
          if (index === 0) return val ? val.toString() : '';
          return parseFloat(val) || 0;
        });
        dataTable.addRow(rowData);
      });

      let chartBuilder;

      switch (chartType) {
      case Charts.ChartType.LINE:
        chartBuilder = Charts.newLineChart();
        break;
      case Charts.ChartType.BAR:
        chartBuilder = Charts.newBarChart();
        break;
      case Charts.ChartType.SCATTER:
        chartBuilder = Charts.newScatterChart();
        break;
      case Charts.ChartType.PIE:
        chartBuilder = Charts.newPieChart();
        break;
      case Charts.ChartType.AREA:
        chartBuilder = Charts.newAreaChart();
        break;
      default:
        chartBuilder = Charts.newColumnChart();
      }

      chartBuilder
        .setDataTable(dataTable.build())
        .setOption('backgroundColor', theme.background)
        .setOption('colors', [theme.primary, theme.accent, theme.secondary])
        .setOption('legend', { position: 'bottom' })
        .setOption('titleTextStyle', {
          color: theme.text,
          fontSize: 16,
          bold: true,
        })
        .setOption('chartArea', { width: '85%', height: '75%' })
        .setDimensions(600, 400);

      return chartBuilder.build();
    },

    formatCellValue: function (value) {
      if (value == null || value === '') return '';

      const str = value.toString();

      if (str.length > 20) {
        return str.substring(0, 17) + '...';
      }

      if (!isNaN(parseFloat(value)) && isFinite(value)) {
        const num = parseFloat(value);
        if (num % 1 === 0) {
          return num.toString();
        }
        return num.toFixed(2);
      }

      return str;
    },

    applyFinishingTouches: function (presentation, config) {
      const slides = presentation.getSlides();
      const totalSlides = slides.length;

      slides.forEach((slide, index) => {
        if (index > 0 && index < totalSlides - 1) {
          const slideNumber = index + 1 + ' / ' + totalSlides;
          const numberBox = slide.insertTextBox(slideNumber, this.SLIDE_WIDTH - 80, this.SLIDE_HEIGHT - 25, 70, 20);
          numberBox
            .getText()
            .getTextStyle()
            .setFontSize(this.TYPOGRAPHY.overline)
            .setFontFamily('Arial')
            .setForegroundColor('#9aa0a6');
          numberBox.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.END);
        }
      });

      presentation.setName(config.title || 'Data Presentation');
      Logger.log('Applied finishing touches to presentation');
    },

    /**
     * Create dashboard layout based on config
     * Phase 2 function for professional dashboard layouts
     */
    createDashboardLayout: function (slide, config) {
      try {
        Logger.log('Creating dashboard layout with config:', config);
        // Extract layout type and theme from config
        const layoutType = config.type || 'executive';
        const theme = config.theme || this.PROFESSIONAL_THEMES.executive;
        const data = config.data || {};

        // Map layout types to dashboard layouts
        const layoutMap = {
          executive: 'executive_summary',
          quad: 'quad_view',
          focus: 'focus_detail',
          timeline: 'timeline_view',
          comparison: 'comparison_view',
        };

        const layoutName = layoutMap[layoutType] || 'executive_summary';

        // For now, use a simple fallback since createDashboardWithLayout may not exist
        // TODO: Implement full dashboard layout system
        Logger.log('Creating dashboard layout type: ' + layoutName);

        // Simple implementation for now
        this.addBackground(slide, theme.background || '#FFFFFF');
        this.addSlideTitle(slide, data.title || 'Executive Summary', theme);

        // Add metrics if provided
        if (data.metrics && data.metrics.length > 0) {
          const metrics = data.metrics;
          const cardWidth = 140;
          const cardHeight = 80;
          const spacing = 20;
          const startX = (this.SLIDE_WIDTH - (metrics.length * cardWidth + (metrics.length - 1) * spacing)) / 2;
          const startY = 120;

          metrics.forEach((metric, index) => {
            const x = startX + index * (cardWidth + spacing);

            // Card background
            const card = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, x, startY, cardWidth, cardHeight);
            card.getFill().setSolidFill(theme.surface || '#F8F9FA');
            card
              .getBorder()
              .getLineFill()
              .setSolidFill(theme.primary || '#4285F4');
            card.getBorder().setWeight(1);

            // Metric value
            const valueBox = slide.insertTextBox(
              metric.value ? metric.value.toString() : '0',
              x,
              startY + 15,
              cardWidth,
              30
            );
            valueBox
              .getText()
              .getTextStyle()
              .setFontSize(24)
              .setFontFamily('Arial')
              .setBold(true)
              .setForegroundColor(theme.primary || '#4285F4');
            valueBox.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

            // Metric label
            const labelBox = slide.insertTextBox(metric.label || '', x, startY + 45, cardWidth, 20);
            labelBox
              .getText()
              .getTextStyle()
              .setFontSize(12)
              .setFontFamily('Arial')
              .setForegroundColor(theme.textLight || '#5F6368');
            labelBox.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
          });
        }

        return true;
      } catch (error) {
        Logger.error('Error in createDashboardLayout:', error);
        // Fallback to simple layout
        this.addBackground(slide, '#FFFFFF');
        this.addSlideTitle(slide, 'Dashboard', {});
        return false;
      }
    },
  },

  /**
   * Wrapper functions for UI calls from CellM8Template.html
   */

  // Extract data for UI (wrapper)
  extractCellM8Data: function (options) {
    return this.extractSheetData(options);
  },

  // Analyze data for UI (wrapper)
  analyzeCellM8Data: function (data) {
    try {
      // Prepare data structure for analysis
      const dataStructure = {
        headers: data.headers || [],
        data: data.data || [],
      };

      const analysis = this.SlideGenerator.analyzeDataIntelligently(dataStructure);

      return {
        success: true,
        analysis: analysis,
      };
    } catch (error) {
      Logger.error('Failed to analyze data:', error);
      return {
        success: false,
        error: error.toString(),
      };
    }
  },

  // Preview presentation for UI (wrapper)
  previewCellM8Presentation: function (config) {
    try {
      // Generate a preview structure without creating actual slides
      const analysis = config.analysis || {};
      const slideCount = config.slideCount || 5;

      const slides = [];

      // Title slide
      slides.push({
        title: config.title,
        subtitle: config.subtitle || 'Data Presentation',
        type: 'title',
      });

      // Data overview slide
      if (analysis.totalRows) {
        slides.push({
          title: 'Data Overview',
          bullets: [
            'Total Records: ' + analysis.totalRows,
            'Data Columns: ' + analysis.totalColumns,
            'Data Quality: ' + (analysis.completeness || 'N/A') + '% complete',
          ],
          type: 'overview',
        });
      }

      // Add more slides based on analysis
      if (analysis.keyMetrics && analysis.keyMetrics.length > 0) {
        slides.push({
          title: 'Key Metrics',
          bullets: analysis.keyMetrics.slice(0, 5),
          type: 'metrics',
        });
      }

      // Conclusion slide
      slides.push({
        title: 'Summary',
        subtitle: 'Key findings from your data',
        type: 'conclusion',
      });

      return {
        success: true,
        slides: slides,
        slideCount: slides.length,
        dataInfo: {
          rows: analysis.totalRows || 0,
          columns: analysis.totalColumns || 0,
        },
      };
    } catch (error) {
      Logger.error('Failed to preview presentation:', error);
      return {
        success: false,
        error: error.toString(),
      };
    }
  },

  // Create presentation for UI (wrapper)
  createCellM8Presentation: function (config) {
    return this.createPresentation(config);
  },

  /**
   * Refresh all linked charts in a presentation
   */
  refreshPresentationCharts: function (presentationId) {
    try {
      const presentation = presentationId ? SlidesApp.openById(presentationId) : SlidesApp.getActivePresentation();

      if (!presentation) {
        return {
          success: false,
          error: 'No presentation found',
        };
      }

      let chartsRefreshed = 0;
      const slides = presentation.getSlides();

      slides.forEach((slide) => {
        const sheetsCharts = slide.getSheetsCharts();
        sheetsCharts.forEach((chart) => {
          try {
            chart.refresh();
            chartsRefreshed++;
          } catch (e) {
            Logger.log('Failed to refresh chart: ' + e.toString());
          }
        });
      });

      return {
        success: true,
        chartsRefreshed: chartsRefreshed,
        message: `Refreshed ${chartsRefreshed} linked charts`,
      };
    } catch (error) {
      Logger.error('Error refreshing presentation charts:', error);
      return {
        success: false,
        error: error.toString(),
      };
    }
  },

  /**
   * Clean up temporary charts created in Sheets
   */
  cleanupSheetsCharts: function (keepCount) {
    try {
      const sheet = SpreadsheetApp.getActiveSheet();
      const charts = sheet.getCharts();
      const toKeep = keepCount || 0;

      // Remove charts beyond the keep count
      if (charts.length > toKeep) {
        for (let i = charts.length - 1; i >= toKeep; i--) {
          sheet.removeChart(charts[i]);
        }
      }

      return {
        success: true,
        removed: Math.max(0, charts.length - toKeep),
      };
    } catch (error) {
      Logger.error('Error cleaning up charts:', error);
      return {
        success: false,
        error: error.toString(),
      };
    }
  },

  /**
   * Create a dashboard slide with multiple charts
   */
  createDashboardSlide: function (slide, data, analysis, theme) {
    try {
      this.addBackground(slide, theme.background);
      this.addSlideTitle(slide, 'Data Dashboard', theme);

      const sheet = SpreadsheetApp.getActiveSheet();
      const charts = [];

      // Create up to 4 different chart types for the dashboard
      const chartTypes = ['column', 'pie', 'line', 'table'];
      const positions = [
        { left: 40, top: 90, width: 320, height: 180 },
        { left: 380, top: 90, width: 320, height: 180 },
        { left: 40, top: 290, width: 320, height: 180 },
        { left: 380, top: 290, width: 320, height: 180 },
      ];

      chartTypes.forEach((type, index) => {
        if (index < Math.min(4, analysis.numericColumns.length)) {
          const chartType = this.mapToSheetsChartType(type + '_chart');
          const dataRange = this.selectChartDataRange(sheet, data, analysis);

          if (dataRange) {
            const chart = sheet
              .newChart()
              .setChartType(chartType)
              .addRange(dataRange)
              .setPosition(1, 15 + index * 2, 0, 0)
              .setOption('title', this.getChartTitle(type, analysis))
              .setOption('width', 300)
              .setOption('height', 200)
              .build();

            sheet.insertChart(chart);

            // Insert into slide
            slide.insertSheetsChart(
              chart,
              positions[index].left,
              positions[index].top,
              positions[index].width,
              positions[index].height
            );

            charts.push(chart);
          }
        }
      });

      return {
        success: true,
        charts: charts.length,
      };
    } catch (error) {
      Logger.error('Error creating dashboard slide:', error);
      return {
        success: false,
        error: error.toString(),
      };
    }
  },

  /**
   * Get appropriate chart title based on type
   */
  getChartTitle: function (type, analysis) {
    const titles = {
      column: 'Category Comparison',
      pie: 'Distribution Analysis',
      line: 'Trend Overview',
      table: 'Key Metrics',
    };
    return titles[type] || 'Data Visualization';
  },

  // Dashboard layout templates
  DASHBOARD_LAYOUTS: {
    executive_summary: {
      name: 'Executive Summary',
      zones: [
        { id: 'kpi', type: 'metrics', position: { x: 40, y: 80, w: 680, h: 60 } },
        { id: 'main', type: 'chart', position: { x: 40, y: 160, w: 440, h: 260 } },
        { id: 'secondary', type: 'chart', position: { x: 500, y: 160, w: 220, h: 260 } },
        { id: 'insights', type: 'text', position: { x: 40, y: 440, w: 680, h: 60 } },
      ],
    },
    quad_view: {
      name: 'Quad View',
      zones: [
        { id: 'tl', type: 'chart', position: { x: 40, y: 90, w: 340, h: 200 } },
        { id: 'tr', type: 'chart', position: { x: 400, y: 90, w: 340, h: 200 } },
        { id: 'bl', type: 'chart', position: { x: 40, y: 310, w: 340, h: 200 } },
        { id: 'br', type: 'chart', position: { x: 400, y: 310, w: 340, h: 200 } },
      ],
    },
    focus_detail: {
      name: 'Focus + Detail',
      zones: [
        { id: 'focus', type: 'chart', position: { x: 40, y: 90, w: 480, h: 340 } },
        { id: 'detail1', type: 'table', position: { x: 540, y: 90, w: 200, h: 160 } },
        { id: 'detail2', type: 'metrics', position: { x: 540, y: 270, w: 200, h: 160 } },
      ],
    },
    timeline: {
      name: 'Timeline View',
      zones: [
        { id: 'timeline', type: 'chart', position: { x: 40, y: 90, w: 680, h: 200 } },
        { id: 'details', type: 'table', position: { x: 40, y: 310, w: 680, h: 120 } },
        { id: 'summary', type: 'metrics', position: { x: 40, y: 450, w: 680, h: 60 } },
      ],
    },
    comparison: {
      name: 'Side-by-Side',
      zones: [
        { id: 'left_chart', type: 'chart', position: { x: 40, y: 120, w: 340, h: 280 } },
        { id: 'right_chart', type: 'chart', position: { x: 400, y: 120, w: 340, h: 280 } },
        { id: 'comparison', type: 'table', position: { x: 40, y: 420, w: 680, h: 80 } },
      ],
    },
  },

  // Dashboard creation helper functions for different layouts
  createExecutiveDashboard: function (slide, data, theme) {
    return this.createDashboardWithLayout(slide, 'executive_summary', data, {}, theme);
  },

  createQuadViewDashboard: function (slide, charts, theme) {
    return this.createDashboardWithLayout(slide, 'quad_view', { charts: charts }, {}, theme);
  },

  createFocusDetailDashboard: function (slide, mainChart, detailData, theme) {
    return this.createDashboardWithLayout(
      slide,
      'focus_detail',
      { mainChart: mainChart, details: detailData },
      {},
      theme
    );
  },

  createTimelineDashboard: function (slide, timelineData, theme) {
    return this.createDashboardWithLayout(slide, 'timeline_view', { timeline: timelineData }, {}, theme);
  },

  createComparisonDashboard: function (slide, comparisonData, theme) {
    return this.createDashboardWithLayout(slide, 'comparison_view', { comparison: comparisonData }, {}, theme);
  },

  /**
   * Create dashboard with specified layout
   */
  createDashboardWithLayout: function (slide, layoutName, data, analysis, theme) {
    try {
      const layout = this.DASHBOARD_LAYOUTS[layoutName] || this.DASHBOARD_LAYOUTS.quad_view;

      // Clear slide and set background
      this.clearSlide(slide);
      this.addBackground(slide, theme.background);
      this.addSlideTitle(slide, layout.name, theme);

      const sheet = SpreadsheetApp.getActiveSheet();
      const createdElements = [];

      // Process each zone
      layout.zones.forEach((zone) => {
        const element = this.createDashboardZone(slide, zone, data, analysis, theme, sheet);
        if (element) {
          createdElements.push(element);
        }
      });

      return {
        success: true,
        layout: layout.name,
        elements: createdElements.length,
      };
    } catch (error) {
      Logger.error('Error creating dashboard with layout:', error);
      return {
        success: false,
        error: error.toString(),
      };
    }
  },

  /**
   * Create a single dashboard zone
   */
  createDashboardZone: function (slide, zone, data, analysis, theme, sheet) {
    try {
      switch (zone.type) {
      case 'chart':
        return this.addChartToZone(slide, zone, data, analysis, theme, sheet);

      case 'table':
        return this.addTableToZone(slide, zone, data, theme);

      case 'metrics':
        return this.addMetricsToZone(slide, zone, analysis.keyMetrics, theme);

      case 'text':
        return this.addTextToZone(slide, zone, analysis.insights, theme);

      default:
        return null;
      }
    } catch (error) {
      Logger.error('Error creating dashboard zone:', error);
      return null;
    }
  },

  /**
   * Add chart to dashboard zone
   */
  addChartToZone: function (slide, zone, data, analysis, theme, sheet) {
    // Determine best chart type for this zone
    const chartType = this.selectZoneChartType(zone.id, analysis);
    const dataRange = this.selectChartDataRange(sheet, data, analysis);

    if (!dataRange) return null;

    // Create chart with zone-specific configuration
    const chartBuilder = sheet
      .newChart()
      .setChartType(chartType)
      .addRange(dataRange)
      .setPosition(1, 20 + Math.floor(zone.position.x / 50), 0, 0)
      .setOption('width', zone.position.w * 0.8)
      .setOption('height', zone.position.h * 0.8);

    // Apply theme styling
    this.configureAdvancedChart(chartBuilder, {
      style: 'dashboard',
      theme: theme.name || 'executive',
      animation: 'subtle',
    });

    const chart = chartBuilder.build();
    sheet.insertChart(chart);

    // Insert into slide
    slide.insertSheetsChart(chart, zone.position.x, zone.position.y, zone.position.w, zone.position.h);

    return { type: 'chart', zone: zone.id };
  },

  /**
   * Add table to dashboard zone
   */
  addTableToZone: function (slide, zone, data, theme) {
    const maxRows = Math.floor(zone.position.h / 25) - 1; // Calculate rows that fit
    const maxCols = Math.min(data.headers.length, 4); // Limit columns for readability

    const table = slide.insertTable(
      Math.min(data.data.length, maxRows) + 1,
      maxCols,
      zone.position.x,
      zone.position.y,
      zone.position.w,
      zone.position.h
    );

    // Style the table
    this.styleCompactTable(table, data, theme);

    return { type: 'table', zone: zone.id };
  },

  /**
   * Add metrics cards to dashboard zone
   */
  addMetricsToZone: function (slide, zone, metrics, theme) {
    if (!metrics || metrics.length === 0) return null;

    const cardCount = Math.min(metrics.length, 4);
    const cardWidth = (zone.position.w - (cardCount - 1) * 10) / cardCount;

    metrics.slice(0, cardCount).forEach((metric, index) => {
      const cardX = zone.position.x + index * (cardWidth + 10);

      // Create metric card
      const card = slide.insertShape(
        SlidesApp.ShapeType.ROUND_RECTANGLE,
        cardX,
        zone.position.y,
        cardWidth,
        zone.position.h
      );
      card.getFill().setSolidFill(theme.surface || '#F8F9FA');
      card
        .getBorder()
        .getLineFill()
        .setSolidFill(theme.primary || '#1B2951');
      card.getBorder().setWeight(1);

      // Add metric value
      const valueText = this.formatMetricValue(metric);
      const valueBox = slide.insertTextBox(
        valueText,
        cardX + 5,
        zone.position.y + 10,
        cardWidth - 10,
        zone.position.h * 0.5
      );
      valueBox
        .getText()
        .getTextStyle()
        .setFontSize(20)
        .setBold(true)
        .setForegroundColor(theme.primary || '#1B2951');
      valueBox.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

      // Add metric label
      const labelBox = slide.insertTextBox(
        metric.label,
        cardX + 5,
        zone.position.y + zone.position.h * 0.6,
        cardWidth - 10,
        zone.position.h * 0.3
      );
      labelBox
        .getText()
        .getTextStyle()
        .setFontSize(10)
        .setForegroundColor(theme.textLight || '#666666');
      labelBox.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    });

    return { type: 'metrics', zone: zone.id };
  },

  /**
   * Add text/insights to dashboard zone
   */
  addTextToZone: function (slide, zone, insights, theme) {
    if (!insights || insights.length === 0) return null;

    const text = insights.slice(0, 3).join(' • ');

    const textBox = slide.insertTextBox(text || '', zone.position.x, zone.position.y, zone.position.w, zone.position.h);

    textBox
      .getText()
      .getTextStyle()
      .setFontSize(11)
      .setForegroundColor(theme.text || '#333333');
    textBox.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.START);

    return { type: 'text', zone: zone.id };
  },

  /**
   * Select appropriate chart type for zone
   */
  selectZoneChartType: function (zoneId, analysis) {
    const zoneChartMapping = {
      main: Charts.ChartType.COLUMN,
      secondary: Charts.ChartType.PIE,
      timeline: Charts.ChartType.LINE,
      focus: Charts.ChartType.COMBO,
      tl: Charts.ChartType.BAR,
      tr: Charts.ChartType.PIE,
      bl: Charts.ChartType.LINE,
      br: Charts.ChartType.SCATTER,
      left_chart: Charts.ChartType.COLUMN,
      right_chart: Charts.ChartType.COLUMN,
    };

    return zoneChartMapping[zoneId] || Charts.ChartType.COLUMN;
  },

  /**
   * Style compact table for dashboard
   */
  styleCompactTable: function (table, data, theme) {
    const numRows = table.getNumRows();
    const numCols = table.getNumColumns();

    // Compact header
    for (let col = 0; col < numCols; col++) {
      const cell = table.getCell(0, col);
      cell.getFill().setSolidFill(theme.primary || '#1B2951');
      cell.getText().setText(data.headers[col] || '');
      cell
        .getText()
        .getTextStyle()
        .setFontSize(9)
        .setForegroundColor(theme.darkMode ? '#FFFFFF' : theme.text || '#2C3E50')
        .setBold(true);
    }

    // Compact data rows
    for (let row = 1; row < numRows; row++) {
      for (let col = 0; col < numCols; col++) {
        const cell = table.getCell(row, col);
        const bgColor = row % 2 === 0 ? theme.background : theme.surface;
        cell.getFill().setSolidFill(bgColor || '#FFFFFF');

        const value = data.data[row - 1] ? data.data[row - 1][col] : '';
        cell.getText().setText(this.formatCellValue(value));
        cell
          .getText()
          .getTextStyle()
          .setFontSize(8)
          .setForegroundColor(theme.text || '#333333');
      }
    }
  },

  /**
   * Format metric value for display
   */
  formatMetricValue: function (metric) {
    if (!metric.value) return '—';

    if (metric.type === 'monetary') {
      return '$' + this.formatNumber(metric.value);
    } else if (metric.type === 'percentage') {
      return metric.value + '%';
    } else {
      return this.formatNumber(metric.value);
    }
  },

  /**
   * Format number with appropriate abbreviation
   */
  formatNumber: function (num) {
    if (num >= 1000000) {
      return (num / 1000000).toFixed(1) + 'M';
    } else if (num >= 1000) {
      return (num / 1000).toFixed(1) + 'K';
    } else {
      return num.toFixed(0);
    }
  },

  /**
   * Clear all elements from slide
   */
  clearSlide: function (slide) {
    const elements = slide.getPageElements();
    elements.forEach((element) => {
      try {
        element.remove();
      } catch (e) {
        // Continue if element cannot be removed
      }
    });
  },
};
