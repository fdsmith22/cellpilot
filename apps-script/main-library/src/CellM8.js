/**
 * CellM8 - Presentation Generator Module
 * Creates professional presentations from spreadsheet data
 */

const CellM8 = {
  /**
   * Initialize CellM8
   */
  init: function() {
    Logger.log('CellM8 initialized');
    return { success: true };
  },
  
  /**
   * Show CellM8 sidebar
   */
  showSidebar: function() {
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
  testFunction: function() {
    return {
      success: true,
      message: 'CellM8 is working properly',
      timestamp: new Date().toISOString()
    };
  },
  
  /**
   * Get current selection info
   */
  getCurrentSelection: function() {
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
          sheet: sheet.getName()
        };
      } else {
        const dataRange = sheet.getDataRange();
        return {
          success: true,
          hasSelection: false,
          range: dataRange.getA1Notation(),
          rows: dataRange.getNumRows(),
          cols: dataRange.getNumColumns(),
          sheet: sheet.getName()
        };
      }
    } catch (error) {
      Logger.error('Error getting selection:', error);
      return {
        success: false,
        error: error.toString()
      };
    }
  },
  
  /**
   * Select entire data range
   */
  selectEntireDataRange: function() {
    try {
      const sheet = SpreadsheetApp.getActiveSheet();
      const dataRange = sheet.getDataRange();
      sheet.setActiveRange(dataRange);
      
      return {
        success: true,
        range: dataRange.getA1Notation(),
        rows: dataRange.getNumRows(),
        cols: dataRange.getNumColumns()
      };
    } catch (error) {
      Logger.error('Error selecting data range:', error);
      return {
        success: false,
        error: error.toString()
      };
    }
  },
  
  /**
   * Select specific range
   */
  selectRange: function(rangeA1) {
    try {
      const sheet = SpreadsheetApp.getActiveSheet();
      const range = sheet.getRange(rangeA1);
      sheet.setActiveRange(range);
      
      return {
        success: true,
        range: range.getA1Notation(),
        rows: range.getNumRows(),
        cols: range.getNumColumns()
      };
    } catch (error) {
      Logger.error('Error selecting range:', error);
      return {
        success: false,
        error: error.toString()
      };
    }
  },
  
  /**
   * Generate preview of presentation
   */
  previewPresentation: function(config) {
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
          headers: dataResult.headers
        },
        slides: []
      };
      
      // Add preview slides
      preview.slides.push({
        type: 'title',
        title: config.title || 'Data Presentation',
        subtitle: config.subtitle || `Analysis of ${dataResult.rowCount} records`
      });
      
      if (config.slideCount >= 3) {
        preview.slides.push({
          type: 'overview',
          title: 'Data Overview',
          content: `Dataset contains ${dataResult.rowCount} rows and ${dataResult.columnCount} columns`
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
   * Create a presentation with data
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

      // Extract data if not provided
      let dataResult = null;
      if (!config.data) {
        dataResult = this.extractSheetData();
        if (!dataResult.success) {
          return {
            success: false,
            error: 'Failed to extract data: ' + dataResult.error
          };
        }
      } else {
        dataResult = config.data;
      }
      
      // Check if CellM8.SlideGenerator exists (as a property of CellM8)
      Logger.log('=== DEBUGGING CellM8.SlideGenerator ===');
      Logger.log('Type of CellM8.SlideGenerator: ' + typeof CellM8.SlideGenerator);
      
      // Use the SlideGenerator if it exists
      if (CellM8.SlideGenerator) {
        try {
          Logger.log('CellM8.SlideGenerator IS DEFINED - Using it');
          Logger.log('Config template selected: ' + config.template);
          Logger.log('Config master template: ' + config.masterTemplate);
          
          const generatorResult = CellM8.SlideGenerator.createPresentation(
            config.title || 'Data Presentation',
            dataResult,
            {
              title: config.title,
              subtitle: config.subtitle,
              slideCount: config.slideCount || 5,
              template: config.template || 'professional',  // Pass template directly
              masterTemplate: config.masterTemplate
            }
          );
          
          Logger.log('Generator returned: ' + JSON.stringify(generatorResult));
          
          if (generatorResult.success) {
            Logger.log('CellM8.SlideGenerator SUCCEEDED - returning result');
            return generatorResult;
          } else {
            Logger.log('Generator returned but success=false');
          }
        } catch (error) {
          Logger.error('CellM8.SlideGenerator threw ERROR: ' + error.toString());
          Logger.error('Stack: ' + error.stack);
        }
      } else {
        Logger.warn('CellM8.SlideGenerator is UNDEFINED');
      }

      // SlideGenerator is required - no fallback
      Logger.error('CRITICAL: CellM8.SlideGenerator is not available!');
      return {
        success: false,
        error: 'Presentation generator not loaded. Please refresh and try again.'
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
   * Extract data from the current sheet or selection
   */
  extractSheetData: function(options) {
    try {
      const sheet = SpreadsheetApp.getActiveSheet();
      let range;
      
      // Check if specific range is requested
      if (options && options.rangeA1) {
        range = sheet.getRange(options.rangeA1);
      } else {
        // Try to get active range first, fall back to data range
        range = sheet.getActiveRange();
        if (!range || range.getNumRows() === 1 && range.getNumColumns() === 1) {
          range = sheet.getDataRange();
        }
      }
      
      const values = range.getValues();
      
      if (values.length === 0) {
        return {
          success: false,
          error: 'No data found in selected range'
        };
      }
      
      // Assume first row is headers
      const headers = values[0];
      const data = values.slice(1);
      
      // Filter out empty rows
      const filteredData = data.filter(row => row.some(cell => cell !== '' && cell !== null));
      
      return {
        success: true,
        headers: headers,
        data: filteredData,
        rowCount: filteredData.length,
        columnCount: headers.length,
        range: range.getA1Notation(),
        sheetName: sheet.getName()
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
   * Analyze data for insights
   */
  analyzeData: function(dataResult) {
    try {
      if (!dataResult || !dataResult.data) {
        return {
          success: false,
          error: 'No data to analyze'
        };
      }
      
      const analysis = {
        totalRows: dataResult.rowCount,
        totalColumns: dataResult.columnCount,
        columnTypes: {},
        statistics: {},
        completeness: 0
      };
      
      // Analyze each column
      dataResult.headers.forEach((header, index) => {
        const columnData = dataResult.data.map(row => row[index]);
        const type = this.detectColumnType(columnData);
        
        analysis.columnTypes[header] = type;
        
        if (type === 'number') {
          analysis.statistics[header] = this.calculateColumnStats(columnData);
        }
      });
      
      // Calculate completeness
      let totalCells = dataResult.rowCount * dataResult.columnCount;
      let filledCells = 0;
      
      dataResult.data.forEach(row => {
        row.forEach(cell => {
          if (cell !== '' && cell !== null) {
            filledCells++;
          }
        });
      });
      
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
      // This function would be implemented with actual slide generation logic
      // For now, it's a placeholder
      return {
        success: true,
        slideCount: config.slideCount || 5
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
   * Apply template to presentation (stub for now)
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
        message: 'Template functionality coming soon'
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
        chartType: chartType,
        dataPoints: data.length
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
      // This would export the presentation to different formats
      return {
        success: true,
        format: format,
        message: 'Export functionality coming soon'
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
      // This would share the presentation
      return {
        success: true,
        sharedWith: emails,
        permission: permission,
        message: 'Share functionality coming soon'
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