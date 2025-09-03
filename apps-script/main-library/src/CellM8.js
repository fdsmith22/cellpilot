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
          sheetName: sheet.getName(),
          hasSelection: true
        };
      }
      
      // If no selection, return info about the entire sheet
      const dataRange = sheet.getDataRange();
      return {
        success: true,
        range: dataRange.getA1Notation(),
        rows: dataRange.getNumRows(),
        cols: dataRange.getNumColumns(),
        sheetName: sheet.getName(),
        hasSelection: false,
        message: 'No range selected - using entire sheet'
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
      Logger.error('Error selecting entire range:', error);
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
   * Get available sheets
   */
  getAvailableSheets: function() {
    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const sheets = spreadsheet.getSheets();
      
      const sheetInfo = sheets.map(sheet => ({
        name: sheet.getName(),
        rows: sheet.getMaxRows(),
        cols: sheet.getMaxColumns(),
        dataRows: sheet.getDataRange().getNumRows(),
        dataCols: sheet.getDataRange().getNumColumns()
      }));
      
      return {
        success: true,
        sheets: sheetInfo,
        activeSheet: spreadsheet.getActiveSheet().getName()
      };
    } catch (error) {
      Logger.error('Error getting sheets:', error);
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
      
      // Try CellM8SlideGenerator FIRST (creates its own presentation)
      // All templates now use the enhanced generator
      if (typeof CellM8SlideGenerator !== 'undefined') {
        try {
          Logger.log('Using CellM8 slide generator with research-based approach');
          Logger.log('Config template selected: ' + config.template);
          Logger.log('Config master template: ' + config.masterTemplate);
          
          const generatorResult = CellM8SlideGenerator.createPresentation(
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
          
          if (generatorResult.success) {
            Logger.log('CellM8 slide generator succeeded');
            return generatorResult;
          }
        } catch (error) {
          Logger.warn('CellM8 slide generator failed, falling back:', error);
        }
      }

      // Fall back to simple presentation if CellM8SlideGenerator is not available or simple template selected
      // This creates a basic presentation using the legacy approach
      Logger.log('Falling back to simple presentation creation');
      
      // Create presentation using Slides API
      const presentation = SlidesApp.create(config.title);
      const presentationId = presentation.getId();
      
      // Get the presentation for editing
      const pres = SlidesApp.openById(presentationId);
      
      // Add title slide
      const slides = pres.getSlides();
      const titleSlide = slides[0];
      
      // Set title and subtitle using page elements
      const pageElements = titleSlide.getPageElements();
      
      // Find title and subtitle shapes
      for (let i = 0; i < pageElements.length; i++) {
        const element = pageElements[i];
        const type = element.getPageElementType();
        
        if (type === SlidesApp.PageElementType.SHAPE) {
          const shape = element.asShape();
          try {
            const placeholder = shape.getPlaceholderType();
            
            if (placeholder === SlidesApp.PlaceholderType.TITLE || 
                placeholder === SlidesApp.PlaceholderType.CENTERED_TITLE) {
              shape.getText().setText(config.title);
            } else if (placeholder === SlidesApp.PlaceholderType.SUBTITLE || 
                       placeholder === SlidesApp.PlaceholderType.BODY) {
              const subtitle = config.subtitle || `Data Analysis - ${dataResult.rowCount} rows × ${dataResult.columnCount} columns`;
              shape.getText().setText(subtitle);
            }
          } catch (e) {
            // Not a placeholder, skip
          }
        }
      }
      
      // Add overview slide
      const overviewSlide = pres.appendSlide(SlidesApp.PredefinedLayout.TITLE_AND_BODY);
      
      // Get page elements and find placeholders
      const overviewElements = overviewSlide.getPageElements();
      for (let i = 0; i < overviewElements.length; i++) {
        const element = overviewElements[i];
        if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
          const shape = element.asShape();
          try {
            const placeholder = shape.getPlaceholderType();
            
            if (placeholder === SlidesApp.PlaceholderType.TITLE || 
                placeholder === SlidesApp.PlaceholderType.CENTERED_TITLE) {
              shape.getText().setText('Data Overview');
            } else if (placeholder === SlidesApp.PlaceholderType.BODY) {
              let overviewText = `Dataset Information:\n`;
              overviewText += `• Total Rows: ${dataResult.rowCount}\n`;
              overviewText += `• Total Columns: ${dataResult.columnCount}\n`;
              overviewText += `• Headers: ${dataResult.headers.slice(0, 5).join(', ')}`;
              if (dataResult.headers.length > 5) {
                overviewText += ` (and ${dataResult.headers.length - 5} more)`;
              }
              shape.getText().setText(overviewText);
            }
          } catch (e) {
            // Not a placeholder
          }
        }
      }
      
      // Add data table slide (if data exists)
      if (dataResult.data && dataResult.data.length > 0) {
        const tableSlide = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
        
        // Add title with proper positioning
        const titleElement = tableSlide.insertTextBox('Sample Data', 30, 20, 400, 40);
        titleElement.getText().getTextStyle().setFontSize(20).setBold(true);
        
        // Create smaller, properly sized table
        const numRows = Math.min(5, dataResult.data.length + 1); // +1 for headers, max 5 rows total
        const numCols = Math.min(4, dataResult.headers.length); // Max 4 columns for readability
        
        // Smaller table that fits on slide
        const table = tableSlide.insertTable(numRows, numCols, 30, 70, 400, 200);
        
        // Add headers with smaller font
        for (let col = 0; col < numCols; col++) {
          const cell = table.getCell(0, col);
          const headerText = String(dataResult.headers[col] || '').substring(0, 15); // Limit header length
          cell.getText().setText(headerText);
          cell.getFill().setSolidFill('#e5e7eb');
          cell.getText().getTextStyle().setBold(true).setFontSize(10);
        }
        
        // Add data rows with truncated text
        for (let row = 0; row < numRows - 1; row++) {
          for (let col = 0; col < numCols; col++) {
            const cell = table.getCell(row + 1, col);
            const value = dataResult.data[row] && dataResult.data[row][col];
            const cellText = String(value || '').substring(0, 20); // Limit cell text length
            cell.getText().setText(cellText);
            cell.getText().getTextStyle().setFontSize(9);
          }
        }
      }
      
      // Calculate how many more slides we need
      const currentSlideCount = pres.getSlides().length;
      const targetSlideCount = config.slideCount || 5;
      const additionalSlidesNeeded = targetSlideCount - currentSlideCount;
      
      // Add content slides if provided
      if (config.content && config.content.length > 0) {
        for (let i = 0; i < Math.min(config.content.length, additionalSlidesNeeded); i++) {
          const slideData = config.content[i];
          const layout = slideData.type === 'table' ? 
            SlidesApp.PredefinedLayout.BLANK : 
            SlidesApp.PredefinedLayout.TITLE_AND_BODY;
          
          const slide = pres.appendSlide(layout);
          
          if (slideData.type === 'table' && slideData.headers && slideData.rows) {
            // Create table slide
            const titleElement = slide.insertTextBox(slideData.title || 'Data Table', 20, 20, 600, 50);
            titleElement.getText().getTextStyle().setFontSize(24).setBold(true);
            
            const table = slide.insertTable(
              Math.min(slideData.rows.length + 1, 10),
              Math.min(slideData.headers.length, 5),
              20, 80, 680, 300
            );
            
            // Add headers
            for (let col = 0; col < Math.min(slideData.headers.length, 5); col++) {
              const cell = table.getCell(0, col);
              cell.getText().setText(String(slideData.headers[col]));
              cell.getFill().setSolidFill('#f0f0f0');
              cell.getText().getTextStyle().setBold(true);
            }
            
            // Add rows
            for (let row = 0; row < Math.min(slideData.rows.length, 9); row++) {
              for (let col = 0; col < Math.min(slideData.headers.length, 5); col++) {
                const cell = table.getCell(row + 1, col);
                cell.getText().setText(String(slideData.rows[row][col] || ''));
              }
            }
          } else {
            // Regular slide - find placeholders safely
            const slideElements = slide.getPageElements();
            for (let j = 0; j < slideElements.length; j++) {
              const element = slideElements[j];
              if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
                const shape = element.asShape();
                try {
                  const placeholder = shape.getPlaceholderType();
                  
                  if (placeholder === SlidesApp.PlaceholderType.TITLE || 
                      placeholder === SlidesApp.PlaceholderType.CENTERED_TITLE) {
                    shape.getText().setText(slideData.title || `Slide ${i + 3}`);
                  } else if (placeholder === SlidesApp.PlaceholderType.BODY) {
                    if (slideData.bullets && slideData.bullets.length > 0) {
                      shape.getText().setText(slideData.bullets.join('\n• '));
                    } else {
                      shape.getText().setText(slideData.content || slideData.subtitle || '');
                    }
                  }
                } catch (e) {
                  // Not a placeholder
                }
              }
            }
          }
        }
      }
      
      // Add additional slides if we haven't reached the target count
      const finalSlideCount = pres.getSlides().length;
      if (finalSlideCount < targetSlideCount) {
        const remainingSlides = targetSlideCount - finalSlideCount;
        
        // Add insight slides based on data analysis
        for (let i = 0; i < remainingSlides; i++) {
          const insightSlide = pres.appendSlide(SlidesApp.PredefinedLayout.TITLE_AND_BODY);
          
          // Find placeholders safely
          const slideElements = insightSlide.getPageElements();
          for (let j = 0; j < slideElements.length; j++) {
            const element = slideElements[j];
            if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
              const shape = element.asShape();
              try {
                const placeholder = shape.getPlaceholderType();
                
                if (placeholder === SlidesApp.PlaceholderType.TITLE || 
                    placeholder === SlidesApp.PlaceholderType.CENTERED_TITLE) {
                  const titles = ['Key Insights', 'Data Summary', 'Analysis Results', 'Next Steps', 'Recommendations'];
                  shape.getText().setText(titles[i % titles.length]);
                } else if (placeholder === SlidesApp.PlaceholderType.BODY) {
                  const content = this.generateInsightContent(dataResult, i);
                  shape.getText().setText(content);
                }
              } catch (e) {
                // Not a placeholder
              }
            }
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
   * Generate insight content for additional slides
   */
  generateInsightContent: function(dataResult, slideIndex) {
    const insights = [
      `• Dataset contains ${dataResult.rowCount} records\n• ${dataResult.columnCount} data fields analyzed\n• Data completeness: ${Math.round((dataResult.rowCount * dataResult.columnCount - this.countEmptyCells(dataResult)) / (dataResult.rowCount * dataResult.columnCount) * 100)}%`,
      `• Primary columns: ${dataResult.headers.slice(0, 3).join(', ')}\n• Total data points: ${dataResult.rowCount * dataResult.columnCount}\n• Sheet name: ${dataResult.sheetName || 'Active Sheet'}`,
      `• Data range: ${dataResult.range || 'Full dataset'}\n• Non-empty rows: ${dataResult.rowCount}\n• Analysis complete`,
      `• Review data quality\n• Identify patterns and trends\n• Create visualizations\n• Share insights with stakeholders`,
      `• Consider adding charts for visual impact\n• Group related data together\n• Focus on key metrics\n• Update regularly for accuracy`
    ];
    
    return insights[slideIndex % insights.length];
  },
  
  /**
   * Count empty cells in data
   */
  countEmptyCells: function(dataResult) {
    let emptyCount = 0;
    if (dataResult.data) {
      for (let row of dataResult.data) {
        for (let cell of row) {
          if (cell === '' || cell === null || cell === undefined) {
            emptyCount++;
          }
        }
      }
    }
    return emptyCount;
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