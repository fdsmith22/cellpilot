// Data Import/Export Pipeline Manager with ML Intelligence
const DataPipelineManager = {
  // Core properties
  supportedFormats: ['csv', 'tsv', 'json', 'xml', 'html', 'api'],
  pipelines: [],
  transformations: [],
  scheduledImports: [],
  
  // ML properties
  mlEnabled: false,
  importHistory: [],
  dataQualityProfiles: {},
  schemaLearning: {},
  automationRules: {},
  
  /**
   * Initialize Data Pipeline Manager
   */
  initialize: function() {
    try {
      // Check ML status
      const mlStatus = PropertiesService.getUserProperties().getProperty('cellpilot_ml_enabled');
      this.mlEnabled = mlStatus === 'true';
      
      if (this.mlEnabled) {
        // Load import history
        const savedHistory = PropertiesService.getUserProperties().getProperty('import_history');
        if (savedHistory) {
          this.importHistory = JSON.parse(savedHistory);
        }
        
        // Load data quality profiles
        const savedProfiles = PropertiesService.getUserProperties().getProperty('data_quality_profiles');
        if (savedProfiles) {
          this.dataQualityProfiles = JSON.parse(savedProfiles);
        }
        
        // Load schema learning data
        const savedSchemas = PropertiesService.getUserProperties().getProperty('schema_learning');
        if (savedSchemas) {
          this.schemaLearning = JSON.parse(savedSchemas);
        }
      }
      
      return true;
    } catch (error) {
      Logger.error('Error initializing Data Pipeline Manager:', error);
      return false;
    }
  },
  
  /**
   * Import data from various sources
   */
  importData: function(source) {
    try {
      const startTime = Date.now();
      let result = null;
      
      // Determine source type and handle accordingly
      switch (source.type) {
        case 'csv':
          result = this.importCSV(source);
          break;
        case 'json':
          result = this.importJSON(source);
          break;
        case 'api':
          result = this.importFromAPI(source);
          break;
        case 'html_table':
          result = this.importHTMLTable(source);
          break;
        case 'xml':
          result = this.importXML(source);
          break;
        default:
          throw new Error('Unsupported data source type: ' + source.type);
      }
      
      if (result.success) {
        const duration = Date.now() - startTime;
        
        // Apply ML enhancements if enabled
        if (this.mlEnabled) {
          result = this.enhanceImportWithML(result, source, duration);
        }
        
        // Track import for learning
        this.trackImport(source, result, duration);
      }
      
      return result;
      
    } catch (error) {
      Logger.error('Error importing data:', error);
      return {
        success: false,
        error: error.message
      };
    }
  },
  
  /**
   * Import CSV data
   */
  importCSV: function(source) {
    try {
      let csvText = '';
      
      // Handle different CSV sources
      if (source.url) {
        const response = UrlFetchApp.fetch(source.url);
        csvText = response.getContentText();
      } else if (source.data) {
        csvText = source.data;
      } else {
        throw new Error('No CSV data or URL provided');
      }
      
      // Parse CSV
      const delimiter = source.delimiter || this.detectDelimiter(csvText);
      const hasHeaders = source.hasHeaders !== false; // Default to true
      
      const rows = this.parseCSV(csvText, delimiter);
      
      if (rows.length === 0) {
        throw new Error('No data found in CSV');
      }
      
      // Extract headers and data
      let headers = [];
      let data = rows;
      
      if (hasHeaders) {
        headers = rows[0];
        data = rows.slice(1);
      } else {
        // Generate default headers
        headers = Array(rows[0].length).fill(0).map((_, i) => `Column${i + 1}`);
      }
      
      // Clean and validate data
      const cleanedData = this.cleanImportedData(data, headers);
      
      // Write to sheet
      const sheet = this.createOrGetSheet(source.sheetName || 'Imported Data');
      const range = sheet.getRange(1, 1, cleanedData.length + 1, headers.length);
      
      const values = [headers].concat(cleanedData);
      range.setValues(values);
      
      // Apply formatting
      this.formatImportedData(sheet, headers, cleanedData);
      
      return {
        success: true,
        rowsImported: data.length,
        columnsImported: headers.length,
        sheetName: sheet.getName(),
        headers: headers,
        dataQuality: this.assessDataQuality(cleanedData, headers)
      };
      
    } catch (error) {
      Logger.error('Error importing CSV:', error);
      return {
        success: false,
        error: error.message
      };
    }
  },
  
  /**
   * Import JSON data
   */
  importJSON: function(source) {
    try {
      let jsonData = null;
      
      // Handle different JSON sources
      if (source.url) {
        const response = UrlFetchApp.fetch(source.url);
        jsonData = JSON.parse(response.getContentText());
      } else if (source.data) {
        jsonData = typeof source.data === 'string' ? JSON.parse(source.data) : source.data;
      } else {
        throw new Error('No JSON data or URL provided');
      }
      
      // Handle different JSON structures
      let flattenedData = [];
      if (Array.isArray(jsonData)) {
        flattenedData = jsonData.map(item => this.flattenObject(item));
      } else if (typeof jsonData === 'object') {
        // Check if it contains an array
        const arrayKey = this.findArrayInObject(jsonData);
        if (arrayKey) {
          flattenedData = jsonData[arrayKey].map(item => this.flattenObject(item));
        } else {
          // Single object
          flattenedData = [this.flattenObject(jsonData)];
        }
      }
      
      if (flattenedData.length === 0) {
        throw new Error('No usable data found in JSON');
      }
      
      // Extract headers from all objects
      const allHeaders = new Set();
      flattenedData.forEach(row => {
        Object.keys(row).forEach(key => allHeaders.add(key));
      });
      
      const headers = Array.from(allHeaders);
      
      // Convert to 2D array
      const data = flattenedData.map(row => 
        headers.map(header => row[header] || '')
      );
      
      // Clean data
      const cleanedData = this.cleanImportedData(data, headers);
      
      // Write to sheet
      const sheet = this.createOrGetSheet(source.sheetName || 'JSON Import');
      const range = sheet.getRange(1, 1, cleanedData.length + 1, headers.length);
      
      const values = [headers].concat(cleanedData);
      range.setValues(values);
      
      // Apply formatting
      this.formatImportedData(sheet, headers, cleanedData);
      
      return {
        success: true,
        rowsImported: data.length,
        columnsImported: headers.length,
        sheetName: sheet.getName(),
        headers: headers,
        dataQuality: this.assessDataQuality(cleanedData, headers)
      };
      
    } catch (error) {
      Logger.error('Error importing JSON:', error);
      return {
        success: false,
        error: error.message
      };
    }
  },
  
  /**
   * Import data from API
   */
  importFromAPI: function(source) {
    try {
      // Prepare request options
      const options = {
        method: source.method || 'GET',
        headers: source.headers || {},
        muteHttpExceptions: true
      };
      
      if (source.payload) {
        options.payload = typeof source.payload === 'string' 
          ? source.payload 
          : JSON.stringify(source.payload);
        
        if (!options.headers['Content-Type']) {
          options.headers['Content-Type'] = 'application/json';
        }
      }
      
      // Add authentication if provided
      if (source.auth) {
        if (source.auth.type === 'bearer') {
          options.headers['Authorization'] = `Bearer ${source.auth.token}`;
        } else if (source.auth.type === 'basic') {
          options.headers['Authorization'] = `Basic ${Utilities.base64Encode(source.auth.username + ':' + source.auth.password)}`;
        } else if (source.auth.type === 'apikey') {
          options.headers[source.auth.headerName || 'X-API-Key'] = source.auth.key;
        }
      }
      
      // Make API call
      const response = UrlFetchApp.fetch(source.url, options);
      const statusCode = response.getResponseCode();
      
      if (statusCode < 200 || statusCode >= 300) {
        throw new Error(`API request failed with status ${statusCode}: ${response.getContentText()}`);
      }
      
      // Parse response
      const responseText = response.getContentText();
      let responseData;
      
      try {
        responseData = JSON.parse(responseText);
      } catch (e) {
        // Try as CSV if JSON parsing fails
        return this.importCSV({
          data: responseText,
          sheetName: source.sheetName || 'API Import'
        });
      }
      
      // Import as JSON
      return this.importJSON({
        data: responseData,
        sheetName: source.sheetName || 'API Import'
      });
      
    } catch (error) {
      Logger.error('Error importing from API:', error);
      return {
        success: false,
        error: error.message
      };
    }
  },
  
  /**
   * Import HTML table
   */
  importHTMLTable: function(source) {
    try {
      let htmlContent = '';
      
      if (source.url) {
        const response = UrlFetchApp.fetch(source.url);
        htmlContent = response.getContentText();
      } else if (source.data) {
        htmlContent = source.data;
      } else {
        throw new Error('No HTML data or URL provided');
      }
      
      // Extract table data using regex (basic parsing)
      const tableRegex = /<table[^>]*>(.*?)<\/table>/gis;
      const rowRegex = /<tr[^>]*>(.*?)<\/tr>/gis;
      const cellRegex = /<t[dh][^>]*>(.*?)<\/t[dh]>/gis;
      
      const tableMatch = tableRegex.exec(htmlContent);
      if (!tableMatch) {
        throw new Error('No table found in HTML');
      }
      
      const tableContent = tableMatch[1];
      const rows = [];
      let rowMatch;
      
      while ((rowMatch = rowRegex.exec(tableContent)) !== null) {
        const rowContent = rowMatch[1];
        const cells = [];
        let cellMatch;
        
        while ((cellMatch = cellRegex.exec(rowContent)) !== null) {
          // Remove HTML tags and decode entities
          let cellText = cellMatch[1]
            .replace(/<[^>]*>/g, '')
            .replace(/&nbsp;/g, ' ')
            .replace(/&amp;/g, '&')
            .replace(/&lt;/g, '<')
            .replace(/&gt;/g, '>')
            .trim();
          
          cells.push(cellText);
        }
        
        if (cells.length > 0) {
          rows.push(cells);
        }
      }
      
      if (rows.length === 0) {
        throw new Error('No data rows found in table');
      }
      
      // Assume first row is headers
      const headers = rows[0];
      const data = rows.slice(1);
      
      // Clean data
      const cleanedData = this.cleanImportedData(data, headers);
      
      // Write to sheet
      const sheet = this.createOrGetSheet(source.sheetName || 'HTML Table Import');
      const range = sheet.getRange(1, 1, cleanedData.length + 1, headers.length);
      
      const values = [headers].concat(cleanedData);
      range.setValues(values);
      
      // Apply formatting
      this.formatImportedData(sheet, headers, cleanedData);
      
      return {
        success: true,
        rowsImported: data.length,
        columnsImported: headers.length,
        sheetName: sheet.getName(),
        headers: headers,
        dataQuality: this.assessDataQuality(cleanedData, headers)
      };
      
    } catch (error) {
      Logger.error('Error importing HTML table:', error);
      return {
        success: false,
        error: error.message
      };
    }
  },
  
  /**
   * Export data to various formats
   */
  exportData: function(options) {
    try {
      const sheet = options.sheet || SpreadsheetApp.getActiveSheet();
      const range = options.range || sheet.getDataRange();
      
      const data = range.getValues();
      const headers = data[0];
      const rows = data.slice(1);
      
      switch (options.format) {
        case 'csv':
          return this.exportToCSV(headers, rows, options);
        case 'json':
          return this.exportToJSON(headers, rows, options);
        case 'xml':
          return this.exportToXML(headers, rows, options);
        case 'html':
          return this.exportToHTML(headers, rows, options);
        default:
          throw new Error('Unsupported export format: ' + options.format);
      }
      
    } catch (error) {
      Logger.error('Error exporting data:', error);
      return {
        success: false,
        error: error.message
      };
    }
  },
  
  /**
   * Export to CSV
   */
  exportToCSV: function(headers, rows, options) {
    const delimiter = options.delimiter || ',';
    const includeHeaders = options.includeHeaders !== false;
    
    let csvContent = '';
    
    if (includeHeaders) {
      csvContent += this.formatCSVRow(headers, delimiter) + '\n';
    }
    
    rows.forEach(row => {
      csvContent += this.formatCSVRow(row, delimiter) + '\n';
    });
    
    // Create downloadable content
    const blob = Utilities.newBlob(csvContent, 'text/csv', options.filename || 'export.csv');
    
    return {
      success: true,
      content: csvContent,
      blob: blob,
      rowsExported: rows.length,
      format: 'csv'
    };
  },
  
  /**
   * Export to JSON
   */
  exportToJSON: function(headers, rows, options) {
    const jsonArray = rows.map(row => {
      const obj = {};
      headers.forEach((header, index) => {
        obj[header] = row[index];
      });
      return obj;
    });
    
    const jsonContent = JSON.stringify(jsonArray, null, options.indent ? 2 : 0);
    const blob = Utilities.newBlob(jsonContent, 'application/json', options.filename || 'export.json');
    
    return {
      success: true,
      content: jsonContent,
      blob: blob,
      rowsExported: rows.length,
      format: 'json'
    };
  },
  
  /**
   * Detect CSV delimiter
   */
  detectDelimiter: function(csvText) {
    const delimiters = [',', ';', '\t', '|'];
    const sample = csvText.split('\n').slice(0, 5).join('\n'); // Check first 5 lines
    
    let bestDelimiter = ',';
    let maxCount = 0;
    
    delimiters.forEach(delimiter => {
      const matches = (sample.match(new RegExp('\\' + delimiter, 'g')) || []).length;
      if (matches > maxCount) {
        maxCount = matches;
        bestDelimiter = delimiter;
      }
    });
    
    return bestDelimiter;
  },
  
  /**
   * Parse CSV content
   */
  parseCSV: function(csvText, delimiter) {
    const rows = [];
    const lines = csvText.split('\n');
    
    lines.forEach(line => {
      if (line.trim()) {
        const row = this.parseCSVRow(line, delimiter);
        rows.push(row);
      }
    });
    
    return rows;
  },
  
  /**
   * Parse single CSV row handling quotes
   */
  parseCSVRow: function(line, delimiter) {
    const row = [];
    let current = '';
    let inQuotes = false;
    
    for (let i = 0; i < line.length; i++) {
      const char = line[i];
      const nextChar = line[i + 1];
      
      if (char === '"') {
        if (inQuotes && nextChar === '"') {
          // Escaped quote
          current += '"';
          i++; // Skip next quote
        } else {
          // Toggle quote state
          inQuotes = !inQuotes;
        }
      } else if (char === delimiter && !inQuotes) {
        // End of field
        row.push(current.trim());
        current = '';
      } else {
        current += char;
      }
    }
    
    // Add last field
    row.push(current.trim());
    
    return row;
  },
  
  /**
   * Format CSV row with proper escaping
   */
  formatCSVRow: function(row, delimiter) {
    return row.map(cell => {
      const cellStr = String(cell);
      
      // Escape if contains delimiter, quotes, or newlines
      if (cellStr.includes(delimiter) || cellStr.includes('"') || cellStr.includes('\n')) {
        return '"' + cellStr.replace(/"/g, '""') + '"';
      }
      
      return cellStr;
    }).join(delimiter);
  },
  
  /**
   * Flatten nested object for JSON import
   */
  flattenObject: function(obj, prefix = '') {
    const flattened = {};
    
    for (const key in obj) {
      if (obj.hasOwnProperty(key)) {
        const newKey = prefix ? `${prefix}.${key}` : key;
        
        if (obj[key] !== null && typeof obj[key] === 'object' && !Array.isArray(obj[key])) {
          // Recursively flatten nested objects
          Object.assign(flattened, this.flattenObject(obj[key], newKey));
        } else if (Array.isArray(obj[key])) {
          // Convert arrays to strings
          flattened[newKey] = obj[key].join(', ');
        } else {
          flattened[newKey] = obj[key];
        }
      }
    }
    
    return flattened;
  },
  
  /**
   * Find array in object for JSON parsing
   */
  findArrayInObject: function(obj) {
    for (const key in obj) {
      if (Array.isArray(obj[key])) {
        return key;
      }
    }
    return null;
  },
  
  /**
   * Clean imported data
   */
  cleanImportedData: function(data, headers) {
    return data.map(row => {
      return row.map((cell, index) => {
        let cleaned = cell;
        
        // Handle null/undefined
        if (cleaned === null || cleaned === undefined) {
          cleaned = '';
        }
        
        // Convert to string and trim
        cleaned = String(cleaned).trim();
        
        // Auto-detect and convert data types
        const header = headers[index];
        if (this.mlEnabled && this.schemaLearning[header]) {
          const expectedType = this.schemaLearning[header].type;
          cleaned = this.convertToType(cleaned, expectedType);
        } else {
          cleaned = this.autoConvertType(cleaned);
        }
        
        return cleaned;
      });
    });
  },
  
  /**
   * Auto-convert data type
   */
  autoConvertType: function(value) {
    const str = String(value).trim();
    
    // Empty values
    if (str === '') return '';
    
    // Numbers
    if (/^-?\d+(\.\d+)?$/.test(str)) {
      return parseFloat(str);
    }
    
    // Dates
    const date = new Date(str);
    if (!isNaN(date.getTime()) && str.length > 6) {
      return date;
    }
    
    // Booleans
    if (/^(true|false|yes|no|y|n)$/i.test(str)) {
      return /^(true|yes|y)$/i.test(str);
    }
    
    return str;
  },
  
  /**
   * Convert to specific type
   */
  convertToType: function(value, type) {
    const str = String(value).trim();
    
    switch (type) {
      case 'number':
        const num = parseFloat(str);
        return isNaN(num) ? str : num;
      case 'date':
        const date = new Date(str);
        return isNaN(date.getTime()) ? str : date;
      case 'boolean':
        return /^(true|yes|y|1)$/i.test(str);
      default:
        return str;
    }
  },
  
  /**
   * Create or get sheet
   */
  createOrGetSheet: function(sheetName) {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      sheet = spreadsheet.insertSheet(sheetName);
    } else {
      // Clear existing content
      sheet.clear();
    }
    
    return sheet;
  },
  
  /**
   * Format imported data
   */
  formatImportedData: function(sheet, headers, data) {
    try {
      // Format header row
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#f0f0f0');
      
      // Auto-resize columns
      for (let i = 1; i <= headers.length; i++) {
        sheet.autoResizeColumn(i);
      }
      
      // Apply data type formatting
      for (let colIndex = 0; colIndex < headers.length; colIndex++) {
        const columnRange = sheet.getRange(2, colIndex + 1, data.length, 1);
        
        // Detect column type from sample data
        const sampleData = data.slice(0, Math.min(100, data.length)).map(row => row[colIndex]);
        const columnType = this.detectColumnDataType(sampleData);
        
        // Apply formatting based on type
        switch (columnType) {
          case 'number':
            columnRange.setNumberFormat('#,##0.00');
            break;
          case 'currency':
            columnRange.setNumberFormat('$#,##0.00');
            break;
          case 'percentage':
            columnRange.setNumberFormat('0.00%');
            break;
          case 'date':
            columnRange.setNumberFormat('mm/dd/yyyy');
            break;
        }
      }
      
      // Freeze header row
      sheet.setFrozenRows(1);
      
    } catch (error) {
      Logger.error('Error formatting imported data:', error);
    }
  },
  
  /**
   * Detect column data type
   */
  detectColumnDataType: function(sampleData) {
    const nonEmptyData = sampleData.filter(val => val !== '' && val !== null);
    
    if (nonEmptyData.length === 0) return 'text';
    
    // Check for currency
    const currencyCount = nonEmptyData.filter(val => 
      /^\$[\d,]+\.?\d*$/.test(String(val))
    ).length;
    
    if (currencyCount / nonEmptyData.length > 0.7) return 'currency';
    
    // Check for percentage
    const percentageCount = nonEmptyData.filter(val => 
      /^\d+\.?\d*%$/.test(String(val))
    ).length;
    
    if (percentageCount / nonEmptyData.length > 0.7) return 'percentage';
    
    // Check for numbers
    const numberCount = nonEmptyData.filter(val => 
      !isNaN(val) && !isNaN(parseFloat(val))
    ).length;
    
    if (numberCount / nonEmptyData.length > 0.7) return 'number';
    
    // Check for dates
    const dateCount = nonEmptyData.filter(val => {
      const date = new Date(val);
      return !isNaN(date.getTime());
    }).length;
    
    if (dateCount / nonEmptyData.length > 0.7) return 'date';
    
    return 'text';
  },
  
  /**
   * Assess data quality
   */
  assessDataQuality: function(data, headers) {
    const quality = {
      completeness: 0,
      consistency: 0,
      validity: 0,
      issues: []
    };
    
    const totalCells = data.length * headers.length;
    let emptyCount = 0;
    let inconsistentCount = 0;
    let invalidCount = 0;
    
    // Check each column
    headers.forEach((header, colIndex) => {
      const columnData = data.map(row => row[colIndex]);
      const nonEmptyData = columnData.filter(val => val !== '' && val !== null);
      
      // Completeness
      const emptyInColumn = columnData.length - nonEmptyData.length;
      emptyCount += emptyInColumn;
      
      if (emptyInColumn > columnData.length * 0.2) {
        quality.issues.push(`Column "${header}" has ${emptyInColumn} empty values (${(emptyInColumn/columnData.length*100).toFixed(1)}%)`);
      }
      
      // Consistency (data type consistency)
      if (nonEmptyData.length > 0) {
        const expectedType = this.detectColumnDataType(nonEmptyData);
        const inconsistent = nonEmptyData.filter(val => {
          const actualType = this.detectColumnDataType([val]);
          return actualType !== expectedType && actualType !== 'text';
        }).length;
        
        inconsistentCount += inconsistent;
        
        if (inconsistent > nonEmptyData.length * 0.1) {
          quality.issues.push(`Column "${header}" has inconsistent data types`);
        }
      }
    });
    
    // Calculate scores
    quality.completeness = Math.max(0, (totalCells - emptyCount) / totalCells * 100);
    quality.consistency = Math.max(0, (totalCells - inconsistentCount) / totalCells * 100);
    quality.validity = Math.max(0, (totalCells - invalidCount) / totalCells * 100);
    
    return quality;
  },
  
  /**
   * Enhance import with ML
   */
  enhanceImportWithML: function(result, source, duration) {
    try {
      // Learn column types for future imports
      if (result.headers) {
        result.headers.forEach((header, index) => {
          if (!this.schemaLearning[header]) {
            this.schemaLearning[header] = {
              type: 'text',
              frequency: 0,
              sources: []
            };
          }
          
          this.schemaLearning[header].frequency++;
          this.schemaLearning[header].sources.push(source.type);
        });
      }
      
      // Update data quality profile
      if (result.dataQuality && source.url) {
        const sourceKey = source.url;
        if (!this.dataQualityProfiles[sourceKey]) {
          this.dataQualityProfiles[sourceKey] = {
            history: [],
            avgQuality: 0
          };
        }
        
        this.dataQualityProfiles[sourceKey].history.push({
          timestamp: new Date().toISOString(),
          quality: result.dataQuality,
          rowCount: result.rowsImported
        });
        
        // Keep recent history
        if (this.dataQualityProfiles[sourceKey].history.length > 10) {
          this.dataQualityProfiles[sourceKey].history = 
            this.dataQualityProfiles[sourceKey].history.slice(-10);
        }
        
        // Calculate average quality
        const avgCompleteness = this.dataQualityProfiles[sourceKey].history
          .reduce((sum, h) => sum + h.quality.completeness, 0) / 
          this.dataQualityProfiles[sourceKey].history.length;
        
        this.dataQualityProfiles[sourceKey].avgQuality = avgCompleteness;
      }
      
      // Add ML insights to result
      result.mlInsights = {
        dataQualityTrend: this.getDataQualityTrend(source.url),
        schemaConfidence: this.getSchemaConfidence(result.headers),
        recommendedTransformations: this.getRecommendedTransformations(result)
      };
      
      return result;
      
    } catch (error) {
      Logger.error('Error enhancing import with ML:', error);
      return result;
    }
  },
  
  /**
   * Track import for learning
   */
  trackImport: function(source, result, duration) {
    try {
      this.importHistory.push({
        timestamp: new Date().toISOString(),
        source: {
          type: source.type,
          url: source.url || 'manual',
          size: result.rowsImported * result.columnsImported
        },
        result: {
          success: result.success,
          rowsImported: result.rowsImported,
          columnsImported: result.columnsImported,
          dataQuality: result.dataQuality
        },
        performance: {
          duration: duration,
          throughput: (result.rowsImported * result.columnsImported) / (duration / 1000)
        }
      });
      
      // Keep recent history
      if (this.importHistory.length > 1000) {
        this.importHistory = this.importHistory.slice(-1000);
      }
      
      // Save to properties
      this.saveMLData();
      
    } catch (error) {
      Logger.error('Error tracking import:', error);
    }
  },
  
  /**
   * Get data quality trend
   */
  getDataQualityTrend: function(sourceUrl) {
    if (!sourceUrl || !this.dataQualityProfiles[sourceUrl]) {
      return 'unknown';
    }
    
    const history = this.dataQualityProfiles[sourceUrl].history;
    if (history.length < 3) return 'insufficient_data';
    
    const recent = history.slice(-3);
    const firstQuality = recent[0].quality.completeness;
    const lastQuality = recent[recent.length - 1].quality.completeness;
    
    if (lastQuality > firstQuality + 5) return 'improving';
    if (lastQuality < firstQuality - 5) return 'declining';
    return 'stable';
  },
  
  /**
   * Get schema confidence
   */
  getSchemaConfidence: function(headers) {
    if (!headers) return 0;
    
    let totalConfidence = 0;
    headers.forEach(header => {
      if (this.schemaLearning[header]) {
        totalConfidence += Math.min(1, this.schemaLearning[header].frequency / 10);
      }
    });
    
    return totalConfidence / headers.length;
  },
  
  /**
   * Get recommended transformations
   */
  getRecommendedTransformations: function(result) {
    const recommendations = [];
    
    if (result.dataQuality) {
      if (result.dataQuality.completeness < 80) {
        recommendations.push({
          type: 'clean_empty_values',
          reason: 'High number of empty values detected',
          priority: 'high'
        });
      }
      
      if (result.dataQuality.consistency < 70) {
        recommendations.push({
          type: 'standardize_data_types',
          reason: 'Inconsistent data types found',
          priority: 'medium'
        });
      }
      
      if (result.dataQuality.issues.length > 5) {
        recommendations.push({
          type: 'data_profiling',
          reason: 'Multiple data quality issues detected',
          priority: 'high'
        });
      }
    }
    
    return recommendations;
  },
  
  /**
   * Save ML data
   */
  saveMLData: function() {
    try {
      PropertiesService.getUserProperties().setProperty(
        'import_history',
        JSON.stringify(this.importHistory)
      );
      
      PropertiesService.getUserProperties().setProperty(
        'data_quality_profiles',
        JSON.stringify(this.dataQualityProfiles)
      );
      
      PropertiesService.getUserProperties().setProperty(
        'schema_learning',
        JSON.stringify(this.schemaLearning)
      );
      
    } catch (error) {
      Logger.error('Error saving ML data:', error);
    }
  },
  
  /**
   * Get pipeline statistics
   */
  getPipelineStats: function() {
    return {
      totalImports: this.importHistory.length,
      supportedFormats: this.supportedFormats,
      averageDataQuality: this.calculateAverageDataQuality(),
      schemaKnowledge: Object.keys(this.schemaLearning).length,
      qualityProfiles: Object.keys(this.dataQualityProfiles).length,
      mlEnabled: this.mlEnabled
    };
  },
  
  /**
   * Calculate average data quality
   */
  calculateAverageDataQuality: function() {
    const qualityScores = this.importHistory
      .filter(h => h.result.dataQuality)
      .map(h => h.result.dataQuality.completeness);
    
    if (qualityScores.length === 0) return 0;
    
    return qualityScores.reduce((sum, score) => sum + score, 0) / qualityScores.length;
  }
};