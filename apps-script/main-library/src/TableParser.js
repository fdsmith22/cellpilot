/**
 * TableParser - Intelligently parse unstructured data into columns
 * Handles various delimiters and patterns to structure data
 */

const TableParser = {
  
  /**
   * Analyze data to detect patterns and suggest parsing options
   */
  analyzeData: function(data) {
    if (!data || data.length === 0) {
      return { hasData: false };
    }
    
    const sampleSize = Math.min(10, data.length);
    const samples = data.slice(0, sampleSize);
    
    // Check for common delimiters
    const delimiters = {
      comma: { char: ',', count: 0, consistent: true },
      tab: { char: '\t', count: 0, consistent: true },
      pipe: { char: '|', count: 0, consistent: true },
      semicolon: { char: ';', count: 0, consistent: true },
      space: { char: ' ', count: 0, consistent: true },
      doubleSpace: { char: '  ', count: 0, consistent: true }
    };
    
    // Count delimiter occurrences
    const counts = {};
    samples.forEach((row, index) => {
      const cellValue = String(row[0] || '');
      
      if (!counts[index]) {
        counts[index] = {};
      }
      
      Object.keys(delimiters).forEach(key => {
        const delimiter = delimiters[key];
        try {
          const count = (cellValue.split(delimiter.char).length - 1);
          delimiter.count += count;
          counts[index][key] = count;
          
          // Check consistency with previous row
          if (index > 0 && counts[index - 1][key] !== count) {
            delimiter.consistent = false;
          }
        } catch (e) {
          Logger.error('Error counting delimiter:', key, e);
          delimiter.count = 0;
          counts[index][key] = 0;
        }
      });
    });
    
    // Find best delimiter
    let bestDelimiter = null;
    let maxScore = 0;
    
    Object.keys(delimiters).forEach(key => {
      const delimiter = delimiters[key];
      // Score based on count and consistency
      const score = delimiter.count * (delimiter.consistent ? 2 : 1);
      if (score > maxScore) {
        maxScore = score;
        bestDelimiter = key;
      }
    });
    
    // Detect patterns
    const patterns = this.detectPatterns(samples);
    
    // Analyze first row for potential headers
    const firstRow = String(data[0][0] || '');
    const hasHeaders = this.detectHeaders(firstRow);
    
    return {
      hasData: true,
      rowCount: data.length,
      suggestedDelimiter: bestDelimiter,
      delimiters: delimiters,
      patterns: patterns,
      hasHeaders: hasHeaders,
      firstRowSample: firstRow.substring(0, 100),
      estimatedColumns: bestDelimiter ? Math.round(delimiters[bestDelimiter].count / sampleSize) + 1 : 1
    };
  },
  
  /**
   * Detect common patterns in data
   */
  detectPatterns: function(samples) {
    const patterns = {
      fixedWidth: false,
      keyValue: false,
      jsonLike: false,
      csvLike: false,
      tsvLike: false
    };
    
    samples.forEach(row => {
      const cellValue = String(row[0] || '');
      
      // Check for key-value patterns (e.g., "Name: John")
      if (cellValue.match(/\w+\s*:\s*\w+/)) {
        patterns.keyValue = true;
      }
      
      // Check for JSON-like patterns
      if (cellValue.includes('{') || cellValue.includes('[')) {
        patterns.jsonLike = true;
      }
      
      // Check for CSV patterns
      if (cellValue.includes(',') && !cellValue.includes('\t')) {
        patterns.csvLike = true;
      }
      
      // Check for TSV patterns
      if (cellValue.includes('\t')) {
        patterns.tsvLike = true;
      }
    });
    
    return patterns;
  },
  
  /**
   * Detect if first row contains headers
   */
  detectHeaders: function(firstRow) {
    // Common header keywords
    const headerKeywords = ['name', 'date', 'time', 'email', 'phone', 'address', 
                           'city', 'state', 'country', 'id', 'number', 'amount',
                           'price', 'quantity', 'total', 'status', 'type', 'category',
                           'description', 'title', 'region', 'sales', 'revenue'];
    
    const lowerRow = firstRow.toLowerCase();
    return headerKeywords.some(keyword => lowerRow.includes(keyword));
  },
  
  /**
   * Parse data using specified delimiter
   */
  parseWithDelimiter: function(data, delimiter, options = {}) {
    const parsed = [];
    const hasHeaders = options.hasHeaders !== false;
    const trimCells = options.trimCells !== false;
    
    // Get delimiter character
    const delimiterChar = this.getDelimiterChar(delimiter);
    
    data.forEach((row, index) => {
      const cellValue = String(row[0] || '');
      
      if (!cellValue.trim()) {
        // Skip empty rows
        return;
      }
      
      // Split by delimiter
      let columns = this.smartSplit(cellValue, delimiterChar);
      
      // Trim cells if requested
      if (trimCells) {
        columns = columns.map(col => col.trim());
      }
      
      // Clean quotes if present
      columns = columns.map(col => {
        if ((col.startsWith('"') && col.endsWith('"')) || 
            (col.startsWith("'") && col.endsWith("'"))) {
          return col.slice(1, -1);
        }
        return col;
      });
      
      parsed.push(columns);
    });
    
    return parsed;
  },
  
  /**
   * Smart split that handles quoted values
   */
  smartSplit: function(str, delimiter) {
    const result = [];
    let current = '';
    let inQuotes = false;
    let quoteChar = null;
    
    for (let i = 0; i < str.length; i++) {
      const char = str[i];
      const nextChar = str[i + 1];
      
      // Handle quotes
      if ((char === '"' || char === "'") && !inQuotes) {
        inQuotes = true;
        quoteChar = char;
        current += char;
      } else if (char === quoteChar && inQuotes) {
        // Check for escaped quote
        if (nextChar === quoteChar) {
          current += char + nextChar;
          i++; // Skip next char
        } else {
          inQuotes = false;
          quoteChar = null;
          current += char;
        }
      } else if (char === delimiter && !inQuotes) {
        // Found delimiter outside quotes
        result.push(current);
        current = '';
      } else {
        current += char;
      }
    }
    
    // Add last segment
    if (current) {
      result.push(current);
    }
    
    return result;
  },
  
  /**
   * Parse key-value pairs
   */
  parseKeyValue: function(data, options = {}) {
    const parsed = [];
    const separator = options.separator || ':';
    
    // Collect all keys first
    const allKeys = new Set();
    
    data.forEach(row => {
      const cellValue = String(row[0] || '');
      const pairs = cellValue.split(/[,;|]/).map(p => p.trim());
      
      pairs.forEach(pair => {
        const sepIndex = pair.indexOf(separator);
        if (sepIndex > 0) {
          const key = pair.substring(0, sepIndex).trim();
          allKeys.add(key);
        }
      });
    });
    
    // Create header row
    const headers = Array.from(allKeys);
    if (headers.length > 0) {
      parsed.push(headers);
    }
    
    // Parse each row
    data.forEach(row => {
      const cellValue = String(row[0] || '');
      const rowData = new Array(headers.length).fill('');
      const pairs = cellValue.split(/[,;|]/).map(p => p.trim());
      
      pairs.forEach(pair => {
        const sepIndex = pair.indexOf(separator);
        if (sepIndex > 0) {
          const key = pair.substring(0, sepIndex).trim();
          const value = pair.substring(sepIndex + 1).trim();
          const keyIndex = headers.indexOf(key);
          if (keyIndex >= 0) {
            rowData[keyIndex] = value;
          }
        }
      });
      
      parsed.push(rowData);
    });
    
    return parsed;
  },
  
  /**
   * Parse fixed-width data
   */
  parseFixedWidth: function(data, columnWidths) {
    const parsed = [];
    
    data.forEach(row => {
      const cellValue = String(row[0] || '');
      const columns = [];
      let startPos = 0;
      
      columnWidths.forEach(width => {
        const column = cellValue.substring(startPos, startPos + width).trim();
        columns.push(column);
        startPos += width;
      });
      
      // Add any remaining text as last column
      if (startPos < cellValue.length) {
        columns.push(cellValue.substring(startPos).trim());
      }
      
      parsed.push(columns);
    });
    
    return parsed;
  },
  
  /**
   * Auto-detect fixed width columns
   */
  detectFixedWidth: function(data) {
    if (data.length < 2) return null;
    
    // Analyze first 10 rows to find consistent spaces
    const sampleSize = Math.min(10, data.length);
    const samples = data.slice(0, sampleSize).map(row => String(row[0] || ''));
    
    // Find positions where spaces consistently appear
    const spacePositions = [];
    const maxLength = Math.max(...samples.map(s => s.length));
    
    for (let pos = 0; pos < maxLength; pos++) {
      let allHaveSpace = true;
      
      for (let i = 0; i < samples.length; i++) {
        if (pos >= samples[i].length || samples[i][pos] !== ' ') {
          allHaveSpace = false;
          break;
        }
      }
      
      if (allHaveSpace) {
        spacePositions.push(pos);
      }
    }
    
    // Group consecutive spaces to find column boundaries
    const columnWidths = [];
    let lastBoundary = 0;
    
    for (let i = 0; i < spacePositions.length; i++) {
      if (i === 0 || spacePositions[i] - spacePositions[i-1] > 1) {
        if (spacePositions[i] - lastBoundary > 1) {
          columnWidths.push(spacePositions[i] - lastBoundary);
          lastBoundary = spacePositions[i];
        }
      }
    }
    
    return columnWidths.length > 0 ? columnWidths : null;
  },
  
  /**
   * Get delimiter character from key
   */
  getDelimiterChar: function(delimiterKey) {
    const delimiters = {
      comma: ',',
      tab: '\t',
      pipe: '|',
      semicolon: ';',
      space: ' ',
      doubleSpace: '  ',
      colon: ':',
      dash: '-'
    };
    
    return delimiters[delimiterKey] || ',';
  },
  
  /**
   * Apply parsing to range
   */
  applyToRange: function(range, parsedData) {
    try {
      // Clear the current range
      range.clear();
      
      if (!parsedData || parsedData.length === 0) {
        return { success: false, error: 'No data to apply' };
      }
      
      // Find max columns
      const maxCols = Math.max(...parsedData.map(row => row.length));
      
      // Pad rows to have same number of columns
      const paddedData = parsedData.map(row => {
        while (row.length < maxCols) {
          row.push('');
        }
        return row;
      });
      
      // Get the sheet and starting position
      const sheet = range.getSheet();
      const startRow = range.getRow();
      const startCol = range.getColumn();
      
      // Set the new values
      const targetRange = sheet.getRange(startRow, startCol, paddedData.length, maxCols);
      targetRange.setValues(paddedData);
      
      // Format headers if present
      if (paddedData.length > 0) {
        const headerRange = sheet.getRange(startRow, startCol, 1, maxCols);
        headerRange.setFontWeight('bold');
        headerRange.setBackground('#f3f3f3');
      }
      
      return {
        success: true,
        rowCount: paddedData.length,
        columnCount: maxCols
      };
      
    } catch (error) {
      Logger.error('Error applying parsed data:', error);
      return { success: false, error: error.message };
    }
  }
};