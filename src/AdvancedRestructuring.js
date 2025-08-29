/**
 * Advanced Data Restructuring Module
 * Handles complex semi-structured data transformation
 */

const AdvancedRestructuring = {
  
  // Temporary storage for multi-step processing
  _sessionData: {
    originalData: null,
    processedSections: null,
    columnConfig: null,
    previewData: null
  },
  
  /**
   * Analyze the structure of selected data
   */
  analyzeDataStructure: function() {
    try {
      const range = SpreadsheetApp.getActiveRange();
      if (!range) {
        return { success: false, error: 'Please select data to restructure' };
      }
      
      const values = range.getValues();
      if (!values || values.length === 0) {
        return { success: false, error: 'No data found in selection' };
      }
      
      // Flatten and clean the data
      const rawLines = [];
      values.forEach(row => {
        // Join all cells in the row and trim
        const line = row.map(cell => String(cell || '')).join(' ').trim();
        if (line) {
          rawLines.push(line);
        }
      });
      
      if (rawLines.length === 0) {
        return { success: false, error: 'No text data found in selection' };
      }
      
      // Store for later use
      this._sessionData.originalData = rawLines;
      
      // Analyze patterns
      const patterns = this._detectPatterns(rawLines);
      
      // Detect data types
      const dataTypes = this._detectDataTypes(rawLines);
      
      // Count potential sections
      const sectionCount = this._countPotentialSections(rawLines);
      
      return {
        success: true,
        totalLines: rawLines.length,
        detectedSections: sectionCount,
        dataTypes: dataTypes.length,
        patterns: patterns,
        sample: rawLines.slice(0, 10)
      };
      
    } catch (error) {
      Logger.error('Analysis error:', error);
      return { success: false, error: error.toString() };
    }
  },
  
  /**
   * Detect patterns in the data
   */
  _detectPatterns: function(lines) {
    const patterns = [];
    
    // Check for SECTION keyword (like in the example)
    const sectionLines = lines.filter(line => /^SECTION\b/i.test(line));
    if (sectionLines.length > 1) {
      patterns.push({
        type: 'KEYWORD_SECTION',
        keyword: 'SECTION',
        confidence: sectionLines.length / lines.length,
        example: sectionLines[0],
        count: sectionLines.length
      });
    }
    
    // Check for INFO prefix
    const infoLines = lines.filter(line => /^INFO\b/i.test(line));
    if (infoLines.length > 0) {
      patterns.push({
        type: 'INFO_PREFIX',
        confidence: infoLines.length / lines.length,
        example: infoLines[0],
        count: infoLines.length
      });
    }
    
    // Check for all-caps headers
    const capsHeaders = lines.filter(line => /^[A-Z\s]{10,}$/.test(line));
    if (capsHeaders.length > 0) {
      patterns.push({
        type: 'CAPS_HEADER',
        confidence: capsHeaders.length / lines.length,
        example: capsHeaders[0],
        count: capsHeaders.length
      });
    }
    
    // Check for key-value pairs
    const keyValueLines = lines.filter(line => /^\w+[\s\t]+\S+/.test(line));
    if (keyValueLines.length > lines.length * 0.3) {
      patterns.push({
        type: 'KEY_VALUE',
        confidence: keyValueLines.length / lines.length,
        example: keyValueLines[0],
        count: keyValueLines.length
      });
    }
    
    // Check for numbered lists
    const numberedLines = lines.filter(line => /^\d+[\s\t]/.test(line));
    if (numberedLines.length > 5) {
      patterns.push({
        type: 'NUMBERED_LIST',
        confidence: numberedLines.length / lines.length,
        example: numberedLines[0],
        count: numberedLines.length
      });
    }
    
    // Check for table-like data (consistent spacing)
    const spacedLines = lines.filter(line => /\s{2,}/.test(line));
    if (spacedLines.length > lines.length * 0.4) {
      patterns.push({
        type: 'SPACED_COLUMNS',
        confidence: spacedLines.length / lines.length,
        example: spacedLines[0],
        count: spacedLines.length
      });
    }
    
    return patterns.sort((a, b) => b.confidence - a.confidence);
  },
  
  /**
   * Detect data types present in the data
   */
  _detectDataTypes: function(lines) {
    const types = new Set();
    
    lines.forEach(line => {
      // Check for various data types
      if (/\d{1,2}[\/\-:]\d{1,2}[\/\-:]\d{2,4}/.test(line)) types.add('DATE');
      if (/\d{1,2}:\d{2}(:\d{2})?/.test(line)) types.add('TIME');
      if (/[+-]?\d+\.?\d*/.test(line)) types.add('NUMBER');
      if (/\$[\d,]+\.?\d*/.test(line)) types.add('CURRENCY');
      if (/\b[A-Z]{2,}\d+\b/.test(line)) types.add('ID');
      if (/\+?\d{10,}/.test(line)) types.add('PHONE');
      if (/[\w\.-]+@[\w\.-]+\.\w+/.test(line)) types.add('EMAIL');
      if (/\b(?:\d{1,3}\.){3}\d{1,3}\b/.test(line)) types.add('IP_ADDRESS');
      if (/\/dev\/\w+/.test(line)) types.add('DEVICE_PATH');
      if (/\b[A-Z]{2,}(?:ED|ING|ABLE|FUL)\b/.test(line)) types.add('STATUS');
      if (/[-+]?\d+\s*dB[m]?/.test(line)) types.add('SIGNAL_MEASUREMENT');
    });
    
    return Array.from(types);
  },
  
  /**
   * Count potential sections in the data
   */
  _countPotentialSections: function(lines) {
    let sectionCount = 0;
    
    // Count SECTION keywords
    lines.forEach(line => {
      if (/^SECTION\b/i.test(line)) {
        sectionCount++;
      }
    });
    
    // If no SECTION keywords, count blank line separated blocks
    if (sectionCount === 0) {
      let inSection = false;
      lines.forEach(line => {
        if (line.trim() === '') {
          inSection = false;
        } else if (!inSection) {
          sectionCount++;
          inSection = true;
        }
      });
    }
    
    return sectionCount || 1;
  },
  
  /**
   * Process section configuration
   */
  processSectionConfiguration: function(config) {
    try {
      const lines = this._sessionData.originalData;
      if (!lines) {
        return { success: false, error: 'No data to process. Please restart.' };
      }
      
      let sections = [];
      
      switch (config.method) {
        case 'none':
          // Single section with all data
          sections = [{
            name: 'Data',
            lines: lines,
            startLine: 0
          }];
          break;
          
        case 'keyword':
          sections = this._splitByKeyword(lines, config.keyword);
          break;
          
        case 'blankline':
          sections = this._splitByBlankLines(lines);
          break;
          
        case 'pattern':
          sections = this._splitByPattern(lines, config.pattern);
          break;
      }
      
      // Store processed sections
      this._sessionData.processedSections = sections;
      
      return {
        success: true,
        sectionCount: sections.length,
        sections: sections.map(s => ({
          name: s.name,
          lineCount: s.lines.length
        }))
      };
      
    } catch (error) {
      Logger.error('Section processing error:', error);
      return { success: false, error: error.toString() };
    }
  },
  
  /**
   * Split data by keyword
   */
  _splitByKeyword: function(lines, keyword) {
    const sections = [];
    let currentSection = null;
    let sectionIndex = 0;
    
    lines.forEach((line, index) => {
      if (line.toLowerCase().includes(keyword.toLowerCase())) {
        // Save previous section if exists
        if (currentSection && currentSection.lines.length > 0) {
          sections.push(currentSection);
        }
        
        // Start new section
        sectionIndex++;
        
        // Extract section name from the line
        const sectionName = line.replace(new RegExp(keyword, 'i'), '').trim() || `Section ${sectionIndex}`;
        
        currentSection = {
          name: sectionName,
          lines: [],
          startLine: index,
          headerLine: line
        };
      } else if (currentSection) {
        currentSection.lines.push(line);
      } else {
        // Lines before first section
        if (sections.length === 0 && !currentSection) {
          currentSection = {
            name: 'Header',
            lines: [line],
            startLine: 0
          };
        }
      }
    });
    
    // Add last section
    if (currentSection && currentSection.lines.length > 0) {
      sections.push(currentSection);
    }
    
    return sections;
  },
  
  /**
   * Split data by blank lines
   */
  _splitByBlankLines: function(lines) {
    const sections = [];
    let currentSection = {
      name: 'Section 1',
      lines: [],
      startLine: 0
    };
    let sectionIndex = 1;
    
    lines.forEach((line, index) => {
      if (line.trim() === '') {
        // Blank line - end current section if it has content
        if (currentSection.lines.length > 0) {
          sections.push(currentSection);
          sectionIndex++;
          currentSection = {
            name: `Section ${sectionIndex}`,
            lines: [],
            startLine: index + 1
          };
        }
      } else {
        currentSection.lines.push(line);
      }
    });
    
    // Add last section
    if (currentSection.lines.length > 0) {
      sections.push(currentSection);
    }
    
    return sections;
  },
  
  /**
   * Split data by regex pattern
   */
  _splitByPattern: function(lines, pattern) {
    const sections = [];
    let currentSection = null;
    let sectionIndex = 0;
    
    const regex = new RegExp(pattern, 'i');
    
    lines.forEach((line, index) => {
      if (regex.test(line)) {
        // Save previous section if exists
        if (currentSection && currentSection.lines.length > 0) {
          sections.push(currentSection);
        }
        
        // Start new section
        sectionIndex++;
        currentSection = {
          name: line.trim() || `Section ${sectionIndex}`,
          lines: [],
          startLine: index,
          headerLine: line
        };
      } else if (currentSection) {
        currentSection.lines.push(line);
      }
    });
    
    // Add last section
    if (currentSection && currentSection.lines.length > 0) {
      sections.push(currentSection);
    }
    
    return sections;
  },
  
  /**
   * Process column configuration
   */
  processColumnConfiguration: function(config) {
    try {
      const sections = this._sessionData.processedSections;
      if (!sections) {
        return { success: false, error: 'No sections to process. Please restart.' };
      }
      
      // Store column config
      this._sessionData.columnConfig = config;
      
      // Process all sections
      const allRows = [];
      let maxColumns = 0;
      
      sections.forEach(section => {
        // Add section header as a marked row
        if (section.headerLine) {
          const headerRow = this._processLine(section.headerLine, config);
          headerRow._isSection = true;
          headerRow._sectionName = section.name;
          allRows.push(headerRow);
        }
        
        // Process section lines
        section.lines.forEach(line => {
          const processedRow = this._processLine(line, config);
          if (processedRow.length > 0 || config.preserveEmpty) {
            allRows.push(processedRow);
            maxColumns = Math.max(maxColumns, processedRow.length);
          }
        });
      });
      
      // Normalize column count
      allRows.forEach(row => {
        while (row.length < maxColumns) {
          row.push('');
        }
      });
      
      // Store preview data
      this._sessionData.previewData = allRows;
      
      // Suggest headers based on data
      const suggestedHeaders = this._suggestHeaders(allRows, maxColumns);
      
      return {
        success: true,
        preview: {
          rows: allRows,
          columnCount: maxColumns,
          sectionCount: sections.length,
          suggestedHeaders: suggestedHeaders
        }
      };
      
    } catch (error) {
      Logger.error('Column processing error:', error);
      return { success: false, error: error.toString() };
    }
  },
  
  /**
   * Process a single line according to column configuration
   */
  _processLine: function(line, config) {
    if (!line || line.trim() === '') {
      return [];
    }
    
    let parts = [];
    
    switch (config.method) {
      case 'smart':
        parts = this._smartSplit(line);
        break;
        
      case 'whitespace':
        parts = line.split(/\s{2,}/);
        break;
        
      case 'delimiter':
        if (config.delimiter === '\\t') {
          parts = line.split('\t');
        } else {
          parts = line.split(config.delimiter);
        }
        break;
        
      case 'fixed':
        parts = this._fixedWidthSplit(line, config.positions);
        break;
        
      default:
        parts = this._smartSplit(line);
    }
    
    // Apply options
    if (config.trimWhitespace) {
      parts = parts.map(p => p.trim());
    }
    
    if (!config.preserveEmpty) {
      parts = parts.filter(p => p !== '');
    }
    
    if (config.mergeConsecutive && config.method === 'delimiter') {
      // Remove empty parts from consecutive delimiters
      parts = parts.filter(p => p !== '');
    }
    
    return parts;
  },
  
  /**
   * Smart split - intelligently detect column boundaries
   */
  _smartSplit: function(line) {
    const parts = [];
    
    // Special handling for lines with clear patterns
    
    // Pattern 1: Key-value pairs with multiple spaces
    if (/\w+\s{2,}\S+/.test(line)) {
      // Split on multiple spaces
      const splitParts = line.split(/\s{2,}/);
      splitParts.forEach(part => {
        // Further split key-value pairs with single tab or consistent spacing
        if (/^\w+\s+\S+$/.test(part) && !/^\d+\s+\d+/.test(part)) {
          const subParts = part.split(/\s+/);
          if (subParts.length === 2) {
            parts.push(...subParts);
          } else {
            parts.push(part);
          }
        } else {
          parts.push(part);
        }
      });
      return parts;
    }
    
    // Pattern 2: Numbered data (like the signal measurements)
    if (/^\d+\s+[\d:]+\s+/.test(line)) {
      // Split on whitespace but preserve groups
      return line.trim().split(/\s+/);
    }
    
    // Pattern 3: Status or result lines
    if (/^\w+\s+(READY|FAILED|OK|YES|NO|PASSED)\s*$/.test(line)) {
      return line.trim().split(/\s+/);
    }
    
    // Pattern 4: Device paths or technical identifiers
    if (/\/dev\/\w+/.test(line) || /\+[A-Z]+:/.test(line)) {
      // Split on first space to separate label from value
      const firstSpace = line.indexOf(' ');
      if (firstSpace > 0) {
        parts.push(line.substring(0, firstSpace));
        parts.push(line.substring(firstSpace + 1).trim());
      } else {
        parts.push(line);
      }
      return parts;
    }
    
    // Pattern 5: Measurement lines (dBm, dB, etc.)
    if (/[-+]?\d+\.?\d*\s*dB[m]?/.test(line)) {
      // Split keeping measurements together
      const tokens = line.split(/\s+/);
      let current = '';
      
      tokens.forEach(token => {
        if (/^[-+]?\d+\.?\d*$/.test(token)) {
          // Number - might be part of measurement
          if (current) {
            parts.push(current);
            current = token;
          } else {
            current = token;
          }
        } else if (/^dB[m]?$/.test(token)) {
          // Unit - combine with previous number
          current += ' ' + token;
          parts.push(current);
          current = '';
        } else {
          if (current) {
            parts.push(current);
            current = '';
          }
          parts.push(token);
        }
      });
      
      if (current) {
        parts.push(current);
      }
      
      return parts;
    }
    
    // Default: Split on multiple spaces or tabs
    return line.trim().split(/\s{2,}|\t+/);
  },
  
  /**
   * Fixed width split
   */
  _fixedWidthSplit: function(line, positions) {
    const parts = [];
    const posArray = positions.split(',').map(p => parseInt(p.trim()));
    
    let start = 0;
    posArray.forEach(pos => {
      if (pos > start && pos <= line.length) {
        parts.push(line.substring(start, pos));
        start = pos;
      }
    });
    
    // Add remaining
    if (start < line.length) {
      parts.push(line.substring(start));
    }
    
    return parts;
  },
  
  /**
   * Suggest column headers based on data
   */
  _suggestHeaders: function(rows, columnCount) {
    const suggestions = [];
    
    // Analyze first few non-section rows
    const dataRows = rows.filter(r => !r._isSection).slice(0, 10);
    
    for (let col = 0; col < columnCount; col++) {
      const columnData = dataRows.map(row => row[col] || '').filter(v => v !== '');
      
      if (columnData.length === 0) {
        suggestions.push('');
        continue;
      }
      
      // Detect column type
      let suggestion = '';
      
      // Check if all values are numbers
      if (columnData.every(v => /^[-+]?\d+\.?\d*$/.test(v))) {
        suggestion = 'Value';
      }
      // Check if all values are dates
      else if (columnData.every(v => /\d{1,2}[\/\-]\d{1,2}/.test(v))) {
        suggestion = 'Date';
      }
      // Check if all values are times
      else if (columnData.every(v => /\d{1,2}:\d{2}/.test(v))) {
        suggestion = 'Time';
      }
      // Check if all values are status-like
      else if (columnData.every(v => /^(READY|FAILED|OK|YES|NO|PASSED|PENDING)$/i.test(v))) {
        suggestion = 'Status';
      }
      // Check for consistent prefixes
      else if (columnData[0] && /^[A-Z]+[-_]/.test(columnData[0])) {
        suggestion = 'ID';
      }
      // Default to generic
      else {
        // Try to detect from position
        if (col === 0) {
          suggestion = 'Name';
        } else if (col === 1) {
          suggestion = 'Value';
        } else {
          suggestion = `Field ${col + 1}`;
        }
      }
      
      suggestions.push(suggestion);
    }
    
    return suggestions;
  },
  
  /**
   * Apply the restructured data to the spreadsheet
   */
  applyRestructuredData: function(config) {
    try {
      const data = this._sessionData.previewData;
      if (!data || data.length === 0) {
        return { success: false, error: 'No preview data available. Please restart.' };
      }
      
      const range = SpreadsheetApp.getActiveRange();
      const sheet = range.getSheet();
      
      // Prepare final data
      let finalData = [...data];
      
      // Add headers if provided
      if (config.includeHeaders && config.headers.length > 0) {
        // Ensure headers match column count
        const headers = [...config.headers];
        while (headers.length < finalData[0].length) {
          headers.push(`Column ${headers.length + 1}`);
        }
        finalData.unshift(headers);
      }
      
      // Clear original selection
      range.clearContent();
      
      // Calculate new range size
      const numRows = finalData.length;
      const numCols = finalData[0].length;
      
      // Get output range
      const outputRange = sheet.getRange(
        range.getRow(),
        range.getColumn(),
        numRows,
        numCols
      );
      
      // Set values
      outputRange.setValues(finalData);
      
      // Format headers if added
      if (config.includeHeaders) {
        const headerRange = sheet.getRange(range.getRow(), range.getColumn(), 1, numCols);
        headerRange.setFontWeight('bold');
        headerRange.setBackground('#f3f4f6');
        headerRange.setBorder(null, null, true, null, null, null, '#d1d5db', SpreadsheetApp.BorderStyle.SOLID);
      }
      
      // Format section headers
      finalData.forEach((row, index) => {
        if (row._isSection) {
          const rowNum = range.getRow() + index;
          const sectionRange = sheet.getRange(rowNum, range.getColumn(), 1, numCols);
          sectionRange.setBackground('#e0e7ff');
          sectionRange.setFontWeight('bold');
        }
      });
      
      // Auto-resize columns
      for (let col = range.getColumn(); col < range.getColumn() + numCols; col++) {
        sheet.autoResizeColumn(col);
      }
      
      // Clear session data
      this._sessionData = {
        originalData: null,
        processedSections: null,
        columnConfig: null,
        previewData: null
      };
      
      return {
        success: true,
        rowsProcessed: numRows,
        columnsCreated: numCols
      };
      
    } catch (error) {
      Logger.error('Apply error:', error);
      return { success: false, error: error.toString() };
    }
  }
};

// Make functions available globally for HTML interface
function analyzeDataStructure() {
  return AdvancedRestructuring.analyzeDataStructure();
}

function processSectionConfiguration(config) {
  return AdvancedRestructuring.processSectionConfiguration(config);
}

function processColumnConfiguration(config) {
  return AdvancedRestructuring.processColumnConfiguration(config);
}

function applyRestructuredData(config) {
  return AdvancedRestructuring.applyRestructuredData(config);
}