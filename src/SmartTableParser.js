/**
 * SmartTableParser - Advanced parsing for complex data structures
 * Handles variable-length fields like multi-word names
 */

const SmartTableParser = {
  
  /**
   * Analyze patterns across all rows to identify column structure
   */
  analyzePatterns: function(data) {
    if (!data || data.length === 0) return null;
    
    const patterns = [];
    
    // Analyze each row
    data.forEach(row => {
      const cellValue = String(row[0] || '').trim();
      if (!cellValue) return;
      
      // Split into tokens
      const tokens = cellValue.split(/\s+/);
      
      // Analyze each token
      const tokenTypes = tokens.map(token => this.detectTokenType(token));
      patterns.push({
        tokens: tokens,
        types: tokenTypes,
        tokenCount: tokens.length
      });
    });
    
    return patterns;
  },
  
  /**
   * Detect the type of a token
   */
  detectTokenType: function(token) {
    // Remove common punctuation for analysis
    const cleanToken = token.replace(/[,;:.!?]/g, '');
    
    // Check if it's a number (including currency)
    if (this.isNumber(cleanToken)) {
      return 'number';
    }
    
    // Check if it's a currency amount
    if (this.isCurrency(cleanToken)) {
      return 'currency';
    }
    
    // Check if it's a date
    if (this.isDate(cleanToken)) {
      return 'date';
    }
    
    // Check if it's an email
    if (this.isEmail(token)) {
      return 'email';
    }
    
    // Check if it's a known region/location
    if (this.isRegion(cleanToken)) {
      return 'region';
    }
    
    // Check if it's a percentage
    if (this.isPercentage(cleanToken)) {
      return 'percentage';
    }
    
    // Check if it's likely a name part (capitalized)
    if (this.isNameLike(cleanToken)) {
      return 'name';
    }
    
    // Default to text
    return 'text';
  },
  
  /**
   * Smart parse that groups tokens into appropriate columns
   */
  smartParse: function(data, options = {}) {
    const patterns = this.analyzePatterns(data);
    if (!patterns || patterns.length === 0) return [];
    
    // Detect column structure from patterns
    const columnStructure = this.detectColumnStructure(patterns);
    
    // Parse each row according to detected structure
    const parsed = [];
    
    patterns.forEach((pattern, index) => {
      if (!pattern.tokens || pattern.tokens.length === 0) return;
      
      const row = this.parseRowWithStructure(
        pattern.tokens, 
        pattern.types, 
        columnStructure
      );
      
      parsed.push(row);
    });
    
    return parsed;
  },
  
  /**
   * Detect the likely column structure from patterns
   */
  detectColumnStructure: function(patterns) {
    // Find the most common pattern
    const typeSequences = {};
    
    patterns.forEach(pattern => {
      // Look for consistent ending patterns (like "region number")
      if (pattern.types.length >= 2) {
        const lastTwo = pattern.types.slice(-2).join('-');
        typeSequences[lastTwo] = (typeSequences[lastTwo] || 0) + 1;
      }
    });
    
    // Determine structure based on common patterns
    let structure = {
      columns: [],
      nameColumns: 0,
      hasRegion: false,
      hasNumber: false
    };
    
    // Check if most rows end with region-number pattern
    const regionNumberPattern = typeSequences['region-number'] || 
                                typeSequences['region-currency'] || 
                                typeSequences['text-number'] || 0;
    
    if (regionNumberPattern > patterns.length * 0.5) {
      // Common pattern: Name(s) Region Sales
      structure.hasRegion = true;
      structure.hasNumber = true;
      structure.columns = ['name', 'region', 'sales'];
      
      // Determine how many columns are for names
      const avgTokens = patterns.reduce((sum, p) => sum + p.tokenCount, 0) / patterns.length;
      structure.nameColumns = Math.max(1, Math.round(avgTokens - 2));
    } else {
      // Try other patterns
      structure = this.detectAlternativeStructure(patterns);
    }
    
    return structure;
  },
  
  /**
   * Parse a row according to detected structure
   */
  parseRowWithStructure: function(tokens, types, structure) {
    const row = [];
    
    if (structure.hasRegion && structure.hasNumber) {
      // Pattern: Name(s) Region Sales
      // Work backwards from the end
      
      // Last token is likely sales/number
      const sales = tokens[tokens.length - 1] || '';
      
      // Second to last is likely region (if it exists and matches pattern)
      let region = '';
      let nameEndIndex = tokens.length - 1;
      
      if (tokens.length >= 2) {
        const secondLast = tokens[tokens.length - 2];
        if (types[tokens.length - 2] === 'region' || 
            types[tokens.length - 2] === 'text') {
          region = secondLast;
          nameEndIndex = tokens.length - 2;
        }
      }
      
      // Everything before is the name
      const name = tokens.slice(0, nameEndIndex).join(' ');
      
      row.push(name, region, sales);
      
    } else if (structure.columns.length > 0) {
      // Use detected columns
      let tokenIndex = 0;
      
      structure.columns.forEach((colType, colIndex) => {
        if (colType === 'name' && tokenIndex < tokens.length) {
          // For name columns, take multiple tokens if needed
          const nameTokens = [];
          while (tokenIndex < tokens.length && 
                 (types[tokenIndex] === 'name' || types[tokenIndex] === 'text') &&
                 tokenIndex < tokens.length - structure.columns.length + colIndex + 1) {
            nameTokens.push(tokens[tokenIndex]);
            tokenIndex++;
          }
          row.push(nameTokens.join(' '));
        } else if (tokenIndex < tokens.length) {
          row.push(tokens[tokenIndex]);
          tokenIndex++;
        } else {
          row.push('');
        }
      });
      
    } else {
      // Fallback: use all tokens as separate columns
      return tokens;
    }
    
    return row;
  },
  
  /**
   * Detect alternative structure patterns
   */
  detectAlternativeStructure: function(patterns) {
    const structure = {
      columns: [],
      nameColumns: 0,
      hasRegion: false,
      hasNumber: false
    };
    
    // Check for consistent column count
    const columnCounts = {};
    patterns.forEach(p => {
      const count = p.tokenCount;
      columnCounts[count] = (columnCounts[count] || 0) + 1;
    });
    
    // Find most common column count
    let mostCommonCount = 0;
    let maxOccurrences = 0;
    
    Object.keys(columnCounts).forEach(count => {
      if (columnCounts[count] > maxOccurrences) {
        maxOccurrences = columnCounts[count];
        mostCommonCount = parseInt(count);
      }
    });
    
    // Analyze the types in rows with the most common count
    const commonPatterns = patterns.filter(p => p.tokenCount === mostCommonCount);
    
    if (commonPatterns.length > 0) {
      // Check last column type
      const lastTypes = {};
      commonPatterns.forEach(p => {
        const lastType = p.types[p.types.length - 1];
        lastTypes[lastType] = (lastTypes[lastType] || 0) + 1;
      });
      
      // If last column is mostly numbers, it's likely a value column
      if ((lastTypes['number'] || 0) + (lastTypes['currency'] || 0) > commonPatterns.length * 0.5) {
        structure.hasNumber = true;
        
        // Check second to last
        if (mostCommonCount >= 3) {
          const secondLastTypes = {};
          commonPatterns.forEach(p => {
            const type = p.types[p.types.length - 2];
            secondLastTypes[type] = (secondLastTypes[type] || 0) + 1;
          });
          
          if ((secondLastTypes['region'] || 0) + (secondLastTypes['text'] || 0) > commonPatterns.length * 0.3) {
            structure.hasRegion = true;
          }
        }
      }
    }
    
    // Build column structure
    if (structure.hasNumber) {
      if (structure.hasRegion && mostCommonCount >= 3) {
        structure.columns = ['name', 'region', 'value'];
        structure.nameColumns = mostCommonCount - 2;
      } else if (mostCommonCount >= 2) {
        structure.columns = ['name', 'value'];
        structure.nameColumns = mostCommonCount - 1;
      }
    }
    
    return structure;
  },
  
  /**
   * Check if token is a number
   */
  isNumber: function(token) {
    // Remove currency symbols and commas
    const cleaned = token.replace(/[$£€¥,]/g, '');
    return !isNaN(cleaned) && cleaned !== '';
  },
  
  /**
   * Check if token is currency
   */
  isCurrency: function(token) {
    return /^[$£€¥]?[\d,]+\.?\d*$/.test(token) || 
           /^[\d,]+\.?\d*[$£€¥]$/.test(token);
  },
  
  /**
   * Check if token is a date
   */
  isDate: function(token) {
    // Simple date patterns
    return /^\d{1,2}[-/]\d{1,2}[-/]\d{2,4}$/.test(token) ||
           /^\d{4}[-/]\d{1,2}[-/]\d{1,2}$/.test(token);
  },
  
  /**
   * Check if token is an email
   */
  isEmail: function(token) {
    return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(token);
  },
  
  /**
   * Check if token is a known region
   */
  isRegion: function(token) {
    const regions = [
      'north', 'south', 'east', 'west', 'central', 'northeast', 'northwest',
      'southeast', 'southwest', 'midwest', 'region', 'area', 'zone', 'district'
    ];
    
    const lowerToken = token.toLowerCase();
    return regions.some(region => lowerToken.includes(region));
  },
  
  /**
   * Check if token is a percentage
   */
  isPercentage: function(token) {
    return /^\d+\.?\d*%$/.test(token);
  },
  
  /**
   * Check if token looks like a name (capitalized)
   */
  isNameLike: function(token) {
    // Check if first letter is capital and rest are letters
    return /^[A-Z][a-z]+$/.test(token) || 
           /^[A-Z]+$/.test(token) || // All caps (like initials)
           /^[A-Z][a-z]+[A-Z][a-z]+$/.test(token); // Like McDonald
  },
  
  /**
   * Interactive column mapping
   */
  createColumnMapping: function(sampleData, suggestedColumns) {
    // This would be used for user to manually map columns
    return {
      samples: sampleData.slice(0, 5),
      suggestedColumns: suggestedColumns,
      mappings: []
    };
  }
};