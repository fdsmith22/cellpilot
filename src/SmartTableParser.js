/**
 * SmartTableParser - Advanced parsing for complex data structures
 * Handles variable-length fields like multi-word names
 */

const SmartTableParser = {
  // ML-Enhanced Properties
  mlEngine: null,
  mlEnabled: false,
  columnTypeConfidence: {},
  mlPredictions: {},
  userTypeOverrides: {},
  learningHistory: [],
  
  /**
   * Initialize ML engine for column detection
   */
  initializeML: async function() {
    try {
      if (typeof window !== 'undefined' && window.CellPilotMLEngine) {
        this.mlEngine = new window.CellPilotMLEngine();
        await this.mlEngine.initialize();
        this.mlEnabled = true;
        console.log('SmartTableParser ML initialized successfully');
      }
    } catch (error) {
      console.warn('ML initialization failed, falling back to rule-based parsing:', error);
      this.mlEnabled = false;
    }
  },
  
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
   * ML-Enhanced column type detection
   */
  detectColumnTypesML: async function(data, headers = []) {
    if (!this.mlEnabled || !this.mlEngine || !data || data.length === 0) {
      return this.detectColumnTypesFallback(data, headers);
    }
    
    try {
      // Use ML engine for column classification
      const mlResults = await this.mlEngine.classifyColumns(data, headers);
      
      // Store confidence scores
      this.columnTypeConfidence = {};
      mlResults.predictions.forEach((pred, idx) => {
        this.columnTypeConfidence[idx] = pred.confidence;
      });
      
      // Apply user overrides if any
      const finalTypes = mlResults.predictions.map((pred, idx) => {
        if (this.userTypeOverrides[idx]) {
          return this.userTypeOverrides[idx];
        }
        return pred.type;
      });
      
      // Store ML predictions for learning
      this.mlPredictions = {
        types: finalTypes,
        confidences: mlResults.predictions.map(p => p.confidence),
        timestamp: Date.now()
      };
      
      return finalTypes;
    } catch (error) {
      console.warn('ML column detection failed, using fallback:', error);
      return this.detectColumnTypesFallback(data, headers);
    }
  },
  
  /**
   * Fallback column type detection (original rule-based)
   */
  detectColumnTypesFallback: function(data, headers) {
    const columnTypes = [];
    const sampleSize = Math.min(10, data.length);
    
    // Analyze each column
    const maxCols = Math.max(...data.map(row => row.length));
    
    for (let col = 0; col < maxCols; col++) {
      const columnSamples = [];
      for (let row = 0; row < sampleSize; row++) {
        if (data[row] && data[row][col]) {
          columnSamples.push(String(data[row][col]));
        }
      }
      
      // Detect type based on samples
      const type = this.detectColumnTypeFromSamples(columnSamples, headers[col]);
      columnTypes.push(type);
    }
    
    return columnTypes;
  },
  
  /**
   * Detect column type from sample values
   */
  detectColumnTypeFromSamples: function(samples, header = '') {
    if (samples.length === 0) return 'text';
    
    const types = samples.map(sample => this.detectTokenType(sample));
    const typeCount = {};
    
    types.forEach(type => {
      typeCount[type] = (typeCount[type] || 0) + 1;
    });
    
    // Find dominant type
    let dominantType = 'text';
    let maxCount = 0;
    
    Object.keys(typeCount).forEach(type => {
      if (typeCount[type] > maxCount) {
        maxCount = typeCount[type];
        dominantType = type;
      }
    });
    
    // Check header hints
    if (header) {
      const headerLower = header.toLowerCase();
      if (headerLower.includes('date')) return 'date';
      if (headerLower.includes('email')) return 'email';
      if (headerLower.includes('phone')) return 'phone';
      if (headerLower.includes('amount') || headerLower.includes('price')) return 'currency';
      if (headerLower.includes('percent')) return 'percentage';
      if (headerLower.includes('name')) return 'name';
      if (headerLower.includes('id') || headerLower.includes('code')) return 'identifier';
    }
    
    return dominantType;
  },
  
  /**
   * Learn from user corrections to column types
   */
  learnFromTypeCorrection: async function(columnIndex, correctType, context = {}) {
    // Store user override
    this.userTypeOverrides[columnIndex] = correctType;
    
    // Track for learning
    this.learningHistory.push({
      columnIndex: columnIndex,
      predicted: this.mlPredictions.types ? this.mlPredictions.types[columnIndex] : null,
      corrected: correctType,
      confidence: this.columnTypeConfidence[columnIndex],
      timestamp: Date.now(),
      context: context
    });
    
    // Send feedback to ML engine if available
    if (this.mlEnabled && this.mlEngine) {
      await this.mlEngine.learnFromFeedback(
        'columnType',
        this.mlPredictions.types[columnIndex],
        correctType,
        {
          confidence: this.columnTypeConfidence[columnIndex],
          columnIndex: columnIndex,
          ...context
        }
      );
    }
    
    // Update backend if available
    if (typeof google !== 'undefined' && google.script && google.script.run) {
      google.script.run.trackMLFeedback(
        'columnType',
        this.mlPredictions.types[columnIndex],
        correctType,
        { columnIndex: columnIndex, confidence: this.columnTypeConfidence[columnIndex] }
      );
    }
  },
  
  /**
   * Get confidence score for column type detection
   */
  getColumnTypeConfidence: function(columnIndex) {
    return this.columnTypeConfidence[columnIndex] || 0;
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
   * ML-Enhanced smart parse with confidence scoring
   */
  smartParseML: async function(data, options = {}) {
    // Initialize ML if not already done
    if (!this.mlEnabled && !options.skipML) {
      await this.initializeML();
    }
    
    // Detect column types using ML
    const columnTypes = await this.detectColumnTypesML(data, options.headers || []);
    
    // Parse with ML-informed structure
    const patterns = this.analyzePatterns(data);
    if (!patterns || patterns.length === 0) return [];
    
    // Enhance column structure detection with ML insights
    const columnStructure = await this.detectColumnStructureML(patterns, columnTypes);
    
    // Parse each row with ML-enhanced structure
    const parsed = [];
    
    for (let i = 0; i < patterns.length; i++) {
      const pattern = patterns[i];
      if (!pattern.tokens || pattern.tokens.length === 0) continue;
      
      const row = await this.parseRowWithMLStructure(
        pattern.tokens,
        pattern.types,
        columnStructure,
        columnTypes
      );
      
      parsed.push(row);
    }
    
    return parsed;
  },
  
  /**
   * Original smart parse (fallback)
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
   * ML-Enhanced column structure detection
   */
  detectColumnStructureML: async function(patterns, mlColumnTypes) {
    // Start with rule-based detection
    const baseStructure = this.detectColumnStructure(patterns);
    
    // Enhance with ML insights if available
    if (mlColumnTypes && mlColumnTypes.length > 0) {
      baseStructure.mlTypes = mlColumnTypes;
      baseStructure.hasMLSupport = true;
      
      // Refine structure based on ML column types
      if (mlColumnTypes.includes('name') && mlColumnTypes.includes('currency')) {
        baseStructure.columns = mlColumnTypes.map(type => {
          switch(type) {
            case 'currency': return 'sales';
            case 'percentage': return 'percent';
            case 'identifier': return 'id';
            default: return type;
          }
        });
      }
      
      // Calculate confidence for structure
      baseStructure.confidence = this.calculateStructureConfidence(patterns, mlColumnTypes);
    }
    
    return baseStructure;
  },
  
  /**
   * Calculate confidence score for detected structure
   */
  calculateStructureConfidence: function(patterns, mlColumnTypes) {
    let confidence = 0.5; // Base confidence
    
    // Check consistency of patterns
    const consistentTokenCounts = patterns.filter(p => 
      p.tokenCount === patterns[0].tokenCount
    ).length / patterns.length;
    
    confidence += consistentTokenCounts * 0.3;
    
    // Check ML confidence scores
    if (this.columnTypeConfidence) {
      const avgMLConfidence = Object.values(this.columnTypeConfidence).reduce(
        (sum, conf) => sum + conf, 0
      ) / Object.values(this.columnTypeConfidence).length;
      
      confidence += avgMLConfidence * 0.2;
    }
    
    return Math.min(confidence, 1.0);
  },
  
  /**
   * Parse row with ML-enhanced structure understanding
   */
  parseRowWithMLStructure: async function(tokens, types, structure, mlColumnTypes) {
    // If we have high confidence ML structure, use it
    if (structure.confidence && structure.confidence > 0.8 && structure.mlTypes) {
      return this.parseRowWithMLTypes(tokens, structure.mlTypes);
    }
    
    // Otherwise fall back to rule-based parsing
    return this.parseRowWithStructure(tokens, types, structure);
  },
  
  /**
   * Parse row using ML column types
   */
  parseRowWithMLTypes: function(tokens, mlTypes) {
    const row = [];
    let tokenIndex = 0;
    
    mlTypes.forEach((type, colIndex) => {
      if (tokenIndex >= tokens.length) {
        row.push('');
        return;
      }
      
      switch(type) {
        case 'name':
          // Names might span multiple tokens
          const nameTokens = [];
          while (tokenIndex < tokens.length && 
                 this.isNameLike(tokens[tokenIndex]) &&
                 tokenIndex < tokens.length - (mlTypes.length - colIndex - 1)) {
            nameTokens.push(tokens[tokenIndex++]);
          }
          row.push(nameTokens.join(' '));
          break;
          
        case 'date':
        case 'email':
        case 'phone':
        case 'currency':
        case 'percentage':
        case 'identifier':
          // Single token types
          row.push(tokens[tokenIndex++]);
          break;
          
        default:
          // Generic text - might be multi-token
          if (colIndex === mlTypes.length - 1) {
            // Last column - take remaining tokens
            row.push(tokens.slice(tokenIndex).join(' '));
            tokenIndex = tokens.length;
          } else {
            row.push(tokens[tokenIndex++]);
          }
      }
    });
    
    return row;
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
   * Interactive column mapping with ML suggestions
   */
  createColumnMapping: async function(sampleData, suggestedColumns) {
    const mapping = {
      samples: sampleData.slice(0, 5),
      suggestedColumns: suggestedColumns,
      mappings: [],
      mlSuggestions: null,
      confidence: {}
    };
    
    // Get ML suggestions if available
    if (this.mlEnabled) {
      const mlTypes = await this.detectColumnTypesML(sampleData);
      mapping.mlSuggestions = mlTypes;
      mapping.confidence = this.columnTypeConfidence;
    }
    
    return mapping;
  },
  
  /**
   * Get parsing suggestions based on data patterns
   */
  getParsingSuggestions: async function(data) {
    const suggestions = [];
    
    // Analyze data patterns
    const patterns = this.analyzePatterns(data);
    
    // Get ML column types
    let mlTypes = null;
    if (this.mlEnabled) {
      mlTypes = await this.detectColumnTypesML(data);
    }
    
    // Generate suggestions
    if (patterns && patterns.length > 0) {
      const avgTokens = patterns.reduce((sum, p) => sum + p.tokenCount, 0) / patterns.length;
      
      if (avgTokens > 3) {
        suggestions.push({
          type: 'multiWordNames',
          message: 'Data appears to contain multi-word names. Consider using smart parsing.',
          confidence: 0.8
        });
      }
      
      // Check for consistent patterns
      const tokenCounts = {};
      patterns.forEach(p => {
        tokenCounts[p.tokenCount] = (tokenCounts[p.tokenCount] || 0) + 1;
      });
      
      const maxCount = Math.max(...Object.values(tokenCounts));
      if (maxCount < patterns.length * 0.7) {
        suggestions.push({
          type: 'inconsistentStructure',
          message: 'Data structure appears inconsistent. Manual mapping may be needed.',
          confidence: 0.7
        });
      }
    }
    
    // Add ML-based suggestions
    if (mlTypes) {
      const lowConfidenceCols = [];
      Object.entries(this.columnTypeConfidence).forEach(([col, conf]) => {
        if (conf < 0.7) {
          lowConfidenceCols.push(parseInt(col));
        }
      });
      
      if (lowConfidenceCols.length > 0) {
        suggestions.push({
          type: 'lowConfidence',
          message: `Columns ${lowConfidenceCols.join(', ')} have low confidence scores. Manual review recommended.`,
          confidence: 0.6,
          affectedColumns: lowConfidenceCols
        });
      }
    }
    
    return suggestions;
  },
  
  /**
   * Export current parsing configuration for reuse
   */
  exportConfiguration: function() {
    return {
      userTypeOverrides: this.userTypeOverrides,
      learningHistory: this.learningHistory,
      mlEnabled: this.mlEnabled,
      columnTypeConfidence: this.columnTypeConfidence,
      timestamp: Date.now()
    };
  },
  
  /**
   * Import saved parsing configuration
   */
  importConfiguration: function(config) {
    if (config.userTypeOverrides) {
      this.userTypeOverrides = config.userTypeOverrides;
    }
    if (config.learningHistory) {
      this.learningHistory = config.learningHistory;
    }
    if (config.columnTypeConfidence) {
      this.columnTypeConfidence = config.columnTypeConfidence;
    }
  }
};