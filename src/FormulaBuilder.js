/**
 * CellPilot Formula Builder Module
 * Handles natural language formula generation and templates
 */

const FormulaBuilder = {
  
  /**
   * Show natural language formula builder interface
   */
  showNaturalLanguageFormulaBuilder: function() {
    return Utils.safeExecute(() => {
      // Check tier access
      const userTier = UserSettings.load('userTier', 'free');
      if (!Config.FEATURE_ACCESS.formula_builder.includes(userTier)) {
        return UIComponents.buildUpgradeCard();
      }
      
      return this.buildFormulaBuilderCard();
    }, UIComponents.buildLoadingCard('Loading formula builder...'), 'showNaturalLanguageFormulaBuilder');
  },
  
  /**
   * Build the formula builder interface card
   */
  buildFormulaBuilderCard: function() {
    const card = CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader()
        .setTitle('Formula Builder')
        .setSubtitle('Describe what you want, get the formula'));
    
    // Input section
    const inputSection = CardService.newCardSection()
      .setHeader('Describe Your Formula');
    
    const descriptionInput = CardService.newTextInput()
      .setFieldName('formulaDescription')
      .setTitle('What do you want to calculate?')
      .setHint('Example: "Sum sales where region is North"')
      .setMultiline(true);
    
    inputSection.addWidget(descriptionInput);
    
    // Examples section
    const examplesSection = CardService.newCardSection()
      .setHeader('Example Descriptions')
      .addWidget(CardService.newTextParagraph()
        .setText(this.getExampleDescriptions()));
    
    // Action buttons
    const actionSection = CardService.newCardSection();
    
    const buildButton = CardService.newTextButton()
      .setText('Build Formula')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('buildFormulaFromDescription'));
    
    actionSection.addWidget(buildButton);
    
    card.addSection(inputSection);
    card.addSection(examplesSection);
    card.addSection(actionSection);
    
    return card.build();
  },
  
  /**
   * Build formula from natural language description
   */
  buildFormulaFromDescription: function(formInputs) {
    return Utils.safeExecute(() => {
      const description = formInputs && formInputs.formulaDescription;
      
      if (!description || description.trim() === '') {
        throw new Error('Please provide a formula description');
      }
      
      // Check and track usage
      const usageCheck = UsageTracker.track('operations');
      if (!usageCheck.allowed) {
        return UIComponents.buildUpgradeCard();
      }
      
      const formula = this.parseNaturalLanguageFormula(description.trim());
      
      if (!formula) {
        return this.buildFormulaNotFoundCard(description);
      }
      
      return this.buildFormulaResultCard(formula, description);
      
    }, Utils.handleError(new Error('Formula generation failed'), 'Failed to generate formula'), 'buildFormulaFromDescription');
  },
  
  /**
   * Parse natural language into spreadsheet formula
   */
  parseNaturalLanguageFormula: function(description) {
    const desc = description.toLowerCase().trim();
    
    // Pattern matching for common formula types
    const patterns = [
      {
        pattern: /sum.*where.*equals?.*|add.*where.*is.*|total.*where.*=/,
        type: 'sumif',
        generator: this.generateSumIfFormula
      },
      {
        pattern: /count.*where.*equals?.*|count.*where.*is.*|how many.*where/,
        type: 'countif',
        generator: this.generateCountIfFormula
      },
      {
        pattern: /average.*where.*equals?.*|mean.*where.*is.*|avg.*where/,
        type: 'averageif',
        generator: this.generateAverageIfFormula
      },
      {
        pattern: /find.*in.*|lookup.*in.*|search.*in.*|get.*from/,
        type: 'vlookup',
        generator: this.generateVLookupFormula
      },
      {
        pattern: /sum.*|add.*|total.*/,
        type: 'sum',
        generator: this.generateSumFormula
      },
      {
        pattern: /count.*|how many/,
        type: 'count',
        generator: this.generateCountFormula
      },
      {
        pattern: /average.*|mean/,
        type: 'average',
        generator: this.generateAverageFormula
      },
      {
        pattern: /max.*|maximum.*|highest.*|largest/,
        type: 'max',
        generator: this.generateMaxFormula
      },
      {
        pattern: /min.*|minimum.*|lowest.*|smallest/,
        type: 'min',
        generator: this.generateMinFormula
      }
    ];
    
    // Find matching pattern
    for (const pattern of patterns) {
      if (pattern.pattern.test(desc)) {
        return pattern.generator.call(this, description, desc);
      }
    }
    
    return null;
  },
  
  /**
   * Generate SUMIF formula
   */
  generateSumIfFormula: function(originalDesc, desc) {
    return {
      formula: '=SUMIF(range, criteria, sum_range)',
      explanation: 'Sums values in sum_range where corresponding cells in range meet the criteria',
      example: '=SUMIF(B:B, "North", C:C)',
      placeholders: {
        range: 'Range containing criteria to check (e.g., B:B)',
        criteria: 'Criteria to match (e.g., "North", ">100")',
        sum_range: 'Range containing values to sum (e.g., C:C)'
      },
      category: 'Conditional Sum'
    };
  },
  
  /**
   * Generate COUNTIF formula
   */
  generateCountIfFormula: function(originalDesc, desc) {
    return {
      formula: '=COUNTIF(range, criteria)',
      explanation: 'Counts cells in range that meet the criteria',
      example: '=COUNTIF(A:A, "Completed")',
      placeholders: {
        range: 'Range to check (e.g., A:A)',
        criteria: 'Criteria to count (e.g., "Completed", ">50")'
      },
      category: 'Conditional Count'
    };
  },
  
  /**
   * Generate AVERAGEIF formula
   */
  generateAverageIfFormula: function(originalDesc, desc) {
    return {
      formula: '=AVERAGEIF(range, criteria, average_range)',
      explanation: 'Calculates average of cells in average_range where corresponding cells in range meet criteria',
      example: '=AVERAGEIF(B:B, "Q1", C:C)',
      placeholders: {
        range: 'Range containing criteria to check (e.g., B:B)',
        criteria: 'Criteria to match (e.g., "Q1", ">0")',
        average_range: 'Range containing values to average (e.g., C:C)'
      },
      category: 'Conditional Average'
    };
  },
  
  /**
   * Generate VLOOKUP formula
   */
  generateVLookupFormula: function(originalDesc, desc) {
    return {
      formula: '=VLOOKUP(lookup_value, table_array, col_index_num, FALSE)',
      explanation: 'Looks up a value in the first column of a table and returns a value in the same row from another column',
      example: '=VLOOKUP(A2, B:D, 3, FALSE)',
      placeholders: {
        lookup_value: 'Value to search for (e.g., A2)',
        table_array: 'Range containing the lookup table (e.g., B:D)',
        col_index_num: 'Column number in table to return value from (e.g., 3)',
        range_lookup: 'FALSE for exact match, TRUE for approximate'
      },
      category: 'Lookup'
    };
  },
  
  /**
   * Generate SUM formula
   */
  generateSumFormula: function(originalDesc, desc) {
    return {
      formula: '=SUM(range)',
      explanation: 'Adds up all numbers in the specified range',
      example: '=SUM(A1:A10)',
      placeholders: {
        range: 'Range to sum (e.g., A1:A10, A:A)'
      },
      category: 'Basic Math'
    };
  },
  
  /**
   * Generate COUNT formula
   */
  generateCountFormula: function(originalDesc, desc) {
    return {
      formula: '=COUNT(range)',
      explanation: 'Counts the number of cells that contain numbers',
      example: '=COUNT(A1:A10)',
      placeholders: {
        range: 'Range to count (e.g., A1:A10, A:A)'
      },
      category: 'Basic Count'
    };
  },
  
  /**
   * Generate AVERAGE formula
   */
  generateAverageFormula: function(originalDesc, desc) {
    return {
      formula: '=AVERAGE(range)',
      explanation: 'Calculates the average (arithmetic mean) of numbers in the range',
      example: '=AVERAGE(A1:A10)',
      placeholders: {
        range: 'Range to average (e.g., A1:A10, A:A)'
      },
      category: 'Basic Math'
    };
  },
  
  /**
   * Generate MAX formula
   */
  generateMaxFormula: function(originalDesc, desc) {
    return {
      formula: '=MAX(range)',
      explanation: 'Returns the largest number in the range',
      example: '=MAX(A1:A10)',
      placeholders: {
        range: 'Range to find maximum in (e.g., A1:A10, A:A)'
      },
      category: 'Basic Math'
    };
  },
  
  /**
   * Generate MIN formula
   */
  generateMinFormula: function(originalDesc, desc) {
    return {
      formula: '=MIN(range)',
      explanation: 'Returns the smallest number in the range',
      example: '=MIN(A1:A10)',
      placeholders: {
        range: 'Range to find minimum in (e.g., A1:A10, A:A)'
      },
      category: 'Basic Math'
    };
  },
  
  /**
   * Build formula result display card
   */
  buildFormulaResultCard: function(formulaData, originalDescription) {
    const card = CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader()
        .setTitle('Generated Formula')
        .setSubtitle(formulaData.category));
    
    // Original request
    const requestSection = CardService.newCardSection()
      .setHeader('Your Request')
      .addWidget(CardService.newTextParagraph()
        .setText(`<i>"${originalDescription}"</i>`));
    
    // Formula section
    const formulaSection = CardService.newCardSection()
      .setHeader('Generated Formula')
      .addWidget(CardService.newTextParagraph()
        .setText(`<font face="Courier New"><b>${formulaData.formula}</b></font>`));
    
    // Explanation
    const explanationSection = CardService.newCardSection()
      .setHeader('How it works')
      .addWidget(CardService.newTextParagraph()
        .setText(formulaData.explanation))
      .addWidget(CardService.newTextParagraph()
        .setText(`<b>Example:</b> <font face="Courier New">${formulaData.example}</font>`));
    
    // Placeholders help
    if (formulaData.placeholders) {
      let placeholderText = '<b>Replace these placeholders:</b><br>';
      Object.entries(formulaData.placeholders).forEach(([key, value]) => {
        placeholderText += `• <font face="Courier New">${key}</font>: ${value}<br>`;
      });
      
      explanationSection.addWidget(CardService.newTextParagraph()
        .setText(placeholderText));
    }
    
    // ML Confidence indicator (if available)
    if (formulaData.confidence) {
      const confidenceSection = CardService.newCardSection()
        .setHeader('Confidence Level');
      
      const confidenceText = this.getConfidenceDescription(formulaData.confidence);
      confidenceSection.addWidget(CardService.newTextParagraph()
        .setText(`${confidenceText} (${Math.round(formulaData.confidence * 100)}%)`));
      
      // Add alternatives if available
      if (formulaData.alternatives && formulaData.alternatives.length > 0) {
        const altText = '<b>Alternative suggestions:</b><br>' +
          formulaData.alternatives.slice(0, 3).map(alt => 
            `• <font face="Courier New">${alt.formula}</font> (${Math.round(alt.confidence * 100)}%)`
          ).join('<br>');
        
        confidenceSection.addWidget(CardService.newTextParagraph()
          .setText(altText));
      }
      
      card.addSection(confidenceSection);
    }
    
    // Action buttons
    const actionSection = CardService.newCardSection();
    
    const copyButton = CardService.newTextButton()
      .setText('Insert Formula')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('insertFormulaWithTracking')
        .setParameters({ 
          formula: formulaData.formula,
          description: originalDescription
        }));
    
    // Add feedback button for ML learning
    const feedbackButton = CardService.newTextButton()
      .setText('This Helped!')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('recordFormulaFeedback')
        .setParameters({ 
          formula: formulaData.formula,
          description: originalDescription,
          helpful: 'true'
        }));
    
    const newSearchButton = CardService.newTextButton()
      .setText('Build Another Formula')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('showNaturalLanguageFormulaBuilder'));
    
    actionSection.addWidget(copyButton);
    if (formulaData.confidence) {
      actionSection.addWidget(feedbackButton);
    }
    actionSection.addWidget(newSearchButton);
    
    card.addSection(requestSection);
    card.addSection(formulaSection);
    card.addSection(explanationSection);
    card.addSection(actionSection);
    
    return card.build();
  },
  
  /**
   * Build "formula not found" card
   */
  buildFormulaNotFoundCard: function(description) {
    const card = CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader()
        .setTitle('Formula Not Found')
        .setSubtitle('Try a different description'));
    
    const section = CardService.newCardSection()
      .addWidget(CardService.newTextParagraph()
        .setText(`I couldn't understand: <i>"${description}"</i>`))
      .addWidget(CardService.newTextParagraph()
        .setText('Try using simpler language or check the examples below:'))
      .addWidget(CardService.newTextParagraph()
        .setText(this.getExampleDescriptions()));
    
    const tryAgainButton = CardService.newTextButton()
      .setText('Try Again')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('showNaturalLanguageFormulaBuilder'));
    
    section.addWidget(tryAgainButton);
    card.addSection(section);
    
    return card.build();
  },
  
  /**
   * Get confidence description
   */
  getConfidenceDescription: function(confidence) {
    if (confidence >= 0.9) return 'Very High Confidence';
    if (confidence >= 0.75) return 'High Confidence';
    if (confidence >= 0.6) return 'Moderate Confidence';
    if (confidence >= 0.4) return 'Low Confidence';
    return 'Uncertain';
  },
  
  /**
   * Insert formula with ML tracking
   */
  insertFormulaWithTracking: function(params) {
    return Utils.safeExecute(() => {
      const formula = params && params.formula;
      const description = params && params.description;
      
      if (!formula) {
        throw new Error('No formula provided');
      }
      
      const activeCell = SpreadsheetApp.getActiveCell();
      activeCell.setFormula(formula);
      
      // Track formula usage for ML learning
      if (description) {
        this.learnFromFormulaSelection(formula, description, true);
      }
      
      return UIComponents.buildSuccessCard(
        'Formula inserted successfully',
        `Formula "${formula}" has been added to cell ${activeCell.getA1Notation()}`
      );
      
    }, Utils.handleError(new Error('Insert failed'), 'Failed to insert formula'), 'insertFormulaWithTracking');
  },
  
  /**
   * Record formula feedback for ML
   */
  recordFormulaFeedback: function(params) {
    const formula = params && params.formula;
    const description = params && params.description;
    const helpful = params && params.helpful === 'true';
    
    if (formula && description) {
      this.learnFromFormulaSelection(formula, description, helpful);
    }
    
    return UIComponents.buildSuccessCard(
      'Thank you!',
      'Your feedback helps improve formula suggestions.'
    );
  },
  
  /**
   * Original insert formula (kept for compatibility)
   */
  insertFormulaToActiveCell: function(params) {
    return Utils.safeExecute(() => {
      const formula = params && params.formula;
      if (!formula) {
        throw new Error('No formula provided');
      }
      
      const activeCell = SpreadsheetApp.getActiveCell();
      activeCell.setFormula(formula);
      
      return UIComponents.buildSuccessCard(
        'Formula inserted successfully',
        `Formula "${formula}" has been added to cell ${activeCell.getA1Notation()}`
      );
      
    }, Utils.handleError(new Error('Insert failed'), 'Failed to insert formula'), 'insertFormulaToActiveCell');
  },
  
  /**
   * Show formula templates library
   */
  showFormulaTemplates: function() {
    return Utils.safeExecute(() => {
      return this.buildTemplatesCard();
    }, UIComponents.buildLoadingCard('Loading formula templates...'), 'showFormulaTemplates');
  },
  
  /**
   * Build ML-enhanced formula templates card
   */
  buildTemplatesCard: async function() {
    // Initialize ML for personalized recommendations
    if (!this.mlEnabled && typeof window !== 'undefined') {
      await this.initializeML();
    }
    
    const card = CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader()
        .setTitle('Formula Templates')
        .setSubtitle('Ready-to-use formulas'));
    
    // Categories of templates with ML personalization
    const categories = ['Recommended', 'Lookup', 'Math', 'Text', 'Date', 'Conditional'];
    
    // Add recommended section if ML is enabled
    if (this.mlEnabled && this.userFormulaHistory.length > 0) {
      const recommendedSection = CardService.newCardSection()
        .setHeader('Recommended for You')
        .setCollapsible(false);
      
      const recommendations = await this.getPersonalizedRecommendations();
      recommendations.forEach(rec => {
        const button = CardService.newDecoratedText()
          .setText(rec.name)
          .setBottomLabel(`${Math.round(rec.confidence * 100)}% match`)
          .setOnClickAction(CardService.newAction()
            .setFunctionName('showTemplateDetail')
            .setParameters({ 
              templateId: rec.id,
              isRecommended: 'true'
            }));
        
        recommendedSection.addWidget(button);
      });
      
      if (recommendations.length > 0) {
        card.addSection(recommendedSection);
      }
    }
    
    // Add standard categories
    categories.filter(cat => cat !== 'Recommended').forEach(category => {
      const section = CardService.newCardSection()
        .setHeader(category)
        .setCollapsible(true);
      
      const templates = this.getTemplatesByCategory(category);
      templates.forEach(template => {
        const button = CardService.newTextButton()
          .setText(template.name)
          .setOnClickAction(CardService.newAction()
            .setFunctionName('showTemplateDetail')
            .setParameters({ templateId: template.id }));
        
        section.addWidget(button);
      });
      
      card.addSection(section);
    });
    
    return card.build();
  },
  
  /**
   * Get formula templates by category
   */
  getTemplatesByCategory: function(category) {
    const allTemplates = {
      'Lookup': [
        { id: 'vlookup_basic', name: 'Basic VLOOKUP', formula: '=VLOOKUP(lookup_value, table_array, col_index, FALSE)' },
        { id: 'vlookup_safe', name: 'VLOOKUP with Error Handling', formula: '=IFERROR(VLOOKUP(lookup_value, table_array, col_index, FALSE), "Not Found")' }
      ],
      'Math': [
        { id: 'sum_basic', name: 'Sum Range', formula: '=SUM(range)' },
        { id: 'sumif', name: 'Conditional Sum', formula: '=SUMIF(range, criteria, sum_range)' },
        { id: 'average_exclude_zero', name: 'Average Excluding Zero', formula: '=AVERAGEIF(range, ">0")' }
      ],
      'Text': [
        { id: 'concatenate', name: 'Combine Text', formula: '=CONCATENATE(text1, " ", text2)' },
        { id: 'proper_case', name: 'Title Case', formula: '=PROPER(text)' }
      ],
      'Date': [
        { id: 'today', name: 'Current Date', formula: '=TODAY()' },
        { id: 'days_between', name: 'Days Between Dates', formula: '=DATEDIF(start_date, end_date, "D")' }
      ],
      'Conditional': [
        { id: 'if_basic', name: 'Basic IF Statement', formula: '=IF(condition, value_if_true, value_if_false)' },
        { id: 'countif', name: 'Count with Condition', formula: '=COUNTIF(range, criteria)' }
      ]
    };
    
    return allTemplates[category] || [];
  },
  
  /**
   * Get personalized formula recommendations
   */
  getPersonalizedRecommendations: async function() {
    if (!this.mlEnabled || !this.mlEngine) {
      return [];
    }
    
    try {
      // Analyze user history to find patterns
      const recentFormulas = this.userFormulaHistory.slice(-20);
      const context = this.getCurrentSheetContext();
      
      // Get ML recommendations
      const recommendations = await this.mlEngine.getPersonalizedFormulas(
        recentFormulas,
        context
      );
      
      // Map to template format
      return recommendations.map(rec => ({
        id: `ml_${rec.type}_${Date.now()}`,
        name: rec.name || rec.formula,
        formula: rec.formula,
        confidence: rec.confidence || 0.5,
        category: rec.category || 'Recommended'
      }));
    } catch (error) {
      console.warn('Failed to get personalized recommendations:', error);
      return [];
    }
  },
  
  /**
   * Predict next formula based on context
   */
  predictNextFormula: async function() {
    if (!this.mlEnabled || !this.mlEngine) {
      return null;
    }
    
    try {
      const context = this.getCurrentSheetContext();
      const recentFormulas = this.userFormulaHistory.slice(-5);
      
      const prediction = await this.mlEngine.predictNextFormula(
        recentFormulas,
        context
      );
      
      return prediction;
    } catch (error) {
      console.warn('Formula prediction failed:', error);
      return null;
    }
  },
  
  /**
   * Get example descriptions for the UI
   */
  getExampleDescriptions: function() {
    return `<b>Try these examples:</b><br>
• "Sum sales where region is North"<br>
• "Count completed tasks"<br>
• "Average revenue by month"<br>
• "Find customer name from ID"<br>
• "Maximum value in column A"<br>
• "Count how many cells contain 'Yes'"`;
  }
};