/**
 * CellM8 Advanced Slide Generator
 * Next-generation intelligent presentation system with comprehensive Google Slides API integration
 */

const CellM8SlideGeneratorAdvanced = {
  
  // Slide dimension constants (in points - 1 point = 1/72 inch)
  SLIDE_WIDTH: 720,
  SLIDE_HEIGHT: 405,
  
  // Professional safe margins with golden ratio
  MARGINS: {
    top: 54,      // 13.3% of height
    bottom: 36,   // 8.9% of height  
    left: 48,     // 6.7% of width
    right: 48     // 6.7% of width
  },
  
  // Content area calculations
  CONTENT: {
    width: 624,   // 720 - 48 - 48
    height: 315,  // 405 - 54 - 36
    centerX: 360, // 720 / 2
    centerY: 202.5 // 405 / 2
  },
  
  // Advanced typography system with hierarchy
  TYPOGRAPHY: {
    hero: { 
      size: 48, 
      family: 'Montserrat', 
      weight: 700,
      lineHeight: 1.1,
      letterSpacing: -0.02
    },
    title: { 
      size: 36, 
      family: 'Montserrat', 
      weight: 600,
      lineHeight: 1.2,
      letterSpacing: -0.01
    },
    subtitle: { 
      size: 24, 
      family: 'Open Sans', 
      weight: 400,
      lineHeight: 1.3,
      letterSpacing: 0
    },
    heading: { 
      size: 20, 
      family: 'Montserrat', 
      weight: 600,
      lineHeight: 1.4,
      letterSpacing: 0
    },
    body: { 
      size: 14, 
      family: 'Open Sans', 
      weight: 400,
      lineHeight: 1.6,
      letterSpacing: 0
    },
    caption: { 
      size: 11, 
      family: 'Open Sans', 
      weight: 400,
      lineHeight: 1.4,
      letterSpacing: 0.01
    },
    data: { 
      size: 10, 
      family: 'Roboto Mono', 
      weight: 400,
      lineHeight: 1.3,
      letterSpacing: 0
    },
    label: {
      size: 9,
      family: 'Open Sans',
      weight: 600,
      lineHeight: 1.2,
      letterSpacing: 0.05
    }
  },
  
  // Professional color palettes with accessibility in mind
  PALETTES: {
    corporate: {
      primary: '#0F172A',     // Deep navy
      secondary: '#334155',   // Dark gray
      accent: '#3B82F6',      // Bright blue
      success: '#10B981',     // Green
      warning: '#F59E0B',     // Amber
      error: '#EF4444',       // Red
      background: '#FFFFFF',
      surface: '#F8FAFC',
      text: '#1E293B',
      textLight: '#64748B',
      border: '#E2E8F0',
      gradient1: '#3B82F6',   // For gradient start
      gradient2: '#8B5CF6'    // For gradient end
    },
    modern: {
      primary: '#6366F1',     // Indigo
      secondary: '#64748B',   // Slate
      accent: '#A855F7',      // Purple
      success: '#22C55E',     // Green
      warning: '#FCD34D',     // Yellow
      error: '#F87171',       // Light red
      background: '#FFFFFF',
      surface: '#F1F5F9',
      text: '#0F172A',
      textLight: '#475569',
      border: '#CBD5E1',
      gradient1: '#6366F1',
      gradient2: '#EC4899'
    },
    elegant: {
      primary: '#1F2937',     // Charcoal
      secondary: '#4B5563',   // Medium gray
      accent: '#059669',      // Emerald
      success: '#34D399',     // Light green
      warning: '#FBBF24',     // Yellow
      error: '#F87171',       // Coral
      background: '#FFFFFF',
      surface: '#F9FAFB',
      text: '#111827',
      textLight: '#6B7280',
      border: '#D1D5DB',
      gradient1: '#047857',
      gradient2: '#0891B2'
    },
    vibrant: {
      primary: '#DC2626',     // Red
      secondary: '#7C3AED',   // Purple
      accent: '#0891B2',      // Cyan
      success: '#16A34A',     // Green
      warning: '#FB923C',     // Orange
      error: '#E11D48',       // Pink
      background: '#FAFAFA',
      surface: '#FEF3C7',
      text: '#18181B',
      textLight: '#52525B',
      border: '#E4E4E7',
      gradient1: '#DC2626',
      gradient2: '#FB923C'
    }
  },
  
  // Predefined professional layouts with smart positioning
  LAYOUTS: {
    titleSlide: {
      type: SlidesApp.PredefinedLayout.TITLE,
      elements: {
        title: { x: 360, y: 162, width: 600, height: 80, align: 'center' },
        subtitle: { x: 360, y: 250, width: 500, height: 40, align: 'center' },
        logo: { x: 630, y: 30, width: 60, height: 60 }
      }
    },
    contentSlide: {
      type: SlidesApp.PredefinedLayout.TITLE_AND_BODY,
      elements: {
        title: { x: 48, y: 30, width: 624, height: 40 },
        content: { x: 48, y: 90, width: 624, height: 280 }
      }
    },
    twoColumn: {
      type: SlidesApp.PredefinedLayout.TITLE_AND_TWO_COLUMNS,
      elements: {
        title: { x: 48, y: 30, width: 624, height: 40 },
        leftColumn: { x: 48, y: 90, width: 300, height: 280 },
        rightColumn: { x: 372, y: 90, width: 300, height: 280 }
      }
    },
    comparison: {
      type: SlidesApp.PredefinedLayout.TITLE_AND_TWO_COLUMNS,
      elements: {
        title: { x: 360, y: 30, width: 624, height: 40, align: 'center' },
        leftTitle: { x: 180, y: 85, width: 280, height: 30, align: 'center' },
        rightTitle: { x: 540, y: 85, width: 280, height: 30, align: 'center' },
        leftContent: { x: 48, y: 125, width: 300, height: 240 },
        rightContent: { x: 372, y: 125, width: 300, height: 240 }
      }
    },
    dashboard: {
      type: SlidesApp.PredefinedLayout.BLANK,
      elements: {
        title: { x: 48, y: 20, width: 400, height: 35 },
        metric1: { x: 48, y: 70, width: 150, height: 100 },
        metric2: { x: 213, y: 70, width: 150, height: 100 },
        metric3: { x: 378, y: 70, width: 150, height: 100 },
        metric4: { x: 543, y: 70, width: 150, height: 100 },
        chart: { x: 48, y: 185, width: 644, height: 200 }
      }
    },
    hero: {
      type: SlidesApp.PredefinedLayout.MAIN_POINT,
      elements: {
        statement: { x: 360, y: 180, width: 600, height: 100, align: 'center' },
        caption: { x: 360, y: 290, width: 400, height: 30, align: 'center' }
      }
    },
    closing: {
      type: SlidesApp.PredefinedLayout.SECTION_HEADER,
      elements: {
        message: { x: 360, y: 150, width: 500, height: 60, align: 'center' },
        contact: { x: 360, y: 230, width: 400, height: 80, align: 'center' }
      }
    }
  },
  
  // Comprehensive chart factory for all 20 Google Charts types
  CHART_FACTORY: {
    
    // Create appropriate chart based on data characteristics
    createSmartChart: function(data, analysis, theme) {
      const chartType = this.determineOptimalChartType(data, analysis);
      return this.createChart(data, chartType, theme);
    },
    
    // Determine best chart type based on data
    determineOptimalChartType: function(data, analysis) {
      if (!analysis) return Charts.ChartType.COLUMN;
      
      // Time series data -> Line or Area chart
      if (analysis.hasTimeSeries) {
        return data.data.length > 50 ? Charts.ChartType.AREA : Charts.ChartType.LINE;
      }
      
      // Categorical comparison -> Column or Bar
      if (analysis.categoricalColumns.length > 0 && analysis.numericColumns.length > 0) {
        const categories = analysis.categoricalColumns[0].uniqueValues;
        return categories > 8 ? Charts.ChartType.BAR : Charts.ChartType.COLUMN;
      }
      
      // Part-to-whole relationship -> Pie or Treemap
      if (analysis.isPartToWhole) {
        const parts = data.data.length;
        return parts <= 8 ? Charts.ChartType.PIE : Charts.ChartType.TREEMAP;
      }
      
      // Distribution data -> Histogram
      if (analysis.isDistribution) {
        return Charts.ChartType.HISTOGRAM;
      }
      
      // Correlation data -> Scatter
      if (analysis.hasCorrelation) {
        return Charts.ChartType.SCATTER;
      }
      
      // Hierarchical data -> Treemap or Org
      if (analysis.isHierarchical) {
        return analysis.isOrganizational ? Charts.ChartType.ORG : Charts.ChartType.TREEMAP;
      }
      
      // Financial data -> Candlestick
      if (analysis.isFinancial) {
        return Charts.ChartType.CANDLESTICK;
      }
      
      // Geographic data -> Geo chart
      if (analysis.isGeographic) {
        return Charts.ChartType.GEO;
      }
      
      // Progress/Goal data -> Gauge
      if (analysis.isProgress) {
        return Charts.ChartType.GAUGE;
      }
      
      // Default to column chart
      return Charts.ChartType.COLUMN;
    },
    
    // Create chart with specific type
    createChart: function(data, chartType, theme) {
      const dataTable = Charts.newDataTable();
      
      // Add headers
      data.headers.forEach((header, index) => {
        const columnType = this.inferColumnType(data.data.map(row => row[index]));
        dataTable.addColumn(columnType, header);
      });
      
      // Add data rows
      data.data.forEach(row => {
        dataTable.addRow(row);
      });
      
      // Build chart with appropriate builder
      const chartBuilder = this.getChartBuilder(chartType, dataTable, theme);
      
      // Apply theme colors
      if (theme && CellM8SlideGeneratorAdvanced.PALETTES[theme]) {
        const palette = CellM8SlideGeneratorAdvanced.PALETTES[theme];
        chartBuilder.setOption('colors', [
          palette.accent,
          palette.primary,
          palette.success,
          palette.warning,
          palette.secondary
        ]);
        chartBuilder.setOption('backgroundColor', palette.background);
        chartBuilder.setOption('chartArea.backgroundColor', palette.surface);
      }
      
      // Set dimensions
      chartBuilder.setDimensions(640, 360);
      
      return chartBuilder.build();
    },
    
    // Get appropriate chart builder
    getChartBuilder: function(chartType, dataTable, theme) {
      switch(chartType) {
        case Charts.ChartType.AREA:
          return Charts.newAreaChart()
            .setDataTable(dataTable)
            .setStacked()
            .setOption('animation.startup', true)
            .setOption('animation.duration', 1000);
            
        case Charts.ChartType.BAR:
          return Charts.newBarChart()
            .setDataTable(dataTable)
            .setOption('bars', 'horizontal')
            .setOption('animation.startup', true);
            
        case Charts.ChartType.BUBBLE:
          return Charts.newBubbleChart()
            .setDataTable(dataTable)
            .setOption('bubble.textStyle.fontSize', 11);
            
        case Charts.ChartType.CANDLESTICK:
          return Charts.newCandlestickChart()
            .setDataTable(dataTable);
            
        case Charts.ChartType.COLUMN:
          return Charts.newColumnChart()
            .setDataTable(dataTable)
            .setOption('animation.startup', true)
            .setOption('animation.duration', 1000);
            
        case Charts.ChartType.COMBO:
          return Charts.newComboChart()
            .setDataTable(dataTable)
            .setOption('seriesType', 'bars');
            
        case Charts.ChartType.GAUGE:
          return Charts.newGaugeChart()
            .setDataTable(dataTable)
            .setOption('animation.startup', true);
            
        case Charts.ChartType.GEO:
          return Charts.newGeoChart()
            .setDataTable(dataTable)
            .setOption('region', 'auto');
            
        case Charts.ChartType.HISTOGRAM:
          return Charts.newHistogram()
            .setDataTable(dataTable)
            .setOption('histogram.bucketSize', 'auto');
            
        case Charts.ChartType.LINE:
          return Charts.newLineChart()
            .setDataTable(dataTable)
            .setCurveStyle(Charts.CurveStyle.SMOOTH)
            .setOption('animation.startup', true);
            
        case Charts.ChartType.ORG:
          return Charts.newOrgChart()
            .setDataTable(dataTable)
            .setOption('allowHtml', true);
            
        case Charts.ChartType.PIE:
          return Charts.newPieChart()
            .setDataTable(dataTable)
            .set3D()
            .setOption('pieSliceText', 'percentage')
            .setOption('animation.startup', true);
            
        case Charts.ChartType.RADAR:
          return Charts.newRadarChart()
            .setDataTable(dataTable);
            
        case Charts.ChartType.SCATTER:
          return Charts.newScatterChart()
            .setDataTable(dataTable)
            .setOption('trendlines', { 0: {} });
            
        case Charts.ChartType.SPARKLINE:
          return Charts.newSparklineChart()
            .setDataTable(dataTable);
            
        case Charts.ChartType.STEPPED_AREA:
          return Charts.newSteppedAreaChart()
            .setDataTable(dataTable);
            
        case Charts.ChartType.TABLE:
          return Charts.newTableChart()
            .setDataTable(dataTable)
            .setOption('showRowNumber', true);
            
        case Charts.ChartType.TIMELINE:
          return Charts.newTimelineChart()
            .setDataTable(dataTable);
            
        case Charts.ChartType.TREEMAP:
          return Charts.newTreemapChart()
            .setDataTable(dataTable)
            .setOption('minColor', '#f0f0f0')
            .setOption('maxColor', theme ? CellM8SlideGeneratorAdvanced.PALETTES[theme].accent : '#3B82F6');
            
        case Charts.ChartType.WATERFALL:
          return Charts.newWaterfallChart()
            .setDataTable(dataTable);
            
        default:
          return Charts.newColumnChart()
            .setDataTable(dataTable);
      }
    },
    
    // Infer column data type for chart
    inferColumnType: function(columnData) {
      const sample = columnData.find(val => val != null && val !== '');
      if (!sample) return Charts.ColumnType.STRING;
      
      if (typeof sample === 'number') return Charts.ColumnType.NUMBER;
      if (sample instanceof Date) return Charts.ColumnType.DATE;
      
      // Try to parse as number
      const parsed = parseFloat(sample);
      if (!isNaN(parsed)) return Charts.ColumnType.NUMBER;
      
      // Try to parse as date
      const dateTest = new Date(sample);
      if (dateTest instanceof Date && !isNaN(dateTest)) return Charts.ColumnType.DATE;
      
      return Charts.ColumnType.STRING;
    }
  },
  
  // Shape composition system for complex visualizations
  SHAPE_COMPOSER: {
    
    // Create metric card with icon
    createMetricCard: function(slide, metric, position, theme) {
      const palette = CellM8SlideGeneratorAdvanced.PALETTES[theme] || CellM8SlideGeneratorAdvanced.PALETTES.corporate;
      
      // Background card
      const card = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 
        position.x, position.y, position.width, position.height);
      card.getFill().setSolidFill(palette.surface);
      card.getBorder().getLineFill().setSolidFill(palette.border);
      card.getBorder().setWeight(1);
      card.setCornerRadius(8);
      
      // Icon background
      const iconBg = slide.insertShape(SlidesApp.ShapeType.ELLIPSE,
        position.x + 12, position.y + 12, 36, 36);
      iconBg.getFill().setSolidFill(palette.accent);
      iconBg.getBorder().setTransparent();
      
      // Metric value
      const valueText = slide.insertTextBox(
        metric.value.toString(),
        position.x + 12, position.y + 56,
        position.width - 24, 32
      );
      valueText.getText().getTextStyle()
        .setFontSize(28)
        .setFontFamily('Montserrat')
        .setBold(true)
        .setForegroundColor(palette.text);
      
      // Metric label
      const labelText = slide.insertTextBox(
        metric.label,
        position.x + 12, position.y + 92,
        position.width - 24, 16
      );
      labelText.getText().getTextStyle()
        .setFontSize(11)
        .setFontFamily('Open Sans')
        .setForegroundColor(palette.textLight);
      
      // Change indicator if provided
      if (metric.change) {
        const changeColor = metric.change > 0 ? palette.success : palette.error;
        const changeIcon = metric.change > 0 ? '↑' : '↓';
        const changeText = slide.insertTextBox(
          changeIcon + ' ' + Math.abs(metric.change) + '%',
          position.x + 12, position.y + position.height - 28,
          position.width - 24, 16
        );
        changeText.getText().getTextStyle()
          .setFontSize(10)
          .setFontFamily('Open Sans')
          .setBold(true)
          .setForegroundColor(changeColor);
      }
      
      return { card, iconBg, valueText, labelText };
    },
    
    // Create progress indicator
    createProgressBar: function(slide, progress, position, theme) {
      const palette = CellM8SlideGeneratorAdvanced.PALETTES[theme] || CellM8SlideGeneratorAdvanced.PALETTES.corporate;
      
      // Background track
      const track = slide.insertShape(SlidesApp.ShapeType.RECTANGLE,
        position.x, position.y, position.width, 8);
      track.getFill().setSolidFill(palette.border);
      track.getBorder().setTransparent();
      
      // Progress fill
      const fillWidth = position.width * (progress / 100);
      const fill = slide.insertShape(SlidesApp.ShapeType.RECTANGLE,
        position.x, position.y, fillWidth, 8);
      fill.getFill().setSolidFill(palette.accent);
      fill.getBorder().setTransparent();
      
      // Percentage label
      const label = slide.insertTextBox(
        progress + '%',
        position.x + position.width + 8, position.y - 4,
        40, 16
      );
      label.getText().getTextStyle()
        .setFontSize(11)
        .setFontFamily('Open Sans')
        .setBold(true)
        .setForegroundColor(palette.text);
      
      return { track, fill, label };
    },
    
    // Create comparison arrows
    createComparisonArrows: function(slide, value1, value2, position, theme) {
      const palette = CellM8SlideGeneratorAdvanced.PALETTES[theme] || CellM8SlideGeneratorAdvanced.PALETTES.corporate;
      
      const comparison = value2 - value1;
      const arrowType = comparison > 0 ? SlidesApp.ShapeType.ARROW_UP : SlidesApp.ShapeType.ARROW_DOWN;
      const arrowColor = comparison > 0 ? palette.success : palette.error;
      
      const arrow = slide.insertShape(arrowType,
        position.x, position.y, 40, 40);
      arrow.getFill().setSolidFill(arrowColor);
      arrow.getBorder().setTransparent();
      
      const percentChange = ((comparison / value1) * 100).toFixed(1);
      const changeText = slide.insertTextBox(
        percentChange + '%',
        position.x + 50, position.y + 10,
        60, 20
      );
      changeText.getText().getTextStyle()
        .setFontSize(16)
        .setFontFamily('Montserrat')
        .setBold(true)
        .setForegroundColor(arrowColor);
      
      return { arrow, changeText };
    },
    
    // Create info callout
    createCallout: function(slide, text, position, theme, type = 'info') {
      const palette = CellM8SlideGeneratorAdvanced.PALETTES[theme] || CellM8SlideGeneratorAdvanced.PALETTES.corporate;
      
      const colors = {
        info: { bg: '#EBF5FF', border: palette.accent, icon: 'ℹ' },
        success: { bg: '#F0FDF4', border: palette.success, icon: '✓' },
        warning: { bg: '#FFFBEB', border: palette.warning, icon: '⚠' },
        error: { bg: '#FEF2F2', border: palette.error, icon: '✕' }
      };
      
      const config = colors[type] || colors.info;
      
      // Callout shape
      const callout = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE,
        position.x, position.y, position.width, position.height);
      callout.getFill().setSolidFill(config.bg);
      callout.getBorder().getLineFill().setSolidFill(config.border);
      callout.getBorder().setWeight(2);
      callout.setCornerRadius(6);
      
      // Icon
      const icon = slide.insertTextBox(
        config.icon,
        position.x + 8, position.y + 8,
        20, 20
      );
      icon.getText().getTextStyle()
        .setFontSize(16)
        .setForegroundColor(config.border);
      
      // Text
      const textBox = slide.insertTextBox(
        text,
        position.x + 36, position.y + 8,
        position.width - 44, position.height - 16
      );
      textBox.getText().getTextStyle()
        .setFontSize(12)
        .setFontFamily('Open Sans')
        .setForegroundColor(palette.text);
      
      return { callout, icon, textBox };
    }
  },
  
  // Intelligent data interpreter for meaningful insights
  DATA_INTERPRETER: {
    
    // Generate intelligent insights from data
    generateInsights: function(data, analysis) {
      const insights = [];
      
      if (!analysis) return insights;
      
      // Completeness insight
      if (analysis.completeness < 100) {
        insights.push({
          type: 'warning',
          title: 'Data Completeness',
          message: `Dataset is ${analysis.completeness}% complete with ${analysis.missingValues} missing values`,
          priority: 2
        });
      }
      
      // Trend insights
      if (analysis.trends) {
        analysis.trends.forEach(trend => {
          insights.push({
            type: trend.direction === 'increasing' ? 'success' : 'info',
            title: `${trend.column} Trend`,
            message: `${trend.column} shows ${trend.strength} ${trend.direction} trend (${trend.percentage}% change)`,
            priority: 1
          });
        });
      }
      
      // Statistical insights
      if (analysis.statistics) {
        Object.keys(analysis.statistics).forEach(column => {
          const stats = analysis.statistics[column];
          if (stats.outliers && stats.outliers > 0) {
            insights.push({
              type: 'warning',
              title: `${column} Outliers`,
              message: `${stats.outliers} outlier values detected in ${column}`,
              priority: 3
            });
          }
          if (stats.cv && stats.cv > 1) {
            insights.push({
              type: 'info',
              title: `${column} Variability`,
              message: `High variability in ${column} (CV: ${stats.cv.toFixed(2)})`,
              priority: 3
            });
          }
        });
      }
      
      // Correlation insights
      if (analysis.correlations) {
        analysis.correlations.forEach(corr => {
          if (Math.abs(corr.value) > 0.7) {
            const strength = Math.abs(corr.value) > 0.9 ? 'Very strong' : 'Strong';
            const direction = corr.value > 0 ? 'positive' : 'negative';
            insights.push({
              type: 'success',
              title: 'Correlation Found',
              message: `${strength} ${direction} correlation between ${corr.column1} and ${corr.column2} (${(corr.value * 100).toFixed(0)}%)`,
              priority: 1
            });
          }
        });
      }
      
      // Category insights
      if (analysis.categories) {
        analysis.categories.forEach(cat => {
          if (cat.dominant) {
            insights.push({
              type: 'info',
              title: `${cat.column} Distribution`,
              message: `"${cat.dominant}" represents ${cat.percentage}% of ${cat.column} values`,
              priority: 2
            });
          }
        });
      }
      
      // Sort by priority
      return insights.sort((a, b) => a.priority - b.priority);
    },
    
    // Generate narrative summary
    generateNarrative: function(data, analysis) {
      const narratives = [];
      
      // Opening statement
      narratives.push(`This dataset contains ${data.data.length} records across ${data.headers.length} fields.`);
      
      // Data quality narrative
      if (analysis.completeness === 100) {
        narratives.push('The dataset is complete with no missing values.');
      } else {
        narratives.push(`The dataset is ${analysis.completeness}% complete.`);
      }
      
      // Key metrics narrative
      if (analysis.keyMetrics) {
        const metrics = analysis.keyMetrics.map(m => `${m.name}: ${m.value}`).join(', ');
        narratives.push(`Key metrics include: ${metrics}.`);
      }
      
      // Trend narrative
      if (analysis.trends && analysis.trends.length > 0) {
        const trend = analysis.trends[0];
        narratives.push(`${trend.column} shows a ${trend.strength} ${trend.direction} trend.`);
      }
      
      // Distribution narrative
      if (analysis.distribution) {
        narratives.push(`The data distribution is ${analysis.distribution.type} with ${analysis.distribution.skew} skew.`);
      }
      
      return narratives.join(' ');
    },
    
    // Generate recommendations
    generateRecommendations: function(data, analysis) {
      const recommendations = [];
      
      // Data quality recommendations
      if (analysis.completeness < 90) {
        recommendations.push({
          title: 'Improve Data Quality',
          action: 'Consider filling missing values or collecting additional data',
          impact: 'high'
        });
      }
      
      // Outlier recommendations
      if (analysis.hasOutliers) {
        recommendations.push({
          title: 'Review Outliers',
          action: 'Investigate and validate outlier values for accuracy',
          impact: 'medium'
        });
      }
      
      // Correlation recommendations
      if (analysis.strongCorrelations) {
        recommendations.push({
          title: 'Leverage Correlations',
          action: 'Use correlated variables for predictive modeling',
          impact: 'high'
        });
      }
      
      // Trend recommendations
      if (analysis.significantTrends) {
        recommendations.push({
          title: 'Monitor Trends',
          action: 'Set up alerts for trend changes and anomalies',
          impact: 'medium'
        });
      }
      
      // Distribution recommendations
      if (analysis.skewedDistribution) {
        recommendations.push({
          title: 'Address Skew',
          action: 'Consider data transformation for better analysis',
          impact: 'low'
        });
      }
      
      return recommendations;
    }
  },
  
  // Master slide template system
  MASTER_TEMPLATES: {
    
    // Apply master template to presentation
    applyMasterTemplate: function(presentation, templateName, theme) {
      const palette = CellM8SlideGeneratorAdvanced.PALETTES[theme] || CellM8SlideGeneratorAdvanced.PALETTES.corporate;
      
      switch(templateName) {
        case 'executive':
          this.applyExecutiveTemplate(presentation, palette);
          break;
        case 'technical':
          this.applyTechnicalTemplate(presentation, palette);
          break;
        case 'sales':
          this.applySalesTemplate(presentation, palette);
          break;
        case 'educational':
          this.applyEducationalTemplate(presentation, palette);
          break;
        default:
          this.applyDefaultTemplate(presentation, palette);
      }
    },
    
    // Executive template - clean and professional
    applyExecutiveTemplate: function(presentation, palette) {
      const slides = presentation.getSlides();
      
      slides.forEach((slide, index) => {
        // Add subtle header line
        const headerLine = slide.insertShape(SlidesApp.ShapeType.RECTANGLE,
          0, 0, 720, 2);
        headerLine.getFill().setSolidFill(palette.accent);
        headerLine.getBorder().setTransparent();
        
        // Add page number
        if (index > 0) {
          const pageNum = slide.insertTextBox(
            (index + 1).toString(),
            680, 375, 30, 20
          );
          pageNum.getText().getTextStyle()
            .setFontSize(10)
            .setFontFamily('Open Sans')
            .setForegroundColor(palette.textLight);
        }
        
        // Add company watermark (subtle)
        const watermark = slide.insertTextBox(
          'CONFIDENTIAL',
          600, 10, 100, 15
        );
        watermark.getText().getTextStyle()
          .setFontSize(8)
          .setFontFamily('Open Sans')
          .setForegroundColor(palette.textLight);
      });
    },
    
    // Technical template - data-focused
    applyTechnicalTemplate: function(presentation, palette) {
      const slides = presentation.getSlides();
      
      slides.forEach((slide, index) => {
        // Add grid background effect
        for (let i = 0; i < 720; i += 60) {
          const gridLine = slide.insertShape(SlidesApp.ShapeType.RECTANGLE,
            i, 0, 1, 405);
          gridLine.getFill().setSolidFill('#F0F0F0');
          gridLine.getBorder().setTransparent();
          gridLine.sendToBack();
        }
        
        // Add technical header
        if (index > 0) {
          const header = slide.insertShape(SlidesApp.ShapeType.RECTANGLE,
            0, 0, 720, 30);
          header.getFill().setSolidFill(palette.primary);
          header.getBorder().setTransparent();
          
          const breadcrumb = slide.insertTextBox(
            `Data Analysis > Slide ${index + 1}`,
            10, 8, 200, 15
          );
          breadcrumb.getText().getTextStyle()
            .setFontSize(10)
            .setFontFamily('Roboto Mono')
            .setForegroundColor('#FFFFFF');
        }
      });
    },
    
    // Sales template - engaging and dynamic
    applySalesTemplate: function(presentation, palette) {
      const slides = presentation.getSlides();
      
      slides.forEach((slide, index) => {
        // Add diagonal accent
        const accent = slide.insertShape(SlidesApp.ShapeType.TRIANGLE,
          620, 0, 100, 100);
        accent.getFill().setSolidFill(palette.accent);
        accent.getBorder().setTransparent();
        accent.setRotation(45);
        
        // Add call-to-action footer
        if (index === slides.length - 1) {
          const ctaBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE,
            260, 340, 200, 50);
          ctaBox.getFill().setSolidFill(palette.accent);
          ctaBox.getBorder().setTransparent();
          ctaBox.setCornerRadius(25);
          
          const ctaText = slide.insertTextBox(
            'Get Started Today',
            260, 355, 200, 20
          );
          ctaText.getText().getTextStyle()
            .setFontSize(14)
            .setFontFamily('Montserrat')
            .setBold(true)
            .setForegroundColor('#FFFFFF');
          ctaText.getText().getParagraphStyle()
            .setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
        }
      });
    },
    
    // Educational template - clear and informative
    applyEducationalTemplate: function(presentation, palette) {
      const slides = presentation.getSlides();
      
      slides.forEach((slide, index) => {
        // Add sidebar for navigation hints
        const sidebar = slide.insertShape(SlidesApp.ShapeType.RECTANGLE,
          0, 0, 40, 405);
        sidebar.getFill().setSolidFill(palette.surface);
        sidebar.getBorder().setTransparent();
        
        // Add section indicator
        const indicator = slide.insertShape(SlidesApp.ShapeType.ELLIPSE,
          15, 180 + (index * 20), 10, 10);
        indicator.getFill().setSolidFill(palette.accent);
        indicator.getBorder().setTransparent();
        
        // Add learning objective box for content slides
        if (index > 0 && index < slides.length - 1) {
          const objectiveBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE,
            50, 360, 200, 35);
          objectiveBox.getFill().setSolidFill('#FFF9E6');
          objectiveBox.getBorder().getLineFill().setSolidFill('#FFB800');
          objectiveBox.getBorder().setWeight(1);
          objectiveBox.setCornerRadius(5);
          
          const objectiveText = slide.insertTextBox(
            'Key Takeaway',
            55, 368, 190, 20
          );
          objectiveText.getText().getTextStyle()
            .setFontSize(11)
            .setFontFamily('Open Sans')
            .setBold(true)
            .setForegroundColor('#B37700');
        }
      });
    },
    
    // Default template - balanced and versatile
    applyDefaultTemplate: function(presentation, palette) {
      const slides = presentation.getSlides();
      
      slides.forEach((slide, index) => {
        // Add footer bar
        const footer = slide.insertShape(SlidesApp.ShapeType.RECTANGLE,
          0, 395, 720, 10);
        footer.getFill().setSolidFill(palette.surface);
        footer.getBorder().setTransparent();
        
        // Add progress indicator
        const progress = slide.insertShape(SlidesApp.ShapeType.RECTANGLE,
          0, 395, (720 / slides.length) * (index + 1), 10);
        progress.getFill().setSolidFill(palette.accent);
        progress.getBorder().setTransparent();
      });
    }
  },
  
  // Visual effects and styling engine
  VISUAL_EFFECTS: {
    
    // Apply gradient-like effect using overlapping shapes
    applyGradientEffect: function(slide, position, colors, theme) {
      const steps = 10;
      const stepHeight = position.height / steps;
      
      for (let i = 0; i < steps; i++) {
        const opacity = 1 - (i * 0.1);
        const rect = slide.insertShape(SlidesApp.ShapeType.RECTANGLE,
          position.x, position.y + (i * stepHeight),
          position.width, stepHeight + 1
        );
        rect.getFill().setSolidFill(colors[0]);
        rect.getBorder().setTransparent();
        rect.sendToBack();
        
        // Note: Direct opacity control is limited, using color lightening instead
        const lightness = Math.floor(255 * (1 - opacity));
        const adjustedColor = this.lightenColor(colors[0], lightness);
        rect.getFill().setSolidFill(adjustedColor);
      }
    },
    
    // Create depth with shadows (simulated)
    addDepthEffect: function(element, depth = 'medium') {
      const offsets = {
        light: { x: 2, y: 2, blur: 4 },
        medium: { x: 4, y: 4, blur: 8 },
        heavy: { x: 8, y: 8, blur: 16 }
      };
      
      const offset = offsets[depth] || offsets.medium;
      
      // Create shadow shape behind element
      const bounds = {
        left: element.getLeft(),
        top: element.getTop(),
        width: element.getWidth(),
        height: element.getHeight()
      };
      
      const shadow = element.getParentPage().insertShape(
        element.getPageElementType() === SlidesApp.PageElementType.SHAPE ? 
          element.asShape().getShapeType() : SlidesApp.ShapeType.RECTANGLE,
        bounds.left + offset.x,
        bounds.top + offset.y,
        bounds.width,
        bounds.height
      );
      
      shadow.getFill().setSolidFill('#00000020');
      shadow.getBorder().setTransparent();
      shadow.sendBackward();
      
      return shadow;
    },
    
    // Create glow effect
    addGlowEffect: function(element, color, intensity = 'medium') {
      const glowSizes = {
        light: 2,
        medium: 4,
        heavy: 8
      };
      
      const size = glowSizes[intensity] || glowSizes.medium;
      const bounds = {
        left: element.getLeft(),
        top: element.getTop(),
        width: element.getWidth(),
        height: element.getHeight()
      };
      
      // Create multiple layers for glow
      for (let i = size; i > 0; i--) {
        const glow = element.getParentPage().insertShape(
          SlidesApp.ShapeType.RECTANGLE,
          bounds.left - i,
          bounds.top - i,
          bounds.width + (i * 2),
          bounds.height + (i * 2)
        );
        
        const opacity = Math.floor(40 / i);
        const glowColor = color + opacity.toString(16).padStart(2, '0');
        glow.getFill().setSolidFill(glowColor);
        glow.getBorder().setTransparent();
        glow.sendBackward();
      }
    },
    
    // Lighten color helper
    lightenColor: function(color, amount) {
      const usePound = color[0] === '#';
      const col = usePound ? color.slice(1) : color;
      const num = parseInt(col, 16);
      const r = Math.min(255, (num >> 16) + amount);
      const g = Math.min(255, ((num >> 8) & 0x00FF) + amount);
      const b = Math.min(255, (num & 0x0000FF) + amount);
      return (usePound ? '#' : '') + (g | (b << 8) | (r << 16)).toString(16).padStart(6, '0');
    }
  },
  
  // Intelligent content generator
  CONTENT_GENERATOR: {
    
    // Generate smart title based on data
    generateSmartTitle: function(data, analysis) {
      if (!data || !data.headers) return 'Data Presentation';
      
      // Look for date columns to add time context
      const timeContext = analysis && analysis.dateColumns && analysis.dateColumns.length > 0 ?
        ' - ' + this.getTimeRange(data, analysis.dateColumns[0].index) : '';
      
      // Look for main metric
      const mainMetric = analysis && analysis.numericColumns && analysis.numericColumns.length > 0 ?
        analysis.numericColumns[0].name : data.headers[0];
      
      // Generate contextual title
      if (analysis && analysis.trends && analysis.trends.length > 0) {
        return `${mainMetric} Analysis${timeContext}`;
      } else if (analysis && analysis.categories && analysis.categories.length > 0) {
        return `${analysis.categories[0].column} Overview${timeContext}`;
      } else {
        return `${mainMetric} Report${timeContext}`;
      }
    },
    
    // Generate executive summary
    generateExecutiveSummary: function(data, analysis) {
      const summary = [];
      
      // Dataset overview
      summary.push(`Dataset contains ${data.data.length} records across ${data.headers.length} metrics`);
      
      // Key findings
      if (analysis) {
        if (analysis.trends && analysis.trends.length > 0) {
          const trend = analysis.trends[0];
          summary.push(`${trend.column} shows ${trend.percentage}% ${trend.direction} trend`);
        }
        
        if (analysis.keyMetrics && analysis.keyMetrics.length > 0) {
          const metric = analysis.keyMetrics[0];
          summary.push(`Average ${metric.name}: ${metric.value}`);
        }
        
        if (analysis.completeness < 100) {
          summary.push(`Data completeness: ${analysis.completeness}%`);
        }
      }
      
      return summary;
    },
    
    // Generate key takeaways
    generateKeyTakeaways: function(data, analysis) {
      const takeaways = [];
      
      if (!analysis) return ['Data has been successfully processed and visualized'];
      
      // Performance takeaways
      if (analysis.performance) {
        if (analysis.performance.growth > 0) {
          takeaways.push(`${analysis.performance.growth}% growth observed in key metrics`);
        }
        if (analysis.performance.efficiency) {
          takeaways.push(`Efficiency improved by ${analysis.performance.efficiency}%`);
        }
      }
      
      // Pattern takeaways
      if (analysis.patterns) {
        takeaways.push(`${analysis.patterns.length} significant patterns identified`);
      }
      
      // Opportunity takeaways
      if (analysis.opportunities) {
        takeaways.push(`${analysis.opportunities.length} improvement opportunities found`);
      }
      
      // Risk takeaways
      if (analysis.risks) {
        takeaways.push(`${analysis.risks.length} potential risks require attention`);
      }
      
      return takeaways.length > 0 ? takeaways : 
        ['Data analysis complete', 'All metrics within expected ranges'];
    },
    
    // Generate next steps
    generateNextSteps: function(data, analysis) {
      const steps = [];
      
      if (!analysis) {
        return ['Review the presented data', 'Identify areas for deeper analysis'];
      }
      
      // Data quality steps
      if (analysis.completeness < 90) {
        steps.push('Improve data collection to address missing values');
      }
      
      // Trend-based steps
      if (analysis.trends && analysis.trends.some(t => t.direction === 'decreasing')) {
        steps.push('Investigate declining metrics and implement corrective actions');
      }
      
      // Outlier steps
      if (analysis.hasOutliers) {
        steps.push('Review and validate outlier data points');
      }
      
      // Opportunity steps
      if (analysis.opportunities && analysis.opportunities.length > 0) {
        steps.push(`Prioritize top ${Math.min(3, analysis.opportunities.length)} improvement opportunities`);
      }
      
      // Monitoring steps
      steps.push('Set up automated monitoring for key metrics');
      steps.push('Schedule follow-up analysis in 30 days');
      
      return steps;
    },
    
    // Get time range from data
    getTimeRange: function(data, dateColumnIndex) {
      const dates = data.data
        .map(row => row[dateColumnIndex])
        .filter(d => d)
        .map(d => new Date(d))
        .filter(d => !isNaN(d));
      
      if (dates.length === 0) return '';
      
      const minDate = new Date(Math.min(...dates));
      const maxDate = new Date(Math.max(...dates));
      
      const formatDate = (date) => {
        const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 
                       'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
        return `${months[date.getMonth()]} ${date.getFullYear()}`;
      };
      
      return formatDate(minDate) + ' to ' + formatDate(maxDate);
    }
  },
  
  /**
   * Main entry point - Generate professional presentation
   */
  generatePresentation: function(presentation, data, config) {
    // Determine theme
    const theme = config.theme || 'corporate';
    const palette = this.PALETTES[theme];
    
    // Analyze data
    const analysis = this.analyzeData(data);
    
    // Apply master template
    const templateType = config.templateType || 'default';
    this.MASTER_TEMPLATES.applyMasterTemplate(presentation, templateType, theme);
    
    // Generate slides
    const slides = presentation.getSlides();
    let slideIndex = 0;
    
    // 1. Title Slide
    if (slideIndex < slides.length) {
      this.createTitleSlide(slides[slideIndex], config, theme, analysis);
      slideIndex++;
    }
    
    // 2. Executive Summary
    if (slideIndex < slides.length) {
      this.createExecutiveSummarySlide(slides[slideIndex], data, analysis, theme);
      slideIndex++;
    }
    
    // 3. Key Metrics Dashboard
    if (slideIndex < slides.length && analysis.keyMetrics) {
      this.createMetricsDashboard(slides[slideIndex], analysis.keyMetrics, theme);
      slideIndex++;
    }
    
    // 4. Data Visualization Slides
    const chartSlides = Math.min(3, slides.length - slideIndex - 2);
    for (let i = 0; i < chartSlides && slideIndex < slides.length; i++) {
      this.createVisualizationSlide(slides[slideIndex], data, analysis, theme, i);
      slideIndex++;
    }
    
    // 5. Insights Slide
    if (slideIndex < slides.length) {
      this.createInsightsSlide(slides[slideIndex], analysis, theme);
      slideIndex++;
    }
    
    // 6. Closing Slide
    if (slideIndex < slides.length) {
      this.createClosingSlide(slides[slideIndex], data, analysis, theme);
      slideIndex++;
    }
    
    return presentation;
  },
  
  // Create title slide with professional design
  createTitleSlide: function(slide, config, theme, analysis) {
    const layout = this.LAYOUTS.titleSlide;
    const palette = this.PALETTES[theme];
    
    // Background gradient effect
    this.VISUAL_EFFECTS.applyGradientEffect(slide, 
      { x: 0, y: 0, width: 720, height: 405 },
      [palette.gradient1, palette.gradient2], theme);
    
    // Title
    const title = this.CONTENT_GENERATOR.generateSmartTitle(config.data, analysis);
    const titleText = slide.insertTextBox(title,
      layout.elements.title.x - layout.elements.title.width/2,
      layout.elements.title.y,
      layout.elements.title.width,
      layout.elements.title.height
    );
    titleText.getText().getTextStyle()
      .setFontSize(this.TYPOGRAPHY.hero.size)
      .setFontFamily(this.TYPOGRAPHY.hero.family)
      .setBold(true)
      .setForegroundColor('#FFFFFF');
    titleText.getText().getParagraphStyle()
      .setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    
    // Subtitle
    if (config.subtitle) {
      const subtitleText = slide.insertTextBox(config.subtitle,
        layout.elements.subtitle.x - layout.elements.subtitle.width/2,
        layout.elements.subtitle.y,
        layout.elements.subtitle.width,
        layout.elements.subtitle.height
      );
      subtitleText.getText().getTextStyle()
        .setFontSize(this.TYPOGRAPHY.subtitle.size)
        .setFontFamily(this.TYPOGRAPHY.subtitle.family)
        .setForegroundColor('#FFFFFFCC');
      subtitleText.getText().getParagraphStyle()
        .setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    }
    
    // Add decorative elements
    const accent1 = slide.insertShape(SlidesApp.ShapeType.ELLIPSE, 600, 300, 150, 150);
    accent1.getFill().setSolidFill(palette.accent + '20');
    accent1.getBorder().setTransparent();
    accent1.sendToBack();
    
    const accent2 = slide.insertShape(SlidesApp.ShapeType.ELLIPSE, -50, 50, 200, 200);
    accent2.getFill().setSolidFill(palette.gradient2 + '15');
    accent2.getBorder().setTransparent();
    accent2.sendToBack();
  },
  
  // Create executive summary slide
  createExecutiveSummarySlide: function(slide, data, analysis, theme) {
    const palette = this.PALETTES[theme];
    
    // Title
    const titleBox = slide.insertTextBox('Executive Summary', 48, 30, 624, 40);
    titleBox.getText().getTextStyle()
      .setFontSize(this.TYPOGRAPHY.title.size)
      .setFontFamily(this.TYPOGRAPHY.title.family)
      .setBold(true)
      .setForegroundColor(palette.primary);
    
    // Generate summary points
    const summaryPoints = this.CONTENT_GENERATOR.generateExecutiveSummary(data, analysis);
    
    // Create summary cards
    summaryPoints.forEach((point, index) => {
      const yPos = 90 + (index * 80);
      
      // Card background
      const card = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE,
        48, yPos, 624, 70);
      card.getFill().setSolidFill(palette.surface);
      card.getBorder().getLineFill().setSolidFill(palette.border);
      card.getBorder().setWeight(1);
      card.setCornerRadius(8);
      
      // Bullet number
      const bullet = slide.insertShape(SlidesApp.ShapeType.ELLIPSE,
        68, yPos + 20, 30, 30);
      bullet.getFill().setSolidFill(palette.accent);
      bullet.getBorder().setTransparent();
      
      const bulletText = slide.insertTextBox((index + 1).toString(),
        68, yPos + 27, 30, 16);
      bulletText.getText().getTextStyle()
        .setFontSize(14)
        .setFontFamily('Montserrat')
        .setBold(true)
        .setForegroundColor('#FFFFFF');
      bulletText.getText().getParagraphStyle()
        .setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
      
      // Summary text
      const summaryText = slide.insertTextBox(point,
        118, yPos + 25, 534, 20);
      summaryText.getText().getTextStyle()
        .setFontSize(this.TYPOGRAPHY.body.size)
        .setFontFamily(this.TYPOGRAPHY.body.family)
        .setForegroundColor(palette.text);
    });
  },
  
  // Create metrics dashboard
  createMetricsDashboard: function(slide, metrics, theme) {
    const layout = this.LAYOUTS.dashboard;
    const palette = this.PALETTES[theme];
    
    // Title
    const titleBox = slide.insertTextBox('Key Metrics Dashboard',
      layout.elements.title.x, layout.elements.title.y,
      layout.elements.title.width, layout.elements.title.height);
    titleBox.getText().getTextStyle()
      .setFontSize(this.TYPOGRAPHY.heading.size)
      .setFontFamily(this.TYPOGRAPHY.heading.family)
      .setBold(true)
      .setForegroundColor(palette.primary);
    
    // Create metric cards
    const metricPositions = [
      layout.elements.metric1,
      layout.elements.metric2,
      layout.elements.metric3,
      layout.elements.metric4
    ];
    
    metrics.slice(0, 4).forEach((metric, index) => {
      if (metricPositions[index]) {
        this.SHAPE_COMPOSER.createMetricCard(
          slide, metric, metricPositions[index], theme
        );
      }
    });
    
    // Add main chart if data available
    if (metrics.length > 0) {
      // Create a simple data visualization
      const chartData = {
        headers: ['Metric', 'Value'],
        data: metrics.slice(0, 6).map(m => [m.name, m.value])
      };
      
      const chart = this.CHART_FACTORY.createChart(
        chartData, 
        Charts.ChartType.COLUMN,
        theme
      );
      
      // Insert chart
      slide.insertChart(chart,
        layout.elements.chart.x,
        layout.elements.chart.y,
        layout.elements.chart.width,
        layout.elements.chart.height
      );
    }
  },
  
  // Create visualization slide
  createVisualizationSlide: function(slide, data, analysis, theme, index) {
    const palette = this.PALETTES[theme];
    const layout = this.LAYOUTS.contentSlide;
    
    // Determine visualization type based on index
    const vizTypes = ['main', 'trend', 'distribution'];
    const vizType = vizTypes[index] || 'main';
    
    // Title based on visualization type
    const titles = {
      main: 'Data Overview',
      trend: 'Trend Analysis',
      distribution: 'Distribution Analysis'
    };
    
    const titleBox = slide.insertTextBox(titles[vizType],
      layout.elements.title.x, layout.elements.title.y,
      layout.elements.title.width, layout.elements.title.height);
    titleBox.getText().getTextStyle()
      .setFontSize(this.TYPOGRAPHY.heading.size)
      .setFontFamily(this.TYPOGRAPHY.heading.family)
      .setBold(true)
      .setForegroundColor(palette.primary);
    
    // Create appropriate chart
    const chart = this.CHART_FACTORY.createSmartChart(data, analysis, theme);
    
    // Insert chart with proper positioning
    slide.insertChart(chart,
      layout.elements.content.x,
      layout.elements.content.y,
      layout.elements.content.width,
      layout.elements.content.height - 40
    );
    
    // Add insight callout
    if (analysis && analysis.insights && analysis.insights[index]) {
      const insight = analysis.insights[index];
      this.SHAPE_COMPOSER.createCallout(
        slide,
        insight.message,
        { x: layout.elements.content.x, 
          y: layout.elements.content.y + layout.elements.content.height - 35,
          width: layout.elements.content.width,
          height: 30 },
        theme,
        insight.type
      );
    }
  },
  
  // Create insights slide
  createInsightsSlide: function(slide, analysis, theme) {
    const palette = this.PALETTES[theme];
    const layout = this.LAYOUTS.contentSlide;
    
    // Title
    const titleBox = slide.insertTextBox('Key Insights & Findings',
      layout.elements.title.x, layout.elements.title.y,
      layout.elements.title.width, layout.elements.title.height);
    titleBox.getText().getTextStyle()
      .setFontSize(this.TYPOGRAPHY.title.size)
      .setFontFamily(this.TYPOGRAPHY.title.family)
      .setBold(true)
      .setForegroundColor(palette.primary);
    
    // Generate insights
    const insights = this.DATA_INTERPRETER.generateInsights(analysis, analysis);
    
    // Display insights
    insights.slice(0, 5).forEach((insight, index) => {
      const yPos = 90 + (index * 55);
      
      // Insight type icon
      const iconColors = {
        success: palette.success,
        warning: palette.warning,
        error: palette.error,
        info: palette.accent
      };
      
      const icon = slide.insertShape(SlidesApp.ShapeType.STAR_5,
        60, yPos, 20, 20);
      icon.getFill().setSolidFill(iconColors[insight.type] || palette.accent);
      icon.getBorder().setTransparent();
      
      // Insight title
      const insightTitle = slide.insertTextBox(insight.title,
        95, yPos, 200, 20);
      insightTitle.getText().getTextStyle()
        .setFontSize(14)
        .setFontFamily('Montserrat')
        .setBold(true)
        .setForegroundColor(palette.text);
      
      // Insight message
      const insightMessage = slide.insertTextBox(insight.message,
        95, yPos + 22, 577, 25);
      insightMessage.getText().getTextStyle()
        .setFontSize(12)
        .setFontFamily('Open Sans')
        .setForegroundColor(palette.textLight);
    });
  },
  
  // Create closing slide
  createClosingSlide: function(slide, data, analysis, theme) {
    const palette = this.PALETTES[theme];
    const layout = this.LAYOUTS.closing;
    
    // Background gradient
    this.VISUAL_EFFECTS.applyGradientEffect(slide,
      { x: 0, y: 0, width: 720, height: 405 },
      [palette.gradient2, palette.gradient1], theme);
    
    // Main message
    const messageBox = slide.insertTextBox('Next Steps',
      layout.elements.message.x - layout.elements.message.width/2,
      layout.elements.message.y,
      layout.elements.message.width,
      layout.elements.message.height);
    messageBox.getText().getTextStyle()
      .setFontSize(this.TYPOGRAPHY.title.size)
      .setFontFamily(this.TYPOGRAPHY.title.family)
      .setBold(true)
      .setForegroundColor('#FFFFFF');
    messageBox.getText().getParagraphStyle()
      .setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    
    // Generate and display next steps
    const nextSteps = this.CONTENT_GENERATOR.generateNextSteps(data, analysis);
    const stepsText = nextSteps.join('\n• ');
    
    const stepsBox = slide.insertTextBox('• ' + stepsText,
      layout.elements.contact.x - layout.elements.contact.width/2,
      layout.elements.contact.y,
      layout.elements.contact.width,
      layout.elements.contact.height);
    stepsBox.getText().getTextStyle()
      .setFontSize(14)
      .setFontFamily('Open Sans')
      .setForegroundColor('#FFFFFFDD');
    stepsBox.getText().getParagraphStyle()
      .setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    
    // Add call-to-action button shape
    const ctaButton = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE,
      260, 340, 200, 45);
    ctaButton.getFill().setSolidFill('#FFFFFF');
    ctaButton.getBorder().setTransparent();
    ctaButton.setCornerRadius(22);
    
    const ctaText = slide.insertTextBox('Get Started',
      260, 352, 200, 20);
    ctaText.getText().getTextStyle()
      .setFontSize(16)
      .setFontFamily('Montserrat')
      .setBold(true)
      .setForegroundColor(palette.primary);
    ctaText.getText().getParagraphStyle()
      .setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
  },
  
  // Analyze data comprehensively
  analyzeData: function(data) {
    if (!data || !data.headers || !data.data) return null;
    
    const analysis = {
      dataTypes: [],
      numericColumns: [],
      categoricalColumns: [],
      dateColumns: [],
      completeness: 100,
      missingValues: 0,
      insights: [],
      trends: [],
      correlations: [],
      statistics: {},
      keyMetrics: [],
      distribution: null,
      hasOutliers: false,
      isHierarchical: false,
      isFinancial: false,
      isGeographic: false,
      hasTimeSeries: false,
      isPartToWhole: false,
      isDistribution: false,
      hasCorrelation: false,
      isProgress: false,
      categories: [],
      performance: {},
      patterns: [],
      opportunities: [],
      risks: []
    };
    
    // Analyze each column
    for (let i = 0; i < data.headers.length; i++) {
      const columnData = data.data.map(row => row[i]);
      const dataType = this.detectDataType(columnData);
      
      analysis.dataTypes.push(dataType);
      
      if (dataType === 'numeric') {
        const stats = this.calculateStatistics(columnData);
        analysis.numericColumns.push({
          index: i,
          name: data.headers[i],
          stats: stats
        });
        analysis.statistics[data.headers[i]] = stats;
        
        // Add to key metrics
        if (stats.average) {
          analysis.keyMetrics.push({
            name: data.headers[i],
            value: Math.round(stats.average),
            change: stats.trend || 0,
            label: 'Average'
          });
        }
        
        // Check for outliers
        if (stats.outliers > 0) {
          analysis.hasOutliers = true;
        }
      } else if (dataType === 'date') {
        analysis.dateColumns.push({
          index: i,
          name: data.headers[i]
        });
        analysis.hasTimeSeries = true;
      } else {
        const uniqueValues = [...new Set(columnData)];
        analysis.categoricalColumns.push({
          index: i,
          name: data.headers[i],
          uniqueValues: uniqueValues.length,
          values: uniqueValues
        });
        
        // Check for geographic data
        if (this.isGeographicData(uniqueValues)) {
          analysis.isGeographic = true;
        }
      }
      
      // Calculate completeness
      const missing = columnData.filter(val => val == null || val === '').length;
      analysis.missingValues += missing;
    }
    
    // Calculate overall completeness
    const totalCells = data.data.length * data.headers.length;
    analysis.completeness = Math.round(((totalCells - analysis.missingValues) / totalCells) * 100);
    
    // Detect trends
    if (analysis.numericColumns.length > 0 && data.data.length > 3) {
      analysis.trends = this.detectTrends(data, analysis.numericColumns);
    }
    
    // Calculate correlations
    if (analysis.numericColumns.length > 1) {
      analysis.correlations = this.calculateCorrelations(data, analysis.numericColumns);
      analysis.hasCorrelation = analysis.correlations.some(c => Math.abs(c.value) > 0.5);
    }
    
    // Determine distribution type
    if (analysis.numericColumns.length > 0) {
      analysis.distribution = this.determineDistribution(
        data.data.map(row => row[analysis.numericColumns[0].index])
      );
      analysis.isDistribution = true;
    }
    
    // Check for financial data patterns
    analysis.isFinancial = this.detectFinancialData(data.headers);
    
    // Check for hierarchical structure
    analysis.isHierarchical = this.detectHierarchicalData(data);
    
    // Detect part-to-whole relationships
    if (analysis.numericColumns.length > 0) {
      analysis.isPartToWhole = this.detectPartToWhole(data, analysis.numericColumns[0].index);
    }
    
    // Generate insights
    analysis.insights = this.DATA_INTERPRETER.generateInsights(data, analysis);
    
    return analysis;
  },
  
  // Helper functions for data analysis
  detectDataType: function(columnData) {
    const nonEmpty = columnData.filter(val => val != null && val !== '');
    if (nonEmpty.length === 0) return 'unknown';
    
    // Check if all numeric
    const allNumeric = nonEmpty.every(val => !isNaN(parseFloat(val)));
    if (allNumeric) return 'numeric';
    
    // Check if all dates
    const allDates = nonEmpty.every(val => {
      const date = new Date(val);
      return date instanceof Date && !isNaN(date);
    });
    if (allDates) return 'date';
    
    return 'categorical';
  },
  
  calculateStatistics: function(columnData) {
    const numbers = columnData
      .filter(val => val != null && val !== '')
      .map(val => parseFloat(val))
      .filter(val => !isNaN(val));
    
    if (numbers.length === 0) return {};
    
    const sum = numbers.reduce((a, b) => a + b, 0);
    const average = sum / numbers.length;
    const sorted = [...numbers].sort((a, b) => a - b);
    const median = sorted[Math.floor(sorted.length / 2)];
    const min = sorted[0];
    const max = sorted[sorted.length - 1];
    
    // Calculate standard deviation
    const variance = numbers.reduce((acc, val) => acc + Math.pow(val - average, 2), 0) / numbers.length;
    const stdDev = Math.sqrt(variance);
    
    // Detect outliers (values beyond 2 standard deviations)
    const outliers = numbers.filter(val => Math.abs(val - average) > 2 * stdDev).length;
    
    // Calculate coefficient of variation
    const cv = stdDev / average;
    
    // Detect trend (simplified)
    const firstHalf = numbers.slice(0, Math.floor(numbers.length / 2));
    const secondHalf = numbers.slice(Math.floor(numbers.length / 2));
    const firstAvg = firstHalf.reduce((a, b) => a + b, 0) / firstHalf.length;
    const secondAvg = secondHalf.reduce((a, b) => a + b, 0) / secondHalf.length;
    const trend = ((secondAvg - firstAvg) / firstAvg) * 100;
    
    return {
      count: numbers.length,
      sum: sum,
      average: average,
      median: median,
      min: min,
      max: max,
      stdDev: stdDev,
      variance: variance,
      cv: cv,
      outliers: outliers,
      trend: trend
    };
  },
  
  detectTrends: function(data, numericColumns) {
    const trends = [];
    
    numericColumns.forEach(col => {
      const values = data.data.map(row => parseFloat(row[col.index])).filter(v => !isNaN(v));
      if (values.length < 3) return;
      
      // Simple linear regression
      const n = values.length;
      const indices = Array.from({length: n}, (_, i) => i);
      const sumX = indices.reduce((a, b) => a + b, 0);
      const sumY = values.reduce((a, b) => a + b, 0);
      const sumXY = indices.reduce((acc, x, i) => acc + x * values[i], 0);
      const sumX2 = indices.reduce((acc, x) => acc + x * x, 0);
      
      const slope = (n * sumXY - sumX * sumY) / (n * sumX2 - sumX * sumX);
      const intercept = (sumY - slope * sumX) / n;
      
      const firstValue = values[0];
      const lastValue = values[values.length - 1];
      const percentageChange = ((lastValue - firstValue) / firstValue) * 100;
      
      trends.push({
        column: col.name,
        direction: slope > 0 ? 'increasing' : 'decreasing',
        strength: Math.abs(slope) > 1 ? 'strong' : 'moderate',
        percentage: Math.round(percentageChange),
        slope: slope,
        intercept: intercept
      });
    });
    
    return trends;
  },
  
  calculateCorrelations: function(data, numericColumns) {
    const correlations = [];
    
    for (let i = 0; i < numericColumns.length - 1; i++) {
      for (let j = i + 1; j < numericColumns.length; j++) {
        const col1 = numericColumns[i];
        const col2 = numericColumns[j];
        
        const values1 = data.data.map(row => parseFloat(row[col1.index])).filter(v => !isNaN(v));
        const values2 = data.data.map(row => parseFloat(row[col2.index])).filter(v => !isNaN(v));
        
        if (values1.length !== values2.length) continue;
        
        const correlation = this.pearsonCorrelation(values1, values2);
        
        correlations.push({
          column1: col1.name,
          column2: col2.name,
          value: correlation
        });
      }
    }
    
    return correlations.sort((a, b) => Math.abs(b.value) - Math.abs(a.value));
  },
  
  pearsonCorrelation: function(x, y) {
    const n = x.length;
    const sumX = x.reduce((a, b) => a + b, 0);
    const sumY = y.reduce((a, b) => a + b, 0);
    const sumXY = x.reduce((acc, val, i) => acc + val * y[i], 0);
    const sumX2 = x.reduce((acc, val) => acc + val * val, 0);
    const sumY2 = y.reduce((acc, val) => acc + val * val, 0);
    
    const numerator = n * sumXY - sumX * sumY;
    const denominator = Math.sqrt((n * sumX2 - sumX * sumX) * (n * sumY2 - sumY * sumY));
    
    return denominator === 0 ? 0 : numerator / denominator;
  },
  
  determineDistribution: function(values) {
    const numbers = values.filter(v => !isNaN(parseFloat(v))).map(v => parseFloat(v));
    if (numbers.length < 3) return { type: 'unknown', skew: 'none' };
    
    const mean = numbers.reduce((a, b) => a + b, 0) / numbers.length;
    const sorted = [...numbers].sort((a, b) => a - b);
    const median = sorted[Math.floor(sorted.length / 2)];
    
    // Simple skewness detection
    const skew = mean > median ? 'positive' : mean < median ? 'negative' : 'none';
    
    // Distribution type (simplified)
    const uniqueValues = [...new Set(numbers)].length;
    const type = uniqueValues < 10 ? 'discrete' : 'continuous';
    
    return { type, skew };
  },
  
  detectFinancialData: function(headers) {
    const financialKeywords = ['price', 'cost', 'revenue', 'profit', 'expense', 
                              'budget', 'sales', 'income', 'balance', 'cash'];
    return headers.some(h => 
      financialKeywords.some(keyword => 
        h.toLowerCase().includes(keyword)
      )
    );
  },
  
  detectHierarchicalData: function(data) {
    // Simple check for parent-child relationships
    const headers = data.headers.map(h => h.toLowerCase());
    return headers.includes('parent') || headers.includes('child') ||
           headers.includes('category') && headers.includes('subcategory');
  },
  
  detectPartToWhole: function(data, numericColumnIndex) {
    const values = data.data.map(row => parseFloat(row[numericColumnIndex])).filter(v => !isNaN(v));
    if (values.length === 0) return false;
    
    // Check if all positive and could represent parts of a whole
    const allPositive = values.every(v => v >= 0);
    const sum = values.reduce((a, b) => a + b, 0);
    
    // Check if values could be percentages
    const couldBePercentages = values.every(v => v >= 0 && v <= 100) && 
                               Math.abs(sum - 100) < 5;
    
    return allPositive && (couldBePercentages || values.length <= 10);
  },
  
  isGeographicData: function(values) {
    const geoKeywords = ['country', 'state', 'city', 'region', 'location', 'address'];
    return values.some(v => 
      typeof v === 'string' && 
      geoKeywords.some(keyword => v.toLowerCase().includes(keyword))
    );
  }
};