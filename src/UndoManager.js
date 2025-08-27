/**
 * UndoManager - Handles undo/redo operations without creating backup sheets
 * Uses PropertiesService to store temporary undo data
 */

const UndoManager = {
  MAX_UNDO_SIZE: 50000, // Max characters for stored data (Properties Service limit)
  
  /**
   * Save state for undo
   */
  saveState: function(range, operationType) {
    try {
      const data = range.getValues();
      const formulas = range.getFormulas();
      
      // Create undo state object
      const undoState = {
        timestamp: Date.now(),
        operationType: operationType,
        rangeA1: range.getA1Notation(),
        sheetName: range.getSheet().getName(),
        rowStart: range.getRow(),
        colStart: range.getColumn(),
        numRows: range.getNumRows(),
        numCols: range.getNumColumns(),
        data: data,
        formulas: formulas
      };
      
      // Compress the data to fit within limits
      const stateString = JSON.stringify(undoState);
      
      // Check size limit
      if (stateString.length > this.MAX_UNDO_SIZE) {
        // If too large, only save a portion of the data
        undoState.data = data.slice(0, Math.min(100, data.length));
        undoState.truncated = true;
      }
      
      // Store in Properties Service (user-specific)
      const userProperties = PropertiesService.getUserProperties();
      userProperties.setProperty('lastUndoState', JSON.stringify(undoState));
      userProperties.setProperty('lastUndoTimestamp', String(Date.now()));
      
      return true;
    } catch (error) {
      console.error('Failed to save undo state:', error);
      return false;
    }
  },
  
  /**
   * Perform undo operation
   */
  undo: function() {
    try {
      const userProperties = PropertiesService.getUserProperties();
      const stateString = userProperties.getProperty('lastUndoState');
      
      if (!stateString) {
        return {
          success: false,
          error: 'No undo history available'
        };
      }
      
      const state = JSON.parse(stateString);
      
      // Check if undo is too old (more than 30 minutes)
      const ageMinutes = (Date.now() - state.timestamp) / (1000 * 60);
      if (ageMinutes > 30) {
        return {
          success: false,
          error: 'Undo history has expired (over 30 minutes old)'
        };
      }
      
      // Get the target sheet and range
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = spreadsheet.getSheetByName(state.sheetName);
      
      if (!sheet) {
        return {
          success: false,
          error: 'Original sheet not found'
        };
      }
      
      // Restore the range
      const range = sheet.getRange(
        state.rowStart,
        state.colStart,
        state.numRows,
        state.numCols
      );
      
      // Clear the range first
      range.clear();
      
      // Restore values
      if (state.data && state.data.length > 0) {
        const targetRange = sheet.getRange(
          state.rowStart,
          state.colStart,
          state.data.length,
          state.data[0].length
        );
        targetRange.setValues(state.data);
      }
      
      // Restore formulas
      if (state.formulas) {
        state.formulas.forEach((row, i) => {
          row.forEach((formula, j) => {
            if (formula) {
              sheet.getRange(state.rowStart + i, state.colStart + j).setFormula(formula);
            }
          });
        });
      }
      
      // Clear undo state after successful undo
      userProperties.deleteProperty('lastUndoState');
      userProperties.deleteProperty('lastUndoTimestamp');
      
      return {
        success: true,
        message: `Undid ${state.operationType} operation`,
        truncated: state.truncated
      };
      
    } catch (error) {
      console.error('Undo failed:', error);
      return {
        success: false,
        error: 'Failed to undo: ' + error.message
      };
    }
  },
  
  /**
   * Check if undo is available
   */
  canUndo: function() {
    try {
      const userProperties = PropertiesService.getUserProperties();
      const timestamp = userProperties.getProperty('lastUndoTimestamp');
      
      if (!timestamp) return false;
      
      // Check if undo is less than 30 minutes old
      const ageMinutes = (Date.now() - parseInt(timestamp)) / (1000 * 60);
      return ageMinutes <= 30;
      
    } catch (error) {
      return false;
    }
  },
  
  /**
   * Get undo information
   */
  getUndoInfo: function() {
    try {
      const userProperties = PropertiesService.getUserProperties();
      const stateString = userProperties.getProperty('lastUndoState');
      
      if (!stateString) return null;
      
      const state = JSON.parse(stateString);
      const ageMinutes = Math.floor((Date.now() - state.timestamp) / (1000 * 60));
      
      // Return null if undo has expired (over 30 minutes)
      if (ageMinutes > 30) {
        return null;
      }
      
      return {
        operationType: state.operationType,
        ageMinutes: ageMinutes,
        rangeA1: state.rangeA1,
        sheetName: state.sheetName
      };
      
    } catch (error) {
      return null;
    }
  },
  
  /**
   * Clear undo history
   */
  clearUndo: function() {
    const userProperties = PropertiesService.getUserProperties();
    userProperties.deleteProperty('lastUndoState');
    userProperties.deleteProperty('lastUndoTimestamp');
  }
};