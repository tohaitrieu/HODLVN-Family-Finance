/**
 * ===============================================
 * DASHBOARDREFRESH.GS
 * ===============================================
 * 
 * Helper để force refresh Dashboard custom functions
 * Vì Google Sheets không tự động recalculate custom functions
 * khi data trong sheets khác thay đổi
 */

/**
 * Force refresh Dashboard by updating a hidden timestamp cell
 * This triggers recalculation of all custom functions
 */
/**
 * Force refresh Dashboard by updating a hidden timestamp cell and custom functions
 * This triggers recalculation of all custom functions
 */
function forceDashboardRecalc() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dashboard = ss.getSheetByName(APP_CONFIG.SHEETS.DASHBOARD);
    
    if (!dashboard) {
      Logger.log('⚠️ Dashboard sheet not found');
      return;
    }
    
    // Method 1: Update timestamp in multiple hidden cells
    const timestamp = new Date().getTime();
    dashboard.getRange('Z1').setValue(timestamp);
    dashboard.getRange('Z2').setValue(timestamp + Math.random()); // Add randomness
    
    // Method 2: Force recalculation by temporarily modifying custom function formulas
    _forceCustomFunctionRecalc(dashboard);
    
    // Method 3: Use Google Sheets' built-in recalculation trigger
    SpreadsheetApp.flush(); // Force all pending changes to be applied
    
    // Optional: Show a quick toast notification
    SpreadsheetApp.getActive().toast('Dashboard đã được cập nhật', '✅ Thành công', 1);
    
    Logger.log('✅ Dashboard refreshed at ' + new Date());
  } catch (error) {
    Logger.log('❌ Error refreshing dashboard: ' + error.message);
  }
}

/**
 * Force custom functions to recalculate by modifying and restoring their formulas
 * This is the most reliable way to ensure custom functions update
 */
function _forceCustomFunctionRecalc(dashboard) {
  try {
    // Find all cells with custom functions (AccPayable, AccReceivable)
    const maxRows = 50;
    const maxCols = 15; 
    const dataRange = dashboard.getRange(1, 1, maxRows, maxCols);
    const formulas = dataRange.getFormulas();
    
    const customFunctionCells = [];
    
    // Scan for custom function formulas
    for (let row = 0; row < formulas.length; row++) {
      for (let col = 0; col < formulas[row].length; col++) {
        const formula = formulas[row][col];
        if (formula && (formula.includes('AccPayable') || formula.includes('AccReceivable'))) {
          customFunctionCells.push({
            row: row + 1,
            col: col + 1,
            formula: formula
          });
        }
      }
    }
    
    Logger.log(`Found ${customFunctionCells.length} custom function cells to refresh`);
    
    // Temporarily replace with a dummy formula, then restore original
    customFunctionCells.forEach(cell => {
      const cellRef = dashboard.getRange(cell.row, cell.col);
      
      // Step 1: Clear the formula temporarily
      cellRef.setFormula('');
      
      // Step 2: Restore the original formula (this forces recalculation)
      cellRef.setFormula(cell.formula);
      
      Logger.log(`Refreshed custom function at R${cell.row}C${cell.col}: ${cell.formula}`);
    });
    
    // Force immediate calculation
    SpreadsheetApp.flush();
    
  } catch (error) {
    Logger.log('❌ Error in _forceCustomFunctionRecalc: ' + error.message);
  }
}

/**
 * Auto-refresh Dashboard after data submission
 * Call this after adding income, expense, debt, etc.
 */
function triggerDashboardRefresh() {
  // Use setTimeout to avoid blocking the main form submission
  Utilities.sleep(100); // Small delay to ensure data is written first
  forceDashboardRecalc();
  
  // Additional refresh with a slight delay to catch any async updates
  Utilities.sleep(500);
  _quickRefreshCustomFunctions();
}

/**
 * Quick refresh method that only updates timestamp cells
 * Faster than full forceDashboardRecalc for frequent updates
 */
function _quickRefreshCustomFunctions() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dashboard = ss.getSheetByName(APP_CONFIG.SHEETS.DASHBOARD);
    
    if (!dashboard) {
      return;
    }
    
    // Update timestamp cells with unique values
    const timestamp = new Date().getTime();
    dashboard.getRange('Z1').setValue(timestamp);
    dashboard.getRange('Z2').setValue(timestamp + 1);
    
    // Force immediate recalculation
    SpreadsheetApp.flush();
    
    Logger.log('✅ Quick refresh completed at ' + new Date());
  } catch (error) {
    Logger.log('❌ Error in quick refresh: ' + error.message);
  }
}
