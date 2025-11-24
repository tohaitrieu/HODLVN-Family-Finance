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
 * Force refresh Dashboard by updating a hidden timestamp cell
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
    
    // Update a hidden cell with current timestamp to force recalc
    // Use cell Z1 (far right, usually hidden)
    dashboard.getRange('Z1').setValue(new Date().getTime());
    
    // Optional: Show a quick toast notification
    SpreadsheetApp.getActive().toast('Dashboard đã được cập nhật', '✅ Thành công', 1);
    
    Logger.log('✅ Dashboard refreshed at ' + new Date());
  } catch (error) {
    Logger.log('❌ Error refreshing dashboard: ' + error.message);
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
}
