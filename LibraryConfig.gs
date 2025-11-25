/**
 * LibraryConfig.gs
 *
 * Configuration module for Google Apps Script library deployment.
 * Handles spreadsheet context when code is used as a library vs standalone.
 *
 * PROBLEM SOLVED:
 * When code is deployed as a library, SpreadsheetApp.getActiveSpreadsheet()
 * returns the CALLER's spreadsheet, not the library's data spreadsheet.
 * This causes all data writes to fail or go to wrong location.
 *
 * SOLUTION:
 * Store target spreadsheet ID in global config, use openById() in library mode.
 * Automatically routes all 45+ functions to correct spreadsheet with zero changes.
 *
 * USAGE:
 *
 * // Standalone Mode (default - no changes needed):
 * addIncome({ date: '2025-11-25', amount: 1000000, source: 'Salary' });
 *
 * // Library Mode (from external spreadsheet):
 * const FinanceLib = YOUR_LIBRARY_NAME;
 * FinanceLib.initLibrary('YOUR_DATA_SPREADSHEET_ID');
 * FinanceLib.addIncome({ date: '2025-11-25', amount: 1000000, source: 'Salary' });
 *
 * Created: 2025-11-25
 * Version: 3.1.0
 */

/**
 * Global configuration object for library mode.
 * DO NOT modify directly - use initLibrary() and resetLibrary() functions.
 */
var LIBRARY_SPREADSHEET_ID = null;

/**
 * Initialize library for use from external spreadsheet.
 * MUST be called before using any library functions when deployed as library.
 *
 * @param {string} spreadsheetId - The ID of the target data spreadsheet
 * @returns {Object} Status object with success flag and message
 *
 * @example
 * // From external spreadsheet:
 * const FinanceLib = YOUR_LIBRARY_NAME;
 * const result = FinanceLib.initLibrary('1a2b3c4d5e6f7g8h9i0j');
 * if (!result.success) {
 *   Logger.log('Error: ' + result.message);
 * }
 */
function initLibrary(spreadsheetId) {
  // === VALIDATION ===

  // Check if spreadsheetId is provided
  if (!spreadsheetId) {
    return {
      success: false,
      message: '❌ Error: spreadsheetId is required. Please provide the ID of your data spreadsheet.'
    };
  }

  // Check if spreadsheetId is a string
  if (typeof spreadsheetId !== 'string') {
    return {
      success: false,
      message: '❌ Error: spreadsheetId must be a string. Received: ' + typeof spreadsheetId
    };
  }

  // Check if spreadsheetId is not empty
  if (spreadsheetId.trim().length === 0) {
    return {
      success: false,
      message: '❌ Error: spreadsheetId cannot be empty.'
    };
  }

  // === SPREADSHEET ACCESS TEST ===

  try {
    // Try to open the spreadsheet
    const testSpreadsheet = SpreadsheetApp.openById(spreadsheetId);

    // Verify spreadsheet is accessible
    const spreadsheetName = testSpreadsheet.getName();

    // === REQUIRED SHEETS VALIDATION ===

    // Define required sheets for HODLVN-Family-Finance
    const requiredSheets = [
      'THU',           // Income
      'CHI',           // Expense
      'BUDGET',        // Budget
      'DASHBOARD'      // Dashboard
    ];

    // Check for missing sheets
    const missingSheets = [];
    for (let i = 0; i < requiredSheets.length; i++) {
      const sheetName = requiredSheets[i];
      const sheet = testSpreadsheet.getSheetByName(sheetName);
      if (!sheet) {
        missingSheets.push(sheetName);
      }
    }

    // If any required sheets are missing, return error
    if (missingSheets.length > 0) {
      return {
        success: false,
        message: '❌ Error: Missing required sheets in spreadsheet "' + spreadsheetName + '": ' + missingSheets.join(', ') +
                 '\n\nThis spreadsheet does not appear to be a HODLVN-Family-Finance data spreadsheet. ' +
                 'Please initialize the spreadsheet first using Setup Wizard.'
      };
    }

    // === INITIALIZATION ===

    // Store the spreadsheet ID in global config
    LIBRARY_SPREADSHEET_ID = spreadsheetId;

    // Return success
    return {
      success: true,
      message: '✅ Library initialized successfully for spreadsheet: "' + spreadsheetName + '" (ID: ' + spreadsheetId + ')',
      spreadsheetId: spreadsheetId,
      spreadsheetName: spreadsheetName,
      mode: 'LIBRARY'
    };

  } catch (error) {
    // Handle errors (invalid ID, permission denied, etc.)
    return {
      success: false,
      message: '❌ Error: Failed to access spreadsheet with ID "' + spreadsheetId + '". ' +
               'Reason: ' + error.message +
               '\n\nPlease verify:' +
               '\n1. The spreadsheet ID is correct' +
               '\n2. You have access to this spreadsheet' +
               '\n3. The spreadsheet is a valid HODLVN-Family-Finance data spreadsheet'
    };
  }
}

/**
 * Get current library status and configuration.
 * Useful for debugging and verifying library initialization.
 *
 * @returns {Object} Status object with mode, spreadsheet info, and initialization status
 *
 * @example
 * const status = FinanceLib.getLibraryStatus();
 * Logger.log('Mode: ' + status.mode);
 * Logger.log('Initialized: ' + status.initialized);
 */
function getLibraryStatus() {
  // Check if library mode is active
  const isLibraryMode = LIBRARY_SPREADSHEET_ID !== null;

  // Build status object
  const status = {
    mode: isLibraryMode ? 'LIBRARY' : 'STANDALONE',
    initialized: isLibraryMode,
    spreadsheetId: LIBRARY_SPREADSHEET_ID || 'NONE (using active spreadsheet)'
  };

  // If library mode, get additional info
  if (isLibraryMode) {
    try {
      const ss = SpreadsheetApp.openById(LIBRARY_SPREADSHEET_ID);
      status.spreadsheetName = ss.getName();
      status.spreadsheetUrl = ss.getUrl();
    } catch (error) {
      status.error = 'Cannot access spreadsheet: ' + error.message;
    }
  }

  return status;
}

/**
 * Reset library configuration back to standalone mode.
 * Useful for testing or switching between spreadsheets.
 *
 * @returns {Object} Status object confirming reset
 *
 * @example
 * const result = FinanceLib.resetLibrary();
 * Logger.log(result.message); // "✅ Library reset to standalone mode"
 */
function resetLibrary() {
  const wasLibraryMode = LIBRARY_SPREADSHEET_ID !== null;
  const previousId = LIBRARY_SPREADSHEET_ID;

  // Reset to standalone mode
  LIBRARY_SPREADSHEET_ID = null;

  return {
    success: true,
    message: wasLibraryMode
      ? '✅ Library reset to standalone mode (was using spreadsheet: ' + previousId + ')'
      : '✅ Already in standalone mode (no reset needed)',
    mode: 'STANDALONE',
    previousSpreadsheetId: previousId
  };
}

/**
 * Check if library is currently in library mode.
 *
 * @returns {boolean} True if library mode is active, false if standalone mode
 */
function isLibraryMode() {
  return LIBRARY_SPREADSHEET_ID !== null;
}

/**
 * Get the target spreadsheet ID (for library mode).
 * Returns null if in standalone mode.
 *
 * @returns {string|null} The target spreadsheet ID, or null if standalone mode
 */
function getLibrarySpreadsheetId() {
  return LIBRARY_SPREADSHEET_ID;
}
