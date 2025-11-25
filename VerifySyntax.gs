/**
 * VerifySyntax.gs
 *
 * Quick syntax verification for library context fix.
 * Run this to verify LibraryConfig.gs and Utils.gs are syntactically correct.
 *
 * HOW TO RUN:
 * 1. Open Apps Script Editor
 * 2. Run: verifySyntax()
 * 3. Check Execution Log
 *
 * Created: 2025-11-25
 */

function verifySyntax() {
  Logger.log('╔═══════════════════════════════════════════════════════════════╗');
  Logger.log('║              SYNTAX VERIFICATION - LIBRARY FIX                ║');
  Logger.log('╚═══════════════════════════════════════════════════════════════╝');
  Logger.log('');

  const tests = [];

  // Test 1: LibraryConfig.gs functions exist
  Logger.log('=== Test 1: LibraryConfig.gs Functions ===');
  try {
    // Check if functions exist
    if (typeof initLibrary !== 'function') throw new Error('initLibrary not found');
    if (typeof getLibraryStatus !== 'function') throw new Error('getLibraryStatus not found');
    if (typeof resetLibrary !== 'function') throw new Error('resetLibrary not found');
    if (typeof isLibraryMode !== 'function') throw new Error('isLibraryMode not found');
    if (typeof getLibrarySpreadsheetId !== 'function') throw new Error('getLibrarySpreadsheetId not found');

    Logger.log('✅ All LibraryConfig.gs functions found');
    tests.push({ name: 'LibraryConfig.gs functions', passed: true });
  } catch (error) {
    Logger.log('❌ Error: ' + error.message);
    tests.push({ name: 'LibraryConfig.gs functions', passed: false, error: error.message });
  }

  // Test 2: Utils.gs functions exist
  Logger.log('');
  Logger.log('=== Test 2: Utils.gs Functions ===');
  try {
    if (typeof getSpreadsheet !== 'function') throw new Error('getSpreadsheet not found');
    if (typeof getSheet !== 'function') throw new Error('getSheet not found');
    if (typeof logSpreadsheetContext !== 'function') throw new Error('logSpreadsheetContext not found');

    Logger.log('✅ All Utils.gs functions found');
    tests.push({ name: 'Utils.gs functions', passed: true });
  } catch (error) {
    Logger.log('❌ Error: ' + error.message);
    tests.push({ name: 'Utils.gs functions', passed: false, error: error.message });
  }

  // Test 3: Global variable exists
  Logger.log('');
  Logger.log('=== Test 3: Global Variable ===');
  try {
    // LIBRARY_SPREADSHEET_ID should exist (initially null)
    if (typeof LIBRARY_SPREADSHEET_ID === 'undefined') {
      throw new Error('LIBRARY_SPREADSHEET_ID not defined');
    }
    Logger.log('✅ LIBRARY_SPREADSHEET_ID exists (value: ' + LIBRARY_SPREADSHEET_ID + ')');
    tests.push({ name: 'Global variable', passed: true });
  } catch (error) {
    Logger.log('❌ Error: ' + error.message);
    tests.push({ name: 'Global variable', passed: false, error: error.message });
  }

  // Test 4: Basic function calls work
  Logger.log('');
  Logger.log('=== Test 4: Basic Function Calls ===');
  try {
    // Test getLibraryStatus (should not throw)
    const status = getLibraryStatus();
    if (!status.mode) throw new Error('getLibraryStatus returned invalid object');

    // Test isLibraryMode (should return boolean)
    const mode = isLibraryMode();
    if (typeof mode !== 'boolean') throw new Error('isLibraryMode should return boolean');

    // Test getLibrarySpreadsheetId (should return null or string)
    const id = getLibrarySpreadsheetId();
    if (id !== null && typeof id !== 'string') throw new Error('getLibrarySpreadsheetId invalid return type');

    Logger.log('✅ Basic function calls work correctly');
    tests.push({ name: 'Basic function calls', passed: true });
  } catch (error) {
    Logger.log('❌ Error: ' + error.message);
    tests.push({ name: 'Basic function calls', passed: false, error: error.message });
  }

  // Test 5: getSpreadsheet() works
  Logger.log('');
  Logger.log('=== Test 5: getSpreadsheet() Function ===');
  try {
    const ss = getSpreadsheet();
    if (!ss) throw new Error('getSpreadsheet returned null/undefined');

    const name = ss.getName();
    if (!name) throw new Error('getSpreadsheet returned invalid object');

    Logger.log('✅ getSpreadsheet() works (Spreadsheet: "' + name + '")');
    tests.push({ name: 'getSpreadsheet()', passed: true });
  } catch (error) {
    Logger.log('❌ Error: ' + error.message);
    tests.push({ name: 'getSpreadsheet()', passed: false, error: error.message });
  }

  // Test 6: logSpreadsheetContext() works
  Logger.log('');
  Logger.log('=== Test 6: logSpreadsheetContext() Function ===');
  try {
    const context = logSpreadsheetContext();
    if (!context.mode) throw new Error('logSpreadsheetContext returned invalid object');

    Logger.log('✅ logSpreadsheetContext() works');
    tests.push({ name: 'logSpreadsheetContext()', passed: true });
  } catch (error) {
    Logger.log('❌ Error: ' + error.message);
    tests.push({ name: 'logSpreadsheetContext()', passed: false, error: error.message });
  }

  // Summary
  Logger.log('');
  Logger.log('╔═══════════════════════════════════════════════════════════════╗');
  Logger.log('║                    VERIFICATION SUMMARY                       ║');
  Logger.log('╚═══════════════════════════════════════════════════════════════╝');

  const passed = tests.filter(t => t.passed).length;
  const failed = tests.filter(t => !t.passed).length;

  Logger.log('');
  Logger.log('Total Tests: ' + tests.length);
  Logger.log('Passed: ' + passed + ' ✅');
  Logger.log('Failed: ' + failed + ' ❌');
  Logger.log('');

  if (failed === 0) {
    Logger.log('🎉 SYNTAX VERIFICATION PASSED!');
    Logger.log('LibraryConfig.gs and Utils.gs are syntactically correct.');
    Logger.log('');
    Logger.log('Next steps:');
    Logger.log('1. Run runAllTests() for comprehensive testing');
    Logger.log('2. Proceed with manual deployment testing');
  } else {
    Logger.log('⚠️  VERIFICATION FAILED!');
    Logger.log('Please fix the errors above before deployment.');
    Logger.log('');
    Logger.log('Failed tests:');
    tests.filter(t => !t.passed).forEach(t => {
      Logger.log('  - ' + t.name + ': ' + t.error);
    });
  }

  Logger.log('');
  Logger.log('═══════════════════════════════════════════════════════════════');

  return {
    passed: passed,
    failed: failed,
    total: tests.length,
    success: failed === 0
  };
}

/**
 * Quick test of library initialization with current spreadsheet
 */
function quickTestInit() {
  Logger.log('=== Quick Library Init Test ===');
  Logger.log('');

  try {
    // Get current spreadsheet ID
    const currentId = SpreadsheetApp.getActiveSpreadsheet().getId();
    Logger.log('Current Spreadsheet ID: ' + currentId);

    // Test initialization
    Logger.log('Initializing library...');
    const result = initLibrary(currentId);
    Logger.log('Result: ' + JSON.stringify(result, null, 2));

    if (result.success) {
      Logger.log('✅ Initialization successful');

      // Check status
      const status = getLibraryStatus();
      Logger.log('Status: ' + JSON.stringify(status, null, 2));

      // Reset
      Logger.log('Resetting...');
      const resetResult = resetLibrary();
      Logger.log('Reset: ' + JSON.stringify(resetResult, null, 2));

      Logger.log('✅ Quick test passed');
    } else {
      Logger.log('❌ Initialization failed: ' + result.message);
    }

  } catch (error) {
    Logger.log('❌ Error: ' + error.message);
    Logger.log('Stack: ' + error.stack);
  }
}
