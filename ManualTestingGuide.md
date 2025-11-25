# Manual Testing Guide - Library Context Fix

**Version:** 3.1.0
**Date:** 2025-11-25
**Purpose:** Step-by-step manual testing for library deployment

---

## Prerequisites

- [ ] Google Account with Apps Script access
- [ ] Access to 2 Google Sheets for testing
- [ ] Basic understanding of Apps Script editor
- [ ] 30-45 minutes for complete testing

---

## Test Scenario 1: Standalone Mode (Backward Compatibility)

**Goal:** Verify existing users unaffected

### Steps:

1. **Open your existing HODLVN-Family-Finance spreadsheet**
   - Or create copy from template

2. **Add test income without any initialization**
   ```
   Menu: HODLVN Finance → Thu nhập (Income)
   Date: Today
   Amount: 1,000,000
   Source: Test Standalone Mode
   Category: Lương
   Submit
   ```

3. **Verify results:**
   - [ ] Data appears in THU sheet
   - [ ] No errors in execution log
   - [ ] Dashboard updates correctly

**Expected:** ✅ Works exactly as before

---

## Test Scenario 2: Library Deployment

**Goal:** Deploy code as library

### Steps:

1. **Open Apps Script project**
   ```
   Extensions → Apps Script
   ```

2. **Deploy as library**
   ```
   Deploy → New deployment
   Type: Library
   Description: HODLVN-Family-Finance v3.1.0
   Version: New version
   Click: Deploy
   ```

3. **Copy Script ID**
   ```
   Format: 1A2B3C4D5E6F7G8H9I0J...
   Save for next steps
   ```

**Expected:** ✅ Deployment succeeds, Script ID obtained

---

## Test Scenario 3: Target Data Spreadsheet Setup

**Goal:** Create spreadsheet that will store data

### Steps:

1. **Create new Google Sheet**
   ```
   Name: "Finance Data - Library Test"
   ```

2. **Open Apps Script**
   ```
   Extensions → Apps Script
   ```

3. **Copy all .gs files** from main project:
   - Main.gs
   - SheetInitializer.gs
   - DataProcessor.gs
   - BudgetManager.gs
   - DashboardManager.gs
   - Utils.gs
   - LibraryConfig.gs
   - AssetsPrice.gs

4. **Copy all .html files** from forms folder

5. **Run Setup Wizard**
   ```
   In spreadsheet: Menu → HODLVN Finance → 🚀 Setup Wizard
   Fill in basic info
   Click: Khởi tạo hệ thống
   ```

6. **Verify sheets created:**
   - [ ] THU (Income)
   - [ ] CHI (Expense)
   - [ ] BUDGET
   - [ ] DASHBOARD
   - [ ] Other required sheets

7. **Copy Spreadsheet ID from URL:**
   ```
   URL: https://docs.google.com/spreadsheets/d/[SPREADSHEET_ID]/edit
   Example: 1a2b3c4d5e6f7g8h9i0j1k2l3m4n5o6p7q8r9s0t
   ```

**Expected:** ✅ Complete finance spreadsheet ready

---

## Test Scenario 4: Consumer Spreadsheet Setup

**Goal:** Create spreadsheet that uses library

### Steps:

1. **Create another new Google Sheet**
   ```
   Name: "Finance Consumer - Library Test"
   ```

2. **Open Apps Script**
   ```
   Extensions → Apps Script
   ```

3. **Add library**
   ```
   Left sidebar: Libraries → +
   Script ID: [Paste from Scenario 2]
   Version: Latest (or specific version)
   Identifier: FinanceLib
   Click: Add
   ```

4. **Create test functions in Code.gs:**
   ```javascript
   function testLibraryInit() {
     const targetDataSpreadsheetId = 'PASTE_YOUR_DATA_SPREADSHEET_ID_HERE';

     Logger.log('Initializing library...');
     const result = FinanceLib.initLibrary(targetDataSpreadsheetId);
     Logger.log(JSON.stringify(result, null, 2));

     if (result.success) {
       Logger.log('✅ Library initialized successfully');

       // Check status
       const status = FinanceLib.getLibraryStatus();
       Logger.log('Status: ' + JSON.stringify(status, null, 2));
     } else {
       Logger.log('❌ Initialization failed: ' + result.message);
     }
   }
   ```

5. **Run testLibraryInit()**
   ```
   Select function: testLibraryInit
   Click: Run
   Authorize if prompted
   ```

6. **Check Execution Log**
   ```
   View → Logs
   ```

**Expected Output:**
```json
{
  "success": true,
  "message": "✅ Library initialized successfully...",
  "spreadsheetId": "1a2b3c...",
  "spreadsheetName": "Finance Data - Library Test",
  "mode": "LIBRARY"
}
```

**Expected:** ✅ Library initializes, status shows LIBRARY mode

---

## Test Scenario 5: Data Operations Through Library

**Goal:** Verify data writes to correct location

### Steps:

1. **In Consumer spreadsheet Apps Script, add test function:**
   ```javascript
   function testAddIncomeViaLibrary() {
     const targetDataSpreadsheetId = 'YOUR_DATA_SPREADSHEET_ID';

     // Initialize library
     FinanceLib.initLibrary(targetDataSpreadsheetId);
     Logger.log('✅ Library initialized');

     // Prepare income data
     const incomeData = {
       date: new Date(),
       amount: 5000000,
       source: 'Library Test Income',
       category: 'Lương',
       notes: 'Testing library data write'
     };

     // Add income through library
     Logger.log('Adding income...');
     FinanceLib.addIncome(incomeData);
     Logger.log('✅ Income added');

     // Verify context
     const status = FinanceLib.getLibraryStatus();
     Logger.log('Final status: ' + JSON.stringify(status, null, 2));
   }
   ```

2. **Run testAddIncomeViaLibrary()**

3. **Check Consumer spreadsheet:**
   - [ ] No THU sheet created
   - [ ] No data written here

4. **Check Target Data spreadsheet:**
   - [ ] THU sheet has new row
   - [ ] Amount: 5,000,000
   - [ ] Source: "Library Test Income"
   - [ ] Timestamp: Current time

**Expected:** ✅ Data written to Target Data spreadsheet ONLY

---

## Test Scenario 6: Validation Tests

**Goal:** Verify error handling

### Steps:

1. **Test invalid spreadsheet ID:**
   ```javascript
   function testInvalidSpreadsheetId() {
     const result = FinanceLib.initLibrary('INVALID_ID_12345');
     Logger.log(JSON.stringify(result, null, 2));
   }
   ```

   **Expected:** ❌ Error message about invalid ID

2. **Test missing spreadsheet ID:**
   ```javascript
   function testMissingSpreadsheetId() {
     const result = FinanceLib.initLibrary();
     Logger.log(JSON.stringify(result, null, 2));
   }
   ```

   **Expected:** ❌ Error message "spreadsheetId is required"

3. **Test blank spreadsheet (no required sheets):**
   ```javascript
   function testBlankSpreadsheet() {
     // Create blank spreadsheet, copy ID
     const blankSpreadsheetId = 'BLANK_SPREADSHEET_ID';
     const result = FinanceLib.initLibrary(blankSpreadsheetId);
     Logger.log(JSON.stringify(result, null, 2));
   }
   ```

   **Expected:** ❌ Error message "Missing required sheets: THU, CHI, BUDGET, DASHBOARD"

---

## Test Scenario 7: Reset and Status

**Goal:** Verify state management

### Steps:

1. **Test status before init:**
   ```javascript
   function testStatusBeforeInit() {
     FinanceLib.resetLibrary(); // Ensure clean state
     const status = FinanceLib.getLibraryStatus();
     Logger.log(JSON.stringify(status, null, 2));
   }
   ```

   **Expected:**
   ```json
   {
     "mode": "STANDALONE",
     "initialized": false,
     "spreadsheetId": "NONE (using active spreadsheet)"
   }
   ```

2. **Test status after init:**
   ```javascript
   function testStatusAfterInit() {
     const targetId = 'YOUR_DATA_SPREADSHEET_ID';
     FinanceLib.initLibrary(targetId);
     const status = FinanceLib.getLibraryStatus();
     Logger.log(JSON.stringify(status, null, 2));
   }
   ```

   **Expected:**
   ```json
   {
     "mode": "LIBRARY",
     "initialized": true,
     "spreadsheetId": "1a2b3c...",
     "spreadsheetName": "Finance Data - Library Test",
     "spreadsheetUrl": "https://..."
   }
   ```

3. **Test reset:**
   ```javascript
   function testReset() {
     const targetId = 'YOUR_DATA_SPREADSHEET_ID';

     // Init
     FinanceLib.initLibrary(targetId);
     Logger.log('After init: ' + FinanceLib.isLibraryMode());

     // Reset
     const resetResult = FinanceLib.resetLibrary();
     Logger.log('Reset result: ' + JSON.stringify(resetResult, null, 2));
     Logger.log('After reset: ' + FinanceLib.isLibraryMode());
   }
   ```

   **Expected:**
   ```
   After init: true
   Reset result: { "success": true, "mode": "STANDALONE", ... }
   After reset: false
   ```

---

## Test Scenario 8: Multiple Operations

**Goal:** Verify multiple operations work correctly

### Steps:

1. **Run comprehensive test:**
   ```javascript
   function testMultipleOperations() {
     const targetId = 'YOUR_DATA_SPREADSHEET_ID';

     // Initialize once
     FinanceLib.initLibrary(targetId);

     // Add multiple transactions
     FinanceLib.addIncome({
       date: new Date(),
       amount: 10000000,
       source: 'Test Income 1',
       category: 'Lương'
     });

     FinanceLib.addExpense({
       date: new Date(),
       amount: 500000,
       item: 'Test Expense 1',
       category: 'Ăn uống'
     });

     FinanceLib.addIncome({
       date: new Date(),
       amount: 2000000,
       source: 'Test Income 2',
       category: 'Thu nhập phụ'
     });

     Logger.log('✅ All operations completed');
     Logger.log('Check Target Data spreadsheet for new rows');
   }
   ```

2. **Verify in Target Data spreadsheet:**
   - [ ] THU sheet: 2 new income rows
   - [ ] CHI sheet: 1 new expense row
   - [ ] Consumer spreadsheet: No data written

**Expected:** ✅ All data in Target Data spreadsheet only

---

## Test Scenario 9: Automated Test Suite

**Goal:** Run comprehensive automated tests

### Steps:

1. **In Target Data spreadsheet Apps Script:**
   - Copy TestLibraryContext.gs content

2. **Run all tests:**
   ```javascript
   runAllTests()
   ```

3. **Review Execution Log:**
   ```
   View → Logs
   ```

**Expected Output:**
```
╔═══════════════════════════════════════════════════════════════╗
║                        TEST SUMMARY                           ║
╚═══════════════════════════════════════════════════════════════╝

Total Tests: 6
Passed: 6 ✅
Failed: 0 ❌
Success Rate: 100.0%

🎉 ALL TESTS PASSED! Library context fix is working correctly.
```

---

## Verification Checklist

### ✅ Functionality Tests

- [ ] Standalone mode works without initialization
- [ ] Library deploys successfully
- [ ] Target data spreadsheet setup completes
- [ ] Consumer spreadsheet can add library
- [ ] Library initialization succeeds with valid ID
- [ ] Data writes to Target Data spreadsheet (not Consumer)
- [ ] Invalid IDs rejected with clear errors
- [ ] Missing required sheets detected
- [ ] Status functions return correct info
- [ ] Reset returns to standalone mode
- [ ] Multiple operations work correctly
- [ ] Automated test suite passes

### ✅ Error Handling Tests

- [ ] Invalid spreadsheet ID handled gracefully
- [ ] Missing spreadsheet ID rejected
- [ ] Blank spreadsheet validation works
- [ ] Permission errors caught and reported
- [ ] Malformed IDs rejected

### ✅ State Management Tests

- [ ] Status before init: STANDALONE
- [ ] Status after init: LIBRARY
- [ ] Status after reset: STANDALONE
- [ ] Helper functions (isLibraryMode, getLibrarySpreadsheetId) work

---

## Troubleshooting

### Issue: "Library not found"
**Solution:** Re-add library with correct Script ID

### Issue: "Cannot access spreadsheet"
**Solution:** Verify spreadsheet ID, check sharing permissions

### Issue: "Missing required sheets"
**Solution:** Run Setup Wizard on Target Data spreadsheet

### Issue: Authorization prompt appears
**Solution:** Grant required permissions (normal for first run)

### Issue: Data appears in wrong spreadsheet
**Solution:** Verify initLibrary() was called with correct ID

---

## Success Criteria

**All tests pass if:**

1. ✅ Standalone mode: Data writes to current spreadsheet
2. ✅ Library mode: Data writes to target spreadsheet only
3. ✅ Error handling: Invalid inputs rejected with clear messages
4. ✅ State management: Init/reset/status work correctly
5. ✅ Automated tests: 100% pass rate

**Manual testing complete:** ________ (Tester signature)

**Date:** ________

**Result:** ☐ PASS  ☐ FAIL  ☐ PARTIAL

---

**End of Manual Testing Guide**
