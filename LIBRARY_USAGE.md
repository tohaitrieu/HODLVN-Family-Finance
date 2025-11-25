# Library Usage Guide - HODLVN-Family-Finance

**Version:** 3.1.0
**Last Updated:** 2025-11-25

---

## Quick Start

### Using as Standalone (Default)

No changes needed. Use as before:

```javascript
// In your spreadsheet
Menu: HODLVN Finance → Thu nhập
// Fill form and submit
```

### Using as Library

From external spreadsheet:

```javascript
// 1. Add library (Script ID: YOUR_SCRIPT_ID)
// 2. Initialize
const FinanceLib = YOUR_LIBRARY_NAME;
FinanceLib.initLibrary('YOUR_DATA_SPREADSHEET_ID');

// 3. Use functions
FinanceLib.addIncome({ date: new Date(), amount: 1000000, source: 'Salary', category: 'Lương' });
```

---

## Deployment

### 1. Deploy as Library

```
Apps Script Editor:
Deploy → New deployment
Type: Library
Version: New version
Description: HODLVN-Family-Finance v3.1.0
Deploy
```

Copy Script ID: `1A2B3C4D5E6F7G8H9I0J...`

### 2. Setup Target Data Spreadsheet

```
1. Create new Google Sheet
2. Extensions → Apps Script → Copy all .gs and .html files
3. Run Setup Wizard in spreadsheet
4. Copy Spreadsheet ID from URL
```

### 3. Add Library to Consumer

```
Consumer Spreadsheet:
Extensions → Apps Script
Libraries → + → Paste Script ID
Identifier: FinanceLib
Add
```

---

## API Reference

### Initialization

#### `initLibrary(spreadsheetId)`

Initialize library for external use. **MUST** be called before any operations.

**Parameters:**
- `spreadsheetId` (string) - Target data spreadsheet ID

**Returns:**
```javascript
{
  success: true/false,
  message: string,
  spreadsheetId: string,  // (if success)
  spreadsheetName: string, // (if success)
  mode: 'LIBRARY'
}
```

**Example:**
```javascript
const result = FinanceLib.initLibrary('1a2b3c4d5e6f7g8h9i0j');
if (!result.success) {
  console.error(result.message);
}
```

---

### Status Management

#### `getLibraryStatus()`

Get current library configuration.

**Returns:**
```javascript
{
  mode: 'LIBRARY' | 'STANDALONE',
  initialized: true/false,
  spreadsheetId: string,
  spreadsheetName: string,  // (if library mode)
  spreadsheetUrl: string    // (if library mode)
}
```

**Example:**
```javascript
const status = FinanceLib.getLibraryStatus();
console.log('Mode:', status.mode);
```

---

#### `resetLibrary()`

Reset to standalone mode.

**Returns:**
```javascript
{
  success: true,
  message: string,
  mode: 'STANDALONE',
  previousSpreadsheetId: string
}
```

**Example:**
```javascript
FinanceLib.resetLibrary();
```

---

#### `isLibraryMode()`

Check if library mode active.

**Returns:** `boolean`

**Example:**
```javascript
if (FinanceLib.isLibraryMode()) {
  console.log('Library mode active');
}
```

---

#### `getLibrarySpreadsheetId()`

Get configured spreadsheet ID.

**Returns:** `string | null`

**Example:**
```javascript
const id = FinanceLib.getLibrarySpreadsheetId();
```

---

### Data Operations

All existing functions work through library:

```javascript
// Income
FinanceLib.addIncome({ date, amount, source, category, notes });

// Expense
FinanceLib.addExpense({ date, amount, item, category, notes });

// Debt Management
FinanceLib.addDebt({ date, amount, type, notes });
FinanceLib.addDebtPayment({ date, amount, debtId, notes });

// Investments
FinanceLib.addStockTransaction({ date, action, symbol, quantity, price });
FinanceLib.addGoldTransaction({ date, action, quantity, price });
FinanceLib.addCryptoTransaction({ date, action, symbol, quantity, price });
FinanceLib.addOtherInvestment({ date, type, amount, notes });

// Dividends
FinanceLib.addDividend({ date, symbol, amount, notes });

// Budget
FinanceLib.setBudget({ category, amount, period });
```

---

## Error Handling

### Common Errors

**Error:** "spreadsheetId is required"
```javascript
// Wrong:
FinanceLib.initLibrary();

// Correct:
FinanceLib.initLibrary('1a2b3c4d...');
```

**Error:** "Failed to access spreadsheet"
```javascript
// Causes:
// 1. Invalid spreadsheet ID
// 2. No access permission
// 3. Spreadsheet deleted

// Solution: Verify ID and permissions
```

**Error:** "Missing required sheets"
```javascript
// Cause: Spreadsheet not initialized

// Solution: Run Setup Wizard in target spreadsheet
```

---

## Best Practices

### 1. Initialize Once

```javascript
// Good:
function myFinanceApp() {
  FinanceLib.initLibrary(SPREADSHEET_ID);
  FinanceLib.addIncome({...});
  FinanceLib.addExpense({...});
}

// Avoid:
function badExample() {
  FinanceLib.initLibrary(SPREADSHEET_ID);
  FinanceLib.addIncome({...});
  FinanceLib.initLibrary(SPREADSHEET_ID);  // Unnecessary
  FinanceLib.addExpense({...});
}
```

### 2. Check Initialization

```javascript
function safeOperation() {
  const result = FinanceLib.initLibrary(SPREADSHEET_ID);

  if (!result.success) {
    Logger.log('Init failed: ' + result.message);
    return;
  }

  // Proceed with operations
  FinanceLib.addIncome({...});
}
```

### 3. Use Status for Debugging

```javascript
function debugLibrary() {
  const status = FinanceLib.getLibraryStatus();
  Logger.log('Mode: ' + status.mode);
  Logger.log('Spreadsheet: ' + status.spreadsheetName);
  Logger.log('Initialized: ' + status.initialized);
}
```

---

## Migration Guide

### From Standalone to Library

**Before (Standalone):**
```javascript
// In your spreadsheet
addIncome({ date: new Date(), amount: 1000000, source: 'Salary', category: 'Lương' });
```

**After (Library):**
```javascript
// In external spreadsheet
const FinanceLib = YOUR_LIBRARY_NAME;
FinanceLib.initLibrary('YOUR_DATA_SPREADSHEET_ID');
FinanceLib.addIncome({ date: new Date(), amount: 1000000, source: 'Salary', category: 'Lương' });
```

**Key Changes:**
1. Add library to project
2. Call `initLibrary()` before operations
3. Prefix all functions with library name

---

## Troubleshooting

### Library Not Found

**Symptoms:** "FinanceLib is not defined"

**Solutions:**
1. Check library added correctly (Script ID)
2. Verify identifier matches usage (e.g., "FinanceLib")
3. Refresh Apps Script editor

---

### Data Written to Wrong Spreadsheet

**Symptoms:** Data appears in consumer spreadsheet

**Solutions:**
1. Verify `initLibrary()` called before operations
2. Check correct spreadsheet ID used
3. Use `getLibraryStatus()` to debug

---

### Permission Errors

**Symptoms:** "Cannot access spreadsheet"

**Solutions:**
1. Verify you own or have edit access to target spreadsheet
2. Re-authorize Apps Script permissions
3. Check spreadsheet not deleted

---

## Examples

### Example 1: Simple Income Tracking

```javascript
function trackIncome() {
  const DATA_SPREADSHEET_ID = '1a2b3c4d5e6f7g8h9i0j';

  // Initialize
  const result = FinanceLib.initLibrary(DATA_SPREADSHEET_ID);
  if (!result.success) {
    Logger.log('Error: ' + result.message);
    return;
  }

  // Add income
  FinanceLib.addIncome({
    date: new Date(),
    amount: 15000000,
    source: 'Monthly Salary',
    category: 'Lương',
    notes: 'November salary'
  });

  Logger.log('✅ Income tracked successfully');
}
```

---

### Example 2: Multiple Operations

```javascript
function monthlyFinances() {
  const DATA_SPREADSHEET_ID = '1a2b3c4d5e6f7g8h9i0j';

  // Initialize once
  FinanceLib.initLibrary(DATA_SPREADSHEET_ID);

  // Record income
  FinanceLib.addIncome({
    date: new Date(),
    amount: 20000000,
    source: 'Salary',
    category: 'Lương'
  });

  // Record expenses
  FinanceLib.addExpense({
    date: new Date(),
    amount: 500000,
    item: 'Groceries',
    category: 'Ăn uống'
  });

  FinanceLib.addExpense({
    date: new Date(),
    amount: 2000000,
    item: 'Electricity bill',
    category: 'Hóa đơn'
  });

  // Check status
  const status = FinanceLib.getLibraryStatus();
  Logger.log('Operations completed on: ' + status.spreadsheetName);
}
```

---

### Example 3: Error Handling

```javascript
function robustFinanceTracking() {
  const DATA_SPREADSHEET_ID = '1a2b3c4d5e6f7g8h9i0j';

  try {
    // Initialize with validation
    const initResult = FinanceLib.initLibrary(DATA_SPREADSHEET_ID);

    if (!initResult.success) {
      throw new Error('Initialization failed: ' + initResult.message);
    }

    Logger.log('✅ Initialized: ' + initResult.spreadsheetName);

    // Perform operations
    FinanceLib.addIncome({
      date: new Date(),
      amount: 10000000,
      source: 'Freelance work',
      category: 'Thu nhập phụ'
    });

    // Verify
    const status = FinanceLib.getLibraryStatus();
    if (status.mode !== 'LIBRARY') {
      throw new Error('Not in library mode!');
    }

    Logger.log('✅ All operations completed successfully');

  } catch (error) {
    Logger.log('❌ Error: ' + error.message);
    Logger.log('Status: ' + JSON.stringify(FinanceLib.getLibraryStatus()));
  }
}
```

---

## Testing

### Quick Test

```javascript
function testLibrarySetup() {
  const DATA_SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID_HERE';

  // Test 1: Initialize
  Logger.log('=== Test 1: Initialize ===');
  const result = FinanceLib.initLibrary(DATA_SPREADSHEET_ID);
  Logger.log(JSON.stringify(result, null, 2));

  // Test 2: Check status
  Logger.log('=== Test 2: Status ===');
  const status = FinanceLib.getLibraryStatus();
  Logger.log(JSON.stringify(status, null, 2));

  // Test 3: Add test income
  Logger.log('=== Test 3: Add Income ===');
  FinanceLib.addIncome({
    date: new Date(),
    amount: 1000000,
    source: 'Test',
    category: 'Lương'
  });
  Logger.log('✅ Income added');

  // Test 4: Reset
  Logger.log('=== Test 4: Reset ===');
  const resetResult = FinanceLib.resetLibrary();
  Logger.log(JSON.stringify(resetResult, null, 2));
}
```

---

## Support

**Documentation:** [README.md](README.md)
**Technical Docs:** [docs/TECHNICAL_DOCUMENTATION.md](docs/TECHNICAL_DOCUMENTATION.md)
**Issues:** [GitHub Issues](https://github.com/tohaitrieu/HODLVN-Family-Finance/issues)

---

**Last Updated:** 2025-11-25
**Version:** 3.1.0
