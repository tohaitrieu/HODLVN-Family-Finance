# 📦 Library Deployment Guide

**HODLVN-Family-Finance v3.1.0+**

Hướng dẫn triển khai và sử dụng HODLVN-Family-Finance như một Google Apps Script Library.

---

## 📋 MỤC LỤC

1. [Giới thiệu](#-giới-thiệu)
2. [Tại sao dùng Library Mode?](#-tại-sao-dùng-library-mode)
3. [Cài đặt Library](#-cài-đặt-library)
4. [Sử dụng từ External Spreadsheet](#-sử-dụng-từ-external-spreadsheet)
5. [API Reference](#-api-reference)
6. [Use Cases](#-use-cases)
7. [Troubleshooting](#-troubleshooting)
8. [Best Practices](#-best-practices)

---

## 🎯 Giới thiệu

Từ version 3.1.0, HODLVN-Family-Finance hỗ trợ **Library Mode** - cho phép code được triển khai một lần và sử dụng từ nhiều spreadsheet khác nhau.

### Điểm khác biệt giữa 2 modes

| Feature | Standalone Mode | Library Mode |
|---------|----------------|--------------|
| **Code location** | Trong spreadsheet | Trong Library project riêng |
| **Số spreadsheet** | 1 (code + data cùng file) | Nhiều (1 library → N spreadsheets) |
| **Update code** | Phải copy vào từng file | Update 1 lần, áp dụng cho tất cả |
| **Use case** | Cá nhân, 1 file dữ liệu | Team, nhiều file, quản lý tập trung |

---

## 🤔 Tại sao dùng Library Mode?

### ✅ Ưu điểm

1. **Quản lý nhiều spreadsheet**
   - Một codebase cho nhiều file dữ liệu khác nhau
   - Mỗi thành viên team có spreadsheet riêng
   - Dễ dàng scale từ 1 → 10 → 100 users

2. **Cập nhật dễ dàng**
   - Update code một lần trong library
   - Tất cả spreadsheet tự động dùng version mới
   - Không cần copy-paste code vào từng file

3. **Tách biệt code & data**
   - Code trong library (read-only for users)
   - Data trong spreadsheet cá nhân
   - Bảo mật code tốt hơn

4. **Version control**
   - Quản lý version library dễ dàng
   - Users có thể chọn version library
   - Rollback khi cần

### ❌ Nhược điểm

1. **Phức tạp hơn**
   - Setup phức tạp hơn standalone
   - Cần hiểu cách deploy library

2. **Permissions**
   - Users cần có quyền truy cập library
   - Phải manage library sharing

3. **Debugging**
   - Debug khó hơn (code không ở trong file)
   - Log messages phải xem ở library project

### 🎯 Khi nào nên dùng Library Mode?

**Dùng Library Mode khi:**
- ✅ Có 2+ spreadsheet cần dùng cùng code
- ✅ Có team nhiều người dùng
- ✅ Cần update code thường xuyên
- ✅ Muốn bảo mật code (users không xem được)

**Dùng Standalone Mode khi:**
- ✅ Chỉ có 1 spreadsheet
- ✅ Sử dụng cá nhân
- ✅ Muốn đơn giản, dễ setup
- ✅ Muốn customize code thoải mái

---

## 🚀 Cài đặt Library

### Bước 1: Tạo Library Project

```
1. Mở Google Apps Script Editor: https://script.google.com
2. New Project → Đặt tên: "HODLVN-Finance-Library"
3. Copy tất cả .gs files từ repo vào project:
   - Main.gs
   - Utils.gs
   - DataProcessor.gs
   - BudgetManager.gs
   - DashboardManager.gs
   - SheetInitializer.gs
   - LibraryConfig.gs (QUAN TRỌNG!)
   - Và tất cả files khác
4. Copy tất cả .html forms vào project:
   - IncomeForm.html
   - ExpenseForm.html
   - Và tất cả forms khác
```

### Bước 2: Deploy as Library

```
1. Trong Apps Script Editor:
   - Click "Deploy" → "New deployment"
   - Chọn type: "Library"
   - Description: "HODLVN-Family-Finance Library v3.1.0"
   - Click "Deploy"

2. Copy thông tin quan trọng:
   - Script ID: 1a2b3c4d5e6f7g8h9i0j... (copy ID này)
   - Deployment ID (nếu có)

3. Cấu hình access:
   - Settings → Share
   - Chọn: "Anyone with the link can view"
   - Hoặc: Chia sẻ với specific users/groups
```

### Bước 3: Get Library Script ID

```
Library Script ID (cần cho bước tiếp theo):
- Trong Library Project
- Settings (gear icon) → Script ID
- Copy Script ID: 1a2b3c4d5e6f7g8h9i0j...
```

---

## 💻 Sử dụng từ External Spreadsheet

### Bước 1: Tạo Data Spreadsheet

```
1. Tạo Google Sheet mới (hoặc dùng existing)
2. Chạy Setup Wizard để khởi tạo sheets
   Option A: Copy code tạm vào spreadsheet, chạy Setup Wizard
   Option B: Tạo sheets thủ công (THU, CHI, BUDGET, DASHBOARD, etc.)
3. Lưu Spreadsheet ID:
   - Lấy từ URL: https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/edit
   - Copy SPREADSHEET_ID này
```

### Bước 2: Add Library Reference

```
1. Mở spreadsheet → Extensions → Apps Script
2. Editor bên trái → Libraries (dấu +)
3. Paste Script ID: 1a2b3c4d5e6f7g8h9i0j...
4. Chọn version (hoặc HEAD for latest)
5. Identifier: "FinanceLib" (hoặc tên bạn muốn)
6. Click "Add"
```

### Bước 3: Initialize Library

```javascript
// Trong Apps Script của spreadsheet này:

function setupLibrary() {
  // Get spreadsheet ID of THIS spreadsheet
  const spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();

  // Initialize library
  const result = FinanceLib.initLibrary(spreadsheetId);

  // Check result
  if (result.success) {
    Logger.log('✅ ' + result.message);
    Logger.log('Spreadsheet: ' + result.spreadsheetName);
  } else {
    Logger.log('❌ ' + result.message);
  }
}
```

### Bước 4: Sử dụng Library Functions

```javascript
// Add income
function testAddIncome() {
  const result = FinanceLib.addIncome({
    date: '2025-11-25',
    amount: 10000000,
    source: 'Lương tháng 11',
    category: 'Lương',
    note: 'Test from library'
  });

  Logger.log(result.success ? '✅ Added' : '❌ Error');
}

// Add expense
function testAddExpense() {
  const result = FinanceLib.addExpense({
    date: '2025-11-25',
    amount: 500000,
    category: 'Ăn uống',
    subcategory: 'Ăn sáng',
    paymentMethod: 'Tiền mặt',
    note: 'Breakfast'
  });

  Logger.log(result.success ? '✅ Added' : '❌ Error');
}

// Check library status
function checkStatus() {
  const status = FinanceLib.getLibraryStatus();
  Logger.log('Mode: ' + status.mode);
  Logger.log('Spreadsheet: ' + status.spreadsheetName);
  Logger.log('ID: ' + status.spreadsheetId);
}
```

---

## 📖 API Reference

### LibraryConfig Functions

#### `initLibrary(spreadsheetId)`

Khởi tạo library với target spreadsheet.

**Parameters:**
- `spreadsheetId` (string): ID của spreadsheet chứa data

**Returns:**
```javascript
{
  success: true/false,
  message: 'Success or error message',
  spreadsheetId: 'abc123...',
  spreadsheetName: 'My Finance Data',
  mode: 'LIBRARY'
}
```

**Example:**
```javascript
const result = FinanceLib.initLibrary('1a2b3c4d5e6f7g8h9i0j');
if (!result.success) {
  throw new Error(result.message);
}
```

#### `getLibraryStatus()`

Kiểm tra trạng thái hiện tại của library.

**Returns:**
```javascript
{
  mode: 'LIBRARY' or 'STANDALONE',
  initialized: true/false,
  spreadsheetId: 'abc123...',
  spreadsheetName: 'My Finance Data',
  spreadsheetUrl: 'https://...'
}
```

**Example:**
```javascript
const status = FinanceLib.getLibraryStatus();
Logger.log('Current mode: ' + status.mode);
```

#### `resetLibrary()`

Reset library về standalone mode.

**Returns:**
```javascript
{
  success: true,
  message: 'Reset message',
  mode: 'STANDALONE',
  previousSpreadsheetId: 'abc123...'
}
```

**Example:**
```javascript
const result = FinanceLib.resetLibrary();
Logger.log(result.message);
```

### All Data Functions

Tất cả functions từ standalone mode đều hoạt động trong library mode:

```javascript
// Income
FinanceLib.addIncome(data)

// Expense
FinanceLib.addExpense(data)

// Debt
FinanceLib.addDebt(data)
FinanceLib.payDebt(data)

// Lending
FinanceLib.addLending(data)
FinanceLib.collectLending(data)

// Investment
FinanceLib.buyStock(data)
FinanceLib.sellStock(data)
FinanceLib.buyGold(data)
FinanceLib.sellGold(data)
FinanceLib.buyCrypto(data)
FinanceLib.sellCrypto(data)
FinanceLib.addOtherInvestment(data)

// Dividend
FinanceLib.addDividend(data)

// Budget
FinanceLib.setBudget(data)
FinanceLib.checkBudget()

// Dashboard
FinanceLib.refreshDashboard()
```

---

## 📋 Use Cases

### Use Case 1: Family Finance (Multiple Members)

**Scenario:** Gia đình 4 người, mỗi người 1 spreadsheet riêng.

**Setup:**
```
1. Deploy 1 library chung
2. Mỗi người tạo spreadsheet riêng:
   - Dad: "Finance - Dad"
   - Mom: "Finance - Mom"
   - Son: "Finance - Son"
   - Daughter: "Finance - Daughter"
3. Mỗi spreadsheet add library reference
4. Mỗi người initialize với spreadsheet ID của mình
5. Mỗi người track tài chính độc lập
```

**Benefits:**
- Code chung, dữ liệu riêng
- Dad update code → tất cả dùng version mới
- Privacy: Mỗi người không xem được data người khác

### Use Case 2: Business (Multiple Businesses)

**Scenario:** Quản lý 5 doanh nghiệp nhỏ.

**Setup:**
```
1. Deploy 1 library
2. Tạo 5 spreadsheets:
   - "Business A Finance"
   - "Business B Finance"
   - "Business C Finance"
   - "Business D Finance"
   - "Business E Finance"
3. Mỗi spreadsheet init với library
4. Tách biệt sổ sách từng doanh nghiệp
```

**Benefits:**
- Không lẫn lộn data giữa các DN
- Báo cáo riêng từng DN
- Quản lý tập trung

### Use Case 3: Template Distribution

**Scenario:** Chia sẻ template cho nhiều người dùng.

**Setup:**
```
1. Deploy library public
2. Tạo template spreadsheet (empty data)
3. User make a copy template
4. User initialize với spreadsheet ID của copy
5. User bắt đầu sử dụng
```

**Benefits:**
- User không cần copy code
- User chỉ cần template spreadsheet
- Easy distribution

---

## 🔧 Troubleshooting

### Error: "spreadsheetId is required"

**Cause:** Chưa gọi `initLibrary()`.

**Fix:**
```javascript
// Must call initLibrary first
const spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
FinanceLib.initLibrary(spreadsheetId);
```

### Error: "Missing required sheets"

**Cause:** Spreadsheet chưa có sheets cần thiết (THU, CHI, BUDGET, DASHBOARD).

**Fix:**
```javascript
// Run Setup Wizard or create sheets manually
FinanceLib.initializeAllSheets();
```

### Error: "Failed to access spreadsheet"

**Cause:**
- Spreadsheet ID sai
- Không có quyền truy cập
- Spreadsheet đã bị xóa

**Fix:**
```javascript
// Check spreadsheet ID is correct
// Check you have Editor access
// Check spreadsheet still exists
const ss = SpreadsheetApp.openById('YOUR_ID');
Logger.log(ss.getName()); // Should work
```

### Error: "ReferenceError: FinanceLib is not defined"

**Cause:** Chưa add library reference vào project.

**Fix:**
```
1. Apps Script Editor → Libraries (dấu +)
2. Paste Script ID
3. Add library
4. Identifier must match: "FinanceLib"
```

### Warning: "Data writes not working"

**Cause:** Quên gọi `initLibrary()` trước khi call data functions.

**Fix:**
```javascript
// CORRECT order:
FinanceLib.initLibrary(spreadsheetId);  // 1. Init first
FinanceLib.addIncome(data);             // 2. Then use

// WRONG order:
FinanceLib.addIncome(data);             // ❌ Will fail
FinanceLib.initLibrary(spreadsheetId);  // Too late
```

---

## ✅ Best Practices

### 1. Always Initialize First

```javascript
// GOOD: Create setup function
function setupMySpreadsheet() {
  const spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  const result = FinanceLib.initLibrary(spreadsheetId);

  if (!result.success) {
    throw new Error('Init failed: ' + result.message);
  }

  return result;
}

// BAD: Call functions without init
function badExample() {
  FinanceLib.addIncome(data); // ❌ Will fail
}
```

### 2. Check Status Before Operations

```javascript
function safeAddIncome(data) {
  // Check if library is initialized
  const status = FinanceLib.getLibraryStatus();

  if (!status.initialized) {
    // Initialize if needed
    const spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
    FinanceLib.initLibrary(spreadsheetId);
  }

  // Now safe to add income
  return FinanceLib.addIncome(data);
}
```

### 3. Handle Errors Gracefully

```javascript
function robustAddIncome(data) {
  try {
    const result = FinanceLib.addIncome(data);

    if (result.success) {
      Logger.log('✅ Income added: ' + result.message);
      return result;
    } else {
      Logger.log('❌ Failed: ' + result.message);
      // Handle error
      return null;
    }

  } catch (error) {
    Logger.log('❌ Exception: ' + error.message);
    // Maybe library not initialized?
    const status = FinanceLib.getLibraryStatus();
    Logger.log('Library status: ' + JSON.stringify(status));
    throw error;
  }
}
```

### 4. Version Pinning

```
When adding library:
- For production: Pin to specific version (v1, v2, etc.)
- For development: Use HEAD (latest)

Update strategy:
1. Test new version in dev spreadsheet
2. If OK, update production spreadsheets
3. Pin to stable version
```

### 5. Document Your Setup

```javascript
/**
 * HODLVN-Family-Finance Library Setup
 *
 * Library: HODLVN-Finance-Library
 * Script ID: 1a2b3c4d5e6f7g8h9i0j
 * Version: 1 (v3.1.0)
 * Identifier: FinanceLib
 *
 * Initialized: 2025-11-25
 * Spreadsheet: My Personal Finance
 * Spreadsheet ID: abc123xyz789
 */

function initLibrary() {
  const spreadsheetId = 'abc123xyz789';
  return FinanceLib.initLibrary(spreadsheetId);
}
```

---

## 🆚 Library Mode vs Standalone Mode

### Decision Matrix

| Requirement | Standalone | Library |
|-------------|-----------|---------|
| **Single user** | ✅ Best | ⚠️ Overkill |
| **Multiple users** | ❌ Hard | ✅ Best |
| **Easy setup** | ✅ Yes | ⚠️ Complex |
| **Easy update** | ❌ Manual | ✅ Auto |
| **Code security** | ❌ Visible | ✅ Hidden |
| **Customization** | ✅ Easy | ⚠️ Limited |
| **Debugging** | ✅ Easy | ⚠️ Harder |
| **Version control** | ❌ Hard | ✅ Easy |

### Migration Path

**Standalone → Library:**
```
1. Deploy code as library
2. Create new empty spreadsheet
3. Copy data from standalone to new spreadsheet
4. Initialize library in new spreadsheet
5. Test all functions
6. Switch to new spreadsheet
```

**Library → Standalone:**
```
1. Copy library code to spreadsheet's Apps Script
2. Remove library reference
3. Remove initLibrary() calls
4. Functions auto-switch to standalone mode
5. Test all functions
```

---

## 📞 Support

### Getting Help

- 📖 **Docs**: [README.md](../README.md)
- 🐛 **Issues**: [GitHub Issues](https://github.com/tohaitrieu/HODLVN-Family-Finance/issues)
- 💬 **Community**: [Facebook Group](https://facebook.com/groups/hodl.vn)
- 📧 **Email**: contact@tohaitrieu.net

### Reporting Library Issues

When reporting issues, include:

```
1. Library version: v3.1.0
2. Library mode: LIBRARY or STANDALONE
3. Spreadsheet ID: abc123... (if comfortable sharing)
4. Error message: Full error text
5. Steps to reproduce: Detailed steps
6. Expected behavior: What should happen
7. Actual behavior: What actually happened
```

---

## 🔄 Updates & Changelog

### Version 3.1.0 (2025-11-25)
- ✨ Initial library mode support
- ✨ LibraryConfig.gs module
- ✨ Complete routing coverage (50+ fixes)
- ✅ 10/10 integration tests passing

### Future Roadmap

**v3.2.0:**
- [ ] Library version auto-check
- [ ] Migration tools (standalone ↔ library)
- [ ] Better error messages for library mode
- [ ] Library usage analytics

**v3.3.0:**
- [ ] Multi-library support (modular libraries)
- [ ] Library marketplace
- [ ] Community-contributed libraries

---

## 📄 License

MIT License - See [LICENSE](../LICENSE) for details.

---

<div align="center">

**Library Mode được phát triển với ❤️ cho cộng đồng Việt Nam**

[⬆ Về đầu trang](#-library-deployment-guide)

</div>
