# 🔧 TÀI LIỆU KỸ THUẬT - HODLVN-Family-Finance

Tài liệu kỹ thuật chi tiết cho developers và contributors.

---

## 📑 MỤC LỤC

1. [Tổng quan kiến trúc](#-tổng-quan-kiến-trúc)
2. [Kiến trúc v3.0 - Modular](#-kiến-trúc-v30---modular)
3. [Modules chi tiết](#-modules-chi-tiết)
4. [Cấu trúc dữ liệu](#-cấu-trúc-dữ-liệu)
5. [API Reference](#-api-reference)
6. [Data Flow](#-data-flow)
7. [Database Schema](#-database-schema)
8. [Formulas & Calculations](#-formulas--calculations)
9. [Code Conventions](#-code-conventions)
10. [Testing Guidelines](#-testing-guidelines)
11. [Deployment](#-deployment)
12. [Performance Optimization](#-performance-optimization)
13. [Security](#-security)
14. [Error Handling](#-error-handling)
15. [Debugging](#-debugging)
16. [Contributing](#-contributing)

---

## 🏗️ Tổng quan kiến trúc

### Technology Stack

```
Platform:      Google Apps Script (JavaScript ES6)
Storage:       Google Sheets
UI:            HTML/CSS/JavaScript
Framework:     None (Vanilla JS)
Version:       3.0 (Modular Architecture)
```

### Design Principles

1. **Modularity**: Tách biệt chức năng thành modules độc lập
2. **Separation of Concerns**: UI, Logic, Data tách riêng
3. **DRY (Don't Repeat Yourself)**: Tái sử dụng code qua Utils
4. **Single Responsibility**: Mỗi function làm 1 việc duy nhất
5. **KISS (Keep It Simple, Stupid)**: Code đơn giản, dễ hiểu

---

## 🎯 Kiến trúc v3.0 - Modular

### Sơ đồ tổng quan

```
┌─────────────────────────────────────────────────────────────┐
│                    GOOGLE SHEETS UI                          │
│         (10 Data Sheets + 1 Dashboard + 1 Config)           │
└────────────────────┬────────────────────────────────────────┘
                     │
┌────────────────────▼────────────────────────────────────────┐
│                 MAIN MENU (Main.gs)                          │
│            Menu Dispatcher & Event Router                    │
└─┬──────┬──────────┬──────────┬──────────┬──────────────────┘
  │      │          │          │          │
┌─▼────┐│┌─▼──────┐│┌─▼──────┐│┌─▼──────┐│┌─▼──────┐
│Sheet ││ │ Data   ││ │Budget ││ │Dashbrd││ │ Utils │
│Init  ││ │Process ││ │Manager││ │Manager││ │       │
└──────┘│└────────┘│└────────┘│└────────┘│└────────┘
        │          │          │          │
        └──────────┴──────────┴──────────┘
               Shared Utilities
```

### So sánh v2.0 vs v3.0

| Aspect | v2.0 (Monolithic) | v3.0 (Modular) |
|--------|-------------------|----------------|
| **Cấu trúc** | 1 file Code.gs lớn | 6 modules riêng biệt |
| **Maintainability** | Khó bảo trì | Dễ bảo trì |
| **Testing** | Khó test | Dễ test từng module |
| **Debugging** | Phức tạp | Đơn giản hơn |
| **Scalability** | Giới hạn | Dễ mở rộng |
| **Collaboration** | Conflict nhiều | Conflict ít |
| **Lines of Code** | ~3000 lines | ~500 lines/module |

### Lợi ích của Modular Architecture

```
✅ Code rõ ràng, dễ đọc
✅ Dễ debug (biết lỗi ở module nào)
✅ Dễ test (test từng module riêng)
✅ Dễ maintain (sửa 1 module không ảnh hưởng khác)
✅ Dễ scale (thêm module mới không ảnh hưởng cũ)
✅ Team collaboration tốt hơn
✅ Reusability cao
```

---

## 📦 Modules chi tiết

### 1. Main.gs - Menu Dispatcher

**Mục đích:** Tạo menu và điều hướng

**Functions chính:**
```javascript
// Entry point - Tự động chạy khi mở file
function onOpen() {
  createMenu();
}

// Tạo menu HODLVN Finance
function createMenu() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('HODLVN Finance')
    .addItem('🚀 Setup Wizard', 'openSetupWizard')
    .addSeparator()
    .addItem('📥 Thu (THU)', 'openIncomeForm')
    .addItem('📤 Chi (CHI)', 'openExpenseForm')
    .addSeparator()
    .addItem('💳 Quản lý nợ', 'openDebtManagementForm')
    .addItem('💰 Trả nợ', 'openDebtPaymentForm')
    .addSeparator()
    .addItem('📈 Chứng khoán', 'openStockForm')
    .addItem('🪙 Vàng', 'openGoldForm')
    .addItem('₿ Crypto', 'openCryptoForm')
    .addItem('💼 Đầu tư khác', 'openOtherInvestmentForm')
    .addItem('💵 Cổ tức', 'openDividendForm')
    .addSeparator()
    .addItem('🎯 Đặt ngân sách', 'openSetBudgetForm')
    .addToUi();
}

// Form openers
function openIncomeForm() {
  showForm('IncomeForm');
}

function openExpenseForm() {
  showForm('ExpenseForm');
}

// Generic form opener
function showForm(formName) {
  var html = HtmlService.createHtmlOutputFromFile(formName)
    .setWidth(600)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, formName);
}
```

**Dependency:** Không phụ thuộc module khác

---

### 2. SheetInitializer.gs - Sheet Setup

**Mục đích:** Khởi tạo và cấu hình sheets

**Functions chính:**

```javascript
/**
 * Khởi tạo tất cả sheets
 */
function initializeAllSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Tạo 10 data sheets + Dashboard + Config
  createSheetIfNotExists(ss, 'Dashboard');
  createSheetIfNotExists(ss, 'THU');
  createSheetIfNotExists(ss, 'CHI');
  createSheetIfNotExists(ss, 'QUẢN LÝ NỢ');
  createSheetIfNotExists(ss, 'TRẢ NỢ');
  createSheetIfNotExists(ss, 'CHỨNG KHOÁN');
  createSheetIfNotExists(ss, 'VÀNG');
  createSheetIfNotExists(ss, 'CRYPTO');
  createSheetIfNotExists(ss, 'ĐẦU TƯ KHÁC');
  createSheetIfNotExists(ss, 'CỔ TỨC');
  createSheetIfNotExists(ss, 'BUDGET');
  createSheetIfNotExists(ss, 'CONFIG');
  
  // Setup từng sheet
  setupDashboard();
  setupIncomeSheet();
  setupExpenseSheet();
  setupDebtManagementSheet();
  setupDebtPaymentSheet();
  setupStockSheet();
  setupGoldSheet();
  setupCryptoSheet();
  setupOtherInvestmentSheet();
  setupDividendSheet();
  setupBudgetSheet();
  setupConfigSheet();
}

/**
 * Tạo sheet nếu chưa tồn tại
 */
function createSheetIfNotExists(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    Logger.log('Created sheet: ' + sheetName);
  }
  return sheet;
}

/**
 * Setup sheet THU
 */
function setupIncomeSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName('THU');
  
  // Headers
  var headers = ['STT', 'Ngày', 'Danh mục', 'Số tiền', 'Ghi chú'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Formatting
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#34A853')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  // Freeze header row
  sheet.setFrozenRows(1);
  
  // Column widths
  sheet.setColumnWidth(1, 60);  // STT
  sheet.setColumnWidth(2, 120); // Ngày
  sheet.setColumnWidth(3, 150); // Danh mục
  sheet.setColumnWidth(4, 150); // Số tiền
  sheet.setColumnWidth(5, 300); // Ghi chú
  
  // Pre-fill formulas (rows 2-1000)
  var formulas = [];
  for (var i = 2; i <= 1000; i++) {
    formulas.push(['=IF(B' + i + '<>"", ROW()-1, "")']); // STT auto
  }
  sheet.getRange(2, 1, 999, 1).setFormulas(formulas);
  
  // Data validation for Danh mục (dropdown)
  var categoryRange = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName('CONFIG')
    .getRange('A2:A10'); // Income categories
  
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(categoryRange)
    .build();
  sheet.getRange(2, 3, 999, 1).setDataValidation(rule);
  
  // Number formatting for Số tiền
  sheet.getRange(2, 4, 999, 1).setNumberFormat('#,##0');
}

/**
 * Setup Wizard - Thu thập thông tin người dùng
 */
function openSetupWizard() {
  var html = HtmlService.createHtmlOutputFromFile('SetupWizard')
    .setWidth(700)
    .setHeight(900);
  SpreadsheetApp.getUi().showModalDialog(html, '🚀 Setup Wizard');
}

/**
 * Xử lý data từ Setup Wizard
 */
function processSetupWizard(data) {
  try {
    // Lưu thông tin user vào Script Properties
    var props = PropertiesService.getScriptProperties();
    props.setProperty('USER_NAME', data.userName);
    props.setProperty('USER_EMAIL', data.userEmail);
    props.setProperty('USER_PHONE', data.userPhone);
    props.setProperty('START_YEAR', data.startYear);
    props.setProperty('MONTHLY_BUDGET', data.monthlyBudget || '0');
    props.setProperty('FINANCIAL_GOAL', data.financialGoal || '');
    props.setProperty('SETUP_COMPLETED', 'true');
    
    // Khởi tạo tất cả sheets
    initializeAllSheets();
    
    // Cập nhật Dashboard với thông tin user
    updateDashboardUserInfo(data);
    
    return {
      success: true,
      message: 'Khởi tạo thành công!'
    };
  } catch (error) {
    Logger.log('Error in processSetupWizard: ' + error);
    return {
      success: false,
      message: 'Lỗi: ' + error.message
    };
  }
}
```

**Dependency:** Utils.gs (helper functions)

---

### 3. DataProcessor.gs - Transaction Handler

**Mục đích:** Xử lý tất cả các loại giao dịch

**Functions chính:**

```javascript
/**
 * Thêm thu nhập
 */
function addIncome(data) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName('THU');
    
    // Tìm row trống (dựa vào cột Ngày, không phải STT)
    var emptyRow = findEmptyRow(sheet, 2); // Column B (Ngày)
    
    // Data array
    var rowData = [
      '', // STT - Formula tự tính
      data.date,
      data.category,
      data.amount,
      data.note || ''
    ];
    
    // Insert data
    sheet.getRange(emptyRow, 1, 1, rowData.length).setValues([rowData]);
    
    // Update Dashboard
    updateDashboard();
    
    Logger.log('Added income: ' + JSON.stringify(data));
    return { success: true };
  } catch (error) {
    Logger.log('Error in addIncome: ' + error);
    return { success: false, message: error.message };
  }
}

/**
 * Thêm chi tiêu
 */
function addExpense(data) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName('CHI');
    
    var emptyRow = findEmptyRow(sheet, 2);
    
    var rowData = [
      '',
      data.date,
      data.category,
      data.amount,
      data.paymentMethod || 'Tiền mặt',
      data.note || ''
    ];
    
    sheet.getRange(emptyRow, 1, 1, rowData.length).setValues([rowData]);
    
    // Update Budget
    updateBudgetCategory(data.category, data.amount);
    
    // Update Dashboard
    updateDashboard();
    
    Logger.log('Added expense: ' + JSON.stringify(data));
    return { success: true };
  } catch (error) {
    Logger.log('Error in addExpense: ' + error);
    return { success: false, message: error.message };
  }
}

/**
 * Thêm nợ - TỰ ĐỘNG TẠO THU NHẬP
 */
function addDebt(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. Thêm vào QUẢN LÝ NỢ
    var debtSheet = ss.getSheetByName('QUẢN LÝ NỢ');
    var emptyRow = findEmptyRow(debtSheet, 2);
    
    var debtData = [
      '',
      data.loanDate,
      data.debtType,
      data.amount,
      data.interestRate,
      data.term,
      data.dueDate,
      data.purpose || '',
      data.note || '',
      data.amount, // Số dư ban đầu = Số tiền vay
      'Chưa thanh toán' // Status
    ];
    
    debtSheet.getRange(emptyRow, 1, 1, debtData.length)
      .setValues([debtData]);
    
    // 2. TỰ ĐỘNG TẠO THU NHẬP
    var incomeData = {
      date: data.loanDate,
      category: 'Vay ' + data.debtType,
      amount: data.amount,
      note: 'Tự động từ khoản nợ: ' + data.purpose
    };
    
    addIncome(incomeData);
    
    Logger.log('Added debt and auto-created income: ' + 
               JSON.stringify(data));
    return { success: true };
  } catch (error) {
    Logger.log('Error in addDebt: ' + error);
    return { success: false, message: error.message };
  }
}

/**
 * Trả nợ
 */
function payDebt(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. Thêm vào TRẢ NỢ
    var paymentSheet = ss.getSheetByName('TRẢ NỢ');
    var emptyRow = findEmptyRow(paymentSheet, 2);
    
    var paymentData = [
      '',
      data.paymentDate,
      data.debtId,
      data.principalPayment,
      data.interestPayment,
      data.principalPayment + data.interestPayment, // Tổng
      data.note || ''
    ];
    
    paymentSheet.getRange(emptyRow, 1, 1, paymentData.length)
      .setValues([paymentData]);
    
    // 2. Cập nhật số dư trong QUẢN LÝ NỢ
    updateDebtBalance(data.debtId, data.principalPayment);
    
    // 3. Update Dashboard
    updateDashboard();
    
    Logger.log('Paid debt: ' + JSON.stringify(data));
    return { success: true };
  } catch (error) {
    Logger.log('Error in payDebt: ' + error);
    return { success: false, message: error.message };
  }
}

/**
 * Cập nhật số dư nợ
 */
function updateDebtBalance(debtId, principalPayment) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName('QUẢN LÝ NỢ');
  
  var data = sheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == debtId) { // Match STT
      var currentBalance = data[i][9]; // Column J - Số dư
      var newBalance = currentBalance - principalPayment;
      
      sheet.getRange(i + 1, 10).setValue(newBalance); // Update số dư
      
      // Nếu trả hết
      if (newBalance <= 0) {
        sheet.getRange(i + 1, 11).setValue('Đã thanh toán');
      }
      
      break;
    }
  }
}

/**
 * Thêm giao dịch chứng khoán
 */
function addStockTransaction(data) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName('CHỨNG KHOÁN');
    
    var emptyRow = findEmptyRow(sheet, 2);
    
    // Tính toán
    var totalValue = data.quantity * data.price;
    var fee = totalValue * (data.feeRate / 100);
    var tax = data.transactionType === 'Bán' ? 
              totalValue * (data.taxRate / 100) : 0;
    var totalCost = totalValue + fee + tax;
    
    var rowData = [
      '',
      data.date,
      data.transactionType,
      data.stockCode,
      data.quantity,
      data.price,
      totalValue,
      fee,
      tax,
      totalCost,
      data.useMargin ? 'Có' : 'Không',
      data.marginAmount || 0,
      data.marginRate || 0,
      data.note || ''
    ];
    
    sheet.getRange(emptyRow, 1, 1, rowData.length).setValues([rowData]);
    
    // Nếu dùng margin → Tạo khoản nợ
    if (data.useMargin && data.marginAmount > 0) {
      var marginDebt = {
        loanDate: data.date,
        debtType: 'Vay margin',
        amount: data.marginAmount,
        interestRate: data.marginRate,
        term: 3, // Default 3 tháng
        dueDate: calculateDueDate(data.date, 3),
        purpose: 'Margin mua ' + data.stockCode,
        note: 'Tự động từ giao dịch chứng khoán'
      };
      
      addDebt(marginDebt);
    }
    
    updateDashboard();
    
    Logger.log('Added stock transaction: ' + JSON.stringify(data));
    return { success: true };
  } catch (error) {
    Logger.log('Error in addStockTransaction: ' + error);
    return { success: false, message: error.message };
  }
}

/**
 * Thêm cổ tức
 */
function addDividend(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. Thêm vào CỔ TỨC
    var dividendSheet = ss.getSheetByName('CỔ TỨC');
    var emptyRow = findEmptyRow(dividendSheet, 2);
    
    var totalDividend = data.dividendPerShare * data.quantity;
    
    var rowData = [
      '',
      data.receiveDate,
      data.stockCode,
      data.dividendPerShare,
      data.quantity,
      totalDividend,
      data.dividendType,
      data.note || ''
    ];
    
    dividendSheet.getRange(emptyRow, 1, 1, rowData.length)
      .setValues([rowData]);
    
    // 2. Nếu là cổ tức tiền mặt → Tạo thu nhập
    if (data.dividendType === 'Tiền mặt' || 
        data.dividendType === 'Cả hai') {
      var incomeData = {
        date: data.receiveDate,
        category: 'Cổ tức',
        amount: totalDividend,
        note: 'Cổ tức ' + data.stockCode
      };
      
      addIncome(incomeData);
    }
    
    // 3. Nếu là cổ tức CP → Cập nhật portfolio
    if (data.dividendType === 'Cổ phiếu' || 
        data.dividendType === 'Cả hai') {
      updateStockPortfolio(data.stockCode, data.stockDividendQuantity);
    }
    
    updateDashboard();
    
    Logger.log('Added dividend: ' + JSON.stringify(data));
    return { success: true };
  } catch (error) {
    Logger.log('Error in addDividend: ' + error);
    return { success: false, message: error.message };
  }
}
```

**Dependency:** 
- Utils.gs (findEmptyRow, calculateDueDate...)
- BudgetManager.gs (updateBudgetCategory)
- DashboardManager.gs (updateDashboard)

---

### 4. BudgetManager.gs - Budget Tracking

**Mục đích:** Quản lý và theo dõi ngân sách

**Functions chính:**

```javascript
/**
 * Đặt ngân sách
 */
function setBudget(data) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName('BUDGET');
    
    // Clear existing budget
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn())
      .clearContent();
    
    // Headers
    var headers = [
      'Danh mục', 'Ngân sách', 'Đã chi', 'Còn lại', '%', 'Trạng thái'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Chi tiêu categories
    var expenseCategories = [
      ['Ăn uống', data.anUong || 0],
      ['Di chuyển', data.diChuyen || 0],
      ['Nhà cửa', data.nhaCua || 0],
      ['Điện nước', data.dienNuoc || 0],
      ['Viễn thông', data.vienThong || 0],
      ['Giáo dục', data.giaoDuc || 0],
      ['Y tế', data.yTe || 0],
      ['Mua sắm', data.muaSam || 0],
      ['Giải trí', data.giaiTri || 0],
      ['Khác', data.khac || 0]
    ];
    
    var startRow = 2;
    for (var i = 0; i < expenseCategories.length; i++) {
      var row = startRow + i;
      var category = expenseCategories[i][0];
      var budget = expenseCategories[i][1];
      
      sheet.getRange(row, 1).setValue(category);
      sheet.getRange(row, 2).setValue(budget);
      
      // Formula: Đã chi
      sheet.getRange(row, 3)
        .setFormula('=SUMIF(CHI!C:C,"' + category + '",CHI!D:D)');
      
      // Formula: Còn lại
      sheet.getRange(row, 4).setFormula('=B' + row + '-C' + row);
      
      // Formula: %
      sheet.getRange(row, 5)
        .setFormula('=IF(B' + row + '=0,0,C' + row + '/B' + row + ')');
      
      // Formula: Trạng thái
      sheet.getRange(row, 6)
        .setFormula(
          '=IF(E' + row + '<0.7,"🟢",' +
          'IF(E' + row + '<0.9,"🟡","🔴"))'
        );
    }
    
    // Formatting
    sheet.getRange(startRow, 5, expenseCategories.length, 1)
      .setNumberFormat('0.0%');
    sheet.getRange(startRow, 2, expenseCategories.length, 3)
      .setNumberFormat('#,##0');
    
    // Conditional formatting for %
    applyBudgetConditionalFormatting(sheet, startRow, 
                                     expenseCategories.length);
    
    Logger.log('Budget set successfully');
    return { success: true };
  } catch (error) {
    Logger.log('Error in setBudget: ' + error);
    return { success: false, message: error.message };
  }
}

/**
 * Cập nhật budget khi có chi tiêu mới
 */
function updateBudgetCategory(category, amount) {
  // Budget tự động cập nhật qua formulas
  // Function này để log hoặc trigger alerts
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName('BUDGET');
  
  var data = sheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === category) {
      var budget = data[i][1];
      var spent = data[i][2]; // Sẽ tự update qua formula
      var percentage = spent / budget;
      
      // Alert nếu vượt 90%
      if (percentage >= 0.9) {
        sendBudgetAlert(category, percentage);
      }
      
      break;
    }
  }
}

/**
 * Gửi cảnh báo ngân sách
 */
function sendBudgetAlert(category, percentage) {
  var message = '⚠️ CẢNH BÁO NGÂN SÁCH\n\n' +
                'Danh mục: ' + category + '\n' +
                'Đã chi: ' + (percentage * 100).toFixed(1) + '%\n';
  
  if (percentage >= 1.0) {
    message += '\n🔴 ĐÃ VƯỢT NGÂN SÁCH!';
  } else {
    message += '\n🟡 SẮP VƯỢT NGÂN SÁCH!';
  }
  
  Logger.log(message);
  
  // TODO: Send email or push notification
  // MailApp.sendEmail(userEmail, 'Budget Alert', message);
}

/**
 * Apply conditional formatting
 */
function applyBudgetConditionalFormatting(sheet, startRow, numRows) {
  var range = sheet.getRange(startRow, 5, numRows, 1);
  
  // Rule 1: < 70% → Green
  var rule1 = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0.7)
    .setBackground('#B7E1CD')
    .setRanges([range])
    .build();
  
  // Rule 2: 70-90% → Yellow
  var rule2 = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(0.7, 0.9)
    .setBackground('#FFE599')
    .setRanges([range])
    .build();
  
  // Rule 3: > 90% → Red
  var rule3 = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0.9)
    .setBackground('#F4CCCC')
    .setRanges([range])
    .build();
  
  var rules = sheet.getConditionalFormatRules();
  rules.push(rule1, rule2, rule3);
  sheet.setConditionalFormatRules(rules);
}
```

**Dependency:** Utils.gs

---

### 5. DashboardManager.gs - Analytics

**Mục đích:** Tạo dashboard và báo cáo

**Functions chính:**

```javascript
/**
 * Cập nhật Dashboard
 */
function updateDashboard() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var dashboard = ss.getSheetByName('Dashboard');
    
    // Calculate totals
    var totals = calculateTotals();
    
    // Update cells
    dashboard.getRange('B5').setValue(totals.totalIncome);
    dashboard.getRange('B6').setValue(totals.totalExpense);
    dashboard.getRange('B7').setValue(totals.totalInvestment);
    dashboard.getRange('B8').setValue(totals.totalDebtPayment);
    dashboard.getRange('B9').setValue(totals.cashFlow);
    
    // Update charts
    updateCharts();
    
    Logger.log('Dashboard updated');
  } catch (error) {
    Logger.log('Error in updateDashboard: ' + error);
  }
}

/**
 * Tính tổng các chỉ số
 */
function calculateTotals() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Thu nhập
  var incomeSheet = ss.getSheetByName('THU');
  var incomeData = incomeSheet.getRange(2, 4, 
                                        incomeSheet.getLastRow() - 1, 1)
    .getValues();
  var totalIncome = sumArray(incomeData);
  
  // Chi tiêu
  var expenseSheet = ss.getSheetByName('CHI');
  var expenseData = expenseSheet.getRange(2, 4, 
                                          expenseSheet.getLastRow() - 1, 1)
    .getValues();
  var totalExpense = sumArray(expenseData);
  
  // Đầu tư (tổng từ 4 loại)
  var totalInvestment = 
    calculateTotalInvestment('CHỨNG KHOÁN') +
    calculateTotalInvestment('VÀNG') +
    calculateTotalInvestment('CRYPTO') +
    calculateTotalInvestment('ĐẦU TƯ KHÁC');
  
  // Trả nợ
  var debtPaymentSheet = ss.getSheetByName('TRẢ NỢ');
  var debtPaymentData = debtPaymentSheet.getRange(2, 6, 
                                                  debtPaymentSheet.getLastRow() - 1, 1)
    .getValues();
  var totalDebtPayment = sumArray(debtPaymentData);
  
  // Cash Flow = Thu - (Chi + Trả nợ + Đầu tư)
  var cashFlow = totalIncome - 
                 (totalExpense + totalDebtPayment + totalInvestment);
  
  return {
    totalIncome: totalIncome,
    totalExpense: totalExpense,
    totalInvestment: totalInvestment,
    totalDebtPayment: totalDebtPayment,
    cashFlow: cashFlow
  };
}

/**
 * Tính tổng đầu tư theo loại
 */
function calculateTotalInvestment(sheetName) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(sheetName);
  
  if (!sheet || sheet.getLastRow() <= 1) return 0;
  
  var data = sheet.getDataRange().getValues();
  var total = 0;
  
  // Cột tổng tiền khác nhau tùy loại
  var amountCol = getInvestmentAmountColumn(sheetName);
  
  for (var i = 1; i < data.length; i++) {
    var transType = data[i][2]; // Column C - Loại GD
    var amount = data[i][amountCol];
    
    if (transType === 'Mua' || transType === 'Mua/Góp vốn') {
      total += amount;
    } else if (transType === 'Bán' || transType === 'Bán/Rút vốn') {
      total -= amount;
    }
  }
  
  return total;
}

/**
 * Cập nhật biểu đồ
 */
function updateCharts() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dashboard = ss.getSheetByName('Dashboard');
  
  // Chart 1: Thu Chi theo tháng
  createIncomeExpenseChart(dashboard);
  
  // Chart 2: Phân bổ chi tiêu
  createExpensePieChart(dashboard);
  
  // Chart 3: Danh mục đầu tư
  createInvestmentChart(dashboard);
}

/**
 * Tạo biểu đồ Thu Chi
 */
function createIncomeExpenseChart(dashboard) {
  // Remove existing chart if any
  var charts = dashboard.getCharts();
  charts.forEach(function(chart) {
    if (chart.getOptions().get('title') === 'Thu Chi theo tháng') {
      dashboard.removeChart(chart);
    }
  });
  
  // Create new chart
  var dataRange = dashboard.getRange('A15:C20'); // Sample range
  
  var chart = dashboard.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(dataRange)
    .setPosition(5, 5, 0, 0)
    .setOption('title', 'Thu Chi theo tháng')
    .setOption('width', 600)
    .setOption('height', 400)
    .build();
  
  dashboard.insertChart(chart);
}

/**
 * Phân tích theo period
 */
function analyzeByPeriod(startDate, endDate) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var result = {
    income: 0,
    expense: 0,
    investment: 0,
    debtPayment: 0,
    cashFlow: 0,
    expenseByCategory: {},
    investmentByType: {}
  };
  
  // Filter and sum by date range
  // Implementation details...
  
  return result;
}
```

**Dependency:** Utils.gs (sumArray, date functions...)

---

### 6. Utils.gs - Helper Functions

**Mục đích:** Các hàm tiện ích dùng chung

**Functions chính:**

```javascript
/**
 * Tìm row trống dựa vào column có data
 * @param {Sheet} sheet - Sheet cần tìm
 * @param {number} dataColumn - Cột chứa data (1-indexed)
 * @return {number} Row number trống
 */
function findEmptyRow(sheet, dataColumn) {
  var data = sheet.getRange(2, dataColumn, 
                             sheet.getMaxRows() - 1, 1)
    .getValues();
  
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === '' || data[i][0] === null) {
      return i + 2; // +2 vì bắt đầu từ row 2
    }
  }
  
  // Nếu không tìm thấy, return row cuối + 1
  return sheet.getLastRow() + 1;
}

/**
 * Sum array values
 */
function sumArray(arr) {
  var sum = 0;
  for (var i = 0; i < arr.length; i++) {
    if (typeof arr[i][0] === 'number') {
      sum += arr[i][0];
    }
  }
  return sum;
}

/**
 * Calculate due date
 */
function calculateDueDate(startDate, months) {
  var date = new Date(startDate);
  date.setMonth(date.getMonth() + months);
  return Utilities.formatDate(date, 'GMT+7', 'yyyy-MM-dd');
}

/**
 * Format currency VND
 */
function formatCurrency(amount) {
  return amount.toLocaleString('vi-VN') + ' VNĐ';
}

/**
 * Format date Vietnamese style
 */
function formatDateVN(date) {
  return Utilities.formatDate(date, 'GMT+7', 'dd/MM/yyyy');
}

/**
 * Get current timestamp
 */
function getCurrentTimestamp() {
  return Utilities.formatDate(new Date(), 'GMT+7', 
                               'yyyy-MM-dd HH:mm:ss');
}

/**
 * Validate email
 */
function isValidEmail(email) {
  var re = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return re.test(email);
}

/**
 * Validate phone number (VN)
 */
function isValidPhone(phone) {
  var re = /^(0|\+84)[0-9]{9,10}$/;
  return re.test(phone);
}

/**
 * Get sheet by name safely
 */
function getSheetSafe(sheetName) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(sheetName);
  
  if (!sheet) {
    throw new Error('Sheet not found: ' + sheetName);
  }
  
  return sheet;
}

/**
 * Log with timestamp
 */
function logWithTimestamp(message) {
  Logger.log('[' + getCurrentTimestamp() + '] ' + message);
}

/**
 * Get investment amount column by sheet name
 */
function getInvestmentAmountColumn(sheetName) {
  switch (sheetName) {
    case 'CHỨNG KHOÁN': return 9; // Column J - Tổng tiền
    case 'VÀNG': return 6; // Column G - Tổng giá trị
    case 'CRYPTO': return 7; // Column H - Giá trị VNĐ
    case 'ĐẦU TƯ KHÁC': return 4; // Column E - Số tiền
    default: return 4;
  }
}

/**
 * Retry function with exponential backoff
 */
function retryWithBackoff(func, maxRetries, initialDelay) {
  maxRetries = maxRetries || 3;
  initialDelay = initialDelay || 1000;
  
  for (var i = 0; i < maxRetries; i++) {
    try {
      return func();
    } catch (error) {
      if (i === maxRetries - 1) throw error;
      
      var delay = initialDelay * Math.pow(2, i);
      Utilities.sleep(delay);
    }
  }
}
```

**Dependency:** Không

---

## 🗄️ Cấu trúc dữ liệu

### Sheet: THU (Income)

| Column | Name | Type | Description | Example |
|--------|------|------|-------------|---------|
| A | STT | Formula | =IF(B2<>"", ROW()-1, "") | 1 |
| B | Ngày | Date | Ngày thu | 2025-10-29 |
| C | Danh mục | String | Loại thu nhập | Lương |
| D | Số tiền | Number | Số tiền VNĐ | 15000000 |
| E | Ghi chú | String | Mô tả chi tiết | Lương tháng 10 |

### Sheet: CHI (Expense)

| Column | Name | Type | Description | Example |
|--------|------|------|-------------|---------|
| A | STT | Formula | Auto increment | 1 |
| B | Ngày | Date | Ngày chi | 2025-10-29 |
| C | Danh mục | String | Loại chi tiêu | Ăn uống |
| D | Số tiền | Number | Số tiền VNĐ | 200000 |
| E | Phương thức | String | Thanh toán | Tiền mặt |
| F | Ghi chú | String | Mô tả | Cơm trưa |

### Sheet: QUẢN LÝ NỢ (Debt Management)

| Column | Name | Type | Description |
|--------|------|------|-------------|
| A | STT | Formula | Auto |
| B | Ngày vay | Date | Loan date |
| C | Loại nợ | String | Debt type |
| D | Số tiền vay | Number | Principal |
| E | Lãi suất (%/năm) | Number | Interest rate |
| F | Kỳ hạn (tháng) | Number | Term |
| G | Ngày đáo hạn | Date | Due date |
| H | Mục đích | String | Purpose |
| I | Ghi chú | String | Note |
| J | Số dư | Formula | Balance |
| K | Trạng thái | String | Status |

### Sheet: CHỨNG KHOÁN (Stocks)

| Column | Name | Type | Description |
|--------|------|------|-------------|
| A | STT | Formula | Auto |
| B | Ngày GD | Date | Transaction date |
| C | Loại GD | String | Buy/Sell |
| D | Mã CP | String | Stock code |
| E | Số lượng | Number | Quantity |
| F | Giá | Number | Price/share |
| G | Tổng giá trị | Formula | =E*F |
| H | Phí | Formula | =G*0.15% |
| I | Thuế | Formula | =IF(C="Bán",G*0.1%,0) |
| J | Tổng tiền | Formula | =G+H+I |
| K | Margin | String | Yes/No |
| L | Số tiền vay | Number | Margin amount |
| M | Lãi suất | Number | Margin rate |
| N | Ghi chú | String | Note |

### Sheet: BUDGET

| Column | Name | Type | Description |
|--------|------|------|-------------|
| A | Danh mục | String | Category |
| B | Ngân sách | Number | Budget |
| C | Đã chi | Formula | =SUMIF(CHI!C:C,A2,CHI!D:D) |
| D | Còn lại | Formula | =B-C |
| E | % | Formula | =C/B |
| F | Trạng thái | Formula | Color indicator |

---

## 🔄 Data Flow

### Flow 1: Thêm Thu nhập

```
User fills IncomeForm.html
        ↓
Form submits to addIncome()
        ↓
DataProcessor validates data
        ↓
findEmptyRow() tìm row trống
        ↓
Insert data vào sheet THU
        ↓
updateDashboard() cập nhật
        ↓
Dashboard shows new totals
```

### Flow 2: Thêm Nợ (có tự động tạo Thu)

```
User fills DebtManagementForm.html
        ↓
Form submits to addDebt()
        ↓
DataProcessor validates
        ↓
Insert vào QUẢN LÝ NỢ sheet
        ↓
AUTO: Create income record
        ↓
addIncome() with "Vay [Type]"
        ↓
Insert vào THU sheet
        ↓
updateDashboard()
        ↓
Both sheets updated + Dashboard
```

### Flow 3: Trả Nợ

```
User fills DebtPaymentForm.html
        ↓
Dropdown loads from QUẢN LÝ NỢ
        ↓
User selects debt & enters amounts
        ↓
Form submits to payDebt()
        ↓
Insert vào TRẢ NỢ sheet
        ↓
updateDebtBalance() updates QUẢN LÝ NỢ
        ↓
If balance = 0 → Status = "Đã thanh toán"
        ↓
updateDashboard()
```

### Flow 4: Chi tiêu với Budget

```
User fills ExpenseForm.html
        ↓
Form submits to addExpense()
        ↓
Insert vào CHI sheet
        ↓
updateBudgetCategory() checks %
        ↓
If % >= 90% → sendBudgetAlert()
        ↓
BUDGET sheet auto-updates (formulas)
        ↓
Conditional formatting updates color
        ↓
updateDashboard()
```

---

## 📐 Formulas & Calculations

### Cash Flow Formula

```
Cash Flow = Tổng Thu - (Tổng Chi + Tổng Trả Nợ + Tổng Đầu Tư)
```

**Implementation:**
```javascript
var cashFlow = totalIncome - 
               (totalExpense + totalDebtPayment + totalInvestment);
```

**In Sheet (Dashboard):**
```
=SUMIF(THU!D:D) - 
 (SUMIF(CHI!D:D) + SUMIF('TRẢ NỢ'!F:F) + 
  [Investment totals])
```

### Budget Percentage

```
% = Đã chi / Ngân sách × 100
```

**In Sheet:**
```
=IF(B2=0, 0, C2/B2)
```

### Interest Calculation (Monthly)

```
Lãi 1 tháng = (Số dư × Lãi suất/năm) / 12
```

**Example:**
```
Số dư: 500,000,000 VNĐ
Lãi suất: 8.5%/năm

Lãi 1 tháng = (500,000,000 × 8.5%) / 12
            = 3,541,667 VNĐ
```

### Stock Transaction Total

```
Khi MUA:
Tổng = (Số lượng × Giá) + Phí
Phí = Tổng giá trị × 0.15%

Khi BÁN:
Tổng = (Số lượng × Giá) - Phí - Thuế
Phí = Tổng giá trị × 0.15%
Thuế = Tổng giá trị × 0.1%
```

### Dividend Yield

```
Dividend Yield = (Cổ tức/CP / Giá mua) × 100%
```

---

## 📝 Code Conventions

### Naming Conventions

**Variables:**
```javascript
// camelCase
var totalIncome = 0;
var userEmail = 'test@example.com';

// Constants: UPPER_SNAKE_CASE
var MAX_RETRIES = 3;
var DEFAULT_INTEREST_RATE = 12;
```

**Functions:**
```javascript
// camelCase, verb-first
function addIncome(data) { }
function updateDashboard() { }
function calculateTotals() { }
```

**Sheet Names:**
```javascript
// UPPER CASE Vietnamese
'THU'
'CHI'
'QUẢN LÝ NỢ'
'CHỨNG KHOÁN'
```

### Comments

```javascript
/**
 * Function description
 * @param {Type} paramName - Description
 * @return {Type} Description
 */
function myFunction(paramName) {
  // Implementation
}

// Single-line comment for clarification
var x = 10; // Inline comment
```

### Error Handling

```javascript
function addIncome(data) {
  try {
    // Main logic
    
    return { success: true };
  } catch (error) {
    Logger.log('Error in addIncome: ' + error);
    return { 
      success: false, 
      message: error.message 
    };
  }
}
```

### Function Structure

```javascript
function functionName(params) {
  // 1. Input validation
  if (!params || !params.required) {
    throw new Error('Missing required parameter');
  }
  
  // 2. Get resources
  var sheet = getSheetSafe('THU');
  
  // 3. Main logic
  var result = doSomething();
  
  // 4. Side effects
  updateDashboard();
  
  // 5. Return
  return result;
}
```

---

## 🧪 Testing Guidelines

### Manual Testing Checklist

```
□ Test all 10 forms individually
□ Test data insertion (check row position)
□ Test formulas auto-calculate
□ Test budget updates when adding expense
□ Test debt auto-creates income
□ Test margin creates debt
□ Test dashboard updates
□ Test with empty sheets
□ Test with existing data
□ Test edge cases (zero amounts, special characters)
```

### Unit Test Example

```javascript
function testAddIncome() {
  var testData = {
    date: '2025-10-29',
    category: 'Lương',
    amount: 15000000,
    note: 'Test'
  };
  
  var result = addIncome(testData);
  
  if (result.success) {
    Logger.log('✅ testAddIncome PASSED');
  } else {
    Logger.log('❌ testAddIncome FAILED: ' + result.message);
  }
}
```

### Integration Test Example

```javascript
function testDebtAutoIncomeFlow() {
  // 1. Add debt
  var debtData = {
    loanDate: '2025-10-29',
    debtType: 'Vay ngân hàng',
    amount: 50000000,
    interestRate: 8.5,
    term: 12,
    dueDate: '2026-10-29',
    purpose: 'Test'
  };
  
  addDebt(debtData);
  
  // 2. Check income auto-created
  var incomeSheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName('THU');
  var lastRow = incomeSheet.getLastRow();
  var incomeCategory = incomeSheet.getRange(lastRow, 3).getValue();
  
  if (incomeCategory === 'Vay Vay ngân hàng') {
    Logger.log('✅ testDebtAutoIncomeFlow PASSED');
  } else {
    Logger.log('❌ testDebtAutoIncomeFlow FAILED');
  }
}
```

---

## 🚀 Deployment

### Step 1: Prepare Code

```
1. Review all .gs files
2. Run all tests
3. Update version number
4. Update CHANGELOG.md
```

### Step 2: Create Template

```
1. Make a clean copy of spreadsheet
2. Remove sample data
3. Keep formulas (rows 2-1000)
4. Keep all sheets structure
```

### Step 3: Deploy Apps Script

```
1. Apps Script Editor
2. Deploy → New deployment
3. Type: Web app (if needed)
4. Version: New version
5. Description: "v3.0 Release"
6. Execute as: Me
7. Who has access: Anyone
8. Deploy
```

### Step 4: Share Template

```
1. File → Share → Get link
2. Set to "Anyone with link can view"
3. Copy link
4. Update README.md with link
```

### Step 5: Documentation

```
1. Update all docs with latest info
2. Create video tutorials
3. Prepare FAQ
4. Post to community
```

---

## ⚡ Performance Optimization

### 1. Batch Operations

**Bad:**
```javascript
for (var i = 0; i < 1000; i++) {
  sheet.getRange(i + 2, 1).setValue(i);
}
```

**Good:**
```javascript
var data = [];
for (var i = 0; i < 1000; i++) {
  data.push([i]);
}
sheet.getRange(2, 1, data.length, 1).setValues(data);
```

### 2. Minimize getRange() calls

**Bad:**
```javascript
var val1 = sheet.getRange('A1').getValue();
var val2 = sheet.getRange('A2').getValue();
var val3 = sheet.getRange('A3').getValue();
```

**Good:**
```javascript
var values = sheet.getRange('A1:A3').getValues();
var val1 = values[0][0];
var val2 = values[1][0];
var val3 = values[2][0];
```

### 3. Cache frequently accessed data

```javascript
var CACHE_KEY = 'dashboard_totals';
var CACHE_TTL = 300; // 5 minutes

function getCachedTotals() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get(CACHE_KEY);
  
  if (cached) {
    return JSON.parse(cached);
  }
  
  var totals = calculateTotals();
  cache.put(CACHE_KEY, JSON.stringify(totals), CACHE_TTL);
  return totals;
}
```

### 4. Use formulas instead of scripts

**Prefer:**
```
=SUMIF(CHI!C:C, "Ăn uống", CHI!D:D)
```

**Over:**
```javascript
function sumExpenseCategory(category) {
  // Script to sum...
}
```

---

## 🔐 Security

### 1. Input Validation

```javascript
function validateExpenseData(data) {
  if (!data.date || !isValidDate(data.date)) {
    throw new Error('Invalid date');
  }
  
  if (!data.amount || data.amount <= 0) {
    throw new Error('Amount must be > 0');
  }
  
  if (!data.category) {
    throw new Error('Category is required');
  }
  
  // Sanitize inputs
  data.note = sanitizeString(data.note);
  
  return true;
}

function sanitizeString(str) {
  if (!str) return '';
  return str.replace(/<script[^>]*>.*?<\/script>/gi, '')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;');
}
```

### 2. Protect Sensitive Data

```javascript
// Store in Script Properties, not in code
var props = PropertiesService.getScriptProperties();
props.setProperty('API_KEY', 'secret_key_here');

// Retrieve
var apiKey = props.getProperty('API_KEY');
```

### 3. Permission Management

```
Apps Script Permissions:
✅ Read/Write Spreadsheet
✅ Display UI
❌ Send Emails (disable if not used)
❌ External API calls (restrict domains)
```

---

## 🐛 Debugging

### Logging

```javascript
// Basic log
Logger.log('Message');

// With timestamp
logWithTimestamp('Action completed');

// Object logging
Logger.log('Data: ' + JSON.stringify(data));

// View logs
Apps Script Editor → View → Logs
```

### Breakpoints

```javascript
// Add debugger statement
function myFunction() {
  var x = 10;
  debugger; // Execution pauses here
  var y = x * 2;
}
```

### Common Issues

**Issue 1: "Cannot read property 'getSheetByName' of null"**
```
Solution: Check sheet name spelling
Check sheet exists
```

**Issue 2: "Authorization required"**
```
Solution: Re-run function to authorize
Clear authorization and re-authorize
```

**Issue 3: "Data appears at row 1001"**
```
Solution: Use findEmptyRow() with data column
Don't rely on getLastRow() when pre-filled formulas exist
```

---

## 🤝 Contributing

### How to Contribute

1. Fork repository
2. Create feature branch
3. Make changes
4. Test thoroughly
5. Submit Pull Request

### Code Review Checklist

```
□ Follows code conventions
□ Includes comments
□ Passes all tests
□ No console.log() left
□ Error handling implemented
□ Performance optimized
□ Documentation updated
```

### Git Commit Messages

```
Format:
[TYPE] Brief description

TYPE:
- feat: New feature
- fix: Bug fix
- docs: Documentation
- refactor: Code refactoring
- test: Tests
- perf: Performance

Example:
[feat] Add cryptocurrency tracking
[fix] Fix appendRow bug with pre-filled formulas
[docs] Update API reference
```

---

## 📚 Resources

### Google Apps Script Docs

- [Official Docs](https://developers.google.com/apps-script)
- [Spreadsheet Service](https://developers.google.com/apps-script/reference/spreadsheet)
- [HTML Service](https://developers.google.com/apps-script/reference/html)

### Community

- GitHub Issues
- Facebook Group
- Stack Overflow Tag: google-apps-script

---

<div align="center">

**Happy Coding! 💻🚀**

[⬆ Về đầu trang](#-tài-liệu-kỹ-thuật---hodlvn-family-finance)

</div>