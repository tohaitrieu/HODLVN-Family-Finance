/**
 * ===============================================
 * SHEETINITIALIZER.GS - MODULE KHỞI TẠO CÁC SHEET
 * ===============================================
 * 
 * Chức năng:
 * - Khởi tạo cấu trúc và format cho từng sheet
 * - Tạo header, validation, công thức
 * - Áp dụng format và màu sắc nhất quán
 * - ✅ SAFE MODE: Không xóa dữ liệu cũ khi chạy lại
 * 
 * VERSION: 4.0 - Non-destructive Updates & Standardized Formats
 */

const SheetInitializer = {
  
  /**
   * Helper: Get or Create Sheet
   */
  _getOrCreateSheet(ss, sheetName) {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }
    return sheet;
  },

  /**
   * Helper: Fix Date Column Format
   * Converts text dates (dd/MM/yyyy) to Date objects and sets format
   * Also strips time components from existing Date objects
   * Also handles Excel serial numbers
   */
  _fixDateColumn(sheet, colIndex) {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;
    
    const range = sheet.getRange(2, colIndex, lastRow - 1, 1);
    const values = range.getValues();
    let hasChange = false;
    
    const fixedValues = values.map(row => {
      const val = row[0];
      if (!val) return [val];

      let dateObj = null;

      // Case 1: String "dd/mm/yyyy"
      if (typeof val === 'string' && val.match(/^\d{1,2}\/\d{1,2}\/\d{4}/)) {
        const parts = val.split('/');
        // Note: Month is 0-indexed in JS Date
        dateObj = new Date(parts[2], parts[1] - 1, parts[0]);
      } 
      // Case 2: Already a Date object (potentially with time)
      else if (val instanceof Date) {
        dateObj = val;
      }
      // Case 3: Excel serial number (e.g., 48030.29166...)
      else if (typeof val === 'number' && val > 1000) {
        // Excel serial date: days since 1900-01-01
        // Google Sheets uses same system
        const excelEpoch = new Date(1899, 11, 30); // Excel epoch (Dec 30, 1899)
        const milliseconds = Math.round((val - Math.floor(val)) * 86400000); // fractional part to ms
        dateObj = new Date(excelEpoch.getTime() + Math.floor(val) * 86400000 + milliseconds);
      }

      if (dateObj && !isNaN(dateObj.getTime())) {
        // Create new date with time set to 00:00:00
        const cleanDate = new Date(dateObj.getFullYear(), dateObj.getMonth(), dateObj.getDate());
        
        // Check if value actually changed (to avoid unnecessary writes)
        if (val instanceof Date && val.getTime() === cleanDate.getTime()) {
          return [val];
        }
        
        hasChange = true;
        return [cleanDate];
      }
      
      return [val];
    });
    
    if (hasChange) {
      range.setValues(fixedValues);
    }
    range.setNumberFormat(APP_CONFIG.FORMATS.DATE);
  },

  /**
   * Helper: Fix Term Column Format
   * Converts Date objects back to plain numbers (months)
   * This fixes cases where Google Sheets auto-converted numbers to dates
   */
  _fixTermColumn(sheet, colIndex) {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;
    
    const range = sheet.getRange(2, colIndex, lastRow - 1, 1);
    const values = range.getValues();
    let hasChange = false;
    
    const fixedValues = values.map(row => {
      const val = row[0];
      if (!val) return [val];

      // If it's a Date object, extract the numeric value
      // Google Sheets might have auto-converted "48" to a date like "Feb 17, 1900"
      if (val instanceof Date) {
        // Get the serial number and use it as the term
        // This is a workaround - ideally we'd have the original number
        // For dates in 1900, the day of year is roughly the number
        const serialNumber = Math.round((val - new Date(1899, 11, 30)) / 86400000);
        hasChange = true;
        return [serialNumber];
      }
      
      // If it's already a number, ensure it's an integer
      if (typeof val === 'number') {
        const intVal = Math.round(val);
        if (intVal !== val) {
          hasChange = true;
          return [intVal];
        }
      }
      
      return [val];
    });
    
    if (hasChange) {
      range.setValues(fixedValues);
    }
    range.setNumberFormat('0'); // Plain integer, no decimals
  },

  /**
   * Cập nhật toàn bộ các Sheet (Non-destructive)
   * Chạy lại logic khởi tạo để cập nhật header, format, validation
   * mà không làm mất dữ liệu cũ.
   */
  updateAllSheets() {
    this.initializeIncomeSheet();
    this.initializeExpenseSheet();
    this.initializeDebtPaymentSheet();
    this.initializeDebtManagementSheet();
    this.initializeLendingSheet();
    this.initializeLendingRepaymentSheet();
    this.initializeStockSheet();
    this.initializeGoldSheet();
    this.initializeCryptoSheet();
    this.initializeOtherInvestmentSheet();
    this.initializeBudgetSheet();
    this.initializeChangelogSheet();
    
    // Cập nhật nội dung Changelog
    if (typeof ChangelogManager !== 'undefined') {
      ChangelogManager.updateChangelogSheet();
    }
    
    // Dashboard được cập nhật riêng qua DashboardManager
    if (typeof DashboardManager !== 'undefined') {
      DashboardManager.setupDashboard();
    }
    
    // Sắp xếp lại thứ tự Sheet
    this.reorderSheets();
  },

  /**
   * Sắp xếp lại thứ tự các Sheet theo quy định
   */
  reorderSheets() {
    const ss = getSpreadsheet();
    const desiredOrder = [
      APP_CONFIG.SHEETS.DASHBOARD,        // 1. Tổng quan
      APP_CONFIG.SHEETS.INCOME,           // 2. Thu
      APP_CONFIG.SHEETS.EXPENSE,          // 3. Chi
      APP_CONFIG.SHEETS.BUDGET,           // 4. Budget
      APP_CONFIG.SHEETS.DEBT_MANAGEMENT,  // 5. Quản lý nợ
      APP_CONFIG.SHEETS.DEBT_PAYMENT,     // 6. Trả nợ
      APP_CONFIG.SHEETS.GOLD,             // 7. Vàng
      APP_CONFIG.SHEETS.STOCK,            // 8. Chứng khoán
      APP_CONFIG.SHEETS.CRYPTO,           // 9. Crypto
      APP_CONFIG.SHEETS.LENDING,          // 10. Cho vay
      APP_CONFIG.SHEETS.LENDING_REPAYMENT,// 11. Thu nợ
      APP_CONFIG.SHEETS.OTHER_INVESTMENT, // 12. Đầu tư khác
      APP_CONFIG.SHEETS.CHANGELOG         // 13. Lịch sử cập nhật
    ];
    
    desiredOrder.forEach((sheetName, index) => {
      const sheet = ss.getSheetByName(sheetName);
      if (sheet) {
        ss.setActiveSheet(sheet);
        ss.moveActiveSheet(index + 1);
      }
    });
    
    // Quay về Dashboard
    const dashboard = ss.getSheetByName(APP_CONFIG.SHEETS.DASHBOARD);
    if (dashboard) {
      ss.setActiveSheet(dashboard);
    }
  },

  /**
   * Khởi tạo Sheet THU
   */
  initializeIncomeSheet() {
    const ss = getSpreadsheet();
    const sheet = this._getOrCreateSheet(ss, APP_CONFIG.SHEETS.INCOME);
    
    // Header
    const headers = ['STT', 'Ngày', 'Số tiền', 'Nguồn thu', 'Ghi chú', 'TransactionID'];
    sheet.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBackground(APP_CONFIG.COLORS.HEADER_BG)
      .setFontColor(APP_CONFIG.COLORS.HEADER_TEXT);
    
    // Column widths
    sheet.setColumnWidth(1, 50);   // STT
    sheet.setColumnWidth(2, 100);  // Ngày
    sheet.setColumnWidth(3, 120);  // Số tiền
    sheet.setColumnWidth(4, 150);  // Nguồn thu
    sheet.setColumnWidth(5, 300);  // Ghi chú
    sheet.hideColumns(6);          // Hide TransactionID
    
    // Format - Apply to whole columns (safe)
    sheet.getRange('A2:A').setNumberFormat('0');
    sheet.getRange('B2:B').setNumberFormat(APP_CONFIG.FORMATS.DATE);
    this._fixDateColumn(sheet, 2);
    sheet.getRange('C2:C').setNumberFormat('#,##0');
    
    // Freeze header
    sheet.setFrozenRows(1);
    
    // Data validation
    const sourceRange = sheet.getRange('D2:D1000');
    const sourceRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(APP_CONFIG.CATEGORIES.INCOME)
      .setAllowInvalid(false)
      .build();
    sourceRange.setDataValidation(sourceRule);
    
    return sheet;
  },
  
  /**
   * Khởi tạo Sheet CHI
   */
  initializeExpenseSheet() {
    const ss = getSpreadsheet();
    const sheet = this._getOrCreateSheet(ss, APP_CONFIG.SHEETS.EXPENSE);
    
    // Header
    const headers = ['STT', 'Ngày', 'Số tiền', 'Danh mục', 'Chi tiết', 'Ghi chú', 'TransactionID'];
    sheet.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBackground(APP_CONFIG.COLORS.HEADER_BG)
      .setFontColor(APP_CONFIG.COLORS.HEADER_TEXT);
    
    // Column widths
    sheet.setColumnWidth(1, 50);
    sheet.setColumnWidth(2, 100);
    sheet.setColumnWidth(3, 120);
    sheet.setColumnWidth(4, 120);
    sheet.setColumnWidth(5, 200);
    sheet.setColumnWidth(6, 250);
    sheet.hideColumns(7);          // Hide TransactionID
    
    // Format
    sheet.getRange('A2:A').setNumberFormat('0');
    sheet.getRange('B2:B').setNumberFormat(APP_CONFIG.FORMATS.DATE);
    this._fixDateColumn(sheet, 2);
    sheet.getRange('C2:C').setNumberFormat('#,##0');
    
    // Freeze header
    sheet.setFrozenRows(1);
    
    // Data validation
    const categoryRange = sheet.getRange('D2:D1000');
    const categoryRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(APP_CONFIG.CATEGORIES.EXPENSE)
      .setAllowInvalid(false)
      .build();
    categoryRange.setDataValidation(categoryRule);
    
    return sheet;
  },
  
  /**
   * Khởi tạo Sheet TRẢ NỢ
   */
  initializeDebtPaymentSheet() {
    const ss = getSpreadsheet();
    const sheet = this._getOrCreateSheet(ss, APP_CONFIG.SHEETS.DEBT_PAYMENT);
    
    // Header
    const headers = ['STT', 'Ngày', 'Khoản nợ', 'Trả gốc', 'Trả lãi', 'Tổng trả', 'Ghi chú', 'TransactionID'];
    sheet.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBackground(APP_CONFIG.COLORS.HEADER_BG)
      .setFontColor(APP_CONFIG.COLORS.HEADER_TEXT);
    
    // Column widths
    sheet.setColumnWidth(1, 50);
    sheet.setColumnWidth(2, 100);
    sheet.setColumnWidth(3, 150);
    sheet.setColumnWidth(4, 120);
    sheet.setColumnWidth(5, 120);
    sheet.setColumnWidth(6, 120);
    sheet.setColumnWidth(7, 250);
    sheet.hideColumns(8);          // Hide TransactionID
    
    // Format
    sheet.getRange('A2:A').setNumberFormat('0');
    sheet.getRange('B2:B').setNumberFormat(APP_CONFIG.FORMATS.DATE);
    this._fixDateColumn(sheet, 2);
    sheet.getRange('D2:F').setNumberFormat('#,##0');
    
    // Formula (Safe to re-apply)
    sheet.getRange('F2:F1000').setFormula('=IFERROR(D2+E2, 0)');
    
    sheet.setFrozenRows(1);
    return sheet;
  },
  
  /**
   * Khởi tạo Sheet QUẢN LÝ NỢ
   */
  initializeDebtManagementSheet() {
    const ss = getSpreadsheet();
    const sheet = this._getOrCreateSheet(ss, APP_CONFIG.SHEETS.DEBT_MANAGEMENT);
    
    // Header
    const headers = [
      'STT', 'Tên khoản nợ', 'Loại hình', 'Nợ gốc ban đầu', 'Lãi suất (%/năm)', 
      'Kỳ hạn (tháng)', 'Ngày vay', 'Ngày đến hạn', 'Đã trả gốc', 
      'Đã trả lãi', 'Còn nợ', 'Trạng thái', 'Ghi chú', 'TransactionID'
    ];
    
    sheet.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBackground(APP_CONFIG.COLORS.HEADER_BG)
      .setFontColor(APP_CONFIG.COLORS.HEADER_TEXT);
    
    // Column widths
    sheet.setColumnWidth(1, 50);
    sheet.setColumnWidth(2, 150);
    sheet.setColumnWidth(3, 250); // Increased width for Type ID/Label
    sheet.setColumnWidth(4, 100);
    sheet.setColumnWidth(5, 100);
    sheet.setColumnWidth(6, 100);
    sheet.setColumnWidth(7, 100);
    sheet.setColumnWidth(8, 120);
    sheet.setColumnWidth(9, 120);
    sheet.setColumnWidth(10, 120);
    sheet.setColumnWidth(11, 100);
    sheet.setColumnWidth(12, 200);
    sheet.hideColumns(14);         // Hide TransactionID
    
    // Format - CRITICAL: Set BEFORE any data validation or formulas
    sheet.getRange('A2:A').setNumberFormat('0');
    sheet.getRange('C2:C').setNumberFormat('@'); // Type is text
    sheet.getRange('D2:D').setNumberFormat('#,##0');
    sheet.getRange('E2:E').setNumberFormat('0.00"%"');
    
    // CRITICAL: Force Kỳ hạn (Term) column to be NUMBER, not DATE
    const termRange = sheet.getRange('F2:F1000');
    termRange.setNumberFormat('0'); // Plain integer
    // Clear any existing format that might interfere
    this._fixTermColumn(sheet, 6); // Column F
    
    // Date columns
    sheet.getRange('G2:H').setNumberFormat(APP_CONFIG.FORMATS.DATE); // Start Date, Maturity Date
    this._fixDateColumn(sheet, 7); // Ngày vay
    this._fixDateColumn(sheet, 8); // Ngày đến hạn
    
    sheet.getRange('I2:K').setNumberFormat('#,##0');
    
    // Formula
    // K: Còn nợ = Gốc (D) - Đã trả gốc (I)
    sheet.getRange('K2:K1000').setFormula('=IFERROR(D2-I2, 0)');
    
    // Validation for Loan Types
    const typeRange = sheet.getRange('C2:C1000');
    // Use IDs from LOAN_TYPES
    const loanTypes = Object.keys(LOAN_TYPES);
    const typeRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(loanTypes)
      .setAllowInvalid(true) // Allow legacy values but warn
      .build();
    typeRange.setDataValidation(typeRule);
    
    // Validation for Status
    const statusRange = sheet.getRange('L2:L1000');
    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Chưa trả', 'Đang trả', 'Đã thanh toán', 'Quá hạn'])
      .setAllowInvalid(false)
      .build();
    statusRange.setDataValidation(statusRule);
    
    sheet.setFrozenRows(1);
    return sheet;
  },
  
  /**
   * Khởi tạo Sheet CHO VAY
   */
  initializeLendingSheet() {
    const ss = getSpreadsheet();
    const sheet = this._getOrCreateSheet(ss, APP_CONFIG.SHEETS.LENDING);
    
    // Header
    const headers = [
      'STT', 'Tên người vay', 'Loại hình', 'Số tiền gốc', 'Lãi suất (%/năm)', 
      'Kỳ hạn (tháng)', 'Ngày vay', 'Ngày đến hạn', 'Gốc đã thu', 'Lãi đã thu', 
      'Còn lại', 'Trạng thái', 'Ghi chú', 'TransactionID'
    ];
    
    sheet.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBackground(APP_CONFIG.COLORS.HEADER_BG)
      .setFontColor(APP_CONFIG.COLORS.HEADER_TEXT);
    
    // Column widths
    sheet.setColumnWidth(1, 50);
    sheet.setColumnWidth(2, 150);
    sheet.setColumnWidth(3, 250); // Increased width for Type ID/Label
    sheet.setColumnWidth(4, 100);
    sheet.setColumnWidth(5, 100);
    sheet.setColumnWidth(6, 100);
    sheet.setColumnWidth(7, 100);
    sheet.setColumnWidth(8, 120);
    sheet.setColumnWidth(9, 120);
    sheet.setColumnWidth(10, 120);
    sheet.setColumnWidth(11, 100);
    sheet.setColumnWidth(12, 200);
    sheet.hideColumns(14);         // Hide TransactionID
    
    // Format - CRITICAL: Set BEFORE any data validation or formulas
    sheet.getRange('A2:A').setNumberFormat('0');
    sheet.getRange('C2:C').setNumberFormat('@'); // Type is text
    sheet.getRange('D2:D').setNumberFormat('#,##0');
    sheet.getRange('E2:E').setNumberFormat('0.00"%"');
    
    // CRITICAL: Force Kỳ hạn (Term) column to be NUMBER, not DATE
    const termRange = sheet.getRange('F2:F1000');
    termRange.setNumberFormat('0'); // Plain integer
    // Clear any existing format that might interfere
    this._fixTermColumn(sheet, 6); // Column F
    
    // Date columns
    sheet.getRange('G2:H').setNumberFormat(APP_CONFIG.FORMATS.DATE); // Start Date, Maturity Date
    this._fixDateColumn(sheet, 7); // Ngày vay
    this._fixDateColumn(sheet, 8); // Ngày đến hạn
    
    sheet.getRange('I2:K').setNumberFormat('#,##0');
    
    // Formula
    // K: Còn lại = Gốc (D) - Gốc đã thu (I)
    sheet.getRange('K2:K1000').setFormula('=IFERROR(D2-I2, 0)');
    
    // Validation for Loan Types
    const typeRange = sheet.getRange('C2:C1000');
    // Use IDs from LOAN_TYPES
    const loanTypes = Object.keys(LOAN_TYPES);
    const typeRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(loanTypes)
      .setAllowInvalid(true) // Allow legacy values but warn
      .build();
    typeRange.setDataValidation(typeRule);
    
    // Validation for Status
    const statusRange = sheet.getRange('L2:L1000');
    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Đang vay', 'Đã tất toán', 'Quá hạn', 'Khó đòi'])
      .setAllowInvalid(false)
      .build();
    statusRange.setDataValidation(statusRule);
    
    sheet.setFrozenRows(1);
    return sheet;
  },
  
  /**
   * Khởi tạo Sheet THU NỢ
   */
  initializeLendingRepaymentSheet() {
    const ss = getSpreadsheet();
    const sheet = this._getOrCreateSheet(ss, APP_CONFIG.SHEETS.LENDING_REPAYMENT);
    
    // Header
    const headers = ['STT', 'Ngày', 'Người vay', 'Thu gốc', 'Thu lãi', 'Tổng thu', 'Ghi chú', 'TransactionID'];
    sheet.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBackground(APP_CONFIG.COLORS.HEADER_BG)
      .setFontColor(APP_CONFIG.COLORS.HEADER_TEXT);
    
    // Column widths
    sheet.setColumnWidth(1, 50);
    sheet.setColumnWidth(2, 100);
    sheet.setColumnWidth(3, 150);
    sheet.setColumnWidth(4, 120);
    sheet.setColumnWidth(5, 120);
    sheet.setColumnWidth(6, 120);
    sheet.setColumnWidth(7, 250);
    sheet.hideColumns(8);          // Hide TransactionID
    
    // Format
    sheet.getRange('A2:A').setNumberFormat('0');
    sheet.getRange('B2:B').setNumberFormat(APP_CONFIG.FORMATS.DATE);
    this._fixDateColumn(sheet, 2);
    sheet.getRange('D2:F').setNumberFormat('#,##0');
    
    // Formula (Safe to re-apply)
    sheet.getRange('F2:F1000').setFormula('=IFERROR(D2+E2, 0)');
    
    sheet.setFrozenRows(1);
    return sheet;
  },
  
  /**
   * Khởi tạo Sheet CHỨNG KHOÁN
   */
  initializeStockSheet() {
    const ss = getSpreadsheet();
    const sheet = this._getOrCreateSheet(ss, APP_CONFIG.SHEETS.STOCK);
    
    // Header
    const headers = [
      'STT', 'Ngày', 'Loại GD', 'Mã CK', 'Số lượng', 'Giá gốc', 'Phí', 
      'Tổng vốn', '💰 Cổ tức TM', '📈 Cổ tức CP', '📊 Giá ĐC', 
      '💹 Giá HT', '💵 Giá trị HT', '📈 Lãi/Lỗ', '📊 % L/L', 'Ghi chú'
    ];
    
    sheet.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBackground(APP_CONFIG.COLORS.HEADER_BG)
      .setFontColor(APP_CONFIG.COLORS.HEADER_TEXT);
    
    // Column widths
    sheet.setColumnWidth(1, 50);
    sheet.setColumnWidth(2, 100);
    sheet.setColumnWidth(3, 80);
    sheet.setColumnWidth(4, 80);
    sheet.setColumnWidth(5, 80);
    sheet.setColumnWidth(6, 100);
    sheet.setColumnWidth(7, 100);
    sheet.setColumnWidth(8, 120);
    sheet.setColumnWidth(9, 110);
    sheet.setColumnWidth(10, 100);
    sheet.setColumnWidth(11, 100);
    sheet.setColumnWidth(12, 100);
    sheet.setColumnWidth(13, 120);
    sheet.setColumnWidth(14, 110);
    sheet.setColumnWidth(15, 80);
    sheet.setColumnWidth(16, 250);
    
    // Format
    sheet.getRange('A2:A').setNumberFormat('0');
    sheet.getRange('B2:B').setNumberFormat(APP_CONFIG.FORMATS.DATE);
    this._fixDateColumn(sheet, 2);
    sheet.getRange('F2:H').setNumberFormat('#,##0');
    sheet.getRange('I2:I').setNumberFormat('#,##0');
    sheet.getRange('J2:J').setNumberFormat('0');
    sheet.getRange('K2:M').setNumberFormat('#,##0');
    sheet.getRange('N2:N').setNumberFormat('#,##0');
    sheet.getRange('O2:O').setNumberFormat('0.00%');
    
    // Formulas - REMOVED PRE-FILL to avoid getLastRow issues
    // Formulas are now set dynamically by DataProcessor.gs
    // sheet.getRange('K2:K1000').setFormula('=IF(E2>0, (H2-I2)/E2, 0)');
    // sheet.getRange('L2:L1000').setFormula('=IF(D2<>"", MPRICE(D2), 0)');
    // sheet.getRange('M2:M1000').setFormula('=IF(AND(E2>0, L2>0), E2*L2, 0)');
    // sheet.getRange('N2:N1000').setFormula('=IF(M2>0, M2-(H2-I2), 0)');
    // sheet.getRange('O2:O1000').setFormula('=IF(AND(N2<>0, (H2-I2)>0), N2/(H2-I2), 0)');
    
    // Conditional Formatting
    sheet.clearConditionalFormatRules();
    const profitLossRange = sheet.getRange('N2:N1000');
    const percentRange = sheet.getRange('O2:O1000');
    
    const rules = [
      SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0).setBackground('#D4EDDA').setFontColor('#155724').setRanges([profitLossRange]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0).setBackground('#F8D7DA').setFontColor('#721C24').setRanges([profitLossRange]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0.1).setBackground('#28A745').setFontColor('#FFFFFF').setBold(true).setRanges([percentRange]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(0, 0.1).setBackground('#D4EDDA').setFontColor('#155724').setRanges([percentRange]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(-0.1, 0).setBackground('#F8D7DA').setFontColor('#721C24').setRanges([percentRange]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(-0.1).setBackground('#DC3545').setFontColor('#FFFFFF').setBold(true).setRanges([percentRange]).build()
    ];
    sheet.setConditionalFormatRules(rules);
    
    sheet.setFrozenRows(1);
    
    // Validation
    const typeRange = sheet.getRange('C2:C1000');
    const typeRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Mua', 'Bán', 'Thưởng'])
      .setAllowInvalid(false)
      .build();
    typeRange.setDataValidation(typeRule);
    
    return sheet;
  },
  
  /**
   * Khởi tạo Sheet VÀNG
   */
  initializeGoldSheet() {
    const ss = getSpreadsheet();
    const sheet = this._getOrCreateSheet(ss, APP_CONFIG.SHEETS.GOLD);
    
    // Header
    // [NEW] Thêm cột Tài sản, Giá vốn, Tổng vốn, Giá HT, Giá trị HT, Lãi/Lỗ
    const headers = [
      'STT', 'Ngày', 'Tài sản', 'Loại GD', 'Loại vàng', 'Số lượng', 'Đơn vị', 
      'Giá vốn', 'Tổng vốn', 'Giá HT', 'Giá trị HT', 'Lãi/Lỗ', '% Lãi/Lỗ', 'Ghi chú'
    ];
    
    sheet.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBackground(APP_CONFIG.COLORS.HEADER_BG)
      .setFontColor(APP_CONFIG.COLORS.HEADER_TEXT);
    
    // Column widths
    sheet.setColumnWidth(1, 50);   // STT
    sheet.setColumnWidth(2, 100);  // Ngày
    sheet.setColumnWidth(3, 80);   // Tài sản
    sheet.setColumnWidth(4, 80);   // Loại GD
    sheet.setColumnWidth(5, 100);  // Loại vàng
    sheet.setColumnWidth(6, 80);   // Số lượng
    sheet.setColumnWidth(7, 70);   // Đơn vị
    sheet.setColumnWidth(8, 100);  // Giá vốn
    sheet.setColumnWidth(9, 120);  // Tổng vốn
    sheet.setColumnWidth(10, 100); // Giá HT
    sheet.setColumnWidth(11, 120); // Giá trị HT
    sheet.setColumnWidth(12, 110); // Lãi/Lỗ
    sheet.setColumnWidth(13, 80);  // % Lãi/Lỗ
    sheet.setColumnWidth(14, 200); // Ghi chú
    
    // Format
    sheet.getRange('A2:A').setNumberFormat('0');
    sheet.getRange('B2:B').setNumberFormat(APP_CONFIG.FORMATS.DATE);
    this._fixDateColumn(sheet, 2);
    sheet.getRange('H2:L').setNumberFormat('#,##0'); // Giá vốn -> Lãi/Lỗ
    sheet.getRange('M2:M').setNumberFormat('0.00%');
    
    // Formulas - REMOVED PRE-FILL
    // Formulas are now set dynamically by DataProcessor.gs
    // J: Giá HT = GPRICE(Loại vàng - Cột E)
    // sheet.getRange('J2:J1000').setFormula('=IF(E2<>"", GPRICE(E2), 0)');
    
    // K: Giá trị HT = Số lượng * Giá HT
    // sheet.getRange('K2:K1000').setFormula('=IF(AND(F2>0, J2>0), F2*J2, 0)');
    
    // L: Lãi/Lỗ = Giá trị HT - Tổng vốn
    // sheet.getRange('L2:L1000').setFormula('=IF(K2>0, K2-I2, 0)');
    
    // M: % Lãi/Lỗ
    // sheet.getRange('M2:M1000').setFormula('=IF(I2>0, L2/I2, 0)');
    
    // Conditional Formatting for Profit/Loss
    sheet.clearConditionalFormatRules();
    const profitLossRange = sheet.getRange('L2:L1000');
    const percentRange = sheet.getRange('M2:M1000');
    
    const rules = [
      SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0).setBackground('#D4EDDA').setFontColor('#155724').setRanges([profitLossRange]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0).setBackground('#F8D7DA').setFontColor('#721C24').setRanges([profitLossRange]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0.1).setBackground('#28A745').setFontColor('#FFFFFF').setBold(true).setRanges([percentRange]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(0, 0.1).setBackground('#D4EDDA').setFontColor('#155724').setRanges([percentRange]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(-0.1, 0).setBackground('#F8D7DA').setFontColor('#721C24').setRanges([percentRange]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(-0.1).setBackground('#DC3545').setFontColor('#FFFFFF').setBold(true).setRanges([percentRange]).build()
    ];
    sheet.setConditionalFormatRules(rules);
    
    sheet.setFrozenRows(1);
    
    // Validations
    sheet.getRange('D2:D1000').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Mua', 'Bán']).build());
    sheet.getRange('E2:E1000').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['SJC', '24K', '18K', '14K', '10K', 'Khác']).build());
    sheet.getRange('G2:G1000').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['chỉ', 'lượng', 'cây', 'gram']).build());
    
    return sheet;
  },
  
  /**
   * Khởi tạo Sheet CRYPTO
   */
  initializeCryptoSheet() {
    const ss = getSpreadsheet();
    const sheet = this._getOrCreateSheet(ss, APP_CONFIG.SHEETS.CRYPTO);
    
    // Header
    // [NEW] Thêm cột Giá HT (USD), Giá trị HT (USD), Giá HT (VND), Giá trị HT (VND), Lãi/Lỗ
    const headers = [
      'STT', 'Ngày', 'Loại GD', 'Coin', 'Số lượng', 'Giá (USD)', 'Tỷ giá', 'Giá (VND)', 'Tổng vốn',
      'Giá HT (USD)', 'Giá trị HT (USD)', 'Giá HT (VND)', 'Giá trị HT (VND)', 'Lãi/Lỗ', '% Lãi/Lỗ',
      'Sàn', 'Ví', 'Ghi chú'
    ];
    
    sheet.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBackground(APP_CONFIG.COLORS.HEADER_BG)
      .setFontColor(APP_CONFIG.COLORS.HEADER_TEXT);
    
    // Column widths
    sheet.setColumnWidth(1, 50);   // STT
    sheet.setColumnWidth(2, 100);  // Ngày
    sheet.setColumnWidth(3, 80);   // Loại GD
    sheet.setColumnWidth(4, 80);   // Coin
    sheet.setColumnWidth(5, 100);  // Số lượng
    sheet.setColumnWidth(6, 100);  // Giá (USD)
    sheet.setColumnWidth(7, 80);   // Tỷ giá
    sheet.setColumnWidth(8, 100);  // Giá (VND)
    sheet.setColumnWidth(9, 120);  // Tổng vốn
    sheet.setColumnWidth(10, 100); // Giá HT (USD)
    sheet.setColumnWidth(11, 120); // Giá trị HT (USD)
    sheet.setColumnWidth(12, 100); // Giá HT (VND)
    sheet.setColumnWidth(13, 120); // Giá trị HT (VND)
    sheet.setColumnWidth(14, 110); // Lãi/Lỗ
    sheet.setColumnWidth(15, 80);  // % Lãi/Lỗ
    sheet.setColumnWidth(16, 100); // Sàn
    sheet.setColumnWidth(17, 150); // Ví
    sheet.setColumnWidth(18, 200); // Ghi chú
    
    // Format
    sheet.getRange('A2:A').setNumberFormat('0');
    sheet.getRange('B2:B').setNumberFormat(APP_CONFIG.FORMATS.DATE);
    this._fixDateColumn(sheet, 2);
    sheet.getRange('F2:F').setNumberFormat('#,##0.00'); // Giá USD
    sheet.getRange('G2:G').setNumberFormat('#,##0');    // Tỷ giá
    sheet.getRange('H2:I').setNumberFormat('#,##0');    // Giá VND, Tổng vốn
    sheet.getRange('J2:J').setNumberFormat('#,##0.00'); // Giá HT USD
    sheet.getRange('K2:K').setNumberFormat('#,##0.00'); // Giá trị HT USD
    sheet.getRange('L2:N').setNumberFormat('#,##0');    // Giá HT VND -> Lãi/Lỗ
    sheet.getRange('O2:O').setNumberFormat('0.00%');    // % Lãi/Lỗ
    
    // Formulas - REMOVED PRE-FILL
    // Formulas are now set dynamically by DataProcessor.gs
    // J: Giá HT (USD) = CPRICE(Coin + "USD")
    // sheet.getRange('J2:J1000').setFormula('=IF(D2<>"", CPRICE(D2&"USD"), 0)');
    
    // K: Giá trị HT (USD) = Số lượng * Giá HT (USD)
    // sheet.getRange('K2:K1000').setFormula('=IF(AND(E2>0, J2>0), E2*J2, 0)');
    
    // L: Giá HT (VND) = Giá HT (USD) * Tỷ giá (Cột G)
    // sheet.getRange('L2:L1000').setFormula('=IF(AND(J2>0, G2>0), J2*G2, 0)');
    
    // M: Giá trị HT (VND) = Giá trị HT (USD) * Tỷ giá
    // sheet.getRange('M2:M1000').setFormula('=IF(AND(K2>0, G2>0), K2*G2, 0)');
    
    // N: Lãi/Lỗ = Giá trị HT (VND) - Tổng vốn
    // sheet.getRange('N2:N1000').setFormula('=IF(M2>0, M2-I2, 0)');
    
    // O: % Lãi/Lỗ
    // sheet.getRange('O2:O1000').setFormula('=IF(I2>0, N2/I2, 0)');
    
    // Conditional Formatting
    sheet.clearConditionalFormatRules();
    const profitLossRange = sheet.getRange('N2:N1000');
    const percentRange = sheet.getRange('O2:O1000');
    
    const rules = [
      SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0).setBackground('#D4EDDA').setFontColor('#155724').setRanges([profitLossRange]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0).setBackground('#F8D7DA').setFontColor('#721C24').setRanges([profitLossRange]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0.1).setBackground('#28A745').setFontColor('#FFFFFF').setBold(true).setRanges([percentRange]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(0, 0.1).setBackground('#D4EDDA').setFontColor('#155724').setRanges([percentRange]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(-0.1, 0).setBackground('#F8D7DA').setFontColor('#721C24').setRanges([percentRange]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(-0.1).setBackground('#DC3545').setFontColor('#FFFFFF').setBold(true).setRanges([percentRange]).build()
    ];
    sheet.setConditionalFormatRules(rules);
    
    sheet.setFrozenRows(1);
    
    // Validation
    sheet.getRange('C2:C1000').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Mua', 'Bán', 'Swap', 'Stake', 'Unstake']).build());
    
    return sheet;
  },
  
  /**
   * Khởi tạo Sheet ĐẦU TƯ KHÁC
   */
  initializeOtherInvestmentSheet() {
    const ss = getSpreadsheet();
    const sheet = this._getOrCreateSheet(ss, APP_CONFIG.SHEETS.OTHER_INVESTMENT);
    
    // Header
    const headers = ['STT', 'Ngày', 'Loại đầu tư', 'Số tiền', 'Lãi suất (%)', 'Kỳ hạn (tháng)', 'Dự kiến thu về', 'Ghi chú'];
    sheet.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBackground(APP_CONFIG.COLORS.HEADER_BG)
      .setFontColor(APP_CONFIG.COLORS.HEADER_TEXT);
    
    // Column widths
    sheet.setColumnWidth(1, 50);
    sheet.setColumnWidth(2, 100);
    sheet.setColumnWidth(3, 150);
    sheet.setColumnWidth(4, 120);
    sheet.setColumnWidth(5, 100);
    sheet.setColumnWidth(6, 100);
    sheet.setColumnWidth(7, 120);
    sheet.setColumnWidth(8, 250);
    
    // Format
    sheet.getRange('A2:A').setNumberFormat('0');
    sheet.getRange('B2:B').setNumberFormat(APP_CONFIG.FORMATS.DATE);
    this._fixDateColumn(sheet, 2);
    sheet.getRange('D2:D').setNumberFormat('#,##0');
    sheet.getRange('E2:E').setNumberFormat('0.00"%"');
    sheet.getRange('G2:G').setNumberFormat('#,##0');
    
    sheet.setFrozenRows(1);
    
    // Validation
    sheet.getRange('C2:C1000').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Gửi tiết kiệm', 'Quỹ đầu tư', 'Bất động sản', 'Trái phiếu', 'P2P Lending', 'Khác']).build());
    
    return sheet;
  },

  /**
   * Khởi tạo Sheet CHANGELOG
   */
  initializeChangelogSheet() {
    const ss = getSpreadsheet();
    const sheet = this._getOrCreateSheet(ss, APP_CONFIG.SHEETS.CHANGELOG);
    
    // Header
    const headers = ['Phiên bản / Tính năng', 'Chi tiết thay đổi', 'Hành động khuyến nghị'];
    sheet.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBackground(APP_CONFIG.COLORS.HEADER_BG)
      .setFontColor(APP_CONFIG.COLORS.HEADER_TEXT);
    
    // Column widths
    sheet.setColumnWidth(1, 250); // Phiên bản
    sheet.setColumnWidth(2, 400); // Chi tiết
    sheet.setColumnWidth(3, 300); // Hành động
    
    // Format
    sheet.setFrozenRows(1);
    sheet.getRange('A:C').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    
    return sheet;
  },
  
  /**
   * Khởi tạo Sheet BUDGET
   */
  initializeBudgetSheet() {
    const ss = getSpreadsheet();
    const sheet = this._getOrCreateSheet(ss, APP_CONFIG.SHEETS.BUDGET);
    
    // Lấy tháng/năm hiện tại
    const now = new Date();
    const currentMonth = now.getMonth() + 1;
    const currentYear = now.getFullYear();
    
    // ========== ROW 1: HEADER ==========
    sheet.getRange('A1:F1').merge()
      .setValue(`💰 NGÂN SÁCH THÁNG ${currentMonth}/${currentYear}`)
      .setFontWeight('bold')
      .setFontSize(14)
      .setHorizontalAlignment('center')
      .setBackground('#4472C4')
      .setFontColor('#FFFFFF');
    
    // ========== ROW 2: THU NHẬP DỰ KIẾN ==========
    sheet.getRange('A2').setValue('Thu nhập dự kiến:')
      .setFontWeight('bold')
      .setHorizontalAlignment('right');
    sheet.getRange('B2').setNumberFormat('#,##0');
    
    // ========== ROW 3-5: PHÂN BỔ NGÂN SÁCH (60/25/15) ==========
    sheet.getRange('A3').setValue('Nhóm Chi tiêu:')
      .setFontWeight('bold')
      .setHorizontalAlignment('right')
      .setBackground('#EA4335')
      .setFontColor('#FFFFFF');
    sheet.getRange('B3').setNumberFormat('0.00%').setValue(0.6); // Default 60%
    
    sheet.getRange('A4').setValue('Nhóm Đầu tư:')
      .setFontWeight('bold')
      .setHorizontalAlignment('right')
      .setBackground('#34A853')
      .setFontColor('#FFFFFF');
    sheet.getRange('B4').setNumberFormat('0.00%').setValue(0.25); // Default 25%
    
    sheet.getRange('A5').setValue('Nhóm Trả nợ:')
      .setFontWeight('bold')
      .setHorizontalAlignment('right')
      .setBackground('#FBBC04')
      .setFontColor('#FFFFFF');
    sheet.getRange('B5').setNumberFormat('0.00%').setValue(0.15); // Default 15%
    
    // Validation: Sum must be 100%
    sheet.getRange('C3').setFormula('=IF(ROUND(SUM(B3:B5), 2)<>1, "⚠️ Tổng phải là 100%", "")');
    
    // ========== TABLE HEADERS (Row 6) ==========
    const headers = ['Danh mục', '% Nhóm', 'Ngân sách', 'Đã chi', 'Còn lại', 'Trạng thái'];
    sheet.getRange('A6:F6').setValues([headers])
      .setFontWeight('bold')
      .setBackground('#4472C4')
      .setFontColor('#FFFFFF')
      .setHorizontalAlignment('center');
      
    // ========== GROUP 1: CHI TIÊU ==========
    // Section Header
    sheet.getRange('A7:F7').merge().setValue('📛 CHI TIÊU')
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBackground('#EA4335')
      .setFontColor('#FFFFFF');

    // Data Rows
    // Filter out 'Trả nợ' and 'Cho vay' from Expense Categories if they exist, 
    // as they have their own sections or logic.
    // User requested "Trả nợ" to be separate.
    const expenseCats = APP_CONFIG.CATEGORIES.EXPENSE.filter(cat => cat !== 'Trả nợ' && cat !== 'Cho vay' && cat !== 'Đầu tư');
    
    expenseCats.forEach((cat, i) => {
      const r = 8 + i;
      sheet.getRange(r, 1, 1, 6).breakApart();
      sheet.getRange(r, 1).setValue(cat);
      
      // Formula: Budget = Income * Group% * Category%
      sheet.getRange(r, 3).setFormula('=IF(B' + r + '<>"", $B$2 * $B$3 * B' + r + ', 0)');
      
      // Formula: Spent (SUMIFS from CHI sheet)
      sheet.getRange(r, 4).setFormulaR1C1(
        '=IFERROR(SUMIFS(CHI!C3, CHI!C4, RC[-3], CHI!C2, ">="&DATE(YEAR(TODAY()), MONTH(TODAY()), 1), CHI!C2, "<"&DATE(YEAR(TODAY()), MONTH(TODAY())+1, 1)), 0)'
      );
      
      // Formula: Status
      const statusFormula = '=IF(D' + r + '=0, "⚪ Chưa chi", IF(D' + r + '>C' + r + ', "🔴 Vượt ngân sách", IF(D' + r + '/C' + r + '>=0.8, "⚠️ Sắp hết", "✅ Trong hạn mức")))';
      sheet.getRange(r, 6).setFormula(statusFormula);
    });
    
    const expEndRow = 8 + expenseCats.length;
    
    // Total Expense Row
    sheet.getRange(expEndRow, 1).setValue('TỔNG CHI').setFontWeight('bold');
    sheet.getRange(expEndRow, 2).setFormula(`=SUM(B8:B${expEndRow-1})`).setNumberFormat('0.00%');
    sheet.getRange(expEndRow, 3).setFormula(`=SUM(C8:C${expEndRow-1})`);
    sheet.getRange(expEndRow, 4).setFormula(`=SUM(D8:D${expEndRow-1})`);
    sheet.getRange(expEndRow, 5).setFormula(`=C${expEndRow}-D${expEndRow}`);
    sheet.getRange(expEndRow, 6).setFormula(`=IF(B${expEndRow}<>1, "⚠️ Tổng % phải là 100%", "✅ OK")`);
    
    // ========== GROUP 2: ĐẦU TƯ ==========
    const investStartRow = expEndRow + 2;
    
    // Section Header
    sheet.getRange(investStartRow, 1, 1, 6).merge().setValue('💰 ĐẦU TƯ')
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBackground('#34A853')
      .setFontColor('#FFFFFF');
      
    // Data Rows
    const investCats = APP_CONFIG.CATEGORIES.INVESTMENT;
    
    investCats.forEach((cat, i) => {
      const r = investStartRow + 1 + i;
      sheet.getRange(r, 1).setValue(cat);
      
      // Formula: Budget = Income * Group% * Category%
      sheet.getRange(r, 3).setFormula('=IF(B' + r + '<>"", $B$2 * $B$4 * B' + r + ', 0)');
      
      // Formula: Spent (Specific logic per type)
      if (cat === 'Chứng khoán') {
        sheet.getRange(r, 4).setFormula(
          '=IFERROR(SUMIFS(\'CHỨNG KHOÁN\'!H:H, \'CHỨNG KHOÁN\'!C:C, "Mua", \'CHỨNG KHOÁN\'!B:B, ">="&DATE(YEAR(TODAY()), MONTH(TODAY()), 1), \'CHỨNG KHOÁN\'!B:B, "<"&DATE(YEAR(TODAY()), MONTH(TODAY())+1, 1)), 0)'
        );
      } else if (cat === 'Vàng') {
        sheet.getRange(r, 4).setFormula(
          '=IFERROR(SUMIFS(\'VÀNG\'!I:I, \'VÀNG\'!D:D, "Mua", \'VÀNG\'!B:B, ">="&DATE(YEAR(TODAY()), MONTH(TODAY()), 1), \'VÀNG\'!B:B, "<"&DATE(YEAR(TODAY()), MONTH(TODAY())+1, 1)), 0)'
        );
      } else if (cat === 'Crypto') {
        sheet.getRange(r, 4).setFormula(
          '=IFERROR(SUMIFS(\'CRYPTO\'!I:I, \'CRYPTO\'!C:C, "Mua", \'CRYPTO\'!B:B, ">="&DATE(YEAR(TODAY()), MONTH(TODAY()), 1), \'CRYPTO\'!B:B, "<"&DATE(YEAR(TODAY()), MONTH(TODAY())+1, 1)), 0)'
        );
      } else if (cat === 'Đầu tư khác') {
        sheet.getRange(r, 4).setFormula(
          '=IFERROR(SUMIFS(\'ĐẦU TƯ KHÁC\'!D:D, \'ĐẦU TƯ KHÁC\'!B:B, ">="&DATE(YEAR(TODAY()), MONTH(TODAY()), 1), \'ĐẦU TƯ KHÁC\'!B:B, "<"&DATE(YEAR(TODAY()), MONTH(TODAY())+1, 1)), 0)'
        );
      } else {
        sheet.getRange(r, 4).setValue(0); // Unknown type
      }
      
      // Formula: Status
      const statusFormula = '=IF(D' + r + '=0, "⚪ Chưa chi", IF(D' + r + '>C' + r + ', "🔴 Vượt ngân sách", IF(D' + r + '/C' + r + '>=0.8, "⚠️ Sắp hết", "✅ Trong hạn mức")))';
      sheet.getRange(r, 6).setFormula(statusFormula);
    });
    
    const investEndRow = investStartRow + 1 + investCats.length;
    
    // Total Investment Row
    sheet.getRange(investEndRow, 1).setValue('TỔNG ĐT').setFontWeight('bold');
    sheet.getRange(investEndRow, 2).setFormula(`=SUM(B${investStartRow+1}:B${investEndRow-1})`).setNumberFormat('0.00%');
    sheet.getRange(investEndRow, 3).setFormula(`=SUM(C${investStartRow+1}:C${investEndRow-1})`);
    sheet.getRange(investEndRow, 4).setFormula(`=SUM(D${investStartRow+1}:D${investEndRow-1})`);
    sheet.getRange(investEndRow, 5).setFormula(`=C${investEndRow}-D${investEndRow}`);
    sheet.getRange(investEndRow, 6).setFormula(`=IF(B${investEndRow}<>1, "⚠️ Tổng % phải là 100%", "✅ OK")`);

    // ========== GROUP 3: TRẢ NỢ ==========
    const debtStartRow = investEndRow + 2;
    
    // Section Header
    sheet.getRange(debtStartRow, 1, 1, 6).merge().setValue('💳 TRẢ NỢ')
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBackground('#FBBC04')
      .setFontColor('#FFFFFF');
      
    // Data Rows
    const debtRow = debtStartRow + 1;
    sheet.getRange(debtRow, 1).setValue('Trả nợ');
    sheet.getRange(debtRow, 2).setValue(1).setNumberFormat('0.00%'); // 100% of Debt Group
    
    // Formula: Budget
    sheet.getRange(debtRow, 3).setFormula('=IF(B' + debtRow + '<>"", $B$2 * $B$5 * B' + debtRow + ', 0)');
    
    // Formula: Spent
    sheet.getRange(debtRow, 4).setFormula(
      '=IFERROR(SUMIFS(\'TRẢ NỢ\'!F:F, \'TRẢ NỢ\'!B:B, ">="&DATE(YEAR(TODAY()), MONTH(TODAY()), 1), \'TRẢ NỢ\'!B:B, "<"&DATE(YEAR(TODAY()), MONTH(TODAY())+1, 1)), 0)'
    );
    
    // Formula: Status
    const debtStatusFormula = '=IF(D' + debtRow + '=0, "⚪ Chưa chi", IF(D' + debtRow + '>C' + debtRow + ', "🔴 Vượt ngân sách", IF(D' + debtRow + '/C' + debtRow + '>=0.8, "⚠️ Sắp hết", "✅ Trong hạn mức")))';
    sheet.getRange(debtRow, 6).setFormula(debtStatusFormula);
    
    // Total Debt Row
    const debtEndRow = debtRow + 1;
    sheet.getRange(debtEndRow, 1).setValue('TỔNG TRẢ NỢ').setFontWeight('bold');
    sheet.getRange(debtEndRow, 2).setFormula(`=SUM(B${debtRow})`).setNumberFormat('0.00%');
    sheet.getRange(debtEndRow, 3).setFormula(`=SUM(C${debtRow})`);
    sheet.getRange(debtEndRow, 4).setFormula(`=SUM(D${debtRow})`);
    sheet.getRange(debtEndRow, 5).setFormula(`=C${debtEndRow}-D${debtEndRow}`);
    
    // Sync Warning for Debt
    sheet.getRange(debtEndRow, 6).setFormula(
      '=IF(C' + debtEndRow + ' < IFERROR(SUM(\'QUẢN LÝ NỢ\'!G:G), 0), "⚠️ Thấp hơn thực tế", "✅ OK")'
    );

    // ========== COMMON FORMULAS ==========
    // Calculate Remaining (Col E) = Budget (C) - Spent (D)
    // Applied to all rows
    const lastRow = debtEndRow;
    sheet.getRange('E8:E' + lastRow).setFormula('=IF(C8>0, C8-D8, 0)');
    
    // ========== FORMATTING ==========
    // Col B: %
    sheet.getRange('B8:B' + lastRow).setNumberFormat('0.00%');
    
    // Col C, D, E: Number (NO CURRENCY)
    sheet.getRange('C8:E' + (lastRow + 1)).setNumberFormat('#,##0');
    
    // Widths
    sheet.setColumnWidth(1, 200); // Danh mục
    sheet.setColumnWidth(2, 80);  // % Nhóm
    sheet.setColumnWidth(3, 120); // Ngân sách
    sheet.setColumnWidth(4, 120); // Đã chi
    sheet.setColumnWidth(5, 120); // Còn lại
    sheet.setColumnWidth(6, 150); // Trạng thái
    
    return sheet;
  }
};