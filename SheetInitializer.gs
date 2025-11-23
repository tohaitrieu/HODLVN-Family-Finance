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

      if (dateObj) {
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
    
    // Format
    sheet.getRange('A2:A').setNumberFormat('0');
    sheet.getRange('C2:C').setNumberFormat('@'); // Type is text
    sheet.getRange('D2:D').setNumberFormat('#,##0');
    sheet.getRange('E2:E').setNumberFormat('0.00"%"');
    sheet.getRange('F2:F').setNumberFormat('0'); // Term (Month) - Number
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
    
    // Format
    sheet.getRange('A2:A').setNumberFormat('0');
    sheet.getRange('C2:C').setNumberFormat('@'); // Type is text
    sheet.getRange('D2:D').setNumberFormat('#,##0');
    sheet.getRange('E2:E').setNumberFormat('0.00"%"');
    sheet.getRange('F2:F').setNumberFormat('0'); // Term - Number
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
    
    // Formulas
    sheet.getRange('K2:K1000').setFormula('=IF(E2>0, (H2-I2)/E2, 0)');
    sheet.getRange('L2:L1000').setFormula('=IF(D2<>"", MPRICE(D2), 0)');
    sheet.getRange('M2:M1000').setFormula('=IF(AND(E2>0, L2>0), E2*L2, 0)');
    sheet.getRange('N2:N1000').setFormula('=IF(M2>0, M2-(H2-I2), 0)');
    sheet.getRange('O2:O1000').setFormula('=IF(AND(N2<>0, (H2-I2)>0), N2/(H2-I2), 0)');
    
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
    
    // Formulas
    // J: Giá HT = GPRICE(Tài sản - Cột C)
    sheet.getRange('J2:J1000').setFormula('=IF(C2<>"", GPRICE(C2), 0)');
    
    // K: Giá trị HT = Số lượng * Giá HT
    // Lưu ý: GPRICE trả về giá VND (thường là cho 1 lượng/chỉ tùy loại). 
    // Giả định Số lượng và Giá HT tương thích đơn vị.
    sheet.getRange('K2:K1000').setFormula('=IF(AND(F2>0, J2>0), F2*J2, 0)');
    
    // L: Lãi/Lỗ = Giá trị HT - Tổng vốn
    sheet.getRange('L2:L1000').setFormula('=IF(K2>0, K2-I2, 0)');
    
    // M: % Lãi/Lỗ
    sheet.getRange('M2:M1000').setFormula('=IF(I2>0, L2/I2, 0)');
    
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
    
    // Formulas
    // J: Giá HT (USD) = CPRICE(Coin)
    // User tự nhập mã (VD: BTC-USD, ETH-USD) hoặc mã Yahoo Finance
    sheet.getRange('J2:J1000').setFormula('=IF(D2<>"", CPRICE(D2), 0)');
    
    // K: Giá trị HT (USD) = Số lượng * Giá HT (USD)
    sheet.getRange('K2:K1000').setFormula('=IF(AND(E2>0, J2>0), E2*J2, 0)');
    
    // L: Giá HT (VND) = Giá HT (USD) * Tỷ giá (Cột G)
    sheet.getRange('L2:L1000').setFormula('=IF(AND(J2>0, G2>0), J2*G2, 0)');
    
    // M: Giá trị HT (VND) = Giá trị HT (USD) * Tỷ giá
    sheet.getRange('M2:M1000').setFormula('=IF(AND(K2>0, G2>0), K2*G2, 0)');
    
    // N: Lãi/Lỗ = Giá trị HT (VND) - Tổng vốn
    sheet.getRange('N2:N1000').setFormula('=IF(M2>0, M2-I2, 0)');
    
    // O: % Lãi/Lỗ
    sheet.getRange('O2:O1000').setFormula('=IF(I2>0, N2/I2, 0)');
    
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
    
    // ========== ROW 3-5: PHÂN BỔ NGÂN SÁCH ==========
    sheet.getRange('A3').setValue('% Chi tiêu:')
      .setFontWeight('bold')
      .setHorizontalAlignment('right');
    sheet.getRange('B3').setNumberFormat('0%').setValue(0.5); // Default 50%
    
    sheet.getRange('A4').setValue('% Đầu tư:')
      .setFontWeight('bold')
      .setHorizontalAlignment('right');
    sheet.getRange('B4').setNumberFormat('0%').setValue(0.3); // Default 30%
    
    sheet.getRange('A5').setValue('% Trả nợ:')
      .setFontWeight('bold')
      .setHorizontalAlignment('right');
    sheet.getRange('B5').setNumberFormat('0%').setValue(0.2); // Default 20%
    
    // Validation: Sum must be 100%
    sheet.getRange('C3').setFormula('=IF(SUM(B3:B5)<>1, "⚠️ Tổng phải là 100%", "✅ OK")');
    
    // ========== TABLE HEADERS (Row 6) ==========
    const headers = ['Danh mục', 'Mục tiêu', 'Thực tế', 'Còn lại', 'Trạng thái', 'Ghi chú'];
    sheet.getRange('A6:F6').setValues([headers])
      .setFontWeight('bold')
      .setBackground('#EEEEEE')
      .setHorizontalAlignment('center');
      
    // ========== GROUP 1: CHI TIÊU (Row 7-16) ==========
    const expenseCats = [
      'Ăn uống', 'Đi lại', 'Nhà ở', 'Điện nước', 'Viễn thông',
      'Giáo dục', 'Y tế', 'Mua sắm', 'Giải trí', 'Khác'
    ];
    
    expenseCats.forEach((cat, i) => {
      sheet.getRange(7 + i, 1).setValue(cat);
    });
    
    // Total Expense Row (17)
    sheet.getRange('A17').setValue('TỔNG CHI TIÊU').setFontWeight('bold');
    sheet.getRange('B17').setFormula('=SUM(B7:B16)');
    sheet.getRange('C17').setFormula('=SUM(C7:C16)');
    sheet.getRange('D17').setFormula('=B17-C17');
    
    // ========== GROUP 2: ĐẦU TƯ (Row 19-23) - MOVED UP ==========
    sheet.getRange('A19').setValue('ĐẦU TƯ').setFontWeight('bold').setBackground('#D4EDDA');
    sheet.getRange('A20').setValue('Chứng khoán');
    sheet.getRange('A21').setValue('Vàng');
    sheet.getRange('A22').setValue('Crypto');
    sheet.getRange('A23').setValue('Đầu tư khác');
    
    // Total Investment Row (24)
    sheet.getRange('A24').setValue('TỔNG ĐẦU TƯ').setFontWeight('bold');
    sheet.getRange('B24').setFormula('=SUM(B20:B23)');
    sheet.getRange('C24').setFormula('=SUM(C20:C23)');
    sheet.getRange('D24').setFormula('=B24-C24');

    // ========== GROUP 3: TRẢ NỢ (Row 26) - MOVED DOWN ==========
    sheet.getRange('A26').setValue('TRẢ NỢ').setFontWeight('bold').setBackground('#F8D7DA');
    sheet.getRange('A27').setValue('Trả nợ (Gốc + Lãi)');
    
    // Sync Warning for Debt
    sheet.getRange('E27').setFormula(
      '=IF(B27 < IFERROR(SUM(\'QUẢN LÝ NỢ\'!G:G), 0), "⚠️ Thấp hơn thực tế", "✅ OK")'
    );

    // ========== FORMULAS FOR ACTUAL (Column C) ==========
    // Expense
    sheet.getRange('C7:C16').setFormulaR1C1(
      '=IFERROR(SUMIFS(CHI!C3, CHI!C4, RC[-2], CHI!C2, ">="&DATE(YEAR(TODAY()), MONTH(TODAY()), 1), CHI!C2, "<"&DATE(YEAR(TODAY()), MONTH(TODAY())+1, 1)), 0)'
    );
    
    // Investment
    // Stock
    sheet.getRange('C20').setFormula(
      '=IFERROR(SUMIFS(\'CHỨNG KHOÁN\'!H:H, \'CHỨNG KHOÁN\'!C:C, "Mua", \'CHỨNG KHOÁN\'!B:B, ">="&DATE(YEAR(TODAY()), MONTH(TODAY()), 1), \'CHỨNG KHOÁN\'!B:B, "<"&DATE(YEAR(TODAY()), MONTH(TODAY())+1, 1)), 0)'
    );
    // Gold
    sheet.getRange('C21').setFormula(
      '=IFERROR(SUMIFS(\'VÀNG\'!I:I, \'VÀNG\'!D:D, "Mua", \'VÀNG\'!B:B, ">="&DATE(YEAR(TODAY()), MONTH(TODAY()), 1), \'VÀNG\'!B:B, "<"&DATE(YEAR(TODAY()), MONTH(TODAY())+1, 1)), 0)'
    );
    // Crypto
    sheet.getRange('C22').setFormula(
      '=IFERROR(SUMIFS(\'CRYPTO\'!I:I, \'CRYPTO\'!C:C, "Mua", \'CRYPTO\'!B:B, ">="&DATE(YEAR(TODAY()), MONTH(TODAY()), 1), \'CRYPTO\'!B:B, "<"&DATE(YEAR(TODAY()), MONTH(TODAY())+1, 1)), 0)'
    );
    // Other
    sheet.getRange('C23').setFormula(
      '=IFERROR(SUMIFS(\'ĐẦU TƯ KHÁC\'!D:D, \'ĐẦU TƯ KHÁC\'!B:B, ">="&DATE(YEAR(TODAY()), MONTH(TODAY()), 1), \'ĐẦU TƯ KHÁC\'!B:B, "<"&DATE(YEAR(TODAY()), MONTH(TODAY())+1, 1)), 0)'
    );
    
    // Debt
    sheet.getRange('C27').setFormula(
      '=IFERROR(SUMIFS(\'TRẢ NỢ\'!F:F, \'TRẢ NỢ\'!B:B, ">="&DATE(YEAR(TODAY()), MONTH(TODAY()), 1), \'TRẢ NỢ\'!B:B, "<"&DATE(YEAR(TODAY()), MONTH(TODAY())+1, 1)), 0)'
    );
    
    // ========== REMAINING & STATUS ==========
    sheet.getRange('D7:D27').setFormula('=IF(B7>0, B7-C7, 0)');
    sheet.getRange('E7:E24').setFormula(
      '=IF(C7=0, "⚪ Chưa chi", IF(C7>B7, "🔴 Vượt ngân sách", IF(C7/B7>=0.8, "⚠️ Sắp hết", "✅ Trong hạn mức")))'
    );
    
    // Formatting
    sheet.getRange('B2:D30').setNumberFormat('#,##0');
    sheet.setColumnWidth(1, 200);
    sheet.setColumnWidth(2, 120);
    sheet.setColumnWidth(3, 120);
    sheet.setColumnWidth(4, 120);
    sheet.setColumnWidth(5, 150);
    sheet.setColumnWidth(6, 200);
    
    return sheet;
  }
};