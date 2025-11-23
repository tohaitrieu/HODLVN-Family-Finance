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
    sheet.setColumnWidth(3, 120);
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
    sheet.getRange('C2:C').setNumberFormat('#,##0');
    sheet.getRange('D2:D').setNumberFormat('0.00"%"');
    sheet.getRange('E2:E').setNumberFormat('0'); // Term (Month) - Number
    sheet.getRange('F2:G').setNumberFormat(APP_CONFIG.FORMATS.DATE); // Start Date, Maturity Date
    this._fixDateColumn(sheet, 6); // Ngày vay
    this._fixDateColumn(sheet, 7); // Ngày đến hạn
    sheet.getRange('H2:J').setNumberFormat('#,##0');
    
    // Formula
    // K: Còn nợ = Gốc (D) - Đã trả gốc (I)
    sheet.getRange('K2:K1000').setFormula('=IFERROR(D2-I2, 0)');
    
    // Validation
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
    sheet.setColumnWidth(3, 120);
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
    sheet.getRange('C2:C').setNumberFormat('#,##0');
    sheet.getRange('D2:D').setNumberFormat('0.00"%"');
    sheet.getRange('E2:E').setNumberFormat('0'); // Term - Number
    sheet.getRange('F2:G').setNumberFormat(APP_CONFIG.FORMATS.DATE); // Start Date, Maturity Date
    this._fixDateColumn(sheet, 6); // Ngày vay
    this._fixDateColumn(sheet, 7); // Ngày đến hạn
    sheet.getRange('H2:J').setNumberFormat('#,##0');
    
    // Formula
    // K: Còn lại = Gốc (D) - Gốc đã thu (I)
    sheet.getRange('K2:K1000').setFormula('=IFERROR(D2-I2, 0)');
    
    // Validation
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
    
    // Only set 0 if empty to avoid overwriting user input
    if (sheet.getRange('B2').getValue() === '') {
      sheet.getRange('B2:F2').merge().setValue(0);
    }
    sheet.getRange('B2').setNumberFormat('#,##0')
      .setFontWeight('bold')
      .setHorizontalAlignment('right')
      .setBackground('#E7E6E6');
    
    // ========== ROW 3: % NHÓM CHI TIÊU ==========
    sheet.getRange('A3').setValue('Nhóm Chi tiêu:')
      .setFontWeight('bold')
      .setFontColor('#E74C3C');
    
    if (sheet.getRange('B3').getValue() === '') {
      sheet.getRange('B3').setValue(0.5);
    }
    sheet.getRange('B3').setNumberFormat('0.00%')
      .setFontWeight('bold')
      .setBackground('#FFF3CD')
      .setHorizontalAlignment('center');
    
    // ========== SECTION 1: CHI TIÊU ==========
    sheet.getRange('A4:F4').merge()
      .setValue('📤 CHI TIÊU')
      .setFontWeight('bold')
      .setBackground('#E74C3C')
      .setFontColor('#FFFFFF')
      .setHorizontalAlignment('center');
    
    // Header columns
    const chiHeaders = ['Danh mục', '% Nhóm', 'Ngân sách', 'Đã chi', 'Còn lại', 'Trạng thái'];
    sheet.getRange('A5:F5').setValues([chiHeaders])
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBackground('#4472C4')
      .setFontColor('#FFFFFF');
    
    // Danh mục chi tiêu với % mặc định
    const expenseCategories = [
      ['Ăn uống', 0.35],
      ['Đi lại', 0.10],
      ['Nhà ở', 0.30],
      ['Y tế', 0.05],
      ['Giáo dục', 0.10],
      ['Mua sắm', 0.07],
      ['Giải trí', 0.02],
      ['Khác', 0.01]
    ];
    
    // Điền dữ liệu chi tiêu
    for (let i = 0; i < expenseCategories.length; i++) {
      const row = 6 + i;
      const category = expenseCategories[i][0];
      const pct = expenseCategories[i][1];
      
      // A: Danh mục
      sheet.getRange(row, 1).setValue(category);
      
      // B: % Nhóm (Only set if empty)
      if (sheet.getRange(row, 2).getValue() === '') {
        sheet.getRange(row, 2).setValue(pct);
      }
      sheet.getRange(row, 2).setNumberFormat('0.00%').setHorizontalAlignment('center');
      
      // C: Ngân sách
      sheet.getRange(row, 3).setFormula(`=$B$2*$B$3*B${row}`).setNumberFormat('#,##0');
      
      // D: Đã chi
      const formulaChi = `=SUMIFS(CHI!C:C, CHI!D:D, A${row}, CHI!B:B, ">="&DATE(${currentYear},${currentMonth},1), CHI!B:B, "<"&DATE(${currentYear},${currentMonth}+1,1))`;
      sheet.getRange(row, 4).setFormula(formulaChi).setNumberFormat('#,##0');
      
      // E: Còn lại
      sheet.getRange(row, 5).setFormula(`=C${row}-D${row}`).setNumberFormat('#,##0');
      
      // F: Trạng thái
      const statusFormula = `=IF(C${row}=0, "⚪ N/A", IF(E${row}<0, "🔴 " & TEXT(D${row}/C${row}, "0.0%"), IF(D${row}/C${row} > DAY(TODAY())/DAY(EOMONTH(TODAY(),0))+0.1, "🔴 " & TEXT(D${row}/C${row}, "0.0%"), IF(D${row}/C${row} > DAY(TODAY())/DAY(EOMONTH(TODAY(),0)), "⚠️ " & TEXT(D${row}/C${row}, "0.0%"), "✅ " & TEXT(D${row}/C${row}, "0.0%")))))`;
      sheet.getRange(row, 6).setFormula(statusFormula);
    }
    
    // TỔNG CHI
    const chiEndRow = 6 + expenseCategories.length;
    sheet.getRange(chiEndRow, 1).setValue('TỔNG CHI').setFontWeight('bold');
    sheet.getRange(chiEndRow, 2).setFormula(`=SUM(B6:B${chiEndRow-1})`).setNumberFormat('0.00%').setFontWeight('bold').setHorizontalAlignment('center');
    sheet.getRange(chiEndRow, 3).setFormula(`=SUM(C6:C${chiEndRow-1})`).setNumberFormat('#,##0').setFontWeight('bold');
    sheet.getRange(chiEndRow, 4).setFormula(`=SUM(D6:D${chiEndRow-1})`).setNumberFormat('#,##0').setFontWeight('bold');
    sheet.getRange(chiEndRow, 5).setFormula(`=SUM(E6:E${chiEndRow-1})`).setNumberFormat('#,##0').setFontWeight('bold');
    
    const tongChiStatus = `=IF(C${chiEndRow}=0, "⚪ N/A", IF(E${chiEndRow}<0, "🔴 " & TEXT(D${chiEndRow}/C${chiEndRow}, "0.0%"), IF(D${chiEndRow}/C${chiEndRow} > DAY(TODAY())/DAY(EOMONTH(TODAY(),0))+0.1, "🔴 " & TEXT(D${chiEndRow}/C${chiEndRow}, "0.0%"), IF(D${chiEndRow}/C${chiEndRow} > DAY(TODAY())/DAY(EOMONTH(TODAY(),0)), "⚠️ " & TEXT(D${chiEndRow}/C${chiEndRow}, "0.0%"), "✅ " & TEXT(D${chiEndRow}/C${chiEndRow}, "0.0%")))))`;
    sheet.getRange(chiEndRow, 6).setFormula(tongChiStatus).setFontWeight('bold');
    
    // ========== ROW: % NHÓM ĐẦU TƯ ==========
    const dautuRow = chiEndRow + 2;
    sheet.getRange(dautuRow, 1).setValue('Nhóm Đầu tư:').setFontWeight('bold').setFontColor('#70AD47');
    
    if (sheet.getRange(dautuRow, 2).getValue() === '') {
      sheet.getRange(dautuRow, 2).setValue(0.3);
    }
    sheet.getRange(dautuRow, 2).setNumberFormat('0.00%').setFontWeight('bold').setBackground('#D4EDDA').setHorizontalAlignment('center');
    
    // ========== SECTION 2: ĐẦU TƯ ==========
    const dautuHeaderRow = dautuRow + 1;
    sheet.getRange(`A${dautuHeaderRow}:F${dautuHeaderRow}`).merge()
      .setValue('💰 ĐẦU TƯ').setFontWeight('bold').setBackground('#70AD47').setFontColor('#FFFFFF').setHorizontalAlignment('center');
    
    const dautuColRow = dautuHeaderRow + 1;
    sheet.getRange(`A${dautuColRow}:F${dautuColRow}`).setValues([chiHeaders])
      .setFontWeight('bold').setHorizontalAlignment('center').setBackground('#4472C4').setFontColor('#FFFFFF');
    
    const investCategories = [
      ['Chứng khoán', 0.50, `=SUMIFS('CHỨNG KHOÁN'!H:H, 'CHỨNG KHOÁN'!C:C, "Mua", 'CHỨNG KHOÁN'!B:B, ">="&DATE(${currentYear},${currentMonth},1), 'CHỨNG KHOÁN'!B:B, "<"&DATE(${currentYear},${currentMonth}+1,1))`],
      ['Vàng', 0.20, `=SUMIFS(VÀNG!H:H, VÀNG!C:C, "Mua", VÀNG!B:B, ">="&DATE(${currentYear},${currentMonth},1), VÀNG!B:B, "<"&DATE(${currentYear},${currentMonth}+1,1))`],
      ['Crypto', 0.20, `=SUMIFS(CRYPTO!I:I, CRYPTO!C:C, "Mua", CRYPTO!B:B, ">="&DATE(${currentYear},${currentMonth},1), CRYPTO!B:B, "<"&DATE(${currentYear},${currentMonth}+1,1))`],
      ['Đầu tư khác', 0.10, `=SUMIFS('ĐẦU TƯ KHÁC'!D:D, 'ĐẦU TƯ KHÁC'!B:B, ">="&DATE(${currentYear},${currentMonth},1), 'ĐẦU TƯ KHÁC'!B:B, "<"&DATE(${currentYear},${currentMonth}+1,1))`]
    ];
    
    for (let i = 0; i < investCategories.length; i++) {
      const row = dautuColRow + 1 + i;
      const category = investCategories[i][0];
      const pct = investCategories[i][1];
      const formula = investCategories[i][2];
      
      sheet.getRange(row, 1).setValue(category);
      if (sheet.getRange(row, 2).getValue() === '') {
        sheet.getRange(row, 2).setValue(pct);
      }
      sheet.getRange(row, 2).setNumberFormat('0.00%').setHorizontalAlignment('center');
      sheet.getRange(row, 3).setFormula(`=$B$2*$B$${dautuRow}*B${row}`).setNumberFormat('#,##0');
      sheet.getRange(row, 4).setFormula(formula).setNumberFormat('#,##0');
      sheet.getRange(row, 5).setFormula(`=C${row}-D${row}`).setNumberFormat('#,##0');
      
      const statusFormula = `=IF(C${row}=0, "⚪ N/A", IF(E${row}<0, "🔴 " & TEXT(D${row}/C${row}, "0.0%"), IF(D${row}/C${row} > DAY(TODAY())/DAY(EOMONTH(TODAY(),0))+0.1, "🔴 " & TEXT(D${row}/C${row}, "0.0%"), IF(D${row}/C${row} > DAY(TODAY())/DAY(EOMONTH(TODAY(),0)), "⚠️ " & TEXT(D${row}/C${row}, "0.0%"), "✅ " & TEXT(D${row}/C${row}, "0.0%")))))`;
      sheet.getRange(row, 6).setFormula(statusFormula);
    }

    // ========== ROW: % NHÓM TRẢ NỢ ==========
    const debtRow = dautuColRow + investCategories.length + 2;
    sheet.getRange(debtRow, 1).setValue('Nhóm Trả nợ:').setFontWeight('bold').setFontColor('#FF9800');
    
    if (sheet.getRange(debtRow, 2).getValue() === '') {
      sheet.getRange(debtRow, 2).setValue(0.2);
    }
    sheet.getRange(debtRow, 2).setNumberFormat('0.00%').setFontWeight('bold').setBackground('#FFE0B2').setHorizontalAlignment('center');
    
    // ========== SECTION 3: TRẢ NỢ ==========
    const debtHeaderRow = debtRow + 1;
    sheet.getRange(`A${debtHeaderRow}:F${debtHeaderRow}`).merge()
      .setValue('💸 TRẢ NỢ').setFontWeight('bold').setBackground('#FF9800').setFontColor('#FFFFFF').setHorizontalAlignment('center');
    
    const debtColRow = debtHeaderRow + 1;
    sheet.getRange(`A${debtColRow}:F${debtColRow}`).setValues([chiHeaders])
      .setFontWeight('bold').setHorizontalAlignment('center').setBackground('#4472C4').setFontColor('#FFFFFF');
      
    // Data Row
    const debtDataRow = debtColRow + 1;
    sheet.getRange(debtDataRow, 1).setValue('Trả nợ (Gốc + Lãi)');
    sheet.getRange(debtDataRow, 2).setValue(1).setNumberFormat('0.00%').setHorizontalAlignment('center');
    
    // Budget
    sheet.getRange(debtDataRow, 3).setFormula(`=$B$2*$B$${debtRow}`).setNumberFormat('#,##0');
    
    // Actual
    const formulaTraNo = `=SUMIFS('TRẢ NỢ'!F:F, 'TRẢ NỢ'!B:B, ">="&DATE(${currentYear},${currentMonth},1), 'TRẢ NỢ'!B:B, "<"&DATE(${currentYear},${currentMonth}+1,1))`;
    sheet.getRange(debtDataRow, 4).setFormula(formulaTraNo).setNumberFormat('#,##0');
    
    // Remaining
    sheet.getRange(debtDataRow, 5).setFormula(`=C${debtDataRow}-D${debtDataRow}`).setNumberFormat('#,##0');
    
    // Status
    const debtStatusFormula = `=IF(C${debtDataRow}=0, "⚪ N/A", IF(E${debtDataRow}<0, "🔴 " & TEXT(D${debtDataRow}/C${debtDataRow}, "0.0%"), IF(D${debtDataRow}/C${debtDataRow} > DAY(TODAY())/DAY(EOMONTH(TODAY(),0))+0.1, "🔴 " & TEXT(D${debtDataRow}/C${debtDataRow}, "0.0%"), IF(D${debtDataRow}/C${debtDataRow} > DAY(TODAY())/DAY(EOMONTH(TODAY(),0)), "⚠️ " & TEXT(D${debtDataRow}/C${debtDataRow}, "0.0%"), "✅ " & TEXT(D${debtDataRow}/C${debtDataRow}, "0.0%")))))`;
    sheet.getRange(debtDataRow, 6).setFormula(debtStatusFormula);
    
    // ========== SYNC WARNING ==========
    // Calculate Expected Debt Payment for current month
    if (typeof DashboardManager !== 'undefined' && typeof DashboardManager._getCalendarEvents === 'function') {
      const events = DashboardManager._getCalendarEvents();
      const currentMonthEvents = events.payables.filter(e => {
        const d = new Date(e.date);
        return d.getMonth() + 1 === currentMonth && d.getFullYear() === currentYear;
      });
      
      const totalExpected = currentMonthEvents.reduce((sum, e) => sum + (e.principalPayment || 0) + (e.interestPayment || 0), 0);
      
      // Get Budgeted Amount (Need to fetch value because formula isn't calculated yet in script)
      // Budget = Income * Debt Ratio
      const income = sheet.getRange('B2').getValue() || 0;
      const debtRatio = sheet.getRange(debtRow, 2).getValue() || 0;
      const budgeted = income * debtRatio;
      
      if (budgeted < totalExpected) {
        SpreadsheetApp.getUi().alert(
          '⚠️ CẢNH BÁO NGÂN SÁCH TRẢ NỢ',
          `Ngân sách trả nợ dự kiến (${Utilities.formatString('#,##0', budgeted)}) thấp hơn nghĩa vụ trả nợ thực tế trong tháng (${Utilities.formatString('#,##0', totalExpected)}).\n\nVui lòng điều chỉnh tỷ lệ phân bổ ngân sách hoặc tăng thu nhập dự kiến!`,
          SpreadsheetApp.getUi().ButtonSet.OK
        );
      }
    }
    
    return sheet;
  },
  /**
   * Sửa lỗi lệch cột do thêm cột "Loại hình"
   */
  fixDebtLendingAlignment() {
    const ss = getSpreadsheet();
    const sheets = [APP_CONFIG.SHEETS.DEBT_MANAGEMENT, APP_CONFIG.SHEETS.LENDING];
    
    sheets.forEach(sheetName => {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) return;
      
      const lastRow = sheet.getLastRow();
      if (lastRow < 2) return;
      
      // Check C2 (Should be Type - String) and D2 (Should be Principal - Number)
      // If C2 is Number (Principal) and D2 is Number (Rate < 1), it's likely misaligned
      const c2 = sheet.getRange('C2').getValue();
      const d2 = sheet.getRange('D2').getValue();
      
      // Logic: Old C was Principal (Large Number), Old D was Rate (Small Number < 1)
      // New C is Type (String), New D is Principal (Large Number)
      // If C2 is Large Number (> 1000), it's likely Principal -> Misaligned
      if (typeof c2 === 'number' && c2 > 1000) {
        Logger.log(`Phát hiện lệch cột tại sheet ${sheetName}. Đang sửa...`);
        
        // Insert cells at C2:C (Shift Right)
        sheet.getRange(2, 3, lastRow - 1, 1).insertCells(SpreadsheetApp.Dimension.COLUMNS);
        
        // Set default value "Khác" for new C column
        sheet.getRange(2, 3, lastRow - 1, 1).setValue('Khác');
        
        // Re-apply Formula for K (Remaining)
        // K = D - I
        sheet.getRange(2, 11, lastRow - 1, 1).setFormula('=IFERROR(D2-I2, 0)');
        
        Logger.log(`Đã sửa xong sheet ${sheetName}`);
      } else {
        Logger.log(`Sheet ${sheetName} có vẻ đã đúng cấu trúc.`);
        // Ensure formula is correct anyway
        sheet.getRange(2, 11, lastRow - 1, 1).setFormula('=IFERROR(D2-I2, 0)');
      }
    });
    
    SpreadsheetApp.getUi().alert('✅ Đã kiểm tra và sửa lỗi lệch cột (nếu có)!');
  }
};