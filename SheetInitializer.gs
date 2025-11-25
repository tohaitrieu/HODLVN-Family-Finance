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
    SheetUtils.applySheetFormat(sheet, 'INCOME');
    this._fixDateColumn(sheet, 2);
    return sheet;
  },
  
  /**
   * Khởi tạo Sheet CHI
   */
  initializeExpenseSheet() {
    const ss = getSpreadsheet();
    const sheet = this._getOrCreateSheet(ss, APP_CONFIG.SHEETS.EXPENSE);
    SheetUtils.applySheetFormat(sheet, 'EXPENSE');
    this._fixDateColumn(sheet, 2);
    return sheet;
  },
  
  /**
   * Khởi tạo Sheet TRẢ NỢ
   */
  initializeDebtPaymentSheet() {
    const ss = getSpreadsheet();
    const sheet = this._getOrCreateSheet(ss, APP_CONFIG.SHEETS.DEBT_PAYMENT);
    SheetUtils.applySheetFormat(sheet, 'DEBT_PAYMENT');
    this._fixDateColumn(sheet, 2);
    return sheet;
  },
  
  /**
   * Khởi tạo Sheet QUẢN LÝ NỢ
   */
  initializeDebtManagementSheet() {
    const ss = getSpreadsheet();
    const sheet = this._getOrCreateSheet(ss, APP_CONFIG.SHEETS.DEBT_MANAGEMENT);
    SheetUtils.applySheetFormat(sheet, 'DEBT_MANAGEMENT');
    this._fixDateColumn(sheet, 7); // Ngày vay
    this._fixDateColumn(sheet, 8); // Ngày đến hạn
    this._fixTermColumn(sheet, 6); // Kỳ hạn
    return sheet;
  },
  
  /**
   * Khởi tạo Sheet CHO VAY
   */
  initializeLendingSheet() {
    const ss = getSpreadsheet();
    const sheet = this._getOrCreateSheet(ss, APP_CONFIG.SHEETS.LENDING);
    SheetUtils.applySheetFormat(sheet, 'LENDING');
    this._fixDateColumn(sheet, 7); // Ngày vay
    this._fixDateColumn(sheet, 8); // Ngày đến hạn
    this._fixTermColumn(sheet, 6); // Kỳ hạn
    return sheet;
  },
  
  /**
   * Khởi tạo Sheet THU NỢ
   */
  initializeLendingRepaymentSheet() {
    const ss = getSpreadsheet();
    const sheet = this._getOrCreateSheet(ss, APP_CONFIG.SHEETS.LENDING_REPAYMENT);
    SheetUtils.applySheetFormat(sheet, 'LENDING_REPAYMENT');
    this._fixDateColumn(sheet, 2);
    return sheet;
  },
  
  /**
   * Khởi tạo Sheet CHỨNG KHOÁN
   */
  initializeStockSheet() {
    const ss = getSpreadsheet();
    const sheet = this._getOrCreateSheet(ss, APP_CONFIG.SHEETS.STOCK);
    SheetUtils.applySheetFormat(sheet, 'STOCK');
    this._fixDateColumn(sheet, 2);
    
    // Conditional Formatting (Specific to Stock)
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
    
    return sheet;
  },
  
  /**
   * Khởi tạo Sheet VÀNG
   */
  initializeGoldSheet() {
    const ss = getSpreadsheet();
    const sheet = this._getOrCreateSheet(ss, APP_CONFIG.SHEETS.GOLD);
    SheetUtils.applySheetFormat(sheet, 'GOLD');
    this._fixDateColumn(sheet, 2);
    
    // Conditional Formatting (Specific to Gold)
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
    
    return sheet;
  },
  
  /**
   * Khởi tạo Sheet CRYPTO
   */
  initializeCryptoSheet() {
    const ss = getSpreadsheet();
    const sheet = this._getOrCreateSheet(ss, APP_CONFIG.SHEETS.CRYPTO);
    SheetUtils.applySheetFormat(sheet, 'CRYPTO');
    this._fixDateColumn(sheet, 2);
    
    // Conditional Formatting (Specific to Crypto)
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
    
    return sheet;
  },
  
  /**
   * Khởi tạo Sheet ĐẦU TƯ KHÁC
   */
  initializeOtherInvestmentSheet() {
    const ss = getSpreadsheet();
    const sheet = this._getOrCreateSheet(ss, APP_CONFIG.SHEETS.OTHER_INVESTMENT);
    SheetUtils.applySheetFormat(sheet, 'OTHER_INVESTMENT');
    this._fixDateColumn(sheet, 2);
    return sheet;
  },

  /**
   * Khởi tạo Sheet CHANGELOG
   */
  initializeChangelogSheet() {
    const ss = getSpreadsheet();
    const sheet = this._getOrCreateSheet(ss, APP_CONFIG.SHEETS.CHANGELOG);
    SheetUtils.applySheetFormat(sheet, 'CHANGELOG');
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
      .setHorizontalAlignment('left')
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
      .setHorizontalAlignment('left')
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
      .setHorizontalAlignment('left')
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
    
    // Apply global formatting using SheetConfig
    SheetUtils.applySheetFormat(sheet, 'BUDGET');

    // Fix v251125: Re-apply B2 format after applySheetFormat() override
    // B2 (Thu nhập dự kiến / Expected Income) should be plain number (#,##0), not percentage (0.0%)
    // Root cause: applySheetFormat() applies column B percentage format to all rows from row 2
    sheet.getRange('B2').setNumberFormat('#,##0');

    // Set all rows to consistent height
    const maxRows = sheet.getMaxRows();
    for (let row = 1; row <= maxRows; row++) {
      const height = (row === 1) ? GLOBAL_SHEET_CONFIG.HEADER_ROW_HEIGHT : GLOBAL_SHEET_CONFIG.DEFAULT_ROW_HEIGHT;
      sheet.setRowHeight(row, height);
    }
    
    return sheet;
  }
};