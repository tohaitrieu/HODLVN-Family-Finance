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
   * Khởi tạo Sheet THU
   */
  initializeIncomeSheet() {
    const ss = getSpreadsheet();
    const sheet = this._getOrCreateSheet(ss, APP_CONFIG.SHEETS.INCOME);
    
    // Header
    const headers = ['STT', 'Ngày', 'Số tiền', 'Nguồn thu', 'Ghi chú'];
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
    const headers = ['STT', 'Ngày', 'Số tiền', 'Danh mục', 'Chi tiết', 'Ghi chú'];
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
    const headers = ['STT', 'Ngày', 'Khoản nợ', 'Trả gốc', 'Trả lãi', 'Tổng trả', 'Ghi chú'];
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
      'STT', 'Tên khoản nợ', 'Nợ gốc ban đầu', 'Lãi suất (%/năm)', 
      'Kỳ hạn (tháng)', 'Ngày vay', 'Ngày đến hạn', 'Đã trả gốc', 
      'Đã trả lãi', 'Còn nợ', 'Trạng thái', 'Ghi chú'
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
    
    // Format
    sheet.getRange('A2:A').setNumberFormat('0');
    sheet.getRange('C2:C').setNumberFormat('#,##0');
    sheet.getRange('D2:D').setNumberFormat('0.00"%"');
    sheet.getRange('E2:E').setNumberFormat('0');
    sheet.getRange('F2:G').setNumberFormat(APP_CONFIG.FORMATS.DATE);
    this._fixDateColumn(sheet, 6); // Ngày vay
    this._fixDateColumn(sheet, 7); // Ngày đến hạn
    sheet.getRange('H2:J').setNumberFormat('#,##0');
    
    // Formula
    sheet.getRange('J2:J1000').setFormula('=IFERROR(C2-H2, 0)');
    
    // Validation
    const statusRange = sheet.getRange('K2:K1000');
    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Chưa trả', 'Đang trả', 'Đã thanh toán', 'Quá hạn'])
      .setAllowInvalid(false)
      .build();
    statusRange.setDataValidation(statusRule);
    
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
    const headers = ['STT', 'Ngày', 'Loại GD', 'Loại vàng', 'Số lượng', 'Đơn vị', 'Giá', 'Tổng', 'Nơi lưu', 'Ghi chú'];
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
    sheet.setColumnWidth(4, 100);
    sheet.setColumnWidth(5, 80);
    sheet.setColumnWidth(6, 70);
    sheet.setColumnWidth(7, 100);
    sheet.setColumnWidth(8, 120);
    sheet.setColumnWidth(9, 120);
    sheet.setColumnWidth(10, 200);
    
    // Format
    sheet.getRange('A2:A').setNumberFormat('0');
    sheet.getRange('B2:B').setNumberFormat(APP_CONFIG.FORMATS.DATE);
    this._fixDateColumn(sheet, 2);
    sheet.getRange('G2:H').setNumberFormat('#,##0');
    
    sheet.setFrozenRows(1);
    
    // Validations
    sheet.getRange('C2:C1000').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Mua', 'Bán']).build());
    sheet.getRange('D2:D1000').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['SJC', '24K', '18K', '14K', '10K', 'Khác']).build());
    sheet.getRange('F2:F1000').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['chỉ', 'lượng', 'cây', 'gram']).build());
    
    return sheet;
  },
  
  /**
   * Khởi tạo Sheet CRYPTO
   */
  initializeCryptoSheet() {
    const ss = getSpreadsheet();
    const sheet = this._getOrCreateSheet(ss, APP_CONFIG.SHEETS.CRYPTO);
    
    // Header
    const headers = ['STT', 'Ngày', 'Loại GD', 'Coin', 'Số lượng', 'Giá (USD)', 'Tỷ giá', 'Giá', 'Tổng', 'Sàn', 'Ví', 'Ghi chú'];
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
    sheet.setColumnWidth(5, 100);
    sheet.setColumnWidth(6, 100);
    sheet.setColumnWidth(7, 80);
    sheet.setColumnWidth(8, 100);
    sheet.setColumnWidth(9, 120);
    sheet.setColumnWidth(10, 100);
    sheet.setColumnWidth(11, 150);
    sheet.setColumnWidth(12, 200);
    
    // Format
    sheet.getRange('A2:A').setNumberFormat('0');
    sheet.getRange('B2:B').setNumberFormat(APP_CONFIG.FORMATS.DATE);
    this._fixDateColumn(sheet, 2);
    sheet.getRange('F2:F').setNumberFormat('#,##0.00');
    sheet.getRange('H2:I').setNumberFormat('#,##0');
    
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
    
    return sheet;
  }
};