/**
 * ===============================================
 * SHEETINITIALIZER.GS - MODULE KHỞI TẠO CÁC SHEET
 * ===============================================
 * 
 * Chức năng:
 * - Khởi tạo cấu trúc và format cho từng sheet
 * - Tạo header, validation, công thức
 * - Áp dụng format và màu sắc nhất quán
 * 
 * VERSION: 3.3 - FIXED: Xóa tất cả dữ liệu mẫu + Thêm trạng thái "Chưa trả"
 */

const SheetInitializer = {
  
  /**
   * Khởi tạo Sheet THU
   * ✅ FIXED: Xóa dữ liệu mẫu
   */
  initializeIncomeSheet() {
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName(APP_CONFIG.SHEETS.INCOME);
    
    if (sheet) {
      ss.deleteSheet(sheet);
    }
    
    sheet = ss.insertSheet(APP_CONFIG.SHEETS.INCOME);
    
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
    
    // Format
    sheet.getRange('A2:A').setNumberFormat('0');
    sheet.getRange('B2:B').setNumberFormat(APP_CONFIG.FORMATS.DATE);
    sheet.getRange('C2:C').setNumberFormat(APP_CONFIG.FORMATS.NUMBER);
    
    // Freeze header
    sheet.setFrozenRows(1);
    
    // Data validation cho Nguồn thu
    const sourceRange = sheet.getRange('D2:D1000');
    const sourceRule = SpreadsheetApp.newDataValidation()
      .requireValueInList([
        'Lương',
        'MMO (Make Money Online)',
        'Thưởng',
        'Bán CK',
        'Bán Vàng',
        'Bán Crypto',
        'Lãi đầu tư',
        'Thu hồi nợ',
        'Vay ngân hàng',
        'Vay cá nhân',
        'Khác'
      ])
      .setAllowInvalid(false)
      .build();
    sourceRange.setDataValidation(sourceRule);
    
    // ✅ FIX: KHÔNG THÊM DỮ LIỆU MẪU
    // Để sheet hoàn toàn trống, chỉ có header
    
    return sheet;
  },
  
  /**
   * Khởi tạo Sheet CHI
   * ✅ FIXED: Xóa dữ liệu mẫu
   */
  initializeExpenseSheet() {
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName(APP_CONFIG.SHEETS.EXPENSE);
    
    if (sheet) {
      ss.deleteSheet(sheet);
    }
    
    sheet = ss.insertSheet(APP_CONFIG.SHEETS.EXPENSE);
    
    // Header
    const headers = ['STT', 'Ngày', 'Số tiền', 'Danh mục', 'Chi tiết', 'Ghi chú'];
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
    sheet.setColumnWidth(4, 120);  // Danh mục
    sheet.setColumnWidth(5, 200);  // Chi tiết
    sheet.setColumnWidth(6, 250);  // Ghi chú
    
    // Format
    sheet.getRange('A2:A').setNumberFormat('0');
    sheet.getRange('B2:B').setNumberFormat(APP_CONFIG.FORMATS.DATE);
    sheet.getRange('C2:C').setNumberFormat(APP_CONFIG.FORMATS.NUMBER);
    
    // Freeze header
    sheet.setFrozenRows(1);
    
    // Data validation cho Danh mục
    const categoryRange = sheet.getRange('D2:D1000');
    const categoryRule = SpreadsheetApp.newDataValidation()
      .requireValueInList([
        'Ăn uống',
        'Đi lại',
        'Nhà ở',
        'Y tế',
        'Giáo dục',
        'Mua sắm',
        'Giải trí',
        'Khác'
      ])
      .setAllowInvalid(false)
      .build();
    categoryRange.setDataValidation(categoryRule);
    
    // ✅ FIX: KHÔNG THÊM DỮ LIỆU MẪU
    // Để sheet hoàn toàn trống, chỉ có header
    
    return sheet;
  },
  
  /**
   * Khởi tạo Sheet TRẢ NỢ
   */
  initializeDebtPaymentSheet() {
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName(APP_CONFIG.SHEETS.DEBT_PAYMENT);
    
    if (sheet) {
      ss.deleteSheet(sheet);
    }
    
    sheet = ss.insertSheet(APP_CONFIG.SHEETS.DEBT_PAYMENT);
    
    // Header
    const headers = ['STT', 'Ngày', 'Khoản nợ', 'Trả gốc', 'Trả lãi', 'Tổng trả', 'Ghi chú'];
    sheet.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBackground(APP_CONFIG.COLORS.HEADER_BG)
      .setFontColor(APP_CONFIG.COLORS.HEADER_TEXT);
    
    // Column widths
    sheet.setColumnWidth(1, 50);   // STT
    sheet.setColumnWidth(2, 100);  // Ngày
    sheet.setColumnWidth(3, 150);  // Khoản nợ
    sheet.setColumnWidth(4, 120);  // Trả gốc
    sheet.setColumnWidth(5, 120);  // Trả lãi
    sheet.setColumnWidth(6, 120);  // Tổng trả
    sheet.setColumnWidth(7, 250);  // Ghi chú
    
    // Format
    sheet.getRange('A2:A').setNumberFormat('0');
    sheet.getRange('B2:B').setNumberFormat(APP_CONFIG.FORMATS.DATE);
    sheet.getRange('D2:F').setNumberFormat(APP_CONFIG.FORMATS.NUMBER);
    
    // Công thức Tổng trả - FIXED: Thêm IFERROR
    sheet.getRange('F2:F1000').setFormula('=IFERROR(D2+E2, 0)');
    
    // Freeze header
    sheet.setFrozenRows(1);
    
    return sheet;
  },
  
  /**
   * Khởi tạo Sheet QUẢN LÝ NỢ
   * ✅ FIXED: Xóa dữ liệu mẫu + Thêm trạng thái "Chưa trả"
   */
  initializeDebtManagementSheet() {
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName(APP_CONFIG.SHEETS.DEBT_MANAGEMENT);
    
    if (sheet) {
      ss.deleteSheet(sheet);
    }
    
    sheet = ss.insertSheet(APP_CONFIG.SHEETS.DEBT_MANAGEMENT);
    
    // Header
    const headers = [
      'STT',
      'Tên khoản nợ',
      'Nợ gốc ban đầu',
      'Lãi suất (%/năm)',
      'Kỳ hạn (tháng)',
      'Ngày vay',
      'Ngày đến hạn',
      'Đã trả gốc',
      'Đã trả lãi',
      'Còn nợ',
      'Trạng thái',
      'Ghi chú'
    ];
    
    sheet.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBackground(APP_CONFIG.COLORS.HEADER_BG)
      .setFontColor(APP_CONFIG.COLORS.HEADER_TEXT);
    
    // Column widths
    sheet.setColumnWidth(1, 50);   // STT
    sheet.setColumnWidth(2, 150);  // Tên khoản nợ
    sheet.setColumnWidth(3, 120);  // Nợ gốc
    sheet.setColumnWidth(4, 100);  // Lãi suất
    sheet.setColumnWidth(5, 100);  // Kỳ hạn
    sheet.setColumnWidth(6, 100);  // Ngày vay
    sheet.setColumnWidth(7, 100);  // Ngày đến hạn
    sheet.setColumnWidth(8, 120);  // Đã trả gốc
    sheet.setColumnWidth(9, 120);  // Đã trả lãi
    sheet.setColumnWidth(10, 120); // Còn nợ
    sheet.setColumnWidth(11, 100); // Trạng thái
    sheet.setColumnWidth(12, 200); // Ghi chú
    
    // Format
    sheet.getRange('A2:A').setNumberFormat('0');
    sheet.getRange('C2:C').setNumberFormat(APP_CONFIG.FORMATS.NUMBER);
    sheet.getRange('D2:D').setNumberFormat('0.00"%"');
    sheet.getRange('E2:E').setNumberFormat('0');
    sheet.getRange('F2:G').setNumberFormat(APP_CONFIG.FORMATS.DATE);
    sheet.getRange('H2:J').setNumberFormat(APP_CONFIG.FORMATS.NUMBER);
    
    // Công thức Còn nợ - FIXED: Thêm IFERROR
    sheet.getRange('J2:J1000').setFormula('=IFERROR(C2-H2, 0)');
    
    // ✅ FIX: Data validation cho Trạng thái - THÊM "Chưa trả"
    const statusRange = sheet.getRange('K2:K1000');
    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Chưa trả', 'Đang trả', 'Đã thanh toán', 'Quá hạn'])
      .setAllowInvalid(false)
      .build();
    statusRange.setDataValidation(statusRule);
    
    // Freeze header
    sheet.setFrozenRows(1);
    
    // ✅ FIX: KHÔNG THÊM DỮ LIỆU MẪU
    // Để sheet hoàn toàn trống, chỉ có header và công thức
    // Công thức sẽ tự động tính toán khi có dữ liệu
    
    return sheet;
  },
  
  /**
   * Khởi tạo Sheet CHỨNG KHOÁN
   */
  initializeStockSheet() {
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName(APP_CONFIG.SHEETS.STOCK);
    
    if (sheet) {
      ss.deleteSheet(sheet);
    }
    
    sheet = ss.insertSheet(APP_CONFIG.SHEETS.STOCK);
    
    // Header
    const headers = ['STT', 'Ngày', 'Loại GD', 'Mã CK', 'Số lượng', 'Giá', 'Phí', 'Tổng', 'Ghi chú'];
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
    sheet.setColumnWidth(4, 80);   // Mã CK
    sheet.setColumnWidth(5, 80);   // Số lượng
    sheet.setColumnWidth(6, 100);  // Giá
    sheet.setColumnWidth(7, 100);  // Phí
    sheet.setColumnWidth(8, 120);  // Tổng
    sheet.setColumnWidth(9, 250);  // Ghi chú
    
    // Format
    sheet.getRange('A2:A').setNumberFormat('0');
    sheet.getRange('B2:B').setNumberFormat(APP_CONFIG.FORMATS.DATE);
    sheet.getRange('F2:H').setNumberFormat(APP_CONFIG.FORMATS.NUMBER);
    
    // Freeze header
    sheet.setFrozenRows(1);
    
    // Data validation cho Loại GD
    const typeRange = sheet.getRange('C2:C1000');
    const typeRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Mua', 'Bán'])
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
    let sheet = ss.getSheetByName(APP_CONFIG.SHEETS.GOLD);
    
    if (sheet) {
      ss.deleteSheet(sheet);
    }
    
    sheet = ss.insertSheet(APP_CONFIG.SHEETS.GOLD);
    
    // Header
    const headers = ['STT', 'Ngày', 'Loại GD', 'Loại vàng', 'Số lượng', 'Đơn vị', 'Giá', 'Tổng', 'Nơi lưu', 'Ghi chú'];
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
    sheet.setColumnWidth(4, 100);  // Loại vàng
    sheet.setColumnWidth(5, 80);   // Số lượng
    sheet.setColumnWidth(6, 70);   // Đơn vị
    sheet.setColumnWidth(7, 100);  // Giá
    sheet.setColumnWidth(8, 120);  // Tổng
    sheet.setColumnWidth(9, 120);  // Nơi lưu
    sheet.setColumnWidth(10, 200); // Ghi chú
    
    // Format
    sheet.getRange('A2:A').setNumberFormat('0');
    sheet.getRange('B2:B').setNumberFormat(APP_CONFIG.FORMATS.DATE);
    sheet.getRange('G2:H').setNumberFormat(APP_CONFIG.FORMATS.NUMBER);
    
    // Freeze header
    sheet.setFrozenRows(1);
    
    // Data validation cho Loại GD
    const typeRange = sheet.getRange('C2:C1000');
    const typeRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Mua', 'Bán'])
      .setAllowInvalid(false)
      .build();
    typeRange.setDataValidation(typeRule);
    
    // Data validation cho Loại vàng
    const goldTypeRange = sheet.getRange('D2:D1000');
    const goldTypeRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['SJC', '24K', '18K', '14K', '10K', 'Khác'])
      .setAllowInvalid(false)
      .build();
    goldTypeRange.setDataValidation(goldTypeRule);
    
    // Data validation cho Đơn vị
    const unitRange = sheet.getRange('F2:F1000');
    const unitRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['chỉ', 'lượng', 'cây', 'gram'])
      .setAllowInvalid(false)
      .build();
    unitRange.setDataValidation(unitRule);
    
    return sheet;
  },
  
  /**
   * Khởi tạo Sheet CRYPTO
   */
  initializeCryptoSheet() {
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName(APP_CONFIG.SHEETS.CRYPTO);
    
    if (sheet) {
      ss.deleteSheet(sheet);
    }
    
    sheet = ss.insertSheet(APP_CONFIG.SHEETS.CRYPTO);
    
    // Header
    const headers = ['STT', 'Ngày', 'Loại GD', 'Coin', 'Số lượng', 'Giá (USD)', 'Tỷ giá', 'Giá (VNĐ)', 'Tổng (VNĐ)', 'Sàn', 'Ví', 'Ghi chú'];
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
    sheet.setColumnWidth(6, 100);  // Giá USD
    sheet.setColumnWidth(7, 80);   // Tỷ giá
    sheet.setColumnWidth(8, 100);  // Giá VNĐ
    sheet.setColumnWidth(9, 120);  // Tổng VNĐ
    sheet.setColumnWidth(10, 100); // Sàn
    sheet.setColumnWidth(11, 150); // Ví
    sheet.setColumnWidth(12, 200); // Ghi chú
    
    // Format
    sheet.getRange('A2:A').setNumberFormat('0');
    sheet.getRange('B2:B').setNumberFormat(APP_CONFIG.FORMATS.DATE);
    sheet.getRange('F2:F').setNumberFormat('#,##0.00" USD"');
    sheet.getRange('H2:I').setNumberFormat(APP_CONFIG.FORMATS.NUMBER);
    
    // Freeze header
    sheet.setFrozenRows(1);
    
    // Data validation cho Loại GD
    const typeRange = sheet.getRange('C2:C1000');
    const typeRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Mua', 'Bán', 'Swap', 'Stake', 'Unstake'])
      .setAllowInvalid(false)
      .build();
    typeRange.setDataValidation(typeRule);
    
    return sheet;
  },
  
  /**
   * Khởi tạo Sheet ĐẦU TƯ KHÁC
   */
  initializeOtherInvestmentSheet() {
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName(APP_CONFIG.SHEETS.OTHER_INVESTMENT);
    
    if (sheet) {
      ss.deleteSheet(sheet);
    }
    
    sheet = ss.insertSheet(APP_CONFIG.SHEETS.OTHER_INVESTMENT);
    
    // Header
    const headers = ['STT', 'Ngày', 'Loại đầu tư', 'Số tiền', 'Lãi suất (%)', 'Kỳ hạn (tháng)', 'Dự kiến thu về', 'Ghi chú'];
    sheet.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBackground(APP_CONFIG.COLORS.HEADER_BG)
      .setFontColor(APP_CONFIG.COLORS.HEADER_TEXT);
    
    // Column widths
    sheet.setColumnWidth(1, 50);   // STT
    sheet.setColumnWidth(2, 100);  // Ngày
    sheet.setColumnWidth(3, 150);  // Loại
    sheet.setColumnWidth(4, 120);  // Số tiền
    sheet.setColumnWidth(5, 100);  // Lãi suất
    sheet.setColumnWidth(6, 100);  // Kỳ hạn
    sheet.setColumnWidth(7, 120);  // Dự kiến
    sheet.setColumnWidth(8, 250);  // Ghi chú
    
    // Format
    sheet.getRange('A2:A').setNumberFormat('0');
    sheet.getRange('B2:B').setNumberFormat(APP_CONFIG.FORMATS.DATE);
    sheet.getRange('D2:D').setNumberFormat(APP_CONFIG.FORMATS.NUMBER);
    sheet.getRange('E2:E').setNumberFormat('0.00"%"');
    sheet.getRange('G2:G').setNumberFormat(APP_CONFIG.FORMATS.NUMBER);
    
    // Freeze header
    sheet.setFrozenRows(1);
    
    // Data validation cho Loại đầu tư
    const typeRange = sheet.getRange('C2:C1000');
    const typeRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Gửi tiết kiệm', 'Quỹ đầu tư', 'Bất động sản', 'Trái phiếu', 'P2P Lending', 'Khác'])
      .setAllowInvalid(false)
      .build();
    typeRange.setDataValidation(typeRule);
    
    return sheet;
  },
  
  /**
   * Khởi tạo Sheet BUDGET
   */
  initializeBudgetSheet() {
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName(APP_CONFIG.SHEETS.BUDGET);
    
    if (sheet) {
      ss.deleteSheet(sheet);
    }
    
    sheet = ss.insertSheet(APP_CONFIG.SHEETS.BUDGET);
    
    // Title
    sheet.getRange('A1:D1').merge()
      .setValue('💰 NGÂN SÁCH THÁNG ' + (new Date().getMonth() + 1) + '/' + new Date().getFullYear())
      .setFontSize(14)
      .setFontWeight('bold')
      .setBackground(APP_CONFIG.COLORS.PRIMARY)
      .setFontColor('#FFFFFF')
      .setHorizontalAlignment('center');
    
    // ========== SECTION 1: CHI TIÊU ==========
    sheet.getRange('A3:D3').merge()
      .setValue('📤 CHI TIÊU')
      .setFontWeight('bold')
      .setBackground('#E74C3C')
      .setFontColor('#FFFFFF');
    
    const expenseHeaders = ['Danh mục', 'Ngân sách', 'Đã chi', 'Tỷ lệ %'];
    sheet.getRange('A4:D4').setValues([expenseHeaders])
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBackground(APP_CONFIG.COLORS.HEADER_BG)
      .setFontColor(APP_CONFIG.COLORS.HEADER_TEXT);
    
    const expenseCategories = [
      ['Ăn uống', 3000000],
      ['Đi lại', 1000000],
      ['Nhà ở', 5000000],
      ['Y tế', 500000],
      ['Giáo dục', 1000000],
      ['Mua sắm', 2000000],
      ['Giải trí', 1000000],
      ['Khác', 500000]
    ];
    
    for (let i = 0; i < expenseCategories.length; i++) {
      const row = 5 + i;
      sheet.getRange(row, 1).setValue(expenseCategories[i][0]);
      sheet.getRange(row, 2).setValue(expenseCategories[i][1]);
      sheet.getRange(row, 3).setValue(0);
      // FIXED: Thêm IFERROR để tránh #DIV/0!
      sheet.getRange(row, 4).setFormula(`=IFERROR(C${row}/B${row}, 0)`);
    }
    
    // Total Chi tiêu
    const expenseEndRow = 5 + expenseCategories.length;
    sheet.getRange(expenseEndRow, 1).setValue('Tổng').setFontWeight('bold');
    sheet.getRange(expenseEndRow, 2).setFormula(`=SUM(B5:B${expenseEndRow-1})`).setFontWeight('bold');
    sheet.getRange(expenseEndRow, 3).setFormula(`=SUM(C5:C${expenseEndRow-1})`).setFontWeight('bold');
    // FIXED: Thêm IFERROR
    sheet.getRange(expenseEndRow, 4).setFormula(`=IFERROR(C${expenseEndRow}/B${expenseEndRow}, 0)`).setFontWeight('bold');
    
    // ========== SECTION 2: Nợ & LÃI ==========
    const debtStartRow = expenseEndRow + 2;
    sheet.getRange(debtStartRow, 1, 1, 4).merge()
      .setValue('💳 Nợ & LÃI')
      .setFontWeight('bold')
      .setBackground('#FF6B6B')
      .setFontColor('#FFFFFF');
    
    sheet.getRange(debtStartRow + 1, 1, 1, 4).setValues([expenseHeaders])
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBackground(APP_CONFIG.COLORS.HEADER_BG)
      .setFontColor(APP_CONFIG.COLORS.HEADER_TEXT);
    
    const debtCategories = [
      ['Trả nợ gốc', 5000000],
      ['Trả lãi', 1000000]
    ];
    
    for (let i = 0; i < debtCategories.length; i++) {
      const row = debtStartRow + 2 + i;
      sheet.getRange(row, 1).setValue(debtCategories[i][0]);
      sheet.getRange(row, 2).setValue(debtCategories[i][1]);
      sheet.getRange(row, 3).setValue(0);
      // FIXED: Thêm IFERROR
      sheet.getRange(row, 4).setFormula(`=IFERROR(C${row}/B${row}, 0)`);
    }
    
    // Total Nợ & Lãi
    const debtEndRow = debtStartRow + 2 + debtCategories.length;
    sheet.getRange(debtEndRow, 1).setValue('Tổng').setFontWeight('bold');
    sheet.getRange(debtEndRow, 2).setFormula(`=SUM(B${debtStartRow+2}:B${debtEndRow-1})`).setFontWeight('bold');
    sheet.getRange(debtEndRow, 3).setFormula(`=SUM(C${debtStartRow+2}:C${debtEndRow-1})`).setFontWeight('bold');
    // FIXED: Thêm IFERROR
    sheet.getRange(debtEndRow, 4).setFormula(`=IFERROR(C${debtEndRow}/B${debtEndRow}, 0)`).setFontWeight('bold');
    
    // ========== SECTION 3: QUỸ DỰ PHÒNG ==========
    const reserveStartRow = debtEndRow + 2;
    sheet.getRange(reserveStartRow, 1, 1, 4).merge()
      .setValue('🛡️ QUỸ DỰ PHÒNG')
      .setFontWeight('bold')
      .setBackground('#4ECDC4')
      .setFontColor('#FFFFFF');
    
    sheet.getRange(reserveStartRow + 1, 1, 1, 4).setValues([expenseHeaders])
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBackground(APP_CONFIG.COLORS.HEADER_BG)
      .setFontColor(APP_CONFIG.COLORS.HEADER_TEXT);
    
    const reserveCategories = [
      ['Quỹ khẩn cấp', 2000000],
      ['Quỹ dự phòng', 1000000]
    ];
    
    for (let i = 0; i < reserveCategories.length; i++) {
      const row = reserveStartRow + 2 + i;
      sheet.getRange(row, 1).setValue(reserveCategories[i][0]);
      sheet.getRange(row, 2).setValue(reserveCategories[i][1]);
      sheet.getRange(row, 3).setValue(0);
      // FIXED: Thêm IFERROR
      sheet.getRange(row, 4).setFormula(`=IFERROR(C${row}/B${row}, 0)`);
    }
    
    // Total Quỹ dự phòng
    const reserveEndRow = reserveStartRow + 2 + reserveCategories.length;
    sheet.getRange(reserveEndRow, 1).setValue('Tổng').setFontWeight('bold');
    sheet.getRange(reserveEndRow, 2).setFormula(`=SUM(B${reserveStartRow+2}:B${reserveEndRow-1})`).setFontWeight('bold');
    sheet.getRange(reserveEndRow, 3).setFormula(`=SUM(C${reserveStartRow+2}:C${reserveEndRow-1})`).setFontWeight('bold');
    // FIXED: Thêm IFERROR
    sheet.getRange(reserveEndRow, 4).setFormula(`=IFERROR(C${reserveEndRow}/B${reserveEndRow}, 0)`).setFontWeight('bold');
    
    // ========== SECTION 4: ĐẦU TƯ ==========
    const investStartRow = reserveEndRow + 2;
    sheet.getRange(investStartRow, 1, 1, 4).merge()
      .setValue('💼 ĐẦU TƯ')
      .setFontWeight('bold')
      .setBackground('#70AD47')
      .setFontColor('#FFFFFF');
    
    sheet.getRange(investStartRow + 1, 1, 1, 4).setValues([expenseHeaders])
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBackground(APP_CONFIG.COLORS.HEADER_BG)
      .setFontColor(APP_CONFIG.COLORS.HEADER_TEXT);
    
    const investCategories = [
      ['Chứng khoán', 3000000],
      ['Vàng', 2000000],
      ['Crypto', 1000000],
      ['Đầu tư khác', 1000000]
    ];
    
    for (let i = 0; i < investCategories.length; i++) {
      const row = investStartRow + 2 + i;
      sheet.getRange(row, 1).setValue(investCategories[i][0]);
      sheet.getRange(row, 2).setValue(investCategories[i][1]);
      sheet.getRange(row, 3).setValue(0);
      // FIXED: Thêm IFERROR
      sheet.getRange(row, 4).setFormula(`=IFERROR(C${row}/B${row}, 0)`);
    }
    
    // Total Đầu tư
    const investEndRow = investStartRow + 2 + investCategories.length;
    sheet.getRange(investEndRow, 1).setValue('Tổng').setFontWeight('bold');
    sheet.getRange(investEndRow, 2).setFormula(`=SUM(B${investStartRow+2}:B${investEndRow-1})`).setFontWeight('bold');
    sheet.getRange(investEndRow, 3).setFormula(`=SUM(C${investStartRow+2}:C${investEndRow-1})`).setFontWeight('bold');
    // FIXED: Thêm IFERROR
    sheet.getRange(investEndRow, 4).setFormula(`=IFERROR(C${investEndRow}/B${investEndRow}, 0)`).setFontWeight('bold');
    
    // ========== TỔNG KẾT ==========
    const summaryRow = investEndRow + 2;
    sheet.getRange(summaryRow, 1, 1, 4).merge()
      .setValue('📊 TỔNG KẾT NGÂN SÁCH')
      .setFontWeight('bold')
      .setFontSize(12)
      .setBackground(APP_CONFIG.COLORS.PRIMARY)
      .setFontColor('#FFFFFF')
      .setHorizontalAlignment('center');
    
    sheet.getRange(summaryRow + 1, 1).setValue('Tổng ngân sách tháng').setFontWeight('bold');
    sheet.getRange(summaryRow + 1, 2).setFormula(`=B${expenseEndRow}+B${debtEndRow}+B${reserveEndRow}+B${investEndRow}`)
      .setFontWeight('bold')
      .setBackground('#E8F5E9');
    
    sheet.getRange(summaryRow + 2, 1).setValue('Đã sử dụng').setFontWeight('bold');
    sheet.getRange(summaryRow + 2, 3).setFormula(`=C${expenseEndRow}+C${debtEndRow}+C${reserveEndRow}+C${investEndRow}`)
      .setFontWeight('bold')
      .setBackground('#FFF3E0');
    
    sheet.getRange(summaryRow + 3, 1).setValue('Tỷ lệ sử dụng').setFontWeight('bold');
    // FIXED: Thêm IFERROR
    sheet.getRange(summaryRow + 3, 4).setFormula(`=IFERROR(C${summaryRow+2}/B${summaryRow+1}, 0)`)
      .setFontWeight('bold')
      .setBackground('#E3F2FD')
      .setNumberFormat(APP_CONFIG.FORMATS.PERCENTAGE);
    
    // Format
    sheet.getRange(`B5:C${investEndRow}`).setNumberFormat(APP_CONFIG.FORMATS.NUMBER);
    sheet.getRange(`D5:D${investEndRow}`).setNumberFormat(APP_CONFIG.FORMATS.PERCENTAGE);
    sheet.getRange(`B${summaryRow+1}:C${summaryRow+2}`).setNumberFormat(APP_CONFIG.FORMATS.NUMBER);
    
    // Column widths
    sheet.setColumnWidth(1, 150);
    sheet.setColumnWidth(2, 120);
    sheet.setColumnWidth(3, 120);
    sheet.setColumnWidth(4, 100);
    
    // Conditional formatting for %
    const percentRange = sheet.getRange(`D5:D${investEndRow}`);
    
    const greenRule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThanOrEqualTo(0.8)
      .setBackground('#D4EDDA')
      .setFontColor('#155724')
      .setRanges([percentRange])
      .build();
    
    const yellowRule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(0.8, 1)
      .setBackground('#FFF3CD')
      .setFontColor('#856404')
      .setRanges([percentRange])
      .build();
    
    const redRule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(1)
      .setBackground('#F8D7DA')
      .setFontColor('#721C24')
      .setRanges([percentRange])
      .build();
    
    sheet.setConditionalFormatRules([greenRule, yellowRule, redRule]);
    
    return sheet;
  }
};