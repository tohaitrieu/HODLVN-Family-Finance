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
    
    sheet.getRange('B2:F2').merge()
      .setValue(0) // Placeholder - user sẽ nhập
      .setNumberFormat('#,##0" VNĐ"')
      .setFontWeight('bold')
      .setHorizontalAlignment('right')
      .setBackground('#E7E6E6');
    
    // ========== ROW 3: % NHÓM CHI TIÊU ==========
    sheet.getRange('A3').setValue('Nhóm Chi tiêu:')
      .setFontWeight('bold')
      .setFontColor('#E74C3C');
    
    sheet.getRange('B3').setValue(0.5) // 50% mặc định
      .setNumberFormat('0.00%')
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
      
      // B: % Nhóm
      sheet.getRange(row, 2).setValue(pct)
        .setNumberFormat('0.00%')
        .setHorizontalAlignment('center');
      
      // C: Ngân sách = Thu nhập × % Nhóm Chi × % Danh mục
      sheet.getRange(row, 3).setFormula(`=$B$2*$B$3*B${row}`)
        .setNumberFormat('#,##0" VNĐ"');
      
      // D: Đã chi = SUMIFS từ sheet CHI
      const formulaChi = `=SUMIFS(CHI!C:C, CHI!D:D, A${row}, CHI!B:B, ">="&DATE(${currentYear},${currentMonth},1), CHI!B:B, "<"&DATE(${currentYear},${currentMonth}+1,1))`;
      sheet.getRange(row, 4).setFormula(formulaChi)
        .setNumberFormat('#,##0" VNĐ"');
      
      // E: Còn lại = Ngân sách - Đã chi
      sheet.getRange(row, 5).setFormula(`=C${row}-D${row}`)
        .setNumberFormat('#,##0" VNĐ"');
      
      // F: Trạng thái thông minh theo ngày
      const statusFormula = `=IF(C${row}=0, "⚪ N/A", IF(E${row}<0, "🔴 " & TEXT(D${row}/C${row}, "0.0%"), IF(D${row}/C${row} > DAY(TODAY())/DAY(EOMONTH(TODAY(),0))+0.1, "🔴 " & TEXT(D${row}/C${row}, "0.0%"), IF(D${row}/C${row} > DAY(TODAY())/DAY(EOMONTH(TODAY(),0)), "⚠️ " & TEXT(D${row}/C${row}, "0.0%"), "✅ " & TEXT(D${row}/C${row}, "0.0%")))))`;
      sheet.getRange(row, 6).setFormula(statusFormula);
    }
    
    // TỔNG CHI
    const chiEndRow = 6 + expenseCategories.length;
    sheet.getRange(chiEndRow, 1).setValue('TỔNG CHI')
      .setFontWeight('bold');
    sheet.getRange(chiEndRow, 2).setFormula(`=SUM(B6:B${chiEndRow-1})`)
      .setNumberFormat('0.00%')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    sheet.getRange(chiEndRow, 3).setFormula(`=SUM(C6:C${chiEndRow-1})`)
      .setNumberFormat('#,##0" VNĐ"')
      .setFontWeight('bold');
    sheet.getRange(chiEndRow, 4).setFormula(`=SUM(D6:D${chiEndRow-1})`)
      .setNumberFormat('#,##0" VNĐ"')
      .setFontWeight('bold');
    sheet.getRange(chiEndRow, 5).setFormula(`=SUM(E6:E${chiEndRow-1})`)
      .setNumberFormat('#,##0" VNĐ"')
      .setFontWeight('bold');
    
    // Trạng thái TỔNG CHI
    const tongChiStatus = `=IF(C${chiEndRow}=0, "⚪ N/A", IF(E${chiEndRow}<0, "🔴 " & TEXT(D${chiEndRow}/C${chiEndRow}, "0.0%"), IF(D${chiEndRow}/C${chiEndRow} > DAY(TODAY())/DAY(EOMONTH(TODAY(),0))+0.1, "🔴 " & TEXT(D${chiEndRow}/C${chiEndRow}, "0.0%"), IF(D${chiEndRow}/C${chiEndRow} > DAY(TODAY())/DAY(EOMONTH(TODAY(),0)), "⚠️ " & TEXT(D${chiEndRow}/C${chiEndRow}, "0.0%"), "✅ " & TEXT(D${chiEndRow}/C${chiEndRow}, "0.0%")))))`;
    sheet.getRange(chiEndRow, 6).setFormula(tongChiStatus)
      .setFontWeight('bold');
    
    // ========== ROW: % NHÓM ĐẦU TƯ ==========
    const dautuRow = chiEndRow + 2;
    sheet.getRange(dautuRow, 1).setValue('Nhóm Đầu tư:')
      .setFontWeight('bold')
      .setFontColor('#70AD47');
    
    sheet.getRange(dautuRow, 2).setValue(0.3) // 30% mặc định
      .setNumberFormat('0.00%')
      .setFontWeight('bold')
      .setBackground('#D4EDDA')
      .setHorizontalAlignment('center');
    
    // ========== SECTION 2: ĐẦU TƯ ==========
    const dautuHeaderRow = dautuRow + 1;
    sheet.getRange(`A${dautuHeaderRow}:F${dautuHeaderRow}`).merge()
      .setValue('💰 ĐẦU TƯ')
      .setFontWeight('bold')
      .setBackground('#70AD47')
      .setFontColor('#FFFFFF')
      .setHorizontalAlignment('center');
    
    const dautuColRow = dautuHeaderRow + 1;
    sheet.getRange(`A${dautuColRow}:F${dautuColRow}`).setValues([chiHeaders])
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBackground('#4472C4')
      .setFontColor('#FFFFFF');
    
    // Danh mục đầu tư với % mặc định
    const investCategories = [
      ['Chứng khoán', 0.50, `=SUMIFS('CHỨNG KHOÁN'!H:H, 'CHỨNG KHOÁN'!C:C, "Mua", 'CHỨNG KHOÁN'!B:B, ">="&DATE(${currentYear},${currentMonth},1), 'CHỨNG KHOÁN'!B:B, "<"&DATE(${currentYear},${currentMonth}+1,1))`],
      ['Vàng', 0.20, `=SUMIFS(VÀNG!H:H, VÀNG!C:C, "Mua", VÀNG!B:B, ">="&DATE(${currentYear},${currentMonth},1), VÀNG!B:B, "<"&DATE(${currentYear},${currentMonth}+1,1))`],
      ['Crypto', 0.20, `=SUMIFS(CRYPTO!I:I, CRYPTO!C:C, "Mua", CRYPTO!B:B, ">="&DATE(${currentYear},${currentMonth},1), CRYPTO!B:B, "<"&DATE(${currentYear},${currentMonth}+1,1))`],
      ['Đầu tư khác', 0.10, `=SUMIFS('ĐẦU TƯ KHÁC'!D:D, 'ĐẦU TƯ KHÁC'!B:B, ">="&DATE(${currentYear},${currentMonth},1), 'ĐẦU TƯ KHÁC'!B:B, "<"&DATE(${currentYear},${currentMonth}+1,1))`]
    ];
    
    // Điền dữ liệu đầu tư
    for (let i = 0; i < investCategories.length; i++) {
      const row = dautuColRow + 1 + i;
      const category = investCategories[i][0];
      const pct = investCategories[i][1];
      const formula = investCategories[i][2];
      
      // A: Loại
      sheet.getRange(row, 1).setValue(category);
      
      // B: % Nhóm
      sheet.getRange(row, 2).setValue(pct)
        .setNumberFormat('0.00%')
        .setHorizontalAlignment('center');
      
      // C: Target = Thu nhập × % Nhóm Đầu tư × % Loại
      sheet.getRange(row, 3).setFormula(`=$B$2*$B$${dautuRow}*B${row}`)
        .setNumberFormat('#,##0" VNĐ"');
      
      // D: Đã đầu tư
      sheet.getRange(row, 4).setFormula(formula)
        .setNumberFormat('#,##0" VNĐ"');
      
      // E: Còn lại
      sheet.getRange(row, 5).setFormula(`=C${row}-D${row}`)
        .setNumberFormat('#,##0" VNĐ"');
      
      // F: Trạng thái
      const investStatusFormula = `=IF(C${row}=0, "⚪ N/A", IF(E${row}<0, "🔴 " & TEXT(D${row}/C${row}, "0.0%"), IF(D${row}/C${row} > DAY(TODAY())/DAY(EOMONTH(TODAY(),0))+0.1, "🔴 " & TEXT(D${row}/C${row}, "0.0%"), IF(D${row}/C${row} > DAY(TODAY())/DAY(EOMONTH(TODAY(),0)), "⚠️ " & TEXT(D${row}/C${row}, "0.0%"), "✅ " & TEXT(D${row}/C${row}, "0.0%")))))`;
      sheet.getRange(row, 6).setFormula(investStatusFormula);
    }
    
    // TỔNG ĐẦU TƯ
    const dautuEndRow = dautuColRow + 1 + investCategories.length;
    sheet.getRange(dautuEndRow, 1).setValue('TỔNG ĐT')
      .setFontWeight('bold');
    sheet.getRange(dautuEndRow, 2).setFormula(`=SUM(B${dautuColRow+1}:B${dautuEndRow-1})`)
      .setNumberFormat('0.00%')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    sheet.getRange(dautuEndRow, 3).setFormula(`=SUM(C${dautuColRow+1}:C${dautuEndRow-1})`)
      .setNumberFormat('#,##0" VNĐ"')
      .setFontWeight('bold');
    sheet.getRange(dautuEndRow, 4).setFormula(`=SUM(D${dautuColRow+1}:D${dautuEndRow-1})`)
      .setNumberFormat('#,##0" VNĐ"')
      .setFontWeight('bold');
    sheet.getRange(dautuEndRow, 5).setFormula(`=SUM(E${dautuColRow+1}:E${dautuEndRow-1})`)
      .setNumberFormat('#,##0" VNĐ"')
      .setFontWeight('bold');
    
    const tongDtStatus = tongChiStatus.replace(new RegExp(`${chiEndRow}`, 'g'), dautuEndRow);
    sheet.getRange(dautuEndRow, 6).setFormula(tongDtStatus)
      .setFontWeight('bold');
    
    // ========== ROW: % NHÓM TRẢ NỢ ==========
    const tranoRow = dautuEndRow + 2;
    sheet.getRange(tranoRow, 1).setValue('Nhóm Trả nợ:')
      .setFontWeight('bold')
      .setFontColor('#F39C12');
    
    sheet.getRange(tranoRow, 2).setValue(0.2) // 20% mặc định
      .setNumberFormat('0.00%')
      .setFontWeight('bold')
      .setBackground('#FFF3CD')
      .setHorizontalAlignment('center');
    
    // ========== SECTION 3: TRẢ NỢ ==========
    const tranoHeaderRow = tranoRow + 1;
    sheet.getRange(`A${tranoHeaderRow}:F${tranoHeaderRow}`).merge()
      .setValue('💳 TRẢ NỢ')
      .setFontWeight('bold')
      .setBackground('#F39C12')
      .setFontColor('#FFFFFF')
      .setHorizontalAlignment('center');
    
    const tranoColRow = tranoHeaderRow + 1;
    sheet.getRange(`A${tranoColRow}:F${tranoColRow}`).setValues([chiHeaders])
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBackground('#4472C4')
      .setFontColor('#FFFFFF');
    
    const tranoDataRow = tranoColRow + 1;
    
    // Trả nợ
    sheet.getRange(tranoDataRow, 1).setValue('Trả nợ');
    sheet.getRange(tranoDataRow, 2).setValue(1.0) // 100% của nhóm
      .setNumberFormat('0.00%')
      .setHorizontalAlignment('center');
    sheet.getRange(tranoDataRow, 3).setFormula(`=$B$2*$B$${tranoRow}`)
      .setNumberFormat('#,##0" VNĐ"');
    
    const tranoFormula = `=SUMIFS('TRẢ NỢ'!D:D, 'TRẢ NỢ'!B:B, ">="&DATE(${currentYear},${currentMonth},1), 'TRẢ NỢ'!B:B, "<"&DATE(${currentYear},${currentMonth}+1,1)) + SUMIFS('TRẢ NỢ'!E:E, 'TRẢ NỢ'!B:B, ">="&DATE(${currentYear},${currentMonth},1), 'TRẢ NỢ'!B:B, "<"&DATE(${currentYear},${currentMonth}+1,1))`;
    sheet.getRange(tranoDataRow, 4).setFormula(tranoFormula)
      .setNumberFormat('#,##0" VNĐ"');
    sheet.getRange(tranoDataRow, 5).setFormula(`=C${tranoDataRow}-D${tranoDataRow}`)
      .setNumberFormat('#,##0" VNĐ"');
    
    const tranoStatusFormula = `=IF(C${tranoDataRow}=0, "⚪ N/A", IF(E${tranoDataRow}<0, "🔴 " & TEXT(D${tranoDataRow}/C${tranoDataRow}, "0.0%"), IF(D${tranoDataRow}/C${tranoDataRow} > DAY(TODAY())/DAY(EOMONTH(TODAY(),0))+0.1, "🔴 " & TEXT(D${tranoDataRow}/C${tranoDataRow}, "0.0%"), IF(D${tranoDataRow}/C${tranoDataRow} > DAY(TODAY())/DAY(EOMONTH(TODAY(),0)), "⚠️ " & TEXT(D${tranoDataRow}/C${tranoDataRow}, "0.0%"), "✅ " & TEXT(D${tranoDataRow}/C${tranoDataRow}, "0.0%")))))`;
    sheet.getRange(tranoDataRow, 6).setFormula(tranoStatusFormula);
    
    // ========== FORMAT ==========
    
    // Column widths
    sheet.setColumnWidth(1, 150);  // Danh mục
    sheet.setColumnWidth(2, 80);   // % Nhóm
    sheet.setColumnWidth(3, 120);  // Ngân sách
    sheet.setColumnWidth(4, 120);  // Đã chi/đầu tư
    sheet.setColumnWidth(5, 120);  // Còn lại
    sheet.setColumnWidth(6, 120);  // Trạng thái
    
    // Borders
    sheet.getRange(`A1:F${tranoDataRow}`).setBorder(
      true, true, true, true, true, true,
      '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID
    );
    
    // Freeze header
    sheet.setFrozenRows(5);
    
    // Validation: Tổng % các nhóm phải = 100%
    const validationCell = sheet.getRange('C2');
    const totalPct = sheet.getRange('D2');
    totalPct.setFormula(`=B3+B${dautuRow}+B${tranoRow}`)
      .setNumberFormat('0.00%')
      .setNote('Tổng % các nhóm (Chi + Đầu tư + Trả nợ). Phải = 100%');
    
    showSuccess('Thành công', '✅ Sheet BUDGET v3.5 đã được khởi tạo!\n\n' +
      '📊 Cấu trúc mới:\n' +
      '• Thu nhập dự kiến (cell B2)\n' +
      '• 3 nhóm: Chi tiêu (B3), Đầu tư (B' + dautuRow + '), Trả nợ (B' + tranoRow + ')\n' +
      '• % chi tiết trong mỗi nhóm\n\n' +
      '⚠️ LƯU Ý:\n' +
      '• Nhập Thu nhập dự kiến vào cell B2\n' +
      '• Điều chỉnh % các nhóm để tổng = 100%\n' +
      '• Điều chỉnh % chi tiết trong từng nhóm\n\n' +
      '💡 Trạng thái tự động cảnh báo theo % ngày trong tháng!');
    
    return sheet;
  }

// ==================== KẾT THÚC HÀM ====================
};