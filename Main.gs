/**
 * ===============================================
 * MAIN.GS - FILE CHÍNH HỆ THỐNG QUẢN LÝ TÀI CHÍNH v3.2
 * ===============================================
 * 
 * Kiến trúc Module:
 * - Main.gs: Điều phối, Menu, UI, Setup Wizard
 * - SheetInitializer.gs: Khởi tạo các sheet
 * - FormHandlers.gs: Xử lý form nhập liệu
 * - DataProcessors.gs: Xử lý giao dịch
 * - BudgetManager.gs: Quản lý ngân sách
 * - DashboardManager.gs: Quản lý dashboard & thống kê
 * 
 * VERSION: 3.2 - Added Setup Wizard
 */

// ==================== CẤU HÌNH TOÀN CỤC ====================

const APP_CONFIG = {
  VERSION: '3.2',
  APP_NAME: '💰 Quản lý Tài chính',
  
  // Danh sách các sheet
  SHEETS: {
    INCOME: 'THU',
    EXPENSE: 'CHI',
    DEBT_PAYMENT: 'TRẢ NỢ',
    STOCK: 'CHỨNG KHOÁN',
    GOLD: 'VÀNG',
    CRYPTO: 'CRYPTO',
    OTHER_INVESTMENT: 'ĐẦU TƯ KHÁC',
    DEBT_MANAGEMENT: 'QUẢN LÝ NỢ',
    BUDGET: 'BUDGET',
    DASHBOARD: 'TỔNG QUAN'
  },
  
  // Màu sắc theme
  COLORS: {
    PRIMARY: '#4472C4',
    SUCCESS: '#70AD47',
    DANGER: '#E74C3C',
    WARNING: '#F39C12',
    INFO: '#3498DB',
    HEADER_BG: '#4472C4',
    HEADER_TEXT: '#FFFFFF'
  },
  
  // Format số
  FORMATS: {
    NUMBER: '#,##0" VNĐ"',
    PERCENTAGE: '0.00%',
    DATE: 'dd/mm/yyyy'
  }
};

// ==================== MENU CHÍNH ====================

/**
 * Tạo menu khi mở file
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu(APP_CONFIG.APP_NAME)
    // === NHÓM 1: NHẬP LIỆU ===
    .addSubMenu(ui.createMenu('📝 Nhập liệu')
      .addItem('📥 Nhập Thu nhập', 'showIncomeForm')
      .addItem('📤 Nhập Chi tiêu', 'showExpenseForm')
      .addItem('💳 Thêm Khoản Nợ', 'showDebtManagementForm')
      .addItem('💳 Trả nợ', 'showDebtPaymentForm')
      .addSeparator()
      .addItem('📈 Giao dịch Chứng khoán', 'showStockForm')
      .addItem('🪙 Giao dịch Vàng', 'showGoldForm')
      .addItem('₿ Giao dịch Crypto', 'showCryptoForm')
      .addItem('💼 Đầu tư khác', 'showOtherInvestmentForm'))
    
    .addSeparator()

    // === NHÓM 2: BUDGET ===
    .addSubMenu(ui.createMenu('💵 Ngân sách')
      .addItem('⚠️ Kiểm tra Budget', 'checkBudgetWarnings')
      .addItem('📊 Báo cáo Chi tiêu', 'showExpenseReport')
      .addItem('💰 Báo cáo Đầu tư', 'showInvestmentReport'))
    
    .addSeparator()
    
    // === NHÓM 3: THỐNG KÊ ===
    .addSubMenu(ui.createMenu('📊 Thống kê & Dashboard')
      .addItem('🔄 Cập nhật Dashboard', 'refreshDashboard')
      .addSeparator()
      .addItem('📅 Xem Tất cả', 'viewAll')
      .addItem('📊 Xem Năm hiện tại', 'viewCurrentYear')
      .addItem('🗓️ Xem Quý hiện tại', 'viewCurrentQuarter')
      .addItem('📆 Xem Tháng hiện tại', 'viewCurrentMonth'))
    
    .addSeparator()
    
    // === NHÓM 4: TIỆN ÍCH ===
    .addSubMenu(ui.createMenu('🛠️ Tiện ích')
      .addItem('🔍 Tìm kiếm giao dịch', 'searchTransaction')
      .addItem('📤 Xuất báo cáo PDF', 'exportToPDF')
      .addItem('🗑️ Xóa dữ liệu test', 'clearTestData'))
    
    .addSeparator()

    // === NHÓM 5: KHỞI TẠO SHEET ===
    .addSubMenu(ui.createMenu('⚙️ Khởi tạo Sheet')
      .addItem('🔄 Khởi tạo TẤT CẢ Sheet', 'initializeAllSheets')
      .addSeparator()
      .addItem('📥 Khởi tạo Sheet THU', 'initializeIncomeSheet')
      .addItem('📤 Khởi tạo Sheet CHI', 'initializeExpenseSheet')
      .addItem('💳 Khởi tạo Sheet TRẢ NỢ', 'initializeDebtPaymentSheet')
      .addItem('📊 Khởi tạo Sheet QUẢN LÝ NỢ', 'initializeDebtManagementSheet')
      .addSeparator()
      .addItem('📈 Khởi tạo Sheet CHỨNG KHOÁN', 'initializeStockSheet')
      .addItem('🪙 Khởi tạo Sheet VÀNG', 'initializeGoldSheet')
      .addItem('₿ Khởi tạo Sheet CRYPTO', 'initializeCryptoSheet')
      .addItem('💼 Khởi tạo Sheet ĐẦU TƯ KHÁC', 'initializeOtherInvestmentSheet')
      .addSeparator()
      .addItem('💰 Khởi tạo Sheet BUDGET', 'initializeBudgetSheet')
      .addItem('📊 Khởi tạo Sheet TỔNG QUAN', 'initializeDashboardSheet'))
    
    .addSeparator()
    
    // === NHÓM 6: TRỢ GIÚP ===
    .addItem('ℹ️ Hướng dẫn sử dụng', 'showInstructions')
    .addItem('📖 Giới thiệu hệ thống', 'showAbout')
    
    .addToUi();
}

// ==================== HIỂN THỊ FORM ====================

/**
 * Hiển thị form nhập thu
 */
function showIncomeForm() {
  showForm('IncomeForm', '📥 Nhập Thu nhập', 450, 400);
}

/**
 * Hiển thị form nhập chi
 */
function showExpenseForm() {
  showForm('ExpenseForm', '📤 Nhập Chi tiêu', 450, 450);
}

/**
 * Hiển thị form trả nợ
 */
function showDebtPaymentForm() {
  showForm('DebtPaymentForm', '💳 Trả nợ', 450, 450);
}
/**
 * Hiển thị form thêm khoản nợ mới
 */
function showDebtManagementForm() {
  showForm('DebtManagementForm', '💳 Thêm Khoản Nợ Mới', 500, 650);
}
/**
 * Hiển thị form giao dịch chứng khoán
 */
function showStockForm() {
  showForm('StockForm', '📈 Giao dịch Chứng khoán', 450, 500);
}

/**
 * Hiển thị form giao dịch vàng
 */
function showGoldForm() {
  showForm('GoldForm', '🪙 Giao dịch Vàng', 450, 550);
}

/**
 * Hiển thị form giao dịch crypto
 */
function showCryptoForm() {
  showForm('CryptoForm', '₿ Giao dịch Crypto', 450, 550);
}

/**
 * Hiển thị form đầu tư khác
 */
function showOtherInvestmentForm() {
  showForm('OtherInvestmentForm', '💼 Đầu tư khác', 450, 550);
}

/**
 * Hàm helper hiển thị form
 */
function showForm(formName, title, width, height) {
  try {
    const html = HtmlService.createHtmlOutputFromFile(formName)
      .setWidth(width)
      .setHeight(height);
    SpreadsheetApp.getUi().showModalDialog(html, title);
  } catch (error) {
    showError(`Không thể mở form ${formName}`, error.message);
  }
}

// ==================== SETUP WIZARD - NEW IN v3.2 ====================

/**
 * Khởi tạo tất cả các sheet - PHIÊN BẢN MỚI VỚI SETUP WIZARD
 * Thay thế quy trình khởi tạo cũ bằng wizard 3 bước
 */
function initializeAllSheets() {
  // Hiển thị Setup Wizard thay vì confirm dialog
  showSetupWizard();
}

/**
 * Hiển thị Setup Wizard
 */
function showSetupWizard() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('SetupWizard')
      .setWidth(700)
      .setHeight(650);
    SpreadsheetApp.getUi().showModalDialog(html, '🚀 Thiết lập Hệ thống Quản lý Tài chính v3.2');
  } catch (error) {
    showError('Không thể mở Setup Wizard', 
      'Vui lòng đảm bảo file SetupWizard.html đã được tạo trong Apps Script.\n\n' +
      'Lỗi: ' + error.message);
  }
}

/**
 * Xử lý dữ liệu từ Setup Wizard
 * @param {Object} setupData - Dữ liệu từ form wizard
 * @return {Object} Kết quả thực thi
 */
function processSetupWizard(setupData) {
  try {
    const ss = getSpreadsheet();
    
    // ============================================================
    // BƯỚC 1: Khởi tạo tất cả sheets (KHÔNG có dữ liệu mẫu)
    // ============================================================
    
    // Sheet THU - Khởi tạo thủ công để không có dữ liệu mẫu
    let sheet = ss.getSheetByName(APP_CONFIG.SHEETS.INCOME);
    if (sheet) ss.deleteSheet(sheet);
    sheet = ss.insertSheet(APP_CONFIG.SHEETS.INCOME);
    
    const incomeHeaders = ['STT', 'Ngày', 'Số tiền', 'Nguồn thu', 'Ghi chú'];
    sheet.getRange(1, 1, 1, incomeHeaders.length)
      .setValues([incomeHeaders])
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBackground(APP_CONFIG.COLORS.HEADER_BG)
      .setFontColor(APP_CONFIG.COLORS.HEADER_TEXT);
    
    sheet.setColumnWidth(1, 50);
    sheet.setColumnWidth(2, 100);
    sheet.setColumnWidth(3, 120);
    sheet.setColumnWidth(4, 150);
    sheet.setColumnWidth(5, 300);
    
    sheet.getRange('A2:A').setNumberFormat('0');
    sheet.getRange('B2:B').setNumberFormat(APP_CONFIG.FORMATS.DATE);
    sheet.getRange('C2:C').setNumberFormat(APP_CONFIG.FORMATS.NUMBER);
    sheet.setFrozenRows(1);
    
    // Data validation cho Nguồn thu
    const sourceRange = sheet.getRange('D2:D1000');
    sourceRange.setNumberFormat('@'); // Set as plain text TRƯỚC validation
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
        'Khác'
      ])
      .setAllowInvalid(false)
      .build();
    sourceRange.setDataValidation(sourceRule);
    
    // Sheet CHI - Khởi tạo thủ công để không có dữ liệu mẫu
    sheet = ss.getSheetByName(APP_CONFIG.SHEETS.EXPENSE);
    if (sheet) ss.deleteSheet(sheet);
    sheet = ss.insertSheet(APP_CONFIG.SHEETS.EXPENSE);
    
    const expenseHeaders = ['STT', 'Ngày', 'Số tiền', 'Danh mục', 'Chi tiết', 'Ghi chú'];
    sheet.getRange(1, 1, 1, expenseHeaders.length)
      .setValues([expenseHeaders])
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBackground(APP_CONFIG.COLORS.HEADER_BG)
      .setFontColor(APP_CONFIG.COLORS.HEADER_TEXT);
    
    sheet.setColumnWidth(1, 50);
    sheet.setColumnWidth(2, 100);
    sheet.setColumnWidth(3, 120);
    sheet.setColumnWidth(4, 120);
    sheet.setColumnWidth(5, 200);
    sheet.setColumnWidth(6, 250);
    
    sheet.getRange('A2:A').setNumberFormat('0');
    sheet.getRange('B2:B').setNumberFormat(APP_CONFIG.FORMATS.DATE);
    sheet.getRange('C2:C').setNumberFormat(APP_CONFIG.FORMATS.NUMBER);
    sheet.setFrozenRows(1);
    
    // Data validation cho Danh mục
    const categoryRange = sheet.getRange('D2:D1000');
    categoryRange.setNumberFormat('@');
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
    
    // Khởi tạo các sheet còn lại bằng SheetInitializer
    SheetInitializer.initializeDebtPaymentSheet();
    SheetInitializer.initializeDebtManagementSheet();
    SheetInitializer.initializeStockSheet();
    SheetInitializer.initializeGoldSheet();
    SheetInitializer.initializeCryptoSheet();
    SheetInitializer.initializeOtherInvestmentSheet();
    
    // ============================================================
    // BƯỚC 2: Thêm số dư ban đầu vào THU
    // ============================================================
    const incomeSheet = ss.getSheetByName(APP_CONFIG.SHEETS.INCOME);
    const incomeData = [
      1, // STT
      new Date(setupData.balance.date),
      setupData.balance.amount,
      setupData.balance.source,
      'Số dư ban đầu (Setup Wizard)'
    ];
    incomeSheet.appendRow(incomeData);
    
    // ============================================================
    // BƯỚC 3: Thêm khoản nợ (nếu có)
    // ============================================================
    if (setupData.debt) {
      const debtSheet = ss.getSheetByName(APP_CONFIG.SHEETS.DEBT_MANAGEMENT);
      const startDate = new Date(setupData.debt.date);
      const dueDate = new Date(startDate);
      dueDate.setMonth(dueDate.getMonth() + setupData.debt.term);
      
      const debtData = [
        1, // STT
        setupData.debt.name,
        setupData.debt.principal,
        setupData.debt.rate,
        setupData.debt.term,
        startDate,
        dueDate,
        0, // Đã trả gốc
        0, // Đã trả lãi
        setupData.debt.principal, // Còn nợ
        'Đang trả',
        'Khởi tạo từ Setup Wizard'
      ];
      debtSheet.appendRow(debtData);
    }
    
    // ============================================================
    // BƯỚC 4: Khởi tạo Budget với giá trị từ form
    // ============================================================
    sheet = ss.getSheetByName(APP_CONFIG.SHEETS.BUDGET);
    if (sheet) ss.deleteSheet(sheet);
    sheet = ss.insertSheet(APP_CONFIG.SHEETS.BUDGET);
    
    // Header
    sheet.getRange('A1').setValue('NGÂN SÁCH THÁNG')
      .setFontSize(16)
      .setFontWeight('bold')
      .setBackground('#4472C4')
      .setFontColor('white');
    sheet.getRange('A1:D1').merge();
    
    // Section 1: CHI TIÊU
    sheet.getRange('A2').setValue('📊 CHI TIÊU')
      .setFontWeight('bold')
      .setBackground('#E7E6E6');
    sheet.getRange('A2:D2').merge();
    
    const categoryHeaders = ['Danh mục', 'Ngân sách', 'Đã chi', 'Tỷ lệ %'];
    sheet.getRange(3, 1, 1, 4)
      .setValues([categoryHeaders])
      .setFontWeight('bold')
      .setBackground('#D9D9D9');
    
    // Categories với budget values từ form
    const categories = [
      ['Ăn uống', setupData.budget.food],
      ['Đi lại', setupData.budget.transport],
      ['Nhà ở', setupData.budget.housing],
      ['Y tế', setupData.budget.health],
      ['Giáo dục', setupData.budget.education],
      ['Mua sắm', setupData.budget.shopping],
      ['Giải trí', setupData.budget.entertainment],
      ['Khác', setupData.budget.other]
    ];
    
    categories.forEach((cat, idx) => {
      const row = 4 + idx;
      sheet.getRange(row, 1).setValue(cat[0]);
      sheet.getRange(row, 2).setValue(cat[1]);
      sheet.getRange(row, 3).setValue(0);
      sheet.getRange(row, 4).setFormula(`=IFERROR(C${row}/B${row}, 0)`);
    });
    
    // Format
    sheet.getRange('B4:C11').setNumberFormat(APP_CONFIG.FORMATS.NUMBER);
    sheet.getRange('D4:D11').setNumberFormat('0.00%');
    
    // Conditional formatting cho tỷ lệ
    const percentRange = sheet.getRange('D4:D11');
    
    // Xanh: < 70%
    const rule1 = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0.7)
      .setBackground('#C6EFCE')
      .setRanges([percentRange])
      .build();
    
    // Vàng: 70-90%
    const rule2 = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(0.7, 0.9)
      .setBackground('#FFEB9C')
      .setRanges([percentRange])
      .build();
    
    // Đỏ: > 90%
    const rule3 = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0.9)
      .setBackground('#FFC7CE')
      .setRanges([percentRange])
      .build();
    
    sheet.setConditionalFormatRules([rule1, rule2, rule3]);
    
    // Tổng cộng
    sheet.getRange('A12').setValue('TỔNG').setFontWeight('bold');
    sheet.getRange('B12').setFormula('=SUM(B4:B11)');
    sheet.getRange('C12').setFormula('=SUM(C4:C11)');
    sheet.getRange('D12').setFormula('=IFERROR(C12/B12, 0)');
    sheet.getRange('B12:C12').setNumberFormat(APP_CONFIG.FORMATS.NUMBER);
    sheet.getRange('D12').setNumberFormat('0.00%');
    
    // Section 2: ĐẦU TƯ & TRẢ NỢ
    sheet.getRange('A14').setValue('💰 ĐẦU TƯ & TRẢ NỢ')
      .setFontWeight('bold')
      .setBackground('#E7E6E6');
    sheet.getRange('A14:D14').merge();
    
    sheet.getRange(15, 1, 1, 4)
      .setValues([['Loại', 'Target', 'Đã đầu tư/trả', 'Tỷ lệ %']])
      .setFontWeight('bold')
      .setBackground('#D9D9D9');
    
    const investmentTypes = [
      ['Chứng khoán', 0],
      ['Vàng', 0],
      ['Crypto', 0],
      ['Đầu tư khác', 0],
      ['Trả nợ', 0]
    ];
    
    investmentTypes.forEach((type, idx) => {
      const row = 16 + idx;
      sheet.getRange(row, 1).setValue(type[0]);
      sheet.getRange(row, 2).setValue(type[1]);
      sheet.getRange(row, 3).setValue(0);
      sheet.getRange(row, 4).setFormula(`=IFERROR(C${row}/B${row}, 0)`);
    });
    
    sheet.getRange('B16:C20').setNumberFormat(APP_CONFIG.FORMATS.NUMBER);
    sheet.getRange('D16:D20').setNumberFormat('0.00%');
    
    // Column widths
    sheet.setColumnWidth(1, 150);
    sheet.setColumnWidth(2, 130);
    sheet.setColumnWidth(3, 130);
    sheet.setColumnWidth(4, 100);
    
    // ============================================================
    // BƯỚC 5: Khởi tạo Dashboard
    // ============================================================
    DashboardManager.setupDashboard();
    
    // ============================================================
    // BƯỚC 6: Thành công
    // ============================================================
    const totalBudget = setupData.budget.food + setupData.budget.transport + 
                        setupData.budget.housing + setupData.budget.health +
                        setupData.budget.education + setupData.budget.shopping +
                        setupData.budget.entertainment + setupData.budget.other;
    
    return {
      success: true,
      message: 
        '✅ Khởi tạo thành công!\n\n' +
        '━━━━━━━━━━━━━━━━━━━━━━\n' +
        '💰 Số dư ban đầu: ' + formatCurrency(setupData.balance.amount) + '\n' +
        (setupData.debt ? 
          '📋 Khoản nợ: ' + setupData.debt.name + ' - ' + formatCurrency(setupData.debt.principal) + '\n' : 
          '📋 Khoản nợ: Không có\n') +
        '📊 Tổng ngân sách/tháng: ' + formatCurrency(totalBudget) + '\n' +
        '━━━━━━━━━━━━━━━━━━━━━━\n\n' +
        '🎉 Hệ thống đã sẵn sàng!\n' +
        'Bạn có thể bắt đầu nhập liệu qua Menu > Nhập liệu'
    };
    
  } catch (error) {
    Logger.log('Error in processSetupWizard: ' + error.toString());
    return {
      success: false,
      message: 'Lỗi khởi tạo: ' + error.message + '\n\nVui lòng thử lại hoặc liên hệ hỗ trợ.'
    };
  }
}

// ==================== KHỞI TẠO TỪNG SHEET ====================

/**
 * Khởi tạo Sheet THU
 */
function initializeIncomeSheet(skipConfirm) {
  if (!skipConfirm && !confirmInitialize('Sheet THU')) return;
  SheetInitializer.initializeIncomeSheet();
  if (!skipConfirm) showSuccess('Thành công', '✅ Đã khởi tạo Sheet THU!');
}

/**
 * Khởi tạo Sheet CHI
 */
function initializeExpenseSheet(skipConfirm) {
  if (!skipConfirm && !confirmInitialize('Sheet CHI')) return;
  SheetInitializer.initializeExpenseSheet();
  if (!skipConfirm) showSuccess('Thành công', '✅ Đã khởi tạo Sheet CHI!');
}

/**
 * Khởi tạo Sheet TRẢ NỢ
 */
function initializeDebtPaymentSheet(skipConfirm) {
  if (!skipConfirm && !confirmInitialize('Sheet TRẢ NỢ')) return;
  SheetInitializer.initializeDebtPaymentSheet();
  if (!skipConfirm) showSuccess('Thành công', '✅ Đã khởi tạo Sheet TRẢ NỢ!');
}

/**
 * Khởi tạo Sheet QUẢN LÝ NỢ
 */
function initializeDebtManagementSheet(skipConfirm) {
  if (!skipConfirm && !confirmInitialize('Sheet QUẢN LÝ NỢ')) return;
  SheetInitializer.initializeDebtManagementSheet();
  if (!skipConfirm) showSuccess('Thành công', '✅ Đã khởi tạo Sheet QUẢN LÝ NỢ!');
}

/**
 * Khởi tạo Sheet CHỨNG KHOÁN
 */
function initializeStockSheet(skipConfirm) {
  if (!skipConfirm && !confirmInitialize('Sheet CHỨNG KHOÁN')) return;
  SheetInitializer.initializeStockSheet();
  if (!skipConfirm) showSuccess('Thành công', '✅ Đã khởi tạo Sheet CHỨNG KHOÁN!');
}

/**
 * Khởi tạo Sheet VÀNG
 */
function initializeGoldSheet(skipConfirm) {
  if (!skipConfirm && !confirmInitialize('Sheet VÀNG')) return;
  SheetInitializer.initializeGoldSheet();
  if (!skipConfirm) showSuccess('Thành công', '✅ Đã khởi tạo Sheet VÀNG!');
}

/**
 * Khởi tạo Sheet CRYPTO
 */
function initializeCryptoSheet(skipConfirm) {
  if (!skipConfirm && !confirmInitialize('Sheet CRYPTO')) return;
  SheetInitializer.initializeCryptoSheet();
  if (!skipConfirm) showSuccess('Thành công', '✅ Đã khởi tạo Sheet CRYPTO!');
}

/**
 * Khởi tạo Sheet ĐẦU TƯ KHÁC
 */
function initializeOtherInvestmentSheet(skipConfirm) {
  if (!skipConfirm && !confirmInitialize('Sheet ĐẦU TƯ KHÁC')) return;
  SheetInitializer.initializeOtherInvestmentSheet();
  if (!skipConfirm) showSuccess('Thành công', '✅ Đã khởi tạo Sheet ĐẦU TƯ KHÁC!');
}

/**
 * Khởi tạo Sheet BUDGET
 */
function initializeBudgetSheet(skipConfirm) {
  if (!skipConfirm && !confirmInitialize('Sheet BUDGET')) return;
  SheetInitializer.initializeBudgetSheet();
  if (!skipConfirm) showSuccess('Thành công', '✅ Đã khởi tạo Sheet BUDGET!');
}

/**
 * Khởi tạo Sheet TỔNG QUAN (Dashboard)
 */
function initializeDashboardSheet(skipConfirm) {
  if (!skipConfirm && !confirmInitialize('Sheet TỔNG QUAN')) return;
  DashboardManager.setupDashboard();
  if (!skipConfirm) showSuccess('Thành công', '✅ Đã khởi tạo Sheet TỔNG QUAN!');
}

// ==================== HƯỚNG DẪN & GIỚI THIỆU ====================

/**
 * Hiển thị hướng dẫn sử dụng
 */
function showInstructions() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'Hướng dẫn sử dụng',
    '📖 HƯỚNG DẪN NHANH:\n\n' +
    '1️⃣ KHỞI TẠO:\n' +
    '   Menu > Khởi tạo Sheet > Khởi tạo TẤT CẢ Sheet\n' +
    '   → Setup Wizard sẽ hướng dẫn bạn từng bước!\n\n' +
    '2️⃣ NHẬP LIỆU:\n' +
    '   Menu > Nhập liệu > Chọn loại giao dịch\n\n' +
    '3️⃣ XEM THỐNG KÊ:\n' +
    '   Vào Sheet TỔNG QUAN\n' +
    '   Chọn Chu kỳ: Tất cả / Năm / Quý / Tháng\n\n' +
    '4️⃣ KIỂM TRA BUDGET:\n' +
    '   Menu > Ngân sách > Kiểm tra Budget\n\n' +
    '📚 Chi tiết xem file README.md',
    ui.ButtonSet.OK
  );
}

/**
 * Hiển thị thông tin hệ thống
 */
function showAbout() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'Giới thiệu hệ thống',
    `💰 ${APP_CONFIG.APP_NAME} v${APP_CONFIG.VERSION}\n\n` +
    '✨ MỚI TRONG v3.2:\n' +
    '   • Setup Wizard 3 bước chuyên nghiệp\n' +
    '   • Khởi tạo với dữ liệu thật của bạn\n' +
    '   • Không còn dữ liệu mẫu gây nhầm lẫn\n\n' +
    '🎯 Tính năng:\n' +
    '   • Quản lý thu chi hàng ngày\n' +
    '   • Theo dõi nợ và lãi\n' +
    '   • Đầu tư đa dạng (CK, Vàng, Crypto)\n' +
    '   • Ngân sách thông minh\n' +
    '   • Dashboard trực quan\n\n' +
    '📊 10 Sheet:\n' +
    '   THU • CHI • TRẢ NỢ • QUẢN LÝ NỢ\n' +
    '   CK • VÀNG • CRYPTO • ĐẦU TƯ KHÁC\n' +
    '   BUDGET • TỔNG QUAN\n\n' +
    '👨‍💻 Phát triển bởi: Claude & Nika\n' +
    '📅 Phiên bản: ' + APP_CONFIG.VERSION,
    ui.ButtonSet.OK
  );
}

// ==================== TIỆN ÍCH ====================

/**
 * Tìm kiếm giao dịch
 */
function searchTransaction() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'Tìm kiếm',
    '🔍 Tính năng đang phát triển...\n\n' +
    'Sẽ cho phép tìm kiếm giao dịch theo:\n' +
    '• Ngày\n' +
    '• Số tiền\n' +
    '• Ghi chú\n' +
    '• Loại giao dịch',
    ui.ButtonSet.OK
  );
}

/**
 * Xuất báo cáo PDF
 */
function exportToPDF() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'Xuất PDF',
    '📤 Tính năng đang phát triển...\n\n' +
    'Sẽ cho phép xuất báo cáo:\n' +
    '• Báo cáo tháng\n' +
    '• Báo cáo quý\n' +
    '• Báo cáo năm',
    ui.ButtonSet.OK
  );
}

/**
 * Xóa dữ liệu test
 */
function clearTestData() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Xác nhận',
    '⚠️ Bạn có chắc muốn xóa TẤT CẢ dữ liệu test?\n\n' +
    'Thao tác này KHÔNG THỂ HOÀN TÁC!',
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    // TODO: Implement clear test data
    ui.alert('Thông báo', '🔍 Tính năng đang phát triển...', ui.ButtonSet.OK);
  }
}

// ==================== UTILITY FUNCTIONS ====================

/**
 * Xác nhận trước khi khởi tạo sheet
 */
function confirmInitialize(sheetName) {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Xác nhận',
    `⚠️ Bạn có chắc muốn khởi tạo ${sheetName}?\n\n` +
    'Lưu ý: Nếu sheet đã tồn tại sẽ BỊ XÓA và tạo lại!',
    ui.ButtonSet.YES_NO
  );
  return response === ui.Button.YES;
}

/**
 * Hiển thị thông báo thành công
 */
function showSuccess(title, message) {
  SpreadsheetApp.getUi().alert(title, message, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Hiển thị thông báo lỗi
 */
function showError(title, message) {
  SpreadsheetApp.getUi().alert(
    title,
    `❌ ${message}\n\nVui lòng thử lại hoặc liên hệ hỗ trợ.`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Lấy spreadsheet hiện tại
 */
function getSpreadsheet() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * Lấy sheet theo tên
 */
function getSheet(sheetName) {
  return getSpreadsheet().getSheetByName(sheetName);
}

/**
 * Force recalculate toàn bộ sheet
 */
function forceRecalculate() {
  SpreadsheetApp.flush();
  getSpreadsheet().getSheets().forEach(sheet => {
    sheet.getDataRange().getValues();
  });
}

/**
 * Lấy ngày hiện tại
 */
function getCurrentDate() {
  return new Date();
}

/**
 * Format số tiền
 * @param {number} amount
 * @return {string}
 */
function formatCurrency(amount) {
  return new Intl.NumberFormat('vi-VN', {
    style: 'currency',
    currency: 'VND'
  }).format(amount);
}