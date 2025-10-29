/**
 * ===============================================
 * MAIN.GS - FILE CHÍNH HỆ THỐNG QUẢN LÝ TÀI CHÍNH v3.4
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
 * VERSION: 3.4 - Unified Budget Sheet Structure
 * CHANGELOG v3.4:
 * - THỐNG NHẤT cấu trúc Sheet BUDGET cho cả Setup Wizard và Menu khởi tạo
 * - processSetupWizard() giờ GỌI SheetInitializer.initializeBudgetSheet() thay vì tự tạo sheet
 * - Điền dữ liệu từ Setup Wizard vào các ô tham số (B2, B3, B16, B25, etc.)
 * - ĐẢM BẢO cả 2 cách đều tạo CÙNG CẤU TRÚC: 6 cột với công thức động
 * - FIX vấn đề khác biệt giữa 2 sheet gây lỗi Dashboard và các công thức
 * 
 * CHANGELOG v3.3:
 * - Fix lỗi appendRow() trong processSetupWizard chèn dữ liệu vào dòng 1001
 * - Thay appendRow() bằng getRange().setValues() để insert đúng vị trí
 * - Sửa trạng thái ban đầu từ "Đang trả" thành "Chưa trả"
 * - Fix format lãi suất (chia 100 để chuyển từ % sang decimal)
 * - Thêm tính năng tự động thêm khoản thu khi thêm nợ trong Setup Wizard
 */

// ==================== CẤU HÌNH TOÀN CỤC ====================

const APP_CONFIG = {
  VERSION: '3.4',
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
      .addItem('📊 Nhập Cổ tức', 'showDividendForm')
      .addItem('🪙 Giao dịch Vàng', 'showGoldForm')
      .addItem('₿ Giao dịch Crypto', 'showCryptoForm')
      .addItem('💼 Đầu tư khác', 'showOtherInvestmentForm'))
    
    .addSeparator()

    // === NHÓM 2: BUDGET ===
    .addSubMenu(ui.createMenu('💵 Ngân sách')
      .addItem('📊 Đặt Ngân sách tháng', 'showSetBudgetForm')
      .addSeparator()
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
  showForm('IncomeForm', '📥 Nhập Thu nhập', 600, 650);
}

/**
 * Hiển thị form nhập chi
 */
function showExpenseForm() {
  showForm('ExpenseForm', '📤 Nhập Chi tiêu', 600, 650);
}

/**
 * Hiển thị form trả nợ
 */
function showDebtPaymentForm() {
  showForm('DebtPaymentForm', '💳 Trả nợ', 600, 650);
}
/**
 * Hiển thị form thêm khoản nợ mới
 */
function showDebtManagementForm() {
  showForm('DebtManagementForm', '💳 Thêm Khoản Nợ Mới', 600, 700);
}
/**
 * Hiển thị form giao dịch chứng khoán
 */
function showStockForm() {
  showForm('StockForm', '📈 Giao dịch Chứng khoán', 600, 650);
}
/**
 * ✅ NEW v3.4: Hiển thị form nhập cổ tức
 */
function showDividendForm() {
  showForm('DividendForm', '📊 Nhập Cổ tức', 600, 650);
}
/**
 * Hiển thị form giao dịch vàng
 */
function showGoldForm() {
  showForm('GoldForm', '🪙 Giao dịch Vàng', 600, 650);
}

/**
 * Hiển thị form giao dịch crypto
 */
function showCryptoForm() {
  showForm('CryptoForm', '₿ Giao dịch Crypto', 600, 650);
}

/**
 * Hiển thị form đầu tư khác
 */
function showOtherInvestmentForm() {
  showForm('OtherInvestmentForm', '💼 Đầu tư khác', 600, 650);
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
    SpreadsheetApp.getUi().showModalDialog(html, '🚀 Thiết lập Hệ thống Quản lý Tài chính v3.3');
  } catch (error) {
    showError('Không thể mở Setup Wizard', 
      'Vui lòng đảm bảo file SetupWizard.html đã được tạo trong Apps Script.\n\n' +
      'Lỗi: ' + error.message);
  }
}

/**
 * ===============================================
 * PROCESSETUPWIZARD - VERSION MỚI v3.4
 * ===============================================
 * 
 * Xử lý dữ liệu từ Setup Wizard với phân bổ ngân sách theo %
 * 
 * THAY ĐỔI SO VỚI v3.3:
 * - Nhận dữ liệu thu nhập + phân bổ % (Chi/Đầu tư/Trả nợ)
 * - Tính toán ngân sách chi tiết dựa trên %
 * - Cập nhật sheet BUDGET với giá trị đã tính
 */

/**
 * Xử lý dữ liệu từ Setup Wizard
 * @param {Object} setupData - Dữ liệu từ form wizard
 * setupData = {
 *   balance: { amount, date, source },
 *   debt: { name, principal, rate, term, date } hoặc null,
 *   budget: {
 *     income: số thu nhập,
 *     pctChi: % chi tiêu,
 *     pctDautu: % đầu tư,
 *     pctTrano: % trả nợ,
 *     chi: { "Ăn uống": 35, "Đi lại": 10, ... },
 *     dautu: { "Chứng khoán": 50, "Vàng": 20, ... }
 *   }
 * }
 * @return {Object} Kết quả thực thi
 */
function processSetupWizard(setupData) {
  try {
    const ss = getSpreadsheet();
    
    // ============================================================
    // BƯỚC 1: Khởi tạo tất cả sheets (KHÔNG có dữ liệu mẫu)
    // ============================================================
    
    // Sheet THU - Khởi tạo thủ công
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
    
    const sourceRange = sheet.getRange('D2:D1000');
    sourceRange.setNumberFormat('@');
    
    // Sheet CHI - Khởi tạo thủ công
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
    
    incomeSheet.getRange(2, 1, 1, incomeData.length).setValues([incomeData]);
    
    // Format dòng vừa thêm
    incomeSheet.getRange(2, 2).setNumberFormat('dd/mm/yyyy');
    incomeSheet.getRange(2, 3).setNumberFormat('#,##0');
    
    // ============================================================
    // BƯỚC 3: Thêm khoản nợ (nếu có)
    // ============================================================
    if (setupData.debt) {
      const debtSheet = ss.getSheetByName(APP_CONFIG.SHEETS.DEBT_MANAGEMENT);
      const startDate = new Date(setupData.debt.date);
      const dueDate = new Date(startDate);
      dueDate.setMonth(dueDate.getMonth() + setupData.debt.term);
      
      // Phần 1: Cột A-I (STT đến Đã trả lãi)
      const debtDataPart1 = [
        1, // A: STT
        setupData.debt.name, // B: Tên khoản nợ
        setupData.debt.principal, // C: Gốc
        setupData.debt.rate / 100, // D: Lãi suất (chuyển % sang decimal)
        setupData.debt.term, // E: Kỳ hạn
        startDate, // F: Ngày vay
        dueDate, // G: Đáo hạn
        0, // H: Đã trả gốc
        0  // I: Đã trả lãi
      ];
      
      // Phần 2: Cột K-L (Trạng thái và Ghi chú)
      const debtDataPart2 = [
        'Chưa trả', // K: Trạng thái
        'Khởi tạo từ Setup Wizard' // L: Ghi chú
      ];
      
      // Insert Phần 1: Cột A-I
      debtSheet.getRange(2, 1, 1, debtDataPart1.length).setValues([debtDataPart1]);
      
      // Insert Phần 2: Cột K-L (bắt đầu từ cột 11)
      debtSheet.getRange(2, 11, 1, debtDataPart2.length).setValues([debtDataPart2]);
      
      // Format các cột
      debtSheet.getRange(2, 3).setNumberFormat('#,##0');
      debtSheet.getRange(2, 4).setNumberFormat('0.00"%"');
      debtSheet.getRange(2, 6).setNumberFormat('dd/mm/yyyy');
      debtSheet.getRange(2, 7).setNumberFormat('dd/mm/yyyy');
      debtSheet.getRange(2, 8).setNumberFormat('#,##0');
      debtSheet.getRange(2, 9).setNumberFormat('#,##0');
      
      // Tự động thêm khoản thu tương ứng
      const incomeRow = 3;
      const autoIncomeData = [
        2, // STT = 2
        startDate,
        setupData.debt.principal,
        'Vay nợ',
        `Vay: ${setupData.debt.name}`
      ];
      
      incomeSheet.getRange(incomeRow, 1, 1, autoIncomeData.length).setValues([autoIncomeData]);
      incomeSheet.getRange(incomeRow, 2).setNumberFormat('dd/mm/yyyy');
      incomeSheet.getRange(incomeRow, 3).setNumberFormat('#,##0');
    }
    
    // ============================================================
    // BƯỚC 3.5: Set Data Validation cho sheet THU
    // ============================================================
    const incomeSourceRange = incomeSheet.getRange('D2:D1000');
    const incomeSourceRule = SpreadsheetApp.newDataValidation()
      .requireValueInList([
        'Lương',
        'MMO (Make Money Online)',
        'Thưởng',
        'Bán CK',
        'Bán Vàng',
        'Bán Crypto',
        'Lãi đầu tư',
        'Thu hồi nợ',
        'Vay nợ',
        'Khác'
      ])
      .setAllowInvalid(true)
      .build();
    incomeSourceRange.setDataValidation(incomeSourceRule);
    
    // ============================================================
    // BƯỚC 4: Khởi tạo Budget bằng SheetInitializer và điền dữ liệu từ Setup Wizard
    // ============================================================
    
    // Gọi hàm khởi tạo chuẩn từ SheetInitializer
    SheetInitializer.initializeBudgetSheet();
    
    // Lấy sheet vừa tạo và điền dữ liệu từ Setup Wizard
    const budgetSheet = ss.getSheetByName(APP_CONFIG.SHEETS.BUDGET);
    
    // Điền Thu nhập dự kiến (cell B2)
    budgetSheet.getRange('B2').setValue(setupData.budget.income);
    
    // Điền % Nhóm Chi tiêu (cell B3)
    budgetSheet.getRange('B3').setValue(setupData.budget.pctChi / 100);
    
    // Điền % chi tiết cho từng danh mục chi tiêu (rows 6-13, column B)
    // Cần map tên danh mục từ setupData với vị trí row
    const chiMapping = {
      'Ăn uống': 6,
      'Đi lại': 7,
      'Nhà ở': 8,
      'Y tế': 9,
      'Giáo dục': 10,
      'Mua sắm': 11,
      'Giải trí': 12,
      'Khác': 13
    };
    
    Object.entries(setupData.budget.chi).forEach(([categoryName, pct]) => {
      const row = chiMapping[categoryName];
      if (row) {
        budgetSheet.getRange(row, 2).setValue(pct / 100); // Chuyển % sang decimal
      }
    });
    
    // Điền % Nhóm Đầu tư (cell B16 - chiEndRow + 2 = 14 + 2 = 16)
    budgetSheet.getRange('B16').setValue(setupData.budget.pctDautu / 100);
    
    // Điền % chi tiết cho đầu tư (rows 19-22, column B)
    const dautuMapping = {
      'Chứng khoán': 19,
      'Vàng': 20,
      'Crypto': 21,
      'Đầu tư khác': 22
    };
    
    Object.entries(setupData.budget.dautu).forEach(([categoryName, pct]) => {
      const row = dautuMapping[categoryName];
      if (row) {
        budgetSheet.getRange(row, 2).setValue(pct / 100); // Chuyển % sang decimal
      }
    });
    
    // Điền % Nhóm Trả nợ (cell B25 - tranoRow)
    budgetSheet.getRange('B25').setValue(setupData.budget.pctTrano / 100);
    
    // ============================================================
    // BƯỚC 5: Khởi tạo Dashboard
    // ============================================================
    DashboardManager.setupDashboard();
    
    return {
      success: true,
      message: '✅ Khởi tạo thành công!\n\n' +
        `💰 Thu nhập: ${formatCurrency(setupData.budget.income)}\n` +
        `📤 Chi tiêu: ${setupData.budget.pctChi}% = ${formatCurrency(setupData.budget.income * setupData.budget.pctChi / 100)}\n` +
        `💰 Đầu tư: ${setupData.budget.pctDautu}% = ${formatCurrency(setupData.budget.income * setupData.budget.pctDautu / 100)}\n` +
        `💳 Trả nợ: ${setupData.budget.pctTrano}% = ${formatCurrency(setupData.budget.income * setupData.budget.pctTrano / 100)}\n\n` +
        'Hệ thống đã sẵn sàng sử dụng!'
    };
    
  } catch (error) {
    Logger.log('Error in processSetupWizard: ' + error.message);
    return {
      success: false,
      message: 'Có lỗi xảy ra: ' + error.message
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
    '✨ MỚI TRONG v3.3:\n' +
    '   • Fix lỗi Setup Wizard chèn dữ liệu sai vị trí\n' +
    '   • Tự động thêm khoản thu khi thêm nợ\n' +
    '   • Cải thiện logic trạng thái nợ\n\n' +
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