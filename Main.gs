/**
 * ===============================================
 * MAIN.GS - FILE CHÍNH HỆ THỐNG QUẢN LÝ TÀI CHÍNH v3.0
 * ===============================================
 * 
 * Kiến trúc Module:
 * - Main.gs: Điều phối, Menu, UI
 * - SheetInitializer.gs: Khởi tạo các sheet
 * - FormHandlers.gs: Xử lý form nhập liệu
 * - DataProcessors.gs: Xử lý giao dịch
 * - BudgetManager.gs: Quản lý ngân sách
 * - DashboardManager.gs: Quản lý dashboard & thống kê
 */

// ==================== CẤU HÌNH TOÀN CỤC ====================

const APP_CONFIG = {
  VERSION: '3.0',
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

// ==================== KHỞI TẠO TẤT CẢ SHEET ====================

/**
 * Khởi tạo tất cả các sheet
 */
function initializeAllSheets() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Xác nhận khởi tạo',
    '⚠️ Bạn có chắc muốn khởi tạo TẤT CẢ sheet?\n\n' +
    'Lưu ý: Các sheet đã tồn tại sẽ BỊ XÓA và tạo lại từ đầu!',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    return;
  }
  
  try {
    const progress = [];
    
    // Khởi tạo từng sheet
    progress.push('📥 Khởi tạo Sheet THU...');
    initializeIncomeSheet(true);
    
    progress.push('📤 Khởi tạo Sheet CHI...');
    initializeExpenseSheet(true);
    
    progress.push('💳 Khởi tạo Sheet TRẢ NỢ...');
    initializeDebtPaymentSheet(true);
    
    progress.push('📊 Khởi tạo Sheet QUẢN LÝ NỢ...');
    initializeDebtManagementSheet(true);
    
    progress.push('📈 Khởi tạo Sheet CHỨNG KHOÁN...');
    initializeStockSheet(true);
    
    progress.push('🪙 Khởi tạo Sheet VÀNG...');
    initializeGoldSheet(true);
    
    progress.push('₿ Khởi tạo Sheet CRYPTO...');
    initializeCryptoSheet(true);
    
    progress.push('💼 Khởi tạo Sheet ĐẦU TƯ KHÁC...');
    initializeOtherInvestmentSheet(true);
    
    progress.push('💰 Khởi tạo Sheet BUDGET...');
    initializeBudgetSheet(true);
    
    progress.push('📊 Khởi tạo Sheet TỔNG QUAN...');
    initializeDashboardSheet(true);
    
    showSuccess(
      'Hoàn thành!',
      '✅ Đã khởi tạo thành công TẤT CẢ sheet!\n\n' +
      progress.join('\n') +
      '\n\n🎉 Hệ thống sẵn sàng sử dụng!'
    );
    
  } catch (error) {
    showError('Lỗi khởi tạo', error.message);
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
    '   Menu > Khởi tạo Sheet > Khởi tạo TẤT CẢ Sheet\n\n' +
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
 */
function formatCurrency(amount) {
  return amount.toLocaleString('vi-VN') + ' VNĐ';
}