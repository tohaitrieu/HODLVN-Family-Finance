/**
 * ===============================================
 * MAIN.GS - FILE CHÍNH HỆ THỐNG QUẢN LÝ TÀI CHÍNH v3.4.1
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
 * VERSION: 3.5.4 - Dashboard Quick Actions
 * CHANGELOG v3.5.4:
 * ✅ NEW: Quick Action Checkboxes on Dashboard
 * ✅ NEW: Trigger integration for fast data entry
 * 
 * CHANGELOG v3.5.3:
 * ✅ REFACTOR: Unified Debt & Lending Type System (Shared Logic)
 * ✅ FIX: Event Calendar for Payables & Receivables
 * ✅ UPDATE: Backward compatibility for old data
 * 
 * CHANGELOG v3.5.2:
 * ✅ NEW: Báo cáo chi phí chi tiết (Ngân sách, Còn lại, Trạng thái)
 * ✅ UPDATE: Tự động lấy dữ liệu từ Sheet BUDGET
 * 
 * CHANGELOG v3.5.1:
 * ✅ FIX: Dashboard Payables Logic (Installment Loans)
 * 
 * CHANGELOG v3.5.0:
 * ✅ NEW: Refactored Lending System with 3 specific types
 * ✅ NEW: Dynamic Event Calendar using Custom Functions
 * ✅ FIX: Debt Payment & Collection forms support precise amounts
 * ✅ FIX: Dashboard tables row offset issues
 * ✅ UPDATE: Event Calendar shows only nearest event per loan
 * 
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
  VERSION: '3.5.5',
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
    LENDING: 'CHO VAY',
    LENDING_REPAYMENT: 'THU NỢ',
    BUDGET: 'BUDGET',
    DASHBOARD: 'TỔNG QUAN',
    CHANGELOG: 'LỊCH SỬ CẬP NHẬT'
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
    NUMBER: '#,##0',
    PERCENTAGE: '0.00%',
    DATE: 'dd/MM/yyyy'
  },
  
  // Danh mục thu chi (Centralized Categories)
  CATEGORIES: {
    INCOME: [
      'Lương', 'MMO (Make Money Online)', 'Thưởng', 'Bán CK', 
      'Bán Vàng', 'Bán Crypto', 'Lãi đầu tư', 'Thu hồi nợ', 
      'Vay ngân hàng', 'Vay cá nhân', 'Khác'
    ],
    EXPENSE: [
      'Ăn uống', 'Đi lại', 'Nhà ở', 'Y tế', 
      'Giáo dục', 'Mua sắm', 'Giải trí', 'Cho vay', 'Trả nợ', 'Khác'
    ]
  }
};

// ==================== LOAN TYPES CONFIGURATION ====================

/**
 * Unified loan type system for both debt and lending
 * Uses type IDs with context-specific display names
 */
const LOAN_TYPES = {
  BULLET: {
    id: 'BULLET',
    debtLabel: 'Vay tất toán (Gốc + lãi cuối kỳ)',
    lendingLabel: 'Cho vay tất toán (Gốc + lãi cuối kỳ)',
    description: 'Trả toàn bộ gốc và lãi vào ngày đáo hạn',
    requiresMaturityDate: true,
    hasInterest: true
  },
  
  INTEREST_ONLY: {
    id: 'INTEREST_ONLY',
    debtLabel: 'Nợ ngân hàng (Trả lãi hàng tháng)',
    lendingLabel: 'Cho vay trả lãi định kỳ',
    description: 'Trả lãi hàng tháng, gốc vào cuối kỳ',
    requiresMaturityDate: true,
    hasInterest: true
  },
  
  EQUAL_PRINCIPAL: {
    id: 'EQUAL_PRINCIPAL',
    debtLabel: 'Vay trả góp (Gốc đều, lãi giảm dần)',
    lendingLabel: 'Cho vay trả góp (Gốc đều, lãi giảm dần)',
    description: 'Trả gốc đều hàng tháng, lãi tính trên số dư',
    requiresMaturityDate: false,
    hasInterest: true
  },
  
  EQUAL_PRINCIPAL_UPFRONT_FEE: {
    id: 'EQUAL_PRINCIPAL_UPFRONT_FEE',
    debtLabel: 'Trả góp qua thẻ (Phí ban đầu)',
    lendingLabel: 'Cho vay có phí trước (Gốc đều)',
    description: 'Phí/lãi tính một lần, gốc chia đều',
    requiresMaturityDate: false,
    hasInterest: true
  },
  
  INTEREST_FREE: {
    id: 'INTEREST_FREE',
    debtLabel: 'Trả góp miễn lãi (Chỉ trả gốc)',
    lendingLabel: 'Cho vay miễn lãi (Chỉ thu gốc)',
    description: 'Chỉ trả/thu gốc chia đều, không lãi',
    requiresMaturityDate: false,
    hasInterest: false
  },
  
  OTHER: {
    id: 'OTHER',
    debtLabel: 'Khác',
    lendingLabel: 'Khác',
    description: 'Loại khác (mặc định: tất toán)',
    requiresMaturityDate: true,
    hasInterest: true
  }
};

/**
 * Get loan type configuration by ID
 * @param {string} typeId - Type ID
 * @return {Object} Type configuration or OTHER if not found
 */
function getLoanType(typeId) {
  return LOAN_TYPES[typeId] || LOAN_TYPES.OTHER;
}

/**
 * Get display label for a loan type based on context
 * @param {string} typeId - Type ID
 * @param {boolean} isDebt - True for debt context, false for lending
 * @return {string} Display label
 */
function getLoanTypeLabel(typeId, isDebt = true) {
  const type = getLoanType(typeId);
  return isDebt ? type.debtLabel : type.lendingLabel;
}

/**
 * Map legacy type names to new type IDs (Backward compatibility)
 * Enhanced to handle full labels with parentheses
 * @param {string} typeName - Old type name, new type ID, or full label
 * @return {string} Type ID
 */
function mapLegacyTypeToId(typeName) {
  if (!typeName) return 'OTHER';
  
  // If already a valid ID, return as-is
  if (LOAN_TYPES[typeName]) {
    return typeName;
  }
  
  // Try exact match with legacy mapping first
  const legacyMapping = {
    // Legacy lending types (from LendingForm.html)
    'Tất toán gốc - lãi cuối kỳ': 'BULLET',
    'Trả lãi hàng tháng, gốc cuối kỳ': 'INTEREST_ONLY',
    'Trả góp gốc - lãi hàng tháng': 'EQUAL_PRINCIPAL',
    
    // Legacy debt types (from DebtManagementForm.html)
    'Nợ ngân hàng': 'INTEREST_ONLY',
    'Vay trả góp': 'EQUAL_PRINCIPAL',
    'Trả góp qua thẻ (Phí ban đầu)': 'EQUAL_PRINCIPAL_UPFRONT_FEE',
    'Trả góp qua thẻ (Lãi giảm dần)': 'EQUAL_PRINCIPAL',
    'Trả góp qua thẻ (Miễn lãi)': 'INTEREST_FREE',
    'Khác': 'OTHER'
  };
  
  if (legacyMapping[typeName]) {
    return legacyMapping[typeName];
  }
  
  // NEW: Try fuzzy matching with current labels (handle full labels with descriptions)
  // Check if typeName matches any debtLabel or lendingLabel
  for (const [id, config] of Object.entries(LOAN_TYPES)) {
    if (typeName === config.debtLabel || typeName === config.lendingLabel) {
      return id;
    }
  }
  
  // NEW: Try partial matching (e.g., "Vay trả góp (Gốc đều, lãi giảm dần)" contains "Vay trả góp")
  const typeNameLower = typeName.toLowerCase();
  
  // Match patterns
  if (typeNameLower.includes('vay trả góp') || typeNameLower.includes('cho vay trả góp')) {
    if (typeNameLower.includes('miễn lãi')) {
      return 'INTEREST_FREE';
    } else if (typeNameLower.includes('phí ban đầu')) {
      return 'EQUAL_PRINCIPAL_UPFRONT_FEE';
    } else {
      return 'EQUAL_PRINCIPAL';
    }
  }
  
  if (typeNameLower.includes('nợ ngân hàng') || typeNameLower.includes('trả lãi hàng tháng') || typeNameLower.includes('trả lãi định kỳ')) {
    return 'INTEREST_ONLY';
  }
  
  if (typeNameLower.includes('tất toán') || typeNameLower.includes('gốc + lãi cuối kỳ') || typeNameLower.includes('gốc, lãi cuối kỳ')) {
    return 'BULLET';
  }
  
  return 'OTHER';
}

// ==================== MENU CHÍNH ====================

/**
 * Tạo menu khi mở file
 */
function onOpen() {
  // Kiểm tra và cập nhật Changelog nếu có phiên bản mới
  ChangelogManager.checkVersionAndUpdate();

  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu(APP_CONFIG.APP_NAME)
    // === NHÓM 1: NHẬP LIỆU ===
    .addSubMenu(ui.createMenu('📝 Nhập liệu')
      .addItem('📥 Nhập Thu nhập', 'showIncomeForm')
      .addItem('📤 Nhập Chi tiêu', 'showExpenseForm')
      .addItem('💳 Thêm Khoản Nợ', 'showDebtManagementForm')
      .addItem('💳 Trả nợ', 'showDebtPaymentForm')
      .addItem('🤝 Cho vay', 'showLendingForm')
      .addItem('💰 Thu nợ & Lãi', 'showLendingPaymentForm')
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
      .addItem('♻️ Cập nhật cấu trúc dữ liệu', 'runMigrations')
      .addSeparator()
      .addItem('📅 Lịch trả nợ dự kiến', 'showDebtScheduleReport')
      .addSeparator()
      .addItem('📅 Xem Tất cả', 'viewAll')
      .addItem('📊 Xem Năm hiện tại', 'viewCurrentYear')
      .addItem('🗓️ Xem Quý hiện tại', 'viewCurrentQuarter')
      .addItem('📆 Xem Tháng hiện tại', 'viewCurrentMonth'))
    
    .addSeparator()
    
    // === NHÓM 4: TIỆN ÍCH ===
    .addSubMenu(ui.createMenu('🛠️ Tiện ích')
      .addItem('✨ Chuẩn hóa dữ liệu', 'normalizeAllData')
      .addItem('🧹 Dọn dẹp dữ liệu mồ côi', 'cleanOrphans')
      .addItem('🔍 Tìm kiếm giao dịch', 'searchTransaction')
      .addItem('📤 Xuất báo cáo PDF', 'exportToPDF')
      .addSeparator()
      .addItem('🐛 Debug - Kiểm tra sự kiện', 'debugEventCalculation')
      .addItem('🔧 Sửa lỗi định dạng cột Gốc', 'fixPrincipalColumnFormat')
      .addSeparator()
      .addItem('🗑️ Xóa dữ liệu test', 'clearTestData'))
    
    .addSeparator()

    // === NHÓM 5: KHỞI TẠO SHEET ===
    .addSubMenu(ui.createMenu('⚙️ Khởi tạo Sheet')
      .addItem('🔄 Cập nhật toàn bộ các Sheet', 'updateAllSheets')
      .addSeparator()
      .addItem('🚀 Khởi tạo TẤT CẢ Sheet (Setup Wizard)', 'initializeAllSheets')
      .addSeparator()
      .addItem('📥 Khởi tạo Sheet THU', 'initializeIncomeSheet')
      .addItem('📤 Khởi tạo Sheet CHI', 'initializeExpenseSheet')
      .addItem('💳 Khởi tạo Sheet TRẢ NỢ', 'initializeDebtPaymentSheet')
      .addItem('📊 Khởi tạo Sheet QUẢN LÝ NỢ', 'initializeDebtManagementSheet')
      .addItem('🤝 Khởi tạo Sheet CHO VAY', 'initializeLendingSheet')
      .addItem('💰 Khởi tạo Sheet THU NỢ', 'initializeLendingRepaymentSheet')
      .addSeparator()
      .addItem('📈 Khởi tạo Sheet CHỨNG KHOÁN', 'initializeStockSheet')
      .addItem('🪙 Khởi tạo Sheet VÀNG', 'initializeGoldSheet')
      .addItem('₿ Khởi tạo Sheet CRYPTO', 'initializeCryptoSheet')
      .addItem('💼 Khởi tạo Sheet ĐẦU TƯ KHÁC', 'initializeOtherInvestmentSheet')
      .addSeparator()
      .addItem('💰 Khởi tạo Sheet BUDGET', 'initializeBudgetSheet')
      .addItem('📊 Khởi tạo Sheet TỔNG QUAN', 'initializeDashboardSheet')
      .addSeparator()
      .addItem('Sửa lỗi lệch cột (Nợ/Cho vay)', 'fixColumnAlignment'))
    
    .addSeparator()
    
    // === NHÓM 6: TRỢ GIÚP ===
    .addItem('ℹ️ Hướng dẫn sử dụng', 'showInstructions')
    .addItem('📜 Lịch sử cập nhật', 'updateChangelog')
    .addItem('📖 Giới thiệu hệ thống', 'showAbout')
    
    .addToUi();
}

/**
 * ⭐ MENU FUNCTION: Cập nhật Dashboard thủ công
 */
function refreshDashboard() {
  DashboardManager.setupDashboard();
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
 * Hiển thị form cho vay
 */
function showLendingForm() {
  showForm('LendingForm', '🤝 Thêm Khoản Cho Vay Mới', 600, 700);
}

/**
 * Hiển thị form thu nợ
 */
function showLendingPaymentForm() {
  showForm('LendingPaymentForm', '💰 Thu Hồi Nợ & Lãi', 600, 650);
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

// ==================== BUDGET MENU FUNCTIONS - NEW v3.4.1 ====================

/**
 * ✅ NEW v3.4.1: Hiển thị form đặt ngân sách tháng
 */
function showSetBudgetForm() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('SetBudgetForm')
      .setWidth(700)
      .setHeight(650);
    SpreadsheetApp.getUi().showModalDialog(html, '📊 Đặt Ngân sách tháng');
  } catch (error) {
    showError('Không thể mở form đặt ngân sách', error.message);
  }
}

/**
 * ✅ NEW v3.4.1: Wrapper function để gọi BudgetManager.checkBudgetWarnings()
 */
function checkBudgetWarnings() {
  BudgetManager.checkBudgetWarnings();
}

/**
 * ✅ NEW v3.4.1: Wrapper function để gọi BudgetManager.showExpenseReport()
 */
function showExpenseReport() {
  BudgetManager.showExpenseReport();
}

/**
 * ✅ NEW v3.4.1: Wrapper function để gọi BudgetManager.showInvestmentReport()
 */
function showInvestmentReport() {
  BudgetManager.showInvestmentReport();
}

/**
 * ✅ NEW v3.4.1: Xử lý dữ liệu từ SetBudgetForm
 * Hàm này được gọi từ SetBudgetForm.html qua google.script.run
 * @param {Object} budgetData - Dữ liệu ngân sách từ form
 * @return {Object} {success, message}
 */
function setBudgetForMonth(budgetData) {
  return BudgetManager.setBudgetForMonth(budgetData);
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
 * Cập nhật toàn bộ các Sheet (Giữ nguyên dữ liệu)
 */
function updateAllSheets() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
    'Cập nhật toàn bộ Sheet',
    'Bạn có chắc chắn muốn cập nhật lại cấu trúc (header, format, công thức) cho TẤT CẢ các sheet?\n\n' +
    'Lưu ý:\n' +
    '- Dữ liệu hiện có sẽ ĐƯỢC GIỮ NGUYÊN.\n' +
    '- Header, độ rộng cột, format sẽ được reset về mặc định.\n' +
    '- Công thức cột tính toán sẽ được áp dụng lại.',
    ui.ButtonSet.YES_NO
  );

  if (result === ui.Button.YES) {
    try {
      const toastId = SpreadsheetApp.getActive().toast('Đang cập nhật...', 'Hệ thống', -1);
      
      SheetInitializer.updateAllSheets();
      DashboardManager.setupDashboard(); // Cập nhật cả Dashboard
      
      SpreadsheetApp.getActive().toast('Đã cập nhật xong!', 'Hệ thống', 3);
      
      ui.alert('Thành công', '✅ Đã cập nhật toàn bộ các Sheet thành công!', ui.ButtonSet.OK);
    } catch (error) {
      showError('Lỗi cập nhật', error.message);
    }
  }
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
    
    initializeIncomeSheet(true);
    initializeExpenseSheet(true);
    initializeDebtPaymentSheet(true);
    initializeDebtManagementSheet(true);
    initializeStockSheet(true);
    initializeGoldSheet(true);
    initializeCryptoSheet(true);
    initializeOtherInvestmentSheet(true);
    initializeLendingRepaymentSheet(true);
    initializeBudgetSheet(true); // ✅ v3.4: Gọi hàm chuẩn từ SheetInitializer
    initializeDashboardSheet(true);
    
    // ============================================================
    // BƯỚC 2: Thêm dữ liệu số dư ban đầu vào sheet THU
    // ============================================================
    
    const incomeSheet = getSheet(APP_CONFIG.SHEETS.INCOME);
    if (incomeSheet && setupData.balance) {
      const balance = setupData.balance;
      const balanceRow = [
        new Date(), // A: Mã GD (auto)
        new Date(balance.date), // B: Ngày
        balance.amount, // C: Số tiền
        'Số dư đầu kỳ', // D: Nguồn
        balance.source || 'Thiết lập ban đầu', // E: Ghi chú
        'Xác nhận' // F: Trạng thái
      ];
      
      // Tìm dòng trống đầu tiên (sau header)
      const lastRow = incomeSheet.getLastRow();
      incomeSheet.getRange(lastRow + 1, 1, 1, 6).setValues([balanceRow]);
      
      Logger.log('✅ Đã thêm số dư ban đầu vào sheet THU');
    }
    
    // ============================================================
    // BƯỚC 3: Thêm khoản nợ vào sheet QUẢN LÝ NỢ (nếu có)
    // ============================================================
    
    if (setupData.debt) {
      const debtSheet = getSheet(APP_CONFIG.SHEETS.DEBT_MANAGEMENT);
      if (debtSheet) {
        const debt = setupData.debt;
        
        // Tính lãi hàng tháng
        const monthlyRate = (debt.rate / 100) / 12;
        const monthlyPayment = (debt.principal * monthlyRate * Math.pow(1 + monthlyRate, debt.term)) / 
                               (Math.pow(1 + monthlyRate, debt.term) - 1);
        
        const debtRow = [
          new Date(), // A: Mã nợ (timestamp)
          debt.name, // B: Tên khoản nợ
          new Date(debt.date), // C: Ngày vay
          debt.principal, // D: Số tiền gốc
          debt.rate / 100, // E: Lãi suất (chia 100 để chuyển từ % sang decimal)
          debt.term, // F: Kỳ hạn
          monthlyPayment, // G: Trả hàng tháng
          0, // H: Đã trả
          debt.principal, // I: Còn lại
          'Chưa trả', // J: Trạng thái - ✅ v3.3: Đổi từ "Đang trả" sang "Chưa trả"
          '' // K: Ghi chú
        ];
        
        // ✅ v3.3: Sử dụng getRange().setValues() thay vì appendRow()
        const lastRow = debtSheet.getLastRow();
        debtSheet.getRange(lastRow + 1, 1, 1, 11).setValues([debtRow]);
        
        Logger.log('✅ Đã thêm khoản nợ vào sheet QUẢN LÝ NỢ');
        
        // ============================================================
        // ✅ v3.3: TỰ ĐỘNG THÊM KHOẢN THU KHI THÊM NỢ
        // ============================================================
        // Lý do: Khi vay nợ = nhận tiền vào = tăng thu nhập
        
        if (incomeSheet) {
          const debtIncomeRow = [
            new Date(), // A: Mã GD
            new Date(debt.date), // B: Ngày vay
            debt.principal, // C: Số tiền vay
            'Vay nợ', // D: Nguồn thu
            `Khoản vay: ${debt.name}`, // E: Ghi chú
            'Xác nhận' // F: Trạng thái
          ];
          
          const lastIncomeRow = incomeSheet.getLastRow();
          incomeSheet.getRange(lastIncomeRow + 1, 1, 1, 6).setValues([debtIncomeRow]);
          
          Logger.log('✅ Đã tự động thêm khoản thu tương ứng với khoản vay');
        }
      }
    }
    
    // ============================================================
    // BƯỚC 4: ✅ v3.4 - CẬP NHẬT BUDGET VỚI PHÂN BỔ % TỪ WIZARD
    // ============================================================
    
    if (setupData.budget) {
      const budgetSheet = getSheet(APP_CONFIG.SHEETS.BUDGET);
      if (budgetSheet) {
        const budget = setupData.budget;
        const income = budget.income;
        
        // Tính số tiền cho mỗi nhóm
        const totalChi = income * (budget.pctChi / 100);
        const totalDautu = income * (budget.pctDautu / 100);
        const totalTrano = income * (budget.pctTrano / 100);
        
        // === CẬP NHẬT THU NHẬP & CÁC THAM SỐ ===
        budgetSheet.getRange('B2').setValue(income); // Thu nhập tháng
        budgetSheet.getRange('B3').setValue(budget.pctChi / 100); // % Chi
        budgetSheet.getRange('B4').setValue(budget.pctDautu / 100); // % Đầu tư
        budgetSheet.getRange('B5').setValue(budget.pctTrano / 100); // % Trả nợ
        
        // === CẬP NHẬT CHI TIÊU (Row 7-16) ===
        const chiCategories = [
          'Ăn uống', 'Đi lại', 'Nhà ở', 'Điện nước', 'Viễn thông',
          'Giáo dục', 'Y tế', 'Mua sắm', 'Giải trí', 'Khác'
        ];
        
        chiCategories.forEach((cat, idx) => {
          const row = 7 + idx;
          const pct = budget.chi[cat] || 0;
          const amount = totalChi * (pct / 100);
          budgetSheet.getRange(row, 2).setValue(amount); // Cột B: Mục tiêu
        });
        
        // === CẬP NHẬT ĐẦU TƯ (Row 25-28) ===
        const investMap = {
          'Chứng khoán': 25,
          'Vàng': 26,
          'Crypto': 27,
          'Đầu tư khác': 28
        };
        
        Object.entries(investMap).forEach(([type, row]) => {
          const pct = budget.dautu[type] || 0;
          const amount = totalDautu * (pct / 100);
          budgetSheet.getRange(row, 2).setValue(amount); // Cột B: Mục tiêu
        });
        
        // === CẬP NHẬT TRẢ NỢ (Row 19) ===
        budgetSheet.getRange('B19').setValue(totalTrano); // Trả nợ gốc
        
        Logger.log('✅ Đã cập nhật BUDGET với phân bổ % từ Setup Wizard');
      }
    }
    
    // ============================================================
    // BƯỚC 5: Cập nhật Dashboard
    // ============================================================
    
    DashboardManager.setupDashboard();
    Logger.log('✅ Đã cập nhật Dashboard');
    
    // ============================================================
    // BƯỚC 6: Sắp xếp lại thứ tự Sheet
    // ============================================================
    
    SheetInitializer.reorderSheets();
    Logger.log('✅ Đã sắp xếp lại thứ tự Sheet');
    
    // ============================================================
    // BƯỚC 7: Trả về kết quả
    // ============================================================
    
    return {
      success: true,
      message: '✅ Khởi tạo hệ thống thành công!\n\n' +
        `💰 Số dư ban đầu: ${formatCurrency(setupData.balance.amount)}\n` +
        (setupData.debt ? `💳 Khoản nợ: ${setupData.debt.name} - ${formatCurrency(setupData.debt.principal)}\n` : '') +
        `\n📊 PHÂN BỔ NGÂN SÁCH:\n` +
        `📤 Chi tiêu: ${setupData.budget.pctChi}% = ${formatCurrency(setupData.budget.income * setupData.budget.pctChi / 100)}\n` +
        `💼 Đầu tư: ${setupData.budget.pctDautu}% = ${formatCurrency(setupData.budget.income * setupData.budget.pctDautu / 100)}\n` +
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
 * Khởi tạo Sheet CHO VAY
 */
function initializeLendingSheet(skipConfirm) {
  if (!skipConfirm && !confirmInitialize('Sheet CHO VAY')) return;
  SheetInitializer.initializeLendingSheet();
  if (!skipConfirm) showSuccess('Thành công', '✅ Đã khởi tạo Sheet CHO VAY!');
}

/**
 * Khởi tạo Sheet THU NỢ
 */
function initializeLendingRepaymentSheet(skipConfirm) {
  if (!skipConfirm && !confirmInitialize('Sheet THU NỢ')) return;
  SheetInitializer.initializeLendingRepaymentSheet();
  if (!skipConfirm) showSuccess('Thành công', '✅ Đã khởi tạo Sheet THU NỢ!');
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
  const htmlContent = `
    <style>
      body { font-family: sans-serif; padding: 15px; line-height: 1.5; font-size: 14px; }
      h3 { color: #4472C4; margin: 0 0 15px 0; border-bottom: 1px solid #ddd; padding-bottom: 10px; }
      .group { margin-bottom: 15px; }
      .label { font-weight: bold; color: #333; }
      ul { margin: 5px 0; padding-left: 20px; }
      li { margin-bottom: 2px; }
      .footer { margin-top: 20px; border-top: 1px solid #eee; padding-top: 15px; color: #666; font-size: 13px; }
      a { color: #4472C4; text-decoration: none; }
    </style>
    
    <h3>💰 ${APP_CONFIG.APP_NAME} v${APP_CONFIG.VERSION}</h3>
    
    <div class="group">
      <div class="label">✨ MỚI TRONG v${APP_CONFIG.VERSION}:</div>
      <ul>
        <li>Fix các menu Budget không hoạt động</li>
        <li>Thêm form đặt ngân sách tháng</li>
        <li>Báo cáo chi tiêu và đầu tư</li>
      </ul>
    </div>
    
    <div class="group">
      <div class="label">🎯 Tính năng:</div>
      <ul>
        <li>Quản lý thu chi, nợ & lãi</li>
        <li>Đầu tư (CK, Vàng, Crypto)</li>
        <li>Ngân sách & Dashboard</li>
      </ul>
    </div>

    <div class="footer">
      👨‍💻 Phát triển bởi: Tô Triều với ❤️ từ <a href="https://hodl.vn" target="_blank"><b>HODL.VN</b></a><br>
      📅 Phiên bản: ${APP_CONFIG.VERSION}
    </div>
  `;
  
  const html = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(400)
    .setHeight(450);
    
  SpreadsheetApp.getUi().showModalDialog(html, 'Giới thiệu hệ thống');
}

/**
 * Cập nhật Changelog thủ công
 */
function updateChangelog() {
  ChangelogManager.updateChangelogSheet();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(APP_CONFIG.SHEETS.CHANGELOG);
  if (sheet) {
    sheet.activate();
  }
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
 * Dọn dẹp dữ liệu mồ côi
 */
function cleanOrphans() {
  SyncManager.cleanOrphans();
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