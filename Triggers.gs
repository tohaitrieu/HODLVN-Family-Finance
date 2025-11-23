/**
 * ===============================================
 * TRIGGERS.GS
 * ===============================================
 * 
 * Xử lý các sự kiện tự động (Triggers)
 */

/**
 * Trigger chạy khi có thay đổi trong Spreadsheet
 * Dùng để cập nhật Budget và Dashboard khi nhập liệu thủ công
 * Đổi tên từ onEdit -> processEdit để dùng với Installable Trigger (cấp quyền UI)
 */
function processEdit(e) {
  try {
    if (!e) return;
    
    const range = e.range;
    const sheet = range.getSheet();
    const sheetName = sheet.getName();
    const row = range.getRow();
    const col = range.getColumn();
    
    // Bỏ qua header
    if (row < 2) return;
    
    // 1. Xử lý cập nhật BUDGET khi nhập liệu CHI TIÊU
    if (sheetName === APP_CONFIG.SHEETS.EXPENSE) {
      // Cột 4 là Danh mục (Category)
      if (col >= 2 && col <= 4) {
        Logger.log('🔄 Data changed in Expense sheet. Budget should update automatically via formulas.');
      }
    }
    
    // 2. Xử lý cập nhật BUDGET khi nhập liệu ĐẦU TƯ
    else if (sheetName === APP_CONFIG.SHEETS.STOCK || 
             sheetName === APP_CONFIG.SHEETS.GOLD || 
             sheetName === APP_CONFIG.SHEETS.CRYPTO || 
             sheetName === APP_CONFIG.SHEETS.OTHER_INVESTMENT) {
      Logger.log('🔄 Data changed in Investment sheet. Budget should update automatically via formulas.');
    }
    
    // 3. Xử lý cập nhật BUDGET khi TRẢ NỢ
    else if (sheetName === APP_CONFIG.SHEETS.DEBT_PAYMENT) {
      Logger.log('🔄 Data changed in Debt Payment sheet. Budget should update automatically via formulas.');
    }
    
    // 4. Cập nhật Dashboard (nếu cần)
    // Dashboard dùng công thức nên thường tự cập nhật.
    
    // 5. Xử lý Quick Actions trên Dashboard (Checkboxes)
    if (sheetName === APP_CONFIG.SHEETS.DASHBOARD) {
      handleDashboardAction(range);
    }
    
    // 6. Xử lý đồng bộ dữ liệu 2 chiều (Transaction ID)
    SyncManager.handleOnEdit(e);
    
  } catch (error) {
    Logger.log('❌ Lỗi processEdit: ' + error.message);
  }
}

/**
 * Xử lý Quick Actions trên Dashboard
 * @param {Range} range - Range được edit
 */
function handleDashboardAction(range) {
  const sheet = range.getSheet();
  const row = range.getRow();
  const col = range.getColumn();
  const a1Notation = range.getA1Notation();
  const value = range.getValue();
  
  // Chỉ xử lý khi Checkbox được tích (TRUE)
  if (value !== true) return;
  
  // Map A1Notation -> Function Name
  const actionMap = {
    'D2': 'showIncomeForm',
    'D4': 'showExpenseForm',
    'F2': 'showDebtManagementForm',
    'F4': 'showLendingForm',
    'H2': 'showGoldForm',
    'H4': 'showStockForm',
    'J2': 'showCryptoForm',
    'J4': 'showOtherInvestmentForm'
  };
  
  const functionName = actionMap[a1Notation];
  
  if (functionName) {
    // 1. Uncheck ngay lập tức để reset
    range.setValue(false);
    
    // 2. Gọi hàm hiển thị form
    try {
      switch (functionName) {
        case 'showIncomeForm': showIncomeForm(); break;
        case 'showExpenseForm': showExpenseForm(); break;
        case 'showDebtManagementForm': showDebtManagementForm(); break;
        case 'showLendingForm': showLendingForm(); break;
        case 'showGoldForm': showGoldForm(); break;
        case 'showStockForm': showStockForm(); break;
        case 'showCryptoForm': showCryptoForm(); break;
        case 'showOtherInvestmentForm': showOtherInvestmentForm(); break;
      }
      
      // Hiển thị thông báo nhỏ (Toast)
      SpreadsheetApp.getActive().toast('Đang mở form...', 'Hệ thống', 1);
      
    } catch (e) {
      Logger.log(`❌ Lỗi mở form ${functionName}: ${e.message}`);
      range.setNote(`Lỗi: ${e.message}`); // Ghi lỗi vào note để debug
    }
  }
}

/**
 * Trigger chạy khi có thay đổi cấu trúc (Insert/Remove Row, etc.)
 * Cần cài đặt thủ công hoặc qua hàm createInstallableTriggers()
 */
function onChange(e) {
  SyncManager.handleOnChange(e);
}

/**
 * Hàm cài đặt Trigger (Chạy 1 lần)
 */
function createInstallableTriggers() {
  const ss = SpreadsheetApp.getActive();
  
  // Xóa trigger cũ để tránh trùng lặp
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    const funcName = triggers[i].getHandlerFunction();
    if (funcName === 'onChange' || funcName === 'processEdit' || funcName === 'onEdit') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  // 1. Tạo trigger onChange
  ScriptApp.newTrigger('onChange')
      .forSpreadsheet(ss)
      .onChange()
      .create();
      
  // 2. Tạo trigger onEdit (Installable) -> Gọi processEdit
  // Điều này cho phép script mở Modal/Sidebar
  ScriptApp.newTrigger('processEdit')
      .forSpreadsheet(ss)
      .onEdit()
      .create();
      
  SpreadsheetApp.getUi().alert('✅ Đã cài đặt Trigger thành công! Các tính năng tự động và Quick Action sẽ hoạt động ngay.');
}
