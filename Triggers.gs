/**
 * ===============================================
 * TRIGGERS.GS
 * ===============================================
 * 
 * Xử lý các sự kiện tự động (Triggers)
 */

/**
 * Simple Trigger - Tự động chạy khi có edit
 * ĐÃ SỬA: Bỏ qua sheet Dashboard để tránh xung đột với Trigger cài đặt
 */
function onEdit(e) {
  // Kiểm tra nếu biến e không tồn tại (khi chạy thủ công)
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  
  // QUAN TRỌNG: Nếu đang ở Dashboard, dừng ngay lập tức!
  // Để nhường quyền cho Installable Trigger (processEdit) xử lý việc mở Form.
  if (sheet.getName() === APP_CONFIG.SHEETS.DASHBOARD) {
    return;
  }

  processEdit(e);
}

/**
 * Trigger chính xử lý các thay đổi trong Spreadsheet
 * Dùng để cập nhật Budget và Dashboard khi nhập liệu thủ công
 * Có thể dùng với Simple Trigger (onEdit) hoặc Installable Trigger (cấp quyền UI)
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
    
    // 4. Cập nhật Dashboard tự động khi có thay đổi dữ liệu
    if (sheetName === APP_CONFIG.SHEETS.INCOME || 
        sheetName === APP_CONFIG.SHEETS.EXPENSE ||
        sheetName === APP_CONFIG.SHEETS.DEBT_MANAGEMENT ||
        sheetName === APP_CONFIG.SHEETS.DEBT_PAYMENT ||
        sheetName === APP_CONFIG.SHEETS.LENDING) {
      // Trigger quick refresh for better custom function updates
      Logger.log(`🔄 Triggering dashboard refresh due to data change in ${sheetName}`);
      try {
        Utilities.sleep(200); // Small delay to ensure data is saved
        _quickRefreshCustomFunctions();
      } catch (error) {
        Logger.log('⚠️ Could not auto-refresh dashboard: ' + error.message);
      }
    }
    
    // 5. Xử lý thay đổi bộ lọc trên Dashboard (B2:B4) 
    if (sheetName === APP_CONFIG.SHEETS.DASHBOARD) {
      // Check if filter cells B2:B4 were changed
      if ((col >= 2 && col <= 4) && (row >= 2 && row <= 4)) {
        Logger.log(`🔄 Filter changed in Dashboard (${range.getA1Notation()}). Triggering refresh...`);
        try {
          Utilities.sleep(100); // Small delay to ensure filter value is saved
          _quickRefreshCustomFunctions();
          SpreadsheetApp.getActive().toast('Dữ liệu đã được cập nhật theo bộ lọc mới', '✅ Thành công', 2);
        } catch (error) {
          Logger.log('⚠️ Could not auto-refresh after filter change: ' + error.message);
        }
      }
      
      // Handle Quick Actions (Checkboxes)
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
 * Hàm cài đặt Trigger (Chạy 1 lần hoặc khi cập nhật Dashboard)
 * @param {boolean} silent - Nếu true, không hiển thị alert
 */
function createInstallableTriggers(silent = false) {
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
  
  if (!silent) {
    SpreadsheetApp.getUi().alert('✅ Đã cài đặt Trigger thành công! Các tính năng tự động và Quick Action sẽ hoạt động ngay.');
  }
  
  Logger.log('✅ Triggers installed successfully (onChange, processEdit)');
}
