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
 */
function onEdit(e) {
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
    // Tuy nhiên, nếu muốn chắc chắn, có thể gọi refreshDashboard()
    // Nhưng cẩn thận hiệu năng.
    
  } catch (error) {
    Logger.log('❌ Lỗi onEdit: ' + error.message);
  }
}
