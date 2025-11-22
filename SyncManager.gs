/**
 * ===============================================
 * SYNC MANAGER
 * ===============================================
 * 
 * Xử lý đồng bộ dữ liệu giữa các sheet và tự động cập nhật STT
 */

var SyncManager = {
  
  // Cấu hình cột ID cho các sheet
  ID_COLUMNS: {
    'THU': 6,             // Col F
    'CHI': 7,             // Col G
    'QUẢN LÝ NỢ': 14,     // Col N
    'TRẢ NỢ': 8,          // Col H
    'CHO VAY': 14         // Col N
  },
  
  // Cấu hình cặp sheet cần đồng bộ
  SYNC_PAIRS: [
    { a: 'TRẢ NỢ', b: 'CHI' },
    { a: 'QUẢN LÝ NỢ', b: 'THU' },
    { a: 'CHO VAY', b: 'CHI' }
  ],

  /**
   * Hàm xử lý chính cho trigger onChange (Xóa/Thêm dòng)
   */
  handleOnChange: function(e) {
    try {
      if (!e) return;
      
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getActiveSheet();
      const sheetName = sheet.getName();
      
      // 1. Luôn cập nhật STT khi có thay đổi dòng
      if (e.changeType === 'REMOVE_ROW' || e.changeType === 'INSERT_ROW') {
        this.updateSheetSTT(sheet);
      }
      
      // 2. Xử lý đồng bộ xóa dữ liệu (Chỉ khi xóa dòng)
      if (e.changeType === 'REMOVE_ROW') {
        this.syncDeletedRows(ss, sheetName);
      }
      
    } catch (error) {
      Logger.log('❌ Lỗi handleOnChange: ' + error.message);
    }
  },
  
  /**
   * Hàm xử lý chính cho trigger onEdit (Sửa dữ liệu)
   */
  handleOnEdit: function(e) {
    try {
      if (!e) return;
      
      const range = e.range;
      const sheet = range.getSheet();
      const sheetName = sheet.getName();
      const row = range.getRow();
      const col = range.getColumn();
      
      if (row < 2) return; // Bỏ qua header
      
      // Kiểm tra xem sheet này có được theo dõi không
      const idCol = this.ID_COLUMNS[sheetName];
      if (!idCol) return;
      
      // Lấy TransactionID của dòng đang sửa
      const transactionId = sheet.getRange(row, idCol).getValue();
      if (!transactionId) return;
      
      // Tìm sheet đối tác để update
      const pair = this.SYNC_PAIRS.find(p => p.a === sheetName || p.b === sheetName);
      if (!pair) return;
      
      const targetSheetName = (pair.a === sheetName) ? pair.b : pair.a;
      const targetSheet = e.source.getSheetByName(targetSheetName);
      if (!targetSheet) return;
      
      // Tìm dòng bên kia có cùng ID
      const targetIdCol = this.ID_COLUMNS[targetSheetName];
      const targetRow = this.findRowById(targetSheet, targetIdCol, transactionId);
      
      if (targetRow !== -1) {
        this.syncRowData(sheet, row, targetSheet, targetRow, sheetName, targetSheetName);
      }
      
    } catch (error) {
      Logger.log('❌ Lỗi handleOnEdit: ' + error.message);
    }
  },
  
  /**
   * Đồng bộ dữ liệu giữa 2 dòng đã được link
   */
  syncRowData: function(sourceSheet, sourceRow, targetSheet, targetRow, sourceName, targetName) {
    // Mapping logic tùy thuộc vào cặp sheet
    
    // Case 1: TRẢ NỢ <-> CHI
    if ((sourceName === 'TRẢ NỢ' && targetName === 'CHI') || (sourceName === 'CHI' && targetName === 'TRẢ NỢ')) {
      // TRẢ NỢ: Col B (Date), Col F (Total)
      // CHI: Col B (Date), Col C (Amount)
      
      if (sourceName === 'TRẢ NỢ') {
        const date = sourceSheet.getRange(sourceRow, 2).getValue();
        const amount = sourceSheet.getRange(sourceRow, 6).getValue();
        
        targetSheet.getRange(targetRow, 2).setValue(date);
        targetSheet.getRange(targetRow, 3).setValue(amount);
      } else {
        const date = sourceSheet.getRange(sourceRow, 2).getValue();
        const amount = sourceSheet.getRange(sourceRow, 3).getValue();
        
        targetSheet.getRange(targetRow, 2).setValue(date);
        // Note: TRẢ NỢ amount is total (Col F). But it is composed of Principal + Interest.
        // Editing Total in Expense is ambiguous for Debt Payment breakdown.
        // Strategy: Update Total (Col F) in Debt Payment, but warn or leave breakdown as is?
        // Better: Just update Total for now, or maybe don't support reverse sync for complex fields.
        // Let's update Total (Col F) and assume it's all Principal update for simplicity or just update the display.
        // Actually, Debt Payment Col F is a formula usually? No, it's value in addDebtPayment.
        targetSheet.getRange(targetRow, 6).setValue(amount);
      }
    }
    
    // Case 2: QUẢN LÝ NỢ <-> THU
    else if ((sourceName === 'QUẢN LÝ NỢ' && targetName === 'THU') || (sourceName === 'THU' && targetName === 'QUẢN LÝ NỢ')) {
      // QUẢN LÝ NỢ: Col G (Date), Col D (Amount)
      // THU: Col B (Date), Col C (Amount)
      
      if (sourceName === 'QUẢN LÝ NỢ') {
        const date = sourceSheet.getRange(sourceRow, 7).getValue(); // Col G
        const amount = sourceSheet.getRange(sourceRow, 4).getValue(); // Col D
        
        targetSheet.getRange(targetRow, 2).setValue(date);
        targetSheet.getRange(targetRow, 3).setValue(amount);
      } else {
        const date = sourceSheet.getRange(sourceRow, 2).getValue();
        const amount = sourceSheet.getRange(sourceRow, 3).getValue();
        
        targetSheet.getRange(targetRow, 7).setValue(date);
        targetSheet.getRange(targetRow, 4).setValue(amount);
      }
    }
    
    // Case 3: CHO VAY <-> CHI
    else if ((sourceName === 'CHO VAY' && targetName === 'CHI') || (sourceName === 'CHI' && targetName === 'CHO VAY')) {
      // CHO VAY: Col G (Date), Col D (Amount)
      // CHI: Col B (Date), Col C (Amount)
      
      if (sourceName === 'CHO VAY') {
        const date = sourceSheet.getRange(sourceRow, 7).getValue(); // Col G
        const amount = sourceSheet.getRange(sourceRow, 4).getValue(); // Col D
        
        targetSheet.getRange(targetRow, 2).setValue(date);
        targetSheet.getRange(targetRow, 3).setValue(amount);
      } else {
        const date = sourceSheet.getRange(sourceRow, 2).getValue();
        const amount = sourceSheet.getRange(sourceRow, 3).getValue();
        
        targetSheet.getRange(targetRow, 7).setValue(date);
        targetSheet.getRange(targetRow, 4).setValue(amount);
      }
    }
    
    Logger.log(`🔄 Đã đồng bộ edit: ${sourceName} -> ${targetName}`);
  },
  
  /**
   * Tìm dòng theo ID
   */
  findRowById: function(sheet, idCol, id) {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return -1;
    
    const ids = sheet.getRange(2, idCol, lastRow - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (ids[i][0] === id) {
        return i + 2;
      }
    }
    return -1;
  },
  
  /**
   * Xử lý xóa dòng đồng bộ dựa trên ID
   */
  syncDeletedRows: function(ss, sheetName) {
    // Tìm sheet đối tác
    const pair = this.SYNC_PAIRS.find(p => p.a === sheetName || p.b === sheetName);
    if (!pair) return;
    
    const targetSheetName = (pair.a === sheetName) ? pair.b : pair.a;
    const targetSheet = ss.getSheetByName(targetSheetName);
    if (!targetSheet) return;
    
    const currentSheet = ss.getSheetByName(sheetName);
    const currentIdCol = this.ID_COLUMNS[sheetName];
    const targetIdCol = this.ID_COLUMNS[targetSheetName];
    
    // Lấy danh sách ID hiện tại của sheet vừa bị xóa dòng
    const currentIds = this.getAllIds(currentSheet, currentIdCol);
    const currentIdSet = new Set(currentIds);
    
    // Lấy danh sách ID của sheet đối tác
    const targetIds = this.getAllIds(targetSheet, targetIdCol);
    
    // Tìm những ID có bên Target mà KHÔNG có bên Current -> Cần xóa
    // (Giả sử ban đầu 2 bên khớp nhau, giờ bên Current mất 1 dòng -> bên Target dư ra dòng đó)
    // Lưu ý: Logic này chỉ đúng nếu trước đó đã đồng bộ. 
    // Để an toàn, ta chỉ xóa những dòng có ID (không rỗng) và không tìm thấy bên kia.
    
    const rowsToDelete = [];
    targetIds.forEach((id, index) => {
      if (id && !currentIdSet.has(id)) {
        rowsToDelete.push(index + 2); // Row index (1-based)
      }
    });
    
    // Xóa từ dưới lên
    for (let i = rowsToDelete.length - 1; i >= 0; i--) {
      targetSheet.deleteRow(rowsToDelete[i]);
      Logger.log(`🗑️ Đã xóa dòng đồng bộ bên ${targetSheetName} (Row ${rowsToDelete[i]})`);
    }
  },
  
  getAllIds: function(sheet, idCol) {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    return sheet.getRange(2, idCol, lastRow - 1, 1).getValues().map(r => r[0]);
  },
  
  updateSheetSTT: function(sheet) {
    try {
      const header = sheet.getRange(1, 1).getValue();
      if (header !== 'STT') return;
      
      const lastRow = sheet.getLastRow();
      if (lastRow < 2) return;
      
      const range = sheet.getRange(2, 1, lastRow - 1, 1);
      const values = range.getValues();
      let hasChange = false;
      
      const newValues = values.map((row, index) => {
        const newSTT = index + 1;
        if (row[0] !== newSTT) {
          hasChange = true;
        }
        return [newSTT];
      });
      
      if (hasChange) {
        range.setValues(newValues);
      }
    } catch (error) {
      Logger.log('Lỗi updateSheetSTT: ' + error.message);
    }
  }
};

