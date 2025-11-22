/**
 * ===============================================
 * MIGRATION.GS
 * ===============================================
 * 
 * Script di chuyển dữ liệu và cập nhật cấu trúc mới (Transaction ID)
 */

var MIGRATION_CONFIG = {
  // Định nghĩa cột Transaction ID cho từng sheet (1-based index)
  // Lưu ý: Cột này nên là cột ẩn hoặc cột cuối cùng
  ID_COLUMNS: {
    'THU': 6,             // Col F
    'CHI': 7,             // Col G
    'QUẢN LÝ NỢ': 14,     // Col N
    'TRẢ NỢ': 8,          // Col H
    'CHO VAY': 14         // Col N
  }
};

/**
 * Chạy migration để thêm Transaction ID cho dữ liệu cũ
 */
function runMigration_AddTransactionIDs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  try {
    // 1. Tạo ID cho từng sheet độc lập trước
    const sheets = Object.keys(MIGRATION_CONFIG.ID_COLUMNS);
    
    sheets.forEach(sheetName => {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) return;
      
      const idCol = MIGRATION_CONFIG.ID_COLUMNS[sheetName];
      const lastRow = sheet.getLastRow();
      
      if (lastRow >= 2) {
        // Đọc cột ID hiện tại
        const idRange = sheet.getRange(2, idCol, lastRow - 1, 1);
        const idValues = idRange.getValues();
        let hasChange = false;
        
        const newValues = idValues.map(row => {
          if (!row[0]) {
            hasChange = true;
            return [Utilities.getUuid()]; // Tạo ID mới nếu chưa có
          }
          return [row[0]];
        });
        
        if (hasChange) {
          idRange.setValues(newValues);
          Logger.log(`✅ Đã tạo ID cho sheet ${sheetName}`);
        }
      }
      
      // Đặt header cho cột ID
      sheet.getRange(1, idCol).setValue('TransactionID').setFontColor('#cccccc');
      // Ẩn cột ID (nếu muốn)
      // sheet.hideColumns(idCol);
    });
    
    // 2. Link dữ liệu (Đồng bộ ID giữa các cặp sheet liên kết)
    // Logic: Tìm các cặp khớp nhau (Date + Amount + Name) và gán cùng 1 ID
    
    // Pair 1: TRẢ NỢ <-> CHI
    syncExistingIds(ss, 'TRẢ NỢ', 'CHI', matchDebtPaymentAndExpense);
    
    // Pair 2: QUẢN LÝ NỢ <-> THU
    syncExistingIds(ss, 'QUẢN LÝ NỢ', 'THU', matchDebtAndIncome);
    
    // Pair 3: CHO VAY <-> CHI
    syncExistingIds(ss, 'CHO VAY', 'CHI', matchLendingAndExpense);
    
    ui.alert('✅ Migration hoàn tất! Đã thêm Transaction ID và liên kết dữ liệu.');
    
  } catch (error) {
    Logger.log('❌ Lỗi Migration: ' + error.message);
    ui.alert('❌ Lỗi Migration: ' + error.message);
  }
}

/**
 * Hàm đồng bộ ID giữa 2 sheet dựa trên hàm match
 */
function syncExistingIds(ss, sheetNameA, sheetNameB, matchFunc) {
  const sheetA = ss.getSheetByName(sheetNameA);
  const sheetB = ss.getSheetByName(sheetNameB);
  
  if (!sheetA || !sheetB) return;
  
  const idColA = MIGRATION_CONFIG.ID_COLUMNS[sheetNameA];
  const idColB = MIGRATION_CONFIG.ID_COLUMNS[sheetNameB];
  
  const dataA = getDataWithId(sheetA, idColA);
  const dataB = getDataWithId(sheetB, idColB);
  
  let updatesB = []; // Danh sách update cho sheet B
  
  // Duyệt qua sheet A, tìm dòng khớp bên sheet B
  dataA.forEach(rowA => {
    // Tìm dòng khớp trong B mà chưa được sync (hoặc ID khác)
    const matchIndex = dataB.findIndex(rowB => matchFunc(rowA, rowB));
    
    if (matchIndex !== -1) {
      const rowB = dataB[matchIndex];
      
      // Nếu ID khác nhau, gán ID của A cho B (để A làm chuẩn)
      if (rowA.id !== rowB.id) {
        // Cập nhật trong mảng dataB để không match lại dòng này
        rowB.id = rowA.id; 
        
        // Lưu lại để update batch sau
        updatesB.push({
          rowIndex: rowB.rowIndex,
          newId: rowA.id
        });
        
        Logger.log(`🔗 Linked: ${sheetNameA} (${rowA.rowIndex}) <-> ${sheetNameB} (${rowB.rowIndex})`);
      }
    }
  });
  
  // Thực hiện update cho sheet B
  updatesB.forEach(update => {
    sheetB.getRange(update.rowIndex, idColB).setValue(update.newId);
  });
}

/**
 * Helper: Lấy dữ liệu cùng với ID
 */
function getDataWithId(sheet, idCol) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  
  const lastCol = sheet.getLastColumn();
  // Lấy toàn bộ dữ liệu
  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  // Lấy cột ID riêng (vì idCol có thể nằm ngoài lastCol nếu chưa có dữ liệu)
  // Nhưng ở bước 1 ta đã fill ID rồi, nên idCol chắc chắn <= lastCol hoặc ta lấy riêng
  const idValues = sheet.getRange(2, idCol, lastRow - 1, 1).getValues();
  
  return values.map((row, index) => {
    return {
      rowIndex: index + 2,
      data: row,
      id: idValues[index][0]
    };
  });
}

// ================= MATCHING FUNCTIONS =================

function matchDebtPaymentAndExpense(rowA, rowB) {
  // A: TRẢ NỢ (Col B: Date, Col C: Name, Col F: Total)
  // B: CHI (Col B: Date, Col C: Amount, Col D: Category, Col E: Subcategory)
  
  const dateA = formatDate(rowA.data[1]);
  const dateB = formatDate(rowB.data[1]);
  const amountA = Math.round(rowA.data[5]); // Total Amount
  const amountB = Math.round(rowB.data[2]); // Amount
  
  if (dateA !== dateB) return false;
  if (amountA !== amountB) return false;
  
  // Check Name
  const nameA = rowA.data[2].toString().toLowerCase();
  const subB = rowB.data[4].toString().toLowerCase(); // Subcategory: "Trả nợ: [Name]"
  
  return subB.includes(nameA);
}

function matchDebtAndIncome(rowA, rowB) {
  // A: QUẢN LÝ NỢ (Col B: Name, Col D: Amount, Col G: Date)
  // B: THU (Col B: Date, Col C: Amount, Col D: Source, Col E: Note)
  
  const dateA = formatDate(rowA.data[6]); // Col G
  const dateB = formatDate(rowB.data[1]);
  const amountA = Math.round(rowA.data[3]); // Col D
  const amountB = Math.round(rowB.data[2]);
  
  if (dateA !== dateB) return false;
  if (amountA !== amountB) return false;
  
  const nameA = rowA.data[1].toString().toLowerCase();
  const noteB = rowB.data[4].toString().toLowerCase(); // Note: "Vay: [Name]"
  
  return noteB.includes(nameA);
}

function matchLendingAndExpense(rowA, rowB) {
  // A: CHO VAY (Col B: Name, Col D: Amount, Col G: Date)
  // B: CHI (Col B: Date, Col C: Amount, Col E: Subcategory)
  
  const dateA = formatDate(rowA.data[6]); // Col G
  const dateB = formatDate(rowB.data[1]);
  const amountA = Math.round(rowA.data[3]); // Col D
  const amountB = Math.round(rowB.data[2]);
  
  if (dateA !== dateB) return false;
  if (amountA !== amountB) return false;
  
  const nameA = rowA.data[1].toString().toLowerCase();
  const subB = rowB.data[4].toString().toLowerCase(); // Subcategory: "Cho vay: [Name]"
  
  return subB.includes(nameA);
}

function formatDate(date) {
  if (!date) return '';
  const d = new Date(date);
  return `${d.getFullYear()}-${d.getMonth()}-${d.getDate()}`;
}
