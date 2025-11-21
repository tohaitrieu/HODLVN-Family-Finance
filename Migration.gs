/**
 * ===============================================
 * MIGRATION.GS - CHUYỂN ĐỔI CẤU TRÚC DỮ LIỆU
 * ===============================================
 */

function runMigrations() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
    'Cập nhật cấu trúc dữ liệu',
    'Bạn có muốn cập nhật cấu trúc Sheet VÀNG và CRYPTO không?\nQuá trình này sẽ thêm cột mới và giữ nguyên dữ liệu cũ.',
    ui.ButtonSet.YES_NO
  );

  if (result == ui.Button.YES) {
    try {
      migrateGoldSheet();
      migrateCryptoSheet();
      ui.alert('✅ Đã cập nhật thành công!');
    } catch (e) {
      ui.alert('❌ Lỗi: ' + e.message);
    }
  }
}

function migrateGoldSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(APP_CONFIG.SHEETS.GOLD);
  if (!sheet) return;

  // Check if already migrated (Check Col 3 Header)
  const h3 = sheet.getRange('C1').getValue();
  if (h3 === 'Tài sản') {
    Logger.log('Sheet VÀNG đã ở cấu trúc mới. Cập nhật lại header để đảm bảo.');
    // Force update headers even if migrated
    const headers = [
      'STT', 'Ngày', 'Tài sản', 'Loại GD', 'Loại vàng', 'Số lượng', 'Đơn vị', 
      'Giá vốn', 'Tổng vốn', 'Giá HT', 'Giá trị HT', 'Lãi/Lỗ', '% Lãi/Lỗ', 'Ghi chú'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    return;
  }

  Logger.log('Đang migrate Sheet VÀNG...');

  // 1. Insert Col 3 (Tài sản)
  sheet.insertColumnAfter(2);
  sheet.getRange('C1').setValue('Tài sản');
  
  // Fill "GOLD" for all existing rows
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    sheet.getRange(2, 3, lastRow - 1, 1).setValue('GOLD');
  }

  // 2. Delete old "Nơi lưu" (Was Col 9 in old, now Col 10 because of insertion)
  // Old: 1:STT, 2:Date, 3:Type, 4:GoldType, 5:Qty, 6:Unit, 7:Price, 8:Total, 9:Location, 10:Note
  // After Insert Col 3: 1:STT, 2:Date, 3:GOLD, 4:Type, 5:GoldType, 6:Qty, 7:Unit, 8:Price, 9:Total, 10:Location, 11:Note
  sheet.deleteColumn(10);

  // Now: 1:STT, 2:Date, 3:GOLD, 4:Type, 5:GoldType, 6:Qty, 7:Unit, 8:Price, 9:Total, 10:Note
  
  // 3. Insert 4 columns before Note (at Col 10)
  // We want: 10:PriceHT, 11:ValueHT, 12:PL, 13:%, 14:Note
  sheet.insertColumnsBefore(10, 4);

  // 4. Update Headers
  const headers = [
    'STT', 'Ngày', 'Tài sản', 'Loại GD', 'Loại vàng', 'Số lượng', 'Đơn vị', 
    'Giá vốn', 'Tổng vốn', 'Giá HT', 'Giá trị HT', 'Lãi/Lỗ', '% Lãi/Lỗ', 'Ghi chú'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // 5. Update Formats & Formulas
  if (lastRow >= 2) {
    // Formulas
    // J: Giá HT = GPRICE(Loại vàng)
    sheet.getRange(2, 10, lastRow - 1).setFormulaR1C1('=IF(R[0]C[-5]<>"", GPRICE(R[0]C[-5]), 0)');
    
    // K: Giá trị HT = Số lượng * Giá HT
    sheet.getRange(2, 11, lastRow - 1).setFormulaR1C1('=IF(AND(R[0]C[-5]>0, R[0]C[-1]>0), R[0]C[-5]*R[0]C[-1], 0)');
    
    // L: Lãi/Lỗ = Giá trị HT - Tổng vốn
    sheet.getRange(2, 12, lastRow - 1).setFormulaR1C1('=IF(R[0]C[-1]>0, R[0]C[-1]-R[0]C[-3], 0)');
    
    // M: % Lãi/Lỗ
    sheet.getRange(2, 13, lastRow - 1).setFormulaR1C1('=IF(R[0]C[-4]>0, R[0]C[-1]/R[0]C[-4], 0)');
    
    // Formats
    sheet.getRange(2, 8, lastRow - 1, 5).setNumberFormat('#,##0'); // H-L
    sheet.getRange(2, 13, lastRow - 1, 1).setNumberFormat('0.00%'); // M
  }
  
  // Re-apply styling
  SheetInitializer.initializeGoldSheet();
}

function migrateCryptoSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(APP_CONFIG.SHEETS.CRYPTO);
  if (!sheet) return;

  // Check if already migrated
  const h10 = sheet.getRange('J1').getValue();
  if (h10 === 'Giá HT (USD)') {
    Logger.log('Sheet CRYPTO đã ở cấu trúc mới. Cập nhật lại header để đảm bảo.');
    const headers = [
      'STT', 'Ngày', 'Loại GD', 'Coin', 'Số lượng', 'Giá (USD)', 'Tỷ giá', 'Giá (VND)', 'Tổng vốn',
      'Giá HT (USD)', 'Giá trị HT (USD)', 'Giá HT (VND)', 'Giá trị HT (VND)', 'Lãi/Lỗ', '% Lãi/Lỗ',
      'Sàn', 'Ví', 'Ghi chú'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    return;
  }

  Logger.log('Đang migrate Sheet CRYPTO...');

  // Old: 1:STT, 2:Date, 3:Type, 4:Coin, 5:Qty, 6:Price, 7:Rate, 8:PriceVND, 9:Total, 10:Exchange, 11:Wallet, 12:Note
  // New: ... 9:Total, 10:PriceHT_USD, 11:ValueHT_USD, 12:PriceHT_VND, 13:ValueHT_VND, 14:PL, 15:%, 16:Exchange...

  // 1. Insert 6 columns at Col 10
  sheet.insertColumnsBefore(10, 6);

  // 2. Update Headers
  const headers = [
    'STT', 'Ngày', 'Loại GD', 'Coin', 'Số lượng', 'Giá (USD)', 'Tỷ giá', 'Giá (VND)', 'Tổng vốn',
    'Giá HT (USD)', 'Giá trị HT (USD)', 'Giá HT (VND)', 'Giá trị HT (VND)', 'Lãi/Lỗ', '% Lãi/Lỗ',
    'Sàn', 'Ví', 'Ghi chú'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // 3. Update Formulas
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    // J: Giá HT (USD)
    sheet.getRange(2, 10, lastRow - 1).setFormulaR1C1('=IF(R[0]C[-6]<>"", CPRICE(R[0]C[-6]&"USD"), 0)');
    
    // K: Giá trị HT (USD)
    sheet.getRange(2, 11, lastRow - 1).setFormulaR1C1('=IF(AND(R[0]C[-6]>0, R[0]C[-1]>0), R[0]C[-6]*R[0]C[-1], 0)');
    
    // L: Giá HT (VND)
    sheet.getRange(2, 12, lastRow - 1).setFormulaR1C1('=IF(AND(R[0]C[-2]>0, R[0]C[-5]>0), R[0]C[-2]*R[0]C[-5], 0)');
    
    // M: Giá trị HT (VND)
    sheet.getRange(2, 13, lastRow - 1).setFormulaR1C1('=IF(AND(R[0]C[-2]>0, R[0]C[-6]>0), R[0]C[-2]*R[0]C[-6], 0)');
    
    // N: Lãi/Lỗ
    sheet.getRange(2, 14, lastRow - 1).setFormulaR1C1('=IF(R[0]C[-1]>0, R[0]C[-1]-R[0]C[-5], 0)');
    
    // O: % Lãi/Lỗ
    sheet.getRange(2, 15, lastRow - 1).setFormulaR1C1('=IF(R[0]C[-6]>0, R[0]C[-1]/R[0]C[-6], 0)');
    
    // Formats
    sheet.getRange(2, 10, lastRow - 1, 2).setNumberFormat('#,##0.00'); // USD
    sheet.getRange(2, 12, lastRow - 1, 3).setNumberFormat('#,##0');    // VND
    sheet.getRange(2, 15, lastRow - 1, 1).setNumberFormat('0.00%');    // %
  }

  // Re-apply styling
  SheetInitializer.initializeCryptoSheet();
}
