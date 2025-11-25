/**
 * ===============================================
 * DEBUG UTILITY - KIỂM TRA CẤU TRÚC SHEET
 * ===============================================
 * 
 * Công cụ debug để kiểm tra dữ liệu trong các sheet
 * Cho vay và Quản lý nợ
 */

/**
 * Debug: Hiển thị thông tin chi tiết về các khoản nợ/cho vay
 * Giúp xác định vấn đề với định dạng dữ liệu
 */
function debugEventCalculation() {
  const ss = getSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  let debugInfo = '=== THÔNG TIN DEBUG - KHOẢN NỢ & CHO VAY ===\n\n';
  
  // 1. Check Debt Management Sheet
  debugInfo += '📋 QUẢN LÝ NỢ:\n';
  debugInfo += checkSheet(ss, APP_CONFIG.SHEETS.DEBT_MANAGEMENT, true);
  
  debugInfo += '\n📋 CHO VAY:\n';
  debugInfo += checkSheet(ss, APP_CONFIG.SHEETS.LENDING, false);
  
  // Show in dialog
  const htmlOutput = HtmlService.createHtmlOutput(
    '<pre style="font-family: monospace; font-size: 11px; white-space: pre-wrap; word-wrap: break-word;">' + 
    debugInfo + 
    '</pre>'
  )
  .setWidth(800)
  .setHeight(600);
  
  ui.showModalDialog(htmlOutput, '🔍 Debug - Event Calendar');
  
  // Also log to console
  Logger.log(debugInfo);
}

function checkSheet(ss, sheetName, isDebt) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return `❌ Sheet "${sheetName}" không tồn tại!\n`;
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return `⚠️ Sheet "${sheetName}" không có dữ liệu.\n`;
  
  const data = sheet.getRange(2, 1, Math.min(lastRow - 1, 5), 12).getValues();
  let info = '';
  
  data.forEach((row, idx) => {
    if (!row[1]) return; // Skip empty rows
    
    const name = row[1];
    const type = row[2];
    const principal = row[3];
    const rate = row[4];
    const term = row[5];
    const startDate = row[6];
    const maturityDate = row[7];
    const remaining = row[10];
    const status = row[11];
    
    info += `\n--- Khoản ${idx + 1}: ${name} ---\n`;
    info += `  Loại hình (raw): "${type}"\n`;
    info += `  Loại hình (mapped): "${mapLegacyTypeToId(type)}"\n`;
    info += `  Gốc ban đầu: ${principal} (type: ${typeof principal})\n`;
    info += `  Lãi suất (raw): ${rate} (type: ${typeof rate})\n`;
    
    // Check rate normalization
    let normalizedRate = parseFloat(rate) || 0;
    if (normalizedRate > 1) {
      info += `  ⚠️ Lãi suất > 1, sẽ chia cho 100: ${normalizedRate} -> ${normalizedRate/100}\n`;
      normalizedRate = normalizedRate / 100;
    }
    info += `  Lãi suất (normalized): ${normalizedRate}\n`;
    
    info += `  Kỳ hạn: ${term} tháng\n`;
    info += `  Ngày vay: ${startDate}\n`;
    info += `  Ngày đến hạn: ${maturityDate}\n`;
    info += `  Còn lại: ${remaining}\n`;
    info += `  Trạng thái: "${status}"\n`;
    
    // Check if active
    let isActive = false;
    if (isDebt) {
      isActive = (status === 'Chưa trả' || status === 'Đang trả');
    } else {
      isActive = (status === 'Đang vay');
    }
    info += `  Active: ${isActive ? '✅ YES' : '❌ NO'}\n`;
    
    // Try to calculate next event
    try {
      const today = new Date();
      today.setHours(0, 0, 0, 0);
      
      const parsedStartDate = parseDate(startDate);
      const parsedMaturityDate = parseDate(maturityDate);
      
      if (isActive && remaining > 0 && parsedStartDate) {
        const nextEvent = calculateNextPayment(mapLegacyTypeToId(type), {
          name,
          isDebt,
          initialPrincipal: parseCurrency(principal),
          remaining: parseCurrency(remaining),
          rate: normalizedRate,
          term: parseInt(term) || 1,
          startDate: parsedStartDate,
          maturityDate: parsedMaturityDate,
          today
        });
        
        if (nextEvent) {
          info += `  ✅ Event calculated:\n`;
          info += `     Ngày: ${nextEvent.date}\n`;
          info += `     Gốc trả: ${nextEvent.principalPayment}\n`;
          info += `     Lãi trả: ${nextEvent.interestPayment}\n`;
        } else {
          info += `  ❌ No event calculated (might be past due date)\n`;
        }
      } else {
        info += `  ⚠️ Cannot calculate: `;
        if (!isActive) info += `not active (status: ${status}), `;
        if (remaining <= 0) info += `remaining = 0, `;
        if (!parsedStartDate) info += `invalid start date, `;
        info += `\n`;
      }
    } catch (e) {
      info += `  ❌ ERROR calculating event: ${e.message}\n`;
    }
  });
  
  return info;
}

/**
 * Debug: Sửa lỗi định dạng cho cột "Nợ gốc ban đầu" / "Số tiền gốc"
 * Nếu cột này đang có format phần trăm thay vì số tiền
 */
function fixPrincipalColumnFormat() {
  const ss = getSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  const result = ui.alert(
    'Sửa định dạng cột Gốc',
    'Công cụ này sẽ:\n' +
    '1. Kiểm tra cột "Nợ gốc ban đầu" / "Số tiền gốc" (cột D)\n' +
    '2. Nếu đang có format phần trăm, sẽ chuyển về số tiền (x100)\n' +
    '3. Áp dụng format #,##0\n\n' +
    'Bạn có muốn tiếp tục?',
    ui.ButtonSet.YES_NO
  );
  
  if (result !== ui.Button.YES) return;
  
  let fixed = 0;
  
  // Fix Debt Management
  fixed += fixPrincipalColumn(ss, APP_CONFIG.SHEETS.DEBT_MANAGEMENT);
  
  // Fix Lending
  fixed += fixPrincipalColumn(ss, APP_CONFIG.SHEETS.LENDING);
  
  ui.alert(
    'Hoàn tất',
    `✅ Đã sửa ${fixed} ô trong các sheet Quản lý nợ và Cho vay.\n\n` +
    'Vui lòng chạy "Cập nhật Dashboard" để xem kết quả.',
    ui.ButtonSet.OK
  );
}

function fixPrincipalColumn(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return 0;
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  
  const range = sheet.getRange(2, 4, lastRow - 1, 1); // Column D
  const values = range.getValues();
  const formats = range.getNumberFormats();
  
  let fixedCount = 0;
  const newValues = values.map((row, idx) => {
    const val = row[0];
    const format = formats[idx][0];
    
    // Check if format contains '%'
    if (format && format.includes('%')) {
      // This value is displayed as percentage
      // E.g., if cell shows "80000000.00%", the actual value is 800000
      // We need to multiply by 100 to get the correct principal
      if (typeof val === 'number' && val > 0 && val < 1000000) {
        fixedCount++;
        return [val * 100];
      }
    }
    
    return [val];
  });
  
  if (fixedCount > 0) {
    range.setValues(newValues);
    range.setNumberFormat('#,##0');
    Logger.log(`Fixed ${fixedCount} cells in ${sheetName}`);
  }
  
  return fixedCount;
}
