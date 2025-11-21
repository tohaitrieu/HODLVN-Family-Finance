/**
 * ===============================================
 * DATA NORMALIZER
 * ===============================================
 */

/**
 * Chuẩn hóa dữ liệu trên toàn bộ hệ thống
 */
function normalizeAllData() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    ui.alert('Đang xử lý...', 'Hệ thống đang chuẩn hóa dữ liệu. Vui lòng đợi...', ui.ButtonSet.OK);
    
    // 1. Define map of Sheet -> Date Columns (1-indexed)
    const dateColumnsMap = {
      [APP_CONFIG.SHEETS.INCOME]: [2],
      [APP_CONFIG.SHEETS.EXPENSE]: [2],
      [APP_CONFIG.SHEETS.DEBT_PAYMENT]: [2],
      [APP_CONFIG.SHEETS.DEBT_MANAGEMENT]: [6, 7], // Ngày vay, Ngày đến hạn
      [APP_CONFIG.SHEETS.LENDING]: [6, 7],         // Ngày vay, Ngày đến hạn
      [APP_CONFIG.SHEETS.STOCK]: [2],
      [APP_CONFIG.SHEETS.GOLD]: [2],
      [APP_CONFIG.SHEETS.CRYPTO]: [2],
      [APP_CONFIG.SHEETS.OTHER_INVESTMENT]: [2]
    };
    
    let totalFixed = 0;
    
    // 2. Iterate and fix
    for (const [sheetName, cols] of Object.entries(dateColumnsMap)) {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) continue;
      
      cols.forEach(col => {
        // Use SheetInitializer's helper if available, or local implementation
        if (SheetInitializer && SheetInitializer._fixDateColumn) {
          SheetInitializer._fixDateColumn(sheet, col);
        } else {
          // Fallback implementation
          _fixDateColumnLocal(sheet, col);
        }
      });
      
      totalFixed++;
    }
    
    ui.alert('Thành công', `✅ Đã chuẩn hóa dữ liệu cho ${totalFixed} sheet!`, ui.ButtonSet.OK);
    
  } catch (error) {
    ui.alert('Lỗi', '❌ Có lỗi xảy ra: ' + error.message, ui.ButtonSet.OK);
    Logger.log(error);
  }
}

/**
 * Local helper to fix date column (fallback)
 */
function _fixDateColumnLocal(sheet, colIndex) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  
  const range = sheet.getRange(2, colIndex, lastRow - 1, 1);
  const values = range.getValues();
  let hasChange = false;
  
  const fixedValues = values.map(row => {
    const val = row[0];
    if (!val) return [val];

    let dateObj = null;

    // Case 1: String "dd/mm/yyyy"
    if (typeof val === 'string' && val.match(/^\d{1,2}\/\d{1,2}\/\d{4}/)) {
      const parts = val.split('/');
      dateObj = new Date(parts[2], parts[1] - 1, parts[0]);
    } 
    // Case 2: Already a Date object
    else if (val instanceof Date) {
      dateObj = val;
    }

    if (dateObj) {
      const cleanDate = new Date(dateObj.getFullYear(), dateObj.getMonth(), dateObj.getDate());
      if (val instanceof Date && val.getTime() === cleanDate.getTime()) {
        return [val];
      }
      hasChange = true;
      return [cleanDate];
    }
    
    return [val];
  });
  
  if (hasChange) {
    range.setValues(fixedValues);
  }
  range.setNumberFormat('dd/mm/yyyy');
}

/**
 * Tính toán lịch trả nợ tiếp theo
 * @returns {Array} List of upcoming payments
 */
function calculateNextDebtPayments() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const debtSheet = ss.getSheetByName(APP_CONFIG.SHEETS.DEBT_MANAGEMENT);
  
  if (!debtSheet) return [];
  
  const data = debtSheet.getDataRange().getValues();
  // Header: STT, Tên, Gốc, Lãi suất, Kỳ hạn, Ngày vay, Đáo hạn, Đã trả gốc, Đã trả lãi, Còn nợ, Trạng thái
  // Indices: 0    1    2    3        4       5         6        7           8           9       10
  
  const payments = [];
  const today = new Date();
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const status = row[10];
    
    if (status === 'Đã thanh toán') continue;
    
    const name = row[1];
    const principal = parseFloat(row[2]);
    const rate = parseFloat(row[3]); // Decimal (e.g. 0.12 for 12%)
    const term = parseInt(row[4]);
    const startDate = new Date(row[5]);
    const remaining = parseFloat(row[9]);
    
    if (remaining <= 0) continue;
    
    // Calculate monthly payment
    // Reducing balance method:
    // Monthly Principal = Total Principal / Term
    // Monthly Interest = Remaining Principal * (Rate / 12)
    
    const monthlyPrincipal = principal / term;
    const monthlyInterest = remaining * (rate / 12);
    const totalMonthly = monthlyPrincipal + monthlyInterest;
    
    // Estimate next payment date (simple logic: same day of next month)
    // In reality, we should track "Last Payment Date".
    // For now, assume payment is due on the same day of the month as Start Date.
    
    let nextPaymentDate = new Date(today.getFullYear(), today.getMonth(), startDate.getDate());
    if (nextPaymentDate < today) {
      nextPaymentDate.setMonth(nextPaymentDate.getMonth() + 1);
    }
    
    payments.push({
      name: name,
      dueDate: nextPaymentDate,
      principal: monthlyPrincipal,
      interest: monthlyInterest,
      total: totalMonthly,
      remaining: remaining
    });
  }
  
  return payments;
}

/**
 * Hiển thị báo cáo lịch trả nợ
 */
function showDebtScheduleReport() {
  const payments = calculateNextDebtPayments();
  
  if (payments.length === 0) {
    SpreadsheetApp.getUi().alert('Thông báo', 'Không có khoản nợ nào cần trả trong thời gian tới!', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  // Create HTML content
  let html = `
    <html>
    <head>
      <style>
        body { font-family: Arial, sans-serif; padding: 10px; }
        table { width: 100%; border-collapse: collapse; margin-top: 10px; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: right; }
        th { background-color: #4472C4; color: white; text-align: center; }
        td:first-child { text-align: left; }
        .total-row { font-weight: bold; background-color: #f0f0f0; }
      </style>
    </head>
    <body>
      <h3>📅 Lịch Trả Nợ Dự Kiến (Tháng tới)</h3>
      <table>
        <tr>
          <th>Khoản nợ</th>
          <th>Ngày trả</th>
          <th>Gốc</th>
          <th>Lãi</th>
          <th>Tổng cộng</th>
        </tr>
  `;
  
  let totalP = 0;
  let totalI = 0;
  let totalT = 0;
  
  payments.forEach(p => {
    html += `
      <tr>
        <td>${p.name}</td>
        <td style="text-align: center">${Utilities.formatDate(p.dueDate, Session.getScriptTimeZone(), 'dd/MM/yyyy')}</td>
        <td>${formatCurrency(p.principal)}</td>
        <td>${formatCurrency(p.interest)}</td>
        <td>${formatCurrency(p.total)}</td>
      </tr>
    `;
    totalP += p.principal;
    totalI += p.interest;
    totalT += p.total;
  });
  
  html += `
      <tr class="total-row">
        <td colspan="2" style="text-align: center">TỔNG CỘNG</td>
        <td>${formatCurrency(totalP)}</td>
        <td>${formatCurrency(totalI)}</td>
        <td>${formatCurrency(totalT)}</td>
      </tr>
    </table>
    <p><i>* Lưu ý: Tính toán dựa trên dư nợ hiện tại và lãi suất giảm dần.</i></p>
    </body>
    </html>
  `;
  
  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(600)
    .setHeight(400);
    
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Lịch Trả Nợ');
}
