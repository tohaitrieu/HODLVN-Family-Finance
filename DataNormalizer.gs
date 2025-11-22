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
    const toast = (msg) => ss.toast(msg, 'Chuẩn hóa dữ liệu', 3);
    toast('Đang xử lý...');
    
    // Configuration for normalization
    const sheetConfigs = [
      {
        name: APP_CONFIG.SHEETS.INCOME,
        dates: ['B2:B'],
        currency: ['C2:C'],
        numbers: []
      },
      {
        name: APP_CONFIG.SHEETS.EXPENSE,
        dates: ['B2:B'],
        currency: ['C2:C'],
        numbers: []
      },
      {
        name: APP_CONFIG.SHEETS.DEBT_PAYMENT,
        dates: ['B2:B'],
        currency: ['D2:F'],
        numbers: []
      },
      {
        name: APP_CONFIG.SHEETS.DEBT_MANAGEMENT,
        dates: ['G2:H'], // Ngày vay, Ngày đến hạn
        currency: ['D2:D', 'I2:K'], // Gốc, Đã trả, Còn lại
        numbers: ['E2:E'], // Lãi suất (decimal) handled separately? No, E is Rate, F is Term.
        // Wait, Debt Management: 
        // D: Principal (Currency)
        // E: Rate (Percent)
        // F: Term (Number)
        // G: Start Date
        // H: Maturity Date
        // I, J, K: Paid Prin, Paid Int, Remaining (Currency)
        custom: (sheet) => {
           sheet.getRange('E2:E').setNumberFormat('0.00"%"');
           sheet.getRange('F2:F').setNumberFormat('0'); // Term
        }
      },
      {
        name: APP_CONFIG.SHEETS.LENDING,
        dates: ['G2:H'],
        currency: ['D2:D', 'I2:K'],
        custom: (sheet) => {
           sheet.getRange('E2:E').setNumberFormat('0.00"%"');
           sheet.getRange('F2:F').setNumberFormat('0'); // Term
        }
      },
      {
        name: APP_CONFIG.SHEETS.LENDING_REPAYMENT,
        dates: ['B2:B'],
        currency: ['D2:F'],
        numbers: []
      },
      {
        name: APP_CONFIG.SHEETS.STOCK,
        dates: ['B2:B'],
        currency: ['F2:I', 'K2:N'], // Prices and Values
        numbers: ['E2:E'], // Quantity
        custom: (sheet) => {
           sheet.getRange('O2:O').setNumberFormat('0.00%');
        }
      },
      {
        name: APP_CONFIG.SHEETS.GOLD,
        dates: ['B2:B'],
        currency: ['H2:L'],
        numbers: ['F2:F'], // Quantity
        custom: (sheet) => {
           sheet.getRange('M2:M').setNumberFormat('0.00%');
        }
      },
      {
        name: APP_CONFIG.SHEETS.CRYPTO,
        dates: ['B2:B'],
        currency: ['H2:I', 'L2:N'], // VND Values
        numbers: ['E2:E'], // Quantity
        custom: (sheet) => {
           sheet.getRange('F2:F').setNumberFormat('#,##0.00'); // USD Price
           sheet.getRange('J2:K').setNumberFormat('#,##0.00'); // USD Values
           sheet.getRange('O2:O').setNumberFormat('0.00%');
        }
      },
      {
        name: APP_CONFIG.SHEETS.OTHER_INVESTMENT,
        dates: ['B2:B'],
        currency: ['D2:D', 'G2:G'],
        numbers: ['F2:F'], // Term
        custom: (sheet) => {
           sheet.getRange('E2:E').setNumberFormat('0.00"%"');
        }
      }
    ];
    
    let totalFixed = 0;
    
    sheetConfigs.forEach(config => {
      const sheet = ss.getSheetByName(config.name);
      if (!sheet) return;
      
      // 1. Fix Dates
      if (config.dates) {
        config.dates.forEach(rangeA1 => {
          const range = sheet.getRange(rangeA1);
          // Apply format
          range.setNumberFormat(APP_CONFIG.FORMATS.DATE);
          // Fix data types (String -> Date)
          _fixDateRange(range);
        });
      }
      
      // 2. Fix Currency
      if (config.currency) {
        config.currency.forEach(rangeA1 => {
          sheet.getRange(rangeA1).setNumberFormat(APP_CONFIG.FORMATS.NUMBER);
        });
      }
      
      // 3. Fix Numbers (General integer/float)
      if (config.numbers) {
        config.numbers.forEach(rangeA1 => {
          sheet.getRange(rangeA1).setNumberFormat('0'); // Or general number
        });
      }
      
      // 4. Custom Rules
      if (config.custom) {
        config.custom(sheet);
      }
      
      totalFixed++;
    });
    
    toast(`✅ Đã chuẩn hóa ${totalFixed} sheet!`);
    ui.alert('Thành công', `✅ Đã chuẩn hóa dữ liệu cho ${totalFixed} sheet!\n\n- Định dạng Ngày: dd/MM/yyyy\n- Định dạng Tiền: #,##0\n- Định dạng Số: Chuẩn hóa theo loại dữ liệu`, ui.ButtonSet.OK);
    
  } catch (error) {
    ui.alert('Lỗi', '❌ Có lỗi xảy ra: ' + error.message, ui.ButtonSet.OK);
    Logger.log(error);
  }
}

/**
 * Helper to fix date values in a range
 */
function _fixDateRange(range) {
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
