/**
 * ===============================================
 * DATAPROCESSORS.GS - XỬ LÝ GIAO DỊCH v3.0
 * ===============================================
 * 
 * Chức năng: Xử lý tất cả các loại giao dịch
 * - Thu nhập
 * - Chi tiêu
 * - Trả nợ
 * - Đầu tư (CK, Vàng, Crypto, Khác)
 */

// ==================== THU NHẬP ====================

/**
 * Thêm khoản thu nhập
 * @param {Object} data - {date, amount, source, note}
 * @return {Object} {success, message}
 */
function addIncome(data) {
  try {
    // Validation
    if (!data.date || !data.amount || !data.source) {
      return {
        success: false,
        message: '❌ Vui lòng điền đầy đủ thông tin bắt buộc!'
      };
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(APP_CONFIG.SHEETS.INCOME);
    
    if (!sheet) {
      return {
        success: false,
        message: '❌ Sheet THU chưa được khởi tạo! Vui lòng khởi tạo sheet trước.'
      };
    }
    
    // Parse dữ liệu
    const date = new Date(data.date);
    const amount = parseFloat(data.amount);
    const source = data.source;
    const note = data.note || '';
    
    // Lấy STT tự động
    const lastRow = sheet.getLastRow();
    const stt = lastRow > 1 ? lastRow : 1;
    
    // Thêm dữ liệu
    sheet.appendRow([
      stt,
      date,
      amount,
      source,
      note
    ]);
    
    // Format dòng mới
    const newRow = sheet.getLastRow();
    sheet.getRange(newRow, 1, 1, 5).setHorizontalAlignment('left');
    sheet.getRange(newRow, 2).setNumberFormat('dd/mm/yyyy');
    sheet.getRange(newRow, 3).setNumberFormat('#,##0" VNĐ"');
    
    return {
      success: true,
      message: `✅ Đã thêm thu nhập ${formatCurrency(amount)} từ ${source}!`
    };
    
  } catch (error) {
    return {
      success: false,
      message: `❌ Lỗi: ${error.message}`
    };
  }
}

// ==================== CHI TIÊU ====================

/**
 * Thêm khoản chi tiêu
 * @param {Object} data - {date, amount, category, detail, note}
 * @return {Object} {success, message}
 */
function addExpense(data) {
  try {
    // Validation
    if (!data.date || !data.amount || !data.category) {
      return {
        success: false,
        message: '❌ Vui lòng điền đầy đủ thông tin bắt buộc!'
      };
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(APP_CONFIG.SHEETS.EXPENSE);
    
    if (!sheet) {
      return {
        success: false,
        message: '❌ Sheet CHI chưa được khởi tạo! Vui lòng khởi tạo sheet trước.'
      };
    }
    
    // Parse dữ liệu
    const date = new Date(data.date);
    const amount = parseFloat(data.amount);
    const category = data.category;
    const detail = data.detail || '';
    const note = data.note || '';
    
    // Lấy STT tự động
    const lastRow = sheet.getLastRow();
    const stt = lastRow > 1 ? lastRow : 1;
    
    // Thêm dữ liệu
    sheet.appendRow([
      stt,
      date,
      amount,
      category,
      detail,
      note
    ]);
    
    // Format dòng mới
    const newRow = sheet.getLastRow();
    sheet.getRange(newRow, 1, 1, 6).setHorizontalAlignment('left');
    sheet.getRange(newRow, 2).setNumberFormat('dd/mm/yyyy');
    sheet.getRange(newRow, 3).setNumberFormat('#,##0" VNĐ"');
    
    // Cập nhật Budget
    try {
      updateBudgetSpent(category);
    } catch (e) {
      Logger.log('Không thể cập nhật budget: ' + e.message);
    }
    
    return {
      success: true,
      message: `✅ Đã thêm chi tiêu ${formatCurrency(amount)} cho ${category}!`
    };
    
  } catch (error) {
    return {
      success: false,
      message: `❌ Lỗi: ${error.message}`
    };
  }
}

// ==================== TRẢ NỢ ====================

/**
 * Thêm khoản trả nợ
 * @param {Object} data - {date, debtName, principal, interest, note}
 * @return {Object} {success, message}
 */
function addDebtPayment(data) {
  try {
    // Validation
    if (!data.date || !data.debtName || (!data.principal && !data.interest)) {
      return {
        success: false,
        message: '❌ Vui lòng điền đầy đủ thông tin bắt buộc!'
      };
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(APP_CONFIG.SHEETS.DEBT_PAYMENT);
    
    if (!sheet) {
      return {
        success: false,
        message: '❌ Sheet TRẢ NỢ chưa được khởi tạo!'
      };
    }
    
    // Parse dữ liệu
    const date = new Date(data.date);
    const debtName = data.debtName;
    const principal = parseFloat(data.principal) || 0;
    const interest = parseFloat(data.interest) || 0;
    const total = principal + interest;
    const note = data.note || '';
    
    // Lấy STT
    const lastRow = sheet.getLastRow();
    const stt = lastRow > 1 ? lastRow : 1;
    
    // Thêm dữ liệu
    sheet.appendRow([
      stt,
      date,
      debtName,
      principal,
      interest,
      total,
      note
    ]);
    
    // Format
    const newRow = sheet.getLastRow();
    sheet.getRange(newRow, 2).setNumberFormat('dd/mm/yyyy');
    sheet.getRange(newRow, 4, 1, 3).setNumberFormat('#,##0" VNĐ"');
    
    return {
      success: true,
      message: `✅ Đã ghi nhận trả nợ ${debtName}: ${formatCurrency(total)}!`
    };
    
  } catch (error) {
    return {
      success: false,
      message: `❌ Lỗi: ${error.message}`
    };
  }
}

// ==================== CHỨNG KHOÁN ====================

/**
 * Thêm giao dịch chứng khoán
 * @param {Object} data - {date, type, symbol, quantity, price, fee, note}
 * @return {Object} {success, message}
 */
function addStock(data) {
  try {
    if (!data.date || !data.type || !data.symbol || !data.quantity || !data.price) {
      return {
        success: false,
        message: '❌ Vui lòng điền đầy đủ thông tin!'
      };
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(APP_CONFIG.SHEETS.STOCK);
    
    if (!sheet) {
      return {
        success: false,
        message: '❌ Sheet CHỨNG KHOÁN chưa được khởi tạo!'
      };
    }
    
    const date = new Date(data.date);
    const type = data.type;
    const symbol = data.symbol.toUpperCase();
    const quantity = parseFloat(data.quantity);
    const price = parseFloat(data.price);
    const fee = parseFloat(data.fee) || 0;
    const total = (quantity * price) + fee;
    const note = data.note || '';
    
    const lastRow = sheet.getLastRow();
    const stt = lastRow > 1 ? lastRow : 1;
    
    sheet.appendRow([
      stt,
      date,
      type,
      symbol,
      quantity,
      price,
      fee,
      total,
      note
    ]);
    
    // Format
    const newRow = sheet.getLastRow();
    sheet.getRange(newRow, 2).setNumberFormat('dd/mm/yyyy');
    sheet.getRange(newRow, 6, 1, 3).setNumberFormat('#,##0" VNĐ"');
    
    // Nếu mua thì trừ Budget Đầu tư
    if (type === 'Mua') {
      try {
        updateInvestmentBudget('Chứng khoán', total);
      } catch (e) {
        Logger.log('Không thể cập nhật budget: ' + e.message);
      }
    }
    
    // Nếu bán thì thêm vào THU
    if (type === 'Bán') {
      try {
        addIncome({
          date: data.date,
          amount: total,
          source: 'Bán chứng khoán',
          note: `Bán ${quantity} ${symbol} @ ${price}`
        });
      } catch (e) {
        Logger.log('Không thể thêm thu nhập: ' + e.message);
      }
    }
    
    return {
      success: true,
      message: `✅ Đã ghi nhận ${type} ${quantity} ${symbol} với giá ${formatCurrency(price)}!`
    };
    
  } catch (error) {
    return {
      success: false,
      message: `❌ Lỗi: ${error.message}`
    };
  }
}

// ==================== VÀNG ====================

/**
 * Thêm giao dịch vàng
 */
function addGold(data) {
  try {
    if (!data.date || !data.type || !data.goldType || !data.quantity || !data.price) {
      return {
        success: false,
        message: '❌ Vui lòng điền đầy đủ thông tin!'
      };
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(APP_CONFIG.SHEETS.GOLD);
    
    if (!sheet) {
      return { success: false, message: '❌ Sheet VÀNG chưa được khởi tạo!' };
    }
    
    const date = new Date(data.date);
    const type = data.type;
    const goldType = data.goldType;
    const quantity = parseFloat(data.quantity);
    const unit = data.unit || 'chỉ';
    const price = parseFloat(data.price);
    const location = data.location || '';
    const total = quantity * price;
    const note = data.note || '';
    
    const lastRow = sheet.getLastRow();
    const stt = lastRow > 1 ? lastRow : 1;
    
    sheet.appendRow([stt, date, type, goldType, quantity, unit, price, total, location, note]);
    
    const newRow = sheet.getLastRow();
    sheet.getRange(newRow, 2).setNumberFormat('dd/mm/yyyy');
    sheet.getRange(newRow, 7, 1, 2).setNumberFormat('#,##0" VNĐ"');
    
    if (type === 'Mua') {
      try { updateInvestmentBudget('Vàng', total); } catch(e) {}
    }
    
    if (type === 'Bán') {
      try {
        addIncome({
          date: data.date,
          amount: total,
          source: 'Bán vàng',
          note: `Bán ${quantity} ${unit} ${goldType}`
        });
      } catch(e) {}
    }
    
    return {
      success: true,
      message: `✅ Đã ghi nhận ${type} ${quantity} ${unit} ${goldType}!`
    };
    
  } catch (error) {
    return { success: false, message: `❌ Lỗi: ${error.message}` };
  }
}

// ==================== CRYPTO ====================

/**
 * Thêm giao dịch crypto
 */
function addCrypto(data) {
  try {
    if (!data.date || !data.type || !data.coin || !data.quantity || !data.priceUSD) {
      return {
        success: false,
        message: '❌ Vui lòng điền đầy đủ thông tin!'
      };
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(APP_CONFIG.SHEETS.CRYPTO);
    
    if (!sheet) {
      return { success: false, message: '❌ Sheet CRYPTO chưa được khởi tạo!' };
    }
    
    const date = new Date(data.date);
    const type = data.type;
    const coin = data.coin.toUpperCase();
    const quantity = parseFloat(data.quantity);
    const priceUSD = parseFloat(data.priceUSD);
    const exchangeRate = parseFloat(data.exchangeRate) || 23000;
    const priceVND = priceUSD * exchangeRate;
    const totalVND = quantity * priceVND;
    const exchange = data.exchange || '';
    const wallet = data.wallet || '';
    const note = data.note || '';
    
    const lastRow = sheet.getLastRow();
    const stt = lastRow > 1 ? lastRow : 1;
    
    sheet.appendRow([stt, date, type, coin, quantity, priceUSD, exchangeRate, priceVND, totalVND, exchange, wallet, note]);
    
    const newRow = sheet.getLastRow();
    sheet.getRange(newRow, 2).setNumberFormat('dd/mm/yyyy');
    sheet.getRange(newRow, 6).setNumberFormat('#,##0.00" USD"');
    sheet.getRange(newRow, 8, 1, 2).setNumberFormat('#,##0" VNĐ"');
    
    if (type === 'Mua') {
      try { updateInvestmentBudget('Crypto', totalVND); } catch(e) {}
    }
    
    if (type === 'Bán') {
      try {
        addIncome({
          date: data.date,
          amount: totalVND,
          source: 'Bán Crypto',
          note: `Bán ${quantity} ${coin}`
        });
      } catch(e) {}
    }
    
    return {
      success: true,
      message: `✅ Đã ghi nhận ${type} ${quantity} ${coin}!`
    };
    
  } catch (error) {
    return { success: false, message: `❌ Lỗi: ${error.message}` };
  }
}

// ==================== ĐẦU TƯ KHÁC ====================

/**
 * Thêm khoản đầu tư khác
 */
function addOtherInvestment(data) {
  try {
    if (!data.date || !data.type || !data.amount) {
      return {
        success: false,
        message: '❌ Vui lòng điền đầy đủ thông tin!'
      };
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(APP_CONFIG.SHEETS.OTHER_INVESTMENT);
    
    if (!sheet) {
      return { success: false, message: '❌ Sheet ĐẦU TƯ KHÁC chưa được khởi tạo!' };
    }
    
    const date = new Date(data.date);
    const type = data.type;
    const amount = parseFloat(data.amount);
    const roi = parseFloat(data.roi) || 0;
    const term = parseInt(data.term) || 0;
    const expectedReturn = amount * (1 + (roi / 100) * (term / 12));
    const note = data.note || '';
    
    const lastRow = sheet.getLastRow();
    const stt = lastRow > 1 ? lastRow : 1;
    
    sheet.appendRow([stt, date, type, amount, roi, term, expectedReturn, note]);
    
    const newRow = sheet.getLastRow();
    sheet.getRange(newRow, 2).setNumberFormat('dd/mm/yyyy');
    sheet.getRange(newRow, 4).setNumberFormat('#,##0" VNĐ"');
    sheet.getRange(newRow, 5).setNumberFormat('0.00"%"');
    sheet.getRange(newRow, 7).setNumberFormat('#,##0" VNĐ"');
    
    try {
      updateInvestmentBudget('Đầu tư khác', amount);
    } catch(e) {}
    
    return {
      success: true,
      message: `✅ Đã ghi nhận đầu tư ${type} với số tiền ${formatCurrency(amount)}!`
    };
    
  } catch (error) {
    return { success: false, message: `❌ Lỗi: ${error.message}` };
  }
}