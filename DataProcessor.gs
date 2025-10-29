/**
 * ===============================================
 * DATAPROCESSORS.GS v3.3 - COMPLETE FIX
 * ===============================================
 */

// ==================== HÀM HỖ TRỢ - DEBT LIST ====================

/**
 * Lấy danh sách các khoản nợ đang có
 * @return {Array} Mảng tên các khoản nợ
 */
function getDebtList() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(APP_CONFIG.SHEETS.DEBT_MANAGEMENT);
    
    if (!sheet) {
      Logger.log('Sheet QUẢN LÝ NỢ không tồn tại');
      return [];
    }
    
    // ✅ FIX: Sử dụng findEmptyRow() để xác định số dòng có dữ liệu thực
    const emptyRow = findEmptyRow(sheet, 2);
    const dataRows = emptyRow - 2; // Trừ header và bắt đầu từ 0
    
    if (dataRows <= 0) {
      return [];
    }
    
    // Lấy dữ liệu từ cột B (Tên khoản nợ)
    const data = sheet.getRange(2, 2, dataRows, 1).getValues();
    
    // Lọc các dòng có tên khoản nợ
    const debtList = data
      .map(row => row[0])
      .filter(name => name && name.toString().trim() !== '');
    
    Logger.log('Danh sách nợ: ' + debtList.join(', '));
    return debtList;
    
  } catch (error) {
    Logger.log('Lỗi getDebtList: ' + error.message);
    return [];
  }
}

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
    
    // ✅ FIX: Sử dụng findEmptyRow() và getNextSTT()
    const emptyRow = findEmptyRow(sheet, 2); // Cột B = Ngày
    const stt = getNextSTT(sheet, 2);
    
    // Thêm dữ liệu
    const rowData = [stt, date, amount, source, note];
    sheet.getRange(emptyRow, 1, 1, rowData.length).setValues([rowData]);
    
    // Format dòng mới
    formatNewRow(sheet, emptyRow, {
      2: 'dd/mm/yyyy',
      3: '#,##0" VNĐ"'
    });
    
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
    
    // ✅ FIX: Sử dụng findEmptyRow() và getNextSTT()
    const emptyRow = findEmptyRow(sheet, 2); // Cột B = Ngày
    const stt = getNextSTT(sheet, 2);
    
    // Thêm dữ liệu
    const rowData = [stt, date, amount, category, detail, note];
    sheet.getRange(emptyRow, 1, 1, rowData.length).setValues([rowData]);
    
    // Format dòng mới
    formatNewRow(sheet, emptyRow, {
      2: 'dd/mm/yyyy',
      3: '#,##0" VNĐ"'
    });
    
    // Cập nhật Budget
    try {
      BudgetManager.updateExpenseBudget(category);
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

// ==================== TRẢ NỢ - FIXED VERSION ====================

/**
 * Thêm khoản trả nợ VÀ tự động cập nhật QUẢN LÝ NỢ
 * ✅ FIX: Logic trạng thái Chưa trả → Đang trả → Đã trả hết
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
    
    // Kiểm tra sheets
    const paymentSheet = ss.getSheetByName(APP_CONFIG.SHEETS.DEBT_PAYMENT);
    const managementSheet = ss.getSheetByName(APP_CONFIG.SHEETS.DEBT_MANAGEMENT);
    
    if (!paymentSheet) {
      return {
        success: false,
        message: '❌ Sheet TRẢ NỢ chưa được khởi tạo!'
      };
    }
    
    if (!managementSheet) {
      return {
        success: false,
        message: '❌ Sheet QUẢN LÝ NỢ chưa được khởi tạo!'
      };
    }
    
    // Parse dữ liệu
    const date = new Date(data.date);
    const debtName = data.debtName;
    const principal = parseFloat(data.principal) || 0;
    const interest = parseFloat(data.interest) || 0;
    const total = principal + interest;
    const note = data.note || '';
    
    // ✅ FIX: Thêm vào Sheet TRẢ NỢ
    const emptyRow = findEmptyRow(paymentSheet, 2); // Cột B = Ngày
    const stt = getNextSTT(paymentSheet, 2);
    
    const rowData = [stt, date, debtName, principal, interest, total, note];
    paymentSheet.getRange(emptyRow, 1, 1, rowData.length).setValues([rowData]);
    
    // Format dòng mới
    formatNewRow(paymentSheet, emptyRow, {
      2: 'dd/mm/yyyy',
      4: '#,##0" VNĐ"',
      5: '#,##0" VNĐ"',
      6: '#,##0" VNĐ"'
    });
    
    // Cập nhật Sheet QUẢN LÝ NỢ
    const debtRow = findDebtRow(managementSheet, debtName);
    
    if (debtRow === -1) {
      return {
        success: false,
        message: `❌ Không tìm thấy khoản nợ "${debtName}" trong QUẢN LÝ NỢ!`
      };
    }
    
    // Lấy thông tin nợ hiện tại
    const debtData = managementSheet.getRange(debtRow, 1, 1, 12).getValues()[0];
    const originalDebt = parseFloat(debtData[2]) || 0;
    const paidPrincipal = parseFloat(debtData[7]) || 0;
    const paidInterest = parseFloat(debtData[8]) || 0;
    
    // Tính toán giá trị mới
    const newPaidPrincipal = paidPrincipal + principal;
    const newPaidInterest = paidInterest + interest;
    const remaining = originalDebt - newPaidPrincipal;
    
    // ✅ FIX: Xác định trạng thái theo logic đúng
    let status = 'Đang trả';
    
    if (remaining <= 0) {
      status = 'Đã trả hết';
    } else if (newPaidPrincipal === 0 && newPaidInterest === 0) {
      // Nếu chưa trả gì cả (cả gốc và lãi đều = 0)
      status = 'Chưa trả';
    }
    
    // Cập nhật các cột H, I, K (cột J có công thức nên tự tính)
    managementSheet.getRange(debtRow, 8).setValue(newPaidPrincipal);  // H: Đã trả gốc
    managementSheet.getRange(debtRow, 9).setValue(newPaidInterest);   // I: Đã trả lãi
    // Cột J (Còn nợ) có công thức =C-H nên không cần set
    managementSheet.getRange(debtRow, 11).setValue(status);           // K: Trạng thái
    
    return {
      success: true,
      message: `✅ Đã ghi nhận trả nợ cho ${debtName}!\n` +
               `💰 Gốc: ${formatCurrency(principal)}\n` +
               `💳 Lãi: ${formatCurrency(interest)}\n` +
               `📊 Còn nợ: ${formatCurrency(remaining)}\n` +
               `📌 Trạng thái: ${status}`
    };
    
  } catch (error) {
    return {
      success: false,
      message: `❌ Lỗi: ${error.message}`
    };
  }
}

/**
 * Tìm dòng chứa khoản nợ theo tên
 * @param {Sheet} sheet - Sheet QUẢN LÝ NỢ
 * @param {string} debtName - Tên khoản nợ
 * @return {number} Row number hoặc -1 nếu không tìm thấy
 */
function findDebtRow(sheet, debtName) {
  const emptyRow = findEmptyRow(sheet, 2);
  const dataRows = emptyRow - 2;
  
  if (dataRows <= 0) {
    return -1;
  }
  
  const data = sheet.getRange(2, 2, dataRows, 1).getValues();
  
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === debtName) {
      return i + 2; // +2 vì bắt đầu từ dòng 2
    }
  }
  
  return -1;
}

// ==================== CHỨNG KHOÁN ====================

/**
 * Thêm giao dịch chứng khoán
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
    
    // ✅ FIX: Sử dụng findEmptyRow()
    const emptyRow = findEmptyRow(sheet, 2); // Cột B = Ngày
    const stt = getNextSTT(sheet, 2);
    
    const rowData = [stt, date, type, symbol, quantity, price, fee, total, note];
    sheet.getRange(emptyRow, 1, 1, rowData.length).setValues([rowData]);
    
    // Format
    formatNewRow(sheet, emptyRow, {
      2: 'dd/mm/yyyy',
      6: '#,##0" VNĐ"',
      7: '#,##0" VNĐ"',
      8: '#,##0" VNĐ"'
    });
    
    // Nếu mua thì trừ Budget
    if (type === 'Mua') {
      try {
        if (typeof BudgetManager !== 'undefined') {
          BudgetManager.updateInvestmentBudget('Chứng khoán', total);
        }
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
          source: 'Bán CK',
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
    
    // ✅ FIX: Sử dụng findEmptyRow()
    const emptyRow = findEmptyRow(sheet, 2);
    const stt = getNextSTT(sheet, 2);
    
    const rowData = [stt, date, type, goldType, quantity, unit, price, total, location, note];
    sheet.getRange(emptyRow, 1, 1, rowData.length).setValues([rowData]);
    
    formatNewRow(sheet, emptyRow, {
      2: 'dd/mm/yyyy',
      7: '#,##0" VNĐ"',
      8: '#,##0" VNĐ"'
    });
    
    if (type === 'Mua') {
      try { 
        if (typeof BudgetManager !== 'undefined') {
          BudgetManager.updateInvestmentBudget('Vàng', total);
        }
      } catch(e) {}
    }
    
    if (type === 'Bán') {
      try {
        addIncome({
          date: data.date,
          amount: total,
          source: 'Bán Vàng',
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
    
    // ✅ FIX: Sử dụng findEmptyRow()
    const emptyRow = findEmptyRow(sheet, 2);
    const stt = getNextSTT(sheet, 2);
    
    const rowData = [stt, date, type, coin, quantity, priceUSD, exchangeRate, priceVND, totalVND, exchange, wallet, note];
    sheet.getRange(emptyRow, 1, 1, rowData.length).setValues([rowData]);
    
    formatNewRow(sheet, emptyRow, {
      2: 'dd/mm/yyyy',
      6: '#,##0.00" USD"',
      8: '#,##0" VNĐ"',
      9: '#,##0" VNĐ"'
    });
    
    if (type === 'Mua') {
      try { 
        if (typeof BudgetManager !== 'undefined') {
          BudgetManager.updateInvestmentBudget('Crypto', totalVND);
        }
      } catch(e) {}
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
    
    // ✅ FIX: Sử dụng findEmptyRow()
    const emptyRow = findEmptyRow(sheet, 2);
    const stt = getNextSTT(sheet, 2);
    
    const rowData = [stt, date, type, amount, roi, term, expectedReturn, note];
    sheet.getRange(emptyRow, 1, 1, rowData.length).setValues([rowData]);
    
    formatNewRow(sheet, emptyRow, {
      2: 'dd/mm/yyyy',
      4: '#,##0" VNĐ"',
      5: '0.00"%"',
      7: '#,##0" VNĐ"'
    });
    
    try {
      if (typeof BudgetManager !== 'undefined') {
        BudgetManager.updateInvestmentBudget('Đầu tư khác', amount);
      }
    } catch(e) {}
    
    return {
      success: true,
      message: `✅ Đã ghi nhận đầu tư ${type} với số tiền ${formatCurrency(amount)}!`
    };
    
  } catch (error) {
    return { success: false, message: `❌ Lỗi: ${error.message}` };
  }
}