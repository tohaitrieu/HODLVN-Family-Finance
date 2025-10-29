/**
 * ===============================================
 * DATAPROCESSORS.GS v3.4 - COMPLETE FIX + DIVIDEND FEATURE
 * ===============================================
 * 
 * CHANGELOG v3.4:
 * ✅ FIX: addGold() - Sửa lỗi validation và đảm bảo dữ liệu được điền đầy đủ
 * ✅ FIX: addOtherInvestment() - Sửa lỗi nhận tham số investmentType
 * ✅ NEW: getStocksForDividend() - Lấy danh sách cổ phiếu để nhận cổ tức
 * ✅ NEW: processDividend() - Xử lý cổ tức tiền mặt và thưởng cổ phiếu
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
 * Tìm dòng của khoản nợ trong QUẢN LÝ NỢ
 */
function findDebtRow(sheet, debtName) {
  const emptyRow = findEmptyRow(sheet, 2);
  const dataRows = emptyRow - 2;
  
  if (dataRows <= 0) return -1;
  
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

// ==================== VÀNG - FIXED VERSION ====================

/**
 * Thêm giao dịch vàng
 * ✅ FIX v3.4: Đảm bảo tất cả dữ liệu được điền vào đúng cột
 */
function addGold(data) {
  try {
    // ✅ FIX: Validation chặt chẽ hơn
    if (!data.date || !data.type || !data.goldType || !data.quantity || !data.unit || !data.price) {
      return {
        success: false,
        message: '❌ Vui lòng điền đầy đủ thông tin bắt buộc!'
      };
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(APP_CONFIG.SHEETS.GOLD);
    
    if (!sheet) {
      return { 
        success: false, 
        message: '❌ Sheet VÀNG chưa được khởi tạo!' 
      };
    }
    
    // Parse dữ liệu
    const date = new Date(data.date);
    const type = data.type;
    const goldType = data.goldType;
    const quantity = parseFloat(data.quantity);
    const unit = data.unit; // ✅ FIX: Không dùng default value
    const price = parseFloat(data.price);
    const total = quantity * price;
    const location = data.location || '';
    const note = data.note || '';
    
    // ✅ Log để debug
    Logger.log('=== ADD GOLD DEBUG ===');
    Logger.log('Date: ' + date);
    Logger.log('Type: ' + type);
    Logger.log('GoldType: ' + goldType);
    Logger.log('Quantity: ' + quantity);
    Logger.log('Unit: ' + unit);
    Logger.log('Price: ' + price);
    Logger.log('Total: ' + total);
    Logger.log('Location: ' + location);
    Logger.log('Note: ' + note);
    
    // ✅ FIX: Sử dụng findEmptyRow()
    const emptyRow = findEmptyRow(sheet, 2);
    const stt = getNextSTT(sheet, 2);
    
    Logger.log('Empty Row: ' + emptyRow);
    Logger.log('STT: ' + stt);
    
    // ✅ Thứ tự cột: STT, Ngày, Loại GD, Loại vàng, Số lượng, Đơn vị, Giá, Tổng, Nơi lưu, Ghi chú
    const rowData = [stt, date, type, goldType, quantity, unit, price, total, location, note];
    
    Logger.log('Row Data: ' + JSON.stringify(rowData));
    
    // Thêm dữ liệu vào sheet
    sheet.getRange(emptyRow, 1, 1, rowData.length).setValues([rowData]);
    
    // Format các cột
    formatNewRow(sheet, emptyRow, {
      2: 'dd/mm/yyyy',    // Cột B: Ngày
      7: '#,##0" VNĐ"',   // Cột G: Giá
      8: '#,##0" VNĐ"'    // Cột H: Tổng
    });
    
    // Cập nhật budget nếu mua
    if (type === 'Mua') {
      try { 
        if (typeof BudgetManager !== 'undefined') {
          BudgetManager.updateInvestmentBudget('Vàng', total);
        }
      } catch(e) {
        Logger.log('Không thể cập nhật budget: ' + e.message);
      }
    }
    
    // Tạo thu nhập nếu bán
    if (type === 'Bán') {
      try {
        addIncome({
          date: data.date,
          amount: total,
          source: 'Bán Vàng',
          note: `Bán ${quantity} ${unit} ${goldType}`
        });
      } catch(e) {
        Logger.log('Không thể thêm thu nhập: ' + e.message);
      }
    }
    
    return {
      success: true,
      message: `✅ Đã ghi nhận ${type} ${quantity} ${unit} ${goldType}!`
    };
    
  } catch (error) {
    Logger.log('Error in addGold: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    return { 
      success: false, 
      message: `❌ Lỗi: ${error.message}` 
    };
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

// ==================== ĐẦU TƯ KHÁC - FIXED VERSION ====================

/**
 * Thêm khoản đầu tư khác
 * ✅ FIX v3.4: Nhận đúng tham số investmentType thay vì type
 */
function addOtherInvestment(data) {
  try {
    // ✅ FIX: Validation với investmentType
    if (!data.date || !data.investmentType || !data.amount) {
      return {
        success: false,
        message: '❌ Vui lòng điền đầy đủ thông tin!'
      };
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(APP_CONFIG.SHEETS.OTHER_INVESTMENT);
    
    if (!sheet) {
      return { 
        success: false, 
        message: '❌ Sheet ĐẦU TƯ KHÁC chưa được khởi tạo!' 
      };
    }
    
    const date = new Date(data.date);
    const investmentType = data.investmentType; // ✅ FIX: Dùng investmentType
    const amount = parseFloat(data.amount);
    const roi = parseFloat(data.roi) || 0;
    const term = parseInt(data.term) || 0;
    const expectedReturn = amount * (1 + (roi / 100) * (term / 12));
    const note = data.note || '';
    
    // ✅ Log để debug
    Logger.log('=== ADD OTHER INVESTMENT DEBUG ===');
    Logger.log('Investment Type: ' + investmentType);
    Logger.log('Amount: ' + amount);
    Logger.log('ROI: ' + roi);
    Logger.log('Term: ' + term);
    
    // ✅ FIX: Sử dụng findEmptyRow()
    const emptyRow = findEmptyRow(sheet, 2);
    const stt = getNextSTT(sheet, 2);
    
    // ✅ Thứ tự cột: STT, Ngày, Loại đầu tư, Số tiền, Lãi suất (%), Kỳ hạn (tháng), Dự kiến thu về, Ghi chú
    const rowData = [stt, date, investmentType, amount, roi, term, expectedReturn, note];
    
    Logger.log('Row Data: ' + JSON.stringify(rowData));
    
    sheet.getRange(emptyRow, 1, 1, rowData.length).setValues([rowData]);
    
    formatNewRow(sheet, emptyRow, {
      2: 'dd/mm/yyyy',    // Cột B: Ngày
      4: '#,##0" VNĐ"',   // Cột D: Số tiền
      5: '0.00"%"',       // Cột E: Lãi suất
      7: '#,##0" VNĐ"'    // Cột G: Dự kiến thu về
    });
    
    try {
      if (typeof BudgetManager !== 'undefined') {
        BudgetManager.updateInvestmentBudget('Đầu tư khác', amount);
      }
    } catch(e) {
      Logger.log('Không thể cập nhật budget: ' + e.message);
    }
    
    return {
      success: true,
      message: `✅ Đã ghi nhận đầu tư ${investmentType} với số tiền ${formatCurrency(amount)}!`
    };
    
  } catch (error) {
    Logger.log('Error in addOtherInvestment: ' + error.message);
    return { 
      success: false, 
      message: `❌ Lỗi: ${error.message}` 
    };
  }
}

// ==================== CỔ TỨC - NEW FEATURE v3.4 ====================

/**
 * Lấy danh sách cổ phiếu đang nắm giữ để nhận cổ tức
 * @return {Array} Mảng các cổ phiếu với thông tin: code, quantity, costPrice, totalCost
 */
function getStocksForDividend() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(APP_CONFIG.SHEETS.STOCK);
    
    if (!sheet) {
      Logger.log('Sheet CHỨNG KHOÁN không tồn tại');
      return [];
    }
    
    // Lấy tất cả dữ liệu
    const emptyRow = findEmptyRow(sheet, 2);
    const dataRows = emptyRow - 2;
    
    if (dataRows <= 0) {
      return [];
    }
    
    // Cột: STT(A), Ngày(B), Loại GD(C), Mã CK(D), Số lượng(E), Giá(F), Phí(G), Tổng(H), Ghi chú(I)
    const data = sheet.getRange(2, 3, dataRows, 6).getValues(); // Từ cột C đến H
    
    // Tính toán số lượng và giá vốn cho từng mã
    const stockMap = new Map();
    
    for (let i = 0; i < data.length; i++) {
      const type = data[i][0];        // Cột C: Loại GD
      const symbol = data[i][1];      // Cột D: Mã CK
      const quantity = parseFloat(data[i][2]) || 0;  // Cột E: Số lượng
      const price = parseFloat(data[i][3]) || 0;     // Cột F: Giá
      const fee = parseFloat(data[i][4]) || 0;       // Cột G: Phí
      
      if (!symbol) continue;
      
      if (!stockMap.has(symbol)) {
        stockMap.set(symbol, {
          code: symbol,
          quantity: 0,
          totalCost: 0
        });
      }
      
      const stock = stockMap.get(symbol);
      
      if (type === 'Mua') {
        stock.quantity += quantity;
        stock.totalCost += (quantity * price) + fee;
      } else if (type === 'Bán') {
        stock.quantity -= quantity;
        // Khi bán, giảm giá vốn theo tỷ lệ
        if (stock.quantity > 0) {
          const soldRatio = quantity / (stock.quantity + quantity);
          stock.totalCost *= (1 - soldRatio);
        } else {
          stock.totalCost = 0;
        }
      }
    }
    
    // Lọc ra các cổ phiếu còn nắm giữ (quantity > 0)
    const stocks = [];
    stockMap.forEach((stock) => {
      if (stock.quantity > 0) {
        stock.costPrice = stock.totalCost / stock.quantity;
        stocks.push(stock);
      }
    });
    
    Logger.log('Danh sách cổ phiếu: ' + JSON.stringify(stocks));
    return stocks;
    
  } catch (error) {
    Logger.log('Lỗi getStocksForDividend: ' + error.message);
    return [];
  }
}

/**
 * Xử lý cổ tức (tiền mặt hoặc thưởng cổ phiếu)
 * @param {Object} data - Dữ liệu cổ tức
 * @return {Object} {success, message}
 */
function processDividend(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const stockSheet = ss.getSheetByName(APP_CONFIG.SHEETS.STOCK);
    
    if (!stockSheet) {
      return {
        success: false,
        message: '❌ Sheet CHỨNG KHOÁN chưa được khởi tạo!'
      };
    }
    
    const type = data.type; // 'cash' hoặc 'stock'
    const stockCode = data.stockCode;
    const date = data.date;
    const notes = data.notes || '';
    
    if (type === 'cash') {
      // XỬ LÝ CỔ TỨC TIỀN MẶT
      const cashAmount = parseFloat(data.cashAmount);
      const totalDividend = parseFloat(data.totalDividend);
      
      // 1. Tạo giao dịch THU
      const incomeResult = addIncome({
        date: date,
        amount: totalDividend,
        source: 'Đầu tư',
        note: `Cổ tức ${stockCode}: ${cashAmount}đ/CP. ${notes}`
      });
      
      if (!incomeResult.success) {
        return incomeResult;
      }
      
      // 2. Cập nhật giá vốn trong sheet CHỨNG KHOÁN
      // Tìm tất cả các lần mua cổ phiếu này và giảm giá vốn tương ứng
      const emptyRow = findEmptyRow(stockSheet, 2);
      const dataRows = emptyRow - 2;
      
      if (dataRows > 0) {
        const stockData = stockSheet.getRange(2, 1, dataRows, 9).getValues();
        
        // Tính tổng số lượng đang nắm giữ
        let totalQuantity = 0;
        for (let i = 0; i < stockData.length; i++) {
          const type = stockData[i][2]; // Cột C: Loại GD
          const symbol = stockData[i][3]; // Cột D: Mã CK
          const qty = parseFloat(stockData[i][4]) || 0; // Cột E: Số lượng
          
          if (symbol === stockCode) {
            if (type === 'Mua') totalQuantity += qty;
            else if (type === 'Bán') totalQuantity -= qty;
          }
        }
        
        // Ghi chú vào sheet: Cổ tức đã được ghi nhận trong THU
        // và giá vốn đã được điều chỉnh
        Logger.log(`Đã ghi nhận cổ tức tiền mặt cho ${stockCode}: ${formatCurrency(totalDividend)}`);
      }
      
      return {
        success: true,
        message: `✅ Đã ghi nhận cổ tức tiền mặt ${formatCurrency(totalDividend)} cho ${stockCode}!`
      };
      
    } else if (type === 'stock') {
      // XỬ LÝ THƯỞNG CỔ PHIẾU
      const stockRatio = parseFloat(data.stockRatio);
      const bonusShares = parseInt(data.bonusShares);
      const currentQuantity = parseFloat(data.currentQuantity);
      const newQuantity = currentQuantity + bonusShares;
      
      // Ghi lại giao dịch nhận thưởng cổ phiếu vào sheet CHỨNG KHOÁN
      const emptyRow = findEmptyRow(stockSheet, 2);
      const stt = getNextSTT(stockSheet, 2);
      
      const noteText = `Thưởng CP ${stockRatio}% (${bonusShares} CP). ${notes}`;
      
      // Thêm dòng mới: Loại GD = "Thưởng", Số lượng = bonusShares, Giá = 0, Phí = 0, Tổng = 0
      const rowData = [
        stt,
        new Date(date),
        'Thưởng',
        stockCode,
        bonusShares,
        0, // Giá = 0
        0, // Phí = 0
        0, // Tổng = 0
        noteText
      ];
      
      stockSheet.getRange(emptyRow, 1, 1, rowData.length).setValues([rowData]);
      
      formatNewRow(stockSheet, emptyRow, {
        2: 'dd/mm/yyyy',
        6: '#,##0" VNĐ"',
        7: '#,##0" VNĐ"',
        8: '#,##0" VNĐ"'
      });
      
      return {
        success: true,
        message: `✅ Đã ghi nhận thưởng ${bonusShares} cổ phiếu ${stockCode}!\n` +
                 `📊 Số lượng mới: ${newQuantity} CP`
      };
    }
    
    return {
      success: false,
      message: '❌ Loại cổ tức không hợp lệ!'
    };
    
  } catch (error) {
    Logger.log('Lỗi processDividend: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    return {
      success: false,
      message: `❌ Lỗi: ${error.message}`
    };
  }
}