/**
 * ===============================================
 * DATAPROCESSORS.GS v3.5.1 - COMPLETE FIX + NEW DIVIDEND LOGIC
 * ===============================================
 * 
 * CHANGELOG v3.5.1:
 * ✅ FIX: addStockTransaction() - Ghi đúng 16 cột theo cấu trúc mới
 * ✅ FIX: processDividend() - Cập nhật CỘT I (Cổ tức TM) thay vì giảm cột H
 * ✅ LOGIC: Cổ tức tiền mặt CỘNG DỒN vào cột I, cột K tự động tính giá điều chỉnh
 * ✅ LOGIC: Thêm lịch sử cổ tức vào cột P (Ghi chú)
 * ✅ FIX: getStocksForDividend() - Đọc đúng cột I (Cổ tức TM) và tính giá điều chỉnh
 * 
 * CHANGELOG v3.5:
 * ✅ FIX: processDividend() - ĐIỀU CHỈNH GIÁ CỔ PHIẾU khi nhận cổ tức tiền mặt
 * ✅ LOGIC: Cổ tức tiền mặt giảm giá vốn trực tiếp cho TẤT CẢ giao dịch mua
 * ✅ LOGIC: Thưởng cổ phiếu tăng số lượng, giữ nguyên tổng giá vốn
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
    
    const emptyRow = findEmptyRow(sheet, 2);
    const dataRows = emptyRow - 2;
    
    if (dataRows <= 0) {
      return [];
    }
    
    const data = sheet.getRange(2, 2, dataRows, 1).getValues();
    
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
    
    const emptyRow = findEmptyRow(sheet, 2);
    const stt = getNextSTT(sheet, 2);
    
    const date = new Date(data.date);
    const amount = parseFloat(data.amount);
    const source = data.source.toString();
    const note = data.note || '';
    
    const rowData = [
      stt,
      date,
      amount,
      source,
      note
    ];
    
    sheet.getRange(emptyRow, 1, 1, rowData.length).setValues([rowData]);
    
    formatNewRow(sheet, emptyRow, {
      2: 'dd/mm/yyyy',
      3: '#,##0" VNĐ"'
    });
    
    Logger.log(`Đã thêm thu nhập: ${formatCurrency(amount)} - ${source}`);
    
    return {
      success: true,
      message: `✅ Đã ghi nhận thu nhập ${formatCurrency(amount)} từ ${source}!`
    };
    
  } catch (error) {
    Logger.log('Error in addIncome: ' + error.message);
    return {
      success: false,
      message: `❌ Lỗi: ${error.message}`
    };
  }
}

// ==================== CHI TIÊU ====================

/**
 * Thêm khoản chi tiêu
 * @param {Object} data - {date, amount, category, subcategory, note}
 * @return {Object} {success, message}
 */
function addExpense(data) {
  try {
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
        message: '❌ Sheet CHI chưa được khởi tạo!'
      };
    }
    
    const emptyRow = findEmptyRow(sheet, 2);
    const stt = getNextSTT(sheet, 2);
    
    const date = new Date(data.date);
    const amount = parseFloat(data.amount);
    const category = data.category.toString();
    const subcategory = data.subcategory || '';
    const note = data.note || '';
    
    const rowData = [
      stt,
      date,
      amount,
      category,
      subcategory,
      note
    ];
    
    sheet.getRange(emptyRow, 1, 1, rowData.length).setValues([rowData]);
    
    formatNewRow(sheet, emptyRow, {
      2: 'dd/mm/yyyy',
      3: '#,##0" VNĐ"'
    });
    
    BudgetManager.updateBudgetSpent(category);
    
    Logger.log(`Đã thêm chi tiêu: ${formatCurrency(amount)} - ${category}`);
    
    return {
      success: true,
      message: `✅ Đã ghi nhận chi tiêu ${formatCurrency(amount)} cho ${category}!`
    };
    
  } catch (error) {
    Logger.log('Error in addExpense: ' + error.message);
    return {
      success: false,
      message: `❌ Lỗi: ${error.message}`
    };
  }
}

// ==================== QUẢN LÝ NỢ ====================

/**
 * Thêm khoản nợ
 * @param {Object} data - {loanDate, debtName, amount, interestRate, term, purpose, note}
 * @return {Object} {success, message}
 */
function addDebt(data) {
  try {
    if (!data.loanDate || !data.debtName || !data.amount) {
      return {
        success: false,
        message: '❌ Vui lòng điền đầy đủ thông tin bắt buộc!'
      };
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(APP_CONFIG.SHEETS.DEBT_MANAGEMENT);
    
    if (!sheet) {
      return {
        success: false,
        message: '❌ Sheet QUẢN LÝ NỢ chưa được khởi tạo!'
      };
    }
    
    const emptyRow = findEmptyRow(sheet, 2);
    const stt = getNextSTT(sheet, 2);
    
    const loanDate = new Date(data.loanDate);
    const debtName = data.debtName.toString();
    const amount = parseFloat(data.amount);
    const interestRate = parseFloat(data.interestRate) || 0;
    const term = parseInt(data.term) || 12;
    const purpose = data.purpose || '';
    const note = data.note || '';
    
    const dueDate = new Date(loanDate);
    dueDate.setMonth(dueDate.getMonth() + term);
    
    const rowData = [
      stt,
      debtName,
      loanDate,
      amount,
      amount,
      interestRate / 100,
      term,
      dueDate,
      purpose,
      'Chưa trả',
      note
    ];
    
    sheet.getRange(emptyRow, 1, 1, rowData.length).setValues([rowData]);
    
    formatNewRow(sheet, emptyRow, {
      3: 'dd/mm/yyyy',
      4: '#,##0" VNĐ"',
      5: '#,##0" VNĐ"',
      6: '0.00%',
      8: 'dd/mm/yyyy'
    });
    
    const incomeResult = addIncome({
      date: loanDate,
      amount: amount,
      source: 'Vay nợ',
      note: `Vay ${debtName}. ${note}`
    });
    
    if (!incomeResult.success) {
      Logger.log('Cảnh báo: Không thể tạo khoản thu tự động cho nợ');
    }
    
    Logger.log(`Đã thêm khoản nợ: ${debtName} - ${formatCurrency(amount)}`);
    
    return {
      success: true,
      message: `✅ Đã ghi nhận khoản nợ ${debtName}: ${formatCurrency(amount)}!\n` +
               `📅 Hạn thanh toán: ${formatDate(dueDate)}`
    };
    
  } catch (error) {
    Logger.log('Error in addDebt: ' + error.message);
    return {
      success: false,
      message: `❌ Lỗi: ${error.message}`
    };
  }
}

/**
 * Trả nợ
 * @param {Object} data - {date, debtName, principalAmount, interestAmount, note}
 * @return {Object} {success, message}
 */
function payDebt(data) {
  try {
    if (!data.date || !data.debtName || !data.principalAmount) {
      return {
        success: false,
        message: '❌ Vui lòng điền đầy đủ thông tin bắt buộc!'
      };
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const paymentSheet = ss.getSheetByName(APP_CONFIG.SHEETS.DEBT_PAYMENT);
    const debtSheet = ss.getSheetByName(APP_CONFIG.SHEETS.DEBT_MANAGEMENT);
    
    if (!paymentSheet || !debtSheet) {
      return {
        success: false,
        message: '❌ Các sheet liên quan chưa được khởi tạo!'
      };
    }
    
    const emptyRow = findEmptyRow(paymentSheet, 2);
    const stt = getNextSTT(paymentSheet, 2);
    
    const date = new Date(data.date);
    const debtName = data.debtName.toString();
    const principalAmount = parseFloat(data.principalAmount);
    const interestAmount = parseFloat(data.interestAmount) || 0;
    const totalAmount = principalAmount + interestAmount;
    const note = data.note || '';
    
    const rowData = [
      stt,
      date,
      debtName,
      principalAmount,
      interestAmount,
      totalAmount,
      note
    ];
    
    paymentSheet.getRange(emptyRow, 1, 1, rowData.length).setValues([rowData]);
    
    formatNewRow(paymentSheet, emptyRow, {
      2: 'dd/mm/yyyy',
      4: '#,##0" VNĐ"',
      5: '#,##0" VNĐ"',
      6: '#,##0" VNĐ"'
    });
    
    const debtEmptyRow = findEmptyRow(debtSheet, 2);
    const debtDataRows = debtEmptyRow - 2;
    
    if (debtDataRows > 0) {
      const debtData = debtSheet.getRange(2, 2, debtDataRows, 10).getValues();
      
      for (let i = 0; i < debtData.length; i++) {
        const rowDebtName = debtData[i][0];
        
        if (rowDebtName === debtName) {
          const currentBalance = parseFloat(debtData[i][3]) || 0;
          const newBalance = currentBalance - principalAmount;
          const row = i + 2;
          
          debtSheet.getRange(row, 5).setValue(newBalance);
          
          if (newBalance <= 0) {
            debtSheet.getRange(row, 10).setValue('Đã thanh toán');
          }
          
          break;
        }
      }
    }
    
    BudgetManager.updateDebtBudget();
    
    Logger.log(`Đã trả nợ: ${debtName} - Gốc: ${formatCurrency(principalAmount)}, Lãi: ${formatCurrency(interestAmount)}`);
    
    return {
      success: true,
      message: `✅ Đã ghi nhận trả nợ ${debtName}!\n` +
               `💰 Gốc: ${formatCurrency(principalAmount)}\n` +
               `📊 Lãi: ${formatCurrency(interestAmount)}\n` +
               `💵 Tổng: ${formatCurrency(totalAmount)}`
    };
    
  } catch (error) {
    Logger.log('Error in payDebt: ' + error.message);
    return {
      success: false,
      message: `❌ Lỗi: ${error.message}`
    };
  }
}

// ==================== CHỨNG KHOÁN ====================

/**
 * Thêm giao dịch chứng khoán
 * @param {Object} data - {date, type, stockCode, quantity, price, fee, useMargin, marginAmount, marginRate, note}
 * @return {Object} {success, message}
 */
function addStock(data) {
  try {
    // ✅ v3.5.1: Sửa validation + hỗ trợ cả symbol và stockCode
    Logger.log('addStock received data: ' + JSON.stringify(data));
    
    // Hỗ trợ cả 2 tên parameter: symbol (từ form) và stockCode (từ code cũ)
    const stockCode = data.stockCode || data.symbol;
    
    if (!data.date) {
      Logger.log('Missing date');
      return { success: false, message: '❌ Thiếu ngày giao dịch!' };
    }
    if (!data.type) {
      Logger.log('Missing type');
      return { success: false, message: '❌ Thiếu loại giao dịch!' };
    }
    if (!stockCode) {
      Logger.log('Missing stockCode/symbol');
      return { success: false, message: '❌ Thiếu mã cổ phiếu!' };
    }
    if (!data.quantity) {
      Logger.log('Missing quantity');
      return { success: false, message: '❌ Thiếu số lượng!' };
    }
    if (data.price === undefined || data.price === null || data.price === '') {
      Logger.log('Missing price');
      return { success: false, message: '❌ Thiếu giá!' };
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(APP_CONFIG.SHEETS.STOCK);
    
    if (!sheet) {
      return {
        success: false,
        message: '❌ Sheet CHỨNG KHOÁN chưa được khởi tạo!'
      };
    }
    
    const emptyRow = findEmptyRow(sheet, 2);
    const stt = getNextSTT(sheet, 2);
    
    const date = new Date(data.date);
    const type = data.type.toString();
    // ✅ Dùng biến stockCode đã được định nghĩa ở trên (hỗ trợ cả symbol và stockCode)
    const stockCodeUpper = stockCode.toString().toUpperCase();
    const quantity = parseInt(data.quantity);
    const price = parseFloat(data.price);
    const fee = parseFloat(data.fee) || 0;
    const total = (quantity * price) + fee;
    const note = data.note || '';
    
    // ✅ v3.5.1: Ghi đúng 16 cột - KHÔNG GHI ĐÈ công thức
    // Chỉ ghi dữ liệu vào cột A-J và P, bỏ qua K-O (để công thức tự động)
    const rowData = [
      stt,           // A: STT
      date,          // B: Ngày
      type,          // C: Loại GD
      stockCodeUpper,     // D: Mã CK
      quantity,      // E: Số lượng
      price,         // F: Giá gốc
      fee,           // G: Phí
      total,         // H: Tổng vốn
      0,             // I: Cổ tức TM (khởi tạo = 0)
      0              // J: Cổ tức CP (khởi tạo = 0)
    ];
    
    // Ghi dữ liệu vào cột A-J (10 cột đầu)
    sheet.getRange(emptyRow, 1, 1, 10).setValues([rowData]);
    
    // Ghi ghi chú vào cột P (cột 16)
    sheet.getRange(emptyRow, 16).setValue(note);
    
    // ✅ SET CÔNG THỨC cho cột K-O
    // K: Giá điều chỉnh = (Tổng vốn - Cổ tức TM) / Số lượng
    sheet.getRange(emptyRow, 11).setFormula(`=IF(E${emptyRow}>0, (H${emptyRow}-I${emptyRow})/E${emptyRow}, 0)`);
    
    // M: Giá trị HT = Số lượng × Giá HT
    sheet.getRange(emptyRow, 13).setFormula(`=IF(AND(E${emptyRow}>0, L${emptyRow}>0), E${emptyRow}*L${emptyRow}, 0)`);
    
    // N: Lãi/Lỗ = Giá trị HT - (Tổng vốn - Cổ tức TM)
    sheet.getRange(emptyRow, 14).setFormula(`=IF(M${emptyRow}>0, M${emptyRow}-(H${emptyRow}-I${emptyRow}), 0)`);
    
    // O: % L/L = Lãi/Lỗ / (Tổng vốn - Cổ tức TM)
    sheet.getRange(emptyRow, 15).setFormula(`=IF(AND(N${emptyRow}<>0, (H${emptyRow}-I${emptyRow})>0), N${emptyRow}/(H${emptyRow}-I${emptyRow}), 0)`);
    
    formatNewRow(sheet, emptyRow, {
      2: 'dd/mm/yyyy',
      6: '#,##0" VNĐ"',
      7: '#,##0" VNĐ"',
      8: '#,##0" VNĐ"',
      9: '#,##0" VNĐ"',
      11: '#,##0" VNĐ"',
      12: '#,##0" VNĐ"',
      13: '#,##0" VNĐ"',
      14: '#,##0" VNĐ"',
      15: '0.00%'
    });
    
    if (data.useMargin && data.marginAmount > 0) {
      const marginDebt = {
        loanDate: date,
        debtName: `Margin ${stockCodeUpper}`,
        amount: parseFloat(data.marginAmount),
        interestRate: parseFloat(data.marginRate) || 8.5,
        term: 3,
        purpose: `Vay margin mua ${stockCodeUpper}`,
        note: 'Tự động từ giao dịch chứng khoán'
      };
      
      addDebt(marginDebt);
    }
    
    BudgetManager.updateInvestmentBudget('Chứng khoán', total);
    
    Logger.log(`Đã thêm giao dịch CK: ${type} ${quantity} ${stockCodeUpper} @ ${formatCurrency(price)}`);
    
    return {
      success: true,
      message: `✅ Đã ghi nhận ${type} ${quantity} CP ${stockCodeUpper}!\n` +
               `💰 Giá: ${formatCurrency(price)}/CP\n` +
               `💵 Tổng: ${formatCurrency(total)}`
    };
    
  } catch (error) {
    Logger.log('Error in addStock: ' + error.message);
    return {
      success: false,
      message: `❌ Lỗi: ${error.message}`
    };
  }
}

// ==================== VÀNG ====================

/**
 * Thêm giao dịch vàng
 * @param {Object} data - {date, type, goldType, unit, quantity, price, note}
 * @return {Object} {success, message}
 */
function addGold(data) {
  try {
    if (!data.date || !data.type || !data.goldType || !data.quantity || !data.price) {
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
    
    const emptyRow = findEmptyRow(sheet, 2);
    const stt = getNextSTT(sheet, 2);
    
    const date = new Date(data.date);
    const type = data.type.toString();
    const goldType = data.goldType.toString();
    const unit = data.unit || 'Lượng';
    const quantity = parseFloat(data.quantity);
    const price = parseFloat(data.price);
    const total = quantity * price;
    const note = data.note || '';
    
    const rowData = [
      stt,
      date,
      type,
      goldType,
      unit,
      quantity,
      price,
      total,
      note
    ];
    
    sheet.getRange(emptyRow, 1, 1, rowData.length).setValues([rowData]);
    
    formatNewRow(sheet, emptyRow, {
      2: 'dd/mm/yyyy',
      7: '#,##0" VNĐ"',
      8: '#,##0" VNĐ"'
    });
    
    BudgetManager.updateInvestmentBudget('Vàng', total);
    
    Logger.log(`Đã thêm giao dịch vàng: ${type} ${quantity} ${unit} ${goldType}`);
    
    return {
      success: true,
      message: `✅ Đã ghi nhận ${type} ${quantity} ${unit} ${goldType}!\n` +
               `💰 Giá: ${formatCurrency(price)}/${unit}\n` +
               `💵 Tổng: ${formatCurrency(total)}`
    };
    
  } catch (error) {
    Logger.log('Error in addGold: ' + error.message);
    return {
      success: false,
      message: `❌ Lỗi: ${error.message}`
    };
  }
}

// ==================== CRYPTO ====================

/**
 * Thêm giao dịch crypto
 * @param {Object} data - {date, type, coin, quantity, price, fee, note}
 * @return {Object} {success, message}
 */
function addCrypto(data) {
  try {
    if (!data.date || !data.type || !data.coin || !data.quantity || !data.price) {
      return {
        success: false,
        message: '❌ Vui lòng điền đầy đủ thông tin bắt buộc!'
      };
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(APP_CONFIG.SHEETS.CRYPTO);
    
    if (!sheet) {
      return {
        success: false,
        message: '❌ Sheet CRYPTO chưa được khởi tạo!'
      };
    }
    
    const emptyRow = findEmptyRow(sheet, 2);
    const stt = getNextSTT(sheet, 2);
    
    const date = new Date(data.date);
    const type = data.type.toString();
    const coin = data.coin.toString().toUpperCase();
    const quantity = parseFloat(data.quantity);
    const price = parseFloat(data.price);
    const fee = parseFloat(data.fee) || 0;
    const total = (quantity * price) + fee;
    const note = data.note || '';
    
    const rowData = [
      stt,
      date,
      type,
      coin,
      quantity,
      price,
      fee,
      total,
      note
    ];
    
    sheet.getRange(emptyRow, 1, 1, rowData.length).setValues([rowData]);
    
    formatNewRow(sheet, emptyRow, {
      2: 'dd/mm/yyyy',
      6: '#,##0" VNĐ"',
      7: '#,##0" VNĐ"',
      8: '#,##0" VNĐ"'
    });
    
    BudgetManager.updateInvestmentBudget('Crypto', total);
    
    Logger.log(`Đã thêm giao dịch crypto: ${type} ${quantity} ${coin}`);
    
    return {
      success: true,
      message: `✅ Đã ghi nhận ${type} ${quantity} ${coin}!\n` +
               `💰 Giá: ${formatCurrency(price)}/${coin}\n` +
               `💵 Tổng: ${formatCurrency(total)}`
    };
    
  } catch (error) {
    Logger.log('Error in addCrypto: ' + error.message);
    return {
      success: false,
      message: `❌ Lỗi: ${error.message}`
    };
  }
}

// ==================== ĐẦU TƯ KHÁC ====================

/**
 * Thêm giao dịch đầu tư khác
 * @param {Object} data - {date, investmentType, amount, note}
 * @return {Object} {success, message}
 */
function addOtherInvestment(data) {
  try {
    if (!data.date || !data.investmentType || !data.amount) {
      return {
        success: false,
        message: '❌ Vui lòng điền đầy đủ thông tin bắt buộc!'
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
    
    const emptyRow = findEmptyRow(sheet, 2);
    const stt = getNextSTT(sheet, 2);
    
    const date = new Date(data.date);
    const investmentType = data.investmentType.toString();
    const amount = parseFloat(data.amount);
    const note = data.note || '';
    
    const rowData = [
      stt,
      date,
      investmentType,
      amount,
      note
    ];
    
    sheet.getRange(emptyRow, 1, 1, rowData.length).setValues([rowData]);
    
    formatNewRow(sheet, emptyRow, {
      2: 'dd/mm/yyyy',
      4: '#,##0" VNĐ"'
    });
    
    BudgetManager.updateInvestmentBudget('Đầu tư khác', amount);
    
    Logger.log(`Đã thêm đầu tư khác: ${investmentType} - ${formatCurrency(amount)}`);
    
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

// ==================== CỔ TỨC - v3.4/v3.5 FEATURE ====================

/**
 * Lấy danh sách cổ phiếu đang nắm giữ để nhận cổ tức
 * ✅ v3.5: Cập nhật đọc từ cột I (Cổ tức TM)
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
    
    const emptyRow = findEmptyRow(sheet, 2);
    const dataRows = emptyRow - 2;
    
    if (dataRows <= 0) {
      return [];
    }
    
    // ✅ v3.5: Đọc đúng theo cấu trúc 16 cột
    // Cột: STT(A), Ngày(B), Loại GD(C), Mã CK(D), Số lượng(E), Giá gốc(F), Phí(G), Tổng vốn(H), Cổ tức TM(I)
    const data = sheet.getRange(2, 3, dataRows, 7).getValues(); // Từ cột C đến I
    
    const stockMap = new Map();
    
    for (let i = 0; i < data.length; i++) {
      const type = data[i][0];        // Cột C: Loại GD
      const symbol = data[i][1];      // Cột D: Mã CK
      const quantity = parseFloat(data[i][2]) || 0;  // Cột E: Số lượng
      const price = parseFloat(data[i][3]) || 0;     // Cột F: Giá gốc
      const fee = parseFloat(data[i][4]) || 0;       // Cột G: Phí
      const totalCost = parseFloat(data[i][5]) || 0; // Cột H: Tổng vốn
      const dividendReceived = parseFloat(data[i][6]) || 0; // Cột I: Cổ tức TM đã nhận
      
      if (!symbol) continue;
      
      if (!stockMap.has(symbol)) {
        stockMap.set(symbol, {
          code: symbol,
          quantity: 0,
          totalCost: 0,
          totalDividend: 0
        });
      }
      
      const stock = stockMap.get(symbol);
      
      if (type === 'Mua') {
        stock.quantity += quantity;
        stock.totalCost += totalCost;
        stock.totalDividend += dividendReceived;
      } else if (type === 'Bán') {
        stock.quantity -= quantity;
        if (stock.quantity > 0) {
          const soldRatio = quantity / (stock.quantity + quantity);
          stock.totalCost *= (1 - soldRatio);
          stock.totalDividend *= (1 - soldRatio);
        } else {
          stock.totalCost = 0;
          stock.totalDividend = 0;
        }
      } else if (type === 'Thưởng') {
        // Thưởng cổ phiếu: tăng số lượng, giá vốn không đổi
        stock.quantity += quantity;
      }
    }
    
    const stocks = [];
    stockMap.forEach((stock) => {
      if (stock.quantity > 0) {
        // Giá vốn điều chỉnh = (Tổng vốn - Cổ tức đã nhận) / Số lượng
        const adjustedCost = stock.totalCost - stock.totalDividend;
        stock.costPrice = adjustedCost / stock.quantity;
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
 * ============================================================
 * XỬ LÝ CỔ TỨC (TIỀN MẶT & THƯỞNG CỔ PHIẾU) - v3.5.1 NEW LOGIC
 * ============================================================
 * 
 * LOGIC MỚI v3.5.1:
 * 
 * 1. CỔ TỨC TIỀN MẶT:
 *    - Tạo khoản THU
 *    - CẬP NHẬT CỘT I (Cổ tức TM): Cộng dồn cổ tức vào cột I
 *    - GHI LỊCH SỬ VÀO CỘT P (Ghi chú): Thêm note về cổ tức
 *    - Cột K (Giá điều chỉnh) tự động tính = (H-I)/E
 * 
 * 2. THƯỞNG CỔ PHIẾU:
 *    - Thêm dòng mới với Loại GD = "Thưởng"
 *    - Giá = 0, Phí = 0, Tổng = 0
 *    - Cột J (Cổ tức CP) = số CP thưởng
 * 
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
      // ============================================================
      // XỬ LÝ CỔ TỨC TIỀN MẶT - v3.5.1 NEW LOGIC
      // ============================================================
      const cashAmount = parseFloat(data.cashAmount);
      const totalDividend = parseFloat(data.totalDividend);
      
      // BƯỚC 1: Tạo giao dịch THU
      const incomeResult = addIncome({
        date: date,
        amount: totalDividend,
        source: 'Đầu tư',
        note: `Cổ tức ${stockCode}: ${formatCurrency(cashAmount)}/CP. ${notes}`
      });
      
      if (!incomeResult.success) {
        return incomeResult;
      }
      
      // BƯỚC 2: CẬP NHẬT CỘT I (Cổ tức TM) VÀ CỘT P (Ghi chú)
      const emptyRow = findEmptyRow(stockSheet, 2);
      const dataRows = emptyRow - 2;
      
      if (dataRows > 0) {
        // Lấy toàn bộ dữ liệu từ sheet (16 cột)
        const stockData = stockSheet.getRange(2, 1, dataRows, 16).getValues();
        
        // Tính tổng số lượng đang nắm giữ và lưu các row MUA
        let totalQuantity = 0;
        const buyRows = [];
        
        for (let i = 0; i < stockData.length; i++) {
          const rowType = stockData[i][2];   // Cột C: Loại GD
          const rowSymbol = stockData[i][3]; // Cột D: Mã CK
          const rowQty = parseFloat(stockData[i][4]) || 0; // Cột E: Số lượng
          
          if (rowSymbol === stockCode) {
            if (rowType === 'Mua') {
              totalQuantity += rowQty;
              buyRows.push({
                row: i + 2, // +2 vì header ở row 1 và array index từ 0
                quantity: rowQty,
                currentDividend: parseFloat(stockData[i][8]) || 0, // Cột I: Cổ tức TM hiện tại
                currentNote: stockData[i][15] || '' // Cột P: Ghi chú hiện tại
              });
            } else if (rowType === 'Bán') {
              totalQuantity -= rowQty;
            } else if (rowType === 'Thưởng') {
              totalQuantity += rowQty;
            }
          }
        }
        
        // Kiểm tra có cổ phiếu hay không
        if (totalQuantity <= 0 || buyRows.length === 0) {
          return {
            success: false,
            message: '❌ Không tìm thấy cổ phiếu MUA để ghi nhận cổ tức!'
          };
        }
        
        // CẬP NHẬT: Cộng cổ tức vào cột I và thêm note vào cột P
        const dateStr = new Date(date).toLocaleDateString('vi-VN');
        
        for (let i = 0; i < buyRows.length; i++) {
          const buyRow = buyRows[i];
          
          // Tính phần cổ tức tương ứng với lô này
          const dividendForThisLot = (buyRow.quantity / totalQuantity) * totalDividend;
          
          // Cột I: Cổ tức TM mới = Cổ tức cũ + Cổ tức lần này
          const newDividend = buyRow.currentDividend + dividendForThisLot;
          stockSheet.getRange(buyRow.row, 9).setValue(newDividend);
          
          // Cột P: Thêm lịch sử cổ tức
          const dividendNote = `CT ${dateStr}: +${formatCurrency(dividendForThisLot)}`;
          const newNote = buyRow.currentNote 
            ? `${buyRow.currentNote} | ${dividendNote}`
            : dividendNote;
          stockSheet.getRange(buyRow.row, 16).setValue(newNote);
          
          Logger.log(`✅ Row ${buyRow.row}: ${stockCode} - Cộng cổ tức ${formatCurrency(dividendForThisLot)} vào cột I`);
        }
        
        Logger.log(`✅ Đã cập nhật cổ tức ${formatCurrency(totalDividend)} cho ${stockCode} vào cột I`);
      }
      
      return {
        success: true,
        message: `✅ Đã ghi nhận cổ tức tiền mặt ${formatCurrency(totalDividend)} cho ${stockCode}!\n` +
                 `📊 Cột "Cổ tức TM" đã được cập nhật.\n` +
                 `💡 Cột "Giá ĐC" tự động tính = (Tổng vốn - Cổ tức TM) / Số lượng`
      };
      
    } else if (type === 'stock') {
      // ============================================================
      // XỬ LÝ THƯỞNG CỔ PHIẾU - v3.5.1
      // ============================================================
      const stockRatio = parseFloat(data.stockRatio);
      const bonusShares = parseInt(data.bonusShares);
      const currentQuantity = parseFloat(data.currentQuantity);
      const newQuantity = currentQuantity + bonusShares;
      
      // Thêm dòng mới: Loại GD = "Thưởng"
      const emptyRow = findEmptyRow(stockSheet, 2);
      const stt = getNextSTT(stockSheet, 2);
      
      const noteText = `Thưởng CP ${stockRatio}% (${bonusShares} CP). ${notes}`;
      
      // ✅ v3.5.1: Ghi đúng 10 cột + ghi chú, sau đó set công thức
      const rowData = [
        stt,            // A: STT
        new Date(date), // B: Ngày
        'Thưởng',       // C: Loại GD
        stockCode,      // D: Mã CK
        bonusShares,    // E: Số lượng
        0,              // F: Giá = 0
        0,              // G: Phí = 0
        0,              // H: Tổng = 0
        0,              // I: Cổ tức TM = 0
        bonusShares     // J: Cổ tức CP = số CP thưởng
      ];
      
      // Ghi dữ liệu vào cột A-J
      stockSheet.getRange(emptyRow, 1, 1, 10).setValues([rowData]);
      
      // Ghi ghi chú vào cột P
      stockSheet.getRange(emptyRow, 16).setValue(noteText);
      
      // ✅ SET CÔNG THỨC cho cột K-O
      stockSheet.getRange(emptyRow, 11).setFormula(`=IF(E${emptyRow}>0, (H${emptyRow}-I${emptyRow})/E${emptyRow}, 0)`);
      stockSheet.getRange(emptyRow, 13).setFormula(`=IF(AND(E${emptyRow}>0, L${emptyRow}>0), E${emptyRow}*L${emptyRow}, 0)`);
      stockSheet.getRange(emptyRow, 14).setFormula(`=IF(M${emptyRow}>0, M${emptyRow}-(H${emptyRow}-I${emptyRow}), 0)`);
      stockSheet.getRange(emptyRow, 15).setFormula(`=IF(AND(N${emptyRow}<>0, (H${emptyRow}-I${emptyRow})>0), N${emptyRow}/(H${emptyRow}-I${emptyRow}), 0)`);
      
      formatNewRow(stockSheet, emptyRow, {
        2: 'dd/mm/yyyy',
        6: '#,##0" VNĐ"',
        7: '#,##0" VNĐ"',
        8: '#,##0" VNĐ"',
        9: '#,##0" VNĐ"',
        11: '#,##0" VNĐ"',
        12: '#,##0" VNĐ"',
        13: '#,##0" VNĐ"',
        14: '#,##0" VNĐ"',
        15: '0.00%'
      });
      
      Logger.log(`✅ Đã ghi nhận thưởng ${bonusShares} CP ${stockCode}`);
      
      return {
        success: true,
        message: `✅ Đã ghi nhận thưởng ${bonusShares} cổ phiếu ${stockCode}!\n` +
                 `📊 Số lượng mới: ${newQuantity} CP\n` +
                 `💡 Giá vốn/CP tự động giảm (vì tổng vốn không đổi, số lượng tăng)`
      };
    }
    
    return {
      success: false,
      message: '❌ Loại cổ tức không hợp lệ!'
    };
    
  } catch (error) {
    Logger.log('Error in processDividend: ' + error.message);
    return {
      success: false,
      message: `❌ Lỗi: ${error.message}`
    };
  }
}