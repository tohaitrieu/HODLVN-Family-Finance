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
      note,
      data.transactionId || Utilities.getUuid() // Col F: TransactionID
    ];
    
    sheet.getRange(emptyRow, 1, 1, rowData.length).setValues([rowData]);
    
    formatNewRow(sheet, emptyRow, {
      2: 'dd/mm/yyyy',
      3: '#,##0'
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
    const subcategory = data.subcategory || data.detail || '';
    const note = data.note || '';
    
    const rowData = [
      stt,
      date,
      amount,
      category,
      subcategory,
      note,
      data.transactionId || Utilities.getUuid() // Col G: TransactionID
    ];
    
    sheet.getRange(emptyRow, 1, 1, rowData.length).setValues([rowData]);
    
    formatNewRow(sheet, emptyRow, {
      2: 'dd/mm/yyyy',
      3: '#,##0'
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
    const debtType = data.debtType || (debtName.toLowerCase().includes('margin') ? 'Margin chứng khoán' : 'Khác');
    const amount = parseFloat(data.amount);
    const interestRate = parseFloat(data.interestRate) || 0;
    const term = parseInt(data.term) || 12;
    const note = data.note || data.purpose || '';
    
    const dueDate = new Date(loanDate);
    dueDate.setMonth(dueDate.getMonth() + term);
    
    // Phần 1: Cột A-J (STT đến Đã trả lãi) - 10 cột
    const rowDataPart1 = [
      stt,                    // A: STT
      debtName,               // B: Tên khoản nợ
      debtType,               // C: Loại hình (NEW)
      amount,                 // D: Gốc
      interestRate / 100,     // E: Lãi suất
      term,                   // F: Kỳ hạn
      loanDate,               // G: Ngày vay
      dueDate,                // H: Đáo hạn
      0,                      // I: Đã trả gốc
      0                       // J: Đã trả lãi
    ];
    
    const transactionId = Utilities.getUuid();
    
    // Phần 2: Cột L-M (Trạng thái và Ghi chú) - 2 cột
    // Col N: TransactionID
    const rowDataPart2 = [
      'Chưa trả',             // L: Trạng thái
      note,                   // M: Ghi chú
      transactionId           // N: TransactionID
    ];
    
    // Insert Phần 1
    sheet.getRange(emptyRow, 1, 1, rowDataPart1.length).setValues([rowDataPart1]);
    
    // Insert Phần 2 (Bỏ qua cột K)
    sheet.getRange(emptyRow, 12, 1, rowDataPart2.length).setValues([rowDataPart2]);
    
    formatNewRow(sheet, emptyRow, {
      4: '#,##0',           // D: Gốc
      5: '0.00"%"',         // E: Lãi suất
      7: 'dd/mm/yyyy',      // G: Ngày vay
      8: 'dd/mm/yyyy',      // H: Đáo hạn
      9: '#,##0',           // I: Đã trả gốc
      10: '#,##0',          // J: Đã trả lãi
      11: '#,##0'           // K: Còn nợ
    });
    
    let incomeSource = 'Khác';
    const nameLower = debtName.toLowerCase();
    const typeLower = (data.debtType || '').toLowerCase();
    
    if (nameLower.includes('margin') || typeLower.includes('margin') || typeLower.includes('ngân hàng')) {
      incomeSource = 'Vay ngân hàng';
    } else if (typeLower.includes('cá nhân')) {
      incomeSource = 'Vay cá nhân';
    }

    const incomeResult = addIncome({
      date: loanDate,
      amount: amount,
      source: incomeSource,
      note: `Vay ${debtName}. ${note}`,
      transactionId: transactionId // Link ID
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
function addDebtPayment(data) {
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
    
    // 1. Find Parent Debt & ID
    let parentId = '';
    const debtData = debtSheet.getRange(2, 1, debtSheet.getLastRow() - 1, 14).getValues(); // Read up to Col N (14)
    
    for (let i = 0; i < debtData.length; i++) {
      if (debtData[i][1] === debtName) { // Col B: Name
        parentId = debtData[i][13]; // Col N: TransactionID (Index 13)
        break;
      }
    }
    
    if (!parentId) {
      // Fallback if no ID found (old data): Generate one on the fly based on name/date? 
      // Or just use UUID. Better to use UUID fallback to avoid collision if logic fails.
      parentId = Utilities.getUuid(); 
    }
    
    // 2. Count existing payments for this Parent ID to generate Suffix
    // Read Debt Payment sheet to count
    const paymentData = paymentSheet.getRange(2, 8, paymentSheet.getLastRow() - 1, 1).getValues(); // Col H: TransactionID
    let count = 0;
    paymentData.forEach(row => {
      if (row[0] && row[0].toString().startsWith(parentId)) {
        count++;
      }
    });
    
    // 3. Generate New ID
    const transactionId = IDGenerator.generateSuffix(parentId, count);

    const rowData = [
      stt,
      date,
      debtName,
      principalAmount,
      interestAmount,
      totalAmount,
      note,
      transactionId // Col H: TransactionID
    ];
    
    paymentSheet.getRange(emptyRow, 1, 1, rowData.length).setValues([rowData]);
    
    formatNewRow(paymentSheet, emptyRow, {
      2: 'dd/mm/yyyy',
      4: '#,##0',
      5: '#,##0',
      6: '#,##0'
    });
    
    // Update Debt Sheet
    const debtEmptyRow = findEmptyRow(debtSheet, 2);
    const debtDataRows = debtEmptyRow - 2;
    
    if (debtDataRows > 0) {
      // Read Col B (Name) to Col L (Status)
      // B(1), C(2), D(3), E(4), F(5), G(6), H(7), I(8), J(9), K(10), L(11)
      // Indices in values: 0=B, ..., 10=L
      // Re-read to be safe or use previous read
      
      for (let i = 0; i < debtData.length; i++) {
        const rowDebtName = debtData[i][1]; // Col B
        
        if (rowDebtName === debtName) {
          const row = i + 2;
          
          // Get current Paid Principal (Col I) & Interest (Col J)
          const paidPrinCell = debtSheet.getRange(row, 9); // Col I
          const paidIntCell = debtSheet.getRange(row, 10); // Col J
          
          const currentPaidPrin = paidPrinCell.getValue() || 0;
          const currentPaidInt = paidIntCell.getValue() || 0;
          
          paidPrinCell.setValue(currentPaidPrin + principalAmount);
          paidIntCell.setValue(currentPaidInt + interestAmount);
          
          // Check Remaining (Col K - 11)
          // Remaining is calculated by formula: D - I
          // We can check if Paid Principal >= Original Principal (Col D)
          const originalPrincipal = parseFloat(debtData[i][3]); // Col D (Index 3 in range A-N)
          
          if (currentPaidPrin + principalAmount >= originalPrincipal) {
            debtSheet.getRange(row, 12).setValue('Đã thanh toán'); // Col L
          }
          
          break;
        }
      }
    }
    
    // [NEW] Add to Expense Sheet (Sync)
    const expenseResult = addExpense({
      date: date,
      amount: totalAmount,
      category: 'Trả nợ',
      subcategory: `Trả nợ: ${debtName}`,
      note: note,
      transactionId: transactionId // Link ID
    });
    
    if (!expenseResult.success) {
      Logger.log('Cảnh báo: Không thể tự động thêm chi phí cho khoản trả nợ');
    } else {
      Logger.log('✅ Đã tự động thêm chi phí: Trả nợ ' + debtName);
    }
    
    BudgetManager.updateDebtBudget();
    
    Logger.log(`Đã trả nợ: ${debtName} - Gốc: ${formatCurrency(principalAmount)}, Lãi: ${formatCurrency(interestAmount)} (ID: ${transactionId})`);
    
    return {
      success: true,
      message: `✅ Đã ghi nhận trả nợ ${debtName}!\n` +
               `💰 Gốc: ${formatCurrency(principalAmount)}\n` +
               `📊 Lãi: ${formatCurrency(interestAmount)}\n` +
               `💵 Tổng: ${formatCurrency(totalAmount)}\n` +
               `🆔 ID: ${transactionId}`
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
      6: '#,##0',
      7: '#,##0',
      8: '#,##0',
      9: '#,##0',
      11: '#,##0',
      12: '#,##0',
      13: '#,##0',
      14: '#,##0',
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
    
    // [NEW] Structure: 
    // A: STT, B: Ngày, C: Tài sản (GOLD), D: Loại GD, E: Loại vàng, F: Số lượng, G: Đơn vị, 
    // H: Giá vốn, I: Tổng vốn, J-M: Formulas, N: Ghi chú
    
    const rowData = [
      stt,
      date,
      'GOLD',
      type,
      goldType,
      quantity,
      unit,
      price,
      total
    ];
    
    // Write A-I (9 columns)
    sheet.getRange(emptyRow, 1, 1, 9).setValues([rowData]);
    
    // Write Note to N (Column 14)
    sheet.getRange(emptyRow, 14).setValue(note);
    
    // Set Formulas for J-M
    // J: Giá HT = GPRICE(Tài sản - Cột C)
    sheet.getRange(emptyRow, 10).setFormula(`=IF(C${emptyRow}<>"", GPRICE(C${emptyRow}), 0)`);
    
    // K: Giá trị HT = Số lượng * Giá HT
    sheet.getRange(emptyRow, 11).setFormula(`=IF(AND(F${emptyRow}>0, J${emptyRow}>0), F${emptyRow}*J${emptyRow}, 0)`);
    
    // L: Lãi/Lỗ = Giá trị HT - Tổng vốn
    sheet.getRange(emptyRow, 12).setFormula(`=IF(K${emptyRow}>0, K${emptyRow}-I${emptyRow}, 0)`);
    
    // M: % Lãi/Lỗ
    sheet.getRange(emptyRow, 13).setFormula(`=IF(I${emptyRow}>0, L${emptyRow}/I${emptyRow}, 0)`);
    
    formatNewRow(sheet, emptyRow, {
      2: 'dd/mm/yyyy',
      8: '#,##0', // Giá vốn
      9: '#,##0', // Tổng vốn
      10: '#,##0', // Giá HT
      11: '#,##0', // Giá trị HT
      12: '#,##0', // Lãi/Lỗ
      13: '0.00%'  // % Lãi/Lỗ
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
/**
 * Thêm giao dịch crypto
 * @param {Object} data - {date, type, coin, quantity, price, fee, note}
 * @return {Object} {success, message}
 */
function addCrypto(data) {
  try {
    if (!data.date || !data.type || !data.coin || !data.quantity || !data.priceUSD) {
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
    const priceUSD = parseFloat(data.priceUSD); // Corrected key from form
    
    const rate = parseFloat(data.exchangeRate) || 25300; // Corrected key from form
    
    const priceVND = priceUSD * rate;
    const fee = parseFloat(data.fee) || 0; 
    
    const totalUSD = (quantity * priceUSD) + fee;
    const totalVND = totalUSD * rate;
    
    const note = data.note || '';
    const san = data.exchange || ''; // Corrected key from form
    const vi = data.wallet || '';   // Corrected key from form
    
    // [NEW] Structure:
    // A: STT, B: Ngày, C: Loại GD, D: Coin, E: Số lượng, F: Giá (USD), G: Tỷ giá, H: Giá (VND), I: Tổng vốn
    // J-O: Formulas
    // P: Sàn, Q: Ví, R: Ghi chú
    
    const rowData = [
      stt,
      date,
      type,
      coin,
      quantity,
      priceUSD,
      rate,
      priceVND,
      totalVND
    ];
    
    // Write A-I (9 columns)
    sheet.getRange(emptyRow, 1, 1, 9).setValues([rowData]);
    
    // Write P-R (3 columns)
    sheet.getRange(emptyRow, 16, 1, 3).setValues([[san, vi, note]]);
    
    // Set Formulas for J-O
    // J: Giá HT (USD)
    sheet.getRange(emptyRow, 10).setFormula(`=IF(D${emptyRow}<>"", CPRICE(D${emptyRow}&"USD"), 0)`);
    
    // K: Giá trị HT (USD)
    sheet.getRange(emptyRow, 11).setFormula(`=IF(AND(E${emptyRow}>0, J${emptyRow}>0), E${emptyRow}*J${emptyRow}, 0)`);
    
    // L: Giá HT (VND)
    sheet.getRange(emptyRow, 12).setFormula(`=IF(AND(J${emptyRow}>0, G${emptyRow}>0), J${emptyRow}*G${emptyRow}, 0)`);
    
    // M: Giá trị HT (VND)
    sheet.getRange(emptyRow, 13).setFormula(`=IF(AND(K${emptyRow}>0, G${emptyRow}>0), K${emptyRow}*G${emptyRow}, 0)`);
    
    // N: Lãi/Lỗ
    sheet.getRange(emptyRow, 14).setFormula(`=IF(M${emptyRow}>0, M${emptyRow}-I${emptyRow}, 0)`);
    
    // O: % Lãi/Lỗ
    sheet.getRange(emptyRow, 15).setFormula(`=IF(I${emptyRow}>0, N${emptyRow}/I${emptyRow}, 0)`);
    
    formatNewRow(sheet, emptyRow, {
      2: 'dd/mm/yyyy',
      6: '#,##0.00', // Giá USD
      7: '#,##0',    // Tỷ giá
      8: '#,##0',    // Giá VND
      9: '#,##0',    // Tổng vốn
      10: '#,##0.00', // Giá HT USD
      11: '#,##0.00', // Giá trị HT USD
      12: '#,##0',    // Giá HT VND
      13: '#,##0',    // Giá trị HT VND
      14: '#,##0',    // Lãi/Lỗ
      15: '0.00%'     // % Lãi/Lỗ
    });
    
    BudgetManager.updateInvestmentBudget('Crypto', totalVND);
    
    Logger.log(`Đã thêm giao dịch crypto: ${type} ${quantity} ${coin}`);
    
    return {
      success: true,
      message: `✅ Đã ghi nhận ${type} ${quantity} ${coin}!\n` +
               `💰 Giá: $${formatCurrency(priceUSD)}\n` +
               `💵 Tổng: ${formatCurrency(totalVND)}`
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
      4: '#,##0'
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
        6: '#,##0',
        7: '#,##0',
        8: '#,##0',
        9: '#,##0',
        11: '#,##0',
        12: '#,##0',
        13: '#,##0',
        14: '#,##0',
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