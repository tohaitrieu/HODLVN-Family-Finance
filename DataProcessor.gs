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
    
    // Trigger dashboard refresh
    triggerDashboardRefresh();
    
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
    
    // Trigger dashboard refresh
    triggerDashboardRefresh();
    
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
    if (!data.loanDate || !data.debtName || !data.principal) {
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
    
    const principal = parseFloat(data.principal);
    
    const dataObj = {
      stt: stt,
      name: data.debtName,
      type: data.debtType,
      principal: principal,
      rate: (parseFloat(data.interestRate) || 0) / 100,
      term: parseInt(data.term) || 0,
      startDate: new Date(data.loanDate),
      endDate: new Date(data.loanDate).setMonth(new Date(data.loanDate).getMonth() + parseInt(data.term)),
      paidPrincipal: 0,
      paidInterest: 0,
      remaining: '', // Formula
      status: 'Chưa trả',
      note: data.note || '',
      transactionId: Utilities.getUuid()
    };
    
    const rowData = SheetUtils.dataToRow('DEBT_MANAGEMENT', dataObj);
    
    sheet.getRange(emptyRow, 1, 1, rowData.length).setValues([rowData]);
    
    // Set formula for Remaining (Col K - 11)
    // Formula: =IFERROR(Principal(D) - PaidPrincipal(I), 0)
    // R1C1: =IFERROR(RC[-7]-RC[-2], 0)
    sheet.getRange(emptyRow, 11).setFormulaR1C1('=IFERROR(RC[-7]-RC[-2], 0)');
    
    formatNewRow(sheet, emptyRow, {
      2: 'dd/mm/yyyy',
      4: '#,##0',
      5: '0.00%',
      7: 'dd/mm/yyyy',
      8: 'dd/mm/yyyy',
      11: '#,##0'
    });
    
    // STEP 1: ALL debts record INCOME (receiving cash from loan)
    let incomeCategory = 'Vay ngân hàng';
    
    // Map debt type to income category
    if (data.debtType === 'INTEREST_ONLY' || data.debtType === 'BULLET') {
      incomeCategory = 'Vay ngân hàng';
    } else if (data.debtType === 'OTHER') {
      incomeCategory = 'Vay cá nhân';
    } else {
      // Installment loans also use "Vay ngân hàng" or appropriate category
      incomeCategory = 'Vay ngân hàng';
    }
    
    // Verify category exists
    if (!APP_CONFIG.CATEGORIES.INCOME.includes(incomeCategory)) {
      incomeCategory = 'Khác';
    }

    addIncome({
      date: data.loanDate,
      amount: principal,
      source: incomeCategory,
      note: `Giải ngân khoản vay: ${data.debtName}`
    });
    
    // STEP 2: For installment loans, ALSO record EXPENSE (spending cash on purchase)
    const isInstallmentLoan = ['EQUAL_PRINCIPAL', 'EQUAL_PRINCIPAL_UPFRONT_FEE', 'INTEREST_FREE'].includes(data.debtType);
    
    Logger.log(`Debt Type: ${data.debtType}, isInstallmentLoan: ${isInstallmentLoan}`);
    
    if (isInstallmentLoan) {
      Logger.log(`Creating EXPENSE for installment loan: ${data.debtName}`);
      addExpense({
        date: data.loanDate,
        amount: principal,
        category: 'Mua sắm', // Or category based on what was purchased
        subcategory: 'Trả góp',
        note: `Mua trả góp: ${data.debtName}`,
        transactionId: Utilities.getUuid()
      });
    } else {
      Logger.log(`NOT creating EXPENSE for bank loan: ${data.debtName}`);
    }
    
    // Trigger dashboard refresh
    triggerDashboardRefresh();
    
    return { success: true, message: '✅ Đã thêm khoản nợ mới' };
    
  } catch (error) {
    return { success: false, message: 'Lỗi: ' + error.message };
  }
}

/**
 * Thêm khoản trả nợ
 * @param {Object} data - {date, debtName, principal, interest, note}
 * @return {Object} {success, message}
 */
function addDebtPayment(data) {
  try {
    if (!data.date || !data.debtName || (!data.principal && !data.interest)) {
      return {
        success: false,
        message: '❌ Vui lòng nhập ngày, khoản nợ và số tiền!'
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
    
    const emptyRow = findEmptyRow(sheet, 2);
    const stt = getNextSTT(sheet, 2);
    
    const principal = parseFloat(data.principal) || 0;
    const interest = parseFloat(data.interest) || 0;
    
    const dataObj = {
      stt: stt,
      date: new Date(data.date),
      debtName: data.debtName,
      principal: principal,
      interest: interest,
      total: '', // Formula will handle this
      note: data.note || '',
      transactionId: Utilities.getUuid()
    };
    
    const rowData = SheetUtils.dataToRow('DEBT_PAYMENT', dataObj);
    
    sheet.getRange(emptyRow, 1, 1, rowData.length).setValues([rowData]);
    
    // Formula re-application is handled by SheetInitializer/SheetUtils now, 
    // but for specific row formulas that depend on relative references, we might need to set them if not using ArrayFormula.
    // However, SheetUtils.applySheetFormat sets formulas for the whole column range (Row 2 to LastRow).
    // Since we just added a row, we might need to extend the formula or set it for this row.
    // For safety, let's set the formula for this specific row as well, using the logic from Config.
    
    // F: Tổng trả = D + E
    sheet.getRange(emptyRow, 6).setFormula(`=IFERROR(D${emptyRow}+E${emptyRow}, 0)`);
    
    formatNewRow(sheet, emptyRow, {
      2: 'dd/mm/yyyy',
      4: '#,##0',
      5: '#,##0',
      6: '#,##0'
    });
    
    // Update Debt Management status
    updateDebtStatus(data.debtName, principal, interest);
    
    // Add Expense record
    addExpense({
      date: data.date,
      amount: principal + interest,
      category: 'Trả nợ',
      subcategory: data.debtName,
      note: `Trả nợ: ${data.debtName} (Gốc: ${formatCurrency(principal)}, Lãi: ${formatCurrency(interest)})`,
      transactionId: Utilities.getUuid()
    });
    
    return {
      success: true,
      message: `✅ Đã ghi nhận trả nợ: ${data.debtName}\n💰 Tổng: ${formatCurrency(principal + interest)}`
    };
    
  } catch (error) {
    Logger.log('Error in addDebtPayment: ' + error.message);
    return {
      success: false,
      message: `❌ Lỗi: ${error.message}`
    };
  }
}

/**
 * Thêm khoản cho vay mới
 * @param {Object} data
 */
function addLending(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(APP_CONFIG.SHEETS.LENDING);
    
    if (!sheet) return { success: false, message: 'Sheet CHO VAY chưa tạo' };
    
    const emptyRow = findEmptyRow(sheet, 2);
    const stt = getNextSTT(sheet, 2);
    
    const principal = parseFloat(data.principal);
    
    // Calculate end date
    const startDate = new Date(data.date);
    const endDate = new Date(data.date);
    endDate.setMonth(endDate.getMonth() + parseInt(data.term));
    
    const dataObj = {
      stt: stt,
      name: data.borrowerName,
      type: data.lendingType,
      principal: principal,
      rate: (parseFloat(data.interestRate) || 0) / 100,
      term: parseInt(data.term) || 0,
      startDate: startDate,
      endDate: endDate,
      paidPrincipal: 0,
      paidInterest: 0,
      remaining: '', // Formula
      status: 'Đang vay',
      note: data.note || '',
      transactionId: Utilities.getUuid()
    };
    
    const rowData = SheetUtils.dataToRow('LENDING', dataObj);
    
    sheet.getRange(emptyRow, 1, 1, rowData.length).setValues([rowData]);
    
    // Set formula for Remaining (Col K - 11)
    // Formula: =IFERROR(Principal(D) - PaidPrincipal(I), 0)
    // R1C1: =IFERROR(RC[-7]-RC[-2], 0)
    sheet.getRange(emptyRow, 11).setFormulaR1C1('=IFERROR(RC[-7]-RC[-2], 0)');
    
    formatNewRow(sheet, emptyRow, {
      2: 'dd/mm/yyyy', // Date
      4: '#,##0',      // Principal
      5: '0.00%',      // Interest Rate
      7: 'dd/mm/yyyy', // Loan Date
      8: 'dd/mm/yyyy', // Due Date
      9: '#,##0',      // Paid Principal
      10: '#,##0',     // Paid Interest
      11: '#,##0'      // Remaining
    });
    
    // Add Expense record (Money out)
    addExpense({
      date: data.date,
      amount: principal,
      category: 'Đầu tư',
      subcategory: 'Cho vay',
      note: `Cho vay: ${data.borrowerName}`,
      transactionId: Utilities.getUuid()
    });
    
    // Trigger dashboard refresh
    triggerDashboardRefresh();
    
    return { success: true, message: '✅ Đã thêm khoản cho vay mới' };
    
  } catch (error) {
    return { success: false, message: 'Lỗi: ' + error.message };
  }
}

/**
 * Thêm giao dịch thu nợ
 * @param {Object} data
 */
function addLendingRepayment(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(APP_CONFIG.SHEETS.LENDING_REPAYMENT);
    
    if (!sheet) return { success: false, message: 'Sheet THU NỢ chưa tạo' };
    
    const emptyRow = findEmptyRow(sheet, 2);
    const stt = getNextSTT(sheet, 2);
    
    const principal = parseFloat(data.principal) || 0;
    const interest = parseFloat(data.interest) || 0;
    
    const dataObj = {
      stt: stt,
      date: new Date(data.date),
      borrower: data.borrower,
      principal: principal,
      interest: interest,
      total: '', // Formula
      note: data.note || '',
      transactionId: Utilities.getUuid()
    };
    
    const rowData = SheetUtils.dataToRow('LENDING_REPAYMENT', dataObj);
    
    sheet.getRange(emptyRow, 1, 1, rowData.length).setValues([rowData]);
    
    // F: Tổng thu = D + E
    sheet.getRange(emptyRow, 6).setFormula(`=IFERROR(D${emptyRow}+E${emptyRow}, 0)`);
    
    formatNewRow(sheet, emptyRow, {
      2: 'dd/mm/yyyy',
      4: '#,##0',
      5: '#,##0',
      6: '#,##0'
    });
    
    // Update Lending Status
    updateLendingStatus(data.borrower, principal, interest);
    
    // Add Income record
    addIncome({
      date: data.date,
      amount: principal + interest,
      category: 'Thu nợ',
      note: `Thu nợ từ: ${data.borrower} (Gốc: ${formatCurrency(principal)}, Lãi: ${formatCurrency(interest)})`
    });
    
    return { success: true, message: '✅ Đã ghi nhận thu nợ' };
    
  } catch (error) {
    return { success: false, message: 'Lỗi: ' + error.message };
  }
}

// ==================== CHỨNG KHOÁN ====================

/**
 * Thêm giao dịch chứng khoán
 * @param {Object} data - {date, type, ticker, quantity, price, fee, note}
 * @return {Object} {success, message}
 */
function addStock(data) {
  try {
    if (!data.date || !data.type || !data.ticker || !data.quantity || !data.price) {
      return {
        success: false,
        message: '❌ Vui lòng điền đầy đủ thông tin bắt buộc!'
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
    
    const emptyRow = findEmptyRow(sheet, 2);
    const stt = getNextSTT(sheet, 2);
    
    const date = new Date(data.date);
    const type = data.type.toString();
    const ticker = data.ticker.toString().toUpperCase();
    const quantity = parseFloat(data.quantity);
    const price = parseFloat(data.price);
    const fee = parseFloat(data.fee) || 0;
    
    // Calculate Total Cost
    let totalCost = 0;
    if (type === 'Mua') {
      totalCost = (quantity * price) + fee;
    } else if (type === 'Bán') {
      totalCost = (quantity * price) - fee; // Net proceeds
    }
    
    const dataObj = {
      stt: stt,
      date: date,
      type: type,
      ticker: ticker,
      quantity: quantity,
      price: price,
      fee: fee,
      totalCost: totalCost,
      divCash: 0,
      divStock: 0,
      adjPrice: '', // Formula
      marketPrice: '', // Formula
      marketValue: '', // Formula
      profit: '', // Formula
      profitPercent: '', // Formula
      note: data.note || ''
    };
    
    const rowData = SheetUtils.dataToRow('STOCK', dataObj);
    
    sheet.getRange(emptyRow, 1, 1, rowData.length).setValues([rowData]);
    
    // Set Formulas
    // K: Giá ĐC = (Tổng vốn - Cổ tức TM) / (Số lượng + Cổ tức CP) -> Simplified: (H-I)/(E+J) ? No, formula in Config is: (H-I)/E (Assuming J is 0 for now or handled differently)
    // Config Formula: =IF(RC[-6]>0, (RC[-3]-RC[-2])/RC[-6], 0) -> (TotalCost - DivCash) / Quantity
    sheet.getRange(emptyRow, 11).setFormula(`=IF(E${emptyRow}>0, (H${emptyRow}-I${emptyRow})/E${emptyRow}, 0)`);
    
    // L: Giá HT = MPRICE(Ticker)
    sheet.getRange(emptyRow, 12).setFormula(`=IF(D${emptyRow}<>"", MPRICE(D${emptyRow}), 0)`);
    
    // M: Giá trị HT = Quantity * MarketPrice
    sheet.getRange(emptyRow, 13).setFormula(`=IF(AND(E${emptyRow}>0, L${emptyRow}>0), E${emptyRow}*L${emptyRow}, 0)`);
    
    // N: Lãi/Lỗ = MarketValue - (TotalCost - DivCash)
    sheet.getRange(emptyRow, 14).setFormula(`=IF(M${emptyRow}>0, M${emptyRow}-(H${emptyRow}-I${emptyRow}), 0)`);
    
    // O: % Lãi/Lỗ
    sheet.getRange(emptyRow, 15).setFormula(`=IF(AND(N${emptyRow}<>0, (H${emptyRow}-I${emptyRow})>0), N${emptyRow}/(H${emptyRow}-I${emptyRow}), 0)`);
    
    formatNewRow(sheet, emptyRow, {
      2: 'dd/mm/yyyy',
      6: '#,##0',
      7: '#,##0',
      8: '#,##0',
      11: '#,##0',
      12: '#,##0',
      13: '#,##0',
      14: '#,##0',
      15: '0.00%'
    });
    
    BudgetManager.updateInvestmentBudget('Chứng khoán', totalCost);
    
    // Auto-add Expense for Buy
    if (type === 'Mua') {
      addExpense({
        date: date,
        amount: totalCost,
        category: 'Đầu tư',
        subcategory: `Mua CK: ${ticker}`,
        note: `Mua ${quantity} ${ticker} giá ${formatCurrency(price)}`,
        transactionId: Utilities.getUuid()
      });
    }
    
    return {
      success: true,
      message: `✅ Đã ghi nhận ${type} ${quantity} ${ticker}!\n` +
               `💰 Giá: ${formatCurrency(price)}\n` +
               `💵 Tổng: ${formatCurrency(totalCost)}`
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
    const quantity = parseFloat(data.quantity);
    const price = parseFloat(data.price);
    const total = quantity * price;
    
    const dataObj = {
      stt: stt,
      date: date,
      assetName: 'GOLD',
      type: data.type,
      goldType: data.goldType,
      quantity: quantity,
      unit: data.unit || 'Lượng',
      price: price,
      totalCost: total,
      marketPrice: '', // Formula
      marketValue: '', // Formula
      profit: '', // Formula
      profitPercent: '', // Formula
      note: data.note || ''
    };
    
    const rowData = SheetUtils.dataToRow('GOLD', dataObj);
    
    sheet.getRange(emptyRow, 1, 1, rowData.length).setValues([rowData]);
    
    // Set Formulas
    // J: Giá HT = GPRICE(Loại vàng - Cột E)
    sheet.getRange(emptyRow, 10).setFormula(`=IF(E${emptyRow}<>"", GPRICE(E${emptyRow}), 0)`);
    
    // K: Giá trị HT = Số lượng * Giá HT
    sheet.getRange(emptyRow, 11).setFormula(`=IF(AND(F${emptyRow}>0, J${emptyRow}>0), F${emptyRow}*J${emptyRow}, 0)`);
    
    // L: Lãi/Lỗ = Giá trị HT - Tổng vốn
    sheet.getRange(emptyRow, 12).setFormula(`=IF(K${emptyRow}>0, K${emptyRow}-I${emptyRow}, 0)`);
    
    // M: % Lãi/Lỗ
    sheet.getRange(emptyRow, 13).setFormula(`=IF(I${emptyRow}>0, L${emptyRow}/I${emptyRow}, 0)`);
    
    formatNewRow(sheet, emptyRow, {
      2: 'dd/mm/yyyy',
      8: '#,##0',
      9: '#,##0',
      10: '#,##0',
      11: '#,##0',
      12: '#,##0',
      13: '0.00%'
    });
    
    BudgetManager.updateInvestmentBudget('Vàng', total);
    
    if (data.type === 'Mua') {
      addExpense({
        date: date,
        amount: total,
        category: 'Đầu tư',
        subcategory: `Mua Vàng: ${data.goldType}`,
        note: `Mua ${quantity} ${data.unit} ${data.goldType} giá ${formatCurrency(price)}`,
        transactionId: Utilities.getUuid()
      });
    }
    
    // Trigger dashboard refresh
    triggerDashboardRefresh();
    
    return {
      success: true,
      message: `✅ Đã ghi nhận ${data.type} ${quantity} ${data.unit} ${data.goldType}!`
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
    const quantity = parseFloat(data.quantity);
    const priceUSD = parseFloat(data.priceUSD);
    const rate = parseFloat(data.exchangeRate) || 25300;
    const fee = parseFloat(data.fee) || 0;
    
    const totalUSD = (quantity * priceUSD) + fee;
    const totalVND = totalUSD * rate;
    const priceVND = priceUSD * rate;
    
    const dataObj = {
      stt: stt,
      date: date,
      type: data.type,
      coin: data.coin.toString().toUpperCase(),
      quantity: quantity,
      priceUSD: priceUSD,
      rate: rate,
      priceVND: priceVND,
      totalCost: totalVND,
      marketPriceUSD: '', // Formula
      marketValueUSD: '', // Formula
      marketPriceVND: '', // Formula
      marketValueVND: '', // Formula
      profit: '', // Formula
      profitPercent: '', // Formula
      exchange: data.exchange || '',
      wallet: data.wallet || '',
      note: data.note || ''
    };
    
    const rowData = SheetUtils.dataToRow('CRYPTO', dataObj);
    
    sheet.getRange(emptyRow, 1, 1, rowData.length).setValues([rowData]);
    
    // Set Formulas
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
      6: '#,##0.00',
      7: '#,##0',
      8: '#,##0',
      9: '#,##0',
      10: '#,##0.00',
      11: '#,##0.00',
      12: '#,##0',
      13: '#,##0',
      14: '#,##0',
      15: '0.00%'
    });
    
    BudgetManager.updateInvestmentBudget('Crypto', totalVND);
    
    if (data.type === 'Mua') {
      addExpense({
        date: date,
        amount: totalVND,
        category: 'Đầu tư',
        subcategory: `Mua Crypto: ${data.coin}`,
        note: `Mua ${quantity} ${data.coin} giá $${formatCurrency(priceUSD)}`,
        transactionId: Utilities.getUuid()
      });
    }
    
    // Trigger dashboard refresh
    triggerDashboardRefresh();
    
    return {
      success: true,
      message: `✅ Đã ghi nhận ${data.type} ${quantity} ${data.coin}!`
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
 * Thêm đầu tư khác
 * @param {Object} data - {date, investmentType, amount, roi, term, note}
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
    
    const amount = parseFloat(data.amount);
    const roi = parseFloat(data.roi) || 0;
    const term = parseFloat(data.term) || 0;
    
    // Calculate Expected Return
    const interest = amount * (roi / 100) * (term / 12);
    const expectedReturn = amount + interest;
    
    const dataObj = {
      stt: stt,
      date: new Date(data.date),
      type: data.investmentType,
      amount: amount,
      rate: roi / 100,
      term: term,
      expectedReturn: expectedReturn,
      note: data.note || ''
    };
    
    const rowData = SheetUtils.dataToRow('OTHER_INVESTMENT', dataObj);
    
    sheet.getRange(emptyRow, 1, 1, rowData.length).setValues([rowData]);
    
    formatNewRow(sheet, emptyRow, {
      2: 'dd/mm/yyyy',
      4: '#,##0',
      5: '0.00%',
      6: '0',
      7: '#,##0'
    });
    
    BudgetManager.updateInvestmentBudget('Đầu tư khác', amount);
    
    addExpense({
      date: data.date,
      amount: amount,
      category: 'Đầu tư',
      subcategory: `Đầu tư khác: ${data.investmentType}`,
      note: `Đầu tư ${data.investmentType}: ${formatCurrency(amount)}`,
      transactionId: Utilities.getUuid()
    });
    
    // Trigger dashboard refresh
    triggerDashboardRefresh();
    
    return {
      success: true,
      message: `✅ Đã thêm đầu tư: ${data.investmentType} - ${formatCurrency(amount)}`
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
      
      // Trigger dashboard refresh
      triggerDashboardRefresh();
      
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
      
      // Trigger dashboard refresh
      triggerDashboardRefresh();
      
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