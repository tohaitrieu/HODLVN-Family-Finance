/**
 * ===============================================
 * LENDING HANDLER
 * ===============================================
 */

/**
 * Thêm khoản cho vay mới vào sheet CHO VAY
 * @param {Object} data - Dữ liệu khoản cho vay
 * @returns {Object} {success: boolean, message: string}
 */
function addLending(data) {
  try {
    // Validation
    if (!data.date || !data.borrowerName || !data.principal || !data.interestRate || !data.term) {
      return {
        success: false,
        message: '❌ Vui lòng điền đầy đủ các trường bắt buộc!'
      };
    }
    
    // Parse dữ liệu
    const date = new Date(data.date);
    const borrowerName = data.borrowerName.trim();
    const lendingType = data.lendingType || 'OTHER';
    const principal = parseFloat(data.principal);
    const interestRate = parseFloat(data.interestRate);
    const term = parseInt(data.term);
    const note = data.note || '';
    
    // Validate số liệu
    if (isNaN(principal) || principal <= 0) {
      return {
        success: false,
        message: '❌ Số tiền gốc không hợp lệ!'
      };
    }
    
    if (isNaN(interestRate) || interestRate < 0) {
      return {
        success: false,
        message: '❌ Lãi suất không hợp lệ!'
      };
    }
    
    if (isNaN(term) || term <= 0) {
      return {
        success: false,
        message: '❌ Kỳ hạn không hợp lệ!'
      };
    }
    
    // ============================================
    // DELEGATE TO DataProcessor.addLending() - SINGLE SOURCE OF TRUTH
    // ============================================
    const result = addLending({
      date: date,
      borrowerName: borrowerName,
      lendingType: lendingType,
      principal: principal,
      interestRate: interestRate,
      term: term,
      note: note
    });
    
    if (!result.success) {
      return result;
    }
    
    // ============================================
    // Enhanced result message for UI
    // ============================================
    const resultMessage = `✅ Đã thêm khoản cho vay: ${borrowerName}\n` +
                `💰 Số tiền: ${principal.toLocaleString('vi-VN')}\n` +
                `📅 Kỳ hạn: ${term} tháng\n` +
                `📊 Trạng thái: Đang vay\n✅ Đã tạo khoản chi: Cho vay`;
    
    return {
      success: true,
      message: resultMessage
    };
    
  } catch (error) {
    Logger.log('ERROR in addLending: ' + error.message);
    return {
      success: false,
      message: '❌ Lỗi: ' + error.message
    };
  }
}

/**
 * Lấy danh sách người vay đang nợ
 */
function getLendingList() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(APP_CONFIG.SHEETS.LENDING);
    if (!sheet) return ['Chưa có khoản cho vay'];
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return ['Chưa có khoản cho vay'];
    
    // Col B (Index 1): Name, Col L (Index 11): Status
    const data = sheet.getRange(2, 1, lastRow - 1, 12).getValues();
    
    const list = data
      .filter(row => row[11] === 'Đang vay' && row[1] !== '')
      .map(row => row[1]); // Name
      
    return [...new Set(list)]; // Unique names
    
  } catch (error) {
    return ['Lỗi tải danh sách'];
  }
}

/**
 * Thêm khoản thu hồi nợ
 */
function addLendingPayment(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const lendingSheet = ss.getSheetByName(APP_CONFIG.SHEETS.LENDING);
    const incomeSheet = ss.getSheetByName(APP_CONFIG.SHEETS.INCOME);
    
    if (!lendingSheet || !incomeSheet) {
      return { success: false, message: '❌ Thiếu sheet dữ liệu!' };
    }
    
    const date = new Date(data.date);
    const borrowerName = data.borrowerName;
    const principal = parseFloat(data.principal) || 0;
    const interest = parseFloat(data.interest) || 0;
    const note = data.note || '';
    
    if (principal === 0 && interest === 0) {
      return { success: false, message: '❌ Vui lòng nhập số tiền!' };
    }
    
    // 1. Update Lending Sheet
    // Find the loan by name (Simple logic: Find first active loan with this name)
    // In reality, one person might have multiple loans. Ideally, select specific loan.
    // For now, we just update the first matching record or ask user to be specific in Name.
    // Or we just update the "Gốc đã thu" and "Lãi đã thu" columns.
    
    const lastRow = lendingSheet.getLastRow();
    const lendingData = lendingSheet.getRange(2, 1, lastRow - 1, 12).getValues();
    let foundRow = -1;
    let currentPrincipalPaid = 0;
    let currentInterestPaid = 0;
    
    for (let i = 0; i < lendingData.length; i++) {
      // Col B (1): Name, Col L (11): Status
      if (lendingData[i][1] === borrowerName && lendingData[i][11] === 'Đang vay') {
        foundRow = i + 2;
        currentPrincipalPaid = parseFloat(lendingData[i][8]) || 0; // Col I (Index 8)
        currentInterestPaid = parseFloat(lendingData[i][9]) || 0;  // Col J (Index 9)
        break;
      }
    }
    
    if (foundRow === -1) {
      return { success: false, message: '❌ Không tìm thấy khoản vay của người này!' };
    }
    
    // Update values
    lendingSheet.getRange(foundRow, 9).setValue(currentPrincipalPaid + principal); // I: Gốc đã thu
    lendingSheet.getRange(foundRow, 10).setValue(currentInterestPaid + interest);  // J: Lãi đã thu
    
    // Check if fully paid (Remaining <= 0)
    // Remaining is calculated by formula in Col K (D - I).
    // We can check if Principal Paid >= Principal
    const originalPrincipal = parseFloat(lendingData[foundRow - 2][3]); // Col D (Index 3)
    if (currentPrincipalPaid + principal >= originalPrincipal) {
      lendingSheet.getRange(foundRow, 12).setValue('Đã tất toán'); // L: Status
    }
    
    // 2. Add to Income Sheet
    // Add Principal (Thu hồi nợ)
    if (principal > 0) {
      const incomeEmptyRow = findEmptyRow(incomeSheet, 2);
      const incomeStt = getNextSTT(incomeSheet, 2);
      
      const rowData = [
        incomeStt,
        date,
        principal,
        'Thu hồi nợ',
        `Thu gốc: ${borrowerName}` + (note ? ` - ${note}` : ''),
        Utilities.getUuid() // TransactionID
      ];
      
      incomeSheet.getRange(incomeEmptyRow, 1, 1, rowData.length).setValues([rowData]);
      formatNewRow(incomeSheet, incomeEmptyRow, { 2: 'dd/mm/yyyy', 3: '#,##0' });
    }
    
    // Add Interest (Lãi đầu tư)
    if (interest > 0) {
      const incomeEmptyRow = findEmptyRow(incomeSheet, 2); // Re-find as it might have changed
      const incomeStt = getNextSTT(incomeSheet, 2);
      
      const rowData = [
        incomeStt,
        date,
        interest,
        'Lãi đầu tư',
        `Thu lãi: ${borrowerName}` + (note ? ` - ${note}` : ''),
        Utilities.getUuid() // TransactionID
      ];
      
      incomeSheet.getRange(incomeEmptyRow, 1, 1, rowData.length).setValues([rowData]);
      formatNewRow(incomeSheet, incomeEmptyRow, { 2: 'dd/mm/yyyy', 3: '#,##0' });
    }

    // ============================================
    // BƯỚC 3: GHI VÀO SHEET THU NỢ (LENDING_REPAYMENT)
    // ============================================
    const repaymentSheet = ss.getSheetByName(APP_CONFIG.SHEETS.LENDING_REPAYMENT);
    if (repaymentSheet) {
      const emptyRow = findEmptyRow(repaymentSheet, 2);
      const stt = getNextSTT(repaymentSheet, 2);
      const total = principal + interest;
      const transactionId = Utilities.getUuid();

      // ['STT', 'Ngày', 'Người vay', 'Thu gốc', 'Thu lãi', 'Tổng thu', 'Ghi chú', 'TransactionID']
      const rowData = [
        stt,
        date,
        borrowerName,
        principal,
        interest,
        total,
        note,
        transactionId
      ];

      repaymentSheet.getRange(emptyRow, 1, 1, rowData.length).setValues([rowData]);
      
      formatNewRow(repaymentSheet, emptyRow, {
        2: 'dd/mm/yyyy',
        4: '#,##0',
        5: '#,##0',
        6: '#,##0'
      });
    } else {
      Logger.log('⚠️ Không tìm thấy sheet THU NỢ để ghi nhận giao dịch');
    }
    
    return {
      success: true,
      message: `✅ Đã ghi nhận thu ${principal > 0 ? 'gốc ' + formatCurrency(principal) : ''} ${interest > 0 ? 'lãi ' + formatCurrency(interest) : ''}`
    };
    
  } catch (error) {
    return { success: false, message: '❌ Lỗi: ' + error.message };
  }
}
