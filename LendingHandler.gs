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
    const lendingType = data.lendingType || 'Khác';
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
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // ============================================
    // BƯỚC 1: THÊM VÀO SHEET CHO VAY
    // ============================================
    const lendingSheet = ss.getSheetByName(APP_CONFIG.SHEETS.LENDING);
    
    if (!lendingSheet) {
      return {
        success: false,
        message: '❌ Không tìm thấy sheet CHO VAY. Vui lòng khởi tạo sheet trước!'
      };
    }
    
    // Tính toán
    const maturityDate = new Date(date);
    maturityDate.setMonth(maturityDate.getMonth() + term);
    
    // Tìm dòng trống
    const emptyRow = findEmptyRow(lendingSheet, 2);
    const stt = getNextSTT(lendingSheet, 2);
    
    // Phần 1: Cột A-J (STT đến Lãi đã thu) - 10 cột
    // 'STT', 'Tên người vay', 'Loại hình', 'Số tiền gốc', 'Lãi suất (%/năm)', 
    // 'Kỳ hạn (tháng)', 'Ngày vay', 'Ngày đến hạn', 'Gốc đã thu', 'Lãi đã thu'
    const rowDataPart1 = [
      stt,                    // A: STT
      borrowerName,           // B: Tên người vay
      lendingType,            // C: Loại hình (NEW)
      principal,              // D: Gốc
      interestRate / 100,     // E: Lãi suất
      term,                   // F: Kỳ hạn
      date,                   // G: Ngày vay
      maturityDate,           // H: Đáo hạn
      0,                      // I: Gốc đã thu
      0                       // J: Lãi đã thu
    ];
    
    // Phần 2: Cột L-M (Trạng thái và Ghi chú) - 2 cột
    const rowDataPart2 = [
      'Đang vay',             // L: Trạng thái
      note                    // M: Ghi chú
    ];
    
    // Insert Phần 1
    lendingSheet.getRange(emptyRow, 1, 1, rowDataPart1.length).setValues([rowDataPart1]);
    
    // Insert Phần 2 (Bỏ qua cột K - Còn lại)
    lendingSheet.getRange(emptyRow, 12, 1, rowDataPart2.length).setValues([rowDataPart2]);
    
    // Format
    formatNewRow(lendingSheet, emptyRow, {
      4: '#,##0',           // D: Gốc
      5: '0.00"%"',         // E: Lãi suất
      7: 'dd/mm/yyyy',      // G: Ngày vay
      8: 'dd/mm/yyyy',      // H: Đáo hạn
      9: '#,##0',           // I: Gốc đã thu
      10: '#,##0',          // J: Lãi đã thu
      11: '#,##0'           // K: Còn lại
    });
    
    // ============================================
    // BƯỚC 2: TỰ ĐỘNG THÊM KHOẢN CHI
    // ============================================
    let autoExpenseMessage = '';
    
    const expenseSheet = ss.getSheetByName(APP_CONFIG.SHEETS.EXPENSE);
    
    if (!expenseSheet) {
      autoExpenseMessage = '\n⚠️ Không tìm thấy sheet CHI. Không thể tự động thêm khoản chi!';
    } else {
      const expenseEmptyRow = findEmptyRow(expenseSheet, 2);
      const expenseStt = getNextSTT(expenseSheet, 2);
      
      // Columns: STT | Ngày | Số tiền | Danh mục | Chi tiết | Ghi chú
      const expenseRowData = [
        expenseStt,
        date,
        principal,
        'Cho vay',
        `Cho vay: ${borrowerName}`,
        `Loại: ${lendingType}`
      ];
      
      expenseSheet.getRange(expenseEmptyRow, 1, 1, expenseRowData.length).setValues([expenseRowData]);
      
      // Format
      formatNewRow(expenseSheet, expenseEmptyRow, {
        2: 'dd/mm/yyyy',
        3: '#,##0'
      });
      
      autoExpenseMessage = `\n✅ Đã TỰ ĐỘNG thêm khoản chi "Cho vay" vào sheet CHI`;
    }
    
    // ============================================
    // BƯỚC 3: TRẢ VỀ KẾT QUẢ
    // ============================================
    const resultMessage = `✅ Đã thêm khoản cho vay: ${borrowerName}\n` +
               `💰 Số tiền: ${principal.toLocaleString('vi-VN')}\n` +
               `📅 Kỳ hạn: ${term} tháng\n` +
               `📊 Trạng thái: Đang vay` +
               autoExpenseMessage;
    
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
    
    // Col B (Index 1): Name, Col K (Index 10): Status
    const data = sheet.getRange(2, 1, lastRow - 1, 11).getValues();
    
    const list = data
      .filter(row => row[10] === 'Đang vay' && row[1] !== '')
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
    const lendingData = lendingSheet.getRange(2, 1, lastRow - 1, 11).getValues();
    let foundRow = -1;
    let currentPrincipalPaid = 0;
    let currentInterestPaid = 0;
    
    for (let i = 0; i < lendingData.length; i++) {
      if (lendingData[i][1] === borrowerName && lendingData[i][10] === 'Đang vay') {
        foundRow = i + 2;
        currentPrincipalPaid = parseFloat(lendingData[i][7]) || 0;
        currentInterestPaid = parseFloat(lendingData[i][8]) || 0;
        break;
      }
    }
    
    if (foundRow === -1) {
      return { success: false, message: '❌ Không tìm thấy khoản vay của người này!' };
    }
    
    // Update values
    lendingSheet.getRange(foundRow, 8).setValue(currentPrincipalPaid + principal); // H: Gốc đã thu
    lendingSheet.getRange(foundRow, 9).setValue(currentInterestPaid + interest);   // I: Lãi đã thu
    
    // Check if fully paid (Remaining <= 0)
    // Remaining is calculated by formula in Col J (C - H).
    // We can check if Principal Paid >= Principal
    const originalPrincipal = parseFloat(lendingData[foundRow - 2][2]);
    if (currentPrincipalPaid + principal >= originalPrincipal) {
      lendingSheet.getRange(foundRow, 11).setValue('Đã tất toán'); // K: Status
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
        `Thu gốc: ${borrowerName}` + (note ? ` - ${note}` : '')
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
        `Thu lãi: ${borrowerName}` + (note ? ` - ${note}` : '')
      ];
      
      incomeSheet.getRange(incomeEmptyRow, 1, 1, rowData.length).setValues([rowData]);
      formatNewRow(incomeSheet, incomeEmptyRow, { 2: 'dd/mm/yyyy', 3: '#,##0' });
    }
    
    return {
      success: true,
      message: `✅ Đã ghi nhận thu ${principal > 0 ? 'gốc ' + formatCurrency(principal) : ''} ${interest > 0 ? 'lãi ' + formatCurrency(interest) : ''}`
    };
    
  } catch (error) {
    return { success: false, message: '❌ Lỗi: ' + error.message };
  }
}
