/**
 * ===============================================
 * DEBT MANAGEMENT HANDLER v3.3.2 - FINAL FIX
 * ===============================================
 * 
 * CHANGELOG v3.3.2:
 * - Fix lỗi mất công thức cột J khi insert dữ liệu
 * - Chia insert thành 2 phần: A-I và K-L, bỏ qua cột J
 */

/**
 * Thêm khoản nợ mới vào sheet QUẢN LÝ NỢ
 * @param {Object} data - Dữ liệu khoản nợ
 * @returns {Object} {success: boolean, message: string}
 */
function addDebtManagement(data) {
  try {
    // Validation
    if (!data.date || !data.debtName || !data.principal || !data.interestRate || !data.term) {
      return {
        success: false,
        message: '❌ Vui lòng điền đầy đủ các trường bắt buộc!'
      };
    }
    
    // Parse dữ liệu
    const date = new Date(data.date);
    const debtName = data.debtName.trim();
    const debtType = data.debtType || 'OTHER';
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
    // DELEGATE TO addDebt() - SINGLE SOURCE OF TRUTH
    // ============================================
    const result = addDebt({
      loanDate: date,
      debtName: debtName,
      debtType: debtType,
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
    const isInstallmentLoan = ['EQUAL_PRINCIPAL', 'EQUAL_PRINCIPAL_UPFRONT_FEE', 'INTEREST_FREE'].includes(debtType);
    
    const resultMessage = `✅ Đã thêm khoản nợ: ${debtName}\n` +
                `💰 Số tiền: ${principal.toLocaleString('vi-VN')}\n` +
                `📅 Kỳ hạn: ${term} tháng\n` +
                `💳 Loại: ${debtType}\n` +
                `📊 Trạng thái: Chưa trả\n` +
                `✅ Đã tạo khoản thu: Vay ngân hàng` +
                (isInstallmentLoan ? `\n➖ Đã tạo khoản chi: Mua sắm (Trả góp)` : '');
    
    return {
      success: true,
      message: resultMessage
    };
    
  } catch (error) {
    Logger.log('ERROR in addDebtManagement: ' + error.message);
    Logger.log('Stack trace: ' + error.stack);
    return {
      success: false,
      message: '❌ Lỗi: ' + error.message
    };
  }
}

/**
 * ============================================
 * FUNCTION TEST
 * ============================================
 */
function testAddDebtManagement() {
  const testData = {
    date: '2025-10-29',
    debtName: 'Test Margin SSI',
    debtType: 'Margin chứng khoán',
    principal: 25000000,
    interestRate: 9.5,
    term: 3,
    note: 'Test với findEmptyRow() và giữ công thức cột J'
  };
  
  Logger.log('=== BẮT ĐẦU TEST ===');
  const result = addDebtManagement(testData);
  Logger.log('Result: ' + JSON.stringify(result));
  
  if (result.success) {
    SpreadsheetApp.getUi().alert('Test thành công!', result.message, SpreadsheetApp.getUi().ButtonSet.OK);
  } else {
    SpreadsheetApp.getUi().alert('Test thất bại!', result.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * ============================================
 * FUNCTION TEST - Kiểm tra công thức cột J
 * ============================================
 */
function testFormulaColumnJ() {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('QUẢN LÝ NỢ');
  
  if (!sheet) {
    Logger.log('ERROR: Không tìm thấy sheet QUẢN LÝ NỢ');
    return;
  }
  
  // Tìm dòng dữ liệu cuối cùng
  const emptyRow = findEmptyRow(sheet, 2);
  const lastDataRow = emptyRow - 1;
  
  if (lastDataRow < 2) {
    Logger.log('Không có dữ liệu để test');
    return;
  }
  
  Logger.log('=== KIỂM TRA CÔNG THỨC CỘT J ===');
  
  for (let row = 2; row <= lastDataRow; row++) {
    const cellJ = sheet.getRange(row, 10); // Cột J
    const formula = cellJ.getFormula();
    const value = cellJ.getValue();
    
    Logger.log(`Dòng ${row}:`);
    Logger.log(`  - Công thức: ${formula || '(không có)'}`);
    Logger.log(`  - Giá trị: ${value}`);
    
    if (!formula) {
      Logger.log(`  ⚠️ CẢNH BÁO: Dòng ${row} không có công thức!`);
    } else {
      Logger.log(`  ✅ OK`);
    }
  }
}

/**
 * Cập nhật trạng thái khoản nợ sau khi trả nợ
 * @param {string} debtName - Tên khoản nợ
 * @param {number} principal - Số tiền gốc vừa trả
 * @param {number} interest - Số tiền lãi vừa trả
 */
function updateDebtStatus(debtName, principal, interest) {
  try {
    const sheet = getSpreadsheet().getSheetByName(APP_CONFIG.SHEETS.DEBT_MANAGEMENT);
    if (!sheet) return;

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    const data = sheet.getRange(2, 1, lastRow - 1, 12).getValues();
    
    // Find the debt
    for (let i = 0; i < data.length; i++) {
      // Col B (1): Name, Col L (11): Status
      if (data[i][1] === debtName && data[i][11] !== 'Đã thanh toán') { 
        const row = i + 2;
        
        // Update Paid Principal (Col I - index 8)
        const currentPaidPrincipal = parseFloat(data[i][8]) || 0;
        sheet.getRange(row, 9).setValue(currentPaidPrincipal + principal);
        
        // Update Paid Interest (Col J - index 9)
        const currentPaidInterest = parseFloat(data[i][9]) || 0;
        sheet.getRange(row, 10).setValue(currentPaidInterest + interest);
        
        // Check if fully paid
        // Original Principal is Col D (index 3)
        const originalPrincipal = parseFloat(data[i][3]);
        if (currentPaidPrincipal + principal >= originalPrincipal) {
          sheet.getRange(row, 12).setValue('Đã thanh toán');
        } else {
            // If it was "Chưa trả", change to "Đang trả"
            if (data[i][11] === 'Chưa trả') {
                sheet.getRange(row, 12).setValue('Đang trả');
            }
        }
        break; // Update the first matching active debt
      }
    }
  } catch (error) {
    Logger.log('Error updating debt status: ' + error.message);
  }
}

/**
 * Cập nhật trạng thái cho vay sau khi thu nợ
 * @param {string} borrowerName - Tên người vay
 * @param {number} principal - Số tiền gốc vừa thu
 * @param {number} interest - Số tiền lãi vừa thu
 */
function updateLendingStatus(borrowerName, principal, interest) {
  try {
    const sheet = getSpreadsheet().getSheetByName(APP_CONFIG.SHEETS.LENDING);
    if (!sheet) return;

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    const data = sheet.getRange(2, 1, lastRow - 1, 12).getValues();
    
    for (let i = 0; i < data.length; i++) {
      // Col B (1): Name, Col L (11): Status
      if (data[i][1] === borrowerName && data[i][11] !== 'Đã tất toán') {
        const row = i + 2;
        
        // Update Paid Principal (Col I - index 8)
        const currentPaidPrincipal = parseFloat(data[i][8]) || 0;
        sheet.getRange(row, 9).setValue(currentPaidPrincipal + principal);
        
        // Update Paid Interest (Col J - index 9)
        const currentPaidInterest = parseFloat(data[i][9]) || 0;
        sheet.getRange(row, 10).setValue(currentPaidInterest + interest);
        
        // Check if fully paid
        const originalPrincipal = parseFloat(data[i][3]); // Col D (index 3)
        if (currentPaidPrincipal + principal >= originalPrincipal) {
          sheet.getRange(row, 12).setValue('Đã tất toán');
        }
        break;
      }
    }
  } catch (error) {
    Logger.log('Error updating lending status: ' + error.message);
  }
}