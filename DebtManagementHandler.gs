/**
 * ===============================================
 * DEBT MANAGEMENT HANDLER v2.0
 * ===============================================
 * 
 * Function xử lý thêm khoản nợ mới vào QUẢN LÝ NỢ
 * TỰ ĐỘNG thêm khoản thu nếu tên nợ có từ khóa "Margin"
 * 
 * FIXED: Trực tiếp thêm vào sheet THU thay vì gọi function
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
    const debtType = data.debtType || 'Khác';
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
    // BƯỚC 1: THÊM VÀO SHEET QUẢN LÝ NỢ
    // ============================================
    const debtSheet = ss.getSheetByName('QUẢN LÝ NỢ');
    
    if (!debtSheet) {
      return {
        success: false,
        message: '❌ Không tìm thấy sheet QUẢN LÝ NỢ. Vui lòng khởi tạo sheet trước!'
      };
    }
    
    // Tính toán
    const maturityDate = new Date(date);
    maturityDate.setMonth(maturityDate.getMonth() + term);
    
    // Tìm dòng trống để thêm (bỏ qua header)
    let lastRow = debtSheet.getLastRow();
    let emptyRow = lastRow + 1;
    
    // Nếu có dữ liệu, tìm dòng trống thực sự
    if (lastRow > 1) {
      const dataRange = debtSheet.getRange(2, 2, lastRow - 1, 1).getValues();
      for (let i = 0; i < dataRange.length; i++) {
        if (!dataRange[i][0]) {
          emptyRow = i + 2;
          break;
        }
      }
    }
    
    // STT tự động
    const stt = emptyRow - 1;
    
    // Thêm dữ liệu vào sheet QUẢN LÝ NỢ
    // Columns: STT | Tên | Gốc | Lãi suất | Kỳ hạn | Ngày vay | Đáo hạn | Đã trả gốc | Đã trả lãi | Còn nợ | Trạng thái | Loại | Ghi chú
    debtSheet.getRange(emptyRow, 1, 1, 13).setValues([[
      stt,                    // A: STT
      debtName,               // B: Tên khoản nợ
      principal,              // C: Gốc
      interestRate,           // D: Lãi suất
      term,                   // E: Kỳ hạn
      date,                   // F: Ngày vay
      maturityDate,           // G: Đáo hạn
      0,                      // H: Đã trả gốc
      0,                      // I: Đã trả lãi
      principal,              // J: Còn nợ (ban đầu = gốc)
      'Đang trả',             // K: Trạng thái (phải khớp với validation!)
      debtType,               // L: Loại nợ
      note                    // M: Ghi chú
    ]]);
    
    // Format
    debtSheet.getRange(emptyRow, 3, 1, 1).setNumberFormat('#,##0'); // Gốc
    debtSheet.getRange(emptyRow, 4, 1, 1).setNumberFormat('0.00"%"'); // Lãi suất
    debtSheet.getRange(emptyRow, 6, 1, 2).setNumberFormat('dd/mm/yyyy'); // Ngày vay & Đáo hạn
    debtSheet.getRange(emptyRow, 8, 1, 3).setNumberFormat('#,##0'); // Đã trả gốc, lãi, còn nợ
    
    // ============================================
    // BƯỚC 2: TỰ ĐỘNG THÊM KHOẢN THU CHO MỌI KHOẢN NỢ
    // ============================================
    // Logic: Khi vay nợ = tiền mặt tăng = phải ghi nhận thu
    let autoIncomeMessage = '';
    
    Logger.log('Bắt đầu thêm khoản thu tự động cho khoản vay...');
    
    // ============================================
    // BƯỚC 3: TỰ ĐỘNG THÊM VÀO SHEET THU
    // ============================================
    // ============================================
    // BƯỚC 3: TỰ ĐỘNG THÊM VÀO SHEET THU
    // ============================================
    const incomeSheet = ss.getSheetByName('THU');
    
    if (!incomeSheet) {
      autoIncomeMessage = '\n⚠️ Không tìm thấy sheet THU. Không thể tự động thêm khoản thu!';
      Logger.log('ERROR: Không tìm thấy sheet THU');
    } else {
      // Tìm dòng trống trong sheet THU
      let incomeLastRow = incomeSheet.getLastRow();
      let incomeEmptyRow = incomeLastRow + 1;
      
      if (incomeLastRow > 1) {
        const incomeDataRange = incomeSheet.getRange(2, 2, incomeLastRow - 1, 1).getValues();
        for (let i = 0; i < incomeDataRange.length; i++) {
          if (!incomeDataRange[i][0]) {
            incomeEmptyRow = i + 2;
            break;
          }
        }
      }
      
      // STT tự động cho sheet THU
      const incomeStt = incomeEmptyRow - 1;
      
      // Xác định nguồn thu dựa vào loại nợ
      let incomeSource = 'Vay ' + debtType;
      
      // Thêm dữ liệu vào sheet THU
      // Columns: STT | Ngày | Số tiền | Nguồn thu | Ghi chú
      incomeSheet.getRange(incomeEmptyRow, 1, 1, 5).setValues([[
        incomeStt,
        date,
        principal,
        incomeSource,
        `Vay: ${debtName}`
      ]]);
      
      // Format
      incomeSheet.getRange(incomeEmptyRow, 2, 1, 1).setNumberFormat('dd/mm/yyyy');
      incomeSheet.getRange(incomeEmptyRow, 3, 1, 1).setNumberFormat('#,##0');
      
      autoIncomeMessage = `\n✅ Đã TỰ ĐỘNG thêm khoản thu "${incomeSource}" vào sheet THU`;
      Logger.log('SUCCESS: Đã thêm khoản thu vào sheet THU tại dòng ' + incomeEmptyRow);
    }
    
    // ============================================
    // BƯỚC 4: TRẢ VỀ KẾT QUẢ
    // ============================================
    const resultMessage = `✅ Đã thêm khoản nợ: ${debtName}\n` +
               `💰 Số tiền: ${principal.toLocaleString('vi-VN')} VNĐ\n` +
               `📅 Kỳ hạn: ${term} tháng\n` +
               `💳 Loại: ${debtType}` +
               autoIncomeMessage;
    
    Logger.log('=== KẾT QUẢ ===');
    Logger.log(resultMessage);
    
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
 * FUNCTION TEST - Để test riêng
 * ============================================
 */
function testAddDebtManagement() {
  const testData = {
    date: '2025-10-29',
    debtName: 'Margin SSI test',
    debtType: 'Margin chứng khoán',
    principal: 30000000,
    interestRate: 9,
    term: 1,
    note: 'Test margin auto'
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