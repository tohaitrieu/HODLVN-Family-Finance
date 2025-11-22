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
    
    // ✅ FIX: Sử dụng findEmptyRow() thay vì getLastRow()
    // Cột 2 (B) = Tên khoản nợ - cột dữ liệu thực
    const emptyRow = findEmptyRow(debtSheet, 2);
    const stt = getNextSTT(debtSheet, 2);
    
    Logger.log('QUẢN LÝ NỢ - Dòng trống tìm được: ' + emptyRow);
    Logger.log('QUẢN LÝ NỢ - STT: ' + stt);
    
    // ============================================
    // CRITICAL FIX v3.3.2: Chia làm 2 phần để KHÔNG ghi đè công thức cột K (Mới)
    // ============================================
    
    // Phần 1: Cột A-J (STT đến Đã trả lãi) - 10 cột
    // STT, Name, Type, Principal, Rate, Term, Date, Maturity, PaidPrin, PaidInt
    const rowDataPart1 = [
      stt,                    // A: STT
      debtName,               // B: Tên khoản nợ
      debtType,               // C: Loại hình (NEW)
      principal,              // D: Gốc
      interestRate / 100,     // E: Lãi suất (chuyển % sang decimal)
      term,                   // F: Kỳ hạn
      date,                   // G: Ngày vay
      maturityDate,           // H: Đáo hạn
      0,                      // I: Đã trả gốc
      0                       // J: Đã trả lãi
    ];
    
    // Phần 2: Cột L-M (Trạng thái và Ghi chú) - 2 cột
    const rowDataPart2 = [
      'Chưa trả',             // L: Trạng thái
      note                    // M: Ghi chú
    ];
    
    // ✅ Insert Phần 1: Cột A-J (10 cột)
    debtSheet.getRange(emptyRow, 1, 1, rowDataPart1.length).setValues([rowDataPart1]);
    
    // ✅ BỎ QUA cột K (cột 11) - GIỮ NGUYÊN CÔNG THỨC =D-I
    
    // ✅ Insert Phần 2: Cột L-M (2 cột, bắt đầu từ cột 12)
    debtSheet.getRange(emptyRow, 12, 1, rowDataPart2.length).setValues([rowDataPart2]);
    
    Logger.log('✅ ĐÃ INSERT XONG! Công thức cột K được giữ nguyên.');
    
    // Format
    formatNewRow(debtSheet, emptyRow, {
      4: '#,##0',           // D: Gốc
      5: '0.00"%"',         // E: Lãi suất
      7: 'dd/mm/yyyy',      // G: Ngày vay
      8: 'dd/mm/yyyy',      // H: Đáo hạn
      9: '#,##0',           // I: Đã trả gốc
      10: '#,##0',          // J: Đã trả lãi
      11: '#,##0'           // K: Còn nợ (công thức đã có sẵn)
    });
    
    // ============================================
    // BƯỚC 2: TỰ ĐỘNG THÊM KHOẢN THU
    // ============================================
    let autoIncomeMessage = '';
    
    const incomeSheet = ss.getSheetByName('THU');
    
    if (!incomeSheet) {
      autoIncomeMessage = '\n⚠️ Không tìm thấy sheet THU. Không thể tự động thêm khoản thu!';
      Logger.log('ERROR: Không tìm thấy sheet THU');
    } else {
      // ✅ FIX: Sử dụng findEmptyRow() thay vì getLastRow()
      // Cột 2 (B) = Ngày - cột dữ liệu thực
      const incomeEmptyRow = findEmptyRow(incomeSheet, 2);
      const incomeStt = getNextSTT(incomeSheet, 2);
      
      Logger.log('THU - Dòng trống tìm được: ' + incomeEmptyRow);
      Logger.log('THU - STT: ' + incomeStt);
      
      // Xác định nguồn thu dựa vào loại nợ
      let incomeSource = 'Vay ' + debtType;
      
      // Thêm dữ liệu vào sheet THU
      // Columns: STT | Ngày | Số tiền | Nguồn thu | Ghi chú
      const incomeRowData = [
        incomeStt,
        date,
        principal,
        incomeSource,
        `Vay: ${debtName}`
      ];
      
      incomeSheet.getRange(incomeEmptyRow, 1, 1, incomeRowData.length).setValues([incomeRowData]);
      
      // Format
      formatNewRow(incomeSheet, incomeEmptyRow, {
        2: 'dd/mm/yyyy',
        3: '#,##0'
      });
      
      autoIncomeMessage = `\n✅ Đã TỰ ĐỘNG thêm khoản thu "${incomeSource}" vào sheet THU`;
      Logger.log('SUCCESS: Đã thêm khoản thu vào sheet THU tại dòng ' + incomeEmptyRow);
    }
    
    // ============================================
    // BƯỚC 3: TRẢ VỀ KẾT QUẢ
    // ============================================
    const resultMessage = `✅ Đã thêm khoản nợ: ${debtName}\n` +
               `💰 Số tiền: ${principal.toLocaleString('vi-VN')}\n` +
               `📅 Kỳ hạn: ${term} tháng\n` +
               `💳 Loại: ${debtType}\n` +
               `📊 Trạng thái: Chưa trả` +
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
  const ss = SpreadsheetApp.getActiveSpreadsheet();
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