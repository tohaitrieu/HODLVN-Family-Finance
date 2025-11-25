/**
 * ===============================================
 * UTILS.GS - MODULE TIỆN ÍCH CHUNG
 * ===============================================
 * 
 * Chức năng:
 * - Lấy danh sách nợ
 * - Validation dữ liệu
 * - Format & conversion
 * - Các helper functions khác
 * Lấy danh sách khoản nợ từ QUẢN LÝ NỢ
 */

/**
 * Tìm dòng trống thực sự trong sheet
 * 
 * VẤN ĐỀ: Khi khởi tạo sheet, các công thức được điền vào 1000 dòng
 * → getLastRow() trả về 1000
 * → appendRow() chèn dữ liệu vào dòng 1001!
 * 
 * GIẢI PHÁP: Tìm dòng trống dựa vào CÁC CỘT DỮ LIỆU THỰC
 * (không phải cột có công thức)
 * 
 * @param {Sheet} sheet - Sheet cần tìm
 * @param {number} dataColumn - Cột chứa dữ liệu thực (không phải công thức)
 *                              Ví dụ: Cột B (Ngày), Cột C (Số tiền), etc.
 * @return {number} Số dòng trống đầu tiên
 */
function findEmptyRow(sheet, dataColumn) {
  dataColumn = dataColumn || 2; // Mặc định cột B
  
  const lastRow = sheet.getLastRow();
  
  // Nếu chỉ có header, trả về dòng 2
  if (lastRow <= 1) {
    return 2;
  }
  
  // Lấy tất cả dữ liệu từ cột dataColumn, bắt đầu từ dòng 2
  const dataRange = sheet.getRange(2, dataColumn, lastRow - 1, 1).getValues();
  
  // Tìm dòng đầu tiên có giá trị rỗng
  for (let i = 0; i < dataRange.length; i++) {
    const cellValue = dataRange[i][0];
    
    // Kiểm tra xem ô có thực sự trống không
    if (!cellValue || cellValue === '' || cellValue === null) {
      return i + 2; // +2 vì: i bắt đầu từ 0, và dòng đầu tiên là dòng 2
    }
  }
  
  // Nếu không tìm thấy dòng trống, trả về dòng tiếp theo
  return lastRow + 1;
}

/**
 * Thêm dữ liệu vào sheet ở dòng trống đầu tiên
 * 
 * @param {Sheet} sheet - Sheet cần thêm
 * @param {Array} rowData - Mảng dữ liệu cần thêm
 * @param {number} dataColumn - Cột dùng để kiểm tra dòng trống (mặc định: 2 = cột B)
 * @return {number} Số dòng đã thêm dữ liệu
 */
function insertRowAtEmptyPosition(sheet, rowData, dataColumn) {
  dataColumn = dataColumn || 2;
  const emptyRow = findEmptyRow(sheet, dataColumn);
  
  // Thêm dữ liệu vào dòng trống
  sheet.getRange(emptyRow, 1, 1, rowData.length).setValues([rowData]);
  
  return emptyRow;
}

/**
 * Lấy STT tự động dựa vào số dòng có dữ liệu thực
 * 
 * @param {Sheet} sheet - Sheet cần đếm
 * @param {number} dataColumn - Cột dùng để kiểm tra dữ liệu
 * @return {number} STT tiếp theo
 */
function getNextSTT(sheet, dataColumn) {
  dataColumn = dataColumn || 2;
  const emptyRow = findEmptyRow(sheet, dataColumn);
  return emptyRow - 1; // STT = số dòng - 1 (vì có header)
}

/**
 * Format dòng dữ liệu mới được thêm
 * 
 * @param {Sheet} sheet - Sheet chứa dữ liệu
 * @param {number} row - Số dòng cần format
 * @param {Object} formats - Object chứa format cho từng cột
 *                           Ví dụ: {2: 'dd/mm/yyyy', 3: '#,##0'}
 */
function formatNewRow(sheet, row, formats) {
  for (const [col, format] of Object.entries(formats)) {
    sheet.getRange(row, parseInt(col)).setNumberFormat(format);
  }
}

/**
 * ========================================
 * HÀM TEST - Kiểm tra findEmptyRow()
 * ========================================
 */
function testFindEmptyRow() {
  const ss = getSpreadsheet();
  
  // Test với sheet THU
  const incomeSheet = ss.getSheetByName('THU');
  if (incomeSheet) {
    Logger.log('=== TEST SHEET THU ===');
    Logger.log('getLastRow(): ' + incomeSheet.getLastRow());
    Logger.log('findEmptyRow(): ' + findEmptyRow(incomeSheet, 2)); // Cột B = Ngày
  }
  
  // Test với sheet QUẢN LÝ NỢ
  const debtSheet = ss.getSheetByName('QUẢN LÝ NỢ');
  if (debtSheet) {
    Logger.log('=== TEST SHEET QUẢN LÝ NỢ ===');
    Logger.log('getLastRow(): ' + debtSheet.getLastRow());
    Logger.log('findEmptyRow(): ' + findEmptyRow(debtSheet, 2)); // Cột B = Tên khoản nợ
  }
  
  // Test với sheet CHỨNG KHOÁN
  const stockSheet = ss.getSheetByName('CHỨNG KHOÁN');
  if (stockSheet) {
    Logger.log('=== TEST SHEET CHỨNG KHOÁN ===');
    Logger.log('getLastRow(): ' + stockSheet.getLastRow());
    Logger.log('findEmptyRow(): ' + findEmptyRow(stockSheet, 2)); // Cột B = Ngày
  }
}

// ==================== ORIGINAL UTILS FUNCTIONS ====================

/**
 * Lấy danh sách khoản nợ từ QUẢN LÝ NỢ - FIXED VERSION
 */
function getDebtList() {
  try {
    const sheet = getSheet(APP_CONFIG.SHEETS.DEBT_MANAGEMENT);
    if (!sheet) {
      return ['Chưa có khoản nợ'];
    }
    
    // ✅ FIX: Sử dụng findEmptyRow() thay vì getDataRange()
    const emptyRow = findEmptyRow(sheet, 2);
    const dataRows = emptyRow - 2; // Trừ header và bắt đầu từ 0
    
    if (dataRows <= 0) {
      return ['Chưa có khoản nợ'];
    }
    
    const data = sheet.getRange(2, 2, dataRows, 1).getValues();
    const debtList = [];
    
    for (let i = 0; i < data.length; i++) {
      if (data[i][0]) { // Cột B: Tên khoản nợ
        debtList.push(data[i][0]);
      }
    }
    
    return debtList.length > 0 ? debtList : ['Chưa có khoản nợ'];
    
  } catch (error) {
    Logger.log('Error getting debt list: ' + error.message);
    return ['Chưa có khoản nợ'];
  }
}

/**
 * Validate số tiền
 */
function validateAmount(amount) {
  const num = parseFloat(amount);
  return !isNaN(num) && num > 0;
}

/**
 * Validate ngày
 */
function validateDate(dateString) {
  try {
    const date = new Date(dateString);
    return date instanceof Date && !isNaN(date);
  } catch (error) {
    return false;
  }
}

/**
 * Parse ngày từ string ISO (YYYY-MM-DD)
 */
function parseISODate(dateString) {
  try {
    const parts = dateString.split('-');
    return new Date(
      parseInt(parts[0]),
      parseInt(parts[1]) - 1,
      parseInt(parts[2])
    );
  } catch (error) {
    return null;
  }
}

/**
 * Format ngày thành string VN
 */
function formatDateVN(date) {
  if (!(date instanceof Date)) return '';
  
  const day = ('0' + date.getDate()).slice(-2);
  const month = ('0' + (date.getMonth() + 1)).slice(-2);
  const year = date.getFullYear();
  
  return `${day}/${month}/${year}`;
}

/**
 * Lấy tháng hiện tại
 */
function getCurrentMonth() {
  return new Date().getMonth() + 1;
}

/**
 * Lấy năm hiện tại
 */
function getCurrentYear() {
  return new Date().getFullYear();
}

/**
 * Lấy quý hiện tại
 */
function getCurrentQuarter() {
  return Math.ceil(getCurrentMonth() / 3);
}

/**
 * Lấy số ngày trong tháng
 */
function getDaysInMonth(month, year) {
  return new Date(year, month, 0).getDate();
}

/**
 * Tính phần trăm
 */
function calculatePercentage(value, total) {
  if (!total || total === 0) return 0;
  return (value / total) * 100;
}

/**
 * Làm tròn số
 */
function roundNumber(number, decimals) {
  const multiplier = Math.pow(10, decimals || 0);
  return Math.round(number * multiplier) / multiplier;
}

/**
 * Kiểm tra sheet tồn tại
 */
function sheetExists(sheetName) {
  return getSheet(sheetName) !== null;
}

/**
 * Tạo sheet nếu chưa tồn tại
 */
function createSheetIfNotExists(sheetName) {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  
  return sheet;
}

/**
 * Xóa sheet nếu tồn tại
 */
function deleteSheetIfExists(sheetName) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  
  if (sheet) {
    ss.deleteSheet(sheet);
  }
}

/**
 * Copy sheet
 */
function copySheet(sourceSheetName, targetSheetName) {
  const ss = getSpreadsheet();
  const sourceSheet = ss.getSheetByName(sourceSheetName);
  
  if (!sourceSheet) {
    throw new Error(`Sheet ${sourceSheetName} không tồn tại`);
  }
  
  const targetSheet = sourceSheet.copyTo(ss);
  targetSheet.setName(targetSheetName);
  
  return targetSheet;
}

/**
 * Lấy dòng cuối cùng có dữ liệu
 */
function getLastDataRow(sheet, column) {
  column = column || 1;
  const maxRows = sheet.getMaxRows();
  const values = sheet.getRange(1, column, maxRows, 1).getValues();
  
  for (let i = values.length - 1; i >= 0; i--) {
    if (values[i][0] !== '' && values[i][0] !== null) {
      return i + 1;
    }
  }
  
  return 1;
}

/**
 * Lấy cột cuối cùng có dữ liệu
 */
function getLastDataColumn(sheet, row) {
  row = row || 1;
  const maxCols = sheet.getMaxColumns();
  const values = sheet.getRange(row, 1, 1, maxCols).getValues()[0];
  
  for (let i = values.length - 1; i >= 0; i--) {
    if (values[i] !== '' && values[i] !== null) {
      return i + 1;
    }
  }
  
  return 1;
}

/**
 * Tự động điều chỉnh độ rộng cột
 */
function autoResizeColumns(sheet, startCol, endCol) {
  for (let col = startCol; col <= endCol; col++) {
    sheet.autoResizeColumn(col);
  }
}

/**
 * Set border cho range
 */
function setBorder(range, color) {
  color = color || 'black';
  range.setBorder(true, true, true, true, true, true, color, SpreadsheetApp.BorderStyle.SOLID);
}

/**
 * Set alternating colors cho range
 */
function setAlternatingColors(range) {
  range.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
}

/**
 * Freeze rows và columns
 */
function freezeHeaderRow(sheet, numRows) {
  numRows = numRows || 1;
  sheet.setFrozenRows(numRows);
}

function freezeHeaderColumn(sheet, numCols) {
  numCols = numCols || 1;
  sheet.setFrozenColumns(numCols);
}

/**
 * Sort sheet by column
 */
function sortSheetByColumn(sheet, columnIndex, ascending) {
  ascending = ascending !== false;
  const lastRow = getLastDataRow(sheet, columnIndex);
  const range = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
  range.sort({column: columnIndex, ascending: ascending});
}

/**
 * Filter unique values
 */
function getUniqueValues(arr) {
  return [...new Set(arr)];
}

/**
 * Sum array
 */
function sumArray(arr) {
  return arr.reduce((a, b) => a + (parseFloat(b) || 0), 0);
}

/**
 * Average array
 */
function averageArray(arr) {
  const validNumbers = arr.filter(x => !isNaN(parseFloat(x)));
  if (validNumbers.length === 0) return 0;
  return sumArray(validNumbers) / validNumbers.length;
}

/**
 * Generate random color
 */
function randomColor() {
  const letters = '0123456789ABCDEF';
  let color = '#';
  for (let i = 0; i < 6; i++) {
    color += letters[Math.floor(Math.random() * 16)];
  }
  return color;
}

/**
 * Lighten color
 */
function lightenColor(color, percent) {
  const num = parseInt(color.replace('#', ''), 16);
  const amt = Math.round(2.55 * percent);
  const R = (num >> 16) + amt;
  const G = (num >> 8 & 0x00FF) + amt;
  const B = (num & 0x0000FF) + amt;
  
  return '#' + (0x1000000 +
    (R < 255 ? R < 1 ? 0 : R : 255) * 0x10000 +
    (G < 255 ? G < 1 ? 0 : G : 255) * 0x100 +
    (B < 255 ? B < 1 ? 0 : B : 255))
    .toString(16).slice(1);
}

/**
 * Log to console
 */
function log(message, data) {
  Logger.log(`[${new Date().toISOString()}] ${message}`);
  if (data) {
    Logger.log(JSON.stringify(data, null, 2));
  }
}

/**
 * Show toast notification
 */
function showToast(message, title, timeoutSeconds) {
  title = title || 'Thông báo';
  timeoutSeconds = timeoutSeconds || 3;
  getSpreadsheet().toast(message, title, timeoutSeconds);
}

/**
 * Create timestamp
 */
function createTimestamp() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}

/**
 * Sleep function
 */
function sleep(milliseconds) {
  Utilities.sleep(milliseconds);
}

/**
 * Batch update (reduce API calls)
 */
function batchUpdate(updates) {
  SpreadsheetApp.flush();
  updates.forEach(update => update());
  SpreadsheetApp.flush();
}

/**
 * Get script properties
 */
function getScriptProperty(key) {
  return PropertiesService.getScriptProperties().getProperty(key);
}

/**
 * Set script property
 */
function setScriptProperty(key, value) {
  PropertiesService.getScriptProperties().setProperty(key, value);
}

/**
 * Delete script property
 */
function deleteScriptProperty(key) {
  PropertiesService.getScriptProperties().deleteProperty(key);
}

/**
 * Get all script properties
 */
function getAllScriptProperties() {
  return PropertiesService.getScriptProperties().getProperties();
}

/**
 * Send email
 */
function sendEmail(recipient, subject, body) {
  GmailApp.sendEmail(recipient, subject, body);
}

/**
 * Create trigger
 */
function createTimeDrivenTrigger(functionName, hours) {
  ScriptApp.newTrigger(functionName)
    .timeBased()
    .everyHours(hours)
    .create();
}

/**
 * Delete all triggers
 */
function deleteAllTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
}

/**
 * Get active user email
 */
function getActiveUserEmail() {
  return Session.getActiveUser().getEmail();
}

/**
 * Check if user is owner
 */
function isOwner() {
  return Session.getActiveUser().getEmail() === getSpreadsheet().getOwner().getEmail();
}

/**
 * Format số tiền (không đơn vị)
 */
function formatCurrency(amount) {
  return new Intl.NumberFormat('vi-VN', { 
    style: 'decimal',
    maximumFractionDigits: 0 
  }).format(amount);
}

/**
 * Lấy spreadsheet hiện tại
 *
 * Supports both standalone and library mode:
 * - Standalone mode: Returns getActiveSpreadsheet() (default behavior)
 * - Library mode: Returns openById(LIBRARY_SPREADSHEET_ID) if initLibrary() was called
 *
 * This fixes the issue where library functions write to caller's spreadsheet
 * instead of the library's data spreadsheet.
 *
 * @returns {Spreadsheet} The target spreadsheet object
 */
function getSpreadsheet() {
  // Check if library mode is active (LIBRARY_SPREADSHEET_ID is defined in LibraryConfig.gs)
  if (typeof LIBRARY_SPREADSHEET_ID !== 'undefined' && LIBRARY_SPREADSHEET_ID !== null) {
    // Library mode: Use the configured spreadsheet ID
    return SpreadsheetApp.openById(LIBRARY_SPREADSHEET_ID);
  }

  // Standalone mode: Use active spreadsheet (backward compatible)
  return SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * Lấy sheet theo tên
 */
function getSheet(sheetName) {
  return getSpreadsheet().getSheetByName(sheetName);
}

/**
 * Log current spreadsheet context for debugging.
 * Useful for troubleshooting library vs standalone mode issues.
 *
 * @example
 * logSpreadsheetContext();
 * // Logs: "📊 Spreadsheet Context: LIBRARY mode using 'My Finance Data' (ID: 1a2b3c...)"
 */
function logSpreadsheetContext() {
  try {
    const ss = getSpreadsheet();
    const mode = (typeof LIBRARY_SPREADSHEET_ID !== 'undefined' && LIBRARY_SPREADSHEET_ID !== null) ? 'LIBRARY' : 'STANDALONE';
    const name = ss.getName();
    const id = ss.getId();

    Logger.log('📊 Spreadsheet Context: ' + mode + ' mode using "' + name + '" (ID: ' + id + ')');
    return {
      mode: mode,
      name: name,
      id: id
    };
  } catch (error) {
    Logger.log('❌ Error getting spreadsheet context: ' + error.message);
    return {
      error: error.message
    };
  }
}

/**
 * Force recalculate toàn bộ sheet
 */
function forceRecalculate() {
  SpreadsheetApp.flush();
  getSpreadsheet().getSheets().forEach(sheet => {
    sheet.getDataRange().getValues();
  });
}