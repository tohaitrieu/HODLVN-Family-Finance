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
 */

/**
 * Lấy danh sách khoản nợ từ QUẢN LÝ NỢ
 */
function getDebtList() {
  try {
    const sheet = getSheet(APP_CONFIG.SHEETS.DEBT_MANAGEMENT);
    if (!sheet) {
      return ['Chưa có khoản nợ'];
    }
    
    const data = sheet.getDataRange().getValues();
    const debtList = [];
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][1]) { // Cột B: Tên khoản nợ
        debtList.push(data[i][1]);
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
  SpreadsheetApp.getActiveSpreadsheet().toast(message, title, timeoutSeconds);
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