function debugCalendarToSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let debugSheet = ss.getSheetByName('DEBUG_CALENDAR');
  if (!debugSheet) {
    debugSheet = ss.insertSheet('DEBUG_CALENDAR');
  }
  debugSheet.clear();
  debugSheet.appendRow(['Timestamp', 'Level', 'Message', 'Data']);
  
  const log = (msg, data = '') => {
    debugSheet.appendRow([new Date(), 'INFO', msg, JSON.stringify(data)]);
  };
  
  log('Starting Debug...');
  
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  log('Today', today);
  
  // Helper to parse currency
  const parseCurrency = (val) => {
    if (typeof val === 'number') return val;
    if (typeof val === 'string') {
      return parseFloat(val.replace(/\D/g, ''));
    }
    return 0;
  };
  
  // Helper to parse date
  const parseDate = (d) => {
    if (d instanceof Date) return d;
    if (typeof d === 'string') {
      if (d.includes('/')) {
        const parts = d.split('/');
        if (parts.length === 3) return new Date(parts[2], parts[1] - 1, parts[0]);
      }
      return new Date(d);
    }
    return null;
  };

  const processSheet = (sheetName, isDebt) => {
    log(`Processing Sheet: ${sheetName}`);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      log('Sheet not found');
      return;
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      log('No data rows');
      return;
    }
    
    const data = sheet.getRange(2, 1, lastRow - 1, 12).getValues();
    
    data.forEach((row, idx) => {
      const rowNum = idx + 2;
      const name = row[1];
      const term = parseInt(row[5]) || 1;
      const rawStartDate = row[6];
      const rawMaturityDate = row[7];
      const remaining = parseCurrency(row[10]);
      const status = String(row[11]).trim(); // Trim status
      
      const startDate = parseDate(rawStartDate);
      const maturityDate = parseDate(rawMaturityDate);
      
      log(`Row ${rowNum}: ${name}`, {
        term, rawStartDate, rawMaturityDate, remaining, status, 
        parsedStart: startDate, parsedMaturity: maturityDate
      });
      
      let isActive = false;
      if (isDebt) {
        isActive = (status === 'Chưa trả' || status === 'Đang trả');
      } else {
        isActive = (status === 'Đang vay');
      }
      
      if (!isActive) {
        log(`  -> Skipped: Status '${status}' not active`);
        return;
      }
      
      if (remaining <= 0) {
        log('  -> Skipped: Remaining <= 0');
        return;
      }
      
      if (!startDate) {
        log('  -> Skipped: Invalid Start Date');
        return;
      }
      
      if (term > 1) {
        log('  -> Checking Installments...');
        let found = false;
        for (let i = 1; i <= term; i++) {
          let payDate = new Date(startDate);
          payDate.setMonth(payDate.getMonth() + i);
          
          if (payDate >= today) {
            log(`    -> Found Future Payment: ${payDate} (Month ${i})`);
            found = true;
            break; // Just check if we find one
          }
        }
        if (!found) log('    -> No future payments found (all past)');
      } else {
        log('  -> Checking Bullet...');
        if (maturityDate && maturityDate >= today) {
          log(`    -> Found Future Payment: ${maturityDate}`);
        } else {
          log(`    -> Skipped: Maturity ${maturityDate} is past or invalid`);
        }
      }
    });
  };
  
  processSheet('QUẢN LÝ NỢ', true);
  processSheet('CHO VAY', false);
  
  log('Debug Finished');
}
