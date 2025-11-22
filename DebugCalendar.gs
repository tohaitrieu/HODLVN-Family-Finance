function debugCalendarEvents() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const events = [];
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  
  Logger.log('Today: ' + today);
  
  // Helper to process installments
  const processInstallments = (sheetName, isDebt) => {
    Logger.log(`Processing Sheet: ${sheetName}`);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      Logger.log('Sheet not found');
      return;
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      Logger.log('No data rows');
      return;
    }
    
    // A(0): STT, B(1): Name, C(2): Type, D(3): Principal, E(4): Rate, F(5): Term, G(6): Date, H(7): Maturity
    // I(8): PaidPrin, J(9): PaidInt, K(10): Remaining, L(11): Status
    const data = sheet.getRange(2, 1, lastRow - 1, 12).getValues();
    
    data.forEach((row, index) => {
      const name = row[1];
      const initialPrincipal = parseFloat(row[3]) || 0;
      const rate = parseFloat(row[4]) || 0; // Annual Rate (decimal)
      const term = parseInt(row[5]) || 1;
      const startDate = row[6];
      const maturityDate = row[7];
      const remaining = parseFloat(row[10]) || 0;
      const status = row[11];
      
      Logger.log(`Row ${index + 2}: ${name} | Type: ${row[2]} | Term: ${term} | Start: ${startDate} | Remaining: ${remaining} | Status: ${status}`);
      
      // Check active status
      const isActive = isDebt ? (status !== 'Đã thanh toán') : (status === 'Đang vay');
      
      if (!isActive) {
        Logger.log('  -> Skipped: Not active');
        return;
      }
      if (remaining <= 0) {
        Logger.log('  -> Skipped: No remaining balance');
        return;
      }
      if (!(startDate instanceof Date)) {
        Logger.log('  -> Skipped: Invalid Start Date');
        return;
      }
      
      if (term > 1) {
        // Installment Logic
        Logger.log('  -> Installment Logic');
        const monthlyPrincipal = initialPrincipal / term;
        let simulatedRemaining = remaining;
        
        for (let i = 1; i <= term; i++) {
          let payDate = new Date(startDate);
          payDate.setMonth(payDate.getMonth() + i);
          
          if (payDate >= today) {
            Logger.log(`    -> Found Future Payment: ${payDate}`);
            // ... (rest of logic)
            events.push({ date: payDate, name: name });
            
            let currentPrincipal = monthlyPrincipal;
            if (simulatedRemaining < currentPrincipal) currentPrincipal = simulatedRemaining;
            simulatedRemaining -= currentPrincipal;
            if (simulatedRemaining <= 0.1) break;
          } else {
             // Logger.log(`    -> Past Payment: ${payDate}`);
          }
        }
      } else {
        // Bullet Logic
        Logger.log('  -> Bullet Logic');
        if (maturityDate instanceof Date && maturityDate >= today) {
          Logger.log(`    -> Found Future Payment: ${maturityDate}`);
          events.push({ date: maturityDate, name: name });
        } else {
          Logger.log(`    -> Skipped: Maturity Date ${maturityDate} is past or invalid`);
        }
      }
    });
  };
  
  processInstallments('QUẢN LÝ NỢ', true);
  processInstallments('CHO VAY', false);
  
  Logger.log(`Total Events Found: ${events.length}`);
}
