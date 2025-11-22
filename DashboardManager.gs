/**
 * ===============================================
 * DASHBOARDMANAGER.GS v5.0 - 2x2 GRID LAYOUT
 * ===============================================
 * 
 * Layout:
 * - Top Left: Income (Thu nhập)
 * - Top Right: Expense (Chi phí)
 * - Bottom Left: Liabilities (Nợ phải trả - Chi tiết)
 * - Bottom Right: Assets (Tài sản)
 * - Bottom: Monthly Statistics
 * - Chart: Top Right (Columns I-N)
 */

const DashboardManager = {
  
  CONFIG: {
    LAYOUT: {
      LEFT_COL: 1,      // A
      RIGHT_COL: 5,     // E
      CALENDAR_COL: 11, // K
      CHART_COL: 11,    // K
      CHART_ROW: 33,    // Row 33
      START_ROW: 6,     // Start data after header/dropdowns
      COL_WIDTH: 3      // Width of each table (A-B-C)
    },
    COLORS: {
      INCOME: '#4CAF50',
      EXPENSE: '#F44336',
      ASSETS: '#2196F3',
      LIABILITIES: '#FF9800',
      HEADER: '#4472C4',
      TEXT: '#FFFFFF',
      CALENDAR: '#9C27B0'
    }
  },
  
  setupDashboard() {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      let sheet = ss.getSheetByName(APP_CONFIG.SHEETS.DASHBOARD);
      
      if (!sheet) {
        sheet = ss.insertSheet(APP_CONFIG.SHEETS.DASHBOARD);
      } else {
        sheet.clear();
        sheet.clearConditionalFormatRules();
        sheet.getRange('B2:B4').setDataValidation(null);
        sheet.getRange('B2:B4').setNumberFormat('@');
        const charts = sheet.getCharts();
        charts.forEach(chart => sheet.removeChart(chart));
      }
      
      // 1. Setup Header & Dropdowns
      this._setupHeader(sheet);
      
      // 2. Setup 2x2 Grid Tables
      // Return the last row used to position the Monthly Table
      const lastRow = this._setupGridTables(sheet, this.CONFIG.LAYOUT.START_ROW);
      
      // 3. Setup Monthly Table (Table 2)
      this._setupTable2(sheet, lastRow + 3);
      
      // 4. Setup Chart
      this._createChart(sheet);
      
      // 5. Format & Finalize
      this._formatSheet(sheet);
      this._setupTriggers();
      
      SpreadsheetApp.flush();
      
      SpreadsheetApp.getUi().alert(
        'Cập nhật Dashboard thành công!',
        '✅ Đã cập nhật giao diện 2x2 Grid:\n' +
        '1. Thu nhập - Chi phí (Hàng trên)\n' +
        '2. Nợ phải trả - Tài sản (Hàng dưới)\n' +
        '3. Danh sách nợ được liệt kê chi tiết.\n' +
        '4. Biểu đồ đã được sửa lỗi hiển thị.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      
      return sheet;
      
    } catch (error) {
      Logger.log('❌ Lỗi setupDashboard: ' + error.message);
      throw new Error('Không thể khởi tạo Dashboard: ' + error.message);
    }
  },
  
  _setupHeader(sheet) {
    // Title
    sheet.getRange('A1:G1').merge()
      .setValue('📊 BÁO CÁO TÀI CHÍNH (CASHFLOW)')
      .setFontSize(14)
      .setFontWeight('bold')
      .setVerticalAlignment('middle')
      .setBackground(this.CONFIG.COLORS.HEADER)
      .setFontColor(this.CONFIG.COLORS.TEXT);
    sheet.setRowHeight(1, 40);
    
    // Dropdowns
    sheet.getRange('A2').setValue('Năm').setFontWeight('bold');
    sheet.getRange('A3').setValue('Quý').setFontWeight('bold');
    sheet.getRange('A4').setValue('Tháng').setFontWeight('bold');
    
    const currentYear = new Date().getFullYear();
    sheet.getRange('B2:B4').setNumberFormat('@');
    
    const yearList = ['Tất cả'];
    for (let y = currentYear - 5; y <= currentYear + 2; y++) yearList.push(y.toString());
    
    const monthList = ['Tất cả'];
    for (let m = 1; m <= 12; m++) monthList.push(`Tháng ${m}`);
    
    sheet.getRange('B2').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(yearList).build());
    sheet.getRange('B3').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Tất cả', 'Quý 1', 'Quý 2', 'Quý 3', 'Quý 4']).build());
    sheet.getRange('B4').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(monthList).build());
    
    sheet.getRange('B2').setValue(currentYear.toString());
    sheet.getRange('B3').setValue('Tất cả');
    sheet.getRange('B4').setValue('Tất cả');
  },
  

  _setupGridTables(sheet, startRow) {
    const cfg = this.CONFIG.LAYOUT;
    let currentRow = startRow;
    
    // === ROW 1: INCOME (Left) & EXPENSE (Middle) & CALENDAR (Right) ===
    
    // 1. Income Table (3 Cols: Name, Value, %)
    const incomeCategories = APP_CONFIG.CATEGORIES.INCOME;
    const incomeRows = [...incomeCategories, 'TỔNG THU NHẬP'];
    const incomeHeight = this._renderTable(sheet, currentRow, cfg.LEFT_COL, '1. Báo cáo Thu nhập', this.CONFIG.COLORS.INCOME, incomeRows, 3, true);
    
    // Formulas for Income
    const incStart = currentRow + 1;
    const incTotalRow = incStart + incomeCategories.length;
    
    incomeCategories.forEach((cat, idx) => {
      const r = incStart + idx;
      // Value
      sheet.getRange(r, cfg.LEFT_COL + 1).setFormula('=' + this._createDynamicSumFormula('THU', 'C', cat, 'D'));
      // % (Value / Total)
      sheet.getRange(r, cfg.LEFT_COL + 2).setFormula(`=IFERROR(R[0]C[-1] / R${incTotalRow}C[-1], 0)`);
    });
    
    // Total Income
    sheet.getRange(incTotalRow, cfg.LEFT_COL + 1).setFormula(`=SUM(R[-${incomeCategories.length}]C:R[-1]C)`);
    sheet.getRange(incTotalRow, cfg.LEFT_COL + 2).setValue(1).setNumberFormat('0%');
    
    // 2. Expense Table (3 Cols: Name, Value, %)
    const expenseCategories = APP_CONFIG.CATEGORIES.EXPENSE;
    const expenseRows = [...expenseCategories, 'Trả nợ (Gốc + Lãi)', 'TỔNG CHI PHÍ'];
    const expenseHeight = this._renderTable(sheet, currentRow, cfg.RIGHT_COL, '2. Báo cáo Chi phí', this.CONFIG.COLORS.EXPENSE, expenseRows, 3, true);
    
    // Formulas for Expense
    const expStart = currentRow + 1;
    const expTotalRow = expStart + expenseCategories.length + 1; // +1 for Debt row
    
    expenseCategories.forEach((cat, idx) => {
      const r = expStart + idx;
      // Value
      sheet.getRange(r, cfg.RIGHT_COL + 1).setFormula('=' + this._createDynamicSumFormula('CHI', 'C', cat, 'D'));
      // %
      sheet.getRange(r, cfg.RIGHT_COL + 2).setFormula(`=IFERROR(R[0]C[-1] / R${expTotalRow}C[-1], 0)`);
    });
    
    // Formula for 'Trả nợ'
    const debtRowIdx = expStart + expenseCategories.length;
    sheet.getRange(debtRowIdx, cfg.RIGHT_COL + 1).setFormula('=' + `${this._createDynamicSumFormula('TRẢ NỢ', 'D')} + ${this._createDynamicSumFormula('TRẢ NỢ', 'E')}`);
    sheet.getRange(debtRowIdx, cfg.RIGHT_COL + 2).setFormula(`=IFERROR(R[0]C[-1] / R${expTotalRow}C[-1], 0)`);
    
    // Total Expense
    sheet.getRange(expTotalRow, cfg.RIGHT_COL + 1).setFormula(`=SUM(R[-${expenseCategories.length + 1}]C:R[-1]C)`);
    sheet.getRange(expTotalRow, cfg.RIGHT_COL + 2).setValue(1).setNumberFormat('0%');
    
    // 3. Calendar of Events (Right - Col K)
    const calendarHeight = this._renderCalendarTable(sheet, currentRow, cfg.CALENDAR_COL);

    // Calculate max height for Row 1
    const row1Height = Math.max(incomeHeight, expenseHeight, calendarHeight);
    currentRow += row1Height + 2; // +2 padding
    
    // === ROW 2: LIABILITIES (Left) & ASSETS (Right) ===
    
    // 4. Liabilities Table (3 Cols: Name, Value, %)
    const debtItems = this._getDebtItems();
    const liabilityRows = [...debtItems.map(d => d.name), 'TỔNG NỢ'];
    const liabilityHeight = this._renderTable(sheet, currentRow, cfg.LEFT_COL, '3. Báo cáo Nợ phải trả', this.CONFIG.COLORS.LIABILITIES, liabilityRows, 3, true);
    
    // Formulas for Liabilities
    const liabStart = currentRow + 1;
    const liabTotalRow = liabStart + debtItems.length;
    
    debtItems.forEach((item, idx) => {
      const r = liabStart + idx;
      const formula = `=IFERROR(SUMIFS('QUẢN LÝ NỢ'!K:K, 'QUẢN LÝ NỢ'!B:B, "${item.name}"), 0)`; // Col K is Remaining
      sheet.getRange(r, cfg.LEFT_COL + 1).setFormula(formula);
      // %
      sheet.getRange(r, cfg.LEFT_COL + 2).setFormula(`=IFERROR(R[0]C[-1] / R${liabTotalRow}C[-1], 0)`);
    });
    
    // Total Liability
    sheet.getRange(liabTotalRow, cfg.LEFT_COL + 1).setFormula(`=SUM(R[-${debtItems.length}]C:R[-1]C)`);
    sheet.getRange(liabTotalRow, cfg.LEFT_COL + 2).setValue(1).setNumberFormat('0%');
    
    // 5. Assets Table (5 Cols: E, F, G, H, I)
    // Name, Capital, P/L, Current, % (on Current)
    const assetRows = ['Tiền mặt (Ròng)', 'Chứng khoán', 'Vàng', 'Crypto', 'Đầu tư khác', 'Cho vay', 'TỔNG TÀI SẢN'];
    
    // Header for Assets
    const assetHeaderRow = currentRow;
    sheet.getRange(assetHeaderRow, cfg.RIGHT_COL, 1, 5).merge()
      .setValue('4. Báo cáo Tài sản')
      .setFontWeight('bold')
      .setBackground(this.CONFIG.COLORS.ASSETS)
      .setFontColor('#FFFFFF')
      .setHorizontalAlignment('left');
      
    // Sub-headers
    sheet.getRange(assetHeaderRow + 1, cfg.RIGHT_COL).setValue('Danh mục').setFontWeight('bold');
    sheet.getRange(assetHeaderRow + 1, cfg.RIGHT_COL + 1).setValue('Tổng vốn').setFontWeight('bold');
    sheet.getRange(assetHeaderRow + 1, cfg.RIGHT_COL + 2).setValue('Lãi/Lỗ').setFontWeight('bold');
    sheet.getRange(assetHeaderRow + 1, cfg.RIGHT_COL + 3).setValue('Giá trị HT').setFontWeight('bold');
    sheet.getRange(assetHeaderRow + 1, cfg.RIGHT_COL + 4).setValue('Tỷ lệ').setFontWeight('bold');
    
    // Rows
    assetRows.forEach((name, idx) => {
      const r = assetHeaderRow + 2 + idx;
      sheet.getRange(r, cfg.RIGHT_COL).setValue(name);
      
      if (idx === assetRows.length - 1) {
        sheet.getRange(r, cfg.RIGHT_COL).setFontWeight('bold');
        sheet.getRange(r, cfg.RIGHT_COL, 1, 5).setBackground('#EEEEEE');
      }
    });
    
    // Border
    sheet.getRange(assetHeaderRow, cfg.RIGHT_COL, assetRows.length + 2, 5)
      .setBorder(true, true, true, true, true, true);
    
    const assetStart = assetHeaderRow + 2;
    const assetTotalRow = assetStart + 6;
    
    // Formulas for Assets
    // Cash
    const totalIncome = this._createDynamicSumFormula('THU', 'C');
    const totalExpense = `(${this._createDynamicSumFormula('CHI', 'C')} + ${this._createDynamicSumFormula('TRẢ NỢ', 'D')} + ${this._createDynamicSumFormula('TRẢ NỢ', 'E')})`;
    const investCK = `SUMIF('CHỨNG KHOÁN'!C:C,"Mua",'CHỨNG KHOÁN'!H:H)`;
    const investGold = `SUMIF(VÀNG!D:D,"Mua",VÀNG!I:I)`;
    const investCrypto = `SUMIF(CRYPTO!C:C,"Mua",CRYPTO!I:I)`;
    const investOther = `SUM('ĐẦU TƯ KHÁC'!D:D)`;
    const totalInvest = `(${investCK} + ${investGold} + ${investCrypto} + ${investOther})`;
    const divestCK = `SUMIF('CHỨNG KHOÁN'!C:C,"Bán",'CHỨNG KHOÁN'!H:H)`;
    const divestGold = `SUMIF(VÀNG!D:D,"Bán",VÀNG!I:I)`;
    const divestCrypto = `SUMIF(CRYPTO!C:C,"Bán",CRYPTO!I:I)`;
    const totalDivest = `(${divestCK} + ${divestGold} + ${divestCrypto})`;
    const netCashFormula = `IFERROR(${totalIncome} - ${totalExpense} - ${totalInvest} + ${totalDivest}, 0)`;
    
    // 1. Tiền mặt
    sheet.getRange(assetStart, cfg.RIGHT_COL + 1).setFormula('=' + netCashFormula);
    sheet.getRange(assetStart, cfg.RIGHT_COL + 2).setValue(0);
    sheet.getRange(assetStart, cfg.RIGHT_COL + 3).setFormula('=' + netCashFormula);
    sheet.getRange(assetStart, cfg.RIGHT_COL + 4).setFormula(`=IFERROR(R[0]C[-1] / R${assetTotalRow}C[-1], 0)`);
    
    // 2. Chứng khoán
    sheet.getRange(assetStart + 1, cfg.RIGHT_COL + 1).setFormula(`=IFERROR(SUM('CHỨNG KHOÁN'!H:H) - SUM('CHỨNG KHOÁN'!I:I), 0)`);
    sheet.getRange(assetStart + 1, cfg.RIGHT_COL + 2).setFormula(`=IFERROR(SUM('CHỨNG KHOÁN'!N:N), 0)`);
    sheet.getRange(assetStart + 1, cfg.RIGHT_COL + 3).setFormula(`=IFERROR(SUM('CHỨNG KHOÁN'!M:M), 0)`);
    sheet.getRange(assetStart + 1, cfg.RIGHT_COL + 4).setFormula(`=IFERROR(R[0]C[-1] / R${assetTotalRow}C[-1], 0)`);
    
    // 3. Vàng
    sheet.getRange(assetStart + 2, cfg.RIGHT_COL + 1).setFormula(`=IFERROR(SUM(VÀNG!I:I), 0)`);
    sheet.getRange(assetStart + 2, cfg.RIGHT_COL + 2).setFormula(`=IFERROR(SUM(VÀNG!L:L), 0)`);
    sheet.getRange(assetStart + 2, cfg.RIGHT_COL + 3).setFormula(`=IFERROR(SUM(VÀNG!K:K), 0)`);
    sheet.getRange(assetStart + 2, cfg.RIGHT_COL + 4).setFormula(`=IFERROR(R[0]C[-1] / R${assetTotalRow}C[-1], 0)`);
    
    // 4. Crypto
    sheet.getRange(assetStart + 3, cfg.RIGHT_COL + 1).setFormula(`=IFERROR(SUM(CRYPTO!I:I), 0)`);
    sheet.getRange(assetStart + 3, cfg.RIGHT_COL + 2).setFormula(`=IFERROR(SUM(CRYPTO!N:N), 0)`);
    sheet.getRange(assetStart + 3, cfg.RIGHT_COL + 3).setFormula(`=IFERROR(SUM(CRYPTO!M:M), 0)`);
    sheet.getRange(assetStart + 3, cfg.RIGHT_COL + 4).setFormula(`=IFERROR(R[0]C[-1] / R${assetTotalRow}C[-1], 0)`);
    
    // 5. Đầu tư khác
    sheet.getRange(assetStart + 4, cfg.RIGHT_COL + 1).setFormula(`=IFERROR(SUM('ĐẦU TƯ KHÁC'!D:D), 0)`);
    sheet.getRange(assetStart + 4, cfg.RIGHT_COL + 2).setValue(0);
    sheet.getRange(assetStart + 4, cfg.RIGHT_COL + 3).setFormula(`=IFERROR(SUM('ĐẦU TƯ KHÁC'!D:D), 0)`);
    sheet.getRange(assetStart + 4, cfg.RIGHT_COL + 4).setFormula(`=IFERROR(R[0]C[-1] / R${assetTotalRow}C[-1], 0)`);
    
    // 6. Cho vay
    sheet.getRange(assetStart + 5, cfg.RIGHT_COL + 1).setFormula(`=IFERROR(SUM('CHO VAY'!K:K), 0)`); // K is Remaining
    sheet.getRange(assetStart + 5, cfg.RIGHT_COL + 2).setValue(0);
    sheet.getRange(assetStart + 5, cfg.RIGHT_COL + 3).setFormula(`=IFERROR(SUM('CHO VAY'!K:K), 0)`);
    sheet.getRange(assetStart + 5, cfg.RIGHT_COL + 4).setFormula(`=IFERROR(R[0]C[-1] / R${assetTotalRow}C[-1], 0)`);
    
    // Total Assets
    sheet.getRange(assetTotalRow, cfg.RIGHT_COL + 1).setFormula(`=SUM(R[-6]C:R[-1]C)`);
    sheet.getRange(assetTotalRow, cfg.RIGHT_COL + 2).setFormula(`=SUM(R[-6]C:R[-1]C)`);
    sheet.getRange(assetTotalRow, cfg.RIGHT_COL + 3).setFormula(`=SUM(R[-6]C:R[-1]C)`);
    sheet.getRange(assetTotalRow, cfg.RIGHT_COL + 4).setValue(1).setNumberFormat('0%');
    
    // Calculate max height for Row 2
    const row2Height = Math.max(liabilityHeight, assetRows.length + 2);
    
    // === FORMAT NUMBERS FOR ALL TABLES ===
    // Income
    sheet.getRange(incStart, cfg.LEFT_COL + 1, incomeCategories.length + 1, 1).setNumberFormat('#,##0');
    sheet.getRange(incStart, cfg.LEFT_COL + 2, incomeCategories.length + 1, 1).setNumberFormat('0.0%');
    // Expense
    sheet.getRange(expStart, cfg.RIGHT_COL + 1, expenseCategories.length + 2, 1).setNumberFormat('#,##0');
    sheet.getRange(expStart, cfg.RIGHT_COL + 2, expenseCategories.length + 2, 1).setNumberFormat('0.0%');
    // Liabilities
    sheet.getRange(liabStart, cfg.LEFT_COL + 1, debtItems.length + 1, 1).setNumberFormat('#,##0');
    sheet.getRange(liabStart, cfg.LEFT_COL + 2, debtItems.length + 1, 1).setNumberFormat('0.0%');
    // Assets
    sheet.getRange(assetStart, cfg.RIGHT_COL + 1, 7, 3).setNumberFormat('#,##0');
    sheet.getRange(assetStart, cfg.RIGHT_COL + 4, 7, 1).setNumberFormat('0.0%');
    
    return currentRow + row2Height;
  },
  
  _renderTable(sheet, startRow, startCol, title, color, rows, numCols = 3, hasPercentage = false) {
    // Header
    sheet.getRange(startRow, startCol, 1, numCols).merge()
      .setValue(title)
      .setFontWeight('bold')
      .setBackground(color)
      .setFontColor('#FFFFFF')
      .setHorizontalAlignment('left');
      
    // Sub-headers (Row 2)
    const subHeaderRow = startRow + 1;
    sheet.getRange(subHeaderRow, startCol).setValue('Danh mục').setFontWeight('bold');
    sheet.getRange(subHeaderRow, startCol + 1).setValue('Giá trị').setFontWeight('bold');
    if (hasPercentage) {
      sheet.getRange(subHeaderRow, startCol + 2).setValue('Tỷ lệ').setFontWeight('bold');
    }
      
    // Rows
    rows.forEach((name, idx) => {
      const r = startRow + 2 + idx;
      sheet.getRange(r, startCol).setValue(name);
      
      // Last row is Total
      if (idx === rows.length - 1) {
        sheet.getRange(r, startCol).setFontWeight('bold');
        sheet.getRange(r, startCol, 1, numCols).setBackground('#EEEEEE');
      }
    });
    
    // Border
    sheet.getRange(startRow, startCol, rows.length + 2, numCols)
      .setBorder(true, true, true, true, true, true);
      
    return rows.length + 2; // Header + SubHeader + Data rows
  },

  _renderCalendarTable(sheet, startRow, startCol) {
    const events = this._getCalendarEvents();
    const title = 'Lịch sự kiện (Sắp tới)';
    const color = this.CONFIG.COLORS.CALENDAR;
    const numCols = 4; // Date, Event, Principal, Interest
    
    // Header
    sheet.getRange(startRow, startCol, 1, numCols).merge()
      .setValue(title)
      .setFontWeight('bold')
      .setBackground(color)
      .setFontColor('#FFFFFF')
      .setHorizontalAlignment('left');
      
    // Sub-header
    sheet.getRange(startRow + 1, startCol).setValue('Ngày').setFontWeight('bold');
    sheet.getRange(startRow + 1, startCol + 1).setValue('Sự kiện').setFontWeight('bold');
    sheet.getRange(startRow + 1, startCol + 2).setValue('Gốc').setFontWeight('bold');
    sheet.getRange(startRow + 1, startCol + 3).setValue('Lãi').setFontWeight('bold');
    
    // Data
    if (events.length === 0) {
        sheet.getRange(startRow + 2, startCol, 1, numCols).merge().setValue('Không có sự kiện sắp tới');
        sheet.getRange(startRow, startCol, 3, numCols).setBorder(true, true, true, true, true, true);
        return 3;
    }
    
    events.forEach((evt, idx) => {
      const r = startRow + 2 + idx;
      sheet.getRange(r, startCol).setValue(evt.date).setNumberFormat('dd/MM/yyyy');
      sheet.getRange(r, startCol + 1).setValue(evt.name);
      sheet.getRange(r, startCol + 2).setValue(evt.principal).setNumberFormat('#,##0');
      sheet.getRange(r, startCol + 3).setValue(evt.interest).setNumberFormat('#,##0');
    });
    
    // Border
    sheet.getRange(startRow, startCol, events.length + 2, numCols)
      .setBorder(true, true, true, true, true, true);
      
    return events.length + 2;
  },

  _getCalendarEvents() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const events = [];
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    
    // 1. Debt Payments (QUẢN LÝ NỢ)
    const debtSheet = ss.getSheetByName(APP_CONFIG.SHEETS.DEBT_MANAGEMENT);
    if (debtSheet) {
      const lastRow = debtSheet.getLastRow();
      if (lastRow >= 2) {
        // Col C (2): Type, Col H (7): Due Date, Col K (10): Remaining, Col L (11): Status
        // Note: Indices changed due to new column structure
        // A(0): STT, B(1): Name, C(2): Type, D(3): Principal, E(4): Rate, F(5): Term, G(6): Date, H(7): Maturity
        // I(8): PaidPrin, J(9): PaidInt, K(10): Remaining, L(11): Status
        const data = debtSheet.getRange(2, 1, lastRow - 1, 12).getValues();
        data.forEach(row => {
          const name = row[1];
          const type = row[2];
          const dueDate = row[7]; // H
          const remaining = row[10]; // K
          const status = row[11]; // L
          
          if (dueDate instanceof Date && dueDate >= today && status !== 'Đã thanh toán' && remaining > 0) {
            let principal = remaining;
            let interest = 0;
            
            // Simple logic for split based on Type
            if (type === 'Nợ ngân hàng') {
               // Bank loan: Pay interest monthly, principal at end.
               // Assuming 'remaining' is Principal.
               // Interest = Principal * Rate / 12.
               // But we don't have Rate here easily without parsing.
               // Let's assume the "Event" is the monthly payment?
               // Or is it the Final Due Date?
               // The code checks `dueDate` (Maturity).
               // If Maturity is coming up, it's likely the Principal payment.
               principal = remaining;
               interest = 0; 
            }
            
            events.push({
              date: dueDate,
              name: `Trả nợ: ${name}`,
              principal: principal,
              interest: interest
            });
          }
        });
      }
    }
    
    // 2. Lending Collections (CHO VAY)
    const lendingSheet = ss.getSheetByName(APP_CONFIG.SHEETS.LENDING);
    if (lendingSheet) {
      const lastRow = lendingSheet.getLastRow();
      if (lastRow >= 2) {
        // Similar structure
        const data = lendingSheet.getRange(2, 1, lastRow - 1, 12).getValues();
        data.forEach(row => {
          const name = row[1];
          const dueDate = row[7]; // H
          const remaining = row[10]; // K
          const status = row[11]; // L
          
          if (dueDate instanceof Date && dueDate >= today && status === 'Đang vay' && remaining > 0) {
            events.push({
              date: dueDate,
              name: `Thu nợ: ${name}`,
              principal: remaining,
              interest: 0
            });
          }
        });
      }
    }
    
    // Sort by Date
    events.sort((a, b) => a.date - b.date);
    
    // Limit to top 10
    return events.slice(0, 10);
  },
  
  _getDebtItems() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const debtSheet = ss.getSheetByName(APP_CONFIG.SHEETS.DEBT_MANAGEMENT);
    if (!debtSheet) return [];
    
    const lastRow = debtSheet.getLastRow();
    if (lastRow < 2) return [];
    
    // Col B (Index 1): Name, Col L (Index 11): Status (New Structure)
    const data = debtSheet.getRange(2, 1, lastRow - 1, 12).getValues();
    // Filter debts that are NOT "Đã thanh toán"
    const debts = data
      .filter(row => row[11] !== 'Đã thanh toán' && row[1] !== '')
      .map(row => ({ name: row[1] }));
      
    return debts.length > 0 ? debts : [{ name: 'Không có khoản nợ' }];
  },
  
  _createDynamicSumFormula(sheetName, col, criteria, criteriaCol) {
    let base = `INDIRECT("'"&"${sheetName}"&"'!${col}:${col}")`;
    let dateCol = `INDIRECT("'"&"${sheetName}"&"'!B:B")`;
    
    let criteriaStr = "";
    if (criteria && criteriaCol) {
      criteriaStr = `, INDIRECT("'"&"${sheetName}"&"'!${criteriaCol}:${criteriaCol}"), "${criteria}"`;
    }
    
    return `IFERROR(
      IF(OR($B$2="", $B$3="", $B$4=""), 0,
        IF($B$2="Tất cả",
          SUMIFS(${base}${criteriaStr}, ${dateCol}, ">0"),
          IF($B$3="Tất cả",
            SUMIFS(${base}${criteriaStr}, 
                   ${dateCol}, ">="&DATE(VALUE($B$2),1,1), 
                   ${dateCol}, "<"&DATE(VALUE($B$2)+1,1,1)),
            IF($B$4="Tất cả",
              SUMIFS(${base}${criteriaStr}, 
                     ${dateCol}, ">="&DATE(VALUE($B$2),(VALUE(RIGHT($B$3,1))-1)*3+1,1), 
                     ${dateCol}, "<"&DATE(VALUE($B$2),(VALUE(RIGHT($B$3,1))-1)*3+4,1)),
              SUMIFS(${base}${criteriaStr}, 
                     ${dateCol}, ">="&DATE(VALUE($B$2),VALUE(RIGHT($B$4,LEN($B$4)-6)),1), 
                     ${dateCol}, "<"&DATE(VALUE($B$2),VALUE(RIGHT($B$4,LEN($B$4)-6))+1,1))
            )
          )
        )
      ), 0)`;
  },
  
  _setupTable2(sheet, startRow) {
    const currentYear = new Date().getFullYear();
    
    // Title
    sheet.getRange(startRow, 1, 1, 10).merge()
      .setValue(`📈 THỐNG KÊ TÀI CHÍNH GIA ĐÌNH NĂM ${currentYear}`)
      .setFontSize(12)
      .setFontWeight('bold')
      .setBackground(this.CONFIG.COLORS.HEADER)
      .setFontColor('#FFFFFF');
    sheet.setRowHeight(startRow, 35);
    
    const headerRow = startRow + 1;
    const dataStart = startRow + 2;
    
    const headers = ['Kỳ', 'Thu', 'Chi', 'Nợ (Gốc)', 'Lãi', 'CK', 'Vàng', 'Crypto', 'ĐT khác', 'Dòng tiền'];
    sheet.getRange(headerRow, 1, 1, 10).setValues([headers])
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBackground(APP_CONFIG.COLORS.HEADER_BG)
      .setFontColor(APP_CONFIG.COLORS.HEADER_TEXT);
      
    // Data
    for (let month = 1; month <= 12; month++) {
      const row = dataStart + month - 1;
      sheet.getRange(row, 1).setValue(`Tháng ${month}`);
      
      sheet.getRange(row, 2).setFormula(`=SUMIFS(THU!C:C, THU!B:B, ">="&DATE(${currentYear},${month},1), THU!B:B, "<"&DATE(${currentYear},${month}+1,1))`);
      sheet.getRange(row, 3).setFormula(`=SUMIFS(CHI!C:C, CHI!B:B, ">="&DATE(${currentYear},${month},1), CHI!B:B, "<"&DATE(${currentYear},${month}+1,1))`);
      sheet.getRange(row, 4).setFormula(`=SUMIFS('TRẢ NỢ'!D:D, 'TRẢ NỢ'!B:B, ">="&DATE(${currentYear},${month},1), 'TRẢ NỢ'!B:B, "<"&DATE(${currentYear},${month}+1,1))`);
      sheet.getRange(row, 5).setFormula(`=SUMIFS('TRẢ NỢ'!E:E, 'TRẢ NỢ'!B:B, ">="&DATE(${currentYear},${month},1), 'TRẢ NỢ'!B:B, "<"&DATE(${currentYear},${month}+1,1))`);
      sheet.getRange(row, 6).setFormula(`=SUMIFS('CHỨNG KHOÁN'!H:H, 'CHỨNG KHOÁN'!C:C, "Mua", 'CHỨNG KHOÁN'!B:B, ">="&DATE(${currentYear},${month},1), 'CHỨNG KHOÁN'!B:B, "<"&DATE(${currentYear},${month}+1,1))`);
      sheet.getRange(row, 7).setFormula(`=SUMIFS(VÀNG!H:H, VÀNG!C:C, "Mua", VÀNG!B:B, ">="&DATE(${currentYear},${month},1), VÀNG!B:B, "<"&DATE(${currentYear},${month}+1,1))`);
      sheet.getRange(row, 8).setFormula(`=SUMIFS(CRYPTO!I:I, CRYPTO!C:C, "Mua", CRYPTO!B:B, ">="&DATE(${currentYear},${month},1), CRYPTO!B:B, "<"&DATE(${currentYear},${month}+1,1))`);
      sheet.getRange(row, 9).setFormula(`=SUMIFS('ĐẦU TƯ KHÁC'!D:D, 'ĐẦU TƯ KHÁC'!B:B, ">="&DATE(${currentYear},${month},1), 'ĐẦU TƯ KHÁC'!B:B, "<"&DATE(${currentYear},${month}+1,1))`);
      
      const b = `B${row}`, c = `C${row}`, d = `D${row}`, e = `E${row}`;
      const f = `F${row}`, g = `G${row}`, h = `H${row}`, i = `I${row}`;
      sheet.getRange(row, 10).setFormula(`=IFERROR(${b}-(${c}+${d}+${e}+${f}+${g}+${h}+${i}), 0)`);
    }
    
    // Summary
    const summaryRow = dataStart + 12;
    sheet.getRange(summaryRow, 1).setValue('Tổng').setFontWeight('bold');
    for (let col = 2; col <= 10; col++) {
      const colLetter = this._getColumnLetter(col);
      sheet.getRange(summaryRow, col).setFormula(`=SUM(${colLetter}${dataStart}:${colLetter}${dataStart + 11})`).setFontWeight('bold').setBackground('#FFE0B2');
    }
    
    sheet.getRange(dataStart, 2, 13, 9).setNumberFormat('#,##0');
    sheet.getRange(headerRow, 1, 14, 10).setBorder(true, true, true, true, true, true);
  },
  
  _createChart(sheet) {
    // Create hidden data range for Chart
    // Z1:AA5 (Removed AB - Style column as it's not supported natively in Sheets Charts this way)
    const chartDataRange = sheet.getRange('Z1:AA5');
    chartDataRange.setValues([
      ['Danh mục', 'Giá trị'],
      ['Thu nhập', 0],
      ['Chi phí', 0],
      ['Tài sản', 0],
      ['Nợ', 0]
    ]);
    
    // Link formulas to hidden data
    // Note: We need to find the Total cells dynamically or use named ranges. 
    // For simplicity, we'll re-calculate totals in the hidden area using the same logic.
    
    // Thu nhập (Total Income)
    sheet.getRange('AA2').setFormula('=' + this._createDynamicSumFormula('THU', 'C'));
    
    // Chi phí (Total Expense)
    // Chi phí = Chi tiêu (CHI) + Trả nợ (TRẢ NỢ)
    const chiFormula = this._createDynamicSumFormula('CHI', 'C');
    const traNoGoc = this._createDynamicSumFormula('TRẢ NỢ', 'D');
    const traNoLai = this._createDynamicSumFormula('TRẢ NỢ', 'E');
    sheet.getRange('AA3').setFormula('=' + `${chiFormula} + ${traNoGoc} + ${traNoLai}`);
    
    // Tài sản (Total Assets)
    // Tài sản = (Thu - Chi - Đầu tư + Thoái vốn) + Giá trị hiện tại các khoản đầu tư
    // Tuy nhiên, để đơn giản và chính xác theo bảng Tài sản, ta sẽ lấy tổng các khoản đầu tư hiện tại + Tiền mặt ròng
    
    // 1. Tiền mặt ròng = Thu - Chi - Đầu tư + Thoái vốn
    const totalIncome = this._createDynamicSumFormula('THU', 'C');
    const totalExpense = `(${chiFormula} + ${traNoGoc} + ${traNoLai})`;
    
    // Đầu tư (Flow out)
    const investCK = this._createDynamicSumFormula('CHỨNG KHOÁN', 'H', 'Mua', 'C');
    const investGold = this._createDynamicSumFormula('VÀNG', 'H', 'Mua', 'C');
    const investCrypto = this._createDynamicSumFormula('CRYPTO', 'I', 'Mua', 'C');
    const investOther = this._createDynamicSumFormula('ĐẦU TƯ KHÁC', 'D');
    const totalInvest = `(${investCK} + ${investGold} + ${investCrypto} + ${investOther})`;
    
    // Thoái vốn (Flow in)
    const divestCK = this._createDynamicSumFormula('CHỨNG KHOÁN', 'H', 'Bán', 'C');
    const divestGold = this._createDynamicSumFormula('VÀNG', 'H', 'Bán', 'C');
    const divestCrypto = this._createDynamicSumFormula('CRYPTO', 'I', 'Bán', 'C');
    const totalDivest = `(${divestCK} + ${divestGold} + ${divestCrypto})`;
    
    const netCash = `(${totalIncome} - ${totalExpense} - ${totalInvest} + ${totalDivest})`;
    
    // 2. Giá trị hiện tại (Current Value) - Cái này thường là Stock, không phụ thuộc dòng tiền quá khứ, 
    // nhưng trong Dashboard lọc theo thời gian, ta có thể chỉ hiển thị dòng tiền ròng tích lũy trong kỳ đó?
    // KHÔNG, Tài sản là Stock (Tích lũy). Nếu lọc theo năm 2024, ta muốn xem Tài sản tăng thêm trong năm 2024 hay Tổng tài sản tại thời điểm đó?
    // Thông thường Dashboard Cashflow theo kỳ sẽ hiển thị Dòng tiền (Income/Expense) và Tài sản Tăng thêm (Net Asset Change).
    // Tuy nhiên, người dùng thường muốn xem Tổng tài sản hiện tại.
    // Nhưng nếu lọc "Tháng 11", mà hiển thị Tổng tài sản tích lũy cả đời thì không khớp.
    // Để thống nhất với logic "Dòng tiền", ta sẽ hiển thị "Tài sản ròng tăng thêm trong kỳ".
    // Hoặc nếu muốn hiển thị Tổng tài sản, ta phải bỏ qua bộ lọc thời gian cho phần Tài sản.
    // Theo yêu cầu "tương quan khi lọc", ta sẽ tính dòng tiền ròng (Net Worth Change) trong kỳ.
    
    // Tuy nhiên, bảng Assets bên dưới lại tính: Tiền mặt ròng (trong kỳ) + Giá trị hiện tại (lấy SUM cột M - Giá trị hiện tại).
    // Cột M của Chứng khoán là (Số lượng * Giá thị trường). Nó không có ngày tháng.
    // Vì vậy, bảng Assets đang hiển thị "Giá trị hiện tại" bất kể bộ lọc thời gian cho phần đầu tư.
    // Chỉ có phần "Tiền mặt ròng" là bị ảnh hưởng bởi bộ lọc.
    
    // Chúng ta sẽ giữ nguyên logic của Bảng Assets:
    // Tiền mặt ròng (theo kỳ) + Giá trị đầu tư (hiện tại - không lọc ngày, hoặc lọc theo ngày mua?)
    // Nếu lọc theo ngày mua thì không ra giá trị hiện tại.
    // Quyết định: Phần Đầu tư giữ nguyên (All time), phần Tiền mặt ròng theo bộ lọc.
    
    // Fix lại công thức Chart cho khớp với Bảng Assets:
    
    // Giá trị hiện tại (All time)
    const currentValCK = `IFERROR(SUM('CHỨNG KHOÁN'!M:M), 0)`;
    // Vàng/Crypto tính theo luỹ kế mua - bán (vì không có cột giá trị thị trường tự động cập nhật realtime trong sheet này, trừ khi có API)
    // Trong sheet VÀNG, không có cột Giá trị hiện tại, chỉ có Mua/Bán. Ta lấy (Mua - Bán) * Giá hiện tại? Không có giá hiện tại.
    // Ta tạm tính theo giá vốn: Sum(Mua) - Sum(Bán).
    const currentValGold = `(SUMIF(VÀNG!C:C, "Mua", VÀNG!H:H) - SUMIF(VÀNG!C:C, "Bán", VÀNG!H:H))`;
    const currentValCrypto = `(SUMIF(CRYPTO!C:C, "Mua", CRYPTO!I:I) - SUMIF(CRYPTO!C:C, "Bán", CRYPTO!I:I))`;
    const currentValOther = `SUM('ĐẦU TƯ KHÁC'!D:D)`; // Giả sử chưa thu về
    
    // Tổng hợp lại cho Chart
    sheet.getRange('AA4').setFormula('=' + `IFERROR(${netCash} + ${currentValCK} + ${currentValGold} + ${currentValCrypto} + ${currentValOther}, 0)`);
    
    // Nợ (Total Liabilities)
    // Nợ cũng là Stock (Tích lũy). Lọc theo thời gian cho Nợ nghĩa là gì? "Nợ phát sinh trong kỳ"?
    // Bảng Liabilities đang dùng: SUMIFS('QUẢN LÝ NỢ'!J:J, 'QUẢN LÝ NỢ'!B:B, name) -> Cột J là "Còn nợ".
    // Nó KHÔNG lọc theo ngày tháng trong Bảng Liabilities (xem dòng 164: formula không có date criteria).
    // Vậy Chart cũng không nên lọc theo ngày tháng cho Nợ.
    sheet.getRange('AA5').setFormula(`=SUM('QUẢN LÝ NỢ'!J:J)`);
    
    // Create Chart
    const chart = sheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(chartDataRange)
      .setPosition(33, 11, 0, 0) // Row 33, Col K
      .setOption('title', 'Tổng quan Tài chính')
      .setOption('width', 600)
      .setOption('height', 400)
      .setOption('legend', { position: 'none' })
      .setOption('useFirstColumnAsDomain', true)
      .setNumHeaders(1)
      .build();
      
    sheet.insertChart(chart);
    sheet.hideColumns(26, 2); // Hide Z, AA
  },
  
  _formatSheet(sheet) {
    // A, B, C: Income / Liabilities (3 cols)
    sheet.setColumnWidth(1, 200); // A: Name
    sheet.setColumnWidth(2, 120); // B: Value
    sheet.setColumnWidth(3, 60);  // C: %
    
    sheet.setColumnWidth(4, 50);  // D: Spacer
    
    // E, F, G, H, I: Expense / Assets
    // Expense: E, F, G (Name, Value, %)
    // Assets: E, F, G, H, I (Name, Cap, P/L, Cur, %)
    sheet.setColumnWidth(5, 200); // E: Name
    sheet.setColumnWidth(6, 120); // F: Value / Capital
    sheet.setColumnWidth(7, 100); // G: % / P/L
    sheet.setColumnWidth(8, 120); // H: Current Val
    sheet.setColumnWidth(9, 60);  // I: % (Assets)
    
    sheet.setColumnWidth(10, 50); // J: Spacer
    
    // K, L, M, N: Calendar / Chart
    sheet.setColumnWidth(11, 100); // K: Date
    sheet.setColumnWidth(12, 200); // L: Event
    sheet.setColumnWidth(13, 100); // M: Principal
    sheet.setColumnWidth(14, 100); // N: Interest
    
    sheet.setFrozenRows(1);
  },
  
  _setupTriggers() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'onDashboardEdit') {
        ScriptApp.deleteTrigger(trigger);
      }
    });
    
    ScriptApp.newTrigger('onDashboardEdit')
      .forSpreadsheet(ss)
      .onEdit()
      .create();
  },
  
  refreshDashboard() {
    this.setupDashboard();
  },
  
  _getColumnLetter(colNum) {
    let letter = '';
    while (colNum > 0) {
      const remainder = (colNum - 1) % 26;
      letter = String.fromCharCode(65 + remainder) + letter;
      colNum = Math.floor((colNum - 1) / 26);
    }
    return letter;
  }
};

function onDashboardEdit(e) {
  try {
    if (!e) return;
    const range = e.range;
    const sheet = range.getSheet();
    if (sheet.getName() === APP_CONFIG.SHEETS.DASHBOARD) {
      const row = range.getRow();
      const col = range.getColumn();
      if (col === 2 && (row === 2 || row === 3 || row === 4)) {
        Utilities.sleep(100);
        SpreadsheetApp.flush();
      }
    }
  } catch (error) {
    Logger.log('❌ Lỗi onDashboardEdit: ' + error.message);
  }
}
