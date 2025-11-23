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
    
    // Fetch events once - REMOVED (Using Custom Functions)
    // const events = this._getCalendarEvents();
    
    // === ROW 1: INCOME (Left) & EXPENSE (Middle) & PAYABLES (Right) ===
    
    // 1. Income Table (3 Cols: Name, Value, %)
    const incomeCategories = APP_CONFIG.CATEGORIES.INCOME;
    const incomeRows = [...incomeCategories, 'TỔNG THU NHẬP'];
    const incomeHeight = this._renderTable(sheet, currentRow, cfg.LEFT_COL, '1. Báo cáo Thu nhập', this.CONFIG.COLORS.INCOME, incomeRows, 3, true);
    
    // Formulas for Income
    const incStart = currentRow + 2;
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
    
    // 2. Expense Table (5 Cols: Name, Spent, Budget, Remaining, Status)
    const expenseCategories = APP_CONFIG.CATEGORIES.EXPENSE;
    const expenseRows = [...expenseCategories, 'Trả nợ (Gốc + Lãi)', 'TỔNG CHI PHÍ'];
    const expenseHeight = this._renderExpenseTable(sheet, currentRow, cfg.RIGHT_COL, '2. Báo cáo Chi phí', this.CONFIG.COLORS.EXPENSE, expenseRows);
    
    // Formulas for Expense are handled inside _renderExpenseTable now
    
    // 3. Payables Table (Right - Col K)
    const payablesHeight = this._renderPayables(sheet, currentRow, cfg.CALENDAR_COL);

    // Calculate max height for Row 1
    const row1Height = Math.max(incomeHeight, expenseHeight, payablesHeight);
    currentRow += row1Height + 2; // +2 padding
    
    // === ROW 2: LIABILITIES (Left) & ASSETS (Middle) & RECEIVABLES (Right) ===
    
    // 4. Liabilities Table (3 Cols: Name, Value, %)
    const debtItems = this._getDebtItems();
    const liabilityRows = [...debtItems.map(d => d.name), 'TỔNG NỢ'];
    const liabilityHeight = this._renderTable(sheet, currentRow, cfg.LEFT_COL, '3. Báo cáo Nợ phải trả', this.CONFIG.COLORS.LIABILITIES, liabilityRows, 3, true);
    
    // Formulas for Liabilities
    const liabStart = currentRow + 2;
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
    
    // Rows for Assets
    const assetStart = assetHeaderRow + 2;
    const assetTotalRow = assetStart + 6; // 6 items
    
    // 1. Cash (Net)
    sheet.getRange(assetStart, cfg.RIGHT_COL).setValue('Tiền mặt (Ròng)');
    sheet.getRange(assetStart, cfg.RIGHT_COL + 1).setValue('-'); // Capital N/A
    sheet.getRange(assetStart, cfg.RIGHT_COL + 2).setValue('-'); // P/L N/A
    sheet.getRange(assetStart, cfg.RIGHT_COL + 3).setFormula('=' + `${this._createDynamicSumFormula('THU', 'C')} - ${this._createDynamicSumFormula('CHI', 'C')}`);
    
    // 2. Stock
    sheet.getRange(assetStart + 1, cfg.RIGHT_COL).setValue('Chứng khoán');
    sheet.getRange(assetStart + 1, cfg.RIGHT_COL + 1).setFormula(`=IFERROR(SUM('CHỨNG KHOÁN'!J:J), 0)`); // Total Cost
    sheet.getRange(assetStart + 1, cfg.RIGHT_COL + 2).setFormula(`=IFERROR(R[0]C[1] - R[0]C[-1], 0)`); // Current - Cost
    sheet.getRange(assetStart + 1, cfg.RIGHT_COL + 3).setFormula(`=IFERROR(SUM('CHỨNG KHOÁN'!M:M), 0)`); // Market Value
    
    // 3. Gold
    sheet.getRange(assetStart + 2, cfg.RIGHT_COL).setValue('Vàng');
    sheet.getRange(assetStart + 2, cfg.RIGHT_COL + 1).setFormula(`=IFERROR(SUM('VÀNG'!J:J), 0)`); // Total Cost
    sheet.getRange(assetStart + 2, cfg.RIGHT_COL + 2).setFormula(`=IFERROR(R[0]C[1] - R[0]C[-1], 0)`);
    sheet.getRange(assetStart + 2, cfg.RIGHT_COL + 3).setFormula(`=IFERROR(SUM('VÀNG'!M:M), 0)`); // Market Value
    
    // 4. Crypto
    sheet.getRange(assetStart + 3, cfg.RIGHT_COL).setValue('Crypto');
    sheet.getRange(assetStart + 3, cfg.RIGHT_COL + 1).setFormula(`=IFERROR(SUM('CRYPTO'!L:L), 0)`); // Total Cost VND
    sheet.getRange(assetStart + 3, cfg.RIGHT_COL + 2).setFormula(`=IFERROR(R[0]C[1] - R[0]C[-1], 0)`);
    sheet.getRange(assetStart + 3, cfg.RIGHT_COL + 3).setFormula(`=IFERROR(SUM('CRYPTO'!N:N), 0)`); // Market Value VND
    
    // 5. Other Investment
    sheet.getRange(assetStart + 4, cfg.RIGHT_COL).setValue('Đầu tư khác');
    sheet.getRange(assetStart + 4, cfg.RIGHT_COL + 1).setFormula(`=IFERROR(SUM('ĐẦU TƯ KHÁC'!D:D), 0)`); // Capital
    sheet.getRange(assetStart + 4, cfg.RIGHT_COL + 2).setFormula(`=IFERROR(SUM('ĐẦU TƯ KHÁC'!H:H), 0)`); // Profit
    sheet.getRange(assetStart + 4, cfg.RIGHT_COL + 3).setFormula(`=IFERROR(R[0]C[-2] + R[0]C[-1], 0)`); // Capital + Profit
    
    // 6. Lending (Receivables)
    sheet.getRange(assetStart + 5, cfg.RIGHT_COL).setValue('Cho vay');
    sheet.getRange(assetStart + 5, cfg.RIGHT_COL + 1).setFormula(`=IFERROR(SUM('CHO VAY'!D:D), 0)`); // Principal
    sheet.getRange(assetStart + 5, cfg.RIGHT_COL + 2).setFormula(`=IFERROR(SUM('CHO VAY'!J:J), 0)`); // Interest Collected
    sheet.getRange(assetStart + 5, cfg.RIGHT_COL + 3).setFormula(`=IFERROR(SUM('CHO VAY'!K:K), 0)`); // Remaining Principal
    
    // Total Assets
    sheet.getRange(assetTotalRow, cfg.RIGHT_COL).setValue('TỔNG TÀI SẢN').setFontWeight('bold');
    sheet.getRange(assetTotalRow, cfg.RIGHT_COL + 1).setFormula(`=SUM(R[-6]C:R[-1]C)`);
    sheet.getRange(assetTotalRow, cfg.RIGHT_COL + 2).setFormula(`=SUM(R[-6]C:R[-1]C)`);
    sheet.getRange(assetTotalRow, cfg.RIGHT_COL + 3).setFormula(`=SUM(R[-6]C:R[-1]C)`);
    
    // % Column
    for (let i = 0; i < 6; i++) {
      sheet.getRange(assetStart + i, cfg.RIGHT_COL + 4).setFormula(`=IFERROR(R[0]C[-1] / R${assetTotalRow}C[-1], 0)`);
    }
    sheet.getRange(assetTotalRow, cfg.RIGHT_COL + 4).setValue(1).setNumberFormat('0%');
    
    sheet.getRange(assetHeaderRow, cfg.RIGHT_COL, 9, 5).setBorder(true, true, true, true, true, true, '#B0B0B0', SpreadsheetApp.BorderStyle.SOLID);
    
    const assetHeight = 9; // Fixed height for Assets table
    
    // 6. Receivables Table (Right - Col K, Row 2)
    const receivablesHeight = this._renderReceivables(sheet, currentRow, cfg.CALENDAR_COL);

    // Calculate max height for Row 2
    const row2Height = Math.max(liabilityHeight, assetHeight, receivablesHeight);
    
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
    // Border - Lighter Color
    sheet.getRange(startRow, startCol, rows.length + 2, numCols)
      .setBorder(true, true, true, true, true, true, '#B0B0B0', SpreadsheetApp.BorderStyle.SOLID);
      
    return rows.length + 2; // Header + SubHeader + Data rows
  },

  _renderExpenseTable(sheet, startRow, startCol, title, color, rows) {
    const numCols = 5; // Name, Spent, Budget, Remaining, Status
    
    // Header
    sheet.getRange(startRow, startCol, 1, numCols).merge()
      .setValue(title)
      .setFontWeight('bold')
      .setBackground(color)
      .setFontColor('#FFFFFF')
      .setHorizontalAlignment('left');
      
    // Sub-headers (Row 2)
    const subHeaderRow = startRow + 1;
    const headers = ['Danh mục', 'Đã chi', 'Ngân sách', 'Còn lại', 'Trạng thái'];
    sheet.getRange(subHeaderRow, startCol, 1, numCols).setValues([headers]).setFontWeight('bold');
      
    // Rows
    const dataStart = startRow + 2;
    const expenseCategories = APP_CONFIG.CATEGORIES.EXPENSE;
    const totalRowIdx = dataStart + rows.length - 1;
    
    rows.forEach((name, idx) => {
      const r = dataStart + idx;
      sheet.getRange(r, startCol).setValue(name);
      
      // 1. Đã chi (Spent)
      if (idx < expenseCategories.length) {
        // Normal categories
        sheet.getRange(r, startCol + 1).setFormula('=' + this._createDynamicSumFormula('CHI', 'C', name, 'D'));
      } else if (name === 'Trả nợ (Gốc + Lãi)') {
        // Debt
        sheet.getRange(r, startCol + 1).setFormula('=' + `${this._createDynamicSumFormula('TRẢ NỢ', 'D')} + ${this._createDynamicSumFormula('TRẢ NỢ', 'E')}`);
      } else {
        // Total
        sheet.getRange(r, startCol + 1).setFormula(`=SUM(R[-${rows.length - 1}]C:R[-1]C)`);
      }
      
      // 2. Ngân sách (Budget)
      if (name === 'TỔNG CHI PHÍ') {
         sheet.getRange(r, startCol + 2).setFormula(`=SUM(R[-${rows.length - 1}]C:R[-1]C)`);
      } else {
         // VLOOKUP from BUDGET sheet. Range A:C. Col 3 is Budget.
         // IF name is "Trả nợ (Gốc + Lãi)", map to "Trả nợ gốc" + "Trả lãi" in Budget?
         // Budget sheet has "Trả nợ gốc" and "Trả lãi" separate.
         // For simplicity, we will try to VLOOKUP the name directly.
         // Note: "Trả nợ (Gốc + Lãi)" won't match directly. We need to handle it.
         
         if (name === 'Trả nợ (Gốc + Lãi)') {
           // Sum Budget of "Trả nợ gốc" and "Trả lãi"
           sheet.getRange(r, startCol + 2).setFormula(`=IFERROR(VLOOKUP("Trả nợ gốc", BUDGET!A:C, 3, 0), 0) + IFERROR(VLOOKUP("Trả lãi", BUDGET!A:C, 3, 0), 0)`);
         } else {
           sheet.getRange(r, startCol + 2).setFormula(`=IFERROR(VLOOKUP("${name}", BUDGET!A:C, 3, 0), 0)`);
         }
      }
      
      // 3. Còn lại (Remaining) = Budget - Spent
      sheet.getRange(r, startCol + 3).setFormula(`=R[0]C[-1] - R[0]C[-2]`);
      
      // 4. Trạng thái (Status)
      // If Spent > Budget -> "Vượt" (Red)
      // If Spent > 80% Budget -> "Sắp hết" (Yellow)
      // Else -> "Trong hạn mức" (Green)
      // Skip for Total row if needed, but useful there too.
      
      const statusFormula = `=IF(R[0]C[-2]=0, "Chưa có NS", IF(R[0]C[-3] > R[0]C[-2], "Vượt ngân sách", IF(R[0]C[-3] > 0.8 * R[0]C[-2], "Sắp hết", "Trong hạn mức")))`;
      sheet.getRange(r, startCol + 4).setFormula(statusFormula);
      
      // Last row styling
      if (idx === rows.length - 1) {
        sheet.getRange(r, startCol).setFontWeight('bold');
        sheet.getRange(r, startCol, 1, numCols).setBackground('#EEEEEE');
      }
    });
    
    // Formatting
    // Number format for Spent, Budget, Remaining
    sheet.getRange(dataStart, startCol + 1, rows.length, 3).setNumberFormat('#,##0');
    
    // Conditional Formatting for Status
    const statusRange = sheet.getRange(dataStart, startCol + 4, rows.length, 1);
    
    // Red - Vượt
    const ruleRed = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Vượt ngân sách')
      .setBackground('#FFEBEE')
      .setFontColor('#C62828')
      .setRanges([statusRange])
      .build();
      
    // Yellow - Sắp hết
    const ruleYellow = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Sắp hết')
      .setBackground('#FFF3E0')
      .setFontColor('#EF6C00')
      .setRanges([statusRange])
      .build();
      
    // Green - Trong hạn mức
    const ruleGreen = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Trong hạn mức')
      .setBackground('#E8F5E9')
      .setFontColor('#2E7D32')
      .setRanges([statusRange])
      .build();
      
    // Add rules to sheet
    const rules = sheet.getConditionalFormatRules();
    rules.push(ruleRed, ruleYellow, ruleGreen);
    sheet.setConditionalFormatRules(rules);
    
    // Border
    sheet.getRange(startRow, startCol, rows.length + 2, numCols)
      .setBorder(true, true, true, true, true, true, '#B0B0B0', SpreadsheetApp.BorderStyle.SOLID);
      
    return rows.length + 2;

  _renderPayables(sheet, startRow, startCol) {
    return this._renderEventTable(sheet, startRow, startCol, '📅 Lịch sự kiện: KHOẢN PHẢI TRẢ (Sắp tới)', this.CONFIG.COLORS.CALENDAR, 'AccPayable', 'QUẢN LÝ NỢ');
  },

  _renderReceivables(sheet, startRow, startCol) {
    return this._renderEventTable(sheet, startRow, startCol, '📅 Lịch sự kiện: KHOẢN PHẢI THU (Sắp tới)', '#70AD47', 'AccReceivable', 'CHO VAY');
  },

  _renderEventTable(sheet, startRow, startCol, title, color, functionName, sourceSheet) {
    const numCols = 6; // Date, Action, Event, Remaining, Principal, Interest
    
    // Header
    sheet.getRange(startRow, startCol, 1, numCols).merge()
      .setValue(title)
      .setFontWeight('bold')
      .setBackground(color)
      .setFontColor('#FFFFFF')
      .setHorizontalAlignment('left');
      
    // Sub-header
    const headers = ['Ngày', 'Hành động', 'Sự kiện', 'Gốc còn lại', 'Gốc trả kỳ này', 'Lãi trả kỳ này'];
    sheet.getRange(startRow + 1, startCol, 1, numCols).setValues([headers]).setFontWeight('bold');
    
    // Formula
    const formulaRow = startRow + 2;
    sheet.getRange(formulaRow, startCol).setFormula(`=${functionName}('${sourceSheet}'!A2:L)`);
    
    // Format columns (Assuming max 12 rows returned)
    const dataRange = sheet.getRange(formulaRow, startCol, 12, numCols);
    dataRange.setBorder(true, true, true, true, true, true, '#B0B0B0', SpreadsheetApp.BorderStyle.SOLID);
    
    // Format Date Column
    sheet.getRange(formulaRow, startCol, 12, 1).setNumberFormat('dd/MM/yyyy');
    // Format Numbers
    sheet.getRange(formulaRow, startCol + 3, 12, 3).setNumberFormat('#,##0');
    
    return 14; // Header + SubHeader + 12 Data Rows
  },

  _getCalendarEvents() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const payables = [];
    const receivables = [];
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    
    // Helper to calculate days between dates
    const getDaysDiff = (d1, d2) => {
      if (!d1 || !d2) return 0;
      const diffTime = Math.abs(d2.getTime() - d1.getTime());
      return Math.ceil(diffTime / (1000 * 60 * 60 * 24)); 
    };
    
    // Helper to parse date safely
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

    // Helper to parse currency safely
    const parseCurrency = (val) => {
      if (typeof val === 'number') return val;
      if (typeof val === 'string') {
        return parseFloat(val.replace(/\D/g, ''));
      }
      return 0;
    };
    
    // Helper to process installments
    const processInstallments = (sheetName, isDebt) => {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) return;
      
      const lastRow = sheet.getLastRow();
      if (lastRow < 2) return;
      
      // A(0): STT, B(1): Name, C(2): Type, D(3): Principal, E(4): Rate, F(5): Term, G(6): Date, H(7): Maturity
      // I(8): PaidPrin, J(9): PaidInt, K(10): Remaining, L(11): Status
      const data = sheet.getRange(2, 1, lastRow - 1, 12).getValues();
      
      data.forEach(row => {
        const name = row[1];
        const type = row[2]; // Lending Type
        const initialPrincipal = parseCurrency(row[3]);
        const rate = parseFloat(row[4]) || 0; // Annual Rate (decimal)
        const term = parseInt(row[5]) || 1;
        const rawStartDate = row[6];
        const rawMaturityDate = row[7];
        const remaining = parseCurrency(row[10]);
        const status = row[11];
        
        const startDate = parseDate(rawStartDate);
        const maturityDate = parseDate(rawMaturityDate);
        
        // Check active status
        let isActive = false;
        if (isDebt) {
          isActive = (status === 'Chưa trả' || status === 'Đang trả');
        } else {
          isActive = (status === 'Đang vay');
        }
        
        if (isActive && remaining > 0 && startDate) {
          const targetList = isDebt ? payables : receivables;
          
          // LOGIC THEO TỪNG LOẠI HÌNH
          
          // 1. Tất toán gốc - lãi cuối kỳ
          if (type === 'Tất toán gốc - lãi cuối kỳ') {
            if (maturityDate && maturityDate >= today) {
               const days = getDaysDiff(startDate, maturityDate);
               // Lãi = Ngày * Gốc * Tỷ lệ / 365
               let currentInterest = days * (rate * remaining) / 365;
               
               targetList.push({
                  date: maturityDate,
                  action: isDebt ? 'Phải trả' : 'Phải thu',
                  name: isDebt ? `${name} (Tất toán)` : `${name} (Đáo hạn)`,
                  remaining: remaining,
                  principalPayment: remaining,
                  interestPayment: currentInterest
               });
            }
          }
          
          // 2. Trả lãi hàng tháng, gốc cuối kỳ (Bao gồm "Nợ ngân hàng")
          else if (type === 'Trả lãi hàng tháng, gốc cuối kỳ' || type === 'Nợ ngân hàng') {
             // Lặp qua từng tháng để tìm kỳ trả lãi tiếp theo
             for (let i = 1; i <= term; i++) {
                let payDate = new Date(startDate);
                payDate.setMonth(payDate.getMonth() + i);
                
                if (payDate >= today) {
                   // Tính số ngày của tháng trước đó
                   let prevDate = new Date(startDate);
                   prevDate.setMonth(prevDate.getMonth() + i - 1);
                   const daysInMonth = getDaysDiff(prevDate, payDate);
                   
                   // Lãi tháng = Số ngày * Gốc * Tỷ lệ / 365
                   let monthlyInterest = daysInMonth * (rate * remaining) / 365;
                   
                   // Nếu là kỳ cuối cùng thì trả cả gốc
                   let principalPay = (i === term) ? remaining : 0;
                   
                   targetList.push({
                      date: payDate,
                      action: isDebt ? 'Phải trả' : 'Phải thu',
                      name: isDebt ? `${name} (Kỳ ${i}/${term})` : `${name} (Kỳ ${i}/${term})`,
                      remaining: remaining,
                      principalPayment: principalPay,
                      interestPayment: monthlyInterest
                   });
                   
                   // Chỉ hiển thị 1 kỳ tiếp theo cho mỗi khoản vay để tránh spam lịch
                   break; 
                }
             }
          }
          
          // 3. Trả góp gốc - lãi hàng tháng (Gốc đều, lãi giảm dần)
          // Bao gồm: Vay trả góp, Trả góp qua thẻ...
          else if (type === 'Trả góp gốc - lãi hàng tháng' || 
                   type === 'Vay trả góp' || 
                   (typeof type === 'string' && type.includes('Trả góp qua thẻ'))) {
             
             const monthlyPrincipal = initialPrincipal / term;
             let simulatedRemaining = initialPrincipal; // Bắt đầu tính từ đầu để khớp lịch
             
             for (let i = 1; i <= term; i++) {
                let payDate = new Date(startDate);
                payDate.setMonth(payDate.getMonth() + i);
                
                // Tính lãi cho kỳ này dựa trên dư nợ đầu kỳ
                let prevDate = new Date(startDate);
                prevDate.setMonth(prevDate.getMonth() + i - 1);
                const daysInMonth = getDaysDiff(prevDate, payDate);
                
                let monthlyInterest = daysInMonth * (rate * simulatedRemaining) / 365;
                
                // Nếu ngày trả >= hôm nay thì hiển thị
                if (payDate >= today) {
                   targetList.push({
                      date: payDate,
                      action: isDebt ? 'Phải trả' : 'Phải thu',
                      name: isDebt ? `${name} (Kỳ ${i}/${term})` : `${name} (Kỳ ${i}/${term})`,
                      remaining: simulatedRemaining, // Dư nợ đầu kỳ
                      principalPayment: monthlyPrincipal,
                      interestPayment: monthlyInterest
                   });
                   
                   break; // Chỉ lấy 1 kỳ tiếp theo
                }
                
                simulatedRemaining -= monthlyPrincipal;
                if (simulatedRemaining < 0) simulatedRemaining = 0;
             }
          }
          
          // Default: Fallback to old logic (Bullet) if type is unknown
          else {
             if (maturityDate && maturityDate >= today) {
               const days = getDaysDiff(startDate, maturityDate);
               let currentInterest = days * (rate * remaining) / 365;
               
               targetList.push({
                  date: maturityDate,
                  action: isDebt ? 'Phải trả' : 'Phải thu',
                  name: isDebt ? `${name} (Tất toán)` : `${name} (Đáo hạn)`,
                  remaining: remaining,
                  principalPayment: remaining,
                  interestPayment: currentInterest
               });
            }
          }
        }
      });
    };
    
    // 1. Debt Payments
    processInstallments(APP_CONFIG.SHEETS.DEBT_MANAGEMENT, true);
    
    // 2. Lending Collections
    processInstallments(APP_CONFIG.SHEETS.LENDING, false);
    
    // Sort and Limit
    payables.sort((a, b) => a.date - b.date);
    receivables.sort((a, b) => a.date - b.date);
    
    return {
      payables: payables.slice(0, 10),
      receivables: receivables.slice(0, 10)
    };
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
    sheet.getRange(headerRow, 1, 14, 10).setBorder(true, true, true, true, true, true, '#B0B0B0', SpreadsheetApp.BorderStyle.SOLID);
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
    
    // Thu nhập (Total Income) / 1000
    sheet.getRange('AA2').setFormula('=' + this._createDynamicSumFormula('THU', 'C') + '/1000');
    
    // Chi phí (Total Expense) / 1000
    // Chi phí = Chi tiêu (CHI) + Trả nợ (TRẢ NỢ)
    const chiFormula = this._createDynamicSumFormula('CHI', 'C');
    const traNoGoc = this._createDynamicSumFormula('TRẢ NỢ', 'D');
    const traNoLai = this._createDynamicSumFormula('TRẢ NỢ', 'E');
    sheet.getRange('AA3').setFormula('=(' + `${chiFormula} + ${traNoGoc} + ${traNoLai}` + ')/1000');
    
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
    
    // Tổng hợp lại cho Chart / 1000
    sheet.getRange('AA4').setFormula('=' + `IFERROR(${netCash} + ${currentValCK} + ${currentValGold} + ${currentValCrypto} + ${currentValOther}, 0)/1000`);
    
    // Nợ (Total Liabilities) / 1000
    // Nợ cũng là Stock (Tích lũy). Lọc theo thời gian cho Nợ nghĩa là gì? "Nợ phát sinh trong kỳ"?
    // Bảng Liabilities đang dùng: SUMIFS('QUẢN LÝ NỢ'!J:J, 'QUẢN LÝ NỢ'!B:B, name) -> Cột J là "Còn nợ".
    // Nó KHÔNG lọc theo ngày tháng trong Bảng Liabilities (xem dòng 164: formula không có date criteria).
    // Vậy Chart cũng không nên lọc theo ngày tháng cho Nợ.
    sheet.getRange('AA5').setFormula(`=SUM('QUẢN LÝ NỢ'!J:J)/1000`);
    
    // Create Chart
    const chart = sheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(chartDataRange)
      .setPosition(39, 12, 0, 0) // Row 39, Col L (Index 12)
      .setOption('title', 'Tổng quan Tài chính (Đơn vị: Nghìn VND)')
      .setOption('width', 600)
      .setOption('height', 400)
      .setOption('legend', { position: 'none' })
      .setOption('useFirstColumnAsDomain', true)
      .setOption('colors', ['#4CAF50', '#F44336', '#2196F3', '#D32F2F']) // Income(Green), Expense(Red), Assets(Blue), Debt(Dark Red)
      .setOption('vAxis.format', '#,##0')
      .setNumHeaders(1)
      .build();
      
    sheet.insertChart(chart);
    sheet.hideColumns(26, 2); // Hide Z, AA
  },
  
  _formatSheet(sheet) {
    // 1. Set columns A-L (1-12) to 120
    sheet.setColumnWidths(1, 12, 120);
    
    // 2. Set column M (13) to 200
    sheet.setColumnWidth(13, 200);
    
    // 3. Set columns N, O, P (14-16) to 120
    sheet.setColumnWidths(14, 3, 120);
    
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
