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
      
      // 3. Setup Monthly Table (Table 2) - Fixed at row 36
      this._setupTable2(sheet, 36);
      
      // 4. Setup Chart - Fixed at K36
      this._createChart(sheet);
      
      // 5. Setup Action Buttons (Checkboxes) - DISABLED
      // this._setupActionButtons(sheet);
      
      // 6. Format & Finalize
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
    sheet.getRange('A1:I1').merge()
      .setValue('📊 BÁO CÁO TÀI CHÍNH (CASHFLOW)')
      .setFontSize(14)
      .setFontWeight('bold')
      .setVerticalAlignment('middle')
      .setBackground(this.CONFIG.COLORS.HEADER)
      .setFontColor(this.CONFIG.COLORS.TEXT);
    sheet.setRowHeight(1, 40);
    
    // Dropdowns Labels
    sheet.getRange('A2').setValue('Năm').setFontWeight('bold');
    sheet.getRange('A3').setValue('Quý').setFontWeight('bold');
    sheet.getRange('A4').setValue('Tháng').setFontWeight('bold');
    
    const currentYear = new Date().getFullYear();
    
    // FIX 1: Apply Text format ONLY to B3:B4. Leave B2 (Year) as Automatic/Number.
    sheet.getRange('B3:B4').setNumberFormat('@');
    sheet.getRange('B2').setNumberFormat('0'); // Format Year as number (no decimals)

    // Setup Year List (Use Numbers)
    const yearList = ['Tất cả'];
    for (let y = currentYear - 5; y <= currentYear + 2; y++) {
      yearList.push(y); // Push NUMBER, not string
    }

    const monthList = ['Tất cả'];
    for (let m = 1; m <= 12; m++) monthList.push(`Tháng ${m}`);
    
    // Set Validation
    sheet.getRange('B2').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(yearList).build());
    sheet.getRange('B3').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Tất cả', 'Quý 1', 'Quý 2', 'Quý 3', 'Quý 4']).build());
    sheet.getRange('B4').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(monthList).build());
    
    // Set Values
    sheet.getRange('B2').setValue(currentYear); // Set NUMBER
    sheet.getRange('B3').setValue('Tất cả');
    sheet.getRange('B4').setValue('Tất cả');
    
    // Chart Data Headers (E2:H2)
    sheet.getRange('F2').setValue('Thu nhập').setFontWeight('bold').setHorizontalAlignment('center');
    sheet.getRange('G2').setValue('Chi phí').setFontWeight('bold').setHorizontalAlignment('center');
    sheet.getRange('H2').setValue('Nợ').setFontWeight('bold').setHorizontalAlignment('center');
    sheet.getRange('I2').setValue('Tài sản').setFontWeight('bold').setHorizontalAlignment('center');
  },

  _setupActionButtons(sheet) {
    // Define buttons configuration
    const buttons = [
      { cell: 'D2', label: '+ NHẬP THU', color: '#4CAF50' },
      { cell: 'D4', label: '+ NHẬP CHI', color: '#F44336' },
      { cell: 'F2', label: 'QUẢN LÝ NỢ', color: '#FF9800' },
      { cell: 'F4', label: 'QUẢN LÝ CHO VAY', color: '#2196F3' },
      { cell: 'H2', label: 'GD VÀNG', color: '#FFC107' },
      { cell: 'H4', label: 'GD CHỨNG KHOÁN', color: '#9C27B0' },
      { cell: 'J2', label: 'GD CRYPTO', color: '#607D8B' },
      { cell: 'J4', label: 'ĐẦU TƯ KHÁC', color: '#795548' }
    ];
    
    buttons.forEach(btn => {
      // Set Checkbox
      const range = sheet.getRange(btn.cell);
      range.insertCheckboxes(); 
      // FIX 2: Removed range.setValue(false); to avoid validation conflict.
      // Checkboxes are FALSE (unchecked) by default anyway.
      
      range.setFontColor(btn.color);
      range.setFontWeight('bold');
      
      // Set Label in the NEXT cell
      const labelRange = range.offset(0, 1);
      labelRange.setValue(btn.label);
      labelRange.setFontColor(btn.color);
      labelRange.setFontWeight('bold');
      labelRange.setHorizontalAlignment('left');
    });
  },
  

  _setupGridTables(sheet, startRow) {
    const cfg = this.CONFIG.LAYOUT;
    let currentRow = startRow;
    
    // Fetch events once - REMOVED (Using Custom Functions)
    // const events = this._getCalendarEvents();
    
    // === ROW 1: INCOME (Left) & EXPENSE (Middle) & PAYABLES (Right) ===
    
    // Calculate expected row counts from config
    const incomeRowCount = APP_CONFIG.CATEGORIES.INCOME.length + 1; // +1 for total
    const expenseRowCount = APP_CONFIG.CATEGORIES.EXPENSE.filter(c => c !== 'Trả nợ' && c !== 'Cho vay').length + 2; // +1 debt payment, +1 total
    
    // 1. Income Table - Using custom function
    const incomeHeight = this._renderCustomFunctionTable(
      sheet, currentRow, cfg.LEFT_COL, 
      '1. Báo cáo Thu nhập', 
      this.CONFIG.COLORS.INCOME, 
      '=hffsIncome()', 
      3, // 3 columns
      true, // has percentage
      incomeRowCount, // exact row count
      ['Danh mục', 'Giá trị', 'Tỷ lệ'] // headers
    );
    
    // 2. Expense Table - Using custom function
    const expenseHeight = this._renderCustomFunctionTable(
      sheet, currentRow, cfg.RIGHT_COL,
      '2. Báo cáo Chi phí',
      this.CONFIG.COLORS.EXPENSE,
      '=hffsExpense()',
      5, // 5 columns
      true, // has percentage
      expenseRowCount, // exact row count
      ['Danh mục', 'Đã chi', 'Ngân sách', 'Còn lại', 'Trạng thái'] // headers
    );
    
    // 3. Payables Table (Right - Col K)
    const payablesHeight = this._renderPayables(sheet, currentRow, cfg.CALENDAR_COL);

    // Calculate max height for Row 1
    const row1Height = Math.max(incomeHeight, expenseHeight, payablesHeight);
    currentRow += row1Height + 2; // +2 padding
    
    // === ROW 2: LIABILITIES (Left) & ASSETS (Middle) & RECEIVABLES (Right) ===
    
    // Get active debt count
    const debtItems = this._getDebtItems();
    const debtRowCount = debtItems.length + 1; // +1 for total
    const assetRowCount = 7; // 6 asset types + total
    
    // 4. Liabilities Table - Using custom function
    const liabilityHeight = this._renderCustomFunctionTable(
      sheet, currentRow, cfg.LEFT_COL,
      '3. Báo cáo Nợ phải trả',
      this.CONFIG.COLORS.LIABILITIES,
      '=hffsDebt()',
      3, // 3 columns
      true, // has percentage
      debtRowCount, // exact row count
      ['Khoản nợ', 'Còn nợ', 'Tỷ lệ'] // headers
    );
    
    // 5. Assets Table - Using custom function
    const assetHeight = this._renderCustomFunctionTable(
      sheet, currentRow, cfg.RIGHT_COL,
      '4. Báo cáo Tài sản',
      this.CONFIG.COLORS.ASSETS,
      '=hffsAssets()',
      5, // 5 columns
      true, // has percentage
      assetRowCount, // exact row count
      ['Danh mục', 'Tổng vốn', 'Lãi/Lỗ', 'Giá trị HT', 'Tỷ lệ'] // headers
    );
    
    // 6. Receivables Table (Right - Col K, Row 2)
    const receivablesHeight = this._renderReceivables(sheet, currentRow, cfg.CALENDAR_COL);


    // Calculate max height for Row 2
    const row2Height = Math.max(liabilityHeight, assetHeight, receivablesHeight);
    
    // Formatting is handled by _renderCustomFunctionTable

    
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
    // Filter out 'Trả nợ' and 'Cho vay'
    const expenseCategories = APP_CONFIG.CATEGORIES.EXPENSE.filter(cat => cat !== 'Trả nợ' && cat !== 'Cho vay');
    const totalRowIdx = dataStart + rows.length - 1;
    
    rows.forEach((name, idx) => {
      const r = dataStart + idx;
      sheet.getRange(r, startCol).setValue(name);
      
      // 1. Đã chi (Spent)
      if (idx < expenseCategories.length) {
        // Normal categories
        sheet.getRange(r, startCol + 1).setFormula('=' + this._createDynamicSumFormula('CHI', 'C', name, 'D'));
      } else if (name === 'Trả nợ (Gốc + Lãi)') {
        // Debt - Gộp Gốc (col D) + Lãi (col E) từ sheet TRẢ NỢ
        sheet.getRange(r, startCol + 1).setFormula('=' + `${this._createDynamicSumFormula('TRẢ NỢ', 'D')} + ${this._createDynamicSumFormula('TRẢ NỢ', 'E')}`);
      } else {
        // Total
        sheet.getRange(r, startCol + 1).setFormula(`=SUM(R[-${rows.length - 1}]C:R[-1]C)`);
      }
      
      // 2. Ngân sách (Budget)
      if (name === 'TỔNG CHI PHÍ') {
         sheet.getRange(r, startCol + 2).setFormula(`=SUM(R[-${rows.length - 1}]C:R[-1]C)`);
      } else if (name === 'Đầu tư') {
         // Map 'Đầu tư' expense to 'TỔNG ĐT' in Budget
         const budgetBase = `IFERROR(VLOOKUP("TỔNG ĐT", BUDGET!A:C, 3, 0), 0)`;
         const multiplier = `IF($B$4<>"Tất cả", 1, IF($B$3<>"Tất cả", 3, 12))`;
         sheet.getRange(r, startCol + 2).setFormula(`=${budgetBase} * ${multiplier}`);
      } else {
         // VLOOKUP từ BUDGET sheet với multiplier động theo bộ lọc thời gian
         // Nếu chọn Tháng: x1, Quý: x3, Năm: x12
         const budgetBase = `IFERROR(VLOOKUP("${name}", BUDGET!A:C, 3, 0), 0)`;
         const multiplier = `IF($B$4<>"Tất cả", 1, IF($B$3<>"Tất cả", 3, 12))`;
         sheet.getRange(r, startCol + 2).setFormula(`=${budgetBase} * ${multiplier}`);
      }
      
      // 3. Còn lại (Remaining) = Budget - Spent
      sheet.getRange(r, startCol + 3).setFormula(`=R[0]C[-1] - R[0]C[-2]`);
      
      // 4. Trạng thái (Status) - Icon + Percent
      const statusFormula = `=IF(R[0]C[-2]=0, "⚪ 0%", 
        IF(R[0]C[-3] / R[0]C[-2] > 1, "🔴 " & TEXT(R[0]C[-3] / R[0]C[-2], "0.0%"), 
        IF(R[0]C[-3] / R[0]C[-2] >= 0.8, "⚠️ " & TEXT(R[0]C[-3] / R[0]C[-2], "0.0%"), 
        "✅ " & TEXT(R[0]C[-3] / R[0]C[-2], "0.0%"))))`;
      sheet.getRange(r, startCol + 4).setFormula(statusFormula);
      
      // Last row styling
      if (idx === rows.length - 1) {
        sheet.getRange(r, startCol).setFontWeight('bold');
        sheet.getRange(r, startCol, 1, numCols).setBackground('#EEEEEE');
      }
    });  // <-- ĐÃ THÊM DẤU ĐÓNG forEach
    
    // Conditional Formatting for Status Column
    const statusRange = sheet.getRange(dataStart, startCol + 4, rows.length, 1);

    // Red - Vượt ngân sách (🔴)
    const ruleRed = SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('🔴')
      .setBackground('#FFEBEE')
      .setFontColor('#C62828')
      .setRanges([statusRange])
      .build();
      
    // Yellow - Sắp hết (⚠️)
    const ruleYellow = SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('⚠️')
      .setBackground('#FFF3E0')
      .setFontColor('#EF6C00')
      .setRanges([statusRange])
      .build();
      
    // Green - Trong hạn mức (✅)
    const ruleGreen = SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('✅')
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
  },
  
  /**
   * Render table using custom function that returns array
   * @param {Sheet} sheet
   * @param {number} startRow
   * @param {number} startCol
   * @param {string} title
   * @param {string} color
   * @param {string} formula - Custom function formula (e.g., '=hffsIncome()')
   * @param {number} numCols - Number of columns
   * @param {boolean} hasPercentage - Whether last column is percentage
   * @param {number} expectedRows - Expected number of data rows (excluding header)
   * @param {Array} subHeaders - Optional column headers
   * @return {number} Height of table
   */
  _renderCustomFunctionTable(sheet, startRow, startCol, title, color, formula, numCols, hasPercentage, expectedRows, subHeaders) {
    // Header
    sheet.getRange(startRow, startCol, 1, numCols).merge()
      .setValue(title)
      .setFontWeight('bold')
      .setBackground(color)
      .setFontColor('#FFFFFF')
      .setHorizontalAlignment('left');
    
    let dataStartRow = startRow + 1;
    
    // Optional sub-headers
    if (subHeaders && subHeaders.length === numCols) {
      sheet.getRange(dataStartRow, startCol, 1, numCols)
        .setValues([subHeaders])
        .setFontWeight('bold')
        .setBackground('#EEEEEE')
        .setHorizontalAlignment('center');
      dataStartRow++;
    }
    
    // Data range - place formula in first cell
    // The formula will spill to adjacent cells automatically
    sheet.getRange(dataStartRow, startCol).setFormula(formula);
    
    // Format numbers based on expected rows
    // Format all columns with numbers except first (category name)
    if (numCols > 1) {
      sheet.getRange(dataStartRow, startCol + 1, expectedRows, numCols - 1).setNumberFormat('#,##0');
    }
    
    // Format percentage column if exists
    if (hasPercentage && numCols > 1) {
      sheet.getRange(dataStartRow, startCol + numCols - 1, expectedRows, 1).setNumberFormat('0.0%');
    }
    
    // Border around actual table area (header + sub-headers + expected rows)
    const totalHeight = (subHeaders ? 2 : 1) + expectedRows;
    sheet.getRange(startRow, startCol, totalHeight, numCols)
      .setBorder(true, true, true, true, true, true, '#B0B0B0', SpreadsheetApp.BorderStyle.SOLID);
    
    // Return actual height
    return totalHeight;
  },
  
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
    sheet.getRange(startRow, 1, 1, 9).merge()
      .setValue(`📈 THỐNG KÊ TÀI CHÍNH GIA ĐÌNH NĂM ${currentYear}`)
      .setFontSize(12)
      .setFontWeight('bold')
      .setBackground(this.CONFIG.COLORS.HEADER)
      .setFontColor('#FFFFFF');
    sheet.setRowHeight(startRow, 35);
    
    const dataStart = startRow + 1;
    
    // Use custom function to get yearly data (uses current year by default)
    sheet.getRange(dataStart, 1).setFormula('=hffsYearly()');
    
    // Format
    sheet.getRange(dataStart, 2, 12, 8).setNumberFormat('#,##0');
    sheet.getRange(dataStart, 1, 13, 9).setBorder(true, true, true, true, true, true, '#B0B0B0', SpreadsheetApp.BorderStyle.SOLID);
  },
  
  _createChart(sheet) {
    // 1. Cấu hình vị trí Chart
    const chartStartRow = 36;
    const chartStartCol = 11; // Column K

    // 2. Calculate data source positions
    // With custom functions, we use fixed positions for totals
    // Income: Last row of hffsIncome() output (varies by categories)
    // Expense: Last row of hffsExpense() output
    // Debt: Last row of hffsDebt() output  
    // Assets: Last row of hffsAssets() output
    
    // Since custom functions auto-fill, we use INDEX to get last row values
    // Or we can use fixed cells if we know the max rows
    
    const cfg = this.CONFIG.LAYOUT;
    const startRow = cfg.START_ROW;
    
    // For now, use simple SUMs to aggregate the custom function outputs
    // Income Total: Sum column B (amounts) where data exists
    sheet.getRange('F3').setFormula(`=IFERROR(INDEX(A:C, MATCH("TỔNG THU NHẬP", A:A, 0), 2), 0)`);
    
    // Expense Total: Sum column F (amounts) in expense table
    sheet.getRange('G3').setFormula(`=IFERROR(INDEX(E:I, MATCH("TỔNG CHI PHÍ", E:E, 0), 2), 0)`);
    
    // Debt Total: Sum column B in debt table  
    const row2Start = startRow + 23; // Approximate based on row1 height
    sheet.getRange('H3').setFormula(`=IFERROR(INDEX(A:C, MATCH("TỔNG NỢ", A:A, 0), 2), 0)`);
    
    // Assets Total: Sum column H in assets table
    sheet.getRange('I3').setFormula(`=IFERROR(INDEX(E:I, MATCH("TỔNG TÀI SẢN", E:E, 0), 4), 0)`);

    // Format F3:I3
    sheet.getRange('F3:I3').setNumberFormat('#,##0');
    
    // 4. TẠO BIỂU ĐỒ
    const chart = sheet.newChart()
        .setChartType(Charts.ChartType.COLUMN)
        
        // --- FIX: Sử dụng range F2:I3 (Headers + Data) ---
        .addRange(sheet.getRange('F2:I3')) 
        
        .setPosition(chartStartRow, chartStartCol, 0, 0)
        .setOption('title', 'TỔNG QUAN TÀI CHÍNH')
        .setOption('titleTextStyle', { fontSize: 14, bold: true, color: '#333333' })
        .setOption('width', 720)
        .setOption('height', 420)
        
        // --- FIX: Transpose = false (Mỗi cột là 1 series) ---
        // F2:I2 là Header, F3:I3 là Data
        .setTransposeRowsAndColumns(false) 
        
        .setOption('legend', { position: 'bottom', textStyle: { fontSize: 11 } })
        
        // --- FIX: Màu sắc cho từng Series ---
        .setOption('series', {
            0: { color: '#4CAF50' }, // Thu nhập
            1: { color: '#F44336' }, // Chi phí
            2: { color: '#FF9800' }, // Nợ
            3: { color: '#2196F3' }  // Tài sản
        })
        
        .setOption('vAxis', { 
            format: 'short',
            gridlines: { count: 5 }
        })
        // Ẩn trục ngang (vì chỉ có 1 nhóm dữ liệu, không cần nhãn trục X)
        .setOption('hAxis', { textPosition: 'none' }) 
        
        .setOption('chartArea', { width: '85%', height: '70%' })
        .build();
      
    sheet.insertChart(chart);
  },
  
  _formatSheet(sheet) {
    // Hide gridlines
    sheet.setHiddenGridlines(true);
    
    // Set tất cả column widths = 120
    for (let col = 1; col <= 20; col++) {
      sheet.setColumnWidth(col, 120);
    }
  },
  
  _setupTriggers() {
    // Tự động cài đặt Installable Triggers để hỗ trợ Quick Actions
    try {
      createInstallableTriggers(true); // silent = true để không hiển thị alert
      Logger.log('✅ Triggers đã được cài đặt tự động');
    } catch (error) {
      Logger.log('⚠️ Không thể cài đặt triggers tự động: ' + error.message);
      Logger.log('Người dùng có thể chạy thủ công từ menu hoặc hàm createInstallableTriggers()');
    }
  },
  
  /**
   * Wrapper function để refresh Dashboard
   * Gọi setupDashboard để đảm bảo triggers được cài đặt
   */
  refreshDashboard() {
    return this.setupDashboard();
  }
};