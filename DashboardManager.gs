/**
 * ===============================================
 * DASHBOARDMANAGER.GS v3.25
 * ===============================================
 * 
 * 
 * A1:C16 - Bảng thống kê (Danh mục, Số tiền, Tỷ lệ)
 * E1:J16 - Biểu đồ cột
 * A19:J34 - Bảng thống kê theo tháng
 */

const DashboardManager = {
  
  CONFIG: {
    TABLE1: {
      START_ROW: 1,
      TITLE_ROW: 1,
      DROPDOWN_ROW: 2,
      HEADER_ROW: 6,
      DATA_START: 7,
      DATA_ROWS: 10,
      CHART_POSITION: {
        ROW: 1,     // ⭐ Row 1 (E1)
        COL: 5,     // ⭐ Column E (cột 5)
        OFFSET_X: 0,
        OFFSET_Y: 0
      },
      CHART_SIZE: {
        WIDTH: 600,   // ⭐ Width để fit từ E đến J
        HEIGHT: 400   // ⭐ Height để fit từ row 1-16
      },
      CHART_DATA: {
        START: 'A7',  // ⭐ Bắt đầu từ A7 (không lấy header)
        END: 'B15'    // ⭐ Kết thúc tại B15
      }
    },
    TABLE2: {
      START_ROW: 19,
      TITLE_ROW: 19,
      HEADER_ROW: 20,
      DATA_START: 21,
      DATA_ROWS: 12,
      SUMMARY_ROW: 33,
      AVG_ROW: 34
    },
    CHART_COLORS: {
      CASH_FLOW: '#64B5F6',
      INCOME: '#81C784',
      EXPENSE: '#4FC3F7',
      DEBT_PRINCIPAL: '#EF5350',
      DEBT_INTEREST: '#E57373',
      STOCK: '#FF9800',
      GOLD: '#FFB74D',
      CRYPTO: '#FFA726',
      OTHER: '#FFCC80'
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
        
        // Xóa tất cả các chart cũ
        const charts = sheet.getCharts();
        charts.forEach(chart => sheet.removeChart(chart));
      }
      
      this._setupTable1(sheet);
      this._setupTable2(sheet);
      this._formatSheet(sheet);
      
      // ⭐⭐ CRITICAL: FLUSH() - Buộc tất cả công thức tính toán xong
      SpreadsheetApp.flush();
      
      // ⭐ DELAY để đảm bảo công thức INDIRECT() phức tạp tính toán xong
      Utilities.sleep(500);
      
      // ⭐ DEBUG: Log dữ liệu trước khi vẽ chart
      const chartData = sheet.getRange('A7:B15').getDisplayValues();
      Logger.log('📊 Chart Data Preview (A7:B15):');
      chartData.forEach((row, idx) => {
        Logger.log(`  Row ${idx + 7}: ${row[0]} = ${row[1]}`);
      });
      
      // ⭐⭐ TẠO CHART - sau khi tất cả đã setup xong
      this._createChart(sheet);
      this._setupTriggers();
      
      SpreadsheetApp.getUi().alert(
        'Khởi tạo thành công!',
        '✅ Dashboard đã được tạo kèm biểu đồ.\n\n' +
        '📊 Biểu đồ: "Báo cáo dòng tiền Gia đình"\n' +
        '📍 Vị trí: E1:J16\n' +
        '📈 Dữ liệu: A7:B15\n\n' +
        '📝 Giá trị mặc định:\n' +
        '- Năm: 2025\n' +
        '- Quý: Tất cả\n' +
        '- Tháng: Tất cả\n\n' +
        '💡 Dashboard sẽ tự động cập nhật khi thay đổi dropdown!',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      
      return sheet;
      
    } catch (error) {
      Logger.log('❌ Lỗi setupDashboard: ' + error.message);
      Logger.log('Stack: ' + error.stack);
      throw new Error('Không thể khởi tạo Dashboard: ' + error.message);
    }
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
    
    Logger.log('✅ Trigger đã được cài đặt');
  },
  
  _setupTable1(sheet) {
    const cfg = this.CONFIG.TABLE1;
    
    // Tiêu đề chính
    sheet.getRange('A1:C1').merge();
    sheet.getRange('A1')
      .setValue('📊 THỐNG KÊ TÀI CHÍNH GIA ĐÌNH')
      .setFontSize(14)
      .setFontWeight('bold')
      .setHorizontalAlignment('left')
      .setVerticalAlignment('middle')
      .setBackground(APP_CONFIG.COLORS.PRIMARY)
      .setFontColor('#FFFFFF');
    sheet.setRowHeight(1, 40);
    
    // Dropdown labels
    sheet.getRange('A2').setValue('Năm').setFontWeight('bold');
    sheet.getRange('A3').setValue('Quý').setFontWeight('bold');
    sheet.getRange('A4').setValue('Tháng').setFontWeight('bold');
    
    const currentYear = new Date().getFullYear();
    
    // 1. Set number format trước
    sheet.getRange('B2').setNumberFormat('@');
    sheet.getRange('B3').setNumberFormat('@');
    sheet.getRange('B4').setNumberFormat('@');
    
    // 2. Tạo data validation
    const yearList = ['Tất cả'];
    for (let y = currentYear - 5; y <= currentYear + 2; y++) {
      yearList.push(y.toString());
    }
    const yearRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(yearList)
      .setAllowInvalid(false)
      .build();
    
    const quarterRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Tất cả', 'Quý 1', 'Quý 2', 'Quý 3', 'Quý 4'])
      .setAllowInvalid(false)
      .build();
    
    const monthList = ['Tất cả'];
    for (let m = 1; m <= 12; m++) {
      monthList.push(`Tháng ${m}`);
    }
    const monthRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(monthList)
      .setAllowInvalid(false)
      .build();
    
    // 3. Set validation
    sheet.getRange('B2').setDataValidation(yearRule);
    sheet.getRange('B3').setDataValidation(quarterRule);
    sheet.getRange('B4').setDataValidation(monthRule);
    
    // 4. ⭐ Set giá trị mặc định SAU KHI đã set validation
    sheet.getRange('B2').setValue('2025');
    sheet.getRange('B3').setValue('Tất cả');
    sheet.getRange('B4').setValue('Tất cả');
    
    // Header row
    const headers1 = ['Danh mục', 'Số tiền', 'Tỷ lệ'];
    sheet.getRange(cfg.HEADER_ROW, 1, 1, 3).setValues([headers1])
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBackground(APP_CONFIG.COLORS.HEADER_BG)
      .setFontColor(APP_CONFIG.COLORS.HEADER_TEXT);
    
    // ⭐ Danh mục categories - 9 dòng (A7:A15)
    const categories = [
      'Tiền mặt', 'Thu', 'Chi', 'Trả nợ', 'Lãi đã trả',
      'Chứng khoán', 'Vàng', 'Crypto', 'Đầu tư khác'
    ];
    
    for (let i = 0; i < categories.length; i++) {
      const row = cfg.DATA_START + i;
      sheet.getRange(row, 1).setValue(categories[i]).setFontWeight('bold');
    }
    
    // ⭐ Dòng TỔNG riêng biệt (A16)
    sheet.getRange(16, 1).setValue('Tổng').setFontWeight('bold');
    
    // Tạo formulas
    this._createTable1Formulas(sheet);
    
    // Format numbers
    sheet.getRange(7, 2, 10, 1).setNumberFormat('#,##0" VNĐ"');
    sheet.getRange(7, 3, 10, 1).setNumberFormat('0.00%');
    
    // Conditional formatting
    const rule1 = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0)
      .setBackground('#D4EDDA')
      .setRanges([sheet.getRange(7, 3, 10, 1)])
      .build();
    
    const rule2 = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0)
      .setBackground('#F8D7DA')
      .setRanges([sheet.getRange(7, 3, 10, 1)])
      .build();
    
    sheet.setConditionalFormatRules([rule1, rule2]);
    
    // Borders
    sheet.getRange(cfg.HEADER_ROW, 1, cfg.DATA_ROWS + 1, 3)
      .setBorder(true, true, true, true, true, true);
  },
  
  _createTable1Formulas(sheet) {
    // B7: Tiền mặt = Thu - Chi - Trả nợ - Lãi - CK - Vàng - Crypto - ĐT khác
    sheet.getRange(7, 2).setFormula(`=IFERROR(B8-B9-B10-B11-B12-B13-B14-B15, 0)`);
    
    // B8: Thu
    sheet.getRange(8, 2).setFormula(this._createDynamicSumFormula('THU', 'C'));
    
    // B9: Chi
    sheet.getRange(9, 2).setFormula(this._createDynamicSumFormula('CHI', 'C'));
    
    // B10: Trả nợ (gốc)
    sheet.getRange(10, 2).setFormula(this._createDynamicSumFormula('TRẢ NỢ', 'D'));
    
    // B11: Lãi đã trả
    sheet.getRange(11, 2).setFormula(this._createDynamicSumFormula('TRẢ NỢ', 'E'));
    
    // B12: Chứng khoán
    sheet.getRange(12, 2).setFormula(this._createDynamicInvestmentFormula('CHỨNG KHOÁN', 'H'));
    
    // B13: Vàng
    sheet.getRange(13, 2).setFormula(this._createDynamicInvestmentFormula('VÀNG', 'H'));
    
    // B14: Crypto
    sheet.getRange(14, 2).setFormula(this._createDynamicInvestmentFormula('CRYPTO', 'I'));
    
    // B15: Đầu tư khác
    sheet.getRange(15, 2).setFormula(this._createDynamicSumFormula('ĐẦU TƯ KHÁC', 'D'));
    
    // B16: Tổng = Thu
    sheet.getRange(16, 2).setFormula(`=IFERROR(B8, 0)`);
    
    // C7:C16: Tỷ lệ % (so với tổng thu)
    for (let row = 7; row <= 16; row++) {
      sheet.getRange(row, 3).setFormula(`=IFERROR(IF($B16<>0, B${row}/$B16, 0), 0)`);
    }
  },
  
  _createDynamicSumFormula(sheetName, col) {
    return `=IFERROR(
      IF(OR($B$2="", $B$3="", $B$4=""), 0,
        IF($B$2="Tất cả",
          SUM(INDIRECT("'"&"${sheetName}"&"'!${col}:${col}")),
          IF($B$3="Tất cả",
            SUMIFS(INDIRECT("'"&"${sheetName}"&"'!${col}:${col}"), 
                   INDIRECT("'"&"${sheetName}"&"'!B:B"), ">="&DATE(VALUE($B$2),1,1), 
                   INDIRECT("'"&"${sheetName}"&"'!B:B"), "<"&DATE(VALUE($B$2)+1,1,1)),
            IF($B$4="Tất cả",
              SUMIFS(INDIRECT("'"&"${sheetName}"&"'!${col}:${col}"), 
                     INDIRECT("'"&"${sheetName}"&"'!B:B"), ">="&DATE(VALUE($B$2),(VALUE(RIGHT($B$3,1))-1)*3+1,1), 
                     INDIRECT("'"&"${sheetName}"&"'!B:B"), "<"&DATE(VALUE($B$2),(VALUE(RIGHT($B$3,1))-1)*3+4,1)),
              SUMIFS(INDIRECT("'"&"${sheetName}"&"'!${col}:${col}"), 
                     INDIRECT("'"&"${sheetName}"&"'!B:B"), ">="&DATE(VALUE($B$2),VALUE(RIGHT($B$4,LEN($B$4)-6)),1), 
                     INDIRECT("'"&"${sheetName}"&"'!B:B"), "<"&DATE(VALUE($B$2),VALUE(RIGHT($B$4,LEN($B$4)-6))+1,1))
            )
          )
        )
      ), 0)`;
  },
  
  _createDynamicInvestmentFormula(sheetName, col) {
    return `=IFERROR(
      IF(OR($B$2="", $B$3="", $B$4=""), 0,
        IF($B$2="Tất cả",
          SUMIF(INDIRECT("'"&"${sheetName}"&"'!C:C"), "Mua", INDIRECT("'"&"${sheetName}"&"'!${col}:${col}")),
          IF($B$3="Tất cả",
            SUMIFS(INDIRECT("'"&"${sheetName}"&"'!${col}:${col}"), 
                   INDIRECT("'"&"${sheetName}"&"'!C:C"), "Mua",
                   INDIRECT("'"&"${sheetName}"&"'!B:B"), ">="&DATE(VALUE($B$2),1,1), 
                   INDIRECT("'"&"${sheetName}"&"'!B:B"), "<"&DATE(VALUE($B$2)+1,1,1)),
            IF($B$4="Tất cả",
              SUMIFS(INDIRECT("'"&"${sheetName}"&"'!${col}:${col}"), 
                     INDIRECT("'"&"${sheetName}"&"'!C:C"), "Mua",
                     INDIRECT("'"&"${sheetName}"&"'!B:B"), ">="&DATE(VALUE($B$2),(VALUE(RIGHT($B$3,1))-1)*3+1,1), 
                     INDIRECT("'"&"${sheetName}"&"'!B:B"), "<"&DATE(VALUE($B$2),(VALUE(RIGHT($B$3,1))-1)*3+4,1)),
              SUMIFS(INDIRECT("'"&"${sheetName}"&"'!${col}:${col}"), 
                     INDIRECT("'"&"${sheetName}"&"'!C:C"), "Mua",
                     INDIRECT("'"&"${sheetName}"&"'!B:B"), ">="&DATE(VALUE($B$2),VALUE(RIGHT($B$4,LEN($B$4)-6)),1), 
                     INDIRECT("'"&"${sheetName}"&"'!B:B"), "<"&DATE(VALUE($B$2),VALUE(RIGHT($B$4,LEN($B$4)-6))+1,1))
            )
          )
        )
      ), 0)`;
  },
  
  _setupTable2(sheet) {
    const cfg = this.CONFIG.TABLE2;
    const currentYear = new Date().getFullYear();
    
    // Title
    sheet.getRange(`A${cfg.TITLE_ROW}:J${cfg.TITLE_ROW}`).merge();
    sheet.getRange(cfg.TITLE_ROW, 1)
      .setValue(`📈 THỐNG KÊ TÀI CHÍNH GIA ĐÌNH NĂM ${currentYear}`)
      .setFontSize(12)
      .setFontWeight('bold')
      .setHorizontalAlignment('left')
      .setVerticalAlignment('middle')
      .setBackground(APP_CONFIG.COLORS.PRIMARY)
      .setFontColor('#FFFFFF');
    sheet.setRowHeight(cfg.TITLE_ROW, 35);
    
    // Headers
    const headers2 = [
      'Kỳ', 'Thu', 'Chi', 'Nợ (Gốc)', 'Lãi', 
      'CK', 'Vàng', 'Crypto', 'ĐT khác', 'Dòng tiền'
    ];
    
    sheet.getRange(cfg.HEADER_ROW, 1, 1, 10).setValues([headers2])
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBackground(APP_CONFIG.COLORS.HEADER_BG)
      .setFontColor(APP_CONFIG.COLORS.HEADER_TEXT);
    
    // Monthly data
    for (let month = 1; month <= 12; month++) {
      const row = cfg.DATA_START + month - 1;
      
      sheet.getRange(row, 1).setValue(`Tháng ${month}`);
      
      // Thu
      sheet.getRange(row, 2).setFormula(
        `=SUMIFS(THU!C:C, THU!B:B, ">="&DATE(${currentYear},${month},1), THU!B:B, "<"&DATE(${currentYear},${month}+1,1))`
      );
      
      // Chi
      sheet.getRange(row, 3).setFormula(
        `=SUMIFS(CHI!C:C, CHI!B:B, ">="&DATE(${currentYear},${month},1), CHI!B:B, "<"&DATE(${currentYear},${month}+1,1))`
      );
      
      // Nợ gốc
      sheet.getRange(row, 4).setFormula(
        `=SUMIFS('TRẢ NỢ'!D:D, 'TRẢ NỢ'!B:B, ">="&DATE(${currentYear},${month},1), 'TRẢ NỢ'!B:B, "<"&DATE(${currentYear},${month}+1,1))`
      );
      
      // Lãi
      sheet.getRange(row, 5).setFormula(
        `=SUMIFS('TRẢ NỢ'!E:E, 'TRẢ NỢ'!B:B, ">="&DATE(${currentYear},${month},1), 'TRẢ NỢ'!B:B, "<"&DATE(${currentYear},${month}+1,1))`
      );
      
      // Chứng khoán
      sheet.getRange(row, 6).setFormula(
        `=SUMIFS('CHỨNG KHOÁN'!H:H, 'CHỨNG KHOÁN'!C:C, "Mua", 'CHỨNG KHOÁN'!B:B, ">="&DATE(${currentYear},${month},1), 'CHỨNG KHOÁN'!B:B, "<"&DATE(${currentYear},${month}+1,1))`
      );
      
      // Vàng
      sheet.getRange(row, 7).setFormula(
        `=SUMIFS(VÀNG!H:H, VÀNG!C:C, "Mua", VÀNG!B:B, ">="&DATE(${currentYear},${month},1), VÀNG!B:B, "<"&DATE(${currentYear},${month}+1,1))`
      );
      
      // Crypto
      sheet.getRange(row, 8).setFormula(
        `=SUMIFS(CRYPTO!I:I, CRYPTO!C:C, "Mua", CRYPTO!B:B, ">="&DATE(${currentYear},${month},1), CRYPTO!B:B, "<"&DATE(${currentYear},${month}+1,1))`
      );
      
      // Đầu tư khác
      sheet.getRange(row, 9).setFormula(
        `=SUMIFS('ĐẦU TƯ KHÁC'!D:D, 'ĐẦU TƯ KHÁC'!B:B, ">="&DATE(${currentYear},${month},1), 'ĐẦU TƯ KHÁC'!B:B, "<"&DATE(${currentYear},${month}+1,1))`
      );
      
      // Dòng tiền
      const b = `B${row}`, c = `C${row}`, d = `D${row}`, e = `E${row}`;
      const f = `F${row}`, g = `G${row}`, h = `H${row}`, i = `I${row}`;
      sheet.getRange(row, 10).setFormula(
        `=IFERROR(${b}-(${c}+${d}+${e}+${f}+${g}+${h}+${i}), 0)`
      );
    }
    
    // Tổng
    sheet.getRange(cfg.SUMMARY_ROW, 1).setValue('Tổng').setFontWeight('bold');
    for (let col = 2; col <= 10; col++) {
      const startCell = this._getColumnLetter(col) + cfg.DATA_START;
      const endCell = this._getColumnLetter(col) + (cfg.DATA_START + 11);
      sheet.getRange(cfg.SUMMARY_ROW, col)
        .setFormula(`=SUM(${startCell}:${endCell})`)
        .setFontWeight('bold')
        .setBackground('#FFE0B2');
    }
    
    // Trung bình
    sheet.getRange(cfg.AVG_ROW, 1).setValue('Trung bình').setFontWeight('bold');
    for (let col = 2; col <= 10; col++) {
      const startCell = this._getColumnLetter(col) + cfg.DATA_START;
      const endCell = this._getColumnLetter(col) + (cfg.DATA_START + 11);
      sheet.getRange(cfg.AVG_ROW, col)
        .setFormula(`=AVERAGE(${startCell}:${endCell})`)
        .setFontWeight('bold')
        .setBackground('#FFF3E0');
    }
    
    // Format
    sheet.getRange(cfg.DATA_START, 2, 14, 9).setNumberFormat('#,##0" VNĐ"');
    sheet.getRange(cfg.HEADER_ROW, 1, 15, 10).setBorder(true, true, true, true, true, true);
  },
  
  /**
   * ⭐⭐ TẠO CHART TẠI VỊ TRÍ E1:J16
   * 
   * THAY ĐỔI:
   * 1. Dữ liệu: A7:B15 (9 dòng, không có header, không có Tổng)
   * 2. Vị trí: E1 (row 1, col 5)
   * 3. Tiêu đề: "Báo cáo dòng tiền Gia đình"
   * 4. Size: 600x400px để fit E1:J16
   */
  _createChart(sheet) {
    const cfg = this.CONFIG.TABLE1;
    
    Logger.log('🔧 Creating chart at E1:J16...');
    Logger.log(`📊 Data range: ${cfg.CHART_DATA.START}:${cfg.CHART_DATA.END}`);
    Logger.log(`📍 Position: Row ${cfg.CHART_POSITION.ROW}, Col ${cfg.CHART_POSITION.COL}`);
    Logger.log(`📐 Size: ${cfg.CHART_SIZE.WIDTH}x${cfg.CHART_SIZE.HEIGHT}`);
    
    // ⭐ Tạo biểu đồ cột với dữ liệu từ A7:B15
    const chart = sheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(sheet.getRange(cfg.CHART_DATA.START + ':' + cfg.CHART_DATA.END))
      .setPosition(
        cfg.CHART_POSITION.ROW, 
        cfg.CHART_POSITION.COL, 
        cfg.CHART_POSITION.OFFSET_X, 
        cfg.CHART_POSITION.OFFSET_Y
      )
      .setOption('title', 'Báo cáo dòng tiền Gia đình')
      .setOption('width', cfg.CHART_SIZE.WIDTH)
      .setOption('height', cfg.CHART_SIZE.HEIGHT)
      .setOption('legend', { position: 'none' })
      .setOption('hAxis', { 
        title: 'Danh mục',
        slantedText: true,
        slantedTextAngle: 45
      })
      .setOption('vAxis', { 
        title: 'Số tiền (VNĐ)',
        format: '#,###'
      })
      .build();
    
    sheet.insertChart(chart);
    
    Logger.log('✅ Chart created successfully at E1:J16');
    Logger.log(`   - Title: "Báo cáo dòng tiền Gia đình"`);
    Logger.log(`   - Data: A7:B15 (9 categories)`);
    Logger.log(`   - Position: E1`);
  },
  
  _formatSheet(sheet) {
    // Set column widths
    sheet.setColumnWidth(1, 120);  // A: Danh mục
    sheet.setColumnWidth(2, 120);  // B: Số tiền
    sheet.setColumnWidth(3, 100);  // C: Tỷ lệ
    for (let col = 4; col <= 10; col++) {
      sheet.setColumnWidth(col, 100);
    }
    sheet.setFrozenRows(1);
  },
  
  refreshDashboard() {
    try {
      SpreadsheetApp.flush();
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(APP_CONFIG.SHEETS.DASHBOARD);
      
      if (sheet) {
        const b2 = sheet.getRange('B2').getValue();
        const b3 = sheet.getRange('B3').getValue();
        const b4 = sheet.getRange('B4').getValue();
        
        // Force recalculation
        sheet.getRange('B2').setValue(b2 + ' ');
        SpreadsheetApp.flush();
        sheet.getRange('B2').setValue(b2);
        
        sheet.getRange('B3').setValue(b3 + ' ');
        SpreadsheetApp.flush();
        sheet.getRange('B3').setValue(b3);
        
        sheet.getRange('B4').setValue(b4 + ' ');
        SpreadsheetApp.flush();
        sheet.getRange('B4').setValue(b4);
        
        SpreadsheetApp.flush();
        
        Logger.log('✅ Dashboard refreshed');
      }
    } catch (error) {
      Logger.log('❌ Lỗi refreshDashboard: ' + error.message);
    }
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

/**
 * ⭐ TRIGGER: Auto-update dashboard khi thay đổi dropdown
 */
function onDashboardEdit(e) {
  try {
    if (!e) return;
    
    const range = e.range;
    const sheet = range.getSheet();
    const sheetName = sheet.getName();
    
    if (sheetName === APP_CONFIG.SHEETS.DASHBOARD) {
      const row = range.getRow();
      const col = range.getColumn();
      
      // Chỉ trigger khi thay đổi dropdown B2, B3, B4
      if (col === 2 && (row === 2 || row === 3 || row === 4)) {
        Logger.log('⭐ Dropdown changed, force recalculate...');
        Utilities.sleep(100);
        SpreadsheetApp.flush();
        sheet.getDataRange().getValues();
        Logger.log('✅ Dashboard auto-updated');
      }
    }
  } catch (error) {
    Logger.log('❌ Lỗi onDashboardEdit: ' + error.message);
  }
}

/**
 * ⭐ MENU FUNCTION: Manual refresh
 */
function refreshDashboard() {
  DashboardManager.refreshDashboard();
  SpreadsheetApp.getUi().alert(
    'Thành công', 
    '✅ Dashboard đã được cập nhật!', 
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}