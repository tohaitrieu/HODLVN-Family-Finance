/**
 * ===============================================
 * SHEETUTILS.GS - TIỆN ÍCH XỬ LÝ SHEET CONFIG
 * ===============================================
 */

const SheetUtils = {
  
  /**
   * Lấy index cột (1-based) dựa trên key
   * @param {string} sheetKey - Key trong SHEET_CONFIG (VD: 'INCOME')
   * @param {string} colKey - Key của cột (VD: 'amount')
   * @return {number} Index cột (1-based), hoặc -1 nếu không tìm thấy
   */
  getColIndex(sheetKey, colKey) {
    const config = SHEET_CONFIG[sheetKey];
    if (!config) throw new Error(`Sheet config not found: ${sheetKey}`);
    
    const index = config.columns.findIndex(col => col.key === colKey);
    return index >= 0 ? index + 1 : -1;
  },

  /**
   * Map object dữ liệu sang mảng row để ghi vào Sheet
   * @param {string} sheetKey - Key trong SHEET_CONFIG
   * @param {Object} dataObj - Object dữ liệu { key: value }
   * @return {Array} Mảng dữ liệu theo đúng thứ tự cột
   */
  dataToRow(sheetKey, dataObj) {
    const config = SHEET_CONFIG[sheetKey];
    if (!config) throw new Error(`Sheet config not found: ${sheetKey}`);
    
    return config.columns.map(col => {
      // Nếu cột là công thức, trả về null hoặc công thức (nếu cần set)
      // Thường thì công thức được set bởi ArrayFormula hoặc setFormula riêng
      // Ở đây ta trả về value nếu có, hoặc ''
      if (col.type === 'formula') return '';
      
      // Lấy giá trị từ dataObj
      let val = dataObj[col.key];
      
      // Xử lý giá trị mặc định hoặc null
      if (val === undefined || val === null) return '';
      
      return val;
    });
  },

  /**
   * Áp dụng định dạng cho Sheet dựa trên Config
   * @param {Sheet} sheet - Google Sheet object
   * @param {string} sheetKey - Key trong SHEET_CONFIG
   */
  applySheetFormat(sheet, sheetKey) {
    const config = SHEET_CONFIG[sheetKey];
    if (!config) return;
    
    const cols = config.columns;
    
    // 1. Set Headers
    const headers = cols.map(c => c.header);
    sheet.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBackground(APP_CONFIG.COLORS.HEADER_BG || '#4472C4')
      .setFontColor(APP_CONFIG.COLORS.HEADER_TEXT || '#FFFFFF');
      
    // 2. Set Column Widths & Hide Columns
    cols.forEach((col, index) => {
      const colIndex = index + 1;
      if (col.width === 0 || col.type === 'hidden') {
        sheet.hideColumns(colIndex);
      } else {
        sheet.setColumnWidth(colIndex, col.width);
      }
    });
    
    // 3. Set Formats (Number, Date...)
    // Áp dụng cho toàn bộ cột (từ dòng 2)
    const lastRow = sheet.getMaxRows();
    if (lastRow > 1) {
      cols.forEach((col, index) => {
        if (col.format) {
          sheet.getRange(2, index + 1, lastRow - 1).setNumberFormat(col.format);
        }
      });
    }
    
    // 4. Set Validations
    cols.forEach((col, index) => {
      if (col.type === 'dropdown' && col.source) {
        let rule = null;
        
        // Source là mảng tĩnh
        if (Array.isArray(col.source)) {
          rule = SpreadsheetApp.newDataValidation()
            .requireValueInList(col.source)
            .setAllowInvalid(false)
            .build();
        } 
        // Source là tham chiếu đến APP_CONFIG (VD: 'INCOME', 'EXPENSE')
        else if (typeof col.source === 'string') {
          let list = [];
          
          // Check APP_CONFIG.CATEGORIES first
          if (APP_CONFIG.CATEGORIES[col.source]) {
            list = APP_CONFIG.CATEGORIES[col.source];
          }
          // Check APP_CONFIG root keys (e.g. STOCK_TRANSACTION_TYPES)
          else if (APP_CONFIG[col.source]) {
            list = APP_CONFIG[col.source];
          }
          // Special case for LOAN_TYPES keys
          else if (col.source === 'LOAN_TYPES') {
            list = Object.keys(LOAN_TYPES || {});
          }
          
          if (list.length > 0) {
            rule = SpreadsheetApp.newDataValidation()
              .requireValueInList(list)
              .setAllowInvalid(false)
              .build();
          }
        }
        
        if (rule) {
          sheet.getRange(2, index + 1, lastRow - 1).setDataValidation(rule);
        }
      }
    });
    
    // 5. Set Formulas
    // Áp dụng công thức cho toàn bộ cột (từ dòng 2) nếu có
    if (lastRow > 1) {
      cols.forEach((col, index) => {
        if (col.formula) {
          // Sử dụng R1C1 notation nếu công thức chứa RC
          if (col.formula.includes('RC')) {
            sheet.getRange(2, index + 1, lastRow - 1).setFormulaR1C1(col.formula);
          } else {
            sheet.getRange(2, index + 1, lastRow - 1).setFormula(col.formula);
          }
        }
      });
    }

    // 6. Freeze Header
    sheet.setFrozenRows(1);
  }
};
