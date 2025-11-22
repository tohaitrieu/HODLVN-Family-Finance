/**
 * ===============================================
 * SYNC MANAGER
 * ===============================================
 * 
 * Xử lý đồng bộ dữ liệu giữa các sheet và tự động cập nhật STT
 */

var SyncManager = {
  
  /**
   * Hàm xử lý chính cho trigger onChange
   */
  handleOnChange: function(e) {
    try {
      if (!e) return;
      
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getActiveSheet();
      const sheetName = sheet.getName();
      
      // 1. Luôn cập nhật STT khi có thay đổi dòng (REMOVE_ROW hoặc INSERT_ROW)
      if (e.changeType === 'REMOVE_ROW' || e.changeType === 'INSERT_ROW') {
        this.updateSheetSTT(sheet);
      }
      
      // 2. Xử lý đồng bộ xóa dữ liệu (Chỉ khi xóa dòng)
      if (e.changeType === 'REMOVE_ROW') {
        
        // Case 1: Trả nợ <-> Chi tiêu
        if (sheetName === APP_CONFIG.SHEETS.DEBT_PAYMENT || sheetName === APP_CONFIG.SHEETS.EXPENSE) {
          this.syncDebtPaymentAndExpense(ss);
        }
        
        // Case 2: Quản lý nợ (Vay) <-> Thu nhập
        else if (sheetName === APP_CONFIG.SHEETS.DEBT_MANAGEMENT || sheetName === APP_CONFIG.SHEETS.INCOME) {
          this.syncDebtAndIncome(ss);
        }
        
        // Case 3: Cho vay <-> Chi tiêu
        else if (sheetName === APP_CONFIG.SHEETS.LENDING || sheetName === APP_CONFIG.SHEETS.EXPENSE) {
          this.syncLendingAndExpense(ss);
        }
        
        // Case 4: Thu hồi nợ (Lending Payment) <-> Thu nhập
        // Note: Thu hồi nợ hiện tại chưa có sheet riêng mà ghi trực tiếp vào Lending/Income
        // Nên logic này phức tạp hơn, tạm thời chưa xử lý tự động xóa ngược từ Income -> Lending update
      }
      
    } catch (error) {
      Logger.log('❌ Lỗi handleOnChange: ' + error.message);
    }
  },
  
  /**
   * Cập nhật lại cột STT cho sheet
   */
  updateSheetSTT: function(sheet) {
    try {
      // Kiểm tra xem cột A có phải là STT không
      const header = sheet.getRange(1, 1).getValue();
      if (header !== 'STT') return;
      
      const lastRow = sheet.getLastRow();
      if (lastRow < 2) return;
      
      // Lấy toàn bộ cột A
      const range = sheet.getRange(2, 1, lastRow - 1, 1);
      const values = range.getValues();
      let hasChange = false;
      
      // Tạo mảng STT mới
      const newValues = values.map((row, index) => {
        const newSTT = index + 1;
        if (row[0] !== newSTT) {
          hasChange = true;
        }
        return [newSTT];
      });
      
      // Chỉ ghi lại nếu có thay đổi để tối ưu hiệu năng
      if (hasChange) {
        range.setValues(newValues);
        Logger.log(`🔄 Đã cập nhật STT cho sheet ${sheet.getName()}`);
      }
      
    } catch (error) {
      Logger.log('Lỗi updateSheetSTT: ' + error.message);
    }
  },
  
  /**
   * Đồng bộ giữa Trả nợ và Chi tiêu
   * Logic: So sánh danh sách giao dịch dựa trên (Ngày + Số tiền + Tên nợ)
   * Nếu bên nào dư ra (do bị xóa bên kia) thì xóa luôn bên này.
   */
  syncDebtPaymentAndExpense: function(ss) {
    const paymentSheet = ss.getSheetByName(APP_CONFIG.SHEETS.DEBT_PAYMENT);
    const expenseSheet = ss.getSheetByName(APP_CONFIG.SHEETS.EXPENSE);
    
    if (!paymentSheet || !expenseSheet) return;
    
    // 1. Lấy dữ liệu Trả nợ
    // Col B (Date), Col C (Debt Name), Col F (Total Amount)
    const paymentRows = this.getSheetData(paymentSheet);
    const paymentSigs = paymentRows.map((row, index) => ({
      id: index + 2, // Row number
      date: row[1],
      name: row[2],
      amount: row[5],
      sig: this.createSignature(row[1], row[5], row[2])
    }));
    
    // 2. Lấy dữ liệu Chi tiêu (Chỉ lấy loại Trả nợ)
    // Col B (Date), Col C (Amount), Col D (Category), Col E (Subcategory)
    const expenseRows = this.getSheetData(expenseSheet);
    const expenseSigs = [];
    
    expenseRows.forEach((row, index) => {
      // Kiểm tra Category là "Trả nợ" hoặc Subcategory bắt đầu bằng "Trả nợ:"
      const category = row[3];
      const subcategory = row[4];
      
      if (category === 'Trả nợ' || (subcategory && subcategory.toString().startsWith('Trả nợ:'))) {
        // Parse Debt Name from Subcategory "Trả nợ: [Name]"
        let debtName = '';
        if (subcategory && subcategory.toString().startsWith('Trả nợ:')) {
          debtName = subcategory.toString().substring(8).trim(); // Length of "Trả nợ: " is 8
        }
        
        expenseSigs.push({
          id: index + 2,
          date: row[1],
          amount: row[2],
          name: debtName,
          sig: this.createSignature(row[1], row[2], debtName)
        });
      }
    });
    
    this.syncLists(paymentSheet, paymentSigs, expenseSheet, expenseSigs);
  },
  
  /**
   * Đồng bộ giữa Quản lý nợ (Vay) và Thu nhập
   */
  syncDebtAndIncome: function(ss) {
    const debtSheet = ss.getSheetByName(APP_CONFIG.SHEETS.DEBT_MANAGEMENT);
    const incomeSheet = ss.getSheetByName(APP_CONFIG.SHEETS.INCOME);
    
    if (!debtSheet || !incomeSheet) return;
    
    // 1. Lấy dữ liệu Nợ
    // Col B (Name), Col D (Amount), Col G (Date)
    const debtRows = this.getSheetData(debtSheet);
    const debtSigs = debtRows.map((row, index) => ({
      id: index + 2,
      date: row[6], // Col G
      name: row[1], // Col B
      amount: row[3], // Col D
      sig: this.createSignature(row[6], row[3], row[1])
    }));
    
    // 2. Lấy dữ liệu Thu nhập (Nguồn = Vay ngân hàng/Vay cá nhân hoặc Note bắt đầu bằng "Vay:")
    // Col B (Date), Col C (Amount), Col D (Source), Col E (Note)
    const incomeRows = this.getSheetData(incomeSheet);
    const incomeSigs = [];
    
    incomeRows.forEach((row, index) => {
      const source = row[3];
      const note = row[4];
      
      if (source === 'Vay ngân hàng' || source === 'Vay cá nhân' || (note && note.toString().startsWith('Vay:'))) {
        // Parse Name from Note "Vay: [Name]"
        let name = '';
        if (note && note.toString().startsWith('Vay:')) {
          // Note format often: "Vay: [Name]. [Other info]"
          const parts = note.toString().split('.');
          name = parts[0].substring(4).trim(); // Remove "Vay:"
        }
        
        incomeSigs.push({
          id: index + 2,
          date: row[1],
          amount: row[2],
          name: name,
          sig: this.createSignature(row[1], row[2], name)
        });
      }
    });
    
    this.syncLists(debtSheet, debtSigs, incomeSheet, incomeSigs);
  },
  
  /**
   * Đồng bộ giữa Cho vay và Chi tiêu
   */
  syncLendingAndExpense: function(ss) {
    const lendingSheet = ss.getSheetByName(APP_CONFIG.SHEETS.LENDING);
    const expenseSheet = ss.getSheetByName(APP_CONFIG.SHEETS.EXPENSE);
    
    if (!lendingSheet || !expenseSheet) return;
    
    // 1. Lấy dữ liệu Cho vay
    // Col B (Name), Col D (Amount), Col G (Date)
    const lendingRows = this.getSheetData(lendingSheet);
    const lendingSigs = lendingRows.map((row, index) => ({
      id: index + 2,
      date: row[6], // Col G
      name: row[1], // Col B
      amount: row[3], // Col D
      sig: this.createSignature(row[6], row[3], row[1])
    }));
    
    // 2. Lấy dữ liệu Chi tiêu (Category = Cho vay)
    // Col B (Date), Col C (Amount), Col E (Subcategory: "Cho vay: [Name]")
    const expenseRows = this.getSheetData(expenseSheet);
    const expenseSigs = [];
    
    expenseRows.forEach((row, index) => {
      const category = row[3];
      const subcategory = row[4];
      
      if (category === 'Cho vay' || (subcategory && subcategory.toString().startsWith('Cho vay:'))) {
        let name = '';
        if (subcategory && subcategory.toString().startsWith('Cho vay:')) {
          name = subcategory.toString().substring(9).trim(); // "Cho vay: " is 9 chars
        }
        
        expenseSigs.push({
          id: index + 2,
          date: row[1],
          amount: row[2],
          name: name,
          sig: this.createSignature(row[1], row[2], name)
        });
      }
    });
    
    this.syncLists(lendingSheet, lendingSigs, expenseSheet, expenseSigs);
  },
  
  /**
   * Hàm so sánh và xóa dòng dư thừa
   */
  syncLists: function(sheetA, listA, sheetB, listB) {
    // Đếm số lượng signature
    const countA = this.countSignatures(listA);
    const countB = this.countSignatures(listB);
    
    // Tìm signature bị lệch
    const allSigs = new Set([...Object.keys(countA), ...Object.keys(countB)]);
    
    allSigs.forEach(sig => {
      const cA = countA[sig] || 0;
      const cB = countB[sig] || 0;
      
      if (cA > cB) {
        // A nhiều hơn B -> Xóa bớt ở A (vì B đã bị xóa)
        const diff = cA - cB;
        this.deleteRowsBySignature(sheetA, listA, sig, diff);
        Logger.log(`🗑️ Đã xóa ${diff} dòng đồng bộ ở ${sheetA.getName()} (Sig: ${sig})`);
      } else if (cB > cA) {
        // B nhiều hơn A -> Xóa bớt ở B (vì A đã bị xóa)
        const diff = cB - cA;
        this.deleteRowsBySignature(sheetB, listB, sig, diff);
        Logger.log(`🗑️ Đã xóa ${diff} dòng đồng bộ ở ${sheetB.getName()} (Sig: ${sig})`);
      }
    });
  },
  
  /**
   * Helper: Lấy data sheet (bỏ header)
   */
  getSheetData: function(sheet) {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    const lastCol = sheet.getLastColumn();
    return sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  },
  
  /**
   * Helper: Tạo signature định danh giao dịch
   */
  createSignature: function(date, amount, name) {
    // Chuẩn hóa date: YYYY-MM-DD
    let d = new Date(date);
    if (isNaN(d.getTime())) d = new Date(); // Fallback
    const dateStr = `${d.getFullYear()}-${d.getMonth()}-${d.getDate()}`;
    
    // Chuẩn hóa amount: integer
    const amt = Math.round(parseFloat(amount) || 0);
    
    // Chuẩn hóa name: lowercase, trim
    const n = (name || '').toString().toLowerCase().trim();
    
    return `${dateStr}_${amt}_${n}`;
  },
  
  /**
   * Helper: Đếm số lượng mỗi signature
   */
  countSignatures: function(list) {
    const count = {};
    list.forEach(item => {
      count[item.sig] = (count[item.sig] || 0) + 1;
    });
    return count;
  },
  
  /**
   * Helper: Xóa dòng theo signature
   * Xóa từ dưới lên để không ảnh hưởng index
   */
  deleteRowsBySignature: function(sheet, list, sig, countToDelete) {
    let deleted = 0;
    // Duyệt ngược từ dưới lên
    for (let i = list.length - 1; i >= 0; i--) {
      if (deleted >= countToDelete) break;
      
      if (list[i].sig === sig) {
        sheet.deleteRow(list[i].id);
        deleted++;
      }
    }
  }
};
