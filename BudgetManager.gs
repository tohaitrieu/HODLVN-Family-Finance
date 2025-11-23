/**
 * ===============================================
 * BUDGETMANAGER.GS v3.5 - MODULE QUẢN LÝ NGÂN SÁCH
 * ===============================================
 * 
 * CHANGELOG v3.5:
 * ✅ NEW: setBudgetForMonth() - Thiết lập ngân sách từ SetBudgetForm
 * ✅ FIX: showExpenseReport() - Hoàn thiện báo cáo chi tiêu
 * ✅ FIX: showInvestmentReport() - Hoàn thiện báo cáo đầu tư
 * ✅ COMPLETE: Tất cả menu Budget hoạt động đầy đủ
 * 
 * Chức năng:
 * - Cập nhật budget chi tiêu
 * - Cập nhật budget đầu tư
 * - Kiểm tra cảnh báo vượt ngân sách
 * - Tạo báo cáo chi tiêu và đầu tư
 * - Thiết lập ngân sách tháng từ form
 */

const BudgetManager = {
  
  /**
   * Cập nhật "Đã chi" cho danh mục
   * @deprecated v3.5.1 - Sheet BUDGET sử dụng công thức tự động cập nhật
   */
  updateBudgetSpent(category) {
    // Không làm gì cả vì Sheet BUDGET đã dùng công thức SUMIFS
    Logger.log('BudgetManager: updateBudgetSpent skipped (using formulas)');
  },
  
  /**
   * Cập nhật "Đã đầu tư" cho loại đầu tư
   * @deprecated v3.5.1 - Sheet BUDGET sử dụng công thức tự động cập nhật
   */
  updateInvestmentBudget(investmentType, amount) {
    // Không làm gì cả vì Sheet BUDGET đã dùng công thức SUMIFS
    Logger.log('BudgetManager: updateInvestmentBudget skipped (using formulas)');
  },
  
  /**
   * Cập nhật "Đã trả nợ" và "Đã trả lãi"
   * @deprecated v3.5.1 - Sheet BUDGET sử dụng công thức tự động cập nhật
   */
  updateDebtBudget() {
    // Không làm gì cả vì Sheet BUDGET đã dùng công thức SUMIFS
    Logger.log('BudgetManager: updateDebtBudget skipped (using formulas)');
  },
  
  /**
   * Cập nhật "Quỹ dự phòng"
   * Lưu ý: Quỹ dự phòng thường được cập nhật thủ công hoặc từ sheet khác
   * Hàm này để dự phòng cho tương lai
   */
  updateReserveBudget(reserveType, amount) {
    try {
      const budgetSheet = getSheet(APP_CONFIG.SHEETS.BUDGET);
      if (!budgetSheet) return;
      
      const row = this._findBudgetRow(budgetSheet, reserveType);
      if (row) {
        // Quỹ dự phòng có thể vẫn dùng value thay vì formula
        const currentValue = budgetSheet.getRange(row, 3).getValue() || 0;
        budgetSheet.getRange(row, 3).setValue(currentValue + amount);
      }
      
    } catch (error) {
      Logger.log('Error updating reserve budget: ' + error.message);
    }
  },
  
  /**
   * Kiểm tra cảnh báo budget
   * ✅ FIXED v3.5.1: Cập nhật đúng cột C (Ngân sách) và D (Đã chi)
   */
  checkBudgetWarnings() {
    try {
      const budgetSheet = getSheet(APP_CONFIG.SHEETS.BUDGET);
      if (!budgetSheet) {
        showError('Lỗi', 'Không tìm thấy sheet BUDGET');
        return;
      }
      
      const data = budgetSheet.getDataRange().getValues();
      
      let warnings = '⚠️ CẢNH BÁO BUDGET:\n\n';
      let hasWarning = false;
      
      // === CHI TIÊU ===
      warnings += '=== 📤 CHI TIÊU ===\n';
      let expenseWarning = false;
      
      for (let i = 4; i < data.length; i++) {
        const category = data[i][0]; // Col A
        const budget = data[i][2];   // Col C: Ngân sách (Target)
        const spent = data[i][3];    // Col D: Đã chi (Spent)
        
        // Dừng khi gặp section tiếp theo
        if (category && (category.includes('NỢ') || category === 'Tổng' || category === 'TỔNG CHI')) {
          if (category === 'TỔNG CHI') continue; // Skip summary row
          break;
        }
        
        if (category && budget && spent !== undefined && category !== 'Tổng') {
          const percentage = spent / budget;
          
          if (percentage > 1) {
            warnings += `🔴 ${category}: Vượt ${((percentage - 1) * 100).toFixed(1)}%\n`;
            hasWarning = true;
            expenseWarning = true;
          } else if (percentage > 0.8) {
            warnings += `🟡 ${category}: Đã dùng ${(percentage * 100).toFixed(1)}%\n`;
            hasWarning = true;
            expenseWarning = true;
          }
        }
      }
      
      if (!expenseWarning) {
        warnings += '✅ Tất cả trong mức an toàn\n';
      }
      
      // === NỢ & LÃI ===
      warnings += '\n=== 💳 NỢ & LÃI ===\n';
      let debtWarning = false;
      
      const debtCategories = ['Trả nợ gốc', 'Trả lãi'];
      for (let cat of debtCategories) {
        const row = this._findBudgetRow(budgetSheet, cat);
        if (row) {
          const budget = budgetSheet.getRange(row, 3).getValue(); // Col C
          const paid = budgetSheet.getRange(row, 4).getValue();   // Col D
          
          if (budget && paid !== null && paid !== undefined) {
            const percentage = paid / budget;
            
            if (percentage > 1) {
              warnings += `🔴 ${cat}: Vượt ${((percentage - 1) * 100).toFixed(1)}%\n`;
              hasWarning = true;
              debtWarning = true;
            } else if (percentage > 0.8) {
              warnings += `🟡 ${cat}: Đã trả ${(percentage * 100).toFixed(1)}%\n`;
              hasWarning = true;
              debtWarning = true;
            }
          }
        }
      }
      
      if (!debtWarning) {
        warnings += '✅ Trả nợ đúng kế hoạch\n';
      }
      
      // === QUỸ DỰ PHÒNG ===
      warnings += '\n=== 🛡️ QUỸ DỰ PHÒNG ===\n';
      let reserveWarning = false;
      
      const reserveCategories = ['Quỹ khẩn cấp', 'Quỹ dự phòng'];
      for (let cat of reserveCategories) {
        const row = this._findBudgetRow(budgetSheet, cat);
        if (row) {
          const target = budgetSheet.getRange(row, 3).getValue(); // Col C
          const saved = budgetSheet.getRange(row, 4).getValue();  // Col D
          
          if (target && saved !== null && saved !== undefined) {
            const percentage = saved / target;
            
            if (percentage < 0.5) {
              warnings += `🔴 ${cat}: Chỉ đạt ${(percentage * 100).toFixed(1)}%\n`;
              hasWarning = true;
              reserveWarning = true;
            } else if (percentage < 0.8) {
              warnings += `🟡 ${cat}: Đạt ${(percentage * 100).toFixed(1)}%\n`;
              hasWarning = true;
              reserveWarning = true;
            }
          }
        }
      }
      
      if (!reserveWarning) {
        warnings += '✅ Quỹ dự phòng đầy đủ\n';
      }
      
      // === ĐẦU TƯ ===
      warnings += '\n=== 💼 ĐẦU TƯ ===\n';
      let investWarning = false;
      
      const investCategories = ['Chứng khoán', 'Vàng', 'Crypto', 'Đầu tư khác'];
      
      for (let cat of investCategories) {
        const row = this._findBudgetRow(budgetSheet, cat);
        if (row) {
          const target = budgetSheet.getRange(row, 3).getValue(); // Col C
          const invested = budgetSheet.getRange(row, 4).getValue(); // Col D
          
          if (target && invested !== null && invested !== undefined) {
            const percentage = invested / target;
            
            if (percentage < 0.8) {
              warnings += `🔴 ${cat}: Chưa đạt (${(percentage * 100).toFixed(1)}%)\n`;
              hasWarning = true;
              investWarning = true;
            } else if (percentage < 1) {
              warnings += `🟡 ${cat}: Gần đạt (${(percentage * 100).toFixed(1)}%)\n`;
              hasWarning = true;
              investWarning = true;
            }
          }
        }
      }
      
      if (!investWarning) {
        warnings += '✅ Đầu tư đạt mục tiêu\n';
      }
      
      if (!hasWarning) {
        warnings = '✅ XUẤT SẮC!\n\nTất cả các mục đều trong tình trạng tốt:\n' +
                   '✓ Chi tiêu hợp lý\n' +
                   '✓ Trả nợ đúng hạn\n' +
                   '✓ Quỹ dự phòng đầy đủ\n' +
                   '✓ Đầu tư đạt mục tiêu';
      }
      
      SpreadsheetApp.getUi().alert('Kiểm tra Budget', warnings, SpreadsheetApp.getUi().ButtonSet.OK);
      
    } catch (error) {
      showError('Lỗi', error.message);
    }
  },
  
  /**
   * Hiển thị báo cáo chi tiêu
   */
  showExpenseReport() {
    try {
      const now = new Date();
      const currentMonth = now.getMonth() + 1;
      const currentYear = now.getFullYear();
      
      let report = `📊 BÁO CÁO CHI TIÊU THÁNG ${currentMonth}/${currentYear}\n\n`;
      
      const categories = [
        'Ăn uống', 'Đi lại', 'Nhà ở', 'Y tế',
        'Giáo dục', 'Mua sắm', 'Giải trí', 'Khác'
      ];
      
      let total = 0;
      
      for (let cat of categories) {
        const spent = this._calculateCategorySpent(cat, currentMonth, currentYear);
        if (spent > 0) {
          report += `${cat}: ${formatCurrency(spent)}\n`;
          total += spent;
        }
      }
      
      report += `\n━━━━━━━━━━━━━━━━━━━━\n`;
      report += `TỔNG: ${formatCurrency(total)}`;
      
      SpreadsheetApp.getUi().alert('Báo cáo Chi tiêu', report, SpreadsheetApp.getUi().ButtonSet.OK);
      
    } catch (error) {
      showError('Lỗi', error.message);
    }
  },
  
  /**
   * Hiển thị báo cáo đầu tư
   */
  showInvestmentReport() {
    try {
      const now = new Date();
      const currentMonth = now.getMonth() + 1;
      const currentYear = now.getFullYear();
      
      let report = `💰 BÁO CÁO ĐẦU TƯ THÁNG ${currentMonth}/${currentYear}\n\n`;
      
      const investments = {
        'Chứng khoán': this._calculateInvestmentSpent('Chứng khoán', currentMonth, currentYear),
        'Vàng': this._calculateInvestmentSpent('Vàng', currentMonth, currentYear),
        'Crypto': this._calculateInvestmentSpent('Crypto', currentMonth, currentYear),
        'Đầu tư khác': this._calculateInvestmentSpent('Đầu tư khác', currentMonth, currentYear)
      };
      
      let total = 0;
      
      for (let [type, amount] of Object.entries(investments)) {
        if (amount > 0) {
          report += `${type}: ${formatCurrency(amount)}\n`;
          total += amount;
        }
      }
      
      report += `\n━━━━━━━━━━━━━━━━━━━━\n`;
      report += `TỔNG: ${formatCurrency(total)}`;
      
      SpreadsheetApp.getUi().alert('Báo cáo Đầu tư', report, SpreadsheetApp.getUi().ButtonSet.OK);
      
    } catch (error) {
      showError('Lỗi', error.message);
    }
  },
  
  /**
   * Thiết lập ngân sách tháng từ SetBudgetForm
   * ✅ FIXED v3.5.1: Cập nhật đúng cột % (Col B) thay vì ghi đè cột Ngân sách (Col C)
   */
  setBudgetForMonth(budgetData) {
    try {
      const budgetSheet = getSheet(APP_CONFIG.SHEETS.BUDGET);
      if (!budgetSheet) {
        return {
          success: false,
          message: '❌ Sheet BUDGET chưa được khởi tạo!'
        };
      }
      
      const income = parseFloat(budgetData.income);
      const pctChi = parseFloat(budgetData.pctChi) / 100;
      const pctDautu = parseFloat(budgetData.pctDautu) / 100;
      const pctTrano = parseFloat(budgetData.pctTrano) / 100;
      
      // Cập nhật Thu nhập (B2) và % Nhóm (B3, B4, B5)
      // Lưu ý: SheetInitializer dùng B3, B{dautuRow} cho % nhóm
      // Cần tìm đúng dòng % Nhóm
      
      budgetSheet.getRange('B2').setValue(income); // Thu nhập
      
      // Tìm dòng % Chi tiêu (Row 3)
      budgetSheet.getRange('B3').setValue(pctChi);
      
      // ===== CẬP NHẬT CHI TIÊU =====
      const expenseCategories = [
        'Ăn uống', 'Đi lại', 'Nhà ở', 'Điện nước', 'Viễn thông',
        'Giáo dục', 'Y tế', 'Mua sắm', 'Giải trí', 'Khác'
      ];
      
      for (let category of expenseCategories) {
        const row = this._findBudgetRow(budgetSheet, category);
        if (row && budgetData.chi && budgetData.chi[category] !== undefined) {
          const categoryPct = parseFloat(budgetData.chi[category]) / 100;
          
          // ✅ FIX: Cập nhật cột B (% Danh mục)
          budgetSheet.getRange(row, 2).setValue(categoryPct);
          
          // Cột C (Ngân sách) sẽ tự động cập nhật qua công thức: =Income * %Group * %Category
          // Cột D (Đã chi) sẽ tự động cập nhật qua công thức SUMIFS
        }
      }
      
      // ===== CẬP NHẬT ĐẦU TƯ =====
      // Tìm dòng % Đầu tư
      // Đầu tư bắt đầu sau phần Chi tiêu. Tìm dòng "Nhóm Đầu tư:"
      const data = budgetSheet.getDataRange().getValues();
      let dautuGroupRow = -1;
      for(let i=0; i<data.length; i++) {
        if(data[i][0] === 'Nhóm Đầu tư:') {
          dautuGroupRow = i + 1;
          break;
        }
      }
      
      if (dautuGroupRow > 0) {
        budgetSheet.getRange(dautuGroupRow, 2).setValue(pctDautu);
      }
      
      const investmentTypes = ['Chứng khoán', 'Vàng', 'Crypto', 'Đầu tư khác'];
      
      for (let type of investmentTypes) {
        const row = this._findBudgetRow(budgetSheet, type);
        if (row && budgetData.dautu && budgetData.dautu[type] !== undefined) {
          const typePct = parseFloat(budgetData.dautu[type]) / 100;
          
          // ✅ FIX: Cập nhật cột B (% Danh mục)
          budgetSheet.getRange(row, 2).setValue(typePct);
        }
      }
      
      // ===== CẬP NHẬT TRẢ NỢ =====
      // Trả nợ thường là số tiền cố định, không theo % nhóm
      // Nhưng trong SetupWizard/SetBudgetForm đang dùng % tổng thu nhập
      const debtRow = this._findBudgetRow(budgetSheet, 'Trả nợ gốc');
      if (debtRow) {
         const totalTrano = income * pctTrano;
         // Trả nợ gốc thường nhập số tiền trực tiếp vào Col C (Ngân sách)
         // Vì không có cột % cho từng khoản nợ trong cấu trúc hiện tại
         budgetSheet.getRange(debtRow, 3).setValue(totalTrano);
      }
      
      Logger.log('✅ Đã thiết lập ngân sách thành công');
      
      return {
        success: true,
        message: '✅ Đã thiết lập ngân sách tháng thành công!'
      };
      
    } catch (error) {
      Logger.log('Error in setBudgetForMonth: ' + error.message);
      return {
        success: false,
        message: `❌ Lỗi: ${error.message}`
      };
    }
  },
  
  // ==================== HELPER FUNCTIONS ====================
  
  /**
   * Tính tổng chi cho danh mục trong tháng
   */
  _calculateCategorySpent(category, month, year) {
    const sheet = getSheet(APP_CONFIG.SHEETS.EXPENSE);
    if (!sheet) return 0;
    
    const data = sheet.getDataRange().getValues();
    let total = 0;
    
    for (let i = 1; i < data.length; i++) {
      const date = new Date(data[i][1]);
      const amount = data[i][2];
      const cat = data[i][3];
      
      if (cat === category &&
          date.getMonth() + 1 === month &&
          date.getFullYear() === year &&
          amount) {
        total += amount;
      }
    }
    
    return total;
  },
  
  /**
   * Tính tổng đầu tư cho loại trong tháng
   */
  _calculateInvestmentSpent(investmentType, month, year) {
    const sheetMap = {
      'Chứng khoán': APP_CONFIG.SHEETS.STOCK,
      'Vàng': APP_CONFIG.SHEETS.GOLD,
      'Crypto': APP_CONFIG.SHEETS.CRYPTO,
      'Đầu tư khác': APP_CONFIG.SHEETS.OTHER_INVESTMENT
    };
    
    const sheetName = sheetMap[investmentType];
    if (!sheetName) return 0;
    
    const sheet = getSheet(sheetName);
    if (!sheet) return 0;
    
    const data = sheet.getDataRange().getValues();
    let total = 0;
    
    for (let i = 1; i < data.length; i++) {
      const date = new Date(data[i][1]);
      const type = data[i][2];
      
      if (date.getMonth() + 1 === month && date.getFullYear() === year) {
        // Đầu tư khác: không có loại GD
        if (sheetName === APP_CONFIG.SHEETS.OTHER_INVESTMENT) {
          const amount = data[i][3];
          if (amount) total += amount;
        }
        // CK, Vàng, Crypto: chỉ tính "Mua"
        else if (type === 'Mua') {
          const amount = data[i][7]; // Cột Tổng giá trị
          if (amount) total += amount;
        }
      }
    }
    
    return total;
  },
  
  /**
   * Tìm row của danh mục trong Budget
   */
  _findBudgetRow(sheet, category) {
    const data = sheet.getDataRange().getValues();
    
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === category) {
        return i + 1;
      }
    }
    
    return null;
  },
  
  /**
   * Tính tổng trả nợ gốc và lãi trong tháng
   */
  _calculateDebtPayment(month, year) {
    const sheet = getSheet(APP_CONFIG.SHEETS.DEBT_PAYMENT);
    if (!sheet) return {principal: 0, interest: 0};
    
    const data = sheet.getDataRange().getValues();
    let principal = 0;
    let interest = 0;
    
    for (let i = 1; i < data.length; i++) {
      const date = new Date(data[i][1]);
      const paidPrincipal = data[i][3]; // Cột D: Trả gốc
      const paidInterest = data[i][4];  // Cột E: Trả lãi
      
      if (date.getMonth() + 1 === month && 
          date.getFullYear() === year) {
        if (paidPrincipal) principal += paidPrincipal;
        if (paidInterest) interest += paidInterest;
      }
    }
    
    return {principal, interest};
  },

  /**
   * Lấy cấu hình ngân sách hiện tại từ sheet BUDGET
   * Để hiển thị lên SetBudgetForm
   */
  getBudgetConfig() {
    const budgetSheet = getSheet(APP_CONFIG.SHEETS.BUDGET);
    if (!budgetSheet) return null;

    const data = budgetSheet.getDataRange().getValues();
    
    // 1. Get General Info
    const income = budgetSheet.getRange('B2').getValue() || 0;
    const pctChi = (budgetSheet.getRange('B3').getValue() || 0) * 100;
    const pctDautu = (budgetSheet.getRange('B4').getValue() || 0) * 100;
    const pctTrano = (budgetSheet.getRange('B5').getValue() || 0) * 100;

    // 2. Get Expense Categories (Rows 7-16)
    const chi = {};
    const expenseCategories = [
      'Ăn uống', 'Đi lại', 'Nhà ở', 'Điện nước', 'Viễn thông',
      'Giáo dục', 'Y tế', 'Mua sắm', 'Giải trí', 'Khác'
    ];
    
    expenseCategories.forEach(cat => {
      const row = this._findBudgetRow(budgetSheet, cat);
      if (row) {
        const pct = (budgetSheet.getRange(row, 2).getValue() || 0) * 100;
        chi[cat] = Math.round(pct);
      }
    });

    // 3. Get Investment Categories (Rows 25-28)
    const dautu = {};
    const investmentTypes = ['Chứng khoán', 'Vàng', 'Crypto', 'Đầu tư khác'];
    
    investmentTypes.forEach(type => {
      const row = this._findBudgetRow(budgetSheet, type);
      if (row) {
        const pct = (budgetSheet.getRange(row, 2).getValue() || 0) * 100;
        dautu[type] = Math.round(pct);
      }
    });

    return {
      income,
      pctChi: Math.round(pctChi),
      pctDautu: Math.round(pctDautu),
      pctTrano: Math.round(pctTrano),
      chi,
      dautu
    };
  }
};