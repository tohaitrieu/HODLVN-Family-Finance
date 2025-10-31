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
   */
  updateBudgetSpent(category) {
    try {
      const now = new Date();
      const currentMonth = now.getMonth() + 1;
      const currentYear = now.getFullYear();
      
      const budgetSheet = getSheet(APP_CONFIG.SHEETS.BUDGET);
      if (!budgetSheet) return;
      
      const budgetData = budgetSheet.getDataRange().getValues();
      
      // Tìm danh mục trong Budget
      for (let i = 4; i < budgetData.length; i++) {
        if (budgetData[i][0] === category) {
          const spent = this._calculateCategorySpent(category, currentMonth, currentYear);
          budgetSheet.getRange(i + 1, 3).setValue(spent);
          break;
        }
      }
      
    } catch (error) {
      Logger.log('Error updating budget: ' + error.message);
    }
  },
  
  /**
   * Cập nhật "Đã đầu tư" cho loại đầu tư
   */
  updateInvestmentBudget(investmentType, amount) {
    try {
      const now = new Date();
      const currentMonth = now.getMonth() + 1;
      const currentYear = now.getFullYear();
      
      const budgetSheet = getSheet(APP_CONFIG.SHEETS.BUDGET);
      if (!budgetSheet) return;
      
      const currentInvested = this._calculateInvestmentSpent(investmentType, currentMonth, currentYear);
      
      // Map loại đầu tư với row trong Budget
      const rowMap = {
        'Chứng khoán': this._findBudgetRow(budgetSheet, 'Chứng khoán'),
        'Vàng': this._findBudgetRow(budgetSheet, 'Vàng'),
        'Crypto': this._findBudgetRow(budgetSheet, 'Crypto'),
        'Đầu tư khác': this._findBudgetRow(budgetSheet, 'Đầu tư khác')
      };
      
      const row = rowMap[investmentType];
      if (row) {
        budgetSheet.getRange(row, 3).setValue(currentInvested);
      }
      
    } catch (error) {
      Logger.log('Error updating investment budget: ' + error.message);
    }
  },
  
  /**
   * Cập nhật "Đã trả nợ" và "Đã trả lãi"
   */
  updateDebtBudget() {
    try {
      const now = new Date();
      const currentMonth = now.getMonth() + 1;
      const currentYear = now.getFullYear();
      
      const budgetSheet = getSheet(APP_CONFIG.SHEETS.BUDGET);
      if (!budgetSheet) return;
      
      // Tính tổng trả nợ gốc và lãi trong tháng
      const {principal, interest} = this._calculateDebtPayment(currentMonth, currentYear);
      
      // Cập nhật "Trả nợ gốc"
      const principalRow = this._findBudgetRow(budgetSheet, 'Trả nợ gốc');
      if (principalRow) {
        budgetSheet.getRange(principalRow, 3).setValue(principal);
      }
      
      // Cập nhật "Trả lãi"
      const interestRow = this._findBudgetRow(budgetSheet, 'Trả lãi');
      if (interestRow) {
        budgetSheet.getRange(interestRow, 3).setValue(interest);
      }
      
    } catch (error) {
      Logger.log('Error updating debt budget: ' + error.message);
    }
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
        const currentValue = budgetSheet.getRange(row, 3).getValue() || 0;
        budgetSheet.getRange(row, 3).setValue(currentValue + amount);
      }
      
    } catch (error) {
      Logger.log('Error updating reserve budget: ' + error.message);
    }
  },
  
  /**
   * Kiểm tra cảnh báo budget
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
        const category = data[i][0];
        const budget = data[i][1];
        const spent = data[i][2];
        
        // Dừng khi gặp section tiếp theo
        if (category && (category.includes('NỢ') || category === 'Tổng')) {
          break;
        }
        
        if (category && budget && spent && category !== 'Tổng') {
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
          const budget = budgetSheet.getRange(row, 2).getValue();
          const paid = budgetSheet.getRange(row, 3).getValue();
          
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
          const target = budgetSheet.getRange(row, 2).getValue();
          const saved = budgetSheet.getRange(row, 3).getValue();
          
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
          const target = budgetSheet.getRange(row, 2).getValue();
          const invested = budgetSheet.getRange(row, 3).getValue();
          
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
   * @param {Object} budgetData - Dữ liệu ngân sách từ form
   * @return {Object} {success, message}
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
      
      // Tính số tiền cho mỗi nhóm
      const totalChi = income * pctChi;
      const totalDautu = income * pctDautu;
      const totalTrano = income * pctTrano;
      
      // ===== CẬP NHẬT CHI TIÊU =====
      const expenseCategories = [
        'Ăn uống', 'Đi lại', 'Nhà ở', 'Điện nước', 'Viễn thông',
        'Giáo dục', 'Y tế', 'Mua sắm', 'Giải trí', 'Khác'
      ];
      
      for (let category of expenseCategories) {
        const row = this._findBudgetRow(budgetSheet, category);
        if (row && budgetData.chi && budgetData.chi[category] !== undefined) {
          const categoryPct = parseFloat(budgetData.chi[category]) / 100;
          const categoryBudget = totalChi * categoryPct;
          
          // Cập nhật cột B (Mục tiêu)
          budgetSheet.getRange(row, 2).setValue(categoryBudget);
          
          // Cột C (Đã chi) sẽ tự động cập nhật qua SUMIFS formula
        }
      }
      
      // ===== CẬP NHẬT ĐẦU TƯ =====
      const investmentTypes = ['Chứng khoán', 'Vàng', 'Crypto', 'Đầu tư khác'];
      
      for (let type of investmentTypes) {
        const row = this._findBudgetRow(budgetSheet, type);
        if (row && budgetData.dautu && budgetData.dautu[type] !== undefined) {
          const typePct = parseFloat(budgetData.dautu[type]) / 100;
          const typeBudget = totalDautu * typePct;
          
          // Cập nhật cột B (Mục tiêu)
          budgetSheet.getRange(row, 2).setValue(typeBudget);
          
          // Cột C (Đã đầu tư) sẽ tự động cập nhật qua hàm updateInvestmentBudget
        }
      }
      
      // ===== CẬP NHẬT TRẢ NỢ =====
      const debtRow = this._findBudgetRow(budgetSheet, 'Trả nợ gốc');
      if (debtRow) {
        budgetSheet.getRange(debtRow, 2).setValue(totalTrano);
      }
      
      // ===== GHI CHÚ THAM SỐ =====
      // Lưu các tham số vào vùng riêng để tham khảo
      const paramStartRow = 2;
      budgetSheet.getRange(paramStartRow, 7).setValue('Thu nhập tháng:');
      budgetSheet.getRange(paramStartRow, 8).setValue(income);
      
      budgetSheet.getRange(paramStartRow + 1, 7).setValue('% Chi tiêu:');
      budgetSheet.getRange(paramStartRow + 1, 8).setValue(pctChi);
      
      budgetSheet.getRange(paramStartRow + 2, 7).setValue('% Đầu tư:');
      budgetSheet.getRange(paramStartRow + 2, 8).setValue(pctDautu);
      
      budgetSheet.getRange(paramStartRow + 3, 7).setValue('% Trả nợ:');
      budgetSheet.getRange(paramStartRow + 3, 8).setValue(pctTrano);
      
      // Format số
      budgetSheet.getRange(paramStartRow, 8).setNumberFormat('#,##0" VNĐ"');
      budgetSheet.getRange(paramStartRow + 1, 8, 3, 1).setNumberFormat('0.0%');
      
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
  }
};