/**
 * ===============================================
 * DASHBOARDFUNCTIONS.GS - CUSTOM FUNCTIONS FOR DASHBOARD
 * ===============================================
 * 
 * Custom functions to display dashboard data in sheets.
 * These functions read filter values from B2-B4 and return detailed arrays.
 * 
 * Functions:
 * - hffsIncome(): Income breakdown by category
 * - hffsExpense(): Expense breakdown with budget comparison
 * - hffsDebt(): List of outstanding debts
 * - hffsAssets(): Asset summary
 * - hffsYearly(year): 12-month statistics
 */

// ==================== HELPER FUNCTIONS ====================

/**
 * Get filter values from dashboard cells B2-B4
 * @return {Object} {year, quarter, month, startDate, endDate}
 * @private
 */
function _getDateFilters() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(APP_CONFIG.SHEETS.DASHBOARD);
    if (!sheet) return null;
    
    const yearFilter = sheet.getRange('B2').getValue();
    const quarterFilter = sheet.getRange('B3').getValue();
    const monthFilter = sheet.getRange('B4').getValue();
    
    let startDate, endDate;
    const currentYear = typeof yearFilter === 'number' ? yearFilter : new Date().getFullYear();
    
    // Priority: Month > Quarter > Year
    if (monthFilter && monthFilter !== 'Tất cả') {
      // Extract month number from "Tháng 1", "Tháng 2", etc.
      const monthNum = parseInt(monthFilter.replace('Tháng ', ''));
      startDate = new Date(currentYear, monthNum - 1, 1);
      endDate = new Date(currentYear, monthNum, 0);
    } else if (quarterFilter && quarterFilter !== 'Tất cả') {
      // Extract quarter number from "Quý 1", "Quý 2", etc.
      const quarterNum = parseInt(quarterFilter.replace('Quý ', ''));
      const startMonth = (quarterNum - 1) * 3;
      startDate = new Date(currentYear, startMonth, 1);
      endDate = new Date(currentYear, startMonth + 3, 0);
    } else {
      // Year filter
      startDate = new Date(currentYear, 0, 1);
      endDate = new Date(currentYear, 11, 31);
    }
    
    return {
      year: currentYear,
      quarter: quarterFilter,
      month: monthFilter,
      startDate: startDate,
      endDate: endDate
    };
  } catch (error) {
    // Fallback to current month
    const today = new Date();
    return {
      year: today.getFullYear(),
      quarter: 'Tất cả',
      month: 'Tất cả',
      startDate: new Date(today.getFullYear(), 0, 1),
      endDate: new Date(today.getFullYear(), 11, 31)
    };
  }
}

// ==================== CUSTOM FUNCTIONS ====================

/**
 * Get income breakdown by category
 * Reads filter from B2-B4
 * @return {Array} 2D array: [["Category", "Amount", "Percentage"], ...]
 * @customfunction
 */
function hffsIncome() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const incomeSheet = ss.getSheetByName(APP_CONFIG.SHEETS.INCOME);
    
    if (!incomeSheet) return [['Lỗi: Sheet THU không tồn tại']];
    
    const filters = _getDateFilters();
    const data = incomeSheet.getDataRange().getValues();
    const categoryTotals = {};
    let totalIncome = 0;
    
    // Calculate totals by category
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const date = row[1]; // Column B: Date
      const amount = parseFloat(row[2]) || 0; // Column C: Amount
      const category = row[3]; // Column D: Category
      
      if (date >= filters.startDate && date <= filters.endDate) {
        categoryTotals[category] = (categoryTotals[category] || 0) + amount;
        totalIncome += amount;
      }
    }
    
    // Build result array
    const result = [];
    const categories = APP_CONFIG.CATEGORIES.INCOME;
    
    for (const category of categories) {
      const amount = categoryTotals[category] || 0;
      const percentage = totalIncome > 0 ? (amount / totalIncome) : 0;
      result.push([category, amount, percentage]);
    }
    
    // Add total row
    result.push(['TỔNG THU NHẬP', totalIncome, 1]);
    
    return result;
    
  } catch (error) {
    return [['Lỗi: ' + error.message]];
  }
}

/**
 * Get expense breakdown with budget comparison
 * Reads filter from B2-B4
 * @return {Array} 2D array: [["Category", "Spent", "Budget", "Remaining", "Status%"], ...]
 * @customfunction
 */
function hffsExpense() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const expenseSheet = ss.getSheetByName(APP_CONFIG.SHEETS.EXPENSE);
    const budgetSheet = ss.getSheetByName(APP_CONFIG.SHEETS.BUDGET);
    const debtPaymentSheet = ss.getSheetByName(APP_CONFIG.SHEETS.DEBT_PAYMENT);
    
    if (!expenseSheet) return [['Lỗi: Sheet CHI không tồn tại']];
    
    const filters = _getDateFilters();
    const categoryTotals = {};
    
    // Calculate spent by category (exclude 'Trả nợ' and 'Cho vay')
    const expenseData = expenseSheet.getDataRange().getValues();
    for (let i = 1; i < expenseData.length; i++) {
      const row = expenseData[i];
      const date = row[1]; // Column B: Date
      const amount = parseFloat(row[2]) || 0; // Column C: Amount
      const category = row[3]; // Column D: Category
      
      if (date >= filters.startDate && date <= filters.endDate &&
          category !== 'Trả nợ' && category !== 'Cho vay') {
        categoryTotals[category] = (categoryTotals[category] || 0) + amount;
      }
    }
    
    // Get budget data
    const budgetData = {};
    if (budgetSheet) {
      const budgetValues = budgetSheet.getDataRange().getValues();
      for (let i = 1; i < budgetValues.length; i++) {
        const cat = budgetValues[i][0]; // Column A: Category
        const budget = parseFloat(budgetValues[i][1]) || 0; // Column B: Budget
        if (cat && budget > 0) {
          budgetData[cat] = budget;
        }
      }
    }
    
    // Calculate debt payments
    let debtPaymentTotal = 0;
    if (debtPaymentSheet) {
      const debtData = debtPaymentSheet.getDataRange().getValues();
      for (let i = 1; i < debtData.length; i++) {
        const row = debtData[i];
        const date = row[1]; // Column B: Date
        const principal = parseFloat(row[3]) || 0; // Column D: Principal
        const interest = parseFloat(row[4]) || 0; // Column E: Interest
        
        if (date >= filters.startDate && date <= filters.endDate) {
          debtPaymentTotal += (principal + interest);
        }
      }
    }
    
    // Build result array
    const result = [];
    const categories = APP_CONFIG.CATEGORIES.EXPENSE.filter(cat => cat !== 'Trả nợ' && cat !== 'Cho vay');
    let totalExpense = 0;
    let totalBudget = 0;
    
    for (const category of categories) {
      const spent = categoryTotals[category] || 0;
      const budget = budgetData[category] || 0;
      const remaining = budget - spent;
      const percentage = budget > 0 ? (spent / budget) : 0;
      
      result.push([category, spent, budget, remaining, percentage]);
      totalExpense += spent;
      totalBudget += budget;
    }
    
    // Add debt payment row
    result.push(['Trả nợ (Gốc + Lãi)', debtPaymentTotal, 0, -debtPaymentTotal, 0]);
    totalExpense += debtPaymentTotal;
    
    // Add total row
    const totalRemaining = totalBudget - totalExpense;
    const totalPercentage = totalBudget > 0 ? (totalExpense / totalBudget) : 0;
    result.push(['TỔNG CHI PHÍ', totalExpense, totalBudget, totalRemaining, totalPercentage]);
    
    return result;
    
  } catch (error) {
    return [['Lỗi: ' + error.message]];
  }
}

/**
 * Get list of outstanding debts
 * @return {Array} 2D array: [["Debt Name", "Remaining", "Percentage"], ...]
 * @customfunction
 */
function hffsDebt() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const debtSheet = ss.getSheetByName(APP_CONFIG.SHEETS.DEBT_MANAGEMENT);
    
    if (!debtSheet) return [['Lỗi: Sheet QUẢN LÝ NỢ không tồn tại']];
    
    const data = debtSheet.getDataRange().getValues();
    const result = [];
    let totalDebt = 0;
    
    // Collect active debts
    const activeDebts = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const name = row[1]; // Column B: Debt name
      const remaining = parseFloat(row[10]) || 0; // Column K: Remaining
      const status = row[11]; // Column L: Status
      
      if (status !== 'Đã thanh toán' && remaining > 0) {
        activeDebts.push({ name, remaining });
        totalDebt += remaining;
      }
    }
    
    // Build result
    for (const debt of activeDebts) {
      const percentage = totalDebt > 0 ? (debt.remaining / totalDebt) : 0;
      result.push([debt.name, debt.remaining, percentage]);
    }
    
    // Add total row
    result.push(['TỔNG NỢ', totalDebt, 1]);
    
    return result;
    
  } catch (error) {
    return [['Lỗi: ' + error.message]];
  }
}

/**
 * Get asset summary
 * @return {Array} 2D array: [["Asset", "Cost", "P&L", "Current Value", "Percentage"], ...]
 * @customfunction
 */
function hffsAssets() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    const result = [];
    
    // Helper to sum column
    const sumColumn = (sheetName, colIndex, filterCol = null, filterValue = null) => {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) return 0;
      
      const data = sheet.getDataRange().getValues();
      let total = 0;
      
      for (let i = 1; i < data.length; i++) {
        if (filterCol === null || data[i][filterCol] === filterValue) {
          total += parseFloat(data[i][colIndex]) || 0;
        }
      }
      return total;
    };
    
    // 1. Net Cash = Income - Expense
    // Note: Debt payments are already in Expense sheet, so don't subtract debt remaining
    // Debt remaining is a liability, not a cash outflow
    const totalIncome = sumColumn(APP_CONFIG.SHEETS.INCOME, 2); // Col C
    const totalExpense = sumColumn(APP_CONFIG.SHEETS.EXPENSE, 2); // Col C
    const netCash = totalIncome - totalExpense;
    
    // 2. Stock
    const stockCost = sumColumn(APP_CONFIG.SHEETS.STOCK, 7); // Col H: Total Cost
    const stockValue = sumColumn(APP_CONFIG.SHEETS.STOCK, 12); // Col M: Market Value
    const stockPL = stockValue - stockCost;
    
    // 3. Gold
    const goldCost = sumColumn(APP_CONFIG.SHEETS.GOLD, 8); // Col I: Total Cost
    const goldValue = sumColumn(APP_CONFIG.SHEETS.GOLD, 10); // Col K: Market Value
    const goldPL = goldValue - goldCost;
    
    // 4. Crypto
    const cryptoCost = sumColumn(APP_CONFIG.SHEETS.CRYPTO, 8); // Col I: Total Cost (VND)
    const cryptoValue = sumColumn(APP_CONFIG.SHEETS.CRYPTO, 12); // Col M: Market Value (VND)
    const cryptoPL = cryptoValue - cryptoCost;
    
    // 5. Other Investment
    const otherInvCost = sumColumn(APP_CONFIG.SHEETS.OTHER_INVESTMENT, 3); // Col D: Amount
    const otherInvReturn = sumColumn(APP_CONFIG.SHEETS.OTHER_INVESTMENT, 6); // Col G: Expected Return
    const otherInvPL = otherInvReturn - otherInvCost;
    
    // 6. Lending
    const lendingCost = sumColumn(APP_CONFIG.SHEETS.LENDING, 3); // Col D: Principal
    const lendingRemaining = sumColumn(APP_CONFIG.SHEETS.LENDING, 10); // Col K: Remaining
    const lendingPL = lendingRemaining - lendingCost; // Simplified, should include interest
    
    // Build result
    result.push(['Tiền mặt (Ròng)', 0, 0, netCash]);
    result.push(['Chứng khoán', stockCost, stockPL, stockValue]);
    result.push(['Vàng', goldCost, goldPL, goldValue]);
    result.push(['Crypto', cryptoCost, cryptoPL, cryptoValue]);
    result.push(['Đầu tư khác', otherInvCost, otherInvPL, otherInvReturn]);
    result.push(['Cho vay', lendingCost, lendingPL, lendingRemaining]);
    
    // Calculate total
    const totalAssets = netCash + stockValue + goldValue + cryptoValue + otherInvReturn + lendingRemaining;
    
    // Add percentages
    const finalResult = [];
    for (const row of result) {
      const percentage = totalAssets > 0 ? (row[3] / totalAssets) : 0;
      finalResult.push([...row, percentage]);
    }
    
    // Add total
    const totalCost = stockCost + goldCost + cryptoCost + otherInvCost + lendingCost;
    const totalPL = stockPL + goldPL + cryptoPL + otherInvPL + lendingPL;
    finalResult.push(['TỔNG TÀI SẢN', totalCost, totalPL, totalAssets, 1]);
    
    return finalResult;
    
  } catch (error) {
    return [['Lỗi: ' + error.message]];
  }
}

/**
 * Get 12-month statistics
 * @param {number} year - Year (default: current year)
 * @return {Array} 2D array with monthly statistics
 * @customfunction
 */
function hffsYearly(year) {
  try {
    const currentYear = year || new Date().getFullYear();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    const result = [];
    
    const incomeSheet = ss.getSheetByName(APP_CONFIG.SHEETS.INCOME);
    const expenseSheet = ss.getSheetByName(APP_CONFIG.SHEETS.EXPENSE);
    const debtPaymentSheet = ss.getSheetByName(APP_CONFIG.SHEETS.DEBT_PAYMENT);
    const stockSheet = ss.getSheetByName(APP_CONFIG.SHEETS.STOCK);
    const goldSheet = ss.getSheetByName(APP_CONFIG.SHEETS.GOLD);
    const cryptoSheet = ss.getSheetByName(APP_CONFIG.SHEETS.CRYPTO);
    const otherInvSheet = ss.getSheetByName(APP_CONFIG.SHEETS.OTHER_INVESTMENT);
    const lendingSheet = ss.getSheetByName(APP_CONFIG.SHEETS.LENDING);
    
    // Helper function to sum by month
    const sumByMonth = (sheet, month, amountCol, dateCol = 1, filterCol = null, excludeValue = null) => {
      if (!sheet) return 0;
      const data = sheet.getDataRange().getValues();
      let total = 0;
      
      for (let i = 1; i < data.length; i++) {
        const date = data[i][dateCol];
        const amount = parseFloat(data[i][amountCol]) || 0;
        
        if (filterCol !== null && data[i][filterCol] === excludeValue) continue;
        
        if (date instanceof Date && 
            date.getFullYear() === currentYear && 
            date.getMonth() === month - 1) {
          total += amount;
        }
      }
      return total;
    };
    
    // Calculate for each month
    for (let m = 1; m <= 12; m++) {
      const monthName = `Tháng ${m}`;
      
      // Income (Column C)
      const income = sumByMonth(incomeSheet, m, 2);
      
      // Expense (Column C) - Exclude "Trả nợ" category
      const expense = sumByMonth(expenseSheet, m, 2, 1, 3, 'Trả nợ');
      
      // Debt payments (Principal + Interest)
      const debtPrincipal = sumByMonth(debtPaymentSheet, m, 3);
      const debtInterest = sumByMonth(debtPaymentSheet, m, 4);
      const totalDebt = debtPrincipal + debtInterest;
      
      // Investment amounts (using total cost for investments made in month)
      const stockInv = sumByMonth(stockSheet, m, 7); // Total Cost
      const goldInv = sumByMonth(goldSheet, m, 8); // Total Cost
      const cryptoInv = sumByMonth(cryptoSheet, m, 8); // Total Cost (VND)
      
      // Other Investment + Lending
      const otherInv = sumByMonth(otherInvSheet, m, 3); // Amount
      const lending = sumByMonth(lendingSheet, m, 3, 6); // Principal, date in col G
      const totalOtherInv = otherInv + lending;
      
      // Cash flow = Income - Expense - Debt Payments
      const cashflow = income - expense - totalDebt;
      
      result.push([
        monthName, 
        income, 
        expense, 
        totalDebt, 
        stockInv, 
        goldInv, 
        cryptoInv, 
        totalOtherInv, 
        cashflow
      ]);
    }
    
    return result;
    
  } catch (error) {
    return [['Lỗi: ' + error.message]];
  }
}
