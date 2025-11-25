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
 * Calculate date filters from parameters (creates proper dependencies)
 * @param {any} yearFilter - Year filter value from B2
 * @param {any} quarterFilter - Quarter filter value from B3  
 * @param {any} monthFilter - Month filter value from B4
 * @return {Object} {year, quarter, month, startDate, endDate}
 * @private
 */
function _getDateFiltersFromParams(yearFilter, quarterFilter, monthFilter) {
  try {
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

/**
 * Get filter values from dashboard cells B2-B4
 * @return {Object} {year, quarter, month, startDate, endDate}
 * @private
 */
function _getDateFilters() {
  try {
    const sheet = getSpreadsheet().getSheetByName(APP_CONFIG.SHEETS.DASHBOARD);
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
 * NEW: Direct filter parameters for auto-refresh when B2:B4 change
 * @param {any} yearFilter - Year filter value from B2
 * @param {any} quarterFilter - Quarter filter value from B3
 * @param {any} monthFilter - Month filter value from B4
 * @param {any} trigger - Dummy parameter to force recalculation (pass $Z$1)
 * @return {Array} 2D array: [["Category", "Amount", "Percentage"], ...]
 * @customfunction
 */
function hffsIncome(yearFilter, quarterFilter, monthFilter, trigger) {
  try {
    const ss = getSpreadsheet();
    const incomeSheet = ss.getSheetByName(APP_CONFIG.SHEETS.INCOME);
    
    if (!incomeSheet) return [['Lỗi: Sheet THU không tồn tại']];
    
    const filters = _getDateFiltersFromParams(yearFilter, quarterFilter, monthFilter);
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
 * NEW: Direct filter parameters for auto-refresh when B2:B4 change
 * @param {any} yearFilter - Year filter value from B2
 * @param {any} quarterFilter - Quarter filter value from B3
 * @param {any} monthFilter - Month filter value from B4
 * @param {any} trigger - Dummy parameter to force recalculation (pass $Z$1)
 * @return {Array} 2D array: [["Category", "Spent", "Budget", "Remaining", "Status%"], ...]
 * @customfunction
 */
function hffsExpense(yearFilter, quarterFilter, monthFilter, trigger) {
  try {
    const ss = getSpreadsheet();
    const expenseSheet = ss.getSheetByName(APP_CONFIG.SHEETS.EXPENSE);
    const budgetSheet = ss.getSheetByName(APP_CONFIG.SHEETS.BUDGET);
    const debtPaymentSheet = ss.getSheetByName(APP_CONFIG.SHEETS.DEBT_PAYMENT);
    
    if (!expenseSheet) return [['Lỗi: Sheet CHI không tồn tại']];
    
    const filters = _getDateFiltersFromParams(yearFilter, quarterFilter, monthFilter);
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
    
    // Get budget data and apply multiplier based on filter
    const budgetData = {};
    let budgetMultiplier = 1; // Default for month
    
    if (budgetSheet) {
      // Determine multiplier based on filter
      if (filters.month && filters.month !== 'Tất cả') {
        budgetMultiplier = 1; // Month filter
      } else if (filters.quarter && filters.quarter !== 'Tất cả') {
        budgetMultiplier = 3; // Quarter filter (3 months)
      } else {
        budgetMultiplier = 12; // Year filter (12 months)
      }
      
      const budgetValues = budgetSheet.getDataRange().getValues();
      for (let i = 1; i < budgetValues.length; i++) {
        const cat = budgetValues[i][0]; // Column A: Category
        const monthlyBudget = parseFloat(budgetValues[i][2]) || 0; // Column C: Ngân sách (monthly)
        if (cat && monthlyBudget > 0) {
          budgetData[cat] = monthlyBudget * budgetMultiplier;
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
 * @param {any} trigger - Dummy parameter to force recalculation (pass $Z$1)
 * @return {Array} 2D array: [["Debt Name", "Remaining", "Percentage"], ...]
 * @customfunction
 */
function hffsDebt(trigger) {
  try {
    const ss = getSpreadsheet();
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
/**
 * Get asset summary
 * @param {any} trigger - Dummy parameter to force recalculation (pass $Z$1)
 * @return {Array} 2D array: [["Asset", "Cost", "P&L", "Current Value", "Percentage"], ...]
 * @customfunction
 */
function hffsAssets(trigger) {
  try {
    const ss = getSpreadsheet();
    
    const result = [];
    
    // Helper to sum column with filter
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
    const totalIncome = sumColumn(APP_CONFIG.SHEETS.INCOME, 2); // Col C (Index 2)
    const totalExpense = sumColumn(APP_CONFIG.SHEETS.EXPENSE, 2); // Col C (Index 2)
    const netCash = totalIncome - totalExpense;
    
    // 2. Stock (Calculate Net Holdings)
    let stockCost = 0;
    let stockValue = 0;
    
    const stockSheet = ss.getSheetByName(APP_CONFIG.SHEETS.STOCK);
    if (stockSheet) {
      const data = stockSheet.getDataRange().getValues();
      const holdings = {}; // { ticker: { qty: 0, cost: 0, currentPrice: 0 } }
      
      for (let i = 1; i < data.length; i++) {
        const type = data[i][2]; // Col C: Type
        const ticker = data[i][3]; // Col D: Ticker
        const qty = parseFloat(data[i][4]) || 0; // Col E: Quantity
        const totalCost = parseFloat(data[i][7]) || 0; // Col H: Total Cost
        const marketPrice = parseFloat(data[i][11]) || 0; // Col L: Market Price
        
        if (!holdings[ticker]) holdings[ticker] = { qty: 0, cost: 0, currentPrice: 0 };
        
        // Update current price (use the latest available)
        if (marketPrice > 0) holdings[ticker].currentPrice = marketPrice;
        
        if (type === 'Mua') {
          holdings[ticker].qty += qty;
          holdings[ticker].cost += totalCost;
        } else if (type === 'Bán') {
          holdings[ticker].qty -= qty;
          holdings[ticker].cost -= totalCost; // Subtract proceeds from cost basis (Net Investment)
        } else if (type === 'Thưởng') {
          holdings[ticker].qty += qty;
          // Bonus shares have 0 cost
        }
      }
      
      // Calculate totals
      for (const ticker in holdings) {
        const h = holdings[ticker];
        if (h.qty > 0) {
          stockCost += h.cost;
          stockValue += h.qty * h.currentPrice;
        }
      }
    }
    const stockPL = stockValue - stockCost;
    
    // 3. Gold (Calculate Net Holdings)
    let goldCost = 0;
    let goldValue = 0;
    
    const goldSheet = ss.getSheetByName(APP_CONFIG.SHEETS.GOLD);
    if (goldSheet) {
      const data = goldSheet.getDataRange().getValues();
      const holdings = {};
      
      for (let i = 1; i < data.length; i++) {
        const type = data[i][3]; // Col D: Type (Mua/Bán)
        const goldType = data[i][4]; // Col E: Gold Type
        const qty = parseFloat(data[i][5]) || 0; // Col F: Quantity
        const totalCost = parseFloat(data[i][8]) || 0; // Col I: Total Cost
        const marketPrice = parseFloat(data[i][9]) || 0; // Col J: Market Price
        
        if (!holdings[goldType]) holdings[goldType] = { qty: 0, cost: 0, currentPrice: 0 };
        
        if (marketPrice > 0) holdings[goldType].currentPrice = marketPrice;
        
        if (type === 'Mua') {
          holdings[goldType].qty += qty;
          holdings[goldType].cost += totalCost;
        } else if (type === 'Bán') {
          holdings[goldType].qty -= qty;
          holdings[goldType].cost -= totalCost;
        }
      }
      
      for (const type in holdings) {
        const h = holdings[type];
        if (h.qty > 0) {
          goldCost += h.cost;
          goldValue += h.qty * h.currentPrice;
        }
      }
    }
    const goldPL = goldValue - goldCost;
    
    // 4. Crypto (Calculate Net Holdings)
    let cryptoCost = 0;
    let cryptoValue = 0;
    
    const cryptoSheet = ss.getSheetByName(APP_CONFIG.SHEETS.CRYPTO);
    if (cryptoSheet) {
      const data = cryptoSheet.getDataRange().getValues();
      const holdings = {};
      
      for (let i = 1; i < data.length; i++) {
        const type = data[i][2]; // Col C: Type
        const coin = data[i][3]; // Col D: Coin
        const qty = parseFloat(data[i][4]) || 0; // Col E: Quantity
        const totalCost = parseFloat(data[i][8]) || 0; // Col I: Total Cost (VND)
        const marketPriceVND = parseFloat(data[i][11]) || 0; // Col L: Market Price VND
        
        if (!holdings[coin]) holdings[coin] = { qty: 0, cost: 0, currentPrice: 0 };
        
        if (marketPriceVND > 0) holdings[coin].currentPrice = marketPriceVND;
        
        if (type === 'Mua') {
          holdings[coin].qty += qty;
          holdings[coin].cost += totalCost;
        } else if (type === 'Bán') {
          holdings[coin].qty -= qty;
          holdings[coin].cost -= totalCost;
        }
      }
      
      for (const coin in holdings) {
        const h = holdings[coin];
        if (h.qty > 0) {
          cryptoCost += h.cost;
          cryptoValue += h.qty * h.currentPrice;
        }
      }
    }
    const cryptoPL = cryptoValue - cryptoCost;
    
    // 5. Other Investment
    const otherInvCost = sumColumn(APP_CONFIG.SHEETS.OTHER_INVESTMENT, 3); // Col D: Amount
    const otherInvReturn = sumColumn(APP_CONFIG.SHEETS.OTHER_INVESTMENT, 6); // Col G: Expected Return
    const otherInvPL = otherInvReturn - otherInvCost;
    
    // 6. Lending
    // Principal is Col D (Index 3)
    // Remaining is Col K (Index 10)
    const lendingCost = sumColumn(APP_CONFIG.SHEETS.LENDING, 3); 
    const lendingRemaining = sumColumn(APP_CONFIG.SHEETS.LENDING, 10);
    const lendingPL = lendingRemaining - lendingCost; // Simplified
    
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
 * @param {any} trigger - Dummy parameter to force recalculation (pass $Z$1)
 * @return {Array} 2D array with monthly statistics
 * @customfunction
 */
function hffsYearly(year, trigger) {
  try {
    const currentYear = year || new Date().getFullYear();
    const ss = getSpreadsheet();
    
    // Initialize monthly totals (index 0 = month 1, index 11 = month 12)
    const monthlyData = Array.from({length: 12}, () => ({
      income: 0,
      expense: 0,
      debtPrincipal: 0,
      debtInterest: 0,
      stock: 0,
      gold: 0,
      crypto: 0,
      otherInv: 0,
      lending: 0
    }));
    
    // Helper to process sheet data once and accumulate by month
    const processSheet = (sheetName, amountCol, dateCol, targetField, filterCol = null, excludeValue = null) => {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) return;
      
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        const date = data[i][dateCol];
        const amount = parseFloat(data[i][amountCol]) || 0;
        
        // Apply filter if specified
        if (filterCol !== null && data[i][filterCol] === excludeValue) continue;
        
        if (date instanceof Date && date.getFullYear() === currentYear) {
          const monthIndex = date.getMonth();
          monthlyData[monthIndex][targetField] += amount;
        }
      }
    };
    
    // Process each sheet ONCE
    processSheet(APP_CONFIG.SHEETS.INCOME, 2, 1, 'income');
    processSheet(APP_CONFIG.SHEETS.EXPENSE, 2, 1, 'expense', 3, 'Trả nợ'); // Exclude debt payments
    processSheet(APP_CONFIG.SHEETS.DEBT_PAYMENT, 3, 1, 'debtPrincipal');
    processSheet(APP_CONFIG.SHEETS.DEBT_PAYMENT, 4, 1, 'debtInterest');
    processSheet(APP_CONFIG.SHEETS.STOCK, 7, 1, 'stock');
    processSheet(APP_CONFIG.SHEETS.GOLD, 8, 1, 'gold');
    processSheet(APP_CONFIG.SHEETS.CRYPTO, 8, 1, 'crypto');
    processSheet(APP_CONFIG.SHEETS.OTHER_INVESTMENT, 3, 1, 'otherInv');
    processSheet(APP_CONFIG.SHEETS.LENDING, 3, 6, 'lending'); // Date in col G (index 6)
    
    // Build result array
    const result = [];
    for (let m = 0; m < 12; m++) {
      const data = monthlyData[m];
      const totalDebt = data.debtPrincipal + data.debtInterest;
      const totalOtherInv = data.otherInv + data.lending;
      const cashflow = data.income - data.expense - totalDebt;
      
      result.push([
        `Tháng ${m + 1}`,
        data.income,
        data.expense,
        totalDebt,
        data.stock,
        data.gold,
        data.crypto,
        totalOtherInv,
        cashflow
      ]);
    }
    
    return result;
    
  } catch (error) {
    return [['Lỗi: ' + error.message]];
  }
}
