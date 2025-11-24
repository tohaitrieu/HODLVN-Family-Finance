/**
 * ===============================================
 * DASHBOARDFUNCTIONS.GS - CUSTOM FUNCTIONS FOR DASHBOARD
 * ===============================================
 * 
 * Custom functions to display dashboard data in sheets.
 * These functions can be used directly in cells like: =hffsIncome()
 * 
 * Functions:
 * - hffsIncome(): Income report for current month
 * - hffsExpense(period): Expense report (month/year)
 * - hffsDebt(): List of outstanding debts
 * - hffsAssets(type): Asset summary by type
 * - hffsYearly(year): 12-month statistics
 */

/**
 * Get income report for current month
 * @return {Array} 2D array: [[Label, Amount], ...]
 * @customfunction
 */
function hffsIncome() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const incomeSheet = ss.getSheetByName(APP_CONFIG.SHEETS.INCOME);
    
    if (!incomeSheet) return [['Lỗi: Sheet THU không tồn tại']];
    
    const today = new Date();
    const firstDay = new Date(today.getFullYear(), today.getMonth(), 1);
    const lastDay = new Date(today.getFullYear(), today.getMonth() + 1, 0);
    
    const data = incomeSheet.getDataRange().getValues();
    let totalIncome = 0;
    const categoryTotals = {};
    
    // Skip header row
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const date = row[1]; // Column B: Date
      const amount = parseFloat(row[2]) || 0; // Column C: Amount
      const category = row[3]; // Column D: Category
      
      if (date >= firstDay && date <= lastDay) {
        totalIncome += amount;
        categoryTotals[category] = (categoryTotals[category] || 0) + amount;
      }
    }
    
    // Find largest category
    let largestCategory = '';
    let largestAmount = 0;
    for (const [cat, amt] of Object.entries(categoryTotals)) {
      if (amt > largestAmount) {
        largestCategory = cat;
        largestAmount = amt;
      }
    }
    
    return [
      ['Thu nhập tháng này', totalIncome],
      ['Nguồn thu lớn nhất', largestCategory, largestAmount]
    ];
    
  } catch (error) {
    return [['Lỗi: ' + error.message]];
  }
}

/**
 * Get expense report
 * @param {string} period - "month" or "year"
 * @return {Array} 2D array with expense data by category
 * @customfunction
 */
function hffsExpense(period) {
  try {
    period = period || 'month';
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const expenseSheet = ss.getSheetByName(APP_CONFIG.SHEETS.EXPENSE);
    const budgetSheet = ss.getSheetByName(APP_CONFIG.SHEETS.BUDGET);
    
    if (!expenseSheet) return [['Lỗi: Sheet CHI không tồn tại']];
    
    const today = new Date();
    let startDate, endDate;
    
    if (period.toLowerCase() === 'month') {
      startDate = new Date(today.getFullYear(), today.getMonth(), 1);
      endDate = new Date(today.getFullYear(), today.getMonth() + 1, 0);
    } else {
      startDate = new Date(today.getFullYear(), 0, 1);
      endDate = new Date(today.getFullYear(), 11, 31);
    }
    
    const expenseData = expenseSheet.getDataRange().getValues();
    const categoryTotals = {};
    
    // Calculate spent by category
    for (let i = 1; i < expenseData.length; i++) {
      const row = expenseData[i];
      const date = row[1]; // Column B: Date
      const amount = parseFloat(row[2]) || 0; // Column C: Amount
      const category = row[3]; // Column D: Category
      
      if (date >= startDate && date <= endDate) {
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
    
    // Build result array
    const result = [['Danh mục', 'Đã chi', 'Ngân sách', 'Còn lại', 'Trạng thái (%)']];
    
    for (const [category, spent] of Object.entries(categoryTotals)) {
      const budget = budgetData[category] || 0;
      const remaining = budget - spent;
      const percentage = budget > 0 ? (spent / budget) : 0;
      
      result.push([category, spent, budget, remaining, percentage]);
    }
    
    return result;
    
  } catch (error) {
    return [['Lỗi: ' + error.message]];
  }
}

/**
 * Get list of outstanding debts
 * @return {Array} 2D array with debt details
 * @customfunction
 */
function hffsDebt() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const debtSheet = ss.getSheetByName(APP_CONFIG.SHEETS.DEBT_MANAGEMENT);
    
    if (!debtSheet) return [['Lỗi: Sheet QUẢN LÝ NỢ không tồn tại']];
    
    const data = debtSheet.getDataRange().getValues();
    const result = [['Khoản nợ', 'Nợ gốc', 'Lãi suất', 'Còn nợ', 'Trạng thái']];
    
    // Skip header row
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const name = row[1]; // Column B: Debt name
      const principal = parseFloat(row[3]) || 0; // Column D: Principal
      const rate = parseFloat(row[4]) || 0; // Column E: Rate
      const remaining = parseFloat(row[10]) || 0; // Column K: Remaining
      const status = row[11]; // Column L: Status
      
      // Only include active debts
      if (status !== 'Đã thanh toán' && remaining > 0) {
        result.push([name, principal, rate, remaining, status]);
      }
    }
    
    return result;
    
  } catch (error) {
    return [['Lỗi: ' + error.message]];
  }
}

/**
 * Get asset summary
 * @param {string} type - "stock", "gold", "crypto", or "all"
 * @return {Array} 2D array with asset details
 * @customfunction
 */
function hffsAssets(type) {
  try {
    type = (type || 'all').toLowerCase();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    const result = [['Tài sản', 'Số lượng', 'Giá vốn', 'Giá trị HT', 'Lãi/Lỗ']];
    
    // Stock assets
    if (type === 'stock' || type === 'all') {
      const stockSheet = ss.getSheetByName(APP_CONFIG.SHEETS.STOCK);
      if (stockSheet) {
        const data = stockSheet.getDataRange().getValues();
        for (let i = 1; i < data.length; i++) {
          const row = data[i];
          const ticker = row[3]; // Column D: Ticker
          const quantity = parseFloat(row[4]) || 0; // Column E: Quantity
          const totalCost = parseFloat(row[7]) || 0; // Column H: Total Cost
          const marketValue = parseFloat(row[12]) || 0; // Column M: Market Value
          const profit = parseFloat(row[13]) || 0; // Column N: Profit
          
          if (quantity > 0) {
            result.push([`CK: ${ticker}`, quantity, totalCost, marketValue, profit]);
          }
        }
      }
    }
    
    // Gold assets
    if (type === 'gold' || type === 'all') {
      const goldSheet = ss.getSheetByName(APP_CONFIG.SHEETS.GOLD);
      if (goldSheet) {
        const data = goldSheet.getDataRange().getValues();
        for (let i = 1; i < data.length; i++) {
          const row = data[i];
          const goldType = row[4]; // Column E: Gold Type
          const quantity = parseFloat(row[5]) || 0; // Column F: Quantity
          const unit = row[6]; // Column G: Unit
          const totalCost = parseFloat(row[8]) || 0; // Column I: Total Cost
          const marketValue = parseFloat(row[10]) || 0; // Column K: Market Value
          const profit = parseFloat(row[11]) || 0; // Column L: Profit
          
          if (quantity > 0) {
            result.push([`Vàng: ${goldType}`, `${quantity} ${unit}`, totalCost, marketValue, profit]);
          }
        }
      }
    }
    
    // Crypto assets
    if (type === 'crypto' || type === 'all') {
      const cryptoSheet = ss.getSheetByName(APP_CONFIG.SHEETS.CRYPTO);
      if (cryptoSheet) {
        const data = cryptoSheet.getDataRange().getValues();
        for (let i = 1; i < data.length; i++) {
          const row = data[i];
          const coin = row[3]; // Column D: Coin
          const quantity = parseFloat(row[4]) || 0; // Column E: Quantity
          const totalCost = parseFloat(row[8]) || 0; // Column I: Total Cost (VND)
          const marketValue = parseFloat(row[12]) || 0; // Column M: Market Value (VND)
          const profit = parseFloat(row[13]) || 0; // Column N: Profit
          
          if (quantity > 0) {
            result.push([`Crypto: ${coin}`, quantity, totalCost, marketValue, profit]);
          }
        }
      }
    }
    
    return result;
    
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
    
    const result = [['Tháng', 'Thu', 'Chi', 'Nợ', 'CK', 'Vàng', 'Crypto', 'ĐT khác', 'Dòng tiền']];
    
    const incomeSheet = ss.getSheetByName(APP_CONFIG.SHEETS.INCOME);
    const expenseSheet = ss.getSheetByName(APP_CONFIG.SHEETS.EXPENSE);
    const debtPaymentSheet = ss.getSheetByName(APP_CONFIG.SHEETS.DEBT_PAYMENT);
    const stockSheet = ss.getSheetByName(APP_CONFIG.SHEETS.STOCK);
    const goldSheet = ss.getSheetByName(APP_CONFIG.SHEETS.GOLD);
    const cryptoSheet = ss.getSheetByName(APP_CONFIG.SHEETS.CRYPTO);
    const otherInvSheet = ss.getSheetByName(APP_CONFIG.SHEETS.OTHER_INVESTMENT);
    const lendingSheet = ss.getSheetByName(APP_CONFIG.SHEETS.LENDING);
    
    // Helper function to sum by month
    const sumByMonth = (sheet, month, amountCol, dateCol = 1) => {
      if (!sheet) return 0;
      const data = sheet.getDataRange().getValues();
      let total = 0;
      
      for (let i = 1; i < data.length; i++) {
        const date = data[i][dateCol];
        const amount = parseFloat(data[i][amountCol]) || 0;
        
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
      const monthName = `T${m}`;
      
      // Income (Column C)
      const income = sumByMonth(incomeSheet, m, 2);
      
      // Expense (Column C) - Exclude debt payments
      let expense = 0;
      if (expenseSheet) {
        const data = expenseSheet.getDataRange().getValues();
        for (let i = 1; i < data.length; i++) {
          const date = data[i][1];
          const amount = parseFloat(data[i][2]) || 0;
          const category = data[i][3];
          
          if (date instanceof Date && 
              date.getFullYear() === currentYear && 
              date.getMonth() === m - 1 &&
              category !== 'Trả nợ') {
            expense += amount;
          }
        }
      }
      
      // Debt payments (Principal + Interest)
      const debtPrincipal = sumByMonth(debtPaymentSheet, m, 3);
      const debtInterest = sumByMonth(debtPaymentSheet, m, 4);
      const totalDebt = debtPrincipal + debtInterest;
      
      // Investment P&L (using total cost for investments made in month)
      const stockInv = sumByMonth(stockSheet, m, 7); // Total Cost
      const goldInv = sumByMonth(goldSheet, m, 8); // Total Cost
      const cryptoInv = sumByMonth(cryptoSheet, m, 8); // Total Cost (VND)
      
      // Other Investment + Lending
      const otherInv = sumByMonth(otherInvSheet, m, 3); // Amount
      const lending = sumByMonth(lendingSheet, m, 3, 6); // Principal, date in col G
      const totalOtherInv = otherInv + lending;
      
      // Cash flow = Income - Expense - Debt
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
