/**
 * ===============================================
 * CUSTOM FUNCTIONS v2.0 - UNIFIED TYPE SYSTEM
 * ===============================================
 * Refactored to use shared LOAN_TYPES configuration
 * from Main.gs for both debt and lending calculations
 */

/**
 * Tính toán lịch trả nợ sắp tới (Khoản phải trả)
 * @param {Array[]} debtData Dữ liệu từ sheet QUẢN LÝ NỢ (A2:L)
 * @return {Array[]} Bảng lịch sự kiện
 * @customfunction
 */
function AccPayable(debtData) {
  if (!Array.isArray(debtData) || debtData.length === 0) {
    return [['Không có dữ liệu']];
  }
  
  return _calculateEvents(debtData, true);
}

/**
 * Tính toán lịch thu nợ sắp tới (Khoản phải thu)
 * @param {Array[]} lendingData Dữ liệu từ sheet CHO VAY (A2:L)
 * @return {Array[]} Bảng lịch sự kiện
 * @customfunction
 */
function AccReceivable(lendingData) {
  if (!Array.isArray(lendingData) || lendingData.length === 0) {
    return [['Không có dữ liệu']];
  }
  
  return _calculateEvents(lendingData, false);
}

/**
 * Helper function to calculate events
 */
function _calculateEvents(data, isDebt) {
  const events = [];
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  
  // A(0): STT, B(1): Name, C(2): Type, D(3): Principal, E(4): Rate, F(5): Term, G(6): Date, H(7): Maturity
  // I(8): PaidPrin, J(9): PaidInt, K(10): Remaining, L(11): Status
  
  data.forEach(row => {
    // Skip empty rows
    if (!row[1]) return;
    
    const name = row[1];
    let type = row[2]; // Might be legacy name or new ID
    
    // Convert legacy type name to ID using backward compatibility function
    type = mapLegacyTypeToId(type);
    
    const initialPrincipal = parseCurrency(row[3]);
    const rate = parseFloat(row[4]) || 0;
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
      // Use shared calculation function
      const nextEvent = calculateNextPayment(type, {
        name,
        isDebt,
        initialPrincipal,
        remaining,
        rate,
        term,
        startDate,
        maturityDate,
        today
      });
      
      if (nextEvent) {
        events.push(nextEvent);
      }
    }
  });
  
  if (events.length === 0) {
    return [['Không có sự kiện sắp tới', '', '', '', '', '']];
  }
  
  // Sort by date
  events.sort((a, b) => a.date - b.date);
  
  // Limit to 10
  const limitedEvents = events.slice(0, 10);
  
  // Convert to 2D array
  const result = limitedEvents.map(evt => [
    evt.date,
    evt.action,
    evt.name,
    evt.remaining,
    evt.principalPayment,
    evt.interestPayment
  ]);
  
  // Add Total Row
  let totalPrincipal = 0;
  let totalInterest = 0;
  limitedEvents.forEach(evt => {
    totalPrincipal += evt.principalPayment;
    totalInterest += evt.interestPayment;
  });
  
  result.push(['TỔNG CỘNG', '', '', '', totalPrincipal, totalInterest]);
  
  return result;
}

// ==================== SHARED CALCULATION FUNCTIONS ====================

/**
 * Calculate next payment based on loan type ID
 * @param {string} typeId - Loan type ID (BULLET, INTEREST_ONLY, etc.)
 * @param {Object} params - Payment parameters
 * @return {Object|null} Next payment event or null
 */
function calculateNextPayment(typeId, params) {
  switch(typeId) {
    case 'BULLET':
      return calculateBulletPayment(params);
      
    case 'INTEREST_ONLY':
      return calculateInterestOnlyPayment(params);
      
    case 'EQUAL_PRINCIPAL':
    case 'EQUAL_PRINCIPAL_UPFRONT_FEE':
      return calculateEqualPrincipalPayment(params);
      
    case 'INTEREST_FREE':
      return calculateInterestFreePayment(params);
      
    case 'OTHER':
    default:
      // Fallback to bullet
      return calculateBulletPayment(params);
  }
}

/**
 * BULLET: Principal and interest at maturity
 * Hiển thị duy nhất một sự kiện cuối kỳ
 */
function calculateBulletPayment(params) {
  const { name, isDebt, remaining, rate, startDate, maturityDate, today } = params;
  
  if (!maturityDate || maturityDate < today) return null;
  
  const days = getDaysDiff(startDate, maturityDate);
  const interest = days * (rate * remaining) / 365;
  
  return {
    date: maturityDate,
    action: isDebt ? 'Phải trả' : 'Phải thu',
    name: `${name} (${isDebt ? 'Tất toán' : 'Đáo hạn'})`,
    remaining: remaining,
    principalPayment: remaining,
    interestPayment: interest
  };
}

/**
 * INTEREST_ONLY: Monthly interest, principal at maturity
 * Lãi kỳ này = Số ngày trong tháng cột Ngày * (Gốc còn lại * lãi)/365
 */
function calculateInterestOnlyPayment(params) {
  const { name, isDebt, remaining, rate, term, startDate, today } = params;
  
  // Loop through each month to find next payment
  for (let i = 1; i <= term; i++) {
    let payDate = new Date(startDate);
    payDate.setMonth(payDate.getMonth() + i);
    
    // Check for invalid date
    if (isNaN(payDate.getTime())) continue;
    
    if (payDate >= today) {
      // Calculate days in the month of the payment date
      const daysInMonth = getDaysInMonth(payDate);
      
      // Monthly interest = Days in Month * (Remaining Principal * Rate) / 365
      // Ensure rate is treated correctly (e.g. 0.1 for 10%)
      let monthlyInterest = daysInMonth * (rate * remaining) / 365;
      
      // Principal payment only on last period
      // Gốc kỳ này = 0. Nếu là kỳ cuối thì = tổng gốc cho vay (remaining)
      let principalPay = (i === term) ? remaining : 0;
      
      return {
        date: payDate,
        action: isDebt ? 'Phải trả' : 'Phải thu',
        name: `${name} (Kỳ ${i}/${term})`,
        remaining: remaining,
        principalPayment: principalPay,
        interestPayment: monthlyInterest
      };
    }
  }
  
  return null;
}

/**
 * EQUAL_PRINCIPAL: Equal principal installments, decreasing interest
 * Cột Gốc kỳ này = Tổng gốc/số kỳ phải trả.
 * Lãi kỳ này = Số ngày trong tháng cột Ngày * (Gốc còn lại * lãi)/365
 */
function calculateEqualPrincipalPayment(params) {
  const { name, isDebt, initialPrincipal, remaining, rate, term, startDate, today } = params;
  
  // Gốc kỳ này = Tổng gốc / số kỳ phải trả
  const monthlyPrincipal = initialPrincipal / term;
  
  for (let i = 1; i <= term; i++) {
    let payDate = new Date(startDate);
    payDate.setMonth(payDate.getMonth() + i);
    
    // Check for invalid date
    if (isNaN(payDate.getTime())) continue;
    
    // If payment date >= today, return this event
    if (payDate >= today) {
      // Calculate days in the month of the payment date
      const daysInMonth = getDaysInMonth(payDate);
      
      // Interest based on ACTUAL remaining principal from sheet
      // Lãi kỳ này = Số ngày trong tháng cột Ngày * (Gốc còn lại * lãi)/365
      let monthlyInterest = daysInMonth * (rate * remaining) / 365;
      
      return {
        date: payDate,
        action: isDebt ? 'Phải trả' : 'Phải thu',
        name: `${name} (Kỳ ${i}/${term})`,
        remaining: remaining, // Use actual remaining
        principalPayment: monthlyPrincipal,
        interestPayment: monthlyInterest
      };
    }
  }
  
  return null;
}

/**
 * INTEREST_FREE: Equal principal installments, zero interest
 */
function calculateInterestFreePayment(params) {
  const { name, isDebt, initialPrincipal, remaining, term, startDate, today } = params;
  
  const monthlyPrincipal = initialPrincipal / term;
  
  for (let i = 1; i <= term; i++) {
    let payDate = new Date(startDate);
    payDate.setMonth(payDate.getMonth() + i);
    
    if (payDate >= today) {
      return {
        date: payDate,
        action: isDebt ? 'Phải trả' : 'Phải thu',
        name: `${name} (Kỳ ${i}/${term})`,
        remaining: remaining, // Use actual remaining
        principalPayment: monthlyPrincipal,
        interestPayment: 0
      };
    }
  }
  
  return null;
}

// ==================== HELPER FUNCTIONS ====================

/**
 * Calculate days between two dates
 */
function getDaysDiff(d1, d2) {
  if (!d1 || !d2) return 0;
  const diffTime = Math.abs(d2.getTime() - d1.getTime());
  return Math.ceil(diffTime / (1000 * 60 * 60 * 24)); 
}

/**
 * Get number of days in the month of a specific date
 */
function getDaysInMonth(date) {
  if (!date || isNaN(date.getTime())) return 30; // Fallback
  return new Date(date.getFullYear(), date.getMonth() + 1, 0).getDate();
}

/**
 * Parse date safely
 */
function parseDate(d) {
  if (d instanceof Date) return d;
  if (typeof d === 'string') {
    if (d.includes('/')) {
      const parts = d.split('/');
      if (parts.length === 3) return new Date(parts[2], parts[1] - 1, parts[0]);
    }
    return new Date(d);
  }
  return null;
}

/**
 * Parse currency safely
 */
function parseCurrency(val) {
  if (typeof val === 'number') return val;
  if (typeof val === 'string') {
    return parseFloat(val.replace(/\D/g, ''));
  }
  return 0;
}
