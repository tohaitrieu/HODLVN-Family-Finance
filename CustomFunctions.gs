/**
 * ===============================================
 * CUSTOM FUNCTIONS v2.0 - UNIFIED TYPE SYSTEM
 * ===============================================
 * Refactored to use shared LOAN_TYPES configuration
 * from Main.gs for both debt and lending calculations
 */

/**
 * Map legacy type names to new type IDs (Backward compatibility)
 * @param {string} typeName - Old type name or new type ID
 * @return {string} Type ID
 */
function mapLegacyTypeToId(typeName) {
  if (!typeName) return 'OTHER';
  
  // Legacy mapping
  const legacyMapping = {
    // Legacy lending types
    'Tất toán gốc - lãi cuối kỳ': 'BULLET',
    'Trả lãi hàng tháng, gốc cuối kỳ': 'INTEREST_ONLY',
    'Trả góp gốc - lãi hàng tháng': 'EQUAL_PRINCIPAL',
    
    // Legacy debt types
    'Nợ ngân hàng': 'INTEREST_ONLY',
    'Vay trả góp': 'EQUAL_PRINCIPAL',
    'Trả góp qua thẻ (Phí ban đầu)': 'EQUAL_PRINCIPAL_UPFRONT_FEE',
    'Trả góp qua thẻ (Lãi giảm dần)': 'EQUAL_PRINCIPAL',
    'Trả góp qua thẻ (Miễn lãi)': 'INTEREST_FREE',
    'Khác': 'OTHER'
  };
  
  if (legacyMapping[typeName]) {
    return legacyMapping[typeName];
  }
  
  // Try partial matching
  const typeNameLower = typeName.toLowerCase();
  
  if (typeNameLower.includes('vay trả góp') || typeNameLower.includes('cho vay trả góp')) {
    if (typeNameLower.includes('miễn lãi')) {
      return 'INTEREST_FREE';
    } else if (typeNameLower.includes('phí ban đầu')) {
      return 'EQUAL_PRINCIPAL_UPFRONT_FEE';
    } else {
      return 'EQUAL_PRINCIPAL';
    }
  }
  
  if (typeNameLower.includes('nợ ngân hàng') || typeNameLower.includes('trả lãi hàng tháng')) {
    return 'INTEREST_ONLY';
  }
  
  if (typeNameLower.includes('tất toán') || typeNameLower.includes('bullet')) {
    return 'BULLET';
  }
  
  if (typeNameLower.includes('miễn lãi')) {
    return 'INTEREST_FREE';
  }
  
  // Default: assume it's already an ID or fallback to OTHER
  return typeName.toUpperCase() === 'BULLET' || 
         typeName.toUpperCase() === 'INTEREST_ONLY' ||
         typeName.toUpperCase() === 'EQUAL_PRINCIPAL' ||
         typeName.toUpperCase() === 'EQUAL_PRINCIPAL_UPFRONT_FEE' ||
         typeName.toUpperCase() === 'INTEREST_FREE' ? typeName.toUpperCase() : 'OTHER';
}


/**
 * Tính toán lịch trả nợ sắp tới (Khoản phải trả)
 * @param {Array[]} debtData Dữ liệu từ sheet QUẢN LÝ NỢ (A2:L)
 * @return {Array[]} Bảng lịch sự kiện
 * @customfunction
 */
function AccPayable(debtData) {
  Logger.log('=== AccPayable CALLED ===');
  Logger.log(`Input data length: ${Array.isArray(debtData) ? debtData.length : 'NOT ARRAY'}`);
  
  if (!Array.isArray(debtData) || debtData.length === 0) {
    Logger.log('⚠️ No data or invalid input');
    return [['Không có dữ liệu', '', '', '', '', '']];
  }
  
  const result = _calculateEvents(debtData, true);
  Logger.log(`AccPayable result rows: ${result.length}`);
  Logger.log(`First row: ${JSON.stringify(result[0])}`);
  
  return result;
}

/**
 * Tính toán lịch thu nợ sắp tới (Khoản phải thu)
 * @param {Array[]} lendingData Dữ liệu từ sheet CHO VAY (A2:L)
 * @return {Array[]} Bảng lịch sự kiện
 * @customfunction
 */
function AccReceivable(lendingData) {
  Logger.log('=== AccReceivable CALLED ===');
  Logger.log(`Input data length: ${Array.isArray(lendingData) ? lendingData.length : 'NOT ARRAY'}`);
  
  if (!Array.isArray(lendingData) || lendingData.length === 0) {
    Logger.log('⚠️ No data or invalid input');
    return [['Không có dữ liệu', '', '', '', '', '']];
  }
  
  const result = _calculateEvents(lendingData, false);
  Logger.log(`AccReceivable result rows: ${result.length}`);
  
  return result;
}

/**
 * Helper function to calculate events
 * ĐÃ SỬA: Đưa cột Lãi suất xuống cuối để không làm lệch cột Gốc/Lãi
 */
function _calculateEvents(data, isDebt) {
  const events = [];
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  
  data.forEach(row => {
    if (!row[1]) return;
    
    const name = row[1];
    let type = row[2];
    
    type = mapLegacyTypeToId(type);
    
    const initialPrincipal = parseCurrency(row[3]);
    let rate = parseFloat(row[4]) || 0;
    
    // Chuẩn hóa lãi suất
    if (rate > 1) {
      rate = rate / 100;
    }
    if (isNaN(rate) || rate < 0) {
      rate = 0;
    }
    
    let term = row[5];
    if (term instanceof Date) {
      term = Math.round((term - new Date(1899, 11, 30)) / 86400000);
    } else {
      term = parseInt(term) || 1;
    }
    
    const rawStartDate = row[6];
    const rawMaturityDate = row[7];
    const remaining = parseCurrency(row[10]);
    const status = row[11];
    
    const startDate = parseDate(rawStartDate);
    const maturityDate = parseDate(rawMaturityDate);
    
    let isActive = false;
    if (isDebt) {
      isActive = (status === 'Chưa trả' || status === 'Đang trả');
    } else {
      isActive = (status === 'Đang vay');
    }
    
    if (isActive && remaining > 0 && startDate) {
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
        nextEvent.rate = rate; // Lưu lãi suất
        events.push(nextEvent);
      }
    }
  });

  if (events.length === 0) {
    // Trả về 7 cột trống
    return [['Không có sự kiện sắp tới', '', '', '', '', '', '']];
  }
  
  events.sort((a, b) => a.date - b.date);
  const limitedEvents = events.slice(0, 10);
  
  // --- SỬA ĐỔI QUAN TRỌNG: Đưa rate xuống cuối ---
  const result = limitedEvents.map(evt => [
    evt.date,              // Cột 1: Ngày
    evt.action,            // Cột 2: Hành động
    evt.name,              // Cột 3: Tên
    evt.remaining,         // Cột 4: Còn lại
    evt.principalPayment,  // Cột 5: Gốc (Giữ nguyên vị trí cũ)
    evt.interestPayment,   // Cột 6: Lãi (Giữ nguyên vị trí cũ)
    evt.rate               // Cột 7: Lãi suất (Mới thêm vào cuối)
  ]);
  
  let totalPrincipal = 0;
  let totalInterest = 0;
  
  limitedEvents.forEach(evt => {
    totalPrincipal += evt.principalPayment;
    totalInterest += evt.interestPayment;
  });
  
  // Dòng tổng cộng: Tổng Gốc (cột 5), Tổng Lãi (cột 6), Trống (cột 7)
  result.push(['TỔNG CỘNG', '', '', '', totalPrincipal, totalInterest, '']);
  
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
 * Lãi = term * (remaining * rate) / 12
 */
function calculateBulletPayment(params) {
  const { name, isDebt, remaining, rate, term, maturityDate, today } = params;
  
  if (!maturityDate || maturityDate < today) return null;
  
  // Validate values to prevent #NUM!
  const validRate = (isNaN(rate) || rate < 0) ? 0 : rate;
  const validRemaining = (isNaN(remaining) || remaining < 0) ? 0 : remaining;
  const validTerm = (isNaN(term) || term <= 0) ? 1 : term;
  
  // Lãi = Số kỳ * (Gốc * lãi_năm) / 12
  const interest = validTerm * (validRemaining * validRate) / 12;
  
  return {
    date: maturityDate,
    action: isDebt ? 'Phải trả' : 'Phải thu',
    name: `${name} (${isDebt ? 'Tất toán' : 'Đáo hạn'})`,
    remaining: validRemaining,
    principalPayment: validRemaining,
    interestPayment: isNaN(interest) ? 0 : interest
  };
}

/**
 * INTEREST_ONLY: Monthly interest, principal at maturity
 * Lãi kỳ này = Số ngày trong tháng cột Ngày * (Gốc còn lại * lãi)/365
 */
function calculateInterestOnlyPayment(params) {
  const { name, isDebt, remaining, rate, term, startDate, today } = params;
  
  // Validate values to prevent #NUM!
  const validRate = (isNaN(rate) || rate < 0) ? 0 : rate;
  const validRemaining = (isNaN(remaining) || remaining < 0) ? 0 : remaining;
  
  // Loop through each month to find next payment
  for (let i = 1; i <= term; i++) {
    // Create new date object for each iteration
    let payDate = new Date(startDate.getFullYear(), startDate.getMonth() + i, startDate.getDate());
    
    // Handle edge case: if original date is 31 but new month has only 30 days
    if (payDate.getDate() !== startDate.getDate()) {
      payDate = new Date(startDate.getFullYear(), startDate.getMonth() + i + 1, 0);
    }
    
    // Check for invalid date
    if (isNaN(payDate.getTime())) continue;
    
    if (payDate >= today) {
      // Calculate days in the month of the payment date
      const daysInMonth = getDaysInMonth(payDate);
      
      // Monthly interest = (Days in Month × Remaining Principal × Rate) / 365
      let monthlyInterest = (daysInMonth * validRemaining * validRate) / 365;
      
      // Validate result
      if (isNaN(monthlyInterest) || monthlyInterest < 0) {
        monthlyInterest = 0;
      }
      
      // Principal payment only on last period
      // Gốc kỳ này = 0. Nếu là kỳ cuối thì = tổng gốc cho vay (remaining)
      let principalPay = (i === term) ? validRemaining : 0;
      
      return {
        date: payDate,
        action: isDebt ? 'Phải trả' : 'Phải thu',
        name: `${name} (Kỳ ${i}/${term})`,
        remaining: validRemaining,
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

  // Validate values
  const validRate = (isNaN(rate) || rate < 0) ? 0 : rate;
  const validInitialPrincipal = (isNaN(initialPrincipal) || initialPrincipal < 0) ? 0 : initialPrincipal;
  const validRemaining = (isNaN(remaining) || remaining < 0) ? 0 : remaining;

  // Monthly principal payment (constant)
  const monthlyPrincipal = term ? validInitialPrincipal / term : 0;

  // Use current remaining from column K for interest calculation
  let currentRemaining = validRemaining;

  for (let i = 1; i <= term; i++) {
    // Create new date object for each iteration to avoid mutation issues
    let payDate = new Date(startDate.getFullYear(), startDate.getMonth() + i, startDate.getDate());
    
    // Handle edge case: if original date is 31 but new month has only 30 days
    if (payDate.getDate() !== startDate.getDate()) {
      // Set to last day of the month
      payDate = new Date(startDate.getFullYear(), startDate.getMonth() + i + 1, 0);
    }

    if (isNaN(payDate.getTime())) continue;

    if (payDate >= today) {
      const daysInMonth = getDaysInMonth(payDate);
      // Lãi = (Số ngày trong tháng × Gốc còn lại × Lãi suất) / 365
      let monthlyInterest = (daysInMonth * currentRemaining * validRate) / 365;
      if (isNaN(monthlyInterest) || monthlyInterest < 0) monthlyInterest = 0;

      return {
        date: payDate,
        action: isDebt ? 'Phải trả' : 'Phải thu',
        name: `${name} (Kỳ ${i}/${term})`,
        remaining: currentRemaining,
        principalPayment: isNaN(monthlyPrincipal) ? 0 : monthlyPrincipal,
        interestPayment: monthlyInterest
      };
    }
    // Decrease remaining after this month
    currentRemaining = Math.max(currentRemaining - monthlyPrincipal, 0);
  }
  return null;
}

/**
 * INTEREST_FREE: Equal principal installments, zero interest
 */
function calculateInterestFreePayment(params) {
  const { name, isDebt, initialPrincipal, remaining, term, startDate, today } = params;
  
  // Validate values
  const validRemaining = (isNaN(remaining) || remaining < 0) ? 0 : remaining;
  const validInitialPrincipal = (isNaN(initialPrincipal) || initialPrincipal < 0) ? 0 : initialPrincipal;
  
  const monthlyPrincipal = validInitialPrincipal / term;
  
  for (let i = 1; i <= term; i++) {
    let payDate = new Date(startDate);
    payDate.setMonth(payDate.getMonth() + i);
    
    if (payDate >= today) {
      return {
        date: payDate,
        action: isDebt ? 'Phải trả' : 'Phải thu',
        name: `${name} (Kỳ ${i}/${term})`,
        remaining: validRemaining,
        principalPayment: isNaN(monthlyPrincipal) ? 0 : monthlyPrincipal,
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
 * Uses hardcoded map for accuracy
 */
function getDaysInMonth(date) {
  Logger.log(`[getDaysInMonth] Input: ${date}`);
  
  if (!date || isNaN(date.getTime())) {
    Logger.log(`[getDaysInMonth] Invalid date, returning 30`);
    return 30; // Fallback
  }
  
  const month = date.getMonth(); // 0-11
  const year = date.getFullYear();
  
  Logger.log(`[getDaysInMonth] Month: ${month}, Year: ${year}`);
  
  // Map số ngày cho từng tháng (index 0-11)
  const daysInMonthMap = [
    31, // Jan (0)
    isLeapYear(year) ? 29 : 28, // Feb (1) - Check năm nhuận
    31, // Mar (2)
    30, // Apr (3)
    31, // May (4)
    30, // Jun (5)
    31, // Jul (6)
    31, // Aug (7)
    30, // Sep (8)
    31, // Oct (9)
    30, // Nov (10)
    31  // Dec (11)
  ];
  
  const result = daysInMonthMap[month];
  Logger.log(`[getDaysInMonth] Result: ${result} days`);
  
  return result;
}

/**
 * Check if a year is a leap year
 */
function isLeapYear(year) {
  return (year % 4 === 0 && year % 100 !== 0) || (year % 400 === 0);
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