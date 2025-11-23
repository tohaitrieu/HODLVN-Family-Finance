/**
 * ===============================================
 * CUSTOM FUNCTIONS v3.5.6 - STABLE RELEASE
 * ===============================================
 * Fix: "Result was not a number" error
 * Fix: Logic tính lãi và hiển thị sự kiện
 */

/**
 * Map tên loại hình cũ sang ID mới
 */
function mapLegacyTypeToId(typeName) {
  if (!typeName) return 'OTHER';
  
  const legacyMapping = {
    'Tất toán gốc - lãi cuối kỳ': 'BULLET',
    'Trả lãi hàng tháng, gốc cuối kỳ': 'INTEREST_ONLY',
    'Trả góp gốc - lãi hàng tháng': 'EQUAL_PRINCIPAL',
    'Nợ ngân hàng': 'INTEREST_ONLY',
    'Vay trả góp': 'EQUAL_PRINCIPAL',
    'Trả góp qua thẻ (Phí ban đầu)': 'EQUAL_PRINCIPAL_UPFRONT_FEE',
    'Trả góp qua thẻ (Lãi giảm dần)': 'EQUAL_PRINCIPAL',
    'Trả góp qua thẻ (Miễn lãi)': 'INTEREST_FREE',
    'Khác': 'OTHER'
  };
  
  if (legacyMapping[typeName]) return legacyMapping[typeName];
  
  const lower = typeName.toLowerCase();
  if (lower.includes('vay trả góp') || lower.includes('cho vay trả góp')) {
    if (lower.includes('miễn lãi')) return 'INTEREST_FREE';
    if (lower.includes('phí ban đầu')) return 'EQUAL_PRINCIPAL_UPFRONT_FEE';
    return 'EQUAL_PRINCIPAL';
  }
  if (lower.includes('nợ ngân hàng') || lower.includes('trả lãi hàng tháng')) return 'INTEREST_ONLY';
  if (lower.includes('tất toán') || lower.includes('bullet')) return 'BULLET';
  if (lower.includes('miễn lãi')) return 'INTEREST_FREE';
  
  const validIds = ['BULLET', 'INTEREST_ONLY', 'EQUAL_PRINCIPAL', 'EQUAL_PRINCIPAL_UPFRONT_FEE', 'INTEREST_FREE'];
  return validIds.includes(typeName.toUpperCase()) ? typeName.toUpperCase() : 'OTHER';
}

/**
 * Tính toán lịch trả nợ sắp tới
 * @customfunction
 */
function AccPayable(debtData) {
  if (!Array.isArray(debtData) || debtData.length === 0) return [['Không có dữ liệu']];
  return _calculateEvents(debtData, true);
}

/**
 * Tính toán lịch thu nợ sắp tới
 * @customfunction
 */
function AccReceivable(lendingData) {
  if (!Array.isArray(lendingData) || lendingData.length === 0) return [['Không có dữ liệu']];
  return _calculateEvents(lendingData, false);
}

/**
 * Hàm xử lý chính
 */
function _calculateEvents(data, isDebt) {
  const events = [];
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  
  data.forEach(row => {
    if (!row[1]) return;
    
    const name = row[1];
    const type = mapLegacyTypeToId(row[2]);
    
    // Sử dụng hàm parseCurrency an toàn mới
    const initialPrincipal = parseCurrency(row[3]);
    const remaining = parseCurrency(row[10]); 
    
    // Xử lý Rate an toàn
    let rate = parseFloat(row[4]);
    if (isNaN(rate)) rate = 0;
    if (rate > 1) rate = rate / 100;
    if (rate < 0) rate = 0;
    
    // Xử lý Term an toàn
    let term = row[5];
    if (term instanceof Date) {
      term = Math.round((term - new Date(1899, 11, 30)) / 86400000);
    } else {
      term = parseInt(term);
    }
    if (isNaN(term) || term <= 0) term = 1;
    
    const startDate = parseDate(row[6]);
    const maturityDate = parseDate(row[7]);
    const status = row[11];
    
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
        nextEvent.rate = rate;
        events.push(nextEvent);
      }
    }
  });

  if (events.length === 0) {
    return [['Không có sự kiện sắp tới', '', '', '', '', '', '']];
  }
  
  events.sort((a, b) => a.date - b.date);
  const limitedEvents = events.slice(0, 10);
  
  // Map kết quả và đảm bảo không có giá trị NaN
  const result = limitedEvents.map(evt => [
    evt.date,
    evt.action || '',
    evt.name || '',
    safeNumber(evt.remaining),
    safeNumber(evt.principalPayment),
    safeNumber(evt.interestPayment),
    safeNumber(evt.rate)
  ]);
  
  let totalPrincipal = 0;
  let totalInterest = 0;
  limitedEvents.forEach(evt => {
    totalPrincipal += safeNumber(evt.principalPayment);
    totalInterest += safeNumber(evt.interestPayment);
  });
  
  result.push(['TỔNG CỘNG', '', '', '', totalPrincipal, totalInterest, '']);
  
  return result;
}

// ==================== LOGIC TÍNH TOÁN ====================

function calculateNextPayment(typeId, params) {
  switch(typeId) {
    case 'BULLET': return calculateBulletPayment(params);
    case 'INTEREST_ONLY': return calculateInterestOnlyPayment(params);
    case 'EQUAL_PRINCIPAL':
    case 'EQUAL_PRINCIPAL_UPFRONT_FEE': return calculateEqualPrincipalPayment(params);
    case 'INTEREST_FREE': return calculateInterestFreePayment(params);
    default: return calculateBulletPayment(params);
  }
}

function calculateEqualPrincipalPayment(params) {
  const { name, isDebt, initialPrincipal, remaining, rate, term, startDate, today } = params;
  const monthlyPrincipal = term ? initialPrincipal / term : 0;
  
  for (let i = 1; i <= term; i++) {
    let payDate = new Date(startDate.getFullYear(), startDate.getMonth() + i, startDate.getDate());
    if (payDate.getDate() !== startDate.getDate()) {
      payDate = new Date(startDate.getFullYear(), startDate.getMonth() + i + 1, 0);
    }
    
    if (isNaN(payDate.getTime())) continue;

    if (payDate >= today) {
      const daysInMonth = getDaysInMonth(payDate);
      // Sử dụng dư nợ thực tế (remaining) để tính lãi
      let monthlyInterest = (daysInMonth * remaining * rate) / 365;
      
      return {
        date: payDate,
        action: isDebt ? 'Phải trả' : 'Phải thu',
        name: `${name} (Kỳ ${i}/${term})`,
        remaining: remaining,
        principalPayment: monthlyPrincipal,
        interestPayment: monthlyInterest
      };
    }
  }
  return null;
}

function calculateInterestOnlyPayment(params) {
  const { name, isDebt, remaining, rate, term, startDate, today } = params;
  
  for (let i = 1; i <= term; i++) {
    let payDate = new Date(startDate.getFullYear(), startDate.getMonth() + i, startDate.getDate());
    if (payDate.getDate() !== startDate.getDate()) {
      payDate = new Date(startDate.getFullYear(), startDate.getMonth() + i + 1, 0);
    }
    
    if (isNaN(payDate.getTime())) continue;

    if (payDate >= today) {
      const daysInMonth = getDaysInMonth(payDate);
      let monthlyInterest = (daysInMonth * remaining * rate) / 365;
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

function calculateBulletPayment(params) {
  const { name, isDebt, remaining, rate, term, maturityDate, today } = params;
  if (!maturityDate || maturityDate < today) return null;
  
  const interest = term * (remaining * rate) / 12; 

  return {
    date: maturityDate,
    action: isDebt ? 'Phải trả' : 'Phải thu',
    name: `${name} (${isDebt ? 'Tất toán' : 'Đáo hạn'})`,
    remaining: remaining,
    principalPayment: remaining,
    interestPayment: interest
  };
}

function calculateInterestFreePayment(params) {
  const { name, isDebt, initialPrincipal, remaining, term, startDate, today } = params;
  const monthlyPrincipal = term ? initialPrincipal / term : 0;

  for (let i = 1; i <= term; i++) {
    let payDate = new Date(startDate);
    payDate.setMonth(payDate.getMonth() + i);
    
    if (payDate >= today) {
      return {
        date: payDate,
        action: isDebt ? 'Phải trả' : 'Phải thu',
        name: `${name} (Kỳ ${i}/${term})`,
        remaining: remaining,
        principalPayment: monthlyPrincipal,
        interestPayment: 0
      };
    }
  }
  return null;
}

// ==================== HELPER FUNCTIONS (FIXED) ====================

/**
 * Hàm parseCurrency an toàn, luôn trả về số, không trả về NaN
 */
function parseCurrency(val) {
  if (typeof val === 'number') {
    return isNaN(val) ? 0 : val;
  }
  if (typeof val === 'string') {
    if (!val.trim()) return 0;
    // Xóa tất cả ký tự không phải số
    const clean = val.replace(/\D/g, '');
    if (!clean) return 0;
    const num = parseFloat(clean);
    return isNaN(num) ? 0 : num;
  }
  return 0;
}

/**
 * Hàm đảm bảo giá trị là số hợp lệ
 */
function safeNumber(val) {
  if (typeof val === 'number') {
    return isNaN(val) ? 0 : val;
  }
  return 0;
}

function getDaysInMonth(date) {
  if (!date || isNaN(date.getTime())) return 30;
  const month = date.getMonth();
  const year = date.getFullYear();
  const daysInMonthMap = [31, isLeapYear(year) ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
  return daysInMonthMap[month];
}

function isLeapYear(year) {
  return (year % 4 === 0 && year % 100 !== 0) || (year % 400 === 0);
}

function parseDate(d) {
  if (!d) return null;
  if (d instanceof Date) return isNaN(d.getTime()) ? null : d;
  if (typeof d === 'string') {
    if (d.includes('/')) {
      const parts = d.split('/');
      if (parts.length === 3) return new Date(parts[2], parts[1] - 1, parts[0]);
    }
    return new Date(d);
  }
  return null;
}