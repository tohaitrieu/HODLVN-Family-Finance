/**
 * ===============================================
 * CUSTOM FUNCTIONS v3.6.4 - CONFLICT FIX
 * ===============================================
 * - CRITICAL FIX: Đổi tên hàm helper thành _getDaysInMonthSafe để tránh xung đột
 * với hàm getDaysInMonth(month, year) đang tồn tại trong file Utils.gs.
 * - Đây là nguyên nhân gốc rễ gây ra lỗi NaN.
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

function AccPayable(debtData, forceRefresh) {
  // Make function volatile by referencing a refresh trigger
  try {
    const ss = getSpreadsheet();
    const dashboard = ss.getSheetByName('TỔNG QUAN');
    if (dashboard) {
      // Read timestamp cell to force dependency (makes function volatile)
      const refreshTrigger = dashboard.getRange('Z1').getValue();
      // This creates a dependency that forces recalculation when Z1 changes
    }
  } catch (e) {
    // Ignore errors to prevent function failure
  }
  
  if (!Array.isArray(debtData) || debtData.length === 0) return [['Không có dữ liệu']];
  return _calculateEvents(debtData, true);
}

function AccReceivable(lendingData, forceRefresh) {
  // Make function volatile by referencing a refresh trigger
  try {
    const ss = getSpreadsheet();
    const dashboard = ss.getSheetByName('TỔNG QUAN');
    if (dashboard) {
      // Read timestamp cell to force dependency (makes function volatile)
      const refreshTrigger = dashboard.getRange('Z2').getValue();
      // This creates a dependency that forces recalculation when Z2 changes
    }
  } catch (e) {
    // Ignore errors to prevent function failure
  }
  
  if (!Array.isArray(lendingData) || lendingData.length === 0) return [['Không có dữ liệu']];
  return _calculateEvents(lendingData, false);
}

function _calculateEvents(data, isDebt) {
  Logger.log("=== RUNNING v3.6.4 (Conflict Fixed) ==="); 
  
  const events = [];
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  
  data.forEach((row, index) => {
    if (!row[1]) return;
    
    const name = row[1];
    const type = mapLegacyTypeToId(row[2]);
    
    const initialPrincipal = parseCurrency(row[3]);
    let rate = parseFloat(row[4]);
    
    if (isNaN(rate)) rate = 0;
    if (rate > 1) rate = rate / 100;
    if (rate < 0) rate = 0;
    
    let term = row[5];
    if (term instanceof Date) {
      term = Math.round((term - new Date(1899, 11, 30)) / 86400000);
    } else {
      term = parseInt(term);
    }
    if (isNaN(term) || term <= 0) term = 1;
    
    const startDate = parseDate(row[6]);
    const maturityDate = parseDate(row[7]);
    const remaining = parseCurrency(row[10]); 
    
    let isActive = false;
    const status = row[11];
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
        events.push(nextEvent);
      }
    }
  });

  if (events.length === 0) {
    return [['Không có sự kiện sắp tới', '', '', '', '', '']];
  }
  
  events.sort((a, b) => a.date - b.date);
  const limitedEvents = events.slice(0, 10);
  
  const result = limitedEvents.map(evt => [
    evt.date,
    evt.action || '',
    evt.name || '',
    safeNumber(evt.remaining),
    safeNumber(evt.principalPayment),
    safeNumber(evt.interestPayment)
  ]);
  
  let totalPrincipal = 0;
  let totalInterest = 0;
  limitedEvents.forEach(evt => {
    totalPrincipal += safeNumber(evt.principalPayment);
    totalInterest += safeNumber(evt.interestPayment);
  });
  
  result.push(['TỔNG CỘNG', '', '', '', totalPrincipal, totalInterest]);
  
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

/**
 * 1. VAY TRẢ GÓP
 */
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
      // GỌI HÀM MỚI (Đã đổi tên)
      const daysInMonth = _getDaysInMonthSafe(payDate);
      
      let monthlyInterest = (daysInMonth * remaining * rate) / 365;
      if (isNaN(monthlyInterest)) monthlyInterest = 0;
      
      Logger.log(`[FIXED TRẢ GÓP] ${name}: PayDate=${payDate.toDateString()}, Days=${daysInMonth}, Interest=${monthlyInterest}`);

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

/**
 * 2. TRẢ LÃI ĐỊNH KỲ
 */
function calculateInterestOnlyPayment(params) {
  const { name, isDebt, remaining, rate, term, startDate, today } = params;
  
  for (let i = 1; i <= term; i++) {
    let payDate = new Date(startDate.getFullYear(), startDate.getMonth() + i, startDate.getDate());
    if (payDate.getDate() !== startDate.getDate()) {
      payDate = new Date(startDate.getFullYear(), startDate.getMonth() + i + 1, 0);
    }
    
    if (isNaN(payDate.getTime())) continue;

    if (payDate >= today) {
      // GỌI HÀM MỚI (Đã đổi tên)
      const daysInMonth = _getDaysInMonthSafe(payDate);
      
      let monthlyInterest = (daysInMonth * remaining * rate) / 365;
      if (isNaN(monthlyInterest)) monthlyInterest = 0;
      
      let principalPay = (i === term) ? remaining : 0;
      
      Logger.log(`[FIXED TRẢ LÃI] ${name}: PayDate=${payDate.toDateString()}, Days=${daysInMonth}, Interest=${monthlyInterest}`);

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
 * 3. TẤT TOÁN CUỐI KỲ
 */
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

// ==================== HELPER FUNCTIONS ====================

/**
 * ĐÃ ĐỔI TÊN HÀM: _getDaysInMonthSafe
 * Lý do: Hàm getDaysInMonth(month, year) đã tồn tại trong Utils.gs.
 * Nếu dùng tên cũ, nó sẽ gọi nhầm hàm bên Utils.gs (yêu cầu 2 tham số), dẫn đến lỗi NaN.
 */
function _getDaysInMonthSafe(date) {
  try {
    if (!date) return 30;
    var d = new Date(date);
    if (isNaN(d.getTime())) return 30;
    
    // new Date(year, month + 1, 0) trả về ngày cuối tháng
    return new Date(d.getFullYear(), d.getMonth() + 1, 0).getDate();
  } catch (e) {
    Logger.log("_getDaysInMonthSafe Error: " + e.toString());
    return 30;
  }
}

function parseCurrency(val) {
  if (typeof val === 'number') return isNaN(val) ? 0 : val;
  if (typeof val === 'string') {
    if (!val.trim()) return 0;
    const clean = val.replace(/\D/g, ''); 
    if (!clean) return 0;
    const num = parseFloat(clean);
    return isNaN(num) ? 0 : num;
  }
  return 0;
}

function safeNumber(val) {
  if (typeof val === 'number') {
    return isNaN(val) || !isFinite(val) ? 0 : val;
  }
  return 0;
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