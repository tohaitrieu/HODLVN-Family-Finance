/**
 * ===============================================
 * CUSTOM FUNCTIONS v3.6.0 - STRING PARSING FIX
 * ===============================================
 * - Cải tiến hàm getDaysInMonth với DUAL strategy:
 *   + Strategy 1: Thử dùng date.getMonth() trước
 *   + Strategy 2: Nếu fail → Parse từ Date.toString()
 * - Hỗ trợ parse format: "Mon Dec 01 2025" và "2025-12-01"
 * - Month mapping: Jan=0, Feb=1, ..., Dec=11
 * - FIX: Giải quyết vấn đề getMonth() trả về NaN/undefined
 * - GIỮ NGUYÊN tên hàm để tương thích với code cũ
 */

/**
 * Map legacy type names to new type IDs
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
  
  data.forEach((row, index) => {
    if (!row[1]) return;
    
    const name = row[1];
    const type = mapLegacyTypeToId(row[2]);
    
    // Parse an toàn
    const initialPrincipal = parseCurrency(row[3]);
    let rate = parseFloat(row[4]);
    
    // Chuẩn hóa Rate
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
    
    // Lấy dư nợ từ cột K
    const remaining = parseCurrency(row[10]); 
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
        events.push(nextEvent);
      }
    }
  });

  if (events.length === 0) {
    // Trả về 6 cột trống (đã bỏ cột lãi suất)
    return [['Không có sự kiện sắp tới', '', '', '', '', '']];
  }
  
  events.sort((a, b) => a.date - b.date);
  const limitedEvents = events.slice(0, 10);
  
  // Map kết quả - BỎ CỘT RATE
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
  
  // Dòng tổng cộng - 6 cột
  result.push(['TỔNG CỘNG', '', '', '', totalPrincipal, totalInterest]);
  
  return result;
}

// ==================== LOGIC TÍNH TOÁN (CÓ DEBUG) ====================

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
      // --- DEBUG: Kiểm tra payDate trước khi gọi getDaysInMonth ---
      Logger.log(`[DEBUG TRƯỚC getDaysInMonth - TRẢ GÓP]`);
      Logger.log(`  - payDate value: ${payDate}`);
      Logger.log(`  - typeof payDate: ${typeof payDate}`);
      Logger.log(`  - payDate instanceof Date: ${payDate instanceof Date}`);
      Logger.log(`  - payDate.getTime(): ${payDate.getTime()}`);
      Logger.log(`  - isNaN(payDate.getTime()): ${isNaN(payDate.getTime())}`);
      
      const daysInMonth = getDaysInMonth(payDate);
      
      // --- LOGIC TÍNH LÃI ---
      let monthlyInterest = (daysInMonth * remaining * rate) / 365;
      
      // --- DEBUG LOG ---
      Logger.log(`[DEBUG TRẢ GÓP] Khoản: ${name}`);
      Logger.log(`  - 1. Lãi suất (rate): ${rate} (${rate * 100}%)`);
      Logger.log(`  - 2. Số tiền còn lại (remaining): ${remaining}`);
      Logger.log(`  - 3. Số ngày (daysInMonth): ${daysInMonth} (Tháng ${payDate.getMonth()+1})`);
      Logger.log(`  - 4. Gốc phải trả: ${monthlyPrincipal}`);
      Logger.log(`  - 5. Lãi phải trả (tính): ${monthlyInterest}`);
      Logger.log(`  - Công thức: (${daysInMonth} * ${remaining} * ${rate}) / 365`);
      // -----------------

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
      // --- DEBUG: Kiểm tra payDate trước khi gọi getDaysInMonth ---
      Logger.log(`[DEBUG TRƯỚC getDaysInMonth - TRẢ LÃI]`);
      Logger.log(`  - payDate value: ${payDate}`);
      Logger.log(`  - typeof payDate: ${typeof payDate}`);
      Logger.log(`  - payDate instanceof Date: ${payDate instanceof Date}`);
      Logger.log(`  - payDate.getTime(): ${payDate.getTime()}`);
      Logger.log(`  - isNaN(payDate.getTime()): ${isNaN(payDate.getTime())}`);
      
      const daysInMonth = getDaysInMonth(payDate);
      
      // --- LOGIC TÍNH LÃI ---
      let monthlyInterest = (daysInMonth * remaining * rate) / 365;
      let principalPay = (i === term) ? remaining : 0;
      
      // --- DEBUG LOG ---
      Logger.log(`[DEBUG TRẢ LÃI] Khoản: ${name}`);
      Logger.log(`  - 1. Lãi suất (rate): ${rate} (${rate * 100}%)`);
      Logger.log(`  - 2. Số tiền còn lại (remaining): ${remaining}`);
      Logger.log(`  - 3. Số ngày (daysInMonth): ${daysInMonth}`);
      Logger.log(`  - 4. Gốc phải trả: ${principalPay}`);
      Logger.log(`  - 5. Lãi phải trả (tính): ${monthlyInterest}`);
      // -----------------

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
  
  // Tính lãi ước lượng: Gốc * Lãi * (Số tháng / 12)
  // Hoặc chính xác hơn: (Gốc * Lãi * Số ngày thực tế) / 365
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

function parseCurrency(val) {
  if (typeof val === 'number') return isNaN(val) ? 0 : val;
  if (typeof val === 'string') {
    if (!val.trim()) return 0;
    const clean = val.replace(/\D/g, ''); // Chỉ giữ lại số
    if (!clean) return 0;
    const num = parseFloat(clean);
    return isNaN(num) ? 0 : num;
  }
  return 0;
}

function safeNumber(val) {
  if (typeof val === 'number') return isNaN(val) ? 0 : val;
  return 0;
}

function getDaysInMonth(date) {
  Logger.log(`[getDaysInMonth] Input:`, date);
  
  if (!date) {
    Logger.log(`  → NULL, return 30`);
    return 30;
  }
  
  // Parse từ Date string nếu cần
  let month, year;
  
  if (date instanceof Date && !isNaN(date.getTime())) {
    // Thử dùng getMonth() trước
    month = date.getMonth();
    year = date.getFullYear();
    
    Logger.log(`  Via getMonth(): month=${month}, year=${year}`);
    
    // Nếu month hợp lệ, dùng luôn
    if (typeof month === 'number' && !isNaN(month) && month >= 0 && month <= 11) {
      Logger.log(`  → getMonth() OK!`);
      const daysInMonthMap = [31, isLeapYear(year) ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
      const result = daysInMonthMap[month];
      Logger.log(`  → Result: ${result}`);
      return result;
    }
  }
  
  // FALLBACK: Parse từ string
  Logger.log(`  → Fallback to string parsing`);
  const dateStr = date.toString();
  Logger.log(`  dateStr: ${dateStr}`);
  
  // Map tháng từ tên
  const monthMap = {
    'Jan': 0, 'Feb': 1, 'Mar': 2, 'Apr': 3, 'May': 4, 'Jun': 5,
    'Jul': 6, 'Aug': 7, 'Sep': 8, 'Oct': 9, 'Nov': 10, 'Dec': 11
  };
  
  // Match pattern: "Mon Dec 01 2025" hoặc "2025-12-01"
  let monthStr = null;
  
  // Try format: "Mon Dec 01 2025"
  const match1 = dateStr.match(/\b(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\b/);
  if (match1) {
    monthStr = match1[1];
    Logger.log(`  Found month string: ${monthStr}`);
  }
  
  // Try format: "2025-12-01"
  const match2 = dateStr.match(/\d{4}-(\d{2})-\d{2}/);
  if (match2) {
    month = parseInt(match2[1]) - 1; // Trừ 1 vì tháng đếm từ 0
    Logger.log(`  Found month from ISO format: ${month}`);
  }
  
  if (monthStr && monthMap[monthStr] !== undefined) {
    month = monthMap[monthStr];
    Logger.log(`  Mapped ${monthStr} → ${month}`);
  }
  
  // Extract year
  const yearMatch = dateStr.match(/\b(20\d{2})\b/);
  if (yearMatch) {
    year = parseInt(yearMatch[1]);
    Logger.log(`  Found year: ${year}`);
  } else {
    year = new Date().getFullYear(); // Fallback to current year
  }
  
  if (typeof month === 'number' && !isNaN(month) && month >= 0 && month <= 11) {
    const daysInMonthMap = [31, isLeapYear(year) ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
    const result = daysInMonthMap[month];
    Logger.log(`  → Final result: ${result} (month=${month}, year=${year})`);
    return result;
  }
  
  Logger.log(`  → Could not determine month, return 30`);
  return 30;
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