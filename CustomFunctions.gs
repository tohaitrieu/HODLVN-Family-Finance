/**
 * ===============================================
 * CUSTOM FUNCTIONS v3.5.5 - REFINED EVENT LOGIC
 * ===============================================
 * Updated logic for Installment & Interest-Only loans
 * based on specific user requirements.
 */

/**
 * Map legacy type names to new type IDs (Backward compatibility)
 */
function mapLegacyTypeToId(typeName) {
  if (!typeName) return 'OTHER';
  
  // 1. Map chính xác từ cấu hình cũ
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
  
  // 2. Map theo từ khóa (Partial matching)
  const lower = typeName.toLowerCase();
  if (lower.includes('vay trả góp') || lower.includes('cho vay trả góp')) {
    if (lower.includes('miễn lãi')) return 'INTEREST_FREE';
    if (lower.includes('phí ban đầu')) return 'EQUAL_PRINCIPAL_UPFRONT_FEE';
    return 'EQUAL_PRINCIPAL';
  }
  if (lower.includes('nợ ngân hàng') || lower.includes('trả lãi hàng tháng')) return 'INTEREST_ONLY';
  if (lower.includes('tất toán') || lower.includes('bullet')) return 'BULLET';
  if (lower.includes('miễn lãi')) return 'INTEREST_FREE';
  
  // 3. Fallback: Giữ nguyên nếu đã là ID đúng
  const validIds = ['BULLET', 'INTEREST_ONLY', 'EQUAL_PRINCIPAL', 'EQUAL_PRINCIPAL_UPFRONT_FEE', 'INTEREST_FREE'];
  return validIds.includes(typeName.toUpperCase()) ? typeName.toUpperCase() : 'OTHER';
}

/**
 * Tính toán lịch trả nợ sắp tới (Khoản phải trả)
 * @customfunction
 */
function AccPayable(debtData) {
  if (!Array.isArray(debtData) || debtData.length === 0) return [['Không có dữ liệu']];
  return _calculateEvents(debtData, true);
}

/**
 * Tính toán lịch thu nợ sắp tới (Khoản phải thu)
 * @customfunction
 */
function AccReceivable(lendingData) {
  if (!Array.isArray(lendingData) || lendingData.length === 0) return [['Không có dữ liệu']];
  return _calculateEvents(lendingData, false);
}

/**
 * Hàm xử lý chính: Tính toán sự kiện
 */
function _calculateEvents(data, isDebt) {
  const events = [];
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  
  // Duyệt qua từng khoản vay/nợ
  data.forEach(row => {
    // Bỏ qua dòng trống
    if (!row[1]) return;
    
    const name = row[1];
    const type = mapLegacyTypeToId(row[2]);
    const initialPrincipal = parseCurrency(row[3]);
    let rate = parseFloat(row[4]) || 0;
    
    // Chuẩn hóa lãi suất (VD: 12.5 -> 0.125)
    if (rate > 1) rate = rate / 100;
    if (isNaN(rate) || rate < 0) rate = 0;
    
    // Xử lý kỳ hạn (Term)
    let term = row[5];
    if (term instanceof Date) {
      // Fix lỗi Google Sheets tự convert số thành ngày
      term = Math.round((term - new Date(1899, 11, 30)) / 86400000);
    } else {
      term = parseInt(term) || 1;
    }
    
    const startDate = parseDate(row[6]);
    const maturityDate = parseDate(row[7]);
    const remaining = parseCurrency(row[10]); // Cột K: Gốc còn lại thực tế
    const status = row[11];
    
    // Kiểm tra trạng thái hoạt động
    let isActive = false;
    if (isDebt) {
      isActive = (status === 'Chưa trả' || status === 'Đang trả');
    } else {
      isActive = (status === 'Đang vay');
    }
    
    // Chỉ tính toán nếu đang hoạt động, còn nợ và có ngày bắt đầu
    if (isActive && remaining > 0 && startDate) {
      const nextEvent = calculateNextPayment(type, {
        name,
        isDebt,
        initialPrincipal,
        remaining, // Truyền dư nợ thực tế từ Sheet
        rate,
        term,
        startDate,
        maturityDate,
        today
      });
      
      if (nextEvent) {
        // Gán rate vào object để hiển thị
        nextEvent.rate = rate;
        events.push(nextEvent);
      }
    }
  });

  // Nếu không có sự kiện nào
  if (events.length === 0) {
    return [['Không có sự kiện sắp tới', '', '', '', '', '', '']];
  }
  
  // Sắp xếp theo ngày (gần nhất trước)
  events.sort((a, b) => a.date - b.date);
  
  // Lấy tối đa 10 sự kiện
  const limitedEvents = events.slice(0, 10);
  
  // Chuyển đổi sang mảng 2 chiều để hiển thị
  // Cấu trúc: [Ngày, Hành động, Tên, Còn lại, Gốc, Lãi, Lãi suất]
  const result = limitedEvents.map(evt => [
    evt.date,
    evt.action,
    evt.name,
    evt.remaining,
    evt.principalPayment,
    evt.interestPayment,
    evt.rate // Đưa Lãi suất xuống cuối cùng (Cột 7)
  ]);
  
  // Tính tổng
  let totalPrincipal = 0;
  let totalInterest = 0;
  limitedEvents.forEach(evt => {
    totalPrincipal += evt.principalPayment;
    totalInterest += evt.interestPayment;
  });
  
  // Thêm dòng tổng cộng
  result.push(['TỔNG CỘNG', '', '', '', totalPrincipal, totalInterest, '']);
  
  return result;
}

// ==================== LOGIC TÍNH TOÁN CHI TIẾT ====================

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
    default:
      return calculateBulletPayment(params);
  }
}

/**
 * 1. VAY TRẢ GÓP (EQUAL_PRINCIPAL)
 * - Hiển thị 1 sự kiện gần nhất.
 * - Gốc kỳ này = Tổng gốc / Số kỳ.
 * - Lãi kỳ này = Số ngày trong tháng * (Gốc còn lại thực tế * Lãi suất) / 365.
 */
function calculateEqualPrincipalPayment(params) {
  const { name, isDebt, initialPrincipal, remaining, rate, term, startDate, today } = params;
  
  // Gốc trả đều hàng tháng
  const monthlyPrincipal = term ? initialPrincipal / term : 0;
  
  // Tìm kỳ trả nợ TIẾP THEO (gần nhất >= hôm nay)
  for (let i = 1; i <= term; i++) {
    // Tính ngày trả nợ của kỳ thứ i
    let payDate = new Date(startDate.getFullYear(), startDate.getMonth() + i, startDate.getDate());
    
    // Xử lý lỗi ngày (ví dụ 31/2 -> 28/2 hoặc 29/2)
    if (payDate.getDate() !== startDate.getDate()) {
      payDate = new Date(startDate.getFullYear(), startDate.getMonth() + i + 1, 0);
    }
    
    if (isNaN(payDate.getTime())) continue;

    // Nếu ngày này ở tương lai (hoặc hôm nay) -> Đây là sự kiện tiếp theo
    if (payDate >= today) {
      const daysInMonth = getDaysInMonth(payDate);
      
      // Tính lãi dựa trên dư nợ THỰC TẾ từ file Sheet (cột K)
      // Công thức: Số ngày trong tháng * (Gốc còn lại * lãi) / 365
      let monthlyInterest = (daysInMonth * remaining * rate) / 365;
      if (monthlyInterest < 0) monthlyInterest = 0;

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
 * 2. TRẢ LÃI HÀNG THÁNG, GỐC CUỐI KỲ (INTEREST_ONLY)
 * - Gốc kỳ này = 0 (trừ kỳ cuối = Gốc còn lại).
 * - Lãi kỳ này = Số ngày trong tháng * (Gốc còn lại * lãi) / 365.
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
      const daysInMonth = getDaysInMonth(payDate);
      
      // Tính lãi
      let monthlyInterest = (daysInMonth * remaining * rate) / 365;
      if (monthlyInterest < 0) monthlyInterest = 0;
      
      // Gốc: Nếu là kỳ cuối cùng thì trả toàn bộ dư nợ còn lại, ngược lại là 0
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
 * 3. TẤT TOÁN GỐC LÃI CUỐI KỲ (BULLET)
 * - Chỉ hiển thị 1 sự kiện vào ngày đáo hạn.
 */
function calculateBulletPayment(params) {
  const { name, isDebt, remaining, rate, term, maturityDate, today } = params;
  
  if (!maturityDate || maturityDate < today) return null;
  
  // Lãi = (Gốc * lãi suất * số ngày vay) / 365 (hoặc công thức đơn giản theo kỳ)
  // Ở đây dùng công thức đơn giản: Term * (Gốc * Lãi) / 12 cho khớp ước lượng
  // Hoặc theo yêu cầu:
  const interest = term * (remaining * rate) / 12; 

  return {
    date: maturityDate,
    action: isDebt ? 'Phải trả' : 'Phải thu',
    name: `${name} (${isDebt ? 'Tất toán' : 'Đáo hạn'})`,
    remaining: remaining,
    principalPayment: remaining, // Trả hết gốc
    interestPayment: interest
  };
}

/**
 * 4. MIỄN LÃI (INTEREST_FREE)
 */
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

function parseCurrency(val) {
  if (typeof val === 'number') return val;
  if (typeof val === 'string') {
    return parseFloat(val.replace(/\D/g, ''));
  }
  return 0;
}