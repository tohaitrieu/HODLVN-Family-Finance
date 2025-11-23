/**
 * ===============================================
 * CUSTOM FUNCTIONS
 * ===============================================
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
  
  // Helper to calculate days between dates
  const getDaysDiff = (d1, d2) => {
    if (!d1 || !d2) return 0;
    const diffTime = Math.abs(d2.getTime() - d1.getTime());
    return Math.ceil(diffTime / (1000 * 60 * 60 * 24)); 
  };
  
  // Helper to parse date safely
  const parseDate = (d) => {
    if (d instanceof Date) return d;
    if (typeof d === 'string') {
      if (d.includes('/')) {
        const parts = d.split('/');
        if (parts.length === 3) return new Date(parts[2], parts[1] - 1, parts[0]);
      }
      return new Date(d);
    }
    return null;
  };

  // Helper to parse currency safely
  const parseCurrency = (val) => {
    if (typeof val === 'number') return val;
    if (typeof val === 'string') {
      return parseFloat(val.replace(/\D/g, ''));
    }
    return 0;
  };
  
  // A(0): STT, B(1): Name, C(2): Type, D(3): Principal, E(4): Rate, F(5): Term, G(6): Date, H(7): Maturity
  // I(8): PaidPrin, J(9): PaidInt, K(10): Remaining, L(11): Status
  
  data.forEach(row => {
    // Skip empty rows
    if (!row[1]) return;
    
    const name = row[1];
    const type = row[2]; // Lending Type
    const initialPrincipal = parseCurrency(row[3]);
    const rate = parseFloat(row[4]) || 0; // Annual Rate (decimal)
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
      
      // LOGIC THEO TỪNG LOẠI HÌNH
      
      // 1. Tất toán gốc - lãi cuối kỳ
      if (type === 'Tất toán gốc - lãi cuối kỳ') {
        if (maturityDate && maturityDate >= today) {
           const days = getDaysDiff(startDate, maturityDate);
           // Lãi = Ngày * Gốc * Tỷ lệ / 365
           let currentInterest = days * (rate * remaining) / 365;
           
           events.push({
              date: maturityDate,
              action: isDebt ? 'Phải trả' : 'Phải thu',
              name: isDebt ? `${name} (Tất toán)` : `${name} (Đáo hạn)`,
              remaining: remaining,
              principalPayment: remaining,
              interestPayment: currentInterest
           });
        }
      }
      
      // 2. Trả lãi hàng tháng, gốc cuối kỳ
      else if (type === 'Trả lãi hàng tháng, gốc cuối kỳ') {
         // Lặp qua từng tháng để tìm kỳ trả lãi tiếp theo
         for (let i = 1; i <= term; i++) {
            let payDate = new Date(startDate);
            payDate.setMonth(payDate.getMonth() + i);
            
            if (payDate >= today) {
               // Tính số ngày của tháng trước đó
               let prevDate = new Date(startDate);
               prevDate.setMonth(prevDate.getMonth() + i - 1);
               const daysInMonth = getDaysDiff(prevDate, payDate);
               
               // Lãi tháng = Số ngày * Gốc * Tỷ lệ / 365
               let monthlyInterest = daysInMonth * (rate * remaining) / 365;
               
               // Nếu là kỳ cuối cùng thì trả cả gốc
               let principalPay = (i === term) ? remaining : 0;
               
               events.push({
                  date: payDate,
                  action: isDebt ? 'Phải trả' : 'Phải thu',
                  name: isDebt ? `${name} (Kỳ ${i}/${term})` : `${name} (Kỳ ${i}/${term})`,
                  remaining: remaining,
                  principalPayment: principalPay,
                  interestPayment: monthlyInterest
               });
               
               // Chỉ hiển thị 1 kỳ tiếp theo cho mỗi khoản vay (One event per loan)
               break; 
            }
         }
      }
      
      // 3. Trả góp gốc - lãi hàng tháng (Gốc đều, lãi giảm dần)
      else if (type === 'Trả góp gốc - lãi hàng tháng') {
         const monthlyPrincipal = initialPrincipal / term;
         let simulatedRemaining = initialPrincipal; // Bắt đầu tính từ đầu để khớp lịch
         
         for (let i = 1; i <= term; i++) {
            let payDate = new Date(startDate);
            payDate.setMonth(payDate.getMonth() + i);
            
            // Tính lãi cho kỳ này dựa trên dư nợ đầu kỳ
            let prevDate = new Date(startDate);
            prevDate.setMonth(prevDate.getMonth() + i - 1);
            const daysInMonth = getDaysDiff(prevDate, payDate);
            
            let monthlyInterest = daysInMonth * (rate * simulatedRemaining) / 365;
            
            // Nếu ngày trả >= hôm nay thì hiển thị
            if (payDate >= today) {
               events.push({
                  date: payDate,
                  action: isDebt ? 'Phải trả' : 'Phải thu',
                  name: isDebt ? `${name} (Kỳ ${i}/${term})` : `${name} (Kỳ ${i}/${term})`,
                  remaining: simulatedRemaining, // Dư nợ đầu kỳ
                  principalPayment: monthlyPrincipal,
                  interestPayment: monthlyInterest
               });
               
               // Chỉ hiển thị 1 kỳ tiếp theo cho mỗi khoản vay (One event per loan)
               break;
            }
            
            simulatedRemaining -= monthlyPrincipal;
            if (simulatedRemaining < 0) simulatedRemaining = 0;
         }
      }
      
      // Default: Fallback to old logic (Bullet) if type is unknown
      else {
         if (maturityDate && maturityDate >= today) {
           const days = getDaysDiff(startDate, maturityDate);
           let currentInterest = days * (rate * remaining) / 365;
           
           events.push({
              date: maturityDate,
              action: isDebt ? 'Phải trả' : 'Phải thu',
              name: isDebt ? `${name} (Tất toán)` : `${name} (Đáo hạn)`,
              remaining: remaining,
              principalPayment: remaining,
              interestPayment: currentInterest
           });
        }
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
