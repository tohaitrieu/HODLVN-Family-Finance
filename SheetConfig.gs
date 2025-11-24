/**
 * ===============================================
 * SHEETCONFIG.GS - CẤU HÌNH CẤU TRÚC DỮ LIỆU
 * ===============================================
 * 
 * Định nghĩa Schema chuẩn cho toàn bộ hệ thống.
 * Mọi thao tác tạo sheet, đọc/ghi dữ liệu phải tham chiếu từ đây.
 */

// GLOBAL FORMATTING CONFIGURATION
const GLOBAL_SHEET_CONFIG = {
  DEFAULT_ROW_HEIGHT: 28,
  HEADER_ROW_HEIGHT: 32,
  TITLE_ROW_HEIGHT: 35,
  DEFAULT_CELL_BACKGROUND: '#fff1e5',
  BORDER_COLOR: '#B0B0B0',
  BORDER_STYLE: SpreadsheetApp.BorderStyle.SOLID
};

const SHEET_CONFIG = {
  
  // 1. THU NHẬP
  INCOME: {
    name: 'THU', // Tên Sheet thực tế
    columns: [
      { key: 'stt', header: 'STT', width: 50, type: 'number', format: '0' },
      { key: 'date', header: 'Ngày', width: 100, type: 'date', format: 'dd/mm/yyyy' },
      { key: 'amount', header: 'Số tiền', width: 120, type: 'number', format: '#,##0' },
      { key: 'category', header: 'Nguồn thu', width: 150, type: 'dropdown', source: 'INCOME' },
      { key: 'note', header: 'Ghi chú', width: 300, type: 'text' },
      { key: 'transactionId', header: 'TransactionID', width: 0, type: 'hidden' }
    ]
  },

  // 2. CHI TIÊU
  EXPENSE: {
    name: 'CHI',
    columns: [
      { key: 'stt', header: 'STT', width: 50, type: 'number', format: '0' },
      { key: 'date', header: 'Ngày', width: 100, type: 'date', format: 'dd/mm/yyyy' },
      { key: 'amount', header: 'Số tiền', width: 120, type: 'number', format: '#,##0' },
      { key: 'category', header: 'Danh mục', width: 120, type: 'dropdown', source: 'EXPENSE' },
      { key: 'subcategory', header: 'Chi tiết', width: 200, type: 'text' },
      { key: 'note', header: 'Ghi chú', width: 250, type: 'text' },
      { key: 'transactionId', header: 'TransactionID', width: 0, type: 'hidden' }
    ]
  },

  // 3. TRẢ NỢ (Lịch sử thanh toán)
  DEBT_PAYMENT: {
    name: 'TRẢ NỢ',
    columns: [
      { key: 'stt', header: 'STT', width: 50, type: 'number', format: '0' },
      { key: 'date', header: 'Ngày', width: 100, type: 'date', format: 'dd/mm/yyyy' },
      { key: 'debtName', header: 'Khoản nợ', width: 150, type: 'text' },
      { key: 'principal', header: 'Trả gốc', width: 120, type: 'number', format: '#,##0' },
      { key: 'interest', header: 'Trả lãi', width: 120, type: 'number', format: '#,##0' },
      { key: 'total', header: 'Tổng trả', width: 120, type: 'formula', formula: '=IFERROR(RC[-2]+RC[-1], 0)', format: '#,##0' },
      { key: 'note', header: 'Ghi chú', width: 250, type: 'text' },
      { key: 'transactionId', header: 'TransactionID', width: 0, type: 'hidden' }
    ]
  },

  // 4. QUẢN LÝ NỢ (Danh sách khoản nợ)
  DEBT_MANAGEMENT: {
    name: 'QUẢN LÝ NỢ',
    columns: [
      { key: 'stt', header: 'STT', width: 50, type: 'number', format: '0' },
      { key: 'name', header: 'Tên khoản nợ', width: 150, type: 'text' },
      { key: 'type', header: 'Loại hình', width: 250, type: 'dropdown', source: 'LOAN_TYPES' },
      { key: 'principal', header: 'Nợ gốc ban đầu', width: 100, type: 'number', format: '#,##0' },
      { key: 'rate', header: 'Lãi suất (%/năm)', width: 100, type: 'number', format: '0.00"%"' },
      { key: 'term', header: 'Kỳ hạn (tháng)', width: 100, type: 'number', format: '0' },
      { key: 'startDate', header: 'Ngày vay', width: 100, type: 'date', format: 'dd/mm/yyyy' },
      { key: 'endDate', header: 'Ngày đến hạn', width: 120, type: 'date', format: 'dd/mm/yyyy' },
      { key: 'paidPrincipal', header: 'Đã trả gốc', width: 120, type: 'number', format: '#,##0' },
      { key: 'paidInterest', header: 'Đã trả lãi', width: 120, type: 'number', format: '#,##0' },
      { key: 'remaining', header: 'Còn nợ', width: 100, type: 'formula', formula: '=IFERROR(RC[-7]-RC[-2], 0)', format: '#,##0' },
      { key: 'status', header: 'Trạng thái', width: 200, type: 'dropdown', source: 'DEBT_STATUS' },
      { key: 'note', header: 'Ghi chú', width: 200, type: 'text' },
      { key: 'transactionId', header: 'TransactionID', width: 0, type: 'hidden' }
    ]
  },

  // 5. CHO VAY (Danh sách khoản cho vay)
  LENDING: {
    name: 'CHO VAY',
    columns: [
      { key: 'stt', header: 'STT', width: 50, type: 'number', format: '0' },
      { key: 'name', header: 'Tên người vay', width: 150, type: 'text' },
      { key: 'type', header: 'Loại hình', width: 250, type: 'dropdown', source: 'LOAN_TYPES' },
      { key: 'principal', header: 'Số tiền gốc', width: 100, type: 'number', format: '#,##0' },
      { key: 'rate', header: 'Lãi suất (%/năm)', width: 100, type: 'number', format: '0.00"%"' },
      { key: 'term', header: 'Kỳ hạn (tháng)', width: 100, type: 'number', format: '0' },
      { key: 'startDate', header: 'Ngày vay', width: 100, type: 'date', format: 'dd/mm/yyyy' },
      { key: 'endDate', header: 'Ngày đến hạn', width: 120, type: 'date', format: 'dd/mm/yyyy' },
      { key: 'paidPrincipal', header: 'Gốc đã thu', width: 120, type: 'number', format: '#,##0' },
      { key: 'paidInterest', header: 'Lãi đã thu', width: 120, type: 'number', format: '#,##0' },
      { key: 'remaining', header: 'Còn lại', width: 100, type: 'formula', formula: '=IFERROR(RC[-7]-RC[-2], 0)', format: '#,##0' },
      { key: 'status', header: 'Trạng thái', width: 200, type: 'dropdown', source: 'LENDING_STATUS' },
      { key: 'note', header: 'Ghi chú', width: 200, type: 'text' },
      { key: 'transactionId', header: 'TransactionID', width: 0, type: 'hidden' }
    ]
  },

  // 6. THU NỢ (Lịch sử thu nợ)
  LENDING_REPAYMENT: {
    name: 'THU NỢ',
    columns: [
      { key: 'stt', header: 'STT', width: 50, type: 'number', format: '0' },
      { key: 'date', header: 'Ngày', width: 100, type: 'date', format: 'dd/mm/yyyy' },
      { key: 'borrower', header: 'Người vay', width: 150, type: 'text' },
      { key: 'principal', header: 'Thu gốc', width: 120, type: 'number', format: '#,##0' },
      { key: 'interest', header: 'Thu lãi', width: 120, type: 'number', format: '#,##0' },
      { key: 'total', header: 'Tổng thu', width: 120, type: 'formula', formula: '=IFERROR(RC[-2]+RC[-1], 0)', format: '#,##0' },
      { key: 'note', header: 'Ghi chú', width: 250, type: 'text' },
      { key: 'transactionId', header: 'TransactionID', width: 0, type: 'hidden' }
    ]
  },

  // 7. CHỨNG KHOÁN
  STOCK: {
    name: 'CHỨNG KHOÁN',
    columns: [
      { key: 'stt', header: 'STT', width: 50, type: 'number', format: '0' },
      { key: 'date', header: 'Ngày', width: 100, type: 'date', format: 'dd/mm/yyyy' },
      { key: 'type', header: 'Loại GD', width: 80, type: 'dropdown', source: 'STOCK_TRANSACTION_TYPES' },
      { key: 'ticker', header: 'Mã CK', width: 80, type: 'text' },
      { key: 'quantity', header: 'Số lượng', width: 80, type: 'number', format: '#,##0' },
      { key: 'price', header: 'Giá gốc', width: 100, type: 'number', format: '#,##0' },
      { key: 'fee', header: 'Phí', width: 100, type: 'number', format: '#,##0' },
      { key: 'totalCost', header: 'Tổng vốn', width: 120, type: 'number', format: '#,##0' },
      { key: 'divCash', header: '💰 Cổ tức TM', width: 110, type: 'number', format: '#,##0' },
      { key: 'divStock', header: '📈 Cổ tức CP', width: 100, type: 'number', format: '0' },
      { key: 'adjPrice', header: '📊 Giá ĐC', width: 100, type: 'formula', formula: '=IF(RC[-6]>0, (RC[-3]-RC[-2])/RC[-6], 0)', format: '#,##0' },
      { key: 'marketPrice', header: '💹 Giá HT', width: 100, type: 'formula', formula: '=IF(RC[-8]<>"", MPRICE(RC[-8]), 0)', format: '#,##0' },
      { key: 'marketValue', header: '💵 Giá trị HT', width: 120, type: 'formula', formula: '=IF(AND(RC[-8]>0, RC[-1]>0), RC[-8]*RC[-1], 0)', format: '#,##0' },
      { key: 'profit', header: '📈 Lãi/Lỗ', width: 110, type: 'formula', formula: '=IF(RC[-1]>0, RC[-1]-(RC[-6]-RC[-5]), 0)', format: '#,##0' },
      { key: 'profitPercent', header: '📊 % L/L', width: 80, type: 'formula', formula: '=IF(AND(RC[-1]<>0, (RC[-7]-RC[-6])>0), RC[-1]/(RC[-7]-RC[-6]), 0)', format: '0.00%' },
      { key: 'note', header: 'Ghi chú', width: 250, type: 'text' }
    ]
  },

  // 8. VÀNG
  GOLD: {
    name: 'VÀNG',
    columns: [
      { key: 'stt', header: 'STT', width: 50, type: 'number', format: '0' },
      { key: 'date', header: 'Ngày', width: 100, type: 'date', format: 'dd/mm/yyyy' },
      { key: 'assetName', header: 'Tài sản', width: 80, type: 'text' }, // Mặc định 'GOLD'
      { key: 'type', header: 'Loại GD', width: 80, type: 'dropdown', source: 'GOLD_TRANSACTION_TYPES' },
      { key: 'goldType', header: 'Loại vàng', width: 100, type: 'dropdown', source: 'GOLD_TYPES' },
      { key: 'quantity', header: 'Số lượng', width: 80, type: 'number', format: '#,##0.00' },
      { key: 'unit', header: 'Đơn vị', width: 70, type: 'dropdown', source: 'GOLD_UNITS' },
      { key: 'price', header: 'Giá vốn', width: 100, type: 'number', format: '#,##0' },
      { key: 'totalCost', header: 'Tổng vốn', width: 120, type: 'number', format: '#,##0' },
      { key: 'marketPrice', header: 'Giá HT', width: 100, type: 'formula', formula: '=IF(RC[-5]<>"", GPRICE(RC[-5]), 0)', format: '#,##0' },
      { key: 'marketValue', header: 'Giá trị HT', width: 120, type: 'formula', formula: '=IF(AND(RC[-5]>0, RC[-1]>0), RC[-5]*RC[-1], 0)', format: '#,##0' },
      { key: 'profit', header: 'Lãi/Lỗ', width: 110, type: 'formula', formula: '=IF(RC[-1]>0, RC[-1]-RC[-3], 0)', format: '#,##0' },
      { key: 'profitPercent', header: '% Lãi/Lỗ', width: 80, type: 'formula', formula: '=IF(RC[-4]>0, RC[-1]/RC[-4], 0)', format: '0.00%' },
      { key: 'note', header: 'Ghi chú', width: 200, type: 'text' }
    ]
  },

  // 9. CRYPTO
  CRYPTO: {
    name: 'CRYPTO',
    columns: [
      { key: 'stt', header: 'STT', width: 50, type: 'number', format: '0' },
      { key: 'date', header: 'Ngày', width: 100, type: 'date', format: 'dd/mm/yyyy' },
      { key: 'type', header: 'Loại GD', width: 80, type: 'dropdown', source: 'CRYPTO_TRANSACTION_TYPES' },
      { key: 'coin', header: 'Coin', width: 80, type: 'text' },
      { key: 'quantity', header: 'Số lượng', width: 100, type: 'number', format: '#,##0.0000' },
      { key: 'priceUSD', header: 'Giá (USD)', width: 100, type: 'number', format: '#,##0.00' },
      { key: 'rate', header: 'Tỷ giá', width: 80, type: 'number', format: '#,##0' },
      { key: 'priceVND', header: 'Giá (VND)', width: 100, type: 'number', format: '#,##0' },
      { key: 'totalCost', header: 'Tổng vốn', width: 120, type: 'number', format: '#,##0' },
      { key: 'marketPriceUSD', header: 'Giá HT (USD)', width: 100, type: 'formula', formula: '=IF(RC[-6]<>"", CPRICE(RC[-6]&"USD"), 0)', format: '#,##0.00' },
      { key: 'marketValueUSD', header: 'Giá trị HT (USD)', width: 120, type: 'formula', formula: '=IF(AND(RC[-6]>0, RC[-1]>0), RC[-6]*RC[-1], 0)', format: '#,##0.00' },
      { key: 'marketPriceVND', header: 'Giá HT (VND)', width: 100, type: 'formula', formula: '=IF(AND(RC[-2]>0, RC[-5]>0), RC[-2]*RC[-5], 0)', format: '#,##0' },
      { key: 'marketValueVND', header: 'Giá trị HT (VND)', width: 120, type: 'formula', formula: '=IF(AND(RC[-2]>0, RC[-6]>0), RC[-2]*RC[-6], 0)', format: '#,##0' },
      { key: 'profit', header: 'Lãi/Lỗ', width: 110, type: 'formula', formula: '=IF(RC[-1]>0, RC[-1]-RC[-5], 0)', format: '#,##0' },
      { key: 'profitPercent', header: '% Lãi/Lỗ', width: 80, type: 'formula', formula: '=IF(RC[-6]>0, RC[-1]/RC[-6], 0)', format: '0.00%' },
      { key: 'exchange', header: 'Sàn', width: 100, type: 'text' },
      { key: 'wallet', header: 'Ví', width: 150, type: 'text' },
      { key: 'note', header: 'Ghi chú', width: 200, type: 'text' }
    ]
  },

  // 10. ĐẦU TƯ KHÁC
  OTHER_INVESTMENT: {
    name: 'ĐẦU TƯ KHÁC',
    columns: [
      { key: 'stt', header: 'STT', width: 50, type: 'number', format: '0' },
      { key: 'date', header: 'Ngày', width: 100, type: 'date', format: 'dd/mm/yyyy' },
      { key: 'type', header: 'Loại đầu tư', width: 150, type: 'dropdown', source: 'OTHER_INVESTMENT_TYPES' },
      { key: 'amount', header: 'Số tiền', width: 120, type: 'number', format: '#,##0' },
      { key: 'rate', header: 'Lãi suất (%)', width: 100, type: 'number', format: '0.00"%"' },
      { key: 'term', header: 'Kỳ hạn (tháng)', width: 100, type: 'number', format: '0' },
      { key: 'expectedReturn', header: 'Dự kiến thu về', width: 120, type: 'number', format: '#,##0' },
      { key: 'note', header: 'Ghi chú', width: 250, type: 'text' }
    ]
  },

  // 11. BUDGET
  BUDGET: {
    name: 'BUDGET',
    columns: [
      { key: 'category', header: 'Danh mục', width: 200, type: 'text' },
      { key: 'percentage', header: '% Nhóm', width: 80, type: 'number', format: '0.0%' },
      { key: 'budget', header: 'Ngân sách', width: 120, type: 'number', format: '#,##0' },
      { key: 'spent', header: 'Đã chi', width: 120, type: 'number', format: '#,##0' },
      { key: 'remaining', header: 'Còn lại', width: 120, type: 'number', format: '#,##0' },
      { key: 'status', header: 'Trạng thái', width: 150, type: 'text' }
    ]
  },

  // 12. CHANGELOG
  CHANGELOG: {
    name: 'LỊCH SỬ CẬP NHẬT',
    columns: [
      { key: 'version', header: 'Phiên bản / Tính năng', width: 250, type: 'text' },
      { key: 'detail', header: 'Chi tiết thay đổi', width: 400, type: 'text' },
      { key: 'action', header: 'Hành động khuyến nghị', width: 300, type: 'text' }
    ]
  }
};
