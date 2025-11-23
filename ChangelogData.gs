/**
 * ===============================================
 * CHANGELOG_DATA.GS
 * ===============================================
 * 
 * Lưu trữ dữ liệu lịch sử cập nhật của hệ thống
 */

const CHANGELOG_HISTORY = [
  {
    version: '3.5.0',
    date: '2025-11-23',
    title: 'Lending Refactor & Dynamic Dashboard',
    changes: [
      '✅ NEW: Refactored Lending System with 3 specific types (Maturity, Interest Only, Installment)',
      '✅ NEW: Dynamic Event Calendar using Custom Functions (AccPayable, AccReceivable)',
      '✅ FIX: Debt Payment & Collection forms now support precise amounts',
      '✅ FIX: Dashboard tables row offset issues',
      '✅ UPDATE: Event Calendar now shows only the nearest upcoming event per loan'
    ],
    actions: [
      'Chạy "Cập nhật Dashboard" để áp dụng công thức lịch sự kiện mới',
      'Kiểm tra lại loại hình cho vay trong sheet CHO VAY nếu cần'
    ]
  },
  {
    version: '3.4.1',
    date: '2025-05-22',
    title: 'Fix Budget Menu & Reports',
    changes: [
      '✅ FIX: Menu "Đặt Ngân sách tháng" không hoạt động',
      '✅ FIX: Menu "Kiểm tra Budget" không hoạt động',
      '✅ FIX: Menu "Báo cáo Chi tiêu" không hoạt động',
      '✅ FIX: Menu "Báo cáo Đầu tư" không hoạt động',
      '✅ NEW: Thêm wrapper functions cho các tính năng Budget'
    ],
    actions: [
      'Không cần thực hiện hành động gì đặc biệt',
      'Thử nghiệm các tính năng trong menu "Ngân sách"'
    ]
  },
  {
    version: '3.4',
    date: '2025-05-20',
    title: 'Unified Budget Structure',
    changes: [
      '✅ UPDATE: Thống nhất cấu trúc Sheet BUDGET cho cả Setup Wizard và Menu khởi tạo',
      '✅ FIX: Lỗi lệch dòng khi chạy Setup Wizard (dữ liệu bị đẩy xuống dòng 1000+)',
      '✅ NEW: Tự động điền dữ liệu từ Setup Wizard vào các ô tham số Budget',
      '✅ NEW: Cập nhật công thức tính toán ngân sách tự động'
    ],
    actions: [
      'Nếu đang dùng bản cũ, hãy chạy "Khởi tạo Sheet BUDGET" để cập nhật cấu trúc mới',
      'Kiểm tra lại các mục tiêu ngân sách trong sheet BUDGET'
    ]
  },
  {
    version: '3.3',
    date: '2025-05-18',
    title: 'Setup Wizard Improvements',
    changes: [
      '✅ FIX: Lỗi appendRow() chèn dữ liệu sai vị trí',
      '✅ UPDATE: Đổi trạng thái nợ mặc định từ "Đang trả" sang "Chưa trả"',
      '✅ FIX: Format lãi suất hiển thị đúng %',
      '✅ NEW: Tự động thêm khoản thu khi khai báo nợ trong Setup Wizard'
    ],
    actions: [
      'Kiểm tra lại các khoản nợ mới thêm để đảm bảo thông tin chính xác'
    ]
  },
  {
    version: '3.2',
    date: '2025-05-15',
    title: 'Setup Wizard Introduction',
    changes: [
      '✅ NEW: Giới thiệu Setup Wizard để thiết lập hệ thống lần đầu',
      '✅ NEW: Hướng dẫn từng bước nhập số dư, nợ và ngân sách',
      '✅ UPDATE: Cải thiện giao diện khởi tạo'
    ],
    actions: [
      'Dùng menu "Khởi tạo TẤT CẢ Sheet" để trải nghiệm Setup Wizard nếu muốn reset hệ thống'
    ]
  }
];
