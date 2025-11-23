/**
 * ===============================================
 * CHANGELOG_DATA.GS
 * ===============================================
 * 
 * Lưu trữ dữ liệu lịch sử cập nhật của hệ thống
 */

const CHANGELOG_HISTORY = [
  {
    version: '3.5.5',
    date: '2025-11-23',
    title: 'Tinh chỉnh Logic Lịch sự kiện',
    changes: [
      '✅ FIX: Cập nhật logic Lịch sự kiện cho "Vay trả góp" và "Trả lãi hàng tháng"',
      '✅ UPDATE: Tính lãi dựa trên số ngày thực tế trong tháng của ngày sự kiện',
      '✅ UPDATE: "Vay trả góp" giờ sử dụng gốc còn lại thực tế từ sheet để tính lãi',
      '✅ UPDATE: Đảm bảo chỉ hiển thị duy nhất 1 sự kiện gần nhất cho mỗi khoản vay'
    ],
    actions: [
      'Chạy "Cập nhật Dashboard" để áp dụng thay đổi',
      'Kiểm tra lại bảng Lịch sự kiện (Khoản phải trả/thu)'
    ]
  },
  {
    version: '3.5.4',
    date: '2025-11-23',
    title: 'Thao tác nhanh trên Dashboard',
    changes: [
      '✅ NEW: Checkbox thao tác nhanh trên Dashboard để nhập liệu nhanh',
      '✅ NEW: Tích hợp Trigger onEdit để mở form từ Dashboard',
      '✅ UI: Thêm nhãn trực quan cho các thao tác nhanh'
    ],
    actions: [
      'Chạy "Cập nhật Dashboard" để hiển thị Checkbox',
      'Cấp quyền cho Script nếu được yêu cầu để mở Form từ Trigger'
    ]
  },
  {
    version: '3.5.3',
    date: '2025-11-23',
    title: 'Tái cấu trúc Hệ thống Nợ',
    changes: [
      '✅ REFACTOR: Thống nhất hệ thống loại hình Nợ & Cho vay (Logic chia sẻ)',
      '✅ FIX: Lịch sự kiện cho Khoản phải trả & Phải thu',
      '✅ UPDATE: Tương thích ngược cho dữ liệu cũ',
      '✅ FIX: Loại nợ trong form giờ khớp với logic tính toán'
    ],
    actions: [
      'Chạy "Cập nhật Dashboard" để áp dụng thay đổi',
      'Kiểm tra lại bảng Lịch sự kiện (Khoản phải trả/thu)'
    ]
  },
  {
    version: '3.5.2',
    date: '2025-11-23',
    title: 'Nâng cấp Báo cáo Chi phí',
    changes: [
      '✅ NEW: Báo cáo chi phí chi tiết (Ngân sách, Còn lại, Trạng thái)',
      '✅ UPDATE: Tự động lấy dữ liệu từ Sheet BUDGET',
      '✅ FIX: Logic Khoản phải trả trên Dashboard (Vay trả góp)'
    ],
    actions: [
      'Chạy "Cập nhật Dashboard" để áp dụng giao diện mới',
      'Kiểm tra lại Sheet BUDGET để đảm bảo ngân sách đã được thiết lập'
    ]
  },
  {
    version: '3.5.1',
    date: '2025-11-23',
    title: 'Tinh chỉnh Ngân sách & Đồng bộ',
    changes: [
      '✅ UPDATE: Di chuyển nhóm "Trả nợ" xuống sau "Đầu tư" trong sheet Budget',
      '✅ NEW: Thêm cảnh báo đồng bộ nếu Ngân sách Trả nợ < Số tiền phải trả dự kiến',
      '✅ FIX: Loại bỏ các mục sai trong nhóm Nợ'
    ],
    actions: [
      'Chạy "Khởi tạo lại toàn bộ" để cập nhật cấu trúc Sheet Budget',
      'Kiểm tra cảnh báo đồng bộ khi khởi tạo'
    ]
  },
  {
    version: '3.5.0',
    date: '2025-11-23',
    title: 'Tái cấu trúc Cho vay & Cải tiến UI',
    changes: [
      '✅ NEW: Tái cấu trúc hệ thống Cho vay với 3 loại hình cụ thể',
      '✅ NEW: Lịch sự kiện động sử dụng Custom Functions',
      '✅ UI: Chuẩn hóa độ rộng cột Dashboard (120px) và mở rộng cột M (200px)',
      '✅ UI: Viền bảng nhạt hơn để dễ nhìn',
      '✅ UI: Di chuyển Biểu đồ xuống L39, trục Y chia theo nghìn đồng, Nợ màu Đỏ'
    ],
    actions: [
      'Chạy "Cập nhật Dashboard" để áp dụng giao diện mới',
      'Kiểm tra lại loại hình cho vay trong sheet CHO VAY nếu cần'
    ]
  },
  {
    version: '3.4.1',
    date: '2025-05-22',
    title: 'Sửa lỗi Menu Ngân sách & Báo cáo',
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
    title: 'Thống nhất Cấu trúc Ngân sách',
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
    title: 'Cải tiến Setup Wizard',
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
    title: 'Giới thiệu Setup Wizard',
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
