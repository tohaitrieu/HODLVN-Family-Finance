# Changelog

Mọi thay đổi đáng chú ý của dự án sẽ được ghi lại trong file này.

## [3.5.5] - 2025-11-23
### Tinh chỉnh Logic Lịch sự kiện
- ✅ **FIX**: Cập nhật logic Lịch sự kiện cho "Vay trả góp" và "Trả lãi hàng tháng".
- ✅ **UPDATE**: Tính lãi dựa trên số ngày thực tế trong tháng của ngày sự kiện.
- ✅ **UPDATE**: "Vay trả góp" giờ sử dụng gốc còn lại thực tế từ sheet để tính lãi.
- ✅ **UPDATE**: Đảm bảo chỉ hiển thị duy nhất 1 sự kiện gần nhất cho mỗi khoản vay.

## [3.5.4] - 2025-11-23
### Thao tác nhanh trên Dashboard
- ✅ **NEW**: Checkbox thao tác nhanh trên Dashboard để nhập liệu nhanh.
- ✅ **NEW**: Tích hợp Trigger `onEdit` để mở form từ Dashboard.
- ✅ **UI**: Thêm nhãn trực quan cho các thao tác nhanh (Thu, Chi, Nợ, Cho vay, Đầu tư).

## [3.5.3] - 2025-11-23
### Tái cấu trúc Cho vay & Dashboard Động
- ✅ **NEW**: Tái cấu trúc hệ thống Cho vay với 3 loại hình cụ thể (Tất toán, Trả lãi, Trả góp).
- ✅ **NEW**: Lịch sự kiện động sử dụng Custom Functions (`AccPayable`, `AccReceivable`).
- ✅ **FIX**: Form Trả nợ & Thu nợ hỗ trợ số tiền chính xác (bỏ bước nhảy 1000).
- ✅ **FIX**: Sửa lỗi lệch dòng bảng Dashboard.
- ✅ **UPDATE**: Lịch sự kiện chỉ hiển thị sự kiện sắp tới gần nhất cho mỗi khoản vay.
- ✅ **UI**: Chuẩn hóa độ rộng cột Dashboard (120px) và mở rộng cột M (200px).
- ✅ **UI**: Viền bảng nhạt hơn để dễ nhìn.
- ✅ **UI**: Di chuyển Biểu đồ xuống L39, trục Y chia theo nghìn đồng, Nợ màu Đỏ.

## [3.5.1] - 2025-11-23
### Tinh chỉnh Ngân sách & Đồng bộ
- ✅ **UPDATE**: Di chuyển nhóm "Trả nợ" xuống sau "Đầu tư" trong sheet Budget.
- ✅ **NEW**: Thêm cảnh báo đồng bộ nếu Ngân sách Trả nợ < Số tiền phải trả dự kiến.
- ✅ **FIX**: Loại bỏ các mục sai trong nhóm Nợ.

## [3.4.1] - 2025-05-22
### Sửa lỗi Menu Ngân sách & Báo cáo
- ✅ **FIX**: Menu "Đặt Ngân sách tháng" không hoạt động.
- ✅ **FIX**: Menu "Kiểm tra Budget" không hoạt động.
- ✅ **FIX**: Menu "Báo cáo Chi tiêu" không hoạt động.
- ✅ **FIX**: Menu "Báo cáo Đầu tư" không hoạt động.
- ✅ **NEW**: Thêm wrapper functions cho các tính năng Budget.

## [3.4] - 2025-05-20
### Thống nhất Cấu trúc Ngân sách
- ✅ **UPDATE**: Thống nhất cấu trúc Sheet BUDGET cho cả Setup Wizard và Menu khởi tạo.
- ✅ **FIX**: Lỗi lệch dòng khi chạy Setup Wizard (dữ liệu bị đẩy xuống dòng 1000+).
- ✅ **NEW**: Tự động điền dữ liệu từ Setup Wizard vào các ô tham số Budget.
- ✅ **NEW**: Cập nhật công thức tính toán ngân sách tự động.

## [3.3] - 2025-05-18
### Cải tiến Setup Wizard
- ✅ **FIX**: Lỗi appendRow() chèn dữ liệu sai vị trí.
- ✅ **UPDATE**: Đổi trạng thái nợ mặc định từ "Đang trả" sang "Chưa trả".
- ✅ **FIX**: Format lãi suất hiển thị đúng %.
- ✅ **NEW**: Tự động thêm khoản thu khi khai báo nợ trong Setup Wizard.

## [3.2] - 2025-05-15
### Giới thiệu Setup Wizard
- ✅ **NEW**: Giới thiệu Setup Wizard để thiết lập hệ thống lần đầu.
- ✅ **NEW**: Hướng dẫn từng bước nhập số dư, nợ và ngân sách.
- ✅ **UPDATE**: Cải thiện giao diện khởi tạo.
