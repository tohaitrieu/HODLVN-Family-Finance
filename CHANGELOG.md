# Changelog

All notable changes to this project will be documented in this file.

## [3.5.0] - 2025-11-23
### Lending Refactor & Dynamic Dashboard
- ✅ **NEW**: Refactored Lending System with 3 specific types (Maturity, Interest Only, Installment).
- ✅ **NEW**: Dynamic Event Calendar using Custom Functions (`AccPayable`, `AccReceivable`).
- ✅ **FIX**: Debt Payment & Collection forms now support precise amounts (removed 1000 step).
- ✅ **FIX**: Dashboard tables row offset issues.
- ✅ **UPDATE**: Event Calendar now shows only the nearest upcoming event per loan.
- ✅ **UI**: Dashboard column widths standardized (120px) and M column widened (200px).
- ✅ **UI**: Lighter table borders for better visibility.
- ✅ **UI**: Chart moved to L39, Y-axis scaled (x1000 VND), and Debt series colored Red.

## [3.5.1] - 2025-11-23
### Budget Refinement & Sync
- ✅ **UPDATE**: Moved "Trả nợ" group after "Đầu tư" in Budget sheet.
- ✅ **NEW**: Added Sync Warning if Budgeted Debt Payment < Expected Debt Payment.
- ✅ **FIX**: Removed incorrect items from Debt group.

## [3.4.1] - 2025-05-22
### Fix Budget Menu & Reports
- ✅ **FIX**: Menu "Đặt Ngân sách tháng" không hoạt động.
- ✅ **FIX**: Menu "Kiểm tra Budget" không hoạt động.
- ✅ **FIX**: Menu "Báo cáo Chi tiêu" không hoạt động.
- ✅ **FIX**: Menu "Báo cáo Đầu tư" không hoạt động.
- ✅ **NEW**: Thêm wrapper functions cho các tính năng Budget.

## [3.4] - 2025-05-20
### Unified Budget Structure
- ✅ **UPDATE**: Thống nhất cấu trúc Sheet BUDGET cho cả Setup Wizard và Menu khởi tạo.
- ✅ **FIX**: Lỗi lệch dòng khi chạy Setup Wizard (dữ liệu bị đẩy xuống dòng 1000+).
- ✅ **NEW**: Tự động điền dữ liệu từ Setup Wizard vào các ô tham số Budget.
- ✅ **NEW**: Cập nhật công thức tính toán ngân sách tự động.

## [3.3] - 2025-05-18
### Setup Wizard Improvements
- ✅ **FIX**: Lỗi appendRow() chèn dữ liệu sai vị trí.
- ✅ **UPDATE**: Đổi trạng thái nợ mặc định từ "Đang trả" sang "Chưa trả".
- ✅ **FIX**: Format lãi suất hiển thị đúng %.
- ✅ **NEW**: Tự động thêm khoản thu khi khai báo nợ trong Setup Wizard.

## [3.2] - 2025-05-15
### Setup Wizard Introduction
- ✅ **NEW**: Giới thiệu Setup Wizard để thiết lập hệ thống lần đầu.
- ✅ **NEW**: Hướng dẫn từng bước nhập số dư, nợ và ngân sách.
- ✅ **UPDATE**: Cải thiện giao diện khởi tạo.
