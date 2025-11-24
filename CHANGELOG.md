# Changelog

Mọi thay đổi đáng chú ý của dự án sẽ được ghi lại trong file này.

## [3.5.8] - 2025-11-24
### Dashboard Optimization & Code Cleanup
- ✅ **OPTIMIZATION**: Custom Dashboard Functions Integration
  - Replaced 200+ lines of SUMIFS formulas with 5 custom functions (`hffsIncome`, `hffsExpense`, `hffsDebt`, `hffsAssets`, `hffsYearly`)
  - Added dynamic column headers to all dashboard tables
  - Implemented dynamic row sizing to eliminate empty rows
  - Optimized `hffsYearly` performance (12x faster, prevents timeout)
  
- ✅ **FIX**: Dashboard Data Accuracy
  - Fixed cash calculation in Assets table (`netCash = Income - Expense`)
  - Fixed budget reading to use correct column (Column C instead of B)
  - Added budget multiplier based on filter period (x1 for month, x3 for quarter, x12 for year)
  - Set default filters to current year/quarter/month
  - Added TỔNG and TRUNG BÌNH rows to 12-month statistics table
  
- ✅ **NEW**: Auto-Refresh System
  - Created `DashboardRefresh.gs` with automatic recalculation after data submission
  - Integrated auto-refresh into 9 major data functions
  - Dashboard now updates automatically without manual refresh
  
- ✅ **FIX**: Debt Transaction Logic
  - ALL debt types now record Income (receiving loan money)
  - Installment loans additionally record Expense (immediate purchase)
  - Corrected cash flow representation
  - Updated SyncManager to handle both THU and CHI deletion for debt records
  
- ✅ **REFACTOR**: Code Quality Improvements
  - Eliminated duplicate logic in `DebtManagementHandler` (-51% code, 180→89 lines)
  - Eliminated duplicate logic in `LendingHandler` (-49% code, 175→89 lines)
  - Removed duplicate utility functions from `Main.gs` (getSpreadsheet, getSheet, formatCurrency, forceRecalculate)
  - Established single source of truth pattern (Forms → Handlers → DataProcessor)
  - Reduced duplicate function rate from 4.7% to 0%

- ✅ **ARCHITECTURE**: Clean Separation of Concerns
  - Forms handle input collection
  - Handlers provide validation and UX
  - DataProcessor is single source of truth for all data mutations
  - Utils.gs consolidates all utility functions

## [3.5.7] - 2025-11-23
### Tự động ghi nhận Chi phí Đầu tư
- ✅ **NEW**: Thêm danh mục "Đầu tư" vào nhóm Chi phí.
- ✅ **LOGIC**: Tự động tạo khoản Chi phí tương ứng khi thực hiện giao dịch MUA (Chứng khoán, Vàng, Crypto, Đầu tư khác).
- ✅ **SYNC**: Đảm bảo dòng tiền ra được ghi nhận đầy đủ vào báo cáo Chi tiêu.

## [3.5.6] - 2025-11-23
### Sửa lỗi Dashboard
- ✅ **FIX**: Sửa lỗi tính toán vị trí dòng Tổng Chi phí (do sai lệch bộ lọc danh mục).
- ✅ **FIX**: Biểu đồ Tổng quan giờ hiển thị Chi phí Thực tế (Cột F) thay vì Ngân sách (Cột G).
- ✅ **UI**: Cập nhật hiển thị biểu đồ chính xác hơn.

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

## [3.5.8] - 2025-11-24
### Fixed
- **Dashboard Data Accuracy**:
  - Updated "12-Month Statistics" formulas for Investment columns (CK, Gold, Crypto) to show **Total Invested Capital** instead of Profit/Loss.
  - Updated "Other Investment" column to sum **Principal Amount** + **Lending Principal**.
  - Fixed "Other Investment" data entry bug where Note was overwriting Interest Rate.
  - Fixed "Gold" price lookup formula to use correct Gold Type column.
  - Fixed "Crypto" price lookup formula to correctly append "USD" for Yahoo Finance.
- **System Stability**:
  - Removed pre-filled formulas in Investment sheets to prevent new data from being added at row 1000+. Formulas are now automatically applied to new rows.

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
