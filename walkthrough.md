# Fix Dashboard & Budget Update Issues

## 1. Vấn đề đã xử lý

### ❌ Lỗi 1: Nhập liệu thủ công không cập nhật Sheet Tổng quan/Budget
**Nguyên nhân:**
- `BudgetManager.gs` cũ ghi đè giá trị tĩnh vào sheet BUDGET, làm mất các công thức tự động.
- `BudgetManager.gs` ghi dữ liệu vào sai cột (Ghi vào cột Ngân sách thay vì cột Đã chi), gây sai lệch dữ liệu.
- Không có trigger `onEdit` để xử lý khi nhập tay (tuy nhiên với việc khôi phục công thức, trigger này không còn bắt buộc nhưng đã được thêm để log).

**Giải pháp:**
- **Sửa `BudgetManager.gs`**: Vô hiệu hóa các hàm ghi đè giá trị thủ công. Chuyển sang sử dụng công thức `SUMIFS` tự động (đã có trong `SheetInitializer`).
- **Sửa `checkBudgetWarnings`**: Cập nhật lại chỉ số cột để đọc đúng dữ liệu từ sheet BUDGET mới.
- **Tạo `Triggers.gs`**: Thêm hàm `onEdit` để theo dõi thay đổi dữ liệu (hiện tại chỉ log, vì công thức sẽ tự tính toán).

### ❌ Lỗi 2: Menu "Cập nhật Dashboard" không hoạt động
**Nguyên nhân:**
- Hàm `refreshDashboard` nằm trong file `DashboardManager.gs` có thể gặp vấn đề về scope khi gọi từ Menu.

**Giải pháp:**
- Chuyển hàm `refreshDashboard` sang `Main.gs` để đảm bảo Menu luôn tìm thấy hàm này.

## 2. Hướng dẫn Cập nhật (QUAN TRỌNG)

Do code cũ đã ghi đè và làm hỏng các công thức trong sheet BUDGET, bạn cần **Khởi tạo lại sheet BUDGET** để khôi phục các công thức tự động.

### 🛠️ Bước 1: Khởi tạo lại Sheet BUDGET
1. Trên thanh menu, chọn **HODLVN Family Finance** (hoặc tên App của bạn).
2. Chọn **⚙️ Khởi tạo Sheet** > **💰 Khởi tạo Sheet BUDGET**.
3. Xác nhận **OK**.
   > ⚠️ Lưu ý: Việc này sẽ reset sheet BUDGET về trạng thái ban đầu (có công thức). Dữ liệu chi tiêu thực tế sẽ tự động được tính toán lại từ sheet CHI.

### 🛠️ Bước 2: Kiểm tra lại
1. Thử nhập liệu thủ công vào sheet **CHI** hoặc **THU**.
2. Sang sheet **BUDGET**, kiểm tra cột **Đã chi** xem có tự động nhảy số không.
3. Sang sheet **TỔNG QUAN**, bấm menu **📊 Thống kê & Dashboard** > **🔄 Cập nhật Dashboard** để kiểm tra nút bấm.

## 3. Chi tiết thay đổi Code

### `BudgetManager.gs`
- `updateBudgetSpent`, `updateInvestmentBudget`, `updateDebtBudget`: Đã vô hiệu hóa (để công thức tự chạy).
- `checkBudgetWarnings`: Sửa index cột để đọc đúng cột Ngân sách (Col C) và Đã chi (Col D).
- `setBudgetForMonth`: Sửa lỗi ghi đè cột Ngân sách. Giờ sẽ cập nhật cột % (Col B) để công thức tự tính ra Ngân sách.

### `Main.gs`
- Thêm hàm `refreshDashboard()` để xử lý menu.

### `Triggers.gs` (Mới)
- Thêm file mới để quản lý các sự kiện `onEdit`.
