# 🤝 Hướng dẫn sử dụng: Cho vay (Lending)

Tính năng **Cho vay** giúp bạn quản lý các khoản tiền cho người khác vay, theo dõi lịch trả nợ, tính lãi và quản lý dòng tiền thu hồi.

---

## 📑 MỤC LỤC

1. [Tổng quan](#-tổng-quan)
2. [Thêm khoản cho vay mới](#-thêm-khoản-cho-vay-mới)
3. [Thu hồi nợ & lãi](#-thu-hồi-nợ--lãi)
4. [Các loại hình cho vay](#-các-loại-hình-cho-vay)
5. [Tác động đến Cash Flow](#-tác-động-đến-cash-flow)

---

## 📋 Tổng quan

Hệ thống quản lý cho vay bao gồm 2 sheet chính:
- **CHO VAY**: Danh sách các khoản đang cho vay, trạng thái và số dư.
- **THU NỢ**: Lịch sử các lần thu hồi gốc và lãi.

---

## ➕ Thêm khoản cho vay mới

### Bước 1: Mở form
Trên thanh menu, chọn **HODLVN Finance** → **🤝 Cho vay** → **➕ Thêm khoản vay**.

### Bước 2: Điền thông tin
- **Ngày cho vay**: Ngày giải ngân.
- **Tên người vay**: Tên cá nhân hoặc tổ chức vay.
- **Loại cho vay**: Chọn hình thức trả nợ (xem chi tiết bên dưới).
- **Số tiền gốc**: Số tiền bạn cho vay.
- **Lãi suất (%/năm)**: Lãi suất thỏa thuận (nhập 0 nếu cho vay không lãi).
- **Kỳ hạn (tháng)**: Thời gian vay.

### Bước 3: Xác nhận
Nhấn **Thêm khoản cho vay**.

> [!NOTE]
> **Tự động tạo Chi tiêu**: Khi bạn thêm khoản cho vay, hệ thống sẽ **TỰ ĐỘNG** tạo một giao dịch Chi tiêu trong sheet **CHI** với danh mục "Cho vay". Điều này giúp dòng tiền (Cash Flow) của bạn giảm đi đúng thực tế.

---

## 💰 Thu hồi nợ & lãi

Khi người vay trả tiền (gốc hoặc lãi), bạn cần ghi nhận vào hệ thống.

### Bước 1: Mở form
Trên thanh menu, chọn **HODLVN Finance** → **🤝 Cho vay** → **💰 Thu nợ & lãi**.

### Bước 2: Điền thông tin
- **Ngày**: Ngày nhận tiền.
- **Người vay**: Chọn từ danh sách (chỉ hiện những người đang nợ).
- **Thu gốc**: Số tiền gốc thu về (nhập 0 nếu chỉ thu lãi).
- **Thu lãi**: Số tiền lãi thu về (nhập 0 nếu chỉ thu gốc).
- **Ghi chú**: Ghi chú thêm (ví dụ: "Trả đợt 1").

### Bước 3: Xác nhận
Nhấn **Xác nhận thu tiền**.

> [!IMPORTANT]
> **Tự động tạo Thu nhập**:
> - Tiền **Gốc** thu về sẽ được ghi nhận là **Thu nhập** (Danh mục: "Thu hồi nợ").
> - Tiền **Lãi** thu về sẽ được ghi nhận là **Thu nhập** (Danh mục: "Lãi đầu tư").

---

## 🔄 Các loại hình cho vay

Hệ thống hỗ trợ 3 hình thức cho vay phổ biến:

### 1. Tất toán gốc - lãi cuối kỳ
- Người vay trả toàn bộ gốc và lãi một lần khi đáo hạn.
- Thường dùng cho các khoản vay ngắn hạn.

### 2. Trả lãi hàng tháng, gốc cuối kỳ
- Người vay trả lãi định kỳ hàng tháng.
- Gốc trả một lần khi đáo hạn.
- Thường dùng cho cho vay lãi ngày, lãi tháng.

### 3. Trả góp gốc - lãi hàng tháng
- Người vay trả một phần gốc và lãi hàng tháng.
- Số tiền trả hàng tháng cố định hoặc giảm dần (tùy thỏa thuận, hệ thống hiện tại chỉ ghi nhận số thực thu).

---

## 💸 Tác động đến Cash Flow

Quy trình dòng tiền hoạt động như sau:

1. **Khi cho vay (Giải ngân)**:
   - Cash Flow **GIẢM** (do phát sinh Chi tiêu "Cho vay").
   - Tài sản "Khoản phải thu" **TĂNG**.

2. **Khi thu nợ (Thu gốc)**:
   - Cash Flow **TĂNG** (do phát sinh Thu nhập "Thu hồi nợ").
   - Tài sản "Khoản phải thu" **GIẢM**.

3. **Khi thu lãi**:
   - Cash Flow **TĂNG** (do phát sinh Thu nhập "Lãi đầu tư").
   - Tài sản ròng (Net Worth) **TĂNG**.

---

[⬅️ Quay lại Hướng dẫn sử dụng](USER_GUIDE.md)
