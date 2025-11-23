# 📈 Đầu tư Chứng khoán

Hướng dẫn theo dõi danh mục đầu tư chứng khoán, tính toán lãi/lỗ và quản lý margin.

---

## 📋 Tổng quan

Sheet **CHỨNG KHOÁN** giúp bạn ghi lại các giao dịch mua/bán cổ phiếu. Hệ thống sẽ tự động tính toán phí giao dịch, thuế (khi bán) và giá vốn trung bình.

## 📝 Giao dịch Mua/Bán

### 1. Mở form
Trên thanh menu, chọn **HODLVN Finance** → **📈 Chứng khoán**.

### 2. Loại giao dịch
- **Mua**: Mua thêm cổ phiếu.
- **Bán**: Bán bớt hoặc bán hết cổ phiếu.

### 3. Công thức tính phí & thuế
Hệ thống sử dụng mức phí tham khảo phổ biến trên thị trường:

**Khi MUA:**
- Phí giao dịch: 0.15% giá trị giao dịch.
- Thuế: 0.
- **Tổng tiền chi** = (Giá x Số lượng) + Phí.

**Khi BÁN:**
- Phí giao dịch: 0.15% giá trị giao dịch.
- Thuế thu nhập cá nhân: 0.1% giá trị giao dịch.
- **Tiền thực nhận** = (Giá x Số lượng) - Phí - Thuế.

## 🏦 Sử dụng Margin (Vay ký quỹ)

Nếu bạn sử dụng margin để mua cổ phiếu:
1. Trong form Mua, tích chọn **"Sử dụng Margin"**.
2. Hệ thống sẽ:
   - Ghi nhận giao dịch mua chứng khoán.
   - **Tự động tạo khoản nợ** trong sheet **QUẢN LÝ NỢ** với loại "Vay margin".
   - Tạo thu nhập tương ứng từ khoản vay này.

## 💡 Ví dụ: Mua cổ phiếu VNM

**Thông tin giao dịch:**
- Mã CP: **VNM**
- Số lượng: **1,000** CP
- Giá khớp: **85,000** VNĐ/CP

**Tính toán tự động:**
- Giá trị lệnh: 85,000,000 VNĐ
- Phí GD (0.15%): 127,500 VNĐ
- **Tổng tiền cần có**: 85,127,500 VNĐ

---

[⬅️ Quay lại Hướng dẫn sử dụng](USER_GUIDE.md)
