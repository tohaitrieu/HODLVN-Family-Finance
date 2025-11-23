# 💳 Quản lý Nợ

Hướng dẫn quản lý các khoản vay nợ và lịch trình trả nợ.

---

## 📋 Tổng quan

Tính năng này giúp bạn theo dõi các khoản bạn đi vay (từ ngân hàng, người thân, thẻ tín dụng...) và quá trình trả nợ dần.

Hệ thống sử dụng 2 sheet:
- **QUẢN LÝ NỢ**: Danh sách các khoản nợ hiện tại.
- **TRẢ NỢ**: Lịch sử các lần trả gốc và lãi.

## ➕ Thêm khoản nợ mới

### 1. Mở form
Trên thanh menu, chọn **HODLVN Finance** → **💳 Quản lý nợ**.

### 2. Các loại nợ hỗ trợ
- 🏦 **Vay ngân hàng**: Vay mua nhà, mua xe, vay tiêu dùng.
- 💳 **Thẻ tín dụng**: Dư nợ thẻ tín dụng cần trả.
- 👥 **Vay cá nhân**: Vay bạn bè, người thân.
- 🏢 **Vay công ty**: Tạm ứng lương, vay ưu đãi công ty.
- 📈 **Vay margin**: Vay ký quỹ chứng khoán (thường được tạo tự động từ module Chứng khoán).
- 🔄 **Khác**: Các khoản nợ khác.

### 3. Tự động tạo Thu nhập
> [!IMPORTANT]
> Khi bạn thêm một khoản nợ mới, hệ thống sẽ **TỰ ĐỘNG** tạo một giao dịch **Thu nhập** trong sheet **THU**.
>
> **Ví dụ:** Bạn vay ngân hàng 500 triệu.
> - Sheet **QUẢN LÝ NỢ**: Ghi nhận khoản nợ 500 triệu.
> - Sheet **THU**: Ghi nhận thu nhập 500 triệu (Nguồn: "Vay Ngân hàng").
>
> **Lý do:** Vay tiền làm tăng lượng tiền mặt hiện có (Cash Flow dương).

## 💰 Trả nợ

Khi bạn thực hiện trả nợ (trả góp hàng tháng hoặc tất toán), hãy ghi nhận vào hệ thống.

### 1. Mở form
Trên thanh menu, chọn **HODLVN Finance** → **💰 Trả nợ**.

### 2. Quy trình nhập liệu
1. **Chọn khoản nợ**: Chọn từ danh sách các khoản nợ đang có.
2. **Nhập tiền trả gốc**: Số tiền gốc bạn trả đợt này.
3. **Nhập tiền trả lãi**: Số tiền lãi bạn trả đợt này.
4. **Submit**.

### 3. Hệ thống xử lý
- **Cập nhật số dư**: Trừ số tiền gốc vừa trả khỏi tổng nợ gốc.
- **Ghi nhận chi tiêu**: Khoản trả nợ được tính vào dòng tiền ra (Cash Flow âm).
- **Cập nhật trạng thái**: Nếu dư nợ về 0, khoản nợ sẽ được đánh dấu là "Đã tất toán".

## 💡 Ví dụ thực tế: Trả nợ mua nhà

**Thông tin khoản vay:**
- Vay Ngân hàng ABC
- Số tiền: 500,000,000 VNĐ
- Lãi suất: 8.5%/năm

**Thực hiện trả nợ tháng 1:**
- Trả gốc: 2,000,000 VNĐ
- Trả lãi: 3,541,667 VNĐ (=(500tr * 8.5%)/12)
- **Tổng chi**: 5,541,667 VNĐ

---

[⬅️ Quay lại Hướng dẫn sử dụng](USER_GUIDE.md)
