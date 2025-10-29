# 📖 HƯỚNG DẪN SỬ DỤNG - HODLVN-Family-Finance

Hướng dẫn chi tiết cách sử dụng tất cả tính năng của hệ thống quản lý tài chính gia đình.

---

## 📑 MỤC LỤC

1. [Bắt đầu sử dụng](#-bắt-đầu-sử-dụng)
2. [Giao diện tổng quan](#-giao-diện-tổng-quan)
3. [Quản lý Thu nhập](#-quản-lý-thu-nhập)
4. [Quản lý Chi tiêu](#-quản-lý-chi-tiêu)
5. [Quản lý Nợ](#-quản-lý-nợ)
6. [Đầu tư - Chứng khoán](#-đầu-tư---chứng-khoán)
7. [Đầu tư - Vàng](#-đầu-tư---vàng)
8. [Đầu tư - Cryptocurrency](#-đầu-tư---cryptocurrency)
9. [Đầu tư - Khác](#-đầu-tư---khác)
10. [Cổ tức](#-cổ-tức)
11. [Quản lý Ngân sách](#-quản-lý-ngân-sách)
12. [Dashboard & Báo cáo](#-dashboard--báo-cáo)
13. [Tips & Tricks](#-tips--tricks)
14. [Best Practices](#-best-practices)
15. [Keyboard Shortcuts](#-keyboard-shortcuts)
16. [Tính năng nâng cao](#-tính-năng-nâng-cao)
17. [Câu hỏi thường gặp](#-câu-hỏi-thường-gặp)

---

## 🚀 Bắt đầu sử dụng

### Mở file lần đầu

```
1. Mở Google Drive (drive.google.com)
2. Tìm file "HODLVN-Finance-[Tên bạn]"
3. Double click để mở
4. Đợi 5-10 giây để menu load
```

**Menu "HODLVN Finance" xuất hiện ở thanh menu.**

### Quy trình làm việc hàng ngày

```
Sáng:
1. Mở file
2. Xem Dashboard → Kiểm tra tình hình tài chính
3. Xem Budget → Check chi tiêu so với kế hoạch

Trong ngày:
4. Phát sinh giao dịch → Mở form tương ứng → Nhập ngay
5. Form tự động tính toán và validate

Tối:
6. Review lại các giao dịch trong ngày
7. Check Dashboard xem Cash Flow
8. Đánh giá chi tiêu có hợp lý không
```

---

## 🎨 Giao diện tổng quan

### Dashboard (Tổng quan)

**Các chỉ số quan trọng:**
- 💰 **Tổng Thu**: Tổng thu nhập (bao gồm cả tiền vay)
- 💸 **Tổng Chi**: Tổng chi tiêu theo danh mục
- 📊 **Đầu tư**: Tổng tiền đã đầu tư
- 💳 **Trả nợ**: Tổng nợ gốc + lãi đã trả
- 💵 **Cash Flow**: Thu - (Chi + Trả nợ + Đầu tư)

### 10 Data Sheets

| Sheet | Mục đích |
|-------|----------|
| **THU** | Thu nhập: Lương, thưởng, thu phụ, tiền vay... |
| **CHI** | Chi tiêu: Ăn uống, đi lại, nhà cửa, giải trí... |
| **QUẢN LÝ NỢ** | Danh sách các khoản vay, số tiền, lãi suất... |
| **TRẢ NỢ** | Lịch sử trả nợ: Trả gốc, trả lãi, ngày trả... |
| **CHỨNG KHOÁN** | Giao dịch cổ phiếu: Mua/bán, margin... |
| **VÀNG** | Giao dịch vàng: SJC, PNJ, DOJI... |
| **CRYPTO** | Tiền ảo: BTC, ETH, altcoins... |
| **ĐẦU TƯ KHÁC** | BĐS, quỹ, trái phiếu, tiết kiệm... |
| **CỔ TỨC** | Cổ tức nhận được từ cổ phiếu |
| **BUDGET** | Kế hoạch và theo dõi ngân sách |

---

## 📥 Quản lý Thu nhập

### Mở form

```
Menu "HODLVN Finance" → "📥 Thu (THU)"
```

### Danh mục thu nhập mặc định

- 💼 **Lương**: Lương tháng, lương tuần
- 💰 **Thưởng**: Thưởng tháng, thưởng dự án, bonus
- 📊 **Đầu tư**: Lãi từ đầu tư
- 🎁 **Quà tặng**: Tiền quà, tiền mừng
- 💵 **Thu nhập phụ**: Freelance, bán hàng online, cho thuê
- 🔄 **Khác**: Các khoản thu khác

### Ví dụ: Nhập lương tháng

```
Ngày: 2025-10-30
Danh mục: Lương
Số tiền: 15,000,000
Ghi chú: Lương tháng 10/2025
```

**Kết quả:**
- Sheet THU: Row mới với data
- Dashboard: Tổng thu +15,000,000
- Cash Flow: +15,000,000

---

## 📤 Quản lý Chi tiêu

### Mở form

```
Menu "HODLVN Finance" → "📤 Chi (CHI)"
```

### Danh mục chi tiêu chi tiết

#### 1. 🍜 Ăn uống
Ăn sáng, trưa, tối, cafe, nhà hàng, mua thực phẩm

#### 2. 🚗 Di chuyển
Xăng xe, Grab, taxi, bảo dưỡng xe, phí đỗ xe

#### 3. 🏠 Nhà cửa
Tiền thuê nhà, sửa chữa, đồ dùng gia đình

#### 4. ⚡ Điện nước
Tiền điện, nước, gas

#### 5. 📱 Viễn thông
Điện thoại, internet, Netflix, Spotify

#### 6. 🎓 Giáo dục
Học phí, sách vở, khóa học online

#### 7. 🏥 Y tế
Khám bệnh, mua thuốc, bảo hiểm y tế

#### 8. 👗 Mua sắm
Quần áo, đồ điện tử, mỹ phẩm

#### 9. 🎬 Giải trí
Xem phim, du lịch, game, sách

#### 10. 🔄 Khác

### Tác động đến Budget

**Màu sắc cảnh báo:**
- < 70%: 🟢 Xanh (An toàn)
- 70-90%: 🟡 Vàng (Cảnh báo)
- > 90%: 🔴 Đỏ (Vượt ngân sách)

---

## 💳 Quản lý Nợ

### Form Quản lý Nợ

```
Menu "HODLVN Finance" → "💳 Quản lý nợ"
```

**Loại nợ:**
- 🏦 Vay ngân hàng
- 💳 Thẻ tín dụng
- 👥 Vay cá nhân
- 🏢 Vay công ty
- 📈 Vay margin (chứng khoán)
- 🔄 Khác

### ⚠️ Tính năng quan trọng: Tự động tạo Thu nhập

Khi thêm khoản nợ, hệ thống **TỰ ĐỘNG** tạo khoản thu nhập:

```
Vay ngân hàng: 50,000,000 VNĐ
→ Tự động tạo thu nhập: "Vay Ngân hàng" - 50,000,000 VNĐ
```

**Lý do:** Vay tiền = Tăng cash flow hiện tại

### Form Trả Nợ

```
Menu "HODLVN Finance" → "💰 Trả nợ"
```

**Quy trình:**
1. Chọn khoản nợ từ dropdown
2. Nhập số tiền trả gốc
3. Nhập số tiền trả lãi
4. Submit

**Hệ thống tự động:**
- Cập nhật số dư nợ
- Giảm Cash Flow
- Đánh dấu "Đã thanh toán" nếu trả hết

### Ví dụ: Trả nợ mua nhà

```
Khoản nợ: Vay NH ABC - 500tr - 8.5%
Trả gốc: 2,000,000
Trả lãi: 3,541,667
Tổng: 5,541,667
```

**Tính lãi 1 tháng:**
```
= (500,000,000 × 8.5%) / 12
= 3,541,667 VNĐ
```

---

## 📈 Đầu tư - Chứng khoán

### Mở form

```
Menu "HODLVN Finance" → "📈 Chứng khoán"
```

### Loại giao dịch

- **Mua**: Mua cổ phiếu
- **Bán**: Bán cổ phiếu

### Tính toán Phí & Thuế

**Khi MUA:**
```
Phí = Tổng giá trị × 0.15%
Thuế = 0
Tổng tiền cần = Tổng giá trị + Phí
```

**Khi BÁN:**
```
Phí = Tổng giá trị × 0.15%
Thuế = Tổng giá trị × 0.1%
Tiền nhận về = Tổng giá trị - Phí - Thuế
```

### Ví dụ: Mua cổ phiếu

```
Mã CP: VNM
Số lượng: 1000 CP
Giá: 85,000 VNĐ/CP
─────────────────────
Tổng giá trị: 85,000,000 VNĐ
Phí GD (0.15%): 127,500 VNĐ
Tổng tiền cần: 85,127,500 VNĐ
```

### Margin Trading

**Khi tick "Sử dụng Margin":**
- Tự động tạo khoản nợ trong QUẢN LÝ NỢ
- Loại nợ: "Vay margin"
- Tạo thu nhập tương ứng

---

## 🪙 Đầu tư - Vàng

### Loại vàng & Đơn vị

**Loại vàng:**
- 🥇 Vàng SJC (thanh/lượng)
- 💍 Vàng SJC 9999 (nhẫn)
- 🏪 Vàng PNJ 9999
- 🏆 Vàng DOJI

**Đơn vị:**
```
1 Lượng = 10 Chỉ = 37.5 Gram
1 Chỉ = 3.75 Gram
```

### Ví dụ: Mua vàng SJC

```
Loại vàng: Vàng SJC (thanh/lượng)
Đơn vị: Lượng
Số lượng: 2 lượng
Giá: 75,000,000 VNĐ/lượng
─────────────────────
Tổng giá trị: 150,000,000 VNĐ
```

---

## ₿ Đầu tư - Cryptocurrency

### Coin/Token phổ biến

```
₿ BTC (Bitcoin) - "Vàng kỹ thuật số"
Ξ ETH (Ethereum) - Smart contract platform
💎 BNB (Binance Coin) - Binance ecosystem
₳ ADA (Cardano) - "ETH killer"
◎ SOL (Solana) - Tốc độ cao, phí thấp
✖️ XRP (Ripple) - Thanh toán xuyên biên giới
💵 USDT (Tether) - Stablecoin
```

### Loại giao dịch

1. **Mua**: Mua coin/token
2. **Bán**: Bán coin/token
3. **Swap**: Đổi coin này sang coin khác

### Ví dụ: Mua Bitcoin (DCA)

```
Coin: BTC
Số lượng: 0.01 BTC
Giá: 68,000 USDT/BTC
─────────────────────
Giá trị (USDT): 680 USDT
Giá trị (VNĐ): 16,660,000 VNĐ
Phí (0.1%): 16,660 VNĐ
─────────────────────
Tổng: 16,676,660 VNĐ
```

### Tips an toàn

```
✅ Bật 2FA
✅ Không share Private Key
✅ Dùng Hardware Wallet cho số lượng lớn
✅ Crypto chỉ nên chiếm 5-10% tổng tài sản
✅ Phân bổ: 50% BTC, 30% ETH, 20% Altcoins
```

---

## 💼 Đầu tư - Khác

### Các loại đầu tư

1. 🏠 **Bất động sản**: Đất nền, nhà đất, chung cư
2. 📊 **Quỹ đầu tư**: Open-end, Closed-end, ETF
3. 📜 **Trái phiếu**: Chính phủ, doanh nghiệp, ngân hàng
4. 🏦 **Gửi tiết kiệm**: Kỳ hạn 1-24 tháng
5. 🚀 **Kinh doanh/Startup**: Góp vốn kinh doanh
6. 🤝 **P2P Lending**: Cho vay qua platform

### Ví dụ: Mua quỹ ETF

```
Loại đầu tư: Quỹ đầu tư (Fund)
Tên: Quỹ VCBF VN30 ETF
Số tiền: 50,000,000 VNĐ
Lợi nhuận KV: 12%/năm
Kỳ hạn: 36 tháng
```

---

## 💵 Cổ tức

### Mở form

```
Menu "HODLVN Finance" → "💵 Cổ tức"
```

### Loại cổ tức

1. **Tiền mặt**: Nhận tiền về tài khoản
2. **Cổ phiếu**: Nhận thêm cổ phiếu
3. **Cả hai**: Vừa tiền, vừa CP

### Ví dụ: Cổ tức tiền mặt

```
Mã CP: VNM
Portfolio: 1,000 CP
Cổ tức/CP: 2,000 VNĐ
─────────────────────
Tổng cổ tức: 2,000,000 VNĐ
Loại: Tiền mặt
```

**Dividend Yield:**
```
= (2,000 / 85,000) × 100%
= 2.35%/năm
```

---

## 🎯 Quản lý Ngân sách

### Phương pháp 50/30/20 Rule

```
Thu nhập: 20,000,000 VNĐ

50% (10tr) → Nhu cầu thiết yếu
30% (6tr) → Mong muốn cá nhân
20% (4tr) → Tiết kiệm & Đầu tư
```

### Màu sắc cảnh báo

```
🟢 Xanh:  < 70% ngân sách (An toàn)
🟡 Vàng:  70-90% (Cảnh báo)
🔴 Đỏ:    > 90% (Vượt ngân sách)
```

### Ví dụ: Ngân sách gia đình trẻ (25tr)

```
NGÂN SÁCH CHI TIÊU:
🍜 Ăn uống:        5,000,000  (20%)
🚗 Di chuyển:      2,000,000  (8%)
🏠 Nhà cửa:        6,000,000  (24%)
⚡ Điện nước:        800,000  (3.2%)
📱 Viễn thông:       500,000  (2%)
🎓 Giáo dục:       1,000,000  (4%)
🏥 Y tế:             500,000  (2%)
👗 Mua sắm:        1,500,000  (6%)
🎬 Giải trí:       1,000,000  (4%)
🔄 Khác:             700,000  (2.8%)
━━━━━━━━━━━━━━━━━━━━━
Tổng Chi:         19,000,000  (76%)

MỤC TIÊU ĐẦU TƯ:
📈 Chứng khoán:    2,500,000  (10%)
🪙 Vàng:           1,500,000  (6%)
₿ Crypto:            500,000  (2%)
💼 Tiết kiệm:      1,500,000  (6%)
━━━━━━━━━━━━━━━━━━━━━
Tổng Đầu tư:       6,000,000  (24%)

🎯 TỔNG:          25,000,000  (100%) ✅
```

---

## 📊 Dashboard & Báo cáo

### Dashboard chính

**Hiển thị:**
- Tổng thu/chi/đầu tư/trả nợ
- Cash Flow = Thu - (Chi + Trả nợ + Đầu tư)
- Biểu đồ thu chi theo tháng
- Biểu đồ phân bổ chi tiêu
- Danh mục đầu tư

### Phân tích theo kỳ

- **Tuần**: 7 ngày gần nhất
- **Tháng**: Tháng hiện tại
- **Quý**: 3 tháng
- **Năm**: 12 tháng

### So sánh YoY (Year over Year)

```
Thu nhập:
- Năm 2024: 200tr
- Năm 2025: 250tr
- Tăng: +25%

Chi tiêu:
- Năm 2024: 150tr
- Năm 2025: 170tr
- Tăng: +13.3%
```

---

## 💡 Tips & Tricks

### 1. Nhập liệu nhanh

```
- Dùng Tab để chuyển giữa các field
- Enter để submit form
- Esc để đóng form
```

### 2. Bookmark file

```
Chrome: Ctrl + D
Truy cập nhanh mỗi ngày
```

### 3. Shortcuts Google Sheets

```
Ctrl + F: Tìm kiếm
Ctrl + H: Tìm và thay thế
Ctrl + Z: Undo
Ctrl + Y: Redo
Ctrl + C: Copy
Ctrl + V: Paste
```

### 4. Filter & Sort

```
Data → Create a filter
→ Filter theo ngày, danh mục, số tiền
```

### 5. Export dữ liệu

```
File → Download → Excel (.xlsx)
→ Backup hoặc phân tích ngoài
```

---

## ✅ Best Practices

### 1. Nhập liệu ngay

```
✅ Nhập giao dịch NGAY khi phát sinh
❌ Đừng để tích lũy rồi nhập 1 lúc
→ Dễ quên, sai số liệu
```

### 2. Review hàng tuần

```
Mỗi Chủ Nhật:
- Xem lại các giao dịch trong tuần
- Check Budget có vượt không
- Điều chỉnh kế hoạch tuần sau
```

### 3. Backup định kỳ

```
Mỗi tháng:
File → Make a copy
Đặt tên: "HODLVN-Finance-Backup-2025-10"
```

### 4. Phân loại rõ ràng

```
✅ Ghi chú chi tiết
✅ Chọn đúng danh mục
✅ Consistent naming
```

### 5. Đặt mục tiêu SMART

```
S - Specific: Cụ thể
M - Measurable: Đo lường được
A - Achievable: Khả thi
R - Relevant: Liên quan
T - Time-bound: Có thời hạn

Ví dụ:
❌ "Tiết kiệm nhiều hơn"
✅ "Tiết kiệm 100tr trong 12 tháng"
```

---

## ⌨️ Keyboard Shortcuts

### Shortcuts trong Google Sheets

```
Ctrl + Home: Về cell A1
Ctrl + End: Đến cell cuối cùng
Ctrl + Arrow: Di chuyển nhanh
Ctrl + Shift + Arrow: Select nhiều cells

Ctrl + Space: Select cả cột
Shift + Space: Select cả hàng

Alt + E, D: Delete row
Alt + I, R: Insert row above
```

### Shortcuts trong Forms

```
Tab: Next field
Shift + Tab: Previous field
Enter: Submit
Esc: Close form
```

---

## 🚀 Tính năng nâng cao

### 1. Conditional Formatting

Tự động tô màu dựa trên điều kiện

### 2. Data Validation

Dropdown lists, date pickers

### 3. IMPORTRANGE

Link dữ liệu từ sheet khác

### 4. Query Functions

```
=QUERY(THU!A:D, "SELECT * WHERE B='Lương'")
```

### 5. Pivot Tables

Phân tích dữ liệu nhanh chóng

### 6. Apps Script Custom Functions

Viết hàm JavaScript tùy chỉnh

### 7. Triggers

Tự động hóa các tác vụ

---

## ❓ Câu hỏi thường gặp

### Q1: Làm sao xóa giao dịch đã nhập?

**A:** 
```
1. Vào sheet tương ứng (THU, CHI...)
2. Right-click vào row cần xóa
3. Delete row
4. Dashboard tự động cập nhật
```

### Q2: Có thể sửa giao dịch đã nhập không?

**A:** Có!
```
1. Double-click vào cell cần sửa
2. Chỉnh sửa
3. Enter để lưu
```

### Q3: Làm sao thêm danh mục mới?

**A:**
```
1. Vào sheet CONFIG
2. Thêm tên danh mục vào cột tương ứng
3. Dropdown tự động cập nhật
```

### Q4: Dashboard không cập nhật?

**A:**
```
1. Kiểm tra formulas có bị xóa không
2. Ctrl + R để refresh
3. Hoặc chạy function updateDashboard()
```

### Q5: Có thể dùng trên điện thoại không?

**A:** Có!
```
1. Cài Google Sheets app
2. Mở file
3. Forms tương thích mobile
```

### Q6: Làm sao chia sẻ với vợ/chồng?

**A:**
```
1. Click "Share" button
2. Add email của người đó
3. Chọn quyền "Editor"
4. Send
```

### Q7: Có mất phí không?

**A:** Không! Hoàn toàn miễn phí.

### Q8: Dữ liệu có bị mất không?

**A:** 
```
- Google tự động lưu
- Version history: File → Version history
- Nên backup định kỳ để an toàn
```

### Q9: Có thể xuất ra Excel không?

**A:** Có!
```
File → Download → Microsoft Excel (.xlsx)
```

### Q10: Làm sao reset lại từ đầu?

**A:**
```
1. Delete file cũ
2. Make copy template mới
3. Chạy Setup Wizard
```

---

## 🎓 Kết luận

Chúc bạn quản lý tài chính hiệu quả với HODLVN-Family-Finance!

**Nhớ:**
- ✅ Nhập liệu đều đặn
- ✅ Review thường xuyên
- ✅ Stick to your budget
- ✅ Đầu tư thông minh
- ✅ Tiết kiệm đều đặn

**Liên hệ hỗ trợ:**
- 💬 [Facebook Group](https://www.facebook.com/groups/hodl.vn)
- 📧 [Email support](mailto:contact@tohaitrieu.net)
- 🐛 GitHub Issues

---

<div align="center">

**Chúc bạn thành công trên con đường tài chính! 💰📊**

[⬆ Về đầu trang](#-hướng-dẫn-sử-dụng---hodlvn-family-finance)

</div>