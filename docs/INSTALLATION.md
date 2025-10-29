# 📥 HƯỚNG DẪN CÀI ĐẶT - HODLVN-Family-Finance

Hướng dẫn cài đặt chi tiết hệ thống quản lý tài chính gia đình từ A đến Z.

---

## 📑 MỤC LỤC

1. [Tổng quan](#-tổng-quan)
2. [Yêu cầu trước khi cài đặt](#-yêu-cầu-trước-khi-cài-đặt)
3. [Phương pháp cài đặt](#-phương-pháp-cài-đặt)
   - [Phương pháp 1: Sao chép Template (Đơn giản)](#phương-pháp-1-sao-chép-template-đơn-giản-khuyến-nghị)
   - [Phương pháp 2: Cài đặt từ đầu (Nâng cao)](#phương-pháp-2-cài-đặt-từ-đầu-nâng-cao)
4. [Cấu hình Google Apps Script](#-cấu-hình-google-apps-script)
5. [Chạy Setup Wizard](#-chạy-setup-wizard)
6. [Cấu hình nâng cao](#-cấu-hình-nâng-cao)
7. [Kiểm tra và xác minh](#-kiểm-tra-và-xác-minh)
8. [Xử lý lỗi thường gặp](#-xử-lý-lỗi-thường-gặp)
9. [Video hướng dẫn](#-video-hướng-dẫn)
10. [Checklist cài đặt](#-checklist-cài-đặt)
11. [Câu hỏi thường gặp](#-câu-hỏi-thường-gặp)

---

## 🎯 Tổng quan

### Thời gian cài đặt

| Phương pháp | Thời gian | Độ khó | Phù hợp với |
|-------------|-----------|--------|-------------|
| **Template** | 5-10 phút | ⭐ Dễ | Người dùng thông thường |
| **Từ đầu** | 30-45 phút | ⭐⭐⭐ Khó | Developers, người muốn tùy chỉnh |

### Quy trình tổng quan

```
1. Chuẩn bị → 2. Tạo Google Sheet → 3. Cài Apps Script → 4. Setup Wizard → 5. Hoàn thành
   (2 phút)      (3 phút)              (5 phút)           (3 phút)         (2 phút)
```

---

## ✅ Yêu cầu trước khi cài đặt

### 1. Tài khoản Google

**Bắt buộc:**
- ✅ Tài khoản Gmail hoạt động
- ✅ Đã đăng nhập vào Google Drive
- ✅ Có quyền tạo và chỉnh sửa Google Sheets

**Khuyến nghị:**
- 📱 Bật xác thực 2 bước (2FA) để bảo mật
- 💾 Đảm bảo còn ít nhất 100MB dung lượng Drive
- 🔒 Xem xét sử dụng tài khoản riêng cho tài chính

### 2. Trình duyệt

**Được hỗ trợ:**
- ✅ Google Chrome (phiên bản 90+)
- ✅ Microsoft Edge (phiên bản 90+)
- ✅ Mozilla Firefox (phiên bản 88+)
- ✅ Safari (phiên bản 14+)

**Không được hỗ trợ:**
- ❌ Internet Explorer
- ❌ Trình duyệt cũ (< 2020)

### 3. Kết nối Internet

- 🌐 Tốc độ tối thiểu: 2 Mbps
- 📶 Kết nối ổn định trong quá trình cài đặt
- ⏱️ Latency thấp để tương tác mượt mà

### 4. Kiến thức cơ bản

**Cho người dùng thông thường (Phương pháp 1):**
- 📊 Biết cách sử dụng Google Sheets cơ bản
- 🖱️ Copy/Paste, điền form

**Cho Developers (Phương pháp 2):**
- 💻 Hiểu cơ bản về Google Apps Script
- 📝 Biết JavaScript
- 🔧 Có kinh nghiệm debug code

---

## 🚀 Phương pháp cài đặt

---

## Phương pháp 1: Sao chép Template (Đơn giản) - KHUYẾN NGHỊ

### Bước 1: Truy cập Template

#### 1.1. Mở link Template

```
👉 Link Template: https://docs.google.com/spreadsheets/d/YOUR_TEMPLATE_ID
(Link sẽ được cập nhật khi release)
```

**Lưu ý:** Template đã có sẵn:
- ✅ Toàn bộ code Apps Script
- ✅ 10 sheets được cấu hình
- ✅ Formulas đã được thiết lập
- ✅ Forms HTML sẵn sàng

#### 1.2. Xem quyền truy cập

Đảm bảo bạn thấy thông báo:
```
"Chế độ xem" hoặc "View only"
```

Nếu gặp lỗi "Access denied":
- 🔍 Kiểm tra link có chính xác không
- 🔐 Đảm bảo đã đăng nhập đúng tài khoản Gmail

---

### Bước 2: Tạo bản sao

#### 2.1. Sao chép file

1. Click **File** (góc trên bên trái)
2. Chọn **Make a copy** (hoặc **Tạo bản sao**)

![Make a copy](images/installation/make-a-copy.png)

#### 2.2. Đặt tên file

```
Tên đề xuất: HODLVN-Finance-[Tên bạn]-2025
Ví dụ: HODLVN-Finance-NikaFamily-2025
```

**Lưu ý đặt tên:**
- ✅ Có thể chứa tiếng Việt có dấu
- ✅ Nên thêm năm để dễ quản lý
- ✅ Tránh ký tự đặc biệt: `/ \ : * ? " < > |`

#### 2.3. Chọn vị trí lưu

```
📁 Đề xuất: 
   My Drive
   └── Tài chính gia đình
       └── HODLVN-Finance-NikaFamily-2025
```

**Tạo folder mới (nếu chưa có):**
1. Click "Organize" hoặc "Folder" icon
2. New folder → "Tài chính gia đình"
3. Select folder vừa tạo

#### 2.4. Hoàn tất sao chép

1. Click **Make a copy** (màu xanh)
2. Đợi 5-10 giây
3. File mới sẽ tự động mở

✅ **Bạn đã có bản sao riêng!**

---

### Bước 3: Kích hoạt Scripts

#### 3.1. Chạy function đầu tiên

Khi mở file lần đầu, bạn sẽ thấy:

```
⚠️ Authorization Required
```

**Thực hiện:**

1. Đợi menu "HODLVN Finance" xuất hiện (10-20 giây)
2. Nếu menu không xuất hiện:
   - Extensions → Apps Script
   - Chọn function `onOpen`
   - Click ▶️ Run

![Run onOpen](images/installation/run-onopen.png)

#### 3.2. Cấp quyền truy cập

Sẽ có popup:

```
⚠️ Authorization needed
This project wants to:
- View and manage your spreadsheets
- Display and run content from the web
```

**Các bước cấp quyền:**

1. Click **Review permissions** (Xem lại quyền)
2. Chọn tài khoản Google của bạn
3. Nếu thấy cảnh báo:

```
⚠️ Google hasn't verified this app
```

**Không sao! Đây là app của bạn.**

4. Click **Advanced** (Nâng cao)
5. Click **Go to [Project Name] (unsafe)**
6. Click **Allow** (Cho phép)

![Authorization steps](images/installation/authorization.png)

**Giải thích:**
- 🔒 Apps Script chưa verify vì là project cá nhân
- ✅ An toàn 100% vì code do bạn kiểm soát
- 🔐 Chỉ bạn có quyền truy cập

#### 3.3. Kiểm tra menu

Sau khi cấp quyền, refresh trang (F5):

```
Menu "HODLVN Finance" xuất hiện ✅
```

Menu bao gồm:
- 🚀 Setup Wizard
- 📥 Thu (THU)
- 📤 Chi (CHI)
- 💳 Quản lý nợ
- 💰 Trả nợ
- 📈 Chứng khoán
- 🪙 Vàng
- ₿ Crypto
- 💼 Đầu tư khác
- 💵 Cổ tức
- 🎯 Đặt ngân sách

---

### Bước 4: Chạy Setup Wizard

#### 4.1. Mở Setup Wizard

```
Menu "HODLVN Finance" → "🚀 Setup Wizard"
```

Một cửa sổ popup sẽ hiện lên.

![Setup Wizard](images/installation/setup-wizard.png)

#### 4.2. Điền thông tin

**Thông tin bắt buộc:**

```
1. Tên chủ hộ: _______________
   Ví dụ: Nguyễn Văn A

2. Email: _______________
   Ví dụ: nguyenvana@gmail.com

3. Số điện thoại: _______________
   Ví dụ: 0912345678

4. Năm bắt đầu: _______________
   Ví dụ: 2025
```

**Thông tin tùy chọn:**

```
5. Ngân sách hàng tháng (VNĐ): _______________
   Ví dụ: 20,000,000
   (Có thể để trống, sẽ cài sau)

6. Mục tiêu tài chính: _______________
   Ví dụ: Tiết kiệm mua nhà
   (Tùy chọn)
```

#### 4.3. Khởi tạo hệ thống

1. Click **"Khởi tạo hệ thống"**
2. Đợi 10-15 giây
3. Thông báo thành công:

```
✅ Khởi tạo thành công!
Hệ thống đã sẵn sàng sử dụng.
```

#### 4.4. Xác minh

Kiểm tra các tabs:
- ✅ Dashboard có tên bạn
- ✅ 10 sheets đã được tạo
- ✅ Budget sheet có cấu trúc

---

### Bước 5: Hoàn thành 🎉

**Xin chúc mừng! Bạn đã cài đặt thành công.**

**Bước tiếp theo:**
1. 📖 Đọc [USER_GUIDE.md](USER_GUIDE.md) để học cách sử dụng
2. 💰 Nhập giao dịch đầu tiên
3. 📊 Xem Dashboard

---

## Phương pháp 2: Cài đặt từ đầu (Nâng cao)

### Bước 1: Tạo Google Sheet mới

#### 1.1. Tạo spreadsheet trống

```
1. Truy cập: https://sheets.google.com
2. Click "+" (Blank spreadsheet)
3. Đặt tên: HODLVN-Finance-[Tên bạn]-2025
```

#### 1.2. Tạo sheet đầu tiên

Mặc định sẽ có sheet "Sheet1"
- Đổi tên thành: **"Dashboard"**

---

### Bước 2: Mở Apps Script Editor

#### 2.1. Truy cập Script Editor

```
Extensions → Apps Script
```

Hoặc dùng phím tắt:
```
Alt + Shift + 11 (Windows)
Cmd + Option + Shift + 11 (Mac)
```

#### 2.2. Giao diện Apps Script

Bạn sẽ thấy:
```javascript
function myFunction() {
  
}
```

**Xóa code mẫu này đi.**

---

### Bước 3: Copy mã nguồn

#### 3.1. Tải source code

**Option A: Từ GitHub**
```
git clone https://github.com/tohaitrieu/HODLVN-Family-Finance.git
```

**Option B: Tải ZIP**
```
https://github.com/tohaitrieu/HODLVN-Family-Finance/archive/refs/heads/main.zip
```

#### 3.2. Tạo các file .gs

Trong Apps Script Editor:

**File 1: Main.gs**
```
1. Click "+" → Script
2. Đặt tên: Main
3. Copy nội dung từ /src/Main.gs
4. Paste vào editor
5. Ctrl + S (Save)
```

**Lặp lại cho các files:**
- SheetInitializer.gs
- DataProcessor.gs
- BudgetManager.gs
- DashboardManager.gs
- Utils.gs

![Apps Script Files](images/installation/apps-script-files.png)

#### 3.3. Tạo các file .html

**File 1: IncomeForm.html**
```
1. Click "+" → HTML
2. Đặt tên: IncomeForm
3. Copy nội dung từ /forms/IncomeForm.html
4. Paste vào editor
5. Ctrl + S (Save)
```

**Lặp lại cho các HTML files:**
- ExpenseForm.html
- DebtManagementForm.html
- DebtPaymentForm.html
- StockForm.html
- GoldForm.html
- CryptoForm.html
- OtherInvestmentForm.html
- DividendForm.html
- SetupWizard.html

---

### Bước 4: Cấu hình Project

#### 4.1. Project settings

```
1. Click ⚙️ (Settings) bên trái
2. Cấu hình:
   - Project name: HODLVN-Finance
   - Time zone: (GMT+07:00) Bangkok, Hanoi, Jakarta
   - Show "appsscript.json" manifest file: ON
```

#### 4.2. Cấu hình appsscript.json

```json
{
  "timeZone": "Asia/Ho_Chi_Minh",
  "dependencies": {
  },
  "exceptionLogging": "STACKDRIVER",
  "runtimeVersion": "V8",
  "oauthScopes": [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/script.container.ui"
  ]
}
```

---

### Bước 5: Chạy initial setup

#### 5.1. Chạy onOpen function

```
1. Chọn function: onOpen
2. Click ▶️ Run
3. Cấp quyền (như Phương pháp 1, Bước 3.2)
```

#### 5.2. Quay lại Sheet và refresh

```
F5 hoặc Ctrl + R
```

Menu "HODLVN Finance" xuất hiện ✅

---

### Bước 6: Chạy Setup Wizard

Làm tương tự **Phương pháp 1, Bước 4**.

---

## 🔧 Cấu hình Google Apps Script

### Cấu hình nâng cao cho Developers

#### 1. Script Properties

Lưu trữ các biến toàn cục:

```javascript
// Truy cập Script Properties
var scriptProperties = PropertiesService.getScriptProperties();

// Set properties
scriptProperties.setProperty('SHEET_VERSION', '3.0');
scriptProperties.setProperty('SETUP_COMPLETED', 'true');
scriptProperties.setProperty('USER_NAME', 'Nika');

// Get properties
var version = scriptProperties.getProperty('SHEET_VERSION');
```

**Cách set qua GUI:**
```
1. Apps Script Editor
2. Project Settings
3. Script Properties
4. Add Property
```

#### 2. Triggers (Tự động hóa)

**Trigger onOpen tự động:**
```javascript
function createOnOpenTrigger() {
  ScriptApp.newTrigger('onOpen')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onOpen()
    .create();
}
```

**Trigger theo thời gian:**
```javascript
// Tự động backup mỗi ngày 2h sáng
function createDailyBackup() {
  ScriptApp.newTrigger('autoBackup')
    .timeBased()
    .atHour(2)
    .everyDays(1)
    .create();
}
```

**Cách tạo trigger qua GUI:**
```
1. Apps Script Editor
2. Click ⏰ (Triggers) bên trái
3. Add Trigger
4. Chọn:
   - Function: onOpen
   - Event source: From spreadsheet
   - Event type: On open
5. Save
```

#### 3. Libraries (Thư viện)

Nếu muốn dùng thư viện external:

```
1. Apps Script Editor
2. Libraries (+)
3. Add library ID
4. Chọn version
5. Save
```

**Ví dụ thư viện phổ biến:**
- Moment.js cho xử lý date
- Lodash cho utilities
- Chart.js cho biểu đồ

---

## 🎨 Chạy Setup Wizard

### Chi tiết các bước Setup Wizard

#### 1. Form thông tin cơ bản

**Các trường dữ liệu:**

| Trường | Bắt buộc | Định dạng | Ví dụ |
|--------|----------|-----------|-------|
| Tên chủ hộ | ✅ | Text | Nguyễn Văn A |
| Email | ✅ | Email | nguyenvana@gmail.com |
| Số điện thoại | ✅ | 10-11 số | 0912345678 |
| Năm bắt đầu | ✅ | YYYY | 2025 |
| Ngân sách tháng | ❌ | Số | 20000000 |
| Mục tiêu | ❌ | Text | Tiết kiệm mua nhà |

**Validation tự động:**
```javascript
// Email validation
if (!email.includes('@')) {
  alert('Email không hợp lệ!');
}

// Phone validation
if (phone.length < 10 || phone.length > 11) {
  alert('Số điện thoại phải 10-11 số!');
}

// Year validation
if (year < 2020 || year > 2030) {
  alert('Năm không hợp lệ!');
}
```

#### 2. Danh mục Thu nhập

Setup Wizard sẽ tạo sẵn các danh mục:

**Danh mục mặc định:**
- 💼 Lương
- 💰 Thưởng
- 📊 Đầu tư
- 🎁 Quà tặng
- 💵 Thu nhập phụ
- 🔄 Khác

**Tùy chỉnh:**
```
- Thêm danh mục: Click "+" 
- Xóa danh mục: Click "🗑️"
- Sửa tên: Double click vào tên
```

#### 3. Danh mục Chi tiêu

**Danh mục mặc định:**
- 🍜 Ăn uống
- 🚗 Di chuyển
- 🏠 Nhà cửa
- ⚡ Điện nước
- 📱 Viễn thông
- 🎓 Giáo dục
- 🏥 Y tế
- 👗 Mua sắm
- 🎬 Giải trí
- 🔄 Khác

#### 4. Loại nợ

**Danh mục mặc định:**
- 🏦 Vay ngân hàng
- 💳 Thẻ tín dụng
- 👥 Vay cá nhân
- 🏢 Vay công ty
- 📈 Vay margin (chứng khoán)
- 🔄 Khác

#### 5. Khởi tạo sheets

Wizard tự động tạo:

```
✅ Dashboard (Tổng quan)
✅ THU (Thu nhập)
✅ CHI (Chi tiêu)
✅ QUẢN LÝ NỢ (Danh sách nợ)
✅ TRẢ NỢ (Lịch sử trả nợ)
✅ CHỨNG KHOÁN (Giao dịch cổ phiếu)
✅ VÀNG (Giao dịch vàng)
✅ CRYPTO (Tiền ảo)
✅ ĐẦU TƯ KHÁC (Đầu tư khác)
✅ CỔ TỨC (Cổ tức)
✅ BUDGET (Ngân sách)
```

**Cấu trúc mỗi sheet:**
- Headers (Tiêu đề cột)
- Sample formulas (Công thức mẫu đến row 1000)
- Data validation (Dropdown lists)
- Conditional formatting (Màu sắc tự động)
- Frozen rows/columns (Đóng băng)

#### 6. Hoàn thành

Thông báo:
```
✅ Khởi tạo thành công!

Các sheet đã được tạo:
- Dashboard ✅
- THU ✅
- CHI ✅
- QUẢN LÝ NỢ ✅
- TRẢ NỢ ✅
- CHỨNG KHOÁN ✅
- VÀNG ✅
- CRYPTO ✅
- ĐẦU TƯ KHÁC ✅
- CỔ TỨC ✅
- BUDGET ✅

Hệ thống đã sẵn sàng sử dụng!
```

---

## ⚙️ Cấu hình nâng cao

### 1. Tùy chỉnh danh mục

#### Thêm danh mục thu nhập mới

```
1. Vào sheet "CONFIG" (ẩn)
2. Tìm bảng "Danh mục thu nhập"
3. Thêm row mới với tên danh mục
4. Sheet THU sẽ tự động cập nhật dropdown
```

**Hoặc dùng code:**
```javascript
function addIncomeCategory(categoryName) {
  var configSheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName('CONFIG');
  
  var lastRow = configSheet.getLastRow();
  configSheet.getRange(lastRow + 1, 1).setValue(categoryName);
  
  Logger.log('Đã thêm danh mục: ' + categoryName);
}
```

#### Thêm danh mục chi tiêu mới

Tương tự như thu nhập, nhưng column khác trong CONFIG sheet.

### 2. Cấu hình màu sắc

#### Thay đổi theme màu

```javascript
// Trong Utils.gs
var COLORS = {
  primary: '#4285F4',      // Màu chính
  success: '#34A853',      // Màu thành công (xanh lá)
  warning: '#FBBC04',      // Màu cảnh báo (vàng)
  danger: '#EA4335',       // Màu nguy hiểm (đỏ)
  income: '#34A853',       // Màu thu nhập
  expense: '#EA4335',      // Màu chi tiêu
  investment: '#4285F4',   // Màu đầu tư
  debt: '#FBBC04'          // Màu nợ
};
```

### 3. Cấu hình công thức

#### Tùy chỉnh công thức tính Cash Flow

Mặc định:
```
= Thu - (Chi + Trả nợ + Đầu tư)
```

Nếu muốn thay đổi:
```
1. Vào sheet CONFIG
2. Tìm cell "CASH_FLOW_FORMULA"
3. Chỉnh sửa công thức
4. Dashboard sẽ tự động cập nhật
```

### 4. Backup tự động

#### Cấu hình auto-backup

```javascript
function setupAutoBackup() {
  // Xóa trigger cũ (nếu có)
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'autoBackup') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // Tạo trigger mới: backup mỗi ngày 2h sáng
  ScriptApp.newTrigger('autoBackup')
    .timeBased()
    .atHour(2)
    .everyDays(1)
    .create();
    
  Logger.log('Đã cấu hình auto-backup');
}

function autoBackup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var file = DriveApp.getFileById(ss.getId());
  var folder = DriveApp.getFoldersByName('HODLVN-Backups').next();
  
  var today = Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd');
  var backupName = 'HODLVN-Backup-' + today;
  
  file.makeCopy(backupName, folder);
  Logger.log('Backup thành công: ' + backupName);
}
```

---

## ✔️ Kiểm tra và xác minh

### Checklist sau khi cài đặt

#### 1. Kiểm tra Menu

```
✅ Menu "HODLVN Finance" hiển thị
✅ Có đầy đủ 11 menu items
✅ Click menu không bị lỗi
```

#### 2. Kiểm tra Sheets

```
✅ Có 11 sheets (Dashboard + 10 data sheets)
✅ Headers đúng format
✅ Formulas hiển thị chính xác
✅ Dropdowns hoạt động
```

#### 3. Test nhập liệu

**Test form Thu nhập:**
```
1. Menu → Thu (THU)
2. Điền form:
   - Ngày: Hôm nay
   - Danh mục: Lương
   - Số tiền: 10,000,000
   - Ghi chú: Test
3. Submit
4. Kiểm tra sheet THU → Row 2 có data ✅
5. Kiểm tra Dashboard → Tổng thu cập nhật ✅
```

**Test form Chi tiêu:**
```
1. Menu → Chi (CHI)
2. Điền form:
   - Ngày: Hôm nay
   - Danh mục: Ăn uống
   - Số tiền: 200,000
   - Ghi chú: Test
3. Submit
4. Kiểm tra sheet CHI → Row 2 có data ✅
5. Kiểm tra Dashboard → Tổng chi cập nhật ✅
6. Kiểm tra Budget → Ăn uống cập nhật ✅
```

#### 4. Kiểm tra Dashboard

```
✅ Tên chủ hộ hiển thị đúng
✅ Tổng thu/chi/đầu tư hiển thị
✅ Cash Flow tính toán chính xác
✅ Biểu đồ hiển thị (nếu có data)
```

#### 5. Kiểm tra Budget

```
✅ Các danh mục chi tiêu hiển thị
✅ Mục tiêu ngân sách đã set
✅ % sử dụng tính toán đúng
✅ Màu sắc cảnh báo hoạt động
```

### Test Cases chi tiết

#### TC01: Nhập thu nhập và kiểm tra Cash Flow

```
Input:
- Thu nhập: 15,000,000 VNĐ

Expected:
- Sheet THU có 1 record
- Dashboard: Tổng thu = 15,000,000
- Dashboard: Cash Flow = 15,000,000
```

#### TC02: Nhập chi tiêu và kiểm tra Budget

```
Input:
- Chi tiêu Ăn uống: 5,000,000 VNĐ
- Budget Ăn uống: 10,000,000 VNĐ

Expected:
- Sheet CHI có 1 record
- Budget: Ăn uống = 5,000,000 / 10,000,000 (50%)
- Budget: Màu xanh (< 70%)
```

#### TC03: Quản lý nợ tự động tạo thu nhập

```
Input:
- Vay ngân hàng: 50,000,000 VNĐ

Expected:
- Sheet QUẢN LÝ NỢ có 1 record
- Sheet THU tự động tạo record: "Vay Ngân hàng"
- Dashboard: Tổng thu tăng 50,000,000
```

#### TC04: Trả nợ và cập nhật Cash Flow

```
Input:
- Trả nợ gốc: 5,000,000 VNĐ
- Trả lãi: 500,000 VNĐ

Expected:
- Sheet TRẢ NỢ có 1 record
- Dashboard: Cash Flow giảm 5,500,000
- QUẢN LÝ NỢ: Số dư nợ giảm 5,000,000
```

---

## ⚠️ Xử lý lỗi thường gặp

### Lỗi 1: Menu không hiển thị

**Triệu chứng:**
```
Menu "HODLVN Finance" không xuất hiện sau khi mở sheet
```

**Nguyên nhân:**
- Script chưa được authorize
- Function onOpen chưa chạy
- Code bị lỗi syntax

**Cách khắc phục:**

```
Cách 1: Refresh page
1. Ctrl + R hoặc F5
2. Đợi 10 giây

Cách 2: Chạy onOpen manual
1. Extensions → Apps Script
2. Chọn function: onOpen
3. Click ▶️ Run
4. Cấp quyền nếu được yêu cầu

Cách 3: Check logs
1. Apps Script Editor
2. View → Logs
3. Xem có lỗi gì không
```

### Lỗi 2: Authorization failed

**Triệu chứng:**
```
Exception: Authorization is required to perform that action
```

**Cách khắc phục:**

```
1. Apps Script Editor
2. Chọn function bất kỳ
3. Click ▶️ Run
4. Popup "Review permissions"
5. Advanced → Go to [Project] (unsafe)
6. Allow
7. Quay lại Sheet và refresh
```

### Lỗi 3: Form submit bị lỗi

**Triệu chứng:**
```
Error: Cannot read property 'getSheetByName' of null
```

**Nguyên nhân:**
- Sheet chưa được tạo
- Tên sheet sai
- Permission bị lỗi

**Cách khắc phục:**

```
1. Kiểm tra sheet có tồn tại không
2. Kiểm tra tên sheet chính xác (có dấu, hoa thường)
3. Chạy lại Setup Wizard
4. Nếu vẫn lỗi, xem Logs trong Apps Script
```

### Lỗi 4: Data không vào đúng vị trí

**Triệu chứng:**
```
Data xuất hiện ở row 1001 thay vì row tiếp theo
```

**Nguyên nhân:**
- Lỗi appendRow với pre-filled formulas
- Function findEmptyRow không hoạt động

**Cách khắc phục:**

```
Đã fix trong v3.0, nếu vẫn gặp:

1. Xóa formulas dư thừa (row > last data row)
2. Hoặc update code DataProcessor.gs với version mới nhất
```

### Lỗi 5: Dashboard không cập nhật

**Triệu chứng:**
```
Nhập dữ liệu nhưng Dashboard vẫn là 0
```

**Nguyên nhân:**
- Formulas trong Dashboard bị xóa
- Range reference sai
- Sheets bị rename

**Cách khắc phục:**

```
1. Kiểm tra formulas trong Dashboard
2. Đảm bảo sheet names chính xác
3. Re-run function updateDashboard() manual:
   - Apps Script Editor
   - Run updateDashboard()
4. Hoặc chạy lại Setup Wizard
```

### Lỗi 6: Budget không có màu

**Triệu chứng:**
```
Budget sheet không có màu cảnh báo (xanh/vàng/đỏ)
```

**Nguyên nhân:**
- Conditional formatting chưa được set
- Data validation bị lỗi

**Cách khắc phục:**

```
1. Select range Budget (cột %)
2. Format → Conditional formatting
3. Set rules:
   - < 70%: Green
   - 70-90%: Yellow
   - > 90%: Red
4. Done
```

### Lỗi 7: Dropdown không có options

**Triệu chứng:**
```
Dropdown danh mục trống, không có gì để chọn
```

**Nguyên nhân:**
- Data validation chưa được cấu hình
- CONFIG sheet bị xóa

**Cách khắc phục:**

```
1. Chạy lại Setup Wizard
2. Hoặc manual add data validation:
   - Select cell
   - Data → Data validation
   - Criteria: List from range
   - Range: CONFIG!A2:A10
   - Save
```

### Lỗi 8: Slow performance

**Triệu chứng:**
```
Sheet load chậm, form submit lâu
```

**Nguyên nhân:**
- Quá nhiều formulas
- File size quá lớn
- Conditional formatting phức tạp

**Cách khắc phục:**

```
1. Xóa formulas row không dùng (> last data)
2. Archive data cũ sang file khác
3. Giảm conditional formatting rules
4. Optimize formulas (SUMIFS thay vì SUMIF nested)
```

---

## 🎥 Video hướng dẫn

### Video chính thức

| Video | Nội dung | Thời lượng | Link |
|-------|----------|------------|------|
| 📺 **Cài đặt cơ bản** | Hướng dẫn cài đặt từ template | 10 phút | [Xem video](#) |
| 📺 **Setup Wizard** | Chi tiết cấu hình ban đầu | 8 phút | [Xem video](#) |
| 📺 **Nhập giao dịch** | Cách sử dụng forms | 12 phút | [Xem video](#) |
| 📺 **Dashboard** | Đọc hiểu báo cáo tài chính | 15 phút | [Xem video](#) |
| 📺 **Troubleshooting** | Xử lý lỗi thường gặp | 10 phút | [Xem video](#) |

### Playlist YouTube

```
🎬 HODLVN-Finance Tutorial Series
https://youtube.com/playlist?list=YOUR_PLAYLIST_ID

- Video 1: Giới thiệu tổng quan
- Video 2: Cài đặt chi tiết
- Video 3: Sử dụng cơ bản
- Video 4: Tính năng nâng cao
- Video 5: Tips & Tricks
```

---

## ✅ Checklist cài đặt

### Pre-installation checklist

```
□ Có tài khoản Gmail
□ Đăng nhập Google Drive
□ Trình duyệt Chrome/Edge mới nhất
□ Kết nối internet ổn định
□ Đã đọc tài liệu cài đặt
```

### Installation checklist (Template method)

```
□ Mở link Template
□ Make a copy
□ Đặt tên file
□ Chọn vị trí lưu
□ File copy thành công
□ Menu HODLVN Finance hiển thị
□ Chạy Setup Wizard
□ Điền thông tin đầy đủ
□ Khởi tạo thành công
□ Kiểm tra 11 sheets đã tạo
```

### Post-installation checklist

```
□ Test form Thu nhập
□ Test form Chi tiêu
□ Kiểm tra Dashboard cập nhật
□ Kiểm tra Budget hoạt động
□ Test ít nhất 1 loại đầu tư
□ Kiểm tra Cash Flow tính toán
□ Set up danh mục riêng (nếu cần)
□ Đọc User Guide
□ Bookmark file để truy cập nhanh
□ Backup file (optional)
```

### Advanced configuration checklist

```
□ Tùy chỉnh danh mục
□ Cấu hình màu sắc
□ Set up auto-backup (nếu cần)
□ Thêm custom formulas (nếu cần)
□ Cấu hình triggers
□ Set up notifications (nếu cần)
```

---

## ❓ Câu hỏi thường gặp

### Q1: Có mất phí không?

**A:** Không, hoàn toàn miễn phí!
- ✅ Google Sheets miễn phí
- ✅ Google Apps Script miễn phí
- ✅ Code mã nguồn mở miễn phí

### Q2: Dữ liệu có bảo mật không?

**A:** Có, rất bảo mật!
- 🔒 Lưu trữ trên Google Drive cá nhân
- 🔐 Chỉ bạn có quyền truy cập
- 🔒 Google có mã hóa end-to-end
- 🔐 Không ai khác thấy được data

### Q3: Có thể dùng trên điện thoại không?

**A:** Có!
- 📱 Cài Google Sheets app
- 📱 Mở file trên app
- 📱 Forms tương thích mobile
- 📱 Dashboard responsive

### Q4: Có thể share với người khác không?

**A:** Có nhưng cẩn thận!
- ✅ Click "Share" trong Google Sheet
- ✅ Add email người dùng
- ⚠️ Cân nhắc quyền Editor hay Viewer
- 🔒 Dữ liệu tài chính nhạy cảm!

### Q5: Làm sao backup dữ liệu?

**A:** Nhiều cách:
```
Cách 1: Manual copy
- File → Make a copy

Cách 2: Download
- File → Download → Excel (.xlsx)

Cách 3: Auto-backup
- Cấu hình trigger (xem phần Cấu hình nâng cao)
```

### Q6: Có giới hạn số giao dịch không?

**A:** Không!
- ✅ Google Sheets hỗ trợ 10 triệu cells
- ✅ Với 10 columns, được ~1 triệu rows
- ✅ Đủ cho hàng chục năm sử dụng

### Q7: Làm sao update version mới?

**A:** 2 cách:
```
Cách 1: Copy code mới vào Apps Script
- So sánh CHANGELOG
- Copy các file thay đổi
- Test kỹ

Cách 2: Make copy file mới, import data cũ
- Copy template mới
- Export data từ file cũ
- Import vào file mới
```

### Q8: Có thể tùy chỉnh không?

**A:** Có hoàn toàn!
- ✅ Code mã nguồn mở
- ✅ Chỉnh sửa thoải mái
- ✅ Thêm tính năng riêng
- ✅ Thay đổi giao diện

### Q9: Có support không?

**A:** Có nhiều kênh:
```
- 📖 Đọc documentation
- 🔍 Xem Troubleshooting
- 💬 Facebook Group
- 📧 Email hỗ trợ
- 🐛 GitHub Issues
```

### Q10: Làm sao xóa và bắt đầu lại?

**A:** Rất đơn giản:
```
1. Delete file Google Sheet
2. Make copy template mới
3. Chạy Setup Wizard
4. Done!
```

---

## 📞 Hỗ trợ cài đặt

### Khi cần trợ giúp

**Trước khi hỏi, hãy:**
1. ✅ Đọc kỹ hướng dẫn này
2. ✅ Xem phần Troubleshooting
3. ✅ Check các video hướng dẫn
4. ✅ Xem logs trong Apps Script

**Liên hệ hỗ trợ:**
- 💬 Facebook Group: [https://www.facebook.com/groups/hodl.vn]
- 📧 Email: contact@tohaitrieu.net
- 🐛 GitHub Issues: [https://github.com/tohaitrieu/HODLVN-Family-Finance/issues]
- 📱 Telegram: [https://t.me/totrieu]

**Khi báo lỗi, cung cấp:**
- 📋 Version đang dùng
- 💻 Trình duyệt + OS
- 📸 Screenshot lỗi
- 📝 Các bước đã làm
- 📊 Logs từ Apps Script (nếu có)

---

## 🎉 Hoàn thành

**Xin chúc mừng!** Bạn đã hoàn thành việc cài đặt HODLVN-Family-Finance.

### Bước tiếp theo

1. 📖 Đọc [User Guide](USER_GUIDE.md) để học cách sử dụng
2. 💰 Nhập giao dịch đầu tiên
3. 🎯 Đặt ngân sách tháng
4. 📊 Xem Dashboard

### Tham gia cộng đồng

- 👥 [Facebook Group](https://www.facebook.com/groups/hodl.vn)
- 🐦 [Twitter](https://x.com/tohaitrieu)
- 📺 [YouTube](https://www.youtube.com/@totrieu)
- 💬 [Telegram](https://t.me/totrieu)

---

<div align="center">

**Chúc bạn quản lý tài chính hiệu quả! 💰📊**

[⬆ Về đầu trang](#-hướng-dẫn-cài-đặt---hodlvn-family-finance)

</div>