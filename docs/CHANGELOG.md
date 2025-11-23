# 📝 CHANGELOG - HODLVN-Family-Finance

Tất cả các thay đổi quan trọng của dự án được ghi lại trong file này.

Format dựa trên [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
và dự án tuân theo [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

---

## 📑 MỤC LỤC

- [Unreleased](#unreleased)
- [3.0.0 - 2025-10-29](#300---2025-10-29)
- [2.5.0 - 2025-09-15](#250---2025-09-15)
- [2.0.0 - 2025-06-01](#200---2025-06-01)
- [1.5.0 - 2025-03-15](#150---2025-03-15)
- [1.0.0 - 2025-01-01](#100---2025-01-01)

---

## [Unreleased]

### Planned Features
- 📱 Mobile PWA (Progressive Web App)
- 🔔 Email/SMS notifications cho budget alerts
- 🌐 Multi-currency support (USD, EUR)
- 🤖 AI insights và recommendations
- 📊 Advanced analytics và forecasting
- 🔄 Auto-sync với banking API
- 🌙 Dark mode
- 📤 Export to Excel với formatting
- 🔐 Two-factor authentication
- 👥 Multi-user collaboration

---

## [3.5.0] - 2025-11-23

### Lending Refactor & Dynamic Dashboard

- ✅ **NEW**: Refactored Lending System with 3 specific types (Maturity, Interest Only, Installment).
- ✅ **NEW**: Dynamic Event Calendar using Custom Functions (`AccPayable`, `AccReceivable`).
- ✅ **FIX**: Debt Payment & Collection forms now support precise amounts (removed 1000 step).
- ✅ **FIX**: Dashboard tables row offset issues.
- ✅ **UPDATE**: Event Calendar now shows only the nearest upcoming event per loan.

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

## [3.0.0] - 2025-10-29

### 🎉 Major Release - Modular Architecture

Đây là bản cập nhật lớn nhất với việc tái cấu trúc hoàn toàn từ monolithic sang modular architecture.

### ✨ Added

#### Core Architecture
- **Modular Architecture v3.0**: Tách code thành 6 modules độc lập
  - `Main.gs`: Menu dispatcher và event router
  - `SheetInitializer.gs`: Sheet setup và configuration
  - `DataProcessor.gs`: Transaction handling
  - `BudgetManager.gs`: Budget tracking và alerts
  - `DashboardManager.gs`: Analytics và reporting
  - `Utils.gs`: Shared utilities

#### Setup Wizard
- **Setup Wizard mới hoàn toàn**:
  - Thu thập thông tin người dùng (tên, email, phone)
  - Cấu hình năm bắt đầu
  - Đặt ngân sách ban đầu
  - Khởi tạo tất cả sheets tự động
  - Loại bỏ sample data (thay bằng user data thực)

#### Budget Management
- **3-level Budget Alerts**:
  - 🟢 Xanh: < 70% ngân sách (An toàn)
  - 🟡 Vàng: 70-90% (Cảnh báo)
  - 🔴 Đỏ: > 90% (Nguy hiểm)
- Conditional formatting tự động
- Real-time budget tracking
- Budget alerts khi vượt ngưỡng

#### Investment Tracking
- **4 loại đầu tư riêng biệt**:
  - 📈 Chứng khoán (với margin support)
  - 🪙 Vàng (SJC, PNJ, DOJI)
  - ₿ Cryptocurrency (BTC, ETH, altcoins)
  - 💼 Đầu tư khác (BĐS, quỹ, trái phiếu, tiết kiệm)
- Swap transactions cho crypto
- Portfolio tracking

#### Debt Management Enhanced
- **Auto-create income khi vay nợ**:
  - Tất cả loại nợ đều tự động tạo thu nhập
  - Income source: "Vay [Debt Type]"
  - Đúng với logic: Vay tiền = Tăng cash flow
- Margin debt tự động từ stock transactions
- Debt balance tracking real-time
- Status: "Chưa thanh toán" / "Đã thanh toán"

#### Dividend Tracking
- **Sheet CỔ TỨC mới**:
  - Theo dõi cổ tức tiền mặt
  - Theo dõi cổ tức cổ phiếu
  - Tự động cập nhật portfolio khi nhận CT CP
  - Tính Dividend Yield

#### UI/UX Improvements
- Form riêng cho từng loại giao dịch (không dùng unified form)
- Color-coded forms (mỗi form 1 màu riêng)
- Auto-calculation trong forms
- Validation tốt hơn
- Vietnamese throughout

### 🐛 Fixed

#### Critical Bug Fixes
- **🔥 CRITICAL: Fix appendRow() bug với pre-filled formulas**:
  - **Vấn đề**: Data xuất hiện ở row 1001 thay vì row tiếp theo
  - **Nguyên nhân**: `appendRow()` coi cells có formula là "occupied"
  - **Giải pháp**: Dùng `getRange().setValues()` với `findEmptyRow()`
  - **Impact**: Ảnh hưởng tất cả 10 data sheets
  - **Files fixed**: 
    - `DataProcessor.gs` (tất cả add* functions)
    - `DebtManagementHandler.gs` (addDebt, payDebt)
    - `SheetInitializer.gs` (tất cả setup functions)

- **Fix findEmptyRow() logic**:
  - Tìm row trống dựa vào cột DATA (column B - Ngày)
  - KHÔNG dựa vào cột STT (có formula)
  - KHÔNG dùng `getLastRow()` (không đáng tin cậy với formulas)

#### Data Integrity Fixes
- Fix STT formula: `=IF(B2<>"", ROW()-1, "")` thay vì `=ROW()-1`
- Fix budget formulas không update
- Fix dashboard không refresh khi có giao dịch mới
- Fix date format inconsistency (unified to ISO 8601)
- Fix number format cho VNĐ

#### Form Fixes
- Fix dropdown không load data
- Fix validation bypass
- Fix form không close sau submit
- Fix parameter mismatch trong investment forms
- Fix margin checkbox không hoạt động

### 🔄 Changed

#### Breaking Changes
- **⚠️ Code structure hoàn toàn mới**:
  - Migration từ v2.0 sang v3.0 yêu cầu re-setup
  - Không backward compatible với v2.x
  - Phải chạy Setup Wizard

- **Sheet structure updates**:
  - Thêm column "Status" trong QUẢN LÝ NỢ
  - Thêm column "Margin" trong CHỨNG KHOÁN
  - Thêm column "Loại CT" trong CỔ TỨC
  - Budget sheet có conditional formatting

- **Function signatures changed**:
  ```javascript
  // Old (v2.0)
  addTransaction(type, data)
  
  // New (v3.0)
  addIncome(data)
  addExpense(data)
  addDebt(data)
  // Mỗi loại có function riêng
  ```

#### Improved Features
- Cash Flow formula chính xác hơn:
  ```
  Old: Thu - Chi
  New: Thu - (Chi + Trả nợ + Đầu tư)
  ```
- Performance optimization (batch operations)
- Better error handling
- Detailed logging
- Code documentation

### 🗑️ Deprecated

- ~~`addTransaction()`~~ → Dùng `addIncome()`, `addExpense()`, etc.
- ~~`UnifiedForm.html`~~ → Dùng separate forms
- ~~Sample data initialization~~ → Dùng Setup Wizard
- ~~Manual formula fill~~ → Auto pre-filled

### 🔒 Security

- Input validation cho tất cả forms
- Sanitize user inputs
- Script properties cho sensitive data
- Permission restrictions

### 📚 Documentation

- README.md hoàn chỉnh với badges
- INSTALLATION.md chi tiết (2 phương pháp)
- USER_GUIDE.md đầy đủ 17 sections
- TECHNICAL_DOCUMENTATION.md (16 sections)
- CHANGELOG.md (file này)
- TROUBLESHOOTING.md (upcoming)

### 🎓 Migration Guide (v2.0 → v3.0)

**Option 1: Clean Install (Recommended)**
```
1. Make a copy of v3.0 template
2. Run Setup Wizard with your info
3. Manual migrate data from v2.0:
   - Export v2.0 data to CSV
   - Import to v3.0 sheets
4. Verify all formulas working
```

**Option 2: In-place Upgrade (Advanced)**
```
1. Backup v2.0 file
2. Copy all v3.0 .gs files to Apps Script
3. Delete old Code.gs
4. Run initializeAllSheets()
5. Fix any formula errors
6. Test thoroughly
```

---

## [2.5.0] - 2025-09-15

### ✨ Added

- **Crypto tracking**: Sheet mới cho cryptocurrency
- **Gold tracking**: Sheet mới cho vàng
- **Other investments**: Sheet cho các kênh đầu tư khác
- Payment method tracking trong expenses
- Budget percentage display

### 🐛 Fixed

- Fix margin calculation errors
- Fix date picker issues
- Fix currency formatting
- Memory leak trong loops

### 🔄 Changed

- Updated Google Apps Script runtime to V8
- Improved form validation
- Better mobile responsiveness

### 📚 Documentation

- Added video tutorials
- Updated FAQ
- Added troubleshooting guide

---

## [2.0.0] - 2025-06-01

### 🎉 Major Release - Budget Tracking

### ✨ Added

#### Budget System
- **Budget sheet mới**:
  - Đặt ngân sách theo danh mục
  - Theo dõi % sử dụng
  - Cảnh báo khi vượt ngân sách
- SetBudgetForm.html

#### Stock Trading
- **CHỨNG KHOÁN sheet**:
  - Mua/bán cổ phiếu
  - Tính phí & thuế tự động
  - Portfolio tracking
- StockForm.html

#### Enhanced Dashboard
- Biểu đồ thu chi theo tháng
- Biểu đồ phân bổ chi tiêu
- Cash Flow visualization
- Period comparison (MoM, YoY)

### 🐛 Fixed

- Fix formula errors trong income sheet
- Fix dashboard không refresh
- Fix dropdown categories bị duplicate

### 🔄 Changed

- Refactored Code.gs (split some functions)
- Updated UI colors
- Improved error messages

### 🗑️ Deprecated

- ~~BasicForm.html~~ → Use specialized forms

---

## [1.5.0] - 2025-03-15

### ✨ Added

#### Debt Management
- **QUẢN LÝ NỢ sheet**:
  - Danh sách các khoản nợ
  - Lãi suất, kỳ hạn
  - Ngày đáo hạn
- **TRẢ NỢ sheet**:
  - Lịch sử trả nợ
  - Trả gốc + lãi
  - Tính lãi tự động

#### Forms
- DebtManagementForm.html
- DebtPaymentForm.html
- Auto-calculate interest

### 🐛 Fixed

- Fix date validation
- Fix amount input không chấp nhận số thập phân
- Fix menu không hiện trên mobile

### 🔄 Changed

- Vietnamese menu items
- Better form layouts
- Mobile-friendly forms

---

## [1.0.0] - 2025-01-01

### 🎉 Initial Release

### ✨ Added

#### Core Features
- **THU sheet**: Thu nhập tracking
- **CHI sheet**: Chi tiêu tracking
- **Dashboard**: Tổng quan tài chính
- Simple menu system

#### Basic Forms
- IncomeForm.html
- ExpenseForm.html

#### Calculations
- Tổng thu
- Tổng chi
- Cash Flow cơ bản (Thu - Chi)

#### Data Structure
- 10 expense categories
- 5 income categories
- STT auto-increment
- Date tracking

### 🎯 Core Principles

- Google Sheets based
- Vietnamese language
- Simple & intuitive
- Free & open source
- Privacy-focused (data on personal Drive)

---

## Version Numbering

Dự án sử dụng [Semantic Versioning](https://semver.org/):

```
MAJOR.MINOR.PATCH

MAJOR: Breaking changes, incompatible API changes
MINOR: New features, backward compatible
PATCH: Bug fixes, backward compatible
```

**Examples:**
- `3.0.0`: Modular architecture (breaking change)
- `2.5.0`: Added crypto tracking (new feature)
- `2.0.1`: Fixed formula bug (bug fix)

---

## Release Schedule

- **Major releases**: 2-3 times per year
- **Minor releases**: Monthly
- **Patch releases**: As needed

---

## How to Upgrade

### Check Current Version

```javascript
// In Apps Script Editor
Logger.log(PropertiesService.getScriptProperties()
  .getProperty('VERSION'));
```

### Upgrade Steps

1. **Backup current file**
   ```
   File → Make a copy
   Rename: "[Filename] - Backup - [Date]"
   ```

2. **Read CHANGELOG**
   - Check breaking changes
   - Review migration guide
   - Note deprecated features

3. **Test in copy first**
   - Don't upgrade production directly
   - Test all features
   - Verify data integrity

4. **Upgrade production**
   - Copy new code
   - Run migration scripts
   - Update formulas if needed

5. **Verify**
   - Check all sheets
   - Test all forms
   - Review dashboard

---

## Reporting Issues

Nếu gặp bug sau khi upgrade:

1. **Check version**:
   ```
   Menu → About → Version
   ```

2. **Check CHANGELOG**:
   - Is it a known issue?
   - Is there a workaround?

3. **Report on GitHub**:
   ```
   Title: [v3.0.0] Brief description
   
   **Version**: 3.0.0
   **Browser**: Chrome 120
   **OS**: Windows 11
   
   **Steps to reproduce**:
   1. ...
   2. ...
   
   **Expected**: ...
   **Actual**: ...
   
   **Screenshots**: ...
   ```

---

## Contributors

### v3.0.0
- **Nika** - Architecture redesign, modular refactoring
- **Community** - Bug reports, feature requests

### v2.0.0
- **Nika** - Budget system, stock tracking

### v1.0.0
- **Nika** - Initial release, core features

---

## Acknowledgments

Cảm ơn:
- Google Apps Script team
- Vietnamese personal finance community
- All beta testers
- Stack Overflow contributors
- Open source community

---

## Future Roadmap

### Q1 2026 (v3.1.0)
- [ ] Export/Import Excel
- [ ] Auto-backup
- [ ] Multi-currency
- [ ] Dark mode
- [ ] Custom templates

### Q2 2026 (v3.2.0)
- [ ] Mobile PWA
- [ ] Email notifications
- [ ] Real-time price API (gold, stocks, crypto)
- [ ] AI insights
- [ ] Sharing & collaboration

### Q3 2026 (v3.5.0)
- [ ] Banking API integration
- [ ] Tax calculation
- [ ] Investment recommendations
- [ ] Community marketplace

### Q4 2026 (v4.0.0)
- [ ] Machine Learning predictions
- [ ] Chatbot assistant
- [ ] Advanced analytics
- [ ] Mobile native app

---

## Breaking Changes History

### v3.0.0
- Code structure: Monolithic → Modular
- Sheet setup: Manual → Setup Wizard
- Forms: Unified → Separate
- Functions: Generic → Specific

### v2.0.0
- Cash Flow: Simple → Comprehensive
- Data entry: Manual → Forms
- Dashboard: Basic → Advanced

---

## Deprecation Policy

- **Announce**: Deprecated features marked in CHANGELOG
- **Support**: 2 versions (6 months minimum)
- **Remove**: After support period ends

**Example:**
```
v2.5: Feature X deprecated
v3.0: Feature X still works (warning)
v3.5: Feature X removed
```

---

## License Changes

- v1.0.0 - v3.0.0: MIT License
- All versions: Free & open source

---

## Statistics

### Lines of Code

| Version | Total LoC | Modules | Forms | Sheets |
|---------|-----------|---------|-------|--------|
| v1.0.0  | ~500      | 1       | 2     | 3      |
| v2.0.0  | ~1500     | 1       | 5     | 7      |
| v3.0.0  | ~3000     | 6       | 10    | 12     |

### Features Count

| Version | Sheets | Forms | Functions | Features |
|---------|--------|-------|-----------|----------|
| v1.0.0  | 3      | 2     | 10        | 3        |
| v2.0.0  | 7      | 5     | 30        | 8        |
| v3.0.0  | 12     | 10    | 60+       | 15+      |

### Community Growth

| Version | Users | GitHub Stars | Contributors |
|---------|-------|--------------|--------------|
| v1.0.0  | 10    | -            | 1            |
| v2.0.0  | 100   | -            | 1            |
| v3.0.0  | 500+  | TBD          | 2+           |

---

## Contact

- **Email**: contact@tohaitrieu.net
- **GitHub**: [Issues](https://github.com/tohaitrieu/HODLVN-Family-Finance/issues)
- **Facebook**: [Group](https://facebook.com/groups/hodl.vn)
- **Telegram**: [Channel](https://t.me/totrieu)

---

<div align="center">

**Cảm ơn bạn đã sử dụng HODLVN-Family-Finance! 🎉**

[⬆ Về đầu trang](#-changelog---hodlvn-family-finance)

</div>