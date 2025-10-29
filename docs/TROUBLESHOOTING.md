# 🔍 XỬ LÝ SỰ CỐ - HODLVN-Family-Finance

Tài liệu tổng hợp tất cả các lỗi đã gặp trong quá trình phát triển và cách khắc phục.

---

## 📑 MỤC LỤC

1. [Lỗi Nghiêm Trọng (Critical Issues)](#-lỗi-nghiêm-trọng-critical-issues)
2. [Lỗi Thường Gặp (Common Bugs)](#-lỗi-thường-gặp-common-bugs)
3. [Lỗi Forms & Data Entry](#-lỗi-forms--data-entry)
4. [Lỗi Dashboard & Validation](#-lỗi-dashboard--validation)
5. [Lỗi Authorization & Permissions](#-lỗi-authorization--permissions)
6. [Lỗi Performance](#-lỗi-performance)
7. [Solutions & Best Practices](#-solutions--best-practices)
8. [Prevention Guide](#-prevention-guide)
9. [Testing Procedures](#-testing-procedures)
10. [FAQ Kỹ Thuật](#-faq-kỹ-thuật)

---

## 🔴 Lỗi Nghiêm Trọng (Critical Issues)

### 1. ❌ BUG NGHIÊM TRỌNG: appendRow() với Pre-filled Formulas

**Triệu chứng:**
```
- Dữ liệu xuất hiện ở dòng 1001 thay vì dòng tiếp theo
- Các dòng 2-1000 chứa công thức nhưng trống dữ liệu
- Sheet THU hoạt động OK nhưng các sheet khác bị lỗi
```

**Nguyên nhân:**
```javascript
// ❌ CODE CŨ - SAI
function addIncome(data) {
  const sheet = getSheet('THU');
  const lastRow = sheet.getLastRow();  // → Trả về 1000 (do formulas)
  const stt = lastRow > 1 ? lastRow : 1;
  
  sheet.appendRow([stt, date, amount, ...]);  
  // → appendRow() thấy dòng 1-1000 đã có formulas
  // → Coi như "occupied" 
  // → Thêm vào dòng 1001 ❌
}
```

**Giải thích chi tiết:**

Google Sheets' `appendRow()` method có behavior đặc biệt:
- Nó coi tất cả cells có **công thức** là "occupied"
- Khi sheet có pre-filled formulas từ dòng 2-1000
- `appendRow()` sẽ skip qua và chèn vào dòng 1001

**Tại sao Sheet THU OK nhưng QUẢN LÝ NỢ lỗi?**
- **THU**: Dữ liệu thêm qua Setup Wizard → không dùng `appendRow()`
- **QUẢN LÝ NỢ**: Dữ liệu thêm qua form → dùng `appendRow()` → lỗi

**Giải pháp:**

```javascript
// ✅ CODE MỚI - ĐÚNG
function addIncome(data) {
  const sheet = getSheet('THU');
  const emptyRow = findEmptyRow(sheet, 2);   // → Tìm dòng trống thực sự
  const stt = getNextSTT(sheet, 2);          // → Tính STT đúng
  const rowData = [stt, date, amount, ...];
  
  // Dùng setValues() thay vì appendRow()
  sheet.getRange(emptyRow, 1, 1, rowData.length)
       .setValues([rowData]);
  
  formatNewRow(sheet, emptyRow, {...});
}

/**
 * Tìm dòng trống dựa trên cột dữ liệu thực (không phải cột công thức)
 */
function findEmptyRow(sheet, startRow, dataColumn = 1) {
  const lastRow = sheet.getMaxRows();
  const values = sheet.getRange(startRow, dataColumn, lastRow - startRow + 1, 1).getValues();
  
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] === '' || values[i][0] === null) {
      return startRow + i;
    }
  }
  
  return lastRow + 1;
}

/**
 * Tính STT tiếp theo
 */
function getNextSTT(sheet, startRow, sttColumn = 1) {
  const emptyRow = findEmptyRow(sheet, startRow, sttColumn);
  if (emptyRow === startRow) return 1;
  
  const lastSTT = sheet.getRange(emptyRow - 1, sttColumn).getValue();
  return (typeof lastSTT === 'number') ? lastSTT + 1 : 1;
}
```

**Files cần fix:**
- ✅ `Utils.gs` - Thêm helper functions
- ✅ `DataProcessor.gs` - Sửa tất cả hàm add*()
- ✅ `DebtManagementHandler.gs` - Sửa addDebtManagement()
- ✅ `SheetInitializer.gs` - Nếu dùng appendRow()

**Test để verify:**
```javascript
function testFindEmptyRow() {
  const sheet = getSheet('THU');
  Logger.log('Empty row: ' + findEmptyRow(sheet, 2));
  Logger.log('Next STT: ' + getNextSTT(sheet, 2));
  
  // Kết quả mong đợi:
  // Empty row: 2 (nếu sheet trống)
  // Empty row: 3 (nếu có 1 dòng dữ liệu)
  // Next STT: 1 (nếu sheet trống)
  // Next STT: 2 (nếu có 1 dòng dữ liệu)
}
```

**Impact:**
- 🔴 **Severity**: CRITICAL
- 🔴 **Priority**: HIGHEST
- 🔴 **Affected**: Tất cả forms nhập liệu (trừ Setup Wizard)

---

### 2. ❌ BUG NGHIÊM TRỌNG: Missing Helper Functions

**Triệu chứng:**
```
ReferenceError: findEmptyRow is not defined
ReferenceError: getNextSTT is not defined
ReferenceError: formatNewRow is not defined
```

**Nguyên nhân:**
File `Utils.gs` trong project thiếu các hàm helper quan trọng:
1. `findEmptyRow()` - Tìm dòng trống thực sự
2. `getNextSTT()` - Tính số thứ tự tiếp theo
3. `formatNewRow()` - Format dòng mới (border, alignment)

**Giải pháp:**

Thêm vào `Utils.gs`:

```javascript
/**
 * ===============================================
 * HELPER FUNCTIONS
 * ===============================================
 */

/**
 * Tìm dòng trống đầu tiên (dựa trên cột dữ liệu, không phải cột công thức)
 * @param {Sheet} sheet - Sheet cần tìm
 * @param {number} startRow - Dòng bắt đầu tìm kiếm (thường là 2)
 * @param {number} dataColumn - Cột dữ liệu để check (mặc định cột 1 - STT)
 * @return {number} Số thứ tự dòng trống
 */
function findEmptyRow(sheet, startRow, dataColumn = 1) {
  const lastRow = sheet.getMaxRows();
  const values = sheet.getRange(startRow, dataColumn, lastRow - startRow + 1, 1).getValues();
  
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] === '' || values[i][0] === null) {
      return startRow + i;
    }
  }
  
  return lastRow + 1; // Nếu không tìm thấy, thêm dòng mới
}

/**
 * Lấy STT tiếp theo
 * @param {Sheet} sheet - Sheet cần lấy STT
 * @param {number} startRow - Dòng bắt đầu
 * @param {number} sttColumn - Cột STT (mặc định 1)
 * @return {number} STT tiếp theo
 */
function getNextSTT(sheet, startRow, sttColumn = 1) {
  const emptyRow = findEmptyRow(sheet, startRow, sttColumn);
  
  if (emptyRow === startRow) {
    return 1; // Dòng đầu tiên
  }
  
  const lastSTT = sheet.getRange(emptyRow - 1, sttColumn).getValue();
  return (typeof lastSTT === 'number') ? lastSTT + 1 : 1;
}

/**
 * Format dòng mới (border, căn giữa, v.v.)
 * @param {Sheet} sheet - Sheet cần format
 * @param {number} row - Số dòng cần format
 * @param {Object} options - Tùy chọn format
 */
function formatNewRow(sheet, row, options = {}) {
  const lastCol = sheet.getLastColumn();
  const range = sheet.getRange(row, 1, 1, lastCol);
  
  // Border
  range.setBorder(
    true, true, true, true, true, true,
    '#000000', SpreadsheetApp.BorderStyle.SOLID
  );
  
  // Căn giữa các cột số (nếu có)
  if (options.centerColumns) {
    options.centerColumns.forEach(col => {
      sheet.getRange(row, col).setHorizontalAlignment('center');
    });
  }
  
  // Number format (nếu có)
  if (options.numberColumns) {
    options.numberColumns.forEach(col => {
      sheet.getRange(row, col).setNumberFormat('#,##0');
    });
  }
}

/**
 * Insert row at specific position với data
 * @param {Sheet} sheet - Sheet cần insert
 * @param {number} row - Vị trí insert
 * @param {Array} rowData - Dữ liệu dòng
 */
function insertRowAtEmptyPosition(sheet, row, rowData) {
  sheet.getRange(row, 1, 1, rowData.length).setValues([rowData]);
}
```

**Test:**
```javascript
function testHelperFunctions() {
  const sheet = getSheet('THU');
  
  Logger.log('=== TEST HELPER FUNCTIONS ===');
  Logger.log('Empty row: ' + findEmptyRow(sheet, 2));
  Logger.log('Next STT: ' + getNextSTT(sheet, 2));
  
  // Test insert
  const testData = [1, new Date(), 1000000, 'Lương', 'Test', '', '', '', '', ''];
  const emptyRow = findEmptyRow(sheet, 2);
  insertRowAtEmptyPosition(sheet, emptyRow, testData);
  formatNewRow(sheet, emptyRow, {
    centerColumns: [1, 2],
    numberColumns: [3]
  });
  
  Logger.log('✅ Insert thành công tại dòng: ' + emptyRow);
}
```

---

## ⚠️ Lỗi Thường Gặp (Common Bugs)

### 3. ❌ Parameter Mismatch trong Investment Forms

**Triệu chứng:**
```
- Form Gold, Crypto hoạt động OK
- Form "Đầu tư khác" submit nhưng không có dữ liệu
- Validation failed: "Vui lòng điền đầy đủ thông tin"
```

**Nguyên nhân:**

```javascript
// ❌ OtherInvestmentForm.html
const data = {
  date: document.getElementById('date').value,
  investmentType: document.getElementById('investmentType').value,  // ← investmentType
  amount: parseFloat(document.getElementById('amount').value),
  note: document.getElementById('note').value
};
google.script.run.addOtherInvestment(data);

// ❌ DataProcessor.gs
function addOtherInvestment(data) {
  // Validation
  if (!data.date || !data.type || !data.amount) {  // ← Tìm "type" không phải "investmentType"
    return {
      success: false,
      message: '❌ Vui lòng điền đầy đủ thông tin bắt buộc!'
    };
  }
  
  const type = data.type;  // ← undefined!
  const rowData = [stt, date, type, amount, note, '', '', '', '', ''];
  // → Chỉ 3 cột đầu có dữ liệu (stt, date, undefined)
}
```

**So sánh với Gold Form (OK):**

```javascript
// ✅ GoldForm.html
const data = {
  type: document.getElementById('type').value,  // ← "type"
  ...
};

// ✅ DataProcessor.gs
function addGold(data) {
  if (!data.type || !data.goldType || !data.amount) {  // ← Tìm "type"
    // ...
  }
  const type = data.type;  // ← OK!
}
```

**Giải pháp:**

**Option 1: Sửa DataProcessor (Khuyến nghị)**
```javascript
// ✅ DataProcessor.gs - FIX
function addOtherInvestment(data) {
  // Validation - Sửa "type" → "investmentType"
  if (!data.date || !data.investmentType || !data.amount) {
    return {
      success: false,
      message: '❌ Vui lòng điền đầy đủ thông tin bắt buộc!'
    };
  }
  
  const type = data.investmentType;  // ← FIX!
  const rowData = [stt, date, type, amount, note, '', '', '', '', ''];
}
```

**Option 2: Sửa Form HTML**
```javascript
// ✅ OtherInvestmentForm.html - FIX
const data = {
  date: document.getElementById('date').value,
  type: document.getElementById('investmentType').value,  // ← Đổi key thành "type"
  amount: parseFloat(document.getElementById('amount').value),
  note: document.getElementById('note').value
};
```

**Recommendation:** Sửa DataProcessor (Option 1) để nhất quán với form HTML.

**Files bị ảnh hưởng:**
- `DataProcessor.gs` - Hàm `addOtherInvestment()`
- `OtherInvestmentForm.html` (nếu chọn Option 2)

---

### 4. ❌ Forms Không Hoạt Động Sau Restructure v3.0

**Triệu chứng:**
```
- Click menu → Form không mở
- Console error: "Cannot find function showIncomeForm"
- Hoặc form mở nhưng submit không có phản hồi
```

**Nguyên nhân:**

Sau khi tái cấu trúc từ v2.0 → v3.0 (modular), các hàm được phân tán vào các module riêng:
- `Main.gs` - Menu & UI handlers
- `DataProcessor.gs` - Data processing functions
- `SheetInitializer.gs` - Sheet initialization
- `BudgetManager.gs` - Budget operations
- `DashboardManager.gs` - Dashboard operations
- `Utils.gs` - Helper functions

Nhưng **các form HTML vẫn gọi hàm theo cách cũ**:

```javascript
// ❌ IncomeForm.html (v2.0 style)
google.script.run.addIncome(data);  
// → Tìm hàm addIncome() trong Main.gs
// → KHÔNG TÌM THẤY (đã chuyển sang DataProcessor.gs)
```

**Giải pháp:**

**Option 1: Namespace (Khuyến nghị cho v3.0+)**

```javascript
// ✅ DataProcessor.gs
var DataProcessor = DataProcessor || {};

DataProcessor.addIncome = function(data) {
  // Implementation...
};

// ✅ IncomeForm.html
google.script.run.DataProcessor.addIncome(data);
```

**Option 2: Wrapper Functions (Đơn giản hơn)**

```javascript
// ✅ Main.gs - Giữ lại wrapper functions
function addIncome(data) {
  return DataProcessor.addIncome(data);
}

function addExpense(data) {
  return DataProcessor.addExpense(data);
}

// Form HTML không cần thay đổi
google.script.run.addIncome(data);  // ← Vẫn hoạt động!
```

**Option 3: Import Pattern**

```javascript
// ✅ Main.gs
// Re-export tất cả functions cần thiết
var addIncome = DataProcessor.addIncome;
var addExpense = DataProcessor.addExpense;
var addDebtPayment = DataProcessor.addDebtPayment;
// ...
```

**Recommendation:** 
- Dự án nhỏ: Option 2 (Wrapper)
- Dự án lớn, nhiều developer: Option 1 (Namespace)

**Checklist sau restructure:**
```
☐ Tất cả menu functions trong Main.gs
☐ Tất cả form submission handlers accessible
☐ Tất cả callback functions (success/failure) hoạt động
☐ Test từng form một
☐ Check console errors
```

---

### 5. ❌ Dữ liệu Không Đầy Đủ trong Gold Form

**Triệu chứng:**
```
- Submit Gold form thành công
- Nhưng chỉ 3 cột đầu có dữ liệu (STT, Ngày, Loại GD)
- Các cột còn lại (Loại vàng, Số lượng, Đơn vị...) trống
```

**Nguyên nhân:**

```javascript
// ❌ DataProcessor.gs - CODE CŨ
function addGold(data) {
  // Validation yếu - dùng default value
  const unit = data.unit || 'chỉ';  // ← Default!
  const goldType = data.goldType || 'SJC';
  
  // Nếu form không gửi unit → data.unit = undefined
  // → unit = 'chỉ' → validation PASS
  // → Nhưng thực tế data.unit vẫn là undefined
  // → Chỉ insert được 3 cột đầu!
  
  const rowData = [stt, date, type, goldType, amount, unit, price, total, location, note];
  // → [1, '2025-01-15', 'Mua', undefined, undefined, undefined, ...]
}
```

**Giải pháp:**

```javascript
// ✅ DataProcessor.gs - FIX
function addGold(data) {
  // Validation chặt chẽ - KHÔNG dùng default
  if (!data.date || !data.type || !data.goldType || !data.amount || !data.unit) {
    return {
      success: false,
      message: '❌ Vui lòng điền đầy đủ thông tin bắt buộc!'
    };
  }
  
  // Không dùng default value
  const unit = data.unit;
  const goldType = data.goldType;
  const location = data.location || '';  // ← Optional field mới dùng default
  
  // Thêm Logger để debug
  Logger.log('=== ADD GOLD DEBUG ===');
  Logger.log('Data received: ' + JSON.stringify(data));
  Logger.log('Unit: ' + unit);
  Logger.log('Gold Type: ' + goldType);
  
  const rowData = [stt, date, type, goldType, amount, unit, price, total, location, note];
  Logger.log('Row Data: ' + JSON.stringify(rowData));
  
  // Verify data before insert
  const emptyCount = rowData.filter(cell => cell === undefined || cell === '').length;
  if (emptyCount > 2) {  // Allow 2 optional fields
    Logger.log('⚠️ WARNING: Too many empty cells: ' + emptyCount);
    return {
      success: false,
      message: '❌ Dữ liệu không đầy đủ. Vui lòng kiểm tra lại!'
    };
  }
  
  // Insert data
  insertRowAtEmptyPosition(sheet, emptyRow, rowData);
}
```

**Form HTML cần kiểm tra:**

```html
<!-- ✅ GoldForm.html - Đảm bảo tất cả required fields -->
<select id="unit" required>
  <option value="">-- Chọn đơn vị --</option>
  <option value="chỉ">Chỉ</option>
  <option value="lượng">Lượng</option>
  <option value="gram">Gram</option>
</select>

<script>
  function submitForm() {
    // Validate required fields
    const requiredFields = ['date', 'type', 'goldType', 'amount', 'unit'];
    for (let field of requiredFields) {
      const value = document.getElementById(field).value;
      if (!value) {
        alert('❌ Vui lòng điền đầy đủ: ' + field);
        return;
      }
    }
    
    // Collect data
    const data = {
      date: document.getElementById('date').value,
      type: document.getElementById('type').value,
      goldType: document.getElementById('goldType').value,
      amount: parseFloat(document.getElementById('amount').value),
      unit: document.getElementById('unit').value,  // ← QUAN TRỌNG
      price: parseFloat(document.getElementById('price').value) || 0,
      location: document.getElementById('location').value || '',
      note: document.getElementById('note').value || ''
    };
    
    console.log('Sending data:', data);  // Debug
    google.script.run.addGold(data);
  }
</script>
```

---

## 📝 Lỗi Forms & Data Entry

### 6. ❌ Form Submit Không Có Phản Hồi

**Triệu chứng:**
```
- Click Submit button
- Loading spinner xuất hiện
- Nhưng không có success/error message
- Form không đóng
```

**Nguyên nhân:**

```javascript
// ❌ IncomeForm.html - Thiếu callback handlers
function submitForm() {
  const data = {...};
  
  google.script.run.addIncome(data);  // ← Không có callback!
  
  // Loading spinner không tắt
  // Form không đóng
  // User không biết thành công hay thất bại
}
```

**Giải pháp:**

```javascript
// ✅ IncomeForm.html - FIX
function submitForm() {
  const data = {...};
  
  // Show loading
  document.getElementById('submitBtn').disabled = true;
  document.getElementById('submitBtn').textContent = 'Đang xử lý...';
  
  // Call với callbacks
  google.script.run
    .withSuccessHandler(onSuccess)
    .withFailureHandler(onFailure)
    .addIncome(data);
}

function onSuccess(response) {
  if (response.success) {
    alert('✅ ' + response.message);
    google.script.host.close();  // Đóng form
  } else {
    alert('❌ ' + response.message);
    resetButton();
  }
}

function onFailure(error) {
  alert('❌ Lỗi: ' + error.message);
  console.error('Error:', error);
  resetButton();
}

function resetButton() {
  document.getElementById('submitBtn').disabled = false;
  document.getElementById('submitBtn').textContent = 'Thêm Khoản Thu';
}
```

**Server-side phải return object:**

```javascript
// ✅ DataProcessor.gs
function addIncome(data) {
  try {
    // Process...
    
    return {
      success: true,
      message: 'Thêm khoản thu thành công!'
    };
  } catch (error) {
    Logger.log('Error in addIncome: ' + error.toString());
    return {
      success: false,
      message: 'Lỗi: ' + error.toString()
    };
  }
}
```

---

### 7. ❌ Form Validation Không Hoạt Động

**Triệu chứng:**
```
- Submit form với dữ liệu thiếu
- Không có cảnh báo
- Data được insert với giá trị rỗng/undefined
```

**Giải pháp:**

**Client-side validation (HTML5):**

```html
<!-- ✅ IncomeForm.html -->
<form id="incomeForm" onsubmit="return false;">
  <input type="date" id="date" required max="2099-12-31">
  <input type="number" id="amount" required min="0" step="1000">
  <select id="source" required>
    <option value="">-- Chọn nguồn thu --</option>
    <option value="Lương">Lương</option>
    <option value="Thưởng">Thưởng</option>
  </select>
</form>

<script>
  function submitForm() {
    // HTML5 validation tự động check
    const form = document.getElementById('incomeForm');
    if (!form.checkValidity()) {
      form.reportValidity();  // Hiện message lỗi
      return;
    }
    
    // Manual validation
    const amount = parseFloat(document.getElementById('amount').value);
    if (isNaN(amount) || amount <= 0) {
      alert('❌ Số tiền phải lớn hơn 0!');
      return;
    }
    
    // Proceed với submission...
  }
</script>
```

**Server-side validation (GAS):**

```javascript
// ✅ DataProcessor.gs
function addIncome(data) {
  // Type checking
  if (typeof data !== 'object' || data === null) {
    return {
      success: false,
      message: '❌ Dữ liệu không hợp lệ!'
    };
  }
  
  // Required fields
  if (!data.date || !data.amount || !data.source) {
    return {
      success: false,
      message: '❌ Vui lòng điền đầy đủ: Ngày, Số tiền, Nguồn thu!'
    };
  }
  
  // Date validation
  const date = new Date(data.date);
  if (isNaN(date.getTime())) {
    return {
      success: false,
      message: '❌ Ngày không hợp lệ!'
    };
  }
  
  // Amount validation
  const amount = parseFloat(data.amount);
  if (isNaN(amount) || amount <= 0) {
    return {
      success: false,
      message: '❌ Số tiền phải là số dương!'
    };
  }
  
  // String length validation
  if (data.note && data.note.length > 500) {
    return {
      success: false,
      message: '❌ Ghi chú không được quá 500 ký tự!'
    };
  }
  
  // Proceed...
}
```

---

## 📊 Lỗi Dashboard & Validation

### 8. ❌ Dashboard Dropdown Validation Failed

**Triệu chứng:**
```
- Dashboard có 2 dropdown: Khoản mục & Thời gian
- Khi chọn giá trị → validation error
- Công thức INDIRECT() không hoạt động
- Bảng kết quả hiện #REF! hoặc trống
```

**Nguyên nhân:**

```javascript
// ❌ DashboardManager.gs - CODE CŨ
function updateDashboard() {
  const dash = getSheet('DASHBOARD');
  
  // Set default value ngay trong cell có Data Validation
  dash.getRange('B2').setValue('Thu nhập');  // ← Conflict!
  
  // Data Validation đã set trên B2
  // → setValue() ghi đè validation
  // → Validation rule bị mất
  // → Formula INDIRECT() lỗi
}
```

**Giải pháp:**

```javascript
// ✅ DashboardManager.gs - FIX
function updateDashboard() {
  const dash = getSheet('DASHBOARD');
  
  // 1. Set validation TRƯỚC
  setupDashboardValidation(dash);
  
  // 2. KHÔNG set default value
  // Để cell trống → User phải chọn
  // → Validation hoạt động bình thường
  
  // 3. Nếu cần default, check validation trước
  const range = dash.getRange('B2');
  const rule = range.getDataValidation();
  if (!rule) {
    // Chỉ set value khi KHÔNG có validation
    range.setValue('Thu nhập');
  }
}

function setupDashboardValidation(dash) {
  // Dropdown 1: Khoản mục
  const categories = ['Thu nhập', 'Chi tiêu', 'Trả nợ', 'Đầu tư'];
  const rule1 = SpreadsheetApp.newDataValidation()
    .requireValueInList(categories, true)
    .setAllowInvalid(false)
    .build();
  dash.getRange('B2').setDataValidation(rule1);
  
  // Dropdown 2: Thời gian (KHÔNG set default value)
  const periods = ['Tháng này', 'Tháng trước', 'Quý này', 'Năm nay'];
  const rule2 = SpreadsheetApp.newDataValidation()
    .requireValueInList(periods, true)
    .setAllowInvalid(false)
    .build();
  dash.getRange('B3').setDataValidation(rule2);
  
  // ❌ KHÔNG LÀM NHƯ NÀY:
  // dash.getRange('B3').setValue('Tháng này');  // ← SAI!
}
```

**Named Range setup:**

```javascript
// ✅ Utils.gs hoặc DashboardManager.gs
function setupNamedRanges() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Tạo Named Range cho dropdown sources
  const thuSheet = getSheet('THU');
  const chiSheet = getSheet('CHI');
  
  // Income sources (cột D, từ dòng 2 đến cuối)
  const incomeRange = thuSheet.getRange('D2:D1000');
  ss.setNamedRange('IncomeSources', incomeRange);
  
  // Expense categories (cột D, từ dòng 2 đến cuối)
  const expenseRange = chiSheet.getRange('D2:D1000');
  ss.setNamedRange('ExpenseCategories', expenseRange);
  
  Logger.log('✅ Named Ranges created successfully');
}
```

**Formula sử dụng INDIRECT:**

```javascript
// Cell công thức trong DASHBOARD
=SUMIFS(INDIRECT("THU[C:C]"), INDIRECT("THU[B:B]"), ">="&DATE_START, INDIRECT("THU[B:B]"), "<="&DATE_END)
```

---

## 🔐 Lỗi Authorization & Permissions

### 9. ❌ Script Không Có Quyền Truy Cập

**Triệu chứng:**
```
Authorization required
You need to authorize this script
Exception: You do not have permission to call SpreadsheetApp.getActiveSpreadsheet
```

**Giải pháp:**

**Bước 1: Authorize script**
```
1. Extensions → Apps Script
2. Select function: onOpen hoặc bất kỳ function nào
3. Click ▶️ Run
4. Click "Review permissions"
5. Chọn tài khoản Google
6. Click "Advanced" → "Go to [Project Name] (unsafe)"
7. Click "Allow"
```

**Bước 2: Re-deploy nếu cần**
```
1. Apps Script Editor
2. Deploy → Manage deployments
3. Click ⚙️ Edit
4. Version: New version
5. Description: "Fix authorization"
6. Click Deploy
```

**Bước 3: Clear authorization và re-authorize**
```
1. Google Account → Security
2. Third-party apps with account access
3. Remove [Project Name]
4. Quay lại Apps Script và re-authorize
```

**Lưu ý:**
- Authorization required mỗi khi thay đổi Scopes
- Scopes được auto-detect từ code
- Check manifest file (`appsscript.json`) nếu cần manual scopes

---

### 10. ❌ Trigger Không Hoạt Động

**Triệu chứng:**
```
- onOpen trigger không chạy khi mở file
- onEdit trigger không chạy khi edit cell
- Time-based trigger không chạy đúng giờ
```

**Giải pháp:**

**Kiểm tra Triggers:**
```
1. Apps Script Editor
2. Left sidebar → Triggers (⏰ icon)
3. Xem list triggers hiện có
4. Xóa triggers cũ nếu có
5. Tạo trigger mới
```

**Tạo onOpen trigger manual:**
```javascript
// Main.gs
function createOnOpenTrigger() {
  // Xóa trigger cũ
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'onOpen') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // Tạo trigger mới
  ScriptApp.newTrigger('onOpen')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onOpen()
    .create();
    
  Logger.log('✅ onOpen trigger created');
}
```

**Debug trigger:**
```javascript
function debugTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  Logger.log('=== TRIGGERS DEBUG ===');
  Logger.log('Total triggers: ' + triggers.length);
  
  triggers.forEach((trigger, index) => {
    Logger.log('\nTrigger ' + (index + 1) + ':');
    Logger.log('- Handler: ' + trigger.getHandlerFunction());
    Logger.log('- Event Type: ' + trigger.getEventType());
    Logger.log('- Trigger Source: ' + trigger.getTriggerSource());
  });
}
```

---

## ⚡ Lỗi Performance

### 11. ❌ Script Timeout (Execution Time Exceeded)

**Triệu chứng:**
```
Exception: Exceeded maximum execution time
Script running time: 6 minutes
```

**Nguyên nhân:**
- Loop qua quá nhiều rows (> 5000)
- Nhiều `getRange()` calls riêng lẻ
- Không sử dụng batch operations

**Giải pháp:**

**BAD Practice:**
```javascript
// ❌ CHẬM - 1000 API calls
function updateAllRows() {
  const sheet = getSheet('THU');
  const lastRow = sheet.getLastRow();
  
  for (let i = 2; i <= lastRow; i++) {
    const value = sheet.getRange(i, 3).getValue();  // ← API call
    const newValue = value * 1.1;
    sheet.getRange(i, 3).setValue(newValue);  // ← API call
  }
}
```

**GOOD Practice:**
```javascript
// ✅ NHANH - 2 API calls
function updateAllRows() {
  const sheet = getSheet('THU');
  const lastRow = sheet.getLastRow();
  
  // 1 API call để lấy tất cả
  const values = sheet.getRange(2, 3, lastRow - 1, 1).getValues();
  
  // Process trong memory
  const newValues = values.map(row => [row[0] * 1.1]);
  
  // 1 API call để update tất cả
  sheet.getRange(2, 3, newValues.length, 1).setValues(newValues);
}
```

**Best Practices:**
1. **Batch reads/writes**: Dùng `getValues()` / `setValues()` thay vì `getValue()` / `setValue()`
2. **Reduce API calls**: Cache sheet references
3. **Use arrays**: Process trong memory thay vì trên sheet
4. **Limit range**: Chỉ get/set range cần thiết

---

## ✅ Solutions & Best Practices

### Helper Functions Template

```javascript
/**
 * ===============================================
 * UTILS.GS - HELPER FUNCTIONS TEMPLATE
 * ===============================================
 */

// ===== SHEET OPERATIONS =====

function getSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    throw new Error('Sheet không tồn tại: ' + sheetName);
  }
  
  return sheet;
}

function findEmptyRow(sheet, startRow, dataColumn = 1) {
  const lastRow = sheet.getMaxRows();
  const values = sheet.getRange(startRow, dataColumn, lastRow - startRow + 1, 1).getValues();
  
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] === '' || values[i][0] === null) {
      return startRow + i;
    }
  }
  
  return lastRow + 1;
}

function getNextSTT(sheet, startRow, sttColumn = 1) {
  const emptyRow = findEmptyRow(sheet, startRow, sttColumn);
  if (emptyRow === startRow) return 1;
  
  const lastSTT = sheet.getRange(emptyRow - 1, sttColumn).getValue();
  return (typeof lastSTT === 'number') ? lastSTT + 1 : 1;
}

// ===== DATA VALIDATION =====

function validateRequiredFields(data, requiredFields) {
  const missing = [];
  
  requiredFields.forEach(field => {
    if (!data[field] || data[field] === '') {
      missing.push(field);
    }
  });
  
  if (missing.length > 0) {
    return {
      valid: false,
      message: '❌ Thiếu thông tin: ' + missing.join(', ')
    };
  }
  
  return { valid: true };
}

function validateNumber(value, fieldName, options = {}) {
  const num = parseFloat(value);
  
  if (isNaN(num)) {
    return {
      valid: false,
      message: `❌ ${fieldName} phải là số!`
    };
  }
  
  if (options.min !== undefined && num < options.min) {
    return {
      valid: false,
      message: `❌ ${fieldName} phải >= ${options.min}!`
    };
  }
  
  if (options.max !== undefined && num > options.max) {
    return {
      valid: false,
      message: `❌ ${fieldName} phải <= ${options.max}!`
    };
  }
  
  return { valid: true, value: num };
}

function validateDate(dateString, fieldName) {
  const date = new Date(dateString);
  
  if (isNaN(date.getTime())) {
    return {
      valid: false,
      message: `❌ ${fieldName} không hợp lệ!`
    };
  }
  
  return { valid: true, value: date };
}

// ===== FORMATTING =====

function formatNewRow(sheet, row, options = {}) {
  const lastCol = sheet.getLastColumn();
  const range = sheet.getRange(row, 1, 1, lastCol);
  
  // Border
  range.setBorder(
    true, true, true, true, true, true,
    '#000000', SpreadsheetApp.BorderStyle.SOLID
  );
  
  // Căn giữa
  if (options.centerColumns) {
    options.centerColumns.forEach(col => {
      sheet.getRange(row, col).setHorizontalAlignment('center');
    });
  }
  
  // Number format
  if (options.numberColumns) {
    options.numberColumns.forEach(col => {
      sheet.getRange(row, col).setNumberFormat('#,##0');
    });
  }
  
  // Date format
  if (options.dateColumns) {
    options.dateColumns.forEach(col => {
      sheet.getRange(row, col).setNumberFormat('dd/mm/yyyy');
    });
  }
}

// ===== ERROR HANDLING =====

function tryCatch(fn, context = 'Operation') {
  try {
    return fn();
  } catch (error) {
    Logger.log(`❌ Error in ${context}: ${error.toString()}`);
    Logger.log('Stack trace: ' + error.stack);
    
    return {
      success: false,
      message: `❌ Lỗi: ${error.message}`
    };
  }
}

// ===== LOGGING =====

function logInfo(message, data = null) {
  Logger.log(`ℹ️ ${message}`);
  if (data) {
    Logger.log(JSON.stringify(data, null, 2));
  }
}

function logError(message, error) {
  Logger.log(`❌ ${message}`);
  Logger.log('Error: ' + error.toString());
  Logger.log('Stack: ' + error.stack);
}

function logSuccess(message) {
  Logger.log(`✅ ${message}`);
}
```

---

### DataProcessor Template

```javascript
/**
 * ===============================================
 * DATAPROCESSOR.GS - TEMPLATE
 * ===============================================
 */

function addTransaction(sheetName, data, config) {
  return tryCatch(() => {
    // 1. Get sheet
    const sheet = getSheet(sheetName);
    
    // 2. Validate
    const validation = validateRequiredFields(data, config.requiredFields);
    if (!validation.valid) {
      return validation;
    }
    
    // 3. Process data
    const processedData = config.processData(data);
    
    // 4. Find empty row & get STT
    const emptyRow = findEmptyRow(sheet, 2);
    const stt = getNextSTT(sheet, 2);
    
    // 5. Build row data
    const rowData = [stt, ...processedData];
    
    // 6. Insert
    sheet.getRange(emptyRow, 1, 1, rowData.length).setValues([rowData]);
    
    // 7. Format
    formatNewRow(sheet, emptyRow, config.formatOptions);
    
    // 8. Post-process (if any)
    if (config.postProcess) {
      config.postProcess(sheet, emptyRow, data);
    }
    
    // 9. Return success
    logSuccess(`${sheetName}: Thêm dữ liệu tại dòng ${emptyRow}`);
    return {
      success: true,
      message: config.successMessage || '✅ Thêm dữ liệu thành công!',
      row: emptyRow
    };
    
  }, `addTransaction (${sheetName})`);
}

// Example usage:
function addIncome(data) {
  return addTransaction('THU', data, {
    requiredFields: ['date', 'amount', 'source'],
    processData: (data) => [
      new Date(data.date),
      parseFloat(data.amount),
      data.source,
      data.note || '',
      '', '', '', '', ''  // Formula columns
    ],
    formatOptions: {
      centerColumns: [1, 2],
      numberColumns: [3],
      dateColumns: [2]
    },
    successMessage: '✅ Thêm khoản thu thành công!'
  });
}
```

---

## 🛡️ Prevention Guide

### Development Checklist

```markdown
## Before Writing Code

☐ Đọc documentation của project
☐ Hiểu rõ data structure
☐ Check helper functions có sẵn
☐ Review coding standards

## While Writing Code

☐ Validate inputs đầy đủ (client + server)
☐ Dùng helper functions thay vì duplicate code
☐ Add Logger.log() cho debug
☐ Handle errors với try-catch
☐ Return consistent response format:
  {
    success: true/false,
    message: string,
    data: object (optional)
  }

## After Writing Code

☐ Test với data hợp lệ
☐ Test với data không hợp lệ
☐ Test edge cases (empty, null, undefined)
☐ Check execution log
☐ Verify data in sheet
☐ Test performance với dataset lớn

## Before Commit/Deploy

☐ Remove console.log() và debug code
☐ Update documentation
☐ Update CHANGELOG
☐ Test trên staging environment (nếu có)
☐ Backup data trước khi deploy
```

### Code Review Checklist

```markdown
## Data Processing

☐ Dùng findEmptyRow() thay vì getLastRow()
☐ Dùng setValues([array]) thay vì appendRow()
☐ Validate tất cả inputs
☐ Handle null/undefined/empty strings
☐ Format data đúng type (number, date, string)

## Performance

☐ Batch operations (getValues/setValues)
☐ Minimize API calls
☐ Cache sheet references
☐ Avoid loops với getRange() inside

## Error Handling

☐ Try-catch cho tất cả operations
☐ Log errors chi tiết
☐ Return user-friendly messages
☐ Don't expose internal errors

## Forms & UI

☐ Client-side validation (HTML5)
☐ Server-side validation
☐ Loading states
☐ Success/error callbacks
☐ Close dialog sau khi thành công

## Consistency

☐ Naming conventions nhất quán
☐ Parameter names match giữa form & server
☐ Response format nhất quán
☐ Error messages rõ ràng
```

---

## 🧪 Testing Procedures

### Unit Testing Functions

```javascript
/**
 * ===============================================
 * TEST FUNCTIONS
 * ===============================================
 */

function runAllTests() {
  Logger.log('=== RUNNING ALL TESTS ===\n');
  
  testFindEmptyRow();
  testGetNextSTT();
  testValidation();
  testDataInsertion();
  
  Logger.log('\n=== ALL TESTS COMPLETED ===');
}

function testFindEmptyRow() {
  Logger.log('--- Test: findEmptyRow() ---');
  
  const sheet = getSheet('THU');
  const emptyRow = findEmptyRow(sheet, 2);
  
  Logger.log('Empty row found: ' + emptyRow);
  Logger.log('Expected: 2 (if empty) or next available row');
  
  // Verify
  const cellValue = sheet.getRange(emptyRow, 1).getValue();
  if (cellValue === '' || cellValue === null) {
    Logger.log('✅ PASS: Cell is empty');
  } else {
    Logger.log('❌ FAIL: Cell is not empty, value: ' + cellValue);
  }
}

function testGetNextSTT() {
  Logger.log('\n--- Test: getNextSTT() ---');
  
  const sheet = getSheet('THU');
  const nextSTT = getNextSTT(sheet, 2);
  
  Logger.log('Next STT: ' + nextSTT);
  
  // Find actual last STT
  const emptyRow = findEmptyRow(sheet, 2);
  if (emptyRow === 2) {
    // Sheet empty
    if (nextSTT === 1) {
      Logger.log('✅ PASS: Next STT is 1 for empty sheet');
    } else {
      Logger.log('❌ FAIL: Expected 1, got ' + nextSTT);
    }
  } else {
    // Sheet has data
    const lastSTT = sheet.getRange(emptyRow - 1, 1).getValue();
    if (nextSTT === lastSTT + 1) {
      Logger.log('✅ PASS: Next STT is correct');
    } else {
      Logger.log('❌ FAIL: Expected ' + (lastSTT + 1) + ', got ' + nextSTT);
    }
  }
}

function testValidation() {
  Logger.log('\n--- Test: Validation Functions ---');
  
  // Test validateRequiredFields
  const data1 = { date: '2025-01-15', amount: 1000, source: 'Test' };
  const result1 = validateRequiredFields(data1, ['date', 'amount', 'source']);
  Logger.log('Test 1 (valid data): ' + (result1.valid ? '✅ PASS' : '❌ FAIL'));
  
  const data2 = { date: '2025-01-15', amount: 1000 };
  const result2 = validateRequiredFields(data2, ['date', 'amount', 'source']);
  Logger.log('Test 2 (missing field): ' + (!result2.valid ? '✅ PASS' : '❌ FAIL'));
  
  // Test validateNumber
  const numResult1 = validateNumber(1000, 'Amount', { min: 0 });
  Logger.log('Test 3 (valid number): ' + (numResult1.valid ? '✅ PASS' : '❌ FAIL'));
  
  const numResult2 = validateNumber(-100, 'Amount', { min: 0 });
  Logger.log('Test 4 (invalid number): ' + (!numResult2.valid ? '✅ PASS' : '❌ FAIL'));
}

function testDataInsertion() {
  Logger.log('\n--- Test: Data Insertion ---');
  
  const testData = {
    date: '2025-01-15',
    amount: 999999,
    source: 'TEST TRANSACTION',
    note: 'Automated test - Please delete'
  };
  
  const result = addIncome(testData);
  
  if (result.success) {
    Logger.log('✅ PASS: Data inserted at row ' + result.row);
    
    // Verify data
    const sheet = getSheet('THU');
    const insertedNote = sheet.getRange(result.row, 5).getValue();
    if (insertedNote === testData.note) {
      Logger.log('✅ PASS: Data verified correct');
    } else {
      Logger.log('❌ FAIL: Data mismatch');
    }
    
    // Clean up
    Logger.log('⚠️ Remember to delete test row: ' + result.row);
  } else {
    Logger.log('❌ FAIL: ' + result.message);
  }
}
```

### Integration Testing

```javascript
function testFullFlow() {
  Logger.log('=== TESTING FULL FLOW ===\n');
  
  // 1. Test Income
  Logger.log('1. Testing Income Form...');
  const incomeData = {
    date: new Date().toISOString().split('T')[0],
    amount: 5000000,
    source: 'Test Income',
    note: 'Test transaction'
  };
  const incomeResult = addIncome(incomeData);
  Logger.log(incomeResult.message);
  
  // 2. Test Expense
  Logger.log('\n2. Testing Expense Form...');
  const expenseData = {
    date: new Date().toISOString().split('T')[0],
    amount: 500000,
    category: 'Ăn uống',
    note: 'Test expense'
  };
  const expenseResult = addExpense(expenseData);
  Logger.log(expenseResult.message);
  
  // 3. Test Budget Update
  Logger.log('\n3. Testing Budget Update...');
  updateBudgetAfterTransaction('CHI', expenseData.category, expenseData.amount);
  Logger.log('✅ Budget updated');
  
  // 4. Test Dashboard
  Logger.log('\n4. Testing Dashboard Update...');
  updateDashboard();
  Logger.log('✅ Dashboard updated');
  
  Logger.log('\n=== FULL FLOW TEST COMPLETED ===');
}
```

---

## ❓ FAQ Kỹ Thuật

### Q1: Khi nào nên dùng appendRow() vs setValues()?

**A:** 
- **appendRow()**: KHÔNG nên dùng khi sheet có pre-filled formulas
- **setValues()**: Luôn an toàn, nên dùng mặc định
- Best practice: Dùng `findEmptyRow()` + `setValues()`

### Q2: Làm sao để debug scripts trong Google Apps Script?

**A:**
```javascript
// Method 1: Logger.log()
Logger.log('Debug: ' + variable);

// Method 2: View logs
// Apps Script Editor → View → Execution log

// Method 3: Breakpoints (không available trong GAS)
// Thay vào đó: Add nhiều Logger.log()

// Method 4: Try-catch
try {
  // Your code
} catch (error) {
  Logger.log('Error: ' + error.toString());
  Logger.log('Stack: ' + error.stack);
}
```

### Q3: Script chạy chậm, làm sao để tối ưu?

**A:**
1. **Batch operations**: Dùng `getValues()` / `setValues()` thay vì loops
2. **Cache references**: Lưu sheet reference thay vì gọi `getSheet()` nhiều lần
3. **Minimize API calls**: Process trong memory (arrays) thay vì trên sheet
4. **Limit ranges**: Chỉ get/set range cần thiết, không get toàn bộ sheet

### Q4: Làm sao để share project với người khác?

**A:**
1. Share Google Sheet: File → Share → Add people
2. Quyền Editor required để chạy scripts
3. User mới phải authorize script lần đầu
4. Nếu deploy as Web App: Có thể set "Anyone" access

### Q5: Có thể version control Apps Script code không?

**A:**
**Option 1: clasp (Command Line Apps Script)**
```bash
npm install -g @google/clasp
clasp login
clasp clone <scriptId>
# Edit code locally
clasp push
```

**Option 2: Manual backup**
- Copy code vào GitHub/GitLab
- Version trong comment
- Use Git for version control

**Option 3: Apps Script Versions**
- File → Manage versions
- Save version trước mỗi major change

### Q6: Error "Service invoked too many times in a short time"

**A:**
- Đây là rate limiting của Google
- Giảm số lần gọi service trong 1 execution
- Sử dụng batch operations
- Add delays giữa các calls: `Utilities.sleep(1000)`
- Xem quotas: https://developers.google.com/apps-script/guides/services/quotas

---

## 📞 Contact & Support

### Khi nào cần liên hệ support?

1. **Lỗi Critical** - Mất dữ liệu, script không chạy
2. **Bug không có trong guide** - Lỗi mới, chưa documented
3. **Feature request** - Đề xuất tính năng mới
4. **Security issues** - Phát hiện lỗ hổng bảo mật

### Thông tin cần cung cấp khi report bug:

```markdown
**Bug Description:**
Mô tả ngắn gọn lỗi

**Steps to Reproduce:**
1. Bước 1
2. Bước 2
3. Bước 3

**Expected Behavior:**
Kết quả mong đợi

**Actual Behavior:**
Kết quả thực tế

**Screenshots:**
(Nếu có)

**Execution Log:**
(Copy từ View → Execution log)

**Environment:**
- Browser: Chrome/Firefox/Safari
- OS: Windows/Mac/Linux
- Sheet version: v3.0/v3.1/v3.2...

**Additional Context:**
Thông tin thêm (nếu có)
```

### GitHub Issues (Preferred)
```
https://github.com/[username]/HODLVN-Family-Finance/issues
```

### Email Support
```
contact@tohaitrieu.net
```

---

## 📚 Related Documentation

- [README.md](README.md) - Tổng quan hệ thống
- [INSTALLATION.md](INSTALLATION.md) - Hướng dẫn cài đặt
- [USER_GUIDE.md](USER_GUIDE.md) - Hướng dẫn sử dụng
- [TECHNICAL_DOCUMENTATION.md](TECHNICAL_DOCUMENTATION.md) - Tài liệu kỹ thuật
- [CHANGELOG.md](CHANGELOG.md) - Lịch sử phiên bản

---

## 🎓 Learning Resources

### Google Apps Script
- [Official Documentation](https://developers.google.com/apps-script)
- [Best Practices Guide](https://developers.google.com/apps-script/guides/best-practices)
- [Performance Tips](https://developers.google.com/apps-script/guides/support/troubleshooting)

### JavaScript
- [MDN Web Docs](https://developer.mozilla.org/en-US/docs/Web/JavaScript)
- [JavaScript.info](https://javascript.info/)

### Google Sheets
- [Function List](https://support.google.com/docs/table/25273)
- [Keyboard Shortcuts](https://support.google.com/docs/answer/181110)

---

<div align="center">

**Version**: 3.0  
**Last Updated**: 29/10/2025  
**Maintainer**: HODLVN Finance Team

---

**Nếu gặp vấn đề không có trong guide này, vui lòng tạo issue trên GitHub! 🚀**

[⬆ Về đầu trang](#-xử-lý-sự-cố---hodlvn-family-finance)

</div>