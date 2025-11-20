# Walkthrough - Remove Currency Units

## Changes Made

### 1. Removed "VNĐ" from Number Formats
- **File:** `Main.gs`, `DataProcessor.gs`
- **Action:** Updated `APP_CONFIG.FORMATS.NUMBER` and all hardcoded formats to use `'#,##0'` instead of `'#,##0" VNĐ"'`.
- **Result:** Numbers in sheets will display as plain numbers (e.g., `1,000,000`) without the " VNĐ" suffix.

### 2. Updated HTML Forms
- **Files:** `IncomeForm.html`, `ExpenseForm.html`, `DebtPaymentForm.html`, `StockForm.html`, `GoldForm.html`, `CryptoForm.html`, `SetBudgetForm.html`, `SetupWizard.html`, `OtherInvestmentForm.html`, `DebtManagementForm.html`
- **Action:** Removed "(VNĐ)" from labels and " VNĐ" from JavaScript display logic.
- **Result:** Forms now show cleaner labels (e.g., "Số tiền" instead of "Số tiền (VNĐ)") and input values.

### 3. Updated Utility Functions
- **File:** `Utils.gs`
- **Action:** Updated `formatCurrency` to return a decimal string without the currency symbol or unit.

## Verification Steps

### 1. Verify Sheet Formatting
1.  Go to **Menu > Khởi tạo Sheet > 📥 Khởi tạo TẤT CẢ Sheet** (or individual sheets).
2.  Enter new data via any form (e.g., Income, Expense).
3.  Check the spreadsheet columns. Values should be formatted as `1,000,000` (no VNĐ).

### 2. Verify Forms
1.  Open any form (e.g., **Menu > Nhập liệu > ➕ Thu nhập**).
2.  Check the "Số tiền" label. It should NOT say "(VNĐ)".
3.  Enter a value. The display (if any) should not append "VNĐ".
