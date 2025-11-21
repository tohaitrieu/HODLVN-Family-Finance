# Walkthrough - New Features: Lending, Debt Schedule, Data Normalization

I have implemented the following features as requested:

## 1. Quản lý Cho vay (Lending Management)

### ✨ Features
- **New Sheet**: `CHO VAY` (Lending) to track loans given to others.
- **Add Loan**: Form to add new loans. Automatically creates an "Expense" entry (Category: "Cho vay") in the `CHI` sheet to reflect cash outflow.
- **Receive Payment**: Form to record Principal and Interest repayment. Automatically creates "Income" entries (Category: "Thu hồi nợ" / "Lãi đầu tư") in the `THU` sheet.
- **Dashboard Integration**: "Cho vay" is now listed in the **Assets** (Tài sản) section of the Dashboard.

### 📂 Files Created/Modified
- `SheetInitializer.gs`: Added `initializeLendingSheet`.
- `LendingForm.html`: Form for adding loans.
- `LendingPaymentForm.html`: Form for receiving payments.
- `LendingHandler.gs`: Logic for adding loans and processing payments.
- `Main.gs`: Added menu items and configuration.
- `DashboardManager.gs`: Added "Cho vay" to Assets table.

## 2. Lịch Trả Nợ (Debt Repayment Schedule)

### ✨ Features
- **Report**: A new report showing upcoming debt payments for the next month.
- **Calculation**: Uses the "Reducing Balance" method (Dư nợ giảm dần) to estimate the next payment (Principal + Interest).
- **Access**: Menu > **Thống kê & Dashboard** > **Lịch trả nợ dự kiến**.

### 📂 Files Created/Modified
- `DataNormalizer.gs`: Implemented `calculateNextDebtPayments` and `showDebtScheduleReport`.
- `Main.gs`: Added menu item.

## 3. Chuẩn hóa Dữ liệu (Data Normalization)

### ✨ Features
- **One-click Fix**: A tool to scan all sheets and normalize Date formats (dd/mm/yyyy) and Number formats.
- **Access**: Menu > **Tiện ích** > **Chuẩn hóa dữ liệu**.

### 📂 Files Created/Modified
- `DataNormalizer.gs`: Implemented `normalizeAllData`.
- `Main.gs`: Added menu item.

## 🚀 How to Use

1.  **Initialize**: Go to **Menu > Khởi tạo Sheet > Khởi tạo Sheet CHO VAY** (if not already done).
2.  **Lending**:
    -   **Add Loan**: Menu > **Nhập liệu > Cho vay**.
    -   **Receive Payment**: Menu > **Nhập liệu > Thu nợ & Lãi**.
3.  **Debt Schedule**: Menu > **Thống kê & Dashboard > Lịch trả nợ dự kiến**.
4.  **Normalize Data**: Menu > **Tiện ích > Chuẩn hóa dữ liệu**.

## 4. Cập nhật Giá Chứng Khoán (Stock Price Update)

### ✨ Features
- **Automatic Price Update**: The "Giá HT" (Current Price) column in the `CHỨNG KHOÁN` sheet now automatically updates using the `MPRICE` custom function.
- **Financial Functions**: Added a suite of financial functions (`TCBS_BARS`, `HODLDATA`, `PIVOTFIB`, `ATR`, `STOCHASTIC`, `RSI`, `EMA`, `MACD`, `AVERAGE_DOWN`) for advanced analysis.

### 📂 Files Created/Modified
- `StockFunctions.gs`: Created to house `MPRICE` and other financial functions.
- `SheetInitializer.gs`: Updated `initializeStockSheet` to apply the `MPRICE` formula to the "Giá HT" column.
