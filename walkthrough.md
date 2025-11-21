# Walkthrough - Add Lending Repayment Sheet

I have added the functionality to initialize a new sheet called "THU NỢ" (Lending Repayment), which is related to the "CHO VAY" (Lending) sheet. This addresses the missing menu item for initializing the sheet related to lending.

## Changes

### 1. `Main.gs`

-   **Updated `APP_CONFIG`**: Added `LENDING_REPAYMENT: 'THU NỢ'` to the `SHEETS` configuration.
-   **Updated Menu**: Added "💰 Khởi tạo Sheet THU NỢ" to the "Khởi tạo Sheet" submenu.
-   **Added Wrapper Function**: Added `initializeLendingRepaymentSheet` function to call the initializer.
-   **Updated `processSetupWizard`**: Included `initializeLendingRepaymentSheet(true)` to ensure the sheet is created during full system setup.

### 2. `SheetInitializer.gs`

-   **Implemented `initializeLendingRepaymentSheet`**:
    -   Creates the "THU NỢ" sheet if it doesn't exist.
    -   Sets up headers: `STT`, `Ngày`, `Người vay`, `Thu gốc`, `Thu lãi`, `Tổng thu`, `Ghi chú`.
    -   Applies formatting (column widths, number formats, date formats).
    -   Adds a formula for "Tổng thu" (`=Thu gốc + Thu lãi`).

## Verification

1.  **Reload the Spreadsheet**: Refresh the page to see the updated menu.
2.  **Check Menu**: Go to `HODLVN Family Finance` > `⚙️ Khởi tạo Sheet`. You should see `💰 Khởi tạo Sheet THU NỢ`.
3.  **Run Initialization**: Click the menu item. It should create a new sheet named "THU NỢ" with the correct structure.
4.  **Check Setup Wizard**: Running "Khởi tạo TẤT CẢ Sheet" will now also create the "THU NỢ" sheet.
