# Walkthrough - Dashboard Calendar & Layout Update

I have updated the Dashboard (Tổng quan) to include a "Calendar of Events" table and optimized the layout to fit more information.

## Changes

### 1. Layout Optimization
-   **Income, Expense, Liabilities Tables**: Reduced from 3 columns to 2 columns (Name, Value). Removed the unused "%" column to save space.
-   **Assets Table**: Optimized to fit the new grid layout (Columns D-G).
-   **New Grid**:
    -   **Row 1**: Income (Left), Expense (Middle), Calendar (Right).
    -   **Row 2**: Liabilities (Left), Assets (Right).

### 2. Calendar of Events
-   Added a new table **"Lịch sự kiện (Sắp tới)"** to the right of the Expense table.
-   **Data Source**:
    -   **Debt Payments**: Upcoming due dates from `QUẢN LÝ NỢ` (Status != "Đã thanh toán").
    -   **Lending Collections**: Upcoming due dates from `CHO VAY` (Status == "Đang vay").
-   **Columns**: Ngày (Date), Sự kiện (Event Name), Số tiền (Amount).
-   **Sorting**: Events are sorted by date (nearest first).
-   **Limit**: Shows up to 10 upcoming events.

## Verification

1.  **Open "Tổng quan" Sheet**:
    -   If the sheet is already open, you might need to refresh it.
2.  **Update Dashboard**:
    -   Click the menu **"HODLVN Family Finance"** > **"Cập nhật Dashboard"**.
    -   Or wait for the `onEdit` trigger if you make changes (but manual update is recommended for layout changes).
3.  **Check Layout**:
    -   Verify that **Income** and **Liabilities** are on the left (Columns A-B).
    -   Verify that **Expense** is in the middle (Columns D-E).
    -   Verify that **Calendar of Events** is on the right (Columns G-I).
    -   Verify that **Assets** is below Expense/Calendar (Columns D-G).
4.  **Check Data**:
    -   Ensure the Calendar lists valid upcoming payments/collections if you have data in "QUẢN LÝ NỢ" or "CHO VAY".
    -   Check that the amounts and dates match the source sheets.
