# Walkthrough - Refactor Event Calendar

## Changes
I have refactored the "Lịch sự kiện" (Event Calendar) in `DashboardManager.gs` to provide a more detailed view of upcoming financial obligations.

### Key Improvements
1.  **New Column Structure (6 Columns)**:
    *   **Ngày** (Date): Bold formatted.
    *   **Hành động** (Action): "Phải trả" (Payable) or "Phải thu" (Receivable).
    *   **Sự kiện** (Event): Name of the debt or loan (e.g., "Trả nợ: Vay mua nhà (Kỳ 1/12)").
    *   **Gốc còn lại** (Remaining Principal): Shows the outstanding balance *before* the payment.
    *   **Gốc trả kỳ này** (Principal Payment): The principal amount due for this specific event.
    *   **Lãi trả kỳ này** (Interest Payment): The interest amount due, calculated based on the declining balance.

2.  **Enhanced Logic**:
    *   **Installment Debts**:
        *   Generates monthly events based on the Start Date and Term.
        *   Calculates interest using `Days * (Rate * Remaining) / 360`.
        *   Simulates the declining principal balance for future projections.
    *   **Lending (Cho vay)**:
        *   Treated as "Bullet" repayment (one-time collection at maturity).
        *   Shows "Phải thu" action.
        *   Interest is calculated for the full duration (Start to Maturity).
    *   **Robust Data Parsing**:
        *   **Currency**: Added `parseCurrency` to correctly handle formatted numbers.
        *   **Date**: Added `parseDate` to correctly handle date strings.
    *   **Status Tracking**:
        *   **Debts**: Now explicitly tracks both **"Chưa trả"** and **"Đang trả"** statuses.
        *   **Lending**: Tracks "Đang vay".

3.  **Total Row**:
    *   Added a summary row at the bottom of the calendar table.
    *   Sums up the total "Principal Payment" and "Interest Payment" for the displayed events.

## Verification Results
### Automated Logic Check
*   **Scenario 1: Installment Loan**
    *   Input: Loan 120M, 12 months, 10% rate.
    *   Result:
        *   Shows 12 monthly events.
        *   "Gốc trả kỳ này" = 10M/month.
        *   "Lãi trả kỳ này" decreases each month as "Gốc còn lại" decreases.
*   **Scenario 2: Lending**
    *   Input: Lend 50M, 6 months, 12% rate.
    *   Result:
        *   Shows 1 event at Maturity Date.
        *   Action: "Phải thu".
        *   "Gốc trả kỳ này" = 50M.
        *   "Lãi trả kỳ này" = Interest for 6 months.

### Manual Verification Required
1.  Open the "Tổng quan" (Dashboard) sheet.
2.  Trigger a dashboard update (Menu > Cập nhật Dashboard).
3.  Inspect the "Lịch sự kiện" table (Columns K-P).
4.  Verify the columns match the new structure.
5.  Check if the calculations for Principal and Interest look correct for your data.
6.  Verify the Total row at the bottom.
