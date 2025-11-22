# Walkthrough - Fix Debt Repayment Calendar Logic

## Changes
I have modified `DashboardManager.gs` to update the `_getCalendarEvents` function.

### Key Improvements
1.  **Monthly Installment Support**:
    *   Previously, the calendar only showed a single "Due Date" event for debts.
    *   Now, for debts with `Term > 1` month, it generates monthly repayment events starting from the `Start Date`.
    *   It projects these events into the future based on the loan term.

2.  **Declining Balance Interest Calculation**:
    *   Previously, interest was hardcoded to `0` or not calculated dynamically.
    *   Now, interest is calculated as: `Remaining Principal * Annual Rate / 12`.
    *   For projected future payments (beyond the immediate next one), the principal balance is simulated to decrease by the monthly principal amount.

3.  **Correct Principal Display**:
    *   Previously, the "Principal" column showed the *entire* remaining balance.
    *   Now, it shows the *monthly* principal portion (`Initial Principal / Term`) for installments.

4.  **Unified Logic**:
    *   Applied the same logic to both **Debt Management** (Quản lý nợ) and **Lending** (Cho vay).

## Verification Results
### Automated Logic Check
*   **Scenario 1: Installment Loan**
    *   Input: Loan 120M, 12 months, 10% rate. Start Jan 1st.
    *   Current Date: June 1st.
    *   Result:
        *   Finds payment dates: June 1st, July 1st, etc.
        *   Calculates Interest on current Remaining Balance.
        *   Projects decreasing interest for July, August, etc.
*   **Scenario 2: Bullet Loan**
    *   Input: Loan 10M, 1 month.
    *   Result: Shows one event at Maturity Date with full Principal and 1-month Interest.

### Manual Verification Required
1.  Open the "Tổng quan" (Dashboard) sheet.
2.  Click "Cập nhật Dashboard" (Update Dashboard) from the menu (if available) or edit a cell to trigger the refresh.
3.  Check the "Lịch sự kiện" (Event Calendar) table.
4.  Verify that installment loans now appear as monthly events with reasonable Interest and Principal amounts.
