# Walkthrough - Refactor Event Calendar

## Changes
I have refactored the "Lịch sự kiện" (Event Calendar) in `DashboardManager.gs` to provide a more detailed and organized view.

### Key Improvements
1.  **Split Tables**:
    *   **KHOẢN PHẢI TRẢ (Sắp tới)**: Displays the top 10 upcoming debt payments.
    *   **KHOẢN PHẢI THU (Sắp tới)**: Displays the top 10 upcoming lending collections.
    *   Each table has its own "TỔNG CỘNG" row.

2.  **New Column Structure (6 Columns)**:
    *   **Ngày** (Date): Bold formatted.
    *   **Hành động** (Action): "Phải trả" or "Phải thu".
    *   **Sự kiện** (Event): Name of the debt or loan.
    *   **Gốc còn lại** (Remaining Principal): Outstanding balance before payment.
    *   **Gốc trả kỳ này** (Principal Payment): Amount due.
    *   **Lãi trả kỳ này** (Interest Payment): Interest due.

3.  **Enhanced Logic**:
    *   **Installment Debts**: Monthly events, declining balance interest.
    *   **Lending**: Bullet repayment at maturity.
    *   **Robust Parsing**: Correctly handles formatted currency and date strings.
    *   **Status Tracking**: Includes "Chưa trả", "Đang trả" (Debts) and "Đang vay" (Lending).

## Verification Results
### Automated Logic Check
*   **Scenario 1: Mixed Debts and Lending**
    *   Input: 1 Installment Loan, 1 Lending.
    *   Result:
        *   **Table 1 (Payables)**: Shows monthly installments for the loan.
        *   **Table 2 (Receivables)**: Shows the maturity event for the lending.

### Manual Verification Required
1.  Open the "Tổng quan" (Dashboard) sheet.
2.  Trigger a dashboard update (Menu > Cập nhật Dashboard).
3.  Inspect the "Lịch sự kiện" area (Columns K-P).
4.  Verify there are **TWO** separate tables.
5.  Check if the data in each table corresponds to the correct category (Payable vs Receivable).
6.  Verify the Total rows for each table.
