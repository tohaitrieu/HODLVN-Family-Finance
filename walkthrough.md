# Walkthrough - Event Calendar Logic Refinement (v3.5.5)

I have refined the logic for the Event Calendar (Phải trả/Phải thu) on the Dashboard to strictly follow your requirements.

## Changes

### 1. Refined Interest Calculation Logic
- **Installment Loans (Vay trả góp)**:
  - **Principal**: `Total Principal / Term` (unchanged).
  - **Interest**: Now uses the **actual remaining principal** from the sheet (Column K) instead of a simulated schedule. This ensures accuracy even if payments are irregular.
  - **Days Calculation**: Now uses the **number of days in the month of the event date** (e.g., 30 for Sept, 31 for Oct) instead of the exact days elapsed between payments.
  
- **Interest Only Loans (Trả lãi hàng tháng)**:
  - **Interest**: Now uses the **number of days in the month of the event date**.
  - **Principal**: 0 until the last period.

### 2. Single Event Display
- The logic now strictly ensures that **only the single nearest upcoming event** is displayed for each loan/debt.
- Loops break immediately after finding the first future payment date.

### 3. Files Updated
- `CustomFunctions.gs`: Core logic update.
- `CHANGELOG.md`: Documentation.
- `ChangelogData.gs`: Internal version tracking.
- `Main.gs`: Version bump to 3.5.5.

## Verification
- **Check Dashboard**: The "Lịch sự kiện" tables should now show more accurate interest amounts matching your formula: `DaysInMonth * Remaining * Rate / 365`.
- **Installment Loans**: Should show interest based on the current "Gốc còn lại" in the sheet.
