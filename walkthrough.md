# Fix Broken Forms Walkthrough

## Issues Identified

1.  **Missing Functions**: `DebtManagementHandler.gs` was missing `updateDebtStatus` and `updateLendingStatus` functions, which were being called by `DataProcessor.gs`. This likely caused the script engine to fail loading `DataProcessor.gs`, breaking all forms dependent on it.
2.  **Incorrect References**: `DataProcessor.gs` was calling `DebtManagementHandler.updateDebtStatus` as if `DebtManagementHandler` was an object, but it was defined as a collection of top-level functions.
3.  **Duplicate Function**: `DashboardRefresh.gs` defined `refreshDashboard`, which conflicted with the menu function of the same name in `Main.gs`.
4.  **Parameter Mismatch**: `StockForm.html` was sending `symbol` instead of `ticker`, which would cause `addStock` to fail validation.

## Changes Made

### 1. `DebtManagementHandler.gs`
-   Added `updateDebtStatus(debtName, principal, interest)` function.
-   Added `updateLendingStatus(borrowerName, principal, interest)` function.

### 2. `DataProcessor.gs`
-   Updated `addDebtPayment` to call `updateDebtStatus` directly (removed `DebtManagementHandler.` prefix).
-   Updated `addLendingRepayment` to call `updateLendingStatus` directly (removed `DebtManagementHandler.` prefix).

### 3. `DashboardRefresh.gs`
-   Renamed `refreshDashboard` to `forceDashboardRecalc` to avoid conflict with `Main.gs`.
-   Updated `triggerDashboardRefresh` to call `forceDashboardRecalc`.

### 4. `StockForm.html`
-   Changed input key from `symbol` to `ticker` to match `DataProcessor.gs` expectation.

### 5. `DashboardFunctions.gs`
-   Rewrote `hffsAssets` function to correctly calculate asset values by aggregating holdings (Buy - Sell) instead of summing transaction costs.
-   Fixed column indices for Stock, Gold, Crypto, and Lending to match the current sheet configuration.

### 6. `DataProcessor.gs` (Additional Fixes)
-   Updated `addDebt` and `addLending` to explicitly set the formula for the "Remaining" (Còn lại) column (Col K) for new rows. This prevents the formula from being overwritten by empty values during data entry.
-   Fixed parameter mismatch in `addDebt`: changed `category` to `source` when calling `addIncome`, which was preventing automatic income entry when adding debt.

### 7. `LendingHandler.gs`
-   Fixed infinite recursion bug: changed `addLending({...})` to `DataProcessor.addLending({...})` at line 56. This was causing "Maximum call stack size exceeded" error when entering lending transactions.

## Verification

-   **All Forms**: Should now load correctly as the script syntax errors/missing references are resolved.
-   **Debt Payment Form**: Should correctly update the "Paid" and "Remaining" columns in "QUẢN LÝ NỢ".
-   **Lending Repayment Form**: Should correctly update the "Paid" and "Remaining" columns in "CHO VAY".
-   **Stock Form**: Should now submit successfully without "Missing information" error.
-   **Dashboard Refresh**: Should work without conflict.

## Next Steps
-   User should reload the spreadsheet to ensure the latest script version is loaded.
-   Test each form to confirm functionality.
