# Implementation Plan - Fix Dashboard Formula Syntax Error

## Problem
The `_createDynamicSumFormula` helper function in `DashboardManager.gs` returns a formula string starting with `=`. When these results are concatenated (e.g., `FormulaA + FormulaB`), it results in invalid syntax like `=IFERROR(...) + =IFERROR(...)`.

## Proposed Changes

### `DashboardManager.gs`

#### [MODIFY] `_createDynamicSumFormula`
- Remove the leading `=` from the return string.
- The function will now return just the expression: `IFERROR(...)`.

#### [MODIFY] `_setupGridTables`
- Prepend `=` to `setFormula` calls where `_createDynamicSumFormula` is the start.
- Example: `.setFormula('=' + this._createDynamicSumFormula(...))`
- For concatenations: `.setFormula('=' + this._createDynamicSumFormula(...) + ' + ' + this._createDynamicSumFormula(...))`

#### [MODIFY] `_createChart`
- Prepend `=` to `setFormula` calls.
- Ensure concatenated formulas in `AA3`, `AA4` are constructed correctly starting with `=`.

### `SheetInitializer.gs`

#### [MODIFY] `_fixDateColumn`
- Update logic to explicitly strip time components from Date objects.
- Ensure `setNumberFormat(APP_CONFIG.FORMATS.DATE)` is applied.
- Logic: `new Date(date.getFullYear(), date.getMonth(), date.getDate())`.

## Verification Plan

### Automated Verification
- None.

### Manual Verification
1.  **Action:** Run **Menu > Thống kê & Dashboard > 🔄 Cập nhật Dashboard**.
2.  **Check:** Verify no formula parse errors.
3.  **Action:** Enter a date like "10/11/2025 17:00:00" in the "THU" sheet (Column B).
4.  **Action:** Run **Menu > Khởi tạo Sheet > 📥 Khởi tạo Sheet THU**.
5.  **Check:** Verify the cell value becomes "10/11/2025" (time removed) and format is correct.
