# Implementation Plan - Luxurious Excel Export for Price Comparison Matrix

Enhance the current Excel download functionality to produce a professional, "luxurious" report with advanced styling, conditional formatting, and summary statistics.

## 1. Dependencies
- Install `exceljs`: For advanced Excel manipulation and styling.
- Install `file-saver`: To handle file saving in the browser.

## 2. Refactor `downloadMatrixAsExcel`
Modify the export logic in `src/App.tsx` to:
- **Initialize Workbook**: Use `ExcelJS.Workbook`.
- **Add Summary Stats**: At the top of the sheet, include global stats like total items, average price across vendors, and overall price variance.
- **Strict Column Widths**: Set column widths to match the UI proportions:
  - Item Name: 30
  - Spec: 20
  - Unit: 10
  - Vendor Columns: 15
- **Header Styling**: Apply a professional dark background (e.g., Slate 800) with white bold text and centered alignment.
- **Data Rows**:
  - Apply clean borders to all cells.
  - Format numeric values as KRW currency (`"₩"#,##0`).
  - **Conditional Formatting**: Iterate through each row and highlight the cell containing the minimum price with a soft emerald background (success indicator).
- **Auto-Filter**: Enable auto-filtering on the header row.

## 3. Verification
- Verify that the exported file opens correctly in Excel/Google Sheets.
- Ensure the minimum price highlighting is accurate per row.
- Confirm the currency formatting and summary stats are present.
- Run `compile_applet` to ensure no build errors.
