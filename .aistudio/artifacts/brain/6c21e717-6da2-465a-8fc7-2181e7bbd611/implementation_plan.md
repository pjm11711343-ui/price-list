# Implementation Plan - Add Maker Name to Comparison Matrix

Add the product maker (제조사) information to the comparison matrix view and Excel export to provide better context for unit price comparisons.

## 1. Update Data Processing
- Modify the `matrixData` useMemo hook in `src/App.tsx`.
- Include the `maker` field from the `PriceItem` into the grouped `prices` object for each vendor.

## 2. Enhance Matrix UI
- Locate the price cell rendering in the comparison matrix table.
- Update the markup below the unit price to display the maker name.
- Styling:
  - Font: Smaller (text-[7px] or text-[8px]).
  - Color: Subdued (text-slate-400).
  - Layout: Displayed next to or below the nego rate with wrapping enabled (`whitespace-normal`).

## 3. Update Excel Export
- Modify `downloadMatrixAsExcel` to include the maker name in the vendor price cells.
- The format will be: `₩1,200 (Maker Name)` or a similar professional layout within the cell.
- Ensure the Excel cell styling accommodates the additional text.

## 4. Verification
- Verify the maker name appears correctly in the Matrix UI.
- Test with short and long maker names to ensure wrapping works as intended.
- Export to Excel and verify the maker name is present and formatted professionally.
