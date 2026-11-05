# SheetCompare

A React, TypeScript, and Tailwind CSS app for comparing Excel workbooks locally in the browser.

Run `npm install` and `npm run dev`. Create a production build with `npm run build`.

Upload an original and updated `.xlsx`, `.xls`, or `.csv` file, or select the sample comparison. Results show only cells containing different non-empty values in both files, with a sheet filter and CSV export. Added and removed cells are excluded.

Rows are matched by Unit No., regardless of row order. The app compares Status, Unit Description, Plot Size, Width, Length, and Dimension Unit. Other columns, including IDs, prices, and timestamps, are ignored. Results identify affected units, fields, and cell addresses in both files. Duplicate unit numbers produce a warning.

Single-sheet files compare directly even when filenames and sheet names differ. Multi-sheet workbooks match sheets by name. Columns match by normalized header name, so their order can differ. Only different non-empty values for units present in both files appear; added and removed units or blank cells are excluded.

The app compares saved cell values, including decimal differences such as 229.44 to 229.5. It does not detect formatting differences or formula changes that produce the same saved value. Files are never uploaded. Save and reselect a file after editing it in Excel.

Run `node check-comparison.mjs` to verify the comparison cases.

