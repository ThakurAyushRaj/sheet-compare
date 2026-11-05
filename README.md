# SheetCompare

A React, TypeScript, and Tailwind CSS app for comparing Excel workbooks locally in the browser.

Run `npm install` and `npm run dev`. Create a production build with `npm run build`.

Upload an original and updated `.xlsx`, `.xls`, or `.csv` file, or select the sample comparison. Results show changed, added, and removed cells across sheets, with sheet and change filters and CSV export.

Comparison matches sheet names and cell addresses, and compares saved cell values. It does not detect formatting differences or formula changes that produce the same saved value. Rows are compared by position; inserted rows can therefore generate multiple differences. Files are never uploaded.

Run `node check-comparison.mjs` to verify the comparison cases.
