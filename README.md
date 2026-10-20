# Lab Grader

Gradescope-style grading for Canvas assignment downloads, built for spreadsheet/doc labs graded on **completeness** and **accuracy** per question (green = full, yellow = partial, red = low effort, black = missing).

Everything runs in the browser. Student files are never uploaded: they're read locally and cached in the browser's IndexedDB so a reload resumes where you left off.

## Use it
1. **Setup**: drop in the Canvas "Download Submissions" `.zip`, the Canvas gradebook export `.csv` (so non-submitters show up), the assignment instructions `.docx` (question labels like `1a.`), and the answer key `.docx`. Void questions or give everyone full credit on a question.
   - **Points** come from two totals (default 10 completeness + 5 accuracy = 15), as in the original gradebook legend (`=10/20`).
     - Full marks = total ÷ graded questions; partial = ½; low effort = ¼; missing = 0.
     - Voiding a question spreads its points over the rest, so the maximum stays the same.
   - **Expected file type** (default `xlsx`): any other format, or more than one file, gets a ⚠ in the student list, panel and file header. The content is still shown.
2. **Grade**: cycle through students and questions. The submission scrolls to the current question, with the key beside it.
   - `1`–`4` set completeness *and* accuracy (full / partial / low / missing).
   - `Shift`+`1`–`4` sets accuracy only.
   - **Excel formulas** for each question are listed beside the key:
     - `✓` when the function matches the key (old names like `NORMSINV` count as `NORM.S.INV`);
     - `≈ key` when the result matches a number in the key;
     - a warning when a value was typed in instead of computed.
     - Click a formula to jump to its cell.
   - `↑`/`↓` changes question, `←`/`→` changes student, and `F` fills the remaining questions with full marks.
3. **Export**:
   - **Gradebook .xlsx**: two rows per student, colored cells, `SUM` formulas.
   - **Canvas CSV**: for Grades → Import. Only fully graded students get a number.

## Develop
No build step. Open `index.html`, or serve the folder. Libraries are vendored in `vendor/` (JSZip, ExcelJS, mammoth).

```sh
node test.js   # logic checks (synthetic data only — never commit real student files)
```
