# Lab Grader

Gradescope-style grading for Canvas assignment downloads, built for spreadsheet/doc labs graded on **completeness** and **accuracy** per question (green = full, yellow = partial, red = low effort, black = missing).

Everything runs in the browser. Student files are never uploaded: they're read locally and cached in the browser's IndexedDB so a reload resumes where you left off.

## Use it
1. **Setup**: drop in the Canvas "Download Submissions" `.zip`, the Canvas gradebook export `.csv` (so non-submitters show up), the assignment instructions `.docx` (question labels like `1a.`), and the answer key `.docx`. Void questions or give everyone full credit on a question.
2. **Grade**: cycle through students and questions. The submission scrolls to the current question, with the key beside it.
   - `1`–`4` set completeness *and* accuracy (full / partial / low / missing).
   - `Shift`+`1`–`4` sets accuracy only.
   - `↑`/`↓` changes question, `←`/`→` changes student, and `F` fills the remaining questions with full marks.
3. **Export**:
   - **Gradebook .xlsx**: two rows per student, colored cells, `SUM` formulas.
   - **Canvas CSV**: for Grades → Import. Only fully graded students get a number.

## Develop
No build step. Open `index.html`, or serve the folder. Libraries are vendored in `vendor/` (JSZip, ExcelJS, mammoth).

```sh
node test.js   # logic checks (synthetic data only — never commit real student files)
```
