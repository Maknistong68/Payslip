# Payslip Generator (offline)

A single web page that turns an Excel sheet into PDF payslips. Built for employers who
do not have a payroll system: fill a simple Excel template, open the page, upload the
file, check it, pick a design and export all payslips at once (one combined PDF or a
ZIP with one PDF per employee). It handles 1,000+ employees in a few seconds.

**No installation, no sign-in, no internet.** Everything runs inside the browser.
The Excel file never leaves the computer.

## How to use

1. Open `index.html` in Chrome, Edge or Firefox (double-click it; it works from a USB stick or a shared drive).
2. **Step 1** – click *Download template (.xlsx)*.
   - `Settings` sheet: company name, CR number, address, pay period, payment date, currency, footer note…
   - `Payslips` sheet: one row per employee (ID, name, Iqama, job title, bank/IBAN, attendance, earnings, deductions).
   - `Instructions` sheet: the same rules in short form.
3. **Step 2** – upload the filled file. The page checks it and lists every problem with its row number
   (missing names/IDs, duplicated IDs, text instead of numbers, negative amounts, deductions bigger than
   earnings, suspicious IBANs, net salary that does not match, Arabic text that the PDF fonts cannot print…).
4. **Step 3** – add your logo (and optionally a stamp/signature image), pick one of the designs
   (Classic, Modern, Minimal), accent colour, paper size. These are remembered on the computer.
5. **Step 4** – preview any employee and click *Download all as one PDF* or *Download ZIP (one PDF per employee)*.

Next month: change the pay period and the amounts in the same Excel file and upload it again.

## Template rules

- Keep the header row of the `Payslips` sheet. Required columns: **Employee ID**, **Employee Name**, **Basic Salary**.
  Everything else is optional and unused columns can be deleted.
- Amounts must be numbers; an empty cell means zero. Net salary is calculated (earnings − deductions).
  If you fill the optional *Net Salary* column it is only cross-checked and a warning is shown on mismatch.
- Extra pay items: add a column named `Earning: <anything>` (e.g. `Earning: Site Allowance`).
  Extra deductions: `Deduction: <anything>` (e.g. `Deduction: Housing Loan`). They are printed under that name.
- Header names are matched loosely (case, spaces and common aliases such as *Iqama*, *Basic*, *HRA*, *Net Pay* are accepted).
- Use English / Latin letters. Arabic script is not supported by the built-in PDF fonts in this version.

Ready-made files are in `samples/`: the empty template, a filled sample with 25 employees and one sample PDF per design.

## Project layout

```
index.html        the page
css/styles.css
js/schema.js      Excel template definition, parser and validation (pure functions)
js/pdf.js         the three PDF designs (jsPDF)
js/app.js         UI
js/vendor/        SheetJS, jsPDF, JSZip (bundled so the page works offline)
samples/          template, sample workbook, sample PDFs
tools/            build-samples.js (regenerates samples/) and test.js (smoke tests)
```

## Development

```
node tools/test.js            # parser/validator/PDF smoke tests, also renders 1,000 payslips
node tools/build-samples.js   # regenerate samples/*.xlsx
WRITE_PDFS=1 node tools/test.js   # also writes samples/Sample_*.pdf
```

No build step; edit the files and reload the page.

## Known limitations

- Latin script only in PDFs (Arabic names/labels show as `?`); a bilingual template needs an embedded Arabic font.
- The logo must be uploaded on the page (Excel images are not read).
- One payslip per page, A4 or Letter, portrait.
