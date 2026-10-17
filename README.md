# Payslip Generator (offline)

A single web page that turns an Excel sheet into PDF payslips. Built for employers who
do not have a payroll system: fill a simple Excel template, open the page, upload the
file, check it, pick a design and export all payslips at once (one combined PDF or a
ZIP with one PDF per employee). It handles 1,000+ employees in a few seconds.

**No installation, no sign-in, no internet.** Everything runs inside the browser.
The Excel file never leaves the computer.

## How to use

Open `index.html` in Chrome, Edge or Firefox (double-click it; it works from a USB stick or a shared drive).
The page is a four-step wizard:

1. **Upload** – drop the Excel file. First time? Download the template (or a filled example). Used it before?
   Load the settings file you saved last time.
2. **Review** – one sentence tells you whether the file is ready. Below it:
   - *Company details*: filled from the file's `Settings` sheet, or typed once here if the file has none
     (saved on the computer and in the settings file).
   - *Columns*: how each column of your sheet was understood. Anything guessed or ignored can be remapped
     in one click, and the choice is remembered for the next upload.
   - *Employees*: every row with its calculated earnings, deductions, net and the checks that failed.
3. **Design** – template (Classic, Modern, Minimal), accent colour, paper size, title, logo and stamp,
   and a checklist of what to print (CR number, IBAN, attendance, amount in words, signature lines…),
   with a live preview. *Save settings file* / *Load settings file* keep every payslip uniform across
   months and computers; *Download my template* gives an Excel file already filled with your company
   details and logo.
4. **Export** – one combined PDF, or a ZIP with one PDF per employee. 1,000 employees take a few seconds.

## What the file can look like

The template gives the best result, but any salary sheet is accepted:

- One row per employee. Required: a name column and at least one salary column. Everything else is optional.
- Headers are matched exactly, by alias (*Iqama*, *HRA*, *Net Payable*…) or by keyword; a merged
  "Earnings / Deductions" title row above the headers is used to classify unknown pay columns.
- Title rows above the header, totals rows, notes rows and the template's example row are skipped.
- No Employee ID column → IDs are generated (EMP-001…). No Settings sheet → company details are typed on the page.
- Extra pay items: a column named `Earning: <anything>`; extra deductions: `Deduction: <anything>`.
- Amounts may contain thousands separators or currency text. Net/Gross/Total Deductions columns in the file
  are only cross-checked and a warning is shown on mismatch.
- Logo: insert the picture on the `Settings` sheet (Insert → Pictures); a second picture becomes the stamp/signature.
  Pictures are read straight out of the .xlsx.
- Use English / Latin letters. Arabic script is not supported by the built-in PDF fonts in this version.

Ready-made files are in `samples/`: the empty template, a filled sample with 25 employees and one sample PDF per design.

## Project layout

```
index.html        the page
css/styles.css
js/schema.js      Excel template definition, tolerant parser, column guessing and validation (pure functions)
js/xlsx-images.js reads pictures out of an .xlsx and embeds pictures into one (JSZip)
js/pdf.js         the three PDF designs and the display options (jsPDF)
js/app.js         the four-step UI, settings file, preview and export
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
- Pictures in formats the browser cannot decode (EMF/WMF) are ignored; use PNG or JPG in the Excel file.
- One payslip per page, A4 or Letter, portrait.
