# Payslip Generator (offline)

A single web page that turns an Excel sheet into PDF payslips. Built for employers who
do not have a payroll system: fill a simple Excel template, open the page, upload the
file, check it, pick a design and export all payslips at once (one combined PDF or a
ZIP with one PDF per employee). It handles 1,000+ employees in a few seconds.

**No installation, no sign-in, no internet.** Everything runs inside the browser.
The Excel file never leaves the computer.

## How to use

Open `index.html` in Chrome, Edge or Firefox (double-click it; it works from a USB stick or a shared drive).
The page is a five-step wizard:

1. **Upload** – drop the Excel file. First time? Download the template (or a filled example). Used it before?
   Load the settings file you saved last time.
2. **Review** – one sentence tells you whether the file is ready. Below it:
   - *Company details*: filled from the file's `Settings` sheet, or typed once here if the file has none
     (saved on the computer and in the settings file).
   - *Columns*: how each column of your sheet was understood. Anything guessed or ignored can be remapped
     in one click, and the choice is remembered for the next upload.
   - *Employees*: every row with its calculated earnings, deductions, net and the checks that failed.
3. **Design** – template (Classic, Modern, Minimal), accent colour, paper size, title, logo and stamp,
   and a checklist of what to print, with a live preview. *Save settings file* / *Load settings file* keep
   every payslip uniform across months and computers; *Download my template* gives an Excel file already
   filled with your company details and logo.
4. **Export** – a folder of PDFs (Chrome/Edge, written straight to a folder you pick) or a ZIP, one file per
   employee named `EMP-001_2026-10_Name.pdf` plus a `manifest.csv`; each file is password protected using
   the *Password Rule* (default: the Iqama / National ID number). A combined PDF for printing or
   the archive is also available.
5. **Send** – Outlook is the default (and for now the only visible) channel: press **Next draft** or the Enter
   key, classic Outlook opens a ready email with the locked PDF attached, you press Send. The other channels
   below are still built in but hidden; open the page with `?channels=all` to show them again
   (`?channels=outlook` hides them). Nothing is sent by the page itself:
   - **WhatsApp**: one button per employee opens the chat with the message typed; on an HTTPS page in
     Chrome/Edge/Safari a *Share PDF* button hands the file itself to WhatsApp through the system share sheet.
   - **Outlook (Windows)**: one `.eml` per employee that classic Outlook opens as a ready draft with the PDF
     attached, plus a pack with batches of 50, a guide and a macro that saves every draft in one run.
   - **Gmail / any mailbox**: a Mail Merge pack (PDFs + `mailmerge.csv` + guide) for the free Thunderbird
     Mail Merge add-on, which sends every payslip with its own attachment in one run.
   - **Print & sign**: an acknowledgement register to collect signatures.
   Rows are ticked as you go; the ticks survive reloads and export as a delivery log CSV with file hashes.

The research behind the delivery design is in `reports/Offline bulk payslip delivery options.md`, and the
build plan in `docs/DELIVERY_PLAN.md`.

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
- Delivery columns (optional): **Mobile** (any format; `05…` becomes `9665…`), **Email**, **Date of Birth**
  (used for the PDF password). Delivery settings on the `Settings` sheet: Sender Email, Message Subject,
  Message Text (with `{name}` `{period}` `{company}` `{net}` `{password_hint}` placeholders), Password Rule
  (`id` default, `id+dob`, `dob`, `custom`, `none`) and Default Country Code.
- Names and text can be in English or Arabic. Arabic is printed with the bundled Amiri font (SIL Open Font License);
  other scripts (Cyrillic, Chinese…) are flagged in the checks and print as `?`.

Ready-made files are in `samples/`: the empty template, a filled sample with 25 employees, one sample PDF per design and an Arabic sample per design.

## Project layout

```
index.html        the page
css/styles.css
js/schema.js      Excel template definition, tolerant parser, column guessing and validation (pure functions)
js/xlsx-images.js reads pictures out of an .xlsx and embeds pictures into one (JSZip)
js/pdf.js         the three PDF designs, display options, PDF passwords, acknowledgement register (jsPDF)
js/delivery.js    .eml drafts, mail-merge CSV, WhatsApp/Gmail links, Outlook macro and guides
js/fonts/         Amiri Regular/Bold as base64 for Arabic text, plus its OFL licence
js/app.js         the five-step UI, settings file, preview, export and send
js/art.js         logo, the mascot "Qasi" and its moods, confetti, count-up numbers, card spotlight
css/fonts/        Sora display font (SIL Open Font License)
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

- Nothing can be sent from a web page without a click per employee. WhatsApp needs one click (plus picking
  the contact in the share sheet, or dragging the file); the only zero-click route is the Thunderbird Mail
  Merge pack, which needs a one-time install.
- The `.eml` drafts open as editable drafts only in classic Outlook for Windows (build 2109 or newer).
  New Outlook, Outlook for Mac, Apple Mail and web mail open them read-only.
- PDF passwords use the 40-bit RC4 standard security handler built into jsPDF: it stops casual opening and
  works in every viewer, but it is not strong encryption. The default password is the Iqama number for
  simplicity; switch the Password Rule to `id+dob` for a harder one. AES-256 is a planned upgrade.
- The folder export and the Share PDF button are Chrome/Edge (and Safari for share) features; Firefox users
  get the ZIP and the "Open chat" button.
- Arabic is supported for values (names, titles, company name, custom pay lines, remarks, footer); the fixed
  labels and the amount in words stay in English.
- A ZIP of 1,000 Arabic payslips takes around 1–2 minutes because each file carries its own font subset.
- Pictures in formats the browser cannot decode (EMF/WMF) are ignored; use PNG or JPG in the Excel file.

## Look and feel

The theme is driven by four colour variables at the top of `css/styles.css` (`--c1` to `--c4`): the moving
aurora background, buttons, logo accents and the mascot follow them. Animations respect the operating system's
"reduce motion" setting. Ideas for what to build next are in `docs/IDEAS.md`.
