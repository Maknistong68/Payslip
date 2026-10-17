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

1. **Upload** – drop the Excel file (or CSV, or paste the cells straight from Excel). First time? Download the
   template (or a filled example). Used it before? Load the settings file you saved last time.
2. **Review** – one sentence tells you whether the file is ready. Below it:
   - *Source bar* (only when needed): which sheet to read (or all sheets combined), which row holds the
     headers, whether to include hidden rows, and a badge when a saved column setup was applied.
   - *Totals*: earnings, deductions and net to pay, with a box for the bank / WPS transfer total that shows
     ✓ or the exact difference.
   - *Company details*: filled from the file's `Settings` sheet, or typed once here if the file has none
     (saved on the computer and in the settings file).
   - *Columns*: how each column of your sheet was understood. Anything guessed or ignored can be remapped
     in one click, and the choice is remembered for the next upload.
   - *Employees*: every row with its calculated earnings, deductions, net and the checks that failed.
3. **Design** – one of 13 templates (Classic, Modern, Minimal, Executive, Steel, Corporate, Bold, Ledger,
   Compact, Split, Formal, Heritage, Receipt), accent colour, paper size, title, logo and stamp,
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

### First visit, save point and never losing work

- **First visit** (nothing saved in this browser yet): a 30-second animated quick start opens by itself:
  download the template, fill it, drop it, save the settings file. It can be replayed from *▶ 30-second guide*.
  Returning visitors get a *Welcome back* card with their company and backup status instead.
- **Save point**: before the Export step, if the setup (company, logo, design, columns, message) has no
  up-to-date settings file, an animated save point offers to save one. Skipping is remembered until something
  changes. The top bar shows the backup state at all times (*Save settings* / *Save changes* / *Backed up*).
  In Chrome and Edge the file is picked once and then **kept up to date automatically**; other browsers download it.
  The pay period and dates do not count as changes, so the monthly routine does not nag.
- **Refresh, crash or browser restart**: the open file, the step, the sheet and header choices, the export
  options and the bank total are kept in the browser (IndexedDB). Within 12 hours the page reopens exactly where
  it was; later, a *Pick up where you left off* card offers to resume (kept up to 30 days).
- **Sending is crash-proof**: every tick is written before the next click (localStorage, with a copy in
  IndexedDB) into an append-only journal, so two tabs or a restored session never erase each other's ticks.
  After a folder export the journal is also written to `delivery-log.csv` next to the PDFs, and *Restore ticks
  from a log* reads it back after a browser reset or on another computer. Sending to someone already ticked asks
  first; a second open tab shows a warning; ticks made with a different version of the Excel file are flagged;
  closing the page during an export asks for confirmation.
- **Your data** (the *Private* button): what this browser keeps and why, a switch to stop remembering the open
  file on a shared computer, and *Erase everything here*.

The research behind the delivery design is in `reports/Offline bulk payslip delivery options.md`, and the
build plan in `docs/DELIVERY_PLAN.md`.

## What the file can look like

The template gives the best result, but any salary sheet is accepted:

- One row per employee. Required: a name column and at least one amount (a Net Salary column alone is enough).
  Everything else is optional. First Name / Last Name columns are joined.
- Several sheets: sheets with the same layout (one per site) are combined; one sheet per month picks the
  pay period's month (or the latest). Hidden sheets, hidden rows and sheets named Lists/Lookup are skipped.
- The header row is searched in the first 200 rows and can be picked by hand.
- Headers are matched exactly, by alias (*Iqama*, *HRA*, *Net Payable*…) or by keyword; a merged
  "Earnings / Deductions" title row above the headers is used to classify unknown pay columns.
- Title rows above the header, totals rows, notes rows and the template's example row are skipped.
  Department heading rows ("PROJECTS") become the department of the rows below them.
- No Employee ID column → IDs are generated (EMP-001…). No Settings sheet → company details are typed on the page.
- Extra pay items: a column named `Earning: <anything>`; extra deductions: `Deduction: <anything>`.
- Amounts may contain thousands separators, currency text, Arabic-Indic digits (`٣٠٠٠`) or a decimal comma
  (`3.000,50`, detected per column). Deductions typed as negatives are accepted. Excel error cells (`#N/A`,
  `#REF!`) are flagged instead of read as zero. Net/Gross/Total Deductions columns in the file are only
  cross-checked and a warning is shown on mismatch.
- Dates: day-first by default; a column that is clearly month-first (`03/15/2024`) is read that way and flagged,
  impossible dates are errors.
- Headers can be English or Arabic (`الاسم`, `الراتب الأساسي`, `بدل السكن`…).
- Logo: insert the picture on the `Settings` sheet (Insert → Pictures); a second picture becomes the stamp/signature.
  Pictures are read straight out of the .xlsx.
- Delivery columns (optional): **Mobile** (any format; `05…` becomes `9665…`), **Email**, **Date of Birth**
  (used for the PDF password). Delivery settings on the `Settings` sheet: Sender Email, Message Subject,
  Message Text (with `{name}` `{period}` `{company}` `{net}` `{password_hint}` placeholders), Password Rule
  (`id` default, `id+dob`, `dob`, `custom`, `none`) and Default Country Code.
- Names and text can be in English or Arabic. Arabic is printed with the bundled Amiri font (SIL Open Font License);
  other scripts (Cyrillic, Chinese…) are flagged in the checks and print as `?`.

The downloaded template is a styled workbook: grouped colour-coded headers, notes on every header, frozen
panes, drop-down lists and data validation on IDs, dates, amounts, mobile, email and IBAN, a live Net Salary
formula and a ✓ / ⚠ check column, protected headers (no password), and optional columns grouped so they can
be collapsed. *Download my template* fills in the company details, logo and stamp.

Ready-made files are in `samples/`: the empty template, a filled sample with 25 employees, `Sample_All_Templates.pdf` (one page per design) and an Arabic sample.

## Install as an app (PWA)

Served over HTTPS (or `http://localhost`), the page is an installable app: an *Install app* button appears in
the top bar (Chrome/Edge; on iPhone use Share → Add to Home Screen). The service worker caches every file, so the
app opens and works without internet, including the template download. Installed on Windows/macOS/ChromeOS, Excel
and CSV files can be opened with it directly (*Open with › Payslips*). When a new version is deployed, a
*New version ready* message offers a reload; the open file and ticks come straight back.

Opened as a file from disk (USB stick, shared drive) everything works as before, without the install option.

When deploying a change, bump `VERSION` in `sw.js` (the test suite checks that every file the page loads is in
its cache list).

## Project layout

```
index.html        the page
css/styles.css
js/schema.js      Excel template definition, tolerant parser, column guessing and validation (pure functions)
js/template-xlsx.js  the styled, validated Excel template (built with ExcelJS, loaded only when downloading)
js/xlsx-images.js reads pictures out of an .xlsx and embeds pictures into one (JSZip)
js/pdf.js         13 PDF templates (each a small design spec), thumbnails, display options, passwords, register
js/delivery.js    .eml drafts, mail-merge CSV, WhatsApp/Gmail links, Outlook macro and guides
js/fonts/         Amiri Regular/Bold as base64 for Arabic text, plus its OFL licence
js/app.js         the five-step UI, settings file, preview, export and send, session restore, backup state
js/art.js         logo, the mascot "Qasi" and its moods, confetti, count-up numbers, card spotlight
js/guide.js       animated quick start, save point, confirm dialogs and toasts (SVG scenes animated with CSS)
js/vault.js       IndexedDB store, append-only delivery journal, first-visit record, fingerprints (also runs in Node)
js/pwa.js         service worker registration, install button, update message, file handling
sw.js             service worker (offline cache); manifest.webmanifest and icons/ for installing
css/fonts/        Sora display font (SIL Open Font License)
js/vendor/        SheetJS, jsPDF, JSZip, ExcelJS (MIT) (bundled so the page works offline)
samples/          template, sample workbook, sample PDFs
tools/            build-samples.js (regenerates samples/), build-icons.js (regenerates icons/), test.js (smoke tests)
```

## Development

```
node tools/test.js            # parser/validator/PDF smoke tests, also renders 1,000 payslips
node tools/build-samples.js   # regenerate samples/*.xlsx
WRITE_PDFS=1 node tools/test.js   # also writes samples/Sample_*.pdf
NODE_PATH=$(npm root -g) node tools/build-icons.js   # regenerate icons/ (needs Playwright + Chromium)
python3 -m http.server 8000   # try the installable app at http://localhost:8000
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
- The open file and progress are kept in the browser's storage without a password, like any local app data.
  On a shared computer, switch off *Remember my open file* in the *Private* panel, or use *Erase everything here*.
- The self-updating settings file and the delivery log written into the export folder need Chrome or Edge;
  after a browser restart they ask once for permission again (one click). Other browsers download the files.

## Look and feel

The theme is driven by four colour variables at the top of `css/styles.css` (`--c1` to `--c4`): the moving
aurora background, buttons, logo accents and the mascot follow them. Animations respect the operating system's
"reduce motion" setting. Ideas for what to build next are in `docs/IDEAS.md`.
