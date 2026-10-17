# Payslip Generator (offline)

A single web page that turns an Excel sheet into PDF payslips. Built for employers who
do not have a payroll system: fill a simple Excel template, open the page, upload the
file, check it, pick a design and export all payslips at once (one combined PDF or a
ZIP with one PDF per employee). It handles 1,000+ employees in a few seconds.

**No installation, no sign-in, no internet.** Everything runs inside the browser.
The Excel file never leaves the computer.

## How to use

Open `index.html` in Chrome, Edge or Firefox (double-click it; it works from a USB stick or a shared drive).
### Company profile: set up once, the same payslip every month

Payslips must not drift from month to month, so everything that makes a payslip look the way it does is
kept in one **company profile**: company details, logo and stamp, design and what to show, the message,
the password per channel and how the contractor's columns are read. Only the Excel file and the pay period
change each month.

- **First time** (no profile in this browser): Upload → Review → **Design (once)** → *Save company profile &
  continue* saves the profile as a file (Chrome/Edge: picked once, then kept up to date) and moves on to Send.
- **Every month after** (profile found in the browser, or loaded from its file on another computer):
  **Upload → Review → Send**. The Design step is gone; company details, message and passwords are shown
  read-only. Review only asks for the **pay period**, with one-click suggestions (from the sheet or file name,
  the month after the last one, or the same as last time). A new file never inherits last month's period.
- **The profile wins**: if the Excel Settings sheet names another company or has other pictures, the profile is
  used and a note says so.
- **Changing it is deliberate**: *Edit profile* (profile button at the top right, or the links on Review and
  Send) opens an edit session with a bar: *Save profile* makes a new version and saves the file, *Cancel changes*
  restores it exactly. The profile panel also saves the file, loads another profile, downloads the Excel
  template pre-filled with the company details and logo, or removes the profile from this computer.
- *Try with sample data* never becomes a profile.

The steps:

1. **Upload** – drop the Excel file (or CSV, or paste the cells straight from Excel). First time? Download the
   template (or a filled example), or load a company profile. With a profile, the template comes pre-filled.
2. **Review** – one sentence tells you whether the file is ready. Below it:
   - *Source bar* (only when needed): which sheet to read (or all sheets combined), which row holds the
     headers, whether to include hidden rows, a badge when a saved column setup was applied, and
     **Other sheets**: overtime, attendance, loans or employee-list sheets found in the same file, each with a
     switch, what it adds and how many employees matched.
   - *Totals*: earnings, deductions and net to pay, with a box for the bank / WPS transfer total that shows
     ✓ or the exact difference.
   - *Company details*: filled from the file's `Settings` sheet, or typed once here if the file has none
     (they become part of the company profile).
   - *Columns*: how each column of your sheet was understood. Anything guessed or ignored can be remapped
     in one click, and the choice is remembered for the next upload.
   - *Employees*: every row with its calculated earnings, deductions, net and the checks that failed.
3. **Design** (first time, or while editing the profile) – one of 13 templates (Classic, Modern, Minimal, Executive, Steel, Corporate, Bold, Ledger,
   Compact, Split, Formal, Heritage, Receipt), accent colour, paper size, title, logo and stamp,
   and a checklist of what to print, with a live preview.
4. **Send** – view first, then send. Nothing is sent by the page itself:
   - **View** opens the payslip in a viewer with *previous / next*, the net pay, the row's warnings and the
     send buttons. *Start* (or *Continue with …*) opens the next person not sent yet.
   - **Send on WhatsApp** opens the chat with the right number and the message typed, and downloads the PDF to
     attach (drag it into the chat). On phones and Chrome/Edge with file sharing, *share the file to WhatsApp*
     hands the PDF over directly.
   - **Send by email** downloads a ready Outlook email (`.eml`) with the PDF attached: open it, press Send.
   - **Download** saves the single PDF.
   - **Password on the PDF**, one switch per channel: WhatsApp on (recommended), email off, downloads on; the
     password is the Iqama / National ID number by default (or Iqama + birth day and month, date of birth, or a
     password column). The message includes the password hint only when that channel's PDF is locked.
   - **More**: all PDFs to a folder or a ZIP, one PDF to print, the signature register, all Outlook emails as a
     ZIP (batches of 50 + macro), the Thunderbird Mail Merge pack for Gmail or any mailbox, the delivery log
     CSV, restore ticks from a log, clear ticks.
   Rows are ticked as you send; the ticks survive reloads and export as a delivery log CSV with file hashes.

### First visit and never losing work

- **First visit** (nothing saved in this browser): a short animated guide opens by itself (template, fill,
  drop, view, WhatsApp, email). It can be replayed from *▶ How it works*.
- The top bar shows the company profile and whether it is saved as a file (green) or not yet (amber).
- **Refresh, crash or browser restart**: the open file, the step, the sheet and header choices, the export
  options, the bank total and this month's pay period are kept in the browser (IndexedDB). Within three days the
  page reopens exactly where it was; older work is dropped (ticks are kept separately, per company and period).
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

The template gives the best result, but any salary sheet is accepted (20 real-world cases are tested; see
`docs/PARSING_SCENARIOS.md`):

- One row per employee. Required: a name column and at least one amount (a Net Salary column alone is enough).
  Everything else is optional. First Name / Last Name columns are joined.
- Several sheets: sheets with the same layout (one per site) are combined; one sheet per month picks the
  pay period's month (or the latest). Hidden sheets, hidden rows and sheets named Lists/Lookup are skipped.
- Extra sheets are used: overtime (amounts and hours, also daily logs), attendance (also day-by-day P/A
  timesheets), loans / deductions and employee lists (names, Iqama, mobile, email, bank), matched by employee
  number, Iqama or name. Details fill empty fields; amounts are added automatically only when the payroll's
  net salary column confirms them, otherwise one switch on the Review screen decides. Summaries and pivots are
  ignored.
- Sideways tables (one column per employee) and one-sheet-per-employee payslip forms are read too. Headers
  repeated on printed pages and page totals are skipped.
- Files that are not a payroll (stock lists, attendance only, an employee list without salaries) get a clear
  message instead of empty payslips.
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
js/app.js         the step-by-step UI, company profile, preview, viewer and sending, session restore
js/art.js         logo, the mascot "Qasi" and its moods, confetti, count-up numbers, card spotlight
js/guide.js       animated guide, dialogs and toasts (SVG scenes animated with CSS)
js/vault.js       IndexedDB store, append-only delivery journal, first-visit record, fingerprints (also runs in Node)
js/pwa.js         service worker registration, install button, update message, file handling
sw.js             service worker (offline cache); manifest.webmanifest and icons/ for installing
css/fonts/        Sora display font (SIL Open Font License)
js/vendor/        SheetJS, jsPDF, JSZip, ExcelJS (MIT) (bundled so the page works offline)
samples/          template, sample workbook, sample PDFs
tools/            build-samples.js (regenerates samples/), build-icons.js (regenerates icons/), test.js (smoke tests),
                  scenarios.js (the 20 workbook cases)
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
- WhatsApp cannot receive a file from a web page directly on most computers: the chat opens with the message,
  and the PDF (already downloaded) is dragged in by hand. On phones the share sheet sends it in one step.
- The viewer shows the PDF with the browser's built-in PDF viewer; some phone browsers show an
  "Open in a new tab" link instead.
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
- The self-updating company profile file and the delivery log written into the export folder need Chrome or Edge;
  after a browser restart they ask once for permission again (one click). Other browsers download the files.

## Look and feel

The theme is driven by four colour variables at the top of `css/styles.css` (`--c1` to `--c4`): the moving
aurora background, buttons, logo accents and the mascot follow them. Animations respect the operating system's
"reduce motion" setting. Ideas for what to build next are in `docs/IDEAS.md`.
