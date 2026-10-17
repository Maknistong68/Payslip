# Delivery plan: from "export" to "send", one click per employee

> Status: phases 1–4 are built (contact columns, password PDFs, folder/ZIP export with manifest,
> Send step with WhatsApp / Outlook drafts / Mail Merge pack / register, delivery log). Phase 5
> (viewer compatibility matrix for encrypted PDFs on phones, share sheet on the hosted URL) is still
> to be done by hand on real devices. Section 4 items remain future work.

Goal: after the PDFs are generated, HR should be able to deliver each payslip to its employee with
one action, using the tools the company already has (WhatsApp Desktop, Outlook, Gmail or any
mailbox), with no server, SMTP setup or API. Background and evidence: `reports/Offline bulk
payslip delivery options.md`.

Hard limits that shape the design (from the research):

- A web page cannot send. It can only prepare a file or open a link, and each outward action
  costs one user click. "One click per employee" is the floor; "zero clicks" is only possible
  by handing a folder plus a CSV to a mail-merge tool (Thunderbird Mail Merge add-on).
- No link format (`mailto:`, `wa.me`, Gmail compose URL) can carry an attachment.
- Classic Outlook for Windows opens a generated `.eml` with `X-Unsent: 1` as a ready draft
  (attachment included). New Outlook is unreliable; Mac, Apple Mail and web mail open it read-only.
- The OS share sheet (Web Share API with files) reaches WhatsApp and Outlook on Windows 11,
  Android and iOS from Chrome/Edge/Safari, but not from Firefox, and needs an HTTPS page
  (the Vercel URL) rather than a file opened from a USB stick.
- Salary data is personal data under the PDPL. Every file that leaves the PC should be
  password-protected and the message text should carry a one-line notice.

---

## 1. Data changes (Excel template and parser)

### New columns on the `Payslips` sheet

| Column | Required for | Accepted forms | Validation |
|---|---|---|---|
| **Mobile** | WhatsApp | `05XXXXXXXX`, `5XXXXXXXX`, `+966 5X…`, `009665…`, any `+<country>` number | Normalised to digits in E.164 without `+` (`9665XXXXXXXX`). Error if not 9–15 digits after normalisation. Saudi numbers must start with 5 after the country code. Warning on duplicates. |
| **Email** | Outlook / Gmail / mail merge | any | RFC-style check (`x@y.z`), lower-cased, warning on duplicates and on obviously personal domains only as info. |
| **Date of Birth** | Password rule | date | Parsed like other dates. Used only to build the PDF password. |

Aliases to recognise: `mobile`, `phone`, `mobile no`, `contact`, `contact no`, `whatsapp`,
`whatsapp no`, `cell`, `tel`, `هاتف`, `جوال`; `email`, `e-mail`, `mail`, `البريد`;
`dob`, `birth date`, `date of birth`, `تاريخ الميلاد`. All three go into the template,
the column-mapping screen and the settings file like any other column.

### New rows on the `Settings` sheet (all optional, also editable on the page)

| Setting | Default | Purpose |
|---|---|---|
| Sender Email | (empty) | `From:` address in generated `.eml` files |
| Message Subject | `Payslip {period} - {company}` | Subject line for email, first line for WhatsApp |
| Message Text | bilingual default (see §3.4) | Body with placeholders `{name}`, `{period}`, `{company}`, `{net}`, `{password_hint}` |
| Password Rule | `id+dob` | How the PDF password is built: `none`, `id` (National ID / Iqama), `id+dob` (ID + DDMM), `dob` (DDMMYYYY), `custom` (a `Payslip Password` column) |

### Review step additions

- A fourth tile: **Contacts**: "812 with mobile · 640 with email · 140 with neither".
- Row checks: invalid mobile, invalid email, missing contact for the channel chosen later.
- Password readiness: rows where the chosen rule cannot be built (no ID or no DOB) are flagged.

---

## 2. Foundations (Export step)

1. **File naming.** `EMP-001_2026-10_Ahmed-Hassan.pdf`: ASCII only, employee ID first so files
   sort and match the send list. Arabic names are transliterated to ASCII in the filename only.
2. **Folder export** with the File System Access API on Chrome/Edge: one folder pick, then all
   PDFs are written directly (no 60 MB ZIP in memory, files ready for drag-and-drop and for
   mail merge). ZIP stays as the fallback for Firefox/Safari and `file://`; switch it to STORE
   (PDFs do not compress) so 1,000 files pack in seconds.
3. **Password-protected PDFs.** Phase A: jsPDF's built-in encryption (40-bit RC4, already in the
   vendored library), per-employee password from the Password Rule, owner password random.
   Label it honestly in the UI: "stops casual opening; not strong encryption". Phase B: AES-256
   via `@pdfsmaller/pdf-encrypt` + `pdf-lib` (MIT, ~450 KB vendored) behind the same switch.
   Compatibility test matrix before release: Acrobat Reader, Chrome, Edge, iOS Files,
   Android WhatsApp document viewer.
4. **Manifest.** `manifest.csv` next to the PDFs: employee ID, name, file name, mobile, email,
   SHA-256 of the file (Web Crypto), password hint. This is the input for mail merge and the
   backbone of the delivery log.

---

## 3. New step 5: Send

After Export. Channel chooser at the top, then a **send list**: one row per employee with
contact, file name, status, and one action button for the chosen channel. "Sent" ticks are
stored per pay period in the browser and can be exported as a delivery log.

### 3.1 WhatsApp (WhatsApp Desktop or phone)

- Row button **Share** when `navigator.canShare({ files })` is true (HTTPS page in Chrome/Edge
  on Windows 11, Android, iOS Safari): opens the OS share sheet with the PDF and the message
  text; HR picks WhatsApp and the contact. One click plus one pick.
- Otherwise row button **Open chat**: `https://wa.me/9665XXXXXXXX?text=<message>` opens the
  employee's chat in WhatsApp Desktop with the text pre-filled; HR drags the file from the
  exported folder (which the page keeps open in a side panel listing file names in the same
  order). Two actions per employee.
- Guard rails baked into the UI: prompt to save all employees as contacts first (reduces spam
  flags), a pace counter that suggests a pause every 50 sends, and a banner explaining what
  WhatsApp's "restricted" state looks like and to stop when it appears. macOS note: use the
  "+" attach button, not drag-and-drop (WhatsApp for Mac 26.14+ deletes the original file).
- Not built: broadcast lists (256 max, same file to all), bulk-sender extensions, unofficial
  WhatsApp Web automation (ban risk).

### 3.2 Outlook (classic Outlook for Windows)

- Row button **Outlook draft**: downloads `EMP-001_2026-10.eml` with `X-Unsent: 1`, unique
  `Message-ID`, RFC 2047 subject, body, PDF as base64 attachment. Double-click opens an
  editable draft with the attachment; HR presses Send. Two clicks.
- **Batch**: "All drafts (.eml)" button writes the `.eml` files into the export folder (or a
  ZIP) in batches of 50 with an index; HR selects a batch and presses Enter to open them all.
- Advanced pack (optional download): a short VBA module for classic Outlook that reads
  `manifest.csv` and saves every message to Drafts in one run, with instructions. In-process
  VBA avoids the Object Model Guard prompt.
- Shown only when the user picks "Outlook (classic, Windows)". New Outlook, Mac and web
  Outlook users are routed to 3.3.

### 3.3 Gmail, Microsoft 365 web, Zoho, any mailbox: Mail Merge pack

The only genuinely zero-clicks-per-employee free route. The page produces a **Mail Merge
pack**: the PDF folder, `mailmerge.csv` (to, subject, body, attachment file name), and a
one-page illustrated guide:

1. Install Thunderbird and sign in with the company mailbox (Gmail and Microsoft 365 both
   work with the normal sign-in; no app passwords needed).
2. Install the free **Mail Merge** add-on.
3. Open the template message the pack provides, choose `mailmerge.csv`, pre-select the PDF
   folder when asked, set "Pause 3 seconds", press Send.

Thunderbird sends at the provider's limits (Microsoft 365: 30 per minute; free Gmail: 500 per
day; Google Workspace: 2,000 per day). The guide states these numbers so HR can plan batches.

For companies that refuse to install anything: row button **Gmail compose** opens
`mail.google.com/mail/?view=cm&to=…&su=…&body=…` with everything but the attachment; HR
attaches the PDF from the folder. Four actions per employee; shown as the fallback it is.

### 3.4 Message text

Default bilingual template, editable in Settings and saved in the settings file:

```
{company} – Payslip {period}
Dear {name}, your payslip for {period} is attached. Password: {password_hint}.
This document contains personal data; please do not forward it.
قسيمة راتب {period} – {company}. كلمة المرور: {password_hint}. هذا المستند يحتوي على بيانات شخصية، يرجى عدم مشاركته.
```

`{password_hint}` prints the rule ("your Iqama number + birth day and month, e.g. 2456789012
1503"), never the password itself.

### 3.5 Proof of delivery

- **Delivery log** (`delivery_log.csv`): employee, channel, timestamp, file hash, operator
  note. Written from the send list ticks; append-only per pay period.
- **Acknowledgement register** (PDF): one printable table per department with employee ID,
  name, period, net, signature column, for supervisors who hand out printed slips or for
  labour-inspection files. Saudi law needs a wage register, paper or electronic, not a
  signature, but a signed register is the strongest evidence in a dispute.

---

## 4. Later (not in the first build)

- Print pack: fold-to-conceal layout and a text QR on the stub for workers without data.
- Kiosk mode: encrypted dataset on a camp-office PC; worker enters Iqama + DOB, prints own slip.
- AES-256 (§2.3 Phase B) if the compatibility matrix is clean.

---

## 5. Build order and estimates

| Phase | Scope | Effort |
|---|---|---|
| 1 | Data: Mobile / Email / DOB columns, aliases, normalisation, Settings rows, Contacts tile, password rule, settings-file fields | half a day |
| 2 | Export: ASCII file names, folder export (FSA) + STORE ZIP, RC4 password PDFs, manifest.csv | half a day |
| 3 | Send step: send list, WhatsApp (share / wa.me), Outlook `.eml` per row + batch, Gmail compose fallback, sent ticks, delivery log | 1 day |
| 4 | Mail Merge pack (Thunderbird CSV + template + guide), Outlook VBA pack, acknowledgement register | half a day |
| 5 | Compatibility tests (Acrobat / iOS / Android viewers for encrypted PDFs, share sheet on Vercel URL vs file://), docs, samples | half a day |

Everything stays a static page: no new hosting, no accounts, no API keys.
