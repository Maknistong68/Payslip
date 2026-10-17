# Excel input assessment: template, worst cases, validation

Tested on 17 Oct 2026 against the current parser (`js/schema.js`) with 16 realistic contractor-sheet cases.

## 1. The template we hand out today

**What works**
- Clear split: `Settings` (company, once), `Payslips` (one row per employee), `Instructions`.
- Forgiving reader: aliases, keyword guessing, example row skipped, extra `Earning:` / `Deduction:` columns.
- Pictures (logo, stamp) are read from the Settings sheet.

**What is weak** (confirmed by opening the file with openpyxl)

| Item | Today |
|---|---|
| Data validation | none (0 rules on every sheet) |
| Header styling, colour groups | none |
| Frozen header row / columns | none |
| Number formats | `General` everywhere |
| Text format for Iqama, mobile, IBAN | none, so Excel turns `0551234567` into `551234567` |
| Net Salary column | empty, so the contractor cannot see the net before uploading |
| Totals row to reconcile with the bank transfer | none |
| Sheet protection of headers | none, headers can be renamed or deleted |
| Column count | 27 columns, all visible, required and optional mixed |
| Cause | the free SheetJS writer cannot write styles, validation or formulas |

## 2. Worst cases: contractor keeps their own sheet

Result legend: **OK** handled; **Blocked** stopped with a clear message (safe, but needs a click or a fix);
**Silent** wrong or missing data with no clear warning (dangerous).

| # | Contractor's sheet | What happens today | Result |
|---|---|---|---|
| 1 | Two sites as two sheets (Site A, Site B) | Only the first sheet is read; Site B employees are dropped with only an info note | **Silent** |
| 2 | One sheet per month (Sep, Oct) | Reads the first sheet (September) | **Silent** |
| 3 | Decimal comma `3.000,50` | Read as 3.0005 → net 3.00 instead of 3,000.50 | **Silent** |
| 4 | US date `03/15/2024` | Read as day 03, month 15 → shows "03 Mar 2025" | **Silent** |
| 5 | Error cell `#N/A` in an allowance | Treated as empty; net is short with no warning | **Silent** |
| 6 | Hidden rows for people who left | Hidden rows still get payslips | **Silent** |
| 7 | Deductions typed as negatives (-100) | Every row flagged as an error and skipped | Blocked |
| 8 | Arabic-Indic digits `٣٠٠٠` | Row flagged "not a number" | Blocked |
| 9 | Only Name + Net Salary columns | "No salary column" error; works only if the user maps Net as an earning | Blocked |
| 10 | Arabic headers only | No header recognised; works after manual mapping (remembered next month) | Blocked |
| 11 | Department heading rows ("PROJECTS") inside the data | Become error rows (skipped by default); department totals are skipped correctly | Blocked |
| 12 | First Name / Last Name columns | Last name ignored, with a warning | Blocked |
| 13 | Header below 40 lines of notes | Header not found | Blocked |
| 14 | Two-row header ("Earnings" over "Basic, Food") | Group row used to classify columns; net correct | OK |
| 15 | "Salary", "Allowance 1/2", "Deduction" generic names | Classified correctly; net correct | OK |
| 16 | Net in file differs by rounding | Warning, calculated net printed | OK |

Also already handled: title rows above the header, totals rows, amounts as text with separators or "SR",
no ID column (IDs generated), no Settings sheet (details typed once on the page), column order, extra columns.

## 3. How to manage contractors who will not change their sheet

Principle: **adopt their file, do not convert it.** The template is the fast lane, not a requirement.
The legal floor is a wage record, not our layout, so the minimum input should be *a name and one amount*.

1. **Fix the six silent failures first** (they produce wrong payslips without anyone noticing):
   - multi-sheet files: show a sheet picker with employee counts, plus "combine all sheets";
   - month sheets: pre-select the sheet whose name matches the Pay Period, else the last sheet, and say which one was used;
   - decimal comma: detect per column (`\d{1,3}(\.\d{3})+,\d{2}` pattern) and convert;
   - ambiguous dates: if the middle part is above 12 treat as US format; if both parts are ≤ 12 and the year is not ISO, warn;
   - error cells (`#N/A`, `#REF!`, `#DIV/0!`): mark the row as an error naming the cell;
   - hidden rows: skip them and list them ("3 hidden rows skipped: …"), with a toggle to include.
2. **Turn the blocked cases into one-click fixes**:
   - all values in a deduction column ≤ 0 → treat as positive deductions automatically (with a note);
   - Arabic-Indic and Persian digits → convert to 0–9;
   - only a net column → "net-only payslip" (one line "Salary", net = that amount), with a warning;
   - Arabic header aliases for the common columns (الاسم، الراتب الأساسي، بدل السكن، بدل النقل، الصافي…);
   - a row with a label but no amounts in any money column → section heading, skipped silently;
   - First/Last name → join into Employee Name;
   - header search beyond row 40, plus a "this row is the header" picker on the review screen.
3. **Map once, remember forever.** Column choices are already saved per header text and travel in the settings file.
   Add a visible "Profile: ABC Trading (saved 3 Oct)" badge so HR trusts it is applied.
4. **Pre-built profiles for large contractors.** The compliance team maps a big contractor's sheet once and sends them
   the settings file; next month the contractor just drops their own file.
5. **Reconciliation guard.** Show "Total net in this file: SAR X" and ask the user to compare it with the bank transfer
   total (WPS file). This single check catches most silent problems, whatever their cause.
6. **Last resort.** If a file is protected, linked or formula-only: "File → Save As → Excel Workbook, values" instruction, and a
   paste box where the user copies the block of cells straight from their sheet.

## 4. Template layout and data validation (proposed)

**Layout**
- `Payslips` as an Excel Table (validation and formats extend to new rows automatically), header row and the first two
  columns frozen, filter buttons on.
- Colour-banded header groups: Employee (navy) · Contact (slate) · Attendance (grey) · Earnings (green) · Deductions (red) ·
  Result (blue).
- Required columns first (ID, Name, Iqama, Basic); rarely used ones (DOB, Nationality, Bank, IBAN, Remarks) in a collapsible
  column group.
- `Net Salary` as a locked formula column (`=SUM(earnings)-SUM(deductions)`) and a `Check` column that shows a warning
  symbol when deductions exceed earnings or Basic is empty, so problems are visible before upload.
- Table total row: total gross, deductions and net, to reconcile with the bank transfer.
- Formats: amounts `#,##0.00`; Iqama, mobile, IBAN, Employee ID as Text (keeps leading zeros); dates `dd-mmm-yyyy`.
- Header notes (cell comments) and validation input messages explain each column on hover; the Instructions sheet shrinks
  to five lines and a filled example.
- Sheet protection without a password: headers and formula columns locked, data cells open.
- `Settings` as a two-column form with drop-downs and a framed "insert logo here" area.

**Validation rules**

| Column | Rule | On violation |
|---|---|---|
| Employee ID | unique in column (`COUNTIF(A:A,A2)=1`) | Stop |
| Employee Name | text length ≥ 2 | Stop |
| National ID / Iqama | 10 digits, starts with 1 (Saudi) or 2 (resident) | Warning |
| Nationality | drop-down list of common nationalities | Warning (other values allowed) |
| Date of Birth | date between 1940 and 16 years ago | Stop |
| Joining Date | date between 1980 and end of pay period | Stop |
| Mobile | 10 digits starting `05`, or starts with `+` | Warning |
| Email | contains `@` and a dot after it | Warning |
| IBAN | 24 characters starting `SA` | Warning |
| Working Days | whole number 0–31 | Stop |
| Absent Days | whole number 0–31 and ≤ Working Days | Stop |
| Overtime Hours | number 0–300 | Stop |
| Every amount column | decimal 0–1,000,000 | Stop |
| Settings: Currency | list SAR, AED, QAR, KWD, BHD, OMR, USD | Stop |
| Settings: Pay Period | list of the last and next 12 months | Warning |
| Settings: Password Rule | list id, id+dob, dob, custom, none | Stop |
| Settings: Payment Date | date ≥ Period End Date | Warning |

Caveat: Excel only enforces validation while typing; pasted values bypass it. The app's own checks remain the authority;
the template rules just catch mistakes earlier.

**How to build it:** the free SheetJS writer cannot write any of this. Options:
- **ExcelJS** (MIT, about 1 MB, offline, vendored and loaded only when the template is downloaded): writes styles, tables,
  validation, formulas, notes and protection. Recommended: keeps the "My template" pre-fill and logo working.
- A master template built once with Python/openpyxl and shipped as a static file, patched in the browser for "My template".
  Zero new runtime code, but fiddly to patch.

## 5. Suggested order

1. Six silent-failure fixes + the reconciliation total (highest risk, small code).
2. Blocked-case auto-fixes (negatives, Arabic digits and headers, net-only, section rows, first/last name, header picker).
3. Redesigned template with validation (ExcelJS).
4. Profile badge and pre-built profiles for large contractors.
