# Parsing re-audit: 20 workbooks

Each case is a real kind of file a contractor may upload. They are built and checked automatically by `tools/scenarios.js` (run `node tools/scenarios.js`; `node tools/test.js` fails if any case regresses).

Result before this round: 2 of 20 correct. Now: 20 of 20.

| # | Workbook | Before | Now | Expected |
|---|---|---|---|---|
| 1 | Our template, filled (25 employees) | OK | ✓ 25 payslips from Payslips | 25 payslips, no errors |
| 2 | Payroll + separate "Overtime" sheet; net already includes overtime | Wrong: overtime sheet read as more employees (5 payslips, errors) | ✓ 3 payslips from Payroll + Overtime | overtime added to the payslips automatically (the net column proves it), hours too |
| 3 | Payroll has an Overtime column + a daily "OT Log" sheet (many rows per person) | Wrong: OT log read as employees; hours lost | ✓ 2 payslips from Salaries + OT Log | hours summed per person; overtime pay not counted twice |
| 4 | Payroll + "Attendance" timesheet grid (days 1-30 with P / A / W) | Wrong: timesheet read as employees; days lost | ✓ 2 payslips from Payroll + Attendance | working days and absent days counted from the grid |
| 5 | Payroll with IDs only + "Employees" master list (names, Iqama, mobile, email) | Wrong: master list read as employees; names missing | ✓ 3 payslips from Pay + Employees | names and contacts filled from the master list by employee number |
| 6 | No IDs anywhere: "Contacts" sheet matched by name (different case and spacing) | Wrong: contacts read as employees | ✓ 3 payslips from Sheet1 + Contacts | matched by name; the person missing from payroll is reported |
| 7 | "Summary" sheet (department totals) comes before the payroll sheet | Wrong: summary combined with payroll | ✓ 2 payslips from Payroll | payroll sheet used alone, summary ignored |
| 8 | Two site sheets + one "Overtime" sheet for both + a "Summary" | Wrong: overtime and summary read as employees | ✓ 4 payslips from All sheets + Overtime | sites combined, overtime added, summary ignored |
| 9 | Monthly sheets (Sep, Oct) with monthly overtime sheets (OT Sep, OT Oct) | Partly: October used, overtime ignored | ✓ 1 payslips from Oct 2026 + OT Oct | October payroll with October overtime only |
| 10 | Not payroll at all: an inventory list | Blocked with generic errors | ✓ 0 payslips from Stock · not-payroll | refused with a clear "not a payroll sheet" message |
| 11 | Attendance only (timesheet, no amounts) | Wrong: 2 empty payslips | ✓ 0 payslips from Timesheet · attendance-only | explains that payslips need amounts, not only attendance |
| 12 | Employee list only (no salaries) | Wrong: 2 empty payslips | ✓ 0 payslips from Staff · no-salary | explains that salary columns are missing |
| 13 | One sheet per employee, laid out like a payslip form | Wrong: forms not understood | ✓ 3 payslips from All sheets | each sheet becomes one payslip |
| 14 | Sideways table: employees in columns, pay items in rows | Blocked: no header found | ✓ 3 payslips from Oct | turned the right way round automatically |
| 15 | Abbreviated headers (Emp#, B.Sal, H.A, T.A, OT Amt, Ded, Net Pay) | Partly: 4 of 8 columns understood | ✓ 2 payslips from Data | all columns understood |
| 16 | Printed report: header repeated on every page, page totals in between | Wrong: repeated headers became error rows | ✓ 5 payslips from Report | repeated headers and page totals skipped, 5 payslips |
| 17 | Broken formulas (#DIV/0!, #REF!) in a few cells | OK | ✓ 3 payslips from Payroll | the two rows are flagged with the cell, the rest are fine |
| 18 | Arabic sheet names: "الرواتب" + "السلف" (loans) sheet | Wrong: loans sheet read as employees | ✓ 2 payslips from الرواتب + السلف | loan deduction added (the net proves it) |
| 19 | CSV from another system: semicolons and decimal commas | Wrong: separator not detected when uploading a CSV | ✓ 2 payslips from Sheet1 | amounts read correctly |
| 20 | Busy workbook: instructions, hidden calc sheet, empty chart sheet, pivot, lists, and the payroll on sheet 6 | Wrong: pivot combined with payroll | ✓ 2 payslips from Data | only the payroll sheet is used |

## How the app now reads a workbook

1. **Every sheet gets a role** from its columns and name: payroll, overtime, attendance (including day-by-day timesheets), deductions / loans, employee details, summary, or not used. Sheet names in Arabic work too.
2. **Payroll sheets** are combined (one per site) unless they list the same people; monthly sheets pick the pay period or the latest month.
3. **The other sheets add to the payroll**, matched by employee number, Iqama or name (case, spacing and word order do not matter):
   - details (mobile, email, Iqama, job, bank, working days, overtime hours) only fill empty fields and are on by default;
   - amounts (overtime pay, loans, deductions) are switched on automatically only when the payroll's own net salary column proves they belong there; otherwise the app warns and the user decides with one switch;
   - amounts already in the payroll are never added twice;
   - people in an overtime or attendance sheet who are missing from the payroll are reported ("they get no payslip").
4. **Unusual layouts**: sideways tables (one column per employee) and one-sheet-per-employee payslip forms are turned into rows; headers repeated on every printed page and page totals are skipped.
5. **Files that are not payroll** (inventory, attendance only, employee list only) are refused with a message that says what was found, instead of producing empty payslips. Mapping the columns by hand still works.
6. **CSV**: the separator (comma, semicolon, tab, pipe) is detected and values are read as typed, so decimal commas and IDs with leading zeros survive; Arabic CSV files saved by older Excel versions (Windows-1256) are decoded.

On the Review screen, the other sheets appear under the sheet picker as switches, each with what it adds and how many employees matched.
