#!/usr/bin/env node
/* Smoke tests for the parser, validator and PDF renderer. Run: node tools/test.js */
'use strict';
var path = require('path');
var fs = require('fs');
var assert = require('assert');
var XLSX = require(path.join(__dirname, '..', 'js', 'vendor', 'xlsx.full.min.js'));
var jspdf = require(path.join(__dirname, '..', 'js', 'vendor', 'jspdf.umd.min.js'));
var Schema = require(path.join(__dirname, '..', 'js', 'schema.js'));
var PDF = require(path.join(__dirname, '..', 'js', 'pdf.js'));
var jsPDF = jspdf.jsPDF;
require(path.join(__dirname, '..', 'js', 'fonts', 'amiri.js'));

function readWb(buf) { return XLSX.read(buf, { type: 'buffer', cellDates: true, cellStyles: true }); }
function wbToBuf(wb) { return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx', cellDates: true }); }

// 1. Template round-trip: template alone has no rows (example row ignored) and reports missing settings.
var tpl = Schema.parseWorkbook(XLSX, readWb(wbToBuf(Schema.buildTemplateWorkbook(XLSX))));
assert.strictEqual(tpl.slips.length, 0, 'example row must be ignored');
assert.strictEqual(tpl.settings.companyName, '', 'empty template has no company name');
assert.ok(tpl.settingsFound, 'settings sheet recognised');
assert.ok(tpl.issues.some(function (i) { return i.code === 'no-rows'; }));

// 2. Sample data parses cleanly.
var sample = Schema.parseWorkbook(XLSX, readWb(wbToBuf(Schema.buildTemplateWorkbook(XLSX, { sampleRows: Schema.sampleRows(25) }))));
assert.strictEqual(sample.slips.length, 25);
assert.strictEqual(sample.stats.fileErrors, 0, JSON.stringify(sample.issues));
assert.strictEqual(sample.stats.withErrors, 0, JSON.stringify(sample.slips.filter(function (s) { return s.hasErrors; }).map(function (s) { return s.issues; })));
assert.strictEqual(sample.settings.companyName, 'Example Contracting Co. Ltd.');
assert.strictEqual(sample.settings.currency, 'SAR');
assert.strictEqual(sample.settings.paymentDate, '05 Nov 2026');
var s0 = sample.slips[0];
assert.strictEqual(s0.employeeId, 'EMP-001');
assert.strictEqual(s0.gross, Schema.round2(s0.earnings.reduce(function (a, e) { return a + e.amount; }, 0)));
assert.strictEqual(s0.net, Schema.round2(s0.gross - s0.totalDeductions));
assert.strictEqual(s0.fields.joiningDate, '01 Jan 2022');

// 3. Validation catches bad data.
var bad = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(bad, XLSX.utils.aoa_to_sheet([
  ['Setting', 'Value'], ['Company Name', 'X Co'], ['CR Number', ''], ['Pay Period', 'May 2026'], ['Currency', 'sar']
]), 'Settings');
XLSX.utils.book_append_sheet(bad, XLSX.utils.aoa_to_sheet([
  ['Employee ID', 'Employee Name', 'Basic Salary', 'Housing Allowance', 'GOSI Deduction', 'Earning: Site Allowance', 'Deduction: Housing Loan', 'Net Salary', 'IBAN', 'Mystery'],
  ['E1', 'A', '4,000', 500, 100, 250, 50, 4600, 'SA12', ''],
  ['E1', 'B', 'abc', 0, 0, 0, 0, '', '', ''],
  ['E3', '', 1000, 0, 5000, 0, 0, '', '', ''],
  ['E4', 'محمد', 1000, 0, 0, 0, 0, '', 'SA0380000000608010167519', ''],
  [null, null, null, null, null, null, null, null, null, null],
  ['E5', 'D', -5, 0, 0, 0, 0, '', '', '']
]), 'Payslips');
var b = Schema.parseWorkbook(XLSX, readWb(wbToBuf(bad)));
assert.strictEqual(b.settings.crNumber, '', 'empty CR number stays empty (UI asks for it)');
assert.ok(b.issues.some(function (i) { return /Mystery/.test(i.message) && i.level === 'warning'; }));
assert.strictEqual(b.settings.currency, 'SAR');
assert.strictEqual(b.slips.length, 5);
var r1 = b.slips[0];
assert.strictEqual(r1.gross, 4750, 'thousand separators + custom earning');
assert.strictEqual(r1.totalDeductions, 150);
assert.strictEqual(r1.net, 4600);
assert.ok(r1.earnings.some(function (e) { return e.label === 'Site Allowance'; }));
assert.ok(r1.deductions.some(function (e) { return e.label === 'Housing Loan'; }));
assert.ok(r1.issues.some(function (i) { return /duplicated/.test(i.message); }));
assert.ok(r1.issues.some(function (i) { return /IBAN/.test(i.message); }));
assert.ok(!r1.issues.some(function (i) { return /does not match/.test(i.message); }), 'net matches');
assert.ok(b.slips[1].issues.some(function (i) { return /not a number/.test(i.message); }));
assert.ok(b.slips[2].issues.some(function (i) { return /name is empty/.test(i.message); }));
assert.ok(b.slips[2].issues.some(function (i) { return /exceed/.test(i.message); }));
assert.ok(!b.slips[3].issues.some(function (i) { return /cannot print/.test(i.message); }), 'Arabic names are supported now');
assert.ok(Schema.hasNonLatin('Привет'), 'Cyrillic still flagged');
assert.ok(!b.slips[3].issues.some(function (i) { return /IBAN/.test(i.message); }), 'valid SA IBAN');
assert.ok(b.slips[4].issues.some(function (i) { return /negative/.test(i.message); }));
assert.strictEqual(b.stats.withErrors, 4);

// 3b. Worst case: a company's own sheet, no Settings sheet, title rows, merged group row, totals row, no ID column
var own = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(own, XLSX.utils.aoa_to_sheet([
  ['ABC Trading Est. - Salary Sheet September 2026'],
  [],
  [null, null, null, null, 'Earnings', null, null, null, 'Deductions', null, null],
  ['S.No', 'Name of Employee', 'Iqama No.', 'Designation', 'Basic (SAR)', 'HRA', 'Site All.', 'Mobile', 'Absent Amt', 'Traffic Fines', 'Total Salary', 'Net Payable', 'Bank A/C'],
  [1, 'Ali Hassan', '2345678901', 'Driver', 3000, 750, 200, 100, 0, 150, 4050, 3900, 'SA1234'],
  [2, 'Juan Dela Cruz', '2345678902', 'Welder', '4,500.00', 1125, 0, 0, 300, 0, 5625, 5325, ''],
  [3, 'Mohamed Saeed', '2345678903', 'Helper', 2000, 500, 0, 0, 0, 0, 2500, 2500, ''],
  ['Total', null, null, null, 9500, 2375, 200, 100, 300, 150, 12175, 11725, ''],
  [null, null, null, null, null, null, null, null, null, null, null, null, null],
  ['Prepared by: Accounts']
]), 'Sheet1');
var o = Schema.parseWorkbook(XLSX, readWb(wbToBuf(own)));
assert.ok(!o.settingsFound);
assert.ok(o.issues.some(function (i) { return i.code === 'no-settings-sheet'; }));
assert.strictEqual(o.headerRow, 4, 'header row detected at row 4');
assert.strictEqual(o.slips.length, 3, 'totals/prepared-by rows skipped: ' + JSON.stringify(o.slips.map(function (s) { return s.name; })));
var s1 = o.slips[0];
assert.strictEqual(s1.employeeId, '1');
assert.strictEqual(s1.name, 'Ali Hassan');
assert.strictEqual(s1.fields.nationalId, '2345678901');
assert.strictEqual(s1.fields.jobTitle, 'Driver');
assert.strictEqual(s1.gross, 4050, 'basic+hra+site+mobile: ' + JSON.stringify(s1.earnings));
assert.strictEqual(s1.totalDeductions, 150, JSON.stringify(s1.deductions));
assert.strictEqual(s1.net, 3900);
assert.ok(s1.earnings.some(function (e) { return e.label === 'Site All.'; }), 'group-row guess: Site All. -> earning');
assert.ok(s1.deductions.some(function (e) { return /Traffic Fines/.test(e.label); }), 'Traffic Fines -> deduction');
assert.ok(!s1.issues.some(function (i) { return /does not match/.test(i.message); }), 'gross and net cross-checks agree: ' + JSON.stringify(s1.issues));
assert.strictEqual(o.slips[1].gross, 5625, 'text amount with thousands separator');
assert.strictEqual(o.slips[1].totalDeductions, 300, 'Absent Amt -> absence deduction');
var ibanCol = o.columns.filter(function (c) { return c.header === 'Bank A/C'; })[0];
assert.ok(ibanCol && ibanCol.kind === 'column' && ibanCol.def.key === 'iban', 'Bank A/C -> IBAN');

// Manual mapping override: treat "Mobile" as ignored and "S.No" as not an id
var o2 = Schema.parseWorkbook(XLSX, readWb(wbToBuf(own)), { overrides: { 'mobile': { kind: 'ignored' }, [Schema.normalizeHeader('S.No')]: { kind: 'ignored' } } });
assert.strictEqual(o2.slips[0].gross, 3950);
assert.strictEqual(o2.slips[0].employeeId, 'EMP-001', 'generated id when no id column');
assert.ok(o2.issues.some(function (i) { return i.code === 'missing-id'; }));
assert.ok(o2.slips[0].issues.some(function (i) { return /Gross salary in the file/.test(i.message); }), 'gross mismatch now reported');

// Arabic headers are recognised on their own now
var ar = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(ar, XLSX.utils.aoa_to_sheet([
  ['الرقم', 'الاسم', 'الراتب الأساسي', 'بدل السكن', 'الصافي'],
  ['1', 'Ahmed', 3000, 500, 3500]
]), 'ورقة1');
var a = Schema.parseWorkbook(XLSX, readWb(wbToBuf(ar)));
assert.strictEqual(a.slips.length, 1, JSON.stringify(a.issues));
assert.strictEqual(a.slips[0].net, 3500);
assert.strictEqual(a.stats.fileErrors, 0, JSON.stringify(a.issues));

// Truly unknown headers -> error, then manual mapping rescues the file
var unk = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(unk, XLSX.utils.aoa_to_sheet([
  ['Col A', 'Col B', 'Col C', 'Col D'],
  ['1', 'Ahmed', 3000, 500]
]), 'Data');
var u = Schema.parseWorkbook(XLSX, readWb(wbToBuf(unk)));
assert.ok(u.issues.some(function (i) { return i.code === 'not-payroll'; }), JSON.stringify(u.issues));
var u2 = Schema.parseWorkbook(XLSX, readWb(wbToBuf(unk)), { overrides: {
  'col b': { kind: 'column', key: 'name' }, 'col c': { kind: 'column', key: 'basicSalary' }, 'col d': { kind: 'column', key: 'housing' }
} });
assert.strictEqual(u2.slips.length, 1);
assert.strictEqual(u2.slips[0].net, 3500);
assert.strictEqual(u2.stats.fileErrors, 0, JSON.stringify(u2.issues));

// CSV-ish single sheet with template headers but no Settings
var csv = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(csv, XLSX.utils.aoa_to_sheet([
  ['Employee ID', 'Employee Name', 'Basic Salary', 'Housing Allowance', 'GOSI Deduction'],
  ['A1', 'X', 1000, 200, 50]
]), 'data');
var c = Schema.parseWorkbook(XLSX, readWb(wbToBuf(csv)));
assert.strictEqual(c.slips.length, 1);
assert.strictEqual(c.slips[0].net, 1150);
assert.strictEqual(c.stats.fileErrors, 0);

// 3c. Contact columns, passwords, file names, eml
var Delivery = require(path.join(__dirname, '..', 'js', 'delivery.js'));
var s0c = sample.slips[0];
assert.strictEqual(sample.stats.withMobile, 25);
assert.strictEqual(sample.stats.withEmail, 25);
assert.ok(/^9665\d{8}$/.test(s0c.fields.mobile), 'mobile normalised: ' + s0c.fields.mobile);
assert.ok(/@example\.com$/.test(s0c.fields.email));
assert.ok(/^\d{4}-\d{2}-\d{2}$/.test(s0c.fields.dobDate), 'dob iso: ' + s0c.fields.dobDate);
assert.strictEqual(Schema.normalizePhone('0551234567').value, '966551234567');
assert.strictEqual(Schema.normalizePhone('+971 50 123 4567').value, '971501234567');
assert.ok(Schema.normalizePhone('12345').error);
assert.strictEqual(Schema.normalizeEmail(' A.B@Example.COM ').value, 'a.b@example.com');
assert.ok(Schema.normalizeEmail('nope').error);
assert.strictEqual(Schema.periodSlug('October 2026'), '2026-10');
assert.strictEqual(Schema.periodSlug('10/2026'), '2026-10');
assert.strictEqual(Schema.slipFileName({ employeeId: 'EMP-001', name: 'Ahmed Hassan', rowNumber: 2 }, { payPeriod: 'October 2026' }), 'EMP-001_2026-10_Ahmed-Hassan.pdf');
assert.strictEqual(Schema.slipFileName({ employeeId: '7', name: 'محمد', rowNumber: 2 }, { payPeriod: 'Oct 2026' }), '7_2026-10.pdf');
var pwSlip = { employeeId: 'E1', fields: { nationalId: '2456789012', dobDate: '1990-03-15' } };
assert.strictEqual(Schema.buildPassword(pwSlip, 'id+dob').password, '24567890121503');
assert.strictEqual(Schema.buildPassword(pwSlip, 'id').password, '2456789012');
assert.strictEqual(Schema.buildPassword(pwSlip, 'dob').password, '15031990');
assert.strictEqual(Schema.buildPassword(pwSlip, 'none').password, '');
assert.ok(Schema.buildPassword({ employeeId: 'E1', fields: {} }, 'id+dob').error);
var encDoc = PDF.createDoc(jsPDF, 'a4', { password: '24567890121503' });
PDF.renderPayslip(encDoc, { slip: sample.slips[0], settings: sample.settings, templateId: 'classic', accent: '#1f4e79', options: {} });
var encBuf = Buffer.from(encDoc.output('arraybuffer')).toString('latin1');
assert.ok(encBuf.indexOf('/Encrypt') >= 0, 'PDF has an /Encrypt dictionary');
assert.ok(encBuf.indexOf('Example Contracting') < 0, 'plain text not visible in encrypted PDF');
var plainDoc = PDF.createDoc(jsPDF, 'a4');
PDF.renderPayslip(plainDoc, { slip: sample.slips[0], settings: sample.settings, templateId: 'classic', accent: '#1f4e79', options: {} });
assert.ok(Buffer.from(plainDoc.output('arraybuffer')).toString('latin1').indexOf('/Encrypt') < 0, 'no /Encrypt without password');
var emlText = Delivery.buildEml({ from: 'payroll@example.com', to: 'a@b.com', subject: 'Payslip October 2026 - قسيمة راتب', text: 'Dear Ahmed,\nline 2', fileName: 'EMP-001_2026-10.pdf', pdfBytes: new Uint8Array(Buffer.from(plainDoc.output('arraybuffer'))) });
var emlLines = emlText.split('\r\n');
assert.strictEqual(emlLines[0], 'X-Unsent: 1');
assert.ok(/^Message-ID: <.+@payslip\.local>$/.test(emlLines[1]));
assert.ok(emlText.indexOf('Subject: =?UTF-8?B?') >= 0, 'RFC 2047 subject');
assert.ok(emlText.indexOf('Content-Disposition: attachment; filename="EMP-001_2026-10.pdf"') >= 0);
assert.ok(!/[^\r]\n/.test(emlText), 'CRLF line endings only');
assert.ok(emlLines.every(function (l) { return l.length <= 998; }), 'lines <= 998 chars');
assert.ok(emlLines.filter(function (l) { return /^[A-Za-z0-9+\/=]{76}$/.test(l); }).length > 10, 'base64 folded at 76');
assert.strictEqual(Delivery.encodeHeader('Plain ASCII'), 'Plain ASCII');
assert.ok(Delivery.encodeHeader('قسيمة راتب لشهر أكتوبر ٢٠٢٦ للموظف محمد عبدالله الحربي').split('\r\n ').every(function (w) { return w.length <= 75; }), 'encoded words <= 75');
assert.strictEqual(Delivery.waLink('966551234567', 'Hi there'), 'https://wa.me/966551234567?text=Hi%20there');
var mm = Delivery.mailMergeCsv([{ slip: sample.slips[0], fileName: 'a.pdf', subject: 'S', body: 'B, "q"' }], sample.settings);
assert.ok(mm.indexOf('to,name,employee_id,subject,body,attachment') >= 0);
assert.ok(mm.indexOf('"B, ""q"""') >= 0, 'csv quoting');
var vba = Delivery.outlookVbaModule(sample.settings);
assert.ok(vba.indexOf('Sub CreatePayslipDrafts()') >= 0 && vba.indexOf('mail.Attachments.Add') >= 0);
var regDoc = PDF.createDoc(jsPDF, 'a4');
PDF.renderRegister(regDoc, { settings: sample.settings, slips: sample.slips });
assert.ok(regDoc.internal.getNumberOfPages() >= 1);


// 3d. Worst cases from docs/EXCEL_ASSESSMENT.md
function wbOf(sheets, mut) {
  var w = XLSX.utils.book_new();
  sheets.forEach(function (sh) { var ws = XLSX.utils.aoa_to_sheet(sh[1]); if (mut) mut(ws, sh[0]); XLSX.utils.book_append_sheet(w, ws, sh[0]); });
  return readWb(wbToBuf(w));
}
function codes(p) { return p.issues.map(function (i) { return i.code; }); }
var WH = ['Name', 'Basic', 'Housing', 'GOSI', 'Net'];
// two sites -> combined, department = sheet name
var w1 = Schema.parseWorkbook(XLSX, wbOf([['Site A', [WH, ['Ali', 3000, 750, 100, 3650], ['Omar', 2000, 500, 0, 2500]]], ['Site B', [WH, ['Raju', 1800, 0, 0, 1800], ['Jose', 2200, 0, 0, 2200]]]]));
assert.strictEqual(w1.slips.length, 4, 'both sheets combined');
assert.strictEqual(w1.sheetMode, 'combined');
assert.ok(codes(w1).indexOf('combined') >= 0);
assert.strictEqual(w1.slips[2].fields.department, 'Site B');
assert.strictEqual(w1.slips[2].employeeId, 'EMP-003', 'generated ids continue across sheets');
assert.strictEqual(w1.sheets.length, 2);
// pick one sheet explicitly
var w1b = Schema.parseWorkbook(XLSX, wbOf([['Site A', [WH, ['Ali', 3000, 750, 100, 3650]]], ['Site B', [WH, ['Raju', 1800, 0, 0, 1800]]]]), { sheet: 'Site B' });
assert.strictEqual(w1b.slips.length, 1);
assert.strictEqual(w1b.slips[0].name, 'Raju');
// month sheets -> latest, or the one matching the pay period
var mSheets = [['Sep 2026', [WH, ['Ali', 3000, 750, 100, 3650]]], ['Oct 2026', [WH, ['Ali', 3100, 750, 100, 3750]]]];
var w2 = Schema.parseWorkbook(XLSX, wbOf(mSheets));
assert.strictEqual(w2.sheetName, 'Oct 2026');
assert.strictEqual(w2.slips[0].net, 3750);
assert.ok(codes(w2).indexOf('month-sheet') >= 0);
var w2b = Schema.parseWorkbook(XLSX, wbOf(mSheets), { payPeriod: 'September 2026' });
assert.strictEqual(w2b.sheetName, 'Sep 2026');
// negative deductions
var w3 = Schema.parseWorkbook(XLSX, wbOf([['S', [WH, ['Ali', 3000, 750, -100, 3650], ['Omar', 2000, 500, -50, 2450]]]]));
assert.strictEqual(w3.stats.withErrors, 0, JSON.stringify(w3.slips[0].issues));
assert.strictEqual(w3.slips[0].net, 3650);
assert.ok(codes(w3).indexOf('negative-deductions') >= 0);
// decimal comma
var w4 = Schema.parseWorkbook(XLSX, wbOf([['S', [['Name', 'Basic', 'Housing'], ['Ali', '3.000,50', '750,25'], ['Omar', '2.500,00', '3.000']]]]));
assert.strictEqual(w4.slips[0].gross, 3750.75, 'euro amounts');
assert.strictEqual(w4.slips[1].gross, 5500, '"3.000" in a decimal-comma column is three thousand');
assert.ok(codes(w4).indexOf('decimal-comma') >= 0);
assert.strictEqual(Schema.parseNumber('1,234.50').value, 1234.5);
assert.strictEqual(Schema.parseNumber('3,000').value, 3000);
assert.strictEqual(Schema.parseNumber('1,5').value, 1.5);
assert.strictEqual(Schema.parseNumber('BHD 1.500').value, 1.5, 'three-decimal currencies stay intact');
assert.strictEqual(Schema.parseNumber('150-').value, -150, 'trailing minus');
// Arabic-Indic digits
var w5 = Schema.parseWorkbook(XLSX, wbOf([['S', [['Name', 'Basic', 'Iqama'], ['Ali', '٣٠٠٠', '٢٤٥٦٧٨٩٠١٢']]]]));
assert.strictEqual(w5.slips[0].gross, 3000);
assert.strictEqual(w5.slips[0].fields.nationalId, '2456789012');
assert.strictEqual(w5.stats.withErrors, 0);
// first / last name
var w6 = Schema.parseWorkbook(XLSX, wbOf([['S', [['First Name', 'Last Name', 'Basic'], ['Ali', 'Hassan', 3000]]]]));
assert.strictEqual(w6.slips[0].name, 'Ali Hassan');
assert.ok(codes(w6).indexOf('name-joined') >= 0);
// section heading rows
var w7 = Schema.parseWorkbook(XLSX, wbOf([['S', [['Name', 'Basic', 'Housing'], ['PROJECTS DEPARTMENT', null, null], ['Ali', 3000, 750], ['Total Projects', 3000, 750], ['FINANCE', null, null], ['Sara', 4000, 1000]]]]));
assert.strictEqual(w7.slips.length, 2);
assert.strictEqual(w7.stats.withErrors, 0);
assert.strictEqual(w7.slips[0].fields.department, 'PROJECTS DEPARTMENT');
assert.strictEqual(w7.slips[1].fields.department, 'FINANCE');
// hidden rows skipped, can be included
var hidMut = function (ws) { ws['!rows'] = [null, null, { hidden: true }]; };
var w8 = Schema.parseWorkbook(XLSX, wbOf([['S', [['Name', 'Basic'], ['Ali', 3000], ['Hidden Guy', 2000], ['Omar', 2500]]]], hidMut));
assert.strictEqual(w8.slips.length, 2, 'hidden row skipped');
assert.strictEqual(w8.hiddenRows[0].label, 'Hidden Guy');
assert.ok(codes(w8).indexOf('hidden-rows') >= 0);
var w8b = Schema.parseWorkbook(XLSX, wbOf([['S', [['Name', 'Basic'], ['Ali', 3000], ['Hidden Guy', 2000], ['Omar', 2500]]]], hidMut), { includeHidden: true });
assert.strictEqual(w8b.slips.length, 3);
// error cells reported
var w9 = Schema.parseWorkbook(XLSX, wbOf([['S', [['Name', 'Basic', 'Housing'], ['Ali', 3000, 750], ['Omar', 2000, 500]]]], function (ws) { ws.C3 = { t: 'e', v: 0x2A, w: '#N/A' }; }));
assert.ok(w9.slips[1].hasErrors, 'error cell flags the row');
assert.ok(w9.slips[1].issues.some(function (i) { return /#N\/A/.test(i.message); }));
// net-only sheets
var w10 = Schema.parseWorkbook(XLSX, wbOf([['S', [['Employee Name', 'Net Salary'], ['Ali', 3650], ['Omar', 2500]]]]));
assert.strictEqual(w10.stats.fileErrors, 0, JSON.stringify(w10.issues));
assert.strictEqual(w10.slips[0].net, 3650);
assert.strictEqual(w10.slips[0].earnings[0].label, 'Net Salary');
assert.ok(codes(w10).indexOf('net-only') >= 0);
// header far below notes
var notes = []; for (var q = 0; q < 60; q++) notes.push(['note ' + q]);
var w12 = Schema.parseWorkbook(XLSX, wbOf([['S', notes.concat([['Name', 'Basic'], ['Ali', 3000]])]]));
assert.strictEqual(w12.slips.length, 1);
assert.strictEqual(w12.headerRow, 61);
// forced header row
var w12b = Schema.parseWorkbook(XLSX, wbOf([['S', [['Report'], ['Who', 'Pay'], ['Ali', 3000]]]]), { sheet: 'S', headerRow: 2, overrides: { who: { kind: 'column', key: 'name' }, pay: { kind: 'column', key: 'basicSalary' } } });
assert.strictEqual(w12b.slips.length, 1);
assert.strictEqual(w12b.slips[0].gross, 3000);
// US dates
var w15 = Schema.parseWorkbook(XLSX, wbOf([['S', [['Name', 'Basic', 'Joining Date'], ['Ali', 3000, '03/15/2024'], ['Omar', 2000, '04/05/2024']]]]));
assert.strictEqual(w15.slips[0].fields.joiningDate, '15 Mar 2024');
assert.strictEqual(w15.slips[1].fields.joiningDate, '05 Apr 2024', 'ambiguous value follows the column (month/day)');
assert.ok(codes(w15).indexOf('us-dates') >= 0);
assert.strictEqual(Schema.readDate('31/02/2024').invalid, true, 'impossible dates are flagged, not rolled over');
// hidden sheets ignored; our template helper columns ignored silently
var w16 = Schema.parseWorkbook(XLSX, (function () {
  var w = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(w, XLSX.utils.aoa_to_sheet([['Name', 'Basic', 'Check'], ['Ali', 3000, '✓']]), 'Payslips');
  XLSX.utils.book_append_sheet(w, XLSX.utils.aoa_to_sheet([['Name', 'Basic'], ['Ghost', 1]]), 'Old');
  var b = readWb(wbToBuf(w)); b.Workbook = b.Workbook || {}; b.Workbook.Sheets = [{ name: 'Payslips', Hidden: 0 }, { name: 'Old', Hidden: 1 }];
  return b;
})());
assert.strictEqual(w16.slips.length, 1);
assert.ok(!w16.issues.some(function (i) { return i.code === 'ignored-columns'; }), 'Check column ignored silently');
assert.ok(w16.issues.some(function (i) { return /Hidden sheet/.test(i.message); }));
assert.strictEqual(Schema.monthInfo('أكتوبر 2026').month, 10);
assert.strictEqual(Schema.monthInfo('Oct-26').year, 2026);

// 4. Amount in words
assert.strictEqual(Schema.amountToWords(1234.5, 'SAR'), 'One Thousand Two Hundred Thirty-Four Saudi Riyals and Fifty Halalas Only');
assert.strictEqual(Schema.amountToWords(1, 'SAR'), 'One Saudi Riyal Only');
assert.strictEqual(Schema.amountToWords(0, 'AED'), 'Zero Dirhams Only');
assert.strictEqual(Schema.amountToWords(1000000, 'USD'), 'One Million US Dollars Only');
assert.strictEqual(Schema.formatMoney(1234567.891), '1,234,567.89');

// 4b. Arabic rendering: a slip with Arabic name/title/custom line and Arabic company name
assert.ok(PDF.arabicFontAvailable(), 'Amiri font bundle loaded');
var arSlip = JSON.parse(JSON.stringify(sample.slips[0]));
arSlip.name = 'محمد عبدالله الحربي';
arSlip.fields.jobTitle = 'مهندس موقع';
arSlip.fields.department = 'المشاريع';
arSlip.fields.remarks = 'يشمل بدل العمل الإضافي للموقع';
arSlip.earnings.push({ key: 'custom', label: 'بدل طبيعة عمل', amount: 250 });
var arSettings = Object.assign({}, sample.settings, { companyName: 'شركة المثال للمقاولات المحدودة', address: 'تبوك، المملكة العربية السعودية', footerNote: 'هذه القسيمة صادرة آلياً ولا تحتاج إلى توقيع' });
PDF.TEMPLATES.forEach(function (t) {
  var d = PDF.createDoc(jsPDF, 'a4');
  PDF.renderPayslip(d, { slip: arSlip, settings: arSettings, templateId: t.id, accent: '#1f4e79', options: { signatureLines: true } });
  var ab = Buffer.from(d.output('arraybuffer'));
  assert.ok(ab.length > 20000, 'embedded font makes the file bigger');
  assert.ok(ab.toString('latin1').indexOf('Amiri') >= 0, 'Amiri embedded');
  if (process.env.WRITE_PDFS && t.id === 'executive') fs.writeFileSync(path.join(__dirname, '..', 'samples', 'Sample_Arabic.pdf'), ab);
});
var plain = PDF.createDoc(jsPDF, 'a4');
PDF.renderPayslip(plain, { slip: sample.slips[0], settings: sample.settings, templateId: 'classic', accent: '#1f4e79', options: {} });
assert.ok(Buffer.from(plain.output('arraybuffer')).toString('latin1').indexOf('Amiri') < 0, 'font not embedded when no Arabic');
var tAr = Date.now();
for (var k = 0; k < 20; k++) { var dk = PDF.createDoc(jsPDF, 'a4'); PDF.renderPayslip(dk, { slip: arSlip, settings: arSettings, templateId: 'modern', accent: '#1f4e79', options: {} }); dk.output('arraybuffer'); }
console.log('20 single Arabic payslips (font embedded each time):', Date.now() - tAr, 'ms');

// 5. Render every template for every sample slip, and a 1000-page document.
var outDir = path.join(__dirname, '..', 'samples');
PDF.TEMPLATES.forEach(function (t) {
  var doc = PDF.createDoc(jsPDF, 'a4');
  sample.slips.forEach(function (slip, i) {
    if (i) doc.addPage();
    PDF.renderPayslip(doc, { slip: slip, settings: sample.settings, templateId: t.id, accent: '#1f4e79', logo: null, stamp: null, options: { amountInWords: true, signatureLines: true } });
  });
  var buf = Buffer.from(doc.output('arraybuffer'));
  assert.ok(buf.length > 5000);
});
if (process.env.WRITE_PDFS) {
  // one page per template, for a quick look at every design
  var all = PDF.createDoc(jsPDF, 'a4');
  PDF.TEMPLATES.forEach(function (t, i) {
    if (i) all.addPage();
    PDF.renderPayslip(all, { slip: sample.slips[0], settings: sample.settings, templateId: t.id, accent: '#1f4e79', options: { signatureLines: true } });
  });
  fs.writeFileSync(path.join(outDir, 'Sample_All_Templates.pdf'), Buffer.from(all.output('arraybuffer')));
}
var big = Schema.parseWorkbook(XLSX, readWb(wbToBuf(Schema.buildTemplateWorkbook(XLSX, { sampleRows: Schema.sampleRows(1000) }))));
assert.strictEqual(big.slips.length, 1000);
var t0 = Date.now();
var bigDoc = PDF.createDoc(jsPDF, 'a4');
big.slips.forEach(function (slip, i) {
  if (i) bigDoc.addPage();
  PDF.renderPayslip(bigDoc, { slip: slip, settings: big.settings, templateId: 'modern', accent: '#1f4e79', options: {} });
});
var bigBuf = Buffer.from(bigDoc.output('arraybuffer'));
console.log('1000 payslips rendered in', Date.now() - t0, 'ms,', (bigBuf.length / 1024 / 1024).toFixed(1), 'MB');

// 20 real-world workbooks (extra overtime / attendance / loans / employee sheets, odd layouts, non-payroll files)
var scenarioResults = require(path.join(__dirname, 'scenarios.js')).run(XLSX, Schema);
scenarioResults.forEach(function (r) {
  assert.ok(r.pass, 'Scenario ' + r.id + ' (' + r.title + '): ' + r.fails.join(' | ') + ' [got ' + r.got + ']');
});

// A payroll sheet whose name mentions overtime is still the payroll; an employee list with a contract salary
// column next to the payroll only supplies details (no double basic salary).
(function () {
  var one = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(one, XLSX.utils.aoa_to_sheet([['Emp ID', 'Name', 'Basic', 'OT Hours', 'OT Amount'], ['E1', 'Ali', 3000, 5, 150]]), 'Salary & OT');
  var p1 = Schema.parseWorkbook(XLSX, readWb(wbToBuf(one)));
  assert.strictEqual(p1.slips.length, 1, JSON.stringify(p1.issues));
  assert.strictEqual(p1.slips[0].net, 3150);
  var two = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(two, XLSX.utils.aoa_to_sheet([['Emp ID', 'Name', 'Basic', 'Housing', 'Net'], ['E1', 'Ali', 3000, 750, 3750], ['E2', 'Omar', 2000, 500, 2500]]), 'October');
  XLSX.utils.book_append_sheet(two, XLSX.utils.aoa_to_sheet([['Emp ID', 'Name', 'Basic Salary', 'Mobile', 'IBAN'], ['E1', 'Ali', 3000, '0551234567', 'SA0380000000608010167519'], ['E2', 'Omar', 2000, '0559876543', '']]), 'Staff');
  var p2 = Schema.parseWorkbook(XLSX, readWb(wbToBuf(two)));
  assert.strictEqual(p2.sheetName, 'October');
  assert.strictEqual(p2.slips[0].net, 3750, 'basic salary not added twice');
  assert.strictEqual(p2.slips[0].fields.mobile, '966551234567', 'mobile from the staff sheet');
  assert.strictEqual(p2.supplements[0].role, 'employees');
})();

// Safety net: the delivery journal never loses a tick or revives an undone one, whatever order copies merge in.
var Vault = require(path.join(__dirname, '..', 'js', 'vault.js'));
var Delivery = require(path.join(__dirname, '..', 'js', 'delivery.js'));
(function () {
  var t = function (m) { return '2026-10-17T10:' + (m < 10 ? '0' : '') + m + ':00.000Z'; };
  var tabA = [{ timestamp: t(1), employeeId: 'E1', action: 'sent', channel: 'outlook' }, { timestamp: t(3), employeeId: 'E2', action: 'sent' }];
  var tabB = [{ timestamp: t(1), employeeId: 'E1', action: 'sent', channel: 'outlook' }, { timestamp: t(2), employeeId: 'E3', action: 'sent' }, { timestamp: t(4), employeeId: 'E1', action: 'undo' }];
  var merged = Vault.mergeLogs(tabA, tabB);
  assert.strictEqual(merged.length, 4, 'duplicates collapse');
  assert.deepStrictEqual(Object.keys(Vault.sentMap(merged)).sort(), ['E2', 'E3'], 'undo wins over the earlier send');
  assert.deepStrictEqual(Vault.mergeLogs(tabB, tabA).map(function (e) { return e.id; }), merged.map(function (e) { return e.id; }), 'merge order does not matter');
  var cleared = Vault.mergeLogs(merged, [{ timestamp: t(5), employeeId: '*', action: 'clear' }, { timestamp: t(6), employeeId: 'E9', action: 'sent' }]);
  assert.deepStrictEqual(Object.keys(Vault.sentMap(cleared)), ['E9'], 'clear resets everything before it');
  // logs saved by older versions: plain list, undone rows deleted or flagged
  var legacy = Vault.normalizeLog([{ timestamp: t(1), employeeId: 'E1', channel: 'outlook' }, { timestamp: t(2), employeeId: 'E2', undone: true }]);
  assert.strictEqual(legacy.length, 1);
  assert.strictEqual(legacy[0].action, 'sent');
  // a delivery log CSV round-trips (restore after a browser reset)
  var csv = Delivery.deliveryLogCsv(cleared);
  var back = Vault.logFromCsv(csv);
  assert.deepStrictEqual(Object.keys(Vault.sentMap(back)), ['E9'], 'CSV restore keeps the same ticks');
  assert.deepStrictEqual(Vault.parseCsv('a,"b ""q"", c"\r\n1,2\n'), [['a', 'b "q", c'], ['1', '2']]);
  assert.throws(function () { Vault.logFromCsv('foo,bar\n1,2'); }, /not a delivery log/);
  // settings fingerprint: key order and the monthly period fields do not matter, real changes do
  var base = { savedAt: 'x', templateId: 'classic', company: { companyName: 'Acme', payPeriod: 'October 2026', crNumber: '1' } };
  var same = { company: { crNumber: '1', payPeriod: 'November 2026', companyName: 'Acme' }, templateId: 'classic', savedAt: 'y' };
  assert.strictEqual(Vault.settingsFingerprint(base), Vault.settingsFingerprint(same));
  assert.notStrictEqual(Vault.settingsFingerprint(base), Vault.settingsFingerprint({ templateId: 'modern', company: base.company }));
  assert.strictEqual(Vault.timeAgo(t(0), Date.parse(t(5))), '5 min ago');
})();

// PWA: every file the page loads is in the service worker cache list, and the manifest icons exist.
(function () {
  var root = path.join(__dirname, '..');
  var sw = fs.readFileSync(path.join(root, 'sw.js'), 'utf8');
  var listed = JSON.parse(sw.match(/var ASSETS = (\[[\s\S]*?\]);/)[1].replace(/'/g, '"'));
  var html = fs.readFileSync(path.join(root, 'index.html'), 'utf8');
  var needed = [];
  html.replace(/<(?:script|link)[^>]+(?:src|href)="([^":#]+)"/g, function (m, u) { needed.push(u); });
  var css = fs.readFileSync(path.join(root, 'css', 'styles.css'), 'utf8');
  css.replace(/url\("(fonts\/[^"]+)"\)/g, function (m, u) { needed.push('css/' + u); });
  needed.push('js/vendor/exceljs.min.js', 'manifest.webmanifest');
  var manifest = JSON.parse(fs.readFileSync(path.join(root, 'manifest.webmanifest'), 'utf8'));
  manifest.icons.forEach(function (i) { needed.push(i.src); });
  needed.forEach(function (u) {
    assert.ok(listed.indexOf(u) >= 0, 'sw.js ASSETS is missing ' + u);
  });
  listed.forEach(function (u) {
    if (u !== './') assert.ok(fs.existsSync(path.join(root, u)), 'sw.js lists a file that does not exist: ' + u);
  });
  assert.ok(manifest.icons.some(function (i) { return i.purpose === 'maskable'; }));
})();

// Styled ExcelJS template: round-trips through the parser, keeps validations, carries the logo.
var ExcelJS = require(path.join(__dirname, '..', 'js', 'vendor', 'exceljs.min.js'));
var JSZip = require(path.join(__dirname, '..', 'js', 'vendor', 'jszip.min.js'));
var PayslipTemplate = require(path.join(__dirname, '..', 'js', 'template-xlsx.js'));
var XlsxImages = require(path.join(__dirname, '..', 'js', 'xlsx-images.js'));
var PNG_1PX = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==';
(async function () {
  var empty = Schema.parseWorkbook(XLSX, readWb(Buffer.from(await PayslipTemplate.build(ExcelJS, {}))));
  assert.strictEqual(empty.slips.length, 0, 'styled template: example row ignored');
  assert.ok(empty.settingsFound);
  assert.ok(!empty.issues.some(function (i) { return i.level === 'warn'; }), JSON.stringify(empty.issues));

  var full = Schema.parseWorkbook(XLSX, readWb(Buffer.from(await PayslipTemplate.build(ExcelJS, { records: Schema.sampleRecords(25), sample: true }))));
  assert.strictEqual(full.slips.length, 25);
  assert.strictEqual(full.stats.withErrors, 0, JSON.stringify(full.slips.filter(function (s) { return s.hasErrors; }).map(function (s) { return s.issues; })));
  assert.strictEqual(full.stats.withWarnings, 0);
  assert.strictEqual(full.settings.companyName, 'Example Contracting Co. Ltd.');
  assert.ok(/\d{4}/.test(full.slips[0].fields.joiningDate), 'dates read back');

  var mineBuf = Buffer.from(await PayslipTemplate.build(ExcelJS, { settings: { companyName: 'Acme Build' }, logo: { dataUrl: PNG_1PX, width: 1, height: 1 }, stamp: { dataUrl: PNG_1PX, width: 1, height: 1 } }));
  assert.strictEqual(Schema.parseWorkbook(XLSX, readWb(mineBuf)).settings.companyName, 'Acme Build');
  var imgs = await XlsxImages.extractImages(JSZip, mineBuf);
  assert.strictEqual(imgs.length, 2, 'logo and stamp embedded');

  var check = new ExcelJS.Workbook();
  await check.xlsx.load(mineBuf);
  var ws = check.getWorksheet('Payslips');
  assert.ok(Object.keys(ws.dataValidations.model).length >= 10, 'data validation kept');
  assert.ok(check.getWorksheet('Lists').state === 'hidden');
  console.log('All tests passed.');
})().catch(function (e) { console.error(e); process.exit(1); });
