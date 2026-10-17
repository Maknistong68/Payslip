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

function readWb(buf) { return XLSX.read(buf, { type: 'buffer', cellDates: true }); }
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
assert.ok(b.slips[3].issues.some(function (i) { return /Arabic/.test(i.message); }));
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

// Totally unrecognisable headers (e.g. Arabic) -> error + mapping choices still work
var ar = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(ar, XLSX.utils.aoa_to_sheet([
  ['الرقم', 'الاسم', 'الراتب الأساسي', 'بدل السكن', 'الصافي'],
  ['1', 'Ahmed', 3000, 500, 3500]
]), 'ورقة1');
var a = Schema.parseWorkbook(XLSX, readWb(wbToBuf(ar)));
assert.ok(a.issues.some(function (i) { return i.code === 'no-header' || i.code === 'missing-name'; }), JSON.stringify(a.issues));
assert.strictEqual(a.columns.length, 5);
var a2 = Schema.parseWorkbook(XLSX, readWb(wbToBuf(ar)), { overrides: {
  [Schema.normalizeHeader('الاسم')]: { kind: 'column', key: 'name' },
  [Schema.normalizeHeader('الراتب الأساسي')]: { kind: 'column', key: 'basicSalary' },
  [Schema.normalizeHeader('بدل السكن')]: { kind: 'column', key: 'housing' }
} });
assert.strictEqual(a2.slips.length, 1);
assert.strictEqual(a2.slips[0].net, 3500);
assert.strictEqual(a2.stats.fileErrors, 0, JSON.stringify(a2.issues));

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

// 4. Amount in words
assert.strictEqual(Schema.amountToWords(1234.5, 'SAR'), 'One Thousand Two Hundred Thirty-Four Saudi Riyals and Fifty Halalas Only');
assert.strictEqual(Schema.amountToWords(1, 'SAR'), 'One Saudi Riyal Only');
assert.strictEqual(Schema.amountToWords(0, 'AED'), 'Zero Dirhams Only');
assert.strictEqual(Schema.amountToWords(1000000, 'USD'), 'One Million US Dollars Only');
assert.strictEqual(Schema.formatMoney(1234567.891), '1,234,567.89');

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
  if (process.env.WRITE_PDFS) fs.writeFileSync(path.join(outDir, 'Sample_' + t.name + '.pdf'), buf);
});
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

console.log('All tests passed.');
