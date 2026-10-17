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
assert.ok(tpl.issues.some(function (i) { return /Company Name/.test(i.message); }), 'missing company name reported');

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
assert.ok(b.issues.some(function (i) { return /CR Number/.test(i.message) && i.level === 'error'; }));
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
assert.ok(b.slips[2].issues.some(function (i) { return /Employee Name/.test(i.message); }));
assert.ok(b.slips[2].issues.some(function (i) { return /exceed/.test(i.message); }));
assert.ok(b.slips[3].issues.some(function (i) { return /Arabic/.test(i.message); }));
assert.ok(!b.slips[3].issues.some(function (i) { return /IBAN/.test(i.message); }), 'valid SA IBAN');
assert.ok(b.slips[4].issues.some(function (i) { return /negative/.test(i.message); }));
assert.strictEqual(b.stats.withErrors, 4);

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
