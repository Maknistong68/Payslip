#!/usr/bin/env node
/* Builds samples/Payslip_Template.xlsx and samples/Sample_Payslips.xlsx with the same
 * styled, validated template the web page downloads. Run: node tools/build-samples.js */
'use strict';
var path = require('path');
var fs = require('fs');
var ExcelJS = require(path.join(__dirname, '..', 'js', 'vendor', 'exceljs.min.js'));
var Schema = require(path.join(__dirname, '..', 'js', 'schema.js'));
var PayslipTemplate = require(path.join(__dirname, '..', 'js', 'template-xlsx.js'));

var outDir = path.join(__dirname, '..', 'samples');
if (!fs.existsSync(outDir)) fs.mkdirSync(outDir);

function write(name, opts) {
  return PayslipTemplate.build(ExcelJS, opts).then(function (buf) {
    var file = path.join(outDir, name);
    fs.writeFileSync(file, Buffer.from(buf));
    console.log('Wrote', file);
  });
}

write('Payslip_Template.xlsx', {})
  .then(function () { return write('Sample_Payslips.xlsx', { records: Schema.sampleRecords(25), sample: true }); })
  .catch(function (e) { console.error(e); process.exit(1); });
