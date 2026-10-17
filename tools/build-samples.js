#!/usr/bin/env node
/* Builds samples/Payslip_Template.xlsx and samples/Sample_Payslips.xlsx using the same
 * schema code that the web page uses. Run: node tools/build-samples.js */
'use strict';
var path = require('path');
var fs = require('fs');
var XLSX = require(path.join(__dirname, '..', 'js', 'vendor', 'xlsx.full.min.js'));
var Schema = require(path.join(__dirname, '..', 'js', 'schema.js'));

var outDir = path.join(__dirname, '..', 'samples');
if (!fs.existsSync(outDir)) fs.mkdirSync(outDir);

var tpl = Schema.buildTemplateWorkbook(XLSX);
fs.writeFileSync(path.join(outDir, 'Payslip_Template.xlsx'), XLSX.write(tpl, { type: 'buffer', bookType: 'xlsx', cellDates: true }));

var sample = Schema.buildTemplateWorkbook(XLSX, { sampleRows: Schema.sampleRows(25) });
fs.writeFileSync(path.join(outDir, 'Sample_Payslips.xlsx'), XLSX.write(sample, { type: 'buffer', bookType: 'xlsx', cellDates: true }));

console.log('Wrote', path.join(outDir, 'Payslip_Template.xlsx'));
console.log('Wrote', path.join(outDir, 'Sample_Payslips.xlsx'));
