#!/usr/bin/env node
/*
 * 20 real-world workbooks a contractor might upload, each with the behaviour we expect.
 * Used by tools/test.js; run directly to print the result table:  node tools/scenarios.js
 */
'use strict';
var path = require('path');

function day(n) { return n; }

// Builds a workbook from [[sheetName, rows, opts]] and reads it back the way the app does.
function book(XLSX, sheets) {
  var wb = XLSX.utils.book_new();
  sheets.forEach(function (s) {
    var ws = XLSX.utils.aoa_to_sheet(s[1]);
    var opts = s[2] || {};
    if (opts.hiddenRows) ws['!rows'] = opts.hiddenRows.map(function (h) { return h ? { hidden: true } : null; });
    XLSX.utils.book_append_sheet(wb, ws, s[0]);
    if (opts.hidden) {
      wb.Workbook = wb.Workbook || { Sheets: [] };
      while (wb.Workbook.Sheets.length < wb.SheetNames.length) wb.Workbook.Sheets.push({});
      wb.Workbook.Sheets[wb.SheetNames.length - 1] = { name: s[0], Hidden: 1 };
    }
  });
  var buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
  return XLSX.read(buf, { type: 'buffer', cellDates: true, cellStyles: true });
}

function timesheet(rows, days) {
  var head = ['Emp ID', 'Name'];
  for (var d = 1; d <= days; d++) head.push(day(d));
  return [head].concat(rows.map(function (r) { return [r[0], r[1]].concat(r[2]); }));
}

function marks(present, absent, off, leave) {
  var out = [];
  for (var i = 0; i < present; i++) out.push('P');
  for (var j = 0; j < absent; j++) out.push('A');
  for (var k = 0; k < (leave || 0); k++) out.push('L');
  for (var m = 0; m < off; m++) out.push('W');
  return out;
}

function slip(p, id) { return p.slips.filter(function (s) { return s.employeeId === id || s.name === id; })[0]; }
function line(s, re) { return s && s.earnings.concat(s.deductions).filter(function (l) { return re.test(l.label); })[0]; }
function codes(p) { return p.issues.map(function (i) { return i.code; }); }
function errors(p) { return p.issues.filter(function (i) { return i.level === 'error'; }); }
function supp(p, name) { return (p.supplements || []).filter(function (s) { return s.sheet === name; })[0]; }
function netMismatch(p) { return p.slips.filter(function (s) { return s.issues.some(function (i) { return /does not match the calculated net/.test(i.message); }); }).length; }

function scenarios(XLSX, Schema) {
  return [
    {
      id: 1, title: 'Our template, filled (25 employees)',
      wb: function () { return XLSX.read(XLSX.write(Schema.buildTemplateWorkbook(XLSX, { sampleRows: Schema.sampleRows(25) }), { type: 'buffer', bookType: 'xlsx', cellDates: true }), { type: 'buffer', cellDates: true, cellStyles: true }); },
      expect: '25 payslips, no errors',
      check: function (p, ok) {
        ok(p.slips.length === 25, '25 slips');
        ok(p.stats.withErrors === 0, 'no error rows');
        ok(!errors(p).length, 'no file errors');
      }
    },
    {
      id: 2, title: 'Payroll + separate "Overtime" sheet; net already includes overtime',
      wb: function () {
        return book(XLSX, [
          ['Payroll', [['Emp ID', 'Employee Name', 'Basic', 'Housing', 'Net Salary'], ['E1', 'Ali Hassan', 3000, 750, 4050], ['E2', 'Omar Khan', 2000, 500, 2650], ['E3', 'Raju Das', 1800, 0, 1800]]],
          ['Overtime', [['Emp ID', 'Name', 'OT Hours', 'OT Amount'], ['E1', 'Ali Hassan', 10, 300], ['E2', 'Omar Khan', 5, 150]]]
        ]);
      },
      expect: 'overtime added to the payslips automatically (the net column proves it), hours too',
      check: function (p, ok) {
        ok(p.slips.length === 3, '3 slips');
        var s = supp(p, 'Overtime');
        ok(s && s.role === 'overtime' && s.enabled && s.auto, 'Overtime sheet used automatically');
        var e1 = slip(p, 'E1');
        ok(line(e1, /overtime/i) && line(e1, /overtime/i).amount === 300, 'E1 has overtime 300');
        ok(e1.fields.overtimeHours === 10, 'E1 overtime hours 10');
        ok(netMismatch(p) === 0, 'net matches for everyone');
      }
    },
    {
      id: 3, title: 'Payroll has an Overtime column + a daily "OT Log" sheet (many rows per person)',
      wb: function () {
        return book(XLSX, [
          ['Salaries', [['Emp ID', 'Name', 'Basic', 'Overtime', 'Net'], ['E1', 'Ali Hassan', 3000, 300, 3300], ['E2', 'Omar Khan', 2000, 0, 2000]]],
          ['OT Log', [['Emp ID', 'Date', 'OT Hours'], ['E1', '03/10/2026', 4], ['E1', '10/10/2026', 6], ['E2', '04/10/2026', 0]]]
        ]);
      },
      expect: 'hours summed per person; overtime pay not counted twice',
      check: function (p, ok) {
        var e1 = slip(p, 'E1');
        ok(e1 && e1.fields.overtimeHours === 10, 'E1 hours summed to 10');
        ok(e1.earnings.filter(function (l) { return /overtime/i.test(l.label); }).length === 1, 'overtime pay once');
        ok(netMismatch(p) === 0, 'net matches');
      }
    },
    {
      id: 4, title: 'Payroll + "Attendance" timesheet grid (days 1-30 with P / A / W)',
      wb: function () {
        return book(XLSX, [
          ['Payroll', [['ID', 'Name', 'Basic', 'Net'], ['E1', 'Ali Hassan', 3000, 3000], ['E2', 'Omar Khan', 2000, 2000]]],
          ['Attendance', timesheet([['E1', 'Ali Hassan', marks(26, 0, 4)], ['E2', 'Omar Khan', marks(23, 2, 4, 1)]], 30)]
        ]);
      },
      expect: 'working days and absent days counted from the grid',
      check: function (p, ok) {
        var s = supp(p, 'Attendance');
        ok(s && s.role === 'attendance' && s.enabled, 'attendance sheet used');
        ok(slip(p, 'E1').fields.workingDays === 26 && slip(p, 'E1').fields.absentDays === 0, 'E1 26/0');
        ok(slip(p, 'E2').fields.workingDays === 23 && slip(p, 'E2').fields.absentDays === 2, 'E2 23/2');
      }
    },
    {
      id: 5, title: 'Payroll with IDs only + "Employees" master list (names, Iqama, mobile, email)',
      wb: function () {
        return book(XLSX, [
          ['Pay', [['Emp No', 'Basic', 'Housing', 'Net'], ['101', 3000, 750, 3750], ['102', 2000, 500, 2500], ['103', 1800, 0, 1800]]],
          ['Employees', [['Emp No', 'Full Name', 'Iqama No', 'Mobile', 'Email', 'Nationality'],
            ['101', 'Ali Hassan', '2456789012', '0551234567', 'ali@example.com', 'Egyptian'],
            ['102', 'Omar Khan', '2456789013', '0559876543', 'omar@example.com', 'Pakistani'],
            ['103', 'Raju Das', '2456789014', '0551112233', '', 'Indian']]]
        ]);
      },
      expect: 'names and contacts filled from the master list by employee number',
      check: function (p, ok) {
        ok(p.slips.length === 3 && p.stats.withErrors === 0, '3 slips, no errors');
        var a = slip(p, '101');
        ok(a && a.name === 'Ali Hassan', 'name filled');
        ok(a.fields.mobile === '966551234567' && a.fields.email === 'ali@example.com' && a.fields.nationalId === '2456789012', 'contacts + Iqama filled');
        ok(p.stats.withMobile === 3, 'all have a mobile');
        ok(codes(p).indexOf('missing-name') < 0, 'no "missing name" error');
      }
    },
    {
      id: 6, title: 'No IDs anywhere: "Contacts" sheet matched by name (different case and spacing)',
      wb: function () {
        return book(XLSX, [
          ['Sheet1', [['Name', 'Basic'], ['Ahmed Hassan', 3000], ['Sara Ali', 2500], ['John Smith', 2000]]],
          ['Contacts', [['Employee Name', 'WhatsApp'], ['AHMED  HASSAN', '0551112222'], ['sara ali', '0553334444'], ['Mike Ross', '0555556666']]]
        ]);
      },
      expect: 'matched by name; the person missing from payroll is reported',
      check: function (p, ok) {
        var s = supp(p, 'Contacts');
        ok(s && s.keyBy === 'name' && s.matched === 2, 'matched 2 by name');
        ok(slip(p, 'Ahmed Hassan').fields.mobile === '966551112222', 'Ahmed mobile');
        ok(!slip(p, 'John Smith').fields.mobile, 'John has none');
        ok(s.unmatched.indexOf('Mike Ross') >= 0, 'Mike Ross reported as not in payroll');
      }
    },
    {
      id: 7, title: '"Summary" sheet (department totals) comes before the payroll sheet',
      wb: function () {
        return book(XLSX, [
          ['Summary', [['Department', 'Headcount', 'Total Salary'], ['Projects', 12, 45000], ['Finance', 3, 15000]]],
          ['Payroll', [['Name', 'Basic', 'Housing'], ['Ali Hassan', 3000, 750], ['Omar Khan', 2000, 500]]]
        ]);
      },
      expect: 'payroll sheet used alone, summary ignored',
      check: function (p, ok) {
        ok(p.sheetName === 'Payroll' && p.slips.length === 2, 'Payroll sheet, 2 slips');
        ok(p.sheets.filter(function (s) { return s.name === 'Summary'; })[0].role === 'summary', 'summary recognised');
        ok(!errors(p).length, 'no errors');
      }
    },
    {
      id: 8, title: 'Two site sheets + one "Overtime" sheet for both + a "Summary"',
      wb: function () {
        var H = ['Emp ID', 'Name', 'Basic', 'Net Pay'];
        return book(XLSX, [
          ['Site A', [H, ['A1', 'Ali Hassan', 3000, 3200], ['A2', 'Omar Khan', 2000, 2000]]],
          ['Site B', [H, ['B1', 'Raju Das', 1800, 1900], ['B2', 'Jose Cruz', 2200, 2200]]],
          ['Overtime', [['Emp ID', 'OT Pay'], ['A1', 200], ['B1', 100]]],
          ['Summary', [['Site', 'Employees', 'Total Net'], ['Site A', 2, 5200], ['Site B', 2, 4100]]]
        ]);
      },
      expect: 'sites combined, overtime added, summary ignored',
      check: function (p, ok) {
        ok(p.sheetMode === 'combined' && p.slips.length === 4, 'combined 4 slips');
        ok(supp(p, 'Overtime') && supp(p, 'Overtime').enabled, 'overtime used');
        ok(netMismatch(p) === 0, 'net matches');
      }
    },
    {
      id: 9, title: 'Monthly sheets (Sep, Oct) with monthly overtime sheets (OT Sep, OT Oct)',
      wb: function () {
        var H = ['Emp ID', 'Name', 'Basic', 'Net'];
        return book(XLSX, [
          ['Sep 2026', [H, ['E1', 'Ali Hassan', 3000, 3100]]],
          ['Oct 2026', [H, ['E1', 'Ali Hassan', 3000, 3250]]],
          ['OT Sep', [['Emp ID', 'OT Amount'], ['E1', 100]]],
          ['OT Oct', [['Emp ID', 'OT Amount'], ['E1', 250]]]
        ]);
      },
      expect: 'October payroll with October overtime only',
      check: function (p, ok) {
        ok(p.sheetName === 'Oct 2026', 'Oct used');
        ok(supp(p, 'OT Oct') && supp(p, 'OT Oct').enabled, 'OT Oct used');
        ok(!supp(p, 'OT Sep') || !supp(p, 'OT Sep').enabled, 'OT Sep not used');
        ok(line(slip(p, 'E1'), /overtime/i).amount === 250, 'overtime 250');
      }
    },
    {
      id: 10, title: 'Not payroll at all: an inventory list',
      wb: function () {
        return book(XLSX, [['Stock', [['Item', 'Qty', 'Unit Price', 'Total'], ['Cement', 10, 20, 200], ['Steel', 5, 300, 1500]]]]);
      },
      expect: 'refused with a clear "not a payroll sheet" message',
      check: function (p, ok) {
        ok(codes(p).indexOf('not-payroll') >= 0, 'not-payroll error');
        ok(p.slips.length === 0, 'no payslips');
      }
    },
    {
      id: 11, title: 'Attendance only (timesheet, no amounts)',
      wb: function () {
        return book(XLSX, [['Timesheet', timesheet([['E1', 'Ali Hassan', marks(26, 0, 4)], ['E2', 'Omar Khan', marks(25, 1, 4)]], 30)]]);
      },
      expect: 'explains that payslips need amounts, not only attendance',
      check: function (p, ok) {
        ok(codes(p).indexOf('attendance-only') >= 0, 'attendance-only error');
        ok(p.slips.length === 0, 'no payslips');
      }
    },
    {
      id: 12, title: 'Employee list only (no salaries)',
      wb: function () {
        return book(XLSX, [['Staff', [['Emp ID', 'Name', 'Iqama', 'Mobile', 'Nationality'], ['E1', 'Ali Hassan', '2456789012', '0551234567', 'Egyptian'], ['E2', 'Omar Khan', '2456789013', '0559876543', 'Pakistani']]]]);
      },
      expect: 'explains that salary columns are missing',
      check: function (p, ok) {
        ok(codes(p).indexOf('no-salary') >= 0, 'no-salary error');
        ok(p.slips.length === 0, 'no payslips');
      }
    },
    {
      id: 13, title: 'One sheet per employee, laid out like a payslip form',
      wb: function () {
        var form = function (name, id, basic, housing, gosi) {
          return [['PAYSLIP - OCTOBER 2026'], [], ['Employee Name', name], ['Employee ID', id], ['Basic Salary', basic], ['Housing Allowance', housing], ['GOSI', gosi], ['Net Salary', basic + housing - gosi]];
        };
        return book(XLSX, [['Ali', form('Ali Hassan', 'E1', 3000, 750, 100)], ['Omar', form('Omar Khan', 'E2', 2000, 500, 0)], ['Sara', form('Sara Ali', 'E3', 4000, 1000, 390)]]);
      },
      expect: 'each sheet becomes one payslip',
      check: function (p, ok) {
        ok(p.slips.length === 3, '3 slips');
        ok(slip(p, 'E3') && slip(p, 'E3').net === 4610, 'Sara net 4610');
        ok(netMismatch(p) === 0 && p.stats.withErrors === 0, 'clean');
        ok(slip(p, 'E1').fields.department !== 'Ali', 'sheet name not used as department');
      }
    },
    {
      id: 14, title: 'Sideways table: employees in columns, pay items in rows',
      wb: function () {
        return book(XLSX, [['Oct', [['', 'Ali Hassan', 'Omar Khan', 'Raju Das'], ['Emp ID', 'E1', 'E2', 'E3'], ['Basic', 3000, 2000, 1800], ['Housing', 750, 500, 0], ['Net', 3750, 2500, 1800]]]]);
      },
      expect: 'turned the right way round automatically',
      check: function (p, ok) {
        ok(p.slips.length === 3, '3 slips');
        ok(slip(p, 'E2') && slip(p, 'E2').name === 'Omar Khan' && slip(p, 'E2').net === 2500, 'Omar 2500');
        ok(netMismatch(p) === 0, 'net matches');
      }
    },
    {
      id: 15, title: 'Abbreviated headers (Emp#, B.Sal, H.A, T.A, OT Amt, Ded, Net Pay)',
      wb: function () {
        return book(XLSX, [['Data', [['Emp#', 'Emp Name', 'B.Sal', 'H.A', 'T.A', 'OT Amt', 'Ded', 'Net Pay'], ['E1', 'Ali Hassan', 3000, 750, 300, 200, 100, 4150], ['E2', 'Omar Khan', 2000, 500, 300, 0, 0, 2800]]]]);
      },
      expect: 'all columns understood',
      check: function (p, ok) {
        var keys = p.columns.filter(function (c) { return c.kind === 'column'; }).map(function (c) { return c.def.key; });
        ok(['employeeId', 'name', 'basicSalary', 'housing', 'transport', 'overtimePay', 'netSalary'].every(function (k) { return keys.indexOf(k) >= 0; }), 'columns mapped: ' + keys.join(','));
        ok(netMismatch(p) === 0 && p.stats.withErrors === 0, 'net matches');
      }
    },
    {
      id: 16, title: 'Printed report: header repeated on every page, page totals in between',
      wb: function () {
        var H = ['Emp ID', 'Name', 'Basic', 'Housing'];
        return book(XLSX, [['Report', [['ACME CONTRACTING - PAYROLL OCT 2026'], H, ['E1', 'Ali Hassan', 3000, 750], ['E2', 'Omar Khan', 2000, 500], ['E3', 'Raju Das', 1800, 0], ['Page Total', '', 6800, 1250], [], H, ['E4', 'Jose Cruz', 2200, 0], ['E5', 'Sara Ali', 4000, 1000], ['Grand Total', '', 13000, 2250]]]]);
      },
      expect: 'repeated headers and page totals skipped, 5 payslips',
      check: function (p, ok) {
        ok(p.slips.length === 5, '5 slips (got ' + p.slips.length + ')');
        ok(p.stats.withErrors === 0, 'no error rows');
      }
    },
    {
      id: 17, title: 'Broken formulas (#DIV/0!, #REF!) in a few cells',
      wb: function () {
        var wb = book(XLSX, [['Payroll', [['Name', 'Basic', 'Housing', 'Transport'], ['Ali Hassan', 3000, 750, 300], ['Omar Khan', 2000, 500, 300], ['Raju Das', 1800, 0, 300]]]]);
        wb.Sheets.Payroll.C3 = { t: 'e', v: 7, w: '#DIV/0!' };
        wb.Sheets.Payroll.D4 = { t: 'e', v: 23, w: '#REF!' };
        return wb;
      },
      expect: 'the two rows are flagged with the cell, the rest are fine',
      check: function (p, ok) {
        ok(p.slips.length === 3 && p.stats.withErrors === 2, '2 error rows');
        ok(slip(p, 'Omar Khan').issues.some(function (i) { return /#DIV\/0!/.test(i.message); }), 'Omar flagged #DIV/0!');
      }
    },
    {
      id: 18, title: 'Arabic sheet names: "الرواتب" + "السلف" (loans) sheet',
      wb: function () {
        return book(XLSX, [
          ['الرواتب', [['الرقم الوظيفي', 'الاسم', 'الراتب الأساسي', 'بدل السكن', 'صافي الراتب'], ['E1', 'علي حسن', 3000, 750, 3250], ['E2', 'عمر خان', 2000, 500, 2500]]],
          ['السلف', [['الرقم الوظيفي', 'السلفة'], ['E1', 500]]]
        ]);
      },
      expect: 'loan deduction added (the net proves it)',
      check: function (p, ok) {
        var s = supp(p, 'السلف');
        ok(s && s.enabled, 'loans sheet used');
        ok(line(slip(p, 'E1'), /loan|سلف/i) && slip(p, 'E1').net === 3250, 'E1 net 3250 with loan');
        ok(netMismatch(p) === 0, 'net matches');
      }
    },
    {
      id: 19, title: 'CSV from another system: semicolons and decimal commas',
      csv: 'Name;Basic;Housing;Net\nAli Hassan;3.000,50;750,00;3.750,50\nOmar Khan;2.000,00;500,00;2.500,00\n',
      expect: 'amounts read correctly',
      check: function (p, ok) {
        ok(p.slips.length === 2, '2 slips');
        ok(slip(p, 'Ali Hassan') && slip(p, 'Ali Hassan').net === 3750.5, 'Ali net 3750.50');
        ok(netMismatch(p) === 0, 'net matches');
      }
    },
    {
      id: 20, title: 'Busy workbook: instructions, hidden calc sheet, empty chart sheet, pivot, lists, and the payroll on sheet 6',
      wb: function () {
        return book(XLSX, [
          ['Instructions', [['How to fill this file'], ['Enter one row per employee.']]],
          ['Calc', [['Rate', 1.5], ['Hours', 240]], { hidden: true }],
          ['Chart', [[]]],
          ['Pivot', [['Department', 'Sum of Basic'], ['Projects', 5000], ['Finance', 4000]]],
          ['Lists', [['Departments'], ['Projects'], ['Finance']]],
          ['Data', [['ACME CONTRACTING'], ['Payroll for October 2026'], [], ['Emp Code', 'Employee', 'Designation', 'Basic Salary', 'Allowances', 'Deductions', 'Net'], ['E1', 'Ali Hassan', 'Mason', 3000, 500, 100, 3400], ['E2', 'Omar Khan', 'Driver', 2000, 300, 0, 2300]]]
        ]);
      },
      expect: 'only the payroll sheet is used',
      check: function (p, ok) {
        ok(p.sheetName === 'Data' && p.sheetMode !== 'combined', 'Data used alone');
        ok(p.slips.length === 2 && p.stats.withErrors === 0 && netMismatch(p) === 0, 'clean 2 slips');
      }
    }
  ];
}

function parse(XLSX, Schema, sc) {
  var wb = sc.csv != null ? Schema.readCsv(XLSX, sc.csv) : sc.wb();
  return Schema.parseWorkbook(XLSX, wb, {});
}

function run(XLSX, Schema) {
  return scenarios(XLSX, Schema).map(function (sc) {
    var fails = [];
    var p;
    try {
      p = parse(XLSX, Schema, sc);
      sc.check(p, function (cond, what) { if (!cond) fails.push(what); });
    } catch (e) {
      fails.push('crashed: ' + e.message);
    }
    return {
      id: sc.id, title: sc.title, expect: sc.expect, pass: !fails.length, fails: fails,
      got: p ? (p.slips.length + ' payslips' + (p.sheetName ? ' from ' + p.sheetName : '') +
        ((p.supplements || []).filter(function (s) { return s.enabled; }).map(function (s) { return ' + ' + s.sheet; }).join('')) +
        (errors(p).length ? ' · ' + errors(p).map(function (i) { return i.code || 'error'; }).join(', ') : '')) : ''
    };
  });
}

module.exports = { run: run, scenarios: scenarios };

if (require.main === module) {
  var XLSX = require(path.join(__dirname, '..', 'js', 'vendor', 'xlsx.full.min.js'));
  var Schema = require(path.join(__dirname, '..', 'js', 'schema.js'));
  var res = run(XLSX, Schema);
  res.forEach(function (r) {
    console.log((r.pass ? 'PASS ' : 'FAIL ') + String(r.id).padStart(2) + '  ' + r.title + '\n        got: ' + r.got + (r.fails.length ? '\n        missing: ' + r.fails.join(' | ') : ''));
  });
  console.log(res.filter(function (r) { return r.pass; }).length + ' / ' + res.length + ' pass');
}
