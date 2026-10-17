/*
 * Payslip Generator - the styled Excel template (ExcelJS).
 * Colour-banded column groups, frozen header, Net / Check formulas, totals, data validation with
 * input messages, header notes, text formats that keep leading zeros, and sheet protection without
 * a password. Works in the browser and in Node (tools/build-samples.js).
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory(require('./schema.js'));
  } else {
    root.PayslipTemplate = factory(root.PayslipSchema);
  }
})(typeof self !== 'undefined' ? self : this, function (Schema) {
  'use strict';

  var FIRST = 4;          // first data row (1: title + totals, 2: group band, 3: headers)
  var CAPACITY = 1000;    // rows pre-formatted with validation and formulas
  var LAST = FIRST + CAPACITY - 1;

  var GROUPS = [
    { name: 'EMPLOYEE', color: '1E3A8A', tint: 'DBEAFE', keys: ['employeeId', 'name', 'nationalId', 'jobTitle', 'department', 'joiningDate'] },
    { name: 'CONTACT', color: '334155', tint: 'E2E8F0', keys: ['email', 'mobile'] },
    { name: 'ATTENDANCE', color: '475569', tint: 'F1F5F9', keys: ['workingDays', 'absentDays', 'overtimeHours'] },
    { name: 'EARNINGS', color: '047857', tint: 'D1FAE5', keys: ['basicSalary', 'housing', 'transport', 'food', 'overtimePay', 'otherEarnings'] },
    { name: 'DEDUCTIONS', color: 'B91C1C', tint: 'FEE2E2', keys: ['gosi', 'absence', 'loan', 'otherDeductions'] },
    { name: 'RESULT', color: '0E7490', tint: 'CFFAFE', keys: ['netSalary', 'check'] },
    { name: 'OPTIONAL', color: '64748B', tint: 'F1F5F9', keys: ['nationality', 'dob', 'bankName', 'iban', 'remarks'], outline: true }
  ];

  var HELP = {
    employeeId: 'Unique code per employee, e.g. EMP-001. Required.',
    name: 'Full name as on the Iqama. English or Arabic. Required.',
    nationalId: '10 digits. Saudi IDs start with 1, Iqama numbers with 2. Also used as the PDF password.',
    jobTitle: 'Job title or profession.',
    department: 'Department, project or site.',
    joiningDate: 'Date the employee joined, e.g. 15-Mar-2024.',
    email: 'Work or personal email. Needed to send the payslip by Outlook.',
    mobile: 'Mobile number, e.g. 0551234567.',
    workingDays: 'Days worked in this month (0 to 31).',
    absentDays: 'Unpaid absent days (0 to 31, not more than working days).',
    overtimeHours: 'Overtime hours worked this month.',
    basicSalary: 'Monthly basic salary. Required.',
    housing: 'Housing allowance for this month.',
    transport: 'Transport allowance for this month.',
    food: 'Food allowance for this month.',
    overtimePay: 'Overtime amount for this month.',
    otherEarnings: 'Bonus or any other earning. For a named item add a column "Earning: Item name".',
    gosi: 'Employee share of GOSI.',
    absence: 'Deduction for absent days.',
    loan: 'Loan or salary advance repayment.',
    otherDeductions: 'Any other deduction. For a named item add a column "Deduction: Item name".',
    netSalary: 'Calculated: earnings minus deductions. Do not type here.',
    check: 'Shows ✓ when the row is fine, or a warning to fix before uploading.',
    nationality: 'Nationality (pick from the list or type).',
    dob: 'Date of birth. Only needed for the "id+dob" password rule.',
    bankName: 'Bank name.',
    iban: 'IBAN, 24 characters starting with SA.',
    remarks: 'Note printed on this employee\'s payslip.'
  };

  var WIDTH = {
    employeeId: 13, name: 28, nationalId: 15, jobTitle: 20, department: 16, joiningDate: 14, email: 28, mobile: 14,
    workingDays: 10, absentDays: 10, overtimeHours: 11, netSalary: 15, check: 26, nationality: 14, dob: 14, bankName: 18, iban: 28, remarks: 30
  };

  var NATIONALITIES = ['Saudi', 'Egyptian', 'Indian', 'Pakistani', 'Bangladeshi', 'Filipino', 'Nepali', 'Sri Lankan', 'Sudanese', 'Yemeni', 'Jordanian', 'Syrian', 'Lebanese', 'Palestinian', 'Indonesian', 'Kenyan', 'Ugandan', 'Ethiopian', 'Nigerian', 'Moroccan', 'Tunisian', 'Turkish', 'Afghan', 'Chinese', 'Vietnamese', 'Thai', 'British', 'American', 'Other'];
  var CURRENCIES = ['SAR', 'AED', 'QAR', 'KWD', 'BHD', 'OMR', 'USD', 'EUR'];
  var RULES = ['id', 'id+dob', 'dob', 'custom', 'none'];
  var MONTH_NAMES = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];

  function colLetter(n) {
    var s = '';
    while (n > 0) {
      var m = (n - 1) % 26;
      s = String.fromCharCode(65 + m) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  }

  function argb(hex) { return 'FF' + hex; }

  function fill(hex) { return { type: 'pattern', pattern: 'solid', fgColor: { argb: argb(hex) } }; }

  var THIN = { style: 'thin', color: { argb: 'FFCBD5E1' } };
  var BORDER = { top: THIN, left: THIN, bottom: THIN, right: THIN };

  function layoutColumns() {
    var cols = [];
    GROUPS.forEach(function (g) {
      g.keys.forEach(function (k) {
        var def = k === 'check' ? { key: 'check', label: 'Check' } : Schema.COLUMNS.filter(function (c) { return c.key === k; })[0];
        cols.push({ key: k, def: def, group: g, index: cols.length + 1 });
      });
    });
    cols.forEach(function (c) { c.letter = colLetter(c.index); });
    return cols;
  }

  function payPeriodOptions(today) {
    var out = [];
    for (var i = -3; i <= 12; i++) {
      var d = new Date(today.getFullYear(), today.getMonth() + i, 1);
      out.push(MONTH_NAMES[d.getMonth()] + ' ' + d.getFullYear());
    }
    return out;
  }

  function isoToDate(v) {
    if (v instanceof Date) return v;
    var m = String(v || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
    return m ? new Date(Date.UTC(+m[1], +m[2] - 1, +m[3])) : v;
  }

  function asUtcDate(d) {
    // ExcelJS writes dates in UTC; keep the calendar day the user sees
    return d instanceof Date ? new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate())) : d;
  }

  function imageFromDataUrl(dataUrl) {
    var m = /^data:image\/(png|jpe?g);base64,(.*)$/i.exec(dataUrl || '');
    if (!m) return null;
    return { base64: m[2], extension: m[1].toLowerCase() === 'png' ? 'png' : 'jpeg' };
  }

  function fitPx(img, maxW, maxH) {
    var w = img.width || maxW;
    var h = img.height || maxH;
    var k = Math.min(maxW / w, maxH / h, 1);
    return { width: Math.round(w * k), height: Math.round(h * k) };
  }

  /**
   * build(ExcelJS, opts) -> Promise<ArrayBuffer|Buffer>
   * opts.settings  company values to pre-fill (key -> value)
   * opts.records   employee rows (objects keyed by column key); none = one example row
   * opts.sample    fill Settings with the example values
   * opts.logo / opts.stamp  { dataUrl, width, height }
   */
  function build(ExcelJS, opts) {
    opts = opts || {};
    var today = opts.today || new Date();
    var wb = new ExcelJS.Workbook();
    wb.creator = 'Payslip Generator';
    wb.created = today;
    var cols = layoutColumns();
    var byKey = {};
    cols.forEach(function (c) { byKey[c.key] = c; });

    // ------------------------------------------------------------------ Lists (hidden)
    var periods = payPeriodOptions(today);
    var lists = wb.addWorksheet('Lists', { state: 'hidden' });
    lists.getRow(1).values = ['Currency', 'Nationality', 'Pay period', 'Password rule'];
    var maxLen = Math.max(CURRENCIES.length, NATIONALITIES.length, periods.length, RULES.length);
    for (var i = 0; i < maxLen; i++) {
      lists.getRow(i + 2).values = [CURRENCIES[i] || null, NATIONALITIES[i] || null, periods[i] || null, RULES[i] || null];
    }
    var listRef = {
      currency: 'Lists!$A$2:$A$' + (CURRENCIES.length + 1),
      nationality: 'Lists!$B$2:$B$' + (NATIONALITIES.length + 1),
      period: 'Lists!$C$2:$C$' + (periods.length + 1),
      rule: 'Lists!$D$2:$D$' + (RULES.length + 1)
    };

    // ------------------------------------------------------------------ Settings
    var st = wb.addWorksheet('Settings', { properties: { tabColor: { argb: argb('1E3A8A') } }, views: [{ state: 'frozen', ySplit: 2, showGridLines: false }] });
    st.columns = [{ width: 26 }, { width: 46 }, { width: 66 }, { width: 3 }, { width: 14 }, { width: 14 }, { width: 14 }];
    st.mergeCells('A1:C1');
    var t1 = st.getCell('A1');
    t1.value = 'COMPANY SETTINGS  ·  fill once, reused every month';
    t1.font = { bold: true, size: 14, color: { argb: 'FFFFFFFF' } };
    t1.fill = fill('1E3A8A');
    t1.alignment = { vertical: 'middle', indent: 1 };
    st.getRow(1).height = 30;
    ['Setting', 'Value', 'Notes'].forEach(function (h, k) {
      var c = st.getCell(2, k + 1);
      c.value = h;
      c.font = { bold: true, color: { argb: 'FF334155' } };
      c.fill = fill('E2E8F0');
      c.border = BORDER;
    });
    var preset = opts.settings || {};
    var settingRow = {};
    var r = 3;
    Schema.SETTINGS_FIELDS.forEach(function (f) {
      var label = st.getCell(r, 1);
      var val = st.getCell(r, 2);
      var note = st.getCell(r, 3);
      label.value = f.label + (f.required ? ' *' : '');
      label.font = { bold: true, color: { argb: 'FF0F172A' } };
      label.border = BORDER;
      note.value = f.note + (f.example && f.type !== 'image' ? '  Example: ' + f.example : '');
      note.font = { italic: true, size: 9, color: { argb: 'FF64748B' } };
      note.alignment = { wrapText: true, vertical: 'top' };
      var v;
      if (f.type === 'image') v = '(put the picture in the box on the right)';
      else if (preset[f.key] != null && preset[f.key] !== '') v = preset[f.key];
      else if (opts.sample && f.example) v = f.example;
      else v = f.default != null ? f.default : null;
      if (f.type === 'date' && v) v = asUtcDate(isoToDate(v));
      val.value = v;
      val.fill = fill(f.type === 'image' ? 'F8FAFC' : 'FFFBEB');
      val.border = BORDER;
      val.alignment = { vertical: 'top', wrapText: !!f.multiline };
      val.protection = { locked: f.type === 'image' };
      if (f.type === 'date') val.numFmt = 'dd-mmm-yyyy';
      if (f.key === 'crNumber' || f.key === 'countryCode') val.numFmt = '@';
      if (f.multiline) st.getRow(r).height = 96;
      settingRow[f.key] = r;
      r++;
    });
    var addV = function (ws, ref, rule) { ws.dataValidations.add(ref, rule); };
    var cell = function (key) { return 'B' + settingRow[key]; };
    addV(st, cell('currency'), { type: 'list', allowBlank: true, formulae: [listRef.currency], showErrorMessage: true, errorStyle: 'stop', errorTitle: 'Currency', error: 'Pick a currency from the list.', showInputMessage: true, promptTitle: 'Currency', prompt: 'Pick from the list. SAR for Saudi Arabia.' });
    addV(st, cell('passwordRule'), { type: 'list', allowBlank: true, formulae: [listRef.rule], showErrorMessage: true, errorStyle: 'stop', errorTitle: 'Password rule', error: 'Pick a rule from the list.', showInputMessage: true, promptTitle: 'PDF password', prompt: 'id = Iqama number (default). none = no password.' });
    addV(st, cell('payPeriod'), { type: 'list', allowBlank: false, formulae: [listRef.period], showErrorMessage: true, errorStyle: 'warning', errorTitle: 'Pay period', error: 'Not in the list. Keep it anyway?', showInputMessage: true, promptTitle: 'Pay period', prompt: 'The month printed on every payslip, e.g. October 2026.' });
    ['periodStart', 'periodEnd', 'paymentDate'].forEach(function (k) {
      addV(st, cell(k), { type: 'date', operator: 'greaterThan', allowBlank: true, formulae: [new Date(Date.UTC(2000, 0, 1))], showErrorMessage: true, errorStyle: 'stop', errorTitle: 'Date', error: 'Enter a date, e.g. 01-Oct-2026.', showInputMessage: true, promptTitle: 'Date', prompt: 'Enter a date, e.g. 01-Oct-2026.' });
    });
    addV(st, cell('crNumber'), { type: 'custom', allowBlank: true, formulae: ['AND(LEN(' + cell('crNumber') + ')=10,ISNUMBER(--' + cell('crNumber') + '))'], showErrorMessage: true, errorStyle: 'warning', errorTitle: 'CR number', error: 'A CR number has 10 digits. Keep it anyway?', showInputMessage: true, promptTitle: 'CR number', prompt: '10-digit Commercial Registration number.' });
    addV(st, cell('senderEmail'), { type: 'custom', allowBlank: true, formulae: ['ISNUMBER(FIND("@",' + cell('senderEmail') + '))'], showErrorMessage: true, errorStyle: 'warning', errorTitle: 'Email', error: 'This does not look like an email address.' });
    addV(st, cell('companyName'), { type: 'textLength', operator: 'greaterThanOrEqual', allowBlank: false, formulae: [2], showErrorMessage: true, errorStyle: 'stop', errorTitle: 'Company name', error: 'Enter the company name.', showInputMessage: true, promptTitle: 'Company name', prompt: 'Legal name as on the CR.' });

    // logo and stamp boxes
    function box(range, title) {
      st.mergeCells(range);
      var c = st.getCell(range.split(':')[0]);
      c.value = title;
      c.font = { italic: true, size: 9, color: { argb: 'FF94A3B8' } };
      c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      var dash = { style: 'dashed', color: { argb: 'FF94A3B8' } };
      c.border = { top: dash, left: dash, bottom: dash, right: dash };
    }
    box('E3:G9', 'LOGO\nInsert your company logo here\n(Insert ▸ Pictures)');
    box('E11:G16', 'STAMP / SIGNATURE (optional)\nsecond picture on this sheet');
    var logo = opts.logo && imageFromDataUrl(opts.logo.dataUrl);
    if (logo) {
      var lid = wb.addImage(logo);
      st.addImage(lid, { tl: { col: 4.15, row: 2.2 }, ext: fitPx(opts.logo, 200, 110) });
    }
    var stamp = opts.stamp && imageFromDataUrl(opts.stamp.dataUrl);
    if (stamp) {
      var sid = wb.addImage(stamp);
      st.addImage(sid, { tl: { col: 4.15, row: 10.2 }, ext: fitPx(opts.stamp, 180, 90) });
    }

    // ------------------------------------------------------------------ Payslips
    var ws = wb.addWorksheet('Payslips', {
      properties: { tabColor: { argb: argb('047857') }, outlineLevelCol: 1 },
      views: [{ state: 'frozen', xSplit: 2, ySplit: 3, zoomScale: 100 }]
    });
    var lastCol = cols[cols.length - 1].letter;
    ws.columns = cols.map(function (c) {
      var t = c.def.type;
      var style = { protection: { locked: c.key === 'netSalary' || c.key === 'check' } };
      if (t === 'money') style.numFmt = '#,##0.00';
      else if (t === 'number') style.numFmt = '0.##';
      else if (t === 'date') style.numFmt = 'dd-mmm-yyyy';
      else if (c.key === 'employeeId' || c.key === 'nationalId' || c.key === 'mobile' || c.key === 'iban') style.numFmt = '@';
      return { key: c.key, width: WIDTH[c.key] || 13, style: style, outlineLevel: c.group.outline ? 1 : 0 };
    });

    // row 1: title + live totals
    ws.mergeCells('A1:F1');
    var title = ws.getCell('A1');
    title.value = 'PAYSLIPS  ·  one row per employee  ·  amounts for this month';
    title.font = { bold: true, size: 14, color: { argb: 'FFFFFFFF' } };
    title.alignment = { vertical: 'middle', indent: 1 };
    for (var c1 = 1; c1 <= cols.length; c1++) ws.getCell(1, c1).fill = fill('0F172A');
    ws.getRow(1).height = 30;
    var net = byKey.netSalary.letter;
    var nameL = byKey.name.letter;
    var eFirst = byKey.basicSalary.letter;
    var eLast = byKey.otherEarnings.letter;
    var dFirst = byKey.gosi.letter;
    var dLast = byKey.otherDeductions.letter;
    var put1 = function (letter, value, fmt) {
      var c = ws.getCell(letter + '1');
      c.value = value;
      c.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      c.alignment = { horizontal: 'right', vertical: 'middle' };
      if (fmt) c.numFmt = fmt;
    };
    var recs = opts.records && opts.records.length ? opts.records : null;
    put1(byKey.basicSalary.letter, 'EMPLOYEES');
    put1(byKey.housing.letter, { formula: 'COUNTA(' + nameL + FIRST + ':' + nameL + LAST + ')-COUNTIF(' + nameL + FIRST + ':' + nameL + LAST + ',"Example*")', result: recs ? recs.length : 0 }, '0');
    put1(byKey.otherDeductions.letter, 'TOTAL NET');
    var totalNet = 0;

    // row 2: group bands
    GROUPS.forEach(function (g) {
      var a = byKey[g.keys[0]].letter;
      var b = byKey[g.keys[g.keys.length - 1]].letter;
      if (a !== b) ws.mergeCells(a + '2:' + b + '2');
      var c = ws.getCell(a + '2');
      c.value = g.name + (g.outline ? '  (click − above to fold)' : '');
      c.font = { bold: true, size: 9, color: { argb: 'FFFFFFFF' } };
      c.fill = fill(g.color);
      c.alignment = { horizontal: 'center', vertical: 'middle' };
      c.numFmt = 'General';
    });
    ws.getRow(2).height = 18;

    // row 3: headers with notes
    cols.forEach(function (c) {
      var h = ws.getCell(3, c.index);
      h.value = c.def.label;
      h.font = { bold: true, color: { argb: argb(c.group.color) } };
      h.fill = fill(c.group.tint);
      h.border = { top: THIN, left: THIN, right: THIN, bottom: { style: 'medium', color: { argb: argb(c.group.color) } } };
      h.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
      h.protection = { locked: true };
      h.numFmt = 'General';
      if (HELP[c.key]) h.note = { texts: [{ text: HELP[c.key] }], margins: { insetmode: 'auto' } };
    });
    ws.getRow(3).height = 34;
    ws.autoFilter = { from: 'A3', to: lastCol + '3' };

    // data rows
    var rows = recs || [(function () {
      var ex = {};
      cols.forEach(function (c) {
        if (c.key === 'employeeId') ex[c.key] = 'EXAMPLE-001';
        else if (c.key === 'name') ex[c.key] = 'Example Row - delete me';
        else if (c.def.example !== undefined && c.def.example !== '') ex[c.key] = c.def.type === 'date' ? isoToDate(c.def.example) : c.def.example;
      });
      return ex;
    })()];
    rows.forEach(function (rec, k) {
      var rowNo = FIRST + k;
      var row = ws.getRow(rowNo);
      var earn = 0;
      var ded = 0;
      cols.forEach(function (c) {
        if (c.key === 'netSalary' || c.key === 'check') return;
        var v = rec[c.key];
        if (v === undefined || v === null || v === '') return;
        if (c.def.type === 'date') v = asUtcDate(v instanceof Date ? v : isoToDate(v));
        else if (c.def.type !== 'money' && c.def.type !== 'number') v = String(v);
        row.getCell(c.index).value = v;
        if (c.def.group === 'earning') earn += +v || 0;
        if (c.def.group === 'deduction') ded += +v || 0;
      });
      rec.__net = Math.round((earn - ded) * 100) / 100;
      rec.__check = !rec.basicSalary && rec.basicSalary !== 0 ? '⚠ Basic salary missing' : ded > earn ? '⚠ Deductions exceed earnings' : '✓';
      if (recs) totalNet += rec.__net;
    });
    put1(byKey.netSalary.letter, { formula: 'SUM(' + net + FIRST + ':' + net + LAST + ')', result: Math.round(totalNet * 100) / 100 }, '#,##0.00');
    for (var rr = FIRST; rr <= LAST; rr++) {
      var rec = rows[rr - FIRST];
      var nCell = ws.getCell(net + rr);
      var nf = 'IF($' + nameL + rr + '="","",SUM(' + eFirst + rr + ':' + eLast + rr + ')-SUM(' + dFirst + rr + ':' + dLast + rr + '))';
      nCell.value = rec ? { formula: nf, result: rec.__net } : { formula: nf };
      nCell.fill = fill('F1F5F9');
      nCell.font = { bold: true, color: { argb: 'FF0E7490' } };
      var cCell = ws.getCell(byKey.check.letter + rr);
      var basic = byKey.basicSalary.letter + rr;
      var idL = byKey.employeeId.letter;
      var cf = 'IF($' + nameL + rr + '="","",IF(' + basic + '="","⚠ Basic salary missing",IF(SUM(' + dFirst + rr + ':' + dLast + rr + ')>SUM(' + eFirst + rr + ':' + eLast + rr + '),"⚠ Deductions exceed earnings",IF(COUNTIF($' + idL + '$' + FIRST + ':$' + idL + '$' + LAST + ',$' + idL + rr + ')>1,"⚠ Duplicate ID","✓"))))';
      cCell.value = rec ? { formula: cf, result: rec.__check } : { formula: cf };
      cCell.fill = fill('F1F5F9');
    }
    ws.addConditionalFormatting({
      ref: byKey.check.letter + FIRST + ':' + byKey.check.letter + LAST,
      rules: [
        { type: 'containsText', operator: 'containsText', text: '⚠', style: { font: { color: { argb: 'FFB91C1C' }, bold: true }, fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFEE2E2' } } } },
        { type: 'containsText', operator: 'containsText', text: '✓', style: { font: { color: { argb: 'FF047857' }, bold: true } } }
      ]
    });

    // validation (rows FIRST..LAST)
    function rng(key) { var l = byKey[key].letter; return l + FIRST + ':' + l + LAST; }
    function at(key) { return byKey[key].letter + FIRST; }
    function prompt(key, title) { return { showInputMessage: true, promptTitle: title || byKey[key].def.label, prompt: HELP[key] || '' }; }
    function rule(key, obj, style, err) {
      var full = Object.assign({ allowBlank: true, showErrorMessage: true, errorStyle: style || 'stop', errorTitle: byKey[key].def.label, error: err }, prompt(key), obj);
      ws.dataValidations.add(rng(key), full);
    }
    var idCol = byKey.employeeId.letter;
    rule('employeeId', { type: 'custom', formulae: ['COUNTIF($' + idCol + '$' + FIRST + ':$' + idCol + '$' + LAST + ',' + at('employeeId') + ')=1'] }, 'stop', 'This Employee ID is already used in another row.');
    rule('name', { type: 'textLength', operator: 'greaterThanOrEqual', formulae: [2] }, 'stop', 'Enter the employee name.');
    rule('nationalId', { type: 'custom', formulae: ['AND(LEN(' + at('nationalId') + ')=10,ISNUMBER(--' + at('nationalId') + '),OR(LEFT(' + at('nationalId') + ',1)="1",LEFT(' + at('nationalId') + ',1)="2"))'] }, 'warning', 'Expected 10 digits starting with 1 or 2. Keep it anyway?');
    var minJoin = new Date(Date.UTC(1980, 0, 1));
    var maxJoin = new Date(Date.UTC(today.getFullYear() + 1, today.getMonth(), today.getDate()));
    rule('joiningDate', { type: 'date', operator: 'between', formulae: [minJoin, maxJoin] }, 'stop', 'Enter a date between 1980 and next year, e.g. 15-Mar-2024.');
    rule('email', { type: 'custom', formulae: ['AND(ISNUMBER(FIND("@",' + at('email') + ')),ISNUMBER(FIND(".",' + at('email') + ',FIND("@",' + at('email') + '))))'] }, 'warning', 'This does not look like an email address. Keep it anyway?');
    rule('mobile', { type: 'custom', formulae: ['OR(AND(LEN(' + at('mobile') + ')=10,LEFT(' + at('mobile') + ',2)="05"),LEFT(' + at('mobile') + ',1)="+",LEFT(' + at('mobile') + ',3)="966")'] }, 'warning', 'Saudi mobile numbers look like 0551234567. Keep it anyway?');
    rule('workingDays', { type: 'whole', operator: 'between', formulae: [0, 31] }, 'stop', 'Working days must be a whole number from 0 to 31.');
    var wd = byKey.workingDays.letter + FIRST;
    rule('absentDays', { type: 'custom', formulae: ['AND(ISNUMBER(' + at('absentDays') + '),' + at('absentDays') + '>=0,' + at('absentDays') + '<=31,' + at('absentDays') + '=INT(' + at('absentDays') + '),OR(' + wd + '="",' + at('absentDays') + '<=' + wd + '))'] }, 'stop', 'Absent days: a whole number from 0 to 31, not more than working days.');
    rule('overtimeHours', { type: 'decimal', operator: 'between', formulae: [0, 300] }, 'stop', 'Overtime hours must be between 0 and 300.');
    ['basicSalary', 'housing', 'transport', 'food', 'overtimePay', 'otherEarnings', 'gosi', 'absence', 'loan', 'otherDeductions'].forEach(function (k) {
      rule(k, { type: 'decimal', operator: 'between', formulae: [0, 1000000] }, 'stop', 'Enter an amount from 0 to 1,000,000 (no negative numbers; deductions are positive).');
    });
    rule('nationality', { type: 'list', formulae: [listRef.nationality] }, 'warning', 'Not in the list. Keep it anyway?');
    var minDob = new Date(Date.UTC(1940, 0, 1));
    var maxDob = new Date(Date.UTC(today.getFullYear() - 16, today.getMonth(), today.getDate()));
    rule('dob', { type: 'date', operator: 'between', formulae: [minDob, maxDob] }, 'stop', 'Enter a date of birth between 1940 and 16 years ago.');
    rule('iban', { type: 'custom', formulae: ['AND(LEN(SUBSTITUTE(' + at('iban') + '," ",""))=24,LEFT(' + at('iban') + ',2)="SA")'] }, 'warning', 'A Saudi IBAN has 24 characters starting with SA. Keep it anyway?');

    // ------------------------------------------------------------------ Instructions
    var ins = wb.addWorksheet('Instructions', { properties: { tabColor: { argb: argb('64748B') } }, views: [{ showGridLines: false }] });
    ins.columns = [{ width: 4 }, { width: 26 }, { width: 12 }, { width: 70 }, { width: 22 }];
    ins.mergeCells('B1:E1');
    var it = ins.getCell('B1');
    it.value = 'HOW TO USE THIS TEMPLATE';
    it.font = { bold: true, size: 14, color: { argb: 'FF0F172A' } };
    ins.getRow(1).height = 28;
    var steps = [
      'Fill the Settings sheet once: company name, CR number, pay period. Put your logo picture in the box.',
      'Payslips sheet: one row per employee. Delete the example row. Net Salary and Check fill themselves.',
      'Fix every row where Check shows ⚠. Hover a column title to see what it needs.',
      'Extra items: add a column named "Earning: Site Allowance" or "Deduction: Housing Loan".',
      'Save, then drop the file on the Payslip Generator page. Next month: change the amounts and the pay period.'
    ];
    steps.forEach(function (s, k) {
      var n = ins.getCell(3 + k, 2);
      n.value = (k + 1) + '.';
      n.font = { bold: true, color: { argb: 'FF2563EB' } };
      n.alignment = { horizontal: 'right' };
      ins.mergeCells('C' + (3 + k) + ':E' + (3 + k));
      ins.getCell(3 + k, 3).value = s;
    });
    var hr = 10;
    ['Column', 'Required', 'What to enter', 'Example'].forEach(function (h, k) {
      var c = ins.getCell(hr, k + 2);
      c.value = h;
      c.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      c.fill = fill('0F172A');
    });
    cols.forEach(function (c, k) {
      var row = hr + 1 + k;
      ins.getCell(row, 2).value = c.def.label;
      ins.getCell(row, 2).font = { bold: true, color: { argb: argb(c.group.color) } };
      ins.getCell(row, 3).value = c.def.required ? 'Yes' : c.key === 'netSalary' || c.key === 'check' ? 'Formula' : 'No';
      ins.getCell(row, 4).value = HELP[c.key] || '';
      ins.getCell(row, 4).alignment = { wrapText: true };
      var ex = c.def.example;
      ins.getCell(row, 5).value = ex === undefined || ex === '' ? '' : String(ex);
      for (var q = 2; q <= 5; q++) ins.getCell(row, q).border = { bottom: THIN };
    });

    // order the tabs: Settings, Payslips, Instructions (Lists stays hidden)
    wb.worksheets.forEach(function (w) {
      w.orderNo = { Settings: 0, Payslips: 1, Instructions: 2, Lists: 3 }[w.name];
    });
    wb.views = [{ activeTab: 1, firstSheet: 0 }];

    return Promise.all([
      st.protect('', { selectLockedCells: true, selectUnlockedCells: true, formatColumns: true, formatRows: true }),
      ws.protect('', { selectLockedCells: true, selectUnlockedCells: true, formatCells: true, formatColumns: true, formatRows: true, insertColumns: true, insertRows: true, deleteRows: true, sort: true, autoFilter: true })
    ]).then(function () {
      return wb.xlsx.writeBuffer();
    });
  }

  return {
    build: build,
    layoutColumns: layoutColumns,
    FIRST_ROW: FIRST,
    CAPACITY: CAPACITY
  };
});
