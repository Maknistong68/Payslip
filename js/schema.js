/*
 * Payslip Generator - Excel schema, template builder, parser and validator.
 * Pure functions only (no DOM) so the same code runs in the browser and in Node tests.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory();
  } else {
    root.PayslipSchema = factory();
  }
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  // ---------------------------------------------------------------------------
  // Schema definitions
  // ---------------------------------------------------------------------------

  var SHEET_SETTINGS = 'Settings';
  var SHEET_PAYSLIPS = 'Payslips';
  var SHEET_INSTRUCTIONS = 'Instructions';

  var SETTINGS_FIELDS = [
    { key: 'companyName', label: 'Company Name', required: true, example: 'Example Contracting Co. Ltd.', note: 'Required. Legal name as it appears on the CR.', aliases: ['company', 'employer', 'employer name', 'establishment name'] },
    { key: 'crNumber', label: 'CR Number', required: true, example: '1010123456', note: 'Required. Commercial Registration number.', aliases: ['cr', 'cr no', 'commercial registration', 'commercial registration number', 'cr number'] },
    { key: 'address', label: 'Company Address', example: 'Tabuk, Kingdom of Saudi Arabia', note: 'Optional.', aliases: ['address'] },
    { key: 'phone', label: 'Company Phone', example: '+966 11 000 0000', note: 'Optional.', aliases: ['phone', 'telephone', 'mobile'] },
    { key: 'email', label: 'Company Email', example: 'hr@example.com', note: 'Optional.', aliases: ['email', 'e-mail'] },
    { key: 'currency', label: 'Currency', example: 'SAR', default: 'SAR', note: 'Optional. 3-letter code, defaults to SAR.', aliases: ['currency code'] },
    { key: 'payPeriod', label: 'Pay Period', required: true, example: 'October 2026', note: 'Required. The month shown on every payslip, e.g. "October 2026".', aliases: ['period', 'salary month', 'month', 'pay month'] },
    { key: 'periodStart', label: 'Period Start Date', example: '2026-10-01', type: 'date', note: 'Optional. First day of the pay period.', aliases: ['start date', 'from', 'period from'] },
    { key: 'periodEnd', label: 'Period End Date', example: '2026-10-31', type: 'date', note: 'Optional. Last day of the pay period.', aliases: ['end date', 'to', 'period to'] },
    { key: 'paymentDate', label: 'Payment Date', example: '2026-11-05', type: 'date', note: 'Optional. Date salaries were transferred.', aliases: ['pay date', 'date of payment', 'transfer date'] },
    { key: 'preparedBy', label: 'Prepared By', example: 'HR Department', note: 'Optional. Printed in the footer.', aliases: ['prepared', 'issued by'] },
    { key: 'footerNote', label: 'Footer Note', example: 'This is a system-generated payslip and does not require a signature.', note: 'Optional. Free text printed at the bottom of each payslip.', aliases: ['note', 'footer', 'disclaimer'] }
  ];

  var COLUMNS = [
    { key: 'employeeId', label: 'Employee ID', group: 'info', required: true, example: 'EMP-001', aliases: ['emp id', 'employee no', 'employee number', 'emp no', 'id', 'staff id', 'staff no'] },
    { key: 'name', label: 'Employee Name', group: 'info', required: true, example: 'Ahmed Ali Hassan', aliases: ['name', 'full name', 'employee', 'staff name'] },
    { key: 'nationalId', label: 'National ID / Iqama', group: 'info', example: '2456789012', aliases: ['iqama', 'iqama no', 'iqama number', 'national id', 'id number', 'national id iqama', 'iqama id'] },
    { key: 'nationality', label: 'Nationality', group: 'info', example: 'Egyptian', aliases: [] },
    { key: 'jobTitle', label: 'Job Title', group: 'info', example: 'Site Engineer', aliases: ['title', 'position', 'designation', 'job', 'profession'] },
    { key: 'department', label: 'Department', group: 'info', example: 'Projects', aliases: ['dept', 'section', 'project'] },
    { key: 'joiningDate', label: 'Joining Date', group: 'info', type: 'date', example: '2024-03-15', aliases: ['date of joining', 'doj', 'hire date', 'start date', 'join date'] },
    { key: 'bankName', label: 'Bank Name', group: 'info', example: 'Al Rajhi Bank', aliases: ['bank'] },
    { key: 'iban', label: 'IBAN', group: 'info', example: 'SA0380000000608010167519', aliases: ['iban no', 'iban number', 'account', 'account number', 'bank account'] },
    { key: 'workingDays', label: 'Working Days', group: 'attendance', type: 'number', example: 30, aliases: ['days worked', 'work days', 'paid days'] },
    { key: 'absentDays', label: 'Absent Days', group: 'attendance', type: 'number', example: 0, aliases: ['absence days', 'absent', 'unpaid days', 'lop days'] },
    { key: 'overtimeHours', label: 'Overtime Hours', group: 'attendance', type: 'number', example: 10, aliases: ['ot hours', 'overtime hrs', 'ot hrs'] },
    { key: 'basicSalary', label: 'Basic Salary', group: 'earning', required: true, type: 'money', example: 4000, aliases: ['basic', 'basic pay', 'base salary', 'basic wage'] },
    { key: 'housing', label: 'Housing Allowance', group: 'earning', type: 'money', example: 1000, aliases: ['housing', 'hra', 'house allowance', 'accommodation allowance', 'accommodation'] },
    { key: 'transport', label: 'Transport Allowance', group: 'earning', type: 'money', example: 400, aliases: ['transport', 'transportation', 'transportation allowance', 'conveyance'] },
    { key: 'food', label: 'Food Allowance', group: 'earning', type: 'money', example: 300, aliases: ['food', 'meal allowance', 'meals'] },
    { key: 'overtimePay', label: 'Overtime Pay', group: 'earning', type: 'money', example: 250, aliases: ['overtime', 'ot', 'ot pay', 'overtime amount'] },
    { key: 'otherEarnings', label: 'Other Earnings', group: 'earning', type: 'money', example: 0, aliases: ['other allowance', 'other allowances', 'bonus', 'other earning', 'additions'] },
    { key: 'gosi', label: 'GOSI Deduction', group: 'deduction', type: 'money', example: 0, aliases: ['gosi', 'social insurance', 'gosi employee'] },
    { key: 'absence', label: 'Absence Deduction', group: 'deduction', type: 'money', example: 0, aliases: ['absence', 'absent deduction', 'lop', 'unpaid leave'] },
    { key: 'loan', label: 'Loan / Advance', group: 'deduction', type: 'money', example: 0, aliases: ['loan', 'advance', 'salary advance', 'loan deduction', 'loan advance'] },
    { key: 'otherDeductions', label: 'Other Deductions', group: 'deduction', type: 'money', example: 0, aliases: ['other deduction', 'deductions', 'penalty', 'fines'] },
    { key: 'netSalary', label: 'Net Salary', group: 'check', type: 'money', example: '', aliases: ['net', 'net pay', 'net amount', 'total net', 'net paid', 'take home'] },
    { key: 'remarks', label: 'Remarks', group: 'info', example: '', aliases: ['notes', 'comment', 'comments', 'remark'] }
  ];

  // Column headers like "Earning: Site Allowance" / "Deduction: Housing Loan" are
  // treated as extra earning / deduction lines without changing the app.
  var CUSTOM_EARNING_RE = /^(earning|earnings|allowance|addition)\s*[:\-]\s*(.+)$/i;
  var CUSTOM_DEDUCTION_RE = /^(deduction|deductions|ded)\s*[:\-]\s*(.+)$/i;

  var EXAMPLE_ID_PREFIX = 'EXAMPLE';

  // ---------------------------------------------------------------------------
  // Small helpers
  // ---------------------------------------------------------------------------

  function normalizeHeader(s) {
    return String(s == null ? '' : s)
      .toLowerCase()
      .replace(/\(.*?\)/g, ' ') // drop "(SAR)" style suffixes
      .replace(/[*]/g, ' ')
      .replace(/[^a-z0-9]+/g, ' ')
      .trim();
  }

  function buildLookup(defs) {
    var map = {};
    defs.forEach(function (d) {
      map[normalizeHeader(d.label)] = d;
      (d.aliases || []).forEach(function (a) {
        var n = normalizeHeader(a);
        if (!map[n]) map[n] = d;
      });
    });
    return map;
  }

  var SETTINGS_LOOKUP = buildLookup(SETTINGS_FIELDS);
  var COLUMN_LOOKUP = buildLookup(COLUMNS);

  function isBlank(v) {
    return v == null || (typeof v === 'string' && v.trim() === '');
  }

  function cleanText(v) {
    if (v == null) return '';
    if (v instanceof Date) return formatDate(v);
    return String(v).replace(/\s+/g, ' ').trim();
  }

  // Returns { value: Number|null, error: String|null }
  function parseNumber(v) {
    if (isBlank(v)) return { value: null, error: null };
    if (typeof v === 'number') {
      if (!isFinite(v)) return { value: null, error: 'is not a valid number' };
      return { value: v, error: null };
    }
    if (typeof v === 'boolean') return { value: null, error: 'is not a number' };
    var s = String(v).trim();
    // strip currency codes/symbols, thousands separators and whitespace
    var cleaned = s.replace(/[^0-9.\-()]/g, '');
    var negative = /^\(.*\)$/.test(cleaned);
    cleaned = cleaned.replace(/[()]/g, '');
    if (cleaned === '' || cleaned === '-' || cleaned === '.') return { value: null, error: 'is not a number ("' + s + '")' };
    var n = Number(cleaned);
    if (!isFinite(n)) return { value: null, error: 'is not a number ("' + s + '")' };
    return { value: negative ? -n : n, error: null };
  }

  function pad2(n) {
    return (n < 10 ? '0' : '') + n;
  }

  var MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

  function formatDate(d) {
    if (!(d instanceof Date) || isNaN(d.getTime())) return '';
    return pad2(d.getDate()) + ' ' + MONTHS[d.getMonth()] + ' ' + d.getFullYear();
  }

  // Excel serial date -> JS Date (1900 date system)
  function excelSerialToDate(serial) {
    if (typeof serial !== 'number' || !isFinite(serial)) return null;
    var utcDays = Math.floor(serial - 25569);
    var ms = utcDays * 86400 * 1000;
    var d = new Date(ms);
    return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
  }

  // Accepts Date, Excel serial, "YYYY-MM-DD", "DD/MM/YYYY", "DD-MM-YYYY" and free text.
  // Returns display string (free text is passed through unchanged).
  function parseDateLike(v) {
    if (isBlank(v)) return '';
    if (v instanceof Date) return formatDate(v);
    if (typeof v === 'number') {
      if (v > 20000 && v < 80000) return formatDate(excelSerialToDate(v));
      return String(v);
    }
    var s = String(v).trim();
    var m = s.match(/^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})/);
    if (m) return formatDate(new Date(+m[1], +m[2] - 1, +m[3]));
    m = s.match(/^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{4})/);
    if (m) return formatDate(new Date(+m[3], +m[2] - 1, +m[1]));
    return s;
  }

  function round2(n) {
    return Math.round((n + Number.EPSILON) * 100) / 100;
  }

  function formatMoney(n) {
    if (n == null || !isFinite(n)) return '';
    var neg = n < 0;
    var fixed = Math.abs(n).toFixed(2);
    var parts = fixed.split('.');
    parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    return (neg ? '-' : '') + parts.join('.');
  }

  // Characters outside Latin-1 / Latin Extended cannot be rendered by the standard PDF fonts.
  var NON_LATIN_RE = /[^\u0000-ɏ -⁯₠-⃏]/;

  function hasNonLatin(s) {
    return NON_LATIN_RE.test(String(s || ''));
  }

  // ---------------------------------------------------------------------------
  // Number to words (English)
  // ---------------------------------------------------------------------------

  var ONES = ['', 'One', 'Two', 'Three', 'Four', 'Five', 'Six', 'Seven', 'Eight', 'Nine', 'Ten', 'Eleven', 'Twelve', 'Thirteen', 'Fourteen', 'Fifteen', 'Sixteen', 'Seventeen', 'Eighteen', 'Nineteen'];
  var TENS = ['', '', 'Twenty', 'Thirty', 'Forty', 'Fifty', 'Sixty', 'Seventy', 'Eighty', 'Ninety'];
  var SCALES = ['', 'Thousand', 'Million', 'Billion', 'Trillion'];

  function chunkToWords(n) {
    var out = [];
    if (n >= 100) {
      out.push(ONES[Math.floor(n / 100)] + ' Hundred');
      n = n % 100;
    }
    if (n >= 20) {
      out.push(TENS[Math.floor(n / 10)] + (n % 10 ? '-' + ONES[n % 10] : ''));
    } else if (n > 0) {
      out.push(ONES[n]);
    }
    return out.join(' ');
  }

  function integerToWords(n) {
    n = Math.floor(Math.abs(n));
    if (n === 0) return 'Zero';
    var parts = [];
    var scale = 0;
    while (n > 0 && scale < SCALES.length) {
      var chunk = n % 1000;
      if (chunk) parts.unshift(chunkToWords(chunk) + (SCALES[scale] ? ' ' + SCALES[scale] : ''));
      n = Math.floor(n / 1000);
      scale++;
    }
    return parts.join(' ');
  }

  var CURRENCY_NAMES = {
    SAR: ['Saudi Riyal', 'Saudi Riyals', 'Halala', 'Halalas'],
    AED: ['Dirham', 'Dirhams', 'Fils', 'Fils'],
    QAR: ['Qatari Riyal', 'Qatari Riyals', 'Dirham', 'Dirhams'],
    KWD: ['Kuwaiti Dinar', 'Kuwaiti Dinars', 'Fils', 'Fils'],
    BHD: ['Bahraini Dinar', 'Bahraini Dinars', 'Fils', 'Fils'],
    OMR: ['Omani Rial', 'Omani Rials', 'Baisa', 'Baisa'],
    JOD: ['Jordanian Dinar', 'Jordanian Dinars', 'Piaster', 'Piasters'],
    EGP: ['Egyptian Pound', 'Egyptian Pounds', 'Piaster', 'Piasters'],
    USD: ['US Dollar', 'US Dollars', 'Cent', 'Cents'],
    EUR: ['Euro', 'Euros', 'Cent', 'Cents'],
    GBP: ['Pound Sterling', 'Pounds Sterling', 'Penny', 'Pence'],
    INR: ['Indian Rupee', 'Indian Rupees', 'Paisa', 'Paise'],
    PKR: ['Pakistani Rupee', 'Pakistani Rupees', 'Paisa', 'Paise'],
    BDT: ['Taka', 'Taka', 'Poisha', 'Poisha'],
    PHP: ['Philippine Peso', 'Philippine Pesos', 'Centavo', 'Centavos'],
    LKR: ['Sri Lankan Rupee', 'Sri Lankan Rupees', 'Cent', 'Cents'],
    NPR: ['Nepalese Rupee', 'Nepalese Rupees', 'Paisa', 'Paise']
  };

  function amountToWords(amount, currency) {
    if (amount == null || !isFinite(amount)) return '';
    var code = String(currency || 'SAR').toUpperCase();
    var names = CURRENCY_NAMES[code] || [code, code, 'Cent', 'Cents'];
    var neg = amount < 0;
    var abs = Math.abs(round2(amount));
    var whole = Math.floor(abs);
    var frac = Math.round((abs - whole) * 100);
    if (frac === 100) {
      whole += 1;
      frac = 0;
    }
    var s = integerToWords(whole) + ' ' + (whole === 1 ? names[0] : names[1]);
    if (frac > 0) s += ' and ' + integerToWords(frac) + ' ' + (frac === 1 ? names[2] : names[3]);
    s += ' Only';
    return (neg ? 'Minus ' : '') + s;
  }

  // ---------------------------------------------------------------------------
  // Template workbook
  // ---------------------------------------------------------------------------

  function buildTemplateWorkbook(XLSX, options) {
    options = options || {};
    var withSample = !!options.sampleRows;
    var wb = XLSX.utils.book_new();

    // Settings sheet
    var settingsAoa = [['Setting', 'Value', 'Notes']];
    SETTINGS_FIELDS.forEach(function (f) {
      var value = withSample ? f.example : f.default != null ? f.default : '';
      settingsAoa.push([f.label, value, f.note + (f.example ? ' Example: ' + f.example : '')]);
    });
    var wsSettings = XLSX.utils.aoa_to_sheet(settingsAoa);
    wsSettings['!cols'] = [{ wch: 22 }, { wch: 40 }, { wch: 80 }];
    XLSX.utils.book_append_sheet(wb, wsSettings, SHEET_SETTINGS);

    // Payslips sheet
    var header = COLUMNS.map(function (c) {
      return c.label;
    });
    var rows = [header];
    if (withSample) {
      rows = rows.concat(options.sampleRows);
    } else {
      // One example row that the parser recognises and ignores.
      rows.push(COLUMNS.map(function (c) {
        if (c.key === 'employeeId') return EXAMPLE_ID_PREFIX + '-001';
        if (c.key === 'name') return 'Example Row - delete me';
        if (c.key === 'netSalary') return '';
        return c.example;
      }));
    }
    var wsPay = XLSX.utils.aoa_to_sheet(rows);
    wsPay['!cols'] = COLUMNS.map(function (c) {
      return { wch: Math.max(12, c.label.length + 2) };
    });
    wsPay['!freeze'] = { xSplit: 2, ySplit: 1 };
    XLSX.utils.book_append_sheet(wb, wsPay, SHEET_PAYSLIPS);

    // Instructions sheet
    var instr = [
      ['HOW TO USE THIS TEMPLATE'],
      [''],
      ['1. Fill the "Settings" sheet with your company details. Company Name, CR Number and Pay Period are required.'],
      ['2. Fill the "Payslips" sheet with one row per employee. Delete the example row.'],
      ['3. Keep the header row exactly as it is. You may delete optional columns you do not use.'],
      ['4. Amounts must be numbers (no text). Leave a cell empty for zero.'],
      ['5. Extra pay items: add a column named "Earning: <name>" (e.g. "Earning: Site Allowance").'],
      ['6. Extra deductions: add a column named "Deduction: <name>" (e.g. "Deduction: Housing Loan").'],
      ['7. Net Salary is calculated automatically (earnings - deductions). If you fill the Net Salary column it is only used as a cross-check.'],
      ['8. Save the file, open the Payslip Generator page, upload this file, review the checks, pick a design and export the PDFs.'],
      ['9. Use English/Latin letters for names and text. Arabic script is not supported by the PDF fonts in this version.'],
      [''],
      ['COLUMNS'],
      ['Column', 'Required', 'Type', 'Group']
    ];
    COLUMNS.forEach(function (c) {
      instr.push([c.label, c.required ? 'Yes' : 'No', c.type || 'text', c.group]);
    });
    var wsInstr = XLSX.utils.aoa_to_sheet(instr);
    wsInstr['!cols'] = [{ wch: 110 }, { wch: 10 }, { wch: 10 }, { wch: 12 }];
    XLSX.utils.book_append_sheet(wb, wsInstr, SHEET_INSTRUCTIONS);

    return wb;
  }

  // ---------------------------------------------------------------------------
  // Parsing
  // ---------------------------------------------------------------------------

  function findSheet(wb, wanted) {
    var target = normalizeHeader(wanted);
    for (var i = 0; i < wb.SheetNames.length; i++) {
      if (normalizeHeader(wb.SheetNames[i]) === target) return wb.Sheets[wb.SheetNames[i]];
    }
    return null;
  }

  function parseSettings(ws, issues) {
    var settings = { currency: 'SAR', extra: [] };
    SETTINGS_FIELDS.forEach(function (f) {
      settings[f.key] = f.default != null ? f.default : '';
    });
    if (!ws) {
      issues.push({ level: 'error', where: 'Settings', message: 'Sheet "' + SHEET_SETTINGS + '" was not found. Download the template and fill the Settings sheet.' });
      return settings;
    }
    var aoa = XLSXRef.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: false });
    aoa.forEach(function (row, idx) {
      if (!row || isBlank(row[0])) return;
      var keyNorm = normalizeHeader(row[0]);
      if (keyNorm === 'setting' && normalizeHeader(row[1]) === 'value') return; // header row
      var def = SETTINGS_LOOKUP[keyNorm];
      var raw = row[1];
      if (!def) {
        if (!isBlank(raw)) settings.extra.push({ label: cleanText(row[0]), value: cleanText(raw) });
        return;
      }
      if (def.type === 'date') settings[def.key] = parseDateLike(raw);
      else settings[def.key] = cleanText(raw);
    });
    SETTINGS_FIELDS.forEach(function (f) {
      if (f.required && isBlank(settings[f.key])) {
        issues.push({ level: 'error', where: 'Settings', message: '"' + f.label + '" is empty in the Settings sheet.' });
      }
    });
    if (!isBlank(settings.currency)) {
      settings.currency = String(settings.currency).toUpperCase().trim();
      if (!/^[A-Z]{3}$/.test(settings.currency)) {
        issues.push({ level: 'warning', where: 'Settings', message: 'Currency "' + settings.currency + '" is not a 3-letter code. It will be printed as typed.' });
      }
    } else {
      settings.currency = 'SAR';
    }
    ['companyName', 'address', 'preparedBy', 'footerNote', 'payPeriod'].forEach(function (k) {
      if (hasNonLatin(settings[k])) {
        issues.push({ level: 'warning', where: 'Settings', message: '"' + settingLabel(k) + '" contains characters (e.g. Arabic) that the PDF fonts cannot print. Use English text.' });
      }
    });
    return settings;
  }

  function settingLabel(key) {
    for (var i = 0; i < SETTINGS_FIELDS.length; i++) if (SETTINGS_FIELDS[i].key === key) return SETTINGS_FIELDS[i].label;
    return key;
  }

  // Maps each header cell to a column definition (or custom earning/deduction).
  function mapHeaders(headerRow, issues) {
    var mapping = []; // index -> {kind:'column', def} | {kind:'earning'|'deduction', label}
    var seen = {};
    headerRow.forEach(function (h, idx) {
      if (isBlank(h)) {
        mapping[idx] = null;
        return;
      }
      var text = String(h).trim();
      var m = text.match(CUSTOM_EARNING_RE);
      if (m) {
        mapping[idx] = { kind: 'earning', label: m[2].trim() };
        return;
      }
      m = text.match(CUSTOM_DEDUCTION_RE);
      if (m) {
        mapping[idx] = { kind: 'deduction', label: m[2].trim() };
        return;
      }
      var def = COLUMN_LOOKUP[normalizeHeader(text)];
      if (!def) {
        mapping[idx] = { kind: 'ignored', label: text };
        issues.push({ level: 'warning', where: 'Payslips', message: 'Column "' + text + '" is not recognised and will be ignored. To add it as a pay item, rename it to "Earning: ' + text + '" or "Deduction: ' + text + '".' });
        return;
      }
      if (seen[def.key]) {
        mapping[idx] = { kind: 'ignored', label: text };
        issues.push({ level: 'warning', where: 'Payslips', message: 'Column "' + text + '" appears more than once. Only the first one is used.' });
        return;
      }
      seen[def.key] = true;
      mapping[idx] = { kind: 'column', def: def };
    });
    COLUMNS.forEach(function (c) {
      if (c.required && !seen[c.key]) {
        issues.push({ level: 'error', where: 'Payslips', message: 'Required column "' + c.label + '" is missing from the Payslips sheet.' });
      }
    });
    return mapping;
  }

  function parseRow(rawRow, mapping, rowNumber, settings) {
    var slip = {
      rowNumber: rowNumber,
      issues: [],
      earnings: [],
      deductions: [],
      fields: {},
      isExample: false
    };
    COLUMNS.forEach(function (c) {
      slip.fields[c.key] = '';
    });
    var netProvided = null;

    mapping.forEach(function (m, idx) {
      if (!m) return;
      var raw = rawRow[idx];
      if (m.kind === 'column') {
        var def = m.def;
        if (def.type === 'money' || def.type === 'number') {
          var p = parseNumber(raw);
          if (p.error) {
            slip.issues.push({ level: 'error', message: '"' + def.label + '" ' + p.error + '.' });
            slip.fields[def.key] = null;
          } else {
            slip.fields[def.key] = p.value;
          }
          if (def.key === 'netSalary') netProvided = p.value;
          if (p.value != null && p.value < 0 && def.type === 'money') {
            slip.issues.push({ level: 'error', message: '"' + def.label + '" is negative (' + formatMoney(p.value) + ').' });
          }
          if (def.group === 'earning' && p.value) slip.earnings.push({ key: def.key, label: def.label, amount: p.value });
          if (def.group === 'deduction' && p.value) slip.deductions.push({ key: def.key, label: def.label, amount: p.value });
          if (def.key === 'basicSalary' && p.value == null && !p.error) {
            slip.issues.push({ level: 'error', message: '"Basic Salary" is empty.' });
          }
        } else if (def.type === 'date') {
          slip.fields[def.key] = parseDateLike(raw);
        } else {
          slip.fields[def.key] = cleanText(raw);
        }
      } else if (m.kind === 'earning' || m.kind === 'deduction') {
        var pn = parseNumber(raw);
        if (pn.error) {
          slip.issues.push({ level: 'error', message: '"' + m.label + '" ' + pn.error + '.' });
        } else if (pn.value != null && pn.value < 0) {
          slip.issues.push({ level: 'error', message: '"' + m.label + '" is negative (' + formatMoney(pn.value) + ').' });
        } else if (pn.value) {
          slip[m.kind === 'earning' ? 'earnings' : 'deductions'].push({ key: 'custom', label: m.label, amount: pn.value });
        }
      }
    });

    // Make sure Basic Salary is always the first earning line even when 0.
    if (!slip.earnings.some(function (e) { return e.key === 'basicSalary'; }) && slip.fields.basicSalary != null) {
      slip.earnings.unshift({ key: 'basicSalary', label: 'Basic Salary', amount: slip.fields.basicSalary || 0 });
    }

    slip.employeeId = cleanText(slip.fields.employeeId);
    slip.name = cleanText(slip.fields.name);
    slip.isExample = slip.employeeId.toUpperCase().indexOf(EXAMPLE_ID_PREFIX) === 0;

    var gross = 0;
    slip.earnings.forEach(function (e) {
      gross += e.amount || 0;
    });
    var totalDed = 0;
    slip.deductions.forEach(function (d) {
      totalDed += d.amount || 0;
    });
    slip.gross = round2(gross);
    slip.totalDeductions = round2(totalDed);
    slip.net = round2(gross - totalDed);
    slip.netProvided = netProvided;

    // Row validation
    if (isBlank(slip.employeeId)) slip.issues.push({ level: 'error', message: '"Employee ID" is empty.' });
    if (isBlank(slip.name)) slip.issues.push({ level: 'error', message: '"Employee Name" is empty.' });
    if (slip.fields.basicSalary === 0) slip.issues.push({ level: 'warning', message: '"Basic Salary" is zero.' });
    if (slip.totalDeductions > slip.gross + 0.001) slip.issues.push({ level: 'error', message: 'Deductions (' + formatMoney(slip.totalDeductions) + ') exceed total earnings (' + formatMoney(slip.gross) + ').' });
    else if (slip.net <= 0 && slip.gross > 0) slip.issues.push({ level: 'warning', message: 'Net salary is zero.' });
    if (netProvided != null && Math.abs(netProvided - slip.net) > 0.005) {
      slip.issues.push({ level: 'warning', message: 'Net Salary in the file (' + formatMoney(netProvided) + ') does not match the calculated net (' + formatMoney(slip.net) + '). The calculated value will be printed.' });
    }
    var iban = String(slip.fields.iban || '').replace(/\s+/g, '').toUpperCase();
    if (iban) {
      slip.fields.iban = iban;
      if (/^SA/.test(iban) && !/^SA\d{2}[0-9A-Z]{20}$/.test(iban)) {
        slip.issues.push({ level: 'warning', message: 'IBAN "' + iban + '" does not look like a valid Saudi IBAN (SA + 22 characters).' });
      } else if (!/^[A-Z]{2}\d{2}[0-9A-Z]{11,30}$/.test(iban)) {
        slip.issues.push({ level: 'warning', message: 'IBAN "' + iban + '" does not look like a valid IBAN.' });
      }
    }
    if (slip.fields.absentDays != null && slip.fields.workingDays != null && slip.fields.absentDays > slip.fields.workingDays) {
      slip.issues.push({ level: 'warning', message: 'Absent Days is greater than Working Days.' });
    }
    var textFields = ['name', 'jobTitle', 'department', 'nationality', 'bankName', 'remarks'];
    var nonLatin = textFields.filter(function (k) {
      return hasNonLatin(slip.fields[k]);
    });
    slip.earnings.concat(slip.deductions).forEach(function (l) {
      if (hasNonLatin(l.label) && nonLatin.indexOf(l.label) < 0) nonLatin.push(l.label);
    });
    if (nonLatin.length) {
      slip.issues.push({ level: 'warning', message: 'Some text (e.g. ' + fieldLabel(nonLatin[0]) + ') contains characters the PDF fonts cannot print, such as Arabic. It will show as garbled text; use English.' });
    }

    slip.hasErrors = slip.issues.some(function (i) {
      return i.level === 'error';
    });
    return slip;
  }

  function fieldLabel(key) {
    for (var i = 0; i < COLUMNS.length; i++) if (COLUMNS[i].key === key) return COLUMNS[i].label;
    return key;
  }

  var XLSXRef = null;

  /**
   * Parse a workbook (already read with XLSX.read(..., {cellDates:true})).
   * Returns { settings, slips, issues, stats, columnsPresent }
   */
  function parseWorkbook(XLSX, wb) {
    XLSXRef = XLSX;
    var issues = [];
    var settings = parseSettings(findSheet(wb, SHEET_SETTINGS), issues);
    var wsPay = findSheet(wb, SHEET_PAYSLIPS);
    var slips = [];
    var mapping = [];
    if (!wsPay) {
      // fall back to the first sheet that is not Settings/Instructions
      var candidates = wb.SheetNames.filter(function (n) {
        var k = normalizeHeader(n);
        return k !== normalizeHeader(SHEET_SETTINGS) && k !== normalizeHeader(SHEET_INSTRUCTIONS);
      });
      if (candidates.length) {
        wsPay = wb.Sheets[candidates[0]];
        issues.push({ level: 'warning', where: 'Payslips', message: 'Sheet "' + SHEET_PAYSLIPS + '" was not found. Using sheet "' + candidates[0] + '" instead.' });
      } else {
        issues.push({ level: 'error', where: 'Payslips', message: 'Sheet "' + SHEET_PAYSLIPS + '" was not found in the file.' });
      }
    }
    if (wsPay) {
      var aoa = XLSX.utils.sheet_to_json(wsPay, { header: 1, raw: true, defval: null, blankrows: true });
      // Find the header row: first row containing "Employee ID"/"Employee Name"-ish headers
      var headerIdx = -1;
      for (var i = 0; i < Math.min(aoa.length, 20); i++) {
        var row = aoa[i] || [];
        var hits = row.filter(function (h) {
          return !isBlank(h) && COLUMN_LOOKUP[normalizeHeader(h)];
        }).length;
        if (hits >= 2) {
          headerIdx = i;
          break;
        }
      }
      if (headerIdx < 0) {
        issues.push({ level: 'error', where: 'Payslips', message: 'Could not find the header row in the Payslips sheet. Keep the header row from the template.' });
      } else {
        mapping = mapHeaders(aoa[headerIdx], issues);
        for (var r = headerIdx + 1; r < aoa.length; r++) {
          var raw = aoa[r] || [];
          var empty = raw.every(isBlank);
          if (empty) continue;
          var slip = parseRow(raw, mapping, r + 1, settings);
          if (slip.isExample) continue;
          slips.push(slip);
        }
      }
      // Duplicate IDs
      var byId = {};
      slips.forEach(function (s) {
        var k = s.employeeId.toUpperCase();
        if (!k) return;
        (byId[k] = byId[k] || []).push(s);
      });
      Object.keys(byId).forEach(function (k) {
        if (byId[k].length > 1) {
          var rowsTxt = byId[k].map(function (s) { return s.rowNumber; }).join(', ');
          byId[k].forEach(function (s) {
            s.issues.push({ level: 'error', message: 'Employee ID "' + s.employeeId + '" is duplicated (rows ' + rowsTxt + ').' });
            s.hasErrors = true;
          });
        }
      });
      if (!slips.length) {
        issues.push({ level: 'error', where: 'Payslips', message: 'No employee rows were found in the Payslips sheet.' });
      }
    }

    var stats = {
      total: slips.length,
      withErrors: slips.filter(function (s) { return s.hasErrors; }).length,
      withWarnings: slips.filter(function (s) { return !s.hasErrors && s.issues.length; }).length,
      fileErrors: issues.filter(function (i) { return i.level === 'error'; }).length,
      fileWarnings: issues.filter(function (i) { return i.level === 'warning'; }).length,
      totalNet: round2(slips.reduce(function (a, s) { return a + (s.hasErrors ? 0 : s.net); }, 0)),
      totalGross: round2(slips.reduce(function (a, s) { return a + (s.hasErrors ? 0 : s.gross); }, 0))
    };
    stats.ok = stats.total - stats.withErrors;

    return {
      settings: settings,
      slips: slips,
      issues: issues,
      stats: stats,
      mapping: mapping.filter(Boolean)
    };
  }

  // ---------------------------------------------------------------------------
  // Sample data (used for the "Load sample" button and the committed sample file)
  // ---------------------------------------------------------------------------

  function sampleRows(count) {
    var first = ['Ahmed', 'Mohammed', 'Fatima', 'Rajesh', 'Maria', 'John', 'Khalid', 'Priya', 'Omar', 'Sara', 'Ali', 'Noor', 'Carlos', 'Aisha', 'Hassan', 'Deepak', 'Lina', 'Yusuf', 'Grace', 'Tariq'];
    var last = ['Hassan', 'Al-Otaibi', 'Khan', 'Kumar', 'Santos', 'Smith', 'Ibrahim', 'Sharma', 'Abdullah', 'Youssef', 'Rahman', 'Siddiqui', 'Reyes', 'Malik', 'Farouk', 'Patel', 'Haddad', 'Mansour', 'Okafor', 'Aziz'];
    var titles = ['Site Engineer', 'Foreman', 'Electrician', 'Driver', 'Accountant', 'Safety Officer', 'Mason', 'Surveyor', 'HR Officer', 'Plumber'];
    var depts = ['Projects', 'Operations', 'Finance', 'HSE', 'Administration'];
    var nats = ['Saudi', 'Egyptian', 'Indian', 'Pakistani', 'Filipino', 'Bangladeshi', 'Sudanese', 'Jordanian'];
    var banks = ['Al Rajhi Bank', 'SNB', 'Riyad Bank', 'Alinma Bank', 'Bank Albilad'];
    var rows = [];
    for (var i = 0; i < count; i++) {
      var basic = 3000 + ((i * 937) % 9000);
      var housing = Math.round(basic * 0.25);
      var transport = 300 + ((i * 50) % 400);
      var food = i % 3 === 0 ? 300 : 0;
      var otHours = (i * 7) % 25;
      var otPay = Math.round(((basic / 240) * 1.5 * otHours) * 100) / 100;
      var absentDays = i % 11 === 0 ? 2 : 0;
      var absence = absentDays ? Math.round(((basic + housing + transport) / 30) * absentDays * 100) / 100 : 0;
      var gosi = nats[i % nats.length] === 'Saudi' ? Math.round((basic + housing) * 0.0975 * 100) / 100 : 0;
      var loan = i % 7 === 0 ? 500 : 0;
      var nationality = nats[i % nats.length];
      rows.push([
        'EMP-' + String(i + 1).padStart(3, '0'),
        first[i % first.length] + ' ' + last[(i * 3) % last.length],
        String(2000000000 + i * 12345),
        nationality,
        titles[i % titles.length],
        depts[i % depts.length],
        new Date(2022 + (i % 4), i % 12, 1 + (i % 27)),
        banks[i % banks.length],
        'SA' + String((1000000000 + i * 987654) % 10000000000).padStart(10, '0') + String((200000000000 + i * 12345) % 1000000000000).padStart(12, '0'),
        30 - absentDays,
        absentDays,
        otHours,
        basic,
        housing,
        transport,
        food,
        otPay,
        0,
        gosi,
        absence,
        loan,
        0,
        '',
        i % 9 === 0 ? 'Includes overtime for site work' : ''
      ]);
    }
    return rows;
  }

  return {
    SHEET_SETTINGS: SHEET_SETTINGS,
    SHEET_PAYSLIPS: SHEET_PAYSLIPS,
    SETTINGS_FIELDS: SETTINGS_FIELDS,
    COLUMNS: COLUMNS,
    buildTemplateWorkbook: buildTemplateWorkbook,
    parseWorkbook: parseWorkbook,
    sampleRows: sampleRows,
    formatMoney: formatMoney,
    formatDate: formatDate,
    amountToWords: amountToWords,
    parseNumber: parseNumber,
    parseDateLike: parseDateLike,
    normalizeHeader: normalizeHeader,
    hasNonLatin: hasNonLatin,
    round2: round2
  };
});
