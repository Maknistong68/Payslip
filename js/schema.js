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
    { key: 'companyName', label: 'Company Name', required: true, example: 'Example Contracting Co. Ltd.', note: 'Required. Legal name as it appears on the CR.', aliases: ['company', 'employer', 'employer name', 'establishment name', 'establishment'] },
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
    { key: 'footerNote', label: 'Footer Note', example: 'This is a system-generated payslip and does not require a signature.', note: 'Optional. Free text printed at the bottom of each payslip.', aliases: ['note', 'footer', 'disclaimer'] },
    { key: 'logo', label: 'Logo', type: 'image', example: '', note: 'Optional. Insert your logo picture anywhere on this sheet (Insert > Pictures). The first picture on this sheet is used as the logo.', aliases: ['company logo'] },
    { key: 'stamp', label: 'Stamp / Signature', type: 'image', example: '', note: 'Optional. The second picture on this sheet is used as the stamp or signature.', aliases: ['stamp', 'signature'] }
  ];

  var COLUMNS = [
    { key: 'employeeId', label: 'Employee ID', group: 'info', required: true, example: 'EMP-001', aliases: ['emp id', 'employee no', 'employee number', 'emp no', 'id', 'staff id', 'staff no', 'employee code', 'emp code', 'code', 'badge no', 'badge'] },
    { key: 'name', label: 'Employee Name', group: 'info', required: true, example: 'Ahmed Ali Hassan', aliases: ['name', 'full name', 'employee', 'staff name', 'worker name', 'name of employee'] },
    { key: 'nationalId', label: 'National ID / Iqama', group: 'info', example: '2456789012', aliases: ['iqama', 'iqama no', 'iqama number', 'national id', 'id number', 'national id iqama', 'iqama id', 'id no', 'identity number', 'resident id', 'nid'] },
    { key: 'nationality', label: 'Nationality', group: 'info', example: 'Egyptian', aliases: ['nation', 'country'] },
    { key: 'jobTitle', label: 'Job Title', group: 'info', example: 'Site Engineer', aliases: ['title', 'position', 'designation', 'job', 'profession', 'occupation', 'role'] },
    { key: 'department', label: 'Department', group: 'info', example: 'Projects', aliases: ['dept', 'section', 'project', 'division', 'site', 'location'] },
    { key: 'joiningDate', label: 'Joining Date', group: 'info', type: 'date', example: '2024-03-15', aliases: ['date of joining', 'doj', 'hire date', 'start date', 'join date', 'date joined', 'hiring date'] },
    { key: 'bankName', label: 'Bank Name', group: 'info', example: 'Al Rajhi Bank', aliases: ['bank'] },
    { key: 'iban', label: 'IBAN', group: 'info', example: 'SA0380000000608010167519', aliases: ['iban no', 'iban number', 'account', 'account number', 'bank account', 'account no', 'acc no'] },
    { key: 'workingDays', label: 'Working Days', group: 'attendance', type: 'number', example: 30, aliases: ['days worked', 'work days', 'paid days', 'days', 'present days', 'no of days', 'number of days'] },
    { key: 'absentDays', label: 'Absent Days', group: 'attendance', type: 'number', example: 0, aliases: ['absence days', 'absent', 'unpaid days', 'lop days', 'leave days', 'absences'] },
    { key: 'overtimeHours', label: 'Overtime Hours', group: 'attendance', type: 'number', example: 10, aliases: ['ot hours', 'overtime hrs', 'ot hrs', 'ot hr', 'overtime hour'] },
    { key: 'basicSalary', label: 'Basic Salary', group: 'earning', required: true, type: 'money', example: 4000, aliases: ['basic', 'basic pay', 'base salary', 'basic wage', 'salary', 'wage', 'base pay', 'monthly salary', 'basic amount'] },
    { key: 'housing', label: 'Housing Allowance', group: 'earning', type: 'money', example: 1000, aliases: ['housing', 'hra', 'house allowance', 'accommodation allowance', 'accommodation', 'house rent', 'rent allowance', 'housing all'] },
    { key: 'transport', label: 'Transport Allowance', group: 'earning', type: 'money', example: 400, aliases: ['transport', 'transportation', 'transportation allowance', 'conveyance', 'travel allowance', 'transport all', 'car allowance'] },
    { key: 'food', label: 'Food Allowance', group: 'earning', type: 'money', example: 300, aliases: ['food', 'meal allowance', 'meals', 'food all'] },
    { key: 'overtimePay', label: 'Overtime Pay', group: 'earning', type: 'money', example: 250, aliases: ['overtime', 'ot', 'ot pay', 'overtime amount', 'ot amount', 'overtime allowance'] },
    { key: 'otherEarnings', label: 'Other Earnings', group: 'earning', type: 'money', example: 0, aliases: ['other allowance', 'other allowances', 'bonus', 'other earning', 'additions', 'other', 'others', 'incentive', 'incentives', 'commission', 'other income', 'other pay', 'other addition'] },
    { key: 'gosi', label: 'GOSI Deduction', group: 'deduction', type: 'money', example: 0, aliases: ['gosi', 'social insurance', 'gosi employee', 'gosi ded', 'pension'] },
    { key: 'absence', label: 'Absence Deduction', group: 'deduction', type: 'money', example: 0, aliases: ['absence', 'absent deduction', 'lop', 'unpaid leave', 'absence amount', 'absent amount', 'leave deduction'] },
    { key: 'loan', label: 'Loan / Advance', group: 'deduction', type: 'money', example: 0, aliases: ['loan', 'advance', 'salary advance', 'loan deduction', 'loan advance', 'advance deduction', 'loan ded'] },
    { key: 'otherDeductions', label: 'Other Deductions', group: 'deduction', type: 'money', example: 0, aliases: ['other deduction', 'deductions', 'deduction', 'penalty', 'penalties', 'fines', 'fine', 'other ded', 'violation', 'violations', 'misc deduction'] },
    { key: 'netSalary', label: 'Net Salary', group: 'check', type: 'money', example: '', aliases: ['net', 'net pay', 'net amount', 'total net', 'net paid', 'take home', 'net payable', 'total payable', 'amount payable', 'payable'] },
    { key: 'grossSalary', label: 'Gross Salary', group: 'check', type: 'money', example: '', inTemplate: false, aliases: ['gross', 'gross pay', 'total earnings', 'total earning', 'total salary', 'gross amount', 'total allowances', 'total'] },
    { key: 'totalDeductions', label: 'Total Deductions', group: 'check', type: 'money', example: '', inTemplate: false, aliases: ['total deduction', 'total ded', 'deductions total', 'sum of deductions'] },
    { key: 'remarks', label: 'Remarks', group: 'info', example: '', aliases: ['notes', 'comment', 'comments', 'remark', 'description'] }
  ];

  // Column headers like "Earning: Site Allowance" / "Deduction: Housing Loan" are
  // treated as extra earning / deduction lines without changing the app.
  var CUSTOM_EARNING_RE = /^(earning|earnings|allowance|addition|add|income|pay)\s*[:\-]\s*(.+)$/i;
  var CUSTOM_DEDUCTION_RE = /^(deduction|deductions|ded|less|minus)\s*[:\-]\s*(.+)$/i;

  // Keyword rules used when a header does not match a known label/alias exactly.
  // Order matters: the first matching rule wins.
  var GUESS_RULES = [
    [/\b(emp|employee|staff|worker|personnel|payroll)\b.*\b(id|no|num|number|code)\b|\b(id|no|num|number|code)\b.*\b(emp|employee|staff|worker)\b|^(s|sl|sr|si)\s*no$|^(sno|slno|srno)$|^serial|^#$|^no$/, 'employeeId'],
    [/iqama|national\s*id|\bnid\b|identity|resident\s*id|\bid\s*(no|num|number)\b|civil\s*id|passport/, 'nationalId'],
    [/\bname\b/, 'name'],
    [/nationality|citizen/, 'nationality'],
    [/title|position|designation|profession|occupation|\bjob\b|\brole\b|trade/, 'jobTitle'],
    [/depart|dept|section|division|\bunit\b|cost\s*cent|\bproject\b|\bsite\b/, 'department'],
    [/joining|hire|hiring|\bdoj\b|date\s*of\s*join|join\s*date|start\s*date|contract\s*start/, 'joiningDate'],
    [/\biban\b|account\s*(no|num|number)?|\bacc\b|\ba\s*c\b/, 'iban'],
    [/\bbank\b/, 'bankName'],
    [/(working|work|paid|present|actual|attendance|no\s*of)\s*days|days\s*(worked|present|paid)/, 'workingDays'],
    [/(absen|unpaid|lop|leave)\w*\s*days|days\s*absen/, 'absentDays'],
    [/(overtime|\bot\b)\s*\(?\s*(hour|hrs|hr)|(hour|hrs)\s*(of\s*)?(overtime|\bot\b)/, 'overtimeHours'],
    [/overtime|\bot\b/, 'overtimePay'],
    [/\bnet\b|take\s*home|payable|to\s*be\s*paid|paid\s*amount|final\s*salary|total\s*pay\b/, 'netSalary'],
    [/total\s*ded|ded\w*\s*total|sum\s*of\s*ded/, 'totalDeductions'],
    [/gross|total\s*earn|total\s*allow|total\s*salary|total\s*income|^total$/, 'grossSalary'],
    [/basic|\bbase\b|^salary$|^wage/, 'basicSalary'],
    [/hous|\bhra\b|accommodat|\brent\b/, 'housing'],
    [/transport|conveyance|travel|\bcar\b|fuel|petrol/, 'transport'],
    [/food|meal|canteen/, 'food'],
    [/gosi|social\s*ins|pension|\bsi\b/, 'gosi'],
    [/absen|\blop\b|unpaid|leave\s*ded/, 'absence'],
    [/loan|advance|\badv\b/, 'loan'],
    [/remark|note|comment|descr/, 'remarks'],
    // generic buckets -> custom lines named after the header
    [/deduct|penalt|fine|violation|medical|insurance|tax|recover|damage|\bless\b|\bded\b/, '*deduction'],
    [/allow|bonus|incentive|commission|reward|arrear|reimburs|ticket|mobile|phone|shift|hardship|danger|special|site\s*all|other\s*(earn|pay|income|add)/, '*earning']
  ];

  var EXAMPLE_ID_PREFIX = 'EXAMPLE';
  var TOTAL_ROW_RE = /^(grand\s*)?(total|totals|sum|subtotal|sub\s*total|net\s*total|g\.?\s*total)\b/i;

  // ---------------------------------------------------------------------------
  // Small helpers
  // ---------------------------------------------------------------------------

  function normalizeHeader(s) {
    return String(s == null ? '' : s)
      .toLowerCase()
      .replace(/\(.*?\)/g, ' ') // drop "(SAR)" style suffixes
      .replace(/[*#:]/g, ' ')
      .replace(/\./g, '') // "S.No" -> "sno", "O.T." -> "ot"
      .replace(/[^\p{L}\p{N}]+/gu, ' ')
      .replace(/\s+/g, ' ')
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
  var COLUMN_BY_KEY = {};
  COLUMNS.forEach(function (c) { COLUMN_BY_KEY[c.key] = c; });

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
    if (/^[-–—]+$/.test(s) || /^(n\/?a|nil|none|null)$/i.test(s)) return { value: null, error: null };
    // strip currency codes/symbols, thousands separators and whitespace
    var cleaned = s.replace(/[^0-9.\-()]/g, '');
    var negative = /^\(.*\)$/.test(cleaned) || /^-/.test(cleaned);
    cleaned = cleaned.replace(/[()\-]/g, '');
    if (cleaned === '' || cleaned === '.') return { value: null, error: 'is not a number ("' + s + '")' };
    // "1.234.567,89" style: more than one dot and no comma handled above; keep last dot as decimal
    if ((cleaned.match(/\./g) || []).length > 1) {
      var last = cleaned.lastIndexOf('.');
      cleaned = cleaned.slice(0, last).replace(/\./g, '') + cleaned.slice(last);
    }
    var n = Number(cleaned);
    if (!isFinite(n)) return { value: null, error: 'is not a number ("' + s + '")' };
    return { value: negative ? -n : n, error: null };
  }

  function looksNumeric(v) {
    if (typeof v === 'number') return true;
    if (isBlank(v)) return false;
    var p = parseNumber(v);
    return p.value != null && !p.error;
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

  // The PDF fonts cover Latin (built-in Helvetica/Times) and Arabic (bundled Amiri).
  // Anything else (Cyrillic, CJK, Devanagari...) prints as "?".
  var UNSUPPORTED_RE = /[^\u0000-ɏ؀-ۿݐ-ݿࢠ-ࣿﭐ-﷿ﹰ-﻿ -⁯₠-⃏]/;

  function hasNonLatin(s) {
    return UNSUPPORTED_RE.test(String(s || ''));
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

  function templateColumns() {
    return COLUMNS.filter(function (c) { return c.inTemplate !== false; });
  }

  /**
   * options.sampleRows   -> array of row arrays (in templateColumns() order)
   * options.settings     -> object of setting values to pre-fill (key -> value)
   */
  function buildTemplateWorkbook(XLSX, options) {
    options = options || {};
    var withSample = !!options.sampleRows;
    var preset = options.settings || null;
    var wb = XLSX.utils.book_new();

    // Settings sheet
    var settingsAoa = [['Setting', 'Value', 'Notes']];
    SETTINGS_FIELDS.forEach(function (f) {
      var value;
      if (f.type === 'image') value = '(optional: insert the picture anywhere on this sheet)';
      else if (preset && preset[f.key] != null && preset[f.key] !== '') value = preset[f.key];
      else if (withSample) value = f.example;
      else value = f.default != null ? f.default : '';
      settingsAoa.push([f.label, value, f.note + (f.example ? ' Example: ' + f.example : '')]);
    });
    var wsSettings = XLSX.utils.aoa_to_sheet(settingsAoa);
    wsSettings['!cols'] = [{ wch: 22 }, { wch: 44 }, { wch: 90 }];
    XLSX.utils.book_append_sheet(wb, wsSettings, SHEET_SETTINGS);

    // Payslips sheet
    var cols = templateColumns();
    var header = cols.map(function (c) { return c.label; });
    var rows = [header];
    if (withSample) {
      rows = rows.concat(options.sampleRows);
    } else {
      // One example row that the parser recognises and ignores.
      rows.push(cols.map(function (c) {
        if (c.key === 'employeeId') return EXAMPLE_ID_PREFIX + '-001';
        if (c.key === 'name') return 'Example Row - delete me';
        if (c.key === 'netSalary') return '';
        return c.example;
      }));
    }
    var wsPay = XLSX.utils.aoa_to_sheet(rows);
    wsPay['!cols'] = cols.map(function (c) {
      return { wch: Math.max(12, c.label.length + 2) };
    });
    wsPay['!freeze'] = { xSplit: 2, ySplit: 1 };
    XLSX.utils.book_append_sheet(wb, wsPay, SHEET_PAYSLIPS);

    // Instructions sheet
    var instr = [
      ['HOW TO USE THIS TEMPLATE'],
      [''],
      ['1. Fill the "Settings" sheet with your company details. Company Name, CR Number and Pay Period are required.'],
      ['2. Optional: insert your logo picture on the "Settings" sheet (Insert > Pictures). A second picture is used as stamp/signature.'],
      ['3. Fill the "Payslips" sheet with one row per employee. Delete the example row.'],
      ['4. Keep the header row. You may delete optional columns you do not use, and reorder columns freely.'],
      ['5. Amounts must be numbers (no text). Leave a cell empty for zero.'],
      ['6. Extra pay items: add a column named "Earning: <name>" (e.g. "Earning: Site Allowance").'],
      ['7. Extra deductions: add a column named "Deduction: <name>" (e.g. "Deduction: Housing Loan").'],
      ['8. Net Salary is calculated automatically (earnings - deductions). If you fill the Net Salary column it is only used as a cross-check.'],
      ['9. Save the file, open the Payslip Generator page, upload this file, review the checks, pick a design and export the PDFs.'],
      ['10. Names and text can be in English or Arabic (or both).'],
      [''],
      ['COLUMNS'],
      ['Column', 'Required', 'Type', 'Group']
    ];
    cols.forEach(function (c) {
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

  var XLSXRef = null;

  function findSheet(wb, wanted) {
    var target = normalizeHeader(wanted);
    for (var i = 0; i < wb.SheetNames.length; i++) {
      if (normalizeHeader(wb.SheetNames[i]) === target) return wb.Sheets[wb.SheetNames[i]];
    }
    return null;
  }

  function settingLabel(key) {
    for (var i = 0; i < SETTINGS_FIELDS.length; i++) if (SETTINGS_FIELDS[i].key === key) return SETTINGS_FIELDS[i].label;
    return key;
  }

  function emptySettings() {
    var settings = { extra: [] };
    SETTINGS_FIELDS.forEach(function (f) {
      if (f.type === 'image') return;
      settings[f.key] = f.default != null ? f.default : '';
    });
    return settings;
  }

  // Looks for key/value pairs in a sheet (two adjacent cells where the first is a known setting name).
  function parseSettings(ws, issues, result) {
    var settings = emptySettings();
    result.settingsFound = false;
    if (!ws) {
      issues.push({ level: 'warning', where: 'Settings', code: 'no-settings-sheet', message: 'No "' + SHEET_SETTINGS + '" sheet was found. Enter the company details on the page (they are saved for next time).' });
      return settings;
    }
    var aoa = XLSXRef.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: false });
    var found = 0;
    aoa.forEach(function (row) {
      if (!row) return;
      for (var c = 0; c < row.length - 1; c++) {
        if (isBlank(row[c])) continue;
        var keyNorm = normalizeHeader(row[c]);
        if (keyNorm === 'setting' && normalizeHeader(row[c + 1]) === 'value') continue; // header row
        var def = SETTINGS_LOOKUP[keyNorm];
        var raw = row[c + 1];
        if (!def) {
          if (c === 0 && !isBlank(raw) && typeof row[c] === 'string') settings.extra.push({ label: cleanText(row[c]), value: cleanText(raw) });
          continue;
        }
        if (def.type === 'image') { found++; break; }
        found++;
        if (def.type === 'date') settings[def.key] = parseDateLike(raw);
        else settings[def.key] = cleanText(raw);
        break;
      }
    });
    result.settingsFound = found > 0;
    if (!result.settingsFound) {
      issues.push({ level: 'warning', where: 'Settings', code: 'no-settings-sheet', message: 'The "' + SHEET_SETTINGS + '" sheet has no recognised settings. Enter the company details on the page.' });
    }
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
        issues.push({ level: 'warning', where: 'Settings', message: '"' + settingLabel(k) + '" contains characters the PDF fonts cannot print (only English and Arabic are supported).' });
      }
    });
    return settings;
  }

  // Guess what a header means. Returns {kind:'column', key} | {kind:'earning'|'deduction', label} | null
  function guessHeader(text) {
    var t = String(text).trim();
    var m = t.match(CUSTOM_EARNING_RE);
    if (m) return { kind: 'earning', label: m[2].trim(), confidence: 'exact' };
    m = t.match(CUSTOM_DEDUCTION_RE);
    if (m) return { kind: 'deduction', label: m[2].trim(), confidence: 'exact' };
    var norm = normalizeHeader(t);
    if (!norm) return null;
    var def = COLUMN_LOOKUP[norm];
    if (def) return { kind: 'column', key: def.key, confidence: 'exact' };
    // also try without trailing words like "amount", "sar", "in sar"
    var stripped = norm.replace(/\b(amount|amt|value|in|sar|aed|riyal|riyals|sr|per month|monthly)\b/g, ' ').replace(/\s+/g, ' ').trim();
    if (stripped && COLUMN_LOOKUP[stripped] && COLUMN_LOOKUP[stripped].type === 'money') return { kind: 'column', key: COLUMN_LOOKUP[stripped].key, confidence: 'exact' };
    for (var i = 0; i < GUESS_RULES.length; i++) {
      if (GUESS_RULES[i][0].test(norm)) {
        var key = GUESS_RULES[i][1];
        if (key === '*earning') return { kind: 'earning', label: t.replace(/\s+/g, ' ').trim(), confidence: 'guess' };
        if (key === '*deduction') return { kind: 'deduction', label: t.replace(/\s+/g, ' ').trim(), confidence: 'guess' };
        return { kind: 'column', key: key, confidence: 'guess' };
      }
    }
    return null;
  }

  function headerScore(row) {
    var score = 0;
    (row || []).forEach(function (h) {
      if (isBlank(h) || typeof h !== 'string' && typeof h !== 'number') return;
      if (looksNumeric(h) && typeof h === 'number') return;
      var g = guessHeader(h);
      if (g && g.confidence === 'exact') score += 1;
      else if (g) score += 0.5;
    });
    return score;
  }

  // Picks the data sheet and the header row in it.
  function locateHeader(XLSX, wb, issues) {
    var skip = [SHEET_SETTINGS, SHEET_INSTRUCTIONS].map(normalizeHeader);
    var candidates = wb.SheetNames.filter(function (n) { return skip.indexOf(normalizeHeader(n)) < 0; });
    if (!candidates.length) candidates = wb.SheetNames.slice();
    var best = null;
    candidates.forEach(function (name) {
      var ws = wb.Sheets[name];
      if (!ws) return;
      var aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: true });
      var limit = Math.min(aoa.length, 40);
      for (var i = 0; i < limit; i++) {
        var s = headerScore(aoa[i]);
        var dataRows = aoa.length - i - 1;
        var isPreferred = normalizeHeader(name) === normalizeHeader(SHEET_PAYSLIPS);
        var total = s + (isPreferred ? 0.75 : 0) + Math.min(dataRows, 5) * 0.01;
        if (!best || total > best.total) best = { name: name, aoa: aoa, headerIdx: i, score: s, total: total };
      }
    });
    if (!best || best.score < 1.5) {
      // Nothing recognisable: fall back to the first non-empty row of the first candidate sheet.
      var name0 = candidates[0];
      var aoa0 = XLSX.utils.sheet_to_json(wb.Sheets[name0], { header: 1, raw: true, defval: null, blankrows: true });
      var idx = 0;
      while (idx < aoa0.length && (aoa0[idx] || []).every(isBlank)) idx++;
      if (idx >= aoa0.length) {
        issues.push({ level: 'error', where: 'Payslips', code: 'no-data', message: 'The file has no data rows.' });
        return null;
      }
      issues.push({ level: 'error', where: 'Payslips', code: 'no-header', message: 'No recognisable header row was found. Map your columns below (Employee Name and at least one salary column are needed), or download the template and copy your data into it.' });
      return { name: name0, aoa: aoa0, headerIdx: idx, score: 0 };
    }
    if (normalizeHeader(best.name) !== normalizeHeader(SHEET_PAYSLIPS)) {
      issues.push({ level: 'info', where: 'Payslips', message: 'Reading employee rows from sheet "' + best.name + '" (row ' + (best.headerIdx + 1) + ' used as the header).' });
    } else if (best.headerIdx > 0) {
      issues.push({ level: 'info', where: 'Payslips', message: 'Row ' + (best.headerIdx + 1) + ' of sheet "' + best.name + '" is used as the header row.' });
    }
    return best;
  }

  // Row above the header may hold merged group titles such as "Earnings" / "Deductions".
  function groupRowHints(aoa, headerIdx) {
    var hints = [];
    if (headerIdx < 1) return hints;
    var row = aoa[headerIdx - 1] || [];
    var current = null;
    var any = false;
    for (var c = 0; c < Math.max(row.length, (aoa[headerIdx] || []).length); c++) {
      var v = row[c];
      if (!isBlank(v) && typeof v === 'string') {
        var n = normalizeHeader(v);
        if (/earn|allow|addition|income|pay\b|salary/.test(n)) { current = 'earning'; any = true; }
        else if (/deduct|less|minus/.test(n)) { current = 'deduction'; any = true; }
        else current = null;
      }
      hints[c] = current;
    }
    return any ? hints : [];
  }

  /**
   * Maps each header cell to a target. overrides: { [normalizedHeader]: target } where target is
   * {kind:'column', key} | {kind:'earning'|'deduction', label} | {kind:'ignored'}
   */
  function mapHeaders(headerRow, groupHints, overrides, issues, sampleRows) {
    var mapping = [];
    var seen = {};
    var columns = [];
    overrides = overrides || {};

    headerRow.forEach(function (h, idx) {
      if (isBlank(h)) {
        mapping[idx] = null;
        return;
      }
      var text = cleanText(h);
      var norm = normalizeHeader(text);
      var samples = sampleRows.map(function (r) { return r[idx]; }).filter(function (v) { return !isBlank(v); });
      var numericRatio = samples.length ? samples.filter(looksNumeric).length / samples.length : 0;
      var target;
      var source;
      if (overrides[norm]) {
        target = overrides[norm];
        source = 'manual';
      } else {
        target = guessHeader(text);
        source = target ? target.confidence : 'none';
        if (!target && groupHints[idx] && numericRatio > 0.5) {
          target = { kind: groupHints[idx], label: text };
          source = 'guess';
        }
        if (!target) target = { kind: 'ignored' };
        // Sanity-check keyword guesses against the actual values in the column.
        if (source === 'guess' && target.kind === 'column') {
          var def0 = COLUMN_BY_KEY[target.key];
          if (def0 && (def0.type === 'money' || def0.type === 'number') && samples.length >= 3 && numericRatio < 0.5) {
            target = { kind: 'ignored' };
            source = 'none';
          } else if (def0 && !def0.type && samples.length >= 2 && numericRatio > 0.8 && def0.key !== 'employeeId' && def0.key !== 'nationalId' && def0.key !== 'iban') {
            // a text field guessed for a column full of numbers: use the group row if any, else treat as unknown
            target = groupHints[idx] ? { kind: groupHints[idx], label: text } : { kind: 'ignored' };
            source = groupHints[idx] ? 'guess' : 'none';
          }
        }
      }
      var finalKind = target.kind;
      var m;
      if (finalKind === 'column') {
        var def = COLUMN_BY_KEY[target.key];
        if (!def) {
          m = { kind: 'ignored', label: text };
        } else if (seen[def.key]) {
          m = { kind: 'ignored', label: text };
          issues.push({ level: 'warning', where: 'Payslips', message: 'Column "' + text + '" also maps to "' + def.label + '" (already used by another column). It is ignored; change the mapping if this is wrong.' });
          source = 'duplicate';
        } else {
          seen[def.key] = text;
          m = { kind: 'column', def: def };
        }
      } else if (finalKind === 'earning' || finalKind === 'deduction') {
        m = { kind: finalKind, label: target.label || text };
      } else {
        m = { kind: 'ignored', label: text };
      }
      m.header = text;
      m.norm = norm;
      m.index = idx;
      m.source = source;
      m.numericRatio = numericRatio;
      m.samples = samples.slice(0, 3).map(cleanText);
      mapping[idx] = m;
      columns.push(m);
    });

    var ignored = columns.filter(function (c) { return c.kind === 'ignored'; });
    if (ignored.length) {
      issues.push({ level: 'warning', where: 'Payslips', code: 'ignored-columns', message: ignored.length + ' column' + (ignored.length === 1 ? ' is' : 's are') + ' not recognised and will be ignored: ' + ignored.map(function (c) { return '"' + c.header + '"'; }).join(', ') + '. Use "Columns" below to map them.' });
    }
    var guessed = columns.filter(function (c) { return c.source === 'guess'; });
    if (guessed.length) {
      issues.push({ level: 'info', where: 'Payslips', code: 'guessed-columns', message: guessed.length + ' column' + (guessed.length === 1 ? ' was' : 's were') + ' matched by keyword: ' + guessed.map(function (c) { return '"' + c.header + '" → ' + describeTarget(c); }).join(', ') + '. Check them under "Columns".' });
    }
    return { mapping: mapping, columns: columns, seen: seen };
  }

  function describeTarget(m) {
    if (m.kind === 'column') return m.def.label;
    if (m.kind === 'earning') return 'Earning "' + m.label + '"';
    if (m.kind === 'deduction') return 'Deduction "' + m.label + '"';
    return 'Ignored';
  }

  function parseRow(rawRow, mapping, rowNumber) {
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
    var checks = {};

    mapping.forEach(function (m, idx) {
      if (!m) return;
      var raw = rawRow[idx];
      if (m.kind === 'column') {
        var def = m.def;
        if (def.type === 'money' || def.type === 'number') {
          var p = parseNumber(raw);
          if (p.error) {
            slip.issues.push({ level: 'error', message: '"' + m.header + '" ' + p.error + '.' });
            slip.fields[def.key] = null;
          } else {
            slip.fields[def.key] = p.value;
          }
          if (def.group === 'check') checks[def.key] = p.value;
          if (p.value != null && p.value < 0 && def.type === 'money' && def.group !== 'check') {
            slip.issues.push({ level: 'error', message: '"' + m.header + '" is negative (' + formatMoney(p.value) + ').' });
          }
          if (def.group === 'earning' && p.value) slip.earnings.push({ key: def.key, label: def.label, amount: p.value });
          if (def.group === 'deduction' && p.value) slip.deductions.push({ key: def.key, label: def.label, amount: p.value });
          if (def.key === 'basicSalary' && p.value == null && !p.error) {
            slip.issues.push({ level: 'error', message: '"' + m.header + '" is empty.' });
          }
        } else if (def.type === 'date') {
          slip.fields[def.key] = parseDateLike(raw);
        } else {
          slip.fields[def.key] = cleanText(raw);
        }
      } else if (m.kind === 'earning' || m.kind === 'deduction') {
        var pn = parseNumber(raw);
        if (pn.error) {
          slip.issues.push({ level: 'error', message: '"' + m.header + '" ' + pn.error + '.' });
        } else if (pn.value != null && pn.value < 0) {
          slip.issues.push({ level: 'error', message: '"' + m.header + '" is negative (' + formatMoney(pn.value) + ').' });
        } else if (pn.value) {
          slip[m.kind === 'earning' ? 'earnings' : 'deductions'].push({ key: 'custom', label: m.label, amount: pn.value });
        }
      }
    });

    // Make sure Basic Salary is always the first earning line even when 0.
    if (!slip.earnings.some(function (e) { return e.key === 'basicSalary'; }) && slip.fields.basicSalary != null && slip.fields.basicSalary !== '') {
      slip.earnings.unshift({ key: 'basicSalary', label: 'Basic Salary', amount: slip.fields.basicSalary || 0 });
    }

    slip.employeeId = cleanText(slip.fields.employeeId);
    slip.name = cleanText(slip.fields.name);
    slip.isExample = slip.employeeId.toUpperCase().indexOf(EXAMPLE_ID_PREFIX) === 0 || /^example row/i.test(slip.name);

    var gross = 0;
    slip.earnings.forEach(function (e) { gross += e.amount || 0; });
    var totalDed = 0;
    slip.deductions.forEach(function (d) { totalDed += d.amount || 0; });
    slip.gross = round2(gross);
    slip.totalDeductions = round2(totalDed);
    slip.net = round2(gross - totalDed);
    slip.netProvided = checks.netSalary != null ? checks.netSalary : null;

    // Row validation
    if (isBlank(slip.name)) slip.issues.push({ level: 'error', message: 'Employee name is empty.' });
    if (slip.fields.basicSalary === 0) slip.issues.push({ level: 'warning', message: 'Basic salary is zero.' });
    if (!slip.earnings.length || slip.gross === 0) slip.issues.push({ level: 'warning', message: 'No earnings found for this row.' });
    if (slip.totalDeductions > slip.gross + 0.001) slip.issues.push({ level: 'error', message: 'Deductions (' + formatMoney(slip.totalDeductions) + ') exceed total earnings (' + formatMoney(slip.gross) + ').' });
    else if (slip.net <= 0 && slip.gross > 0) slip.issues.push({ level: 'warning', message: 'Net salary is zero.' });
    if (checks.netSalary != null && Math.abs(checks.netSalary - slip.net) > 0.005) {
      slip.issues.push({ level: 'warning', message: 'Net salary in the file (' + formatMoney(checks.netSalary) + ') does not match the calculated net (' + formatMoney(slip.net) + '). The calculated value will be printed; check for an earning or deduction column that was not recognised.' });
    }
    if (checks.grossSalary != null && Math.abs(checks.grossSalary - slip.gross) > 0.005) {
      slip.issues.push({ level: 'warning', message: 'Gross salary in the file (' + formatMoney(checks.grossSalary) + ') does not match the sum of earnings (' + formatMoney(slip.gross) + ').' });
    }
    if (checks.totalDeductions != null && Math.abs(checks.totalDeductions - slip.totalDeductions) > 0.005) {
      slip.issues.push({ level: 'warning', message: 'Total deductions in the file (' + formatMoney(checks.totalDeductions) + ') does not match the sum of deductions (' + formatMoney(slip.totalDeductions) + ').' });
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
    if (slip.fields.absentDays != null && slip.fields.workingDays != null && slip.fields.absentDays !== '' && slip.fields.workingDays !== '' && slip.fields.absentDays > slip.fields.workingDays) {
      slip.issues.push({ level: 'warning', message: 'Absent days is greater than working days.' });
    }
    var textFields = ['name', 'jobTitle', 'department', 'nationality', 'bankName', 'remarks'];
    var nonLatin = textFields.filter(function (k) { return hasNonLatin(slip.fields[k]); });
    slip.earnings.concat(slip.deductions).forEach(function (l) {
      if (hasNonLatin(l.label) && nonLatin.indexOf(l.label) < 0) nonLatin.push(l.label);
    });
    if (nonLatin.length) {
      slip.issues.push({ level: 'warning', message: 'Some text (e.g. ' + fieldLabel(nonLatin[0]) + ') contains characters the PDF fonts cannot print; only English and Arabic are supported. They will show as "?".' });
    }

    slip.hasErrors = slip.issues.some(function (i) { return i.level === 'error'; });
    return slip;
  }

  function fieldLabel(key) {
    return COLUMN_BY_KEY[key] ? COLUMN_BY_KEY[key].label : key;
  }

  function isTotalsRow(raw, mapping) {
    var textCells = [];
    mapping.forEach(function (m, idx) {
      if (m && m.kind === 'column' && !m.def.type && !isBlank(raw[idx])) textCells.push(cleanText(raw[idx]));
    });
    // also look at the first few cells for files without name mapping
    for (var i = 0; i < Math.min(raw.length, 3); i++) if (!isBlank(raw[i]) && typeof raw[i] === 'string') textCells.push(cleanText(raw[i]));
    return textCells.some(function (t) { return TOTAL_ROW_RE.test(t); });
  }

  /**
   * Parse a workbook (already read with XLSX.read(..., {cellDates:true})).
   * options.overrides: { [normalizedHeader]: target } manual column mapping
   * Returns { settings, settingsFound, slips, issues, stats, columns, sheetName, headerRow, missingRequired }
   */
  function parseWorkbook(XLSX, wb, options) {
    XLSXRef = XLSX;
    options = options || {};
    var issues = [];
    var result = {};
    var settings = parseSettings(findSheet(wb, SHEET_SETTINGS), issues, result);
    var slips = [];
    var columns = [];
    var seen = {};
    var located = locateHeader(XLSX, wb, issues);
    var skipped = { totals: 0, unlabelled: 0, example: 0 };

    if (located) {
      var aoa = located.aoa;
      var headerIdx = located.headerIdx;
      var headerRow = aoa[headerIdx] || [];
      var sampleRows = aoa.slice(headerIdx + 1, headerIdx + 21);
      var hints = groupRowHints(aoa, headerIdx);
      var mapped = mapHeaders(headerRow, hints, options.overrides, issues, sampleRows);
      var mapping = mapped.mapping;
      columns = mapped.columns;
      seen = mapped.seen;

      var hasIdCol = !!seen.employeeId;
      var hasEarningCol = columns.some(function (c) { return (c.kind === 'column' && c.def.group === 'earning') || c.kind === 'earning'; });
      if (seen.name && hasEarningCol) {
        // manual mapping rescued an unrecognisable header row
        issues = issues.filter(function (i) { return i.code !== 'no-header'; });
      }

      if (!seen.name) issues.push({ level: 'error', where: 'Payslips', code: 'missing-name', message: 'No "Employee Name" column was found. Map one of your columns to Employee Name under "Columns".' });
      if (!hasEarningCol) issues.push({ level: 'error', where: 'Payslips', code: 'missing-earnings', message: 'No salary column was found (Basic Salary or any earning). Map one of your columns under "Columns".' });
      else if (!seen.basicSalary) issues.push({ level: 'warning', where: 'Payslips', code: 'missing-basic', message: 'No "Basic Salary" column was found. Earnings are taken from the other pay columns.' });
      if (!hasIdCol) issues.push({ level: 'warning', where: 'Payslips', code: 'missing-id', message: 'No "Employee ID" column was found. IDs are generated from the row order (EMP-001, EMP-002…). Map a column to Employee ID if you have one.' });

      var seq = 0;
      for (var r = headerIdx + 1; r < aoa.length; r++) {
        var raw = aoa[r] || [];
        if (raw.every(isBlank)) continue;
        if (isTotalsRow(raw, mapping)) { skipped.totals++; continue; }
        var slip = parseRow(raw, mapping, r + 1);
        if (slip.isExample) { skipped.example++; continue; }
        if ((isBlank(slip.name) && isBlank(slip.employeeId)) || (isBlank(slip.name) && slip.gross === 0 && slip.totalDeductions === 0)) {
          // amounts without any employee, or a label without any amounts: a totals/notes row
          skipped.unlabelled++;
          continue;
        }
        seq++;
        if (!hasIdCol || isBlank(slip.employeeId)) {
          slip.employeeId = 'EMP-' + String(seq).padStart(3, '0');
          slip.generatedId = true;
          if (hasIdCol) slip.issues.push({ level: 'warning', message: 'Employee ID is empty; "' + slip.employeeId + '" was generated.' });
        }
        slips.push(slip);
      }
      if (skipped.totals || skipped.unlabelled) {
        issues.push({ level: 'info', where: 'Payslips', message: (skipped.totals + skipped.unlabelled) + ' row' + (skipped.totals + skipped.unlabelled === 1 ? '' : 's') + ' without an employee (totals or notes) were skipped.' });
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
        issues.push({ level: 'error', where: 'Payslips', code: 'no-rows', message: 'No employee rows were found under the header row.' });
      }
    }

    var stats = {
      total: slips.length,
      withErrors: slips.filter(function (s) { return s.hasErrors; }).length,
      withWarnings: slips.filter(function (s) { return !s.hasErrors && s.issues.length; }).length,
      fileErrors: issues.filter(function (i) { return i.level === 'error'; }).length,
      fileWarnings: issues.filter(function (i) { return i.level === 'warning'; }).length,
      totalNet: round2(slips.reduce(function (a, s) { return a + (s.hasErrors ? 0 : s.net); }, 0)),
      totalGross: round2(slips.reduce(function (a, s) { return a + (s.hasErrors ? 0 : s.gross); }, 0)),
      skipped: skipped
    };
    stats.ok = stats.total - stats.withErrors;

    return {
      settings: settings,
      settingsFound: !!result.settingsFound,
      slips: slips,
      issues: issues,
      stats: stats,
      columns: columns,
      sheetName: located ? located.name : null,
      headerRow: located ? located.headerIdx + 1 : null,
      missingRequired: ['name', 'basicSalary', 'employeeId'].filter(function (k) { return !seen[k]; })
    };
  }

  // Choices offered by the column-mapping UI.
  function mappingChoices() {
    var list = [{ value: 'ignored', label: 'Ignore this column' }];
    var groups = { info: 'Employee details', attendance: 'Attendance', earning: 'Earnings', deduction: 'Deductions', check: 'Cross-check only' };
    Object.keys(groups).forEach(function (g) {
      COLUMNS.filter(function (c) { return c.group === g; }).forEach(function (c) {
        list.push({ value: 'column:' + c.key, label: c.label, group: groups[g] });
      });
    });
    list.push({ value: 'earning', label: 'Extra earning (named after the column)', group: 'Custom' });
    list.push({ value: 'deduction', label: 'Extra deduction (named after the column)', group: 'Custom' });
    return list;
  }

  function targetToChoice(m) {
    if (m.kind === 'column') return 'column:' + m.def.key;
    return m.kind;
  }

  function choiceToTarget(choice, header) {
    if (choice === 'ignored') return { kind: 'ignored' };
    if (choice === 'earning' || choice === 'deduction') return { kind: choice, label: header };
    if (choice.indexOf('column:') === 0) return { kind: 'column', key: choice.slice(7) };
    return { kind: 'ignored' };
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
    templateColumns: templateColumns,
    buildTemplateWorkbook: buildTemplateWorkbook,
    parseWorkbook: parseWorkbook,
    guessHeader: guessHeader,
    mappingChoices: mappingChoices,
    targetToChoice: targetToChoice,
    choiceToTarget: choiceToTarget,
    describeTarget: describeTarget,
    sampleRows: sampleRows,
    formatMoney: formatMoney,
    formatDate: formatDate,
    amountToWords: amountToWords,
    parseNumber: parseNumber,
    parseDateLike: parseDateLike,
    normalizeHeader: normalizeHeader,
    hasNonLatin: hasNonLatin,
    round2: round2,
    emptySettings: emptySettings
  };
});
