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
    { key: 'companyName', label: 'Company Name', required: true, example: 'Example Contracting Co. Ltd.', note: 'Required. Legal name as it appears on the CR.', aliases: ['company', 'employer', 'employer name', 'establishment name', 'establishment', 'اسم الشركة', 'الشركة', 'اسم المنشأة', 'المنشأة'] },
    { key: 'crNumber', label: 'CR Number', required: true, example: '1010123456', note: 'Required. Commercial Registration number.', aliases: ['cr', 'cr no', 'commercial registration', 'commercial registration number', 'cr number', 'السجل التجاري', 'رقم السجل', 'رقم السجل التجاري'] },
    { key: 'address', label: 'Company Address', example: 'Tabuk, Kingdom of Saudi Arabia', note: 'Optional.', aliases: ['address', 'العنوان'] },
    { key: 'phone', label: 'Company Phone', example: '+966 11 000 0000', note: 'Optional.', aliases: ['phone', 'telephone', 'mobile', 'الهاتف', 'رقم الهاتف'] },
    { key: 'email', label: 'Company Email', example: 'hr@example.com', note: 'Optional.', aliases: ['email', 'e-mail'] },
    { key: 'currency', label: 'Currency', example: 'SAR', default: 'SAR', note: 'Optional. 3-letter code, defaults to SAR.', aliases: ['currency code', 'العملة'] },
    { key: 'payPeriod', label: 'Pay Period', required: true, example: 'October 2026', note: 'Required. The month shown on every payslip, e.g. "October 2026".', aliases: ['period', 'salary month', 'month', 'pay month', 'الشهر', 'فترة الراتب', 'شهر الراتب'] },
    { key: 'periodStart', label: 'Period Start Date', example: '2026-10-01', type: 'date', note: 'Optional. First day of the pay period.', aliases: ['start date', 'from', 'period from'] },
    { key: 'periodEnd', label: 'Period End Date', example: '2026-10-31', type: 'date', note: 'Optional. Last day of the pay period.', aliases: ['end date', 'to', 'period to'] },
    { key: 'paymentDate', label: 'Payment Date', example: '2026-11-05', type: 'date', note: 'Optional. Date salaries were transferred.', aliases: ['pay date', 'date of payment', 'transfer date', 'تاريخ الصرف', 'تاريخ الدفع'] },
    { key: 'preparedBy', label: 'Prepared By', example: 'HR Department', note: 'Optional. Printed in the footer.', aliases: ['prepared', 'issued by'] },
    { key: 'footerNote', label: 'Footer Note', example: 'This is a system-generated payslip and does not require a signature.', note: 'Optional. Free text printed at the bottom of each payslip.', aliases: ['note', 'footer', 'disclaimer'] },
    { key: 'countryCode', label: 'Default Country Code', example: '966', default: '966', note: 'Optional. Added to mobile numbers written without a country code (966 = Saudi Arabia).', aliases: ['country code', 'phone country code', 'dial code'] },
    { key: 'senderEmail', label: 'Sender Email', example: 'payroll@example.com', note: 'Optional. The From address used in Outlook drafts and mail-merge files.', aliases: ['from email', 'from', 'sender', 'payroll email'] },
    { key: 'messageSubject', label: 'Message Subject', example: 'Payslip {period} - {company}', default: 'Payslip {period} - {company}', note: 'Optional. Subject line for emails and first line of WhatsApp messages. Placeholders: {name} {period} {company} {net} {password_hint}', aliases: ['subject', 'email subject'] },
    { key: 'messageText', label: 'Message Text', multiline: true, example: '', default: 'Dear {first_name},\n\nPlease find attached your payslip for {period}.\n\n{password_hint}\n\nBest regards', note: 'Optional. Body of the email / WhatsApp message. Placeholders: {first_name} {name} {period} {company} {net} {password_hint} {password_hint_ar}', aliases: ['message', 'email body', 'body', 'message body'] },
    { key: 'passwordRule', label: 'Password Rule', example: 'id', default: 'id', note: 'Optional. PDF password: id (National ID / Iqama number, default), id+dob (ID + birth day and month DDMM), dob (DDMMYYYY), custom (uses the "Payslip Password" column), none.', aliases: ['pdf password', 'password', 'password scheme'] },
    { key: 'logo', label: 'Logo', type: 'image', example: '', note: 'Optional. Insert your logo picture anywhere on this sheet (Insert > Pictures). The first picture on this sheet is used as the logo.', aliases: ['company logo'] },
    { key: 'stamp', label: 'Stamp / Signature', type: 'image', example: '', note: 'Optional. The second picture on this sheet is used as the stamp or signature.', aliases: ['stamp', 'signature'] }
  ];

  var COLUMNS = [
    { key: 'employeeId', label: 'Employee ID', group: 'info', required: true, example: 'EMP-001', aliases: ['emp id', 'employee no', 'employee number', 'emp no', 'id', 'staff id', 'staff no', 'employee code', 'emp code', 'code', 'badge no', 'badge', 'الرقم الوظيفي', 'رقم الموظف', 'الرقم', 'رقم وظيفي', 'كود الموظف'] },
    { key: 'name', label: 'Employee Name', group: 'info', required: true, example: 'Ahmed Ali Hassan', aliases: ['name', 'full name', 'employee', 'staff name', 'worker name', 'name of employee', 'الاسم', 'اسم الموظف', 'الإسم', 'اسم العامل', 'الاسم الكامل'] },
    { key: 'nationalId', label: 'National ID / Iqama', group: 'info', example: '2456789012', aliases: ['iqama', 'iqama no', 'iqama number', 'national id', 'id number', 'national id iqama', 'iqama id', 'id no', 'identity number', 'resident id', 'nid', 'رقم الهوية', 'رقم الإقامة', 'الهوية', 'الإقامة', 'رقم الاقامة', 'رقم الهويه', 'السجل المدني', 'هوية / إقامة'] },
    { key: 'nationality', label: 'Nationality', group: 'info', example: 'Egyptian', aliases: ['nation', 'country', 'الجنسية'] },
    { key: 'dob', label: 'Date of Birth', group: 'info', type: 'date', example: '1990-03-15', aliases: ['dob', 'birth date', 'birthdate', 'date of birth', 'birthday', 'تاريخ الميلاد'] },
    { key: 'jobTitle', label: 'Job Title', group: 'info', example: 'Site Engineer', aliases: ['title', 'position', 'designation', 'job', 'profession', 'occupation', 'role', 'المهنة', 'الوظيفة', 'المسمى الوظيفي', 'المهنه'] },
    { key: 'department', label: 'Department', group: 'info', example: 'Projects', aliases: ['dept', 'section', 'project', 'division', 'site', 'location', 'القسم', 'الإدارة', 'المشروع', 'الموقع'] },
    { key: 'joiningDate', label: 'Joining Date', group: 'info', type: 'date', example: '2024-03-15', aliases: ['date of joining', 'doj', 'hire date', 'start date', 'join date', 'date joined', 'hiring date', 'تاريخ المباشرة', 'تاريخ الالتحاق', 'تاريخ التعيين', 'تاريخ الانضمام'] },
    { key: 'mobile', label: 'Mobile', group: 'contact', type: 'phone', example: '0551234567', aliases: ['mobile no', 'mobile number', 'phone', 'phone no', 'phone number', 'contact', 'contact no', 'contact number', 'whatsapp', 'whatsapp no', 'whatsapp number', 'cell', 'cell phone', 'tel', 'telephone', 'gsm', 'جوال', 'رقم الجوال', 'هاتف', 'واتساب'] },
    { key: 'email', label: 'Email', group: 'contact', type: 'email', example: 'ahmed@example.com', aliases: ['e-mail', 'mail', 'email address', 'e mail', 'البريد', 'البريد الإلكتروني', 'ايميل'] },
    { key: 'bankName', label: 'Bank Name', group: 'info', example: 'Al Rajhi Bank', aliases: ['bank', 'البنك', 'اسم البنك'] },
    { key: 'iban', label: 'IBAN', group: 'info', example: 'SA0380000000608010167519', aliases: ['iban no', 'iban number', 'account', 'account number', 'bank account', 'account no', 'acc no', 'الآيبان', 'رقم الحساب', 'الايبان', 'آيبان', 'رقم الآيبان'] },
    { key: 'workingDays', label: 'Working Days', group: 'attendance', type: 'number', example: 30, aliases: ['days worked', 'work days', 'paid days', 'days', 'present days', 'no of days', 'number of days', 'أيام العمل', 'عدد الأيام', 'ايام العمل', 'أيام الحضور'] },
    { key: 'absentDays', label: 'Absent Days', group: 'attendance', type: 'number', example: 0, aliases: ['absence days', 'absent', 'unpaid days', 'lop days', 'leave days', 'absences', 'أيام الغياب', 'ايام الغياب'] },
    { key: 'overtimeHours', label: 'Overtime Hours', group: 'attendance', type: 'number', example: 10, aliases: ['ot hours', 'overtime hrs', 'ot hrs', 'ot hr', 'overtime hour', 'ساعات إضافية', 'ساعات العمل الإضافي', 'ساعات الاضافي'] },
    { key: 'basicSalary', label: 'Basic Salary', group: 'earning', required: true, type: 'money', example: 4000, aliases: ['basic', 'basic pay', 'base salary', 'basic wage', 'salary', 'wage', 'base pay', 'monthly salary', 'basic amount', 'الراتب الأساسي', 'الأساسي', 'الراتب', 'الراتب الاساسي', 'الاساسي'] },
    { key: 'housing', label: 'Housing Allowance', group: 'earning', type: 'money', example: 1000, aliases: ['housing', 'hra', 'house allowance', 'accommodation allowance', 'accommodation', 'house rent', 'rent allowance', 'housing all', 'بدل السكن', 'السكن'] },
    { key: 'transport', label: 'Transport Allowance', group: 'earning', type: 'money', example: 400, aliases: ['transport', 'transportation', 'transportation allowance', 'conveyance', 'travel allowance', 'transport all', 'car allowance', 'بدل النقل', 'بدل المواصلات', 'النقل', 'المواصلات'] },
    { key: 'food', label: 'Food Allowance', group: 'earning', type: 'money', example: 300, aliases: ['food', 'meal allowance', 'meals', 'food all', 'بدل الطعام', 'بدل الإعاشة', 'الإعاشة', 'بدل الاعاشة'] },
    { key: 'overtimePay', label: 'Overtime Pay', group: 'earning', type: 'money', example: 250, aliases: ['overtime', 'ot', 'ot pay', 'overtime amount', 'ot amount', 'overtime allowance', 'العمل الإضافي', 'الإضافي', 'أجر إضافي', 'الاضافي', 'بدل العمل الإضافي'] },
    { key: 'otherEarnings', label: 'Other Earnings', group: 'earning', type: 'money', example: 0, aliases: ['other allowance', 'other allowances', 'bonus', 'other earning', 'additions', 'other', 'others', 'incentive', 'incentives', 'commission', 'other income', 'other pay', 'other addition', 'بدلات أخرى', 'مكافأة', 'حوافز', 'بدلات اخرى'] },
    { key: 'gosi', label: 'GOSI Deduction', group: 'deduction', type: 'money', example: 0, aliases: ['gosi', 'social insurance', 'gosi employee', 'gosi ded', 'pension', 'التأمينات', 'التأمينات الاجتماعية', 'خصم التأمينات', 'التامينات'] },
    { key: 'absence', label: 'Absence Deduction', group: 'deduction', type: 'money', example: 0, aliases: ['absence', 'absent deduction', 'lop', 'unpaid leave', 'absence amount', 'absent amount', 'leave deduction', 'خصم الغياب', 'الغياب'] },
    { key: 'loan', label: 'Loan / Advance', group: 'deduction', type: 'money', example: 0, aliases: ['loan', 'advance', 'salary advance', 'loan deduction', 'loan advance', 'advance deduction', 'loan ded', 'سلفة', 'السلف', 'قرض', 'السلفة'] },
    { key: 'otherDeductions', label: 'Other Deductions', group: 'deduction', type: 'money', example: 0, aliases: ['other deduction', 'deductions', 'deduction', 'penalty', 'penalties', 'fines', 'fine', 'other ded', 'violation', 'violations', 'misc deduction', 'خصومات أخرى', 'الخصومات', 'جزاءات', 'خصم', 'خصومات اخرى'] },
    { key: 'netSalary', label: 'Net Salary', group: 'check', type: 'money', example: '', aliases: ['net', 'net pay', 'net amount', 'total net', 'net paid', 'take home', 'net payable', 'total payable', 'amount payable', 'payable', 'الصافي', 'صافي الراتب', 'الراتب الصافي', 'المستحق'] },
    { key: 'grossSalary', label: 'Gross Salary', group: 'check', type: 'money', example: '', inTemplate: false, aliases: ['gross', 'gross pay', 'total earnings', 'total earning', 'total salary', 'gross amount', 'total allowances', 'total', 'الإجمالي', 'إجمالي الراتب', 'الاجمالي', 'اجمالي الراتب'] },
    { key: 'totalDeductions', label: 'Total Deductions', group: 'check', type: 'money', example: '', inTemplate: false, aliases: ['total deduction', 'total ded', 'deductions total', 'sum of deductions', 'إجمالي الخصومات', 'اجمالي الخصومات'] },
    { key: 'remarks', label: 'Remarks', group: 'info', example: '', aliases: ['notes', 'comment', 'comments', 'remark', 'description', 'ملاحظات', 'ملاحظة'] },
    { key: 'payslipPassword', label: 'Payslip Password', group: 'contact', inTemplate: false, example: '', aliases: ['pdf password', 'password', 'passcode'] }
  ];

  // Column headers like "Earning: Site Allowance" / "Deduction: Housing Loan" are
  // treated as extra earning / deduction lines without changing the app.
  var CUSTOM_EARNING_RE = /^(earning|earnings|allowance|addition|add|income|pay)\s*[:\-]\s*(.+)$/i;
  var CUSTOM_DEDUCTION_RE = /^(deduction|deductions|ded|less|minus)\s*[:\-]\s*(.+)$/i;

  // Keyword rules used when a header does not match a known label/alias exactly.
  // Order matters: the first matching rule wins.
  var GUESS_RULES = [
    [/\b(emp|employee|staff|worker|personnel|payroll)\b.*\b(id|no|num|number|code)\b|\b(id|no|num|number|code)\b.*\b(emp|employee|staff|worker)\b|^(s|sl|sr|si)\s*no$|^(sno|slno|srno)$|^serial|^#$|^no$/, 'employeeId'],
    [/iqama|national\s*id|\bnid\b|identity|resident\s*id|\bid\s*(no|num|number)\b|civil\s*id|passport|اقامه|هويه|سجل مدني/, 'nationalId'],
    [/(first|given|middle|last|family|father|second)\s*name|surname|الاسم الاول|اسم العائله|اسم الاب|اللقب/, 'name'],
    [/\bname\b|اسم/, 'name'],
    [/nationality|citizen/, 'nationality'],
    [/title|position|designation|profession|occupation|\bjob\b|\brole\b|trade/, 'jobTitle'],
    [/depart|dept|section|division|\bunit\b|cost\s*cent|\bproject\b|\bsite\b/, 'department'],
    [/\bdob\b|birth/, 'dob'],
    [/e\s*-?\s*mail|\bmail\b|بريد|ايميل|إيميل/, 'email'],
    [/mobile|phone|whats\s*app|\bcell\b|\btel\b|\bgsm\b|contact\s*(no|num|number)?$|جوال|هاتف|واتس/, 'mobile'],
    [/joining|hire|hiring|\bdoj\b|date\s*of\s*join|join\s*date|start\s*date|contract\s*start/, 'joiningDate'],
    [/\biban\b|account\s*(no|num|number)?|\bacc\b|\ba\s*c\b/, 'iban'],
    [/\bbank\b/, 'bankName'],
    [/(working|work|paid|present|actual|attendance|no\s*of)\s*days|days\s*(worked|present|paid)/, 'workingDays'],
    [/(absen|unpaid|lop|leave)\w*\s*days|days\s*absen/, 'absentDays'],
    [/(overtime|\bot\b)\s*\(?\s*(hour|hrs|hr)|(hour|hrs)\s*(of\s*)?(overtime|\bot\b)/, 'overtimeHours'],
    [/overtime|\bot\b/, 'overtimePay'],
    [/\bnet\b|take\s*home|payable|to\s*be\s*paid|paid\s*amount|final\s*salary|total\s*pay\b|صافي|المستحق/, 'netSalary'],
    [/total\s*ded|ded\w*\s*total|sum\s*of\s*ded/, 'totalDeductions'],
    [/gross|total\s*earn|total\s*allow|total\s*salary|total\s*income|^total$/, 'grossSalary'],
    [/basic|\bbase\b|^salary$|^wage|اساسي/, 'basicSalary'],
    [/hous|\bhra\b|accommodat|\brent\b/, 'housing'],
    [/transport|conveyance|travel|\bcar\b|fuel|petrol/, 'transport'],
    [/food|meal|canteen/, 'food'],
    [/gosi|social\s*ins|pension|\bsi\b/, 'gosi'],
    [/absen|\blop\b|unpaid|leave\s*ded/, 'absence'],
    [/loan|advance|\badv\b/, 'loan'],
    [/remark|note|comment|descr/, 'remarks'],
    // generic buckets -> custom lines named after the header
    [/deduct|penalt|fine|violation|medical|insurance|tax|recover|damage|\bless\b|\bded\b|خصم|جزاء|غرامه|سلف|استقطاع/, '*deduction'],
    [/allow|bonus|incentive|commission|reward|arrear|reimburs|ticket|mobile|phone|shift|hardship|danger|special|site\s*all|other\s*(earn|pay|income|add)|بدل|مكافا|حافز|علاوه/, '*earning']
  ];

  var EXAMPLE_ID_PREFIX = 'EXAMPLE';
  var TOTAL_ROW_RE = /^(grand\s*)?(total|totals|sum|subtotal|sub\s*total|net\s*total|g\.?\s*total)\b/i;

  // ---------------------------------------------------------------------------
  // Small helpers
  // ---------------------------------------------------------------------------

  function normalizeHeader(s) {
    return toAsciiDigits(String(s == null ? '' : s))
      .replace(/[\u064B-\u0652\u0640]/g, '') // Arabic diacritics and tatweel
      .replace(/[أإآ]/g, 'ا')
      .replace(/ة/g, 'ه')
      .replace(/ى/g, 'ي')
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

  // Arabic-Indic (٠-٩) and Persian (۰-۹) digits, Arabic decimal/thousands marks -> ASCII.
  function toAsciiDigits(s) {
    return String(s)
      .replace(/[\u0660-\u0669]/g, function (d) { return String(d.charCodeAt(0) - 0x0660); })
      .replace(/[\u06F0-\u06F9]/g, function (d) { return String(d.charCodeAt(0) - 0x06F0); })
      .replace(/\u066B/g, '.')
      .replace(/\u066C/g, ',');
  }

  var EXCEL_ERROR_RE = /^#(N\/A|REF!|VALUE!|DIV\/0!|NAME\?|NUM!|NULL!|SPILL!|CALC!|GETTING_DATA|ERROR!?)/i;

  // "3.000,50" or "3000,50" style (comma as decimal separator). "3.000" alone is ambiguous and is decided per column.
  function looksDecimalComma(cleaned) {
    return /^\d{1,3}(\.\d{3})+,\d+$/.test(cleaned) || /^\d+,\d{1,2}$/.test(cleaned);
  }

  function isBlank(v) {
    return v == null || (typeof v === 'string' && v.trim() === '');
  }

  function cleanText(v) {
    if (v == null) return '';
    if (v instanceof Date) return formatDate(v);
    return String(v).replace(/\s+/g, ' ').trim();
  }

  // Returns { value: Number|null, error: String|null }. opts.decimalComma: the column uses "," as decimal separator.
  function parseNumber(v, opts) {
    opts = opts || {};
    if (isBlank(v)) return { value: null, error: null };
    if (typeof v === 'number') {
      if (!isFinite(v)) return { value: null, error: 'is not a valid number' };
      return { value: v, error: null };
    }
    if (typeof v === 'boolean') return { value: null, error: 'is not a number' };
    var s = toAsciiDigits(String(v)).trim();
    var err = s.match(EXCEL_ERROR_RE);
    if (err) return { value: null, error: 'contains the Excel error ' + err[0] + ' (fix the formula in Excel)' };
    if (/^[-–—]+$/.test(s) || /^(n\/?a|nil|none|null)$/i.test(s)) return { value: null, error: null };
    // strip currency codes/symbols and whitespace
    var cleaned = s.replace(/[^0-9.,\-()]/g, '');
    var negative = /^\(.*\)$/.test(cleaned) || /^-/.test(cleaned) || /\d-$/.test(cleaned);
    cleaned = cleaned.replace(/[()\-]/g, '');
    if (!/\d/.test(cleaned)) return { value: null, error: 'is not a number ("' + s + '")' };
    if (opts.decimalComma || looksDecimalComma(cleaned)) {
      cleaned = cleaned.replace(/\./g, '');
      if ((cleaned.match(/,/g) || []).length > 1) return { value: null, error: 'is not a number ("' + s + '")' };
      cleaned = cleaned.replace(',', '.');
    } else {
      cleaned = cleaned.replace(/,/g, '');
      // "1.234.567" with several dots: keep the last one as decimal point
      if ((cleaned.match(/\./g) || []).length > 1) {
        var last = cleaned.lastIndexOf('.');
        cleaned = cleaned.slice(0, last).replace(/\./g, '') + cleaned.slice(last);
      }
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

  function buildDate(y, mo, d, text) {
    var dt = new Date(y, mo - 1, d);
    if (dt.getFullYear() !== y || dt.getMonth() !== mo - 1 || dt.getDate() !== d) return { date: null, text: text, invalid: true };
    return { date: dt, text: formatDate(dt) };
  }

  /**
   * Reads a date from a Date, an Excel serial or text. order: 'dmy' (default, Saudi convention) or 'mdy'.
   * A value whose second part is above 12 (03/15/2024) is read as month/day whatever the order.
   * Returns { date, text, invalid?, ambiguous? }.
   */
  function readDate(v, order) {
    if (isBlank(v)) return { date: null, text: '' };
    if (v instanceof Date) return isNaN(v.getTime()) ? { date: null, text: '', invalid: true } : { date: v, text: formatDate(v) };
    if (typeof v === 'number') {
      if (v > 20000 && v < 80000) {
        var sd = excelSerialToDate(v);
        return { date: sd, text: formatDate(sd) };
      }
      return { date: null, text: String(v) };
    }
    var s = toAsciiDigits(String(v)).trim();
    var m = s.match(/^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})(?:$|[T\s])/);
    if (m) return buildDate(+m[1], +m[2], +m[3], s);
    m = s.match(/^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{4}|\d{2})(?:$|\s)/);
    if (m) {
      var a = +m[1];
      var b = +m[2];
      var y = +m[3];
      if (m[3].length === 2) y += y > (new Date().getFullYear() % 100) + 5 ? 1900 : 2000;
      var mdy = (b > 12 && a <= 12) || (order === 'mdy' && a <= 12);
      var res = buildDate(y, mdy ? a : b, mdy ? b : a, s);
      res.ambiguous = a <= 12 && b <= 12 && a !== b;
      return res;
    }
    if (/[A-Za-z]/.test(s)) {
      var d = new Date(s.replace(/-/g, ' '));
      if (!isNaN(d.getTime())) return { date: d, text: formatDate(d) };
    }
    return { date: null, text: s, invalid: /\d/.test(s) };
  }

  // Display string for a date-like value (free text is passed through unchanged).
  function parseDateLike(v, order) {
    return readDate(v, order).text;
  }

  // Returns a Date for Date objects, Excel serials and common string formats, else null.
  function toDate(v, order) {
    return readDate(v, order).date || null;
  }

  function isoDate(d) {
    return d ? d.getFullYear() + '-' + pad2(d.getMonth() + 1) + '-' + pad2(d.getDate()) : '';
  }

  // Mobile numbers -> international digits without "+" (what wa.me expects), e.g. 9665XXXXXXXX.
  function normalizePhone(v, countryCode) {
    if (isBlank(v)) return { value: '', error: null };
    var raw = typeof v === 'number' ? String(Math.round(v)) : toAsciiDigits(String(v)).trim();
    // Excel may turn 0551234567 into 551234567 (number) or 5.51234567E8
    if (/\d[eE][+-]?\d/.test(raw)) raw = String(Number(raw));
    var hasPlus = /^\s*\+/.test(raw);
    var digits = raw.replace(/\D/g, '');
    if (!digits) return { value: '', error: 'is not a phone number ("' + raw + '")' };
    var cc = String(countryCode || '966').replace(/\D/g, '') || '966';
    var out;
    if (hasPlus) out = digits;
    else if (digits.indexOf('00') === 0) out = digits.slice(2);
    else if (digits.indexOf(cc) === 0 && digits.length >= cc.length + 8) out = digits;
    else if (digits.charAt(0) === '0') out = cc + digits.slice(1);
    else if (digits.length >= 8 && digits.length <= 10) out = cc + digits;
    else out = digits;
    if (out.length < 9 || out.length > 15) return { value: out, error: 'does not look like a valid mobile number ("' + raw + '")' };
    if (out.indexOf('966') === 0 && !(out.length === 12 && out.charAt(3) === '5')) {
      return { value: out, error: 'does not look like a valid Saudi mobile number ("' + raw + '", expected 05XXXXXXXX)' };
    }
    return { value: out, error: null };
  }

  function normalizeEmail(v) {
    if (isBlank(v)) return { value: '', error: null };
    var s = String(v).trim().toLowerCase().replace(/^mailto:/, '');
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(s)) return { value: s, error: 'is not a valid email address ("' + String(v).trim() + '")' };
    return { value: s, error: null };
  }

  // "October 2026" / "Oct-2026" / "10/2026" / "2026-10" -> "2026-10"; otherwise a slug of the text.
  var MONTH_INDEX = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, sept: 9, oct: 10, nov: 11, dec: 12 };
  var AR_MONTHS = [['يناير', 1], ['فبراير', 2], ['مارس', 3], ['ابريل', 4], ['أبريل', 4], ['مايو', 5], ['يونيو', 6], ['يوليو', 7], ['اغسطس', 8], ['أغسطس', 8], ['سبتمبر', 9], ['اكتوبر', 10], ['أكتوبر', 10], ['نوفمبر', 11], ['ديسمبر', 12]];

  // Month (and year when present) in a sheet name or pay period: "Oct 2026", "October", "10-2026", "أكتوبر 2026".
  function monthInfo(text) {
    var s = toAsciiDigits(String(text || '')).trim();
    if (!s) return null;
    var yearM = s.match(/(19|20)\d{2}/);
    var year = yearM ? +yearM[0] : null;
    var low = s.toLowerCase();
    var words = low.match(/[a-z]{3,9}/g) || [];
    for (var i = 0; i < words.length; i++) {
      var w = words[i];
      var idx = MONTH_INDEX[w.slice(0, 4)] || MONTH_INDEX[w.slice(0, 3)];
      if (idx && /^(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)/.test(w)) {
        if (!year) { var y2 = low.match(new RegExp(w + "\\W*'?(\\d{2})(?!\\d)")); if (y2) year = 2000 + +y2[1]; }
        return { year: year, month: idx };
      }
    }
    for (var j = 0; j < AR_MONTHS.length; j++) if (s.indexOf(AR_MONTHS[j][0]) >= 0) return { year: year, month: AR_MONTHS[j][1] };
    var m = s.match(/^(\d{1,2})[-\/.](\d{4})$/) || null;
    if (m && +m[1] >= 1 && +m[1] <= 12) return { year: +m[2], month: +m[1] };
    m = s.match(/^(\d{4})[-\/.](\d{1,2})$/);
    if (m && +m[2] >= 1 && +m[2] <= 12) return { year: +m[1], month: +m[2] };
    return null;
  }
  function periodSlug(payPeriod, periodStart) {
    var s = String(payPeriod || '').trim();
    var m = s.match(/([A-Za-z]{3,9})\.?\s*[-\/,]?\s*(\d{4})/) || s.match(/(\d{4})\s*[-\/,]?\s*([A-Za-z]{3,9})/);
    if (m) {
      var name = (isNaN(+m[1]) ? m[1] : m[2]).slice(0, 4).toLowerCase().replace(/\.$/, '');
      var year = isNaN(+m[1]) ? m[2] : m[1];
      var mi = MONTH_INDEX[name] || MONTH_INDEX[name.slice(0, 3)];
      if (mi) return year + '-' + pad2(mi);
    }
    m = s.match(/^(\d{1,2})[-\/](\d{4})$/);
    if (m) return m[2] + '-' + pad2(+m[1]);
    m = s.match(/^(\d{4})[-\/](\d{1,2})$/);
    if (m) return m[1] + '-' + pad2(+m[2]);
    var d = toDate(periodStart);
    if (d) return d.getFullYear() + '-' + pad2(d.getMonth() + 1);
    return asciiSlug(s) || 'period';
  }

  function asciiSlug(s) {
    return String(s || '')
      .normalize('NFKD')
      .replace(/[\u0300-\u036f]/g, '')
      .replace(/[^A-Za-z0-9]+/g, '-')
      .replace(/^-+|-+$/g, '')
      .slice(0, 40);
  }

  // EMP-001_2026-10_Ahmed-Hassan.pdf : ASCII only, sorts by employee id, matches the send list.
  function slipFileName(slip, settings, ext) {
    var parts = [asciiSlug(slip.employeeId) || 'row' + slip.rowNumber, periodSlug(settings.payPeriod, settings.periodStart)];
    var nm = asciiSlug(slip.name);
    if (nm) parts.push(nm);
    return parts.join('_') + (ext || '.pdf');
  }

  // Builds the PDF password for a slip according to the rule. Returns { password, hint, hintAr, error }.
  function buildPassword(slip, rule) {
    rule = String(rule || 'id').toLowerCase().replace(/\s+/g, '');
    var id = String(slip.fields.nationalId || '').replace(/\s+/g, '') || String(slip.employeeId || '');
    var dob = toDate(slip.fields.dobDate || slip.fields.dob);
    var ddmm = dob ? pad2(dob.getDate()) + pad2(dob.getMonth() + 1) : '';
    var ddmmyyyy = dob ? ddmm + dob.getFullYear() : '';
    if (rule === 'none' || rule === 'no' || rule === 'off') return { password: '', hint: '', hintAr: '', error: null };
    if (rule === 'custom') {
      var pw = String(slip.fields.payslipPassword || '').trim();
      return pw ? { password: pw, hint: 'Password: given to you by HR.', hintAr: 'كلمة المرور: من الموارد البشرية.', error: null } : { password: '', hint: '', hintAr: '', error: 'no value in the "Payslip Password" column' };
    }
    if (rule === 'dob') {
      return ddmmyyyy ? { password: ddmmyyyy, hint: 'Password: your date of birth (DDMMYYYY).', hintAr: 'كلمة المرور: تاريخ ميلادك (يوم شهر سنة).', error: null } : { password: '', hint: '', hintAr: '', error: 'no date of birth' };
    }
    if (rule === 'id') {
      return id ? { password: id, hint: 'Password: your Iqama number.', hintAr: 'كلمة المرور: رقم الإقامة.', error: null } : { password: '', hint: '', hintAr: '', error: 'no National ID / Iqama' };
    }
    // default id+dob
    if (!id) return { password: '', hint: '', hintAr: '', error: 'no National ID / Iqama' };
    if (!ddmm) return { password: '', hint: '', hintAr: '', error: 'no date of birth' };
    return { password: id + ddmm, hint: 'Password: your Iqama number + birth day and month (DDMM).', hintAr: 'كلمة المرور: رقم الإقامة + يوم وشهر الميلاد.', error: null };
  }

  // {name} {period} {company} {net} {currency} {password_hint} {password_hint_ar} {employee_id}
  function renderTemplate(tpl, vars) {
    return String(tpl == null ? '' : tpl).replace(/\{(\w+)\}/g, function (m, k) {
      return vars[k] != null ? String(vars[k]) : '';
    });
  }

  function messageVars(slip, settings, pw) {
    return {
      name: slip.name,
      first_name: String(slip.name || '').trim().split(/\s+/)[0] || slip.name,
      employee_id: slip.employeeId,
      period: settings.payPeriod || '',
      company: settings.companyName || '',
      net: formatMoney(slip.net),
      currency: settings.currency || 'SAR',
      password_hint: pw && pw.hint ? pw.hint : '',
      password_hint_ar: pw && pw.hintAr ? pw.hintAr : ''
    };
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
      else if (withSample && f.example) value = f.example;
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
  var SKIP_SHEETS = [SHEET_SETTINGS, SHEET_INSTRUCTIONS, 'Lists', 'Lookup', 'Lookups'].map(normalizeHeader);
  // Helper columns of our own template (and similar) that carry no payslip data: ignored without a warning.
  var IGNORE_SILENT = ['check', 'checks', 'validation', 'row check', 'template check'].map(normalizeHeader);
  var NAME_PART_RE = /(first|given|middle|last|family|father|second)\s*name|surname|الاسم الاول|اسم العائله|اسم الاب|اللقب/;
  var ERROR_CODES = { 0: '#NULL!', 7: '#DIV/0!', 15: '#VALUE!', 23: '#REF!', 29: '#NAME?', 36: '#NUM!', 42: '#N/A', 43: '#GETTING_DATA' };
  var HEADER_SEARCH_ROWS = 200;

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
    exposeErrorCells(ws);
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
        else if (def.multiline) settings[def.key] = String(raw == null ? '' : raw).replace(/\r\n?/g, '\n').replace(/[ \t]+/g, ' ').replace(/ *\n */g, '\n').trim();
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

  // Excel error cells (#N/A, #REF!…) become visible text so they are reported instead of being read as empty.
  function exposeErrorCells(ws) {
    if (!ws || ws['!errorsExposed']) return;
    Object.keys(ws).forEach(function (k) {
      if (k.charAt(0) === '!') return;
      var c = ws[k];
      if (c && c.t === 'e') {
        c.t = 's';
        c.v = c.w || ERROR_CODES[c.v] || '#ERROR';
      }
      // text in a date-formatted column (a header, a note) would otherwise be turned into an invalid date
      if (c && c.t === 's' && c.z) delete c.z;
    });
    ws['!errorsExposed'] = true;
  }

  function isHiddenSheet(wb, name) {
    var list = wb.Workbook && wb.Workbook.Sheets;
    if (!list) return false;
    for (var i = 0; i < list.length; i++) if (list[i].name === name) return !!list[i].Hidden;
    return false;
  }

  // Reads one sheet: rows, the sheet row number of the first row, hidden rows, best header row.
  function analyzeSheet(XLSX, wb, name, forcedHeaderRow) {
    var ws = wb.Sheets[name];
    exposeErrorCells(ws);
    var aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: true });
    var range = ws['!ref'] ? XLSX.utils.decode_range(ws['!ref']) : { s: { r: 0, c: 0 } };
    var hidden = {};
    (ws['!rows'] || []).forEach(function (r, i) { if (r && r.hidden) hidden[i] = true; });
    var best = { idx: -1, score: 0, total: -1 };
    var limit = Math.min(aoa.length, HEADER_SEARCH_ROWS);
    for (var i = 0; i < limit; i++) {
      var sc = headerScore(aoa[i]);
      var total = sc + Math.min(aoa.length - i - 1, 5) * 0.01;
      if (total > best.total) best = { idx: i, score: sc, total: total };
    }
    var forced = forcedHeaderRow != null ? forcedHeaderRow - 1 - range.s.r : null;
    if (forced != null && forced >= 0 && forced < aoa.length) best = { idx: forced, score: Math.max(headerScore(aoa[forced]), 1.5), forced: true };
    if (best.idx < 0) {
      var first = 0;
      while (first < aoa.length && (aoa[first] || []).every(isBlank)) first++;
      best = { idx: first < aoa.length ? first : 0, score: 0 };
    }
    var rows = aoa.slice(best.idx + 1).filter(function (r) { return r && !r.every(isBlank); }).length;
    return {
      name: name, aoa: aoa, rowOffset: range.s.r, hidden: hidden,
      headerIdx: best.idx, headerRow: best.idx + range.s.r + 1, score: best.score, forced: !!best.forced,
      isData: best.score >= 1.5, rows: rows, month: monthInfo(name),
      preferred: normalizeHeader(name) === normalizeHeader(SHEET_PAYSLIPS)
    };
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
        if (/earn|allow|addition|income|pay\b|salary|استحقاق|بدلات|الراتب/.test(n)) { current = 'earning'; any = true; }
        else if (/deduct|less|minus|خصم|خصومات|استقطاع/.test(n)) { current = 'deduction'; any = true; }
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
  function mapHeaders(headerRow, groupHints, overrides, issues, sampleRows, prefix) {
    var mapping = [];
    var seen = {};
    var columns = [];
    overrides = overrides || {};
    prefix = prefix || '';

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
      } else if (IGNORE_SILENT.indexOf(norm) >= 0) {
        target = { kind: 'ignored' };
        source = 'silent';
      } else {
        target = guessHeader(text);
        source = target ? target.confidence : 'none';
        if (!target && groupHints[idx] && numericRatio > 0.5) {
          target = { kind: groupHints[idx], label: text };
          source = 'guess';
        }
        if (!target) target = { kind: 'ignored' };
        // Contact columns: check the values look like phones / emails even on an exact header match
        // ("Mobile" full of small numbers is a mobile allowance, not phone numbers).
        if (target.kind === 'column' && COLUMN_BY_KEY[target.key] && COLUMN_BY_KEY[target.key].type === 'phone' && samples.length >= 2 &&
            samples.filter(function (v) { return !normalizePhone(v).error; }).length / samples.length < 0.5) {
          target = groupHints[idx] ? { kind: groupHints[idx], label: text } : (numericRatio > 0.8 ? { kind: 'earning', label: text } : { kind: 'ignored' });
          source = target.kind === 'ignored' ? 'none' : 'guess';
        }
        if (target.kind === 'column' && COLUMN_BY_KEY[target.key] && COLUMN_BY_KEY[target.key].type === 'email' && samples.length >= 2 &&
            samples.filter(function (v) { return String(v).indexOf('@') >= 0; }).length / samples.length < 0.5) {
          target = { kind: 'ignored' };
          source = 'none';
        }
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
      var m;
      if (target.kind === 'column') {
        var def = COLUMN_BY_KEY[target.key];
        if (!def) {
          m = { kind: 'ignored', label: text };
        } else if (seen[def.key] && def.key === 'name' && (NAME_PART_RE.test(norm) || NAME_PART_RE.test(normalizeHeader(seen.name)))) {
          // "First Name" + "Last Name": joined into one name
          m = { kind: 'column', def: def, namePart: true };
          issues.push({ level: 'info', where: 'Payslips', code: 'name-joined', message: prefix + 'Name columns "' + seen.name + '" and "' + text + '" are joined into one employee name.' });
        } else if (seen[def.key]) {
          m = { kind: 'ignored', label: text };
          issues.push({ level: 'warning', where: 'Payslips', message: prefix + 'Column "' + text + '" also maps to "' + def.label + '" (already used by another column). It is ignored; change the mapping if this is wrong.' });
          source = 'duplicate';
        } else {
          seen[def.key] = text;
          m = { kind: 'column', def: def };
        }
      } else if (target.kind === 'earning' || target.kind === 'deduction') {
        m = { kind: target.kind, label: target.label || text };
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

    var ignored = columns.filter(function (c) { return c.kind === 'ignored' && c.source !== 'silent'; });
    if (ignored.length) {
      issues.push({ level: 'warning', where: 'Payslips', code: 'ignored-columns', message: prefix + ignored.length + ' column' + (ignored.length === 1 ? ' is' : 's are') + ' not recognised and will be ignored: ' + ignored.map(function (c) { return '"' + c.header + '"'; }).join(', ') + '. Use "Columns" below to map them.' });
    }
    var guessed = columns.filter(function (c) { return c.source === 'guess'; });
    if (guessed.length) {
      issues.push({ level: 'info', where: 'Payslips', code: 'guessed-columns', message: prefix + guessed.length + ' column' + (guessed.length === 1 ? ' was' : 's were') + ' matched by keyword: ' + guessed.map(function (c) { return '"' + c.header + '" → ' + describeTarget(c); }).join(', ') + '. Check them under "Columns".' });
    }
    return { mapping: mapping, columns: columns, seen: seen };
  }

  function describeTarget(m) {
    if (m.kind === 'column') return m.def.label;
    if (m.kind === 'earning') return 'Earning "' + m.label + '"';
    if (m.kind === 'deduction') return 'Deduction "' + m.label + '"';
    return 'Ignored';
  }

  function isMoneyMapping(m) {
    return !!m && ((m.kind === 'column' && (m.def.type === 'money' || m.def.type === 'number')) || m.kind === 'earning' || m.kind === 'deduction');
  }

  function isDeductionMapping(m) {
    return !!m && ((m.kind === 'column' && m.def.group === 'deduction') || m.kind === 'deduction');
  }

  // Per-column decisions taken from all values of the column: decimal comma, negative deductions, date order.
  function inspectColumns(mapping, body, notes) {
    mapping.forEach(function (m) {
      if (!m || m.kind === 'ignored') return;
      var vals = [];
      for (var i = 0; i < body.length; i++) {
        var v = body[i] && body[i][m.index];
        if (!isBlank(v)) vals.push(v);
      }
      if (isMoneyMapping(m)) {
        var strs = vals.filter(function (v) { return typeof v === 'string'; }).map(function (v) { return toAsciiDigits(v).replace(/[^0-9.,]/g, ''); });
        if (strs.some(looksDecimalComma)) {
          m.decimalComma = true;
          notes.decimalComma.push(m.header);
        }
        if (isDeductionMapping(m)) {
          var nums = vals.map(function (v) { return parseNumber(v, { decimalComma: m.decimalComma }); })
            .filter(function (p) { return !p.error && p.value != null && p.value !== 0; })
            .map(function (p) { return p.value; });
          if (nums.length && nums.every(function (x) { return x < 0; })) {
            m.negated = true;
            notes.negated.push(m.header);
          }
        }
      }
      if (m.kind === 'column' && m.def.type === 'date') {
        var dayFirst = false;
        var monthFirst = false;
        vals.forEach(function (v) {
          if (typeof v !== 'string') return;
          var mm = toAsciiDigits(v).trim().match(/^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{2}|\d{4})/);
          if (!mm) return;
          if (+mm[1] > 12) dayFirst = true;
          if (+mm[2] > 12) monthFirst = true;
        });
        m.dateOrder = monthFirst && !dayFirst ? 'mdy' : 'dmy';
        if (monthFirst && !dayFirst) notes.mdy.push(m.header);
        if (monthFirst && dayFirst) notes.mixedDates.push(m.header);
      }
    });
  }

  function parseRow(rawRow, mapping, rowNumber, countryCode) {
    var slip = {
      _countryCode: countryCode,
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
          var p = parseNumber(raw, { decimalComma: m.decimalComma });
          if (m.negated && p.value != null && p.value < 0) p.value = -p.value;
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
          var rd = readDate(raw, m.dateOrder);
          slip.fields[def.key] = rd.text;
          if (rd.invalid && !isBlank(raw)) slip.issues.push({ level: 'warning', message: '"' + m.header + '" is not a valid date ("' + cleanText(raw) + '").' });
          if (def.key === 'dob') slip.fields.dobDate = isoDate(rd.date);
        } else if (def.type === 'phone') {
          var ph = normalizePhone(raw, slip._countryCode);
          slip.fields[def.key] = ph.value;
          if (ph.error) slip.issues.push({ level: 'warning', message: '"' + m.header + '" ' + ph.error + '. WhatsApp delivery will not work for this row.' });
        } else if (def.type === 'email') {
          var em = normalizeEmail(raw);
          slip.fields[def.key] = em.value;
          if (em.error) slip.issues.push({ level: 'warning', message: '"' + m.header + '" ' + em.error + '. Email delivery will not work for this row.' });
        } else {
          var txt = cleanText(raw);
          if (EXCEL_ERROR_RE.test(txt)) {
            slip.issues.push({ level: def.key === 'name' ? 'error' : 'warning', message: '"' + m.header + '" contains the Excel error ' + txt + '.' });
            txt = '';
          }
          if (def.key === 'employeeId' || def.key === 'nationalId' || def.key === 'iban') txt = toAsciiDigits(txt);
          if (m.namePart) slip.fields.name = [slip.fields.name, txt].filter(Boolean).join(' ');
          else slip.fields[def.key] = txt;
        }
      } else if (m.kind === 'earning' || m.kind === 'deduction') {
        var pn = parseNumber(raw, { decimalComma: m.decimalComma });
        if (m.negated && pn.value != null && pn.value < 0) pn.value = -pn.value;
        if (pn.error) {
          slip.issues.push({ level: 'error', message: '"' + m.header + '" ' + pn.error + '.' });
        } else if (pn.value != null && pn.value < 0) {
          slip.issues.push({ level: 'error', message: '"' + m.header + '" is negative (' + formatMoney(pn.value) + ').' });
        } else if (pn.value) {
          slip[m.kind === 'earning' ? 'earnings' : 'deductions'].push({ key: m.netOnly ? 'netOnly' : 'custom', label: m.label, amount: pn.value });
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

  // Best label to show for a skipped row.
  function rowLabel(raw, mapping) {
    for (var i = 0; i < mapping.length; i++) {
      var m = mapping[i];
      if (m && m.kind === 'column' && m.def.key === 'name' && !isBlank(raw[i])) return cleanText(raw[i]);
    }
    for (var j = 0; j < raw.length; j++) if (!isBlank(raw[j]) && typeof raw[j] === 'string') return cleanText(raw[j]);
    return '';
  }

  // Parses the employee rows of one sheet.
  function parseSheet(sh, ctx) {
    var aoa = sh.aoa;
    var issues = ctx.issues;
    var prefix = ctx.combined ? 'Sheet "' + sh.name + '": ' : '';
    var headerRow = aoa[sh.headerIdx] || [];
    var body = aoa.slice(sh.headerIdx + 1);
    var sampleRows = body.filter(function (r) { return r && !r.every(isBlank); }).slice(0, 30);
    var mapped = mapHeaders(headerRow, groupRowHints(aoa, sh.headerIdx), ctx.overrides, issues, sampleRows, prefix);
    var mapping = mapped.mapping;
    var seen = mapped.seen;
    var columns = mapped.columns;

    var hasEarningCol = columns.some(function (c) { return (c.kind === 'column' && c.def.group === 'earning') || c.kind === 'earning'; });
    if (!hasEarningCol && seen.netSalary) {
      // Only a net column: one payslip line "Net Salary" instead of nothing at all.
      columns.forEach(function (c) {
        if (c.kind === 'column' && c.def.key === 'netSalary') {
          c.kind = 'earning';
          c.label = 'Net Salary';
          c.netOnly = true;
          delete c.def;
        }
      });
      delete seen.netSalary;
      hasEarningCol = true;
      issues.push({ level: 'warning', where: 'Payslips', code: 'net-only', message: prefix + 'Only a net salary column was found, so each payslip shows one "Net Salary" line. Add basic salary and allowance columns for a full breakdown.' });
    }
    inspectColumns(mapping, body, ctx.notes);

    if (!seen.name) issues.push({ level: 'error', where: 'Payslips', code: 'missing-name', message: prefix + 'No "Employee Name" column was found. Map one of your columns to Employee Name under "Columns".' });
    if (!hasEarningCol) issues.push({ level: 'error', where: 'Payslips', code: 'missing-earnings', message: prefix + 'No salary column was found (Basic Salary or any earning). Map one of your columns under "Columns".' });
    else if (!seen.basicSalary && !columns.some(function (c) { return c.netOnly; })) issues.push({ level: 'warning', where: 'Payslips', code: 'missing-basic', message: prefix + 'No "Basic Salary" column was found. Earnings are taken from the other pay columns.' });
    if (!seen.employeeId) issues.push({ level: 'warning', where: 'Payslips', code: 'missing-id', message: prefix + 'No "Employee ID" column was found. IDs are generated from the row order (EMP-001, EMP-002…). Map a column to Employee ID if you have one.' });

    var section = '';
    var slips = [];
    for (var i = 0; i < body.length; i++) {
      var r = sh.headerIdx + 1 + i;
      var rowNo = r + sh.rowOffset + 1;
      var raw = body[i] || [];
      if (raw.every(isBlank)) continue;
      if (sh.hidden[r + sh.rowOffset] && !ctx.includeHidden) {
        ctx.hiddenRows.push({ sheet: sh.name, row: rowNo, label: rowLabel(raw, mapping) });
        continue;
      }
      var nonBlank = raw.filter(function (v) { return !isBlank(v); });
      if (nonBlank.length === 1 && typeof nonBlank[0] === 'string' && !looksNumeric(nonBlank[0]) && !TOTAL_ROW_RE.test(cleanText(nonBlank[0])) && !EXCEL_ERROR_RE.test(cleanText(nonBlank[0]))) {
        // A lone label such as "PROJECTS DEPARTMENT": a section heading, used as department for the rows below.
        section = cleanText(nonBlank[0]);
        ctx.sections.push(section);
        continue;
      }
      if (isTotalsRow(raw, mapping)) { ctx.skipped.totals++; continue; }
      var slip = parseRow(raw, mapping, rowNo, ctx.countryCode);
      if (slip.isExample) { ctx.skipped.example++; continue; }
      if ((isBlank(slip.name) && isBlank(slip.employeeId)) || (isBlank(slip.name) && slip.gross === 0 && slip.totalDeductions === 0)) {
        ctx.skipped.unlabelled++;
        continue;
      }
      slip.sheet = sh.name;
      slip.rowRef = ctx.combined ? sh.name + ' · ' + rowNo : String(rowNo);
      if (!slip.fields.department && section) slip.fields.department = section;
      else if (!slip.fields.department && ctx.combined) slip.fields.department = sh.name;
      ctx.seq++;
      if (!seen.employeeId || isBlank(slip.employeeId)) {
        slip.employeeId = 'EMP-' + String(ctx.seq).padStart(3, '0');
        slip.generatedId = true;
        if (seen.employeeId) slip.issues.push({ level: 'warning', message: 'Employee ID is empty; "' + slip.employeeId + '" was generated.' });
      }
      slips.push(slip);
    }
    sh.employees = slips.length;
    return { slips: slips, columns: columns, seen: seen };
  }

  // Decides which sheets hold the employees: one sheet, all sheets combined, or the right month.
  function chooseSheets(sheets, options, payPeriod, issues) {
    var data = sheets.filter(function (s) { return s.isData; });
    if (options.sheet && options.sheet !== 'auto') {
      if (options.sheet === '__all__') return { used: data.length ? data : sheets.slice(0, 1), mode: 'combined', how: 'chosen' };
      var one = sheets.filter(function (s) { return s.name === options.sheet; })[0];
      if (one) return { used: [one], mode: 'single', how: 'chosen' };
    }
    if (data.length <= 1) {
      var pick = data[0] || sheets.filter(function (s) { return s.rows > 0; })[0] || sheets[0];
      return { used: pick ? [pick] : [], mode: 'single', how: 'only' };
    }
    var preferred = data.filter(function (s) { return s.preferred; })[0];
    var months = data.filter(function (s) { return s.month; });
    if (months.length >= Math.ceil(data.length / 2)) {
      var want = monthInfo(payPeriod);
      var match = want && months.filter(function (s) { return s.month.month === want.month && (!s.month.year || !want.year || s.month.year === want.year); })[0];
      var chosen = match;
      var why = 'it matches the pay period "' + payPeriod + '"';
      if (!chosen) {
        chosen = months.slice().sort(function (a, b) {
          var ka = (a.month.year || 0) * 12 + a.month.month;
          var kb = (b.month.year || 0) * 12 + b.month.month;
          return ka - kb || sheets.indexOf(a) - sheets.indexOf(b);
        })[months.length - 1];
        why = 'it is the latest month';
      }
      issues.push({ level: 'warning', where: 'Payslips', code: 'month-sheet', message: 'This file has ' + months.length + ' monthly sheets. Using "' + chosen.name + '" because ' + why + '. Pick another sheet at the top if needed.' });
      return { used: [chosen], mode: 'month', how: 'auto' };
    }
    if (preferred && data.length > 1 && data.every(function (s) { return s === preferred || s.rows === 0; })) return { used: [preferred], mode: 'single', how: 'only' };
    return { used: data, mode: 'combined', how: 'auto' };
  }

  /**
   * Parse a workbook (read with XLSX.read(..., { cellDates: true, cellStyles: true })).
   * options.overrides   { [normalizedHeader]: target } manual column mapping
   * options.sheet       'auto' (default) | '__all__' (combine) | a sheet name
   * options.headerRow   sheet row number (1-based) to use as header for the chosen sheet
   * options.includeHidden  give payslips to hidden rows as well
   * options.payPeriod   hint for picking the right monthly sheet
   */
  function parseWorkbook(XLSX, wb, options) {
    XLSXRef = XLSX;
    options = options || {};
    var issues = [];
    var result = {};
    var settings = parseSettings(findSheet(wb, SHEET_SETTINGS), issues, result);
    var visible = wb.SheetNames.filter(function (n) { return !isHiddenSheet(wb, n); });
    var hiddenSheets = wb.SheetNames.filter(function (n) { return isHiddenSheet(wb, n) && SKIP_SHEETS.indexOf(normalizeHeader(n)) < 0; });
    var candidates = visible.filter(function (n) { return SKIP_SHEETS.indexOf(normalizeHeader(n)) < 0; });
    if (!candidates.length) candidates = visible.length ? visible : wb.SheetNames.slice();
    var forcedSheet = options.sheet && options.sheet !== 'auto' && options.sheet !== '__all__' ? options.sheet : null;
    var sheets = candidates.map(function (n) {
      return analyzeSheet(XLSX, wb, n, n === forcedSheet ? options.headerRow : null);
    });
    if (forcedSheet && options.headerRow == null) {
      // nothing forced
    }
    var choice = chooseSheets(sheets, options, options.payPeriod || settings.payPeriod, issues);
    var ctx = {
      issues: issues, overrides: options.overrides, countryCode: settings.countryCode,
      includeHidden: !!options.includeHidden, combined: choice.mode === 'combined' && choice.used.length > 1,
      seq: 0, hiddenRows: [], sections: [], skipped: { totals: 0, unlabelled: 0, example: 0 },
      notes: { decimalComma: [], negated: [], mdy: [], mixedDates: [] }
    };
    var slips = [];
    var columnsByNorm = {};
    var columns = [];
    var seenAll = {};

    if (!choice.used.length || choice.used.every(function (s) { return !s.aoa.length; })) {
      issues.push({ level: 'error', where: 'Payslips', code: 'no-data', message: 'The file has no data rows.' });
    } else {
      if (choice.used.length === 1 && !choice.used[0].isData) {
        issues.push({ level: 'error', where: 'Payslips', code: 'no-header', message: 'No recognisable header row was found. Map your columns below (Employee Name and at least one salary column are needed), pick the header row at the top, or download the template and copy your data into it.' });
      }
      choice.used.forEach(function (sh) {
        var res = parseSheet(sh, ctx);
        slips = slips.concat(res.slips);
        Object.keys(res.seen).forEach(function (k) { seenAll[k] = true; });
        res.columns.forEach(function (c) {
          if (!columnsByNorm[c.norm]) {
            columnsByNorm[c.norm] = c;
            columns.push(c);
          }
        });
      });
      if (seenAll.name && columns.some(function (c) { return c.kind === 'earning' || (c.kind === 'column' && c.def && c.def.group === 'earning'); })) {
        // manual mapping rescued an unrecognisable header row
        issues = issues.filter(function (i) { return i.code !== 'no-header'; });
        ctx.issues = issues;
      }

      if (ctx.combined) {
        issues.push({ level: 'info', where: 'Payslips', code: 'combined', message: 'Combined ' + choice.used.length + ' sheets: ' + choice.used.map(function (s) { return s.name + ' (' + s.employees + ')'; }).join(', ') + '. Pick one sheet at the top to use only that one.' });
      } else if (choice.used[0] && normalizeHeader(choice.used[0].name) !== normalizeHeader(SHEET_PAYSLIPS) && choice.mode !== 'month') {
        issues.push({ level: 'info', where: 'Payslips', message: 'Reading employee rows from sheet "' + choice.used[0].name + '" (row ' + choice.used[0].headerRow + ' used as the header).' });
      } else if (choice.used[0] && choice.used[0].headerRow > 1 && choice.mode !== 'month') {
        issues.push({ level: 'info', where: 'Payslips', message: 'Row ' + choice.used[0].headerRow + ' of sheet "' + choice.used[0].name + '" is used as the header row.' });
      }
      if (hiddenSheets.length) {
        issues.push({ level: 'info', where: 'Payslips', message: 'Hidden sheet' + (hiddenSheets.length === 1 ? '' : 's') + ' ignored: ' + hiddenSheets.map(function (n) { return '"' + n + '"'; }).join(', ') + '.' });
      }
      var n = ctx.notes;
      if (n.decimalComma.length) issues.push({ level: 'info', where: 'Payslips', code: 'decimal-comma', message: 'Amounts written with a decimal comma (3.000,50) were read correctly in: ' + n.decimalComma.join(', ') + '.' });
      if (n.negated.length) issues.push({ level: 'info', where: 'Payslips', code: 'negative-deductions', message: 'Deductions written as negative numbers were read as deductions in: ' + n.negated.join(', ') + '.' });
      if (n.mdy.length) issues.push({ level: 'info', where: 'Payslips', code: 'us-dates', message: 'Dates in ' + n.mdy.join(', ') + ' are month/day/year (like 03/15/2024) and were read that way.' });
      if (n.mixedDates.length) issues.push({ level: 'warning', where: 'Payslips', code: 'mixed-dates', message: 'Dates in ' + n.mixedDates.join(', ') + ' mix day/month and month/day order. Check them in Excel.' });
      if (ctx.hiddenRows.length) {
        issues.push({ level: 'warning', where: 'Payslips', code: 'hidden-rows', message: ctx.hiddenRows.length + ' hidden row' + (ctx.hiddenRows.length === 1 ? ' was' : 's were') + ' skipped (' + ctx.hiddenRows.slice(0, 5).map(function (h) { return (ctx.combined ? h.sheet + ' ' : '') + 'row ' + h.row + (h.label ? ' ' + h.label : ''); }).join(', ') + (ctx.hiddenRows.length > 5 ? ', …' : '') + '). Tick "Include hidden rows" if they should get payslips.' });
      }
      if (ctx.sections.length) {
        issues.push({ level: 'info', where: 'Payslips', code: 'sections', message: ctx.sections.length + ' heading row' + (ctx.sections.length === 1 ? ' was' : 's were') + ' skipped and used as department: ' + ctx.sections.slice(0, 6).map(function (s) { return '"' + s + '"'; }).join(', ') + (ctx.sections.length > 6 ? ', …' : '') + '.' });
      }
      var skipped = ctx.skipped;
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
          var rowsTxt = byId[k].map(function (s) { return s.rowRef; }).join(', ');
          byId[k].forEach(function (s) {
            s.issues.push({ level: 'error', message: 'Employee ID "' + s.employeeId + '" is duplicated (rows ' + rowsTxt + ').' });
            s.hasErrors = true;
          });
        }
      });
      // Duplicate contacts
      ['mobile', 'email'].forEach(function (k) {
        var seenC = {};
        slips.forEach(function (sl) {
          var v = sl.fields[k];
          if (!v) return;
          (seenC[v] = seenC[v] || []).push(sl);
        });
        Object.keys(seenC).forEach(function (v) {
          if (seenC[v].length > 1) {
            var rowsTxt = seenC[v].map(function (sl) { return sl.rowRef; }).join(', ');
            seenC[v].forEach(function (sl) {
              sl.issues.push({ level: 'warning', message: (k === 'mobile' ? 'Mobile' : 'Email') + ' "' + v + '" is shared with other rows (' + rowsTxt + '). Each employee would receive the others\' payslips.' });
            });
          }
        });
      });
      slips.forEach(function (sl) { delete sl._countryCode; });
      if (!slips.length) {
        issues.push({ level: 'error', where: 'Payslips', code: 'no-rows', message: 'No employee rows were found under the header row.' });
      }
    }

    // the same message can come from several combined sheets
    var seenMsg = {};
    issues = issues.filter(function (i) {
      if (seenMsg[i.level + i.message]) return false;
      seenMsg[i.level + i.message] = true;
      return true;
    });

    var stats = {
      total: slips.length,
      withErrors: slips.filter(function (s) { return s.hasErrors; }).length,
      withWarnings: slips.filter(function (s) { return !s.hasErrors && s.issues.length; }).length,
      fileErrors: issues.filter(function (i) { return i.level === 'error'; }).length,
      fileWarnings: issues.filter(function (i) { return i.level === 'warning'; }).length,
      totalNet: round2(slips.reduce(function (a, s) { return a + (s.hasErrors ? 0 : s.net); }, 0)),
      totalGross: round2(slips.reduce(function (a, s) { return a + (s.hasErrors ? 0 : s.gross); }, 0)),
      totalDeductions: round2(slips.reduce(function (a, s) { return a + (s.hasErrors ? 0 : s.totalDeductions); }, 0)),
      skipped: ctx.skipped,
      hiddenRows: ctx.hiddenRows.length,
      withMobile: slips.filter(function (s) { return !!s.fields.mobile; }).length,
      withEmail: slips.filter(function (s) { return !!s.fields.email; }).length
    };
    stats.withNoContact = slips.filter(function (s) { return !s.fields.mobile && !s.fields.email; }).length;
    stats.ok = stats.total - stats.withErrors;

    var first = choice.used[0];
    return {
      settings: settings,
      settingsFound: !!result.settingsFound,
      slips: slips,
      issues: issues,
      stats: stats,
      columns: columns,
      sheets: sheets.map(function (s) {
        return { name: s.name, isData: s.isData, headerRow: s.headerRow, rows: s.rows, employees: s.employees || 0, month: s.month, used: choice.used.indexOf(s) >= 0, firstRow: s.rowOffset + 1, lastRow: s.rowOffset + s.aoa.length };
      }),
      sheetMode: choice.mode,
      sheetName: ctx.combined ? 'All sheets' : (first ? first.name : null),
      headerRow: first ? first.headerRow : null,
      headerPreview: first ? headerPreview(first) : [],
      hiddenRows: ctx.hiddenRows,
      missingRequired: ['name', 'basicSalary', 'employeeId'].filter(function (k) { return !seenAll[k]; })
    };
  }

  // Rows offered by the "header row" picker: sheet row number and a short preview of its cells.
  function headerPreview(sh) {
    var out = [];
    for (var i = 0; i < Math.min(sh.aoa.length, 40); i++) {
      var cells = (sh.aoa[i] || []).filter(function (v) { return !isBlank(v); }).slice(0, 4).map(function (v) { return cleanText(v).slice(0, 18); });
      if (cells.length) out.push({ row: i + sh.rowOffset + 1, text: cells.join(' · ') });
    }
    return out;
  }

  // Choices offered by the column-mapping UI.
  function mappingChoices() {
    var list = [{ value: 'ignored', label: 'Ignore this column' }];
    var groups = { info: 'Employee details', contact: 'Contact / delivery', attendance: 'Attendance', earning: 'Earnings', deduction: 'Deductions', check: 'Cross-check only' };
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
        new Date(1975 + (i % 25), (i * 5) % 12, 1 + ((i * 3) % 28)),
        titles[i % titles.length],
        depts[i % depts.length],
        new Date(2022 + (i % 4), i % 12, 1 + (i % 27)),
        '05' + String(50000000 + i * 7919).slice(-8),
        (first[i % first.length] + '.' + last[(i * 3) % last.length].replace(/[^A-Za-z]/g, '') + (i + 1) + '@example.com').toLowerCase(),
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

  // Same sample data keyed by column key (used by the styled template builder).
  function sampleRecords(count) {
    var cols = templateColumns();
    return sampleRows(count).map(function (row) {
      var rec = {};
      cols.forEach(function (c, i) { rec[c.key] = row[i]; });
      return rec;
    });
  }

  // Earlier default message texts; a saved copy equal to one of these is upgraded to the current default.
  var OLD_MESSAGE_DEFAULTS = [
    'Dear {name}, your payslip for {period} is attached. {password_hint}\nThis document contains personal data; please do not forward it.\n\nعزيزي {name}، مرفق قسيمة راتبك لشهر {period}. {password_hint_ar}\nهذا المستند يحتوي على بيانات شخصية، يرجى عدم مشاركته.',
    'Dear {name},\n\nPlease find attached your payslip for {period}.\n\n{password_hint}\n\n{company}'
  ];

  return {
    OLD_MESSAGE_DEFAULTS: OLD_MESSAGE_DEFAULTS,
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
    emptySettings: emptySettings,
    toDate: toDate,
    isoDate: isoDate,
    normalizePhone: normalizePhone,
    toAsciiDigits: toAsciiDigits,
    readDate: readDate,
    monthInfo: monthInfo,
    sampleRecords: sampleRecords,
    normalizeEmail: normalizeEmail,
    periodSlug: periodSlug,
    asciiSlug: asciiSlug,
    slipFileName: slipFileName,
    buildPassword: buildPassword,
    renderTemplate: renderTemplate,
    messageVars: messageVars
  };
});
