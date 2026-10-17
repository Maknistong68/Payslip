/*
 * Payslip Generator - PDF rendering (jsPDF).
 * renderPayslip(doc, ctx) draws ONE payslip on the current page of `doc`.
 *
 * ctx = {
 *   slip, settings,
 *   templateId: 'classic' | 'modern' | 'minimal',
 *   accent: '#hex',
 *   logo:  { dataUrl, width, height } | null,
 *   stamp: { dataUrl, width, height } | null,
 *   options: { amountInWords, signatureLines, showAttendance, showBank }
 * }
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory(require('./schema.js'), typeof global !== 'undefined' ? global : root);
  } else {
    root.PayslipPDF = factory(root.PayslipSchema, root);
  }
})(typeof self !== 'undefined' ? self : this, function (Schema, root) {
  'use strict';

  // ---- Arabic support -------------------------------------------------------
  // The built-in PDF fonts only cover Latin. When a string contains Arabic we switch to
  // Amiri (bundled in js/fonts/amiri.js); jsPDF's Arabic plugin does the letter shaping and
  // right-to-left ordering. The font is registered on a document only when first needed.
  var ARABIC_RE = /[\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF\uFB50-\uFDFF\uFE70-\uFEFF]/;

  function hasArabic(s) {
    return ARABIC_RE.test(String(s == null ? '' : s));
  }

  function arabicFontAvailable() {
    return !!(root && root.PayslipFonts && root.PayslipFonts.Amiri);
  }

  function ensureArabicFont(doc) {
    if (doc.__payslipArabic !== undefined) return doc.__payslipArabic;
    if (!arabicFontAvailable()) {
      doc.__payslipArabic = false;
      return false;
    }
    var f = root.PayslipFonts.Amiri;
    doc.addFileToVFS('Amiri-Regular.ttf', f.regular);
    doc.addFont('Amiri-Regular.ttf', 'Amiri', 'normal');
    doc.addFileToVFS('Amiri-Bold.ttf', f.bold);
    doc.addFont('Amiri-Bold.ttf', 'Amiri', 'bold');
    doc.__payslipArabic = true;
    return true;
  }

  var INK = [15, 23, 42];
  var STEEL = [51, 65, 85];
  var GOLD = [184, 140, 40];
  var BLACK = [0, 0, 0];
  var WHITE = [255, 255, 255];

  /*
   * Each template is a design spec read by renderPayslip():
   *   header  rule | band | split | sidebar | topbar | boxed | centered | receipt
   *   info    outline | soft | lines | accentLeft | plain        (employee details box)
   *   head    grey | accent | dark | line                        (table header row)
   *   totals  soft | line | accent ;  tables side | stacked
   *   net     box | band | pill | dark | outline | lines
   *   font / bodyFont helvetica | times | courier ; accent fixes the colour, otherwise the user's accent is used
   */
  var TEMPLATES = [
    { id: 'classic', name: 'Classic', description: 'Black and white, bordered. Prints anywhere.', font: 'times', accent: BLACK, header: 'rule', ruleW: 0.6, info: 'outline', head: 'grey', headFill: [225, 225, 225], borders: true, totals: 'soft', net: 'box', rule: [120, 120, 120] },
    { id: 'modern', name: 'Modern', description: 'Colour band header, highlighted net pay.', header: 'band', info: 'soft', head: 'accent', borders: true, zebra: true, totals: 'soft', net: 'band', radius: 1.5 },
    { id: 'minimal', name: 'Minimal', description: 'Thin lines, lots of white space.', header: 'rule', info: 'lines', head: 'line', borders: false, totals: 'line', net: 'lines', rule: [222, 226, 230] },
    { id: 'executive', name: 'Executive', description: 'Navy band with gold details.', header: 'band', headerFill: INK, accent: GOLD, goldLine: true, info: 'lines', head: 'line', borders: false, totals: 'line', net: 'dark' },
    { id: 'steel', name: 'Steel', description: 'Slate header and tables, hard contrast.', header: 'band', headerFill: STEEL, info: 'accentLeft', head: 'dark', headDark: STEEL, borders: true, zebra: true, totals: 'accent', net: 'dark', darkFill: STEEL },
    { id: 'corporate', name: 'Corporate', description: 'Colour strip down the left edge.', header: 'sidebar', info: 'plain', head: 'grey', borders: true, totals: 'soft', net: 'outline' },
    { id: 'bold', name: 'Bold', description: 'Heavy top bar, big title, solid totals.', header: 'topbar', titleSize: 26, info: 'accentLeft', head: 'accent', borders: false, zebra: true, totals: 'accent', net: 'band' },
    { id: 'ledger', name: 'Ledger', description: 'Typewriter type, full grid. Accountant style.', font: 'courier', bodyFont: 'courier', accent: BLACK, header: 'boxed', info: 'outline', head: 'grey', borders: true, totals: 'soft', net: 'box', rule: [90, 90, 90] },
    { id: 'compact', name: 'Compact', description: 'Stacked tables for long pay item lists.', header: 'rule', compact: true, info: 'soft', head: 'accent', borders: false, zebra: true, totals: 'soft', tables: 'stacked', net: 'pill', radius: 1.5 },
    { id: 'split', name: 'Split', description: 'Two-tone header: colour block and white.', header: 'split', info: 'soft', head: 'accent', borders: true, totals: 'soft', net: 'band', radius: 1.5 },
    { id: 'formal', name: 'Formal', description: 'Centred letterhead, serif type.', font: 'times', header: 'centered', info: 'outline', head: 'line', borders: false, totals: 'line', net: 'outline' },
    { id: 'heritage', name: 'Heritage', description: 'Double frame border, certificate feel.', font: 'times', header: 'centered', frame: true, info: 'lines', head: 'grey', borders: true, totals: 'soft', net: 'box' },
    { id: 'receipt', name: 'Receipt', description: 'Narrow till-receipt column, dashed cuts.', font: 'courier', bodyFont: 'courier', accent: BLACK, header: 'receipt', info: 'plain', head: 'line', borders: false, totals: 'line', tables: 'stacked', net: 'lines', dashed: true }
  ];

  function templateSpec(id) {
    for (var i = 0; i < TEMPLATES.length; i++) if (TEMPLATES[i].id === id) return TEMPLATES[i];
    return TEMPLATES[0];
  }

  // Everything the user can switch on/off on the payslip. Grouped for the settings UI.
  var DISPLAY_OPTIONS = [
    { key: 'showCr', label: 'CR number', group: 'Company header', def: true },
    { key: 'showAddress', label: 'Company address', group: 'Company header', def: true },
    { key: 'showContact', label: 'Phone / email', group: 'Company header', def: true },
    { key: 'showPeriodDates', label: 'Period start / end dates', group: 'Company header', def: true },
    { key: 'showPaymentDate', label: 'Payment date', group: 'Company header', def: true },
    { key: 'f_employeeId', label: 'Employee ID', group: 'Employee details', def: true },
    { key: 'f_nationalId', label: 'National ID / Iqama', group: 'Employee details', def: true },
    { key: 'f_nationality', label: 'Nationality', group: 'Employee details', def: true },
    { key: 'f_jobTitle', label: 'Job title', group: 'Employee details', def: true },
    { key: 'f_department', label: 'Department', group: 'Employee details', def: true },
    { key: 'f_joiningDate', label: 'Joining date', group: 'Employee details', def: true },
    { key: 'f_dob', label: 'Date of birth', group: 'Employee details', def: false },
    { key: 'f_mobile', label: 'Mobile', group: 'Employee details', def: false },
    { key: 'f_email', label: 'Email', group: 'Employee details', def: false },
    { key: 'f_bankName', label: 'Bank name', group: 'Employee details', def: true },
    { key: 'f_iban', label: 'IBAN', group: 'Employee details', def: true },
    { key: 'f_workingDays', label: 'Working days', group: 'Attendance', def: true },
    { key: 'f_absentDays', label: 'Absent days', group: 'Attendance', def: true },
    { key: 'f_overtimeHours', label: 'Overtime hours', group: 'Attendance', def: true },
    { key: 'amountInWords', label: 'Net amount in words', group: 'Totals & footer', def: true },
    { key: 'showRemarks', label: 'Remarks line', group: 'Totals & footer', def: true },
    { key: 'signatureLines', label: 'Signature lines', group: 'Totals & footer', def: true },
    { key: 'showFooterNote', label: 'Footer note', group: 'Totals & footer', def: true },
    { key: 'showPreparedBy', label: '"Prepared by"', group: 'Totals & footer', def: true },
    { key: 'showGeneratedOn', label: '"Generated on" date', group: 'Totals & footer', def: false }
  ];

  var DEFAULT_OPTIONS = { title: 'PAYSLIP' };
  DISPLAY_OPTIONS.forEach(function (o) { DEFAULT_OPTIONS[o.key] = o.def; });

  function mergeOptions(options) {
    var out = {};
    Object.keys(DEFAULT_OPTIONS).forEach(function (k) { out[k] = DEFAULT_OPTIONS[k]; });
    if (options) Object.keys(options).forEach(function (k) { if (options[k] !== undefined) out[k] = options[k]; });
    if (!out.title || !String(out.title).trim()) out.title = 'PAYSLIP';
    return out;
  }

  var PAPERS = {
    a4: { format: 'a4', width: 210, height: 297 },
    letter: { format: 'letter', width: 215.9, height: 279.4 }
  };

  function hexToRgb(hex) {
    var h = String(hex || '#1f4e79').replace('#', '');
    if (h.length === 3) h = h.split('').map(function (c) { return c + c; }).join('');
    var n = parseInt(h, 16);
    if (isNaN(n)) return [31, 78, 121];
    return [(n >> 16) & 255, (n >> 8) & 255, n & 255];
  }

  function mix(rgb, white) {
    // blend towards white by factor 0..1
    return rgb.map(function (c) { return Math.round(c + (255 - c) * white); });
  }

  function fmt(n) {
    return Schema.formatMoney(n);
  }

  function fitImage(img, maxW, maxH) {
    var w = img.width || 1;
    var h = img.height || 1;
    var scale = Math.min(maxW / w, maxH / h, 1e9);
    return { w: w * scale, h: h * scale };
  }

  function imageFormat(dataUrl) {
    return /^data:image\/jpe?g/i.test(dataUrl) ? 'JPEG' : 'PNG';
  }

  function sanitize(s) {
    // Replace characters neither font can render with '?' to avoid garbage glyph runs.
    return String(s == null ? '' : s).replace(/[^\u0000-ɏ؀-ۿݐ-ݿࢠ-ࣿﭐ-﷿ﹰ-﻿ -⁯₠-⃏]/g, '?');
  }

  // ---------------------------------------------------------------------------
  // Drawing primitives
  // ---------------------------------------------------------------------------

  function Painter(doc, style) {
    this.doc = doc;
    this.style = style;
  }

  Painter.prototype.text = function (str, x, y, opts) {
    opts = opts || {};
    var d = this.doc;
    var s = sanitize(str);
    var arabic = hasArabic(s) && ensureArabicFont(d);
    if (arabic) {
      d.setFont('Amiri', opts.weight === 'bold' ? 'bold' : 'normal');
      d.setFontSize((opts.size || 9) * 1.12); // Amiri runs visually smaller than Helvetica
    } else {
      d.setFont(opts.font || this.style.font, opts.weight || 'normal');
      d.setFontSize(opts.size || 9);
    }
    var c = opts.color || this.style.text;
    d.setTextColor(c[0], c[1], c[2]);
    if (opts.maxWidth) {
      // truncate with ellipsis instead of wrapping when asked
      if (opts.truncate) {
        while (s.length > 1 && d.getTextWidth(s) > opts.maxWidth) s = s.slice(0, -2) + '…';
        d.text(s, x, y, { align: opts.align || 'left', baseline: opts.baseline || 'alphabetic' });
        return 1;
      }
      var lines = d.splitTextToSize(s, opts.maxWidth);
      if (opts.maxLines && lines.length > opts.maxLines) {
        lines = lines.slice(0, opts.maxLines);
        lines[lines.length - 1] = lines[lines.length - 1].replace(/.{2}$/, '') + '…';
      }
      d.text(lines, x, y, { align: opts.align || 'left', baseline: opts.baseline || 'alphabetic', lineHeightFactor: 1.25 });
      return lines.length;
    }
    d.text(s, x, y, { align: opts.align || 'left', baseline: opts.baseline || 'alphabetic' });
    return 1;
  };

  Painter.prototype.line = function (x1, y1, x2, y2, color, width, dashed) {
    var c = color || this.style.rule;
    var d = this.doc;
    d.setDrawColor(c[0], c[1], c[2]);
    d.setLineWidth(width || 0.2);
    if (dashed && d.setLineDashPattern) d.setLineDashPattern([1.2, 1], 0);
    d.line(x1, y1, x2, y2);
    if (dashed && d.setLineDashPattern) d.setLineDashPattern([], 0);
  };

  Painter.prototype.rect = function (x, y, w, h, fill, stroke, radius, lineWidth) {
    var d = this.doc;
    var mode = '';
    if (fill) {
      d.setFillColor(fill[0], fill[1], fill[2]);
      mode += 'F';
    }
    if (stroke) {
      d.setDrawColor(stroke[0], stroke[1], stroke[2]);
      d.setLineWidth(lineWidth || 0.2);
      mode += 'D';
    }
    if (!mode) return;
    if (radius) d.roundedRect(x, y, w, h, radius, radius, mode);
    else d.rect(x, y, w, h, mode);
  };

  // ---------------------------------------------------------------------------
  // Shared layout
  // ---------------------------------------------------------------------------

  function buildStyle(specOrId, accentHex) {
    var spec = typeof specOrId === 'string' ? templateSpec(specOrId) : specOrId;
    var accent = spec.accent || hexToRgb(accentHex);
    var s = {
      spec: spec,
      font: spec.bodyFont || 'helvetica',
      headingFont: spec.font || 'helvetica',
      text: [33, 37, 41],
      muted: [110, 117, 125],
      rule: spec.rule || [200, 204, 208],
      accent: accent,
      accentSoft: spec.accent === BLACK ? [235, 235, 235] : mix(accent, 0.88),
      borders: spec.borders !== false,
      zebra: !!spec.zebra
    };
    if (spec.head === 'accent') { s.tableHeadFill = accent; s.tableHeadText = WHITE; }
    else if (spec.head === 'dark') { s.tableHeadFill = spec.headDark || INK; s.tableHeadText = WHITE; }
    else if (spec.head === 'line') { s.tableHeadFill = null; s.tableHeadText = s.muted; }
    else { s.tableHeadFill = spec.headFill || [236, 238, 241]; s.tableHeadText = s.text; }
    return s;
  }

  function infoPairs(slip, settings, o) {
    var f = slip.fields;
    var pairs = [
      ['Employee ID', o.f_employeeId ? slip.employeeId : ''],
      ['Employee Name', slip.name],
      ['National ID / Iqama', o.f_nationalId ? f.nationalId : ''],
      ['Job Title', o.f_jobTitle ? f.jobTitle : ''],
      ['Department', o.f_department ? f.department : ''],
      ['Nationality', o.f_nationality ? f.nationality : ''],
      ['Joining Date', o.f_joiningDate ? f.joiningDate : ''],
      ['Date of Birth', o.f_dob ? f.dob : ''],
      ['Mobile', o.f_mobile && f.mobile ? '+' + f.mobile : ''],
      ['Email', o.f_email ? f.email : ''],
      ['Bank', o.f_bankName ? f.bankName : ''],
      ['IBAN', o.f_iban ? f.iban : ''],
      ['Working Days', o.f_workingDays ? f.workingDays : ''],
      ['Absent Days', o.f_absentDays ? f.absentDays : ''],
      ['Overtime Hours', o.f_overtimeHours ? f.overtimeHours : '']
    ];
    return pairs.filter(function (p) {
      return p[1] !== '' && p[1] != null;
    });
  }

  function renderPayslip(doc, ctx) {
    var spec = templateSpec(ctx.templateId || 'classic');
    var style = buildStyle(spec, ctx.accent);
    var P = new Painter(doc, style);
    var settings = ctx.settings || {};
    var slip = ctx.slip;
    var options = mergeOptions(ctx.options);
    var pageW = doc.internal.pageSize.getWidth();
    var pageH = doc.internal.pageSize.getHeight();
    var M = 14; // margin mm
    var L = M;
    var R = pageW - M;
    var currency = settings.currency || 'SAR';
    var radius = spec.radius || 0;
    var compact = !!spec.compact;
    var dashed = !!spec.dashed;
    var hv = spec.header;

    // ----- Page decorations -------------------------------------------------
    if (hv === 'sidebar') {
      P.rect(0, 0, 9, pageH, style.accent);
      L = 9 + 10;
    }
    if (hv === 'receipt') {
      L = (pageW - 124) / 2;
      R = L + 124;
    }
    if (spec.frame) {
      P.rect(6, 6, pageW - 12, pageH - 12, null, style.text, 0, 0.7);
      P.rect(8, 8, pageW - 16, pageH - 16, null, style.accent, 0, 0.3);
    }
    var contentW = R - L;
    var cx = (L + R) / 2;
    var y = M;
    if (hv === 'topbar') {
      P.rect(0, 0, pageW, 6, style.accent);
      y = M + 3;
    }
    if (spec.frame) y = M + 2;

    // ----- Header content -----------------------------------------------------
    var showRange = options.showPeriodDates && [settings.periodStart, settings.periodEnd].some(Boolean);
    var showPay = options.showPaymentDate && settings.paymentDate;
    var infoLines = [];
    if (options.showCr && settings.crNumber) infoLines.push('CR No. ' + settings.crNumber);
    if (options.showAddress && settings.address) infoLines.push(settings.address);
    var contact = [settings.phone, settings.email].filter(Boolean).join('  |  ');
    if (options.showContact && contact) infoLines.push(contact);
    var rightLines = [{ t: settings.payPeriod || '', size: 11, main: true }];
    if (showRange) rightLines.push({ t: [settings.periodStart, settings.periodEnd].filter(Boolean).join(' - '), size: 8.5 });
    if (showPay) rightLines.push({ t: 'Payment date: ' + settings.paymentDate, size: 8.5 });
    var title = String(options.title).toUpperCase();
    var titleSize = spec.titleSize || 20;
    var companyName = settings.companyName || 'Company Name';
    var hasLogo = !!(ctx.logo && ctx.logo.dataUrl);

    function addLogo(x, y0, maxW, maxH) {
      var dim = fitImage(ctx.logo, maxW, maxH);
      try {
        doc.addImage(ctx.logo.dataUrl, imageFormat(ctx.logo.dataUrl), x, y0, dim.w, dim.h, 'company-logo', 'FAST');
        return dim;
      } catch (e) {
        return null;
      }
    }

    // logo + company + CR/address/contact, left aligned; returns the bottom y
    function leftBlock(x, y0, maxW, cText, cMuted) {
      var tx = x;
      var logoBottom = 0;
      if (hasLogo) {
        var dim = addLogo(x, y0, 38, 20);
        if (dim) {
          tx = x + dim.w + 5;
          maxW -= dim.w + 5;
          logoBottom = y0 + dim.h;
        }
      }
      var n = P.text(companyName, tx, y0 + 6, { size: 14, weight: 'bold', font: style.headingFont, color: cText, maxWidth: maxW, maxLines: 2 });
      var ly = y0 + 11.5 + (n - 1) * 6;
      infoLines.forEach(function (l) {
        var k = P.text(l, tx, ly, { size: 8.5, color: cMuted, maxWidth: maxW, maxLines: 2 });
        ly += 4.2 * k;
      });
      return Math.max(ly, logoBottom);
    }

    // title + period lines, right aligned; returns the bottom y
    function rightBlock(xr, y0, cTitle, cText, cMuted) {
      P.text(title, xr, y0 + 8 + (titleSize - 20) * 0.25, { size: titleSize, weight: 'bold', font: style.headingFont, align: 'right', color: cTitle, maxWidth: contentW * 0.38, truncate: true });
      var ry = y0 + 14.5 + (titleSize - 20) * 0.35;
      rightLines.forEach(function (r, i) {
        P.text(r.t, xr, ry, { size: r.size, align: 'right', color: r.main ? cText : cMuted });
        ry += i === 0 ? 5 : 4.2;
      });
      return ry;
    }

    var estLeft = 11.5 + 4.2 * infoLines.length;
    var estRight = 14.5 + (titleSize - 20) * 0.35 + 5 + 4.2 * Math.max(0, rightLines.length - 2);
    var estLogo = hasLogo ? fitImage(ctx.logo, 38, 20).h : 0;
    var headH = Math.max(estLeft, estRight, estLogo) + 4;

    if (hv === 'band' || hv === 'split') {
      var fill = spec.headerFill || style.accent;
      var bandH = Math.max(38, M + headH + 4);
      var light = mix(fill, 0.75);
      if (hv === 'band') P.rect(0, 0, pageW, bandH, fill);
      else P.rect(0, 0, pageW * 0.56, bandH, fill);
      if (spec.goldLine) P.rect(0, bandH - 1.2, pageW, 1.2, style.accent);
      leftBlock(L, M, hv === 'split' ? pageW * 0.56 - L - 6 : contentW * 0.62, WHITE, light);
      if (hv === 'band') rightBlock(R, M, spec.goldLine ? style.accent : WHITE, WHITE, light);
      else rightBlock(R, M, style.accent, style.text, style.muted);
      y = bandH + 8;
    } else if (hv === 'boxed') {
      var boxH = headH + 6;
      P.rect(L, y, contentW, boxH, null, style.text, 0, 0.4);
      leftBlock(L + 4, y + 2, contentW * 0.6, style.text, style.muted);
      rightBlock(R - 4, y + 2, style.accent, style.text, style.muted);
      y += boxH + 7;
    } else if (hv === 'centered' || hv === 'receipt') {
      var yy = y;
      if (hasLogo && hv === 'centered') {
        var cd = fitImage(ctx.logo, 34, 16);
        var placed = addLogo(cx - cd.w / 2, yy, 34, 16);
        if (placed) yy += placed.h + 3;
      }
      var cn = P.text(companyName, cx, yy + 5, { size: hv === 'receipt' ? 12 : 15, weight: 'bold', font: style.headingFont, align: 'center', maxWidth: contentW, maxLines: 2 });
      yy += 5 + cn * 6 - 1;
      if (infoLines.length) {
        var ck = P.text(infoLines.join('  ·  '), cx, yy, { size: 8, color: style.muted, align: 'center', maxWidth: contentW, maxLines: 3 });
        yy += ck * 4 + 1;
      }
      if (hv === 'receipt') {
        P.line(L, yy + 1, R, yy + 1, style.text, 0.3, true);
        yy += 9;
      } else {
        P.line(cx - 30, yy + 1, cx + 30, yy + 1, style.accent, 0.5);
        yy += 10;
      }
      P.text(title, cx, yy, { size: hv === 'receipt' ? 16 : 18, weight: 'bold', font: style.headingFont, align: 'center', color: hv === 'receipt' ? style.text : style.accent, maxWidth: contentW, truncate: true });
      yy += 6;
      P.text(rightLines.map(function (r) { return r.t; }).filter(Boolean).join('  ·  '), cx, yy, { size: 9, align: 'center', maxWidth: contentW, maxLines: 2 });
      y = yy + 6;
      if (hv === 'receipt') P.line(L, y, R, y, style.text, 0.3, true);
      else P.line(L, y, R, y, style.rule, 0.3);
      y += 7;
    } else {
      // rule, sidebar, topbar (also used inside the heritage frame)
      var lb = leftBlock(L, y, contentW * 0.62, style.text, style.muted);
      var rb = rightBlock(R, y, style.accent, style.text, style.muted);
      y = Math.max(y + 30, lb + 2, rb + 2);
      P.line(L, y, R, y, style.accent, spec.ruleW || (hv === 'topbar' ? 0.8 : 0.3));
      y += 7;
    }

    // ----- Employee details ----------------------------------------------
    var pairs = infoPairs(slip, settings, options);
    var cols = hv === 'receipt' ? 2 : 3;
    var colW = contentW / cols;
    var rowH = compact ? 9.5 : 10.5;
    var rows = Math.ceil(pairs.length / cols);
    var boxH2 = rows * rowH + 4;
    var info = spec.info;
    if (info === 'soft') P.rect(L, y, contentW, boxH2, style.accentSoft, null, radius || 1.5);
    else if (info === 'outline') P.rect(L, y, contentW, boxH2, null, style.rule, radius);
    else if (info === 'lines') { P.line(L, y, R, y); P.line(L, y + boxH2, R, y + boxH2); }
    else if (info === 'accentLeft') { P.rect(L, y, contentW, boxH2, style.accentSoft); P.rect(L, y, 1.6, boxH2, style.accent); }
    var padX = info === 'accentLeft' ? 5 : info === 'plain' ? 0 : 3;
    pairs.forEach(function (p, i) {
      var c = i % cols;
      var r = Math.floor(i / cols);
      var x = L + c * colW + padX;
      var yy2 = y + 4 + r * rowH;
      P.text(p[0].toUpperCase(), x, yy2 + 2.5, { size: 6.5, color: style.muted });
      P.text(String(p[1]), x, yy2 + 7, { size: 9.5, weight: p[0] === 'Employee Name' ? 'bold' : 'normal', maxWidth: colW - 6, truncate: true });
    });
    y += boxH2 + (compact ? 6 : 8);
    if (dashed) {
      P.line(L, y - 3, R, y - 3, style.text, 0.3, true);
      y += 3;
    }

    // ----- Earnings / Deductions tables -----------------------------------
    var earnings = slip.earnings.length ? slip.earnings : [{ label: 'Basic Salary', amount: 0 }];
    var deductions = slip.deductions.length ? slip.deductions : [{ label: 'No deductions', amount: 0, placeholder: true }];
    var lineH = compact ? 6.2 : 7;
    var headRowH = compact ? 7 : 8;
    var zebraFill = spec.accent === BLACK ? [246, 246, 246] : mix(style.accent, 0.95);

    function drawTable(x, ty, w, heading, lines, nRows, total, totalLabel) {
      if (style.tableHeadFill) P.rect(x, ty, w, headRowH, style.tableHeadFill, style.borders ? style.rule : null);
      else P.line(x, ty + headRowH, x + w, ty + headRowH, style.rule, 0.3, dashed);
      P.text(heading.toUpperCase(), x + 3, ty + headRowH - 2.6, { size: 8, weight: 'bold', color: style.tableHeadText });
      P.text('AMOUNT (' + currency + ')', x + w - 3, ty + headRowH - 2.6, { size: 7, weight: 'bold', align: 'right', color: style.tableHeadText });
      ty += headRowH;
      for (var i = 0; i < nRows; i++) {
        var l = lines[i];
        if (style.zebra && i % 2 === 1) P.rect(x, ty, w, lineH, zebraFill);
        if (style.borders) P.rect(x, ty, w, lineH, null, style.rule);
        if (l) {
          P.text(l.label, x + 3, ty + lineH - 2.2, { size: 9, color: l.placeholder ? style.muted : style.text, maxWidth: w - 32, truncate: true });
          if (!l.placeholder) P.text(fmt(l.amount), x + w - 3, ty + lineH - 2.2, { size: 9, align: 'right' });
        }
        ty += lineH;
      }
      var tColor = style.text;
      if (spec.totals === 'accent') {
        P.rect(x, ty, w, lineH, style.accent);
        tColor = WHITE;
      } else if (spec.totals === 'line') {
        P.line(x, ty, x + w, ty, style.text, 0.3, dashed);
      } else {
        P.rect(x, ty, w, lineH, style.accentSoft, style.borders ? style.rule : null);
      }
      P.text(totalLabel, x + 3, ty + lineH - 2.2, { size: 9, weight: 'bold', color: tColor });
      P.text(fmt(total), x + w - 3, ty + lineH - 2.2, { size: 9, weight: 'bold', align: 'right', color: tColor });
      return ty + lineH;
    }

    if (spec.tables === 'stacked') {
      y = drawTable(L, y, contentW, 'Earnings', earnings, earnings.length, slip.gross, 'Total Earnings') + (compact ? 4 : 5);
      y = drawTable(L, y, contentW, 'Deductions', deductions, deductions.length, slip.totalDeductions, 'Total Deductions') + 8;
    } else {
      var gap = 6;
      var tableW = (contentW - gap) / 2;
      var nRows = Math.max(earnings.length, deductions.length);
      drawTable(L, y, tableW, 'Earnings', earnings, nRows, slip.gross, 'Total Earnings');
      y = drawTable(L + tableW + gap, y, tableW, 'Deductions', deductions, nRows, slip.totalDeductions, 'Total Deductions') + 8;
    }

    // ----- Net pay --------------------------------------------------------
    var words = options.amountInWords !== false;
    var netH = words ? 20 : 14;
    var ns = spec.net;
    var pad = 5;
    var cLabel = style.text;
    var cAmt = style.text;
    var cWords = style.muted;
    if (ns === 'band' || ns === 'pill') {
      P.rect(L, y, contentW, netH, style.accent, null, ns === 'pill' ? netH / 2 : radius);
      cLabel = cAmt = WHITE;
      cWords = mix(style.accent, 0.85);
      pad = ns === 'pill' ? 9 : 5;
    } else if (ns === 'dark') {
      var df = spec.darkFill || INK;
      P.rect(L, y, contentW, netH, df, null, radius);
      cLabel = WHITE;
      cAmt = spec.goldLine ? style.accent : WHITE;
      cWords = mix(df, 0.7);
    } else if (ns === 'box') {
      P.rect(L, y, contentW, netH, style.accentSoft, style.text, radius);
      pad = 4;
    } else if (ns === 'outline') {
      P.rect(L, y, contentW, netH, null, style.accent, radius || 1.5, 0.7);
      cAmt = style.accent;
    } else {
      P.line(L, y, R, y, style.text, 0.5, dashed);
      P.line(L, y + netH, R, y + netH, style.text, 0.5, dashed);
      cAmt = style.accent;
      pad = 4;
    }
    P.text('NET PAY', L + pad, y + 8, { size: 9.5, weight: 'bold', color: cLabel });
    P.text(currency + ' ' + fmt(slip.net), R - pad, y + 9, { size: 16, weight: 'bold', align: 'right', color: cAmt });
    if (words) P.text(Schema.amountToWords(slip.net, currency), L + pad, y + 15.5, { size: 8, color: cWords, maxWidth: contentW - pad * 2, maxLines: 1 });
    y += netH + 6;

    if (options.showRemarks && slip.fields.remarks) {
      var rn = P.text('Remarks: ' + slip.fields.remarks, L, y + 3, { size: 8.5, color: style.muted, maxWidth: contentW, maxLines: 3 });
      y += rn * 4.5 + 4;
    }

    // ----- Footer ---------------------------------------------------------
    var footerTop = pageH - M - 30;
    if (y > footerTop) footerTop = y; // never overlap; tall slips just push the footer down
    var sigW = Math.min(55, contentW * 0.42);

    if (options.signatureLines) {
      var sigY = footerTop + 14;
      P.line(L, sigY, L + sigW, sigY, style.text, 0.3);
      P.text('Employer signature / stamp', L, sigY + 4, { size: 7.5, color: style.muted });
      P.line(R - sigW, sigY, R, sigY, style.text, 0.3);
      P.text('Employee signature', R, sigY + 4, { size: 7.5, color: style.muted, align: 'right' });
      if (ctx.stamp && ctx.stamp.dataUrl) {
        var sd = fitImage(ctx.stamp, 40, 20);
        try {
          doc.addImage(ctx.stamp.dataUrl, imageFormat(ctx.stamp.dataUrl), L + 2, sigY - sd.h - 1, sd.w, sd.h, 'company-stamp', 'FAST');
        } catch (e) { /* ignore broken image */ }
      }
    } else if (ctx.stamp && ctx.stamp.dataUrl) {
      var sd2 = fitImage(ctx.stamp, 40, 20);
      try {
        doc.addImage(ctx.stamp.dataUrl, imageFormat(ctx.stamp.dataUrl), R - sd2.w, footerTop - 2, sd2.w, sd2.h, 'company-stamp', 'FAST');
      } catch (e) { /* ignore */ }
    }

    var bottomY = pageH - M;
    P.line(L, bottomY - 9, R, bottomY - 9, style.rule, 0.2, dashed);
    var footLeft = [];
    if (options.showFooterNote && settings.footerNote) footLeft.push(settings.footerNote);
    if (options.showPreparedBy && settings.preparedBy) footLeft.push('Prepared by: ' + settings.preparedBy);
    if (options.showGeneratedOn) footLeft.push('Generated on ' + Schema.formatDate(new Date()));
    if (footLeft.length) {
      P.text(footLeft.join('   ·   '), L, bottomY - 4.5, { size: 7.5, color: style.muted, maxWidth: Math.max(40, contentW - 60), maxLines: 2 });
    }
    P.text([slip.employeeId, settings.payPeriod].filter(Boolean).join(' · '), R, bottomY - 4.5, { size: 7, color: style.muted, align: 'right', maxWidth: 58, truncate: true });
  }

  // ---------------------------------------------------------------------------
  // Template thumbnails (small SVG built from the same spec as the PDF)
  // ---------------------------------------------------------------------------
  function rgbHex(c) {
    return '#' + c.map(function (v) { return ('0' + Math.max(0, Math.min(255, Math.round(v))).toString(16)).slice(-2); }).join('');
  }

  function thumbSvg(id, accentHex) {
    var spec = templateSpec(id);
    var accRgb = spec.accent || hexToRgb(accentHex);
    var a = rgbHex(accRgb);
    var soft = rgbHex(spec.accent === BLACK ? [232, 232, 232] : mix(accRgb, 0.85));
    var hf = rgbHex(spec.headerFill || accRgb);
    var dark = '#334155';
    var light = '#cbd5e1';
    var faint = '#e2e8f0';
    var L = 6;
    var R = 54;
    var hv = spec.header;
    if (hv === 'sidebar') L = 9;
    if (hv === 'receipt') { L = 15; R = 45; }
    var W = R - L;
    var out = [];
    function rect(x, y, w, h, fill, stroke, rx, dash) {
      out.push('<rect x="' + x + '" y="' + y + '" width="' + w + '" height="' + h + '"' + (rx ? ' rx="' + rx + '"' : '') +
        ' fill="' + (fill || 'none') + '"' + (stroke ? ' stroke="' + stroke + '" stroke-width="0.6"' + (dash ? ' stroke-dasharray="1.2 1"' : '') : '') + '/>');
    }
    function line(x1, y1, x2, y2, c, w, dash) {
      out.push('<line x1="' + x1 + '" y1="' + y1 + '" x2="' + x2 + '" y2="' + y2 + '" stroke="' + c + '" stroke-width="' + (w || 0.6) + '"' + (dash ? ' stroke-dasharray="1.2 1"' : '') + '/>');
    }
    rect(0, 0, 60, 76, '#ffffff');
    if (spec.frame) { rect(2, 2, 56, 72, null, dark); rect(3.4, 3.4, 53.2, 69.2, null, a); }
    if (hv === 'sidebar') rect(0, 0, 4, 76, a);
    if (hv === 'topbar') rect(0, 0, 60, 3, a);
    var y = 6;
    if (hv === 'band' || hv === 'split') {
      rect(0, 0, hv === 'band' ? 60 : 33, 15, hf);
      if (spec.goldLine) rect(0, 14.2, 60, 0.8, a);
      rect(L, 5, 16, 2.6, '#ffffff', null, 0.6);
      rect(L, 9.4, 11, 1.4, 'rgba(255,255,255,.65)', null, 0.5);
      rect(R - 13, 5, 13, 3.2, hv === 'band' ? (spec.goldLine ? a : '#ffffff') : a, null, 0.6);
      rect(R - 9, 9.8, 9, 1.4, hv === 'band' ? 'rgba(255,255,255,.65)' : light, null, 0.5);
      y = 19;
    } else if (hv === 'boxed') {
      rect(L, 4, W, 12, null, dark);
      rect(L + 2, 6.5, 15, 2.4, dark, null, 0.5);
      rect(L + 2, 10.5, 10, 1.4, light, null, 0.5);
      rect(R - 15, 6.5, 13, 3, dark, null, 0.5);
      y = 19;
    } else if (hv === 'centered' || hv === 'receipt') {
      rect(30 - 9, y, 18, 2.4, dark, null, 0.5);
      rect(30 - 12, y + 4, 24, 1.3, light, null, 0.5);
      if (hv === 'receipt') line(L, y + 7.5, R, y + 7.5, dark, 0.5, true);
      else line(24, y + 7.5, 36, y + 7.5, a, 0.7);
      rect(30 - 8, y + 10, 16, 2.8, hv === 'receipt' ? dark : a, null, 0.5);
      if (hv === 'receipt') line(L, y + 16, R, y + 16, dark, 0.5, true);
      else line(L, y + 16, R, y + 16, faint, 0.5);
      y += 19;
    } else {
      rect(L, y, 16, 2.6, dark, null, 0.6);
      rect(L, y + 4.2, 11, 1.4, light, null, 0.5);
      rect(R - 13, y, 13, spec.titleSize ? 4 : 3.2, a, null, 0.6);
      line(L, y + 9, R, y + 9, a, spec.ruleW || hv === 'topbar' ? 1 : 0.6);
      y += 13;
    }
    // employee details
    var ih = 9;
    if (spec.info === 'soft') rect(L, y, W, ih, soft, null, 1.2);
    else if (spec.info === 'outline') rect(L, y, W, ih, null, light, spec.radius ? 1 : 0);
    else if (spec.info === 'lines') { line(L, y, R, y, light); line(L, y + ih, R, y + ih, light); }
    else if (spec.info === 'accentLeft') { rect(L, y, W, ih, soft); rect(L, y, 1.2, ih, a); }
    var cols = hv === 'receipt' ? 2 : 3;
    for (var c = 0; c < cols; c++) {
      var x = L + 2 + c * (W / cols);
      rect(x, y + 2.4, W / cols - 5, 1.3, light, null, 0.4);
      rect(x, y + 5.2, W / cols - 7, 1.6, dark, null, 0.4);
    }
    y += ih + 4;
    // tables
    var headFill = spec.head === 'accent' ? a : spec.head === 'dark' ? rgbHex(spec.headDark || INK) : spec.head === 'grey' ? '#e5e7eb' : null;
    function table(x, ty, w, n) {
      if (headFill) rect(x, ty, w, 3, headFill);
      else line(x, ty + 3, x + w, ty + 3, light, 0.5, spec.dashed);
      for (var i = 0; i < n; i++) {
        var ry = ty + 3 + i * 2.6;
        if (spec.zebra && i % 2) rect(x, ry, w, 2.6, '#f1f5f9');
        if (spec.borders !== false) rect(x, ry, w, 2.6, null, faint);
        rect(x + 1.2, ry + 0.9, w * 0.45, 0.9, light, null, 0.3);
        rect(x + w - 1.2 - w * 0.2, ry + 0.9, w * 0.2, 0.9, light, null, 0.3);
      }
      var tt = ty + 3 + n * 2.6;
      if (spec.totals === 'accent') rect(x, tt, w, 2.6, a);
      else if (spec.totals === 'line') line(x, tt, x + w, tt, dark, 0.5, spec.dashed);
      else rect(x, tt, w, 2.6, soft);
      return tt + 2.6;
    }
    if (spec.tables === 'stacked') {
      y = table(L, y, W, 3) + 2;
      y = table(L, y, W, 2) + 3;
    } else {
      var tw = (W - 3) / 2;
      table(L, y, tw, 4);
      y = table(L + tw + 3, y, tw, 4) + 3;
    }
    // net pay
    var nh = 6;
    var ns = spec.net;
    if (ns === 'band') rect(L, y, W, nh, a, null, spec.radius ? 1 : 0);
    else if (ns === 'pill') rect(L, y, W, nh, a, null, 3);
    else if (ns === 'dark') rect(L, y, W, nh, rgbHex(spec.darkFill || INK));
    else if (ns === 'box') rect(L, y, W, nh, soft, dark);
    else if (ns === 'outline') rect(L, y, W, nh, null, a, 1);
    else { line(L, y, R, y, dark, 0.7, spec.dashed); line(L, y + nh, R, y + nh, dark, 0.7, spec.dashed); }
    var amtColor = (ns === 'band' || ns === 'pill' || ns === 'dark') ? (spec.goldLine ? a : '#ffffff') : (ns === 'outline' || ns === 'lines' ? a : dark);
    rect(R - 2 - W * 0.32, y + 2, W * 0.32, 2, amtColor, null, 0.5);
    rect(L + 2, y + 2.4, W * 0.18, 1.3, (ns === 'band' || ns === 'pill' || ns === 'dark') ? 'rgba(255,255,255,.75)' : light, null, 0.4);
    // footer
    line(L, 70, R, 70, faint, 0.5, spec.dashed);
    rect(L, 71.8, W * 0.4, 1, light, null, 0.3);
    return '<svg viewBox="0 0 60 76" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">' + out.join('') + '</svg>';
  }

  // ---------------------------------------------------------------------------
  // Document helpers
  // ---------------------------------------------------------------------------

  /**
   * createDoc(jsPDF, paper, { password }) - with a password the PDF is encrypted with jsPDF's
   * built-in standard security handler (40-bit RC4): it stops casual opening and works in every
   * viewer, but it is not strong encryption. Printing stays allowed; copying/modifying is not.
   */
  function createDoc(jsPDF, paper, opts) {
    var p = PAPERS[paper] || PAPERS.a4;
    var init = { unit: 'mm', format: p.format, compress: true };
    if (opts && opts.password) {
      init.encryption = {
        userPassword: String(opts.password),
        ownerPassword: opts.ownerPassword || randomPassword(),
        userPermissions: ['print']
      };
    }
    return new jsPDF(init);
  }

  function randomPassword() {
    var chars = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghjkmnpqrstuvwxyz23456789';
    var s = '';
    for (var i = 0; i < 20; i++) s += chars.charAt(Math.floor(Math.random() * chars.length));
    return s;
  }

  // ---------------------------------------------------------------------------
  // Acknowledgement register: a printable table (one block per department) that employees
  // sign when they receive their payslip. Proof of delivery for paper hand-outs.
  // ---------------------------------------------------------------------------
  function renderRegister(doc, ctx) {
    var settings = ctx.settings || {};
    var slips = ctx.slips || [];
    var style = buildStyle('classic', '#000000');
    var P = new Painter(doc, style);
    var pageW = doc.internal.pageSize.getWidth();
    var pageH = doc.internal.pageSize.getHeight();
    var M = 12;
    var cols = [
      { label: '#', w: 9, align: 'left' },
      { label: 'Employee ID', w: 26, align: 'left' },
      { label: 'Name', w: 54, align: 'left' },
      { label: 'National ID / Iqama', w: 30, align: 'left' },
      { label: 'Net (' + (settings.currency || 'SAR') + ')', w: 24, align: 'right' },
      { label: 'Signature', w: 0, align: 'left' },
      { label: 'Date', w: 20, align: 'left' }
    ];
    var fixed = cols.reduce(function (a, c) { return a + c.w; }, 0);
    cols[5].w = pageW - M * 2 - fixed;
    var rowH = 9;
    var groups = {};
    slips.forEach(function (s) {
      var k = s.fields.department || '';
      (groups[k] = groups[k] || []).push(s);
    });
    var keys = Object.keys(groups).sort();
    var page = 1;
    var y = 0;

    function header(dept) {
      y = M;
      P.text(String(settings.companyName || ''), M, y + 5, { size: 12, weight: 'bold' });
      P.text('Payslip acknowledgement register', pageW - M, y + 5, { size: 12, weight: 'bold', align: 'right' });
      P.text(['CR ' + (settings.crNumber || '-'), 'Pay period: ' + (settings.payPeriod || '-'), dept ? 'Department: ' + dept : ''].filter(Boolean).join('   ·   '), M, y + 10.5, { size: 8.5, color: style.muted });
      P.text('I confirm that I received my payslip for this period.', pageW - M, y + 10.5, { size: 8, color: style.muted, align: 'right' });
      y += 15;
      var x = M;
      P.rect(M, y, pageW - M * 2, 7, [230, 230, 230], style.rule);
      cols.forEach(function (c) {
        P.text(c.label, c.align === 'right' ? x + c.w - 2 : x + 2, y + 4.8, { size: 7.5, weight: 'bold', align: c.align });
        x += c.w;
      });
      y += 7;
    }

    function footer() {
      P.text('Page ' + page, pageW - M, pageH - 6, { size: 7.5, color: style.muted, align: 'right' });
      P.text('Generated ' + Schema.formatDate(new Date()) + ' · Signed register to be kept with the wage records.', M, pageH - 6, { size: 7.5, color: style.muted });
    }

    header('');
    keys.forEach(function (dept) {
      if (y + rowH * 2 + 6 > pageH - 14) {
        footer();
        doc.addPage();
        page++;
        header('');
      }
      if (keys.length > 1 || dept) {
        P.rect(M, y, pageW - M * 2, 6, [245, 245, 245], style.rule);
        P.text((dept || 'No department') + '  (' + groups[dept].length + ')', M + 2, y + 4.2, { size: 8, weight: 'bold' });
        y += 6;
      }
      groups[dept].forEach(function (s, i) {
        if (y + rowH > pageH - 14) {
          footer();
          doc.addPage();
          page++;
          header('');
        }
        var x = M;
        P.rect(M, y, pageW - M * 2, rowH, i % 2 ? [248, 248, 248] : null, style.rule);
        var vals = [String(i + 1), s.employeeId, s.name, s.fields.nationalId || '', fmt(s.net), '', ''];
        cols.forEach(function (c, ci) {
          if (vals[ci]) P.text(vals[ci], c.align === 'right' ? x + c.w - 2 : x + 2, y + 5.8, { size: 8.5, align: c.align, maxWidth: c.w - 4, truncate: true });
          x += c.w;
        });
        y += rowH;
      });
    });
    if (!keys.length) P.text('No employees.', M, y + 8, { size: 9, color: style.muted });
    footer();
  }

  function setMeta(doc, settings, title) {
    try {
      doc.setProperties({
        title: title,
        subject: 'Payslip ' + (settings.payPeriod || ''),
        author: settings.companyName || 'Payslip Generator',
        creator: 'Payslip Generator (offline)'
      });
    } catch (e) { /* ignore */ }
  }

  function safeFileName(s) {
    return String(s || '')
      .replace(/[\\/:*?"<>|]+/g, '-')
      .replace(/\s+/g, ' ')
      .trim()
      .slice(0, 80) || 'payslip';
  }

  return {
    TEMPLATES: TEMPLATES,
    templateSpec: templateSpec,
    thumbSvg: thumbSvg,
    PAPERS: PAPERS,
    DISPLAY_OPTIONS: DISPLAY_OPTIONS,
    hasArabic: hasArabic,
    arabicFontAvailable: arabicFontAvailable,
    DEFAULT_OPTIONS: DEFAULT_OPTIONS,
    mergeOptions: mergeOptions,
    renderPayslip: renderPayslip,
    renderRegister: renderRegister,
    createDoc: createDoc,
    setMeta: setMeta,
    safeFileName: safeFileName,
    hexToRgb: hexToRgb
  };
});
