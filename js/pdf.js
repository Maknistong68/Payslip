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

  var TEMPLATES = [
    { id: 'classic', name: 'Classic', description: 'Bordered tables, black and white. Prints well on any printer.' },
    { id: 'modern', name: 'Modern', description: 'Coloured header band and highlighted net pay.' },
    { id: 'minimal', name: 'Minimal', description: 'Clean, lots of white space, thin lines.' }
  ];

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

  Painter.prototype.line = function (x1, y1, x2, y2, color, width) {
    var c = color || this.style.rule;
    this.doc.setDrawColor(c[0], c[1], c[2]);
    this.doc.setLineWidth(width || 0.2);
    this.doc.line(x1, y1, x2, y2);
  };

  Painter.prototype.rect = function (x, y, w, h, fill, stroke, radius) {
    var d = this.doc;
    var mode = '';
    if (fill) {
      d.setFillColor(fill[0], fill[1], fill[2]);
      mode += 'F';
    }
    if (stroke) {
      d.setDrawColor(stroke[0], stroke[1], stroke[2]);
      d.setLineWidth(0.2);
      mode += 'D';
    }
    if (!mode) return;
    if (radius) d.roundedRect(x, y, w, h, radius, radius, mode);
    else d.rect(x, y, w, h, mode);
  };

  // ---------------------------------------------------------------------------
  // Shared layout
  // ---------------------------------------------------------------------------

  function buildStyle(templateId, accentHex) {
    var accent = hexToRgb(accentHex);
    var base = {
      font: 'helvetica',
      headingFont: 'helvetica',
      text: [33, 37, 41],
      muted: [110, 117, 125],
      rule: [200, 204, 208],
      accent: accent,
      accentSoft: mix(accent, 0.88),
      tableHeadFill: [240, 242, 244],
      tableHeadText: [33, 37, 41],
      borders: true,
      zebra: false
    };
    if (templateId === 'classic') {
      base.headingFont = 'times';
      base.accent = [0, 0, 0];
      base.accentSoft = [235, 235, 235];
      base.rule = [120, 120, 120];
      base.tableHeadFill = [225, 225, 225];
    } else if (templateId === 'modern') {
      base.tableHeadFill = accent;
      base.tableHeadText = [255, 255, 255];
      base.zebra = true;
    } else if (templateId === 'minimal') {
      base.borders = false;
      base.rule = [222, 226, 230];
      base.tableHeadFill = null;
      base.tableHeadText = base.muted;
    }
    return base;
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
    var templateId = ctx.templateId || 'classic';
    var style = buildStyle(templateId, ctx.accent);
    var P = new Painter(doc, style);
    var settings = ctx.settings || {};
    var slip = ctx.slip;
    var options = mergeOptions(ctx.options);
    var pageW = doc.internal.pageSize.getWidth();
    var pageH = doc.internal.pageSize.getHeight();
    var M = 14; // margin mm
    var contentW = pageW - M * 2;
    var currency = settings.currency || 'SAR';
    var y = M;

    // ----- Header ---------------------------------------------------------
    var headerH = 30;
    var showRange = options.showPeriodDates && [settings.periodStart, settings.periodEnd].some(Boolean);
    var showPay = options.showPaymentDate && settings.paymentDate;
    var rightLines = (showRange ? 1 : 0) + (showPay ? 1 : 0);
    var leftLines = (options.showCr && settings.crNumber ? 1 : 0) + (options.showAddress && settings.address ? 1 : 0) + (options.showContact && (settings.phone || settings.email) ? 1 : 0);
    var bandH = Math.max(headerH + 8, y + 19.5 + 4.2 * rightLines + 4, y + 11.5 + 4.2 * leftLines + 4);
    if (templateId === 'modern') {
      P.rect(0, 0, pageW, bandH, style.accent);
    }
    var headerTextColor = templateId === 'modern' ? [255, 255, 255] : style.text;
    var headerMuted = templateId === 'modern' ? mix(style.accent, 0.8) : style.muted;

    var textX = M;
    if (ctx.logo && ctx.logo.dataUrl) {
      var dim = fitImage(ctx.logo, 38, 20);
      try {
        doc.addImage(ctx.logo.dataUrl, imageFormat(ctx.logo.dataUrl), M, y, dim.w, dim.h, 'company-logo', 'FAST');
        textX = M + dim.w + 5;
      } catch (e) {
        textX = M;
      }
    }
    var rightColX = pageW - M;
    var leftColW = contentW * 0.62 - (textX - M);

    P.text(settings.companyName || 'Company Name', textX, y + 6, { size: 14, weight: 'bold', font: style.headingFont, color: headerTextColor, maxWidth: leftColW, maxLines: 2 });
    var ly = y + 11.5;
    var infoLines = [];
    if (options.showCr && settings.crNumber) infoLines.push('CR No. ' + settings.crNumber);
    if (options.showAddress && settings.address) infoLines.push(settings.address);
    var contact = [settings.phone, settings.email].filter(Boolean).join('  |  ');
    if (options.showContact && contact) infoLines.push(contact);
    infoLines.forEach(function (l) {
      var n = P.text(l, textX, ly, { size: 8.5, color: headerMuted, maxWidth: leftColW, maxLines: 2 });
      ly += 4.2 * n;
    });

    P.text(String(options.title).toUpperCase(), rightColX, y + 8, { size: 20, weight: 'bold', font: style.headingFont, align: 'right', color: templateId === 'modern' ? [255, 255, 255] : style.accent, maxWidth: contentW * 0.38, truncate: true });
    P.text(settings.payPeriod || '', rightColX, y + 14.5, { size: 11, align: 'right', color: headerTextColor });
    var subY = y + 19.5;
    var periodRange = [settings.periodStart, settings.periodEnd].filter(Boolean).join(' - ');
    if (showRange) {
      P.text(periodRange, rightColX, subY, { size: 8.5, align: 'right', color: headerMuted });
      subY += 4.2;
    }
    if (showPay) {
      P.text('Payment date: ' + settings.paymentDate, rightColX, subY, { size: 8.5, align: 'right', color: headerMuted });
      subY += 4.2;
    }

    y = Math.max(y + headerH, ly + 2, subY + 2);
    if (templateId === 'modern') y = bandH + 8;
    else {
      P.line(M, y, pageW - M, y, style.accent, templateId === 'classic' ? 0.6 : 0.3);
      y += 7;
    }

    // ----- Employee details ----------------------------------------------
    var pairs = infoPairs(slip, settings, options);
    var cols = 3;
    var colW = contentW / cols;
    var rowH = 10.5;
    var rows = Math.ceil(pairs.length / cols);
    var boxH = rows * rowH + 4;

    if (templateId === 'modern') {
      P.rect(M, y, contentW, boxH, style.accentSoft, null, 1.5);
    } else if (templateId === 'classic') {
      P.rect(M, y, contentW, boxH, null, style.rule);
    } else {
      P.line(M, y, pageW - M, y);
      P.line(M, y + boxH, pageW - M, y + boxH);
    }

    pairs.forEach(function (p, i) {
      var c = i % cols;
      var r = Math.floor(i / cols);
      var x = M + c * colW + 3;
      var yy = y + 4 + r * rowH;
      P.text(p[0].toUpperCase(), x, yy + 2.5, { size: 6.5, color: style.muted });
      P.text(String(p[1]), x, yy + 7, { size: 9.5, weight: p[0] === 'Employee Name' ? 'bold' : 'normal', maxWidth: colW - 6, truncate: true });
    });
    y += boxH + 8;

    // ----- Earnings / Deductions tables -----------------------------------
    var gap = 6;
    var tableW = (contentW - gap) / 2;
    var earnings = slip.earnings.length ? slip.earnings : [{ label: 'Basic Salary', amount: 0 }];
    var deductions = slip.deductions.length ? slip.deductions : [{ label: 'No deductions', amount: 0, placeholder: true }];
    var lineH = 7;
    var headH = 8;
    var nRows = Math.max(earnings.length, deductions.length);
    var tableH = headH + nRows * lineH + lineH; // + totals row

    function drawTable(x, title, lines, total, totalLabel) {
      var ty = y;
      // header
      if (style.tableHeadFill) {
        P.rect(x, ty, tableW, headH, style.tableHeadFill, style.borders ? style.rule : null);
      } else {
        P.line(x, ty + headH, x + tableW, ty + headH, style.rule, 0.3);
      }
      P.text(title.toUpperCase(), x + 3, ty + 5.4, { size: 8, weight: 'bold', color: style.tableHeadText });
      P.text('AMOUNT (' + currency + ')', x + tableW - 3, ty + 5.4, { size: 7, weight: 'bold', align: 'right', color: style.tableHeadText });
      ty += headH;
      for (var i = 0; i < nRows; i++) {
        var l = lines[i];
        if (style.zebra && i % 2 === 1) P.rect(x, ty, tableW, lineH, [248, 249, 250]);
        if (style.borders) P.rect(x, ty, tableW, lineH, null, style.rule);
        if (l) {
          P.text(l.label, x + 3, ty + 4.8, { size: 9, color: l.placeholder ? style.muted : style.text, maxWidth: tableW - 32, truncate: true });
          if (!l.placeholder) P.text(fmt(l.amount), x + tableW - 3, ty + 4.8, { size: 9, align: 'right' });
        }
        ty += lineH;
      }
      // totals
      P.rect(x, ty, tableW, lineH, templateId === 'minimal' ? null : style.accentSoft, style.borders ? style.rule : null);
      if (templateId === 'minimal') P.line(x, ty, x + tableW, ty, style.text, 0.3);
      P.text(totalLabel, x + 3, ty + 4.8, { size: 9, weight: 'bold' });
      P.text(fmt(total), x + tableW - 3, ty + 4.8, { size: 9, weight: 'bold', align: 'right' });
    }

    drawTable(M, 'Earnings', earnings, slip.gross, 'Total Earnings');
    drawTable(M + tableW + gap, 'Deductions', deductions, slip.totalDeductions, 'Total Deductions');
    y += tableH + 8;

    // ----- Net pay --------------------------------------------------------
    var netH = options.amountInWords !== false ? 20 : 14;
    if (templateId === 'modern') {
      P.rect(M, y, contentW, netH, style.accent, null, 1.5);
      P.text('NET PAY', M + 5, y + 8, { size: 9, weight: 'bold', color: [255, 255, 255] });
      P.text(currency + ' ' + fmt(slip.net), pageW - M - 5, y + 9, { size: 16, weight: 'bold', align: 'right', color: [255, 255, 255] });
      if (options.amountInWords !== false) {
        P.text(Schema.amountToWords(slip.net, currency), M + 5, y + 15.5, { size: 8, color: mix(style.accent, 0.85), maxWidth: contentW - 10, maxLines: 1 });
      }
    } else {
      P.rect(M, y, contentW, netH, templateId === 'classic' ? style.accentSoft : null, templateId === 'classic' ? style.text : null);
      if (templateId === 'minimal') {
        P.line(M, y, pageW - M, y, style.text, 0.5);
        P.line(M, y + netH, pageW - M, y + netH, style.text, 0.5);
      }
      P.text('NET PAY', M + 4, y + 8, { size: 10, weight: 'bold' });
      P.text(currency + ' ' + fmt(slip.net), pageW - M - 4, y + 9, { size: 16, weight: 'bold', align: 'right', color: templateId === 'minimal' ? style.accent : style.text });
      if (options.amountInWords !== false) {
        P.text(Schema.amountToWords(slip.net, currency), M + 4, y + 15.5, { size: 8, color: style.muted, maxWidth: contentW - 8, maxLines: 1 });
      }
    }
    y += netH + 6;

    if (options.showRemarks && slip.fields.remarks) {
      var n = P.text('Remarks: ' + slip.fields.remarks, M, y + 3, { size: 8.5, color: style.muted, maxWidth: contentW, maxLines: 3 });
      y += n * 4.5 + 4;
    }

    // ----- Footer ---------------------------------------------------------
    var footerTop = pageH - M - 30;
    if (y > footerTop) footerTop = y; // never overlap; tall slips just push the footer down

    if (options.signatureLines) {
      var sigY = footerTop + 14;
      var sigW = 55;
      P.line(M, sigY, M + sigW, sigY, style.text, 0.3);
      P.text('Employer signature / stamp', M, sigY + 4, { size: 7.5, color: style.muted });
      P.line(pageW - M - sigW, sigY, pageW - M, sigY, style.text, 0.3);
      P.text('Employee signature', pageW - M, sigY + 4, { size: 7.5, color: style.muted, align: 'right' });
      if (ctx.stamp && ctx.stamp.dataUrl) {
        var sd = fitImage(ctx.stamp, 40, 20);
        try {
          doc.addImage(ctx.stamp.dataUrl, imageFormat(ctx.stamp.dataUrl), M + 2, sigY - sd.h - 1, sd.w, sd.h, 'company-stamp', 'FAST');
        } catch (e) { /* ignore broken image */ }
      }
    } else if (ctx.stamp && ctx.stamp.dataUrl) {
      var sd2 = fitImage(ctx.stamp, 40, 20);
      try {
        doc.addImage(ctx.stamp.dataUrl, imageFormat(ctx.stamp.dataUrl), pageW - M - sd2.w, footerTop - 2, sd2.w, sd2.h, 'company-stamp', 'FAST');
      } catch (e) { /* ignore */ }
    }

    var bottomY = pageH - M;
    P.line(M, bottomY - 9, pageW - M, bottomY - 9);
    var footLeft = [];
    if (options.showFooterNote && settings.footerNote) footLeft.push(settings.footerNote);
    if (options.showPreparedBy && settings.preparedBy) footLeft.push('Prepared by: ' + settings.preparedBy);
    if (options.showGeneratedOn) footLeft.push('Generated on ' + Schema.formatDate(new Date()));
    if (footLeft.length) {
      P.text(footLeft.join('   ·   '), M, bottomY - 4.5, { size: 7.5, color: style.muted, maxWidth: contentW - 60, maxLines: 2 });
    }
    P.text([slip.employeeId, settings.payPeriod].filter(Boolean).join(' · '), pageW - M, bottomY - 4.5, { size: 7, color: style.muted, align: 'right', maxWidth: 58, truncate: true });
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
