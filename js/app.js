/* Payslip Generator - UI. Everything runs in the browser; no network calls. */
(function () {
  'use strict';

  var $ = function (id) { return document.getElementById(id); };

  if (typeof XLSX === 'undefined' || typeof jspdf === 'undefined' || typeof JSZip === 'undefined' ||
      typeof PayslipSchema === 'undefined' || typeof PayslipPDF === 'undefined' || typeof XlsxImages === 'undefined' ||
      typeof PayslipDelivery === 'undefined') {
    $('libError').classList.remove('hidden');
    return;
  }

  var Schema = PayslipSchema;
  var PDF = PayslipPDF;
  var Delivery = PayslipDelivery;
  var jsPDF = jspdf.jsPDF;
  var STORAGE_KEY = 'payslip-generator-profile-v2';
  var LOG_KEY_PREFIX = 'payslip-generator-log-';
  var SETTINGS_FILE_VERSION = 2;
  var EML_BATCH = 50;
  var Art = window.PayslipArt || { say: function () {}, confetti: function () {}, countUp: function (el, v) { el.textContent = v; }, init: function () {} };

  // ---------------------------------------------------------------------------
  // State
  // ---------------------------------------------------------------------------
  function defaultProfile() {
    return {
      templateId: 'classic',
      accent: '#1f4e79',
      paper: 'a4',
      display: PDF.mergeOptions({}),
      logo: null,
      stamp: null,
      company: {},
      columnOverrides: {},
      channel: 'outlook',
      savedAt: null
    };
  }

  var state = {
    step: 1,
    wb: null,
    parsed: null,
    fileName: '',
    company: {},
    companyFromExcel: {},
    excelImages: { logo: null, stamp: null },
    profile: defaultProfile(),
    previewUrl: null,
    cancelExport: false,
    exporting: false,
    cache: {},          // employeeId -> { bytes, fileName, password, hash }
    dirHandle: null,    // File System Access directory handle after a folder export
    exportedTo: '',     // human text: where the PDFs went
    log: []             // delivery log entries for the current company+period
  };

  // ---------------------------------------------------------------------------
  // Helpers
  // ---------------------------------------------------------------------------
  function escapeHtml(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }

  function downloadBlob(blob, filename) {
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    setTimeout(function () {
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    }, 2000);
  }

  function nextTick() {
    return new Promise(function (r) { setTimeout(r, 0); });
  }

  function plural(n, one, many) {
    return n + ' ' + (n === 1 ? one : many || one + 's');
  }

  function setAlert(el, html, kind) {
    el.innerHTML = html ? '<div class="alert alert-' + (kind || 'ok') + '">' + html + '</div>' : '';
  }

  function saveProfile() {
    state.profile.savedAt = new Date().toISOString();
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(state.profile));
    } catch (e) { /* storage unavailable or full */ }
  }

  function loadProfile() {
    try {
      var raw = localStorage.getItem(STORAGE_KEY);
      if (raw) applyProfile(JSON.parse(raw), { silent: true });
    } catch (e) { /* ignore */ }
  }

  function applyProfile(p, opts) {
    opts = opts || {};
    var base = defaultProfile();
    if (!p || typeof p !== 'object') return;
    ['templateId', 'accent', 'paper', 'channel'].forEach(function (k) { if (p[k]) base[k] = p[k]; });
    base.display = PDF.mergeOptions(p.display || {});
    if (p.logo && p.logo.dataUrl) base.logo = p.logo;
    if (p.stamp && p.stamp.dataUrl) base.stamp = p.stamp;
    if (p.company && typeof p.company === 'object') base.company = upgradeMessageText(p.company);
    if (p.columnOverrides && typeof p.columnOverrides === 'object') base.columnOverrides = p.columnOverrides;
    base.savedAt = p.savedAt || null;
    state.profile = base;
    if (!opts.silent) saveProfile();
  }

  function normalizeImage(src, preferJpeg) {
    return new Promise(function (resolve, reject) {
      var img = new Image();
      img.onerror = function () { reject(new Error('This image format cannot be opened by the browser. Use PNG or JPG.')); };
      img.onload = function () {
        var MAX = 900;
        var scale = Math.min(1, MAX / Math.max(img.naturalWidth, img.naturalHeight));
        var canvas = document.createElement('canvas');
        canvas.width = Math.max(1, Math.round(img.naturalWidth * scale));
        canvas.height = Math.max(1, Math.round(img.naturalHeight * scale));
        canvas.getContext('2d').drawImage(img, 0, 0, canvas.width, canvas.height);
        resolve({ dataUrl: preferJpeg ? canvas.toDataURL('image/jpeg', 0.9) : canvas.toDataURL('image/png'), width: canvas.width, height: canvas.height });
      };
      img.src = src;
    });
  }

  function readFile(file, as) {
    return new Promise(function (resolve, reject) {
      var reader = new FileReader();
      reader.onerror = function () { reject(new Error('Could not read the file.')); };
      reader.onload = function () { resolve(reader.result); };
      if (as === 'text') reader.readAsText(file);
      else if (as === 'dataUrl') reader.readAsDataURL(file);
      else reader.readAsArrayBuffer(file);
    });
  }

  function fileToImage(file) {
    return readFile(file, 'dataUrl').then(function (dataUrl) { return normalizeImage(dataUrl, /jpe?g/i.test(file.type)); });
  }

  function sha256Hex(bytes) {
    if (!(window.crypto && crypto.subtle)) return Promise.resolve('');
    return crypto.subtle.digest('SHA-256', bytes).then(function (buf) {
      return Array.prototype.map.call(new Uint8Array(buf), function (b) { return ('0' + b.toString(16)).slice(-2); }).join('');
    }).catch(function () { return ''; });
  }

  // ---------------------------------------------------------------------------
  // Steps
  // ---------------------------------------------------------------------------
  function goStep(n) {
    if (n > 1 && !state.parsed) return;
    var prev = state.step;
    state.step = n;
    document.body.setAttribute('data-step', String(n));
    for (var i = 1; i <= 5; i++) {
      var panel = $('panel' + i);
      panel.classList.toggle('hidden', i !== n);
      panel.classList.remove('enter-fwd', 'enter-back');
      if (i === n && prev !== n) {
        void panel.offsetWidth; // restart the slide animation
        panel.classList.add(n >= prev ? 'enter-fwd' : 'enter-back');
      }
    }
    buddyForStep(n);
    Array.prototype.forEach.call($('stepper').children, function (li) {
      var s = +li.getAttribute('data-step');
      li.classList.toggle('active', s === n);
      li.classList.toggle('done', s < n && !!state.parsed);
      li.classList.toggle('disabled', s > 1 && !state.parsed);
    });
    window.scrollTo({ top: 0, behavior: 'smooth' });
    if (n === 3) schedulePreview();
    if (n === 4) renderExportPanel();
    if (n === 5) renderSendPanel();
  }

  // The mascot's line for each step.
  function buddyForStep(n) {
    if (!state.parsed) { Art.say('idle', null); return; }
    var st = state.parsed.stats;
    if (n === 2) {
      var errs = state.parsed.issues.filter(function (i) { return i.level === 'error'; }).length;
      if (errs || st.withErrors) Art.say('oops', 'I checked every row. <b>' + (errs ? errs + ' file problem' + (errs === 1 ? '' : 's') : st.withErrors + ' row' + (st.withErrors === 1 ? '' : 's')) + '</b> need a look; the rest are fine.');
      else if (companyMissing().length) Art.say('look', '<b>' + st.ok + ' payslips</b> look good. Just tell me the company details below.');
      else Art.say('look', 'All <b>' + st.ok + ' payslips</b> check out. Total net <b>' + Schema.formatMoney(st.totalNet) + ' ' + escapeHtml(state.company.currency || 'SAR') + '</b>.');
    } else if (n === 3) {
      Art.say('design', 'Pick a look. I remember it for <b>next month</b>.');
    } else if (n === 4) {
      Art.say('export', 'Each PDF gets locked with the employee\'s <b>Iqama number</b>.', { settle: 'idle', after: 3500 });
    } else if (n === 5) {
      var left = exportableSlips().filter(function (s) { return s.fields.email && !isSent(s); }).length;
      Art.say('send', left ? 'Press <b>Next draft</b> (or Enter). Outlook opens it, you press Send. <b>' + left + '</b> to go.' : 'Everyone has their payslip. <b>Nice work!</b>');
    }
  }

  Array.prototype.forEach.call(document.querySelectorAll('[data-go]'), function (btn) {
    btn.addEventListener('click', function () { goStep(+btn.getAttribute('data-go')); });
  });
  Array.prototype.forEach.call($('stepper').children, function (li) {
    li.addEventListener('click', function () { goStep(+li.getAttribute('data-step')); });
  });
  $('brandLink').addEventListener('click', function (e) { e.preventDefault(); goStep(1); });
  $('btnChangeFile').addEventListener('click', function () { goStep(1); });
  $('btnStartOver').addEventListener('click', function () { goStep(1); });

  // ---------------------------------------------------------------------------
  // Step 1: template downloads, upload
  // ---------------------------------------------------------------------------
  var XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

  function workbookToArray(wb) {
    return XLSX.write(wb, { type: 'array', bookType: 'xlsx', cellDates: true });
  }

  $('btnTemplate').addEventListener('click', function () {
    downloadBlob(new Blob([workbookToArray(Schema.buildTemplateWorkbook(XLSX))], { type: XLSX_MIME }), 'Payslip_Template.xlsx');
  });
  $('btnSampleXlsx').addEventListener('click', function () {
    downloadBlob(new Blob([workbookToArray(Schema.buildTemplateWorkbook(XLSX, { sampleRows: Schema.sampleRows(25) }))], { type: XLSX_MIME }), 'Sample_Payslips.xlsx');
  });
  $('btnLoadSample').addEventListener('click', function () {
    loadWorkbookData(workbookToArray(Schema.buildTemplateWorkbook(XLSX, { sampleRows: Schema.sampleRows(25) })), 'Sample data (25 employees)');
  });

  $('btnMyTemplate').addEventListener('click', function () {
    var buf = workbookToArray(Schema.buildTemplateWorkbook(XLSX, { settings: state.company }));
    var images = [];
    var logo = effectiveImage('logo');
    var stamp = effectiveImage('stamp');
    if (logo) images.push({ dataUrl: logo.dataUrl, width: logo.width, height: logo.height, name: 'Logo', row: 1, col: 4, maxPx: 220 });
    if (stamp) images.push({ dataUrl: stamp.dataUrl, width: stamp.width, height: stamp.height, name: 'Stamp', row: 8, col: 4, maxPx: 160 });
    XlsxImages.embedImages(JSZip, buf, Schema.SHEET_SETTINGS, images).then(function (out) {
      downloadBlob(new Blob([out], { type: XLSX_MIME }), PDF.safeFileName('Payslip_Template - ' + (state.company.companyName || 'my company')) + '.xlsx');
    }).catch(function (e) { alert('Could not build the template: ' + e.message); });
  });

  var dropzone = $('dropzone');
  var fileInput = $('fileInput');
  ['dragenter', 'dragover'].forEach(function (ev) {
    dropzone.addEventListener(ev, function (e) { e.preventDefault(); dropzone.classList.add('drag'); });
  });
  ['dragleave', 'drop'].forEach(function (ev) {
    dropzone.addEventListener(ev, function (e) { e.preventDefault(); dropzone.classList.remove('drag'); });
  });
  dropzone.addEventListener('drop', function (e) {
    var f = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
    if (f) handleFile(f);
  });
  fileInput.addEventListener('change', function () {
    if (fileInput.files && fileInput.files[0]) handleFile(fileInput.files[0]);
    fileInput.value = '';
  });

  function handleFile(file) {
    if (/\.json$/i.test(file.name)) return loadSettingsFile(file);
    readFile(file).then(function (buf) { loadWorkbookData(buf, file.name); })
      .catch(function (e) { setAlert($('loadError'), escapeHtml(e.message), 'error'); });
  }

  function loadWorkbookData(arrayBuffer, name) {
    $('loadError').innerHTML = '';
    var wb;
    try {
      wb = XLSX.read(arrayBuffer, { type: 'array', cellDates: true });
    } catch (e) {
      setAlert($('loadError'), 'This file could not be opened as a spreadsheet. ' + escapeHtml(e.message || ''), 'error');
      return;
    }
    state.wb = wb;
    state.fileName = name;
    state.excelImages = { logo: null, stamp: null };
    state.dirHandle = null;
    state.exportedTo = '';
    $('exportMessage').innerHTML = '';
    $('sendMessage').innerHTML = '';
    reparse();
    XlsxImages.extractImages(JSZip, arrayBuffer).then(function (images) {
      if (!images.length) return;
      var onSettings = images.filter(function (i) { return Schema.normalizeHeader(i.sheet) === Schema.normalizeHeader(Schema.SHEET_SETTINGS); });
      var pick = onSettings.length ? onSettings : images;
      var tasks = [];
      if (pick[0]) tasks.push(normalizeImage(XlsxImages.toDataUrl(pick[0]), pick[0].mime === 'image/jpeg').then(function (img) { state.excelImages.logo = img; }, function () {}));
      if (pick[1]) tasks.push(normalizeImage(XlsxImages.toDataUrl(pick[1]), pick[1].mime === 'image/jpeg').then(function (img) { state.excelImages.stamp = img; }, function () {}));
      return Promise.all(tasks).then(function () { renderImages(); invalidate(); });
    }).catch(function () { /* csv / xls / no pictures */ });
    goStep(2);
  }

  function reparse() {
    state.parsed = Schema.parseWorkbook(XLSX, state.wb, { overrides: state.profile.columnOverrides });
    mergeCompany();
    loadLog();
    invalidate();
    $('fileName').textContent = state.fileName + (state.parsed.sheetName ? ' · sheet "' + state.parsed.sheetName + '"' : '');
    renderReview();
    populatePreviewSelect();
  }

  // Company settings: Excel values win, blanks are filled from the saved profile.
  function currentDefault(key) {
    var f = Schema.SETTINGS_FIELDS.filter(function (x) { return x.key === key; })[0];
    return f && f.default != null ? f.default : '';
  }

  // Message texts saved from older versions are replaced by the current short default.
  function upgradeMessageText(obj) {
    if (obj && obj.messageText && Schema.OLD_MESSAGE_DEFAULTS.indexOf(String(obj.messageText).replace(/\r\n/g, '\n')) >= 0) {
      obj.messageText = currentDefault('messageText');
    }
    return obj;
  }

  function mergeCompany() {
    var excel = upgradeMessageText(state.parsed.settings);
    var saved = upgradeMessageText(state.profile.company || {});
    var out = {};
    state.companyFromExcel = {};
    Schema.SETTINGS_FIELDS.forEach(function (f) {
      if (f.type === 'image') return;
      var v = excel[f.key];
      var isDefaultOnly = f.default != null && v === f.default && !state.parsed.settingsFound;
      var fromExcel = v !== '' && v != null && !isDefaultOnly;
      if (fromExcel) {
        out[f.key] = v;
        state.companyFromExcel[f.key] = true;
      } else if (saved[f.key] != null && saved[f.key] !== '') {
        out[f.key] = saved[f.key];
      } else {
        out[f.key] = f.default != null ? f.default : '';
      }
    });
    state.company = out;
    state.profile.company = Object.assign({}, saved, out);
    saveProfile();
  }

  function companyMissing() {
    return Schema.SETTINGS_FIELDS.filter(function (f) { return f.required && !String(state.company[f.key] || '').trim(); });
  }

  function effectiveImage(kind) {
    return state.excelImages[kind] || state.profile[kind] || null;
  }

  function setCompanyValue(k, v) {
    state.company[k] = v;
    state.profile.company[k] = v;
    saveProfile();
    invalidate();
  }

  // Anything that changes the rendered PDFs clears the per-employee cache.
  function invalidate() {
    state.cache = {};
    schedulePreview();
  }

  // ---------------------------------------------------------------------------
  // Step 2: review
  // ---------------------------------------------------------------------------
  function renderReview() {
    var p = state.parsed;
    var st = p.stats;
    var errs = p.issues.filter(function (i) { return i.level === 'error'; });
    var warns = p.issues.filter(function (i) { return i.level === 'warning'; });
    var infos = p.issues.filter(function (i) { return i.level === 'info'; });
    var missing = companyMissing();

    var banner = '';
    var kind = 'ok';
    if (errs.length) {
      kind = 'error';
      banner = '<strong>The file needs attention before payslips can be made.</strong> ' + escapeHtml(errs[0].message) + (errs.length > 1 ? ' (+' + (errs.length - 1) + ' more under Messages)' : '');
    } else if (!st.total) {
      kind = 'error';
      banner = '<strong>No employees were found.</strong> Check the Columns section below.';
    } else if (missing.length) {
      kind = 'warn';
      banner = '<strong>' + plural(st.ok, 'payslip') + ' ready.</strong> Please fill in the company details below (' + missing.map(function (f) { return f.label; }).join(', ') + ').';
    } else if (st.withErrors) {
      kind = 'warn';
      banner = '<strong>' + plural(st.ok, 'payslip') + ' ready.</strong> ' + plural(st.withErrors, 'row has', 'rows have') + ' errors and will be skipped unless fixed in Excel. Tick "Only rows with problems" in the Employees list to see them.';
    } else if (st.withWarnings || warns.length) {
      banner = '<strong>All ' + plural(st.ok, 'payslip') + ' ready.</strong> A few warnings, nothing blocking.';
    } else {
      banner = '<strong>All ' + plural(st.ok, 'payslip') + ' ready.</strong>';
    }
    setAlert($('statusBanner'), banner, kind);

    var contactsCls = st.withNoContact ? 'warn' : 'ok';
    var tiles = [
      { label: 'Employees', value: st.total, cls: '' },
      { label: 'Ready', value: st.ok, cls: st.ok ? 'ok' : '' },
      { label: 'Rows with errors', value: st.withErrors, cls: st.withErrors ? 'err' : '' },
      channelEnabled('whatsapp')
        ? { label: 'Contacts', value: '<span title="with mobile">' + st.withMobile + ' 📱</span> · <span title="with email">' + st.withEmail + ' ✉</span>', cls: contactsCls, sub: st.withNoContact ? st.withNoContact + ' with neither' : 'all reachable' }
        : { label: 'Total net (' + escapeHtml(state.company.currency || 'SAR') + ')', value: Schema.formatMoney(st.totalNet), money: st.totalNet, cls: '', sub: st.withEmail + ' of ' + st.total + ' have an email' }
    ];
    $('summaryGrid').innerHTML = tiles.map(function (t) {
      var num = typeof t.value === 'number' ? ' data-count="' + t.value + '"' : (t.money != null ? ' data-money="' + t.money + '"' : '');
      return '<div class="tile spot ' + t.cls + '"><div class="tile-label">' + t.label + '</div><div class="tile-value"' + num + '>' + (typeof t.value === 'number' ? '0' : t.value) + '</div>' + (t.sub ? '<div class="muted small">' + t.sub + '</div>' : '') + '</div>';
    }).join('');
    Array.prototype.forEach.call($('summaryGrid').querySelectorAll('[data-count]'), function (el) {
      Art.countUp(el, +el.getAttribute('data-count'), { format: function (v) { return String(Math.round(v)); } });
    });
    Array.prototype.forEach.call($('summaryGrid').querySelectorAll('[data-money]'), function (el) {
      Art.countUp(el, +el.getAttribute('data-money'), { duration: 1100, format: function (v) { return Schema.formatMoney(v); } });
    });

    renderCompanyForm();
    renderColumns();

    var all = errs.concat(warns, infos);
    $('messagesMeta').innerHTML = all.length
      ? (errs.length ? '<span class="badge err">' + plural(errs.length, 'error') + '</span>' : '') +
        (warns.length ? '<span class="badge warn">' + plural(warns.length, 'warning') + '</span>' : '') +
        (infos.length ? '<span class="badge neutral">' + plural(infos.length, 'note') + '</span>' : '')
      : '<span class="badge ok">none</span>';
    $('fileIssues').innerHTML = all.length
      ? '<ul class="msg-list">' + all.map(function (i) {
          return '<li class="issue-' + (i.level === 'error' ? 'err' : i.level === 'warning' ? 'warn' : 'info') + '">' + escapeHtml(i.message) + '</li>';
        }).join('') + '</ul>'
      : '<p class="muted small">No messages.</p>';
    $('secMessages').open = errs.length > 0;

    renderTable();
    refreshExportState();
  }

  function companyMetaHtml() {
    var missing = companyMissing();
    return missing.length
      ? '<span class="badge err">' + plural(missing.length, 'required field') + ' missing</span>'
      : '<span class="badge ok">complete</span> <span>' + escapeHtml(state.company.companyName || '') + ' · ' + escapeHtml(state.company.payPeriod || '') + '</span>';
  }

  function renderCompanyForm() {
    var missing = companyMissing();
    var fromExcelCount = Object.keys(state.companyFromExcel).length;
    $('companyMeta').innerHTML = companyMetaHtml();
    $('companySource').textContent = state.parsed.settingsFound
      ? 'Filled from the Settings sheet of the file' + (fromExcelCount < 3 ? ' (partly)' : '') + '. You can correct anything here; edits are kept for next time.'
      : 'This file has no Settings sheet. Enter the details here once; they are saved on this computer and in your settings file.';
    $('secCompany').open = missing.length > 0;

    var form = $('companyForm');
    // Delivery-related settings live on the Send step; keep this form about the company.
    var hidden = { messageSubject: 1, messageText: 1, senderEmail: 1 };
    form.innerHTML = Schema.SETTINGS_FIELDS.filter(function (f) { return f.type !== 'image' && !hidden[f.key]; }).map(function (f) {
      var v = state.company[f.key] || '';
      var isMissing = f.required && !String(v).trim();
      return '<label class="' + (f.key === 'footerNote' || f.key === 'address' ? 'wide' : '') + '"><span>' +
        escapeHtml(f.label) + (f.required ? ' <span class="req">*</span>' : '') +
        (state.companyFromExcel[f.key] ? ' <span class="badge neutral" title="Value from the Excel file">Excel</span>' : '') +
        '</span><input type="text" data-key="' + f.key + '" value="' + escapeHtml(v) + '" placeholder="' + escapeHtml(f.example ? 'e.g. ' + f.example : '') + '" class="' + (isMissing ? 'missing' : '') + '"></label>';
    }).join('');
    Array.prototype.forEach.call(form.querySelectorAll('input'), function (inp) {
      inp.addEventListener('input', function () {
        var k = inp.getAttribute('data-key');
        var def = Schema.SETTINGS_FIELDS.filter(function (f) { return f.key === k; })[0];
        setCompanyValue(k, k === 'currency' ? inp.value.toUpperCase() : inp.value);
        inp.classList.toggle('missing', !!(def && def.required) && !inp.value.trim());
        $('companyMeta').innerHTML = companyMetaHtml();
        refreshExportState();
      });
    });
  }

  function renderColumns() {
    var p = state.parsed;
    var cols = p.columns || [];
    var ignored = cols.filter(function (c) { return c.kind === 'ignored'; });
    var guessed = cols.filter(function (c) { return c.source === 'guess'; });
    var missingReq = [];
    var seenKeys = {};
    cols.forEach(function (c) { if (c.kind === 'column') seenKeys[c.def.key] = true; });
    if (!seenKeys.name) missingReq.push('Employee Name');
    var hasEarning = cols.some(function (c) { return (c.kind === 'column' && c.def.group === 'earning') || c.kind === 'earning'; });
    if (!hasEarning) missingReq.push('Basic Salary (or any earning)');
    var optionalMissing = [];
    if (!seenKeys.employeeId) optionalMissing.push('Employee ID');
    if (!seenKeys.basicSalary && hasEarning) optionalMissing.push('Basic Salary');
    if (!seenKeys.mobile && channelEnabled('whatsapp')) optionalMissing.push('Mobile (needed for WhatsApp)');
    if (!seenKeys.email) optionalMissing.push('Email (needed for email delivery)');
    if (!seenKeys.dob) optionalMissing.push('Date of Birth (used for the PDF password)');

    var meta = '';
    if (missingReq.length) meta += '<span class="badge err">' + plural(missingReq.length, 'required column') + ' missing</span>';
    if (ignored.length) meta += '<span class="badge warn">' + plural(ignored.length, 'column') + ' ignored</span>';
    if (guessed.length) meta += '<span class="badge info">' + plural(guessed.length, 'guess', 'guesses') + '</span>';
    if (!meta) meta = '<span class="badge ok">all ' + cols.length + ' recognised</span>';
    $('columnsMeta').innerHTML = meta;
    $('secColumns').open = missingReq.length > 0 || ignored.length > 0 || guessed.length > 0;

    var mm = '';
    if (missingReq.length) mm += '<div class="alert alert-error">Missing: <strong>' + missingReq.map(escapeHtml).join('</strong>, <strong>') + '</strong>. Pick the matching column in the "Understood as" list below, or add the column in Excel.</div>';
    if (optionalMissing.length) mm += '<div class="alert alert-info">Not found (optional): ' + optionalMissing.map(escapeHtml).join(', ') + '.' + (optionalMissing[0] === 'Employee ID' ? ' IDs will be generated from the row order.' : '') + '</div>';
    $('missingColumns').innerHTML = mm;

    var choices = Schema.mappingChoices();
    var groups = [];
    choices.forEach(function (c) {
      var g = c.group || '';
      var last = groups[groups.length - 1];
      if (!last || last.name !== g) groups.push({ name: g, items: [c] });
      else last.items.push(c);
    });

    $('columnsTable').querySelector('tbody').innerHTML = cols.map(function (c) {
      var current = Schema.targetToChoice(c);
      var optionsHtml = groups.map(function (g) {
        var inner = g.items.map(function (it) {
          return '<option value="' + it.value + '"' + (it.value === current ? ' selected' : '') + '>' + escapeHtml(it.label) + '</option>';
        }).join('');
        return g.name ? '<optgroup label="' + escapeHtml(g.name) + '">' + inner + '</optgroup>' : inner;
      }).join('');
      var how = c.source === 'manual' ? '<span class="badge info">your choice</span>'
        : c.source === 'exact' ? '<span class="badge ok">matched</span>'
        : c.source === 'guess' ? '<span class="badge warn">guessed</span>'
        : c.source === 'duplicate' ? '<span class="badge warn">duplicate</span>'
        : '<span class="badge neutral">not recognised</span>';
      return '<tr class="' + (c.kind === 'ignored' ? 'row-warn' : '') + '">' +
        '<td><strong>' + escapeHtml(c.header) + '</strong></td>' +
        '<td><select data-norm="' + escapeHtml(c.norm) + '" data-header="' + escapeHtml(c.header) + '">' + optionsHtml + '</select></td>' +
        '<td class="samples">' + escapeHtml(c.samples.join(' · ')) + '</td>' +
        '<td>' + how + '</td></tr>';
    }).join('');

    Array.prototype.forEach.call($('columnsTable').querySelectorAll('select'), function (sel) {
      sel.addEventListener('change', function () {
        state.profile.columnOverrides[sel.getAttribute('data-norm')] = Schema.choiceToTarget(sel.value, sel.getAttribute('data-header'));
        saveProfile();
        var wasOpen = $('secColumns').open;
        reparse();
        $('secColumns').open = wasOpen;
      });
    });
  }

  function renderTable() {
    var p = state.parsed;
    if (!p) return;
    var q = $('searchBox').value.trim().toLowerCase();
    var onlyProblems = $('onlyProblems').checked;
    var rows = p.slips.filter(function (s) {
      if (onlyProblems && !s.issues.length) return false;
      if (!q) return true;
      return (s.employeeId + ' ' + s.name).toLowerCase().indexOf(q) >= 0;
    });
    var LIMIT = 500;
    $('employeeTable').querySelector('tbody').innerHTML = rows.slice(0, LIMIT).map(function (s) {
      var cls = s.hasErrors ? 'row-err' : s.issues.length ? 'row-warn' : '';
      var badge = s.hasErrors ? '<span class="badge err">Error</span>' : s.issues.length ? '<span class="badge warn">Check</span>' : '<span class="badge ok">OK</span>';
      var checks = s.issues.map(function (i) {
        return '<div class="issue-' + (i.level === 'error' ? 'err' : 'warn') + '">' + escapeHtml(i.message) + '</div>';
      }).join('');
      return '<tr class="' + cls + '">' +
        '<td>' + s.rowNumber + '</td><td>' + badge + '</td>' +
        '<td>' + escapeHtml(s.employeeId) + (s.generatedId ? ' <span class="badge neutral" title="Generated">auto</span>' : '') + '</td>' +
        '<td>' + escapeHtml(s.name) + '</td>' +
        '<td class="num">' + Schema.formatMoney(s.gross) + '</td>' +
        '<td class="num">' + Schema.formatMoney(s.totalDeductions) + '</td>' +
        '<td class="num"><strong>' + Schema.formatMoney(s.net) + '</strong></td>' +
        '<td class="checks">' + (checks || '<span class="muted">—</span>') + '</td></tr>';
    }).join('');
    $('rowCount').innerHTML = '<span class="badge neutral">' + (rows.length === p.slips.length ? p.slips.length : rows.length + ' of ' + p.slips.length) + '</span>';
    $('tableNote').textContent = rows.length > LIMIT ? 'Showing the first ' + LIMIT + ' of ' + rows.length + ' rows. Use the search box to find an employee; all rows are still exported.' : '';
  }

  $('searchBox').addEventListener('input', renderTable);
  $('onlyProblems').addEventListener('change', renderTable);

  // ---------------------------------------------------------------------------
  // Step 3: design
  // ---------------------------------------------------------------------------
  function renderTemplateCards() {
    $('templateCards').innerHTML = PDF.TEMPLATES.map(function (t) {
      return '<button type="button" class="template-card' + (t.id === state.profile.templateId ? ' active' : '') + '" data-id="' + t.id + '">' +
        '<div class="thumb">' + PDF.thumbSvg(t.id, state.profile.accent) + '</div>' +
        '<div class="name">' + escapeHtml(t.name) + '</div><div class="desc">' + escapeHtml(t.description) + '</div></button>';
    }).join('');
    Array.prototype.forEach.call($('templateCards').querySelectorAll('.template-card'), function (btn) {
      btn.addEventListener('click', function () {
        state.profile.templateId = btn.getAttribute('data-id');
        renderTemplateCards();
        saveProfile();
        invalidate();
      });
    });
    document.documentElement.style.setProperty('--accent', state.profile.accent);
  }

  function renderDesignControls() {
    $('accent').value = state.profile.accent;
    $('paper').value = state.profile.paper;
    $('docTitle').value = state.profile.display.title || 'PAYSLIP';
    renderTemplateCards();
    renderDisplayOptions();
    renderImages();
  }

  var thumbTimer = null;
  $('accent').addEventListener('input', function () {
    state.profile.accent = $('accent').value;
    document.documentElement.style.setProperty('--accent', state.profile.accent);
    saveProfile();
    invalidate();
    clearTimeout(thumbTimer);
    thumbTimer = setTimeout(renderTemplateCards, 120);
  });
  $('paper').addEventListener('change', function () {
    state.profile.paper = $('paper').value;
    saveProfile();
    invalidate();
  });
  $('docTitle').addEventListener('input', function () {
    state.profile.display.title = $('docTitle').value;
    saveProfile();
    invalidate();
  });

  function renderDisplayOptions() {
    var groups = [];
    PDF.DISPLAY_OPTIONS.forEach(function (o) {
      var last = groups[groups.length - 1];
      if (!last || last.name !== o.group) groups.push({ name: o.group, items: [o] });
      else last.items.push(o);
    });
    $('displayOptions').innerHTML = groups.map(function (g) {
      return '<div class="display-group"><h4>' + escapeHtml(g.name) + '</h4>' + g.items.map(function (o) {
        return '<label><input type="checkbox" data-opt="' + o.key + '"' + (state.profile.display[o.key] ? ' checked' : '') + '> ' + escapeHtml(o.label) + '</label>';
      }).join('') + '</div>';
    }).join('');
    Array.prototype.forEach.call($('displayOptions').querySelectorAll('input'), function (cb) {
      cb.addEventListener('change', function () {
        state.profile.display[cb.getAttribute('data-opt')] = cb.checked;
        saveProfile();
        updateDisplayMeta();
        invalidate();
      });
    });
    updateDisplayMeta();
  }

  function updateDisplayMeta() {
    var off = PDF.DISPLAY_OPTIONS.filter(function (o) { return !state.profile.display[o.key]; }).length;
    $('displayMeta').innerHTML = off ? '<span class="badge neutral">' + plural(off, 'item') + ' hidden</span>' : '<span class="badge neutral">everything shown</span>';
  }

  function renderImages() {
    ['logo', 'stamp'].forEach(function (kind) {
      var img = effectiveImage(kind);
      $(kind + 'Preview').innerHTML = img ? '<img src="' + img.dataUrl + '" alt="' + kind + '">' : (kind === 'logo' ? 'No logo' : 'No stamp / signature');
      var src = $(kind + 'Source');
      if (state.excelImages[kind]) src.textContent = 'Taken from the picture in the Excel file.';
      else if (state.profile[kind]) src.textContent = 'Saved on this computer. Tip: insert it on the Settings sheet of the Excel file to keep it with the data.';
      else src.textContent = kind === 'logo' ? 'PNG or JPG. Printed top-left of every payslip.' : 'Optional company stamp or signature, printed by the signature line.';
    });
  }

  function bindImage(kind) {
    $(kind + 'Input').addEventListener('change', function (e) {
      var f = e.target.files && e.target.files[0];
      e.target.value = '';
      if (!f) return;
      fileToImage(f).then(function (img) {
        state.profile[kind] = img;
        state.excelImages[kind] = null;
        renderImages();
        saveProfile();
        invalidate();
      }).catch(function (err) { alert(err.message); });
    });
    $(kind + 'Clear').addEventListener('click', function () {
      state.profile[kind] = null;
      state.excelImages[kind] = null;
      renderImages();
      saveProfile();
      invalidate();
    });
  }

  // ---- settings file ---------------------------------------------------------
  function settingsFileObject() {
    var p = state.profile;
    return {
      app: 'payslip-generator',
      version: SETTINGS_FILE_VERSION,
      savedAt: new Date().toISOString(),
      company: state.company && Object.keys(state.company).length ? state.company : p.company,
      templateId: p.templateId,
      accent: p.accent,
      paper: p.paper,
      display: p.display,
      channel: p.channel,
      logo: effectiveImage('logo'),
      stamp: effectiveImage('stamp'),
      columnOverrides: p.columnOverrides
    };
  }

  $('btnSaveSettings').addEventListener('click', function () {
    var obj = settingsFileObject();
    downloadBlob(new Blob([JSON.stringify(obj, null, 2)], { type: 'application/json' }), PDF.safeFileName('payslip-settings - ' + (obj.company.companyName || 'my company')) + '.json');
  });

  function loadSettingsFile(file) {
    return readFile(file, 'text').then(function (text) {
      var obj = JSON.parse(text);
      if (!obj || obj.app !== 'payslip-generator') throw new Error('This is not a Payslip Generator settings file.');
      applyProfile(obj);
      state.excelImages = { logo: null, stamp: null };
      if (state.parsed) reparse();
      renderDesignControls();
      updateSavedHint();
      setAlert($('loadError'), 'Settings loaded from ' + escapeHtml(file.name) + (obj.company && obj.company.companyName ? ' (' + escapeHtml(obj.company.companyName) + ')' : '') + '.' + (state.parsed ? '' : ' Now upload the Excel file.'), 'ok');
      invalidate();
    }).catch(function (e) {
      setAlert($('loadError'), 'Could not load settings: ' + escapeHtml(e.message), 'error');
      goStep(1);
    });
  }

  ['settingsInput', 'settingsInput2'].forEach(function (id) {
    $(id).addEventListener('change', function (e) {
      var f = e.target.files && e.target.files[0];
      e.target.value = '';
      if (f) loadSettingsFile(f);
    });
  });

  function updateSavedHint() {
    var c = state.profile.company || {};
    var hasSomething = c.companyName || state.profile.logo || Object.keys(state.profile.columnOverrides).length;
    $('savedHint').textContent = hasSomething
      ? 'This computer already remembers ' + (c.companyName ? '"' + c.companyName + '"' : 'your settings') + (state.profile.logo ? ' and the logo' : '') + '. They are applied automatically.'
      : '';
  }

  // ---------------------------------------------------------------------------
  // PDF generation (shared by preview, export and send)
  // ---------------------------------------------------------------------------
  function exportableSlips() {
    if (!state.parsed) return [];
    var skip = $('skipErrors').checked;
    return state.parsed.slips.filter(function (s) { return !skip || !s.hasErrors; });
  }

  function renderContext(slip) {
    var p = state.profile;
    return {
      slip: slip,
      settings: state.company,
      templateId: p.templateId,
      accent: p.accent,
      logo: effectiveImage('logo'),
      stamp: effectiveImage('stamp'),
      options: p.display
    };
  }

  function protectEnabled() {
    return $('protectPdfs').checked && String(state.company.passwordRule || 'id').toLowerCase() !== 'none';
  }

  function passwordFor(slip) {
    if (!protectEnabled()) return { password: '', hint: '', hintAr: '', error: null, off: true };
    return Schema.buildPassword(slip, state.company.passwordRule);
  }

  // One employee's PDF (encrypted when enabled), cached until something changes.
  function artifactFor(slip) {
    var key = slip.employeeId + '|' + slip.rowNumber + '|' + (protectEnabled() ? 1 : 0);
    if (state.cache[key]) return state.cache[key];
    var pw = passwordFor(slip);
    var doc = PDF.createDoc(jsPDF, state.profile.paper, { password: pw.password });
    PDF.renderPayslip(doc, renderContext(slip));
    PDF.setMeta(doc, state.company, 'Payslip ' + slip.employeeId);
    var art = {
      bytes: new Uint8Array(doc.output('arraybuffer')),
      fileName: Schema.slipFileName(slip, state.company),
      password: pw,
      hash: ''
    };
    state.cache[key] = art;
    return art;
  }

  function messageFor(slip) {
    var vars = Schema.messageVars(slip, state.company, passwordFor(slip));
    return {
      subject: Schema.renderTemplate(state.company.messageSubject || 'Payslip {period} - {company}', vars),
      text: Schema.renderTemplate(state.company.messageText || '', vars)
    };
  }

  function populatePreviewSelect() {
    var sel = $('previewSelect');
    var slips = state.parsed ? state.parsed.slips : [];
    sel.innerHTML = slips.slice(0, 2000).map(function (s, i) {
      return '<option value="' + i + '">' + escapeHtml(s.employeeId + ' — ' + s.name) + (s.hasErrors ? ' (has errors)' : '') + '</option>';
    }).join('');
    sel.disabled = !slips.length;
    $('btnOpenPreview').disabled = !slips.length;
  }

  var previewTimer = null;
  function schedulePreview() {
    clearTimeout(previewTimer);
    previewTimer = setTimeout(renderPreview, 150);
  }

  function renderPreview() {
    if (state.step !== 3) return;
    var slips = state.parsed ? state.parsed.slips : [];
    var frame = $('previewFrame');
    if (!slips.length) {
      $('previewEmpty').classList.remove('hidden');
      frame.removeAttribute('src');
      return;
    }
    var idx = parseInt($('previewSelect').value, 10) || 0;
    var slip = slips[idx] || slips[0];
    try {
      var doc = PDF.createDoc(jsPDF, state.profile.paper); // preview is never password protected
      PDF.renderPayslip(doc, renderContext(slip));
      PDF.setMeta(doc, state.company, 'Payslip preview');
      var blob = doc.output('blob');
      if (state.previewUrl) URL.revokeObjectURL(state.previewUrl);
      state.previewUrl = URL.createObjectURL(blob);
      frame.src = state.previewUrl + '#toolbar=0&navpanes=0&view=FitH';
      $('previewEmpty').classList.add('hidden');
    } catch (e) {
      $('previewEmpty').textContent = 'Preview failed: ' + e.message;
      $('previewEmpty').classList.remove('hidden');
    }
  }

  $('previewSelect').addEventListener('change', renderPreview);
  $('btnOpenPreview').addEventListener('click', function () {
    if (state.previewUrl) window.open(state.previewUrl, '_blank');
  });

  // ---------------------------------------------------------------------------
  // Step 4: export
  // ---------------------------------------------------------------------------
  var fsaSupported = typeof window.showDirectoryPicker === 'function';

  function renderExportPanel() {
    if (!state.parsed) return;
    var slips = exportableSlips();
    var t = PDF.TEMPLATES.filter(function (x) { return x.id === state.profile.templateId; })[0];
    var items = [
      ['Company', state.company.companyName || '—'],
      ['Pay period', state.company.payPeriod || '—'],
      ['Payslips to export', String(slips.length) + (state.parsed.stats.withErrors ? ' (' + state.parsed.stats.withErrors + ' with errors' + ($('skipErrors').checked ? ' skipped' : ' included') + ')' : '')],
      ['Design', (t ? t.name : '') + ' · ' + state.profile.paper.toUpperCase() + (effectiveImage('logo') ? ' · logo' : ' · no logo')]
    ];
    $('exportSummary').innerHTML = items.map(function (it) {
      return '<div><div class="k">' + it[0] + '</div><div class="v">' + escapeHtml(it[1]) + '</div></div>';
    }).join('');

    var rule = String(state.company.passwordRule || 'id');
    var cannot = slips.filter(function (s) { return Schema.buildPassword(s, rule).error; }).length;
    var ruleText = { 'id+dob': 'Iqama + birth day/month', id: 'Iqama / National ID number', dob: 'date of birth', custom: '"Payslip Password" column', none: 'off' }[rule.toLowerCase()] || rule;
    $('protectHint').textContent = rule.toLowerCase() === 'none'
      ? '(off in company settings)'
      : '· ' + ruleText + (cannot ? ' · ' + cannot + ' row' + (cannot === 1 ? '' : 's') + ' without data stay open' : '');
    $('protectPdfs').disabled = rule.toLowerCase() === 'none';
    if (!fsaSupported) {
      $('btnExportFolder').classList.add('hidden');
    }
    refreshExportState();
  }

  function refreshExportState() {
    var missing = state.parsed ? companyMissing().filter(function (f) { return f.key !== 'crNumber'; }) : [];
    var slips = exportableSlips();
    var can = !!(state.parsed && slips.length && !missing.length && !state.exporting);
    ['btnExportPdf', 'btnExportZip', 'btnExportFolder'].forEach(function (id) { $(id).disabled = !can; });
    var el = $('exportState');
    if (!state.parsed) { el.innerHTML = ''; return; }
    if (missing.length) {
      setAlert(el, 'Fill in ' + missing.map(function (f) { return '<strong>' + f.label + '</strong>'; }).join(' and ') + ' first (Review → Company details).' +
        '<div class="alert-actions"><button type="button" class="btn btn-small" data-go="2">Go to Review</button></div>', 'warn');
      el.querySelector('[data-go]').addEventListener('click', function () { goStep(2); $('secCompany').open = true; });
    } else if (!slips.length) {
      setAlert(el, 'No rows can be exported. Fix the errors in Excel or untick "Skip rows with errors".', 'warn');
    } else if (!String(state.company.crNumber || '').trim()) {
      setAlert(el, 'CR number is empty. Payslips can be exported, but a CR number is recommended.', 'info');
    } else {
      el.innerHTML = '';
    }
  }

  $('skipErrors').addEventListener('change', function () { renderExportPanel(); });
  $('protectPdfs').addEventListener('change', function () { state.cache = {}; renderExportPanel(); });

  function setProgress(done, total, label, which) {
    var suffix = which === 5 ? '5' : '';
    $('progress' + suffix).classList.remove('hidden');
    $('progressFill' + suffix).style.width = (total ? Math.round((done / total) * 100) : 0) + '%';
    $('progressText' + suffix).textContent = label || (done + ' / ' + total);
  }

  function beginExport(which) {
    Art.say('export', which === 5 ? 'Packing every draft with its locked PDF…' : 'Building and locking every PDF…');
    state.exporting = true;
    state.cancelExport = false;
    $('btnCancel').classList.toggle('hidden', which === 5);
    $(which === 5 ? 'sendMessage' : 'exportMessage').innerHTML = '';
    refreshExportState();
  }

  function endExport(which) {
    state.exporting = false;
    $('btnCancel').classList.add('hidden');
    $('progress').classList.add('hidden');
    $('progress5').classList.add('hidden');
    refreshExportState();
  }

  $('btnCancel').addEventListener('click', function () { state.cancelExport = true; });

  function baseName() {
    return PDF.safeFileName(['Payslips', state.company.companyName, state.company.payPeriod].filter(Boolean).join(' - '));
  }

  function reportError(el, e) {
    var cancelled = e && e.message === 'cancelled';
    if (e && (e.name === 'AbortError')) { el.innerHTML = ''; return; }
    setAlert(el, cancelled ? 'Export cancelled.' : 'Failed: ' + escapeHtml(e.message), cancelled ? 'warn' : 'error');
  }

  // Generates every exportable slip's artifact with progress; returns items [{slip, art}]
  async function buildAll(slips, which) {
    var items = [];
    for (var i = 0; i < slips.length; i++) {
      if (state.cancelExport) throw new Error('cancelled');
      var art = artifactFor(slips[i]);
      if (!art.hash) art.hash = await sha256Hex(art.bytes);
      items.push({ slip: slips[i], art: art, fileName: art.fileName, hash: art.hash, password: art.password });
      if (i % 5 === 0 || i === slips.length - 1) {
        setProgress(i + 1, slips.length, 'Rendering payslip ' + (i + 1) + ' of ' + slips.length + '…', which);
        await nextTick();
      }
    }
    return items;
  }

  async function exportCombinedPdf() {
    var slips = exportableSlips();
    if (!slips.length) return;
    beginExport(4);
    var started = Date.now();
    try {
      var doc = PDF.createDoc(jsPDF, state.profile.paper);
      for (var i = 0; i < slips.length; i++) {
        if (state.cancelExport) throw new Error('cancelled');
        if (i) doc.addPage();
        PDF.renderPayslip(doc, renderContext(slips[i]));
        if (i % 10 === 0 || i === slips.length - 1) {
          setProgress(i + 1, slips.length, 'Rendering payslip ' + (i + 1) + ' of ' + slips.length + '…');
          await nextTick();
        }
      }
      PDF.setMeta(doc, state.company, baseName());
      setProgress(slips.length, slips.length, 'Building PDF file…');
      await nextTick();
      var blob = doc.output('blob');
      downloadBlob(blob, baseName() + '.pdf');
      celebrateDone();
      setAlert($('exportMessage'), '<strong>Done.</strong> ' + plural(slips.length, 'payslip') + ' in one PDF (' + (blob.size / 1024 / 1024).toFixed(1) + ' MB, ' + ((Date.now() - started) / 1000).toFixed(1) + ' s). Check your Downloads folder.', 'ok');
    } catch (e) {
      reportError($('exportMessage'), e);
    } finally {
      endExport(4);
    }
  }

  async function exportZip() {
    var slips = exportableSlips();
    if (!slips.length) return;
    beginExport(4);
    var started = Date.now();
    try {
      var items = await buildAll(slips, 4);
      var zip = new JSZip();
      var folder = zip.folder(baseName());
      items.forEach(function (it) { folder.file(it.fileName, it.art.bytes); });
      folder.file('manifest.csv', Delivery.manifestCsv(items, state.company));
      setProgress(1, 1, 'Packing ZIP…');
      await nextTick();
      var blob = await zip.generateAsync({ type: 'blob', compression: 'STORE' });
      downloadBlob(blob, baseName() + '.zip');
      state.exportedTo = 'zip';
      celebrateDone();
      setAlert($('exportMessage'), '<strong>Done.</strong> ' + plural(items.length, 'PDF file') + ' + manifest.csv in a ZIP (' + (blob.size / 1024 / 1024).toFixed(1) + ' MB, ' + ((Date.now() - started) / 1000).toFixed(1) + ' s).' + protectNote(items), 'ok');
    } catch (e) {
      reportError($('exportMessage'), e);
    } finally {
      endExport(4);
    }
  }

  async function writeToDir(dir, name, data) {
    var fh = await dir.getFileHandle(name, { create: true });
    var w = await fh.createWritable();
    await w.write(data);
    await w.close();
  }

  async function exportFolder() {
    var slips = exportableSlips();
    if (!slips.length) return;
    var dir;
    try {
      dir = await window.showDirectoryPicker({ mode: 'readwrite', id: 'payslips' });
    } catch (e) {
      return; // user cancelled
    }
    beginExport(4);
    var started = Date.now();
    try {
      var sub = await dir.getDirectoryHandle(baseName(), { create: true });
      var items = [];
      for (var i = 0; i < slips.length; i++) {
        if (state.cancelExport) throw new Error('cancelled');
        var art = artifactFor(slips[i]);
        if (!art.hash) art.hash = await sha256Hex(art.bytes);
        await writeToDir(sub, art.fileName, art.bytes);
        items.push({ slip: slips[i], art: art, fileName: art.fileName, hash: art.hash, password: art.password });
        if (i % 5 === 0 || i === slips.length - 1) {
          setProgress(i + 1, slips.length, 'Writing ' + art.fileName + ' (' + (i + 1) + ' of ' + slips.length + ')…');
          await nextTick();
        }
      }
      await writeToDir(sub, 'manifest.csv', Delivery.manifestCsv(items, state.company));
      state.dirHandle = sub;
      state.exportedTo = 'folder';
      celebrateDone();
      setAlert($('exportMessage'), '<strong>Done.</strong> ' + plural(items.length, 'PDF file') + ' + manifest.csv written to the folder <code>' + escapeHtml(baseName()) + '</code> (' + ((Date.now() - started) / 1000).toFixed(1) + ' s).' + protectNote(items) + ' <button type="button" class="btn btn-small" data-go="5" style="margin-left:8px">Continue to send →</button>', 'ok');
      $('exportMessage').querySelector('[data-go]').addEventListener('click', function () { goStep(5); });
    } catch (e) {
      reportError($('exportMessage'), e);
    } finally {
      endExport(4);
    }
  }

  function celebrateDone() {
    Art.confetti();
    Art.say('celebrate', state.step === 5 ? 'Pack ready! Open <b>READ ME</b> inside the ZIP.' : 'All PDFs are ready. <b>Next: Send.</b>', { settle: state.step === 5 ? 'send' : 'idle', after: 3200 });
  }

  function protectNote(items) {
    if (!protectEnabled()) return ' Files are not password protected.';
    var without = items.filter(function (it) { return !it.art.password.password; }).length;
    return ' Each file is password protected' + (without ? ' except ' + without + ' without the needed data' : '') + '.';
  }

  $('btnExportPdf').addEventListener('click', exportCombinedPdf);
  $('btnExportZip').addEventListener('click', exportZip);
  $('btnExportFolder').addEventListener('click', exportFolder);

  // ---------------------------------------------------------------------------
  // Step 5: send
  // ---------------------------------------------------------------------------
  var CHANNELS = [
    { id: 'whatsapp', name: 'WhatsApp', sub: 'WhatsApp Desktop or the phone', contact: 'mobile' },
    { id: 'outlook', name: 'Outlook (Windows)', sub: 'Ready drafts for classic Outlook', contact: 'email' },
    { id: 'mailmerge', name: 'Gmail / any mailbox', sub: 'Mail-merge pack, zero clicks per employee', contact: 'email' },
    { id: 'print', name: 'Print & sign', sub: 'Paper hand-out with a signature register', contact: 'none' }
  ];

  var canShareFiles = (function () {
    try {
      return !!(navigator.canShare && navigator.canShare({ files: [new File([new Blob(['x'], { type: 'application/pdf' })], 'x.pdf', { type: 'application/pdf' })] }));
    } catch (e) { return false; }
  })();

  // Outlook is the only visible channel for now. The others stay in the code: open the page with
  // ?channels=all (or ?channels=outlook,whatsapp) to bring them back; the choice is remembered.
  var ENABLED_CHANNELS = (function () {
    var v = null;
    try {
      var m = location.search.match(/[?&]channels=([^&]+)/);
      if (m) { v = decodeURIComponent(m[1]); localStorage.setItem('payslip-channels', v); }
      else v = localStorage.getItem('payslip-channels');
    } catch (e) { /* ignore */ }
    if (v === 'all') return CHANNELS.map(function (c) { return c.id; });
    var list = (v || 'outlook').split(',').filter(function (id) { return CHANNELS.some(function (c) { return c.id === id; }); });
    return list.length ? list : ['outlook'];
  })();

  function channelEnabled(id) {
    return ENABLED_CHANNELS.indexOf(id) >= 0;
  }

  function visibleChannels() {
    return CHANNELS.filter(function (c) { return channelEnabled(c.id); });
  }

  function currentChannel() {
    var vis = visibleChannels();
    return vis.filter(function (c) { return c.id === state.profile.channel; })[0] || vis[0];
  }

  function logKey() {
    return LOG_KEY_PREFIX + Schema.asciiSlug(state.company.companyName || 'company').toLowerCase() + '-' + Schema.periodSlug(state.company.payPeriod, state.company.periodStart);
  }

  function loadLog() {
    try {
      state.log = JSON.parse(localStorage.getItem(logKey()) || '[]');
    } catch (e) { state.log = []; }
  }

  function saveLog() {
    try { localStorage.setItem(logKey(), JSON.stringify(state.log)); } catch (e) { /* ignore */ }
  }

  function isSent(slip) {
    return state.log.some(function (e) { return e.employeeId === slip.employeeId && !e.undone; });
  }

  function markSent(slip, channel, note) {
    var art = state.cache[Object.keys(state.cache).filter(function (k) { return k.indexOf(slip.employeeId + '|') === 0; })[0]];
    state.log.push({
      timestamp: new Date().toISOString(),
      employeeId: slip.employeeId,
      name: slip.name,
      channel: channel,
      contact: channel === 'whatsapp' ? slip.fields.mobile : slip.fields.email,
      file: art ? art.fileName : Schema.slipFileName(slip, state.company),
      hash: art ? art.hash : '',
      note: note || ''
    });
    saveLog();
    renderSendList();
    renderSummaryChips();
    updateNextMeter();
    var row = $('sendTable').querySelector('tr[data-emp="' + cssEscape(slip.employeeId) + '"]');
    if (row) { row.classList.remove('flash'); void row.offsetWidth; row.classList.add('flash'); }
    var targets = exportableSlips().filter(function (s) { return currentChannel().contact === 'email' ? s.fields.email : true; });
    if (targets.length && targets.every(isSent)) {
      Art.confetti();
      Art.say('celebrate', 'Done! <b>' + targets.length + ' payslips</b> delivered. Download the delivery log for your records.', { settle: 'send', after: 4000 });
    }
  }

  function cssEscape(v) {
    return String(v).replace(/["\\]/g, '\\$&');
  }

  // Downloads the next unsent Outlook draft (button or Enter key).
  function nextDraft() {
    var next = exportableSlips().filter(function (s) { return s.fields.email && !isSent(s); })[0];
    if (!next) {
      Art.say('celebrate', 'Nobody left. <b>Everyone has a draft.</b>');
      return;
    }
    downloadDraft(next);
  }

  function downloadDraft(slip) {
    var a1 = artifactFor(slip);
    var m1 = messageFor(slip);
    var eml = Delivery.buildEml({ from: state.company.senderEmail, to: slip.fields.email, subject: m1.subject, text: m1.text, fileName: a1.fileName, pdfBytes: a1.bytes });
    downloadBlob(new Blob([eml], { type: 'message/rfc822' }), a1.fileName.replace(/\.pdf$/, '.eml'));
    markSent(slip, 'outlook', 'draft downloaded');
  }

  function updateNextMeter() {
    if (!$('nextMeter')) return;
    var all = exportableSlips().filter(function (s) { return s.fields.email; });
    var done = all.filter(isSent).length;
    $('nextMeter').style.width = (all.length ? Math.round(done / all.length * 100) : 0) + '%';
    $('nextCount').textContent = all.length ? done + ' of ' + all.length + ' drafted' : 'No employee has an email address in the file';
    $('btnNextDraft').disabled = !all.length || done === all.length;
  }

  document.addEventListener('keydown', function (e) {
    if (state.step !== 5 || e.key !== 'Enter' || e.repeat || e.ctrlKey || e.metaKey || e.altKey) return;
    var t = e.target;
    if (t && (t.tagName === 'INPUT' || t.tagName === 'TEXTAREA' || t.tagName === 'SELECT' || t.tagName === 'BUTTON' || t.tagName === 'A' || t.isContentEditable)) return;
    if (currentChannel().id !== 'outlook' || !$('btnNextDraft') || $('btnNextDraft').disabled) return;
    e.preventDefault();
    nextDraft();
  });

  function unmarkSent(slip) {
    state.log = state.log.filter(function (e) { return e.employeeId !== slip.employeeId; });
    saveLog();
    renderSendList();
    renderSummaryChips();
    updateNextMeter();
  }

  function renderSummaryChips() {
    var slips = exportableSlips();
    var sent = slips.filter(isSent).length;
    $('contactSummary').innerHTML = [
      '<span class="chip"><strong>' + slips.length + '</strong> payslips</span>',
      channelEnabled('whatsapp') ? '<span class="chip"><strong>' + slips.filter(function (s) { return s.fields.mobile; }).length + '</strong> with mobile</span>' : '',
      '<span class="chip"><strong>' + slips.filter(function (s) { return s.fields.email; }).length + '</strong> with email</span>',
      '<span class="chip"><strong>' + sent + '</strong> marked sent</span>',
      state.exportedTo ? '<span class="chip ok">PDFs exported</span>' : '<span class="chip warn">PDFs not exported yet</span>'
    ].join('');
  }

  function renderSendPanel() {
    if (!state.parsed) return;
    var slips = exportableSlips();
    renderSummaryChips();
    $('channelCards').classList.toggle('hidden', visibleChannels().length < 2);
    $('channelCards').innerHTML = visibleChannels().map(function (c) {
      var count = c.contact === 'mobile' ? slips.filter(function (s) { return s.fields.mobile; }).length
        : c.contact === 'email' ? slips.filter(function (s) { return s.fields.email; }).length : slips.length;
      return '<button type="button" class="channel-card' + (c.id === state.profile.channel ? ' active' : '') + '" data-id="' + c.id + '">' +
        '<span class="ch-name">' + escapeHtml(c.name) + '</span><span class="ch-sub">' + escapeHtml(c.sub) + '</span>' +
        '<span class="ch-count">' + count + ' of ' + slips.length + ' reachable</span></button>';
    }).join('');
    Array.prototype.forEach.call($('channelCards').querySelectorAll('.channel-card'), function (btn) {
      btn.addEventListener('click', function () {
        state.profile.channel = btn.getAttribute('data-id');
        saveProfile();
        renderSendPanel();
      });
    });
    $('msgMeta').textContent = visibleChannels().length > 1 ? 'WhatsApp · email · mail merge' : 'subject · body · sender';
    $('msgSubject').value = state.company.messageSubject || '';
    $('msgFrom').value = state.company.senderEmail || '';
    $('msgText').value = state.company.messageText || '';
    renderChannelHelp();
    renderMessagePreview();
    renderSendList();
  }

  var ICONS = {
    chat: '<svg viewBox="0 0 24 24"><path d="M21 11.5a8.4 8.4 0 0 1-9 8.4 8.6 8.6 0 0 1-3.7-.8L3 21l1.9-5.1A8.4 8.4 0 0 1 3 11.5a8.4 8.4 0 0 1 9-8.4 8.4 8.4 0 0 1 9 8.4z"/></svg>',
    clip: '<svg viewBox="0 0 24 24"><path d="m21.4 11.1-9.2 9.2a6 6 0 0 1-8.5-8.5l9.2-9.2a4 4 0 0 1 5.7 5.7l-9.2 9.2a2 2 0 0 1-2.8-2.8l8.5-8.5"/></svg>',
    send: '<svg viewBox="0 0 24 24"><path d="M22 2 11 13"/><path d="M22 2 15 22l-4-9-9-4z"/></svg>',
    download: '<svg viewBox="0 0 24 24"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><path d="m7 10 5 5 5-5"/><path d="M12 15V3"/></svg>',
    click: '<svg viewBox="0 0 24 24"><path d="m9 9 5 12 1.8-5.2L21 14z"/><path d="M7.2 2.2 8 5.1"/><path d="m5.1 8-2.9-.8"/><path d="M14 4.1 12 6"/><path d="m6 12-1.9 2"/></svg>',
    mail: '<svg viewBox="0 0 24 24"><rect x="2" y="4" width="20" height="16" rx="2"/><path d="m22 7-10 7L2 7"/></svg>',
    printer: '<svg viewBox="0 0 24 24"><path d="M6 9V2h12v7"/><path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2"/><rect x="6" y="14" width="12" height="8"/></svg>',
    hand: '<svg viewBox="0 0 24 24"><path d="M18 11V6a2 2 0 0 0-4 0v5"/><path d="M14 10V4a2 2 0 0 0-4 0v2"/><path d="M10 10.5V6a2 2 0 0 0-4 0v8"/><path d="M18 8a2 2 0 1 1 4 0v6a8 8 0 0 1-8 8h-2c-2.8 0-4.5-.9-5.9-2.5L2.7 15.3a2 2 0 0 1 3-2.6L6 14"/></svg>',
    pen: '<svg viewBox="0 0 24 24"><path d="M12 20h9"/><path d="M16.5 3.5a2.1 2.1 0 0 1 3 3L7 19l-4 1 1-4z"/></svg>',
    share: '<svg viewBox="0 0 24 24"><circle cx="18" cy="5" r="3"/><circle cx="6" cy="12" r="3"/><circle cx="18" cy="19" r="3"/><path d="m8.6 13.5 6.8 4"/><path d="m15.4 6.5-6.8 4"/></svg>',
    folder: '<svg viewBox="0 0 24 24"><path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"/></svg>'
  };

  function flowHtml(steps, extra) {
    return '<div class="flow">' + steps.map(function (st, i) {
      return '<div class="flow-step"><span class="flow-icon">' + ICONS[st.icon] + '<span class="flow-num">' + (i + 1) + '</span></span><span class="flow-title">' + escapeHtml(st.title) + '</span><span class="flow-sub">' + escapeHtml(st.sub || '') + '</span></div>';
    }).join('<span class="flow-arrow" aria-hidden="true">›</span>') + '</div>' + (extra || '');
  }

  function renderChannelHelp() {
    var ch = currentChannel();
    var html = '';
    var where = state.exportedTo ? 'from the export folder' : 'export the PDFs first';
    if (ch.id === 'whatsapp') {
      html = flowHtml([
        canShareFiles ? { icon: 'share', title: 'Share PDF', sub: 'pick WhatsApp + contact' } : { icon: 'chat', title: 'Open chat', sub: 'message pre-typed' },
        canShareFiles ? { icon: 'chat', title: 'Or Open chat', sub: 'then drag the PDF' } : { icon: 'clip', title: 'Attach the PDF', sub: where },
        { icon: 'send', title: 'Send', sub: 'row ticks itself' }
      ], '<p class="flow-note">Save contacts first · pause every 50 · stop if WhatsApp says "restricted"' + (/Mac/.test(navigator.platform) ? ' · on Mac use the + button, not drag-and-drop' : '') + '</p>');
    } else if (ch.id === 'outlook') {
      html = '<div class="ch-hero"><span class="ch-logo">' + ICONS.mail + '</span><div><h3>Send with Outlook</h3><p>Every draft opens ready: address, subject, message and the locked PDF.</p></div></div>' +
        flowHtml([
          { icon: 'download', title: 'Next draft', sub: 'one click or Enter' },
          { icon: 'click', title: 'Open it', sub: 'classic Outlook, Windows' },
          { icon: 'send', title: 'Send', sub: 'row ticks itself' }
        ]) +
        '<div class="next-wrap"><button type="button" class="btn btn-primary btn-xl" id="btnNextDraft">Next draft <kbd>Enter</kbd></button>' +
        '<div class="next-meter"><i id="nextMeter"></i></div><div class="next-count" id="nextCount"></div></div>' +
        '<div class="flow-actions"><button type="button" class="btn btn-small" id="btnEmlPack">All drafts at once (ZIP)</button><span class="muted small">batches of ' + EML_BATCH + ' + macro for Drafts · Microsoft 365 sends ≈30 per minute</span></div>';
    } else if (ch.id === 'mailmerge') {
      html = flowHtml([
        { icon: 'download', title: 'Pack', sub: 'PDFs + CSV + guide' },
        { icon: 'mail', title: 'Thunderbird', sub: 'free, any mailbox' },
        { icon: 'send', title: 'Send all', sub: 'zero clicks per employee' }
      ], '<div class="flow-actions"><button type="button" class="btn btn-primary" id="btnMergePack">Mail Merge pack</button><span class="muted small">Gmail 500/day · Workspace 2,000 · M365 ≈30/min · no install? use Gmail compose per row</span></div>');
    } else {
      html = flowHtml([
        { icon: 'printer', title: 'Print', sub: 'combined PDF (Export)' },
        { icon: 'hand', title: 'Hand out', sub: 'folded, by supervisor' },
        { icon: 'pen', title: 'Sign', sub: 'register = proof' }
      ], '<div class="flow-actions"><button type="button" class="btn btn-primary" id="btnRegister">Acknowledgement register</button></div>');
    }
    $('channelHelp').innerHTML = html;
    if ($('btnEmlPack')) $('btnEmlPack').addEventListener('click', downloadEmlPack);
    if ($('btnNextDraft')) $('btnNextDraft').addEventListener('click', nextDraft);
    updateNextMeter();
    if ($('btnMergePack')) $('btnMergePack').addEventListener('click', downloadMergePack);
    if ($('btnRegister')) $('btnRegister').addEventListener('click', downloadRegister);
    $('sendContactHead').textContent = ch.contact === 'mobile' ? 'Mobile' : ch.contact === 'email' ? 'Email' : 'Department';
  }

  function renderMessagePreview() {
    var slips = exportableSlips();
    if (!slips.length) { $('msgPreview').textContent = ''; return; }
    var m = messageFor(slips[0]);
    $('msgPreview').innerHTML = '<span class="subj">' + escapeHtml(m.subject) + '</span>' + escapeHtml(m.text);
  }

  ['msgSubject', 'msgFrom', 'msgText'].forEach(function (id) {
    $(id).addEventListener('input', function () {
      var key = { msgSubject: 'messageSubject', msgFrom: 'senderEmail', msgText: 'messageText' }[id];
      setCompanyValue(key, $(id).value);
      renderMessagePreview();
      renderSendList();
    });
  });

  function renderSendList() {
    var ch = currentChannel();
    var slips = exportableSlips();
    var q = $('sendSearch').value.trim().toLowerCase();
    var hideSent = $('hideSent').checked;
    var rows = slips.filter(function (s) {
      if (hideSent && isSent(s)) return false;
      if (!q) return true;
      return (s.employeeId + ' ' + s.name).toLowerCase().indexOf(q) >= 0;
    });
    var LIMIT = 300;
    var tbody = $('sendTable').querySelector('tbody');
    tbody.innerHTML = rows.slice(0, LIMIT).map(function (s, i) {
      var sent = isSent(s);
      var contact = ch.contact === 'mobile' ? s.fields.mobile : ch.contact === 'email' ? s.fields.email : (s.fields.department || '');
      var missing = ch.contact !== 'none' && !contact;
      var fileName = Schema.slipFileName(s, state.company);
      var acts = '';
      if (ch.id === 'whatsapp') {
        if (!missing) {
          var m = messageFor(s);
          if (canShareFiles) acts += '<button type="button" class="btn btn-small btn-primary" data-act="share" data-i="' + i + '">Share PDF</button>';
          acts += '<a class="btn btn-small" href="' + escapeHtml(Delivery.waLink(contact, m.subject + '\n' + m.text)) + '" target="_blank" rel="noopener" data-act="wa" data-i="' + i + '">Open chat</a>';
        }
        acts += '<button type="button" class="btn btn-small btn-ghost" data-act="pdf" data-i="' + i + '">PDF</button>';
      } else if (ch.id === 'outlook') {
        if (!missing) acts += '<button type="button" class="btn btn-small btn-primary" data-act="eml" data-i="' + i + '">Outlook draft</button>';
        acts += '<button type="button" class="btn btn-small btn-ghost" data-act="pdf" data-i="' + i + '">PDF</button>';
      } else if (ch.id === 'mailmerge') {
        if (!missing) {
          var mm = messageFor(s);
          acts += '<a class="btn btn-small" href="' + escapeHtml(Delivery.gmailComposeLink(contact, mm.subject, mm.text)) + '" target="_blank" rel="noopener" data-act="gmail" data-i="' + i + '">Gmail compose</a>';
          acts += '<a class="btn btn-small btn-ghost" href="' + escapeHtml(Delivery.mailtoLink(contact, mm.subject, mm.text)) + '" data-act="mailto" data-i="' + i + '">Other mail app</a>';
        }
        acts += '<button type="button" class="btn btn-small btn-ghost" data-act="pdf" data-i="' + i + '">PDF</button>';
      } else {
        acts += '<button type="button" class="btn btn-small btn-ghost" data-act="pdf" data-i="' + i + '">PDF</button>';
      }
      return '<tr class="' + (sent ? 'sent' : '') + '" data-i="' + i + '" data-emp="' + escapeHtml(s.employeeId) + '">' +
        '<td><input type="checkbox" data-act="tick" data-i="' + i + '"' + (sent ? ' checked' : '') + ' title="Mark as sent"></td>' +
        '<td>' + escapeHtml(s.employeeId) + '</td>' +
        '<td class="name">' + escapeHtml(s.name) + '</td>' +
        '<td>' + (missing ? '<span class="nocontact">no ' + (ch.contact === 'mobile' ? 'mobile' : 'email') + ' in file</span>' : escapeHtml(ch.contact === 'mobile' ? '+' + contact : contact)) + '</td>' +
        '<td class="muted small">' + escapeHtml(fileName) + '</td>' +
        '<td class="acts">' + acts + '</td></tr>';
    }).join('');
    tbody.__rows = rows;
    $('sendCount').textContent = rows.length === slips.length ? slips.length : rows.length + ' of ' + slips.length;
    $('sendNote').textContent = rows.length > LIMIT ? 'Showing the first ' + LIMIT + ' rows. Use the search box or tick "Hide sent" to see the rest.' : '';
  }

  $('sendSearch').addEventListener('input', renderSendList);
  $('hideSent').addEventListener('change', renderSendList);

  // One delegated handler for every row action. Links (wa.me / Gmail / mailto) are real anchors so
  // browsers do not treat them as popups; we only record the click.
  $('sendTable').addEventListener('click', function (e) {
    var el = e.target.closest('[data-act]');
    if (!el) return;
    var rows = $('sendTable').querySelector('tbody').__rows || [];
    var slip = rows[+el.getAttribute('data-i')];
    if (!slip) return;
    var act = el.getAttribute('data-act');
    if (act === 'tick') {
      if (el.checked) markSent(slip, currentChannel().id, 'manual');
      else unmarkSent(slip);
      return;
    }
    if (act === 'pdf') {
      var art = artifactFor(slip);
      downloadBlob(new Blob([art.bytes], { type: 'application/pdf' }), art.fileName);
      return;
    }
    if (act === 'eml') {
      downloadDraft(slip);
      return;
    }
    if (act === 'share') {
      e.preventDefault();
      var a2 = artifactFor(slip);
      var m2 = messageFor(slip);
      var file = new File([a2.bytes], a2.fileName, { type: 'application/pdf' });
      navigator.share({ files: [file], title: m2.subject, text: m2.subject + '\n' + m2.text }).then(function () {
        markSent(slip, 'whatsapp', 'shared');
      }).catch(function (err) {
        if (err && err.name !== 'AbortError') setAlert($('sendMessage'), 'Share failed: ' + escapeHtml(err.message) + '. Use "Open chat" and attach the file by hand.', 'warn');
      });
      return;
    }
    if (act === 'wa' || act === 'gmail' || act === 'mailto') {
      // let the anchor open, then record it
      setTimeout(function () { markSent(slip, act === 'wa' ? 'whatsapp' : 'email', act === 'wa' ? 'chat opened' : 'compose opened'); }, 50);
    }
  });

  async function downloadEmlPack() {
    var slips = exportableSlips().filter(function (s) { return s.fields.email; });
    if (!slips.length) { setAlert($('sendMessage'), 'No employee has an email address in the file.', 'warn'); return; }
    beginExport(5);
    try {
      var items = await buildAll(slips, 5);
      var zip = new JSZip();
      var root = zip.folder(baseName() + ' - Outlook drafts');
      items.forEach(function (it, i) {
        var m = messageFor(it.slip);
        var eml = Delivery.buildEml({ from: state.company.senderEmail, to: it.slip.fields.email, subject: m.subject, text: m.text, fileName: it.fileName, pdfBytes: it.art.bytes });
        root.folder('batch-' + ('0' + (Math.floor(i / EML_BATCH) + 1)).slice(-2)).file(it.fileName.replace(/\.pdf$/, '.eml'), eml);
      });
      var pdfs = root.folder('pdf');
      items.forEach(function (it) { pdfs.file(it.fileName, it.art.bytes); });
      pdfs.file('manifest.csv', Delivery.manifestCsv(items, state.company));
      pdfs.file('OutlookDrafts.bas', Delivery.outlookVbaModule(state.company));
      root.file('READ ME - how to send.html', Delivery.outlookGuideHtml(state.company, EML_BATCH));
      setProgress(1, 1, 'Packing ZIP…', 5);
      await nextTick();
      var blob = await zip.generateAsync({ type: 'blob', compression: 'STORE' });
      downloadBlob(blob, baseName() + ' - Outlook drafts.zip');
      celebrateDone();
      setAlert($('sendMessage'), '<strong>Done.</strong> ' + plural(items.length, 'draft') + ' in ' + Math.ceil(items.length / EML_BATCH) + ' batch folder(s), plus the PDFs, manifest and macro (' + (blob.size / 1024 / 1024).toFixed(1) + ' MB). Open "READ ME - how to send.html" first.', 'ok');
    } catch (e) {
      reportError($('sendMessage'), e);
    } finally {
      endExport(5);
    }
  }

  async function downloadMergePack() {
    var all = exportableSlips();
    var slips = all.filter(function (s) { return s.fields.email; });
    if (!slips.length) { setAlert($('sendMessage'), 'No employee has an email address in the file.', 'warn'); return; }
    beginExport(5);
    try {
      var items = await buildAll(slips, 5);
      items.forEach(function (it) {
        var m = messageFor(it.slip);
        it.subject = m.subject;
        it.body = m.text;
      });
      var zip = new JSZip();
      var root = zip.folder(baseName() + ' - Mail Merge');
      items.forEach(function (it) { root.file(it.fileName, it.art.bytes); });
      root.file('mailmerge.csv', Delivery.mailMergeCsv(items, state.company));
      root.file('manifest.csv', Delivery.manifestCsv(items, state.company));
      root.file('READ ME - how to send.html', Delivery.mailMergeGuideHtml(state.company, { withEmail: items.length, withoutEmail: all.length - items.length }));
      setProgress(1, 1, 'Packing ZIP…', 5);
      await nextTick();
      var blob = await zip.generateAsync({ type: 'blob', compression: 'STORE' });
      downloadBlob(blob, baseName() + ' - Mail Merge.zip');
      celebrateDone();
      setAlert($('sendMessage'), '<strong>Done.</strong> Mail Merge pack with ' + plural(items.length, 'PDF') + ', mailmerge.csv and the guide (' + (blob.size / 1024 / 1024).toFixed(1) + ' MB). Open "READ ME - how to send.html" first.' + (all.length - items.length ? ' ' + (all.length - items.length) + ' employee(s) without email are not included.' : ''), 'ok');
    } catch (e) {
      reportError($('sendMessage'), e);
    } finally {
      endExport(5);
    }
  }

  function downloadRegister() {
    var slips = exportableSlips();
    var doc = PDF.createDoc(jsPDF, state.profile.paper);
    PDF.renderRegister(doc, { settings: state.company, slips: slips });
    PDF.setMeta(doc, state.company, 'Acknowledgement register');
    downloadBlob(doc.output('blob'), baseName() + ' - Acknowledgement register.pdf');
  }

  $('btnLogCsv').addEventListener('click', function () {
    downloadBlob(new Blob([Delivery.deliveryLogCsv(state.log)], { type: 'text/csv' }), baseName() + ' - delivery log.csv');
  });
  $('btnClearSent').addEventListener('click', function () {
    if (!state.log.length || confirm('Clear all "sent" ticks for ' + (state.company.payPeriod || 'this period') + '? The downloaded delivery log keeps the history.')) {
      state.log = [];
      saveLog();
      renderSendPanel();
    }
  });

  // ---------------------------------------------------------------------------
  // Init
  // ---------------------------------------------------------------------------
  Art.init();
  loadProfile();
  renderDesignControls();
  bindImage('logo');
  bindImage('stamp');
  updateSavedHint();
  goStep(1);

  window.PayslipApp = {
    state: state,
    loadWorkbookData: loadWorkbookData,
    exportableSlips: exportableSlips,
    artifactFor: artifactFor,
    messageFor: messageFor,
    goStep: goStep,
    applyProfile: applyProfile
  };
})();
