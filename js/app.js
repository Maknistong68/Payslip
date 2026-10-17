/* Payslip Generator - UI. Everything runs in the browser; no network calls. */
(function () {
  'use strict';

  var $ = function (id) { return document.getElementById(id); };

  if (typeof XLSX === 'undefined' || typeof jspdf === 'undefined' || typeof JSZip === 'undefined' ||
      typeof PayslipSchema === 'undefined' || typeof PayslipPDF === 'undefined' || typeof XlsxImages === 'undefined' ||
      typeof PayslipDelivery === 'undefined' || typeof PayslipVault === 'undefined' || typeof PayslipGuide === 'undefined') {
    $('libError').classList.remove('hidden');
    return;
  }

  var Schema = PayslipSchema;
  var PDF = PayslipPDF;
  var Delivery = PayslipDelivery;
  var Vault = PayslipVault;
  var Guide = PayslipGuide;
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
    log: [],            // delivery log entries for the current company+period
    parseOptions: { sheet: 'auto', headerRow: null, includeHidden: false },
    bankTotal: '',
    sent: {},           // employeeId -> the delivery journal entry that counts (see PayslipVault.sentMap)
    logKey: '',
    sourceHash: '',     // SHA-256 of the loaded Excel file (or pasted text)
    isSample: false,
    settingsHandle: null,   // File System Access handle of the settings file, kept up to date automatically
    savePointSkipped: ''    // settings fingerprint the user chose not to back up
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
    scheduleBackupCheck();
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
  function goStep(n, opts) {
    if (n > 1 && !state.parsed) return;
    if (n >= 4 && state.step <= 3 && !(opts && opts.force) && shouldOfferSavePoint()) { offerSavePoint(n); return; }
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
    if (n === 1) renderWelcome();
    saveSessionState();
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
  var SAMPLE_NAME = 'Sample data (25 employees)';

  function workbookToArray(wb) {
    return XLSX.write(wb, { type: 'array', bookType: 'xlsx', cellDates: true });
  }

  // The styled template (validation, formulas, notes) needs ExcelJS, loaded only when a template is downloaded.
  var excelJsPromise = null;
  function loadExcelJS() {
    if (window.ExcelJS) return Promise.resolve(window.ExcelJS);
    if (!excelJsPromise) {
      excelJsPromise = new Promise(function (resolve, reject) {
        var sc = document.createElement('script');
        sc.src = 'js/vendor/exceljs.min.js';
        sc.onload = function () { window.ExcelJS ? resolve(window.ExcelJS) : reject(new Error('ExcelJS did not load')); };
        sc.onerror = function () { excelJsPromise = null; reject(new Error('js/vendor/exceljs.min.js is missing')); };
        document.head.appendChild(sc);
      });
    }
    return excelJsPromise;
  }

  function downloadTemplate(opts, fileName, fallback) {
    return loadExcelJS().then(function (ExcelJS) {
      return PayslipTemplate.build(ExcelJS, opts);
    }).then(function (buf) {
      downloadBlob(new Blob([buf], { type: XLSX_MIME }), fileName);
    }).catch(function (e) {
      // plain template without styling if the styled one cannot be built
      if (fallback) fallback();
      else alert('Could not build the template: ' + e.message);
    });
  }

  function downloadEmptyTemplate() {
    return downloadTemplate({}, 'Payslip_Template.xlsx', function () {
      downloadBlob(new Blob([workbookToArray(Schema.buildTemplateWorkbook(XLSX))], { type: XLSX_MIME }), 'Payslip_Template.xlsx');
    });
  }
  function downloadExampleTemplate() {
    return downloadTemplate({ records: Schema.sampleRecords(25), sample: true }, 'Sample_Payslips.xlsx', function () {
      downloadBlob(new Blob([workbookToArray(Schema.buildTemplateWorkbook(XLSX, { sampleRows: Schema.sampleRows(25) }))], { type: XLSX_MIME }), 'Sample_Payslips.xlsx');
    });
  }
  $('btnTemplate').addEventListener('click', downloadEmptyTemplate);
  $('btnSampleXlsx').addEventListener('click', downloadExampleTemplate);
  $('btnLoadSample').addEventListener('click', function () {
    loadWorkbookData(workbookToArray(Schema.buildTemplateWorkbook(XLSX, { sampleRows: Schema.sampleRows(25) })), SAMPLE_NAME);
  });
  $('btnTour').addEventListener('click', function () { openTour(); });

  // Template pre-filled with the company details and the logo/stamp pictures.
  $('btnMyTemplate').addEventListener('click', function () {
    var name = PDF.safeFileName('Payslip_Template - ' + (state.company.companyName || 'my company')) + '.xlsx';
    downloadTemplate({ settings: state.company, logo: effectiveImage('logo'), stamp: effectiveImage('stamp') }, name, function () {
      var buf = workbookToArray(Schema.buildTemplateWorkbook(XLSX, { settings: state.company }));
      var images = [];
      var logo = effectiveImage('logo');
      var stamp = effectiveImage('stamp');
      if (logo) images.push({ dataUrl: logo.dataUrl, width: logo.width, height: logo.height, name: 'Logo', row: 1, col: 4, maxPx: 220 });
      if (stamp) images.push({ dataUrl: stamp.dataUrl, width: stamp.width, height: stamp.height, name: 'Stamp', row: 8, col: 4, maxPx: 160 });
      XlsxImages.embedImages(JSZip, buf, Schema.SHEET_SETTINGS, images).then(function (out) {
        downloadBlob(new Blob([out], { type: XLSX_MIME }), name);
      });
    });
  });

  // Paste cells straight from Excel (last resort when the file itself cannot be used).
  function readPasted(text) {
    return XLSX.read(text, { type: 'string', FS: text.indexOf('\t') >= 0 ? '\t' : ',', cellDates: true, raw: false });
  }

  $('btnUsePaste').addEventListener('click', function () {
    var text = $('pasteBox').value;
    if (!text.trim()) { setAlert($('loadError'), 'Paste the cells first (select them in Excel, Ctrl+C, then Ctrl+V here).', 'warn'); return; }
    try {
      var wb = readPasted(text);
      rememberFile({ kind: 'paste', name: 'Pasted cells', text: text });
      loadWorkbookObject(wb, 'Pasted cells', null);
    } catch (e) {
      setAlert($('loadError'), 'These cells could not be read: ' + escapeHtml(e.message), 'error');
    }
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
    if (/delivery log.*\.csv$/i.test(file.name) && state.parsed) return restoreLogFile(file);
    readFile(file).then(function (buf) { loadWorkbookData(buf, file.name); })
      .catch(function (e) { setAlert($('loadError'), escapeHtml(e.message), 'error'); });
  }

  function loadWorkbookData(arrayBuffer, name) {
    $('loadError').innerHTML = '';
    var wb;
    try {
      wb = XLSX.read(arrayBuffer, { type: 'array', cellDates: true, cellStyles: true });
    } catch (e) {
      var msg = /password|encrypt/i.test(e.message || '') ? 'This file is password protected. Open it in Excel, remove the password (File ▸ Info ▸ Protect Workbook) or Save As a new file, then try again.' : 'This file could not be opened as a spreadsheet. ' + escapeHtml(e.message || '');
      setAlert($('loadError'), msg, 'error');
      return;
    }
    rememberFile({ kind: 'file', name: name, bytes: arrayBuffer });
    loadWorkbookObject(wb, name, arrayBuffer);
  }

  // opts.restore: the saved session (step, sheet, toggles) when coming back after a refresh or restart.
  function loadWorkbookObject(wb, name, arrayBuffer, opts) {
    var restore = (opts && opts.restore) || null;
    $('loadError').innerHTML = '';
    state.parseOptions = { sheet: 'auto', headerRow: null, includeHidden: false };
    state.bankTotal = '';
    state.wb = wb;
    state.fileName = name;
    state.isSample = name === SAMPLE_NAME;
    state.excelImages = { logo: null, stamp: null };
    state.dirHandle = null;
    state.exportedTo = '';
    $('exportMessage').innerHTML = '';
    $('sendMessage').innerHTML = '';
    if (restore) {
      var po = restore.parseOptions || {};
      state.parseOptions = { sheet: typeof po.sheet === 'string' ? po.sheet : 'auto', headerRow: po.headerRow > 0 ? +po.headerRow : null, includeHidden: !!po.includeHidden };
      state.bankTotal = restore.bankTotal || '';
      state.exportedTo = restore.exportedTo || '';
      if (typeof restore.skipErrors === 'boolean') $('skipErrors').checked = restore.skipErrors;
      if (typeof restore.protectPdfs === 'boolean') $('protectPdfs').checked = restore.protectPdfs;
    }
    reparse();
    if (arrayBuffer) XlsxImages.extractImages(JSZip, arrayBuffer).then(function (images) {
      if (!images.length) return;
      var onSettings = images.filter(function (i) { return Schema.normalizeHeader(i.sheet) === Schema.normalizeHeader(Schema.SHEET_SETTINGS); });
      var pick = onSettings.length ? onSettings : images;
      var tasks = [];
      if (pick[0]) tasks.push(normalizeImage(XlsxImages.toDataUrl(pick[0]), pick[0].mime === 'image/jpeg').then(function (img) { state.excelImages.logo = img; }, function () {}));
      if (pick[1]) tasks.push(normalizeImage(XlsxImages.toDataUrl(pick[1]), pick[1].mime === 'image/jpeg').then(function (img) { state.excelImages.stamp = img; }, function () {}));
      return Promise.all(tasks).then(function () { renderImages(); invalidate(); scheduleBackupCheck(); });
    }).catch(function () { /* csv / xls / no pictures */ });
    goStep(restore ? Math.min(5, Math.max(2, restore.step || 2)) : 2, { force: !!restore });
  }

  function reparse() {
    var o = state.parseOptions;
    state.parsed = Schema.parseWorkbook(XLSX, state.wb, { overrides: state.profile.columnOverrides, sheet: o.sheet, headerRow: o.headerRow, includeHidden: o.includeHidden });
    mergeCompany();
    loadLog();
    invalidate();
    $('fileName').textContent = state.fileName + (state.parsed.sheetName ? ' · ' + (state.parsed.sheetMode === 'combined' ? state.parsed.sheetName.toLowerCase() : 'sheet "' + state.parsed.sheetName + '"') : '');
    renderSourceBar();
    renderReview();
    populatePreviewSelect();
    saveSessionState();
  }

  // Which sheet, which header row, hidden rows, saved setup: the controls for files that are not our template.
  function renderSourceBar() {
    var p = state.parsed;
    var o = state.parseOptions;
    var bar = $('sourceBar');
    var parts = [];
    var dataSheets = p.sheets.filter(function (sh) { return sh.isData || sh.used; });
    if (p.sheets.length > 1) {
      var current = p.sheetMode === 'combined' ? '__all__' : (p.sheets.filter(function (sh) { return sh.used; })[0] || {}).name;
      var opts = '';
      if (dataSheets.length > 1) opts += '<option value="__all__"' + (current === '__all__' ? ' selected' : '') + '>All sheets combined</option>';
      p.sheets.forEach(function (sh) {
        opts += '<option value="' + escapeHtml(sh.name) + '"' + (current === sh.name ? ' selected' : '') + '>' + escapeHtml(sh.name) + ' · ' + (sh.used ? sh.employees + ' employees' : sh.rows + ' rows') + '</option>';
      });
      parts.push('<label class="src-ctl"><span>Sheet</span><select id="srcSheet">' + opts + '</select></label>');
    }
    if (p.sheetMode !== 'combined' && p.headerPreview && p.headerPreview.length) {
      var hopts = p.headerPreview.map(function (h) {
        return '<option value="' + h.row + '"' + (h.row === p.headerRow ? ' selected' : '') + '>Row ' + h.row + ': ' + escapeHtml(h.text) + '</option>';
      }).join('');
      parts.push('<label class="src-ctl"><span>Header row</span><select id="srcHeader">' + hopts + '</select></label>');
    }
    if (p.stats.hiddenRows || o.includeHidden) {
      parts.push('<label class="check src-ctl"><input type="checkbox" id="srcHidden"' + (o.includeHidden ? ' checked' : '') + '> Include hidden rows' + (p.stats.hiddenRows ? ' (' + p.stats.hiddenRows + ')' : '') + '</label>');
    }
    var manual = p.columns.filter(function (c) { return c.source === 'manual'; }).length;
    if (manual) {
      parts.push('<span class="badge info" title="Your column choices from an earlier upload or settings file">Saved column setup applied · ' + plural(manual, 'column') + (state.profile.savedAt ? ' · ' + Schema.formatDate(new Date(state.profile.savedAt)) : '') + '</span>');
    }
    bar.innerHTML = parts.join('');
    bar.classList.toggle('hidden', !parts.length);
    if ($('srcSheet')) $('srcSheet').addEventListener('change', function () { o.sheet = this.value; o.headerRow = null; reparse(); });
    if ($('srcHeader')) $('srcHeader').addEventListener('change', function () {
      var used = p.sheets.filter(function (sh) { return sh.used; })[0];
      o.sheet = used ? used.name : 'auto';
      o.headerRow = +this.value;
      reparse();
    });
    if ($('srcHidden')) $('srcHidden').addEventListener('change', function () { o.includeHidden = this.checked; reparse(); });
  }

  // Totals to compare with the bank / WPS transfer: the one check that catches any silent problem.
  function renderReconcile() {
    var st = state.parsed.stats;
    var cur = escapeHtml(state.company.currency || 'SAR');
    var bank = Schema.parseNumber(state.bankTotal).value;
    var status = '';
    if (bank != null) {
      var diff = Schema.round2(bank - st.totalNet);
      status = Math.abs(diff) < 0.01
        ? '<span class="rec-ok">✓ Matches the bank transfer</span>'
        : '<span class="rec-bad">Difference ' + cur + ' ' + Schema.formatMoney(diff) + ': look for a missing employee, a skipped row or an unmapped column.</span>';
    }
    $('reconcile').innerHTML =
      '<div class="rec-figs"><div><span class="k">Earnings</span><span class="v">' + Schema.formatMoney(st.totalGross) + '</span></div>' +
      '<div><span class="k">Deductions</span><span class="v">' + Schema.formatMoney(st.totalDeductions) + '</span></div>' +
      '<div><span class="k">Net to pay</span><span class="v strong">' + cur + ' ' + Schema.formatMoney(st.totalNet) + '</span></div></div>' +
      '<label class="rec-input"><span>Bank / WPS transfer total</span><input type="text" id="bankTotal" inputmode="decimal" placeholder="e.g. ' + Schema.formatMoney(st.totalNet) + '" value="' + escapeHtml(state.bankTotal) + '"></label>' +
      '<div class="rec-status">' + status + '</div>';
    $('bankTotal').addEventListener('change', function () { state.bankTotal = this.value; renderReconcile(); saveSessionState(); });
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
    if (k === 'companyName' || k === 'payPeriod' || k === 'periodStart') scheduleRekey();
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

    renderReconcile();
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
        : c.source === 'silent' ? '<span class="badge neutral">helper column</span>'
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
        '<td>' + escapeHtml(s.rowRef || s.rowNumber) + '</td><td>' + badge + '</td>' +
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

  $('btnSaveSettings').addEventListener('click', function () { saveSettingsFile(); });

  function loadSettingsFile(file) {
    return readFile(file, 'text').then(function (text) {
      var obj = JSON.parse(text);
      if (!obj || obj.app !== 'payslip-generator') throw new Error('This is not a Payslip Generator settings file.');
      applyProfile(obj);
      state.excelImages = { logo: null, stamp: null };
      if (state.parsed) reparse();
      renderDesignControls();
      updateSavedHint();
      markBackedUp('loaded');
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

  $('skipErrors').addEventListener('change', function () { renderExportPanel(); saveSessionState(); });
  $('protectPdfs').addEventListener('change', function () { state.cache = {}; renderExportPanel(); saveSessionState(); });

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
      saveSessionState();
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
      state.dirWritable = true;
      state.exportedTo = 'folder';
      saveSessionState();
      Vault.set('session-dir', { dir: sub, hash: state.sourceHash }).catch(function () { /* handle not storable in this browser */ });
      journalToDisk();
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

  // The delivery journal is append-only (sent / undo / clear entries). It lives in localStorage (written
  // synchronously, so a tick survives a crash right after the click) with a copy in IndexedDB, and every
  // write merges what is stored first, so two tabs or a restored session never erase each other's ticks.
  function setLog(entries) {
    state.log = entries;
    state.sent = Vault.sentMap(entries);
  }

  function loadLog() {
    var key = logKey();
    state.logKey = key;
    setLog(Vault.normalizeLog(Vault.readJson(key, [])));
    Vault.get('log:' + key).then(function (copy) {
      if (!copy || key !== state.logKey) return;
      var merged = Vault.mergeLogs(state.log, copy);
      if (merged.length !== state.log.length) { setLog(merged); saveLog(); refreshSendViews(); }
    }).catch(function () { /* no IndexedDB */ });
  }

  function saveLog() {
    var key = state.logKey || logKey();
    setLog(Vault.mergeLogs(Vault.readJson(key, []), state.log));
    if (!Vault.writeJson(key, state.log)) {
      Guide.toast({ id: 'logFull', kind: 'warn', html: '<b>Browser storage is full.</b> Ticks are kept in this tab only: download the delivery log now.', timeout: 0,
        actions: [{ label: 'Download log', primary: true, onClick: function () { $('btnLogCsv').click(); } }] });
    }
    Vault.set('log:' + key, state.log).catch(function () { /* memory only */ });
    journalToDisk();
    saveSessionState();
  }

  function addLogEntry(entry) {
    entry.timestamp = new Date().toISOString();
    entry.source = (state.sourceHash || '').slice(0, 12);
    state.log = state.log.concat([entry]);
    saveLog();
  }

  function isSent(slip) {
    return !!state.sent[slip.employeeId];
  }

  function refreshSendViews() {
    if (!state.parsed) return;
    if (state.step === 5) {
      renderSendList();
      renderSummaryChips();
      updateNextMeter();
      renderSendNotice();
    }
  }

  // Company name or period edited while sending: the ticks follow (after typing stops).
  var rekeyTimer = null;
  function scheduleRekey() {
    clearTimeout(rekeyTimer);
    rekeyTimer = setTimeout(function () {
      if (!state.parsed || logKey() === state.logKey) return;
      var carry = state.log;
      loadLog();
      if (!state.log.length && Object.keys(Vault.sentMap(carry)).length) {
        setLog(carry.slice());
        saveLog();
      }
      refreshSendViews();
    }, 800);
  }

  // After a folder export the journal is also written next to the PDFs, so even a full browser reset
  // cannot lose who already received a payslip.
  var journalTimer = null;
  function journalToDisk() {
    var dir = state.dirHandle;
    if (!dir || !state.log.length) return;
    clearTimeout(journalTimer);
    journalTimer = setTimeout(function () {
      var check = dir.queryPermission ? dir.queryPermission({ mode: 'readwrite' }) : Promise.resolve('granted');
      check.then(function (perm) {
        state.dirWritable = perm === 'granted';
        if (!state.dirWritable) { renderSummaryChips(); return; }
        return writeToDir(dir, 'delivery-log.csv', Delivery.deliveryLogCsv(state.log));
      }).catch(function () { state.dirWritable = false; });
    }, 400);
  }

  function reconnectFolder() {
    var dir = state.dirHandle;
    if (!dir || !dir.requestPermission) return;
    dir.requestPermission({ mode: 'readwrite' }).then(function (perm) {
      state.dirWritable = perm === 'granted';
      if (state.dirWritable) {
        journalToDisk();
        Guide.toast({ kind: 'ok', html: 'Folder reconnected. Every tick is written to <b>delivery-log.csv</b> again.' });
      }
      renderSummaryChips();
    }).catch(function () { /* refused */ });
  }

  function restoreLogFile(file) {
    return readFile(file, 'text').then(function (text) {
      var entries = Vault.logFromCsv(text);
      var fromCsv = Vault.sentMap(entries);
      // whoever the log shows as sent and is not ticked here gets a fresh tick (so a later "clear" or
      // a reset browser cannot hide it), with the original date kept in the note
      var now = Date.now();
      var restored = Object.keys(fromCsv).filter(function (id) { return !state.sent[id]; }).map(function (id, k) {
        var e = fromCsv[id];
        return { action: 'sent', timestamp: new Date(now + k).toISOString(), employeeId: id, name: e.name, channel: e.channel, contact: e.contact, file: e.file, hash: e.hash, source: e.source,
          note: 'restored from delivery log (sent ' + String(e.timestamp).slice(0, 16).replace('T', ' ') + (e.note ? ', ' + e.note : '') + ')' };
      });
      setLog(Vault.mergeLogs(state.log, restored));
      saveLog();
      refreshSendViews();
      Guide.toast({ kind: 'ok', html: '<b>Delivery log read.</b> ' + (restored.length ? plural(restored.length, 'employee is', 'employees are') + ' marked as sent again.' : 'Everyone in it was already marked as sent.') });
    }).catch(function (e) {
      Guide.toast({ kind: 'warn', html: 'Could not read that log: ' + escapeHtml(e.message) });
    });
  }

  function markSent(slip, channel, note) {
    var art = state.cache[Object.keys(state.cache).filter(function (k) { return k.indexOf(slip.employeeId + '|') === 0; })[0]];
    addLogEntry({
      action: 'sent',
      employeeId: slip.employeeId,
      name: slip.name,
      channel: channel,
      contact: channel === 'whatsapp' ? slip.fields.mobile : slip.fields.email,
      file: art ? art.fileName : Schema.slipFileName(slip, state.company),
      hash: art ? art.hash : '',
      note: note || ''
    });
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
    addLogEntry({ action: 'undo', employeeId: slip.employeeId, name: slip.name, channel: currentChannel().id, note: 'tick removed' });
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
      state.exportedTo ? '<span class="chip ok">PDFs exported</span>' : '<span class="chip warn">PDFs not exported yet</span>',
      state.dirHandle ? (state.dirWritable ? '<span class="chip ok" title="Every tick is also written to delivery-log.csv in the export folder">Log saved in folder</span>'
        : '<button type="button" class="chip warn" id="btnReconnect" title="After a restart the browser needs your OK to write to the folder again">Reconnect folder for the log</button>') : ''
    ].join('');
    if ($('btnReconnect')) $('btnReconnect').addEventListener('click', reconnectFolder);
  }

  function renderSendPanel() {
    if (!state.parsed) return;
    var slips = exportableSlips();
    renderSummaryChips();
    renderSendNotice();
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
    var sending = act === 'eml' || act === 'share' || act === 'wa' || act === 'gmail' || act === 'mailto';
    if (sending && isSent(slip)) {
      // already sent (maybe before a refresh or in another tab): ask before making a second email
      e.preventDefault();
      confirmResend(slip).then(function (yes) {
        if (!yes) return;
        if (act === 'eml') downloadDraft(slip);
        else if (act === 'share') shareSlip(slip);
        else {
          window.open(el.href, act === 'mailto' ? '_self' : '_blank', 'noopener');
          markSent(slip, act === 'wa' ? 'whatsapp' : 'email', 'sent again');
        }
      });
      return;
    }
    if (act === 'eml') {
      downloadDraft(slip);
      return;
    }
    if (act === 'share') {
      e.preventDefault();
      shareSlip(slip);
      return;
    }
    if (act === 'wa' || act === 'gmail' || act === 'mailto') {
      // let the anchor open, then record it
      setTimeout(function () { markSent(slip, act === 'wa' ? 'whatsapp' : 'email', act === 'wa' ? 'chat opened' : 'compose opened'); }, 50);
    }
  });

  function shareSlip(slip) {
    var a2 = artifactFor(slip);
    var m2 = messageFor(slip);
    var file = new File([a2.bytes], a2.fileName, { type: 'application/pdf' });
    navigator.share({ files: [file], title: m2.subject, text: m2.subject + '\n' + m2.text }).then(function () {
      markSent(slip, 'whatsapp', 'shared');
    }).catch(function (err) {
      if (err && err.name !== 'AbortError') setAlert($('sendMessage'), 'Share failed: ' + escapeHtml(err.message) + '. Use "Open chat" and attach the file by hand.', 'warn');
    });
  }

  function confirmResend(slip) {
    var e = state.sent[slip.employeeId] || {};
    var when = e.timestamp ? new Date(e.timestamp) : null;
    var whenText = when ? when.toLocaleDateString(undefined, { day: 'numeric', month: 'short' }) + ', ' + when.toLocaleTimeString(undefined, { hour: '2-digit', minute: '2-digit' }) : 'earlier';
    return Guide.confirm({
      title: 'Already sent',
      html: '<b>' + escapeHtml(slip.name) + '</b> was marked as sent on <b>' + escapeHtml(whenText) + '</b>' + (e.note ? ' (' + escapeHtml(e.note) + ')' : '') + '. Sending again means a second email for this person.',
      ok: 'Send again', cancel: 'Keep one'
    });
  }

  // A warning when some ticks were made while a different version of the Excel file was loaded.
  function renderSendNotice() {
    var el = $('sendNotice');
    var cur = (state.sourceHash || '').slice(0, 12);
    var visible = {};
    exportableSlips().forEach(function (sl) { visible[sl.employeeId] = sl.name; });
    var other = Object.keys(state.sent).filter(function (k) { var e = state.sent[k]; return cur && e.source && e.source !== cur && visible[k] != null; });
    el.innerHTML = other.length
      ? '<div class="alert alert-warn"><strong>' + plural(other.length, 'payslip was', 'payslips were') + ' sent while a different version of the Excel file was open</strong> (' +
        escapeHtml(other.slice(0, 3).map(function (k) { return visible[k]; }).join(', ')) + (other.length > 3 ? '…' : '') + '). If their amounts changed, untick them and send again.</div>'
      : '';
  }

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
    var n = Object.keys(state.sent).length;
    if (!n) return;
    Guide.confirm({ title: 'Clear all ticks?', danger: true, ok: 'Clear ' + plural(n, 'tick'),
      html: 'Every employee of <b>' + escapeHtml(state.company.payPeriod || 'this period') + '</b> will show as not sent. The delivery log keeps the history.' }).then(function (yes) {
      if (!yes) return;
      addLogEntry({ action: 'clear', employeeId: '*', note: 'all ticks cleared' });
      renderSendPanel();
    });
  });
  $('logInput').addEventListener('change', function (e) {
    var f = e.target.files && e.target.files[0];
    e.target.value = '';
    if (f) restoreLogFile(f);
  });

  // ---------------------------------------------------------------------------
  // Settings backup: a file on disk, so clearing the browser or a new computer loses nothing
  // ---------------------------------------------------------------------------
  var ICON_SAVE = '<svg viewBox="0 0 24 24"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><path d="M17 21v-8H7v8M7 3v5h8"/></svg>';
  var ICON_SHIELD = '<svg viewBox="0 0 24 24"><path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/><path d="m9 12 2 2 4-4"/></svg>';

  function settingsHash() {
    return Vault.settingsFingerprint(settingsFileObject());
  }

  function backupStatus() {
    var c = state.profile.company || {};
    var worth = c.companyName || effectiveImage('logo') || Object.keys(state.profile.columnOverrides).length || (state.parsed && !state.isSample);
    if (state.isSample || !worth) return { state: 'empty' };
    var b = Vault.backup();
    if (!b.hash) return { state: 'never' };
    if (b.hash !== settingsHash()) return { state: 'stale', at: b.at, how: b.how };
    return { state: 'ok', at: b.at, how: b.how };
  }

  function markBackedUp(how) {
    Vault.setBackup({ hash: settingsHash(), at: new Date().toISOString(), how: how });
    renderBackupChip(true);
  }

  var backupTimer = null;
  function scheduleBackupCheck() {
    clearTimeout(backupTimer);
    backupTimer = setTimeout(function () {
      if (state.settingsHandle && backupStatus().state === 'stale') autoSaveSettings();
      else renderBackupChip();
    }, 700);
  }

  function renderBackupChip(bump) {
    var chip = $('backupChip');
    var st = backupStatus();
    chip.className = 'top-chip backup-chip state-' + st.state + (st.state === 'empty' ? ' hidden' : '') + (bump ? ' bump' : '');
    if (st.state === 'ok') {
      chip.innerHTML = ICON_SHIELD + '<span class="lbl">' + (st.how === 'auto' ? 'Auto-saved' : 'Backed up') + '</span>';
      chip.title = 'Settings file saved ' + Vault.timeAgo(st.at) + '. Click to see or save again.';
    } else {
      chip.innerHTML = ICON_SAVE + '<span class="lbl">' + (st.state === 'stale' ? 'Save changes' : 'Save settings') + '</span>';
      chip.title = st.state === 'stale' ? 'You changed settings since the last settings file.' : 'Your setup lives only in this browser. Save a settings file.';
    }
    chip.setAttribute('aria-label', chip.title);
    if (state.step === 1) renderWelcome();
  }

  function settingsJson() {
    return JSON.stringify(settingsFileObject(), null, 2);
  }

  function settingsFileName() {
    return PDF.safeFileName('payslip-settings - ' + (state.company.companyName || state.profile.company.companyName || 'my company')) + '.json';
  }

  function writeSettingsTo(handle) {
    return handle.createWritable().then(function (w) {
      return w.write(settingsJson()).then(function () { return w.close(); });
    });
  }

  function autoSaveSettings() {
    var h = state.settingsHandle;
    (h.queryPermission ? h.queryPermission({ mode: 'readwrite' }) : Promise.resolve('denied')).then(function (perm) {
      if (perm !== 'granted') { renderBackupChip(); return; }
      return writeSettingsTo(h).then(function () { markBackedUp('auto'); });
    }).catch(function () { renderBackupChip(); });
  }

  // Chrome/Edge: a Save dialog once, then that same file is kept up to date. Others: a download.
  function saveSettingsFile() {
    var h = state.settingsHandle;
    var viaHandle = h && h.requestPermission
      ? h.requestPermission({ mode: 'readwrite' }).then(function (perm) {
          if (perm !== 'granted') throw new Error('no permission');
          return writeSettingsTo(h).then(function () { return 'file'; });
        })
      : Promise.reject(new Error('no handle'));
    function download() {
      downloadBlob(new Blob([settingsJson()], { type: 'application/json' }), settingsFileName());
      return 'download';
    }
    return viaHandle.catch(function () {
      if (typeof window.showSaveFilePicker !== 'function') return download();
      return window.showSaveFilePicker({ suggestedName: settingsFileName(), id: 'payslip-settings', types: [{ description: 'Payslip settings', accept: { 'application/json': ['.json'] } }] }).then(function (handle) {
        return writeSettingsTo(handle).then(function () {
          state.settingsHandle = handle;
          Vault.set('settings-handle', handle).catch(function () { /* handle not storable */ });
          return 'file';
        });
      }, function (err) {
        return err && err.name === 'AbortError' ? null : download();
      });
    }).then(function (how) {
      if (!how) return false;
      markBackedUp(how);
      Guide.toast({ id: 'savedToast', kind: 'save', icon: '💾', html: how === 'download'
        ? '<b>Settings file saved</b> to your Downloads. Keep it with your payroll files; next month just load it.'
        : '<b>Settings file saved.</b> It now updates itself whenever you change something.' });
      return true;
    });
  }

  $('backupChip').addEventListener('click', function () { openSavePoint(null); });

  // ---------------------------------------------------------------------------
  // Save point: before exporting, offer to save the setup (once per set of changes)
  // ---------------------------------------------------------------------------
  function shouldOfferSavePoint() {
    var st = backupStatus().state;
    return (st === 'never' || st === 'stale') && state.savePointSkipped !== settingsHash();
  }

  function savePointItems() {
    var c = state.company && Object.keys(state.company).length ? state.company : (state.profile.company || {});
    var t = PDF.TEMPLATES.filter(function (x) { return x.id === state.profile.templateId; })[0];
    var hiddenOpts = PDF.DISPLAY_OPTIONS.filter(function (o) { return !state.profile.display[o.key]; }).length;
    var cols = Object.keys(state.profile.columnOverrides).length;
    var logo = effectiveImage('logo');
    return [
      { label: 'Company', value: [c.companyName, c.crNumber ? 'CR ' + c.crNumber : ''].filter(Boolean).join(' · ') || 'not filled in yet', ok: !!c.companyName },
      { label: 'Logo', value: logo ? 'included' : 'none yet', ok: !!logo, thumb: logo ? logo.dataUrl : '' },
      { label: 'Design', value: (t ? t.name : '') + ' · ' + String(state.profile.paper).toUpperCase() + (hiddenOpts ? ' · ' + plural(hiddenOpts, 'item') + ' hidden' : ''), ok: true },
      { label: 'Column choices', value: cols ? plural(cols, 'column') + ' remembered' : 'none needed', ok: cols > 0 },
      { label: 'Email message', value: c.messageText && c.messageText !== currentDefault('messageText') ? 'your own text' : 'standard text', ok: true }
    ];
  }

  function openSavePoint(target) {
    var st = backupStatus();
    var fsa = typeof window.showSaveFilePicker === 'function';
    return Guide.savePoint({
      title: st.state === 'ok' ? 'Your setup is backed up' : target ? (st.state === 'stale' ? 'You changed a few things. Save them first?' : 'Save point! Keep your setup safe') : 'Save your setup',
      text: st.state === 'ok'
        ? 'Saved ' + escapeHtml(Vault.timeAgo(st.at)) + (st.how === 'auto' ? ' and kept up to date automatically' : '') + '. Save again any time.'
        : 'Right now your company details, logo and design live only in this browser. Clearing it or switching computers would lose them. One small file keeps it all for next month.',
      items: savePointItems(),
      saveLabel: st.state === 'ok' ? 'Save again' : 'Save settings file',
      skipLabel: target ? 'Continue without saving' : 'Not now',
      note: fsa ? 'Pick a place once, for example your payroll folder. After that the file updates itself.' : 'The file goes to your Downloads folder. Move it next to your payroll files.',
      onSave: saveSettingsFile
    }).then(function (r) {
      if (!target) return;
      if (r === 'skip') state.savePointSkipped = settingsHash();
      if (r === 'skip' || r === 'saved') goStep(target, { force: true });
    });
  }

  function offerSavePoint(target) {
    openSavePoint(target);
  }

  // ---------------------------------------------------------------------------
  // Session: the open file and where the user was, kept in IndexedDB so a refresh, a crash or a
  // browser restart (even mid-sending) brings everything back exactly as it was.
  // ---------------------------------------------------------------------------
  var AUTO_RESUME_MS = 12 * 3600 * 1000;

  function rememberOn() {
    return Vault.prefs().remember !== false;
  }

  function rememberFile(rec) {
    state.sourceRec = rec;
    state.sourceHash = '';
    state.pendingSession = null;
    Vault.del('session-dir').catch(function () {});
    var data = rec.kind === 'paste' ? new TextEncoder().encode(rec.text) : rec.bytes;
    sha256Hex(data).then(function (hash) {
      state.sourceHash = hash;
      rec.hash = hash;
      return storeSessionFile(rec);
    }).catch(function () { /* storage blocked: the app still works, it just cannot resume */ });
  }

  function storeSessionFile(rec) {
    if (!rec || !rememberOn()) return Promise.resolve();
    return Vault.saveSessionFile({ savedAt: new Date().toISOString(), kind: rec.kind, name: rec.name, bytes: rec.kind === 'file' ? rec.bytes : null, text: rec.kind === 'paste' ? rec.text : null, hash: rec.hash || '' })
      .then(function () { saveSessionState(); requestPersistence(); });
  }

  var sessionTimer = null;
  function saveSessionState() {
    if (!state.parsed || !rememberOn()) return;
    clearTimeout(sessionTimer);
    sessionTimer = setTimeout(writeSessionState, 250);
  }

  function writeSessionState() {
    clearTimeout(sessionTimer);
    if (!state.parsed || !rememberOn()) return Promise.resolve();
    var slips = exportableSlips();
    return Vault.saveSessionState({
      savedAt: new Date().toISOString(),
      fileName: state.fileName,
      step: state.step,
      parseOptions: state.parseOptions,
      skipErrors: $('skipErrors').checked,
      protectPdfs: $('protectPdfs').checked,
      bankTotal: state.bankTotal,
      exportedTo: state.exportedTo,
      sourceHash: state.sourceHash,
      company: state.company.companyName || '',
      period: state.company.payPeriod || '',
      total: slips.length,
      sent: slips.filter(isSent).length
    }).catch(function () { /* memory only */ });
  }

  var persistAsked = false;
  function requestPersistence() {
    if (persistAsked) return;
    persistAsked = true;
    Vault.persist();
  }

  function restoreSession(s, auto) {
    var st = s.state || {};
    var wb;
    try {
      wb = s.file.kind === 'paste' ? readPasted(s.file.text) : XLSX.read(s.file.bytes, { type: 'array', cellDates: true, cellStyles: true });
    } catch (e) {
      Vault.clearSession();
      state.pendingSession = null;
      renderWelcome();
      return;
    }
    state.pendingSession = null;
    state.sourceRec = { kind: s.file.kind, name: s.file.name, bytes: s.file.bytes, text: s.file.text, hash: s.file.hash };
    state.sourceHash = s.file.hash || '';
    loadWorkbookObject(wb, s.file.name, s.file.kind === 'paste' ? null : s.file.bytes, { restore: st });
    Vault.get('session-dir').then(function (rec) {
      if (!rec || !rec.dir || rec.hash !== state.sourceHash) return;
      state.dirHandle = rec.dir;
      state.dirWritable = false;
      return rec.dir.queryPermission ? rec.dir.queryPermission({ mode: 'readwrite' }).then(function (p) { state.dirWritable = p === 'granted'; }) : null;
    }).then(function () { if (state.step === 5) renderSummaryChips(); }).catch(function () { /* no handle */ });
    var sent = Object.keys(state.sent).length;
    Art.say('celebrate', null, { settle: state.step === 5 ? 'send' : 'look', after: 2600 });
    Guide.toast({
      id: 'restored', kind: 'ok', icon: '↺', timeout: 9000,
      html: '<b>Welcome back, nothing was lost.</b> ' + escapeHtml(s.file.name) + ' is open again on the ' + ['', 'Upload', 'Review', 'Design', 'Export', 'Send'][state.step] + ' step' + (sent ? ', ' + plural(sent, 'payslip') + ' already marked as sent' : '') + '.',
      actions: auto ? [{ label: 'Start fresh instead', onClick: discardSession }] : []
    });
  }

  function discardSession() {
    Promise.all([Vault.clearSession(), Vault.del('session-dir')]).then(function () {
      if (state.parsed) location.reload();
      else { state.pendingSession = null; renderWelcome(); }
    });
  }

  window.addEventListener('beforeunload', function (e) {
    writeSessionState();
    if (state.exporting) { e.preventDefault(); e.returnValue = ''; return ''; }
  });
  document.addEventListener('visibilitychange', function () {
    if (document.visibilityState === 'hidden') writeSessionState();
  });

  // Another tab ticked someone: show it here too.
  window.addEventListener('storage', function (e) {
    if (!e.key || e.key !== state.logKey) return;
    setLog(Vault.mergeLogs(state.log, Vault.readJson(e.key, [])));
    refreshSendViews();
  });

  // Two tabs (a restored browser often reopens one more) could send the same payslips twice.
  (function watchOtherTabs() {
    if (typeof BroadcastChannel === 'undefined') return;
    var id = Math.random().toString(36).slice(2);
    var warned = false;
    var bc;
    try { bc = new BroadcastChannel('payslip-generator'); } catch (e) { return; }
    bc.onmessage = function (ev) {
      var d = ev.data || {};
      if (d.id === id) return;
      if (d.type === 'hello') bc.postMessage({ type: 'here', id: id });
      if ((d.type === 'hello' || d.type === 'here') && !warned) {
        warned = true;
        Guide.toast({ id: 'otherTab', kind: 'warn', timeout: 12000, html: '<b>Also open in another tab.</b> Ticks stay in sync between tabs, but send from one tab only so nobody gets two emails.' });
      }
    };
    bc.postMessage({ type: 'hello', id: id });
  })();

  // ---------------------------------------------------------------------------
  // Welcome card (returning visitors), quick start (first visit), data panel
  // ---------------------------------------------------------------------------
  function ringSvg(pct, label) {
    var c = 2 * Math.PI * 24;
    return '<svg class="wl-ring" viewBox="0 0 58 58" aria-hidden="true"><circle class="trk" cx="29" cy="29" r="24" fill="none" stroke-width="6"/>' +
      '<circle class="val" cx="29" cy="29" r="24" fill="none" stroke-width="6" stroke-linecap="round" stroke-dasharray="' + c.toFixed(1) + '" stroke-dashoffset="' + c.toFixed(1) + '" data-off="' + (c * (1 - Math.max(0, Math.min(1, pct)))).toFixed(1) + '"/>' +
      '<text x="29" y="33.5" text-anchor="middle">' + escapeHtml(label) + '</text></svg>';
  }

  function renderWelcome() {
    var el = $('welcome');
    var s = state.pendingSession;
    var c = state.profile.company || {};
    var html = '';
    if (s && !state.parsed) {
      var st = s.state || {};
      var steps = ['', 'Upload', 'Review', 'Design', 'Export', 'Send'];
      var sending = st.step === 5 && st.total;
      html = ringSvg(sending ? (st.sent || 0) / st.total : (st.step || 2) / 5, sending ? (st.sent || 0) + '/' + st.total : (st.step || 2) + '/5') +
        '<div class="wl-text"><div class="wl-kicker">Pick up where you left off</div><div class="wl-title">' + escapeHtml(s.file.name) + '</div>' +
        '<div class="wl-meta">' + escapeHtml([st.company, st.period, st.step ? steps[st.step] + ' step' : '', sending ? (st.sent || 0) + ' of ' + st.total + ' sent' : ''].filter(Boolean).join(' · ')) +
        ' · saved ' + escapeHtml(Vault.timeAgo(st.savedAt || s.file.savedAt)) + '</div></div>' +
        '<div class="wl-actions"><button type="button" class="btn btn-light" id="btnResume">Resume →</button><button type="button" class="btn" id="btnDiscard">Start new</button></div>';
    } else if (state.returning && (c.companyName || state.profile.logo)) {
      var b = backupStatus();
      var line = b.state === 'ok' ? 'settings file saved ' + escapeHtml(Vault.timeAgo(b.at)) + ' ✓'
        : b.state === 'stale' ? '<span class="warn">changes since your last settings file</span>'
        : b.state === 'never' ? '<span class="warn">no settings file saved yet</span>' : '';
      html = '<div class="wl-art">' + Art.mascot('q-own') + '</div>' +
        '<div class="wl-text"><div class="wl-kicker">Welcome back</div><div class="wl-title">' + escapeHtml(c.companyName || 'Good to see you again') + '</div>' +
        '<div class="wl-meta">Company, logo and design are remembered on this computer' + (line ? ' · ' + line : '') + '</div></div>' +
        (b.state === 'never' || b.state === 'stale' ? '<div class="wl-actions"><button type="button" class="btn btn-light" id="btnWlSave">' + ICON_SAVE.replace('<svg', '<svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="2.2"') + ' Save settings file</button></div>' : '');
    }
    el.innerHTML = html;
    el.classList.toggle('hidden', !html);
    if (!html) return;
    var ring = el.querySelector('.wl-ring .val');
    if (ring) requestAnimationFrame(function () { requestAnimationFrame(function () { ring.style.strokeDashoffset = ring.getAttribute('data-off'); }); });
    if ($('btnResume')) $('btnResume').addEventListener('click', function () { restoreSession(s, false); });
    if ($('btnDiscard')) $('btnDiscard').addEventListener('click', discardSession);
    if ($('btnWlSave')) $('btnWlSave').addEventListener('click', function () { openSavePoint(null); });
  }

  function openTour(start) {
    return Guide.tour({
      start: start || 0,
      onDownload: function (example) { if (example) downloadExampleTemplate(); else downloadEmptyTemplate(); },
      canInstall: function () { return !!(window.PayslipPWA && window.PayslipPWA.canInstall()); },
      onInstall: function () { window.PayslipPWA.install(); }
    }).then(function (result) {
      Vault.updateVisit({ tourDone: true, tourResult: result });
      if (result === 'done') Art.say('idle', null);
    });
  }

  function openDataPanel() {
    var b = backupStatus();
    var prefs = Vault.prefs();
    Promise.all([Vault.persisted(), Vault.loadSession()]).then(function (r) {
      var persisted = r[0];
      var sess = r[1];
      var sentCount = Object.keys(state.sent).length;
      var li = function (cls, icon, title, text) {
        return '<li><span class="di ' + cls + '">' + icon + '</span><div><b>' + title + '</b><span>' + text + '</span></div></li>';
      };
      var m = Guide.modal({
        className: 'data-panel',
        label: 'Your data on this computer',
        html: '<h2>Your data stays on this computer</h2><p class="muted">Nothing is uploaded anywhere. This is what the browser keeps, and why.</p><ul class="data-list">' +
          li(b.state === 'ok' ? 'ok' : b.state === 'empty' ? '' : 'warn', '⚙️', 'Settings', 'Company, logo, design and column choices. ' +
            (b.state === 'ok' ? 'Settings file saved ' + escapeHtml(Vault.timeAgo(b.at)) + '.' : b.state === 'empty' ? 'Nothing set up yet.' : 'No up-to-date settings file yet: save one as a backup.')) +
          li('', '📄', 'Open file and progress', sess ? escapeHtml(sess.file.name) + ': kept up to ' + Vault.SESSION_TTL_DAYS + ' days so a refresh, crash or restart loses nothing.' : prefs.remember ? 'Nothing open yet.' : 'Not kept (switched off below).') +
          li('', '✉️', 'Delivery ticks', state.parsed ? plural(sentCount, 'employee') + ' marked as sent for ' + escapeHtml(state.company.payPeriod || 'this period') + (state.dirHandle && state.dirWritable ? ', also written to delivery-log.csv in your export folder.' : '. Download the delivery log to keep a copy.') : 'Kept per company and pay period, even when you close the browser.') +
          li(persisted ? 'ok' : '', persisted ? '🛡️' : 'ℹ️', 'Protected storage', persisted ? 'The browser will not clear this data by itself.' : 'The browser may clear this data when the disk runs full. The settings file is your backup.') +
          '</ul><label class="check"><input type="checkbox" id="optRemember"' + (prefs.remember ? ' checked' : '') + '><span>Remember my open file and progress on this computer. Turn this off on a shared computer.</span></label>' +
          '<div class="actions"><button type="button" class="btn btn-primary" id="dpSave">Save settings file</button><button type="button" class="btn btn-danger" id="dpErase">Erase everything here</button></div>'
      });
      m.el.querySelector('#optRemember').addEventListener('change', function () {
        Vault.setPrefs({ remember: this.checked });
        if (this.checked) storeSessionFile(state.sourceRec);
        else Promise.all([Vault.clearSession(), Vault.del('session-dir')]);
      });
      m.el.querySelector('#dpSave').addEventListener('click', function () { saveSettingsFile(); });
      m.el.querySelector('#dpErase').addEventListener('click', function () {
        m.close();
        Guide.confirm({
          title: 'Erase everything on this computer?', danger: true, ok: 'Erase everything',
          html: 'Settings, logo, the open file and all delivery ticks are removed from this browser. Files you saved yourself (settings file, PDFs, delivery log) are not touched.'
        }).then(function (yes) {
          if (!yes) return;
          Vault.eraseAll().then(function () { location.reload(); });
        });
      });
    });
  }
  $('privacyBtn').addEventListener('click', openDataPanel);

  // ---------------------------------------------------------------------------
  // Init
  // ---------------------------------------------------------------------------
  Art.init();
  // the quick start only opens by itself if the visitor has not started doing something already
  var interacted = false;
  ['pointerdown', 'keydown', 'drop'].forEach(function (ev) {
    document.addEventListener(ev, function () { interacted = true; }, { capture: true, once: true });
  });
  var visitInfo = Vault.touchVisit(new Date().toISOString());
  var hadProfile = !!Vault.readJson(STORAGE_KEY, null);
  loadProfile();
  state.returning = !visitInfo.first || hadProfile;
  renderDesignControls();
  bindImage('logo');
  bindImage('stamp');
  updateSavedHint();
  goStep(1);
  renderBackupChip();
  Vault.get('settings-handle').then(function (h) { if (h && h.createWritable) state.settingsHandle = h; }).catch(function () { /* none */ });
  Vault.loadSession().then(function (s) {
    if (s && rememberOn()) {
      state.pendingSession = s;
      if (s.ageMs < AUTO_RESUME_MS && !state.parsed) { restoreSession(s, true); return; }
    }
    renderWelcome();
    if (!state.returning && !s) {
      setTimeout(function () {
        if (interacted || state.parsed || document.querySelector('.modal-backdrop')) return;
        openTour();
      }, 700);
    }
  });

  window.PayslipApp = {
    state: state,
    handleFile: handleFile,
    flushSession: writeSessionState,
    openTour: openTour,
    loadWorkbookData: loadWorkbookData,
    exportableSlips: exportableSlips,
    artifactFor: artifactFor,
    messageFor: messageFor,
    goStep: goStep,
    applyProfile: applyProfile
  };
})();
