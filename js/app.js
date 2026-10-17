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
  var SETTINGS_FILE_VERSION = 3;
  // Fields that change every month; never kept in the company profile.
  var MONTHLY_KEYS = ['payPeriod', 'periodStart', 'periodEnd', 'paymentDate'];
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
      lock: { whatsapp: true, email: false, download: true },
      channel: 'outlook',
      locked: false,      // true once the company profile is set up: design and details stay fixed
      version: 0,
      lastPeriod: '',
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
    editing: null,      // { snapshot, returnStep } while the company profile is being edited
    monthly: {},        // this month's pay period and dates (not part of the profile)
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
    if (p.lock && typeof p.lock === 'object') ['whatsapp', 'email', 'download'].forEach(function (k) { if (typeof p.lock[k] === 'boolean') base.lock[k] = p.lock[k]; });
    // the pay period and dates belong to a month, not to the profile
    base.lastPeriod = p.lastPeriod || base.company.payPeriod || '';
    MONTHLY_KEYS.forEach(function (k) { delete base.company[k]; });
    if (opts.fromFile) {
      base.locked = true;
      base.version = p.profileVersion || p.version || 1;
    } else {
      // profiles saved before the profile existed as a concept: in use already, so treat them as set up
      base.locked = typeof p.locked === 'boolean' ? p.locked : !!base.company.companyName;
      base.version = p.version || (base.locked ? 1 : 0);
    }
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
    // with a company profile the design is fixed: Review goes straight to Send and back
    if (n === 3 && profileLocked()) n = state.step >= 4 ? 2 : 4;
    var prev = state.step;
    state.step = n;
    document.body.setAttribute('data-step', String(n));
    for (var i = 1; i <= 4; i++) {
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
    if (n === 4) renderSendPanel();
    if (n === 1) renderStartCards();
    saveSessionState();
  }

  // The mascot's line for each step.
  function buddyForStep(n) {
    if (!state.parsed) { Art.say('idle', null); return; }
    var st = state.parsed.stats;
    if (n === 2) {
      var errs = state.parsed.issues.filter(function (i) { return i.level === 'error'; }).length;
      if (errs || st.withErrors) Art.say('oops', 'I checked every row. <b>' + (errs ? errs + ' file problem' + (errs === 1 ? '' : 's') : st.withErrors + ' row' + (st.withErrors === 1 ? '' : 's')) + '</b> need a look; the rest are fine.');
      else if (companyMissing().length) Art.say('look', '<b>' + st.ok + ' payslips</b> look good. ' + (profileLocked() ? 'Just enter the <b>pay period</b>.' : 'Just tell me the company details below.'));
      else Art.say('look', 'All <b>' + st.ok + ' payslips</b> check out. Total net <b>' + Schema.formatMoney(st.totalNet) + ' ' + escapeHtml(state.company.currency || 'SAR') + '</b>.');
    } else if (n === 3) {
      Art.say('design', state.editing ? 'Change what you need, then <b>Save profile</b> (or Cancel).' : state.isSample ? 'Try any look: sample data is never saved.' : 'Pick a look <b>once</b>. It becomes your company profile.');
    } else if (n === 4) {
      var left = exportableSlips().filter(function (s) { return !isSent(s); }).length;
      Art.say('send', left ? 'Open a payslip with <b>View</b>, check it, then send it on <b>WhatsApp</b> or by <b>email</b>. ' + left + ' to go.' : 'Everyone has their payslip.');
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
  $('panel1').addEventListener('click', function (e) {
    var b = e.target.closest('[data-action]');
    if (!b) return;
    var a = b.getAttribute('data-action');
    if (a === 'template') downloadEmptyTemplate();
    else if (a === 'example') downloadExampleTemplate();
    else if (a === 'sample') loadWorkbookData(workbookToArray(Schema.buildTemplateWorkbook(XLSX, { sampleRows: Schema.sampleRows(25) })), SAMPLE_NAME);
    else if (a === 'tour') openTour();
    else if (a === 'mytemplate') downloadMyTemplate();
    else if (a === 'loadprofile') $('profileInput').click();
    else if (a === 'profile') openProfilePanel();
  });
  $('profileInput').addEventListener('change', function (e) {
    var f = e.target.files && e.target.files[0];
    e.target.value = '';
    if (f) loadSettingsFile(f);
  });

  // Template pre-filled with the company details and the logo/stamp pictures.
  function downloadMyTemplate() {
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
  }

  // Paste cells straight from Excel (last resort when the file itself cannot be used).
  function readPasted(text) {
    return Schema.readCsv(XLSX, text);
  }

  // CSV bytes: UTF-8, or the Arabic Windows code page that older Excel versions use.
  function decodeText(buf) {
    try { return new TextDecoder('utf-8', { fatal: true }).decode(buf); } catch (e) { /* not UTF-8 */ }
    try { return new TextDecoder('windows-1256').decode(buf); } catch (e2) { return new TextDecoder().decode(buf); }
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
      wb = /\.(csv|txt)$/i.test(name || '') ? Schema.readCsv(XLSX, decodeText(arrayBuffer)) : XLSX.read(arrayBuffer, { type: 'array', cellDates: true, cellStyles: true });
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
    state.parseOptions = { sheet: 'auto', headerRow: null, includeHidden: false, supplements: {} };
    state.bankTotal = '';
    state.wb = wb;
    state.fileName = name;
    state.isSample = name === SAMPLE_NAME;
    state.excelImages = { logo: null, stamp: null };
    state.dirHandle = null;
    state.exportedTo = '';
    $('exportMessage').innerHTML = '';
    // a new file never inherits a pay period: it comes from the file or is chosen again
    state.monthly = {};
    if (restore) {
      if (restore.monthly && typeof restore.monthly === 'object') state.monthly = restore.monthly;
      var po = restore.parseOptions || {};
      state.parseOptions = { sheet: typeof po.sheet === 'string' ? po.sheet : 'auto', headerRow: po.headerRow > 0 ? +po.headerRow : null, includeHidden: !!po.includeHidden, supplements: po.supplements && typeof po.supplements === 'object' ? po.supplements : {} };
      state.bankTotal = restore.bankTotal || '';
      state.exportedTo = restore.exportedTo || '';
      if (typeof restore.skipErrors === 'boolean') $('skipErrors').checked = restore.skipErrors;
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
    applyMode();
    goStep(restore ? Math.min(4, Math.max(2, restore.step || 2)) : 2, { force: !!restore });
  }

  function reparse() {
    var o = state.parseOptions;
    state.parsed = Schema.parseWorkbook(XLSX, state.wb, { overrides: state.profile.columnOverrides, sheet: o.sheet, headerRow: o.headerRow, includeHidden: o.includeHidden, supplements: o.supplements });
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
    if (p.sheets.length > 1 && (dataSheets.length > 1 || !(p.supplements || []).length || p.sheets.some(function (sh) { return sh.month; }))) {
      var current = p.sheetMode === 'combined' ? '__all__' : (p.sheets.filter(function (sh) { return sh.used; })[0] || {}).name;
      var opts = '';
      if (dataSheets.length > 1) opts += '<option value="__all__"' + (current === '__all__' ? ' selected' : '') + '>All sheets combined</option>';
      p.sheets.forEach(function (sh) {
        opts += '<option value="' + escapeHtml(sh.name) + '"' + (current === sh.name ? ' selected' : '') + '>' + escapeHtml(sh.name) + ' · ' + (sh.used ? sh.employees + ' employees' : sh.role && sh.role !== 'payroll' ? sh.roleLabel.toLowerCase() : sh.rows + ' rows') + '</option>';
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
    var extra = (p.supplements || []);
    if (extra.length) {
      parts.push('<div class="src-extra"><span class="src-lbl">Other sheets</span>' + extra.map(function (x) {
        var what = x.money.concat(x.fields).slice(0, 3).join(', ');
        var detail = x.matched ? (what ? what + ' · ' : '') + x.matched + ' of ' + x.total + ' matched' : x.reason;
        var tip = x.matched ? 'Matched by ' + { id: 'employee number', nid: 'Iqama / ID number', name: 'name' }[x.keyBy] + '. ' + (x.auto ? (x.enabled ? 'On because ' : 'Off because ') + x.reason + '.' : 'Your choice.') + (x.unmatched.length ? ' Not in the payroll: ' + x.unmatched.slice(0, 6).join(', ') + (x.unmatched.length > 6 ? '…' : '') : '') : x.reason;
        var usable = x.matched > 0 && (x.money.length || x.fields.length);
        return '<label class="xsheet' + (x.enabled ? ' on' : '') + (usable ? '' : ' off') + '" title="' + escapeHtml(tip) + '"><input type="checkbox" data-sheet="' + escapeHtml(x.sheet) + '"' + (x.enabled ? ' checked' : '') + (usable ? '' : ' disabled') + '>' +
          '<span class="sw"></span><b>' + escapeHtml(x.sheet) + '</b><span>' + escapeHtml(x.roleLabel) + (detail ? ' · ' + escapeHtml(detail) : '') + '</span></label>';
      }).join('') + (function () {
        var unused = p.sheets.filter(function (sh) { return !sh.used && (sh.role === 'summary' || sh.role === 'other') && sh.rows; });
        return unused.length ? '<span class="muted small">Not used: ' + unused.map(function (sh) { return escapeHtml(sh.name) + ' (' + sh.roleLabel.toLowerCase() + ')'; }).join(', ') + '</span>' : '';
      })() + '</div>');
    }
    bar.innerHTML = parts.join('');
    bar.classList.toggle('hidden', !parts.length);
    Array.prototype.forEach.call(bar.querySelectorAll('[data-sheet]'), function (cb) {
      cb.addEventListener('change', function () {
        o.supplements = Object.assign({}, o.supplements || {});
        o.supplements[cb.getAttribute('data-sheet')] = cb.checked;
        reparse();
      });
    });
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

  // Company details for this file. Before the profile is set up, the Excel Settings sheet fills them in.
  // Once it is set up, the profile wins (so payslips stay the same every month) and the Excel file only
  // supplies this month's pay period and dates.
  function mergeCompany() {
    var excel = upgradeMessageText(state.parsed.settings);
    var saved = upgradeMessageText(state.profile.company || {});
    var fixed = !!state.profile.locked;
    var out = {};
    var differs = [];
    state.companyFromExcel = {};
    Schema.SETTINGS_FIELDS.forEach(function (f) {
      if (f.type === 'image') return;
      var v = excel[f.key];
      var isDefaultOnly = f.default != null && v === f.default && !state.parsed.settingsFound;
      var fromExcel = v !== '' && v != null && !isDefaultOnly;
      if (MONTHLY_KEYS.indexOf(f.key) >= 0) {
        if (fromExcel) { out[f.key] = v; state.companyFromExcel[f.key] = true; }
        else out[f.key] = state.monthly[f.key] || '';
        return;
      }
      var mine = saved[f.key] != null && saved[f.key] !== '' ? saved[f.key] : null;
      if (fixed) {
        out[f.key] = mine != null ? mine : (f.default != null ? f.default : '');
        if (fromExcel && mine != null && f.key !== 'messageText' && f.key !== 'messageSubject' && String(v).trim().toLowerCase() !== String(mine).trim().toLowerCase()) differs.push(f.label);
      } else if (fromExcel) {
        out[f.key] = v;
        state.companyFromExcel[f.key] = true;
      } else {
        out[f.key] = mine != null ? mine : (f.default != null ? f.default : '');
      }
    });
    state.company = out;
    MONTHLY_KEYS.forEach(function (k) { if (out[k]) state.monthly[k] = out[k]; });
    if (differs.length) {
      state.parsed.issues.push({ level: 'info', where: 'Settings', code: 'profile-wins', message: 'The Settings sheet in this file differs from your company profile (' + differs.join(', ') + '). The profile is used, so the payslips look the same every month. Change it with Edit profile if needed.' });
    }
    if (fixed && (state.excelImages.logo || state.excelImages.stamp)) {
      // pictures in the file are ignored once the profile is set up
    }
    if (!fixed) {
      var keep = Object.assign({}, saved, out);
      MONTHLY_KEYS.forEach(function (k) { delete keep[k]; });
      state.profile.company = keep;
      saveProfile();
    }
  }

  function companyMissing() {
    return Schema.SETTINGS_FIELDS.filter(function (f) { return f.required && !String(state.company[f.key] || '').trim(); });
  }

  function effectiveImage(kind) {
    if (state.profile.locked) return state.profile[kind] || null;
    return state.excelImages[kind] || state.profile[kind] || null;
  }

  function setCompanyValue(k, v) {
    state.company[k] = v;
    if (MONTHLY_KEYS.indexOf(k) >= 0) {
      state.monthly[k] = v;
      if (k === 'payPeriod' && v && state.profile.locked) { state.profile.lastPeriod = v; saveProfile(); }
      saveSessionState();
    } else {
      state.profile.company[k] = v;
      saveProfile();
    }
    invalidate();
    if (k === 'companyName' || k === 'payPeriod' || k === 'periodStart') scheduleRekey();
  }

  function profileLocked() {
    return !!(state.profile.locked && !state.editing);
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
      banner = '<strong>' + plural(st.ok, 'payslip') + ' ready.</strong> ' + (profileLocked() ? 'Enter the pay period below.' : 'Please fill in the company details below (' + missing.map(function (f) { return f.label; }).join(', ') + ').');
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
      true
        ? { label: 'Mobile · email', value: st.withMobile + ' · ' + st.withEmail, cls: contactsCls, sub: st.withNoContact ? st.withNoContact + ' with neither' : 'everyone reachable' }
        : { label: 'Total net (' + escapeHtml(state.company.currency || 'SAR') + ')', value: Schema.formatMoney(st.totalNet), money: st.totalNet, cls: '', sub: st.withEmail + ' of ' + st.total + ' have an email' }
    ];
    $('summaryGrid').innerHTML = tiles.map(function (t) {
      var v = t.money != null ? Schema.formatMoney(t.money) : t.value;
      return '<div class="tile ' + t.cls + '"><div class="tile-label">' + t.label + '</div><div class="tile-value">' + v + '</div>' + (t.sub ? '<div class="muted small">' + t.sub + '</div>' : '') + '</div>';
    }).join('');

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

  var MONTH_NAMES = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];

  // Likely pay periods: from the sheet or file name, the month after the last one, or the same as last time.
  function periodSuggestions() {
    var out = [];
    var add = function (text, why) { if (text && !out.some(function (o) { return o.text === text; })) out.push({ text: text, why: why }); };
    var used = state.parsed && state.parsed.sheets.filter(function (sh) { return sh.used && sh.month && sh.month.year; })[0];
    var m = used ? used.month : Schema.monthInfo(String(state.fileName || '').replace(/\.[a-z0-9]+$/i, ''));
    if (m && m.year) add(MONTH_NAMES[m.month - 1] + ' ' + m.year, used ? 'from the sheet name' : 'from the file name');
    var lastText = state.profile.lastPeriod;
    var last = Schema.monthInfo(lastText);
    if (last && last.year) {
      var mo = last.month % 12 + 1;
      add(MONTH_NAMES[mo - 1] + ' ' + (last.year + (last.month === 12 ? 1 : 0)), 'the month after ' + lastText);
    }
    if (lastText) add(lastText, 'same as last time');
    return out;
  }

  function renderCompanyForm() {
    var missing = companyMissing();
    var fixed = profileLocked();
    var c = state.company;
    $('companyMeta').innerHTML = companyMetaHtml();
    $('secCompany').querySelector('.sec-title').textContent = fixed ? 'Pay period' : 'Company details';
    if (fixed) {
      $('companySource').innerHTML = 'Company details, design and message come from your company profile (<b>' + escapeHtml(c.companyName) + '</b>' + (c.crNumber ? ', CR ' + escapeHtml(c.crNumber) : '') + '). Only this month\'s dates are entered here. <button type="button" class="link inline" data-profile-edit>Edit profile</button>';
      $('secCompany').open = true;
    } else {
      $('companySource').textContent = state.editing
        ? 'You are editing the company profile. These details are printed on every payslip.'
        : (state.parsed.settingsFound ? 'Filled from the Settings sheet of the file. ' : '') + 'Entered once: these details become part of your company profile.';
      $('secCompany').open = missing.length > 0 || !!state.editing;
    }

    var form = $('companyForm');
    // Delivery-related settings live on the Send step; keep this form about the company.
    var hidden = { messageSubject: 1, messageText: 1, senderEmail: 1 };
    var fields = Schema.SETTINGS_FIELDS.filter(function (f) { return f.type !== 'image' && !hidden[f.key] && (!fixed || MONTHLY_KEYS.indexOf(f.key) >= 0); });
    var sugs = fixed && !String(c.payPeriod || '').trim() ? periodSuggestions() : [];
    form.innerHTML = fields.map(function (f) {
      var v = c[f.key] || '';
      var isMissing = f.required && !String(v).trim();
      return '<label class="' + (f.key === 'footerNote' || f.key === 'address' ? 'wide' : '') + '"><span>' +
        escapeHtml(f.label) + (f.required ? ' <span class="req">*</span>' : '') +
        (state.companyFromExcel[f.key] ? ' <span class="badge neutral" title="Value from the Excel file">Excel</span>' : '') +
        '</span><input type="text" data-key="' + f.key + '" value="' + escapeHtml(v) + '" placeholder="' + escapeHtml(f.example ? 'e.g. ' + f.example : '') + '" class="' + (isMissing ? 'missing' : '') + '"></label>';
    }).join('') + (sugs.length ? '<div class="wide period-hint"><span class="muted small">Which month is this file?</span>' + sugs.map(function (o, k) {
      return '<button type="button" class="btn btn-small' + (k === 0 ? ' btn-primary' : '') + '" data-use-period="' + escapeHtml(o.text) + '" title="' + escapeHtml(o.why) + '">' + escapeHtml(o.text) + '</button>';
    }).join('') + '<span class="muted small">' + escapeHtml(sugs[0].why) + '</span></div>' : '');
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

  $('secCompany').addEventListener('click', function (e) {
    if (e.target.closest('[data-profile-edit]')) { e.preventDefault(); startEdit(); return; }
    var use = e.target.closest('[data-use-period]');
    if (use) {
      setCompanyValue('payPeriod', use.getAttribute('data-use-period'));
      renderReview();
    }
  });

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
    if (!seenKeys.mobile) optionalMissing.push('Mobile (needed for WhatsApp)');
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
    var src = p.company && Object.keys(p.company).length ? p.company : state.company;
    var company = {};
    Object.keys(src || {}).forEach(function (k) { if (MONTHLY_KEYS.indexOf(k) < 0) company[k] = src[k]; });
    return {
      app: 'payslip-generator',
      kind: 'company-profile',
      version: SETTINGS_FILE_VERSION,
      profileVersion: p.version || 1,
      savedAt: new Date().toISOString(),
      company: company,
      templateId: p.templateId,
      accent: p.accent,
      paper: p.paper,
      display: p.display,
      lock: p.lock,
      logo: p.logo || effectiveImage('logo'),
      stamp: p.stamp || effectiveImage('stamp'),
      columnOverrides: p.columnOverrides
    };
  }

  function loadSettingsFile(file) {
    return readFile(file, 'text').then(function (text) {
      var obj = JSON.parse(text);
      if (!obj || obj.app !== 'payslip-generator') throw new Error('This is not a Payslip Generator company profile.');
      applyProfile(obj, { fromFile: true });
      state.editing = null;
      state.excelImages = { logo: null, stamp: null };
      renderDesignControls();
      applyMode();
      if (state.parsed) reparse();
      markBackedUp('loaded');
      setAlert($('loadError'), 'Company profile loaded: <b>' + escapeHtml((obj.company && obj.company.companyName) || file.name) + '</b>.' + (state.parsed ? '' : ' Now drop this month\'s Excel file.'), 'ok');
      invalidate();
    }).catch(function (e) {
      setAlert($('loadError'), 'Could not load the profile: ' + escapeHtml(e.message), 'error');
      goStep(1);
    });
  }

  // ---------------------------------------------------------------------------
  // PDF generation (shared by the design preview, the viewer, downloads and sending)
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

  // Password per channel: WhatsApp locked by default (phones get shared and seen), email open by
  // default (it goes to a personal inbox), downloads locked. Each one can be switched.
  var LOCK_DEFAULTS = { whatsapp: true, email: false, download: true };
  function lockOn(channel) {
    if (String(state.company.passwordRule || 'id').toLowerCase() === 'none') return false;
    var l = state.profile.lock || {};
    return typeof l[channel] === 'boolean' ? l[channel] : !!LOCK_DEFAULTS[channel];
  }

  function passwordFor(slip, locked) {
    if (!locked) return { password: '', hint: '', hintAr: '', error: null, off: true };
    return Schema.buildPassword(slip, state.company.passwordRule);
  }

  // One employee's PDF, locked or open, cached until something changes.
  function artifactFor(slip, locked) {
    var key = slip.employeeId + '|' + slip.rowNumber + '|' + (locked ? 1 : 0);
    if (state.cache[key]) return state.cache[key];
    var pw = passwordFor(slip, locked);
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

  // The message for a channel; the password line disappears when that channel's PDF is not locked.
  function messageFor(slip, channel) {
    var vars = Schema.messageVars(slip, state.company, passwordFor(slip, lockOn(channel)));
    return {
      subject: Schema.renderTemplate(state.company.messageSubject || 'Payslip {period} - {company}', vars),
      text: Schema.renderTemplate(state.company.messageText || '', vars).replace(/[ \t]+\n/g, '\n').replace(/\n{3,}/g, '\n\n').trim()
    };
  }

  // Unlocked copy, for looking at only.
  function previewBlob(slip) {
    var doc = PDF.createDoc(jsPDF, state.profile.paper);
    PDF.renderPayslip(doc, renderContext(slip));
    PDF.setMeta(doc, state.company, 'Payslip ' + slip.employeeId);
    return doc.output('blob');
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
      if (state.previewUrl) URL.revokeObjectURL(state.previewUrl);
      state.previewUrl = URL.createObjectURL(previewBlob(slip));
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
  // Step 4: send. View each payslip, then send it on WhatsApp or by email.
  // Bulk downloads, printing and records sit under "More".
  // ---------------------------------------------------------------------------
  var fsaSupported = typeof window.showDirectoryPicker === 'function';
  var canShareFiles = (function () {
    try {
      return !!(navigator.canShare && navigator.canShare({ files: [new File([new Blob(['x'], { type: 'application/pdf' })], 'x.pdf', { type: 'application/pdf' })] }));
    } catch (e) { return false; }
  })();

  var ICONS = {
    wa: '<svg viewBox="0 0 24 24"><path d="M20.5 11.6a8.5 8.5 0 0 1-12.6 7.4L3.5 20.5l1.6-4.2a8.5 8.5 0 1 1 15.4-4.7z"/><path d="M9 8.6c.3 2.9 2.6 5.3 5.6 5.8l1.1-1.2-1.8-.9-.8.8c-1-.5-1.9-1.3-2.4-2.4l.8-.8-.9-1.8z"/></svg>',
    mail: '<svg viewBox="0 0 24 24"><rect x="2.5" y="4.5" width="19" height="15" rx="2.5"/><path d="m3 7 9 6.5L21 7"/></svg>',
    download: '<svg viewBox="0 0 24 24"><path d="M12 3.5v12M7 11l5 5 5-5"/><path d="M4.5 20.5h15"/></svg>',
    eye: '<svg viewBox="0 0 24 24"><path d="M2 12s3.6-7 10-7 10 7 10 7-3.6 7-10 7S2 12 2 12z"/><circle cx="12" cy="12" r="3"/></svg>',
    lock: '<svg viewBox="0 0 24 24"><rect x="4.5" y="10.5" width="15" height="10" rx="2.5"/><path d="M8 10.5V8a4 4 0 0 1 8 0v2.5"/></svg>',
    share: '<svg viewBox="0 0 24 24"><circle cx="18" cy="5.5" r="2.5"/><circle cx="6" cy="12" r="2.5"/><circle cx="18" cy="18.5" r="2.5"/><path d="m8.2 13.2 7.6 4.1M15.8 6.7l-7.6 4.1"/></svg>'
  };

  var CHANNEL_NAMES = { whatsapp: 'WhatsApp', outlook: 'email', email: 'email', mailmerge: 'mail merge', manual: 'marked by hand' };

  function refreshExportState() {
    var missing = state.parsed ? companyMissing().filter(function (f) { return f.key !== 'crNumber'; }) : [];
    var slips = exportableSlips();
    state.canSend = !!(state.parsed && slips.length && !missing.length && !state.exporting);
    Array.prototype.forEach.call(document.querySelectorAll('[data-needs-ready]'), function (b) { b.disabled = !state.canSend; });
    var el = $('exportState');
    if (!state.parsed) { el.innerHTML = ''; return; }
    if (missing.length) {
      setAlert(el, 'Fill in ' + missing.map(function (f) { return '<strong>' + f.label + '</strong>'; }).join(' and ') + ' first (on the Review step). You can already look at the payslips.' +
        '<div class="alert-actions"><button type="button" class="btn btn-small" data-go="2">Go to Review</button></div>', 'warn');
      el.querySelector('[data-go]').addEventListener('click', function () { goStep(2); $('secCompany').open = true; });
    } else if (!slips.length) {
      setAlert(el, 'No rows can be sent. Fix the errors in Excel or untick "Skip rows with errors" under More.', 'warn');
    } else {
      el.innerHTML = '';
    }
  }

  function setProgress(done, total, label) {
    $('progress').classList.remove('hidden');
    $('progressFill').style.width = (total ? Math.round((done / total) * 100) : 0) + '%';
    $('progressText').textContent = label || (done + ' / ' + total);
  }

  function beginExport() {
    Art.say('export', 'Building the PDFs…');
    state.exporting = true;
    state.cancelExport = false;
    $('btnCancel').classList.remove('hidden');
    $('exportMessage').innerHTML = '';
    refreshExportState();
  }

  function endExport() {
    state.exporting = false;
    $('btnCancel').classList.add('hidden');
    $('progress').classList.add('hidden');
    refreshExportState();
    if (state.step === 4) renderSendSummary();
  }

  $('btnCancel').addEventListener('click', function () { state.cancelExport = true; });

  function baseName() {
    return PDF.safeFileName(['Payslips', state.company.companyName, state.company.payPeriod].filter(Boolean).join(' - '));
  }

  function reportError(el, e) {
    var cancelled = e && e.message === 'cancelled';
    if (e && (e.name === 'AbortError')) { el.innerHTML = ''; return; }
    setAlert(el, cancelled ? 'Cancelled.' : 'Failed: ' + escapeHtml(e.message), cancelled ? 'warn' : 'error');
  }

  // Every exportable slip's PDF with progress; returns items [{slip, art}]
  async function buildAll(slips, locked) {
    var items = [];
    for (var i = 0; i < slips.length; i++) {
      if (state.cancelExport) throw new Error('cancelled');
      var art = artifactFor(slips[i], locked);
      if (!art.hash) art.hash = await sha256Hex(art.bytes);
      items.push({ slip: slips[i], art: art, fileName: art.fileName, hash: art.hash, password: art.password });
      if (i % 5 === 0 || i === slips.length - 1) {
        setProgress(i + 1, slips.length, 'Making payslip ' + (i + 1) + ' of ' + slips.length + '…');
        await nextTick();
      }
    }
    return items;
  }

  function doneMessage(html) {
    Art.say('celebrate', 'Done.', { settle: 'send', after: 2400 });
    setAlert($('exportMessage'), html, 'ok');
  }

  async function exportCombinedPdf() {
    var slips = exportableSlips();
    if (!slips.length) return;
    beginExport();
    try {
      var doc = PDF.createDoc(jsPDF, state.profile.paper);
      for (var i = 0; i < slips.length; i++) {
        if (state.cancelExport) throw new Error('cancelled');
        if (i) doc.addPage();
        PDF.renderPayslip(doc, renderContext(slips[i]));
        if (i % 10 === 0 || i === slips.length - 1) {
          setProgress(i + 1, slips.length, 'Making payslip ' + (i + 1) + ' of ' + slips.length + '…');
          await nextTick();
        }
      }
      PDF.setMeta(doc, state.company, baseName());
      setProgress(slips.length, slips.length, 'Building PDF file…');
      await nextTick();
      var blob = doc.output('blob');
      downloadBlob(blob, baseName() + '.pdf');
      doneMessage('<strong>Done.</strong> ' + plural(slips.length, 'payslip') + ' in one PDF for printing (' + (blob.size / 1024 / 1024).toFixed(1) + ' MB, no password).');
    } catch (e) {
      reportError($('exportMessage'), e);
    } finally {
      endExport();
    }
  }

  async function exportZip() {
    var slips = exportableSlips();
    if (!slips.length) return;
    var locked = lockOn('download');
    beginExport();
    try {
      var items = await buildAll(slips, locked);
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
      doneMessage('<strong>Done.</strong> ' + plural(items.length, 'PDF file') + ' + manifest.csv in a ZIP (' + (blob.size / 1024 / 1024).toFixed(1) + ' MB).' + protectNote(items, locked));
    } catch (e) {
      reportError($('exportMessage'), e);
    } finally {
      endExport();
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
    var locked = lockOn('download');
    beginExport();
    try {
      var sub = await dir.getDirectoryHandle(baseName(), { create: true });
      var items = [];
      for (var i = 0; i < slips.length; i++) {
        if (state.cancelExport) throw new Error('cancelled');
        var art = artifactFor(slips[i], locked);
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
      doneMessage('<strong>Done.</strong> ' + plural(items.length, 'PDF file') + ' + manifest.csv in the folder <code>' + escapeHtml(baseName()) + '</code>.' + protectNote(items, locked) + ' The delivery log is kept in the same folder from now on.');
    } catch (e) {
      reportError($('exportMessage'), e);
    } finally {
      endExport();
    }
  }

  function protectNote(items, locked) {
    if (!locked) return ' Files are not password protected.';
    var without = items.filter(function (it) { return !it.art.password.password; }).length;
    return ' Each file is password protected' + (without ? ' except ' + without + ' without the needed data' : '') + '.';
  }

  async function downloadEmlPack() {
    var slips = exportableSlips().filter(function (s) { return s.fields.email; });
    if (!slips.length) { setAlert($('exportMessage'), 'No employee has an email address in the file.', 'warn'); return; }
    beginExport();
    try {
      var items = await buildAll(slips, lockOn('email'));
      var zip = new JSZip();
      var root = zip.folder(baseName() + ' - Outlook drafts');
      items.forEach(function (it, i) {
        var m = messageFor(it.slip, 'email');
        var eml = Delivery.buildEml({ from: state.company.senderEmail, to: it.slip.fields.email, subject: m.subject, text: m.text, fileName: it.fileName, pdfBytes: it.art.bytes });
        root.folder('batch-' + ('0' + (Math.floor(i / EML_BATCH) + 1)).slice(-2)).file(it.fileName.replace(/\.pdf$/, '.eml'), eml);
      });
      var pdfs = root.folder('pdf');
      items.forEach(function (it) { pdfs.file(it.fileName, it.art.bytes); });
      pdfs.file('manifest.csv', Delivery.manifestCsv(items, state.company));
      pdfs.file('OutlookDrafts.bas', Delivery.outlookVbaModule(state.company));
      root.file('READ ME - how to send.html', Delivery.outlookGuideHtml(state.company, EML_BATCH));
      setProgress(1, 1, 'Packing ZIP…');
      await nextTick();
      var blob = await zip.generateAsync({ type: 'blob', compression: 'STORE' });
      downloadBlob(blob, baseName() + ' - Outlook drafts.zip');
      doneMessage('<strong>Done.</strong> ' + plural(items.length, 'draft') + ' in ' + Math.ceil(items.length / EML_BATCH) + ' batch folder(s), plus the PDFs, manifest and macro. Open "READ ME - how to send.html" first.');
    } catch (e) {
      reportError($('exportMessage'), e);
    } finally {
      endExport();
    }
  }

  async function downloadMergePack() {
    var all = exportableSlips();
    var slips = all.filter(function (s) { return s.fields.email; });
    if (!slips.length) { setAlert($('exportMessage'), 'No employee has an email address in the file.', 'warn'); return; }
    beginExport();
    try {
      var items = await buildAll(slips, lockOn('email'));
      items.forEach(function (it) {
        var m = messageFor(it.slip, 'email');
        it.subject = m.subject;
        it.body = m.text;
      });
      var zip = new JSZip();
      var root = zip.folder(baseName() + ' - Mail Merge');
      items.forEach(function (it) { root.file(it.fileName, it.art.bytes); });
      root.file('mailmerge.csv', Delivery.mailMergeCsv(items, state.company));
      root.file('manifest.csv', Delivery.manifestCsv(items, state.company));
      root.file('READ ME - how to send.html', Delivery.mailMergeGuideHtml(state.company, { withEmail: items.length, withoutEmail: all.length - items.length }));
      setProgress(1, 1, 'Packing ZIP…');
      await nextTick();
      var blob = await zip.generateAsync({ type: 'blob', compression: 'STORE' });
      downloadBlob(blob, baseName() + ' - Mail Merge.zip');
      doneMessage('<strong>Done.</strong> Mail Merge pack with ' + plural(items.length, 'PDF') + ', mailmerge.csv and the guide. Open "READ ME - how to send.html" first.');
    } catch (e) {
      reportError($('exportMessage'), e);
    } finally {
      endExport();
    }
  }

  function downloadRegister() {
    var slips = exportableSlips();
    var doc = PDF.createDoc(jsPDF, state.profile.paper);
    PDF.renderRegister(doc, { settings: state.company, slips: slips });
    PDF.setMeta(doc, state.company, 'Acknowledgement register');
    downloadBlob(doc.output('blob'), baseName() + ' - Acknowledgement register.pdf');
  }

  $('btnExportPdf').addEventListener('click', exportCombinedPdf);
  $('btnExportZip').addEventListener('click', exportZip);
  $('btnExportFolder').addEventListener('click', exportFolder);
  $('btnEmlPack').addEventListener('click', downloadEmlPack);
  $('btnMergePack').addEventListener('click', downloadMergePack);
  $('btnRegister').addEventListener('click', downloadRegister);
  if (!fsaSupported) $('btnExportFolder').classList.add('hidden');
  $('skipErrors').addEventListener('change', function () { renderSendPanel(); saveSessionState(); });

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
        if (!state.dirWritable) { refreshSendViews(); return; }
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
      refreshSendViews();
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
    var locked = channel === 'whatsapp' ? lockOn('whatsapp') : lockOn('email');
    var art = state.cache[slip.employeeId + '|' + slip.rowNumber + '|' + (locked ? 1 : 0)];
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
    refreshSendViews();
    var row = $('sendTable').querySelector('tr[data-emp="' + cssEscape(slip.employeeId) + '"]');
    if (row) { row.classList.remove('flash'); void row.offsetWidth; row.classList.add('flash'); }
    var all = exportableSlips();
    if (all.length && all.every(isSent)) {
      Art.say('celebrate', 'Everyone has their payslip. Download the <b>delivery log</b> under More for your records.', { settle: 'send', after: 4000 });
    }
  }

  function unmarkSent(slip) {
    addLogEntry({ action: 'undo', employeeId: slip.employeeId, name: slip.name, channel: 'manual', note: 'tick removed' });
    refreshSendViews();
  }

  function cssEscape(v) {
    return String(v).replace(/["\\]/g, '\\$&');
  }

  function sentText(slip) {
    var e = state.sent[slip.employeeId];
    if (!e) return '';
    var d = new Date(e.timestamp);
    return 'Sent ' + (CHANNEL_NAMES[e.channel] && e.channel !== 'manual' ? 'by ' + CHANNEL_NAMES[e.channel] : '(' + (CHANNEL_NAMES[e.channel] || e.channel) + ')') + ', ' +
      d.toLocaleDateString(undefined, { day: 'numeric', month: 'short' }) + ' ' + d.toLocaleTimeString(undefined, { hour: '2-digit', minute: '2-digit' });
  }

  function nextUnsent(after) {
    var slips = exportableSlips();
    var start = after ? slips.indexOf(after) + 1 : 0;
    for (var k = 0; k < slips.length; k++) {
      var s = slips[(start + k) % slips.length];
      if (!isSent(s)) return s;
    }
    return null;
  }

  // --- the three ways out: WhatsApp, email, download ---------------------------
  // Already-sent people get a question first (maybe sent before a refresh or from another tab).
  function guarded(slip, action) {
    if (!isSent(slip)) { action(); return; }
    confirmResend(slip).then(function (yes) { if (yes) action(); });
  }

  function sendWhatsApp(slip) {
    if (!slip.fields.mobile || !state.canSend) return;
    guarded(slip, function () {
      var locked = lockOn('whatsapp');
      var art = artifactFor(slip, locked);
      var m = messageFor(slip, 'whatsapp');
      // the chat opens with the right person and the message typed; the PDF lands in Downloads to attach
      window.open(Delivery.waLink(slip.fields.mobile, m.subject + '\n\n' + m.text), '_blank', 'noopener');
      downloadBlob(new Blob([art.bytes], { type: 'application/pdf' }), art.fileName);
      markSent(slip, 'whatsapp', (locked ? 'locked ' : '') + 'PDF downloaded, chat opened');
    });
  }

  function shareWhatsApp(slip) {
    if (!state.canSend) return;
    guarded(slip, function () {
      var art = artifactFor(slip, lockOn('whatsapp'));
      var m = messageFor(slip, 'whatsapp');
      var file = new File([art.bytes], art.fileName, { type: 'application/pdf' });
      navigator.share({ files: [file], title: m.subject, text: m.subject + '\n\n' + m.text }).then(function () {
        markSent(slip, 'whatsapp', 'file shared');
      }).catch(function (err) {
        if (err && err.name !== 'AbortError') Guide.toast({ kind: 'warn', html: 'Sharing did not work (' + escapeHtml(err.message) + '). Use "Send on WhatsApp" instead.' });
      });
    });
  }

  function sendEmail(slip) {
    if (!slip.fields.email || !state.canSend) return;
    guarded(slip, function () {
      var art = artifactFor(slip, lockOn('email'));
      var m = messageFor(slip, 'email');
      var eml = Delivery.buildEml({ from: state.company.senderEmail, to: slip.fields.email, subject: m.subject, text: m.text, fileName: art.fileName, pdfBytes: art.bytes });
      downloadBlob(new Blob([eml], { type: 'message/rfc822' }), art.fileName.replace(/\.pdf$/, '.eml'));
      markSent(slip, 'outlook', 'Outlook email downloaded');
    });
  }

  function downloadPdf(slip) {
    var art = artifactFor(slip, lockOn('download'));
    downloadBlob(new Blob([art.bytes], { type: 'application/pdf' }), art.fileName);
  }

  function confirmResend(slip) {
    return Guide.confirm({
      title: 'Already sent',
      html: '<b>' + escapeHtml(slip.name) + '</b>: ' + escapeHtml(sentText(slip).replace(/^Sent/, 'sent')) + '. Sending again means a second message for this person.',
      ok: 'Send again', cancel: 'Keep one', mood: 'look'
    });
  }

  // --- panel ----------------------------------------------------------------
  function renderSendPanel() {
    if (!state.parsed) return;
    refreshExportState();
    renderLockCard();
    renderSendSummary();
    renderSendNotice();
    renderSendList();
    $('msgSubject').value = state.company.messageSubject || '';
    $('msgFrom').value = state.company.senderEmail || '';
    $('msgText').value = state.company.messageText || '';
    var fixed = profileLocked();
    ['msgSubject', 'msgFrom', 'msgText'].forEach(function (id) { $(id).readOnly = fixed; });
    $('msgFixed').innerHTML = fixed ? 'The message is part of your company profile. <button type="button" class="link inline" data-profile-edit>Edit profile</button>' : '';
    renderMessagePreview();
  }
  $('panel4').addEventListener('click', function (e) {
    if (e.target.closest('[data-profile-edit]')) { e.preventDefault(); startEdit(); }
  });

  function refreshSendViews() {
    if (!state.parsed) return;
    if (state.step === 4) {
      renderSendSummary();
      renderSendList();
      renderSendNotice();
    }
    if (viewer) viewer.refresh();
  }

  function renderSendSummary() {
    var slips = exportableSlips();
    var sent = slips.filter(isSent).length;
    var mob = slips.filter(function (s) { return s.fields.mobile; }).length;
    var mail = slips.filter(function (s) { return s.fields.email; }).length;
    var none = slips.filter(function (s) { return !s.fields.mobile && !s.fields.email; }).length;
    $('sentCount').textContent = sent;
    $('sendTotal').textContent = slips.length;
    $('sendMeter').style.width = (slips.length ? Math.round(sent / slips.length * 100) : 0) + '%';
    $('sendFacts').innerHTML = [
      '<span><b>' + mob + '</b> with a mobile</span>',
      '<span><b>' + mail + '</b> with an email</span>',
      none ? '<span class="warn"><b>' + none + '</b> with neither (download or print)</span>' : '',
      state.dirHandle ? (state.dirWritable ? '<span class="ok" title="Every tick is also written to delivery-log.csv in the export folder">log saved in your folder</span>'
        : '<button type="button" class="link" id="btnReconnect" title="After a restart the browser needs your OK to write to the folder again">reconnect the folder for the log</button>') : ''
    ].filter(Boolean).join('<i class="dot-sep"></i>');
    if ($('btnReconnect')) $('btnReconnect').addEventListener('click', reconnectFolder);
    var next = nextUnsent();
    var btn = $('btnStartSending');
    btn.innerHTML = ICONS.eye + '<span>' + (!slips.length ? 'Nothing to send' : !next ? 'Everyone is done' : sent ? 'Continue with ' + escapeHtml(next.name) : 'Start: view the first payslip') + '</span>';
    btn.disabled = !next;
  }

  function renderLockCard() {
    var rule = String(state.company.passwordRule || 'id').toLowerCase();
    var none = rule === 'none';
    $('lockWhatsapp').checked = lockOn('whatsapp');
    $('lockEmail').checked = lockOn('email');
    $('lockDownload').checked = lockOn('download');
    var fixed = profileLocked();
    ['lockWhatsapp', 'lockEmail', 'lockDownload'].forEach(function (id) { $(id).disabled = none || fixed; });
    $('passwordRule').value = ['id', 'id+dob', 'dob', 'custom'].indexOf(rule) >= 0 ? rule : 'id';
    $('passwordRule').disabled = none || fixed;
    $('lockCard').classList.toggle('fixed', fixed);
    var ruleText = { 'id+dob': 'Iqama number or date of birth', id: 'Iqama / National ID number', dob: 'date of birth', custom: 'password column' }[rule] || 'Iqama / National ID number';
    var cannot = exportableSlips().filter(function (s) { return Schema.buildPassword(s, rule).error; }).length;
    var notes = [];
    if (none) notes.push('The Settings sheet says Password Rule "none", so no PDF is locked.');
    else {
      if (!lockOn('whatsapp')) notes.push('<span class="warn">WhatsApp PDFs are not locked: anyone holding the phone can open them.</span>');
      if (cannot && (lockOn('whatsapp') || lockOn('email') || lockOn('download'))) notes.push(plural(cannot, 'employee has', 'employees have') + ' no ' + ruleText + ' in the file; their PDFs stay open.');
      if (!notes.length) notes.push('Locked messages tell the employee the password.');
    }
    if (fixed) notes.push('Part of your company profile. <button type="button" class="link inline" data-profile-edit>Edit profile</button>');
    $('lockNote').innerHTML = notes.join(' ');
  }

  function setLock(channel, on) {
    state.profile.lock = Object.assign({}, LOCK_DEFAULTS, state.profile.lock || {});
    state.profile.lock[channel] = on;
    saveProfile();
    state.cache = {};
    renderLockCard();
    renderMessagePreview();
  }
  $('lockWhatsapp').addEventListener('change', function () { setLock('whatsapp', this.checked); });
  $('lockEmail').addEventListener('change', function () { setLock('email', this.checked); });
  $('lockDownload').addEventListener('change', function () { setLock('download', this.checked); });
  $('passwordRule').addEventListener('change', function () {
    setCompanyValue('passwordRule', this.value);
    state.cache = {};
    renderLockCard();
    renderMessagePreview();
  });

  function renderMessagePreview() {
    var slips = exportableSlips();
    if (!slips.length) { $('msgPreview').innerHTML = ''; return; }
    var wa = messageFor(slips[0], 'whatsapp');
    var em = messageFor(slips[0], 'email');
    $('msgPreview').innerHTML =
      '<div class="msg-card"><span class="msg-k">' + ICONS.wa + 'WhatsApp' + (lockOn('whatsapp') ? ' · PDF locked' : '') + '</span><pre>' + escapeHtml(wa.subject + '\n\n' + wa.text) + '</pre></div>' +
      '<div class="msg-card"><span class="msg-k">' + ICONS.mail + 'Email' + (lockOn('email') ? ' · PDF locked' : ' · PDF not locked') + '</span><pre><b>' + escapeHtml(em.subject) + '</b>\n\n' + escapeHtml(em.text) + '</pre></div>';
  }

  ['msgSubject', 'msgFrom', 'msgText'].forEach(function (id) {
    $(id).addEventListener('input', function () {
      var key = { msgSubject: 'messageSubject', msgFrom: 'senderEmail', msgText: 'messageText' }[id];
      setCompanyValue(key, $(id).value);
      renderMessagePreview();
    });
  });

  function sendRows() {
    var slips = exportableSlips();
    var q = $('sendSearch').value.trim().toLowerCase();
    var f = $('sendFilter').value;
    return slips.filter(function (s) {
      if (f === 'todo' && isSent(s)) return false;
      if (f === 'sent' && !isSent(s)) return false;
      if (f === 'nocontact' && (s.fields.mobile || s.fields.email)) return false;
      if (!q) return true;
      return (s.employeeId + ' ' + s.name).toLowerCase().indexOf(q) >= 0;
    });
  }

  function renderSendList() {
    var rows = sendRows();
    var LIMIT = 300;
    var tbody = $('sendTable').querySelector('tbody');
    var dis = state.canSend ? '' : ' disabled';
    tbody.innerHTML = rows.slice(0, LIMIT).map(function (s, i) {
      var sent = isSent(s);
      var mob = s.fields.mobile;
      var mail = s.fields.email;
      return '<tr class="' + (sent ? 'sent' : '') + '" data-emp="' + escapeHtml(s.employeeId) + '">' +
        '<td class="tick"><input type="checkbox" data-act="tick" data-i="' + i + '"' + (sent ? ' checked' : '') + ' title="Sent"></td>' +
        '<td class="who"><b>' + escapeHtml(s.name) + '</b><span>' + escapeHtml(s.employeeId) + (sent ? ' · <em>' + escapeHtml(sentText(s)) + '</em>' : '') + '</span></td>' +
        '<td class="contact">' + (mob ? '+' + escapeHtml(mob) : '<span class="nocontact">no mobile</span>') + '<br>' + (mail ? escapeHtml(mail) : '<span class="nocontact">no email</span>') + '</td>' +
        '<td class="acts">' +
        '<button type="button" class="btn btn-small" data-act="view" data-i="' + i + '">' + ICONS.eye + 'View</button>' +
        '<button type="button" class="btn btn-small btn-wa" data-act="wa" data-i="' + i + '"' + (mob ? dis : ' disabled') + ' title="' + (mob ? 'Send on WhatsApp' : 'No mobile number in the file') + '">' + ICONS.wa + 'WhatsApp</button>' +
        '<button type="button" class="btn btn-small" data-act="eml" data-i="' + i + '"' + (mail ? dis : ' disabled') + ' title="' + (mail ? 'Send by email (Outlook)' : 'No email address in the file') + '">' + ICONS.mail + 'Email</button>' +
        '<button type="button" class="btn btn-small btn-ghost icon-only" data-act="pdf" data-i="' + i + '" title="Download the PDF" aria-label="Download the PDF">' + ICONS.download + '</button>' +
        '</td></tr>';
    }).join('') || '<tr><td colspan="4" class="muted">Nobody matches.</td></tr>';
    tbody.__rows = rows;
    $('sendCount').textContent = rows.length === exportableSlips().length ? rows.length : rows.length + ' of ' + exportableSlips().length;
    $('sendNote').textContent = rows.length > LIMIT ? 'Showing the first ' + LIMIT + ' rows. Use the search box or the filter to see the rest.' : '';
  }

  $('sendSearch').addEventListener('input', renderSendList);
  $('sendFilter').addEventListener('change', renderSendList);
  $('btnStartSending').addEventListener('click', function () {
    var next = nextUnsent();
    if (next) openViewer(exportableSlips(), exportableSlips().indexOf(next));
  });

  $('sendTable').addEventListener('click', function (e) {
    var el = e.target.closest('[data-act]');
    if (!el) return;
    var rows = $('sendTable').querySelector('tbody').__rows || [];
    var slip = rows[+el.getAttribute('data-i')];
    if (!slip) return;
    var act = el.getAttribute('data-act');
    if (act === 'tick') { if (el.checked) markSent(slip, 'manual', 'ticked by hand'); else unmarkSent(slip); }
    else if (act === 'view') openViewer(rows, rows.indexOf(slip));
    else if (act === 'wa') sendWhatsApp(slip);
    else if (act === 'eml') sendEmail(slip);
    else if (act === 'pdf') downloadPdf(slip);
  });

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

  // --- viewer: look first, then send ----------------------------------------
  var viewer = null;
  function openViewer(list, index) {
    if (!list.length) return;
    var i = Math.max(0, index);
    var url = null;
    var justSent = false;
    var m = Guide.modal({
      className: 'viewer',
      label: 'Payslip',
      html: '<div class="vw-top"><button type="button" class="btn btn-small" data-v="prev" aria-label="Previous">←</button><div class="vw-pos"></div>' +
        '<button type="button" class="btn btn-small" data-v="next" aria-label="Next">→</button></div>' +
        '<div class="vw-grid"><div class="vw-doc"><iframe title="Payslip"></iframe><a class="vw-open" target="_blank" rel="noopener">Open in a new tab</a></div><div class="vw-side"></div></div>',
      onKey: function (e) {
        if (e.target && /INPUT|TEXTAREA|SELECT/.test(e.target.tagName)) return;
        if (e.key === 'ArrowRight') { e.preventDefault(); go(1); }
        if (e.key === 'ArrowLeft') { e.preventDefault(); go(-1); }
      },
      onClose: function () {
        if (url) URL.revokeObjectURL(url);
        viewer = null;
      }
    });
    var el = m.el;

    function go(d) {
      if (!list.length) return;
      i = (i + d + list.length) % list.length;
      justSent = false;
      show(true);
    }

    function show(loadPdf) {
      var s = list[i];
      if (loadPdf) {
        if (url) URL.revokeObjectURL(url);
        url = URL.createObjectURL(previewBlob(s));
        el.querySelector('iframe').src = url + '#toolbar=0&navpanes=0&view=FitH';
        el.querySelector('.vw-open').href = url;
      }
      el.querySelector('.vw-pos').innerHTML = '<b>' + (i + 1) + '</b> of ' + list.length;
      var sent = isSent(s);
      var cur = escapeHtml(state.company.currency || 'SAR');
      var waLocked = lockOn('whatsapp');
      var emLocked = lockOn('email');
      var issues = s.issues.filter(function (x) { return x.level !== 'info'; });
      var dis = state.canSend ? '' : ' disabled';
      el.querySelector('.vw-side').innerHTML =
        '<div class="vw-who"><b>' + escapeHtml(s.name) + '</b><span>' + escapeHtml([s.employeeId, s.fields.jobTitle, s.fields.department].filter(Boolean).join(' · ')) + '</span></div>' +
        '<div class="vw-net"><span>Net pay</span><b>' + cur + ' ' + Schema.formatMoney(s.net) + '</b></div>' +
        '<div class="vw-status ' + (sent ? 'is-sent' : '') + '">' + (sent ? '✓ ' + escapeHtml(sentText(s)) : 'Not sent yet') + '</div>' +
        (issues.length ? '<ul class="vw-issues">' + issues.slice(0, 3).map(function (x) { return '<li class="' + (x.level === 'error' ? 'err' : 'warn') + '">' + escapeHtml(x.message) + '</li>'; }).join('') + '</ul>' : '') +
        '<div class="vw-actions">' +
        '<button type="button" class="btn btn-wa btn-block" data-v="wa"' + (s.fields.mobile ? dis : ' disabled') + '>' + ICONS.wa + 'Send on WhatsApp</button>' +
        '<span class="vw-sub">' + (s.fields.mobile ? '+' + escapeHtml(s.fields.mobile) + (waLocked ? ' · ' + ICONS.lock + 'locked' : ' · not locked') : 'No mobile number in the file') + '</span>' +
        (canShareFiles && s.fields.mobile ? '<button type="button" class="link vw-share" data-v="share"' + dis + '>' + ICONS.share + 'or share the file to WhatsApp</button>' : '') +
        '<button type="button" class="btn btn-block" data-v="eml"' + (s.fields.email ? dis : ' disabled') + '>' + ICONS.mail + 'Send by email</button>' +
        '<span class="vw-sub">' + (s.fields.email ? escapeHtml(s.fields.email) + (emLocked ? ' · ' + ICONS.lock + 'locked' : ' · not locked') : 'No email address in the file') + '</span>' +
        '<button type="button" class="btn btn-ghost btn-block" data-v="pdf">' + ICONS.download + 'Download PDF</button>' +
        '</div>' +
        (justSent === 'wa' ? '<p class="vw-hint">WhatsApp is opening with the message typed. The PDF is in your Downloads: drag it into the chat, then press Send.</p>' : '') +
        (justSent === 'eml' ? '<p class="vw-hint">Open the downloaded email (bottom of the browser). Outlook shows it ready: press Send.</p>' : '') +
        (justSent ? '<button type="button" class="btn btn-primary btn-block" data-v="nextUnsent">Next payslip →</button>' : '');
    }

    el.addEventListener('click', function (e) {
      var b = e.target.closest('[data-v]');
      if (!b || b.disabled) return;
      var v = b.getAttribute('data-v');
      var s = list[i];
      if (v === 'prev') go(-1);
      else if (v === 'next') go(1);
      else if (v === 'wa') { sendWhatsApp(s); justSent = 'wa'; show(false); }
      else if (v === 'share') { shareWhatsApp(s); }
      else if (v === 'eml') { sendEmail(s); justSent = 'eml'; show(false); }
      else if (v === 'pdf') downloadPdf(s);
      else if (v === 'nextUnsent') {
        var n = nextUnsent(s);
        if (!n) { m.close(); return; }
        var k = list.indexOf(n);
        if (k < 0) { list = exportableSlips(); k = list.indexOf(n); }
        i = k;
        justSent = false;
        show(true);
      }
    });
    viewer = { refresh: function () { show(false); } };
    show(true);
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
    if (!state.profile.locked) return { state: 'empty' };
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
    var p = state.profile;
    if (!p.locked) { chip.className = 'top-chip backup-chip hidden'; return; }
    var st = backupStatus();
    var ok = st.state === 'ok' && !state.editing;
    chip.className = 'top-chip backup-chip ' + (ok ? 'state-ok' : 'state-stale') + (bump ? ' bump' : '');
    chip.innerHTML = (ok ? ICON_SHIELD : ICON_SAVE) + '<span class="lbl">' + escapeHtml(state.editing ? 'Editing profile' : (p.company.companyName || 'Company profile')) + '</span>';
    chip.title = state.editing ? 'Editing the company profile' : ok ? 'Company profile, saved as a file ' + Vault.timeAgo(st.at) + (st.how === 'auto' ? ' (kept up to date)' : '') : 'Company profile: not saved as a file yet. Click to save it.';
    chip.setAttribute('aria-label', chip.title);
    if (state.step === 1) renderStartCards();
  }

  function settingsJson() {
    return JSON.stringify(settingsFileObject(), null, 2);
  }

  function settingsFileName() {
    return PDF.safeFileName('Company profile - ' + (state.profile.company.companyName || state.company.companyName || 'my company')) + '.json';
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
      return window.showSaveFilePicker({ suggestedName: settingsFileName(), id: 'payslip-settings', types: [{ description: 'Company profile', accept: { 'application/json': ['.json'] } }] }).then(function (handle) {
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
        ? '<b>Company profile saved</b> to your Downloads. Keep it with your payroll files; load it on any computer.'
        : '<b>Company profile saved.</b> The file is kept up to date from now on.' });
      return true;
    });
  }

  $('backupChip').addEventListener('click', function () { openProfilePanel(); });

  // ---------------------------------------------------------------------------
  // Company profile: set up once (company details, design, message, passwords, columns), saved as a
  // file and loaded when needed. Every month then needs only the Excel file and the pay period, so
  // payslips cannot drift from month to month. Changing it is a deliberate edit with Save or Cancel.
  // ---------------------------------------------------------------------------
  var ICON_DOC = '<svg viewBox="0 0 24 24"><path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z"/><path d="M14 3v5h5"/><path d="M9 13h6M9 17h4"/></svg>';
  var ICON_ID = '<svg viewBox="0 0 24 24"><rect x="3" y="5" width="18" height="14" rx="2.5"/><circle cx="9" cy="11" r="2.2"/><path d="M5.8 16c.6-1.6 1.8-2.4 3.2-2.4s2.6.8 3.2 2.4M14.5 10h4M14.5 13.5h3"/></svg>';

  function applyMode() {
    var fixed = profileLocked();
    var steps = fixed ? [1, 2, 4] : [1, 2, 3, 4];
    Array.prototype.forEach.call($('stepper').children, function (li) {
      var k = steps.indexOf(+li.getAttribute('data-step'));
      li.classList.toggle('hidden', k < 0);
      if (k >= 0) li.querySelector('.dot').textContent = k + 1;
    });
    $('panel2').querySelector('.panel-nav [data-go="3"]').textContent = fixed ? 'Continue to send →' : 'Continue to design →';
    $('btnDesignNext').textContent = state.editing ? 'Save profile' : state.isSample ? 'Continue to send →' : 'Save company profile & continue →';
    $('designSub').textContent = state.editing ? 'Editing the company profile: changes apply to every payslip from now on.'
      : state.isSample ? 'Sample data: try any design. Nothing is saved as your profile.'
      : 'Set once: this becomes your company profile, so every payslip looks the same each month.';
    $('editBar').classList.toggle('hidden', !state.editing);
    document.body.classList.toggle('editing', !!state.editing);
    renderBackupChip();
    renderStartCards();
  }

  function renderStartCards() {
    var p = state.profile;
    var c = p.company || {};
    var A = $('startCardA');
    var B = $('startCardB');
    if (p.locked) {
      var t = PDF.TEMPLATES.filter(function (x) { return x.id === p.templateId; })[0];
      var b = backupStatus();
      A.innerHTML = '<span class="sc-icon">' + ICON_DOC + '</span><h3>Your Excel template</h3><p class="muted">Already filled in with ' + escapeHtml(c.companyName) + '\'s details and logo. One row per employee.</p>' +
        '<div class="actions"><button type="button" class="btn btn-primary" data-action="mytemplate">Download my template</button><button type="button" class="btn" data-action="template">Blank template</button></div>' +
        '<button type="button" class="link" data-action="tour">▶ How it works</button>';
      B.innerHTML = '<span class="sc-icon">' + ICON_ID + '</span><h3>Company profile <span class="badge ok">set up</span></h3>' +
        '<p class="muted"><b>' + escapeHtml(c.companyName) + '</b> · ' + escapeHtml(t ? t.name : '') + ' design · version ' + (p.version || 1) + '</p>' +
        '<p class="small ' + (b.state === 'ok' ? 'muted' : 'warn-text') + '">' + (b.state === 'ok' ? 'Saved as a file ' + escapeHtml(Vault.timeAgo(b.at)) + '.' : 'Not saved as a file yet: save it so it is never lost.') + '</p>' +
        '<div class="actions"><button type="button" class="btn" data-action="profile">View profile</button><button type="button" class="btn btn-ghost" data-action="loadprofile">Load another</button></div>';
    } else {
      A.innerHTML = '<span class="sc-icon">' + ICON_DOC + '</span><h3>First time?</h3><p class="muted">Fill the template: company details once, one row per employee.</p>' +
        '<div class="actions"><button type="button" class="btn btn-primary" data-action="template">Download template</button><button type="button" class="btn" data-action="example">Filled example</button></div>' +
        '<button type="button" class="link" data-action="sample">Try with sample data <span class="arrow">→</span></button><button type="button" class="link guide-link" data-action="tour">▶ How it works</button>';
      B.innerHTML = '<span class="sc-icon">' + ICON_ID + '</span><h3>Have a company profile?</h3><p class="muted">Load it once. Every month then uses the same company details and design.</p>' +
        '<div class="actions"><button type="button" class="btn" data-action="loadprofile">Load company profile</button></div>' +
        '<p class="muted small">No profile yet? Drop your Excel file: the first run sets it up.</p>';
    }
  }

  function profileMissing() {
    return Schema.SETTINGS_FIELDS.filter(function (f) { return f.key === 'companyName' && !String(state.profile.company[f.key] || state.company[f.key] || '').trim(); });
  }

  $('btnDesignNext').addEventListener('click', function () {
    if (state.editing) { saveEdit(); return; }
    if (state.isSample || profileLocked()) { goStep(4); return; }
    var missing = profileMissing();
    if (missing.length) {
      Guide.toast({ kind: 'warn', html: 'Fill in <b>' + missing.map(function (f) { return f.label; }).join(', ') + '</b> first (Review → Company details).' });
      goStep(2);
      $('secCompany').open = true;
      return;
    }
    var p = state.profile;
    p.logo = state.excelImages.logo || p.logo;
    p.stamp = state.excelImages.stamp || p.stamp;
    p.locked = true;
    p.version = Math.max(1, p.version || 0);
    p.lastPeriod = state.company.payPeriod || p.lastPeriod;
    saveProfile();
    // straight from the click, so the browser's Save dialog is allowed to open
    saveSettingsFile().then(function (ok) {
      if (!ok) Guide.toast({ id: 'notSaved', kind: 'warn', timeout: 0, html: '<b>The profile is kept in this browser only.</b> Save it as a file from the profile button at the top right so it is never lost.' });
    });
    applyMode();
    reparse();
    goStep(4);
  });

  function startEdit() {
    if (!state.parsed) {
      Guide.toast({ kind: 'info', html: 'Drop this month\'s Excel file first: the profile is edited with a live preview of a real payslip.' });
      return;
    }
    if (state.editing) { goStep(3); return; }
    state.editing = { snapshot: JSON.stringify(state.profile), returnStep: state.step === 3 ? 2 : state.step };
    applyMode();
    reparse();
    renderDesignControls();
    goStep(3);
  }

  function profileJson(p) {
    var copy = JSON.parse(JSON.stringify(p));
    delete copy.savedAt;
    delete copy.lastPeriod;
    return JSON.stringify(copy);
  }

  function saveEdit() {
    if (!state.editing) return;
    if (profileMissing().length) { Guide.toast({ kind: 'warn', html: 'The company name cannot be empty.' }); return; }
    var changed = profileJson(JSON.parse(state.editing.snapshot)) !== profileJson(state.profile);
    var back = state.editing.returnStep || 2;
    if (changed) state.profile.version = (state.profile.version || 1) + 1;
    state.editing = null;
    saveProfile();
    if (changed) saveSettingsFile();
    applyMode();
    reparse();
    if (!changed) Guide.toast({ kind: 'info', html: 'No changes; the profile stays as it was.' });
    goStep(back);
  }

  function cancelEdit() {
    if (!state.editing) return;
    var back = state.editing.returnStep || 2;
    state.profile = JSON.parse(state.editing.snapshot);
    state.editing = null;
    saveProfile();
    state.cache = {};
    renderDesignControls();
    applyMode();
    reparse();
    goStep(back);
    Guide.toast({ kind: 'info', html: 'Changes cancelled. The profile is exactly as before.' });
  }
  $('btnSaveEdit').addEventListener('click', saveEdit);
  $('btnCancelEdit').addEventListener('click', cancelEdit);

  function removeProfile() {
    state.profile = defaultProfile();
    state.editing = null;
    saveProfile();
    Vault.setBackup({});
    state.settingsHandle = null;
    Vault.del('settings-handle').catch(function () {});
    renderDesignControls();
    applyMode();
    if (state.parsed) reparse();
    goStep(state.parsed ? 2 : 1);
  }

  function openProfilePanel() {
    var p = state.profile;
    var c = p.company || {};
    var b = backupStatus();
    var t = PDF.TEMPLATES.filter(function (x) { return x.id === p.templateId; })[0];
    var rule = { 'id+dob': 'Iqama + birth day and month', id: 'Iqama / National ID', dob: 'date of birth', custom: 'password column', none: 'none' }[String(c.passwordRule || 'id').toLowerCase()] || 'Iqama / National ID';
    var row = function (k, v) { return '<div><dt>' + k + '</dt><dd>' + v + '</dd></div>'; };
    var m = Guide.modal({
      className: 'profile-panel',
      label: 'Company profile',
      html: '<h2>Company profile</h2><p class="muted">Set once and used for every payslip, so they look the same every month. Only the Excel file and the pay period change.</p>' +
        '<div class="pf-grid"><div class="pf-thumb">' + PDF.thumbSvg(p.templateId, p.accent) + '</div><dl class="pf-list">' +
        row('Company', escapeHtml(c.companyName || '—') + (c.crNumber ? ' · CR ' + escapeHtml(c.crNumber) : '')) +
        row('Design', escapeHtml((t ? t.name : '') + ' · ' + String(p.paper).toUpperCase()) + (p.logo ? ' · <img class="pf-logo" src="' + p.logo.dataUrl + '" alt="Logo">' : ' · no logo')) +
        row('Message', c.messageText && c.messageText !== currentDefault('messageText') ? 'your own text' : 'standard text') +
        row('Password', 'WhatsApp ' + (lockOn('whatsapp') ? 'locked' : 'open') + ' · email ' + (lockOn('email') ? 'locked' : 'open') + ' · downloads ' + (lockOn('download') ? 'locked' : 'open') + ' (' + rule + ')') +
        row('Columns', Object.keys(p.columnOverrides).length ? plural(Object.keys(p.columnOverrides).length, 'column choice') + ' remembered' : 'read automatically') +
        row('Version', (p.version || 1) + ' · ' + (b.state === 'ok' ? 'saved as a file ' + escapeHtml(Vault.timeAgo(b.at)) : '<span class="warn-text">not saved as a file</span>')) +
        '</dl></div>' +
        '<div class="actions pf-actions"><button type="button" class="btn btn-primary" data-p="save">Save profile file</button><button type="button" class="btn" data-p="edit">Edit profile</button>' +
        '<button type="button" class="btn" data-p="load">Load another profile</button><button type="button" class="btn" data-p="mytemplate">Download my Excel template</button>' +
        '<button type="button" class="btn btn-ghost pf-remove" data-p="remove">Remove from this computer</button></div>'
    });
    m.el.addEventListener('click', function (e) {
      var btn = e.target.closest('[data-p]');
      if (!btn) return;
      var a = btn.getAttribute('data-p');
      if (a === 'save') saveSettingsFile().then(function (ok) { if (ok) m.close(); });
      else if (a === 'edit') { m.close(); startEdit(); }
      else if (a === 'load') { m.close(); $('profileInput').click(); }
      else if (a === 'mytemplate') downloadMyTemplate();
      else if (a === 'remove') {
        m.close();
        Guide.confirm({ title: 'Remove the company profile?', danger: true, ok: 'Remove',
          html: 'The profile is removed from this browser and the next file starts the setup again. Profile files you saved are not touched; load one any time.' }).then(function (yes) { if (yes) removeProfile(); });
      }
    });
  }

  // ---------------------------------------------------------------------------
  // Session: the open file and where the user was, kept in IndexedDB so a refresh, a crash or a
  // browser restart (even mid-sending) brings everything back exactly as it was.
  // ---------------------------------------------------------------------------
  // A refresh, crash or restart within three days reopens exactly where the user was (sending can take
  // a few days); older work is dropped. Ticks are kept separately per company and pay period anyway.
  var AUTO_RESUME_MS = 3 * 864e5;

  function rememberOn() {
    return Vault.prefs().remember !== false;
  }

  function rememberFile(rec) {
    state.sourceRec = rec;
    state.sourceHash = '';
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
      bankTotal: state.bankTotal,
      exportedTo: state.exportedTo,
      sourceHash: state.sourceHash,
      monthly: state.monthly,
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
      wb = s.file.kind === 'paste' ? readPasted(s.file.text)
        : /\.(csv|txt)$/i.test(s.file.name) ? Schema.readCsv(XLSX, decodeText(s.file.bytes))
        : XLSX.read(s.file.bytes, { type: 'array', cellDates: true, cellStyles: true });
    } catch (e) {
      Vault.clearSession();
      return;
    }
    state.sourceRec = { kind: s.file.kind, name: s.file.name, bytes: s.file.bytes, text: s.file.text, hash: s.file.hash };
    state.sourceHash = s.file.hash || '';
    loadWorkbookObject(wb, s.file.name, s.file.kind === 'paste' ? null : s.file.bytes, { restore: st });
    Vault.get('session-dir').then(function (rec) {
      if (!rec || !rec.dir || rec.hash !== state.sourceHash) return;
      state.dirHandle = rec.dir;
      state.dirWritable = false;
      return rec.dir.queryPermission ? rec.dir.queryPermission({ mode: 'readwrite' }).then(function (p) { state.dirWritable = p === 'granted'; }) : null;
    }).then(function () { refreshSendViews(); }).catch(function () { /* no handle */ });
    var sent = Object.keys(state.sent).length;
    Art.say('celebrate', null, { settle: state.step === 4 ? 'send' : 'look', after: 2600 });
    Guide.toast({
      id: 'restored', kind: 'ok', icon: '↺', timeout: 9000,
      html: '<b>Welcome back, nothing was lost.</b> ' + escapeHtml(s.file.name) + ' is open again on the ' + ['', 'Upload', 'Review', 'Design', 'Send'][state.step] + ' step' + (sent ? ', ' + plural(sent, 'payslip') + ' already marked as sent' : '') + '.',
      actions: auto ? [{ label: 'Start fresh instead', onClick: discardSession }] : []
    });
  }

  function discardSession() {
    Promise.all([Vault.clearSession(), Vault.del('session-dir')]).then(function () {
      if (state.parsed) location.reload();
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
          li(b.state === 'ok' ? 'ok' : b.state === 'empty' ? '' : 'warn', '⚙️', 'Company profile', 'Company details, logo, design, message, passwords and column choices. ' +
            (b.state === 'ok' ? 'Saved as a file ' + escapeHtml(Vault.timeAgo(b.at)) + '.' : b.state === 'empty' ? 'Not set up yet.' : 'Not saved as a file yet: save one as a backup.')) +
          li('', '📄', 'Open file and progress', sess ? escapeHtml(sess.file.name) + ': kept for 3 days so a refresh, crash or restart loses nothing.' : prefs.remember ? 'Nothing open yet.' : 'Not kept (switched off below).') +
          li('', '✉️', 'Delivery ticks', state.parsed ? plural(sentCount, 'employee') + ' marked as sent for ' + escapeHtml(state.company.payPeriod || 'this period') + (state.dirHandle && state.dirWritable ? ', also written to delivery-log.csv in your export folder.' : '. Download the delivery log to keep a copy.') : 'Kept per company and pay period, even when you close the browser.') +
          li(persisted ? 'ok' : '', persisted ? '🛡️' : 'ℹ️', 'Protected storage', persisted ? 'The browser will not clear this data by itself.' : 'The browser may clear this data when the disk runs full. The settings file is your backup.') +
          '</ul><label class="check"><input type="checkbox" id="optRemember"' + (prefs.remember ? ' checked' : '') + '><span>Remember my open file and progress on this computer. Turn this off on a shared computer.</span></label>' +
          '<div class="actions"><button type="button" class="btn btn-primary" id="dpSave">Save company profile</button><button type="button" class="btn btn-danger" id="dpErase">Erase everything here</button></div>'
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
  applyMode();
  goStep(1);
  Vault.get('settings-handle').then(function (h) { if (h && h.createWritable) state.settingsHandle = h; }).catch(function () { /* none */ });
  Vault.loadSession().then(function (s) {
    if (s && rememberOn() && s.ageMs < AUTO_RESUME_MS && !state.parsed) { restoreSession(s, true); return; }
    if (s) Vault.clearSession();
    if (!state.returning && !state.profile.locked) {
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
