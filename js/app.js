/* Payslip Generator - UI. Everything runs in the browser; no network calls. */
(function () {
  'use strict';

  var $ = function (id) { return document.getElementById(id); };

  if (typeof XLSX === 'undefined' || typeof jspdf === 'undefined' || typeof JSZip === 'undefined' ||
      typeof PayslipSchema === 'undefined' || typeof PayslipPDF === 'undefined' || typeof XlsxImages === 'undefined') {
    $('libError').classList.remove('hidden');
    return;
  }

  var Schema = PayslipSchema;
  var PDF = PayslipPDF;
  var jsPDF = jspdf.jsPDF;
  var STORAGE_KEY = 'payslip-generator-profile-v2';
  var SETTINGS_FILE_VERSION = 2;

  // ---------------------------------------------------------------------------
  // State
  // ---------------------------------------------------------------------------
  function defaultProfile() {
    return {
      templateId: 'classic',
      accent: '#1f4e79',
      paper: 'a4',
      display: PDF.mergeOptions({}),
      logo: null,           // { dataUrl, width, height }
      stamp: null,
      company: {},          // last used company settings (key -> value)
      columnOverrides: {},  // normalized header -> target
      savedAt: null
    };
  }

  var state = {
    step: 1,
    wb: null,               // parsed workbook (SheetJS)
    parsed: null,           // Schema.parseWorkbook result
    fileName: '',
    company: {},            // effective company settings used for rendering
    companyFromExcel: {},   // keys that came from the Excel Settings sheet
    excelImages: { logo: null, stamp: null },
    profile: defaultProfile(),
    previewUrl: null,
    cancelExport: false,
    exporting: false
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
    ['templateId', 'accent', 'paper'].forEach(function (k) { if (p[k]) base[k] = p[k]; });
    base.display = PDF.mergeOptions(p.display || {});
    if (p.logo && p.logo.dataUrl) base.logo = p.logo;
    if (p.stamp && p.stamp.dataUrl) base.stamp = p.stamp;
    if (p.company && typeof p.company === 'object') base.company = p.company;
    if (p.columnOverrides && typeof p.columnOverrides === 'object') base.columnOverrides = p.columnOverrides;
    base.savedAt = p.savedAt || null;
    state.profile = base;
    if (!opts.silent) saveProfile();
  }

  // Draw any browser-decodable image onto a canvas and return a PNG/JPEG data URL + size.
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
        var ctx = canvas.getContext('2d');
        ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
        var dataUrl = preferJpeg ? canvas.toDataURL('image/jpeg', 0.9) : canvas.toDataURL('image/png');
        resolve({ dataUrl: dataUrl, width: canvas.width, height: canvas.height });
      };
      img.src = src;
    });
  }

  function fileToImage(file) {
    return new Promise(function (resolve, reject) {
      var reader = new FileReader();
      reader.onerror = function () { reject(new Error('Could not read the image file.')); };
      reader.onload = function () { resolve(reader.result); };
      reader.readAsDataURL(file);
    }).then(function (dataUrl) {
      return normalizeImage(dataUrl, /jpe?g/i.test(file.type));
    });
  }

  function readFileAsArrayBuffer(file) {
    return new Promise(function (resolve, reject) {
      var reader = new FileReader();
      reader.onerror = function () { reject(new Error('Could not read the file.')); };
      reader.onload = function () { resolve(reader.result); };
      reader.readAsArrayBuffer(file);
    });
  }

  function readFileAsText(file) {
    return new Promise(function (resolve, reject) {
      var reader = new FileReader();
      reader.onerror = function () { reject(new Error('Could not read the file.')); };
      reader.onload = function () { resolve(reader.result); };
      reader.readAsText(file);
    });
  }

  // ---------------------------------------------------------------------------
  // Steps
  // ---------------------------------------------------------------------------
  function goStep(n) {
    if (n > 1 && !state.parsed) return;
    state.step = n;
    for (var i = 1; i <= 4; i++) {
      $('panel' + i).classList.toggle('hidden', i !== n);
    }
    Array.prototype.forEach.call($('stepper').children, function (li) {
      var s = +li.getAttribute('data-step');
      li.classList.toggle('active', s === n);
      li.classList.toggle('done', s < n && !!state.parsed);
      li.classList.toggle('disabled', s > 1 && !state.parsed);
    });
    window.scrollTo({ top: 0, behavior: 'smooth' });
    if (n === 3) schedulePreview();
    if (n === 4) renderExportPanel();
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
  function workbookToArray(wb) {
    return XLSX.write(wb, { type: 'array', bookType: 'xlsx', cellDates: true });
  }

  var XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

  $('btnTemplate').addEventListener('click', function () {
    downloadBlob(new Blob([workbookToArray(Schema.buildTemplateWorkbook(XLSX))], { type: XLSX_MIME }), 'Payslip_Template.xlsx');
  });

  $('btnSampleXlsx').addEventListener('click', function () {
    downloadBlob(new Blob([workbookToArray(Schema.buildTemplateWorkbook(XLSX, { sampleRows: Schema.sampleRows(25) }))], { type: XLSX_MIME }), 'Sample_Payslips.xlsx');
  });

  $('btnLoadSample').addEventListener('click', function () {
    var wb = Schema.buildTemplateWorkbook(XLSX, { sampleRows: Schema.sampleRows(25) });
    loadWorkbookData(workbookToArray(wb), 'Sample data (25 employees)');
  });

  // Template pre-filled with the company details and the logo/stamp pictures.
  $('btnMyTemplate').addEventListener('click', function () {
    var wb = Schema.buildTemplateWorkbook(XLSX, { settings: state.company });
    var buf = workbookToArray(wb);
    var images = [];
    var logo = effectiveImage('logo');
    var stamp = effectiveImage('stamp');
    if (logo) images.push({ dataUrl: logo.dataUrl, width: logo.width, height: logo.height, name: 'Logo', row: 1, col: 4, maxPx: 220 });
    if (stamp) images.push({ dataUrl: stamp.dataUrl, width: stamp.width, height: stamp.height, name: 'Stamp', row: 8, col: 4, maxPx: 160 });
    XlsxImages.embedImages(JSZip, buf, Schema.SHEET_SETTINGS, images).then(function (out) {
      var name = PDF.safeFileName('Payslip_Template - ' + (state.company.companyName || 'my company')) + '.xlsx';
      downloadBlob(new Blob([out], { type: XLSX_MIME }), name);
    }).catch(function (e) {
      alert('Could not build the template: ' + e.message);
    });
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
    if (/\.json$/i.test(file.name)) {
      return loadSettingsFile(file);
    }
    readFileAsArrayBuffer(file).then(function (buf) {
      loadWorkbookData(buf, file.name);
    }).catch(function (e) {
      setAlert($('loadError'), escapeHtml(e.message), 'error');
    });
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
    $('exportMessage').innerHTML = '';
    reparse();
    // Pictures (logo / stamp) embedded in the file, read in the background.
    XlsxImages.extractImages(JSZip, arrayBuffer).then(function (images) {
      if (!images.length) return;
      var onSettings = images.filter(function (i) { return Schema.normalizeHeader(i.sheet) === Schema.normalizeHeader(Schema.SHEET_SETTINGS); });
      var pick = onSettings.length ? onSettings : images;
      var tasks = [];
      if (pick[0]) tasks.push(normalizeImage(XlsxImages.toDataUrl(pick[0]), pick[0].mime === 'image/jpeg').then(function (img) { state.excelImages.logo = img; }, function () {}));
      if (pick[1]) tasks.push(normalizeImage(XlsxImages.toDataUrl(pick[1]), pick[1].mime === 'image/jpeg').then(function (img) { state.excelImages.stamp = img; }, function () {}));
      return Promise.all(tasks).then(function () {
        renderImages();
        schedulePreview();
      });
    }).catch(function () { /* not a zip (csv/xls) or unreadable pictures: ignore */ });
    goStep(2);
  }

  // Re-run the parser with the current column overrides and refresh step 2.
  function reparse() {
    state.parsed = Schema.parseWorkbook(XLSX, state.wb, { overrides: state.profile.columnOverrides });
    mergeCompany();
    $('fileName').textContent = state.fileName + (state.parsed.sheetName ? ' · sheet "' + state.parsed.sheetName + '"' : '');
    renderReview();
    populatePreviewSelect();
  }

  // Company settings: Excel values win, blanks are filled from the saved profile.
  function mergeCompany() {
    var excel = state.parsed.settings;
    var saved = state.profile.company || {};
    var out = {};
    state.companyFromExcel = {};
    Schema.SETTINGS_FIELDS.forEach(function (f) {
      if (f.type === 'image') return;
      var v = excel[f.key];
      var fromExcel = v !== '' && v != null && !(f.key === 'currency' && !state.parsed.settingsFound);
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
    return Schema.SETTINGS_FIELDS.filter(function (f) {
      return f.required && !String(state.company[f.key] || '').trim();
    });
  }

  function effectiveImage(kind) {
    return state.excelImages[kind] || state.profile[kind] || null;
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

    // Banner: one sentence that tells the user what to do next.
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
      kind = 'ok';
      banner = '<strong>All ' + plural(st.ok, 'payslip') + ' ready.</strong> There are a few warnings worth a look, but nothing blocking.';
    } else {
      banner = '<strong>All ' + plural(st.ok, 'payslip') + ' ready.</strong> Everything checks out.';
    }
    setAlert($('statusBanner'), banner, kind);

    // Tiles
    var tiles = [
      { label: 'Employees', value: st.total, cls: '' },
      { label: 'Ready', value: st.ok, cls: st.ok ? 'ok' : '' },
      { label: 'Rows with errors', value: st.withErrors, cls: st.withErrors ? 'err' : '' },
      { label: 'Total net (' + escapeHtml(state.company.currency || 'SAR') + ')', value: Schema.formatMoney(st.totalNet), cls: '' }
    ];
    $('summaryGrid').innerHTML = tiles.map(function (t) {
      return '<div class="tile ' + t.cls + '"><div class="tile-label">' + t.label + '</div><div class="tile-value">' + t.value + '</div></div>';
    }).join('');

    renderCompanyForm();
    renderColumns();

    // Messages
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

  function renderCompanyForm() {
    var missing = companyMissing();
    var fromExcelCount = Object.keys(state.companyFromExcel).length;
    $('companyMeta').innerHTML = missing.length
      ? '<span class="badge err">' + plural(missing.length, 'required field') + ' missing</span>'
      : '<span class="badge ok">complete</span> <span>' + escapeHtml(state.company.companyName || '') + ' · ' + escapeHtml(state.company.payPeriod || '') + '</span>';
    $('companySource').textContent = state.parsed.settingsFound
      ? 'Filled from the Settings sheet of the file' + (fromExcelCount < 3 ? ' (partly)' : '') + '. You can correct anything here; edits are kept for next time.'
      : 'This file has no Settings sheet. Enter the details here once; they are saved on this computer and in your settings file.';
    $('secCompany').open = missing.length > 0;

    var form = $('companyForm');
    form.innerHTML = Schema.SETTINGS_FIELDS.filter(function (f) { return f.type !== 'image'; }).map(function (f) {
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
        state.company[k] = inp.value;
        if (k === 'currency') state.company[k] = inp.value.toUpperCase();
        state.profile.company[k] = state.company[k];
        inp.classList.toggle('missing', !!(Schema.SETTINGS_FIELDS.filter(function (f) { return f.key === k; })[0].required) && !inp.value.trim());
        saveProfile();
        var missingNow = companyMissing();
        $('companyMeta').innerHTML = missingNow.length
          ? '<span class="badge err">' + plural(missingNow.length, 'required field') + ' missing</span>'
          : '<span class="badge ok">complete</span> <span>' + escapeHtml(state.company.companyName || '') + ' · ' + escapeHtml(state.company.payPeriod || '') + '</span>';
        refreshExportState();
        schedulePreview();
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

    var meta = '';
    if (missingReq.length) meta += '<span class="badge err">' + plural(missingReq.length, 'required column') + ' missing</span>';
    if (ignored.length) meta += '<span class="badge warn">' + plural(ignored.length, 'column') + ' ignored</span>';
    if (guessed.length) meta += '<span class="badge info">' + plural(guessed.length, 'guess', 'guesses') + '</span>';
    if (!meta) meta = '<span class="badge ok">all ' + cols.length + ' recognised</span>';
    $('columnsMeta').innerHTML = meta;
    $('secColumns').open = missingReq.length > 0 || ignored.length > 0 || guessed.length > 0;

    var mm = '';
    if (missingReq.length) mm += '<div class="alert alert-error">Missing: <strong>' + missingReq.map(escapeHtml).join('</strong>, <strong>') + '</strong>. Pick the matching column in the "Understood as" list below, or add the column in Excel.</div>';
    if (optionalMissing.length) mm += '<div class="alert alert-info">Not found (optional): ' + optionalMissing.map(escapeHtml).join(', ') + '. ' + (optionalMissing.indexOf('Employee ID') >= 0 ? 'IDs will be generated from the row order. ' : '') + '</div>';
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
        var norm = sel.getAttribute('data-norm');
        var header = sel.getAttribute('data-header');
        state.profile.columnOverrides[norm] = Schema.choiceToTarget(sel.value, header);
        saveProfile();
        var wasOpen = $('secColumns').open;
        reparse();
        $('secColumns').open = wasOpen;
        schedulePreview();
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
    var shown = rows.slice(0, LIMIT);
    $('employeeTable').querySelector('tbody').innerHTML = shown.map(function (s) {
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
        '<div class="thumb ' + t.id + '"><div class="t-head"><span></span><span></span></div><div class="t-box"></div><div class="t-tables"><span></span><span></span></div><div class="t-net"></div></div>' +
        '<div class="name">' + escapeHtml(t.name) + '</div><div class="desc">' + escapeHtml(t.description) + '</div></button>';
    }).join('');
    Array.prototype.forEach.call($('templateCards').querySelectorAll('.template-card'), function (btn) {
      btn.addEventListener('click', function () {
        state.profile.templateId = btn.getAttribute('data-id');
        renderTemplateCards();
        saveProfile();
        schedulePreview();
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

  $('accent').addEventListener('input', function () {
    state.profile.accent = $('accent').value;
    document.documentElement.style.setProperty('--accent', state.profile.accent);
    saveProfile();
    schedulePreview();
  });
  $('paper').addEventListener('change', function () {
    state.profile.paper = $('paper').value;
    saveProfile();
    schedulePreview();
  });
  $('docTitle').addEventListener('input', function () {
    state.profile.display.title = $('docTitle').value;
    saveProfile();
    schedulePreview();
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
        schedulePreview();
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
      var el = $(kind + 'Preview');
      el.innerHTML = img ? '<img src="' + img.dataUrl + '" alt="' + kind + '">' : (kind === 'logo' ? 'No logo' : 'No stamp / signature');
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
        state.excelImages[kind] = null; // the user's choice wins over the Excel picture
        renderImages();
        saveProfile();
        schedulePreview();
      }).catch(function (err) { alert(err.message); });
    });
    $(kind + 'Clear').addEventListener('click', function () {
      state.profile[kind] = null;
      state.excelImages[kind] = null;
      renderImages();
      saveProfile();
      schedulePreview();
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
      logo: effectiveImage('logo'),
      stamp: effectiveImage('stamp'),
      columnOverrides: p.columnOverrides
    };
  }

  $('btnSaveSettings').addEventListener('click', function () {
    var obj = settingsFileObject();
    var name = PDF.safeFileName('payslip-settings - ' + (obj.company.companyName || 'my company')) + '.json';
    downloadBlob(new Blob([JSON.stringify(obj, null, 2)], { type: 'application/json' }), name);
  });

  function loadSettingsFile(file) {
    return readFileAsText(file).then(function (text) {
      var obj = JSON.parse(text);
      if (!obj || obj.app !== 'payslip-generator') throw new Error('This is not a Payslip Generator settings file.');
      applyProfile(obj);
      state.excelImages = { logo: null, stamp: null };
      if (state.parsed) {
        reparse();
      }
      renderDesignControls();
      updateSavedHint();
      var msg = 'Settings loaded from ' + escapeHtml(file.name) + (obj.company && obj.company.companyName ? ' (' + escapeHtml(obj.company.companyName) + ')' : '') + '.';
      setAlert($('loadError'), msg + (state.parsed ? '' : ' Now upload the Excel file.'), 'ok');
      schedulePreview();
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
  // Preview + export
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
      var doc = PDF.createDoc(jsPDF, state.profile.paper);
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
    refreshExportState();
  }

  function refreshExportState() {
    var missing = state.parsed ? companyMissing().filter(function (f) { return f.key !== 'crNumber'; }) : [];
    var slips = exportableSlips();
    var can = !!(state.parsed && slips.length && !missing.length && !state.exporting);
    $('btnExportPdf').disabled = !can;
    $('btnExportZip').disabled = !can;
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

  $('skipErrors').addEventListener('change', renderExportPanel);

  function setProgress(done, total, label) {
    $('progress').classList.remove('hidden');
    $('progressFill').style.width = (total ? Math.round((done / total) * 100) : 0) + '%';
    $('progressText').textContent = label || (done + ' / ' + total);
  }

  function beginExport() {
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
  }

  $('btnCancel').addEventListener('click', function () { state.cancelExport = true; });

  function baseName() {
    return PDF.safeFileName(['Payslips', state.company.companyName, state.company.payPeriod].filter(Boolean).join(' - '));
  }

  function slipFileName(slip) {
    return PDF.safeFileName([slip.employeeId, slip.name, state.company.payPeriod].filter(Boolean).join(' - ')) + '.pdf';
  }

  function reportExportError(e) {
    var cancelled = e && e.message === 'cancelled';
    setAlert($('exportMessage'), cancelled ? 'Export cancelled.' : 'Export failed: ' + escapeHtml(e.message), cancelled ? 'warn' : 'error');
  }

  async function exportCombinedPdf() {
    var slips = exportableSlips();
    if (!slips.length) return;
    beginExport();
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
      setAlert($('exportMessage'), '<strong>Done.</strong> ' + plural(slips.length, 'payslip') + ' in one PDF (' + (blob.size / 1024 / 1024).toFixed(1) + ' MB, ' + ((Date.now() - started) / 1000).toFixed(1) + ' s). Check your Downloads folder.', 'ok');
    } catch (e) {
      reportExportError(e);
    } finally {
      endExport();
    }
  }

  async function exportZip() {
    var slips = exportableSlips();
    if (!slips.length) return;
    beginExport();
    var started = Date.now();
    try {
      var zip = new JSZip();
      var folder = zip.folder(baseName());
      var usedNames = {};
      for (var i = 0; i < slips.length; i++) {
        if (state.cancelExport) throw new Error('cancelled');
        var doc = PDF.createDoc(jsPDF, state.profile.paper);
        PDF.renderPayslip(doc, renderContext(slips[i]));
        PDF.setMeta(doc, state.company, 'Payslip ' + slips[i].employeeId);
        var name = slipFileName(slips[i]);
        if (usedNames[name]) name = name.replace(/\.pdf$/, ' (' + (++usedNames[name]) + ').pdf');
        else usedNames[name] = 1;
        folder.file(name, doc.output('arraybuffer'));
        if (i % 5 === 0 || i === slips.length - 1) {
          setProgress(i + 1, slips.length, 'Rendering payslip ' + (i + 1) + ' of ' + slips.length + '…');
          await nextTick();
        }
      }
      var blob = await zip.generateAsync({ type: 'blob', compression: 'DEFLATE', compressionOptions: { level: 6 } }, function (meta) {
        setProgress(Math.round(meta.percent), 100, 'Compressing ZIP… ' + Math.round(meta.percent) + '%');
      });
      downloadBlob(blob, baseName() + '.zip');
      setAlert($('exportMessage'), '<strong>Done.</strong> ' + plural(slips.length, 'PDF file') + ' in a ZIP (' + (blob.size / 1024 / 1024).toFixed(1) + ' MB, ' + ((Date.now() - started) / 1000).toFixed(1) + ' s). Check your Downloads folder.', 'ok');
    } catch (e) {
      reportExportError(e);
    } finally {
      endExport();
    }
  }

  $('btnExportPdf').addEventListener('click', exportCombinedPdf);
  $('btnExportZip').addEventListener('click', exportZip);

  // ---------------------------------------------------------------------------
  // Init
  // ---------------------------------------------------------------------------
  loadProfile();
  renderDesignControls();
  bindImage('logo');
  bindImage('stamp');
  updateSavedHint();
  goStep(1);

  // Expose for tests / debugging
  window.PayslipApp = {
    state: state,
    loadWorkbookData: loadWorkbookData,
    exportableSlips: exportableSlips,
    goStep: goStep,
    applyProfile: applyProfile
  };
})();
