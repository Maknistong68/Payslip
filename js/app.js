/* Payslip Generator - UI glue. Everything runs in the browser; no network calls. */
(function () {
  'use strict';

  var $ = function (id) { return document.getElementById(id); };

  if (typeof XLSX === 'undefined' || typeof jspdf === 'undefined' || typeof JSZip === 'undefined' || typeof PayslipSchema === 'undefined' || typeof PayslipPDF === 'undefined') {
    $('libError').classList.remove('hidden');
    return;
  }

  var Schema = PayslipSchema;
  var PDF = PayslipPDF;
  var jsPDF = jspdf.jsPDF;
  var STORAGE_KEY = 'payslip-generator-branding-v1';

  // ---------------------------------------------------------------------------
  // State
  // ---------------------------------------------------------------------------
  var state = {
    parsed: null, // result of Schema.parseWorkbook
    fileName: '',
    branding: {
      templateId: 'classic',
      accent: '#1f4e79',
      paper: 'a4',
      amountInWords: true,
      signatureLines: true,
      showBank: true,
      showAttendance: true,
      logo: null, // {dataUrl, width, height}
      stamp: null
    },
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

  function saveBranding() {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(state.branding));
    } catch (e) { /* storage may be unavailable or full (huge logo) */ }
  }

  function loadBranding() {
    try {
      var raw = localStorage.getItem(STORAGE_KEY);
      if (!raw) return;
      var b = JSON.parse(raw);
      Object.keys(state.branding).forEach(function (k) {
        if (b[k] !== undefined) state.branding[k] = b[k];
      });
    } catch (e) { /* ignore */ }
  }

  function setMessage(el, html, kind) {
    el.innerHTML = html ? '<div class="alert alert-' + (kind || 'ok') + '">' + html + '</div>' : '';
  }

  // Draw any browser-decodable image onto a canvas and return a PNG/JPEG data URL + size.
  function fileToImage(file) {
    return new Promise(function (resolve, reject) {
      var reader = new FileReader();
      reader.onerror = function () { reject(new Error('Could not read the image file.')); };
      reader.onload = function () {
        var img = new Image();
        img.onerror = function () { reject(new Error('This file is not an image the browser can open.')); };
        img.onload = function () {
          var MAX = 900;
          var scale = Math.min(1, MAX / Math.max(img.naturalWidth, img.naturalHeight));
          var canvas = document.createElement('canvas');
          canvas.width = Math.max(1, Math.round(img.naturalWidth * scale));
          canvas.height = Math.max(1, Math.round(img.naturalHeight * scale));
          var ctx = canvas.getContext('2d');
          ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
          var isJpeg = /jpe?g/i.test(file.type);
          var dataUrl = isJpeg ? canvas.toDataURL('image/jpeg', 0.9) : canvas.toDataURL('image/png');
          resolve({ dataUrl: dataUrl, width: canvas.width, height: canvas.height });
        };
        img.src = reader.result;
      };
      reader.readAsDataURL(file);
    });
  }

  // ---------------------------------------------------------------------------
  // Step 1: template downloads
  // ---------------------------------------------------------------------------
  function workbookToBlob(wb) {
    var out = XLSX.write(wb, { type: 'array', bookType: 'xlsx', cellDates: true });
    return new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  }

  $('btnTemplate').addEventListener('click', function () {
    downloadBlob(workbookToBlob(Schema.buildTemplateWorkbook(XLSX)), 'Payslip_Template.xlsx');
  });

  $('btnSampleXlsx').addEventListener('click', function () {
    downloadBlob(workbookToBlob(Schema.buildTemplateWorkbook(XLSX, { sampleRows: Schema.sampleRows(25) })), 'Sample_Payslips.xlsx');
  });

  $('btnLoadSample').addEventListener('click', function () {
    var wb = Schema.buildTemplateWorkbook(XLSX, { sampleRows: Schema.sampleRows(25) });
    // round-trip through the writer so the sample goes through exactly the same path as a real file
    var buf = XLSX.write(wb, { type: 'array', bookType: 'xlsx', cellDates: true });
    handleWorkbookData(buf, 'Sample data (25 employees)');
    $('step2').scrollIntoView({ behavior: 'smooth', block: 'start' });
  });

  // ---------------------------------------------------------------------------
  // Step 2: upload + parse
  // ---------------------------------------------------------------------------
  var dropzone = $('dropzone');
  var fileInput = $('fileInput');

  ['dragenter', 'dragover'].forEach(function (ev) {
    dropzone.addEventListener(ev, function (e) {
      e.preventDefault();
      dropzone.classList.add('drag');
    });
  });
  ['dragleave', 'drop'].forEach(function (ev) {
    dropzone.addEventListener(ev, function (e) {
      e.preventDefault();
      dropzone.classList.remove('drag');
    });
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
    var reader = new FileReader();
    reader.onerror = function () {
      setMessage($('fileIssues'), 'Could not read the file.', 'error');
    };
    reader.onload = function () {
      handleWorkbookData(reader.result, file.name);
    };
    reader.readAsArrayBuffer(file);
  }

  function handleWorkbookData(arrayBuffer, name) {
    var wb;
    try {
      wb = XLSX.read(arrayBuffer, { type: 'array', cellDates: true });
    } catch (e) {
      state.parsed = null;
      $('results').classList.remove('hidden');
      $('summaryGrid').innerHTML = '';
      setMessage($('fileIssues'), 'This file could not be opened as a spreadsheet. ' + escapeHtml(e.message || ''), 'error');
      refreshExportState();
      return;
    }
    var parsed = Schema.parseWorkbook(XLSX, wb);
    state.parsed = parsed;
    state.fileName = name;
    $('exportMessage').innerHTML = '';
    $('fileName').textContent = 'Loaded: ' + name;
    renderResults();
    populatePreviewSelect();
    refreshExportState();
    schedulePreview();
  }

  function renderResults() {
    var p = state.parsed;
    $('results').classList.remove('hidden');
    var st = p.stats;

    // Summary tiles
    var tiles = [
      { label: 'Employees', value: st.total, cls: '' },
      { label: 'Ready to print', value: st.ok, cls: st.ok ? 'ok' : '' },
      { label: 'Rows with errors', value: st.withErrors, cls: st.withErrors ? 'err' : 'ok' },
      { label: 'Rows with warnings', value: st.withWarnings, cls: st.withWarnings ? 'warn' : '' },
      { label: 'Total net (' + escapeHtml(p.settings.currency) + ')', value: Schema.formatMoney(st.totalNet), cls: '' }
    ];
    $('summaryGrid').innerHTML = tiles.map(function (t) {
      return '<div class="tile ' + t.cls + '"><div class="tile-label">' + t.label + '</div><div class="tile-value">' + t.value + '</div></div>';
    }).join('');

    // File-level issues
    var errs = p.issues.filter(function (i) { return i.level === 'error'; });
    var warns = p.issues.filter(function (i) { return i.level === 'warning'; });
    var html = '';
    if (errs.length) {
      html += '<div class="alert alert-error"><strong>Fix these in the Excel file and upload again:</strong><ul>' +
        errs.map(function (i) { return '<li>' + escapeHtml(i.message) + '</li>'; }).join('') + '</ul></div>';
    }
    if (warns.length) {
      html += '<div class="alert alert-warn"><strong>Warnings:</strong><ul>' +
        warns.map(function (i) { return '<li>' + escapeHtml(i.message) + '</li>'; }).join('') + '</ul></div>';
    }
    if (!errs.length && !st.withErrors) {
      html += '<div class="alert alert-ok">File looks good. ' + st.ok + ' payslip' + (st.ok === 1 ? '' : 's') + ' ready to export.</div>';
    } else if (!errs.length && st.withErrors) {
      html += '<div class="alert alert-warn">' + st.withErrors + ' row' + (st.withErrors === 1 ? ' has' : 's have') + ' errors and will be skipped unless you untick "Skip rows with errors". See the table below.</div>';
    }
    $('fileIssues').innerHTML = html;

    // Settings
    $('settingsList').innerHTML = Schema.SETTINGS_FIELDS.map(function (f) {
      var v = p.settings[f.key];
      var missing = v === '' || v == null;
      return '<div><dt>' + escapeHtml(f.label) + '</dt><dd class="' + (missing && f.required ? 'missing' : '') + '">' +
        (missing ? (f.required ? 'missing' : '—') : escapeHtml(v)) + '</dd></div>';
    }).join('');

    renderTable();
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
      var cls = s.hasErrors ? 'row-err' : s.issues.length ? 'row-warn' : 'row-ok';
      var badge = s.hasErrors ? '<span class="status-badge err">Error</span>' : s.issues.length ? '<span class="status-badge warn">Check</span>' : '<span class="status-badge ok">OK</span>';
      var checks = s.issues.map(function (i) {
        return '<div class="issue-' + (i.level === 'error' ? 'err' : 'warn') + '">' + escapeHtml(i.message) + '</div>';
      }).join('');
      return '<tr class="' + cls + '">' +
        '<td>' + s.rowNumber + '</td>' +
        '<td class="status">' + badge + '</td>' +
        '<td>' + escapeHtml(s.employeeId) + '</td>' +
        '<td>' + escapeHtml(s.name) + '</td>' +
        '<td>' + escapeHtml(s.fields.jobTitle) + '</td>' +
        '<td class="num">' + Schema.formatMoney(s.gross) + '</td>' +
        '<td class="num">' + Schema.formatMoney(s.totalDeductions) + '</td>' +
        '<td class="num"><strong>' + Schema.formatMoney(s.net) + '</strong></td>' +
        '<td class="checks">' + (checks || '<span class="muted">—</span>') + '</td>' +
        '</tr>';
    }).join('');
    $('rowCount').textContent = rows.length === p.slips.length ? p.slips.length : rows.length + ' of ' + p.slips.length;
    $('tableNote').textContent = rows.length > LIMIT ? 'Showing the first ' + LIMIT + ' of ' + rows.length + ' rows. Use the search box to find a specific employee. All rows are still exported.' : '';
  }

  $('searchBox').addEventListener('input', renderTable);
  $('onlyProblems').addEventListener('change', renderTable);

  // ---------------------------------------------------------------------------
  // Step 3: branding
  // ---------------------------------------------------------------------------
  function renderTemplateCards() {
    $('templateCards').innerHTML = PDF.TEMPLATES.map(function (t) {
      return '<button type="button" class="template-card' + (t.id === state.branding.templateId ? ' active' : '') + '" data-id="' + t.id + '">' +
        '<div class="thumb ' + t.id + '"><div class="t-head"><span></span><span></span></div><div class="t-box"></div><div class="t-tables"><span></span><span></span></div><div class="t-net"></div></div>' +
        '<div class="name">' + escapeHtml(t.name) + '</div><div class="desc">' + escapeHtml(t.description) + '</div></button>';
    }).join('');
    Array.prototype.forEach.call($('templateCards').querySelectorAll('.template-card'), function (btn) {
      btn.addEventListener('click', function () {
        state.branding.templateId = btn.getAttribute('data-id');
        renderTemplateCards();
        saveBranding();
        schedulePreview();
      });
    });
    document.documentElement.style.setProperty('--accent', state.branding.accent);
  }

  function bindOption(id, key, isCheckbox) {
    var el = $(id);
    if (isCheckbox) el.checked = !!state.branding[key];
    else el.value = state.branding[key];
    el.addEventListener(isCheckbox ? 'change' : 'input', function () {
      state.branding[key] = isCheckbox ? el.checked : el.value;
      if (key === 'accent') document.documentElement.style.setProperty('--accent', el.value);
      saveBranding();
      schedulePreview();
    });
  }

  function renderImagePreview(kind) {
    var el = $(kind + 'Preview');
    var img = state.branding[kind];
    el.innerHTML = img ? '<img src="' + img.dataUrl + '" alt="' + kind + '">' : '<span class="muted">No ' + (kind === 'logo' ? 'logo' : 'stamp / signature') + '</span>';
  }

  function bindImage(kind) {
    $(kind + 'Input').addEventListener('change', function (e) {
      var f = e.target.files && e.target.files[0];
      e.target.value = '';
      if (!f) return;
      fileToImage(f).then(function (img) {
        state.branding[kind] = img;
        renderImagePreview(kind);
        saveBranding();
        schedulePreview();
      }).catch(function (err) {
        alert(err.message);
      });
    });
    $(kind + 'Clear').addEventListener('click', function () {
      state.branding[kind] = null;
      renderImagePreview(kind);
      saveBranding();
      schedulePreview();
    });
  }

  // ---------------------------------------------------------------------------
  // Step 4: preview + export
  // ---------------------------------------------------------------------------
  function exportableSlips() {
    if (!state.parsed) return [];
    var skip = $('skipErrors').checked;
    return state.parsed.slips.filter(function (s) { return !skip || !s.hasErrors; });
  }

  function renderContext(slip) {
    var b = state.branding;
    return {
      slip: slip,
      settings: state.parsed.settings,
      templateId: b.templateId,
      accent: b.accent,
      logo: b.logo,
      stamp: b.stamp,
      options: { amountInWords: b.amountInWords, signatureLines: b.signatureLines, showBank: b.showBank, showAttendance: b.showAttendance }
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
    previewTimer = setTimeout(renderPreview, 200);
  }

  function renderPreview() {
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
      var doc = PDF.createDoc(jsPDF, state.branding.paper);
      PDF.renderPayslip(doc, renderContext(slip));
      PDF.setMeta(doc, state.parsed.settings, 'Payslip preview');
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

  function refreshExportState() {
    var can = !!(state.parsed && exportableSlips().length && !state.exporting);
    $('btnExportPdf').disabled = !can;
    $('btnExportZip').disabled = !can;
    var msgEl = $('exportState');
    if (!state.parsed) {
      msgEl.innerHTML = '';
      return;
    }
    var fileErrors = state.parsed.stats.fileErrors;
    if (fileErrors) {
      setMessage(msgEl, 'The file has ' + fileErrors + ' problem' + (fileErrors === 1 ? '' : 's') + ' listed in step 2 (missing company settings or columns). Payslips can still be exported, but please fix them for a complete payslip.', 'warn');
    } else if (!exportableSlips().length) {
      setMessage(msgEl, 'No rows can be exported. Fix the errors or untick "Skip rows with errors".', 'warn');
    } else {
      msgEl.innerHTML = '';
    }
  }

  $('skipErrors').addEventListener('change', refreshExportState);

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

  $('btnCancel').addEventListener('click', function () {
    state.cancelExport = true;
  });

  function baseName() {
    var s = state.parsed.settings;
    return PDF.safeFileName(['Payslips', s.companyName, s.payPeriod].filter(Boolean).join(' - '));
  }

  function slipFileName(slip) {
    var s = state.parsed.settings;
    return PDF.safeFileName([slip.employeeId, slip.name, s.payPeriod].filter(Boolean).join(' - ')) + '.pdf';
  }

  async function exportCombinedPdf() {
    var slips = exportableSlips();
    if (!slips.length) return;
    beginExport();
    var started = Date.now();
    try {
      var doc = PDF.createDoc(jsPDF, state.branding.paper);
      for (var i = 0; i < slips.length; i++) {
        if (state.cancelExport) throw new Error('cancelled');
        if (i) doc.addPage();
        PDF.renderPayslip(doc, renderContext(slips[i]));
        if (i % 10 === 0 || i === slips.length - 1) {
          setProgress(i + 1, slips.length, 'Rendering payslip ' + (i + 1) + ' of ' + slips.length + '…');
          await nextTick();
        }
      }
      PDF.setMeta(doc, state.parsed.settings, baseName());
      setProgress(slips.length, slips.length, 'Building PDF file…');
      await nextTick();
      var blob = doc.output('blob');
      downloadBlob(blob, baseName() + '.pdf');
      setMessage($('exportMessage'), 'Done: ' + slips.length + ' payslips in one PDF (' + (blob.size / 1024 / 1024).toFixed(1) + ' MB, ' + ((Date.now() - started) / 1000).toFixed(1) + ' s). Check your Downloads folder.', 'ok');
    } catch (e) {
      setMessage($('exportMessage'), e.message === 'cancelled' ? 'Export cancelled.' : 'Export failed: ' + escapeHtml(e.message), e.message === 'cancelled' ? 'warn' : 'error');
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
        var doc = PDF.createDoc(jsPDF, state.branding.paper);
        PDF.renderPayslip(doc, renderContext(slips[i]));
        PDF.setMeta(doc, state.parsed.settings, 'Payslip ' + slips[i].employeeId);
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
      setMessage($('exportMessage'), 'Done: ' + slips.length + ' PDF files in a ZIP (' + (blob.size / 1024 / 1024).toFixed(1) + ' MB, ' + ((Date.now() - started) / 1000).toFixed(1) + ' s). Check your Downloads folder.', 'ok');
    } catch (e) {
      setMessage($('exportMessage'), e.message === 'cancelled' ? 'Export cancelled.' : 'Export failed: ' + escapeHtml(e.message), e.message === 'cancelled' ? 'warn' : 'error');
    } finally {
      endExport();
    }
  }

  $('btnExportPdf').addEventListener('click', exportCombinedPdf);
  $('btnExportZip').addEventListener('click', exportZip);

  // ---------------------------------------------------------------------------
  // Init
  // ---------------------------------------------------------------------------
  loadBranding();
  renderTemplateCards();
  bindOption('accent', 'accent', false);
  bindOption('paper', 'paper', false);
  bindOption('optWords', 'amountInWords', true);
  bindOption('optSign', 'signatureLines', true);
  bindOption('optBank', 'showBank', true);
  bindOption('optAttendance', 'showAttendance', true);
  bindImage('logo');
  bindImage('stamp');
  renderImagePreview('logo');
  renderImagePreview('stamp');
  refreshExportState();

  // Expose for tests / debugging
  window.PayslipApp = { state: state, handleWorkbookData: handleWorkbookData, exportableSlips: exportableSlips };
})();
