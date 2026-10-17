/*
 * Payslip Generator - animated guidance: the first-visit quick start (download the template, fill it,
 * drop it, save the settings file), the "save point" before exporting, confirm dialogs and toasts.
 * Scenes are inline SVG animated with CSS, so they scale with the card and work offline. With reduced
 * motion every scene shows its finished state.
 */
(function (root) {
  'use strict';

  var Art = root.PayslipArt || { mascot: function () { return ''; }, confetti: function () {} };

  function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }

  // ---------------------------------------------------------------------------
  // Modal shell
  // ---------------------------------------------------------------------------
  var openCount = 0;

  function modal(opts) {
    opts = opts || {};
    var lastFocus = document.activeElement;
    var back = document.createElement('div');
    back.className = 'modal-backdrop';
    back.innerHTML = '<div class="modal ' + (opts.className || '') + '" role="dialog" aria-modal="true"' + (opts.label ? ' aria-label="' + esc(opts.label) + '"' : '') + '>' +
      (opts.dismissible === false ? '' : '<button type="button" class="modal-x" aria-label="Close">×</button>') +
      opts.html + '</div>';
    document.body.appendChild(back);
    openCount++;
    document.body.classList.add('modal-open');
    var box = back.querySelector('.modal');
    var closed = false;

    function close(result) {
      if (closed) return;
      closed = true;
      document.removeEventListener('keydown', onKey, true);
      back.classList.add('leaving');
      setTimeout(function () {
        back.remove();
        openCount--;
        if (!openCount) document.body.classList.remove('modal-open');
        if (lastFocus && lastFocus.focus) try { lastFocus.focus(); } catch (e) { /* gone */ }
      }, 260);
      if (opts.onClose) opts.onClose(result);
    }

    function onKey(e) {
      var all = document.querySelectorAll('.modal-backdrop:not(.leaving)');
      if (all[all.length - 1] !== back) return; // only the dialog on top listens
      if (e.key === 'Escape' && opts.dismissible !== false) { e.preventDefault(); close('dismiss'); return; }
      if (opts.onKey) opts.onKey(e);
      if (e.key === 'Tab') {
        // keep focus inside the dialog
        var f = box.querySelectorAll('button:not([disabled]), a[href], input, select, textarea, [tabindex]:not([tabindex="-1"])');
        if (!f.length) return;
        var first = f[0], last = f[f.length - 1];
        if (e.shiftKey && document.activeElement === first) { e.preventDefault(); last.focus(); }
        else if (!e.shiftKey && document.activeElement === last) { e.preventDefault(); first.focus(); }
      }
    }
    document.addEventListener('keydown', onKey, true);
    back.addEventListener('mousedown', function (e) { if (e.target === back && opts.dismissible !== false) close('dismiss'); });
    var x = box.querySelector('.modal-x');
    if (x) x.addEventListener('click', function () { close('dismiss'); });
    setTimeout(function () {
      var target = box.querySelector('[data-autofocus]') || box.querySelector('.btn-primary') || box;
      if (target === box) box.setAttribute('tabindex', '-1');
      try { target.focus({ preventScroll: true }); } catch (e) { target.focus(); }
    }, 60);
    return { el: box, close: close };
  }

  // ---------------------------------------------------------------------------
  // Scenes (viewBox 360 x 240)
  // ---------------------------------------------------------------------------
  var INK = '#0f172a';
  function defs(id) {
    return '<defs>' +
      '<linearGradient id="' + id + 'b" x1="0" y1="0" x2="1" y2="0"><stop offset="0" stop-color="#1e3a8a"/><stop offset=".55" stop-color="#2563eb"/><stop offset="1" stop-color="#0ea5e9"/></linearGradient>' +
      '<linearGradient id="' + id + 'a" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#fcd34d"/><stop offset="1" stop-color="#f59e0b"/></linearGradient>' +
      '<linearGradient id="' + id + 'g" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#34d399"/><stop offset="1" stop-color="#059669"/></linearGradient>' +
      '</defs>';
  }

  function fileIcon(cls, label, color, x, y) {
    return '<g class="' + cls + '"' + (x != null ? ' transform="translate(' + x + ' ' + y + ')"' : '') + '><g class="sc-file-in">' +
      '<path d="M0 0h24l11 11v35H0z" fill="#fff" stroke="' + INK + '" stroke-width="2.5" stroke-linejoin="round"/>' +
      '<path d="M24 0v11h11" fill="#e2e8f0" stroke="' + INK + '" stroke-width="2.5" stroke-linejoin="round"/>' +
      '<rect x="4" y="25" width="27" height="13" rx="3.5" fill="' + color + '"/>' +
      '<text x="17.5" y="34.5" text-anchor="middle" font-size="8" font-weight="800" fill="#fff" font-family="Segoe UI, system-ui, sans-serif">' + label + '</text>' +
      '</g></g>';
  }

  var uid = 0;
  var SCENES = {
    download: function () {
      var id = 'gs' + (++uid);
      return '<svg class="scene sc-dl" viewBox="0 0 360 240" aria-hidden="true">' + defs(id) +
        '<g class="sc-win">' +
        '<rect x="18" y="20" width="236" height="156" rx="16" fill="#fff" stroke="' + INK + '" stroke-opacity=".14" stroke-width="1.5"/>' +
        '<path d="M18 36a16 16 0 0 1 16-16h204a16 16 0 0 1 16 16v10H18z" fill="#eef2f7"/>' +
        '<circle cx="34" cy="33" r="4.2" fill="#f87171"/><circle cx="47" cy="33" r="4.2" fill="#fbbf24"/><circle cx="60" cy="33" r="4.2" fill="#34d399"/>' +
        '<rect x="40" y="62" width="112" height="10" rx="5" fill="#cbd5e1"/>' +
        '<rect x="40" y="80" width="176" height="7" rx="3.5" fill="#e2e8f0"/>' +
        '<g class="sc-cta"><rect x="58" y="106" width="156" height="38" rx="12" fill="url(#' + id + 'b)"/>' +
        '<text x="136" y="130" text-anchor="middle" font-size="13" font-weight="700" fill="#fff" font-family="Segoe UI, system-ui, sans-serif">Download template</text></g>' +
        '</g>' +
        '<g class="sc-tray"><rect x="270" y="164" width="70" height="56" rx="16" fill="' + INK + '"/>' +
        '<path d="M305 176v20M296 188l9 9 9-9" fill="none" stroke="#fff" stroke-width="3.2" stroke-linecap="round" stroke-linejoin="round"/>' +
        '<path d="M290 205h30" stroke="#94a3b8" stroke-width="3.2" stroke-linecap="round"/></g>' +
        '<g class="sc-badge"><circle cx="338" cy="166" r="11" fill="url(#' + id + 'a)" stroke="#fff" stroke-width="2.5"/><text x="338" y="170.5" text-anchor="middle" font-size="12" font-weight="800" fill="' + INK + '" font-family="Segoe UI, system-ui, sans-serif">1</text></g>' +
        fileIcon('sc-file', 'XLSX', '#10b981') +
        '<g class="sc-cursor"><path d="M0 0v24l6.5-6 4.2 9.6 4.6-2-4.2-9.4h9z" fill="' + INK + '" stroke="#fff" stroke-width="1.6" stroke-linejoin="round"/></g>' +
        '<g class="sc-note"><rect x="22" y="190" width="226" height="32" rx="11" fill="#ecfdf5" stroke="#10b981" stroke-width="1.5"/>' +
        '<circle cx="40" cy="206" r="8" fill="#10b981"/><path d="m36 206 3 3 5-5.5" fill="none" stroke="#fff" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"/>' +
        '<text x="56" y="210.5" font-size="11.5" font-weight="700" fill="#047857" font-family="Segoe UI, system-ui, sans-serif">Payslip_Template.xlsx saved</text></g>' +
        '</svg>';
    },

    fill: function () {
      var id = 'gs' + (++uid);
      var cols = [[26, 32, 'ID'], [62, 74, 'Name'], [140, 38, 'Basic'], [182, 38, 'Housing'], [224, 32, 'GOSI'], [260, 44, 'Net'], [308, 24, '✓']];
      var bands = [[22, 116, '#dbeafe', '#1e3a8a', 'EMPLOYEE'], [138, 84, '#d1fae5', '#047857', 'EARNINGS'], [222, 36, '#fee2e2', '#b91c1c', 'DED.'], [258, 76, '#fef3c7', '#92400e', 'RESULT']];
      var nets = ['3,650', '2,500', '4,120'];
      var s = '<svg class="scene sc-fill" viewBox="0 0 360 240" aria-hidden="true">' + defs(id) +
        '<rect x="14" y="12" width="332" height="176" rx="16" fill="#fff" stroke="' + INK + '" stroke-opacity=".14" stroke-width="1.5"/>';
      bands.forEach(function (b) {
        s += '<rect x="' + b[0] + '" y="22" width="' + b[1] + '" height="14" rx="4" fill="' + b[2] + '"/>' +
          '<text x="' + (b[0] + b[1] / 2) + '" y="32" text-anchor="middle" font-size="7.5" font-weight="800" letter-spacing=".08em" fill="' + b[3] + '" font-family="Segoe UI, system-ui, sans-serif">' + b[4] + '</text>';
      });
      s += '<rect x="22" y="40" width="312" height="18" rx="4" fill="#1e293b"/>';
      cols.forEach(function (c) {
        s += '<text x="' + (c[0] + 2) + '" y="52.5" font-size="8.5" font-weight="700" fill="#fff" font-family="Segoe UI, system-ui, sans-serif">' + c[2] + '</text>';
      });
      for (var r = 0; r < 3; r++) {
        var y = 64 + r * 26;
        var d = (0.35 + r * 0.9).toFixed(2);
        s += '<rect x="22" y="' + (y - 2) + '" width="312" height="22" rx="4" fill="' + (r % 2 ? '#f8fafc' : '#fff') + '"/>';
        cols.slice(0, 5).forEach(function (c, i) {
          s += '<rect class="sc-type" style="--d:' + (+d + i * 0.12).toFixed(2) + 's" x="' + (c[0] + 2) + '" y="' + (y + 5) + '" width="' + (c[1] - 8 - (i === 1 ? r * 10 : 0)) + '" height="8" rx="4" fill="' + (i === 0 ? '#94a3b8' : i === 1 ? '#475569' : '#cbd5e1') + '"/>';
        });
        s += '<text class="sc-net" style="--d:' + (+d + 0.75).toFixed(2) + 's" x="' + (cols[5][0] + 2) + '" y="' + (y + 12.5) + '" font-size="10" font-weight="800" fill="#1e3a8a" font-family="Segoe UI, system-ui, sans-serif">' + nets[r] + '</text>';
        s += '<g class="sc-ok" style="--d:' + (+d + 0.9).toFixed(2) + 's"><circle cx="' + (cols[6][0] + 10) + '" cy="' + (y + 9) + '" r="7.5" fill="url(#' + id + 'g)"/>' +
          '<path d="m' + (cols[6][0] + 6.5) + ' ' + (y + 9) + ' 2.5 2.5 4.5-5" fill="none" stroke="#fff" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></g>';
      }
      // sheet tabs + the logo going onto the Settings sheet
      s += '<g font-family="Segoe UI, system-ui, sans-serif" font-size="10" font-weight="700">' +
        '<rect x="22" y="154" width="70" height="24" rx="8" fill="#eef2f7" stroke="' + INK + '" stroke-opacity=".12"/><text x="57" y="170" text-anchor="middle" fill="#475569">Settings</text>' +
        '<rect x="98" y="154" width="70" height="24" rx="8" fill="#1e3a8a"/><text x="133" y="170" text-anchor="middle" fill="#fff">Payslips</text>' +
        '</g>' +
        '<g class="sc-logo"><rect x="30" y="196" width="54" height="34" rx="9" fill="url(#' + id + 'b)" stroke="#fff" stroke-width="2"/>' +
        '<path d="M40 222l9-10 7 7 4-4 9 7" fill="none" stroke="#fff" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"/><circle cx="66" cy="206" r="3" fill="#fde68a"/></g>' +
        '<g class="sc-tag"><text x="96" y="218" font-size="11" font-weight="700" fill="#475569" font-family="Segoe UI, system-ui, sans-serif">← your logo goes on the Settings sheet</text></g>' +
        '</svg>';
      return s;
    },

    drop: function () {
      var id = 'gs' + (++uid);
      return '<svg class="scene sc-drop" viewBox="0 0 360 240" aria-hidden="true">' + defs(id) +
        '<g class="sc-zone"><rect x="96" y="20" width="244" height="142" rx="18" fill="#f8fbff" stroke="#2563eb" stroke-width="2.5" stroke-dasharray="9 7"/>' +
        '<circle cx="218" cy="72" r="20" fill="url(#' + id + 'b)"/><path d="M218 82V63m-8 8 8-8 8 8" fill="none" stroke="#fff" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"/>' +
        '<text x="218" y="118" text-anchor="middle" font-size="13" font-weight="800" fill="' + INK + '" font-family="Segoe UI, system-ui, sans-serif">Drop your Excel file here</text>' +
        '<text x="218" y="136" text-anchor="middle" font-size="10" fill="#64748b" font-family="Segoe UI, system-ui, sans-serif">nothing leaves this computer</text></g>' +
        '<rect class="sc-glow" x="96" y="20" width="244" height="142" rx="18" fill="rgba(37,99,235,.10)" stroke="#2563eb" stroke-width="3"/>' +
        fileIcon('sc-fly', 'XLSX', '#10b981') +
        '<g class="sc-res" font-family="Segoe UI, system-ui, sans-serif" font-size="11" font-weight="800">' +
        '<g class="r1"><rect x="20" y="182" width="104" height="30" rx="15" fill="#ecfdf5" stroke="#10b981"/><text x="72" y="201" text-anchor="middle" fill="#047857">✓ 25 employees</text></g>' +
        '<g class="r2"><rect x="132" y="182" width="88" height="30" rx="15" fill="#eff6ff" stroke="#2563eb"/><text x="176" y="201" text-anchor="middle" fill="#1d4ed8">0 errors</text></g>' +
        '<g class="r3"><rect x="228" y="182" width="112" height="30" rx="15" fill="#fffbeb" stroke="#f59e0b"/><text x="284" y="201" text-anchor="middle" fill="#92400e">Net 112,450</text></g>' +
        '</g>' +
        '</svg>';
    },

    save: function () {
      var id = 'gs' + (++uid);
      var chip = function (cls, x, y, w, label, color) {
        return '<g class="sc-chip ' + cls + '"><rect x="' + x + '" y="' + y + '" width="' + w + '" height="28" rx="14" fill="#fff" stroke="' + color + '" stroke-width="2"/>' +
          '<circle cx="' + (x + 14) + '" cy="' + (y + 14) + '" r="5" fill="' + color + '"/>' +
          '<text x="' + (x + 25) + '" y="' + (y + 18.5) + '" font-size="11" font-weight="800" fill="' + INK + '" font-family="Segoe UI, system-ui, sans-serif">' + label + '</text></g>';
      };
      return '<svg class="scene sc-save" viewBox="0 0 360 240" aria-hidden="true">' + defs(id) +
        '<circle class="sc-ring" cx="180" cy="104" r="62" fill="none" stroke="#2563eb" stroke-width="2" stroke-dasharray="4 8"/>' +
        chip('c1', 22, 30, 96, 'Company', '#2563eb') +
        chip('c2', 250, 30, 80, 'Logo', '#f59e0b') +
        chip('c3', 22, 150, 84, 'Design', '#0ea5e9') +
        chip('c4', 248, 150, 92, 'Columns', '#10b981') +
        '<g class="sc-doc"><path d="M152 62h42l16 16v62h-58z" fill="#fff" stroke="' + INK + '" stroke-width="3" stroke-linejoin="round"/>' +
        '<path d="M194 62v16h16" fill="#e2e8f0" stroke="' + INK + '" stroke-width="3" stroke-linejoin="round"/>' +
        '<rect x="162" y="92" width="30" height="6" rx="3" fill="#cbd5e1"/><rect x="162" y="104" width="38" height="6" rx="3" fill="#e2e8f0"/>' +
        '<rect x="160" y="118" width="42" height="14" rx="4" fill="#1e3a8a"/><text x="181" y="128.5" text-anchor="middle" font-size="8" font-weight="800" fill="#fff" font-family="Segoe UI, system-ui, sans-serif">JSON</text></g>' +
        '<g class="sc-shield"><path d="M206 108l20 7v14c0 13-9 21-20 25-11-4-20-12-20-25v-14z" fill="url(#' + id + 'g)" stroke="#fff" stroke-width="3" stroke-linejoin="round"/>' +
        '<path d="m198 131 6 6 10-11" fill="none" stroke="#fff" stroke-width="3.2" stroke-linecap="round" stroke-linejoin="round"/></g>' +
        '<g class="sc-stamp"><rect x="118" y="186" width="124" height="34" rx="10" fill="none" stroke="#059669" stroke-width="3"/>' +
        '<text x="180" y="208.5" text-anchor="middle" font-size="14" font-weight="900" letter-spacing=".12em" fill="#059669" font-family="Segoe UI, system-ui, sans-serif">BACKED UP</text></g>' +
        '</svg>';
    }
  };

  // ---------------------------------------------------------------------------
  // Quick start tour
  // ---------------------------------------------------------------------------
  var TOUR = [
    { scene: 'download', mood: 'idle', kicker: 'Step 1', title: 'Grab the template',
      text: 'One Excel file with everything ready: colour-coded columns, drop-downs, and checks that catch mistakes while you type.',
      action: 'download' },
    { scene: 'fill', mood: 'design', kicker: 'Step 2', title: 'Fill it in Excel',
      text: 'Company details once on <b>Settings</b>, one row per employee on <b>Payslips</b>. The Net column works itself out and ✓ means the row is good.' },
    { scene: 'drop', mood: 'look', kicker: 'Step 3', title: 'Drop it on this page',
      text: 'Every row is checked in a second. Your own salary sheet works too; anything unclear is asked on the next screen.' },
    { scene: 'save', mood: 'export', kicker: 'Step 4', title: 'Save your settings file',
      text: 'Before exporting, keep a <b>settings file</b>: logo, company, design and columns in one small file. Next month, load it and you are done in a minute. If this browser is ever cleared, nothing is lost.',
      action: 'finish' }
  ];

  function tour(opts) {
    opts = opts || {};
    return new Promise(function (resolve) {
      var i = Math.max(0, Math.min(TOUR.length - 1, opts.start || 0));
      var downloaded = false;
      var dots = TOUR.map(function (t, k) { return '<button type="button" class="tour-dot" data-k="' + k + '" aria-label="Step ' + (k + 1) + '"></button>'; }).join('');
      var m = modal({
        className: 'tour',
        label: 'Quick start',
        html:
          '<div class="tour-top"><span class="tour-kicker">Quick start · 30 seconds</span><div class="tour-bar"><i></i></div></div>' +
          '<div class="tour-grid">' +
          '<div class="tour-stage"><div class="tour-scene"></div></div>' +
          '<div class="tour-side"><div class="tour-qasi">' + Art.mascot('q-own') + '</div><div class="tour-copy" aria-live="polite"></div></div>' +
          '</div>' +
          '<div class="tour-foot"><button type="button" class="btn btn-ghost btn-small" data-act="skip">Skip</button>' +
          '<div class="tour-dots">' + dots + '</div>' +
          '<div class="tour-nav"><button type="button" class="btn btn-small" data-act="back">Back</button>' +
          '<button type="button" class="btn btn-primary" data-act="next" data-autofocus>Next</button></div></div>',
        onKey: function (e) {
          if (e.target && /INPUT|TEXTAREA|SELECT/.test(e.target.tagName)) return;
          if (e.key === 'ArrowRight') { e.preventDefault(); go(i + 1); }
          if (e.key === 'ArrowLeft') { e.preventDefault(); go(i - 1); }
        },
        onClose: function (result) { resolve(result === 'done' ? 'done' : 'skipped'); }
      });
      var el = m.el;

      function go(k) {
        if (k >= TOUR.length) { finish(); return; }
        if (k < 0) return;
        var dir = k >= i ? 'fwd' : 'back';
        i = k;
        render(dir);
      }

      function finish() {
        Art.confetti({ count: 120, y: innerHeight / 2 });
        m.close('done');
      }

      function render(dir) {
        var t = TOUR[i];
        var stage = el.querySelector('.tour-scene');
        stage.innerHTML = SCENES[t.scene]();
        stage.className = 'tour-scene enter-' + (dir || 'fwd');
        el.querySelector('.tour-qasi .qasi').setAttribute('data-mood', t.mood);
        var actionHtml = '';
        if (t.action === 'download') {
          actionHtml = '<div class="tour-action"><button type="button" class="btn btn-primary" data-act="download">' +
            (downloaded ? '✓ Downloaded, get it again' : 'Download template now') + '</button>' +
            '<button type="button" class="btn btn-small btn-ghost" data-act="sample">or a filled example</button></div>';
        } else if (t.action === 'finish' && opts.canInstall && opts.canInstall()) {
          actionHtml = '<div class="tour-action"><button type="button" class="btn btn-small" data-act="install">Install as an app</button><span class="muted small">opens from the desktop, works offline</span></div>';
        }
        el.querySelector('.tour-copy').innerHTML =
          '<span class="tour-step">' + esc(t.kicker) + ' of ' + TOUR.length + '</span>' +
          '<h2>' + esc(t.title) + '</h2><p>' + t.text + '</p>' + actionHtml;
        el.querySelector('.tour-copy').className = 'tour-copy enter-' + (dir || 'fwd');
        el.querySelector('.tour-bar i').style.width = ((i + 1) / TOUR.length * 100) + '%';
        Array.prototype.forEach.call(el.querySelectorAll('.tour-dot'), function (d, k) {
          d.classList.toggle('on', k === i);
          d.classList.toggle('done', k < i);
        });
        el.querySelector('[data-act="back"]').disabled = i === 0;
        el.querySelector('[data-act="next"]').textContent = i === TOUR.length - 1 ? 'Let’s go' : 'Next';
      }

      el.addEventListener('click', function (e) {
        var b = e.target.closest('[data-act], .tour-dot');
        if (!b) return;
        if (b.classList.contains('tour-dot')) { go(+b.getAttribute('data-k')); return; }
        var act = b.getAttribute('data-act');
        if (act === 'next') go(i + 1);
        else if (act === 'back') go(i - 1);
        else if (act === 'skip') m.close('skip');
        else if (act === 'download' || act === 'sample') {
          if (opts.onDownload) opts.onDownload(act === 'sample');
          downloaded = true;
          b.classList.add('pressed');
          var stage = el.querySelector('.tour-scene');
          stage.innerHTML = SCENES.download();
          setTimeout(function () { if (i === 0) render('fwd'); }, 900);
        } else if (act === 'install' && opts.onInstall) opts.onInstall();
      });
      render('fwd');
    });
  }

  // ---------------------------------------------------------------------------
  // Save point: shown before export while the settings have no backup file.
  // ---------------------------------------------------------------------------
  function savePoint(opts) {
    opts = opts || {};
    return new Promise(function (resolve) {
      var items = (opts.items || []).map(function (it, k) {
        return '<li style="--k:' + k + '" class="' + (it.ok ? 'ok' : 'off') + '"><span class="sp-ico">' + (it.thumb ? '<img src="' + esc(it.thumb) + '" alt="">' : it.ok ? '✓' : '–') + '</span>' +
          '<span class="sp-lbl">' + esc(it.label) + '</span><span class="sp-val">' + esc(it.value || '') + '</span></li>';
      }).join('');
      var m = modal({
        className: 'savepoint',
        label: 'Save point',
        html:
          '<div class="sp-ribbon"><span>Save point</span></div>' +
          '<div class="sp-grid">' +
          '<div class="sp-stage">' + SCENES.save() + '<div class="sp-qasi">' + Art.mascot('q-own') + '</div></div>' +
          '<div class="sp-copy">' +
          '<h2>' + esc(opts.title || 'Save your progress before exporting') + '</h2>' +
          '<p>' + (opts.text || 'Everything you set up is in this browser only. One small file keeps it safe for next month and any other computer.') + '</p>' +
          '<ul class="sp-list">' + items + '</ul>' +
          '<div class="sp-actions"><button type="button" class="btn btn-primary btn-xl" data-act="save">' + esc(opts.saveLabel || 'Save settings file') + '</button>' +
          '<button type="button" class="btn btn-ghost" data-act="skip">' + esc(opts.skipLabel || 'Continue without saving') + '</button></div>' +
          (opts.note ? '<p class="muted small sp-note">' + opts.note + '</p>' : '') +
          '</div></div>',
        onClose: function (result) { resolve(result || 'dismiss'); }
      });
      m.el.querySelector('.qasi').setAttribute('data-mood', 'export');
      m.el.addEventListener('click', function (e) {
        var b = e.target.closest('[data-act]');
        if (!b) return;
        if (b.getAttribute('data-act') === 'skip') { m.close('skip'); return; }
        b.disabled = true;
        Promise.resolve(opts.onSave ? opts.onSave() : true).then(function (ok) {
          if (!ok) { b.disabled = false; return; }
          m.el.classList.add('saved');
          m.el.querySelector('.qasi').setAttribute('data-mood', 'celebrate');
          b.textContent = '✓ Saved';
          Art.confetti({ count: 140, y: innerHeight / 2.4 });
          setTimeout(function () { m.close('saved'); }, 1300);
        }, function () { b.disabled = false; });
      });
    });
  }

  // ---------------------------------------------------------------------------
  // Confirm dialog
  // ---------------------------------------------------------------------------
  function confirmBox(opts) {
    opts = opts || {};
    return new Promise(function (resolve) {
      var m = modal({
        className: 'confirm' + (opts.danger ? ' danger' : ''),
        label: opts.title,
        html: '<div class="cf-head"><div class="cf-qasi">' + Art.mascot('q-own') + '</div><h2>' + esc(opts.title || 'Are you sure?') + '</h2></div>' +
          '<div class="cf-body">' + (opts.html || '') + '</div>' +
          '<div class="cf-actions"><button type="button" class="btn" data-act="no">' + esc(opts.cancel || 'Cancel') + '</button>' +
          '<button type="button" class="btn ' + (opts.danger ? 'btn-danger' : 'btn-primary') + '" data-act="yes" data-autofocus>' + esc(opts.ok || 'OK') + '</button></div>',
        onClose: function (result) { resolve(result === 'yes'); }
      });
      m.el.querySelector('.qasi').setAttribute('data-mood', opts.mood || 'oops');
      m.el.addEventListener('click', function (e) {
        var b = e.target.closest('[data-act]');
        if (b) m.close(b.getAttribute('data-act'));
      });
    });
  }

  // ---------------------------------------------------------------------------
  // Toasts
  // ---------------------------------------------------------------------------
  function toast(opts) {
    opts = typeof opts === 'string' ? { html: opts } : (opts || {});
    var stack = document.getElementById('toastStack');
    if (!stack) {
      stack = document.createElement('div');
      stack.id = 'toastStack';
      stack.className = 'toast-stack';
      stack.setAttribute('role', 'status');
      stack.setAttribute('aria-live', 'polite');
      document.body.appendChild(stack);
    }
    if (opts.id) {
      var old = document.getElementById(opts.id);
      if (old) old.remove();
    }
    var t = document.createElement('div');
    t.className = 'toast toast-' + (opts.kind || 'info');
    if (opts.id) t.id = opts.id;
    var icons = { ok: '✓', info: 'i', warn: '!', save: '💾' };
    t.innerHTML = '<span class="toast-ico">' + (opts.icon || icons[opts.kind || 'info'] || 'i') + '</span><div class="toast-body">' + opts.html +
      (opts.actions && opts.actions.length ? '<div class="toast-actions">' + opts.actions.map(function (a, k) {
        return '<button type="button" class="btn btn-small' + (a.primary ? ' btn-primary' : '') + '" data-k="' + k + '">' + esc(a.label) + '</button>';
      }).join('') + '</div>' : '') + '</div><button type="button" class="toast-x" aria-label="Dismiss">×</button>' +
      (opts.timeout !== 0 ? '<i class="toast-timer" style="animation-duration:' + (opts.timeout || 6000) + 'ms"></i>' : '');
    stack.appendChild(t);
    var timer = null;
    function close() {
      clearTimeout(timer);
      t.classList.add('leaving');
      setTimeout(function () { t.remove(); }, 300);
    }
    t.addEventListener('click', function (e) {
      if (e.target.closest('.toast-x')) { close(); return; }
      var b = e.target.closest('[data-k]');
      if (!b) return;
      var a = opts.actions[+b.getAttribute('data-k')];
      if (a && a.onClick) a.onClick();
      if (!a || a.keepOpen !== true) close();
    });
    if (opts.timeout !== 0) {
      timer = setTimeout(close, opts.timeout || 6000);
      t.addEventListener('mouseenter', function () { clearTimeout(timer); t.classList.add('hold'); });
      t.addEventListener('mouseleave', function () { t.classList.remove('hold'); timer = setTimeout(close, 2500); });
    }
    return { el: t, close: close };
  }

  root.PayslipGuide = {
    modal: modal,
    tour: tour,
    savePoint: savePoint,
    confirm: confirmBox,
    toast: toast,
    scene: function (name) { return SCENES[name] ? SCENES[name]() : ''; },
    TOUR: TOUR
  };
})(typeof self !== 'undefined' ? self : this);
