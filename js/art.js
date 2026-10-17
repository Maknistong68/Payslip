/*
 * Payslip Generator - art layer: logo, the mascot "Qasi", confetti, count-up numbers, card spotlight.
 * No dependencies, works offline. Everything respects prefers-reduced-motion.
 */
(function (root) {
  'use strict';

  var reduced = !!(root.matchMedia && root.matchMedia('(prefers-reduced-motion: reduce)').matches);
  var uid = 0;

  // ---------------------------------------------------------------------------
  // Logo: a torn-edge receipt (the payslip) leaning on an aurora squircle, with a "sent" check badge.
  // ---------------------------------------------------------------------------
  function logo(size) {
    var id = 'lg' + (++uid);
    size = size || 36;
    return '<svg class="logo-mark" width="' + size + '" height="' + size + '" viewBox="0 0 48 48" aria-hidden="true">' +
      '<defs>' +
      '<linearGradient id="' + id + 'a" x1="0" y1="0" x2="1" y2="1">' +
      '<stop offset="0" stop-color="#1e3a8a"><animate attributeName="stop-color" values="#1e3a8a;#2563eb;#0e7490;#1e3a8a" dur="9s" repeatCount="indefinite"/></stop>' +
      '<stop offset="1" stop-color="#0ea5e9"><animate attributeName="stop-color" values="#0ea5e9;#1e3a8a;#2563eb;#0ea5e9" dur="9s" repeatCount="indefinite"/></stop>' +
      '</linearGradient>' +
      '<linearGradient id="' + id + 'b" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#fff"/><stop offset="1" stop-color="#e2e8f0"/></linearGradient>' +
      '</defs>' +
      '<rect x="2" y="2" width="44" height="44" rx="14" fill="url(#' + id + 'a)"/>' +
      '<rect x="2" y="2" width="44" height="22" rx="14" fill="#fff" opacity=".12"/>' +
      '<g transform="rotate(-8 24 25)">' +
      '<path d="M15 10.5h14.5l5 5V35l-2.5 2.5-2.5-2.5-2.5 2.5-2.5-2.5-2.5 2.5-2.5-2.5-2.5 2.5-2.5-2.5V12.5a2 2 0 0 1 2-2z" fill="url(#' + id + 'b)"/>' +
      '<path d="M29.5 10.5v3.5a1.5 1.5 0 0 0 1.5 1.5h3.5z" fill="#cbd5e1"/>' +
      '<rect x="17" y="17" width="9" height="2.4" rx="1.2" fill="#93c5fd"/>' +
      '<rect x="17" y="22" width="13" height="2.4" rx="1.2" fill="#cbd5e1"/>' +
      '<rect x="17" y="27" width="7" height="2.4" rx="1.2" fill="url(#' + id + 'a)"/>' +
      '</g>' +
      '<circle cx="35" cy="34" r="7.5" fill="#10b981" stroke="#fff" stroke-width="2.2"/>' +
      '<path d="m31.6 34 2.3 2.3 4.4-4.6" fill="none" stroke="#fff" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"/>' +
      '</svg>';
  }

  // ---------------------------------------------------------------------------
  // Mascot: Qasi, a living payslip (torn receipt edge, folded corner, little feet).
  // Moods (data-mood): idle, look, design, export, send, celebrate, oops.
  // ---------------------------------------------------------------------------
  function mascot(cls) {
    var id = 'qs' + (++uid);
    var ink = '#0f172a';
    return '<svg class="qasi ' + (cls || '') + '" viewBox="0 -18 220 248" data-mood="idle" role="img" aria-label="Qasi, the payslip mascot">' +
      '<defs>' +
      '<linearGradient id="' + id + 'b" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#ffffff"/><stop offset="1" stop-color="#e2e8f0"/></linearGradient>' +
      '<linearGradient id="' + id + 'h" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#fcd34d"/><stop offset="1" stop-color="#f59e0b"/></linearGradient>' +
      '<linearGradient id="' + id + 'g" x1="0" y1="0" x2="1" y2="0">' +
      '<stop offset="0" stop-color="#2563eb"><animate attributeName="stop-color" values="#2563eb;#0ea5e9;#0e7490;#2563eb" dur="8s" repeatCount="indefinite"/></stop>' +
      '<stop offset="1" stop-color="#0e7490"><animate attributeName="stop-color" values="#0e7490;#2563eb;#0ea5e9;#0e7490" dur="8s" repeatCount="indefinite"/></stop>' +
      '</linearGradient>' +
      '</defs>' +
      '<ellipse class="q-shadow" cx="110" cy="220" rx="48" ry="7" fill="' + ink + '" opacity=".14"/>' +
      '<g class="q-float">' +
      // feet
      '<g class="q-feet"><ellipse cx="90" cy="203" rx="10" ry="6" fill="' + ink + '"/><ellipse cx="130" cy="203" rx="10" ry="6" fill="' + ink + '"/></g>' +
      // left arm
      '<g class="q-arm-l"><path d="M50 118 Q30 122 26 106" fill="none" stroke="' + ink + '" stroke-width="5" stroke-linecap="round"/><circle cx="26" cy="103" r="7" fill="#fff" stroke="' + ink + '" stroke-width="3.5"/></g>' +
      // body
      '<path d="M62 22 H146 L170 46 V178 L160 190 L150 178 L140 190 L130 178 L120 190 L110 178 L100 190 L90 178 L80 190 L70 178 L60 190 L50 178 V34 Q50 22 62 22 Z" fill="url(#' + id + 'b)" stroke="' + ink + '" stroke-width="4" stroke-linejoin="round"/>' +
      '<path d="M146 22 V40 Q146 46 152 46 H170 Z" fill="#cbd5e1" stroke="' + ink + '" stroke-width="4" stroke-linejoin="round"/>' +
      '<g class="q-hat">' +
      '<path d="M66 22 Q66 -12 108 -12 Q150 -12 150 22 Z" fill="url(#' + id + 'h)" stroke="' + ink + '" stroke-width="4" stroke-linejoin="round"/>' +
      '<path d="M100 -11 Q100 2 100 21 M116 -11 Q116 2 116 21" stroke="#d97706" stroke-width="3" fill="none"/>' +
      '<rect x="50" y="17" width="116" height="10" rx="5" fill="#f59e0b" stroke="' + ink + '" stroke-width="4"/>' +
      '<path d="M76 6 Q80 -4 90 -7" stroke="#fff" stroke-width="3" stroke-linecap="round" fill="none" opacity=".7"/>' +
      '</g>' +
      // face
      '<g class="q-eyes"><g class="q-pupils">' +
      '<ellipse cx="90" cy="78" rx="8" ry="10" fill="' + ink + '"/><ellipse cx="130" cy="78" rx="8" ry="10" fill="' + ink + '"/>' +
      '<circle cx="93" cy="74" r="3" fill="#fff"/><circle cx="133" cy="74" r="3" fill="#fff"/>' +
      '</g></g>' +
      '<g class="q-brows"><path d="M78 58 L100 63" stroke="' + ink + '" stroke-width="5" stroke-linecap="round"/><path d="M120 63 L142 58" stroke="' + ink + '" stroke-width="5" stroke-linecap="round"/></g>' +
      '<path class="q-mouth q-m-smile" d="M98 99 Q112 108 125 95" fill="none" stroke="' + ink + '" stroke-width="4" stroke-linecap="round"/>' +
      '<ellipse class="q-mouth q-m-o" cx="110" cy="101" rx="6" ry="7" fill="' + ink + '"/>' +
      '<path class="q-mouth q-m-grin" d="M94 94 Q110 118 126 94 Z" fill="' + ink + '" stroke="' + ink + '" stroke-width="3" stroke-linejoin="round"/>' +
      '<path class="q-mouth q-m-flat" d="M99 101 Q110 96 121 101" fill="none" stroke="' + ink + '" stroke-width="4" stroke-linecap="round"/>' +
      // body lines + the "net pay" bar
      '<rect x="72" y="128" width="76" height="8" rx="4" fill="#cbd5e1"/>' +
      '<rect x="72" y="144" width="54" height="8" rx="4" fill="#e2e8f0"/>' +
      '<rect class="q-net" x="72" y="160" width="76" height="10" rx="5" fill="url(#' + id + 'g)"/>' +
      // right arm (waves)
      '<g class="q-arm-r"><path d="M170 118 Q192 112 196 94" fill="none" stroke="' + ink + '" stroke-width="5" stroke-linecap="round"/><circle cx="196" cy="91" r="7" fill="#fff" stroke="' + ink + '" stroke-width="3.5"/></g>' +
      // props
      '<g class="q-prop q-glass"><circle cx="198" cy="72" r="15" fill="rgba(255,255,255,.55)" stroke="' + ink + '" stroke-width="4"/><path d="M207 84 L216 96" stroke="' + ink + '" stroke-width="5" stroke-linecap="round"/><path d="M191 66 q4-5 10-3" fill="none" stroke="#fff" stroke-width="3" stroke-linecap="round"/></g>' +
      '<g class="q-prop q-plane"><path d="M188 70 L218 58 L206 88 L200 76 Z" fill="#fff" stroke="' + ink + '" stroke-width="3.5" stroke-linejoin="round"/><path d="M200 76 L218 58" stroke="' + ink + '" stroke-width="3" stroke-linecap="round"/></g>' +
      '<g class="q-prop q-lock"><rect x="14" y="96" width="24" height="20" rx="5" fill="url(#' + id + 'g)" stroke="' + ink + '" stroke-width="3.5"/><path d="M19 96 v-5 a7 7 0 0 1 14 0 v5" fill="none" stroke="' + ink + '" stroke-width="3.5"/><circle cx="26" cy="106" r="2.6" fill="#fff"/></g>' +
      '<g class="q-prop q-spark">' +
      '<path class="s1" d="M40 30 l3 8 8 3 -8 3 -3 8 -3-8 -8-3 8-3z" fill="#fbbf24"/>' +
      '<path class="s2" d="M184 20 l2.4 6 6 2.4 -6 2.4 -2.4 6 -2.4-6 -6-2.4 6-2.4z" fill="#0ea5e9"/>' +
      '<path class="s3" d="M28 150 l2 5 5 2 -5 2 -2 5 -2-5 -5-2 5-2z" fill="#10b981"/>' +
      '</g>' +
      '<g class="q-prop q-bang"><circle cx="188" cy="34" r="14" fill="#f59e0b" stroke="' + ink + '" stroke-width="3.5"/><path d="M188 26 v9" stroke="#fff" stroke-width="4" stroke-linecap="round"/><circle cx="188" cy="41" r="2.4" fill="#fff"/></g>' +
      '</g>' +
      '</svg>';
  }

  function renderArt(scope) {
    var r = scope || document;
    Array.prototype.forEach.call(r.querySelectorAll('[data-art="logo"]'), function (el) {
      el.innerHTML = logo(+el.getAttribute('data-size') || 36);
    });
    Array.prototype.forEach.call(r.querySelectorAll('[data-art="mascot"]'), function (el) {
      el.innerHTML = mascot(el.getAttribute('data-class') || '');
    });
  }

  // ---------------------------------------------------------------------------
  // Buddy: the floating mascot with a speech bubble.
  // ---------------------------------------------------------------------------
  var buddyTimer = null;
  var quietTimer = null;
  function say(mood, html, opts) {
    opts = opts || {};
    var buddy = document.getElementById('buddy');
    if (!buddy) return;
    // dialogs own their mascot (class q-own); the buddy and hero follow the app
    Array.prototype.forEach.call(document.querySelectorAll('.qasi:not(.q-own)'), function (svg) {
      svg.setAttribute('data-mood', mood || 'idle');
    });
    var bubble = document.getElementById('buddyBubble');
    if (html != null && bubble) {
      bubble.innerHTML = html;
      bubble.classList.remove('pop');
      void bubble.offsetWidth; // restart the pop animation
      bubble.classList.add('pop');
      buddy.classList.remove('quiet');
    }
    clearTimeout(buddyTimer);
    clearTimeout(quietTimer);
    if (html != null) {
      // the bubble fades away on its own so it never covers the tables for long; click Qasi to bring it back
      quietTimer = setTimeout(function () {
        if (!buddy.matches(':hover')) buddy.classList.add('quiet');
      }, opts.hold || 9000);
    }
    if (opts.settle) {
      buddyTimer = setTimeout(function () {
        Array.prototype.forEach.call(document.querySelectorAll('.qasi:not(.q-own)'), function (svg) {
          svg.setAttribute('data-mood', opts.settle);
        });
      }, opts.after || 2600);
    }
  }

  function setupBuddy() {
    var buddy = document.getElementById('buddy');
    if (!buddy) return;
    buddy.querySelector('.buddy-bot').addEventListener('click', function () {
      buddy.classList.toggle('quiet');
      buddy.classList.remove('hop');
      void buddy.offsetWidth;
      buddy.classList.add('hop');
    });
  }

  // ---------------------------------------------------------------------------
  // Confetti (canvas, ~2.5 s).
  // ---------------------------------------------------------------------------
  function confetti(opts) {
    if (reduced) return;
    opts = opts || {};
    var canvas = document.createElement('canvas');
    canvas.className = 'confetti-canvas';
    document.body.appendChild(canvas);
    var dpr = Math.min(window.devicePixelRatio || 1, 2);
    var W = canvas.width = innerWidth * dpr;
    var H = canvas.height = innerHeight * dpr;
    var ctx = canvas.getContext('2d');
    var colors = ['#2563eb', '#0ea5e9', '#f59e0b', '#10b981', '#0f172a', '#94a3b8'];
    var parts = [];
    var count = opts.count || 160;
    var ox = (opts.x != null ? opts.x : innerWidth / 2) * dpr;
    var oy = (opts.y != null ? opts.y : innerHeight / 3) * dpr;
    for (var i = 0; i < count; i++) {
      var a = Math.random() * Math.PI * 2;
      var sp = (4 + Math.random() * 9) * dpr;
      parts.push({
        x: ox, y: oy,
        vx: Math.cos(a) * sp, vy: Math.sin(a) * sp - 6 * dpr,
        w: (6 + Math.random() * 6) * dpr, h: (8 + Math.random() * 10) * dpr,
        r: Math.random() * Math.PI, vr: (Math.random() - 0.5) * 0.3,
        c: colors[i % colors.length], shape: i % 3
      });
    }
    var start = performance.now();
    (function frame(t) {
      var el = t - start;
      ctx.clearRect(0, 0, W, H);
      parts.forEach(function (p) {
        p.vy += 0.32 * dpr;
        p.vx *= 0.985;
        p.x += p.vx;
        p.y += p.vy;
        p.r += p.vr;
        ctx.save();
        ctx.globalAlpha = Math.max(0, 1 - el / 2600);
        ctx.translate(p.x, p.y);
        ctx.rotate(p.r);
        ctx.fillStyle = p.c;
        if (p.shape === 0) ctx.fillRect(-p.w / 2, -p.h / 2, p.w, p.h);
        else if (p.shape === 1) { ctx.beginPath(); ctx.arc(0, 0, p.w / 2, 0, Math.PI * 2); ctx.fill(); }
        else { ctx.beginPath(); ctx.moveTo(0, -p.h / 2); ctx.lineTo(p.w / 2, p.h / 2); ctx.lineTo(-p.w / 2, p.h / 2); ctx.closePath(); ctx.fill(); }
        ctx.restore();
      });
      if (el < 2700) requestAnimationFrame(frame);
      else canvas.remove();
    })(start);
  }

  // ---------------------------------------------------------------------------
  // Count-up numbers.
  // ---------------------------------------------------------------------------
  function countUp(el, to, opts) {
    opts = opts || {};
    var decimals = opts.decimals || 0;
    var fmt = opts.format || function (v) { return v.toFixed(decimals); };
    if (reduced || !isFinite(to)) { el.textContent = fmt(to); return; }
    var dur = opts.duration || 900;
    var t0 = performance.now();
    (function step(t) {
      var k = Math.min(1, (t - t0) / dur);
      var e = 1 - Math.pow(1 - k, 3);
      el.textContent = fmt(to * e);
      if (k < 1) requestAnimationFrame(step);
    })(t0);
  }

  // ---------------------------------------------------------------------------
  // Spotlight: cards glow where the pointer is.
  // ---------------------------------------------------------------------------
  function setupSpotlight() {
    document.addEventListener('pointermove', function (e) {
      var card = e.target.closest && e.target.closest('.spot');
      if (!card) return;
      var r = card.getBoundingClientRect();
      card.style.setProperty('--mx', (e.clientX - r.left) + 'px');
      card.style.setProperty('--my', (e.clientY - r.top) + 'px');
    }, { passive: true });
  }

  function init() {
    renderArt();
    setupBuddy();
    setupSpotlight();
  }

  root.PayslipArt = {
    logo: logo,
    mascot: mascot,
    renderArt: renderArt,
    say: say,
    confetti: confetti,
    countUp: countUp,
    reducedMotion: reduced,
    init: init
  };
})(typeof self !== 'undefined' ? self : this);
