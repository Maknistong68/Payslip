/*
 * Payslip Generator - installable app (PWA). When the page is served over http(s) it registers the
 * service worker (every file cached, so it opens without internet), offers "Install app", tells the
 * user when a new version is ready, and opens .xlsx/.csv files handed to the installed app.
 * Opened as a file (USB stick, shared drive) none of this applies and the page works as before.
 */
(function () {
  'use strict';

  var Guide = window.PayslipGuide;
  var App = window.PayslipApp;
  var served = /^https?:$/.test(location.protocol);
  var deferredPrompt = null;
  var installBtn = document.getElementById('btnInstall');

  function standalone() {
    return (window.matchMedia && matchMedia('(display-mode: standalone)').matches) || navigator.standalone === true;
  }

  function install() {
    if (!deferredPrompt) return Promise.resolve(false);
    var p = deferredPrompt;
    deferredPrompt = null;
    installBtn.classList.add('hidden');
    p.prompt();
    return p.userChoice.then(function (c) { return c && c.outcome === 'accepted'; }).catch(function () { return false; });
  }

  window.PayslipPWA = {
    canInstall: function () { return !!deferredPrompt; },
    install: install,
    standalone: standalone
  };

  if (!served) return;

  // The manifest is added here, not in the HTML, so pages opened from disk do not log a fetch error.
  var link = document.createElement('link');
  link.rel = 'manifest';
  link.href = 'manifest.webmanifest';
  document.head.appendChild(link);

  window.addEventListener('beforeinstallprompt', function (e) {
    e.preventDefault();
    deferredPrompt = e;
    if (!standalone()) installBtn.classList.remove('hidden');
  });
  installBtn.addEventListener('click', install);
  window.addEventListener('appinstalled', function () {
    deferredPrompt = null;
    installBtn.classList.add('hidden');
    if (Guide) Guide.toast({ kind: 'ok', html: '<b>Installed.</b> Open Payslips from your desktop or start menu; it works without internet.' });
  });

  // Files opened with the installed app ("Open with › Payslips").
  if ('launchQueue' in window && App && App.handleFile) {
    window.launchQueue.setConsumer(function (params) {
      if (!params.files || !params.files.length) return;
      params.files[0].getFile().then(function (f) { App.handleFile(f); });
    });
  }

  if (!('serviceWorker' in navigator)) return;

  var reloading = false;
  var userAskedUpdate = false;
  navigator.serviceWorker.addEventListener('controllerchange', function () {
    if (!userAskedUpdate || reloading) return;
    reloading = true;
    location.reload();
  });

  function offerUpdate(reg) {
    if (!Guide || !reg.waiting) return;
    Guide.toast({
      id: 'updateToast', kind: 'info', icon: '★', timeout: 0,
      html: '<b>A new version is ready.</b> Your open file and ticks are saved and come straight back after the reload.',
      actions: [
        { label: 'Reload now', primary: true, keepOpen: true, onClick: function () {
          if (App && App.state && App.state.exporting) {
            Guide.toast({ kind: 'warn', html: 'Export in progress. Reload when it has finished.' });
            return;
          }
          userAskedUpdate = true;
          Promise.resolve(App && App.flushSession ? App.flushSession() : null).then(function () {
            if (reg.waiting) reg.waiting.postMessage({ type: 'SKIP_WAITING' });
            else location.reload();
          });
        } },
        { label: 'Later' }
      ]
    });
  }

  window.addEventListener('load', function () {
    navigator.serviceWorker.register('sw.js').then(function (reg) {
      if (reg.waiting && navigator.serviceWorker.controller) offerUpdate(reg);
      reg.addEventListener('updatefound', function () {
        var w = reg.installing;
        if (!w) return;
        w.addEventListener('statechange', function () {
          if (w.state === 'installed' && navigator.serviceWorker.controller) offerUpdate(reg);
        });
      });
      // look for a new version when the user comes back to the tab, at most once an hour
      var last = Date.now();
      document.addEventListener('visibilitychange', function () {
        if (document.visibilityState === 'visible' && Date.now() - last > 3600e3) {
          last = Date.now();
          reg.update().catch(function () { /* offline */ });
        }
      });
    }).catch(function () { /* private mode or blocked: the page still works */ });
  });
})();
