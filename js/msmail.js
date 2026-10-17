/*
 * Payslip Generator - "Sign in with Microsoft": sends the emails straight from this page through the
 * user's own Microsoft 365 or Outlook.com mailbox (Microsoft Graph sendMail). There is no server in
 * between: the sign-in goes from the browser to Microsoft, the sign-in token is kept only in this page's
 * memory (gone when the page closes), and each email goes from the browser to Microsoft. Nothing is
 * stored or sent anywhere else. Standard OAuth for web pages: authorization code with PKCE in a popup.
 * The pure parts (links, messages, retries) also run in Node for the tests.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.PayslipMs = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  var GRAPH_MAIL_SEND = 'https://graph.microsoft.com/Mail.Send';
  var cfg = {
    clientId: '',
    authority: 'https://login.microsoftonline.com/common',
    graph: 'https://graph.microsoft.com/v1.0',
    scope: 'openid profile offline_access ' + GRAPH_MAIL_SEND,
    redirectUri: '',
    fetch: null,
    sleep: function (ms) { return new Promise(function (r) { setTimeout(r, ms); }); },
    waitMs: 10 * 60 * 1000
  };
  var session = null; // { token, refresh, expires, account: { name, username } }
  var pending = null;  // the sign-in in progress: { state, verifier, resolve, reject, cleanup }

  function configure(opts) {
    Object.keys(opts || {}).forEach(function (k) { if (opts[k] !== undefined && opts[k] !== null && opts[k] !== '') cfg[k] = opts[k]; });
  }
  function doFetch(url, init) { return (cfg.fetch || fetch)(url, init); }

  // Only on a real web address (not file://) and once the app registration's ID is filled in.
  function available() {
    return !!cfg.clientId && typeof location !== 'undefined' && /^https?:$/.test(location.protocol);
  }
  function redirectUri() {
    return cfg.redirectUri || new URL('auth.html', location.href).href;
  }

  // --- PKCE --------------------------------------------------------------------------------------
  function b64url(bytes) {
    var s = '';
    for (var i = 0; i < bytes.length; i++) s += String.fromCharCode(bytes[i]);
    return btoa(s).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
  }
  function randomString(n) {
    var a = new Uint8Array(n);
    crypto.getRandomValues(a);
    return b64url(a);
  }
  function challengeFor(verifier) {
    return crypto.subtle.digest('SHA-256', new TextEncoder().encode(verifier)).then(function (h) { return b64url(new Uint8Array(h)); });
  }
  function query(obj) {
    return Object.keys(obj).map(function (k) { return encodeURIComponent(k) + '=' + encodeURIComponent(obj[k]); }).join('&');
  }
  function authorizeUrl(p) {
    return cfg.authority + '/oauth2/v2.0/authorize?' + query({
      client_id: cfg.clientId, response_type: 'code', redirect_uri: redirectUri(), response_mode: 'query',
      scope: cfg.scope, state: p.state, code_challenge: p.challenge, code_challenge_method: 'S256',
      prompt: 'select_account'
    });
  }
  // The link a company's IT clicks once when users there may not approve apps themselves.
  function adminConsentUrl() {
    return 'https://login.microsoftonline.com/organizations/v2.0/adminconsent?' + query({ client_id: cfg.clientId, scope: GRAPH_MAIL_SEND, redirect_uri: redirectUri() });
  }

  function decodeJwt(t) {
    try {
      var part = String(t).split('.')[1].replace(/-/g, '+').replace(/_/g, '/');
      while (part.length % 4) part += '=';
      var bin = atob(part);
      var bytes = new Uint8Array(bin.length);
      for (var i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
      return JSON.parse(new TextDecoder().decode(bytes));
    } catch (e) { return {}; }
  }

  function oauthError(json, status) {
    var e = new Error((json && (json.error_description || json.error)) || ('Microsoft answered ' + status));
    e.code = (json && json.error) || 'http_' + status;
    return e;
  }

  function tokenRequest(fields) {
    fields.client_id = cfg.clientId;
    fields.scope = cfg.scope;
    return doFetch(cfg.authority + '/oauth2/v2.0/token', {
      method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body: query(fields)
    }).then(function (res) {
      return res.json().catch(function () { return {}; }).then(function (json) {
        if (!res.ok || !json.access_token) throw oauthError(json, res.status);
        return json;
      });
    });
  }
  function keep(json) {
    var claims = json.id_token ? decodeJwt(json.id_token) : {};
    var acct = session && session.account;
    session = {
      token: json.access_token,
      refresh: json.refresh_token || (session && session.refresh) || '',
      expires: Date.now() + (Number(json.expires_in || 3600) - 120) * 1000,
      account: claims.preferred_username || claims.email ? { name: claims.name || '', username: claims.preferred_username || claims.email } : acct
    };
    return session.account;
  }

  // --- sign-in -----------------------------------------------------------------------------------
  // Opens Microsoft's page in a popup straight from the click (so it is not blocked); the page comes
  // back to auth.html, which hands the answer over by BroadcastChannel (and window.opener).
  function signIn() {
    if (!available()) return Promise.reject(new Error('Sign in with Microsoft is not set up for this copy of the app.'));
    if (pending) pending.reject(Object.assign(new Error('A new sign-in was started.'), { code: 'superseded' }));
    var popup = window.open('', 'payslip-ms-signin', 'width=520,height=700');
    if (!popup) return Promise.reject(Object.assign(new Error('The browser blocked the sign-in window. Allow pop-ups for this page and try again.'), { code: 'popup_blocked' }));
    try { popup.document.title = 'Signing in…'; popup.document.body.innerHTML = '<p style="font:16px sans-serif;padding:24px">Opening Microsoft sign-in…</p>'; } catch (e) { /* already elsewhere */ }
    var state = randomString(18);
    var verifier = randomString(48);
    return new Promise(function (resolve, reject) {
      var channel = typeof BroadcastChannel === 'function' ? new BroadcastChannel('payslip-ms-auth') : null;
      var timer = setTimeout(function () { fail(Object.assign(new Error('The sign-in took too long. Try again.'), { code: 'timeout' })); }, cfg.waitMs);
      function cleanup() {
        clearTimeout(timer);
        if (channel) channel.close();
        window.removeEventListener('message', onMessage);
        window.removeEventListener('storage', onStorage);
        pending = null;
      }
      function fail(err) { cleanup(); try { popup.close(); } catch (e) { /* gone */ } reject(err); }
      var settled = false;
      function onResult(data) {
        if (settled || !data || data.type !== 'payslip-ms-auth' || data.state !== state) return;
        settled = true;
        cleanup();
        if (data.error) { reject(oauthError(data, 400)); return; }
        tokenRequest({ grant_type: 'authorization_code', code: data.code, redirect_uri: redirectUri(), code_verifier: verifier })
          .then(function (json) { resolve(keep(json)); }, reject);
      }
      function onMessage(ev) { if (ev.origin === location.origin) onResult(ev.data); }
      function onStorage(ev) { if (ev.key === 'payslip-ms-auth' && ev.newValue) { try { onResult(JSON.parse(ev.newValue)); } catch (e) { /* not ours */ } } }
      if (channel) channel.onmessage = function (ev) { onResult(ev.data); };
      window.addEventListener('message', onMessage);
      window.addEventListener('storage', onStorage);
      pending = { state: state, reject: fail, cancel: function () { fail(Object.assign(new Error('Sign-in cancelled.'), { code: 'cancelled' })); } };
      challengeFor(verifier).then(function (challenge) {
        popup.location.href = authorizeUrl({ state: state, challenge: challenge });
      }, fail);
    });
  }
  function cancelSignIn() { if (pending) pending.cancel(); }
  function signOut() { session = null; }
  function account() { return session ? session.account : null; }

  function accessToken() {
    if (!session) return Promise.reject(Object.assign(new Error('Not signed in.'), { code: 'signed_out' }));
    if (Date.now() < session.expires || !session.refresh) return Promise.resolve(session.token);
    return refresh().then(function () { return session.token; });
  }
  function refresh() {
    return tokenRequest({ grant_type: 'refresh_token', refresh_token: session.refresh }).then(keep);
  }

  // --- sending -----------------------------------------------------------------------------------
  function base64(bytes) {
    var s = '';
    for (var i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
    return btoa(s);
  }
  function mailBody(m) {
    return JSON.stringify({
      message: {
        subject: m.subject,
        body: { contentType: 'Text', content: m.text },
        toRecipients: [{ emailAddress: { address: m.to } }],
        attachments: [{ '@odata.type': '#microsoft.graph.fileAttachment', name: m.fileName, contentType: 'application/pdf', contentBytes: base64(m.bytes) }]
      },
      saveToSentItems: true
    });
  }

  // Sends one email with its PDF from the signed-in mailbox (a copy lands in Sent Items). Waits and
  // retries when Microsoft asks to slow down; signs in again quietly when the token has expired.
  function sendMail(m) {
    var body = mailBody(m);
    var attempt = function (n, refreshed) {
      return accessToken().then(function (token) {
        return doFetch(cfg.graph + '/me/sendMail', {
          method: 'POST', headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' }, body: body
        });
      }).then(function (res) {
        if (res.ok) return true;
        if ((res.status === 429 || res.status === 503 || res.status === 504) && n < 4) {
          var wait = Number(res.headers && res.headers.get && res.headers.get('Retry-After')) || 20 * n;
          if (m.onWait) m.onWait(wait);
          return cfg.sleep(wait * 1000).then(function () { return attempt(n + 1, refreshed); });
        }
        if (res.status === 401 && !refreshed && session && session.refresh) {
          return refresh().then(function () { return attempt(n, true); });
        }
        return res.json().catch(function () { return {}; }).then(function (json) {
          var err = json && json.error ? json.error : {};
          var e = new Error((err.code ? err.code + ': ' : '') + (err.message || 'Microsoft answered ' + res.status));
          e.code = err.code || 'http_' + res.status;
          throw e;
        });
      });
    };
    return attempt(1, false);
  }

  // Plain words for what went wrong, and what to do.
  function hint(err) {
    var t = String((err && err.code) || '') + ' ' + String((err && err.message) || err || '');
    if (/popup_blocked/.test(t)) return 'The browser blocked the sign-in window. Allow pop-ups for this page (the icon at the end of the address bar), then try again.';
    if (/AADSTS65001|AADSTS90094|AADSTS90095|admin(istrator)? (approval|consent)|consent_required/i.test(t)) return 'Your company lets only its IT approve new apps. Send your IT the approval link; after they approve once, everyone there can sign in.';
    if (/AADSTS53003|AADSTS50105|AADSTS530|Conditional Access|blocked/i.test(t)) return 'Your company\'s sign-in rules block this sign-in. Ask IT, or use the send-all pack or one email at a time.';
    if (/access_denied|AADSTS65004|declined|cancelled/i.test(t)) return 'The sign-in was cancelled. Try again and accept "Send mail as you".';
    if (/MailboxNotEnabledForRESTAPI|MailboxNotHostedInExchangeOnline|ResourceNotFound/i.test(t)) return 'This mailbox is not on Microsoft\'s online service, so it cannot send from here. Use the send-all pack or one email at a time.';
    if (/ErrorSendAsDenied|ErrorAccessDenied|Forbidden|http_403/i.test(t)) return 'Microsoft refused to send from this mailbox. Sign in with the payroll email address.';
    if (/ErrorInvalidRecipients|InvalidRecipients/i.test(t)) return 'Microsoft did not accept this email address. Check it in the Excel file.';
    if (/QuotaExceeded|ErrorQuotaExceeded|SubmissionQuotaExceeded|daily/i.test(t)) return 'The mailbox reached its daily sending limit. Continue tomorrow: everyone sent is already ticked.';
    if (/Failed to fetch|NetworkError|network|offline/i.test(t)) return 'Microsoft could not be reached. Check the internet connection and try again.';
    if (/signed_out|invalid_grant|interaction_required|AADSTS50173|AADSTS700082/i.test(t)) return 'The sign-in has ended. Sign in with Microsoft again.';
    return '';
  }
  // Errors after which trying the next person is pointless.
  function stopsRun(err) {
    return !/ErrorInvalidRecipients|InvalidRecipients/i.test(String((err && err.code) || '') + ' ' + String((err && err.message) || ''));
  }

  return {
    configure: configure,
    available: available,
    redirectUri: redirectUri,
    authorizeUrl: authorizeUrl,
    adminConsentUrl: adminConsentUrl,
    signIn: signIn,
    cancelSignIn: cancelSignIn,
    signOut: signOut,
    account: account,
    sendMail: sendMail,
    hint: hint,
    stopsRun: stopsRun,
    _test: { keep: keep, decodeJwt: decodeJwt, mailBody: mailBody, challengeFor: challengeFor, setSession: function (s) { session = s; } }
  };
});
