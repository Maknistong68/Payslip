// Test helper: stands in for Microsoft sign-in (device code for the packs; authorization code with PKCE
// for the page) and Microsoft Graph sendMail. Writes its port to <dir>/port, each request to
// <dir>/requests.jsonl and each sent message to <dir>/<n>.json.
// mode: ok (sign-in pending once, then a token; the first sendMail is throttled once), deny, admin.
// Usage: node tools/fake-graph.js <dir> [mode] [fail-for-address]
'use strict';
var http = require('http');
var crypto = require('crypto');
var fs = require('fs');
var path = require('path');
var qs = require('querystring');

var dir = process.argv[2];
var mode = process.argv[3] || 'ok';
var failFor = process.argv[4] || '';
var polls = 0;
var sends = 0;
var throttled = false;
var challenges = {};   // browser sign-in: code -> PKCE challenge

var CORS = { 'Access-Control-Allow-Origin': '*', 'Access-Control-Allow-Headers': 'Authorization, Content-Type', 'Access-Control-Expose-Headers': 'Retry-After' };
function reply(res, status, obj, headers) {
  res.writeHead(status, Object.assign({ 'Content-Type': 'application/json' }, CORS, headers || {}));
  res.end(obj === undefined ? '' : JSON.stringify(obj));
}

http.createServer(function (req, res) {
  var chunks = [];
  req.on('data', function (c) { chunks.push(c); });
  req.on('end', function () {
    var raw = Buffer.concat(chunks).toString('utf8');
    var form = /x-www-form-urlencoded/.test(req.headers['content-type'] || '') ? qs.parse(raw) : null;
    fs.appendFileSync(path.join(dir, 'requests.jsonl'), JSON.stringify({ method: req.method, url: req.url, form: form, auth: req.headers.authorization || '' }) + '\n');
    if (req.method === 'OPTIONS') { res.writeHead(204, Object.assign({ 'Access-Control-Allow-Methods': 'POST, GET' }, CORS)); return res.end(); }
    // browser sign-in (authorization code + PKCE): Microsoft's page "signs in" at once and goes back
    var az = req.url.match(/^\/(organizations|consumers|common)\/oauth2\/v2\.0\/authorize\?(.*)$/);
    if (az) {
      var q = qs.parse(az[2]);
      var back = q.redirect_uri + '?' + (mode === 'deny' ? qs.stringify({ error: 'access_denied', error_description: 'AADSTS65004: User declined to consent to access the app.', state: q.state })
        : mode === 'admin' ? qs.stringify({ error: 'access_denied', error_description: 'AADSTS90094: An administrator of Contoso has set a policy that prevents you from granting this app the permissions. admin approval needed.', state: q.state })
        : qs.stringify({ code: 'code-' + (++polls), state: q.state }));
      if (mode !== 'deny' && mode !== 'admin') challenges['code-' + polls] = q.code_challenge;
      res.writeHead(302, { Location: back });
      return res.end();
    }
    var m = req.url.match(/^\/(organizations|consumers|common)\/oauth2\/v2\.0\/(devicecode|token)$/);
    if (m && m[2] === 'token' && form.grant_type === 'authorization_code') {
      var want = challenges[form.code];
      var got = crypto.createHash('sha256').update(form.code_verifier || '').digest('base64url');
      if (!want || want !== got) return reply(res, 400, { error: 'invalid_grant', error_description: 'AADSTS501481: The Code_Verifier does not match the code_challenge.' });
      delete challenges[form.code];
      var idt = 'h.' + Buffer.from(JSON.stringify({ name: 'Payroll Team', preferred_username: 'payroll@contoso.com' })).toString('base64url') + '.s';
      return reply(res, 200, { access_token: 'tok-1', refresh_token: 'ref-1', expires_in: 3600, id_token: idt });
    }
    if (m && m[2] === 'devicecode') {
      return reply(res, 200, { device_code: 'dc-1', user_code: 'ABCD-EFGH', verification_uri: 'https://microsoft.com/devicelogin', expires_in: 900, interval: 5 });
    }
    if (m && m[2] === 'token') {
      if (form.grant_type === 'refresh_token') return reply(res, 200, { access_token: 'tok-2', refresh_token: 'ref-2', expires_in: 3600 });
      if (polls++ === 0) return reply(res, 400, { error: 'authorization_pending', error_description: 'AADSTS70016: pending' });
      if (mode === 'deny') return reply(res, 400, { error: 'access_denied', error_description: 'AADSTS65004: User declined to consent to access the app.' });
      if (mode === 'admin') return reply(res, 400, { error: 'invalid_grant', error_description: 'AADSTS65001: The user or administrator has not consented to use the application.' });
      return reply(res, 200, { access_token: 'tok-1', refresh_token: 'ref-1', expires_in: 3600 });
    }
    if (req.url === '/v1.0/me/sendMail' && req.method === 'POST') {
      if (!/^Bearer tok-/.test(req.headers.authorization || '')) return reply(res, 401, { error: { code: 'InvalidAuthenticationToken', message: 'Access token is empty.' } });
      if (!throttled) { throttled = true; return reply(res, 429, { error: { code: 'ApplicationThrottled', message: 'Too many requests' } }, { 'Retry-After': '1' }); }
      var body = JSON.parse(raw);
      var to = body.message.toRecipients[0].emailAddress.address;
      if (failFor && to === failFor) return reply(res, 400, { error: { code: 'ErrorInvalidRecipients', message: 'At least one recipient is not valid.' } });
      sends++;
      fs.writeFileSync(path.join(dir, sends + '.json'), raw);
      return reply(res, 202);
    }
    reply(res, 404, { error: { code: 'NotFound', message: 'no such endpoint ' + req.url } });
  });
}).listen(0, '127.0.0.1', function () { fs.writeFileSync(path.join(dir, 'port'), String(this.address().port)); });
