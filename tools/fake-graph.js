// Test helper: stands in for Microsoft sign-in (device code) and Microsoft Graph sendMail. Writes its
// port to <dir>/port, each request to <dir>/requests.jsonl and each sent message to <dir>/<n>.json.
// mode: ok (sign-in pending once, then a token; the first sendMail is throttled once), deny, admin.
// Usage: node tools/fake-graph.js <dir> [mode] [fail-for-address]
'use strict';
var http = require('http');
var fs = require('fs');
var path = require('path');
var qs = require('querystring');

var dir = process.argv[2];
var mode = process.argv[3] || 'ok';
var failFor = process.argv[4] || '';
var polls = 0;
var sends = 0;
var throttled = false;

function reply(res, status, obj) {
  res.writeHead(status, { 'Content-Type': 'application/json' });
  res.end(obj === undefined ? '' : JSON.stringify(obj));
}

http.createServer(function (req, res) {
  var chunks = [];
  req.on('data', function (c) { chunks.push(c); });
  req.on('end', function () {
    var raw = Buffer.concat(chunks).toString('utf8');
    var form = /x-www-form-urlencoded/.test(req.headers['content-type'] || '') ? qs.parse(raw) : null;
    fs.appendFileSync(path.join(dir, 'requests.jsonl'), JSON.stringify({ method: req.method, url: req.url, form: form, auth: req.headers.authorization || '' }) + '\n');
    var m = req.url.match(/^\/(organizations|consumers|common)\/oauth2\/v2\.0\/(devicecode|token)$/);
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
      if (!throttled) { throttled = true; return reply(res, 429, { error: { code: 'ApplicationThrottled', message: 'Too many requests' } }); }
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
