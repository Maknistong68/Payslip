// Test helper: a tiny mail server that accepts mail only from a signed-in client, like Microsoft 365
// (AUTH LOGIN/PLAIN with any password but the refuse-password). Writes each message to <dir>/<n>.eml and
// its port to <dir>/port, then runs until killed. Usage: node tools/fake-smtp.js <dir> [refuse-password]
'use strict';
var net = require('net');
var fs = require('fs');
var path = require('path');

var dir = process.argv[2];
var refuse = process.argv[3] || '';
var count = 0;

net.createServer(function (sock) {
  var buf = '';
  var inData = false;
  var authStep = '';
  var user = '';
  var msg = { from: '', to: [], user: '' };
  var say = function (line) { sock.write(line + '\r\n'); };
  var plainAuth = function (b64) {
    var plain = Buffer.from(b64, 'base64').toString('utf8').split('\0');
    if (refuse && plain[2] === refuse) say('535 5.7.8 Username and Password not accepted');
    else { msg.user = plain[1]; say('235 2.7.0 Authentication successful'); }
  };
  say('220 fake-smtp ready');
  sock.on('data', function (chunk) {
    buf += chunk.toString('latin1');
    for (;;) {
      if (inData) {
        var end = buf.indexOf('\r\n.\r\n');
        if (end < 0) return;
        var data = buf.slice(0, end);
        buf = buf.slice(end + 5);
        inData = false;
        count++;
        var head = 'X-Envelope-From: ' + msg.from + '\r\nX-Envelope-To: ' + msg.to.join(',') + '\r\nX-Auth-User: ' + msg.user + '\r\n';
        fs.writeFileSync(path.join(dir, count + '.eml'), Buffer.from(head + data.replace(/\r\n\.\./g, '\r\n.'), 'latin1'));
        msg = { from: '', to: [], user: msg.user };
        say('250 queued as ' + count);
        continue;
      }
      var nl = buf.indexOf('\r\n');
      if (nl < 0) return;
      var line = buf.slice(0, nl);
      buf = buf.slice(nl + 2);
      var decode = function (b64) { return Buffer.from(b64, 'base64').toString('utf8'); };
      if (authStep === 'plain') { authStep = ''; plainAuth(line); continue; }
      if (authStep === 'user') { user = decode(line); authStep = 'pass'; say('334 UGFzc3dvcmQ6'); continue; }
      if (authStep === 'pass') {
        authStep = '';
        if (refuse && decode(line) === refuse) { say('535 5.7.8 Username and Password not accepted'); continue; }
        msg.user = user; say('235 2.7.0 Authentication successful'); continue;
      }
      var cmd = line.slice(0, 4).toUpperCase();
      if (cmd === 'EHLO') { say('250-fake-smtp'); say('250-AUTH LOGIN PLAIN'); say('250 8BITMIME'); }
      else if (cmd === 'HELO') say('250 fake-smtp');
      else if (cmd === 'AUTH') {
        var parts = line.split(' ');
        if (parts[1].toUpperCase() === 'LOGIN') {
          if (parts[2]) { user = decode(parts[2]); authStep = 'pass'; say('334 UGFzc3dvcmQ6'); } else { authStep = 'user'; say('334 VXNlcm5hbWU6'); }
        } else if (parts[2]) plainAuth(parts[2]);
        else { authStep = 'plain'; say('334 '); }
      }
      else if (cmd === 'MAIL' && !msg.user) say('530 5.7.57 Client not authenticated to send mail');
      else if (cmd === 'MAIL') { msg.from = line.replace(/^MAIL FROM:\s*/i, ''); say('250 OK'); }
      else if (cmd === 'RCPT') { msg.to.push(line.replace(/^RCPT TO:\s*/i, '')); say('250 OK'); }
      else if (cmd === 'DATA') { inData = true; say('354 go ahead'); }
      else if (cmd === 'QUIT') { say('221 bye'); sock.end(); return; }
      else say('250 OK');
    }
  });
  sock.on('error', function () {});
}).listen(0, '127.0.0.1', function () { fs.writeFileSync(path.join(dir, 'port'), String(this.address().port)); });
