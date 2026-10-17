/*
 * Payslip Generator - delivery helpers: .eml drafts, mail-merge pack, WhatsApp / Gmail links,
 * Outlook VBA pack. Pure functions (no DOM) so they run in the browser and in Node tests.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory(require('./schema.js'));
  } else {
    root.PayslipDelivery = factory(root.PayslipSchema);
  }
})(typeof self !== 'undefined' ? self : this, function (Schema) {
  'use strict';

  var CRLF = '\r\n';

  // ---- encoding helpers ------------------------------------------------------

  function utf8Bytes(str) {
    if (typeof TextEncoder !== 'undefined') return new TextEncoder().encode(str);
    return new Uint8Array(Buffer.from(str, 'utf8'));
  }

  function bytesToBase64(bytes) {
    if (typeof Buffer !== 'undefined' && Buffer.from) return Buffer.from(bytes).toString('base64');
    var s = '';
    for (var i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
    return btoa(s);
  }

  // base64 folded at 76 characters per line (RFC 2045)
  function base64Lines(bytes) {
    var b64 = bytesToBase64(bytes);
    var out = [];
    for (var i = 0; i < b64.length; i += 76) out.push(b64.slice(i, i + 76));
    return out.join(CRLF);
  }

  function isAscii(s) {
    return /^[\x20-\x7e]*$/.test(s);
  }

  // RFC 2047 encoded-words for headers with non-ASCII text (Arabic subjects).
  // Each encoded word stays <= 75 chars and never splits a UTF-8 character.
  function encodeHeader(value) {
    value = String(value == null ? '' : value).replace(/[\r\n]+/g, ' ');
    if (isAscii(value)) return value;
    var words = [];
    var chunk = '';
    var chunkBytes = 0;
    var chars = Array.from(value);
    for (var i = 0; i < chars.length; i++) {
      var ch = chars[i];
      var len = utf8Bytes(ch).length;
      if (chunkBytes + len > 45 && chunk) { // 45 raw bytes -> 60 base64 chars -> 72 with wrapper
        words.push('=?UTF-8?B?' + bytesToBase64(utf8Bytes(chunk)) + '?=');
        chunk = '';
        chunkBytes = 0;
      }
      chunk += ch;
      chunkBytes += len;
    }
    if (chunk) words.push('=?UTF-8?B?' + bytesToBase64(utf8Bytes(chunk)) + '?=');
    return words.join(CRLF + ' ');
  }

  function rfc2822Date(d) {
    return (d || new Date()).toUTCString().replace('GMT', '+0000');
  }

  function randomToken() {
    var s = '';
    for (var i = 0; i < 16; i++) s += Math.floor(Math.random() * 16).toString(16);
    return s;
  }

  // ---- .eml draft ------------------------------------------------------------

  /**
   * Builds an RFC 5322 message with the PDF attached. Classic Outlook for Windows opens it as an
   * editable draft because of X-Unsent: 1 (other clients open it read-only).
   * opts: { from, to, subject, text, fileName, pdfBytes (Uint8Array), messageId }
   */
  // Plain text -> simple HTML: blank line = new paragraph, single newline = <br>. dir="auto" keeps Arabic right-to-left.
  function textToHtml(text) {
    var paras = String(text || '').replace(/\r\n/g, '\n').split(/\n{2,}/);
    return '<!DOCTYPE html><html><head><meta charset="UTF-8"></head><body style="font-family:Segoe UI,Arial,sans-serif;font-size:14px;line-height:1.5;color:#1f2933">' +
      paras.map(function (p) {
        return '<p dir="auto" style="margin:0 0 14px">' + esc(p).replace(/\n/g, '<br>') + '</p>';
      }).join('') + '</body></html>';
  }

  function buildEml(opts) {
    var boundary = '----=_Payslip_' + randomToken();
    var messageId = opts.messageId || ('<payslip-' + Schema.asciiSlug(opts.fileName || 'slip').toLowerCase() + '-' + randomToken() + '@payslip.local>');
    var fileName = opts.fileName || 'payslip.pdf';
    var lines = [
      'X-Unsent: 1',
      'Message-ID: ' + messageId,
      'Date: ' + rfc2822Date(),
      'MIME-Version: 1.0'
    ];
    if (opts.from) lines.push('From: ' + opts.from);
    if (opts.to) lines.push('To: ' + opts.to);
    lines.push('Subject: ' + encodeHeader(opts.subject || ''));
    lines.push('Content-Type: multipart/mixed; boundary="' + boundary + '"');
    lines.push('');
    lines.push('This is a multi-part message in MIME format.');
    lines.push('');
    // HTML body only: Outlook's "remove extra line breaks" flattens plain-text drafts, and when both
    // parts are present it may still build the draft from the plain one.
    lines.push('--' + boundary);
    lines.push('Content-Type: text/html; charset="UTF-8"');
    lines.push('Content-Transfer-Encoding: base64');
    lines.push('');
    lines.push(base64Lines(utf8Bytes(textToHtml(opts.text))));
    lines.push('');
    lines.push('--' + boundary);
    lines.push('Content-Type: application/pdf; name="' + fileName + '"');
    lines.push('Content-Transfer-Encoding: base64');
    lines.push('Content-Disposition: attachment; filename="' + fileName + '"');
    lines.push('');
    lines.push(base64Lines(opts.pdfBytes));
    lines.push('');
    lines.push('--' + boundary + '--');
    lines.push('');
    return lines.join(CRLF);
  }

  // ---- links -----------------------------------------------------------------

  function waLink(phoneDigits, text) {
    var p = String(phoneDigits || '').replace(/\D/g, '');
    return 'https://wa.me/' + p + (text ? '?text=' + encodeURIComponent(text) : '');
  }

  function gmailComposeLink(to, subject, body) {
    return 'https://mail.google.com/mail/?view=cm&fs=1&to=' + encodeURIComponent(to || '') +
      '&su=' + encodeURIComponent(subject || '') + '&body=' + encodeURIComponent(body || '');
  }

  function mailtoLink(to, subject, body) {
    return 'mailto:' + encodeURIComponent(to || '') + '?subject=' + encodeURIComponent(subject || '') + '&body=' + encodeURIComponent(body || '');
  }

  // Outlook on the web (work or school accounts): opens the compose window with everything typed.
  function outlookWebLink(to, subject, body) {
    return 'https://outlook.office.com/mail/deeplink/compose?to=' + encodeURIComponent(to || '') + '&subject=' + encodeURIComponent(subject || '') + '&body=' + encodeURIComponent(body || '');
  }

  // ---- CSV -------------------------------------------------------------------

  function csvCell(v) {
    var s = String(v == null ? '' : v);
    if (/[",\r\n]/.test(s)) s = '"' + s.replace(/"/g, '""') + '"';
    return s;
  }

  function toCsv(rows) {
    return '\ufeff' + rows.map(function (r) { return r.map(csvCell).join(','); }).join(CRLF) + CRLF;
  }

  /**
   * manifest rows: one per slip. items: [{ slip, fileName, hash, password }]
   */
  function manifestCsv(items, settings) {
    var rows = [['employee_id', 'name', 'email', 'mobile', 'file', 'sha256', 'net', 'currency', 'period', 'password_hint']];
    items.forEach(function (it) {
      var pw = it.password || {};
      rows.push([it.slip.employeeId, it.slip.name, it.slip.fields.email || '', it.slip.fields.mobile || '', it.fileName, it.hash || '', Schema.formatMoney(it.slip.net), settings.currency || 'SAR', settings.payPeriod || '', pw.hint || '']);
    });
    return toCsv(rows);
  }

  /** Thunderbird Mail Merge CSV: to, subject, body, attachment (file name only), name, employee_id */
  function mailMergeCsv(items, settings) {
    var rows = [['to', 'name', 'employee_id', 'subject', 'body', 'attachment']];
    items.forEach(function (it) {
      if (!it.slip.fields.email) return;
      rows.push([it.slip.fields.email, it.slip.name, it.slip.employeeId, it.subject, it.body, it.fileName]);
    });
    return toCsv(rows);
  }

  function deliveryLogCsv(entries) {
    // action: sent, undo (tick removed) or clear (all ticks removed); source: which Excel file was loaded
    var rows = [['timestamp', 'employee_id', 'name', 'channel', 'contact', 'file', 'sha256', 'note', 'action', 'source']];
    entries.forEach(function (e) {
      rows.push([e.timestamp, e.employeeId, e.name || '', e.channel || '', e.contact || '', e.file || '', e.hash || '', e.note || '', e.action || 'sent', e.source || '']);
    });
    return toCsv(rows);
  }

  // ---- packs: guides ---------------------------------------------------------

  function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; });
  }

  function mailMergeGuideHtml(settings, counts) {
    return '<!doctype html><html lang="en"><head><meta charset="utf-8"><title>How to send the payslips with Thunderbird Mail Merge</title>' +
      '<style>body{font-family:Segoe UI,Arial,sans-serif;max-width:760px;margin:40px auto;line-height:1.55;color:#1f2933;padding:0 16px}h1{font-size:1.5rem}h2{font-size:1.1rem;margin-top:28px}code,pre{background:#eef2f6;border-radius:4px;padding:2px 6px}pre{padding:12px;white-space:pre-wrap}li{margin:6px 0}.note{background:#fff6d6;border:1px solid #f0dc9a;border-radius:8px;padding:10px 14px}</style></head><body>' +
      '<h1>Send ' + counts.withEmail + ' payslips by email in one run</h1>' +
      '<p>' + esc(settings.companyName || '') + ' · ' + esc(settings.payPeriod || '') + '. This pack contains the PDF files, <code>mailmerge.csv</code> with one row per employee, and this guide. It works with Gmail, Microsoft 365 / Outlook.com, Zoho or any mailbox, using the free Thunderbird email program. Setup is done once; next month you only repeat step 4.</p>' +
      (counts.withoutEmail ? '<p class="note">' + counts.withoutEmail + ' employee(s) have no email address in the Excel file and are not in <code>mailmerge.csv</code>. Send theirs by WhatsApp or on paper.</p>' : '') +
      '<h2>1. Install Thunderbird (once)</h2><ol><li>Download Thunderbird from <code>thunderbird.net</code> and install it.</li><li>On first start, enter the company mailbox address (for example <code>' + esc(settings.senderEmail || 'payroll@yourcompany.com') + '</code>) and sign in. Gmail and Microsoft 365 open their normal sign-in window; no special password is needed.</li></ol>' +
      '<h2>2. Install the Mail Merge add-on (once)</h2><ol><li>In Thunderbird open <b>Tools → Add-ons and Themes</b>.</li><li>Search for <b>Mail Merge</b> (by Alexander Bergmann) and click <b>Add to Thunderbird</b>.</li></ol>' +
      '<h2>3. Unzip this pack</h2><p>Extract the ZIP to a folder, for example <code>Documents\\Payslips\\' + esc(Schema.periodSlug(settings.payPeriod, settings.periodStart)) + '</code>. Keep the PDF files and <code>mailmerge.csv</code> together.</p>' +
      '<h2>4. Send</h2><ol>' +
      '<li>Click <b>Write</b> to open a new message.</li>' +
      '<li>In <b>To</b> type exactly: <code>{{to}}</code></li>' +
      '<li>In <b>Subject</b> type exactly: <code>{{subject}}</code></li>' +
      '<li>In the body type exactly: <code>{{body}}</code></li>' +
      '<li>Open <b>File → Mail Merge</b>. Choose <b>Source: CSV</b>, click <b>Browse</b> and pick <code>mailmerge.csv</code>. Set <b>Character set</b> UTF-8, <b>Field delimiter</b> comma.</li>' +
      '<li>In <b>Attachments</b> type <code>{{attachment}}</code>, then use the Browse button next to it and select <b>all the PDF files</b> in the folder (Ctrl+A). Thunderbird needs you to pre-select the files once; the column tells it which file goes to whom.</li>' +
      '<li>Set <b>Deliver mode: Send Later</b> for a first test, or <b>Send Now</b>. Set <b>Pause: 3</b> seconds between messages.</li>' +
      '<li>Click <b>OK</b>. With Send Later the messages wait in the Outbox: open two or three to check them, then <b>File → Send Unsent Messages</b>.</li></ol>' +
      '<h2>Sending limits of the mailbox</h2><ul><li>Microsoft 365 business mailbox: about 30 messages per minute, 10,000 recipients per day. ' + counts.withEmail + ' messages take at least ' + Math.ceil(counts.withEmail / 30) + ' minutes.</li><li>Free Gmail: 500 messages per day. Google Workspace: 2,000 per day.</li><li>A tenant still on the default <code>onmicrosoft.com</code> address is limited to 100 external recipients per day; use the company domain.</li></ul>' +
      '<h2>Privacy</h2><p>Every PDF is password protected with the rule shown in the message. Do not send the password in the same message. Keep <code>mailmerge.csv</code> and the PDFs in a folder only payroll staff can open and delete them after sending.</p>' +
      '</body></html>';
  }

  // ---- "Send all" packs: one double-click sends every email from the user's own mailbox ----------
  // items: [{ employeeId, name, email, subject, body, file, source }]. Both packs ask for YES first,
  // pause between messages (Microsoft 365 allows about 30 a minute), skip anyone already sent when run
  // again, and write delivery-log.csv in the format "Restore ticks from a log" reads.

  function psString(v) { return "'" + String(v == null ? '' : v).replace(/'/g, "''") + "'"; }
  function shString(v) { return "'" + String(v == null ? '' : v).replace(/'/g, "'\\''") + "'"; }

  function sendListCsv(items) {
    var rows = [['employee_id', 'name', 'email', 'subject', 'body', 'file', 'source']];
    items.forEach(function (it) {
      rows.push([it.employeeId, it.name, it.email, it.subject, String(it.body || '').replace(/\r?\n/g, '\r\n'), it.file, it.source || '']);
    });
    return toCsv(rows);
  }

  // Tab separated, one line per employee: line breaks in the message become the two characters \n.
  function sendListTsv(items) {
    var clean = function (v) { return String(v == null ? '' : v).replace(/\t/g, ' ').replace(/\r?\n/g, '\\n'); };
    return ['employee_id\tname\temail\tsubject\tbody\tfile\tsource'].concat(items.map(function (it) {
      return [it.employeeId, it.name, it.email, it.subject, it.body, it.file, it.source || ''].map(clean).join('\t');
    })).join('\n') + '\n';
  }

  // Windows: PowerShell (built in, no installs) sending through classic Outlook on this computer, or
  // without Outlook through a Microsoft sign-in or the email provider's mail server.
  function windowsSendScript(info) {
    return '\uFEFF' + [
      '# Payslip Generator - sends every payslip email from this computer, whatever email it uses:',
      '# classic Outlook if it answers, otherwise a Microsoft sign-in (Microsoft 365, Outlook.com, New Outlook)',
      '# or your email provider\'s mail server (Gmail, Zoho, company mail).',
      '# Start it with "2 - Send all.bat" (try "1 - Send test to me.bat" first).',
      'param(',
      '  [switch]$Test,          # one email to yourself with the first payslip; nothing is logged',
      '  [switch]$Drafts,        # save every email in Outlook Drafts instead of sending (classic Outlook only)',
      '  [switch]$Check,         # check the pack and how it will send; nothing is sent',
      '  [ValidateSet(\'\', \'auto\', \'outlook\', \'microsoft\', \'server\')]',
      '  [string]$Route = \'\',    # outlook, microsoft (sign-in), server (mail server); auto = by address, skip Outlook',
      '  [switch]$Smtp,          # same as -Route server',
      '  [string]$From = \'\',     # your email address',
      '  [string]$Server = \'\',   # mail server: host or host:port',
      '  [switch]$NoTls,         # mail server without encryption (local test servers only)',
      '  [switch]$Yes,           # no questions (for automated checks)',
      '  [double]$Delay = 2.5,',
      '  $Outlook = $null',
      ')',
      '$ErrorActionPreference = \'Stop\'',
      '$Company = ' + psString(info.company),
      '$Period = ' + psString(info.period),
      '$Sender = ' + psString(info.sender),
      '$here = Split-Path -Parent $MyInvocation.MyCommand.Path',
      '$self = $MyInvocation.MyCommand.Path',
      '$bound = $PSBoundParameters',
      '$listPath = Join-Path $here \'send-list.csv\'',
      '$logPath = Join-Path $here \'delivery-log.csv\'',
      '$errPath = Join-Path $here \'errors.csv\'',
      '$pdfDir = Join-Path $here \'pdf\'',
      '# Microsoft sign-in: Microsoft\'s own public app for command-line access to your mailbox (the one',
      '# Microsoft Graph PowerShell uses); it only gets "send mail as you", after you agree on Microsoft\'s page.',
      '$MsClient = \'14d82eec-204b-4c2f-b7e8-296a70dab67e\'',
      '$MsScope = \'https://graph.microsoft.com/Mail.Send offline_access\'',
      '$AuthBase = if ($env:PAYSLIP_AUTH_BASE) { $env:PAYSLIP_AUTH_BASE } else { \'https://login.microsoftonline.com\' }',
      '$GraphBase = if ($env:PAYSLIP_GRAPH_BASE) { $env:PAYSLIP_GRAPH_BASE } else { \'https://graph.microsoft.com/v1.0\' }',
      '',
      'function Say($text, $color) { if ($color) { Write-Host $text -ForegroundColor $color } else { Write-Host $text } }',
      'function Finish { if (-not $Yes) { Write-Host \'\'; [void](Read-Host \'Press Enter to close\') } }',
      'function CsvCell($v) { $s = [string]$v; if ($s -match \'[",\\r\\n]\') { return \'"\' + $s.Replace(\'"\', \'""\') + \'"\' } return $s }',
      'function CsvLine($values) { return (($values | ForEach-Object { CsvCell $_ }) -join \',\') }',
      '',
      'try { [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12 } catch { }',
      '',
      '# --- classic Outlook: is it installed, is New Outlook here, can we talk to it? ---------------------',
      'function Find-Outlook {',
      '  $res = @{ ok = $false; app = $null; installed = $false; newOutlook = $false; error = \'\' }',
      '  try { if (Get-Process -Name \'olk\' -ErrorAction SilentlyContinue) { $res.newOutlook = $true } } catch { }',
      '  if (-not $res.newOutlook) { try { if (Get-AppxPackage -Name \'Microsoft.OutlookForWindows\' -ErrorAction SilentlyContinue) { $res.newOutlook = $true } } catch { } }',
      '  foreach ($key in @(\'Registry::HKEY_CLASSES_ROOT\\Outlook.Application\', \'HKLM:\\SOFTWARE\\Classes\\Outlook.Application\', \'HKLM:\\SOFTWARE\\WOW6432Node\\Classes\\Outlook.Application\')) {',
      '    try { if (Test-Path $key) { $res.installed = $true } } catch { }',
      '  }',
      '  # always try: the Outlook that is already open first, otherwise start it',
      '  try { $res.app = [Runtime.InteropServices.Marshal]::GetActiveObject(\'Outlook.Application\') } catch { }',
      '  if (-not $res.app) { try { $res.app = New-Object -ComObject Outlook.Application } catch { $res.error = $_.Exception.Message } }',
      '  if ($res.app) { $res.installed = $true; $res.ok = $true }',
      '  return $res',
      '}',
      '',
      'function Explain-Outlook($o) {',
      '  if ($o.error -match \'80080005\') {',
      '    Say \'Classic Outlook is installed but would not answer: Outlook or this window runs "as administrator" and the other does not.\' Yellow',
      '    Say \'  To use Outlook: close it completely, open it normally (not as administrator), and run this again.\' Yellow',
      '  } elseif ($o.newOutlook -and -not $o.installed) {',
      '    Say \'This computer uses the NEW Outlook, which Microsoft does not let scripts control.\' Yellow',
      '  } elseif ($o.installed) {',
      '    Say \'Classic Outlook is installed but would not answer.\' Yellow',
      '    if ($o.newOutlook) { Say \'  To use it: open CLASSIC Outlook (turn off the "New Outlook" switch), wait until it has loaded, and run this again.\' Yellow }',
      '    elseif ($o.error -match \'80040154|not registered\') { Say \'  Office may need a repair: Windows Settings > Apps > Microsoft 365 / Office > Modify > Quick Repair.\' Yellow }',
      '    else { Say \'  To use it: open Outlook, wait until it has loaded, and run this again.\' Yellow }',
      '  } else {',
      '    Say \'Classic Outlook is not on this computer.\' Yellow',
      '  }',
      '  if ($o.error) { Say (\'  (Windows said: \' + $o.error + \')\') DarkGray }',
      '}',
      '',
      '# 32-bit Office on 64-bit Windows sometimes answers only 32-bit PowerShell: try once more there.',
      '$script:relaunchCode = 0',
      'function Use-32BitPowerShell($o) {',
      '  if ($o.ok -or $env:PAYSLIP_32BIT -or $o.error -notmatch \'80040154|not registered\') { return $false }',
      '  if (-not [Environment]::Is64BitProcess -or -not $env:WINDIR) { return $false }',
      '  $ps32 = Join-Path $env:WINDIR \'SysWOW64\\WindowsPowerShell\\v1.0\\powershell.exe\'',
      '  if (-not (Test-Path $ps32)) { return $false }',
      '  Say \'Trying again with 32-bit PowerShell (for 32-bit Office)...\' DarkGray',
      '  $argList = @(\'-NoProfile\', \'-ExecutionPolicy\', \'Bypass\', \'-File\', $self)',
      '  foreach ($k in $bound.Keys) {',
      '    $v = $bound[$k]',
      '    if ($k -eq \'Outlook\') { continue }',
      '    if ($v -is [System.Management.Automation.SwitchParameter]) { if ($v.IsPresent) { $argList += \'-\' + $k } }',
      '    else { $argList += \'-\' + $k; $argList += [string]$v }',
      '  }',
      '  $env:PAYSLIP_32BIT = \'1\'',
      '  # same window, its own input and output (so its messages and questions show as they happen)',
      '  $line = ($argList | ForEach-Object { \'"\' + ([string]$_).Replace(\'"\', \'\\"\') + \'"\' }) -join \' \'',
      '  $p = Start-Process -FilePath $ps32 -ArgumentList $line -NoNewWindow -Wait -PassThru',
      '  $script:relaunchCode = $p.ExitCode',
      '  return $true',
      '}',
      '',
      'function Outlook-Hint($message) {',
      '  if ($message -match \'80004004|Operation aborted|programmatic|not trusted|denied\') { return \'Outlook blocked this program from sending. In Outlook: File > Options > Trust Center > Trust Center Settings > Programmatic Access. If it says the antivirus is not up to date, update it (or ask IT), then run this again.\' }',
      '  return \'\'',
      '}',
      '',
      '# --- who hosts your email? Decides how to send without Outlook --------------------------------------',
      'function Find-Provider($address) {',
      '  $domain = ($address -split \'@\')[-1].Trim().ToLower()',
      '  $res = @{ kind = \'other\'; domain = $domain; mx = \'\' }',
      '  if ($domain -match \'^(gmail|googlemail)\\.com$\') { $res.kind = \'google\'; return $res }',
      '  if ($domain -match \'^(outlook|hotmail|live|msn|windowslive|passport)\\.[a-z.]+$\') { $res.kind = \'microsoft-personal\'; return $res }',
      '  $mx = \'\'',
      '  if ($env:PAYSLIP_MX) { $mx = $env:PAYSLIP_MX }',
      '  else {',
      '    try { $mx = (@(Resolve-DnsName -Name $domain -Type MX -ErrorAction Stop | Where-Object { $_.NameExchange } | ForEach-Object { $_.NameExchange }) -join \' \') } catch { }',
      '    if (-not $mx) { try { $mx = ((& nslookup -type=mx $domain 2>$null) -join \' \') } catch { } }',
      '  }',
      '  $res.mx = $mx.ToLower()',
      '  if ($res.mx -match \'protection\\.outlook\\.com|\\.outlook\\.com|office365\') { $res.kind = \'microsoft\' }',
      '  elseif ($res.mx -match \'google\\.com|googlemail\\.com\') { $res.kind = \'google\' }',
      '  return $res',
      '}',
      '',
      'function Provider-Name($p) {',
      '  switch ($p.kind) {',
      '    \'microsoft\' { return \'Microsoft 365\' }',
      '    \'microsoft-personal\' { return \'Outlook.com (Microsoft)\' }',
      '    \'google\' { return \'Google (Gmail)\' }',
      '    default { return \'your email provider (\' + $p.domain + \')\' }',
      '  }',
      '}',
      '',
      'function Guess-Server($p) {',
      '  switch ($p.kind) {',
      '    \'google\' { return \'smtp.gmail.com:587\' }',
      '    \'microsoft\' { return \'smtp.office365.com:587\' }',
      '    \'microsoft-personal\' { return \'smtp-mail.outlook.com:587\' }',
      '  }',
      '  $mx = $p.mx',
      '  if ($mx -match \'zoho\\.(com|eu|in|sa|com\\.au|jp|com\\.cn)\') { return \'smtp.zoho.\' + $Matches[1] + \':587\' }',
      '  if ($mx -match \'secureserver\\.net\') { return \'smtpout.secureserver.net:587\' }',
      '  if ($mx -match \'hostinger\') { return \'smtp.hostinger.com:587\' }',
      '  if ($mx -match \'yahoodns|yahoo\\.com\') { return \'smtp.mail.yahoo.com:587\' }',
      '  if ($mx -match \'icloud\\.com|me\\.com\') { return \'smtp.mail.me.com:587\' }',
      '  return \'mail.\' + $p.domain + \':587\'',
      '}',
      '',
      '# --- web requests: read Microsoft\'s error text the same way on PowerShell 5 and 7 --------------------',
      'function Read-WebError($err) {',
      '  $res = @{ status = 0; text = \'\'; json = $null; message = $err.Exception.Message }',
      '  try { $res.status = [int]$err.Exception.Response.StatusCode } catch { }',
      '  if ($err.ErrorDetails -and $err.ErrorDetails.Message) { $res.text = $err.ErrorDetails.Message }',
      '  elseif ($err.Exception.Response) {',
      '    try { $sr = New-Object IO.StreamReader($err.Exception.Response.GetResponseStream()); $res.text = $sr.ReadToEnd() } catch { }',
      '  }',
      '  try { $res.json = $res.text | ConvertFrom-Json } catch { }',
      '  if ($res.json) {',
      '    if ($res.json.error_description) { $res.message = [string]$res.json.error_description }',
      '    elseif ($res.json.error.message) { $res.message = [string]$res.json.error.message }',
      '  }',
      '  return $res',
      '}',
      '',
      'function Post-Form($url, $fields) {',
      '  $body = ($fields.Keys | ForEach-Object { $_ + \'=\' + [Uri]::EscapeDataString([string]$fields[$_]) }) -join \'&\'',
      '  return Invoke-RestMethod -Method Post -Uri $url -ContentType \'application/x-www-form-urlencoded\' -Body $body -UseBasicParsing',
      '}',
      '',
      '# --- Microsoft sign-in (works with New Outlook, Outlook on the web, Microsoft 365 and Outlook.com) -----',
      'function Microsoft-Hint($message) {',
      '  if ($message -match \'AADSTS65001|AADSTS90094|AADSTS90095|admin(istrator)? (approval|consent)|needs permission\') { return \'Your company only lets an administrator approve this sign-in ("Microsoft Graph Command Line Tools"). Ask IT to approve it, or use classic Outlook or the mail server.\' }',
      '  if ($message -match \'AADSTS53003|AADSTS50105|AADSTS530|Conditional Access|blocked\') { return \'Your company\'\'s sign-in rules block this sign-in. Ask IT, or use classic Outlook.\' }',
      '  if ($message -match \'expired_token|AADSTS70020\') { return \'The code expired before the sign-in finished. Run it again.\' }',
      '  if ($message -match \'access_denied|AADSTS65004|declined\') { return \'The sign-in was cancelled. Run it again and accept "send mail as you".\' }',
      '  if ($message -match \'MailboxNotEnabledForRESTAPI|MailboxNotHostedInExchangeOnline\') { return \'This mailbox is not on Microsoft\'\'s online service. Use classic Outlook or the mail server.\' }',
      '  if ($message -match \'ErrorSendAsDenied|ErrorAccessDenied\') { return \'Microsoft refused to send from this mailbox. Check that you signed in with the payroll email address.\' }',
      '  if ($message -match \'remote name could not be resolved|No such host|Unable to connect|actively refused|timed out\') { return \'Microsoft could not be reached. Check the internet connection.\' }',
      '  return \'\'',
      '}',
      '',
      'function Connect-Microsoft($address, $personal) {',
      '  $tenant = if ($personal) { \'consumers\' } else { \'organizations\' }',
      '  $base = $AuthBase + \'/\' + $tenant + \'/oauth2/v2.0/\'',
      '  $dc = Post-Form ($base + \'devicecode\') @{ client_id = $MsClient; scope = $MsScope }',
      '  Say \'\'',
      '  Say \'Sign in with Microsoft (once for this run; no password is typed here):\' Cyan',
      '  Say (\'  1. Your browser opens \' + $dc.verification_uri + \' (open it yourself if it does not).\')',
      '  Say (\'  2. Enter this code:  \' + $dc.user_code) Green',
      '  Say (\'  3. Sign in as \' + $address + \' and accept "send mail as you". Then come back to this window.\')',
      '  try { Set-Clipboard -Value $dc.user_code; Say \'     (The code is copied: paste it with Ctrl+V.)\' DarkGray } catch { }',
      '  if (-not $env:PAYSLIP_NO_BROWSER) { try { Start-Process $dc.verification_uri } catch { } }',
      '  $interval = 5; if ($dc.interval) { $interval = [int]$dc.interval }',
      '  if ($env:PAYSLIP_POLL_MS) { $wait = [int]$env:PAYSLIP_POLL_MS } else { $wait = $interval * 1000 }',
      '  $deadline = (Get-Date).AddSeconds([int]$dc.expires_in)',
      '  while ((Get-Date) -lt $deadline) {',
      '    Start-Sleep -Milliseconds $wait',
      '    try {',
      '      $tok = Post-Form ($base + \'token\') @{ grant_type = \'urn:ietf:params:oauth:grant-type:device_code\'; client_id = $MsClient; device_code = $dc.device_code }',
      '      Say \'Signed in.\' Green',
      '      return @{ token = $tok.access_token; refresh = $tok.refresh_token; expires = (Get-Date).AddSeconds([int]$tok.expires_in - 120); base = $base }',
      '    } catch {',
      '      $w = Read-WebError $_',
      '      $code = if ($w.json) { [string]$w.json.error } else { \'\' }',
      '      if ($code -eq \'authorization_pending\') { continue }',
      '      if ($code -eq \'slow_down\') { $wait += 5000; continue }',
      '      throw (\'Microsoft sign-in failed: \' + $(if ($code) { $code + \': \' } else { \'\' }) + $w.message)',
      '    }',
      '  }',
      '  throw \'Microsoft sign-in failed: expired_token (no sign-in within the time allowed)\'',
      '}',
      '',
      'function Refresh-Microsoft {',
      '  $tok = Post-Form ($ms.base + \'token\') @{ grant_type = \'refresh_token\'; client_id = $MsClient; refresh_token = $ms.refresh; scope = $MsScope }',
      '  $ms.token = $tok.access_token',
      '  if ($tok.refresh_token) { $ms.refresh = $tok.refresh_token }',
      '  $ms.expires = (Get-Date).AddSeconds([int]$tok.expires_in - 120)',
      '}',
      '',
      'function Send-Graph($to, $subject, $body, $file) {',
      '  if ($ms.refresh -and (Get-Date) -gt $ms.expires) { Refresh-Microsoft }',
      '  $att = @{ \'@odata.type\' = \'#microsoft.graph.fileAttachment\'; name = [IO.Path]::GetFileName($file); contentType = \'application/pdf\'; contentBytes = [Convert]::ToBase64String([IO.File]::ReadAllBytes($file)) }',
      '  $msg = @{ subject = $subject; body = @{ contentType = \'Text\'; content = $body }; toRecipients = @(@{ emailAddress = @{ address = $to } }); attachments = @($att) }',
      '  # sent as UTF-8 bytes: PowerShell 5 would otherwise spoil Arabic names',
      '  $json = [Text.Encoding]::UTF8.GetBytes((ConvertTo-Json -InputObject @{ message = $msg; saveToSentItems = $true } -Depth 6 -Compress))',
      '  for ($try = 1; $try -le 4; $try++) {',
      '    try {',
      '      [void](Invoke-RestMethod -Method Post -Uri ($GraphBase + \'/me/sendMail\') -Headers @{ Authorization = \'Bearer \' + $ms.token } -ContentType \'application/json; charset=utf-8\' -Body $json -UseBasicParsing)',
      '      return',
      '    } catch {',
      '      $w = Read-WebError $_',
      '      if (($w.status -eq 429 -or $w.status -eq 503 -or $w.status -eq 504) -and $try -lt 4) {',
      '        Say (\'  Microsoft asks to slow down; waiting \' + (20 * $try) + \' seconds...\') DarkGray',
      '        if (-not $env:PAYSLIP_POLL_MS) { Start-Sleep -Seconds (20 * $try) }',
      '        continue',
      '      }',
      '      if ($w.status -eq 401 -and $try -eq 1 -and $ms.refresh) { Refresh-Microsoft; continue }',
      '      $code = if ($w.json -and $w.json.error.code) { [string]$w.json.error.code + \': \' } else { \'\' }',
      '      throw ($code + $w.message)',
      '    }',
      '  }',
      '}',
      '',
      '# --- mail server: password asked here, used for this run only, never saved ------------------------',
      'function Open-MailServer($address, $p) {',
      '  $srv = $Server',
      '  if (-not $srv) {',
      '    $srv = Guess-Server $p',
      '    if (-not $Yes) {',
      '      $typed = Read-Host (\'Mail server: \' + $srv + \'  (press Enter to keep it, or type the one your email provider gives)\')',
      '      if ($typed) { $srv = $typed.Trim() }',
      '    }',
      '  }',
      '  $parts = $srv -split \':\'',
      '  $hostName = $parts[0]',
      '  $port = 587',
      '  if ($parts.Count -gt 1) { $port = [int]$parts[1] }',
      '  if ($port -eq 465) { Say \'Port 465 is not supported here; using 587 (the same server accepts it).\' DarkGray; $port = 587 }',
      '  $client = New-Object System.Net.Mail.SmtpClient($hostName, $port)',
      '  $client.EnableSsl = -not $NoTls',
      '  $client.Timeout = 60000',
      '  if ($env:PAYSLIP_SMTP_PASSWORD) {',
      '    $client.Credentials = New-Object System.Net.NetworkCredential($address, $env:PAYSLIP_SMTP_PASSWORD)',
      '  } elseif (-not $Yes) {',
      '    if ($p.kind -eq \'google\') { Say \'Gmail needs an APP PASSWORD, not your normal password: create one at https://myaccount.google.com/apppasswords (2-Step Verification must be on).\' Yellow }',
      '    $secure = Read-Host (\'Password for \' + $address + \' (not shown while you type)\') -AsSecureString',
      '    $plain = (New-Object System.Net.NetworkCredential(\'\', $secure)).Password',
      '    if (-not $plain) { throw \'No password given.\' }',
      '    $client.Credentials = New-Object System.Net.NetworkCredential($address, $plain)',
      '  }',
      '  return @{ client = $client; from = $address; server = ($hostName + \':\' + $port) }',
      '}',
      '',
      'function Mail-ServerHint($message) {',
      '  if ($message -match \'basic authentication is disabled|5\\.7\\.139|SmtpClientAuthentication is disabled\') { return \'Microsoft no longer accepts a password here. Run it again and let it use the Microsoft sign-in (it does so by itself for Microsoft addresses).\' }',
      '  if ($message -match \'Username and Password not accepted|5\\.7\\.8|BadCredentials|535\') { return \'The password was not accepted. Gmail and accounts with two-step sign-in need an app password (Gmail: https://myaccount.google.com/apppasswords).\' }',
      '  if ($message -match \'5\\.7\\.0|5\\.7\\.57|530|Authentication|authenticat\') { return \'The mail server refused the sign-in. Check the address and password; Gmail needs an app password.\' }',
      '  if ($message -match \'remote name could not be resolved|No such host|actively refused|timed out|Timeout|Failure sending mail\') { return \'The mail server could not be reached. Check the internet connection and the server name (your email provider\'\'s help pages list it as "SMTP server").\' }',
      '  return \'\'',
      '}',
      '',
      '# --- the list ------------------------------------------------------------------------------------',
      'Say \'\'',
      'Say (\'Payslips \' + $Company + \' - \' + $Period) Cyan',
      'if (-not (Test-Path $listPath)) { Say \'send-list.csv is missing. Keep all files of this pack together.\' Red; Finish; exit 1 }',
      '$rows = @(Import-Csv -Path $listPath -Encoding UTF8)',
      'if ($rows.Count -eq 0) { Say \'send-list.csv has no employees.\' Red; Finish; exit 1 }',
      '$missing = @($rows | Where-Object { -not (Test-Path (Join-Path $pdfDir $_.file)) })',
      '$done = @{}',
      'if (Test-Path $logPath) {',
      '  Import-Csv -Path $logPath -Encoding UTF8 | ForEach-Object { if ($_.action -eq \'sent\') { $done[$_.employee_id] = $true } }',
      '}',
      '$todo = @($rows | Where-Object { -not $done.ContainsKey($_.employee_id) })',
      '$address = if ($From) { $From } else { $Sender }',
      '',
      'if ($Check) {',
      '  foreach ($r in $rows) {',
      '    $state = if ($done.ContainsKey($r.employee_id)) { \'already sent\' } elseif (Test-Path (Join-Path $pdfDir $r.file)) { \'ready\' } else { \'PDF MISSING\' }',
      '    Say (\'  \' + $r.name + \' <\' + $r.email + \'>  \' + $r.file + \'  [\' + $state + \']\')',
      '  }',
      '  Say (\'\' + $rows.Count + \' employees, \' + $todo.Count + \' to send, \' + $missing.Count + \' PDF missing.\') Cyan',
      '  if ($Outlook) { Say \'Outlook: provided.\' Green } else {',
      '    $o = Find-Outlook',
      '    if (Use-32BitPowerShell $o) { exit $script:relaunchCode }',
      '    if ($o.ok) {',
      '      $names = @(); foreach ($a in $o.app.GetNamespace(\'MAPI\').Accounts) { $names += $a.SmtpAddress }',
      '      Say (\'Ready: classic Outlook answers (\' + ($names -join \', \') + \'). Send all uses it.\') Green',
      '    } else {',
      '      Explain-Outlook $o',
      '      if ($address) {',
      '        $p = Find-Provider $address',
      '        if ($p.kind -like \'microsoft*\') { Say (\'Ready another way: \' + $address + \' is \' + (Provider-Name $p) + \', so Send all signs in with Microsoft (a code in your browser).\') Green }',
      '        else { Say (\'Ready another way: \' + $address + \' is \' + (Provider-Name $p) + \', so Send all uses its mail server \' + (Guess-Server $p) + \' and asks the password.\') Green }',
      '      } else { Say \'Ready another way: Send all asks your email address, then signs in with Microsoft or uses your mail server.\' Green }',
      '    }',
      '  }',
      '  Say \'Nothing was sent.\'',
      '  Finish; exit 0',
      '}',
      'if ($missing.Count -gt 0) {',
      '  Say (\'These PDF files are missing in the pdf folder: \' + (($missing | ForEach-Object { $_.file }) -join \', \')) Red',
      '  Finish; exit 1',
      '}',
      '',
      '# --- how to send: classic Outlook, Microsoft sign-in, or the mail server --------------------------',
      '$route = $Route',
      'if ($Smtp) { $route = \'server\' }',
      'if ($Outlook -and -not $route) { $route = \'outlook\' }',
      'if ($route -eq \'auto\') { $route = \'\' }   # "Send without Outlook": decide by the email address',
      'elseif (-not $route -or ($route -eq \'outlook\' -and -not $Outlook)) {',
      '  $o = Find-Outlook',
      '  if (Use-32BitPowerShell $o) { exit $script:relaunchCode }',
      '  if ($o.ok) { $Outlook = $o.app; $route = \'outlook\' }',
      '  else {',
      '    Explain-Outlook $o',
      '    if ($route -eq \'outlook\') { Finish; exit 1 }',
      '    if ($Drafts) { Say \'Saving drafts needs classic Outlook. "2 - Send all" works without it.\' Red; Finish; exit 1 }',
      '  }',
      '}',
      '$provider = $null',
      'if ($route -ne \'outlook\') {',
      '  if (-not $address -and -not $Yes) { $address = (Read-Host \'Your email address (the payslips are sent from it)\').Trim() }',
      '  if (-not $address) { Say \'Start it by double-clicking "2 - Send all.bat" to type your email address.\' Yellow; Finish; exit 1 }',
      '  $provider = Find-Provider $address',
      '  if (-not $route) {',
      '    $route = if ($provider.kind -like \'microsoft*\') { \'microsoft\' } else { \'server\' }',
      '    Say \'\'',
      '    if ($route -eq \'microsoft\') {',
      '      Say ($address + \' is \' + (Provider-Name $provider) + \': it sends through a Microsoft sign-in, without Outlook.\') Cyan',
      '      if (-not $Yes) {',
      '        $go = Read-Host \'Press Enter to sign in with Microsoft (or type P to use a mail server and password instead)\'',
      '        if ($go -match \'^[Pp]\') { $route = \'server\' }',
      '      }',
      '    } else {',
      '      Say ($address + \' is \' + (Provider-Name $provider) + \': it sends through its mail server with your email password.\') Cyan',
      '    }',
      '  }',
      '}',
      '$account = $null',
      '$me = $address',
      '$ns = $null',
      '$mail = $null',
      '$ms = $null',
      'try {',
      '  if ($route -eq \'outlook\') {',
      '    $ns = $Outlook.GetNamespace(\'MAPI\')',
      '    $me = $null',
      '    foreach ($a in $ns.Accounts) {',
      '      if (-not $me) { $me = $a.SmtpAddress }',
      '      if ($Sender -and $a.SmtpAddress -eq $Sender) { $account = $a; $me = $a.SmtpAddress }',
      '    }',
      '  } elseif ($route -eq \'microsoft\') {',
      '    $ms = Connect-Microsoft $address ($provider.kind -eq \'microsoft-personal\')',
      '  } else {',
      '    $mail = Open-MailServer $address $provider',
      '    Say (\'Sending through \' + $mail.server + \' as \' + $address) DarkGray',
      '  }',
      '} catch {',
      '  Say $_.Exception.Message Red',
      '  $hint = if ($route -eq \'microsoft\') { Microsoft-Hint $_.Exception.Message } else { Mail-ServerHint $_.Exception.Message }',
      '  if ($hint) { Say $hint Yellow }',
      '  Finish; exit 1',
      '}',
      '',
      '# Sends (or saves as a draft) one email; throws on failure.',
      'function Send-One($to, $subject, $body, $file) {',
      '  if ($route -eq \'outlook\') {',
      '    $item = $Outlook.CreateItem(0)',
      '    $item.To = $to',
      '    $item.Subject = $subject',
      '    $item.Body = $body',
      '    [void]$item.Attachments.Add($file)',
      '    if ($account) { $item.SendUsingAccount = $account }',
      '    if ($Drafts) { $item.Save() } else { $item.Send() }',
      '  } elseif ($route -eq \'microsoft\') {',
      '    Send-Graph $to $subject $body $file',
      '  } else {',
      '    $msg = New-Object System.Net.Mail.MailMessage',
      '    try {',
      '      $msg.From = New-Object System.Net.Mail.MailAddress($mail.from)',
      '      $msg.To.Add($to)',
      '      $msg.Subject = $subject',
      '      $msg.SubjectEncoding = [Text.Encoding]::UTF8',
      '      $msg.Body = $body',
      '      $msg.BodyEncoding = [Text.Encoding]::UTF8',
      '      $att = New-Object System.Net.Mail.Attachment($file, \'application/pdf\')',
      '      $att.ContentDisposition.FileName = [IO.Path]::GetFileName($file)',
      '      $msg.Attachments.Add($att)',
      '      $mail.client.Send($msg)',
      '    } finally { $msg.Dispose() }',
      '  }',
      '}',
      'function Route-Hint($message) {',
      '  switch ($route) {',
      '    \'outlook\' { return Outlook-Hint $message }',
      '    \'microsoft\' { return Microsoft-Hint $message }',
      '    default { return Mail-ServerHint $message }',
      '  }',
      '}',
      '',
      'if ($Test) {',
      '  $r = $rows[0]',
      '  if (-not $me -and -not $Yes) { $me = Read-Host \'Your email address for the test\' }',
      '  try {',
      '    Send-One $me (\'[TEST] \' + $r.subject) (\'TEST: this is what \' + $r.name + \' <\' + $r.email + \'> will receive.\' + "`r`n`r`n" + $r.body) (Join-Path $pdfDir $r.file)',
      '    if ($ns) { try { $ns.SendAndReceive($false) } catch { } }',
      '    Say (\'Test email sent to \' + $me + \'. Check it in your inbox, then run "2 - Send all".\') Green',
      '  } catch {',
      '    Say (\'The test email could not be sent: \' + $_.Exception.Message) Red',
      '    $hint = Route-Hint $_.Exception.Message',
      '    if ($hint) { Say $hint Yellow }',
      '  }',
      '  Finish; exit 0',
      '}',
      '',
      '$verb = if ($Drafts) { \'save as drafts\' } else { \'send\' }',
      'if ($todo.Count -eq 0) { Say \'Everyone in this pack has been sent already.\' Green; Finish; exit 0 }',
      'Say (\'Ready to \' + $verb + \' \' + $todo.Count + \' emails\' + $(if ($me) { \' from \' + $me } else { \'\' }) + \'.\')',
      'if ($done.Count -gt 0) { Say (\'\' + $done.Count + \' were sent earlier and are skipped.\') }',
      'Say (\'First: \' + $todo[0].name + \' <\' + $todo[0].email + \'>\')',
      'Say \'Press Ctrl+C at any time to stop. Run it again later to continue where it stopped.\' DarkGray',
      'if ($route -eq \'outlook\') { Say \'If Outlook asks "A program is trying to send an e-mail", choose Allow.\' DarkGray }',
      'if (-not $Yes) {',
      '  $answer = Read-Host (\'Type YES to \' + $verb + \' \' + $todo.Count + \' emails\')',
      '  if ($answer -ne \'YES\') { Say \'Nothing was sent.\'; Finish; exit 0 }',
      '}',
      'if (-not $Drafts -and -not (Test-Path $logPath)) { Set-Content -Path $logPath -Value \'timestamp,employee_id,name,channel,contact,file,sha256,note,action,source\' -Encoding UTF8 }',
      '$how = switch ($route) { \'outlook\' { \'sent by Outlook (send-all pack)\' } \'microsoft\' { \'sent by Microsoft sign-in (send-all pack)\' } default { \'sent by mail server (send-all pack)\' } }',
      '',
      '$i = 0; $ok = 0; $failed = 0',
      'foreach ($r in $todo) {',
      '  $i++',
      '  $file = Join-Path $pdfDir $r.file',
      '  $tag = \'[\' + $i + \'/\' + $todo.Count + \'] \'',
      '  try {',
      '    Send-One $r.email $r.subject $r.body $file',
      '    if ($Drafts) {',
      '      Say ($tag + \'Draft saved for \' + $r.name + \' <\' + $r.email + \'>\') Green',
      '    } else {',
      '      $hash = (Get-FileHash -Algorithm SHA256 -Path $file).Hash.ToLower()',
      '      $stamp = (Get-Date).ToUniversalTime().ToString(\'yyyy-MM-ddTHH:mm:ss.fffZ\')',
      '      Add-Content -Path $logPath -Value (CsvLine @($stamp, $r.employee_id, $r.name, \'outlook\', $r.email, $r.file, $hash, $how, \'sent\', $r.source)) -Encoding UTF8',
      '      Say ($tag + \'Sent to \' + $r.name + \' <\' + $r.email + \'>\') Green',
      '    }',
      '    $ok++',
      '  } catch {',
      '    $failed++',
      '    $why = $_.Exception.Message',
      '    Say ($tag + \'FAILED \' + $r.name + \' <\' + $r.email + \'>: \' + $why) Red',
      '    Add-Content -Path $errPath -Value (CsvLine @((Get-Date).ToString(\'s\'), $r.employee_id, $r.name, $r.email, $why)) -Encoding UTF8',
      '    $hint = Route-Hint $why',
      '    if ($hint -and $ok -eq 0) { Say $hint Yellow; Say \'Stopped before trying the others.\' Yellow; break }',
      '  }',
      '  if ($i -lt $todo.Count -and $Delay -gt 0) { Start-Sleep -Milliseconds ([int]($Delay * 1000)) }',
      '}',
      'if ($ns) { try { $ns.SendAndReceive($false) } catch { } }',
      'Say \'\'',
      'if ($Drafts) {',
      '  Say (\'Done: \' + $ok + \' drafts saved in Outlook, \' + $failed + \' failed. Check the Drafts folder and send them from Outlook.\') Cyan',
      '} else {',
      '  Say (\'Done: \' + $ok + \' sent, \' + $failed + \' failed.\') Cyan',
      '  if ($route -eq \'outlook\') { Say \'Keep Outlook open until the Outbox is empty.\' }',
      '  Say \'Then load delivery-log.csv on the Send step (More > Restore ticks from a log) to tick everyone.\'',
      '}',
      'if ($failed -gt 0) { Say \'The failed ones are listed in errors.csv. Fix them and run Send all again: the others are skipped.\' Yellow }',
      'Finish'
    ].join('\r\n') + '\r\n';
  }

  function windowsBat(args, title) {
    return ['@echo off', 'title ' + title, 'powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0send-all.ps1"' + (args ? ' ' + args : ''), ''].join('\r\n');
  }

  // Mac and Linux: a bash script (bash and curl only) that hands each email to Apple Mail through
  // AppleScript, or without Apple Mail sends through a Microsoft sign-in or the provider's mail server.
  function macSendScript(info) {
    return [
      '#!/bin/bash',
      '# Payslip Generator - sends every payslip email from this Mac or Linux computer, whatever email it uses:',
      '# Apple Mail if it has an account, otherwise a Microsoft sign-in (Microsoft 365, Outlook.com) or your',
      '# email provider\'s mail server (Gmail, Zoho, company mail). Needs only bash and curl.',
      '# Start it with "2 - Send all.command" (try "1 - Send test to me.command" first); on Linux: bash send-all.sh',
      'cd "$(dirname "$0")" || exit 1',
      'MODE="${1:-send}"',
      'OSA="${PAYSLIP_OSASCRIPT:-osascript}"',
      'DELAY="${PAYSLIP_DELAY:-2.5}"',
      'ROUTE="${PAYSLIP_ROUTE:-}"',
      'COMPANY=' + shString(info.company),
      'PERIOD=' + shString(info.period),
      'SENDER=' + shString(info.sender),
      'LIST="send-list.tsv"',
      'LOG="delivery-log.csv"',
      'SENT="sent-ids.txt"',
      '# Microsoft sign-in: Microsoft\'s own public app for command-line access to your mailbox (the one',
      '# Microsoft Graph PowerShell uses); it only gets "send mail as you", after you agree on Microsoft\'s page.',
      'MS_CLIENT="14d82eec-204b-4c2f-b7e8-296a70dab67e"',
      'MS_SCOPE="https://graph.microsoft.com/Mail.Send offline_access"',
      'AUTH="${PAYSLIP_AUTH_BASE:-https://login.microsoftonline.com}"',
      'GRAPH="${PAYSLIP_GRAPH_BASE:-https://graph.microsoft.com/v1.0}"',
      'TMP=$(mktemp -d 2>/dev/null || mktemp -d -t payslips)',
      'trap \'rm -rf "$TMP"\' EXIT',
      '',
      'finish() { echo; [ -n "$PAYSLIP_YES" ] || read -r -p "Press Enter to close" _; }',
      'csv() { local s="$1"; case "$s" in *[\\",]*) s="${s//\\"/\\"\\"}"; printf \'"%s"\' "$s" ;; *) printf \'%s\' "$s" ;; esac; }',
      'lower() { printf \'%s\' "$1" | tr \'[:upper:]\' \'[:lower:]\'; }',
      '# a "string" or number field from a small JSON reply',
      'json_get() { printf \'%s\' "$1" | tr -d \'\\r\\n\' | sed -n \'s/.*"\'"$2"\'"[[:space:]]*:[[:space:]]*"\\([^"]*\\)".*/\\1/p\' | sed \'s#\\\\/#/#g\'; }',
      'json_num() { printf \'%s\' "$1" | tr -d \'\\r\\n\' | sed -n \'s/.*"\'"$2"\'"[[:space:]]*:[[:space:]]*\\([0-9][0-9]*\\).*/\\1/p\'; }',
      '# text into a JSON string; the \\n of the send list stay line breaks',
      'json_str() { printf \'%s\' "$1" | sed -e \'s/\\\\/\\\\\\\\/g\' -e \'s/\\\\\\\\n/\\\\n/g\' -e \'s/"/\\\\"/g\'; }',
      'pause_ms() { if [ -n "$PAYSLIP_POLL_MS" ]; then sleep 0.01; else sleep "$1"; fi; }',
      '',
      '# --- who hosts your email? Decides how to send without Apple Mail ----------------------------------',
      'find_provider() {',
      '  DOMAIN=$(lower "${1##*@}")',
      '  MX=""',
      '  case "$DOMAIN" in',
      '    gmail.com|googlemail.com) PROVIDER=google; return ;;',
      '    outlook.*|hotmail.*|live.*|msn.com|windowslive.com|passport.com) PROVIDER=microsoft-personal; return ;;',
      '  esac',
      '  if [ -n "$PAYSLIP_MX" ]; then MX="$PAYSLIP_MX"',
      '  elif command -v dig >/dev/null 2>&1; then MX=$(dig +short MX "$DOMAIN" 2>/dev/null | tr \'\\n\' \' \')',
      '  elif command -v host >/dev/null 2>&1; then MX=$(host -t MX "$DOMAIN" 2>/dev/null | tr \'\\n\' \' \')',
      '  elif command -v nslookup >/dev/null 2>&1; then MX=$(nslookup -type=mx "$DOMAIN" 2>/dev/null | tr \'\\n\' \' \')',
      '  fi',
      '  MX=$(lower "$MX")',
      '  case "$MX" in',
      '    *protection.outlook.com*|*.outlook.com*|*office365*) PROVIDER=microsoft ;;',
      '    *google.com*|*googlemail.com*) PROVIDER=google ;;',
      '    *) PROVIDER=other ;;',
      '  esac',
      '}',
      'provider_name() {',
      '  case "$PROVIDER" in',
      '    microsoft) echo "Microsoft 365" ;;',
      '    microsoft-personal) echo "Outlook.com (Microsoft)" ;;',
      '    google) echo "Google (Gmail)" ;;',
      '    *) echo "your email provider ($DOMAIN)" ;;',
      '  esac',
      '}',
      'guess_server() {',
      '  case "$PROVIDER" in',
      '    google) echo "smtp.gmail.com:587"; return ;;',
      '    microsoft) echo "smtp.office365.com:587"; return ;;',
      '    microsoft-personal) echo "smtp-mail.outlook.com:587"; return ;;',
      '  esac',
      '  case "$MX" in',
      '    *zoho.com.au*) echo "smtp.zoho.com.au:587" ;;',
      '    *zoho.sa*) echo "smtp.zoho.sa:587" ;;',
      '    *zoho.eu*) echo "smtp.zoho.eu:587" ;;',
      '    *zoho.in*) echo "smtp.zoho.in:587" ;;',
      '    *zoho.com*) echo "smtp.zoho.com:587" ;;',
      '    *secureserver.net*) echo "smtpout.secureserver.net:587" ;;',
      '    *hostinger*) echo "smtp.hostinger.com:587" ;;',
      '    *yahoodns*|*yahoo.com*) echo "smtp.mail.yahoo.com:587" ;;',
      '    *icloud.com*|*me.com*) echo "smtp.mail.me.com:587" ;;',
      '    *) echo "mail.$DOMAIN:587" ;;',
      '  esac',
      '}',
      '',
      '# --- Apple Mail -------------------------------------------------------------------------------------',
      'mail_accounts() {',
      '  [ -n "$PAYSLIP_OSASCRIPT" ] || [ "$(uname)" = "Darwin" ] || { echo 0; return; }',
      '  n=$("$OSA" -e \'tell application "Mail" to count of accounts\' 2>/dev/null)',
      '  case "$n" in \'\'|*[!0-9]*) echo 0 ;; *) echo "$n" ;; esac',
      '}',
      '',
      '# --- Microsoft sign-in (works with Microsoft 365 and Outlook.com, no Outlook needed) ---------------',
      'ms_hint() {',
      '  case "$1" in',
      '    *AADSTS65001*|*AADSTS90094*|*AADSTS90095*|*"admin approval"*|*"administrator"*) echo "Your company only lets an administrator approve this sign-in (\\"Microsoft Graph Command Line Tools\\"). Ask IT to approve it, or use the mail server." ;;',
      '    *AADSTS53003*|*AADSTS50105*|*AADSTS530*|*"Conditional Access"*|*blocked*) echo "Your company\'s sign-in rules block this sign-in. Ask IT, or use Apple Mail." ;;',
      '    *expired_token*|*AADSTS70020*) echo "The code expired before the sign-in finished. Run it again." ;;',
      '    *access_denied*|*AADSTS65004*|*declined*) echo "The sign-in was cancelled. Run it again and accept \\"send mail as you\\"." ;;',
      '    *MailboxNotEnabledForRESTAPI*|*MailboxNotHostedInExchangeOnline*) echo "This mailbox is not on Microsoft\'s online service. Use Apple Mail or the mail server." ;;',
      '    *ErrorSendAsDenied*|*ErrorAccessDenied*) echo "Microsoft refused to send from this mailbox. Check that you signed in with the payroll email address." ;;',
      '    *"could not be reached"*) echo "Check the internet connection." ;;',
      '  esac',
      '}',
      'ms_connect() {',
      '  local tenant=organizations resp code uri dc interval expires waited at err',
      '  [ "$PROVIDER" = "microsoft-personal" ] && tenant=consumers',
      '  MS_BASE="$AUTH/$tenant/oauth2/v2.0"',
      '  resp=$(curl -sS -X POST "$MS_BASE/devicecode" --data-urlencode "client_id=$MS_CLIENT" --data-urlencode "scope=$MS_SCOPE" 2>&1) || { ERR="Microsoft could not be reached: $resp"; return 1; }',
      '  code=$(json_get "$resp" user_code); uri=$(json_get "$resp" verification_uri); dc=$(json_get "$resp" device_code)',
      '  [ -n "$code" ] || { ERR="Microsoft sign-in failed: $(json_get "$resp" error): $(json_get "$resp" error_description)"; return 1; }',
      '  interval=$(json_num "$resp" interval); [ -n "$interval" ] || interval=5',
      '  expires=$(json_num "$resp" expires_in); [ -n "$expires" ] || expires=900',
      '  echo',
      '  echo "Sign in with Microsoft (once for this run; no password is typed here):"',
      '  echo "  1. Your browser opens $uri (open it yourself if it does not)."',
      '  echo "  2. Enter this code:  $code"',
      '  echo "  3. Sign in as $ADDRESS and accept \\"send mail as you\\". Then come back to this window."',
      '  if command -v pbcopy >/dev/null 2>&1; then printf \'%s\' "$code" | pbcopy && echo "     (The code is copied: paste it with Cmd+V.)"; fi',
      '  if [ -z "$PAYSLIP_NO_BROWSER" ]; then',
      '    if command -v open >/dev/null 2>&1 && [ "$(uname)" = "Darwin" ]; then open "$uri"; elif command -v xdg-open >/dev/null 2>&1; then xdg-open "$uri" >/dev/null 2>&1 & fi',
      '  fi',
      '  waited=0',
      '  while [ "$waited" -lt "$expires" ]; do',
      '    pause_ms "$interval"; waited=$((waited + interval))',
      '    resp=$(curl -sS -X POST "$MS_BASE/token" --data-urlencode "grant_type=urn:ietf:params:oauth:grant-type:device_code" --data-urlencode "client_id=$MS_CLIENT" --data-urlencode "device_code=$dc" 2>&1) || continue',
      '    at=$(json_get "$resp" access_token)',
      '    if [ -n "$at" ]; then',
      '      TOKEN="$at"; REFRESH=$(json_get "$resp" refresh_token)',
      '      TOKEN_UNTIL=$(( $(date +%s) + $(json_num "$resp" expires_in) - 120 ))',
      '      echo "Signed in."',
      '      return 0',
      '    fi',
      '    err=$(json_get "$resp" error)',
      '    case "$err" in',
      '      authorization_pending) continue ;;',
      '      slow_down) interval=$((interval + 5)); continue ;;',
      '      *) ERR="Microsoft sign-in failed: $err: $(json_get "$resp" error_description)"; return 1 ;;',
      '    esac',
      '  done',
      '  ERR="Microsoft sign-in failed: expired_token (no sign-in within the time allowed)"; return 1',
      '}',
      'ms_refresh() {',
      '  local resp at',
      '  resp=$(curl -sS -X POST "$MS_BASE/token" --data-urlencode "grant_type=refresh_token" --data-urlencode "client_id=$MS_CLIENT" --data-urlencode "refresh_token=$REFRESH" --data-urlencode "scope=$MS_SCOPE" 2>&1) || return 1',
      '  at=$(json_get "$resp" access_token); [ -n "$at" ] || return 1',
      '  TOKEN="$at"; [ -n "$(json_get "$resp" refresh_token)" ] && REFRESH=$(json_get "$resp" refresh_token)',
      '  TOKEN_UNTIL=$(( $(date +%s) + $(json_num "$resp" expires_in) - 120 ))',
      '}',
      'ms_send() { # to subject body file',
      '  local try status resp',
      '  [ -n "$REFRESH" ] && [ "$(date +%s)" -gt "$TOKEN_UNTIL" ] && ms_refresh',
      '  {',
      '    printf \'{"message":{"subject":"%s","body":{"contentType":"Text","content":"%s"},\' "$(json_str "$2")" "$(json_str "$3")"',
      '    printf \'"toRecipients":[{"emailAddress":{"address":"%s"}}],\' "$(json_str "$1")"',
      '    printf \'"attachments":[{"@odata.type":"#microsoft.graph.fileAttachment","name":"%s","contentType":"application/pdf","contentBytes":"\' "$(json_str "$(basename "$4")")"',
      '    base64 < "$4" | tr -d \'\\r\\n\'',
      '    printf \'"}]},"saveToSentItems":true}\'',
      '  } > "$TMP/msg.json"',
      '  for try in 1 2 3 4; do',
      '    status=$(curl -sS -o "$TMP/resp" -w \'%{http_code}\' -X POST "$GRAPH/me/sendMail" -H "Authorization: Bearer $TOKEN" -H "Content-Type: application/json; charset=utf-8" --data-binary @"$TMP/msg.json" 2>"$TMP/curl-err") || status=000',
      '    case "$status" in',
      '      2*) return 0 ;;',
      '      429|503|504) if [ "$try" -lt 4 ]; then echo "  Microsoft asks to slow down; waiting $((20 * try)) seconds..."; pause_ms $((20 * try)); continue; fi ;;',
      '      401) if [ "$try" -eq 1 ] && [ -n "$REFRESH" ] && ms_refresh; then continue; fi ;;',
      '      000) ERR="Microsoft could not be reached: $(cat "$TMP/curl-err")"; return 1 ;;',
      '    esac',
      '    resp=$(cat "$TMP/resp" 2>/dev/null)',
      '    ERR="$(json_get "$resp" code): $(json_get "$resp" message)"',
      '    return 1',
      '  done',
      '}',
      '',
      '# --- mail server: password asked here, used for this run only, never saved -------------------------',
      'server_hint() {',
      '  case "$1" in',
      '    *"(67)"*|*"Login denied"*|*535*) echo "The password was not accepted. Gmail and accounts with two-step sign-in need an app password (Gmail: https://myaccount.google.com/apppasswords). Microsoft addresses use the Microsoft sign-in instead." ;;',
      '    *"(6)"*|*"(7)"*|*"(28)"*|*"resolve"*|*"connect"*) echo "The mail server could not be reached. Check the internet connection and the server name (your email provider\'s help pages list it as \\"SMTP server\\")." ;;',
      '    *"(64)"*|*"(35)"*) echo "The mail server would not start a secure connection. Check the server name and port (587)." ;;',
      '  esac',
      '}',
      'server_open() {',
      '  SRV="${PAYSLIP_SERVER:-}"',
      '  if [ -z "$SRV" ]; then',
      '    SRV=$(guess_server)',
      '    if [ -z "$PAYSLIP_YES" ]; then',
      '      read -r -p "Mail server: $SRV  (press Enter to keep it, or type the one your email provider gives): " typed',
      '      [ -n "$typed" ] && SRV="$typed"',
      '    fi',
      '  fi',
      '  case "$SRV" in *:*) ;; *) SRV="$SRV:587" ;; esac',
      '  case "$SRV" in *:465) SMTP_URL="smtps://$SRV" ;; *) SMTP_URL="smtp://$SRV" ;; esac',
      '  TLS="--ssl-reqd"; [ -n "$PAYSLIP_NO_TLS" ] && TLS=""',
      '  PASS="${PAYSLIP_SMTP_PASSWORD:-}"',
      '  if [ -z "$PASS" ] && [ -z "$PAYSLIP_YES" ]; then',
      '    [ "$PROVIDER" = "google" ] && echo "Gmail needs an APP PASSWORD, not your normal password: create one at https://myaccount.google.com/apppasswords (2-Step Verification must be on)."',
      '    read -r -s -p "Password for $ADDRESS (not shown while you type): " PASS; echo',
      '  fi',
      '  [ -n "$PASS" ] || { ERR="No password given."; return 1; }',
      '  # the password goes to curl in a private file, not on the command line',
      '  (umask 077; printf \'user = "%s:%s"\\n\' "$(printf \'%s\' "$ADDRESS" | sed \'s/[\\\\"]/\\\\&/g\')" "$(printf \'%s\' "$PASS" | sed \'s/[\\\\"]/\\\\&/g\')" > "$TMP/curl.conf")',
      '  echo "Sending through $SRV as $ADDRESS"',
      '}',
      'server_send() { # to subject body file',
      '  printf \'%s\\n\' "$3" | awk \'{ gsub(/\\\\n/, "\\r\\n"); printf "%s\\r\\n", $0 }\' > "$TMP/body.txt"',
      '  local subj="=?UTF-8?B?$(printf \'%s\' "$2" | base64 | tr -d \'\\r\\n\')?="',
      '  if curl -sS $TLS --url "$SMTP_URL" -K "$TMP/curl.conf" --mail-from "$ADDRESS" --mail-rcpt "$1" \\',
      '      -H "Subject: $subj" -H "From: $ADDRESS" -H "To: $1" \\',
      '      -F "=<$TMP/body.txt;type=text/plain; charset=utf-8;encoder=quoted-printable" \\',
      '      -F "=@$4;type=application/pdf;encoder=base64" > /dev/null 2>"$TMP/curl-err"; then',
      '    return 0',
      '  fi',
      '  ERR=$(tr -d \'\\r\\n\' < "$TMP/curl-err")',
      '  return 1',
      '}',
      '',
      '# --- one email through the chosen route --------------------------------------------------------------',
      'send_one() { # to subject body file',
      '  ERR=""',
      '  case "$ROUTE" in',
      '    mail) "$OSA" send-one.applescript "$1" "$2" "$3" "$PWD/$4" "$SENDER" >/dev/null 2>"$TMP/osa-err" || { ERR=$(tr -d \'\\r\\n\' < "$TMP/osa-err"); return 1; } ;;',
      '    microsoft) ms_send "$1" "$2" "$3" "$4" ;;',
      '    *) server_send "$1" "$2" "$3" "$4" ;;',
      '  esac',
      '}',
      'route_hint() {',
      '  case "$ROUTE" in',
      '    mail) case "$1" in *-1743*|*"Not authorized"*|*"not allowed"*) echo "macOS blocked control of Mail: System Settings > Privacy & Security > Automation > Terminal > turn on Mail." ;; esac ;;',
      '    microsoft) ms_hint "$1" ;;',
      '    *) server_hint "$1" ;;',
      '  esac',
      '}',
      '',
      '# --- the list ----------------------------------------------------------------------------------------',
      'echo',
      'echo "Payslips $COMPANY - $PERIOD"',
      '[ -f "$LIST" ] || { echo "send-list.tsv is missing. Keep all files of this pack together."; finish; exit 1; }',
      'command -v curl >/dev/null 2>&1 || { echo "curl is missing on this computer: install it (for example: sudo apt install curl), then run this again."; finish; exit 1; }',
      'touch "$SENT"',
      'ADDRESS="${PAYSLIP_FROM:-$SENDER}"',
      'total=0; todo=0; first=""; missing=0',
      'while IFS=$\'\\t\' read -r -u 3 id name email subject body file source; do',
      '  total=$((total + 1))',
      '  if ! grep -qxF "$id" "$SENT"; then',
      '    todo=$((todo + 1))',
      '    [ -n "$first" ] || first="$name <$email>"',
      '    if [ ! -f "pdf/$file" ]; then missing=$((missing + 1)); echo "Missing PDF: pdf/$file"; fi',
      '  fi',
      'done 3< <(tail -n +2 "$LIST")',
      '',
      '# --- how to send: Apple Mail, Microsoft sign-in, or the mail server -----------------------------------',
      'if [ "$ROUTE" = "auto" ]; then ROUTE=""   # "Send without Apple Mail": decide by the email address',
      'elif [ -z "$ROUTE" ]; then',
      '  accounts=$(mail_accounts)',
      '  if [ "$accounts" -gt 0 ]; then ROUTE=mail',
      '  elif [ "$(uname)" = "Darwin" ]; then echo "Apple Mail has no email account set up, so it sends another way."',
      '  fi',
      'fi',
      '',
      'if [ "$MODE" = "check" ]; then',
      '  echo "$total employees, $todo to send, $missing PDF missing."',
      '  if [ "$ROUTE" = "mail" ]; then echo "Ready: Apple Mail has $accounts email account(s). Send all uses it."',
      '  elif [ -n "$ADDRESS" ]; then',
      '    find_provider "$ADDRESS"',
      '    case "$PROVIDER" in',
      '      microsoft*) echo "Ready: $ADDRESS is $(provider_name), so Send all signs in with Microsoft (a code in your browser)." ;;',
      '      *) echo "Ready: $ADDRESS is $(provider_name), so Send all uses its mail server $(guess_server) and asks the password." ;;',
      '    esac',
      '  else echo "Ready: Send all asks your email address, then signs in with Microsoft or uses your mail server."',
      '  fi',
      '  echo "Nothing was sent."',
      '  finish; exit 0',
      'fi',
      '[ "$missing" -eq 0 ] || { echo "Put the missing PDF files back in the pdf folder, then run this again."; finish; exit 1; }',
      '',
      'if [ "$ROUTE" != "mail" ]; then',
      '  if [ -z "$ADDRESS" ] && [ -z "$PAYSLIP_YES" ]; then read -r -p "Your email address (the payslips are sent from it): " ADDRESS; fi',
      '  [ -n "$ADDRESS" ] || { echo "Start it by double-clicking \\"2 - Send all.command\\" to type your email address."; finish; exit 1; }',
      '  find_provider "$ADDRESS"',
      '  if [ -z "$ROUTE" ]; then',
      '    case "$PROVIDER" in microsoft*) ROUTE=microsoft ;; *) ROUTE=server ;; esac',
      '    echo',
      '    if [ "$ROUTE" = "microsoft" ]; then',
      '      echo "$ADDRESS is $(provider_name): it sends through a Microsoft sign-in."',
      '      if [ -z "$PAYSLIP_YES" ]; then',
      '        read -r -p "Press Enter to sign in with Microsoft (or type P to use a mail server and password instead): " go',
      '        case "$go" in [Pp]*) ROUTE=server ;; esac',
      '      fi',
      '    else',
      '      echo "$ADDRESS is $(provider_name): it sends through its mail server with your email password."',
      '    fi',
      '  fi',
      '  if [ "$ROUTE" = "microsoft" ]; then ms_connect || { echo "$ERR"; route_hint "$ERR"; finish; exit 1; }',
      '  else server_open || { echo "$ERR"; finish; exit 1; }',
      '  fi',
      'fi',
      '',
      'if [ "$MODE" = "test" ]; then',
      '  ME="$ADDRESS"',
      '  if [ "$ROUTE" = "mail" ] && [ -z "$ME" ]; then ME=$("$OSA" -e \'tell application "Mail" to get item 1 of (email addresses of item 1 of accounts)\' 2>/dev/null); fi',
      '  [ -n "$ME" ] || read -r -p "Your email address for the test: " ME',
      '  IFS=$\'\\t\' read -r id name email subject body file source < <(sed -n 2p "$LIST")',
      '  if send_one "$ME" "[TEST] $subject" "TEST: this is what $name <$email> will receive.\\n\\n$body" "pdf/$file"; then',
      '    echo "Test email sent to $ME. Check it in your inbox, then run \\"2 - Send all\\"."',
      '  else',
      '    echo "The test email could not be sent: $ERR"',
      '    route_hint "$ERR"',
      '  fi',
      '  finish; exit 0',
      'fi',
      '',
      'if [ "$todo" -eq 0 ]; then echo "Everyone in this pack has been sent already."; finish; exit 0; fi',
      'case "$ROUTE" in mail) how="Apple Mail" ;; microsoft) how="Microsoft sign-in" ;; *) how="mail server" ;; esac',
      'echo "Ready to send $todo emails by $how."',
      '[ "$todo" -lt "$total" ] && echo "$((total - todo)) were sent earlier and are skipped."',
      'echo "First: $first"',
      'echo "Press Ctrl+C at any time to stop. Run it again later to continue where it stopped."',
      'if [ -z "$PAYSLIP_YES" ]; then',
      '  read -r -p "Type YES to send $todo emails: " answer',
      '  [ "$(printf \'%s\' "$answer" | tr \'[:lower:]\' \'[:upper:]\')" = "YES" ] || { echo "Nothing was sent."; finish; exit 0; }',
      'fi',
      '[ -f "$LOG" ] || echo "timestamp,employee_id,name,channel,contact,file,sha256,note,action,source" > "$LOG"',
      '',
      'i=0; ok=0; failed=0',
      'while IFS=$\'\\t\' read -r -u 3 id name email subject body file source; do',
      '  grep -qxF "$id" "$SENT" && continue',
      '  i=$((i + 1))',
      '  if send_one "$email" "$subject" "$body" "pdf/$file"; then',
      '    hash=$(shasum -a 256 "pdf/$file" 2>/dev/null | cut -d" " -f1)',
      '    [ -n "$hash" ] || hash=$(sha256sum "pdf/$file" | cut -d" " -f1)',
      '    stamp=$(date -u +%Y-%m-%dT%H:%M:%S.000Z)',
      '    echo "$(csv "$stamp"),$(csv "$id"),$(csv "$name"),mail,$(csv "$email"),$(csv "$file"),$hash,sent by $how (send-all pack),sent,$(csv "$source")" >> "$LOG"',
      '    echo "$id" >> "$SENT"',
      '    ok=$((ok + 1))',
      '    echo "[$i/$todo] Sent to $name <$email>"',
      '  else',
      '    failed=$((failed + 1))',
      '    echo "[$i/$todo] FAILED $name <$email>: $ERR"',
      '    echo "$(date +%Y-%m-%dT%H:%M:%S)	$id	$name	$email	$ERR" >> errors.txt',
      '    hint=$(route_hint "$ERR")',
      '    if [ -n "$hint" ] && [ "$ok" -eq 0 ]; then echo "$hint"; echo "Stopped before trying the others."; break; fi',
      '  fi',
      '  [ "$i" -lt "$todo" ] && sleep "$DELAY"',
      'done 3< <(tail -n +2 "$LIST")',
      'echo',
      'echo "Done: $ok sent, $failed failed."',
      '[ "$ROUTE" = "mail" ] && echo "Mail sends them from its Outbox; keep Mail open until it is empty."',
      'echo "Then load delivery-log.csv on the Send step (More > Restore ticks from a log) to tick everyone."',
      '[ "$failed" -gt 0 ] && echo "The failed ones are listed in errors.txt. Fix them and run Send all again: the others are skipped."',
      'finish',
      ''
    ].join('\n');
  }

  function macSendOneAppleScript() {
    return [
      '-- Payslip Generator: one email with its PDF through Apple Mail.',
      '-- arguments: to, subject, body (line breaks written as \\n), PDF path, sender (may be empty)',
      'on run argv',
      '\tset toAddr to item 1 of argv',
      '\tset subj to item 2 of argv',
      '\tset bodyText to my swapText(item 3 of argv, "\\\\n", linefeed)',
      '\tset pdfFile to (POSIX file (item 4 of argv)) as alias',
      '\tset fromAddr to item 5 of argv',
      '\ttell application "Mail"',
      '\t\tset msg to make new outgoing message with properties {subject:subj, content:bodyText & linefeed & linefeed, visible:false}',
      '\t\tif fromAddr is not "" then',
      '\t\t\ttry',
      '\t\t\t\tset sender of msg to fromAddr',
      '\t\t\tend try',
      '\t\tend if',
      '\t\ttell msg',
      '\t\t\tmake new to recipient at end of to recipients with properties {address:toAddr}',
      '\t\t\ttell content',
      '\t\t\t\tmake new attachment with properties {file name:pdfFile} at after the last paragraph',
      '\t\t\tend tell',
      '\t\tend tell',
      '\t\tdelay 2 -- let Mail finish attaching before sending',
      '\t\tsend msg',
      '\tend tell',
      'end run',
      '',
      'on swapText(t, a, b)',
      '\tset AppleScript\'s text item delimiters to a',
      '\tset parts to text items of t',
      '\tset AppleScript\'s text item delimiters to b',
      '\tset t to parts as text',
      '\tset AppleScript\'s text item delimiters to ""',
      '\treturn t',
      'end swapText',
      ''
    ].join('\n');
  }

  function macCommand(mode, route) {
    return '#!/bin/bash\ncd "$(dirname "$0")" && ' + (route ? 'PAYSLIP_ROUTE=' + route + ' ' : '') + 'exec bash send-all.sh ' + mode + '\n';
  }

  function sendPackGuideHtml(info, platform) {
    var win = platform === 'windows';
    var app = win ? 'classic Outlook' : 'Apple Mail';
    var ext = win ? '.bat' : '.command';
    var run = function (name) { return win ? '<code>' + name + '.bat</code>' : '<code>' + name + '.command</code> (Linux: <code>bash send-all.sh ' + ({ '1 - Send test to me': 'test', '2 - Send all': 'send', 'Check the pack': 'check' }[name] || 'send') + '</code>)'; };
    return '<!doctype html><html lang="en"><head><meta charset="utf-8"><title>Send all payslips</title>' +
      '<style>body{font-family:Segoe UI,-apple-system,Arial,sans-serif;max-width:760px;margin:40px auto;line-height:1.55;color:#1f2933;padding:0 16px}h1{font-size:1.5rem}h2{font-size:1.1rem;margin-top:28px}code{background:#eef2f6;border-radius:4px;padding:2px 6px}li{margin:6px 0}table{border-collapse:collapse;width:100%}td,th{border:1px solid #d9e0e7;padding:6px 10px;text-align:left;vertical-align:top}.note{background:#fff6d6;border:1px solid #f0dc9a;border-radius:8px;padding:10px 14px}</style></head><body>' +
      '<h1>Send all ' + info.count + ' payslips</h1>' +
      '<p>' + esc(info.company) + ' · ' + esc(info.period) + '. Every email goes from your own mailbox with the employee\'s PDF attached' + (info.locked ? ' (locked with the password rule in the message)' : '') + '. Nothing else leaves this computer.</p>' +
      '<h2>How it sends (it chooses by itself)</h2><table><tr><th>If this computer has…</th><th>It sends with…</th></tr>' +
      '<tr><td>' + (win ? 'Classic Outlook (with a File menu)' : 'Apple Mail with your email account (Mac)') + '</td><td>' + app + ', no questions.</td></tr>' +
      '<tr><td>' + (win ? 'New Outlook, Outlook on the web, or no Outlook' : 'No Apple Mail account, or Linux') + ', and a Microsoft 365 or Outlook.com address</td><td>A <b>Microsoft sign-in</b>: it shows a code, your browser opens Microsoft\'s page, you sign in and accept "send mail as you". No password is typed in the window.</td></tr>' +
      '<tr><td>A Gmail or Google Workspace address</td><td>Google\'s mail server with an <b>app password</b> (create one at myaccount.google.com/apppasswords; 2-Step Verification must be on).</td></tr>' +
      '<tr><td>Any other email (Zoho, GoDaddy, Hostinger, company mail)</td><td>Your email provider\'s <b>mail server</b> with your email password. It suggests the server; your provider\'s help pages list it as "SMTP server".</td></tr></table>' +
      '<p>Passwords and sign-ins are used for this run only and never saved. ' + run('Check the pack') + ' tells you which way it will use.</p>' +
      '<h2>Steps</h2><ol>' +
      '<li><b>Unzip</b> this pack into a folder (' + (win ? 'right-click → Extract All' : 'double-click the ZIP') + '). Keep the files together.</li>' +
      '<li>Double-click ' + run('Check the pack') + '. It lists everyone and says how it will send.</li>' +
      '<li>Double-click ' + run('1 - Send test to me') + '. You receive the first employee\'s email. Check it.</li>' +
      '<li>Double-click ' + run('2 - Send all') + '. It shows how many emails and to whom, and waits for you to type <b>YES</b>.</li>' +
      '<li>Leave the window open. It sends one email every few seconds (about 24 a minute) and lists each one.</li>' +
      '<li>When it says Done, open the payslip page, go to <b>Send → More → Restore ticks from a log</b> and choose <code>delivery-log.csv</code> from this folder. Everyone sent is ticked.</li></ol>' +
      '<h2>If you stop it</h2><p>Press <b>Ctrl+C</b> or close the window. Run ' + run('2 - Send all') + ' again later: everyone already sent is skipped.</p>' +
      (win ? '<p>Prefer to check every email first? <code>Save as drafts instead.bat</code> puts them in the classic Outlook Drafts folder. <code>Send without Outlook.bat</code> skips Outlook and uses the sign-in or the mail server, for when Outlook misbehaves.</p>'
        : '<p><code>Send without Apple Mail.command</code> skips Apple Mail and uses the sign-in or the mail server, for when Mail is not set up the way you want.</p>') +
      '<h2>If it does not start</h2><ul>' +
      (win
        ? '<li><b>Windows protected your PC</b>: click <b>More info → Run anyway</b>.</li><li><b>Running scripts is disabled</b> or the window closes at once: your IT department blocks scripts. Use the buttons on the payslip page instead (one email at a time).</li>'
        : '<li><b>Cannot be opened because it is from an unidentified developer</b>: right-click the file → <b>Open</b> → <b>Open</b>. Needed once.</li><li><b>Terminal wants to control Mail</b>: click <b>OK</b>. If you clicked Don\'t Allow: System Settings → Privacy &amp; Security → Automation → Terminal → turn on Mail.</li><li><b>Linux</b>: open a terminal in this folder and run <code>bash send-all.sh</code> (needs curl).</li>') +
      '</ul><h2>If a way does not work</h2><ul>' +
      (win ? '<li><b>Classic Outlook would not answer (80080005)</b>: Outlook and the script run with different rights. Close Outlook, open it normally (not as administrator), and run it again.</li>' +
        '<li><b>Class not registered (80040154)</b>: it retries with 32-bit PowerShell by itself; if it still fails, repair Office (Settings → Apps → Microsoft 365 / Office → Modify → Quick Repair), or use <code>Send without Outlook.bat</code>.</li>' +
        '<li>Outlook says <b>a program is trying to send email</b>: choose Allow. Blocked by <b>Programmatic Access</b>: update the antivirus, or ask IT.</li>' : '') +
      '<li><b>Microsoft sign-in: administrator approval needed</b>: your company lets only IT approve "Microsoft Graph Command Line Tools" (Microsoft\'s own app for scripts). Ask IT to approve it for you, or use ' + (win ? 'classic Outlook' : 'Apple Mail') + '.</li>' +
      '<li><b>Microsoft sign-in blocked by your company\'s rules</b>: ask IT, or use ' + (win ? 'classic Outlook' : 'Apple Mail') + '.</li>' +
      '<li><b>Password not accepted</b>: Gmail and accounts with two-step sign-in need an <b>app password</b>. Microsoft addresses do not take passwords here any more (Microsoft turned that off); they use the Microsoft sign-in.</li>' +
      '<li><b>Mail server could not be reached</b>: check the internet connection and the server name; the port is 587.</li>' +
      '<li>Still stuck: the buttons on the payslip page send one email at a time on any computer or phone.</li>' +
      '</ul><h2>Sending limits</h2><ul><li>Microsoft 365: about 30 emails a minute and 10,000 recipients a day.</li><li>Gmail: 500 a day; Google Workspace: 2,000 a day.</li></ul>' +
      '<p class="note">Delete this folder after sending: it contains the payslips.</p></body></html>';
  }


  return {
    buildEml: buildEml,
    textToHtml: textToHtml,
    encodeHeader: encodeHeader,
    base64Lines: base64Lines,
    waLink: waLink,
    gmailComposeLink: gmailComposeLink,
    mailtoLink: mailtoLink,
    outlookWebLink: outlookWebLink,
    toCsv: toCsv,
    manifestCsv: manifestCsv,
    mailMergeCsv: mailMergeCsv,
    deliveryLogCsv: deliveryLogCsv,
    mailMergeGuideHtml: mailMergeGuideHtml,
    sendListCsv: sendListCsv,
    sendListTsv: sendListTsv,
    windowsSendScript: windowsSendScript,
    windowsBat: windowsBat,
    macSendScript: macSendScript,
    macSendOneAppleScript: macSendOneAppleScript,
    macCommand: macCommand,
    sendPackGuideHtml: sendPackGuideHtml
  };
});
