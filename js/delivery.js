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

  // Windows, classic Outlook: PowerShell driving Outlook on this computer (no macros, no add-ins).
  function windowsSendScript(info) {
    return '\uFEFF' + [
      '# Payslip Generator - sends every payslip email through classic Outlook on this computer.',
      '# Start it with "2 - Send all.bat" (try "1 - Send test to me.bat" first). Nothing leaves this',
      '# computer except the emails themselves, sent from your own Outlook.',
      'param(',
      '  [switch]$Test,      # one email to yourself with the first payslip; nothing is logged',
      '  [switch]$Drafts,    # save every email in Outlook Drafts instead of sending',
      '  [switch]$Check,     # only check the pack; Outlook is not needed',
      '  [switch]$Yes,       # no questions (for automated checks)',
      '  [double]$Delay = 2.5,',
      '  $Outlook = $null',
      ')',
      "$ErrorActionPreference = 'Stop'",
      '$Company = ' + psString(info.company),
      '$Period = ' + psString(info.period),
      '$Sender = ' + psString(info.sender),
      '$here = Split-Path -Parent $MyInvocation.MyCommand.Path',
      "$listPath = Join-Path $here 'send-list.csv'",
      "$logPath = Join-Path $here 'delivery-log.csv'",
      "$errPath = Join-Path $here 'errors.csv'",
      "$pdfDir = Join-Path $here 'pdf'",
      '',
      'function Say($text, $color) { if ($color) { Write-Host $text -ForegroundColor $color } else { Write-Host $text } }',
      "function Finish { if (-not $Yes) { Write-Host ''; [void](Read-Host 'Press Enter to close') } }",
      'function CsvCell($v) { $s = [string]$v; if ($s -match \'[",\\r\\n]\') { return \'"\' + $s.Replace(\'"\', \'""\') + \'"\' } return $s }',
      'function CsvLine($values) { return (($values | ForEach-Object { CsvCell $_ }) -join \',\') }',
      '',
      "Say ''",
      "Say ('Payslips ' + $Company + ' - ' + $Period) Cyan",
      'if (-not (Test-Path $listPath)) { Say \'send-list.csv is missing. Keep all files of this pack together.\' Red; Finish; exit 1 }',
      '$rows = @(Import-Csv -Path $listPath -Encoding UTF8)',
      "if ($rows.Count -eq 0) { Say 'send-list.csv has no employees.' Red; Finish; exit 1 }",
      '$missing = @($rows | Where-Object { -not (Test-Path (Join-Path $pdfDir $_.file)) })',
      '',
      '$done = @{}',
      'if (Test-Path $logPath) {',
      "  Import-Csv -Path $logPath -Encoding UTF8 | ForEach-Object { if ($_.action -eq 'sent') { $done[$_.employee_id] = $true } }",
      '}',
      '$todo = @($rows | Where-Object { -not $done.ContainsKey($_.employee_id) })',
      '',
      'if ($Check) {',
      '  foreach ($r in $rows) {',
      "    $state = if ($done.ContainsKey($r.employee_id)) { 'already sent' } elseif (Test-Path (Join-Path $pdfDir $r.file)) { 'ready' } else { 'PDF MISSING' }",
      "    Say ('  ' + $r.name + ' <' + $r.email + '>  ' + $r.file + '  [' + $state + ']')",
      '  }',
      "  Say ('' + $rows.Count + ' employees, ' + $todo.Count + ' to send, ' + $missing.Count + ' PDF missing. Nothing was sent.') Cyan",
      '  Finish; exit 0',
      '}',
      'if ($missing.Count -gt 0) {',
      "  Say ('These PDF files are missing in the pdf folder: ' + (($missing | ForEach-Object { $_.file }) -join ', ')) Red",
      '  Finish; exit 1',
      '}',
      '',
      'if (-not $Outlook) {',
      "  try { $Outlook = New-Object -ComObject Outlook.Application }",
      "  catch { Say 'Classic Outlook was not found on this computer. Use the Outlook button on the page instead (New Outlook and Outlook on the web are not supported by this pack).' Red; Finish; exit 1 }",
      '}',
      "$ns = $Outlook.GetNamespace('MAPI')",
      '$account = $null',
      '$me = $null',
      'foreach ($a in $ns.Accounts) {',
      '  if (-not $me) { $me = $a.SmtpAddress }',
      '  if ($Sender -and $a.SmtpAddress -eq $Sender) { $account = $a; $me = $a.SmtpAddress }',
      '}',
      '',
      'if ($Test) {',
      '  $r = $rows[0]',
      "  if (-not $me) { $me = Read-Host 'Your email address for the test' }",
      '  $mail = $Outlook.CreateItem(0)',
      '  $mail.To = $me',
      "  $mail.Subject = '[TEST] ' + $r.subject",
      "  $mail.Body = 'TEST: this is what ' + $r.name + ' <' + $r.email + '> will receive.' + \"`r`n`r`n\" + $r.body",
      '  [void]$mail.Attachments.Add((Join-Path $pdfDir $r.file))',
      '  if ($account) { $mail.SendUsingAccount = $account }',
      '  $mail.Send()',
      '  try { $ns.SendAndReceive($false) } catch { }',
      "  Say ('Test email sent to ' + $me + '. Check it in your inbox, then run \"2 - Send all\".') Green",
      '  Finish; exit 0',
      '}',
      '',
      "$verb = if ($Drafts) { 'save as drafts' } else { 'send' }",
      "if ($todo.Count -eq 0) { Say 'Everyone in this pack has been sent already.' Green; Finish; exit 0 }",
      "Say ('Ready to ' + $verb + ' ' + $todo.Count + ' emails' + $(if ($me) { ' from ' + $me } else { '' }) + '.')",
      "if ($done.Count -gt 0) { Say ('' + $done.Count + ' were sent earlier and are skipped.') }",
      "Say ('First: ' + $todo[0].name + ' <' + $todo[0].email + '>')",
      "Say 'Press Ctrl+C at any time to stop. Run it again later to continue where it stopped.' DarkGray",
      'if (-not $Yes) {',
      "  $answer = Read-Host ('Type YES to ' + $verb + ' ' + $todo.Count + ' emails')",
      "  if ($answer -ne 'YES') { Say 'Nothing was sent.'; Finish; exit 0 }",
      '}',
      "if (-not $Drafts -and -not (Test-Path $logPath)) { Set-Content -Path $logPath -Value 'timestamp,employee_id,name,channel,contact,file,sha256,note,action,source' -Encoding UTF8 }",
      '',
      '$i = 0; $ok = 0; $failed = 0',
      'foreach ($r in $todo) {',
      '  $i++',
      '  $file = Join-Path $pdfDir $r.file',
      "  $tag = '[' + $i + '/' + $todo.Count + '] '",
      '  try {',
      '    $mail = $Outlook.CreateItem(0)',
      '    $mail.To = $r.email',
      '    $mail.Subject = $r.subject',
      '    $mail.Body = $r.body',
      '    [void]$mail.Attachments.Add($file)',
      '    if ($account) { $mail.SendUsingAccount = $account }',
      '    if ($Drafts) {',
      '      $mail.Save()',
      "      Say ($tag + 'Draft saved for ' + $r.name + ' <' + $r.email + '>') Green",
      '    } else {',
      '      $mail.Send()',
      '      $hash = (Get-FileHash -Algorithm SHA256 -Path $file).Hash.ToLower()',
      "      $stamp = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ss.fffZ')",
      "      Add-Content -Path $logPath -Value (CsvLine @($stamp, $r.employee_id, $r.name, 'outlook', $r.email, $r.file, $hash, 'sent by Outlook (send-all pack)', 'sent', $r.source)) -Encoding UTF8",
      "      Say ($tag + 'Sent to ' + $r.name + ' <' + $r.email + '>') Green",
      '    }',
      '    $ok++',
      '  } catch {',
      '    $failed++',
      "    Say ($tag + 'FAILED ' + $r.name + ' <' + $r.email + '>: ' + $_.Exception.Message) Red",
      "    Add-Content -Path $errPath -Value (CsvLine @((Get-Date).ToString('s'), $r.employee_id, $r.name, $r.email, $_.Exception.Message)) -Encoding UTF8",
      '  }',
      '  if ($i -lt $todo.Count -and $Delay -gt 0) { Start-Sleep -Milliseconds ([int]($Delay * 1000)) }',
      '}',
      'try { $ns.SendAndReceive($false) } catch { }',
      "Say ''",
      'if ($Drafts) {',
      "  Say ('Done: ' + $ok + ' drafts saved in Outlook, ' + $failed + ' failed. Check the Drafts folder and send them from Outlook.') Cyan",
      '} else {',
      "  Say ('Done: ' + $ok + ' sent, ' + $failed + ' failed.') Cyan",
      "  Say 'Keep Outlook open until the Outbox is empty.'",
      "  Say 'Then load delivery-log.csv on the Send step (More > Restore ticks from a log) to tick everyone.'",
      '}',
      "if ($failed -gt 0) { Say 'The failed ones are listed in errors.csv. Fix them and run Send all again: the others are skipped.' Yellow }",
      'Finish'
    ].join('\r\n') + '\r\n';
  }

  function windowsBat(args, title) {
    return ['@echo off', 'title ' + title, 'powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0send-all.ps1"' + (args ? ' ' + args : ''), ''].join('\r\n');
  }

  // Mac, Apple Mail: a shell script that hands each email to Mail through AppleScript.
  function macSendScript(info) {
    return [
      '#!/bin/bash',
      '# Payslip Generator - sends every payslip email through Apple Mail on this Mac.',
      '# Start it with "2 - Send all.command" (try "1 - Send test to me.command" first).',
      'cd "$(dirname "$0")" || exit 1',
      'MODE="${1:-send}"',
      'OSA="${PAYSLIP_OSASCRIPT:-osascript}"',
      'DELAY="${PAYSLIP_DELAY:-2.5}"',
      'COMPANY=' + shString(info.company),
      'PERIOD=' + shString(info.period),
      'SENDER=' + shString(info.sender),
      'LIST="send-list.tsv"',
      'LOG="delivery-log.csv"',
      'SENT="sent-ids.txt"',
      '',
      'finish() { echo; [ -n "$PAYSLIP_YES" ] || read -r -p "Press Enter to close" _; }',
      'csv() { local s="$1"; case "$s" in *[\\",]*) s="${s//\\"/\\"\\"}"; printf \'"%s"\' "$s" ;; *) printf \'%s\' "$s" ;; esac; }',
      '',
      'echo',
      'echo "Payslips $COMPANY - $PERIOD"',
      '[ -f "$LIST" ] || { echo "send-list.tsv is missing. Keep all files of this pack together."; finish; exit 1; }',
      'touch "$SENT"',
      '',
      'if [ "$MODE" = "test" ]; then',
      '  ME=$("$OSA" -e \'tell application "Mail" to get item 1 of (email addresses of item 1 of accounts)\' 2>/dev/null)',
      '  [ -n "$ME" ] || read -r -p "Your email address for the test: " ME',
      '  IFS=$\'\\t\' read -r id name email subject body file source < <(sed -n 2p "$LIST")',
      '  if "$OSA" send-one.applescript "$ME" "[TEST] $subject" "TEST: this is what $name <$email> will receive.\\n\\n$body" "$PWD/pdf/$file" "$SENDER" >/dev/null; then',
      '    echo "Test email sent to $ME. Check it in your inbox, then run \\"2 - Send all\\"."',
      '  else',
      '    echo "The test could not be sent. Is Apple Mail set up with your email account?"',
      '  fi',
      '  finish; exit 0',
      'fi',
      '',
      'total=0; todo=0; first=""',
      'while IFS=$\'\\t\' read -r -u 3 id name email subject body file source; do',
      '  total=$((total + 1))',
      '  if ! grep -qxF "$id" "$SENT"; then',
      '    todo=$((todo + 1))',
      '    [ -n "$first" ] || first="$name <$email>"',
      '    [ -f "pdf/$file" ] || { echo "Missing PDF: pdf/$file"; finish; exit 1; }',
      '  fi',
      'done 3< <(tail -n +2 "$LIST")',
      'if [ "$todo" -eq 0 ]; then echo "Everyone in this pack has been sent already."; finish; exit 0; fi',
      'echo "Ready to send $todo emails with Apple Mail."',
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
      '  if "$OSA" send-one.applescript "$email" "$subject" "$body" "$PWD/pdf/$file" "$SENDER" >/dev/null 2>>errors.txt; then',
      '    hash=$(shasum -a 256 "pdf/$file" 2>/dev/null | cut -d" " -f1)',
      '    [ -n "$hash" ] || hash=$(sha256sum "pdf/$file" | cut -d" " -f1)',
      '    stamp=$(date -u +%Y-%m-%dT%H:%M:%S.000Z)',
      '    echo "$(csv "$stamp"),$(csv "$id"),$(csv "$name"),mail,$(csv "$email"),$(csv "$file"),$hash,sent by Apple Mail (send-all pack),sent,$(csv "$source")" >> "$LOG"',
      '    echo "$id" >> "$SENT"',
      '    ok=$((ok + 1))',
      '    echo "[$i/$todo] Sent to $name <$email>"',
      '  else',
      '    failed=$((failed + 1))',
      '    echo "[$i/$todo] FAILED $name <$email> (see errors.txt)"',
      '  fi',
      '  [ "$i" -lt "$todo" ] && sleep "$DELAY"',
      'done 3< <(tail -n +2 "$LIST")',
      'echo',
      'echo "Done: $ok sent, $failed failed."',
      'echo "Mail sends them from its Outbox; keep Mail open until it is empty."',
      'echo "Then load delivery-log.csv on the Send step (More > Restore ticks from a log) to tick everyone."',
      '[ "$failed" -gt 0 ] && echo "Run Send all again after fixing the failed ones: the others are skipped."',
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

  function macCommand(mode) {
    return '#!/bin/bash\ncd "$(dirname "$0")" && exec bash send-all.sh ' + mode + '\n';
  }

  function sendPackGuideHtml(info, platform) {
    var win = platform === 'windows';
    var app = win ? 'classic Outlook' : 'Apple Mail';
    var ext = win ? '.bat' : '.command';
    return '<!doctype html><html lang="en"><head><meta charset="utf-8"><title>Send all payslips with ' + app + '</title>' +
      '<style>body{font-family:Segoe UI,-apple-system,Arial,sans-serif;max-width:760px;margin:40px auto;line-height:1.55;color:#1f2933;padding:0 16px}h1{font-size:1.5rem}h2{font-size:1.1rem;margin-top:28px}code{background:#eef2f6;border-radius:4px;padding:2px 6px}li{margin:6px 0}.note{background:#fff6d6;border:1px solid #f0dc9a;border-radius:8px;padding:10px 14px}</style></head><body>' +
      '<h1>Send all ' + info.count + ' payslips with ' + app + '</h1>' +
      '<p>' + esc(info.company) + ' · ' + esc(info.period) + '. Every email goes from your own ' + (win ? 'Outlook' : 'Mail') + ' with the employee\'s PDF attached' + (info.locked ? ' (locked with the password rule in the message)' : '') + '. Nothing else leaves this computer.</p>' +
      '<h2>Steps</h2><ol>' +
      '<li><b>Unzip</b> this pack into a folder (' + (win ? 'right-click → Extract All' : 'double-click the ZIP') + '). Keep the files together.</li>' +
      (win ? '<li>Open <b>classic Outlook</b> and leave it open.</li>' : '<li>Make sure <b>Apple Mail</b> is set up with your email account.</li>') +
      '<li>Double-click <code>1 - Send test to me' + ext + '</code>. You receive the first employee\'s email. Check it.</li>' +
      '<li>Double-click <code>2 - Send all' + ext + '</code>. It shows how many emails and to whom, and waits for you to type <b>YES</b>.</li>' +
      '<li>Leave the window open. It sends one email every few seconds (about 24 a minute) and lists each one.</li>' +
      '<li>When it says Done, open the payslip page, go to <b>Send → More → Restore ticks from a log</b> and choose <code>delivery-log.csv</code> from this folder. Everyone sent is ticked.</li></ol>' +
      '<h2>If you stop it</h2><p>Press <b>Ctrl+C</b> or close the window. Run <code>2 - Send all' + ext + '</code> again later: everyone already sent is skipped.</p>' +
      (win ? '<p>Prefer to check every email first? Double-click <code>Save as drafts instead.bat</code>: the emails go to the Outlook Drafts folder and you press Send in Outlook. <code>Check the pack.bat</code> lists everything without sending.</p>' : '') +
      '<h2>If it does not start</h2><ul>' +
      (win
        ? '<li><b>Windows protected your PC</b>: click <b>More info → Run anyway</b>.</li><li><b>Running scripts is disabled</b> or nothing happens: your IT department blocks scripts. Use the Outlook button on the page instead.</li><li>Outlook says <b>a program is trying to send email</b>: choose Allow (this appears when the antivirus is out of date).</li><li><b>New Outlook</b> or Outlook on the web: this pack needs classic Outlook. Use the Outlook button on the page.</li>'
        : '<li><b>Cannot be opened because it is from an unidentified developer</b>: right-click the file → <b>Open</b> → <b>Open</b>. Needed once.</li><li><b>Terminal wants to control Mail</b>: click <b>OK</b>. Needed once.</li><li>Outlook for Mac is not supported by this pack; use Apple Mail, or the Outlook button on the page.</li>') +
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
