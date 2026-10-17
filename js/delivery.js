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
    lines.push('--' + boundary);
    lines.push('Content-Type: text/plain; charset="UTF-8"');
    lines.push('Content-Transfer-Encoding: base64');
    lines.push('');
    lines.push(base64Lines(utf8Bytes(String(opts.text || '').replace(/\r?\n/g, CRLF))));
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
    var rows = [['timestamp', 'employee_id', 'name', 'channel', 'contact', 'file', 'sha256', 'note']];
    entries.forEach(function (e) {
      rows.push([e.timestamp, e.employeeId, e.name, e.channel, e.contact || '', e.file || '', e.hash || '', e.note || '']);
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

  function outlookGuideHtml(settings, batchSize) {
    return '<!doctype html><html lang="en"><head><meta charset="utf-8"><title>How to send the payslips from Outlook</title>' +
      '<style>body{font-family:Segoe UI,Arial,sans-serif;max-width:760px;margin:40px auto;line-height:1.55;color:#1f2933;padding:0 16px}h1{font-size:1.5rem}h2{font-size:1.1rem;margin-top:28px}code,pre{background:#eef2f6;border-radius:4px;padding:2px 6px}pre{padding:12px;white-space:pre-wrap}li{margin:6px 0}.note{background:#fff6d6;border:1px solid #f0dc9a;border-radius:8px;padding:10px 14px}</style></head><body>' +
      '<h1>Outlook drafts (.eml files)</h1>' +
      '<p>' + esc(settings.companyName || '') + ' · ' + esc(settings.payPeriod || '') + '. Each <code>.eml</code> file is a ready email with the payslip attached. This works with <b>classic Outlook for Windows</b> (Microsoft 365, Outlook 2021/2019). The new Outlook, Outlook for Mac and web mail open these files read-only.</p>' +
      '<h2>Option A: open and send, ' + batchSize + ' at a time</h2><ol><li>Open one <code>batch-XX</code> folder.</li><li>Select all files (Ctrl+A) and press <b>Enter</b>. Outlook opens each one as a draft with the attachment.</li><li>Check the address, press <b>Send</b> (Ctrl+Enter) on each window.</li><li>Repeat for the next batch. Outlook sends about 30 messages per minute.</li></ol>' +
      '<h2>Option B: save everything to Drafts with the macro</h2><ol><li>In Outlook press <b>Alt+F11</b>, <b>Insert → Module</b>, paste the content of <code>OutlookDrafts.bas</code> (or File → Import File).</li><li>Change the folder path at the top of the macro to this folder.</li><li>Press <b>F5</b>. All messages are created in <b>Drafts</b> with the right address and attachment; nothing is sent.</li><li>Review the Drafts folder and press Send on each, or change <code>.Save</code> to <code>.Send</code> in the macro to send directly.</li></ol>' +
      '<p class="note">If Outlook says macros are disabled: File → Options → Trust Center → Trust Center Settings → Macro Settings → "Notifications for all macros", restart Outlook and choose Enable when asked.</p>' +
      '</body></html>';
  }

  // Classic Outlook VBA: reads manifest.csv next to the PDFs and saves one draft per row.
  function outlookVbaModule(settings) {
    var subject = String(settings.messageSubject || 'Payslip {period} - {company}');
    return [
      'Attribute VB_Name = "PayslipDrafts"',
      "' Payslip Generator - creates one Outlook draft per row of manifest.csv (same folder as the PDFs).",
      "' 1. Set FolderPath below to the folder that holds manifest.csv and the PDF files.",
      "' 2. Press F5. Drafts appear in the Drafts folder; nothing is sent until you press Send.",
      "'    To send immediately, replace .Save with .Send.",
      'Option Explicit',
      '',
      'Sub CreatePayslipDrafts()',
      '    Dim FolderPath As String',
      '    FolderPath = "C:\\Payslips\\' + Schema.periodSlug(settings.payPeriod, settings.periodStart) + '\\"',
      '',
      '    Dim fso As Object, ts As Object, line As String, cols() As String',
      '    Dim mail As Object, created As Long, skipped As Long, first As Boolean',
      '    Set fso = CreateObject("Scripting.FileSystemObject")',
      '    If Not fso.FileExists(FolderPath & "manifest.csv") Then',
      '        MsgBox "manifest.csv not found in " & FolderPath, vbExclamation',
      '        Exit Sub',
      '    End If',
      '    Set ts = fso.OpenTextFile(FolderPath & "manifest.csv", 1, False, -1)',
      '    first = True',
      '    Do Until ts.AtEndOfStream',
      '        line = ts.ReadLine',
      '        If first Then',
      '            first = False',
      "        ElseIf Len(Trim(line)) > 0 Then",
      '            cols = SplitCsv(line)',
      "            ' columns: 0 employee_id, 1 name, 2 email, 3 mobile, 4 file, 5 sha256, 6 net, 7 currency, 8 period, 9 password_hint",
      '            If UBound(cols) >= 4 And Len(cols(2)) > 0 And fso.FileExists(FolderPath & cols(4)) Then',
      '                Set mail = Application.CreateItem(0)',
      '                mail.To = cols(2)',
      '                mail.Subject = "' + subject.replace(/"/g, '""').replace('{period}', '" & cols(8) & "').replace('{company}', String(settings.companyName || '').replace(/"/g, '""')).replace('{name}', '" & cols(1) & "') + '"',
      '                mail.Body = "Dear " & cols(1) & "," & vbCrLf & "Your payslip for " & cols(8) & " is attached. " & cols(9) & vbCrLf & "This document contains personal data; please do not forward it."',
      '                mail.Attachments.Add FolderPath & cols(4)',
      '                mail.Save',
      '                created = created + 1',
      '            Else',
      '                skipped = skipped + 1',
      '            End If',
      '        End If',
      '    Loop',
      '    ts.Close',
      '    MsgBox created & " drafts created, " & skipped & " rows skipped (no email or file missing).", vbInformation',
      'End Sub',
      '',
      "' Minimal CSV splitter that understands double quotes.",
      'Private Function SplitCsv(ByVal s As String) As String()',
      '    Dim result() As String, i As Long, ch As String, cur As String, inQ As Boolean, n As Long',
      '    ReDim result(0 To 0)',
      '    For i = 1 To Len(s)',
      '        ch = Mid$(s, i, 1)',
      '        If inQ Then',
      '            If ch = """" Then',
      '                If i < Len(s) And Mid$(s, i + 1, 1) = """" Then',
      '                    cur = cur & """": i = i + 1',
      '                Else',
      '                    inQ = False',
      '                End If',
      '            Else',
      '                cur = cur & ch',
      '            End If',
      '        ElseIf ch = """" Then',
      '            inQ = True',
      '        ElseIf ch = "," Then',
      '            ReDim Preserve result(0 To n): result(n) = cur: n = n + 1: cur = ""',
      '        Else',
      '            cur = cur & ch',
      '        End If',
      '    Next i',
      '    ReDim Preserve result(0 To n): result(n) = cur',
      '    SplitCsv = result',
      'End Function',
      ''
    ].join(CRLF);
  }

  return {
    buildEml: buildEml,
    encodeHeader: encodeHeader,
    base64Lines: base64Lines,
    waLink: waLink,
    gmailComposeLink: gmailComposeLink,
    mailtoLink: mailtoLink,
    toCsv: toCsv,
    manifestCsv: manifestCsv,
    mailMergeCsv: mailMergeCsv,
    deliveryLogCsv: deliveryLogCsv,
    mailMergeGuideHtml: mailMergeGuideHtml,
    outlookGuideHtml: outlookGuideHtml,
    outlookVbaModule: outlookVbaModule
  };
});
