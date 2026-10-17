# Sending ~1,000 personalised payslip emails with per-employee PDF attachments through the employer's existing mail client (no SMTP config, no API, no server)

Scope: files produced by an offline browser page on a Windows/Mac office PC; delivery handed to Outlook (classic / new / web / Mac), Gmail in the browser, Thunderbird or Apple Mail. State as of October 2026. Audience: HR/accounts clerk of a small Saudi employer on Microsoft 365 or Gmail.

Summary scorecard (details and sources in the sections below):

| Route | Carries per-employee PDF? | Setup effort | Clicks per employee | Free? | Works in new Outlook? | Works on Mac? |
|---|---|---|---|---|---|---|
| (1) `mailto:` links | **No** (RFC 6068 forbids MIME/attachment headers) | none | 3+ (click link, attach PDF manually, Send) | yes | yes | yes |
| (2) Generated `.eml` with `X-Unsent: 1` | yes (base64 MIME part) | low (generator writes files) | 2 (double-click, Send) | yes | yes, but intermittently broken 2024–2026 and needs a `Message-ID` line | **No** (Outlook for Mac, Apple Mail, Thunderbird all open as read-only) |
| (3) Word mail merge + Merge Tools add-in / Mail Merge Toolkit | yes (column in data source) | medium (Word+Excel+classic Outlook) | 0 (bulk) | Merge Tools: free; MAPILab: ~US$24 | **No** (needs classic Outlook, VBA) | No |
| (4) Outlook VBA / PowerShell COM loop | yes | medium (macro security, maybe one prompt per message) | 0 if trusted, else 1 "Allow" click + 5 s wait per message | yes | **No** (no COM/VBA/MAPI) | No |
| (5) Thunderbird Mail Merge add-on | yes (`{{Attachment}}` column; TB ≥102 needs files pre-selected via Browse) | medium (install TB, add account, CSV) | 0 (bulk) | yes (donation-ware) | n/a | yes |
| (6) new Outlook / OWA | only via .eml route or drag-drop drafts | — | — | — | — | — |
| (7) Gmail web | native mail merge = one attachment for all; compose URL = no attachments | — | 3+ manual | yes | — | yes |

Provider limits (all routes, because the mailbox is the sender): Exchange Online 10,000 recipients/day and 30 messages/minute; Outlook.com with M365 subscription 5,000 recipients/day (free tier unpublished, ~300); Gmail free 500 messages/day; Google Workspace 2,000/day (1,500 for native mail merge, 500 for trial). 1,000 one-to-one messages in one go is inside the Microsoft 365 and Workspace numbers but **over** free Gmail and the Outlook.com free tier, and at 30 msg/min takes ≥34 minutes on Exchange Online.

---

## Key question 1: What exactly can a `mailto:` URL carry, and what are the practical length limits in Outlook, the Gmail web handler and default Windows mail apps?

### Takeaway
A `mailto:` URI can pre-fill To/Cc/Bcc/Subject/Body (and arbitrary header fields the client is told to distrust) but by specification cannot carry an attachment or any MIME structure; practical length is ~2,000 characters for safety (Chrome-on-Windows misbehaves at 2,048; Gmail's handler ≈4,096 after encoding). It is therefore useless for 1,000 PDF payslips except as a "pre-addressed message, attach the PDF yourself" convenience.

### Cited Findings
- RFC 6068 §2: addresses go in the path; `cc`, `bcc`, `subject`, `in-reply-to` and any `hfname=hfvalue` pair are allowed; the `to` hfname is "NOT RECOMMENDED because different user agents handle this case differently". §2 says `body` is "primarily intended for the generation of short text messages", no transfer encoding is applied to it, and "Content-Transfer-Encoding and other message-encoding headers in a mailto URI are irrelevant and MUST be ignored". — [RFC 6068](https://datatracker.ietf.org/doc/html/rfc6068)
- RFC 6068 §3: originator fields (From, Date), routing/trace fields and MIME fields (MIME-Version, Content-*) MUST be ignored by the client; §4: a user agent SHOULD NOT create a message if any header field is dangerous and can only be assumed to understand Subject and body; §7: clients SHOULD NOT send without showing the user the full decoded message and getting approval. Hence no standards-based way to pass an attachment. — [RFC 6068](https://datatracker.ietf.org/doc/html/rfc6068)
- RFC 6068 §2/§5/§6.3: non-ASCII (e.g. Arabic) must be UTF-8 then percent-encoded (`subject=caf%C3%A9`); line breaks in body MUST be `%0D%0A`; `?`, `=`, `&` must be escaped; spaces SHOULD be `%20`. — [RFC 6068](https://datatracker.ietf.org/doc/html/rfc6068)
- Gmail-as-handler: "The limit is ~4,096 characters for Gmail - this count includes the subject, recipients, and body of the intended email (after URI encoding)"; the project chose a 2,500-char warning threshold because clients differ. No figures given for Outlook/Apple Mail/Thunderbird. — [codeforboston/maple issue #1964](https://github.com/codeforboston/maple/issues/1964)
- Browser side: mailto handling "starts misbehaving on Windows 10 / Chrome 73 once the mailto url character count reaches 2048", while Firefox on Windows handles much longer mailto URLs. — [personaldata.io wiki](https://wiki.personaldata.io/wiki/Special:MobileDiff/2578)
- The "2083" figure often quoted comes from Internet Explorer's URL maximum, not from a mailto spec; "most browsers support lengths up to around 2000 characters". — [Anaplan community](https://community.anaplan.com/discussion/comment/45350/)
- Outlook's own hyperlink length limit (for links inside messages, not mailto handling) was 2084 characters and "the fix ... increases the URL limit from 2084 to 8192 characters" from Outlook Desktop build 16116.10000, requiring both Outlook and Windows updates. — [Microsoft Support, "Cannot send this item"](https://support.microsoft.com/en-us/support/known-issues/outlook-error-cannot-send-this-item-when-replying-to-an-email); Sophos separately states "Microsoft Outlook does not support URLs longer than 2048 characters" — [Sophos community](https://community.sophos.com/sophos-email/b/blog/posts/enhanced-time-of-click-protection) (conflicting; neither is specific to mailto).
- Gmail compose deep-link alternative (`https://mail.google.com/mail/?view=cm&fs=1&to=...&su=...&body=...`) is undocumented, pre-fills to/cc/bcc/su/body only, and nothing indicates it accepts attachments. — [Simon Willison TIL](https://til.simonwillison.net/google/gmail-compose-url); [Alec Jacobson](https://www.cs.toronto.edu/~jacobson/weblog/210.html)

### Inferences
- For 1,000 payslips a mailto page would still require the clerk to drag each PDF into each message by hand: at least three actions per employee plus finding the right file, i.e. several hours of error-prone work and the main risk of sending the wrong payslip to the wrong person.
- Keep any generated mailto URL under ~2,000 encoded characters; Arabic text triples in length after UTF-8 percent-encoding (each Arabic letter = 2 bytes = 6 characters), so an Arabic body of ~300 characters already approaches the limit.

### Gaps
- No Microsoft documentation states a mailto-specific maximum for classic/new Outlook or the (now retired) Windows Mail app; the figures above are community measurements.
- No official Google documentation on the Gmail mailto handler length limit; ~4,096 is a developer report.

---

## Key question 2: Does double-clicking a generated `.eml` produce an editable, sendable draft in classic Outlook, new Outlook, Outlook for Mac, Apple Mail, Thunderbird, Windows Mail? What headers make it a draft? Can 1,000 be opened in bulk?

### Takeaway
Only Outlook for Windows treats an `.eml` as a draft, and only when the non-standard header `X-Unsent: 1` is present (required since classic Outlook 2109, Oct 2021); new Outlook additionally needs a `Message-ID` and has had two regressions (May 2024 and Sep 2025–Jan 2026). Outlook for Mac, Apple Mail and Thunderbird open `.eml` as received/read-only mail ("Edit as New" in Thunderbird is the manual workaround). There is no documented bulk "open 1,000 .eml as drafts" mechanism; each file is one window.

### Cited Findings
- Classic Outlook for Windows: from Version 2109 (Build 16.0.14430.20256) `.eml` files that previously opened as editable drafts (Outlook 2010–2019) began opening "as read-only sent messages that can only be replied to or forwarded"; the accepted fix (2021-10-18) is to add the header `X-Unsent: 1`. A macOS user (2021-12-27) reported `X-Unsent` did not work in Outlook for macOS, with no answer. — [Microsoft Q&A 4562029](https://learn.microsoft.com/en-us/answers/questions/4562029/eml-files-no-longer-opening-as-a-draft-e-mail-to-s)
- New Outlook, 2024: report of 2024-05-29 that with `X-Unsent: 1` "the message opens, but the user cannot edit or send it"; a Microsoft Community Moderator replied the same day that "this isn't currently feasible in New Outlook for Windows" and pointed to the feedback portal. — [Microsoft Q&A 4635248](https://learn.microsoft.com/en-us/answers/questions/4635248/header-x-unsent-for-eml-files-is-not-recognized-in)
- New Outlook, 2025: a user builds `.eml` templates with `X-Unsent: 1` as the first line; "Opening one gives an editable draft, as intended" but a given file can be sent only once; a Microsoft Q&A moderator reproduced it (2025-09-23) and found the fix is to delete the `Message-ID:` line from the file, after which it can be sent repeatedly. Same workflow worked in classic Outlook with `.oft` files. — [Microsoft Q&A 5562698](https://learn.microsoft.com/en-au/answers/questions/5562698/new-outlook-eml-x-unsent-1-sending-not-working-mul)
- New Outlook, Sep 2025 regression: newly generated `.eml` with `X-Unsent: 1` failed to save to Drafts ("This message can't be saved right now"); a Microsoft Q&A reply on 2025-11-11 said Outlook now requires a valid `Message-ID` and the header alone no longer triggers draft saving; users worked around by adding `Message-ID: <unique-id@xxxx.outlook.com>` next to `X-Unsent: 1`; the poster edited on 2026-01-20 that it appears fixed in Outlook Version 1.2026.105.100 / Client 20260109002.06. No official acknowledgement. — [Microsoft Q&A 5594291](https://learn.microsoft.com/en-us/answers/questions/5594291/new-eml-files-will-not-save-to-draft-folder-(fixed)
- Microsoft's support page for new Outlook / Outlook on the web describes only *opening* `.eml/.msg/.oft` (drag onto the reading pane, right-click > Open with > Outlook (new), or double-click when it is the default app) and notes `.msg/.oft` need "a qualifying Microsoft 365 subscription"; it says nothing about editing or sending. — [Microsoft Support](https://support.microsoft.com/en-us/office/open-eml-msg-and-oft-files-in-new-outlook-60f71e69-e9a5-445b-b4dd-2e0d5aaf21d6)
- Dragging `.eml` files into new Outlook "just opens each message in its own window and does not add them to a mailbox"; no bulk import method. — [Microsoft Q&A 1483732](https://learn.microsoft.com/en-gb/answers/questions/1483732/eml-file-import-into-the-new-outlook-for-windows-()
- Outlook for Mac: "the issue is actually with Outlook for Mac which doesn't care for the unsent flag" (vendor support reply). — [Aspose forum](https://forum.aspose.com/t/saving-eml-as-unsent-not-working-on-mac/17643)
- Apple Mail: `X-Unsent: 1` is not honoured; the developer found drafts carry `Uniform-Type-Identifier: com.apple.mail-draft`, but even a saved draft re-opened from disk "is not unsent anymore". The header "is not a part of the standard, and e-mail clients don't have to support this". — [Rebex forum](https://forum.rebex.net/9290/save-email-as-unsent-using-eml-in-apple-mail)
- Thunderbird: "the X-Unsent flag set in EML is only understandable by MS Outlook"; Thunderbird opens `.eml` as a received message and the user must use Edit as New (Ctrl-E); a request for a command-line switch to open directly as a new message appears unresolved. — [Aspose forum](https://forum.aspose.com/t/open-eml-file-in-edit-mode-for-ibm-notes-and-thunderbird/160296); [Mozilla support question](https://support.mozilla.org/de/questions/1022415)
- eM Client (not in scope but shows the header is client-specific) added `X-Unsent` support so `.eml` loads in compose mode. — [eM Client forum](https://forum.emclient.com/t/add-support-for-x-unsent-header/37519)
- Windows Mail: the service behind the built-in Mail/Calendar/People apps was shut down on 31 December 2024; users are moved to new Outlook. — [macmyths.com](https://macmyths.com/farewell-mail-and-calendar-apps-how-to-move-safely-to-outlook-on-windows-10-and-11/) (third-party; Microsoft announcement not fetched)
- Message-format basis: RFC 2047 encoded-words `=?UTF-8?B?...?=` are required for non-ASCII Subject/display names (each encoded-word ≤75 chars, header lines ≤76 chars; not allowed inside addr-spec or quoted-string). — [RFC 2047](https://datatracker.ietf.org/doc/html/rfc2047)

### Inferences
- A generator that targets Windows Outlook users should emit, per employee, an `.eml` whose first lines are `X-Unsent: 1`, a unique `Message-ID: <uuid@employer-domain>` (needed by new Outlook since late 2025; harmless in classic), `MIME-Version: 1.0`, RFC 2047-encoded `Subject`, `To`, a `multipart/mixed` body with a `text/plain; charset=utf-8` part and an `application/pdf` part with `Content-Transfer-Encoding: base64` and `Content-Disposition: attachment; filename="..."`. This is the only "hand to the client" route that needs zero add-ins, but it costs 2 clicks per employee (open, Send) = ~2,000 clicks for 1,000 staff, and the clerk must watch for the new-Outlook "sent once only / can't save" regressions.
- Because Outlook for Mac, Apple Mail and Thunderbird all open `.eml` read-only, Mac offices need route (5) Thunderbird Mail Merge or manual forwarding; there is no draft-from-file path on macOS.
- "Bulk open": selecting 1,000 `.eml` files and pressing Enter would spawn 1,000 inspector windows; nothing suggests Outlook queues them to Drafts. Practical batches are 20–50 at a time.

### Gaps
- No Microsoft documentation defines `X-Unsent`; all behaviour is inferred from community/Q&A threads and may change with any monthly new-Outlook build.
- Did not find any report that confirms or denies draft behaviour for `.eml` in Outlook on the web (browser) when a file is dragged in; Microsoft's page only covers opening.
- Apple Mail's `com.apple.mail-draft` route was tried and failed by one developer; no working recipe found.

---

## Key question 3: Minimal working Outlook VBA / PowerShell snippet to create emails from a CSV + PDF folder, and which Microsoft security prompts or admin policies block it in 2025–2026 (object model guard, macro settings, new Outlook lacking COM)?

### Takeaway
A ~20-line loop using `Application.CreateItem(olMailItem)`, `.Attachments.Add path`, then `.Save`/`.Display`/`.Send` works in classic Outlook for Windows; the blockers are (a) Outlook's default macro setting that disables unsigned macros, (b) the Object Model Guard "send message" prompt (5-second timer, one Allow per message) when the code runs out-of-process (PowerShell/Excel) or antivirus status is not "Good", and (c) new Outlook, which supports no VBA, COM or MAPI at all.

### Cited Findings
- `Attachments.Add(Source, Type, Position, DisplayName)`: Source is "the full file system path with a file name"; Microsoft's own example is `Set myItem = Application.CreateItem(olMailItem)` / `myItem.Attachments.Add "C:\Test.doc", olByValue, 1, "Test"` / `myItem.Display`. Remark: "always save an item before adding or removing objects in the Attachments collection" for consistent results. — [Microsoft Learn, Attachments.Add](https://learn.microsoft.com/en-us/office/vba/api/outlook.attachments.add)
- Object Model Guard: three prompts (address book, send message, execute action). Send Message Warning: "invoked when an untrusted solution attempts to send an item programmatically ... The user must wait five seconds before clicking Allow"; Allow permits "only that call" so every subsequent `Send` prompts again; Deny returns `MAPI_E_NOT_SUPPORTED`. The address-book prompt (triggered e.g. by reading `.To`/recipients) has an "Allow access for N minutes" option that then covers all callers. — [Microsoft Learn, OOM Security Warnings](https://learn.microsoft.com/en-us/office/vba/outlook/how-to/security/outlook-object-model-security-warnings)
- When prompts are suppressed: in-process COM add-ins are trusted by default; cross-process callers (PowerShell, Excel VBA, scripts) run without warnings only if Windows Security Center reports antivirus "Good", and Outlook uses default settings or a Group Policy without a programmatic-access policy. Trust Center > Programmatic Access offers "Warn me ... when my antivirus software is inactive or out-of-date (recommended)" (default), "Always warn me", "Never warn me (not recommended)"; only local administrators can change it; Group Policy can lock it. Objects created via `CreateObject("Outlook.Application")` are "not trusted". — [Microsoft Learn, Security Behavior of the Outlook Object Model](https://learn.microsoft.com/en-us/office/vba/outlook/how-to/security/security-behavior-of-the-outlook-object-model)
- Macro setting default in classic Outlook: "Notifications for digitally signed macros, all other macros disabled" (older label "Warnings for signed macros; all unsigned macros are disabled"); path File > Options > Trust Center > Trust Center Settings > Macro Settings, must be checked from the main Outlook window. — [Tenable audit of Microsoft 365 Apps baseline](https://tenable.com/audits/items/MSCT_Microsoft_365_Apps_for_enterprise_2306_v1.0.0.audit:07395353b305ccf2597fc17f9c92ef09); [robert365.com self-signing walkthrough](https://robert365.com/article/selfcert); [DISA STIG Outlook 2016 V-228464](https://www.stigviewer.com/stigs/microsoft_outlook_2016/2022-03-11/finding/V-228464)
- Slipstick's `Application_ItemSend` pattern in `ThisOutlookSession` attaches a per-recipient file by display name — `.Attachments.Add "D:\For merge\" & Item.To & ".docx"` — triggered by a `[merge]` codeword in the subject; the author calls display-name matching "prone to problems"; readers report Outlook asks to enable macros after each restart. — [Slipstick, Mail merge messages](https://www.slipstick.com/developer/code-samples/mail-merge-messages/)
- Microsoft DevBlogs shows PowerShell driving Outlook via the interop assembly to attach a PDF to a template email for each recipient and send. — [Microsoft DevBlogs, Outlook email automation with PowerShell](https://devblogs.microsoft.com/premier-developer/outlook-email-automation-with-powershell/)
- New Outlook: "VBA and Macros aren't supported in new Outlook" (alternatives: Power Automate, Microsoft Graph, Office.js). — [Microsoft Learn, VBA alternatives](https://learn.microsoft.com/en-us/microsoft-365-apps/outlook/get-started/vba-alternatives); "the new Outlook for Windows doesn't support COM add-ins" but supports web add-ins — [Microsoft Learn, State of COM add-ins](https://learn.microsoft.com/zh-tw/microsoft-365-apps/outlook/get-started/state-of-com-add-ins); VSTO/COM still supported in classic Outlook — [Microsoft Learn, One Outlook](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/one-outlook); a Microsoft Q&A moderator (2026) states new Outlook "does not support COM, VBA, or MAPI automation" — [Microsoft Q&A 5951650](https://learn.microsoft.com/en-in/answers/questions/5951650/need-a-support-in-new-outlook-automation); Excel VBA `CreateObject("Outlook.Application")` fails to target new Outlook — [Microsoft Q&A 5286064](https://learn.microsoft.com/en-us/answers/questions/5286064/excel-vba-macros-new-outlook)

### Example code (written for this report from the documented API above; not executed here)

VBA, paste into a module in classic Outlook (Alt+F11), run `SendPayslips`. CSV columns: `Email,Name,File` where `File` is the PDF filename.

```vb
Sub SendPayslips()
    Const CSV = "C:\Payslips\list.csv", PDFDIR = "C:\Payslips\pdf\"
    Dim f As Integer, line As String, p() As String, m As Outlook.MailItem
    f = FreeFile: Open CSV For Input As #f
    Line Input #f, line                       ' skip header
    Do While Not EOF(f)
        Line Input #f, line: p = Split(line, ",")
        If Dir(PDFDIR & p(2)) <> "" Then
            Set m = Application.CreateItem(olMailItem)   ' in-process: trusted
            m.To = p(0)
            m.Subject = "Payslip " & Format(Date, "mmmm yyyy") & " - " & p(1)
            m.Body = "Dear " & p(1) & "," & vbCrLf & "Your payslip is attached."
            m.Save                                       ' per Attachments.Add remarks
            m.Attachments.Add PDFDIR & p(2)
            m.Save            ' leaves it in Drafts for review; use m.Send to send
        End If
    Loop
    Close #f
End Sub
```

PowerShell (runs *out-of-process*, so the Object Model Guard applies unless antivirus status is "Good" and policy allows):

```powershell
$ol = New-Object -ComObject Outlook.Application      # classic Outlook only
Import-Csv 'C:\Payslips\list.csv' | ForEach-Object {
    $pdf = Join-Path 'C:\Payslips\pdf' $_.File
    if (Test-Path $pdf) {
        $m = $ol.CreateItem(0)                         # 0 = olMailItem
        $m.To = $_.Email
        $m.Subject = "Payslip $(Get-Date -Format 'MMMM yyyy') - $($_.Name)"
        $m.Body = "Dear $($_.Name),`r`nYour payslip is attached."
        $m.Save(); [void]$m.Attachments.Add($pdf)
        $m.Save()                                      # or $m.Send()
        Start-Sleep -Seconds 2                         # stay under 30 msg/min
    }
}
```

### Inferences
- Running the loop *inside* Outlook VBA (not from Excel/PowerShell) avoids the Object Model Guard on `.Send` because in-process code uses the trusted `Application` object; the remaining hurdle is the macro setting, solved by self-signing the VBA project (SelfCert) or temporarily choosing "Enable all macros" (admin-controlled on many corporate PCs).
- `.Save` (leave in Drafts) instead of `.Send` gives the clerk a review step and sidesteps the per-message send prompt; the clerk then selects all drafts and sends, which still goes through Exchange's 30 msg/min throttle.
- `.Send` from PowerShell on a PC without "Good" antivirus status means 1,000 prompts x 5 seconds = ≥83 minutes of clicking; CDO/Redemption workarounds are out of scope (they require extra components).
- Any organisation that has already switched users to new Outlook (or whose admin has removed classic Outlook) cannot use this route at all.

### Gaps
- Microsoft's OOM security pages are dated 2019 and describe behaviour that still applies to classic Outlook, but no 2025–2026 Microsoft page was found restating them; treat defaults as unchanged.
- Microsoft has not published a firm date when classic Outlook stops being available for Microsoft 365 commercial tenants (the Learn pages fetched do not state one); flag as deprecated-trajectory.

---

## Key question 4: Which free mail-merge-with-attachment tools exist for Outlook and Thunderbird, and are they still maintained? Why can't Word mail merge attach files natively?

### Takeaway
Word's "Merge to E-Mail" has no attachment field at all; the practical fixes are Doug Robbins' free Merge Tools Add-in (Word VBA, Windows + classic Outlook only), the free GPLv3 OMMA script (sends keystrokes, last tested Office 2013), the paid MAPILab Mail Merge Toolkit (~US$24/user), and on Thunderbird the actively maintained free Mail Merge add-on by Alexander Bergmann (v11.5.0, 2026-03-31) with a `{{Attachment}}` column.

### Cited Findings
- Word/Outlook native limitation: "While Microsoft Outlook supports mail merge, it does not support merging attachments with the message. You'll need to either use one of the mail merge tools ... or use VBA." — [Slipstick, Mail merge using Outlook](https://www.slipstick.com/outlook/mail-merge-using-outlook); Outlook-tips: "two possible solutions if you need to send an attachment- use a mail merge utility that is designed to send attachments or use VBA" — [Outlook-tips tip 392](https://www.outlook-tips.net/tips/tip-392-mail-merge-attachments)
- Merge Tools Add-in (Doug Robbins, Word MVP): features "Mail Merge with Attachments, CC, BCC and custom Subject", merge to individual Word/PDF files named from a data field, many-to-one; distributed as a zip from mergetoolsaddin.com (or by emailing mergetools@gmail.com). — [mergetoolsaddin.com](https://www.mergetoolsaddin.com/); Microsoft Answers threads point to its "Merge with Attachments" dialog, where attachments come "from a column in the data source, or one file can be added to every message"; Robbins stated it "cannot be used on a Mac because Visual Basic on a Mac is not equivalent to the Windows implementation"; one user reported a PC lock-up after downloading from a non-official mirror. — [Microsoft Q&A 5040737](https://learn.microsoft.com/en-us/answers/questions/5040737/mail-merge-attachment); [Microsoft Q&A 4885883](https://learn.microsoft.com/en-us/answers/questions/4885883/adding-attachments-in-mail-merge-in-word-2016?page=2); [Kevin Stratvert tutorial](https://kevinstratvert.com/2021/08/18/how-to-mail-merge-with-attachments-custom-subject-cc-bcc-using-word-excel-outlook/)
- OMMA (Outlook Mail Merge Attachment): free, GPLv3, v1.1.9 beta by Wouter Westerveld; workflow = set Outlook offline, run Word merge so messages sit in Outbox, run the .vbs to attach the same file(s) to every Outbox message; "tested on Microsoft Office 2003, 2007, 2010 and 2013"; it works by sending keystrokes (ALT+S), so non-English Outlook and any keyboard/mouse use during processing break it; the author notes MergeTools "provides better integration ... and also the option to send personalized attachments". — [omma.sourceforge.net](https://omma.sourceforge.net); [troubleshooting](https://omma.sourceforge.net/troubleshooting.php)
- MAPILab Mail Merge Toolkit: listings show starting price US$24 per user, perpetual licence with one year of updates, free trial; supports personalised attachments including password-protected PDF/DOCX; a 2017 review found it failed when a digital-signing certificate was active in Outlook. — [Software Advice listing](https://www.softwareadvice.co.uk/software/319051/mail-merge-toolkit); [TechByter review 2017](https://techbyter.com/2017/0827.html); [Techjockey listing](https://techjockey.com/detail/mapilab-mail-merge)
- Thunderbird Mail Merge add-on (Alexander Bergmann): version 11.1.0 (2025-06-10) added Thunderbird 140 support; 11.5.0 (2026-03-31) added Thunderbird 150 support and is listed compatible with TB 128 ESR, 140 ESR, 151 release, 152 beta; 11.4.0 added a warning that merged messages are never OpenPGP/S-MIME signed or encrypted; developer notes a TB 151 bug when re-using a saved draft/template that already has attachments (workaround: compose fresh or re-add attachments; 140 ESR unaffected); and that "legacy add-ons using MailExtension Experiments will no longer run in the Thunderbird release channel starting with version 153" (ESR 140 unaffected). — [addons.thunderbird.net version history](https://addons.thunderbird.net/hu/thunderbird/addon/mail-merge/versions); [addon listing](https://reviewers.addons.thunderbird.net/af/thunderbird/addon/mail-merge/)
- Per-row attachments in Thunderbird Mail Merge: CSV with an `ATTACHMENT` column holding a filename per recipient; in the dialog, Source = CSV, Attachments field = folder path + `{{ATTACHMENT}}`; Start/Stop line numbers (header = line 1); Pause in seconds between mails (author used 7–10 s); Deliver mode "Send Now" / "Send Later" (queues in Outbox) / "Save as Draft". — [FSFE wiki, Personalised mass-mailing with Thunderbird](https://wiki.fsfe.org/TechDocs/MailMerge) (2018 page)
- Since Thunderbird 102 the attachment field only accepts the filename (e.g. `MyFile.pdf`) and the files must first be selected with the Browse button, "caused by a limitation of the new MailExtension APIs". — [Mozilla support thread](https://support.mozilla.org/ro/questions/1325172) (summary via search; page itself failed to load during this research)
- Other: Mail Merge P (siefkenj) is a fork "based off of Mail Merge by Alexander Bergmann". — [GitHub MailMergeP](https://github.com/siefkenj/MailMergeP)

### Inferences
- Effort ranking for a clerk: Thunderbird Mail Merge is the only free, maintained, cross-platform, zero-clicks-per-employee option, but requires installing Thunderbird and adding the work account (Microsoft 365 via OAuth2 IMAP/SMTP, or Gmail with IMAP enabled) and, on TB ≥102, multi-selecting all 1,000 PDFs in the Browse dialog so that `{{Attachment}}` can resolve filenames.
- Merge Tools Add-in is free and purpose-built but Windows + Word + classic Outlook only, macro-dependent (same Trust Center hurdles as Q3) and distributed as a zip by an individual MVP, which many IT departments will not allow.
- OMMA cannot do per-recipient attachments (same file to every Outbox message) and is effectively unmaintained; exclude it for payslips.

### Gaps
- Could not load the Thunderbird add-on's main listing page (HTTP 403) to confirm licence text and exact Thunderbird 153 wording; the version-history and reviewer mirror pages were used.
- No current (2024+) confirmation that Merge Tools Add-in works with the latest Microsoft 365 Word builds; Microsoft Answers threads still recommend it in 2024–2025 but give no version test.
- Mail Merge Toolkit's current official price page was not fetched; US$24 is from reseller listings.

---

## Key question 5: Daily sending limits for Microsoft 365 business mailboxes, Outlook.com, Gmail free and Google Workspace; do 1,000 messages in one go trip throttling?

### Takeaway
1,000 one-recipient messages fit within Exchange Online (10,000 recipients/day, throttled to 30 messages/minute, so ≥34 min) and Google Workspace (2,000/day; 1,500 if using Gmail's native mail merge), but exceed free Gmail (500/day) and the unpublished free Outlook.com quota (~300); Outlook.com with a Microsoft 365 subscription allows 5,000/day but only 1,000 "non-relationship" recipients/day.

### Cited Findings
- Exchange Online (all Microsoft 365 Business Basic/Standard/Premium, E3/E5, E1, F3): Recipient rate limit 10,000 recipients per day (24-hour rolling); Recipient limit per message "Customizable up to 1000"; Message rate limit 30 messages per minute (excess "throttled and successively carried over to the following minutes ... Exchange Online isn't suited to accommodate bulk-mailing scenarios"); distribution groups in the GAL count as one recipient, personal lists count members individually; "Exchange Online customers who need to send legitimate bulk commercial email ... should use third-party providers". — [Microsoft Learn, Exchange Online limits (ms.date 2026-04-07)](https://learn.microsoft.com/en-us/office365/servicedescriptions/exchange-online-service-description/exchange-online-limits)
- Tenant External Recipient Rate Limit (TERRL): tenant-wide cap on external recipients per 24 h, depends on licence count, trial tenants capped at 5,000/day; visible in EAC Reports > Mail flow; Azure Communication Services recommended above it. Separately, mail sent from the default `*.onmicrosoft.com` domain is limited to 100 external recipients per organisation per 24 h (NDR 550 5.7.236). — [Microsoft Learn, Exchange Online limits](https://learn.microsoft.com/en-us/office365/servicedescriptions/exchange-online-service-description/exchange-online-limits)
- Per-mailbox External Recipient Rate limit (2,000/24 h as a sublimit of 10,000): Message Center MC787382 announced it with phased dates (trial/new tenants first, existing tenants later; the June 2025 revision said 1 Apr 2026 and 1 Oct 2026), but the latest revision of MC787382 states "we have decided not to proceed with the rollout of this limit". University IT pages still quote April 2026 — treat as outdated. — [MC787382 (merill.net archive)](https://mc.merill.net/message/MC787382); [HKU ITS notice](https://its.hku.hk/?p=57870) (conflicting/outdated)
- Outlook.com with Microsoft 365 subscription: 5,000 daily recipients, 500 recipients per message, 1,000 "daily non-relationship recipients"; non-subscribers have lower, unpublished limits that vary with usage history; new accounts have a temporary "low sending quota". — [Microsoft Support, Sending limits in Outlook.com](https://support.microsoft.com/en-us/office/sending-limits-in-outlook-com-279ee200-594c-40f0-9ec8-bb6af7735c2e); third-party estimate for free accounts ~300 recipients/day, 100 per message — [Mailmeteor](https://mailmeteor.com/blog/outlook-send-limit)
- Google Workspace (per user, rolling 24 h): 2,000 messages/day (1,500 for mail merge; 500 for trial accounts); 10,000 total recipients/day (1,500 for mail merge); 3,000 external recipients/day; 3,000 unique recipients/day (2,000 external; 500 external for trial); 2,000 recipients per message (max 500 external); 100 per message via SMTP relay by POP/IMAP users; when exceeded "sending is blocked for up to 24 hours"; trial limits lift only after the domain has cumulatively paid ≥US$100, up to 75 days later. — [Google Workspace Admin Help, Gmail sending limits](https://knowledge.workspace.google.com/admin/gmail/gmail-sending-limits-in-google-workspace)
- Gmail free: 500 messages/day; limits are per account and shared across all clients/add-ons/SMTP use. — [Mailmeteor sending-limits guide](https://mailmeteor.com/blog/email-sending-limits); [Labnol](https://labnol.org/docs/20309-gmail-message-not-sent)
- Google native mail merge (Workspace Individual/Business/Enterprise): shares the account's 2,000 limit, mail-merge sends stop at 1,500/day; the same attachment goes to every recipient. — [Google Workspace Individual Help](https://support.google.com/google-workspace-individual/answer/12921167); [REMC1 guide mirror](https://remc1.atlassian.net/wiki/spaces/PS/pages/695402504/Gmail+Mail+Merge)

### Inferences
- Microsoft 365 tenant, internal employee addresses on the company domain: 1,000 payslips count as 1,000 internal recipients — no external-limit exposure, but a VBA `.Send` loop or "send all drafts" will be metered at 30/min. Sending from the `onmicrosoft.com` default domain to employees' personal Gmail/Hotmail addresses would hit the 100-external-recipient tenant cap on the first ~100 messages.
- Gmail free account (common for very small Saudi employers): 1,000 payslips must be split over 2–3 days or the account is blocked for 24 h mid-run; a Workspace account with trial status (500/day) has the same problem.
- Thunderbird Mail Merge's Pause setting (e.g. 3–7 s) keeps Exchange under 30/min and reduces Gmail heuristics risk.

### Gaps
- Microsoft publishes no number for free Outlook.com accounts; ~300/day is vendor/community data.
- Whether 1,000 distinct one-to-one messages with identical-looking PDF attachments trigger Exchange Online Protection or Gmail spam heuristics (as opposed to the published quotas) is not documented by either provider; Google only says heuristics exist and to follow sender guidelines.

---

## Key question 6: Saudi/GCC-specific considerations (Arabic subject lines, employees without email, company domains)

### Takeaway
Arabic subjects and names are handled by UTF-8 percent-encoding in mailto and RFC 2047 encoded-words in `.eml` headers (Outlook, Gmail, Thunderbird all decode these); no source was found that imposes a Saudi legal requirement to deliver payslips by email, and the Wage Protection System (WPS/Mudad) governs salary payment, not payslip delivery, so the practical issue is that many blue-collar employees have no corporate email address.

### Cited Findings
- `.eml` headers: non-ASCII Subject and display-name text must be RFC 2047 encoded-words (`=?UTF-8?B?...?=`), max 75 chars per encoded-word, 76 per line, never inside the `addr-spec`; MIME Content-Type/Content-Disposition parameters (e.g. Arabic PDF filenames) may not use encoded-words (RFC 2231 applies instead). — [RFC 2047](https://datatracker.ietf.org/doc/html/rfc2047)
- `mailto:` with Arabic: "Non-ASCII characters MUST first be encoded according to UTF-8" then percent-encoded (§2); encoded-words are allowed in hfvalues but not in `body`. — [RFC 6068](https://datatracker.ietf.org/doc/html/rfc6068)
- WPS/Mudad: every private-sector employer must pay salaries through Mudad-approved banking channels and submit a monthly Salary Information File at least one business day before payday; employers must be registered on Mudad, Qiwa, Muqeem and GOSI; cash payment is not permitted. None of these sources mention emailing payslips. — [Setup in Saudi, WPS overview](https://www.setupinsaudi.com/en/news/announcements/understanding-the-wage-protection-system-wps-in-saudi-arabia); [Mercans, WPS rules](https://mercans.com/?p=12439); [Al Rajhi Bank WPS](https://www.alrajhibank.com.sa/en/Business/Cash-Management/payroll/WPS)
- Vendor guidance describes an electronic salary receipt as relying on "a certified digital signature plus digital approval in the employee self-service portal" but cites no regulation. — [Qoyod salary receipt template](https://academy.qoyod.com/en/templates/salary-receipt/)
- Google's Workspace sending-limits page is published in Arabic (hl=ar), i.e. the same quotas apply to Saudi tenants. — [Google Workspace Admin Help (ar)](https://knowledge.workspace.google.com/admin/gmail/gmail-sending-limits-in-google-workspace?hl=ar)

### Inferences
- Generator should write `.eml` files as UTF-8 throughout (RFC 2047 B-encoded Subject, `Content-Type: text/plain; charset=utf-8` body, ASCII-safe PDF filename such as `payslip_2026-09_EMP001.pdf` plus optional RFC 2231 `filename*=UTF-8''...` for an Arabic display name) to avoid mojibake in Outlook classic, which is the most charset-sensitive of the clients.
- Right-to-left rendering of the body is a client concern; plain-text parts render in Outlook/Gmail with auto-direction, but a `text/html` part with `dir="rtl"` is safer for mixed Arabic/English payslip text (inference, no source).
- Employees without email addresses (common for site/driver/domestic categories) cannot be served by any email route; the CSV should flag missing/invalid addresses before the run so the clerk can print or hand-deliver those payslips, and the per-recipient limits above mean sending to personal Gmail/Hotmail addresses counts as *external* recipients for TERRL/onmicrosoft.com caps.
- Company-domain hygiene: a Saudi SME sending from `company.onmicrosoft.com` rather than a verified custom domain is limited to 100 external recipients/day tenant-wide (Microsoft Learn, Q5), which would block payslips to employees' personal mailboxes.

### Gaps
- No MHRSD regulation, Labour Law article or Mudad documentation on payslip delivery format was found in this research; vendor pages only. Report-writer should not assert a legal requirement either way.
- No data found on the share of Saudi private-sector employees without an email address; the "many lack email" point is an inference from the workforce profile, not a cited statistic.
- Arabic-interface Outlook (RTL UI) specifics for the `.eml` draft behaviour were not tested or documented anywhere found.

---

## Key question 7: New Outlook (Monarch) limitations and Gmail-web "multiple drafts" tricks

### Takeaway
New Outlook has no VBA/COM/MAPI, no add-in route for local automation, and only partially honours `.eml` drafts (needs `X-Unsent: 1` plus a `Message-ID`, with regressions); Gmail's browser UI offers no attachment-capable deep link and its native mail merge sends one shared attachment, so per-employee PDFs in Gmail web require either manual compose per employee or a Workspace add-on (outside the no-API constraint).

### Cited Findings
- "VBA and Macros aren't supported in new Outlook" — [Microsoft Learn](https://learn.microsoft.com/en-us/microsoft-365-apps/outlook/get-started/vba-alternatives); no COM add-ins — [Microsoft Learn](https://learn.microsoft.com/zh-tw/microsoft-365-apps/outlook/get-started/state-of-com-add-ins); no COM/VBA/MAPI automation per Microsoft Q&A moderator — [Microsoft Q&A 5951650](https://learn.microsoft.com/en-in/answers/questions/5951650/need-a-support-in-new-outlook-automation)
- New Outlook `.eml`: opens via double-click/drag; draft editing depends on `X-Unsent: 1` and `Message-ID` (see Q2 citations: Q&A 4635248, 5562698, 5594291); users have an open feature request "Requesting Functionality to Open, Edit, and Send .eml or .msg Files in 'The New Outlook'". — [Microsoft Q&A 4620562](https://learn.microsoft.com/en-us/answers/questions/4620562/requesting-functionality-to-open-edit-and-send-eml)
- Gmail compose URL (`view=cm&fs=1&to&su&body`) is undocumented, pre-fills text fields only, and may need URI-escaping; no attachment parameter. — [Simon Willison TIL](https://til.simonwillison.net/google/gmail-compose-url); [Airtable community formula example](https://community.airtable.com/fid-10/tid-32785)
- Gmail native "mail merge"/multi-send: 1,500/day cap within 2,000; same attachment to all recipients; "no limit to the number of unique recipients you can contact per month". — [Google Workspace Individual Help](https://support.google.com/google-workspace-individual/answer/12921167); per-recipient files need third-party add-ons (e.g. Digital Inspiration Mail Merge, paid tier for dynamic Drive attachments) — [Digital Inspiration pricing](https://digitalinspiration.com/docs/mail-merge/pricing)
- Classic Outlook can still be reached by toggling off "new Outlook"; Microsoft positions web add-ins/Graph/Power Automate as the replacements (all outside this report's "no API" constraint). — [Microsoft Learn, VBA alternatives](https://learn.microsoft.com/en-us/microsoft-365-apps/outlook/get-started/vba-alternatives)

### Inferences
- For a Gmail-in-browser office the only no-API route is manual: open Gmail, click Compose (or a generated compose link that pre-fills address/subject/body), drag the named PDF from a file-explorer window, Send — about 4 actions per employee; a generator can at least sort PDFs and produce a checklist to reduce mis-sends.
- In new Outlook the `.eml` route is the only hands-off-generator option; it should be tested on the exact build in use the morning of the payroll run, with the fallback of toggling to classic Outlook if the admin still allows it.

### Gaps
- No documentation found for whether Outlook on the web (OWA in browser) can turn a dragged `.eml` into a draft; Microsoft's page covers "open" only.
- Gmail's browser UI "multiple drafts" tricks beyond the compose URL (e.g. Chrome protocol-handler registration for mailto) were not located in official documentation during this research.
