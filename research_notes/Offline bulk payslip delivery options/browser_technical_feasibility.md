# Browser-side technical feasibility: handing ~1,000 generated payslip PDFs to the employer's own tools (no backend)

Scope: static single-page app (vanilla JS, jsPDF, JSZip, SheetJS), opened from `file://` (USB / shared drive) or a static HTTPS host (Vercel). State as of Oct 2026 unless noted. Context discovered locally: the project currently vendors **jsPDF 2.5.1 and 2.5.2 UMD builds** (both contain the `userPermissions` encryption code path), **JSZip 3.10.1** and SheetJS `xlsx.full.min.js` (observed in the session scratchpad copies of the project libraries; the JSZip version string also appears in the project tree).

---

## Q1. Web Share API Level 2 (`navigator.share({files})`, `navigator.canShare`) — support, secure context / `file://`, limits, gestures, WhatsApp/Outlook targets

### Takeaway
File sharing via `navigator.share` works in Chrome/Edge on Windows, macOS, ChromeOS and Android, and in Safari on macOS/iOS, but **not in Firefox on any platform**; Chromium caps a single share at **10 files / 50 MiB**, every call needs a fresh user gesture, and `file://` is formally a secure context in both the spec and Chromium — but whether a given browser actually shows the share sheet from a local file has to be verified empirically (no authoritative source found). Targets (WhatsApp, Outlook, Mail) are whatever the OS share sheet offers; there is no way to pick one programmatically.

### Cited Findings
**Support matrix (MDN browser-compat-data, raw JSON):**
- `Navigator.share()`: Chrome desktop **89–127 partial (ChromeOS & Windows only)**, **128 full**; Chrome Android 61; Edge **81–92 Windows-only**, **93 full**; Firefox desktop 71 but only behind `dom.webshare.enabled` pref; Firefox Android 79; Safari 12.1 (iOS mirrors Safari); Android WebView: not supported (crbug 40540400). — [MDN BCD api/Navigator.json](https://raw.githubusercontent.com/mdn/browser-compat-data/main/api/Navigator.json)
- `share()` **`files` parameter**: Chrome desktop 89, Chrome Android 76, Edge 81, **Firefox: not supported (desktop or Android)**, Safari 14 (iOS mirrors), Samsung Internet 11.0, WebView: no. — [MDN BCD api/Navigator.json](https://raw.githubusercontent.com/mdn/browser-compat-data/main/api/Navigator.json)
- `Navigator.canShare()`: Chrome 128 (89–127 partial, ChromeOS/Windows), Chrome Android 75, Edge 93 (81–92 Windows), Firefox 96 behind flag, Firefox Android 96, Safari 14; `canShare({files})`: Chrome 89, Chrome Android 76, Edge 81, Safari 14, Firefox: no. — [MDN BCD api/Navigator.json](https://raw.githubusercontent.com/mdn/browser-compat-data/main/api/Navigator.json)
- caniuse "Web Share API" (current data): Chrome 155–157 yes, Edge 152–154 yes, **Firefox desktop 158–160 no**, Safari 27.x yes, iOS Safari 27.x yes, Chrome Android yes, Firefox Android yes, Samsung yes; historical note #4 "Only supported on Windows and Chrome OS" for older Chrome desktop. — [caniuse web-share.json](https://raw.githubusercontent.com/Fyrd/caniuse/main/features-json/web-share.json)
- A 2022 BCD bug report noted `navigator.share` "not working in Chrome on Linux" and desktop support then limited to Windows/ChromeOS. — [mdn/browser-compat-data issue #16823](https://github.com/mdn/browser-compat-data/issues/16823)

**Requirements and behaviour (spec/MDN/Chromium):**
- MDN: available only in secure contexts; requires the `web-share` Permissions Policy and **transient activation** ("cannot be launched at arbitrary points by a script"); `NotAllowedError` if no transient activation or a file share is blocked; `TypeError` if `files` given but unsupported; always test `navigator.canShare({files})` first (pass only `files`). Shareable types "usually" include `application/pdf`. — [MDN Navigator.share()](https://developer.mozilla.org/en-US/docs/Web/API/Navigator/share)
- The Web Share IDL marks `share()` and `canShare()` as `[SecureContext]`. — [web-share.idl (WPT mirror)](https://chromium.googlesource.com/external/w3c/web-platform-tests.git/+/HEAD/interfaces/web-share.idl)
- Chromium hard limits: `kMaxSharedFileCount = 10` and `kMaxSharedFileBytes = 50 * 1024 * 1024` (50 MiB). — [chromium share_service_impl.h](https://chromium.googlesource.com/chromium/src/+/main/chrome/browser/webshare/share_service_impl.h)
- Chromium `ShareServiceImpl::Share()` rejects when: frame inactive; `web-share` policy disabled; a non-empty share URL is not http/https; **no transient user activation (and the activation is consumed, so one gesture cannot trigger repeated share sheets)**; tab not visible/active; more than `kMaxSharedFileCount` files; filename/MIME not on the allowlist (`.pdf` / `application/pdf` **is allowed**); Safe Browsing flags a file. Platform back-ends exist for Windows (`ShareOperation`), ChromeOS (sharesheet) and Mac (`SharingServiceOperation`); other platforms hit `NOTREACHED()`. The file itself does not check secure context (that is done renderer-side). — [chromium share_service_impl.cc](https://chromium.googlesource.com/chromium/src/+/main/chrome/browser/webshare/share_service_impl.cc)
- `navigator.share()` is on MDN's list of APIs that **require transient activation**; only `window.open()` is explicitly listed there as *consuming* it. Activation-triggering events: trusted `keydown`, `mousedown`, `pointerdown` (mouse), `pointerup` (non-mouse), `touchend`. — [MDN User activation](https://developer.mozilla.org/en-US/docs/Web/Security/User_activation)
- W3C spec: share targets are chosen by the user agent/OS ("apps, websites or contacts"); the UA may simply forward the data to the OS share system. — [W3C Web Share REC 2023](https://www.w3.org/TR/2023/REC-web-share-20230530/)

**`file://` and secure contexts:**
- MDN: an origin is potentially trustworthy if its scheme is `https`, `wss` **or `file`**; table shows `file:///path/to/resource.html` = Secure. — [MDN Secure contexts](https://developer.mozilla.org/en-US/docs/Web/Security/Secure_Contexts)
- Chromium's secure-origin definition explicitly lists `(file, *, —)` alongside https, wss, localhost, 127/8, ::1 and chrome-extension. — [Chromium: Prefer Secure Origins For Powerful New Features](https://www.chromium.org/Home/chromium-security/prefer-secure-origins-for-powerful-new-features/)

**Windows share-sheet targets:**
- Windows 11 Share window: WhatsApp was added to the "Share using" section (with an install prompt if absent); Outlook appears as a target; Edge's own share list includes Outlook, Gmail, LinkedIn, Facebook, Twitter and WhatsApp. — [ElevenForum: Share file with app in Windows 11](https://www.elevenforum.com/t/share-file-with-app-in-windows-11.8014/post-168559); [ElevenForum: Share link in Edge](https://www.elevenforum.com/t/share-link-with-app-email-or-nearby-device-in-microsoft-edge.21540/post-399091); [Winaero build 26020 share UI](https://winaero.com/windows-11-build-26020-introduces-a-new-share-ui)
- Microsoft: all apps can participate in the Share dialog as targets, including PWAs installed through Edge that implement Web Share Target. — [Windows Developer Blog](https://blogs.windows.com/windowsdeveloper/?p=56290)

### Inferences
- For 1,000 payslips, Web Share is a **one-tap-per-recipient** channel at best (one activation = one share sheet) and a **≤10 files per share** channel in Chromium; it cannot "bulk send" 1,000 PDFs. It fits a "send list" UI where HR taps *Share* on one employee's row, picks WhatsApp/Outlook in the OS sheet, picks the contact, and sends.
- On Windows the OS sheet lists WhatsApp (Store app) and Outlook only if those apps are installed and registered; the web page cannot pre-select the contact/recipient — the user must do it in the target app.
- Firefox users (desktop and Android) get no file sharing at all; feature-detect with `navigator.canShare && navigator.canShare({files:[pdfFile]})` and fall back to download/ZIP.
- `file://` satisfies the spec's and Chromium's secure-context definition, so `navigator.share` should be *defined* on a local page in Chromium and Safari; whether the Windows/macOS share sheet actually appears from a `file://` page is untested here (see Gaps).

### Gaps
- No authoritative source found confirming or denying that `navigator.share({files})` actually opens the OS share sheet from a `file://` page in Chrome/Edge/Safari; must be tested (`window.isSecureContext`, `typeof navigator.share`, then a real call) on each target OS.
- Chrome on **Linux** desktop: BCD marks Chrome 128+ as full support, but current Chromium source still shows share back-ends only for Windows/ChromeOS/Mac — unresolved conflict; irrelevant for the stated Windows/macOS/mobile users.
- Exact transient-activation timeout per browser (HTML spec allows up to a few seconds; Chrome's blog mentions "about a second" in its UAv2 example) not confirmed from a primary source in this run.
- Whether WhatsApp Desktop (non-Store/Electron build) or classic Outlook register as Windows share targets: not found; only the Store WhatsApp and the in-box Outlook app are described.
- Android-specific Web Share file limits (Android `ShareServiceImpl.java`) were not read; Chromium comments say its lists must stay in sync with the desktop allowlists.

---

## Q2. File System Access API (`showDirectoryPicker`, writable streams) to write 1,000 PDFs into a folder tree — support matrix and `file://` restrictions

### Takeaway
`showDirectoryPicker({mode:'readwrite'})` is Chromium-only: Chrome/Edge 86+ on desktop, Chrome 132+ on Android (via Android intents, with caveats), **never Firefox or Safari**. One picker gesture grants read-write access to the chosen folder for the lifetime of the tab, so all 1,000 `getDirectoryHandle(..., {create:true}) → getFileHandle(..., {create:true}) → createWritable()` writes need no further prompts. `file://` satisfies Chromium's secure-origin rule, but no source was found that confirms the picker works from a local HTML file.

### Cited Findings
- BCD: `Window.showDirectoryPicker`, `showSaveFilePicker`, `showOpenFilePicker`: Chrome 86, **Chrome Android 132**, Edge mirrors Chrome (86), Opera/Samsung/WebView mirror, **Firefox: not supported, Safari: not supported** (desktop and iOS). — [MDN BCD api/Window.json](https://raw.githubusercontent.com/mdn/browser-compat-data/main/api/Window.json)
- caniuse "File System Access API": Chrome 155–157 yes, Edge 152–154 yes, Firefox no, Safari/iOS no, Samsung no, Chrome Android listed **no** (caniuse note: "Parts of the API are still missing"). — [caniuse native-filesystem-api.json](https://raw.githubusercontent.com/Fyrd/caniuse/main/features-json/native-filesystem-api.json)
- Chrome 132 release notes: the File System Access API "is now available on Android and in WebViews". — [New in Chrome 132](https://developer.chrome.com/blog/new-in-chrome-132)
- blink-dev Intent to Ship (Sept–Oct 2024): enabled for Android in M132; Android uses Intents (`GET_CONTENT` read-only or `OPEN_DOCUMENT` writable), files are content-URIs, **atomic writes and renames unsupported** (move = copy); early Canary testing found MIME filters ignored, a huge folder (DCIM) froze the browser, save picker could only overwrite existing files, and "folder save threw an InvalidStateError". — [blink-dev: Intent to Ship: File System Access on Android and WebView](https://groups.google.com/a/chromium.org/g/blink-dev/c/x3IcFv2jY6c/m/HxuEfzBHAAAJ)
- MDN `showDirectoryPicker()`: secure context only; **transient user activation required** (else `SecurityError`); options `id`, `mode: "read" | "readwrite"`, `startIn` (`desktop`, `documents`, `downloads`, … or a handle); `AbortError` if the user cancels, if the browser deems the directory too sensitive, or if permission state is not `granted` for the requested mode. — [MDN showDirectoryPicker](https://developer.mozilla.org/en-US/docs/Web/API/Window/showDirectoryPicker)
- Chrome capabilities article: pickers must run in a secure context and from a user gesture; **calling a picker after slow async work can fail because the gesture expired — obtain the handle first, then generate data**; `showDirectoryPicker()` grants read by default, pass `{mode:'readwrite'}` for write; **access lasts until all tabs for the origin are closed**; `queryPermission()`/`requestPermission()` to re-check; `createWritable()` prompts if write permission not yet granted and "changes are not written to disk until the stream is closed"; `getFileHandle(name,{create:true})` and `getDirectoryHandle(name,{create:true})` create entries; the browser may refuse core system folders (Windows system folders, macOS Library) and ask for another folder; `move()` is not yet supported for directories. — [Chrome: The File System Access API](https://developer.chrome.com/docs/capabilities/web-apis/file-system-access)
- MDN lists `showOpenFilePicker()`, `showSaveFilePicker()` and `showDirectoryPicker()` among APIs requiring transient activation. — [MDN User activation](https://developer.mozilla.org/en-US/docs/Web/Security/User_activation)
- Chromium treats `(file, *, —)` as a secure origin. — [Chromium secure origins](https://www.chromium.org/Home/chromium-security/prefer-secure-origins-for-powerful-new-features/)
- Searching WICG/file-system-access issues for `file:// origin` surfaced no issue stating the pickers are blocked on local files; closest is #474 "Define getAsFileSystemHandle() behavior in insecure contexts" (open). — [WICG/file-system-access issues](https://github.com/WICG/file-system-access/issues?q=file%3A%2F%2F+origin)

### Inferences
- Recommended flow for Chromium desktop: button click → `await showDirectoryPicker({mode:'readwrite', id:'payslips', startIn:'documents'})` **before** generating PDFs → loop over employees creating `Department/Employee/` sub-handles and writing each PDF with `createWritable()` → `close()`; no further prompts. Because writes commit on `close()`, each small PDF is written atomically on desktop.
- On Safari/Firefox (macOS users on Safari, any Firefox) this path is unavailable; fall back to a single ZIP download (Q8) which the user extracts.
- Android Chrome 132+ nominally supports it but the intent-based implementation (content URIs, no atomic writes, early InvalidStateError on folder save) makes a 1,000-file folder tree unreliable on phones; treat ZIP as the mobile fallback.
- `file://` origin: by Chromium's own definition it is a secure origin, so `window.showDirectoryPicker` should be defined; the open question is whether Chromium's permission context accepts a `file://` origin for directory grants.

### Gaps
- No primary source found confirming that `showDirectoryPicker` succeeds from a `file://` page in Chrome/Edge (needs a 1-minute empirical test).
- Safe Browsing / download-protection scanning of FSA writes was not documented in the fetched Chrome article; unknown whether 1,000 `.pdf` writes incur per-file checks.
- No published throughput numbers for writing ~1,000 small files via FSA.

---

## Q3. Client-side `.eml` generation (MIME, base64 PDF, `X-Unsent: 1`) and what happens on double-click; `.msg` feasibility

### Takeaway
An `.eml` is plain RFC 5322/MIME text and is trivial to build in the browser (CRLF lines, `multipart/mixed` with a base64 `application/pdf` part at ≤76 chars/line); **classic Outlook for Windows opens it as an editable, sendable draft only if the header `X-Unsent: 1` is present** (behaviour changed in Outlook build 16.0.14430.20256 / Version 2109); **New Outlook for Windows, Outlook for macOS, Apple Mail and Thunderbird do not honour `X-Unsent`** (Thunderbird needs "Edit As New Message", Ctrl-E). `.msg` (MS-OXMSG compound file) is not practical in-browser: the only open-source writer (tutao/oxmsg) is GPL-3.0, Node-only and early-stage.

### Cited Findings
**Format rules:**
- RFC 5322 §2.1.1: lines "MUST be no more than 998 characters" and "SHOULD be no more than 78 characters, excluding the CRLF"; lines delimited by CRLF (§2.1), header fields terminated by CRLF (§2.2); only `Date:` and `From:` are required; every message SHOULD have `Message-ID:` (§3.6.4); `Subject:` is unstructured US-ASCII (§3.6.5), `To:` an address-list (§3.6.3). — [RFC 5322](https://www.rfc-editor.org/rfc/rfc5322)
- RFC 2045 §6.8: base64 "encoded output stream must be represented in lines of no more than 76 characters each"; §5: `boundary` parameter required for any multipart subtype; §6.8: base64 output contains no hyphens so boundary collision is not a concern for base64 bodies; multipart delimiter syntax is defined in RFC 2046. — [RFC 2045](https://www.rfc-editor.org/rfc/rfc2045)
- RFC 6068 (mailto) notes header values with non-ASCII must be "suitably encoded (for example, as MIME encoded words)" when composing. — [RFC 6068](https://www.rfc-editor.org/rfc/rfc6068)

**Outlook behaviour:**
- Microsoft Q&A (Oct 2021): a hosted web app generating "plaintext MIME .eml files with attachments" that Outlook 2010–2019 opened as editable drafts; since **Outlook build 16.0.14430.20256 (Version 2109)** they open as already-sent messages; adding **`X-Unsent: 1`** makes that build open the EML as a draft again (answer marked helpful by 20+); a later commenter reports **Outlook for macOS does not honour the header**; Microsoft's only suggestion was reverting Office. — [Microsoft Q&A: EML files no longer opening as a draft](https://learn.microsoft.com/en-us/answers/questions/4562029/eml-files-no-longer-opening-as-a-draft-e-mail-to-s)
- Microsoft Q&A (May 2024): with `X-Unsent: 1`, classic Outlook lets the user edit and send; **New Outlook for Windows opens the message but the user cannot edit or send it**; Microsoft community moderator: "not feasible in New Outlook for Windows at this time", file feedback. — [Microsoft Q&A: X-Unsent not recognized in New Outlook](https://learn.microsoft.com/en-us/answers/questions/4635248/header-x-unsent-for-eml-files-is-not-recognized-in)
- Sept 2025 thread: New Outlook `.eml` with `X-Unsent: 1` drafts open but sending fails repeatedly; moderator said it worked in their testing. — [Microsoft Q&A 5562698](https://learn.microsoft.com/en-au/answers/questions/5562698/new-outlook-eml-x-unsent-1-sending-not-working-mul); Oct 2025: new .eml files not saving to Drafts in new Outlook web, reported fixed in a later build. — [Microsoft Q&A 5594291](https://learn.microsoft.com/en-us/answers/questions/5594291/new-eml-files-will-not-save-to-draft-folder-(fixed)
- `X-Unsent` is a custom header, not a standard; Apple Mail ignores it and uses its own `Uniform-Type-Identifier: com.apple.mail-draft`. — [Rebex forum](https://forum.rebex.net/9290/save-email-as-unsent-using-eml-in-apple-mail?show=9291); eM Client has an open feature request to support `X-Unsent`. — [eM Client forum](https://forum.emclient.com/t/add-support-for-x-unsent-header/37519)

**Thunderbird:**
- Opening an `.eml` shows it as a received message with Send disabled; workaround is **Message → Edit As New Message (Ctrl-E)** to get a compose window. — [Mozilla Support question 1022415](https://support.mozilla.org/de/questions/1022415)
- Aspose support: "tried different scenarios for saving the EML so that Thunderbird can open it as draft, but it seems this is not possible"; `X-Unsent` "only understandable by MS Outlook". — [Aspose forum](https://forum.aspose.com/t/open-eml-file-in-edit-mode-for-ibm-notes-and-thunderbird/160296)

**`.msg`:**
- tutao/oxmsg: "write Microsoft .msg Outlook Item files" (CFB compound file via `cfb`), **GPL-3.0**, based on MsgKit (C#); `Email` → `.msg()` returns `Uint8Array`; **Node-only** — browser use would require replacing Buffer-dependent parsers (`iconv-lite`), polyfilling BigInt and removing BigInt literals; TODOs unfinished, no releases listed. — [github.com/tutao/oxmsg](https://github.com/tutao/oxmsg)
- `@independentsoft/msg` is a commercial JS/TS library (30-day evaluation) that can create/modify Outlook items. — [npm @independentsoft/msg](https://npmjs.com/package/@independentsoft/msg); `msg2eml.js` runs in the browser but only converts .msg→.eml; `msg-parser` is read-only. — [msg2eml.js](https://github.com/master131/msg2eml.js); [msg-parser](https://npmjs.com/package/msg-parser)

**Libraries / prior art for client-side MIME building:**
- MIMEText (muratgozel): "RFC-2822, RFC-2045 and RFC-2049 compliant raw email message generator", runs in Node, browsers and Apps Script, supports regular and inline attachments. — [MIMEText](https://github.com/muratgozel/MIMEText)
- `eml-parse-js` (@vortiq fork) also exposes a `buildEml` function and added base64 attachment support. — [@vortiq/eml-parse-js README](https://unpkg.com/@vortiq/eml-parse-js@0.0.5/README.md)

### Inferences
- A minimal Outlook-classic-compatible draft (from the cited RFC rules + Q&A reports):
  `X-Unsent: 1` / `To:` / `Subject:` (RFC 2047 encoded-word for Arabic) / `Date:` / `From:` (omit or set the employer address) / `MIME-Version: 1.0` / `Content-Type: multipart/mixed; boundary="..."`, then a `text/plain; charset=utf-8` (or `text/html`) body part and an `application/pdf; name="..."` part with `Content-Transfer-Encoding: base64`, `Content-Disposition: attachment; filename="..."`, base64 wrapped at 76 chars, **all lines CRLF**, closing `--boundary--`. Build as a `Blob` and download; because base64 has no hyphens, any random `----=_Part_<random>` boundary is safe.
- Double-click outcome by client: classic Outlook (Windows, ≥2109 with `X-Unsent: 1`) → editable draft with attachment; New Outlook → read-only; Outlook macOS / Apple Mail → read-only; Thunderbird → read view, Ctrl-E to compose. So `.eml` is a reliable "one-click send" path **only for classic Outlook on Windows**; everywhere else it is a "forward/edit-as-new" path.
- `.msg` generation should be declared out of scope (GPL-3.0 + Node-only + beta).

### Gaps
- No open-source project found that specifically does "client-side mail-merge to `.eml` with attachments and `X-Unsent`"; the closest building block is MIMEText.
- Whether New Outlook will ever honour `X-Unsent` is unknown (Microsoft points to feedback portal only).
- Non-ASCII attachment file names in `.eml` (RFC 2231 `filename*=UTF-8''…`) and Outlook's handling of them were not verified in this run.
- Outlook's handling of a `From:`-less or mismatched-`From:` draft was not verified.

---

## Q4. PDF encryption in the browser — jsPDF built-in vs alternatives (AES-256, size, licence) and viewer compatibility

### Takeaway
jsPDF (current release 4.2.1; encryption added in v2.2.0 via PR #2959, and present in the project's vendored 2.5.x builds) encrypts with the **Standard handler V=1/R=2, 40-bit RC4 (MD5-derived key)** and does not write `/Length`; permissions are `print|modify|copy|annot-forms`. PDFium (Chrome/Edge) and pdf.js (Firefox) both still open R2/40-bit files. For AES-256 (R6) in a vendored, offline bundle, the smallest permissive options are **`@pdfsmaller/pdf-encrypt` (MIT, ~17 KB gz, needs pdf-lib) or the `@cantoo/pdf-lib` fork (MIT, AES-256 R6 default)**; qpdf-based WASM (`pdfstudio`, Apache-2.0) costs ~2 MB; mupdf.js is AGPL.

### Cited Findings
**jsPDF built-in encryption (source, master = 4.2.1):**
- Constructor options: `options.encryption.userPassword`, `ownerPassword` ("Both … should be set for proper authentication"), `userPermissions` array of `"print"`, `"modify"`, `"copy"`, `"annot-forms"`. — [jsPDF src/jspdf.js](https://raw.githubusercontent.com/parallax/jsPDF/master/src/jspdf.js) (JSDoc lines ~190–193); current `package.json` version `4.2.1`. — [jsPDF package.json](https://raw.githubusercontent.com/parallax/jsPDF/master/package.json)
- `pdfsecurity.js`: `this.v = 1; // algorithm 1, future work can add in more recent encryption schemes`, `this.r = 2; // revision 2`; file key = `md5Bin(paddedUserPassword + O + P(LE) + fileId).substr(0, 5)` (**5 bytes = 40-bit**); `this.U = rc4(this.encryptionKey, this.padding)`; O computed with `md5Bin(paddedOwnerPassword).substr(0, 5)` then `rc4(key, paddedUserPassword)`; permission base `192` plus print=4, modify=8, copy=16, annot-forms=32, `this.P = -((protection ^ 255) + 1)`; standard 32-byte Adobe padding string. Bug: validation uses `typeof permissionOptions.perm` so invalid permission names never throw. — [jsPDF src/libs/pdfsecurity.js](https://raw.githubusercontent.com/parallax/jsPDF/master/src/libs/pdfsecurity.js); docs mirror — [libs_pdfsecurity.js.html](https://raw.githack.com/parallax/jsPDF/master/docs/libs_pdfsecurity.js.html)
- `putEncryptionDict` writes exactly `/Filter /Standard`, `/V`, `/R`, `/U <…>`, `/O <…>`, `/P` — **no `/Length`**; every stream passes through `encryption.encryptor(objectId, 0)` in `putStream` ("ObjectId must be passed to putStream for file encryption"), and Info-dictionary strings are encrypted too. — [jsPDF src/jspdf.js](https://raw.githubusercontent.com/parallax/jsPDF/master/src/jspdf.js)
- Encryption shipped in **jsPDF v2.2.0** ("#2959 Add support for PDF encryption"). — [jsPDF release v2.2.0](https://github.com/parallax/jsPDF/releases/tag/v2.2.0); PR #2959 "Add functions to encrypt PDFs for pdf version 1.3" merged Oct 28 2020. — [jsPDF issues search](https://github.com/parallax/jsPDF/issues?q=encryption+password)
- Related issues: #3593 "PDF is not asking for password in Safari browser" (closed Sep 2023); #3611 "2.5.1 Using a custom Chinese font, when using userPassword, the bookmarks are garbled" (closed Oct 2023). — [jsPDF issues search](https://github.com/parallax/jsPDF/issues?q=encryption+password)

**Viewer support for the algorithms:**
- PDFium (Chrome/Edge built-in viewer): branches `if (revision_ == 2)` (single RC4 pass), R≥3 (MD5 loop + 20 RC4 rounds), `revision_ >= 5` → AES-256 (`AES256_CheckPassword`, R6 uses `Revision6_Hash`); RC4 key length accepted 5–16 bytes; **missing `/Length` defaults to 40-bit for V<4**, 128 for V4, 256 for V≥5; AESV2/AESV3 both map to AES. — [pdfium cpdf_security_handler.cpp](https://pdfium.googlesource.com/pdfium/+/main/core/fpdfapi/parser/cpdf_security_handler.cpp)
- pdf.js (Firefox built-in viewer): accepts V 1, 2, 4, 5 (else "unsupported encryption algorithm"); R2 simple RC4 path, R≥3 50 MD5 rounds; V=1/2 default 40-bit ("For 1 and 2 it's fixed to 40-bit"); AESV2 = AES-128; `algorithm === 5 || cfm.name === "AESV3"` → AES-256 with `revision === 6 ? new PDF20() : new PDF17()`. — [pdf.js src/core/crypto.js](https://raw.githubusercontent.com/mozilla/pdf.js/master/src/core/crypto.js)
- MeSciUS docs (for their Rev4 handler): "tests show that current/recent versions of Adobe Acrobat Reader DC … cannot handle files encrypted with RC4 if the key length is not equal to 128" — context is the **Revision 4** handler; they recommend Rev3 for other RC4 lengths. — [MeSciUS StandardSecurityHandlerRev4](https://developer.mescius.com/document-solutions/javascript-pdf-api/api/classes/StandardSecurityHandlerRev4)

**In-browser alternatives:**
- **`@cantoo/pdf-lib`** (fork of pdf-lib): `doc.encrypt({userPassword, ownerPassword, permissions:{printing:'highResolution'}, algorithm?})` before `save()`; **AES-256 revision 6 (ISO 32000-2) by default**, `algorithm:'AES-128'` available, RC4 refused unless `allowWeakCryptography:true`; needs Web Crypto `crypto.getRandomValues`; in-tree AES/RC4/MD5/SHA-256 (no crypto-js); also decrypts (`password:` on load); **MIT**. Version not shown on the README (npm page returned 403 to this run). — [@cantoo/pdf-lib README (jsDelivr)](https://cdn.jsdelivr.net/npm/@cantoo/pdf-lib/README.md)
- **`@pdfsmaller/pdf-encrypt`**: `encryptPDF(pdfBytes, userPassword, {ownerPassword, algorithm:'AES-256'|'RC4', allowPrinting, allowHighQualityPrint, allowModifying, allowCopying, allowAnnotating, allowFillingForms, allowExtraction, allowAssembly})`; **AES-256 (V=5,R=6) default; RC4-128 (V=2,R=3) legacy**; **~17 KB gzipped**, zero deps besides **pdf-lib peer dependency**; UMD build exposes global `PDFEncrypt` and reads `PDFLib`; uses Web Crypto, so **AES-256 requires a secure context**; **MIT**; 1.0.x corrupted PDFs containing literal strings — use ≥1.1.0. — [@pdfsmaller/pdf-encrypt README](https://cdn.jsdelivr.net/npm/@pdfsmaller/pdf-encrypt@1.0.2/README.md)
- **`pdfstudio`** (qpdf compiled to WASM): **Apache-2.0** (qpdf licence; wasm also bundles zlib and libjpeg), `qpdf.wasm` **~2.1 MB**, qpdf **12.3.2** with native crypto; `createPdfToolkit({wasmUrl:'/assets/qpdf.wasm'})` for self-hosting; `lock(doc,{userPassword, ownerPassword, keyLength:256|128|40, permissions:{print:'full'|'low'|'none', modify:'all'|'annotate'|'form'|'assembly'|'none', extract:boolean}})`; runs on main thread or Web Worker, Node 18+, Cloudflare Workers. — [pdfstudio README](https://cdn.jsdelivr.net/npm/pdfstudio@0.4.0/README.md)
- **`qpdf-wasm`** npm: described only as "qpdf CLI compiled to WASM"; licence/size/usage not available in fetched README (GitHub repo URL 404). — [qpdf-wasm README](https://cdn.jsdelivr.net/npm/qpdf-wasm/README.md)
- **mupdf.js**: npm `mupdf` **1.28.1**, licence **AGPL-3.0-or-later**. — [mupdf package.json](https://cdn.jsdelivr.net/npm/mupdf/package.json); PDF write options: `encrypt=none|keep|rc4-40|rc4-128|aes-128|aes-256`, `user-password=`, `owner-password=`, `permissions=NUMBER`, plus `garbage`, `compress`. — [MuPDF PDF write options](https://mupdf.readthedocs.io/en/latest/reference/common/pdf-write-options.html)
- Upstream pdf-lib: the Cantoo README lists `encrypt()` among features "upstream does not (or no longer) provide". — [@cantoo/pdf-lib README](https://cdn.jsdelivr.net/npm/@cantoo/pdf-lib/README.md)
- `pdf-oxide-wasm` advertises "Encryption | Yes | AES-256" for browsers (not evaluated further). — [PDF Oxide WASM docs](https://pdf.oxide.fyi/docs/getting-started/javascript-wasm)
- No npm package for a pdfcpu WASM build surfaced in search. — [search context: pdfstudio README](https://cdn.jsdelivr.net/npm/pdfstudio@0.4.0/README.md)
- Adobe Acrobat integrated into WhatsApp **Web and Windows** (view/markup PDFs in chat; password-protected files keep security) — not the Android app. — [ground.news summary](https://ground.news/article/how-to-login-to-two-whatsapp-accounts-on-one-phone-without-additional-apps-use-whatsapp-web)
- Opening a password-protected PDF on Android requires a reader that supports it (e.g. Acrobat Reader prompts for the password). — [Adobe: open password protected PDFs on Android](https://www.adobe.com/acrobat/hub/how-to/open-password-protected-pdfs-on-android.html)

### Inferences
- jsPDF's built-in RC4-40/R2 is the **zero-extra-bytes** option and is readable by Chrome/Edge (PDFium) and Firefox (pdf.js) per their sources; R2 predates the Rev4 RC4-length quirk MeSciUS describes, so it is probably fine in Acrobat Reader, but 40-bit RC4 is cryptographically weak — adequate only as a "stop casual opening" measure for payslips, not real confidentiality.
- Cheapest AES-256 upgrade path that stays offline/vendored: keep jsPDF for layout, then post-process each `doc.output('arraybuffer')` with `@pdfsmaller/pdf-encrypt` (+ pdf-lib, both MIT) or load it in `@cantoo/pdf-lib` and `encrypt()`; this adds a parse+re-serialize per PDF (cost to measure for 1,000 files) and requires Web Crypto, which `file://` satisfies as a secure context (Q1).
- qpdf-WASM (`pdfstudio`) gives the most battle-tested encryption engine but a ~2 MB wasm that must be vendored next to the HTML; mupdf.js is excluded for a proprietary/internal tool unless AGPL is acceptable.
- The two jsPDF issues suggest testing: (a) password prompt behaviour in Safari, (b) custom (Arabic) font + `userPassword` for garbled outline/metadata strings.

### Gaps
- No source confirming that jsPDF R2/RC4-40 PDFs open correctly in **Adobe Reader DC**, **iOS Files/Quick Look**, or the **Android WhatsApp in-app viewer**; nor whether those viewers open **AES-256 R6**. Apple documentation found only covers Preview *writing* passwords. Empirical test required.
- `@cantoo/pdf-lib` version/unpacked size not captured (npm page blocked); `qpdf-wasm` licence/size unavailable.
- No measured timing for encrypting 1,000 PDFs with any of these libraries in a browser.

---

## Q5. Offline QR-code libraries and embedding QR into jsPDF

### Takeaway
`qrcode-generator` (kazuhikoarase, MIT, single script, no deps) and `node-qrcode` (soldair, MIT, `build/qrcode.js` global) both run fully offline and emit PNG data URLs/canvas/SVG; `qrcodejs` (davidshimjs, MIT) is old and DOM-oriented. Any PNG data URL or canvas drops straight into `doc.addImage()`.

### Cited Findings
- **node-qrcode**: MIT; browser bundle `/build/qrcode.js` exposes global `QRCode`; `toCanvas(canvas, text, opts)`, `toDataURL(text, {type:'image/png'|'image/jpeg'|'image/webp'})`, `toString(text,{type:'svg'})`; error-correction L/M/Q/H (~7/15/25/30 %); versions 1–40 auto-selected; byte mode carries arbitrary bytes and multibyte UTF-8 (ECI not fully implemented); supports IE10+/Safari 5.1+/evergreen. — [github.com/soldair/node-qrcode](https://github.com/soldair/node-qrcode)
- **qrcode-generator**: MIT; usage `qrcode(typeNumber /*1–40, 0 = auto*/, 'L'|'M'|'Q'|'H')`, `addData(data, mode /*Numeric|Alphanumeric|Byte(default)|Kanji*/)`, `make()`, then `createDataURL(cellSize, margin)`, `createImgTag`, `createSvgTag({cellSize, margin, scalable})`, `createASCII`; include via `<script src="qrcode.js">`; `qrcode.stringToBytes(s)` can be overwritten "to encode using a multibyte charset". — [qrcode-generator js/README](https://raw.githubusercontent.com/kazuhikoarase/qrcode-generator/master/js/README.md); repo licence MIT, implementations based on JIS X 0510:1999. — [github.com/kazuhikoarase/qrcode-generator](https://github.com/kazuhikoarase/qrcode-generator)
- **qrcodejs** (davidshimjs): MIT; `new QRCode(el,{text,width,height,colorDark,correctLevel})`, canvas with table fallback, targets IE6–10-era browsers; 177 open issues / 61 open PRs (maintenance unclear). — [github.com/davidshimjs/qrcodejs](https://github.com/davidshimjs/qrcodejs)
- jsPDF `addImage(imageData, format, x, y, width, height, alias, compression, rotation)`: `imageData` may be a base64 data-URL string, `HTMLImageElement`, `HTMLCanvasElement` or `Uint8Array`; `alias` = "alias of the image (if used multiple times)"; compression `NONE|FAST|MEDIUM|SLOW` (JPEG). — [jsPDF addImage docs](https://artskydj.github.io/jsPDF/docs/module-addImage.html)

### Inferences
- Preferred: `qrcode-generator` (smallest, no build step) → `qr.createDataURL(4, 0)` → `doc.addImage(dataUrl, 'PNG', x, y, w, h)`; for Arabic/UTF-8 payloads override `qrcode.stringToBytes` with `new TextEncoder().encode(s)` or use node-qrcode byte mode. QR payload candidates that need no server: a `wa.me/<phone>?text=…` link, a verification string (employee ID + period + hash), or a `mailto:`.
- Because `addImage` supports an `alias`, a per-document QR (unique per employee) cannot be shared, but a company logo can, which keeps per-PDF size down.

### Gaps
- Exact minified sizes of the three libraries were not confirmed from primary pages.
- PNG decoding cost of 1,000 QR `addImage` calls in jsPDF was not benchmarked (jsPDF's PNG path re-encodes pixel data).

---

## Q6. Opening `wa.me` / `whatsapp://` / `mailto:` / `ms-outlook:` links in bulk — popup blockers, gesture-per-link, UX

### Takeaway
Browsers allow **one popup/navigation per user activation** (`window.open` consumes it), so a loop that opens 1,000 `wa.me`/`mailto:` links is blocked after the first; URL schemes can only pre-fill text (never attach a file). The workable UX is a "send list" where each row's button opens exactly one link (ideally a plain `<a target="_blank">`), and the attachment reaches the recipient via the OS share sheet (Q1), a saved file (Q2/Q8) or an `.eml` draft (Q3).

### Cited Findings
- Chrome: "only one popup is allowed per user activation because `window.open()` consumes user activation"; transient activation expires after a short interval or when consumed; a popup from a delayed callback still succeeds if within the window. — [Chrome: Making user activation consistent across APIs](https://developer.chrome.com/blog/user-activation)
- MDN: `Window.open()` requires transient activation and is the API explicitly described as consuming it; activation comes only from trusted `mousedown`/`pointerdown`/`pointerup`/`touchend`/`keydown`; `navigator.userActivation.isActive` exposes state. — [MDN User activation](https://developer.mozilla.org/en-US/docs/Web/Security/User_activation)
- WhatsApp click-to-chat: `https://wa.me/<number>?text=<URL-encoded>` pre-fills an editable message; the person must press send; encode with `encodeURIComponent`; no parameter for files/documents exists in the documented link forms. — [Unipile: wa.me & api.whatsapp.com explained](https://www.unipile.com/de/?p=275019); [Dotdigital click-to-chat](https://support.dotdigital.com/en/articles/11331301-click-to-chat-for-whatsapp)
- `mailto:` (RFC 6068) defines no attachment mechanism; `body` is for short text "not for general MIME bodies"; `From`/`Date` in the URI MUST be ignored; only `Subject`, `Keywords`, `Body` are "believed to be both safe and useful"; non-ASCII must be UTF-8 then percent-encoded (e.g. `subject=caf%C3%A9`), line breaks as `%0D%0A`; mail clients SHOULD NOT send without showing the message. — [RFC 6068](https://www.rfc-editor.org/rfc/rfc6068)
- Outlook desktop handles only ~2000 characters in a `mailto:` link; other reports cite 1840–2048 depending on client/version. — [Atlassian support note](https://su-support.atlassian.net/wiki/x/hqZmB); [evolt thelist 2001](https://lists.evolt.org/archive/Week-of-Mon-20010813/028060.html)
- Edge `RegisteredProtocolHandlers` policy can map `mailto` to an Outlook-on-the-web compose URL with `%s`. — [Microsoft Edge policy: RegisteredProtocolHandlers](https://learn.microsoft.com/en-my/deployEdge/microsoft-edge-policies/registeredprotocolhandlers)
- Chromium `ShareServiceImpl` consumes transient activation on each `share()`. — [chromium share_service_impl.cc](https://chromium.googlesource.com/chromium/src/+/main/chrome/browser/webshare/share_service_impl.cc)

### Inferences
- "Send list" pattern: a table of employees with per-row buttons — *Share PDF* (`navigator.share`, Chromium/Safari), *WhatsApp text* (`<a href="https://wa.me/966…?text=…" target="_blank" rel="noopener">`), *Email* (`<a href="mailto:…?subject=…&body=…">`), *Download .eml* (classic-Outlook users), *Save PDF* (FSA or `<a download>`). Each click is one activation → one action, so it never trips popup blockers; use real anchors instead of `window.open` to avoid activation consumption and to keep middle-click/long-press behaviour. Mark rows done (persist in `localStorage`) so HR can resume after interruptions.
- Keep `mailto:` bodies under ~1,800 characters (Outlook) and put the Arabic text UTF-8 percent-encoded; attachments must travel separately (shared folder, ZIP, `.eml`).
- Automating 1,000 sends is impossible client-side; the achievable goal is to reduce each send to 2–3 clicks.

### Gaps
- No documentation found for an `ms-outlook:` URL scheme on Windows / New Outlook (results only cover `mailto` → OWA mapping); treat `ms-outlook:` as unsupported unless tested.
- `whatsapp://send?phone=…&text=…` desktop-scheme behaviour on Windows/macOS not verified against an official source.
- Per-browser transient-activation durations not confirmed from primary sources.

---

## Q7. Bulk printing (one long PDF vs browser viewer) and iframe PDF preview behaviour

### Takeaway
Printing a PDF from a hidden/visible iframe via `contentWindow.print()` works in Chrome/Edge (PDFium) but is a long-standing failure in Firefox (pdf.js, bug 856116) and unreliable in Safari; iOS Safari renders an embedded PDF as a **single first page**, and Android Chrome tends to download rather than render an iframed PDF. The pragmatic bulk-print path is one merged multi-page PDF opened in a new tab and printed from the browser's own viewer.

### Cited Findings
- Firefox: "Automatic print PDF files loaded into an iframe via javascript do not pop up the print dialog using pdf.js" (Bugzilla 856116, 2013, closed as duplicate). — [bugzilla.mozilla.org 856116](https://bugzilla.mozilla.org/show_bug.cgi?id=856116); Mozilla support thread: JavaScript printing the iframe's PDF "no longer works". — [support.mozilla.org 798300](https://support.mozilla.org/en-US/questions/798300); 2024 Discourse thread: MDN pattern (iframe `src` → `contentWindow.print()` on load) "In Chrome everything works well, but in Firefox an error occurs". — [discourse.mozilla.org](https://discourse.mozilla.org/t/print-an-external-page-without-opening-it-not-working-in-firefox/124853)
- pdf.js issue: `window.document.execCommand('print')` workaround for other browsers "is not supported in Firefox". — [lightrun / pdf.js issue mirror](https://lightrun.com/answers/mozilla-pdf-js-unable-to-print-pdf-when-loaded-in-iframe)
- iOS Safari shows only the first page of an embedded/iframed PDF; desktop shows all pages. — [WebKit bug 253671](https://onbugs.webkit.org/show_bug.cgi?id=253671); Apple Developer Forums: "the only workaround currently known is to provide a download-link that opens the PDF full-page instead of embedding it". — [developer.apple.com forums 649982](https://developer.apple.com/forums/thread/649982); Pega KB: known Apple iOS/Safari issue, nothing can be done on the page side. — [Pega community](https://community.pega.com/node/239651)
- Android: an iframed PDF "prompts to download but is never displayed … In chrome it does the same" (single forum report). — [TMS Software support](https://support.tmssoftware.com/t/using-twebbrowsercontrol-to-view-multi-page-pdf/13619)

### Inferences
- Desktop Chrome/Edge: build one combined PDF (jsPDF `addPage` per payslip) → `URL.createObjectURL(blob)` → open in a new tab (user click) → user presses Ctrl+P; or iframe + `contentWindow.print()` for a one-click flow (Chromium only; feature-detect by UA and fall back to new tab for Firefox/Safari).
- A single 1,000-page combined PDF with one embedded font subset is much smaller than 1,000 separate files (font embedded once) and is the best "print everything" artifact; for per-department printing, generate one combined PDF per department.
- Mobile preview: do not iframe PDFs on iOS/Android; render previews with a canvas (pdf.js) or offer "Open PDF" in a new tab.

### Gaps
- jsPDF's `autoPrint()` plugin (injects an OpenAction print JS into the PDF) was not verified in this run; viewer honouring of OpenAction/JavaScript varies and was not sourced.
- Safari desktop iframe PDF printing behaviour not sourced beyond the general Firefox/Safari failure reports.

---

## Q8. Performance/memory for 1,000 PDFs (~25 KB Arabic font subset each) in JSZip; streaming alternatives (client-zip, fflate); multi-download limits

### Takeaway
A 30–60 MB ZIP of 1,000 small PDFs is well within JSZip 3.10.1's in-memory model (JSZip holds the whole result in RAM, a maintainer estimates ~3× the archive size), so the realistic concern is time, not crashes; fflate (MIT, ~8–33 KB) is ~25–50 % faster and streams, client-zip (MIT, 6.4 KB) is store-only and streams via `ReadableStream`. Use STORE (level 0) for PDFs, whose streams are already Flate-compressed. Triggering 1,000 individual downloads is blocked by Chromium's download limiter after the first file (prompt "download multiple files"), so **ZIP (+ FSA folder write on Chromium) is the pragmatic path**.

### Cited Findings
**JSZip:**
- JSZip docs: `async()`/`generateAsync()` keep the full result in memory; a 10 MB ASCII string costs ~20 MB (UTF-16); prefer `type:"uint8array"|"blob"|"arraybuffer"`; use `StreamHelper` (`pause()`/`resume()`) for results too big; reading DEFLATE→writing DEFLATE or STORE→STORE avoids recompression; no encryption/multi-volume. — [JSZip limitations](https://stuk.github.io/jszip/documentation/limitations.html)
- Issue #617: `generateAsync` in JSZip 3.2.x ~50× slower than 3.1.5 (0.8 s vs 49 s for 137.65 MB), reproduced by maintainers; later versions reportedly mixed. — [Stuk/jszip #617](https://redirect.github.com/Stuk/jszip/issues/617)
- Issue #343: maintainer estimates generating a 1.3 GB zip needs ~3× that size in memory; `generateInternalStream` reduces pressure. — [Stuk/jszip #343](https://github.com/Stuk/jszip/issues/343)

**client-zip vs fflate vs JSZip (client-zip README benchmarks, mean s over 10 runs, Blob output, no compression):**
- Sizes: client-zip **6.4 kB min / 2.6 kB gz**; fflate **29.8 kB min / 11 kB gz**; JSZip **94.9 kB min / 27.6 kB gz**. Exp 2 (6,214 small files) Chrome: client-zip 3.567 s, fflate 3.506 s, JSZip 4.602 s; Safari: 2.173 / 2.157 / 2.631 s. Exp 3 (12,044 files) Chrome: 4.604 / 3.972 / 6.297 s. client-zip is MIT, **store-only (no compression)**, Zip64, returns a `Response`/`ReadableStream` immediately, accepts (async) generators so files can be produced lazily, needs BigInt (ES2020); Safari <15.4 lacks Service-Worker downloads. — [github.com/Touffy/client-zip](https://github.com/Touffy/client-zip)
- fflate: MIT (SheetJS exempt to 0BSD); core ~8 kB, max ~33 kB (12.5 kB gz); up to 50 % faster compression than pako; streaming `Zip` + `ZipDeflate`/`ZipPassThrough` (level 0, tree-shakable) + `AsyncZipDeflate` (Web Worker, ~50 ms startup; sync faster under ~50 kB payloads); "supports files up to 4GB"; recommends **level 0 for already-compressed formats (PNG, JPEG, PDF)**; keep chunks <5,000 files when unzipping. — [github.com/101arrowz/fflate](https://github.com/101arrowz/fflate)

**jsPDF per-file size:**
- jsPDF subsets embedded TTF fonts: `ttffont.js` creates a `Subset` per font, generates a cmap of used glyphs and re-encodes `glyf`/`loca`/`head` tables, prefixing the PostScript name with a `AAAAAA+` subset tag. — [jsPDF src/libs/ttffont.js](https://raw.githubusercontent.com/parallax/jsPDF/master/src/libs/ttffont.js)

**Multi-download behaviour:**
- Chromium `DownloadRequestLimiter`: each navigation starts in `ALLOW_ONE_DOWNLOAD`; after the first download it moves to `PROMPT_BEFORE_DOWNLOAD`; the user's choice persists until navigating to a different host; a mouse click / Enter / space / navigation resets to `ALLOW_ONE_DOWNLOAD`; `kMaxDownloadsAtOnce = 50` ("Max number of downloads before a 'Prompt Before Download' Dialog is shown"); designed "to keep pages from downloading multiple files without user interaction". — [chromium download_request_limiter.h](https://chromium.googlesource.com/chromium/src/+/main/chrome/browser/download/download_request_limiter.h)
- Chrome setting: Privacy and security → Site settings → Additional permissions → **Automatic downloads** (ask / block, per-site allow lists). — [How-To Geek](https://www.howtogeek.com/428416/how-to-enabledisable-multiple-file-downloads-in-chrome/)
- Developer test (Chrome): a loop of `<a download>.click()` left "only the last file downloaded"; 1-second `setTimeout` staggering was flaky for large files; a hidden `<iframe sandbox="allow-downloads">` per file was reliable in Chrome; Safari lacked `allow-downloads` (a commenter says Safari 17+ supports it). — [zenn.dev multiple-download](https://zenn.dev/foxtail88/articles/multiple-download)
- Safari: per-site download permission under Safari → Settings → Websites → Downloads; "Safari does not consistently support multi-file downloads"; Firefox: blocked pop-ups can leave only one file downloading — add to pop-up exceptions. — [Ottomatic community](https://community.ottomatic.cloud/t/fmsvg-page-wording/714); [Storyblocks help](https://help.storyblocks.com/en/articles/6001903-where-are-all-the-downloads-i-initiated-from-my-folder)

### Inferences
- Size budget: with jsPDF subsetting, each payslip ≈ 25 KB font subset + a few KB of content ≈ 30–60 KB → 30–60 MB for 1,000 files; JSZip 3.10.1 in STORE mode (`compression:"STORE"`) keeps ~3× (≈100–200 MB) transient RAM — fine on desktop, borderline on low-end phones. Generation time will be dominated by jsPDF (font subsetting + optional encryption) rather than zipping, given 6,000-file ZIPs build in ~3.5–4.6 s in the cited benchmarks.
- If JSZip proves slow, switch to **fflate `Zip` + `ZipPassThrough`** (streams chunks as each PDF is produced; chunks can be piped to an FSA `createWritable()` or accumulated into a Blob) or **client-zip** (`makeZip(asyncGenerator)` → `Response` → `.blob()` or pipe to a writable). Both are MIT and vendorable as a single file.
- Do not attempt 1,000 `<a download>` clicks: Chromium prompts after the first and silently cancels fast successive clicks; Safari/Firefox are even less predictable. Offer (a) ZIP download for everyone, (b) FSA folder tree on Chromium desktop, (c) per-row single downloads from the send list (each click is its own activation, so no prompt).
- Yield to the event loop every N PDFs (`await new Promise(r=>setTimeout(r))` or a Web Worker for zipping via fflate's async API) to keep the UI responsive and the progress bar truthful.

### Gaps
- No published benchmark of JSZip 3.10.1 specifically with ~1,000 × 50 KB inputs; the 50× regression (#617) was reported for 3.2.x and its status on 3.10.x is unclear — measure in-app.
- fflate's README does not benchmark against JSZip directly; only client-zip's table provides the three-way comparison, and it uses no compression.
- Firefox's exact multi-download policy (no limiter equivalent documented) and Edge's (assumed identical to Chrome as Chromium) were not confirmed from primary sources.
- Chrome's in-memory Blob limits/paging thresholds were not sourced (not expected to matter at 60 MB).
