# "Pull" and Physical Payslip Distribution for ~1,000 Low-Literacy, Multilingual Blue-Collar Workers (Saudi Arabia / NEOM contractors), 2025–2026

Scope: no email servers, no messaging APIs, no paid portals; the payslip generator is a static offline web page. Options researched: password-protected PDFs on a shared location, per-employee cloud folders (OneDrive/SharePoint/Google Drive) with manual sharing, USB/kiosk PC, printed payslip + QR code, QR encoding the payslip data or a verification/WhatsApp link, SMS-with-code, and notice-board/supervisor hand-out schemes. Delivery proof and privacy (PDF passwords) covered throughout.

---

## Q1. QR code capacity — can a whole payslip (data, not the PDF) fit? wa.me / tel / sms link QRs? Offline browser QR libraries?

### Takeaway
A QR code can hold a short, text-only payslip summary (name, month, gross/net/deductions) but **cannot** hold a PDF file. Capacity tops out at ~2,953 bytes (binary, lowest error-correction), and real-world scannability for low-literacy workers argues for keeping QR content to a few hundred bytes or a short link. wa.me / tel / sms links encode in a QR trivially. Fully-offline browser QR generation is a solved problem (qrcode.js, EasyQRCodeJS) with no CDN or server needed.

### Cited Findings
- QR max capacity (version 40, error-correction level L): **2,953 bytes binary (8-bit); 7,089 numeric digits; 4,296 alphanumeric characters; 1,817 Kanji**. Version 40 is 177×177 modules. — [Keyence](https://www.keyence.com/ss/products/auto_id/codereader/basic_2d/qr.jsp); [Simple Wikipedia](https://simple.wikipedia.com/wiki/QR_code)
- Capacity falls sharply as error-correction rises; level H restores ~30% of codewords, so usable byte capacity is far below the L figure. Higher error correction = less storage. — [Keyence](https://www.keyence.com/ss/products/auto_id/codereader/basic_2d/qr.jsp)
- A payslip's textual data (employee name, ID, month, basic, allowances, overtime, deductions, net) is typically a few hundred bytes, well within QR capacity — but a 177×177 dense code is hard to scan on cheap phones in bright outdoor/site light, so practical designs keep content small and error correction moderate. (Inference from capacity/module data above.)
- wa.me link format: `https://wa.me/<number>` with the full international number, **digits only, no "+", no leading zeros, no brackets/dashes**; optional pre-filled message via URL-encoded `?text=` (spaces → %20). A WhatsApp QR is just this link as an image; scanning opens a chat with that number and the scanner needn't have the contact saved, but **must have WhatsApp installed**. — [Linkly](https://linklyhq.com/blog/wa-me-link); [toools.cloud](https://toools.cloud/phone/qr-code-generator/whatsapp)
- qrcode.js (davidshimjs) runs **fully client-side in the browser with no dependencies**, rendering via HTML5 Canvas or DOM table; MIT license; supports old and modern browsers. Original author no longer updates it. — [GitHub: davidshimjs/qrcodejs](https://github.com/davidshimjs/qrcodejs/); [npm: qrcodejs](https://npmjs.com/package/qrcodejs)
- EasyQRCodeJS: cross-browser, downloadable for direct in-browser use, renders Canvas/SVG/Table, supports logos/styling. — [Packagist: net-tools/easy-qrcode-js](https://root.packagist.org/packages/net-tools/easy-qrcode-js)
- Many tutorials (e.g. qr-code-styling) load the library from a CDN; for a fully offline page you must download the JS file and reference it locally. — [freeCodeCamp](https://www.freecodecamp.org/news/how-to-build-a-qr-code-generator-using-javascript)

### Inferences
- **Embedding the whole PDF in a QR is not feasible.** Even a 1-page payslip PDF is tens of kilobytes; QR max is ~2.9 KB. A QR can only hold (a) the payslip's *text data* as a compact string the worker reads on-screen after scanning, or (b) a *link/URI* (http(s), wa.me, tel:, sms:).
- A **text-data QR is the most "any-phone, offline, zero-data-cost" option**: any QR scanner/camera decodes it to plain text with no internet, no app login, no file to open. Downsides: no formatting, no official-looking document, limited to short content, and the worker must be able to read the digits/labels (mitigate with Arabic/Urdu/Hindi labels, but every language added consumes bytes).
- A **link QR (wa.me / https / tel / sms)** requires connectivity (and the relevant app) when scanned, so it shifts data cost and app-dependency onto the worker — contradicting the "low data / any phone" goal for the whole-payslip use case, though fine for a *verification* or *help* link.
- The static offline generator can bundle qrcode.js/EasyQRCodeJS locally and generate per-worker QR codes entirely in-browser with no network — a clean fit for the stated constraint.

### Gaps
- No source gives a verified byte capacity at error-correction level H (one sample suggested ~1,852 characters, unconfirmed). Treat high-ECC capacity as materially lower than the L figures.
- No source on `tel:`/`sms:` URI behavior in QR specifically; these are standard URI schemes recognized by phone cameras, but I found no citation in this search.

---

## Q2. PDF user-password security (AES-128/256, RC4), brute-forcing a 10-digit Iqama, and GCC payroll default-password practice

### Takeaway
Modern PDFs use AES-128 or AES-256; older ones used weak RC4. AES-256 revision 6 (PDF 2.0 / ISO 32000-2) hardens the key derivation, but **encryption strength does not fix a weak password**: a purely numeric Iqama password has a keyspace of at most 10^10, which is brute-forceable in minutes-to-days on a single consumer GPU depending on the PDF revision. Using the Iqama (or last-4+DOB) as the password is common GCC practice and acceptable as a convenience/privacy measure against casual snooping, but it is **not** strong confidentiality — the number is widely known/shared and the keyspace is tiny.

### Cited Findings
- PDF encryption types: modern PDFs use AES-128 or AES-256; older versions used the weaker RC4 algorithm. — [rootcircle blog](https://rootcircle.is-a.dev/blog/notes/pdfcrack.html)
- Revision 6 is the hardened key-derivation algorithm standardized in ISO 32000-2 for PDF 2.0, intended to fix a weakness in revision 5's password verification; older viewers/tools may fail to open rev-6 files. — [LosLab/HotPDF](https://blog.loslab.com/en-gb/pdf-programming/hotpdf-aes256-encryption-delphi-security.html)
- PDF 2.0 replaced the older MD5/RC4-based hashing with SHA-256-based hashing; AES-256 is described as the strongest standard PDF encryption. — [Mapsoft](https://mapsoft.com/importance-of-document-security)
- AES-256 rev 6 passwords are UTF-8, capped at 127 bytes, with Unicode support. — [Syncfusion](https://syncfusion.com/blogs/?p=17671); [LosLab/HotPDF](https://blog.loslab.com/en-gb/pdf-programming/hotpdf-aes256-encryption-delphi-security.html)
- Hashcat modes: 10500 = PDF 1.4–1.7 (Acrobat 5–8); 10600 = PDF 1.7 Level 3; 10700 = PDF 1.7 Level 8 (Acrobat 10–11, the AES-256 family). Workflow: extract hash with pdf2john, then run hashcat/John. — [rootcircle](https://rootcircle.is-a.dev/blog/notes/pdfcrack.html); [fr0stb1rd](https://fr0stb1rd.is-a.dev/posts/Cracking-PDF-Password-with-Hashcat/)
- Observed speeds (community, not a clean benchmark): mode 10500 on an RTX 3090 ≈ 66–74 kH/s in a mask attack; two GTX 1050 Ti ≈ 3.1 MH/s each ("perfect" benchmark speed via `hashcat -b -m 10500`). Mode 10700 on an old GTX 970 (hashcat v1.35) ≈ 14,537 H/s (14.5 kH/s). — [hashcat forum tid=11838](https://hashcat.net/forum/printthread.php?tid=11838); [hashcat forum tid=7021](https://hashcat.net/forum/printthread.php?tid=7021)
- GCC/Gulf payroll default-password schemes actually used (none confirmed to use Emirates/Iqama *last-4*; they use full employee ID, DOB, or passport):
  - HRWorks: full employee ID/No. + date and month of birth (DDMM); case-sensitive if ID alphanumeric. — [hinote/HRWorks](https://hinote.atlassian.net/wiki/spaces/HUG/pages/590058/Payslip+and+Tax+Sheet+Password)
  - BrioHR: password generated from employee data — date of birth or identification number — configurable per organization. — [BrioHR](https://support.briohr.com/knowledge/enhance-payslip-security-with-password-protected-payslips)
  - Sage Payroll Professional: employee ID number, passport number, or DOB (YYYYMMDD) depending on config. — [Sage community](https://communityhub.sage.com/cfs-file/__key/communityserver-discussions-components-files/386/Adding-passwords-to-payslips.pdf)

### Inferences (brute-force math — derived from the cited speeds)
- Saudi Iqama / National ID is **10 digits**; residents' Iqama conventionally starts with 2 (Saudi National IDs with 1), so if the leading digit is fixed the realistic keyspace is ~10^9; worst case (all 10 unknown) is 10^10.
- **AES-128 / older PDF (mode 10500), assume ~3 MH/s on one mid-range GPU:** 10^9 ≈ 333 s (~5.5 min); 10^10 ≈ 3,333 s (~55 min). On an RTX 4090-class card, minutes. → A numeric-only password on an AES-128/legacy PDF is effectively no protection against a motivated attacker.
- **AES-256 rev 6 (mode 10700), the hardened KDF, using the old 14.5 kH/s figure:** 10^9 ≈ ~69,000 s (~19 h); 10^10 ≈ ~8 days on a single old GPU. A modern GPU or small cluster collapses this to hours. The rev-6 KDF only buys a constant-factor slowdown; it does **not** enlarge a numeric keyspace.
- **Conclusion:** PDF-password protection with a numeric Iqama is reasonable *privacy* (stops a coworker or clerk casually opening someone else's slip) but is **weak confidentiality**. If stronger protection is wanted: use AES-256 rev 6 *and* a password with non-numeric entropy (e.g., Iqama + a site-specific word, or last-4 + DDMM-of-birth which still is only ~10^4×366 ≈ 3.7M keyspace — crackable in seconds at MH/s, so still weak). The honest framing is that any password a low-literacy worker can memorize from their own ID documents has small keyspace and is brute-forceable; the password's real job is deterrence and "need-to-know," not cryptographic secrecy.
- "Iqama number as password" is acceptable *practice* in the GCC convenience sense (it's memorable and unique per worker), but privacy-weak because the Iqama is printed on badges, shared with HR/banks/telecoms, and is short/numeric.

### Gaps
- No clean modern benchmark (RTX 3090/4090) for mode 10700 (AES-256 rev 6) was retrievable — the its5Q benchmark repo requires GitHub access not enabled in this session, and only an old GTX 970 figure was found. The time estimates above for AES-256 are order-of-magnitude, not precise; run `hashcat -b -m 10700` for exact current numbers.
- No source confirms any specific KDF iteration count for rev 6 in these results; the "iterated SHA-256/384/512 mix" is general knowledge, not cited here.
- Found no Saudi/GCC source confirming "Iqama number" as the literal default payslip password for a named vendor; the confirmed schemes use employee ID / DOB / passport.

---

## Q3. Per-person cloud folders (OneDrive / SharePoint / Google Drive) with manual sharing, no API

### Takeaway
Manual per-employee folder sharing is technically possible but **does not scale to 1,000 workers without scripting**, and it assumes each worker has a managed account and a smartphone/data to reach the file — a poor fit for a low-literacy, device-poor camp workforce. OneDrive's "Request files" is inbound/upload-only and therefore **useless for distributing** payslips. SharePoint with a script (PnP PowerShell / Graph) is the only realistic cloud route, but that edges into "API" territory the constraints exclude.

### Cited Findings
- Sharing a folder makes every file inside inherit the folder's permissions; "Anyone" links are view-only external links, "people in organization" links are internal; help guidance says to set a password (called mandatory for files with personal data — payroll qualifies). — [St Albans Diocese help](https://help.stalbansdiocese.org/how-do-i-share-a-file-or-folder-on-sharepoint-and-onedrive-to-a-colleague-or-external-user)
- A OneDrive folder can be locked so other licensed users can't see it by setting unique permissions on that folder; after unique permissions are set, other licensed users cannot access it. — [Microsoft Q&A](https://learn.microsoft.com/en-us/answers/questions/1137450/how-can-i-give-access-to-our-workers-to-only-some)
- Bulk-creating per-user folders via PowerShell is problematic; a Microsoft Tech Community poster reported multiple issues across PowerShell versions and no confirmed working one-click solution for user OneDrives (SharePoint libraries are the more scriptable target). — [Microsoft Tech Community](https://techcommunity.microsoft.com/discussions/windowspowershell/using-powershell-to-create-folders-within-users-onedrive/4181973/replies/4183232)
- Third-party automation (e.g., Yoom) offers prebuilt flows that create a folder per employee in OneDrive or SharePoint when an employee is added to an HR system — i.e., folder-per-employee is a recognized pattern but relies on connectors/automation. — [Yoom template](https://lp.yoom.fun/en/fb-templates/165043)
- OneDrive "Request files": creates an **upload-only** secure link to a chosen folder; recipients need no Microsoft account and cannot view/edit/delete folder contents; every uploaded file gets an uploader-name prefix. — [SharePoint Diary](https://www.sharepointdiary.com/?p=21714); [Chris Menard](https://chrismenardtraining.com/post/onedrive-request-files-how-to-use-onedrive-to-collect-files-securely/)

### Inferences
- **"Request files" is for collecting files, not distributing them** — it cannot be used to hand out payslips. Noting this prevents the report from recommending it.
- A no-API, fully manual workflow = HR creates 1,000 folders by hand, drops each slip in, and generates/sends 1,000 individual share links every month. That is on the order of several minutes per worker per month → tens of hours of HR labor monthly, plus a channel to deliver 1,000 unique links (which, without email/messaging automation, becomes printed link cards or QR-of-link — looping back to the QR/print options).
- Workers reaching a cloud link need: a smartphone, data, often a Microsoft/Google account or at least the ability to follow an "Anyone" link, and literacy to navigate the UI. For a multilingual low-literacy camp population this is the weakest of the surveyed options on usability.
- Google Drive behaves analogously (per-file/folder share links, "Anyone with the link" view-only) but has the same scale and account/data problems.

### Gaps
- No single authoritative Microsoft step-by-step for bulk per-employee folder + restricted link creation surfaced; the practical route (PnP PowerShell / Graph) was only inferred, and it uses an API, conflicting with the "no API" constraint.
- Microsoft 365 "Delve"/"Viva" payslip-specific features were asked about but no source surfaced; Delve has largely been deprecated/replaced by Viva, and no payslip-distribution feature was confirmed.

---

## Q4. Do Saudi worker populations have smartphones/data, and which messaging apps dominate?

### Takeaway
General Saudi connectivity is near-universal and WhatsApp-dominated, but **no survey broken down by migrant/blue-collar status** was found. Industry sources indicate low-income migrants typically own smartphones; data cost and WhatsApp voice/video restrictions are the recurring frictions.

### Cited Findings
- Saudi general internet penetration ~99.6% in 2025 (99% in 2024); average individual mobile data use ~53 GB/month. — [CST, via search summary](https://www.cst.gov.sa/sl/n05)
- Most-used apps (general population): **WhatsApp 93%**, Snapchat 83.6%, TikTok 77%. Not broken down by nationality/occupation. — [CST, via search summary](https://www.cst.gov.sa/sl/n05)
- WhatsApp voice/video calling has been restricted in Saudi Arabia for years; some recent user reports suggest calls may work for some, but no confirmed lasting policy change; availability varies by network. — [Roamless travel guide](https://roamless.com/ar/blog/does-whatsapp-work-in-saudi-arabia)
- Low-income migrants "typically have smartphones" but often earn below the salary floor to open a bank account (a Dhs5,000 floor cited in the UAE; an estimated 4.5M in the UAE and 26M across the GCC unable to get accounts). — [Gulf Business / CFI (commercial source)](https://cfi.co/awards/wp-json/wp/v2/posts/9570)
- Migrants are ~70% of the GCC workforce (UN data, via industry source). — [Gulf Business](https://gulfbusiness.com/the-app-connecting-the-gccs-low-paid-workers-with-bank-accounts/)
- GSMA (2025): mobile-money remittance inflows to the Middle East reached $1.8bn in 2024 (3.5× the 2021 value), indicating substantial smartphone-based money movement by migrants. — [GSMA](https://www.gsma.com/solutions-and-impact/connectivity-for-good/mobile-for-development/mobile-for-development-2/how-mobile-money-empowers-social-security-payments-via-migrants-remittances/)

### Inferences
- Smartphone ownership among the target workforce is probably high (text-chat and remittance behavior support this), but **data allowance is constrained and metered** — favoring zero-data options (text-in-QR, kiosk print, paper) over options that require downloading a file or opening a cloud link.
- WhatsApp is the near-certain dominant app, so a wa.me help/verification link QR would be recognizable to most workers — but the brief excludes messaging APIs, so WhatsApp can only be used as a *manually-scanned link*, not automated delivery.
- IMO and Botim (asked about) are historically popular among GCC migrants precisely because of VoIP restrictions, but **no citation was found in this search** to confirm current dominance.

### Gaps
- No ILO/World Bank/GASTAT survey of smartphone/data ownership specific to Saudi migrant or construction/services workers was found; general-population and commercial-source figures are the best available and may overstate the blue-collar camp segment.
- No citation confirming IMO/Botim usage levels among Saudi blue-collar workers.

---

## Q5. What do NEOM / MHRSD / Mudad / Qiwa / GCC bodies say about acceptable payslip delivery? Is printed payslip + signature sheet required?

### Takeaway
Saudi law requires paying wages via bank (WPS) in the official currency and, per vendor readings of the labour law, giving the worker a wage statement and retaining a copy — but **no source specifies the delivery medium** (paper vs. electronic vs. portal), and **no explicit requirement for printed payslip + signature sheets** was found. Mudad is a reporting/reconciliation platform, not a payslip-distribution channel. NEOM's own worker-welfare standards/payslip rules were not retrievable.

### Cited Findings
- Saudi Labour Law Article 90 (official): wages paid in official currency, monthly-wage workers paid once/month, via banks accredited in the Kingdom (WPS). — [HRSD](https://www.hrsd.gov.sa/en/%D8%B4%D8%B1%D9%88%D8%B7-%D8%A7%D9%84%D8%B9%D9%85%D9%84-%D9%88%D8%B8%D8%B1%D9%88%D9%81%D9%87)
- Vendor reading (Qoyod): the law obliges the company to hand the employee a document proving the wage and its details and to keep a copy in the personnel file for employment + 5 years; one Qoyod page attributes the payslip duty to **Article 70** (not 90). This is a vendor interpretation, not the statute text. — [Qoyod salary receipt](https://academy.qoyod.com/en/templates/salary-receipt/); [Qoyod wage dates](https://www.qoyod.com/en/learn/wage-payment-dates-final-settlement-saudi-labor-law/)
- Mudad: MHRSD wage-protection hub (with SAMA); all private-sector entities must report monthly payroll; monthly file includes each employee's ID details, contracted salary, actual salary paid, and authorized deductions; **cash payments not permitted**. Employees can confirm or dispute receipt through Mudad. — [greytHR Mudad/WPS guide](https://www.greythr.com/middle-east/blog/saudi-arabia-wps-mudad-compliance-guide-2026/); [Mercans](https://mercans.com/glossary/mudad-payroll-platform/)
- MHRSD mandates an ongoing monthly WPS compliance rate (sources conflict: one says ≥80%, another says falling below 90% triggers automated penalties as of 2026 — verify). — [greytHR](https://www.greythr.com/middle-east/blog/saudi-arabia-wps-mudad-compliance-guide-2026/); [myndsolution](https://www.myndsolution.com/best-practices/mudad-wps-salary-compliance-saudi-arabia/)
- NEOM: a spokesperson stated contractors and third parties must comply with NEOM's welfare standards; a WSJ-sourced report (via Business & Human Rights Resource Centre) documents welfare/safety problems and that NEOM reported 8 workplace fatalities in the prior year — but the underlying welfare standards and any payslip rule are not in the excerpt. — [Business & Human Rights Resource Centre](https://www.business-humanrights.org/en/latest-news/epic-construction-site-in-the-saudi-desert-is-a-hazard-for-workers/)

### Inferences
- Because Mudad already captures per-employee salary/deduction data and lets workers *confirm or dispute receipt*, MHRSD's visibility into "did the worker get paid" is at the payment level, and the *payslip document* delivery method appears to be left to the employer — so paper, PDF, kiosk, or portal are all plausibly compliant as long as a wage statement exists and a copy is retained.
- The retention duty (copy in personnel file + 5 years, per vendor reading) means whatever channel is chosen, HR must keep an auditable copy and ideally proof of delivery — supporting acknowledgement sheets / download logs as best practice even if not explicitly mandated.
- No evidence that labour inspectors *require* printed-slip-plus-signature specifically; but a signed acknowledgement sheet is the simplest defensible proof-of-delivery artifact for an inspection, which is why many camp operators use it. (Inference; not independently cited.)

### Gaps
- NEOM's actual supplier/contractor code or worker-welfare standard text, and any payslip-delivery clause, were not retrievable in this search.
- No official MHRSD/Mudad/Qiwa document specifying payslip *delivery medium* or mandating signature sheets was found — only vendor interpretations. The Article 70 vs 90 attribution should be verified against official HRSD text.
- WPS compliance-threshold figures conflict (80% vs 90%) across secondary sources.

---

## Q6. Kiosk / shared-PC patterns (worker enters ID to print own slip) and offline tooling

### Takeaway
Self-service HR kiosks that authenticate a worker (ID + OTP, RFID, or fingerprint) and print their own payslip are an established product pattern and map well to a camp admin-office PC — but no labour-camp-specific deployment was documented, and the per-slip print model needs an offline lookup+auth layer the static generator would have to provide.

### Cited Findings
- HR kiosk vendors support authentication by RFID card, fingerprint, or mobile phone, with an optional **national ID number + OTP via SMS**, and let employees print documents directly including payslips; recommended placement in production halls / access areas / high-traffic zones. — [Creasoft employee info kiosk](https://creasoft.ro/en/employees-info-kiosk)
- Workday admin guide: employers can let workers print their own payslips and can control whether online viewing/printing is enabled (e.g., "Disable Payslip Viewing (Online Version)" in self-service). — [Workday docs](https://doc.workday.com/admin-guide/en-us/payroll/payslips/aop1491007974493.html)

### Inferences
- A camp admin-office PC running the static offline generator could implement a kiosk mode: worker enters Iqama → page looks up and renders/prints that worker's slip locally, no internet. Authentication is the weak point — Iqama alone is guessable, so pair it with a second factor (DOB, or a PIN handed out by the supervisor) for privacy.
- This is the strongest fit for the zero-data, low-literacy, any-device constraints because the worker needs **no phone at all** — but it requires a staffed/located PC + printer and a queue-management process, and throughput for ~1,000 workers monthly must be planned (a few hundred prints over several days).
- Proof-of-delivery integrates naturally: the kiosk can log each print (timestamp + worker ID) and/or the worker signs an acknowledgement sheet at collection — producing both a download/print log and a signature record.

### Gaps
- No labour-camp/worker-accommodation-specific kiosk deployment or offline payslip-print toolkit was found; the pattern is assembled from general HR-kiosk and Workday self-service sources.
- No source on fingerprint/RFID feasibility or cost in a Saudi camp setting.

---

## Cross-cutting notes for the report writer

- **Proof of delivery options surfaced:** (a) signed acknowledgement sheets at hand-out/collection; (b) kiosk/print logs (timestamp + worker ID); (c) Mudad's built-in worker "confirm/dispute receipt" (payment-level, not payslip-document-level); (d) cloud share "download" telemetry (only if using OneDrive/SharePoint, which is otherwise a weak fit). The retention duty (copy + 5 years, per vendor reading of the labour law) reinforces keeping an auditable copy regardless of channel.
- **Best fit ranking for THIS workforce (inference from all findings):** (1) Kiosk/shared-PC print at the camp office — zero worker data/device, offline, logs built in; (2) Printed slip handed out by supervisor with signature sheet — zero tech, inspector-friendly, but labor-intensive and paper-privacy-weak unless slips are sealed; (3) Text-data QR on a printed stub the worker scans to read their numbers on any phone offline; (4) Password-protected PDF on a shared folder / USB — privacy is only as strong as the (numeric) password; (5) Per-employee cloud folders — weakest due to account/data/literacy burden and scale without scripting. OneDrive "Request files" is unusable for distribution (upload-only).
- **Security framing to carry forward:** AES-256 rev 6 is the strongest PDF option but only slows per-guess rate; a numeric Iqama/DOB password has a tiny keyspace (≤10^10, often ~10^9) and is brute-forceable in minutes (legacy/AES-128) to hours/days (AES-256 rev 6) on one consumer GPU. Treat PDF passwords as "need-to-know" privacy, not confidentiality.
