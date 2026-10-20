/**
 * OKV Technology Consults - OKV CBT back office (Google Sheets + Apps Script only)
 * ---------------------------------------------------------------------------
 * Handles: website + in-app forms, trial downloads, purchases (manual + Paystack),
 * activation codes (1000+ ready), email + alerts, customers, messaging, admin panel.
 * Setup: see SETUP.md. Run setup() once, then Deploy > Web app (Execute as: Me, Access: Anyone).
 */
const ALPHA = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
const INITIAL_ADMIN_PASSWORD = 'ChangeMe-OKV-2026';          // you must change it at first login
const DEFAULT_CODE_SECRET = 'OKV-CBT-LICENSE-SECRET-CHANGE-BEFORE-SELLING'; // MUST equal LICENSE_SECRET in okv_cbt.py
const SHEETS = {
  Submissions: ['id', 'time', 'type', 'status', 'name', 'school', 'email', 'phone', 'source', 'amount', 'pay_ref', 'pay_date', 'install_id', 'method', 'receipt_url', 'details', 'code', 'notes'],
  Customers: ['id', 'created', 'name', 'school', 'email', 'phone', 'status', 'install_id', 'code', 'last_contact', 'notes'],
  Codes: ['code', 'status', 'sent_to', 'sent_at', 'submission_id', 'activated_at', 'install_id', 'created'],
  Payments: ['id', 'time', 'method', 'reference', 'name', 'school', 'email', 'phone', 'amount', 'status', 'submission_id', 'code', 'raw'],
  Messages: ['id', 'time', 'channel', 'to', 'subject', 'body', 'status', 'by'],
  Alerts: ['id', 'time', 'type', 'title', 'body', 'read']
};
const DEFAULTS = {
  PRICE: '120000', OLD_PRICE: '', INSTALL_PRICE: 'Quote on request', TRIAL_DAYS: '12', TRIAL_ENABLED: 'true', PAYSTACK_ENABLED: 'true', MANUAL_ENABLED: 'true', PAYSTACK_PUBLIC: '', NO_REPLY: 'false',
  HERO_TITLE: '', HERO_SUB: '', DOWNLOAD_URL_WIN: '', DOWNLOAD_URL_MAC_ARM: '', DOWNLOAD_URL_MAC_INTEL: '', ANNOUNCEMENT: '',
  DOWNLOAD_URL: '', SITE_URL: 'https://www.okvtechnology.com', PAYSTACK_CALLBACK: '', WHATSAPP: '2348104141138',
  SENDER_NAME: 'OKV CBT System (OKV Technology Consults)', SIGNATURE: 'OKV Technology Consults\nWhatsApp: +234 810 414 1138\nEmail: technologyokv@gmail.com\nWebsite: www.okvtechnology.com',
  BANK: JSON.stringify({ bank_name: '', account_name: 'OKV Technology Consults', account_number: '', instructions: 'Use your organisation name as the transfer narration, then upload your receipt.' })
};
const TEMPLATES = {
  trial: ['Your OKV CBT free trial download', 'Hello {{name}},\n\nThank you for trying OKV CBT. Your 12-day free trial is ready.\n\nDownload: {{download_url}}\n\nOther computers:\n{{download_list}}\n\nHow to install:\nWindows - unzip the file, double-click OKV-CBT-Setup.exe and follow the wizard.\nMac - open the .dmg and drag OKV CBT into Applications (first time: right-click the app and choose Open).\nNothing else needs to be installed.\n\nThen sign in with admin / admin123 and change the password.\n4. Your trial starts on first run and runs for {{trial_days}} days.\n\nTo keep using it after the trial, activate from the Licence page on your admin dashboard (one-time {{price}}).\n\nWant us to install it for you? OKV Technology Consults offers an installation service (remote or on-site). Just reply to this email.\n\n{{signature}}'],
  paid: ['Payment confirmed - your OKV CBT activation code', 'Hello {{name}},\n\nThank you! Your Paystack payment of {{amount}} was received (reference {{ref}}). Here is your activation code:\n\n    {{code}}\n\nHow to activate:\n1. Sign in to OKV CBT as admin.\n2. Open Licence in the left menu.\n3. Type or paste the code in the box and press Activate.\n\nThis is a lifetime licence for one organisation installation. Keep this email safe and do not share the code. The code is also shown on the pricing page right after payment.\n\nNeed help installing or training your staff? Reply to this email - OKV Technology Consults offers an installation service.\n\n{{signature}}'],
  ack_purchase: ['We received your OKV CBT payment details', 'Hello {{name}},\n\nThank you. We have received your payment details{{school_line}}. We are verifying your payment and will email your activation code within 24 hours.\n\nReference: {{ref}}\n\n{{signature}}'],
  code: ['Your OKV CBT activation code', 'Hello {{name}},\n\nYour payment is confirmed. Here is your activation code:\n\n    {{code}}\n\nHow to activate:\n1. Sign in to OKV CBT as admin.\n2. Open Licence in the left menu.\n3. Type the code in the box and press Activate.\n\nThis is a lifetime licence for one organisation installation. Please keep this email safe and do not share the code.\n\nNeed help installing or training your staff? Reply to this email - OKV Technology Consults offers installation service.\n\n{{signature}}'],
  install_ack: ['We received your installation request', 'Hello {{name}},\n\nThank you for requesting the OKV CBT installation service. Our team will contact you shortly with a quote and available dates.\n\n{{signature}}'],
  contact_ack: ['We received your message', 'Hello {{name}},\n\nThanks for contacting OKV Technology Consults. We will reply shortly.\n\n{{signature}}']
};
const API_ = {};

/* ------------------------------------------------------------------ setup */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('OKV Admin').addItem('1. Run first-time setup', 'setup').addItem('Generate 1000 more codes', 'menuGen').addItem('Install automatic triggers', 'installTriggers').addToUi();
}
function menuGen() { generateCodes_(1000); SpreadsheetApp.getUi().alert('1000 codes added.'); }

function setup() {
  const P = PropertiesService.getScriptProperties();
  // The data sheet is created by this function. If the script is attached to a spreadsheet, that one is used;
  // otherwise the spreadsheet made by an earlier run is reused, or a new one is created.
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    const existing = P.getProperty('SHEET_ID');
    ss = existing ? SpreadsheetApp.openById(existing) : SpreadsheetApp.create('OKV CBT Back Office');
  }
  P.setProperty('SHEET_ID', ss.getId());
  Object.keys(DEFAULTS).forEach(k => { if (P.getProperty(k) === null) P.setProperty(k, DEFAULTS[k]); });
  if (!P.getProperty('CODE_SECRET')) P.setProperty('CODE_SECRET', DEFAULT_CODE_SECRET);
  if (!P.getProperty('ALERT_EMAIL')) P.setProperty('ALERT_EMAIL', Session.getEffectiveUser().getEmail() || 'technologyokv@gmail.com');
  if (!P.getProperty('ADMIN_HASH')) { P.setProperty('ADMIN_HASH', hash_(INITIAL_ADMIN_PASSWORD)); P.setProperty('ADMIN_MUST_CHANGE', '1'); }
  Object.keys(SHEETS).forEach(n => {
    let sh = ss.getSheetByName(n) || ss.insertSheet(n);
    if (sh.getLastRow() === 0) { sh.getRange(1, 1, 1, SHEETS[n].length).setValues([SHEETS[n]]).setFontWeight('bold').setBackground('#e8eefc'); sh.setFrozenRows(1); }
    sh.getRange(2, 1, Math.max(1, sh.getMaxRows() - 1), SHEETS[n].length).setNumberFormat('@');
  });
  const def = ss.getSheetByName('Sheet1'); if (def && ss.getSheets().length > 1 && def.getLastRow() === 0) ss.deleteSheet(def);
  const have = rows_('Codes').length;
  if (have < 1000) generateCodes_(1200 - have);
  installTriggers();
  Logger.log('Data sheet (open this link to see it): ' + ss.getUrl());
  Logger.log('Setup complete. Codes in pool: ' + rows_('Codes').length + '. Now Deploy > New deployment > Web app.');
  return ss.getUrl();
}
function installTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'verifyPending_') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('verifyPending_').timeBased().everyMinutes(10).create();
}

/* ------------------------------------------------------------------ helpers */
function P_(k, d) { const v = PropertiesService.getScriptProperties().getProperty(k); return v === null || v === '' ? (d === undefined ? '' : d) : v; }
function setP_(k, v) { PropertiesService.getScriptProperties().setProperty(k, String(v)); }
function ss_() { return SpreadsheetApp.openById(P_('SHEET_ID')); }
function sheet_(n) { return ss_().getSheetByName(n); }
function hash_(s) { return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'okv|' + s)); }
function uid_(p) { return p + Utilities.getUuid().replace(/-/g, '').slice(0, 8).toUpperCase(); }
function nowIso_() { return new Date().toISOString(); }
function str_(v, n) { return String(v == null ? '' : v).replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, '').trim().slice(0, n || 300); }
function rows_(name) {
  const sh = sheet_(name), vals = sh.getDataRange().getValues(), h = vals[0], out = [];
  for (let i = 1; i < vals.length; i++) {
    const o = { _row: i + 1 };
    for (let j = 0; j < h.length; j++) { let v = vals[i][j]; o[h[j]] = v instanceof Date ? v.toISOString() : (v === null ? '' : String(v)); }
    if (name === 'Codes' ? o.code : o.id) out.push(o);
  }
  return out;
}
function append_(name, obj) { const h = SHEETS[name]; sheet_(name).appendRow(h.map(k => obj[k] == null ? '' : String(obj[k]))); return obj; }
function update_(name, row, obj) {
  const sh = sheet_(name), h = SHEETS[name];
  Object.keys(obj).forEach(k => { const c = h.indexOf(k); if (c >= 0) sh.getRange(row, c + 1).setValue(String(obj[k] == null ? '' : obj[k])); });
}
function find_(name, key, val) { return rows_(name).filter(r => r[key] === val)[0]; }
function fill_(tpl, vars) { return tpl.replace(/\{\{(\w+)\}\}/g, (m, k) => vars[k] == null ? '' : vars[k]); }
function money_(n) { return '\u20A6' + Number(n || 0).toLocaleString('en-NG'); }
function getTemplate_(key) { try { const o = JSON.parse(P_('TEMPLATES', '{}')); if (o[key]) return o[key]; } catch (e) {} return TEMPLATES[key]; }
function mail_(to, subject, body, html) {
  if (!to) return false;
  try {
    const opts = { to: to, subject: subject, body: body, name: P_('SENDER_NAME') };
    if (P_('NO_REPLY') === 'true') opts.noReply = true; else opts.replyTo = P_('ALERT_EMAIL');
    opts.htmlBody = html || '<div style="font-family:Segoe UI,Arial,sans-serif;font-size:14px;line-height:1.55;color:#10213f;max-width:560px">' + body.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/\n/g, '<br>').replace(/ {4}(OKV-[A-Z0-9-]+|RST-[A-Z0-9-]+)/, '<span style="display:inline-block;font:700 20px monospace;background:#e8eefc;padding:8px 14px;border-radius:8px;margin:6px 0">$1</span>') + '</div>';
    MailApp.sendEmail(opts); return true;
  } catch (e) { Logger.log('mail failed: ' + e); return false; }
}
function sendTpl_(key, to, vars) {
  const t = getTemplate_(key), v = Object.assign({ signature: P_('SIGNATURE'), price: money_(P_('PRICE')), trial_days: P_('TRIAL_DAYS'), download_url: P_('DOWNLOAD_URL'), download_list: downloadList_(), ref: '', amount: '' }, vars);
  return mail_(to, fill_(t[0], v), fill_(t[1], v));
}
function dlUrl_(os) { return ({ windows: P_('DOWNLOAD_URL_WIN'), mac_arm: P_('DOWNLOAD_URL_MAC_ARM'), mac_intel: P_('DOWNLOAD_URL_MAC_INTEL') })[os] || P_('DOWNLOAD_URL'); }
function downloadList_() {
  return [['Windows 7/8/10/11', dlUrl_('windows')], ['Mac (Apple Silicon: M1/M2/M3/M4)', dlUrl_('mac_arm')], ['Mac (Intel)', dlUrl_('mac_intel')]].filter(r => r[1]).map(r => '  ' + r[0] + ': ' + r[1]).join('\n') || '  (see the website)';
}
function notify_(type, title, body) {
  append_('Alerts', { id: uid_('A'), time: nowIso_(), type: type, title: title, body: body, read: '' });
  try { mail_(P_('ALERT_EMAIL'), '[OKV CBT alert] ' + title, body + '\n\nOpen your admin panel to act on it.'); } catch (e) {}
  const tg = P_('TELEGRAM_TOKEN'), chat = P_('TELEGRAM_CHAT');
  if (tg && chat) { try { UrlFetchApp.fetch('https://api.telegram.org/bot' + tg + '/sendMessage', { method: 'post', payload: { chat_id: chat, text: title + '\n' + body }, muteHttpExceptions: true }); } catch (e) {} }
  hook_({ event: type, title: title, body: body });
}
function hook_(obj) {
  const u = P_('WEBHOOK_URL'); if (!u) return;
  try { UrlFetchApp.fetch(u, { method: 'post', contentType: 'application/json', payload: JSON.stringify(obj), muteHttpExceptions: true }); } catch (e) {}
}
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/* ------------------------------------------------------------------ activation codes (same algorithm as okv_cbt.py) */
function codeCheck_(body) { const sig = Utilities.computeHmacSha256Signature(body, P_('CODE_SECRET', DEFAULT_CODE_SECRET)); let s = ''; for (let i = 0; i < 4; i++) s += ALPHA[(sig[i] & 255) % 32]; return s; }
function resetCode_(installId) { const sig = Utilities.computeHmacSha256Signature('RESET' + installId, P_('CODE_SECRET', DEFAULT_CODE_SECRET)); let s = ''; for (let i = 0; i < 8; i++) s += ALPHA[(sig[i] & 255) % 32]; return 'RST-' + s.slice(0, 4) + '-' + s.slice(4); }
function normCode_(c) { return String(c || '').toUpperCase().replace(/[^A-Z0-9]/g, ''); }
function fmtCode_(c) { c = normCode_(c); return 'OKV-' + c.slice(3, 7) + '-' + c.slice(7, 11) + '-' + c.slice(11, 15) + '-' + c.slice(15, 19); }
function validCode_(c) { c = normCode_(c); return c.length === 19 && c.indexOf('OKV') === 0 && codeCheck_(c.slice(3, 15)) === c.slice(15, 19) && c.slice(3).split('').every(ch => ALPHA.indexOf(ch) >= 0); }
function generateCodes_(n) {
  const have = {}; rows_('Codes').forEach(r => have[normCode_(r.code)] = 1);
  const out = []; let guard = 0;
  while (out.length < n && guard++ < n * 5) {
    let body = ''; for (let i = 0; i < 12; i++) body += ALPHA[Math.floor(Math.random() * 32)];
    const c = 'OKV' + body + codeCheck_(body); if (have[c]) continue; have[c] = 1;
    out.push([fmtCode_(c), 'Unused', '', '', '', '', '', nowIso_()]);
  }
  if (out.length) { const sh = sheet_('Codes'); sh.getRange(sh.getLastRow() + 1, 1, out.length, 8).setNumberFormat('@').setValues(out); }
  return out.length;
}
function takeCode_(to, submissionId) {
  const free = rows_('Codes').filter(r => r.status === 'Unused')[0];
  if (!free) { generateCodes_(200); return takeCode_(to, submissionId); }
  update_('Codes', free._row, { status: 'Sent', sent_to: to, sent_at: nowIso_(), submission_id: submissionId || '' });
  return free.code;
}
function issueCode_(name, email, organisation, submissionId, paymentId, tplKey, vars) {
  const code = takeCode_(email, submissionId);
  const ok = sendTpl_(tplKey || 'code', email, Object.assign({ name: name || 'there', code: code, school: organisation }, vars || {}));
  if (submissionId) { const s = find_('Submissions', 'id', submissionId); if (s) update_('Submissions', s._row, { code: code, status: 'Handled' }); }
  if (paymentId) { const p = find_('Payments', 'id', paymentId); if (p) update_('Payments', p._row, { code: code, status: p.status === 'Paid' ? 'Paid' : 'Approved' }); }
  const c = rows_('Customers').filter(x => x.email.toLowerCase() === String(email).toLowerCase())[0];
  if (c) update_('Customers', c._row, { status: 'Paying', code: code, last_contact: nowIso_() });
  return { code: code, emailed: ok };
}

/* ------------------------------------------------------------------ public web endpoints */
function doGet(e) {
  const a = e && e.parameter && e.parameter.action;
  if (a === 'config') return out_({ ok: true, config: publicConfig_() });
  if (a === 'verify') return out_(verifyPaystack_(e.parameter.reference));
  if (a === 'ping') return out_({ ok: true, time: nowIso_() });
  return HtmlService.createHtmlOutputFromFile('Admin').setTitle('OKV CBT - Admin').addMetaTag('viewport', 'width=device-width, initial-scale=1').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function publicConfig_() {
  let bank = {}; try { bank = JSON.parse(P_('BANK', '{}')); } catch (e) {}
  return { price: Number(P_('PRICE', '120000')), old_price: Number(P_('OLD_PRICE', '0')) || 0, install_price: P_('INSTALL_PRICE', 'Quote on request'), trial_days: Number(P_('TRIAL_DAYS', '12')), announcement: P_('ANNOUNCEMENT'), download_url: P_('DOWNLOAD_URL'),
    downloads: { windows: dlUrl_('windows'), mac_arm: dlUrl_('mac_arm'), mac_intel: dlUrl_('mac_intel') }, hero_title: P_('HERO_TITLE'), hero_sub: P_('HERO_SUB'), trial_enabled: P_('TRIAL_ENABLED') !== 'false',
    manual_enabled: P_('MANUAL_ENABLED') !== 'false', paystack_enabled: P_('PAYSTACK_ENABLED') !== 'false' && !!P_('PAYSTACK_SECRET'), bank: P_('MANUAL_ENABLED') !== 'false' ? bank : {}, whatsapp: P_('WHATSAPP') };
}
function doPost(e) {
  try {
    const d = JSON.parse((e.postData && e.postData.contents) || '{}');
    const type = str_(d.type, 30);
    const email = str_(d.email, 120).toLowerCase();
    if (type !== 'activated' && !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) return out_({ ok: false, error: 'Please enter a valid email address.' });
    const cache = CacheService.getScriptCache(), key = 'rl_' + email, n = Number(cache.get(key) || 0);
    if (n >= 8) return out_({ ok: false, error: 'Too many requests from this email. Please try again in an hour or WhatsApp us.' });
    cache.put(key, String(n + 1), 3600);
    if (type === 'trial') return out_(onTrial_(d, email));
    if (type === 'installation' || type === 'contact') return out_(onSimple_(d, email, type));
    if (type === 'purchase' || type === 'activation') return out_(onPurchase_(d, email, type));
    if (type === 'paystack_init') return out_(onPaystackInit_(d, email));
    if (type === 'activated') return out_(onActivated_(d));
    return out_({ ok: false, error: 'Unknown request.' });
  } catch (err) { Logger.log(err); return out_({ ok: false, error: 'Server error. Please try again or WhatsApp us.' }); }
}
function upsertCustomer_(d, email, status) {
  const ex = rows_('Customers').filter(c => c.email.toLowerCase() === email)[0];
  if (ex) { const u = { last_contact: nowIso_() }; if (d.install_id) u.install_id = str_(d.install_id, 40); if (status && ex.status !== 'Active' && !(ex.status === 'Paying' && status === 'Trial')) u.status = status; update_('Customers', ex._row, u); return ex.id; }
  const id = uid_('C');
  append_('Customers', { id: id, created: nowIso_(), name: str_(d.name || d.contact_name, 100), school: str_(d.school || d.school_name, 150), email: email, phone: str_(d.phone, 40), status: status || 'Lead', install_id: str_(d.install_id, 40), code: '', last_contact: nowIso_(), notes: '' });
  return id;
}
function saveSub_(d, email, type, extra) {
  const id = uid_('S'), name = str_(d.name || d.contact_name, 100), organisation = str_(d.school || d.school_name, 150);
  const known = ['type', 'name', 'contact_name', 'school', 'school_name', 'email', 'phone', 'source', 'amount', 'pay_ref', 'pay_date', 'install_id', 'receipt', 'consent'];
  const det = {}; Object.keys(d).forEach(k => { if (known.indexOf(k) < 0) det[k] = str_(d[k], 500); });
  append_('Submissions', Object.assign({ id: id, time: nowIso_(), type: type, status: 'New', name: name, school: organisation, email: email, phone: str_(d.phone, 40), source: str_(d.source || 'website', 30), amount: str_(d.amount, 20), pay_ref: str_(d.pay_ref, 80), pay_date: str_(d.pay_date, 20), install_id: str_(d.install_id, 40), method: '', receipt_url: '', details: JSON.stringify(det), code: '', notes: '' }, extra || {}));
  return { id: id, name: name, school: organisation };
}
function onTrial_(d, email) {
  if (P_('TRIAL_ENABLED') === 'false') return { ok: false, error: 'Trial downloads are paused. Please WhatsApp us.' };
  const s = saveSub_(d, email, 'trial'); upsertCustomer_(d, email, 'Trial');
  sendTpl_('trial', email, { name: s.name || 'there', download_url: dlUrl_(str_(d.os, 20)) });
  const os = str_(d.os, 20), url = dlUrl_(os);
  notify_('trial', 'New free trial download: ' + (s.school || s.name), 'Name: ' + s.name + '\nOrganisation: ' + s.school + '\nEmail: ' + email + '\nPhone: ' + str_(d.phone, 40) + '\nPlatform: ' + os + '\nRole: ' + str_(d.role, 60) + '\nStudents: ' + str_(d.students, 40) + (d.install_help ? '\n** Wants installation help **' : ''));
  return { ok: true, download_url: url };
}
function onSimple_(d, email, type) {
  const s = saveSub_(d, email, type); upsertCustomer_(d, email, 'Lead');
  sendTpl_(type === 'installation' ? 'install_ack' : 'contact_ack', email, { name: s.name || 'there' });
  notify_(type, (type === 'installation' ? 'INSTALLATION request: ' : 'New message: ') + (s.school || s.name), 'Name: ' + s.name + '\nOrganisation: ' + s.school + '\nEmail: ' + email + '\nPhone: ' + str_(d.phone, 40) + '\n' + (d.message ? 'Message: ' + str_(d.message, 1000) + '\n' : '') + JSON.stringify(JSON.parse(find_('Submissions', 'id', s.id).details || '{}')));
  return { ok: true };
}
function saveReceipt_(r, name) {
  if (!r || !r.data) return '';
  if (r.data.length > 6 * 1024 * 1024) throw new Error('Receipt is too large.');
  let fid = P_('RECEIPT_FOLDER_ID'), folder;
  try { folder = fid ? DriveApp.getFolderById(fid) : null; } catch (e) { folder = null; }
  if (!folder) { folder = DriveApp.createFolder('OKV CBT Receipts'); setP_('RECEIPT_FOLDER_ID', folder.getId()); }
  const f = folder.createFile(Utilities.newBlob(Utilities.base64Decode(r.data), r.mime || 'application/octet-stream', str_(name, 40).replace(/[^\w.-]+/g, '_') + '_' + str_(r.name, 60)));
  return f.getUrl();
}
function onPurchase_(d, email, type) {
  if (P_('MANUAL_ENABLED') === 'false') return { ok: false, error: 'Bank-transfer payments are turned off. Please pay with Paystack.' };
  const url = saveReceipt_(d.receipt, d.school || d.school_name || d.name);
  const s = saveSub_(d, email, type, { method: 'bank_transfer', receipt_url: url }); upsertCustomer_(d, email, 'Lead');
  const pid = uid_('P');
  append_('Payments', { id: pid, time: nowIso_(), method: 'bank_transfer', reference: str_(d.pay_ref, 80), name: s.name, school: s.school, email: email, phone: str_(d.phone, 40), amount: str_(d.amount || P_('PRICE'), 20), status: 'Pending', submission_id: s.id, code: '', raw: url });
  sendTpl_('ack_purchase', email, { name: s.name || 'there', ref: str_(d.pay_ref, 80) || s.id, school_line: s.school ? ' for ' + s.school : '' });
  notify_('payment', 'PAYMENT receipt: ' + (s.school || s.name) + ' ' + money_(d.amount), 'Needs your approval within 24 hours.\nName: ' + s.name + '\nOrganisation: ' + s.school + '\nEmail: ' + email + '\nPhone: ' + str_(d.phone, 40) + '\nAmount: ' + money_(d.amount) + '\nReference: ' + str_(d.pay_ref, 80) + '\nInstall ID: ' + str_(d.install_id, 40) + '\nReceipt: ' + url + '\nFrom: ' + str_(d.source, 20));
  return { ok: true };
}
function onActivated_(d) {
  const c = fmtCode_(d.code || '');
  if (!validCode_(c)) return { ok: false, error: 'bad code' };
  const r = find_('Codes', 'code', c);
  if (r) update_('Codes', r._row, { status: 'Activated', activated_at: nowIso_(), install_id: str_(d.install_id, 40) });
  const cu = rows_('Customers').filter(x => x.code === c)[0]; if (cu) update_('Customers', cu._row, { status: 'Active', install_id: str_(d.install_id, 40) });
  notify_('activated', 'Activated: ' + str_(d.school_name, 100), 'Code ' + c + ' was activated on install ' + str_(d.install_id, 40));
  return { ok: true };
}

/* ------------------------------------------------------------------ Paystack */
function ps_(path, opt) {
  const sk = P_('PAYSTACK_SECRET'); if (!sk) throw new Error('Paystack secret key is not set.');
  const o = Object.assign({ headers: { Authorization: 'Bearer ' + sk }, muteHttpExceptions: true }, opt || {});
  return JSON.parse(UrlFetchApp.fetch('https://api.paystack.co' + path, o).getContentText());
}
function onPaystackInit_(d, email) {
  if (!P_('PAYSTACK_SECRET') || P_('PAYSTACK_ENABLED') === 'false') return { ok: false, error: 'Online payment is not available right now. Please pay by bank transfer.' };
  const price = Number(P_('PRICE', '120000')), ref = 'OKV' + Date.now().toString(36).toUpperCase() + Utilities.getUuid().replace(/-/g, '').slice(0, 8).toUpperCase();
  const cb = P_('PAYSTACK_CALLBACK') || (P_('SITE_URL') + '/pricing.html');
  const r = ps_('/transaction/initialize', { method: 'post', contentType: 'application/json', payload: JSON.stringify({ email: email, amount: price * 100, currency: 'NGN', reference: ref, callback_url: cb, metadata: { name: str_(d.name, 100), school: str_(d.school, 150), phone: str_(d.phone, 40), product: 'OKV CBT lifetime licence' } }) });
  if (!r.status) return { ok: false, error: 'Paystack: ' + (r.message || 'could not start payment') };
  const s = saveSub_(d, email, 'purchase', { method: 'paystack', pay_ref: ref, amount: String(price) }); upsertCustomer_(d, email, 'Lead');
  append_('Payments', { id: uid_('P'), time: nowIso_(), method: 'paystack', reference: ref, name: s.name, school: s.school, email: email, phone: str_(d.phone, 40), amount: String(price), status: 'Pending', submission_id: s.id, code: '', raw: '' });
  notify_('payment_started', 'Paystack checkout started: ' + (s.school || s.name), 'Reference ' + ref + ' - ' + email);
  return { ok: true, authorization_url: r.data.authorization_url, reference: ref };
}
function verifyPaystack_(ref) {
  ref = str_(ref, 80); if (!ref) return { ok: false };
  const p = find_('Payments', 'reference', ref); if (!p) return { ok: false, error: 'Unknown reference' };
  if (p.status === 'Paid' && p.code) return { ok: true, paid: true, email: p.email, code: p.code };
  const r = ps_('/transaction/verify/' + encodeURIComponent(ref));
  if (r.status && r.data && r.data.status === 'success' && Number(r.data.amount) >= Number(P_('PRICE', '120000')) * 100) {
    update_('Payments', p._row, { status: 'Paid', raw: JSON.stringify({ paid_at: r.data.paid_at, channel: r.data.channel, amount: r.data.amount }) });
    const res = issueCode_(p.name, p.email, p.school, p.submission_id, p.id, 'paid', { ref: ref, amount: money_(Number(r.data.amount) / 100) });
    notify_('paid', 'PAID via Paystack: ' + (p.school || p.name) + ' ' + money_(Number(r.data.amount) / 100), 'Reference ' + ref + '\nActivation code ' + res.code + ' was shown to the customer and emailed to ' + p.email + ' automatically.');
    return { ok: true, paid: true, email: p.email, code: res.code };
  }
  if (r.status && r.data && (r.data.status === 'failed' || r.data.status === 'abandoned')) update_('Payments', p._row, { status: 'Failed' });
  return { ok: true, paid: false };
}
function verifyPending_() {
  rows_('Payments').filter(p => p.method === 'paystack' && p.status === 'Pending' && (Date.now() - new Date(p.time).getTime()) < 3 * 86400000).forEach(p => { try { verifyPaystack_(p.reference); } catch (e) {} });
}

/* ------------------------------------------------------------------ admin API */
function adminLogin(pw) {
  const cache = CacheService.getScriptCache(), fails = Number(cache.get('fails') || 0);
  if (fails >= 6) throw new Error('Too many wrong attempts. Wait 15 minutes.');
  if (hash_(pw || '') !== P_('ADMIN_HASH')) { cache.put('fails', String(fails + 1), 900); throw new Error('Wrong password.'); }
  const t = Utilities.getUuid(); cache.put('tok_' + t, '1', 21600); cache.remove('fails');
  return { token: t, must_change: P_('ADMIN_MUST_CHANGE') === '1' };
}
function adminApi(token, method, args) {
  if (!token || !CacheService.getScriptCache().get('tok_' + token)) throw new Error('Session expired. Please sign in again.');
  if (!API_[method]) throw new Error('Unknown action.');
  return API_[method].apply(null, args || []);
}
const day_ = iso => String(iso).slice(0, 10);
API_.stats = function () {
  const sub = rows_('Submissions'), pay = rows_('Payments'), codes = rows_('Codes'), cust = rows_('Customers'), al = rows_('Alerts');
  const cnt = k => sub.filter(s => s.type === k).length, days = [];
  for (let i = 13; i >= 0; i--) { const d = new Date(Date.now() - i * 86400000).toISOString().slice(0, 10); days.push({ l: d.slice(5), v: sub.filter(s => day_(s.time) === d).length }); }
  const months = []; for (let i = 5; i >= 0; i--) { const dt = new Date(); dt.setMonth(dt.getMonth() - i); const m = dt.toISOString().slice(0, 7); months.push({ l: m.slice(2), v: pay.filter(p => (p.status === 'Paid' || p.status === 'Approved') && p.time.slice(0, 7) === m).reduce((a, p) => a + Number(p.amount || 0), 0) }); }
  const bc = s => codes.filter(c => c.status === s).length;
  return { total: sub.length, fresh: sub.filter(s => s.status === 'New').length, trial: cnt('trial'), purchase: cnt('purchase') + cnt('activation'), installation: cnt('installation'), contact: cnt('contact'),
    revenue: pay.filter(p => p.status === 'Paid' || p.status === 'Approved').reduce((a, p) => a + Number(p.amount || 0), 0), pendingPay: pay.filter(p => p.status === 'Pending').length,
    codes: { total: codes.length, unused: bc('Unused'), sent: bc('Sent'), activated: bc('Activated') }, customers: cust.length, active: cust.filter(c => c.status === 'Active').length,
    custBy: ['Lead', 'Trial', 'Paying', 'Active'].map(s => ({ l: s, v: cust.filter(c => c.status === s).length })), days: days, months: months, unread: al.filter(a => !a.read).length, recent: al.slice(-8).reverse(), pendingList: pay.filter(p => p.status === 'Pending' && p.method !== 'paystack').slice(-6).reverse() };
};
API_.listSubmissions = function (f) {
  f = f || {}; const q = (f.q || '').toLowerCase();
  return rows_('Submissions').filter(s => (!f.type || s.type === f.type) && (!f.status || s.status === f.status) && (!q || (s.name + s.school + s.email + s.phone).toLowerCase().indexOf(q) >= 0)).reverse().slice(0, 400);
};
API_.updateSubmission = function (id, u) { const s = find_('Submissions', 'id', id); if (!s) throw new Error('Not found'); const o = {}; ['status', 'notes'].forEach(k => { if (u[k] != null) o[k] = str_(u[k], 1000); }); update_('Submissions', s._row, o); return 1; };
API_.deleteSubmission = function (id) { const s = find_('Submissions', 'id', id); if (s) sheet_('Submissions').deleteRow(s._row); return 1; };
API_.sendCodeForSubmission = function (id) {
  const s = find_('Submissions', 'id', id); if (!s) throw new Error('Not found');
  if (s.code) throw new Error('A code was already sent: ' + s.code);
  const p = rows_('Payments').filter(x => x.submission_id === id)[0];
  return issueCode_(s.name, s.email, s.school, id, p && p.id);
};
API_.approveAndSend = function (submissionId) {
  const s = find_('Submissions', 'id', submissionId); if (!s) throw new Error('Not found');
  if (s.code) throw new Error('A code was already sent: ' + s.code);
  let p = rows_('Payments').filter(x => x.submission_id === submissionId)[0];
  if (!p) { const id = uid_('P'); append_('Payments', { id: id, time: nowIso_(), method: s.method || 'bank_transfer', reference: s.pay_ref, name: s.name, school: s.school, email: s.email, phone: s.phone, amount: s.amount || P_('PRICE'), status: 'Pending', submission_id: submissionId, code: '', raw: s.receipt_url }); p = { id: id }; }
  return issueCode_(s.name, s.email, s.school, submissionId, p.id);
};
API_.listCustomers = function (q) { q = (q || '').toLowerCase(); return rows_('Customers').filter(c => !q || (c.name + c.school + c.email + c.phone).toLowerCase().indexOf(q) >= 0).reverse().slice(0, 500); };
API_.saveCustomer = function (c) {
  c.email = str_(c.email, 120).toLowerCase(); if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(c.email)) throw new Error('Enter a valid email.');
  const o = { name: str_(c.name, 100), school: str_(c.school, 150), email: c.email, phone: str_(c.phone, 40), status: str_(c.status || 'Lead', 20), install_id: str_(c.install_id, 40), notes: str_(c.notes, 1000) };
  if (c.id) { const x = find_('Customers', 'id', c.id); if (!x) throw new Error('Not found'); update_('Customers', x._row, o); } else append_('Customers', Object.assign({ id: uid_('C'), created: nowIso_(), code: '', last_contact: '' }, o));
  return 1;
};
API_.deleteCustomer = function (id) { const x = find_('Customers', 'id', id); if (x) sheet_('Customers').deleteRow(x._row); return 1; };
API_.listPayments = function () { return rows_('Payments').reverse().slice(0, 400); };
API_.approvePayment = function (id) { const p = find_('Payments', 'id', id); if (!p) throw new Error('Not found'); if (p.code) throw new Error('Code already sent: ' + p.code); return issueCode_(p.name, p.email, p.school, p.submission_id, id); };
API_.rejectPayment = function (id, reason) { const p = find_('Payments', 'id', id); if (!p) throw new Error('Not found'); update_('Payments', p._row, { status: 'Rejected', raw: str_(reason, 300) }); if (reason) mail_(p.email, 'About your OKV CBT payment', 'Hello ' + p.name + ',\n\nWe could not confirm your payment: ' + reason + '\n\nPlease reply or WhatsApp us on +234 810 414 1138.\n\n' + P_('SIGNATURE')); return 1; };
API_.verifyPayment = function (id) { const p = find_('Payments', 'id', id); if (!p) throw new Error('Not found'); if (p.method !== 'paystack') throw new Error('Only Paystack payments can be auto-verified.'); return verifyPaystack_(p.reference); };
API_.verifyAllPending = function () { verifyPending_(); return rows_('Payments').filter(p => p.status === 'Pending').length; };
API_.addManualPayment = function (p) {
  p.email = str_(p.email, 120).toLowerCase(); if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(p.email)) throw new Error('Enter a valid email.');
  const id = uid_('P'); append_('Payments', { id: id, time: nowIso_(), method: str_(p.method || 'cash/transfer', 30), reference: str_(p.reference, 80), name: str_(p.name, 100), school: str_(p.school, 150), email: p.email, phone: str_(p.phone, 40), amount: str_(p.amount || P_('PRICE'), 20), status: 'Pending', submission_id: '', code: '', raw: 'added by admin' });
  upsertCustomer_(p, p.email, 'Lead'); if (p.issue) return issueCode_(p.name, p.email, p.school, '', id); return { code: '' };
};
API_.codeStats = function () { const c = rows_('Codes'); const b = s => c.filter(x => x.status === s).length; return { total: c.length, unused: b('Unused'), sent: b('Sent'), activated: b('Activated') }; };
API_.listCodes = function (f) {
  f = f || {}; const q = (f.q || '').toUpperCase(); let r = rows_('Codes').filter(c => (!f.status || c.status === f.status) && (!q || (c.code + c.sent_to).toUpperCase().indexOf(q) >= 0));
  return { total: r.length, rows: r.slice(0, 300) };
};
API_.exportCodes = function (status) { return rows_('Codes').filter(c => !status || c.status === status).map(c => [c.code, c.status, c.sent_to, c.sent_at, c.activated_at]); };
API_.generateCodes = function (n) { n = Math.max(1, Math.min(Math.floor(Number(n)) || 1000, 3000)); const made = generateCodes_(n); const c = rows_('Codes'); return { made: made, total: c.length, unused: c.filter(x => x.status === 'Unused').length }; };
API_.sendCode = function (code, email, name) {
  email = str_(email, 120).toLowerCase(); if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) throw new Error('Enter a valid email.');
  const r = find_('Codes', 'code', code); if (!r) throw new Error('Code not found.'); if (r.status !== 'Unused') throw new Error('This code is already ' + r.status + '.');
  update_('Codes', r._row, { status: 'Sent', sent_to: email, sent_at: nowIso_() });
  const ok = sendTpl_('code', email, { name: name || 'there', code: r.code });
  const c = rows_('Customers').filter(x => x.email.toLowerCase() === email)[0]; if (c) update_('Customers', c._row, { code: r.code, status: 'Paying', last_contact: nowIso_() });
  return { code: r.code, emailed: ok };
};
API_.sendNextCode = function (email, name) { const r = rows_('Codes').filter(c => c.status === 'Unused')[0]; if (!r) throw new Error('No unused codes. Generate more.'); return API_.sendCode(r.code, email, name); };
API_.markCode = function (code, status) { const r = find_('Codes', 'code', code); if (!r) throw new Error('Not found'); update_('Codes', r._row, { status: status }); return 1; };
API_.sendMessage = function (m) {
  const ch = m.channel || 'email', body = str_(m.body, 5000), subj = str_(m.subject, 150);
  let people = [];
  if (m.audience === 'custom') people = String(m.emails || '').split(/[\s,;]+/).filter(x => /@/.test(x)).map(e => ({ name: '', school: '', email: e.toLowerCase(), phone: '' }));
  else people = rows_('Customers').filter(c => m.audience === 'all' || (m.audience === 'ids' && (m.ids || []).indexOf(c.id) >= 0) || c.status.toLowerCase() === String(m.audience).toLowerCase());
  if (!people.length) throw new Error('No recipients match.');
  const v = p => ({ name: p.name || 'there', school: p.school, signature: P_('SIGNATURE') });
  if (ch === 'email') {
    const left = MailApp.getRemainingDailyQuota(); if (people.length > left) throw new Error('Only ' + left + ' emails left in today\'s Google limit. Select fewer people.');
    let sent = 0; people.forEach(p => { if (p.email && mail_(p.email, fill_(subj, v(p)), fill_(body, v(p)))) sent++; });
    append_('Messages', { id: uid_('M'), time: nowIso_(), channel: 'email', to: people.length + ' recipient(s)', subject: subj, body: body, status: 'Sent ' + sent, by: 'admin' });
    return { sent: sent, left: MailApp.getRemainingDailyQuota() };
  }
  const links = people.filter(p => p.phone).map(p => { const ph = String(p.phone).replace(/\D/g, '').replace(/^0/, '234'); const t = encodeURIComponent(fill_(body, v(p))); return { name: p.name || p.email, link: ch === 'whatsapp' ? 'https://wa.me/' + ph + '?text=' + t : ch === 'sms' ? 'sms:+' + ph + '?body=' + t : 'tel:+' + ph }; });
  append_('Messages', { id: uid_('M'), time: nowIso_(), channel: ch, to: links.length + ' recipient(s)', subject: subj, body: body, status: 'Links created', by: 'admin' });
  return { links: links };
};
API_.listMessages = function () { return rows_('Messages').reverse().slice(0, 200); };
API_.getAlerts = function (afterId) { const a = rows_('Alerts'); const i = afterId ? a.findIndex(x => x.id === afterId) : -1; return { alerts: a.slice(i + 1).slice(-20), unread: a.filter(x => !x.read).length }; };
API_.listAlerts = function () { return rows_('Alerts').reverse().slice(0, 100); };
API_.markAlertsRead = function () { const sh = sheet_('Alerts'), n = sh.getLastRow() - 1; if (n > 0) sh.getRange(2, 6, n, 1).setValue('1'); return 1; };
API_.getSettings = function () {
  let bank = {}; try { bank = JSON.parse(P_('BANK', '{}')); } catch (e) {} let tpl = {}; try { tpl = JSON.parse(P_('TEMPLATES', '{}')); } catch (e) {}
  const t = {}; Object.keys(TEMPLATES).forEach(k => t[k] = tpl[k] || TEMPLATES[k]);
  return { PRICE: P_('PRICE'), OLD_PRICE: P_('OLD_PRICE'), INSTALL_PRICE: P_('INSTALL_PRICE'), HERO_TITLE: P_('HERO_TITLE'), HERO_SUB: P_('HERO_SUB'), PAYSTACK_PUBLIC: P_('PAYSTACK_PUBLIC'), DOWNLOAD_URL_WIN: P_('DOWNLOAD_URL_WIN'), DOWNLOAD_URL_MAC_ARM: P_('DOWNLOAD_URL_MAC_ARM'), DOWNLOAD_URL_MAC_INTEL: P_('DOWNLOAD_URL_MAC_INTEL'),
    MANUAL_ENABLED: P_('MANUAL_ENABLED') !== 'false', NO_REPLY: P_('NO_REPLY') === 'true', TRIAL_DAYS: P_('TRIAL_DAYS'), TRIAL_ENABLED: P_('TRIAL_ENABLED') !== 'false', PAYSTACK_ENABLED: P_('PAYSTACK_ENABLED') !== 'false', ANNOUNCEMENT: P_('ANNOUNCEMENT'), DOWNLOAD_URL: P_('DOWNLOAD_URL'),
    SITE_URL: P_('SITE_URL'), PAYSTACK_CALLBACK: P_('PAYSTACK_CALLBACK'), WHATSAPP: P_('WHATSAPP'), SENDER_NAME: P_('SENDER_NAME'), SIGNATURE: P_('SIGNATURE'), ALERT_EMAIL: P_('ALERT_EMAIL'), BANK: bank, TEMPLATES: t,
    WEBHOOK_URL: P_('WEBHOOK_URL'), TELEGRAM_CHAT: P_('TELEGRAM_CHAT'), has_paystack_secret: !!P_('PAYSTACK_SECRET'), has_telegram: !!P_('TELEGRAM_TOKEN'), has_code_secret: P_('CODE_SECRET') !== DEFAULT_CODE_SECRET,
    secret_fp: Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, P_('CODE_SECRET', DEFAULT_CODE_SECRET))).slice(0, 8), webapp: ScriptApp.getService().getUrl(), quota: MailApp.getRemainingDailyQuota() };
};
API_.saveSettings = function (s) {
  ['PRICE', 'OLD_PRICE', 'INSTALL_PRICE', 'HERO_TITLE', 'HERO_SUB', 'PAYSTACK_PUBLIC', 'DOWNLOAD_URL_WIN', 'DOWNLOAD_URL_MAC_ARM', 'DOWNLOAD_URL_MAC_INTEL', 'TRIAL_DAYS', 'ANNOUNCEMENT', 'DOWNLOAD_URL', 'SITE_URL', 'PAYSTACK_CALLBACK', 'WHATSAPP', 'SENDER_NAME', 'SIGNATURE', 'ALERT_EMAIL', 'WEBHOOK_URL', 'TELEGRAM_CHAT'].forEach(k => { if (s[k] != null) setP_(k, str_(s[k], 1500)); });
  ['TRIAL_ENABLED', 'PAYSTACK_ENABLED', 'MANUAL_ENABLED', 'NO_REPLY'].forEach(k => { if (s[k] != null) setP_(k, s[k] ? 'true' : 'false'); });
  if (s.PAYSTACK_SECRET) setP_('PAYSTACK_SECRET', str_(s.PAYSTACK_SECRET, 200)); if (s.TELEGRAM_TOKEN) setP_('TELEGRAM_TOKEN', str_(s.TELEGRAM_TOKEN, 200));
  if (s.CODE_SECRET) { if (rows_('Codes').filter(c => c.status !== 'Unused').length) throw new Error('Codes were already sent. Changing the secret would break them.'); setP_('CODE_SECRET', str_(s.CODE_SECRET, 200)); sheet_('Codes').getRange(2, 1, Math.max(1, sheet_('Codes').getLastRow() - 1), 8).clearContent(); generateCodes_(1200); }
  if (s.BANK) setP_('BANK', JSON.stringify({ bank_name: str_(s.BANK.bank_name, 80), account_name: str_(s.BANK.account_name, 100), account_number: str_(s.BANK.account_number, 30), instructions: str_(s.BANK.instructions, 400) }));
  if (s.TEMPLATES) setP_('TEMPLATES', JSON.stringify(s.TEMPLATES));
  return 1;
};
API_.changePassword = function (oldPw, newPw) { if (hash_(oldPw) !== P_('ADMIN_HASH')) throw new Error('Current password is wrong.'); if (String(newPw).length < 8) throw new Error('Use at least 8 characters.'); setP_('ADMIN_HASH', hash_(newPw)); setP_('ADMIN_MUST_CHANGE', '0'); return 1; };
API_.testEmail = function () { return mail_(P_('ALERT_EMAIL'), 'OKV CBT test email', 'This is a test from your OKV admin panel.\n\n' + P_('SIGNATURE')); };
API_.testPaystack = function () { const r = ps_('/balance'); if (!r.status) throw new Error(r.message || 'Paystack rejected the key'); return 'Paystack key works.'; };
API_.testWebhook = function () { if (!P_('WEBHOOK_URL')) throw new Error('No webhook URL set.'); hook_({ event: 'test', title: 'OKV test' }); return 'Sent.'; };
API_.testTelegram = function () { if (!P_('TELEGRAM_TOKEN')) throw new Error('No Telegram token set.'); const r = JSON.parse(UrlFetchApp.fetch('https://api.telegram.org/bot' + P_('TELEGRAM_TOKEN') + '/sendMessage', { method: 'post', payload: { chat_id: P_('TELEGRAM_CHAT'), text: 'OKV CBT test alert' }, muteHttpExceptions: true }).getContentText()); if (!r.ok) throw new Error(r.description); return 'Sent.'; };
API_.resetCodeFor = function (installId) { installId = str_(installId, 40).toUpperCase(); if (!/^OKV-[0-9A-F]{8}$/.test(installId)) throw new Error('An Install ID looks like OKV-1A2B3C4D.'); return resetCode_(installId); };
API_.checkCode = function (c) { const r = find_('Codes', 'code', fmtCode_(c)); return { valid: validCode_(c), in_pool: !!r, status: r ? r.status : '', sent_to: r ? r.sent_to : '' }; };
API_.exportSubmissions = function () { return rows_('Submissions').map(s => SHEETS.Submissions.filter(k => k !== 'details').map(k => s[k])); };
