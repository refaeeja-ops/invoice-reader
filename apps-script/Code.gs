/**
 * قارئ الفواتير ⇄ Google Sheet (مجاني: Google Apps Script)
 *
 * يُلصق هذا الملف في الشيت: Extensions (الإضافات) ← Apps Script.
 * - يستقبل الفواتير المعتمدة من التطبيق ويضيفها إلى InvoiceT و VatReport (بدون تكرار).
 * - يعطي التطبيق البيانات المرجعية (الموردين والمشاريع والمنتجات) ليقرأها مباشرة من الشيت.
 * - يحفظ «قوالب الموردين» (أماكن الخانات) في ورقة VendorTemplates حتى تظهر على كل الأجهزة.
 * - التحديث مسموح فقط لبريد إلكتروني أضافه المالك من قائمة «قارئ الفواتير» (لكل بريد رمز خاص)، ويُسجَّل البريد في EnteredBy.
 */

const REF_SHEETS = ['SuppliersT', 'ProjectsT', 'ProductsT', 'SupplierProductT'];
// ورقة بصمات الملفات (تُنشأ تلقائيًا): تمنع إرسال نفس ملف الفاتورة مرتين من أي جهاز أو مستخدم
const HASH_SHEET = 'FileHashes', HASH_HEAD = ['FileHash', 'InvID', 'FilePath', 'EnteredBy', 'EnteredAt'];
function hashSheet_(ss) {
  let sh = sheet_(ss, HASH_SHEET);
  if (!sh) { sh = ss.insertSheet(HASH_SHEET); sh.getRange(1, 1, 1, HASH_HEAD.length).setValues([HASH_HEAD]); }
  return sh;
}
function hashes_(ss) { // {hash: InvID}
  const out = {}; values_(sheet_(ss, HASH_SHEET)).slice(1).forEach(r => { const h = str_(r[0]).toLowerCase(); if (h) out[h] = r[1]; }); return out;
}

// ورقة قوالب الموردين (تُنشأ تلقائيًا ولا تمس أي ورقة أخرى): صف لكل خانة في قالب كل مورد
const TPL_SHEET = 'VendorTemplates', TPL_HEAD = ['Key', 'VendorName', 'VATID', 'NameKey', 'Field', 'Page', 'X', 'Y', 'W', 'H', 'UpdatedAt', 'UpdatedBy', 'Deleted'];
function tplSheet_(ss) {
  let sh = sheet_(ss, TPL_SHEET);
  if (!sh) { sh = ss.insertSheet(TPL_SHEET); sh.getRange(1, 1, 1, TPL_HEAD.length).setValues([TPL_HEAD]); sh.setFrozenRows(1); }
  return sh;
}
const iso_ = v => v instanceof Date ? v.toISOString() : str_(v);
function zonesRead_(ss) { // {zones: {key: {name, vat, nameKey, at, fields}}, deleted: {key: at}}
  const zones = {}, deleted = {};
  values_(sheet_(ss, TPL_SHEET)).slice(1).forEach(r => {
    const key = str_(r[0]); if (!key) return;
    if (r[12] === true || str_(r[12]).toUpperCase() === 'TRUE') { deleted[key] = iso_(r[10]); return; }
    const t = zones[key] = zones[key] || { name: str_(r[1]), vat: str_(r[2]).replace(/\D/g, ''), nameKey: str_(r[3]), at: iso_(r[10]), fields: {} };
    const f = str_(r[4]); if (!f) return;
    const pg = str_(r[5]);
    t.fields[f] = { pg: pg === 'last' ? 'last' : Number(pg) || 0, x: Number(r[6]), y: Number(r[7]), w: Number(r[8]), h: Number(r[9]) };
  });
  return { zones, deleted };
}
// put: {key: قالب أو null للحذف}. يستبدل صفوف المورد فقط
function zonesWrite_(ss, put, who) {
  const sh = tplSheet_(ss), rows = values_(sh), keys = Object.keys(put || {});
  if (!keys.length) return;
  const keep = rows.slice(1).filter(r => keys.indexOf(str_(r[0])) < 0);
  const now = new Date();
  keys.forEach(k => {
    const t = put[k];
    if (!t) { keep.push([k, '', '', '', '', '', '', '', '', '', now, who, true]); return; }
    const at = t.at ? new Date(t.at) : now;
    Object.keys(t.fields || {}).forEach(f => { const z = t.fields[f];
      keep.push([k, t.name || '', str_(t.vat).replace(/\D/g, ''), t.nameKey || '', f, String(z.pg), z.x, z.y, z.w, z.h, at, who, false]); });
  });
  if (rows.length > 1) sh.getRange(2, 1, rows.length - 1, TPL_HEAD.length).clearContent();
  if (keep.length) { sh.getRange(2, 3, keep.length, 1).setNumberFormat('@'); sh.getRange(2, 6, keep.length, 1).setNumberFormat('@'); sh.getRange(2, 1, keep.length, TPL_HEAD.length).setValues(keep); }
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('قارئ الفواتير')
    .addItem('إضافة مستخدم مسموح له', 'addUser')
    .addItem('عرض المستخدمين المسموح لهم', 'listUsers')
    .addItem('إيقاف مستخدم', 'removeUser')
    .addSeparator()
    .addItem('طريقة الإعداد', 'showHelp')
    .addToUi();
}

// المستخدمون المسموح لهم بالتحديث: {email: {token, added}} محفوظة في خصائص السكربت (لا تظهر في الشيت)
function users_() { try { return JSON.parse(PropertiesService.getScriptProperties().getProperty('USERS') || '{}'); } catch (e) { return {}; } }
function saveUsers_(u) { PropertiesService.getScriptProperties().setProperty('USERS', JSON.stringify(u)); }
const normEmail_ = e => String(e || '').trim().toLowerCase();
function auth_(email, token) {
  const u = users_()[normEmail_(email)];
  return !!(u && token && u.token === String(token).trim());
}
function addUser() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('إضافة مستخدم', 'اكتب البريد الإلكتروني للشخص المسموح له بتحديث الشيت:', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const email = normEmail_(r.getResponseText());
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) { ui.alert('البريد غير صحيح'); return; }
  const u = users_(); const token = Utilities.getUuid().replace(/-/g, '').slice(0, 16);
  u[email] = { token, added: new Date().toISOString() }; saveUsers_(u);
  ui.alert('تمت الإضافة',
    'أرسل هذه البيانات إلى ' + email + ' ليضعها في التطبيق (ربط Google Sheet):\n\n' +
    'البريد: ' + email + '\nرمز الدخول: ' + token + '\nالرابط: ' + (ScriptApp.getService().getUrl() || '(رابط Web app بعد النشر)') +
    '\n\nالرمز خاص بهذا الشخص. إن أضفته مرة أخرى يتغير رمزه ويتوقف القديم.', ui.ButtonSet.OK);
}
function listUsers() {
  const u = users_(), list = Object.keys(u);
  SpreadsheetApp.getUi().alert('المستخدمون المسموح لهم (' + list.length + ')',
    list.length ? list.map(e => '• ' + e + '  —  الرمز: ' + u[e].token).join('\n') : 'لا يوجد أحد بعد. أضف المستخدمين من «إضافة مستخدم مسموح له».',
    SpreadsheetApp.getUi().ButtonSet.OK);
}
function removeUser() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('إيقاف مستخدم', 'اكتب البريد الإلكتروني الذي تريد إيقافه:', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const u = users_(), email = normEmail_(r.getResponseText());
  if (!u[email]) { ui.alert('هذا البريد غير موجود في القائمة'); return; }
  delete u[email]; saveUsers_(u); ui.alert('أُوقف ' + email + '. لن يستطيع التحديث بعد الآن.');
}
function showHelp() {
  SpreadsheetApp.getUi().alert('طريقة الإعداد (مرة واحدة)',
    '1) من محرر Apps Script: Deploy (نشر) ← New deployment (عملية نشر جديدة).\n' +
    '2) النوع: Web app. Execute as: Me. Who has access: Anyone.\n' +
    '3) اضغط Deploy ووافق على الصلاحيات، ثم انسخ رابط Web app.\n' +
    '4) من قائمة «قارئ الفواتير» ← «إضافة مستخدم مسموح له» لكل بريد (ابدأ ببريدك).\n' +
    '5) كل شخص يضع في التطبيق: الرابط + بريده + رمزه.\n' +
    'وللحماية الكاملة: شارك الشيت نفسه (زر Share) مع نفس الأشخاص فقط.',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function sheet_(ss, name) {
  return ss.getSheets().filter(s => s.getName().trim().toLowerCase() === name.toLowerCase())[0] || null;
}
function values_(sh) {
  if (!sh || sh.getLastRow() < 1 || sh.getLastColumn() < 1) return [];
  return sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
}
const str_ = v => String(v == null ? '' : v).trim().replace(/\.0+$/, '');
const invKey_ = s => str_(s).toLowerCase().replace(/[^0-9a-zء-ي]/g, '').replace(/^0+(?=.)/, '');
const dKey_ = v => v instanceof Date ? Utilities.formatDate(v, (typeof Session !== 'undefined' && Session.getScriptTimeZone()) || 'GMT', 'yyyy-MM-dd') : str_(v).slice(0, 10);
// مفتاح التكرار: المورّد + رقم الفاتورة. للمورد غير المسجل (SupID فارغ) نضيف التاريخ حتى لا تتصادم فواتير موردين مختلفين
const dupKey_ = (sup, no, date) => str_(sup) === '' ? '|' + invKey_(no) + '|' + dKey_(date) : str_(sup) + '|' + invKey_(no);

// التطبيق يطلب البيانات المرجعية: GET ?email=...&token=...&action=ref
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (!auth_(p.email, p.token)) return json_({ ok: false, error: 'هذا البريد غير مسموح له أو الرمز غير صحيح' });
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (p.action === 'hashes') return json_({ ok: true, hashes: hashes_(ss) });
  if (p.action === 'zones') { const z = zonesRead_(ss); return json_({ ok: true, title: ss.getName(), zones: z.zones, deleted: z.deleted }); }
  const out = { ok: true, title: ss.getName(), sheets: {}, hashes: hashes_(ss) };
  REF_SHEETS.forEach(n => { out.sheets[n] = values_(sheet_(ss, n)).map(r => r.map(v => v instanceof Date ? Utilities.formatDate(v, 'GMT', 'yyyy-MM-dd') : v)); });
  const inv = values_(sheet_(ss, 'InvoiceT'));
  if (inv.length) {
    const h = inv[0].map(str_), iId = h.indexOf('InvID'), iNo = h.indexOf('InvoiceNo'), iSup = h.indexOf('SupID');
    out.sheets.InvoiceT = [['InvID', 'InvoiceNo', 'SupID']].concat(inv.slice(1).filter(r => r[iNo] !== '').map(r => [r[iId], r[iNo], r[iSup]]));
  }
  const vr = values_(sheet_(ss, 'VatReport'));
  out.sheets.VatReport = vr.length ? [vr[0]] : [];
  return json_(out);
}

// التطبيق يرسل الفواتير المعتمدة: POST {email, token, invoices:[...]}
// يضيف عمودًا في نهاية الترويسة إن لم يكن موجودًا (مثل VATID)
function ensureCol_(sh, head, name) {
  if (head.some(h => h.toLowerCase() === name.toLowerCase())) { const i = head.findIndex(h => h.toLowerCase() === name.toLowerCase()); head[i] = name; return; }
  head.push(name); sh.getRange(1, head.length).setValue(name);
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'بيانات غير صالحة' }); }
  if (!auth_(body.email, body.token)) return json_({ ok: false, error: 'هذا البريد غير مسموح له أو الرمز غير صحيح' });
  const who = normEmail_(body.email);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (body.action === 'zones') { zonesWrite_(ss, body.put, who); const z = zonesRead_(ss); return json_({ ok: true, zones: z.zones, deleted: z.deleted }); }
    const invSh = sheet_(ss, 'InvoiceT'), vatSh = sheet_(ss, 'VatReport');
    if (!invSh || !vatSh) return json_({ ok: false, error: 'لم أجد ورقة InvoiceT أو VatReport' });
    const inv = values_(invSh), ih = inv[0].map(str_);
    ensureCol_(invSh, ih, 'VATID');
    const col = n => ih.indexOf(n);
    const existing = {}; let maxId = 0;
    inv.slice(1).forEach(r => {
      const id = Number(r[col('InvID')]); if (id > maxId) maxId = id;
      if (r[col('InvoiceNo')] !== '') existing[dupKey_(r[col('SupID')], r[col('InvoiceNo')], r[col('InvDate')] || r[col('Date')])] = r[col('InvID')];
    });
    const vh = values_(vatSh)[0].map(str_);
    ensureCol_(vatSh, vh, 'VATID');
    const today = new Date();
    const known = hashes_(ss), hashRows = [];
    const added = [], skipped = [], invRows = [], vatRows = [];
    (body.invoices || []).forEach(d => {
      const key = dupKey_(d.SupID, d.InvoiceNo, d.Date);
      if (!d.InvoiceNo) { skipped.push({ id: d.id, reason: 'بدون رقم فاتورة' }); return; }
      const fh = str_(d.FileHash).toLowerCase();
      if (fh && known[fh]) { skipped.push({ id: d.id, reason: 'نفس ملف الفاتورة أُرسل سابقًا (InvID ' + known[fh] + ')', invId: known[fh], byFile: true }); return; }
      if (existing[key]) { skipped.push({ id: d.id, reason: 'موجودة مسبقًا (InvID ' + existing[key] + ')', invId: existing[key] }); return; }
      const date = d.Date && /^\d{4}-\d{2}-\d{2}$/.test(d.Date) ? new Date(d.Date + 'T00:00:00') : (d.Date || '');
      const vals = {
        InvID: ++maxId, InvoiceNo: d.InvoiceNo, SupID: d.SupID === '' ? '' : d.SupID, ProjectID: d.ProjectID, InvDate: date, Date: date,
        Invamount: d.Invamount, invVat: d.invVat, Product: d.ProductID, VATID: str_(d.VATID).replace(/\D/g, ''), ProductID: d.ProductID, ProductName: d.ProductName,
        SupfName: d.SupfName, ProjectName: d.ProjectName, InvEnteryDate: today, EnteredBy: who, FilePath: d.FilePath || '', FileName: d.FilePath || '',
        InvNote: [d.Note || '', ih.indexOf('FilePath') < 0 && ih.indexOf('FileName') < 0 && d.FilePath ? 'الملف: ' + d.FilePath : ''].filter(Boolean).join(' | ')
      };
      invRows.push(ih.map(h => vals[h] !== undefined && vals[h] !== null ? vals[h] : ''));
      // VatReport: «ProductName» المكرر بعد ProjectID = اسم المشروع
      vatRows.push(vh.map((h, i) => (h === 'ProductName' && vh[i - 1] === 'ProjectID') ? (d.ProjectName || '') : (vals[h] !== undefined && vals[h] !== null ? vals[h] : '')));
      existing[key] = maxId;
      if (fh) { known[fh] = maxId; hashRows.push([fh, maxId, d.FilePath || '', who, today]); }
      added.push({ id: d.id, invId: maxId });
    });
    // رقم الفاتورة نص (حتى لا تضيع الأصفار في أوله مثل 03337)
    if (invRows.length) { const r0 = invSh.getLastRow() + 1; if (col('InvoiceNo') >= 0) invSh.getRange(r0, col('InvoiceNo') + 1, invRows.length, 1).setNumberFormat('@'); invSh.getRange(r0, 1, invRows.length, ih.length).setValues(invRows); }
    if (vatRows.length) { const r0 = vatSh.getLastRow() + 1; const vi = vh.indexOf('InvoiceNo'); if (vi >= 0) vatSh.getRange(r0, vi + 1, vatRows.length, 1).setNumberFormat('@'); vatSh.getRange(r0, 1, vatRows.length, vh.length).setValues(vatRows); }
    if (hashRows.length) { const hs = hashSheet_(ss); hs.getRange(hs.getLastRow() + 1, 1, hashRows.length, HASH_HEAD.length).setValues(hashRows); }
    return json_({ ok: true, added, skipped });
  } finally { lock.releaseLock(); }
}
