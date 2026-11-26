/**
 * قارئ الفواتير ⇄ Google Sheet (مجاني: Google Apps Script)
 *
 * يُلصق هذا الملف في الشيت: Extensions (الإضافات) ← Apps Script.
 * - يستقبل الفواتير المعتمدة من التطبيق ويضيفها إلى InvoiceT و VatReport (بدون تكرار).
 * - يعطي التطبيق البيانات المرجعية (الموردين والمشاريع والمنتجات) ليقرأها مباشرة من الشيت.
 * - يحفظ «قوالب الموردين» (أماكن الخانات) في ورقة VendorTemplates حتى تظهر على كل الأجهزة.
 * - يستقبل فواتير الواجهة المبسطة في ورقة «فواتير التطبيق» (صف لكل فاتورة، يُحدَّث إن أُرسلت مرة أخرى).
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

// ورقة فواتير الواجهة المبسطة (تُنشأ تلقائيًا ولا تمس InvoiceT أو VatReport)
const APP_SHEET = 'فواتير التطبيق';
const APP_HEAD = ['معرّف التطبيق', 'رقم الفاتورة', 'تاريخ الفاتورة', 'المبلغ الإجمالي', 'مبلغ الضريبة', 'اسم المورد', 'الرقم الضريبي', 'طريقة الدفع', 'المشروع', 'ملاحظات', 'الملف', 'الحالة', 'أُرسلت بواسطة', 'آخر تحديث'];
function appSheet_(ss) {
  let sh = sheet_(ss, APP_SHEET);
  if (!sh) { sh = ss.insertSheet(APP_SHEET); sh.getRange(1, 1, 1, APP_HEAD.length).setValues([APP_HEAD]).setFontWeight('bold'); sh.setFrozenRows(1); sh.setRightToLeft(true); }
  return sh;
}
// نص يبدأ بـ = أو + أو - أو @ قد يُفسَّر معادلة: نضيف قبله علامة '
const safe_ = v => { const t = String(v == null ? '' : v); return /^[=+\-@]/.test(t) && !/^-?\d+(\.\d+)?$/.test(t) ? "'" + t : t; };
const num_ = v => v === '' || v == null || isNaN(Number(v)) ? safe_(v) : Number(v);
const vendKey_ = (vat, name) => { const d = str_(vat).replace(/\D/g, ''); return d.length === 15 ? 'v' + d : 'n' + str_(name).toLowerCase().replace(/[^0-9a-z\u0600-\u06ff]/g, ''); };
// عمود «الرقم الضريبي المختصر» = أول 10 خانات من اليسار من الرقم الضريبي. يُضاف في نهاية الترويسة إن لم يوجد، ويُعبّأ لكل الصفوف
const APP_SHORT_COL = 'الرقم الضريبي المختصر';
function fillAppShort_(sh) {
  const v = values_(sh); if (!v.length) return;
  const h = v[0].map(str_); let c = h.indexOf(APP_SHORT_COL);
  if (c < 0) { c = h.length; sh.getRange(1, c + 1).setValue(APP_SHORT_COL); }
  if (v.length < 2) return;
  const col = v.slice(1).map(r => [str_(r[6]).replace(/\D/g, '').slice(0, 10)]);
  const rg = sh.getRange(2, c + 1, col.length, 1); rg.setNumberFormat('@'); rg.setValues(col);
}
function appInvoices_(ss, list, who) {
  const sh = appSheet_(ss), rows = values_(sh), now = new Date();
  const byId = {}, byKey = {};
  rows.slice(1).forEach((r, i) => { const id = str_(r[0]); if (id) byId[id] = i + 2;
    if (str_(r[1]) && !/^محذوفة/.test(str_(r[11]))) byKey[vendKey_(r[6], r[5]) + '|' + invKey_(r[1])] = id; });
  const out = [], add = [];
  (list || []).forEach(d => {
    const id = str_(d.id); if (!id) return;
    const key = vendKey_(d.VATID, d.Vendor) + '|' + invKey_(d.InvoiceNo);
    if (!d.Deleted && str_(d.InvoiceNo) && byKey[key] && byKey[key] !== id) { out.push({ id, skipped: 'نفس رقم الفاتورة ونفس المورد موجودة في الشيت' }); return; }
    const date = d.Date && /^\d{4}-\d{2}-\d{2}$/.test(d.Date) ? new Date(d.Date + 'T00:00:00') : safe_(d.Date);
    const row = [id, safe_(d.InvoiceNo), date, num_(d.Total), num_(d.Vat), safe_(d.Vendor), str_(d.VATID).replace(/\D/g, ''), safe_(d.Payment), safe_(d.Project), safe_(d.Note), safe_(d.File),
      d.Deleted ? 'محذوفة' + (d.DeleteReason ? ': ' + safe_(d.DeleteReason) : '') : 'نشطة', who, now];
    if (byId[id]) { sh.getRange(byId[id], 2, 1, 1).setNumberFormat('@'); sh.getRange(byId[id], 7, 1, 1).setNumberFormat('@'); sh.getRange(byId[id], 1, 1, row.length).setValues([row]); out.push({ id, row: byId[id], updated: true }); }
    else { add.push(row); out.push({ id, add: add.length - 1 }); }
    if (!d.Deleted && str_(d.InvoiceNo)) byKey[key] = id;
  });
  if (add.length) {
    const r0 = Math.max(sh.getLastRow(), 1) + 1;
    sh.getRange(r0, 2, add.length, 1).setNumberFormat('@'); sh.getRange(r0, 7, add.length, 1).setNumberFormat('@');
    sh.getRange(r0, 1, add.length, APP_HEAD.length).setValues(add);
    out.forEach(o => { if (o.add !== undefined) { o.row = r0 + o.add; delete o.add; } });
  }
  fillAppShort_(sh);
  return out;
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('قارئ الفواتير')
    .addItem('إضافة مستخدم مسموح له', 'addUser')
    .addItem('عرض المستخدمين المسموح لهم', 'listUsers')
    .addItem('إيقاف مستخدم', 'removeUser')
    .addSeparator()
    .addItem('نسخ «فواتير التطبيق» إلى InvoiceT', 'copyAppToInvoiceT')
    .addItem('تعبئة الرقم الضريبي (VATID) في InvoiceT', 'fillVatIds')
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
// يملأ خانات VATID الفارغة في InvoiceT من SuppliersT حسب SupID. لا يضيف صفوفًا ولا يغيّر قيمة موجودة.
function fillVatIds() {
  const ui = SpreadsheetApp.getUi(), ss = SpreadsheetApp.getActiveSpreadsheet();
  const invSh = sheet_(ss, 'InvoiceT'), supSh = sheet_(ss, 'SuppliersT');
  if (!invSh || !supSh) { ui.alert('لم أجد ورقة InvoiceT أو SuppliersT'); return; }
  const sup = values_(supSh), sh = sup[0].map(str_), si = sh.indexOf('SupID'), sv = sh.findIndex(h => h.toLowerCase() === 'vatid');
  if (si < 0 || sv < 0) { ui.alert('ورقة SuppliersT يجب أن تحتوي العمودين SupID و VATID'); return; }
  const vat = {}; sup.slice(1).forEach(r => { const id = str_(r[si]), v = str_(r[sv]).replace(/\D/g, ''); if (id && v) vat[id] = v; });
  const inv = values_(invSh); if (inv.length < 2) { ui.alert('لا توجد فواتير في InvoiceT'); return; }
  const ih = inv[0].map(str_); ensureCol_(invSh, ih, 'VATID');
  const ii = ih.indexOf('SupID'), iv = ih.indexOf('VATID');
  if (ii < 0) { ui.alert('ورقة InvoiceT لا تحتوي العمود SupID'); return; }
  const col = inv.slice(1).map(r => [iv < r.length ? r[iv] : '']);
  let filled = 0, missing = 0, kept = 0;
  inv.slice(1).forEach((r, k) => {
    if (str_(col[k][0])) { kept++; return; }
    const v = vat[str_(r[ii])]; if (v) { col[k][0] = v; filled++; } else if (str_(r[ii])) missing++;
  });
  const rg = invSh.getRange(2, iv + 1, col.length, 1); rg.setNumberFormat('@'); rg.setValues(col);
  ui.alert('تعبئة الرقم الضريبي', 'عُبّئت ' + filled + ' فاتورة.\nبقيت كما هي (فيها رقم مسبقًا): ' + kept + '.\nموردها بلا رقم ضريبي في SuppliersT: ' + missing + '.', ui.ButtonSet.OK);
}
// ينسخ فواتير ورقة «فواتير التطبيق» إلى InvoiceT. كل فاتورة تُنسخ مرة واحدة فقط:
// يُكتب رقمها في InvoiceT في عمود «InvID في InvoiceT» بورقة التطبيق. المورد يُطابق برقمه الضريبي المختصر في SuppliersT،
// وإن لم يوجد يُضاف مورد جديد برقم Max(SupID)+1.
const APP_COPY_COL = 'InvID في InvoiceT';
const hIdx_ = (head, name) => head.findIndex(h => h.toLowerCase().replace(/\s/g, '') === name.toLowerCase().replace(/\s/g, ''));
function copyAppToInvoiceT() {
  const ui = SpreadsheetApp.getUi(), lock = LockService.getScriptLock(); lock.waitLock(20000);
  try { const r = copyApp_(SpreadsheetApp.getActiveSpreadsheet()); ui.alert('نسخ فواتير التطبيق إلى InvoiceT', r.msg, ui.ButtonSet.OK); }
  catch (e) { ui.alert('تعذّر النسخ: ' + (e && e.message || e)); }
  finally { lock.releaseLock(); }
}
function copyApp_(ss) {
  const app = sheet_(ss, APP_SHEET), invSh = sheet_(ss, 'InvoiceT'), supSh = sheet_(ss, 'SuppliersT');
  if (!app) return { msg: 'لا توجد ورقة «' + APP_SHEET + '» بعد. أرسل الفواتير من التطبيق أولًا.' };
  if (!invSh || !supSh) throw new Error('لم أجد ورقة InvoiceT أو SuppliersT');
  fillAppShort_(app);
  // ورقة التطبيق: عمود لتسجيل InvID بعد النسخ
  const av = values_(app), ah = av[0].map(str_);
  const shc = ah.indexOf(APP_SHORT_COL); // عمود الرقم الضريبي المختصر: هو المستخدم للمطابقة
  let ac = ah.indexOf(APP_COPY_COL); if (ac < 0) { ac = ah.length; app.getRange(1, ac + 1).setValue(APP_COPY_COL); }
  // الموردون
  const sv = values_(supSh), sh = sv[0].map(str_);
  ['SupID', 'SupfName', 'VATID', 'Active', 'EnteredBy'].forEach(n => { if (hIdx_(sh, n) < 0) ensureCol_(supSh, sh, n); });
  const S = n => hIdx_(sh, n);
  const short_ = v => str_(v).replace(/\D/g, '').slice(0, 10); // الرقم الضريبي المختصر = أول 10 أرقام
  const supByVat = {}; let maxSup = 0;
  sv.slice(1).forEach(r => { const id = Number(r[S('SupID')]); if (id > maxSup) maxSup = id; const k = short_(r[S('VATID')]); if (k.length === 10 && !supByVat[k]) supByVat[k] = r[S('SupID')]; });
  // الفواتير
  const iv = values_(invSh), ih = iv[0].map(str_);
  ['InvID', 'InvoiceNo', 'SupID', 'InvDate', 'Invamount', 'InvVat', 'Paymethod', 'EnteredBy', 'InvNote', 'RevDate', 'VATID', 'ProjectID'].forEach(n => { if (hIdx_(ih, n) < 0) ensureCol_(invSh, ih, n); });
  const I = n => hIdx_(ih, n);
  let maxInv = 0; const existing = {};
  iv.slice(1).forEach(r => { const id = Number(r[I('InvID')]); if (id > maxInv) maxInv = id; if (str_(r[I('InvoiceNo')])) existing[str_(r[I('SupID')]) + '|' + invKey_(r[I('InvoiceNo')])] = r[I('InvID')]; });
  const newSup = [], newInv = [], marks = [], now = new Date();
  let copied = 0, dup = 0, skippedDel = 0, skippedNo = 0, already = 0;
  const supLines = [], invLines = [], skipLines = [];
  av.slice(1).forEach((r, k) => {
    const rowNo = k + 2;
    if (str_(r[ac])) { already++; return; }
    if (/^محذوفة/.test(str_(r[11]))) { skippedDel++; skipLines.push('صف ' + rowNo + ' (' + str_(r[5]) + ' ' + str_(r[1]) + '): محذوفة في التطبيق'); return; }
    const no = str_(r[1]); if (!no) { skippedNo++; skipLines.push('صف ' + rowNo + ' (' + str_(r[5]) + '): بلا رقم فاتورة'); return; }
    const vat = str_(r[6]).replace(/\D/g, ''), sk = short_(shc >= 0 ? r[shc] : vat), name = str_(r[5]), who = str_(r[12]);
    let sup = sk.length === 10 ? supByVat[sk] : undefined;
    if (sup === undefined || sup === '') {
      sup = ++maxSup;
      const row = sh.map(() => '');
      row[S('SupID')] = sup; row[S('SupfName')] = name; row[S('VATID')] = sk; row[S('Active')] = true; row[S('EnteredBy')] = 'by application >> ' + who;
      newSup.push(row); if (sk.length === 10) supByVat[sk] = sup;
      supLines.push('SupID ' + sup + ': ' + name + ' | الرقم المختصر ' + (sk || '(لا يوجد)') + ' | ' + row[S('EnteredBy')]);
    }
    const key = str_(sup) + '|' + invKey_(no);
    if (existing[key]) { dup++; marks.push([rowNo, existing[key]]); skipLines.push('صف ' + rowNo + ' (' + name + ' ' + no + '): موجودة مسبقًا في InvoiceT برقم InvID ' + existing[key]); return; }
    const id = ++maxInv, row = ih.map(() => '');
    row[I('InvID')] = id; row[I('InvoiceNo')] = no; row[I('SupID')] = sup; row[I('InvDate')] = r[2]; row[I('Invamount')] = r[3]; row[I('InvVat')] = r[4];
    row[I('Paymethod')] = r[7]; row[I('EnteredBy')] = 'application'; row[I('InvNote')] = [str_(r[10]), str_(r[0])].filter(Boolean).join(' | ');
    row[I('RevDate')] = r[13] || now; row[I('VATID')] = vat;
    const pj = str_(r[8]); row[I('ProjectID')] = /^\d+$/.test(pj) ? Number(pj) : pj;
    newInv.push(row); existing[key] = id; marks.push([rowNo, id]); copied++;
    invLines.push('InvID ' + id + ': فاتورة ' + no + ' | ' + name + ' | SupID ' + sup);
  });
  if (newSup.length) { const r0 = supSh.getLastRow() + 1; supSh.getRange(r0, S('VATID') + 1, newSup.length, 1).setNumberFormat('@'); supSh.getRange(r0, 1, newSup.length, sh.length).setValues(newSup); }
  if (newInv.length) { const r0 = invSh.getLastRow() + 1; invSh.getRange(r0, I('InvoiceNo') + 1, newInv.length, 1).setNumberFormat('@'); invSh.getRange(r0, I('VATID') + 1, newInv.length, 1).setNumberFormat('@'); invSh.getRange(r0, 1, newInv.length, ih.length).setValues(newInv); }
  marks.forEach(([rn, id]) => app.getRange(rn, ac + 1).setValue(id));
  const lim = (a, n) => a.length > n ? a.slice(0, n).concat(['... و' + (a.length - n) + ' أخرى (التفاصيل كاملة في ورقة «سجل النسخ»)']) : a;
  const msg = 'نُسخت ' + copied + ' فاتورة إلى InvoiceT' + (copied ? ' (InvID من ' + (maxInv - copied + 1) + ' إلى ' + maxInv + ')' : '') + '.' +
    '\n\nموردون جدد أُضيفوا إلى SuppliersT: ' + newSup.length + (supLines.length ? '\n' + lim(supLines, 15).join('\n') : ' (كل الموردين موجودون مسبقًا)') +
    (invLines.length ? '\n\nالفواتير المنسوخة:\n' + lim(invLines, 15).join('\n') : '') +
    (already ? '\n\nنُسخت في مرة سابقة فلم تُكرر: ' + already : '') +
    (skipLines.length ? '\n\nلم تُنسخ:\n' + lim(skipLines, 15).join('\n') : '');
  log_(ss, [supLines.map(l => ['مورد جديد', l]), invLines.map(l => ['فاتورة منسوخة', l]), skipLines.map(l => ['لم تُنسخ', l])].flat(), copied, newSup.length);
  return { copied, dup, newSup: newSup.length, msg };
}
// سجل دائم لكل عملية نسخ في ورقة «سجل النسخ»
function log_(ss, lines, copied, nSup) {
  let sh = sheet_(ss, 'سجل النسخ');
  if (!sh) { sh = ss.insertSheet('سجل النسخ'); sh.getRange(1, 1, 1, 3).setValues([['الوقت', 'النوع', 'التفاصيل']]); if (sh.setRightToLeft) sh.setRightToLeft(true); }
  const now = new Date(), rows = [[now, 'تشغيل', 'نُسخت ' + copied + ' فاتورة، وأُضيف ' + nSup + ' مورد']].concat(lines.map(l => [now, l[0], l[1]]));
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, 3).setValues(rows);
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
    if (body.action === 'appInvoices') return json_({ ok: true, results: appInvoices_(ss, body.invoices, who) });
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
