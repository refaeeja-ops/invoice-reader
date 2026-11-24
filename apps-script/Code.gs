/**
 * قارئ الفواتير ⇄ Google Sheet (مجاني: Google Apps Script)
 *
 * يُلصق هذا الملف في الشيت: Extensions (الإضافات) ← Apps Script.
 * - يستقبل الفواتير المعتمدة من التطبيق ويضيفها إلى InvoiceT و VatReport (بدون تكرار).
 * - يعطي التطبيق البيانات المرجعية (الموردين والمشاريع والمنتجات) ليقرأها مباشرة من الشيت.
 * - التحديث مسموح فقط لبريد إلكتروني أضافه المالك من قائمة «قارئ الفواتير» (لكل بريد رمز خاص)، ويُسجَّل البريد في EnteredBy.
 */

const REF_SHEETS = ['SuppliersT', 'ProjectsT', 'ProductsT', 'SupplierProductT'];

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

// التطبيق يطلب البيانات المرجعية: GET ?email=...&token=...&action=ref
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (!auth_(p.email, p.token)) return json_({ ok: false, error: 'هذا البريد غير مسموح له أو الرمز غير صحيح' });
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const out = { ok: true, title: ss.getName(), sheets: {} };
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
function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'بيانات غير صالحة' }); }
  if (!auth_(body.email, body.token)) return json_({ ok: false, error: 'هذا البريد غير مسموح له أو الرمز غير صحيح' });
  const who = normEmail_(body.email);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const invSh = sheet_(ss, 'InvoiceT'), vatSh = sheet_(ss, 'VatReport');
    if (!invSh || !vatSh) return json_({ ok: false, error: 'لم أجد ورقة InvoiceT أو VatReport' });
    const inv = values_(invSh), ih = inv[0].map(str_);
    const col = n => ih.indexOf(n);
    const existing = {}; let maxId = 0;
    inv.slice(1).forEach(r => {
      const id = Number(r[col('InvID')]); if (id > maxId) maxId = id;
      if (r[col('InvoiceNo')] !== '') existing[str_(r[col('SupID')]) + '|' + invKey_(r[col('InvoiceNo')])] = r[col('InvID')];
    });
    const vh = values_(vatSh)[0].map(str_);
    const today = new Date();
    const added = [], skipped = [], invRows = [], vatRows = [];
    (body.invoices || []).forEach(d => {
      const key = str_(d.SupID) + '|' + invKey_(d.InvoiceNo);
      if (!d.InvoiceNo) { skipped.push({ id: d.id, reason: 'بدون رقم فاتورة' }); return; }
      if (existing[key]) { skipped.push({ id: d.id, reason: 'موجودة مسبقًا (InvID ' + existing[key] + ')', invId: existing[key] }); return; }
      const date = d.Date && /^\d{4}-\d{2}-\d{2}$/.test(d.Date) ? new Date(d.Date + 'T00:00:00') : (d.Date || '');
      const vals = {
        InvID: ++maxId, InvoiceNo: d.InvoiceNo, SupID: d.SupID === '' ? '' : d.SupID, ProjectID: d.ProjectID, InvDate: date, Date: date,
        Invamount: d.Invamount, invVat: d.invVat, Product: d.ProductID, ProductID: d.ProductID, ProductName: d.ProductName,
        SupfName: d.SupfName, ProjectName: d.ProjectName, InvEnteryDate: today, EnteredBy: who, FilePath: d.FilePath || '', FileName: d.FilePath || '',
        InvNote: [d.Note || '', ih.indexOf('FilePath') < 0 && ih.indexOf('FileName') < 0 && d.FilePath ? 'الملف: ' + d.FilePath : ''].filter(Boolean).join(' | ')
      };
      invRows.push(ih.map(h => vals[h] !== undefined && vals[h] !== null ? vals[h] : ''));
      // VatReport: «ProductName» المكرر بعد ProjectID = اسم المشروع
      vatRows.push(vh.map((h, i) => (h === 'ProductName' && vh[i - 1] === 'ProjectID') ? (d.ProjectName || '') : (vals[h] !== undefined && vals[h] !== null ? vals[h] : '')));
      existing[key] = maxId;
      added.push({ id: d.id, invId: maxId });
    });
    // رقم الفاتورة نص (حتى لا تضيع الأصفار في أوله مثل 03337)
    if (invRows.length) { const r0 = invSh.getLastRow() + 1; if (col('InvoiceNo') >= 0) invSh.getRange(r0, col('InvoiceNo') + 1, invRows.length, 1).setNumberFormat('@'); invSh.getRange(r0, 1, invRows.length, ih.length).setValues(invRows); }
    if (vatRows.length) { const r0 = vatSh.getLastRow() + 1; const vi = vh.indexOf('InvoiceNo'); if (vi >= 0) vatSh.getRange(r0, vi + 1, vatRows.length, 1).setNumberFormat('@'); vatSh.getRange(r0, 1, vatRows.length, vh.length).setValues(vatRows); }
    return json_({ ok: true, added, skipped });
  } finally { lock.releaseLock(); }
}
