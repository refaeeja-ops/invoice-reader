/* قارئ الفواتير (النسخة المبسطة): قراءة الفواتير وتنزيل ملف Excel بالبيانات الأساسية. كل المعالجة داخل المتصفح ومجانية. */
(function () {
  'use strict';
  const P = window.InvoiceParse, X = window.InvoiceExport;
  const $ = id => document.getElementById(id);
  const esc = s => String(s ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const abs = p => new URL(p, location.href).href;
  const pathOf = f => f.webkitRelativePath || f.name || '';
  const setProg = p => progress(p);

  if (window.pdfjsLib) pdfjsLib.GlobalWorkerOptions.workerSrc = abs('vendor/pdf.worker.min.js');


  const toast = m => { const t = $('toast'); t.textContent = m; t.classList.add('on'); clearTimeout(toast.h); toast.h = setTimeout(() => t.classList.remove('on'), 1800); };
  function status(m, err) { $('status').textContent = m || ''; $('status').classList.toggle('err', !!err); }
  function progress(p) { $('prog').classList.toggle('on', p != null); $('barI').style.width = Math.round((p || 0) * 100) + '%'; }

  // ---------- تحويل الملف إلى صور ----------
  function loadImg(src) { return new Promise((ok, no) => { const i = new Image(); i.onload = () => ok(i); i.onerror = () => no(new Error('img')); i.src = src; }); }
  function toCanvas(img, maxSide) {
    const w = img.naturalWidth || img.width, h = img.naturalHeight || img.height;
    const r = Math.min(maxSide / Math.max(w, h), 1);
    const c = document.createElement('canvas'); c.width = Math.round(w * r); c.height = Math.round(h * r);
    const x = c.getContext('2d'); x.fillStyle = '#fff'; x.fillRect(0, 0, c.width, c.height);
    x.imageSmoothingQuality = 'high'; x.drawImage(img, 0, 0, c.width, c.height); return c;
  }
  // صفحة PDF واحدة: صورة + النص الرقمي + أماكن الكلمات (نسبة من الصفحة) لقراءة منطقة محددة بدقة
  async function renderPdfPage(doc, p) {
    const pg = await doc.getPage(p), v0 = pg.getViewport({ scale: 1 });
    const v = pg.getViewport({ scale: Math.min(3, 2200 / Math.max(v0.width, v0.height)) });
    const c = document.createElement('canvas'); c.width = Math.round(v.width); c.height = Math.round(v.height);
    const x = c.getContext('2d'); x.fillStyle = '#fff'; x.fillRect(0, 0, c.width, c.height);
    await pg.render({ canvasContext: x, viewport: v }).promise;
    let text = '', items = [];
    try {
      const tc = await pg.getTextContent(), rows = [];
      tc.items.forEach(it => {
        if (!it.str || !it.str.trim()) return;
        const y = it.transform[5], xx = it.transform[4];
        let row = rows.find(r => Math.abs(r.y - y) < Math.max(3, Math.abs(it.transform[3]) * 0.5));
        if (!row) rows.push(row = { y, parts: [] });
        row.parts.push({ x: xx, s: it.str });
        const [px, py] = v.convertToViewportPoint(xx, y), h = Math.hypot(it.transform[2], it.transform[3]) * v.scale, w = (it.width || 0) * v.scale;
        items.push({ s: it.str, x: px / c.width, y: (py - h) / c.height, w: w / c.width, h: h / c.height });
      });
      rows.sort((a, b) => b.y - a.y);
      text = rows.map(r => {
        const rtl = /[؀-ۿ]/.test(r.parts.map(p => p.s).join(''));
        r.parts.sort((a, b) => rtl ? b.x - a.x : a.x - b.x);
        return r.parts.map(p => p.s.trim()).join('  ');
      }).join('\n');
    } catch (e) { items = []; }
    return { canvas: c, text, items };
  }
  async function pdfPages(file) {
    if (!window.pdfjsLib) throw new Error('تعذّر تحميل قارئ PDF.');
    const doc = await pdfjsLib.getDocument({ data: await file.arrayBuffer() }).promise;
    const n = Math.min(doc.numPages, 5), out = [];
    for (let p = 1; p <= n; p++) out.push(await renderPdfPage(doc, p));
    return { pages: out, total: doc.numPages };
  }

  // ---------- رمز QR ----------
  function scanQR(canvas) {
    if (!window.jsQR) return null;
    const tryRegion = (sx, sy, sw, sh, target) => {
      const r = Math.min(1.5, target / Math.max(sw, sh));
      const c = document.createElement('canvas'); c.width = Math.max(1, Math.round(sw * r)); c.height = Math.max(1, Math.round(sh * r));
      const x = c.getContext('2d', { willReadFrequently: true }); x.drawImage(canvas, sx, sy, sw, sh, 0, 0, c.width, c.height);
      const d = x.getImageData(0, 0, c.width, c.height);
      const res = jsQR(d.data, c.width, c.height, { inversionAttempts: 'attemptBoth' });
      return res && res.data ? res.data : null;
    };
    const W = canvas.width, H = canvas.height;
    const regions = [[0, 0, W, H, 1000], [0, 0, W, H, 1800], [0, H / 2, W, H / 2, 1400], [0, 0, W, H / 2, 1400],
      [0, H / 2, W / 2, H / 2, 1000], [W / 2, H / 2, W / 2, H / 2, 1000], [0, 0, W / 2, H / 2, 1000], [W / 2, 0, W / 2, H / 2, 1000],
      [0, H / 3, W, H / 3, 1400], [0, 2 * H / 3, W, H / 3, 1400]];
    let firstText = null;
    for (const [sx, sy, sw, sh, t] of regions) {
      const s = tryRegion(sx, sy, sw, sh, t);
      if (s) { const z = P.parseZatcaQR(s); if (z) return { zatca: z, raw: s }; if (!firstText) firstText = s; }
    }
    return firstText ? { zatca: null, raw: firstText } : null;
  }

  // ---------- قراءة النص (Tesseract) ----------
  // نموذجان: «ara+eng» للنص العام، و«Arabic+eng» (نموذج الخط العربي) يُستدعى عند وجود أرقام هندية لأنه أدق فيها.
  const workers = {};
  let deep = false; // إعادة القراءة: كل المحاولات بدل التوقف عند أول نتيجة مقبولة
  let ocrBase = 0, ocrSpan = 1;
  async function getWorker(lang) {
    if (workers[lang]) return workers[lang];
    status(lang === 'ara+eng' ? 'جارٍ تجهيز قارئ النص العربي والإنجليزي…' : 'جارٍ تجهيز قارئ الأرقام الهندية… (مرة واحدة فقط)');
    workers[lang] = Tesseract.createWorker(lang, 1, {
      workerPath: abs('vendor/worker.min.js'), corePath: abs('vendor/core'), langPath: abs('vendor/lang'), gzip: true,
      logger: m => { if (m.status === 'recognizing text') setProg(ocrBase + m.progress * ocrSpan); }
    }).then(async w => { await w.setParameters({ preserve_interword_spaces: '1' }); return w; });
    try { return await workers[lang]; } catch (e) { delete workers[lang]; throw e; }
  }
  function prep(canvas, size) {
    // تغيير الحجم بحيث يكون الضلع الأطول = size، ثم تحويل إلى رمادي
    const long = Math.max(canvas.width, canvas.height);
    const r = size ? Math.min(3, size / long) : (long < 1400 ? 1400 / long : 1);
    const c = document.createElement('canvas'); c.width = Math.round(canvas.width * r); c.height = Math.round(canvas.height * r);
    const x = c.getContext('2d', { willReadFrequently: true }); x.imageSmoothingQuality = 'high'; x.drawImage(canvas, 0, 0, c.width, c.height);
    const d = x.getImageData(0, 0, c.width, c.height), a = d.data;
    for (let i = 0; i < a.length; i += 4) { const g = 0.299 * a[i] + 0.587 * a[i + 1] + 0.114 * a[i + 2]; a[i] = a[i + 1] = a[i + 2] = g; }
    x.putImageData(d, 0, 0); return c;
  }
  // محاولات قراءة بإعدادات مختلفة؛ نتوقف عند أول نتيجة متسقة
  const PASSES = [{ size: 0, psm: '4' }, { size: 2400, psm: '6' }, { size: 1800, psm: '4' }, { size: 2400, psm: '4' }];
  const EAST_PASSES = [{ size: 2000, psm: '4' }, { size: 2800, psm: '4' }];
  const eastern = t => (t.match(/[٠-٩۰-۹]/g) || []).length >= 3;
  const quality = (r, q) => P.score(r) + (P.itemsOk(r) ? 20 : 0) - 3 * P.checks(r, q).length;
  const good = r => P.itemsOk(r) && r.invoice_number && r.total != null && r.vendor_tax_number && !P.checkVat(r.vendor_tax_number);

  async function ocrPage(canvas, q, say) {
    const runs = [];
    const add = (t, tag) => { const r = P.parseInvoice(t, q), s = quality(r, q); runs.push({ t, r, s }); return r; };
    const run = async (lang, pass) => {
      const w = await getWorker(lang);
      await w.setParameters({ tessedit_pageseg_mode: pass.psm });
      const img = prep(canvas, pass.size);
      const { data } = await w.recognize(img, {}, { text: true, blocks: true });
      let r = add(data.text, lang + ' ' + pass.size + '/' + pass.psm);
      // أرقام هندية: نعيد قراءة الكلمات الرقمية بالمصنّف الخاص بها
      if (eastern(data.text) && window.EasternDigits) {
        try {
          await EasternDigits.load(abs('vendor/digits-model.json'));
          const rb = EasternDigits.rebuild(data, img);
          if (rb) { const r2 = add(rb.text, 'digits ' + lang + ' ' + pass.size); if (quality(r2, q) >= quality(r, q)) r = r2; }
        } catch (e) { console.warn(e); }
      }
      return r;
    };
    for (let k = 0; k < PASSES.length; k++) {
      if (k) say(`محاولة قراءة ${k + 1} لتحسين النتيجة…`);
      ocrBase = 0.1 + 0.8 * k / (PASSES.length + EAST_PASSES.length); ocrSpan = 0.8 / (PASSES.length + EAST_PASSES.length);
      const r = await run('ara+eng', PASSES[k]);
      if (!deep && (good(r) || (k >= 1 && P.itemsOk(r)))) break;
      if (!deep && eastern(runs[0].t) && k >= 1) break; // للأرقام الهندية ننتقل للنموذج المتخصص
    }
    if (runs.some(x => eastern(x.t)) && (deep || !runs.some(x => good(x.r)))) {
      for (let k = 0; k < EAST_PASSES.length; k++) {
        say('وُجدت أرقام هندية، جارٍ قراءتها بنموذج متخصص…');
        ocrBase = 0.1 + 0.8 * (PASSES.length + k) / (PASSES.length + EAST_PASSES.length);
        const r = await run('Arabic+eng', EAST_PASSES[k]);
        if (!deep && good(r)) break;
      }
    }
    runs.sort((a, b) => b.s - a.s);
    return { best: runs[0].t, extra: runs.slice(1).map(x => x.t) };
  }

  // ---------- معالجة فاتورة واحدة ----------
  async function processFile(file, say) {
    const isPdf = file.type === 'application/pdf' || /\.pdf$/i.test(file.name || '');
    let list;
    if (isPdf) { const r = await pdfPages(file); list = r.pages; if (r.total > 5) toast('ستُقرأ أول 5 صفحات فقط'); }
    else {
      const url = URL.createObjectURL(file);
      try { list = [{ canvas: toCanvas(await loadImg(url), 2400), text: '' }]; } finally { URL.revokeObjectURL(url); }
    }
    say('جارٍ البحث عن رمز QR الضريبي…'); setProg(0.06);
    let q = null, qRaw = null;
    for (const p of list) { const s = scanQR(p.canvas); if (s) { qRaw = s.raw; if (s.zatca) { q = s.zatca; break; } } }
    const texts = [], extras = [];
    for (let i = 0; i < list.length; i++) {
      const p = list[i];
      if ((p.text || '').replace(/\s/g, '').length >= 40) { texts.push(p.text); continue; }
      say(list.length > 1 ? `جارٍ قراءة الصفحة ${i + 1} من ${list.length}…` : 'جارٍ قراءة نص الفاتورة…');
      const o = await ocrPage(p.canvas, q, say);
      texts.push(o.best); if (list.length === 1) extras.push(...o.extra);
    }
    const raw = texts.join('\n\n');
    const d = P.parseInvoice(raw, q);
    extras.forEach(t => { // ملء الحقول الناقصة من المحاولات الأخرى
      const o = P.parseInvoice(t, q);
      Object.keys(o).forEach(k => { if (k !== 'items' && k !== 'src' && k[0] !== '_' && (d[k] === '' || d[k] == null) && o[k] !== '' && o[k] != null) { d[k] = o[k]; d.src[k] = o.src[k]; } });
    });
    // اسم المورّد: نختار أنظف مرشّح من كل المحاولات (أكثر حروفًا وأقل رموزًا)
    if (d.src.vendor_name !== 'qr') {
      const clean = v => { v = v || ''; return (v.match(/[A-Za-z\u0621-\u064A]/g) || []).length - 3 * (v.match(/[^A-Za-z\u0621-\u064A\u064B-\u0652\s&'.-]/g) || []).length - 2 * (v.match(/(^|\s)\S(?=\s|$)/g) || []).length; };
      const cands = [d.vendor_name, ...extras.map(t => P.parseInvoice(t, q).vendor_name)].filter(Boolean);
      const v = cands.sort((a, b) => clean(b) - clean(a))[0]; if (v) d.vendor_name = v;
    }
    d.qr_raw = qRaw || ''; d._qr = q; d._raw = raw.slice(0, 20000);
    d.file = pathOf(file) || 'صورة من الكاميرا'; d.file_name = file.name || '';
    return { d, pages: list.map(p => toUrl(p.canvas)), src: list };
  }
  function toUrl(c) { // صورة للعرض فقط، بحجم أصغر لتوفير الذاكرة
    const r = Math.min(1, 1600 / Math.max(c.width, c.height));
    if (r === 1) return c.toDataURL('image/jpeg', 0.82);
    const s = document.createElement('canvas'); s.width = Math.round(c.width * r); s.height = Math.round(c.height * r);
    s.getContext('2d').drawImage(c, 0, 0, s.width, s.height); return s.toDataURL('image/jpeg', 0.82);
  }


  // ---------- الحفظ: البيانات في localStorage، والصور والملف الأصلي في IndexedDB ----------
  const KEY = 'invoice-simple-v1';
  let list = [];
  try { list = JSON.parse(localStorage.getItem(KEY)) || []; } catch { list = []; }
  const TK = 'invoice-simple-trash-v1';
  let trash; try { trash = JSON.parse(localStorage.getItem(TK)) || []; } catch { trash = []; }
  const saveTrash = () => { try { localStorage.setItem(TK, JSON.stringify(trash)); } catch { toast('مساحة المتصفح ممتلئة'); } };
  const trOpen = {}; let trAsk = null, delAsk = null;
  const save = () => { try { localStorage.setItem(KEY, JSON.stringify(list)); } catch { toast('مساحة المتصفح ممتلئة: نزّل ملف Excel ثم احذف الفواتير القديمة'); } };
  const DB = (name, store) => {
    let p = null;
    const db = () => p || (p = new Promise((ok, no) => { try { const r = indexedDB.open(name, 1); r.onupgradeneeded = () => r.result.createObjectStore(store); r.onsuccess = () => ok(r.result); r.onerror = () => no(r.error); } catch (e) { no(e); } }));
    const tx = async (mode, fn) => { const d = await db(); return new Promise((ok, no) => { const t = d.transaction(store, mode), r = fn(t.objectStore(store)); t.oncomplete = () => ok(r && r.result); t.onerror = () => no(t.error); }); };
    return { put: (k, v) => tx('readwrite', s => s.put(v, k)).catch(() => { }), put2: (k, v) => tx('readwrite', s => s.put(v, k)).then(() => true, e => { console.warn(e); return false; }), get: k => tx('readonly', s => s.get(k)).catch(() => null), del: k => tx('readwrite', s => s.delete(k)).catch(() => { }), clear: () => tx('readwrite', s => s.clear()).catch(() => { }) };
  };
  const IMG0 = DB('invoice-simple-files', 'f'), OLD = DB('invoice-reader-img', 'img');
  // سفاري (آيباد/آيفون) قد يرفض حفظ كائن «ملف» في IndexedDB فيضيع السجل كله بصمت،
  // لذلك نحفظ الملف الأصلي كبيانات خام (ArrayBuffer)، وإن فشل نحفظ صور الصفحات وحدها وننبّه المستخدم
  let storeWarned = false;
  const storeWarn = m => { if (storeWarned) return; storeWarned = true; toast(m); };
  const IMG = {
    get: async k => { const r = await IMG0.get(k); if (r && r.buf && !r.blob) { try { r.blob = new File([r.buf], r.name || 'invoice', { type: r.type || '' }); } catch { r.blob = new Blob([r.buf], { type: r.type || '' }); } } return r; },
    put: async (k, rec) => {
      const v = { pages: (rec && rec.pages) || [] };
      if (rec && rec.blob) { try { v.buf = await rec.blob.arrayBuffer(); v.type = rec.blob.type || ''; v.name = rec.blob.name || ''; } catch (e) { console.warn(e); } }
      if (await IMG0.put2(k, v)) return true;
      if (v.buf) { delete v.buf; if (await IMG0.put2(k, v)) { storeWarn('حُفظت صور الفواتير بدون الملف الأصلي: مساحة المتصفح لا تكفي'); return true; } }
      storeWarn('تعذّر حفظ صورة الفاتورة في هذا المتصفح. يمكنك إرفاق الملف لاحقًا من «مراجعة الفاتورة»'); return false;
    },
    del: k => IMG0.del(k), clear: () => IMG0.clear()
  };

  // أول تشغيل: نأخذ الفواتير الموجودة من النسخة الكاملة حتى لا تضيع
  let firstRun = false; try { firstRun = localStorage.getItem(KEY) === null; } catch { }
  if (firstRun) {
    try {
      const old = JSON.parse(localStorage.getItem('invoice-reader-history-v1')) || [];
      list = old.map(x => ({ id: x._id, at: x._at || '', file: x.file || '', invoice_number: x.invoice_number || '', invoice_date: x.invoice_date || '', total: x.total ?? '', tax_amount: x.tax_amount ?? '', vendor_name: x.vendor_name || '', vendor_tax_number: x.vendor_tax_number || '', payment_method: x.payment_method || '', note: x.note || '', old: true }));
    } catch { list = []; }
    save();
  }

  // ---------- طريقة الدفع: قائمة ثابتة، وأي قيمة مقروءة تُحوَّل لأقرب خيار ----------
  const PAYS = ['نقدي', 'بطاقة- مدى', 'Apple Pay', 'تحويل بنكي', 'آجل', 'حساب دائن'];
  const PAYMAP = [[/حساب\s*دائن|دائن/i, 'حساب دائن'], [/آجل|اجل|on\s*account|credit(?!\s*card)/i, 'آجل'], [/apple/i, 'Apple Pay'],
    [/تحويل|transfer|bank|بنكي|حوالة/i, 'تحويل بنكي'], [/مدى|mada|بطاق|card|visa|master|ائتمان|فيزا|شبكة/i, 'بطاقة- مدى'], [/نقد|cash|كاش|كاشير/i, 'نقدي']];
  const payNorm = v => { const t = String(v ?? '').trim(); if (!t || PAYS.includes(t)) return t; const m = PAYMAP.find(([re]) => re.test(t)); return m ? m[1] : t; };
  list.forEach(x => { if (x && x.payment_method) x.payment_method = payNorm(x.payment_method); });

  // ---------- الخانات ----------
  const F = [
    ['invoice_number', 'رقم الفاتورة', 'text'], ['invoice_date', 'تاريخ الفاتورة', 'date'],
    ['total', 'المبلغ الإجمالي', 'num'], ['tax_amount', 'مبلغ الضريبة', 'num'],
    ['vendor_name', 'اسم المورد أو التاجر', 'text'], ['vendor_tax_number', 'الرقم الضريبي للمورد', 'vat'],
    ['payment_method', 'طريقة الدفع', 'pay']];
  const blank = v => v === '' || v == null;
  const n2 = v => { if (blank(v)) return ''; const s = P.normDigits(String(v)).replace(/[٬,\s]/g, '').replace('٫', '.'); return s !== '' && isFinite(+s) ? +s : String(v); };
  const money = v => blank(v) ? '' : (isFinite(+v) ? (+v).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) : String(v));

  // ملاحظات البرنامج عن الفاتورة (تتحدث مع كل تعديل)
  function issues(x) {
    const w = [];
    const miss = F.filter(([k]) => k !== 'payment_method' && blank(x[k])).map(f => f[1]);
    if (miss.length) w.push('لم تُقرأ: ' + miss.join('، '));
    const vm = P.checkVat(x.vendor_tax_number); if (vm) w.push(vm);
    const T = +x.total, V = +x.tax_amount;
    if (!blank(x.total) && !blank(x.tax_amount) && isFinite(T) && isFinite(V)) {
      if (V >= T) w.push('مبلغ الضريبة أكبر من الإجمالي أو يساويه.');
      else if (T - V > 0 && Math.abs(V - (T - V) * 0.15) > Math.max(0.05, T * 0.002)) w.push(`الضريبة لا تساوي 15% (المتوقع تقريبًا ${money(Math.round(T / 1.15 * 0.15 * 100) / 100)}).`);
    }
    if (x.invoice_date && !/^\d{4}-\d{2}-\d{2}$/.test(x.invoice_date)) w.push('تاريخ الفاتورة غير واضح.');
    const twin = list.find(y => y !== x && y.invoice_number && y.invoice_number === x.invoice_number && (y.vendor_tax_number || y.vendor_name) === (x.vendor_tax_number || x.vendor_name));
    if (twin) w.push('مكررة: نفس رقم الفاتورة ونفس المورد موجودة في القائمة.');
    return w;
  }
  const noteOf = x => [x.note, ...issues(x)].filter(Boolean).join(' | ');

  // ---------- العرض ----------
  const PICK_ICON = '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true"><path d="M4 8V4h4M16 4h4v4M20 16v4h-4M8 20H4v-4"/><circle cx="12" cy="12" r="2.2" fill="currentColor"/></svg>';
  const fieldHtml = (x, [k, label, type]) => {
    const v = x[k] ?? '', miss = blank(v) && k !== 'payment_method';
    const attrs = `data-id="${esc(x.id)}" data-k="${k}" id="f-${esc(x.id)}-${k}"`;
    let input;
    if (type === 'date') input = `<input type="date" ${attrs} value="${esc(/^\d{4}-\d{2}-\d{2}$/.test(v) ? v : '')}">`;
    else if (type === 'pay') input = `<select ${attrs}><option value="">اختر طريقة الدفع</option>${(blank(v) || PAYS.includes(v) ? PAYS : [v, ...PAYS]).map(o => `<option${o === v ? ' selected' : ''}>${esc(o)}</option>`).join('')}</select>`;
    else input = `<input ${attrs} value="${esc(type === 'num' && !blank(v) && isFinite(+v) ? (+v).toFixed(2) : v)}"${type !== 'text' ? ' inputmode="decimal" class="num"' : ''} placeholder="${miss ? 'لم يُقرأ — أدخله' : ''}">`;
    const zs = x.zsrc && x.zsrc[k];
    const badge = zs ? `<em class="zb${zs === 'pick' ? ' m' : zs === 'new' ? ' g' : ''}">${zs === 'pick' ? 'محددة يدويًا' : zs === 'new' ? 'حُدّث بالقالب' : 'من قالب المورد'}</em>` : '';
    return `<div class="f f-${k}${miss ? ' miss' : ''}${k === 'payment_method' && blank(v) ? ' nopay' : ''}${rv.cur[x.id] === k && isOpen(x.id) ? ' cur' : ''}"><span class="fh"><label for="f-${esc(x.id)}-${k}">${label}</label>${badge}<button type="button" class="pkb${miss ? ' need' : ''}" data-pick="${esc(x.id)}" data-pk="${k}" aria-label="حدد مكان ${label} على الفاتورة" title="حدد مكان ${label} على الفاتورة">${PICK_ICON}</button></span>${input}</div>`;
  };
  function card(x, tr) {
    const w = issues(x);
    if (tr) return `<article class="inv trash" id="c-${esc(x.id)}" aria-label="فاتورة محذوفة">
      <header><div><b>${esc(x.vendor_name || 'مورد غير معروف')}</b><small class="fp">📄 ${esc(x.file || '')}</small></div><span class="amt">${esc(money(x.total))}</span></header>
      <fieldset disabled><div class="grid">${F.map(f => fieldHtml(x, f)).join('')}${projHtml(x)}
        <label class="f wide"><span>ملاحظات</span><textarea rows="2">${esc(x.note || '')}</textarea></label></div></fieldset>
      ${w.length ? `<ul class="warn">${w.map(t => `<li>${esc(t)}</li>`).join('')}</ul>` : ''}
      ${rvHtml(x)}</article>`;
    return `<article class="inv" id="c-${esc(x.id)}">
      <header><div><b>${esc(x.vendor_name || 'مورد غير معروف')}</b><small class="fp">📄 ${esc(x.file || '')}</small>${sentBadge(x)}</div><span class="amt">${esc(money(x.total))}</span></header>
      <div class="grid">${F.map(f => fieldHtml(x, f)).join('')}${projHtml(x)}
        <label class="f wide"><span>ملاحظات</span><textarea data-id="${esc(x.id)}" data-k="note" rows="2" placeholder="اكتب أي إشكالية في الفاتورة">${esc(x.note || '')}</textarea></label></div>
      ${w.length ? `<ul class="warn">${w.map(t => `<li>${esc(t)}</li>`).join('')}</ul>` : '<p class="okline">✓ البيانات الأساسية مكتملة</p>'}
      ${rrHtml(x)}${rvHtml(x)}
      <footer><button data-view="${esc(x.id)}">فتح الملف للمراجعة</button><button data-reread="${esc(x.id)}"${live[x.id] ? ' disabled' : ''}>${live[x.id] ? (live[x.id].status === 'working' ? 'جارٍ إعادة القراءة…' : 'بانتظار إعادة القراءة…') : '🔄 إعادة قراءة الفاتورة'}</button><button class="del" data-del="${esc(x.id)}"${delAsk === x.id ? ' hidden' : ''}>حذف الفاتورة</button></footer>
      ${delAsk === x.id ? `<div class="delask" role="alert"><label for="dr-${esc(x.id)}"><b>سبب حذف هذه الفاتورة (إلزامي)</b></label>
        <textarea id="dr-${esc(x.id)}" data-delwhy="${esc(x.id)}" rows="2" placeholder="مثال: فاتورة مكررة، أو ليست لنا"></textarea>
        <small>تنتقل الفاتورة إلى «الفواتير المحذوفة» أسفل الصفحة مع السبب، ويمكن استرجاعها.</small>
        <div class="zbtns"><button type="button" class="del" data-delok="${esc(x.id)}" disabled>احذف الفاتورة</button><button type="button" data-delno="">تراجع</button></div></div>` : ''}
    </article>`;
  }
  function draw() {
    const dw = delAsk && $('dr-' + delAsk), keepWhy = dw ? dw.value : '';
    const pend = Object.values(live).filter(v => !v.reread);
    $('count').textContent = list.length ? `(${list.length})` : '';
    $('xl').disabled = !list.length; $('xl').textContent = `تنزيل ملف Excel${list.length ? ' (' + list.length + ')' : ''}`;
    const nIss = list.filter(x => issues(x).length).length;
    $('sum').textContent = list.length ? (nIss ? `${nIss} فاتورة فيها ملاحظات أو بيانات ناقصة` : 'كل الفواتير مكتملة') : '';
    $('list').innerHTML = pend.map(v => `<article class="inv pend"><b>📄 ${esc(v.file)}</b> <span>${v.status === 'working' ? 'جارٍ القراءة…' : 'بالانتظار'}</span></article>`).join('')
      + (list.map(x => card(x)).join('') || (pend.length ? '' : '<p class="empty">لا توجد فواتير بعد. ارفع فاتورة أو صوّرها للبدء.</p>'));
    drawZones(); drawTrash(); drawWipe(); rvAllLabel(); sendLabel(); mountViewers();
    if (keepWhy) { const t = $('dr-' + delAsk), ok = document.querySelector(`[data-delok="${CSS.escape(delAsk)}"]`); if (t) t.value = keepWhy; if (ok) ok.disabled = !keepWhy.trim(); }
  }
  function drawTrash() {
    $('trSec').hidden = !trash.length; $('trCount').textContent = trash.length ? `(${trash.length})` : '';
    const d = s => { try { return new Date(s).toLocaleString('ar-SA-u-nu-latn', { dateStyle: 'medium', timeStyle: 'short' }); } catch { return ''; } };
    $('trList').innerHTML = trash.map(x => { const id = esc(x.id), o = !!trOpen[x.id];
      return `<li><div class="trrow"><span class="trn"><b>📄 ${esc(x.file || x.file_name || 'بدون اسم')}</b><small>${esc(x.vendor_name || 'مورد غير معروف')}${x.total !== '' && x.total != null ? ' · ' + esc(money(x.total)) : ''}${x.deletedAt ? ' · حُذفت ' + esc(d(x.deletedAt)) : ''}</small>${x.delReason ? `<span class="trwhy">سبب الحذف: ${esc(x.delReason)}</span>` : ''}</span>
        <button type="button" data-trview="${id}" aria-expanded="${o}">${o ? 'إغلاق' : 'مراجعة'}</button><button type="button" data-trback="${id}">↩ استرجاع</button><button type="button" class="del" data-trdel="${id}">حذف نهائي</button>
        ${trAsk === x.id ? `<div class="trask" role="alert"><b>حذف هذه الفاتورة نهائيًا؟</b>ستُحذف بياناتها وصورتها من المتصفح ولا تظهر هنا بعد ذلك.<div class="zbtns"><button type="button" class="del" data-trok="${id}">نعم، احذف نهائيًا</button><button type="button" data-trno="">تراجع</button></div></div>` : ''}</div>
        ${o ? card(x, true) : ''}</li>`; }).join('');
  }
  // تحديث ملاحظات وحالة بطاقة واحدة بدون إعادة رسم (حتى لا يضيع مكان الكتابة)
  function refreshCard(x) {
    const c = $('c-' + x.id); if (!c) return;
    const w = issues(x), box = c.querySelector('.warn, .okline');
    const html = w.length ? `<ul class="warn">${w.map(t => `<li>${esc(t)}</li>`).join('')}</ul>` : '<p class="okline">✓ البيانات الأساسية مكتملة</p>';
    box.outerHTML = html;
    c.querySelector('header b').textContent = x.vendor_name || 'مورد غير معروف';
    c.querySelector('.amt').textContent = money(x.total);
    const hd = c.querySelector('header > div'); if (hd) { const ob = hd.querySelector('.sb'); if (ob) ob.remove(); hd.insertAdjacentHTML('beforeend', sentBadge(x)); } sendLabel();
    F.forEach(([k]) => { const el = $(`f-${x.id}-${k}`); if (!el) return; const f = el.closest('.f'), m = blank(x[k]) && k !== 'payment_method';
      f.classList.toggle('miss', m); if (k === 'payment_method') f.classList.toggle('nopay', blank(x[k])); const pb = f.querySelector('.pkb'); if (pb) pb.classList.toggle('need', m);
      if (!(x.zsrc && x.zsrc[k])) { const zb = f.querySelector('.zb'); if (zb) zb.remove(); } });
    const nIss = list.filter(y => issues(y).length).length;
    $('sum').textContent = nIss ? `${nIss} فاتورة فيها ملاحظات أو بيانات ناقصة` : 'كل الفواتير مكتملة';
  }
  document.addEventListener('input', e => {
    if (e.target.dataset.delwhy) { const ok = document.querySelector(`[data-delok="${CSS.escape(e.target.dataset.delwhy)}"]`); if (ok) ok.disabled = !e.target.value.trim(); return; }
    const el = e.target, id = el.dataset.id, k = el.dataset.k; if (!id || !k) return;
    const x = list.find(y => y.id === id); if (!x) return;
    let v = el.value.trim();
    if (k === 'total' || k === 'tax_amount') v = n2(v);
    if (k === 'vendor_tax_number') v = P.normDigits(v).replace(/\s/g, '');
    x[k] = v; (x.edited = x.edited || {})[k] = 1; if (x.zsrc) delete x.zsrc[k]; save(); refreshCard(x);
  });
  document.addEventListener('click', async e => {
    const b = e.target.closest('button'); if (!b) return;
    if (b.dataset.del) {
      const x = list.find(y => y.id === b.dataset.del); if (!x) return;
      delAsk = x.id; draw(); const t = $('dr-' + x.id); if (t) { t.focus({ preventScroll: true }); t.scrollIntoView({ block: 'center', behavior: RM ? 'auto' : 'smooth' }); }
    }
    if (b.dataset.delno !== undefined) { const id = delAsk; delAsk = null; draw(); const d = document.querySelector(`[data-del="${CSS.escape(id || '')}"]`); if (d) d.focus({ preventScroll: true }); }
    if (b.dataset.delok) {
      const x = list.find(y => y.id === b.dataset.delok); if (!x) return;
      const why = ($('dr-' + x.id) || {}).value; if (!why || !why.trim()) { const t = $('dr-' + x.id); if (t) t.focus(); return; }
      x.delReason = why.trim(); delAsk = null;
      const i = list.indexOf(x); list.splice(i, 1); x.deletedAt = new Date().toISOString(); trash.unshift(x); delete rrS[x.id]; save(); saveTrash(); draw();
      undoBar(`نُقلت الفاتورة إلى «الفواتير المحذوفة» أسفل الصفحة`, () => { trash.splice(trash.indexOf(x), 1); delete x.deletedAt; delete x.delReason; list.splice(i, 0, x); save(); saveTrash(); draw(); }, () => { });
    }
    if (b.dataset.view) view(b.dataset.view);
    if (b.dataset.pick) openPick(b.dataset.pick, b.dataset.pk);
    if (b.dataset.zdel) { zConfirm = b.dataset.zdel; drawZones(); const n = document.querySelector('[data-zno]'); if (n) n.focus(); }
    if (b.dataset.zno !== undefined) { const k = zConfirm; zConfirm = null; drawZones(); const d = document.querySelector(`[data-zdel="${CSS.escape(k || '')}"]`); if (d) d.focus(); }
    if (b.dataset.zok) {
      const k = b.dataset.zok, t = zones[k]; if (!t) return;
      delete zones[k]; zConfirm = null; saveZones(); zMark(k); drawZones();
      undoBar(`حُذف قالب «${t.name || 'المورد'}»`, () => { zones[k] = t; saveZones(); zMark(k); drawZones(); }, () => { });
    }
    if (b.dataset.reread) reread(b.dataset.reread);
    if (b.dataset.trview) { const id = b.dataset.trview; trOpen[id] = !trOpen[id]; if (trOpen[id]) rv.open[id] = true; draw(); }
    if (b.dataset.trback) {
      const x = trash.find(y => y.id === b.dataset.trback); if (!x) return;
      trash.splice(trash.indexOf(x), 1); delete x.deletedAt; delete x.delReason; delete trOpen[x.id]; list.unshift(x); save(); saveTrash(); draw();
      toast('استُرجعت الفاتورة إلى القائمة'); const c = $('c-' + x.id); if (c) c.scrollIntoView({ block: 'start', behavior: RM ? 'auto' : 'smooth' });
    }
    if (b.dataset.trdel) { trAsk = b.dataset.trdel; draw(); const n = document.querySelector('[data-trno]'); if (n) n.focus(); }
    if (b.dataset.trno !== undefined) { const id = trAsk; trAsk = null; draw(); const d = document.querySelector(`[data-trdel="${CSS.escape(id || '')}"]`); if (d) d.focus(); }
    if (b.dataset.trok) {
      const x = trash.find(y => y.id === b.dataset.trok); if (!x) return;
      const i = trash.indexOf(x); trash.splice(i, 1); trAsk = null; delete trOpen[x.id]; saveTrash(); draw();
      undoBar('حُذفت الفاتورة نهائيًا', () => { trash.splice(i, 0, x); saveTrash(); draw(); }, () => IMG.del(x.id));
    }
    if (b.dataset.rvtg) { const id = b.dataset.rvtg; rv.open[id] = !isOpen(id); draw(); const t = document.querySelector(`[data-rvtg="${CSS.escape(id)}"]`); if (t) t.focus({ preventScroll: true }); }
    if (b.id === 'rvAll') { rv.hideAll = !rv.hideAll; rv.open = {}; draw(); }
    if (b.dataset.rr) rerunVendor(b.dataset.rr);
    if (b.dataset.rrno) { delete rrS[b.dataset.rrno]; draw(); }
    if (b.dataset.rrundo) undoRerun(b.dataset.rrundo);
  });

  // ---------- عرض الملف ----------
  let pages = [], pIdx = 0, zoom = 1, fileUrl = null;
  async function view(id) {
    const x = list.find(y => y.id === id); if (!x) return;
    let rec = await IMG.get(id);
    if (!rec && x.old) { const u = await OLD.get(id); if (u && u.length) rec = { pages: u.map(v => typeof v === 'string' ? v : URL.createObjectURL(v)) }; }
    pages = (rec && rec.pages) || []; pIdx = 0; zoom = 1;
    if (fileUrl) { URL.revokeObjectURL(fileUrl); fileUrl = null; }
    if (rec && rec.blob) fileUrl = URL.createObjectURL(rec.blob);
    $('lbT').textContent = x.file || '';
    $('lbOrig').hidden = !fileUrl; if (fileUrl) { $('lbOrig').href = fileUrl; $('lbOrig').download = x.file_name || 'invoice'; }
    $('lbNone').hidden = !!pages.length;
    showPage(); $('lb').hidden = false; document.body.style.overflow = 'hidden'; $('lbClose').focus();
  }
  function showPage() {
    $('lbImg').hidden = !pages.length; if (pages.length) $('lbImg').src = pages[pIdx];
    $('lbImg').style.width = (zoom * 100) + '%';
    $('lbPg').textContent = pages.length > 1 ? `صفحة ${pIdx + 1} من ${pages.length}` : '';
    $('lbPrev').hidden = $('lbNext').hidden = pages.length < 2; $('lbZ').textContent = Math.round(zoom * 100) + '%';
  }
  const closeLb = () => { $('lb').hidden = true; document.body.style.overflow = ''; };
  $('lbClose').onclick = closeLb;
  $('lb').addEventListener('click', e => { if (e.target === $('lb')) closeLb(); });
  document.addEventListener('keydown', e => { if (e.key === 'Escape' && !$('lb').hidden) closeLb(); });
  $('lbIn').onclick = () => { zoom = Math.min(4, zoom + 0.5); showPage(); };
  $('lbOut').onclick = () => { zoom = Math.max(1, zoom - 0.5); showPage(); };
  $('lbPrev').onclick = () => { pIdx = (pIdx - 1 + pages.length) % pages.length; showPage(); };
  $('lbNext').onclick = () => { pIdx = (pIdx + 1) % pages.length; showPage(); };

  // ---------- الرفع والقراءة ----------
  const queue = [], live = {}; let busy = false;
  const sig = f => String(f.name || '').trim().toLowerCase(); // منع التكرار باسم الملف فقط
  // ملخص آخر عملية رفع: عدد الملفات، وما أُضيف، والمكرر، وغير المدعوم، ونتيجة القراءة
  let upSum = null;
  const isInv = f => /^image\//.test(f.type) || /\.(pdf|jpe?g|png|webp|heic|bmp|gif)$/i.test(f.name || '') || f.type === 'application/pdf';
  function addFiles(files, fromCam, kind) {
    const seen = new Set(list.map(x => x.sig).filter(Boolean).concat(queue.map(j => sig(j.file))));
    const all = [...(files || [])].filter(f => !/(^|\/)\.[^/]*$/.test(pathOf(f))); // نتجاهل ملفات النظام المخفية مثل .DS_Store
    const b = { kind: fromCam ? 'cam' : kind || 'files', total: all.length, added: 0, dups: [], other: [], read: 0, failed: 0 };
    all.forEach((f, i) => {
      if (!isInv(f)) { b.other.push(pathOf(f)); return; }
      const s = sig(f); if (!fromCam && seen.has(s)) { b.dups.push(pathOf(f)); return; } seen.add(s);
      const id = Date.now().toString(36) + '-' + i + '-' + Math.random().toString(36).slice(2, 6);
      queue.push({ id, file: f, batch: b }); live[id] = { status: 'queued', file: pathOf(f) || 'صورة من الكاميرا' }; b.added++;
    });
    if (!fromCam || b.dups.length || b.other.length) upSum = b;
    drawSum(); draw(); pump();
  }
  function drawSum() {
    const el = $('upSum'), b = upSum; el.hidden = !b; if (!b) return;
    const where = b.kind === 'dir' ? 'عدد الملفات في المجلد' : b.kind === 'drop' ? 'عدد الملفات المسحوبة' : b.kind === 'cam' ? 'عدد الصور من الكاميرا' : 'عدد الملفات المختارة';
    const left = b.added - b.read - b.failed;
    const names = (t, a) => a.length ? `<details><summary>${t} (${a.length})</summary><ul>${a.map(n => `<li>${esc(n)}</li>`).join('')}</ul></details>` : '';
    el.innerHTML = `<div class="ushead"><b>ملخص آخر رفع</b><button type="button" class="usx" id="upSumX" aria-label="إغلاق الملخص">✕</button></div>
      <ul class="usn"><li><span>${where}</span><b>${b.total}</b></li>
        <li class="ok"><span>أُضيفت للقراءة</span><b>${b.added}</b></li>
        <li class="${b.dups.length ? 'bad' : ''}"><span>لم تُضف لأنها مكررة (نفس اسم ملف في القائمة)</span><b>${b.dups.length}</b></li>
        ${b.other.length ? `<li><span>ليست صورة أو PDF فتم تجاهلها</span><b>${b.other.length}</b></li>` : ''}
        ${b.added ? `<li><span>${left > 0 ? 'قُرئت حتى الآن' : 'قُرئت'}</span><b>${b.read}${b.failed ? ` <small>(وتعذّرت قراءة ${b.failed})</small>` : ''}${left > 0 ? ` <small>من ${b.added}</small>` : ''}</b></li>` : ''}</ul>
      ${names('أسماء الملفات المكررة', b.dups)}${names('الملفات التي تم تجاهلها', b.other)}`;
  }
  document.addEventListener('click', e => { if (e.target.closest('#upSumX')) { upSum = null; drawSum(); } });
  async function reread(id) {
    const x = list.find(y => y.id === id); if (!x || live[id]) return;
    const rec = await IMG.get(id);
    let file = rec && rec.blob;
    if (!file) { // فواتير منقولة من النسخة السابقة: نعيد القراءة من الصورة المحفوظة
      let pg = rec && rec.pages && rec.pages[0];
      if (!pg && x.old) { const u = await OLD.get(id); pg = u && u[0]; }
      if (pg) file = typeof pg === 'string' ? await (await fetch(pg)).blob() : pg;
    }
    if (!file) { toast('لا يوجد ملف محفوظ لهذه الفاتورة. ارفعها من جديد.'); return; }
    if (!(file instanceof File)) file = new File([file], x.file_name || x.file || 'invoice.jpg', { type: file.type || 'image/jpeg' });
    queue.push({ id, file, reread: true }); live[id] = { status: 'queued', file: x.file, reread: true };
    draw(); pump();
  }
  async function pump() {
    if (busy) return; busy = true;
    let done = 0;
    while (queue.length) {
      const job = queue.shift(); live[job.id].status = 'working'; draw();
      const N = done + 1 + queue.length;
      const say = m => status(N > 1 ? `فاتورة ${done + 1} من ${N}: ${m}` : m);
      progress(0.02);
      if (job.reread) {
        const x = list.find(y => y.id === job.id);
        deep = true;
        try {
          const { d, pages, src } = await processFile(job.file, say);
          let changed = 0;
          F.forEach(([k]) => { const nv = k === 'total' || k === 'tax_amount' ? (d[k] ?? '') : k === 'payment_method' ? payNorm(d[k]) : (d[k] || ''); if (x && !(x.edited && x.edited[k]) && !blank(nv) && nv !== x[k]) { x[k] = nv; if (x.zsrc) delete x.zsrc[k]; changed++; } });
          if (x && zoneFor(x)) { say('جارٍ تطبيق قالب المورد…'); try { changed += await applyZones(x, src); } catch (e) { console.warn(e); } }
          if (x) { save(); delete rvPages[job.id]; IMG.get(job.id).then(r0 => IMG.put(job.id, { pages, blob: (r0 && r0.blob) || job.file })); }
          toast(changed ? `أُعيدت القراءة: تحدّثت ${changed} من الخانات` : 'أُعيدت القراءة ولم تتغير البيانات');
        } catch (e) { console.error(e); toast('تعذّرت إعادة القراءة'); }
        deep = false; delete live[job.id]; done++; draw(); continue;
      }
      try {
        const { d, pages, src } = await processFile(job.file, say);
        const x = { id: job.id, at: new Date().toISOString(), file: d.file, file_name: d.file_name, sig: sig(job.file),
          invoice_number: d.invoice_number || '', invoice_date: d.invoice_date || '', total: d.total ?? '', tax_amount: d.tax_amount ?? '',
          vendor_name: d.vendor_name || '', vendor_tax_number: d.vendor_tax_number || '', payment_method: payNorm(d.payment_method), note: '' };
        if (zoneFor(x)) { say('جارٍ تطبيق قالب المورد…'); try { await applyZones(x, src); } catch (e) { console.warn(e); } }
        list.unshift(x); save(); if (job.batch) { job.batch.read++; if (job.batch === upSum) drawSum(); }
        IMG.put(job.id, { pages, blob: job.file.size < 25e6 ? job.file : null });
      } catch (e) {
        console.error(e);
        const x = { id: job.id, at: new Date().toISOString(), file: live[job.id].file, sig: sig(job.file), invoice_number: '', invoice_date: '', total: '', tax_amount: '', vendor_name: '', vendor_tax_number: '', payment_method: '', note: 'تعذّرت قراءة الملف تلقائيًا: أدخل البيانات يدويًا' };
        list.unshift(x); save(); IMG.put(job.id, { pages: [], blob: job.file }); if (job.batch) { job.batch.failed++; if (job.batch === upSum) drawSum(); }
      }
      delete live[job.id]; done++; draw();
    }
    busy = false; progress(null);
    status(done ? (done > 1 ? `انتهت القراءة: ${done} فواتير. راجع الخانات الحمراء وأكملها.` : 'تمت القراءة. راجع الخانات وأكمل الناقص.') : '');
  }
  $('pick').addEventListener('change', e => { addFiles(e.target.files, false, 'files'); e.target.value = ''; });
  $('dir').addEventListener('change', e => { addFiles(e.target.files, false, 'dir'); e.target.value = ''; });
  $('cam').addEventListener('change', e => { addFiles(e.target.files, true); e.target.value = ''; });
  const drop = $('drop');
  ['dragenter', 'dragover'].forEach(t => drop.addEventListener(t, e => { e.preventDefault(); drop.classList.add('over'); }));
  ['dragleave', 'drop'].forEach(t => drop.addEventListener(t, e => { e.preventDefault(); drop.classList.remove('over'); }));
  drop.addEventListener('drop', e => addFiles(e.dataTransfer.files, false, 'drop'));

  // ---------- Excel ----------
  const XL = [['invoice_number', 'رقم الفاتورة', 18], ['invoice_date', 'تاريخ الفاتورة', 14], ['total', 'المبلغ الإجمالي', 16], ['tax_amount', 'مبلغ الضريبة', 14],
    ['vendor_name', 'اسم المورد', 32], ['vendor_tax_number', 'الرقم الضريبي', 20], ['payment_method', 'طريقة الدفع', 14], ['project', 'المشروع', 12], ['note', 'ملاحظات', 50], ['file', 'الملف', 30]];
  const xlCell = (x, k) => {
    if (k === 'note') return noteOf(x);
    if (k === 'total' || k === 'tax_amount') { const v = n2(x[k]); return typeof v === 'number' ? v : (v || ''); }
    return x[k] || '';
  };
  function download(blob, name) {
    const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name;
    document.body.append(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 4000);
  }
  $('xl').onclick = () => {
    const rows = [XL.map(c => c[1]), ...list.map(x => XL.map(([k]) => xlCell(x, k)))];
    download(X.xlsx([{ name: 'الفواتير', rows, widths: XL.map(c => c[2]) }]), `الفواتير-${new Date().toISOString().slice(0, 10)}.xlsx`);
    toast('جارٍ تنزيل ملف Excel');
  };

  // ---------- تصدير كل الفواتير ثم حذفها (يطلب تأكيدًا ورمزًا) ----------
  // الرمز لا يُحفظ نصًا في الكود: نقارن بصمته فقط (SHA-256)
  const WIPE_HASH = '411530bcc824c1ca5dfed1470d536c65212254a4bfe79ff4e766eb1c1ea5b417';
  async function wipeCodeOk(v) {
    const t = P.normDigits(String(v || '')).replace(/\s/g, '').replace(/[‐-―−ـ]/g, '-'); if (!t) return false;
    try { const d = await crypto.subtle.digest('SHA-256', new TextEncoder().encode('invoice-wipe:' + t)); return [...new Uint8Array(d)].map(b => b.toString(16).padStart(2, '0')).join('') === WIPE_HASH; } catch { return false; }
  }
  const stamp = s => { if (!s) return ''; const d = new Date(s); return isNaN(d) ? s : d.toISOString().slice(0, 16).replace('T', ' '); };
  const XL2 = [...XL.slice(0, -1), ['file', 'الملف', 30], ['at', 'وقت القراءة', 17], ['edited', 'خانات عُدّلت يدويًا', 26], ['tpl', 'خانات من قالب المورد', 26]];
  const xlFull = (x, k) => k === 'at' ? stamp(x.at) : k === 'edited' ? Object.keys(x.edited || {}).map(label).join('، ')
    : k === 'tpl' ? Object.keys(x.zsrc || {}).filter(f => x.zsrc[f] !== 'pick').map(label).join('، ') : xlCell(x, k);
  function wipeBook() {
    const sh = [{ name: 'الفواتير', rows: [XL2.map(c => c[1]), ...list.map(x => XL2.map(([k]) => xlFull(x, k)))], widths: XL2.map(c => c[2]) }];
    if (trash.length) { const T = [...XL2, ['delReason', 'سبب الحذف', 30], ['deletedAt', 'وقت الحذف', 17]];
      sh.push({ name: 'المحذوفة', rows: [T.map(c => c[1]), ...trash.map(x => T.map(([k]) => k === 'delReason' ? x.delReason || '' : k === 'deletedAt' ? stamp(x.deletedAt) : xlFull(x, k)))], widths: T.map(c => c[2]) }); }
    return X.xlsx(sh);
  }
  function drawWipe() {
    const n = list.length, m = trash.length, el = $('wpBox');
    $('wpBtn').disabled = !n && !m; $('wpBtn').hidden = !el.hidden;
    if (el.hidden) return;
    el.innerHTML = `<b>تصدير وحذف جميع الفواتير؟</b>
      <p>سيُنزَّل ملف Excel فيه ${n} فاتورة${m ? ` و${m} فاتورة محذوفة (في ورقة مستقلة)` : ''} بكل بياناتها، ثم تُحذف كل الفواتير وصورها من هذا المتصفح. قوالب الموردين وقائمة المشاريع تبقى كما هي.</p>
      <label for="wpCode">اكتب رمز التأكيد للمتابعة</label><input id="wpCode" inputmode="numeric" autocomplete="off" placeholder="الرمز">
      <div class="zbtns"><button type="button" class="del" id="wpGo" disabled>صدّر إلى Excel واحذف الكل</button><button type="button" id="wpNo">تراجع</button></div>`;
  }
  $('wpBtn').onclick = () => { if (busy || queue.length) { toast('انتظر حتى تنتهي قراءة الفواتير الجارية'); return; } $('wpBox').hidden = false; drawWipe(); $('wpCode').focus(); };
  document.addEventListener('input', async e => { if (e.target.id !== 'wpCode') return; const ok = await wipeCodeOk(e.target.value); const g = $('wpGo'); if (g && $('wpCode') === e.target) g.disabled = !ok; });
  document.addEventListener('click', async e => {
    const b = e.target.closest('button'); if (!b) return;
    if (b.id === 'wpNo') { $('wpBox').hidden = true; drawWipe(); $('wpBtn').focus(); }
    if (b.id === 'wpGo') {
      if (!(await wipeCodeOk($('wpCode').value))) { toast('الرمز غير صحيح'); return; }
      if (busy || queue.length) { toast('انتظر حتى تنتهي قراءة الفواتير الجارية'); return; }
      let blob; try { blob = wipeBook(); } catch (er) { console.error(er); toast('تعذّر إنشاء ملف Excel، لم يُحذف شيء'); return; }
      download(blob, `all-invoices-${new Date().toISOString().slice(0, 16).replace(/[T:]/g, '-')}.xlsx`);
      const oldL = list.slice(), oldT = trash.slice();
      list.length = 0; trash.length = 0; for (const k in rrS) delete rrS[k]; upSum = null; drawSum();
      save(); saveTrash(); $('wpBox').hidden = true; draw();
      undoBar(`صُدّرت ${oldL.length + oldT.length} فاتورة إلى Excel وحُذفت`, () => { list.push(...oldL); trash.push(...oldT); save(); saveTrash(); draw(); },
        () => { for (const k in rvPages) delete rvPages[k]; IMG.clear(); });
    }
  });

  // ---------- شريط التراجع ----------
  function undoBar(msg, undo, commit) {
    const u = $('undo'); clearTimeout(undoBar.t); if (undoBar.c) undoBar.c();
    u.querySelector('span').textContent = msg; u.hidden = false; undoBar.c = commit;
    u.querySelector('button').onclick = () => { clearTimeout(undoBar.t); undoBar.c = null; u.hidden = true; undo(); };
    undoBar.t = setTimeout(() => { u.hidden = true; if (undoBar.c) { undoBar.c(); undoBar.c = null; } }, 7000);
  }


  // ---------- المشروع المرتبط بالفاتورة ----------
  const PK = 'invoice-simple-projects-v1';
  let projects;
  try { projects = JSON.parse(localStorage.getItem(PK)); } catch { projects = null; }
  if (!Array.isArray(projects)) projects = ['3', '5', '6', '7'].map(name => ({ name, hidden: false }));
  const saveProjects = () => { try { localStorage.setItem(PK, JSON.stringify(projects)); } catch { } };
  const projKey = s => P.normDigits(String(s || '')).trim().toLowerCase();
  function projHtml(x) {
    const cur = x.project || '', vis = projects.filter(p => !p.hidden);
    const extra = cur && !vis.some(p => p.name === cur) ? `<option value="${esc(cur)}" selected>${esc(cur)} (مخفي)</option>` : '';
    return `<div class="f f-project${cur ? '' : ' nopj'}"><span class="fh"><label for="pj-${esc(x.id)}">المشروع</label></span>
      <select id="pj-${esc(x.id)}" data-proj="${esc(x.id)}"><option value=""${cur ? '' : ' selected'}>اختر المشروع</option>${vis.map(p => `<option value="${esc(p.name)}"${p.name === cur ? ' selected' : ''}>${esc(p.name)}</option>`).join('')}${extra}
      <option disabled>──────</option><option value="__add">＋ إضافة مشروع جديد…</option><option value="__manage">إدارة قائمة المشاريع…</option></select></div>`;
  }
  document.addEventListener('change', e => {
    const el = e.target; if (!el.dataset || !el.dataset.proj) return;
    const x = list.find(y => y.id === el.dataset.proj); if (!x) return;
    if (el.value === '__add' || el.value === '__manage') { const m = el.value; el.value = x.project || ''; openProj(m === '__add' ? 'add' : 'manage', x.id); return; }
    x.project = el.value; save(); el.closest('.f').classList.toggle('nopj', !x.project);
    toast(x.project ? `رُبطت الفاتورة بالمشروع ${x.project}` : 'أُزيل المشروع من الفاتورة');
  });
  // نافذة إضافة المشاريع وإخفائها، وكل تغيير يمر بخطوة تأكيد
  const pj = { forId: null, ask: null };
  function openProj(mode, forId) {
    pj.forId = forId || null; pj.ask = null; pj.last = document.activeElement;
    $('pjNew').value = ''; drawProj(); $('pj').hidden = false; document.body.style.overflow = 'hidden';
    (mode === 'add' ? $('pjNew') : $('pjClose')).focus();
  }
  function closeProj() { $('pj').hidden = true; document.body.style.overflow = ''; pj.ask = null; if (pj.last && document.contains(pj.last)) pj.last.focus(); }
  function drawProj() {
    const used = n => list.filter(x => x.project === n).length;
    const a = pj.ask;
    $('pjAsk').hidden = !a;
    if (a) {
      $('pjAskT').textContent = a.type === 'add' ? `تأكيد إضافة المشروع «${a.name}» إلى القائمة؟` : a.type === 'hide' ? `تأكيد إخفاء المشروع «${a.name}» من القائمة؟` : `إظهار المشروع «${a.name}» في القائمة من جديد؟`;
      $('pjAskS').textContent = a.type === 'add' ? (pj.forId ? 'سيُضاف للقائمة ويُربط بهذه الفاتورة.' : 'سيظهر في قائمة المشاريع لكل الفواتير.') : a.type === 'hide' ? (used(a.name) ? `لن يظهر عند اختيار مشروع لفاتورة جديدة. الفواتير المرتبطة به (${used(a.name)}) تبقى كما هي.` : 'لن يظهر عند اختيار مشروع. يمكنك إظهاره لاحقًا من هنا.') : 'سيعود للظهور عند اختيار مشروع.';
      $('pjYes').textContent = a.type === 'add' ? 'نعم، أضف' : a.type === 'hide' ? 'نعم، أخفِ' : 'نعم، أظهر';
      $('pjYes').className = a.type === 'hide' ? 'del' : 'primary';
    }
    $('pjList').innerHTML = projects.map((p, i) => `<li class="${p.hidden ? 'off' : ''}"><b>${esc(p.name)}</b><small>${p.hidden ? 'مخفي' : 'ظاهر'}${used(p.name) ? ' · ' + used(p.name) + ' فاتورة' : ''}</small>
      <button data-pjt="${i}"${a ? ' disabled' : ''}>${p.hidden ? 'إظهار' : 'إخفاء'}</button></li>`).join('');
    $('pjAdd').disabled = !!a; $('pjNew').disabled = !!a;
  }
  $('pjForm').addEventListener('submit', e => {
    e.preventDefault();
    const name = P.normDigits($('pjNew').value).trim().replace(/\s+/g, ' ');
    if (!name) { $('pjNew').focus(); return; }
    const ex = projects.find(p => projKey(p.name) === projKey(name));
    if (ex && !ex.hidden) { toast(`المشروع ${ex.name} موجود في القائمة`); return; }
    pj.ask = ex ? { type: 'show', name: ex.name } : { type: 'add', name }; drawProj(); $('pjYes').focus();
  });
  $('pjList').addEventListener('click', e => {
    const b = e.target.closest('[data-pjt]'); if (!b) return;
    const p = projects[+b.dataset.pjt]; pj.ask = { type: p.hidden ? 'show' : 'hide', name: p.name }; drawProj(); $('pjYes').focus();
  });
  $('pjNo').onclick = () => { pj.ask = null; drawProj(); $('pjNew').focus(); };
  $('pjYes').onclick = () => {
    const a = pj.ask; if (!a) return;
    if (a.type === 'add') { projects.push({ name: a.name, hidden: false }); $('pjNew').value = ''; }
    else { const p = projects.find(q => q.name === a.name); if (p) p.hidden = a.type === 'hide'; }
    saveProjects();
    if (a.type !== 'hide' && pj.forId) { const x = list.find(y => y.id === pj.forId); if (x) { x.project = a.name; save(); } }
    pj.ask = null;
    const msg = a.type === 'add' ? `أُضيف المشروع ${a.name}` : a.type === 'hide' ? `أُخفي المشروع ${a.name}` : `ظهر المشروع ${a.name}`;
    if (a.type === 'add' && pj.forId) { closeProj(); draw(); const el = $('pj-' + pj.forId); if (el) { el.scrollIntoView({ block: 'center' }); el.focus({ preventScroll: true }); } }
    else { drawProj(); draw(); }
    toast(msg);
  };
  $('pjClose').onclick = closeProj;
  $('pj').addEventListener('click', e => { if (e.target === $('pj')) closeProj(); });
  document.addEventListener('keydown', e => { if (e.key === 'Escape' && !$('pj').hidden) { if (pj.ask) { pj.ask = null; drawProj(); } else closeProj(); } });
  $('pjOpen').onclick = () => openProj('manage');

  // ---------- قوالب الموردين: مكان كل خانة على فاتورة المورد ----------
  // يُحفظ المكان كنسبة من عرض الصفحة وطولها (لا بالبكسل)، مع رقم الصفحة أو «الأخيرة».
  const ZK = 'invoice-simple-zones-v1';
  let zones = {}, zConfirm = null;
  try { zones = JSON.parse(localStorage.getItem(ZK)) || {}; } catch { zones = {}; }
  const saveZones = () => { try { localStorage.setItem(ZK, JSON.stringify(zones)); } catch { toast('تعذّر حفظ القالب: مساحة المتصفح ممتلئة'); } };
  const vatOf = x => { const d = P.normDigits(x.vendor_tax_number || '').replace(/\D/g, ''); return d.length === 15 ? d : ''; };
  const nameKey = s => P.normDigits(s || '').normalize('NFKC').toLowerCase().replace(/[^\p{L}\p{N}]+/gu, '');
  function vendorKey(x) { const v = vatOf(x); if (v) return 'vat:' + v; const n = nameKey(x.vendor_name); return n.length >= 3 ? 'name:' + n : ''; }
  function zoneFor(x) { // نبحث بالرقم الضريبي أولًا ثم باسم المورد
    const v = vatOf(x); if (v && zones['vat:' + v]) return 'vat:' + v;
    const n = nameKey(x.vendor_name); if (n.length < 3) return '';
    if (zones['name:' + n]) return 'name:' + n;
    return Object.keys(zones).find(k => zones[k].nameKey === n) || '';
  }
  const pgIndex = (pg, n) => pg === 'last' ? n - 1 : (pg < n ? pg : -1);
  // الصفحة الأولى للبيانات العلوية، و«الأخيرة» للمبالغ؛ فيصلح القالب لفاتورة من صفحة أو عدة صفحات
  const pgFor = (k, i, n) => n === 1 ? (['total', 'tax_amount', 'payment_method'].includes(k) ? 'last' : 0) : (i === 0 ? 0 : i === n - 1 ? 'last' : i);
  const pgName = pg => pg === 'last' ? 'الصفحة الأخيرة' : 'صفحة ' + (pg + 1);
  const label = k => (F.find(f => f[0] === k) || [, k])[1];

  // قارئ مستقل للمناطق حتى لا يتداخل مع قراءة الفواتير الجارية
  let zw = null, zq = Promise.resolve();
  async function zoneWorker() {
    if (!zw) zw = Tesseract.createWorker('ara+eng', 1, { workerPath: abs('vendor/worker.min.js'), corePath: abs('vendor/core'), langPath: abs('vendor/lang'), gzip: true })
      .then(async w => { await w.setParameters({ preserve_interword_spaces: '1' }); return w; });
    try { return await zw; } catch (e) { zw = null; throw e; }
  }
  const serial = fn => (zq = zq.then(fn, fn));
  // قراءة النص داخل مستطيل: من نص PDF الرقمي إن وُجد، وإلا بقراءة المنطقة المقصوصة فقط
  function readRegion(src, r) {
    if (src.items && src.items.length) {
      const hit = src.items.filter(it => { const cx = it.x + it.w / 2, cy = it.y + it.h / 2; return cx >= r.x && cx <= r.x + r.w && cy >= r.y && cy <= r.y + r.h; });
      if (hit.length) {
        const rows = [];
        hit.sort((a, b) => a.y - b.y).forEach(it => { let row = rows.find(q => Math.abs(q.y - it.y) < it.h * 0.6); if (!row) rows.push(row = { y: it.y, p: [] }); row.p.push(it); });
        return Promise.resolve(rows.map(q => { const rtl = /[؀-ۿ]/.test(q.p.map(i => i.s).join('')); q.p.sort((a, b) => rtl ? b.x - a.x : a.x - b.x); return q.p.map(i => i.s.trim()).join(' '); }).join('\n'));
      }
    }
    return serial(async () => {
      const C = src.canvas, W = C.width, H = C.height, pad = 0.004;
      const sx = Math.max(0, (r.x - pad) * W), sy = Math.max(0, (r.y - pad) * H);
      const sw = Math.min(W, (r.x + r.w + pad) * W) - sx, sh = Math.min(H, (r.y + r.h + pad) * H) - sy;
      if (sw < 4 || sh < 4) return '';
      const k = Math.min(4, Math.max(1, 70 / sh), 2400 / sw), m = 24;
      const c = document.createElement('canvas'); c.width = Math.round(sw * k) + 2 * m; c.height = Math.round(sh * k) + 2 * m;
      const g = c.getContext('2d', { willReadFrequently: true }); g.fillStyle = '#fff'; g.fillRect(0, 0, c.width, c.height);
      g.imageSmoothingQuality = 'high'; g.drawImage(C, sx, sy, sw, sh, m, m, c.width - 2 * m, c.height - 2 * m);
      const d = g.getImageData(0, 0, c.width, c.height), a = d.data;
      for (let i = 0; i < a.length; i += 4) { const y = 0.299 * a[i] + 0.587 * a[i + 1] + 0.114 * a[i + 2]; a[i] = a[i + 1] = a[i + 2] = y; }
      g.putImageData(d, 0, 0);
      const w = await zoneWorker();
      await w.setParameters({ tessedit_pageseg_mode: sh * k < 130 ? '7' : '6' });
      const { data } = await w.recognize(c);
      return data.text || '';
    });
  }
  // استخراج قيمة الخانة من النص المقروء
  function extract(k, raw) {
    const t = P.normDigits(raw || '').replace(/[ \t]+/g, ' ').trim(); if (!t) return '';
    if (k === 'total' || k === 'tax_amount') {
      const ns = P.numsIn(t).filter(o => o.n > 0 && !/^\d{1,2}%/.test(t.slice(o.i)));
      if (!ns.length) return '';
      const dec = ns.filter(o => /[.,]\d{2}$/.test(o.raw));
      return Math.round((dec.length ? dec : ns).reduce((a, b) => b.n > a.n ? b : a).n * 100) / 100;
    }
    if (k === 'invoice_date') return P.findDate(t) || '';
    if (k === 'vendor_tax_number') {
      const runs = (t.match(/\d[\d ]{12,}\d/g) || []).map(s => s.replace(/ /g, ''));
      return runs.find(s => /^(3\d{13}3|100\d{12})$/.test(s)) || runs.find(s => s.length === 15) || runs.sort((a, b) => b.length - a.length)[0] || '';
    }
    if (k === 'invoice_number') {
      const c = P.invCands ? P.invCands(t.replace(/\n/g, ' ')) : [];
      const lab = v => v.replace(/^(?:ref|reference|no|num|number|invoice|inv|bill|doc)\s*[.:#]\s*(?=[A-Za-z0-9/-]*\d)/i, '');
      if (c.length) return lab(c.sort((a, b) => b.t.length - a.t.length)[0].t);
      const m = t.match(/[A-Za-z0-9][A-Za-z0-9\-\/]*\d[A-Za-z0-9\-\/]*/g); return m ? m.sort((a, b) => b.length - a.length)[0] : '';
    }
    if (k === 'payment_method') { const p = PAYMAP.find(([re]) => re.test(t)); if (p) return p[1]; }
    return t.replace(/\n+/g, ' ').replace(/^(اسم\s*)?(المورد|البائع|التاجر|الشركة|طريقة\s*الدفع|الدفع)\s*[:：-]?\s*/, '').replace(/^(vendor|seller|supplier|payment(\s*method)?)(\s*name)?\s*[:：-]?\s*/i, '').trim();
  }
  // تطبيق القالب على فاتورة مقروءة. يعيد عدد الخانات التي أُخذت من القالب.
  async function applyZones(x, src) {
    const key = zoneFor(x); if (!key || !src || !src.length) return 0;
    let n = 0;
    for (const [k, r] of Object.entries(zones[key].fields || {})) {
      if (x.edited && x.edited[k]) continue;
      const i = pgIndex(r.pg, src.length); if (i < 0) continue;
      const v = extract(k, await readRegion(src[i], r));
      if (!blank(v)) { if (x[k] !== v) n++; x[k] = v; (x.zsrc = x.zsrc || {})[k] = 'tpl'; (x.zrect = x.zrect || {})[k] = { ...r }; }
    }
    return n;
  }
  function drawZones() {
    const keys = Object.keys(zones).filter(k => Object.keys(zones[k].fields || {}).length);
    $('zSec').hidden = false; $('zCount').textContent = keys.length ? `(${keys.length})` : ''; drawSync();
    if (!keys.length) { $('zList').innerHTML = '<p class="note">لا توجد قوالب بعد. حدّد مكان أي خانة بزر التحديد بجانبها ليُحفظ قالب للمورد.</p>'; return; }
    $('zList').innerHTML = keys.map(k => { const t = zones[k];
      return `<div class="zc"><b>${esc(t.name || (t.vat ? 'مورد بالرقم الضريبي' : 'مورد'))}</b><small>${t.vat ? 'الرقم الضريبي ' + esc(t.vat) : 'بدون رقم ضريبي، يُعرف بالاسم'}</small>
        <div>${Object.entries(t.fields).map(([f, r]) => `<span class="zchip">${esc(label(f))} · ${pgName(r.pg)}</span>`).join('')}</div>
        ${zConfirm === k ? `<div class="zask" role="alert"><b>حذف قالب هذا المورد؟</b>ستعود فواتيره الجديدة للقراءة العادية، وتحتاج لتحديد الخانات من جديد.<div class="zbtns"><button class="del" data-zok="${esc(k)}">نعم، احذف</button><button data-zno="">تراجع</button></div></div>`
          : `<div class="zbtns"><button class="del" data-zdel="${esc(k)}">حذف القالب</button></div>`}</div>`; }).join('');
  }

  // ---------- نافذة تحديد مكان الخانة ----------
  const pk = { x: null, k: null, pages: [], src: [], rec: null, pdf: null, pg: 0, rect: null, zoom: 1, mode: 'draw', tok: 0 };
  const pkVp = $('pkVp'), pkPaper = $('pkPaper'), pkImg = $('pkImg'), pkSel = $('pkSel'), pkPad = $('pkPad');
  async function getRec(id, x) {
    let rec = await IMG.get(id);
    if (!rec && x.old) { const u = await OLD.get(id); if (u && u.length) rec = { pages: u.map(v => typeof v === 'string' ? v : URL.createObjectURL(v)) }; }
    return rec;
  }
  const pageSrc = i => loadSrc(pk, i);
  async function loadSrc(c, i) { // نسخة عالية الدقة من الصفحة للقراءة (تُحمّل عند الحاجة فقط). c = {rec, pages, src, pdf}
    if (c.src[i]) return c.src[i];
    const b = c.rec && c.rec.blob;
    let s = null;
    try {
      if (b && (b.type === 'application/pdf' || /\.pdf$/i.test(b.name || '')) && window.pdfjsLib) {
        c.pdf = c.pdf || await pdfjsLib.getDocument({ data: await b.arrayBuffer() }).promise;
        s = await renderPdfPage(c.pdf, i + 1);
      } else if (b && i === 0 && /^image\//.test(b.type || 'image/')) {
        const u = URL.createObjectURL(b); try { s = { canvas: toCanvas(await loadImg(u), 2400) }; } finally { URL.revokeObjectURL(u); }
      }
    } catch (e) { console.warn(e); }
    if (!s) s = { canvas: toCanvas(await loadImg(c.pages[i]), 2400) };
    return (c.src[i] = s);
  }
  async function openPick(id, k) {
    const x = list.find(y => y.id === id); if (!x) return;
    const rec = await getRec(id, x);
    if (!rec || !rec.pages || !rec.pages.length) { // نفتح «مراجعة الفاتورة» وفيها زر لإرفاق الملف
      toast('لا توجد صورة محفوظة لهذه الفاتورة على هذا الجهاز. أرفق ملفها من «مراجعة الفاتورة»');
      rv.open[id] = true; draw(); const h = $('rvh-' + id); if (h) h.scrollIntoView({ block: 'center', behavior: RM ? 'auto' : 'smooth' }); return; }
    Object.assign(pk, { x, k, rec, pages: rec.pages, src: [], pdf: null, rect: null, mode: 'draw', last: document.activeElement });
    const key = zoneFor(x), z = key && zones[key].fields[k];
    $('pkT').textContent = 'حدد مكان: ' + label(k);
    $('pkP').textContent = 'كبّر الصفحة وحرّكها حتى ترى ' + label(k) + ' بوضوح، ثم اسحب مستطيلًا حول القيمة. لا يضر أن يدخل معها اسم الخانة.';
    $('pkRo').hidden = true; $('pkOk').disabled = true; $('pkVal').value = '';
    const canSave = !!(vendorKey(x) || key);
    $('pkSaveL').hidden = !canSave; $('pkSave').checked = true;
    $('pkSaveT').textContent = 'احفظ هذا المكان لكل فواتير «' + (x.vendor_name || 'هذا المورد') + '»';
    $('pkNo').textContent = canSave ? '' : 'لن يُحفظ المكان قالبًا لأن اسم المورد ورقمه الضريبي غير معروفين. أكملهما أولًا ليُحفظ.';
    $('pk').hidden = false; document.body.style.overflow = 'hidden';
    setMode('draw');
    await showPg(z ? Math.max(0, pgIndex(z.pg, pk.pages.length)) : 0);
    pkVp.focus({ preventScroll: true });
  }
  function closePick() { $('pk').hidden = true; document.body.style.overflow = ''; pk.tok++; pk.src = []; pk.pdf = null; }
  function showPg(i) {
    const n = pk.pages.length; pk.pg = Math.max(0, Math.min(n - 1, i));
    pk.rect = null; pkSel.hidden = true; $('pkRo').hidden = true; $('pkOk').disabled = true;
    $('pkPg').textContent = n > 1 ? `صفحة ${pk.pg + 1} من ${n}` : ''; $('pkPrev').hidden = $('pkNext').hidden = n < 2;
    $('pkPrev').disabled = pk.pg === 0; $('pkNext').disabled = pk.pg === n - 1;
    const key = zoneFor(pk.x), z = key && zones[key].fields[pk.k], old = $('pkOld');
    old.hidden = !(z && pgIndex(z.pg, n) === pk.pg);
    if (!old.hidden) Object.assign(old.style, { left: z.x * 100 + '%', top: z.y * 100 + '%', width: z.w * 100 + '%', height: z.h * 100 + '%' });
    return new Promise(ok => { pkImg.onload = () => { setZoom(1); pkVp.scrollTop = 0; ok(); }; pkImg.onerror = ok; pkImg.src = pk.pages[pk.pg]; });
  }
  // التكبير والتحريك
  const ZMIN = 1, ZMAX = 6, RM = matchMedia('(prefers-reduced-motion: reduce)').matches;
  const fitW = () => { const a = (pkImg.naturalWidth || 1) / (pkImg.naturalHeight || 1); return Math.max(160, Math.min(pkVp.clientWidth - 24, (pkVp.clientHeight - 24) * a)); };
  function setZoom(z, cx, cy) {
    z = Math.max(ZMIN, Math.min(ZMAX, z));
    const a = pkPaper.getBoundingClientRect();
    if (cx == null) { const v = pkVp.getBoundingClientRect(); cx = v.left + v.width / 2; cy = v.top + v.height / 2; }
    const fx = a.width ? (cx - a.left) / a.width : .5, fy = a.height ? (cy - a.top) / a.height : .5;
    pk.zoom = z; pkPaper.style.width = fitW() * z + 'px';
    const b = pkPaper.getBoundingClientRect();
    pkVp.scrollLeft += b.left + fx * b.width - cx; pkVp.scrollTop += b.top + fy * b.height - cy;
    $('pkZ').textContent = Math.round(z * 100) + '%'; $('pkOut').disabled = z <= ZMIN; $('pkIn').disabled = z >= ZMAX;
  }
  let zAnim = 0, glide = 0, rafId = 0, job = null, spaceHeld = false;
  const frame = fn => { job = fn; if (!rafId) rafId = requestAnimationFrame(() => { rafId = 0; const f = job; job = null; if (f) f(); }); };
  function animZoom(to, cx, cy) {
    to = Math.max(ZMIN, Math.min(ZMAX, to)); cancelAnimationFrame(zAnim);
    if (RM) { setZoom(to, cx, cy); return; }
    const from = pk.zoom, t0 = performance.now();
    const step = t => { const q = Math.min(1, (t - t0) / 170), e = 1 - (1 - q) ** 3; setZoom(from + (to - from) * e, cx, cy); if (q < 1) zAnim = requestAnimationFrame(step); };
    zAnim = requestAnimationFrame(step);
  }
  function startGlide(vx, vy) {
    cancelAnimationFrame(glide); if (RM) return; let last = performance.now();
    const step = t => { const dt = Math.min(32, t - last); last = t; pkVp.scrollLeft -= vx * dt; pkVp.scrollTop -= vy * dt; const f = Math.pow(.94, dt / 16); vx *= f; vy *= f; if (Math.hypot(vx, vy) > .02) glide = requestAnimationFrame(step); };
    glide = requestAnimationFrame(step);
  }
  const effMode = () => spaceHeld ? 'move' : pk.mode;
  function setMode(m) { pk.mode = m; $('pkDraw').setAttribute('aria-pressed', m === 'draw'); $('pkMove').setAttribute('aria-pressed', m === 'move'); pkPad.classList.toggle('move', effMode() === 'move'); }
  {
    const pts = new Map(); let st = null, pan = null, pinch = null, hold = false, tmp = null, lastPt = null, edge = 0;
    const pt = c => { const b = pkPad.getBoundingClientRect(); return { x: Math.max(0, Math.min(1, (c.x - b.left) / b.width)), y: Math.max(0, Math.min(1, (c.y - b.top) / b.height)) }; };
    const two = () => { const [a, b] = [...pts.values()]; return { d: Math.hypot(a.x - b.x, a.y - b.y), x: (a.x + b.x) / 2, y: (a.y + b.y) / 2 }; };
    const showSel = r => { pkSel.hidden = false; Object.assign(pkSel.style, { left: r.x * 100 + '%', top: r.y * 100 + '%', width: r.w * 100 + '%', height: r.h * 100 + '%' }); };
    const upd = p => { tmp = { x: Math.min(st.x, p.x), y: Math.min(st.y, p.y), w: Math.abs(p.x - st.x), h: Math.abs(p.y - st.y) }; showSel(tmp); };
    function edgeLoop() { // عند الاقتراب من حافة العارض أثناء الرسم تتحرك الصفحة تلقائيًا
      if (!st || !lastPt) { edge = 0; return; }
      const v = pkVp.getBoundingClientRect(), M = 36, sp = d => d < M ? (M - d) / M * 14 : 0;
      const dx = sp(lastPt.x - v.left) - sp(v.right - lastPt.x), dy = sp(lastPt.y - v.top) - sp(v.bottom - lastPt.y);
      if (dx || dy) { pkVp.scrollLeft -= dx; pkVp.scrollTop -= dy; upd(pt(lastPt)); }
      edge = requestAnimationFrame(edgeLoop);
    }
    pkPad.addEventListener('pointerdown', e => {
      cancelAnimationFrame(glide); cancelAnimationFrame(zAnim);
      pkPad.setPointerCapture(e.pointerId); pts.set(e.pointerId, { x: e.clientX, y: e.clientY });
      if (pts.size === 2) { if (st) { st = null; cancelAnimationFrame(edge); if (pk.rect) showSel(pk.rect); else pkSel.hidden = true; } pan = null; const t = two(); pinch = { d: t.d, z: pk.zoom, x: t.x, y: t.y }; hold = true; return; }
      if (pts.size > 2 || hold) return;
      e.preventDefault();
      if (effMode() === 'draw' && e.button === 0) { st = pt({ x: e.clientX, y: e.clientY }); lastPt = { x: e.clientX, y: e.clientY }; upd(st); edge = requestAnimationFrame(edgeLoop); }
      else pan = { x: e.clientX, y: e.clientY, l: pkVp.scrollLeft, t: pkVp.scrollTop, h: [{ x: e.clientX, y: e.clientY, t: e.timeStamp }] };
    });
    pkPad.addEventListener('pointermove', e => {
      if (!pts.has(e.pointerId)) return; pts.set(e.pointerId, { x: e.clientX, y: e.clientY });
      if (pinch && pts.size === 2) { frame(() => { if (!pinch || pts.size < 2) return; const t = two(); setZoom(pinch.z * t.d / pinch.d, t.x, t.y); pkVp.scrollLeft -= t.x - pinch.x; pkVp.scrollTop -= t.y - pinch.y; pinch.x = t.x; pinch.y = t.y; }); return; }
      if (st) { lastPt = { x: e.clientX, y: e.clientY }; frame(() => { if (st) upd(pt(lastPt)); }); }
      else if (pan) { const c = { x: e.clientX, y: e.clientY }; pan.h.push({ ...c, t: e.timeStamp }); if (pan.h.length > 6) pan.h.shift(); frame(() => { if (!pan) return; pkVp.scrollLeft = pan.l - (c.x - pan.x); pkVp.scrollTop = pan.t - (c.y - pan.y); }); }
    });
    const end = e => {
      pts.delete(e.pointerId);
      if (pinch) { if (pts.size < 2) pinch = null; if (!pts.size) hold = false; return; }
      if (!pts.size) hold = false;
      if (pan) { const h = pan.h, a = h[0], b = h[h.length - 1], dt = b.t - a.t; pan = null; if (e.type === 'pointerup' && dt > 0 && e.timeStamp - b.t < 80) startGlide((b.x - a.x) / dt, (b.y - a.y) / dt); return; }
      if (!st) return;
      cancelAnimationFrame(edge); upd(pt({ x: e.clientX, y: e.clientY })); const r = tmp; st = null; lastPt = null;
      if (r.w < 0.012 || r.h < 0.006) { if (pk.rect) showSel(pk.rect); else pkSel.hidden = true; return; }
      pk.rect = { ...r, pg: pgFor(pk.k, pk.pg, pk.pages.length) };
      readPicked();
    };
    pkPad.addEventListener('pointerup', end); pkPad.addEventListener('pointercancel', end);
    pkPad.addEventListener('contextmenu', e => e.preventDefault());
  }
  async function readPicked() {
    const tok = ++pk.tok, r = pk.rect, k = pk.k;
    $('pkRo').hidden = false; $('pkRoS').textContent = 'جارٍ قراءة المنطقة…'; $('pkVal').value = ''; $('pkOk').disabled = true; $('pkRaw').textContent = '';
    try {
      const raw = await readRegion(await pageSrc(pk.pg), r); if (tok !== pk.tok) return;
      const v = extract(k, raw);
      $('pkVal').value = blank(v) ? '' : (typeof v === 'number' ? v.toFixed(2) : v);
      $('pkRoS').textContent = blank(v) ? 'لم أجد قيمة واضحة داخل المستطيل. ارسمه مرة أخرى، أو اكتب القيمة هنا:' : 'القيمة المقروءة (يمكنك تصحيحها):';
      $('pkRaw').textContent = raw.trim() ? 'النص في المنطقة: ' + raw.replace(/\s+/g, ' ').trim().slice(0, 140) : '';
    } catch (e) { console.error(e); if (tok !== pk.tok) return; $('pkRoS').textContent = 'تعذّرت قراءة المنطقة. اكتب القيمة هنا:'; }
    $('pkOk').disabled = !$('pkVal').value.trim();
    if (tok === pk.tok) $('pkOk').scrollIntoView({ block: 'nearest' });
  }
  $('pkVal').addEventListener('input', () => { $('pkOk').disabled = !$('pkVal').value.trim() || !pk.rect; });
  $('pkOk').onclick = () => {
    const x = pk.x, k = pk.k, r = pk.rect; if (!x || !r) return;
    let v = $('pkVal').value.trim();
    if (k === 'total' || k === 'tax_amount') v = n2(v);
    if (k === 'vendor_tax_number') v = P.normDigits(v).replace(/\D/g, '');
    if (k === 'invoice_date') v = P.findDate(v) || v;
    if (k === 'payment_method') v = payNorm(v);
    const oldKey = zoneFor(x);
    x[k] = v; (x.edited = x.edited || {})[k] = 1;
    let saved = false;
    if (!$('pkSaveL').hidden && $('pkSave').checked) {
      const key = vendorKey(x) || oldKey;
      if (key) {
        if (oldKey && oldKey !== key && zones[oldKey]) { zones[key] = { ...zones[oldKey], fields: { ...zones[oldKey].fields, ...((zones[key] || {}).fields || {}) } }; delete zones[oldKey]; zMark(oldKey); }
        const t = zones[key] = zones[key] || { fields: {} };
        Object.assign(t, { name: x.vendor_name || t.name || '', vat: vatOf(x) || t.vat || '', nameKey: nameKey(x.vendor_name) || t.nameKey || '', at: new Date().toISOString() });
        t.fields[k] = { pg: r.pg, x: +r.x.toFixed(4), y: +r.y.toFixed(4), w: +r.w.toFixed(4), h: +r.h.toFixed(4) };
        saveZones(); zMark(key); saved = true;
        const n = list.filter(y => y !== x && zoneFor(y) === key).length;
        if (n && !(rrS[x.id] && rrS[x.id].state === 'run')) rrS[x.id] = { state: 'offer', key, ks: [...new Set([...((rrS[x.id] && rrS[x.id].state === 'offer' && rrS[x.id].ks) || []), k])] };
      }
    }
    (x.zsrc = x.zsrc || {})[k] = saved ? 'tpl' : 'pick';
    (x.zrect = x.zrect || {})[k] = { pg: r.pg, x: r.x, y: r.y, w: r.w, h: r.h }; rv.cur[x.id] = k;
    save(); closePick(); draw(); backTo(x.id, k);
    toast(saved ? `حُفظ مكان «${label(k)}» لفواتير ${x.vendor_name || 'المورد'} القادمة` : `تم تحديث «${label(k)}»`);
  };
  $('pkCancel').onclick = $('pkX').onclick = () => { const x = pk.x, k = pk.k; closePick(); if (x) backTo(x.id, k); };
  function backTo(id, k) { // العودة للخانة نفسها في بطاقة الفاتورة
    const el = $(`f-${id}-${k}`); if (!el) return;
    const f = el.closest('.f'); f.classList.remove('flash'); void f.offsetWidth; f.classList.add('flash');
    el.scrollIntoView({ block: 'center', behavior: RM ? 'auto' : 'smooth' });
    setTimeout(() => el.focus({ preventScroll: true }), RM ? 0 : 350);
  }
  $('pkDraw').onclick = () => setMode('draw'); $('pkMove').onclick = () => setMode('move');
  $('pkIn').onclick = () => animZoom(pk.zoom * 1.4); $('pkOut').onclick = () => animZoom(pk.zoom / 1.4); $('pkFit').onclick = () => animZoom(1);
  $('pkPrev').onclick = () => showPg(pk.pg - 1); $('pkNext').onclick = () => showPg(pk.pg + 1);
  pkVp.addEventListener('wheel', e => { if (!e.ctrlKey && !e.metaKey) return; e.preventDefault(); cancelAnimationFrame(zAnim); const z = pk.zoom * Math.exp(-e.deltaY * .0025), x = e.clientX, y = e.clientY; frame(() => setZoom(z, x, y)); }, { passive: false });
  new ResizeObserver(() => { if (!$('pk').hidden) pkPaper.style.width = fitW() * pk.zoom + 'px'; }).observe(pkVp);
  document.addEventListener('keydown', e => {
    if ($('pk').hidden) return;
    if (e.key === 'Escape') { closePick(); return; }
    if (e.target.closest('input')) return;
    if (e.code === 'Space' && !e.target.closest('button')) { e.preventDefault(); if (!spaceHeld) { spaceHeld = true; setMode(pk.mode); } return; }
    if (e.key === 'PageDown') { e.preventDefault(); showPg(pk.pg + 1); }
    else if (e.key === 'PageUp') { e.preventDefault(); showPg(pk.pg - 1); }
    else if (e.target === pkVp) {
      const m = { ArrowLeft: [-60, 0], ArrowRight: [60, 0], ArrowUp: [0, -60], ArrowDown: [0, 60] }[e.key];
      if (m) { e.preventDefault(); pkVp.scrollBy({ left: m[0], top: m[1], behavior: RM ? 'auto' : 'smooth' }); }
      else if (e.key === '+' || e.key === '=') { e.preventDefault(); animZoom(pk.zoom * 1.25); }
      else if (e.key === '-') { e.preventDefault(); animZoom(pk.zoom / 1.25); }
      else if (e.key === '0') animZoom(1);
    }
  });
  document.addEventListener('keyup', e => { if (e.code === 'Space' && spaceHeld) { spaceHeld = false; setMode(pk.mode); } });

  // ---------- مراجعة الفاتورة داخل البطاقة ----------
  // الصور مخفية عند فتح التطبيق، ويمكن إظهارها لكل فاتورة أو للكل
  const rv = { hideAll: true, open: {}, cur: {}, st: {} }, rvPages = {}, rvLive = {};
  const isOpen = id => rv.open[id] ?? !rv.hideAll;
  function rvHtml(x) {
    const o = isOpen(x.id), id = esc(x.id);
    return `<section class="rv${o ? '' : ' off'}" aria-label="مراجعة الفاتورة"><h3><span>مراجعة الفاتورة ${o ? '<small>اضغط أي خانة لترى مكانها على الفاتورة</small>' : ''}</span>
      <button type="button" class="tg" data-rvtg="${id}" aria-expanded="${o}">${o ? 'إخفاء الصورة' : '👁 إظهار صورة الفاتورة'}</button></h3>${o ? `<div class="rvh" id="rvh-${id}"></div>` : ''}</section>`;
  }
  function boxFor(x, k, n) {
    if (!k) return null;
    let r = x.zrect && x.zrect[k];
    if (!r && x.zsrc && x.zsrc[k] && x.zsrc[k] !== 'pick') { const key = zoneFor(x); r = key && zones[key].fields && zones[key].fields[k]; }
    if (!r) return null;
    const pg = pgIndex(r.pg, n); if (pg < 0) return null;
    return { x: r.x, y: r.y, w: r.w, h: r.h, pg, tag: label(k) };
  }
  async function rvLoadPages(x) {
    if (rvPages[x.id]) return rvPages[x.id];
    const rec = await getRec(x.id, x);
    return (rvPages[x.id] = (rec && rec.pages) || []);
  }
  function mountViewers() {
    for (const id in rvLive) { rvLive[id].ro.disconnect(); delete rvLive[id]; }
    list.concat(trash).forEach(x => { const h = document.getElementById('rvh-' + x.id); if (h) rvLive[x.id] = makeViewer(h, x); });
  }
  function makeViewer(host, x) {
    const id = x.id, st = rv.st[id] = rv.st[id] || { pg: 0, zoom: 1, l: 0, t: 0 };
    host.innerHTML = `<div class="rvtb" role="toolbar" aria-label="أدوات عرض الفاتورة">
      <button type="button" data-z="-" aria-label="تصغير">−</button><span class="zl">100%</span><button type="button" data-z="+" aria-label="تكبير">+</button><button type="button" data-z="0">ملاءمة</button>
      <span class="sp"></span><button type="button" data-p="-1" aria-label="الصفحة السابقة">›</button><span class="pg"></span><button type="button" data-p="1" aria-label="الصفحة التالية">‹</button>
      <span class="cap" aria-live="polite"></span></div>
      <div class="rvvp" dir="ltr" tabindex="0" aria-label="صورة الفاتورة. اسحب للتحريك، والأسهم أيضًا"><div class="rvcv"><div class="rvpaper"><img alt="صفحة الفاتورة"><div class="rvpad"></div></div></div></div>`;
    const vp = host.querySelector('.rvvp'), paper = host.querySelector('.rvpaper'), img = paper.querySelector('img'), pad = paper.querySelector('.rvpad');
    const zl = host.querySelector('.zl'), pgl = host.querySelector('.pg'), cap = host.querySelector('.cap'), bPrev = host.querySelector('[data-p="-1"]'), bNext = host.querySelector('[data-p="1"]');
    let pages = [], box = null;
    const fitW = () => { const a = (img.naturalWidth || 1) / (img.naturalHeight || 1.3); return Math.max(160, Math.min(vp.clientWidth - 24, (vp.clientHeight - 24) * a)); };
    function setZoom(z, cx, cy) {
      z = Math.max(1, Math.min(6, z)); const a = paper.getBoundingClientRect();
      if (cx == null) { const v = vp.getBoundingClientRect(); cx = v.left + v.width / 2; cy = v.top + v.height / 2; }
      const fx = a.width ? (cx - a.left) / a.width : .5, fy = a.height ? (cy - a.top) / a.height : .5;
      st.zoom = z; paper.style.width = fitW() * z + 'px';
      const b = paper.getBoundingClientRect(); vp.scrollLeft += b.left + fx * b.width - cx; vp.scrollTop += b.top + fy * b.height - cy;
      zl.textContent = Math.round(z * 100) + '%'; host.querySelector('[data-z="-"]').disabled = z <= 1; host.querySelector('[data-z="+"]').disabled = z >= 6;
    }
    let za = 0, gl = 0;
    const anim = to => { to = Math.max(1, Math.min(6, to)); cancelAnimationFrame(za); if (RM) { setZoom(to); return; } const f = st.zoom, t0 = performance.now();
      const s = t => { const q = Math.min(1, (t - t0) / 170); setZoom(f + (to - f) * (1 - (1 - q) ** 3)); if (q < 1) za = requestAnimationFrame(s); }; za = requestAnimationFrame(s); };
    function drawBox() {
      paper.querySelectorAll('.rvbox').forEach(e => e.remove());
      if (box && box.pg === st.pg) { const d = document.createElement('div'); d.className = 'rvbox';
        Object.assign(d.style, { left: (box.x - .006) * 100 + '%', top: (box.y - .004) * 100 + '%', width: (box.w + .012) * 100 + '%', height: (box.h + .008) * 100 + '%' });
        d.innerHTML = `<span>${esc(box.tag)}</span>`; paper.insertBefore(d, pad); }
    }
    function scrollToBox() {
      if (!box || box.pg !== st.pg) return;
      const b = paper.getBoundingClientRect(), v = vp.getBoundingClientRect();
      vp.scrollTop += b.top + b.height * box.y - v.top - v.height / 3; vp.scrollLeft += b.left + b.width * (box.x + box.w / 2) - v.left - v.width / 2;
    }
    function showPg(i, keep) {
      const n = pages.length; st.pg = Math.max(0, Math.min(n - 1, i));
      pgl.textContent = n > 1 ? `صفحة ${st.pg + 1} من ${n}` : ''; bPrev.hidden = bNext.hidden = n < 2; bPrev.disabled = st.pg === 0; bNext.disabled = st.pg === n - 1;
      return new Promise(ok => { img.onload = () => { paper.style.width = fitW() * st.zoom + 'px'; zl.textContent = Math.round(st.zoom * 100) + '%';
        host.querySelector('[data-z="-"]').disabled = st.zoom <= 1; host.querySelector('[data-z="+"]').disabled = st.zoom >= 6;
        drawBox(); if (keep) { vp.scrollLeft = st.l; vp.scrollTop = st.t; } else { vp.scrollTop = 0; scrollToBox(); } ok(); }; img.onerror = ok;
        if (img.getAttribute('src') === pages[st.pg]) img.onload(); else img.src = pages[st.pg]; });
    }
    const me = {
      focus(k) {
        box = boxFor(x, k, pages.length);
        cap.textContent = !k ? '' : box ? `المكان المظلل: ${label(k)}` : `مكان «${label(k)}» غير معروف بعد. استخدم زر التحديد بجانب الخانة.`;
        if (!pages.length) return;
        if (box && box.pg !== st.pg) showPg(box.pg); else { drawBox(); scrollToBox(); }
      }
    };
    rvLoadPages(x).then(p => {
      if (rvLive[id] !== me) return; pages = p;
      if (!pages.length) { host.querySelector('.rvtb').hidden = true;
        vp.innerHTML = `<div class="rvnone"><p>لا توجد صورة محفوظة لهذه الفاتورة على هذا الجهاز.</p><p>اختر ملفها (${esc(x.file_name || x.file || 'ملف الفاتورة')}) لإرفاقه، دون تغيير البيانات المقروءة:</p>
          <label class="btn primary">📎 إرفاق ملف الفاتورة<input type="file" accept="image/*,application/pdf" data-attach="${esc(id)}"></label></div>`; return; }
      const k = rv.cur[id]; box = boxFor(x, k, pages.length);
      if (k) cap.textContent = box ? `المكان المظلل: ${label(k)}` : '';
      showPg(st.pg, true);
    });
    vp.addEventListener('scroll', () => { st.l = vp.scrollLeft; st.t = vp.scrollTop; }, { passive: true });
    host.querySelector('.rvtb').addEventListener('click', e => { const b = e.target.closest('button'); if (!b) return;
      if (b.dataset.z) anim(b.dataset.z === '+' ? st.zoom * 1.4 : b.dataset.z === '-' ? st.zoom / 1.4 : 1);
      if (b.dataset.p) showPg(st.pg + +b.dataset.p); });
    vp.addEventListener('wheel', e => { if (!e.ctrlKey && !e.metaKey) return; e.preventDefault(); cancelAnimationFrame(za); setZoom(st.zoom * Math.exp(-e.deltaY * .0025), e.clientX, e.clientY); }, { passive: false });
    vp.addEventListener('keydown', e => { const m = { ArrowLeft: [-60, 0], ArrowRight: [60, 0], ArrowUp: [0, -60], ArrowDown: [0, 60] }[e.key];
      if (m) { e.preventDefault(); vp.scrollBy(m[0], m[1]); } else if (e.key === '+' || e.key === '=') { e.preventDefault(); anim(st.zoom * 1.25); }
      else if (e.key === '-') { e.preventDefault(); anim(st.zoom / 1.25); } else if (e.key === '0') anim(1);
      else if (e.key === 'PageDown') { e.preventDefault(); showPg(st.pg + 1); } else if (e.key === 'PageUp') { e.preventDefault(); showPg(st.pg - 1); } });
    me.ro = new ResizeObserver(() => { if (img.naturalWidth) paper.style.width = fitW() * st.zoom + 'px'; }); me.ro.observe(vp);
    // سحب للتحريك، وإصبعان للتكبير، مع انزلاق خفيف بعد الإفلات
    const pts = new Map(); let pan = null, pinch = null, hold = false;
    const two = () => { const [a, b] = [...pts.values()]; return { d: Math.hypot(a.x - b.x, a.y - b.y), x: (a.x + b.x) / 2, y: (a.y + b.y) / 2 }; };
    vp.addEventListener('pointerdown', e => {
      if (e.pointerType === 'mouse' && e.button !== 0) return;
      cancelAnimationFrame(gl); cancelAnimationFrame(za); vp.setPointerCapture(e.pointerId); pts.set(e.pointerId, { x: e.clientX, y: e.clientY });
      if (pts.size === 2) { pan = null; const t = two(); pinch = { d: t.d || 1, z: st.zoom, x: t.x, y: t.y }; hold = true; return; }
      if (pts.size > 2 || hold) return; e.preventDefault();
      pan = { x: e.clientX, y: e.clientY, l: vp.scrollLeft, t: vp.scrollTop, h: [{ x: e.clientX, y: e.clientY, t: e.timeStamp }] };
    });
    vp.addEventListener('pointermove', e => {
      if (!pts.has(e.pointerId)) return; pts.set(e.pointerId, { x: e.clientX, y: e.clientY });
      if (pinch && pts.size === 2) { const t = two(); setZoom(pinch.z * t.d / pinch.d, t.x, t.y); vp.scrollLeft -= t.x - pinch.x; vp.scrollTop -= t.y - pinch.y; pinch.x = t.x; pinch.y = t.y; return; }
      if (pan) { pan.h.push({ x: e.clientX, y: e.clientY, t: e.timeStamp }); if (pan.h.length > 6) pan.h.shift(); vp.scrollLeft = pan.l - (e.clientX - pan.x); vp.scrollTop = pan.t - (e.clientY - pan.y); }
    });
    const end = e => {
      pts.delete(e.pointerId);
      if (pinch) { if (pts.size < 2) pinch = null; if (!pts.size) hold = false; return; }
      if (!pts.size) hold = false;
      if (!pan) return; const h = pan.h, a = h[0], b = h[h.length - 1], dt = b.t - a.t; pan = null;
      if (RM || e.type !== 'pointerup' || !(dt > 0) || e.timeStamp - b.t > 80) return;
      let vx = (b.x - a.x) / dt, vy = (b.y - a.y) / dt, l = performance.now();
      const s = t => { const d = Math.min(32, t - l); l = t; vp.scrollLeft -= vx * d; vp.scrollTop -= vy * d; const f = Math.pow(.94, d / 16); vx *= f; vy *= f; if (Math.hypot(vx, vy) > .02) gl = requestAnimationFrame(s); };
      gl = requestAnimationFrame(s);
    };
    vp.addEventListener('pointerup', end); vp.addEventListener('pointercancel', end);
    return me;
  }
  // الضغط على خانة يُظهر مكانها على الصورة
  document.addEventListener('focusin', e => {
    const el = e.target.closest('.inv [data-k][data-id]'); if (!el) return;
    const id = el.dataset.id, k = el.dataset.k; if (!F.some(f => f[0] === k)) return;
    const x = list.find(y => y.id === id); if (!x || !isOpen(id)) return;
    const c = $('c-' + id); if (c) c.querySelectorAll('.f.cur').forEach(f => f.classList.remove('cur'));
    el.closest('.f').classList.add('cur'); rv.cur[id] = k;
    if (rvLive[id]) rvLive[id].focus(k);
  });
  $('rvAll').textContent = '👁 إظهار صور كل الفواتير';
  const rvAllLabel = () => { const b = $('rvAll'); b.textContent = rv.hideAll ? '👁 إظهار صور كل الفواتير' : 'إخفاء صور كل الفواتير'; b.setAttribute('aria-pressed', !rv.hideAll); b.hidden = !list.length; };

  // ---------- إعادة قراءة فواتير المورد نفسه بعد تعديل القالب ----------
  const rrS = {};
  const shown = (k, v) => blank(v) ? 'فارغ' : (k === 'total' || k === 'tax_amount') ? money(v) : String(v);
  function rrHtml(x) {
    const r = rrS[x.id]; if (!r) return '';
    const t = zones[r.key], vn = esc((t && t.name) || x.vendor_name || 'المورد'), id = esc(x.id);
    const others = list.filter(y => y !== x && zoneFor(y) === r.key);
    if (r.state === 'offer') {
      if (!others.length) return '';
      return `<div class="rr" role="status"><b>حُفظ مكان ${r.ks.map(k => '«' + esc(label(k)) + '»').join(' و')} في قالب ${vn}</b>
        <p>في القائمة ${others.length === 1 ? 'فاتورة أخرى' : others.length + ' فواتير أخرى'} من نفس المورد. يمكن إعادة قراءتها الآن بالأماكن المحفوظة. الخانات التي عدّلتها بيدك لن تتغير.</p>
        <div class="row"><button type="button" class="primary" data-rr="${id}">أعد قراءة ${others.length === 1 ? 'الفاتورة' : others.length + ' فواتير'} وطبّق التعديل</button><button type="button" data-rrno="${id}">لاحقًا</button></div></div>`;
    }
    if (r.state === 'run') return `<div class="rr" role="status"><b>جارٍ إعادة قراءة فواتير ${vn}… ${r.done} من ${r.n}</b><div class="bar2"><i style="width:${r.n ? r.done / r.n * 100 : 0}%"></i></div></div>`;
    const fn = cid => { const y = list.find(q => q.id === cid); return esc((y && (y.invoice_number || y.file)) || 'محذوفة'); };
    const nInv = new Set(r.changes.map(c => c.id)).size;
    return `<div class="rr" role="status"><b>${r.changes.length ? `انتهت إعادة القراءة: تغيّرت ${r.changes.length} خانة في ${nInv} فاتورة` : 'انتهت إعادة القراءة: لم تتغير أي خانة'}</b>
      ${r.changes.length || r.skipped.length ? `<ul>${r.changes.map(c => `<li><a href="#c-${esc(c.id)}">فاتورة ${fn(c.id)}</a>: ${esc(label(c.k))} <s>${esc(shown(c.k, c.from))}</s> ← <ins>${esc(shown(c.k, c.to))}</ins></li>`).join('')}
        ${r.skipped.map(c => `<li class="sk">فاتورة ${fn(c.id)}: ${esc(label(c.k))} ${c.why}</li>`).join('')}</ul>` : ''}
      <div class="row">${r.changes.length ? `<button type="button" data-rrundo="${id}">↩ تراجع عن التطبيق</button>` : ''}<button type="button" data-rrno="${id}">إغلاق</button></div></div>`;
  }
  async function rerunVendor(id) {
    const x = list.find(y => y.id === id), r0 = rrS[id]; if (!x || !r0 || r0.state !== 'offer') return;
    const key = r0.key, T = zones[key]; if (!T) return;
    const others = list.filter(y => y !== x && zoneFor(y) === key);
    const r = rrS[id] = { state: 'run', key, n: others.length, done: 0, changes: [], skipped: [], before: [] }; draw();
    for (const y of others) {
      r.before.push({ id: y.id, vals: Object.fromEntries(F.map(([k]) => [k, y[k]])), zsrc: { ...(y.zsrc || {}) }, zrect: JSON.parse(JSON.stringify(y.zrect || {})) });
      try {
        const rec = await getRec(y.id, y);
        if (!rec || !rec.pages || !rec.pages.length) r.skipped.push({ id: y.id, k: Object.keys(T.fields)[0], why: 'لا توجد صورة محفوظة لها' });
        else {
          const c = { rec, pages: rec.pages, src: [], pdf: null };
          for (const [k, z] of Object.entries(T.fields || {})) {
            if (y.edited && y.edited[k]) { r.skipped.push({ id: y.id, k, why: 'لم يتغير لأنك عدّلته بيدك' }); continue; }
            const i = pgIndex(z.pg, c.pages.length); if (i < 0) continue;
            const v = extract(k, await readRegion(await loadSrc(c, i), z));
            if (blank(v)) { r.skipped.push({ id: y.id, k, why: 'لم تُقرأ قيمة واضحة في المكان المحفوظ، فبقيت كما هي' }); continue; }
            (y.zrect = y.zrect || {})[k] = { ...z };
            if (String(v) !== String(y[k] ?? '')) { r.changes.push({ id: y.id, k, from: y[k], to: v }); y[k] = v; (y.zsrc = y.zsrc || {})[k] = 'new'; }
            else (y.zsrc = y.zsrc || {})[k] = (y.zsrc && y.zsrc[k] === 'new') ? 'new' : 'tpl';
          }
          if (c.pdf) try { c.pdf.destroy(); } catch { }
        }
      } catch (e) { console.warn(e); r.skipped.push({ id: y.id, k: Object.keys(T.fields)[0], why: 'تعذّرت قراءتها' }); }
      r.done++; save(); if (rrS[id] === r) draw();
    }
    r.state = 'done'; save(); draw();
    toast(r.changes.length ? `تحدّثت ${r.changes.length} خانة بالقالب` : 'انتهت إعادة القراءة ولم تتغير البيانات');
  }
  function undoRerun(id) {
    const r = rrS[id]; if (!r || r.state !== 'done') return;
    r.before.forEach(b => { const y = list.find(q => q.id === b.id); if (!y) return; F.forEach(([k]) => { y[k] = b.vals[k]; }); y.zsrc = b.zsrc; y.zrect = b.zrect; });
    delete rrS[id]; save(); draw(); toast('أُعيدت الفواتير كما كانت قبل التطبيق');
  }

  // ---------- مزامنة قوالب الموردين مع Google Sheet (ورقة VendorTemplates) ----------
  // رابط الشيت من config.js، والبريد والرمز يُدخلهما المستخدم مرة واحدة (نفس الربط في النسخة الكاملة)
  const GKEY = 'invoice-reader-gsheet-v1', ZPK = 'invoice-simple-zones-pending-v1';
  const GURL = String((window.APP_CONFIG || {}).gsheetUrl || '').trim();
  let gs = null; try { gs = JSON.parse(localStorage.getItem(GKEY)); } catch { }
  if (!gs || !gs.email || !gs.token || !GURL) gs = null;
  let zpend = {}; try { zpend = JSON.parse(localStorage.getItem(ZPK)) || {}; } catch { }
  const zst = { busy: false, err: '', at: '' };
  let zTimer = 0;
  function zMark(k) { if (!k) return; zpend[k] = 1; try { localStorage.setItem(ZPK, JSON.stringify(zpend)); } catch { } clearTimeout(zTimer); if (gs) zTimer = setTimeout(zSync, 8000); drawSync(); }
  async function gsCall(cfg, body) {
    const r = body ? await fetch(GURL, { method: 'POST', body: JSON.stringify({ email: cfg.email, token: cfg.token, ...body }), redirect: 'follow' })
      : await fetch(GURL + (GURL.includes('?') ? '&' : '?') + 'action=zones&email=' + encodeURIComponent(cfg.email) + '&token=' + encodeURIComponent(cfg.token), { redirect: 'follow' });
    let j; try { j = await r.json(); } catch { throw new Error('رد غير متوقع من الشيت. تأكد أن Apps Script حُدّث بالنسخة الجديدة'); }
    if (!j.ok) throw new Error(j.error || 'تعذّر الاتصال بالشيت');
    if (!j.zones) throw new Error('الشيت لا يدعم القوالب بعد: حدّث كود Apps Script وانشر نسخة جديدة');
    return j;
  }
  function zMerge(j) { // يدمج القوالب: الأحدث يفوز، والتعديلات المحلية غير المرسلة لها الأولوية
    const put = {}, R = j.zones || {}, Dl = j.deleted || {};
    for (const k in Dl) if (zones[k] && !zpend[k] && (zones[k].at || '') <= (Dl[k] || '')) delete zones[k];
    for (const k in R) { if (zpend[k]) continue; if (!zones[k] || (R[k].at || '') > (zones[k].at || '')) zones[k] = R[k]; }
    for (const k in zones) if (!R[k] || zpend[k]) put[k] = zones[k];
    for (const k in zpend) if (!(k in put)) put[k] = zones[k] || null;
    return put;
  }
  async function zSync() {
    if (!gs || zst.busy) return; clearTimeout(zTimer);
    zst.busy = true; zst.err = ''; drawSync();
    try {
      const sent = { ...zpend };
      let j = await gsCall(gs);
      const put = zMerge(j);
      if (Object.keys(put).length) j = await gsCall(gs, { action: 'zones', put });
      for (const k in sent) delete zpend[k];
      try { localStorage.setItem(ZPK, JSON.stringify(zpend)); } catch { }
      saveZones(); zst.at = new Date().toLocaleTimeString('ar-SA-u-nu-latn', { hour: 'numeric', minute: '2-digit' });
    } catch (e) { console.warn(e); zst.err = (e && e.message) || 'تعذّر الاتصال بالشيت'; }
    zst.busy = false; draw();
    if (Object.keys(zpend).length && !zst.err) { clearTimeout(zTimer); zTimer = setTimeout(zSync, 3000); }
  }
  function drawSync() {
    const b = $('zSync'); if (!b) return;
    if (!GURL) { b.innerHTML = ''; return; }
    const n = Object.keys(zpend).length;
    if (!gs) {
      if (b.querySelector('form')) return; // لا نمسح ما يكتبه المستخدم
      b.innerHTML = `<p><b>القوالب محفوظة في هذا المتصفح فقط.</b> اربط التطبيق بـ Google Sheet لتُحفظ القوالب دائمًا وتظهر على كل أجهزتك، ولتتمكن من إرسال الفواتير إلى الشيت.</p>
        <form id="zsForm" class="zsform"><input id="zsEmail" type="email" autocomplete="email" placeholder="البريد الإلكتروني" required><input id="zsToken" autocomplete="off" placeholder="رمز الدخول" required><button class="primary" id="zsGo">ربط بـ Google Sheet</button></form>
        <small>البريد والرمز من قائمة «قارئ الفواتير» ← «إضافة مستخدم مسموح له» في الشيت.</small>`;
      return;
    }
    b.innerHTML = `<p class="zson">☁️ <b>القوالب محفوظة في Google Sheet</b> باسم ${esc(gs.email)}${zst.busy ? ' · جارٍ المزامنة…' : zst.err ? '' : zst.at ? ' · آخر مزامنة ' + esc(zst.at) : ''}${n && !zst.busy ? ` · ${n} تعديل بانتظار الإرسال` : ''}</p>
      ${zst.err ? `<p class="zserr" role="alert">تعذّرت المزامنة: ${esc(zst.err)}. القوالب باقية في هذا المتصفح وستُرسل لاحقًا.</p>` : ''}
      <div class="zbtns"><button type="button" id="zsNow"${zst.busy ? ' disabled' : ''}>🔄 مزامنة الآن</button><button type="button" id="zsOff">إلغاء الربط</button></div>`;
  }
  document.addEventListener('submit', async e => {
    if (e.target.id !== 'zsForm') return; e.preventDefault();
    const email = $('zsEmail').value.trim().toLowerCase(), token = $('zsToken').value.trim();
    if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email) || !token) { toast('اكتب البريد ورمز الدخول'); return; }
    $('zsGo').disabled = true; $('zsGo').textContent = 'جارٍ الربط…';
    try { await gsCall({ email, token }); gs = { url: GURL, email, token }; try { localStorage.setItem(GKEY, JSON.stringify(gs)); } catch { }
      $('zSync').innerHTML = ''; await zSync(); toast('تم الربط: القوالب محفوظة في Google Sheet'); }
    catch (er) { toast((er && er.message) || 'تعذّر الربط'); const g = $('zsGo'); if (g) { g.disabled = false; g.textContent = 'ربط بـ Google Sheet'; } }
  });
  document.addEventListener('click', e => {
    const b = e.target.closest('button'); if (!b) return;
    if (b.id === 'zsNow') zSync();
    if (b.id === 'zsOff') { const old = gs; gs = null; try { localStorage.setItem(GKEY, JSON.stringify({ url: GURL })); } catch { } drawSync();
      undoBar('أُلغي الربط مع Google Sheet. القوالب باقية في هذا المتصفح', () => { gs = old; try { localStorage.setItem(GKEY, JSON.stringify(old)); } catch { } drawSync(); }, () => { }); }
  });
  if (gs) setTimeout(zSync, 1500);

  // ---------- إرسال الفواتير إلى Google Sheet (ورقة «فواتير التطبيق») ----------
  const sentSig = x => JSON.stringify([x.invoice_number, x.invoice_date, x.total, x.tax_amount, x.vendor_name, x.vendor_tax_number, x.payment_method, x.project || '', noteOf(x), x.file || '', !!x.deletedAt, x.delReason || '']);
  const sendState = x => !x.sentAt ? (x.sendErr ? 'err' : 'new') : x.sentSig === sentSig(x) ? 'ok' : 'dirty';
  function sentBadge(x) {
    if (!GURL) return '';
    const st = sendState(x);
    if (st === 'ok') return `<small class="sb ok">✓ في Google Sheet${x.sentRow ? ' (صف ' + x.sentRow + ')' : ''}</small>`;
    if (st === 'dirty') return '<small class="sb dirty">⟳ عُدّلت بعد الإرسال، ستُحدَّث في الشيت عند الإرسال</small>';
    if (st === 'err') return `<small class="sb err">لم تُرسل: ${esc(x.sendErr)}</small>`;
    return '';
  }
  const toSend = () => list.filter(x => sendState(x) !== 'ok').concat(trash.filter(x => x.sentAt && sendState(x) === 'dirty'));
  let sending = false;
  function sendLabel() {
    const b = $('gsSend'); if (!b) return;
    b.hidden = !GURL || (!list.length && !trash.some(x => x.sentAt));
    const n = toSend().length;
    b.disabled = sending;
    b.textContent = sending ? 'جارٍ الإرسال والنقل إلى InvoiceT…' : n ? `☁️ إرسال ونقل إلى InvoiceT (${n})` : '☁️ نقل إلى InvoiceT';
  }
  $('gsSend').onclick = async () => {
    if (!gs) { toast('اربط التطبيق بـ Google Sheet أولًا: اكتب بريدك ورمز الدخول'); showTab('z'); const f = $('zsEmail'); if (f) { f.scrollIntoView({ block: 'center', behavior: RM ? 'auto' : 'smooth' }); setTimeout(() => f.focus({ preventScroll: true }), 300); } return; }
    if (sending) return; const items = toSend();
    sending = true; sendLabel();
    const sig = new Map(items.map(x => [x.id, sentSig(x)]));
    try {
      const n2s = v => { const t = n2(v); return typeof t === 'number' ? t : (t || ''); };
      const invoices = items.map(x => ({ id: x.id, InvoiceNo: String(x.invoice_number || ''), Date: x.invoice_date || '', Total: n2s(x.total), Vat: n2s(x.tax_amount), Vendor: x.vendor_name || '',
        VATID: x.vendor_tax_number || '', Payment: x.payment_method || '', Project: x.project || '', Note: noteOf(x), File: x.file || '', Deleted: !!x.deletedAt, DeleteReason: x.delReason || '' }));
      const r = await fetch(GURL, { method: 'POST', body: JSON.stringify(items.length ? { email: gs.email, token: gs.token, action: 'appInvoices', invoices, copy: true } : { email: gs.email, token: gs.token, action: 'copyToInvoiceT' }), redirect: 'follow' });
      let j; try { j = await r.json(); } catch { throw new Error('رد غير متوقع من الشيت'); }
      if (!j.ok) throw new Error(j.error || 'تعذّر الإرسال');
      if (!j.results) throw new Error('الشيت لا يستقبل الفواتير بعد: حدّث كود Apps Script وانشر نسخة جديدة');
      let ok = 0, sk = 0; const now = new Date().toISOString(), all = list.concat(trash);
      j.results.forEach(o => { const x = all.find(y => y.id === o.id); if (!x) return;
        if (o.skipped) { x.sendErr = o.skipped; sk++; return; }
        x.sentAt = now; x.sentRow = o.row; x.sentSig = sig.get(x.id); delete x.sendErr; ok++; });
      save(); saveTrash();
      const head = items.length ? (sk ? `أُرسلت ${ok} فاتورة إلى «فواتير التطبيق»، ولم تُرسل ${sk} لأنها مكررة في الشيت.` : `أُرسلت ${ok} فاتورة إلى «فواتير التطبيق».`) : '';
      if (!j.copy) showRep(head + '\n\nلم يتم النقل إلى InvoiceT: حدّث كود Apps Script في الشيت وانشر نسخة جديدة (New version).', true);
      else if (j.copy.error) showRep(head + '\n\nتعذّر النقل إلى InvoiceT: ' + j.copy.error, true);
      else showRep((head ? head + '\n\n' : '') + j.copy.msg, false);
    } catch (e) { console.warn(e); showRep('تعذّر الإرسال: ' + ((e && e.message) || 'تحقق من الاتصال'), true); }
    sending = false; draw();
  };

  // تقرير الإرسال والنقل (يبقى ظاهرًا حتى يُغلق)
  function showRep(t, bad) {
    const b = $('gsRep'); if (!b) { toast(t); return; }
    b.hidden = false; b.classList.toggle('bad', !!bad);
    b.innerHTML = `<div class="gsrep-h"><b>${bad ? 'نتيجة الإرسال' : 'تم الإرسال والنقل إلى InvoiceT'}</b><button type="button" class="ghost" id="gsRepX" aria-label="إغلاق">✕ إغلاق</button></div><div class="gsrep-b">${esc(t)}</div>`;
    $('gsRepX').onclick = () => { b.hidden = true; };
    b.scrollIntoView({ block: 'nearest', behavior: RM ? 'auto' : 'smooth' });
  }

  // إرفاق ملف لفاتورة موجودة ليس لها صورة محفوظة (لا يعيد القراءة ولا يغير البيانات)
  document.addEventListener('change', async e => {
    const inp = e.target; if (!inp.dataset || !inp.dataset.attach) return;
    const id = inp.dataset.attach, f = inp.files && inp.files[0]; inp.value = ''; if (!f) return;
    const x = list.find(y => y.id === id) || trash.find(y => y.id === id); if (!x) return;
    toast('جارٍ تجهيز صفحات الفاتورة…');
    try {
      let pg;
      if (f.type === 'application/pdf' || /\.pdf$/i.test(f.name || '')) pg = (await pdfPages(f)).pages.map(p => p.canvas);
      else { const u = URL.createObjectURL(f); try { pg = [toCanvas(await loadImg(u), 2400)]; } finally { URL.revokeObjectURL(u); } }
      const ok = await IMG.put(id, { pages: pg.map(toUrl), blob: f.size < 25e6 ? f : null });
      delete rvPages[id]; draw(); if (ok) toast('أُرفق ملف الفاتورة');
    } catch (er) { console.error(er); toast('تعذّر فتح هذا الملف'); }
  });

  // ---------- التبويبات: الفواتير / قوالب الموردين ----------
  function showTab(t) {
    document.body.classList.toggle('tab-z', t === 'z');
    $('tabInv').setAttribute('aria-selected', t !== 'z'); $('tabZ').setAttribute('aria-selected', t === 'z');
    if (location.hash !== (t === 'z' ? '#templates' : '')) history.replaceState(null, '', t === 'z' ? '#templates' : location.pathname + location.search);
  }
  $('tabInv').onclick = () => { showTab('i'); window.scrollTo(0, 0); };
  $('tabZ').onclick = () => { showTab('z'); window.scrollTo(0, 0); };
  if (location.hash === '#templates') showTab('z');

  // إزالة أي Service Worker قديم حتى تظهر آخر نسخة دائمًا
  if ('serviceWorker' in navigator) navigator.serviceWorker.getRegistrations().then(rs => rs.forEach(r => r.unregister())).catch(() => { });
  draw();
  window.__invoiceApp = { addFiles, zones: () => zones, extract, get busy() { return busy || queue.length > 0; }, get list() { return list; } };
})();
