/* قارئ الفواتير: الواجهة والتشغيل. كل المعالجة داخل المتصفح. */
(function () {
  'use strict';
  const P = window.InvoiceParse;
  const $ = id => document.getElementById(id);
  const esc = s => String(s ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const abs = p => new URL(p, location.href).href;
  // الحالة: قائمة الفواتير (محفوظة في المتصفح) + صور الجلسة الحالية في الذاكرة فقط
  let busy = false, pages = [], pageIdx = 0, cur = null, qr = null, textOnly = null, rawText = '';
  const queue = [], imgs = {}, live = {}; // live[id] = {status, err}

  if (window.pdfjsLib) pdfjsLib.GlobalWorkerOptions.workerSrc = abs('vendor/pdf.worker.min.js');

  const FIELDS = [
    ['vendor_name', 'المورّد / البائع'], ['vendor_tax_number', 'الرقم الضريبي للمورّد'],
    ['invoice_number', 'رقم الفاتورة'], ['invoice_date', 'تاريخ الفاتورة'],
    ['due_date', 'تاريخ الاستحقاق'], ['customer_name', 'العميل / المشتري'],
    ['currency', 'العملة'], ['payment_method', 'طريقة الدفع']];
  const SUMS = [['subtotal', 'المجموع قبل الضريبة'], ['discount', 'الخصم'], ['tax_rate', 'نسبة الضريبة %'], ['tax_amount', 'مبلغ الضريبة'], ['total', 'الإجمالي']];
  const SRC = { qr: 'QR', text: 'نص', calc: 'محسوب' };

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
  async function pdfPages(file) {
    if (!window.pdfjsLib) throw new Error('تعذّر تحميل قارئ PDF.');
    const doc = await pdfjsLib.getDocument({ data: await file.arrayBuffer(), isEvalSupported: false }).promise;
    const n = Math.min(doc.numPages, 5), out = [];
    for (let p = 1; p <= n; p++) {
      const pg = await doc.getPage(p), v0 = pg.getViewport({ scale: 1 });
      const v = pg.getViewport({ scale: Math.min(3, 2200 / Math.max(v0.width, v0.height)) });
      const c = document.createElement('canvas'); c.width = Math.round(v.width); c.height = Math.round(v.height);
      const x = c.getContext('2d'); x.fillStyle = '#fff'; x.fillRect(0, 0, c.width, c.height);
      await pg.render({ canvasContext: x, viewport: v }).promise;
      // نص PDF الرقمي (أدق من قراءة الصورة إن وُجد)
      let text = '';
      try {
        const tc = await pg.getTextContent(), rows = [];
        tc.items.forEach(it => {
          if (!it.str || !it.str.trim()) return;
          const y = it.transform[5], xx = it.transform[4];
          let row = rows.find(r => Math.abs(r.y - y) < Math.max(3, Math.abs(it.transform[3]) * 0.5));
          if (!row) rows.push(row = { y, parts: [] });
          row.parts.push({ x: xx, s: it.str });
        });
        rows.sort((a, b) => b.y - a.y);
        text = rows.map(r => {
          const rtl = /[؀-ۿ]/.test(r.parts.map(p => p.s).join(''));
          r.parts.sort((a, b) => rtl ? b.x - a.x : a.x - b.x);
          return r.parts.map(p => p.s.trim()).join('  ');
        }).join('\n');
      } catch (e) { }
      out.push({ canvas: c, text });
    }
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
  let ocrBase = 0, ocrSpan = 1;
  async function getWorker(lang) {
    if (workers[lang]) return workers[lang];
    status(lang === 'ara+eng' ? 'جارٍ تجهيز قارئ النص العربي والإنجليزي…' : 'جارٍ تجهيز قارئ الأرقام الهندية… (مرة واحدة فقط)');
    workers[lang] = Tesseract.createWorker(lang, 1, {
      workerPath: abs('vendor/worker.min.js'), workerBlobURL: false, corePath: abs('vendor/core'), langPath: abs('vendor/lang'), gzip: true,
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

  // الدفعات الكبيرة (أكثر من 10 ملفات): عدد محاولات أقل حتى لا تطول القراءة كثيرًا
  let fastMode = false;
  // إغلاق قارئ النص بعد دقيقتين بلا استخدام لتحرير ذاكرة الجوال (يُعاد تحميله من الذاكرة المؤقتة عند الحاجة)
  let idleT = null;
  function idleWorkers() {
    clearTimeout(idleT);
    idleT = setTimeout(() => {
      if (busy) return;
      Object.keys(workers).forEach(k => { const p = workers[k]; delete workers[k]; Promise.resolve(p).then(w => w && w.terminate()).catch(() => { }); });
    }, 120000);
  }
  async function ocrPage(canvas, q, say) {
    const runs = [], PS = fastMode ? PASSES.slice(0, 3) : PASSES, ES = fastMode ? EAST_PASSES.slice(0, 1) : EAST_PASSES;
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
    for (let k = 0; k < PS.length; k++) {
      if (k) say(`محاولة قراءة ${k + 1} لتحسين النتيجة…`);
      ocrBase = 0.1 + 0.8 * k / (PS.length + ES.length); ocrSpan = 0.8 / (PS.length + ES.length);
      const r = await run('ara+eng', PS[k]);
      if (good(r) || (k >= 1 && P.itemsOk(r))) break;
      if (eastern(runs[0].t) && k >= 1) break; // للأرقام الهندية ننتقل للنموذج المتخصص
    }
    if (runs.some(x => eastern(x.t)) && !runs.some(x => good(x.r))) {
      for (let k = 0; k < ES.length; k++) {
        say('وُجدت أرقام هندية، جارٍ قراءتها بنموذج متخصص…');
        ocrBase = 0.1 + 0.8 * (PS.length + k) / (PS.length + ES.length);
        const r = await run('Arabic+eng', ES[k]);
        if (good(r)) break;
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
      try { list = [{ canvas: toCanvas(await loadImg(url), 2400), text: '' }]; }
      catch (e) { if (/\.hei[cf]$/i.test(file.name || '') || /hei[cf]/i.test(file.type || '')) throw new Error('صيغة HEIC لا يفتحها هذا المتصفح. حوّل الصورة إلى JPG (أو اضبط كاميرا الآيفون على «الأكثر توافقًا») ثم ارفعها.'); throw e; }
      finally { URL.revokeObjectURL(url); }
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
    return { d, pages: await Promise.all(list.map(p => toBlob(p.canvas))) };
  }
  function toBlob(c) { // صورة للعرض فقط، بحجم أصغر لتوفير الذاكرة. تُحفظ Blob (أصغر بالثلث من نص base64)
    const r = Math.min(1, 1600 / Math.max(c.width, c.height));
    let s = c;
    if (r < 1) { s = document.createElement('canvas'); s.width = Math.round(c.width * r); s.height = Math.round(c.height * r); s.getContext('2d').drawImage(c, 0, 0, s.width, s.height); }
    return new Promise(ok => s.toBlob(b => ok(b || s.toDataURL('image/jpeg', 0.82)), 'image/jpeg', 0.82));
  }
  // الصورة المحفوظة قد تكون Blob (الجديدة) أو نص data URL (القديمة)
  const srcOf = v => typeof v === 'string' ? v : URL.createObjectURL(v);

  // ---------- طابور الدفعة ----------
  const isInv = f => /^image\/|pdf/.test(f.type) || /\.(pdf|jpe?g|png|webp|heic)$/i.test(f.name || '');
  const pathOf = f => f.webkitRelativePath || f._path || f.name || '';
  // بصمة الملف: الاسم + الحجم + تاريخ التعديل (لمنع اختيار نفس الملف أكثر من مرة)
  const sigOf = f => [String(f.name || '').toLowerCase(), f.size || 0, f.lastModified || 0].join('|');
  const isCam = f => !f.webkitRelativePath && !f._path && /^image\.(jpe?g|png)$/i.test(f.name || '');
  async function hashOf(f) {
    try { const b = await crypto.subtle.digest('SHA-256', await f.arrayBuffer()); return [...new Uint8Array(b)].map(x => x.toString(16).padStart(2, '0')).join(''); } catch { return ''; }
  }
  // آخر عملية تحميل (تُحفظ في المتصفح)
  const LKEY = 'invoice-reader-lastrun-v1';
  let run = null; try { run = JSON.parse(localStorage.getItem(LKEY)) || null; } catch { run = null; }
  const saveRun = () => { try { localStorage.setItem(LKEY, JSON.stringify(run)); } catch { } };
  const runAdd = (file, st, msg, id) => { run.items.push({ file, st, msg: msg || '', id: id || '' }); saveRun(); };
  // صور الفواتير تُحفظ في IndexedDB داخل المتصفح (مجانًا) لعرضها لاحقًا في المراجعة والتصدير
  const IDB = (() => {
    let dbp = null;
    const db = () => dbp || (dbp = new Promise((ok, no) => { try { const r = indexedDB.open('invoice-reader-img', 1); r.onupgradeneeded = () => r.result.createObjectStore('img'); r.onsuccess = () => ok(r.result); r.onerror = () => no(r.error); } catch (e) { no(e); } }));
    const tx = async (mode, fn) => { const d = await db(); return new Promise((ok, no) => { const t = d.transaction('img', mode), st = t.objectStore('img'); const r = fn(st); t.oncomplete = () => ok(r && r.result); t.onerror = () => no(t.error); }); };
    return {
      put: (id, v) => tx('readwrite', st => st.put(v, id)).catch(() => { }),
      get: id => tx('readonly', st => st.get(id)).catch(() => null),
      del: id => tx('readwrite', st => st.delete(id)).catch(() => { }),
      keys: () => tx('readonly', st => st.getAllKeys()).catch(() => [])
    };
  })();
  async function pagesOf(id) { if (imgs[id] && imgs[id].length) return imgs[id]; const u = await IDB.get(id); if (u && u.length) { imgs[id] = u.map(srcOf); return imgs[id]; } return []; }
  function addFiles(files, fromFolder) {
    const allF = [...(files || [])];
    if (!busy && !queue.length) { run = { at: new Date().toISOString(), items: [] }; saveRun(); }
    files = allF.filter(f => isInv(f) && !/^\./.test(f.name || ''));
    // منع التكرار: ملف سبق اختياره (في الملخص أو الدفعة الحالية) أو مكرر في نفس الاختيار
    const nsOf = f => String(f.name || '').toLowerCase() + '|' + (f.size || 0);
    // المسار الكامل يُقارن فقط لملفات المجلدات (فيه «/»)؛ الاسم وحده لا يكفي، فقد تحمل فاتورتان مختلفتان نفس الاسم مثل Invoice.pdf
    const pKey = p => /[\\/]/.test(p || '') ? 'p:' + String(p).toLowerCase() : '';
    const seen = new Set([...all().flatMap(x => [x._sig, x._ns, pKey(x.file)]).filter(Boolean),
      ...queue.flatMap(j => [j.sig, j.ns]), ...Object.values(live).flatMap(v => [v.sig, v.ns]).filter(Boolean)]);
    const dupF = [];
    files = files.filter(f => {
      const keys = [sigOf(f), nsOf(f), pKey(pathOf(f))].filter(Boolean);
      if (!isCam(f) && keys.some(k => seen.has(k))) { dupF.push(pathOf(f)); return false; }
      keys.forEach(k => seen.add(k)); return true;
    });
    const notInv = allF.length - files.length - dupF.length;
    const dupTxt = dupF.length ? `تم تجاهل ${dupF.length} ملف سبق رفعه (${dupF.slice(0, 3).join('، ')}${dupF.length > 3 ? '…' : ''})` : '';
    if (fromFolder) files.sort((a, b) => pathOf(a).localeCompare(pathOf(b), 'ar', { numeric: true }));
    if (!files.length) {
      dupF.forEach(f => runAdd(f, 'dup', 'الملف سبق رفعه')); drawList();
      if (dupF.length) return toast(fromFolder ? `لا توجد ملفات جديدة في المجلد. ${dupTxt}` : dupTxt);
      if (allF.length || fromFolder) toast(fromFolder ? 'لا توجد صور أو ملفات PDF في هذا المجلد' : 'الملف ليس صورة أو PDF'); return;
    }
    if (fromFolder) setTimeout(() => toast(`وُجدت ${files.length} فاتورة جديدة في المجلد` + (dupF.length ? `، و${dupTxt}` : '') + (notInv ? ` (وتُجوهل ${notInv} ملف ليس صورة أو PDF)` : '')), 50);
    else if (dupF.length) setTimeout(() => toast(dupTxt), 50);
    dupF.forEach(f => runAdd(f, 'dup', 'الملف سبق رفعه'));
    files.forEach((f, i) => { const id = Date.now() + '-' + i + '-' + Math.random().toString(36).slice(2, 6); const sig = sigOf(f), ns = nsOf(f); queue.push({ id, file: f, sig, ns }); live[id] = { status: 'queued', file: pathOf(f) || 'صورة من الكاميرا', sig, ns }; });
    drawList(); if (files.length > 1 && !fromFolder) toast(`أُضيفت ${files.length} فواتير إلى الدفعة`);
    pump();
  }
  let setProg = () => { };
  async function sheetHashes() { // بصمات الملفات المرسلة إلى الشيت من كل الأجهزة والمستخدمين
    if (!gs) return (ref && ref.hashes) || {};
    try { const j = await gsCall(gs, 'hashes', 10000); if (j.ok && j.hashes) { if (ref) { ref.hashes = j.hashes; saveRef(); } return j.hashes; } } catch { }
    return (ref && ref.hashes) || {};
  }
  async function pump() {
    if (busy) return; busy = true; clearTimeout(idleT);
    await ready;
    status('جارٍ التحقق من الملفات المرسلة سابقًا…'); const sh = await sheetHashes();
    let done = 0, perr = ''; const batch = [];
    try { // أي خطأ غير متوقع لا يترك التطبيق عالقًا في «جارٍ القراءة»
    const total = () => done + 1 + queue.length;
    while (queue.length) {
      const job = queue.shift(); live[job.id].status = 'working'; drawList();
      const n = done + 1, N = total(); fastMode = N > 10;
      const say = m => status(N > 1 ? `فاتورة ${n} من ${N}: ${m}` : m);
      setProg = p => progress(((n - 1) + Math.min(1, p)) / N);
      setProg(0.02); say('جارٍ فتح الملف…');
      const hash = await hashOf(job.file);
      const twin = hash && all().find(x => x._hash === hash);
      if (hash && sh[hash] && !twin) { runAdd(live[job.id].file, 'dup', 'هذا الملف أُرسل سابقًا إلى Google Sheet (InvID ' + sh[hash] + ')'); delete live[job.id]; done++; drawList(); continue; }
      if (twin) { runAdd(live[job.id].file, 'dup', 'نفس محتوى الملف «' + (twin.file || '') + '»'); delete live[job.id]; done++; drawList(); continue; }
      try {
        const { d, pages: urls } = await processFile(job.file, say);
        d._id = job.id; d._at = new Date().toISOString(); d._run = run && run.at; d._sig = job.sig; d._ns = job.ns; d._hash = hash; d._warnings = P.checks(d, d._qr);
        const same = all().find(x => x._id !== d._id && invNo(x.invoice_number) && invNo(x.invoice_number) === invNo(d.invoice_number) && dKey(x.invoice_date) === dKey(d.invoice_date) && sameVendor(x, d));
        if (same) { runAdd(live[job.id].file, 'dup', 'نفس الفاتورة محمّلة سابقًا من «' + (same.file || same.vendor_name || '') + '» (' + (d.invoice_number || '') + ')'); delete live[job.id]; done++; drawList(); continue; }
        imgs[job.id] = urls.map(srcOf); IDB.put(job.id, urls); runAdd(live[job.id].file, 'ok', '', job.id); delete live[job.id];
        store(d); batch.push(job.id);
        if (N === 1) open(job.id, true); // الدفعة لا تفتح شاشة المراجعة تلقائيًا؛ النتائج تظهر في الملخص
      } catch (e) {
        console.error(e); runAdd(live[job.id].file, 'error', e && e.message === 'img' ? 'تعذّر فتح الصورة' : ((e && e.message) || 'خطأ')); delete live[job.id];
      }
      done++; drawList();
    }
    } catch (e) { console.error(e); queue.length = 0; Object.keys(live).forEach(k => { if (live[k].status !== 'error') delete live[k]; }); perr = 'حدث خطأ غير متوقع أثناء القراءة: ' + ((e && e.message) || e); }
    busy = false; progress(null); setProg = () => { }; fastMode = false; idleWorkers(); if (batch.length) askPersist();
    if (perr) { status(perr, true); try { drawList(); } catch { } return; }
    const dm = dupMap(), nd = batch.filter(id => dm.has(id)).length;
    status((done > 1 ? `انتهت الدفعة: ${done} فواتير. التفاصيل في «آخر عملية تحميل»، وراجعها من تبويب «المراجعة».` : (cur && cur._qr ? 'تمت القراءة. وُجد رمز QR الضريبي.' : 'تمت القراءة. راجع الحقول وصحّح ما يلزم.'))
      + (nd ? ` تنبيه: ${nd === 1 ? 'فاتورة مكررة' : nd + ' فواتير مكررة'} (نفس المورّد والرقم والتاريخ)، انظر القائمة.` : ''), !!nd);
    if (cur) refresh();
    if (done > 1) $('lastRun').scrollIntoView({ behavior: matchMedia('(prefers-reduced-motion:reduce)').matches ? 'auto' : 'smooth' });
  }

  // ---------- كشف الفواتير المكررة: نفس المورّد (الرقم الضريبي أو الاسم) ونفس رقم الفاتورة ونفس التاريخ ----------
  const nk = s => P.normDigits(s).toLowerCase().replace(/[أإآ]/g, 'ا').replace(/ة/g, 'ه').replace(/ى/g, 'ي').replace(/[^0-9a-z\u0621-\u064A]/g, '');
  const vName = s => nk(String(s || '').replace(/^\s*(شركة|مؤسسة|مؤسسه|شركه)\s+/, '').replace(/\b(co|company|ltd|llc|est)\b\.?/gi, ''));
  // الرقم الضريبي للمورد: من SuppliersT إن كان مسجلًا، وإلا من الفاتورة
  const vatOf = x => { const s = !x._unreg && ref && x._supId != null && (ref.suppliers || []).find(z => z.id === x._supId); return (s && s.vat) || P.normDigits(String(x.vendor_tax_number || '')).replace(/\D/g, ''); };
  const vVat = s => { const d = P.normDigits(s).replace(/\D/g, ''); return d.length >= 10 ? d : ''; };
  const invNo = s => nk(s).replace(/^0+(?=.)/, '');
  const dKey = s => { const t = String(s || '').trim(); return t ? (P.findDate(t) || nk(t)) : ''; };
  function sameVendor(a, b) {
    const va = vVat(a.vendor_tax_number), vb = vVat(b.vendor_tax_number);
    if (va && vb) return va === vb;
    const na = vName(a.vendor_name), nb = vName(b.vendor_name);
    return !!na && na === nb;
  }
  // يعيد Map: رقم الفاتورة الداخلي -> قائمة الفواتير المكررة معها
  function dupMap() {
    const groups = {}, out = new Map();
    all().forEach(x => { const n = invNo(x.invoice_number), d = dKey(x.invoice_date); if (n && d) (groups[n + '|' + d] = groups[n + '|' + d] || []).push(x); });
    Object.values(groups).forEach(g => {
      if (g.length < 2) return;
      g.forEach(a => g.forEach(b => { if (a !== b && sameVendor(a, b)) { if (!out.has(a._id)) out.set(a._id, []); out.get(a._id).push(b); } }));
    });
    return out;
  }
  const dupMsg = list => `فاتورة مكررة: يوجد ${list.length === 1 ? 'فاتورة أخرى' : list.length + ' فواتير أخرى'} بنفس المورّد والرقم والتاريخ (${list.map(x => '«\u2068' + (x.file || x.vendor_name || '') + '\u2069»').join('، ')}).`;

  // ---------- العرض ----------
  function showPage() {
    const n = pages.length;
    $('pimg').src = n ? pages[pageIdx] : '';
    $('pimg').hidden = !n;
    $('pinfo').textContent = n > 1 ? `صفحة ${pageIdx + 1} من ${n}` : (n ? '' : 'لا توجد صورة محفوظة لهذه الفاتورة');
    $('prevP').hidden = $('nextP').hidden = n < 2;
    $('prevP').disabled = pageIdx === 0; $('nextP').disabled = pageIdx >= n - 1;
  }
  const numStr = v => (v == null || v === '') ? '' : String(v);
  const tag = k => cur.src && cur.src[k] ? `<span class="tag ${cur.src[k]}">${SRC[cur.src[k]]}</span>` : '';

  function open(id, scroll) {
    const d = all().find(x => x._id === id); if (!d) return;
    cur = d; qr = d._qr || null; rawText = d._raw || ''; textOnly = P.parseInvoice(rawText, null);
    cur.src = cur.src || {}; cur.items = cur.items || [];
    pages = imgs[id] || []; pageIdx = 0; showPage(); render();
    if (!pages.length) pagesOf(id).then(u => { if (cur && cur._id === id && u.length) { pages = u; showPage(); } });
    const want = d._approved ? 'appr' : 'rev'; if (tab !== want) setTab(want, true); else drawList();
    if (scroll) $('review').scrollIntoView({ behavior: matchMedia('(prefers-reduced-motion:reduce)').matches ? 'auto' : 'smooth' });
  }
  function render() {
    $('review').classList.add('on');
    $('rt').textContent = 'راجع: ' + (cur.file || 'الفاتورة');
    $('qrState').innerHTML = qr ? `<span class="qrok">رمز QR ضريبي صالح${qr.phase === 2 ? ' (المرحلة الثانية)' : ''}</span>` : (cur.qr_raw ? '<span class="note">وُجد رمز QR لكنه ليس رمز فاتورة ضريبية سعودية.</span>' : '<span class="note qrno">لم يُعثر على رمز QR ضريبي.</span>');
    $('fields').innerHTML = FIELDS.map(([k, l]) => `<div class="f"><label for="f_${k}">${l} ${tag(k)}</label><input id="f_${k}" data-k="${k}" value="${esc(cur[k])}"${k === 'vendor_tax_number' ? ' inputmode="numeric" class="num"' : ''}></div>`).join('');
    $('sums').innerHTML = SUMS.map(([k, l]) => `<div class="f"><label for="s_${k}">${l} ${tag(k)}</label><input id="s_${k}" data-s="${k}" class="num" inputmode="decimal" value="${esc(numStr(cur[k]))}"></div>`).join('');
    drawItems(); refresh();
    $('rawText').textContent = rawText || '(لا يوجد نص)';
  }
  function drawItems() {
    $('items').tBodies[0].innerHTML = cur.items.map((it, i) => `<tr data-i="${i}">
      <td class="d"><input aria-label="الوصف" data-c="description" value="${esc(it.description)}"></td>
      <td data-l="الكمية"><input aria-label="الكمية" class="num" inputmode="decimal" data-c="quantity" value="${esc(numStr(it.quantity))}"></td>
      <td data-l="سعر الوحدة"><input aria-label="سعر الوحدة" class="num" inputmode="decimal" data-c="unit_price" value="${esc(numStr(it.unit_price))}"></td>
      <td data-l="المجموع"><input aria-label="المجموع" class="num" inputmode="decimal" data-c="total" value="${esc(numStr(it.total))}"></td>
      <td class="x"><button data-del="${i}" aria-label="حذف البند">✕</button></td></tr>`).join('') ||
      '<tr><td colspan="5" class="note">لم تُقرأ بنود. يمكنك إضافتها يدويًا.</td></tr>';
  }
  function refresh() {
    const w = P.checks(cur, qr, { text: textOnly }); cur._warnings = w;
    const items = cur.items.filter(it => it.total !== '' && it.total != null && isFinite(+it.total));
    $('isum').textContent = items.length ? 'مجموع البنود: ' + P.fmt(items.reduce((a, it) => a + +it.total, 0)) : '';
    const dl = dupMap().get(cur._id), ws = dl ? [dupMsg(dl), ...w] : w;
    $('warns').innerHTML = ws.length ? ws.map(m => `<li>${esc(m)}</li>`).join('') : '<li class="good">المبالغ متطابقة والرقم الضريبي بصيغة صحيحة، ولا توجد فاتورة مكررة.</li>';
    $('bigTotal').textContent = cur.total != null && cur.total !== '' ? P.fmt(cur.total) + ' ' + (cur.currency || '') : '';
    const vi = $('f_vendor_tax_number'); if (vi) vi.classList.toggle('bad', !!P.checkVat(cur.vendor_tax_number));
  }
  const asNum = v => { const n = P.toNum(v); return v === '' ? '' : (n == null ? v : n); };
  const edited = () => { refresh(); store(cur, true); };

  $('fields').addEventListener('input', e => { const k = e.target.dataset.k; if (!k) return; cur[k] = e.target.value; delete cur.src[k]; edited(); });
  $('sums').addEventListener('input', e => { const k = e.target.dataset.s; if (!k) return; cur[k] = asNum(e.target.value); delete cur.src[k]; edited(); });
  $('items').addEventListener('input', e => {
    const tr = e.target.closest('tr'), c = e.target.dataset.c; if (!tr || !c) return;
    const it = cur.items[+tr.dataset.i]; it[c] = c === 'description' ? e.target.value : asNum(e.target.value);
    if ((c === 'quantity' || c === 'unit_price') && isFinite(+it.quantity) && isFinite(+it.unit_price) && it.quantity !== '' && it.unit_price !== '' && it.unit_price != null) {
      it.total = Math.round(it.quantity * it.unit_price * 100) / 100; tr.querySelector('[data-c=total]').value = it.total;
    }
    edited();
  });
  $('items').addEventListener('click', e => { const d = e.target.dataset.del; if (d == null) return; cur.items.splice(+d, 1); drawItems(); edited(); });
  $('addRow').onclick = () => { if (!cur) return; cur.items.push({ description: '', quantity: 1, unit_price: '', total: '' }); drawItems(); edited(); const ins = $('items').querySelectorAll('tbody tr:last-child input'); if (ins[0]) ins[0].focus(); };

  // ---------- النسخ والتنزيل ----------
  function asText(d) {
    const L = [];
    FIELDS.forEach(([k, l]) => { if (d[k]) L.push(`${l}: ${d[k]}`); });
    if (d.items.length) { L.push('', 'البنود:'); d.items.forEach((it, i) => L.push(`${i + 1}. ${it.description || ''} | الكمية: ${numStr(it.quantity)} | السعر: ${P.fmt(it.unit_price)} | المجموع: ${P.fmt(it.total)}`)); }
    const s = SUMS.filter(([k]) => d[k] !== null && d[k] !== undefined && d[k] !== '');
    if (s.length) { L.push(''); s.forEach(([k, l]) => L.push(`${l}: ${k === 'tax_rate' ? d[k] + '%' : P.fmt(d[k]) + ' ' + (d.currency || '')}`)); }
    return L.join('\n');
  }
  function asJson(d) {
    const o = {}; FIELDS.forEach(([k]) => o[k] = d[k] || ''); o.items = d.items; SUMS.forEach(([k]) => o[k] = d[k] === '' ? null : d[k]);
    const q = d._qr; o.zatca_qr = q ? { seller: q.seller, vat_number: q.vat_number, timestamp: q.timestamp, total: q.total, vat: q.vat } : null;
    return JSON.stringify(o, null, 2);
  }
  async function copy(t) {
    try { await navigator.clipboard.writeText(t); toast('تم النسخ'); }
    catch { const a = document.createElement('textarea'); a.value = t; document.body.append(a); a.select(); let ok = false; try { ok = document.execCommand('copy'); } catch { } a.remove(); toast(ok ? 'تم النسخ' : 'تعذّر النسخ'); }
  }
  function download(name, data, type) {
    try {
      const blob = data instanceof Blob ? data : new Blob([data], { type });
      const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name; document.body.append(a); a.click();
      setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000); toast('جارٍ التنزيل'); return true;
    } catch (e) { console.error(e); toast('تعذّر التنزيل هنا، استخدم النسخ'); return false; }
  }
  const stamp = () => new Date().toISOString().slice(0, 10);
  const safe = s => String(s || '').replace(/[^\w؀-ۿ-]+/g, '_').slice(0, 40);
  $('copyText').onclick = () => cur && copy(asText(cur));
  $('copyJson').onclick = () => cur && copy(asJson(cur));
  $('copyRaw').onclick = () => copy(rawText);
  $('dlOneXlsx').onclick = () => cur && download(`فاتورة-${safe(cur.invoice_number || cur.vendor_name || cur._id)}.xlsx`, X.toXlsx([cur], ref && ref.vatCols));
  $('dlOneCsv').onclick = () => cur && download(`بنود-${safe(cur.invoice_number || cur._id)}.csv`, X.itemsCsv([cur]), 'text/csv;charset=utf-8');
  $('dlJson').onclick = () => cur && download(`invoice-${safe(cur.invoice_number || cur._id)}.json`, asJson(cur), 'application/json');

  // ---------- قائمة الفواتير (محفوظة في المتصفح) ----------
  const X = window.InvoiceExport;
  // الفواتير محفوظة في IndexedDB (store.js) بلا حد لعددها. القائمة في الذاكرة (cache) ونحفظ ما تغيّر فقط.
  const S = window.InvoiceStore;
  const KEY = 'invoice-reader-history-v1'; // المكان القديم في localStorage: يُنقل منه مرة واحدة
  let cache = [];
  const all = () => cache;
  const snap = new Map(); // آخر نسخة محفوظة من كل فاتورة (لمعرفة ما تغيّر)
  const ser = x => { try { return JSON.stringify(x); } catch { return ''; } };
  // تعدد التبويبات: كل تبويب يُعلم الآخرين بما حفظه أو حذفه فيحدّثون قوائمهم
  const bc = typeof BroadcastChannel !== 'undefined' ? new BroadcastChannel('invoice-reader') : null;
  let saveOk = true;
  function saveFail(e) {
    console.error(e); saveOk = false; const w = $('saveWarn'); if (!w) return;
    w.textContent = '⚠ تعذّر حفظ آخر التعديلات في المتصفح (قد تكون مساحة الجهاز ممتلئة). صدّر الفواتير أو نزّل نسخة احتياطية من تبويب «التصدير» قبل إغلاق الصفحة.'; w.hidden = false;
  }
  function saved() { if (!saveOk) { saveOk = true; $('saveWarn').hidden = true; } }
  function write(list) { if (!list.length) return; S.putInv(list).then(() => { saved(); if (bc) bc.postMessage({ put: list.map(x => x._id) }); }, saveFail); }
  // يحفظ الفواتير التي تغيّرت ويحذف المحذوفة من القائمة
  function persist() {
    const ch = [], have = new Set();
    cache.forEach(x => { have.add(x._id); const s = ser(x); if (snap.get(x._id) !== s) { snap.set(x._id, s); ch.push(x); } });
    const gone = [...snap.keys()].filter(id => !have.has(id));
    write(ch);
    if (gone.length) { gone.forEach(id => snap.delete(id)); S.delInv(gone).then(() => { if (bc) bc.postMessage({ del: gone }); }, saveFail); }
  }
  // حفظ فاتورة واحدة بعد توقف الكتابة قليلًا (بدون تحويل كل القائمة)
  const pendSave = new Set();
  function saveSoon(x, ms) {
    pendSave.add(x); clearTimeout(saveSoon.t);
    saveSoon.t = setTimeout(() => { const l = [...pendSave].filter(v => cache.includes(v)); pendSave.clear(); l.forEach(v => snap.set(v._id, ser(v))); write(l); }, ms == null ? 400 : ms);
  }
  function store(d, quiet) {
    const h = all(), i = h.findIndex(x => x._id === d._id);
    if (i >= 0) h[i] = d; else h.unshift(d);
    saveSoon(d, quiet ? 400 : 0);
    if (quiet) { clearTimeout(store.h); store.h = setTimeout(drawList, 500); } else drawList();
  }
  window.addEventListener('pagehide', () => { if (pendSave.size) { clearTimeout(saveSoon.t); const l = [...pendSave]; pendSave.clear(); l.forEach(v => snap.set(v._id, ser(v))); write(l); } });
  // ---------- البيانات المرجعية (الموردين والمشاريع والمنتجات) ----------
  const R = window.InvoiceRef, RKEY = 'invoice-reader-ref-v1';
  let ref = null; // يُحمَّل من IndexedDB عند بدء التشغيل
  const saveRef = () => S.setKv('ref', ref).then(() => { if (bc) bc.postMessage({ ref: 1 }); }, saveFail);
  let supEdit = null; // فاتورة يجري تغيير مورّدها يدويًا
  // مطابقة المورّد محفوظة مؤقتًا حسب (الرقم الضريبي + الاسم) حتى لا تُعاد مع كل رسم للقائمة
  let lkMemo = new Map(), lkRef = null, supSorted = null;
  function matchSup(x) {
    if (lkRef !== ref) { lkMemo = new Map(); lkRef = ref; supSorted = null; }
    const k = String(x.vendor_tax_number || '') + '\u0001' + String(x.vendor_name || '');
    if (!lkMemo.has(k)) lkMemo.set(k, R.matchSupplier(ref, x));
    return lkMemo.get(k);
  }
  const sortedSup = () => { if (lkRef !== ref) matchSup({}); return supSorted || (supSorted = ref.suppliers.slice().sort((a, b) => a.name.localeCompare(b.name, 'ar'))); };
  function link(x) {
    if (!ref) return;
    if (!x._supManual) { const m = matchSup(x); x._supId = m ? m.id : null; x._supBy = m ? m.by : null; x._unreg = !m; }
    // المورّد غير موجود في SuppliersT: يُسجَّل «مورد غير مسجل» و«منتج غير مسجل» ويُسمح بالتصدير
    if (x._unreg) { x._supId = null; x._supName = 'مورد غير مسجل'; x.product_id = null; const pr0 = R.project(ref, x.project_id); x.project_name = pr0 ? pr0.name : ''; x.product_name = 'منتج غير مسجل'; x._exists = null; return; }
    const s = R.supplier(ref, x._supId); x._supName = s ? s.name : '';
    const prods = x._supId != null ? R.productsFor(ref, x._supId) : [];
    if (x.product_id != null && prods.length && !prods.some(p => p.id === x.product_id) && !x._prodAny) { x.product_id = null; }
    if (x.product_id == null && prods.length === 1 && !x._prodCleared) x.product_id = prods[0].id; // لا نعيد الاختيار التلقائي إذا مسحه المستخدم
    const pr = R.project(ref, x.project_id), pd = R.product(ref, x.product_id);
    x.project_name = pr ? pr.name : ''; x.product_name = pd ? pd.name : '';
    const ex = R.existsInvoice(ref, x._supId, x.invoice_number); x._exists = ex || null;
  }
  function refNotes(x) {
    if (!ref) return [];
    const n = [];
    if (x._exists && String(x._exists) !== String(x._sent)) n.push(['wn', `مسجلة سابقًا في InvoiceT${x._exists !== true ? ' (InvID ' + x._exists + ')' : ''}`]);
    if (x._unreg) n.push(['wn', 'المورّد غير موجود في SuppliersT: سيُسجَّل «مورد غير مسجل» والمنتج «منتج غير مسجل»، واسم المورّد ورقمه الضريبي في الملاحظات']);
    else if (x._supId == null) n.push(['wn', 'اختر المورّد المسجل']);
    else if (x._supBy === 'name') n.push(['wn', 'طوبق المورّد بالاسم فقط، تأكد منه']);
    if (x.project_id == null) n.push(['todo', 'اختر المشروع']);
    if (x.product_id == null && !x._unreg) n.push(['todo', 'اختر المنتج']);
    return n;
  }
  const opt = (v, label, sel) => `<option value="${esc(v)}"${sel ? ' selected' : ''}>${esc(label)}</option>`;
  function supCell(x) {
    if (!ref) return '';
    const id = esc(x._id);
    if (x._unreg && supEdit !== x._id) return `<small class="supunreg">مورد غير مسجل في SuppliersT</small><button class="linkb" data-supedit="${id}">اختيار مورّد مسجل</button>`;
    // قائمة الموردين الكاملة تُبنى فقط عند الضغط (توفّر آلاف العناصر في الصفحة)
    if (x._supId == null && !x._unreg && supEdit !== x._id) return `<button class="linkb unsetb" data-supedit="${id}">⚠ اختر المورّد المسجل</button>`;
    if (x._supId != null && supEdit !== x._id) return `<small class="${x._supBy === 'name' ? 'supname' : 'supok'}">SupID ${x._supId}: ${esc(x._supName)}${x._supBy === 'name' ? ' (بالاسم)' : x._supBy === 'manual' ? ' (يدوي)' : ''}</small><button class="linkb" data-supedit="${id}">تغيير</button>`;
    const list = sortedSup();
    return `<select data-sup="${id}" aria-label="اختر المورّد المسجل"${x._supId == null ? ' class="unset"' : ''}>${opt('', '⚠ — اختر المورّد المسجل —', x._supId == null && !x._unreg)}${opt('unreg', '✱ مورد غير مسجل', !!x._unreg)}${x._supManual ? opt('auto', '↺ مطابقة تلقائية', false) : ''}${list.map(s => opt(s.id, s.id + ' - ' + s.name, s.id === x._supId)).join('')}</select>`;
  }
  function linkCell(x) {
    if (!ref) return '<span class="todo">حمّل ملف البيانات</span>';
    const id = esc(x._id);
    const projs = R.activeProjects(ref).slice(); const cp = R.project(ref, x.project_id); if (cp && !cp.active) projs.unshift(cp);
    let prods = x._supId != null ? R.productsFor(ref, x._supId) : [], all = false;
    if (!prods.length || x._prodAny) { prods = ref.products.slice().sort((a, b) => a.name.localeCompare(b.name, 'ar')); all = true; }
    const cpd = R.product(ref, x.product_id); if (cpd && !prods.some(p => p.id === cpd.id)) prods.unshift(cpd);
    return `<label>المشروع<select data-proj="${id}"${x.project_id == null ? ' class="unset"' : ''}>${opt('', '⚠ — اختر المشروع —', x.project_id == null)}${projs.map(p => opt(p.id, p.id + ' - ' + p.name, p.id === x.project_id)).join('')}</select></label>
      ${x._unreg ? '<label>المنتج<span class="unregp">منتج غير مسجل</span></label>' : `<label>المنتج${all && x._supId != null && !x._prodAny ? ' (لا منتجات مرتبطة بالمورّد)' : ''}<select data-prod="${id}"${x.product_id == null ? ' class="unset"' : ''}>${opt('', '⚠ — اختر المنتج —', x.product_id == null)}${prods.map(p => opt(p.id, p.id + ' - ' + p.name, p.id === x.product_id)).join('')}${!all ? opt('any', '… كل المنتجات', false) : ''}</select></label>`}`;
  }
  function refInfo() {
    $('refClear').hidden = !ref;
    $('refInfo').parentNode.classList.toggle('ok', !!ref);
    $('refInfo').textContent = ref ? `${ref.fileName}: ${ref.suppliers.length} مورّد، ${R.activeProjects(ref).length} مشاريع فعّالة، ${ref.products.length} منتج، ${Object.keys(ref.inv).length} فاتورة سابقة للتحقق من التكرار. محفوظ في هذا المتصفح فقط.` : 'لم يُحمَّل ملف بعد. حمّل ملف Excel الذي فيه أوراق SuppliersT و ProjectsT و ProductsT و SupplierProductT لمطابقة الموردين واختيار المشروع والمنتج. يُقرأ الملف داخل جهازك ولا يُرفع لأي مكان.';
  }
  $('refFile').onchange = async e => {
    const f = e.target.files[0]; e.target.value = ''; if (!f) return;
    try {
      const r = await R.load(f); ref = r;
      saveRef();
      all().forEach(x => { if (!x._supManual) { delete x._supId; delete x._unreg; } }); persist();
      refInfo(); drawList(); toast('تم تحميل ملف البيانات ومطابقة الفواتير');
    } catch (err) { console.error(err); toast((err && err.message) || 'تعذّرت قراءة الملف'); }
  };
  $('refClear').onclick = () => { ref = null; saveRef(); refInfo(); drawList(); toast('أُزيل ملف البيانات من هذا المتصفح'); };
  $('hist').addEventListener('change', e => {
    const t = e.target, d = t.dataset, id = d.sup || d.proj || d.prod; if (!id) return;
    const x = all().find(v => v._id === id); if (!x) return;
    if (d.sup) { if (t.value === 'auto') { x._supManual = false; } else if (t.value === 'unreg') { x._supManual = true; x._unreg = true; x._supId = null; x._supBy = 'manual'; } else { x._supManual = true; x._unreg = false; x._supId = t.value ? +t.value : null; x._supBy = 'manual'; } x._prodAny = false; supEdit = null; }
    if (d.proj) x.project_id = t.value ? +t.value : null;
    if (d.prod) { if (t.value === 'any') { x._prodAny = true; } else { x.product_id = t.value ? +t.value : null; x._prodCleared = !t.value; } }
    persist(); drawList();
  });

  // ---------- ربط Google Sheet (مجاني عبر Google Apps Script) ----------
  const GKEY = 'invoice-reader-gsheet-v1';
  // رابط الشيت ثابت من config.js (يحدده المالك في GitHub)، والمستخدم يُدخل بريده ورمزه فقط
  let gs = null; try { gs = JSON.parse(localStorage.getItem(GKEY)) || null; } catch { gs = null; }
  // إن لم يُضبط الرابط في config.js بعد، نستمر بالرابط المحفوظ سابقًا في هذا المتصفح (بدون أي خانة لتغييره)
  const GURL = String((window.APP_CONFIG || {}).gsheetUrl || '').trim() || String((gs && gs.url) || '').trim();
  if (gs) { gs.url = GURL; if (!GURL || !gs.token) gs = null; }
  function gsUi() {
    $('gsBox').hidden = !GURL;
    $('gsState').textContent = gs ? `(مربوط باسم ${gs.email || '—'} ✓)` : '(غير مربوط)'; $('gsState').classList.toggle('on', !!gs);
    $('gsSend').hidden = !GURL; $('gsLoad').hidden = !gs; $('gsForget').hidden = !gs;
    if (gs) { $('gsEmail').value = gs.email || ''; $('gsToken').value = gs.token; }
  }
  // طلب مع مهلة: إن لم يرد الشيت خلال المهلة نتوقف بدل الانتظار بلا نهاية
  async function fetchT(url, opt, ms) {
    const c = typeof AbortController !== 'undefined' ? new AbortController() : null, t = c && setTimeout(() => c.abort(), ms);
    try { return await fetch(url, Object.assign({ redirect: 'follow' }, opt, c ? { signal: c.signal } : {})); }
    catch (e) { if (e && e.name === 'AbortError') throw new Error('انتهت مهلة الاتصال بـ Google Sheet، تحقق من الإنترنت وحاول مرة أخرى'); throw new Error('تعذّر الاتصال بـ Google Sheet، تحقق من الإنترنت'); }
    finally { if (t) clearTimeout(t); }
  }
  // قراءة رد الشيت: إن لم يكن JSON (صفحة خطأ من Google مثلًا) نعطي رسالة مفهومة بدل «Unexpected token»
  async function readJson(r) {
    const t = await r.text();
    try { return JSON.parse(t); } catch {
      throw new Error(r.status >= 400 ? `ردّ Google Sheet بخطأ (${r.status}). أعد نشر السكربت أو حاول لاحقًا.` : 'ردّ Google Sheet بصفحة غير متوقعة. تأكد أن السكربت منشور (Deploy) كـ Web app والوصول «Anyone»، وأن بريدك ورمزك صحيحان.');
    }
  }
  // الطلبات ترسل البريد والرمز داخل جسم POST (لا في الرابط حتى لا يُسجَّلا في سجلات الروابط).
  // إن كان السكربت المنشور قديمًا (لا يعرف action) نرجع إلى الطريقة القديمة GET.
  async function gsCall(cfg, action, ms) {
    const body = JSON.stringify({ action, email: cfg.email || '', token: cfg.token });
    let j = await readJson(await fetchT(cfg.url, { method: 'POST', body }, ms));
    const okShape = action === 'ref' ? j.sheets : j.hashes;
    if (j.ok && !okShape) {
      const u = cfg.url + (cfg.url.includes('?') ? '&' : '?') + 'action=' + action + '&email=' + encodeURIComponent(cfg.email || '') + '&token=' + encodeURIComponent(cfg.token);
      j = await readJson(await fetchT(u, {}, ms));
    }
    return j;
  }
  async function gsGet(cfg) {
    const j = await gsCall(cfg, 'ref', 30000);
    if (!j.ok) throw new Error(j.error || 'تعذّر الاتصال بالشيت'); return j;
  }
  async function gsLoadRef(cfg, quiet) {
    const j = await gsGet(cfg);
    const r = R.fromSheets(j.sheets, 'Google Sheet: ' + (j.title || '')); r.hashes = j.hashes || {};
    ref = r; saveRef();
    all().forEach(x => { if (!x._supManual) { delete x._supId; delete x._unreg; } }); persist(); refInfo(); drawList();
    if (!quiet) toast('حُدّثت البيانات المرجعية من Google Sheet');
  }
  $('gsSave').onclick = async () => {
    const url = GURL, email = $('gsEmail').value.trim().toLowerCase(), token = $('gsToken').value.trim();
    if (!/^https:\/\/script\.google(usercontent)?\.com\//.test(url)) return toast('الرابط يجب أن يكون رابط Web app من script.google.com');
    if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) return toast('اكتب بريدك الإلكتروني'); if (!token) return toast('الصق رمز الدخول');
    $('gsSave').disabled = true; status('جارٍ اختبار الربط مع Google Sheet…');
    try { await gsLoadRef({ url, email, token }, true); gs = { url, email, token }; localStorage.setItem(GKEY, JSON.stringify({ url, email, token })); gsUi(); $('gsBox').open = false; status('تم الربط مع Google Sheet، وحُمّلت منه البيانات المرجعية.'); }
    catch (e) { console.error(e); status('تعذّر الربط: ' + ((e && e.message) || 'تحقق من الرابط والبريد والرمز، وأن النشر «Anyone»'), true); }
    finally { $('gsSave').disabled = false; }
  };
  $('gsForget').onclick = () => { gs = null; localStorage.setItem(GKEY, JSON.stringify({ url: GURL })); $('gsEmail').value = ''; $('gsToken').value = ''; gsUi(); toast('أُلغي الربط مع Google Sheet'); };
  $('gsLoad').onclick = async () => { try { await gsLoadRef(gs); } catch (e) { toast((e && e.message) || 'تعذّر التحديث من الشيت'); } };
  $('gsSend').onclick = async () => {
    if (!gs) return amsg('Google Sheet غير مربوط. اربطه من تبويب «تحميل الفواتير» ← «البيانات المرجعية والربط».', true);
    const appr = all().filter(x => x._approved && !x._sent);
    if (!appr.length) return amsg(all().some(x => x._approved) ? 'كل الفواتير المعتمدة أُرسلت إلى Google Sheet سابقًا.' : 'لا توجد فواتير معتمدة. اعتمد الفواتير أولًا من تبويب «المراجعة».', true);
    const ready = appr.filter(x => (x._supId != null || x._unreg) && x.project_id != null && (x.product_id != null || x._unreg)), missing = appr.length - ready.length;
    if (!ready.length) return amsg(`لم يُرسل شيء: ${missing} فاتورة معتمدة ينقصها المورّد المسجل أو المشروع أو المنتج. اخترها من القوائم الحمراء.`, true);
    const payload = { email: gs.email || '', token: gs.token, invoices: ready.map(x => ({ id: x._id, InvoiceNo: String(x.invoice_number || ''), SupID: x._unreg ? '' : x._supId, Unreg: !!x._unreg, SupfName: x._supName || x.vendor_name || '', ProjectID: x.project_id, ProjectName: x.project_name || '', ProductID: x._unreg ? '' : (x.product_id ?? ''), ProductName: x._unreg ? 'منتج غير مسجل' : (x.product_name || ''), Date: x.invoice_date || '', Invamount: x.total === '' ? '' : x.total, invVat: x.tax_amount === '' ? '' : x.tax_amount, Note: [x._unreg ? `مورد غير مسجل: ${x.vendor_name || '—'}${x.vendor_tax_number ? ' (الرقم الضريبي ' + x.vendor_tax_number + ')' : ''}` : '', x.note || ''].filter(Boolean).join(' | '), VATID: vatOf(x), FilePath: x.file || '', FileHash: x._hash || '' })) };
    $('gsSend').disabled = true; amsg(`جارٍ إرسال ${ready.length} فاتورة إلى Google Sheet…`);
    try {
      const j = await readJson(await fetchT(gs.url, { method: 'POST', body: JSON.stringify(payload) }, 90000));
      if (!j.ok) throw new Error(j.error);
      const h = all(), now = new Date().toISOString();
      (j.added || []).forEach(a => { const x = h.find(v => v._id === a.id); if (x) { x._sent = a.invId; x._sentAt = now; if (ref && x._hash) (ref.hashes = ref.hashes || {})[x._hash] = a.invId; if (ref) ref.inv[x._supId + '|' + String(x.invoice_number).toLowerCase().replace(/[^0-9a-zء-ي]/g, '').replace(/^0+(?=.)/, '')] = a.invId; } });
      (j.skipped || []).forEach(a => { const x = h.find(v => v._id === a.id); if (x && a.invId) { x._sent = a.invId; x._sentAt = now; x._sentNote = a.byFile ? 'نفس الملف أُرسل سابقًا' : 'كانت موجودة مسبقًا'; delete x._sendErr; } else if (x) x._sendErr = a.reason || 'رفضها الشيت'; });
      (j.added || []).forEach(a => { const x = h.find(v => v._id === a.id); if (x) delete x._sendErr; });
      persist(); if (ref) saveRef(); drawList();
      const sk = (j.skipped || []).length;
      toast(`أُرسلت ${(j.added || []).length} فاتورة إلى Google Sheet`);
      amsg(`أُضيفت ${(j.added || []).length} فاتورة إلى InvoiceT و VatReport.` + (sk ? ` لم تُضف ${sk} لأنها موجودة مسبقًا أو بلا رقم.` : '') + (missing ? ` لم تُرسل ${missing} لأن المورّد أو المشروع أو المنتج غير محدد.` : ''), !!missing);
    } catch (e) { console.error(e); const m = (e && e.message) || 'تحقق من الاتصال'; ready.forEach(x => x._sendErr = 'فشل الإرسال إلى Google Sheet: ' + m); persist(); drawList(); amsg('تعذّر الإرسال: ' + m, true); }
    finally { $('gsSend').disabled = false; }
  };

  let tab = 'up';
  const isExported = x => !!(x._exportedAt || x._sent);
  // شارة التصدير بجانب «معتمدة»: أخضر «مصدّرة» أو أحمر «غير مصدّرة»
  const expBadge = x => isExported(x)
    ? `<span class="xb yes" style="background:#1e7a46;color:#fff" title="${esc([x._sent ? 'في Google Sheet (InvID ' + x._sent + ')' : '', x._exportedAt ? 'صُدّرت إلى Excel' : ''].filter(Boolean).join(' · '))}">مصدّرة</span>`
    : '<span class="xb no" style="background:#c0392b;color:#fff" title="لم تُصدّر ولم تُرسل إلى Google Sheet">غير مصدّرة</span>';
  function amsg(t, bad) { const e = $('apprMsg'); e.textContent = t; e.className = 'amsg' + (bad ? ' bad' : ''); e.hidden = !t; toast(t); }
  // سبب عدم تصدير الفاتورة (يظهر في التحذيرات والجداول)
  function whyNot(x) {
    const r = [];
    if (!x._approved) r.push('لم تُعتمد بعد، راجعها واعتمدها من تبويب «المراجعة»');
    else {
      const miss = [x._supId == null && !x._unreg && 'المورّد المسجل', x.project_id == null && 'المشروع', x.product_id == null && !x._unreg && 'المنتج'].filter(Boolean);
      if (x._sendErr) r.push(x._sendErr);
      if (miss.length) r.push('ينقصها ' + miss.join(' و') + '، فلا يمكن إرسالها إلى Google Sheet');
      if (!r.length) r.push('جاهزة، ولم يُضغط زر التصدير بعد (تبويب «المعتمدة»)');
    }
    if (x._dup) r.push('مكررة مع فاتورة أخرى');
    return r;
  }
  const shownList = () => all().filter(x => tab === 'appr' ? x._approved && !isExported(x) : !x._approved);
  const money = x => x.total != null && x.total !== '' ? P.fmt(x.total) + ' ' + (x.currency || '') : '—';
  const sumBy = list => Object.entries(list.reduce((a, x) => { const c = x.currency || '—'; a[c] = (a[c] || 0) + (+x.total || 0); return a; }, {})).map(([c, v]) => P.fmt(v) + ' ' + c).join(' · ');
  function drawList() {
    const h = all(), pend = Object.entries(live);
    $('count').textContent = `(${shownList().length})`;
    $('exportBar').hidden = !h.length;
    const dm = dupMap();
    h.forEach(x => { const d = dm.get(x._id); if (d) x._dup = dupMsg(d); else delete x._dup; link(x); });
    const nA = h.filter(x => x._approved).length, nP = h.length - nA;
    $('ht').firstChild.textContent = tab === 'appr' ? 'معتمدة لم تُصدّر بعد ' : 'مرفوعة لم تُعتمد بعد ';
    const STAT = { queued: 'بالانتظار', working: 'جارٍ القراءة…', error: 'تعذّرت القراءة', dup: 'لم يُحمّل (مكرر)' };
    const rowsP = ''; void pend.map(([id, v]) => `<tr class="pend ${v.status}"><td>${esc(v.file)}</td><td colspan="6">${STAT[v.status]}${v.err ? ': ' + esc(v.err) : ''}</td><td class="act">${v.status === 'error' || v.status === 'dup' ? `<button data-dismiss="${esc(id)}">إخفاء</button>` : ''}</td></tr>`).join('');
    // نرسم فقط التبويب الظاهر (أسرع بكثير على الجوال مع كثرة الفواتير)
    const onList = tab === 'rev' || tab === 'appr';
    const shown = onList ? shownList() : [];
    const rows = shown.map(x => {
      const id = esc(x._id), dup = dm.get(x._id), w = x._warnings || [];
      const notes = [...(dup ? ['<span class="dupb">مكررة</span>'] : []), ...refNotes(x).map(([c, m]) => `<span class="${c}">• ${esc(m)}</span>`), ...w.map(m => `<span class="wn">• ${esc(m)}</span>`)].join('');
      return `<tr class="${cur && cur._id === x._id ? 'sel' : ''}${dup ? ' dup' : ''}${x._approved ? ' appr' : ''}"${dup ? ` title="${esc(x._dup)}"` : ''}>
        <td data-l="المورّد"><b>${esc(x.vendor_name || 'بدون اسم')}</b><small class="fpath" title="${esc(x.file || '')}">📄 ${esc(x.file || '')}</small>${supCell(x)}</td>
        <td data-l="الرقم الضريبي" class="mono">${esc(x.vendor_tax_number || '—')}</td>
        <td data-l="رقم الفاتورة" class="nw">${esc(x.invoice_number || '—')}</td>
        <td data-l="التاريخ" class="nw">${esc(x.invoice_date || '—')}</td>
        <td data-l="الإجمالي" class="n"><b>${esc(money(x))}</b><small>الضريبة: ${x.tax_amount != null && x.tax_amount !== '' ? esc(P.fmt(x.tax_amount)) : '—'}</small></td>
        <td data-l="المشروع والمنتج" class="lnk">${linkCell(x)}</td>
        <td data-l="الملاحظات" class="notes">${notes ? `<div class="auto">${notes}</div>` : '<div class="auto"><span class="ok">لا تنبيهات</span></div>'}<input class="unote" data-note="${id}" value="${esc(x.note || '')}" placeholder="أضف ملاحظتك…" aria-label="ملاحظتك على الفاتورة"></td>
        <td class="act"><div class="st">${x._approved ? `<span class="okb">معتمدة</span>${expBadge(x)}${x._approvedAt ? `<small>${esc(new Date(x._approvedAt).toLocaleDateString('ar-SA-u-ca-gregory-nu-latn'))}</small>` : ''}` : '<span class="pendb">بانتظار الاعتماد</span>'}</div>${x._approved ? `<button data-unapprove="${id}">إلغاء الاعتماد</button>` : `<button class="primary" data-approve="${id}">اعتماد</button>`}<button data-open="${id}">مراجعة</button><button data-rm="${id}" class="del">حذف</button></td></tr>`;
    }).join('');
    if (onList) $('hist').tBodies[0].innerHTML = rowsP + rows || `<tr><td colspan="8" class="empty">${tab === 'appr' ? 'لا توجد فواتير معتمدة بعد.' : h.length ? 'لا توجد فواتير بانتظار المراجعة. كل الفواتير معتمدة ✓' : 'لا توجد فواتير بعد. حمّلها من تبويب «تحميل الفواتير».'}</td></tr>`;
    const nd = dm.size;
    $('dupNote').hidden = !nd;
    $('dupNote').innerHTML = nd ? `تنبيه: ${nd} فواتير في القائمة مكررة (نفس المورّد ورقم الفاتورة والتاريخ). <button id="dupClean" class="del">حذف المكرر وإبقاء نسخة واحدة</button>` : '';
    if (nd) $('dupClean').onclick = () => {
      const h = all(), before = h.slice(), keep = new Set(), drop = new Set();
      // نُبقي من كل مجموعة: المرسلة للشيت، ثم المعتمدة، ثم الأقدم
      const rank = x => (x._sent ? 0 : 2) + (x._approved ? 0 : 1);
      [...h].sort((a, b) => rank(a) - rank(b) || String(a._at || '').localeCompare(String(b._at || ''))).forEach(x => {
        if (drop.has(x._id)) return; keep.add(x._id); (dm.get(x._id) || []).forEach(o => { if (!keep.has(o._id)) drop.add(o._id); });
      });
      for (let i = h.length - 1; i >= 0; i--) if (drop.has(h[i]._id)) h.splice(i, 1);
      persist(); drawList();
      undo(`حُذفت ${drop.size} فاتورة مكررة`, () => { h.splice(0, h.length, ...before); persist(); drawList(); });
    };
    drawQty();
    $('sumAll').textContent = shownList().length ? `المجموع: ${sumBy(shownList())}` : '';
    void nP; drawTabs();
  }
  // ---------- جدول الكميات (مستقل) ----------
  const qOpen = new Set();
  const f2 = v => v == null ? '—' : P.fmt(v);
  function drawQty() {
    const h = all(), groups = X.qtyGroups(h);
    $('qEmpty').hidden = !!groups.length; $('qtyBody').hidden = !groups.length; $('bQty').textContent = groups.filter(g => g.diff).length || '';
    if (!groups.length || tab !== 'qty') return;
    const nd = groups.filter(g => g.diff).length;
    $('qcount').textContent = `(${groups.length} صنف)`;
    $('qsum').textContent = 'إجمالي البنود: ' + P.fmt(groups.reduce((a, g) => a + g.total, 0));
    $('qwarn').hidden = !nd;
    $('qwarn').textContent = nd ? `تنبيه: ${nd === 1 ? 'صنف واحد اختلف سعره' : nd + ' أصناف اختلف سعرها'} من فاتورة إلى أخرى. هي في أعلى الجدول باللون الأحمر.` : '';
    const only = $('qOnlyDiff').checked, q = $('qSearch').value.trim().toLowerCase();
    const shown = groups.filter(g => (!only || g.diff) && (!q || g.name.toLowerCase().includes(q) || g.key.includes(q)));
    $('qty').tBodies[0].innerHTML = shown.map(g => {
      const k = esc(g.key), open = qOpen.has(g.key);
      const price = g.min == null ? '—' : g.diff ? `${f2(g.min)} – ${f2(g.max)}` : f2(g.min);
      const det = open ? `<tr class="det"><td colspan="8"><table class="qd"><thead><tr><th>المورّد</th><th>رقم الفاتورة</th><th>التاريخ</th><th class="n">الكمية</th><th class="n">سعر الوحدة</th><th class="n">المجموع</th><th></th></tr></thead><tbody>${g.lines.map(l => `<tr>
          <td data-l="المورّد">${esc(l.inv._supName || l.inv.vendor_name || '—')}</td><td data-l="رقم الفاتورة">${esc(l.inv.invoice_number || l.inv.file || '—')}</td><td data-l="التاريخ">${esc(l.inv.invoice_date || '—')}</td>
          <td data-l="الكمية" class="n">${f2(l.qty)}</td><td data-l="سعر الوحدة" class="n ${g.diff && l.price === g.max ? 'hi' : g.diff && l.price === g.min ? 'lo' : ''}">${f2(l.price)}</td><td data-l="المجموع" class="n">${f2(l.total)}</td>
          <td><button class="linkb" data-open="${esc(l.inv._id)}">فتح الفاتورة</button></td></tr>`).join('')}</tbody></table></td></tr>` : '';
      return `<tr class="${g.diff ? 'pdiff' : ''}">
        <td data-l="الصنف"><b>${esc(g.name)}</b>${g.products.length ? `<small>${esc(g.products.join('، '))}</small>` : ''}</td>
        <td data-l="الفواتير" class="n">${g.invs}</td>
        <td data-l="الكمية" class="n">${f2(g.qty)}</td>
        <td data-l="سعر الوحدة" class="n pr">${price}${g.diff ? `<small>المتوسط ${f2(g.avg)}</small>` : ''}</td>
        <td data-l="الإجمالي" class="n"><b>${f2(g.total)}</b></td>
        <td data-l="الموردون">${esc(g.vendors.join('، ') || '—')}</td>
        <td data-l="تنبيه السعر">${g.diff ? `<span class="wn">اختلف السعر بفرق ${f2(g.diff)} (${g.pct}%)</span>` : (g.lines.length > 1 ? '<span class="ok">السعر ثابت</span>' : '<span class="todo">فاتورة واحدة</span>')}</td>
        <td class="act"><button data-qdet="${k}" aria-expanded="${open}">${open ? 'إخفاء' : 'التفاصيل'}</button></td></tr>${det}`;
    }).join('') || '<tr><td colspan="8" class="empty">لا توجد أصناف مطابقة.</td></tr>';
  }
  $('qty').addEventListener('click', e => {
    const k = e.target.dataset.qdet, o = e.target.dataset.open;
    if (k != null) { qOpen.has(k) ? qOpen.delete(k) : qOpen.add(k); drawQty(); }
    if (o) open(o, true);
  });
  $('qOnlyDiff').onchange = drawQty;
  $('qSearch').oninput = drawQty;
  $('qXlsx').onclick = () => all().length && download(`الكميات-${stamp()}.xlsx`, X.qtyXlsx(all()));
  $('qCsv').onclick = () => all().length && download(`الكميات-${stamp()}.csv`, X.qtyCsv(all()), 'text/csv;charset=utf-8');

  $('hist').addEventListener('input', e => {
    const id = e.target.dataset.note; if (!id) return;
    const x = all().find(v => v._id === id); if (!x) return;
    x.note = e.target.value; saveSoon(x); // بدون إعادة رسم حتى لا يضيع المؤشر
  });
  $('hist').addEventListener('click', e => {
    if (e.target.dataset.supedit) { supEdit = e.target.dataset.supedit; drawList(); const s = document.querySelector(`select[data-sup="${CSS.escape(supEdit)}"]`); if (s) s.focus(); return; }
    const a = e.target.dataset.approve, u = e.target.dataset.unapprove;
    if (!a && !u) return;
    const x = all().find(v => v._id === (a || u)); if (!x) return;
    x._approved = !!a; if (a) x._approvedAt = new Date().toISOString(); else delete x._approvedAt;
    persist(); if (cur && cur._id === x._id) { $('review').classList.remove('on'); cur = null; } drawList(); toast(a ? 'اعتُمدت الفاتورة وانتقلت إلى «المعتمدة»' : 'أُلغي الاعتماد وعادت إلى «المراجعة»');
  });
  $('hist').addEventListener('click', e => {
    const o = e.target.dataset.open, r = e.target.dataset.rm, ds = e.target.dataset.dismiss;
    if (o) open(o, true);
    if (r) {
      const h = all(), i = h.findIndex(x => x._id === r); if (i < 0) return;
      const [gone] = h.splice(i, 1); persist(); drawList();
      if (cur && cur._id === r) { $('review').classList.remove('on'); cur = null; }
      undo(`حُذفت فاتورة ${gone.vendor_name || ''}`, () => { h.splice(i, 0, gone); persist(); drawList(); });
    }
    if (ds) { delete live[ds]; drawList(); }
  });
  function undo(msg, fn) {
    const t = $('toast'); t.innerHTML = `${esc(msg)} <button id="undoBtn">تراجع</button>`; t.classList.add('on', 'act');
    $('undoBtn').onclick = () => { fn(); t.classList.remove('on', 'act'); };
    clearTimeout(toast.h); toast.h = setTimeout(() => t.classList.remove('on', 'act'), 5000);
  }
  $('xAll').onclick = () => all().length && download(`فواتير-${stamp()}.xlsx`, X.toXlsx(all(), ref && ref.vatCols));
  $('xVat').onclick = () => { const l = all().filter(x => x._approved); if (!l.length) return toast('لا توجد فواتير معتمدة'); if (!ref) toast('تنبيه: لم يُحمَّل ملف البيانات، فستكون SupID و ProjectID فارغة'); download(`VatReport-${stamp()}.xlsx`, X.vatXlsx(l, ref && ref.vatCols)); };
  $('cAll').onclick = () => all().length && download(`فواتير-${stamp()}.csv`, X.invoicesCsv(all()), 'text/csv;charset=utf-8');
  $('cItems').onclick = () => all().length && download(`بنود-${stamp()}.csv`, X.itemsCsv(all()), 'text/csv;charset=utf-8');
  $('clearAll').onclick = () => {
    const h = all(); if (!h.length) return; const keep = h.slice();
    h.length = 0; persist(); drawList(); $('review').classList.remove('on'); cur = null;
    undo('حُذفت كل الفواتير', () => { h.push(...keep); persist(); drawList(); });
  };

  // ---------- الأحداث ----------
  $('cam').onchange = e => { addFiles(e.target.files); e.target.value = ''; };
  $('pick').onchange = e => { addFiles(e.target.files); e.target.value = ''; };
  $('folder').onchange = e => { addFiles(e.target.files, true); e.target.value = ''; };
  if (!('webkitdirectory' in document.createElement('input'))) $('folderBtn').hidden = true;
  // تكبير الصورة: أزرار + و − ، أو اضغط الصورة لفتحها بملء الشاشة
  let zl = 1; const setZ = z => { zl = Math.max(1, Math.min(4, z)); $('pimg').style.width = (zl * 100) + '%'; $('zinfo').textContent = Math.round(zl * 100) + '%'; };
  $('zoomIn').onclick = () => setZ(zl + 0.5); $('zoomOut').onclick = () => setZ(zl - 0.5);
  $('zoom').onclick = () => lightbox(pages, pageIdx);
  $('pimg').onclick = () => lightbox(pages, pageIdx);
  $('prevP').onclick = () => { if (pageIdx > 0) { pageIdx--; showPage(); } };
  $('nextP').onclick = () => { if (pageIdx < pages.length - 1) { pageIdx++; showPage(); } };
  $('closeV').onclick = () => $('again').click();
  $('again').onclick = () => { $('review').classList.remove('on'); cur = null; drawList(); };
  const dz = $('drop');
  ['dragenter', 'dragover'].forEach(t => dz.addEventListener(t, e => { e.preventDefault(); dz.classList.add('over'); }));
  ['dragleave', 'drop'].forEach(t => dz.addEventListener(t, e => { e.preventDefault(); dz.classList.remove('over'); }));
  // سحب مجلد: نقرأ كل الملفات داخله (والمجلدات الفرعية)
  function readEntry(en, out) {
    if (en.isFile) return new Promise(ok => en.file(f => { f._path = en.fullPath.replace(/^\//, ''); out.push(f); ok(); }, () => ok()));
    if (!en.isDirectory) return Promise.resolve();
    const rd = en.createReader();
    return new Promise(ok => {
      const step = () => rd.readEntries(async ents => { if (!ents.length) return ok(); for (const c of ents) await readEntry(c, out); step(); }, () => ok());
      step();
    });
  }
  dz.addEventListener('drop', async e => {
    const items = [...(e.dataTransfer.items || [])].map(i => i.webkitGetAsEntry && i.webkitGetAsEntry()).filter(Boolean);
    if (items.some(en => en.isDirectory)) { const out = []; for (const en of items) await readEntry(en, out); addFiles(out, true); }
    else addFiles(e.dataTransfer.files);
  });
  document.addEventListener('paste', e => { const f = e.clipboardData?.files; if (f && f.length) addFiles(f); });
  // ---------- التبويبات ----------
  function setTab(t, keep) {
    if (t !== tab) $('apprMsg').hidden = true;
    tab = t;
    document.querySelectorAll('#tabs [data-tab]').forEach(b => { const on = b.dataset.tab === t; b.setAttribute('aria-selected', on); b.tabIndex = on ? 0 : -1; });
    document.querySelectorAll('.pane').forEach(p => p.classList.toggle('on', p.dataset.pane.split(' ').includes(t)));
    if (!keep && cur && (t === 'up' || t === 'exp' || t === 'qty' || (t === 'appr') !== !!cur._approved)) { $('review').classList.remove('on'); cur = null; }
    drawList();
    if (!keep && t === 'rev' && !cur) { const f = all().find(x => !x._approved); if (f) open(f._id); }
    try { sessionStorage.setItem('invoice-reader-tab', t); } catch { }
  }
  document.addEventListener('click', e => { const g = e.target.closest('[data-goto]'); if (g) { setTab(g.dataset.goto); window.scrollTo(0, 0); } });
  $('tabs').addEventListener('click', e => { const b = e.target.closest('[data-tab]'); if (b) { setTab(b.dataset.tab); window.scrollTo(0, 0); } });
  $('tabs').addEventListener('keydown', e => {
    const ts = [...document.querySelectorAll('#tabs [data-tab]')], i = ts.findIndex(b => b.dataset.tab === tab);
    const j = e.key === 'ArrowLeft' ? i + 1 : e.key === 'ArrowRight' ? i - 1 : -1; if (j < 0 || j >= ts.length) return;
    e.preventDefault(); setTab(ts[j].dataset.tab); ts[j].focus();
  });
  // أزرار شاشة المراجعة: اعتماد / إلغاء الاعتماد، حذف، التالي
  const nextOf = id => { const l = shownList(), i = l.findIndex(x => x._id === id); return l[i + 1] || l[i - 1] || null; };
  $('vAppr').onclick = () => {
    if (!cur) return; const x = cur, nx = nextOf(x._id);
    x._approved = !x._approved; if (x._approved) x._approvedAt = new Date().toISOString(); else delete x._approvedAt;
    persist(); toast(x._approved ? 'اعتُمدت الفاتورة' : 'أُلغي الاعتماد وعادت إلى «المراجعة»');
    $('review').classList.remove('on'); cur = null;
    if (nx) open(nx._id); else drawList();
  };
  $('vDel').onclick = () => {
    if (!cur) return; const h = all(), i = h.findIndex(x => x._id === cur._id); if (i < 0) return; const nx = nextOf(cur._id);
    const [gone] = h.splice(i, 1); persist(); $('review').classList.remove('on'); cur = null;
    if (nx) open(nx._id); else drawList();
    undo(`حُذفت فاتورة ${gone.vendor_name || ''}`, () => { h.splice(i, 0, gone); persist(); drawList(); });
  };
  $('vNext').onclick = () => { if (!cur) return; const l = shownList(), i = l.findIndex(x => x._id === cur._id); const nx = l[i + 1] || l[0]; if (nx && nx._id !== cur._id) open(nx._id); else toast('لا توجد فاتورة أخرى'); };
  // عارض الصورة بملء الشاشة (تكبير وتصغير)
  let lbPages = [], lbI = 0, lbZ = 1;
  function lbShow() {
    $('lbImg').src = lbPages[lbI] || ''; $('lbImg').style.width = (lbZ * 100) + '%';
    $('lbInfo').textContent = (lbPages.length > 1 ? `صفحة ${lbI + 1} من ${lbPages.length} · ` : '') + Math.round(lbZ * 100) + '%';
    $('lbPrev').hidden = $('lbNext').hidden = lbPages.length < 2;
  }
  function lightbox(p, i) {
    if (!p || !p.length) return toast('لا توجد صورة محفوظة لهذه الفاتورة');
    lbPages = p; lbI = i || 0; lbZ = 1; lbShow(); $('lb').hidden = false; document.body.classList.add('lbon'); $('lbClose').focus();
  }
  const lbClose = () => { $('lb').hidden = true; document.body.classList.remove('lbon'); };
  $('lbClose').onclick = lbClose;
  $('lbIn').onclick = () => { lbZ = Math.min(5, lbZ + 0.5); lbShow(); };
  $('lbOut').onclick = () => { lbZ = Math.max(1, lbZ - 0.5); lbShow(); };
  $('lbPrev').onclick = () => { if (lbI > 0) { lbI--; lbShow(); } };
  $('lbNext').onclick = () => { if (lbI < lbPages.length - 1) { lbI++; lbShow(); } };
  $('lbImg').ondblclick = () => { lbZ = lbZ > 1 ? 1 : 2.5; lbShow(); };
  document.addEventListener('keydown', e => { if (e.key === 'Escape' && !$('lb').hidden) lbClose(); });
  const viewImg = async id => lightbox(await pagesOf(id), 0);
  // جداول التبويبات: آخر عملية، المحمّلة سابقًا، التصدير
  const ST = x => x._approved ? `<span class="okb">معتمدة</span> ${expBadge(x)}` : '<span class="pendb">بانتظار المراجعة</span>';
  const dt = s => s ? new Date(s).toLocaleString('ar-SA-u-ca-gregory-nu-latn', { dateStyle: 'medium', timeStyle: 'short' }) : '';
  const invRow = (x, extra) => `<tr class="pick" data-row="${esc(x._id)}" title="اضغط لعرض ملف الفاتورة"><td data-l="الملف"><small class="fpath">📄 ${esc(x.file || '')}</small></td><td data-l="المورّد"><b>${esc(x._supName || x.vendor_name || '—')}</b></td><td data-l="رقم الفاتورة">${esc(x.invoice_number || '—')}</td><td data-l="التاريخ">${esc(x.invoice_date || '—')}</td><td data-l="الإجمالي" class="n">${esc(money(x))}</td><td data-l="الحالة">${extra || ST(x)}</td><td class="act"><button data-img="${esc(x._id)}">عرض الصورة</button><button data-go="${esc(x._id)}">${x._approved ? 'فتح' : 'مراجعة'}</button></td></tr>`;
  const HEAD = last => `<thead><tr><th>الملف</th><th>المورّد</th><th>رقم الفاتورة</th><th>التاريخ</th><th class="n">الإجمالي</th><th>${last || 'الحالة'}</th><th></th></tr></thead>`;
  function drawTabs() {
    const h = all(), byId = Object.fromEntries(h.map(x => [x._id, x]));
    const nRev = h.filter(x => !x._approved).length, nAppr = h.filter(x => x._approved).length, notExp = h.filter(x => x._approved && !isExported(x)), exp = h.filter(isExported);
    const notAll = h.filter(x => !isExported(x));
    $('bRev').textContent = nRev || ''; $('bAppr').textContent = notExp.length || ''; $('bExp').textContent = exp.length || '';
    // بطاقات الحالة: مرفوعة / غير معتمدة / معتمدة غير مصدّرة / مصدّرة
    const card = (k, n, l, t) => `<button class="stat s-${k}" data-goto="${t}"><b>${n}</b><span>${l}</span></button>`;
    const cards = card('all', h.length, 'كل الفواتير المرفوعة', 'up') + card('rev', nRev, 'مرفوعة لم تُعتمد', 'rev') + card('appr', notExp.length, 'معتمدة لم تُصدّر', 'appr') + card('exp', exp.length, 'تم تصديرها', 'exp');
    document.querySelectorAll('.stats').forEach(e => e.innerHTML = cards);
    // شريط تبويب «المعتمدة»: أزرار التصدير + تحذير
    $('apprBar').hidden = tab !== 'appr';
    $('needRef').hidden = !!ref || !(tab === 'rev' || tab === 'appr');
    $('xNew').textContent = `تصدير المعتمدة غير المصدّرة إلى Excel (${notExp.length})`;
    const warnList = (list, title) => list.length ? `<b>⚠ ${title}</b><ul>${list.slice(0, 50).map(x => `<li><b>${esc(x._supName || x.vendor_name || x.file || '—')}</b> · ${esc(x.invoice_number || 'بدون رقم')} · ${esc(money(x))}: ${whyNot(x).map(esc).join('؛ ')}</li>`).join('')}${list.length > 50 ? '<li>…</li>' : ''}</ul>` : '';
    $('apprWarn').innerHTML = warnList(notExp, `${notExp.length} فاتورة معتمدة لم تُصدّر بعد:`); $('apprWarn').hidden = !notExp.length;
    $('exWarn').innerHTML = warnList(notAll, `${notAll.length} فاتورة لم تُصدّر بعد، وهذه الأسباب:`); $('exWarn').hidden = !notAll.length;
    keepNote();
    if (tab === 'up') { // جداول تبويب التحميل تُرسم فقط عند ظهوره
    // آخر عملية: من سجل العملية، وإن لم يوجد (متصفح آخر) فآخر دفعة حسب وقت التحميل
    let R0 = run;
    if (!R0 || !R0.items.length) {
      const ts = h.map(x => x._at || '').filter(Boolean).sort(); const last = ts[ts.length - 1];
      if (last) {
        const t0 = Date.parse(last), grp = h.filter(x => x._at && (x._run ? x._run === (h.find(y => y._at === last) || {})._run : t0 - Date.parse(x._at) <= 10 * 60e3));
        R0 = { at: grp.map(x => x._at).sort()[0] || last, items: grp.sort((a, b) => (a._at < b._at ? -1 : 1)).map(x => ({ file: x.file || '', st: 'ok', msg: '', id: x._id })) };
      }
    }
    const items = R0 ? R0.items : [], ok = items.filter(i => i.st === 'ok'), dup = items.filter(i => i.st === 'dup'), err = items.filter(i => i.st === 'error');
    const pend = Object.values(live).filter(v => v.status === 'queued' || v.status === 'working');
    $('lastRun').hidden = false;
    $('lrCount').textContent = `(${items.length + pend.length})`;
    $('lrSum').innerHTML = R0 ? `<span>${esc(dt(R0.at))}</span> · <b class="okc">حُمّلت ${ok.length}</b>${dup.length ? ` · <b class="dupc">لم تُحمّل لأنها مكررة ${dup.length}</b>` : ''}${err.length ? ` · <b class="errc">تعذّرت قراءتها ${err.length}</b>` : ''}${pend.length ? ` · <b>قيد القراءة ${pend.length}</b>` : ''}` : '';
    const lrRows = items.map(i => {
      const x = i.id && byId[i.id];
      if (x) return invRow(x, `<span class="okc">حُمّلت</span> ${ST(x)}`);
      return `<tr class="pend ${i.st === 'dup' ? 'dup' : 'error'}"><td data-l="الملف"><small class="fpath">📄 ${esc(i.file)}</small></td><td colspan="5" data-l="النتيجة">${i.st === 'ok' ? 'حُمّلت ثم حُذفت من القائمة' : i.st === 'dup' ? '<b class="dupc">لم تُحمّل (مكررة)</b>: ' + esc(i.msg) : '<b class="errc">تعذّرت القراءة</b>: ' + esc(i.msg)}</td><td></td></tr>`;
    }).concat(pend.map(v => `<tr class="pend"><td data-l="الملف"><small class="fpath">📄 ${esc(v.file)}</small></td><td colspan="5">${v.status === 'working' ? 'جارٍ القراءة…' : 'بالانتظار'}</td><td></td></tr>`)).join('');
    $('lrTable').innerHTML = HEAD() + `<tbody>${lrRows || '<tr><td colspan="7" class="empty">لم تُحمّل أي فاتورة بعد.</td></tr>'}</tbody>`;
    // المحمّلة سابقًا
    const inRun = new Set(ok.map(i => i.id)), prev = h.filter(x => !inRun.has(x._id)).sort((a, b) => ((b._at || '') > (a._at || '') ? 1 : -1));
    $('pvCount').textContent = `(${prev.length})`;
    $('pvSum').textContent = prev.length ? `بانتظار المراجعة ${prev.filter(x => !x._approved).length} · معتمدة ${prev.filter(x => x._approved && !isExported(x)).length} · مُصدّرة ${prev.filter(isExported).length} · المجموع ${sumBy(prev)}` : '';
    $('pvTable').innerHTML = HEAD() + `<tbody>${prev.map(x => invRow(x)).join('') || '<tr><td colspan="7" class="empty">لا توجد فواتير محمّلة قبل آخر عملية.</td></tr>'}</tbody>`;
    }
    if (tab !== 'exp') return;
    storageInfo();
    // التصدير
    $('exNote').textContent = nRev ? `ملاحظة: ${nRev} فاتورة ما زالت بانتظار المراجعة ولن تُصدّر قبل اعتمادها.` : '';
    const revL = h.filter(x => !x._approved);
    $('naCount').textContent = `(${revL.length})`; $('neCount').textContent = `(${notExp.length})`; $('exCount').textContent = `(${exp.length})`;
    $('naTable').innerHTML = HEAD('السبب') + `<tbody>${revL.map(x => invRow(x, `${ST(x)}<div class="why">${whyNot(x).map(esc).join('<br>')}</div>`)).join('') || '<tr><td colspan="7" class="empty">لا توجد فواتير غير معتمدة ✓</td></tr>'}</tbody>`;
    $('neTable').innerHTML = HEAD('السبب') + `<tbody>${notExp.map(x => invRow(x, `${ST(x)}<div class="why">${whyNot(x).map(esc).join('<br>')}</div>`)).join('') || '<tr><td colspan="7" class="empty">لا توجد فواتير معتمدة غير مصدّرة ✓</td></tr>'}</tbody>`;
    $('exTable').innerHTML = HEAD('طريقة التصدير') + `<tbody>${exp.map(x => invRow(x, ST(x) + '<div class="why ok">' + [x._sent ? `Google Sheet (InvID ${esc(x._sent)})` : '', x._exportedAt ? `Excel · ${esc(dt(x._exportedAt))}` : ''].filter(Boolean).join('<br>') + '</div>')).join('') || '<tr><td colspan="7" class="empty">لم تُصدّر أي فاتورة بعد.</td></tr>'}</tbody>`;
  }
  document.querySelectorAll('.tabtable').forEach(t => t.addEventListener('click', e => {
    const im = e.target.dataset.img, go = e.target.dataset.go;
    if (im) return viewImg(im);
    if (go) return open(go, true);
    const r = e.target.closest('tr[data-row]'); if (r && !e.target.closest('button,input,select,a')) viewImg(r.dataset.row);
  }));
  $('xNew').onclick = () => {
    const l = all().filter(x => x._approved && !isExported(x));
    if (!l.length) return amsg(all().some(x => x._approved) ? 'كل الفواتير المعتمدة صُدّرت سابقًا (انظر «تم تصديرها» في تبويب «التصدير»). لإعادة تصدير الكل استخدم «تصدير الكل إلى Excel» هناك.' : 'لا توجد فواتير معتمدة بعد. اعتمد الفواتير من تبويب «المراجعة».', true);
    if (!ref) toast('تنبيه: لم يُحمَّل ملف البيانات، فستكون SupID و ProjectID فارغة');
    let ok = false; try { ok = download(`فواتير-معتمدة-${stamp()}.xlsx`, X.toXlsx(l, ref && ref.vatCols)); } catch (e) { console.error(e); }
    if (!ok) return amsg('تعذّر إنشاء ملف Excel، فلم تُنقل الفواتير إلى «تم تصديرها». حاول مرة أخرى.', true);
    const now = new Date().toISOString(); l.forEach(x => x._exportedAt = now); persist(); drawList();
    amsg(`صُدّرت ${l.length} فاتورة إلى Excel وانتقلت إلى «تم تصديرها» في تبويب «التصدير».`);
    undo(`صُدّرت ${l.length} فاتورة`, () => { l.forEach(x => delete x._exportedAt); persist(); drawList(); });
  };
  document.addEventListener('click', e => { const u = e.target.dataset.unexport; if (!u) return; const x = all().find(v => v._id === u); if (x) { delete x._exportedAt; persist(); drawList(); } });
  // ---------- الحفظ الدائم والنسخ الاحتياطي ----------
  const OLD_DAYS = 3;
  function keepNote() { // تذكير: معتمدة لم تُرسل ولم تُصدّر منذ أيام (قد تضيع إذا مسح المتصفح بياناته)
    const w = $('keepWarn'); if (!w) return;
    const old = all().filter(x => x._approved && !isExported(x) && Date.now() - (Date.parse(x._approvedAt || x._at || '') || Date.now()) > OLD_DAYS * 864e5);
    w.hidden = !old.length;
    if (old.length) w.innerHTML = `⚠ لديك <b>${old.length}</b> فاتورة معتمدة لم تُرسل إلى Google Sheet ولم تُصدّر منذ أكثر من ${OLD_DAYS} أيام. هي محفوظة في هذا المتصفح فقط، وقد يمسحها المتصفح (خصوصًا Safari على الآيفون إذا لم يُفتح الموقع أسبوعًا). <button data-goto="appr">أرسلها أو صدّرها الآن</button>`;
  }
  let persisted = false;
  async function askPersist() { // نطلب من المتصفح ألا يمسح بيانات التطبيق تلقائيًا
    try { if (!persisted && navigator.storage && navigator.storage.persist) persisted = (await navigator.storage.persisted()) || (await navigator.storage.persist()); } catch { }
    storageInfo();
  }
  async function storageInfo() {
    const e = $('bkInfo'); if (!e) return;
    let t = `${all().length} فاتورة محفوظة في هذا المتصفح`;
    try { if (navigator.storage && navigator.storage.estimate) { const q = await navigator.storage.estimate(); if (q && q.usage != null) t += ` (تشغل ${(q.usage / 1048576).toFixed(1)} ميغابايت)`; } } catch { }
    t += persisted ? '. الحفظ دائم: لن يمسحها المتصفح تلقائيًا، لكن خذ نسخة احتياطية قبل مسح بيانات المتصفح أو تغيير الجهاز.' : '. قد يمسح المتصفح هذه البيانات إذا امتلأ الجهاز أو لم يُفتح الموقع مدة طويلة، فأرسل المعتمدة إلى الشيت أو خذ نسخة احتياطية بانتظام.';
    if (S.mode === 'ls') t += ' تنبيه: هذا المتصفح لا يدعم قاعدة البيانات الداخلية، فالمساحة محدودة (نحو 5 ميغابايت).';
    e.textContent = t;
  }
  const blobToData = b => typeof b === 'string' ? Promise.resolve(b) : new Promise(ok => { const r = new FileReader(); r.onload = () => ok(r.result); r.onerror = () => ok(''); r.readAsDataURL(b); });
  $('bkSave').onclick = async () => {
    const h = all(), withImg = $('bkImg').checked;
    if (!h.length) return toast('لا توجد فواتير لحفظها');
    $('bkSave').disabled = true; status('جارٍ تجهيز النسخة الاحتياطية…');
    try {
      const out = { app: 'invoice-reader', v: 1, at: new Date().toISOString(), invoices: h, images: {} };
      if (withImg) for (const x of h) { const u = await IDB.get(x._id); if (u && u.length) out.images[x._id] = await Promise.all(u.map(blobToData)); }
      if (download(`نسخة-احتياطية-فواتير-${stamp()}.json`, JSON.stringify(out), 'application/json')) status(`نُزّلت نسخة احتياطية فيها ${h.length} فاتورة${withImg ? ' مع صورها' : ''}. احفظ الملف في مكان آمن (Drive مثلًا).`);
    } catch (e) { console.error(e); status('تعذّر تجهيز النسخة الاحتياطية' + (withImg ? '، جرّب بدون الصور' : ''), true); }
    finally { $('bkSave').disabled = false; }
  };
  $('bkFile').onchange = async e => {
    const f = e.target.files[0]; e.target.value = ''; if (!f) return;
    try {
      const o = JSON.parse(await f.text());
      if (!o || o.app !== 'invoice-reader' || !Array.isArray(o.invoices)) throw new Error('هذا ليس ملف نسخة احتياطية من قارئ الفواتير');
      await ready;
      const have = new Set(all().map(x => x._id)); let n = 0, ni = 0;
      for (const x of o.invoices) {
        if (!x || !x._id || have.has(x._id)) continue;
        cache.push(x); have.add(x._id); n++;
        const im = o.images && o.images[x._id];
        if (Array.isArray(im) && im.length) { const bl = await Promise.all(im.map(u => fetch(u).then(r => r.blob()).catch(() => u))); await IDB.put(x._id, bl); ni++; }
      }
      cache.sort((a, b) => String(b._at || '').localeCompare(String(a._at || '')));
      persist(); drawList(); storageInfo();
      toast(n ? `استُرجعت ${n} فاتورة${ni ? ` (${ni} منها مع صورها)` : ''}` : 'كل فواتير هذه النسخة موجودة أصلًا');
    } catch (err) { console.error(err); toast((err && err.message) || 'تعذّرت قراءة الملف'); }
  };
  // تبويب آخر حفظ أو حذف فواتير: نحدّث هذا التبويب بنفس التغييرات حتى لا يكتب أحدهما فوق الآخر
  if (bc) bc.onmessage = async e => {
    const m = e.data || {};
    try {
      if (m.ref) { ref = (await S.getKv('ref')) || null; refInfo(); }
      for (const id of m.put || []) {
        const x = await S.getInv(id); if (!x) continue;
        const i = cache.findIndex(v => v._id === id); if (i >= 0) cache[i] = x; else cache.unshift(x);
        snap.set(id, ser(x)); delete imgs[id];
        if (cur && cur._id === id) { cur = x; qr = x._qr || null; cur.src = cur.src || {}; cur.items = cur.items || []; render(); }
      }
      for (const id of m.del || []) {
        const i = cache.findIndex(v => v._id === id); if (i >= 0) cache.splice(i, 1); snap.delete(id);
        if (cur && cur._id === id) { $('review').classList.remove('on'); cur = null; }
      }
      drawList();
    } catch (err) { console.error(err); }
  };
  // ---------- بدء التشغيل: تحميل الفواتير والبيانات المرجعية من IndexedDB (ونقلها من localStorage أول مرة) ----------
  const ready = (async () => {
    try { await S.init(); cache = (await S.allInv()) || []; } catch (e) { saveFail(e); cache = []; }
    if (!cache.length && S.mode === 'idb') {
      try {
        const old = JSON.parse(localStorage.getItem(KEY) || '[]');
        if (Array.isArray(old) && old.length) { await S.putInv(old); cache = old; localStorage.removeItem(KEY); }
      } catch (e) { console.error('تعذّر نقل الفواتير القديمة', e); }
    }
    cache.sort((a, b) => String(b._at || '').localeCompare(String(a._at || '')));
    cache.forEach(x => snap.set(x._id, ser(x)));
    try {
      ref = (await S.getKv('ref')) || null;
      if (!ref) { const o = JSON.parse(localStorage.getItem(RKEY) || 'null'); if (o) { ref = o; await S.setKv('ref', o); localStorage.removeItem(RKEY); } }
    } catch (e) { console.error(e); }
  })();
  ready.then(() => {
    // حذف الصور المحفوظة لفواتير لم تعد موجودة
    IDB.keys().then(ks => { const have = new Set(all().map(x => x._id)); (ks || []).forEach(k => { if (!have.has(k)) IDB.del(k); }); });
    refInfo(); gsUi();
    // عند فتح التطبيق: إن كان مربوطًا بالشيت نحدّث البيانات المرجعية منه تلقائيًا (حتى تظهر خانات المشروع والمنتج على أي جهاز)
    if (gs) gsLoadRef(gs, true).catch(e => { if (!ref) amsg('تعذّر تحميل الموردين والمشاريع من Google Sheet: ' + ((e && e.message) || 'تحقق من الاتصال'), true); });
    let t0 = 'up'; try { t0 = sessionStorage.getItem('invoice-reader-tab') || 'up'; } catch { }
    setTab(t0);
    if (all().length) askPersist(); else storageInfo();
  });
  // ---------- العمل بدون إنترنت (Service Worker) وتجهيز ملفات القراءة مسبقًا ----------
  if ('serviceWorker' in navigator && (location.protocol === 'https:' || location.hostname === 'localhost')) {
    navigator.serviceWorker.register('sw.js').then(() => navigator.serviceWorker.ready).then(() => {
      // أول زيارة: نحمّل ملفات قراءة النص في الخلفية لتكون أول فاتورة أسرع (مرة واحدة، وليس مع توفير البيانات)
      const c = navigator.connection; if ((c && c.saveData) || localStorage.getItem('invoice-reader-prefetch-v1')) return;
      const go = () => Promise.all(['vendor/worker.min.js', 'vendor/core/tesseract-core-simd-lstm.wasm.js', 'vendor/lang/ara.traineddata.gz', 'vendor/lang/eng.traineddata.gz'].map(u => fetch(abs(u)).then(r => r.ok && r.blob())))
        .then(() => localStorage.setItem('invoice-reader-prefetch-v1', '1')).catch(() => { });
      (window.requestIdleCallback || (f => setTimeout(f, 3000)))(go);
    }).catch(e => console.warn('SW', e));
  }
  window.__invoiceApp = { addFiles, get cur() { return cur; }, get qr() { return qr; }, get raw() { return rawText; }, get busy() { return busy || queue.length > 0; }, all, dupMap, setTab, get tab() { return tab; } };
})();
