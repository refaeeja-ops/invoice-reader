/* قارئ الفواتير (النسخة المبسطة): قراءة الفواتير وتنزيل ملف Excel بالبيانات الأساسية. كل المعالجة داخل المتصفح ومجانية. */
(function () {
  'use strict';
  const P = window.InvoiceParse, X = window.InvoiceExport;
  const $ = id => document.getElementById(id);
  const esc = s => String(s ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const abs = p => new URL(p, location.href).href;
  const pathOf = f => f.relPath || f.webkitRelativePath || f.name || '';
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
    const add = (t, tag, lines) => { const r = P.parseInvoice(t, q), s = quality(r, q); runs.push({ t, r, s, lines: lines || [] }); return r; };
    const run = async (lang, pass) => {
      const w = await getWorker(lang);
      await w.setParameters({ tessedit_pageseg_mode: pass.psm });
      const img = prep(canvas, pass.size);
      const { data } = await w.recognize(img, {}, { text: true, blocks: true });
      const lines = ocrLines(data, img);
      let r = add(data.text, lang + ' ' + pass.size + '/' + pass.psm, lines);
      // أرقام هندية: نعيد قراءة الكلمات الرقمية بالمصنّف الخاص بها
      if (eastern(data.text) && window.EasternDigits) {
        try {
          await EasternDigits.load(abs('vendor/digits-model.json'));
          const rb = EasternDigits.rebuild(data, img);
          if (rb) { const r2 = add(rb.text, 'digits ' + lang + ' ' + pass.size, lines); if (quality(r2, q) >= quality(r, q)) r = r2; }
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
    return { best: runs[0].t, extra: runs.slice(1).map(x => x.t), lines: runs.map(x => x.lines) };
  }

  // كلمات Tesseract مع أماكنها (نسبة من الصفحة)، سطرًا سطرًا
  function ocrLines(data, img) {
    const out = [], W = img.width, H = img.height;
    try {
      (data.blocks || []).forEach(b => (b.paragraphs || []).forEach(p => (p.lines || []).forEach(l => {
        const ws = (l.words || []).filter(w => w.text && w.text.trim() && w.bbox).map(w => ({ s: w.text, x: w.bbox.x0 / W, y: w.bbox.y0 / H, w: (w.bbox.x1 - w.bbox.x0) / W, h: (w.bbox.y1 - w.bbox.y0) / H }));
        if (ws.length) out.push(ws);
      })));
    } catch (e) { }
    return out;
  }
  // نصوص PDF: نجمع العناصر في أسطر بحسب ارتفاعها
  function pdfLines(items) {
    const rows = [];
    (items || []).slice().sort((a, b) => a.y - b.y).forEach(it => {
      const r = rows.find(r => Math.abs(r.y - it.y) < Math.max(0.004, it.h * 0.5));
      if (r) r.ws.push(it); else rows.push({ y: it.y, ws: [it] });
    });
    return rows.map(r => r.ws.sort((a, b) => a.x - b.x));
  }
  // ---------- موضع القيمة المستخرجة على الفاتورة (للمقتطع تحت الخانة) ----------
  const LOC_KEYS = ['vendor_tax_number', 'invoice_number', 'total', 'invoice_date'];
  const LOC_KW = { vendor_tax_number: /ضريب|vat|tax|trn|tin/i, invoice_number: /فاتور|invoice|inv|رقم|no\b|#/i, total: /إجمالي|اجمالي|الاجمالي|total|المستحق|المجموع|الصافي|net|grand/i, invoice_date: /تاريخ|date/i };
  const dgt = s => P.normDigits(String(s ?? '')).replace(/\D/g, '');
  const aln = s => P.normDigits(String(s ?? '')).toLowerCase().replace(/[^0-9a-z]/g, '');
  function locTargets(k, v) {
    if (v === '' || v == null) return [];
    if (k === 'vendor_tax_number') { const d = dgt(v); return d.length >= 10 ? [d] : []; }
    if (k === 'invoice_number') { const a = aln(v); return a ? [a] : []; }
    if (k === 'total') { const n = +n2(v); if (!isFinite(n) || n <= 0) return []; const t = [dgt(n.toFixed(2))]; if (Number.isInteger(n)) t.push(String(n)); if (Math.round(n * 10) === n * 10) t.push(dgt(n.toFixed(1))); return t; }
    if (k === 'invoice_date') {
      const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(v); if (!m) return [];
      const [, y, mo, d] = m, a = [y + mo + d, d + mo + y, mo + d + y, d + mo + y.slice(2), y.slice(2) + mo + d, (+d) + '' + (+mo) + y, y + (+mo) + (+d), (+mo) + '' + (+d) + y];
      return [...new Set(a)];
    }
    return [];
  }
  function locHit(k, ts, span) {
    const d = dgt(span), a = aln(span);
    if (k === 'vendor_tax_number') return ts.some(t => d.includes(t));
    if (k === 'invoice_number') return ts.some(t => a === t || d === t || (t.length >= 4 && a.includes(t)));
    return ts.some(t => d === t); // المبلغ والتاريخ: تطابق تام للأرقام
  }
  // pagesLines: [[سطر من كلمات {s,x,y,w,h}] لكل صفحة] (قد تتعدد محاولات القراءة للصفحة الواحدة)
  function locate(k, v, pagesLines) {
    const ts = locTargets(k, v); if (!ts.length) return null;
    let best = null;
    pagesLines.forEach((sets, pg) => sets.forEach((lines, si) => lines.forEach((ws, li) => {
      const kw = LOC_KW[k].test(ws.map(w => w.s).join(' '));
      for (let i = 0; i < ws.length; i++) {
        let span = '';
        for (let j = i; j < Math.min(ws.length, i + 6); j++) {
          span += ' ' + ws[j].s;
          if (!locHit(k, ts, span)) continue;
          const sc = (kw ? 10 : 0) - (j - i) - si * 0.5 + (k === 'total' ? pg + li * 0.001 : -(pg + li * 0.001));
          if (!best || sc > best.sc) {
            const g = ws.slice(i, j + 1), x0 = Math.min(...g.map(w => w.x)), y0 = Math.min(...g.map(w => w.y));
            best = { sc, pg, x: x0, y: y0, w: Math.max(...g.map(w => w.x + w.w)) - x0, h: Math.max(...g.map(w => w.y + w.h)) - y0 };
          }
          break;
        }
      }
    })));
    if (!best || !(best.w > 0 && best.h > 0)) return null;
    const r = n => Math.round(n * 10000) / 10000;
    return { pg: best.pg, x: r(best.x), y: r(best.y), w: r(best.w), h: r(best.h) };
  }
  function locAll(x, pagesLines) {
    const loc = {};
    LOC_KEYS.forEach(k => { try { const r = locate(k, x[k], pagesLines); if (r) loc[k] = r; } catch (e) { } });
    return loc;
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
    const texts = [], extras = [], plines = [];
    for (let i = 0; i < list.length; i++) {
      const p = list[i];
      if ((p.text || '').replace(/\s/g, '').length >= 40) { texts.push(p.text); plines.push([pdfLines(p.items)]); continue; }
      say(list.length > 1 ? `جارٍ قراءة الصفحة ${i + 1} من ${list.length}…` : 'جارٍ قراءة نص الفاتورة…');
      const o = await ocrPage(p.canvas, q, say);
      texts.push(o.best); plines.push(o.lines || []); if (list.length === 1) extras.push(...o.extra);
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
    return { d, pages: list.map(p => toUrl(p.canvas)), src: list, plines };
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
    if (P.blockedVat(x.vendor_tax_number)) w.push('الرقم الضريبي ' + x.vendor_tax_number + ' لا يُقبل كرقم للمورد (رقمنا كمشترٍ). صحّحه إلى رقم المورد.');
    const T = +x.total, V = +x.tax_amount;
    if (!blank(x.total) && !blank(x.tax_amount) && isFinite(T) && isFinite(V)) {
      if (V >= T) w.push('مبلغ الضريبة أكبر من الإجمالي أو يساويه.');
      else if (T - V > 0 && Math.abs(V - (T - V) * 0.15) > Math.max(0.05, T * 0.002)) w.push(`الضريبة لا تساوي 15% (المتوقع تقريبًا ${money(Math.round(T / 1.15 * 0.15 * 100) / 100)}).`);
    }
    if (x.invoice_date && !/^\d{4}-\d{2}-\d{2}$/.test(x.invoice_date)) w.push('تاريخ الفاتورة غير واضح.');
    const tw = twinsOf(x);
    if (tw.length) w.push(`مكررة: نفس المورد ونفس رقم الفاتورة (${x.invoice_number}) في ${tw.length === 1 ? 'فاتورة أخرى' : tw.length + ' فواتير أخرى'}: ${tw.map(y => y.file || 'بدون اسم').join('، ')}`);
    return w;
  }
  // مفتاح المورد: أول 10 أرقام من الرقم الضريبي إن كان صحيحًا، وإلا الاسم بعد التوحيد
  const vNameKey = s => P.normDigits(String(s || '')).toLowerCase().replace(/[إأآا]/g, 'ا').replace(/ة/g, 'ه').replace(/ى/g, 'ي').replace(/[^\p{L}\p{N}]+/gu, '');
  function vKey(x) {
    const d = P.normDigits(String(x.vendor_tax_number ?? '')).replace(/\.0+$/, '').replace(/\D/g, '').slice(0, 10);
    if (/^3\d{9}$/.test(d) && !P.blockedVat(d)) return 'v:' + d;
    const n = vNameKey(x.vendor_name); return n ? 'n:' + n : '';
  }
  const invKey = v => P.normDigits(String(v ?? '')).toUpperCase().replace(/[\s\-_\/#.]+/g, '').replace(/^0+(?=.)/, '');
  const dupKey = x => { const v = vKey(x), n = invKey(x.invoice_number); return v && n ? v + '|' + n : ''; };
  const twinsOf = x => { const k = dupKey(x); if (!k) return []; const a = list.filter(y => dupKey(y) === k); return a.length < 2 || a.every(y => y.notDup === k) ? [] : a.filter(y => y !== x); };
  // مجموعات الفواتير المكررة (نفس المورد ونفس رقم الفاتورة)
  function dupGroups() {
    const m = new Map(); list.forEach(x => { const k = dupKey(x); if (k) (m.get(k) || m.set(k, []).get(k)).push(x); });
    return [...m.entries()].filter(([k, a]) => a.length > 1 && !a.every(y => y.notDup === k)).map(e => e[1]);
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
    return `<div class="f f-${k}${miss ? ' miss' : ''}${(k === 'vendor_tax_number' || k === 'vendor_name') && noSup(x) ? ' nosup' : ''}${k === 'payment_method' && blank(v) ? ' nopay' : ''}${rv.cur[x.id] === k && isOpen(x.id) ? ' cur' : ''}"><span class="fh"><label for="f-${esc(x.id)}-${k}">${label}</label>${badge}<button type="button" class="pkb${miss ? ' need' : ''}" data-pick="${esc(x.id)}" data-pk="${k}" aria-label="حدد مكان ${label} على الفاتورة" title="حدد مكان ${label} على الفاتورة">${PICK_ICON}</button></span>${input}</div>`;
  };
  // الموردون المسجلون في SuppliersT (أرقامهم المختصرة) من Google Sheet، لتنبيه المورد غير المسجل
  const SUPK = 'invoice-simple-sup-v1';
  let supVats = null; try { const o = JSON.parse(localStorage.getItem(SUPK) || 'null'); if (o && Array.isArray(o.v)) supVats = new Set(o.v); } catch (e) { }
  const vatShort = v => { const d = P.normDigits(String(v ?? '')).replace(/\.0+$/, '').replace(/\D/g, '').slice(0, 10); return /^3\d{9}$/.test(d) ? d : ''; };
  const noSup = x => P.blockedVat(x.vendor_tax_number) || !!(supVats && vatShort(x.vendor_tax_number) && !supVats.has(vatShort(x.vendor_tax_number)));
  const noSupHtml = x => P.blockedVat(x.vendor_tax_number) ? '' : noSup(x) ? `<p class="nosupline" role="note">⚠ الرقم الضريبي ${esc(vatShort(x.vendor_tax_number))} غير مسجل في قائمة الموردين في الإقرار: مورد جديد، سيُضاف عند النقل إلى الإقرار.</p>` : '';
  async function supSync() {
    if (!gs || !GURL) return;
    try {
      const r = await fetch(GURL + (GURL.includes('?') ? '&' : '?') + 'email=' + encodeURIComponent(gs.email) + '&token=' + encodeURIComponent(gs.token), { redirect: 'follow' });
      const j = await r.json(); if (!j.ok || !j.sheets || !j.sheets.SuppliersT) return;
      const t = j.sheets.SuppliersT, h = (t[0] || []).map(v => String(v).trim().toLowerCase()), c = h.indexOf('vatid'); if (c < 0) return;
      const v = [...new Set(t.slice(1).map(r => vatShort(r[c])).filter(Boolean))];
      supVats = new Set(v); try { localStorage.setItem(SUPK, JSON.stringify({ at: Date.now(), v })); } catch (e) { }
      draw();
    } catch (e) { console.warn(e); }
  }
  const imgCache = new Map();
  const loadPageImg = u => { if (!imgCache.has(u)) { if (imgCache.size > 40) imgCache.clear(); imgCache.set(u, loadImg(u).catch(() => null)); } return imgCache.get(u); };
  function card(x, tr, opt) {
    const hide = (opt && opt.hide) || [], showF = F.filter(([k]) => !hide.includes(k));
    const w = issues(x);
    if (tr) return `<article class="inv trash" id="c-${esc(x.id)}" aria-label="فاتورة محذوفة">
      <header><div><b>${esc(x.vendor_name || 'مورد غير معروف')}</b><small class="fp">📄 ${esc(x.file || '')}</small>${sentBadge(x)}</div><span class="amt">${esc(money(x.total))}</span></header>
      <fieldset disabled><div class="grid">${F.map(f => fieldHtml(x, f)).join('')}${projHtml(x)}
        <label class="f wide"><span>ملاحظات</span><textarea rows="2">${esc(x.note || '')}</textarea></label></div></fieldset>
      ${w.length ? `<ul class="warn">${w.map(t => `<li>${esc(t)}</li>`).join('')}</ul>` : ''}
      ${rvHtml(x)}</article>`;
    const dup = twinsOf(x).length > 0;
    return `<article class="inv${dup ? ' isdup' : ''}${opt && opt.odd ? ' odd' : ''}" id="c-${esc(x.id)}">
      <header><div>${hide.includes('vendor_name') ? '' : `<b>${esc(x.vendor_name || 'مورد غير معروف')}</b>`}<small class="fp">📄 ${esc(x.file || '')}</small>${dupBadge(x)}${hide.includes('sent') ? '' : sentBadge(x)}</div><span class="hr"><span class="amt">${esc(money(x.total))}</span>${verSeg(x)}</span></header>
      ${x.vNote ? `<p class="vnote">${esc(x.vNote)}</p>` : ''}
      <div class="grid">${showF.map(f => fieldHtml(x, f)).join('')}${hide.includes('project') ? '' : projHtml(x)}
        <label class="f wide"><span>ملاحظات</span><textarea data-id="${esc(x.id)}" data-k="note" rows="2" placeholder="اكتب أي إشكالية في الفاتورة">${esc(x.note || '')}</textarea></label></div>
      <div class="nosupbox">${noSupHtml(x)}</div>${w.length ? `<ul class="warn">${w.map(t => `<li>${esc(t)}</li>`).join('')}</ul>` : '<p class="okline">✓ البيانات الأساسية مكتملة</p>'}
      ${rrHtml(x)}${rvHtml(x)}
      <footer><button class="primary" data-rvw="${esc(x.id)}" data-rm="inv">🧾 مراجعة بيانات الفاتورة</button><button data-rvw="${esc(x.id)}" data-rm="ven">🏪 مراجعة بيانات المورد</button><button data-reread="${esc(x.id)}"${live[x.id] ? ' disabled' : ''}>${live[x.id] ? (live[x.id].status === 'working' ? 'جارٍ إعادة القراءة…' : 'بانتظار إعادة القراءة…') : '🔄 إعادة قراءة الفاتورة'}</button><button class="del" data-del="${esc(x.id)}"${delAsk === x.id ? ' hidden' : ''}>حذف الفاتورة</button></footer>
      ${delAsk === x.id ? `<div class="delask" role="alert"><label for="dr-${esc(x.id)}"><b>سبب حذف هذه الفاتورة (إلزامي)</b></label>
        <textarea id="dr-${esc(x.id)}" data-delwhy="${esc(x.id)}" rows="2" placeholder="مثال: فاتورة مكررة، أو ليست لنا"></textarea>
        <small>تنتقل الفاتورة إلى «الفواتير المحذوفة» أسفل الصفحة مع السبب، ويمكن استرجاعها.</small>
        <div class="zbtns"><button type="button" class="del" data-delok="${esc(x.id)}" disabled>احذف الفاتورة</button><button type="button" data-delno="">تراجع</button></div></div>` : ''}
    </article>`;
  }
  // ---------- التدقيق: مدقق / غير مدقق (اللون على الزر فقط) ----------
  const verSeg = x => `<span class="vseg" role="group" aria-label="حالة التدقيق"><button type="button" class="ok" data-ver="${esc(x.id)}" data-v="1" aria-pressed="${!!x.verified}">✓ مدقق</button><button type="button" class="no" data-ver="${esc(x.id)}" data-v="0" aria-pressed="${!x.verified}">✗ غير مدقق</button></span>`;
  // أي تعديل على فاتورة مدققة يرجعها «غير مدققة» مع إمكانية التراجع
  function unverify(x, undoFn) {
    if (!x || !x.verified) return;
    x.verified = false; x.vNote = 'عُدّلت بعد التدقيق، فرجعت «غير مدققة». دقّقها من جديد.'; save();
    undoBar('عدّلت فاتورة مدققة فرجعت غير مدققة', () => { if (undoFn) undoFn(); x.verified = true; x.vNote = ''; save(); draw(); });
  }
  function setVerified(x, v) {
    const prev = { verified: !!x.verified, vNote: x.vNote || '' }; if (prev.verified === v) return;
    x.verified = v; x.vNote = ''; save(); draw();
    const w = v ? issues(x) : [];
    undoBar(v ? (w.length ? '⚠ عُلّمت مدققة وفيها ملاحظات: ' + w[0] : 'عُلّمت الفاتورة مدققة') : 'عُلّمت الفاتورة غير مدققة', () => { Object.assign(x, prev); save(); draw(); });
    const n = document.querySelector(`[data-ver="${CSS.escape(x.id)}"][data-v="${v ? 1 : 0}"]`); if (n) n.focus({ preventScroll: true });
  }
  const dupBadge = x => twinsOf(x).length ? `<small class="db">⚠ مكررة: نفس المورد ونفس رقم الفاتورة</small>` : '';
  // ---------- التجميع (المورد، المشروع، طريقة الدفع، الشهر، التصدير) والترتيب ----------
  const VK = 'invoice-simple-view-v2';
  let vw = { group: '', sort: 'new', show: '' }; try { Object.assign(vw, JSON.parse(localStorage.getItem(VK) || '{}')); } catch { }
  const saveVw = () => { try { localStorage.setItem(VK, JSON.stringify(vw)); } catch { } };
  const vLabel = x => x.vendor_name || vatShort(x.vendor_tax_number) || 'مورد غير معروف';
  const cmpAr = (a, b) => String(a).localeCompare(String(b), 'ar', { numeric: true, sensitivity: 'base' });
  const MONTHS = ['يناير', 'فبراير', 'مارس', 'أبريل', 'مايو', 'يونيو', 'يوليو', 'أغسطس', 'سبتمبر', 'أكتوبر', 'نوفمبر', 'ديسمبر'];
  const dOf = x => /^\d{4}-\d{2}-\d{2}$/.test(x.invoice_date || '') ? x.invoice_date : '';
  const vendorName = k => { if (k === '?') return 'مورد غير معروف'; const c = new Map(); list.forEach(z => { if ((vKey(z) || '?') === k) { const n = vLabel(z); c.set(n, (c.get(n) || 0) + 1); } }); return [...c.entries()].sort((a, b) => b[1] - a[1])[0]?.[0] || 'مورد غير معروف'; };
  // لكل تجميع: مفتاح المجموعة، اسمها، والخانات التي تظهر في رأس المجموعة فلا تتكرر في كل فاتورة
  const GR = {
    vendor: { k: x => vKey(x) || '?', name: k => vendorName(k), meta: k => k.startsWith('v:') ? 'الرقم الضريبي ' + ((list.find(z => vKey(z) === k) || {}).vendor_tax_number || '') : '', hide: ['vendor_name', 'vendor_tax_number'], order: (a, b) => a === '?' ? 1 : b === '?' ? -1 : cmpAr(vendorName(a), vendorName(b)) },
    project: { k: x => x.project || '?', name: k => k === '?' ? 'بدون مشروع' : 'المشروع ' + k, hide: ['project'], order: (a, b) => a === '?' ? 1 : b === '?' ? -1 : cmpAr(a, b) },
    pay: { k: x => x.payment_method || '?', name: k => k === '?' ? 'بدون طريقة دفع' : k, hide: ['payment_method'], order: (a, b) => a === '?' ? 1 : b === '?' ? -1 : (PAYS.indexOf(a) + 1 || 99) - (PAYS.indexOf(b) + 1 || 99) || cmpAr(a, b) },
    month: { k: x => dOf(x).slice(0, 7) || '?', name: k => k === '?' ? 'بدون تاريخ' : MONTHS[+k.slice(5, 7) - 1] + ' ' + k.slice(0, 4), hide: [], order: (a, b) => a === '?' ? 1 : b === '?' ? -1 : cmpAr(b, a) },
    ver: { k: x => x.verified ? '1' : '0', name: k => k === '1' ? '✓ المدققة' : '✗ غير المدققة', hide: [], order: (a, b) => cmpAr(a, b) },
    sent: { k: x => sendState(x) === 'ok' ? '1' : '0', name: k => k === '1' ? '✓ المصدرة إلى الإقرار' : 'لم تُصدّر إلى الإقرار بعد', hide: ['sent'], order: (a, b) => cmpAr(a, b) },
  };
  if (vw.group && !GR[vw.group]) vw.group = '';
  const gClosed = new Set();
  function viewList() {
    let a = list.slice();
    if (vw.show === 'sent') a = a.filter(x => sendState(x) === 'ok');
    else if (vw.show === 'unsent') a = a.filter(x => sendState(x) !== 'ok');
    else if (vw.show === 'dup') a = a.filter(x => twinsOf(x).length);
    else if (vw.show === 'iss') a = a.filter(x => issues(x).length);
    else if (vw.show === 'ver') a = a.filter(x => x.verified);
    else if (vw.show === 'unver') a = a.filter(x => !x.verified);
    if (vw.sort === 'date') a.sort((p, q) => (dOf(q) ? 1 : 0) - (dOf(p) ? 1 : 0) || cmpAr(dOf(q), dOf(p)));
    else if (vw.sort === 'amount') a.sort((p, q) => (+q.total || 0) - (+p.total || 0));
    else if (vw.sort === 'no') a.sort((p, q) => cmpAr(invKey(p.invoice_number), invKey(q.invoice_number)));
    else if (vw.sort === 'old') a.reverse();
    // الفواتير المكررة (نفس المورد ونفس الرقم) تأتي متتالية: كل نسخة تلي أول ظهور لها
    const put = new Set(), out = [];
    a.forEach(x => { if (put.has(x)) return; put.add(x); out.push(x);
      if (twinsOf(x).length) { const k = dupKey(x); a.forEach(y => { if (!put.has(y) && dupKey(y) === k) { put.add(y); out.push(y); } }); } });
    return out;
  }
  const sumOf = a => ({ n: a.length, t: a.reduce((s, x) => s + (+x.total || 0), 0), v: a.reduce((s, x) => s + (+x.tax_amount || 0), 0) });
  const statsHtml = s => `<span>${s.n} فاتورة</span><span>الإجمالي <b>${esc(money(s.t))}</b></span><span>الضريبة <b>${esc(money(s.v))}</b></span>`;
  function drawBar() {
    const bar = $('lbar'), tot = $('ltot'); if (!bar) return; bar.hidden = list.length < 2 && !vw.show && !vw.group;
    const nd = list.filter(x => twinsOf(x).length).length, ns = list.filter(x => sendState(x) === 'ok').length, nv = list.filter(x => x.verified).length;
    const o = (v, t, cur) => `<option value="${esc(v)}"${v === cur ? ' selected' : ''}>${esc(t)}</option>`;
    const act = document.activeElement && ['vwGroup', 'vwSort', 'vwShow'].includes(document.activeElement.id) ? document.activeElement.id : '';
    bar.innerHTML = `<label><span>تجميع حسب</span><select id="vwGroup">${o('vendor', 'المورد', vw.group)}${o('project', 'المشروع', vw.group)}${o('pay', 'طريقة الدفع', vw.group)}${o('month', 'شهر الفاتورة', vw.group)}${o('sent', 'التصدير إلى الإقرار', vw.group)}${o('ver', 'التدقيق', vw.group)}${o('', 'الكل', vw.group)}</select></label>
      <label><span>${vw.group ? 'الترتيب داخل المجموعة' : 'الترتيب'}</span><select id="vwSort">${o('new', 'الأحدث إضافة', vw.sort)}${o('old', 'الأقدم إضافة', vw.sort)}${o('date', 'تاريخ الفاتورة (الأحدث)', vw.sort)}${o('amount', 'المبلغ (الأكبر)', vw.sort)}${o('no', 'رقم الفاتورة', vw.sort)}</select></label>
      <label><span>عرض</span><select id="vwShow">${o('', 'كل الفواتير', vw.show)}${o('dup', 'المكررة فقط (' + nd + ')', vw.show)}${GURL || ns ? o('sent', 'المصدرة إلى الإقرار (' + ns + ')', vw.show) + o('unsent', 'غير المصدرة إلى الإقرار (' + (list.length - ns) + ')', vw.show) : ''}${o('iss', 'التي فيها ملاحظات', vw.show)}${o('ver', 'المدققة فقط (' + nv + ')', vw.show)}${o('unver', 'غير المدققة فقط (' + (list.length - nv) + ')', vw.show)}</select></label>`;
    if (act) $(act).focus({ preventScroll: true });
    if (tot) { const a = viewList(), ds = a.map(dOf).filter(Boolean).sort(); tot.hidden = !list.length;
      tot.innerHTML = `<span>المعروض:</span>${statsHtml(sumOf(a))}${ds.length ? `<span>تاريخ أول فاتورة <b>${ds[0]}</b></span><span>تاريخ آخر فاتورة <b>${ds[ds.length - 1]}</b></span>` : ''}<span class="vg">✓ مدققة ${nv}</span><span class="vr">✗ غير مدققة ${list.length - nv}</span>`; }
  }
  // نسخ مطابقة تمامًا: نفس اسم الملف ونفس البيانات. نُبقي واحدة (المرسلة أو المعدّلة أولًا) وننقل الباقي إلى المحذوفات
  const copyKey = x => JSON.stringify([String(x.file || '').toLowerCase(), invKey(x.invoice_number), x.invoice_date || '', String(x.total ?? ''), vKey(x)]);
  function exactCopies() {
    const m = new Map(); list.forEach(x => { if (!x.file) return; const k = copyKey(x); (m.get(k) || m.set(k, []).get(k)).push(x); });
    const extra = [];
    m.forEach(a => { if (a.length < 2) return; const keep = a.slice().sort((p, q) => (q.sentAt ? 2 : 0) + (q.edited ? 1 : 0) - (p.sentAt ? 2 : 0) - (p.edited ? 1 : 0) || String(p.at).localeCompare(String(q.at)))[0]; a.forEach(x => { if (x !== keep) extra.push(x); }); });
    return extra;
  }
  let copyAsk = false;
  document.addEventListener('click', e => {
    const b = e.target.closest('button'); if (!b) return;
    if (b.id === 'cpDel') { copyAsk = true; drawDups(); const n = $('cpNo'); if (n) n.focus(); }
    if (b.id === 'cpNo') { copyAsk = false; drawDups(); }
    if (b.id === 'cpOk') {
      copyAsk = false; const ex = exactCopies(); if (!ex.length) return drawDups();
      const pos = ex.map(x => [x, list.indexOf(x)]), now = new Date().toISOString();
      ex.forEach(x => { list.splice(list.indexOf(x), 1); x.deletedAt = now; x.delReason = 'نسخة مكررة من نفس الملف (حُذفت تلقائيًا)'; trash.unshift(x); delete rrS[x.id]; });
      save(); saveTrash(); draw();
      undoBar(`نُقلت ${ex.length} نسخة مكررة إلى «الفواتير المحذوفة»`, () => { pos.sort((a, b) => a[1] - b[1]).forEach(([x, i]) => { trash.splice(trash.indexOf(x), 1); delete x.deletedAt; delete x.delReason; list.splice(i, 0, x); }); save(); saveTrash(); draw(); }, () => { });
    }
  });
  function drawDups() {
    const el = $('dupBox'); if (!el) return; const g = dupGroups(); el.hidden = !g.length; if (!g.length) return;
    const nc = exactCopies().length;
    const n = g.reduce((s, a) => s + a.length, 0);
    el.innerHTML = `<b>⚠ تنبيه: ${g.length === 1 ? 'توجد فواتير مكررة' : g.length + ' مجموعات من الفواتير المكررة'} (${n} فاتورة لها نفس المورد ونفس رقم الفاتورة)</b>
      <ul>${g.map(a => { const same = a.every(x => x.file && x.file === a[0].file), sp = !same && a.every(x => x.sig && x.sig === a[0].sig);
        return `<li><span><b>${esc(vLabel(a[0]))}</b> · فاتورة رقم <bdi>${esc(a[0].invoice_number)}</bdi> · ${a.length} مرات${same ? ' (نفس الملف أُضيف أكثر من مرة)' : ''}${sp ? ' (كلها من نفس الملف المقسّم: قد يكون الرقم مقروءًا خطأ في إحداها، أو تكون الصفحات من فاتورة واحدة)' : ''}:</span> ${a.map(x => `<button type="button" class="dupgo" data-go="${esc(x.id)}">📄 ${esc(x.file || 'بدون اسم')}${x.invoice_date ? ' · ' + esc(x.invoice_date) : ''}${!blank(x.total) ? ' · ' + esc(money(x.total)) : ''}</button>`).join('')}<button type="button" class="dupno" data-notdup="${esc(dupKey(a[0]))}">ليست مكررة</button></li>`; }).join('')}</ul>
      ${nc ? (copyAsk ? `<div class="cpask" role="alert"><b>نقل ${nc} نسخة مكررة إلى «الفواتير المحذوفة»؟</b> هذه نسخ من نفس الملف بنفس البيانات تمامًا. تبقى نسخة واحدة من كل فاتورة (المرسلة إلى الإقرار أو المعدّلة إن وجدت)، ويمكن استرجاع المحذوف.<div class="zbtns"><button type="button" class="del" id="cpOk">نعم، انقلها</button><button type="button" id="cpNo">تراجع</button></div></div>`
        : `<div class="zbtns"><button type="button" class="del" id="cpDel">🗑 حذف النسخ المطابقة تمامًا (${nc})</button><small>نفس الملف ونفس البيانات، أُضيف أكثر من مرة.</small></div>`) : ''}
      <div class="zbtns">${vw.show !== 'dup' ? '<button type="button" id="dupOnly">عرض المكررة فقط</button>' : ''}<small>راجعها واحذف المكرر (مع ذكر السبب)، أو صحّح رقم الفاتورة إن كان مقروءًا خطأ.</small></div>`;
  }
  document.addEventListener('change', e => {
    const t = e.target; if (!['vwSort', 'vwGroup', 'vwShow'].includes(t.id)) return;
    vw[{ vwSort: 'sort', vwGroup: 'group', vwShow: 'show' }[t.id]] = t.value; saveVw(); draw(); const n = $(t.id); if (n) n.focus();
  });
  document.addEventListener('click', e => {
    const b = e.target.closest('button'); if (!b) return;
    if (b.id === 'dupOnly') { vw.show = 'dup'; vw.group = 'vendor'; saveVw(); draw(); }
    if (b.dataset.grp) { const id = b.dataset.grp; gClosed.has(id) ? gClosed.delete(id) : gClosed.add(id); draw(); const n = document.querySelector(`[data-grp="${CSS.escape(id)}"]`); if (n) n.focus({ preventScroll: true }); }
    if (b.dataset.go) { let c = $('c-' + b.dataset.go);
      if (!c) { vw.show = ''; gClosed.clear(); saveVw(); draw(); c = $('c-' + b.dataset.go); }
      if (c) { c.scrollIntoView({ block: 'start', behavior: RM ? 'auto' : 'smooth' }); c.classList.add('flash'); setTimeout(() => c.classList.remove('flash'), 1600); } }
  });
  function listHtml() {
    const a = viewList();
    if (!a.length) return list.length ? '<p class="empty">لا توجد فواتير تطابق الاختيار الحالي.</p>' : '';
    if (!vw.group) return `<div class="flat">${a.map((x, i) => card(x, false, { odd: i % 2 })).join('')}</div>`;
    const G = GR[vw.group], m = new Map(); a.forEach(x => { const k = G.k(x); (m.get(k) || m.set(k, []).get(k)).push(x); });
    return [...m.keys()].sort(G.order).map(k => { const xs = m.get(k), id = vw.group + ':' + k, c = gClosed.has(id), meta = G.meta ? G.meta(k) : '';
      const hide = k === '?' || (vw.group === 'sent' && k === '0') ? [] : G.hide;
      return `<section class="grp${c ? ' closed' : ''}"><button type="button" class="gh" data-grp="${esc(id)}" aria-expanded="${!c}"><span class="ar" aria-hidden="true">${c ? '◀' : '▼'}</span><span class="gn">${esc(G.name(k))}</span>${meta ? `<span class="gm">${esc(meta)}</span>` : ''}<span class="gs">${statsHtml(sumOf(xs))}</span></button>
        <div class="gb">${c ? '' : xs.map((x, i) => card(x, false, { odd: i % 2, hide })).join('')}</div></section>`; }).join('');
  }
  function draw() {
    const dw = delAsk && $('dr-' + delAsk), keepWhy = dw ? dw.value : '';
    const pend = Object.values(live).filter(v => !v.reread);
    $('count').textContent = list.length ? `(${list.length})` : '';
    $('xl').disabled = !list.length; $('xl').textContent = `تنزيل ملف Excel${list.length ? ' (' + list.length + ')' : ''}`;
    const nIss = list.filter(x => issues(x).length).length;
    $('sum').textContent = list.length ? (nIss ? `${nIss} فاتورة فيها ملاحظات أو بيانات ناقصة` : 'كل الفواتير مكتملة') : '';
    $('list').innerHTML = pend.map(v => `<article class="inv pend"><b>📄 ${esc(v.file)}</b> <span>${v.status === 'working' ? 'جارٍ القراءة…' : v.status === 'split' ? 'بانتظار اختيار طريقة قراءة الملف' : 'بالانتظار'}</span></article>`).join('')
      + (listHtml() || (pend.length ? '' : '<p class="empty">لا توجد فواتير بعد. ارفع فاتورة أو صوّرها للبدء.</p>'));
    drawBar(); drawDups(); drawZones(); drawTrash(); drawWipe(); rvAllLabel(); sendLabel(); mountViewers();
    if (keepWhy) { const t = $('dr-' + delAsk), ok = document.querySelector(`[data-delok="${CSS.escape(delAsk)}"]`); if (t) t.value = keepWhy; if (ok) ok.disabled = !keepWhy.trim(); }
  }
  function drawTrash() {
    $('trSec').hidden = !trash.length; $('trCount').textContent = trash.length ? `(${trash.length})` : '';
    const d = s => { try { return new Date(s).toLocaleString('ar-SA-u-nu-latn', { dateStyle: 'medium', timeStyle: 'short' }); } catch { return ''; } };
    $('trList').innerHTML = trash.map(x => { const id = esc(x.id), o = !!trOpen[x.id];
      return `<li><div class="trrow"><span class="trn"><b>📄 ${esc(x.file || x.file_name || 'بدون اسم')}</b><small>${esc(x.vendor_name || 'مورد غير معروف')}${x.total !== '' && x.total != null ? ' · ' + esc(money(x.total)) : ''}${x.deletedAt ? ' · حُذفت ' + esc(d(x.deletedAt)) : ''}</small>${x.delReason ? `<span class="trwhy">سبب الحذف: ${esc(x.delReason)}</span>` : ''}${sentBadge(x)}</span>
        <button type="button" data-trview="${id}" aria-expanded="${o}">${o ? 'إغلاق' : 'مراجعة'}</button><button type="button" data-trback="${id}">↩ استرجاع</button><button type="button" class="del" data-trdel="${id}">حذف نهائي</button>
        ${trAsk === x.id ? `<div class="trask" role="alert"><b>حذف هذه الفاتورة نهائيًا؟</b>ستُحذف بياناتها وصورتها من المتصفح ولا تظهر هنا بعد ذلك.<div class="zbtns"><button type="button" class="del" data-trok="${id}">نعم، احذف نهائيًا</button><button type="button" data-trno="">تراجع</button></div></div>` : ''}</div>
        ${o ? card(x, true) : ''}</li>`; }).join('');
  }
  // تحديث ملاحظات وحالة بطاقة واحدة بدون إعادة رسم (حتى لا يضيع مكان الكتابة)
  function refreshCard(x, only) {
    const c = $('c-' + x.id); if (!c) return;
    c.classList.toggle('isdup', twinsOf(x).length > 0);
    { const hd = c.querySelector('header > div'), od = hd && hd.querySelector('.db'); if (od) od.remove(); if (hd) hd.querySelector('.fp').insertAdjacentHTML('afterend', dupBadge(x)); }
    if (!only) { list.forEach(y => { if (y !== x) refreshCard(y, true); }); drawDups(); }
    const w = issues(x), box = c.querySelector('.warn, .okline');
    const html = w.length ? `<ul class="warn">${w.map(t => `<li>${esc(t)}</li>`).join('')}</ul>` : '<p class="okline">✓ البيانات الأساسية مكتملة</p>';
    box.outerHTML = html;
    { const hb = c.querySelector('header b'); if (hb) hb.textContent = x.vendor_name || 'مورد غير معروف'; }
    c.querySelector('.amt').textContent = money(x.total);
    const hd = c.querySelector('header > div'); if (hd) { const ob = hd.querySelector('.sb'); if (ob) { ob.remove(); hd.insertAdjacentHTML('beforeend', sentBadge(x)); } } sendLabel(); drawBar();
    const ns = noSup(x), nb = c.querySelector('.nosupbox'); if (nb) nb.innerHTML = noSupHtml(x);
    F.forEach(([k]) => { const el = $(`f-${x.id}-${k}`); if (!el) return; const f = el.closest('.f'), m = blank(x[k]) && k !== 'payment_method';
      if (k === 'vendor_tax_number' || k === 'vendor_name') f.classList.toggle('nosup', ns);
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
    if (x.verified && String(x[k] ?? '') !== String(v ?? '')) { const old = x[k], oe = !!(x.edited && x.edited[k]);
      unverify(x, () => { x[k] = old; if (!oe && x.edited) delete x.edited[k]; });
      x[k] = v; (x.edited = x.edited || {})[k] = 1; if (x.zsrc) delete x.zsrc[k]; save(); draw(); const f = document.querySelector(`[data-id="${CSS.escape(id)}"][data-k="${CSS.escape(k)}"]`); if (f) { f.focus({ preventScroll: true }); try { f.setSelectionRange(f.value.length, f.value.length); } catch { } } return; }
    x[k] = v; (x.edited = x.edited || {})[k] = 1; if (x.zsrc) delete x.zsrc[k]; save(); refreshCard(x);
  });
  document.addEventListener('click', async e => {
    const b = e.target.closest('button'); if (!b) return;
    if (b.dataset.ver) { const x = list.find(y => y.id === b.dataset.ver); if (x) setVerified(x, b.dataset.v === '1'); return; }
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
    if (b.dataset.rvw) openReview(b.dataset.rvw, b.dataset.rm);
    if (b.dataset.pick) { const k = b.dataset.pk, m = MR.inv.ks.includes(k) ? 'inv' : MR.ven.ks.includes(k) ? 'ven' : ''; if (m) openReview(b.dataset.pick, m, k); else openPick(b.dataset.pick, k); }
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
    if (b.dataset.cz) { // تكبير: نفتح صورة الفاتورة في البطاقة على موضع الخانة
      const id = b.dataset.cz, k = b.dataset.ck; rv.open[id] = true; rv.cur[id] = k; draw();
      const h = $('rvh-' + id); if (h) { h.scrollIntoView({ block: 'center', behavior: RM ? 'auto' : 'smooth' }); setTimeout(() => { if (rvLive[id]) rvLive[id].focus(k); }, 350); }
    }
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
  const sig = f => String(pathOf(f)).trim().toLowerCase();
  const fkeyOf = f => String(f.name || '').trim().toLowerCase() + '|' + (f.size || 0); // نفس الملف حتى لو رُفع من مجلد آخر // منع التكرار: نفس الملف بنفس المسار (ملفان بنفس الاسم في مجلدين فرعيين مختلفين لا يُعدّان تكرارًا)
  // ملخص آخر عملية رفع: عدد الملفات، وما أُضيف، والمكرر، وغير المدعوم، ونتيجة القراءة
  let upSum = null;
  const isInv = f => /^image\//.test(f.type) || /\.(pdf|jpe?g|png|webp|heic|bmp|gif)$/i.test(f.name || '') || f.type === 'application/pdf';
  function addFiles(files, fromCam, kind) {
    // نقارن بما في القائمة وما في الانتظار وما يُقرأ الآن (ومنه ملف ينتظر اختيار طريقة التقسيم)
    const seen = new Set(list.flatMap(x => [x.sig, String(x.file || '').toLowerCase(), x.fkey]).filter(Boolean)
      .concat(queue.flatMap(j => [j.sig0 || sig(j.file), j.fkey0 || fkeyOf(j.file)]), Object.values(live).flatMap(v => [v.sig, v.fkey]).filter(Boolean)));
    const all = [...(files || [])].filter(f => !/(^|\/)\./.test(pathOf(f))); // نتجاهل الملفات والمجلدات المخفية مثل .DS_Store
    const b = { kind: fromCam ? 'cam' : kind || 'files', total: all.length, added: 0, dups: [], other: [], read: 0, failed: 0 };
    const sub = new Set(all.map(f => pathOf(f).split('/')).filter(a => a.length > 2).map(a => a.slice(0, -1).join('/')));
    b.subDirs = sub.size; b.subFiles = all.filter(f => pathOf(f).split('/').length > 2).length;
    all.forEach((f, i) => {
      if (!isInv(f)) { b.other.push(pathOf(f)); return; }
      const s = sig(f), fk = fkeyOf(f); if (!fromCam && (seen.has(s) || seen.has(fk))) { b.dups.push(pathOf(f)); return; } seen.add(s); seen.add(fk);
      const id = Date.now().toString(36) + '-' + i + '-' + Math.random().toString(36).slice(2, 6);
      queue.push({ id, file: f, batch: b }); live[id] = { status: 'queued', file: pathOf(f) || 'صورة من الكاميرا', sig: fromCam ? '' : s, fkey: fromCam ? '' : fk }; b.added++;
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
        ${b.subDirs ? `<li><span>منها في المجلدات الفرعية (${b.subDirs} مجلد فرعي)</span><b>${b.subFiles}</b></li>` : ''}
        <li class="${b.dups.length ? 'bad' : ''}"><span>لم تُضف لأنها مكررة (نفس الملف موجود في القائمة أو قيد القراءة)</span><b>${b.dups.length}</b></li>
        ${b.splitFiles ? `<li><span>ملفات قُسّمت إلى عدة فواتير</span><b>${b.splitFiles}</b></li>` : ''}
        ${b.skipped ? `<li><span>ملفات تم تخطيها</span><b>${b.skipped}</b></li>` : ''}
        ${b.other.length ? `<li><span>ليست صورة أو PDF فتم تجاهلها</span><b>${b.other.length}</b></li>` : ''}
        ${b.added ? `<li><span>${left > 0 ? 'قُرئت حتى الآن' : 'قُرئت'}</span><b>${b.read}${b.failed ? ` <small>(وتعذّرت قراءة ${b.failed})</small>` : ''}${left > 0 ? ` <small>من ${b.added}</small>` : ''}</b></li>` : ''}</ul>
      ${names('أسماء الملفات المكررة', b.dups)}${names('الملفات التي تم تجاهلها', b.other)}`;
  }
  document.addEventListener('click', e => { if (e.target.closest('#upSumX')) { upSum = null; drawSum(); } });
  async function reread(id) {
    const x = list.find(y => y.id === id); if (!x || live[id]) return;
    if (x.verified) { toast('هذه الفاتورة مدققة فلا تتغير. علّمها «غير مدقق» أولًا ثم أعد قراءتها.'); return; }
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
          const { d, pages, src, plines } = await processFile(job.file, say);
          let changed = 0;
          F.forEach(([k]) => { const nv = k === 'total' || k === 'tax_amount' ? (d[k] ?? '') : k === 'payment_method' ? payNorm(d[k]) : (d[k] || ''); if (x && !(x.edited && x.edited[k]) && !blank(nv) && nv !== x[k]) { x[k] = nv; if (x.zsrc) delete x.zsrc[k]; changed++; } });
          if (x && zoneFor(x)) { say('جارٍ تطبيق قالب المورد…'); try { changed += await applyZones(x, src, plines); } catch (e) { console.warn(e); } }
          if (x) x.loc = locAll(x, plines);
          if (x) { save(); delete rvPages[job.id]; IMG.get(job.id).then(r0 => IMG.put(job.id, { pages, blob: (r0 && r0.blob) || job.file })); }
          toast(changed ? `أُعيدت القراءة: تحدّثت ${changed} من الخانات` : 'أُعيدت القراءة ولم تتغير البيانات');
        } catch (e) { console.error(e); toast('تعذّرت إعادة القراءة'); }
        deep = false; delete live[job.id]; done++; draw(); continue;
      }
      // ملف PDF بأكثر من صفحة: نسأل المستخدم هل هو فاتورة واحدة أو كل صفحة فاتورة أو يحدد بنفسه
      if (!job.split && window.InvoiceSplit && InvoiceSplit.isPdf(job.file)) {
        job.split = true; live[job.id].status = 'split'; draw(); say('جارٍ فحص صفحات الملف…');
        const more = queue.filter(j => !j.reread && !j.split && InvoiceSplit.isPdf(j.file)).length;
        let r; try { r = await InvoiceSplit.ask(job.file, { more, mode: job.batch && job.batch.splitMode }); } catch (e) { console.warn(e); r = { files: [job.file] }; }
        if (!r) { delete live[job.id]; if (job.batch) { job.batch.skipped = (job.batch.skipped || 0) + 1; job.batch.added--; if (job.batch === upSum) drawSum(); } draw(); continue; }
        if (r.applyAll && job.batch) job.batch.splitMode = r.applyAll;
        if (r.files.length !== 1 || r.files[0] !== job.file) { // كل فاتورة تصبح ملفًا مستقلًا يُقرأ لوحده
          delete live[job.id];
          const subs = r.files.map((f, i) => ({ id: job.id + '-' + (i + 1), file: f, batch: job.batch, split: true, sig0: sig(job.file), fkey0: fkeyOf(job.file) }));
          subs.forEach(j => { live[j.id] = { status: 'queued', file: j.file.name, sig: j.sig0, fkey: j.fkey0 }; });
          if (job.batch) { job.batch.added += subs.length - 1; job.batch.splitFiles = (job.batch.splitFiles || 0) + 1; if (job.batch === upSum) drawSum(); }
          queue.unshift(...subs); draw(); continue;
        }
        live[job.id].status = 'working'; draw();
      }
      try {
        const { d, pages, src, plines } = await processFile(job.file, say);
        const x = { id: job.id, at: new Date().toISOString(), file: d.file, file_name: d.file_name, sig: job.sig0 || sig(job.file), fkey: job.fkey0 || fkeyOf(job.file),
          invoice_number: d.invoice_number || '', invoice_date: d.invoice_date || '', total: d.total ?? '', tax_amount: d.tax_amount ?? '',
          vendor_name: d.vendor_name || '', vendor_tax_number: d.vendor_tax_number || '', payment_method: payNorm(d.payment_method), note: '' };
        if (zoneFor(x)) { say('جارٍ تطبيق قالب المورد…'); try { await applyZones(x, src, plines); } catch (e) { console.warn(e); } }
        x.loc = locAll(x, plines);
        list.unshift(x); save(); if (job.batch) { job.batch.read++; if (job.batch === upSum) drawSum(); }
        IMG.put(job.id, { pages, blob: job.file.size < 25e6 ? job.file : null });
      } catch (e) {
        console.error(e);
        const x = { id: job.id, at: new Date().toISOString(), file: live[job.id].file, sig: job.sig0 || sig(job.file), fkey: job.fkey0 || fkeyOf(job.file), invoice_number: '', invoice_date: '', total: '', tax_amount: '', vendor_name: '', vendor_tax_number: '', payment_method: '', note: 'تعذّرت قراءة الملف تلقائيًا: أدخل البيانات يدويًا' };
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
  // السحب والإفلات: نقرأ المجلدات وكل المجلدات الفرعية تحتها
  drop.addEventListener('drop', async e => {
    const items = [...(e.dataTransfer.items || [])], ents = items.map(i => i.webkitGetAsEntry && i.webkitGetAsEntry()).filter(Boolean);
    if (!ents.some(en => en.isDirectory)) { addFiles(e.dataTransfer.files, false, 'drop'); return; }
    status('جارٍ قراءة المجلد والمجلدات الفرعية…');
    const out = [];
    const readAll = r => new Promise(res => { const acc = []; const step = () => r.readEntries(es => { if (!es.length) return res(acc); acc.push(...es); step(); }, () => res(acc)); step(); });
    const walk = async en => {
      if (en.isFile) { try { const f = await new Promise((res, rej) => en.file(res, rej)); Object.defineProperty(f, 'relPath', { value: en.fullPath.replace(/^\//, '') }); out.push(f); } catch { } }
      else if (en.isDirectory && !/^\./.test(en.name)) for (const c of await readAll(en.createReader())) await walk(c);
    };
    for (const en of ents) await walk(en);
    status(''); addFiles(out, false, 'dir');
  });

  // ---------- Excel ----------
  const XL = [['invoice_number', 'رقم الفاتورة', 18], ['invoice_date', 'تاريخ الفاتورة', 14], ['total', 'المبلغ الإجمالي', 16], ['tax_amount', 'مبلغ الضريبة', 14],
    ['vendor_name', 'اسم المورد', 32], ['vendor_tax_number', 'الرقم الضريبي', 20], ['payment_method', 'طريقة الدفع', 14], ['project', 'المشروع', 12], ['verified', 'التدقيق', 12], ['note', 'ملاحظات', 50], ['file', 'الملف', 30]];
  const xlCell = (x, k) => {
    if (k === 'note') return noteOf(x);
    if (k === 'verified') return x.verified ? 'مدقق' : 'غير مدقق';
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
  // نص بحروف متباعدة (ا س م ا ل م و ر د) يأتي أحيانًا من ملفات PDF: نجمع الحروف
  const unspace = s => { const tk = s.split(' ').filter(Boolean); return tk.length >= 6 && tk.filter(w => [...w].length === 1).length / tk.length > 0.6 ? tk.join('') : s; };
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
    return unspace(t.replace(/\n+/g, ' ')).replace(/\n+/g, ' ').replace(/^(اسم\s*)?(المورد|البائع|التاجر|الشركة|طريقة\s*الدفع|الدفع)\s*[:：-]?\s*/, '').replace(/^(vendor|seller|supplier|payment(\s*method)?)(\s*name)?\s*[:：-]?\s*/i, '').trim();
  }
  // ---------- القالب الذكي: النص المجاور للخانة (النص المرجعي) ----------
  // كلمات منطقة من الصفحة بأماكنها (نسبة من الصفحة): من طبقة نص PDF أو بقراءة الصورة
  function regionWords(src, r) {
    const R = { x: Math.max(0, r.x), y: Math.max(0, r.y) }; R.w = Math.min(1, r.x + r.w) - R.x; R.h = Math.min(1, r.y + r.h) - R.y;
    const inR = w => { const cx = w.x + w.w / 2, cy = w.y + w.h / 2; return cx >= R.x && cx <= R.x + R.w && cy >= R.y && cy <= R.y + R.h; };
    if (src.items && src.items.length) return Promise.resolve(src.items.filter(it => it.s && it.s.trim() && inR(it)).map(it => ({ s: it.s.trim(), x: it.x, y: it.y, w: it.w, h: it.h })));
    if (src.words && R.x <= 0.001 && R.w >= 0.998 && R.y <= 0.001 && R.h >= 0.998) return Promise.resolve(src.words);
    return serial(async () => {
      const C = src.canvas, W = C.width, H = C.height, sx = R.x * W, sy = R.y * H, sw = R.w * W, sh = R.h * H;
      if (sw < 8 || sh < 8) return [];
      const k = Math.min(2, 2400 / sw, 3000 / sh), m = 16;
      const c = document.createElement('canvas'); c.width = Math.round(sw * k) + 2 * m; c.height = Math.round(sh * k) + 2 * m;
      const g = c.getContext('2d'); g.fillStyle = '#fff'; g.fillRect(0, 0, c.width, c.height); g.imageSmoothingQuality = 'high'; g.drawImage(C, sx, sy, sw, sh, m, m, c.width - 2 * m, c.height - 2 * m);
      const w = await zoneWorker(); await w.setParameters({ tessedit_pageseg_mode: '11' });
      const { data } = await w.recognize(c, {}, { text: true, blocks: true });
      const out = [];
      (data.blocks || []).forEach(b => (b.paragraphs || []).forEach(p => (p.lines || []).forEach(l => (l.words || []).forEach(o => {
        if (!o.text || !o.text.trim() || !o.bbox) return;
        const X = sx + (o.bbox.x0 - m) / k, Y = sy + (o.bbox.y0 - m) / k;
        out.push({ s: o.text.trim(), x: X / W, y: Y / H, w: (o.bbox.x1 - o.bbox.x0) / k / W, h: (o.bbox.y1 - o.bbox.y0) / k / H });
      }))));
      if (R.x <= 0.001 && R.w >= 0.998 && R.y <= 0.001 && R.h >= 0.998) src.words = out;
      return out;
    });
  }
  const aNorm = s => P.normDigits(String(s).normalize('NFKC')).replace(/[إأآا]/g, 'ا').replace(/ة/g, 'ه').replace(/ى/g, 'ي').replace(/[^\p{L}\p{N}]+/gu, '').toLowerCase();
  const isLabW = w => /\p{L}/u.test(w.s) && !/\d{4,}/.test(P.normDigits(w.s));
  // الكلمات سطرًا سطرًا
  function wLines(ws) {
    const rows = [];
    ws.slice().sort((a, b) => a.y - b.y).forEach(w => { const r = rows.find(q => Math.abs((q.y + q.h / 2) - (w.y + w.h / 2)) < Math.max(q.h, w.h) * 0.6); if (r) r.ws.push(w); else rows.push({ y: w.y, h: w.h, ws: [w] }); });
    return rows.map(r => r.ws.sort((a, b) => b.x - a.x)); // من اليمين لليسار
  }
  const spanBox = g => { const x0 = Math.min(...g.map(w => w.x)), y0 = Math.min(...g.map(w => w.y)); return { t: g.map(w => w.s).join(' '), x: x0, y: y0, w: Math.max(...g.map(w => w.x + w.w)) - x0, h: Math.max(...g.map(w => w.y + w.h)) - y0 }; };
  // عبارات التسميات المتجاورة على السطر (كلمات متتالية فيها حروف)
  function labPhrases(ws) {
    const out = [];
    wLines(ws).forEach(line => { let cur = [];
      const flush = () => { if (cur.length) out.push(spanBox(cur)); cur = []; };
      line.forEach(w => { if (!isLabW(w)) return flush(); const p = cur[cur.length - 1]; if (p && Math.min(Math.abs(p.x - (w.x + w.w)), Math.abs(w.x - (p.x + p.w))) > Math.max(p.h, w.h) * 2.2) flush(); cur.push(w); });
      flush(); });
    return out;
  }
  // يختار النص المرجعي لمستطيل الخانة: أقرب تسمية على نفس السطر، وإلا فوقه. ar = نسبة الطول للعرض للصفحة
  // f = صندوق القيمة نفسها داخل المستطيل (إن عُرف)، لاختيار التسمية الموجودة على سطر القيمة لا على وسط المستطيل
  function pickAnchor(ws, r, ar, f) {
    const v = f || r, cyr = v.y + v.h / 2, inR = p => p.x + p.w / 2 > r.x && p.x + p.w / 2 < r.x + r.w && p.y + p.h / 2 > r.y && p.y + p.h / 2 < r.y + r.h;
    const ps = labPhrases(ws).filter(p => !inR(p) && aNorm(p.t).length >= 2);
    // التسمية العربية تكون يمين القيمة والإنجليزية يسارها: لا نربط بتسمية في الجهة الأخرى (مثل «السجل التجاري» يسار رقم الفاتورة)
    const tol = 0.01, side = p => /[\u0600-\u06FF\uFB50-\uFDFF\uFE70-\uFEFF]/.test(p.t) ? p.x >= v.x + v.w - tol : p.x + p.w <= v.x + tol;
    const same = ps.filter(p => Math.abs(p.y + p.h / 2 - cyr) < Math.max(v.h, p.h) * 0.6 && side(p));
    const above = ps.filter(p => p.y + p.h / 2 < v.y && p.y + p.h / 2 > v.y - Math.max(4 * v.h, 0.06) && p.x < r.x + r.w && p.x + p.w > r.x);
    const d = p => Math.hypot((p.x + p.w / 2 - (v.x + v.w / 2)), (p.y + p.h / 2 - cyr) * ar);
    const a = (same.length ? same : above).sort((p, q) => d(p) - d(q))[0];
    return a ? anchorFrom(a, r, ar) : null;
  }
  // النص المرجعي من تسمية محددة (تلقائيًا أو باختيار المستخدم). b = مكان التسمية للعرض فقط ولا يُحفظ
  function anchorFrom(a, r, ar) {
    if (!a || !(a.h > 0)) return null;
    const u = a.h * ar; // ارتفاع النص المرجعي بوحدة عرض الصفحة
    const n = v => +v.toFixed(3);
    const o = { t: a.t.slice(0, 80), y0: n(a.y), off: { x: n((r.x - a.x) / u), y: n((r.y - a.y) / a.h), w: n(r.w / u), h: n(r.h / a.h) } };
    Object.defineProperty(o, 'b', { value: { x: a.x, y: a.y, w: a.w, h: a.h }, enumerable: false });
    return o;
  }
  function anchorSim(a, b) {
    const toks = s => String(s).split(/\s+/).map(w => aNorm(w).replace(/^ال/, '')).filter(Boolean);
    const chr = (x, y) => { x = aNorm(x); y = aNorm(y); if (!x || !y) return 0; if (x === y) return 1; const m = Array.from({ length: x.length + 1 }, (_, i) => [i]); for (let j = 1; j <= y.length; j++) m[0][j] = j;
      for (let i = 1; i <= x.length; i++) for (let j = 1; j <= y.length; j++) m[i][j] = Math.min(m[i - 1][j] + 1, m[i][j - 1] + 1, m[i - 1][j - 1] + (x[i - 1] === y[j - 1] ? 0 : 1)); return 1 - m[x.length][y.length] / Math.max(x.length, y.length); };
    const ta = toks(a), tb = toks(b), rec = ta.length ? ta.filter(x => tb.some(y => chr(x, y) >= 0.8)).length / ta.length : 0;
    return Math.max(chr(a, b), rec * 0.95);
  }
  // يحدد مكان الخانة في صفحة جديدة بالبحث عن النص المرجعي. يعيد {r, line} أو null
  async function anchorRect(src, z) {
    if (!z || !z.a || !src || !src.canvas) return null;
    const ws = await regionWords(src, { x: 0, y: 0, w: 1, h: 1 }); if (!ws.length) return null;
    const ar = src.canvas.height / src.canvas.width;
    // كل العبارات المتتالية (حتى 7 كلمات) على كل سطر
    const cands = [];
    const L = aNorm(z.a.t).length * 1.4 + 3; // بعض ملفات PDF تكتب كل حرف وحده، فنحدّ العبارة بطول النص لا بعدد الكلمات
    wLines(ws).forEach(line => { for (let i = 0; i < line.length; i++) for (let j = i; j < Math.min(line.length, i + 40); j++) { const g = line.slice(i, j + 1); if (!g.every(isLabW) && j > i) break; if (!isLabW(line[i])) break; const sb = spanBox(g); if (j > i + 6 && aNorm(sb.t).length > L) break; cands.push({ ...sb, line }); } });
    let best = null;
    cands.forEach(c => { const s = anchorSim(z.a.t, c.t); if (s < 0.75) return; const sc = s - Math.abs(c.y - (z.a.y0 ?? c.y)) * 0.3; if (!best || sc > best.sc) best = { ...c, s, sc }; });
    if (!best) return null;
    const u = best.h * ar, o = z.a.off;
    const r = { pg: z.pg, x: best.x + o.x * u, y: best.y + o.y * best.h, w: o.w * u, h: o.h * best.h };
    r.x = Math.max(0, Math.min(0.99, r.x)); r.y = Math.max(0, Math.min(0.99, r.y)); r.w = Math.min(1 - r.x, r.w); r.h = Math.min(1 - r.y, r.h);
    return { r, line: best.line, anchor: best };
  }
  // إن قطع المستطيل قيمة في ملف PDF (حروف وأرقام متفرقة)، نوسّعه على نفس السطر ليشمل بقية القيمة الملاصقة
  function growLine(src, r) {
    if (!src.items || !src.items.length) return r;
    r = { ...r };
    for (let n = 0; n < 40; n++) {
      const cy = r.y + r.h / 2, it = src.items.find(i => i.s && i.s.trim() && Math.abs(i.y + i.h / 2 - cy) < Math.max(i.h, r.h) * 0.5 && i.x < r.x + r.w + i.h * 0.6 && i.x + i.w > r.x - i.h * 0.6 && (i.x < r.x - 1e-4 || i.x + i.w > r.x + r.w + 1e-4) && i.x + i.w / 2 > r.x - i.h && i.x + i.w / 2 < r.x + r.w + i.h);
      if (!it || isLabW({ s: it.s }) && /\p{L}{2,}/u.test(it.s)) break;
      const x0 = Math.min(r.x, it.x), x1 = Math.max(r.x + r.w, it.x + it.w); r.x = x0; r.w = x1 - x0;
    }
    return r;
  }
  // قراءة خانة بالقالب: بالنص المرجعي أولًا، ثم القيمة على نفس السطر، ثم المكان الثابت
  async function readZone(src, k, z) {
    try {
      const A = await anchorRect(src, z);
      if (A) {
        if (k !== 'vendor_name') A.r = growLine(src, A.r);
        let v = extract(k, await readRegion(src, A.r));
        if (!blank(v) && !(k === 'vendor_tax_number' && P.blockedVat(v))) return { v, r: A.r, how: 'anchor' };
        const cx = A.r.x + A.r.w / 2, near = A.line.filter(w => !A.anchor.t.includes(w.s)).sort((p, q) => Math.abs(p.x + p.w / 2 - cx) - Math.abs(q.x + q.w / 2 - cx));
        for (const w of near) { v = extract(k, w.s); if (!blank(v) && !(k === 'vendor_tax_number' && P.blockedVat(v))) return { v, r: { pg: z.pg, x: w.x, y: w.y, w: w.w, h: w.h }, how: 'line' }; }
      }
    } catch (e) { console.warn(e); }
    return { v: extract(k, await readRegion(src, z)), r: { ...z }, how: 'fixed' };
  }
  // تطبيق القالب على فاتورة مقروءة. يعيد عدد الخانات التي أُخذت من القالب.
  async function applyZones(x, src, plines) {
    const key = zoneFor(x); if (!key || !src || !src.length) return 0;
    // كلمات الصفحة المقروءة مسبقًا (حتى لا نعيد قراءة الصورة كاملة للبحث عن النص المرجعي)
    if (plines) src.forEach((s, i) => { if (!s.words && !(s.items && s.items.length) && plines[i] && plines[i][0]) s.words = plines[i][0].flat(); });
    let n = 0;
    for (const [k, r] of Object.entries(zones[key].fields || {})) {
      if (x.edited && x.edited[k]) continue;
      const i = pgIndex(r.pg, src.length); if (i < 0) continue;
      const o = await readZone(src[i], k, r), v = o.v;
      // قالب قديم بلا نص مرجعي: نضيفه تلقائيًا من هذه الفاتورة متى قُرئت القيمة من مكانها
      if (!r.a && !r.fx && o.how === 'fixed' && !blank(v) && src[i].canvas) try {
        const up = Math.max(4 * r.h, 0.06), a = pickAnchor(await regionWords(src[i], { x: 0, y: r.y - up, w: 1, h: r.h + up + r.h * 0.5 }), r, src[i].canvas.height / src[i].canvas.width);
        if (a) { r.a = a; saveZones(); zMark(key); }
      } catch (e) { console.warn(e); }
      if (!blank(v) && !(k === 'vendor_tax_number' && P.blockedVat(v))) { if (x[k] !== v) n++; x[k] = v; (x.zsrc = x.zsrc || {})[k] = 'tpl'; (x.zrect = x.zrect || {})[k] = { pg: r.pg, x: o.r.x, y: o.r.y, w: o.r.w, h: o.r.h }; }
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
    Object.assign(pk, { x, k, rec, pages: rec.pages, src: [], pdf: null, rect: null, mode: 'draw', last: document.activeElement, multi: '', link: '', cands: null });
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
  function closePick() { $('pk').hidden = true; document.body.style.overflow = ''; pk.tok++; pk.src = []; pk.pdf = null; $('pk').classList.remove('multi'); $('pkTile').hidden = true; $('mrHint').hidden = true; $('pkMk').innerHTML = ''; pk.link = ''; pk.cands = null; pk.multi = ''; }
  function showPg(i) {
    const n = pk.pages.length; pk.pg = Math.max(0, Math.min(n - 1, i));
    pk.rect = null; pkSel.hidden = true; $('pkRo').hidden = true; $('pkOk').disabled = true;
    $('pkPg').textContent = `صفحة ${pk.pg + 1} من ${n}`; $('pkPrev').hidden = $('pkNext').hidden = false;
    $('pkPrev').disabled = pk.pg === 0; $('pkNext').disabled = pk.pg === n - 1;
    const key = zoneFor(pk.x), z = key && zones[key].fields[pk.k], old = $('pkOld');
    old.hidden = !!pk.multi || !(z && pgIndex(z.pg, n) === pk.pg);
    if (pk.multi) { pk.link = pk.link && pk.d[pk.link].pi === pk.pg ? pk.link : ''; mrMarks(); }
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
  function setMode(m) { pk.mode = m; $('pkDraw').setAttribute('aria-pressed', m === 'draw'); $('pkMove').setAttribute('aria-pressed', m === 'move'); pkPad.classList.toggle('move', effMode() === 'move'); if (pk.multi) drawTile(); }
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
      if (r.w < 0.012 || r.h < 0.006) { if (pk.multi) { pkSel.hidden = true; mrTap(r); return; } if (pk.rect) showSel(pk.rect); else pkSel.hidden = true; return; }
      if (pk.multi) { pkSel.hidden = true; mrPicked(r); return; }
      pk.rect = { ...r, pg: pgFor(pk.k, pk.pg, pk.pages.length) };
      readPicked();
    };
    pkPad.addEventListener('pointerup', end); pkPad.addEventListener('pointercancel', end);
    pkPad.addEventListener('contextmenu', e => e.preventDefault());
  }
  async function readPicked() {
    const tok = ++pk.tok, r = pk.rect, k = pk.k;
    $('pkRo').hidden = false; $('pkRoS').textContent = 'جارٍ قراءة المنطقة…'; $('pkVal').value = ''; $('pkOk').disabled = true; $('pkRaw').textContent = ''; $('pkAnc').textContent = '';
    // النص المجاور للخانة: يُحفظ مع القالب ليجد البرنامج الخانة نسبةً له في الفواتير القادمة
    pk.anchor = { rect: r, a: null };
    pk.anchorP = (async () => { try {
      const src = await pageSrc(pk.pg), ar = src.canvas.height / src.canvas.width, up = Math.max(4 * r.h, 0.06);
      const ws = await regionWords(src, { x: 0, y: r.y - up, w: 1, h: r.h + up + r.h * 0.5 });
      const a = pickAnchor(ws, r, ar); if (pk.anchor && pk.anchor.rect === r) pk.anchor.a = a;
      if (tok === pk.tok) $('pkAnc').textContent = a ? `النص المجاور الذي سيُبحث عنه في الفواتير القادمة: «${a.t}»` : 'لم يُوجد نص مجاور، فسيُستخدم المكان كما هو.';
    } catch (e) { console.warn(e); } })();
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
  $('pkOk').onclick = async () => {
    const x = pk.x, k = pk.k, r = pk.rect; if (!x || !r) return;
    if (pk.anchorP && !(pk.anchor && pk.anchor.a)) { $('pkOk').disabled = true; await Promise.race([pk.anchorP, new Promise(res => setTimeout(res, 8000))]); $('pkOk').disabled = false; if (pk.x !== x || pk.rect !== r) return; }
    let v = $('pkVal').value.trim();
    if (k === 'total' || k === 'tax_amount') v = n2(v);
    if (k === 'vendor_tax_number') v = P.normDigits(v).replace(/\D/g, '');
    if (k === 'vendor_tax_number' && P.blockedVat(v)) { toast('هذا الرقم الضريبي لا يُقبل كرقم للمورد (رقمنا كمشترٍ). حدّد رقم المورد.'); $('pkVal').focus(); return; }
    if (k === 'invoice_date') v = P.findDate(v) || v;
    if (k === 'payment_method') v = payNorm(v);
    const oldKey = zoneFor(x);
    if (String(x[k] ?? '') !== String(v ?? '')) { const old = x[k], oe = !!(x.edited && x.edited[k]); unverify(x, () => { x[k] = old; if (!oe && x.edited) delete x.edited[k]; }); }
    x[k] = v; (x.edited = x.edited || {})[k] = 1;
    let saved = false;
    if (!$('pkSaveL').hidden && $('pkSave').checked) {
      const key = vendorKey(x) || oldKey;
      if (key) {
        if (oldKey && oldKey !== key && zones[oldKey]) { zones[key] = { ...zones[oldKey], fields: { ...zones[oldKey].fields, ...((zones[key] || {}).fields || {}) } }; delete zones[oldKey]; zMark(oldKey); }
        const t = zones[key] = zones[key] || { fields: {} };
        Object.assign(t, { name: x.vendor_name || t.name || '', vat: vatOf(x) || t.vat || '', nameKey: nameKey(x.vendor_name) || t.nameKey || '', at: new Date().toISOString() });
        const fld = t.fields[k] = { pg: r.pg, x: +r.x.toFixed(4), y: +r.y.toFixed(4), w: +r.w.toFixed(4), h: +r.h.toFixed(4) };
        if (pk.anchor && pk.anchor.rect === r && pk.anchor.a) fld.a = pk.anchor.a;
        saveZones(); zMark(key); saved = true;
        const n = list.filter(y => y !== x && zoneFor(y) === key).length;
        if (n && !(rrS[x.id] && (rrS[x.id].state === 'run' || rrS[x.id].state === 'pick'))) rrS[x.id] = { state: 'offer', key, ks: [...new Set([...((rrS[x.id] && rrS[x.id].state === 'offer' && rrS[x.id].ks) || []), k])] };
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

  // ---------- مراجعة بيانات الفاتورة / المورد: الملف كاملًا مع بلاطة عائمة تُسحب بحرية ----------
  const MR = {
    inv: { t: 'مراجعة بيانات الفاتورة', h: 'بيانات الإقرار', ks: ['invoice_number', 'invoice_date', 'total', 'tax_amount'] },
    ven: { t: 'مراجعة بيانات المورد', h: 'بيانات المورد', ks: ['vendor_name', 'vendor_tax_number'] },
  };
  const MRC = { invoice_number: '#1838b8', invoice_date: '#7a2fb8', total: '#0b6a35', tax_amount: '#b35a00', vendor_name: '#1838b8', vendor_tax_number: '#b3121f' };
  const MRL = { invoice_number: 'رقم الفاتورة', invoice_date: 'تاريخ الفاتورة', total: 'الإجمالي', tax_amount: 'الضريبة', vendor_name: 'اسم المورد', vendor_tax_number: 'الرقم الضريبي' };
  const MRN = { invoice_number: 'رقم فاتورة', invoice_date: 'تاريخًا', total: 'مبلغًا', tax_amount: 'مبلغًا', vendor_name: 'نصًا', vendor_tax_number: 'رقمًا ضريبيًا من 15 رقمًا' };
  const nice = s => unspace(String(s).normalize('NFKC').replace(/\s+/g, ' ').trim());
  const isAr = s => /[\u0600-\u06FF\uFB50-\uFDFF\uFE70-\uFEFF]/.test(s);
  const fmtV = (k, v) => blank(v) ? '' : (k === 'total' || k === 'tax_amount') && typeof n2(v) === 'number' ? n2(v).toFixed(2) : k === 'invoice_date' ? (P.findDate(String(v)) || String(v)) : String(v);
  const pkTile = $('pkTile'), pkMk = $('pkMk');
  function mrFind(id) { return list.find(y => y.id === id); }
  async function openReview(id, m, k0) {
    const x = mrFind(id); if (!x) return;
    const rec = await getRec(id, x);
    if (!rec || !rec.pages || !rec.pages.length) {
      toast('لا توجد صورة محفوظة لهذه الفاتورة على هذا الجهاز. أرفق ملفها من «مراجعة الفاتورة»');
      rv.open[id] = true; draw(); const h = $('rvh-' + id); if (h) h.scrollIntoView({ block: 'center', behavior: RM ? 'auto' : 'smooth' }); return; }
    const n = rec.pages.length;
    Object.assign(pk, { x, k: '', rec, pages: rec.pages, src: [], pdf: null, rect: null, mode: 'draw', last: document.activeElement, multi: m, link: '', cands: null, d: {} });
    MR[m].ks.forEach(k => { const zr = x.zrect && x.zrect[k]; pk.d[k] = { v: fmtV(k, x[k]), src: '', r: zr ? { ...zr } : null, pi: zr ? pgIndex(zr.pg, n) : -1 }; });
    pk.k = MR[m].ks.includes(k0) ? k0 : MR[m].ks.find(k => mrStat(k)[0] === 'bad') || '';
    $('pkT').textContent = MR[m].t;
    $('pk').classList.add('multi'); pkTile.hidden = false; pkTile.classList.remove('min');
    $('pk').hidden = false; document.body.style.overflow = 'hidden';
    setMode('draw'); drawTile(); tilePlace();
    const d = pk.k && pk.d[pk.k], first = MR[m].ks.map(k => pk.d[k].pi).find(i => i >= 0);
    await showPg(d && d.pi >= 0 ? d.pi : first >= 0 ? first : 0);
    if (d && d.r) mrShow(pk.k);
    pkVp.focus({ preventScroll: true });
  }
  function mrStat(k) {
    const d = pk.d[k], v = String(d.v ?? '').trim(), x = pk.x;
    if (d.busy) return ['', d.busy];
    if (k === 'vendor_tax_number' && v) { const s = P.normDigits(v).replace(/\D/g, ''); if (P.blockedVat(s)) return ['bad', '⚠ هذا رقمنا كمشترٍ، لا يُقبل']; if (!/^3\d{13}3$/.test(s)) return ['bad', '⚠ يجب أن يكون 15 رقمًا يبدأ وينتهي بـ 3']; }
    if (k === 'invoice_date' && v && !P.findDate(v)) return ['bad', '⚠ التاريخ غير واضح'];
    if ((k === 'total' || k === 'tax_amount') && v && typeof n2(v) !== 'number') return ['bad', '⚠ ليس رقمًا'];
    if (!v) return ['bad', 'فارغة، حدّد مكانها'];
    if (d.src === 'pick') return ['ok', '✓ قُرئت من المكان المحدد'];
    if (d.src === 'hand') return ['ok', '✎ كتابة يدوية'];
    if (x.zsrc && x.zsrc[k] === 'tpl') return ['', 'من قالب المورد'];
    if (x.edited && x.edited[k]) return ['', 'معدّلة سابقًا'];
    return ['', 'كما قرأها التطبيق'];
  }
  function mrLinkTxt(k) {
    const d = pk.d[k];
    if (pk.link === k) return pk.cands ? '👆 المس العنوان الصحيح على الفاتورة (المحاط بإطار رمادي)' : 'جارٍ قراءة نصوص الصفحة…';
    if (!d.picked) return '';
    if (d.aBusy) return 'جارٍ البحث عن العنوان المجاور…';
    return d.a ? `🔗 مرتبط بالعنوان «${nice(d.a.t)}»` : '📍 بلا عنوان، يُحفظ بمكانه في الصفحة';
  }
  function drawTile() {
    const m = pk.multi; if (!m) return;
    const ks = MR[m].ks, cur = pk.k, focus = !!cur && (pk.mode === 'draw' || pk.link === cur);
    pkTile.classList.toggle('focus', focus);
    const rows = ks.map(k => { const d = pk.d[k], [c, t] = mrStat(k), lk = mrLinkTxt(k);
      const lb = !d.picked || d.aBusy ? '' : pk.link === k ? `<button type="button" class="lkb" data-mnolink="${k}">بلا عنوان</button><button type="button" class="lkb" data-mlinkx="1">إلغاء</button>` : `<button type="button" class="lkb" data-mlink="${k}">✎ ${d.a ? 'غيّر العنوان' : 'اربط بعنوان'}</button>`;
      return `<div class="mr-r${cur === k ? ' on' : ''}" data-mk="${k}" style="--c:${MRC[k]}"><span class="sw"></span>
        <span class="mrlb"><label for="mr-${k}">${MRL[k]}</label> <span class="st ${c}">${esc(t)}</span></span>
        <button type="button" class="pick" data-mpick="${k}" aria-pressed="${cur === k}">${cur === k ? (pk.mode === 'draw' ? '✕ إيقاف' : '▭ ارسم') : d.r && d.pi >= 0 ? '↻ أعد التحديد' : '▭ حدّد'}</button>
        <input id="mr-${k}" data-min="${k}" value="${esc(d.v)}" dir="${k === 'vendor_name' ? 'rtl' : 'ltr'}" inputmode="${k === 'total' || k === 'tax_amount' ? 'decimal' : k === 'vendor_tax_number' ? 'numeric' : 'text'}" autocomplete="off">
        ${lk ? `<span class="lk">${esc(lk)} ${lb}</span>` : ''}</div>`; }).join('');
    const chips = `<div class="mr-ch">خانة أخرى: ${ks.filter(k => k !== cur).map(k => `<button type="button" data-mpick="${k}" style="--c:${MRC[k]}">${MRL[k]}${mrStat(k)[0] === 'bad' ? ' ⚠' : ''}</button>`).join('')}<button type="button" class="primary" data-mdone="1">✓ تم</button></div>`;
    let ck = '', bad = false;
    if (m === 'inv') { const t = n2(pk.d.total.v), v = n2(pk.d.tax_amount.v);
      if (typeof t === 'number' && typeof v === 'number' && t > 0 && v > 0) {
        const exp = t * 15 / 115, ok = Math.abs(exp - v) < 0.05, pre = Math.abs(t * 0.15 - v) < 0.05, sug = (t + v).toFixed(2);
        ck = ok ? `✓ الضريبة = 15% من المبلغ قبل الضريبة (${(t - v).toFixed(2)})`
          : pre ? `⚠ الإجمالي (${t.toFixed(2)}) هو المبلغ قبل الضريبة، والإقرار يحتاج الإجمالي شامل الضريبة = ${sug} <button type="button" class="lkb" data-mfix="${sug}">استخدم ${sug}</button>`
          : `⚠ الضريبة لا تساوي 15% من المبلغ قبل الضريبة (المتوقع ${exp.toFixed(2)}). راجع الإجمالي والضريبة.`;
        bad = !ok; } }
    else ck = 'الرقم الضريبي للمورد هو المكتوب بجانب «البائع» أو «المورد»، وليس رقمنا كمشترٍ.';
    const x = pk.x, canSave = !!(vendorKey(x) || zoneFor(x) || (m === 'ven' && (nameKey(pk.d.vendor_name.v).length >= 3 || /^3\d{13}3$/.test(P.normDigits(String(pk.d.vendor_tax_number.v)).replace(/\D/g, '')))));
    const nm = (m === 'ven' && String(pk.d.vendor_name.v).trim()) || x.vendor_name || 'هذا المورد';
    const kept = $('mrTpl') ? $('mrTpl').checked : true;
    $('mrH').textContent = MR[m].h;
    $('mrB').innerHTML = rows + chips; $('mrF').innerHTML = `<span class="chk${bad ? ' bad' : ''}">${ck}</span>
      ${canSave ? `<label class="pkchk"><input type="checkbox" id="mrTpl"${kept ? ' checked' : ''}><span>احفظ الأماكن المحددة لكل فواتير «${esc(nm)}»</span></label>` : '<small class="note">لن تُحفظ الأماكن قالبًا لأن اسم المورد ورقمه الضريبي غير معروفين.</small>'}
      <div class="pkbtns"><button type="button" class="primary" id="mrSave">حفظ التعديلات</button><button type="button" id="mrCancel">إلغاء</button></div>`;
    mrMarks(); mrHint(); tileClamp();
  }
  function mrHint() {
    const h = $('mrHint'), f = pk.link || (pk.mode === 'draw' && pk.k);
    h.hidden = !pk.multi || !f; if (h.hidden) return;
    h.textContent = pk.link ? (pk.cands ? `المس العنوان الصحيح لـ «${MRL[pk.link]}»` : 'جارٍ قراءة نصوص الصفحة…') : `ارسم مستطيلًا حول: ${MRL[pk.k]}`;
  }
  const pct = r => `left:${r.x * 100}%;top:${r.y * 100}%;width:${r.w * 100}%;height:${r.h * 100}%`;
  function mrMarks() {
    if (!pk.multi) { pkMk.innerHTML = ''; return; }
    let h = '';
    MR[pk.multi].ks.forEach(k => { const d = pk.d[k]; if (!d.r || d.pi !== pk.pg) return;
      h += `<div class="mk${pk.k === k ? ' cur' : ''}" style="--c:${MRC[k]};${pct(d.r)}"><i>${MRL[k]}</i></div>`;
      if (d.a && d.a.b && pk.link !== k) h += `<div class="mka" style="--c:${MRC[k]};${pct({ x: d.a.b.x - 0.004, y: d.a.b.y - 0.003, w: d.a.b.w + 0.008, h: d.a.b.h + 0.006 })}" title="العنوان المرتبط بـ ${MRL[k]}"><i>🔗</i></div>`; });
    if (pk.link && pk.cands) { const sel = pk.d[pk.link].a && pk.d[pk.link].a.b;
      h += pk.cands.map((p, i) => `<button type="button" class="mkc${sel && Math.abs(sel.x - p.x) < 1e-4 && Math.abs(sel.y - p.y) < 1e-4 ? ' sel' : ''}" data-mcand="${i}" style="--c:${MRC[pk.link]};${pct({ x: p.x - 0.004, y: p.y - 0.003, w: p.w + 0.008, h: p.h + 0.006 })}" aria-label="اجعل «${esc(nice(p.t))}» العنوان"></button>`).join(''); }
    pkMk.innerHTML = h;
  }
  function mrShow(k) { // نحرك العارض إلى مكان الخانة
    const d = pk.d[k]; if (!d || !d.r || d.pi !== pk.pg) return;
    const b = pkPaper.getBoundingClientRect(), v = pkVp.getBoundingClientRect();
    pkVp.scrollTo({ top: pkVp.scrollTop + b.top + d.r.y * b.height - v.top - 70, left: pkVp.scrollLeft + b.left + d.r.x * b.width - v.left - 40, behavior: RM ? 'auto' : 'smooth' });
  }
  async function mrGo(k) { const d = pk.d[k]; if (d.r && d.pi >= 0 && d.pi !== pk.pg) await showPg(d.pi); mrShow(k); }
  const union = ws => { const x0 = Math.min(...ws.map(w => w.x)), y0 = Math.min(...ws.map(w => w.y)); return { x: x0, y: y0, w: Math.max(...ws.map(w => w.x + w.w)) - x0, h: Math.max(...ws.map(w => w.y + w.h)) - y0 }; };
  const lineTxt = ws => wLines(ws).map(l => (isAr(l.map(w => w.s).join('')) ? l : l.slice().reverse()).map(w => w.s).join(' ')).join('\n');
  async function mrPicked(r) {
    const k = pk.k; if (pk.link) return;
    if (!k) { toast('اختر الخانة من البلاطة أولًا، ثم ارسم حولها'); return; }
    const tok = ++pk.tok, d = pk.d[k], prev = { r: d.r, pi: d.pi }, pi = pk.pg;
    const ok = v => !blank(v) && !(k === 'vendor_tax_number' && P.blockedVat(String(v)));
    let rr = { ...r, pg: pgFor(k, pi, pk.pages.length) };
    d.r = rr; d.pi = pi; d.busy = 'جارٍ قراءة المنطقة…'; drawTile();
    let v = '', fb = null, viaLab = null, src = null;
    try {
      src = await pageSrc(pi);
      const ws = await regionWords(src, rr); if (tok !== pk.tok) return;
      v = extract(k, lineTxt(ws));
      const vw = ws.filter(w => ok(extract(k, w.s))); if (vw.length) fb = union(vw);
      if (!ok(v) && !(src.items && src.items.length)) { v = extract(k, await readRegion(src, rr)); if (tok !== pk.tok) return; }
      if (!ok(v) && k !== 'vendor_name') { // حُدد العنوان فقط: نقرأ القيمة التي بجانبه على نفس السطر
        for (const L of labPhrases(ws)) {
          const band = await regionWords(src, { x: 0, y: L.y - L.h, w: 1, h: L.h * 3 }); if (tok !== pk.tok) return;
          const cy = L.y + L.h / 2, ar = isAr(L.t);
          const c = band.filter(w => Math.abs(w.y + w.h / 2 - cy) < Math.max(L.h, w.h) * 0.7 && (ar ? w.x + w.w <= L.x + 0.01 : w.x >= L.x + L.w - 0.01) && ok(extract(k, w.s)))
            .sort((p, q) => Math.abs(p.x + p.w / 2 - (L.x + L.w / 2)) - Math.abs(q.x + q.w / 2 - (L.x + L.w / 2)));
          if (c.length) { const w = c[0]; v = extract(k, w.s); viaLab = L; fb = w;
            rr = d.r = { pg: rr.pg, x: Math.max(0, w.x - 0.006), y: Math.max(0, w.y - 0.004), w: Math.min(1, w.w + 0.012), h: Math.min(1, w.h + 0.008) }; break; }
        }
      }
    } catch (e) { console.error(e); }
    if (tok !== pk.tok) return;
    d.busy = '';
    if (!ok(v)) { Object.assign(d, prev); drawTile();
      toast(k === 'vendor_tax_number' && P.blockedVat(String(v)) ? 'هذا رقمنا كمشترٍ ولا يُقبل رقمًا للمورد. حدّد رقم المورد.' : `لم أجد ${MRN[k]} داخل المستطيل. ارسمه حول القيمة نفسها أو حول عنوانها.`); return; }
    d.v = fmtV(k, v); d.src = 'pick'; d.picked = true; d.nolink = false; d.a = null;
    if (viaLab) toast(`حدّدت العنوان «${nice(viaLab.t)}»، فقرأت القيمة التي بجانبه`);
    // العنوان المجاور: يُحفظ مع القالب ليجد البرنامج الخانة نسبةً له في الفواتير القادمة
    const ar = src.canvas.height / src.canvas.width, at = d.at = (d.at || 0) + 1;
    if (viaLab) d.a = anchorFrom(viaLab, rr, ar);
    else { d.aBusy = true;
      d.aP = (async () => { try { const up = Math.max(4 * rr.h, 0.06); const band = await regionWords(src, { x: 0, y: rr.y - up, w: 1, h: rr.h + up + rr.h * 0.5 });
        if (d.at === at) d.a = pickAnchor(band, rr, ar, fb); } catch (e) { console.warn(e); } finally { if (d.at === at) { d.aBusy = false; d.aP = null; if (pk.multi) drawTile(); } } })(); }
    drawTile();
  }
  function mrTap(r) { // لمس الإطار المتقطع للعنوان المرتبط يفتح تغييره
    if (!pk.multi || pk.link) return;
    const k = MR[pk.multi].ks.find(k => { const d = pk.d[k], b = d.a && d.a.b; return b && d.pi === pk.pg && r.x >= b.x - 0.006 && r.x <= b.x + b.w + 0.006 && r.y >= b.y - 0.006 && r.y <= b.y + b.h + 0.006; });
    if (k) mrLink(k);
  }
  async function mrLink(k) {
    const d = pk.d[k]; if (!d.r) return;
    pk.link = k; pk.k = k; pk.cands = null; d.at = (d.at || 0) + 1; d.aBusy = false;
    if (d.pi !== pk.pg) await showPg(d.pi);
    drawTile();
    try { const src = await pageSrc(pk.pg), ws = await regionWords(src, { x: 0, y: 0, w: 1, h: 1 }); if (pk.link !== k) return;
      const r = d.r, inR = p => p.x + p.w / 2 > r.x && p.x + p.w / 2 < r.x + r.w && p.y + p.h / 2 > r.y && p.y + p.h / 2 < r.y + r.h;
      pk.cands = labPhrases(ws).filter(p => aNorm(p.t).length >= 2 && !inR(p)); pk.ar = src.canvas.height / src.canvas.width;
    } catch (e) { console.error(e); pk.cands = []; }
    drawTile();
  }
  pkMk.addEventListener('click', e => {
    const c = e.target.closest('[data-mcand]'); if (!c || !pk.link || !pk.cands) return;
    const k = pk.link, d = pk.d[k], p = pk.cands[+c.dataset.mcand];
    d.a = anchorFrom(p, d.r, pk.ar); d.nolink = false; pk.link = ''; pk.cands = null; drawTile();
    toast(`رُبطت «${MRL[k]}» بالعنوان «${nice(p.t)}»`);
  });
  pkTile.addEventListener('click', e => {
    const t = e.target, b = t.closest('button');
    if (b && b.dataset.mpick) { const k = b.dataset.mpick; pk.link = ''; pk.cands = null;
      if (pk.k === k && pk.mode === 'draw') pk.k = ''; else { pk.k = k; setMode('draw'); }
      drawTile(); if (pk.k) mrGo(k); return; }
    if (b && b.dataset.mdone) { pk.k = ''; pk.link = ''; pk.cands = null; drawTile(); return; }
    if (b && b.dataset.mlink) { mrLink(b.dataset.mlink); return; }
    if (b && b.dataset.mlinkx) { pk.link = ''; pk.cands = null; drawTile(); return; }
    if (b && b.dataset.mnolink) { const d = pk.d[b.dataset.mnolink]; d.a = null; d.nolink = true; d.at = (d.at || 0) + 1; pk.link = ''; pk.cands = null; drawTile(); toast('أُلغي الربط، سيُحفظ المكان ثابتًا في الصفحة'); return; }
    if (b && b.dataset.mfix) { pk.d.total.v = b.dataset.mfix; pk.d.total.src = 'hand'; drawTile(); return; }
    if (b && b.id === 'mrSave') { mrSave(); return; }
    if (b && b.id === 'mrCancel') { const x = pk.x, k0 = MR[pk.multi].ks[0]; closePick(); if (x) backTo(x.id, k0); return; }
    if (b && b.id === 'mrMin') { const m = pkTile.classList.toggle('min'); b.textContent = m ? '▴ إظهار' : '▾ تصغير'; b.setAttribute('aria-expanded', !m); tileClamp(); return; }
    if (b && b.id === 'mrDock') { tpos = null; tilePlace(); try { localStorage.removeItem(TPK); } catch (e) { } return; }
    const row = t.closest('[data-mk]'); if (row && !b && !t.closest('input,label')) { pk.k = row.dataset.mk; pk.link = ''; drawTile(); mrGo(pk.k); }
  });
  pkTile.addEventListener('input', e => { const i = e.target.closest('[data-min]'); if (!i) return; const k = i.dataset.min, d = pk.d[k]; d.v = i.value; d.src = 'hand';
    const st = i.parentElement.querySelector('.st'), [c, t] = mrStat(k); st.className = 'st ' + c; st.textContent = t; });
  pkTile.addEventListener('change', e => { if (e.target.closest('[data-min]')) drawTile(); });
  pkTile.addEventListener('focusin', e => { const i = e.target.closest('[data-min]'); if (i && pk.k !== i.dataset.min) { pk.k = i.dataset.min; pkTile.querySelectorAll('[data-mk]').forEach(r => r.classList.toggle('on', r.dataset.mk === pk.k)); mrMarks(); mrHint(); } });
  async function mrSave() {
    const x = pk.x, m = pk.multi, ks = MR[m].ks; if (!x) return;
    const bad = ks.filter(k => pk.d[k].src && mrStat(k)[0] === 'bad');
    if (bad.length) { toast('صحّح أولًا: ' + bad.map(k => MRL[k]).join('، ')); pk.k = bad[0]; drawTile(); return; }
    const pend = ks.map(k => pk.d[k].aP).filter(Boolean);
    if (pend.length) { $('mrSave').disabled = true; $('mrSave').textContent = 'جارٍ حفظ العناوين…'; await Promise.race([Promise.all(pend), new Promise(res => setTimeout(res, 8000))]); if (pk.x !== x) return; }
    const oldKey = zoneFor(x), tplOn = !!($('mrTpl') && $('mrTpl').checked);
    const changed = ks.filter(k => pk.d[k].src);
    const before = Object.fromEntries(ks.map(k => [k, x[k]])), bEd = { ...(x.edited || {}) };
    changed.forEach(k => { let v = String(pk.d[k].v).trim();
      if (k === 'total' || k === 'tax_amount') v = n2(v);
      if (k === 'vendor_tax_number') v = P.normDigits(v).replace(/\D/g, '');
      if (k === 'invoice_date') v = P.findDate(v) || v;
      x[k] = v; (x.edited = x.edited || {})[k] = 1; (x.zsrc = x.zsrc || {})[k] = 'pick'; });
    if (changed.some(k => String(before[k] ?? '') !== String(x[k] ?? ''))) unverify(x, () => { Object.assign(x, before); x.edited = bEd; });
    const picked = ks.filter(k => pk.d[k].picked && pk.d[k].r);
    picked.forEach(k => { const r = pk.d[k].r; (x.zrect = x.zrect || {})[k] = { pg: r.pg, x: r.x, y: r.y, w: r.w, h: r.h }; });
    let saved = 0;
    if (tplOn && picked.length) {
      const key = vendorKey(x) || oldKey;
      if (key) {
        if (oldKey && oldKey !== key && zones[oldKey]) { zones[key] = { ...zones[oldKey], fields: { ...zones[oldKey].fields, ...((zones[key] || {}).fields || {}) } }; delete zones[oldKey]; zMark(oldKey); }
        const t = zones[key] = zones[key] || { fields: {} };
        Object.assign(t, { name: x.vendor_name || t.name || '', vat: vatOf(x) || t.vat || '', nameKey: nameKey(x.vendor_name) || t.nameKey || '', at: new Date().toISOString() });
        picked.forEach(k => { const d = pk.d[k], r = d.r;
          const fld = t.fields[k] = { pg: r.pg, x: +r.x.toFixed(4), y: +r.y.toFixed(4), w: +r.w.toFixed(4), h: +r.h.toFixed(4) };
          if (d.a) fld.a = JSON.parse(JSON.stringify(d.a)); else if (d.nolink) fld.fx = 1;
          x.zsrc[k] = 'tpl'; saved++; });
        saveZones(); zMark(key);
        const n = list.filter(y => y !== x && zoneFor(y) === key).length;
        if (n && !(rrS[x.id] && (rrS[x.id].state === 'run' || rrS[x.id].state === 'pick'))) rrS[x.id] = { state: 'offer', key, ks: [...new Set([...((rrS[x.id] && rrS[x.id].state === 'offer' && rrS[x.id].ks) || []), ...picked])] };
      }
    }
    if (changed.length || picked.length) save();
    closePick(); draw(); backTo(x.id, (changed[0] || ks[0]));
    toast(!changed.length && !picked.length ? 'لم يتغير شيء' : saved ? `حُفظت التعديلات، وحُفظ ${saved === 1 ? 'مكان واحد' : saved === 2 ? 'مكانان' : saved + ' أماكن'} في قالب ${x.vendor_name || 'المورد'}` : 'حُفظت التعديلات');
  }
  // سحب البلاطة بحرية من شريط عنوانها؛ إفلاتها قرب أسفل الشاشة يعيدها لمكانها
  const TPK = 'invoice-simple-tile-v1';
  let tpos = null; try { tpos = JSON.parse(localStorage.getItem(TPK) || 'null'); } catch (e) { }
  function tilePlace() { pkTile.classList.toggle('free', !!tpos); $('mrDock').hidden = !tpos; if (!tpos) { pkTile.style.left = pkTile.style.top = ''; return; } tileClamp(); }
  function tileClamp() { if (!tpos || pkTile.hidden) return; const w = pkTile.offsetWidth, h = pkTile.offsetHeight;
    tpos.x = Math.max(4, Math.min(innerWidth - w - 4, tpos.x)); tpos.y = Math.max(4, Math.min(innerHeight - Math.min(h, 56) - 4, tpos.y));
    pkTile.style.left = tpos.x + 'px'; pkTile.style.top = tpos.y + 'px'; }
  {
    const th = $('mrTh'); let tg = null, raf = 0;
    th.addEventListener('pointerdown', e => { if (e.target.closest('button')) return; e.preventDefault(); th.setPointerCapture(e.pointerId);
      const r = pkTile.getBoundingClientRect(); tg = { dx: e.clientX - r.left, dy: e.clientY - r.top };
      if (!tpos) { tpos = { x: r.left, y: r.top }; tilePlace(); } pkTile.classList.add('drag'); });
    th.addEventListener('pointermove', e => { if (!tg) return; tpos.x = e.clientX - tg.dx; tpos.y = e.clientY - tg.dy; cancelAnimationFrame(raf); raf = requestAnimationFrame(tileClamp); });
    const te = () => { if (!tg) return; tg = null; pkTile.classList.remove('drag');
      if (tpos && tpos.y + pkTile.offsetHeight > innerHeight - 24) tpos = null;
      tilePlace(); try { tpos ? localStorage.setItem(TPK, JSON.stringify(tpos)) : localStorage.removeItem(TPK); } catch (e) { } };
    th.addEventListener('pointerup', te); th.addEventListener('pointercancel', te);
    th.addEventListener('dblclick', e => { if (!e.target.closest('button')) $('mrMin').click(); });
    addEventListener('resize', tileClamp);
  }
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
    if (!r && x.loc) r = x.loc[k];
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
    const vv = (k, v) => `<bdi dir="${k === 'vendor_name' ? 'auto' : 'ltr'}">${esc(shown(k, v))}</bdi>`;
    if (r.state === 'pick') {
      const on = r.changes.filter(c => c.on).length, byX = new Map(); r.changes.forEach(c => (byX.get(c.id) || byX.set(c.id, []).get(c.id)).push(c));
      const ks = [...new Set(r.changes.map(c => c.k))], kAll = k => r.changes.filter(c => c.k === k).every(c => c.on);
      return `<div class="rr rrpick" role="group" aria-label="التعديلات المقترحة"><b>أُعيدت قراءة فواتير ${vn}: ${r.changes.length} تعديل مقترح في ${byX.size} ${byX.size === 1 ? 'فاتورة' : 'فواتير'}</b>
        <p>علّم ما تريد تطبيقه فقط. التعديلات غير المعلّمة لا تتغير.</p>
        <div class="chips"><span>حسب الخانة:</span>${ks.map(k => `<button type="button" class="chip" data-rrk="${id}|${k}" aria-pressed="${kAll(k)}">${kAll(k) ? '✓ ' : ''}${esc(label(k))} (${r.changes.filter(c => c.k === k).length})</button>`).join('')}</div>
        <div class="row"><button type="button" data-rrall="${id}">تحديد الكل</button><button type="button" data-rrnone="${id}">إلغاء تحديد الكل</button><button type="button" data-rrrevall="${id}">👁 مراجعة واحدًا واحدًا</button></div>
        ${[...byX.entries()].map(([xid, cs]) => { const y = list.find(q => q.id === xid) || {}, all = cs.every(c => c.on);
          return `<div class="rrinv"><h4>📄 <a href="#c-${esc(xid)}">${esc(y.file || 'فاتورة')}</a> <small>${cs.length} تعديل</small><label><input type="checkbox" data-rrx="${id}|${esc(xid)}"${all ? ' checked' : ''}> كل تعديلات الفاتورة</label></h4>
            ${cs.map((c, i) => `<div class="rrch${c.on ? '' : ' off'}${i % 2 ? ' odd' : ''}"><input type="checkbox" data-rron="${id}|${esc(c.cid)}"${c.on ? ' checked' : ''} aria-label="${esc(label(c.k))}: ${esc(shown(c.k, c.from))} إلى ${esc(shown(c.k, c.to))}">
              <span><span class="rrlb">${esc(label(c.k))}:</span> <s>${vv(c.k, c.from)}</s> ← <ins>${vv(c.k, c.to)}</ins></span>
              <span class="row"><button type="button" class="sm" data-rrrev="${id}|${esc(c.cid)}">👁 مراجعة</button></span>
</div>`).join('')}
            ${r.skipped.filter(s => s.id === xid).map(s => `<div class="rrch sk">${esc(label(s.k))}: ${s.why}</div>`).join('')}</div>`; }).join('')}
        ${r.skipped.filter(s => !byX.has(s.id)).length ? `<ul>${r.skipped.filter(s => !byX.has(s.id)).map(c => `<li class="sk">فاتورة ${fn(c.id)}: ${esc(label(c.k))} ${c.why}</li>`).join('')}</ul>` : ''}
        <div class="row rrapply"><button type="button" class="primary" data-rrapply="${id}"${on ? '' : ' disabled'}>طبّق المحدد (${on} من ${r.changes.length})</button><button type="button" data-rrno="${id}">إغلاق بدون تطبيق</button></div></div>`;
    }
    const ap = r.applied || [];
    return `<div class="rr" role="status"><b>${ap.length ? `طُبّق ${ap.length} من ${r.changes.length} تعديل مقترح` : r.changes.length ? 'لم يُطبّق أي تعديل' : 'انتهت إعادة القراءة: لم تتغير أي خانة'}</b>
      ${ap.length || r.skipped.length ? `<ul>${ap.map(c => `<li><a href="#c-${esc(c.id)}">فاتورة ${fn(c.id)}</a>: ${esc(label(c.k))} <s>${vv(c.k, c.from)}</s> ← <ins>${vv(c.k, c.to)}</ins></li>`).join('')}
        ${r.skipped.map(c => `<li class="sk">فاتورة ${fn(c.id)}: ${esc(label(c.k))} ${c.why}</li>`).join('')}</ul>` : ''}
      <div class="row">${ap.length ? `<button type="button" data-rrundo="${id}">↩ تراجع عن التطبيق</button>` : ''}<button type="button" data-rrno="${id}">إغلاق</button></div></div>`;
  }
  async function rerunVendor(id) {
    const x = list.find(y => y.id === id), r0 = rrS[id]; if (!x || !r0 || r0.state !== 'offer') return;
    const key = r0.key, T = zones[key]; if (!T) return;
    const others = list.filter(y => y !== x && zoneFor(y) === key);
    const r = rrS[id] = { state: 'run', key, n: others.length, done: 0, changes: [], skipped: [], same: [], before: [] }; draw();
    for (const y of others) {
      if (y.verified) { // المدققة لا تتغير أبدًا: نقرأ اقتراح القالب ونعرضه فقط
        try { const rec = await getRec(y.id, y);
          if (rec && rec.pages && rec.pages.length) { const c = { rec, pages: rec.pages, src: [], pdf: null };
            for (const [k, z] of Object.entries(T.fields || {})) { const i = pgIndex(z.pg, c.pages.length); if (i < 0) continue;
              const v = (await readZone(await loadSrc(c, i), k, z)).v;
              if (!blank(v) && String(v) !== String(y[k] ?? '')) r.skipped.push({ id: y.id, k, why: '🔒 لم يتغير لأن الفاتورة مدققة. اقترح القالب: ' + esc(shown(k, v)) }); }
            if (c.pdf) try { c.pdf.destroy(); } catch { } } } catch (e) { console.warn(e); }
        r.done++; if (rrS[id] === r) draw(); continue; }
      r.before.push({ id: y.id, vals: Object.fromEntries(F.map(([k]) => [k, y[k]])), zsrc: { ...(y.zsrc || {}) }, zrect: JSON.parse(JSON.stringify(y.zrect || {})) });
      try {
        const rec = await getRec(y.id, y);
        if (!rec || !rec.pages || !rec.pages.length) r.skipped.push({ id: y.id, k: Object.keys(T.fields)[0], why: 'لا توجد صورة محفوظة لها' });
        else {
          const c = { rec, pages: rec.pages, src: [], pdf: null };
          for (const [k, z] of Object.entries(T.fields || {})) {
            if (y.edited && y.edited[k]) { r.skipped.push({ id: y.id, k, why: 'لم يتغير لأنك عدّلته بيدك' }); continue; }
            const i = pgIndex(z.pg, c.pages.length); if (i < 0) continue;
            const o = await readZone(await loadSrc(c, i), k, z), v = o.v;
            if (blank(v)) { r.skipped.push({ id: y.id, k, why: 'لم تُقرأ قيمة واضحة في المكان المحفوظ، فبقيت كما هي' }); continue; }
            // اقتراح فقط: لا يتغير شيء حتى يختار المستخدم التعديلات ويضغط «طبّق المحدد»
            if (String(v) !== String(y[k] ?? '')) r.changes.push({ cid: y.id + ':' + k, id: y.id, k, from: y[k], to: v, z: { pg: z.pg, x: o.r.x, y: o.r.y, w: o.r.w, h: o.r.h }, pg: i, on: true });
            else r.same.push({ id: y.id, k, z: { pg: z.pg, x: o.r.x, y: o.r.y, w: o.r.w, h: o.r.h } });
          }
          if (c.pdf) try { c.pdf.destroy(); } catch { }
        }
      } catch (e) { console.warn(e); r.skipped.push({ id: y.id, k: Object.keys(T.fields)[0], why: 'تعذّرت قراءتها' }); }
      r.done++; if (rrS[id] === r) draw();
    }
    // الخانات التي قرأها القالب بنفس قيمتها: نحفظ مكانها فقط (لا تغيير في القيم)
    r.same.forEach(o => { const y = list.find(q => q.id === o.id); if (!y) return; (y.zrect = y.zrect || {})[o.k] = o.z; (y.zsrc = y.zsrc || {})[o.k] = y.zsrc[o.k] === 'new' ? 'new' : 'tpl'; });
    r.state = r.changes.length ? 'pick' : 'done'; save(); draw();
    toast(r.changes.length ? `اقتُرح ${r.changes.length} تعديلًا: اختر ما تريد تطبيقه` : 'انتهت إعادة القراءة ولم تتغير البيانات');
  }
  function applyPicked(id) {
    const r = rrS[id]; if (!r || r.state !== 'pick') return;
    const on = r.changes.filter(c => c.on);
    on.forEach(c => { const y = list.find(q => q.id === c.id); if (!y) return; y[c.k] = c.to; (y.zsrc = y.zsrc || {})[c.k] = 'new'; (y.zrect = y.zrect || {})[c.k] = c.z; });
    r.applied = on; r.state = 'done'; save(); draw();
    toast(on.length ? `طُبّق ${on.length} من ${r.changes.length} تعديل` : 'لم يُطبّق أي تعديل');
  }
  // مقتطع خانة من صفحة الفاتورة
  async function cropChange(cv, c, maxW) {
    const y = list.find(q => q.id === c.id); if (!y) return;
    const pages = await rvLoadPages(y), u = pages[c.pg]; if (!u) return;
    const im = await loadPageImg(u); if (!im || !cv.isConnected) return;
    const W = im.naturalWidth, H = im.naturalHeight, b = c.z, px = 0.012, py = 0.006;
    const sx = Math.max(0, b.x - px) * W, sy = Math.max(0, b.y - py) * H, sw = Math.min(W - sx, (b.w + 2 * px) * W), sh = Math.min(H - sy, (b.h + 2 * py) * H);
    if (sw < 2 || sh < 2) return;
    const sc = Math.min(3, maxW / sw); cv.width = Math.round(sw * sc); cv.height = Math.round(sh * sc);
    const g = cv.getContext('2d'); g.imageSmoothingQuality = 'high'; g.drawImage(im, sx, sy, sw, sh, 0, 0, cv.width, cv.height); return true;
  }
  // نافذة مراجعة كل تعديل مع ملفه
  let rvw = null;
  async function openRvw(id, i) {
    const r = rrS[id]; if (!r || r.state !== 'pick') return; const c = r.changes[i]; if (!c) return;
    rvw = { id, i }; const y = list.find(q => q.id === c.id) || {};
    $('rwT').textContent = label(c.k); $('rwS').textContent = '📄 ' + (y.file || '');
    $('rwO').textContent = shown(c.k, c.from); $('rwN').textContent = shown(c.k, c.to);
    $('rwA').setAttribute('aria-pressed', c.on); $('rwR').setAttribute('aria-pressed', !c.on);
    $('rwPos').textContent = `${i + 1} من ${r.changes.length}`; $('rwPrev').disabled = i === 0; $('rwNext').textContent = i === r.changes.length - 1 ? 'إنهاء المراجعة' : 'التالي ←';
    const h = $('rwH'); h.style.left = c.z.x * 100 + '%'; h.style.top = c.z.y * 100 + '%'; h.style.width = c.z.w * 100 + '%'; h.style.height = c.z.h * 100 + '%';
    const img = $('rwImg'), cv = $('rwC'); img.removeAttribute('src'); img.parentNode.hidden = true; cv.parentNode.hidden = true;
    $('rw').hidden = false; document.body.style.overflow = 'hidden'; $('rwA').focus();
    cv.width = 1; cv.height = 1; cropChange(cv, c, 640).then(ok => { if (ok && rvw && rvw.i === i) cv.parentNode.hidden = false; });
    const pages = await rvLoadPages(y); if (rvw && rvw.i === i && pages[c.pg]) { img.onload = () => { img.parentNode.hidden = false; }; img.src = pages[c.pg]; }
  }
  const closeRvw = () => { $('rw').hidden = true; document.body.style.overflow = ''; const id = rvw && rvw.id; rvw = null; draw(); const b = id && document.querySelector(`[data-rrall="${CSS.escape(id)}"]`); if (b) b.scrollIntoView({ block: 'center' }); };
  const stepRvw = d => { const r = rvw && rrS[rvw.id]; if (!r) return closeRvw(); const n = rvw.i + d; if (n >= r.changes.length) return closeRvw(); openRvw(rvw.id, Math.max(0, n)); };
  $('rwA').onclick = () => { const r = rrS[rvw.id]; r.changes[rvw.i].on = true; stepRvw(1); };
  $('rwR').onclick = () => { const r = rrS[rvw.id]; r.changes[rvw.i].on = false; stepRvw(1); };
  $('rwPrev').onclick = () => stepRvw(-1); $('rwNext').onclick = () => stepRvw(1); $('rwX').onclick = closeRvw;
  $('rw').addEventListener('click', e => { if (e.target === $('rw')) closeRvw(); });
  document.addEventListener('keydown', e => { if (e.key === 'Escape' && !$('rw').hidden) closeRvw(); });
  const rrFocus = sel => { const n = document.querySelector(sel); if (n) n.focus({ preventScroll: true }); };
  document.addEventListener('change', e => {
    const t = e.target, d = t.dataset;
    if (d.rron) { const [id, cid] = d.rron.split('|'), r = rrS[id], c = r && r.changes.find(q => q.cid === cid); if (!c) return; c.on = t.checked; draw(); rrFocus(`[data-rron="${CSS.escape(d.rron)}"]`); }
    if (d.rrx) { const [id, xid] = d.rrx.split('|'), r = rrS[id]; if (!r) return; r.changes.filter(c => c.id === xid).forEach(c => c.on = t.checked); draw(); rrFocus(`[data-rrx="${CSS.escape(d.rrx)}"]`); }
  });
  document.addEventListener('click', e => {
    const b = e.target.closest('button'); if (!b) return; const d = b.dataset;
    if (d.rrk) { const [id, k] = d.rrk.split('|'), r = rrS[id]; if (!r) return; const a = r.changes.filter(c => c.k === k), v = !a.every(c => c.on); a.forEach(c => c.on = v); draw(); rrFocus(`[data-rrk="${CSS.escape(d.rrk)}"]`); }
    if (d.rrall) { const r = rrS[d.rrall]; if (r) { r.changes.forEach(c => c.on = true); draw(); rrFocus(`[data-rrall="${CSS.escape(d.rrall)}"]`); } }
    if (d.rrnone) { const r = rrS[d.rrnone]; if (r) { r.changes.forEach(c => c.on = false); draw(); rrFocus(`[data-rrnone="${CSS.escape(d.rrnone)}"]`); } }
    if (d.rrrev) { const [id, cid] = d.rrrev.split('|'), r = rrS[id]; if (r) openRvw(id, r.changes.findIndex(c => c.cid === cid)); }
    if (d.rrrevall) openRvw(d.rrrevall, 0);
    if (d.rrapply) applyPicked(d.rrapply);
    if (d.notdup) { const k = d.notdup; list.forEach(y => { if (dupKey(y) === k) y.notDup = k; }); save(); draw(); toast('عُلّمت «ليست مكررة». إن تغيّر الرقم أو المورد يعود الفحص تلقائيًا'); }
  });
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
    let j; try { j = await r.json(); } catch { throw new Error('رد غير متوقع من الإقرار. تأكد أن Apps Script حُدّث بالنسخة الجديدة'); }
    if (!j.ok) throw new Error(j.error || 'تعذّر الاتصال بالإقرار');
    if (!j.zones) throw new Error('الإقرار لا يدعم القوالب بعد: حدّث كود Apps Script وانشر نسخة جديدة');
    return j;
  }
  function zMerge(j) { // يدمج القوالب: الأحدث يفوز، والتعديلات المحلية غير المرسلة لها الأولوية
    const put = {}, R = j.zones || {}, Dl = j.deleted || {};
    for (const k in Dl) if (zones[k] && !zpend[k] && (zones[k].at || '') <= (Dl[k] || '')) delete zones[k];
    for (const k in R) { if (zpend[k]) continue; if (!zones[k] || (R[k].at || '') > (zones[k].at || '')) {
      const old = zones[k]; zones[k] = R[k];
      if (old) for (const f in R[k].fields || {}) { const a = old.fields && old.fields[f], b = R[k].fields[f]; if (a && a.a && !b.a && Math.abs(a.x - b.x) < 1e-3 && Math.abs(a.y - b.y) < 1e-3) b.a = a.a; } } }
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
    } catch (e) { console.warn(e); zst.err = (e && e.message) || 'تعذّر الاتصال بالإقرار'; }
    zst.busy = false; draw();
    if (Object.keys(zpend).length && !zst.err) { clearTimeout(zTimer); zTimer = setTimeout(zSync, 3000); }
  }
  function drawSync() {
    const b = $('zSync'); if (!b) return;
    if (!GURL) { b.innerHTML = ''; return; }
    const n = Object.keys(zpend).length;
    if (!gs) {
      if (b.querySelector('form')) return; // لا نمسح ما يكتبه المستخدم
      b.innerHTML = `<p><b>القوالب محفوظة في هذا المتصفح فقط.</b> اربط التطبيق بالإقرار لتُحفظ القوالب دائمًا وتظهر على كل أجهزتك، ولتتمكن من إرسال الفواتير إلى الإقرار.</p>
        <form id="zsForm" class="zsform"><input id="zsEmail" type="email" autocomplete="email" placeholder="البريد الإلكتروني" required><input id="zsToken" autocomplete="off" placeholder="رمز الدخول" required><button class="primary" id="zsGo">ربط بالإقرار</button></form>
        <small>البريد والرمز من قائمة «قارئ الفواتير» ← «إضافة مستخدم مسموح له» في الإقرار.</small>`;
      return;
    }
    b.innerHTML = `<p class="zson">☁️ <b>القوالب محفوظة في الإقرار</b> باسم ${esc(gs.email)}${zst.busy ? ' · جارٍ المزامنة…' : zst.err ? '' : zst.at ? ' · آخر مزامنة ' + esc(zst.at) : ''}${n && !zst.busy ? ` · ${n} تعديل بانتظار الإرسال` : ''}</p>
      ${zst.err ? `<p class="zserr" role="alert">تعذّرت المزامنة: ${esc(zst.err)}. القوالب باقية في هذا المتصفح وستُرسل لاحقًا.</p>` : ''}
      <div class="zbtns"><button type="button" id="zsNow"${zst.busy ? ' disabled' : ''}>🔄 مزامنة الآن</button><button type="button" id="zsOff">إلغاء الربط</button></div>`;
  }
  document.addEventListener('submit', async e => {
    if (e.target.id !== 'zsForm') return; e.preventDefault();
    const email = $('zsEmail').value.trim().toLowerCase(), token = $('zsToken').value.trim();
    if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email) || !token) { toast('اكتب البريد ورمز الدخول'); return; }
    $('zsGo').disabled = true; $('zsGo').textContent = 'جارٍ الربط…';
    try { await gsCall({ email, token }); gs = { url: GURL, email, token }; try { localStorage.setItem(GKEY, JSON.stringify(gs)); } catch { }
      $('zSync').innerHTML = ''; await zSync(); toast('تم الربط: القوالب محفوظة في الإقرار'); }
    catch (er) { toast((er && er.message) || 'تعذّر الربط'); const g = $('zsGo'); if (g) { g.disabled = false; g.textContent = 'ربط بالإقرار'; } }
  });
  document.addEventListener('click', e => {
    const b = e.target.closest('button'); if (!b) return;
    if (b.id === 'zsNow') zSync();
    if (b.id === 'zsOff') { const old = gs; gs = null; try { localStorage.setItem(GKEY, JSON.stringify({ url: GURL })); } catch { } drawSync();
      undoBar('أُلغي الربط مع الإقرار. القوالب باقية في هذا المتصفح', () => { gs = old; try { localStorage.setItem(GKEY, JSON.stringify(old)); } catch { } drawSync(); }, () => { }); }
  });
  if (gs) setTimeout(zSync, 1500);
  if (gs) setTimeout(supSync, 2500);

  // ---------- إرسال الفواتير إلى Google Sheet (ورقة «فواتير التطبيق») ----------
  const sentSig = x => JSON.stringify([x.invoice_number, x.invoice_date, x.total, x.tax_amount, x.vendor_name, x.vendor_tax_number, x.payment_method, x.project || '', noteOf(x), x.file || '', !!x.deletedAt, x.delReason || '']) + (x.verified ? '|v' : '');
  const sendState = x => !x.sentAt ? (x.sendErr ? 'err' : 'new') : x.sentSig === sentSig(x) ? 'ok' : 'dirty';
  function sentBadge(x) {
    if (!GURL && !x.sentAt) return '';
    const st = sendState(x);
    const when = x.sentAt ? ' · ' + sentWhen(x.sentAt) : '', row = x.sentRow ? ' · صف ' + x.sentRow : '';
    if (st === 'ok') return `<small class="sb ok">✓ صُدّرت إلى الإقرار${esc(when + row)}</small>`;
    if (st === 'dirty') return `<small class="sb dirty">⟳ صُدّرت إلى الإقرار${esc(when)} ثم عُدّلت: ستُحدَّث عند الإرسال</small>`;
    if (st === 'err') return `<small class="sb err">لم تُرسل: ${esc(String(x.sendErr).replace(/الإقرار|الإقرار/g, 'الإقرار'))}</small>`;
    return '';
  }
  function sentWhen(t) { try { return new Date(t).toLocaleString('ar-SA-u-nu-latn-ca-gregory', { dateStyle: 'medium', timeStyle: 'short' }); } catch { return ''; } }
  const toSend = () => list.filter(x => sendState(x) !== 'ok').concat(trash.filter(x => x.sentAt && sendState(x) === 'dirty'));
  let sending = false;
  function sendLabel() {
    const b = $('gsSend'); if (!b) return;
    b.hidden = !GURL || (!list.length && !trash.some(x => x.sentAt));
    const n = toSend().length;
    b.disabled = sending;
    b.textContent = sending ? 'جارٍ النقل إلى الإقرار…' : n ? `☁️ نقل إلى الإقرار (${n})` : '☁️ نقل إلى الإقرار';
  }
  $('gsSend').onclick = async () => {
    if (!gs) { toast('اربط التطبيق بالإقرار أولًا: اكتب بريدك ورمز الدخول'); showTab('z'); const f = $('zsEmail'); if (f) { f.scrollIntoView({ block: 'center', behavior: RM ? 'auto' : 'smooth' }); setTimeout(() => f.focus({ preventScroll: true }), 300); } return; }
    if (sending) return; const items = toSend();
    sending = true; sendLabel();
    const sig = new Map(items.map(x => [x.id, sentSig(x)]));
    try {
      const n2s = v => { const t = n2(v); return typeof t === 'number' ? t : (t || ''); };
      const invoices = items.map(x => ({ id: x.id, InvoiceNo: String(x.invoice_number || ''), Date: x.invoice_date || '', Total: n2s(x.total), Vat: n2s(x.tax_amount), Vendor: x.vendor_name || '',
        VATID: x.vendor_tax_number || '', Payment: x.payment_method || '', Project: x.project || '', Note: noteOf(x), File: x.file || '', Deleted: !!x.deletedAt, DeleteReason: x.delReason || '', Verified: !!x.verified }));
      const r = await fetch(GURL, { method: 'POST', body: JSON.stringify(items.length ? { email: gs.email, token: gs.token, action: 'appInvoices', invoices, copy: true } : { email: gs.email, token: gs.token, action: 'copyToInvoiceT' }), redirect: 'follow' });
      let j; try { j = await r.json(); } catch { throw new Error('رد غير متوقع من الإقرار'); }
      if (!j.ok) throw new Error(j.error || 'تعذّر الإرسال');
      if (!j.results) throw new Error('الإقرار لا يستقبل الفواتير بعد: حدّث كود Apps Script وانشر نسخة جديدة');
      let ok = 0, sk = 0; const now = new Date().toISOString(), all = list.concat(trash);
      j.results.forEach(o => { const x = all.find(y => y.id === o.id); if (!x) return;
        if (o.skipped) { x.sendErr = String(o.skipped).replace(/الإقرار/g, 'الإقرار'); sk++; return; }
        x.sentAt = now; x.sentRow = o.row; x.sentSig = sig.get(x.id); delete x.sendErr; ok++; });
      save(); saveTrash();
      const head = items.length ? (sk ? `أُرسلت ${ok} فاتورة إلى «فواتير التطبيق»، ولم تُرسل ${sk} لأنها مكررة في الإقرار.` : `أُرسلت ${ok} فاتورة إلى «فواتير التطبيق».`) : '';
      if (j.copy && !j.copy.error && !j.ver && items.some(x => x.verified)) j.copy.msg += '\n\n⚠ حالة التدقيق لم تُنقل: حدّث كود Apps Script في الإقرار وانشر نسخة جديدة (New version).';
      if (!j.copy) showRep(head + '\n\nلم يتم النقل إلى الإقرار: حدّث كود Apps Script في الإقرار وانشر نسخة جديدة (New version).', true);
      else if (j.copy.error) showRep(head + '\n\nتعذّر النقل إلى الإقرار: ' + j.copy.error, true);
      else showRep((head ? head + '\n\n' : '') + j.copy.msg, false);
    } catch (e) { console.warn(e); showRep('تعذّر الإرسال: ' + ((e && e.message) || 'تحقق من الاتصال'), true); }
    sending = false; draw();
  };

  // تقرير الإرسال والنقل (يبقى ظاهرًا حتى يُغلق)
  function showRep(t, bad) {
    const b = $('gsRep'); if (!b) { toast(t); return; }
    b.hidden = false; b.classList.toggle('bad', !!bad);
    b.innerHTML = `<div class="gsrep-h"><b>${bad ? 'نتيجة الإرسال' : 'تم النقل إلى الإقرار'}</b><button type="button" class="ghost" id="gsRepX" aria-label="إغلاق">✕ إغلاق</button></div><div class="gsrep-b">${esc(t)}</div>`;
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
  window.__invoiceApp = { addFiles, zones: () => zones, rr: rrS, tpl: { pickAnchor, readZone, regionWords, anchorRect, loadSrc, getRec }, redraw: () => draw(), extract, get busy() { return busy || queue.length > 0; }, get list() { return list; } };
})();
