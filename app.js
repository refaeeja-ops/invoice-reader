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
    const doc = await pdfjsLib.getDocument({ data: await file.arrayBuffer() }).promise;
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
      if (good(r) || (k >= 1 && P.itemsOk(r))) break;
      if (eastern(runs[0].t) && k >= 1) break; // للأرقام الهندية ننتقل للنموذج المتخصص
    }
    if (runs.some(x => eastern(x.t)) && !runs.some(x => good(x.r))) {
      for (let k = 0; k < EAST_PASSES.length; k++) {
        say('وُجدت أرقام هندية، جارٍ قراءتها بنموذج متخصص…');
        ocrBase = 0.1 + 0.8 * (PASSES.length + k) / (PASSES.length + EAST_PASSES.length);
        const r = await run('Arabic+eng', EAST_PASSES[k]);
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
    d.file = pathOf(file) || 'صورة من الكاميرا';
    return { d, pages: list.map(p => toUrl(p.canvas)) };
  }
  function toUrl(c) { // صورة للعرض فقط، بحجم أصغر لتوفير الذاكرة
    const r = Math.min(1, 1600 / Math.max(c.width, c.height));
    if (r === 1) return c.toDataURL('image/jpeg', 0.82);
    const s = document.createElement('canvas'); s.width = Math.round(c.width * r); s.height = Math.round(c.height * r);
    s.getContext('2d').drawImage(c, 0, 0, s.width, s.height); return s.toDataURL('image/jpeg', 0.82);
  }

  // ---------- طابور الدفعة ----------
  const isInv = f => /^image\/|pdf/.test(f.type) || /\.(pdf|jpe?g|png|webp|heic)$/i.test(f.name || '');
  const pathOf = f => f.webkitRelativePath || f._path || f.name || '';
  function addFiles(files, fromFolder) {
    const allF = [...(files || [])];
    files = allF.filter(f => isInv(f) && !/^\./.test(f.name || ''));
    if (fromFolder) files.sort((a, b) => pathOf(a).localeCompare(pathOf(b), 'ar', { numeric: true }));
    const skipped = allF.length - files.length;
    if (!files.length) { if (allF.length || fromFolder) toast(fromFolder ? 'لا توجد صور أو ملفات PDF في هذا المجلد' : 'الملف ليس صورة أو PDF'); return; }
    if (fromFolder) setTimeout(() => toast(`وُجدت ${files.length} فاتورة في المجلد` + (skipped ? ` (تُجوهل ${skipped} ملف ليس صورة أو PDF)` : '')), 50);
    files.forEach((f, i) => { const id = Date.now() + '-' + i + '-' + Math.random().toString(36).slice(2, 6); queue.push({ id, file: f }); live[id] = { status: 'queued', file: pathOf(f) || 'صورة من الكاميرا' }; });
    drawList(); if (files.length > 1 && !fromFolder) toast(`أُضيفت ${files.length} فواتير إلى الدفعة`);
    pump();
  }
  let setProg = () => { };
  async function pump() {
    if (busy) return; busy = true;
    let done = 0; const batch = [];
    const total = () => done + queue.length;
    while (queue.length) {
      const job = queue.shift(); live[job.id].status = 'working'; drawList();
      const n = done + 1, N = total();
      const say = m => status(N > 1 ? `فاتورة ${n} من ${N}: ${m}` : m);
      setProg = p => progress(((n - 1) + Math.min(1, p)) / N);
      setProg(0.02); say('جارٍ فتح الملف…');
      try {
        const { d, pages: urls } = await processFile(job.file, say);
        d._id = job.id; d._at = new Date().toISOString(); d._warnings = P.checks(d, d._qr);
        imgs[job.id] = urls; delete live[job.id];
        store(d); batch.push(job.id);
        if (!cur || N === 1 || n === 1) open(job.id, n === 1);
      } catch (e) {
        console.error(e); live[job.id].status = 'error';
        live[job.id].err = e && e.message === 'img' ? 'تعذّر فتح الصورة' : ((e && e.message) || 'خطأ');
      }
      done++; drawList();
    }
    busy = false; progress(null); setProg = () => { };
    const dm = dupMap(), nd = batch.filter(id => dm.has(id)).length;
    status((done > 1 ? `انتهت الدفعة: ${done} فواتير. افتح أي فاتورة من القائمة لمراجعتها.` : (cur && cur._qr ? 'تمت القراءة. وُجد رمز QR الضريبي.' : 'تمت القراءة. راجع الحقول وصحّح ما يلزم.'))
      + (nd ? ` تنبيه: ${nd === 1 ? 'فاتورة مكررة' : nd + ' فواتير مكررة'} (نفس المورّد والرقم والتاريخ)، انظر القائمة.` : ''), !!nd);
    if (cur) refresh();
  }

  // ---------- كشف الفواتير المكررة: نفس المورّد (الرقم الضريبي أو الاسم) ونفس رقم الفاتورة ونفس التاريخ ----------
  const nk = s => P.normDigits(s).toLowerCase().replace(/[أإآ]/g, 'ا').replace(/ة/g, 'ه').replace(/ى/g, 'ي').replace(/[^0-9a-z\u0621-\u064A]/g, '');
  const vName = s => nk(String(s || '').replace(/^\s*(شركة|مؤسسة|مؤسسه|شركه)\s+/, '').replace(/\b(co|company|ltd|llc|est)\b\.?/gi, ''));
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
    $('pinfo').textContent = n > 1 ? `صفحة ${pageIdx + 1} من ${n}` : (n ? '' : 'الصورة غير محفوظة بعد إغلاق الصفحة');
    $('prevP').hidden = $('nextP').hidden = n < 2;
    $('prevP').disabled = pageIdx === 0; $('nextP').disabled = pageIdx >= n - 1;
  }
  const numStr = v => (v == null || v === '') ? '' : String(v);
  const tag = k => cur.src && cur.src[k] ? `<span class="tag ${cur.src[k]}">${SRC[cur.src[k]]}</span>` : '';

  function open(id, scroll) {
    const d = all().find(x => x._id === id); if (!d) return;
    cur = d; qr = d._qr || null; rawText = d._raw || ''; textOnly = P.parseInvoice(rawText, null);
    cur.src = cur.src || {}; cur.items = cur.items || [];
    pages = imgs[id] || []; pageIdx = 0; showPage(); render(); drawList();
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
      setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000); toast('جارٍ التنزيل');
    } catch { toast('التنزيل غير متاح هنا، استخدم النسخ'); }
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
  const KEY = 'invoice-reader-history-v1', MAX = 200;
  let cache = null;
  const all = () => { if (cache) return cache; try { cache = JSON.parse(localStorage.getItem(KEY)) || []; } catch { cache = []; } return cache; };
  function persist() { try { localStorage.setItem(KEY, JSON.stringify(all().slice(0, MAX))); } catch { toast('مساحة المتصفح ممتلئة: صدّر الفواتير ثم احذف القديمة'); } }
  function store(d, quiet) {
    const h = all(), i = h.findIndex(x => x._id === d._id);
    if (i >= 0) h[i] = d; else h.unshift(d);
    clearTimeout(store.t); store.t = setTimeout(persist, quiet ? 400 : 0);
    if (quiet) { clearTimeout(store.h); store.h = setTimeout(drawList, 500); } else drawList();
  }
  // ---------- البيانات المرجعية (الموردين والمشاريع والمنتجات) ----------
  const R = window.InvoiceRef, RKEY = 'invoice-reader-ref-v1';
  let ref = null; try { ref = JSON.parse(localStorage.getItem(RKEY)) || null; } catch { ref = null; }
  let supEdit = null; // فاتورة يجري تغيير مورّدها يدويًا
  function link(x) {
    if (!ref) return;
    if (!x._supManual) { const m = R.matchSupplier(ref, x); x._supId = m ? m.id : null; x._supBy = m ? m.by : null; }
    const s = R.supplier(ref, x._supId); x._supName = s ? s.name : '';
    const prods = x._supId != null ? R.productsFor(ref, x._supId) : [];
    if (x.product_id != null && prods.length && !prods.some(p => p.id === x.product_id) && !x._prodAny) { x.product_id = null; }
    if (x.product_id == null && prods.length === 1) x.product_id = prods[0].id;
    const pr = R.project(ref, x.project_id), pd = R.product(ref, x.product_id);
    x.project_name = pr ? pr.name : ''; x.product_name = pd ? pd.name : '';
    const ex = R.existsInvoice(ref, x._supId, x.invoice_number); x._exists = ex || null;
  }
  function refNotes(x) {
    if (!ref) return [];
    const n = [];
    if (x._exists) n.push(['wn', `مسجلة سابقًا في InvoiceT${x._exists !== true ? ' (InvID ' + x._exists + ')' : ''}`]);
    if (x._supId == null) n.push(['wn', 'المورّد غير موجود في SuppliersT: اختره من القائمة']);
    else if (x._supBy === 'name') n.push(['wn', 'طوبق المورّد بالاسم فقط، تأكد منه']);
    if (x.project_id == null) n.push(['todo', 'اختر المشروع']);
    if (x.product_id == null) n.push(['todo', 'اختر المنتج']);
    return n;
  }
  const opt = (v, label, sel) => `<option value="${esc(v)}"${sel ? ' selected' : ''}>${esc(label)}</option>`;
  function supCell(x) {
    if (!ref) return '';
    const id = esc(x._id);
    if (x._supId != null && supEdit !== x._id) return `<small class="${x._supBy === 'name' ? 'supname' : 'supok'}">SupID ${x._supId}: ${esc(x._supName)}${x._supBy === 'name' ? ' (بالاسم)' : x._supBy === 'manual' ? ' (يدوي)' : ''}</small><button class="linkb" data-supedit="${id}">تغيير</button>`;
    const list = ref.suppliers.slice().sort((a, b) => a.name.localeCompare(b.name, 'ar'));
    return `<select data-sup="${id}" aria-label="اختر المورّد المسجل">${opt('', '— اختر المورّد المسجل —', x._supId == null)}${x._supManual ? opt('auto', '↺ مطابقة تلقائية', false) : ''}${list.map(s => opt(s.id, s.id + ' - ' + s.name, s.id === x._supId)).join('')}</select>`;
  }
  function linkCell(x) {
    if (!ref) return '<span class="todo">حمّل ملف البيانات</span>';
    const id = esc(x._id);
    const projs = R.activeProjects(ref).slice(); const cp = R.project(ref, x.project_id); if (cp && !cp.active) projs.unshift(cp);
    let prods = x._supId != null ? R.productsFor(ref, x._supId) : [], all = false;
    if (!prods.length || x._prodAny) { prods = ref.products.slice().sort((a, b) => a.name.localeCompare(b.name, 'ar')); all = true; }
    const cpd = R.product(ref, x.product_id); if (cpd && !prods.some(p => p.id === cpd.id)) prods.unshift(cpd);
    return `<label>المشروع<select data-proj="${id}">${opt('', '— اختر المشروع —', x.project_id == null)}${projs.map(p => opt(p.id, p.id + ' - ' + p.name, p.id === x.project_id)).join('')}</select></label>
      <label>المنتج${all && x._supId != null && !x._prodAny ? ' (لا منتجات مرتبطة بالمورّد)' : ''}<select data-prod="${id}">${opt('', '— اختر المنتج —', x.product_id == null)}${prods.map(p => opt(p.id, p.id + ' - ' + p.name, p.id === x.product_id)).join('')}${!all ? opt('any', '… كل المنتجات', false) : ''}</select></label>`;
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
      try { localStorage.setItem(RKEY, JSON.stringify(r)); } catch { toast('الملف كبير على ذاكرة المتصفح: سيُستخدم حتى إغلاق الصفحة فقط'); }
      all().forEach(x => { if (!x._supManual) delete x._supId; }); persist();
      refInfo(); drawList(); toast('تم تحميل ملف البيانات ومطابقة الفواتير');
    } catch (err) { console.error(err); toast((err && err.message) || 'تعذّرت قراءة الملف'); }
  };
  $('refClear').onclick = () => { ref = null; localStorage.removeItem(RKEY); refInfo(); drawList(); toast('أُزيل ملف البيانات من هذا المتصفح'); };
  $('hist').addEventListener('change', e => {
    const t = e.target, d = t.dataset, id = d.sup || d.proj || d.prod; if (!id) return;
    const x = all().find(v => v._id === id); if (!x) return;
    if (d.sup) { if (t.value === 'auto') { x._supManual = false; } else { x._supManual = true; x._supId = t.value ? +t.value : null; x._supBy = 'manual'; } x._prodAny = false; supEdit = null; }
    if (d.proj) x.project_id = t.value ? +t.value : null;
    if (d.prod) { if (t.value === 'any') { x._prodAny = true; } else x.product_id = t.value ? +t.value : null; }
    persist(); drawList();
  });

  let filter = 'all';
  const shownList = () => all().filter(x => filter === 'all' || (filter === 'approved') === !!x._approved);
  const money = x => x.total != null && x.total !== '' ? P.fmt(x.total) + ' ' + (x.currency || '') : '—';
  const sumBy = list => Object.entries(list.reduce((a, x) => { const c = x.currency || '—'; a[c] = (a[c] || 0) + (+x.total || 0); return a; }, {})).map(([c, v]) => P.fmt(v) + ' ' + c).join(' · ');
  function drawList() {
    const h = all(), pend = Object.entries(live);
    $('count').textContent = h.length ? `(${h.length})` : '';
    $('exportBar').hidden = !h.length;
    const dm = dupMap();
    h.forEach(x => { const d = dm.get(x._id); if (d) x._dup = dupMsg(d); else delete x._dup; link(x); });
    $('xVat').textContent = `VatReport (${filter === 'all' ? 'الكل' : filter === 'approved' ? 'المعتمدة' : 'بانتظار الاعتماد'})`;
    const nA = h.filter(x => x._approved).length, nP = h.length - nA;
    $('filters').hidden = !h.length;
    $('filters').innerHTML = [['all', 'الكل', h.length], ['pending', 'بانتظار الاعتماد', nP], ['approved', 'معتمدة', nA]]
      .map(([k, l, n]) => `<button data-filter="${k}" aria-pressed="${filter === k}">${l} (${n})</button>`).join('');
    const STAT = { queued: 'بالانتظار', working: 'جارٍ القراءة…', error: 'تعذّرت القراءة' };
    const rowsP = pend.map(([id, v]) => `<tr class="pend ${v.status}"><td>${esc(v.file)}</td><td colspan="6">${STAT[v.status]}${v.err ? ': ' + esc(v.err) : ''}</td><td class="act">${v.status === 'error' ? `<button data-dismiss="${esc(id)}">إخفاء</button>` : ''}</td></tr>`).join('');
    const shown = shownList();
    const rows = shown.map(x => {
      const id = esc(x._id), dup = dm.get(x._id), w = x._warnings || [];
      const notes = [...(dup ? ['<span class="dupb">مكررة</span>'] : []), ...refNotes(x).map(([c, m]) => `<span class="${c}">• ${esc(m)}</span>`), ...w.map(m => `<span class="wn">• ${esc(m)}</span>`)].join('');
      return `<tr class="${cur && cur._id === x._id ? 'sel' : ''}${dup ? ' dup' : ''}${x._approved ? ' appr' : ''}"${dup ? ` title="${esc(x._dup)}"` : ''}>
        <td data-l="المورّد"><b>${esc(x.vendor_name || 'بدون اسم')}</b><small>${esc(x.file || '')}</small>${supCell(x)}</td>
        <td data-l="الرقم الضريبي" class="mono">${esc(x.vendor_tax_number || '—')}</td>
        <td data-l="رقم الفاتورة" class="nw">${esc(x.invoice_number || '—')}</td>
        <td data-l="التاريخ" class="nw">${esc(x.invoice_date || '—')}</td>
        <td data-l="الإجمالي" class="n"><b>${esc(money(x))}</b><small>الضريبة: ${x.tax_amount != null && x.tax_amount !== '' ? esc(P.fmt(x.tax_amount)) : '—'}</small></td>
        <td data-l="المشروع والمنتج" class="lnk">${linkCell(x)}</td>
        <td data-l="الملاحظات" class="notes">${notes ? `<div class="auto">${notes}</div>` : '<div class="auto"><span class="ok">لا تنبيهات</span></div>'}<input class="unote" data-note="${id}" value="${esc(x.note || '')}" placeholder="أضف ملاحظتك…" aria-label="ملاحظتك على الفاتورة"></td>
        <td class="act"><div class="st">${x._approved ? `<span class="okb">معتمدة</span>${x._approvedAt ? `<small>${esc(new Date(x._approvedAt).toLocaleDateString('ar-SA-u-ca-gregory-nu-latn'))}</small>` : ''}` : '<span class="pendb">بانتظار الاعتماد</span>'}</div>${x._approved ? `<button data-unapprove="${id}">إلغاء الاعتماد</button>` : `<button class="primary" data-approve="${id}">اعتماد</button>`}<button data-open="${id}">مراجعة</button><button data-rm="${id}" class="del">حذف</button></td></tr>`;
    }).join('');
    $('hist').tBodies[0].innerHTML = rowsP + rows || `<tr><td colspan="8" class="empty">${h.length ? 'لا توجد فواتير في هذا التصنيف.' : 'لا توجد فواتير بعد. يمكنك اختيار عدة ملفات أو مجلد كامل.'}</td></tr>`;
    const nd = dm.size;
    $('dupNote').hidden = !nd;
    $('dupNote').textContent = nd ? `تنبيه: ${nd} فواتير في القائمة مكررة (نفس المورّد ورقم الفاتورة والتاريخ). هي معلّمة بكلمة «مكررة»؛ راجعها واحذف الزائد منها.` : '';
    $('sumAll').textContent = h.length ? `مجموع الكل: ${sumBy(h)}` + (nA ? ` · المعتمد: ${sumBy(h.filter(x => x._approved))}` : '') : '';
  }
  $('filters').addEventListener('click', e => { const f = e.target.dataset.filter; if (f) { filter = f; drawList(); } });
  $('hist').addEventListener('input', e => {
    const id = e.target.dataset.note; if (!id) return;
    const x = all().find(v => v._id === id); if (!x) return;
    x.note = e.target.value; clearTimeout(store.t); store.t = setTimeout(persist, 400); // بدون إعادة رسم حتى لا يضيع المؤشر
  });
  $('hist').addEventListener('click', e => {
    if (e.target.dataset.supedit) { supEdit = e.target.dataset.supedit; drawList(); const s = document.querySelector(`select[data-sup="${CSS.escape(supEdit)}"]`); if (s) s.focus(); return; }
    const a = e.target.dataset.approve, u = e.target.dataset.unapprove;
    if (!a && !u) return;
    const x = all().find(v => v._id === (a || u)); if (!x) return;
    x._approved = !!a; if (a) x._approvedAt = new Date().toISOString(); else delete x._approvedAt;
    persist(); drawList(); toast(a ? 'اعتُمدت الفاتورة' : 'أُلغي الاعتماد');
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
  $('xVat').onclick = () => { const l = shownList(); if (!l.length) return toast('لا توجد فواتير في هذا التصنيف'); if (!ref) toast('تنبيه: لم يُحمَّل ملف البيانات، فستكون SupID و ProjectID فارغة'); download(`VatReport-${stamp()}.xlsx`, X.vatXlsx(l, ref && ref.vatCols)); };
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
  $('zoom').onclick = () => { const z = $('stage').classList.toggle('zoom'); $('zoom').setAttribute('aria-pressed', z); $('zoom').textContent = z ? 'تصغير' : 'تكبير'; };
  $('prevP').onclick = () => { if (pageIdx > 0) { pageIdx--; showPage(); } };
  $('nextP').onclick = () => { if (pageIdx < pages.length - 1) { pageIdx++; showPage(); } };
  $('again').onclick = () => { $('review').classList.remove('on'); cur = null; drawList(); status(''); window.scrollTo(0, 0); };
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
  refInfo(); drawList();
  window.__invoiceApp = { addFiles, get cur() { return cur; }, get qr() { return qr; }, get raw() { return rawText; }, get busy() { return busy || queue.length > 0; }, all, dupMap };
})();
