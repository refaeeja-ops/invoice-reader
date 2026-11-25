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
    return { d, pages: list.map(p => toUrl(p.canvas)) };
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
  const save = () => { try { localStorage.setItem(KEY, JSON.stringify(list)); } catch { toast('مساحة المتصفح ممتلئة: نزّل ملف Excel ثم احذف الفواتير القديمة'); } };
  const DB = (name, store) => {
    let p = null;
    const db = () => p || (p = new Promise((ok, no) => { try { const r = indexedDB.open(name, 1); r.onupgradeneeded = () => r.result.createObjectStore(store); r.onsuccess = () => ok(r.result); r.onerror = () => no(r.error); } catch (e) { no(e); } }));
    const tx = async (mode, fn) => { const d = await db(); return new Promise((ok, no) => { const t = d.transaction(store, mode), r = fn(t.objectStore(store)); t.oncomplete = () => ok(r && r.result); t.onerror = () => no(t.error); }); };
    return { put: (k, v) => tx('readwrite', s => s.put(v, k)).catch(() => { }), get: k => tx('readonly', s => s.get(k)).catch(() => null), del: k => tx('readwrite', s => s.delete(k)).catch(() => { }) };
  };
  const IMG = DB('invoice-simple-files', 'f'), OLD = DB('invoice-reader-img', 'img');

  // أول تشغيل: نأخذ الفواتير الموجودة من النسخة الكاملة حتى لا تضيع
  if (localStorage.getItem(KEY) === null) {
    try {
      const old = JSON.parse(localStorage.getItem('invoice-reader-history-v1')) || [];
      list = old.map(x => ({ id: x._id, at: x._at || '', file: x.file || '', invoice_number: x.invoice_number || '', invoice_date: x.invoice_date || '', total: x.total ?? '', tax_amount: x.tax_amount ?? '', vendor_name: x.vendor_name || '', vendor_tax_number: x.vendor_tax_number || '', payment_method: x.payment_method || '', note: x.note || '', old: true }));
    } catch { list = []; }
    save();
  }

  // ---------- الخانات ----------
  const F = [
    ['invoice_number', 'رقم الفاتورة', 'text'], ['invoice_date', 'تاريخ الفاتورة', 'date'],
    ['total', 'المبلغ الإجمالي', 'num'], ['tax_amount', 'مبلغ الضريبة', 'num'],
    ['vendor_name', 'اسم المورد أو التاجر', 'text'], ['vendor_tax_number', 'الرقم الضريبي للمورد', 'vat'],
    ['payment_method', 'طريقة الدفع', 'pay']];
  const PAYS = ['نقدًا', 'مدى', 'بطاقة ائتمان', 'تحويل بنكي', 'آجل', 'شيك', 'Apple Pay'];
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
  const fieldHtml = (x, [k, label, type]) => {
    const v = x[k] ?? '', miss = blank(v) && k !== 'payment_method';
    const attrs = `data-id="${esc(x.id)}" data-k="${k}" id="f-${esc(x.id)}-${k}"`;
    let input;
    if (type === 'date') input = `<input type="date" ${attrs} value="${esc(/^\d{4}-\d{2}-\d{2}$/.test(v) ? v : '')}">`;
    else if (type === 'pay') input = `<input ${attrs} list="pays" value="${esc(v)}" placeholder="مثال: مدى">`;
    else input = `<input ${attrs} value="${esc(type === 'num' && !blank(v) && isFinite(+v) ? (+v).toFixed(2) : v)}"${type !== 'text' ? ' inputmode="decimal" class="num"' : ''} placeholder="${miss ? 'لم يُقرأ — أدخله' : ''}">`;
    return `<label class="f f-${k}${miss ? ' miss' : ''}"><span>${label}</span>${input}</label>`;
  };
  function card(x) {
    const w = issues(x);
    return `<article class="inv" id="c-${esc(x.id)}">
      <header><div><b>${esc(x.vendor_name || 'مورد غير معروف')}</b><small class="fp">📄 ${esc(x.file || '')}</small></div><span class="amt">${esc(money(x.total))}</span></header>
      <div class="grid">${F.map(f => fieldHtml(x, f)).join('')}
        <label class="f wide"><span>ملاحظات</span><textarea data-id="${esc(x.id)}" data-k="note" rows="2" placeholder="اكتب أي إشكالية في الفاتورة">${esc(x.note || '')}</textarea></label></div>
      ${w.length ? `<ul class="warn">${w.map(t => `<li>${esc(t)}</li>`).join('')}</ul>` : '<p class="okline">✓ البيانات الأساسية مكتملة</p>'}
      <footer><button data-view="${esc(x.id)}">فتح الملف للمراجعة</button><button data-reread="${esc(x.id)}"${live[x.id] ? ' disabled' : ''}>${live[x.id] ? (live[x.id].status === 'working' ? 'جارٍ إعادة القراءة…' : 'بانتظار إعادة القراءة…') : '🔄 إعادة قراءة الفاتورة'}</button><button class="del" data-del="${esc(x.id)}">حذف الفاتورة</button></footer>
    </article>`;
  }
  function draw() {
    const pend = Object.values(live).filter(v => !v.reread);
    $('count').textContent = list.length ? `(${list.length})` : '';
    $('xl').disabled = !list.length; $('xl').textContent = `تنزيل ملف Excel${list.length ? ' (' + list.length + ')' : ''}`;
    const nIss = list.filter(x => issues(x).length).length;
    $('sum').textContent = list.length ? (nIss ? `${nIss} فاتورة فيها ملاحظات أو بيانات ناقصة` : 'كل الفواتير مكتملة') : '';
    $('list').innerHTML = pend.map(v => `<article class="inv pend"><b>📄 ${esc(v.file)}</b> <span>${v.status === 'working' ? 'جارٍ القراءة…' : 'بالانتظار'}</span></article>`).join('')
      + (list.map(card).join('') || (pend.length ? '' : '<p class="empty">لا توجد فواتير بعد. ارفع فاتورة أو صوّرها للبدء.</p>'));
  }
  // تحديث ملاحظات وحالة بطاقة واحدة بدون إعادة رسم (حتى لا يضيع مكان الكتابة)
  function refreshCard(x) {
    const c = $('c-' + x.id); if (!c) return;
    const w = issues(x), box = c.querySelector('.warn, .okline');
    const html = w.length ? `<ul class="warn">${w.map(t => `<li>${esc(t)}</li>`).join('')}</ul>` : '<p class="okline">✓ البيانات الأساسية مكتملة</p>';
    box.outerHTML = html;
    c.querySelector('header b').textContent = x.vendor_name || 'مورد غير معروف';
    c.querySelector('.amt').textContent = money(x.total);
    F.forEach(([k]) => { const el = $(`f-${x.id}-${k}`); if (el) el.closest('.f').classList.toggle('miss', blank(x[k]) && k !== 'payment_method'); });
    const nIss = list.filter(y => issues(y).length).length;
    $('sum').textContent = nIss ? `${nIss} فاتورة فيها ملاحظات أو بيانات ناقصة` : 'كل الفواتير مكتملة';
  }
  document.addEventListener('input', e => {
    const el = e.target, id = el.dataset.id, k = el.dataset.k; if (!id || !k) return;
    const x = list.find(y => y.id === id); if (!x) return;
    let v = el.value.trim();
    if (k === 'total' || k === 'tax_amount') v = n2(v);
    if (k === 'vendor_tax_number') v = P.normDigits(v).replace(/\s/g, '');
    x[k] = v; (x.edited = x.edited || {})[k] = 1; save(); refreshCard(x);
  });
  document.addEventListener('click', async e => {
    const b = e.target.closest('button'); if (!b) return;
    if (b.dataset.del) {
      const x = list.find(y => y.id === b.dataset.del); if (!x) return;
      if (!confirm(`حذف فاتورة ${x.vendor_name || ''} رقم ${x.invoice_number || '—'}؟`)) return;
      const i = list.indexOf(x); list.splice(i, 1); save(); draw();
      undoBar(`حُذفت الفاتورة`, () => { list.splice(i, 0, x); save(); draw(); }, () => IMG.del(x.id));
    }
    if (b.dataset.view) view(b.dataset.view);
    if (b.dataset.reread) reread(b.dataset.reread);
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
  function addFiles(files, fromCam) {
    const seen = new Set(list.map(x => x.sig).filter(Boolean).concat(queue.map(j => sig(j.file))));
    let skipped = 0;
    [...(files || [])].filter(f => /^image\//.test(f.type) || /\.(pdf|jpe?g|png|webp|heic|bmp|gif)$/i.test(f.name || '') || f.type === 'application/pdf').forEach((f, i) => {
      const s = sig(f); if (!fromCam && seen.has(s)) { skipped++; return; } seen.add(s);
      const id = Date.now().toString(36) + '-' + i + '-' + Math.random().toString(36).slice(2, 6);
      queue.push({ id, file: f }); live[id] = { status: 'queued', file: pathOf(f) || 'صورة من الكاميرا' };
    });
    if (skipped) toast(`${skipped} ملف لم يُضف لأن ملفًا بنفس الاسم موجود في القائمة`);
    draw(); pump();
  }
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
          const { d, pages } = await processFile(job.file, say);
          let changed = 0;
          F.forEach(([k]) => { const nv = k === 'total' || k === 'tax_amount' ? (d[k] ?? '') : (d[k] || ''); if (x && !(x.edited && x.edited[k]) && !blank(nv) && nv !== x[k]) { x[k] = nv; changed++; } });
          if (x) { save(); IMG.get(job.id).then(r0 => IMG.put(job.id, { pages, blob: (r0 && r0.blob) || job.file })); }
          toast(changed ? `أُعيدت القراءة: تحدّثت ${changed} من الخانات` : 'أُعيدت القراءة ولم تتغير البيانات');
        } catch (e) { console.error(e); toast('تعذّرت إعادة القراءة'); }
        deep = false; delete live[job.id]; done++; draw(); continue;
      }
      try {
        const { d, pages } = await processFile(job.file, say);
        const x = { id: job.id, at: new Date().toISOString(), file: d.file, file_name: d.file_name, sig: sig(job.file),
          invoice_number: d.invoice_number || '', invoice_date: d.invoice_date || '', total: d.total ?? '', tax_amount: d.tax_amount ?? '',
          vendor_name: d.vendor_name || '', vendor_tax_number: d.vendor_tax_number || '', payment_method: d.payment_method || '', note: '' };
        list.unshift(x); save();
        IMG.put(job.id, { pages, blob: job.file.size < 25e6 ? job.file : null });
      } catch (e) {
        console.error(e);
        const x = { id: job.id, at: new Date().toISOString(), file: live[job.id].file, sig: sig(job.file), invoice_number: '', invoice_date: '', total: '', tax_amount: '', vendor_name: '', vendor_tax_number: '', payment_method: '', note: 'تعذّرت قراءة الملف تلقائيًا: أدخل البيانات يدويًا' };
        list.unshift(x); save(); IMG.put(job.id, { pages: [], blob: job.file });
      }
      delete live[job.id]; done++; draw();
    }
    busy = false; progress(null);
    status(done ? (done > 1 ? `انتهت القراءة: ${done} فواتير. راجع الخانات الحمراء وأكملها.` : 'تمت القراءة. راجع الخانات وأكمل الناقص.') : '');
  }
  $('pick').addEventListener('change', e => { addFiles(e.target.files); e.target.value = ''; });
  $('dir').addEventListener('change', e => { addFiles(e.target.files); e.target.value = ''; });
  $('cam').addEventListener('change', e => { addFiles(e.target.files, true); e.target.value = ''; });
  const drop = $('drop');
  ['dragenter', 'dragover'].forEach(t => drop.addEventListener(t, e => { e.preventDefault(); drop.classList.add('over'); }));
  ['dragleave', 'drop'].forEach(t => drop.addEventListener(t, e => { e.preventDefault(); drop.classList.remove('over'); }));
  drop.addEventListener('drop', e => addFiles(e.dataTransfer.files));

  // ---------- Excel ----------
  const XL = [['invoice_number', 'رقم الفاتورة', 18], ['invoice_date', 'تاريخ الفاتورة', 14], ['total', 'المبلغ الإجمالي', 16], ['tax_amount', 'مبلغ الضريبة', 14],
    ['vendor_name', 'اسم المورد', 32], ['vendor_tax_number', 'الرقم الضريبي', 20], ['payment_method', 'طريقة الدفع', 14], ['note', 'ملاحظات', 50], ['file', 'الملف', 30]];
  $('xl').onclick = () => {
    const rows = [XL.map(c => c[1]), ...list.map(x => XL.map(([k]) => {
      if (k === 'note') return noteOf(x);
      if (k === 'total' || k === 'tax_amount') { const v = n2(x[k]); return typeof v === 'number' ? v : (v || ''); }
      return x[k] || '';
    }))];
    const blob = X.xlsx([{ name: 'الفواتير', rows, widths: XL.map(c => c[2]) }]);
    const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = `الفواتير-${new Date().toISOString().slice(0, 10)}.xlsx`;
    document.body.append(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
    toast('جارٍ تنزيل ملف Excel');
  };

  // ---------- شريط التراجع ----------
  function undoBar(msg, undo, commit) {
    const u = $('undo'); clearTimeout(undoBar.t); if (undoBar.c) undoBar.c();
    u.querySelector('span').textContent = msg; u.hidden = false; undoBar.c = commit;
    u.querySelector('button').onclick = () => { clearTimeout(undoBar.t); undoBar.c = null; u.hidden = true; undo(); };
    undoBar.t = setTimeout(() => { u.hidden = true; if (undoBar.c) { undoBar.c(); undoBar.c = null; } }, 7000);
  }

  // إزالة أي Service Worker قديم حتى تظهر آخر نسخة دائمًا
  if ('serviceWorker' in navigator) navigator.serviceWorker.getRegistrations().then(rs => rs.forEach(r => r.unregister())).catch(() => { });
  draw();
  window.__invoiceApp = { addFiles, get busy() { return busy || queue.length > 0; }, get list() { return list; } };
})();
