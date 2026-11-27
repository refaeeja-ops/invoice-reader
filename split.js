/* تقسيم ملف PDF فيه عدة فواتير: فاتورة واحدة، أو كل صفحة فاتورة، أو يحدد المستخدم بداية كل فاتورة.
   كل فاتورة تصبح ملف PDF مستقلًا (مكتبة pdf-lib المجانية) يُقرأ كأنه ملف مرفوع لوحده، ويمكن تنزيلها كملفات منفصلة.
   كل المعالجة داخل المتصفح ومجانية. */
(function () {
  'use strict';
  const MAXP = 60, PER = 5; // أقصى صفحات تُعرض للتقسيم، وأقصى صفحات تُقرأ لكل فاتورة
  const abs = p => new URL(p, location.href).href;
  const esc = s => String(s ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const isPdf = f => !!f && (f.type === 'application/pdf' || /\.pdf$/i.test(f.name || ''));

  // ---------- pdf-lib تُحمّل عند الحاجة فقط ----------
  let libP = null;
  const lib = () => window.PDFLib ? Promise.resolve(window.PDFLib) : (libP = libP || new Promise((ok, no) => {
    const s = document.createElement('script'); s.src = abs('vendor/pdf-lib.min.js');
    s.onload = () => window.PDFLib ? ok(window.PDFLib) : no(new Error('pdf-lib')); s.onerror = () => { libP = null; no(new Error('pdf-lib')); };
    document.head.appendChild(s);
  }));

  // ---------- رمز QR الضريبي (لاقتراح بدايات الفواتير وعرض اسم المورد) ----------
  function scanQR(canvas) {
    if (!window.jsQR) return null;
    const W = canvas.width, H = canvas.height, P = window.InvoiceParse;
    const regs = [[0, 0, W, H, 1000], [0, H / 2, W, H / 2, 1300], [0, 0, W, H / 2, 1300], [0, 0, W, H, 1600]];
    let any = null;
    for (const [sx, sy, sw, sh, t] of regs) {
      const r = Math.min(1.5, t / Math.max(sw, sh));
      const c = document.createElement('canvas'); c.width = Math.max(1, Math.round(sw * r)); c.height = Math.max(1, Math.round(sh * r));
      const x = c.getContext('2d', { willReadFrequently: true }); x.drawImage(canvas, sx, sy, sw, sh, 0, 0, c.width, c.height);
      const res = jsQR(x.getImageData(0, 0, c.width, c.height).data, c.width, c.height, { inversionAttempts: 'attemptBoth' });
      if (res && res.data) { const z = P && P.parseZatcaQR(res.data); if (z) return z; any = any || {}; }
    }
    return any;
  }

  // ---------- التقسيم ----------
  function groupsOf(S) { // الصفحات المتجاهلة لا تدخل في أي فاتورة
    if (!S.mode) return [];
    const isStart = i => i === 0 || S.mode === 'each' || (S.mode === 'manual' && S.pages[i].mark);
    const isSkip = i => S.mode === 'manual' && S.pages[i].skip;
    const out = []; let cur = null;
    S.pages.forEach((p, i) => {
      if (isStart(i) || !cur) { cur = { pages: [], qr: null, color: out.length }; out.push(cur); }
      if (!isSkip(i)) { cur.pages.push(i); if (p.qr && p.qr.seller && !cur.qr) cur.qr = p.qr; }
    });
    return out.filter(g => g.pages.length);
  }
  const rng = a => { // [0,1,2,4] → «1–3، 5»
    const r = []; let s = a[0], e = a[0];
    for (let i = 1; i <= a.length; i++) { if (a[i] === e + 1) { e = a[i]; continue; } r.push(s === e ? `${s + 1}` : `${s + 1}–${e + 1}`); s = e = a[i]; }
    return r.join('، ');
  };
  const baseOf = name => String(name || '').replace(/\.pdf$/i, '').trim() || 'فواتير';
  const safe = t => String(t || '').replace(/[\\/:*?"<>|\n\r\t]+/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 60);
  // الاسم داخل التطبيق وداخل الملف المضغوط: «ملف - فاتورة 2 - المورد - ص2-3.pdf»
  const nameOf = (S, g, k) => [baseOf(S.name), `فاتورة ${k + 1}`, g.qr && safe(g.qr.seller), 'ص' + rng(g.pages).replace(/–/g, '-').replace(/،\s*/g, ',')].filter(Boolean).join(' - ') + '.pdf';
  // اسم التنزيل المباشر بحروف إنجليزية: بعض المتصفحات تتجاهل الاسم العربي وتحفظ الملف بدون امتداد
  const asciiBase = S => baseOf(S.name).replace(/[^A-Za-z0-9 _.-]+/g, ' ').replace(/\s+/g, ' ').replace(/^[\s_.-]+|[\s_.-]+$/g, '').slice(0, 40);
  const dlName = (S, g, k) => [asciiBase(S), `invoice-${k + 1}`, 'p' + rng(g.pages).replace(/–/g, '-').replace(/،\s*/g, '_')].filter(Boolean).join('_') + '.pdf';

  async function makePdf(S, g) { // ننسخ صفحات الملف الأصلي كما هي (يبقى النص والجودة)
    const L = await lib();
    S.src = S.src || await L.PDFDocument.load(S.buf.slice(0), { ignoreEncryption: true });
    const out = await L.PDFDocument.create();
    (await out.copyPages(S.src, g.pages)).forEach(p => out.addPage(p));
    return out.save();
  }
  function download(data, name, type) {
    const u = URL.createObjectURL(new Blob([data], { type })), a = document.createElement('a');
    a.href = u; a.download = name; document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(u), 4000);
  }
  const CRC = (() => { const t = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
  const crc32 = b => { let c = 0xFFFFFFFF; for (let i = 0; i < b.length; i++) c = CRC[(c ^ b[i]) & 0xFF] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; };
  function zip(files) { // ملف ZIP بدون ضغط (ملفات PDF مضغوطة أصلًا)، والأسماء UTF-8
    const enc = new TextEncoder(), parts = [], cen = []; let off = 0;
    const h = (n, v) => { const a = new Uint8Array(n); for (let i = 0; i < n; i++) a[i] = (v >>> (8 * i)) & 255; return a; };
    const cat = a => { const n = a.reduce((s, x) => s + x.length, 0), o = new Uint8Array(n); let p = 0; a.forEach(x => { o.set(x, p); p += x.length; }); return o; };
    files.forEach(f => {
      const nm = enc.encode(f.name), d = f.data, c = crc32(d);
      const common = [h(2, 20), h(2, 0x0800), h(2, 0), h(2, 0), h(2, 0x21), h(4, c), h(4, d.length), h(4, d.length), h(2, nm.length), h(2, 0)];
      const loc = cat([h(4, 0x04034b50), ...common, nm]);
      cen.push(cat([h(4, 0x02014b50), h(2, 20), ...common, h(2, 0), h(2, 0), h(2, 0), h(4, 0), h(4, off), nm]));
      parts.push(loc, d); off += loc.length + d.length;
    });
    const cd = cat(cen);
    return cat([...parts, cd, h(4, 0x06054b50), h(2, 0), h(2, 0), h(2, files.length), h(2, files.length), h(4, cd.length), h(4, off), h(2, 0)]);
  }
  async function toFiles(S) { // كل فاتورة ملف PDF مستقل؛ وإن كانت فاتورة واحدة بكل الصفحات نعيد الملف الأصلي كما هو
    const gs = groupsOf(S);
    if (gs.length === 1 && gs[0].pages.length === S.pages.length && S.pages.length === S.total) return [S.file];
    const out = [];
    for (let k = 0; k < gs.length; k++) out.push(new File([await makePdf(S, gs[k])], nameOf(S, gs[k], k), { type: 'application/pdf', lastModified: S.file.lastModified || Date.now() }));
    return out;
  }

  // ---------- النافذة ----------
  const CSS = `
.spdlg{position:fixed;inset:0;z-index:56;background:#000b;display:flex;align-items:center;justify-content:center;padding:10px}
.spin{background:var(--paper);color:var(--ink);border-radius:8px;width:min(1100px,100%);max-height:100%;display:grid;grid-template-rows:auto minmax(0,1fr) auto;overflow:hidden}
.sphd{padding:14px 16px 10px;border-bottom:1px solid var(--cardline);display:grid;gap:2px}
.sphd h2{font-size:21px}.sphd small{color:var(--mute);font-size:13px;overflow-wrap:anywhere}
.spbody{overflow:auto;padding:14px 16px;display:grid;gap:14px;align-content:start;overscroll-behavior:contain}
.spft{padding:10px 16px;border-top:1px solid var(--cardline);display:flex;flex-wrap:wrap;gap:10px;align-items:center;background:var(--paper)}
.spft .grow{flex:1 1 auto}
.spmodes{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:10px;margin:0;padding:0;border:0}
.spmodes legend{font-weight:700;margin-bottom:8px;font-size:17px;padding:0}
.spmode{min-height:62px;flex-direction:column;gap:0;border-width:2px;font-size:16px;font-weight:700;text-align:center;padding:8px}
.spmode small{font-weight:400;font-size:13px;color:var(--mute)}
.spmode[aria-pressed=true]{background:var(--blue);border-color:var(--blue);color:#fff}.spmode[aria-pressed=true] small{color:#fff}
.sphint{margin:0;padding:8px 12px;background:#fffbd6;border:1px solid #111;border-radius:4px;font-size:15px;display:flex;flex-wrap:wrap;gap:8px;align-items:center}
.sphint button{min-height:36px;font-size:14px}
.spst{margin:0;color:var(--mute);font-size:14px}
.sppages{display:grid;grid-template-columns:repeat(auto-fill,minmax(150px,1fr));gap:10px}
.sppg{--gc:var(--line);border:3px solid var(--gc);border-radius:8px;background:var(--paper);display:grid;grid-template-rows:auto auto 1fr auto;overflow:hidden}
.sppg.skip{border-style:dashed;border-color:var(--line)}
.sppg .gh{background:var(--gc);color:#fff;font-weight:700;font-size:14px;padding:3px 8px;display:flex;justify-content:space-between;gap:6px}
.sppg .gh.cont{background:var(--paper);color:var(--gc);border-bottom:1px solid var(--cardline)}
.sppg .gh.none{background:var(--soft);color:var(--ink)}
.sppg.skip .gh{background:var(--line);color:var(--ink)}
.sppg .meta{display:flex;justify-content:space-between;align-items:center;padding:3px 8px;font-size:13px;color:var(--mute)}
.spqr{font-size:11px;font-weight:700;padding:0 6px;border-radius:3px;background:#111;color:#ffe600}
.spth{display:block;width:100%;padding:6px;border:0;border-radius:0;background:var(--soft);min-height:110px;cursor:zoom-in}
.spth img{display:block;width:100%;height:auto;box-shadow:0 0 0 1px var(--cardline);background:#fff}
.sppg.skip .spth img{opacity:.35}
.sppg .acts{display:grid;gap:6px;padding:8px}
.sppg .acts button{min-height:40px;font-size:14px;padding:6px 8px;width:100%}
.spstart[aria-pressed=true]{background:var(--gc);border-color:var(--gc);color:#fff;font-weight:700}
.spskip[aria-pressed=true]{background:#ffe600;border-color:#111;color:#111;font-weight:700}
.spsum h3{font-size:18px;margin-bottom:8px}
.spsum ul{display:grid;gap:8px;margin:0;padding:0;list-style:none}
.spsum li{border-inline-start:6px solid var(--gc);background:var(--soft);border-radius:4px;padding:8px 12px;display:grid;gap:2px}
.spsum small{color:var(--mute);font-size:14px}.spsum .w{color:var(--stamp);font-weight:500}
.spsum li .dl{justify-self:start;min-height:38px;font-size:14px;margin-top:4px}
.spsave{display:flex;flex-wrap:wrap;gap:10px;align-items:center;margin-top:10px}
.spsave span{color:var(--mute);font-size:14px}
.spall{display:flex;gap:8px;align-items:center;font-size:14px;flex-basis:100%}.spall input{width:20px;height:20px}
.splb{position:fixed;inset:0;z-index:57;background:#000d;display:flex;flex-direction:column}
.splb .bar3{display:flex;flex-wrap:wrap;gap:8px;align-items:center;padding:8px 10px;background:var(--paper)}
.splb .bar3 b{flex:1 1 160px}
.splb .stage{flex:1;overflow:auto;padding:10px;text-align:center}
.splb .stage img{max-width:100%;height:auto;background:#fff}
.splb .stage.big img{max-width:none;width:min(1600px,200%)}
@media (max-width:560px){.spmodes{grid-template-columns:1fr}.spmode{min-height:50px}.sppages{grid-template-columns:repeat(2,minmax(0,1fr));gap:8px}.sppg .acts button{font-size:13px}.spft button{flex:1 1 auto}}`;
  const GC = ['#1838b8', '#a34400', '#0b6a35', '#6a1b9a', '#8f0d18', '#00606b'];

  // يفتح الملف ويعرض نافذة التقسيم إن كان أكثر من صفحة.
  // النتيجة: { files: [File...], applyAll: 'one'|'each'|null } أو null إن اختار المستخدم تخطي الملف
  async function ask(file, opt) {
    opt = opt || {};
    if (!isPdf(file) || !window.pdfjsLib) return { files: [file] };
    const buf = await file.arrayBuffer();
    let doc; try { doc = await pdfjsLib.getDocument({ data: buf.slice(0) }).promise; } catch (e) { return { files: [file] }; }
    if (doc.numPages <= 1) { try { doc.destroy(); } catch { } return { files: [file] }; }
    const S = { file, name: file.name || 'ملف PDF', buf, total: doc.numPages, pages: [], mode: opt.mode || null };
    if (opt.mode === 'one') { try { doc.destroy(); } catch { } return { files: [file] }; }
    if (opt.mode === 'each') { // نفس اختيار الملف السابق في هذه الدفعة: بدون نافذة
      for (let i = 0; i < Math.min(doc.numPages, MAXP); i++) S.pages.push({});
      try { doc.destroy(); } catch { }
      try { return { files: await toFiles(S) }; } catch (e) { console.warn(e); return { files: [file] }; }
    }
    if (!document.getElementById('spCss')) { const st = document.createElement('style'); st.id = 'spCss'; st.textContent = CSS; document.head.appendChild(st); }
    lib().catch(() => { }); // نبدأ تحميل المكتبة مبكرًا

    return new Promise(resolve => {
      const last = document.activeElement;
      const el = document.createElement('div');
      el.className = 'spdlg'; el.setAttribute('role', 'dialog'); el.setAttribute('aria-modal', 'true'); el.setAttribute('aria-labelledby', 'spT');
      el.innerHTML = `<div class="spin">
        <div class="sphd"><h2 id="spT">الملف فيه ${doc.numPages} صفحات</h2><small>${esc(file.webkitRelativePath || file.name || '')}</small></div>
        <div class="spbody">
          <fieldset class="spmodes"><legend>كيف تريد قراءة هذا الملف؟</legend>
            <button type="button" class="spmode" data-spmode="one" aria-pressed="false">فاتورة واحدة<small>كل الصفحات فاتورة واحدة</small></button>
            <button type="button" class="spmode" data-spmode="each" aria-pressed="false">كل صفحة فاتورة<small>كل صفحة تُقرأ لوحدها</small></button>
            <button type="button" class="spmode" data-spmode="manual" aria-pressed="false">أحدد بنفسي<small>أختار بداية كل فاتورة</small></button>
          </fieldset>
          <p class="sphint" data-r="hint" hidden></p>
          <p class="spst" data-r="st" role="status">جارٍ تجهيز الصفحات…</p>
          <div class="sppages" data-r="pages"></div>
          <section class="spsum" data-r="sum" hidden></section>
        </div>
        <div class="spft">
          <button type="button" class="primary" data-r="ok" disabled>اختر طريقة القراءة</button>
          <button type="button" data-r="cancel">تخطي هذا الملف</button>
          <span class="grow"></span>
          ${opt.more ? `<label class="spall" data-r="allL" hidden><input type="checkbox" data-r="all"> <span>طبّق نفس الاختيار على باقي الملفات متعددة الصفحات في هذه الدفعة (${opt.more})</span></label>` : ''}
        </div></div>`;
      document.body.appendChild(el); document.body.style.overflow = 'hidden';
      const R = k => el.querySelector(`[data-r="${k}"]`);
      let closed = false, srcDoc = doc;
      const close = res => {
        if (closed) return; closed = true; el.remove(); lbClose(); document.body.style.overflow = '';
        try { srcDoc.destroy(); } catch { }
        if (last && document.contains(last)) try { last.focus(); } catch { }
        resolve(res);
      };

      function draw() {
        const man = S.mode === 'manual', gs = groupsOf(S);
        el.querySelectorAll('[data-spmode]').forEach(b => b.setAttribute('aria-pressed', b.dataset.spmode === S.mode));
        const qrs = S.pages.slice(0, -1).filter(p => p.qr).length;
        R('hint').hidden = !(man && qrs);
        if (man && qrs) R('hint').innerHTML = `<span>وُجد رمز QR في ${qrs} من الصفحات. يمكن أن يقترح البرنامج البدايات: الصفحة التي بعد رمز QR تبدأ فاتورة جديدة.</span><button type="button" data-spact="sug">طبّق الاقتراح</button><button type="button" data-spact="clr">مسح التحديد</button>`;
        const gOf = {}; gs.forEach(g => g.pages.forEach(i => { gOf[i] = g; }));
        let g = -1;
        R('pages').innerHTML = S.pages.map((p, i) => {
          const st = i === 0 || S.mode === 'each' || (man && p.mark), sk = man && p.skip; if (st) g++;
          const gc = S.mode ? GC[g % 6] : 'var(--line)';
          const head = !S.mode ? `<span>صفحة ${i + 1}</span>` : sk ? '<span>متجاهلة</span>' : `<span>فاتورة ${g + 1}</span><span>${st ? 'بداية' : 'تكملة'}</span>`;
          return `<article class="sppg${sk ? ' skip' : ''}" style="--gc:${gc}">
            <div class="gh${!S.mode ? ' none' : st ? '' : ' cont'}">${head}</div>
            <div class="meta"><span>صفحة ${i + 1}</span>${p.qr ? '<span class="spqr" title="وُجد رمز QR في هذه الصفحة">QR</span>' : ''}</div>
            <button type="button" class="spth" data-spview="${i}" aria-label="تكبير الصفحة ${i + 1}"><img src="${p.img}" alt=""></button>
            ${man ? `<div class="acts">
              ${i > 0 ? `<button type="button" class="spstart" data-spstart="${i}" aria-pressed="${!!p.mark}">${p.mark ? '✓ بداية فاتورة جديدة' : 'بداية فاتورة جديدة'}</button>` : ''}
              <button type="button" class="spskip" data-spskip="${i}" aria-pressed="${!!p.skip}">${p.skip ? 'متجاهلة: إلغاء' : 'تجاهل الصفحة'}</button>
            </div>` : ''}</article>`;
        }).join('');
        // الملخص والحفظ
        const sum = R('sum'); sum.hidden = !S.mode;
        if (S.mode) {
          const anyQr = S.pages.some(p => p.qr), skipped = man ? S.pages.filter(p => p.skip).length : 0;
          sum.innerHTML = `<h3>الفواتير التي ستُقرأ (${gs.length})</h3><ul>${gs.map((g, k) => {
            const w = [];
            if (g.pages.length > PER) w.push(`فيها ${g.pages.length} صفحات: ستُقرأ أول ${PER} فقط`);
            if (anyQr && !g.pages.some(i => S.pages[i].qr)) w.push('لا يوجد فيها رمز QR: تأكد أنها فاتورة كاملة');
            const q = g.qr ? `من رمز QR: ${esc(g.qr.seller)}${g.qr.total ? ` · الإجمالي ${(+g.qr.total).toLocaleString('en-US', { minimumFractionDigits: 2 })}` : ''}` : '';
            return `<li style="--gc:${GC[g.color % 6]}"><b>فاتورة ${k + 1}: ${g.pages.length === 1 ? 'الصفحة' : 'الصفحات'} ${rng(g.pages)}</b>${q ? `<small>${q}</small>` : ''}${w.map(t => `<small class="w">⚠ ${t}</small>`).join('')}${gs.length > 1 ? `<button type="button" class="dl" data-spdl="${k}">تنزيل هذه الفاتورة PDF</button>` : ''}</li>`;
          }).join('') || '<li><b>كل الصفحات متجاهلة</b></li>'}</ul>
          ${gs.length > 1 ? `<div class="spsave"><button type="button" data-spact="zip">تنزيل الفواتير كملفات منفصلة (ZIP)</button><span data-r="saveNote" role="status">${skipped ? `${skipped} ${skipped === 1 ? 'صفحة متجاهلة لن تُحفظ' : 'صفحات متجاهلة لن تُحفظ'}` : ''}</span></div>` : ''}`;
        }
        const ok = R('ok'); ok.disabled = !S.mode || !gs.length || S.loading;
        ok.textContent = !S.mode ? 'اختر طريقة القراءة' : S.loading ? 'انتظر تجهيز الصفحات…' : gs.length > 1 ? `قراءة ${gs.length} فواتير` : 'قراءة الفاتورة';
        const al = R('allL'); if (al) al.hidden = !(S.mode === 'one' || S.mode === 'each');
      }
      const note = t => { const n = R('saveNote'); if (n) n.textContent = t; };

      el.addEventListener('click', async e => {
        const b = e.target.closest('button'); if (!b) return;
        const d = b.dataset;
        if (d.spmode) { S.mode = d.spmode; draw(); }
        else if (d.spstart != null) { const p = S.pages[+d.spstart]; p.mark = !p.mark; draw(); }
        else if (d.spskip != null) { const p = S.pages[+d.spskip]; p.skip = !p.skip; draw(); }
        else if (d.spview != null) lbOpen(+d.spview);
        else if (d.spact === 'sug') { S.pages.forEach((p, i) => { p.mark = i > 0 && !!(S.pages[i - 1].qr); }); draw(); }
        else if (d.spact === 'clr') { S.pages.forEach(p => { p.mark = false; p.skip = false; }); draw(); }
        else if (d.spdl != null) {
          const gs = groupsOf(S), k = +d.spdl;
          try { download(await makePdf(S, gs[k]), dlName(S, gs[k], k), 'application/pdf'); note('تم تنزيل: ' + dlName(S, gs[k], k)); } catch (err) { console.error(err); note('تعذّر إنشاء الملف.'); }
        } else if (d.spact === 'zip') {
          const gs = groupsOf(S); b.disabled = true;
          try {
            const files = [];
            for (let k = 0; k < gs.length; k++) { note(`جارٍ تجهيز فاتورة ${k + 1} من ${gs.length}…`); files.push({ name: nameOf(S, gs[k], k), data: await makePdf(S, gs[k]) }); }
            download(zip(files), (asciiBase(S) || 'invoices') + '_split.zip', 'application/zip');
            note(`تم تنزيل ملف مضغوط فيه ${gs.length} ملفات PDF.`);
          } catch (err) { console.error(err); note('تعذّر إنشاء الملفات.'); }
          b.disabled = false;
        } else if (b === R('cancel')) close(null);
        else if (b === R('ok')) {
          b.disabled = true; b.textContent = 'جارٍ تجهيز الفواتير…';
          const all = R('all'), applyAll = all && all.checked && (S.mode === 'one' || S.mode === 'each') ? S.mode : null;
          try { close({ files: S.mode === 'one' ? [file] : await toFiles(S), applyAll }); }
          catch (err) { console.error(err); R('st').textContent = 'تعذّر تقسيم الملف. يمكنك اختيار «فاتورة واحدة».'; b.disabled = false; draw(); }
        }
      });

      // تكبير صفحة
      const lb = document.createElement('div'); lb.className = 'splb'; lb.hidden = true; lb.setAttribute('role', 'dialog'); lb.setAttribute('aria-modal', 'true');
      lb.innerHTML = `<div class="bar3"><b></b><button type="button" data-lb="p">→ السابقة</button><button type="button" data-lb="n">التالية ←</button><button type="button" data-lb="z" aria-pressed="false">تكبير</button><button type="button" data-lb="x">إغلاق</button></div><div class="stage"></div>`;
      document.body.appendChild(lb);
      let lbI = 0;
      function lbShow() {
        lb.querySelector('b').textContent = `صفحة ${lbI + 1} من ${S.pages.length}`;
        lb.querySelector('.stage').innerHTML = `<img src="${S.pages[lbI].img}" alt="الصفحة ${lbI + 1}">`;
        lb.querySelector('[data-lb="p"]').disabled = lbI === 0; lb.querySelector('[data-lb="n"]').disabled = lbI === S.pages.length - 1;
      }
      function lbOpen(i) { lbI = i; lb.hidden = false; lbShow(); lb.querySelector('[data-lb="x"]').focus(); }
      function lbClose() { if (!lb.hidden) { lb.hidden = true; const t = el.querySelector(`[data-spview="${lbI}"]`); if (t) t.focus(); } if (closed) lb.remove(); }
      lb.addEventListener('click', e => {
        const b = e.target.closest('[data-lb]'); if (!b) return;
        const k = b.dataset.lb;
        if (k === 'x') lbClose(); else if (k === 'p' && lbI > 0) { lbI--; lbShow(); } else if (k === 'n' && lbI < S.pages.length - 1) { lbI++; lbShow(); }
        else if (k === 'z') { const on = lb.querySelector('.stage').classList.toggle('big'); b.setAttribute('aria-pressed', on); b.textContent = on ? 'تصغير' : 'تكبير'; }
      });
      lb.addEventListener('keydown', e => {
        if (e.key === 'Escape') { e.stopPropagation(); lbClose(); }
        else if (e.key === 'ArrowLeft') lb.querySelector('[data-lb="n"]').click();
        else if (e.key === 'ArrowRight') lb.querySelector('[data-lb="p"]').click();
      });

      // تجهيز الصفحات المصغّرة تدريجيًا مع البحث عن رمز QR
      (async () => {
        const n = Math.min(doc.numPages, MAXP); S.loading = true; draw();
        for (let p = 1; p <= n && !closed; p++) {
          R('st').textContent = `جارٍ تجهيز الصفحة ${p} من ${n}…`;
          try {
            const pg = await doc.getPage(p), v0 = pg.getViewport({ scale: 1 });
            const v = pg.getViewport({ scale: Math.min(2.5, 1400 / Math.max(v0.width, v0.height)) });
            const c = document.createElement('canvas'); c.width = Math.round(v.width); c.height = Math.round(v.height);
            const x = c.getContext('2d'); x.fillStyle = '#fff'; x.fillRect(0, 0, c.width, c.height);
            await pg.render({ canvasContext: x, viewport: v }).promise;
            let qr = null; try { qr = scanQR(c); } catch { }
            S.pages.push({ img: c.toDataURL('image/jpeg', 0.75), qr, mark: false, skip: false });
            c.width = c.height = 0;
          } catch (e) { console.warn(e); S.pages.push({ img: '', qr: null, mark: false, skip: false }); }
          if (p === 1 || p % 3 === 0 || p === n) draw();
        }
        if (closed) return;
        S.loading = false;
        R('st').textContent = doc.numPages > MAXP ? `تُعرض أول ${MAXP} صفحة فقط من ${doc.numPages}.` : 'اضغط على أي صفحة لتكبيرها.';
        draw();
      })();
      el.querySelector('[data-spmode="one"]').focus();
    });
  }

  window.InvoiceSplit = { ask, isPdf };
})();
