/* قارئ الفواتير: قراءة ملف البيانات المرجعية (Excel) داخل المتصفح، ومطابقة الفواتير مع الموردين والمشاريع والمنتجات.
   الملف لا يُرفع لأي مكان: يُقرأ هنا ويُحفظ في هذا المتصفح فقط. */
(function (root) {
  'use strict';

  // ---------- قارئ xlsx بسيط (ZIP + XML) بدون مكتبات ----------
  async function unzip(buf) {
    const dv = new DataView(buf), td = new TextDecoder();
    let e = buf.byteLength - 22;
    while (e >= 0 && dv.getUint32(e, true) !== 0x06054b50) e--;
    if (e < 0) throw new Error('الملف ليس ملف Excel (xlsx) صالحًا.');
    const n = dv.getUint16(e + 10, true); let off = dv.getUint32(e + 16, true);
    const files = {};
    for (let i = 0; i < n; i++) {
      const method = dv.getUint16(off + 10, true), csize = dv.getUint32(off + 20, true);
      const nl = dv.getUint16(off + 28, true), xl = dv.getUint16(off + 30, true), cl = dv.getUint16(off + 32, true), lo = dv.getUint32(off + 42, true);
      files[td.decode(new Uint8Array(buf, off + 46, nl))] = { method, csize, lo };
      off += 46 + nl + xl + cl;
    }
    return async name => {
      const f = files[name]; if (!f) return null;
      const lnl = dv.getUint16(f.lo + 26, true), lxl = dv.getUint16(f.lo + 28, true);
      const data = new Uint8Array(buf, f.lo + 30 + lnl + lxl, f.csize);
      if (f.method === 0) return td.decode(data);
      if (typeof DecompressionStream === 'undefined') throw new Error('المتصفح قديم ولا يدعم قراءة ملفات Excel. حدّث المتصفح.');
      return new Response(new Blob([data]).stream().pipeThrough(new DecompressionStream('deflate-raw'))).text();
    };
  }
  const xml = s => new DOMParser().parseFromString(s, 'application/xml');
  const tags = (node, name) => [...node.getElementsByTagName(name)];
  const colIdx = ref => { const m = /^[A-Z]+/.exec(ref); let n = 0; for (const ch of m[0]) n = n * 26 + ch.charCodeAt(0) - 64; return n - 1; };

  async function readXlsx(buf) {
    const get = await unzip(buf);
    const wb = xml(await get('xl/workbook.xml'));
    const rels = xml(await get('xl/_rels/workbook.xml.rels') || '<r/>');
    const target = {}; tags(rels, 'Relationship').forEach(r => target[r.getAttribute('Id')] = r.getAttribute('Target'));
    const sstX = await get('xl/sharedStrings.xml');
    const sst = sstX ? tags(xml(sstX), 'si').map(si => tags(si, 't').filter(t => t.parentNode.nodeName !== 'rPh').map(t => t.textContent).join('')) : [];
    const out = {};
    for (const sh of tags(wb, 'sheet')) {
      const name = sh.getAttribute('name');
      let t = target[sh.getAttribute('r:id')] || ''; t = t.replace(/^\/?xl\//, '').replace(/^\//, '');
      const sx = await get('xl/' + t); if (!sx) continue;
      const rows = [];
      tags(xml(sx), 'row').forEach(r => {
        const arr = [];
        tags(r, 'c').forEach(c => {
          const ty = c.getAttribute('t'), v = c.getElementsByTagName('v')[0], vs = v ? v.textContent : null;
          let val = null;
          if (ty === 's') val = vs == null ? null : sst[+vs];
          else if (ty === 'b') val = vs === '1';
          else if (ty === 'inlineStr') val = tags(c, 't').map(x => x.textContent).join('');
          else if (ty === 'str' || ty === 'e') val = vs;
          else if (vs != null && vs !== '') val = +vs;
          arr[colIdx(c.getAttribute('r'))] = val;
        });
        rows[(+r.getAttribute('r') || rows.length + 1) - 1] = arr;
      });
      const head = (rows[0] || []).map(h => String(h ?? '').trim());
      out['#' + name] = head;
      out[name] = rows.slice(1).filter(r => r && r.some(v => v != null && v !== '')).map(r => { const o = {}; head.forEach((h, i) => { if (h) o[h] = r[i] ?? null; }); return o; });
    }
    return out;
  }

  // ---------- تحويل الأوراق إلى بيانات مرجعية ----------
  const str = v => String(v ?? '').trim().replace(/\.0+$/, '');
  const int = v => { const s = str(v); return /^-?\d+$/.test(s) ? +s : null; };
  const digits = v => str(v).replace(/[٠-٩]/g, d => '٠١٢٣٤٥٦٧٨٩'.indexOf(d)).replace(/\D/g, '');
  const yes = v => v === true || /^(true|1|نعم)$/i.test(String(v ?? '').trim());
  const no = v => v === false || /^(false|0|لا)$/i.test(String(v ?? '').trim());
  const invKey = s => str(s).toLowerCase().replace(/[^0-9a-zء-ي]/g, '').replace(/^0+(?=.)/, '');
  const pick = (book, name) => { const k = Object.keys(book).find(x => x[0] !== '#' && x.trim().toLowerCase() === name.toLowerCase()); return k ? book[k] : null; };

  function build(book, fileName) {
    const need = ['SuppliersT', 'ProjectsT', 'ProductsT', 'SupplierProductT'];
    const missing = need.filter(n => !pick(book, n));
    if (missing.length) throw new Error('الملف لا يحتوي الأوراق: ' + missing.join('، '));
    const suppliers = pick(book, 'SuppliersT').map(r => ({ id: int(r.SupID), name: str(r.SupfName), vat: digits(r.VATID), active: !no(r.Active) })).filter(s => s.id != null);
    const projects = pick(book, 'ProjectsT').map(r => ({ id: int(r.ProjectID), name: str(r.ProjectName), active: yes(r.ProjectStatus) })).filter(p => p.id != null);
    const products = pick(book, 'ProductsT').map(r => ({ id: int(r.ProductID), name: str(r.ProductName), merge: int(r.MergTargetProductID) })).filter(p => p.id != null);
    const supProd = pick(book, 'SupplierProductT').filter(r => !no(r.Active)).map(r => [int(r.SupID), int(r.ProductID)]).filter(x => x[0] != null && x[1] != null);
    // أعمدة ورقة VatReport كما في ملف المستخدم. «ProductName» المكرر بعد ProjectID يُعامل كاسم المشروع.
    let vatCols = null; const vk = Object.keys(book).find(k => k.trim().toLowerCase() === '#vatreport');
    if (vk) {
      vatCols = book[vk].filter(Boolean);
      const pi = vatCols.findIndex(c => /^projectid$/i.test(c));
      if (pi >= 0 && /^productname$/i.test(vatCols[pi + 1] || '') && vatCols.filter(c => /^productname$/i.test(c)).length > 1 && !vatCols.some(c => /^projectname$/i.test(c))) vatCols[pi + 1] = 'ProjectName';
      if (!vatCols.length) vatCols = null;
    }
    const inv = {};
    (pick(book, 'InvoiceT') || []).forEach(r => { const s = int(r.SupID), k = invKey(r.InvoiceNo); if (s != null && k) inv[s + '|' + k] = int(r.InvID) ?? true; });
    return { fileName, loadedAt: new Date().toISOString(), suppliers, projects, products, supProd, inv, vatCols };
  }

  async function load(file) {
    const book = await readXlsx(await file.arrayBuffer());
    return build(book, file.name);
  }

  // ---------- المطابقة ----------
  const nName = s => String(s || '').replace(/[٠-٩]/g, d => '٠١٢٣٤٥٦٧٨٩'.indexOf(d)).toLowerCase()
    .replace(/[أإآ]/g, 'ا').replace(/ة/g, 'ه').replace(/ى/g, 'ي')
    .replace(/(^|\s)(شركه|شركة|مؤسسه|مؤسسة|مصنع|المحدوده|المحدودة|ذات المسؤوليه المحدوده|للتجاره|التجاريه|co|company|ltd|llc|est|trading)(?=\s|$)/g, ' ')
    .replace(/[^0-9a-zء-ي]/g, '');

  function matchSupplier(ref, x) {
    if (!ref) return null;
    const v = digits(x.vendor_tax_number);
    if (v.length >= 9) {
      for (const s of ref.suppliers) {
        if (!s.vat || s.vat.length < 9) continue;
        if (v === s.vat || v.startsWith(s.vat) || s.vat.startsWith(v) || (v.length >= 10 && s.vat.length >= 10 && v.slice(0, 10) === s.vat.slice(0, 10))) return { id: s.id, by: 'vat' };
      }
    }
    const n = nName(x.vendor_name);
    if (n.length >= 4) {
      let s = ref.suppliers.find(s => nName(s.name) === n);
      if (!s) s = ref.suppliers.find(s => { const m = nName(s.name); return m.length >= 5 && (m.includes(n) || n.includes(m)); });
      if (s) return { id: s.id, by: 'name' };
    }
    return null;
  }
  const supplier = (ref, id) => ref && ref.suppliers.find(s => s.id === id) || null;
  const project = (ref, id) => ref && ref.projects.find(p => p.id === id) || null;
  const product = (ref, id) => ref && ref.products.find(p => p.id === id) || null;
  const activeProjects = ref => ref ? ref.projects.filter(p => p.active) : [];
  // منتجات المورّد من SupplierProductT (مع تحويل المنتجات المدمجة إلى المنتج الهدف)
  function productsFor(ref, supId) {
    if (!ref) return [];
    const ids = [...new Set(ref.supProd.filter(([s]) => s === supId).map(([, p]) => { const pr = product(ref, p); return pr && pr.merge ? pr.merge : p; }))];
    return ids.map(id => product(ref, id) || { id, name: 'منتج ' + id });
  }
  const existsInvoice = (ref, supId, invNo) => ref && supId != null && invKey(invNo) ? ref.inv[supId + '|' + invKey(invNo)] || null : null;

  root.InvoiceRef = { load, readXlsx, build, matchSupplier, supplier, project, product, activeProjects, productsFor, existsInvoice, digits, int };
})(typeof self !== 'undefined' ? self : this);
