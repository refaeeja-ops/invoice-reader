/* قارئ الفواتير: حفظ الفواتير والبيانات المرجعية في IndexedDB داخل المتصفح (مساحة أكبر بكثير من localStorage).
   إن لم يتوفر IndexedDB (وضع خاص قديم مثلًا) نرجع إلى localStorage. */
(function (root) {
  'use strict';
  const LS_INV = 'invoice-reader-history-v1', LS_KV = 'invoice-reader-kv-';
  let dbp = null;
  function open() {
    if (dbp) return dbp;
    dbp = new Promise((ok, no) => {
      try {
        const r = indexedDB.open('invoice-reader-data', 1);
        r.onupgradeneeded = () => { const d = r.result; if (!d.objectStoreNames.contains('inv')) d.createObjectStore('inv', { keyPath: '_id' }); if (!d.objectStoreNames.contains('kv')) d.createObjectStore('kv'); };
        r.onsuccess = () => ok(r.result); r.onerror = () => no(r.error); r.onblocked = () => no(new Error('blocked'));
      } catch (e) { no(e); }
    });
    return dbp;
  }
  const req = r => new Promise((ok, no) => { r.onsuccess = () => ok(r.result); r.onerror = () => no(r.error); });
  async function tx(store, mode, fn) {
    const d = await open();
    return new Promise((ok, no) => {
      const t = d.transaction(store, mode), st = t.objectStore(store); let out;
      Promise.resolve(fn(st)).then(v => { out = v; }, no);
      t.oncomplete = () => ok(out); t.onerror = () => no(t.error); t.onabort = () => no(t.error || new Error('abort'));
    });
  }
  // بديل 1: localStorage عند تعذّر IndexedDB. بديل 2: الذاكرة فقط (حين يمنع المتصفح كل أنواع الحفظ،
  // مثل Safari مع «حظر كل ملفات تعريف الارتباط»: يرمي SecurityError «The operation is insecure»)
  const mem = { inv: [], kv: {} };
  const lsInv = () => { if (mode === 'mem') return mem.inv; try { return JSON.parse(localStorage.getItem(LS_INV)) || []; } catch { return []; } };
  const lsSave = l => { if (mode === 'mem') { mem.inv = l; return; } localStorage.setItem(LS_INV, JSON.stringify(l)); };
  let mode = null, why = ''; // 'idb' أو 'ls' أو 'mem'
  const errTxt = e => (e && (e.name || '') + (e.message ? ': ' + e.message : '')) || String(e);
  async function init() {
    if (mode) return mode;
    try { // اختبار حقيقي: فتح القاعدة وكتابة وحذف
      await open();
      await tx('kv', 'readwrite', st => { st.put(1, '__test'); st.delete('__test'); });
      mode = 'idb';
    } catch (e) {
      console.warn('IndexedDB غير متاح', e); why = errTxt(e); dbp = null;
      try { localStorage.setItem('__ir_test', '1'); localStorage.removeItem('__ir_test'); mode = 'ls'; }
      catch (e2) {
        console.warn('localStorage غير متاح', e2); why += ' / ' + errTxt(e2); mode = 'mem';
        // إن كانت القراءة ممكنة (مثل امتلاء المساحة) نعرض الفواتير الموجودة ولا نحذف شيئًا
        try { mem.inv = JSON.parse(localStorage.getItem(LS_INV)) || []; } catch { mem.inv = []; }
      }
    }
    return mode;
  }
  const api = {
    init, get mode() { return mode; }, get why() { return why; },
    async allInv() { await init(); if (mode !== 'idb') return lsInv(); return tx('inv', 'readonly', st => req(st.getAll())); },
    async getInv(id) { await init(); if (mode !== 'idb') return lsInv().find(x => x._id === id) || null; return tx('inv', 'readonly', st => req(st.get(id))); },
    async putInv(list) {
      await init();
      if (mode !== 'idb') { const l = lsInv(); list.forEach(x => { const i = l.findIndex(y => y._id === x._id); if (i >= 0) l[i] = x; else l.unshift(x); }); lsSave(l); return; }
      return tx('inv', 'readwrite', st => { list.forEach(x => st.put(x)); });
    },
    async delInv(ids) {
      await init();
      if (mode !== 'idb') { const s = new Set(ids); lsSave(lsInv().filter(x => !s.has(x._id))); return; }
      return tx('inv', 'readwrite', st => { ids.forEach(id => st.delete(id)); });
    },
    async getKv(k) {
      await init();
      if (mode === 'mem') return mem.kv[k] ?? null;
      if (mode === 'ls') { try { return JSON.parse(localStorage.getItem(LS_KV + k)); } catch { return null; } }
      return tx('kv', 'readonly', st => req(st.get(k)));
    },
    async setKv(k, v) {
      await init();
      if (mode === 'mem') { if (v == null) delete mem.kv[k]; else mem.kv[k] = v; return; }
      if (mode === 'ls') { if (v == null) localStorage.removeItem(LS_KV + k); else localStorage.setItem(LS_KV + k, JSON.stringify(v)); return; }
      return tx('kv', 'readwrite', st => { if (v == null) st.delete(k); else st.put(v, k); });
    }
  };
  root.InvoiceStore = api;
})(typeof self !== 'undefined' ? self : this);
