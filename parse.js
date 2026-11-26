/* قارئ الفواتير: تحليل رمز QR (هيئة الزكاة) ونص الفاتورة بقواعد بسيطة. مجاني وبدون خوادم. */
(function (root) {
  'use strict';

  // ---------- أرقام ----------
  const AR_DIGITS = '٠١٢٣٤٥٦٧٨٩', FA_DIGITS = '۰۱۲۳۴۵۶۷۸۹';
  function normDigits(s) {
    return String(s || '')
      .replace(/[٠-٩]/g, d => AR_DIGITS.indexOf(d))
      .replace(/[۰-۹]/g, d => FA_DIGITS.indexOf(d))
      .replace(/٫/g, '.').replace(/٬/g, ',')
      .replace(/[\u200e\u200f\u202a-\u202e\u2066-\u2069]/g, '')
      .replace(/ی/g, 'ي').replace(/ک/g, 'ك').replace(/ە/g, 'ه') // حروف فارسية يخرجها نموذج الخط العربي
      .replace(/(\d)ر(\d)/g, '$1.$2') // الفاصلة العشرية العربية تُقرأ أحيانًا «ر»
      .replace(/(\d\.\d{2})(?=\d+\.\d)/g, '$1 '); // "25.0050.00" -> "25.00 50.00" (أعمدة ملتصقة)
  }
  // يحوّل "1,234.50" أو "1 234,50" إلى رقم
  function toNum(tok) {
    if (tok == null) return null;
    let t = normDigits(tok).replace(/[^\d.,-]/g, '');
    if (!/\d/.test(t)) return null;
    const lastDot = t.lastIndexOf('.'), lastComma = t.lastIndexOf(',');
    if (lastComma > lastDot) {
      // فاصلة عشرية أوروبية إذا تبعها رقمان فقط
      if (/,\d{1,2}$/.test(t) && lastDot === -1) t = t.replace(/,/g, '.');
      else if (/,\d{1,2}$/.test(t)) t = t.replace(/\./g, '').replace(',', '.');
      else t = t.replace(/,/g, '');
    } else t = t.replace(/,/g, '');
    const n = parseFloat(t);
    return isFinite(n) ? n : null;
  }
  const NUM_RE = /-?\d{1,3}(?:[,\u202f]\d{3})+(?:\.\d{1,3})?|-?\d+(?:[.,]\d{1,3})?/g;
  // وضع الأرقام الهندية: قارئ النص قد يعكس ترتيب أجزاء الرقم (مثل ۰۰ر۱۵ بدل ۱۵٫۰۰)،
  // فنولّد بدائل محتملة لكل رقم ونختار ما تتسق به الحسابات.
  let EASTERN = false;
  const rev = t => t.split('').reverse().join('');
  function altsOf(raw, n) {
    const out = [n];
    if (!EASTERN) return out;
    const t = raw.replace(/,/g, '');
    const m = t.match(/^(-?)(\d+)\.(\d+)$/);
    const add = v => { const x = parseFloat(v); if (isFinite(x) && !out.some(y => Math.abs(y - x) < 1e-9)) out.push(x); };
    if (m) {
      const [, , a, b] = m;
      if (/^0{2,}$/.test(a)) { out.length = 0; add(b + '.' + a); add(rev(b) + '.' + a); add(t); } // «۰۰ر۱۵» = ۱۵٫۰۰
      add(b + '.' + a); add(rev(b) + '.' + a); add(rev(b) + '.' + rev(a)); add(rev(a) + '.' + b);
      if (b.length > 2) add(b.slice(0, -2) + '.' + b.slice(-2));
    } else if (/^\d{2,4}$/.test(t)) add(rev(t));
    return out;
  }
  function numsIn(line) {
    const out = [];
    const s = normDigits(line);
    let m; NUM_RE.lastIndex = 0;
    while ((m = NUM_RE.exec(s))) { const n = toNum(m[0].replace(/\s/g, '')); if (n !== null) { const alts = altsOf(m[0], n); out.push({ n: alts[0], raw: m[0], i: m.index, alts }); } }
    return out;
  }
  const round2 = n => Math.round(n * 100) / 100;
  const near = (a, b, tol) => Math.abs(a - b) <= Math.max(tol || 0.05, Math.abs(b) * 0.002);

  // ---------- رمز QR الخاص بهيئة الزكاة (TLV بترميز Base64) ----------
  function b64ToBytes(b64) {
    const bin = (typeof atob === 'function') ? atob(b64) : Buffer.from(b64, 'base64').toString('binary');
    const u = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i); return u;
  }
  function utf8(bytes) {
    if (typeof TextDecoder !== 'undefined') return new TextDecoder('utf-8').decode(bytes);
    return Buffer.from(bytes).toString('utf8');
  }
  function parseZatcaQR(text) {
    if (!text) return null;
    const clean = String(text).trim().replace(/\s+/g, '');
    if (!/^[A-Za-z0-9+/=_-]+$/.test(clean) || clean.length < 20) return null;
    let bytes;
    try { bytes = b64ToBytes(clean.replace(/-/g, '+').replace(/_/g, '/')); } catch (e) { return null; }
    const tags = {}; let i = 0;
    while (i + 2 <= bytes.length) {
      const tag = bytes[i], len = bytes[i + 1];
      if (!tag || i + 2 + len > bytes.length) break;
      tags[tag] = bytes.slice(i + 2, i + 2 + len); i += 2 + len;
    }
    if (!tags[1] || !tags[2]) return null;
    const str = t => tags[t] ? utf8(tags[t]) : '';
    const r = {
      seller: str(1).trim(), vat_number: normDigits(str(2)).trim(),
      timestamp: str(3).trim(), total: toNum(str(4)), vat: toNum(str(5)),
      phase: tags[6] ? 2 : 1
    };
    if (!/^\d{10,20}$/.test(r.vat_number)) return null;
    return r;
  }

  // ---------- التحقق من الرقم الضريبي ----------
  function checkVat(v) {
    const d = normDigits(v || '').replace(/\D/g, '');
    if (!d) return null;
    if (d.length !== 15) return 'الرقم الضريبي يجب أن يكون 15 رقمًا (المقروء ' + d.length + ' رقمًا).';
    if (d[0] === '3' && d[14] !== '3') return 'الرقم الضريبي السعودي يبدأ بـ 3 وينتهي بـ 3، والمقروء لا ينتهي بـ 3.';
    if (d[0] !== '3' && d.slice(0, 3) !== '100') return 'الرقم الضريبي لا يطابق صيغة السعودية (يبدأ وينتهي بـ 3) ولا الإمارات (يبدأ بـ 100).';
    return null;
  }

  // ---------- الكلمات المفتاحية ----------
  const K = {
    invNo: /(رقم\s*الفاتور[ةه]|رقم\s*فاتور[ةه]|فاتور[ةه]\s*(?:ضريبي[ةه]\s*)?(?:مبسط[ةه]\s*)?رقم|رقم\s*الإيصال|رقم\s*الايصال|رقم\s*السند|رقم\s*المستند|(?:tax\s*)?invoice\s*(?:n[o0]|number|num|#|id)|inv\.?\s*(?:n[o0]|#|number)|bill\s*(?:n[o0]|#|number)|receipt\s*(?:n[o0]|#|number)|doc(?:ument)?\.?\s*(?:n[o0]|#|number)|ref(?:erence)?\s*(?:n[o0]|#))\.?/i,
    date: /(تاريخ\s*الفاتور[ةه]|تاريخ\s*الإصدار|تاريخ\s*الاصدار|التاريخ|تاريخ|invoice\s*date|issue\s*date|date)/i,
    due: /(تاريخ\s*الاستحقاق|due\s*date|payment\s*due)/i,
    vat: /(الرقم\s*الضريبي|رقم\s*التسجيل\s*الضريبي|الرقم\s*الضريبى|vat\s*(?:reg(?:istration)?\.?\s*)?(?:no|number|#)|tax\s*(?:id|no|number|registration)|trn)/i,
    customer: /(اسم\s*العميل|العميل|المشتري|عميل|customer|bill\s*to|client|buyer|sold\s*to)/i,
    grand: /(الإجمالي\s*شامل|الاجمالي\s*شامل|شامل\s*الضريب[ةه]|الإجمالي\s*المستحق|الاجمالي\s*المستحق|المبلغ\s*المستحق|صافي\s*المبلغ|الصافي|المجموع\s*الكلي|الإجمالي\s*الكلي|الاجمالي\s*الكلي|grand\s*total|total\s*(?:amount\s*)?(?:incl|including|inc\.|with)|amount\s*due|net\s*amount|balance\s*due|total\s*due|total\s*payable)/i,
    sub: /(قبل\s*الضريب[ةه]|غير\s*شامل|بدون\s*ضريب[ةه]|المجموع\s*الفرعي|الإجمالي\s*الفرعي|الاجمالي\s*الفرعي|sub\s*-?\s*total|total\s*(?:excl|excluding|before|w\/o|without)|taxable\s*amount|net\s*total)/i,
    tax: /(ضريب[ةه]\s*القيم[ةه]\s*المضاف[ةه]|قيم[ةه]\s*الضريب[ةه]|مبلغ\s*الضريب[ةه]|الضريب[ةه]|\bvat\b|\btax\b)/i,
    total: /(الإجمالي|الاجمالي|المجموع|الاجمال|إجمالي|اجمالي|\btotal\b|\bamount\b)/i,
    discount: /(الخصم|خصم|discount)/i,
    pay: /(نقد[اًا]?|نقدي|كاش|شبك[ةه]|مدى|بطاق[ةه]|فيزا|ماستر|تحويل\s*بنكي|cash|card|visa|mastercard|mada|bank\s*transfer|apple\s*pay|stc\s*pay)/i,
    skipItem: /(القيم[ةه]\s*المضاف[ةه]|الإجمالي|الاجمالي|المجموع|ضريب|الخصم|خصم|total|subtotal|vat|tax|discount|amount\s*due|balance|change|الباقي|المدفوع|paid|cash|نقد|شبكة|مدى|رقم|تاريخ|date|invoice|فاتورة|هاتف|جوال|phone|tel|mobile|س\.ت|السجل|c\.?r|www|http|@|ص\.ب|p\.?o\.?\s*box|الرمز البريدي|iban|ايبان)/i,
    header: /(الصنف|الوصف|البيان|الكمية|السعر|المنتج|description|item|qty|quantity|unit\s*price|price|amount)/i
  };
  // «ى» تُكتب أحيانًا بدل «ي» في آخر الكلمة (الإجمالى): نقبل الاثنين في الكلمات المفتاحية
  Object.keys(K).forEach(k => { K[k] = new RegExp(K[k].source.replace(/ي/g, '[يى]'), K[k].flags); });
  const CUR = [
    [/(ر\.\s?س|\dر\.?س|ريال\s*سعودي|ريال(?!\s*(قطري|عماني|يمني))|﷼|\bSAR\b|\bSR\b)/gi, 'SAR'],
    [/(د\.\s?إ|درهم|\bAED\b)/gi, 'AED'], [/(د\.\s?ك|دينار\s*كويتي|\bKWD\b)/gi, 'KWD'],
    [/(ر\.\s?ق|ريال\s*قطري|\bQAR\b)/gi, 'QAR'], [/(ر\.\s?ع|ريال\s*عماني|\bOMR\b)/gi, 'OMR'],
    [/(د\.\s?ب|دينار\s*بحريني|\bBHD\b)/gi, 'BHD'], [/(ج\.\s?م|جنيه|\bEGP\b)/gi, 'EGP'],
    [/(دينار\s*أردني|\bJOD\b)/gi, 'JOD'], [/(€|\bEUR\b)/gi, 'EUR'], [/(£|\bGBP\b)/gi, 'GBP'], [/(\$|\bUSD\b)/gi, 'USD']
  ];

  function lastTok(line, opts) {
    const ns = numsIn(line).filter(x => !(opts && opts.noPercent && /%/.test(line.slice(x.i + x.raw.length, x.i + x.raw.length + 2))));
    const money = ns.filter(x => !/^\d{9,}$/.test(x.raw)); // تجاهل الأرقام الطويلة (أرقام ضريبية/هواتف)
    if (!money.length) return null;
    // نفضّل الأرقام ذات الكسور العشرية
    const dec = money.filter(x => /[.,]\d{1,2}$/.test(x.raw));
    return (dec.length ? dec : money).reduce((a, b) => (b.n > a.n ? b : a));
  }
  const lastNum = (line, opts) => { const t = lastTok(line, opts); return t ? t.n : null; };
  function valueAfter(line, re) {
    const m = line.match(re); if (!m) return '';
    let rest = line.slice(m.index + m[0].length).replace(/^[\s:：#.\-–|/]+/, '').trim();
    if (!rest) { // القيمة قبل التسمية (ترتيب عربي معكوس)
      rest = line.slice(0, m.index).replace(/[\s:：#.\-–|/]+$/, '').trim();
    }
    return rest;
  }
  // ---------- رقم الفاتورة ----------
  // المرشحون: كلمات من حروف لاتينية وأرقام تفصلها - / : . _ # وفيها رقم واحد على الأقل (مثل INV-1001، SIF-DM:26111020، 02/2026/00576)
  function invCands(s) {
    const out = [], re = /(?:[A-Za-z]{1,4}\s(?=\d))?[A-Za-z0-9]+(?:\s?[-_/:.#]\s?[A-Za-z0-9]+)*/g; let m;
    while ((m = re.exec(s))) {
      let t = m[0].replace(/^(n[o0]|num|inv|bill|ref|doc|am|pm)\s(?=\d)/i, '').replace(/\s?([-_/:.#])\s?/g, '$1').replace(/^[-_/:.#]+|[-_/:.#]+$/g, '');
      if (!/\d/.test(t) || t.length < 2 || t.length > 32) continue;
      if (/^(n[o0]|number|num|inv|invoice|bill|ref|doc)[.:#]?$/i.test(t)) continue;
      t = t.replace(/^(?:n[o0]|#)[.:#]+/i, ''); // «No.123» → 123
      if (!/\d/.test(t)) continue;
      out.push({ t, i: m.index });
    }
    return out.filter(({ t }) => !badInv(t));
  }
  function badInv(t) {
    const d = t.replace(/\D/g, '');
    if (/^\d{1,4}[-/.]\d{1,2}[-/.]\d{1,4}$/.test(t) && findDate(t)) return true; // تاريخ
    if (/^\d{1,2}:\d{2}(:\d{2})?$/.test(t)) return true; // وقت
    if (/^(3\d{13}3|100\d{12})$/.test(t)) return true; // رقم ضريبي
    if (/^(\+?966|0)?5\d{8}$/.test(t)) return true; // جوال سعودي
    if (/^\d{1,3}(,\d{3})*\.\d{2}$/.test(t) || /^\d+\.\d{1,2}$/.test(t)) return true; // مبلغ
    if (/^\d{1,2}%?$/.test(t)) return true; // رقم قصير جدًا
    if (/^(sar|sr|vat|cr|kg|pcs|qty)\d*$/i.test(t)) return true;
    return d.length === 0;
  }
  // القيمة بعد التسمية أولًا (الأقرب)، ثم قبلها (ترتيب عربي معكوس)، ثم السطر التالي
  function invNoNear(line, next, prev) {
    const m = line.match(K.invNo); if (!m) return '';
    const after = line.slice(m.index + m[0].length), before = line.slice(0, m.index);
    // لا نتجاوز تسمية حقل آخر في نفس السطر (مثل: Invoice NO ... Date ...)
    const stopAt = s => { const k = s.search(/(تاريخ|date|vat|الرقم\s*الضريبي|العميل|customer|client|mobile|phone|هاتف|جوال|c\.?\s*r\b|السجل)/i); return k > 0 ? s.slice(0, k) : (k === 0 ? '' : s); };
    const a = invCands(stopAt(after));
    if (a.length) return a[0].t;
    const b = invCands(before.split(/(تاريخ|date|vat|الرقم\s*الضريبي|customer|client|العميل)/i).pop());
    if (b.length) return b[b.length - 1].t;
    const n = invCands(stopAt(next)); if (n.length && n[0].t.length >= 3) return n[0].t;
    return '';
  }
  // بدون تسمية: رمز مثل INV-1001 أو SIF-DM:26111020 في أعلى الفاتورة
  function invNoGuess(lines) {
    const top = lines.slice(0, Math.max(12, Math.ceil(lines.length * 0.4)));
    for (const l of top) {
      if (/(vat|tax\s*(no|number|id)|الرقم\s*الضريبي|c\.?\s*r\b|السجل|phone|mobile|tel|هاتف|جوال|iban|po\s*box|ص\.ب)/i.test(l)) continue;
      const c = invCands(l).find(({ t }) => /^(inv|sif|si|tin|ti|bill|rec|rcpt|fp|fps|ord|so|cs|sa)[A-Za-z]{0,4}[-_/:.#]?[A-Za-z]{0,4}[-_/:.#]?\d{3,}/i.test(t));
      if (c) return c.t;
    }
    return '';
  }
  function findDate(s) {
    s = normDigits(s);
    let m = s.match(/\b(20\d{2}|19\d{2})[-/.](\d{1,2})[-/.](\d{1,2})\b/);
    if (m) return iso(m[1], m[2], m[3]);
    m = s.match(/\b(\d{1,2})[-/.](\d{1,2})[-/.](20\d{2}|19\d{2}|14\d{2})\b/);
    if (m) { if (+m[3] >= 1400 && +m[3] < 1500) return m[0] + ' هـ'; return +m[2] > 12 ? iso(m[3], m[1], m[2]) : iso(m[3], m[2], m[1]); }
    m = s.match(/\b(14\d{2})[-/.](\d{1,2})[-/.](\d{1,2})\b/);
    if (m) return m[0] + ' هـ';
    const mon = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };
    m = s.match(/\b(\d{1,2})\s+(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\.?,?\s+(\d{4})/i);
    if (m) return iso(m[3], mon[m[2].toLowerCase()], m[1]);
    m = s.match(/\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\.?\s+(\d{1,2}),?\s+(\d{4})/i);
    if (m) return iso(m[3], mon[m[1].toLowerCase()], m[2]);
    return '';
  }
  function iso(y, mo, d) {
    mo = +mo; d = +d; if (mo < 1 || mo > 12 || d < 1 || d > 31) return '';
    return y + '-' + String(mo).padStart(2, '0') + '-' + String(d).padStart(2, '0');
  }

  // ---------- استخراج البنود ----------
  function parseItem(line) {
    const L = normDigits(line);
    if (K.skipItem.test(L) || K.header.test(L) && numsIn(L).length < 2) return null;
    const letters = (L.match(/[A-Za-z؀-ۿ]/g) || []).length;
    if (letters < 2) return null;
    const ns = numsIn(L).filter(x => !/^\d{7,}$/.test(x.raw) && !/%/.test(L.slice(x.i + x.raw.length, x.i + x.raw.length + 2)));
    if (ns.length < 2) return null;
    const vals = ns.map(x => x.n);
    let qty = null, price = null, total = null, alts = null;
    // ابحث عن ثلاثية: كمية × سعر = مجموع (بأي ترتيب، لأن قراءة الأسطر العربية قد تعكس الأعمدة)
    // مع تجربة البدائل المحتملة لكل رقم، وتفضيل الأقل تعديلًا
    let best = null;
    for (let a = 0; a < vals.length; a++) for (let b = 0; b < vals.length; b++) for (let c = 0; c < vals.length; c++) {
      if (a === b || b === c || a === c) continue;
      ns[a].alts.forEach((va, ia) => ns[b].alts.forEach((vb, ib) => ns[c].alts.forEach((vc, ic) => {
        const cost = ia + ib + ic;
        if (best && cost >= best.cost) return;
        if (va > 0 && vb > 0 && near(va * vb, vc, 0.02) && vc >= vb) best = { a, b, c, va, vb, vc, cost };
      })));
    }
    if (best) {
      const { a, b, c } = best;
      const intA = Number.isInteger(best.va), intB = Number.isInteger(best.vb);
      const qa = intA && !intB ? a : intB && !intA ? b : (best.va <= best.vb ? a : b);
      qty = qa === a ? best.va : best.vb; price = qa === a ? best.vb : best.va; total = best.vc;
      // إن وُجد رقم أكبر بعد إضافة الضريبة فهو المجموع الشامل
      const bigger = vals.filter((v, i) => i !== a && i !== b && i !== c && v > total && v <= total * 1.2);
      if (bigger.length) total = Math.max(...bigger);
    }
    if (total === null) {
      const decs = ns.filter(x => /[.,]\d{1,2}$/.test(x.raw));
      if (!decs.length && ns.length < 2) return null;
      const big = ns.reduce((x, y) => (y.n > x.n ? y : x)); total = big.n; alts = big.alts;
      const small = vals.filter(v => Number.isInteger(v) && v > 0 && v < 1000 && v !== total);
      if (small.length) qty = small[0];
    }
    // الوصف: النص بعد حذف الأرقام والرموز
    let desc = L;
    if ((L.match(/[\u0600-\u06FF]/g) || []).length >= (L.match(/[A-Za-z]/g) || []).length) desc = desc.replace(/(^|\s)[A-Za-z]{1,3}(?=\s|$)/g, ' '); // حروف لاتينية دخيلة في سطر عربي
    desc = desc.replace(NUM_RE, ' ').replace(/(ر\.?\s?س|SAR|SR|﷼|\$|%|x|×|@)/gi, ' ').replace(/[|_:;*=\[\]{}()]+/g, ' ').replace(/\s{2,}/g, ' ').trim();
    desc = desc.replace(/^[\-–.,\s]+|[\-–.,\s]+$/g, '');
    if (desc.replace(/[^A-Za-z؀-ۿ]/g, '').length < 2) return null;
    const it = { description: desc, quantity: qty, unit_price: price, total: round2(total) };
    if (alts && alts.length > 1) Object.defineProperty(it, '_alts', { value: alts, enumerable: false });
    return it;
  }

  // ---------- التحليل الكامل ----------
  function parseInvoice(text, qr) {
    const eastCount = (String(text || '').match(/[٠-٩۰-۹]/g) || []).length;
    EASTERN = eastCount >= 4;
    try { return parseInvoiceInner(text, qr); } finally { EASTERN = false; }
  }
  function parseInvoiceInner(text, qr) {
    const raw = normDigits(text || '');
    const lines = raw.split(/\r?\n/).map(l => l.replace(/\s+/g, ' ').trim()).filter(Boolean);
    const r = {
      vendor_name: '', vendor_tax_number: '', invoice_number: '', invoice_date: '', due_date: '',
      customer_name: '', currency: '', payment_method: '', items: [],
      subtotal: null, discount: null, tax_rate: null, tax_amount: null, total: null, src: {}
    };
    const ALT = {};
    const setTok = (k, t) => { if (t && (r[k] == null)) { r[k] = t.n; r.src[k] = 'text'; ALT[k] = t.alts; } };
    const set = (k, v, s) => { if (v !== '' && v != null && (r[k] === '' || r[k] == null)) { r[k] = v; r.src[k] = s; } };

    // 1) من رمز QR أولًا (أدق مصدر)
    if (qr) {
      set('vendor_name', qr.seller, 'qr'); set('vendor_tax_number', qr.vat_number, 'qr');
      const d = findDate(qr.timestamp) || (qr.timestamp || '').slice(0, 10); set('invoice_date', d, 'qr');
      set('total', qr.total, 'qr'); set('tax_amount', qr.vat, 'qr');
    }

    // 2) من النص
    const vatCands = [];
    lines.forEach((line, idx) => {
      const L = line;
      // الرقم الضريبي
      const vm = L.replace(/[\s-]/g, '').match(/3\d{13}3|100\d{12}|\d{15}/g);
      if (vm) vm.forEach(v => vatCands.push({ v, labeled: K.vat.test(L) || K.vat.test(lines[idx - 1] || '') }));
      // رقم الفاتورة
      if (!r.invoice_number && K.invNo.test(L)) {
        const tok = invNoNear(L, lines[idx + 1] || '', lines[idx - 1] || '');
        if (tok) set('invoice_number', tok, 'text');
      }
      // التواريخ
      if (K.due.test(L)) set('due_date', findDate(L) || findDate(lines[idx + 1] || ''), 'text');
      else if (!r.invoice_date && K.date.test(L)) set('invoice_date', findDate(L) || findDate(lines[idx + 1] || ''), 'text');
      // العميل
      if (!r.customer_name && K.customer.test(L) && !K.vat.test(L)) {
        let v = valueAfter(L, K.customer).replace(/^(name|الاسم)\s*[:：]?\s*/i, '');
        if (v && v.replace(/[^A-Za-z؀-ۿ]/g, '').length >= 2 && !/\d{5,}/.test(v)) set('customer_name', v.slice(0, 80), 'text');
      }
      // الدفع
      if (!r.payment_method) { const pm = L.match(K.pay); if (pm && !K.grand.test(L) && !/(طريق[ةه]\s*الدفع|payment\s*method)/i.test(L) || pm && /(طريق[ةه]\s*الدفع|payment\s*method|الدفع)/i.test(L)) set('payment_method', pm[0], 'text'); }
      // المبالغ
      const pct = L.match(/(\d{1,2}(?:\.\d+)?)\s*%/);
      if (K.grand.test(L)) setTok('total', lastTok(L, { noPercent: true }));
      else if (K.sub.test(L)) setTok('subtotal', lastTok(L, { noPercent: true }));
      else if (K.discount.test(L)) { const n = lastNum(L, { noPercent: true }); if (n !== null) set('discount', n, 'text'); }
      else if (K.tax.test(L) && !K.vat.test(L)) {
        if (pct) set('tax_rate', toNum(pct[1]), 'text');
        setTok('tax_amount', lastTok(L, { noPercent: true }));
      }
      else if (K.total.test(L) && !K.tax.test(L)) { const t = lastTok(L, { noPercent: true }); if (t) (r._plainTotals = r._plainTotals || []).push(t); }
      if (pct && K.tax.test(L)) set('tax_rate', toNum(pct[1]), 'text');
    });
    if (!r.invoice_number) { const g = invNoGuess(lines); if (g) set('invoice_number', g, 'text'); }

    // الرقم الضريبي: الأفضل ما كان بجانب تسمية، ثم ما يطابق صيغة السعودية
    if (vatCands.length) {
      const pick = vatCands.find(c => c.labeled && /^3\d{13}3$/.test(c.v)) || vatCands.find(c => /^3\d{13}3$/.test(c.v)) || vatCands.find(c => c.labeled) ;
      if (pick) set('vendor_tax_number', pick.v, 'text');
      r._vatText = vatCands.map(c => c.v);
    }
    // الإجمالي إن لم يوجد سطر "شامل الضريبة"
    if (r.total == null && r._plainTotals) setTok('total', r._plainTotals.reduce((a, b) => (b.n > a.n ? b : a)));
    if (r.subtotal == null && r._plainTotals && r._plainTotals.length > 1) {
      const t = r.total, s = r._plainTotals.filter(v => v.n < t); if (s.length) setTok('subtotal', s.reduce((a, b) => (b.n > a.n ? b : a)));
    }
    delete r._plainTotals;

    // العملة
    let bestCur = null, bestN = 0;
    for (const [re, code] of CUR) { const n = (raw.match(re) || []).length; if (n > bestN) { bestN = n; bestCur = code; } }
    if (bestCur) { r.currency = bestCur; r.src.currency = 'text'; }
    if (!r.currency && qr) { r.currency = 'SAR'; r.src.currency = 'qr'; }
    if (!r.currency && /^3\d{13}3$/.test(r.vendor_tax_number)) { r.currency = 'SAR'; r.src.currency = 'calc'; } // رقم ضريبي سعودي

    // المورّد: أول سطر نصي في الأعلى ليس تسمية أو رقمًا
    if (!r.vendor_name) {
      const top = lines.slice(0, 8).find(l => {
        const letters = (l.match(/[A-Za-z؀-ۿ]/g) || []).length;
        return letters >= 3 && !/^(فاتور[ةه]|invoice|tax invoice|simplified|receipt|إيصال|ايصال)\b/i.test(l) && !/(فاتور[ةه]\s*ضريبي[ةه]|tax\s*invoice|simplified)/i.test(l) && !K.date.test(l) && !K.invNo.test(l) && !K.vat.test(l) && (l.match(/\d/g) || []).length < 6;
      });
      if (top) set('vendor_name', top.slice(0, 80), 'text');
    }

    // البنود
    lines.forEach(l => { const it = parseItem(l); if (it) r.items.push(it); });
    if (EASTERN) repairAmounts(r, ALT);
    // استبعاد بند يساوي الإجمالي بالخطأ
    r.items = r.items.filter(it => !(r.total != null && (near(it.total, r.total) && r.items.length > 1 || it.total > r.total * 1.01)));

    // إن لم تتطابق البنود مع المجموع، جرّب حذف الأسطر غير المتسقة (لا كمية ولا سعر)
    if (!itemsOk(r)) {
      const consistent = r.items.filter(it => it.quantity != null && it.unit_price != null);
      if (consistent.length && consistent.length < r.items.length && itemsOk({ ...r, items: consistent })) r.items = consistent;
    }
    // الضريبة: استنتاج النسبة
    if (r.tax_rate == null && r.tax_amount != null) {
      const base = r.subtotal != null ? r.subtotal : (r.total != null ? r.total - r.tax_amount : null);
      if (base > 0) { const rate = round2(r.tax_amount / base * 100); if ([5, 10, 14, 15].some(x => Math.abs(x - rate) < 0.3)) { r.tax_rate = Math.round(rate); r.src.tax_rate = 'calc'; } }
    }
    if (r.subtotal == null && r.total != null && r.tax_amount != null) { r.subtotal = round2(r.total - r.tax_amount + (r.discount || 0)); r.src.subtotal = 'calc'; }
    return r;
  }

  // ---------- إصلاح المبالغ بالحساب ----------
  // يختار من بدائل كل مبلغ ما يجعل: المجموع + الضريبة = الإجمالي، والضريبة = النسبة × المجموع، ومجموع البنود = المجموع.
  function repairAmounts(r, ALT) {
    const cand = k => r.src[k] === 'qr' ? [r[k]] : (ALT[k] || (r[k] != null ? [r[k]] : [null]));
    const rateHint = r.tax_rate != null ? [r.tax_rate] : [15, 5];
    const S0 = cand('subtotal'), V0 = cand('tax_amount'), T0 = cand('total');
    let best = null;
    const consider = (S, V, T, cost, its, ic) => {
      let sc = -cost;
      if (S != null && V != null && T != null && near(S + V, T, 0.1)) sc += 6;
      if (S != null && V != null && rateHint.some(q => near(S * q / 100, V, 0.1))) sc += 4;
      if (S == null && V != null && T != null && rateHint.some(q => near((T - V) * q / 100, V, 0.1))) sc += 4;
      if (its != null && (S != null && near(its, S, 0.1) || T != null && !V && near(its, T, 0.1))) sc += 5;
      if (!best || sc > best.sc) best = { S, V, T, sc, ic };
    };
    // بدائل بنود بلا كمية/سعر: نجرّب الأصل فقط، أو بدائلها إن كانت قليلة
    const flex = r.items.filter(it => it._alts);
    const itemCombos = [];
    if (flex.length && flex.length <= 6) {
      const rec = (i, acc, cost, picks) => {
        if (i === flex.length) { itemCombos.push({ sum: acc, cost, picks }); return; }
        flex[i]._alts.slice(0, 4).forEach((v, j) => rec(i + 1, acc + v, cost + j, picks.concat(v)));
      };
      const fixed = r.items.filter(it => !it._alts).reduce((a, it) => a + (+it.total || 0), 0);
      rec(0, fixed, 0, []);
    } else itemCombos.push({ sum: r.items.length ? r.items.reduce((a, it) => a + (+it.total || 0), 0) : null, cost: 0, picks: null });
    S0.forEach((S, i) => V0.forEach((V, j) => T0.forEach((T, k) => itemCombos.forEach(ic => {
      consider(S, V, T, i + j + k + ic.cost * 0.5, ic.sum, ic);
    }))));
    if (!best) return;
    [['subtotal', best.S], ['tax_amount', best.V], ['total', best.T]].forEach(([k, v]) => { if (v != null && r[k] !== v) { r[k] = round2(v); } });
    if (best.ic && best.ic.picks) flex.forEach((it, i) => { it.total = round2(best.ic.picks[i]); });
  }

  // ---------- التنبيهات ----------
  function checks(d, qr, extra) {
    const w = [];
    const vatMsg = checkVat(d.vendor_tax_number); if (vatMsg) w.push(vatMsg);
    if (!d.vendor_tax_number) w.push('لم يُعثر على الرقم الضريبي للمورّد.');
    if (d.total == null || d.total === '') w.push('لم يُعثر على الإجمالي. أدخله يدويًا.');
    const items = (d.items || []).filter(it => it.total != null && it.total !== '' && isFinite(+it.total));
    const sumItems = round2(items.reduce((a, it) => a + +it.total, 0));
    const T = +d.total, S = d.subtotal != null && d.subtotal !== '' ? +d.subtotal : null, V = d.tax_amount != null && d.tax_amount !== '' ? +d.tax_amount : null, D = +(d.discount || 0);
    if (items.length && isFinite(T) && d.total !== '' && d.total != null) {
      const okSub = S != null && (near(sumItems, S, 0.1) || near(sumItems - D, S, 0.1));
      const okTot = near(sumItems, T, 0.1) || near(sumItems - D, T, 0.1) || (V != null && (near(sumItems + V, T, 0.1) || near(sumItems - D + V, T, 0.1)));
      if (!okSub && !okTot) w.push(`مجموع البنود (${fmt(sumItems)}) لا يطابق ${S != null ? 'المجموع قبل الضريبة (' + fmt(S) + ') ولا ' : ''}الإجمالي (${fmt(T)}). راجع البنود أو المبالغ.`);
    }
    if (S != null && V != null && isFinite(T) && d.total != null && d.total !== '' && !near(S - D + V, T, 0.1) && !near(S + V, T, 0.1))
      w.push(`المجموع قبل الضريبة + الضريبة = ${fmt(round2(S + V - D))}، والإجمالي المكتوب ${fmt(T)}.`);
    if (d.tax_rate != null && d.tax_rate !== '' && V != null && S != null && S > 0 && !near(S * d.tax_rate / 100, V, 0.1) && !near((S - D) * d.tax_rate / 100, V, 0.1))
      w.push(`الضريبة ${fmt(V)} لا تساوي ${d.tax_rate}% من ${fmt(S)} (المتوقع ${fmt(round2(S * d.tax_rate / 100))}).`);
    if (d.tax_rate != null && d.tax_rate !== '' && +d.tax_rate !== 15 && (d.currency === 'SAR' || /^3\d{13}3$/.test(d.vendor_tax_number || '')))
      w.push(`نسبة الضريبة المقروءة ${d.tax_rate}%، والنسبة المعتادة في السعودية 15%. راجع المبالغ.`);
    if (qr && extra && extra.text) {
      if (extra.text.total != null && qr.total != null && !near(extra.text.total, qr.total, 0.1)) w.push(`الإجمالي في رمز QR (${fmt(qr.total)}) يختلف عن المكتوب في الفاتورة (${fmt(extra.text.total)}).`);
      if (extra.text._vatText && extra.text._vatText.length && !extra.text._vatText.includes(qr.vat_number)) w.push('الرقم الضريبي في رمز QR يختلف عن المكتوب في الفاتورة.');
    }
    return w;
  }
  const fmt = n => (n == null || n === '' || !isFinite(+n)) ? '' : (+n).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

  // درجة جودة القراءة: لاختيار أفضل محاولة عند تكرار القراءة بإعدادات مختلفة
  function score(r) {
    let sc = 0; const it = r.items || [];
    sc += it.filter(x => x.quantity != null && x.unit_price != null).length * 2 + it.length;
    const sum = it.reduce((a, x) => a + (+x.total || 0), 0);
    if (it.length && (r.subtotal != null && near(sum, r.subtotal, 0.1) || r.total != null && near(sum, r.total, 0.1))) sc += 10;
    ['invoice_number', 'invoice_date', 'vendor_tax_number'].forEach(k => { if (r[k]) sc += 2; });
    ['total', 'subtotal', 'tax_amount'].forEach(k => { if (r[k] != null) sc += 2; });
    return sc;
  }
  function itemsOk(r) {
    const it = r.items || []; if (!it.length) return false;
    const sum = it.reduce((a, x) => a + (+x.total || 0), 0);
    return (r.subtotal != null && near(sum, r.subtotal, 0.1)) || (r.total != null && near(sum, r.total, 0.1)) || (r.total != null && r.tax_amount != null && near(sum + r.tax_amount, r.total, 0.1));
  }

  const api = { score, itemsOk, invCands, normDigits, toNum, numsIn, parseZatcaQR, checkVat, parseInvoice, parseItem, checks, findDate, fmt };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.InvoiceParse = api;
})(typeof self !== 'undefined' ? self : this);
