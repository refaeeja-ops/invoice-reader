/* قارئ الأرقام الهندية (٠١٢٣٤٥٦٧٨٩): مصنّف صغير يعمل داخل المتصفح.
   يقطّع كل كلمة رقمية إلى مكوّنات متصلة، ثم يصنّف كل مكوّن بشبكة عصبية صغيرة مدرّبة على خطوط عربية مجانية.
   مطابق لخوارزمية التدريب في tools/seg.py. */
(function (root) {
  'use strict';
  let M = null, loading = null;
  async function load(url) {
    if (M) return M;
    if (!loading) loading = fetch(url).then(r => { if (!r.ok) throw new Error('digits model'); return r.json(); }).then(j => {
      const f = a => Float32Array.from(a.flat ? a.flat() : a);
      M = { cls: j.cls, N: j.N, mu: f(j.mu), sd: f(j.sd), W1: f(j.W1), b1: f(j.b1), W2: f(j.W2), b2: f(j.b2), H: j.b1.length, C: j.cls.length, D: j.mu.length };
      return M;
    }).catch(e => { loading = null; throw e; });
    return loading;
  }

  function otsu(g) {
    const h = new Float64Array(256); for (let i = 0; i < g.length; i++) h[g[i]]++;
    let sumAll = 0; for (let t = 0; t < 256; t++) sumAll += t * h[t];
    let sb = 0, wb = 0, best = 0, thr = 127; const tot = g.length;
    for (let t = 0; t < 256; t++) {
      wb += h[t]; if (!wb) continue; const wf = tot - wb; if (!wf) break;
      sb += t * h[t]; const mb = sb / wb, mf = (sumAll - sb) / wf, v = wb * wf * (mb - mf) ** 2;
      if (v > best) { best = v; thr = t; }
    }
    return thr;
  }
  // تسمية المكوّنات المتصلة (جوار ثماني)
  function label(ink, W, H) {
    const lab = new Int32Array(W * H), comps = []; let n = 0; const st = [];
    for (let p = 0; p < W * H; p++) {
      if (!ink[p] || lab[p]) continue;
      n++; lab[p] = n; st.push(p); let y0 = 1e9, y1 = -1, x0 = 1e9, x1 = -1, area = 0;
      while (st.length) {
        const q = st.pop(), y = (q / W) | 0, x = q - y * W; area++;
        if (y < y0) y0 = y; if (y > y1) y1 = y; if (x < x0) x0 = x; if (x > x1) x1 = x;
        for (let dy = -1; dy <= 1; dy++) for (let dx = -1; dx <= 1; dx++) {
          const yy = y + dy, xx = x + dx; if (yy < 0 || yy >= H || xx < 0 || xx >= W) continue;
          const r = yy * W + xx; if (ink[r] && !lab[r]) { lab[r] = n; st.push(r); }
        }
      }
      comps.push({ ids: [n], y0, y1: y1 + 1, x0, x1: x1 + 1, area });
    }
    return { lab, comps };
  }
  const median = a => { const s = a.slice().sort((x, y) => x - y), m = s.length >> 1; return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2; };

  function segment(gray, W, H, N) {
    const thr = otsu(gray); let ink = new Uint8Array(W * H), cnt = 0;
    for (let i = 0; i < gray.length; i++) { ink[i] = gray[i] <= thr ? 1 : 0; cnt += ink[i]; }
    if (cnt > 0.5 * gray.length) for (let i = 0; i < ink.length; i++) ink[i] = 1 - ink[i];
    const { lab, comps: all } = label(ink, W, H);
    const minA = Math.max(3, 0.006 * H * H);
    const comps = all.filter(c => c.area >= minA).sort((a, b) => a.x0 - b.x0);
    if (!comps.length) return [];
    const merged = [];
    for (const c of comps) {
      const m = merged[merged.length - 1];
      if (m) {
        const ov = Math.min(m.x1, c.x1) - Math.max(m.x0, c.x0);
        const oy = Math.min(m.y1, c.y1) - Math.max(m.y0, c.y0);
        if (ov >= 0.5 * Math.min(m.x1 - m.x0, c.x1 - c.x0) && oy <= 0.2 * Math.min(m.y1 - m.y0, c.y1 - c.y0)) {
          m.x0 = Math.min(m.x0, c.x0); m.x1 = Math.max(m.x1, c.x1); m.y0 = Math.min(m.y0, c.y0); m.y1 = Math.max(m.y1, c.y1); m.area += c.area; m.ids = m.ids.concat(c.ids); continue;
        }
      }
      merged.push({ ...c });
    }
    const mx = Math.max(...merged.map(m => m.y1 - m.y0));
    const tall = merged.filter(m => m.y1 - m.y0 >= 0.5 * mx);
    const R = median(tall.map(m => m.y1 - m.y0)), mid = median(tall.map(m => (m.y0 + m.y1) / 2));
    for (const m of merged) {
      const w = m.x1 - m.x0, h = m.y1 - m.y0, S = Math.max(w, h), oy = (S - h) >> 1, ox = (S - w) >> 1;
      const ids = new Set(m.ids);
      const cs = new Float32Array((S + 1) * (S + 1)); // مجموع تراكمي لصورة المكوّن داخل مربع
      for (let y = 0; y < S; y++) {
        let row = 0;
        for (let x = 0; x < S; x++) {
          const sy = y - oy, sx = x - ox; let v = 0;
          if (sy >= 0 && sy < h && sx >= 0 && sx < w) v = ids.has(lab[(m.y0 + sy) * W + m.x0 + sx]) ? 1 : 0;
          row += v; cs[(y + 1) * (S + 1) + x + 1] = cs[y * (S + 1) + x + 1] + row;
        }
      }
      const f = new Float32Array(N * N + 4);
      for (let i = 0; i < N; i++) {
        const y0 = Math.floor(i * S / N), y1 = Math.max(y0 + 1, Math.floor((i + 1) * S / N));
        for (let j = 0; j < N; j++) {
          const x0 = Math.floor(j * S / N), x1 = Math.max(x0 + 1, Math.floor((j + 1) * S / N));
          f[i * N + j] = (cs[y1 * (S + 1) + x1] - cs[y0 * (S + 1) + x1] - cs[y1 * (S + 1) + x0] + cs[y0 * (S + 1) + x0]) / ((y1 - y0) * (x1 - x0));
        }
      }
      const clip = v => Math.max(-3, Math.min(3, v));
      f[N * N] = clip(h / R); f[N * N + 1] = clip(w / R); f[N * N + 2] = clip(((m.y0 + m.y1) / 2 - mid) / R); f[N * N + 3] = m.area / (w * h);
      m.feat = f;
    }
    return merged;
  }
  function classify(f) {
    const { mu, sd, W1, b1, W2, b2, H, C, D } = M;
    const x = new Float32Array(D); for (let i = 0; i < D; i++) x[i] = (f[i] - mu[i]) / sd[i];
    const h = new Float32Array(H);
    for (let j = 0; j < H; j++) { let s = b1[j]; for (let i = 0; i < D; i++) s += x[i] * W1[i * H + j]; h[j] = s > 0 ? s : 0; }
    const o = new Float32Array(C); let mx = -1e9;
    for (let k = 0; k < C; k++) { let s = b2[k]; for (let j = 0; j < H; j++) s += h[j] * W2[j * C + k]; o[k] = s; if (s > mx) mx = s; }
    let z = 0; for (let k = 0; k < C; k++) { o[k] = Math.exp(o[k] - mx); z += o[k]; }
    let best = 0; for (let k = 1; k < C; k++) if (o[k] > o[best]) best = k;
    return { c: M.cls[best], p: o[best] / z };
  }
  // يقرأ كلمة رقمية من صورة رمادية (Uint8 لكل بكسل). يعيد نصًّا بأرقام غربية أو null إن لم تكن رقمية.
  function readWord(gray, W, H) {
    if (!M) return null;
    const parts = segment(gray, W, H, M.N);
    if (!parts.length || parts.length > 24) return null;
    let out = '', digits = 0, eastern = 0, other = 0, low = 0;
    const res = parts.map(p => classify(p.feat));
    // رموز غير رقمية في طرفي الكلمة (مثل × أو حرف ملتصق) تُفصل بمسافة بدل رفض الكلمة
    let a = 0, z = parts.length;
    while (a < z && res[a].c === 'x') a++;
    while (z > a && res[z - 1].c === 'x') z--;
    for (let i = a; i < z; i++) {
      const p = parts[i], r = res[i], N2 = M.N * M.N;
      // علامة صغيرة بين رقمين تُعدّ فاصلة عشرية (٫)
      if ((r.c === 'x' || r.p < 0.5) && i > a && i < z - 1 && (p.feat[N2] < 0.6 || p.feat[N2 + 2] > 0.3) && p.feat[N2] < 1) { r.c = '.'; r.p = 1; }
      if (r.p < 0.5) low++;
      if (/^\d$/.test(r.c)) { out += r.c; digits++; }
      else if (/^e\d$/.test(r.c)) { out += r.c[1]; digits++; eastern++; }
      else if (r.c === 'x') other++;
      else out += r.c;
    }
    if (!digits || other > 0 || low > (z - a) / 3) return null;
    if (!/^\d+(\.\d+)?$|^\d{1,4}\/\d{1,2}\/\d{1,4}$|^\d{1,2}:\d{2}(:\d{2})?$|^\d+-\d+(-\d+)*$/.test(out)) return null;
    return { text: out, eastern, digits };
  }
  // يستبدل الكلمات الرقمية في نتيجة Tesseract بقراءة المصنّف، ويعيد بناء النص سطرًا سطرًا.
  function rebuild(data, canvas) {
    if (!M || !data || !data.blocks) return null;
    const ctx = canvas.getContext('2d', { willReadFrequently: true });
    const lines = []; let changed = 0;
    for (const b of data.blocks) for (const par of b.paragraphs) {
      for (const ln of par.lines) {
        const words = ln.words.map(w => {
          const t = w.text;
          // نجرّب كل كلمة ليست كلمة عربية واضحة (قارئ النص قد يقرأ الأرقام الهندية حروفًا)
          if ((t.match(/[\u0621-\u064A]/g) || []).length >= 3 && !/[0-9٠-٩۰-۹]/.test(t)) return t;
          const bb = w.bbox, h = bb.y1 - bb.y0, pad = Math.max(2, Math.round(0.15 * h));
          const x0 = Math.max(0, bb.x0 - pad), y0 = Math.max(0, bb.y0 - pad);
          const x1 = Math.min(canvas.width, bb.x1 + pad), y1 = Math.min(canvas.height, bb.y1 + pad);
          const W = x1 - x0, H = y1 - y0; if (W < 4 || H < 6) return t;
          const id = ctx.getImageData(x0, y0, W, H).data, g = new Uint8Array(W * H);
          for (let i = 0; i < g.length; i++) g[i] = id[i * 4];
          const r = readWord(g, W, H);
          if (root.__digitsDebug) root.__digitsDebug.push({ t, bbox: bb, r: r && r.text });
          if (!r || !r.eastern) return t; // نستبدل فقط الكلمات ذات الأرقام الهندية
          if (r.text === '0' && !/[0-9٠-٩۰-۹]/.test(t)) return t; // نقطة منفردة ليست صفرًا غالبًا
          changed++;
          return r.text;
        });
        lines.push(words.join(' '));
      }
      lines.push('');
    }
    return changed ? { text: lines.join('\n'), changed } : null;
  }
  root.EasternDigits = { load, readWord, rebuild, _segment: segment, _classify: f => classify(f), get ready() { return !!M; } };
})(typeof self !== 'undefined' ? self : this);
