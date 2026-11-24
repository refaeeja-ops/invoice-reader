/* قارئ الفواتير: تصدير Excel (xlsx) و CSV بدون مكتبات خارجية. */
(function (root) {
  'use strict';

  // ---------- ZIP بدون ضغط (يكفي لملف xlsx) ----------
  const CRC = (() => { const t = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
  function crc32(b) { let c = 0xFFFFFFFF; for (let i = 0; i < b.length; i++) c = CRC[(c ^ b[i]) & 0xFF] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; }
  const enc = s => new TextEncoder().encode(s);
  function zip(files) { // files: [{name, data:string}]
    const parts = [], central = []; let off = 0;
    const u16 = v => [v & 255, (v >>> 8) & 255], u32 = v => [v & 255, (v >>> 8) & 255, (v >>> 16) & 255, (v >>> 24) & 255];
    files.forEach(f => {
      const name = enc(f.name), data = enc(f.data), crc = crc32(data);
      const head = [...u16(20), ...u16(0x0800), ...u16(0), ...u16(0), ...u16(0x21), ...u32(crc), ...u32(data.length), ...u32(data.length), ...u16(name.length), ...u16(0)];
      const local = new Uint8Array([...u32(0x04034b50), ...head]);
      parts.push(local, name, data);
      central.push(new Uint8Array([...u32(0x02014b50), ...u16(20), ...head, ...u16(0), ...u16(0), ...u16(0), ...u32(0), ...u32(off)]), name);
      off += local.length + name.length + data.length;
    });
    const cSize = central.reduce((a, b) => a + b.length, 0);
    const end = new Uint8Array([...u32(0x06054b50), ...u16(0), ...u16(0), ...u16(files.length), ...u16(files.length), ...u32(cSize), ...u32(off), ...u16(0)]);
    return new Blob([...parts, ...central, end], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  }

  // ---------- XLSX ----------
  const xesc = s => String(s ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])).replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, '');
  function colName(i) { let s = ''; i++; while (i) { const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s; }
  function sheetXml(rows, widths) {
    const body = rows.map((row, r) => `<row r="${r + 1}">` + row.map((v, c) => {
      const ref = colName(c) + (r + 1), st = r === 0 ? ' s="1"' : '';
      if (typeof v === 'number' && isFinite(v)) return `<c r="${ref}"${r === 0 ? st : ' s="2"'}><v>${v}</v></c>`;
      if (v == null || v === '') return '';
      return `<c r="${ref}" t="inlineStr"${st}><is><t xml:space="preserve">${xesc(v)}</t></is></c>`;
    }).join('') + '</row>').join('');
    const cols = '<cols>' + widths.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join('') + '</cols>';
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView rightToLeft="1" workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>${cols}<sheetData>${body}</sheetData><autoFilter ref="A1:${colName(Math.max(0, rows[0].length - 1))}${Math.max(1, rows.length)}"/></worksheet>`;
  }
  function xlsx(sheets) { // sheets: [{name, rows, widths}]
    const files = [
      { name: '[Content_Types].xml', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${sheets.map((s, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')}</Types>` },
      { name: '_rels/.rels', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>` },
      { name: 'xl/workbook.xml', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView/></bookViews><sheets>${sheets.map((s, i) => `<sheet name="${xesc(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('')}</sheets></workbook>` },
      { name: 'xl/_rels/workbook.xml.rels', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${sheets.map((s, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('')}<Relationship Id="rId${sheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>` },
      { name: 'xl/styles.xml', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Arial"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Arial"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF2340C8"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/><xf numFmtId="4" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>` }
    ];
    sheets.forEach((s, i) => files.push({ name: `xl/worksheets/sheet${i + 1}.xml`, data: sheetXml(s.rows, s.widths) }));
    return zip(files);
  }

  // ---------- تحويل الفواتير إلى صفوف ----------
  const num = v => (v === '' || v == null || !isFinite(+v)) ? (v == null ? '' : v) : +v;
  const INV_COLS = [
    ['file', 'الملف', 22], ['vendor_name', 'المورّد', 28], ['vendor_tax_number', 'الرقم الضريبي', 18], ['invoice_number', 'رقم الفاتورة', 16],
    ['invoice_date', 'التاريخ', 12], ['due_date', 'الاستحقاق', 12], ['customer_name', 'العميل', 22], ['currency', 'العملة', 8],
    ['subtotal', 'قبل الضريبة', 13, 1], ['discount', 'الخصم', 10, 1], ['tax_rate', 'نسبة الضريبة %', 10, 1], ['tax_amount', 'الضريبة', 12, 1], ['total', 'الإجمالي', 13, 1],
    ['payment_method', 'طريقة الدفع', 12], ['items_count', 'عدد البنود', 9, 1], ['qr', 'رمز QR', 8], ['status', 'الحالة', 16], ['approved_at', 'تاريخ الاعتماد', 12], ['warnings', 'تنبيهات', 40], ['note', 'ملاحظات', 30]];
  const ITEM_COLS = [['invoice_number', 'رقم الفاتورة', 16], ['vendor_name', 'المورّد', 26], ['invoice_date', 'التاريخ', 12], ['n', '#', 5, 1],
    ['description', 'الوصف', 34], ['quantity', 'الكمية', 9, 1], ['unit_price', 'سعر الوحدة', 12, 1], ['total', 'المجموع', 12, 1], ['currency', 'العملة', 8]];

  function invRows(list) {
    return [INV_COLS.map(c => c[1]), ...list.map(d => INV_COLS.map(([k, , , isNum]) => {
      if (k === 'items_count') return (d.items || []).length;
      if (k === 'qr') return d.src && d.src.vendor_tax_number === 'qr' || d._qr ? 'نعم' : 'لا';
      if (k === 'warnings') return [d._dup, ...(d._warnings || [])].filter(Boolean).join(' | ').replace(/[\u2068\u2069]/g, '');
      if (k === 'status') return d._approved ? 'معتمدة' : 'بانتظار الاعتماد';
      if (k === 'approved_at') return d._approved && d._approvedAt ? String(d._approvedAt).slice(0, 10) : '';
      return isNum ? num(d[k]) : (d[k] ?? '');
    }))];
  }
  function itemRows(list) {
    const rows = [ITEM_COLS.map(c => c[1])];
    list.forEach(d => (d.items || []).forEach((it, i) => rows.push(ITEM_COLS.map(([k, , , isNum]) => {
      if (k === 'n') return i + 1;
      const v = ['description', 'quantity', 'unit_price', 'total'].includes(k) ? it[k] : d[k];
      return isNum ? num(v) : (v ?? '');
    }))));
    return rows;
  }
  function toXlsx(list) {
    return xlsx([{ name: 'الفواتير', rows: invRows(list), widths: INV_COLS.map(c => c[2]) }, { name: 'البنود', rows: itemRows(list), widths: ITEM_COLS.map(c => c[2]) }]);
  }
  function csv(rows) {
    const q = v => { const s = String(v ?? ''); return /[",\n\r]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s; };
    return '﻿' + rows.map(r => r.map(q).join(',')).join('\r\n');
  }
  const api = { toXlsx, invoicesCsv: list => csv(invRows(list)), itemsCsv: list => csv(itemRows(list)), xlsx, zip, crc32 };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.InvoiceExport = api;
})(typeof self !== 'undefined' ? self : this);
