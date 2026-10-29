// رفيق — a small .xlsx / .csv reader (no dependencies, nothing leaves the device).
// Returns the first sheet with data as rows of strings; numbers stay numbers so dates can be recognised by their column.
const dec = new TextDecoder();

async function inflate(bytes) {
  const ds = new DecompressionStream('deflate-raw');
  const out = new Response(new Blob([bytes]).stream().pipeThrough(ds));
  return new Uint8Array(await out.arrayBuffer());
}
/** ZIP central directory → { name: async () => Uint8Array } */
function unzip(buf) {
  const v = new DataView(buf.buffer, buf.byteOffset, buf.byteLength);
  let eocd = -1;
  for (let i = buf.length - 22; i >= Math.max(0, buf.length - 66000); i--) if (v.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
  if (eocd < 0) throw new Error('not_zip');
  const n = v.getUint16(eocd + 10, true); let p = v.getUint32(eocd + 16, true);
  const files = {};
  for (let k = 0; k < n; k++) {
    if (v.getUint32(p, true) !== 0x02014b50) throw new Error('bad_zip');
    const method = v.getUint16(p + 10, true), csize = v.getUint32(p + 20, true), nlen = v.getUint16(p + 28, true), xlen = v.getUint16(p + 30, true), clen = v.getUint16(p + 32, true), off = v.getUint32(p + 42, true);
    const name = dec.decode(buf.subarray(p + 46, p + 46 + nlen));
    files[name] = async () => {
      const ln = v.getUint16(off + 26, true), lx = v.getUint16(off + 28, true);
      const data = buf.subarray(off + 30 + ln + lx, off + 30 + ln + lx + csize);
      if (method === 0) return data;
      if (method === 8) return inflate(data);
      throw new Error('zip_method');
    };
    p += 46 + nlen + xlen + clen;
  }
  return files;
}
const xmlOf = async (files, name) => files[name] ? new DOMParser().parseFromString(dec.decode(await files[name]()), 'application/xml') : null;
const all = (doc, tag) => doc ? [...doc.getElementsByTagNameNS('*', tag)] : [];
const textOf = el => all(el, 't').map(t => t.textContent).join('');
const colIdx = ref => { const m = /^([A-Z]+)/.exec(ref || ''); if (!m) return -1; let n = 0; for (const ch of m[1]) n = n * 26 + ch.charCodeAt(0) - 64; return n - 1; };

/** Rows of one worksheet document. */
function rowsOf(doc, shared) {
  const out = [];
  all(doc, 'row').forEach(r => {
    const ri = (parseInt(r.getAttribute('r'), 10) || out.length + 1) - 1, row = [];
    all(r, 'c').forEach((c, j) => {
      const ci = c.getAttribute('r') ? colIdx(c.getAttribute('r')) : j, ty = c.getAttribute('t') || '';
      const vEl = all(c, 'v')[0]; const raw = vEl ? vEl.textContent : '';
      let val = '';
      if (ty === 's') val = shared[parseInt(raw, 10)] ?? '';
      else if (ty === 'inlineStr') val = textOf(c);
      else if (ty === 'b') val = raw === '1' ? 'TRUE' : 'FALSE';
      else if (ty === 'str' || ty === 'e' || ty === 'd') val = raw;
      else val = raw === '' ? '' : (isFinite(+raw) ? +raw : raw);
      row[ci] = val;
    });
    out[ri] = row;
  });
  return out.filter(Boolean).map(r => Array.from(r, x => x ?? '')).filter(r => r.some(x => String(x).trim() !== ''));
}

/** CSV with quotes; the separator is guessed (comma, semicolon or tab). */
function parseCSV(text) {
  text = text.replace(/^﻿/, '');
  const first = text.split(/\r?\n/)[0] || '';
  const sep = [',', ';', '\t'].map(s => [s, first.split(s).length]).sort((a, b) => b[1] - a[1])[0][0];
  const rows = []; let row = [], cur = '', q = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (q) { if (ch === '"') { if (text[i + 1] === '"') { cur += '"'; i++; } else q = false; } else cur += ch; continue; }
    if (ch === '"') q = true; else if (ch === sep) { row.push(cur); cur = ''; } else if (ch === '\n' || ch === '\r') { if (ch === '\r' && text[i + 1] === '\n') i++; row.push(cur); rows.push(row); row = []; cur = ''; } else cur += ch;
  }
  if (cur || row.length) { row.push(cur); rows.push(row); }
  return rows.filter(r => r.some(x => String(x).trim() !== ''));
}

/** → { sheet, rows } ; throws Error('unsupported' | 'empty' | 'not_zip' …) */
export async function readSheet(file) {
  const name = (file.name || '').toLowerCase();
  if (/\.csv$/.test(name) || file.type === 'text/csv') { const rows = parseCSV(await file.text()); if (!rows.length) throw new Error('empty'); return { sheet: file.name, rows }; }
  if (!/\.xlsx$|\.xlsm$/.test(name) && !/spreadsheetml/.test(file.type || '')) throw new Error('unsupported');
  if (typeof DecompressionStream === 'undefined') throw new Error('old_browser');
  const files = unzip(new Uint8Array(await file.arrayBuffer()));
  const sst = await xmlOf(files, 'xl/sharedStrings.xml');
  const shared = all(sst, 'si').map(textOf);
  const wb = await xmlOf(files, 'xl/workbook.xml'), rels = await xmlOf(files, 'xl/_rels/workbook.xml.rels');
  const relMap = Object.fromEntries(all(rels, 'Relationship').map(r => [r.getAttribute('Id'), r.getAttribute('Target')]));
  const sheets = all(wb, 'sheet').map(s => ({ name: s.getAttribute('name'), path: (t => t ? (t.startsWith('/') ? t.slice(1) : 'xl/' + t.replace(/^\.\//, '')) : '')(relMap[s.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id') || s.getAttribute('r:id')]) }));
  const list = sheets.length ? sheets : Object.keys(files).filter(k => /^xl\/worksheets\/sheet\d+\.xml$/.test(k)).sort().map((p, i) => ({ name: 'Sheet' + (i + 1), path: p }));
  for (const s of list) {
    const doc = await xmlOf(files, s.path); const rows = rowsOf(doc, shared);
    if (rows.length) return { sheet: s.name, rows };
  }
  throw new Error('empty');
}

/** Excel serial day (1900 system) → YYYY-MM-DD */
export function serialDate(n) {
  if (typeof n !== 'number' || !isFinite(n) || n < 20000 || n > 80000) return null;
  const d = new Date(Date.UTC(1899, 11, 30) + Math.round(n) * 86400000);
  return d.toISOString().slice(0, 10);
}
