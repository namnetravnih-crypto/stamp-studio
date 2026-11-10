/* =====================================================================
   Ассистент — чтение старого формата Word (.doc, Word 97–2003) прямо в браузере.
   • text(buf)        — текст документа: абзацы и таблицы («ячейка | ячейка»)
   • blocks(buf)      — структура: абзацы и таблицы
   • toDocx(buf)      — упрощённый .docx (текст и таблицы без исходного оформления)
   • images(buf)      — картинки (PNG / JPEG), встроенные в .doc
   Нужны библиотеки cfb.min.js (контейнер) и jszip (только для toDocx).
   ===================================================================== */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(require('cfb'), require('jszip'));
  else root.Doc97 = factory(root.CFB, root.JSZip);
})(typeof self !== 'undefined' ? self : this, function (CFB, JSZip) {
  'use strict';
  const u16 = (b, o) => b[o] | (b[o + 1] << 8);
  const u32 = (b, o) => (b[o] | (b[o + 1] << 8) | (b[o + 2] << 16) | (b[o + 3] << 24)) >>> 0;
  const toU8 = c => (c instanceof Uint8Array ? c : Uint8Array.from(c));

  function openCf(buf) {
    const u8 = buf instanceof Uint8Array ? buf : new Uint8Array(buf);
    if (!(u8[0] === 0xD0 && u8[1] === 0xCF && u8[2] === 0x11 && u8[3] === 0xE0)) throw new Error('Это не файл Word 97–2003 (.doc). Откройте его в Word и сохраните как .docx.');
    return CFB.read(u8, { type: 'array' });
  }
  const stream = (cf, name) => { const e = CFB.find(cf, '/' + name); return e && e.content ? toU8(e.content) : null; };

  /* ---------- FIB, Clx, куски текста ---------- */
  function readFib(cf) {
    const wd = stream(cf, 'WordDocument'); if (!wd || u16(wd, 0) !== 0xA5EC) throw new Error('Не удалось прочитать .doc: повреждён или не Word 97–2003. Сохраните файл как .docx.');
    const flags = u16(wd, 0x0A);
    if (flags & 0x0100) throw new Error('Файл .doc зашифрован паролем — снимите пароль и сохраните как .docx.');
    const tbl = stream(cf, (flags & 0x0200) ? '1Table' : '0Table'); if (!tbl) throw new Error('В .doc нет таблицы текста. Сохраните файл как .docx.');
    const csw = u16(wd, 32), posLw = 34 + csw * 2, cslw = u16(wd, posLw), lw = posLw + 2;
    const ccpText = u32(wd, lw + 12);
    const posFc = lw + cslw * 4, cb = u16(wd, posFc), pairs = posFc + 2;
    const pair = i => (i < cb) ? [u32(wd, pairs + i * 8), u32(wd, pairs + i * 8 + 4)] : [0, 0];
    return { wd, tbl, ccpText, clx: pair(33), btePapx: pair(13) };
  }
  function readPieces(fib) {
    const { tbl } = fib, [fc, lcb] = fib.clx; let p = fc; const end = fc + lcb;
    while (p < end && tbl[p] === 1) p += 3 + u16(tbl, p + 1);
    if (tbl[p] !== 2) throw new Error('Не удалось прочитать структуру текста .doc.');
    const len = u32(tbl, p + 1), n = Math.floor((len - 4) / 12), base = p + 5, pieces = [];
    for (let i = 0; i < n; i++) {
      const cp0 = u32(tbl, base + i * 4), cp1 = u32(tbl, base + (i + 1) * 4), d = base + (n + 1) * 4 + i * 8, raw = u32(tbl, d + 2);
      const comp = !!(raw & 0x40000000), fcv = raw & 0x3FFFFFFF;
      pieces.push({ cp0, cp1, comp, fc: comp ? fcv / 2 : fcv });
    }
    return pieces;
  }
  const dec1252 = (typeof TextDecoder !== 'undefined') ? new TextDecoder('windows-1252') : null;
  function mainText(fib, pieces) {
    const { wd } = fib; let out = '';
    for (const pc of pieces) {
      if (pc.cp0 >= fib.ccpText) break;
      const cp1 = Math.min(pc.cp1, fib.ccpText), n = cp1 - pc.cp0; if (n <= 0) continue;
      if (pc.comp) {
        const bytes = wd.subarray(pc.fc, pc.fc + n);
        if (dec1252) out += dec1252.decode(bytes); else { let s = ''; for (let i = 0; i < bytes.length; i++) s += String.fromCharCode(bytes[i]); out += s; }
      } else {
        let s = ''; for (let i = 0; i < n; i++) s += String.fromCharCode(u16(wd, pc.fc + i * 2)); out += s;
      }
    }
    return out;
  }

  /* ---------- PAPX: какие абзацы лежат в таблице ---------- */
  function papxTable(fib, pieces) {
    try {
      const { wd, tbl } = fib, [fc, lcb] = fib.btePapx; if (!lcb) return null;
      const n = Math.floor((lcb - 4) / 8), runs = [];
      for (let i = 0; i < n; i++) {
        const fcStart = u32(tbl, fc + i * 4), pn = u32(tbl, fc + (n + 1) * 4 + i * 4) & 0x3FFFFF, pg = pn * 512;
        if (pg + 512 > wd.length) continue;
        const crun = wd[pg + 511];
        for (let r = 0; r < crun; r++) {
          const a = u32(wd, pg + r * 4), b = u32(wd, pg + (r + 1) * 4), off = wd[pg + (crun + 1) * 4 + r * 13] * 2;
          let inTable = false, ttp = false;
          if (off) {
            let cbp = wd[pg + off], q = pg + off + 1;
            if (cbp === 0) { cbp = wd[q] * 2; q += 1; } else cbp = cbp * 2 - 1;
            const gEnd = q + cbp; let g = q + 2;       // 2 байта istd
            while (g + 2 <= gEnd) {
              const sprm = u16(wd, g), spra = sprm >> 13; g += 2; let sz;
              if (spra === 0 || spra === 1) sz = 1; else if (spra === 2 || spra === 4 || spra === 5) sz = 2; else if (spra === 3) sz = 4; else if (spra === 7) sz = 3; else { sz = 1 + wd[g]; if (sprm === 0xD608) sz = 1 + u16(wd, g); }
              if (sprm === 0x2416 && wd[g]) inTable = true;
              if (sprm === 0x2417 && wd[g]) ttp = true;
              g += sz;
            }
          }
          runs.push([a, b, inTable, ttp]);
        }
      }
      runs.sort((x, y) => x[0] - y[0]);
      const fcOf = cp => { for (const pc of pieces) if (cp >= pc.cp0 && cp < pc.cp1) return pc.comp ? pc.fc + (cp - pc.cp0) : pc.fc + (cp - pc.cp0) * 2; return -1; };
      return cp => {
        const f = fcOf(cp); if (f < 0) return null;
        let lo = 0, hi = runs.length - 1;
        while (lo <= hi) { const m = (lo + hi) >> 1, r = runs[m]; if (f < r[0]) hi = m - 1; else if (f >= r[1]) lo = m + 1; else return { inTable: r[2], ttp: r[3] }; }
        return null;
      };
    } catch (e) { return null; }
  }

  /* ---------- блоки: абзацы и таблицы ---------- */
  function cleanText(raw) {
    let out = '', depth = [];     // стек полей: true — идёт код поля (его пропускаем)
    for (let i = 0; i < raw.length; i++) {
      const ch = raw[i], c = ch.charCodeAt(0);
      if (c === 0x13) { depth.push(true); continue; }
      if (c === 0x14) { if (depth.length) depth[depth.length - 1] = false; continue; }
      if (c === 0x15) { depth.pop(); continue; }
      if (depth.some(Boolean)) continue;
      if (c === 0x0B) { out += '\n'; continue; }
      if (c === 0x1E) { out += '-'; continue; }
      if (c === 0x1F || c === 0x01 || c === 0x08 || c === 0x02 || c === 0x05 || c === 0x03 || c === 0x04) continue;
      if (c === 0x0C) { out += '\f'; continue; }
      out += ch;
    }
    return out.replace(/\u0000/g, '');
  }
  function parseBlocks(buf) {
    const cf = openCf(buf), fib = readFib(cf), pieces = readPieces(fib), raw = mainText(fib, pieces), look = papxTable(fib, pieces);
    const blocks = []; let table = null, row = [], cell = [], seg = '', pending = false;
    const flushTable = () => { if (table) { if (row.length) table.rows.push(row); blocks.push(table); table = null; row = []; cell = []; } };
    for (let i = 0; i < raw.length; i++) {
      const ch = raw[i];
      if (ch !== '\r' && ch !== '\x07') { seg += ch; continue; }
      const text = cleanText(seg); seg = '';
      if (ch === '\r') {
        const info = look ? look(i) : null;
        if (info && info.inTable) { if (!table) { table = { type: 'table', rows: [] }; row = []; cell = []; } cell.push(text); }
        else if (!look && table && false) { cell.push(text); }
        else { flushTable(); text.split('\f').forEach((t, k, a) => { blocks.push({ type: 'p', text: t }); }); }
      } else {
        const info = look ? look(i) : null;
        const isRowEnd = info ? !!info.ttp : (!text && !cell.length && row.length > 0);
        if (!table) { table = { type: 'table', rows: [] }; row = []; cell = []; }
        if (isRowEnd) { table.rows.push(row); row = []; cell = []; }
        else { cell.push(text); row.push(cell); cell = []; }
      }
    }
    if (seg) { flushTable(); const t = cleanText(seg); if (t.trim()) blocks.push({ type: 'p', text: t }); }
    flushTable();
    return blocks;
  }
  function blocksToText(blocks) {
    const out = [];
    for (const b of blocks) {
      if (b.type === 'p') { const t = b.text.replace(/\s+/g, ' ').trim(); if (t) out.push(t); }
      else for (const r of b.rows) { const cells = r.map(c => c.join(' ').replace(/\s+/g, ' ').trim()); if (cells.some(Boolean)) out.push(cells.join(' | ')); }
    }
    return out.join('\n');
  }

  /* ---------- .docx из блоков ---------- */
  const esc = s => String(s).replace(/[&<>]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c])).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');
  const runs = text => String(text).split('\n').map((l, i) => (i ? '<w:r><w:br/></w:r>' : '') + '<w:r><w:t xml:space="preserve">' + esc(l) + '</w:t></w:r>').join('');
  async function toDocx(buf) {
    if (!JSZip) throw new Error('Нет библиотеки JSZip');
    const blocks = parseBlocks(buf); let body = '';
    for (const b of blocks) {
      if (b.type === 'p') body += '<w:p><w:r><w:t xml:space="preserve">' + '</w:t></w:r>' .replace('<w:r><w:t xml:space="preserve"></w:t></w:r>', '') + runs(b.text) + '</w:p>';
      else {
        body += '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>' + ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(s => '<w:' + s + ' w:val="single" w:sz="4" w:space="0" w:color="808080"/>').join('') + '</w:tblBorders></w:tblPr>';
        const cols = Math.max.apply(null, b.rows.map(r => r.length).concat([1]));
        body += '<w:tblGrid>' + Array(cols).fill('<w:gridCol w:w="' + Math.floor(9000 / cols) + '"/>').join('') + '</w:tblGrid>';
        for (const r of b.rows) {
          body += '<w:tr>';
          for (let i = 0; i < cols; i++) { const c = r[i] || ['']; body += '<w:tc><w:tcPr><w:tcW w:w="' + Math.floor(9000 / cols) + '" w:type="dxa"/></w:tcPr>' + c.map(t => '<w:p>' + runs(t) + '</w:p>').join('') + '</w:tc>'; }
          body += '</w:tr>';
        }
        body += '</w:tbl><w:p/>';
      }
    }
    const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
    const z = new JSZip();
    z.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>');
    z.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
    z.file('word/_rels/document.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>');
    z.file('word/styles.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + '><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman" w:eastAsia="Times New Roman"/><w:sz w:val="24"/><w:szCs w:val="24"/><w:lang w:val="ru-RU"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>');
    z.file('word/document.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS + '><w:body>' + body + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="850" w:bottom="1134" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>');
    return z.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  }

  /* ---------- картинки ---------- */
  function pngEnd(u, i) {
    let p = i + 8;
    while (p + 12 <= u.length) { const len = u32(Object.assign ? new Uint8Array([u[p + 3], u[p + 2], u[p + 1], u[p]]) : u, 0); const type = String.fromCharCode(u[p + 4], u[p + 5], u[p + 6], u[p + 7]); p += 12 + len; if (type === 'IEND') return p; if (len > u.length) return -1; }
    return -1;
  }
  function jpegEnd(u, i) {
    let p = i + 2;
    while (p + 4 <= u.length) {
      if (u[p] !== 0xFF) return -1;
      const m = u[p + 1];
      if (m === 0xD9) return p + 2;
      if (m === 0xFF) { p++; continue; }
      if (m === 0x01 || (m >= 0xD0 && m <= 0xD8)) { p += 2; continue; }
      const len = (u[p + 2] << 8) | u[p + 3];
      p += 2 + len;
      if (m === 0xDA) { while (p + 1 < u.length && !(u[p] === 0xFF && u[p + 1] !== 0 && !(u[p + 1] >= 0xD0 && u[p + 1] <= 0xD7))) p++; }
    }
    return -1;
  }
  function scanImages(u, out, seen) {
    for (let i = 0; i + 8 < u.length; i++) {
      if (u[i] === 0x89 && u[i + 1] === 0x50 && u[i + 2] === 0x4E && u[i + 3] === 0x47 && u[i + 4] === 0x0D && u[i + 5] === 0x0A) {
        const e = pngEnd(u, i); if (e > i + 60) { const b = u.slice(i, e), k = 'p' + b.length + ':' + b[40] + b[b.length - 20]; if (!seen.has(k)) { seen.add(k); out.push({ mime: 'image/png', ext: 'png', bytes: b }); } i = e - 1; }
      } else if (u[i] === 0xFF && u[i + 1] === 0xD8 && u[i + 2] === 0xFF && (u[i + 3] === 0xE0 || u[i + 3] === 0xE1 || u[i + 3] === 0xDB || u[i + 3] === 0xEE)) {
        const e = jpegEnd(u, i); if (e > i + 200) { const b = u.slice(i, e), k = 'j' + b.length + ':' + b[100] + b[b.length - 30]; if (!seen.has(k)) { seen.add(k); out.push({ mime: 'image/jpeg', ext: 'jpg', bytes: b }); } i = e - 1; }
      }
    }
  }
  function images(buf) {
    const cf = openCf(buf), out = [], seen = new Set();
    for (const n of ['Data', 'WordDocument', '1Table', '0Table']) { const s = stream(cf, n); if (s) scanImages(s, out, seen); }
    return out;
  }

  return { text: buf => blocksToText(parseBlocks(buf)), blocks: parseBlocks, blocksToText, toDocx, images };
});
