// PDFから文字・位置・フォント・色・行を取り出す（pdf.js）
export const nz = s => String(s ?? '').normalize('NFKC').replace(/\s+/g, '');

export async function openPdf(pdfjs, data, opt = {}) {
  return await pdfjs.getDocument({ data, useSystemFonts: false, disableFontFace: true, isEvalSupported: false,
    cMapUrl: opt.cMapUrl, cMapPacked: true, standardFontDataUrl: opt.standardFontDataUrl, verbosity: 0 }).promise;
}

function hexOf(a) {
  if (typeof a[0] === 'string') return parseInt(a[0].replace('#', ''), 16);
  return (Math.round(a[0]) << 16) + (Math.round(a[1]) << 8) + Math.round(a[2]);
}

export async function readPage(pdfjs, page) {
  const OPS = pdfjs.OPS;
  const vp = page.getViewport({ scale: 1 });
  const H = vp.height;
  const ol = await page.getOperatorList();
  // 描画命令から1文字ずつの色を拾う
  const seq = []; let fill = 0; const stack = [];
  for (let i = 0; i < ol.fnArray.length; i++) {
    const fn = ol.fnArray[i], a = ol.argsArray[i];
    if (fn === OPS.save) stack.push(fill);
    else if (fn === OPS.restore) fill = stack.length ? stack.pop() : fill;
    else if (fn === OPS.setFillRGBColor) fill = hexOf(a);
    else if (fn === OPS.setFillGray) { const v = Math.round((typeof a[0] === 'number' ? a[0] : 0) * 255); fill = (v << 16) + (v << 8) + v; }
    else if (fn === OPS.showText || fn === OPS.showSpacedText) {
      for (const g of a[0]) {
        if (g && typeof g === 'object' && g.unicode) {
          if (!g.unicode.trim()) { seq.push([' ', fill]); continue; }
          for (const ch of nz(g.unicode)) seq.push([ch, fill]);
        }
      }
    }
  }
  const tc = await page.getTextContent({ disableNormalization: true });
  const items = []; let j = 0;
  for (const it of tc.items) {
    if (!it.str) continue;
    if (!it.str.trim()) {   // 空白は行をつなぐ手がかりとしてだけ使う
      const t = it.transform; if (it.width > 0) items.push({ s: it.str, x: t[4], x1: t[4] + it.width, base: H - t[5], space: true, idx: items.length });
      continue;
    }
    let font = it.fontName;
    try { font = page.commonObjs.get(it.fontName).name || font; } catch (e) {}
    font = String(font).replace(/^[A-Z]{6}\+/, '');
    const t = it.transform;
    const size = Math.hypot(t[2], t[3]) || it.height;
    const x = t[4], base = H - t[5];
    const w = it.width;
    // 色：1文字ずつ描画命令の文字の並びに当てはめ、色が変わるところで分ける
    const raw = [...it.str]; const cols = [];
    let last = null, spaceBefore = false, firstChar = true;
    for (const c of raw) {
      const n = nz(c); let col = last;
      if (n) {
        let k = j, found = -1;
        while (k < seq.length && k < j + 60) { if (seq[k][0] === n[0]) { found = k; break; } k++; }
        if (found >= 0) {
          if (firstChar && found > 0 && seq[found - 1][0] === ' ') spaceBefore = true;
          col = seq[found][1]; j = found + n.length;
        }
        firstChar = false;
      }
      cols.push(col); if (col !== null) last = col;
    }
    const first = cols.find(c => c !== null) ?? 0;
    for (let i = 0; i < cols.length; i++) if (cols[i] === null) cols[i] = i ? cols[i - 1] : first;
    const vertical = Math.abs(t[1]) > Math.abs(t[0]);
    let st = 0;
    for (let i = 1; i <= raw.length; i++) {
      if (i === raw.length || cols[i] !== cols[st]) {
        const s = raw.slice(st, i).join('');
        if (s.trim()) {
          const xa = vertical ? x : x + w * st / raw.length, xb = vertical ? x + size : x + w * i / raw.length;
          items.push({ s, x: xa, x1: xb, y: Math.round(base - size * 0.88), base, size, font, color: cols[st], vertical, idx: items.length, sp: st === 0 && (spaceBefore || /^\s/.test(it.str)) });
        }
        st = i;
      }
    }
  }
  // 行にまとめる：同じ高さ（ベースライン）で、すき間が小さく続くもの。行の順番は描画の順
  const hor = items.filter(i => !i.vertical).sort((a, b) => a.base - b.base);
  const rows = [];
  for (const it of hor) {
    const r = rows.length && rows[rows.length - 1];
    if (r && Math.abs(it.base - r.base) <= 2.5) r.items.push(it); else rows.push({ base: it.base, items: [it] });
  }
  let lines = [];
  for (const r of rows) {
    r.items.sort((a, b) => a.x - b.x);
    let cur = null, end = -1e9;
    for (const it of r.items) {
      const gap = it.x - end;
      if (cur && (gap <= 6 || (it.sp && gap <= 1.6 * (it.size || 10)))) cur.push(it); else { cur = [it]; lines.push(cur); }
      end = Math.max(cur === null ? -1e9 : end, it.x1); if (cur.length === 1) end = it.x1;
    }
  }
  for (const it of items) if (it.vertical) lines.push([it]);
  lines = lines.map(l => l.filter(i => !i.space)).filter(l => l.length);
  // 同じ書式で続く文字列は1つにまとめる（PyMuPDFの span に相当）
  lines = lines.map(l => {
    const out = [];
    for (const it of l) {
      const q = out[out.length - 1];
      if (q && q.font === it.font && Math.abs(q.size - it.size) < 0.2 && q.color === it.color && !q.vertical
          && (it.x - q.x1 <= 1.5 || (it.sp && it.x - q.x1 <= 1.6 * it.size))) {
        out[out.length - 1] = { ...q, s: q.s + it.s, x1: Math.max(q.x1, it.x1), idx: Math.min(q.idx, it.idx) };
      } else out.push(it);
    }
    return out;
  });
  lines.sort((a, b) => Math.min(...a.map(i => i.idx)) - Math.min(...b.map(i => i.idx)));
  items.length = 0; for (const l of lines) items.push(...l);
  const text = lines.map(l => l.map(i => i.s).join('')).join('\n');
  return { H, W: vp.width, items, lines, text };
}

export async function readAll(pdfjs, doc, onProgress) {
  const out = [];
  for (let p = 1; p <= doc.numPages; p++) {
    const page = await doc.getPage(p);
    out.push(await readPage(pdfjs, page));
    page.cleanup();
    if (onProgress) onProgress(p, doc.numPages);
  }
  return out;
}
