// 小道具：difflib.SequenceMatcher.ratio と同じ計算など
export function ratio(a, b) {
  a = [...a]; b = [...b];
  const la = a.length, lb = b.length;
  if (!la && !lb) return 1;
  const b2j = new Map();
  b.forEach((ch, j) => { if (!b2j.has(ch)) b2j.set(ch, []); b2j.get(ch).push(j); });
  // autojunk（Pythonは長さ200以上で人気の文字を除く）
  if (lb >= 200) { const ntest = Math.floor(lb / 100) + 1; for (const [ch, idx] of [...b2j]) if (idx.length > ntest) b2j.delete(ch); }
  function longest(alo, ahi, blo, bhi) {
    let besti = alo, bestj = blo, bestsize = 0; let j2len = new Map();
    for (let i = alo; i < ahi; i++) {
      const nj = new Map();
      for (const j of (b2j.get(a[i]) || [])) {
        if (j < blo) continue; if (j >= bhi) break;
        const k = (j2len.get(j - 1) || 0) + 1; nj.set(j, k);
        if (k > bestsize) { besti = i - k + 1; bestj = j - k + 1; bestsize = k; }
      }
      j2len = nj;
    }
    while (besti > alo && bestj > blo && !b2j.has(b[bestj - 1]) && a[besti - 1] === b[bestj - 1]) { besti--; bestj--; bestsize++; }
    while (besti + bestsize < ahi && bestj + bestsize < bhi && !b2j.has(b[bestj + bestsize]) && a[besti + bestsize] === b[bestj + bestsize]) bestsize++;
    return [besti, bestj, bestsize];
  }
  let M = 0; const q = [[0, la, 0, lb]];
  while (q.length) {
    const [alo, ahi, blo, bhi] = q.pop();
    const [i, j, k] = longest(alo, ahi, blo, bhi);
    if (k) { M += k; if (alo < i && blo < j) q.push([alo, i, blo, j]); if (i + k < ahi && j + k < bhi) q.push([i + k, ahi, j + k, bhi]); }
  }
  return 2 * M / (la + lb);
}
export const uniq = arr => [...new Set(arr)];
export const sortNum = arr => [...arr].sort((x, y) => x - y);
export const clen = s => [...s].length;
