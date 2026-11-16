// 施設基準の章：ページごとの見出しを取り出し、前回の（施設基準）の行の今回のページを探す（Python版 sisetsu.py の移植）
import { nz } from './pdftext.js';
import { pageNumbers } from './draft.js';
import { ratio } from './util.js';

export function titlesOf(P) {
  const out = [];
  for (const l of P.lines) {
    const sp = l.filter(s => s.s.trim()); if (!sp.length) continue;
    const s = sp[0]; let t = nz(sp.map(x => x.s).join('')); const f = s.font;
    if (f.includes('ShinGo') && s.size < 13 && s.x > 540) continue;
    if (s.vertical) continue;
    if ((f.includes('HiraKakuPro-W6') && s.size >= 9) || (f.includes('HiraMinPro-W6') && /^[0-9０-９]+/.test(t))
      || (f.includes('Jun34') && s.size >= 11) || (f.includes('ShinGoPro-Bold') && s.size > 13 && s.size < 30)) {
      if (t === '告示' || t === '通知' || t.startsWith('■') || [...t].length < 3) continue;
      t = t.replace(/➡.*$/, '').replace(/^[0-9０-９]+/, '');
      out.push(t);
    }
  }
  const j = [];
  for (const t of out) {
    if (j.length && (t === '施設基準' || t === '施設基準等' || /(に関する|及び)$/.test(j[j.length - 1]))) j[j.length - 1] += t;
    else j.push(t);
  }
  return j.map(t => t.replace(/(に関する)?施設基準(等)?$/, '') || t);
}

// 施設基準の章の原稿PDF → [{p, titles, text}]
export function readSisetsu(pages) {
  const { nums } = pageNumbers(pages);
  const out = [];
  pages.forEach((P, i) => { if (nums[i] != null) out.push({ p: nums[i], titles: titlesOf(P), text: nz(P.text) }); });
  return out;
}

export class Sisetsu {
  // neu: readSisetsu の結果、oldTitles: {前回のページ: [見出し]}
  constructor(neu, oldTitles) { this.neu = neu; this.oldT = oldTitles || {}; this.lo = Math.min(...neu.map(x => x.p)); }
  oldTitles(p) { return [...(this.oldT[p] || [])]; }
  find(parent, oldp = null, name = null, alts = [], kinyu = false) {
    let P = [parent, ...alts].flatMap(t => t.split(/[（）()]/)).map(nz).filter(x => [...x].length >= 2);
    P = P.map(x => x.replace(/[0-9０-９・]+$/, '') || x);
    if (kinyu) {   // 施設基準記入例：まず本体のページを探し、その近くの「記入例」のページ
      const [p0, nt] = this.find(parent, oldp, name, alts);
      if (!p0) return [null, nt];
      for (const { p, titles, text } of this.neu)
        if (p >= p0 && p <= p0 + 3 && text.includes('記入例') && (p === p0 || !titles.length || titles.every(t => t.includes('届出'))))
          return [p, nt.replace('のページ（仮）', 'の記入例のページ（仮）')];
      return [null, `施設基準の章に記入例が見つからない（本体は p.${p0}）`];
    }
    const keys = [];
    if (name) keys.push(nz(name));
    if (oldp) {
      const ot = this.oldTitles(oldp);
      const sc = t => Math.max(0, ...P.map(x => ratio(t, x)));
      ot.sort((a, b) => sc(b) - sc(a));
      keys.push(...ot.slice(0, 2));
    }
    keys.push(...P);
    for (const k of keys) {
      if ([...k].length < 2) continue;
      let best = null, bs = 0;
      for (const { p, titles } of this.neu) for (const t of titles) {
        let s = t === k ? 1 : ratio(t, k);
        if (t.includes(k) || k.includes(t)) s = Math.max(s, Math.min([...t].length, [...k].length) >= 4 ? 0.9 : 0);
        if (s > bs) { bs = s; best = [p, t]; }
      }
      if (best && bs >= 0.8) {
        const pg = best[0];
        const ts = this.neu.find(x => x.p === pg).titles;
        const lab = ts.reduce((a, t) => { const v = Math.max(0, ...P.map(x => ratio(t, x) + ((x.includes(t) || t.includes(x)) ? 1 : 0))); return v > a[0] ? [v, t] : a; }, [-1, ts[0]])[1];
        return [pg, `施設基準の章の「${lab}」のページ（仮）`];
      }
    }
    for (const k of keys) {   // 見出しに無ければ本文の中の語で
      if ([...k].length >= 6) {
        const hit = this.neu.filter(x => x.text.includes(k) && x.p > this.lo + 1).map(x => x.p);
        if (hit.length >= 1 && hit.length <= 3) return [hit[0], `施設基準の章の本文に「${k}」があるページ（仮）`];
      }
    }
    return [null, '施設基準の章で見つからない（仮入れ不可）'];
  }
}
