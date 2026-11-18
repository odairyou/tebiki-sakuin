// 前回の完成品PDFの読み取り：扉ページ・巻末索引・章の範囲・施設基準ページの見出し（Python版 base.py の移植＋自動判定）
import { nz } from './pdftext.js';
import { pageNumbers, activeTab, bigTitle, readDraft } from './draft.js';
import { titlesOf } from './sisetsu.js';
import { ratio } from './util.js';

const LEADER = /^(.*?)\s*…+\s*([\d・\s]+)$/;
function pageLines(P) {
  return P.lines.map(l => l.filter(s => s.s.trim() && !s.font.includes('ShinGo') && s.x < 560 && !s.vertical))
    .filter(l => l.length).map(l => l.map(s => s.s).join('').trim());
}

export function parseLeaderLines(lines, backIndex = false) {
  const ents = []; let buf = ''; let head = null;
  for (const l of lines) {
    if (/^−\s*\d+\s*−$/.test(l) || nz(l) === '索引') continue;
    let m = l.match(/^【(.+)】$/);
    if (m && backIndex) { head = m[1]; continue; }
    const l2 = buf + l;
    m = l2.match(LEADER);
    if (m) {
      buf = '';
      let txt = m[1].trim();
      if (!txt && ents.length && !ents[ents.length - 1].pages.length) txt = ents.pop().text;
      ents.push({ text: txt, pages: (m[2].match(/\d+/g) || []).map(Number), head });
    } else if (l2.includes('…')) {
      buf = l2.replace(/\s*…+\s*$/, '');
    } else {
      if (!buf && [...nz(l)].length === 1) continue;   // 「索」「引」など1文字だけの行
      if (buf || l.startsWith('（') || (backIndex && [...l].length > 22 && !l.endsWith('）'))) buf = l2;
      else ents.push({ text: l, pages: [], head });
    }
  }
  for (const e of ents) e.sub = e.text.startsWith('（');
  return ents;
}

export function tabCanon(known) {
  return t => {
    if (!t) return t;
    let best = t, bs = 0;
    for (const k of known) { const r = ratio(t, k); if (r > bs) { bs = r; best = k; } }
    return bs >= 0.6 ? best : t;
  };
}
const leaders = P => (P.text.match(/…/g) || []).length;

// pages: readAll の結果
export function readOld(pages) {
  const { nums } = pageNumbers(pages);
  const info = pages.map((P, i) => ({ i, p: nums[i], tab: activeTab(P), lead: leaders(P), kana: /^【.】$/m.test(P.text) }));
  // 巻末索引：後ろのほうの、点線の多いページが続くところ（【ア】などの見出しがある）
  let bStart = info.findIndex(x => x.lead >= 20 && x.kana && !x.tab);
  if (bStart < 0) bStart = info.findIndex(x => x.lead >= 20 && x.kana);
  let bEnd = bStart;
  while (bEnd + 1 < info.length && info[bEnd + 1].lead >= 20) bEnd++;
  const tobPages = info.filter(x => x.lead >= 20 && x.tab && (bStart < 0 || x.i < bStart));
  // タブの文字は取り出しで一部欠けることがあるので、扉のタブにいちばん近いものにそろえる
  const canon = tabCanon(tobPages.map(x => x.tab));
  for (const x of info) x.tab = canon(x.tab);
  // 章：扉ページごと。範囲は「同じタブが反転しているページ」の最初から、次の章の手前まで
  const firstOf = tab => Math.min(...info.filter(y => y.tab === tab && y.p != null).map(y => y.p));
  const chapters = tobPages.map(x => {
    const lines = pageLines(pages[x.i]).filter(l => [...nz(l)].length > 1);
    const ents = parseLeaderLines(lines).filter(e => e.pages.length);
    return { tab: x.tab, title: bigTitle(pages[x.i]), tob: x.p, lo: Math.min(x.p, firstOf(x.tab)), hi: null, entries: ents };
  });
  // 施設基準の章（扉の無い章）
  const cnt = {};
  for (const x of info) if (x.tab && x.p != null && !chapters.some(c => c.tab === x.tab)) cnt[x.tab] = (cnt[x.tab] || 0) + 1;
  const sTab = Object.keys(cnt).find(t => t.includes('施設基準')) || Object.entries(cnt).sort((a, b) => b[1] - a[1])[0]?.[0];
  const sisetsu = sTab ? { tab: sTab, lo: firstOf(sTab), hi: Math.max(...info.filter(y => y.tab === sTab && y.p != null).map(y => y.p)) } : null;
  const endAll = bStart >= 0 ? info[bStart].p - 1 : info[info.length - 1].p;
  chapters.forEach((c, k) => {
    const nx = chapters[k + 1]?.lo ?? (sisetsu && sisetsu.lo > c.tob ? sisetsu.lo : endAll + 1);
    c.hi = nx - 1;
  });
  if (sisetsu) { const after = chapters.find(c => c.lo > sisetsu.lo); sisetsu.hi = after ? after.lo - 1 : endAll; }
  // 章ごとの見出し・略称（前回その章に出ていた語。索引に載せたかどうかは問わない）
  for (const c of chapters) {
    const sub = info.filter(x => x.p != null && x.p >= c.lo && x.p <= c.hi).map(x => pages[x.i]);
    try {
      const R = readDraft(sub);
      c.heads = [...new Set([...R.heads.map(h => nz(h[2])), ...R.headsMain.map(h => nz(h[2])), ...R.abbrs.map(a => nz(a[2])), ...R.abbrTable.flatMap(([a, b]) => [nz(a), nz(b)])])];
    } catch (e) { c.heads = []; }
  }
  const oldTitles = {};
  if (sisetsu) info.forEach(x => { if (x.p != null && x.p >= sisetsu.lo - 20 && x.p <= sisetsu.hi) oldTitles[x.p] = titlesOf(pages[x.i]); });
  // 巻末索引
  let lines = [];
  if (bStart >= 0) for (let i = bStart; i <= bEnd; i++) lines = lines.concat(pageLines(pages[i]));
  let back = parseLeaderLines(lines, true);
  // 外字（丸囲みの「未」）：PDFからは「未請求」や「未」「○請求」に分かれて取れるので「○未請求」にそろえる
  back = back.filter((e, k) => !(nz(e.text) === '未' && !e.pages.length && back[k + 1] && /^○請求$/.test(nz(back[k + 1].text))));
  for (const e of back) if (['未請求', '○請求', '○未請求'].includes(nz(e.text))) e.text = '○未請求';
  let main = null;
  for (const e of back) { if (!e.sub) main = e.text; e.main = main; }
  return { chapters, back, sisetsu, oldTitles, backPages: bStart >= 0 ? [info[bStart].p, info[bEnd].p] : null, nPages: pages.length };
}
