// 新原稿PDFが前回のどの章にあたるか：タブ（反転している章名）と扉ページの中身で判断
import { nz } from './pdftext.js';
import { ratio } from './util.js';
export function matchChapter(old, R, tobEntries = null) {
  let best = null, bs = -1;
  for (const c of old.chapters) {
    let s = R.tab ? ratio(nz(R.tab), nz(c.tab)) : 0;
    if (R.tobiraPage != null) s += Math.max(0, 1 - Math.abs(R.tobiraPage - c.tob) / 60) * 0.3;
    if (s > bs) { bs = s; best = c; }
  }
  return { chapter: best, score: bs };
}
