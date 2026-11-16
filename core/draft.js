// 新原稿PDFの読み取り：ページごとの領域・見出し・略称・目印（Python版 newdraft.py の移植）
import { nz } from './pdftext.js';

const RED = 0xE60012, WHITE = 0xFFFFFF;
const SKIP_HEAD = /^(点|\(.*?につき\)|（.*?につき）|レセプト記載|[+＋]?[\d,，]+|注\d*|略)$/;
export const isBold = f => ['Jun34', 'JunNum34', 'W6', 'Bold', 'Jun501', 'Medium'].some(k => f.includes(k));
export const pageNo = text => { const m = text.match(/−\s*(\d+)\s*−/); return m ? +m[1] : null; };

// ページ番号（ノンブル）の無いページは前後から補う
export function pageNumbers(pages) {
  const nums = pages.map(P => pageNo(P.text));
  for (let i = 0; i < nums.length; i++) if (nums[i] == null && i > 0 && nums[i - 1] != null) nums[i] = nums[i - 1] + 1;
  for (let i = nums.length - 1; i >= 0; i--) if (nums[i] == null && i + 1 < nums.length && nums[i + 1] != null) nums[i] = nums[i + 1] - 1;
  const inferred = new Set(nums.filter((n, i) => n != null && pageNo(pages[i].text) == null));
  return { nums, inferred };
}

// タブ（ページ右端の章の見出し）で、色が反転している＝今の章
export function activeTab(P) {
  const ch = P.items.filter(i => i.x > 540 && i.color === WHITE && i.size < 13);
  ch.sort((a, b) => (Math.round(b.x / 4) - Math.round(a.x / 4)) || (a.y - b.y));
  return nz(ch.map(i => i.s).join(''));
}
// 扉ページ・章頭の大きなタイトル（例「８． 処置」）
export function bigTitle(P) {
  const t = P.items.filter(i => i.size > 20 && i.x < 540);
  return t.map(i => i.s).join('').replace(/\s+/g, '');
}

function bodyLines(P) {
  return P.lines.map(l => l.filter(s => s.s.trim() && !s.font.includes('ShinGo') && s.x < 560 && !s.vertical && s.y > 45)).filter(l => l.length);
}

export function readDraft(pages) {
  const Z = {};
  const Zp = p => (Z[p] ||= { LB: '', AB: '', LN: '', RB: '', RN: '', ALL: '' });
  const heads = [], abbrs = [], boxes = [], titles = [], labels = [], abbrTable = [];
  let tobiraPage = null; const tablePages = new Set(); const pgs = [];
  const { nums, inferred } = pageNumbers(pages);
  const tabs = {};
  pages.forEach((P, idx) => {
    const p = nums[idx];
    if (p == null || !P.text.trim()) return;
    const tab = activeTab(P); if (tab) tabs[tab] = (tabs[tab] || 0) + 1;
    if ((P.text.match(/…/g) || []).length >= 20) { tobiraPage = p; return; }
    const lines = bodyLines(P);
    const flat = nz(lines.flat().map(s => s.s).join('')).slice(0, 40);
    if (flat.includes('項目') && flat.includes('略称')) {
      tablePages.add(p);
      const rows = new Map();
      for (const sp of lines) for (const s of sp) {
        const y = Math.round(s.y), x = s.x, t = nz(s.s);
        if (!t.replace(/項目|略称/g, '')) continue;
        if (!rows.has(y)) rows.set(y, []);
        rows.get(y).push([x, t]);
      }
      for (const [, parts] of rows) {
        parts.sort((a, b) => a[0] - b[0]);
        const names = parts.filter(([x]) => x < 290).map(a => a[1]);
        const abs = parts.filter(([x]) => x >= 290).map(a => a[1]);
        if (names.length && abs.length && names.join('').length > 1) abbrTable.push([names.join(''), abs.join('')]);
      }
      return;
    }
    pgs.push(p); const z = Zp(p);
    for (const sp0 of lines) {
      // 外字（丸囲みの「未」）は「未」と「○」に分かれて取れるので「未」にそろえる
      const sp = sp0.map(s => ({ ...s, s: s.s.replace(/未\s*○|○\s*未/g, '未') }))
        .filter((s, k, a) => !(nz(s.s) === '○' && ((k > 0 && nz(a[k - 1].s).endsWith('未')) || (k + 1 < a.length && nz(a[k + 1].s).startsWith('未')))));
      const x = sp[0].x, y = Math.round(sp[0].y);
      for (const s of sp) {
        const t = nz(s.s), f = s.font, c = s.color, sx = s.x;
        z.ALL += t;
        if (sx < 190) {
          if (c === RED && nz(sp[0].s) === '略') { z.AB += '|' + t + '|'; abbrs.push([p, Math.round(s.y), t]); }
          else if (isBold(f) && c !== WHITE) z.LB += t;
          else z.LN += t;
        } else {
          if (isBold(f)) z.RB += t; else z.RN += t;
        }
      }
      const line = nz(sp.map(s => s.s).join(''));
      const s0 = sp[0];
      if (x < 190 && isBold(s0.font) && s0.color !== WHITE && s0.color !== RED) {
        let hl = line.replace(/\(.*?につき\)|（.*?につき）/g, '');
        hl = hl.replace(/[\d,，]+点.*$/, '').replace(/^[+＋]+/, '').replace(/[+＋]+$/, '');
        if (hl && !SKIP_HEAD.test(hl) && !/^[\d,，]/.test(hl)) heads.push([p, y, hl, x]);
      }
      if (sp.some(s => s.color === WHITE) && x >= 190) boxes.push([p, y, line]);
      if (sp.some(s => s.size >= 14 && s.font.includes('Bold'))) titles.push([p, y, line]);
      if (/〈.*?〉/.test(line) && isBold(s0.font)) labels.push([p, y, line]);
    }
  });
  // 見出しの続き行をつなぐ
  const merged = [];
  for (const h of heads) {
    const [p, y, t, x] = h; const q = merged[merged.length - 1];
    if (q && q[0] === p && y - q[1] > 0 && y - q[1] <= 15 && !/^[(<+]/.test(t) && !q[2].endsWith(')') && t !== '+')
      merged[merged.length - 1] = [p, y, q[2] + t, q[3]];
    else merged.push(h);
  }
  const BAD = /[。]|^[+＋]$|^[IⅠⅡⅢ]+$|^<|(で|を|に|は|が|と)$|^\(|の点数$/;
  const CH = new Set(['処置', '手術', '麻酔', '歯冠修復及び欠損補綴', '検査', '投薬', '注射', '画像診断', '在宅医療', '医学管理等', '歯周病', '歯科矯正', '病理診断', 'リハビリテーション']);
  const len = t => [...t].length;
  const headsF = merged.filter(([p, y, t]) => !CH.has(t) && len(t) >= 2 && len(t) <= 30 && !BAD.test(t)).map(([p, y, t]) => [p, y, t]);
  const STOP = new Set(['通則', '項目', '区分', '材料', '注意事項', 'その他', '製品名', '薬価', '一般的名称', '項目名', '薬剤', '備考', '摘要', '算定できる場合']);
  const headsMain = merged.filter(([p, y, t, x]) => ((x >= 44 && x <= 49) || (x >= 61 && x <= 66)) && !CH.has(t) && !STOP.has(t)
    && len(t) >= 2 && len(t) <= 30 && !BAD.test(t) && !inferred.has(p)).map(([p, y, t]) => [p, y, t]);
  const tab = Object.entries(tabs).sort((a, b) => b[1] - a[1])[0]?.[0] || '';
  pgs.sort((a, b) => a - b);
  return { Z, pages: pgs, heads: headsF, headsMain, abbrs, boxes, titles, labels, abbrTable, tobiraPage, tablePages, inferred, tab };
}
