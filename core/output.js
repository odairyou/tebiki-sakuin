// 業者に渡すデータ：扉ページ（章ごと）、巻末索引（全体）、引き継ぎファイル
import { nz } from './pdftext.js';
import { pagesText, cellText } from './workbook.js';
import { hira } from './yomi.js';

const isUndecided = r => !r.dec;
const isOn = r => r.dec === '載せる';
const pageOrder = (a, b) => (typeof a === 'number' && typeof b === 'number') ? a - b : (typeof a === 'number' ? -1 : (typeof b === 'number' ? 1 : String(a).localeCompare(String(b))));

// ---- 扉ページ
export function checkTobira(W) {
  const undecided = W.up.filter(isUndecided);
  const noPage = W.up.filter(r => isOn(r) && !r.pages.length);
  return { undecided, noPage, ok: !undecided.length && !noPage.length };
}
export async function writeTobira(ExcelJS, W) {
  const wb = new ExcelJS.Workbook();
  const name = (W.meta.title || W.meta.short || '扉').replace(/[\\/?*\[\]:]/g, '').slice(0, 31);
  const ws = wb.addWorksheet(name);
  ws.getColumn(1).width = 60; ws.getColumn(2).width = 12;
  ws.addRow([W.meta.title || W.meta.short]); ws.getRow(1).font = { bold: true, size: 12 };
  ws.addRow(['項目', 'ページ']); ws.getRow(2).font = { bold: true };
  for (const r of W.up.filter(isOn)) ws.addRow([r.text, pagesText(r.pages)]);
  ws.eachRow(r => r.eachCell(c => { c.font = { ...(c.font || {}), name: '游ゴシック' }; }));
  return await wb.xlsx.writeBuffer();
}

// ---- 50音の並べ方：小さい字は大きく、濁点・半濁点は外し、長音は前の母音として比べる（同じなら清音→濁音→半濁音）
const SMALL = { 'ぁ': 'あ', 'ぃ': 'い', 'ぅ': 'う', 'ぇ': 'え', 'ぉ': 'お', 'っ': 'つ', 'ゃ': 'や', 'ゅ': 'ゆ', 'ょ': 'よ', 'ゎ': 'わ', 'ゕ': 'か', 'ゖ': 'け' };
const VOW = {};
'あかさたなはまやらわがざだばぱぁゃゎ'.split('').forEach(c => VOW[c] = 'あ'); 'いきしちにひみりぎじぢびぴぃ'.split('').forEach(c => VOW[c] = 'い');
'うくすつぬふむゆるぐずづぶぷぅっゅゔ'.split('').forEach(c => VOW[c] = 'う'); 'えけせてねへめれげぜでべぺぇ'.split('').forEach(c => VOW[c] = 'え');
'おこそとのほもよろをごぞどぼぽぉょ'.split('').forEach(c => VOW[c] = 'お'); VOW['ん'] = 'ん';
const seion = c => { const d = c.normalize('NFD').replace(/[゙゚]/g, ''); return SMALL[d] || d; };
export function yomiKey(y) {
  // 数字は並び順の目印として残す（前回の読み「はりは１1」「はりは１2」など）
  const s = hira(String(y || '').normalize('NFKC')).replace(/[^ぁ-んー0-9]/g, '');
  let k1 = '', prev = '';
  for (const c of s) { const b = c === 'ー' ? (VOW[prev] || 'ー') : (/[0-9]/.test(c) ? c : seion(c)); k1 += b; if (c !== 'ー') prev = c; }
  const k2 = [...s].map(c => c.normalize('NFD').includes('゚') ? '2' : (c.normalize('NFD').includes('゙') ? '1' : '0')).join('');
  return [k1, k2];
}
export const cmpYomi = (a, b) => { const [a1, a2] = yomiKey(a), [b1, b2] = yomiKey(b); return a1 < b1 ? -1 : a1 > b1 ? 1 : (a2 < b2 ? -1 : a2 > b2 ? 1 : 0); };
const GYO = 'アイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワヲン';
export function headOf(y) {
  const [k1] = yomiKey(y); const c = k1.replace(/[0-9]/g, '')[0]; if (!c) return '他';
  const kk = String.fromCharCode(c.charCodeAt(0) + 0x60);
  return GYO.includes(kk) ? kk : '他';
}

// ---- 巻末索引のまとめ
// works: readWork の結果の配列、ctx: {chapters(前回の章の一覧), sis(Sisetsu or null), old(前回データ)}
export function mergeKanmatsu(works, ctx) {
  const problems = []; const notes = [];
  const byMain = new Map(); const order = [];
  let undecided = 0; const undecidedRows = [];
  for (const W of works) {
    const ch = W.meta.short || W.meta.title || W.file;
    let cur = null;
    for (const r of W.low) {
      if (isUndecided(r)) { undecided++; undecidedRows.push(`${ch}：${r.text}`); }
      if (!r.sub) {
        const k = nz(r.text);
        if (!byMain.has(k)) { byMain.set(k, { text: r.text, yomi: '', pages: [], subs: [], subKey: new Map(), on: false, chapters: new Set(), pending: [] }); order.push(k); }
        cur = byMain.get(k); cur.chapters.add(ch);
        if (r.yomi && !cur.yomi) cur.yomi = r.yomi;
        else if (r.yomi && cur.yomi && r.yomi !== cur.yomi) notes.push(`読みが章によって違う：${r.text}（${cur.yomi} / ${r.yomi}）→ ${cur.yomi} を使った`);
        if (isOn(r)) { cur.on = true; for (const p of r.pages) if (!cur.pages.includes(p)) cur.pages.push(p); }
        if (isUndecided(r)) cur.pending.push(ch);
        cur._row = r; cur._ch = ch;
      } else {
        if (!cur) { problems.push(`${ch}：副項目「${r.text}」の上に主項目が無い`); continue; }
        if (!isOn(r)) continue;
        const k = nz(r.text);
        if (!cur.subKey.has(k)) { const s = { text: r.text, pages: [], kijun: [] }; cur.subKey.set(k, s); cur.subs.push(s); }
        const s = cur.subKey.get(k);
        let pages = r.pages;
        // （施設基準）の行：施設基準の章の原稿でページを入れ直す（手で直したページはそのまま）
        const isK = r.id && (r.id.k || r.id.st === 'kari');
        if (isK && ctx.sis) {
          const tp = (r.id.tp || []).map(String), cp = r.pages.map(String);
          const hand = JSON.stringify(tp) !== JSON.stringify(cp);
          if (!hand) {
            const parent = cur._row.text;
            const oldp = (String(r.oldp).match(/\d+/) || [null])[0];
            const kin = r.text.includes('記入例');
            let base = oldp ? +oldp : null;
            if (kin) { const sib = W.low.find(x => x.sub && nz(x.text) === nz('（施設基準）') && W.low.indexOf(x) > W.low.indexOf(cur._row) && W.low.indexOf(x) < W.low.indexOf(r)); const m = sib && String(sib.oldp).match(/\d+/); if (m) base = +m[0]; }
            const alts = ctx.aliasK ? ctx.aliasK(parent, r.text, oldp ? +oldp : null) : [];
            const [np] = ctx.sis.find(parent, base, r.id.st === 'new' ? parent : null, alts, kin);
            if (np) { if (String(np) !== cp.join()) notes.push(`施設基準のページを入れ直し：${parent} ${r.text} p.${cp.join('・') || 'なし'} → p.${np}`); pages = [np]; }
            else notes.push(`施設基準の章で見つからない：${parent} ${r.text}（作業Excelのページ p.${cp.join('・') || 'なし'} のまま）`);
          }
        } else if (isK && !ctx.sis) s.kijun.push('施設基準の章のPDFが無いので仮のページのまま');
        for (const p of pages) if (!s.pages.includes(p)) s.pages.push(p);
      }
    }
  }
  const entries = [];
  for (const k of order) {
    const m = byMain.get(k);
    const subs = m.subs;
    if (!m.on && !subs.length) continue;   // どの章でも載せない
    m.pages.sort(pageOrder); subs.forEach(s => s.pages.sort(pageOrder));
    if (!m.yomi) problems.push(`読みが空：${m.text}`);
    for (const s of subs) if (!s.pages.length) problems.push(`ページが空：${m.text} ${s.text}`);
    if (m.on && !m.pages.length && !subs.length) problems.push(`ページが空：${m.text}`);
    entries.push({ text: m.text, yomi: m.yomi, pages: m.on ? m.pages : [], subs: subs.map(s => ({ text: s.text, pages: s.pages, memo: s.kijun.join('') })), chapters: [...m.chapters], pending: m.pending });
  }
  entries.sort((a, b) => cmpYomi(a.yomi, b.yomi) || a.text.localeCompare(b.text, 'ja'));
  return { entries, problems, notes, undecided, undecidedRows };
}

export async function writeKanmatsu(ExcelJS, M, trial) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('索引');
  ws.getColumn(1).width = 60; ws.getColumn(2).width = 14; if (trial) ws.getColumn(3).width = 50;
  ws.addRow(trial ? ['項目', 'ページ', '（試し出し）メモ'] : ['項目', 'ページ']); ws.getRow(1).font = { bold: true, name: '游ゴシック' };
  let head = null;
  for (const e of M.entries) {
    const h = headOf(e.yomi);
    if (h !== head) { head = h; const r = ws.addRow([`【${h}】`]); r.font = { bold: true, name: '游ゴシック' }; }
    ws.addRow(trial ? [e.text, pagesText(e.pages), e.pending.length ? '未決の章あり：' + e.pending.join('・') : null] : [e.text, pagesText(e.pages)]);
    for (const s of e.subs) ws.addRow(trial ? [s.text, pagesText(s.pages), s.memo || null] : [s.text, pagesText(s.pages)]);
  }
  ws.eachRow(r => r.eachCell(c => { c.font = { ...(c.font || {}), name: '游ゴシック' }; }));
  if (trial) {
    const n = wb.addWorksheet('確認すること');
    n.getColumn(1).width = 110;
    n.addRow(['試し出し（社内確認用）です。業者には渡さないでください。']).font = { bold: true };
    for (const t of [...M.missing.map(x => '入っていない章：' + x), ...M.problems, ...M.notes, ...M.undecidedRows.slice(0, 500).map(x => '未決：' + x)]) n.addRow([t]);
  }
  return await wb.xlsx.writeBuffer();
}

// ---- 引き継ぎファイル（次の改定で「設定」に入れる）：巻末の語と読み、載せないと決めた語
export async function writeHandover(ExcelJS, M, works, edition) {
  const wb = new ExcelJS.Workbook();
  const a = wb.addWorksheet('読み');
  a.getColumn(1).width = 60; a.getColumn(2).width = 40;
  a.addRow(['語', '読み']); a.getRow(1).font = { bold: true };
  const seen = new Set();
  for (const e of M.entries) {
    a.addRow([e.text, e.yomi]); seen.add(nz(e.text));
    for (const s of e.subs) a.addRow([s.text, e.yomi]);
  }
  // 載せなかった語の読みも残す（次の改定で読みの推定に使う）
  for (const W of works) for (const r of W.low) if (!r.sub && r.yomi && !seen.has(nz(r.text))) { a.addRow([r.text, r.yomi]); seen.add(nz(r.text)); }
  const b = wb.addWorksheet('載せない');
  b.getColumn(1).width = 24; b.getColumn(2).width = 8; b.getColumn(3).width = 60; b.getColumn(4).width = 50;
  b.addRow(['章', '区分', '語', '打合せメモ']); b.getRow(1).font = { bold: true };
  for (const W of works) for (const r of [...W.up, ...W.low]) if (r.dec === '載せない' && (r.stLabel || '').startsWith('＋')) b.addRow([W.meta.short || '', r.sec, r.text, r.memo || null]);
  const info = wb.addWorksheet('_info', { state: 'veryHidden' });
  [['tool', '手引き索引ツール'], ['kind', '引き継ぎ'], ['edition', edition || '']].forEach((x, i) => { info.getCell(i + 1, 1).value = x[0]; info.getCell(i + 1, 2).value = x[1]; });
  return await wb.xlsx.writeBuffer();
}
