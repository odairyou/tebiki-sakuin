// 業者に渡すデータ：扉ページ（章ごと）、巻末索引（全体）、引き継ぎファイル
import { nz } from './pdftext.js';
import { pagesText, cellText } from './workbook.js';
import { hira } from './yomi.js';

const isUndecided = r => !r.dec || r.dec === '役員確認';   // 役員確認は返事待ち＝まだ決まっていない扱い
const isOn = r => r.dec === '載せる';
const pageOrder = (a, b) => (typeof a === 'number' && typeof b === 'number') ? a - b : (typeof a === 'number' ? -1 : (typeof b === 'number' ? 1 : String(a).localeCompare(String(b))));

// ---- 扉ページ
export function checkTobira(W) {
  const undecided = W.up.filter(r => !r.dec);
  const asking = W.up.filter(r => r.dec === '役員確認');
  const noPage = W.up.filter(r => isOn(r) && !r.pages.length);
  return { undecided, asking, noPage, ok: !undecided.length && !asking.length && !noPage.length };
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
      if (isUndecided(r)) { undecided++; undecidedRows.push(`${ch}：${r.text}` + (r.dec === '役員確認' ? '（役員確認）' : '')); }
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
  // 打合せメモのうち【Claude案】の書き込みは引き継がない（人の判断の記録だけ残す）
  const ownMemo = m => String(m || '').split(/(?=【Claude)/).filter(x => !x.startsWith('【Claude')).join('').trim() || null;
  for (const W of works) for (const r of [...W.up, ...W.low]) if (r.dec === '載せない' && (r.stLabel || '').startsWith('＋')) b.addRow([W.meta.short || '', r.sec, r.text, ownMemo(r.memo)]);
  const info = wb.addWorksheet('_info', { state: 'veryHidden' });
  [['tool', '手引き索引ツール'], ['kind', '引き継ぎ'], ['edition', edition || '']].forEach((x, i) => { info.getCell(i + 1, 1).value = x[0]; info.getCell(i + 1, 2).value = x[1]; });
  return await wb.xlsx.writeBuffer();
}

// ---- 役員確認用（章ごと）：「載せる」「役員確認」の行だけ。列は ★・語・今回P・前回P・事務局より
export function checkOfficer(W) {
  return { undecided: [...W.up, ...W.low].filter(r => !r.dec) };
}
// 役員確認の行に出す決まった文と□の選択肢。戻り値 {mark, lines:[説明…], choice, memo}
export function officerMsg(r, sec) {
  const lab = r.stLabel || '', note = r.note || '';
  const ownMemo = String(r.memo || '').split(/(?=【Claude)/).filter(x => !x.startsWith('【Claude')).join('').trim();
  const old = String(r.oldp || '').trim();
  const put = sec === '扉' ? '扉に載せる' : '索引に載せる';
  if (lab.startsWith('×')) {
    const lines = [/\d/.test(old) ? `前回 p.${old} に掲載。今回の原稿では見当たりません。` : '今回の原稿では見当たりません。'];
    let m = note.match(/よく似た「(.+?)」が p\.(\d+) にあり/) || note.match(/新原稿の「(.+?)」に名称変更かも/);
    if (m) {
      lines.push(`今回は「${m[1]}」${m[2] ? '（p.' + m[2] + '）' : ''}に変わった可能性があります。`);
      return { mark: '無', lines, choice: '□ 外す　　□ 新しい名前で残す　　□ このまま残す（p.＿＿）', memo: ownMemo };
    }
    return { mark: '無', lines, choice: '□ 外す　　□ 残す（該当ページ p.＿＿）', memo: ownMemo };
  }
  const xm = note.match(/前回の索引では(「.+?」)の章（p\.([\d・]+)）にある語/);
  if (lab.startsWith('＋') && xm) {
    const here = (r.pages || []).join('・') || '＿＿';
    if (sec === '扉') return { mark: '新', lines: [`この語は、主に${xm[1]}の章で扱っています。この章では関連として出てくるだけのため、扉には載せなくてよいと考えています。`],
      choice: '□ 載せない（このままでよい）　　□ 扉に載せる', memo: ownMemo };
    return { mark: '新', lines: [`巻末索引では、この語に${xm[1]}の章のページ（前回 p.${xm[2]}）を載せています。`,
      `説明している章のページだけを載せる方針のため、この章のp.${here}は足さなくてよいと考えています。`,
      { small: true, text: `足した場合の表示：${r.text} …… ${xm[2]}・${here}` }],
      choice: `□ 足さない（このままでよい）　　□ p.${here} も足す`, memo: ownMemo, after: `※${xm[1]}の章で見当たらない場合は、こちらを足してください` };
  }
  if (lab.startsWith('＋')) {
    let l;
    if (note.includes('大見出し') && note.includes('早見表')) l = '新しく入った早見表です。';
    else if (note.includes('施設基準') && note.includes('枠あり')) l = '今回の原稿で「施設基準 ➡P参照」が付いた項目です。';
    else if (note.includes('左欄の見出し')) l = '今回の原稿で新しく見出しになった語です。';
    else if (note.includes('略称')) { const m = note.match(/「([^」]+)」の略称/); l = '今回の原稿に新しく出てきた略称です。' + (m ? `（近くの見出し：「${m[1]}」）` : ''); }
    else l = '今回の原稿で新しく見出しになった語です。';
    const lines = [l];
    const rn = note.match(/前回の「(.+?)」と同じ/);
    if (rn) lines.push(`前回の「${rn[1]}」の名前が変わった可能性があります。`);
    return { mark: '新', lines, choice: `□ ${put}　　□ 載せない`, memo: ownMemo };
  }
  return { mark: '★', lines: [], choice: `□ ${put}　　□ 載せない`, memo: ownMemo };
}

// 事務局で載せない予定の、ほかの章の語（取り消し線で出す行）の説明。□は付けない
export function strikeMsg(r, sec) {
  const m = String(r.note || '').match(/前回の索引では(「.+?」)の章（p\.([\d・]+)）にある語/);
  const ch = m ? m[1] : 'ほかの', pg = m ? m[2] : '';
  const A = (r.id && r.id.x === 'A');
  if (sec === '扉') return A ? `主に${ch}の章で扱っている語です。前回もこの章に出ていましたが、扉には載せていません。今回もこのままの予定です。`
    : `主に${ch}の章で扱っている語です。この章では関連として出てくるだけのため、扉には載せない予定です。`;
  return A ? `主に${ch}の章で扱っている語です。前回もこの章に出ていましたが、巻末索引には${ch}の章のページ（前回 p.${pg}）だけを載せています。今回もこのままの予定です。`
    : `主に${ch}の章で扱っている語です。巻末索引には${ch}の章のページ（前回 p.${pg}）だけを載せる予定です。`;
}
const isStrike = r => r.dec === '載せない' && r.id && r.id.x;

export async function writeOfficer(ExcelJS, W) {
  const wb = new ExcelJS.Workbook();
  const JP = '游ゴシック', title = W.meta.title || W.meta.short || '';
  const thin = { style: 'thin', color: { argb: 'FF999999' } }, B = { left: thin, right: thin, top: thin, bottom: thin };
  const isNum = p => typeof p === 'number' || /^\d+$/.test(String(p));
  const MARKC = { '無': 'FFC00000', '新': 'FF1F4E79', '★': 'FFB8860B' };
  const PER_LINE = 26;   // 「事務局より」の欄に1行で入るおよその字数
  const nLines = t => String(t).split('\n').reduce((a, x) => a + Math.max(1, Math.ceil([...x].length / PER_LINE)), 0);
  let asking = 0;
  for (const [sec, name, label] of [['扉', '扉ページ', '扉ページ（この章のはじめのページ）'], ['索引', '巻末索引', '巻末索引（この章の分）']]) {
    const rows = (sec === '扉' ? W.up : W.low).filter(r => r.dec === '載せる' || r.dec === '役員確認' || isStrike(r));
    const nAsk = rows.filter(r => r.dec === '役員確認').length; asking += nAsk;
    const ws = wb.addWorksheet(name, { views: [{ showGridLines: false }] });
    [5, 50, 9, 9, 54].forEach((w, i) => ws.getColumn(i + 1).width = w);
    ws.getCell('A1').value = `索引 確認のお願い　${title}　― ${label}`; ws.getCell('A1').font = { name: JP, size: 15, bold: true };
    ws.getCell('A2').value = '全体をご確認ください。ページの誤り・言い回しの直し・足りない語は、該当する箇所に赤で直接書き込んでください。';
    ws.getCell('A3').value = '印（無・新・★）の行は、特にご判断をお願いしたい行です。「事務局より」の□にレ点を入れてください。　提出：　　月　　日（　）まで';
    for (const a of ['A2', 'A3']) ws.getCell(a).font = { name: JP, size: 10 };
    ws.getCell('A4').value = `このシートで判断をお願いする行：${nAsk}件　（無＝前回あって今回見当たらない語　新＝今回新しく出てきた語）` + (rows.some(isStrike) ? '　取り消し線の行は、ほかの章で扱っているため事務局で載せない予定の語です' : '');
    ws.getCell('A4').font = { name: JP, size: 10, bold: true, color: { argb: 'FF7F6000' } };
    ['', '語', '今回P', '前回P', '事務局より'].forEach((h, j) => {
      const c = ws.getCell(5, j + 1); c.value = h || null; c.font = { name: JP, size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF404040' } }; c.alignment = { horizontal: 'center', vertical: 'middle' }; c.border = B;
    });
    ws.getRow(5).height = 20;
    let rn = 6;
    for (const r of rows) {
      const ask = r.dec === '役員確認';
      const sub = /^[（(]/.test(r.text);
      const pg = r.pages.length ? r.pages.join('・') : (ask ? 'なし' : (/\d/.test(r.oldp) ? '他の章' : ''));
      const old = String(r.oldp || '').trim() || '—';
      const M = ask ? officerMsg(r, sec) : null;
      const strike = isStrike(r);
      let msg = null, h = 30;
      if (strike) {
        const t = strikeMsg(r, sec);
        msg = { richText: [{ text: t, font: { name: JP, size: 9, color: { argb: 'FF666666' } } }] };
        h = Math.max(30, nLines(t) * 13 + 8);
      }
      if (M) {
        const parts = [];
        M.lines.forEach(l => typeof l === 'string'
          ? parts.push({ text: l + '\n', font: { name: JP, size: 10, color: { argb: 'FF262626' } } })
          : parts.push({ text: '　' + l.text + '\n', font: { name: JP, size: 9, color: { argb: 'FF666666' } } }));
        parts.push({ text: M.choice, font: { name: JP, size: 11, bold: true, color: { argb: 'FF000000' } } });
        if (M.after) parts.push({ text: '\n' + M.after, font: { name: JP, size: 9, color: { argb: 'FF666666' } } });
        if (M.memo) parts.push({ text: '\n' + M.memo, font: { name: JP, size: 10, color: { argb: 'FF7F6000' } } });
        msg = { richText: parts };
        const all = [...M.lines.map(l => typeof l === 'string' ? l : l.text), M.choice, ...(M.after ? [M.after] : []), ...(M.memo ? [M.memo] : [])].join('\n');
        h = Math.max(36, nLines(all) * 15 + 10);
      }
      const vals = [ask ? M.mark : (strike ? '外' : null), r.text, pg, old, msg];
      vals.forEach((v, j) => {
        const c = ws.getCell(rn, j + 1); c.value = (j === 2 && isNum(v)) ? +v : v; c.border = B;
        if (j !== 4) c.font = strike
          ? { name: JP, size: j === 0 ? 11 : 11, bold: j === 0, strike: j > 0, color: { argb: j === 0 ? 'FF808080' : 'FF808080' } }
          : { name: JP, size: j === 0 ? 12 : 11, bold: (j === 1 && !sub) || j === 0, color: { argb: j === 0 ? (MARKC[v] || 'FF000000') : (j === 2 && !r.pages.length ? 'FF808080' : 'FF000000') } };
        c.alignment = { vertical: 'middle', wrapText: j === 1 || j === 4, horizontal: [0, 2, 3].includes(j) ? 'center' : 'left', indent: j === 1 && sub ? 2 : 0 };
        if (ask) c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } };
      });
      ws.getRow(rn).height = h; rn++;
    }
    if (!rows.length) { ws.getCell(rn, 2).value = '（この章で載せる語はありません）'; ws.getCell(rn, 2).font = { name: JP, size: 11, color: { argb: 'FF808080' } }; }
    ws.pageSetup = { orientation: 'portrait', paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0, printTitlesRow: '5:5',
      margins: { left: 0.5, right: 0.5, top: 0.6, bottom: 0.6, header: 0.3, footer: 0.3 } };
    ws.headerFooter.oddFooter = `&C${title}　${name}　&P / &N`;
  }
  const buf = await wb.xlsx.writeBuffer();
  return { buf, asking, up: W.up.filter(r => r.dec === '載せる' || r.dec === '役員確認').length, low: W.low.filter(r => r.dec === '載せる' || r.dec === '役員確認').length };
}
