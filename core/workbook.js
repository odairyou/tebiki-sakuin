// 作業Excel（上下2段：上＝扉、下＝巻末）の書き出しと読み戻し
import { nz } from './pdftext.js';

export const TOOL_VER = '1.3';
export const LAB = { ok: '○ 特定', chk: '？ 要確認', ng: '× 見つからず', new: '＋ 新規候補', auto: '－ 自動', head: '－ 親見出し', kari: '△ 仮（施設基準）' };
const STC = { '○': '375623', '？': '7F6000', '×': 'A61C28', '＋': '1F4E79', '－': '999999', '△': 'C55A11' };
const SECC = { '扉': ['1F4E79', 'DDEBF7'], '索引': ['375623', 'E2EFDA'] };
const PREVC = { '両方': '375623', '扉のみ': '1F4E79', '巻末のみ': '7F6000', '—': '999999' };
const COLS = ['区分', 'No', '語', '読み', '今回P', '', '', '前回P', '前回の掲載', 'ツールの判定', '今回の掲載\n▼選ぶ', 'ツールのメモ（なぜそう判断したか）', '打合せメモ', 'ID'];
const WID = [6, 5, 44, 30, 6, 6, 6, 9, 11, 13, 11, 60, 26, 4];
const JP = '游ゴシック';
const argb = c => 'FF' + c;
const F = (o = {}) => ({ name: JP, size: o.size || 10, bold: !!o.bold, italic: !!o.italic, strike: !!o.strike, underline: o.underline, color: { argb: argb(o.color || '000000') } });
const fill = c => ({ type: 'pattern', pattern: 'solid', fgColor: { argb: argb(c) } });
const thin = { style: 'thin', color: { argb: 'FFD9D9D9' } }, thick = { style: 'medium', color: { argb: 'FF808080' } };
export const pagesText = ps => (ps || []).join('・');

// 表示用の語：前回の読みファイルにある元の書き方（全角スペース・外字の指示など）があればそれを使う
export function displayText(text, disp) {
  if (!disp) return text;
  return disp[nz(text)] || text;
}

// C: Chapter.run() の結果、meta: {idx, title, short, tab, oldTob, range, edition}, ctx: {disp, carry}
export async function writeWork(ExcelJS, C, meta, ctx = {}) {
  const disp = ctx.disp || {};
  const mk = (r, sec) => ({
    sec, no: r.no, sub: r.sub, text: r.new ? r.text : displayText(r.text, disp),
    yomi: sec === '索引' && !r.sub ? (r.yomi || '') : '', yflag: sec === '索引' ? (r.yflag || []) : [],
    pages: [...r.newp], oldp: r.oldp, prev: r.prev, st: r.st, note: r.note || '', first: r.first, isNew: r.new,
    dec: ['ok', 'head', 'auto', 'kari'].includes(r.st) ? '載せる' : '', memo: '',
    id: { st: r.st, tp: [...r.newp], op: r.oldp, k: r.kijun ? 1 : 0, other: r.other || [] },
  });
  let up = C.TOP.map(r => mk(r, '扉')), low = C.LOW.map(r => mk(r, '索引'));
  for (const r of [...up, ...low]) {
    if (r.id.other.length && !r.note.includes('他の章')) r.note += (r.note ? '。' : '') + 'p.' + pagesText(r.id.other) + ' は他の章';
    if (r.sec === '扉') r.note = r.note.replace(/。?読みの「.*?（要確認）/, '');
  }
  if (ctx.dropped) markDropped([...up, ...low], ctx.dropped);
  let report = null;
  if (ctx.carry) ({ up, low, report } = carryOver(up, low, ctx.carry));

  const wb = new ExcelJS.Workbook();
  wb.creator = '手引き索引ツール'; wb.calcProperties.fullCalcOnLoad = true;
  const g = wb.addWorksheet('見方', { properties: { tabColor: { argb: 'FF404040' } }, views: [{ showGridLines: false }] });
  const SH = meta.short;
  const ws = wb.addWorksheet(SH, { properties: { tabColor: { argb: 'FFBF8F00' } }, views: [{ state: 'frozen', xSplit: 3, ySplit: 2, activeCell: 'D4' }] });
  COLS.forEach((h, j) => {
    const c = ws.getCell(2, j + 1); c.value = h || null; c.font = F({ bold: true, color: 'FFFFFF' }); c.fill = fill('404040');
    c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    ws.getColumn(j + 1).width = WID[j];
  });
  ws.getColumn(14).hidden = true;
  ws.mergeCells('E2:G2'); ws.getRow(2).height = 32;
  const r0 = 4, lowbanner = r0 + up.length + 1, lowstart = lowbanner + 1, lowend = lowstart + low.length - 1;
  const q = SH.replace(/'/g, "''");
  // シート内へのリンクはHYPERLINK関数で（Excelで開けないファイルにならないように）
  ws.getCell('A1').value = { formula: `HYPERLINK("#'${q}'!A${r0 - 1}","▲ 扉ページへ")`, result: '▲ 扉ページへ' };
  ws.getCell('A1').font = F({ bold: true, color: '1F4E79', underline: true });
  ws.getCell('D1').value = { formula: `HYPERLINK("#'${q}'!A${lowbanner}","▼ 索引（巻末）へ")`, result: '▼ 索引（巻末）へ' };
  ws.getCell('D1').font = F({ bold: true, color: '375623', underline: true });
  ws.mergeCells('A1:C1');
  const banner = (rr, t, sec) => {
    for (let j = 1; j <= COLS.length - 1; j++) ws.getCell(rr, j).fill = fill(SECC[sec][0]);
    const c = ws.getCell(rr, 1); c.value = t; c.font = F({ bold: true, size: 12, color: 'FFFFFF' }); ws.getRow(rr).height = 22;
  };
  banner(r0 - 1, `■ 扉ページ（${meta.title}・ページ順）　― この段の色帯は青`, '扉');
  banner(lowbanner, '■ 索引（巻末に入れる語。扉と同じNoの順）　― この段の色帯は緑', '索引');
  const put = (xr, r) => {
    const sec = r.sec; const st = r.st ? LAB[r.st] : '';
    const newno = !r.sub && (sec === '扉' || r.first) && r.st;
    const p = r.pages || [];
    const vals = [r.st ? sec : null, r.no ?? null, r.text, (sec === '索引' && !r.sub) ? (r.yomi || null) : null, p[0] ?? null, p[1] ?? null, p[2] ?? null,
      r.st ? (pagesText(r.oldp) || '—') : null, r.prev || null, st || null, r.dec || null, r.note || null, r.memo || null, r.st ? JSON.stringify(r.id) : null];
    vals.forEach((v, j) => {
      const c = ws.getCell(xr, j + 1); c.value = (typeof v === 'string' && /^\d+$/.test(v) && j >= 4 && j <= 6) ? +v : v; c.font = F();
      c.border = { left: thin, right: thin, bottom: thin, top: newno && j > 0 ? thick : thin };
      c.alignment = { vertical: 'middle', horizontal: [2, 3, 11, 12].includes(j) ? 'left' : 'center', wrapText: j === 11 || j === 12, indent: (j === 2 && r.sub) ? 2 : 0 };
    });
    ws.getCell(xr, 1).fill = fill(SECC[sec][1]); ws.getCell(xr, 1).font = F({ bold: true, size: 9, color: SECC[sec][0] });
    ws.getCell(xr, 2).font = F({ bold: !!newno, size: 9, color: newno ? '000000' : 'BFBFBF' });
    ws.getCell(xr, 3).font = F({ bold: !!newno && sec === '扉' });
    for (let k = 5; k <= 7; k++) ws.getCell(xr, k).font = F({ bold: true });
    ws.getCell(xr, 8).font = F({ color: '808080' }); if (r.prev) ws.getCell(xr, 9).font = F({ bold: true, color: PREVC[r.prev] || '000000' });
    if (st) ws.getCell(xr, 10).font = F({ bold: true, color: STC[st[0]] });
    ws.getCell(xr, 11).fill = fill('FFFFE0'); ws.getCell(xr, 11).font = F({ bold: true }); ws.getCell(xr, 13).fill = fill('FFFFE0');
    ws.getCell(xr, 12).font = F({ size: 9, color: '595959' });
    ws.getCell(xr, 14).font = F({ size: 6, color: 'FFFFFF' }); ws.getCell(xr, 14).alignment = { horizontal: 'left', vertical: 'middle' };
    if (sec === '索引' && r.yflag && r.yflag.length) ws.getCell(xr, 4).font = F({ color: 'C55A11', bold: true });
    if (r.changed) for (let k = 5; k <= 7; k++) ws.getCell(xr, k).fill = fill('FFD966');
    if ((r.note || '').length > 40) ws.getRow(xr).height = 28;
  };
  up.forEach((r, i) => put(r0 + i, r));
  low.forEach((r, i) => put(lowstart + i, r));
  for (let i = r0; i <= lowbanner - 2; i++) ws.getRow(i).outlineLevel = 1;
  for (let i = lowstart; i <= lowend; i++) ws.getRow(i).outlineLevel = 1;
  const END = lowend + 200;
  ws.dataValidations.add(`K${r0}:K${END}`, { type: 'list', allowBlank: true, formulae: ['"載せる,載せない"'], showErrorMessage: true, errorTitle: '今回の掲載', error: '「載せる」か「載せない」を選んでください' });
  let prio = 1;
  const cf = (ref, formula, style, stop = false) => ws.addConditionalFormatting({ ref, rules: [{ type: 'expression', formulae: [formula], style, stopIfTrue: stop, priority: prio++ }] });
  cf(`B${r0}:M${END}`, `$K${r0}="載せない"`, { font: { color: { argb: 'FFA6A6A6' }, strike: true }, fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFF2F2F2' } } }, true);
  for (const [ch, col] of [['？', 'FFF2CC'], ['×', 'F8D7DA'], ['＋', 'DDEBF7']])
    cf(`B${r0}:J${END}`, `AND($K${r0}="",LEFT($J${r0},1)="${ch}")`, { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: argb(col) } } });
  cf(`B${r0}:K${END}`, `AND($C${r0}<>"",$J${r0}="",$K${r0}="")`, { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFCE4D6' } } });
  ws.autoFilter = `A2:M${lowend}`;
  ws.pageSetup = { orientation: 'landscape', paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0, printTitlesRow: '2:2' };

  // ---- 見方
  [['A', 3], ['B', 16], ['C', 9], ['D', 9], ['E', 82]].forEach(([c, w]) => g.getColumn(c).width = w);
  g.getCell('B2').value = `索引 作業シート　${meta.title}`; g.getCell('B2').font = F({ size: 15, bold: true });
  g.getCell('B3').value = `上の段＝扉ページ（青）、下の段＝巻末の索引（緑）。前回の索引に、新原稿でのページを入れてあります。`; g.getCell('B3').font = F({ color: '595959' });
  g.getCell('B5').value = 'まだ決めていない行'; g.getCell('B5').font = F({ size: 12, bold: true });
  ['ツールの判定', '扉', '索引', '意味・やること'].forEach((h, j) => { const c = g.getCell(6, 2 + j); c.value = h; c.font = F({ bold: true, color: 'FFFFFF' }); c.fill = fill('404040'); c.alignment = { horizontal: 'center' }; });
  const leg = [['？ 要確認', 'FFF2CC', '語はあるがページに自信なし。どこを指すか決めて、合っていれば「載せる」、違えばページを直して「載せる」。'],
    ['× 見つからず', 'F8D7DA', '新原稿に見つからない（削除・名前の変更・図の中など）。載せるならページを手で入れて「載せる」、やめるなら「載せない」。'],
    ['＋ 新規候補', 'DDEBF7', '新原稿の見出し・略称で、前回の索引に無い語。載せるなら「載せる」、不要なら「載せない」。橙色の読みは自動推定の部分あり。'],
    ['（空欄）', 'FCE4D6', '人が追加した行で、まだ「今回の掲載」を選んでいないもの。']];
  const sh = `'${q}'`;
  leg.forEach(([a, col, t], i) => {
    const rr = 7 + i; const c = g.getCell(rr, 2); c.value = a; c.fill = fill(col); c.font = F({ bold: true }); c.alignment = { horizontal: 'center', vertical: 'middle' };
    ['扉', '索引'].forEach((sec, k) => {
      let f;
      if (a === '（空欄）') f = k === 1 ? `COUNTIFS(${sh}!$A:$A,"",${sh}!$C:$C,"<>",${sh}!$J:$J,"",${sh}!$K:$K,"")-COUNTIFS(${sh}!$A:$A,"",${sh}!$C:$C,"語",${sh}!$K:$K,"")` : null;
      else f = `COUNTIFS(${sh}!$A:$A,"${sec}",${sh}!$J:$J,B${rr},${sh}!$K:$K,"")`;
      const n = g.getCell(rr, 3 + k); n.value = f ? { formula: f } : '—'; n.font = F({ bold: true, size: 12 }); n.alignment = { horizontal: 'center', vertical: 'middle' };
    });
    const d = g.getCell(rr, 5); d.value = t; d.font = F(); d.alignment = { wrapText: true, vertical: 'middle' }; g.getRow(rr).height = 30;
  });
  g.getCell('B12').value = '「今回の掲載」を選ぶとその行の色が消えます。件数がすべて0になれば、この章は完了です。'; g.getCell('B12').font = F({ bold: true, color: 'BF8F00' });
  const secs = [['上下の見方', ['No は扉と索引で共通です。扉のNo.3の語は、索引でもNo.3の行にまとまっています。Noが変わるところに太線が入ります。',
      'シート最上段の「▲扉ページへ」「▼索引へ」で、それぞれの段の先頭へ飛べます。左端の「－」で段ごと折りたためます。',
      '「前回の掲載」は前回の手引きでどこに載っていたか（両方／扉のみ／巻末のみ／—＝前回なし）。',
      '扉には載せず巻末だけに載せるときは、上の段を「載せない」、下の段を「載せる」にします。',
      'ページを空けておきたい（あとで業者に入れてもらう）ときは、今回Pに「●●」と書きます。']],
    ['行を足すとき', ['空の行を挿入して、語（索引の段は読みも）・今回Pを書き、「今回の掲載」を選びます。No・ツールの判定は空欄でOK。',
      '既存の行をコピーして作った場合は、「ツールの判定」「ツールのメモ」を消してください。',
      '索引の段の副項目（カッコの行）は「すぐ上の主項目の子」になります。主項目は前の項目の副項目の途中に挟まないでください。']],
    ['（施設基準）の行', ['「△ 仮（施設基準）」は、施設基準の章の原稿から見つけたページを仮に入れたものです。業者に渡すデータを作るとき、施設基準の章の最新の原稿でもう一度入れ直します（手でページを直した行は、そのページを使います）。',
      '施設基準の章で見つからなかったものは「× 見つからず」になっています。']]];
  let rr = 14;
  for (const [t0, items] of secs) {
    g.getCell(rr, 2).value = t0; g.getCell(rr, 2).font = F({ size: 12, bold: true }); rr++;
    for (const t of items) {
      g.getCell(rr, 2).value = '・'; g.getCell(rr, 2).alignment = { horizontal: 'right', vertical: 'top' };
      const c = g.getCell(rr, 3); c.value = t; c.font = F(); c.alignment = { wrapText: true, vertical: 'top' };
      g.mergeCells(rr, 3, rr, 5); g.getRow(rr).height = t.length > 52 ? 30 : 16; rr++;
    }
    rr++;
  }
  g.pageSetup = { orientation: 'landscape', paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0 };
  // ---- 隠しシート：ツールが読み戻すための情報
  const info = wb.addWorksheet('_info', { state: 'veryHidden' });
  const kv = { tool: '手引き索引ツール', ver: TOOL_VER, kind: '作業', idx: meta.idx, title: meta.title, short: meta.short, tab: meta.tab, oldTob: meta.oldTob, range: meta.range, created: meta.created || '', sheet: SH, edition: meta.edition || '' };
  Object.entries(kv).forEach(([k, v], i) => { info.getCell(i + 1, 1).value = k; info.getCell(i + 1, 2).value = v == null ? '' : String(v); });
  wb.views = [{ activeTab: 1 }];
  const buf = await wb.xlsx.writeBuffer();
  return { buf, up, low, report };
}

// 作業Excelを読み戻す
export async function readWork(ExcelJS, data) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(data);
  const meta = {};
  const info = wb.getWorksheet('_info');
  if (info) info.eachRow(r => { meta[String(r.getCell(1).value)] = r.getCell(2).text; });
  let ws = meta.sheet ? wb.getWorksheet(meta.sheet) : null;
  if (!ws) ws = wb.worksheets.find(w => cellText(w.getCell(2, 3)) === '語' && cellText(w.getCell(2, 1)) === '区分');
  if (!ws) throw new Error('作業Excelの形ではありません（「区分」「語」の見出しが見つからない）');
  const up = [], low = []; let cur = null;
  ws.eachRow({ includeEmpty: false }, (row, rn) => {
    if (rn < 3) return;
    const a = cellText(row.getCell(1));
    if (a.startsWith('■ 扉')) { cur = up; return; }
    if (a.startsWith('■ 索引')) { cur = low; return; }
    if (!cur) return;
    const text = cellText(row.getCell(3)).trim();
    if (!text) return;
    const pages = [5, 6, 7].map(k => cellText(row.getCell(k)).trim()).filter(Boolean).flatMap(v => v.split(/[・,、，\s]+/)).filter(Boolean)
      .map(v => /^\d+$/.test(nz(v)) ? +nz(v) : v);
    let id = null; try { const t = cellText(row.getCell(14)); if (t) id = JSON.parse(t); } catch (e) { }
    cur.push({ row: rn, sec: cur === up ? '扉' : '索引', no: cellText(row.getCell(2)), text, sub: /^[（(]/.test(text), yomi: cellText(row.getCell(4)).trim(),
      pages, oldp: cellText(row.getCell(8)), prev: cellText(row.getCell(9)), stLabel: cellText(row.getCell(10)), dec: cellText(row.getCell(11)).trim(),
      note: cellText(row.getCell(12)), memo: cellText(row.getCell(13)), id, human: !cellText(row.getCell(10)) });
  });
  return { meta, up, low, sheet: ws.name };
}
export function cellText(c) {
  const v = c.value;
  if (v == null) return '';
  if (typeof v === 'object') {
    if (v.richText) return v.richText.map(x => x.text).join('');
    if ('result' in v) return v.result == null ? '' : String(v.result);
    if (v.text != null) return String(v.text);
    if (v instanceof Date) return v.toISOString();
  }
  return String(v);
}

// 前回「載せない」にした語（引き継ぎファイル）を新規候補のメモに出す
function markDropped(rows, dropped) {
  for (const r of rows) if (r.st === 'new' && dropped.has(nz(r.text))) r.note += '。前回の改定で「載せない」と判断した語';
}

// 原稿が変わったときの作り直し：前に出したExcelの判定・打合せメモ・人が足した行を引き継ぐ
export function carryOver(up, low, prev) {
  const report = { kept: 0, changed: 0, human: 0, lost: [] };
  const sec = (rows, prevRows) => {
    const key = r => (r.sub ? 'S' : 'M') + nz(r.text);
    const pool = new Map();
    prevRows.forEach((p, i) => { if (p.human) return; const k = key(p); if (!pool.has(k)) pool.set(k, []); pool.get(k).push(i); });
    const used = new Set();
    for (const r of rows) {
      const lst = pool.get(key(r)); if (!lst || !lst.length) continue;
      const i = lst.shift(); used.add(i); const p = prevRows[i];
      r.dec = p.dec; r.memo = p.memo; report.kept++;
      if (r.sec === '索引' && !r.sub && p.yomi && p.yomi !== r.yomi) { r.yomi = p.yomi; r.yflag = []; }
      const toolPrev = p.id ? (p.id.tp || []) : [];
      const same = (a, b) => JSON.stringify(a.map(String)) === JSON.stringify(b.map(String));
      const handEdited = p.id && !same(p.pages, toolPrev);
      if (handEdited) {
        r.pages = p.pages; r.note = (r.note ? r.note + '。' : '') + '前のExcelで手で入れたページを残した（新原稿での推定は p.' + (pagesText(r.id.tp) || 'なし') + '）';
      } else if (!same(p.pages, r.pages)) {
        r.changed = true; report.changed++;
        r.note = `ページが変わった（前のExcel p.${pagesText(p.pages) || 'なし'} → p.${pagesText(r.pages) || 'なし'}）` + (r.note ? '。' + r.note : '');
      }
    }
    // 人が足した行：前のExcelですぐ上にあった行の後ろに入れる
    const out = [...rows];
    prevRows.forEach((p, i) => {
      if (!p.human) return;
      report.human++;
      let anchor = -1;
      for (let j = i - 1; j >= 0; j--) {
        const a = prevRows[j];
        const k = out.findIndex(r => (r.sub ? 'S' : 'M') + nz(r.text) === (a.sub ? 'S' : 'M') + nz(a.text));
        if (k >= 0) { anchor = k; break; }
      }
      const r = { sec: p.sec, no: null, sub: p.sub, text: p.text, yomi: p.yomi, pages: p.pages, oldp: [], prev: '', st: '', note: p.note, dec: p.dec, memo: p.memo, id: null, human: true };
      out.splice(anchor + 1, 0, r);
    });
    prevRows.forEach((p, i) => { if (!p.human && !used.has(i) && p.dec) report.lost.push(`${p.sec}：${p.text}（${p.dec}）`); });
    return out;
  };
  return { up: sec(up, prev.up), low: sec(low, prev.low), report };
}
