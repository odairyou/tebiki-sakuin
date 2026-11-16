// 読みの入ったExcel（引き継ぎファイル、または前回の索引作業のExcel）の読み取り
import { nz } from './pdftext.js';
import { cellText } from './workbook.js';

export async function readYomiBook(ExcelJS, data) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(data);
  const yomi = {}, disp = {}, dropped = []; let rows = 0;
  for (const ws of wb.worksheets) {
    if (['見方', '_info'].includes(ws.name) || ws.state === 'veryHidden') continue;
    if (ws.name === '載せない') {
      ws.eachRow((r, rn) => { if (rn > 1) { const t = cellText(r.getCell(3)).trim(); if (t) dropped.push(nz(t)); } });
      continue;
    }
    // 見出し行（「語」「読み」）を探す。無ければA列＝語、B列＝読み
    let cT = 1, cY = 2, start = 1;
    for (let rn = 1; rn <= 5; rn++) {
      const row = ws.getRow(rn); let a = 0, b = 0;
      row.eachCell((c, cn) => { const t = nz(cellText(c)); if (t === '語' || t === '項目') a = cn; if (t === '読み' || t === 'よみ' || t === '読み仮名') b = cn; });
      if (a && b) { cT = a; cY = b; start = rn + 1; break; }
    }
    ws.eachRow((r, rn) => {
      if (rn < start) return;
      const a = cellText(r.getCell(cT)), b = cellText(r.getCell(cY));
      if (!a || !b || a.startsWith('【')) return;
      const y = b.replace(/[\s　]+/g, '');
      if (!/[ぁ-んァ-ヶ]/.test(y)) return;
      const k = nz(a); rows++;
      if (!(k in yomi)) yomi[k] = y;
      if (!(k in disp)) disp[k] = a.trim();
      // 外字などの指示が付いた語（「○未請求←【トーキさんへ】…」）は、指示の前の部分でも引けるようにする
      const base = a.replace(/←.*$/, '').replace(/「未」を○で.*$/, '').trim();
      if (base !== a.trim()) {
        for (const b2 of [base, base.replace(/^○/, ''), base.replace(/^○/, '○未')]) {
          const kb = nz(b2); if (!kb) continue;
          if (!(kb in yomi)) yomi[kb] = y;
          if (!(kb in disp)) disp[kb] = a.trim();
        }
      }
    });
  }
  return { yomi, disp, dropped, rows };
}
