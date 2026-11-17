// 画面から呼ぶ処理のまとめ（ブラウザでもNodeでも同じ）
import { openPdf, readAll, nz } from './pdftext.js';
import { readOld } from './old.js';
import { readDraft } from './draft.js';
import { readSisetsu, Sisetsu } from './sisetsu.js';
import { Chapter, makeAliasK } from './chapter.js';
import { matchChapter } from './match.js';
import { Yomi } from './yomi.js';
import { writeWork, readWork } from './workbook.js';
import { readYomiBook } from './settings.js';
import { checkTobira, writeTobira, mergeKanmatsu, writeKanmatsu, writeHandover, checkOfficer, writeOfficer } from './output.js';

export const shortOf = title => String(title || '').replace(/^[0-9０-９]+[．.]\s*/, '').replace(/[\\/?*\[\]:]/g, '') || '章';
const pad = n => String(n).padStart(2, '0');

// 設定：前回の完成品PDF（必須）＋ 読みの入ったExcel（任意）
export async function prepare(env, oldPdf, yomiBook, progress = () => { }) {
  const doc = await openPdf(env.pdfjs, oldPdf.data, env.pdfOpt);
  const pages = await readAll(env.pdfjs, doc, (i, n) => progress(`前回のPDFを読んでいます（${i}/${n}ページ）`));
  const old = readOld(pages);
  if (!old.chapters.length) throw new Error('前回のPDFから扉ページが見つかりませんでした。手引きの完成品のPDFか確かめてください。');
  if (!old.back.length) throw new Error('前回のPDFから巻末索引が見つかりませんでした。手引きの完成品のPDFか確かめてください。');
  let Y = { yomi: {}, disp: {}, dropped: [], rows: 0 };
  if (yomiBook) {
    progress('読みのExcelを読んでいます');
    Y = await readYomiBook(env.ExcelJS, yomiBook.data);
    if (!Y.rows) throw new Error('このExcelから読みが取れませんでした（「語」と「読み」の列が見つからない）。引き継ぎファイルか、前回の索引作業のExcelを入れてください。');
  }
  const ym = (oldPdf.name.match(/(20\d\d)/) || [])[1] || '';
  return {
    created: new Date().toISOString(), oldName: oldPdf.name, yomiName: yomiBook ? yomiBook.name : '', oldEdition: ym,
    old: { chapters: old.chapters, back: old.back, sisetsu: old.sisetsu, oldTitles: old.oldTitles },
    yomi: Y.yomi, disp: Y.disp, dropped: Y.dropped, yomiRows: Y.rows,
  };
}

// 新原稿PDFを読んで、章の原稿か施設基準の章かを見分ける
export async function readNewPdf(env, file, progress = () => { }) {
  const doc = await openPdf(env.pdfjs, file.data, env.pdfOpt);
  const pages = await readAll(env.pdfjs, doc, (i, n) => progress(`${file.name} を読んでいます（${i}/${n}ページ）`));
  return pages;
}

export function classify(settings, pages) {
  const R = readDraft(pages);
  const s = settings.old.sisetsu;
  const isSisetsu = !R.tobiraPage && ((s && R.tab && nz(R.tab) === nz(s.tab)) || (!R.tab && pages.some(P => P.text.includes('施設基準'))));
  if (isSisetsu) return { kind: 'sisetsu', sis: readSisetsu(pages), range: [R.pages[0], R.pages[R.pages.length - 1]] };
  const m = matchChapter(settings.old, R);
  return { kind: 'chapter', R, match: m };
}

// 1つの章の作業Excelを作る
export async function makeWork(env, settings, R, och, opt = {}) {
  const Y = new Yomi(settings.yomi, env.tokenizer);
  const sis = opt.sisetsu ? new Sisetsu(opt.sisetsu, settings.old.oldTitles) : null;
  const C = new Chapter(settings.old, och, R, Y, sis).run();
  const idx = settings.old.chapters.indexOf(och);
  const meta = { idx, title: och.title, short: shortOf(och.title), tab: och.tab, oldTob: och.tob, range: `${C.new_lo}-${C.new_hi}`, created: new Date().toISOString(), edition: settings.oldEdition };
  const res = await writeWork(env.ExcelJS, C, meta, { disp: settings.disp, carry: opt.carry, dropped: new Set(settings.dropped || []) });
  const count = (rows) => rows.reduce((o, r) => { if (!r.dec) o[r.st || 'human'] = (o[r.st || 'human'] || 0) + 1; return o; }, {});
  return { name: `索引作業_${pad(idx)}${meta.short}.xlsx`, buf: res.buf, meta, C, report: res.report, undecided: { up: count(res.up), low: count(res.low) }, rows: res.up.length + res.low.length };
}

export { checkOfficer, writeOfficer, readWork, checkTobira, writeTobira, mergeKanmatsu, writeKanmatsu, writeHandover, makeAliasK, Sisetsu };
