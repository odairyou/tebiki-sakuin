// 手引き索引ツール：画面の動き
import * as P from './core/pipeline.js';
import { TOOL_VER } from './core/workbook.js';
import { makeZip } from './core/zip.js';

const $ = id => document.getElementById(id);
const esc = s => String(s ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const fmt = n => Number(n).toLocaleString('ja-JP');
const XLSX = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
$('ver').textContent = 'ver ' + TOOL_VER;

// ---------- ライブラリ
let envP = null;
function getEnv() {
  if (!envP) envP = (async () => {
    const pdfjs = await import('./pdfjs/pdf.min.mjs');
    pdfjs.GlobalWorkerOptions.workerSrc = new URL('pdfjs/pdf.worker.min.mjs', location.href).href;
    return {
      pdfjs, ExcelJS: window.ExcelJS, tokenizer: null,
      pdfOpt: { cMapUrl: new URL('pdfjs/cmaps/', location.href).href, standardFontDataUrl: new URL('pdfjs/standard_fonts/', location.href).href },
    };
  })();
  return envP;
}
let tkP = null;
function getTokenizer() {   // 読みの推定に使う辞書（はじめの1回だけ少し時間がかかる）
  if (!tkP) tkP = new Promise(res => {
    try { window.kuromoji.builder({ dicPath: 'kuromoji/dict/' }).build((err, tk) => res(err ? null : tk)); }
    catch (e) { res(null); }
  });
  return tkP;
}

// ---------- このパソコンに覚えておく（IndexedDB）
const DB = {
  open() {
    return new Promise((res, rej) => {
      const r = indexedDB.open('tebiki-sakuin', 1);
      r.onupgradeneeded = () => r.result.createObjectStore('kv');
      r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error);
    });
  },
  async get(k) { const db = await this.open(); return new Promise((res, rej) => { const q = db.transaction('kv').objectStore('kv').get(k); q.onsuccess = () => res(q.result); q.onerror = () => rej(q.error); }); },
  async set(k, v) { const db = await this.open(); return new Promise((res, rej) => { const t = db.transaction('kv', 'readwrite'); t.objectStore('kv').put(v, k); t.oncomplete = () => res(); t.onerror = () => rej(t.error); }); },
  async del(k) { const db = await this.open(); return new Promise((res, rej) => { const t = db.transaction('kv', 'readwrite'); t.objectStore('kv').delete(k); t.oncomplete = () => res(); t.onerror = () => rej(t.error); }); },
};
let SETTINGS = null, SIS = null;

// ---------- 共通の部品
async function readFile(f) { return { name: f.name, data: new Uint8Array(await f.arrayBuffer()), lastModified: f.lastModified }; }
function download(name, buf, type = XLSX) {
  const url = URL.createObjectURL(new Blob([buf], { type }));
  const a = document.createElement('a'); a.href = url; a.download = name; document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 60000);
}
function dlLink(name, buf, label = 'ダウンロード') {
  const a = document.createElement('a'); a.className = 'dl'; a.href = '#';
  a.innerHTML = '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M12 4v12"></path><path d="M7 11l5 5 5-5"></path><path d="M4 20h16"></path></svg>' + esc(label);
  a.addEventListener('click', e => { e.preventDefault(); download(name, buf); });
  return a;
}
function progress(id, msg) { const el = $(id); if (msg == null) { el.hidden = true; return; } el.hidden = false; el.textContent = msg; }
const tick = () => new Promise(r => setTimeout(r, 0));
function setupDrop(zone, onFiles) {
  const input = $(zone.dataset.input);
  zone.tabIndex = 0; zone.setAttribute('role', 'button');
  zone.addEventListener('click', () => input.click());
  zone.addEventListener('keydown', e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); input.click(); } });
  input.addEventListener('change', () => { if (input.files.length) onFiles([...input.files]); input.value = ''; });
  zone.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('over'); });
  zone.addEventListener('dragleave', () => zone.classList.remove('over'));
  zone.addEventListener('drop', e => { e.preventDefault(); zone.classList.remove('over'); if (e.dataTransfer.files.length) onFiles([...e.dataTransfer.files]); });
}
const isPdf = f => /\.pdf$/i.test(f.name), isXlsx = f => /\.xlsx$/i.test(f.name);
function box(kind, title, lines = [], extra = null) {
  const d = document.createElement('div'); d.className = 'box ' + kind;
  d.innerHTML = `<span class="t">${esc(title)}</span>` + lines.map(l => `<span class="sub">${l}</span>`).join('');
  if (extra) d.append(extra);
  return d;
}
function listDetails(summary, items, max = 300) {
  const d = document.createElement('details'); d.className = 'list';
  d.innerHTML = `<summary>${esc(summary)}</summary><ul>${items.slice(0, max).map(x => `<li>${esc(x)}</li>`).join('')}${items.length > max ? `<li>ほか ${items.length - max} 件</li>` : ''}</ul>`;
  return d;
}
function errText(e) { console.error(e); return (e && e.message) ? e.message : String(e); }

// ---------- 画面の切り替え
function route() {
  const h = (location.hash || '#home').slice(1);
  const id = ['home', 'prepare', 'chapter', 'output'].includes(h) ? h : 'home';
  document.querySelectorAll('.screen').forEach(s => s.hidden = s.id !== 's-' + id);
  $('ch-need').hidden = !!SETTINGS; $('out-need').hidden = !!SETTINGS;
  window.scrollTo(0, 0);
}
window.addEventListener('hashchange', route);

function settingsSummary(S) {
  const d = new Date(S.created);
  return [`前回＝${S.oldEdition ? S.oldEdition + '年版' : '（年版不明）'}（${esc(S.oldName)}）`,
    `扉 ${S.old.chapters.length}章・巻末 ${fmt(S.old.back.length)}語・読み ${S.yomiRows ? fmt(Object.keys(S.yomi).length) + '語（' + esc(S.yomiName) + '）' : 'なし（推定になります）'}`,
    `取り込んだ日：${d.toLocaleDateString('ja-JP')}`];
}
function renderSettings() {
  const st = $('home-status');
  if (SETTINGS) { st.textContent = `取り込み済み：前回＝${SETTINGS.oldEdition ? SETTINGS.oldEdition + '年版' : SETTINGS.oldName}`; st.classList.remove('no'); }
  else { st.textContent = 'まだ取り込んでいません'; st.classList.add('no'); }
  const c = $('prep-current');
  if (SETTINGS) { c.hidden = false; c.innerHTML = '<b>取り込み済みです。</b>' + settingsSummary(SETTINGS).map(x => `<span>${x}</span>`).join('') + '<span class="sub">入れ直すときは、下にファイルを入れて「取り込む」を押してください。</span>'; }
  else c.hidden = true;
  renderSis();
}
function renderSis() {
  const el = $('ch-sis');
  el.textContent = SIS ? `施設基準の章：${SIS.name}（${new Date(SIS.date).toLocaleDateString('ja-JP')}に入れたもの）を使います。新しい原稿があれば一緒に入れてください。` : '施設基準の章のPDFはまだ入っていません。入れると（施設基準）の行のページを仮に入れます。';
}

// ---------- 設定
const prep = { pdf: null, xls: null };
setupDrop(document.querySelector('[data-input="prep-pdf"]'), async fs => {
  const f = fs.find(isPdf); if (!f) { alert('PDFを選んでください'); return; }
  prep.pdf = f; $('prep-pdf-name').textContent = f.name; document.querySelector('[data-input="prep-pdf"]').classList.add('done'); $('prep-go').disabled = false;
});
setupDrop(document.querySelector('[data-input="prep-xls"]'), async fs => {
  const f = fs.find(isXlsx); if (!f) { alert('Excel（.xlsx）を選んでください'); return; }
  prep.xls = f; $('prep-xls-name').textContent = f.name; document.querySelector('[data-input="prep-xls"]').classList.add('done');
});
$('prep-go').addEventListener('click', async () => {
  const btn = $('prep-go'); btn.disabled = true;
  try {
    progress('prep-progress', '準備しています');
    const env = await getEnv();
    const S = await P.prepare(env, await readFile(prep.pdf), prep.xls ? await readFile(prep.xls) : null, m => progress('prep-progress', m));
    await DB.set('settings', S); SETTINGS = S;
    progress('prep-progress', null); renderSettings();
    const c = $('prep-current'); c.innerHTML = '<b>取り込みました。</b>' + settingsSummary(S).map(x => `<span>${x}</span>`).join('') + '<a href="#chapter">次へ：確認用データを出す →</a>';
    c.scrollIntoView({ behavior: 'smooth' });
  } catch (e) { progress('prep-progress', null); alert('取り込めませんでした：' + errText(e)); }
  btn.disabled = false;
});

// ---------- 確認用データ
const made = [];   // [{name, buf}]
function chapterOptions(sel) {
  return SETTINGS.old.chapters.map((c, i) => `<option value="${i}"${i === sel ? ' selected' : ''}>${esc(c.title)}</option>`).join('');
}
function undecidedText(u) {
  const n = o => Object.values(o).reduce((a, b) => a + b, 0);
  return `まだ決めていない行：扉 <b>${n(u.up)}</b>行・索引 <b>${n(u.low)}</b>行`;
}
async function runChapter(item, env) {
  const och = SETTINGS.old.chapters[item.idx];
  const W = await P.makeWork(env, SETTINGS, item.R, och, { sisetsu: SIS ? SIS.sis : null, carry: item.carry || null });
  item.out = W;
  const old = made.findIndex(m => m.name === W.name); if (old >= 0) made.splice(old, 1);
  made.push({ name: W.name, buf: W.buf });
  return W;
}
function renderChapterRow(item, env) {
  const W = item.out; const row = item.el || document.createElement('div'); item.el = row;
  row.className = 'row'; row.innerHTML = '';
  const conf = item.match.score >= 0.8;
  const info = [`${esc(item.file)} → <b>${esc(SETTINGS.old.chapters[item.idx].title)}</b>（p.${esc(W.meta.range.replace('-', '〜'))}）`, undecidedText(W.undecided)];
  if (item.carry) { const r = W.report; info.push(`前のExcelから引き継ぎ：判定 ${r.kept}行・人が足した行 ${r.human}行・ページが変わった行 ${r.changed}行` + (r.lost.length ? `・今回の原稿に無くなった行 ${r.lost.length}行` : '')); }
  if (!SIS) info.push('施設基準の章のPDFが無いので、（施設基準）の行のページは空です');
  if (W.C && W.C.xchNoInfo) info.push('<b style="color:#A61C28">「設定」を取り込み直してください。</b>ツールを新しくしたため、取り込み直すと、ほかの章から来た語も候補に出せるようになります');
  row.innerHTML = `<span class="name">${esc(W.name)}</span><span class="info">${info.join('<br>')}</span>`;
  row.append(dlLink(W.name, W.buf));
  const fix = document.createElement('div'); fix.className = 'fix';
  fix.innerHTML = `${conf ? '' : '<b style="color:#A61C28">章の見分けに自信がありません。</b>'}章が違うときは選び直し：<select>${chapterOptions(item.idx)}</select>`;
  fix.querySelector('select').addEventListener('change', async e => {
    item.idx = +e.target.value; item.match.score = 1;
    progress('ch-progress', '作り直しています'); await tick();
    try { await runChapter(item, env); renderChapterRow(item, env); } catch (er) { alert(errText(er)); }
    progress('ch-progress', null);
  });
  row.append(fix);
  if (item.carry && W.report.lost.length) row.append(listDetails('今回の原稿に無くなった行（前のExcelで判定済みだったもの）', W.report.lost));
  return row;
}
setupDrop(document.querySelector('[data-input="ch-files"]'), async files => {
  if (!SETTINGS) { alert('先に「設定」で前回の索引を取り込んでください。'); return; }
  const pdfs = files.filter(isPdf), xlss = files.filter(isXlsx);
  if (!pdfs.length) { alert('新原稿のPDFを入れてください。（Excelだけでは作れません。作り直すときは、PDFと前のExcelを一緒に入れます）'); return; }
  const list = $('ch-list'); $('ch-results').hidden = false;
  try {
    const env = await getEnv();
    progress('ch-progress', '読みの辞書を用意しています（はじめの1回は少し時間がかかります）');
    env.tokenizer = await getTokenizer();
    if (!env.tokenizer) list.prepend(box('warn', '読みの辞書を読み込めませんでした。新しい語の読みは空になります。', ['インターネットにつながっているか、ツールのページを開き直してみてください。']));
    // 前のExcel（作り直し用）
    const prevs = [];
    for (const f of xlss) {
      try { const W = await P.readWork(env.ExcelJS, (await readFile(f)).data); W.file = f.name; prevs.push(W); }
      catch (e) { list.prepend(box('ng', `${f.name}：作業Excelとして読めませんでした`, [esc(errText(e))])); }
    }
    // PDFを読んで、章か施設基準かを見分ける
    const items = [];
    for (const f of pdfs) {
      try {
        const file = await readFile(f);
        const pages = await P.readNewPdf(env, file, m => progress('ch-progress', m));
        const c = P.classify(SETTINGS, pages);
        if (c.kind === 'sisetsu') {
          SIS = { name: f.name, date: new Date().toISOString(), sis: c.sis }; await DB.set('sisetsu', SIS); renderSis();
          list.prepend(box('ok', `${f.name}：施設基準の章として覚えました`, [`p.${c.range[0]}〜${c.range[1]}。（施設基準）の行のページを仮に入れるのに使います。`]));
        } else items.push({ file: f.name, R: c.R, match: c.match, idx: SETTINGS.old.chapters.indexOf(c.match.chapter) });
      } catch (e) { list.prepend(box('ng', `${f.name}：読めませんでした`, [esc(errText(e))])); }
    }
    for (const item of items) {
      const prev = prevs.find(W => String(W.meta.idx) === String(item.idx));
      if (prev) { item.carry = prev; prev.used = true; }
      progress('ch-progress', `${item.file} の作業Excelを作っています`); await tick();
      try { await runChapter(item, env); list.append(renderChapterRow(item, env)); }
      catch (e) { const r = box('ng', `${item.file}：作業Excelを作れませんでした`, [esc(errText(e))]); list.append(r); }
    }
    for (const W of prevs) if (!W.used) list.append(box('warn', `${W.file}：同じ章の新原稿PDFが無いので使いませんでした`));
    const dup = items.map(i => i.idx).filter((x, i, a) => a.indexOf(x) !== i);
    if (dup.length) list.append(box('warn', '同じ章に見分けられたPDFが複数あります。あとに作ったほうで上書きしています。章が違うものは選び直してください。'));
  } catch (e) { alert(errText(e)); }
  progress('ch-progress', null);
});
$('ch-all').addEventListener('click', async () => {
  if (!made.length) return;
  const d = new Date(), ymd = `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, '0')}${String(d.getDate()).padStart(2, '0')}`;
  const zip = await makeZip(window.JSZip, made);
  download(`索引作業_${ymd}.zip`, zip, 'application/zip');
});

// ---------- 役員確認用
setupDrop(document.querySelector('[data-input="of-files"]'), async files => {
  const res = $('of-res'); const env = await getEnv();
  for (const f of files.filter(isXlsx)) {
    try {
      const W = await P.readWork(env.ExcelJS, (await readFile(f)).data);
      const ch = W.meta.title || W.meta.short || f.name;
      const c = P.checkOfficer(W);
      if (c.undecided.length) {
        const b = box('ng', `${ch}：まだ出せません`, [esc(`「今回の掲載」が空の行が${c.undecided.length}行あります。「載せる」「載せない」「役員確認」のどれかを選んでください`)]);
        b.append(listDetails('空の行を見る', c.undecided.map(r => `${r.row}行目　${r.sec}　${r.text}`)));
        res.prepend(b); continue;
      }
      const O = await P.writeOfficer(env.ExcelJS, W);
      const name = `索引確認_${String(W.meta.idx).padStart(2, '0')}${W.meta.short || ''}.xlsx`;
      res.prepend(box('ok', `${ch}：役員確認用を出しました`, [`扉 ${O.up}行・索引 ${O.low}行。うち「役員確認」${O.asking}行`], dlLink(name, O.buf, name)));
    } catch (e) { res.prepend(box('ng', `${f.name}：読めませんでした`, [esc(errText(e))])); }
  }
});

// ---------- 業者に渡すデータ
// 最後に入れた扉のExcel。チェックを切り替えたら、入れ直さずにその場で出し直す
let TOB_LAST = [];
async function renderTobira(items) {
  const res = $('out-tob-res'); const env = await getEnv();
  const sort = $('out-tob-sort').checked;
  for (const it of items) {
    if (it.el) it.el.remove();
    const { f, W } = it; const ch = W.meta.title || W.meta.short || f.name;
    const c = P.checkTobira(W);
    if (c.ok) {
      const name = `業者渡し_扉_${String(W.meta.idx).padStart(2, '0')}${W.meta.short || ''}.xlsx`;
      const O = P.orderTobira(W, sort);
      const buf = await P.writeTobira(env.ExcelJS, W, O.rows);
      const n = O.rows.length;
      const offLines = O.off.map(o => `${o.text}　p.${o.page}（作業Excelでは「${o.after || '先頭'}」の次にあった）`);
      let b;
      if (!O.off.length) {
        b = box('ok', `${ch}：扉ページを出しました（${n}行）`, [sort ? 'ページ順になっていました（並べ直した行はありません）' : '作業Excelの並びのまま出しました'].map(esc));
      } else if (sort) {
        b = box('warn', `${ch}：扉ページを出しました（${n}行）。ページ順に並べ直した行が${O.off.length}行あります`, offLines.slice(0, 10).map(esc));
      } else {
        b = box('warn', `${ch}：扉ページを出しました（${n}行）。ページ順になっていない行が${O.off.length}行あります（作業Excelの並びのまま出しました）`, offLines.slice(0, 10).map(esc));
      }
      if (offLines.length > 10) b.append(listDetails(`ほかの${offLines.length - 10}行を見る`, offLines.slice(10)));
      if (sort && O.noNum.length) b.append(listDetails(`ページが数字でない${O.noNum.length}行は、作業Excelの位置のままにしました`, O.noNum));
      b.append(dlLink(name, buf));
      it.el = b; res.prepend(b);
    } else {
      const lines = [];
      if (c.undecided.length) lines.push(`まだ決まっていない行が${c.undecided.length}行あるため、出せません`);
      if (c.asking.length) lines.push(`「役員確認」のままの行が${c.asking.length}行あります（返事が来たら「載せる」か「載せない」に変えてください）`);
      if (c.noPage.length) lines.push(`「載せる」なのにページが空の行が${c.noPage.length}行あります（あとで入れる場合は「●●」と書いてください）`);
      const b = box('ng', `${ch}：まだ出せません`, lines.map(esc));
      if (c.undecided.length) b.append(listDetails(`まだ決まっていない${c.undecided.length}行を見る`, c.undecided.map(r => `${r.row}行目　${r.text}`)));
      if (c.asking.length) b.append(listDetails(`役員確認の${c.asking.length}行を見る`, c.asking.map(r => `${r.row}行目　${r.text}`)));
      if (c.noPage.length) b.append(listDetails(`ページが空の${c.noPage.length}行を見る`, c.noPage.map(r => `${r.row}行目　${r.text}`)));
      it.el = b; res.prepend(b);
    }
  }
}
$('out-tob-sort').addEventListener('change', () => { if (TOB_LAST.length) renderTobira(TOB_LAST); });
setupDrop(document.querySelector('[data-input="out-tob"]'), async files => {
  if (!SETTINGS) { alert('先に「設定」で前回の索引を取り込んでください。'); return; }
  const res = $('out-tob-res');
  const env = await getEnv();
  $('out-tob-sort').checked = true;   // 次の章は、いつも「並べ直す」から始める
  TOB_LAST = [];
  for (const f of files.filter(isXlsx)) {
    try {
      const W = await P.readWork(env.ExcelJS, (await readFile(f)).data);
      TOB_LAST.push({ f, W });
    } catch (e) { res.prepend(box('ng', `${f.name}：読めませんでした`, [esc(errText(e))])); }
  }
  await renderTobira(TOB_LAST);
});
setupDrop(document.querySelector('[data-input="out-kan"]'), async files => {
  if (!SETTINGS) { alert('先に「設定」で前回の索引を取り込んでください。'); return; }
  const res = $('out-kan-res'); res.innerHTML = '';
  try {
    const env = await getEnv();
    let sis = null, sisNote = '';
    for (const f of files.filter(isPdf)) {
      const pages = await P.readNewPdf(env, await readFile(f), m => progress('out-progress', m));
      const c = P.classify(SETTINGS, pages);
      if (c.kind === 'sisetsu') { sis = c.sis; sisNote = `施設基準の章：${f.name}`; SIS = { name: f.name, date: new Date().toISOString(), sis }; await DB.set('sisetsu', SIS); renderSis(); }
      else res.append(box('warn', `${f.name}：施設基準の章のPDFではないので使いませんでした`));
    }
    if (!sis && SIS) { sis = SIS.sis; sisNote = `施設基準の章：前に入れた ${SIS.name}（${new Date(SIS.date).toLocaleDateString('ja-JP')}）を使いました`; }
    progress('out-progress', '判定済みExcelを読んでいます');
    const works = []; const bad = [];
    for (const f of files.filter(isXlsx)) {
      try { const W = await P.readWork(env.ExcelJS, (await readFile(f)).data); W.file = f.name; works.push(W); }
      catch (e) { bad.push(`${f.name}：${errText(e)}`); }
    }
    // 同じ章のExcelが2つあれば新しいほう
    const byIdx = new Map(); const dupNotes = [];
    for (const W of works) {
      const k = String(W.meta.idx); const o = byIdx.get(k);
      if (o) { const keep = (W.meta.created || '') > (o.meta.created || '') ? W : o; dupNotes.push(`同じ章のExcelが2つ：${o.file} と ${W.file} → ${keep.file} を使いました`); byIdx.set(k, keep); }
      else byIdx.set(k, W);
    }
    const use = [...byIdx.values()].sort((a, b) => (+a.meta.idx) - (+b.meta.idx));
    if (!use.length) { progress('out-progress', null); res.append(box('ng', '判定済みExcelが入っていません', bad.map(esc))); return; }
    const missing = SETTINGS.old.chapters.map((c, i) => byIdx.has(String(i)) ? null : c.title).filter(Boolean);
    progress('out-progress', '巻末索引をまとめています'); await tick();
    const sisObj = sis ? new P.Sisetsu(sis, SETTINGS.old.oldTitles) : null;
    const M = P.mergeKanmatsu(use, { sis: sisObj, aliasK: P.makeAliasK(SETTINGS.old) });
    M.missing = missing; M.notes.unshift(...dupNotes);
    if (!sis) M.notes.unshift('施設基準の章のPDFが無いので、（施設基準）のページは作業Excelの仮のページのままです');
    const final = !missing.length && !M.undecided && !!sis;
    const year = new Date().getFullYear();
    const lines = [`入れたもの：${use.map(W => W.meta.short).join('・')}` + (sisNote ? `、${esc(sisNote)}` : '')];
    if (missing.length) lines.push(`足りないもの：${esc(missing.join('・'))}`);
    if (!sis) lines.push('足りないもの：施設基準の章のPDF');
    if (M.undecided) lines.push(`まだ決まっていない行：${M.undecided}行`);
    lines.push(`巻末の語：${fmt(M.entries.length)}語`);
    if (final) {
      const buf = await P.writeKanmatsu(env.ExcelJS, M, false);
      const hb = await P.writeHandover(env.ExcelJS, M, use, String(year));
      const b = box('ok', '業者に渡す巻末索引を出しました', lines);
      const w = document.createElement('div'); w.className = 'stack';
      w.append(dlLink('業者渡し_巻末索引.xlsx', buf, '業者渡し_巻末索引をダウンロード'));
      w.append(dlLink(`引き継ぎ_${year}年版.xlsx`, hb, `引き継ぎ_${year}年版をダウンロード`));
      const s = document.createElement('span'); s.className = 'sub'; s.textContent = '引き継ぎファイルは、完成品のPDFと一緒に共有フォルダの今回の年版のフォルダに保管してください。次の改定の「設定」で使います。';
      w.append(s); b.append(w); res.append(b);
    } else {
      const buf = await P.writeKanmatsu(env.ExcelJS, M, true);
      const b = box('warn', '全部そろっていないので、試し出し（社内確認用）になります', lines);
      b.append(dlLink('試し出し_巻末索引.xlsx', buf, '試し出しをダウンロード'));
      const s = document.createElement('span'); s.className = 'sub'; s.textContent = '全章そろって、まだ決まっていない行が0になり、施設基準の章のPDFがあれば、業者渡し用と次回のための引き継ぎファイルが出ます。';
      b.append(s); res.append(b);
      if (M.undecidedRows.length) res.append(listDetails(`まだ決まっていない${M.undecidedRows.length}行を見る`, M.undecidedRows));
    }
    if (M.problems.length) res.append(box('warn', `確かめてほしいこと（${M.problems.length}件）`, [], listDetails('一覧を見る', M.problems)));
    if (M.notes.length) res.append(box('', `ツールがしたこと（${M.notes.length}件）`, [], listDetails('一覧を見る', M.notes)));
    if (bad.length) res.append(box('ng', '読めなかったExcel', bad.map(esc)));
  } catch (e) { res.append(box('ng', '出せませんでした', [esc(errText(e))])); }
  progress('out-progress', null);
});

// ---------- はじめに
$('forget').addEventListener('click', async e => {
  e.preventDefault();
  if (!confirm('このパソコンに覚えている設定（前回の索引・読み・施設基準の章）を消します。よろしいですか？')) return;
  await DB.del('settings'); await DB.del('sisetsu'); SETTINGS = null; SIS = null; renderSettings(); route();
});
(async () => {
  try { SETTINGS = await DB.get('settings') || null; SIS = await DB.get('sisetsu') || null; } catch (e) { console.warn(e); }
  renderSettings(); route();
})();
