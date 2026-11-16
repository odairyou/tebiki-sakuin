// 章ごとの照合：前回の扉・巻末の語に新原稿のページを当て、新規候補を挿入する（Python版 chapter.py の移植）
import { nz } from './pdftext.js';
import { ratio, uniq } from './util.js';

const num = (a, b) => a - b;
const closest = (arr, exp) => arr.reduce((b, p) => (b === null || Math.abs(p - exp) < Math.abs(b - exp)) ? p : b, null);
const pj = arr => arr.join('・');

export function pcs(t, strip = false) {
  t = nz(t);
  if (t.startsWith('(') && t.endsWith(')')) t = t.slice(1, -1);
  const out = [t];
  const m = t.match(/^([^()]+)\((.+)\)$/);
  if (m) out.push(m[1], m[2]);
  for (const o of [...out]) {
    const o4 = o.replace(/(の術式|の考え方|について|の算定)$/, '');
    if (strip && o4 !== o && o4.length >= 3) out.push(o4);
  }
  for (const o of [...out]) {
    const segs = o.split('・');
    if (segs.length > 1 && segs.every(x => x.length >= 2 && !/^\d+$/.test(x))) out.push(...segs);
  }
  for (const o of [...out]) {
    const m2 = o.match(/^(.*?\d)(・\d)+$/);
    if (m2) out.push(m2[1]);
    const o2 = o.replace(/[\d・]+$/, '');
    if (o2 !== o && o2.length >= 2) out.push(o2);
    const o3 = (o2 !== o ? o2 : o).replace(/加算$/, '');
    if (o3 !== o && o3.length >= 2) out.push(o3);
  }
  const res = uniq(out).filter(o => o.length >= 2);
  return res.length ? res : (t ? [t] : []);
}

export class Chapter {
  // old: readOld の結果、och: その章（old.chapters の1つ）、R: readDraft の結果、Y: Yomi、sis: Sisetsu（無ければ null）
  constructor(old, och, R, Y, sis) {
    this.old = old; this.och = och; this.R = R; this.Y = Y; this.sis = sis;
    this.Z = R.Z; this.PG = R.pages;
    this.old_lo = och.lo; this.old_hi = och.hi;
    this.off0 = (R.tobiraPage || och.tob) - och.tob;
    this.new_lo = this.PG[0]; this.new_hi = this.PG[this.PG.length - 1];
    this.tob = och.entries;
    const s = old.sisetsu;
    this.isK = p => s ? (p >= s.lo && p <= s.hi) : false;
  }
  inr(p) { return this.old_lo <= p && p <= this.old_hi; }
  z(p, k) { return (this.Z[p] && this.Z[p][k]) || ''; }

  strong(ps) {
    const out = new Set();
    for (const p of this.PG) for (const c of ps) {
      const cl = c.toLowerCase();
      if (this.z(p, 'LB').toLowerCase().includes(cl) || this.z(p, 'AB').toLowerCase().includes('|' + cl + '|')) out.add(p);
    }
    return [...out].sort(num);
  }
  zone(ps, zone, pages = null) {
    const out = new Set();
    for (const p of (pages && pages.length ? pages : this.PG)) for (const c of ps) if (this.z(p, zone).toLowerCase().includes(c.toLowerCase())) out.add(p);
    return [...out].sort(num);
  }

  resMain(text, oldp) {
    const ps = pcs(text), exp = oldp + this.off;
    const st = this.strong(ps);
    if (st.length) {
      const best = closest(st, exp), near = st.filter(p => Math.abs(p - exp) <= 2);
      if (st.length === 1 && Math.abs(best - exp) <= 3) { this.off = best - oldp; return [best, 'ok', '']; }
      if (st.length === 1) return [best, 'chk', `掲載位置が移動（前回の並びなら p.${exp} 付近 → p.${best}）`];
      if (near.length === 1) { this.off = best - oldp; return [best, 'ok', '見出し・略称が p.' + pj(st) + ' にあり、前回の位置に近いページを採用']; }
      return [best, 'chk', '候補が複数：p.' + pj(st)];
    }
    const ps2 = pcs(text, true).filter(x => !ps.includes(x));
    for (const [zone, lab] of [['RB', '右欄の囲み・表'], ['LN', '左欄の注'], ['STRIP', ''], ['ALL', '本文']]) {
      if (zone === 'STRIP') {
        const st2 = ps2.length ? this.strong(ps2) : [];
        if (st2.length) return [closest(st2, exp), 'chk', `「${ps2[0]}」の見出しのページ（語そのものは見出しに無い）`];
        continue;
      }
      const w = this.zone(ps, zone);
      if (w.length) return [closest(w, exp), 'chk', `見出しには無く${lab}にあり` + (w.length > 1 ? '。候補 p.' + pj(w.slice(0, 6)) : '')];
    }
    const parts = nz(text).replace(/\(.*?\)/g, '').split(/を行った場合の|と|及び/).filter(x => x.length >= 2);
    if (parts.length === 2) {
      const both = this.PG.filter(p => parts.every(x => this.z(p, 'ALL').toLowerCase().includes(x.toLowerCase())));
      if (both.length) return [closest(both, exp), 'chk', `「${parts[0]}」と「${parts[1]}」が両方出てくるページ` + (both.length > 1 ? '。候補 p.' + pj(both.slice(0, 6)) : '')];
    }
    const t0 = nz(text).replace(/\(.*?\)/g, '');
    const core = t0.replace(/(の)?(ループ)?(除去|撤去)$/, '');
    if (core !== nz(text) && core.length >= 2) {
      const near = this.PG.filter(p => Math.abs(p - exp) <= 4 && this.z(p, 'ALL').toLowerCase().includes(core.toLowerCase()));
      if (near.length) { const b = closest(near, exp); return [b, 'chk', `「${core}」が p.${b} の表・本文にあり（前回も表から起こした語）`]; }
    }
    const f = this.fuzzy(text);
    if (f) return [f[0], 'chk', `同じ語は無いが、よく似た「${f[1]}」が p.${f[0]} にあり（名前の変更かも）`];
    return [null, 'ng', '新原稿に見つからない'];
  }
  section(start) {
    const nx = this.topPages.filter(p => p > start);
    return [start, Math.min(nx.length ? nx[0] - 1 : this.new_hi, start + 12)];
  }
  resSub(text, parNew, oldp) {
    const ps = pcs(text);
    let lo = parNew || Math.max(this.new_lo, Math.min(this.new_hi, oldp + this.off));
    let hi; [lo, hi] = this.section(lo); hi = Math.max(hi, lo);
    const rng = this.PG.filter(p => lo <= p && p <= hi);
    for (const [zone, lab] of [['LB', '左欄の見出し'], ['AB', '略称'], ['RB', '囲み・表'], ['ALL', '本文']]) {
      const h = this.zone(ps, zone, rng);
      if (h.length && h.includes(parNew) && zone !== 'ALL') return [parNew, 'ok', ''];
      if (h.length) {
        const best = closest(h, oldp + this.off);
        if (lab !== '本文' && h.length === 1) return [best, 'ok', ''];
        return [best, 'chk', `親項目の区間（p.${lo}〜${hi}）の${lab}で発見` + (h.length > 1 ? '。候補 p.' + pj(h) : '')];
      }
    }
    const f = this.fuzzy(text, rng);
    if (f) return [f[0], 'chk', `同じ語は無いが、よく似た「${f[1]}」が p.${f[0]} にあり（名前の変更かも）`];
    return [null, 'ng', `親項目の区間（p.${lo}〜${hi}）に見つからない`];
  }
  fuzzy(text, pages = null, th = 0.72) {
    const p0 = pcs(text); if (!p0.length) return null;
    const t = p0[0].replace(/[\d・]/g, ''); if (t.length < 3) return null;
    const R = this.R; let best = null, bs = 0;
    for (const [p, , x] of [...R.heads, ...R.titles, ...R.boxes]) {
      if (pages !== null && !pages.includes(p)) continue;
      const x2 = nz(x).replace(/[\d・]/g, '');
      let sc = ratio(t, x2);
      if (x2.includes(t) && t.length >= 4) sc = Math.max(sc, 0.8);
      if (sc > bs) { bs = sc; best = [p, nz(x)]; }
    }
    return bs >= th ? best : null;
  }
  resSubOfHeadless(parent, text, oldp) {
    const sub = pcs(text), par = pcs(parent);
    const st = this.strong(sub);
    if (!st.length) return this.resSub(text, null, oldp);
    const exp = oldp + this.off;
    const tops = st.filter(p => this.topPages.includes(p));
    const s = closest(tops.length ? tops : st, exp);
    const [lo, hi] = this.section(s); const rng = this.PG.filter(p => lo <= p && p <= hi);
    for (const zone of ['LB', 'AB', 'RB', 'ALL']) {
      const h = this.zone(par, zone, rng);
      if (h.length) {
        const best = closest(h, exp);
        return [best, zone !== 'ALL' ? 'ok' : 'chk', zone !== 'ALL' ? '' : `「${text}」の区間（p.${lo}〜${hi}）の本文で親の語を発見`];
      }
    }
    return [s, 'chk', `「${text}」の区間（p.${lo}〜${hi}）に親の語が無いため、その見出しのページを仮に採用`];
  }
  aliasK(e, s) {
    if (!this._kmap) {
      this._kmap = new Map(); let m = null;
      for (const x of this.old.back) {
        if (!x.sub) m = x;
        else if (m && x.text.includes('施設基準')) {
          const k = JSON.stringify([m.pages, x.text, x.pages]);
          if (!this._kmap.has(k)) this._kmap.set(k, []); this._kmap.get(k).push(m.text);
        }
      }
    }
    return (this._kmap.get(JSON.stringify([e.pages, s.text, s.pages])) || []).filter(t => t !== e.text);
  }
  sisFind(parent, oldp, opt = {}) {
    if (!this.sis) return [null, '施設基準の章のPDFが無いため、ページは入れていない（業者データを出すときに入れる）'];
    return this.sis.find(parent, oldp, opt.name || null, opt.alts || [], !!opt.kinyu);
  }

  run() {
    const old = this.old; this.off = this.off0;
    const pre = this.tob.map(e => e.sub ? null : this.resMain(e.text, e.pages[0]));
    this.topPages = uniq(pre.filter(r => r && r[0] && r[1] === 'ok').map(r => r[0])).sort(num);
    this.off = this.off0;
    const TOP = []; let par = null, no = 0;
    this.tob.forEach((e, i) => {
      let n, st, note;
      if (!e.sub) { no++; [n, st, note] = pre[i]; par = n; }
      else [n, st, note] = this.resSub(e.text, par, e.pages[0]);
      TOP.push({ text: e.text, sub: e.sub, oldp: e.pages.slice(0, 1), newp: n ? [n] : [], other: [], st, note, no, new: false });
    });
    // 巻末
    const groups = []; let cur = null;
    for (const e of old.back) {
      if (!e.sub) { cur = { e, subs: [] }; groups.push(cur); }
      else if (cur) cur.subs.push(e);
    }
    const BOT = []; this.off = this.off0;
    for (const g of groups) {
      const e = g.e; const own = e.pages.filter(p => this.inr(p));
      const subsIn = g.subs.filter(s => s.pages.some(p => this.inr(p)));
      if (!own.length && !(!e.pages.length && subsIn.length)) continue;
      const headless = !e.pages.length;
      const newp = [], notes0 = [], sts = [];
      let notes = notes0;
      for (const p of e.pages) if (this.inr(p)) {
        const [n, st, note] = this.resMain(e.text, p); sts.push(st);
        if (n && !newp.includes(n)) newp.push(n);
        if (note && !notes.includes(note)) notes.push(note);
      }
      const other = e.pages.filter(p => !this.inr(p));
      if (other.length) notes.push('p.' + pj(other) + ' は他の章（その章の処理で更新）');
      const st = headless ? 'head' : (!newp.length ? 'ng' : (sts.includes('chk') || sts.includes('ng') ? 'chk' : 'ok'));
      if (headless) notes = ['ページを持たない親見出し（副項目だけにページがある）'];
      const key = own.length ? Math.min(...own) : Math.min(...subsIn.map(s => Math.min(...s.pages.filter(p => this.inr(p)))));
      const blk = [{ text: e.text, sub: false, oldp: e.pages, newp: [...newp].sort(num), other, st, note: notes.join('。'), new: false, key }];
      const parentNew = newp.length ? newp[0] : null;
      for (const s of (own.length ? g.subs : subsIn)) {
        const sp = s.pages; const r = { text: s.text, sub: true, oldp: sp, newp: [], other: [], st: '', note: '', new: false, key };
        if (sp.length && sp.every(p => this.isK(p))) {
          const kin = s.text.includes('記入例');
          const base = kin ? g.subs.filter(x => x.text === '（施設基準）' && x.pages.length).map(x => x.pages) : [];
          const [np, nt] = this.sisFind(e.text, base.length ? Math.min(...base[0]) : Math.min(...sp), { alts: this.aliasK(e, s), kinyu: kin });
          Object.assign(r, { st: np ? 'kari' : (this.sis ? 'ng' : 'auto'), newp: np ? [np] : [], note: nt, kijun: true });
        } else if (sp.some(p => this.inr(p))) {
          for (const p of sp) {
            if (this.inr(p)) {
              let n, st2, note2;
              if (headless) [n, st2, note2] = this.resSubOfHeadless(e.text, s.text, p);
              else {
                [n, st2, note2] = this.resSub(s.text, parentNew, p);
                if (st2 === 'ng') { const [n2, , note3] = this.resMain(s.text, p); if (n2) { n = n2; st2 = 'chk'; note2 = note3; } }
              }
              if (n) r.newp.push(n);
              r.st = st2; r.note = note2;
            } else r.other.push(p);
          }
        } else Object.assign(r, { st: 'auto', other: sp, note: '他の章のページ（その章の処理で更新）' });
        blk.push(r);
      }
      BOT.push(blk);
    }
    this.linkSynonyms(BOT);
    this.TOP = TOP; this.BOT = BOT;
    this.addNew();
    for (const [olds, newt] of this.renames) {
      for (const r of [...this.TOP, ...this.BOT.flat()]) {
        if ((r.st === 'ng' || r.st === 'chk') && !r.sub && pcs(r.text).some(x => olds.has(x)))
          r.note = (r.note ? r.note + '。' : '') + `新原稿の「${newt}」に名称変更かも`;
      }
    }
    this.assignNo();
    return this;
  }

  linkSynonyms(BOT) {
    const sig = b => JSON.stringify(b.slice(1).map(s => [nz(s.text), s.oldp]));
    const by = new Map();
    for (const b of BOT) if (b.length >= 3) { const k = sig(b); if (!by.has(k)) by.set(k, []); by.get(k).push(b); }
    for (const bs of by.values()) {
      if (bs.length < 2) continue;
      for (let i = 1; i < bs[0].length; i++) {
        const rows = bs.map(b => b[i]);
        let good = rows.filter(r => r.st === 'ok'); if (!good.length) good = rows.filter(r => r.st === 'chk');
        if (!good.length) continue;
        const g = good[0];
        for (const r of rows) {
          if (r === g) continue;
          if (r.st === 'ng' || r.st === 'chk' || JSON.stringify(r.newp) !== JSON.stringify(g.newp)) {
            const partner = bs.find(b => b[i] === g)[0].text;
            Object.assign(r, { newp: [...g.newp], st: g.st, note: `「${partner}」の同じ副項目と連動（前回の索引で全く同じページを指していたため）` });
          }
        }
      }
    }
  }

  addNew() {
    const old = this.old, R = this.R;
    const oldall = new Set();
    for (const e of old.back) pcs(e.text).forEach(x => oldall.add(x));
    for (const c of old.chapters) for (const e of c.entries) pcs(e.text).forEach(x => oldall.add(x));
    const oldArr = [...oldall];
    const known = t => {
      for (const u of [t, t.replace(/\d+$/, '')])
        if (u.length >= 2 && (oldall.has(u) || (u.length >= 3 && oldArr.some(o => o.includes(u))))) return true;
      return false;
    };
    const table = {}; for (const [a, b] of R.abbrTable) table[nz(a)] = nz(b);
    const heads = []; const seen = new Set();
    for (const [p, y, t] of R.headsMain) {
      const t0 = nz(t);
      if (seen.has(t0) || known(t0) || known(t0.replace(/\(.*?\)$/, ''))) continue;
      if (heads.some(([q, , x]) => q === p && ratio(t0, x) > 0.85)) continue;
      seen.add(t0); heads.push([p, y, t0]);
    }
    const abbrOf = {};
    for (const [p, y, t] of heads) {
      let a = table[t] || table[t.replace(/\(.*?\)$/, '')];
      if (!a) { const below = R.abbrs.filter(([q, yy]) => q === p && yy - y > 0 && yy - y <= 40).map(x => x[2]); a = below[0] || null; }
      if (a) abbrOf[t] = a;
    }
    const abbrs = [];
    const abVals = Object.values(abbrOf);
    for (const [p, y, a] of R.abbrs) {
      if (known(a) || abVals.includes(a) || abbrs.some(x => x[2] === a)) continue;
      const par = R.heads.filter(([q, yy]) => q === p && yy < y).map(x => x[2]);
      abbrs.push([p, y, a, par.length ? par[par.length - 1] : null]);
    }
    const boxes = R.boxes;
    const kijun = t => boxes.some(([, , b]) => b.includes('施設基準') && (b.includes(t) || b.includes(t.replace(/\d+$/, ''))));
    this.newTop = []; this.newBot = []; this.renames = [];
    const allTob = old.chapters.flatMap(c => c.entries);
    for (const [p, y, t] of heads) {
      const a = abbrOf[t] || null; const [yv, gs] = this.Y.get(t);
      let note = '新原稿の左欄の見出し。前回の索引に無い';
      if (a && a.length >= 2) {
        const same = [...allTob.filter(e => pcs(e.text).includes(a) && !pcs(e.text).includes(t) && nz(e.text) !== a).map(e => e.text),
          ...old.back.filter(e => !e.sub && pcs(e.text).includes(a) && !pcs(e.text).includes(t) && nz(e.text) !== a).map(e => e.text)];
        if (same.length) {
          const oldT = same.reduce((m, x) => x.length > m.length ? x : m, '');
          note = `新原稿の左欄の見出し。略称「${a}」は前回の「${oldT}」と同じ → 名称変更かも`;
          this.renames.push([new Set(pcs(oldT).filter(x => x !== a)), t + `（${a}）`]);
        }
      }
      if (gs.length) note += '。読みの「' + gs.join('」「') + '」は自動推定（要確認）';
      this.newTop.push({ text: t + (a ? `（${a}）` : ''), sub: false, oldp: [], newp: [p], other: [], st: 'new', note, new: true, yflag: gs });
      const blk = [{ text: t, sub: false, oldp: [], newp: [p], other: [], st: 'new', note, new: true, yomi: yv, yflag: gs }];
      if (kijun(t)) {
        const [np, nt] = this.sisFind(t, null, { name: t });
        const k = { text: '（施設基準）', sub: true, oldp: [], newp: np ? [np] : [], other: [], st: 'new', note: '新原稿に「…の施設基準 ➡P参照」の枠あり。' + nt, new: true, kijun: true };
        blk.push(k); this.newTop.push({ ...k });
      }
      this.newBot.push(blk);
      if (a) {
        const [ya, ga] = this.Y.get(a);
        const n2 = `新原稿の略称（赤字）。「${t}」の略称` + (ga.length ? '。読みの「' + ga.join('」「') + '」は自動推定（要確認）' : '');
        const b2 = [{ text: a, sub: false, oldp: [], newp: [p], other: [], st: 'new', note: n2, new: true, yomi: ya, yflag: ga }];
        if (kijun(t)) b2.push({ ...blk[blk.length - 1] });
        this.newBot.push(b2);
      }
    }
    // 早見表：大見出しの「○○早見表」のうち、前回の索引の早見表の行がまだ拾っていないもの
    const hayamiDone = new Set([...this.TOP, ...this.BOT.flat()].filter(r => r.text.includes('早見表') && r.newp.length).flatMap(r => r.newp));
    const hseen = new Set();
    for (const [p, y, t0] of (R.titlesJoined || [])) {
      const t = nz(t0);
      if (!t.includes('早見表') || hayamiDone.has(p) || hseen.has(t) || known(t)) continue;
      hseen.add(t);
      const disp = t.replace(/\(/g, '（').replace(/\)/g, '）');
      const note = `新原稿の大見出し「${disp}」。前回の索引に無い早見表`;
      this.newTop.push({ text: '（早見表）', sub: true, oldp: [], newp: [p], other: [], st: 'new', note: note + '（扉では見出しの下の副項目として入れた）', new: true, hayami: true });
      const [yv, gs] = this.Y.get(disp);
      this.newBot.push([{ text: disp, sub: false, oldp: [], newp: [p], other: [], st: 'new', new: true, yomi: yv, yflag: gs,
        note: note + (gs.length ? '。読みの「' + gs.join('」「') + '」は自動推定（要確認）' : '') }]);
    }
    const exist = new Set(this.BOT.map(b => nz(b[0].text)));
    this.newBot = this.newBot.filter(b => !exist.has(nz(b[0].text)));
    for (const [p, y, a, par] of abbrs) {
      if (exist.has(a)) continue;
      const [ya, ga] = this.Y.get(a);
      const n2 = '新原稿の略称（赤字）。' + (par ? `「${par}」の略称` : '') + (ga.length ? '。読みの「' + ga.join('」「') + '」は自動推定（要確認）' : '');
      this.newBot.push([{ text: a, sub: false, oldp: [], newp: [p], other: [], st: 'new', note: n2, new: true, yomi: ya, yflag: ga, par }]);
    }
  }

  assignNo() {
    let TOP = this.TOP;
    const oldM = this.tob.filter(e => !e.sub).length, oldS = this.tob.filter(e => e.sub).length;
    this.subMode = oldS / Math.max(oldM, 1) >= 5;
    // 早見表：ページ順で直前の扉の主項目の副項目に入れる
    for (const c of this.newTop) {
      if (!c.hayami) continue;
      const p = c.newp[0];
      const idx = TOP.map((r, i) => i).filter(i => !TOP[i].sub && TOP[i].newp.length && TOP[i].newp[0] <= p && TOP[i].st !== 'chk');
      const i = idx.length ? idx.reduce((b, i) => (TOP[i].newp[0] > TOP[b].newp[0] || (TOP[i].newp[0] === TOP[b].newp[0] && i > b)) ? i : b, idx[0]) : -1;
      if (i < 0) { TOP.unshift({ ...c, sub: false, text: c.note.match(/「(.*?)」/)[1] }); continue; }
      c.note = '「' + TOP[i].text + '」の副項目として入れた。' + c.note.replace('（扉では見出しの下の副項目として入れた）', '');
      let j = i + 1, k = i + 1;
      while (j < TOP.length && TOP[j].sub) { if (TOP[j].newp.length && TOP[j].newp[0] <= p) k = j + 1; j++; }
      TOP.splice(k, 0, c);
    }
    for (const c of this.newTop) {
      if (c.sub) continue;
      const p = c.newp[0];
      const idx = TOP.map((r, i) => i).filter(i => !TOP[i].sub && TOP[i].newp.length && TOP[i].newp[0] <= p && TOP[i].st !== 'chk');
      const i = idx.length ? idx.reduce((b, i) => (TOP[i].newp[0] > TOP[b].newp[0] || (TOP[i].newp[0] === TOP[b].newp[0] && i > b)) ? i : b, idx[0]) : -1;
      if (this.subMode && i >= 0) {
        c.sub = true; c.text = '（' + c.text + '）'; c.asSub = true;
        c.note = '前回の扉にならい、「' + TOP[i].text + '」の副項目として入れた。' + c.note;
        let j = i + 1, k = i + 1;
        while (j < TOP.length && TOP[j].sub) { if (TOP[j].newp.length && TOP[j].newp[0] <= p) k = j + 1; j++; }
        TOP.splice(k, 0, c); continue;
      }
      let j = i + 1; while (j < TOP.length && TOP[j].sub) j++;
      TOP.splice(j, 0, c);
    }
    const newsub = new Map(); let last = null;
    for (const c of this.newTop) {
      if (c.hayami) continue;
      if (c.asSub) last = null;
      else if (!c.sub) last = c.text;
      else if (last) { if (!newsub.has(last)) newsub.set(last, []); newsub.get(last).push(c); }
    }
    const out = [];
    for (const r of TOP) { out.push(r); if (r.new && !r.asSub) for (const s of (newsub.get(r.text) || [])) out.push(s); }
    this.TOP = TOP = out;
    let no = 0; for (const r of TOP) { if (!r.sub) no++; r.no = no; }
    const gmain = new Map(); for (const r of TOP) if (!r.sub) gmain.set(r.no, r);
    const blocks = [...this.BOT, ...this.newBot].map((rows, i) => ({ rows, no: null, head: false, i }));
    for (const [n, m] of gmain) {
      const P = pcs(m.text);
      for (let k = 0; k < P.length; k++) {
        const c = blocks.filter(b => b.no === null && nz(b.rows[0].text) === P[k] && b.rows[0].new === m.new);
        if (c.length) { c[0].no = n; c[0].head = true; break; }
      }
    }
    for (const b of blocks) {
      if (b.no !== null) continue;
      const r0 = b.rows[0]; const par = r0.par;
      for (const [n, m] of gmain) {
        const pm = pcs(m.text);
        if ((par && pm.includes(par)) || pcs(r0.text).some(x => pm.includes(x)) || (r0.new && m.text.includes(r0.text))) { b.no = n; break; }
      }
    }
    for (const b of blocks) {
      if (b.no !== null) continue;
      const r0 = b.rows[0]; let c;
      if (r0.new) c = [...gmain].filter(([, m]) => m.newp.length && m.newp[0] <= r0.newp[0]).map(x => x[0]);
      else c = [...gmain].filter(([, m]) => m.oldp.length && m.oldp[0] <= r0.key).map(x => x[0]);
      b.no = c.length ? Math.max(...c) : 1;
    }
    const pkey = b => {
      const r = b.rows[0];
      if (r.newp.length) return r.newp[0];
      const sp = b.rows.slice(1).filter(x => x.newp.length).map(x => x.newp[0]);
      return sp.length ? Math.min(...sp) : (r.key || 0) + this.off0;
    };
    blocks.sort((a, b) => (a.no - b.no) || ((!a.head) - (!b.head)) || (pkey(a) - pkey(b)) || (a.i - b.i));
    const topSub = new Map();
    for (const r of TOP) if (r.sub) { if (!topSub.has(r.no)) topSub.set(r.no, new Set()); topSub.get(r.no).add(nz(r.text)); }
    const heads = new Set(blocks.filter(b => b.head).map(b => b.no));
    const LOW = [];
    for (const b of blocks) b.rows.forEach((r, k) => {
      r.no = b.no; r.first = k === 0;
      if (r.new) r.prev = '—';
      else if (k === 0) r.prev = b.head ? '両方' : '巻末のみ';
      else r.prev = (b.head && topSub.get(b.no)?.has(nz(r.text))) ? '両方' : '巻末のみ';
      if (!r.sub && !('yomi' in r)) { const [y, g] = this.Y.get(r.text); r.yomi = y; r.yflag = g; }
      LOW.push(r);
    });
    const lowSub = new Map();
    for (const r of LOW) if (r.sub) { if (!lowSub.has(r.no)) lowSub.set(r.no, new Set()); lowSub.get(r.no).add(nz(r.text)); }
    for (const r of TOP) {
      if (r.new) r.prev = '—';
      else if (!r.sub) r.prev = heads.has(r.no) ? '両方' : '扉のみ';
      else r.prev = lowSub.get(r.no)?.has(nz(r.text)) ? '両方' : '扉のみ';
    }
    this.LOW = LOW;
  }
}

// 業者データを出すときの（施設基準）の入れ直し用：同じページ・同じ施設基準ページを持つ別名の主項目
export function makeAliasK(old) {
  const byMain = new Map(); const kmap = new Map(); let m = null;
  for (const x of old.back) {
    if (!x.sub) m = x;
    else if (m && x.text.includes('施設基準')) {
      const k = JSON.stringify([m.pages, x.text, x.pages]);
      if (!kmap.has(k)) kmap.set(k, []); kmap.get(k).push(m.text);
      byMain.set(nz(m.text) + '|' + nz(x.text), k);
    }
  }
  return (parent, subText) => {
    const k = byMain.get(nz(parent) + '|' + nz(subText));
    return k ? kmap.get(k).filter(t => nz(t) !== nz(parent)) : [];
  };
}
