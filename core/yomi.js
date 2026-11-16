// 読みがな：前回の読み → 前回の読みから学んだ辞書 → 形態素解析（kuromoji）による推定（Python版 yomi.py の移植）
import { nz } from './pdftext.js';

const N = s => String(s ?? '').normalize('NFKC').trim();
export const hira = s => [...s].map(c => (c >= 'ァ' && c <= 'ヶ') ? String.fromCharCode(c.charCodeAt(0) - 0x60) : c).join('');
export const clean = y => hira(N(y)).replace(/[^ぁ-んー]/g, '');
const esc = s => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
const AL = Object.fromEntries([...'abcdefghijklmnopqrstuvwxyz'].map((c, i) => [c, 'えー びー しー でぃー いー えふ じー えいち あい じぇー けー える えむ えぬ おー ぴー きゅー あーる えす てぃー ゆー ぶい だぶりゅー えっくす わい ぜっと'.split(' ')[i]]));
const NUM = { 1: 'いち', 2: 'に', 3: 'さん', 4: 'よん', 5: 'ご', 6: 'ろく', 7: 'なな', 8: 'はち', 9: 'きゅう', 0: 'ぜろ' };
export const SEED = { '填塞': 'てんそく', '窩洞': 'かどう', '義歯床': 'ぎししょう', '鉤歯': 'こうし', '鉤': 'こう', '線副子': 'せんふくし', '副子': 'ふくし',
  '舌腫瘍': 'ぜつしゅよう', '舌接触': 'ぜつせっしょく', '口腔': 'こうくう', '補綴': 'ほてつ', '歯冠': 'しかん', '歯髄': 'しずい', '歯周': 'ししゅう',
  '歯肉': 'しにく', '歯槽': 'しそう', '根面': 'こんめん', '根管': 'こんかん', '咬合': 'こうごう', '頬': 'きょう', '抜髄': 'ばつずい', '貼薬': 'ちょうやく',
  '嚢胞': 'のうほう', 'う蝕': 'うしょく', 'フッ化物': 'ふっかぶつ', 'cond': 'こんど', '単処': 'たんしょ', '処置': 'しょち', '破折': 'はせつ',
  '歯科': 'しか', '加算': 'かさん', '初診': 'しょしん', '再診': 'さいしん', '掻爬': 'そうは', '搔爬': 'そうは', '剥離': 'はくり', '小帯': 'しょうたい', '弁': 'べん',
  '電子的': 'でんしてき', '診療情報': 'しんりょうじょうほう', '連携': 'れんけい', '体制': 'たいせい', '整備': 'せいび', '地域': 'ちいき', '算定': 'さんてい',
  '基礎知識': 'きそちしき', '次手術': 'じしゅじゅつ', '口画像': 'こうがぞう', '顕画像': 'けんがぞう', '画像': 'がぞう', '乳': 'にゅう', '情報通信機器': 'じょうほうつうしんきき',
  '用いた': 'もちいた', '時間外': 'じかんがい', '休日': 'きゅうじつ', '深夜': 'しんや', '医療': 'いりょう', '早期': 'そうき', '単純': 'たんじゅん', '評価料': 'ひょうかりょう',
  '外来': 'がいらい', '在宅': 'ざいたく', '原法': 'げんぽう', '改良法': 'かいりょうほう', '挿入': 'そうにゅう', '場合': 'ばあい', '中断': 'ちゅうだん', '離脱': 'りだつ',
  '継続支援': 'けいぞくしえん', '治療': 'ちりょう', '明細': 'めいさい', '物価対応料': 'ぶっかたいおうりょう', '重症化予防': 'じゅうしょうかよぼう', '強化': 'きょうか',
  '掲げる': 'かかげる', '単位': 'たんい', '期間': 'きかん', '点数': 'てんすう', '感染症': 'かんせんしょう', '種類': 'しゅるい' };

const isKana = c => /^[ぁ-んァ-ヶー]/.test(c);
const chunks = t => t.match(/[ぁ-んァ-ヶー]+|[^ぁ-んァ-ヶー]+/g) || [];
export function numY(n) {
  n = parseInt(n, 10); const d = ['', 'いち', 'に', 'さん', 'よん', 'ご', 'ろく', 'なな', 'はち', 'きゅう'];
  if (n === 0) return 'ぜろ';
  if (n >= 1000) return [...String(n)].map(c => NUM[c]).join('');
  let out = ''; const h = Math.floor(n / 100), r = n % 100, t = Math.floor(r / 10), o = r % 10;
  if (h) out += ({ 1: 'ひゃく', 3: 'さんびゃく', 6: 'ろっぴゃく', 8: 'はっぴゃく' })[h] || d[h] + 'ひゃく';
  if (t) out += (t === 1 ? '' : d[t]) + 'じゅう';
  if (o) out += d[o];
  return out;
}
function learn(pairs) {
  const D = new Map(); const multi = [];
  const key = c => c.replace(/[・\s]/g, '');
  for (const [t, y] of pairs) {
    const cs = chunks(t); const nk = cs.filter(c => !isKana(c));
    if (nk.length === 1) {
      const pat = cs.map(c => isKana(c) ? esc(hira(c)) : '(.+)').join('');
      const m = y.match(new RegExp('^' + pat + '$', 'u'));
      const k = key(nk[0]);
      if (m && k) { if (!D.has(k)) D.set(k, new Map()); const c = D.get(k); c.set(m[1], (c.get(m[1]) || 0) + 1); }
    } else if (nk.length) multi.push([cs, y]);
  }
  const known = {};
  for (const [k, c] of D) { let best = null, bn = -1; for (const [y, n] of c) if (n > bn) { bn = n; best = y; } known[k] = best; }
  Object.assign(known, SEED);
  for (let it = 0; it < 2; it++) {
    for (const [cs, y] of multi) {
      const unk = cs.filter(c => !isKana(c) && !(key(c) in known));
      if (unk.length !== 1) continue;
      const pat = cs.map(c => isKana(c) ? esc(hira(c)) : (key(c) in known ? esc(known[key(c)]) : '(.+)')).join('');
      const m = y.match(new RegExp('^' + pat + '$', 'u'));
      if (m) known[key(unk[0])] = m[1];
    }
  }
  return known;
}

export class Yomi {
  // oldYomi: {nz(語): 読み}、tokenizer: kuromoji（無ければ漢字はそのまま）
  constructor(oldYomi, tokenizer) {
    this.old = oldYomi; this.tk = tokenizer;
    const pairs = [];
    for (const [k, v] of Object.entries(oldYomi)) {
      if (k.startsWith('(')) continue;
      const t = N(k.replace(/[（(].*$/, '')) || k; const y = clean(v);
      if (t && y) pairs.push([t, y]);
    }
    this.D = learn(pairs);
    this.keys = Object.keys(this.D).filter(k => [...k].length >= 2).sort((a, b) => [...b].length - [...a].length);
  }
  morph(t) {
    if (!this.tk) return clean(t);
    return clean(this.tk.tokenize(t).map(x => (x.reading && x.reading !== '*') ? x.reading : x.surface_form).join(''));
  }
  baseY(t) {
    let s = '';
    for (const ch of N(t)) {
      if (/[A-Za-z]/.test(ch)) s += AL[ch.toLowerCase()];
      else if (ch in NUM) s += NUM[ch];
      else s += ch;
    }
    return this.morph(s);
  }
  dictY(t) {
    const c = N(t).replace(/[・\s()（）]/g, ''); let out = ''; const guessed = [];
    const D = this.D, keys = this.keys;
    const keyAt = i => keys.find(k => c.startsWith(k, i));
    let i = 0;
    while (i < c.length) {
      const hit = keyAt(i);
      if (hit) { out += D[hit]; i += hit.length; continue; }
      if (isKana(c[i])) { out += hira(c[i]); i++; continue; }
      const m = c.slice(i).match(/^\d+/);
      if (m) { out += numY(m[0]); i += m[0].length; continue; }
      let j = i + 1;
      while (j < c.length && !keyAt(j) && !isKana(c[j]) && !/\d/.test(c[j])) j++;
      const seg = c.slice(i, j);
      if (seg in D) out += D[seg];
      else { out += this.baseY(seg); if (/[一-龥A-Za-z]/.test(seg)) guessed.push(seg); }
      i = j;
    }
    return [out, guessed];
  }
  get(text) {
    const k = nz(text);
    if (/^○?未?請求/.test(k) && k.includes('請求') && k.length <= 4) return ['まるみせいきゅう', []];
    if (k.startsWith('○') && k.includes('請求')) return ['まるみせいきゅう', []];
    if (k in this.old) return [clean(this.old[k]), []];
    const base = N(text).replace(/[（(].*$/, '') || N(text);
    return this.dictY(base);
  }
}
