// まとめてダウンロード用のzip。ファイル名はWindowsの標準（Shift_JIS）で入れ、ほかのソフト向けにUTF-8の名前も添える
let SJIS = null;
function sjisTable() {
  if (SJIS) return SJIS;
  SJIS = new Map();
  const dec = new TextDecoder('shift_jis');
  for (let b = 0xA1; b <= 0xDF; b++) SJIS.set(dec.decode(new Uint8Array([b])), [b]);
  const leads = [];
  for (let b = 0x81; b <= 0x9F; b++) leads.push(b);
  for (let b = 0xE0; b <= 0xFC; b++) leads.push(b);
  for (const l of leads) for (let t = 0x40; t <= 0xFC; t++) {
    if (t === 0x7F) continue;
    const ch = dec.decode(new Uint8Array([l, t]));
    if (ch.length === 1 && ch !== '�' && !SJIS.has(ch)) SJIS.set(ch, [l, t]);
  }
  return SJIS;
}
export function toSjis(name) {
  const T = sjisTable(); const out = [];
  for (const ch of name) {
    const c = ch.codePointAt(0);
    if (c < 0x80) out.push(c);
    else if (T.has(ch)) out.push(...T.get(ch));
    else out.push(0x5F);   // Shift_JISに無い字は「_」
  }
  return new Uint8Array(out);
}
// files: [{name, buf}]
export async function makeZip(JSZip, files) {
  const z = new JSZip();
  // zipの日時は時差なしで入るので、日本時間に合わせる
  const now = new Date(); const local = new Date(now.getTime() - now.getTimezoneOffset() * 60000);
  for (const f of files) z.file(f.name, f.buf, { date: local });
  return await z.generateAsync({ type: 'uint8array', compression: 'DEFLATE', encodeFileName: toSjis });
}
