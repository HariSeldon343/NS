// Converts **marked** segments to LinkedIn Unicode Mathematical Sans-Serif Bold,
// in place, then reports the raw codepoint count of each POST body (hook+body+CTA,
// excluding metadata header lines and hashtags). Long-form (POST 3) must be <= 2800.
// Usage: node tools/finalize_md.mjs <file.md>
import { readFileSync, writeFileSync } from 'node:fs';

const file = process.argv[2];
if (!file) { console.error('missing file'); process.exit(1); }
let txt = readFileSync(file, 'utf8');

const boldChar = (c) => {
  const cp = c.codePointAt(0);
  if (cp >= 48 && cp <= 57) return String.fromCodePoint(0x1D7EC + (cp - 48));   // 0-9
  if (cp >= 65 && cp <= 90) return String.fromCodePoint(0x1D5D4 + (cp - 65));   // A-Z
  if (cp >= 97 && cp <= 122) return String.fromCodePoint(0x1D5EE + (cp - 97));  // a-z
  return c; // leave accents, punctuation, spaces untouched
};

txt = txt.replace(/\*\*([\s\S]+?)\*\*/g, (_, p) => [...p].map(boldChar).join(''));
writeFileSync(file, txt);

const lines = txt.split('\n');
const isHeader = (l) => /^##\s+POST\s+(\d)/i.exec(l.trim());
const isMeta = (l) => /^(Tema|Fonte|URL|Orario|Lunghezza|Settore SCO|Titolo)\s*:/i.test(l.trim());
const isHashtags = (l) => {
  const t = l.trim();
  if (!t) return false;
  return t.split(/\s+/).every((w) => w.startsWith('#'));
};

let cur = null, body = [];
const flush = () => {
  if (cur === null) return;
  const text = body.join('\n').trim();
  const count = [...text].length;
  const limit = cur === '3' ? '  (limite 2800)' : '';
  const flag = cur === '3' && count > 2800 ? '  *** OVER LIMIT ***' : '';
  console.log(`POST ${cur}: ${count} caratteri raw${limit}${flag}`);
};
for (const l of lines) {
  const h = isHeader(l);
  if (h) { flush(); cur = h[1]; body = []; continue; }
  if (cur === null) continue;
  if (/^##\s/.test(l.trim())) { flush(); cur = null; continue; } // left posts (e.g. Note)
  if (isMeta(l) || isHashtags(l)) continue;
  body.push(l);
}
flush();
