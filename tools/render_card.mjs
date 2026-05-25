// Card renderer 1080x1080 for LinkedIn workflow (Amodeo / S.Co Solution Consulting).
// Usage: node tools/render_card.mjs <spec.json>
// spec.json = array of cards: { brand, theme, label, grad1, grad2, title, footer, out, fontSize? }
//   brand: "amodeo" | "sco"
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { createRequire } from 'node:module';

const require = createRequire(import.meta.url);
let chromium;
try { ({ chromium } = require('/opt/node22/lib/node_modules/playwright')); }
catch { ({ chromium } = require('playwright')); }

const specPath = process.argv[2];
if (!specPath) { console.error('Missing spec path'); process.exit(1); }
const cards = JSON.parse(readFileSync(specPath, 'utf8'));

const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

function brandHeader(brand) {
  if (brand === 'sco') {
    return `<div class="wm"><span class="wm-strong">S.Co</span><span class="wm-light">Solution Consulting</span></div>`;
  }
  return `<div class="wm"><span class="wm-strong">Antonio S. Amodeo</span><span class="wm-light">SCO Consulting · Fortibyte</span></div>`;
}
function brandFooterRight(brand) {
  if (brand === 'sco') return `<div class="claim">Dare to change</div>`;
  return `<div class="claim claim-domain">fortibyte.it</div>`;
}

function html(card) {
  const fs = card.fontSize || (card.title.length > 90 ? 52 : card.title.length > 55 ? 60 : 70);
  return `<!doctype html><html><head><meta charset="utf-8"><style>
  * { margin:0; padding:0; box-sizing:border-box; }
  html,body { width:1080px; height:1080px; }
  body {
    font-family:'Liberation Sans','DejaVu Sans','FreeSans',sans-serif;
    color:#f8fafc;
    background:
      radial-gradient(900px 620px at 82% 12%, ${card.grad2}33, transparent 60%),
      radial-gradient(760px 760px at 8% 96%, ${card.grad1}26, transparent 62%),
      linear-gradient(150deg, #0b1220 0%, #0f172a 52%, #111827 100%);
    position:relative; overflow:hidden;
  }
  .frame { position:absolute; inset:54px; display:flex; flex-direction:column; }
  .top { display:flex; align-items:flex-start; justify-content:space-between; }
  .wm { display:flex; flex-direction:column; line-height:1.1; }
  .wm-strong { font-size:38px; font-weight:700; letter-spacing:.3px; color:#ffffff; }
  .wm-light { font-size:21px; font-weight:400; color:#94a3b8; margin-top:4px; letter-spacing:.4px; }
  .badge {
    display:inline-flex; align-items:center; align-self:flex-start;
    padding:14px 26px; border-radius:999px;
    background:linear-gradient(90deg, ${card.grad1}, ${card.grad2});
    font-size:23px; font-weight:700; letter-spacing:2.4px; text-transform:uppercase; color:#ffffff;
    box-shadow:0 10px 30px ${card.grad1}40;
    margin-top:90px;
  }
  .mid { flex:1; display:flex; flex-direction:column; justify-content:center; }
  .title { font-size:${fs}px; font-weight:700; line-height:1.16; letter-spacing:-.4px; max-width:980px; }
  .accent { width:120px; height:7px; border-radius:6px; margin-top:40px;
    background:linear-gradient(90deg, ${card.grad1}, ${card.grad2}); }
  .bottom { display:flex; align-items:flex-end; justify-content:space-between; }
  .footer { font-size:22px; color:#9aa6b8; max-width:680px; line-height:1.35; }
  .claim { font-size:30px; font-style:italic; font-weight:600; color:#e2e8f0; opacity:.92; }
  .claim-domain { font-style:normal; font-weight:600; color:#cbd5e1; letter-spacing:.5px; }
  </style></head><body>
    <div class="frame">
      <div class="top">${brandHeader(card.brand)}</div>
      <span class="badge">${esc(card.label)}</span>
      <div class="mid">
        <div class="title">${esc(card.title)}</div>
        <div class="accent"></div>
      </div>
      <div class="bottom">
        <div class="footer">${esc(card.footer || '')}</div>
        ${brandFooterRight(card.brand)}
      </div>
    </div>
  </body></html>`;
}

const browser = await chromium.launch({ args: ['--no-sandbox'] });
const page = await browser.newPage({ viewport: { width: 1080, height: 1080 }, deviceScaleFactor: 1 });
for (const card of cards) {
  await page.setContent(html(card), { waitUntil: 'networkidle' });
  await page.waitForTimeout(1600);
  const out = resolve(card.out);
  await page.screenshot({ path: out, type: 'png' });
  console.log('rendered ->', out);
}
await browser.close();
