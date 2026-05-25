// Publer publisher for the LinkedIn automation.
//   node tools/publish_publer.mjs --check
//       -> read-only: lists connected accounts (validates key + connectivity)
//   node tools/publish_publer.mjs <manifest.json>
//       -> uploads media and creates SCHEDULED posts, then reconciles the queue
//
// Reads the API key from env PUBLER_API_KEY. Manifest item:
//   { "account_id": "...", "text": "...", "media_path": "/abs/card.png",
//     "scheduled_at": "2026-05-26T10:00:00+02:00" }
import { readFileSync } from 'node:fs';
import { basename } from 'node:path';

const KEY = process.env.PUBLER_API_KEY;
const WSID = process.env.PUBLER_WORKSPACE_ID || '6a13de8cadf35ab5536f6cf6';
const BASE = 'https://app.publer.com/api/v1';
if (!KEY) { console.error('Missing PUBLER_API_KEY'); process.exit(1); }

// Browser-like UA: Publer sits behind Cloudflare, which blocks default
// library user-agents (Python urllib got error 1010). Node + this UA passes.
const UA = 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/141.0.0.0 Safari/537.36';
const authHeaders = (extra = {}) => ({
  'Authorization': `Bearer-API ${KEY}`,
  'Publer-Workspace-Id': WSID,
  'User-Agent': UA,
  'Accept': 'application/json',
  ...extra,
});

async function listAccounts() {
  const r = await fetch(`${BASE}/accounts`, { headers: authHeaders() });
  if (!r.ok) throw new Error(`/accounts HTTP ${r.status}: ${await r.text()}`);
  return r.json();
}

async function uploadMedia(path) {
  const fd = new FormData();
  fd.append('file', new Blob([readFileSync(path)]), basename(path));
  const r = await fetch(`${BASE}/media`, { method: 'POST', headers: authHeaders(), body: fd });
  if (!r.ok) throw new Error(`/media HTTP ${r.status}: ${await r.text()}`);
  let d = await r.json();
  if (Array.isArray(d)) d = d[0] || {};
  const id = d.id || (Array.isArray(d.media) && d.media[0] && d.media[0].id);
  if (!id) throw new Error(`no media id in response: ${JSON.stringify(d).slice(0, 200)}`);
  return id;
}

async function schedule(manifest) {
  const posts = [];
  for (const it of manifest) {
    const mediaId = await uploadMedia(it.media_path);
    posts.push({
      networks: { linkedin: { type: 'photo', text: it.text, media: [{ id: mediaId, type: 'image' }] } },
      accounts: [{ id: it.account_id, scheduled_at: it.scheduled_at }],
    });
    console.log(`  uploaded ${basename(it.media_path)} -> ${mediaId} | ${it.scheduled_at}`);
  }
  const payload = { bulk: { state: 'scheduled', posts } };
  const r = await fetch(`${BASE}/posts/schedule`, {
    method: 'POST', headers: authHeaders({ 'Content-Type': 'application/json' }),
    body: JSON.stringify(payload),
  });
  const body = await r.text();
  console.log('schedule HTTP', r.status, body.slice(0, 200));
  const jobId = JSON.parse(body).job_id;
  if (jobId) {
    for (let i = 0; i < 8; i++) {
      await new Promise((s) => setTimeout(s, 2000));
      const jr = await fetch(`${BASE}/job_status/${jobId}`, { headers: authHeaders() });
      const jb = await jr.text();
      console.log(`  [job ${i + 1}] ${jb.slice(0, 200)}`);
      if (/"status":"(complete|completed|failed|error)"/i.test(jb)) break;
    }
  }
  // Reconcile: the bulk job can report transient failures; the queue is the source of truth.
  const sched = await (await fetch(`${BASE}/posts?state=scheduled`, { headers: authHeaders() })).json();
  const list = sched.posts || sched;
  console.log(`RECONCILE: ${list.length} post in coda (scheduled)`);
  return list;
}

const arg = process.argv[2];
if (arg === '--check') {
  const accts = await listAccounts();
  console.log('Account collegati:');
  for (const a of accts) console.log(`  ${a.provider}/${a.type} | ${a.name} | id=${a.id}`);
} else if (arg) {
  await schedule(JSON.parse(readFileSync(arg, 'utf8')));
} else {
  console.error('Usage: --check | <manifest.json>'); process.exit(1);
}
