'use strict';
/**
 * Spårbar demo server. No dependencies: Node 20+ only.
 *   - Serves the portal from ./public
 *   - POST /api/ai calls Claude with YOUR Anthropic API key, which stays on this machine
 *   - Without a key it runs in mock mode, so the demo still works offline
 * Settings come from .env (see .env.example).
 */
const http = require('http');
const fs = require('fs');
const path = require('path');

// ---------- settings ----------
(function loadEnv() {
  const f = path.join(__dirname, '.env');
  if (!fs.existsSync(f)) return;
  for (const line of fs.readFileSync(f, 'utf8').split(/\r?\n/)) {
    const m = /^\s*([A-Z0-9_]+)\s*=\s*(.*)\s*$/.exec(line);
    if (m && !(m[1] in process.env)) process.env[m[1]] = m[2].replace(/^["']|["']$/g, '');
  }
})();
const PORT = Number(process.env.PORT) || 3000;
const HOST = process.env.HOST || '127.0.0.1';
const KEY = (process.env.ANTHROPIC_API_KEY || '').trim();
const MODEL = process.env.ANTHROPIC_MODEL || 'claude-sonnet-5-5';
const API_BASE = (process.env.ANTHROPIC_BASE_URL || 'https://api.anthropic.com').replace(/\/$/, '');
const WORKSPACE = (process.env.ANTHROPIC_WORKSPACE_ID || '').trim(); // only for keys not scoped to a workspace
const MAX_TOKENS = Number(process.env.MAX_TOKENS) || 1500;
const PASSWORD = process.env.DEMO_PASSWORD || '';
const CALLS_PER_HOUR = Number(process.env.MAX_AI_CALLS_PER_HOUR) || 120;
const PUBLIC = path.join(__dirname, 'public');
const MODE = KEY ? 'live' : 'mock';

const MIME = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.css': 'text/css; charset=utf-8', '.json': 'application/json', '.svg': 'image/svg+xml', '.png': 'image/png', '.ico': 'image/x-icon' };
const IMAGE_TYPES = new Set(['image/jpeg', 'image/png', 'image/webp', 'image/gif']);

// ---------- helpers ----------
const sendJson = (res, status, obj) => { res.writeHead(status, { 'content-type': 'application/json; charset=utf-8' }); res.end(JSON.stringify(obj)); };
function readBody(req, limit) {
  return new Promise((resolve, reject) => {
    let size = 0; const chunks = [];
    req.on('data', (c) => { size += c.length; if (size > limit) { reject(Object.assign(new Error('too large'), { status: 413 })); req.destroy(); } else chunks.push(c); });
    req.on('end', () => resolve(Buffer.concat(chunks).toString('utf8')));
    req.on('error', reject);
  });
}
function authorised(req) {
  if (!PASSWORD) return true;
  const h = req.headers.authorization || '';
  if (!h.startsWith('Basic ')) return false;
  const pass = Buffer.from(h.slice(6), 'base64').toString('utf8').split(':').slice(1).join(':');
  return pass === PASSWORD;
}
const calls = [];
function withinBudget() {
  const hourAgo = Date.now() - 3600e3;
  while (calls.length && calls[0] < hourAgo) calls.shift();
  if (calls.length >= CALLS_PER_HOUR) return false;
  calls.push(Date.now()); return true;
}

/** Turn the portal's input (a prompt or user/assistant turns) into Messages API messages. */
function toMessages(input, images) {
  let turns = typeof input === 'string' ? [{ role: 'user', content: input }] : Array.isArray(input) ? input : [];
  turns = turns.filter((t) => t && (t.role === 'user' || t.role === 'assistant') && typeof t.content === 'string' && t.content.trim());
  const merged = [];
  for (const t of turns) {
    const last = merged[merged.length - 1];
    if (last && last.role === t.role) last.content += '\n\n' + t.content; else merged.push({ role: t.role, content: t.content });
  }
  if (!merged.length || merged[0].role !== 'user' || merged[merged.length - 1].role !== 'user') throw Object.assign(new Error('Input must start and end with a user turn.'), { status: 400, code: 'invalid_request' });
  const last = merged[merged.length - 1];
  const imgs = (images || []).slice(0, 5).filter((i) => i && IMAGE_TYPES.has(i.mediaType) && typeof i.data === 'string');
  last.content = imgs.map((i) => ({ type: 'image', source: { type: 'base64', media_type: i.mediaType, data: i.data } })).concat([{ type: 'text', text: last.content }]);
  return merged;
}

// ---------- Claude ----------
async function callClaude(body, res, signal) {
  const messages = toMessages(body.input, body.images);
  const payload = { model: MODEL, max_tokens: MAX_TOKENS, messages, stream: !!body.stream };
  if (body.mode === 'json') payload.system = 'Reply with only valid JSON, no prose and no code fences.';
  const r = await fetch(API_BASE + '/v1/messages', {
    method: 'POST', signal,
    headers: Object.assign({ 'x-api-key': KEY, 'anthropic-version': '2023-06-01', 'content-type': 'application/json' }, WORKSPACE ? { 'anthropic-workspace-id': WORKSPACE } : {}),
    body: JSON.stringify(payload)
  });
  if (!r.ok) {
    const detail = await r.json().catch(() => ({}));
    const msg = (detail.error && detail.error.message) || r.statusText;
    const code = r.status === 429 || r.status === 529 ? 'rate_limited' : r.status === 400 ? 'invalid_request' : 'upstream_error';
    console.error(`[ai] Anthropic ${r.status}: ${msg}`);
    return sendJson(res, r.status === 401 ? 502 : r.status, { code, message: r.status === 401 ? 'The API key was not accepted. Check ANTHROPIC_API_KEY in .env.' : msg });
  }
  if (!body.stream) {
    const j = await r.json();
    res.writeHead(200, { 'content-type': 'text/plain; charset=utf-8' });
    return res.end((j.content || []).filter((b) => b.type === 'text').map((b) => b.text).join(''));
  }
  res.writeHead(200, { 'content-type': 'text/plain; charset=utf-8', 'cache-control': 'no-cache', 'x-accel-buffering': 'no' });
  const reader = r.body.getReader(), dec = new TextDecoder();
  let buf = '';
  for (;;) {
    const { done, value } = await reader.read();
    if (done) break;
    buf += dec.decode(value, { stream: true });
    let i;
    while ((i = buf.indexOf('\n')) >= 0) {
      const line = buf.slice(0, i).trim(); buf = buf.slice(i + 1);
      if (!line.startsWith('data:')) continue;
      let ev; try { ev = JSON.parse(line.slice(5)); } catch (e) { continue; }
      if (ev.type === 'content_block_delta' && ev.delta && ev.delta.type === 'text_delta') res.write(ev.delta.text);
      else if (ev.type === 'error') console.error('[ai] stream error:', ev.error && ev.error.message);
    }
  }
  res.end();
}

// ---------- mock mode (no API key) ----------
const MOCK = {
  sv: {
    summary: 'Under perioden utförde vi transporterna åt er med i huvudsak lastväxlare och kranbilar. Klimatpåverkan från källa till hjul framgår av rapportens nyckeltal, och fordonen körde på HVO100.\n\nSiffrorna kan användas direkt i er klimatredovisning under Scope 3, kategori 4. [Demoläge: lägg in en API-nyckel för att låta Claude skriva texten utifrån era siffror.]',
    explain: 'Kontrollen visar en avvikelse i underlaget som påverkar rapportens siffror.\n- Kontrollera raderna bakom kontrollen.\n- Rätta i källsystemet eller justera inställningen.\n- Kör beräkningen igen innan rapporten låses.\n[Demoläge utan API-nyckel.]',
    chat: 'Det här är ett demosvar. Med en API-nyckel svarar Claude utifrån rapportens siffror, till exempel vilka fordonstyper som står för mest utsläpp. [Demoläge utan API-nyckel.]'
  },
  en: {
    summary: 'During the period we carried out your transports mainly with hook-lift trucks and crane trucks. The well-to-wheel climate impact is shown in the key figures, and the vehicles ran on HVO100.\n\nThe figures can be used directly in your GHG accounts under Scope 3, Category 4. [Demo mode: add an API key to let Claude write this from your figures.]',
    explain: 'The check points to an irregularity in the data that affects the report.\n- Look at the rows behind the check.\n- Fix the source system or adjust the setting.\n- Recalculate before locking the report.\n[Demo mode without an API key.]',
    chat: 'This is a demo answer. With an API key, Claude answers from the report\'s figures, for example which vehicle types cause the most emissions. [Demo mode without an API key.]'
  }
};
async function mock(body, res) {
  if (body.mode === 'json') { res.writeHead(200, { 'content-type': 'text/plain; charset=utf-8' }); return res.end('[]'); }
  const all = typeof body.input === 'string' ? body.input : (body.input || []).map((t) => t.content).join('\n');
  const lastQ = Array.isArray(body.input) ? String((body.input[body.input.length - 1] || {}).content || '') : all;
  const lang = /sammanfattningen|KONTROLL \(|[åäö]/.test(Array.isArray(body.input) ? lastQ : all) ? 'sv' : 'en';
  const kind = /summary paragraph|sammanfattningen i en klimatrapport/.test(all) ? 'summary' : /CHECK \(|KONTROLL \(/.test(all) ? 'explain' : 'chat';
  const text = MOCK[lang][kind];
  res.writeHead(200, { 'content-type': 'text/plain; charset=utf-8', 'cache-control': 'no-cache' });
  if (!body.stream) return res.end(text);
  for (const word of text.split(/(?<=\s)/)) { res.write(word); await new Promise((r) => setTimeout(r, 18)); }
  res.end();
}

// ---------- server ----------
const server = http.createServer(async (req, res) => {
  res.setHeader('x-content-type-options', 'nosniff');
  res.setHeader('referrer-policy', 'no-referrer');
  if (!authorised(req)) { res.writeHead(401, { 'www-authenticate': 'Basic realm="Sparbar demo"' }); return res.end('Password required'); }
  const url = new URL(req.url, 'http://localhost');

  if (url.pathname === '/api/status') return sendJson(res, 200, { ai: MODE, model: MODE === 'live' ? MODEL : null });

  if (url.pathname === '/api/ai' && req.method === 'POST') {
    let body;
    try { body = JSON.parse(await readBody(req, 25 * 1024 * 1024)); } catch (e) { return sendJson(res, e.status || 400, { code: 'invalid_request', message: e.status === 413 ? 'Request too large' : 'Bad JSON' }); }
    if (!withinBudget()) return sendJson(res, 429, { code: 'rate_limited', message: `More than ${CALLS_PER_HOUR} AI calls in the last hour.` });
    const started = Date.now();
    const ctl = new AbortController();
    res.on('close', () => { if (!res.writableEnded) ctl.abort(); });
    try {
      if (MODE === 'mock') await mock(body, res); else await callClaude(body, res, ctl.signal);
      console.log(`[ai] ${MODE} ${body.mode || 'text'}${body.stream ? ' stream' : ''} ${body.images && body.images.length ? body.images.length + ' image(s) ' : ''}${Date.now() - started} ms`);
    } catch (e) {
      if (ctl.signal.aborted) return;
      console.error('[ai] failed:', e.message);
      if (!res.headersSent) sendJson(res, e.status || 502, { code: e.code || 'upstream_error', message: e.message }); else res.end();
    }
    return;
  }

  if (req.method !== 'GET' && req.method !== 'HEAD') { res.writeHead(405); return res.end(); }
  const rel = decodeURIComponent(url.pathname === '/' ? '/index.html' : url.pathname);
  const file = path.normalize(path.join(PUBLIC, rel));
  if (!file.startsWith(PUBLIC + path.sep)) { res.writeHead(403); return res.end(); }
  fs.readFile(file, (err, buf) => {
    if (err) { res.writeHead(404); return res.end('Not found'); }
    res.writeHead(200, { 'content-type': MIME[path.extname(file)] || 'application/octet-stream', 'cache-control': 'no-cache' });
    res.end(buf);
  });
});

server.listen(PORT, HOST, () => {
  console.log(`Spårbar demo on http://${HOST === '0.0.0.0' ? 'localhost' : HOST}:${PORT}`);
  console.log(MODE === 'live' ? `AI: live, model ${MODEL}, max ${CALLS_PER_HOUR} calls/hour` : 'AI: mock mode (no ANTHROPIC_API_KEY in .env)');
  if (HOST !== '127.0.0.1' && !PASSWORD) console.log('Warning: reachable from other machines without DEMO_PASSWORD.');
});
