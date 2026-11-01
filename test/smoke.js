'use strict';
// npm test — starts the demo server against a fake Anthropic API and in mock mode. No network, no real key.
const { spawn } = require('child_process');
const path = require('path');
const assert = require('assert/strict');
const fake = require('./fake-anthropic');

const root = path.join(__dirname, '..');
const wait = (ms) => new Promise((r) => setTimeout(r, ms));
function startServer(port, env) {
  const p = spawn(process.execPath, ['server.js'], { cwd: root, env: Object.assign({}, process.env, { PORT: String(port), HOST: '127.0.0.1', ANTHROPIC_API_KEY: '', DEMO_PASSWORD: '' }, env), stdio: ['ignore', 'pipe', 'pipe'] });
  let out = ''; p.stdout.on('data', (d) => { out += d; }); p.stderr.on('data', (d) => { out += d; });
  return { p, log: () => out };
}
async function up(port) { for (let i = 0; i < 50; i++) { try { await fetch(`http://127.0.0.1:${port}/api/status`); return; } catch (e) { await wait(100); } } throw new Error('server did not start'); }
const ai = (port, body) => fetch(`http://127.0.0.1:${port}/api/ai`, { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify(body) });

(async () => {
  await fake.start(4599);
  const live = startServer(4601, { ANTHROPIC_API_KEY: 'test-key', ANTHROPIC_BASE_URL: 'http://127.0.0.1:4599', ANTHROPIC_MODEL: 'claude-sonnet-5-5' });
  const mock = startServer(4602, {});
  const bad = startServer(4603, { ANTHROPIC_API_KEY: 'wrong', ANTHROPIC_BASE_URL: 'http://127.0.0.1:4599' });
  try {
    await Promise.all([up(4601), up(4602), up(4603)]);

    assert.deepEqual(await (await fetch('http://127.0.0.1:4601/api/status')).json(), { ai: 'live', model: 'claude-sonnet-5-5' });
    assert.equal((await (await fetch('http://127.0.0.1:4602/api/status')).json()).ai, 'mock');

    // streaming text, chat turns merged, image attached to the last user turn
    const r1 = await ai(4601, { mode: 'text', stream: true, input: [{ role: 'user', content: 'RULES' }, { role: 'user', content: 'Vilken fordonstyp?' }], images: [{ mediaType: 'image/png', data: 'iVBORw0KGgo=' }] });
    assert.equal(r1.status, 200);
    assert.match(await r1.text(), /Lastväxlare 3-axl står för störst andel/);
    const sent = fake.requests[fake.requests.length - 1];
    assert.equal(sent.headers['anthropic-version'], '2023-06-01');
    assert.equal(sent.body.model, 'claude-sonnet-5-5');
    assert.equal(sent.body.messages.length, 1, 'consecutive user turns merged');
    assert.equal(sent.body.messages[0].content[0].type, 'image');
    assert.equal(sent.body.messages[0].content[1].text, 'RULES\n\nVilken fordonstyp?');

    // JSON mode (receipt reading) without streaming
    const r2 = await ai(4601, { mode: 'json', stream: false, input: 'Read these receipts' });
    assert.equal(JSON.parse(await r2.text())[0].liters, 412.5);
    assert.match(fake.requests[fake.requests.length - 1].body.system, /JSON/);

    // a wrong key gives a clear message, not a crash
    const r3 = await ai(4603, { mode: 'text', input: 'hej' });
    assert.equal(r3.status, 502);
    assert.match((await r3.json()).message, /API key was not accepted/);

    // bad input is refused
    assert.equal((await ai(4601, { mode: 'text', input: [{ role: 'assistant', content: 'x' }] })).status, 400);

    // mock mode streams a labelled demo answer
    const r4 = await ai(4602, { mode: 'text', stream: true, input: 'Du skriver sammanfattningen i en klimatrapport som ...' });
    assert.match(await r4.text(), /Demoläge/);

    // static files, and no way out of public/
    const idx = await fetch('http://127.0.0.1:4602/');
    assert.match(await idx.text(), /Spårbar Klimatportal \(demo\)/);
    assert.equal((await fetch('http://127.0.0.1:4602/..%2fserver.js')).status, 403);

    console.log('PASS: live streaming, JSON mode, wrong key, bad input, mock mode, static files, path traversal');
  } catch (e) {
    console.error('FAIL:', e.message); console.error(live.log()); process.exitCode = 1;
  } finally {
    live.p.kill(); mock.p.kill(); bad.p.kill(); fake.stop();
  }
})();
