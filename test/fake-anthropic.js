'use strict';
// A stand-in for the Anthropic Messages API, for tests only. Records requests and streams a fixed answer.
const http = require('http');
const requests = [];
const server = http.createServer((req, res) => {
  let b = ''; req.on('data', (c) => { b += c; }); req.on('end', () => {
    const body = JSON.parse(b || '{}'); requests.push({ headers: req.headers, body });
    if (req.headers['x-api-key'] !== 'test-key') { res.writeHead(401, { 'content-type': 'application/json' }); return res.end(JSON.stringify({ error: { message: 'invalid x-api-key' } })); }
    const text = body.system && /JSON/.test(body.system) ? '[{"date":"2026-03-04","fuel":"HVO100","product":"HVO100","liters":412.5,"unit_price":21.6,"amount":8910,"amount_includes_vat":false,"vehicle":null,"supplier":"Test AB","file":"kvitto.png","confidence":0.9}]' : 'Lastväxlare 3-axl står för störst andel av utsläppen. Svar från testservern.';
    if (!body.stream) { res.writeHead(200, { 'content-type': 'application/json' }); return res.end(JSON.stringify({ content: [{ type: 'text', text }] })); }
    res.writeHead(200, { 'content-type': 'text/event-stream' });
    res.write('event: message_start\ndata: {"type":"message_start","message":{"id":"m1"}}\n\n');
    res.write('event: content_block_start\ndata: {"type":"content_block_start","index":0,"content_block":{"type":"text","text":""}}\n\n');
    for (const part of text.match(/.{1,12}/g)) res.write(`event: content_block_delta\ndata: ${JSON.stringify({ type: 'content_block_delta', index: 0, delta: { type: 'text_delta', text: part } })}\n\n`);
    res.write('event: content_block_stop\ndata: {"type":"content_block_stop","index":0}\n\nevent: message_stop\ndata: {"type":"message_stop"}\n\n');
    res.end();
  });
});
module.exports = { start: (port) => new Promise((r) => server.listen(port, '127.0.0.1', r)), stop: () => server.close(), requests };
