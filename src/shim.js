/* Spårbar demo: stands in for the Claude artifact runtime when the portal runs on its own.
 *   sample    -> POST api/ai on the demo server (Claude via your API key; the key never reaches the browser)
 *   downloads -> normal browser downloads
 *   db        -> this browser's localStorage (archive, receipts, settings; not shared between people)
 *   user      -> a single local demo user
 * Inside Claude the real runtime exists and this file does nothing. */
(function () {
  'use strict';
  if (window.claude && typeof window.claude.use === 'function') return;
  // Bundled copies of the export and PDF libraries, so the demo works without CDNs or Wi-Fi.
  window.SPARBAR_LIBS = { xlsx: 'vendor/xlsx.full.min.js', jspdf: 'vendor/jspdf.umd.min.js', autotable: 'vendor/jspdf.plugin.autotable.min.js', pdfjs: 'vendor/pdf.min.js', pdfworker: 'vendor/pdf.worker.min.js' };

  // ---------- db (localStorage) ----------
  const KEY = 'sparbar-demo-db';
  let data = {};
  try { data = JSON.parse(localStorage.getItem(KEY) || '{}') || {}; } catch (e) { data = {}; }
  const persist = () => { try { localStorage.setItem(KEY, JSON.stringify(data)); } catch (e) { /* storage full or blocked: keep in memory */ } };
  const listeners = new Set();
  const notify = () => listeners.forEach((fn) => { try { fn(); } catch (e) { console.error(e); } });
  const meta = Object.freeze({ fromCache: false, hasPendingWrites: false });
  const copy = (v) => JSON.parse(JSON.stringify(v));
  const snap = (path) => ({ id: path.split('/').pop(), exists: Object.prototype.hasOwnProperty.call(data, path), data: () => (path in data ? copy(data[path]) : undefined), metadata: meta });

  function docRef(path) {
    return {
      id: path.split('/').pop(), path,
      async get() { return snap(path); },
      async set(v) { data[path] = copy(v); persist(); notify(); },
      async update(v) { if (!(path in data)) throw { code: 'invalid_argument', message: 'Document does not exist' }; Object.assign(data[path], copy(v)); persist(); notify(); },
      async delete() { delete data[path]; persist(); notify(); },
      onSnapshot(next) { const fn = () => next(snap(path)); listeners.add(fn); setTimeout(fn, 0); return () => listeners.delete(fn); },
      collection(sub) { return collRef(path + '/' + sub); }
    };
  }
  function collRef(path, q) {
    q = q || {};
    const list = () => {
      let docs = Object.keys(data).filter((k) => k.startsWith(path + '/') && !k.slice(path.length + 1).includes('/')).map(snap);
      if (q.order) {
        const [f, dir] = q.order;
        docs.sort((a, b) => { const x = a.data()[f], y = b.data()[f]; return (x < y ? -1 : x > y ? 1 : 0) * (dir === 'desc' ? -1 : 1); });
      } else docs.sort((a, b) => (a.id < b.id ? -1 : 1));
      if (q.limit) docs = docs.slice(0, q.limit);
      return { docs, size: docs.length, empty: !docs.length, docChanges: () => docs.map((d, i) => ({ type: 'added', doc: d, oldIndex: -1, newIndex: i })), metadata: meta };
    };
    return {
      path,
      doc(id) { return docRef(path + '/' + (id || Date.now().toString(36) + Math.random().toString(36).slice(2, 8))); },
      async add(v) { const r = this.doc(); await r.set(v); return r; },
      where() { return this; },
      orderBy(f, dir) { return collRef(path, Object.assign({}, q, { order: [f, dir || 'asc'] })); },
      limit(n) { return collRef(path, Object.assign({}, q, { limit: n })); },
      async get() { return list(); },
      onSnapshot(next) { const fn = () => next(list()); listeners.add(fn); setTimeout(fn, 0); return () => listeners.delete(fn); }
    };
  }
  const db = { doc: docRef, collection: (p) => collRef(p) };

  // ---------- downloads ----------
  const downloads = {
    async save({ filename, data: payload }) {
      const blob = payload instanceof Blob ? payload : new Blob([payload]);
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a'); a.href = url; a.download = filename; document.body.appendChild(a); a.click(); a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 5000);
      return { status: 'saved' };
    }
  };

  // ---------- user ----------
  const user = {
    async id() { return 'demo-user'; }, async me() { return { id: 'demo-user', name: 'Demo' }; },
    async profiles(ids) { const o = {}; ids.forEach((i) => { o[i] = { id: i, name: 'Demo' }; }); return o; },
    isOwner: () => true, canEdit: () => true, can: () => true
  };

  // ---------- sample (Claude via the demo server) ----------
  const toB64 = (b) => new Promise((resolve, reject) => { const r = new FileReader(); r.onload = () => resolve(String(r.result).split(',')[1]); r.onerror = reject; r.readAsDataURL(b); });
  async function call(input, opts, mode) {
    opts = opts || {};
    const list = opts.images ? (opts.images instanceof Blob ? [opts.images] : Array.from(opts.images)) : [];
    const images = await Promise.all(list.map(async (b) => ({ mediaType: b.type || 'image/png', data: await toB64(b) })));
    const stream = !!opts.onText && mode === 'text';
    let res;
    try {
      res = await fetch('api/ai', { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ input, images, mode, stream }), signal: opts.signal });
    } catch (e) {
      throw { code: e && e.name === 'AbortError' ? 'cancelled' : 'upstream_error', message: 'The demo server did not answer.' };
    }
    if (!res.ok) { const e = await res.json().catch(() => ({})); throw { code: e.code || 'upstream_error', message: e.message || 'HTTP ' + res.status }; }
    let text = '';
    if (stream && res.body) {
      const reader = res.body.getReader(), dec = new TextDecoder();
      try {
        for (;;) {
          const { done, value } = await reader.read();
          if (done) break;
          const delta = dec.decode(value, { stream: true });
          if (!delta) continue;
          text += delta; opts.onText({ text, delta });
        }
      } catch (e) { throw { code: e && e.name === 'AbortError' ? 'cancelled' : 'upstream_error', message: 'Interrupted', text }; }
    } else text = await res.text();
    if (!text.trim()) throw { code: 'empty_completion', message: 'Empty answer' };
    return { text, truncated: false, modelTierApplied: 'default' };
  }
  function parseJson(t) {
    try { return JSON.parse(t); } catch (e) { /* try fences and brackets below */ }
    const f = /```(?:json)?\s*([\s\S]*?)```/.exec(t);
    if (f) { try { return JSON.parse(f[1]); } catch (e) { /* fall through */ } }
    const s = t.search(/[[{]/), e = Math.max(t.lastIndexOf(']'), t.lastIndexOf('}'));
    if (s >= 0 && e > s) { try { return JSON.parse(t.slice(s, e + 1)); } catch (err) { /* fall through */ } }
    return undefined;
  }
  const sample = (input, opts) => call(input, opts, 'text');
  sample.json = async (input, opts) => {
    const r = await call(input, Object.assign({}, opts, { onText: undefined }), 'json');
    const v = parseJson(r.text);
    if (v === undefined) throw { code: 'invalid_json', message: 'The answer was not JSON', text: r.text };
    return v;
  };
  sample.limits = async () => ({ maxPromptBytes: 262144, images: { maxCount: 5, maxInputBytes: 10 * 1024 * 1024, mediaTypes: ['image/jpeg', 'image/png', 'image/webp', 'image/gif'] } });

  const caps = { db, downloads, user, sample };
  window.claude = { use: (name) => Promise.resolve(caps[name] || null) };
})();
