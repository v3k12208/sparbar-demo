(function () {
  'use strict';
  const E = window.Engine, Mo = window.Model, X = window.Exports;
  const SAMPLE = JSON.parse(document.getElementById('sample-data').textContent);
  const LIBS = Object.assign({
    xlsx: 'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js',
    jspdf: 'https://cdn.jsdelivr.net/npm/jspdf@2.5.1/dist/jspdf.umd.min.js',
    autotable: 'https://cdn.jsdelivr.net/npm/jspdf-autotable@3.8.2/dist/jspdf.plugin.autotable.min.js',
    pdfjs: 'https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.min.js',
    pdfworker: 'https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.worker.min.js'
  }, window.SPARBAR_LIBS || {});

  // ---------- utilities ----------
  const $ = (s, el = document) => el.querySelector(s);
  const esc = (s) => String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const clone = (o) => JSON.parse(JSON.stringify(o));
  const store = {
    get(k) { try { return localStorage.getItem('sparbar.' + k); } catch (e) { return null; } },
    set(k, v) { try { localStorage.setItem('sparbar.' + k, v); } catch (e) { /* storage unavailable */ } }
  };
  const parseNum = (v) => { if (v == null) return null; const n = Number(String(v).replace(/\s/g, '').replace(',', '.')); return Number.isFinite(n) ? n : null; };
  const today = () => { const d = new Date(); return new Date(d.getTime() - d.getTimezoneOffset() * 6e4).toISOString().slice(0, 10); };
  const nowIso = () => { const d = new Date(); return new Date(d.getTime() - d.getTimezoneOffset() * 6e4).toISOString().slice(0, 19); };
  const scripts = {};
  function loadScript(url) {
    if (!scripts[url]) scripts[url] = new Promise((res, rej) => { const s = document.createElement('script'); s.src = url; s.onload = res; s.onerror = () => { delete scripts[url]; rej(new Error('load ' + url)); }; document.head.appendChild(s); });
    return scripts[url];
  }
  async function sha256(text) {
    try { const b = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(text)); return [...new Uint8Array(b)].map((x) => x.toString(16).padStart(2, '0')).join(''); } catch (e) { return null; }
  }

  // ---------- state ----------
  const S = {
    lang: store.get('lang') === 'en' ? 'en' : 'sv', view: 'oversikt',
    dataset: null, ctx: null, ectx: null, rtOverrides: {},
    settings: { issuer: '', customerNames: {}, reportLang: 'sv', showCost: true },
    receipts: { local: [], shared: [], pending: [] }, useReceiptPrice: false, useReceiptMix: false, vatIncluded: false,
    result: null, checks: [], parity: null, scenarioKg: null, rstats: null,
    archive: [], archiveState: 'idle', caps: {}, me: null, names: {}, hashes: {},
    ui: { chartTable: {}, dPage: 0, dStatus: '', dType: '', dSearch: '', rtOnlyUsed: true, rtSearch: '', confirmLock: false, verify: {}, busy: '', aiState: '', aiMsg: '', importMsg: null, sharedMsg: '' },
    dirtyShared: false, sharedLoaded: false
  };
  const L = (sv, en) => (S.lang === 'en' ? en : sv);
  const loc = (lang) => (lang === 'en' ? 'en-GB' : 'sv-SE');
  const nfc = {};
  function nf(n, d = 0, lang = S.lang) {
    if (n == null || !Number.isFinite(n)) return '–';
    const k = lang + d; if (!nfc[k]) nfc[k] = new Intl.NumberFormat(loc(lang), { minimumFractionDigits: d, maximumFractionDigits: d });
    return nfc[k].format(n);
  }
  const pct = (x, d = 0) => nf(x * 100, d) + (S.lang === 'en' ? '%' : ' %');
  const kgParts = (kg) => (Math.abs(kg) >= 1000 ? [nf(kg / 1000, 1), L('ton CO₂e', 'tonnes CO₂e')] : [nf(kg, 0), 'kg CO₂e']);
  const kgText = (kg) => kgParts(kg).join(' ');
  const monthShort = (mk) => new Date(Date.UTC(+mk.slice(0, 4), +mk.slice(5, 7) - 1, 1)).toLocaleDateString(loc(S.lang), { month: 'short', timeZone: 'UTC' }).replace('.', '');

  // ---------- icons ----------
  const IP = {
    grid: '<rect x="3" y="3" width="7" height="9" rx="1.5"/><rect x="14" y="3" width="7" height="5" rx="1.5"/><rect x="14" y="12" width="7" height="9" rx="1.5"/><rect x="3" y="16" width="7" height="5" rx="1.5"/>',
    file: '<path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z"/><path d="M14 3v5h5M9 13h6M9 17h6"/>',
    shield: '<path d="M12 3l7 3v5c0 4.5-3 8.5-7 10-4-1.5-7-5.5-7-10V6z"/><path d="M9 12l2 2 4-4"/>',
    db: '<ellipse cx="12" cy="5.5" rx="7" ry="2.5"/><path d="M5 5.5v13c0 1.4 3.1 2.5 7 2.5s7-1.1 7-2.5v-13M5 12c0 1.4 3.1 2.5 7 2.5s7-1.1 7-2.5"/>',
    receipt: '<path d="M6 3h12v18l-3-2-3 2-3-2-3 2z"/><path d="M9 8h6M9 12h6M9 16h3"/>',
    sliders: '<path d="M4 6h10M18 6h2M4 12h4M12 12h8M4 18h12"/><circle cx="16" cy="6" r="2"/><circle cx="10" cy="12" r="2"/><circle cx="18" cy="18" r="2"/>',
    scale: '<path d="M12 4v16M7 20h10M5 8h14M5 8l-3 6a3 3 0 0 0 6 0zM19 8l-3 6a3 3 0 0 0 6 0z"/>',
    flow: '<rect x="3" y="4" width="6" height="6" rx="1.5"/><rect x="15" y="14" width="6" height="6" rx="1.5"/><path d="M9 7h4a2 2 0 0 1 2 2v5"/>',
    archive: '<rect x="3" y="4" width="18" height="5" rx="1"/><path d="M5 9v10a1 1 0 0 0 1 1h12a1 1 0 0 0 1-1V9M10 13h4"/>',
    lock: '<rect x="5" y="11" width="14" height="10" rx="2"/><path d="M8 11V8a4 4 0 0 1 8 0v3"/>',
    download: '<path d="M12 4v11M7 10l5 5 5-5M5 20h14"/>',
    upload: '<path d="M12 20V9M7 14l5-5 5 5M5 4h14"/>',
    trace: '<circle cx="11" cy="11" r="6"/><path d="M20 20l-4.5-4.5"/>',
    x: '<path d="M6 6l12 12M18 6L6 18"/>',
    check: '<path d="M5 12.5l4.5 4.5L19 7.5"/>',
    critical: '<path d="M8 3h8l5 5v8l-5 5H8l-5-5V8z"/><path d="M9.5 9.5l5 5M14.5 9.5l-5 5"/>',
    serious: '<path d="M12 4l9 16H3z"/><path d="M12 10v4M12 17.5v.01"/>',
    warning: '<circle cx="12" cy="12" r="9"/><path d="M12 7.5v5.5M12 16.5v.01"/>',
    info: '<circle cx="12" cy="12" r="9"/><path d="M12 11v6M12 7.5v.01"/>',
    ok: '<path d="M5 12.5l4.5 4.5L19 7.5"/>',
    spark: '<path d="M12 3l1.8 5.2L19 10l-5.2 1.8L12 17l-1.8-5.2L5 10l5.2-1.8z"/>',
    yes: '<circle cx="12" cy="12" r="9"/><path d="M8 12.5l2.7 2.7L16 10"/>',
    part: '<circle cx="12" cy="12" r="9"/><path d="M12 3a9 9 0 0 1 0 18z" fill="currentColor"/>',
    plan: '<circle cx="12" cy="12" r="9" stroke-dasharray="3 3"/>'
  };
  const icon = (n) => `<svg viewBox="0 0 24 24" aria-hidden="true">${IP[n] || ''}</svg>`;
  const sevLabel = (s) => ({ critical: L('Kritisk', 'Critical'), serious: L('Allvarlig', 'Serious'), warning: L('Varning', 'Warning'), info: L('Info', 'Info'), ok: L('OK', 'OK') }[s]);
  const sevChip = (s) => `<span class="sev ${s}">${icon(s)}${sevLabel(s)}</span>`;

  // ---------- dataset & context ----------
  function loadSample() {
    const lines = E.fromCompact(SAMPLE);
    const mt = SAMPLE.meta;
    if (mt.customerNames) S.settings.customerNames = Object.assign({}, mt.customerNames, S.settings.customerNames);
    if (mt.issuer && !S.settings.issuer) S.settings.issuer = mt.issuer;
    const excel = mt.excel ? Object.assign({ customer: mt.customer, from: mt.period ? mt.period.split('/')[0] : '2026-01-01', to: mt.period ? mt.period.split('/')[1] : '2026-06-30' }, mt.excel) : null;
    setDataset({ name: mt.datasetName || 'Milj.xlsx', isSample: true, source: mt.source, lines, rtBase: E.resourceTypesFrom(SAMPLE.resourceTypes), tables: mt.tables || ['Antal ton (orderrader)', 'Resursdata (drivmedelsandelar)'], excel });
  }
  function setDataset(ds) {
    S.dataset = ds;
    const dates = ds.lines.map((l) => l.date).filter(Boolean).sort();
    ds.minDate = dates[0]; ds.maxDate = dates[dates.length - 1];
    ds.customers = [...new Set(ds.lines.map((l) => l.customer).filter((c) => c != null))].sort();
    ds.presets = ds.minDate ? E.periodPresets(ds.minDate, ds.maxDate) : [];
    const prev = S.ctx;
    const p = ds.presets[0] || { id: 'custom', from: ds.minDate, to: ds.maxDate };
    S.ctx = {
      customer: ds.customers[0] || null, periodId: p.id, from: p.from, to: p.to, statuses: ['F'],
      resourceTypes: null, fuels: prev ? prev.fuels : clone(E.DEFAULT_FUELS),
      prices: prev ? prev.prices : { hvo: 21.6, diesel: 15.92 },
      fuelPlan: prev ? prev.fuelPlan : { hvoFuel: 'HVO100', dieselFuel: 'DIESEL_RP', hvoPct: 1, byType: {} },
      tonnageBasis: prev ? prev.tonnageBasis : 'excel'
    };
    rebuildRt();
    S.ui.dPage = 0;
  }
  function rebuildRt() {
    const m = new Map();
    for (const [k, v] of S.dataset.rtBase) {
      const o = S.rtOverrides[k];
      m.set(k, o ? Object.assign({}, v, o) : v);
    }
    for (const [k, o] of Object.entries(S.rtOverrides)) if (!m.has(k) && o.name) m.set(k, Object.assign({ code: k, diesel: null, hvo: null }, o));
    S.ctx.resourceTypes = m;
  }
  const currentPreset = () => S.dataset.presets.find((p) => p.id === S.ctx.periodId) || null;
  function periodInfo() {
    const p = currentPreset();
    if (p) return { id: p.id, label: p.label[S.settings.reportLang] || p.label.sv, uiLabel: p.label[S.lang], short: p.short };
    const lbl = `${S.ctx.from} – ${S.ctx.to}`;
    return { id: `${S.ctx.from}_${S.ctx.to}`, label: lbl, uiLabel: lbl, short: lbl };
  }
  const custName = (no, lang = S.lang) => (S.settings.customerNames[no] || (lang === 'en' ? `Customer ${no}` : `Kund ${no}`));
  const allReceipts = () => S.receipts.local.concat(S.receipts.shared);

  function effectiveCtx() {
    const c = Object.assign({}, S.ctx, { prices: Object.assign({}, S.ctx.prices), fuelPlan: Object.assign({}, S.ctx.fuelPlan) });
    S.rstats = E.receiptStats(allReceipts(), c.from, c.to, S.vatIncluded);
    S.receiptApplied = { price: false, mix: false };
    if (S.useReceiptPrice) {
      if (S.rstats.per.hvo && S.rstats.per.hvo.avgPrice) { c.prices.hvo = S.rstats.per.hvo.avgPrice; S.receiptApplied.price = true; }
      if (S.rstats.per.diesel && S.rstats.per.diesel.avgPrice) { c.prices.diesel = S.rstats.per.diesel.avgPrice; S.receiptApplied.price = true; }
    }
    if (S.useReceiptMix && S.rstats.hvoShareOfVolume != null) { c.fuelPlan.hvoPct = S.rstats.hvoShareOfVolume; c.fuelPlan.byType = {}; S.receiptApplied.mix = true; }
    return c;
  }
  function recompute() {
    const c = effectiveCtx();
    S.ectx = c;
    S.result = E.compute(S.dataset.lines, c);
    S.scenarioKg = E.compute(S.dataset.lines, Object.assign({}, c, { fuelPlan: Object.assign({}, c.fuelPlan, { hvoPct: 0, byType: {} }) })).totals.kgWtw;
    const ex = S.dataset.excel;
    const applicable = !!(ex && c.customer === ex.customer && c.from === ex.from && c.to === ex.to && c.fuelPlan.hvoFuel === 'HVO100' &&
      c.fuels.HVO100.wtw === ex.efHvo && c.prices.hvo === ex.price && c.fuelPlan.hvoPct === 1 && !Object.keys(c.fuelPlan.byType).length && !Object.keys(S.rtOverrides).length);
    S.parity = applicable ? { applicable, portal: S.result.totals.kgWtw, excel: ex.totalKg, diff: Math.abs(S.result.totals.kgWtw - ex.totalKg) } : { applicable: false };
    if (applicable) S.parity.ok = S.parity.diff < 0.01;
    S.checks = E.checks(S.dataset.lines, S.result, c, { parity: S.parity, receiptsUsed: S.receiptApplied.mix || (S.rstats.n > 0 && S.useReceiptPrice), priceFromReceipts: S.receiptApplied.price });
  }

  // ---------- report model & hashes ----------
  function reportKey() { return `${S.ectx.customer}|${S.ectx.from}|${S.ectx.to}`; }
  function reportId() { const p = periodInfo(); return `KR-${S.ectx.customer}-${p.id}`.replace(/[^A-Za-z0-9_\-.]/g, '_'); }
  function modelFor(lang, withLines) {
    const p = periodInfo();
    const pl = currentPreset();
    const period = { id: p.id, short: p.short, label: pl ? (pl.label[lang] || pl.label.sv) : p.label };
    const inv = withLines ? E.inScope(S.dataset.lines, S.ectx).filter((l) => S.ectx.statuses.includes(l.status)) : [];
    const rt = S.ectx.resourceTypes;
    const st = reportStatus();
    const model = Mo.buildModel({
      lang, ctx: S.ectx, result: S.result, settings: S.settings, scenarioKg: S.scenarioKg, id: reportId(), period,
      status: st.status, version: st.version, issued: today(), generatedAt: nowIso(),
      receipts: { price: S.receiptApplied.price, mix: S.receiptApplied.mix, n: S.rstats.n },
      hashes: { input: S.hashes.input && S.hashes.input.key === inputSig() ? S.hashes.input.v : null, report: S.hashes.report && S.hashes.report.key === reportSig() ? S.hashes.report.v : null },
      checks: S.checks.map((c) => ({ sev: c.sev, t: (lang === 'en' ? c.en : c.sv).t, d: (lang === 'en' ? c.en : c.sv).d })),
      lines: inv.map((l) => [l.date, l.status, l.resource, l.rtype, rt.get(l.rtype) ? rt.get(l.rtype).name : '', l.orderId, l.article, l.km, l.qty, l.unit, l.amount])
    });
    const ai = S.ai && S.ai.summary[`${S.ectx.customer}|${S.ectx.from}|${S.ectx.to}|${lang}`];
    if (ai && ai.text && !ai.live && ai.sig === reportSig()) { model.narrative = ai.text; model.aiNarrative = true; }
    return model;
  }
  function inputSig() { return `${S.dataset.name}|${S.dataset.lines.length}|${S.ectx.customer}|${S.ectx.from}|${S.ectx.to}`; }
  function reportCore() {
    const c = S.ectx, t = S.result.totals;
    return {
      id: reportId(), customer: c.customer, from: c.from, to: c.to, issuer: S.settings.issuer || '', customerName: S.settings.customerNames[c.customer] || '',
      totals: { kgWtw: t.kgWtw, kgTtw: t.kgTtw, liters: t.liters, tonnes: t.tonnes, tkm: t.tkm, mwh: t.mwh },
      rows: S.result.rows.map((r) => [r.code, r.fuel, r.amount, r.share, r.price, r.liters, r.ef, r.kgWtw]),
      factors: [...new Set(S.result.rows.map((r) => r.fuel))].map((k) => [k, c.fuels[k].wtw, c.fuels[k].ttw, c.fuels[k].validFrom]),
      prices: c.prices, hvoPct: c.fuelPlan.hvoPct, tonnageBasis: c.tonnageBasis
    };
  }
  function reportSig() { return E.canonical(reportCore()); }
  let hashTimer = null;
  function ensureHashes() {
    const iKey = inputSig(), rKey = reportSig();
    const need = (!S.hashes.input || S.hashes.input.key !== iKey) || (!S.hashes.report || S.hashes.report.key !== rKey);
    if (!need) return Promise.resolve();
    clearTimeout(hashTimer);
    return new Promise((resolve) => {
      hashTimer = setTimeout(async () => {
        const scope = E.inScope(S.dataset.lines, S.ectx);
        const [hi, hr] = await Promise.all([sha256(E.canonical(E.linesFingerprint(scope))), sha256(rKey)]);
        S.hashes.input = { key: iKey, v: hi }; S.hashes.report = { key: rKey, v: hr };
        resolve(); softRefresh();
      }, 30);
    });
  }
  function reportStatus() {
    const key = reportKey();
    const same = S.archive.filter((a) => a.key === key).sort((a, b) => b.version - a.version);
    if (!same.length) return { status: 'draft', version: null };
    const h = S.hashes.report && S.hashes.report.key === reportSig() ? S.hashes.report.v : null;
    if (h && same[0].reportHash === h) return { status: 'locked', version: same[0].version, rec: same[0] };
    return { status: 'changed', version: same[0].version, rec: same[0] };
  }

  // ---------- capabilities ----------
  function initCaps() {
    const c = window.claude;
    if (!c || typeof c.use !== 'function') { S.archiveState = 'none'; return; }
    c.use('downloads').then((d) => { S.caps.downloads = d; softRefresh(); }).catch(() => {});
    c.use('user').then(async (u) => { S.caps.user = u; if (u) { try { S.me = await u.id(); } catch (e) { S.me = null; } } }).catch(() => {});
    c.use('sample').then((s) => { S.caps.sample = s; if (s && s.limits) s.limits().then((l) => { S.caps.sampleLimits = l; softRefresh(); }).catch(() => {}); softRefresh(); }).catch(() => {});
    S.archiveState = 'loading';
    c.use('db').then((db) => { S.caps.db = db; if (db) subscribeDb(); else S.archiveState = 'none'; softRefresh(); }).catch(() => { S.archiveState = 'none'; });
  }
  function subscribeDb() {
    const db = S.caps.db;
    const timer = setTimeout(() => { if (S.archiveState === 'loading') { S.archiveState = 'slow'; softRefresh(); } }, 10000);
    db.doc('app/settings').onSnapshot((snap) => {
      if (snap.exists && !snap.metadata.hasPendingWrites) {
        const d = snap.data();
        const first = !S.sharedLoaded;
        if (first || !S.dirtyShared) { applyShared(d); if (!first && d.updatedBy !== S.me) toast(L('Gemensamma inställningar uppdaterades.', 'Shared settings were updated.')); }
      }
      S.sharedLoaded = true;
    }, () => {});
    db.collection('reports').orderBy('createdAt', 'desc').limit(300).onSnapshot((qs) => {
      clearTimeout(timer);
      S.archive = qs.docs.map((d) => d.data());
      S.archiveState = 'ready';
      resolveNames(S.archive.map((a) => a.createdBy).filter(Boolean));
      softRefresh();
    }, () => { clearTimeout(timer); S.archiveState = 'error'; softRefresh(); });
    db.collection('receipts').onSnapshot((qs) => {
      S.receipts.shared = [];
      S.receiptBatches = qs.docs.map((d) => { const b = d.data(); (b.rows || []).forEach((r) => S.receipts.shared.push(Object.assign({ batch: d.id }, r))); return Object.assign({ id: d.id }, b); });
      recompute(); softRefresh();
    }, () => {});
  }
  async function resolveNames(ids) {
    const u = S.caps.user; if (!u || !u.profiles || !ids.length) return;
    try { const ps = await u.profiles([...new Set(ids)]); for (const id of Object.keys(ps || {})) S.names[id] = (ps[id] && ps[id].name) || ''; softRefresh(); } catch (e) { /* names stay unresolved */ }
  }
  function sharedPayload() {
    const c = S.ctx;
    return {
      issuer: S.settings.issuer, customerNames: S.settings.customerNames, reportLang: S.settings.reportLang, showCost: S.settings.showCost,
      prices: c.prices, fuelPlan: c.fuelPlan, tonnageBasis: c.tonnageBasis, fuels: c.fuels, rtOverrides: S.rtOverrides,
      updatedAt: new Date().toISOString(), updatedBy: S.me || null
    };
  }
  function applyShared(d) {
    if (!d) return;
    if (typeof d.issuer === 'string') S.settings.issuer = d.issuer;
    if (d.customerNames && typeof d.customerNames === 'object') S.settings.customerNames = Object.assign({}, d.customerNames);
    if (d.reportLang === 'sv' || d.reportLang === 'en') S.settings.reportLang = d.reportLang;
    if (typeof d.showCost === 'boolean') S.settings.showCost = d.showCost;
    if (d.prices && Number.isFinite(d.prices.hvo) && Number.isFinite(d.prices.diesel)) S.ctx.prices = { hvo: d.prices.hvo, diesel: d.prices.diesel };
    if (d.fuels && typeof d.fuels === 'object') {
      const f = clone(E.DEFAULT_FUELS);
      for (const [k, v] of Object.entries(d.fuels)) if (v && Number.isFinite(v.wtw) && Number.isFinite(v.ttw) && Number.isFinite(v.mj) && (v.shareCol === 'hvo' || v.shareCol === 'diesel')) f[k] = v;
      S.ctx.fuels = f;
    }
    if (d.fuelPlan && S.ctx.fuels[d.fuelPlan.hvoFuel] && S.ctx.fuels[d.fuelPlan.dieselFuel]) S.ctx.fuelPlan = { hvoFuel: d.fuelPlan.hvoFuel, dieselFuel: d.fuelPlan.dieselFuel, hvoPct: Math.min(1, Math.max(0, Number(d.fuelPlan.hvoPct) || 0)), byType: d.fuelPlan.byType || {} };
    if (d.tonnageBasis === 'excel' || d.tonnageBasis === 'invoiced') S.ctx.tonnageBasis = d.tonnageBasis;
    if (d.rtOverrides && typeof d.rtOverrides === 'object') { S.rtOverrides = d.rtOverrides; rebuildRt(); }
    S.dirtyShared = false;
    recompute(); render();
  }
  async function saveShared() {
    const db = S.caps.db;
    if (!db) { toast(L('Gemensam lagring kräver att sidan är öppen i Claude.', 'Shared storage needs the page open in Claude.')); return; }
    try { await db.doc('app/settings').set(sharedPayload()); S.dirtyShared = false; toast(L('Sparat som gemensam standard.', 'Saved as the shared default.')); render(); }
    catch (e) { toast(e && e.code === 'invalid_argument' ? L('Du har inte behörighet att ändra gemensamma inställningar.', 'You do not have permission to change shared settings.') : L('Kunde inte spara. Försök igen.', 'Could not save. Try again.')); }
  }
  const markDirty = () => { S.dirtyShared = true; };

  // ---------- downloads ----------
  async function saveFile(filename, data) {
    const d = S.caps.downloads;
    if (!d) { toast(L('Nedladdning fungerar när sidan är öppen i Claude.', 'Downloads work when the page is open in Claude.')); return false; }
    try { await d.save({ filename, data }); toast(L('Filen är sparad.', 'File saved.')); return true; }
    catch (e) {
      if (e && e.code === 'declined') return false;
      if (e && e.code === 'rate_limited') toast(L('En nedladdning väntar redan på svar.', 'A download is already waiting for an answer.'));
      else toast(L('Nedladdningen gick inte att starta här.', 'The download could not start here.'));
      return false;
    }
  }
  async function doExport(kind) {
    if (!S.result.rows.length) { toast(L('Inget att exportera för vald period.', 'Nothing to export for this period.')); return; }
    await ensureHashes();
    const lang = S.settings.reportLang;
    const m = modelFor(lang, kind === 'xlsx');
    const base = `${lang === 'en' ? 'Climate_report' : 'Klimatrapport'}_${m.customer.no}_${m.period.short}`.replace(/[^A-Za-z0-9_\-.]/g, '_');
    S.ui.busy = kind; softRefresh();
    try {
      if (kind === 'xlsx') { await loadScript(LIBS.xlsx); await saveFile(base + '.xlsx', new Uint8Array(X.buildXlsx(m, window.XLSX))); }
      else if (kind === 'pdf') { await loadScript(LIBS.jspdf); await loadScript(LIBS.autotable); await saveFile(base + '.pdf', new Uint8Array(X.buildPdf(m, window.jspdf.jsPDF))); }
      else if (kind === 'csv') await saveFile(base + '_scope3.csv', X.buildCsv(m));
      else if (kind === 'json') await saveFile(base + '_iso14083.json', X.buildJson(m));
    } catch (e) { toast(L('Exporten misslyckades: biblioteket kunde inte laddas.', 'Export failed: the library could not load.')); }
    S.ui.busy = ''; softRefresh();
  }

  // ---------- lock & verify ----------
  async function lockReport() {
    const db = S.caps.db; if (!db) return;
    await ensureHashes();
    const key = reportKey();
    const version = Math.max(0, ...S.archive.filter((a) => a.key === key).map((a) => a.version)) + 1;
    const c = S.ectx, t = S.result.totals, p = periodInfo();
    const id = `${reportId()}-v${version}`;
    const body = {
      id, key, version, reportId: reportId(), customer: c.customer, customerName: custName(c.customer, S.settings.reportLang), period: p.uiLabel, from: c.from, to: c.to,
      issuer: S.settings.issuer || '', lang: S.settings.reportLang,
      totals: { kgWtw: t.kgWtw, kgTtw: t.kgTtw, kgWtt: t.kgWtt, liters: t.liters, tonnes: t.tonnes, tkm: t.tkm, mwh: t.mwh, renewableShare: t.renewableShare, amount: t.amount },
      rows: S.result.rows.map((r) => ({ code: r.code, name: r.name, fuel: r.fuel, liters: r.liters, ef: r.ef, kgWtw: r.kgWtw, amount: r.amount })),
      factors: [...new Set(S.result.rows.map((r) => r.fuel))].map((k) => ({ key: k, wtw: c.fuels[k].wtw, ttw: c.fuels[k].ttw, validFrom: c.fuels[k].validFrom || '', source: c.fuels[k].source ? c.fuels[k].source.sv : '' })),
      prices: c.prices, hvoPct: c.fuelPlan.hvoPct, tonnageBasis: c.tonnageBasis, receipts: S.receiptApplied,
      inputHash: S.hashes.input.v, reportHash: S.hashes.report.v, dataset: { name: S.dataset.name, rows: S.dataset.lines.length, sample: !!S.dataset.isSample },
      openChecks: S.checks.filter((x) => x.sev !== 'ok' && x.sev !== 'info').map((x) => x.sv.t),
      createdAt: new Date().toISOString(), createdBy: S.me || null
    };
    S.ui.busy = 'lock'; softRefresh();
    try { await db.collection('reports').doc(id.replace(/[^A-Za-z0-9_\-.~:@+]/g, '_')).set(body); S.ui.confirmLock = false; toast(L(`Rapporten är låst som version ${version}.`, `The report is locked as version ${version}.`)); }
    catch (e) { toast(e && e.code === 'invalid_argument' ? L('Du har inte behörighet att låsa rapporter.', 'You do not have permission to lock reports.') : L('Låsningen misslyckades. Försök igen.', 'Locking failed. Try again.')); }
    S.ui.busy = ''; render();
  }
  async function verifyArchived(id) {
    const a = S.archive.find((x) => x.id === id); if (!a) return;
    const lines = S.dataset.lines.filter((l) => l.customer === a.customer && l.date >= a.from && l.date <= a.to);
    if (!lines.length) { S.ui.verify[id] = { sev: 'warning', t: L('Underlaget för perioden finns inte i laddad data.', 'The period\'s data is not in the loaded dataset.') }; render(); return; }
    const h = await sha256(E.canonical(E.linesFingerprint(lines)));
    S.ui.verify[id] = h === a.inputHash ? { sev: 'ok', t: L('Underlaget är oförändrat sedan låsning.', 'The input is unchanged since locking.') } : { sev: 'serious', t: L('Underlaget har ändrats sedan låsning.', 'The input has changed since locking.') };
    render();
  }

  // ---------- imports ----------
  async function readTables(file) {
    await loadScript(LIBS.xlsx);
    const buf = await file.arrayBuffer();
    const isCsv = /\.(csv|txt)$/i.test(file.name);
    const wb = isCsv ? window.XLSX.read(new TextDecoder('utf-8').decode(buf), { type: 'string', raw: false }) : window.XLSX.read(buf, { type: 'array' });
    return wb.SheetNames.map((n) => ({ sheet: n, rows: window.XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: true, defval: null }) }));
  }
  async function importData(files) {
    const found = { orders: [], invoiceLines: [], resourceTypes: null, names: [], notes: [], excel: null };
    for (const f of files) {
      if (!/\.(xlsx|xlsm|xls|csv|txt)$/i.test(f.name)) { found.notes.push(L(`${f.name}: filtypen stöds inte.`, `${f.name}: file type not supported.`)); continue; }
      let tables;
      try { tables = await readTables(f); } catch (e) { found.notes.push(L(`${f.name}: kunde inte läsas.`, `${f.name}: could not be read.`)); continue; }
      for (const tb of tables) {
        if (/^rapport$/i.test(tb.sheet)) {
          const rr = tb.rows, total = rr[0] && Number(rr[0][8]), ef = rr[3] && Number(rr[3][4]), price = rr[3] && Number(rr[3][9]);
          const hm = /^([12])-(\d{4})$/.exec(String((rr[3] && rr[3][0]) || '').trim());
          if (Number.isFinite(total) && total > 0 && ef > 0 && price > 0 && hm) found.excel = { totalKg: total, efHvo: ef, price, from: hm[1] === '1' ? `${hm[2]}-01-01` : `${hm[2]}-07-01`, to: hm[1] === '1' ? `${hm[2]}-06-30` : `${hm[2]}-12-31` };
        }
        const c = E.classifyTable(tb.rows);
        if (c.kind === 'orders') { found.orders.push(...c.lines); found.names.push(`${tb.sheet} (${L('orderrader', 'order lines')}, ${c.lines.length})`); }
        else if (c.kind === 'invoiceLines') { found.invoiceLines.push(Object.assign(c, { sheet: tb.sheet })); found.names.push(`${tb.sheet} (${L('fakturarader', 'invoice lines')}, ${c.lines.length}${c.blankAmounts ? ', ' + L('belopp saknas', 'amounts missing') : ''})`); }
        else if (c.kind === 'resourceTypes') { found.resourceTypes = c.rows; found.names.push(`${tb.sheet} (${L('resurstyper', 'vehicle types')}, ${c.rows.length})`); }
      }
    }
    let lines = found.orders;
    const inv = found.invoiceLines.filter((x) => !x.blankAmounts).flatMap((x) => x.lines);
    const invAll = found.invoiceLines.flatMap((x) => x.lines);
    if (!lines.length && inv.length) lines = inv;
    if (!lines.length) { S.ui.importMsg = { sev: 'serious', t: L('Hittade inga orderrader eller fakturarader med belopp.', 'Found no order lines or invoice lines with amounts.'), d: found.names.concat(found.notes) }; render(); return; }
    // fill customer & supplier from invoice lines when orders lack them
    const invCustomers = [...new Set(invAll.map((l) => l.customer).filter(Boolean))];
    const supByRes = new Map(invAll.filter((l) => l.resource != null && l.supplier != null).map((l) => [l.resource, l.supplier]));
    let filled = 0;
    for (const l of lines) {
      if (!l.customer) { l.customer = invCustomers.length === 1 ? invCustomers[0] : (S.ui.importCustomer || L('okänd', 'unknown')); filled++; }
      if (l.supplier == null && supByRes.has(l.resource)) l.supplier = supByRes.get(l.resource);
    }
    if (filled) found.notes.push(L(`Kundnummer saknades på ${filled} rader och sattes till ${lines[0].customer}.`, `Customer number was missing on ${filled} rows and was set to ${lines[0].customer}.`));
    const rtBase = found.resourceTypes ? E.resourceTypesFrom(found.resourceTypes) : S.dataset.rtBase;
    if (!found.resourceTypes) found.notes.push(L('Ingen resurstypstabell i filen; nuvarande drivmedelsandelar används.', 'No vehicle-type table in the file; current fuel shares are used.'));
    if (found.excel) {
      found.excel.customer = invCustomers.length === 1 ? invCustomers[0] : lines[0].customer;
      found.notes.push(L(`Fliken Rapport hittades: Excel visar ${nf(found.excel.totalKg, 2)} kg CO₂e. Portalen jämför automatiskt.`, `Found the Rapport sheet: Excel shows ${nf(found.excel.totalKg, 2)} kg CO₂e. The portal compares automatically.`));
    }
    setDataset({ name: files.map((f) => f.name).join(', '), isSample: false, source: L('Importerad fil', 'Imported file'), lines, rtBase, tables: found.names, excel: found.excel || null });
    S.ui.importMsg = { sev: 'ok', t: L(`Importerade ${nf(lines.length)} rader.`, `Imported ${nf(lines.length)} rows.`), d: found.names.concat(found.notes) };
    recompute(); render();
  }

  const EXAMPLE_RECEIPTS = [['2026-01-14', 'HVO100', 412.6, 8913.0], ['2026-02-03', 'HVO100', 388.2, 8268.7], ['2026-02-21', 'Diesel MK1', 301.5, 4848.1], ['2026-03-11', 'HVO100', 455.0, 9873.5],
    ['2026-04-02', 'HVO100', 398.4, 8565.6], ['2026-04-29', 'HVO100', 430.9, 9393.6], ['2026-05-19', 'Diesel MK1', 275.0, 4400.0], ['2026-06-09', 'HVO100', 441.3, 9620.3]];
  function exampleReceipts() {
    return EXAMPLE_RECEIPTS.map(([date, fuel, liters, amount]) => ({ date, fuel, fuelClass: E.normalizeFuel(fuel), liters, amount, vatIncluded: false, supplier: L('Exempelstation', 'Example station'), vehicle: 'EXEMPEL', source: 'example' }));
  }
  const R_HEAD = {
    date: ['datum', 'date', 'transaktionsdatum', 'köpdatum', 'kopdatum', 'tankdatum'], fuel: ['drivmedel', 'produkt', 'artikel', 'product', 'fuel', 'vara'],
    liters: ['liter', 'volym', 'antal', 'kvantitet', 'volume', 'litres', 'liters', 'mängd'], amount: ['belopp_exkl_moms', 'belopp exkl moms', 'belopp', 'summa', 'totalt', 'amount', 'netto', 'belopp_inkl_moms', 'belopp inkl moms'],
    unitPrice: ['pris', 'pris/l', 'literpris', 'á-pris', 'a-pris', 'unit price', 'pris per liter'], vehicle: ['regnr', 'registreringsnummer', 'fordon', 'kort', 'kortnummer', 'vehicle'], supplier: ['leverantör', 'leverantor', 'station', 'bolag', 'supplier']
  };
  async function importReceiptTable(file) {
    const tables = await readTables(file);
    const out = [];
    for (const tb of tables) {
      if (!tb.rows.length) continue;
      const H = tb.rows[0].map((h) => String(h == null ? '' : h).trim().toLowerCase());
      const ix = {}; for (const [k, alts] of Object.entries(R_HEAD)) { const i = H.findIndex((h) => alts.includes(h)); if (i >= 0) ix[k] = i; }
      if (ix.date == null || ix.liters == null) continue;
      const vatIncl = ix.amount != null && /inkl/.test(H[ix.amount]);
      for (const r of tb.rows.slice(1)) {
        const date = E.toISODate(r[ix.date]);
        const liters = parseNum(r[ix.liters]); if (!date || !liters) continue;
        const fuel = ix.fuel != null ? String(r[ix.fuel] || '') : 'Diesel';
        out.push({ date, fuel, fuelClass: E.normalizeFuel(fuel), liters, amount: ix.amount != null ? parseNum(r[ix.amount]) : null, unitPrice: ix.unitPrice != null ? parseNum(r[ix.unitPrice]) : null, vatIncluded: vatIncl, vehicle: ix.vehicle != null ? String(r[ix.vehicle] || '') : '', supplier: ix.supplier != null ? String(r[ix.supplier] || '') : '', source: 'csv', file: file.name });
      }
    }
    return out;
  }
  async function pdfToImages(file, maxPages) {
    await loadScript(LIBS.pdfworker); await loadScript(LIBS.pdfjs);
    const lib = window.pdfjsLib; lib.GlobalWorkerOptions.workerSrc = LIBS.pdfworker;
    const pdf = await lib.getDocument({ data: new Uint8Array(await file.arrayBuffer()) }).promise;
    const out = [];
    for (let i = 1; i <= Math.min(pdf.numPages, maxPages); i++) {
      const page = await pdf.getPage(i); const vp = page.getViewport({ scale: 1.8 });
      const cv = document.createElement('canvas'); cv.width = vp.width; cv.height = vp.height;
      await page.render({ canvasContext: cv.getContext('2d'), viewport: vp }).promise;
      out.push(await new Promise((r) => cv.toBlob(r, 'image/png')));
    }
    return out;
  }
  async function importReceipts(files) {
    const tableFiles = files.filter((f) => /\.(csv|xlsx|xls|txt)$/i.test(f.name));
    const imageFiles = files.filter((f) => /^image\//.test(f.type) || /\.pdf$/i.test(f.name));
    let added = 0;
    for (const f of tableFiles) {
      try { const rows = await importReceiptTable(f); S.receipts.pending.push(...rows); added += rows.length; } catch (e) { toast(L(`${f.name} kunde inte läsas.`, `${f.name} could not be read.`)); }
    }
    if (imageFiles.length) await aiReadReceipts(imageFiles);
    if (tableFiles.length && !added) toast(L('Hittade inga kvittorader (kolumnerna datum och liter krävs).', 'Found no receipt rows (date and litre columns are required).'));
    render();
  }
  let aiCtl = null;
  async function aiReadReceipts(files) {
    const sample = S.caps.sample, lim = S.caps.sampleLimits && S.caps.sampleLimits.images;
    if (!sample || !lim) { S.ui.aiState = 'off'; S.ui.aiMsg = L('AI-avläsning av bilder och PDF fungerar när sidan är öppen i Claude.', 'AI reading of images and PDFs works when the page is open in Claude.'); render(); return; }
    S.ui.aiState = 'busy'; S.ui.aiMsg = L('Förbereder filer…', 'Preparing files…'); render();
    const items = [];
    try {
      for (const f of files) {
        if (/\.pdf$/i.test(f.name) || f.type === 'application/pdf') { (await pdfToImages(f, 3)).forEach((b) => items.push({ blob: b, name: f.name })); }
        else if (lim.mediaTypes.includes(f.type) && f.size <= lim.maxInputBytes) items.push({ blob: f, name: f.name });
      }
    } catch (e) { S.ui.aiState = 'error'; S.ui.aiMsg = L('PDF-filen kunde inte öppnas.', 'The PDF could not be opened.'); render(); return; }
    if (!items.length) { S.ui.aiState = 'error'; S.ui.aiMsg = L('Inga bilder eller PDF-sidor att läsa.', 'No images or PDF pages to read.'); render(); return; }
    const chunks = []; for (let i = 0; i < items.length; i += lim.maxCount) chunks.push(items.slice(i, i + lim.maxCount));
    let found = 0;
    for (const [ci, ch] of chunks.entries()) {
      S.ui.aiMsg = L(`Claude läser kvitton (${ci + 1}/${chunks.length})…`, `Claude is reading receipts (${ci + 1}/${chunks.length})…`); render();
      aiCtl = new AbortController();
      const prompt = `You read fuel receipts and fuel invoices for a Swedish haulage company. The images are, in order, pages from these files: ${ch.map((c) => c.name).join(', ')}.
Reply with ONLY a JSON array. One object per fuel purchase line (skip AdBlue, washer fluid and anything that is not vehicle fuel):
{"date":"YYYY-MM-DD","product":"text as printed","fuel":"HVO100"|"DIESEL"|"B100"|"OTHER","liters":number,"unit_price":number|null,"amount":number|null,"amount_includes_vat":true|false|null,"vehicle":"registration or card number"|null,"supplier":"company or station"|null,"file":"file name","confidence":number between 0 and 1}
Use a dot as decimal separator. Swedish receipts use comma decimals and "inkl. moms"/"exkl. moms" for VAT. If a value cannot be read, use null. Never guess litres. If nothing is a fuel receipt, reply [].
Example: [{"date":"2026-03-04","product":"HVO100","fuel":"HVO100","liters":412.5,"unit_price":21.6,"amount":8910.0,"amount_includes_vat":false,"vehicle":"ABC123","supplier":"Example AB","file":"kvitto1.pdf","confidence":0.9}]`;
      try {
        const arr = await sample.json(prompt, { images: ch.map((c) => c.blob), signal: aiCtl.signal });
        for (const r of Array.isArray(arr) ? arr : []) {
          const liters = parseNum(r.liters); const date = /^\d{4}-\d{2}-\d{2}$/.test(String(r.date || '')) ? r.date : null;
          if (!liters || !date) continue;
          S.receipts.pending.push({ date, fuel: String(r.product || r.fuel || ''), fuelClass: E.normalizeFuel(r.fuel + ' ' + (r.product || '')), liters, amount: parseNum(r.amount), unitPrice: parseNum(r.unit_price), vatIncluded: r.amount_includes_vat === true, vehicle: r.vehicle ? String(r.vehicle) : '', supplier: r.supplier ? String(r.supplier) : '', source: 'ai', file: String(r.file || ''), confidence: Number(r.confidence) || null });
          found++;
        }
      } catch (e) {
        const code = e && e.code;
        S.ui.aiState = 'error';
        S.ui.aiMsg = code === 'cancelled' ? L('Avbrutet.', 'Stopped.') : code === 'not_granted' || code === 'sampling_disabled' ? L('Claude har inte tillåtits för den här sidan.', 'Claude has not been allowed for this page.') : code === 'rate_limited' ? L('För många förfrågningar just nu. Försök igen om en stund.', 'Too many requests right now. Try again in a while.') : code === 'image_rejected' ? L('En fil kunde inte läsas som bild.', 'A file could not be read as an image.') : code === 'invalid_json' ? L('Svaret gick inte att tolka. Försök igen.', 'The answer could not be parsed. Try again.') : L('Avläsningen avbröts. Försök igen.', 'Reading was interrupted. Try again.');
        render(); return;
      }
    }
    S.ui.aiState = 'done'; S.ui.aiMsg = L(`Claude hittade ${found} ${found === 1 ? 'kvittorad' : 'kvittorader'}. Granska innan du sparar.`, `Claude found ${found} receipt ${found === 1 ? 'line' : 'lines'}. Review before saving.`); render();
  }
  async function savePending(shared) {
    const rows = S.receipts.pending.splice(0);
    if (!rows.length) return;
    if (shared && S.caps.db) {
      try {
        for (let i = 0; i < rows.length; i += 600) await S.caps.db.collection('receipts').add({ name: `${rows.length} ${L('kvitton', 'receipts')} · ${today()}`, createdAt: new Date().toISOString(), createdBy: S.me || null, rows: rows.slice(i, i + 600).map((r) => { const o = Object.assign({}, r); delete o.batch; return o; }) });
        toast(L('Kvittona är sparade i portalen.', 'Receipts saved in the portal.'));
      } catch (e) { S.receipts.local.push(...rows); toast(L('Kunde inte spara gemensamt; kvittona finns kvar lokalt.', 'Could not save shared; receipts kept locally.')); }
    } else S.receipts.local.push(...rows);
    recompute(); render();
  }


  // ---------- AI (Claude) ----------
  // Claude writes and explains; every number it uses comes from the calculation above.
  S.ai = { summary: {}, explain: {}, chat: [], busy: '' };
  let aiAbort = null;
  const aiReady = () => !!S.caps.sample;
  const rnd = (n, d = 0) => (n == null || !Number.isFinite(n) ? null : Math.round(n * 10 ** d) / 10 ** d);
  function mdLite(text) {
    const lines = String(text || '').split(/\n/);
    let html = '', inList = false;
    const inline = (t) => esc(t).replace(/\*\*(.+?)\*\*/g, '<b>$1</b>');
    for (const raw of lines) {
      const l = raw.trim();
      if (/^[-•*]\s+/.test(l)) { if (!inList) { html += '<ul>'; inList = true; } html += `<li>${inline(l.replace(/^[-•*]\s+/, ''))}</li>`; continue; }
      if (inList) { html += '</ul>'; inList = false; }
      if (l) html += `<p>${inline(l.replace(/^#+\s*/, ''))}</p>`;
    }
    if (inList) html += '</ul>';
    return html;
  }
  function aiErrText(e) {
    const c = e && e.code;
    if (c === 'not_granted' || c === 'sampling_disabled') return L('AI är inte tillåtet för den här sidan.', 'AI is not allowed for this page.');
    if (c === 'rate_limited') return L('För många AI-anrop just nu. Försök igen om en stund.', 'Too many AI calls right now. Try again shortly.');
    if (c === 'cancelled') return L('Avbrutet.', 'Stopped.');
    return L('AI-anropet misslyckades', 'The AI call failed') + (e && e.message ? ': ' + e.message : '.');
  }
  function aiContext(lang) {
    const r = S.result, t = r.totals, c = S.ectx, p = periodInfo();
    const mkt = c.fuels.HVO100_SE ? E.compute(S.dataset.lines, Object.assign({}, c, { fuelPlan: Object.assign({}, c.fuelPlan, { hvoFuel: 'HVO100_SE' }) })).totals.kgWtw : null;
    const byType = new Map();
    for (const x of r.rows) {
      const a = byType.get(x.code) || { type: x.name, code: x.code, litres: 0, kg: 0, sek: 0, fuels: new Set(), orders: x.n, tonnes: x.tonnes, km: x.km };
      a.litres += x.liters; a.kg += x.kgWtw; a.sek += x.amount; a.fuels.add(Mo.fuelShort(x.fuel, c.fuels, 'en')); byType.set(x.code, a);
    }
    return {
      customer: { number: c.customer, name: custName(c.customer, lang) }, period: { label: p.label, from: c.from, to: c.to },
      method: 'Cost-based: invoiced amount x fuel share of price / fuel price = litres; litres x well-to-wheel factor = kg CO2e. Same method as the haulier\'s Excel report.',
      totals: { kgCO2e_WTW: rnd(t.kgWtw, 1), kgCO2e_TTW: rnd(t.kgTtw, 1), litres: rnd(t.liters), MWh: rnd(t.mwh, 1), renewableShareOfEnergy: rnd(t.renewableShare, 3), tonnesGoods: rnd(t.tonnes), tonneKm: rnd(t.tkm), kgCO2ePerTonne: rnd(t.kgPerTonne, 3), gCO2ePerTonneKm: t.tkmCoverage >= 0.8 ? rnd(t.gPerTkm, 1) : null, distanceCoverageShareOfAmount: rnd(t.tkmCoverage, 2), invoicedOrders: r.invoicedCount, invoicedSEK: rnd(t.totalInvoiced) },
      byVehicleType: [...byType.values()].sort((a, b) => b.kg - a.kg).map((a) => ({ type: a.type, code: a.code, fuel: [...a.fuels].join('+'), litres: rnd(a.litres), kgCO2e: rnd(a.kg, 1), invoicedSEK: rnd(a.sek), orders: a.orders, tonnes: rnd(a.tonnes), km: rnd(a.km) })),
      byMonth: r.monthly.map((m) => ({ month: m.month, litres: rnd(m.liters), kgCO2e: rnd(m.kgWtw, 1), invoicedSEK: rnd(m.amount) })),
      fuelPrices_SEK_per_litre: { hvo: rnd(c.prices.hvo, 2), diesel: rnd(c.prices.diesel, 2) }, hvoShareOfVolume: rnd(c.fuelPlan.hvoPct, 3),
      emissionFactors: [...new Set(r.rows.map((x) => x.fuel))].map((k) => ({ fuel: c.fuels[k].label.en, wtw_kg_per_litre: c.fuels[k].wtw, ttw_kg_per_litre: c.fuels[k].ttw, source: c.fuels[k].source ? c.fuels[k].source.en : '', validFrom: c.fuels[k].validFrom })),
      excludedFromCalculation: r.excluded.map((e) => ({ type: e.name, code: e.code, invoicedSEK: rnd(e.amount), rows: e.n, reason: 'no fuel share for this vehicle type' })),
      checks: S.checks.filter((x) => x.sev !== 'ok').map((x) => ({ level: x.sev, title: x.en.t, detail: x.en.d })),
      scenarios: { kgCO2e_ifAllDiesel: rnd(S.scenarioKg, 1), kgCO2e_withSwedishMarketAverageHVOFactor: rnd(mkt, 1), tonnes_invoicedTransportsOnly: rnd(t.tonnesInvoiced), tonnes_asExcel_allRowsWithWeight: rnd(t.tonnesExcelBasis), tonnageBasisInUse: c.tonnageBasis }
    };
  }
  const sumKey = () => `${S.ectx.customer}|${S.ectx.from}|${S.ectx.to}|${S.settings.reportLang}`;
  async function aiSummary() {
    if (!aiReady() || S.ai.busy) return;
    const lang = S.settings.reportLang, key = sumKey(), sig = reportSig();
    const data = JSON.stringify(aiContext(lang));
    const prompt = lang === 'en'
      ? `You write the summary paragraph of a climate report that a Swedish haulier sends to a customer. The reader is the customer's sustainability or purchasing lead.\nWrite 2-3 short paragraphs, at most 120 words in total, in English. Plain text: no heading, no greeting, no bullet points, no markdown.\nUse only numbers that appear in DATA and round them sensibly (e.g. 25.1 tonnes). Do not invent causes, targets or actions that are not in DATA.\nInclude: total climate impact well-to-wheel, the two or three vehicle types with the largest share, the fuel and its renewable share, and one sentence on using the figures in the customer's GHG accounts (Scope 3, Category 4).\nWrite as "we" (the haulier) to "you" (the customer).\n\nDATA:\n${data}`
      : `Du skriver sammanfattningen i en klimatrapport som ett svenskt åkeri skickar till en kund. Läsaren är kundens hållbarhets- eller inköpsansvariga.\nSkriv 2–3 korta stycken, högst 120 ord totalt, på svenska. Vanlig text: ingen rubrik, ingen hälsningsfras, inga punktlistor, ingen markdown.\nAnvänd bara siffror som finns i DATA och avrunda rimligt (till exempel 25,1 ton). Hitta inte på orsaker, mål eller åtgärder som inte står i DATA.\nTa med: total klimatpåverkan från källa till hjul (WTW), de två–tre fordonstyper som står för störst del, drivmedlet och dess förnybara andel, och en mening om att siffrorna kan användas i kundens klimatredovisning (Scope 3, kategori 4).\nSkriv i vi-form från åkeriet till kunden (ni).\n\nDATA:\n${data}`;
    S.ai.busy = 'summary'; S.ai.summary[key] = { text: '', sig, live: true }; render();
    aiAbort = new AbortController();
    try {
      const res = await S.caps.sample(prompt, { signal: aiAbort.signal, cache: false, onText: ({ text }) => {
        S.ai.summary[key].text = text;
        const el = document.getElementById('ai-sum-live'); if (el) el.textContent = text;
        const d = document.getElementById('doc-narr'); if (d) d.textContent = text;
      } });
      S.ai.summary[key] = { text: res.text.trim(), sig, live: false };
      toast(L('Claude har skrivit sammanfattningen. Granska den innan du låser.', 'Claude wrote the summary. Review it before locking.'));
    } catch (e) {
      if (e && e.text && e.code !== 'refused') S.ai.summary[key] = { text: e.text, sig, live: false }; else delete S.ai.summary[key];
      if (!e || e.code !== 'cancelled') toast(aiErrText(e));
    }
    S.ai.busy = ''; render();
  }
  function aiSummaryPanel() {
    if (!aiReady()) return `<p class="small muted">${L('AI-funktionerna fungerar när sidan är öppen i Claude, eller i demo-appen med API-nyckel.', 'The AI features work when the page is open in Claude, or in the demo app with an API key.')}</p>`;
    const a = S.ai.summary[sumKey()];
    if (a && a.live) return `<div class="ai-box"><span class="ai-tag">${icon('spark')}Claude ${L('skriver…', 'is writing…')}</span><div id="ai-sum-live" class="ai-live">${esc(a.text)}</div><div class="row"><button class="btn small" data-act="aiStop">${L('Stoppa', 'Stop')}</button></div></div>`;
    if (a && a.text) {
      const stale = a.sig !== reportSig();
      return `<div class="ai-box"><span class="ai-tag">${icon('spark')}${L('AI-utkast', 'AI draft')}</span>${stale ? `<span class="small" style="color:var(--warning-text)">${L('Siffrorna har ändrats sedan texten skrevs; rapporten använder standardtexten tills du skriver om.', 'The figures changed since this was written; the report uses the standard text until you rewrite.')}</span>` : ''}
        <textarea id="ai-sum-edit" rows="8" aria-label="${L('AI-sammanfattning', 'AI summary')}">${esc(a.text)}</textarea>
        <div class="row"><button class="btn small" data-act="aiSummary" ${S.ai.busy ? 'disabled' : ''}>${icon('spark')}${L('Skriv om', 'Rewrite')}</button><button class="btn ghost small" data-act="aiSummaryReset">${L('Använd standardtext', 'Use standard text')}</button></div>
        <small class="muted">${L('Texten ersätter sammanfattningen i rapporten och exporterna. Siffrorna i tabellerna kommer alltid från beräkningen.', 'The text replaces the summary in the report and exports. The figures in the tables always come from the calculation.')}</small></div>`;
    }
    return `<button class="btn" data-act="aiSummary" ${S.ai.busy ? 'disabled' : ''}>${icon('spark')}${L('Skriv sammanfattning med AI', 'Write summary with AI')}</button><p class="tiny muted">${L('Claude skriver texten utifrån de beräknade siffrorna, på rapportens språk.', 'Claude writes the text from the calculated figures, in the report language.')}</p>`;
  }
  async function aiExplain(id) {
    const chk = S.checks.find((x) => x.id === id);
    if (!chk || !aiReady() || S.ai.busy) return;
    const lang = S.lang, T = lang === 'en' ? chk.en : chk.sv;
    const data = JSON.stringify(aiContext(lang));
    const prompt = lang === 'en'
      ? `You help the environmental lead at a Swedish haulier understand an automatic check in the climate-report tool Spårbar.\nAnswer in English, at most 110 words: first one sentence on what the check means, then one sentence on how it affects the customer report, then 2-3 concrete actions as lines starting with "- ".\nUse only facts in CHECK and DATA. No heading.\n\nCHECK (${chk.sev}): ${T.t}\n${T.d}\n\nDATA:\n${data}`
      : `Du hjälper den miljöansvariga på ett svenskt åkeri att förstå en automatisk kontroll i klimatrapportverktyget Spårbar.\nSvara på svenska, högst 110 ord: först en mening om vad kontrollen betyder, sedan en mening om hur den påverkar kundrapporten, sist 2–3 konkreta åtgärder som rader som börjar med "- ".\nAnvänd bara fakta i KONTROLL och DATA. Ingen rubrik.\n\nKONTROLL (${chk.sev}): ${T.t}\n${T.d}\n\nDATA:\n${data}`;
    S.ai.busy = 'explain'; S.ai.explain[id] = { text: '', live: true, t: T.t }; render();
    aiAbort = new AbortController();
    try {
      const res = await S.caps.sample(prompt, { signal: aiAbort.signal, onText: ({ text }) => {
        S.ai.explain[id].text = text; const el = document.getElementById('ai-exp-' + id); if (el) el.innerHTML = mdLite(text);
      } });
      S.ai.explain[id] = { text: res.text.trim(), live: false, t: T.t };
    } catch (e) {
      if (e && e.text && e.code !== 'refused') S.ai.explain[id] = { text: e.text, live: false, t: T.t }; else delete S.ai.explain[id];
      if (!e || e.code !== 'cancelled') toast(aiErrText(e));
    }
    S.ai.busy = ''; render();
  }
  const SUGGEST = [
    ['Vilka fordonstyper står för mest utsläpp?', 'Which vehicle types cause the most emissions?'],
    ['Varför är utsläppen högre på våren än i januari?', 'Why are emissions higher in spring than in January?'],
    ['Vad händer om vi använder svenskt marknadssnitt för HVO?', 'What if we use the Swedish market average for HVO?'],
    ['Vad betyder WTW och TTW för kunden?', 'What do WTW and TTW mean for the customer?'],
    ['Vilka kontroller bör vi lösa innan rapporten skickas?', 'Which checks should we fix before sending the report?']
  ];
  async function aiAsk(q) {
    q = String(q || '').trim();
    if (!q || !aiReady() || S.ai.busy) return;
    const rules = `You are the AI assistant inside Spårbar, the climate-report tool of a Swedish haulier. You answer questions from the haulier's staff about one customer report.
Rules:
- Use only the facts in DATA. If DATA cannot answer, say so and name the data that would be needed.
- Do not recalculate emissions with your own factors. You may do simple arithmetic on numbers in DATA (sums, shares, differences) and use DATA.scenarios.
- Answer in the language of the question (Swedish or English), in at most 150 words. Plain text; use "- " for short lists; no headings.
- Round numbers sensibly and give units (kg CO2e, litres, SEK, tonnes).

DATA (JSON):
${JSON.stringify(aiContext(S.lang))}`;
    const history = S.ai.chat.filter((t) => t.content && !t.live).slice(-10).map((t) => ({ role: t.role, content: t.content }));
    S.ai.chat.push({ role: 'user', content: q });
    const live = { role: 'assistant', content: '', live: true };
    S.ai.chat.push(live);
    S.ai.busy = 'chat'; render(); scrollChat();
    aiAbort = new AbortController();
    try {
      const res = await S.caps.sample([{ role: 'user', content: rules }].concat(history, [{ role: 'user', content: q }]), { signal: aiAbort.signal, cache: false, onText: ({ text }) => {
        live.content = text; const el = document.getElementById('ai-chat-live'); if (el) { el.innerHTML = mdLite(text); scrollChat(); }
      } });
      live.content = res.text.trim(); live.live = false;
    } catch (e) {
      live.live = false;
      if (e && e.text && e.code !== 'refused') live.content = e.text; else { live.content = aiErrText(e); live.error = true; }
    }
    S.ai.busy = ''; render(); scrollChat();
  }
  function scrollChat() { const c = document.getElementById('chat'); if (c) c.scrollTop = c.scrollHeight; }
  function vAssistant() {
    const c = S.ectx, p = periodInfo();
    const i = S.lang === 'en' ? 1 : 0;
    const bubbles = S.ai.chat.map((t) => t.role === 'user'
      ? `<div class="bubble user">${esc(t.content)}</div>`
      : `<div class="bubble ai ${t.error ? 'err' : ''}"><span class="ai-tag">${icon('spark')}Claude</span><div ${t.live ? 'id="ai-chat-live" class="ai-live"' : ''}>${mdLite(t.content)}</div></div>`).join('');
    return `<div class="vhead"><div><h1>${L('AI-assistent', 'AI assistant')}</h1><p>${L('Ställ frågor om rapporten för', 'Ask about the report for')} ${esc(custName(c.customer))}, ${esc(p.uiLabel)}. ${L('Svaren bygger på rapportens beräknade siffror; AI:n räknar inte om utsläppen.', 'Answers are based on the report\'s calculated figures; the AI does not recalculate emissions.')}</p></div>
      ${S.ai.chat.length ? `<button class="btn ghost small" data-act="aiClear" ${S.ai.busy ? 'disabled' : ''}>${L('Rensa', 'Clear')}</button>` : ''}</div>
    <div class="card stack">
      ${aiReady() ? '' : `<div class="note warn">${L('AI-funktionerna fungerar när sidan är öppen i Claude, eller i demo-appen med API-nyckel.', 'The AI features work when the page is open in Claude, or in the demo app with an API key.')}</div>`}
      <div class="row suggest">${SUGGEST.map((q) => `<button class="btn small" data-act="aiAsk" data-q="${esc(q[i])}" ${!aiReady() || S.ai.busy ? 'disabled' : ''}>${esc(q[i])}</button>`).join('')}</div>
      <div class="chat" id="chat" aria-live="polite">${bubbles || `<div class="empty">${icon('spark').replace('<svg', '<svg width="28" height="28" style="stroke:var(--accent);fill:none;stroke-width:1.6"')}<span>${L('Välj en fråga ovan eller skriv en egen.', 'Pick a question above or write your own.')}</span></div>`}</div>
      <form class="chat-in" id="ask-form"><input type="text" id="ai-q" placeholder="${L('Fråga om rapporten…', 'Ask about the report…')}" autocomplete="off" ${!aiReady() ? 'disabled' : ''}><button class="btn primary" type="submit" ${!aiReady() || S.ai.busy ? 'disabled' : ''}>${L('Fråga', 'Ask')}</button>${S.ai.busy === 'chat' ? `<button class="btn" type="button" data-act="aiStop">${L('Stoppa', 'Stop')}</button>` : ''}</form>
      <p class="tiny muted">${L('AI: Claude från Anthropic. Svar kan innehålla fel; kontrollera mot rapporten innan något skickas vidare.', 'AI: Claude by Anthropic. Answers can contain mistakes; check against the report before passing anything on.')}</p>
    </div>`;
  }

  // ---------- rendering ----------
  const VIEWS = [
    ['oversikt', () => L('Översikt', 'Overview'), 'grid'], ['rapport', () => L('Kundrapport', 'Customer report'), 'file'], ['kontroller', () => L('Kontroller', 'Checks'), 'shield'],
    ['assistent', () => L('AI-assistent', 'AI assistant'), 'spark'], ['data', () => L('Data', 'Data'), 'db'], ['kvitton', () => L('Kvitton', 'Receipts'), 'receipt'], ['faktorer', () => L('Faktorer', 'Factors'), 'sliders'], null,
    ['regelverk', () => L('Regelverk', 'Regulations'), 'scale'], ['automatik', () => L('Automatik', 'Automation'), 'flow'], ['arkiv', () => L('Arkiv', 'Archive'), 'archive']
  ];
  function renderNav() {
    const open = S.checks.filter((c) => c.sev === 'critical' || c.sev === 'serious').length;
    $('#nav').innerHTML = VIEWS.map((v) => {
      if (!v) return '<div class="nav-sep" role="separator"></div>';
      const badge = v[0] === 'kontroller' && open ? `<span class="badge" aria-label="${open} ${L('allvarliga', 'serious')}">${open}</span>` : '';
      return `<a href="#${v[0]}" ${S.view === v[0] ? 'aria-current="page"' : ''}>${icon(v[2])}<span>${v[1]()}</span>${badge}</a>`;
    }).join('');
    $('#railfoot').innerHTML = `<div class="seg" role="group" aria-label="${L('Språk', 'Language')}"><button data-act="lang" data-v="sv" aria-pressed="${S.lang === 'sv'}">SV</button><button data-act="lang" data-v="en" aria-pressed="${S.lang === 'en'}">EN</button></div>
      <span class="hide-sm">${L('Prototyp · beräkningar körs i webbläsaren', 'Prototype · calculations run in the browser')}</span>`;
  }
  function statusPill() {
    const st = reportStatus();
    if (st.status === 'locked') return `<span class="pill locked"><span class="dot"></span>${L('Låst', 'Locked')} v${st.version}</span>`;
    if (st.status === 'changed') return `<span class="pill draft"><span class="dot"></span>${L('Ändrad sedan', 'Changed since')} v${st.version}</span>`;
    return `<span class="pill draft"><span class="dot"></span>${L('Utkast', 'Draft')}</span>`;
  }
  function renderCtx() {
    const ds = S.dataset, c = S.ctx;
    const custOpts = ds.customers.map((k) => `<option value="${esc(k)}" ${k === c.customer ? 'selected' : ''}>${esc(k)} · ${esc(custName(k))}</option>`).join('');
    const perOpts = ds.presets.map((p) => `<option value="${p.id}" ${p.id === c.periodId ? 'selected' : ''}>${esc(p.label[S.lang])}</option>`).join('') + `<option value="custom" ${c.periodId === 'custom' ? 'selected' : ''}>${L('Eget intervall…', 'Custom range…')}</option>`;
    $('#ctxbar').innerHTML = `
      <div class="ctx-field"><label for="ctx-customer">${L('Kund', 'Customer')}</label><select id="ctx-customer">${custOpts}</select></div>
      <div class="ctx-field"><label for="ctx-period">${L('Period', 'Period')}</label><select id="ctx-period">${perOpts}</select></div>
      ${c.periodId === 'custom' ? `<div class="ctx-field"><label for="ctx-from">${L('Från – till', 'From – to')}</label><div class="ctx-dates"><input type="date" id="ctx-from" value="${c.from}" min="${ds.minDate}" max="${ds.maxDate}"><input type="date" id="ctx-to" aria-label="${L('Till', 'To')}" value="${c.to}" min="${ds.minDate}" max="${ds.maxDate}"></div></div>` : ''}
      <div class="ctx-spacer"></div>
      <span class="pill ${ds.isSample ? 'sample' : ''}" title="${esc(ds.source)}">${icon('db').replace('<svg', '<svg width="13" height="13" style="stroke:currentColor;fill:none;stroke-width:2"')}${ds.isSample ? L('Exempeldata', 'Sample data') + ': ' : ''}${esc(ds.name)}</span>
      ${statusPill()}`;
  }

  function hbarChart(id) {
    const agg = new Map();
    for (const r of S.result.rows) { const a = agg.get(r.code) || { code: r.code, name: r.name, kg: 0, l: 0, amt: 0 }; a.kg += r.kgWtw; a.l += r.liters; a.amt += r.amount; agg.set(r.code, a); }
    let items = [...agg.values()].sort((a, b) => b.kg - a.kg);
    if (S.ui.chartTable[id]) {
      return `<div class="tablewrap"><table><thead><tr><th>${L('Resurstyp', 'Vehicle type')}</th><th class="r">${L('Liter', 'Litres')}</th><th class="r">kg CO₂e</th><th class="r">${L('Andel', 'Share')}</th></tr></thead><tbody>${items.map((a) => `<tr class="click" tabindex="0" data-act="trace" data-code="${esc(a.code)}"><td>${esc(a.name)}</td><td class="r">${nf(a.l)}</td><td class="r">${nf(a.kg, 1)}</td><td class="r">${pct(a.kg / S.result.totals.kgWtw, 1)}</td></tr>`).join('')}</tbody></table></div>`;
    }
    const top = items.slice(0, 8), rest = items.slice(8);
    if (rest.length) top.push({ code: '', name: L(`Övriga ${rest.length} typer`, `Other ${rest.length} types`), kg: rest.reduce((s, a) => s + a.kg, 0), l: rest.reduce((s, a) => s + a.l, 0), other: true });
    const max = Math.max(...top.map((a) => a.kg), 1);
    const tot = S.result.totals.kgWtw;
    return `<div class="hbars">${top.map((a) => `<button class="hbar" ${a.other ? '' : `data-act="trace" data-code="${esc(a.code)}"`} data-tip="<b>${esc(a.name)}</b><br>${nf(a.kg, 1)} kg CO₂e · ${pct(a.kg / tot, 1)}<br>${nf(a.l)} ${L('liter', 'litres')}${a.other ? '' : '<br>' + L('Klicka för att spåra', 'Click to trace')}" aria-label="${esc(a.name)}: ${nf(a.kg, 0)} kg CO₂e">
      <span class="hbar-name">${esc(a.name)}</span><span class="hbar-track"><span class="hbar-fill ${a.other ? 'other' : ''}" style="width:${(Math.max(0, a.kg) / max * 100).toFixed(2)}%"></span></span><span class="hbar-val">${nf(a.kg / 1000, 2)} t</span></button>`).join('')}</div>`;
  }
  function niceMax(v) { if (v <= 0) return 1; const p = Math.pow(10, Math.floor(Math.log10(v))); const f = v / p; return (f <= 1 ? 1 : f <= 2 ? 2 : f <= 2.5 ? 2.5 : f <= 5 ? 5 : 10) * p; }
  function colChart(id) {
    const m = S.result.monthly;
    if (S.ui.chartTable[id]) {
      return `<div class="tablewrap"><table><thead><tr><th>${L('Månad', 'Month')}</th><th class="r">${L('Liter', 'Litres')}</th><th class="r">kg CO₂e</th></tr></thead><tbody>${m.map((x) => `<tr><td>${esc(monthShort(x.month))} ${x.month.slice(0, 4)}</td><td class="r">${nf(x.liters)}</td><td class="r">${nf(x.kgWtw, 1)}</td></tr>`).join('')}</tbody></table></div>`;
    }
    if (!m.length) return `<div class="empty">${L('Ingen data i perioden.', 'No data in the period.')}</div>`;
    const mx = niceMax(Math.max(...m.map((x) => x.kgWtw)));
    const ticks = [0, 0.5, 1].map((f) => f * mx);
    const maxIdx = m.reduce((bi, x, i) => (x.kgWtw > m[bi].kgWtw ? i : bi), 0);
    return `<div class="cols" role="img" aria-label="${L('Utsläpp per månad', 'Emissions by month')}"><div class="cols-grid">${ticks.map((v) => `<div style="bottom:${(v / mx * 100).toFixed(2)}%"><span>${nf(v / 1000, (mx / 2) % 1000 === 0 ? 0 : 1)} t</span></div>`).join('')}</div>
      ${m.map((x, i) => { const h = Math.max(0, x.kgWtw) / mx * 100; const lab = i === maxIdx || i === m.length - 1; return `<button class="col" data-tip="<b>${esc(monthShort(x.month))} ${x.month.slice(0, 4)}</b><br>${nf(x.kgWtw, 1)} kg CO₂e<br>${nf(x.liters)} ${L('liter', 'litres')}" aria-label="${esc(monthShort(x.month))}: ${nf(x.kgWtw, 0)} kg"><span class="col-fill" style="height:${h.toFixed(2)}%"></span>${lab ? `<span class="col-val" style="bottom:calc(${h.toFixed(2)}% + 4px)">${nf(x.kgWtw / 1000, 1)}</span>` : ''}<span class="col-lbl">${esc(monthShort(x.month))}</span></button>`; }).join('')}</div>`;
  }
  function fuelMixBar() {
    const t = S.result.totals; const bf = t.byFuel; const tot = t.mj || 1;
    const parts = Object.entries(bf).map(([k, v]) => ({ k, mj: v.mj, label: Mo.fuelShort(k, S.ectx.fuels, S.lang), color: S.ectx.fuels[k].shareCol === 'hvo' ? 'var(--hvo)' : 'var(--diesel)', ink: S.ectx.fuels[k].shareCol === 'hvo' ? '#ffffff' : '#1d1408' })).filter((p) => p.mj > 0).sort((a, b) => b.mj - a.mj);
    if (!parts.length) return '';
    return `<div class="segbar" role="img" aria-label="${parts.map((p) => `${p.label} ${pct(p.mj / tot)}`).join(', ')}">${parts.map((p) => `<div style="flex:${p.mj / tot};background:${p.color};color:${p.ink}" data-tip="<b>${esc(p.label)}</b><br>${pct(p.mj / tot, 1)} ${L('av energin', 'of energy')}<br>${nf(p.mj / 3600, 1)} MWh">${p.mj / tot > 0.12 ? pct(p.mj / tot) : ''}</div>`).join('')}</div>
      <div class="legend">${parts.map((p) => `<span><i style="background:${p.color}"></i>${esc(p.label)}</span>`).join('')}</div>`;
  }
  function chartCard(id, title, sub, body) {
    return `<div class="card"><div class="card-h"><div><h3>${title}</h3><p>${sub}</p></div><button class="btn ghost small chart-toggle" data-act="chartTable" data-id="${id}" aria-pressed="${!!S.ui.chartTable[id]}">${S.ui.chartTable[id] ? L('Visa diagram', 'Show chart') : L('Visa tabell', 'Show table')}</button></div>${body}</div>`;
  }
  function levelLadder() {
    const lvl = S.receiptApplied.mix || S.receiptApplied.price ? 2 : 1;
    return `<div class="ladder">
      <div class="${lvl === 1 ? 'on' : ''}"><b>${L('Nivå 3 · Kostnad', 'Tier 3 · Cost')}</b>${L('Belopp × andel ÷ pris', 'Amount × share ÷ price')}</div>
      <div class="${lvl === 2 ? 'on' : ''}"><b>${L('Nivå 2 · Kalibrerad', 'Tier 2 · Calibrated')}</b>${L('Pris och mix från kvitton', 'Price and mix from receipts')}</div>
      <div><b>${L('Nivå 1 · Primärdata', 'Tier 1 · Primary')}</b>${L('Liter per fordon (fas 3)', 'Litres per vehicle (phase 3)')}</div></div>`;
  }

  function vOverview() {
    const r = S.result, t = r.totals, c = S.ectx, p = periodInfo();
    if (!r.invoicedCount) return `<div class="card empty"><h3>${L('Ingen data för vald kund och period', 'No data for this customer and period')}</h3><p>${L('Välj en annan period, eller importera data under Data.', 'Pick another period, or import data under Data.')}</p></div>`;
    const [hv, hu] = kgParts(t.kgWtw);
    const parity = S.parity.applicable ? (S.parity.ok ? `<span class="pill locked" data-tip="${L('Samma resultat som Milj.xlsx, flik Rapport', 'Same result as Milj.xlsx, Rapport sheet')}: ${nf(S.parity.excel, 2)} kg">${icon('check').replace('<svg', '<svg width="13" height="13" style="stroke:currentColor;fill:none;stroke-width:2.4"')}${L('Matchar Excel', 'Matches Excel')}</span>` : `<span class="pill draft">${L('Avviker från Excel', 'Differs from Excel')}</span>`) : '';
    const saved = S.scenarioKg - t.kgWtw;
    const sevCount = (s) => S.checks.filter((x) => x.sev === s).length;
    const topChecks = S.checks.filter((x) => x.sev !== 'ok').slice(0, 4);
    return `
    <div class="vhead"><div><h1>${esc(custName(c.customer))}</h1><p>${esc(p.uiLabel)} · ${c.from} – ${c.to} · ${L('kostnadsbaserad modell, samma som nuvarande Excel-rapport', 'cost-based model, same as the current Excel report')}</p></div></div>
    <section class="grid g-hero">
      <div class="card hero-card">
        <div class="row" style="justify-content:space-between"><span class="hero-label">${L('Klimatpåverkan från källa till hjul (WTW)', 'Climate impact well-to-wheel (WTW)')}</span>${parity}</div>
        <div class="hero-num">${hv}<span class="unit">${hu}</span></div>
        <div class="hero-sub"><span>TTW <b>${kgText(t.kgTtw)}</b></span><span>${L('Drivmedel', 'Fuel')} <b>${nf(t.liters)} ${L('liter', 'litres')}</b></span><span>${L('Energi', 'Energy')} <b>${nf(t.mwh, 0)} MWh</b></span><span>${L('Underlag', 'Basis')} <b>${nf(r.invoicedCount)} ${L('fakturerade rader', 'invoiced rows')}</b></span></div>
        <div class="row"><a class="btn primary" href="#rapport">${icon('file')}${L('Öppna kundrapport', 'Open customer report')}</a><button class="btn" data-act="export" data-kind="pdf" ${S.ui.busy ? 'disabled' : ''}>${icon('download')}PDF</button><button class="btn" data-act="export" data-kind="xlsx" ${S.ui.busy ? 'disabled' : ''}>${icon('download')}Excel</button>${aiReady() ? `<a class="btn ghost" href="#assistent">${icon('spark')}${L('Fråga AI om rapporten', 'Ask AI about the report')}</a>` : ''}</div>
      </div>
      <div class="card stack">
        <h4>${L('Drivmedel, andel av energin', 'Fuel, share of energy')}</h4>
        ${fuelMixBar()}
        <div class="sep"></div>
        <h4>${L('Jämförelse', 'Comparison')}</h4>
        ${saved > 1 ? `<p>${L('Med enbart', 'With only')} ${esc(Mo.fuelShort(c.fuelPlan.dieselFuel, c.fuels, S.lang))}: <b class="num">${kgText(S.scenarioKg)}</b>. ${L('Valet av HVO100 motsvarar', 'Choosing HVO100 corresponds to')} <b class="num">${kgText(-saved)}</b> (${pct(-saved / S.scenarioKg)}).</p><p class="tiny muted">${L('Jämförelsen ingår inte i kundens klimatredovisning.', 'The comparison is not part of the customer\'s GHG accounts.')}</p>` : `<p class="muted">${L('Ingen förnybar andel i perioden.', 'No renewable share in the period.')}</p>`}
      </div>
    </section>
    <section class="grid g-4">
      <div class="card tile"><span class="k">${L('Transporterat gods', 'Goods transported')}</span><span class="v">${nf(t.tonnes)}<small>${L('ton', 'tonnes')}</small></span><span class="d">${c.tonnageBasis === 'excel' ? L('Som Excel: alla rader med vikt', 'As Excel: all rows with weight') : L('Fakturerade transporter, utan bomkörning', 'Invoiced transports, no empty runs')}</span></div>
      <div class="card tile"><span class="k">${L('Intensitet per ton', 'Intensity per tonne')}</span><span class="v">${nf(t.kgPerTonne, 2)}<small>kg CO₂e/t</small></span><span class="d">${L('WTW delat med transporterat gods', 'WTW divided by goods transported')}</span></div>
      <div class="card tile"><span class="k">${L('Intensitet per tonkilometer', 'Intensity per tonne-km')}</span><span class="v">${nf(t.gPerTkm, 1)}<small>g CO₂e/tkm</small></span><span class="d">${t.tkmCoverage < 0.8 ? `<span style="color:var(--warning-text)">${L('Avstånd finns för', 'Distance exists for')} ${pct(t.tkmCoverage)} ${L('av underlaget', 'of the data')}</span>` : `${nf(t.tkm)} tkm`}</span></div>
      <div class="card tile"><span class="k">${L('Förnybart drivmedel', 'Renewable fuel')}</span><span class="v">${pct(t.renewableShare)}</span><span class="d">${L('av energin', 'of energy')} · ${nf(t.mwhRenewable, 0)} MWh</span></div>
    </section>
    <section class="grid g-2">
      ${chartCard('types', L('Utsläpp per fordonstyp', 'Emissions by vehicle type'), L('Ton CO₂e (WTW). Klicka på en stapel för att se orderraderna bakom.', 'Tonnes CO₂e (WTW). Click a bar to see the order lines behind it.'), hbarChart('types'))}
      ${chartCard('months', L('Utsläpp per månad', 'Emissions by month'), L('Ton CO₂e (WTW) per månad i perioden', 'Tonnes CO₂e (WTW) per month in the period'), colChart('months'))}
    </section>
    <section class="grid g-2">
      <div class="card"><div class="card-h"><div><h3>${L('Kontroller före utskick', 'Checks before sending')}</h3><p>${sevCount('serious') + sevCount('critical')} ${L('allvarliga', 'serious')} · ${sevCount('warning')} ${L('varningar', 'warnings')} · ${sevCount('info')} ${L('info', 'info')}</p></div><a class="btn ghost small" href="#kontroller">${L('Alla kontroller', 'All checks')}</a></div>
        <div class="checklist">${topChecks.map(checkRow).join('')}</div></div>
      <div class="card stack"><div class="card-h" style="margin:0"><div><h3>${L('Datakvalitet', 'Data quality')}</h3><p>${L('Nivåer enligt ISO 14083: primärdata går före modellerade data', 'Tiers per ISO 14083: primary data comes before modelled data')}</p></div></div>
        ${levelLadder()}
        <p class="small muted">${L('Nuvarande Excel-metod är nivå 3. Med tankkvitton räknas pris och drivmedelsmix per period automatiskt (nivå 2). Liter per fordon från kvitton, fördelat per kund efter tonkilometer, ger nivå 1.', 'The current Excel method is tier 3. With fuel receipts, price and fuel mix per period are calculated automatically (tier 2). Litres per vehicle from receipts, allocated per customer by tonne-km, gives tier 1.')}</p>
        <div class="row"><a class="btn small" href="#kvitton">${icon('receipt')}${L('Lägg till kvitton', 'Add receipts')}</a><a class="btn ghost small" href="#automatik">${L('Se färdplanen', 'See the roadmap')}</a></div></div>
    </section>`;
  }
  function checkAction(c) {
    if (c.id === 'tonnage' && S.ctx.tonnageBasis === 'excel') return `<button class="btn small" data-act="tonnage" data-v="invoiced">${L('Räkna bara fakturerade', 'Count invoiced only')}</button>`;
    if (c.id === 'tonnage') return `<button class="btn ghost small" data-act="tonnage" data-v="excel">${L('Tillbaka till Excel-läge', 'Back to Excel mode')}</button>`;
    if (c.id === 'hvoEvidence' || c.id === 'price') return `<a class="btn ghost small" href="#kvitton">${L('Till kvitton', 'Go to receipts')}</a>`;
    if (/^excluded-/.test(c.id) || c.id === 'hvoFactor' || c.id === 'factorAge') return `<a class="btn ghost small" href="#faktorer">${L('Till faktorer', 'Go to factors')}</a>`;
    if (c.id === 'distance') return `<a class="btn ghost small" href="#automatik">${L('Lösning i produktion', 'Fix in production')}</a>`;
    return '';
  }
  function checkRow(c) {
    const T = S.lang === 'en' ? c.en : c.sv;
    const ex = S.ai.explain[c.id];
    const exHtml = ex && ex.t === T.t ? `<div class="ai-box" style="margin-top:8px"><span class="ai-tag">${icon('spark')}Claude${ex.live ? ' · ' + L('skriver…', 'writing…') : ''}</span><div id="ai-exp-${esc(c.id)}" class="${ex.live ? 'ai-live' : ''}">${mdLite(ex.text)}</div></div>` : '';
    const aiBtn = c.sev !== 'ok' && aiReady() && !(ex && ex.t === T.t) ? `<button class="btn ghost small" data-act="aiExplain" data-id="${esc(c.id)}" ${S.ai.busy ? 'disabled' : ''}>${icon('spark')}${L('Förklara med AI', 'Explain with AI')}</button>` : '';
    const acts = [checkAction(c), aiBtn].filter(Boolean).join('');
    return `<div class="check"><div>${sevChip(c.sev)}</div><div><b>${esc(T.t)}</b><p>${esc(T.d)}</p>${acts ? `<div class="row" style="margin-top:8px">${acts}</div>` : ''}${exHtml}</div></div>`;
  }

  // ----- report -----
  function docHtml(m) {
    const R = (sv, en) => (m.lang === 'en' ? en : sv);
    const n = (x, d = 0) => nf(x, d, m.lang);
    const pc = (x, d = 0) => n(x * 100, d) + (m.lang === 'en' ? '%' : ' %');
    const t = m.totals;
    const big = Math.abs(t.kgWtw) >= 1000;
    const mark = m.status === 'locked' ? `<span class="draftmark locked">${R('Låst', 'Locked')} v${m.version}</span>` : `<span class="draftmark">${m.status === 'changed' ? R(`Utkast · ändrad sedan v${m.version}`, `Draft · changed since v${m.version}`) : R('Utkast', 'Draft')}</span>`;
    const cost = m.showCost;
    const rows = m.rows.map((r) => `<tr class="click" tabindex="0" data-act="trace" data-code="${esc(r.code)}" title="${R('Spåra raden', 'Trace this row')}"><td>${esc(r.name)}<span class="trace-mark">${esc(r.code)}</span></td><td>${esc(r.fuelLabel)}</td><td class="r">${n(r.liters)}</td><td class="r">${n(r.ef, 3)}</td><td class="r">${n(r.kgWtw, 1)}</td>${cost ? `<td class="r">${n(r.amount)}</td><td class="r">${pc(r.share, 1)}</td><td class="r">${n(r.price, 2)}</td>` : ''}</tr>`).join('');
    return `<article class="doc" aria-label="${R('Förhandsvisning av kundrapport', 'Customer report preview')}">${mark}
      <div class="doc-top"><div><div class="doc-issuer">${esc(m.issuer)}</div><div class="doc-title">${esc(m.title)}</div></div></div>
      <div class="doc-meta">
        <div><div class="lbl">${R('Kund', 'Customer')}</div><div>${esc(m.customer.name)} <span style="color:var(--paper-muted);font-weight:400">(${R('kundnr', 'no.')} ${esc(m.customer.no)})</span></div></div>
        <div><div class="lbl">${R('Period', 'Period')}</div><div>${esc(m.period.label)} · ${m.period.from} – ${m.period.to}</div></div>
        <div><div class="lbl">${R('Rapport-ID', 'Report ID')}</div><div>${esc(m.id)}${m.version ? ' · v' + m.version : ''} · ${m.issued}</div></div>
      </div>
      <div class="doc-kpis">
        <div><div class="k">${R('Klimatpåverkan WTW', 'Climate impact WTW')}</div><div class="v">${big ? n(t.kgWtw / 1000, 1) : n(t.kgWtw)} <small>${big ? R('ton CO₂e', 'tonnes CO₂e') : 'kg CO₂e'}</small></div></div>
        <div><div class="k">${R('Transporterat gods', 'Goods transported')}</div><div class="v">${n(t.tonnes)} <small>${R('ton', 'tonnes')}</small></div></div>
        <div><div class="k">${R('Intensitet', 'Intensity')}</div><div class="v">${n(t.kgPerTonne, 2)} <small>kg CO₂e/t</small></div></div>
        <div><div class="k">${R('Förnybart drivmedel', 'Renewable fuel')}</div><div class="v">${pc(t.renewableShare)} <small>${R('av energin', 'of energy')}</small></div></div>
      </div>
      <p id="doc-narr">${esc(m.narrative)}</p>${m.aiNarrative ? `<span class="ai-mark">${icon('spark')}${R('AI-text, granskad före låsning', 'AI text, reviewed before locking')}</span>` : ''}
      <h3>${R('Klimatpåverkan per fordonstyp', 'Climate impact by vehicle type')}</h3>
      <div style="overflow-x:auto"><table><thead><tr><th>${R('Resurstyp', 'Vehicle type')}</th><th>${R('Drivmedel', 'Fuel')}</th><th class="r">${R('Volym (l)', 'Volume (l)')}</th><th class="r">${R('Faktor kg/l', 'Factor kg/l')}</th><th class="r">kg CO₂e WTW</th>${cost ? `<th class="r">${R('Kundbelopp (kr)', 'Amount (SEK)')}</th><th class="r">${R('Andel', 'Share')}</th><th class="r">${R('Pris kr/l', 'SEK/l')}</th>` : ''}</tr></thead>
      <tbody>${rows}</tbody><tfoot><tr><td>${R('Totalt', 'Total')}</td><td></td><td class="r">${n(t.liters)}</td><td></td><td class="r">${n(t.kgWtw, 1)}</td>${cost ? `<td class="r">${n(t.amount)}</td><td></td><td></td>` : ''}</tr></tfoot></table></div>
      <h3>${R('Per månad', 'By month')}</h3>
      <table style="max-width:420px"><thead><tr><th>${R('Månad', 'Month')}</th><th class="r">${R('Volym (l)', 'Volume (l)')}</th><th class="r">kg CO₂e WTW</th></tr></thead><tbody>${m.monthly.map((x) => `<tr><td>${esc(x.label)}</td><td class="r">${n(x.liters)}</td><td class="r">${n(x.kgWtw, 1)}</td></tr>`).join('')}</tbody></table>
      <h3>${R('Metod', 'Method')}</h3>${m.method.map((x) => `<p>${esc(x)}</p>`).join('')}
      <h3>${R('Emissionsfaktorer', 'Emission factors')}</h3>
      <table><thead><tr><th>${R('Drivmedel', 'Fuel')}</th><th class="r">WTW kg/l</th><th class="r">TTW kg/l</th><th>${R('Gäller från', 'Valid from')}</th><th>${R('Källa', 'Source')}</th></tr></thead><tbody>${m.factors.map((f) => `<tr><td>${esc(f.label)}</td><td class="r">${n(f.wtw, 3)}</td><td class="r">${n(f.ttw, 3)}</td><td>${esc(f.validFrom)}</td><td>${esc(f.source)}</td></tr>`).join('')}</tbody></table>
      <h3>${R('Antaganden och datakvalitet', 'Assumptions and data quality')}</h3><ul>${m.assumptions.map((a) => `<li>${esc(a)}</li>`).join('')}</ul>
      <div class="doc-foot"><div>${R('Verifieringskod (SHA-256)', 'Verification code (SHA-256)')}<br><span class="mono">${esc(m.hashes.report || R('beräknas…', 'calculating…'))}</span></div><div>${R('Underlagets kontrollsumma', 'Input checksum')}<br><span class="mono">${esc(m.hashes.input ? m.hashes.input.slice(0, 16) + '…' : '…')}</span></div><div>${R('Genererad med Spårbar', 'Generated with Spårbar')}<br>${m.generatedAt.replace('T', ' ').slice(0, 16)}</div></div>
    </article>`;
  }
  function vReport() {
    const s = S.settings, c = S.ctx, ec = S.ectx;
    if (!S.result.rows.length) return `<div class="card empty"><h3>${L('Inget att rapportera för vald kund och period', 'Nothing to report for this customer and period')}</h3></div>`;
    ensureHashes();
    const m = modelFor(s.reportLang, false);
    const db = S.caps.db, st = reportStatus();
    const lockBox = !db ? `<p class="small muted">${L('Låsning och arkiv kräver att sidan är öppen i Claude med inloggning.', 'Locking and the archive need the page open in Claude, signed in.')}</p>`
      : S.ui.confirmLock ? `<div class="confirm"><b>${L(`Lås som version ${(st.version || 0) + 1}?`, `Lock as version ${(st.version || 0) + 1}?`)}</b><span>${L('Siffror, faktorer och underlagets kontrollsumma sparas i arkivet och kan inte ändras. En ny beräkning blir en ny version.', 'Figures, factors and the input checksum are stored in the archive and cannot change. A new calculation becomes a new version.')}</span><div class="row"><button class="btn primary small" data-act="lockGo" ${S.ui.busy ? 'disabled' : ''}>${icon('lock')}${L('Lås rapporten', 'Lock the report')}</button><button class="btn small" data-act="lockCancel">${L('Avbryt', 'Cancel')}</button></div></div>`
      : `<button class="btn" data-act="lockAsk" ${st.status === 'locked' ? 'disabled' : ''}>${icon('lock')}${st.status === 'locked' ? L(`Låst som v${st.version}`, `Locked as v${st.version}`) : L('Lås och arkivera', 'Lock and archive')}</button>`;
    const fuelOpt = (cls, sel) => Object.entries(c.fuels).filter(([, f]) => f.shareCol === cls).map(([k, f]) => `<option value="${esc(k)}" ${k === sel ? 'selected' : ''}>${esc(f.label[S.lang] || f.label.sv)} · ${nf(f.wtw, 3)}</option>`).join('');
    return `
    <div class="vhead"><div><h1>${L('Kundrapport', 'Customer report')}</h1><p>${L('Förhandsvisning av det kunden får. Klicka på en rad i tabellen för att se exakt hur den räknats.', 'Preview of what the customer receives. Click a row in the table to see exactly how it was calculated.')}</p></div></div>
    <div class="grid g-side">
      <div class="card stack">
        <div class="panel-sec"><h4>${L('Innehåll', 'Content')}</h4>
          <label class="field"><span>${L('Utfärdare', 'Issuer')}</span><input type="text" id="r-issuer" value="${esc(s.issuer)}" placeholder="${L('Ert företagsnamn', 'Your company name')}"></label>
          <label class="field"><span>${L('Kundnamn i rapporten', 'Customer name in the report')}</span><input type="text" id="r-custname" value="${esc(s.customerNames[c.customer] || '')}" placeholder="${esc(custName(c.customer))}"></label>
          <div class="field"><span>${L('Rapportens språk', 'Report language')}</span><div class="seg" role="group"><button data-act="rlang" data-v="sv" aria-pressed="${s.reportLang === 'sv'}">Svenska</button><button data-act="rlang" data-v="en" aria-pressed="${s.reportLang === 'en'}">English</button></div></div>
          <label class="check-row"><input type="checkbox" id="r-showcost" ${s.showCost ? 'checked' : ''}><span>${L('Visa kostnadsunderlag (belopp, andel, pris) som i nuvarande Excel', 'Show cost basis (amount, share, price) as in the current Excel')}</span></label>
        </div>
        <div class="panel-sec"><h4>${L('Sammanfattning med AI', 'Summary with AI')}</h4>${aiSummaryPanel()}</div>
        <div class="panel-sec"><h4>${L('Beräkning', 'Calculation')}</h4>
          <div class="field"><span>${L('Tonnage', 'Tonnage')}</span><div class="radio-row">
            <label class="check-row"><input type="radio" name="tonnage" value="excel" ${c.tonnageBasis === 'excel' ? 'checked' : ''}><span>${L('Som Excel: alla rader med vikt', 'As Excel: all rows with a weight')} (${nf(S.result.totals.tonnesExcelBasis)} t)</span></label>
            <label class="check-row"><input type="radio" name="tonnage" value="invoiced" ${c.tonnageBasis === 'invoiced' ? 'checked' : ''}><span>${L('Endast fakturerade transporter', 'Invoiced transports only')} (${nf(S.result.totals.tonnesInvoiced)} t)</span></label></div></div>
          <label class="field"><span>${L('Andel HVO100 av volymen', 'HVO100 share of volume')}: <b id="hvopct-val">${pct(ec.fuelPlan.hvoPct)}</b></span><input type="range" id="r-hvopct" min="0" max="100" step="1" value="${Math.round(c.fuelPlan.hvoPct * 100)}" ${S.receiptApplied.mix ? 'disabled' : ''}>${S.receiptApplied.mix ? `<small>${L('Styrs av kvitton', 'Set by receipts')}</small>` : ''}</label>
          <label class="field"><span>${L('Faktor för HVO100', 'HVO100 factor')}</span><select id="f-hvofuel">${fuelOpt('hvo', c.fuelPlan.hvoFuel)}</select></label>
          <label class="field"><span>${L('Faktor för diesel', 'Diesel factor')}</span><select id="f-dieselfuel">${fuelOpt('diesel', c.fuelPlan.dieselFuel)}</select></label>
          <p class="small muted">${L('Pris', 'Price')}: HVO100 ${nf(ec.prices.hvo, 2)} kr/l${S.receiptApplied.price ? ' (' + L('kvitton', 'receipts') + ')' : ''} · ${L('diesel', 'diesel')} ${nf(ec.prices.diesel, 2)} kr/l. <a href="#faktorer">${L('Ändra', 'Change')}</a></p>
        </div>
        <div class="panel-sec"><h4>${L('Exportera', 'Export')}</h4>
          <div class="btn-col">
            <button class="btn" data-act="export" data-kind="pdf" ${S.ui.busy ? 'disabled' : ''}>${icon('download')}PDF</button>
            <button class="btn" data-act="export" data-kind="xlsx" ${S.ui.busy ? 'disabled' : ''}>${icon('download')}Excel</button>
            <button class="btn" data-act="export" data-kind="csv" ${S.ui.busy ? 'disabled' : ''} data-tip="${L('En rad per fordonstyp, för kundens Scope 3-system', 'One row per vehicle type, for the customer\'s Scope 3 system')}">${icon('download')}CSV Scope 3</button>
            <button class="btn" data-act="export" data-kind="json" ${S.ui.busy ? 'disabled' : ''} data-tip="${L('Maskinläsbar deklaration med ISO 14083-struktur', 'Machine-readable declaration with ISO 14083 structure')}">${icon('download')}JSON</button>
          </div>
          ${S.ui.busy && S.ui.busy !== 'lock' ? `<p class="small muted">${L('Skapar fil…', 'Building file…')}</p>` : ''}
          ${!S.caps.downloads ? `<p class="tiny muted">${L('Nedladdning fungerar när sidan är öppen i Claude.', 'Downloads work when the page is open in Claude.')}</p>` : ''}
          <p class="tiny muted">${L('Excel-filen har samma flik "Rapport" som idag, med formler, plus underlag, metod och kontroller.', 'The Excel file has the same "Rapport" sheet as today, with formulas, plus source lines, method and checks.')}</p>
        </div>
        <div class="panel-sec"><h4>${L('Lås och arkivera', 'Lock and archive')}</h4>${lockBox}
          ${S.checks.some((x) => x.sev === 'serious' || x.sev === 'critical') ? `<p class="small" style="color:var(--serious-text)">${L('Det finns allvarliga kontroller kvar.', 'There are serious checks left.')} <a href="#kontroller">${L('Granska', 'Review')}</a></p>` : ''}
          ${S.dirtyShared && db ? `<button class="btn ghost small" data-act="saveShared">${L('Spara inställningarna som gemensam standard', 'Save settings as the shared default')}</button>` : ''}
        </div>
      </div>
      <div class="docwrap" id="docwrap">${docHtml(m)}</div>
    </div>`;
  }

  function vChecks() {
    const by = (s) => S.checks.filter((x) => x.sev === s);
    const sum = ['critical', 'serious', 'warning', 'info', 'ok'].map((s) => `<div class="card tile"><span class="k">${sevChip(s)}</span><span class="v">${by(s).length}</span></div>`).join('');
    return `<div class="vhead"><div><h1>${L('Kontroller', 'Checks')}</h1><p>${L('Körs automatiskt vid varje beräkning. Allvarliga avvikelser bör vara lösta eller förklarade innan rapporten låses.', 'Run automatically on every calculation. Serious issues should be fixed or explained before the report is locked.')}</p></div></div>
      <section class="grid" style="grid-template-columns:repeat(auto-fit,minmax(150px,1fr))">${sum}</section>
      <div class="card"><div class="checklist">${S.checks.map(checkRow).join('')}</div></div>`;
  }

  function vData() {
    const ds = S.dataset, lines = ds.lines;
    const st = {}; lines.forEach((l) => { st[l.status] = (st[l.status] || 0) + 1; });
    const msg = S.ui.importMsg;
    return `<div class="vhead"><div><h1>${L('Data', 'Data')}</h1><p>${L('Underlaget som rapporterna räknas från. I produktion hämtas det nattligt från ert SQL-system; i prototypen läser ni in samma Excel-fil som idag.', 'The data the reports are calculated from. In production it is pulled nightly from your SQL system; in the prototype you load the same Excel file as today.')}</p></div></div>
    <section class="grid g-2">
      <div class="card stack"><div class="card-h" style="margin:0"><h3>${L('Laddat underlag', 'Loaded data')}</h3>${ds.isSample ? `<span class="pill sample">${L('Exempeldata', 'Sample data')}</span>` : `<button class="btn small" data-act="resetSample">${L('Tillbaka till exempeldata', 'Back to sample data')}</button>`}</div>
        <dl class="kv"><dt>${L('Fil', 'File')}</dt><dd>${esc(ds.name)}</dd><dt>${L('Källa', 'Source')}</dt><dd>${esc(ds.source)}</dd><dt>${L('Rader', 'Rows')}</dt><dd>${nf(lines.length)} (${Object.entries(st).map(([k, v]) => `${esc(k)}: ${nf(v)}`).join(' · ')})</dd>
          <dt>${L('Datum', 'Dates')}</dt><dd>${ds.minDate} – ${ds.maxDate}</dd><dt>${L('Kunder', 'Customers')}</dt><dd>${ds.customers.map(esc).join(', ') || '–'}</dd><dt>${L('Tabeller', 'Tables')}</dt><dd>${(ds.tables || []).map(esc).join('<br>')}</dd></dl>
        <p class="small muted">${L('Status F = fakturerad. Bara fakturerade rader bär belopp och ingår i utsläppen. Fritextfältet Interntext läses aldrig in.', 'Status F = invoiced. Only invoiced rows carry amounts and count toward emissions. The Interntext free-text field is never imported.')}</p>
      </div>
      <div class="card stack"><h3>${L('Läs in ny fil', 'Load a new file')}</h3>
        <label class="drop" data-drop="data" for="file-data">${icon('upload')}<b>${L('Släpp Milj.xlsx eller en SQL-export här', 'Drop Milj.xlsx or an SQL export here')}</b><span class="small muted">${L('Excel eller CSV. Flikar med orderrader (som "Antal ton"), fakturarader (som "Indata") och resurstyper (som "Resursdata") hittas automatiskt.', 'Excel or CSV. Sheets with order lines (like "Antal ton"), invoice lines (like "Indata") and vehicle types (like "Resursdata") are found automatically.')}</span><span class="btn small">${L('Välj fil', 'Choose file')}</span><input type="file" id="file-data" accept=".xlsx,.xlsm,.xls,.csv,.txt" multiple></label>
        <p class="tiny muted">${L('Filen läses i din webbläsare och skickas inte vidare.', 'The file is read in your browser and is not sent anywhere.')}</p>
        ${msg ? `<div class="note ${msg.sev === 'ok' ? 'good' : 'warn'}"><b>${esc(msg.t)}</b>${msg.d && msg.d.length ? `<br>${msg.d.map((x) => esc(x)).join('<br>')}` : ''}</div>` : ''}
      </div>
    </section>
    <div class="card"><div class="card-h"><div><h3>${L('Orderrader', 'Order lines')}</h3><p>${L('Filtrera för att hitta enskilda ordrar.', 'Filter to find individual orders.')}</p></div></div>
      <div class="row" style="margin-bottom:12px">
        <select id="d-status" aria-label="Status"><option value="">${L('Alla statusar', 'All statuses')}</option>${Object.keys(st).sort().map((k) => `<option ${S.ui.dStatus === k ? 'selected' : ''}>${esc(k)}</option>`).join('')}</select>
        <select id="d-type" aria-label="${L('Resurstyp', 'Vehicle type')}"><option value="">${L('Alla resurstyper', 'All vehicle types')}</option>${[...new Set(lines.map((l) => l.rtype).filter(Boolean))].sort((a, b) => Number(a) - Number(b)).map((k) => `<option value="${esc(k)}" ${S.ui.dType === k ? 'selected' : ''}>${esc(k)} ${esc(S.ctx.resourceTypes.get(k) ? S.ctx.resourceTypes.get(k).name : '')}</option>`).join('')}</select>
        <input type="search" id="d-search" placeholder="${L('Order, resurs, artikel…', 'Order, resource, article…')}" value="${esc(S.ui.dSearch)}">
      </div><div id="dtable">${dataTable()}</div></div>
    <div class="card stack"><h3>${L('Fält som portalen läser', 'Fields the portal reads')}</h3><p class="small muted">${L('Samma kolumner som i flikarna "Antal ton" och "Resursdata". SQL-vyn i produktion levererar dem med dessa namn.', 'The same columns as the "Antal ton" and "Resursdata" sheets. The production SQL view delivers them under these names.')}</p>
      <div class="tablewrap"><table><thead><tr><th>${L('Kolumn', 'Column')}</th><th>${L('Används till', 'Used for')}</th><th>${L('Krav', 'Required')}</th></tr></thead><tbody>
      ${[['Startdatum', L('Period och månad', 'Period and month'), L('Ja', 'Yes')], ['Status', L('Endast F (fakturerad) bär belopp', 'Only F (invoiced) carries amounts'), L('Ja', 'Yes')], ['Resurstyp', L('Koppling till drivmedelsandel', 'Link to fuel cost share'), L('Ja', 'Yes')], ['Kund, ordertotal / Kundbelopp', L('Kostnadsbaserad volym', 'Cost-based volume'), L('Ja', 'Yes')], ['Kundnummer', L('Vilken kund rapporten gäller', 'Which customer the report is for'), L('Ja', 'Yes')], ['Avstånd', 'tkm, g CO₂e/tkm (ISO 14083)', L('Rekommenderas', 'Recommended')], ['Taxerad kvantitet + enhet', L('Ton gods och tkm', 'Tonnes and tkm'), L('Rekommenderas', 'Recommended')], ['ArtikelId namn', L('Bomkörning känns igen', 'Detects empty runs'), L('Rekommenderas', 'Recommended')], ['Resurs, Leverantörsnummer', L('Spårning och koppling till kvitton', 'Tracing and receipt matching'), L('Rekommenderas', 'Recommended')], ['Resursdata: Andel drivmedel diesel/hvo', L('Drivmedelsandel av kostnaden', 'Fuel share of cost'), L('Ja', 'Yes')]].map((r) => `<tr><td class="mono">${esc(r[0])}</td><td>${r[1]}</td><td>${r[2]}</td></tr>`).join('')}
      </tbody></table></div></div>`;
  }
  function dataTable() {
    const q = S.ui.dSearch.trim().toLowerCase();
    const rows = S.dataset.lines.filter((l) => (!S.ui.dStatus || l.status === S.ui.dStatus) && (!S.ui.dType || l.rtype === S.ui.dType) && (!q || `${l.orderId} ${l.resource} ${l.article} ${l.supplier}`.toLowerCase().includes(q)));
    const per = 50, pages = Math.max(1, Math.ceil(rows.length / per)); S.ui.dPage = Math.min(S.ui.dPage, pages - 1);
    const slice = rows.slice(S.ui.dPage * per, S.ui.dPage * per + per);
    const rt = S.ctx.resourceTypes;
    return `<div class="tablewrap" style="max-height:520px"><table><thead><tr><th>${L('Datum', 'Date')}</th><th>Status</th><th>${L('Order', 'Order')}</th><th>${L('Resurs', 'Resource')}</th><th>${L('Resurstyp', 'Vehicle type')}</th><th>${L('Artikel', 'Article')}</th><th class="r">km</th><th class="r">${L('Kvantitet', 'Qty')}</th><th class="r">${L('Kundbelopp', 'Amount')}</th></tr></thead><tbody>
      ${slice.map((l) => `<tr><td class="num">${l.date}</td><td>${esc(l.status)}</td><td class="code">${esc(l.orderId)}</td><td class="code">${esc(l.resource)}</td><td>${esc(l.rtype)} <span class="muted">${esc(rt.get(l.rtype) ? rt.get(l.rtype).name : '')}</span></td><td>${esc(l.article)}</td><td class="r">${l.km == null ? '' : nf(l.km, 1)}</td><td class="r">${l.qty == null ? '' : nf(l.qty, 2)} ${esc(l.unit || '')}</td><td class="r">${nf(l.amount, 2)}</td></tr>`).join('') || `<tr><td colspan="9" class="muted">${L('Inga rader matchar.', 'No rows match.')}</td></tr>`}
      </tbody></table></div><div class="pager"><span>${nf(rows.length)} ${L('rader', 'rows')} · ${L('sida', 'page')} ${S.ui.dPage + 1}/${pages}</span><button class="btn small" data-act="dpage" data-d="-1" ${S.ui.dPage === 0 ? 'disabled' : ''}>${L('Föregående', 'Previous')}</button><button class="btn small" data-act="dpage" data-d="1" ${S.ui.dPage >= pages - 1 ? 'disabled' : ''}>${L('Nästa', 'Next')}</button></div>`;
  }

  function vReceipts() {
    const rs = S.rstats, all = allReceipts(), pend = S.receipts.pending;
    const aiOk = !!(S.caps.sample && S.caps.sampleLimits && S.caps.sampleLimits.images);
    const fuelRow = (k, label) => { const p = rs.per[k]; return p ? `<tr><td>${label}</td><td class="r">${nf(p.n)}</td><td class="r">${nf(p.liters)}</td><td class="r">${p.avgPrice ? nf(p.avgPrice, 2) : '–'}</td></tr>` : ''; };
    const srcChip = (s) => s === 'example' ? `<span class="pill sample">${L('Exempel', 'Example')}</span>` : s === 'ai' ? `<span class="pill">${icon('spark').replace('<svg', '<svg width="12" height="12" style="stroke:currentColor;fill:none;stroke-width:2"')}AI</span>` : `<span class="pill">${s === 'csv' ? L('Fil', 'File') : esc(s)}</span>`;
    const rrow = (r, i, pending) => `<tr><td class="num">${esc(r.date)}</td><td>${esc(r.fuel)}</td><td class="r">${nf(r.liters, 1)}</td><td class="r">${r.amount != null ? nf(r.amount, 2) : '–'}</td><td class="r">${r.amount != null && r.liters ? nf((r.vatIncluded ? r.amount / 1.25 : r.amount) / r.liters, 2) : r.unitPrice != null ? nf(r.unitPrice, 2) : '–'}</td><td>${esc(r.vehicle || '')}</td><td>${esc(r.supplier || '')}</td><td>${srcChip(r.source)}${r.confidence != null ? ` <span class="tiny muted">${pct(r.confidence)}</span>` : ''}</td>${pending ? `<td><button class="btn ghost small" data-act="dropPending" data-i="${i}" aria-label="${L('Ta bort', 'Remove')}">${icon('x')}</button></td>` : ''}</tr>`;
    const head = (pending) => `<thead><tr><th>${L('Datum', 'Date')}</th><th>${L('Drivmedel', 'Fuel')}</th><th class="r">${L('Liter', 'Litres')}</th><th class="r">${L('Belopp', 'Amount')}</th><th class="r">kr/l ${L('exkl. moms', 'excl. VAT')}</th><th>${L('Fordon', 'Vehicle')}</th><th>${L('Leverantör', 'Supplier')}</th><th>${L('Källa', 'Source')}</th>${pending ? '<th></th>' : ''}</tr></thead>`;
    return `<div class="vhead"><div><h1>${L('Kvitton', 'Receipts')}</h1><p>${L('Tankkvitton gör beräkningen bättre på tre sätt: priset per liter räknas fram per period, drivmedelsmixen bygger på det som faktiskt tankats, och HVO-påståendet har ett underlag.', 'Fuel receipts improve the calculation three ways: the price per litre is derived per period, the fuel mix is based on what was actually bought, and the HVO claim has evidence.')}</p></div></div>
    <section class="grid g-2">
      <div class="card stack"><h3>${L('Läs in kvitton', 'Load receipts')}</h3>
        <label class="drop" data-drop="receipts" for="file-receipts">${icon('upload')}<b>${L('Släpp kvitton här', 'Drop receipts here')}</b><span class="small muted">${L('CSV eller Excel från drivmedelskortet läses direkt. Bilder och PDF läses av Claude.', 'CSV or Excel from the fuel card is read directly. Images and PDFs are read by Claude.')}</span><span class="btn small">${L('Välj filer', 'Choose files')}</span><input type="file" id="file-receipts" accept=".csv,.xlsx,.xls,.txt,.pdf,image/jpeg,image/png,image/webp" multiple></label>
        <div class="row"><button class="btn small" data-act="receiptTemplate">${icon('download')}${L('Mall (CSV)', 'Template (CSV)')}</button><button class="btn ghost small" data-act="exampleReceipts">${L('Visa med exempelkvitton', 'Try with example receipts')}</button></div>
        <p class="small ${aiOk ? '' : 'muted'}">${icon('spark').replace('<svg', '<svg width="14" height="14" style="stroke:var(--accent-text);fill:none;stroke-width:2;vertical-align:-2px"')} ${aiOk ? L('AI-avläsning är på. Claude läser datum, drivmedel, liter, belopp och moms; du granskar innan något sparas.', 'AI reading is on. Claude reads date, fuel, litres, amount and VAT; you review before anything is saved.') : L('AI-avläsning av bilder och PDF fungerar när sidan är öppen i Claude.', 'AI reading of images and PDFs works when the page is open in Claude.')}</p>
        ${S.ui.aiState ? `<div class="note ${S.ui.aiState === 'error' ? 'warn' : S.ui.aiState === 'done' ? 'good' : ''}">${esc(S.ui.aiMsg)}${S.ui.aiState === 'busy' ? ` <button class="btn ghost small" data-act="aiStop">${L('Stoppa', 'Stop')}</button>` : ''}</div>` : ''}
      </div>
      <div class="card stack"><h3>${L('Perioden enligt kvitton', 'The period according to receipts')}</h3>
        ${rs.n ? `<div class="tablewrap"><table><thead><tr><th>${L('Drivmedel', 'Fuel')}</th><th class="r">${L('Kvitton', 'Receipts')}</th><th class="r">${L('Liter', 'Litres')}</th><th class="r">${L('Snitt kr/l', 'Avg SEK/l')}</th></tr></thead><tbody>${fuelRow('hvo', 'HVO100')}${fuelRow('b100', 'B100')}${fuelRow('diesel', 'Diesel')}</tbody></table></div>
          <p class="small">${L('HVO-andel av volymen', 'HVO share of volume')}: <b>${rs.hvoShareOfVolume == null ? '–' : pct(rs.hvoShareOfVolume, 1)}</b></p>` : `<p class="muted">${L('Inga kvitton i vald period ännu.', 'No receipts in the selected period yet.')}</p>`}
        <label class="check-row"><input type="checkbox" id="k-useprice" ${S.useReceiptPrice ? 'checked' : ''} ${rs.n ? '' : 'disabled'}><span>${L('Använd kvittonas snittpris i beräkningen', 'Use the receipts\' average price in the calculation')}</span></label>
        <label class="check-row"><input type="checkbox" id="k-usemix" ${S.useReceiptMix ? 'checked' : ''} ${rs.hvoShareOfVolume == null ? 'disabled' : ''}><span>${L('Använd kvittonas drivmedelsmix för alla fordonstyper', 'Use the receipts\' fuel mix for all vehicle types')}</span></label>
        <label class="check-row"><input type="checkbox" id="k-vat" ${S.vatIncluded ? 'checked' : ''}><span>${L('Belopp i filerna är inklusive moms (25 %)', 'Amounts in the files include VAT (25%)')}</span></label>
        ${S.receiptApplied.price || S.receiptApplied.mix ? `<div class="note good">${L('Kvittona används i rapporten. Beräkningen är nu kalibrerad (nivå 2).', 'Receipts are used in the report. The calculation is now calibrated (tier 2).')}</div>` : ''}
      </div>
    </section>
    ${pend.length ? `<div class="card stack"><div class="card-h" style="margin:0"><div><h3>${L('Att granska', 'To review')} (${pend.length})</h3><p>${L('Kontrollera mot kvittot. Ta bort felaktiga rader innan du sparar.', 'Check against the receipt. Remove wrong rows before saving.')}</p></div><div class="row">${S.caps.db ? `<button class="btn primary small" data-act="savePending" data-shared="1">${L('Spara i portalen', 'Save in the portal')}</button>` : ''}<button class="btn small" data-act="savePending">${L('Använd i den här sessionen', 'Use in this session')}</button><button class="btn ghost small" data-act="clearPending">${L('Släng', 'Discard')}</button></div></div>
      <div class="tablewrap"><table>${head(true)}<tbody>${pend.map((r, i) => rrow(r, i, true)).join('')}</tbody></table></div></div>` : ''}
    <div class="card stack"><div class="card-h" style="margin:0"><div><h3>${L('Alla kvitton', 'All receipts')} (${all.length})</h3><p>${S.caps.db ? L('Sparade kvitton delas med alla som har åtkomst till portalen.', 'Saved receipts are shared with everyone who has access to the portal.') : L('Utan inloggning finns kvittona bara i den här sessionen.', 'Without sign-in, receipts only live in this session.')}</p></div>${S.receipts.local.length ? `<button class="btn ghost small" data-act="clearLocal">${L('Rensa lokala', 'Clear local')}</button>` : ''}</div>
      ${all.length ? `<div class="tablewrap" style="max-height:420px"><table>${head(false)}<tbody>${all.slice().sort((a, b) => (a.date < b.date ? -1 : 1)).map((r) => rrow(r, 0, false)).join('')}</tbody></table></div>` : `<div class="empty">${icon('receipt').replace('<svg', '<svg width="28" height="28" style="stroke:var(--muted);fill:none;stroke-width:1.5"')}<span>${L('Inga kvitton ännu. Släpp en fil ovan eller prova med exempelkvitton.', 'No receipts yet. Drop a file above or try the example receipts.')}</span></div>`}
      ${(S.receiptBatches || []).length ? `<div class="row">${S.receiptBatches.map((b) => `<span class="pill">${esc(b.name)} <button class="btn ghost small" data-act="delBatch" data-id="${esc(b.id)}" aria-label="${L('Ta bort', 'Delete')}">${icon('x')}</button></span>`).join('')}</div>` : ''}
    </div>`;
  }

  function vFactors() {
    const c = S.ctx, rt = c.resourceTypes;
    const used = new Map(); for (const l of E.inScope(S.dataset.lines, S.ectx)) if (S.ectx.statuses.includes(l.status)) used.set(l.rtype, (used.get(l.rtype) || 0) + l.amount);
    const q = S.ui.rtSearch.trim().toLowerCase();
    const types = [...rt.values()].filter((r) => (!S.ui.rtOnlyUsed || used.has(r.code)) && (!q || `${r.code} ${r.name}`.toLowerCase().includes(q))).sort((a, b) => (used.get(b.code) || 0) - (used.get(a.code) || 0) || String(a.code).localeCompare(String(b.code), 'sv', { numeric: true }));
    const pv = (x) => (x == null ? '' : nf(x * 100, 1).replace(/\s/g, ''));
    return `<div class="vhead"><div><h1>${L('Faktorer och antaganden', 'Factors and assumptions')}</h1><p>${L('Allt som påverkar siffrorna finns här, med källa och datum. Ändringar syns direkt i rapporten. Spara som gemensam standard när de ska gälla för alla.', 'Everything that affects the figures is here, with source and date. Changes show in the report at once. Save as the shared default when they should apply to everyone.')}</p></div>
      <div class="row">${S.caps.db ? `<button class="btn ${S.dirtyShared ? 'primary' : ''}" data-act="saveShared">${L('Spara som gemensam standard', 'Save as shared default')}</button>` : ''}<button class="btn ghost" data-act="resetFactors">${L('Återställ till Excel-värden', 'Reset to Excel values')}</button></div></div>
    <section class="grid g-2">
      <div class="card stack"><h3>${L('Drivmedelspris för perioden', 'Fuel price for the period')}</h3>
        <div class="grid g-2" style="gap:12px">
          <label class="field"><span>HVO100 (kr/l)</span><input type="text" inputmode="decimal" id="f-price-hvo" value="${nf(c.prices.hvo, 2)}"></label>
          <label class="field"><span>Diesel MK1 (kr/l)</span><input type="text" inputmode="decimal" id="f-price-diesel" value="${nf(c.prices.diesel, 2)}"></label></div>
        <p class="small muted">${L('Idag hämtas HVO-priset från Skogsprislistan (Team Produktion / Prislistor). Dieselpriset är härlett ur förhållandet mellan drivmedelsandelarna (21,60 × 11,5/15,6); ange verkligt pris eller använd kvitton.', 'Today the HVO price comes from the forestry price list (Team Produktion / Prislistor). The diesel price is derived from the ratio of the fuel shares (21.60 × 11.5/15.6); enter the real price or use receipts.')}</p>
        ${S.receiptApplied.price ? `<div class="note good">${L('Kvittopriser används just nu', 'Receipt prices are in use')}: HVO100 ${nf(S.ectx.prices.hvo, 2)} · diesel ${nf(S.ectx.prices.diesel, 2)} kr/l</div>` : ''}
      </div>
      <div class="card stack"><h3>${L('Drivmedelsmix', 'Fuel mix')}</h3>
        <label class="field"><span>${L('Andel HVO100 av volymen, alla fordonstyper', 'HVO100 share of volume, all vehicle types')}: <b id="hvopct-val">${pct(S.ectx.fuelPlan.hvoPct)}</b></span><input type="range" id="r-hvopct" min="0" max="100" step="1" value="${Math.round(c.fuelPlan.hvoPct * 100)}" ${S.receiptApplied.mix ? 'disabled' : ''}></label>
        <div class="grid g-2" style="gap:12px">
          <label class="field"><span>${L('Faktor för HVO100', 'HVO100 factor')}</span><select id="f-hvofuel">${Object.entries(c.fuels).filter(([, f]) => f.shareCol === 'hvo').map(([k, f]) => `<option value="${esc(k)}" ${k === c.fuelPlan.hvoFuel ? 'selected' : ''}>${esc(f.label[S.lang] || f.label.sv)}</option>`).join('')}</select></label>
          <label class="field"><span>${L('Faktor för diesel', 'Diesel factor')}</span><select id="f-dieselfuel">${Object.entries(c.fuels).filter(([, f]) => f.shareCol === 'diesel').map(([k, f]) => `<option value="${esc(k)}" ${k === c.fuelPlan.dieselFuel ? 'selected' : ''}>${esc(f.label[S.lang] || f.label.sv)}</option>`).join('')}</select></label></div>
        <p class="small muted">${L('Avvikande mix per fordonstyp (till exempel Lejd) anges i tabellen nedan.', 'A different mix per vehicle type (for example hired hauliers) is set in the table below.')}</p>
      </div>
    </section>
    <div class="card stack"><div class="card-h" style="margin:0"><div><h3>${L('Emissionsfaktorer', 'Emission factors')}</h3><p>${L('kg CO₂e per liter. WTW = från källa till hjul, TTW = avgasrör.', 'kg CO₂e per litre. WTW = well-to-wheel, TTW = tank-to-wheel.')}</p></div></div>
      <div class="tablewrap"><table><thead><tr><th>${L('Drivmedel', 'Fuel')}</th><th class="r">WTW</th><th class="r">TTW</th><th class="r">MJ/l</th><th>${L('Gäller från', 'Valid from')}</th><th>${L('Källa', 'Source')}</th></tr></thead><tbody>
      ${Object.entries(c.fuels).map(([k, f]) => `<tr><td><b style="font-weight:600">${esc(f.label[S.lang] || f.label.sv)}</b><br><span class="code">${esc(k)} · ${f.shareCol === 'hvo' ? L('HVO-klass', 'HVO class') : L('dieselklass', 'diesel class')}</span>${k === c.fuelPlan.hvoFuel || k === c.fuelPlan.dieselFuel ? ` <span class="pill locked" style="padding:0 7px">${L('används', 'in use')}</span>` : ''}</td>
        <td class="r"><input type="text" inputmode="decimal" id="ff-${esc(k)}-wtw" data-fuel="${esc(k)}" data-field="wtw" value="${nf(f.wtw, 3)}" aria-label="WTW ${esc(k)}"></td>
        <td class="r"><input type="text" inputmode="decimal" id="ff-${esc(k)}-ttw" data-fuel="${esc(k)}" data-field="ttw" value="${nf(f.ttw, 3)}" aria-label="TTW ${esc(k)}"></td>
        <td class="r"><input type="text" inputmode="decimal" id="ff-${esc(k)}-mj" data-fuel="${esc(k)}" data-field="mj" value="${nf(f.mj, 1)}" aria-label="MJ ${esc(k)}"></td>
        <td><input type="date" id="ff-${esc(k)}-validFrom" data-fuel="${esc(k)}" data-field="validFrom" value="${esc(f.validFrom || '')}" aria-label="${L('Gäller från', 'Valid from')} ${esc(k)}"></td>
        <td class="small" style="min-width:220px">${esc((f.source && (f.source[S.lang] || f.source.sv)) || '')}</td></tr>`).join('')}
      </tbody></table></div>
      <details><summary class="btn ghost small" style="display:inline-flex">${L('Lägg till leverantörsfaktor', 'Add a supplier factor')}</summary>
        <div class="grid g-4" style="gap:10px;margin-top:10px">
          <label class="field"><span>${L('Namn', 'Name')}</span><input type="text" id="nf-label" placeholder="HVO100 Preem 2026"></label>
          <label class="field"><span>${L('Klass', 'Class')}</span><select id="nf-class"><option value="hvo">HVO</option><option value="diesel">Diesel</option></select></label>
          <label class="field"><span>WTW kg/l</span><input type="text" inputmode="decimal" id="nf-wtw"></label>
          <label class="field"><span>TTW kg/l</span><input type="text" inputmode="decimal" id="nf-ttw" value="0"></label>
          <label class="field"><span>MJ/l</span><input type="text" inputmode="decimal" id="nf-mj" value="34"></label>
          <label class="field"><span>${L('Gäller från', 'Valid from')}</span><input type="date" id="nf-date" value="${today()}"></label>
          <label class="field" style="grid-column:span 2"><span>${L('Källa (dokument)', 'Source (document)')}</span><input type="text" id="nf-source" placeholder="${L('Hållbarhetsbesked, leverantör, datum', 'Sustainability declaration, supplier, date')}"></label>
        </div><div class="row" style="margin-top:10px"><button class="btn small" data-act="addFactor">${L('Lägg till', 'Add')}</button></div></details>
    </div>
    <div class="card stack"><div class="card-h" style="margin:0"><div><h3>${L('Resurstyper och drivmedelsandel av kostnaden', 'Vehicle types and fuel share of cost')}</h3><p>${L('Från Resursdata. Andelen anger hur stor del av kundpriset som är drivmedel, för diesel respektive HVO.', 'From Resursdata. The share is the part of the customer price that is fuel, for diesel and HVO.')}</p></div></div>
      <div class="row"><label class="check-row"><input type="checkbox" id="rt-only" ${S.ui.rtOnlyUsed ? 'checked' : ''}><span>${L('Bara typer med fakturerat belopp i perioden', 'Only types with an invoiced amount in the period')}</span></label><input type="search" id="rt-search" placeholder="${L('Sök typ…', 'Search type…')}" value="${esc(S.ui.rtSearch)}"></div>
      <div class="tablewrap" style="max-height:520px"><table><thead><tr><th>${L('Kod', 'Code')}</th><th>${L('Resurstyp', 'Vehicle type')}</th><th class="r">${L('Diesel %', 'Diesel %')}</th><th class="r">HVO %</th><th class="r">${L('HVO-mix %', 'HVO mix %')}</th><th class="r">${L('Fakturerat', 'Invoiced')}</th></tr></thead><tbody>
      ${types.map((r) => { const ov = c.fuelPlan.byType[r.code]; const miss = r.diesel == null && r.hvo == null; return `<tr><td class="code">${esc(r.code)}</td><td>${esc(r.name)}${miss && used.has(r.code) ? ` ${sevChip('serious')}` : ''}${S.rtOverrides[r.code] ? ` <span class="pill draft" style="padding:0 7px">${L('ändrad', 'edited')}</span>` : ''}</td>
        <td class="r"><input type="text" inputmode="decimal" id="rt-${esc(r.code)}-diesel" data-rt="${esc(r.code)}" data-field="diesel" value="${pv(r.diesel)}" placeholder="–" aria-label="Diesel % ${esc(r.name)}"></td>
        <td class="r"><input type="text" inputmode="decimal" id="rt-${esc(r.code)}-hvo" data-rt="${esc(r.code)}" data-field="hvo" value="${pv(r.hvo)}" placeholder="–" aria-label="HVO % ${esc(r.name)}"></td>
        <td class="r"><input type="text" inputmode="decimal" id="rt-${esc(r.code)}-mix" data-rt="${esc(r.code)}" data-field="mix" value="${ov == null ? '' : nf(ov * 100, 0)}" placeholder="${nf(c.fuelPlan.hvoPct * 100, 0)}" aria-label="${L('HVO-mix', 'HVO mix')} ${esc(r.name)}"></td>
        <td class="r">${used.has(r.code) ? nf(used.get(r.code), 0) + ' kr' : '–'}</td></tr>`; }).join('') || `<tr><td colspan="6" class="muted">${L('Inga typer matchar.', 'No types match.')}</td></tr>`}
      </tbody></table></div></div>`;
  }

  const LAWS = [
    { n: ['CSRD efter Omnibus I', 'CSRD after Omnibus I'], ref: 'Direktiv (EU) 2026/470 · ÅRL 6 kap.', cov: 'yes',
      w: ['I kraft 18 mars 2026. Regeringens förslag: krav bara för företag med fler än 1 000 anställda och över 5 mdkr i omsättning, från räkenskapsår 2026 (ikraftträdande 19 mars 2027).', 'In force 18 March 2026. Swedish government proposal: applies only above 1,000 employees and SEK 5 bn turnover, from financial year 2026 (in force 19 March 2027).'],
      h: ['Era stora kunder redovisar inköpta transporter. Portalen ger dem WTW och TTW per period, en metodbeskrivning och ett spår ned till orderraderna som deras revisor kan granska.', 'Your large customers report purchased transport. The portal gives them WTW and TTW per period, a method statement and a trail down to the order lines that their auditor can check.'],
      s: [['Linklaters: Omnibus I i EUT', 'https://sustainablefutures.linklaters.com/post/102mkjf/eu-omnibus-i-directive-published-in-the-official-journal-of-the-eu'], ['Realtid 2026-10-27', 'https://www.realtid.se/juridik/farre-bolag-maste-hallbarhetsrapportera-galler-redan-i-ar/']] },
    { n: ['Värdekedjetaket och den frivilliga standarden', 'Value-chain cap and the voluntary standard'], ref: 'Delegerade akter 3 juli 2026', cov: 'yes',
      w: ['Från räkenskapsår 2027 får rapporterande företag inte begära mer av leverantörer med upp till 1 000 anställda än den frivilliga standarden omfattar. Reviderade ESRS gäller från samma år.', 'From financial year 2027 reporting companies may not ask suppliers with up to 1,000 employees for more than the voluntary standard covers. Revised ESRS apply from the same year.'],
      h: ['Rapporten ger det kunden behöver om transporterna, energi, förnybar andel och växthusgaser, i ett format ni kan skicka till alla kunder utan att bygga nytt för varje förfrågan.', 'The report gives the customer what it needs about the transports, energy, renewable share and greenhouse gases, in one format you can send to every customer without rebuilding for each request.'],
      s: [['Linklaters: reviderade ESRS', 'https://sustainablefutures.linklaters.com/post/102n88h/eu-csrd-commission-adopts-revised-esrs-and-voluntary-reporting-standard']] },
    { n: ['GHG Protocol Scope 3, kategori 4', 'GHG Protocol Scope 3, Category 4'], ref: 'Corporate Value Chain (Scope 3) Standard', cov: 'yes',
      w: ['Kundens redovisning av köpta transporter, uppströms.', 'The customer\'s accounting of purchased upstream transport.'],
      h: ['CSV-exporten har en rad per fordonstyp och drivmedel med WTW, TTW och WTT, klar att läsa in i kundens system.', 'The CSV export has one row per vehicle type and fuel with WTW, TTW and WTT, ready to load into the customer\'s system.'], s: [] },
    { n: ['ISO 14083 och GLEC', 'ISO 14083 and GLEC'], ref: 'EN ISO 14083:2023', cov: 'part',
      w: ['Standarden för utsläpp från transportkedjor, och den metod CountEmissionsEU pekar ut.', 'The standard for transport-chain emissions, and the method CountEmissionsEU points to.'],
      h: ['Klart: WTW och TTW, kg CO₂e per ton, g CO₂e per tonkilometer, datatyp och faktorkällor. Saknas: avstånd på en del ordrar och primärdata per fordon (fas 3).', 'Done: WTW and TTW, kg CO₂e per tonne, g CO₂e per tonne-km, data type and factor sources. Missing: distances on some orders and primary data per vehicle (phase 3).'], s: [] },
    { n: ['CountEmissionsEU', 'CountEmissionsEU'], ref: 'Förordning (EU) 2026/1030', cov: 'part',
      w: ['I kraft 1 juni 2026. Huvuddelen gäller från 2 december 2030. EU:s databaser för faktorer kommer 2028 och 2029.', 'In force 1 June 2026. Most obligations apply from 2 December 2030. The EU factor databases arrive in 2028 and 2029.'],
      h: ['Gäller när transportutsläpp redovisas, även frivilligt eller på kundens begäran. Primärdata ska prioriteras, därför kvittoflödet. Företag som inte är SME behöver extern verifiering. Bevaka om en egen portal räknas som beräkningsverktyg som måste certifieras.', 'Applies whenever transport emissions are disclosed, also voluntarily or at a customer\'s request. Primary data must come first, hence the receipt flow. Non-SMEs need external verification. Watch whether an in-house portal counts as a calculation tool that must be certified.'],
      s: [['ETP-ALICE: förordning (EU) 2026/1030', 'https://www.etp-logistics.eu/?p=29626']] },
    { n: ['Reduktionsplikten', 'Reduction obligation'], ref: 'Lag (2017:1201)', cov: 'yes',
      w: ['10 % för diesel och bensin sedan 1 juli 2025, delvis med elkrediter.', '10% for diesel and petrol since 1 July 2025, partly through electricity credits.'],
      h: ['Dieselns faktor ändras när nivån ändras. Varje faktor har ett giltighetsdatum och portalen varnar när faktorn är äldre än perioden.', 'The diesel factor changes with the level. Each factor carries a valid-from date and the portal warns when a factor is older than the period.'],
      s: [['Drivkraft Sverige: Beräkningsfaktorer 2026', 'https://drivkraftsverige.se/wp-content/uploads/2026/07/Berakningsfaktorer-2026_V1.pdf']] },
    { n: ['Drivmedelslagen och nationella faktorer', 'Fuels Act and national factors'], ref: 'Lag (2011:319) · Energimyndigheten', cov: 'yes',
      w: ['Leverantörer rapporterar växthusgasutsläpp per drivmedel. Snitten publiceras årligen.', 'Suppliers report greenhouse gas emissions per fuel. Averages are published yearly.'],
      h: ['Biblioteket har både leverantörsvärden, som i nuvarande Excel, och svenska marknadssnitt för 2025. Källa och datum följer med i varje rapport.', 'The library holds both supplier values, as in the current Excel, and Swedish 2025 market averages. Source and date follow into every report.'], s: [] },
    { n: ['EU ETS2', 'EU ETS2'], ref: 'Direktiv 2003/87/EG kap. IVa', cov: 'part',
      w: ['Utsläppshandel för vägtrafikens bränslen från 2028. Europeiska rådet bad i oktober 2026 om en ny konsekvensanalys; tidplanen ligger fast.', 'Emissions trading for road fuels from 2028. In October 2026 the European Council asked for a new impact assessment; the timetable stands.'],
      h: ['Fossila liter per kund visas i varje rapport, vilket är underlaget om ni vill föra in en ETS2-klausul i kundavtalen.', 'Fossil litres per customer show in every report, the basis if you want an ETS2 clause in customer contracts.'],
      s: [['Il Foglio 2026-10-16', 'https://www.ilfoglio.it/en/world/2026/10/16/news/what-european-leaders-have-decided-regarding-the-emissions-trading-market--410764']] },
    { n: ['Trafikverkets klimatkrav', 'Swedish Transport Administration climate requirements'], ref: 'TDOK 2015:0480 m.fl.', cov: 'part',
      w: ['Entreprenörer redovisar drivmedelsvolymer via mall eller digitalt. Kravet på förnybart ökar och når 100 % 2030.', 'Contractors report fuel volumes by template or digitally. The renewable requirement rises and reaches 100% in 2030.'],
      h: ['Liter per drivmedel och förnybar andel finns per kund och period. Per kontrakt kräver ett kontraktsfält i SQL (fas 2).', 'Litres per fuel and renewable share exist per customer and period. Per contract needs a contract field in SQL (phase 2).'],
      s: [['Trafikverket: klimatkrav', 'https://bransch.trafikverket.se/contentassets/bd04374d86074cb0aa46a7f540338747/trafikverkets_klimatkrav_original_251028-skrivskyddad.pdf']] },
    { n: ['Klimatdeklaration för byggnader', 'Climate declarations for buildings'], ref: 'Lag (2021:787) · Boverkets förslag', cov: 'plan',
      w: ['Förslag: livscykelberäkning för byggnader över 1 000 m² från 1 januari 2028 och gränsvärden från 2030.', 'Proposal: life-cycle calculation for buildings over 1,000 m² from 1 January 2028 and limit values from 2030.'],
      h: ['Byggkunder behöver transport till byggplats (modul A4) per projekt. Det kräver projektnummer på ordern; exporten planeras i fas 2.', 'Construction customers need transport to site (module A4) per project. That needs a project number on the order; the export is planned in phase 2.'],
      s: [['Boverket', 'https://www.boverket.se/sv/klimatdeklaration/om-klimatdeklaration/nyheter/regler-om-livscykel-gwp-och-gransvarden-for-byggnaders--klimatpaverkan/']] },
    { n: ['Energikartläggning i stora företag', 'Energy audits in large companies'], ref: 'Lag (2014:266)', cov: 'yes',
      w: ['Stora företag kartlägger sin energianvändning, transporter inräknade.', 'Large companies audit their energy use, transport included.'],
      h: ['Energin redovisas i MWh med förnybar andel per kund och period.', 'Energy is reported in MWh with renewable share per customer and period.'], s: [] },
    { n: ['Bokföringslagen', 'Swedish Bookkeeping Act'], ref: 'Lag (1999:1078)', cov: 'part',
      w: ['Räkenskapsinformation sparas i sju år.', 'Accounting records are kept for seven years.'],
      h: ['Låsta rapporter sparas med kontrollsumma för underlaget. I produktion läggs PDF, underlag och kvitton i oföränderlig lagring i sju år.', 'Locked reports are stored with a checksum of the input. In production the PDF, input and receipts go to immutable storage for seven years.'], s: [] },
    { n: ['GDPR', 'GDPR'], ref: 'Förordning (EU) 2016/679', cov: 'yes',
      w: ['Fritext på ordrar kan innehålla namn på personer.', 'Free text on orders can contain people\'s names.'],
      h: ['Fritext läses aldrig in, filer läses i webbläsaren, kunder ser bara sina egna rapporter och produktionen körs i EU.', 'Free text is never imported, files are read in the browser, customers see only their own reports and production runs in the EU.'], s: [] }
  ];
  function covChip(k) { const t = { yes: L('Täckt', 'Covered'), part: L('Delvis', 'Partly'), plan: L('Planerat', 'Planned') }[k]; return `<span class="cov ${k}">${icon(k).replace('<svg', '<svg style="stroke:currentColor;fill:none;stroke-width:2"')}${t}</span>`; }
  function vRules() {
    const i = S.lang === 'en' ? 1 : 0;
    const counts = { yes: 0, part: 0, plan: 0 }; LAWS.forEach((l) => counts[l.cov]++);
    return `<div class="vhead"><div><h1>${L('Regelverk', 'Regulations')}</h1><p>${L('Vad som styr transporternas klimatredovisning i Sverige, läget i november 2026, och hur portalen svarar mot det. Kontrollera alltid mot källan innan ni lovar något i avtal.', 'What governs transport climate reporting in Sweden, as of November 2026, and how the portal answers it. Always check the source before you promise anything in a contract.')}</p></div>
      <div class="row">${covChip('yes')}<span class="muted small">${counts.yes}</span>${covChip('part')}<span class="muted small">${counts.part}</span>${covChip('plan')}<span class="muted small">${counts.plan}</span></div></div>
    <div class="card">${LAWS.map((l) => `<div class="law"><div><h3>${esc(l.n[i])}</h3><div class="ref">${esc(l.ref)}</div><p class="when">${esc(l.w[i])}</p></div><div class="how">${covChip(l.cov)}<p>${esc(l.h[i])}</p>${l.s.length ? `<p class="tiny">${l.s.map((s) => `<a href="${esc(s[1])}" target="_blank" rel="noopener">${esc(s[0])}</a>`).join(' · ')}</p>` : ''}</div></div>`).join('')}</div>`;
  }

  function vAutomation() {
    const steps = [
      [L('Hämta', 'Fetch'), L('Fakturerade orderrader och resurstyper läses nattligt från en läsbehörig SQL-vy.', 'Invoiced order lines and vehicle types are read nightly from a read-only SQL view.'), false],
      [L('Kvitton', 'Receipts'), L('Tankkvitton i e-postlådan läses av AI och kopplas till åkeri, fordon och period.', 'Fuel receipts in the mailbox are read by AI and linked to haulier, vehicle and period.'), false],
      [L('Kontrollera', 'Check'), L('Saknade andelar, krediteringar, dubbelt tonnage, avstånd och faktorernas ålder.', 'Missing shares, credit notes, double tonnage, distances and factor age.'), false],
      [L('Granska', 'Review'), L('Miljöansvarig går igenom avvikelser och godkänner utkasten.', 'The environmental lead reviews issues and approves the drafts.'), true],
      [L('Lås', 'Lock'), L('Version, SHA-256 och underlag sparas oföränderligt i sju år.', 'Version, SHA-256 and inputs are stored immutably for seven years.'), false],
      [L('Leverera', 'Deliver'), L('Kunden får PDF, Excel och CSV via e-post eller kundportal.', 'The customer gets PDF, Excel and CSV by email or customer portal.'), false]
    ];
    const sql = `-- ${L('Läsbehörig vy för Spårbar. Kolumnnamnen följer flikarna "Antal ton" och "Resursdata".', 'Read-only view for Spårbar. Column names follow the "Antal ton" and "Resursdata" sheets.')}
-- ${L('Tabell- och kolumnnamn till höger är exempel och byts mot ert schema.', 'Table and column names on the right are examples; replace with your schema.')}
CREATE VIEW miljo.v_orderrader AS
SELECT
  o.Startdatum          AS [Startdatum],
  o.Status              AS [Status],          -- F = ${L('fakturerad', 'invoiced')}
  o.Resurs              AS [Resurs],
  o.Resurstyp           AS [Resurstyp],
  o.OrderId             AS [OrderId],
  o.KundOrdertotal      AS [Kund, ordertotal],
  o.Artikel             AS [ArtikelId namn],
  o.Avstand             AS [Avstånd],
  o.TaxeradKvantitet    AS [Taxerad kvantitet],
  o.TaxeradEnhet        AS [Taxerad enhet],
  o.BetalareKundId      AS [Kundnummer],
  r.Leverantorsnummer   AS [Leverantörsnummer]
FROM dbo.[Order] o
LEFT JOIN dbo.Resurs r ON r.Resurs = o.Resurs
WHERE o.Startdatum >= DATEADD(year, -2, CAST(GETDATE() AS date));
-- ${L('Interntext tas medvetet inte med: fältet kan innehålla personnamn.', 'Interntext is left out on purpose: the field can hold personal names.')}

CREATE VIEW miljo.v_resurstyper AS
SELECT Resurstyp, Namn, [Andel drivmedel diesel], [Andel drivmedel hvo]
FROM dbo.Resurstyp;`;
    return `<div class="vhead"><div><h1>${L('Automatik och färdplan', 'Automation and roadmap')}</h1><p>${L('Så går rapporten från manuell Excel till ett flöde där människan bara granskar. Prototypen visar beräkning, kontroller, rapport och arkiv; resten byggs i fas 2.', 'How the report moves from manual Excel to a flow where people only review. The prototype shows calculation, checks, report and archive; the rest is built in phase 2.')}</p></div></div>
    <div class="card stack"><div class="card-h" style="margin:0"><div><h3>${L('Flödet varje halvår', 'The flow every half-year')}</h3><p>${L('Körs femte arbetsdagen efter periodens slut för alla kunder med fakturerade transporter. Kvartal och månad på begäran.', 'Runs on the fifth working day after the period ends for every customer with invoiced transports. Quarter and month on request.')}</p></div></div>
      <div class="pipe">${steps.map((s, i) => `<div class="pipe-step ${s[2] ? 'human' : ''}"><span class="n">${String(i + 1).padStart(2, '0')}</span><b>${s[0]}</b><p>${s[1]}</p><span class="who">${s[2] ? L('Människa', 'Person') : L('Automatiskt', 'Automatic')}</span></div>`).join('')}</div>
      <div class="sep"></div>
      <div class="grid g-2"><div><h4>${L('Idag i Excel', 'Today in Excel')}</h4><ul class="small" style="margin:8px 0 0;padding-left:18px;color:var(--ink-2)"><li>${L('Kopiera in SQL-utdrag i fliken "Kopiera in hit" och uppdatera Power Query.', 'Paste the SQL extract into "Kopiera in hit" and refresh Power Query.')}</li><li>${L('Ändra rapportperiod för hand på varje rad.', 'Change the report period by hand on every row.')}</li><li>${L('Kontrollera emissionsvärdet och skriv in snittpris från Skogsprislistan.', 'Check the emission factor and type in the average price from the forestry price list.')}</li><li>${L('Lägg till rader när antalet fordonstyper ändras.', 'Add rows when the number of vehicle types changes.')}</li><li>${L('Tonnage från en separat pivot, skrivet in som ett tal.', 'Tonnage from a separate pivot, typed in as a number.')}</li></ul></div>
        <div><h4>${L('Med portalen', 'With the portal')}</h4><ul class="small" style="margin:8px 0 0;padding-left:18px;color:var(--ink-2)"><li>${L('Underlaget hämtas automatiskt; perioden väljs i en meny.', 'Data is fetched automatically; the period is a menu choice.')}</li><li>${L('Faktorer med källa och datum; varning när de är gamla.', 'Factors with source and date; a warning when they are old.')}</li><li>${L('Pris och drivmedelsmix från kvitton.', 'Price and fuel mix from receipts.')}</li><li>${L('Alla fordonstyper kommer med, matchade på kod i stället för radposition.', 'Every vehicle type is included, matched on code instead of row position.')}</li><li>${L('Tonnage och tkm räknas från samma orderrader, med kontroll av dubbletter.', 'Tonnage and tkm come from the same order lines, with a duplicate check.')}</li></ul></div></div>
    </div>
    <div class="card stack"><h3>${L('Färdplan', 'Roadmap')}</h3><div class="phases">
      <div class="phase now"><h4>${L('Fas 1 · Prototyp (nu)', 'Phase 1 · Prototype (now)')}</h4><ul><li>${L('Samma resultat som Excel, till decimalen', 'Same result as Excel, to the decimal')}</li><li>${L('Kontroller, faktorbibliotek, kvitton', 'Checks, factor library, receipts')}</li><li>${L('PDF, Excel, CSV och JSON', 'PDF, Excel, CSV and JSON')}</li><li>${L('Låsning med version och SHA-256', 'Locking with version and SHA-256')}</li></ul></div>
      <div class="phase"><h4>${L('Fas 2 · Produktion', 'Phase 2 · Production')}</h4><ul><li>${L('SQL-vy och nattlig synk', 'SQL view and nightly sync')}</li><li>${L('Inloggning med Microsoft Entra ID', 'Sign-in with Microsoft Entra ID')}</li><li>${L('Kvitton från e-post via Microsoft Graph', 'Receipts from email via Microsoft Graph')}</li><li>${L('Kundportal och schemalagda utskick', 'Customer portal and scheduled sending')}</li><li>${L('Projekt- och kontraktsnivå (A4, Trafikverket)', 'Project and contract level (A4, Transport Administration)')}</li></ul></div>
      <div class="phase"><h4>${L('Fas 3 · Primärdata', 'Phase 3 · Primary data')}</h4><ul><li>${L('Liter per fordon från kvitton', 'Litres per vehicle from receipts')}</li><li>${L('Fördelning per kund efter tonkilometer', 'Allocation per customer by tonne-km')}</li><li>${L('Avstånd för alla ordrar', 'Distances on every order')}</li><li>${L('Redo för CountEmissionsEU-verifiering', 'Ready for CountEmissionsEU verification')}</li></ul></div>
    </div></div>
    <section class="grid g-2">
      <div class="card stack"><h3>${L('Arkitektur i produktion', 'Production architecture')}</h3>
        <dl class="kv"><dt>${L('Källa', 'Source')}</dt><dd>${L('Ordersystemets SQL Server, två läsbehöriga vyer', 'The order system\'s SQL Server, two read-only views')}</dd>
        <dt>${L('Synk', 'Sync')}</dt><dd>${L('Azure Function med timer, nattligt, till en databas i Sweden Central', 'Azure Function on a timer, nightly, into a database in Sweden Central')}</dd>
        <dt>${L('Kvitton', 'Receipts')}</dt><dd>${L('Delad postlåda → Microsoft Graph → lagring med sju års låsning → AI-avläsning → granskning', 'Shared mailbox → Microsoft Graph → storage locked for seven years → AI reading → review')}</dd>
        <dt>${L('Inloggning', 'Sign-in')}</dt><dd>${L('Entra ID för personal; kunder som gäster eller med engångslänk', 'Entra ID for staff; customers as guests or by one-time link')}</dd>
        <dt>${L('Beräkning', 'Calculation')}</dt><dd>${L('Samma motor som i prototypen (engine.js) körs på servern för PDF och Excel', 'The same engine as the prototype (engine.js) runs on the server for PDF and Excel')}</dd>
        <dt>${L('Spårbarhet', 'Traceability')}</dt><dd>${L('Varje rapport: version, faktorversion, SHA-256 för underlag och rapport', 'Every report: version, factor version, SHA-256 for input and report')}</dd></dl></div>
      <div class="card stack"><h3>SQL</h3><pre class="codeblock">${esc(sql)}</pre></div>
    </section>`;
  }

  function vArchive() {
    const db = S.caps.db;
    if (!db) return `<div class="vhead"><div><h1>${L('Arkiv', 'Archive')}</h1></div></div><div class="card empty">${icon('archive').replace('<svg', '<svg width="30" height="30" style="stroke:var(--muted);fill:none;stroke-width:1.5"')}<h3>${L('Arkivet kräver inloggning', 'The archive needs sign-in')}</h3><p>${L('Öppna sidan i Claude för att låsa rapporter och se tidigare versioner.', 'Open the page in Claude to lock reports and see earlier versions.')}</p></div>`;
    const st = S.archiveState;
    const rows = S.archive;
    const body = st === 'loading' ? `<div class="empty">${L('Hämtar arkivet…', 'Loading the archive…')}</div>`
      : st === 'slow' ? `<div class="empty">${L('Arkivet svarar inte än. Det visas när anslutningen är tillbaka.', 'The archive is not answering yet. It will show when the connection is back.')}</div>`
      : st === 'error' ? `<div class="empty">${L('Arkivet kunde inte läsas.', 'The archive could not be read.')}</div>`
      : !rows.length ? `<div class="empty">${icon('lock').replace('<svg', '<svg width="28" height="28" style="stroke:var(--muted);fill:none;stroke-width:1.5"')}<b>${L('Inga låsta rapporter ännu', 'No locked reports yet')}</b><span>${L('Lås en rapport under Kundrapport så hamnar den här med version och verifieringskod.', 'Lock a report under Customer report and it lands here with version and verification code.')}</span><a class="btn small" href="#rapport">${L('Till kundrapporten', 'Go to the customer report')}</a></div>`
      : `<div class="tablewrap"><table><thead><tr><th>${L('Rapport', 'Report')}</th><th>${L('Kund', 'Customer')}</th><th>${L('Period', 'Period')}</th><th class="r">kg CO₂e</th><th>${L('Låst', 'Locked')}</th><th>${L('Verifieringskod', 'Verification code')}</th><th></th></tr></thead><tbody>
        ${rows.map((a) => { const v = S.ui.verify[a.id]; return `<tr><td><b style="font-weight:600">${esc(a.reportId)}</b> <span class="pill locked" style="padding:0 7px">v${esc(a.version)}</span>${a.dataset && a.dataset.sample ? ` <span class="pill sample" style="padding:0 7px">${L('exempeldata', 'sample data')}</span>` : ''}</td><td>${esc(a.customerName)}</td><td>${esc(a.period)}<br><span class="tiny muted">${esc(a.from)} – ${esc(a.to)}</span></td><td class="r">${nf(a.totals && a.totals.kgWtw, 1)}</td><td class="small">${esc((a.createdAt || '').replace('T', ' ').slice(0, 16))}<br><span class="muted">${esc(a.createdBy ? (S.names[a.createdBy] || L('Kollega', 'Colleague')) : '–')}</span></td><td class="code" title="${esc(a.reportHash)}">${esc(String(a.reportHash || '').slice(0, 12))}…</td>
          <td><div class="row"><button class="btn small" data-act="verify" data-id="${esc(a.id)}">${L('Verifiera', 'Verify')}</button><button class="btn ghost small" data-act="openArchived" data-id="${esc(a.id)}">${L('Visa', 'View')}</button></div>${v ? `<div style="margin-top:6px">${sevChip(v.sev)} <span class="small">${esc(v.t)}</span></div>` : ''}</td></tr>`; }).join('')}</tbody></table></div>`;
    return `<div class="vhead"><div><h1>${L('Arkiv', 'Archive')}</h1><p>${L('Låsta rapporter med version, beräkningsunderlag och kontrollsumma. Verifiera jämför dagens underlag med det som låstes, till exempel inför revision.', 'Locked reports with version, calculation basis and checksum. Verify compares today\'s data with what was locked, for example before an audit.')}</p></div></div><div class="card">${body}</div>`;
  }

  const VIEW_FNS = { oversikt: vOverview, rapport: vReport, kontroller: vChecks, assistent: vAssistant, data: vData, kvitton: vReceipts, faktorer: vFactors, regelverk: vRules, automatik: vAutomation, arkiv: vArchive };
  function render() {
    const fid = document.activeElement && document.activeElement.id;
    ensureHashes();
    renderNav(); renderCtx();
    $('#view').innerHTML = (VIEW_FNS[S.view] || vOverview)();
    if (fid) { const el = document.getElementById(fid); if (el && el !== document.activeElement) try { el.focus({ preventScroll: true }); } catch (e) { /* ignore */ } }
  }
  let softT = null;
  function softRefresh() { clearTimeout(softT); softT = setTimeout(() => { const a = document.activeElement; if (a && /INPUT|TEXTAREA|SELECT/.test(a.tagName) && a.type !== 'checkbox' && a.type !== 'radio') { renderCtx(); return; } render(); }, 40); }

  // ---------- trace drawer ----------
  function renderDrawer() {
    let root = $('#drawer-root');
    const d = S.drawer;
    if (!d) { root.innerHTML = ''; return; }
    let html = '';
    if (d.kind === 'trace') {
      const rows = S.result.rows.filter((r) => r.code === d.code);
      const g = S.result.types.get(d.code);
      if (!g) { S.drawer = null; root.innerHTML = ''; return; }
      const name = rows[0] ? rows[0].name : (S.ectx.resourceTypes.get(d.code) || {}).name || d.code;
      const lines = g.lines.slice().sort((a, b) => (a.date < b.date ? -1 : 1));
      const shown = d.all ? lines : lines.slice(0, 120);
      html = `<div class="row" style="justify-content:space-between;align-items:flex-start"><div><h4>${L('Spårning', 'Trace')}</h4><h2>${esc(name)}</h2><p class="muted small">${L('Resurstyp', 'Vehicle type')} ${esc(d.code)} · ${nf(g.n)} ${L('fakturerade rader', 'invoiced rows')} · ${esc(custName(S.ectx.customer))} · ${S.ectx.from} – ${S.ectx.to}</p></div><button class="btn ghost" data-act="closeDrawer" aria-label="${L('Stäng', 'Close')}">${icon('x')}</button></div>
        ${rows.map((r) => { const f = S.ectx.fuels[r.fuel]; return `<div class="formula">${L('Kundbelopp', 'Customer amount')} ${nf(r.amount, 2)} kr${r.amount !== r.groupAmount ? ` <span class="op">(${pct(r.amount / r.groupAmount)} ${L('av', 'of')} ${nf(r.groupAmount, 2)})</span>` : ''}<br><span class="op">×</span> ${L('drivmedelsandel', 'fuel share')} ${pct(r.share, 1)} <span class="op">=</span> ${nf(r.fuelCost, 2)} kr<br><span class="op">÷</span> ${L('pris', 'price')} ${nf(r.price, 2)} kr/l <span class="op">=</span> ${nf(r.liters, 2)} ${L('liter', 'litres')} ${esc(Mo.fuelShort(r.fuel, S.ectx.fuels, S.lang))}<br><span class="op">×</span> ${L('faktor', 'factor')} ${nf(r.ef, 3)} kg CO₂e/l <span class="op">=</span> <span class="res">${nf(r.kgWtw, 2)} kg CO₂e WTW</span></div>
          <p class="tiny muted">${L('Faktor', 'Factor')}: ${esc(f.label[S.lang] || f.label.sv)} · ${esc((f.source && (f.source[S.lang] || f.source.sv)) || '')}</p>`; }).join('') || `<div class="note warn">${L('Typen saknar drivmedelsandel och ingår inte i utsläppen.', 'The type has no fuel share and is not included in emissions.')}</div>`}
        <dl class="kv"><dt>${L('Ton gods', 'Tonnes')}</dt><dd>${nf(g.tonnes, 1)}</dd><dt>${L('Avstånd (summa)', 'Distance (sum)')}</dt><dd>${nf(g.km, 0)} km</dd><dt>tkm</dt><dd>${nf(g.tkm, 0)}</dd></dl>
        <h3>${L('Orderrader', 'Order lines')}</h3>
        <div class="tablewrap"><table><thead><tr><th>${L('Datum', 'Date')}</th><th>${L('Order', 'Order')}</th><th>${L('Resurs', 'Resource')}</th><th>${L('Artikel', 'Article')}</th><th class="r">km</th><th class="r">${L('Kvantitet', 'Qty')}</th><th class="r">${L('Belopp', 'Amount')}</th></tr></thead><tbody>
        ${shown.map((l) => `<tr><td class="num">${l.date}</td><td class="code">${esc(l.orderId)}</td><td class="code" style="white-space:nowrap">${esc(l.resource)}${l.supplier != null ? ` <span class="tiny">· ${L('lev.', 'sup.')} ${esc(l.supplier)}</span>` : ''}</td><td class="small">${esc(l.article)}</td><td class="r">${l.km == null ? '' : nf(l.km, 1)}</td><td class="r">${l.qty == null ? '' : nf(l.qty, 2)} ${esc(l.unit || '')}</td><td class="r">${nf(l.amount, 2)}</td></tr>`).join('')}
        </tbody><tfoot><tr><td colspan="6">${L('Summa', 'Sum')}</td><td class="r">${nf(g.amount, 2)}</td></tr></tfoot></table></div>
        ${lines.length > shown.length ? `<button class="btn small" data-act="traceAll">${L(`Visa alla ${lines.length} rader`, `Show all ${lines.length} rows`)}</button>` : ''}`;
    } else if (d.kind === 'archived') {
      const a = S.archive.find((x) => x.id === d.id);
      if (!a) { S.drawer = null; root.innerHTML = ''; return; }
      html = `<div class="row" style="justify-content:space-between;align-items:flex-start"><div><h4>${L('Låst rapport', 'Locked report')}</h4><h2>${esc(a.reportId)} · v${esc(a.version)}</h2><p class="muted small">${esc(a.customerName)} · ${esc(a.from)} – ${esc(a.to)}</p></div><button class="btn ghost" data-act="closeDrawer" aria-label="${L('Stäng', 'Close')}">${icon('x')}</button></div>
        <dl class="kv"><dt>kg CO₂e WTW</dt><dd>${nf(a.totals.kgWtw, 2)}</dd><dt>kg CO₂e TTW</dt><dd>${nf(a.totals.kgTtw, 2)}</dd><dt>${L('Liter', 'Litres')}</dt><dd>${nf(a.totals.liters, 1)}</dd><dt>${L('Ton gods', 'Tonnes')}</dt><dd>${nf(a.totals.tonnes, 1)}</dd><dt>${L('Pris', 'Price')}</dt><dd>HVO ${nf(a.prices && a.prices.hvo, 2)} · diesel ${nf(a.prices && a.prices.diesel, 2)} kr/l</dd><dt>${L('HVO-andel', 'HVO share')}</dt><dd>${pct(a.hvoPct || 0)}</dd><dt>${L('Underlag', 'Dataset')}</dt><dd>${esc(a.dataset && a.dataset.name)} (${nf(a.dataset && a.dataset.rows)} ${L('rader', 'rows')})</dd><dt>${L('Rapportens SHA-256', 'Report SHA-256')}</dt><dd class="code">${esc(a.reportHash)}</dd><dt>${L('Underlagets SHA-256', 'Input SHA-256')}</dt><dd class="code">${esc(a.inputHash)}</dd></dl>
        <h3>${L('Rader', 'Rows')}</h3><div class="tablewrap"><table><thead><tr><th>${L('Resurstyp', 'Vehicle type')}</th><th>${L('Drivmedel', 'Fuel')}</th><th class="r">${L('Liter', 'Litres')}</th><th class="r">${L('Faktor', 'Factor')}</th><th class="r">kg CO₂e</th></tr></thead><tbody>${(a.rows || []).map((r) => `<tr><td>${esc(r.name)}</td><td>${esc(r.fuel)}</td><td class="r">${nf(r.liters, 0)}</td><td class="r">${nf(r.ef, 3)}</td><td class="r">${nf(r.kgWtw, 1)}</td></tr>`).join('')}</tbody></table></div>
        ${(a.openChecks || []).length ? `<h3>${L('Öppna kontroller vid låsning', 'Open checks when locked')}</h3><ul class="small">${a.openChecks.map((x) => `<li>${esc(x)}</li>`).join('')}</ul>` : ''}`;
    }
    const had = !!root.firstChild;
    root.innerHTML = `<div class="drawer-back" data-act="closeDrawer"></div><aside class="drawer" role="dialog" aria-modal="true" aria-label="${L('Detaljer', 'Details')}" tabindex="-1">${html}</aside>`;
    if (!had) { const dr = $('.drawer', root); dr && dr.focus(); }
  }

  // ---------- toast & tooltip ----------
  let toastT = null;
  function toast(msg) { const r = $('#toast-root'); r.innerHTML = `<div class="toast" role="status">${esc(msg)}</div>`; clearTimeout(toastT); toastT = setTimeout(() => { r.innerHTML = ''; }, 3600); }
  const tip = $('#tip');
  function showTip(el, x, y) {
    tip.innerHTML = el.getAttribute('data-tip'); tip.hidden = false;
    const w = tip.offsetWidth, h = tip.offsetHeight;
    let left = x + 14, top = y + 14;
    if (left + w > window.innerWidth - 8) left = x - w - 14;
    if (top + h > window.innerHeight - 8) top = y - h - 14;
    tip.style.left = Math.max(8, left) + 'px'; tip.style.top = Math.max(8, top) + 'px';
  }
  document.addEventListener('mousemove', (e) => { const el = e.target.closest && e.target.closest('[data-tip]'); if (el) showTip(el, e.clientX, e.clientY); else tip.hidden = true; });
  document.addEventListener('focusin', (e) => { const el = e.target.closest && e.target.closest('[data-tip]'); if (el) { const r = el.getBoundingClientRect(); showTip(el, r.left + r.width / 2, r.bottom); } else tip.hidden = true; });
  document.addEventListener('scroll', () => { tip.hidden = true; }, true);

  // ---------- events ----------
  const actions = {
    lang: (el) => { S.lang = el.dataset.v; store.set('lang', S.lang); render(); },
    rlang: (el) => { S.settings.reportLang = el.dataset.v; markDirty(); render(); },
    export: (el) => doExport(el.dataset.kind),
    chartTable: (el) => { S.ui.chartTable[el.dataset.id] = !S.ui.chartTable[el.dataset.id]; render(); },
    trace: (el) => { if (!el.dataset.code) return; S.drawer = { kind: 'trace', code: el.dataset.code }; renderDrawer(); },
    traceAll: () => { S.drawer.all = true; renderDrawer(); },
    closeDrawer: () => { S.drawer = null; renderDrawer(); },
    tonnage: (el) => { S.ctx.tonnageBasis = el.dataset.v; markDirty(); recompute(); render(); toast(el.dataset.v === 'invoiced' ? L('Tonnage räknas nu bara på fakturerade transporter.', 'Tonnage now counts invoiced transports only.') : L('Tonnage räknas som i Excel.', 'Tonnage counts as in Excel.')); },
    lockAsk: () => { S.ui.confirmLock = true; render(); },
    lockCancel: () => { S.ui.confirmLock = false; render(); },
    lockGo: () => lockReport(),
    saveShared: () => saveShared(),
    verify: (el) => verifyArchived(el.dataset.id),
    openArchived: (el) => { S.drawer = { kind: 'archived', id: el.dataset.id }; renderDrawer(); },
    resetSample: () => { loadSample(); S.ui.importMsg = null; recompute(); render(); },
    dpage: (el) => { S.ui.dPage = Math.max(0, S.ui.dPage + Number(el.dataset.d)); $('#dtable').innerHTML = dataTable(); },
    exampleReceipts: () => { S.receipts.local = S.receipts.local.filter((r) => r.source !== 'example').concat(exampleReceipts()); recompute(); render(); toast(L('Åtta exempelkvitton är inlästa. Slå på kryssrutorna för att använda dem.', 'Eight example receipts loaded. Tick the boxes to use them.')); },
    receiptTemplate: () => saveFile('kvitton_mall.csv', '﻿datum;drivmedel;liter;belopp_exkl_moms;regnr;leverantör\r\n2026-01-15;HVO100;412,5;8910,00;ABC123;Leverantör AB\r\n'),
    dropPending: (el) => { S.receipts.pending.splice(Number(el.dataset.i), 1); render(); },
    clearPending: () => { S.receipts.pending = []; S.ui.aiState = ''; render(); },
    savePending: (el) => { S.ui.aiState = ''; savePending(!!el.dataset.shared); },
    clearLocal: () => { S.receipts.local = []; recompute(); render(); },
    delBatch: async (el) => { if (!S.caps.db) return; try { await S.caps.db.doc('receipts/' + el.dataset.id).delete(); toast(L('Kvittobatchen är borttagen.', 'The receipt batch was deleted.')); } catch (e) { toast(L('Kunde inte ta bort.', 'Could not delete.')); } },
    aiStop: () => { if (aiCtl) aiCtl.abort(); if (aiAbort) aiAbort.abort(); },
    aiSummary: () => aiSummary(),
    aiSummaryReset: () => { delete S.ai.summary[sumKey()]; render(); },
    aiExplain: (el) => aiExplain(el.dataset.id),
    aiAsk: (el) => aiAsk(el.dataset.q),
    aiClear: () => { S.ai.chat = []; render(); },
    resetFactors: () => { S.ctx.fuels = clone(E.DEFAULT_FUELS); S.ctx.prices = { hvo: 21.6, diesel: 15.92 }; S.ctx.fuelPlan = { hvoFuel: 'HVO100', dieselFuel: 'DIESEL_RP', hvoPct: 1, byType: {} }; S.rtOverrides = {}; rebuildRt(); markDirty(); recompute(); render(); toast(L('Faktorerna är återställda till värdena i Excel.', 'Factors reset to the Excel values.')); },
    addFactor: () => {
      const label = ($('#nf-label').value || '').trim(), cls = $('#nf-class').value, wtw = parseNum($('#nf-wtw').value), ttw = parseNum($('#nf-ttw').value), mj = parseNum($('#nf-mj').value);
      if (!label || wtw == null || ttw == null || !mj) { toast(L('Fyll i namn, WTW, TTW och MJ/l.', 'Fill in name, WTW, TTW and MJ/l.')); return; }
      const key = 'CUSTOM_' + label.toUpperCase().replace(/[^A-Z0-9]+/g, '_').slice(0, 24);
      S.ctx.fuels[key] = { label: { sv: label, en: label }, wtw, ttw, mj, shareCol: cls, bio: cls === 'hvo' ? 1 : 0, source: { sv: $('#nf-source').value || '', en: $('#nf-source').value || '' }, validFrom: $('#nf-date').value || today(), kind: 'custom' };
      markDirty(); recompute(); render(); toast(L('Faktorn är tillagd. Välj den under Drivmedelsmix.', 'Factor added. Choose it under Fuel mix.'));
    }
  };
  document.addEventListener('click', (e) => {
    const el = e.target.closest('[data-act]');
    if (!el || el.disabled) return;
    const fn = actions[el.dataset.act];
    if (fn) { if (el.tagName === 'A' || el.tagName === 'BUTTON' || el.classList.contains('drawer-back') || el.tagName === 'TR') e.preventDefault(); fn(el); }
  });
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && S.drawer) { S.drawer = null; renderDrawer(); }
    if ((e.key === 'Enter' || e.key === ' ') && e.target.matches && e.target.matches('tr[data-act]')) { e.preventDefault(); actions[e.target.dataset.act](e.target); }
  });
  let searchT = null;
  document.addEventListener('input', (e) => {
    const t = e.target;
    if (t.id === 'r-hvopct') { const v = $('#hvopct-val'); if (v) v.textContent = pct(Number(t.value) / 100); }
    if (t.id === 'd-search') { S.ui.dSearch = t.value; S.ui.dPage = 0; clearTimeout(searchT); searchT = setTimeout(() => { $('#dtable').innerHTML = dataTable(); }, 150); }
    if (t.id === 'rt-search') { S.ui.rtSearch = t.value; clearTimeout(searchT); searchT = setTimeout(render, 250); }
  });
  document.addEventListener('change', (e) => {
    const t = e.target, c = S.ctx;
    let recalc = true;
    switch (t.id) {
      case 'ctx-customer': c.customer = t.value; S.ui.confirmLock = false; S.ai.chat = []; S.ai.explain = {}; break;
      case 'ctx-period': {
        if (t.value === 'custom') { c.periodId = 'custom'; } else { const p = S.dataset.presets.find((x) => x.id === t.value); if (p) { c.periodId = p.id; c.from = p.from; c.to = p.to; } }
        S.ui.confirmLock = false; S.ai.chat = []; S.ai.explain = {}; break;
      }
      case 'ctx-from': if (t.value) c.from = t.value; break;
      case 'ctx-to': if (t.value) c.to = t.value; break;
      case 'r-issuer': S.settings.issuer = t.value.trim(); markDirty(); recalc = false; break;
      case 'r-custname': if (t.value.trim()) S.settings.customerNames[c.customer] = t.value.trim(); else delete S.settings.customerNames[c.customer]; markDirty(); recalc = false; break;
      case 'r-showcost': S.settings.showCost = t.checked; markDirty(); recalc = false; break;
      case 'ai-sum-edit': { const a = S.ai.summary[sumKey()]; if (a) a.text = t.value; recalc = false; break; }
      case 'r-hvopct': c.fuelPlan.hvoPct = Number(t.value) / 100; markDirty(); break;
      case 'f-hvofuel': c.fuelPlan.hvoFuel = t.value; markDirty(); break;
      case 'f-dieselfuel': c.fuelPlan.dieselFuel = t.value; markDirty(); break;
      case 'f-price-hvo': { const v = parseNum(t.value); if (v > 0) c.prices.hvo = v; markDirty(); break; }
      case 'f-price-diesel': { const v = parseNum(t.value); if (v > 0) c.prices.diesel = v; markDirty(); break; }
      case 'k-useprice': S.useReceiptPrice = t.checked; break;
      case 'k-usemix': S.useReceiptMix = t.checked; break;
      case 'k-vat': S.vatIncluded = t.checked; break;
      case 'rt-only': S.ui.rtOnlyUsed = t.checked; recalc = false; break;
      case 'd-status': S.ui.dStatus = t.value; S.ui.dPage = 0; $('#dtable').innerHTML = dataTable(); return;
      case 'd-type': S.ui.dType = t.value; S.ui.dPage = 0; $('#dtable').innerHTML = dataTable(); return;
      case 'file-data': if (t.files && t.files.length) importData([...t.files]); return;
      case 'file-receipts': if (t.files && t.files.length) importReceipts([...t.files]); return;
      default:
        if (t.name === 'tonnage') { c.tonnageBasis = t.value; markDirty(); break; }
        if (t.dataset.fuel) {
          const f = c.fuels[t.dataset.fuel]; if (!f) return;
          if (t.dataset.field === 'validFrom') f.validFrom = t.value; else { const v = parseNum(t.value); if (v == null || v < 0) { render(); return; } f[t.dataset.field] = v; }
          markDirty(); break;
        }
        if (t.dataset.rt) {
          const code = t.dataset.rt, fld = t.dataset.field, v = t.value.trim() === '' ? null : parseNum(t.value);
          if (fld === 'mix') { if (v == null) delete c.fuelPlan.byType[code]; else c.fuelPlan.byType[code] = Math.min(1, Math.max(0, v / 100)); }
          else {
            const base = S.dataset.rtBase.get(code) || {};
            const o = Object.assign({}, S.rtOverrides[code] || {}); o[fld] = v == null ? null : v / 100;
            if ((o.diesel === undefined || o.diesel === base.diesel) && (o.hvo === undefined || o.hvo === base.hvo)) delete S.rtOverrides[code]; else S.rtOverrides[code] = o;
            rebuildRt();
          }
          markDirty(); break;
        }
        return;
    }
    if (recalc) recompute();
    render();
  });
  document.addEventListener('submit', (e) => {
    if (e.target && e.target.id === 'ask-form') { e.preventDefault(); const i = document.getElementById('ai-q'); const q = i ? i.value : ''; if (i) i.value = ''; aiAsk(q); }
  });
  // drag & drop
  document.addEventListener('dragover', (e) => { const z = e.target.closest && e.target.closest('[data-drop]'); if (z) { e.preventDefault(); z.classList.add('over'); } });
  document.addEventListener('dragleave', (e) => { const z = e.target.closest && e.target.closest('[data-drop]'); if (z && !z.contains(e.relatedTarget)) z.classList.remove('over'); });
  document.addEventListener('drop', (e) => {
    const z = e.target.closest && e.target.closest('[data-drop]'); if (!z) return;
    e.preventDefault(); z.classList.remove('over');
    const files = [...(e.dataTransfer.files || [])]; if (!files.length) return;
    if (z.dataset.drop === 'data') importData(files); else importReceipts(files);
  });

  // ---------- routing & boot ----------
  function route() {
    const h = (location.hash || '').replace('#', '');
    S.view = VIEW_FNS[h] ? h : 'oversikt';
    S.drawer = null; renderDrawer();
    render();
    window.scrollTo(0, 0);
  }
  window.addEventListener('hashchange', route);
  loadSample();
  recompute();
  S.view = VIEW_FNS[(location.hash || '').replace('#', '')] ? location.hash.slice(1) : 'oversikt';
  render();
  initCaps();
})();
