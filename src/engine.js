/* Spårbar calculation engine — pure functions, no DOM. Runs in the browser and in Node.
 * Method: cost-based fuel model (the method in the current Excel report), extended with
 * activity metrics (tonnes, tkm), a TTW/WTT split, energy, data checks and a fuel-mix plan.
 *
 *   liters  = invoiced customer amount × fuel cost share (per vehicle type) ÷ fuel price (kr/l)
 *   kg CO2e = liters × emission factor (WTW, kg CO2e/l)
 */
(function (root) {
  'use strict';

  // ---------- Default emission-factor library (kg CO2e per litre) ----------
  // shareCol: which fuel-cost-share column of the vehicle-type table applies; also the price class.
  const DEFAULT_FUELS = {
    HVO100: {
      label: { sv: 'HVO100 (leverantörsvärde)', en: 'HVO100 (supplier value)' },
      wtw: 0.56, ttw: 0, mj: 34.0, shareCol: 'hvo', bio: 1,
      source: { sv: 'Leverantörsvärde från nuvarande Excel-rapport (avläst 2026-04-15)', en: 'Supplier value from the current Excel report (read 2026-04-15)' },
      validFrom: '2026-04-15', kind: 'supplier'
    },
    DIESEL_RP: {
      label: { sv: 'Diesel MK1 med reduktionsplikt (leverantörsvärde)', en: 'Diesel MK1 with reduction obligation (supplier value)' },
      wtw: 3.02, ttw: 2.24, mj: 35.7, shareCol: 'diesel', bio: 0,
      source: { sv: 'WTW: leverantörsvärde i nuvarande Excel (2026-04-15). TTW: Drivkraft Sverige, Beräkningsfaktorer 2026 V1', en: 'WTW: supplier value in current Excel (2026-04-15). TTW: Drivkraft Sverige, Calculation factors 2026 V1' },
      validFrom: '2026-04-15', kind: 'supplier'
    },
    HVO100_SE: {
      label: { sv: 'HVO100 (svenskt marknadssnitt 2025)', en: 'HVO100 (Swedish market average 2025)' },
      wtw: 0.350, ttw: 0, mj: 34.0, shareCol: 'hvo', bio: 1,
      source: { sv: 'Drivkraft Sverige, Beräkningsfaktorer 2026 V1 (2026-07-07), Energimyndighetens rapporteringsstatistik 2025', en: 'Drivkraft Sverige, Calculation factors 2026 V1 (2026-07-07), Swedish Energy Agency reporting statistics 2025' },
      validFrom: '2026-07-07', kind: 'national'
    },
    DIESEL_SE26: {
      label: { sv: 'Diesel MK1, prognos 2026 (10 % reduktionsplikt)', en: 'Diesel MK1, 2026 forecast (10% reduction obligation)' },
      wtw: 3.055, ttw: 2.24, mj: 35.7, shareCol: 'diesel', bio: 0,
      source: { sv: 'Drivkraft Sverige, Beräkningsfaktorer 2026 V1 (2026-07-07)', en: 'Drivkraft Sverige, Calculation factors 2026 V1 (2026-07-07)' },
      validFrom: '2026-07-07', kind: 'national'
    },
    DIESEL_FOSSIL: {
      label: { sv: 'Diesel MK1, fossil referens', en: 'Diesel MK1, fossil reference' },
      wtw: 3.384, ttw: 2.54, mj: 36.0, shareCol: 'diesel', bio: 0,
      source: { sv: 'Drivkraft Sverige, Beräkningsfaktorer 2026 V1 (2026-07-07)', en: 'Drivkraft Sverige, Calculation factors 2026 V1 (2026-07-07)' },
      validFrom: '2026-07-07', kind: 'national'
    },
    B100: {
      label: { sv: 'B100 / FAME (marknadssnitt 2025)', en: 'B100 / FAME (market average 2025)' },
      wtw: 0.983, ttw: 0, mj: 33.0, shareCol: 'hvo', bio: 1,
      source: { sv: 'Drivkraft Sverige, Beräkningsfaktorer 2026 V1 (2026-07-07)', en: 'Drivkraft Sverige, Calculation factors 2026 V1 (2026-07-07)' },
      validFrom: '2026-07-07', kind: 'national'
    }
  };

  const UNIT_TONNES = { TON: 1, KG: 0.001 };
  const isEmptyRun = (art) => /bomk/i.test(art || '');

  // ---------- Normalisation ----------
  // Accepts the sample's compact orders, or rows parsed from xlsx/csv (header objects).
  function fromCompact(data) {
    const arts = data.articles, units = data.units, cust = data.meta.customer;
    return data.orders.map((o) => ({
      date: o[0], status: o[1], resource: o[2], rtype: o[3] == null ? null : String(o[3]), orderId: o[4],
      amount: o[5], supplierAmount: o[6], article: arts[o[7]], km: o[8], qty: o[9], unit: units[o[10]],
      supplier: o[11], customer: o[12] != null ? String(o[12]) : cust
    }));
  }

  function resourceTypesFrom(rows) {
    // rows: [code, name, dieselPct, hvoPct] with percentages as in the Excel table (11.5 = 11.5 %)
    const m = new Map();
    for (const r of rows) {
      if (r[0] == null || r[0] === '') continue;
      m.set(String(r[0]), {
        code: String(r[0]), name: r[1] || String(r[0]),
        diesel: r[2] == null || r[2] === '' ? null : Number(r[2]) / 100,
        hvo: r[3] == null || r[3] === '' ? null : Number(r[3]) / 100
      });
    }
    return m;
  }

  const HEADER_MAP = {
    date: ['startdatum', 'datum', 'årmånad', 'armanad', 'date', 'orderdatum', 'leveransdatum'],
    year: ['år', 'ar', 'year'], month: ['månad', 'manad', 'month'],
    status: ['status'], resource: ['resurs', 'fordon', 'resource'],
    rtype: ['resurstyp', 'resurstypid', 'resourcetype'], rtypeName: ['resurstypsbenämning', 'namn'],
    amount: ['kund, ordertotal', 'kundbelopp', 'kund ordertotal', 'belopp', 'amount'],
    supplierAmount: ['leverantör, ordertotal', 'avrbelopp', 'leverantor, ordertotal'],
    article: ['artikelid namn', 'artikel', 'article'], km: ['avstånd', 'avstand', 'km', 'distance'],
    qty: ['taxerad kvantitet', 'kvantitet', 'quantity'], unit: ['taxerad enhet', 'enhet', 'unit'],
    customer: ['kundnummer', 'betalare kundid', 'kundid', 'kund', 'customer'],
    supplier: ['leverantörsnummer', 'leverantorsnummer', 'åkeri', 'supplier'], orderId: ['orderid', 'ordernr']
  };
  function mapHeaders(headers) {
    const norm = headers.map((h) => String(h == null ? '' : h).trim().toLowerCase());
    const idx = {};
    for (const [k, alts] of Object.entries(HEADER_MAP)) {
      const i = norm.findIndex((h) => alts.includes(h));
      if (i >= 0) idx[k] = i;
    }
    return idx;
  }
  function toISODate(v) {
    if (v == null || v === '') return null;
    if (v instanceof Date) { // local-midnight dates from parsers: shift to midday so any timezone lands on the right day
      const d = new Date(v.getTime() - v.getTimezoneOffset() * 6e4 + 12 * 3600e3); return d.toISOString().slice(0, 10);
    }
    if (typeof v === 'number') { // Excel serial date
      const d = new Date(Math.round((Math.floor(v) - 25569) * 864e5)); return d.toISOString().slice(0, 10);
    }
    const s = String(v).trim();
    const m = s.match(/^(\d{4})-(\d{1,2})(?:-(\d{1,2}))?/);
    if (m) return `${m[1]}-${m[2].padStart(2, '0')}-${(m[3] || '1').padStart(2, '0')}`;
    const m2 = s.match(/^(\d{1,2})[./](\d{1,2})[./](\d{4})/);
    if (m2) return `${m2[3]}-${m2[2].padStart(2, '0')}-${m2[1].padStart(2, '0')}`;
    return null;
  }
  const num = (v) => {
    if (v == null || v === '') return null;
    if (typeof v === 'number') return v;
    const n = Number(String(v).replace(/\s/g, '').replace(',', '.'));
    return Number.isFinite(n) ? n : null;
  };
  // Detect what a table is and normalise it. rows[0] = headers.
  function classifyTable(rows) {
    if (!rows || rows.length < 2) return { kind: 'empty' };
    const headers = rows[0];
    const idx = mapHeaders(headers);
    const H = headers.map((h) => String(h || '').toLowerCase());
    if (idx.amount != null && idx.rtype != null && (idx.date != null || idx.year != null)) {
      const lines = [];
      for (const r of rows.slice(1)) {
        let date = idx.date != null ? toISODate(r[idx.date]) : null;
        if (!date && idx.year != null && idx.month != null && r[idx.year]) date = `${r[idx.year]}-${String(r[idx.month]).padStart(2, '0')}-01`;
        if (!date) continue;
        const amount = num(r[idx.amount]);
        lines.push({
          date, status: idx.status != null ? String(r[idx.status] || 'F') : 'F',
          resource: idx.resource != null ? r[idx.resource] : null,
          rtype: r[idx.rtype] == null || r[idx.rtype] === '' ? null : String(r[idx.rtype]),
          orderId: idx.orderId != null ? r[idx.orderId] : null,
          amount: amount == null ? 0 : amount,
          supplierAmount: idx.supplierAmount != null ? num(r[idx.supplierAmount]) : null,
          article: idx.article != null ? String(r[idx.article] || '') : '',
          km: idx.km != null ? num(r[idx.km]) : null, qty: idx.qty != null ? num(r[idx.qty]) : null,
          unit: idx.unit != null ? String(r[idx.unit] || '').toUpperCase() : null,
          customer: idx.customer != null && r[idx.customer] != null && r[idx.customer] !== '' ? String(r[idx.customer]) : null,
          supplier: idx.supplier != null ? r[idx.supplier] : null
        });
      }
      const hasActivity = idx.qty != null && idx.unit != null;
      if (lines.length) return { kind: hasActivity ? 'orders' : 'invoiceLines', lines, blankAmounts: lines.every((l) => !l.amount) };
    }
    if (H.includes('andel drivmedel diesel') && H.includes('andel drivmedel hvo')) {
      const iC = H.indexOf('resurstyp'), iN = H.indexOf('namn'), iD = H.indexOf('andel drivmedel diesel'), iH = H.indexOf('andel drivmedel hvo');
      if (iN >= 0 && iC >= 0) {
        const out = rows.slice(1).filter((r) => r[iC] != null && r[iC] !== '').map((r) => {
          // Excel table uses percent (11.5); Rapportdata uses fractions (0.115) — normalise to percent
          let d = num(r[iD]), h = num(r[iH]);
          if (d != null && d < 1) d *= 100; if (h != null && h < 1) h *= 100;
          return [String(r[iC]), r[iN], d, h];
        });
        return { kind: 'resourceTypes', rows: out };
      }
    }
    return { kind: 'unknown', headers };
  }

  // ---------- Periods ----------
  function periodPresets(minDate, maxDate) {
    const y0 = Number(minDate.slice(0, 4)), y1 = Number(maxDate.slice(0, 4));
    const out = [];
    for (let y = y1; y >= y0; y--) {
      out.push({ id: `H1-${y}`, from: `${y}-01-01`, to: `${y}-06-30`, label: { sv: `Halvår 1 ${y}`, en: `H1 ${y}` }, short: `1-${y}` });
      out.push({ id: `H2-${y}`, from: `${y}-07-01`, to: `${y}-12-31`, label: { sv: `Halvår 2 ${y}`, en: `H2 ${y}` }, short: `2-${y}` });
      out.push({ id: `Y-${y}`, from: `${y}-01-01`, to: `${y}-12-31`, label: { sv: `Helår ${y}`, en: `Full year ${y}` }, short: `${y}` });
      for (let q = 1; q <= 4; q++) {
        const m0 = (q - 1) * 3 + 1, m1 = m0 + 2;
        const last = new Date(Date.UTC(y, m1, 0)).getUTCDate();
        out.push({ id: `Q${q}-${y}`, from: `${y}-${String(m0).padStart(2, '0')}-01`, to: `${y}-${String(m1).padStart(2, '0')}-${last}`, label: { sv: `Kvartal ${q} ${y}`, en: `Q${q} ${y}` }, short: `Q${q}-${y}` });
      }
    }
    return out.filter((p) => p.to >= minDate && p.from <= maxDate);
  }

  // ---------- Core calculation ----------
  function inScope(lines, ctx) {
    return lines.filter((l) => l.date >= ctx.from && l.date <= ctx.to && (ctx.customer == null || l.customer === ctx.customer));
  }

  function hvoPctFor(ctx, code) {
    const p = ctx.fuelPlan.byType && ctx.fuelPlan.byType[code];
    return p == null ? ctx.fuelPlan.hvoPct : p;
  }

  function calcGroup(amount, code, ctx) {
    // returns fuel parts for one vehicle type given an invoiced amount
    const rt = ctx.resourceTypes.get(code);
    const parts = [];
    if (!rt) return { parts, excluded: amount, reason: 'unknownType' };
    const pH = hvoPctFor(ctx, code);
    const plan = [[ctx.fuelPlan.hvoFuel, pH], [ctx.fuelPlan.dieselFuel, 1 - pH]];
    let excluded = 0, reason = null;
    for (const [fk, p] of plan) {
      if (p <= 0) continue;
      const f = ctx.fuels[fk];
      const share = rt[f.shareCol];
      const amt = amount * p;
      if (share == null) { excluded += amt; reason = 'noShare'; continue; }
      const price = ctx.prices[f.shareCol];
      const fuelCost = amt * share;
      const liters = fuelCost / price;
      parts.push({
        fuel: fk, amount: amt, share, fuelCost, price, liters, ef: f.wtw, eft: f.ttw,
        kgWtw: liters * f.wtw, kgTtw: liters * f.ttw, kgWtt: liters * (f.wtw - f.ttw), mj: liters * f.mj, bioMj: liters * f.mj * f.bio
      });
    }
    return { parts, excluded, reason };
  }

  function tonnesOf(l) { const k = UNIT_TONNES[l.unit]; return k && l.qty ? l.qty * k : 0; }

  function compute(lines, ctx) {
    const scope = inScope(lines, ctx);
    const inv = scope.filter((l) => ctx.statuses.includes(l.status));
    const types = new Map();
    const months = new Map();
    for (const l of inv) {
      const code = l.rtype == null ? '—' : l.rtype;
      let g = types.get(code);
      if (!g) { g = { code, amount: 0, n: 0, km: 0, tonnes: 0, tkm: 0, kmCovered: 0, lines: [] }; types.set(code, g); }
      g.amount += l.amount; g.n++; g.lines.push(l);
      const t = isEmptyRun(l.article) ? 0 : tonnesOf(l);
      if (l.km) g.km += l.km;
      g.tonnes += t;
      if (l.km && t) { g.tkm += t * l.km; g.kmCovered += l.amount; }
      const mk = l.date.slice(0, 7);
      let m = months.get(mk); if (!m) { m = new Map(); months.set(mk, m); }
      m.set(code, (m.get(code) || 0) + l.amount);
    }
    const rows = [], excluded = [];
    const tot = { amount: 0, liters: 0, kgWtw: 0, kgTtw: 0, kgWtt: 0, mj: 0, bioMj: 0, fuelCost: 0, byFuel: {}, excludedAmount: 0, tkm: 0, km: 0, kmCoveredAmount: 0, orders: inv.length };
    for (const g of types.values()) {
      const rt = ctx.resourceTypes.get(g.code);
      const name = rt ? rt.name : (g.code === '—' ? '(saknar resurstyp)' : g.code);
      const c = calcGroup(g.amount, g.code, ctx);
      tot.tkm += g.tkm; tot.km += g.km; tot.kmCoveredAmount += g.kmCovered;
      if (c.excluded) { excluded.push({ code: g.code, name, amount: c.excluded, n: g.n, reason: c.reason }); tot.excludedAmount += c.excluded; }
      for (const p of c.parts) {
        if (Math.abs(g.amount) < 0.005) continue; // types with orders but nothing invoiced add no rows
        const row = Object.assign({ code: g.code, name, n: g.n, km: g.km, tonnes: g.tonnes, tkm: g.tkm, groupAmount: g.amount }, p);
        rows.push(row);
        tot.amount += p.amount; tot.liters += p.liters; tot.kgWtw += p.kgWtw; tot.kgTtw += p.kgTtw; tot.kgWtt += p.kgWtt; tot.mj += p.mj; tot.bioMj += p.bioMj; tot.fuelCost += p.fuelCost;
        const bf = tot.byFuel[p.fuel] || (tot.byFuel[p.fuel] = { liters: 0, kgWtw: 0, mj: 0 });
        bf.liters += p.liters; bf.kgWtw += p.kgWtw; bf.mj += p.mj;
      }
    }
    rows.sort((a, b) => b.kgWtw - a.kgWtw);
    // Monthly series
    const monthly = [...months.keys()].sort().map((mk) => {
      let kg = 0, liters = 0, amount = 0;
      for (const [code, amt] of months.get(mk)) {
        const c = calcGroup(amt, code, ctx);
        for (const p of c.parts) { kg += p.kgWtw; liters += p.liters; amount += p.amount; }
      }
      return { month: mk, kgWtw: kg, liters, amount };
    });
    // Tonnage
    const tonnesExcelBasis = scope.reduce((s, l) => s + tonnesOf(l), 0);
    const tonnesInvoiced = inv.reduce((s, l) => s + (isEmptyRun(l.article) ? 0 : tonnesOf(l)), 0);
    const tonnes = ctx.tonnageBasis === 'excel' ? tonnesExcelBasis : tonnesInvoiced;
    const totalInvoiced = inv.reduce((s, l) => s + l.amount, 0);
    tot.tonnes = tonnes; tot.tonnesExcelBasis = tonnesExcelBasis; tot.tonnesInvoiced = tonnesInvoiced;
    tot.kgPerTonne = tonnes ? tot.kgWtw / tonnes : null;
    tot.gPerTkm = tot.tkm ? tot.kgWtw * 1000 / tot.tkm : null;
    tot.tkmCoverage = totalInvoiced ? tot.kmCoveredAmount / totalInvoiced : 0;
    tot.renewableShare = tot.mj ? tot.bioMj / tot.mj : 0;
    tot.mwh = tot.mj / 3600;
    tot.mwhRenewable = tot.bioMj / 3600;
    tot.totalInvoiced = totalInvoiced;
    return { rows, excluded, monthly, totals: tot, scopeCount: scope.length, invoicedCount: inv.length, types };
  }

  // ---------- Receipts ----------
  function normalizeFuel(s) {
    const t = String(s || '').toLowerCase();
    if (/hvo|evolution diesel|neste my|förnybar diesel|renewable diesel/.test(t)) return 'hvo';
    if (/b100|fame|rme/.test(t)) return 'b100';
    if (/diesel|mk1|dieselbr|eco ?diesel/.test(t)) return 'diesel';
    if (/adblue/.test(t)) return 'adblue';
    if (/bensin|95|98|petrol/.test(t)) return 'petrol';
    return 'other';
  }
  function receiptStats(receipts, from, to, vatIncluded) {
    const per = {}; let n = 0;
    for (const r of receipts) {
      if (!r.date || r.date < from || r.date > to) continue;
      const k = r.fuelClass || normalizeFuel(r.fuel);
      if (!['hvo', 'diesel', 'b100'].includes(k)) continue;
      const liters = Number(r.liters) || 0; if (liters <= 0) continue;
      let amount = r.amount != null ? Number(r.amount) : (r.unitPrice != null ? Number(r.unitPrice) * liters : null);
      if (amount != null && (r.vatIncluded != null ? r.vatIncluded : vatIncluded)) amount = amount / 1.25;
      const p = per[k] || (per[k] = { liters: 0, amount: 0, pricedLiters: 0, n: 0 });
      p.liters += liters; p.n++; n++;
      if (amount != null && Number.isFinite(amount)) { p.amount += amount; p.pricedLiters += liters; }
    }
    for (const p of Object.values(per)) p.avgPrice = p.pricedLiters ? p.amount / p.pricedLiters : null;
    const lHvo = (per.hvo ? per.hvo.liters : 0) + (per.b100 ? per.b100.liters : 0);
    const lDie = per.diesel ? per.diesel.liters : 0;
    return { per, n, hvoShareOfVolume: lHvo + lDie ? lHvo / (lHvo + lDie) : null };
  }

  // ---------- Checks ----------
  function checks(lines, res, ctx, extra) {
    extra = extra || {};
    const out = [];
    const add = (sev, id, sv, en, data) => out.push({ sev, id, sv, en, data: data || null });
    const scope = inScope(lines, ctx);
    const inv = scope.filter((l) => ctx.statuses.includes(l.status));
    const fmt = (n, d = 0) => new Intl.NumberFormat('sv-SE', { maximumFractionDigits: d, minimumFractionDigits: d }).format(n);
    const fmtE = (n, d = 0) => new Intl.NumberFormat('en-GB', { maximumFractionDigits: d, minimumFractionDigits: d }).format(n);

    if (!inv.length) {
      add('critical', 'noData', { t: 'Inga fakturerade rader i vald period', d: 'Välj en annan period eller kund, eller importera data för perioden.' }, { t: 'No invoiced rows in the selected period', d: 'Pick another period or customer, or import data for the period.' });
      return out;
    }
    // Excluded vehicle types
    for (const e of res.excluded) {
      add(Math.abs(e.amount) > 10000 ? 'serious' : 'warning', 'excluded-' + e.code,
        { t: `${e.name} (${e.code}) saknar drivmedelsandel`, d: `${fmt(e.amount, 2)} kr på ${e.n} rader räknas inte med. Lägg till andelar under Faktorer om typen förbrukar drivmedel.` },
        { t: `${e.name} (${e.code}) has no fuel cost share`, d: `${fmtE(e.amount, 2)} SEK on ${e.n} rows is left out. Add shares under Factors if this type uses fuel.` }, { code: e.code });
    }
    // Credit notes / negative amounts on included types
    const neg = inv.filter((l) => l.amount < 0 && !res.excluded.some((e) => e.code === l.rtype));
    if (neg.length) {
      const s = neg.reduce((a, l) => a + l.amount, 0);
      add('warning', 'negatives', { t: `${neg.length} krediteringar minskar utsläppen`, d: `Negativa belopp på ${fmt(s, 2)} kr dras av i kostnadsmodellen. Kontrollera att de avser prisjusteringar och inte ångrade körningar.` },
        { t: `${neg.length} credit notes reduce emissions`, d: `Negative amounts of ${fmtE(s, 2)} SEK are subtracted in the cost model. Check that they are price adjustments, not cancelled trips.` });
    }
    // Tonnage double counting
    const nonInvTon = scope.filter((l) => !ctx.statuses.includes(l.status)).reduce((s, l) => s + tonnesOf(l), 0);
    if (nonInvTon > 0) {
      const sevT = ctx.tonnageBasis === 'excel' ? 'serious' : 'info';
      add(sevT, 'tonnage', { t: `${fmt(nonInvTon)} ton kommer från ej fakturerade rader`, d: `Rader med status ${[...new Set(scope.filter((l) => !ctx.statuses.includes(l.status)).map((l) => l.status))].join(', ')} (t.ex. materialrader) bär vikt men inget transportbelopp. Nuvarande Excel räknar med dem (${fmt(res.totals.tonnesExcelBasis)} t). Endast fakturerade transporter: ${fmt(res.totals.tonnesInvoiced)} t.` + (ctx.tonnageBasis === 'excel' ? ' Byt till "Endast fakturerade" om de är dubbletter.' : '') },
        { t: `${fmtE(nonInvTon)} t come from non-invoiced rows`, d: `Rows with status ${[...new Set(scope.filter((l) => !ctx.statuses.includes(l.status)).map((l) => l.status))].join(', ')} (e.g. material lines) carry weight but no transport amount. The current Excel counts them (${fmtE(res.totals.tonnesExcelBasis)} t). Invoiced transports only: ${fmtE(res.totals.tonnesInvoiced)} t.` + (ctx.tonnageBasis === 'excel' ? ' Switch to "Invoiced only" if they are duplicates.' : '') });
    }
    // Empty runs
    const bom = inv.filter((l) => isEmptyRun(l.article));
    if (bom.length) {
      const t = bom.reduce((s, l) => s + tonnesOf(l), 0);
      add('info', 'emptyRuns', { t: `${bom.length} bomkörningar`, d: `Tomkörning ger utsläpp men inget gods. ${fmt(t)} t på dessa rader räknas inte som transporterat gods i rensat läge.` },
        { t: `${bom.length} empty runs`, d: `Empty running causes emissions but carries no goods. ${fmtE(t)} t on these rows is not counted as goods in clean mode.` });
    }
    // Non-weight units
    const other = inv.filter((l) => l.unit && !UNIT_TONNES[l.unit]);
    if (other.length) {
      const by = {}; other.forEach((l) => { by[l.unit] = (by[l.unit] || 0) + 1; });
      add('info', 'units', { t: `${other.length} rader debiteras per ${Object.keys(by).join('/')}`, d: 'De ingår i utsläppen men inte i tonnage och tonkilometer.' },
        { t: `${other.length} rows are charged per ${Object.keys(by).join('/')}`, d: 'They count toward emissions but not toward tonnes and tonne-km.' });
    }
    // Distance coverage
    const zeroKm = inv.filter((l) => !l.km);
    if (zeroKm.length) {
      const amt = zeroKm.reduce((s, l) => s + l.amount, 0);
      add(res.totals.tkmCoverage < 0.8 ? 'warning' : 'info', 'distance', { t: `${zeroKm.length} rader saknar avstånd`, d: `${fmt(amt / res.totals.totalInvoiced * 100)} % av beloppet har 0 km, så g CO₂e/tkm bygger på ${fmt(res.totals.tkmCoverage * 100)} % av underlaget. Komplettera avstånd i ordersystemet för ISO 14083-intensitet.` },
        { t: `${zeroKm.length} rows have no distance`, d: `${fmtE(amt / res.totals.totalInvoiced * 100)}% of the amount has 0 km, so g CO₂e/tkm rests on ${fmtE(res.totals.tkmCoverage * 100)}% of the data. Fill in distances in the order system for ISO 14083 intensity.` });
    }
    // Fuel assumption evidence
    const pH = ctx.fuelPlan.hvoPct;
    if (pH > 0 && !extra.receiptsUsed) {
      add('warning', 'hvoEvidence', { t: `${fmt(pH * 100)} % HVO100 antas utan kvitton`, d: 'Påståendet om förnybart drivmedel behöver styrkas med tankkvitton eller leverantörens hållbarhetsbesked, särskilt för Lejd (inhyrda åkerier). Importera kvitton under Kvitton.' },
        { t: `${fmtE(pH * 100)}% HVO100 assumed without receipts`, d: 'The renewable-fuel claim needs fuel receipts or the supplier\'s sustainability declaration, especially for hired hauliers (Lejd). Import receipts under Receipts.' });
    }
    // Factor sanity
    const fH = ctx.fuels[ctx.fuelPlan.hvoFuel];
    const ref = DEFAULT_FUELS.HVO100_SE.wtw;
    if (fH && Math.abs(fH.wtw - ref) / ref > 0.25 && pH > 0) {
      add('info', 'hvoFactor', { t: `HVO-faktorn ${fmt(fH.wtw, 2)} avviker från marknadssnittet ${fmt(ref, 2)}`, d: 'Svenskt snitt 2025 är 0,35 kg CO₂e/l (Drivkraft Sverige 2026). Ett leverantörsspecifikt värde är bättre om det finns ett hållbarhetsbesked som styrker det; spara det dokumentet med rapporten.' },
        { t: `HVO factor ${fmtE(fH.wtw, 2)} differs from the market average ${fmtE(ref, 2)}`, d: 'The Swedish 2025 average is 0.35 kg CO₂e/l (Drivkraft Sverige 2026). A supplier-specific value is better if a sustainability declaration backs it; keep that document with the report.' });
    }
    if (fH && fH.validFrom && ctx.to && fH.validFrom.slice(0, 4) < ctx.to.slice(0, 4)) {
      add('warning', 'factorAge', { t: 'Emissionsfaktorn är från ett tidigare år', d: `Reduktionsplikten ändrar dieselns faktor årligen. Faktorn gäller från ${fH.validFrom}; uppdatera för perioden.` },
        { t: 'The emission factor is from an earlier year', d: `The reduction obligation changes diesel factors yearly. The factor is valid from ${fH.validFrom}; update it for the period.` });
    }
    // Price
    if (!extra.priceFromReceipts) {
      add('info', 'price', { t: `Ett snittpris (${fmt(ctx.prices.hvo, 2)} kr/l) för hela perioden`, d: 'Priset skrivs in för hand idag. Med kvitton räknas pris per drivmedel och period automatiskt.' },
        { t: `One average price (${fmtE(ctx.prices.hvo, 2)} SEK/l) for the whole period`, d: 'The price is typed in by hand today. With receipts the price per fuel and period is calculated automatically.' });
    }
    // Parity
    if (extra.parity) {
      const p = extra.parity;
      if (p.applicable) add(p.ok ? 'ok' : 'warning', 'parity', p.ok ? { t: 'Matchar nuvarande Excel-rapport', d: `${fmt(p.portal, 2)} kg CO₂e mot ${fmt(p.excel, 2)} kg i Excel-rapporten (avvikelse ${fmt(p.diff, 4)} kg).` } : { t: 'Avviker från Excel-rapporten', d: `${fmt(p.portal, 2)} kg mot ${fmt(p.excel, 2)} kg.` },
        p.ok ? { t: 'Matches the current Excel report', d: `${fmtE(p.portal, 2)} kg CO₂e vs ${fmtE(p.excel, 2)} kg in the Excel report (difference ${fmtE(p.diff, 4)} kg).` } : { t: 'Differs from the Excel report', d: `${fmtE(p.portal, 2)} kg vs ${fmtE(p.excel, 2)} kg.` });
    }
    const order = { critical: 0, serious: 1, warning: 2, info: 3, ok: 4 };
    out.sort((a, b) => order[a.sev] - order[b.sev]);
    return out;
  }

  // ---------- Canonical JSON (for hashing) ----------
  function canonical(v) {
    if (v === null || typeof v !== 'object') return JSON.stringify(typeof v === 'number' ? Math.round(v * 1e6) / 1e6 : v);
    if (Array.isArray(v)) return '[' + v.map(canonical).join(',') + ']';
    return '{' + Object.keys(v).sort().filter((k) => v[k] !== undefined).map((k) => JSON.stringify(k) + ':' + canonical(v[k])).join(',') + '}';
  }
  function linesFingerprint(lines) {
    return lines.map((l) => [l.date, l.status, l.resource, l.rtype, l.orderId, l.amount, l.km, l.qty, l.unit, l.article]);
  }

  const api = { DEFAULT_FUELS, toISODate, fromCompact, resourceTypesFrom, classifyTable, periodPresets, compute, inScope, checks, receiptStats, normalizeFuel, canonical, linesFingerprint, isEmptyRun, tonnesOf };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.Engine = api;
})(typeof window !== 'undefined' ? window : globalThis);
