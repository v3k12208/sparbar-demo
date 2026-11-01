/* Spårbar exports — build files from a plain report model. No DOM. */
(function (root) {
  'use strict';
  const R = (m, sv, en) => (m.lang === 'en' ? en : sv);
  const loc = (m) => (m.lang === 'en' ? 'en-GB' : 'sv-SE');
  const nf = (m, n, d = 0) => (n == null || !Number.isFinite(n) ? '–' : new Intl.NumberFormat(loc(m), { minimumFractionDigits: d, maximumFractionDigits: d }).format(n));
  const pdfSafe = (s) => String(s == null ? '' : s)
    .replace(/[−–—]/g, '-').replace(/[   ]/g, ' ').replace(/₂/g, '2')
    .replace(/…/g, '...').replace(/≈/g, '~').replace(/→/g, '->').replace(/[‘’]/g, "'").replace(/[“”]/g, '"')
    .replace(/✓/g, 'OK').replace(/[^\x00-\xff]/g, '');

  function statusLabel(m) {
    if (m.status === 'locked') return R(m, `Låst version ${m.version}`, `Locked version ${m.version}`);
    if (m.status === 'changed') return R(m, `Utkast (ändrad sedan v${m.version})`, `Draft (changed since v${m.version})`);
    return R(m, 'Utkast', 'Draft');
  }

  // ---------- Excel (same layout as the current Rapport sheet, plus appendices) ----------
  function buildXlsx(m, XLSX) {
    const wb = XLSX.utils.book_new();
    const H = [R(m, 'Halvår', 'Period'), R(m, 'Resurstyp', 'Vehicle type'), R(m, 'Drivmedel', 'Fuel'), R(m, 'Volym [liter]', 'Volume [litres]'),
      R(m, 'Emissionsvärde \n(kg CO2ekv/l)', 'Emission factor \n(kg CO2e/l)'), R(m, 'Klimatpåverkan WTW (kg CO2ekv) ', 'Climate impact WTW (kg CO2e)'),
      R(m, 'Kundbelopp', 'Customer amount'), R(m, 'Andel drivmedel %', 'Fuel cost share'), R(m, 'Andel kostnad drivmedel', 'Fuel cost'), R(m, 'Pris kr/l', 'Price SEK/l')];
    const aoa = [[R(m, 'Kund', 'Customer'), `${m.customer.no} ${m.customer.name}`, null, null, null, null, R(m, 'Totalt WTW (KG)', 'Total WTW (kg)'), null, m.totals.kgWtw], [], H];
    m.rows.forEach((r) => aoa.push([m.period.short, r.name, r.fuelLabel, r.liters, r.ef, r.kgWtw, r.amount, r.share, r.fuelCost, r.price]));
    const last = 3 + m.rows.length; // last data row (1-based)
    aoa.push([]); aoa.push([]);
    aoa.push([R(m, 'Totalt antal ton transporterat gods', 'Total tonnes of goods transported'), null, null, Math.round(m.totals.tonnes * 100) / 100]);
    aoa.push([R(m, 'Rapport-ID', 'Report ID'), m.id, null, statusLabel(m)]);
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    for (let i = 0; i < m.rows.length; i++) {
      const r = 4 + i;
      ws['D' + r].f = `I${r}/J${r}`; ws['F' + r].f = `D${r}*E${r}`; ws['I' + r].f = `G${r}*H${r}`;
      ['D', 'F', 'G', 'I'].forEach((c) => { ws[c + r].z = '#,##0.00'; });
      ws['H' + r].z = '0.0%'; ws['J' + r].z = '0.00'; ws['E' + r].z = '0.000';
    }
    if (m.rows.length) { ws.I1.f = `SUM(F4:F${last})`; ws.I1.z = '#,##0.00'; }
    ws['!cols'] = [{ wch: 10 }, { wch: 26 }, { wch: 12 }, { wch: 14 }, { wch: 14 }, { wch: 18 }, { wch: 14 }, { wch: 12 }, { wch: 14 }, { wch: 10 }];
    XLSX.utils.book_append_sheet(wb, ws, R(m, 'Rapport', 'Report'));

    const mon = [[R(m, 'Månad', 'Month'), R(m, 'Volym [liter]', 'Volume [litres]'), 'kg CO2e WTW']].concat(m.monthly.map((x) => [x.month, x.liters, x.kgWtw]));
    mon.push([R(m, 'Totalt', 'Total'), m.totals.liters, m.totals.kgWtw]);
    const wsM = XLSX.utils.aoa_to_sheet(mon); wsM['!cols'] = [{ wch: 10 }, { wch: 14 }, { wch: 14 }];
    XLSX.utils.book_append_sheet(wb, wsM, R(m, 'Per månad', 'By month'));

    const sum = [
      [R(m, 'Sammanfattning', 'Summary')], [],
      [R(m, 'Utfärdare', 'Issuer'), m.issuer], [R(m, 'Kund', 'Customer'), `${m.customer.no} ${m.customer.name}`],
      [R(m, 'Period', 'Period'), `${m.period.label} (${m.period.from} – ${m.period.to})`],
      ['kg CO2e WTW', m.totals.kgWtw], ['kg CO2e TTW', m.totals.kgTtw], ['kg CO2e WTT', m.totals.kgWtt],
      [R(m, 'Liter drivmedel', 'Litres of fuel'), m.totals.liters], ['MWh', m.totals.mwh], [R(m, 'Förnybar andel (energi)', 'Renewable share (energy)'), m.totals.renewableShare],
      [R(m, 'Ton gods', 'Tonnes of goods'), m.totals.tonnes], ['tkm', m.totals.tkm], ['kg CO2e/t', m.totals.kgPerTonne],
      ['g CO2e/tkm', m.showIntensityTkm ? m.totals.gPerTkm : R(m, 'redovisas ej (låg täckning)', 'not reported (low coverage)')],
      [R(m, 'Täckning avstånd (andel belopp)', 'Distance coverage (share of amount)'), m.totals.tkmCoverage], [],
      [R(m, 'Metod', 'Method')], ...m.method.map((p) => [p]), [],
      [R(m, 'Antaganden', 'Assumptions')], ...m.assumptions.map((p) => ['• ' + p]), [],
      [R(m, 'Emissionsfaktorer', 'Emission factors'), 'WTW kg/l', 'TTW kg/l', 'MJ/l', R(m, 'Gäller från', 'Valid from'), R(m, 'Källa', 'Source')],
      ...m.factors.map((f) => [f.label, f.wtw, f.ttw, f.mj, f.validFrom, f.source]),
      [R(m, 'Pris HVO-klass kr/l', 'Price HVO class SEK/l'), m.prices.hvo], [R(m, 'Pris dieselklass kr/l', 'Price diesel class SEK/l'), m.prices.diesel], [],
      [R(m, 'Verifiering', 'Verification')], [R(m, 'Underlagets SHA-256', 'Input SHA-256'), m.hashes.input || ''], [R(m, 'Rapportens SHA-256', 'Report SHA-256'), m.hashes.report || ''],
      [R(m, 'Genererad', 'Generated'), m.generatedAt]
    ];
    const wsS = XLSX.utils.aoa_to_sheet(sum); wsS['!cols'] = [{ wch: 34 }, { wch: 22 }, { wch: 12 }, { wch: 10 }, { wch: 12 }, { wch: 70 }];
    XLSX.utils.book_append_sheet(wb, wsS, R(m, 'Metod och faktorer', 'Method and factors'));

    if (m.checks && m.checks.length) {
      const c = [[R(m, 'Nivå', 'Level'), R(m, 'Kontroll', 'Check'), R(m, 'Detalj', 'Detail')]].concat(m.checks.map((x) => [x.sev, x.t, x.d]));
      const wsC = XLSX.utils.aoa_to_sheet(c); wsC['!cols'] = [{ wch: 10 }, { wch: 50 }, { wch: 110 }];
      XLSX.utils.book_append_sheet(wb, wsC, R(m, 'Kontroller', 'Checks'));
    }
    if (m.lines && m.lines.length) {
      const u = [[R(m, 'Datum', 'Date'), 'Status', R(m, 'Resurs', 'Resource'), R(m, 'Resurstyp', 'Type code'), R(m, 'Resurstypsbenämning', 'Type'), 'OrderId', R(m, 'Artikel', 'Article'), R(m, 'Avstånd km', 'Distance km'), R(m, 'Kvantitet', 'Quantity'), R(m, 'Enhet', 'Unit'), R(m, 'Kundbelopp', 'Customer amount')]].concat(m.lines);
      const wsU = XLSX.utils.aoa_to_sheet(u); wsU['!cols'] = [{ wch: 11 }, { wch: 7 }, { wch: 9 }, { wch: 9 }, { wch: 24 }, { wch: 10 }, { wch: 24 }, { wch: 10 }, { wch: 10 }, { wch: 7 }, { wch: 12 }];
      XLSX.utils.book_append_sheet(wb, wsU, R(m, 'Underlag', 'Source lines'));
    }
    return XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
  }

  // ---------- PDF ----------
  function buildPdf(m, jsPDF) {
    const doc = new jsPDF({ unit: 'mm', format: 'a4' });
    const W = 210, M = 16, CW = W - 2 * M;
    const P = pdfSafe;
    const ink = [17, 17, 17], mut = [95, 95, 95], acc = [240, 0, 55], line = [224, 224, 224], soft = [244, 244, 244];
    const setC = (c) => doc.setTextColor(c[0], c[1], c[2]);
    let y = 18;
    doc.setFont('helvetica', 'bold'); doc.setFontSize(10); setC(ink); doc.text(P(m.issuer), M, y);
    doc.setFont('helvetica', 'normal'); doc.setFontSize(8); setC(m.status === 'locked' ? [11, 110, 11] : [154, 107, 0]);
    doc.text(P(statusLabel(m).toUpperCase()), W - M, y, { align: 'right' });
    y += 9; doc.setFont('helvetica', 'bold'); doc.setFontSize(20); setC(ink); doc.text(P(m.title), M, y);
    y += 4; doc.setDrawColor(acc[0], acc[1], acc[2]); doc.setLineWidth(0.9); doc.line(M, y, W - M, y);
    y += 6;
    const meta = [
      [R(m, 'KUND', 'CUSTOMER'), `${m.customer.name} (${R(m, 'kundnr', 'no.')} ${m.customer.no})`],
      [R(m, 'PERIOD', 'PERIOD'), `${m.period.label} · ${m.period.from} – ${m.period.to}`],
      [R(m, 'RAPPORT-ID', 'REPORT ID'), `${m.id}${m.version ? ' · v' + m.version : ''} · ${R(m, 'utfärdad', 'issued')} ${m.issued}`]
    ];
    const colW = CW / 3;
    meta.forEach((x, i) => {
      doc.setFont('helvetica', 'bold'); doc.setFontSize(6.8); setC(mut); doc.text(P(x[0]), M + i * colW, y);
      doc.setFont('helvetica', 'normal'); doc.setFontSize(8.6); setC(ink);
      doc.text(doc.splitTextToSize(P(x[1]), colW - 4), M + i * colW, y + 4.2);
    });
    y += 13; doc.setDrawColor(line[0], line[1], line[2]); doc.setLineWidth(0.25); doc.line(M, y, W - M, y);
    // KPI boxes
    y += 5; const bw = CW / 4, bh = 19;
    const t = m.totals;
    const kpis = [
      [R(m, 'KLIMATPÅVERKAN WTW', 'CLIMATE IMPACT WTW'), t.kgWtw >= 1000 ? nf(m, t.kgWtw / 1000, 1) : nf(m, t.kgWtw, 0), t.kgWtw >= 1000 ? R(m, 'ton CO2e', 'tonnes CO2e') : 'kg CO2e'],
      [R(m, 'TRANSPORTERAT GODS', 'GOODS TRANSPORTED'), nf(m, t.tonnes, 0), R(m, 'ton', 'tonnes')],
      [R(m, 'INTENSITET', 'INTENSITY'), nf(m, t.kgPerTonne, 2), R(m, 'kg CO2e per ton', 'kg CO2e per tonne')],
      [R(m, 'FÖRNYBART DRIVMEDEL', 'RENEWABLE FUEL'), nf(m, t.renewableShare * 100, 0) + (m.lang === 'en' ? '%' : ' %'), R(m, 'av energin', 'of energy')]
    ];
    doc.setFillColor(soft[0], soft[1], soft[2]); doc.rect(M, y, bw, bh, 'F');
    doc.setDrawColor(line[0], line[1], line[2]); doc.rect(M, y, CW, bh);
    kpis.forEach((k, i) => {
      const x = M + i * bw;
      if (i) doc.line(x, y, x, y + bh);
      doc.setFont('helvetica', 'bold'); doc.setFontSize(6.4); setC(mut); doc.text(P(k[0]), x + 3.5, y + 5);
      doc.setFont('helvetica', 'bold'); doc.setFontSize(15); setC(ink); doc.text(P(k[1]), x + 3.5, y + 12.2);
      doc.setFont('helvetica', 'normal'); doc.setFontSize(7); setC(mut); doc.text(P(k[2]), x + 3.5, y + 16.4);
    });
    y += bh + 7;
    doc.setFont('helvetica', 'normal'); doc.setFontSize(9); setC(ink);
    const narr = doc.splitTextToSize(P(m.narrative), CW); doc.text(narr, M, y); y += narr.length * 4.1 + 4;

    const head = [R(m, 'Resurstyp', 'Vehicle type'), R(m, 'Drivmedel', 'Fuel'), R(m, 'Volym (l)', 'Volume (l)'), R(m, 'Faktor kg/l', 'Factor kg/l'), 'kg CO2e WTW'];
    if (m.showCost) head.push(R(m, 'Kundbelopp (kr)', 'Amount (SEK)'), R(m, 'Andel', 'Share'), R(m, 'Pris kr/l', 'SEK/l'));
    const body = m.rows.map((r) => {
      const a = [r.name, r.fuelLabel, nf(m, r.liters, 0), nf(m, r.ef, 3), nf(m, r.kgWtw, 1)];
      if (m.showCost) a.push(nf(m, r.amount, 0), nf(m, r.share * 100, 1) + (m.lang === 'en' ? '%' : ' %'), nf(m, r.price, 2));
      return a.map(P);
    });
    const foot = [R(m, 'Totalt', 'Total'), '', nf(m, t.liters, 0), '', nf(m, t.kgWtw, 1)];
    if (m.showCost) foot.push(nf(m, t.amount, 0), '', '');
    const right = {}; for (let i = 2; i < head.length; i++) right[i] = { halign: 'right' };
    const tableStyle = {
      theme: 'plain', margin: { left: M, right: M }, styles: { font: 'helvetica', fontSize: 7.8, cellPadding: { top: 1.6, bottom: 1.6, left: 1.8, right: 1.8 }, textColor: ink, lineColor: line, lineWidth: { bottom: 0.2 } },
      headStyles: { fontStyle: 'bold', fontSize: 6.8, textColor: mut, lineColor: ink, lineWidth: { bottom: 0.4 } },
      footStyles: { fontStyle: 'bold', fillColor: soft, lineColor: ink, lineWidth: { top: 0.4 } }, columnStyles: right,
    };
    const withCols = (cols, extra) => Object.assign({}, tableStyle, { columnStyles: cols, didParseCell: (d) => { if (d.section !== 'body' && cols[d.column.index] && cols[d.column.index].halign === 'right') d.cell.styles.halign = 'right'; } }, extra);
    doc.setFont('helvetica', 'bold'); doc.setFontSize(9); setC(ink); doc.text(P(R(m, 'KLIMATPÅVERKAN PER FORDONSTYP', 'CLIMATE IMPACT BY VEHICLE TYPE')), M, y); y += 2;
    doc.autoTable(withCols(right, { startY: y, head: [head.map(P)], body, foot: [foot.map(P)] }));
    y = doc.lastAutoTable.finalY + 8;
    const need = (h) => { if (y + h > 272) { doc.addPage(); y = 20; } };
    need(14 + m.monthly.length * 6.5);
    doc.setFont('helvetica', 'bold'); doc.setFontSize(9); setC(ink); doc.text(P(R(m, 'PER MÅNAD', 'BY MONTH')), M, y); y += 2;
    doc.autoTable(withCols({ 1: { halign: 'right' }, 2: { halign: 'right' } }, {
      startY: y, tableWidth: CW * 0.55,
      head: [[R(m, 'Månad', 'Month'), R(m, 'Volym (l)', 'Volume (l)'), 'kg CO2e WTW'].map(P)],
      body: m.monthly.map((x) => [x.label, nf(m, x.liters, 0), nf(m, x.kgWtw, 1)].map(P)),
      foot: [[R(m, 'Totalt', 'Total'), nf(m, t.liters, 0), nf(m, t.kgWtw, 1)].map(P)]
    }));
    y = doc.lastAutoTable.finalY + 8;
    const para = (title, paras, bullet) => {
      need(18);
      doc.setFont('helvetica', 'bold'); doc.setFontSize(9); setC(ink); doc.text(P(title), M, y); y += 5;
      doc.setFont('helvetica', 'normal'); doc.setFontSize(8.4); setC(ink);
      for (const p of paras) {
        const lines = doc.splitTextToSize(P((bullet ? '- ' : '') + p), CW - (bullet ? 2 : 0));
        need(lines.length * 3.8 + 2);
        doc.text(lines, M + (bullet ? 1 : 0), y); y += lines.length * 3.8 + 1.8;
      }
      y += 3;
    };
    para(R(m, 'METOD', 'METHOD'), m.method);
    need(30);
    doc.setFont('helvetica', 'bold'); doc.setFontSize(9); setC(ink); doc.text(P(R(m, 'EMISSIONSFAKTORER', 'EMISSION FACTORS')), M, y); y += 2;
    doc.autoTable(withCols({ 1: { halign: 'right' }, 2: { halign: 'right' }, 3: { halign: 'right' }, 4: { cellWidth: 70 } }, {
      startY: y,
      head: [[R(m, 'Drivmedel', 'Fuel'), 'WTW kg/l', 'TTW kg/l', R(m, 'Gäller från', 'Valid from'), R(m, 'Källa', 'Source')].map(P)],
      body: m.factors.map((f) => [f.label, nf(m, f.wtw, 3), nf(m, f.ttw, 3), f.validFrom, f.source].map(P))
    }));
    y = doc.lastAutoTable.finalY + 8;
    para(R(m, 'ANTAGANDEN OCH DATAKVALITET', 'ASSUMPTIONS AND DATA QUALITY'), m.assumptions, true);

    const pages = doc.getNumberOfPages();
    for (let i = 1; i <= pages; i++) {
      doc.setPage(i);
      doc.setDrawColor(line[0], line[1], line[2]); doc.setLineWidth(0.25); doc.line(M, 284, W - M, 284);
      doc.setFont('helvetica', 'normal'); doc.setFontSize(6.6); setC(mut);
      doc.text(P(R(m, 'Verifieringskod (SHA-256): ', 'Verification code (SHA-256): ') + (m.hashes.report || R(m, 'beräknas vid låsning', 'set when locked'))), M, 288);
      doc.text(P(`${m.id} · ${R(m, 'sida', 'page')} ${i}/${pages}`), W - M, 288, { align: 'right' });
      setC(acc); doc.text(P(R(m, 'Genererad med Spårbar · ', 'Generated with Spårbar · ') + m.generatedAt.replace('T', ' ').slice(0, 16)), M, 291.5);
    }
    return doc.output('arraybuffer');
  }

  // ---------- CSV (customer Scope 3 cat. 4 import) ----------
  function buildCsv(m) {
    const sv = m.lang !== 'en', sep = sv ? ';' : ',';
    const n = (x, d) => (x == null || !Number.isFinite(x) ? '' : (sv ? x.toFixed(d).replace('.', ',') : x.toFixed(d)));
    const q = (s) => { s = String(s == null ? '' : s); return (s.includes(sep) || /["\r\n]/.test(s)) ? '"' + s.replace(/"/g, '""') + '"' : s; };
    const head = ['period_from', 'period_to', 'report_id', 'version', 'issuer', 'customer_no', 'customer_name', 'ghg_scope', 'ghg_category', 'vehicle_type_code', 'vehicle_type', 'fuel', 'litres', 'ef_wtw_kg_per_l', 'ef_ttw_kg_per_l', 'kg_co2e_wtw', 'kg_co2e_ttw', 'kg_co2e_wtt', 'tonnes', 'tkm', 'method', 'data_type', 'factor_source'];
    const out = [head.join(sep)];
    for (const r of m.rows) {
      const f = m.factors.find((x) => x.key === r.fuel) || {};
      out.push([m.period.from, m.period.to, m.id, m.version || 'draft', m.issuer, m.customer.no, m.customer.name, '3', '4 Upstream transportation and distribution', r.code, r.name, r.fuelLabel,
        n(r.liters, 2), n(r.ef, 4), n(r.eft, 4), n(r.kgWtw, 2), n(r.kgTtw, 2), n(r.kgWtw - r.kgTtw, 2), n(r.tonnes, 2), n(r.tkm, 1), 'cost-based fuel model', 'modelled (ISO 14083)', f.source || ''].map(q).join(sep));
    }
    return '﻿' + out.join('\r\n') + '\r\n';
  }

  // ---------- JSON (ISO 14083-structured declaration) ----------
  function buildJson(m) {
    const t = m.totals;
    const o = {
      schema: 'sparbar.emissions-declaration/0.1', methodology: 'EN ISO 14083:2023 (well-to-wheel), modelled data via cost-based fuel model',
      report_id: m.id, version: m.version || null, status: m.status, issued: m.issued, generated_at: m.generatedAt,
      issuer: m.issuer, customer: { number: m.customer.no, name: m.customer.name }, period: { from: m.period.from, to: m.period.to, label: m.period.label },
      reporting_use: 'Customer GHG Protocol Scope 3, Category 4 (upstream transportation and distribution)',
      totals: {
        kg_co2e_wtw: t.kgWtw, kg_co2e_ttw: t.kgTtw, kg_co2e_wtt: t.kgWtt, litres: t.liters, energy_mwh: t.mwh, renewable_energy_share: t.renewableShare,
        tonnes: t.tonnes, tonne_km: t.tkm, kg_co2e_per_tonne: t.kgPerTonne, g_co2e_per_tkm: m.showIntensityTkm ? t.gPerTkm : null, distance_coverage_share_of_amount: t.tkmCoverage
      },
      by_vehicle_type: m.rows.map((r) => ({ code: r.code, name: r.name, fuel: r.fuel, litres: r.liters, ef_wtw_kg_per_l: r.ef, ef_ttw_kg_per_l: r.eft, kg_co2e_wtw: r.kgWtw, kg_co2e_ttw: r.kgTtw, tonnes: r.tonnes, tonne_km: r.tkm, invoiced_rows: r.n })),
      by_month: m.monthly.map((x) => ({ month: x.month, litres: x.liters, kg_co2e_wtw: x.kgWtw })),
      emission_factors: m.factors.map((f) => ({ key: f.key, label: f.label, wtw_kg_per_l: f.wtw, ttw_kg_per_l: f.ttw, mj_per_l: f.mj, valid_from: f.validFrom, source: f.source })),
      data_quality: { data_type: 'modelled', primary_data_share: 0, method: 'invoiced amount × fuel cost share ÷ fuel price × WTW factor', tonnage_basis: m.tonnageBasis },
      assumptions: m.assumptions,
      verification: { input_sha256: m.hashes.input || null, report_sha256: m.hashes.report || null }
    };
    return JSON.stringify(o, null, 2);
  }

  const api = { buildXlsx, buildPdf, buildCsv, buildJson, pdfSafe, statusLabel };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.Exports = api;
})(typeof window !== 'undefined' ? window : globalThis);
