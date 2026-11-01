/* Spårbar report model — turns a calculation result into the plain object every output renders from. */
(function (root) {
  'use strict';
  function fuelShort(key, fuels, lang) {
    if (/^HVO/.test(key)) return 'HVO100';
    if (/^DIESEL/.test(key)) return lang === 'en' ? 'Diesel MK1' : 'Diesel MK1';
    if (key === 'B100') return 'B100';
    const f = fuels[key]; return f ? (f.label[lang] || f.label.sv) : key;
  }
  function buildModel(o) {
    const lang = o.lang, R = (sv, en) => (lang === 'en' ? en : sv);
    const loc = lang === 'en' ? 'en-GB' : 'sv-SE';
    const nf = (n, d = 0) => (n == null || !Number.isFinite(n) ? '–' : new Intl.NumberFormat(loc, { minimumFractionDigits: d, maximumFractionDigits: d }).format(n));
    const pc = (x, d = 0) => nf(x * 100, d) + (lang === 'en' ? '%' : ' %');
    const ctx = o.ctx, res = o.result, t = res.totals;
    const custNo = ctx.customer;
    const custName = (o.settings.customerNames && o.settings.customerNames[custNo]) || R(`Kund ${custNo}`, `Customer ${custNo}`);
    const issuer = (o.settings.issuer || '').trim() || R('Ert företagsnamn', 'Your company name');
    const monthName = (mk) => new Date(Date.UTC(+mk.slice(0, 4), +mk.slice(5, 7) - 1, 1)).toLocaleDateString(loc, { month: 'long', year: 'numeric', timeZone: 'UTC' });
    const usedFuels = [...new Set(res.rows.map((r) => r.fuel))];
    const rows = res.rows.map((r) => Object.assign({}, r, { fuelLabel: fuelShort(r.fuel, ctx.fuels, lang) }));
    const factors = usedFuels.map((k) => { const f = ctx.fuels[k]; return { key: k, label: f.label[lang] || f.label.sv, wtw: f.wtw, ttw: f.ttw, mj: f.mj, validFrom: f.validFrom || '', source: (f.source && (f.source[lang] || f.source.sv)) || '' }; });
    const showIntensityTkm = t.tkmCoverage >= 0.8 && t.tkm > 0;
    const hvoPct = ctx.fuelPlan.hvoPct;
    const typeCount = new Set(res.rows.map((r) => r.code)).size;
    const kgLabel = (kg) => (Math.abs(kg) >= 1000 ? `${nf(kg / 1000, 1)} ${R('ton', 'tonnes')} CO₂e` : `${nf(kg, 0)} kg CO₂e`);

    let narrative = R(
      `Under ${o.period.label.toLowerCase()} utförde vi ${nf(res.invoicedCount)} fakturerade transportuppdrag åt er med ${typeCount} fordonstyper. Klimatpåverkan uppgick till ${kgLabel(t.kgWtw)} räknat från källa till hjul (WTW), varav ${kgLabel(t.kgTtw)} direkta fossila utsläpp från fordonen (TTW).`,
      `During ${o.period.label} we carried out ${nf(res.invoicedCount)} invoiced transport assignments for you with ${typeCount} vehicle types. The climate impact was ${kgLabel(t.kgWtw)} well-to-wheel (WTW), of which ${kgLabel(t.kgTtw)} were direct fossil emissions from the vehicles (TTW).`);
    if (hvoPct >= 0.995) narrative += R(` Fordonen körde på HVO100.`, ` The vehicles ran on HVO100.`);
    else if (hvoPct > 0) narrative += R(` ${pc(hvoPct)} av drivmedelsvolymen var HVO100 och resten diesel.`, ` ${pc(hvoPct)} of the fuel volume was HVO100 and the rest diesel.`);
    if (o.scenarioKg != null && o.scenarioKg - t.kgWtw > 1 && hvoPct > 0) {
      narrative += R(` Med enbart ${fuelShort(ctx.fuelPlan.dieselFuel, ctx.fuels, lang).replace(/^D/, 'd')} hade samma transporter gett ungefär ${kgLabel(o.scenarioKg)}.`, ` With ${fuelShort(ctx.fuelPlan.dieselFuel, ctx.fuels, lang)} only, the same transports would have caused about ${kgLabel(o.scenarioKg)}.`);
    }
    narrative += R(` Totalt transporterades ${nf(t.tonnes)} ton gods, vilket ger ${nf(t.kgPerTonne, 2)} kg CO₂e per ton.`, ` In total ${nf(t.tonnes)} tonnes of goods were moved, which gives ${nf(t.kgPerTonne, 2)} kg CO₂e per tonne.`);

    const method = [
      R(`Beräkningen är kostnadsbaserad. För varje fordonstyp multipliceras fakturerat kundbelopp med fordonstypens drivmedelsandel av kostnaden. Drivmedelskostnaden divideras med genomsnittligt drivmedelspris för perioden (HVO100 ${nf(ctx.prices.hvo, 2)} kr/l${hvoPct < 1 ? `, diesel ${nf(ctx.prices.diesel, 2)} kr/l` : ''}) och ger volym i liter. Volymen multipliceras med drivmedlets emissionsfaktor från källa till hjul (WTW).`,
        `The calculation is cost-based. For each vehicle type the invoiced customer amount is multiplied by the vehicle type's fuel share of cost. The fuel cost is divided by the average fuel price for the period (HVO100 ${nf(ctx.prices.hvo, 2)} SEK/l${hvoPct < 1 ? `, diesel ${nf(ctx.prices.diesel, 2)} SEK/l` : ''}) to give litres. Litres are multiplied by the fuel's well-to-wheel (WTW) emission factor.`),
      R(`Resultatet är modellerade data enligt ISO 14083:2023 och redovisas både som WTW och TTW. Kunden kan använda siffrorna i sin klimatredovisning under GHG Protocol Scope 3, kategori 4 (uppströms transporter). Varje rad kan spåras till fakturerade orderrader, faktorernas källa och versionen av denna rapport.`,
        `The result is modelled data under ISO 14083:2023 and is reported as both WTW and TTW. The customer can use the figures in GHG Protocol Scope 3, Category 4 (upstream transportation). Every row traces back to invoiced order lines, the factor sources and the version of this report.`)
    ];

    const assumptions = [];
    if (o.receipts && o.receipts.mix) assumptions.push(R(`Drivmedelsmixen (${pc(hvoPct)} HVO100) är hämtad från ${nf(o.receipts.n)} tankkvitton för perioden.`, `The fuel mix (${pc(hvoPct)} HVO100) comes from ${nf(o.receipts.n)} fuel receipts for the period.`));
    else if (hvoPct >= 0.995) assumptions.push(R('Alla fordonstyper antas ha kört på HVO100. Underlag: tankkvitton och leverantörens hållbarhetsbesked.', 'All vehicle types are assumed to have run on HVO100. Evidence: fuel receipts and the supplier\'s sustainability declaration.'));
    else assumptions.push(R(`${pc(hvoPct)} av volymen antas vara HVO100 och ${pc(1 - hvoPct)} diesel.`, `${pc(hvoPct)} of volume is assumed to be HVO100 and ${pc(1 - hvoPct)} diesel.`));
    assumptions.push(o.receipts && o.receipts.price
      ? R('Drivmedelspriset är ett volymvägt snitt av periodens tankkvitton exklusive moms.', 'The fuel price is a volume-weighted average of the period\'s fuel receipts excluding VAT.')
      : R('Drivmedelspriset är ett angivet snittpris för perioden.', 'The fuel price is an entered average price for the period.'));
    assumptions.push(ctx.tonnageBasis === 'excel'
      ? R('Tonnage omfattar alla orderrader i perioden med vikt (ton och kg), oavsett status.', 'Tonnage covers all order lines in the period with a weight (tonnes and kg), regardless of status.')
      : R('Tonnage omfattar fakturerade transporter med vikt, utan bomkörningar.', 'Tonnage covers invoiced transports with a weight, excluding empty runs.'));
    for (const e of res.excluded) assumptions.push(R(`${e.name} (${nf(e.amount, 0)} kr) saknar drivmedelsandel och ingår inte i beräkningen.`, `${e.name} (${nf(e.amount, 0)} SEK) has no fuel cost share and is not included.`));
    if (!showIntensityTkm) assumptions.push(R(`Utsläpp per tonkilometer redovisas inte eftersom avstånd finns för ${pc(t.tkmCoverage)} av underlaget.`, `Emissions per tonne-km are not reported because distances exist for ${pc(t.tkmCoverage)} of the data.`));
    else assumptions.push(R(`Avstånd är enkel väg enligt ordern; ${pc(t.tkmCoverage)} av underlaget har avstånd.`, `Distances are one-way as per the order; ${pc(t.tkmCoverage)} of the data has a distance.`));

    return {
      lang, id: o.id, version: o.version || null, status: o.status || 'draft', issuer,
      title: R('Klimatrapport – transporter', 'Climate report – transport'),
      customer: { no: custNo, name: custName },
      period: { label: o.period.label, from: ctx.from, to: ctx.to, short: o.period.short, id: o.period.id },
      issued: o.issued, generatedAt: o.generatedAt,
      totals: t, showIntensityTkm, rows,
      monthly: res.monthly.map((x) => Object.assign({ label: monthName(x.month) }, x)),
      factors, prices: { hvo: ctx.prices.hvo, diesel: ctx.prices.diesel },
      plan: { hvoPct, hvoFuel: ctx.fuelPlan.hvoFuel, dieselFuel: ctx.fuelPlan.dieselFuel },
      tonnageBasis: ctx.tonnageBasis, narrative, method, assumptions,
      scenario: { dieselKg: o.scenarioKg }, showCost: !!o.settings.showCost,
      hashes: o.hashes || {}, checks: o.checks || [], lines: o.lines || []
    };
  }
  const api = { buildModel, fuelShort };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.Model = api;
})(typeof window !== 'undefined' ? window : globalThis);
