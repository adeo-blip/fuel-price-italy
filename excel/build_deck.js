const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");

const DIR = process.env.FUEL_OUTDIR || __dirname;
const NAVY = "1E2761";
const NAVY_DEEP = "161C4C";
const ICE = "CADCFC";
const WHITE = "FFFFFF";
const MUTED = "8E9BD0";
const BLUE = "2A78D6";    // Benzina region
const ORANGE = "EB6834";  // Gasolio region
const NATBLUE = "7C93C9"; // Benzina national (muted)
const NATORANGE = "E8AD8C"; // Gasolio national (muted)
const GOOD = "34C79A";
const BAD = "E8654F";

function readJson(name) {
  let t = fs.readFileSync(path.join(DIR, name), "utf8");
  if (t.charCodeAt(0) === 0xFEFF) t = t.slice(1);
  return JSON.parse(t);
}

const regional = readJson("trends.json");
const nationalRaw = readJson("national.json");

const FLOOR_DATE = "2026-07-16";
const WINDOW_DAYS = 45;
const allNat = nationalRaw.filter(r => r.date >= FLOOR_DATE).map(r => r.date).sort();
const START_DATE = allNat[Math.max(0, allNat.length - WINDOW_DAYS)];
const LAST_DATE = allNat[allNat.length - 1];
const LAST_YEAR = LAST_DATE.slice(0, 4);
const TODAY_LABEL = new Date().toLocaleDateString("en-US", { year: "numeric", month: "long", day: "numeric" });

function byDate(rows) {
  const m = {};
  rows.filter(r => r.date >= START_DATE).forEach(r => { m[r.date] = r; });
  return m;
}
const nationalByDate = byDate(nationalRaw);
const dates = Object.keys(nationalByDate).sort();

const regions = Object.keys(regional).sort((a, b) => a.localeCompare(b, "it"));

function fmtEuro(v) { return "€" + v.toFixed(3); }
function fmtPct(v) { const s = v >= 0 ? "+" : ""; return s + v.toFixed(1) + "%"; }
function shortDate(d) {
  const months = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  const [y, m, day] = d.split("-").map(Number);
  return `${day} ${months[m - 1]}`;
}

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
const PW = 13.33, PH = 7.5;

function baseChartOpts(extra) {
  return Object.assign({
    showTitle: false,
    showLegend: true,
    legendPos: "t",
    legendColor: ICE,
    legendFontSize: 9,
    showValue: false,
    catAxisLabelColor: MUTED,
    catAxisLabelFontSize: 6.5,
    catAxisLabelRotate: 45,
    catAxisLineColor: "3A3F73",
    valAxisLabelColor: MUTED,
    valAxisLabelFontSize: 9,
    valGridLine: { color: "2B316B", size: 0.75 },
    catGridLine: { style: "none" },
    plotArea: { fill: { color: NAVY_DEEP } },
    chartArea: { fill: { color: NAVY_DEEP } },
  }, extra);
}

regions.forEach((region, idx) => {
  const regionByDateFull = byDate(regional[region]);
  const dates = Object.keys(nationalByDate).filter(d => regionByDateFull[d]).sort();
  const regionByDate = regionByDateFull;
  const cats = dates.map(shortDate);

  // A missing price (null or 0, e.g. no Gasolio sample in a small region on a given day) is a gap, not a EUR 0 price.
  const rnd = v => (v == null || !(v > 0) ? null : Math.round(v * 1000) / 1000);
  const gap = (a, b) => (a == null || b == null ? null : Math.round((a - b) * 1000) / 1000);
  const lastVal = arr => { for (let i = arr.length - 1; i >= 0; i--) if (arr[i] != null) return arr[i]; return null; };

  const benzinaRegion = dates.map(d => rnd(regionByDate[d].benzina));
  const gasolioRegion = dates.map(d => rnd(regionByDate[d].gasolio));
  const benzinaNat = dates.map(d => rnd(nationalByDate[d].benzina));
  const gasolioNat = dates.map(d => rnd(nationalByDate[d].gasolio));

  const benzinaGap = dates.map((d, i) => gap(benzinaRegion[i], benzinaNat[i]));
  const gasolioGap = dates.map((d, i) => gap(gasolioRegion[i], gasolioNat[i]));
  const bGapPos = benzinaGap.map(v => (v != null && v >= 0 ? v : null));
  const bGapNeg = benzinaGap.map(v => (v != null && v < 0 ? v : null));
  const gGapPos = gasolioGap.map(v => (v != null && v >= 0 ? v : null));
  const gGapNeg = gasolioGap.map(v => (v != null && v < 0 ? v : null));

  const lastB = lastVal(benzinaRegion);
  const lastBNat = benzinaNat[benzinaNat.length - 1];
  const lastG = lastVal(gasolioRegion);
  const lastGNat = gasolioNat[gasolioNat.length - 1];
  const pctB = ((lastB - lastBNat) / lastBNat) * 100;
  const pctG = ((lastG - lastGNat) / lastGNat) * 100;

  const slide = pres.addSlide();
  slide.background = { color: NAVY };

  slide.addShape("roundRect", {
    x: 0.45, y: 1.5, w: PW - 0.9, h: 5.55,
    rectRadius: 0.14,
    fill: { color: NAVY_DEEP },
    line: { type: "none" },
  });

  slide.addText("ITALY FUEL PRICE VS. NATIONAL AVERAGE", {
    x: 0.5, y: 0.4, w: 8, h: 0.3,
    fontFace: "Calibri", fontSize: 12, bold: true, color: GOOD, charSpacing: 2,
  });
  slide.addText(`REGION ${idx + 1} OF ${regions.length}`, {
    x: PW - 3.0, y: 0.4, w: 2.5, h: 0.3,
    align: "right",
    fontFace: "Calibri", fontSize: 12, bold: true, color: MUTED, charSpacing: 2,
  });

  const longName = region.length > 22;
  slide.addText(region, {
    x: 0.5, y: 0.68, w: 11.5, h: 0.62,
    fontFace: "Cambria", fontSize: longName ? 25 : 32, bold: true, color: WHITE,
  });
  slide.addText(`Daily average price vs. Italy national average • ${shortDate(dates[0])} – ${shortDate(dates[dates.length - 1])}, ${LAST_YEAR}`, {
    x: 0.5, y: 1.22, w: 11.5, h: 0.3,
    fontFace: "Calibri", fontSize: 12.5, color: ICE,
  });

  const kpis = [
    { label: "Benzina (this region)", value: fmtEuro(lastB) },
    { label: "Benzina vs national avg", value: fmtPct(pctB), color: pctB <= 0 ? GOOD : BAD },
    { label: "Gasolio (this region)", value: fmtEuro(lastG) },
    { label: "Gasolio vs national avg", value: fmtPct(pctG), color: pctG <= 0 ? GOOD : BAD },
  ];
  const kpiW = 2.75, kpiGap = 0.25, kpiX0 = 0.7, kpiY = 1.72;
  kpis.forEach((k, i) => {
    const x = kpiX0 + i * (kpiW + kpiGap);
    slide.addText(k.label, {
      x, y: kpiY, w: kpiW, h: 0.26,
      fontFace: "Calibri", fontSize: 11, color: MUTED,
    });
    slide.addText(k.value, {
      x, y: kpiY + 0.24, w: kpiW, h: 0.42,
      fontFace: "Calibri", fontSize: 20, bold: true, color: k.color || WHITE,
    });
  });

  // Column layout: left = Benzina (trend + gap), right = Gasolio (trend + gap)
  const colW = (PW - 1.4 - 0.3) / 2;
  const leftX = 0.7, rightX = leftX + colW + 0.3;
  const trendY = 2.65, trendH = 1.95;
  const gapY = trendY + trendH + 0.15, gapH = 1.45;

  slide.addText("BENZINA", { x: leftX, y: trendY - 0.24, w: colW, h: 0.22, fontFace: "Calibri", fontSize: 10, bold: true, color: BLUE, charSpacing: 1 });
  slide.addText("GASOLIO", { x: rightX, y: trendY - 0.24, w: colW, h: 0.22, fontFace: "Calibri", fontSize: 10, bold: true, color: ORANGE, charSpacing: 1 });

  slide.addChart(pres.ChartType.line, [
    { name: region, labels: cats, values: benzinaRegion },
    { name: "Italy avg", labels: cats, values: benzinaNat },
  ], baseChartOpts({
    x: leftX, y: trendY, w: colW, h: trendH,
    chartColors: [BLUE, NATBLUE],
    lineSize: 2.25,
    lineDataSymbol: "none",
    dashType: ["solid", "dash"],
    valAxisLabelFormatCode: "€0.00",
  }));

  slide.addChart(pres.ChartType.line, [
    { name: region, labels: cats, values: gasolioRegion },
    { name: "Italy avg", labels: cats, values: gasolioNat },
  ], baseChartOpts({
    x: rightX, y: trendY, w: colW, h: trendH,
    chartColors: [ORANGE, NATORANGE],
    lineSize: 2.25,
    lineDataSymbol: "none",
    dashType: ["solid", "dash"],
    valAxisLabelFormatCode: "€0.00",
  }));

  slide.addText("Gap vs national avg (€)", { x: leftX, y: gapY - 0.2, w: colW, h: 0.2, fontFace: "Calibri", fontSize: 9, color: MUTED });
  slide.addText("Gap vs national avg (€)", { x: rightX, y: gapY - 0.2, w: colW, h: 0.2, fontFace: "Calibri", fontSize: 9, color: MUTED });

  slide.addChart(pres.ChartType.bar, [
    { name: "Above national", labels: cats, values: bGapPos },
    { name: "Below national", labels: cats, values: bGapNeg },
  ], baseChartOpts({
    x: leftX, y: gapY, w: colW, h: gapH,
    chartColors: [BAD, GOOD],
    barGapWidthPct: 30,
    valAxisLabelFormatCode: "€0.00",
    showLegend: false,
    catAxisLabelFontSize: 6,
  }));

  slide.addChart(pres.ChartType.bar, [
    { name: "Above national", labels: cats, values: gGapPos },
    { name: "Below national", labels: cats, values: gGapNeg },
  ], baseChartOpts({
    x: rightX, y: gapY, w: colW, h: gapH,
    chartColors: [BAD, GOOD],
    barGapWidthPct: 30,
    valAxisLabelFormatCode: "€0.00",
    showLegend: false,
    catAxisLabelFontSize: 6,
  }));

  slide.addText(TODAY_LABEL, {
    x: 0.5, y: 7.12, w: 4, h: 0.28,
    fontFace: "Calibri", fontSize: 10, color: MUTED,
  });
  slide.addText("Adeodat Turatsinze  ·  Powered by Claude AI", {
    x: 6.8, y: 7.12, w: 6.0, h: 0.28,
    align: "right",
    fontFace: "Calibri", fontSize: 10, bold: true, color: ICE,
  });
});

const outFile = path.join(DIR, "Italy_Fuel_Price_Trends_20_Regions.pptx");
pres.writeFile({ fileName: outFile }).then(() => {
  console.log("written:", outFile);
});
