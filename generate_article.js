// Renders the day's Fuel Dashboard regional report — a ~10-page article covering the
// weekly trend of all 4 main fuels (Benzina, Gasolio, GPL, Metano) for that day's
// rotating region — into a .docx. Title/dek/section prose/pull-quote/caption are
// composed fresh each day by the scheduled agent and handed in via
// data/article_content.json; the numbers (current prices, weekly region-vs-national
// series, station counts) come from data/region_of_day.json.
'use strict';
const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, AlignmentType,
  Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle,
  PageBreak, VerticalAlign,
} = require('docx');

const BASE_DIR = __dirname;
const REGION_JSON = path.join(BASE_DIR, 'data', 'region_of_day.json');
const CONTENT_JSON = path.join(BASE_DIR, 'data', 'article_content.json');
const OUT_DIR = path.join(BASE_DIR, 'slides');
const SITE_URL = 'https://adeo-blip.github.io/fuel-price-italy/';

const COLOR = {
  ink: '14120F',
  muted: '5B564C',
  accent: 'B9790A',
  accentSoft: 'F7EEDD',
  rule: 'E6E2D8',
  headerFill: 'F0EDE6',
};
const FUEL_COLOR = { Benzina: '2A78D6', Gasolio: '1BAF7A', GPL: 'C98500', Metano: '1F9C1F' };
const FUEL_GLOSS = { Benzina: 'petrol', Gasolio: 'diesel', GPL: 'LPG', Metano: 'CNG' };
const FUELS = ['Benzina', 'Gasolio', 'GPL', 'Metano'];

function fmtPrice(v) { return v == null ? '—' : `€${v.toFixed(3)}`; }
function fmtDelta(v, dec = 3) {
  if (v == null) return '—';
  if (Math.abs(v) < 5 * 10 ** -(dec + 1)) return `→ ${v.toFixed(dec)}`;
  const sign = v >= 0 ? '+' : '';
  const arrow = v < 0 ? '↓' : '↑';
  return `${arrow} ${sign}${v.toFixed(dec)}`;
}
function fmtPct(v) {
  if (v == null) return '—';
  const sign = v >= 0 ? '+' : '';
  return `${sign}${v.toFixed(1)}%`;
}
function fmtDateShort(iso) {
  const [y, m, d] = iso.split('-').map(Number);
  const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  return `${months[m - 1]} ${d}`;
}

function para(text, opts = {}) {
  return new Paragraph({
    spacing: { after: opts.after ?? 200 },
    alignment: opts.align,
    border: opts.border,
    children: [new TextRun({ text, size: opts.size ?? 22, bold: !!opts.bold, italics: !!opts.italics, color: opts.color ?? COLOR.ink, font: opts.font ?? 'Calibri' })],
  });
}
function heading(text, opts = {}) {
  return new Paragraph({
    spacing: { before: opts.before ?? 0, after: opts.after ?? 180 },
    children: [new TextRun({ text, bold: true, size: opts.size ?? 34, color: opts.color ?? COLOR.ink, font: 'Cambria' })],
  });
}
function pageBreakPara() {
  return new Paragraph({ children: [new PageBreak()] });
}
function eyebrow(text) {
  return new Paragraph({ spacing: { after: 80 }, children: [new TextRun({ text, bold: true, size: 18, color: COLOR.accent, font: 'Calibri' })] });
}

function statCell(label, value, sub, widthDxa) {
  return new TableCell({
    width: { size: widthDxa, type: WidthType.DXA },
    shading: { type: ShadingType.CLEAR, fill: COLOR.accentSoft },
    verticalAlign: VerticalAlign.TOP,
    margins: { top: 160, bottom: 160, left: 180, right: 180 },
    children: [
      new Paragraph({ children: [new TextRun({ text: label.toUpperCase(), bold: true, size: 15, color: COLOR.muted, font: 'Calibri' })] }),
      new Paragraph({ spacing: { before: 60 }, children: [new TextRun({ text: value, bold: true, size: 28, color: COLOR.ink, font: 'Calibri' })] }),
      new Paragraph({ spacing: { before: 40 }, children: [new TextRun({ text: sub || ' ', size: 17, color: COLOR.muted, font: 'Calibri' })] }),
    ],
  });
}

function noBorders() {
  const none = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
  return { top: none, bottom: none, left: none, right: none, insideHorizontal: none, insideVertical: none };
}

function headerCell(text, widthDxa) {
  return new TableCell({
    width: { size: widthDxa, type: WidthType.DXA },
    shading: { type: ShadingType.CLEAR, fill: COLOR.headerFill },
    children: [new Paragraph({ children: [new TextRun({ text, bold: true, size: 18, font: 'Calibri' })] })],
  });
}
function bodyCell(text, widthDxa, color) {
  return new TableCell({
    width: { size: widthDxa, type: WidthType.DXA },
    children: [new Paragraph({ children: [new TextRun({ text, size: 18, font: 'Calibri', color: color || COLOR.ink })] })],
  });
}

function trendSummary(fuel, weekly) {
  const vals = weekly.filter(w => w[fuel] != null);
  if (vals.length < 2) return null;
  const first = vals[0], last = vals[vals.length - 1];
  const delta = last[fuel] - first[fuel];
  const pct = (delta / first[fuel]) * 100;
  return { first, last, delta, pct };
}

function weeklyTable(fuel, regionWeekly, nationalWeekly) {
  const natByWeek = Object.fromEntries(nationalWeekly.map(w => [w.week, w]));
  const rows = [new TableRow({
    children: [
      headerCell('Week of', 1800),
      headerCell(`${fuel} — Region`, 2400),
      headerCell(`${fuel} — Italy`, 2400),
      headerCell('Region vs. Italy', 2400),
    ],
  })];
  regionWeekly.forEach(w => {
    const nat = natByWeek[w.week];
    const regionVal = w[fuel], natVal = nat ? nat[fuel] : null;
    const diff = (regionVal != null && natVal != null) ? regionVal - natVal : null;
    rows.push(new TableRow({
      children: [
        bodyCell(fmtDateShort(w.week_start), 1800),
        bodyCell(fmtPrice(regionVal), 2400, FUEL_COLOR[fuel]),
        bodyCell(fmtPrice(natVal), 2400),
        bodyCell(diff == null ? '—' : `${fmtDelta(diff)} €/L`, 2400),
      ],
    }));
  });
  return new Table({ width: { size: 9000, type: WidthType.DXA }, columnWidths: [1800, 2400, 2400, 2400], rows });
}

function fuelSection(section, region) {
  const fuel = section.fuel;
  const color = FUEL_COLOR[fuel] || COLOR.accent;
  const regionTrend = trendSummary(fuel, region.weekly);
  const nationalTrend = trendSummary(fuel, region.national_weekly);
  const latest = region.latest[fuel] || {};
  const previous = region.previous ? region.previous[fuel] : null;
  const dayDelta = previous && latest.self != null ? latest.self - previous.self : null;

  const children = [
    eyebrow(`FUEL SPOTLIGHT · ${fuel.toUpperCase()} (${FUEL_GLOSS[fuel]})`),
    heading(section.heading, { color }),
  ];

  const rows = [new TableRow({ children: [
    statCell('Latest self-service', fmtPrice(latest.self), region.latest.date ? fmtDateShort(region.latest.date) : '', 2250),
    statCell('Vs. yesterday', fmtDelta(dayDelta) + ' €/L', 'day-over-day', 2250),
    statCell(`Since ${regionTrend ? fmtDateShort(regionTrend.first.week_start) : '—'}`, regionTrend ? fmtPct(regionTrend.pct) : '—', regionTrend ? `${region.region}, weekly avg` : '', 2250),
    statCell('Italy, same span', nationalTrend ? fmtPct(nationalTrend.pct) : '—', 'national weekly avg', 2250),
  ] })];
  children.push(new Table({ width: { size: 9000, type: WidthType.DXA }, columnWidths: [2250, 2250, 2250, 2250], borders: noBorders(), rows }));
  children.push(para('', { after: 200 }));

  section.paragraphs.forEach(p => children.push(para(p)));

  children.push(para(`Weekly self-service average, ${region.region} vs. Italy`, { bold: true, size: 18, color: COLOR.muted, after: 120 }));
  children.push(weeklyTable(fuel, region.weekly, region.national_weekly));

  return children;
}

function main() {
  const region = JSON.parse(fs.readFileSync(REGION_JSON, 'utf-8'));
  const content = JSON.parse(fs.readFileSync(CONTENT_JSON, 'utf-8'));

  const latest = region.latest;
  const coverStats = [new TableRow({ children: FUELS.map(f =>
    statCell(f, fmtPrice(latest[f] ? latest[f].self : null), FUEL_GLOSS[f], 2250)
  ) })];

  const sections = [];

  // --- Cover ---
  sections.push(
    eyebrow('FUEL DASHBOARD · REGIONE SPOTLIGHT'),
    new Paragraph({ spacing: { before: 400, after: 100 }, children: [new TextRun({ text: content.title, bold: true, size: 60, color: COLOR.ink, font: 'Cambria' })] }),
    para(content.dek, { italics: true, size: 24, color: COLOR.muted, after: 320 }),
    new Table({ width: { size: 9000, type: WidthType.DXA }, columnWidths: [2250, 2250, 2250, 2250], borders: noBorders(), rows: coverStats }),
    para('', { after: 200 }),
    para(`Registered stations: ${latest.stations.registered.toLocaleString('en-US')} · Reporting today: ${latest.stations.reporting.toLocaleString('en-US')} (${Math.round(100 * latest.stations.reporting / latest.stations.registered)}%) · ${fmtDateShort(latest.date)}`, { size: 18, color: COLOR.muted }),
    pageBreakPara(),
  );

  // --- Executive summary ---
  sections.push(
    eyebrow('EXECUTIVE SUMMARY'),
    heading(`${region.region} at a glance`),
    ...content.executive_summary.map(p => para(p)),
    new Paragraph({
      spacing: { before: 200, after: 200 }, alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: content.pull_quote, italics: true, bold: true, size: 30, color: COLOR.accent, font: 'Cambria' })],
    }),
    pageBreakPara(),
  );

  // --- One page per fuel ---
  content.fuel_sections.forEach((section, i) => {
    sections.push(...fuelSection(section, region));
    if (i < content.fuel_sections.length - 1) sections.push(pageBreakPara());
  });
  sections.push(pageBreakPara());

  // --- Station network ---
  sections.push(
    eyebrow('STATION NETWORK'),
    heading(content.network_section.heading),
    ...content.network_section.paragraphs.map(p => para(p)),
    pageBreakPara(),
  );

  // --- Methodology ---
  sections.push(
    eyebrow('HOW THIS REPORT IS BUILT'),
    heading(content.methodology_section.heading),
    ...content.methodology_section.paragraphs.map(p => para(p)),
    pageBreakPara(),
  );

  // --- Closing / CTA ---
  sections.push(
    eyebrow('VISIT THE DASHBOARD'),
    heading(content.closing.heading),
    ...content.closing.paragraphs.map(p => para(p)),
    para(SITE_URL, { bold: true, size: 24, color: COLOR.accent, after: 320 }),
    new Paragraph({
      border: { top: { style: BorderStyle.SINGLE, size: 6, color: COLOR.rule, space: 8 } },
      spacing: { before: 200 },
      children: [new TextRun({ text: 'Adeo Patapo', bold: true, size: 20, color: COLOR.ink, font: 'Calibri' })],
    }),
    para('adeopatapo@gmail.com · +39 366 378 6189', { size: 18, color: COLOR.muted, after: 0 }),
  );

  const doc = new Document({
    sections: [{
      properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1080, bottom: 1080, left: 1080, right: 1080 } } },
      children: sections,
    }],
  });

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const outPath = path.join(OUT_DIR, `article-${region.date}.docx`);
  Packer.toBuffer(doc).then(buf => {
    fs.writeFileSync(outPath, buf);
    console.log(`Wrote ${outPath}`);

    const captionPath = path.join(OUT_DIR, `linkedin-draft-${region.date}.txt`);
    fs.writeFileSync(captionPath, content.caption, 'utf-8');
    console.log(`Wrote ${captionPath}`);
  });
}

main();
