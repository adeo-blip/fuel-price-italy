// Generates the daily 2-slide "Fuel Dashboard" deck (Benzina trend, Gasolio trend) from
// data.json, plus a plain-text LinkedIn caption draft for human review before posting.
// Run after update_data.py so data.json reflects today's figures.
'use strict';
const fs = require('fs');
const path = require('path');
const pptxgen = require('pptxgenjs');

const BASE_DIR = __dirname;
const DATA_JSON = path.join(BASE_DIR, 'data.json');
const OUT_DIR = path.join(BASE_DIR, 'slides');
const SITE_URL = 'https://adeo-blip.github.io/fuel-price-italy/';

const COLORS = {
  bg: '14120F',       // matches the dashboard's dark theme background
  text: 'F5F1E8',     // warm off-white body/heading text
  muted: 'C9C2B4',    // muted secondary text
  benzina: '3987E5',  // same blue used for Benzina on the dashboard
  gasolio: '199E70',  // same green used for Gasolio on the dashboard
  gold: 'F2A900',     // dashboard brand accent
  good: '3DDC5A',
  bad: 'E66767',
};

const FUEL_META = {
  Benzina: { label: 'Benzina', gloss: 'petrol', color: COLORS.benzina },
  Gasolio: { label: 'Gasolio', gloss: 'diesel', color: COLORS.gasolio },
};

function loadHistory() {
  const doc = JSON.parse(fs.readFileSync(DATA_JSON, 'utf-8'));
  return doc.history || [];
}

function fmtShortDate(iso) {
  const [y, m, d] = iso.split('-').map(Number);
  const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  return `${months[m - 1]} ${d}`;
}

function buildSlide(pres, fuel, history) {
  const meta = FUEL_META[fuel];
  const recent = history.slice(-14);
  const series = recent
    .map(h => ({ date: h.date, self: h[fuel] && h[fuel].self != null ? h[fuel].self : null }))
    .filter(p => p.self != null);

  const latest = series.length ? series[series.length - 1] : null;
  const prev = series.length > 1 ? series[series.length - 2] : null;
  const delta = latest && prev ? latest.self - prev.self : null;
  const deltaPct = delta != null ? (delta / prev.self) * 100 : null;

  const slide = pres.addSlide();
  slide.background = { color: COLORS.bg };

  // Brand mark (top-left)
  slide.addShape('ellipse', { x: 0.6, y: 0.55, w: 0.16, h: 0.16, fill: { color: COLORS.gold }, line: { type: 'none' } });
  slide.addText('FUEL DASHBOARD', {
    x: 0.86, y: 0.45, w: 4.5, h: 0.35, fontFace: 'Calibri', fontSize: 13, bold: true,
    color: COLORS.muted, charSpacing: 2, margin: 0,
  });

  // Title
  slide.addText(`${meta.label} price trend`, {
    x: 0.6, y: 0.95, w: 5.4, h: 0.65, fontFace: 'Calibri', fontSize: 34, bold: true,
    color: COLORS.text, margin: 0, valign: 'top',
  });
  slide.addText(`National self-service average · ${meta.gloss}`, {
    x: 0.6, y: 1.75, w: 5.4, h: 0.3, fontFace: 'Calibri', fontSize: 13, color: COLORS.muted, margin: 0,
  });

  // Big stat
  slide.addText(latest ? `€${latest.self.toFixed(3)}` : '—', {
    x: 0.6, y: 2.25, w: 5.4, h: 1.2, fontFace: 'Calibri', fontSize: 72, bold: true,
    color: meta.color, margin: 0, valign: 'top',
  });
  slide.addText('per litre', {
    x: 0.6, y: 3.55, w: 5.4, h: 0.3, fontFace: 'Calibri', fontSize: 12, color: COLORS.muted, margin: 0,
  });

  if (delta != null) {
    const down = delta < 0;
    const arrow = down ? '↓' : '↑';
    const sign = delta >= 0 ? '+' : '';
    slide.addText(`${arrow} ${sign}${delta.toFixed(3)} €/L  (${sign}${deltaPct.toFixed(2)}%) vs. previous day`, {
      x: 0.6, y: 3.95, w: 5.4, h: 0.4, fontFace: 'Calibri', fontSize: 15, bold: true,
      color: down ? COLORS.good : COLORS.bad, margin: 0,
    });
  }

  // CTA / brand block (bottom-left)
  slide.addText([
    { text: 'See the full daily, weekly and monthly trend\n', options: { fontSize: 15, color: COLORS.text, bold: true, breakLine: true } },
    { text: SITE_URL, options: { fontSize: 13, color: COLORS.gold, bold: true } },
  ], { x: 0.6, y: 5.85, w: 5.4, h: 0.75, fontFace: 'Calibri', margin: 0, valign: 'top' });

  // Trend chart (right)
  if (series.length >= 2) {
    slide.addChart('line', [{
      name: meta.label,
      labels: series.map(p => fmtShortDate(p.date)),
      values: series.map(p => p.self),
    }], {
      x: 6.15, y: 0.9, w: 6.55, h: 5.9,
      showTitle: true, title: `Last ${series.length} days · €/L`, titleColor: COLORS.muted, titleFontSize: 13,
      showLegend: false,
      chartColors: [meta.color],
      lineSize: 2.75, lineDataSymbol: 'circle', lineDataSymbolSize: 5,
      showValue: false,
      catAxisLabelColor: COLORS.muted, catAxisLabelFontSize: 10, catAxisLineColor: COLORS.muted,
      valAxisLabelColor: COLORS.muted, valAxisLabelFontSize: 10, valAxisLineColor: COLORS.muted,
      valAxisLabelFormatCode: '€0.00',
      valGridLine: { color: '332F28', size: 1 },
      catGridLine: { style: 'none' },
      plotArea: { fill: { color: COLORS.bg } },
      chartArea: { fill: { color: COLORS.bg } },
    });
  } else {
    slide.addText('History for this chart starts building from today.', {
      x: 6.15, y: 3.0, w: 6.55, h: 0.6, fontFace: 'Calibri', fontSize: 14, color: COLORS.muted,
      align: 'center', margin: 0,
    });
  }

  slide.addNotes(
    `Fuel Dashboard daily update. ${meta.label}: current €${latest ? latest.self.toFixed(3) : 'n/a'}/L. ` +
    `Invite LinkedIn viewers to explore the interactive dashboard at ${SITE_URL}.`
  );
}

function buildLinkedInDraft(history) {
  const last = history[history.length - 1] || {};
  const b = last.Benzina && last.Benzina.self != null ? `€${last.Benzina.self.toFixed(3)}` : 'n/a';
  const g = last.Gasolio && last.Gasolio.self != null ? `€${last.Gasolio.self.toFixed(3)}` : 'n/a';
  const stations = last.stations || {};
  const fmt = n => (typeof n === 'number' ? n.toLocaleString('en-US') : '?');
  return `Today's Italy fuel prices — Benzina ${b}/L, Gasolio ${g}/L (national self-service average).

I built Fuel Dashboard (${SITE_URL}) to track this automatically — daily, weekly and monthly trends for the whole country, or drilled down to a specific Regione, Provincia, Comune or Gestore, with live counts of registered vs. actively reporting stations (${fmt(stations.registered)} registered / ${fmt(stations.reporting)} reporting today).

What used to be a manual spreadsheet refreshed by hand is now a live, multi-language, self-updating dashboard — built entirely with Claude Code. No manual data pulls, no stale Excel file: the pipeline fetches, computes and republishes itself every day.

Take a look: ${SITE_URL}

#FuelPrices #Italy #DataVisualization #ClaudeCode #Automation #Dashboard

--
Draft only — review before posting. Adeo Patapo · adeopatapo@gmail.com
`;
}

async function main() {
  const history = loadHistory();
  if (!history.length) {
    console.error('ERROR: data.json has no history — run update_data.py first.');
    process.exit(1);
  }

  const pres = new pptxgen();
  pres.layout = 'LAYOUT_WIDE';
  pres.author = 'Fuel Dashboard';
  pres.company = 'Fuel Dashboard';
  pres.title = 'Fuel Dashboard — Daily Update';

  buildSlide(pres, 'Benzina', history);
  buildSlide(pres, 'Gasolio', history);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const dateStr = (history[history.length - 1].date || 'latest');
  const outPath = path.join(OUT_DIR, `fuel-dashboard-${dateStr}.pptx`);
  await pres.writeFile({ fileName: outPath });
  console.log(`Wrote ${outPath}`);

  const draftPath = path.join(OUT_DIR, `linkedin-draft-${dateStr}.txt`);
  fs.writeFileSync(draftPath, buildLinkedInDraft(history), 'utf-8');
  console.log(`Wrote ${draftPath}`);
}

main().catch(err => { console.error(err); process.exit(1); });
