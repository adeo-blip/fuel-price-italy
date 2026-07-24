// Renders the day's Fuel Dashboard regional article (title/dek/body/pull-quote/caption
// are composed fresh each day by the scheduled agent and handed in via
// data/article_content.json) into a polished ~2-page .docx, using the numbers in
// data/region_of_day.json for the stat callout and recent-history table.
'use strict';
const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle,
  PageBreak,
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
  benzina: '2A78D6',
  gasolio: '1BAF7A',
  rule: 'E6E2D8',
};

function fmtPrice(v) { return v == null ? '—' : `€${v.toFixed(3)}`; }
function fmtDelta(v) {
  if (v == null) return '';
  const sign = v >= 0 ? '+' : '';
  const arrow = v < 0 ? '↓' : '↑';
  return `${arrow} ${sign}${v.toFixed(3)} €/L`;
}
function fmtDateShort(iso) {
  const [y, m, d] = iso.split('-').map(Number);
  const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  return `${months[m - 1]} ${d}`;
}

function statCell(label, value, sub) {
  return new TableCell({
    width: { size: 2500, type: WidthType.DXA },
    shading: { type: ShadingType.CLEAR, fill: COLOR.accentSoft },
    margins: { top: 160, bottom: 160, left: 180, right: 180 },
    children: [
      new Paragraph({ children: [new TextRun({ text: label.toUpperCase(), bold: true, size: 16, color: COLOR.muted, font: 'Calibri' })] }),
      new Paragraph({ spacing: { before: 60 }, children: [new TextRun({ text: value, bold: true, size: 32, color: COLOR.ink, font: 'Calibri' })] }),
      sub ? new Paragraph({ spacing: { before: 40 }, children: [new TextRun({ text: sub, size: 18, color: COLOR.muted, font: 'Calibri' })] }) : new Paragraph({ children: [] }),
    ],
  });
}

function main() {
  const region = JSON.parse(fs.readFileSync(REGION_JSON, 'utf-8'));
  const content = JSON.parse(fs.readFileSync(CONTENT_JSON, 'utf-8'));

  const latest = region.latest;
  const previous = region.previous;
  const benzinaDelta = previous ? latest.Benzina.self - previous.Benzina.self : null;
  const gasolioDelta = previous ? latest.Gasolio.self - previous.Gasolio.self : null;

  const historyRows = region.history.slice(-7).map(h => new TableRow({
    children: [
      new TableCell({ width: { size: 2000, type: WidthType.DXA }, children: [new Paragraph({ children: [new TextRun({ text: fmtDateShort(h.date), size: 20, font: 'Calibri' })] })] }),
      new TableCell({ width: { size: 2000, type: WidthType.DXA }, children: [new Paragraph({ children: [new TextRun({ text: fmtPrice(h.Benzina.self), size: 20, font: 'Calibri', color: COLOR.benzina })] })] }),
      new TableCell({ width: { size: 2000, type: WidthType.DXA }, children: [new Paragraph({ children: [new TextRun({ text: fmtPrice(h.Gasolio.self), size: 20, font: 'Calibri', color: COLOR.gasolio })] })] }),
    ],
  }));

  const doc = new Document({
    sections: [{
      properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1080, bottom: 1080, left: 1080, right: 1080 } } },
      children: [
        new Paragraph({
          children: [new TextRun({ text: 'FUEL DASHBOARD · REGIONE SPOTLIGHT', bold: true, size: 18, color: COLOR.accent, font: 'Calibri' })],
        }),
        new Paragraph({
          spacing: { before: 120, after: 80 },
          children: [new TextRun({ text: content.title, bold: true, size: 56, color: COLOR.ink, font: 'Cambria' })],
        }),
        new Paragraph({
          spacing: { after: 240 },
          children: [new TextRun({ text: content.dek, italics: true, size: 24, color: COLOR.muted, font: 'Calibri' })],
        }),

        new Table({
          width: { size: 10000, type: WidthType.DXA },
          columnWidths: [2500, 2500, 2500, 2500],
          borders: {
            top: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
            bottom: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
            left: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
            right: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
            insideHorizontal: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
            insideVertical: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
          },
          rows: [new TableRow({
            children: [
              statCell('Benzina', fmtPrice(latest.Benzina.self), fmtDelta(benzinaDelta)),
              statCell('Gasolio', fmtPrice(latest.Gasolio.self), fmtDelta(gasolioDelta)),
              statCell('Registered stations', latest.stations.registered.toLocaleString('en-US'), region.region),
              statCell('Reporting today', latest.stations.reporting.toLocaleString('en-US'), fmtDateShort(latest.date)),
            ],
          })],
        }),

        new Paragraph({ spacing: { before: 320 }, children: [] }),

        ...content.body.map(para => new Paragraph({
          spacing: { after: 200 },
          children: [new TextRun({ text: para, size: 22, color: COLOR.ink, font: 'Calibri' })],
        })),

        new Paragraph({
          spacing: { before: 120, after: 240 },
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: content.pull_quote, italics: true, bold: true, size: 30, color: COLOR.accent, font: 'Cambria' })],
        }),

        new Paragraph({
          spacing: { before: 120, after: 120 },
          children: [new TextRun({ text: `Last ${historyRows.length} tracked days in ${region.region} — self-service average`, bold: true, size: 18, color: COLOR.muted, font: 'Calibri' })],
        }),
        new Table({
          width: { size: 6000, type: WidthType.DXA },
          columnWidths: [2000, 2000, 2000],
          rows: [
            new TableRow({ children: [
              new TableCell({ width: { size: 2000, type: WidthType.DXA }, shading: { type: ShadingType.CLEAR, fill: 'F0EDE6' }, children: [new Paragraph({ children: [new TextRun({ text: 'Date', bold: true, size: 20, font: 'Calibri' })] })] }),
              new TableCell({ width: { size: 2000, type: WidthType.DXA }, shading: { type: ShadingType.CLEAR, fill: 'F0EDE6' }, children: [new Paragraph({ children: [new TextRun({ text: 'Benzina', bold: true, size: 20, font: 'Calibri' })] })] }),
              new TableCell({ width: { size: 2000, type: WidthType.DXA }, shading: { type: ShadingType.CLEAR, fill: 'F0EDE6' }, children: [new Paragraph({ children: [new TextRun({ text: 'Gasolio', bold: true, size: 20, font: 'Calibri' })] })] }),
            ] }),
            ...historyRows,
          ],
        }),

        new Paragraph({ children: [new PageBreak()] }),

        new Paragraph({
          spacing: { after: 160 },
          children: [new TextRun({ text: 'See the full trend — daily, weekly, monthly', bold: true, size: 30, color: COLOR.ink, font: 'Cambria' })],
        }),
        new Paragraph({
          spacing: { after: 200 },
          children: [new TextRun({ text: `Fuel Dashboard tracks Italy's pump prices nationally and down to any Regione, Provincia, Comune or Gestore, with live counts of registered vs. actively reporting stations — free to browse in English, Italiano, Français, Русский or 中文.`, size: 22, color: COLOR.ink, font: 'Calibri' })],
        }),
        new Paragraph({
          spacing: { after: 320 },
          children: [new TextRun({ text: SITE_URL, bold: true, size: 24, color: COLOR.accent, font: 'Calibri' })],
        }),
        new Paragraph({
          border: { top: { style: BorderStyle.SINGLE, size: 6, color: COLOR.rule, space: 8 } },
          spacing: { before: 200 },
          children: [new TextRun({ text: 'Adeo Patapo', bold: true, size: 20, color: COLOR.ink, font: 'Calibri' })],
        }),
        new Paragraph({
          children: [new TextRun({ text: 'adeopatapo@gmail.com · +39 366 378 6189', size: 18, color: COLOR.muted, font: 'Calibri' })],
        }),
      ],
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
