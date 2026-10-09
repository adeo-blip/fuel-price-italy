const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");

const OUT = process.env.FUEL_OUTDIR || __dirname;
const BEFORE = path.join(__dirname, "assets", "before.png");
const AFTER = path.join(OUT, "after.png");

const NAVY = "1E2761", NAVY_DEEP = "161C4C", ICE = "CADCFC", WHITE = "FFFFFF";
const MUTED = "8E9BD0", CORAL = "E8654F", MINT = "34C79A";

function pngSize(file) {
  const b = fs.readFileSync(file);
  return { w: b.readUInt32BE(16), h: b.readUInt32BE(20) };
}

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
const PW = 13.33;
const slide = pres.addSlide();
slide.background = { color: NAVY };

slide.addShape("roundRect", { x: 0.45, y: 1.55, w: PW - 0.9, h: 4.55, rectRadius: 0.14, fill: { color: NAVY_DEEP }, line: { type: "none" } });
slide.addText("AI-POWERED DASHBOARD OPTIMIZATION", { x: 0.5, y: 0.42, w: 9, h: 0.32, fontFace: "Calibri", fontSize: 12, bold: true, color: MINT, charSpacing: 2 });
slide.addText("Fuel Price Italy Dashboard — Before vs. After", { x: 0.5, y: 0.72, w: 12.3, h: 0.62, fontFace: "Cambria", fontSize: 30, bold: true, color: WHITE });
slide.addText("Redesigned with Claude AI: live KPIs, region and brand slicers, and every fuel compared with the national average.", { x: 0.5, y: 1.28, w: 12.3, h: 0.32, fontFace: "Calibri", fontSize: 13, color: ICE });

const panelInset = 0.3, colGap = 0.35;
const colWidth = (PW - 0.9 - panelInset * 2 - colGap) / 2;
const leftX = 0.45 + panelInset, rightX = leftX + colWidth + colGap;
const imgTop = 2.15, imgH = 3.15;

function fit(file) {
  const { w, h } = pngSize(file);
  const aspect = w / h;
  const width = Math.min(colWidth, imgH * aspect);
  return { w: width, h: width / aspect };
}
const b = fit(BEFORE), a = fit(AFTER);
const bx = leftX + (colWidth - b.w) / 2, ax = rightX + (colWidth - a.w) / 2;

slide.addShape("roundRect", { x: leftX, y: imgTop - 0.46, w: 1.15, h: 0.34, rectRadius: 0.17, fill: { color: CORAL }, line: { type: "none" } });
slide.addText("BEFORE", { x: leftX, y: imgTop - 0.46, w: 1.15, h: 0.34, align: "center", valign: "middle", margin: 0, fontFace: "Calibri", fontSize: 12, bold: true, color: WHITE });
slide.addShape("roundRect", { x: rightX, y: imgTop - 0.46, w: 1.05, h: 0.34, rectRadius: 0.17, fill: { color: MINT }, line: { type: "none" } });
slide.addText("AFTER", { x: rightX, y: imgTop - 0.46, w: 1.05, h: 0.34, align: "center", valign: "middle", margin: 0, fontFace: "Calibri", fontSize: 12, bold: true, color: NAVY_DEEP });

slide.addShape("roundRect", { x: bx - 0.06, y: imgTop - 0.06, w: b.w + 0.12, h: b.h + 0.12, rectRadius: 0.08, fill: { color: WHITE }, line: { type: "none" }, shadow: { type: "outer", color: "000000", opacity: 0.35, blur: 10, offset: 4, angle: 90 } });
slide.addImage({ path: BEFORE, x: bx, y: imgTop, w: b.w, h: b.h });
slide.addShape("roundRect", { x: ax - 0.06, y: imgTop - 0.06, w: a.w + 0.12, h: a.h + 0.12, rectRadius: 0.08, fill: { color: WHITE }, line: { type: "none" }, shadow: { type: "outer", color: "000000", opacity: 0.35, blur: 10, offset: 4, angle: 90 } });
slide.addImage({ path: AFTER, x: ax, y: imgTop, w: a.w, h: a.h });

slide.addText("8 sprawling slicers, truncated axes, no headline number — reading it takes effort.", { x: leftX, y: imgTop + b.h + 0.14, w: colWidth, h: 0.5, fontFace: "Calibri", fontSize: 11, color: MUTED, italic: true });
slide.addText("2 focused slicers, live KPI tiles, region and brand drill-down, each fuel vs. the national average.", { x: rightX, y: imgTop + a.h + 0.14, w: colWidth, h: 0.5, fontFace: "Calibri", fontSize: 11, color: ICE, italic: true });

const today = new Date().toLocaleDateString("en-US", { year: "numeric", month: "long", day: "numeric" });
slide.addText(today, { x: 0.5, y: 7.02, w: 4, h: 0.32, fontFace: "Calibri", fontSize: 11, color: MUTED });
slide.addText("Adeodat Turatsinze  ·  Powered by Claude AI", { x: 6.8, y: 7.02, w: 6.0, h: 0.32, align: "right", fontFace: "Calibri", fontSize: 11, bold: true, color: ICE });

pres.writeFile({ fileName: path.join(OUT, "Fuel_Price_Dashboard_Before_After.pptx") }).then(() => console.log("slide written"));
