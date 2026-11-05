// CLICX customer insights deck — question-led storyline, data from
// dtr_strategy2027_v4.pptx and CLICX_5-Year_Business_Plan_V1-Oct_2027.pptx
const pptxgen = require("pptxgenjs");
const { applyTheme } = require(process.env.SKILL + "/scripts/apply_theme.js");

const OUT = process.argv[2] || "CLICX_Customer_Insights.pptx";

const THEME = {
  name: "CLICX Blue Grey Teal",
  headFontFace: "Arial",
  bodyFontFace: "Arial",
  colors: {
    dk1: "1C2733", lt1: "FFFFFF", dk2: "1F3A5F", lt2: "EEF2F6",
    accent1: "2E6DB4", accent2: "1AA39A", accent3: "1F3A5F",
    accent4: "8A97A6", accent5: "5B6B7C", accent6: "9FD8D3",
    hlink: "2E6DB4", folHlink: "5B6B7C",
  },
};
// Hex mirrors for hex-only chart options
const NAVY = "1F3A5F", BLUE = "2E6DB4", CYAN = "1AA39A", GREY = "8A97A6",
  LGREY = "D3D9E0", RED = "5B6B7C", LCYAN = "9FD8D3", TXT = "1C2733", MUTED = "5B6B7C", TINT = "EEF2F6";

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "CLICX 2027 one-page summary";
pres.company = "CLICX";
const C = pres.SchemeColor;

// ---------- Layouts ----------
pres.defineSlideMaster({
  title: "TITLE",
  background: { color: C.text2 },
  objects: [
    { placeholder: { options: { name: "kicker", type: "body", x: 0.8, y: 2.0, w: 11.5, h: 0.4, fontSize: 14, bold: true, color: C.accent2, charSpacing: 2 }, text: "" } },
    { placeholder: { options: { name: "title", type: "title", x: 0.8, y: 2.45, w: 11.5, h: 1.4, fontSize: 40, bold: true, color: C.background1, valign: "top" }, text: "" } },
    { placeholder: { options: { name: "sub", type: "body", x: 0.8, y: 3.95, w: 10.5, h: 1.0, fontSize: 18, color: C.accent6, valign: "top" }, text: "" } },
    { placeholder: { options: { name: "date", type: "body", x: 0.8, y: 6.3, w: 8, h: 0.4, fontSize: 12, color: C.background1 }, text: "" } },
    { image: { path: __dirname + "/logo_white.png", x: 11.13, y: 0.5, w: 1.7, h: 0.426, altText: "CLICX logo" } },
  ],
});

pres.defineSlideMaster({
  title: "QUESTION",
  background: { color: C.background1 },
  margin: [0.5, 0.5, 0.6, 0.5],
  objects: [
    { placeholder: { options: { name: "tracker", type: "body", x: 0.5, y: 0.3, w: 9, h: 0.3, fontSize: 11, bold: true, color: C.accent1, charSpacing: 1, margin: 0 }, text: "" } },
    { placeholder: { options: { name: "title", type: "title", x: 0.5, y: 0.62, w: 12.33, h: 0.9, fontSize: 24, bold: true, color: C.text2, align: "left", valign: "top", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "answer", type: "body", x: 0.5, y: 1.62, w: 12.33, h: 0.65, fontSize: 16, color: C.text1, valign: "top", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "foot", type: "body", x: 0.5, y: 6.92, w: 11.2, h: 0.3, fontSize: 10, color: "5B6B7C", valign: "top", margin: 0 }, text: "" } },
    { image: { path: __dirname + "/logo.png", x: 11.63, y: 0.2, w: 1.2, h: 0.3, altText: "CLICX logo" } },
  ],
  slideNumber: { x: 12.33, y: 6.95, w: 0.5, h: 0.3, fontSize: 10, color: "5B6B7C", align: "right" },
});

// ---------- Helpers ----------
const AX = { catAxisLabelColor: MUTED, valAxisLabelColor: MUTED, catAxisLabelFontSize: 11, valAxisLabelFontSize: 10,
  catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt", dataLabelFontFace: "+mn-lt", legendFontFace: "+mn-lt",
  titleFontFace: "+mn-lt" };
const chartTitle = (t) => ({ showTitle: true, title: t, titleFontSize: 13, titleColor: NAVY, titleBold: true, titleAlign: "left" });
const quiet = () => ({ valGridLine: { style: "none" }, catGridLine: { style: "none" }, valAxisHidden: true,
  catAxisLineShow: true, catAxisLineColor: LGREY });

function qSlide(section, tracker, question, answer, foot) {
  const s = pres.addSlide({ masterName: "QUESTION", sectionTitle: section });
  s.addText(tracker, { placeholder: "tracker" });
  s.addText(question, { placeholder: "title" });
  s.addText(answer, { placeholder: "answer" });
  s.addText(foot, { placeholder: "foot" });
  return s;
}

// Insight panel: tinted box with numbered points
function insights(s, items, x = 9.05, y = 2.35, w = 3.78, h = 4.35, head = "So what") {
  s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: C.background2 }, line: { type: "none" }, objectName: "Insight panel" });
  s.addText(head.toUpperCase(), { x: x + 0.25, y: y + 0.18, w: w - 0.5, h: 0.3, fontSize: 11, bold: true, color: C.accent1, charSpacing: 1, margin: 0, isTextBox: true, objectName: "Insight header" });
  const step = (h - 0.65) / items.length;
  items.forEach((it, i) => {
    const yy = y + 0.6 + i * step;
    s.addShape(pres.shapes.OVAL, { x: x + 0.25, y: yy + 0.02, w: 0.32, h: 0.32, fill: { color: C.text2 }, line: { type: "none" }, objectName: `Badge ${i + 1}` });
    s.addText(String(i + 1), { x: x + 0.25, y: yy + 0.02, w: 0.32, h: 0.32, fontSize: 11, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: `Badge num ${i + 1}` });
    s.addText([
      { text: it[0], options: { bold: true, fontSize: 14, color: C.text2, breakLine: true } },
      { text: it[1], options: { fontSize: 12, color: C.text1 } },
    ], { x: x + 0.7, y: yy - 0.03, w: w - 0.9, h: step - 0.08, valign: "top", margin: 0, paraSpaceAfter: 3, isTextBox: true, objectName: `Insight ${i + 1}` });
  });
}

// Big stat callout
function stat(s, x, y, w, big, label, color = C.accent1) {
  s.addText(big, { x, y, w, h: 0.7, fontSize: 40, bold: true, color, margin: 0, isTextBox: true, objectName: `Stat ${big}` });
  s.addText(label, { x, y: y + 0.72, w, h: 0.55, fontSize: 12, color: C.text1, margin: 0, valign: "top", isTextBox: true, objectName: `Stat label ${big}` });
}

// =====================================================================
// One-page summary: five key questions, one chart each
pres.addSection({ title: "Summary" });
{
  const s = qSlide("Summary", "ONE-PAGE SUMMARY  |  DATA FOR STRATEGY 2027",
    "What have our first 2.1 million customers taught us for 2027?",
    "Partners bring in almost everyone and we reached our targets. Our customers are ready leads for 2027 products, but only 15% are active, and a loan is what makes them active",
    "Source: Data for Strategy 2027 (Nov 2026); CLICX data as of Sep 2026; AIS, KTB, OR data as of Apr 2026. Monthly active = transacted in Sep 2026");
  const cols = [
    { q: "Do our customers come from our partners?", a: "Almost all: 96%",
      title: "% of customers also with partner", type: "hbar", labels: ["KTB", "AIS", "OR"], values: [0.88, 0.67, 0.40], fmt: "0%", max: 1.15, colors: [NAVY, BLUE, CYAN],
      next: "Grow AIS: only 3.7% of its base has joined" },
    { q: "Are we winning the people we targeted?", a: "Mostly",
      title: "% of each group who joined", type: "hbar", labels: ["Tech-savvy", "Gen Z 23–29", "Average", "Multi-bank"], values: [0.085, 0.066, 0.038, 0.024], fmt: "0.0%", max: 0.11, colors: [CYAN, CYAN, LGREY, GREY],
      next: "82% of loan applicants are underserved, as planned. Win multi-bank users next" },
    { q: "Which customers lead to 2027 products?", a: "Three lead pools",
      title: "Lead pools, thousand customers", type: "hbar", labels: ["Office workers (card)", "Savers (wealth)", "Merchants (SME)"], values: [700, 643, 293], fmt: "#,##0", max: 900, colors: [NAVY, BLUE, CYAN],
      next: "Set 2027 targets for each product. 5.5B THB of deposits to build on" },
    { q: "Why are most customers inactive?", a: "Loans drive activity",
      title: "% active each month", type: "hbar", labels: ["Borrowers", "Money in, no loan", "No money in"], values: [0.88, 0.51, 0.04], fmt: "0%", max: 1.15, colors: [CYAN, BLUE, GREY],
      next: "Activate 1.8M idle customers with a first deposit and a first loan" },
    { q: "Do we take customers from KTB?", a: "No, we recover them",
      title: "Turned down by KTB's Jai-pump, thousand", type: "col", labels: ["Rejected", "Applied to us", "Got our loan"], values: [1900, 450, 51], fmt: "#,##0", max: 2350, colors: [GREY, BLUE, CYAN],
      next: "1.1M shared customers were KTB-inactive. Work the reject list before True" },
  ];
  const cw = 2.35, gap = 0.145, top = 2.3;
  cols.forEach((c, i) => {
    const x = 0.5 + i * (cw + gap);
    s.addShape(pres.shapes.RECTANGLE, { x, y: top, w: cw, h: 4.45, fill: { color: C.background2 }, line: { type: "none" }, objectName: `Col ${i + 1}` });
    s.addShape(pres.shapes.OVAL, { x: x + 0.15, y: top + 0.15, w: 0.36, h: 0.36, fill: { color: C.text2 }, line: { type: "none" }, objectName: `Badge ${i + 1}` });
    s.addText(String(i + 1), { x: x + 0.15, y: top + 0.15, w: 0.36, h: 0.36, fontSize: 12, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: `Badge num ${i + 1}` });
    s.addText(c.q, { x: x + 0.6, y: top + 0.1, w: cw - 0.72, h: 0.62, fontSize: 12, bold: true, color: C.text2, valign: "middle", margin: 0, isTextBox: true, objectName: `Question ${i + 1}` });
    s.addText(c.a, { x: x + 0.15, y: top + 0.8, w: cw - 0.3, h: 0.36, fontSize: 14, bold: true, color: C.accent2, margin: 0, isTextBox: true, objectName: `Answer ${i + 1}` });
    s.addChart(pres.charts.BAR, [{ name: c.title, labels: c.labels, values: c.values }], {
      x: x + 0.05, y: top + 1.2, w: cw - 0.1, h: 2.45, barDir: c.type === "col" ? "col" : "bar", ...AX, ...quiet(), ...chartTitle(c.title), titleFontSize: 10,
      chartColors: c.colors, showValue: true, dataLabelPosition: "outEnd", dataLabelColor: TXT, dataLabelFontSize: 10, dataLabelFormatCode: c.fmt,
      showLegend: false, barGapWidthPct: 45, valAxisMaxVal: c.max, catAxisOrientation: c.type === "col" ? "minMax" : "maxMin", catAxisLabelFontSize: 9, objectName: `Chart ${i + 1}` });
    s.addText([
      { text: "NEXT  ", options: { bold: true, fontSize: 9.5, color: C.accent1, charSpacing: 1 } },
      { text: c.next, options: { fontSize: 10.5, color: C.text1 } },
    ], { x: x + 0.15, y: top + 3.72, w: cw - 0.3, h: 0.68, valign: "top", margin: 0, isTextBox: true, objectName: `Next ${i + 1}` });
  });
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
