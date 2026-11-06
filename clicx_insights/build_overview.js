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
pres.title = "CLICX customer overview";
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
// Board storyline: governing question -> 5 sub-questions -> engines + ranked bets
const SRC = "Source: Data Room; Data for Strategy 2027 (Nov 2026); CLICX 5-Year Business Plan; shareholder benefit analysis. CLICX data as of Sep 2026; AIS, KTB, OR as of Apr 2026";
const IC = (f) => __dirname + "/icons/" + f;
const STAT_BG = "E5F1F0";
// =====================================================================
// One-slide customer overview: age and occupation, segment actions, next steps
pres.addSection({ title: "Overview" });
{
  const s = pres.addSlide({ masterName: "QUESTION", sectionTitle: "Overview" });
  s.addText("CUSTOMER OVERVIEW  |  AGE AND OCCUPATION", { placeholder: "tracker" });
  s.addText("Most customers are 20–49; salaried and merchants engage most, gig and blue-collar least, so we act segment by segment", { placeholder: "title" });
  s.addText("1.51M customers as of 31 Oct 2026. Paid = paid with CLICX in October; teal = 15%+ above the 8.4% average, rose = 15%+ below. Salaried = corporate + government. Segments overlap. Source: payment-by-age and occupation files; Partner Payment Story v1", { placeholder: "foot" });
  const UP = "CDEAE7", DOWN = "F4DCDC", AVG = 8.4;
  const chip = (x, y, v, obj) => {
    const fill = v >= AVG * 1.15 ? UP : v <= AVG * 0.85 ? DOWN : "FFFFFF";
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: y + 0.03, w: 0.52, h: 0.2, rectRadius: 0.05, fill: { color: fill }, line: { type: "none" }, objectName: obj + " chip" });
    s.addText(v.toFixed(1) + "%", { x, y: y + 0.03, w: 0.52, h: 0.2, fontSize: 9, bold: true, color: NAVY, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: obj });
  };
  const listPanel = (x, y, w, h, title, rows, max, n) => {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.08, fill: { color: C.background2 }, line: { type: "none" }, objectName: n + " panel" });
    s.addText(title, { x: x + 0.15, y: y + 0.1, w: w - 0.3, h: 0.3, fontSize: 12, bold: true, color: NAVY, margin: 0, valign: "middle", isTextBox: true, objectName: n + " title" });
    const lw = rows.lw, bw = w - lw - 1.35 - 0.15;
    s.addText("Customers", { x: x + 0.15 + lw, y: y + 0.42, w: bw + 0.6, h: 0.2, fontSize: 8.5, bold: true, color: MUTED, margin: 0, isTextBox: true, objectName: n + " col1" });
    s.addText("Paid", { x: x + w - 0.67, y: y + 0.42, w: 0.52, h: 0.2, fontSize: 8.5, bold: true, color: MUTED, align: "center", margin: 0, isTextBox: true, objectName: n + " col2" });
    rows.data.forEach((r, k) => {
      const yy = y + 0.66 + k * rows.rh;
      s.addText(r[0], { x: x + 0.15, y: yy, w: lw, h: 0.26, fontSize: 9.5, color: TXT, bold: !!r[3], margin: 0, valign: "middle", isTextBox: true, objectName: `${n} label ${k}` });
      s.addShape(pres.shapes.RECTANGLE, { x: x + 0.15 + lw, y: yy + 0.07, w: Math.max(0.03, bw * r[1] / max), h: 0.13, fill: { color: r[3] ? CYAN : BLUE }, line: { type: "none" }, objectName: `${n} bar ${k}` });
      s.addText(r[1] >= 1000 ? (r[1] / 1000).toFixed(2) + "M" : r[1].toFixed(r[1] < 100 ? 1 : 0) + "K", { x: x + 0.15 + lw + Math.max(0.03, bw * r[1] / max) + 0.05, y: yy, w: 0.6, h: 0.26, fontSize: 9, color: TXT, margin: 0, valign: "middle", isTextBox: true, objectName: `${n} value ${k}` });
      chip(x + w - 0.67, yy, r[2], `${n} paid ${k}`);
    });
  };
  const top = 1.5, H = 5.28;
  // Age panel
  listPanel(0.5, top, 3.0, H, "By age", { lw: 0.55, rh: 0.42, data: [
    ["10–19", 10.9, 14.3], ["20–29", 395.3, 7.6], ["30–39", 580.3, 8.8, true], ["40–49", 369.3, 8.7], ["50–59", 126.5, 8.1], ["60–69", 22.2, 8.0], ["70+", 2.9, 10.4]] }, 580.3, "Age");
  s.addText([{ text: "89%", options: { bold: true, fontSize: 18, color: BLUE, breakLine: true } }, { text: "are aged 20–49; 80% of all customers hold no active product", options: { fontSize: 9, color: TXT } }],
    { x: 0.65, y: top + 3.75, w: 2.7, h: 1.35, margin: 0, valign: "top", isTextBox: true, objectName: "Age takeaway" });
  // Occupation panel
  listPanel(3.62, top, 3.4, H, "By occupation", { lw: 1.05, rh: 0.42, data: [
    ["Salaried", 575.4, 9.4, true], ["Blue collar", 224.1, 6.4], ["Freelance", 217.6, 7.3], ["Offline merchant", 148.9, 9.7], ["Wage earner", 104.7, 7.1],
    ["Online merchant", 94.8, 10.2], ["Rider", 58.3, 7.1], ["Agriculture", 55.6, 6.6], ["Other / unknown", 28.1, 11.5]] }, 575.4, "Occupation");
  s.addText([{ text: "44%", options: { bold: true, fontSize: 18, color: BLUE, breakLine: true } }, { text: "work without a regular wage (freelance, gig, blue-collar, farm), and pay least", options: { fontSize: 9, color: TXT } }],
    { x: 3.77, y: top + 4.55, w: 3.1, h: 0.68, margin: 0, valign: "top", isTextBox: true, objectName: "Occupation takeaway" });

  // Actions by segment
  const AX0 = 7.17, AW = 5.66;
  s.addText("WHAT TO DO, BY SEGMENT", { x: AX0, y: top, w: AW, h: 0.26, fontSize: 10, bold: true, color: MUTED, charSpacing: 1, margin: 0, isTextBox: true, objectName: "Actions label" });
  const seg = [
    ["p_corp", "Salaried workers", "575K · pay 9.4% · approved most for loans", "Pre-approved salary-bridged loans and a card (P1, P2)"],
    ["p_sme", "Merchants", "244K · pay 9.7–10.2% · save most", "Take QR payments in, then offer SME cash-flow loans"],
    ["p_blue", "Gig, blue-collar and farm", "660K · pay 6.4–7.3% · often turned down", "Savings and insurance first (P3), then a small starter loan"],
    ["p_tech", "Young adults, 20–29", "395K · pay 7.6% · 83% hold nothing", "First-deposit bonus and Scan and Pay rewards"],
    ["p_wealth", "Savers, 50+", "152K · saving rises with age", "Savings tiers, then investment (P4)"],
  ];
  const rh = 0.62;
  seg.forEach((g, k) => {
    const y = top + 0.32 + k * (rh + 0.06);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: AX0, y, w: AW, h: rh, rectRadius: 0.06, fill: { color: C.background2 }, line: { type: "none" }, objectName: `Seg ${k}` });
    s.addShape(pres.shapes.OVAL, { x: AX0 + 0.12, y: y + 0.12, w: 0.38, h: 0.38, fill: { color: k === 2 ? NAVY : CYAN }, line: { type: "none" }, objectName: `Seg icon bg ${k}` });
    s.addImage({ path: IC(g[0] + ".png"), x: AX0 + 0.21, y: y + 0.21, w: 0.2, h: 0.2, altText: g[1] + " icon", objectName: `Seg icon ${k}` });
    s.addText([{ text: g[1], options: { bold: true, fontSize: 10.5, color: NAVY, breakLine: true } }, { text: g[2], options: { fontSize: 8.5, color: MUTED } }],
      { x: AX0 + 0.6, y: y + 0.04, w: 2.15, h: rh - 0.08, margin: 0, valign: "middle", isTextBox: true, objectName: `Seg who ${k}` });
    s.addText([{ text: "→  ", options: { bold: true, color: CYAN } }, { text: g[3], options: { bold: true, color: BLUE } }],
      { x: AX0 + 2.8, y: y + 0.04, w: AW - 2.9, h: rh - 0.08, fontSize: 10, margin: 0, valign: "middle", isTextBox: true, objectName: `Seg action ${k}` });
  });
  // Next steps
  const ny = top + 0.32 + 5 * (rh + 0.06) + 0.06, nh = top + H - ny;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: AX0, y: ny, w: AW, h: nh, rectRadius: 0.08, fill: { color: C.text2 }, line: { type: "none" }, objectName: "Next steps box" });
  s.addText("NEXT STEPS", { x: AX0 + 0.15, y: ny + 0.08, w: 2, h: 0.24, fontSize: 9.5, bold: true, color: C.accent6, charSpacing: 1, margin: 0, isTextBox: true, objectName: "Next label" });
  const steps = [["30 days", "Name an owner and target per segment; add revenue and credit losses to the data room"],
    ["60 days", "Run three pilots with control groups: deposit bonus, salary pre-approval, savings for gig workers"],
    ["90 days", "Scale what works into the 2027 plan; align saver and depositor definitions"]];
  const sw = (AW - 0.3 - 0.2) / 3;
  steps.forEach((t, k) => {
    s.addText([{ text: t[0], options: { bold: true, fontSize: 11, color: "FFFFFF", breakLine: true } }, { text: t[1], options: { fontSize: 9, color: "DCE6EF" } }],
      { x: AX0 + 0.15 + k * (sw + 0.1), y: ny + 0.36, w: sw, h: nh - 0.44, margin: 0, valign: "top", isTextBox: true, objectName: `Step ${k}` });
  });
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
