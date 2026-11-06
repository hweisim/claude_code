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
pres.title = "CLICX customer landscape";
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
// =====================================================================
// One-slide landscape: age x job usage matrix with product holding, inferred to partner customers not yet on CLICX
pres.addSection({ title: "Landscape" });
{
  const s = pres.addSlide({ masterName: "QUESTION", sectionTitle: "Landscape" });
  s.addText("CUSTOMER LANDSCAPE  |  AGE × JOB × PRODUCT", { placeholder: "tracker" });
  s.addText("Age and job together show who uses what today; applying that to partner customers not yet on CLICX shows where to grow", { placeholder: "title" });
  s.addText("Paid in Oct 2026: teal 15%+ above the 8.4% average, rose 15%+ below; * under 1,000 customers. Holdings: by age, product segments (31 Oct); by job, Data for Strategy 2027 (Sep). Inference assumes joiners behave like today's customers in the same group; groups overlap", { placeholder: "foot" });
  const UP = "CDEAE7", DOWN = "F4DCDC", AVG = 8.4;
  const top = 1.5;
  // ---- Left: matrix
  s.addText("TODAY'S CUSTOMERS  ·  % who paid in October, and what each group holds", { x: 0.5, y: top, w: 7.9, h: 0.26, fontSize: 10, bold: true, color: MUTED, charSpacing: 1, margin: 0, isTextBox: true, objectName: "Matrix label" });
  const H = (t, al = "center") => ({ text: t, options: { bold: true, fontSize: 9.5, color: NAVY, align: al, valign: "bottom", border: [{ type: "none" }, { type: "none" }, { pt: 1.25, color: NAVY }, { type: "none" }] } });
  const rowLine = { border: [{ type: "none" }, { type: "none" }, { pt: 0.5, color: "E3E8EE" }, { type: "none" }] };
  const cell = (v) => {
    const small = typeof v === "string";
    const n = parseFloat(v);
    const o = { fontSize: 10.5, align: "center", color: small ? GREY : TXT, ...rowLine };
    if (!small) { if (n >= AVG * 1.15) o.fill = { color: UP }; else if (n <= AVG * 0.85) o.fill = { color: DOWN }; }
    return { text: small ? v : n.toFixed(1) + "%", options: o };
  };
  const lab = (t, bold) => ({ text: t, options: { fontSize: 10, bold: true, color: NAVY, align: "left", ...rowLine } });
  const hold = (b, sv, p) => ({ text: [{ text: b.replace("%", ""), options: { color: NAVY } }, { text: " · ", options: { color: MUTED } }, { text: sv.replace("%", ""), options: { color: CYAN } }, { text: " · ", options: { color: MUTED } }, { text: p.replace("%", ""), options: { color: BLUE } }],
    options: { fontSize: 10, bold: true, align: "center", fill: { color: "F5F7FA" }, ...rowLine } });
  const rows = [
    [H("Age", "left"), H("Salaried"), H("Offline merchant"), H("Online merchant"), H("Freelance"), H("Rider"), H("Blue collar"), H("Holds, %: borrow · save · pay")],
    [lab("20–29"), cell(8.5), cell(8.4), cell(9.5), cell(6.8), cell(5.6), cell(6.0), hold("11%", "4%", "7%")],
    [lab("30–39"), cell(9.8), cell(9.9), cell(10.6), cell(7.5), cell(7.1), cell(7.0), hold("12%", "7%", "8%")],
    [lab("40–49"), cell(9.7), cell(10.2), cell(10.5), cell(7.5), cell(8.2), cell(6.3), hold("11%", "11%", "7%")],
    [lab("50–59"), cell(8.8), cell(9.6), cell(9.7), cell(7.3), cell(8.6), cell(5.3), hold("9%", "17%", "7%")],
    [lab("60–69"), cell(7.5), cell(8.0), cell(7.0), cell(8.1), cell("9.0%*"), cell("3.3%*"), hold("0%", "28%", "7%")],
    [lab("All ages"), cell(9.4), cell(9.7), cell(10.2), cell(7.3), cell(7.1), cell(6.4), hold("11%", "8%", "7%")],
    [{ text: "Holds, %", options: { fontSize: 9.5, bold: true, color: NAVY, fill: { color: "F5F7FA" }, align: "left" } },
     hold("11%", "18%", "14%"), { ...hold("7%", "20%", "14%"), options: { ...hold("7%", "20%", "14%").options, colspan: 2 } },
     hold("6%", "16%", "12%"), hold("5%", "13%", "11%"), hold("7%", "12%", "10%"),
     { text: "", options: { fill: { color: "F5F7FA" } } }],
  ];
  // bold the "All ages" row
  rows[6].forEach((c) => { if (c.options) c.options.bold = true; });
  s.addTable(rows, { x: 0.5, y: top + 0.32, w: 7.9, colW: [0.8, 0.95, 0.95, 0.95, 0.95, 0.85, 0.95, 1.5], fontFace: "Arial",
    rowH: [0.52, 0.4, 0.4, 0.4, 0.4, 0.4, 0.4, 0.5], margin: [0.03, 0.05, 0.03, 0.05], valign: "middle", objectName: "Landscape matrix" });
  // legend
  const ly = top + 0.32 + 3.42 + 0.08;
  s.addText([{ text: "Holds: ", options: { bold: true, color: MUTED } }, { text: "borrow", options: { bold: true, color: NAVY } }, { text: " · ", options: { color: MUTED } },
    { text: "save / deposit", options: { bold: true, color: CYAN } }, { text: " · ", options: { color: MUTED } }, { text: "pay", options: { bold: true, color: BLUE } },
    { text: "   % of each group's customers holding each product", options: { color: MUTED } }],
    { x: 0.5, y: ly, w: 7.9, h: 0.24, fontSize: 9, margin: 0, isTextBox: true, objectName: "Holds legend" });
  // what we see
  const wy = ly + 0.34, ww = (7.9 - 0.24) / 3;
  const see = [["Steady earners pay most", "Merchants and salaried pay 9–11% at almost every age; blue-collar 5–7%"],
    ["Products follow life stage", "Borrowing peaks at 30–49 (11–12%); saving rises to 28% at 60–69"],
    ["Most hold nothing yet", "74–83% of each working-age group hold no active product"]];
  see.forEach((t, k) => {
    const x = 0.5 + k * (ww + 0.12);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: wy, w: ww, h: 6.78 - wy, rectRadius: 0.08, fill: { color: STAT_BG }, line: { type: "none" }, objectName: `See ${k}` });
    s.addText([{ text: t[0], options: { bold: true, fontSize: 11, color: NAVY, breakLine: true } }, { text: t[1], options: { fontSize: 9.5, color: TXT } }],
      { x: x + 0.14, y: wy + 0.08, w: ww - 0.28, h: 6.78 - wy - 0.16, margin: 0, valign: "top", isTextBox: true, objectName: `See text ${k}` });
  });

  // ---- Right: inference to partner customers not yet on CLICX
  const RX = 8.6, RW = 4.23;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: RX, y: top, w: RW, h: 5.28, rectRadius: 0.08, fill: { color: C.background2 }, line: { type: "none" }, objectName: "Infer panel" });
  s.addText("IF PARTNER CUSTOMERS NOT YET ON CLICX JOIN", { x: RX + 0.15, y: top + 0.1, w: RW - 0.3, h: 0.26, fontSize: 9.5, bold: true, color: BLUE, charSpacing: 1, margin: 0, isTextBox: true, objectName: "Infer label" });
  s.addText("If they joined and behaved like today's customers, per 100K joiners:", { x: RX + 0.15, y: top + 0.38, w: RW - 0.3, h: 0.24, fontSize: 9, color: MUTED, margin: 0, isTextBox: true, objectName: "Infer sub" });
  const inf = [
    ["p_corp", "Salaried", "7.5M", 7.5, "11K borrow · 18K save · 14K pay", "Biggest, best all-round pool: card and salary-bridged loans"],
    ["p_sme", "Merchants", "3.0M", 3.0, "7K borrow · 20K save · 14K pay", "Savers and payers: QR, deposits, SME credit"],
    ["p_blue", "Blue collar", "2.1M", 2.1, "7K borrow · 12K save · 10K pay", "Lowest use: savings and insurance first"],
    ["p_rider", "Freelance and riders", "0.6M", 0.6, "6K borrow · 15K save · 12K pay", "34% already joined; small pool left"],
    ["p_wealth", "Aged 60+", "9.3M", 9.3, "0 borrow · 28K save · 7K pay", "Savings-led offer; no loans (age limits)"],
  ];
  const rh = 0.9, barMax = RW - 2.8;
  inf.forEach((r, k) => {
    const y = top + 0.7 + k * rh;
    if (k) s.addShape(pres.shapes.LINE, { x: RX + 0.15, y: y - 0.04, w: RW - 0.3, h: 0, line: { color: "D9DFE6", width: 0.75 }, objectName: `Infer rule ${k}` });
    s.addShape(pres.shapes.OVAL, { x: RX + 0.15, y: y + 0.05, w: 0.34, h: 0.34, fill: { color: k === 2 ? NAVY : CYAN }, line: { type: "none" }, objectName: `Infer icon bg ${k}` });
    s.addImage({ path: IC(r[0] + ".png"), x: RX + 0.23, y: y + 0.13, w: 0.18, h: 0.18, altText: r[1] + " icon", objectName: `Infer icon ${k}` });
    s.addText(r[1], { x: RX + 0.58, y: y + 0.02, w: 1.45, h: 0.4, fontSize: 10.5, bold: true, color: NAVY, margin: 0, valign: "middle", isTextBox: true, objectName: `Infer seg ${k}` });
    s.addShape(pres.shapes.RECTANGLE, { x: RX + 2.05, y: y + 0.13, w: Math.max(0.05, barMax * r[3] / 9.3), h: 0.17, fill: { color: k === 2 ? GREY : BLUE }, line: { type: "none" }, objectName: `Infer bar ${k}` });
    s.addText(r[2], { x: RX + 2.05 + Math.max(0.05, barMax * r[3] / 9.3) + 0.05, y: y + 0.04, w: 0.6, h: 0.36, fontSize: 11, bold: true, color: NAVY, margin: 0, valign: "middle", isTextBox: true, objectName: `Infer pool ${k}` });
    s.addText([{ text: r[4], options: { fontSize: 8.5, color: TXT, breakLine: true } }, { text: "→ " + r[5], options: { fontSize: 8.5, bold: true, color: CYAN } }],
      { x: RX + 0.58, y: y + 0.42, w: RW - 0.73, h: 0.44, margin: 0, valign: "top", isTextBox: true, objectName: `Infer mix ${k}` });
  });
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
