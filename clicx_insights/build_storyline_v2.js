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
pres.title = "CLICX board storyline v2";
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
pres.addSection({ title: "Storyline" });
// =====================================================================
// Storyline v2 — built from "DRAFT 20261105 - Data for Strategy 2027"
const ORANGE = "E3A15F", TRUEC = "D9822B", SOFT = "8FB3DA";
const SRC2 = "Source: Data for Business Strategy (draft, 5 Nov 2026). CLICX data as of Sep–Oct 2026; AIS, KTB, OR as of Apr 2026";
const card = (s, x, y, w, h, fill, obj) => s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.08, fill: { color: fill }, line: { type: "none" }, objectName: obj });
const txt = (s, t, o) => s.addText(t, { margin: 0, isTextBox: true, ...o });
const badge = (s, x, y, n, col = NAVY, d = 0.36) => {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: col }, line: { type: "none" }, objectName: `Badge ${n}` });
  txt(s, String(n), { x, y, w: d, h: d, fontSize: 12, bold: true, color: "FFFFFF", align: "center", valign: "middle", objectName: `Badge num ${n}` });
};
const circIcon = (s, x, y, d, icon, col, obj) => {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: col }, line: { type: "none" }, objectName: obj + " bg" });
  s.addImage({ path: IC(icon), x: x + d * 0.25, y: y + d * 0.25, w: d * 0.5, h: d * 0.5, altText: obj, objectName: obj });
};
const soBar = (s, y, t, x = 0.5, w = 12.33) => {
  card(s, x, y, w, 0.44, "FFFFFF", "So box");
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h: 0.44, rectRadius: 0.08, fill: { type: "none" }, line: { color: CYAN, width: 1 }, objectName: "So outline" });
  txt(s, [{ text: "SO  ", options: { bold: true, color: CYAN, fontSize: 10, charSpacing: 1 } }, { text: t, options: { color: TXT, fontSize: 11.5 } }],
    { x: x + 0.18, y, w: w - 0.36, h: 0.44, valign: "middle", objectName: "So text" });
};
// drawn horizontal bars; rows = [label, value, valueText, color]
function hbars(s, x, y, w, rows, o = {}) {
  const lw = o.lw || 1.8, vw = o.vw || 0.6, rh = o.rh || 0.42, max = o.max || 1, bw = w - lw - vw - 0.08, fs = o.fs || 10.5;
  rows.forEach((r, k) => {
    const yy = y + k * rh;
    txt(s, r[0], { x, y: yy, w: lw - 0.08, h: rh, fontSize: fs, color: TXT, valign: "middle", objectName: `${o.name || "Bar"} label ${k}` });
    s.addShape(pres.shapes.RECTANGLE, { x: x + lw, y: yy + rh * 0.2, w: Math.max(0.03, bw * r[1] / max), h: rh * 0.6, fill: { color: r[3] }, line: { type: "none" }, objectName: `${o.name || "Bar"} ${k}` });
    txt(s, r[2], { x: x + lw + bw * r[1] / max + 0.06, y: yy, w: vw + 0.4, h: rh, fontSize: fs, bold: true, color: NAVY, valign: "middle", objectName: `${o.name || "Bar"} value ${k}` });
  });
  if (o.bench) {
    const bx = x + lw + bw * o.bench.v / max, by = y - 0.05, bh = rows.length * rh + 0.1;
    s.addShape(pres.shapes.LINE, { x: bx, y: by, w: 0, h: bh, line: { color: MUTED, width: 1.25, dashType: "dash" }, objectName: `${o.name || "Bar"} bench` });
    txt(s, o.bench.label, { x: bx - 1.0, y: by + bh + 0.02, w: 2.0, h: 0.2, fontSize: 9, color: MUTED, align: "center", objectName: `${o.name || "Bar"} bench label` });
  }
}
const panelHead = (s, x, y, w, n, head, sub, col = NAVY) => {
  if (n) badge(s, x + 0.18, y + 0.16, n, col);
  const hx = n ? x + 0.66 : x + 0.2;
  txt(s, head, { x: hx, y: y + 0.12, w: w - (hx - x) - 0.15, h: 0.46, fontSize: 13.5, bold: true, color: NAVY, valign: "middle", objectName: `Panel head ${head.slice(0, 12)}` });
  if (sub) txt(s, sub, { x: x + 0.2, y: y + 0.62, w: w - 0.4, h: 0.24, fontSize: 9.5, color: MUTED, objectName: `Panel sub ${head.slice(0, 12)}` });
};

// ---------- Cover
{
  const s = pres.addSlide({ masterName: "TITLE" });
  s.addText("DATA FOR STRATEGY 2027  ·  BOARD STORYLINE V2", { placeholder: "kicker" });
  s.addText("From 2.1M sign-ups to an active, growing and defended customer base", { placeholder: "title", align: "left" });
  s.addText("Consortium penetration · persona and portfolio · activating inactive customers · low KTB cannibalization · defending against True", { placeholder: "sub" });
  s.addText("November 2026", { placeholder: "date" });
}

// ---------- 1 Executive summary
{
  const s = pres.addSlide({ masterName: "QUESTION", sectionTitle: "Storyline" });
  s.addText("1 / 10  |  EXECUTIVE SUMMARY", { placeholder: "tracker" });
  s.addText("We built a 2.1M base through our partners; next we activate it, launch the next products and defend it against True", { placeholder: "title" });
  s.addText(SRC2, { placeholder: "foot" });
  card(s, 0.5, 1.55, 12.33, 0.42, NAVY, "Governing question");
  txt(s, [{ text: "THE QUESTION  ", options: { bold: true, fontSize: 10, color: LCYAN, charSpacing: 1 } },
    { text: "How do we turn 2.1M sign-ups into an active, profitable and defended customer base?", options: { bold: true, fontSize: 14, color: "FFFFFF" } }],
    { x: 0.7, y: 1.55, w: 12.0, h: 0.42, valign: "middle", objectName: "Governing question text" });
  const cx = [1.1, 3.35, 8.15, 9.85];
  ["AREA", "WHAT WE FOUND", "KEY NUMBER", "WHAT WE WILL DO"].forEach((h, k) =>
    txt(s, h, { x: cx[k], y: 2.07, w: 2.5, h: 0.24, fontSize: 9, bold: true, color: MUTED, charSpacing: 1, objectName: `Col head ${k}` }));
  const rows = [
    ["Consortium penetration", "96% of customers come from AIS, KTB or OR, but we have reached only 4–7% of each partner's base", "96%", "from partners", "Keep acquiring through partners, aimed at tech-savvy and Gen Z look-alikes"],
    ["Persona and portfolio", "We hit our targets (82% underserved; tech-savvy join most), but 72% came for cash and only 7% got a loan", "7%", "got a loan", "Credit card for 1.1M, investing for 528K, SME loans for 106K merchants"],
    ["Activate inactive customers", "Only 1 in 5 really uses CLICX; partner deals start payments, but 19% of partner payers stopped after August", "309K", "active of 1.51M", "Lead with AIS prepaid top-up and Amazon offers; sell savings to borrowers"],
    ["Low KTB cannibalization", "64% of shared customers were inactive at KTB, and 450K of KTB's loan rejects applied with us", "1.1M", "KTB-inactive", "Make the Paotang reject list a standing channel: under-40s and riders first"],
    ["Defend against True", "1.26M True and TrueMoney customers hold ฿1.98B, 63% of our deposits", "฿1.98B", "at risk", "Reach the 313K TrueMoney users KTB turned down before True does"],
  ];
  rows.forEach((r, k) => {
    const y = 2.36 + k * 0.86, h = 0.8;
    card(s, 0.5, y, 12.33, h, C.background2, `Row ${k}`);
    badge(s, 0.62, y + 0.22, k + 1, k === 4 ? TRUEC : NAVY);
    txt(s, r[0], { x: cx[0], y, w: 2.15, h, fontSize: 12.5, bold: true, color: NAVY, valign: "middle", objectName: `Area ${k}` });
    txt(s, r[1], { x: cx[1], y, w: 4.6, h, fontSize: 11, color: TXT, valign: "middle", objectName: `Found ${k}` });
    txt(s, [{ text: r[2], options: { bold: true, fontSize: 20, color: k === 4 ? TRUEC : BLUE, breakLine: true } }, { text: r[3], options: { fontSize: 9, color: MUTED } }],
      { x: cx[2], y, w: 1.55, h, valign: "middle", objectName: `Key ${k}` });
    s.addShape(pres.shapes.RECTANGLE, { x: cx[3] - 0.15, y: y + 0.14, w: 0.05, h: h - 0.28, fill: { color: CYAN }, line: { type: "none" }, objectName: `Do bar ${k}` });
    txt(s, r[4], { x: cx[3], y, w: 2.85, h, fontSize: 10.5, bold: true, color: NAVY, valign: "middle", objectName: `Do ${k}` });
  });
}

// ---------- 2 Consortium penetration
{
  const s = pres.addSlide({ masterName: "QUESTION", sectionTitle: "Storyline" });
  s.addText("2 / 10  |  CONSORTIUM PENETRATION", { placeholder: "tracker" });
  s.addText("96% of CLICX customers come from AIS, KTB or OR, yet we have reached only 4–7% of each partner's base", { placeholder: "title" });
  s.addText("Join rate = CLICX customers ÷ the group's size in the 49.5M consortium base. Tech-savvy = uses Make, Kept, Dime or crypto. CLICX data as of Sep 2026; AIS, KTB, OR as of Apr 2026. Source: Data for Business Strategy (draft)", { placeholder: "foot" });
  const y = 1.6, H = 5.15, LW = 6.0, RX = 6.65, RW = 12.83 - 6.65;
  card(s, 0.5, y, LW, H, C.background2, "Left panel");
  panelHead(s, 0.5, y, LW, 1, "Partners bring nearly every customer", "CLICX customers, Sep 2026");
  const fb = [["2.1M", "total CLICX customers", 1, NAVY], ["1.86M · 86%", "gave partner consent", 0.86, BLUE], ["1.8M · 85%", "are AIS, KTB or OR customers", 0.85, CYAN]];
  fb.forEach((f, k) => {
    const yy = y + 0.98 + k * 0.48, ww = (LW - 0.4) * f[2];
    s.addShape(pres.shapes.RECTANGLE, { x: 0.7, y: yy, w: ww, h: 0.4, fill: { color: f[3] }, line: { type: "none" }, objectName: `Funnel ${k}` });
    txt(s, [{ text: f[0] + "  ", options: { bold: true } }, { text: f[1] }], { x: 0.82, y: yy, w: ww - 0.2, h: 0.4, fontSize: 11, color: "FFFFFF", valign: "middle", objectName: `Funnel text ${k}` });
  });
  txt(s, "96% of customers who gave consent are AIS, KTB or OR customers", { x: 0.7, y: y + 2.42, w: LW - 0.4, h: 0.26, fontSize: 10.5, bold: true, color: NAVY, objectName: "Funnel note" });
  const pc = [["logo_ais.png", 2.74, "1.2M", "67% of consented customers", "3.7%", "of AIS's base has joined"],
    ["logo_ktb.jpeg", 2.58, "1.6M", "88% of consented customers", "4.7%", "of KTB's base has joined"],
    ["logo_or.png", 1.93, "725K", "40% of consented customers", "7.4%", "of OR's base has joined, the best rate"]];
  const pw = (LW - 0.4 - 0.2) / 3;
  pc.forEach((p, k) => {
    const x = 0.7 + k * (pw + 0.1), yy = y + 2.82;
    card(s, x, yy, pw, 2.15, "FFFFFF", `Partner card ${k}`);
    s.addImage({ path: IC(p[0]), x: x + 0.15, y: yy + 0.15, h: 0.28, w: 0.28 * p[1], altText: "Partner logo", objectName: `Partner logo ${k}` });
    txt(s, [{ text: p[2], options: { bold: true, fontSize: 20, color: NAVY, breakLine: true } }, { text: p[3], options: { fontSize: 9, color: MUTED } }], { x: x + 0.15, y: yy + 0.5, w: pw - 0.3, h: 0.72, valign: "top", objectName: `Partner size ${k}` });
    s.addShape(pres.shapes.LINE, { x: x + 0.15, y: yy + 1.25, w: pw - 0.3, h: 0, line: { color: LGREY, width: 1 }, objectName: `Partner rule ${k}` });
    txt(s, [{ text: p[4], options: { bold: true, fontSize: 18, color: CYAN, breakLine: true } }, { text: p[5], options: { fontSize: 9, color: TXT } }], { x: x + 0.15, y: yy + 1.32, w: pw - 0.3, h: 0.75, valign: "top", objectName: `Partner rate ${k}` });
  });
  card(s, RX, y, RW, H, C.background2, "Right panel");
  panelHead(s, RX, y, RW, 2, "Tech-savvy and young working customers join fastest", "% of each consortium group now on CLICX (customers)");
  hbars(s, RX + 0.2, y + 1.0, RW - 0.4, [
    ["Tech-savvy (533K)", 8.5, "8.5%", CYAN], ["Working Gen Z, 25–29 (145K)", 5.6, "5.6%", CYAN], ["Working, 30–59 (580K)", 4.1, "4.1%", BLUE],
    ["KTB Next only (193K)", 3.9, "3.9%", BLUE], ["Multi-bank (194K)", 2.4, "2.4%", GREY], ["Aged 60+ (15K)", 0.2, "0.2%", GREY]],
    { lw: 2.4, rh: 0.44, max: 10, name: "Join", bench: { v: 3.8, label: "All groups 3.8%" } });
  card(s, RX + 0.2, y + 4.0, RW - 0.4, 0.95, "FFFFFF", "Join so");
  txt(s, [{ text: "SO  ", options: { bold: true, color: CYAN, fontSize: 10, charSpacing: 1 } }, { text: "Partners still bring almost everyone. Aim partner campaigns at tech-savvy and young working customers, who join at up to twice the average rate", options: { fontSize: 11, color: TXT } }],
    { x: RX + 0.35, y: y + 4.0, w: RW - 0.7, h: 0.95, valign: "middle", objectName: "Join so text" });
}

// ---------- 3 Persona vs target
{
  const s = pres.addSlide({ masterName: "QUESTION", sectionTitle: "Storyline" });
  s.addText("3 / 10  |  PERSONA: WHO WE REACHED", { placeholder: "tracker" });
  s.addText("We reached the customers we aimed for: young digital early adopters and the underserved, mostly outside Bangkok", { placeholder: "title" });
  s.addText("Join rate = CLICX customers ÷ the group's size in the consortium base. Underserved = NCB classification of loan applicants, 27 Oct 2026. 40% = tech-savvy (29%) + multi-bank (11%). Source: Data for Business Strategy (draft)", { placeholder: "foot" });
  const chips = [["2.1M", "customers"], ["71%", "aged 25–44"], ["70%", "live outside Bangkok, mostly small Isan provinces"], ["40%", "already bank digitally"]];
  const cw = (12.33 - 0.45) / 4;
  chips.forEach((c, k) => {
    const x = 0.5 + k * (cw + 0.15);
    card(s, x, 1.58, cw, 0.7, STAT_BG, `Chip ${k}`);
    txt(s, [{ text: c[0] + "  ", options: { bold: true, fontSize: 20, color: BLUE } }, { text: c[1], options: { fontSize: 10.5, color: TXT } }], { x: x + 0.15, y: 1.58, w: cw - 0.3, h: 0.7, valign: "middle", objectName: `Chip text ${k}` });
  });
  const y = 2.42, H = 3.85, PW = (12.33 - 0.15) / 2;
  // panel 1: age
  card(s, 0.5, y, PW, H, C.background2, "Age panel");
  panelHead(s, 0.5, y, PW, 1, "Early adopters: 25–39s join most", "% of each age group in partner bases now on CLICX");
  const ages = ["<20", "20–24", "25–29", "30–34", "35–39", "40–44", "45–49", "50–59", "60+"], av = [1, 5, 8, 9, 8, 6, 4, 2, 1];
  const gx = 0.75, gw = PW - 0.5, gy = y + 1.15, gh = 1.9, colw = gw / ages.length;
  ages.forEach((a, k) => {
    const bh = gh * av[k] / 10, hi = av[k] >= 8;
    s.addShape(pres.shapes.RECTANGLE, { x: gx + k * colw + colw * 0.18, y: gy + gh - bh, w: colw * 0.64, h: bh, fill: { color: hi ? CYAN : SOFT }, line: { type: "none" }, objectName: `Age bar ${k}` });
    txt(s, av[k] + "%", { x: gx + k * colw, y: gy + gh - bh - 0.26, w: colw, h: 0.24, fontSize: 10.5, bold: hi, color: NAVY, align: "center", objectName: `Age value ${k}` });
    txt(s, a, { x: gx + k * colw, y: gy + gh + 0.04, w: colw, h: 0.24, fontSize: 10, color: MUTED, align: "center", objectName: `Age label ${k}` });
  });
  txt(s, "Tech-savvy partner customers join at 8.5%, more than twice the 3.8% average", { x: 0.7, y: y + H - 0.5, w: PW - 0.4, h: 0.34, fontSize: 10.5, italic: true, color: MUTED, valign: "middle", objectName: "Age note" });
  // panel 2: underserved
  const X2 = 0.5 + PW + 0.15;
  card(s, X2, y, PW, H, C.background2, "Underserved panel");
  panelHead(s, X2, y, PW, 2, "Underserved: gig workers join most", "% of each job group in partner bases now on CLICX");
  card(s, X2 + 0.2, y + 1.0, 1.55, 2.3, STAT_BG, "Underserved stat");
  txt(s, [{ text: "82%", options: { bold: true, fontSize: 26, color: BLUE, breakLine: true } }, { text: "of loan applicants are underserved (1.45M, NCB)", options: { fontSize: 9.5, color: TXT } }], { x: X2 + 0.3, y: y + 1.1, w: 1.35, h: 2.1, valign: "top", objectName: "Underserved stat text" });
  hbars(s, X2 + 1.95, y + 1.0, PW - 2.15, [["Freelancer (232K)", 40, "40%", CYAN], ["Rider, driver (66K)", 22, "22%", CYAN], ["Blue collar (401K)", 16, "16%", BLUE], ["White collar (700K)", 9, "9%", BLUE], ["Merchant (293K)", 9, "9%", BLUE]],
    { lw: 1.55, rh: 0.42, max: 45, fs: 10, name: "Job", bench: { v: 3.8, label: "All 3.8%" } });
  txt(s, "White and blue collar are 60% of customers (1.1M); freelancers and riders fit best", { x: X2 + 0.2, y: y + H - 0.5, w: PW - 0.4, h: 0.34, fontSize: 10.5, italic: true, color: MUTED, valign: "middle", objectName: "Job note" });
  soBar(s, 6.37, "Products must work for people without payslips or bank statements: alternative-data approval is core to CLICX, not a niche");
}

// ---------- Slide 2: what customers hold, product segments sized by customers, in-app payers outlined
{
  const s = pres.addSlide({ masterName: "QUESTION", sectionTitle: "Storyline" });
  s.addText("4 / 10  |  PORTFOLIO: WHAT THEY HOLD", { placeholder: "tracker" });
  s.addText("Only 1 in 5 customers really uses CLICX; 127K of them pay in-app, spread across every segment", { placeholder: "title" });
  s.addText("Box size = customers · orange outline = paid in-app in Oct · active = balance over ฿100, payment or loan · Bangkok area = Bangkok, Nonthaburi, Pathum Thani, Samut Prakan · saver on balance only: core over ฿20K, emerging ฿101–20K (Save Max/More flag not used) · 31 Oct 2026", { placeholder: "foot" });
  const ORANGE = "E3A15F", WHITE = "FFFFFF";
  const TOTAL = 1507.4, ZERO = 1061.3, MONEY = 136.7, ACT = 309.4;
  const X0 = 0.5, W = 8.95;
  // whole base strip
  const sy = 1.62, sh = 0.46, zw = W * ZERO / TOTAL, mw = W * MONEY / TOTAL, aw = W - zw - mw;
  s.addText([{ text: "No product, ฿0 balance ", options: { bold: true } }, { text: "1.06M · 70%" }], { x: X0, y: sy, w: zw, h: sh, fontSize: 11, color: WHITE, fill: { color: GREY }, line: { color: WHITE, width: 1.5 }, margin: 7, valign: "middle", isTextBox: true, objectName: "Strip zero" });
  s.addText([{ text: "137K · 9%", options: { bold: true, fontSize: 9, breakLine: true } }, { text: "bal >฿0 & ≤฿100", options: { fontSize: 7.5 } }], { x: X0 + zw, y: sy, w: mw, h: sh, fontSize: 9, color: TXT, align: "center", fill: { color: "C8CFD6" }, line: { color: WHITE, width: 1.5 }, margin: 0, valign: "middle", isTextBox: true, objectName: "Strip money" });
  s.addText([{ text: "309K ", options: { bold: true } }, { text: "20.5%" }], { x: X0 + zw + mw, y: sy, w: aw, h: sh, fontSize: 10.5, color: WHITE, align: "center", fill: { color: "3A8A80" }, line: { color: WHITE, width: 1.5 }, margin: 0, valign: "middle", isTextBox: true, objectName: "Strip active" });
  // treemap geometry
  const ty = 2.4, th = 3.62;
  const fam = [["Borrowers", 165.161, NAVY, " · 70% up-country"], ["Savers", 104.733, CYAN, " · 47% Bangkok area"], ["Payers only", 39.492, MUTED]];
  const cwF = fam.map((f) => W * f[1] / ACT);
  const cx = [X0, X0 + cwF[0], X0 + cwF[0] + cwF[1]];
  fam.forEach((f, a) => {
    const lab = a === 2 ? [{ text: f[0], options: { bold: true } }] : [{ text: f[0], options: { bold: true } }, { text: "  " + Math.round(f[1]) + "K · " + Math.round(100 * f[1] / ACT) + "%" + (f[3] || "") }];
    s.addText(lab, { x: cx[a], y: ty - 0.27, w: cwF[a], h: 0.24, fontSize: 10, color: f[2], margin: [0, 0, 0, 0.02], valign: "bottom", isTextBox: true, objectName: `Family ${a}` });
  });
  const box = (x, y, w, h, fill, txtCol, runs, pay, obj, opts = {}) =>
    s.addText(runs, { x, y, w, h, color: txtCol, fill: { color: fill }, line: { color: pay ? ORANGE : WHITE, width: pay ? 2.25 : 2 }, margin: opts.margin ?? 8, valign: opts.valign || "top", align: opts.align || "left", isTextBox: true, objectName: obj });
  const H = (v, f) => th * v / f;
  // borrowers column
  { const x = cx[0], w = cwF[0], f = 165.161; let y = ty;
    const h1 = H(102.665, f), h2 = H(40.526, f), h3 = H(21.97, f);
    box(x, y, w, h1, NAVY, WHITE, [{ text: "Borrower only", options: { bold: true, fontSize: 14, breakLine: true } }, { text: "103K", options: { bold: true, fontSize: 30 } }, { text: " 33%", options: { fontSize: 11, breakLine: true } }, { text: "Mostly salaried, aged 20–39; 6 in 10 keep ฿0", options: { fontSize: 11 } }], false, "Tree Borrower only");
    y += h1;
    box(x, y, w, h2, "3A5A85", WHITE, [{ text: "Borrower + transactor", options: { bold: true, fontSize: 12, breakLine: true } }, { text: "41K", options: { bold: true, fontSize: 18 } }, { text: " 13% · salaried, 30–39", options: { fontSize: 10.5 } }], true, "Tree Borrower + transactor", { valign: "middle" });
    y += h2;
    const w1 = w * 11.326 / 21.97, w2 = w - w1;
    box(x, y, w1, h3, "5374A8", WHITE, [{ text: "11.3K · 4%", options: { bold: true, fontSize: 12 } }], false, "Tree Saver + borrower", { align: "center", valign: "middle", margin: 2 });
    box(x + w1, y, w2, h3, "5374A8", WHITE, [{ text: "10.6K · 3%", options: { bold: true, fontSize: 12 } }], true, "Tree Saver + borrower who pay", { align: "center", valign: "middle", margin: 2 });
    var callB = [[x + w1 / 2, "Saver + borrower", "Salaried, aged 30–49"], [x + w1 + w2 / 2, "Saver + borrower who pay", "Salaried, aged 30–49"]];
  }
  // savers column
  { const x = cx[1], w = cwF[1], f = 104.733; let y = ty;
    const h1 = H(47.265, f), h2 = H(30.043, f), h3 = H(27.425, f);
    box(x, y, w, h1, "4BA394", WHITE, [{ text: "Emerging saver only", options: { bold: true, fontSize: 14, breakLine: true } }, { text: "47K", options: { bold: true, fontSize: 30 } }, { text: " 15%", options: { fontSize: 11, breakLine: true } }, { text: "Salaried and merchants, 40–59; 44% in Bangkok area", options: { fontSize: 11 } }], false, "Tree Emerging saver only");
    y += h1;
    box(x, y, w, h2, "66B3A6", WHITE, [{ text: "Emerging saver + transactor", options: { bold: true, fontSize: 12, breakLine: true } }, { text: "30K", options: { bold: true, fontSize: 18 } }, { text: " 10%", options: { fontSize: 10.5, breakLine: true } }, { text: "Salaried and merchants, 30–49", options: { fontSize: 10.5 } }], true, "Tree Emerging saver + transactor");
    y += h2;
    const w1 = w * 21.282 / 27.425, w2 = w - w1;
    box(x, y, w1, h3, "9FCFC7", TXT, [{ text: "Core saver", options: { bold: true, fontSize: 12, breakLine: true } }, { text: "21.3K", options: { bold: true, fontSize: 18 } }, { text: " 7%", options: { fontSize: 10.5, breakLine: true } }, { text: "Aged 40+; half live in Bangkok", options: { fontSize: 10 } }], false, "Tree Core saver", { valign: "middle" });
    box(x + w1, y, w2, h3, "9FCFC7", TXT, [{ text: "6.1K", options: { bold: true, fontSize: 11.5, breakLine: true } }, { text: "2%", options: { fontSize: 10 } }], true, "Tree Core savers who pay", { align: "center", valign: "middle", margin: 2 });
    var callS = [[x + w1 + w2 / 2, "Core savers who pay", "Salaried, aged 30–49"]];
  }
  // payers-only column
  box(cx[2], ty, cwF[2], th, "8395AB", WHITE, [{ text: "Pays only", options: { bold: true, fontSize: 12, breakLine: true } }, { text: "39K", options: { bold: true, fontSize: 18 } }, { text: " 13%", options: { fontSize: 10.5, breakLine: true } }, { text: "Wage, gig and online sellers", options: { fontSize: 10.5 } }], true, "Tree Pays only");
  // callouts under the split boxes
  const cy = ty + th;
  [...callB, ...callS].forEach((c, k) => {
    s.addShape(pres.shapes.LINE, { x: c[0], y: cy, w: 0, h: 0.09, line: { color: TXT, width: 1 }, objectName: `Callout tick ${k}` });
    const lw = 2.2, lx = k === 2 ? c[0] - lw + 0.05 : c[0] - lw / 2;
    s.addText([{ text: c[1], options: { bold: true, breakLine: true } }, { text: c[2], options: { color: MUTED } }], { x: lx, y: cy + 0.1, w: lw, h: 0.4, fontSize: 9.5, color: TXT, align: k === 2 ? "right" : "center", margin: 0, valign: "top", isTextBox: true, objectName: `Callout ${k}` });
  });
  // right panel
  const RX = 9.7, RW = 3.13;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: RX, y: sy, w: RW, h: 0.9, rectRadius: 0.08, fill: { color: STAT_BG }, line: { type: "none" }, objectName: "Active stat" });
  s.addText([{ text: "309K", options: { bold: true, fontSize: 28, color: NAVY, breakLine: true } }, { text: "active: 20% of 1.51M", options: { fontSize: 11, color: TXT } }],
    { x: RX + 0.15, y: sy + 0.04, w: RW - 0.3, h: 0.82, margin: 0, valign: "middle", isTextBox: true, objectName: "Active stat text" });
  const tk = [["127K pay in-app (41%)", "88K also borrow or save", ORANGE], ["165K borrowers (53%)", "6 in 10 keep ฿0: sell savings", NAVY],
    ["105K savers (34%)", "1 in 5 core savers pays in-app", CYAN], ["137K hold ฿1–100 only", "No loan, no payment", "C8CFD6"], ["Saving is a Bangkok story", "46% of core savers live in Bangkok, vs 18% of customers", "0F6D64"]];
  tk.forEach((t, k) => {
    const y = 2.72 + k * 0.7;
    s.addShape(pres.shapes.RECTANGLE, { x: RX, y, w: 0.05, h: 0.56, fill: { color: t[2] }, line: { type: "none" }, objectName: `Takeaway bar ${k}` });
    s.addText([{ text: t[0], options: { bold: true, fontSize: 12, color: NAVY, breakLine: true } }, { text: t[1], options: { fontSize: 10.5, color: TXT } }],
      { x: RX + 0.17, y, w: RW - 0.2, h: 0.56, margin: 0, valign: "middle", isTextBox: true, objectName: `Takeaway ${k}` });
  });
}



// ---------- 5 Loans drive activity
{
  const s = pres.addSlide({ masterName: "QUESTION", sectionTitle: "Storyline" });
  s.addText("5 / 10  |  PORTFOLIO: LOANS AND ACTIVITY", { placeholder: "tracker" });
  s.addText("Loans bring customers in and keep them active: 72% came for cash, only 7% got a loan, and those turned down go quiet", { placeholder: "title" });
  s.addText("Active = any payment, transfer in, ATM use, pocket, coupon, loan booking or repayment, or mission in Sep 2026. Money in = balance over ฿0. Applied = ever applied for a loan in the app. Source: Data for Business Strategy (draft), Sep 2026", { placeholder: "foot" });
  const y = 1.6, H = 4.65, PW = (12.33 - 0.15) / 2;
  card(s, 0.5, y, PW, H, C.background2, "Loan panel");
  panelHead(s, 0.5, y, PW, 1, "Almost everyone applies; few get a loan", "% of each job group on CLICX");
  [["Applied for a loan", SOFT], ["Got a loan", NAVY]].forEach((l, k) => {
    s.addShape(pres.shapes.RECTANGLE, { x: 0.72 + k * 1.9, y: y + 0.98, w: 0.16, h: 0.16, fill: { color: l[1] }, line: { type: "none" }, objectName: `Legend key ${k}` });
    txt(s, l[0], { x: 0.94 + k * 1.9, y: y + 0.93, w: 1.7, h: 0.26, fontSize: 10, color: MUTED, valign: "middle", objectName: `Legend ${k}` });
  });
  const lj = [["White collar", 85, 11], ["Merchant", 84, 7], ["Blue collar", 92, 7], ["Freelancer", 88, 6], ["Rider, driver", 91, 5]];
  const lx = 0.72, lw = 1.35, bw = PW - 0.44 - lw - 0.6;
  lj.forEach((r, k) => {
    const yy = y + 1.32 + k * 0.56;
    txt(s, r[0], { x: lx, y: yy, w: lw, h: 0.5, fontSize: 10.5, color: TXT, valign: "middle", objectName: `Job ${k}` });
    s.addShape(pres.shapes.RECTANGLE, { x: lx + lw, y: yy + 0.04, w: bw * r[1] / 100, h: 0.2, fill: { color: SOFT }, line: { type: "none" }, objectName: `Applied ${k}` });
    txt(s, r[1] + "%", { x: lx + lw + bw * r[1] / 100 + 0.05, y: yy, w: 0.6, h: 0.28, fontSize: 9.5, color: MUTED, valign: "middle", objectName: `Applied value ${k}` });
    s.addShape(pres.shapes.RECTANGLE, { x: lx + lw, y: yy + 0.26, w: bw * r[2] / 100, h: 0.2, fill: { color: NAVY }, line: { type: "none" }, objectName: `Booked ${k}` });
    txt(s, r[2] + "%", { x: lx + lw + bw * r[2] / 100 + 0.05, y: yy + 0.22, w: 0.6, h: 0.28, fontSize: 10, bold: true, color: NAVY, valign: "middle", objectName: `Booked value ${k}` });
  });
  txt(s, "Office workers get a loan most often (11%); riders least (5%)", { x: 0.7, y: y + H - 0.5, w: PW - 0.4, h: 0.34, fontSize: 10.5, italic: true, color: MUTED, valign: "middle", objectName: "Loan note" });
  const X2 = 0.5 + PW + 0.15;
  card(s, X2, y, PW, H, C.background2, "Active panel");
  panelHead(s, X2, y, PW, 2, "Customers stay active only if they borrow or keep money with us", "% active in Sep 2026, by what customers hold");
  hbars(s, X2 + 0.2, y + 1.0, PW - 0.4, [["Borrower (158K)", 88, "88%", CYAN], ["Money in, applied (161K)", 50, "50%", BLUE], ["Money in, never applied (92K)", 52, "52%", BLUE],
    ["No money in, never applied (113K)", 7, "7%", GREY], ["No money in, applied (924K)", 4, "4%", GREY]], { lw: 2.55, rh: 0.42, max: 100, fs: 10, name: "Active" });
  card(s, X2 + 0.2, y + 3.25, PW - 0.4, 1.2, "FFFFFF", "Turned down box");
  txt(s, [{ text: "924K", options: { bold: true, fontSize: 22, color: NAVY } }, { text: "  applied, didn't get a loan and keep no money with us: 96% of them are inactive. Another 694K hold no product at all", options: { fontSize: 11, color: TXT } }],
    { x: X2 + 0.35, y: y + 3.25, w: PW - 0.7, h: 1.2, valign: "middle", objectName: "Turned down text" });
  soBar(s, 6.37, "Approve more, safely (alternative data, smaller first loans), and give turned-down applicants a reason to stay: savings and payments");
}

// ---------- 6 Next products
{
  const s = pres.addSlide({ masterName: "QUESTION", sectionTitle: "Storyline" });
  s.addText("6 / 10  |  PORTFOLIO: NEXT PRODUCTS", { placeholder: "tracker" });
  s.addText("Customers already point to our next products: a credit card for 1.1M, investing for 528K and SME loans for 106K merchants", { placeholder: "title" });
  s.addText("Pools overlap. Tech-savvy = uses Make, Kept, Dime or crypto (AIS and KTB data). Income from CLICX profile. Source: Data for Business Strategy (draft), Sep 2026", { placeholder: "foot" });
  const cards = [
    ["card.png", NAVY, "Digital credit card", "1.1M", "earn over ฿15K a month, aged 20–44: 482K salaried, 637K without a steady income",
      ["57% have no steady income", "85% have a high need for cash", "41% (489K) pay for streaming", "66% are AIS and 38% OR customers"],
      "Alternative-data approval · 0% micro-instalments for bills, PTT fuel and daily spend · streaming subscription perks"],
    ["invest.png", CYAN, "Investing and wealth", "528K", "tech-savvy customers: 63K under 24, 379K aged 25–44, 86K aged 45–59",
      ["110K already hold over ฿20K or Save Max/More", "71–73% are AIS, KTB or OR customers", "Under-24s: high cash need, low wealth, pay more"],
      "Under 24: small-ticket, low-risk investing, learn-and-earn · 25–44: zero-fee fractional shares in ADVANC, OR, KTB, GULF · 45–59: pre-retirement plans with insurance"],
    ["sme.png", BLUE, "SME loan for merchants", "288K", "merchants: 102K online, 186K offline; half earn over ฿30K a month",
      ["222K applied without getting a loan in all", "106K applied for a loan without getting one and earn over ฿30K: first wave", "More active: 44% open the app vs 36% of others"],
      "Loan based on sales and partner data · start in Isan (21%), Greater Bangkok (16%) and Bangkok (15%)"],
  ];
  const cw = (12.33 - 0.3) / 3, y = 1.6, H = 5.15;
  cards.forEach((c, k) => {
    const x = 0.5 + k * (cw + 0.15);
    card(s, x, y, cw, H, C.background2, `Product card ${k}`);
    circIcon(s, x + 0.2, y + 0.18, 0.5, c[0], c[1], `Product icon ${k}`);
    txt(s, c[2], { x: x + 0.82, y: y + 0.18, w: cw - 1.0, h: 0.5, fontSize: 15, bold: true, color: NAVY, valign: "middle", objectName: `Product name ${k}` });
    card(s, x + 0.2, y + 0.85, cw - 0.4, 1.05, STAT_BG, `Focus box ${k}`);
    txt(s, [{ text: c[3], options: { bold: true, fontSize: 26, color: BLUE, breakLine: true } }, { text: c[4], options: { fontSize: 9.5, color: TXT } }], { x: x + 0.32, y: y + 0.88, w: cw - 0.64, h: 1.0, valign: "middle", objectName: `Focus ${k}` });
    txt(s, "WHAT WE KNOW", { x: x + 0.2, y: y + 2.02, w: 2, h: 0.22, fontSize: 9, bold: true, color: MUTED, charSpacing: 1, objectName: `Know head ${k}` });
    txt(s, c[5].map((t, i) => ({ text: t, options: { bullet: { indent: 12 }, breakLine: i < c[5].length - 1 } })), { x: x + 0.2, y: y + 2.26, w: cw - 0.4, h: 1.42, fontSize: 10.5, color: TXT, valign: "top", paraSpaceAfter: 3, objectName: `Know ${k}` });
    txt(s, "LEAD WITH", { x: x + 0.2, y: y + 3.72, w: 2, h: 0.22, fontSize: 9, bold: true, color: CYAN, charSpacing: 1, objectName: `Lead head ${k}` });
    card(s, x + 0.2, y + 3.96, cw - 0.4, 1.0, "FFFFFF", `Lead box ${k}`);
    txt(s, c[6], { x: x + 0.32, y: y + 3.96, w: cw - 0.64, h: 1.0, fontSize: 10, bold: true, color: NAVY, valign: "middle", objectName: `Lead ${k}` });
  });
}

// ---------- Slide 3: what we know, payments (3 insights)
{
  const s = pres.addSlide({ masterName: "QUESTION", sectionTitle: "Storyline" });
  s.addText("7 / 10  |  ACTIVATE INACTIVE CUSTOMERS", { placeholder: "tracker" });
  s.addText("3 in 4 payers come through AIS or OR; 19% of them stopped after the August deal, but those who stay pay 39% more", { placeholder: "title" });
  s.addText("Partner payers = ever paid AIS or PTT OR. 19% = drop in partner payers Aug → Oct (114.3K → 93.2K); 39% = spend per partner payer, ฿2.1K → ฿3.0K; +22K assumes prepaid pays at the postpaid rate. Points 1–2: Partner Payment Story (8.4% of 1.51M paid in Oct); point 3: Oct segment extract (7.8% of 1.50M); gig and wage = freelance, rider, wage earner", { placeholder: "foot" });
  const cw = (12.33 - 0.3) / 3, top = 1.5, ch = 5.28;
  const cols = [
    { head: "Partners start payments, but promotions don't keep them", big: "73%", lbl: "of October payers came through AIS or OR, making 67% of spend",
      chart: { type: "col", t: "Partner customers' share of payers and of spend", labels: ["Jun", "Jul", "Aug", "Sep", "Oct"], fmt: "0%", max: 0.9, colors: [NAVY, CYAN],
        series: [{ name: "Share of payers", values: [0.57, 0.67, 0.70, 0.73, 0.73] }, { name: "Share of spend", values: [0.51, 0.62, 0.66, 0.67, 0.67] }] },
      table: [["Partner payers", "12.8K", "31.6K", "114.3K", "95.0K", "93.2K"], ["Their spend", "฿20M", "฿88M", "฿244M", "฿262M", "฿275M"]],
      so: "Keep AIS and OR promotions as the way in; reward the second and third month" },
    { head: "AIS prepaid is the biggest payment gap", big: "516.5K", lbl: "AIS prepaid customers, only 8.4% pay",
      chart: { type: "bar", t: "% who paid in October", labels: ["AIS postpaid", "AIS prepaid", "PTT + Amazon", "Amazon only", "PTT only"], values: [0.128, 0.084, 0.120, 0.107, 0.083], fmt: "0.0%", max: 0.155, colors: [NAVY, CYAN, BLUE, "5A86BE", "8FB3DA"],
        layout: { x: 0.3, y: 0.17, w: 0.6, h: 0.68 }, bench: { v: 0.084, label: "All customers 8.4%" } },
      note: "The gap is mostly bill pay and top-up: 4.6% vs 7.8%",
      so: "Top-up-first offers for prepaid: up to +22K payers a month" },
    { head: "Teens, merchants and 30–49s pay most; blue-collar pay least", big: "13.5%", lbl: "of 10–19s paid in Oct; merchants 9.1–9.6%, vs 7.8% of all customers",
      chart: { type: "bar", t: "% who paid in October, by age and job", labels: ["Age 10–19", "Age 20–29", "Age 30–49", "Age 50–69", "Online merchant", "Offline merchant", "Salaried", "Gig and wage", "Blue collar"],
        values: [0.135, 0.070, 0.081, 0.076, 0.096, 0.091, 0.086, 0.068, 0.059], fmt: "0.0%", max: 0.165, fs: 9,
        colors: [CYAN, "B8C2CC", CYAN, "B8C2CC", CYAN, CYAN, CYAN, "B8C2CC", "B8C2CC"],
        layout: { x: 0.36, y: 0.15, w: 0.52, h: 0.8 } },
      note: "Teal = above the 7.8% all-customer rate; salaried and merchants aged 30–49 are where the volume is",
      so: "Aim payment offers at merchants, salaried 30–49s and students; give blue-collar and gig workers an easy first payment" },
  ];
  cols.forEach((c, i) => {
    const x = 0.5 + i * (cw + 0.15);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: top, w: cw, h: ch, rectRadius: 0.08, fill: { color: C.background2 }, line: { type: "none" }, objectName: `Pay col ${i}` });
    s.addShape(pres.shapes.OVAL, { x: x + 0.16, y: top + 0.16, w: 0.4, h: 0.4, fill: { color: C.text2 }, line: { type: "none" }, objectName: `Pay badge ${i}` });
    s.addText(String(i + 1), { x: x + 0.16, y: top + 0.16, w: 0.4, h: 0.4, fontSize: 12.5, bold: true, color: "FFFFFF", align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: `Pay num ${i}` });
    s.addText(c.head, { x: x + 0.66, y: top + 0.1, w: cw - 0.8, h: 0.58, fontSize: 12.5, bold: true, color: NAVY, margin: 0, valign: "middle", isTextBox: true, objectName: `Pay head ${i}` });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + 0.16, y: top + 0.8, w: cw - 0.32, h: 0.72, rectRadius: 0.06, fill: { color: STAT_BG }, line: { type: "none" }, objectName: `Pay stat box ${i}` });
    s.addText([{ text: c.big, options: { bold: true, fontSize: 20, color: BLUE, breakLine: true } }, { text: c.lbl, options: { fontSize: 10, color: TXT } }],
      { x: x + 0.28, y: top + 0.8, w: cw - 0.56, h: 0.72, valign: "middle", margin: 0, isTextBox: true, objectName: `Pay stat ${i}` });
    const ch2 = c.chart;
    const multi = !!ch2.series;
    if (ch2.bench) {
      const gx = x + 0.08, gw = cw - 0.16, gy = top + 1.62, gh = 2.35, L = ch2.layout;
      const bx = gx + gw * (L.x + L.w * ch2.bench.v / ch2.max);
      s.addShape(pres.shapes.LINE, { x: bx, y: gy + gh * L.y, w: 0, h: gh * L.h, line: { color: MUTED, width: 1.25, dashType: "dash" }, objectName: `Pay bench ${i}` });
      s.addText(ch2.bench.label, { x: bx - 0.8, y: gy + gh * (L.y + L.h) + 0.02, w: 1.6, h: 0.2, fontSize: 8.5, color: MUTED, align: "center", margin: 0, isTextBox: true, objectName: `Pay bench label ${i}` });
    }
    s.addChart(pres.charts.BAR, multi ? ch2.series.map((x2) => ({ name: x2.name, labels: ch2.labels, values: x2.values })) : [{ name: ch2.t, labels: ch2.labels, values: ch2.values }], { x: x + 0.08, y: top + 1.62, w: cw - 0.16, h: c.table ? 2.2 : 2.35, barDir: ch2.type, barGrouping: "clustered", ...AX, ...quiet(), ...chartTitle(ch2.t), titleFontSize: 10.5,
      chartColors: ch2.colors, showValue: true, dataLabelPosition: "outEnd", dataLabelColor: TXT, dataLabelFontSize: ch2.fs || (multi ? 8 : 10), dataLabelFormatCode: ch2.fmt,
      showLegend: multi, legendPos: "b", legendFontSize: 9, legendColor: MUTED, barGapWidthPct: 40, valAxisMaxVal: ch2.max, catAxisOrientation: ch2.type === "bar" ? "maxMin" : "minMax", catAxisLabelFontSize: ch2.fs || 10, ...(ch2.layout ? { layout: ch2.layout } : {}), objectName: `Pay chart ${i}` });
    if (c.table) {
      const hd = ["", "Jun", "Jul", "Aug", "Sep", "Oct"].map((t) => ({ text: t, options: { bold: true, color: MUTED, align: "right" } }));
      const rows = c.table.map((r) => r.map((t, j) => ({ text: t, options: { bold: j === 0, color: j === 0 ? NAVY : TXT, align: j === 0 ? "left" : "right" } })));
      const fw = 1.0, ow = (cw - 0.36 - fw) / 5;
      s.addTable([hd, ...rows], { x: x + 0.18, y: top + 3.84, w: cw - 0.36, colW: [fw, ow, ow, ow, ow, ow], rowH: 0.18, fontSize: 8.5, fontFace: "Arial", margin: [0, 0.04, 0, 0.04], valign: "middle", border: { type: "none" }, objectName: `Pay table ${i}` });
    }
    if (c.note) s.addText(c.note, { x: x + 0.18, y: top + 4.0, w: cw - 0.36, h: 0.4, fontSize: 9.5, italic: true, color: MUTED, margin: 0, valign: "top", isTextBox: true, objectName: `Pay note ${i}` });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + 0.16, y: top + 4.45, w: cw - 0.32, h: 0.68, rectRadius: 0.06, fill: { color: "FFFFFF" }, line: { type: "none" }, objectName: `Pay so box ${i}` });
    s.addText([{ text: "SO  ", options: { bold: true, color: CYAN, fontSize: 9.5, charSpacing: 1 } }, { text: c.so, options: { color: TXT, fontSize: 10.5 } }],
      { x: x + 0.28, y: top + 4.45, w: cw - 0.56, h: 0.68, valign: "middle", margin: 0, isTextBox: true, objectName: `Pay so ${i}` });
  });
}



// ---------- 8 Low KTB cannibalization
{
  const s = pres.addSlide({ masterName: "QUESTION", sectionTitle: "Storyline" });
  s.addText("8 / 10  |  LOW KTB CANNIBALIZATION", { placeholder: "tracker" });
  s.addText("CLICX adds to KTB rather than taking from it: 64% of shared customers were inactive at KTB, and 450K of KTB's loan rejects applied with us", { placeholder: "title" });
  s.addText("KTB-inactive = no KTB activity before joining CLICX. Jai-pump rejects Jun 2025–May 2026. FPDX = early default, measured on borrowers who reached their first bill. CLICX data as of Sep 2026. Source: Data for Business Strategy (draft)", { placeholder: "foot" });
  const y = 1.6, H = 5.15, LW = 5.15, RX = 5.8, RW = 12.83 - 5.8;
  card(s, 0.5, y, LW, H, C.background2, "KTB inactive panel");
  panelHead(s, 0.5, y, LW, 1, "1.1M shared customers were inactive at KTB", "KTB customers on CLICX");
  const kb = [["1.6M", "KTB customers on CLICX", 1, NAVY], ["1.1M · 64%", "were inactive at KTB", 0.64, BLUE]];
  kb.forEach((f, k) => {
    const yy = y + 1.0 + k * 0.5, ww = (LW - 0.4) * f[2];
    s.addShape(pres.shapes.RECTANGLE, { x: 0.7, y: yy, w: ww, h: 0.42, fill: { color: f[3] }, line: { type: "none" }, objectName: `KTB bar ${k}` });
    txt(s, [{ text: f[0] + "  ", options: { bold: true } }, { text: f[1] }], { x: 0.82, y: yy, w: ww - 0.2, h: 0.42, fontSize: 11, color: "FFFFFF", valign: "middle", objectName: `KTB bar text ${k}` });
  });
  card(s, 0.7, y + 2.12, LW - 0.4, 0.95, STAT_BG, "KTB stat");
  txt(s, [{ text: "84%  ", options: { bold: true, fontSize: 26, color: BLUE } }, { text: "of them came to CLICX for a loan (borrowed or applied)", options: { fontSize: 11, color: TXT } }], { x: 0.85, y: y + 2.12, w: LW - 0.7, h: 0.95, valign: "middle", objectName: "KTB stat text" });
  const kf = [["515K", "bank mainly with other banks"], ["548K", "were dormant or used only Paotang"], ["37%", "open the CLICX app; 4% use Save Max/More"]];
  kf.forEach((f, k) => txt(s, [{ text: f[0] + "  ", options: { bold: true, fontSize: 14, color: NAVY } }, { text: f[1], options: { fontSize: 11, color: TXT } }],
    { x: 0.7, y: y + 3.25 + k * 0.55, w: LW - 0.4, h: 0.5, valign: "middle", objectName: `KTB fact ${k}` }));
  card(s, RX, y, RW, H, C.background2, "Reject panel");
  panelHead(s, RX, y, RW, 2, "KTB's loan rejects become our customers", "Jai-pump rejects, last 12 months");
  const fn = [["1.9M", "Jai-pump rejects"], ["450K", "applied with CLICX, 24% of rejects"], ["51K", "got a loan, 11% of applicants"], ["20.2%", "early default (FPDX)"]];
  const fw = (RW - 0.4 - 0.45) / 4;
  fn.forEach((f, k) => {
    const x = RX + 0.2 + k * (fw + 0.15);
    card(s, x, y + 1.0, fw, 1.15, k === 3 ? "FFFFFF" : STAT_BG, `Funnel box ${k}`);
    txt(s, [{ text: f[0], options: { bold: true, fontSize: 22, color: k === 3 ? NAVY : BLUE, breakLine: true } }, { text: f[1], options: { fontSize: 9.5, color: TXT } }], { x: x + 0.12, y: y + 1.05, w: fw - 0.24, h: 1.05, valign: "middle", objectName: `Funnel step ${k}` });
    if (k < 3) txt(s, "›", { x: x + fw - 0.02, y: y + 1.35, w: 0.19, h: 0.4, fontSize: 20, bold: true, color: BLUE, align: "center", valign: "middle", objectName: `Funnel arrow ${k}` });
  });
  const ins = [["Paotang converts better", "Rejected via Paotang apply at 30% vs 23% via NEXT, and default less (18% vs 21%)"],
    ["Riders apply most, get a loan least", "64% of rider rejects apply, but only 7% of them get a loan"],
    ["Core age is 25–39", "61% of applicants; apply rate peaks at 29% for ages 25–34"]];
  const iw = (RW - 0.4 - 0.3) / 3;
  ins.forEach((t, k) => {
    const x = RX + 0.2 + k * (iw + 0.15);
    card(s, x, y + 2.3, iw, 1.55, "FFFFFF", `Reject insight ${k}`);
    badge(s, x + 0.12, y + 2.42, k + 1, BLUE, 0.3);
    txt(s, t[0], { x: x + 0.5, y: y + 2.38, w: iw - 0.6, h: 0.42, fontSize: 10.5, bold: true, color: NAVY, valign: "middle", objectName: `Reject insight head ${k}` });
    txt(s, t[1], { x: x + 0.12, y: y + 2.86, w: iw - 0.24, h: 0.92, fontSize: 9.5, color: TXT, valign: "top", objectName: `Reject insight text ${k}` });
  });
  card(s, RX + 0.2, y + 4.0, RW - 0.4, 0.95, NAVY, "Reject action");
  txt(s, [{ text: "ACTION  ", options: { bold: true, fontSize: 10, color: LCYAN, charSpacing: 1 } }, { text: "Promote CLICX on the Paotang reject list, under-40s and freelancers (especially riders) first, and price loans for the 20% early-default rate", options: { fontSize: 11.5, color: "FFFFFF" } }],
    { x: RX + 0.38, y: y + 4.0, w: RW - 0.76, h: 0.95, valign: "middle", objectName: "Reject action text" });
}

// ---------- 9 Defend against True
{
  const s = pres.addSlide({ masterName: "QUESTION", sectionTitle: "Storyline" });
  s.addText("9 / 10  |  DEFEND AGAINST TRUE", { placeholder: "tracker" });
  s.addText("True's new virtual bank puts ฿1.98B (63%) of our deposits at risk; defend first where KTB customers already use TrueMoney", { placeholder: "title" });
  s.addText("AIS no TrueMoney = AIS customers who don't use TrueMoney Wallet; True = True and TrueMoney customers (561,177 and 1,257,522 customers; 1.82M in all). Tech-savvy = uses Make, Kept, Dime or crypto. TrueMoney use and Jai-pump rejects from KTB data. Source: Data for Business Strategy (draft), Sep 2026", { placeholder: "foot" });
  const y = 1.6, H = 4.65, PW = (12.33 - 0.15) / 2;
  card(s, 0.5, y, PW, H, C.background2, "Deposit panel");
  panelHead(s, 0.5, y, PW, 1, "True and TrueMoney customers hold almost two-thirds of our deposits", "CLICX customers and balances, AIS without TrueMoney vs True", TRUEC);
  const sb = [["Customers", 0.3086, "561K · 31%", "1.26M · 69%", 0.6914], ["Deposits", 0.37, "฿1.17B · 37%", "฿1.98B · 63%", 0.63]];
  const bx = 1.85, bw = PW - 1.75;
  sb.forEach((r, k) => {
    const yy = y + 1.05 + k * 0.72;
    txt(s, r[0], { x: 0.7, y: yy, w: 1.1, h: 0.56, fontSize: 11, bold: true, color: TXT, valign: "middle", objectName: `Split label ${k}` });
    s.addShape(pres.shapes.RECTANGLE, { x: bx, y: yy, w: bw * r[1], h: 0.56, fill: { color: NAVY }, line: { color: "FFFFFF", width: 1.5 }, objectName: `AIS part ${k}` });
    if (r[1] + r[4] < 0.995) s.addShape(pres.shapes.RECTANGLE, { x: bx + bw * (r[1] + r[4]), y: yy, w: bw * (1 - r[1] - r[4]), h: 0.56, fill: { color: LGREY }, line: { color: "FFFFFF", width: 1.5 }, objectName: `Other part ${k}` });
    s.addShape(pres.shapes.RECTANGLE, { x: bx + bw * r[1], y: yy, w: bw * r[4], h: 0.56, fill: { color: TRUEC }, line: { color: "FFFFFF", width: 1.5 }, objectName: `True part ${k}` });
    txt(s, r[2], { x: bx + 0.1, y: yy, w: bw * r[1] - 0.2, h: 0.56, fontSize: 10.5, bold: true, color: "FFFFFF", valign: "middle", objectName: `AIS value ${k}` });
    txt(s, r[3], { x: bx + bw * r[1], y: yy, w: bw * r[4], h: 0.56, fontSize: 10, bold: true, color: "FFFFFF", align: "center", valign: "middle", objectName: `True value ${k}` });
  });
  [["AIS, no TrueMoney", NAVY], ["True and TrueMoney", TRUEC]].forEach((l, k) => {
    s.addShape(pres.shapes.RECTANGLE, { x: bx + k * 2.0, y: y + 2.55, w: 0.16, h: 0.16, fill: { color: l[1] }, line: { type: "none" }, objectName: `Split key ${k}` });
    txt(s, l[0], { x: bx + 0.22 + k * 2.0, y: y + 2.5, w: 1.75, h: 0.26, fontSize: 10, color: MUTED, valign: "middle", objectName: `Split legend ${k}` });
  });
  const tf = [["67% vs 58%", "of True vs AIS customers are under 40: True's base is younger"], ["33% vs 20%", "are tech-savvy: more digital, so easier to lose"]];
  tf.forEach((f, k) => {
    card(s, 0.7, y + 2.95 + k * 0.78, PW - 0.4, 0.68, "FFFFFF", `True fact box ${k}`);
    txt(s, [{ text: f[0] + "  ", options: { bold: true, fontSize: 15, color: TRUEC } }, { text: f[1], options: { fontSize: 10.5, color: TXT } }], { x: 0.85, y: y + 2.95 + k * 0.78, w: PW - 0.7, h: 0.68, valign: "middle", objectName: `True fact ${k}` });
  });
  const X2 = 0.5 + PW + 0.15;
  card(s, X2, y, PW, H, C.background2, "Overlap panel");
  panelHead(s, X2, y, PW, 2, "Act first where KTB and True overlap", "KTB customers in the data room", TRUEC);
  const fw2 = PW - 0.4;
  card(s, X2 + 0.2, y + 0.98, fw2, 0.5, NAVY, "KTB base");
  txt(s, [{ text: "35M  ", options: { bold: true, fontSize: 14 } }, { text: "KTB customers  →  ", options: { fontSize: 11 } }, { text: "2.5M (7%) ", options: { bold: true, fontSize: 14 } }, { text: "use TrueMoney", options: { fontSize: 11 } }],
    { x: X2 + 0.35, y: y + 0.98, w: fw2 - 0.3, h: 0.5, color: "FFFFFF", valign: "middle", objectName: "KTB base text" });
  const hw = (fw2 - 0.15) / 2;
  [["Already on CLICX", "225K", "144K KTB-active · 81K not", STAT_BG], ["Not on CLICX", "2.31M", "1.65M KTB-active · 660K not", "FFFFFF"]].forEach((b, k) => {
    const x = X2 + 0.2 + k * (hw + 0.15);
    card(s, x, y + 1.6, hw, 0.95, b[3], `Overlap box ${k}`);
    txt(s, [{ text: b[0], options: { fontSize: 9.5, color: MUTED, breakLine: true } }, { text: b[1], options: { bold: true, fontSize: 18, color: NAVY, breakLine: true } }, { text: b[2], options: { fontSize: 9, color: TXT } }],
      { x: x + 0.12, y: y + 1.6, w: hw - 0.24, h: 0.95, valign: "middle", objectName: `Overlap text ${k}` });
  });
  txt(s, "TURNED DOWN BY JAI-PUMP, NOT ON CLICX", { x: X2 + 0.2, y: y + 2.68, w: fw2, h: 0.24, fontSize: 9, bold: true, color: TRUEC, charSpacing: 1, objectName: "Priority head" });
  [["Priority 1", "99K", "KTB-inactive: easy to win, no KTB relationship to lose"], ["Priority 2", "214K", "KTB-active: protect them from True"]].forEach((p, k) => {
    const yy = y + 2.98 + k * 0.72;
    card(s, X2 + 0.2, yy, fw2, 0.64, "FFFFFF", `Priority box ${k}`);
    txt(s, [{ text: p[0] + "   ", options: { bold: true, fontSize: 10, color: TRUEC } }, { text: p[1] + "  ", options: { bold: true, fontSize: 16, color: NAVY } }, { text: p[2], options: { fontSize: 10, color: TXT } }],
      { x: X2 + 0.35, y: yy, w: fw2 - 0.3, h: 0.64, valign: "middle", objectName: `Priority text ${k}` });
  });
  soBar(s, 6.37, "Hold deposits with savings rewards and a credit card for young True customers; offer loans to 313K TrueMoney users KTB turned down");
}

// ---------- 10 What we will do
{
  const s = pres.addSlide({ masterName: "QUESTION", sectionTitle: "Storyline" });
  s.addText("10 / 10  |  WHAT WE WILL DO", { placeholder: "tracker" });
  s.addText("Four moves for 2027: acquire through partners, activate with loans and payments, launch three products, and defend against True", { placeholder: "title" });
  s.addText(SRC2, { placeholder: "foot" });
  const mv = [
    ["acquire.png", NAVY, "ACQUIRE", "Turn KTB's rejects into customers", ["Paotang reject list: 1.9M a year, 24% apply", "Under-40s, freelancers and riders first", "Partner campaigns aimed at tech-savvy and Gen Z, who join at up to 2× the average"]],
    ["activate.png", CYAN, "ACTIVATE", "Give every customer a reason to come back", ["AIS prepaid top-up first: up to +22K payers a month", "Amazon + PTT offers; reward the second and third month", "Savings for the 103K who only borrow", "Smaller first loans and savings for 924K turned-down applicants"]],
    ["chart.png", BLUE, "GROW", "Launch the next three products", ["Digital credit card: 1.1M earning over ฿15K", "Investing: 528K tech-savvy, 110K already saving", "SME loan: 106K merchants earning over ฿30K"]],
    ["shield.png", TRUEC, "DEFEND", "Hold deposits as True launches", ["฿1.98B held by 1.26M True and TrueMoney customers", "Savings rewards and a credit card for under-40s", "Reach 313K TrueMoney users KTB turned down"]],
  ];
  const cw = (12.33 - 0.45) / 4, y = 1.6, H = 4.5;
  mv.forEach((m, k) => {
    const x = 0.5 + k * (cw + 0.15);
    card(s, x, y, cw, H, C.background2, `Move card ${k}`);
    s.addShape(pres.shapes.RECTANGLE, { x, y, w: cw, h: 0.08, fill: { color: m[1] }, line: { type: "none" }, objectName: `Move strip ${k}` });
    circIcon(s, x + 0.2, y + 0.28, 0.52, m[0], m[1], `Move icon ${k}`);
    txt(s, m[2], { x: x + 0.85, y: y + 0.3, w: cw - 1.0, h: 0.24, fontSize: 10, bold: true, color: m[1], charSpacing: 1, objectName: `Move tag ${k}` });
    txt(s, m[3], { x: x + 0.85, y: y + 0.52, w: cw - 1.0, h: 0.5, fontSize: 12.5, bold: true, color: NAVY, valign: "top", objectName: `Move head ${k}` });
    txt(s, m[4].map((t, i) => ({ text: t, options: { bullet: { indent: 12 }, breakLine: i < m[4].length - 1 } })), { x: x + 0.2, y: y + 1.25, w: cw - 0.4, h: H - 1.4, fontSize: 11, color: TXT, valign: "top", paraSpaceAfter: 6, objectName: `Move points ${k}` });
  });
  card(s, 0.5, 6.22, 12.33, 0.55, NAVY, "Board asks");
  txt(s, [{ text: "BOARD ASKS  ", options: { bold: true, fontSize: 10, color: LCYAN, charSpacing: 1 } }, { text: "Risk appetite for alternative-data lending · partner offer budget aimed at prepaid and Amazon · revenue, cost and credit-loss data in the data room", options: { fontSize: 11.5, color: "FFFFFF" } }],
    { x: 0.7, y: 6.22, w: 12.0, h: 0.55, valign: "middle", objectName: "Board asks text" });
}

// =====================================================================
pres.addSection({ title: "Appendix" });

// ---------- A1 Six segments
{
  const s = pres.addSlide({ masterName: "QUESTION", sectionTitle: "Appendix" });
  s.addText("APPENDIX A1  |  SIX CUSTOMER SEGMENTS BY AGE", { placeholder: "tracker" });
  s.addText("30–44s are half the base; balances and saving rise with age, while loan demand peaks between 20 and 44", { placeholder: "title" });
  s.addText("Active = monthly transacting user. Balance = average (median in brackets). Tech-savvy from AIS and KTB data. Segments sum to 2.0M; the rest have no age on file. Source: Data for Business Strategy (draft), Sep 2026", { placeholder: "foot" });
  const hd = ["Segment", "Customers", "Active monthly", "Paying", "Save Max/ More", "Avg balance (median)", "Applied for a loan", "Tech-savvy", "What they look like"];
  const rows = [
    ["Student (under 20)", "14K", "21%", "18%", "9%", "฿1,019 (฿18)", "45%", "32%", "Wealth-aware minors: nurture as a future main bank"],
    ["Gen Z first jobber (20–24)", "169K", "13%", "10%", "4%", "฿612 (฿2)", "85%", "36%", "High cash need, tech-savvy: first credit card, first investing"],
    ["Gen Z career builder (25–29)", "350K", "15%", "12%", "5%", "฿1,419 (฿4)", "87%", "31%", "Wealth builders with a high need for cash"],
    ["Experienced worker (30–44)", "1,051K", "16%", "13%", "5%", "฿3,394 (฿10)", "83%", "28%", "The core of the base"],
    ["Working senior (45–59)", "369K", "17%", "13%", "7%", "฿7,902 (฿94)", "74%", "24%", "Affluent caregivers"],
    ["Senior (60+)", "39K", "14%", "11%", "12%", "฿35,202 (฿1,000)", "41%", "2%", "Capital preservers; high cost of funds"],
  ];
  const H = (t, al) => ({ text: t, options: { bold: true, fontSize: 10, color: "FFFFFF", fill: { color: NAVY }, align: al, valign: "middle" } });
  const hi = { 2: [0, 4], 4: [5], 6: [1, 2], 7: [0, 1] };
  const tbl = [hd.map((h, j) => H(h, j === 0 || j === 8 ? "left" : "center")), ...rows.map((r, i) => r.map((v, j) => ({ text: v, options: {
    fontSize: 10.5, color: TXT, bold: j === 0, align: j === 0 || j === 8 ? "left" : "center", valign: "middle",
    fill: { color: (hi[j] || []).includes(i) ? "CDEAE7" : (i % 2 ? "FFFFFF" : "F4F6F9") } } })))];
  s.addTable(tbl, { x: 0.5, y: 1.65, w: 12.33, colW: [2.05, 1.0, 0.95, 0.8, 0.9, 1.45, 0.95, 0.9, 3.33], rowH: [0.55, ...rows.map(() => 0.62)], fontFace: "Arial", border: { type: "solid", pt: 0.5, color: "E3E8EE" }, margin: [0.04, 0.08, 0.04, 0.08], objectName: "Segments table" });
  txt(s, "Teal = highest in each column", { x: 0.5, y: 6.0, w: 4, h: 0.24, fontSize: 9.5, italic: true, color: MUTED, objectName: "Table key" });
}

// ---------- A2 KTB inactive vs all customers
{
  const s = pres.addSlide({ masterName: "QUESTION", sectionTitle: "Appendix" });
  s.addText("APPENDIX A2  |  WHAT KTB-INACTIVE CUSTOMERS DO ON CLICX", { placeholder: "tracker" });
  s.addText("84% of the 1.1M KTB-inactive customers came to CLICX for a loan, against 73% of all customers", { placeholder: "title" });
  s.addText("Share of each base. Have balance = balance over ฿0. Loan applicants = applied but no loan booked. KTB-inactive base 1.06M; all customers 2.14M. CLICX data as of Sep 2026; KTB as of Apr 2026. Source: Data for Business Strategy (draft)", { placeholder: "foot" });
  const rows = [
    ["", "Borrower", 74, 7, 158, 7], ["", "Have balance", 26, 2, 92, 4], ["", "No balance", 54, 5, 111, 5], ["", "No deposit", 84, 8, 375, 17],
    ["Loan applicants (not booked)", "Have balance", 74, 7, 161, 8], ["", "No balance", 534, 50, 874, 41], ["", "No deposit", 216, 20, 371, 17]];
  const y0 = 1.95, rh = 0.6, lx = 0.5, gx = 2.6, bx = 4.4, bw = 4.8;
  [["KTB-inactive (1.1M)", CYAN], ["All customers (2.1M)", "B8C2CC"]].forEach((l, k) => {
    s.addShape(pres.shapes.RECTANGLE, { x: bx + k * 2.4, y: 1.7, w: 0.16, h: 0.16, fill: { color: l[1] }, line: { type: "none" }, objectName: `A2 key ${k}` });
    txt(s, l[0], { x: bx + 0.22 + k * 2.4, y: 1.65, w: 2.1, h: 0.26, fontSize: 10, color: MUTED, valign: "middle", objectName: `A2 legend ${k}` });
  });
  rows.forEach((r, k) => {
    const yy = y0 + k * rh + (k >= 4 ? 0.15 : 0);
    if (r[0]) txt(s, r[0], { x: lx, y: yy, w: 1.9, h: rh * 3, fontSize: 11, bold: true, color: NAVY, valign: "middle", objectName: `A2 group ${k}` });
    txt(s, r[1], { x: gx, y: yy, w: 1.7, h: rh, fontSize: 11, color: TXT, valign: "middle", objectName: `A2 label ${k}` });
    const w1 = bw * r[3] / 55, w2 = bw * r[5] / 55;
    s.addShape(pres.shapes.RECTANGLE, { x: bx, y: yy + 0.06, w: Math.max(0.03, w1), h: 0.22, fill: { color: k === 5 ? "3A8A80" : CYAN }, line: { type: "none" }, objectName: `A2 bar ${k}` });
    txt(s, `${r[2]}K (${r[3]}%)`, { x: bx + w1 + 0.06, y: yy + 0.02, w: 1.6, h: 0.3, fontSize: 10, bold: true, color: NAVY, valign: "middle", objectName: `A2 value ${k}` });
    s.addShape(pres.shapes.RECTANGLE, { x: bx, y: yy + 0.31, w: Math.max(0.03, w2), h: 0.18, fill: { color: "B8C2CC" }, line: { type: "none" }, objectName: `A2 base bar ${k}` });
    txt(s, `${r[4]}K (${r[5]}%)`, { x: bx + w2 + 0.06, y: yy + 0.26, w: 1.6, h: 0.26, fontSize: 9, color: MUTED, valign: "middle", objectName: `A2 base value ${k}` });
  });
  s.addShape(pres.shapes.LINE, { x: 0.5, y: y0 + 4 * rh + 0.07, w: 9.9, h: 0, line: { color: LGREY, width: 1 }, objectName: "A2 divider" });
  card(s, 10.55, 1.95, 2.28, 4.4, C.background2, "A2 side");
  txt(s, [{ text: "84%", options: { bold: true, fontSize: 26, color: BLUE, breakLine: true } }, { text: "of KTB-inactive customers borrowed or applied, vs 73% overall", options: { fontSize: 10.5, color: TXT, breakLine: true } },
    { text: " ", options: { fontSize: 8, breakLine: true } }, { text: "50%", options: { bold: true, fontSize: 26, color: NAVY, breakLine: true } }, { text: "applied, didn't get a loan and keep no balance: the biggest group to re-engage", options: { fontSize: 10.5, color: TXT } }],
    { x: 10.7, y: 2.05, w: 2.0, h: 4.2, valign: "top", objectName: "A2 side text" });
}

// ---------- A3 Payment segments
{
  const s = pres.addSlide({ masterName: "QUESTION", sectionTitle: "Appendix" });
  s.addText("APPENDIX A3  |  PAYMENT SEGMENTS, AUG–OCT 2026", { placeholder: "tracker" });
  s.addText("Win back the 114K who paid recently and stopped before chasing the 1.27M who never paid", { placeholder: "title" });
  s.addText("Each customer sits in one segment by the months they paid in-app (Aug, Sep, Oct). This extract counts 117K Oct payers; the partner analysis on slide 7 counts 127K from a different extract. Next steps are suggestions to test. Source: Data for Business Strategy (draft)", { placeholder: "foot" });
  const rows = [["Habitual", "Aug, Sep, Oct", 57652, "Offer the payment type they don't use yet", CYAN], ["Building", "Sep, Oct", 20220, "Reward a third month in a row", CYAN],
    ["New", "Oct only", 24502, "Cashback on a second payment within 30 days", CYAN], ["Reactivated", "Aug, Oct", 14672, "Autopay or recurring bills", CYAN],
    ["At risk", "Sep, not Oct", 52561, "Win back now: nudge before the bill date", ORANGE], ["Lapsed", "Aug only", 61323, "Low-cost win-back offer and survey", ORANGE],
    ["Never paid", "none", 1270212, "First payment via loan repayment (75K) or deposit (240K); 955K need a first deposit", GREY]];
  const y0 = 1.95, rh = 0.56, bx = 3.6, bw = 3.2;
  ["SEGMENT", "PAID IN", "CUSTOMERS", "NEXT STEP"].forEach((h, k) => txt(s, h, { x: [0.5, 2.0, bx, 7.9][k], y: 1.65, w: 2, h: 0.24, fontSize: 9, bold: true, color: MUTED, charSpacing: 1, objectName: `A3 head ${k}` }));
  rows.forEach((r, k) => {
    const yy = y0 + k * rh + (k >= 4 ? 0.5 : 0) + (k >= 6 ? 0.1 : 0);
    if (k === 4 || k === 6 || k === 0) s.addShape(pres.shapes.LINE, { x: 0.5, y: yy - 0.04, w: 12.33, h: 0, line: { color: LGREY, width: 1 }, objectName: `A3 rule ${k}` });
    if (k >= 4 && k < 6) card(s, 0.45, yy + 0.02, 12.43, rh - 0.04, "FBEFE2", `A3 risk ${k}`);
    txt(s, r[0], { x: 0.55, y: yy, w: 1.45, h: rh, fontSize: 12, bold: true, color: NAVY, valign: "middle", objectName: `A3 seg ${k}` });
    txt(s, r[1], { x: 2.0, y: yy, w: 1.5, h: rh, fontSize: 10.5, color: MUTED, valign: "middle", objectName: `A3 months ${k}` });
    const w = k === 6 ? bw : bw * r[2] / 61323;
    s.addShape(pres.shapes.RECTANGLE, { x: bx, y: yy + 0.15, w, h: rh - 0.3, fill: { color: r[4] }, line: { type: "none" }, objectName: `A3 bar ${k}` });
    txt(s, r[2].toLocaleString("en-US") + (k === 6 ? "  (bar off scale)" : ""), { x: k === 6 ? bx + 0.1 : bx + w + 0.08, y: yy, w: 2.2, h: rh, fontSize: 11, bold: true, color: k === 6 ? "FFFFFF" : NAVY, valign: "middle", objectName: `A3 value ${k}` });
    txt(s, r[3], { x: 7.9, y: yy, w: 4.9, h: rh, fontSize: 10.5, color: TXT, valign: "middle", objectName: `A3 next ${k}` });
  });
  const sy = y0 + 4 * rh;
  txt(s, [{ text: "117,046 paid in Oct", options: { bold: true, color: NAVY } }, { text: "  ·  59.7% of Sep payers paid again in Oct", options: { color: MUTED } }], { x: 2.0, y: sy + 0.05, w: 8, h: 0.4, fontSize: 11, valign: "middle", objectName: "A3 subtotal" });
}

// ---------- Slide 3: what we know, who holds what
{
  const s = pres.addSlide({ masterName: "QUESTION", sectionTitle: "Appendix" });
  s.addText("APPENDIX A4  |  PRODUCT HOLDING BY AGE AND JOB", { placeholder: "tracker" });
  s.addText("Age and job shape what people use: people in their 30s and 40s borrow, those over 60 save; salaried workers borrow, business owners save", { placeholder: "title" });
  s.addText("Share of each group's customers; a customer can hold several. By age: product segments, 31 Oct 2026 (borrow = any loan; save = balance over ฿100 or Save Max/More; pay = in-app payments). By job: Data for Strategy 2027, Sep 2026 (save = balance over ฿0). * under 11K customers", { placeholder: "foot" });
  const pw = (12.33 - 0.15) / 2, py = 1.6, ph = 5.18;
  const ser = (labels, b, sv, p) => [{ name: "Borrow", labels, values: b }, { name: "Save", labels, values: sv }, { name: "Pay", labels, values: p }];
  const panels = [
    { n: "1", head: "By age: teens pay, 30–49s borrow, and saving rises to 28% at 60–69", chartT: "% of each age group holding each product",
      data: ser(["10–19*", "20–29", "30–39", "40–49", "50–59", "60–69"], [0, 0.109, 0.116, 0.113, 0.094, 0], [0.113, 0.043, 0.07, 0.105, 0.165, 0.28], [0.142, 0.069, 0.077, 0.074, 0.067, 0.066]),
      max: 0.32, so: "Payments for the young, credit for 20–59s, savings and then investment for 50+ (P4). No loans under 20 or over 59 (age limits)" },
    { n: "2", head: "By job: salaried borrow most, merchants save most; blue-collar and riders use least of everything", chartT: "% of each job group holding each product",
      data: ser(["Salaried", "Merchant", "Freelance", "Rider", "Blue collar"], [0.11, 0.07, 0.06, 0.05, 0.07], [0.18, 0.20, 0.16, 0.13, 0.12], [0.14, 0.14, 0.12, 0.11, 0.10]),
      max: 0.24, so: "Card and salary-bridged loans for salaried (P1, P2); deposits and SME credit for merchants; savings and insurance first for blue-collar (P3)" },
  ];
  panels.forEach((p, i) => {
    const x = 0.5 + i * (pw + 0.15);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: py, w: pw, h: ph, rectRadius: 0.08, fill: { color: C.background2 }, line: { type: "none" }, objectName: `Product panel ${i}` });
    s.addShape(pres.shapes.OVAL, { x: x + 0.18, y: py + 0.16, w: 0.42, h: 0.42, fill: { color: C.text2 }, line: { type: "none" }, objectName: `Insight badge ${i}` });
    s.addText(p.n, { x: x + 0.18, y: py + 0.16, w: 0.42, h: 0.42, fontSize: 13, bold: true, color: "FFFFFF", align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: `Insight num ${i}` });
    s.addText(p.head, { x: x + 0.72, y: py + 0.1, w: pw - 0.9, h: 0.6, fontSize: 13.5, bold: true, color: NAVY, margin: 0, valign: "middle", isTextBox: true, objectName: `Insight head ${i}` });
    s.addChart(pres.charts.BAR, p.data, { x: x + 0.1, y: py + 0.78, w: pw - 0.2, h: ph - 1.73, barDir: "col", barGrouping: "clustered", ...AX, ...quiet(), ...chartTitle(p.chartT), titleFontSize: 11,
      chartColors: [NAVY, CYAN, BLUE], showValue: true, dataLabelPosition: "outEnd", dataLabelColor: TXT, dataLabelFontSize: 9, dataLabelFormatCode: "0%",
      showLegend: true, legendPos: "b", legendFontSize: 10, legendColor: MUTED, barGapWidthPct: 45, valAxisMaxVal: p.max, catAxisLabelFontSize: 10.5, objectName: `Product chart ${i}` });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + 0.18, y: py + ph - 0.86, w: pw - 0.36, h: 0.72, rectRadius: 0.06, fill: { color: "FFFFFF" }, line: { type: "none" }, objectName: `So box ${i}` });
    s.addText([{ text: "SO  ", options: { bold: true, color: CYAN, fontSize: 9.5, charSpacing: 1 } }, { text: p.so, options: { color: TXT, fontSize: 10.5 } }],
      { x: x + 0.32, y: py + ph - 0.86, w: pw - 0.64, h: 0.72, valign: "middle", margin: 0, isTextBox: true, objectName: `So text ${i}` });
  });
}



require(__dirname + "/region_slide.js")(pres, { master: "QUESTION", tracker: "APPENDIX A5  |  PRODUCT HOLDING BY REGION" });

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
