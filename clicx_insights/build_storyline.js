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
pres.title = "CLICX 2027 board storyline";
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

// ---------- Slide 1: diagnosis
{
  const s = pres.addSlide({ masterName: "QUESTION", sectionTitle: "Storyline" });
  s.addText("1 / 2  |  WHAT WE KNOW", { placeholder: "tracker" });
  s.addText("Partners gave us scale fast, but only 15% of customers are active, and loans are both our hook and our constraint", { placeholder: "title" });
  s.addText(SRC, { placeholder: "foot" });
  // governing question banner
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 1.5, w: 12.33, h: 0.42, rectRadius: 0.08, fill: { color: C.text2 }, line: { type: "none" }, objectName: "Governing question" });
  s.addText([{ text: "THE QUESTION  ", options: { bold: true, fontSize: 10, color: C.accent6, charSpacing: 1 } },
    { text: "How do we turn 2.1M sign-ups into a profitable, active customer base by 2029?", options: { bold: true, fontSize: 14, color: "FFFFFF" } }],
    { x: 0.7, y: 1.5, w: 12.0, h: 0.42, valign: "middle", margin: 0, isTextBox: true, objectName: "Governing question text" });

  const panel = (x, y, w, h, n, q, ans, big, bigLbl, n2) => {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.08, fill: { color: C.background2 }, line: { type: "none" }, objectName: `Panel ${n}` });
    s.addText([{ text: n + "  ", options: { color: CYAN, bold: true } }, { text: q, options: { color: MUTED } }],
      { x: x + 0.18, y: y + 0.1, w: w - 0.36, h: 0.28, fontSize: 10.5, margin: 0, valign: "middle", isTextBox: true, objectName: `SubQ ${n}` });
    s.addText(ans, { x: x + 0.18, y: y + 0.38, w: w - 0.36, h: 0.5, fontSize: 13, bold: true, color: NAVY, margin: 0, valign: "top", isTextBox: true, objectName: `Answer ${n}` });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + 0.18, y: y + 0.95, w: 1.5, h: h - 1.1, rectRadius: 0.08, fill: { color: STAT_BG }, line: { type: "none" }, objectName: `Stat box ${n}` });
    s.addText(big, { x: x + 0.26, y: y + 1.0, w: 1.38, h: 0.55, fontSize: big.length > 5 ? 22 : 28, bold: true, color: BLUE, margin: 0, valign: "middle", isTextBox: true, objectName: `Stat ${n}` });
    s.addText(bigLbl, { x: x + 0.26, y: y + 1.55, w: 1.38, h: h - 1.75, fontSize: 9.5, color: TXT, margin: 0, valign: "top", isTextBox: true, objectName: `Stat label ${n}` });
  };
  const bars = (x, y, w, rows, n, o = {}) => {
    const lw = o.lw || 1.05, vw = o.vw || 0.55, bw = w - lw - vw - 0.06, rh = o.rh || 0.27;
    if (o.head) s.addText(o.head, { x, y: y - 0.25, w, h: 0.22, fontSize: 9, bold: true, color: MUTED, margin: 0, isTextBox: true, objectName: `Bars head ${n}` });
    rows.forEach((r, k) => {
      const yy = y + k * rh;
      if (r[4]) s.addImage({ path: IC(r[4]), x, y: yy + 0.04, h: 0.19, w: 0.19 * r[5], altText: r[0] + " logo", objectName: `Bar logo ${n}-${k}` });
      else s.addText(r[0], { x, y: yy, w: lw, h: rh, fontSize: 9.5, color: TXT, margin: 0, valign: "middle", isTextBox: true, objectName: `Bar label ${n}-${k}` });
      s.addShape(pres.shapes.RECTANGLE, { x: x + lw, y: yy + rh * 0.3, w: bw, h: rh * 0.4, fill: { color: "E3E8EE" }, line: { type: "none" }, objectName: `Track ${n}-${k}` });
      s.addShape(pres.shapes.RECTANGLE, { x: x + lw, y: yy + rh * 0.3, w: Math.max(0.03, bw * r[1]), h: rh * 0.4, fill: { color: r[3] }, line: { type: "none" }, objectName: `Bar ${n}-${k}` });
      s.addText(r[2], { x: x + lw + bw + 0.06, y: yy, w: vw, h: rh, fontSize: 9.5, bold: true, color: NAVY, margin: 0, valign: "middle", isTextBox: true, objectName: `Bar value ${n}-${k}` });
    });
  };
  const W3 = 4.01, G = 0.15, y1 = 2.05, h1 = 2.3, y2 = 4.48, h2 = 2.3;
  // 1 consortium
  let x = 0.5;
  panel(x, y1, W3, h1, 1, "How dependent are we on the consortium?", "96% come from partners; AIS is the biggest untapped base", "96%", "of customers are AIS, KTB or OR customers", 1);
  bars(x + 1.82, y1 + 1.25, W3 - 2.0, [["KTB", 0.47, "4.7%", BLUE, "logo_ktb.jpeg", 2.58], ["AIS", 0.37, "3.7%", CYAN, "logo_ais.png", 2.74], ["OR", 0.74, "7.4%", NAVY, "logo_or.png", 1.93]], 1, { head: "% of partner base on CLICX", lw: 0.62 });
  // 2 targets
  x = 0.5 + W3 + G;
  panel(x, y1, W3, h1, 2, "Are we winning the customers we planned?", "Yes for tech-savvy and underserved; not yet for multi-bank", "82%", "of loan applicants are underserved, as planned", 2);
  bars(x + 1.82, y1 + 1.25, W3 - 2.0, [["Tech-savvy", 0.85, "8.5%", CYAN], ["Average", 0.38, "3.8%", GREY], ["Multi-bank", 0.24, "2.4%", NAVY]], 2, { head: "% of each group who joined" });
  // 3 activity
  x = 0.5 + 2 * (W3 + G);
  panel(x, y1, W3, h1, 3, "What do customers use, and what drives activity?", "Everyone comes for a loan; money in and a loan drive activity", "84–92%", "of every job type applies for a loan", 3);
  bars(x + 1.82, y1 + 1.25, W3 - 2.0, [["Borrowers*", 0.88, "88%", CYAN], ["Money in", 0.51, "51%", BLUE], ["No money in", 0.04, "4%", GREY]], 3, { head: "% active each month" });
  s.addText("* includes loan repayment", { x: x + 1.82, y: y1 + 2.05, w: W3 - 2.0, h: 0.18, fontSize: 8, italic: true, color: MUTED, margin: 0, isTextBox: true, objectName: "Activity caveat" });
  // 4 economics
  const W2 = (12.33 - G) / 2;
  x = 0.5;
  panel(x, y2, W2, h2, 4, "Can we serve them profitably?", "Approval and credit risk are the constraint; revenue per customer is not yet measured", "20.2%", "early default on loans to KTB rejects", 4);
  bars(x + 1.82, y2 + 1.25, 2.05, [["Office worker", 0.55, "11%", NAVY], ["Merchant", 0.35, "7%", BLUE], ["Blue-collar", 0.35, "7%", BLUE], ["Rider", 0.25, "5%", GREY]], 4, { head: "% of applicants who got a loan", rh: 0.26 });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + 4.0, y: y2 + 0.95, w: W2 - 4.18, h: h2 - 1.1, rectRadius: 0.08, fill: { color: "FFFFFF" }, line: { color: GREY, width: 1, dashType: "dash" }, objectName: "Economics gap" });
  s.addText([{ text: "DATA TO COME", options: { bold: true, fontSize: 9, color: BLUE, charSpacing: 1, breakLine: true } },
    { text: "· Revenue per customer", options: { fontSize: 10, breakLine: true } },
    { text: "· Cost to acquire and serve", options: { fontSize: 10, breakLine: true } },
    { text: "· Credit losses by segment", options: { fontSize: 10 } }],
    { x: x + 4.12, y: y2 + 1.02, w: W2 - 4.4, h: h2 - 1.25, valign: "top", margin: 0, paraSpaceAfter: 2, isTextBox: true, objectName: "Economics gap text" });
  // 5 shareholders
  x = 0.5 + W2 + G;
  panel(x, y2, W2, h2, 5, "What value do we create for each shareholder?", "We recover KTB's lapsed and rejected customers, and bring AIS and OR new customers and spend", "1.1M", "shared customers were inactive with KTB", 5);
  const sh = [["logo_ktb.jpeg", 2.58, "450K people KTB's Jai-pump turned down applied to us; 51K got a loan"],
    ["logo_ais.png", 2.74, "0.66M CLICX customers to convert to AIS; THB 261K top-up fees saved"],
    ["logo_or.png", 1.93, "173K active customers without BLUE+; 1,327 new BLUE+ members"]];
  sh.forEach((r, k) => {
    const yy = y2 + 1.0 + k * 0.42;
    s.addImage({ path: IC(r[0]), x: x + 1.85, y: yy + 0.06, h: 0.22, w: 0.22 * r[1], altText: "Partner logo", objectName: `Value logo ${k}` });
    s.addText(r[2], { x: x + 2.55, y: yy, w: W2 - 2.7, h: 0.38, fontSize: 9.5, color: TXT, margin: 0, valign: "middle", isTextBox: true, objectName: `Value text ${k}` });
  });
}

// ---------- Slide 2: resolution
{
  const s = pres.addSlide({ masterName: "QUESTION", sectionTitle: "Storyline" });
  s.addText("2 / 2  |  WHAT WE WILL DO", { placeholder: "tracker" });
  s.addText("Two engines and three ranked bets turn sign-ups into a profitable base; we need the board's call on risk, budget and data", { placeholder: "title" });
  s.addText("Loan book targets from the CLICX 5-Year Business Plan (end-2029). Card and wealth value still to be sized. Icons: Font Awesome (open licence). " + SRC, { placeholder: "foot" });
  const MW = 8.4, x0 = 0.5;
  // Engines
  s.addText("ENGINES  ·  serve every product", { x: x0, y: 1.5, w: MW, h: 0.26, fontSize: 10, bold: true, color: MUTED, charSpacing: 1, margin: 0, isTextBox: true, objectName: "Engines label" });
  const eng = [["partners", "Acquire through AIS and OR", "Only 3.7% of AIS's base has joined; put bill pay and top-up in the app", [["logo_ais.png", 2.74], ["logo_or.png", 1.93]]],
    ["activate", "Activate: first deposit, first loan", "1.8M customers are idle; money in and a loan are what keep people active", null]];
  eng.forEach((e, i) => {
    const w = (MW - 0.15) / 2, x = x0 + i * (w + 0.15), y = 1.8, h = 1.15;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.08, fill: { color: C.background2 }, line: { type: "none" }, objectName: `Engine ${i}` });
    s.addShape(pres.shapes.OVAL, { x: x + 0.15, y: y + 0.25, w: 0.6, h: 0.6, fill: { color: BLUE }, line: { type: "none" }, objectName: `Engine icon bg ${i}` });
    s.addImage({ path: IC(e[0] + ".png"), x: x + 0.29, y: y + 0.39, w: 0.32, h: 0.32, altText: e[1] + " icon", objectName: `Engine icon ${i}` });
    s.addText(e[1], { x: x + 0.9, y: y + 0.12, w: w - 1.05, h: 0.34, fontSize: 12.5, bold: true, color: NAVY, margin: 0, valign: "middle", isTextBox: true, objectName: `Engine title ${i}` });
    s.addText(e[2], { x: x + 0.9, y: y + 0.48, w: w - (e[3] ? 2.25 : 1.05), h: 0.58, fontSize: 10, color: TXT, margin: 0, valign: "top", isTextBox: true, objectName: `Engine text ${i}` });
    if (e[3]) { let lx = x + w - 0.15; e[3].slice().reverse().forEach(([f, ar], k) => { const hh = 0.22, ww = hh * ar; lx -= ww; s.addImage({ path: IC(f), x: lx, y: y + 0.66, w: ww, h: hh, altText: "Partner logo", objectName: `Engine logo ${i}-${k}` }); lx -= 0.06; }); }
  });
  // Bets
  s.addText("BETS  ·  ranked by value and readiness", { x: x0, y: 3.1, w: MW, h: 0.26, fontSize: 10, bold: true, color: MUTED, charSpacing: 1, margin: 0, isTextBox: true, objectName: "Bets label" });
  const bets = [
    ["lend", "1", "Proxy-income loans for the underserved", "1.45M underserved applicants; 90%+ of gig workers apply, 5–7% approved", "25.4B THB", "retail loan book by 2029", "Fix first: approval rules and the 20% early default"],
    ["sme", "2", "SME cash-flow loans for merchants", "293K merchants, 245K already applied; they save and spend more", "12.2B THB", "SME loan book by 2029", "Ready: lead list in hand"],
    ["card", "3", "Digital credit card for office workers", "700K office workers, borrow most (11%); 239K AIS Serenade", "To size", "card revenue by 2029", "Ready: partner perks with AIS and OR"],
  ];
  const bw = (MW - 0.3) / 3;
  bets.forEach((b, i) => {
    const x = x0 + i * (bw + 0.15), y = 3.4, h = 3.38;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: bw, h, rectRadius: 0.08, fill: { color: C.background2 }, line: { type: "none" }, objectName: `Bet ${i}` });
    s.addShape(pres.shapes.OVAL, { x: x + 0.15, y: y + 0.15, w: 0.56, h: 0.56, fill: { color: CYAN }, line: { type: "none" }, objectName: `Bet icon bg ${i}` });
    s.addImage({ path: IC(b[0] + ".png"), x: x + 0.28, y: y + 0.28, w: 0.3, h: 0.3, altText: b[2] + " icon", objectName: `Bet icon ${i}` });
    s.addText("#" + b[1], { x: x + bw - 0.7, y: y + 0.15, w: 0.55, h: 0.4, fontSize: 20, bold: true, color: BLUE, align: "right", margin: 0, isTextBox: true, objectName: `Bet rank ${i}` });
    s.addText(b[2], { x: x + 0.15, y: y + 0.8, w: bw - 0.3, h: 0.5, fontSize: 12, bold: true, color: NAVY, margin: 0, valign: "top", isTextBox: true, objectName: `Bet title ${i}` });
    s.addText(b[3], { x: x + 0.15, y: y + 1.32, w: bw - 0.3, h: 0.62, fontSize: 9.5, color: MUTED, margin: 0, valign: "top", isTextBox: true, objectName: `Bet why ${i}` });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + 0.15, y: y + 1.98, w: bw - 0.3, h: 0.78, rectRadius: 0.06, fill: { color: STAT_BG }, line: { type: "none" }, objectName: `Bet value box ${i}` });
    s.addText(b[4], { x: x + 0.25, y: y + 2.0, w: bw - 0.5, h: 0.42, fontSize: 20, bold: true, color: b[4] === "To size" ? GREY : BLUE, margin: 0, valign: "middle", isTextBox: true, objectName: `Bet value ${i}` });
    s.addText(b[5], { x: x + 0.25, y: y + 2.42, w: bw - 0.5, h: 0.28, fontSize: 9.5, color: TXT, margin: 0, isTextBox: true, objectName: `Bet value label ${i}` });
    s.addText(b[6], { x: x + 0.15, y: y + 2.84, w: bw - 0.3, h: 0.46, fontSize: 9.5, bold: true, color: i === 0 ? NAVY : CYAN, margin: 0, valign: "top", isTextBox: true, objectName: `Bet readiness ${i}` });
  });
  // Right panel: asks, later, won't play
  const px = 9.05, pw = 3.78;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: px, y: 1.5, w: pw, h: 5.28, rectRadius: 0.08, fill: { color: C.text2 }, line: { type: "none" }, objectName: "Asks panel" });
  s.addText([
    { text: "WHAT WE ASK THE BOARD", options: { bold: true, fontSize: 10, color: C.accent6, charSpacing: 1, breakLine: true } },
    { text: " ", options: { fontSize: 8, breakLine: true } },
    { text: "1  Risk appetite", options: { bold: true, fontSize: 13, color: "FFFFFF", breakLine: true } },
    { text: "Limits for test-and-learn lending to the underserved", options: { fontSize: 11, color: "DCE6EF", breakLine: true } },
    { text: " ", options: { fontSize: 8, breakLine: true } },
    { text: "2  Budget shift", options: { bold: true, fontSize: 13, color: "FFFFFF", breakLine: true } },
    { text: "Move part of sign-up spend into activation", options: { fontSize: 11, color: "DCE6EF", breakLine: true } },
    { text: " ", options: { fontSize: 8, breakLine: true } },
    { text: "3  One scorecard", options: { bold: true, fontSize: 13, color: "FFFFFF", breakLine: true } },
    { text: "Active, funded customers next to the performing loan book; one definition per metric", options: { fontSize: 11, color: "DCE6EF", breakLine: true } },
    { text: " ", options: { fontSize: 8, breakLine: true } },
    { text: "4  Data", options: { bold: true, fontSize: 13, color: "FFFFFF", breakLine: true } },
    { text: "Balances, revenue and credit losses per customer in the data room", options: { fontSize: 11, color: "DCE6EF", breakLine: true } },
    { text: " ", options: { fontSize: 8, breakLine: true } },
    { text: "NEXT WAVE (2028)", options: { bold: true, fontSize: 10, color: C.accent6, charSpacing: 1, breakLine: true } },
    { text: "Wealth for 643K savers on 5.5B THB", options: { fontSize: 11, color: "FFFFFF", breakLine: true } },
    { text: " ", options: { fontSize: 8, breakLine: true } },
    { text: "WHERE WE WON'T PLAY", options: { bold: true, fontSize: 10, color: C.accent6, charSpacing: 1, breakLine: true } },
    { text: "Branch lending · secured home and car loans · prime wealth needing RMs · revolving cash cards", options: { fontSize: 11, color: "FFFFFF" } },
  ], { x: px + 0.2, y: 1.65, w: pw - 0.4, h: 5.0, valign: "top", margin: 0, isTextBox: true, objectName: "Asks text" });
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
