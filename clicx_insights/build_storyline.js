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
  s.addText("1 / 3  |  WHAT WE KNOW", { placeholder: "tracker" });
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

// ---------- Slide 2: from answers to initiatives
const POOL_NOTE = "Prospect pools = CLICX customers by occupation and profile (2.14M customers, Sep 2026); pools overlap";
{
  const s = pres.addSlide({ masterName: "QUESTION", sectionTitle: "Storyline" });
  s.addText("2 / 3  |  WHAT WE WILL DO", { placeholder: "tracker" });
  s.addText("Our answers point to four initiatives, in priority order, built on two engines and one fix to how we approve loans", { placeholder: "title" });
  s.addText(POOL_NOTE + ". Icons: Font Awesome (open licence). " + SRC, { placeholder: "foot" });
  // Foundation row: engines + enabler
  s.addText("FOUNDATION  ·  serves every initiative", { x: 0.5, y: 1.5, w: 12.33, h: 0.26, fontSize: 10, bold: true, color: MUTED, charSpacing: 1, margin: 0, isTextBox: true, objectName: "Foundation label" });
  const fnd = [["partners", "Acquire through AIS and OR", "Only 3.7% of AIS's base has joined", "Q1", [["logo_ais.png", 2.74], ["logo_or.png", 1.93]]],
    ["activate", "Activate: first deposit, first loan", "1.8M idle; money in and a loan keep people active", "Q3", null],
    ["enabler", "Fix approvals and risk", "Gig workers: 90%+ apply, 5–7% approved; 20% early default", "Q4", null]];
  const fw = (12.33 - 0.3) / 3;
  fnd.forEach((f, i) => {
    const x = 0.5 + i * (fw + 0.15), y = 1.8, h = 1.0;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: fw, h, rectRadius: 0.08, fill: { color: C.background2 }, line: { type: "none" }, objectName: `Foundation ${i}` });
    s.addShape(pres.shapes.OVAL, { x: x + 0.15, y: y + 0.2, w: 0.58, h: 0.58, fill: { color: i === 2 ? NAVY : BLUE }, line: { type: "none" }, objectName: `Foundation icon bg ${i}` });
    s.addImage({ path: IC(f[0] + ".png"), x: x + 0.29, y: y + 0.34, w: 0.3, h: 0.3, altText: f[1] + " icon", objectName: `Foundation icon ${i}` });
    s.addText(f[1], { x: x + 0.88, y: y + 0.12, w: fw - 1.5, h: 0.34, fontSize: 12, bold: true, color: NAVY, margin: 0, valign: "middle", isTextBox: true, objectName: `Foundation title ${i}` });
    s.addText(f[3], { x: x + fw - 0.55, y: y + 0.12, w: 0.42, h: 0.26, fontSize: 9, bold: true, color: BLUE, align: "right", margin: 0, isTextBox: true, objectName: `Foundation from ${i}` });
    s.addText(f[2], { x: x + 0.88, y: y + 0.48, w: fw - (f[4] ? 2.15 : 1.05), h: 0.45, fontSize: 9.5, color: TXT, margin: 0, valign: "top", isTextBox: true, objectName: `Foundation text ${i}` });
    if (f[4]) { let lx = x + fw - 0.15; f[4].slice().reverse().forEach(([fl, ar], k) => { const hh = 0.2, ww = hh * ar; lx -= ww; s.addImage({ path: IC(fl), x: lx, y: y + 0.62, w: ww, h: hh, altText: "Partner logo", objectName: `Foundation logo ${i}-${k}` }); lx -= 0.06; }); }
  });
  // Initiatives row
  s.addText("INITIATIVES  ·  in priority order", { x: 0.5, y: 2.95, w: 12.33, h: 0.26, fontSize: 10, bold: true, color: MUTED, charSpacing: 1, margin: 0, isTextBox: true, objectName: "Initiatives label" });
  const ini = [
    ["P1", "salary", "Salary-Bridged Financing", "Q2 · Q3", "Everyone comes for a loan, and salaried workers get approved most (11%)", "607K", "salaried and wage earners · 28% of base", "Corporate crew · SME staffer · Helper"],
    ["P2", "card", "Credit Card & Apple Pay", "Q3 · Q5", "Salary earners borrow and pay most; 158K proven borrowers and 239K AIS Serenade members", "984K", "salary earners and business owners · 46%", "Salary earners · Business owners"],
    ["P3", "insure", "Insurance e-Brokerage", "Q2 · Q4", "Riders and blue-collar workers rarely get a loan; insurance keeps them with us", "836K", "riders, blue- and white-collar workers · 39%", "Riders · Blue collars · White collars"],
    ["P4", "invest", "Digital Investment", "Q2 · Q3", "Tech-savvy people join most (8.5%); savers hold 5.5B THB", "533K", "tech-savvy customers · 25%", "Aspiring wealth · Tech savvy"],
  ];
  const iw = (12.33 - 0.45) / 4;
  ini.forEach((c, i) => {
    const x = 0.5 + i * (iw + 0.15), y = 3.25, h = 3.15;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: iw, h, rectRadius: 0.08, fill: { color: C.background2 }, line: { type: "none" }, objectName: `Initiative ${i}` });
    s.addShape(pres.shapes.OVAL, { x: x + 0.15, y: y + 0.15, w: 0.56, h: 0.56, fill: { color: CYAN }, line: { type: "none" }, objectName: `Initiative icon bg ${i}` });
    s.addImage({ path: IC(c[1] + ".png"), x: x + 0.28, y: y + 0.28, w: 0.3, h: 0.3, altText: c[2] + " icon", objectName: `Initiative icon ${i}` });
    s.addShape(pres.shapes.OVAL, { x: x + iw - 0.66, y: y + 0.17, w: 0.5, h: 0.5, fill: { color: BLUE }, line: { type: "none" }, objectName: `Priority badge ${i}` });
    s.addText(c[0], { x: x + iw - 0.66, y: y + 0.17, w: 0.5, h: 0.5, fontSize: 12, bold: true, color: "FFFFFF", align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: `Priority ${i}` });
    s.addText(c[2], { x: x + 0.15, y: y + 0.8, w: iw - 0.3, h: 0.36, fontSize: 13, bold: true, color: NAVY, margin: 0, valign: "middle", isTextBox: true, objectName: `Initiative name ${i}` });
    s.addText([{ text: c[3] + "  ", options: { bold: true, color: BLUE } }, { text: c[4], options: { color: MUTED } }],
      { x: x + 0.15, y: y + 1.2, w: iw - 0.3, h: 0.78, fontSize: 9.5, margin: 0, valign: "top", isTextBox: true, objectName: `Initiative why ${i}` });
    s.addText([{ text: "FOR  ", options: { bold: true, color: CYAN, fontSize: 8.5, charSpacing: 1 } }, { text: c[7], options: { color: NAVY, bold: true, fontSize: 9.5 } }],
      { x: x + 0.15, y: y + 1.68, w: iw - 0.3, h: 0.32, margin: 0, valign: "middle", isTextBox: true, objectName: `Initiative persona ${i}` });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + 0.15, y: y + 2.05, w: iw - 0.3, h: 0.95, rectRadius: 0.06, fill: { color: STAT_BG }, line: { type: "none" }, objectName: `Pool box ${i}` });
    s.addText("PROSPECTS IN OUR BASE", { x: x + 0.25, y: y + 2.1, w: iw - 0.5, h: 0.2, fontSize: 8, bold: true, color: MUTED, charSpacing: 1, margin: 0, isTextBox: true, objectName: `Pool head ${i}` });
    s.addText(c[5], { x: x + 0.25, y: y + 2.3, w: iw - 0.5, h: 0.38, fontSize: 20, bold: true, color: BLUE, margin: 0, valign: "middle", isTextBox: true, objectName: `Pool ${i}` });
    s.addText(c[6], { x: x + 0.25, y: y + 2.68, w: iw - 0.5, h: 0.28, fontSize: 9, color: TXT, margin: 0, valign: "top", isTextBox: true, objectName: `Pool label ${i}` });
  });
  s.addText([{ text: "BOARD ASKS  ", options: { bold: true, color: C.accent1, charSpacing: 1, fontSize: 9.5 } },
    { text: "Risk appetite for underserved lending  ·  Shift sign-up budget to activation  ·  One scorecard  ·  Balances, revenue and losses in the data room", options: { fontSize: 10.5, color: TXT } }],
    { x: 0.5, y: 6.47, w: 12.33, h: 0.3, margin: 0, valign: "middle", isTextBox: true, objectName: "Board asks" });
}

// ---------- Slide 3: prioritization table (connects to business cases)
{
  const s = pres.addSlide({ masterName: "QUESTION", sectionTitle: "Storyline" });
  s.addText("3 / 3  |  PRIORITIZATION", { placeholder: "tracker" });
  s.addText("Prioritization for upcoming initiatives: the customers are already in our base; financials follow the business cases", { placeholder: "title" });
  s.addText(POOL_NOTE + ". xx = to be sized; figures to be filled in once business cases are complete. Icons: Font Awesome (open licence)", { placeholder: "foot" });
  const cols = [["Upcoming initiatives", "in priority order", 0.5, 3.0], ["Persona", "target archetypes", 3.6, 3.15], ["Prospect customers", "in existing CLICX base", 6.85, 2.15],
    ["Key financial outcome", "annual revenue", 9.1, 1.95], ["Payback & IRR", "break-even point", 11.15, 1.68]];
  cols.forEach((c, k) => s.addText([{ text: c[0], options: { bold: true, fontSize: 12, color: NAVY, breakLine: true } }, { text: c[1], options: { fontSize: 9.5, color: MUTED } }],
    { x: c[2], y: 1.5, w: c[3], h: 0.5, margin: 0, valign: "bottom", isTextBox: true, objectName: `Head ${k}` }));
  s.addShape(pres.shapes.LINE, { x: 0.5, y: 2.08, w: 12.33, h: 0, line: { color: NAVY, width: 1.5 }, objectName: "Header rule" });
  const rows = [
    ["P1", "Salary-Bridged Financing", [["p_corp", "Corporate crew"], ["p_sme", "SME staffer"], ["p_helper", "Personal helper"]], "607K", "salaried & wage earners · 28%", "Loan outstanding per year"],
    ["P2", "Credit Card & Apple Pay", [["p_earner", "Salary earners"], ["p_owner", "Business owners"]], "984K", "salary earners & owners · 46%", "by year xx"],
    ["P3", "Insurance e-Brokerage", [["p_rider", "Riders / drivers"], ["p_blue", "Blue collars"], ["p_white", "White collars"]], "836K", "riders, blue & white collars · 39%", "by year xx"],
    ["P4", "Digital Investment", [["p_wealth", "Aspiring wealth"], ["p_tech", "Tech savvy"]], "533K", "tech-savvy customers · 25%", "by year xx"],
  ];
  const rh = 1.1, y0 = 2.15;
  rows.forEach((r, k) => {
    const y = y0 + k * rh;
    if (k) s.addShape(pres.shapes.LINE, { x: 0.5, y, w: 12.33, h: 0, line: { color: LGREY, width: 0.75 }, objectName: `Row rule ${k}` });
    s.addShape(pres.shapes.OVAL, { x: 0.5, y: y + 0.3, w: 0.48, h: 0.48, fill: { color: BLUE }, line: { type: "none" }, objectName: `P badge ${k}` });
    s.addText(r[0], { x: 0.5, y: y + 0.3, w: 0.48, h: 0.48, fontSize: 12, bold: true, color: "FFFFFF", align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: `P text ${k}` });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 1.1, y: y + 0.3, w: 2.35, h: 0.48, rectRadius: 0.24, fill: { color: "FFFFFF" }, line: { color: BLUE, width: 1.25, dashType: "dash" }, objectName: `Initiative pill ${k}` });
    s.addText(r[1], { x: 1.1, y: y + 0.3, w: 2.35, h: 0.48, fontSize: 12, bold: true, color: BLUE, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: `Initiative name ${k}` });
    r[2].forEach((p, j) => {
      const px = 3.6 + (j % 2) * 1.6, py = y + (r[2].length > 2 ? 0.12 : 0.35) + Math.floor(j / 2) * 0.46;
      s.addShape(pres.shapes.OVAL, { x: px, y: py, w: 0.38, h: 0.38, fill: { color: j % 2 ? BLUE : CYAN }, line: { type: "none" }, objectName: `Persona icon bg ${k}-${j}` });
      s.addImage({ path: IC(p[0] + ".png"), x: px + 0.09, y: py + 0.09, w: 0.2, h: 0.2, altText: p[1] + " icon", objectName: `Persona icon ${k}-${j}` });
      s.addText(p[1], { x: px + 0.45, y: py, w: 1.1, h: 0.38, fontSize: 10, color: NAVY, bold: true, margin: 0, valign: "middle", isTextBox: true, objectName: `Persona ${k}-${j}` });
    });
    if (r[2].length === 2) { /* keep two personas on one line, centred vertically */ }
    s.addText([{ text: r[3], options: { fontSize: 22, bold: true, color: CYAN, breakLine: true } }, { text: r[4], options: { fontSize: 9.5, color: TXT } }],
      { x: 6.85, y: y + 0.12, w: 2.15, h: 0.85, margin: 0, valign: "middle", isTextBox: true, objectName: `Prospects ${k}` });
    s.addText([{ text: "THB xx M", options: { fontSize: 20, bold: true, color: GREY, breakLine: true } }, { text: r[5], options: { fontSize: 9.5, color: MUTED } }],
      { x: 9.1, y: y + 0.12, w: 1.95, h: 0.85, margin: 0, valign: "middle", isTextBox: true, objectName: `Financial ${k}` });
    s.addText([{ text: "xx years", options: { fontSize: 20, bold: true, color: GREY, breakLine: true } }, { text: "break-even · IRR xx%", options: { fontSize: 9.5, color: MUTED } }],
      { x: 11.15, y: y + 0.12, w: 1.68, h: 0.85, margin: 0, valign: "middle", isTextBox: true, objectName: `Payback ${k}` });
  });
  s.addShape(pres.shapes.LINE, { x: 0.5, y: y0 + 4 * rh, w: 12.33, h: 0, line: { color: LGREY, width: 0.75 }, objectName: "Bottom rule" });
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
