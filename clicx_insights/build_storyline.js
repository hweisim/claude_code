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
  s.addText("1 / 4  |  WHAT WE KNOW: CUSTOMERS", { placeholder: "tracker" });
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
  panel(x, y1, W3, h1, 1, "How much do we rely on the consortium, and how?", "96% come from partners and 90% sign up through partner channels", "96%", "of customers are AIS, KTB or OR customers", 1);
  bars(x + 1.82, y1 + 1.2, W3 - 2.0, [["Paotang", 0.80, "80%", BLUE], ["myAIS", 0.06, "6%", CYAN], ["Dipchip", 0.04, "4%", CYAN], ["NDID", 0.10, "10%", GREY]], 1, { head: "Sign-up channel", lw: 0.7, rh: 0.21 });
  s.addText("73% of payers came via AIS or OR", { x: x + 1.82, y: y1 + 2.06, w: W3 - 2.0, h: 0.18, fontSize: 8, italic: true, color: MUTED, margin: 0, isTextBox: true, objectName: "AIS headroom" });
  // 2 targets
  x = 0.5 + W3 + G;
  panel(x, y1, W3, h1, 2, "Are we reaching the customers we planned?", "Yes: 40% have no regular income and 82% of applicants are underserved", "40%", "are self-employed, with no regular income", 2);
  bars(x + 1.82, y1 + 1.25, W3 - 2.0, [["Self-employed", 0.40, "40%", CYAN], ["White-collar", 0.38, "38%", BLUE], ["Blue-collar", 0.22, "22%", NAVY]], 2, { head: "Customers by type of work" });
  s.addText("Joined: tech-savvy 8.5%, multi 2.4%", { x: x + 1.82, y: y1 + 2.06, w: W3 - 2.0, h: 0.18, fontSize: 8, italic: true, color: MUTED, margin: 0, isTextBox: true, objectName: "Target note" });
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
  panel(x, y2, W2, h2, 5, "What value do we create for each shareholder?", "We reach customers KTB doesn't, and bring AIS and OR new customers and spend", "81%", "of our 164K borrowers never had a KTB Jai-pump loan", 5);
  const sh = [["logo_ktb.jpeg", 2.58, "65% of CLICX depositors are not active with KTB; 1.1M shared customers were KTB-inactive"],
    ["logo_ais.png", 2.74, "0.6M CLICX customers not on AIS yet; ฿47M paid to AIS in Oct, down 40% after the promotion"],
    ["logo_or.png", 1.93, "฿19M paid to OR in Oct; 173K active customers without BLUE+"]];
  sh.forEach((r, k) => {
    const yy = y2 + 1.0 + k * 0.42;
    s.addImage({ path: IC(r[0]), x: x + 1.85, y: yy + 0.06, h: 0.22, w: 0.22 * r[1], altText: "Partner logo", objectName: `Value logo ${k}` });
    s.addText(r[2], { x: x + 2.55, y: yy, w: W2 - 2.7, h: 0.38, fontSize: 9.5, color: TXT, margin: 0, valign: "middle", isTextBox: true, objectName: `Value text ${k}` });
  });
}

// ---------- Slide 2: what we know, payments
{
  const s = pres.addSlide({ masterName: "QUESTION", sectionTitle: "Storyline" });
  s.addText("2 / 4  |  WHAT WE KNOW: USAGE AND PAYMENTS", { placeholder: "tracker" });
  s.addText("Use follows life stage and payments follow steady income; promotions alone don't build the habit, and AIS prepaid is the biggest gap", { placeholder: "title" });
  s.addText("Depositor = balance over ฿100; saver = Save Max/More over ฿10K. Ever partner = paid AIS or PTT OR. Category figures cover Scan and Pay to other merchants. Source: Partner Payment Story v1 and payment-by-age files (1.51M customers, 31 Oct 2026). Full tables in appendix A1–A2", { placeholder: "foot" });
  const panel = (x, y, w, h, tag, q, ans, big, bigLbl, n) => {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.08, fill: { color: C.background2 }, line: { type: "none" }, objectName: `Pay panel ${n}` });
    s.addText([{ text: tag + "  ", options: { color: CYAN, bold: true } }, { text: q, options: { color: MUTED } }],
      { x: x + 0.18, y: y + 0.1, w: w - 0.36, h: 0.26, fontSize: 10, margin: 0, valign: "middle", isTextBox: true, objectName: `Pay q ${n}` });
    s.addText(ans, { x: x + 0.18, y: y + 0.36, w: w - 0.36, h: 0.5, fontSize: 12, bold: true, color: NAVY, margin: 0, valign: "top", isTextBox: true, objectName: `Pay ans ${n}` });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + 0.18, y: y + 0.92, w: 1.45, h: h - 1.07, rectRadius: 0.08, fill: { color: STAT_BG }, line: { type: "none" }, objectName: `Pay stat box ${n}` });
    s.addText(big, { x: x + 0.26, y: y + 0.98, w: 1.32, h: 0.5, fontSize: big.length > 5 ? 20 : 24, bold: true, color: BLUE, margin: 0, valign: "middle", isTextBox: true, objectName: `Pay stat ${n}` });
    s.addText(bigLbl, { x: x + 0.26, y: y + 1.48, w: 1.32, h: h - 1.6, fontSize: 9, color: TXT, margin: 0, valign: "top", isTextBox: true, objectName: `Pay stat lbl ${n}` });
  };
  const bars = (x, y, w, rows, n, o = {}) => {
    const lw = o.lw || 0.9, vw = o.vw || 0.62, bw = w - lw - vw - 0.06, rh = o.rh || 0.26;
    if (o.head) s.addText(o.head, { x, y: y - 0.24, w, h: 0.22, fontSize: 8.5, bold: true, color: MUTED, margin: 0, isTextBox: true, objectName: `Pay bars head ${n}` });
    rows.forEach((r, k) => {
      const yy = y + k * rh;
      if (r[4]) s.addImage({ path: IC(r[4]), x, y: yy + 0.04, h: 0.18, w: 0.18 * r[5], altText: r[0] + " logo", objectName: `Pay logo ${n}-${k}` });
      else s.addText(r[0], { x, y: yy, w: lw, h: rh, fontSize: 9, color: TXT, margin: 0, valign: "middle", isTextBox: true, objectName: `Pay label ${n}-${k}` });
      s.addShape(pres.shapes.RECTANGLE, { x: x + lw, y: yy + rh * 0.3, w: bw, h: rh * 0.4, fill: { color: "E3E8EE" }, line: { type: "none" }, objectName: `Pay track ${n}-${k}` });
      s.addShape(pres.shapes.RECTANGLE, { x: x + lw, y: yy + rh * 0.3, w: Math.max(0.03, bw * r[1]), h: rh * 0.4, fill: { color: r[3] }, line: { type: "none" }, objectName: `Pay bar ${n}-${k}` });
      s.addText(r[2], { x: x + lw + bw + 0.05, y: yy, w: vw, h: rh, fontSize: 9, bold: true, color: NAVY, margin: 0, valign: "middle", isTextBox: true, objectName: `Pay val ${n}-${k}` });
    });
  };
  const W = 4.01, G = 0.15, H = 2.48, ys = [1.5, 4.1];
  const xs = [0.5, 0.5 + W + G, 0.5 + 2 * (W + G)];
  const bx = (i) => xs[i % 3] + 1.78, by = (i) => ys[Math.floor(i / 3)] + 1.18, bw = W - 1.95;
  // 1 holding by age
  panel(xs[0], ys[0], W, H, "Q3", "What do customers hold, by age?", "80% hold nothing; teens pay, 20–59s borrow, 60+ save", "79.5%", "of customers hold no active product", 1);
  bars(bx(0), by(0), bw, [["Teens: pay only", 0.096 / 0.2, "9.6%", CYAN], ["20–29: borrow", 0.073 / 0.2, "7.3%", BLUE], ["60–69: deposit", 0.142 / 0.2, "14.2%", NAVY]], 1, { head: "Main product for each age group", lw: 1.05, vw: 0.48, rh: 0.27 });
  s.addText("No loans under 20 or over 59 (age limits)", { x: bx(0), y: by(0) + 0.84, w: bw, h: 0.2, fontSize: 8, italic: true, color: MUTED, margin: 0, isTextBox: true, objectName: "Age note" });
  // 2 what they pay for
  panel(xs[1], ys[0], W, H, "Q3", "What do customers pay for?", "Mostly everyday Scan and Pay, not partner bills", "฿408M", "paid in October, +11% vs August", 2);
  bars(bx(1), by(1), bw, [["Scan & Pay", 0.84, "84%", CYAN], ["AIS", 0.11, "11%", BLUE, "logo_ais.png", 2.74], ["OR", 0.05, "5%", NAVY, "logo_or.png", 1.93]], 2, { head: "Share of October spend", lw: 0.72, vw: 0.42, rh: 0.3 });
  // 3 who pays
  panel(xs[2], ys[0], W, H, "Q3", "Who pays with CLICX?", "Merchants and salaried pay most; blue-collar and gig workers least", "8.4%", "of all customers paid in October", 3);
  bars(bx(2), by(2) - 0.06, bw, [["Online merchant", 0.102 / 0.12, "10.2%", CYAN], ["Salaried", 0.094 / 0.12, "9.4%", CYAN], ["Freelance", 0.073 / 0.12, "7.3%", GREY], ["Rider", 0.071 / 0.12, "7.1%", NAVY], ["Blue-collar", 0.064 / 0.12, "6.4%", NAVY]], 3, { head: "% who paid in October, by job", lw: 0.95, vw: 0.48, rh: 0.21 });
  // 4 promotion
  panel(xs[0], ys[1], W, H, "Q3", "Did the payment promotion build a habit?", "Partly: a third of payers left, but those who stayed pay more", "59.9K", "paid only in August and never came back", 4);
  bars(bx(3), by(3), bw, [["Aug", 1.0, "164K", GREY], ["Sep", 0.80, "130K", BLUE], ["Oct", 0.77, "127K", BLUE]], 4, { head: "Customers paying each month", lw: 0.4, vw: 0.5, rh: 0.26 });
  s.addText("Stayers: ฿2.1K → ฿3.0K each", { x: bx(3), y: by(3) + 0.82, w: bw, h: 0.22, fontSize: 9, bold: true, color: CYAN, margin: 0, isTextBox: true, objectName: "Stayers spend" });
  // 5 biggest gaps
  panel(xs[1], ys[1], W, H, "Q1", "Where are the biggest payment gaps?", "AIS prepaid customers, and OR customers who don't pay", "516.5K", "AIS prepaid customers; +22K payers a month at the postpaid rate", 5);
  bars(bx(4), by(4) - 0.05, bw, [["Postpaid", 0.128 / 0.14, "12.8%", NAVY], ["Prepaid", 0.084 / 0.14, "8.4%", CYAN], ["OR", 0.113 / 0.14, "11.3%", BLUE], ["Not OR", 0.074 / 0.14, "7.4%", GREY]], 5, { head: "% who paid in October", lw: 0.62, vw: 0.5, rh: 0.25 });
  // 6 product leads
  panel(xs[2], ys[1], W, H, "Q4", "What do non-partner payers spend on?", "Repaying other banks and investing: leads for a card and investment", "26%", "of their spend repays other banks' cards and loans, ฿4.3K each", 6);
  bars(bx(5), by(5) - 0.05, bw, [["Repay banks", 0.26 / 0.3, "26%", NAVY], ["Crypto, gold, stocks", 0.10 / 0.3, "10%", CYAN], ["Local shops", 0.25 / 0.3, "25%", GREY]], 6, { head: "Share of their spend", lw: 1.15, vw: 0.38, rh: 0.27 });
  s.addText("Partner payers: 20%, 9%, 27%", { x: bx(5), y: by(5) + 0.78, w: bw, h: 0.22, fontSize: 9, color: MUTED, margin: 0, isTextBox: true, objectName: "Partner comparison" });
}

// ---------- Slide 2: what we will do, in two parts
const POOL_NOTE = "Prospect pools = CLICX customers by occupation and profile (2.14M customers, Sep 2026); pools overlap";
{
  const s = pres.addSlide({ masterName: "QUESTION", sectionTitle: "Storyline" });
  s.addText("3 / 4  |  WHAT WE WILL DO", { placeholder: "tracker" });
  s.addText("We will grow in two ways: keep acquiring through our partners, and activate the 1.8M customers we already have", { placeholder: "title" });
  s.addText(POOL_NOTE + ". Icons: Font Awesome (open licence). " + SRC, { placeholder: "foot" });
  const pill = (x, y, w, n, t, col, obj) => {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h: 0.42, rectRadius: 0.21, fill: { color: col }, line: { type: "none" }, objectName: obj });
    s.addText([{ text: n + "   ", options: { bold: true, fontSize: 14, color: "FFFFFF" } }, { text: t, options: { bold: true, fontSize: 12.5, color: "FFFFFF", charSpacing: 1 } }],
      { x: x + 0.2, y, w: w - 0.3, h: 0.42, valign: "middle", margin: 0, isTextBox: true, objectName: obj + " text" });
  };
  const circIcon = (x, y, d, icon, col, alt, obj) => {
    s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: col }, line: { type: "none" }, objectName: obj + " bg" });
    s.addImage({ path: IC(icon), x: x + d * 0.24, y: y + d * 0.24, w: d * 0.52, h: d * 0.52, altText: alt, objectName: obj });
  };

  // ---- Part 1: continue to acquire
  const LX = 0.5, LW = 3.95;
  pill(LX, 1.5, LW, "1", "CONTINUE TO ACQUIRE", NAVY, "Part 1 pill");
  s.addText("New customers from partner bases  ·  from Q1, Q2, Q5", { x: LX, y: 1.97, w: LW, h: 0.24, fontSize: 9.5, color: MUTED, margin: 0, isTextBox: true, objectName: "Part 1 sub" });
  const acq = [
    ["logo_ais.png", 2.74, "3.7%", "of AIS's ~32M base joined", "Sign up through myAIS and AIS shops; put bill pay and top-up in the app"],
    ["logo_or.png", 1.93, "7.4%", "of OR's base joined, the best rate", "OR customers pay 1.5× more; 344K didn't pay in Oct. Use BLUE+ and fuel offers"],
    ["logo_ktb.jpeg", 2.58, "1.9M", "turned down by Jai-pump a year", "Work the Paotang reject list: 24% apply; focus on under-40s and riders"],
  ];
  acq.forEach((a, k) => {
    const y = 2.28 + k * 0.92;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: LX, y, w: LW, h: 0.84, rectRadius: 0.08, fill: { color: C.background2 }, line: { type: "none" }, objectName: `Acquire ${k}` });
    s.addImage({ path: IC(a[0]), x: LX + 0.15, y: y + 0.12, h: 0.24, w: 0.24 * a[1], altText: "Partner logo", objectName: `Acquire logo ${k}` });
    s.addText([{ text: a[2] + "  ", options: { bold: true, fontSize: 15, color: BLUE } }, { text: a[3], options: { fontSize: 9, color: MUTED } }],
      { x: LX + 1.0, y: y + 0.06, w: LW - 1.1, h: 0.34, margin: 0, valign: "middle", isTextBox: true, objectName: `Acquire stat ${k}` });
    s.addText(a[4], { x: LX + 0.15, y: y + 0.42, w: LW - 0.3, h: 0.38, fontSize: 9.5, color: TXT, margin: 0, valign: "top", isTextBox: true, objectName: `Acquire text ${k}` });
  });
  // target personas
  const ty = 2.28 + 3 * 0.92;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: LX, y: ty, w: LW, h: 0.92, rectRadius: 0.08, fill: { color: STAT_BG }, line: { type: "none" }, objectName: "Target card" });
  circIcon(LX + 0.15, ty + 0.17, 0.5, "target.png", CYAN, "Target icon", "Target icon");
  s.addText([{ text: "Who to target", options: { bold: true, fontSize: 11, color: NAVY, breakLine: true } },
    { text: "Look-alikes of tech-savvy 25–39s (8.5% join) and working Gen Z (6.6%); give multi-bank users a reason to switch (2.4%)", options: { fontSize: 9.5, color: TXT } }],
    { x: LX + 0.78, y: ty + 0.06, w: LW - 0.9, h: 0.82, margin: 0, valign: "middle", isTextBox: true, objectName: "Target text" });
  // board asks (compact)
  const by = ty + 1.02;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: LX, y: by, w: LW, h: 6.75 - by, rectRadius: 0.08, fill: { color: C.text2 }, line: { type: "none" }, objectName: "Asks box" });
  s.addText([{ text: "BOARD ASKS", options: { bold: true, fontSize: 9, color: C.accent6, charSpacing: 1, breakLine: true } },
    { text: "Risk appetite · budget shift to activation · one scorecard · balances and revenue data", options: { fontSize: 9.5, color: "FFFFFF" } }],
    { x: LX + 0.15, y: by + 0.05, w: LW - 0.3, h: 6.75 - by - 0.1, margin: 0, valign: "middle", isTextBox: true, objectName: "Asks text" });

  // ---- Part 2: activate current customers
  const RX = 4.6, RW = 8.23;
  pill(RX, 1.5, RW, "2", "ACTIVATE CURRENT CLICX CUSTOMERS", CYAN, "Part 2 pill");
  s.addText("1.8M idle customers (694K no product, 1.1M not active) to turn into active, paying users  ·  from Q3, Q4", { x: RX, y: 1.97, w: RW, h: 0.24, fontSize: 9.5, color: MUTED, margin: 0, isTextBox: true, objectName: "Part 2 sub" });
  // four activation levers
  const lev = [["deposit.png", "Get money in first", "51% active with money in vs 4% without", BLUE],
    ["enabler.png", "Fix approvals and risk", "Gig workers: 5–7% approved; 20% early default", NAVY],
    ["prepaid.png", "Push AIS prepaid", "516.5K prepaid: 8.4% pay vs 12.8% postpaid; top-up first", CYAN],
    ["reward.png", "Reward the second month", "59.9K paid only in the promotion month; keep AIS and OR promos as the way in", BLUE]];
  const lw2 = (RW - 0.3) / 4;
  lev.forEach((l, k) => {
    const x = RX + k * (lw2 + 0.1);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 2.28, w: lw2, h: 0.98, rectRadius: 0.08, fill: { color: C.background2 }, line: { type: "none" }, objectName: `Lever ${k}` });
    circIcon(x + 0.1, 2.36, 0.38, l[0], l[3], l[1] + " icon", `Lever icon ${k}`);
    s.addText(l[1], { x: x + 0.55, y: 2.34, w: lw2 - 0.6, h: 0.42, fontSize: 10, bold: true, color: NAVY, margin: 0, valign: "middle", isTextBox: true, objectName: `Lever title ${k}` });
    s.addText(l[2], { x: x + 0.1, y: 2.8, w: lw2 - 0.18, h: 0.44, fontSize: 8.5, color: TXT, margin: 0, valign: "top", isTextBox: true, objectName: `Lever text ${k}` });
  });
  // initiatives P1-P4 (2x2)
  s.addText("PRODUCTS FOR CUSTOMERS ALREADY IN OUR BASE  ·  in priority order", { x: RX, y: 3.36, w: RW, h: 0.24, fontSize: 9.5, bold: true, color: MUTED, charSpacing: 1, margin: 0, isTextBox: true, objectName: "Initiatives label" });
  const ini = [
    ["P1", "salary", "Salary-Bridged Financing", "Corporate & SME staff · Helpers", "Salaried workers get approved most (11%)", "607K", "salaried and wage earners · 28%"],
    ["P2", "card", "Credit Card & Apple Pay", "Salary earners · Business owners", "Payers repay other banks' cards and loans: 26% of spend, ฿4.3K each", "984K", "salary earners and owners · 46%"],
    ["P3", "insure", "Insurance e-Brokerage", "Riders · Blue & white collars", "Riders rarely get a loan; insurance keeps them", "836K", "riders, blue and white collars · 39%"],
    ["P4", "invest", "Digital Investment", "Tech savvy · 50+ savers", "10% of payment spend goes to crypto, gold and stocks; savers hold 5.5B THB", "533K", "tech-savvy customers · 25%"],
  ];
  const cw = (RW - 0.12) / 2, chh = 1.5;
  ini.forEach((c, i) => {
    const x = RX + (i % 2) * (cw + 0.12), y = 3.66 + Math.floor(i / 2) * (chh + 0.1);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: cw, h: chh, rectRadius: 0.08, fill: { color: C.background2 }, line: { type: "none" }, objectName: `Initiative ${i}` });
    circIcon(x + 0.12, y + 0.14, 0.48, c[1] + ".png", CYAN, c[2] + " icon", `Initiative icon ${i}`);
    s.addText([{ text: c[0] + "  ", options: { bold: true, color: BLUE } }, { text: c[2], options: { bold: true, color: NAVY } }],
      { x: x + 0.7, y: y + 0.12, w: cw - 2.2, h: 0.5, fontSize: 12, margin: 0, valign: "middle", isTextBox: true, objectName: `Initiative name ${i}` });
    s.addText([{ text: "FOR  ", options: { bold: true, color: CYAN, fontSize: 8, charSpacing: 1 } }, { text: c[3], options: { bold: true, color: NAVY, fontSize: 9 } }],
      { x: x + 0.12, y: y + 0.7, w: cw - 1.65, h: 0.28, margin: 0, valign: "middle", isTextBox: true, objectName: `Initiative persona ${i}` });
    s.addText(c[4], { x: x + 0.12, y: y + 1.0, w: cw - 1.65, h: 0.42, fontSize: 9, color: MUTED, margin: 0, valign: "top", isTextBox: true, objectName: `Initiative why ${i}` });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + cw - 1.45, y: y + 0.12, w: 1.33, h: chh - 0.24, rectRadius: 0.06, fill: { color: STAT_BG }, line: { type: "none" }, objectName: `Pool box ${i}` });
    s.addText([{ text: "PROSPECTS", options: { bold: true, fontSize: 7.5, color: MUTED, charSpacing: 1, breakLine: true } },
      { text: c[5], options: { bold: true, fontSize: 20, color: BLUE, breakLine: true } }, { text: c[6], options: { fontSize: 8.5, color: TXT } }],
      { x: x + cw - 1.38, y: y + 0.16, w: 1.2, h: chh - 0.32, margin: 0, valign: "middle", isTextBox: true, objectName: `Pool ${i}` });
  });
}

// ---------- Slide 3: prioritization table (connects to business cases)
{
  const s = pres.addSlide({ masterName: "QUESTION", sectionTitle: "Storyline" });
  s.addText("4 / 4  |  PRIORITIZATION", { placeholder: "tracker" });
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
    ["P4", "Digital Investment", [["p_wealth", "50+ savers"], ["p_tech", "Tech savvy"]], "533K", "tech-savvy customers · 25%", "by year xx"],
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

// ---------- Appendix A1, A2: full tables
const heat = (s, head, rows, colW, base, opts) => {
  const ABOVE = "CDEAE7", BELOW = "F4DCDC";
  const hdr = head.map((h) => ({ text: h, options: { bold: true, fontSize: 9.5, color: NAVY, align: "right", valign: "bottom", border: [{ type: "none" }, { type: "none" }, { pt: 1.25, color: NAVY }, { type: "none" }] } }));
  hdr[0].options.align = "left";
  const body = rows.map((r, ri) => {
    const last = ri === rows.length - 1;
    return r.map((v, ci) => {
      const o = { fontSize: 10, color: TXT, align: ci === 0 ? "left" : "right", bold: ci === 0 || last, border: last ? [{ pt: 1.25, color: NAVY }, { type: "none" }, { type: "none" }, { type: "none" }] : [{ type: "none" }, { type: "none" }, { pt: 0.5, color: "E3E8EE" }, { type: "none" }] };
      if (ci === 0 || ci < (opts.dataFrom || 1)) return { text: String(v), options: o };
      const small = typeof v === "string" && v.endsWith("*");
      const num = parseFloat(v);
      const b = opts.baseFor(ri, ci);
      if (small) { o.color = GREY; return { text: v, options: o }; }
      if (!opts.noColor || !opts.noColor(ci)) {
        if (b > 0 && num >= b * 1.15) o.fill = { color: ABOVE };
        else if (num <= b * 0.85) o.fill = { color: BELOW };
      }
      return { text: (opts.fmt || ((x) => x.toFixed(1) + "%"))(num), options: o };
    });
  });
  s.addTable([hdr, ...body], { x: 0.5, y: opts.y || 2.05, w: 12.33, colW, fontFace: "Arial", rowH: [0.62, ...rows.map(() => 0.42)], margin: [0.03, 0.08, 0.03, 0.08], valign: "middle", objectName: opts.name });
};
pres.addSection({ title: "Appendix" });
{
  const s = pres.addSlide({ masterName: "QUESTION", sectionTitle: "Appendix" });
  s.addText("APPENDIX A1  |  WHO PAYS · AGE × JOB · % PAID IN OCTOBER", { placeholder: "tracker" });
  s.addText("Merchants and salaried customers pay most; blue-collar and gig workers pay least, at every age", { placeholder: "title" });
  s.addText("Shaded: 15%+ above (teal) or below (rose) all customers, 8.4%. * fewer than 1,000 customers, not shaded. Salaried = corporate + government. Data as of 31 Oct 2026", { placeholder: "foot" });
  const head = ["Age (customers)", "Salaried\n575.4K", "Blue collar\n224.1K", "Offline merchant\n148.9K", "Online merchant\n94.8K", "Freelance\n217.6K", "Wage earner\n104.7K", "Rider\n58.3K", "Agriculture\n55.6K", "Other / unknown\n28.1K", "All\n1.51M"];
  const rows = [
    ["10–19 (10.9K)", "6.6*", "5.7*", "15.2*", 14.0, 11.5, 11.4, "9.0*", "14.6*", 21.2, 14.3],
    ["20–29 (395.3K)", 8.5, 6.0, 8.4, 9.5, 6.8, 7.2, 5.6, 6.5, 9.7, 7.6],
    ["30–39 (580.3K)", 9.8, 7.0, 9.9, 10.6, 7.5, 7.4, 7.1, 6.7, 9.4, 8.8],
    ["40–49 (369.3K)", 9.7, 6.3, 10.2, 10.5, 7.5, 6.8, 8.2, 6.6, 9.9, 8.7],
    ["50–59 (126.5K)", 8.8, 5.3, 9.6, 9.7, 7.3, 6.1, 8.6, 6.6, 11.0, 8.1],
    ["60–69 (22.2K)", 7.5, "3.3*", 8.0, 7.0, 8.1, 6.5, "9.0*", 5.3, 11.7, 8.0],
    ["70+ (2.9K)", "6.5*", "8.3*", "14.5*", "8.9*", "11.7*", "10.4*", "18.2*", "8.3*", "10.3*", 10.4],
    ["All (1.51M)", 9.4, 6.4, 9.7, 10.2, 7.3, 7.1, 7.1, 6.6, 11.5, 8.4],
  ];
  s.addText([{ text: "STABLE INCOME", options: { bold: true, color: MUTED } }], { x: 2.05, y: 1.72, w: 2.1, h: 0.26, fontSize: 9.5, align: "center", charSpacing: 1, margin: 0, isTextBox: true, objectName: "Stable label" });
  s.addText([{ text: "UNSTABLE INCOME", options: { bold: true, color: CYAN } }], { x: 4.15, y: 1.72, w: 6.3, h: 0.26, fontSize: 9.5, align: "center", charSpacing: 1, margin: 0, isTextBox: true, objectName: "Unstable label" });
  heat(s, head, rows, [1.55, 1.05, 1.05, 1.1, 1.1, 1.05, 1.05, 1.0, 1.05, 1.15, 1.18], 8.4, { y: 2.0, name: "Who pays table", baseFor: () => 8.4 });
}
{
  const s = pres.addSlide({ masterName: "QUESTION", sectionTitle: "Appendix" });
  s.addText("APPENDIX A2  |  PRODUCT SEGMENTS BY AGE · SHARE OF EACH AGE GROUP", { placeholder: "tracker" });
  s.addText("Teens just pay, 20–59s borrow, and customers 60+ save; 80% of every age group hold nothing active", { placeholder: "title" });
  s.addText("Rows add to 100%. Shaded: 15%+ above (teal) or below (rose) all customers; No segment not shaded. Saver = Save Max/More over ฿10K; depositor = other balance over ฿100. No loans under 20 or over 59 (age limits). Groups 60+ are small (25.1K). As of 31 Oct 2026", { placeholder: "foot" });
  const head = ["Age", "Customers", "Borrower only", "Borrower + transactor", "Depositor + borrower", "Depositor + transactor", "Depositor only", "Saver", "Transactor only", "No segment"];
  const rows = [
    ["10–19", "10.9K", 0.0, 0.0, 0.0, 4.6, 6.0, 0.7, 9.6, 79.2],
    ["20–29", "395.3K", 7.3, 2.6, 1.0, 1.5, 1.3, 0.5, 2.8, 82.9],
    ["30–39", "580.3K", 7.1, 3.1, 1.4, 2.0, 2.5, 1.1, 2.6, 80.1],
    ["40–49", "369.3K", 6.8, 2.7, 1.8, 2.3, 4.4, 2.0, 2.4, 77.5],
    ["50–59", "126.5K", 5.9, 1.5, 2.0, 3.0, 7.7, 3.8, 2.2, 73.9],
    ["60–69", "22.2K", 0.0, 0.0, 0.0, 4.4, 14.2, 9.4, 2.2, 69.8],
    ["70+", "2.9K", 0.0, 0.0, 0.0, 5.4, 18.7, 15.0, 3.1, 57.7],
    ["All", "1.51M", 6.8, 2.7, 1.4, 2.1, 3.3, 1.5, 2.6, 79.5],
  ];
  const all = rows[rows.length - 1];
  s.addText("BORROWER", { x: 2.95, y: 1.72, w: 2.6, h: 0.26, fontSize: 9.5, bold: true, color: MUTED, align: "center", charSpacing: 1, margin: 0, isTextBox: true, objectName: "Borrower label" });
  s.addText("DEPOSITOR OR SAVER", { x: 5.55, y: 1.72, w: 5.2, h: 0.26, fontSize: 9.5, bold: true, color: CYAN, align: "center", charSpacing: 1, margin: 0, isTextBox: true, objectName: "Depositor label" });
  heat(s, head, rows, [1.0, 1.2, 1.25, 1.35, 1.3, 1.35, 1.3, 1.25, 1.3, 1.03], 0, { y: 2.0, name: "Segments by age table", dataFrom: 2, baseFor: (ri, ci) => parseFloat(all[ci]), noColor: (ci) => ci === 9 || false });
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
