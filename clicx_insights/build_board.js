// CLICX customer insights deck — question-led storyline, data from
// dtr_strategy2027_v4.pptx and CLICX_5-Year_Business_Plan_V1-Oct_2027.pptx
const pptxgen = require("pptxgenjs");
const { applyTheme } = require(process.env.SKILL + "/scripts/apply_theme.js");

const OUT = process.argv[2] || "CLICX_Customer_Insights.pptx";

const THEME = {
  name: "CLICX Insights",
  headFontFace: "Arial",
  bodyFontFace: "Arial",
  colors: {
    dk1: "1A1A1A", lt1: "FFFFFF", dk2: "051C2C", lt2: "F2F4F7",
    accent1: "2251FF", accent2: "00A9F4", accent3: "051C2C",
    accent4: "A2AAAD", accent5: "D9475C", accent6: "7FD1F7",
    hlink: "2251FF", folHlink: "6E7C87",
  },
};
// Hex mirrors for hex-only chart options
const NAVY = "051C2C", BLUE = "2251FF", CYAN = "00A9F4", GREY = "A2AAAD",
  LGREY = "D5DADF", RED = "D9475C", LCYAN = "7FD1F7", TXT = "1A1A1A", MUTED = "5F6B76", TINT = "F2F4F7";

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "CLICX existing customers: board summary";
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
    { placeholder: { options: { name: "foot", type: "body", x: 0.5, y: 6.92, w: 11.2, h: 0.3, fontSize: 10, color: "6E7C87", valign: "top", margin: 0 }, text: "" } },
    { image: { path: __dirname + "/logo.png", x: 11.63, y: 0.2, w: 1.2, h: 0.3, altText: "CLICX logo" } },
  ],
  slideNumber: { x: 12.33, y: 6.95, w: 0.5, h: 0.3, fontSize: 10, color: "6E7C87", align: "right" },
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
// Board summary: existing customers' product holding, balance and payments
const SRC = "Source: CLICX product and payment segments, 1 Oct 2026 (1.50M customers with an account)";
pres.addSection({ title: "Board summary" });
{
  const s = qSlide("Board summary", "BOARD SUMMARY  (1/4)  |  WHAT EXISTING CUSTOMERS DO",
    "Do our 1.5 million existing customers actually use CLICX?",
    "Only 1 in 6 do. Most never put money in, and paying users fell 29% since August. Putting money in and getting a loan are what turn sign-ups into users",
    SRC);
  const c = { x: 0.5, y: 2.35, h: 4.35 };
  s.addChart(pres.charts.BAR, [{ name: "Customers, K", labels: ["Zero balance, no use", "Money only, under 1K", "Money only, 1K–20K", "Borrows only", "Pays only", "Borrows and pays", "Saves only (20K+)", "Saves and more"], values: [1058, 166, 20, 117, 66, 47, 22, 5] }], {
    x: c.x, y: c.y, w: 2.9, h: c.h, barDir: "bar", ...AX, ...quiet(), ...chartTitle("What they hold, thousand customers"),
    chartColors: [RED, CYAN, CYAN, NAVY, NAVY, BLUE, NAVY, NAVY], showValue: true, dataLabelPosition: "outEnd", dataLabelColor: TXT, dataLabelFontSize: 10,
    dataLabelFormatCode: "#,##0", showLegend: false, barGapWidthPct: 35, valAxisMaxVal: 1400, catAxisOrientation: "maxMin", catAxisLabelFontSize: 10, objectName: "Product holding" });
  s.addChart(pres.charts.DOUGHNUT, [{ name: "Customers", labels: ["Never put money in (976K)", "Put money in (525K)"], values: [976, 525] }], {
    x: 3.45, y: c.y, w: 2.6, h: c.h, ...AX, ...chartTitle("Have they put money in?"),
    chartColors: [RED, BLUE], holeSize: 60, showPercent: true, showValue: false, dataLabelColor: "FFFFFF", dataLabelFontSize: 12,
    showLegend: true, legendPos: "b", legendFontSize: 10, legendColor: MUTED, objectName: "First deposit" });
  s.addText([
    { text: "1.50M", options: { bold: true, fontSize: 18, color: C.text2, breakLine: true } },
    { text: "accounts", options: { fontSize: 10, color: MUTED } },
  ], { x: 4.2, y: 4.05, w: 1.1, h: 0.6, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: "Accounts total" });
  s.addChart(pres.charts.BAR, [{ name: "Paying users, K", labels: ["Aug", "Sep", "Oct"], values: [164, 130, 117] }], {
    x: 6.2, y: c.y, w: 2.7, h: c.h, barDir: "col", ...AX, ...quiet(), ...chartTitle("Customers paying with CLICX, thousand"),
    chartColors: [NAVY, NAVY, RED], showValue: true, dataLabelPosition: "outEnd", dataLabelColor: TXT, dataLabelFontSize: 12,
    dataLabelFormatCode: "#,##0", showLegend: false, barGapWidthPct: 50, valAxisMaxVal: 200, catAxisLabelFontSize: 11, objectName: "Paying users trend" });
  insights(s, [
    ["Money in is the dividing line", "22% of people who put money in pay with CLICX, vs 0.1% who never did"],
    ["A loan builds the habit", "28.5% of borrowers pay in the app, vs 5.2% of everyone else"],
    ["A 'no' loses the customer", "1.13M applied for a loan and never got one. Only 4.3% of them pay"],
    ["Two ways to pay stick", "People who scan and pay bills stop half as often (24% vs 41–48%)"],
  ], undefined, undefined, undefined, undefined, "What drives usage");
}
{
  const s = qSlide("Board summary", "BOARD SUMMARY  (2/4)  |  HOW PRODUCTS CONNECT",
    "How are money in, loans and payments connected?",
    "Tightly: a customer who has any one of them is 2.5–3.7x more likely to have the other two. But only 35% of accounts ever get money in, so most customers never start the chain",
    "x = times as likely as the average customer. Balance rows use current balance; a 100-baht cut is not in the data yet. Ladder % = share of the step before; main bank = regular income in and payments out. Source: CLICX product and payment segments, 1 Oct 2026 (1.50M accounts); registration funnel from CLICX 5-Year Business Plan");
  // Connection matrix
  s.addText("Has one product: how likely are the others?", { x: 0.5, y: 2.4, w: 4.5, h: 0.32, fontSize: 13, bold: true, color: C.text2, margin: 0, isTextBox: true, objectName: "Matrix title" });
  const H = (t) => ({ text: t, options: { bold: true, color: "FFFFFF", fill: { color: NAVY }, fontSize: 10.5, align: "center" } });
  const rowH = (t) => ({ text: t, options: { bold: true, fontSize: 11, color: NAVY } });
  const base = (t) => ({ text: t, options: { fontSize: 12, bold: true, color: TXT, align: "center", fill: { color: "F2F4F7" } } });
  const dash = { text: "–", options: { fontSize: 12, color: MUTED, align: "center" } };
  const lift = (pct, x, fs = 15) => {
    const v = parseFloat(x);
    if (v < 1.5) return { text: [{ text: pct, options: { fontSize: fs, bold: true, breakLine: true } }, { text: v < 1 ? x + " as likely" : "about average", options: { fontSize: 9.5 } }],
      options: { color: v < 1 ? RED : TXT, align: "center", fill: { color: "F2F4F7" } } };
    return { text: [{ text: pct, options: { fontSize: fs, bold: true, breakLine: true } }, { text: x + " more likely", options: { fontSize: 9.5 } }],
      options: { color: "FFFFFF", align: "center", fill: { color: v >= 3.5 ? NAVY : BLUE } } };
  };
  const subH = (t) => ({ text: t, options: { fontSize: 10, color: NAVY, indentLevel: 1 } });
  const all = { text: "all", options: { fontSize: 10, color: MUTED, align: "center" } };
  const rows = [
    [H("If the customer…"), H("…% who put money in"), H("…% with a loan"), H("…% who paid in Oct")],
    [rowH("All customers (average)"), base("35%"), base("11%"), base("7.8%")],
    [rowH("Has put money in (525K)"), dash, lift("27%", "2.5x"), lift("22%", "2.8x")],
    [subH("↳ Balance 1K+ today (71K)"), all, lift("12%", "1.1x", 13), lift("35%", "4.5x", 13)],
    [subH("↳ Balance 20K+ today (29K)"), all, lift("4%", "0.4x", 13), lift("22%", "2.8x", 13)],
    [rowH("Has a loan (164K)"), lift("87%", "2.5x"), dash, lift("28.5%", "3.7x")],
    [rowH("Paid in Oct (117K)"), lift("99%", "2.8x"), lift("40%", "3.6x"), dash],
  ];
  s.addTable(rows, { x: 0.5, y: 2.8, w: 4.5, colW: [1.35, 1.05, 1.05, 1.05], fontFace: "Arial", border: { type: "solid", pt: 1.5, color: "FFFFFF" },
    valign: "middle", rowH: [0.46, 0.42, 0.6, 0.55, 0.55, 0.6, 0.6], margin: [0.04, 0.06, 0.04, 0.06], objectName: "Connection matrix" });

  // Engagement ladder (shapes)
  const lx = 5.25;
  s.addText("How far customers climb, thousand", { x: lx, y: 2.4, w: 3.65, h: 0.32, fontSize: 13, bold: true, color: C.text2, margin: 0, isTextBox: true, objectName: "Ladder title" });
  const steps = [["Registered", 2100, ""], ["Account opened", 1500, "71%"], ["Money in", 525, "35%"], ["Paying (Oct)", 117, "22%"], ["2+ products", 51, "44%"]];
  const labW = 1.15, maxW = 1.5, rh = 0.62;
  steps.forEach((st, k) => {
    const y = 2.85 + k * rh, bw = Math.max(0.06, st[1] / 2100 * maxW);
    s.addText(st[0], { x: lx, y, w: labW, h: 0.45, fontSize: 11, color: C.text1, align: "right", valign: "middle", margin: 0, isTextBox: true, objectName: `Ladder label ${k}` });
    s.addShape(pres.shapes.RECTANGLE, { x: lx + labW + 0.1, y: y + 0.05, w: bw, h: 0.35, fill: { color: k >= 3 ? C.accent1 : C.text2 }, line: { type: "none" }, objectName: `Ladder bar ${k}` });
    s.addText([
      { text: st[1].toLocaleString("en-US"), options: { fontSize: 12, bold: true, color: C.text1 } },
      ...(st[2] ? [{ text: "  " + st[2], options: { fontSize: 10, color: MUTED } }] : []),
    ], { x: lx + labW + 0.15 + bw, y, w: 3.65 - labW - 0.15 - bw, h: 0.45, valign: "middle", margin: 0, isTextBox: true, objectName: `Ladder value ${k}` });
  });
  const yy = 2.85 + steps.length * rh;
  s.addText("Main bank", { x: lx, y: yy, w: labW, h: 0.5, fontSize: 11, color: C.text1, align: "right", valign: "middle", margin: 0, isTextBox: true, objectName: "Ladder label main" });
  s.addShape(pres.shapes.RECTANGLE, { x: lx + labW + 0.1, y: yy + 0.05, w: 2.4, h: 0.4, fill: { color: C.background2 }, line: { color: GREY, width: 1, dashType: "dash" }, objectName: "Ladder main placeholder" });
  s.addText("Not measured yet", { x: lx + labW + 0.1, y: yy + 0.05, w: 2.4, h: 0.4, fontSize: 10, bold: true, color: C.accent5, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: "Ladder main text" });

  insights(s, [
    ["Each product pulls in the others", "Holding one makes the other two 2.5–3.7x more likely. Borrowers and payers are the most connected"],
    ["Savers pay, but rarely borrow", "35% of customers with 1K+ pay with CLICX (4.5x), but only 4% with 20K+ have a loan. They need a different offer"],
    ["The chain breaks at step one", "Only 35% of accounts ever get money in, yet 99% of payers have. Money in comes first"],
  ], undefined, undefined, undefined, undefined, "What it means");
}
{
  const s = qSlide("Board summary", "BOARD SUMMARY  (3/4)  |  BALANCE AND USE",
    "Does how much money customers keep with us change what else they do?",
    "Yes. Payments peak among customers keeping 1K–20K (44% pay), while big savers (20K+) pay less and rarely borrow. The 1.14M with a zero balance barely use anything",
    "Balance = current total balance, 1 Oct 2026. Zero and under-1K figures are estimates from rounded shares in the source; 1K+ figures are exact. A 100-baht cut is not in the data yet. Source: CLICX product and payment segments");
  const bands = ["Zero", "Under 1K", "1K–20K", "20K+"];
  s.addChart(pres.charts.BAR, [
    { name: "Has a CLICX loan", labels: bands, values: [0.066, 0.28, 0.18, 0.04] },
    { name: "Paid with CLICX in Oct", labels: bands, values: [0.018, 0.25, 0.44, 0.22] },
  ], { x: 0.5, y: 2.35, w: 5.3, h: 4.35, barDir: "col", barGrouping: "clustered", ...AX, ...quiet(), ...chartTitle("What customers do, by how much they keep with us, %"),
    chartColors: [NAVY, CYAN], showValue: true, dataLabelPosition: "outEnd", dataLabelColor: TXT, dataLabelFontSize: 11, dataLabelFormatCode: "0%",
    showLegend: true, legendPos: "b", legendFontSize: 10, legendColor: MUTED, barGapWidthPct: 50, valAxisMaxVal: 0.52, catAxisLabelFontSize: 11, objectName: "Use by balance band" });
  s.addChart(pres.charts.BAR, [{ name: "Customers, K", labels: bands, values: [1142, 288, 42, 29] }], {
    x: 5.9, y: 2.35, w: 3.0, h: 4.35, barDir: "bar", ...AX, ...quiet(), ...chartTitle("Where customers sit, thousand"),
    chartColors: [RED, GREY, BLUE, NAVY], showValue: true, dataLabelPosition: "outEnd", dataLabelColor: TXT, dataLabelFontSize: 11,
    dataLabelFormatCode: "#,##0", showLegend: false, barGapWidthPct: 40, valAxisMaxVal: 1450, catAxisOrientation: "maxMin", catAxisLabelFontSize: 11, objectName: "Customers by balance band" });
  insights(s, [
    ["Small balances follow loans", "28% of customers holding under 1K have a loan, the highest of any group"],
    ["1K–20K is the everyday sweet spot", "44% pay with CLICX, 6x the average. Grow this group from the 288K holding under 1K"],
    ["Big savers want a savings home", "20K+ customers rarely borrow (4%). Offer savings and rewards, not loans"],
  ], undefined, undefined, undefined, undefined, "What it means");
}
{
  const s = qSlide("Board summary", "BOARD SUMMARY  (4/4)  |  WHAT WE WILL DO",
    "What can we do in the next 90 days, and what do we need from the board?",
    "Six actions aimed at customers we already have: get money in, bring loan payments into the app, and keep payers from dropping off",
    SRC);
  const hdr = (t) => ({ text: t, options: { bold: true, color: "FFFFFF", fill: { color: NAVY }, fontSize: 11 } });
  const a = (t) => ({ text: t, options: { bold: true, fontSize: 11, color: NAVY } });
  const cell = (t) => ({ text: t, options: { fontSize: 10.5, color: TXT } });
  const n = (t) => ({ text: t, options: { bold: true, fontSize: 11, color: BLUE, align: "right" } });
  const rows = [
    [hdr("Action"), hdr("Who"), { text: "How many", options: { bold: true, color: "FFFFFF", fill: { color: NAVY }, fontSize: 11, align: "right" } }, hdr("Why it should work")],
    [a("1  Give turned-down applicants a reason to put money in (bonus save pocket, small starter loan)"), cell("Applied, never got a loan"), n("1.13M"), cell("Money in is the strongest sign someone will pay")],
    [a("2  Bring loan repayment into the app (autopay, due-date reminders)"), cell("Borrowers who never paid in the app"), n("75K"), cell("Payers who come back follow due dates")],
    [a("3  Win back last month's payers every month"), cell("Paid last month, not this month"), n("53K"), cell("Cheaper to keep a payer than to find one")],
    [a("4  Offer AIS bill pay and top-up to scan-to-pay users"), cell("Scan-to-pay only users"), n("65K"), cell("Two ways to pay halves drop-off")],
    [a("5  Nudge small savers to the next savings tier"), cell("Holding 1K–20K"), n("42K"), cell("Our 20K+ savers fell from 34.6K to 28.8K")],
    [a("6  Reward the 2nd and 3rd payment, not just the 1st"), cell("New payers each month"), n("25K"), cell("Bill pay fell 45% when the promotion ended")],
  ];
  s.addTable(rows, { x: 0.5, y: 2.4, w: 8.4, colW: [3.5, 1.9, 0.95, 2.05], fontFace: "Arial", border: { type: "solid", pt: 0.75, color: LGREY },
    valign: "middle", rowH: [0.34, 0.56, 0.5, 0.5, 0.5, 0.5, 0.5], margin: [0.04, 0.08, 0.04, 0.08], objectName: "Action table" });
  s.addText([
    { text: "TRACK MONTHLY   ", options: { bold: true, fontSize: 10, color: C.accent1, charSpacing: 1 } },
    { text: "Money in 35%  ·  Paying 117K  ·  Pay every month 57.7K  ·  2+ products 51K  ·  Savers 1K+ 71K", options: { fontSize: 11, color: C.text1 } },
  ], { x: 0.5, y: 6.25, w: 8.4, h: 0.4, valign: "middle", margin: 0, isTextBox: true, objectName: "Scorecard strip" });
  insights(s, [
    ["Agree the measure", "Track 'customers with money in who pay' next to loan growth"],
    ["Move budget to activation", "Shift part of sign-up spend into actions 1–4"],
    ["Open the balance data", "Put CLICX balances in the shared data room to track savings and the KTB deposit question"],
  ], undefined, undefined, undefined, undefined, "What we ask the board");
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
