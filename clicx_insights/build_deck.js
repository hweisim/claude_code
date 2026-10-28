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
pres.title = "CLICX customer insights";
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
// 1. Title
pres.addSection({ title: "Opening" });
{
  const s = pres.addSlide({ masterName: "TITLE", sectionTitle: "Opening" });
  s.addText("CUSTOMER INSIGHTS", { placeholder: "kicker" });
  s.addText("Who is CLICX really winning, and at whose expense?", { placeholder: "title" });
  s.addText("Five simple questions about our first 1.8 million customers: where they come from, who they are, what they use, who keeps using us, and whether we take customers away from KTB", { placeholder: "sub" });
  s.addText("October 2026  |  CLICX data as of Sep 2026; AIS, KTB, OR data as of Apr 2026", { placeholder: "date" });
}

// 2. Executive summary
{
  const s = pres.addSlide({ masterName: "QUESTION", sectionTitle: "Opening" });
  s.addText("EXECUTIVE SUMMARY", { placeholder: "tracker" });
  s.addText("What have our first 1.8 million customers told us?", { placeholder: "title" });
  s.addText("Our partners bring in almost everyone. Customers are mostly 25–39 and come to borrow. We mostly win people KTB had already lost, not KTB's loyal customers", { placeholder: "answer" });
  s.addText("Source: CLICX, AIS, KTB and OR customer data", { placeholder: "foot" });
  const hdr = (t) => ({ text: t, options: { bold: true, color: "FFFFFF", fill: { color: NAVY }, fontSize: 12 } });
  const verdict = (t, col) => ({ text: t, options: { bold: true, color: col, fontSize: 13 } });
  const cell = (t) => ({ text: t, options: { fontSize: 12, color: TXT } });
  const rows = [
    [hdr("#"), hdr("Question"), hdr("Short answer"), hdr("Why we say so")],
    [cell("1"), cell("Do all CLICX customers come from our partners' customers?"), verdict("Almost all", BLUE), cell("96% were already KTB, AIS or OR customers. Only 51K came from outside. But just 3.7% of partner customers have joined so far")],
    [cell("2"), cell("Are our customers the people we planned for?"), verdict("Mostly", BLUE), cell("Tech-savvy people and freelancers sign up far more than others, as planned. People who use several banks and over-40s are hard to win")],
    [cell("3"), cell("Do different customers come for different products?"), verdict("Same reason, different use", BLUE), cell("Almost everyone (84–92%) applies for a loan. Office workers get loans, shop owners keep savings, gig workers are mostly turned down")],
    [cell("4"), cell("Are active customers different from inactive ones?"), verdict("Yes", BLUE), cell("Active customers are mostly office workers and shop owners with monthly phone bills. Inactive ones are often blue-collar workers or people turned down for a loan")],
    [cell("5"), cell("Do we take customers or products away from KTB?"), verdict("Very little", RED), cell("Only 13% of shared customers are loyal to KTB. Most used other banks or had stopped using KTB. 116K of our borrowers have no KTB loan")],
  ];
  s.addTable(rows, { x: 0.5, y: 2.4, w: 12.33, colW: [0.45, 3.6, 2.1, 6.18], fontFace: "Arial", border: { type: "solid", pt: 0.75, color: LGREY },
    valign: "middle", rowH: [0.4, 0.78, 0.78, 0.78, 0.78, 0.78], margin: [0.05, 0.1, 0.05, 0.1], objectName: "Summary table" });
}

// =====================================================================
// Q1
pres.addSection({ title: "Q1 Source of customers" });
{
  const s = qSlide("Q1 Source of customers", "Q1  |  WHERE CUSTOMERS COME FROM",
    "Do all current CLICX customers come from our partners' customers?",
    "Almost all. 96% were already customers of KTB, AIS or OR, and only 51K came from outside. Yet only 3.7% of partner customers have joined, so there is plenty of room to grow",
    "Source: CLICX customer data (Sep 2026); AIS, KTB, OR customer data (Apr 2026)");
  stat(s, 0.5, 2.4, 2.6, "96%", "of our 1.8M customers were already partner customers");
  stat(s, 0.5, 3.85, 2.6, "51K", "customers came from outside our partners", C.text2);
  stat(s, 0.5, 5.3, 2.6, "3.7%", "of 49M partner customers have joined CLICX", C.text2);
  s.addChart(pres.charts.BAR, [{ name: "Share of CLICX customers", labels: ["Also a KTB customer", "Also an AIS customer", "Also an OR customer"], values: [0.88, 0.67, 0.40] }], {
    x: 3.3, y: 2.35, w: 5.6, h: 4.35, barDir: "bar", ...AX, ...quiet(), ...chartTitle("% of CLICX customers who are also with each partner (many are with more than one)"),
    chartColors: [NAVY, NAVY, NAVY], showValue: true, dataLabelPosition: "outEnd", dataLabelColor: TXT, dataLabelFontSize: 12,
    dataLabelFormatCode: "0%", showLegend: false, barGapWidthPct: 60, valAxisMinVal: 0, valAxisMaxVal: 1.05, catAxisOrientation: "maxMin", catAxisLabelFontSize: 12, objectName: "Partner reach" });
  insights(s, [
    ["KTB is the front door", "9 in 10 of our customers bank with KTB, mostly joining through the Paotang app"],
    ["Most know us through 2+ partners", "That gives us more data to judge who to lend to and what else to offer"],
    ["AIS is the biggest opportunity", "Only 3.7% of AIS's ~32M customers have joined, against 7.4% for OR and 4.7% for KTB"],
  ]);
}

// =====================================================================
// Q2
pres.addSection({ title: "Q2 Who our customers are" });
{
  const s = qSlide("Q2 Who our customers are", "Q2  |  WHO OUR CUSTOMERS ARE  (1/3)",
    "Are the people who come to CLICX the ones we planned for?",
    "Mostly yes. Tech-savvy people, our main target, join at twice the average rate, and one-bank users bring the most customers. People who use several banks are not yet won",
    "Source: CLICX customer data (Sep 2026); AIS, KTB, OR customer data (Apr 2026)");
  const labels = ["Tech-savvy (8.5% joined)", "One-bank users (4.1%)", "Multi-bank users (2.4%)", "Unknown (2.0%)"];
  s.addChart(pres.charts.BAR, [
    { name: "Share of all partner customers (49.5M)", labels, values: [0.127, 0.415, 0.165, 0.292] },
    { name: "Share of CLICX customers (1.87M)", labels, values: [0.285, 0.455, 0.104, 0.156] },
  ], { x: 0.5, y: 2.35, w: 8.4, h: 4.35, barDir: "col", barGrouping: "clustered", ...AX, ...quiet(), ...chartTitle("Customer type: all partner customers vs CLICX customers, %"),
    chartColors: [LGREY, BLUE], showValue: true, dataLabelPosition: "outEnd", dataLabelColor: TXT, dataLabelFontSize: 11, dataLabelFormatCode: "0%",
    showLegend: true, legendPos: "b", legendFontSize: 10, legendColor: MUTED, barGapWidthPct: 60, valAxisMaxVal: 0.55, objectName: "Persona mix" });
  insights(s, [
    ["Tech-savvy: target hit", "28% of our customers vs 13% of partner customers. Best of all at ages 30–34, where 13% have joined"],
    ["One-bank users: the biggest group", "850K customers, almost half of CLICX. 577K are KTB customers who do not use AIS"],
    ["Multi-bank users: not won yet", "8.2M people, but only 2.4% have joined. People who already shop around need a stronger reason"],
  ]);
}
{
  const s = qSlide("Q2 Who our customers are", "Q2  |  WHO OUR CUSTOMERS ARE  (2/3)",
    "Do our customers' ages match the people we can reach?",
    "No. Most partner customers are over 40, but most of ours are 25–39 (56% of CLICX vs 29% of partner customers). Few people over 45 join",
    "Source: CLICX customer data (Sep 2026); AIS, KTB, OR customer data (Apr 2026)");
  const ages = ["<20", "20–24", "25–29", "30–34", "35–39", "40–44", "45–49", "50–54", "55–59", "60+"];
  const clicx = [12.4, 163, 334, 385, 328, 268, 189, 104, 51, 31];
  const base = [2280, 3569, 4162, 4496, 4211, 4271, 4366, 4165, 3954, 9355];
  const sum = (a) => a.reduce((x, y) => x + y, 0);
  const share = (a) => a.map((v) => +(v / sum(a)).toFixed(3));
  s.addChart(pres.charts.BAR, [
    { name: "Share of partner customers (age known)", labels: ages, values: share(base) },
    { name: "Share of CLICX customers", labels: ages, values: share(clicx) },
  ], { x: 0.5, y: 2.35, w: 8.4, h: 4.35, barDir: "col", barGrouping: "clustered", ...AX, ...quiet(), ...chartTitle("Age: all partner customers vs CLICX customers, %"),
    chartColors: [LGREY, BLUE], showValue: true, dataLabelPosition: "outEnd", dataLabelColor: TXT, dataLabelFontSize: 9, dataLabelFormatCode: "0%",
    showLegend: true, legendPos: "b", legendFontSize: 10, legendColor: MUTED, barGapWidthPct: 40, valAxisMaxVal: 0.25, catAxisLabelFontSize: 11, objectName: "Age mix" });
  insights(s, [
    ["25–39 is our sweet spot", "1.05M customers. About 8–9% of people this age have joined, more than twice the average"],
    ["Older customers are untouched", "Over-40s are 58% of partner customers but 34% of ours. Only 0.3% of 9.4M over-60s have joined"],
    ["Plenty left in our core group", "2.1M tech-savvy 25–39s and 5.4M one-bank 25–39s have not joined yet"],
  ]);
}
{
  const s = qSlide("Q2 Who our customers are", "Q2  |  WHO OUR CUSTOMERS ARE  (3/3)",
    "Are we reaching people with real income but no payslip, as our strategy says?",
    "Yes. Freelancers (40%) and riders/drivers (22%) join far more than anyone else, and 35% of our loans go to people banks usually turn away, above our 30% goal",
    "Source: CLICX, AIS, KTB, OR matched occupation data (Oct 2026); CLICX 5-Year Business Plan");
  const occ = ["Freelancer (232K)", "Rider / driver (66K)", "Blue-collar worker (401K)", "Office worker (700K)", "Shop owner (293K)", "Unknown (129K)"];
  s.addChart(pres.charts.BAR, [{ name: "Joined CLICX", labels: occ, values: [0.40, 0.22, 0.16, 0.09, 0.09, 0.004] }], {
    x: 0.5, y: 2.35, w: 8.4, h: 4.35, barDir: "bar", ...AX, ...quiet(), ...chartTitle("% of partner customers in each job who joined CLICX (number of CLICX customers)"),
    chartColors: [BLUE, BLUE, NAVY, NAVY, NAVY, GREY], showValue: true, dataLabelPosition: "outEnd", dataLabelColor: TXT, dataLabelFontSize: 12,
    dataLabelFormatCode: "0%", showLegend: false, barGapWidthPct: 45, valAxisMaxVal: 0.46, catAxisOrientation: "maxMin", catAxisLabelFontSize: 12, objectName: "Occupation penetration" });
  insights(s, [
    ["Gig workers love CLICX", "4 in 10 freelancers and 2 in 10 riders we can reach have joined. Proof that lending on phone and spending data works"],
    ["Workers are the bulk", "Office and blue-collar workers are 1.1M customers, over 60% of CLICX"],
    ["Shop owners: room to grow", "293K have joined out of 3.3M. A route into small-business loans"],
  ]);
}

// =====================================================================
// Q3
pres.addSection({ title: "Q3 What they come for" });
{
  const s = qSlide("Q3 What they come for", "Q3  |  WHAT THEY COME FOR",
    "Do different types of customer come for the same product or different ones?",
    "The same: almost everyone comes to borrow. But what they end up using differs. Office workers get loans, shop owners keep savings, and gig workers are mostly turned down",
    "Source: CLICX, AIS, KTB, OR matched occupation data (Oct 2026)");
  const occ = ["Office worker", "Shop owner", "Blue-collar worker", "Freelancer", "Rider / driver"];
  s.addChart(pres.charts.BAR, [
    { name: "Applied for a loan", labels: occ, values: [0.85, 0.84, 0.92, 0.88, 0.91] },
    { name: "Got a loan", labels: occ, values: [0.11, 0.07, 0.07, 0.06, 0.05] },
  ], { x: 0.5, y: 2.35, w: 4.2, h: 4.35, barDir: "bar", barGrouping: "clustered", ...AX, ...quiet(), ...chartTitle("Loans: applied vs got one, % of each job type"),
    chartColors: [LGREY, NAVY], showValue: true, dataLabelPosition: "outEnd", dataLabelColor: TXT, dataLabelFontSize: 10, dataLabelFormatCode: "0%",
    showLegend: true, legendPos: "b", legendFontSize: 10, legendColor: MUTED, barGapWidthPct: 40, valAxisMaxVal: 1.15, catAxisOrientation: "maxMin", objectName: "Loans applied vs got" });
  s.addChart(pres.charts.BAR, [
    { name: "Keeps money in CLICX", labels: occ, values: [0.18, 0.20, 0.12, 0.16, 0.13] },
    { name: "Paid with CLICX in Sep", labels: occ, values: [0.14, 0.14, 0.10, 0.12, 0.11] },
  ], { x: 4.8, y: 2.35, w: 4.1, h: 4.35, barDir: "bar", barGrouping: "clustered", ...AX, ...quiet(), ...chartTitle("Saving and paying, % of each job type"),
    chartColors: [BLUE, CYAN], showValue: true, dataLabelPosition: "outEnd", dataLabelColor: TXT, dataLabelFontSize: 10, dataLabelFormatCode: "0%",
    showLegend: true, legendPos: "b", legendFontSize: 10, legendColor: MUTED, barGapWidthPct: 40, valAxisMaxVal: 0.26, catAxisOrientation: "maxMin", objectName: "Saving and paying" });
  insights(s, [
    ["Loans bring everyone in", "1.52M customers (84%) have applied for a loan, whatever their job"],
    ["Office workers use the most", "Most likely to get a loan (11%) and to pay with CLICX (14%)"],
    ["Shop owners save", "Most likely to keep money with us (20%), a natural fit for business accounts and loans"],
    ["Gig workers get turned down", "Over 90% apply but only 5–7% get a loan. The problem is our approval rules, not demand"],
  ]);
}

// =====================================================================
// Q4
pres.addSection({ title: "Q4 Active vs inactive" });
{
  const s = qSlide("Q4 Active vs inactive", "Q4  |  WHO KEEPS USING US  (1/3)",
    "Are active CLICX customers a different type of person from inactive ones?",
    "Yes. Only 1 in 5 customers is active (has money in CLICX or a CLICX loan). Active customers are mostly office workers and shop owners; inactive ones are more often blue-collar workers",
    "Source: CLICX, AIS, KTB, OR matched occupation data (Oct 2026)");
  const act = [164, 68, 64, 42, 10, 10], tot = [700, 293, 401, 232, 66, 129];
  const occ = ["Office worker (23% active)", "Shop owner (23%)", "Blue-collar worker (16%)", "Freelancer (18%)", "Rider / driver (15%)", "Unknown job (8%)"];
  const sa = 358, si = 1462;
  s.addChart(pres.charts.BAR, [
    { name: "Active customers (358K)", labels: occ, values: act.map((v) => +(v / sa).toFixed(3)) },
    { name: "Inactive customers (1.46M)", labels: occ, values: act.map((v, i) => +((tot[i] - v) / si).toFixed(3)) },
  ], { x: 0.5, y: 2.35, w: 8.4, h: 4.35, barDir: "bar", barGrouping: "clustered", ...AX, ...quiet(), ...chartTitle("Job type of active vs inactive customers, %"),
    chartColors: [NAVY, LGREY], showValue: true, dataLabelPosition: "outEnd", dataLabelColor: TXT, dataLabelFontSize: 11, dataLabelFormatCode: "0%",
    showLegend: true, legendPos: "b", legendFontSize: 10, legendColor: MUTED, barGapWidthPct: 40, valAxisMaxVal: 0.55, catAxisOrientation: "maxMin", objectName: "Active vs inactive mix" });
  insights(s, [
    ["Office workers and shop owners stay", "65% of active customers vs 52% of inactive ones. About 1 in 4 of them is active"],
    ["Blue-collar workers drift away", "They apply for a loan, get turned down and stop using CLICX"],
    ["Unknown job, unknown customer", "Only 8% are active. Most inactive customers never got a loan or put money in"],
  ]);
}
{
  const s = qSlide("Q4 Active vs inactive", "Q4  |  WHO KEEPS USING US  (2/3)",
    "What else separates active customers from inactive ones?",
    "Two things: how they pay for their phone and whether they got a loan. Monthly-bill AIS customers are 3x more active than lapsed ones, and borrowers pay with CLICX 6x more often",
    "Source: AIS and CLICX matched customer data (Sep 2026)");
  const seg = ["Home internet only", "Monthly phone bill", "Prepaid, pay-per-use", "Prepaid, little spend", "Prepaid, regular top-up", "Prepaid, occasional top-up", "Stopped using AIS"];
  s.addChart(pres.charts.BAR, [{ name: "Active in CLICX", labels: seg, values: [0.41, 0.28, 0.22, 0.21, 0.19, 0.15, 0.09] }], {
    x: 0.5, y: 2.35, w: 4.6, h: 4.35, barDir: "bar", ...AX, ...quiet(), ...chartTitle("% active in CLICX, by type of AIS customer"),
    chartColors: [NAVY, NAVY, GREY, GREY, GREY, GREY, RED], showValue: true, dataLabelPosition: "outEnd", dataLabelColor: TXT, dataLabelFontSize: 11,
    dataLabelFormatCode: "0%", showLegend: false, barGapWidthPct: 40, valAxisMaxVal: 0.5, catAxisOrientation: "maxMin", catAxisLabelFontSize: 10, objectName: "AIS segment active rate" });
  s.addChart(pres.charts.BAR, [{ name: "Pays with CLICX", labels: ["Got a loan", "Never turned down", "Turned down for a loan"], values: [0.51, 0.16, 0.08] }], {
    x: 5.2, y: 2.35, w: 3.7, h: 4.35, barDir: "col", ...AX, ...quiet(), ...chartTitle("% who pay with CLICX"),
    chartColors: [NAVY, GREY, RED], showValue: true, dataLabelPosition: "outEnd", dataLabelColor: TXT, dataLabelFontSize: 12,
    dataLabelFormatCode: "0%", showLegend: false, barGapWidthPct: 50, valAxisMaxVal: 0.6, catAxisLabelFontSize: 10, objectName: "Paying by product status" });
  insights(s, [
    ["Steady phone payers stay", "Home internet (41%) and monthly-bill (28%) customers are the most active"],
    ["A loan builds the habit", "Half of borrowers pay with CLICX, against 8% of people we turned down"],
    ["Saying 'no' loses customers", "2 in 3 customers have been turned down for a loan at some point, and most of them never use CLICX again"],
  ]);
}

{
  const s = qSlide("Q4 Active vs inactive", "Q4  |  WHO KEEPS USING US  (3/3)",
    "Who pays bills, tops up or scans to pay with CLICX, and are they different from those who don't?",
    "Having a CLICX loan matters most: 54% of borrowers paid with CLICX in Aug–Oct vs 11% of everyone else. Job and region matter less, but government staff, merchants and Chiang Mai lead",
    "% paid = share who paid a bill, topped up or scanned to pay with CLICX at least once in Aug–Oct 2026 (average 15.4%). Source: CLICX payment segments (1.50M customers)");
  const fmt = { showValue: true, dataLabelPosition: "outEnd", dataLabelColor: TXT, dataLabelFontSize: 10, dataLabelFormatCode: "0%", showLegend: false };
  s.addChart(pres.charts.BAR, [{ name: "Paid Aug–Oct", labels: ["Has a CLICX loan", "No CLICX loan"], values: [0.543, 0.106] }], {
    x: 0.5, y: 2.35, w: 2.0, h: 4.35, barDir: "col", ...AX, ...quiet(), ...chartTitle("% paid, by loan"), ...fmt,
    chartColors: [NAVY, GREY], barGapWidthPct: 50, valAxisMaxVal: 0.65, catAxisLabelFontSize: 10, objectName: "Paid by loan status" });
  const occ = ["Government staff", "Online seller", "Shop owner", "Company employee", "Freelancer", "Rider / driver", "Daily wage earner", "Farmer", "Blue-collar worker"];
  s.addChart(pres.charts.BAR, [{ name: "Paid Aug–Oct", labels: occ, values: [0.183, 0.181, 0.177, 0.169, 0.132, 0.130, 0.125, 0.124, 0.118] }], {
    x: 2.6, y: 2.35, w: 3.2, h: 4.35, barDir: "bar", ...AX, ...quiet(), ...chartTitle("% paid, by job"), ...fmt, dataLabelFormatCode: "0.0%",
    chartColors: [NAVY, NAVY, NAVY, NAVY, GREY, GREY, GREY, GREY, RED], barGapWidthPct: 35, valAxisMaxVal: 0.24, catAxisOrientation: "maxMin", catAxisLabelFontSize: 10, objectName: "Paid by occupation" });
  const reg = ["Chiang Mai", "Nonthaburi", "Bangkok", "Pathum Thani", "Chon Buri", "Rayong", "Nakhon Ratchasima", "Other provinces", "Samut Prakan"];
  s.addChart(pres.charts.BAR, [{ name: "Paid Aug–Oct", labels: reg, values: [0.185, 0.175, 0.169, 0.155, 0.151, 0.151, 0.151, 0.147, 0.144] }], {
    x: 5.9, y: 2.35, w: 3.0, h: 4.35, barDir: "bar", ...AX, ...quiet(), ...chartTitle("% paid, by region"), ...fmt, dataLabelFormatCode: "0.0%",
    chartColors: [NAVY, NAVY, NAVY, GREY, GREY, GREY, GREY, GREY, GREY], barGapWidthPct: 35, valAxisMaxVal: 0.24, catAxisOrientation: "maxMin", catAxisLabelFontSize: 10, objectName: "Paid by region" });
  insights(s, [
    ["A loan is the strongest signal", "Borrowers are 5x more likely to pay with CLICX (54% vs 11%)"],
    ["Steady earners pay more", "Government staff, sellers and company employees: 17–18%. Blue-collar workers and farmers: about 12%"],
    ["Chiang Mai and Nonthaburi lead", "18.5% and 17.5%, vs 14–15% in most provinces. Region matters less than job"],
  ]);
}

// =====================================================================
// Q5
pres.addSection({ title: "Q5 Impact on KTB" });
{
  const s = qSlide("Q5 Impact on KTB", "Q5  |  IMPACT ON KTB  (1/2)",
    "Does CLICX take customers away from KTB?",
    "Very little. Of 1.6M customers we share with KTB, only 13% are loyal to KTB. 61% mainly used other banks or had stopped using KTB, so we are winning them back for the group",
    "Source: KTB and CLICX matched customer data (Oct 2026)");
  const b = ["Loyal to KTB (at risk)", "Uses KTB and other banks", "Loyal to other banks (won back)", "Stopped using KTB (brought back)"];
  s.addChart(pres.charts.BAR, [{ name: "Share of shared customers", labels: b, values: [0.13, 0.26, 0.31, 0.30] }], {
    x: 0.5, y: 2.35, w: 5.0, h: 4.35, barDir: "bar", ...AX, ...quiet(), ...chartTitle("Customers we share with KTB, by how they use KTB, %"),
    chartColors: [RED, GREY, BLUE, BLUE], showValue: true, dataLabelPosition: "outEnd", dataLabelColor: TXT, dataLabelFontSize: 12,
    dataLabelFormatCode: "0%", showLegend: false, barGapWidthPct: 45, valAxisMaxVal: 0.4, catAxisOrientation: "maxMin", objectName: "Overlap by KTB relationship" });
  s.addChart(pres.charts.BAR, [{ name: "Joined CLICX", labels: ["Loyal to KTB", "KTB and others", "Other banks", "Stopped KTB"], values: [0.033, 0.074, 0.079, 0.030] }], {
    x: 5.6, y: 2.35, w: 3.3, h: 4.35, barDir: "bar", ...AX, ...quiet(), ...chartTitle("% of each KTB group who joined CLICX"),
    chartColors: [RED, GREY, BLUE, BLUE], showValue: true, dataLabelPosition: "outEnd", dataLabelColor: TXT, dataLabelFontSize: 12,
    dataLabelFormatCode: "0.0%", showLegend: false, barGapWidthPct: 45, valAxisMaxVal: 0.1, catAxisOrientation: "maxMin", catAxisLabelFontSize: 10, objectName: "Penetration by KTB group" });
  insights(s, [
    ["KTB loyalists rarely switch", "Only 3.3% of them joined, less than half the rate of people who bank elsewhere (7.9%)"],
    ["We win back other banks' customers", "515K customers mainly banked elsewhere. Now they are back in the KTB group through CLICX"],
    ["We wake up lapsed customers", "488K had not used KTB for a year or more, then joined CLICX"],
  ]);
}
{
  const s = qSlide("Q5 Impact on KTB", "Q5  |  IMPACT ON KTB  (2/2)",
    "Does CLICX take products away from KTB?",
    "Not for loans: we mostly lend to people KTB's own app loan (Jai-pump) does not serve or turns down. We cannot check savings yet",
    "Source: KTB and CLICX matched customer data (Oct 2026)");
  s.addChart(pres.charts.BAR, [{ name: "CLICX customers, K", labels: ["Our borrowers with no KTB loan", "Turned down by KTB, applied to us", "Turned down by KTB, got our loan"], values: [116, 156, 23] }], {
    x: 0.5, y: 2.35, w: 5.4, h: 4.35, barDir: "bar", catAxisOrientation: "maxMin", ...AX, ...quiet(), ...chartTitle("CLICX lending that KTB's Jai-pump loan does not cover, thousand customers"),
    chartColors: [BLUE, NAVY, NAVY], showValue: true, dataLabelPosition: "outEnd", dataLabelColor: TXT, dataLabelFontSize: 12,
    dataLabelFormatCode: "#,##0", showLegend: false, barGapWidthPct: 50, valAxisMaxVal: 190, catAxisLabelFontSize: 11, objectName: "Lending outside KTB" });
  s.addShape(pres.shapes.RECTANGLE, { x: 6.1, y: 2.45, w: 2.8, h: 4.15, fill: { color: C.background2 }, line: { color: GREY, width: 1, dashType: "dash" }, objectName: "Deposit gap box" });
  s.addText([
    { text: "SAVINGS: NOT KNOWN YET", options: { bold: true, fontSize: 11, color: C.accent5, charSpacing: 1, breakLine: true } },
    { text: " ", options: { fontSize: 8, breakLine: true } },
    { text: "We cannot yet see CLICX balances, so we cannot tell if money is moving out of KTB accounts.", options: { fontSize: 12, color: C.text1, breakLine: true } },
    { text: " ", options: { fontSize: 8, breakLine: true } },
    { text: "Next step: once the data is in, compare shared customers' KTB balances before and after they joined CLICX", options: { fontSize: 12, bold: true, color: C.text2 } },
  ], { x: 6.3, y: 2.65, w: 2.4, h: 3.8, valign: "top", margin: 0, isTextBox: true, objectName: "Deposit gap text" });
  insights(s, [
    ["New loans for the group", "116K of our borrowers have no KTB Jai-pump loan"],
    ["A safety net for KTB", "156K people KTB turned down in the last 3 months applied to us instead of going to another bank"],
    ["Few loyal KTB users among our best customers", "Only 15% of our most engaged shared customers are loyal to KTB"],
  ]);
}

// =====================================================================
// Implications
pres.addSection({ title: "What to do next" });
{
  const s = qSlide("What to do next", "WHAT TO DO NEXT",
    "What should we do with these answers?",
    "Keep using partners to bring people in, and focus on turning sign-ups into regular users, starting with the customer types that already work",
    "Source: CLICX team analysis");
  const cards = [
    ["Q1", "Reach more partner customers", "Start with AIS: only 3.7% of its ~32M customers have joined. Sign them up through the myAIS app and AIS shops"],
    ["Q2", "Find more people like our best customers", "Target tech-savvy 30–34s. Create a clear reason for multi-bank users to switch. Make it easy for gig workers to prove income"],
    ["Q3", "Say 'yes' more safely", "Over 90% of blue-collar and gig workers apply but only 5–7% get a loan. Improve approval rules and offer small starter loans"],
    ["Q4", "Keep new customers using us", "Get them using 3+ features in the first month. Give people we turn down another reason to stay"],
    ["Q5", "Show KTB we grow the group", "Share the 515K won back and 488K brought back. Set up referrals for people KTB turns down. Check savings once data is in"],
  ];
  const cw = 2.33, gap = 0.17;
  cards.forEach((c, i) => {
    const x = 0.5 + i * (cw + gap), y = 2.4;
    s.addShape(pres.shapes.RECTANGLE, { x, y, w: cw, h: 4.25, fill: { color: C.background2 }, line: { type: "none" }, objectName: `Card ${c[0]}` });
    s.addShape(pres.shapes.OVAL, { x: x + 0.2, y: y + 0.22, w: 0.55, h: 0.55, fill: { color: C.text2 }, line: { type: "none" }, objectName: `Card badge ${c[0]}` });
    s.addText(c[0], { x: x + 0.2, y: y + 0.22, w: 0.55, h: 0.55, fontSize: 13, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: `Card badge text ${c[0]}` });
    s.addText(c[1], { x: x + 0.2, y: y + 0.95, w: cw - 0.4, h: 0.85, fontSize: 15, bold: true, color: C.text2, valign: "top", margin: 0, isTextBox: true, objectName: `Card title ${c[0]}` });
    s.addText(c[2], { x: x + 0.2, y: y + 1.85, w: cw - 0.4, h: 2.25, fontSize: 12, color: C.text1, valign: "top", margin: 0, isTextBox: true, objectName: `Card body ${c[0]}` });
  });
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
