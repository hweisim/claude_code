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
    { placeholder: { options: { name: "foot", type: "body", x: 0.5, y: 6.82, w: 11.2, h: 0.45, fontSize: 10, color: "6E7C87", valign: "top", margin: 0 }, text: "" } },
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
  s.addText("Five questions on where 1.8M CLICX customers come from, who they are, what they use, who stays active, and whether CLICX cannibalizes KTB", { placeholder: "sub" });
  s.addText("October 2026  |  CLICX data as of Sep 2026; AIS, KTB, OR data as of Apr 2026", { placeholder: "date" });
}

// 2. Executive summary
{
  const s = pres.addSlide({ masterName: "QUESTION", sectionTitle: "Opening" });
  s.addText("EXECUTIVE SUMMARY", { placeholder: "tracker" });
  s.addText("What have CLICX's first 1.8M customers told us?", { placeholder: "title" });
  s.addText("CLICX is an ecosystem bank: it wins consortium customers aged 25–39 who come for credit, and mostly grows the pie for KTB rather than taking from it", { placeholder: "answer" });
  s.addText("Source: CLICX data room (Sep 2026); AIS, KTB, OR captive data (Apr 2026); Data for Business Strategy (14 Oct 2026)", { placeholder: "foot" });
  const hdr = (t) => ({ text: t, options: { bold: true, color: "FFFFFF", fill: { color: NAVY }, fontSize: 12 } });
  const verdict = (t, col) => ({ text: t, options: { bold: true, color: col, fontSize: 13 } });
  const cell = (t) => ({ text: t, options: { fontSize: 12, color: TXT } });
  const rows = [
    [hdr("#"), hdr("Question"), hdr("Answer"), hdr("Evidence")],
    [cell("1"), cell("Do all CLICX customers come from captive consortium customers?"), verdict("Almost all", BLUE), cell("96% come from the AIS x KTB x OR base; only 51K are outside it. Yet CLICX reaches just 3.7% of the 49M base")],
    [cell("2"), cell("Are the personas who come the ones we planned for?"), verdict("Mostly, with gaps", BLUE), cell("Tech-Savvy over-index 2.2x and gig workers convert best (Freelancer 40%); Multiple-Bank and 40+ customers under-index")],
    [cell("3"), cell("Do different personas come for the same or different products?"), verdict("Same entry, different use", BLUE), cell("84–92% of every occupation apply for a loan; White-collar books loans (11%), Merchants hold deposits (20%), gig workers rarely get booked (5–6%)")],
    [cell("4"), cell("Is the persona different for active vs inactive customers?"), verdict("Yes", BLUE), cell("Active base skews to White-collar and Merchant (65% vs 52%), post-paid and FBB users; inactive skews to Blue-collar, prepaid and unknown")],
    [cell("5"), cell("Does CLICX cannibalize KTB customers or products?"), verdict("Limited", RED), cell("Only 13% of overlap (212K) are KTB-loyal; 61% are competitor-loyal or KTB-inactive. 116K CLICX borrowers have no KTB Jai-pump loan")],
  ];
  s.addTable(rows, { x: 0.5, y: 2.4, w: 12.33, colW: [0.45, 3.6, 2.1, 6.18], fontFace: "Arial", border: { type: "solid", pt: 0.75, color: LGREY },
    valign: "middle", rowH: [0.4, 0.78, 0.78, 0.78, 0.78, 0.78], margin: [0.05, 0.1, 0.05, 0.1], objectName: "Summary table" });
}

// =====================================================================
// Q1
pres.addSection({ title: "Q1 Source of customers" });
{
  const s = qSlide("Q1 Source of customers", "Q1  |  SOURCE OF CUSTOMERS",
    "Do all current CLICX customers come from captive consortium customers?",
    "Almost all: 96% of CLICX customers already bank, phone or fill up with KTB, AIS or OR; only 51K come from outside. But the consortium is far from tapped, with only 3.7% penetration",
    "Note: one customer can belong to several partners, so partner shares add up to 195%. Penetration = CLICX customers ÷ partner customer base. 85% of CLICX customers gave consent for data matching\nSource: CLICX data (Sep 2026); AIS, KTB, OR data (Apr 2026)");
  stat(s, 0.5, 2.4, 2.6, "96%", "of 1.8M CLICX customers come from the consortium base");
  stat(s, 0.5, 3.85, 2.6, "51K", "CLICX customers who are not AIS, KTB or OR customers", C.text2);
  stat(s, 0.5, 5.3, 2.6, "47M", "consortium customers not yet on CLICX", C.text2);
  s.addChart(pres.charts.BAR, [{ name: "CLICX customers, K", labels: ["KTB", "AIS", "OR", "Outside"], values: [1600, 1200, 725, 51] }], {
    x: 3.2, y: 2.35, w: 3.0, h: 4.35, barDir: "col", ...AX, ...quiet(), ...chartTitle("CLICX customers by partner, K"),
    chartColors: [NAVY, NAVY, NAVY, GREY], showValue: true, dataLabelPosition: "outEnd", dataLabelColor: TXT, dataLabelFontSize: 11,
    dataLabelFormatCode: "#,##0", showLegend: false, barGapWidthPct: 60, catAxisLabelFontSize: 10, objectName: "Customers by partner" });
  s.addChart(pres.charts.BAR, [{ name: "Penetration of partner base", labels: ["KTB", "AIS", "OR"], values: [0.047, 0.037, 0.074] }], {
    x: 6.3, y: 2.35, w: 2.6, h: 4.35, barDir: "col", ...AX, ...quiet(), ...chartTitle("Penetration of partner base"),
    chartColors: [CYAN], showValue: true, dataLabelPosition: "outEnd", dataLabelColor: TXT, dataLabelFontSize: 11,
    dataLabelFormatCode: "0.0%", showLegend: false, barGapWidthPct: 60, valAxisMinVal: 0, valAxisMaxVal: 0.09, objectName: "Penetration by partner" });
  insights(s, [
    ["KTB is the main gateway", "88% of CLICX customers are KTB customers (1.6M), helped by Paotang onboarding"],
    ["Customers sit in several partners", "Shares add up to 195%: most customers are known to 2+ partners, giving rich data for credit"],
    ["AIS is the biggest headroom", "Lowest penetration (3.7%) on the largest base (~32M). OR converts best at 7.4%"],
  ]);
}

// =====================================================================
// Q2
pres.addSection({ title: "Q2 Personas vs plan" });
{
  const s = qSlide("Q2 Personas vs plan", "Q2  |  PERSONAS VS PLAN  (1/3)",
    "Are the personas of customers who come to CLICX the same as we planned?",
    "Mostly yes. The planned Tech-Savvy target over-indexes 2.2x and Single-Bank users bring the volume, but Multiple-Bank users, also a target, are under-won",
    "Note: personas assigned by AIS and KTB. Single-Bank = uses one bank, or KTB customer holding Next but inactive. Index = share of CLICX ÷ share of consortium base\nSource: CLICX data (Sep 2026); AIS, KTB, OR data (Apr 2026); initial target persona benchmark");
  const labels = ["Tech-Savvy", "Single-Bank", "Multiple-Bank", "Unknown"];
  s.addChart(pres.charts.BAR, [
    { name: "Share of consortium base (49.5M)", labels, values: [0.127, 0.415, 0.165, 0.292] },
    { name: "Share of CLICX customers (1.87M)", labels, values: [0.285, 0.455, 0.104, 0.156] },
  ], { x: 0.5, y: 2.35, w: 5.4, h: 4.35, barDir: "col", barGrouping: "clustered", ...AX, ...quiet(), ...chartTitle("Persona mix: consortium base vs CLICX, %"),
    chartColors: [LGREY, BLUE], showValue: true, dataLabelPosition: "outEnd", dataLabelColor: TXT, dataLabelFontSize: 10, dataLabelFormatCode: "0%",
    showLegend: true, legendPos: "b", legendFontSize: 10, legendColor: MUTED, barGapWidthPct: 50, valAxisMaxVal: 0.55, objectName: "Persona mix" });
  s.addChart(pres.charts.BAR, [{ name: "Penetration", labels, values: [0.085, 0.041, 0.024, 0.020] }], {
    x: 6.0, y: 2.35, w: 2.9, h: 4.35, barDir: "bar", ...AX, ...quiet(), ...chartTitle("Penetration by persona"),
    chartColors: [BLUE, NAVY, GREY, GREY], showValue: true, dataLabelPosition: "outEnd", dataLabelColor: TXT, dataLabelFontSize: 11,
    dataLabelFormatCode: "0.0%", showLegend: false, barGapWidthPct: 50, valAxisMaxVal: 0.11, catAxisOrientation: "maxMin", objectName: "Persona penetration" });
  insights(s, [
    ["Tech-Savvy: target hit", "533K customers, 28% of CLICX from 13% of the base; penetration 8.5% (2.2x average), peaking at 13% for ages 30–34"],
    ["Single-Bank: volume engine", "850K (45% of CLICX). 577K are KTB customers without AIS, a pool to cross-sell AIS"],
    ["Multiple-Bank: the miss", "8.2M base but only 2.4% penetration. CLICX is not yet winning customers who already shop around banks"],
  ]);
}
{
  const s = qSlide("Q2 Personas vs plan", "Q2  |  PERSONAS VS PLAN  (2/3)",
    "Does the age profile of CLICX customers match the target pool?",
    "No. The target pool skews older, but CLICX concentrates in ages 25–39: 56% of customers vs 29% of the base. Penetration falls off sharply after 45",
    "Note: penetration = CLICX customers ÷ consortium captive customers in that age band; base share excludes 4.7M customers of unknown age\nSource: CLICX data (Sep 2026); AIS, KTB, OR data (Apr 2026)");
  const ages = ["<20", "20–24", "25–29", "30–34", "35–39", "40–44", "45–49", "50–54", "55–59", "60+"];
  const clicx = [12.4, 163, 334, 385, 328, 268, 189, 104, 51, 31];
  const base = [2280, 3569, 4162, 4496, 4211, 4271, 4366, 4165, 3954, 9355];
  const pen = clicx.map((v, i) => +(v / base[i]).toFixed(4));
  s.addChart(pres.charts.BAR, [{ name: "CLICX customers, K", labels: ages, values: clicx }], {
    x: 0.5, y: 2.35, w: 8.4, h: 2.15, barDir: "col", ...AX, ...quiet(), ...chartTitle("CLICX customers by age, K"),
    chartColors: [NAVY, NAVY, BLUE, BLUE, BLUE, NAVY, NAVY, NAVY, NAVY, NAVY], showValue: true, dataLabelPosition: "outEnd", dataLabelColor: TXT,
    dataLabelFontSize: 10, dataLabelFormatCode: "#,##0", showLegend: false, barGapWidthPct: 40, valAxisMaxVal: 480, catAxisLabelFontSize: 10, objectName: "Customers by age" });
  s.addChart(pres.charts.BAR, [{ name: "Penetration", labels: ages, values: pen }], {
    x: 0.5, y: 4.55, w: 8.4, h: 2.15, barDir: "col", ...AX, ...quiet(), ...chartTitle("Penetration of captive base by age, %"),
    chartColors: [GREY, GREY, CYAN, CYAN, CYAN, GREY, GREY, GREY, GREY, GREY], showValue: true, dataLabelPosition: "outEnd", dataLabelColor: TXT,
    dataLabelFontSize: 10, dataLabelFormatCode: "0.0%", showLegend: false, barGapWidthPct: 40, valAxisMaxVal: 0.11, catAxisLabelFontSize: 10, objectName: "Penetration by age" });
  insights(s, [
    ["Sweet spot is 25–39", "1.05M customers (56% of CLICX); penetration of 8–9% vs 3.8% overall"],
    ["Older base is untouched", "Customers aged 40+ make up 58% of the known-age base but only 34% of CLICX; 60+ penetration is 0.3% on 9.4M people"],
    ["Room left in the core", "2.1M unconverted Tech-Savvy aged 25–39, 1.3M young tech users under 25 and 5.4M Single-Bank millennials"],
  ]);
}
{
  const s = qSlide("Q2 Personas vs plan", "Q2  |  PERSONAS VS PLAN  (3/3)",
    "Are we reaching the thin-file, real-income customers our strategy targets?",
    "Yes. Gig workers convert 5–10x better than average (Freelancer 40%, Rider/Driver 22%), and unserved/underserved borrowers are 35% of loans vs a 30% target",
    "Note: occupation from AIS x KTB x OR x CLICX matched data; penetration = CLICX customers ÷ captive customers in that occupation; Others = unknown occupation\nSource: Data team occupation analysis (Oct 2026); CLICX 5-Year Business Plan scorecard (Dec 2026 target)");
  const occ = ["Freelancer", "Rider / Driver", "Blue collar", "White collar", "Merchant", "Others"];
  s.addChart(pres.charts.BAR, [{ name: "Penetration", labels: occ, values: [0.40, 0.22, 0.16, 0.09, 0.09, 0.004] }], {
    x: 0.5, y: 2.35, w: 4.2, h: 4.35, barDir: "bar", ...AX, ...quiet(), ...chartTitle("Penetration of captive base, %"),
    chartColors: [BLUE, BLUE, NAVY, NAVY, NAVY, GREY], showValue: true, dataLabelPosition: "outEnd", dataLabelColor: TXT, dataLabelFontSize: 11,
    dataLabelFormatCode: "0%", showLegend: false, barGapWidthPct: 45, valAxisMaxVal: 0.5, catAxisOrientation: "maxMin", objectName: "Occupation penetration" });
  s.addChart(pres.charts.BAR, [{ name: "CLICX customers, K", labels: occ, values: [232, 66, 401, 700, 293, 129] }], {
    x: 4.8, y: 2.35, w: 4.1, h: 4.35, barDir: "bar", ...AX, ...quiet(), ...chartTitle("CLICX customers, K"),
    chartColors: [GREY, GREY, NAVY, NAVY, NAVY, GREY], showValue: true, dataLabelPosition: "outEnd", dataLabelColor: TXT, dataLabelFontSize: 11,
    dataLabelFormatCode: "#,##0", showLegend: false, barGapWidthPct: 45, valAxisMaxVal: 850, catAxisOrientation: "maxMin", objectName: "Occupation volume" });
  insights(s, [
    ["Gig workers: best fit", "Freelancers and riders convert at 40% and 22% of their small captive pools, which validates the alternative-data lending thesis"],
    ["Collars: the volume", "White- and Blue-collar make up 1.1M customers, over 60% of the portfolio"],
    ["Merchants: proven, not scaled", "293K customers at 9% penetration from a 3.3M base, an entry point for SME supply-chain financing"],
  ]);
}

// =====================================================================
// Q3
pres.addSection({ title: "Q3 Persona x product" });
{
  const s = qSlide("Q3 Persona x product", "Q3  |  PERSONA x PRODUCT",
    "For different customer personas, do they come for the same or different products?",
    "They come for the same thing, credit: 84–92% of every occupation applies for a loan. What they end up using differs: White-collar books loans, Merchants keep deposits, gig workers are mostly declined",
    "Note: % of CLICX customers in each occupation. Deposit = balance > 0; Payment = transacted in Sep 2026; Booked = loan booked\nSource: Data team occupation analysis, AIS x KTB x OR x CLICX (Oct 2026)");
  const occ = ["White collar", "Merchant", "Blue collar", "Freelancer", "Rider / Driver"];
  s.addChart(pres.charts.BAR, [
    { name: "Applied for loan", labels: occ, values: [0.85, 0.84, 0.92, 0.88, 0.91] },
    { name: "Loan booked", labels: occ, values: [0.11, 0.07, 0.07, 0.06, 0.05] },
  ], { x: 0.5, y: 2.35, w: 4.2, h: 4.35, barDir: "bar", barGrouping: "clustered", ...AX, ...quiet(), ...chartTitle("Credit: applied vs booked, %"),
    chartColors: [LGREY, NAVY], showValue: true, dataLabelPosition: "outEnd", dataLabelColor: TXT, dataLabelFontSize: 10, dataLabelFormatCode: "0%",
    showLegend: true, legendPos: "b", legendFontSize: 10, legendColor: MUTED, barGapWidthPct: 40, valAxisMaxVal: 1.15, catAxisOrientation: "maxMin", objectName: "Credit funnel by occupation" });
  s.addChart(pres.charts.BAR, [
    { name: "Deposit > 0", labels: occ, values: [0.18, 0.20, 0.12, 0.16, 0.13] },
    { name: "Payment (Sep)", labels: occ, values: [0.14, 0.14, 0.10, 0.12, 0.11] },
  ], { x: 4.8, y: 2.35, w: 4.1, h: 4.35, barDir: "bar", barGrouping: "clustered", ...AX, ...quiet(), ...chartTitle("Save & pay usage, %"),
    chartColors: [BLUE, CYAN], showValue: true, dataLabelPosition: "outEnd", dataLabelColor: TXT, dataLabelFontSize: 10, dataLabelFormatCode: "0%",
    showLegend: true, legendPos: "b", legendFontSize: 10, legendColor: MUTED, barGapWidthPct: 40, valAxisMaxVal: 0.26, catAxisOrientation: "maxMin", objectName: "Save and pay by occupation" });
  insights(s, [
    ["Credit is the hook for all", "1.52M customers (84%) have applied for a loan, whatever their occupation"],
    ["White-collar: borrow and pay", "Highest booking rate (11%, 79K loans) and 14% pay monthly. The most complete product user"],
    ["Merchant: deposits", "Highest deposit rate (20%) and payment use (14%), a natural fit for business accounts and SME credit"],
    ["Gig and Blue-collar: leakage", "Over 90% apply, but only 5–7% are booked. Underwriting, not demand, is the bottleneck"],
  ]);
}

// =====================================================================
// Q4
pres.addSection({ title: "Q4 Active vs inactive" });
{
  const s = qSlide("Q4 Active vs inactive", "Q4  |  ACTIVE VS INACTIVE  (1/2)",
    "For active vs inactive CLICX customers, is the persona different?",
    "Yes. Only 1 in 5 customers is active, and the active base skews to salaried White-collar and Merchants, while inactive customers skew to Blue-collar and unknown occupations",
    "Note: Active = CLICX deposit > 0 or a booked CLICX loan. Inactive = all other CLICX customers. Totals: active 358K, inactive 1,462K\nSource: Data team occupation analysis, AIS x KTB x OR x CLICX (Oct 2026)");
  const grp = ["Active (358K)", "Inactive (1,462K)"];
  const act = { WC: 164, M: 68, BC: 64, F: 42, R: 10, O: 10 }, tot = { WC: 700, M: 293, BC: 401, F: 232, R: 66, O: 129 };
  const sa = 358, si = 1462, sh = (k) => [+(act[k] / sa).toFixed(3), +((tot[k] - act[k]) / si).toFixed(3)];
  s.addChart(pres.charts.BAR, [
    { name: "White collar", labels: grp, values: sh("WC") },
    { name: "Merchant", labels: grp, values: sh("M") },
    { name: "Blue collar", labels: grp, values: sh("BC") },
    { name: "Freelancer", labels: grp, values: sh("F") },
    { name: "Rider / Driver", labels: grp, values: sh("R") },
    { name: "Others", labels: grp, values: sh("O") },
  ], { x: 0.5, y: 2.35, w: 4.6, h: 4.35, barDir: "col", barGrouping: "percentStacked", ...AX, ...quiet(), ...chartTitle("Occupation mix, % of segment"),
    chartColors: [NAVY, BLUE, CYAN, LCYAN, GREY, LGREY], showValue: true, dataLabelPosition: "ctr", dataLabelColor: "FFFFFF", dataLabelFontSize: 10,
    dataLabelFormatCode: "0%", showLegend: true, legendPos: "r", legendFontSize: 10, legendColor: MUTED, barGapWidthPct: 45, objectName: "Active vs inactive mix" });
  const occ = ["White collar", "Merchant", "Freelancer", "Blue collar", "Rider / Driver", "Others"];
  s.addChart(pres.charts.BAR, [{ name: "Active rate", labels: occ, values: [0.23, 0.23, 0.18, 0.16, 0.15, 0.08] }], {
    x: 5.2, y: 2.35, w: 3.7, h: 4.35, barDir: "bar", ...AX, ...quiet(), ...chartTitle("Active rate by occupation, %"),
    chartColors: [NAVY, NAVY, GREY, GREY, GREY, GREY], showValue: true, dataLabelPosition: "outEnd", dataLabelColor: TXT, dataLabelFontSize: 11,
    dataLabelFormatCode: "0%", showLegend: false, barGapWidthPct: 45, valAxisMaxVal: 0.3, catAxisOrientation: "maxMin", objectName: "Active rate" });
  insights(s, [
    ["Salaried and Merchants stay", "White-collar and Merchant make up 65% of active customers but 52% of inactive ones, with an active rate of 23% each"],
    ["Blue-collar drops out", "23% of inactive customers vs 18% of active ones. They apply for credit, get declined and leave"],
    ["Unknown means inactive", "Customers with no occupation data are only 8% active, and 3 in 4 inactive customers never got a loan or funded an account"],
  ]);
}
{
  const s = qSlide("Q4 Active vs inactive", "Q4  |  ACTIVE VS INACTIVE  (2/2)",
    "What else separates active from inactive customers?",
    "Value-tier telco users and borrowers stay active. FBB and post-paid AIS customers are 2–4x more active than AIS-inactive ones, and borrowers pay 6x more often than declined depositors",
    "Note: Active = CLICX deposit > 0 or a booked CLICX loan; AIS segment base = 1.2M CLICX customers with AIS. Right chart: share of users by payment behaviour in CLICX\nSource: AIS x CLICX matched data; Data for Business Strategy (14 Oct 2026)");
  const seg = ["FBB only", "Post-paid", "Prepaid PPU", "Prepaid zero-ARPU", "Prepaid recurring", "Prepaid non-recurring", "AIS inactive"];
  s.addChart(pres.charts.BAR, [{ name: "CLICX active rate", labels: seg, values: [0.41, 0.28, 0.22, 0.21, 0.19, 0.15, 0.09] }], {
    x: 0.5, y: 2.35, w: 4.2, h: 4.35, barDir: "bar", ...AX, ...quiet(), ...chartTitle("CLICX active rate by AIS segment, %"),
    chartColors: [NAVY, NAVY, GREY, GREY, GREY, GREY, RED], showValue: true, dataLabelPosition: "outEnd", dataLabelColor: TXT, dataLabelFontSize: 11,
    dataLabelFormatCode: "0%", showLegend: false, barGapWidthPct: 40, valAxisMaxVal: 0.5, catAxisOrientation: "maxMin", catAxisLabelFontSize: 10, objectName: "AIS segment active rate" });
  const st = ["Loan booked (11%)", "Deposit, never rejected (22%)", "Deposit, ever rejected (67%)"];
  s.addChart(pres.charts.BAR, [
    { name: "No txn, no balance", labels: st, values: [0.26, 0.61, 0.83] },
    { name: "No txn, has balance", labels: st, values: [0.23, 0.23, 0.08] },
    { name: "Txn in 1 month", labels: st, values: [0.25, 0.08, 0.05] },
    { name: "Txn in 2+ months", labels: st, values: [0.26, 0.08, 0.03] },
  ], { x: 4.8, y: 2.35, w: 4.1, h: 4.35, barDir: "bar", barGrouping: "percentStacked", ...AX, ...quiet(), ...chartTitle("Payment behaviour by product status"),
    chartColors: [LGREY, GREY, CYAN, NAVY], showValue: true, dataLabelPosition: "ctr", dataLabelColor: TXT, dataLabelFontSize: 10, dataLabelFormatCode: "0%",
    showLegend: true, legendPos: "b", legendFontSize: 10, legendColor: MUTED, barGapWidthPct: 35, catAxisOrientation: "maxMin", catAxisLabelFontSize: 10, objectName: "Payment by product status" });
  insights(s, [
    ["Telco tier predicts activity", "FBB (41%) and post-paid (28%) users are most active; post-paid is 44% of the active AIS base"],
    ["A loan creates the habit", "51% of borrowers pay through CLICX vs 8% of customers whose loan was rejected"],
    ["Rejection drives inactivity", "67% of customers were rejected for credit at some point, and 83% of them never transact"],
  ]);
}

// =====================================================================
// Q5
pres.addSection({ title: "Q5 KTB cannibalization" });
{
  const s = qSlide("Q5 KTB cannibalization", "Q5  |  KTB CANNIBALIZATION  (1/2)",
    "Does CLICX cannibalize KTB customers?",
    "Limited. Only 13% of the 1.6M KTB x CLICX overlap are KTB-loyal customers; 61% are competitor-loyal or inactive with KTB, so CLICX mostly wins back or reactivates them",
    "Note: KTB active = holds at least 1 of 4 core KTB products. Inactive = no KTB transaction in 1 year (incl. Paotang-only without PaotangPay). Shared wallet = KTB swing, low engagement and PaotangPay users\nSource: KTB x CLICX engagement segmentation (Oct 2026)");
  const b = ["KTB-loyal (at risk)", "Shared wallet", "Competitor-loyal (win-back)", "KTB-inactive (reactivated)"];
  s.addChart(pres.charts.BAR, [{ name: "CLICX customers, K", labels: b, values: [212, 434, 515, 488] }], {
    x: 0.5, y: 2.35, w: 4.6, h: 4.35, barDir: "bar", ...AX, ...quiet(), ...chartTitle("Overlap by KTB relationship, K"),
    chartColors: [RED, GREY, BLUE, BLUE], showValue: true, dataLabelPosition: "outEnd", dataLabelColor: TXT, dataLabelFontSize: 11,
    dataLabelFormatCode: "#,##0", showLegend: false, barGapWidthPct: 40, valAxisMaxVal: 650, catAxisOrientation: "maxMin", objectName: "Overlap by KTB relationship" });
  const g = ["PaotangPay active", "KTB swing", "Competitor-loyal", "Paotang, no Pay", "KTB-loyal", "Inactive both", "KTB low engaged", "KTB inactive"];
  s.addChart(pres.charts.BAR, [{ name: "Penetration", labels: g, values: [0.12, 0.085, 0.079, 0.054, 0.033, 0.026, 0.010, 0.009] }], {
    x: 5.2, y: 2.35, w: 3.7, h: 4.35, barDir: "bar", ...AX, ...quiet(), ...chartTitle("Penetration by KTB group"),
    chartColors: [GREY, GREY, BLUE, GREY, RED, GREY, GREY, GREY], showValue: true, dataLabelPosition: "outEnd", dataLabelColor: TXT, dataLabelFontSize: 10,
    dataLabelFormatCode: "0.0%", showLegend: false, barGapWidthPct: 35, valAxisMaxVal: 0.15, catAxisOrientation: "maxMin", catAxisLabelFontSize: 10, objectName: "Penetration by KTB group" });
  insights(s, [
    ["Loyalists are hard to pull", "KTB-loyal penetration is 3.3%, less than half that of competitor-loyal (7.9%) and swing (8.5%) customers"],
    ["Defensive growth for KTB", "515K competitor-loyal customers, 79K of them quality users, now bank with the KTB ecosystem through CLICX"],
    ["Reactivation, not substitution", "488K were inactive with KTB for a year or more but signed up for CLICX"],
  ]);
}
{
  const s = qSlide("Q5 KTB cannibalization", "Q5  |  KTB CANNIBALIZATION  (2/2)",
    "Does CLICX cannibalize KTB products?",
    "Not on lending, where CLICX is additive. Its borrowers largely sit outside KTB's Jai-pump digital loan, and it catches customers Jai-pump rejects. Deposit cannibalization cannot yet be measured",
    "Note: Jai-pump = KTB's digital personal loan. Rejected = rejected by Jai-pump in the last 3 months. Quality = payment in last month, deposit > THB 100 or a performing loan. CLICX balances are not yet in the data room\nSource: KTB x CLICX matched data (Oct 2026)");
  s.addChart(pres.charts.BAR, [{ name: "CLICX customers, K", labels: ["CLICX borrowers with no Jai-pump loan", "CLICX applicants rejected by Jai-pump", "CLICX borrowers rejected by Jai-pump"], values: [116, 156, 23] }], {
    x: 0.5, y: 2.35, w: 4.3, h: 4.35, barDir: "col", ...AX, ...quiet(), ...chartTitle("Lending gained outside KTB, K"),
    chartColors: [BLUE, NAVY, NAVY], showValue: true, dataLabelPosition: "outEnd", dataLabelColor: TXT, dataLabelFontSize: 12,
    dataLabelFormatCode: "#,##0", showLegend: false, barGapWidthPct: 50, valAxisMaxVal: 190, catAxisLabelFontSize: 10, objectName: "Lending outside KTB" });
  s.addChart(pres.charts.DOUGHNUT, [{ name: "Quality CLICX customers", labels: ["Competitor-loyal", "KTB swing", "KTB-loyal", "KTB-inactive", "Other"], values: [79, 76, 38, 48, 10] }], {
    x: 4.9, y: 2.35, w: 4.0, h: 4.35, ...AX, ...chartTitle("Quality customers (251K) by KTB group"),
    chartColors: [BLUE, GREY, RED, CYAN, LGREY], holeSize: 55, showPercent: true, showValue: false, dataLabelColor: "FFFFFF", dataLabelFontSize: 10,
    showLegend: true, legendPos: "b", legendFontSize: 10, legendColor: MUTED, objectName: "Quality customers by KTB group" });
  insights(s, [
    ["Additive loan book", "116K CLICX borrowers hold no KTB Jai-pump loan, new credit for the ecosystem"],
    ["A safety net for KTB rejects", "156K applicants and 23K borrowers were rejected by Jai-pump in the last 3 months, demand that would otherwise go to competitors"],
    ["Value overlap is modest", "Only 15% of quality CLICX customers (38K) are KTB-loyal. Deposit overlap still needs testing once CLICX balances reach the data room"],
  ]);
}

// =====================================================================
// Implications
pres.addSection({ title: "Implications" });
{
  const s = qSlide("Implications", "IMPLICATIONS",
    "What should we do with these answers?",
    "Keep the consortium as the acquisition engine, but move from sign-ups to active, credit-led customers in the personas that already work",
    "Source: synthesis of Q1–Q5; actions build on Data for Business Strategy (14 Oct 2026) and CLICX 5-Year Business Plan");
  const cards = [
    ["Q1", "Mine the 47M base", "Go after AIS first (3.7% penetration on ~32M), through myAIS and Dipchip, starting with 1.1M active AIS customers"],
    ["Q2", "Double down on what fits", "Use Tech-Savvy aged 30–34 as the look-alike seed. Build a Multiple-Bank offer and a proxy-income journey for gig workers"],
    ["Q3", "Fix the credit funnel", "Over 90% of Blue-collar and gig workers apply but 5–7% are booked. Recalibrate underwriting and offer deposit-backed credit-builder loans"],
    ["Q4", "Activate after onboarding", "Get new customers onto 3+ use cases in 30 days. Give declined applicants a non-loan reason to stay; target prepaid and AIS-inactive customers"],
    ["Q5", "Position CLICX as KTB's growth arm", "Report win-back (515K) and reactivation (488K) to KTB. Formalise Jai-pump reject referral; measure deposit overlap once data lands"],
  ];
  const cw = 2.33, gap = 0.17;
  cards.forEach((c, i) => {
    const x = 0.5 + i * (cw + gap), y = 2.4;
    s.addShape(pres.shapes.RECTANGLE, { x, y, w: cw, h: 4.25, fill: { color: C.background2 }, line: { type: "none" }, objectName: `Card ${c[0]}` });
    s.addShape(pres.shapes.OVAL, { x: x + 0.2, y: y + 0.22, w: 0.55, h: 0.55, fill: { color: C.text2 }, line: { type: "none" }, objectName: `Card badge ${c[0]}` });
    s.addText(c[0], { x: x + 0.2, y: y + 0.22, w: 0.55, h: 0.55, fontSize: 13, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: `Card badge text ${c[0]}` });
    s.addText(c[1], { x: x + 0.2, y: y + 0.95, w: cw - 0.4, h: 0.85, fontSize: 15, bold: true, color: C.text2, valign: "top", margin: 0, isTextBox: true, objectName: `Card title ${c[0]}` });
    s.addText(c[2], { x: x + 0.2, y: y + 1.75, w: cw - 0.4, h: 2.35, fontSize: 12, color: C.text1, valign: "top", margin: 0, isTextBox: true, objectName: `Card body ${c[0]}` });
  });
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
