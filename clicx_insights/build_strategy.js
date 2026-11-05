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
pres.title = "CLICX Data for Strategy 2027";
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
// Data for Strategy 2027: storyline deck (source: 20261105 - Data for Strategy 2027)
const SRC = "Source: Data for Strategy 2027 (Nov 2026); CLICX data as of Sep 2026; AIS, KTB, OR data as of Apr 2026";
const hbar = (s, labels, values, o) => s.addChart(pres.charts.BAR, [{ name: o.name || "Value", labels, values }], {
  x: o.x, y: o.y || 2.35, w: o.w, h: o.h || 4.35, barDir: "bar", ...AX, ...quiet(), ...chartTitle(o.title),
  chartColors: o.colors, showValue: true, dataLabelPosition: "outEnd", dataLabelColor: TXT, dataLabelFontSize: o.fs || 11,
  dataLabelFormatCode: o.fmt || "0%", showLegend: false, barGapWidthPct: o.gap || 40, valAxisMaxVal: o.max, catAxisOrientation: "maxMin",
  catAxisLabelFontSize: o.lfs || 11, objectName: o.obj });

pres.addSection({ title: "Opening" });
{
  const s = pres.addSlide({ masterName: "TITLE", sectionTitle: "Opening" });
  s.addText("DATA FOR STRATEGY 2027", { placeholder: "kicker" });
  s.addText("What have our first 2.1 million customers taught us for 2027?", { placeholder: "title" });
  s.addText("Where customers come from, whether we reached our targets, which customers lead to new products, why most are inactive, and how we protect KTB", { placeholder: "sub" });
  s.addText("November 2026  |  CLICX data as of Sep 2026; AIS, KTB, OR data as of Apr 2026", { placeholder: "date" });
}
{
  const s = pres.addSlide({ masterName: "QUESTION", sectionTitle: "Opening" });
  s.addText("EXECUTIVE SUMMARY", { placeholder: "tracker" });
  s.addText("What does the data tell us for 2027?", { placeholder: "title" });
  s.addText("Partners bring in almost everyone and we reached the people we targeted. Our customers are ready lead pools for 2027 products, but only 15% are active, and a loan is what makes them active", { placeholder: "answer" });
  s.addText(SRC, { placeholder: "foot" });
  const hdr = (t) => ({ text: t, options: { bold: true, color: "FFFFFF", fill: { color: NAVY }, fontSize: 12 } });
  const v = (t, col = BLUE) => ({ text: t, options: { bold: true, color: col, fontSize: 13 } });
  const c = (t) => ({ text: t, options: { fontSize: 12, color: TXT } });
  s.addTable([
    [hdr("#"), hdr("Question"), hdr("Short answer"), hdr("Why we say so")],
    [c("1"), c("Do our customers come from our partners?"), v("Almost all"), c("96% of the 1.86M customers who gave consent are AIS, KTB or OR customers. KTB brings 88%")],
    [c("2"), c("Are we winning the people we targeted?"), v("Mostly"), c("Tech-savvy people join at 8.5%, over 2x the average. 82% of loan applicants are underserved. Multi-bank users lag at 2.4%")],
    [c("3"), c("Which customers lead to our 2027 products?"), v("Clear lead pools"), c("293K merchants for SME loans, 700K office workers for cards, 5.5B THB of deposits for wealth")],
    [c("4"), c("Why are most customers inactive?"), v("No loan, no money in", RED), c("Only 15% transact monthly. 88% of borrowers are active, vs 4% of people turned down who never put money in")],
    [c("5"), c("Do we take from KTB, and should we fear True?"), v("We recover, not take"), c("1.1M shared customers were inactive with KTB. 450K people KTB turned down applied to us. True sizing is still to come")],
  ], { x: 0.5, y: 2.4, w: 12.33, colW: [0.45, 3.6, 2.1, 6.18], fontFace: "Arial", border: { type: "solid", pt: 0.75, color: LGREY },
    valign: "middle", rowH: [0.4, 0.78, 0.78, 0.78, 0.78, 0.78], margin: [0.05, 0.1, 0.05, 0.1], objectName: "Summary table" });
}

// ---------- 1 Partners
pres.addSection({ title: "1 Where customers come from" });
{
  const s = qSlide("1 Where customers come from", "1  |  WHERE CUSTOMERS COME FROM",
    "Do our customers come from our partners?",
    "Almost all. Of 2.1M CLICX users, 1.86M gave partner consent and 96% of them are AIS, KTB or OR customers. Yet each partner has only 4–7% of its base on CLICX",
    "Note: one customer can be with several partners, so shares add up to more than 100%. " + SRC);
  stat(s, 0.5, 2.4, 2.6, "2.1M", "CLICX users; 1.86M (86%) gave partner consent", C.text2);
  stat(s, 0.5, 3.85, 2.6, "96%", "of consented customers are AIS, KTB or OR customers");
  stat(s, 0.5, 5.3, 2.6, "3.7%", "of AIS customers have joined: the lowest of the three", C.text2);
  hbar(s, ["KTB (4.7% of its base joined)", "AIS (3.7%)", "OR (7.4%)"], [0.88, 0.67, 0.40],
    { x: 3.3, w: 5.6, title: "% of consented CLICX customers who are also with each partner", colors: [NAVY, NAVY, NAVY], max: 1.05, lfs: 12, fs: 12, gap: 60, obj: "Partner reach" });
  insights(s, [
    ["KTB is the front door", "1.6M customers, 88% of those who gave consent"],
    ["OR converts best", "7.4% of OR's base joined, though OR brings the fewest (725K)"],
    ["AIS is the biggest gap", "1.2M joined, only 3.7% of AIS's base; 0.66M of our customers are not AIS yet"],
  ]);
}

// ---------- 2 Targets
pres.addSection({ title: "2 Target groups" });
{
  const s = qSlide("2 Target groups", "2  |  TARGET GROUPS  (1/2)",
    "Are we winning the people we targeted?",
    "Mostly. Tech-savvy people join at 8.5%, over twice the average, and 82% of loan applicants are underserved. Multi-bank users and over-60s are not won yet",
    "Underserved = unstable income and no income documents, by credit bureau (NCB) classification, 27 Oct 2026. " + SRC);
  hbar(s, ["Tech-savvy (533K)", "Working Gen Z, 23–29 (176K)", "Working, 30–59 (670K)", "One bank, AIS (193K)", "Multi-bank (194K)", "Under 23 (26K)", "60+ (15K)", "KTB Next only (10K)"],
    [0.085, 0.066, 0.047, 0.039, 0.024, 0.015, 0.003, 0.003],
    { x: 0.5, w: 4.9, title: "Early-adopter target: % of partner customers who joined (average 3.8%)", colors: [BLUE, BLUE, NAVY, NAVY, RED, GREY, GREY, GREY], max: 0.105, fmt: "0.0%", lfs: 10, obj: "Target penetration" });
  s.addChart(pres.charts.DOUGHNUT, [{ name: "Loan applicants", labels: ["Underserved (1,450K)", "Served (304K)", "Unserved (16K)"], values: [1450, 304, 16] }], {
    x: 5.5, y: 2.35, w: 3.4, h: 4.35, ...AX, ...chartTitle("Underserved target: loan applicants by credit access"),
    chartColors: [BLUE, GREY, NAVY], holeSize: 58, showPercent: true, showValue: false, dataLabelColor: "FFFFFF", dataLabelFontSize: 11,
    showLegend: true, legendPos: "b", legendFontSize: 10, legendColor: MUTED, objectName: "Credit access" });
  insights(s, [
    ["Tech-savvy: target hit", "533K customers, 8.5% of the group joined. Working Gen Z (23–29) is next at 6.6%"],
    ["Underserved: target hit", "82% of 1.77M loan applicants have no easy access to bank credit"],
    ["Multi-bank: not won yet", "8.2M people, only 2.4% joined. They need a stronger reason to switch"],
  ]);
}
{
  const s = qSlide("2 Target groups", "2  |  TARGET GROUPS  (2/2)",
    "Do the underserved get what they came for?",
    "Not yet. Over 90% of blue-collar and gig workers apply for a loan, but only 5–7% get one. Merchants are different: they keep more money with us and spend more",
    SRC);
  const occ = ["Office worker", "Merchant", "Blue-collar worker", "Freelancer", "Rider / driver"];
  s.addChart(pres.charts.BAR, [
    { name: "Applied for a loan", labels: occ, values: [0.85, 0.84, 0.92, 0.88, 0.91] },
    { name: "Got a loan", labels: occ, values: [0.11, 0.07, 0.07, 0.06, 0.05] },
  ], { x: 0.5, y: 2.35, w: 4.4, h: 4.35, barDir: "bar", barGrouping: "clustered", ...AX, ...quiet(), ...chartTitle("Loans: applied vs got one, % of each job type"),
    chartColors: [LGREY, NAVY], showValue: true, dataLabelPosition: "outEnd", dataLabelColor: TXT, dataLabelFontSize: 10, dataLabelFormatCode: "0%",
    showLegend: true, legendPos: "b", legendFontSize: 10, legendColor: MUTED, barGapWidthPct: 40, valAxisMaxVal: 1.15, catAxisOrientation: "maxMin", objectName: "Applied vs got" });
  hbar(s, ["Office worker", "Merchant", "Freelancer", "Blue-collar worker", "Rider / driver"], [4550, 3303, 2854, 518, 275],
    { x: 5.0, w: 3.9, title: "Average balance per customer, THB", colors: [NAVY, BLUE, NAVY, GREY, GREY], max: 5600, fmt: "#,##0", obj: "Avg balance by job" });
  insights(s, [
    ["Demand is there, approval isn't", "Riders: 91% apply, 5% get a loan. Blue-collar: 92% apply, 7% get one"],
    ["Merchants save and spend", "20% keep money with us (highest) and hold 951M THB, avg 3,303 THB"],
    ["Office workers hold the most", "3,164M THB in total, avg 4,550 THB, and the highest loan rate (11%)"],
  ]);
}

// ---------- 3 Portfolio to products
pres.addSection({ title: "3 Customers to products" });
{
  const s = qSlide("3 Customers to products", "3  |  CUSTOMERS TO PRODUCTS  (1/2)",
    "Who are our customers by life stage, and what should we offer each?",
    "Career builders (25–39) are our core at 1.0M, but working seniors (40–59) hold over half of all deposits. Each life stage needs a different offer",
    "Average balance shown with the median in brackets: most customers keep very little. " + SRC);
  const H = (t) => ({ text: t, options: { bold: true, color: "FFFFFF", fill: { color: NAVY }, fontSize: 10.5 } });
  const c = (t, o = {}) => ({ text: t, options: { fontSize: 10.5, color: TXT, ...o } });
  const b = (t) => ({ text: t, options: { fontSize: 10.5, bold: true, color: NAVY } });
  s.addTable([
    [H("Life stage"), H("Count"), H("Applied for loan"), H("Avg balance, THB"), H("What to offer")],
    [b("Student (<20)"), c("12K"), c("45%"), c("674 (18)"), c("Gamified saving")],
    [b("First jobber (20–24)"), c("163K"), c("85%"), c("415 (2)"), c("Lifestyle spending")],
    [b("Career builder (25–39)"), c("1.0M", { bold: true }), c("85%"), c("1,689 (7)"), c("Core lending and card")],
    [b("Working senior (40–59)"), c("612K"), c("77%"), c("4,496 (52)"), c("Wealth and refinance")],
    [b("Senior (60+)"), c("31K"), c("41%"), c("21,362 (1,000)"), c("Keep deposits; watch funding cost")],
  ], { x: 0.5, y: 2.45, w: 5.3, colW: [1.45, 0.8, 0.8, 1.05, 1.2], fontFace: "Arial", border: { type: "solid", pt: 0.75, color: LGREY },
    valign: "middle", rowH: [0.5, 0.7, 0.7, 0.7, 0.7, 0.7], margin: [0.04, 0.06, 0.04, 0.06], objectName: "Life stage table" });
  hbar(s, ["Student", "First jobber", "Career builder", "Working senior", "Senior"], [9.4, 70, 1774, 2916, 822],
    { x: 5.95, w: 2.95, title: "Total deposits, million THB", colors: [GREY, GREY, NAVY, BLUE, NAVY], max: 3600, fmt: "#,##0", obj: "Deposits by life stage" });
  insights(s, [
    ["Career builders are the core", "1.0M customers, 85% applied for a loan: the base for lending and cards"],
    ["Working seniors hold the money", "612K customers hold 2.9B THB, 52% of deposits: the base for wealth"],
    ["Seniors: small but costly", "31K customers hold 822M THB. Keep them, but watch the cost of funds"],
  ]);
}
{
  const s = qSlide("3 Customers to products", "3  |  CUSTOMERS TO PRODUCTS  (2/2)",
    "Which customers can we turn into 2027 revenue?",
    "Three ready lead pools already in our base: merchants for SME loans, office workers for a credit card, and depositors for wealth products",
    "Revenue and conversion targets are still to be set. " + SRC);
  const cards = [
    ["SME lending", "293K", "merchants", "245K already applied for a loan; mostly outside Bangkok", "Merchants save and spend more than other underserved groups"],
    ["Digital credit card", "700K", "office workers", "158K proven borrowers + 239K AIS Serenade members", "Young office workers who never had a card; partner perks with AIS and PTT"],
    ["Wealth and investment", "5.5B THB", "of deposits", "643K working seniors and seniors hold most of it", "Tech-savvy savers ready to move from Save Max to funds"],
  ];
  cards.forEach((cd, i) => {
    const x = 0.5 + i * 2.85, y = 2.4, w = 2.7;
    s.addShape(pres.shapes.RECTANGLE, { x, y, w, h: 4.25, fill: { color: C.background2 }, line: { type: "none" }, objectName: `Pool ${i}` });
    s.addText(cd[0].toUpperCase(), { x: x + 0.2, y: y + 0.2, w: w - 0.4, h: 0.3, fontSize: 11, bold: true, color: C.accent1, charSpacing: 1, margin: 0, isTextBox: true, objectName: `Pool name ${i}` });
    s.addText(cd[1], { x: x + 0.2, y: y + 0.55, w: w - 0.4, h: 0.65, fontSize: 36, bold: true, color: C.text2, margin: 0, isTextBox: true, objectName: `Pool size ${i}` });
    s.addText(cd[2], { x: x + 0.2, y: y + 1.2, w: w - 0.4, h: 0.3, fontSize: 13, color: C.text1, margin: 0, isTextBox: true, objectName: `Pool unit ${i}` });
    s.addText([
      { text: "Leads", options: { bold: true, fontSize: 12, color: C.text2, breakLine: true } },
      { text: cd[3], options: { fontSize: 12, color: C.text1, breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "Why it fits", options: { bold: true, fontSize: 12, color: C.text2, breakLine: true } },
      { text: cd[4], options: { fontSize: 12, color: C.text1, breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "2027 target: to be set", options: { fontSize: 12, bold: true, color: C.accent5 } },
    ], { x: x + 0.2, y: y + 1.65, w: w - 0.4, h: 2.45, valign: "top", margin: 0, isTextBox: true, objectName: `Pool body ${i}` });
  });
  s.addText([
    { text: "WHAT WE NEED", options: { bold: true, fontSize: 11, color: C.accent1, charSpacing: 1, breakLine: true } },
    { text: " ", options: { fontSize: 6, breakLine: true } },
    { text: "Targets for each product: leads, conversion and revenue", options: { fontSize: 12, bullet: true, breakLine: true } },
    { text: "Merchant detail: real job, age and repayment record", options: { fontSize: 12, bullet: true, breakLine: true } },
    { text: "How many customers already invest elsewhere", options: { fontSize: 12, bullet: true } },
  ], { x: 9.05, y: 2.4, w: 3.78, h: 4.25, fill: { color: C.background2 }, margin: [12, 18, 12, 18], valign: "top", paraSpaceAfter: 6, isTextBox: true, objectName: "Needs panel" });
}

// ---------- 4 Activity
pres.addSection({ title: "4 Activity" });
{
  const s = qSlide("4 Activity", "4  |  WHO IS ACTIVE",
    "Why are most of our customers inactive?",
    "Only 15% transact each month. A loan makes the difference: 88% of borrowers are active, but people turned down who never put money in are almost all inactive (4%)",
    "Monthly transacting = any payment, money in, ATM use, pocket opened, coupon, loan taken or repaid, or mission joined (Sep 2026). " + SRC);
  hbar(s, ["All customers", "Have a deposit account", "Opened the app this month", "Transacted this month"], [2144, 1449, 690, 312],
    { x: 0.5, w: 3.9, title: "From customer to active user, thousand", colors: [NAVY, NAVY, BLUE, RED], max: 2600, fmt: "#,##0", lfs: 10, obj: "Activity funnel" });
  hbar(s, ["Borrowers (158K)", "Money in, never applied (48K)", "Money in, applied (80K)", "No money in, never applied (113K)", "No money in, applied (924K)"], [0.88, 0.52, 0.50, 0.07, 0.04],
    { x: 4.5, w: 4.4, title: "% transacting monthly, by what they hold", colors: [NAVY, BLUE, BLUE, GREY, RED], max: 1.05, lfs: 10, obj: "Activity by holding" });
  insights(s, [
    ["1.8M customers are idle", "694K hold no product; 1.1M hold one but don't use it"],
    ["A loan keeps people active", "88% of borrowers transact monthly, more than any other group"],
    ["Rejection loses customers", "924K applied, got no loan and put no money in: only 4% are active"],
  ]);
}

// ---------- 5 KTB and True
pres.addSection({ title: "5 KTB and True" });
{
  const s = qSlide("5 KTB and True", "5  |  KTB AND TRUE  (1/2)",
    "Does CLICX take customers away from KTB?",
    "No, it recovers them. 1.1M of the 1.6M customers we share with KTB were inactive with KTB, and 450K people KTB's Jai-pump loan turned down applied to us",
    "Active with KTB = 1.6M minus the two inactive groups. Jai-pump rejects from Jun 2025 to May 2026. Early default = missed first bill (FPDX), on borrowers who reached their first bill. " + SRC);
  hbar(s, ["Active with KTB", "Loyal to other banks", "Dormant or Paotang only"], [537, 515, 548],
    { x: 0.5, w: 3.9, title: "1.6M customers we share with KTB, thousand", colors: [GREY, BLUE, BLUE], max: 700, fmt: "#,##0", lfs: 11, obj: "Shared with KTB" });
  const cx = 6.65;
  s.addText("From Jai-pump rejection to CLICX loan", { x: 4.5, y: 2.4, w: 4.4, h: 0.35, fontSize: 13, bold: true, color: C.text2, align: "center", margin: 0, isTextBox: true, objectName: "Funnel title" });
  const stage = (w, y, fill, l, v, name) => {
    s.addShape(pres.shapes.RECTANGLE, { x: cx - w / 2, y, w, h: 0.8, fill: { color: fill }, line: { type: "none" }, objectName: name });
    s.addText([{ text: l, options: { fontSize: 12, color: "FFFFFF", breakLine: true } }, { text: v, options: { fontSize: 18, bold: true, color: "FFFFFF" } }],
      { x: cx - 2.1, y, w: 4.2, h: 0.8, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: name + " text" });
  };
  const step = (t, y) => s.addText(t, { x: 4.5, y, w: 4.4, h: 0.28, fontSize: 11, bold: true, color: C.accent1, align: "center", margin: 0, isTextBox: true, objectName: "Step " + y });
  stage(4.2, 2.9, C.text2, "Turned down by Jai-pump", "1.9M", "Stage 1");
  step("24% applied to CLICX", 3.73);
  stage(3.0, 4.03, C.text2, "Applied to CLICX", "450K", "Stage 2");
  step("11% approved", 4.86);
  stage(1.9, 5.16, C.accent1, "Got a CLICX loan", "51K", "Stage 3");
  s.addText("Early default: 20.2%", { x: 4.5, y: 6.05, w: 4.4, h: 0.3, fontSize: 11, color: MUTED, align: "center", margin: 0, isTextBox: true, objectName: "FPDX" });
  insights(s, [
    ["We bring back lapsed KTB users", "515K mainly bank elsewhere and 548K were dormant or used Paotang only"],
    ["Paotang rejects are the best leads", "They apply more (30% vs 23% via NEXT) and default less (18% vs 21%)"],
    ["Next: work the reject list", "Target Paotang rejects under 40, especially freelancers and riders (64% apply, 7% approved)"],
  ]);
}
{
  const s = qSlide("5 KTB and True", "5  |  KTB AND TRUE  (2/2)",
    "Should we worry about True's new virtual bank?",
    "Partly. 2.5M of KTB's 35M consented customers already use TrueMoney. Customers KTB turns down are the ones most open to a new bank, so CLICX should reach them first",
    "KTB customers who gave consent in the data room. Turned down = rejected by Jai-pump. " + SRC);
  stat(s, 0.5, 2.4, 2.4, "35M", "KTB customers who gave consent", C.text2);
  stat(s, 0.5, 3.85, 2.4, "2.5M", "of them already use TrueMoney", C.accent5);
  s.addChart(pres.charts.BAR, [{ name: "Turned down by Jai-pump", labels: ["Have CLICX, KTB active (144K)", "Have CLICX, KTB inactive (81K)", "No CLICX, KTB active (1,652K)", "No CLICX, KTB inactive (660K)"], values: [0.53, 0.55, 0.13, 0.15] }], {
    x: 3.0, y: 2.35, w: 3.3, h: 4.35, barDir: "col", ...AX, ...quiet(), ...chartTitle("TrueMoney users turned down by Jai-pump, %"),
    chartColors: [BLUE, BLUE, GREY, GREY], showValue: true, dataLabelPosition: "outEnd", dataLabelColor: TXT, dataLabelFontSize: 12,
    dataLabelFormatCode: "0%", showLegend: false, barGapWidthPct: 50, valAxisMaxVal: 0.7, catAxisLabelFontSize: 9, objectName: "Reject share" });
  s.addShape(pres.shapes.RECTANGLE, { x: 6.45, y: 2.45, w: 2.45, h: 4.15, fill: { color: C.background2 }, line: { color: GREY, width: 1, dashType: "dash" }, objectName: "True gap box" });
  s.addText([
    { text: "STILL TO SIZE", options: { bold: true, fontSize: 11, color: C.accent5, charSpacing: 1, breakLine: true } },
    { text: " ", options: { fontSize: 6, breakLine: true } },
    { text: "AIS vs True: who would move their deposits", options: { fontSize: 12, bullet: true, breakLine: true } },
    { text: "Age and profile of TrueMoney users", options: { fontSize: 12, bullet: true, breakLine: true } },
    { text: "True customers: most applied for a loan, and loans drive activity", options: { fontSize: 12, bullet: true } },
  ], { x: 6.6, y: 2.6, w: 2.2, h: 3.9, valign: "top", margin: 0, paraSpaceAfter: 6, isTextBox: true, objectName: "True gap text" });
  insights(s, [
    ["TrueMoney users we have are KTB rejects", "53–55% of TrueMoney users with CLICX were turned down by Jai-pump"],
    ["1.6M KTB-active TrueMoney users are not on CLICX", "Only 13% of them were rejected: a different, safer group to win"],
    ["Rather CLICX than True", "Reach KTB's turned-down and TrueMoney customers before True launches"],
  ]);
}

// ---------- Next steps
pres.addSection({ title: "Next steps" });
{
  const s = qSlide("Next steps", "NEXT STEPS",
    "What should we do in 2027?",
    "Grow from partners, fix loan approvals, activate idle customers, turn lead pools into new products, and get to KTB's turned-down customers before True does",
    SRC);
  const cards = [
    ["1", "Reach more partner customers", "Start with AIS: 3.7% joined and 0.66M of our customers are not AIS yet"],
    ["2", "Say 'yes' more safely", "Over 90% of gig and blue-collar workers apply but only 5–7% get a loan"],
    ["3", "Activate idle customers", "1.8M are idle. Get money in and a first loan: 88% of borrowers stay active"],
    ["4", "Launch to lead pools", "SME loans to 293K merchants, card to 700K office workers, wealth on 5.5B THB"],
    ["5", "Defend KTB against True", "Work the Paotang reject list and size the True threat before launch"],
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
