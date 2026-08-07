/**
 * Book Loan - single-slide summary of Lending_2026.08.07.pptx.
 *
 * Every figure on the slide comes from that deck: the Summary slides
 * (total booked, consortium penetration, cash need, risk level) and the
 * four AIS Base Information slides (charge type, nano loan, true, tenure).
 * Percentages not printed in the source are computed from its own counts.
 *
 *   node build_lending_slide.js
 */

const path = require("path");
const PptxGenJS = require(path.join(
  "/tmp/claude-0/-home-user-claude-code/a5758311-6270-5365-a679-5836c8e934f0/scratchpad",
  "node_modules",
  "pptxgenjs"
));

/* ---------------------------------------------------------------- palette */

const NAVY = "0E1E3C";
const INK = "16233D";
const TEAL = "00A896";
const MINT = "5FD6BE";
const AMBER = "E9A03B";
const PAPER = "FFFFFF";
const TINT = "F1F4F9";
const TINT_COOL = "E8F6F3";
const MUTED = "6B7A90";
const MUTED_LT = "A8B4C6";
// sequential ramp, dark enough that white data labels read on every step
const R1 = "8DA0B8"; // null / none
const R2 = "3FB89F"; // low
const R3 = "1E8E86"; // mid
const R4 = "165A73"; // high
const RULE = "D8DFE9"; // table borders

const H = "Cambria";
const B = "Calibri";

const fs = require("fs");
function tsv(name) {
  const lines = fs.readFileSync(path.join(__dirname, "data", name), "utf8").trim().split("\n");
  const head = lines[0].split("\t");
  return lines.slice(1).map((l) => {
    const c = l.split("\t");
    const o = {};
    head.forEach((h, i) => (o[h] = c[i]));
    return o;
  });
}
const num = (x) => Number(String(x).replace(/,/g, ""));
const fmt = (v) => v.toLocaleString("en-US");
const pct = (a, b) => (b ? (a / b) * 100 : 0).toFixed(1) + "%";
const M = (v) => Number((v / 1e6).toFixed(1)).toLocaleString("en-US") + "M";


const pres = new PptxGenJS();
pres.layout = "LAYOUT_WIDE";
pres.author = "Lending";
pres.title = "Book Loan - summary as of 5 Aug";

const s = pres.addSlide();
s.background = { color: PAPER };

/* ------------------------------------------------------------------ title */

s.addText("BOOK LOAN  ·  TOTAL BOOKED AS OF 5 AUG", {
  x: 0.6,
  y: 0.36,
  w: 12.1,
  h: 0.26,
  margin: 0,
  fontFace: B,
  fontSize: 11.5,
  bold: true,
  charSpacing: 2.2,
  color: TEAL,
});
s.addText("Nine in ten booked loans are DL", {
  x: 0.6,
  y: 0.66,
  w: 12.1,
  h: 0.72,
  margin: 0,
  valign: "top",
  fontFace: H,
  fontSize: 32,
  bold: true,
  color: INK,
});

/* ------------------------------------------------------------ stat tiles */

const tiles = [
  ["94,371", "total booked", "as of 5 Aug", TEAL],
  ["90.4%", "of the book is DL", "85,317 DL vs 9,054 PL", TEAL],
  ["64.5%", "sit at high cash need", "60,839 customers", AMBER],
  ["99.2%", "are in a consortium", "only 794 in none", TEAL],
];
tiles.forEach(([v, l, sub, c], i) => {
  const x = 0.6 + i * 3.1;
  s.addShape("roundRect", {
    x,
    y: 1.46,
    w: 2.8,
    h: 1.1,
    rectRadius: 0.09,
    fill: { color: TINT },
    line: { type: "none" },
  });
  s.addText(v, {
    x: x + 0.25,
    y: 1.56,
    w: 2.3,
    h: 0.44,
    margin: 0,
    fontFace: H,
    fontSize: 25,
    bold: true,
    color: c,
  });
  s.addText(l, {
    x: x + 0.25,
    y: 2.0,
    w: 2.3,
    h: 0.28,
    margin: 0,
    valign: "top",
    fontFace: B,
    fontSize: 11.5,
    bold: true,
    color: INK,
  });
  s.addText(sub, {
    x: x + 0.25,
    y: 2.26,
    w: 2.3,
    h: 0.26,
    margin: 0,
    valign: "top",
    fontFace: B,
    fontSize: 10,
    color: MUTED,
  });
});

/* ------------------------------------------- left panel: consortium bars */

s.addShape("roundRect", {
  x: 0.6,
  y: 2.76,
  w: 5.95,
  h: 2.52,
  rectRadius: 0.09,
  fill: { color: TINT },
  line: { type: "none" },
});
s.addText("Consortium penetration", {
  x: 0.9,
  y: 2.94,
  w: 5.35,
  h: 0.3,
  margin: 0,
  fontFace: H,
  fontSize: 15,
  bold: true,
  color: INK,
});
s.addText("Share of the 94,371 booked customers found in each base", {
  x: 0.9,
  y: 3.24,
  w: 5.35,
  h: 0.26,
  margin: 0,
  fontFace: B,
  fontSize: 10,
  color: MUTED,
});

const consortium = [
  ["KTB", 88, "83,362", R4],
  ["AIS", 74, "70,100", R3],
  ["OR", 57, "53,712", R2],
  ["Not in any", 0.8, "794", R1],
];
consortium.forEach(([name, pct, count, color], i) => {
  const y = 3.62 + i * 0.4;
  s.addText(name, {
    x: 0.9,
    y,
    w: 1.15,
    h: 0.3,
    margin: 0,
    valign: "middle",
    fontFace: B,
    fontSize: 11.5,
    bold: true,
    color: INK,
  });
  s.addShape("roundRect", {
    x: 2.15,
    y: y + 0.05,
    w: Math.max((pct / 100) * 2.6, 0.06),
    h: 0.2,
    rectRadius: 0.04,
    fill: { color },
    line: { type: "none" },
  });
  s.addText(`${count}  ·  ${pct}%`, {
    x: 4.9,
    y,
    w: 1.35,
    h: 0.3,
    margin: 0,
    align: "right",
    valign: "middle",
    fontFace: H,
    fontSize: 11.5,
    bold: true,
    color: INK,
  });
});

/* ------------------------------- right panel: cash need and risk mix */

s.addShape("roundRect", {
  x: 6.75,
  y: 2.76,
  w: 5.95,
  h: 2.52,
  rectRadius: 0.09,
  fill: { color: TINT },
  line: { type: "none" },
});

const cats = ["Cash need · DL", "Cash need · PL", "Risk · DL", "Risk · PL"];
s.addChart(
  pres.charts.BAR,
  [
    { name: "Null", labels: cats, values: [13, 18, 2, 6] },
    { name: "Low", labels: cats, values: [5, 9, 29, 31] },
    { name: "Mid", labels: cats, values: [16, 21, 25, 23] },
    { name: "High", labels: cats, values: [66, 52, 44, 40] },
  ],
  {
    x: 6.88,
    y: 2.86,
    w: 5.72,
    h: 2.34,
    barDir: "bar",
    barGrouping: "percentStacked",
    barGapWidthPct: 45,
    chartColors: [R1, R2, R3, R4],
    showTitle: true,
    title: "Cash need and risk mix (% of customers)",
    titleFontFace: B,
    titleFontSize: 12,
    titleColor: INK,
    showValue: true,
    dataLabelPosition: "ctr",
    dataLabelFontFace: B,
    dataLabelFontSize: 9,
    dataLabelColor: PAPER,
    dataLabelFormatCode: '0"%"',
    showLegend: true,
    legendPos: "b",
    legendFontFace: B,
    legendFontSize: 10,
    legendColor: INK,
    catAxisLabelFontFace: B,
    catAxisLabelFontSize: 10.5,
    catAxisLabelColor: INK,
    valAxisHidden: true,
    valGridLine: { style: "none" },
    catGridLine: { style: "none" },
  }
);

/* ------------------------------------------------- bottom band: AIS base */

s.addShape("roundRect", {
  x: 0.6,
  y: 5.48,
  w: 12.1,
  h: 1.26,
  rectRadius: 0.09,
  fill: { color: NAVY },
  line: { type: "none" },
});
s.addText("AIS BASE INFORMATION  ·  70.1K", {
  x: 1.0,
  y: 5.66,
  w: 11.3,
  h: 0.26,
  margin: 0,
  fontFace: B,
  fontSize: 10.5,
  bold: true,
  charSpacing: 1.8,
  color: MINT,
});
s.addText(
  [
    { text: "68% sit in the TRUE ecosystem", options: { color: MINT, bold: true } },
    { text: " (47,840) · ", options: { color: "C3D0E2" } },
    { text: "61% carry the nano-loan flag", options: { color: MINT, bold: true } },
    { text: " (42,806) · ", options: { color: "C3D0E2" } },
    { text: "66% have tenure of 5 years or more", options: { color: MINT, bold: true } },
    { text: " (46,441), against 11% at 0–1 years", options: { color: "C3D0E2" } },
    {
      text: "\nCharge type splits 49% prepaid / 44% postpaid / 6% FBB, and 33,341 of the 70.1K — 47.6% — booked on 2 Aug alone.",
      options: { color: "C3D0E2" },
    },
  ],
  {
    x: 1.0,
    y: 5.92,
    w: 11.3,
    h: 0.76,
    margin: 0,
    valign: "top",
    lineSpacing: 16,
    fontFace: B,
    fontSize: 11,
  }
);

s.addText(
  "Source: Lending_2026.08.07.pptx (Book Loan). DL and PL are not expanded anywhere in that deck. Cash-need and risk percentages are as printed; all others are computed from its own counts and " +
    "reconcile to 94,371 and 70,100. Consortium bases overlap, so they sum past 100%. AIS-base figures cover the 70.1K AIS base only, not the whole book.",
  {
    x: 0.6,
    y: 6.88,
    w: 12.1,
    h: 0.42,
    margin: 0,
    fontFace: B,
    fontSize: 9,
    italic: true,
    color: MUTED,
  }
);

s.addNotes(
  "Every number is from Lending_2026.08.07.pptx. Reconciliation checks: DL 85,317 + PL 9,054 = 94,371, matching the stated total booked; " +
    "the cash-need and risk rows each sum to their own grand totals; and all four AIS breakdowns (charge type, nano loan, true, tenure) " +
    "sum to 70,100, matching the stated AIS consortium count. The deck's text says '47K are in TRUE ecosystem'; the true table gives 47,840. " +
    "Consortium bases overlap — KTB 88%, AIS 74% and OR 57% sum well past 100% because customers appear in more than one."
);

/* ================================================================= SLIDE 2
   AIS base information, plotted the same way as cash need and risk       */

const s2 = pres.addSlide();
s2.background = { color: PAPER };

s2.addText("AIS BASE INFORMATION  ·  70.1K  ·  DL 63,760 / PL 6,340", {
  x: 0.6,
  y: 0.36,
  w: 12.1,
  h: 0.26,
  margin: 0,
  fontFace: B,
  fontSize: 11.5,
  bold: true,
  charSpacing: 2.2,
  color: TEAL,
});
s2.addText("Long-tenured, prepaid-leaning, mostly TRUE", {
  x: 0.6,
  y: 0.66,
  w: 12.1,
  h: 0.62,
  margin: 0,
  valign: "top",
  fontFace: H,
  fontSize: 32,
  bold: true,
  color: INK,
});
s2.addText(
  "DL and PL profile almost identically on tenure and TRUE; they part company on charge type and the nano-loan flag.",
  {
    x: 0.6,
    y: 1.2,
    w: 12.1,
    h: 0.28,
    margin: 0,
    valign: "top",
    fontFace: B,
    fontSize: 12.5,
    color: MUTED,
  }
);

const aisCats = ["DL", "PL"];
const panels = [
  {
    title: "Charge type (% of customers)",
    colors: [R2, R3, R4],
    series: [
      { name: "FBB", labels: aisCats, values: [6.5, 6.2] },
      { name: "Prepaid", labels: aisCats, values: [48.8, 56.0] },
      { name: "Postpaid", labels: aisCats, values: [44.7, 37.8] },
    ],
  },
  {
    title: "Tenure in years (% of customers)",
    colors: [R2, R3, R4],
    series: [
      { name: "0–1", labels: aisCats, values: [11.0, 10.8] },
      { name: "2–4", labels: aisCats, values: [22.7, 23.6] },
      { name: "5 or more", labels: aisCats, values: [66.3, 65.6] },
    ],
  },
  {
    title: "Nano-loan flag (% of customers)",
    colors: [R1, R4],
    series: [
      { name: "N", labels: aisCats, values: [38.2, 46.2] },
      { name: "Y", labels: aisCats, values: [61.8, 53.8] },
    ],
  },
  {
    title: "TRUE ecosystem (% of customers)",
    colors: [R1, R4],
    series: [
      { name: "N", labels: aisCats, values: [31.5, 34.7] },
      { name: "Y", labels: aisCats, values: [68.5, 65.3] },
    ],
  },
];

panels.forEach((p, i) => {
  const col = i % 2;
  const row = Math.floor(i / 2);
  const px = 0.6 + col * 6.15;
  const py = 1.6 + row * 2.6;
  s2.addShape("roundRect", {
    x: px,
    y: py,
    w: 5.95,
    h: 2.44,
    rectRadius: 0.09,
    fill: { color: TINT },
    line: { type: "none" },
  });
  s2.addChart(pres.charts.BAR, p.series, {
    x: px + 0.12,
    y: py + 0.1,
    w: 5.71,
    h: 2.24,
    barDir: "bar",
    barGrouping: "percentStacked",
    barGapWidthPct: 45,
    chartColors: p.colors,
    showTitle: true,
    title: p.title,
    titleFontFace: B,
    titleFontSize: 12,
    titleColor: INK,
    showValue: true,
    dataLabelPosition: "ctr",
    dataLabelFontFace: B,
    dataLabelFontSize: 9.5,
    dataLabelColor: PAPER,
    dataLabelFormatCode: '0.0"%"',
    showLegend: true,
    legendPos: "b",
    legendFontFace: B,
    legendFontSize: 10,
    legendColor: INK,
    catAxisLabelFontFace: B,
    catAxisLabelFontSize: 11,
    catAxisLabelColor: INK,
    valAxisHidden: true,
    valGridLine: { style: "none" },
    catGridLine: { style: "none" },
  });
});

s2.addText(
  "Source: Lending_2026.08.07.pptx, the four AIS Base Information slides (charge type, nano loan, true, tenure), aggregated across all book dates. " +
    "Each breakdown sums to 70,100, matching the stated AIS consortium count. Percentages are computed from the deck's own counts and may differ by 1pt from its per-date rounding.",
  {
    x: 0.6,
    y: 6.9,
    w: 12.1,
    h: 0.42,
    margin: 0,
    fontFace: B,
    fontSize: 9,
    italic: true,
    color: MUTED,
  }
);

s2.addNotes(
  "Charts use the same percent-stacked form as the cash-need and risk chart on slide 1. " +
    "Biggest DL/PL gaps: nano-loan flag Y is 61.8% for DL against 53.8% for PL, and PL skews further to prepaid (56.0% vs 48.8%) " +
    "with correspondingly less postpaid (37.8% vs 44.7%). Tenure and TRUE are within ~3 points across the two loan types. " +
    "Counts behind the percentages - DL 63,760 and PL 6,340, totalling 70,100."
);

/* ================================================================= SLIDE 3
   Same grid as slide 2, measured by credit line and outstanding          */

const s2b = pres.addSlide();
s2b.background = { color: PAPER };
s2b.addText("AIS BASE INFORMATION  ·  SHARE OF CREDIT LINE AND OUTSTANDING", {
  x: 0.6, y: 0.36, w: 12.1, h: 0.26, margin: 0,
  fontFace: B, fontSize: 11.5, bold: true, charSpacing: 2.2, color: TEAL,
});
s2b.addText("Postpaid punches far above its headcount", {
  x: 0.6, y: 0.66, w: 12.1, h: 0.62, margin: 0, valign: "top",
  fontFace: H, fontSize: 32, bold: true, color: INK,
});
s2b.addText(
  "Prepaid is 48.8% of DL customers but 30.9% of DL credit line; postpaid runs 44.7% to 61.8%.",
  { x: 0.6, y: 1.2, w: 12.1, h: 0.28, margin: 0, valign: "top", fontFace: B, fontSize: 12.5, color: MUTED }
);

const valueCats = ["DL · CL", "DL · OS", "PL · CL", "PL · OS"];
const valuePanels = [
  { file: "charge.tsv", dim: "charge_type", order: ["FBB", "Prepaid", "Postpaid"], title: "Charge type · share of value (%)", colors: [R2, R3, R4] },
  { file: "tenure.tsv", dim: "tenure", order: ["0-1", "2-4", ">=5"], title: "Tenure in years · share of value (%)", colors: [R2, R3, R4], rename: { "0-1": "0–1", "2-4": "2–4", ">=5": "5 or more" } },
  { file: "nano.tsv", dim: "nano_f", order: ["N", "Y"], title: "Nano-loan flag · share of value (%)", colors: [R1, R4] },
  { file: "true.tsv", dim: "true_f", order: ["N", "Y"], title: "TRUE ecosystem · share of value (%)", colors: [R1, R4] },
];

valuePanels.forEach((p, i) => {
  const ds = tsv(p.file);
  // one bar per loan type x measure; each bar is the mix across the dimension
  const share = (lt, meas, k) => {
    const tot = ds.filter((r) => r.loan_type === lt).reduce((a, r) => a + num(r[meas]), 0);
    const part = ds.filter((r) => r.loan_type === lt && r[p.dim] === k).reduce((a, r) => a + num(r[meas]), 0);
    return Number(((part / tot) * 100).toFixed(1));
  };
  const series = p.order.map((k) => ({
    name: (p.rename && p.rename[k]) || k,
    labels: valueCats,
    values: [share("DL", "cl", k), share("DL", "os", k), share("PL", "cl", k), share("PL", "os", k)],
  }));

  const col = i % 2, row = Math.floor(i / 2);
  const px = 0.6 + col * 6.15, py = 1.6 + row * 2.6;
  s2b.addShape("roundRect", {
    x: px, y: py, w: 5.95, h: 2.44, rectRadius: 0.09,
    fill: { color: TINT }, line: { type: "none" },
  });
  s2b.addChart(pres.charts.BAR, series, {
    x: px + 0.12, y: py + 0.1, w: 5.71, h: 2.24,
    barDir: "bar",
    barGrouping: "percentStacked",
    barGapWidthPct: 40,
    chartColors: p.colors,
    showTitle: true,
    title: p.title,
    titleFontFace: B, titleFontSize: 12, titleColor: INK,
    showValue: true,
    dataLabelPosition: "ctr",
    dataLabelFontFace: B, dataLabelFontSize: 8.5, dataLabelColor: PAPER,
    dataLabelFormatCode: '0.0"%"',
    showLegend: true, legendPos: "b",
    legendFontFace: B, legendFontSize: 10, legendColor: INK,
    catAxisLabelFontFace: B, catAxisLabelFontSize: 10, catAxisLabelColor: INK,
    valAxisHidden: true,
    valGridLine: { style: "none" },
    catGridLine: { style: "none" },
  });
});

s2b.addText(
  "Source: figures as supplied, AIS base, aggregated across all book dates. CL = credit line, OS = outstanding. " +
    "Each bar is a mix within that loan type and measure, so all four bars in a panel read to 100%. Compare against slide 2, which shows the same four splits by customer count.",
  { x: 0.6, y: 6.9, w: 12.1, h: 0.44, margin: 0, fontFace: B, fontSize: 9, italic: true, color: MUTED }
);
s2b.addNotes(
  "Same four dimensions as slide 2, measured by value instead of headcount. " +
    "Charge type is the only mix that moves materially: DL prepaid falls from 48.8% of customers to 30.9% of credit line while postpaid rises 44.7% to 61.8% - postpaid customers carry roughly twice the line each. " +
    "The other three shift only a few points toward the favourable category: tenure 5+ years 66.3% to 71.3% of DL line, nano Y 61.8% to 66.6%, TRUE Y 68.5% to 73.3%. " +
    "PL is flatter throughout, and its OS mix leans more to prepaid (52.2%) than its CL mix (46.0%)."
);

/* ================================================================= SLIDE 4
   Book date - volume, and the mixes that actually move across cohorts    */

const s3 = pres.addSlide();
s3.background = { color: PAPER };

s3.addText("AIS BASE INFORMATION  ·  BY BOOK DATE", {
  x: 0.6,
  y: 0.36,
  w: 12.1,
  h: 0.26,
  margin: 0,
  fontFace: B,
  fontSize: 11.5,
  bold: true,
  charSpacing: 2.2,
  color: TEAL,
});
s3.addText("Nearly half the base booked on 2 August", {
  x: 0.6,
  y: 0.66,
  w: 12.1,
  h: 0.62,
  margin: 0,
  valign: "top",
  fontFace: H,
  fontSize: 32,
  bold: true,
  color: INK,
});
s3.addText(
  "PL arrives almost entirely on 5 Aug, and each later cohort is a little less nano-loan — the flag falls 9.5 points from 1 to 5 Aug.",
  {
    x: 0.6,
    y: 1.2,
    w: 12.1,
    h: 0.28,
    margin: 0,
    valign: "top",
    fontFace: B,
    fontSize: 12.5,
    color: MUTED,
  }
);

const bookDates = ["bef 1 Aug", "1-Aug", "2-Aug", "3-Aug", "4-Aug", "5-Aug"];

/* volume ---------------------------------------------------------------- */
s3.addShape("roundRect", {
  x: 0.6,
  y: 1.56,
  w: 12.1,
  h: 2.34,
  rectRadius: 0.09,
  fill: { color: TINT },
  line: { type: "none" },
});
s3.addChart(
  pres.charts.BAR,
  [
    { name: "DL", labels: bookDates, values: [641, 12182, 32182, 2060, 7173, 9522] },
    { name: "PL", labels: bookDates, values: [67, 398, 1159, 90, 371, 4255] },
  ],
  {
    x: 0.72,
    y: 1.66,
    w: 11.86,
    h: 2.14,
    barDir: "col",
    barGrouping: "clustered",
    barGapWidthPct: 55,
    chartColors: [R4, R2],
    showTitle: true,
    title: "Bookings by book date (customers)",
    titleFontFace: B,
    titleFontSize: 12,
    titleColor: INK,
    showValue: true,
    dataLabelPosition: "outEnd",
    dataLabelFontFace: B,
    dataLabelFontSize: 9.5,
    dataLabelColor: INK,
    dataLabelFormatCode: "#,##0",
    showLegend: true,
    legendPos: "b",
    legendFontFace: B,
    legendFontSize: 10,
    legendColor: INK,
    catAxisLabelFontFace: B,
    catAxisLabelFontSize: 11,
    catAxisLabelColor: INK,
    valAxisHidden: true,
    valAxisMinVal: 0,
    valAxisMaxVal: 36000,
    valGridLine: { style: "none" },
    catGridLine: { style: "none" },
  }
);

/* the two mixes that move ----------------------------------------------- */
const dateMixes = [
  {
    title: "Nano-loan flag by book date (% of customers)",
    series: [
      { name: "N", labels: bookDates, values: [37.1, 35.2, 38.0, 36.8, 39.6, 44.7] },
      { name: "Y", labels: bookDates, values: [62.9, 64.8, 62.0, 63.2, 60.4, 55.3] },
    ],
  },
  {
    title: "TRUE ecosystem by book date (% of customers)",
    series: [
      { name: "N", labels: bookDates, values: [18.1, 29.7, 31.9, 31.3, 32.7, 33.4] },
      { name: "Y", labels: bookDates, values: [81.9, 70.3, 68.1, 68.7, 67.3, 66.6] },
    ],
  },
];
dateMixes.forEach((p, i) => {
  const px = 0.6 + i * 6.15;
  s3.addShape("roundRect", {
    x: px,
    y: 4.06,
    w: 5.95,
    h: 2.46,
    rectRadius: 0.09,
    fill: { color: TINT },
    line: { type: "none" },
  });
  s3.addChart(pres.charts.BAR, p.series, {
    x: px + 0.12,
    y: 4.16,
    w: 5.71,
    h: 2.26,
    barDir: "col",
    barGrouping: "percentStacked",
    barGapWidthPct: 45,
    chartColors: [R1, R4],
    showTitle: true,
    title: p.title,
    titleFontFace: B,
    titleFontSize: 12,
    titleColor: INK,
    showValue: true,
    dataLabelPosition: "ctr",
    dataLabelFontFace: B,
    dataLabelFontSize: 9,
    dataLabelColor: PAPER,
    dataLabelFormatCode: '0"%"',
    showLegend: true,
    legendPos: "b",
    legendFontFace: B,
    legendFontSize: 10,
    legendColor: INK,
    catAxisLabelFontFace: B,
    catAxisLabelFontSize: 10,
    catAxisLabelColor: INK,
    valAxisHidden: true,
    valGridLine: { style: "none" },
    catGridLine: { style: "none" },
  });
});

s3.addText(
  "Source: Lending_2026.08.07.pptx, AIS Base Information slides. Volumes are the grand totals per book date; the two mixes combine DL and PL. " +
    "'bef 1 Aug' holds only 708 customers, so its mix is a small-sample outlier. " +
    "Charge type and tenure move far less across the same window and are charted on the next slide.",
  {
    x: 0.6,
    y: 6.66,
    w: 12.1,
    h: 0.56,
    margin: 0,
    fontFace: B,
    fontSize: 9,
    italic: true,
    color: MUTED,
  }
);

s3.addNotes(
  "2 Aug alone accounts for 33,341 of the 70,100 AIS bookings (47.6%). PL is back-loaded: 4,255 of its 6,340 bookings (67%) land on 5 Aug. " +
    "Nano-loan Y falls from 64.8% on 1 Aug to 55.3% on 5 Aug, a 9.5 point drop, and the PL surge on 5 Aug is part of why - PL runs a lower nano share than DL. " +
    "TRUE Y drifts down more gently, 70.3% to 66.6%. Charge type and tenure are effectively flat across the same window."
);

/* ================================================================= SLIDE 5
   Charge type and tenure by book date - the mixes that hold steady       */

const s4 = pres.addSlide();
s4.background = { color: PAPER };

s4.addText("AIS BASE INFORMATION  ·  BY BOOK DATE", {
  x: 0.6,
  y: 0.36,
  w: 12.1,
  h: 0.26,
  margin: 0,
  fontFace: B,
  fontSize: 11.5,
  bold: true,
  charSpacing: 2.2,
  color: TEAL,
});
s4.addText("Charge type and tenure hold steady", {
  x: 0.6,
  y: 0.66,
  w: 12.1,
  h: 0.62,
  margin: 0,
  valign: "top",
  fontFace: H,
  fontSize: 32,
  bold: true,
  color: INK,
});
s4.addText(
  "Unlike the nano-loan flag, these two barely move from 1 Aug onward — only the small pre-August cohort looks different.",
  {
    x: 0.6,
    y: 1.2,
    w: 12.1,
    h: 0.28,
    margin: 0,
    valign: "top",
    fontFace: B,
    fontSize: 12.5,
    color: MUTED,
  }
);

const bookDates4 = ["bef 1 Aug", "1-Aug", "2-Aug", "3-Aug", "4-Aug", "5-Aug"];
const steady = [
  {
    title: "Charge type by book date (% of customers)",
    series: [
      { name: "FBB", labels: bookDates4, values: [3.8, 6.9, 6.9, 7.6, 6.9, 4.9] },
      { name: "Prepaid", labels: bookDates4, values: [17.5, 46.4, 49.4, 47.9, 47.9, 55.0] },
      { name: "Postpaid", labels: bookDates4, values: [78.7, 46.7, 43.8, 44.5, 45.1, 40.1] },
    ],
  },
  {
    title: "Tenure in years by book date (% of customers)",
    series: [
      { name: "0–1", labels: bookDates4, values: [4.4, 11.6, 11.6, 10.9, 10.5, 9.7] },
      { name: "2–4", labels: bookDates4, values: [14.7, 22.7, 23.3, 23.1, 22.7, 21.9] },
      { name: "5 or more", labels: bookDates4, values: [80.9, 65.7, 65.1, 66.0, 66.8, 68.4] },
    ],
  },
];
steady.forEach((p, i) => {
  const px = 0.6 + i * 6.15;
  s4.addShape("roundRect", {
    x: px,
    y: 1.58,
    w: 5.95,
    h: 3.42,
    rectRadius: 0.09,
    fill: { color: TINT },
    line: { type: "none" },
  });
  s4.addChart(pres.charts.BAR, p.series, {
    x: px + 0.12,
    y: 1.68,
    w: 5.71,
    h: 3.22,
    barDir: "col",
    barGrouping: "percentStacked",
    barGapWidthPct: 45,
    chartColors: [R2, R3, R4],
    showTitle: true,
    title: p.title,
    titleFontFace: B,
    titleFontSize: 12,
    titleColor: INK,
    showValue: true,
    dataLabelPosition: "ctr",
    dataLabelFontFace: B,
    dataLabelFontSize: 9,
    dataLabelColor: PAPER,
    dataLabelFormatCode: '0"%"',
    showLegend: true,
    legendPos: "b",
    legendFontFace: B,
    legendFontSize: 10,
    legendColor: INK,
    catAxisLabelFontFace: B,
    catAxisLabelFontSize: 10,
    catAxisLabelColor: INK,
    valAxisHidden: true,
    valGridLine: { style: "none" },
    catGridLine: { style: "none" },
  });
});

s4.addShape("roundRect", {
  x: 0.6,
  y: 5.2,
  w: 12.1,
  h: 1.3,
  rectRadius: 0.09,
  fill: { color: NAVY },
  line: { type: "none" },
});
s4.addText("HOW MUCH EACH MIX MOVES, 1 AUG TO 5 AUG", {
  x: 1.0,
  y: 5.38,
  w: 11.3,
  h: 0.26,
  margin: 0,
  fontFace: B,
  fontSize: 10.5,
  bold: true,
  charSpacing: 1.8,
  color: MINT,
});
s4.addText(
  [
    { text: "Prepaid 46–55% (8.6pt) · Tenure 5+ years 65–68% (3.3pt)", options: { color: MINT, bold: true } },
    {
      text: " — against the nano-loan flag's 9.5pt slide on the previous slide. Charge type drifts toward prepaid on 5 Aug, tracking the PL surge; tenure is flat throughout.",
      options: { color: "C3D0E2" },
    },
    {
      text: "\nThe pre-1-Aug cohort is the exception — 79% postpaid, 81% long-tenured — but holds just 708 of the 70,100 customers.",
      options: { color: "C3D0E2" },
    },
  ],
  {
    x: 1.0,
    y: 5.66,
    w: 11.3,
    h: 0.76,
    margin: 0,
    valign: "top",
    lineSpacing: 16,
    fontFace: B,
    fontSize: 11,
  }
);

s4.addText(
  "Source: Lending_2026.08.07.pptx, the charge type and tenure AIS Base Information slides, DL and PL combined and aggregated to each book date. " +
    "Percentages are computed from the deck's own counts; each date column sums to that date's grand total, and the six together sum to 70,100.",
  {
    x: 0.6,
    y: 6.66,
    w: 12.1,
    h: 0.56,
    margin: 0,
    fontFace: B,
    fontSize: 9,
    italic: true,
    color: MUTED,
  }
);

s4.addNotes(
  "Completes the by-book-date view: charge type and tenure alongside the nano-loan flag and TRUE on slide 3. " +
    "From 1 Aug onward prepaid runs 46.4, 49.4, 47.9, 47.9, 55.0 and tenure 5+ years runs 65.7, 65.1, 66.0, 66.8, 68.4 - " +
    "spreads of 8.6 and 3.3 points against 9.5 for the nano-loan flag. The 5 Aug prepaid uptick coincides with the PL surge, " +
    "and PL runs a higher prepaid share than DL (56.0% vs 48.8%)."
);

/* ================================================================= SLIDE 6
   The four book-date mixes, each measured on credit line AND outstanding */

{
  const sv = pres.addSlide();
  sv.background = { color: PAPER };

  // horizontal bars plot the first label at the bottom, so build bottom-up:
  // newest first, and OS above CL within each date
  const PAIRS = [];
  ["5-Aug", "4-Aug", "3-Aug", "2-Aug", "1-Aug", "bef 1 Aug"].forEach((d) => {
    PAIRS.push({ date: d, meas: "os" }, { date: d, meas: "cl" });
  });
  const catLabels = PAIRS.map((p) => `${p.date} ${p.meas.toUpperCase()}`);

  sv.addText("AIS BASE INFORMATION  ·  BY BOOK DATE  ·  CREDIT LINE vs OUTSTANDING", {
    x: 0.6, y: 0.36, w: 12.1, h: 0.26, margin: 0,
    fontFace: B, fontSize: 11.5, bold: true, charSpacing: 2.2, color: TEAL,
  });
  sv.addText("Weaker segments draw more of their line", {
    x: 0.6, y: 0.66, w: 12.1, h: 0.62, margin: 0, valign: "top",
    fontFace: H, fontSize: 30, bold: true, color: INK,
  });
  sv.addText(
    "Prepaid, short tenure, no nano flag and non-TRUE each take a larger share of OS than of CL — they use more of what they are granted.",
    { x: 0.6, y: 1.18, w: 12.1, h: 0.28, margin: 0, valign: "top", fontFace: B, fontSize: 12, color: MUTED }
  );

  const mixes = [
    { file: "charge.tsv", dim: "charge_type", order: ["FBB", "Prepaid", "Postpaid"], title: "Charge type", colors: [R2, R3, R4] },
    { file: "tenure.tsv", dim: "tenure", order: ["0-1", "2-4", ">=5"], title: "Tenure in years", colors: [R2, R3, R4], rename: { "0-1": "0–1", "2-4": "2–4", ">=5": "5+" } },
    { file: "nano.tsv", dim: "nano_f", order: ["N", "Y"], title: "Nano-loan flag", colors: [R1, R4] },
    { file: "true.tsv", dim: "true_f", order: ["N", "Y"], title: "TRUE ecosystem", colors: [R1, R4] },
  ];

  mixes.forEach((p, i) => {
    const ds = tsv(p.file);
    const share = (date, meas, k) => {
      const rows = ds.filter((r) => r.book_date === date);
      const tot = rows.reduce((a, r) => a + num(r[meas]), 0);
      const part = rows.filter((r) => r[p.dim] === k).reduce((a, r) => a + num(r[meas]), 0);
      return Number(((part / tot) * 100).toFixed(1));
    };
    const px = 0.6 + i * 3.075;
    sv.addShape("roundRect", {
      x: px, y: 1.56, w: 2.875, h: 4.94, rectRadius: 0.09,
      fill: { color: TINT }, line: { type: "none" },
    });
    sv.addChart(
      pres.charts.BAR,
      p.order.map((k) => ({
        name: (p.rename && p.rename[k]) || k,
        labels: catLabels,
        values: PAIRS.map((c) => share(c.date, c.meas, k)),
      })),
      {
        x: px + 0.09, y: 1.64, w: 2.695, h: 4.78,
        barDir: "bar",
        barGrouping: "percentStacked",
        barGapWidthPct: 35,
        chartColors: p.colors,
        showTitle: true,
        title: `${p.title} (CL / OS)`,
        titleFontFace: B, titleFontSize: 11, titleColor: INK,
        showValue: true,
        dataLabelPosition: "ctr",
        dataLabelFontFace: B, dataLabelFontSize: 7.5, dataLabelColor: PAPER,
        dataLabelFormatCode: '0"%"',
        showLegend: true, legendPos: "b",
        legendFontFace: B, legendFontSize: 9, legendColor: INK,
        catAxisLabelFontFace: B, catAxisLabelFontSize: 7.5, catAxisLabelColor: INK,
        valAxisHidden: true,
        valGridLine: { style: "none" },
        catGridLine: { style: "none" },
      }
    );
  });

  sv.addText(
    "Source: figures as supplied, AIS base. CL = credit line, OS = outstanding; each bar is a mix within that date and measure, so every bar reads to 100%. " +
      "The OS-above-CL pattern holds in all five August cohorts on charge type and tenure, and in four of five on the nano and TRUE flags — the exceptions are the 708-customer pre-August cohort " +
      "and a 0.2-point wobble on TRUE on 3 Aug. Slides 4 and 5 show the same four cuts by customer count.",
    { x: 0.6, y: 6.62, w: 12.1, h: 0.5, margin: 0, fontFace: B, fontSize: 8.5, italic: true, color: MUTED }
  );
  sv.addNotes(
    "Pairing CL and OS on the same axis makes the utilisation story readable per segment. " +
      "In every August cohort the prepaid share of outstanding exceeds its share of credit line - by 1.6pt on 1 Aug widening to 6.0pt on 5 Aug - " +
      "and the same direction holds for 0-1 year tenure, no-nano and non-TRUE. " +
      "Read plainly: the segments granted smaller lines draw a higher fraction of them, so outstanding is tilted toward the weaker end of every cut."
  );
}

/* ================================================================= SLIDE 7
   Slide 1's cash-need / risk mix chart, measured on credit line and OS  */

{
  const sv = pres.addSlide();
  sv.background = { color: PAPER };
  const cr = tsv("cash_risk.tsv");
  const LEVELS = ["Null", "Low", "Mid", "High"];
  // bottom-up, so DL - CL ends up on top
  const cats = [
    { lt: "PL", m: "os" }, { lt: "PL", m: "cl" },
    { lt: "DL", m: "os" }, { lt: "DL", m: "cl" },
  ];
  const catLabels = cats.map((c) => `${c.lt} · ${c.m.toUpperCase()}`);
  const share = (field, lvl, lt, m) => {
    const rows = cr.filter((r) => r.loan_type === lt);
    const tot = rows.reduce((a, r) => a + num(r[m]), 0);
    const part = rows.filter((r) => r[field] === lvl).reduce((a, r) => a + num(r[m]), 0);
    return Number(((part / tot) * 100).toFixed(1));
  };

  sv.addText("CASH NEED AND RISK LEVEL  ·  CREDIT LINE vs OUTSTANDING", {
    x: 0.6, y: 0.36, w: 12.1, h: 0.26, margin: 0,
    fontFace: B, fontSize: 11.5, bold: true, charSpacing: 2.2, color: TEAL,
  });
  sv.addText("Riskier customers draw a bigger share", {
    x: 0.6, y: 0.66, w: 12.1, h: 0.62, margin: 0, valign: "top",
    fontFace: H, fontSize: 30, bold: true, color: INK,
  });
  sv.addText(
    `High risk is ${share("risk_level", "High", "DL", "cl")}% of DL credit line but ${share("risk_level", "High", "DL", "os")}% of DL outstanding, ` +
      `while low risk falls from ${share("risk_level", "Low", "DL", "cl")}% to ${share("risk_level", "Low", "DL", "os")}%.`,
    { x: 0.6, y: 1.18, w: 12.1, h: 0.28, margin: 0, valign: "top", fontFace: B, fontSize: 12, color: MUTED }
  );

  [
    // labelMin hides labels too narrow to render inside their segment
    { field: "cash_need", title: "Cash need · % of CL / OS", labelMin: 3 },
    { field: "risk_level", title: "Risk level · % of CL / OS", labelMin: 6 },
  ].forEach((p, i) => {
    const px = 0.6 + i * 6.15;
    sv.addShape("roundRect", {
      x: px, y: 1.58, w: 5.95, h: 4.7, rectRadius: 0.09,
      fill: { color: TINT }, line: { type: "none" },
    });
    sv.addChart(
      pres.charts.BAR,
      LEVELS.map((lvl) => ({
        name: lvl,
        labels: catLabels,
        values: cats.map((c) => share(p.field, lvl, c.lt, c.m)),
      })),
      {
        x: px + 0.12, y: 1.68, w: 5.71, h: 4.5,
        barDir: "bar",
        barGrouping: "percentStacked",
        barGapWidthPct: 90,
        chartColors: [R1, R2, R3, R4],
        showTitle: true,
        title: p.title,
        titleFontFace: B, titleFontSize: 12, titleColor: INK,
        showValue: true,
        dataLabelPosition: "ctr",
        dataLabelFontFace: B, dataLabelFontSize: 9.5, dataLabelColor: PAPER,
        dataLabelFormatCode: `[<${p.labelMin}]"";0.0"%"`,
        showLegend: true, legendPos: "b",
        legendFontFace: B, legendFontSize: 10, legendColor: INK,
        catAxisLabelFontFace: B, catAxisLabelFontSize: 11, catAxisLabelColor: INK,
        valAxisHidden: true,
        valGridLine: { style: "none" },
        catGridLine: { style: "none" },
      }
    );
  });

  sv.addText(
    "Source: figures as supplied. CL = credit line, OS = outstanding; each bar is a mix within that loan type and measure, so every bar reads to 100%. " +
      "Slide 1 shows the same two cuts by customer count, where cash need and risk are the only figures the source states as percentages.",
    { x: 0.6, y: 6.44, w: 12.1, h: 0.44, margin: 0, fontFace: B, fontSize: 9, italic: true, color: MUTED }
  );
  sv.addNotes(
    "Same chart form as slide 1, on value instead of headcount, with CL and OS paired. " +
      "Risk: High rises from 36.3% of DL credit line to 39.1% of outstanding, and 31.7% to 36.9% on PL; Low falls 35.8% to 33.1% and 39.9% to 34.3%. " +
      "Cash need moves less - High 67.8% to 68.6% on DL, 52.2% to 55.7% on PL. " +
      "This is the same direction as the book-date cuts on the previous slide: the weaker segment of every dimension draws a higher fraction of its line."
  );
}

/* ============================================================ DATA TABLES
   Slides 8-13: the source tables, reproduced from data/*.tsv              */

/* generic table slide ---------------------------------------------------- */
function tableSlide({ kicker, title, subtitle, headers, colW, body, total, footnote, notes, numericFrom }) {
  const sl = pres.addSlide();
  sl.background = { color: PAPER };
  sl.addText(kicker.toUpperCase(), {
    x: 0.6, y: 0.36, w: 12.1, h: 0.26, margin: 0,
    fontFace: B, fontSize: 11.5, bold: true, charSpacing: 2.2, color: TEAL,
  });
  sl.addText(title, {
    x: 0.6, y: 0.66, w: 12.1, h: 0.62, margin: 0, valign: "top",
    fontFace: H, fontSize: 30, bold: true, color: INK,
  });
  sl.addText(subtitle, {
    x: 0.6, y: 1.2, w: 12.1, h: 0.28, margin: 0, valign: "top",
    fontFace: B, fontSize: 12, color: MUTED,
  });

  const y = 1.58;
  const nRows = body.length + (total ? 1 : 0) + 1;
  const rowH = Math.min(0.42, (6.76 - y) / nRows);
  const fs2 = rowH >= 0.36 ? 11 : rowH >= 0.30 ? 10 : 9;

  const cell = (t, j, opts = {}) => ({
    text: t,
    options: Object.assign(
      {
        fontFace: j >= numericFrom ? H : B,
        fontSize: fs2,
        align: j >= numericFrom ? "right" : "left",
        color: INK,
      },
      opts
    ),
  });
  const rows = [
    headers.map((t, j) => ({
      text: t,
      options: {
        bold: true, color: PAPER, fill: { color: NAVY }, fontFace: B,
        fontSize: fs2, align: j >= numericFrom ? "right" : "left",
      },
    })),
    ...body.map((r, i) =>
      r.map((t, j) => cell(t, j, { fill: { color: i % 2 ? TINT : PAPER }, bold: j === 0 }))
    ),
  ];
  if (total) {
    rows.push(total.map((t, j) => cell(t, j, { fill: { color: TINT_COOL }, bold: true })));
  }
  sl.addTable(rows, {
    x: 0.6, y, w: 12.1, colW, rowH,
    border: { type: "solid", color: RULE, pt: 0.5 },
    valign: "middle",
    margin: [0.02, 0.07, 0.02, 0.07],
  });
  sl.addText(footnote, {
    x: 0.6, y: 6.88, w: 12.1, h: 0.44, margin: 0,
    fontFace: B, fontSize: 9, italic: true, color: MUTED,
  });
  if (notes) sl.addNotes(notes);
  return sl;
}

const SRC = "Source: figures as supplied. cl = credit line, os = outstanding. ";
const DATES = ["bef 1 Aug", "1-Aug", "2-Aug", "3-Aug", "4-Aug", "5-Aug"];

/* --- slide 8: customer base -------------------------------------------- */
{
  const cb = tsv("cust_base.tsv").map((r) => ({
    base: r.cust_base, cl: num(r.cl), os: num(r.os), cnt: num(r.cust_cnt),
  }));
  const membership = {
    Null: "None", A1: "AIS only", A1_o: "AIS + OR", J: "AIS + KTB",
    J_O: "AIS + KTB + OR", K1: "KTB only", K1_O: "KTB + OR", O1: "OR only",
  };
  const T = cb.reduce((a, r) => ({ cl: a.cl + r.cl, os: a.os + r.os, cnt: a.cnt + r.cnt }), { cl: 0, os: 0, cnt: 0 });
  const util = cb.map((r) => (r.os / r.cl) * 100);
  tableSlide({
    kicker: "Customer base · credit line and outstanding",
    title: "Two-thirds of the book is dual AIS–KTB",
    subtitle: `${M(T.cl)} of credit line against ${M(T.os)} outstanding — ${pct(T.os, T.cl)} utilised, and strikingly even across every segment (${Math.min(...util).toFixed(1)}–${Math.max(...util).toFixed(1)}%).`,
    headers: ["cust_base", "Consortium membership", "Credit line", "Outstanding", "Util.", "Customers", "% of book", "CL / customer"],
    colW: [1.1, 2.05, 1.9, 1.9, 1.0, 1.35, 1.0, 1.8],
    numericFrom: 2,
    body: cb.map((r) => [
      r.base, membership[r.base], fmt(r.cl), fmt(r.os), pct(r.os, r.cl),
      fmt(r.cnt), pct(r.cnt, T.cnt), fmt(Math.round(r.cl / r.cnt)),
    ]),
    total: ["Total", "", fmt(T.cl), fmt(T.os), pct(T.os, T.cl), fmt(T.cnt), "100.0%", fmt(Math.round(T.cl / T.cnt))],
    footnote:
      SRC +
      "Consortium membership is inferred from the segment codes, not stated in the source: the groupings reproduce the deck's AIS 70,100, KTB 83,362, OR 53,712 and 'not in any' 794 exactly. " +
      "Customers total 94,371, matching the stated total booked.",
    notes:
      "The membership column is an inference that checks out four ways: A1+A1_o+J+J_O = 70,100 (AIS), K1+K1_O+J+J_O = 83,362 (KTB), " +
      "O1+A1_o+K1_O+J_O = 53,712 (OR), and Null = 794. J and J_O together hold 61,771 customers (65.5%) and 696.5M of credit line (63.8%).",
  });
}

/* --- slide 9: cash need and risk level, as two separate tables --------- */
{
  const cr = tsv("cash_risk.tsv").map((r) => ({
    cash: r.cash_need, risk: r.risk_level, lt: r.loan_type, cl: num(r.cl), os: num(r.os),
  }));
  const LEVELS = ["Null", "Low", "Mid", "High"];

  // roll the 4x4 grid up along one axis at a time
  const rollup = (field) =>
    LEVELS.map((lvl) => {
      const pick = (lt) =>
        cr.filter((r) => r[field] === lvl && r.lt === lt)
          .reduce((a, r) => ({ cl: a.cl + r.cl, os: a.os + r.os }), { cl: 0, os: 0 });
      const d = pick("DL"), pl = pick("PL");
      return { lvl, d, pl, cl: d.cl + pl.cl, os: d.os + pl.os };
    });
  const byCash = rollup("cash");
  const byRisk = rollup("risk");
  const GT = byCash.reduce((a, r) => ({ cl: a.cl + r.cl, os: a.os + r.os }), { cl: 0, os: 0 });

  const sl = pres.addSlide();
  sl.background = { color: PAPER };
  sl.addText("CASH NEED AND RISK LEVEL  ·  CREDIT LINE AND OUTSTANDING", {
    x: 0.6, y: 0.36, w: 12.1, h: 0.26, margin: 0,
    fontFace: B, fontSize: 11.5, bold: true, charSpacing: 2.2, color: TEAL,
  });
  const hiCash = byCash.find((r) => r.lvl === "High");
  const lowRisk = byRisk.find((r) => r.lvl === "Low");
  const hiRisk = byRisk.find((r) => r.lvl === "High");
  const midRisk = byRisk.find((r) => r.lvl === "Mid");
  sl.addText("Cash need concentrates, risk does not", {
    x: 0.6, y: 0.66, w: 12.1, h: 0.62, margin: 0, valign: "top",
    fontFace: H, fontSize: 30, bold: true, color: INK,
  });
  sl.addText(
    `High cash need holds ${pct(hiCash.cl, GT.cl)} of the credit line; risk splits ` +
      `${pct(lowRisk.cl, GT.cl)} low / ${pct(midRisk.cl, GT.cl)} mid / ${pct(hiRisk.cl, GT.cl)} high.`,
    { x: 0.6, y: 1.18, w: 12.1, h: 0.28, margin: 0, valign: "top", fontFace: B, fontSize: 12, color: MUTED }
  );

  const headers = ["Level", "DL credit line", "DL outstanding", "DL util.", "PL credit line", "PL outstanding", "PL util.", "Total credit line", "% of line"];
  const colW = [1.3, 1.65, 1.65, 0.9, 1.65, 1.65, 0.9, 1.45, 0.95];

  function block(label, data, y) {
    sl.addText(label, {
      x: 0.6, y, w: 12.1, h: 0.24, margin: 0,
      fontFace: H, fontSize: 14, bold: true, color: INK,
    });
    const T = data.reduce(
      (a, r) => ({ dcl: a.dcl + r.d.cl, dos: a.dos + r.d.os, pcl: a.pcl + r.pl.cl, pos: a.pos + r.pl.os }),
      { dcl: 0, dos: 0, pcl: 0, pos: 0 }
    );
    const cell = (t, j, o = {}) => ({
      text: t,
      options: Object.assign(
        { fontFace: j >= 1 ? H : B, fontSize: 10.5, align: j >= 1 ? "right" : "left", color: INK },
        o
      ),
    });
    const rows = [
      headers.map((t, j) => ({
        text: t,
        options: { bold: true, color: PAPER, fill: { color: NAVY }, fontFace: B, fontSize: 10.5, align: j >= 1 ? "right" : "left" },
      })),
      ...data.map((r, i) =>
        [r.lvl, fmt(r.d.cl), fmt(r.d.os), pct(r.d.os, r.d.cl), fmt(r.pl.cl), fmt(r.pl.os), pct(r.pl.os, r.pl.cl), fmt(r.cl), pct(r.cl, GT.cl)]
          .map((t, j) => cell(t, j, { fill: { color: i % 2 ? TINT : PAPER }, bold: j === 0 }))
      ),
      ["Total", fmt(T.dcl), fmt(T.dos), pct(T.dos, T.dcl), fmt(T.pcl), fmt(T.pos), pct(T.pos, T.pcl), fmt(T.dcl + T.pcl), "100.0%"]
        .map((t, j) => cell(t, j, { fill: { color: TINT_COOL }, bold: true })),
    ];
    sl.addTable(rows, {
      x: 0.6, y: y + 0.26, w: 12.1, colW, rowH: 0.37,
      border: { type: "solid", color: RULE, pt: 0.5 },
      valign: "middle", margin: [0.02, 0.07, 0.02, 0.07],
    });
  }
  block("By cash need", byCash, 1.54);
  block("By risk level", byRisk, 4.24);

  sl.addText(
    SRC +
      "Each table rolls the same 4×4 cash-need × risk-level grid up along one axis, so both total to 1,092,262,000 of credit line and 660,543,706 outstanding — matching the customer-base table. " +
      "The source has no customer counts at this cut, so utilisation and share of line are the only derived columns. Low risk holds 37.0% of the line against 29.0% of customers in the deck's risk counts — low-risk customers carry larger lines.",
    { x: 0.6, y: 6.9, w: 12.1, h: 0.44, margin: 0, fontFace: B, fontSize: 9, italic: true, color: MUTED }
  );
  sl.addNotes(
    "Split out of the combined 4x4 table so each dimension reads on its own. " +
      "Cash need is heavily concentrated - High alone is 63.3% of the credit line. Risk is not: Low 37.0%, Mid 25.5%, High 35.0%. " +
      "Low risk holding 37.0% of the line against 29.0% of customers means low-risk customers are granted larger lines on average. " +
      "The DL/PL utilisation gap persists on both cuts: DL 65.2% against PL 48.7%."
  );
}

/* --- slides 10-13: the four AIS breakdowns by book date ------------------ */
function aisTableSlide(file, dimField, dimOrder, dimLabel, kicker, titleFn, subtitleFn, notes) {
  const ds = tsv(file);
  const get = (d, k, lt) => {
    const r = ds.find((x) => x.book_date === d && x[dimField] === k && x.loan_type === lt);
    return r ? { c: num(r.cust_cnt), cl: num(r.cl), os: num(r.os) } : { c: 0, cl: 0, os: 0 };
  };
  const body = [];
  const T = { dc: 0, dcl: 0, dos: 0, pc: 0, pcl: 0, pos: 0 };
  const byDim = {};
  DATES.forEach((d) =>
    dimOrder.forEach((k) => {
      const a = get(d, k, "DL"), b = get(d, k, "PL");
      T.dc += a.c; T.dcl += a.cl; T.dos += a.os;
      T.pc += b.c; T.pcl += b.cl; T.pos += b.os;
      byDim[k] = byDim[k] || { c: 0, cl: 0 };
      byDim[k].c += a.c + b.c;
      byDim[k].cl += a.cl + b.cl;
      body.push([d, k, fmt(a.c), fmt(a.cl), fmt(a.os), fmt(b.c), fmt(b.cl), fmt(b.os)]);
    })
  );
  const totCl = T.dcl + T.pcl, totC = T.dc + T.pc;
  tableSlide({
    kicker,
    title: titleFn(byDim, totC, totCl),
    subtitle: subtitleFn(byDim, totC, totCl),
    headers: ["Book date", dimLabel, "DL customers", "DL credit line", "DL outstanding", "PL customers", "PL credit line", "PL outstanding"],
    colW: [1.25, 1.35, 1.35, 1.9, 1.9, 1.35, 1.5, 1.5],
    numericFrom: 2,
    body,
    total: ["Total", "", fmt(T.dc), fmt(T.dcl), fmt(T.dos), fmt(T.pc), fmt(T.pcl), fmt(T.pos)],
    footnote:
      SRC +
      "Customers total 70,100 with 801,481,000 of credit line and 476,272,677 outstanding — identical across all four AIS breakdowns, and 74.3% of the 94,371-customer book.",
    notes,
  });
}

aisTableSlide(
  "charge.tsv", "charge_type", ["FBB", "Prepaid", "Postpaid"], "Charge type",
  "Charge type by book date · AIS base",
  (d, tc, tcl) => `Prepaid leads on customers, postpaid on value`,
  (d, tc, tcl) =>
    `Prepaid is ${pct(d.Prepaid.c, tc)} of customers but only ${pct(d.Prepaid.cl, tcl)} of credit line; postpaid is ${pct(d.Postpaid.c, tc)} of customers and ${pct(d.Postpaid.cl, tcl)} of the line.`,
  "Postpaid customers carry a materially larger line each than prepaid customers, which is why the customer and value rankings invert."
);

aisTableSlide(
  "nano.tsv", "nano_f", ["N", "Y"], "Nano flag",
  "Nano-loan flag by book date · AIS base",
  (d, tc, tcl) => `Nano flag: ${pct(d.Y.c, tc)} of customers, ${pct(d.Y.cl, tcl)} of line`,
  (d, tc, tcl) => `Credit line tracks headcount almost exactly, so the flag does not by itself mark a larger-line population.`,
  "Nano flag Y: 42,806 customers and 503.8M of credit line. Broadly proportional, so the flag does not by itself mark a larger-line population."
);

aisTableSlide(
  "true.tsv", "true_f", ["N", "Y"], "TRUE flag",
  "TRUE ecosystem by book date · AIS base",
  (d, tc, tcl) => `TRUE customers hold ${pct(d.Y.cl, tcl)} of the line`,
  (d, tc, tcl) => `They are ${pct(d.Y.c, tc)} of customers, so TRUE membership skews toward the larger lines.`,
  "TRUE Y: 47,840 customers (68.2%) holding 573.3M of credit line (71.5%) - a few points more than their headcount share."
);

aisTableSlide(
  "tenure.tsv", "tenure", ["0-1", "2-4", ">=5"], "Tenure",
  "Tenure by book date · AIS base",
  (d, tc, tcl) => `5+ year tenure holds ${pct(d[">=5"].cl, tcl)} of the line`,
  (d, tc, tcl) =>
    `They are ${pct(d[">=5"].c, tc)} of customers, while the 0–1 year cohort is ${pct(d["0-1"].c, tc)} of customers and just ${pct(d["0-1"].cl, tcl)} of the line.`,
  "Long tenure earns a bigger line: 5+ years is 66.2% of customers but 71.0% of credit line, while 0-1 years is 11.0% of customers and only 9.0% of the line."
);

const out = path.join(__dirname, "Lending_Book_Loan_Summary.pptx");
pres.writeFile({ fileName: out }).then(() => console.log("wrote", out));
