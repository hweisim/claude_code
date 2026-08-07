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

/* ================================================================= SLIDE 4
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

/* ============================================================ DATA TABLES
   Slides 5-10: the source tables, reproduced from data/*.tsv              */

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

/* --- slide 5: customer base -------------------------------------------- */
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

/* --- slide 6: cash need x risk level ------------------------------------ */
{
  const cr = tsv("cash_risk.tsv");
  const order = ["Null", "Low", "Mid", "High"];
  const get = (c, r, lt) => {
    const row = cr.find((x) => x.cash_need === c && x.risk_level === r && x.loan_type === lt);
    return { cl: num(row.cl), os: num(row.os) };
  };
  const body = [];
  const T = { dcl: 0, dos: 0, pcl: 0, pos: 0 };
  order.forEach((c) =>
    order.forEach((r) => {
      const d = get(c, r, "DL"), p = get(c, r, "PL");
      T.dcl += d.cl; T.dos += d.os; T.pcl += p.cl; T.pos += p.os;
      body.push([c, r, fmt(d.cl), fmt(d.os), pct(d.os, d.cl), fmt(p.cl), fmt(p.os), pct(p.os, p.cl)]);
    })
  );
  const highCl = order.reduce((a, r) => a + get("High", r, "DL").cl + get("High", r, "PL").cl, 0);
  tableSlide({
    kicker: "Cash need × risk level · credit line and outstanding",
    title: `High cash need carries ${pct(highCl, T.dcl + T.pcl)} of the credit line`,
    subtitle: `DL utilises ${pct(T.dos, T.dcl)} of its ${M(T.dcl)} line against PL's ${pct(T.pos, T.pcl)} of ${M(T.pcl)} — PL draws down far less of what it is granted.`,
    headers: ["Cash need", "Risk level", "DL credit line", "DL outstanding", "DL util.", "PL credit line", "PL outstanding", "PL util."],
    colW: [1.3, 1.3, 1.9, 1.9, 0.95, 1.9, 1.9, 0.95],
    numericFrom: 2,
    body,
    total: ["Total", "", fmt(T.dcl), fmt(T.dos), pct(T.dos, T.dcl), fmt(T.pcl), fmt(T.pos), pct(T.pos, T.pcl)],
    footnote:
      SRC +
      "Credit line and outstanding total 1,092,262,000 and 660,543,706, matching the customer-base table exactly. " +
      "This table has no customer counts in the source, so utilisation is the only derived column.",
    notes:
      "The DL/PL utilisation gap is the story: DL draws 61.9% of its line, PL only 55.2%. " +
      "High cash need holds 691.7M of the 1,092.3M credit line across both loan types.",
  });
}

/* --- slides 7-10: the four AIS breakdowns by book date ------------------ */
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
