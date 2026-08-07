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

const out = path.join(__dirname, "Lending_Book_Loan_Summary.pptx");
pres.writeFile({ fileName: out }).then(() => console.log("wrote", out));
