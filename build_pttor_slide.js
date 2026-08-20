/**
 * PTT OR - one-page summary, data as of Apr 2026.
 * Every figure derives from data/pttor.tsv exactly as supplied.
 *
 *   node build_pttor_slide.js
 */

const path = require("path");
const fs = require("fs");
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
const MUTED = "6B7A90";
const R1 = "8DA0B8";
const R2 = "3FB89F";
const R4 = "165A73";
const H = "Cambria";
const B = "Calibri";

/* ------------------------------------------------------------------- data */
const D = {};
fs.readFileSync(path.join(__dirname, "data", "pttor.tsv"), "utf8")
  .trim().split("\n").slice(1)
  .forEach((l) => { const [k, v] = l.split("\t"); D[k] = Number(v); });

const base = D.base, mtu = D.mtu_all, both = D.mtu_both;
const amzOnly = D.mtu_amazon, oilOnly = D.mtu_oil;
const amzUsers = amzOnly + both, oilUsers = oilOnly + both;
const cntA = D.txn_cnt_amz, amtA = D.txn_amt_amz;
const cntO = D.txn_cnt_oil, amtO = D.txn_amt_oil;
const cntT = cntA + cntO, amtT = amtA + amtO;

const p1 = (a, b) => ((a / b) * 100).toFixed(1) + "%";
const M = (v) => (v / 1e6).toFixed(2) + "M";
const Bn = (v) => (v / 1e9).toFixed(2) + "B";
const n0 = (v) => Math.round(v).toLocaleString("en-US");

/* ------------------------------------------------------------------- deck */
const pres = new PptxGenJS();
pres.layout = "LAYOUT_WIDE";
pres.author = "PTT OR";
pres.title = "PTT OR - summary as of Apr 2026";

const s = pres.addSlide();
s.background = { color: PAPER };

s.addText("PTT OR  ·  DATA AS OF APR 2026", {
  x: 0.6, y: 0.36, w: 12.1, h: 0.26, margin: 0,
  fontFace: B, fontSize: 11.5, bold: true, charSpacing: 2.2, color: TEAL,
});
s.addText("Half the transactions, nine-tenths of the value", {
  x: 0.6, y: 0.66, w: 12.1, h: 0.62, margin: 0, valign: "top",
  fontFace: H, fontSize: 32, bold: true, color: INK,
});
s.addText(
  `Amazon and Oil run almost the same number of transactions, but Oil's ticket is ${(amtO / cntO / (amtA / cntA)).toFixed(0)}× larger, so it carries ${p1(amtO, amtT)} of the value.`,
  { x: 0.6, y: 1.18, w: 12.1, h: 0.28, margin: 0, valign: "top", fontFace: B, fontSize: 12.5, color: MUTED }
);

/* stat tiles ------------------------------------------------------------- */
const tiles = [
  [(base / 1e6).toFixed(1) + "M", "total base", "customers", TEAL],
  [M(mtu), "transacting users", `monthly · ${p1(mtu, base)} of base`, TEAL],
  [M(cntT), "transactions", `avg ticket ${(amtT / cntT).toFixed(0)}`, TEAL],
  [Bn(amtT), "total value", `Oil ${p1(amtO, amtT)} of it`, AMBER],
];
tiles.forEach(([v, l, sub, c], i) => {
  const x = 0.6 + i * 3.1;
  s.addShape("roundRect", {
    x, y: 1.56, w: 2.8, h: 1.1, rectRadius: 0.09,
    fill: { color: TINT }, line: { type: "none" },
  });
  s.addText(v, { x: x + 0.25, y: 1.66, w: 2.3, h: 0.44, margin: 0, fontFace: H, fontSize: 25, bold: true, color: c });
  s.addText(l, { x: x + 0.25, y: 2.1, w: 2.3, h: 0.28, margin: 0, valign: "top", fontFace: B, fontSize: 11.5, bold: true, color: INK });
  s.addText(sub, { x: x + 0.25, y: 2.36, w: 2.3, h: 0.26, margin: 0, valign: "top", fontFace: B, fontSize: 10, color: MUTED });
});

/* left panel - who transacts --------------------------------------------- */
s.addShape("roundRect", {
  x: 0.6, y: 2.86, w: 5.95, h: 2.44, rectRadius: 0.09,
  fill: { color: TINT }, line: { type: "none" },
});
s.addText("Who transacts", {
  x: 0.9, y: 3.02, w: 5.35, h: 0.3, margin: 0, fontFace: H, fontSize: 15, bold: true, color: INK,
});
s.addText(`The ${M(mtu)} monthly transacting users, split by what they buy`, {
  x: 0.9, y: 3.32, w: 5.35, h: 0.26, margin: 0, fontFace: B, fontSize: 10, color: MUTED,
});

const seg = [
  ["Oil only", oilOnly, R4],
  ["Both", both, R2],
  ["Amazon only", amzOnly, R1],
];
const segMax = Math.max(...seg.map(([, v]) => v));
seg.forEach(([name, v, color], i) => {
  const y = 3.72 + i * 0.44;
  s.addText(name, {
    x: 0.9, y, w: 1.5, h: 0.32, margin: 0, valign: "middle",
    fontFace: B, fontSize: 11.5, bold: true, color: INK,
  });
  s.addShape("roundRect", {
    x: 2.5, y: y + 0.06, w: (v / segMax) * 2.15, h: 0.2, rectRadius: 0.04,
    fill: { color }, line: { type: "none" },
  });
  s.addText(`${M(v)}  ·  ${p1(v, mtu)}`, {
    x: 4.75, y, w: 1.5, h: 0.32, margin: 0, align: "right", valign: "middle",
    fontFace: H, fontSize: 11.5, bold: true, color: INK,
  });
});

/* right panel - transactions vs value ------------------------------------ */
s.addShape("roundRect", {
  x: 6.75, y: 2.86, w: 5.95, h: 2.44, rectRadius: 0.09,
  fill: { color: TINT }, line: { type: "none" },
});
const chartCats = [`Value  ${Bn(amtT)}`, `Transactions  ${M(cntT)}`];
s.addChart(
  pres.charts.BAR,
  [
    { name: "Amazon", labels: chartCats, values: [Number(p1(amtA, amtT).replace("%", "")), Number(p1(cntA, cntT).replace("%", ""))] },
    { name: "Oil", labels: chartCats, values: [Number(p1(amtO, amtT).replace("%", "")), Number(p1(cntO, cntT).replace("%", ""))] },
  ],
  {
    x: 6.88, y: 2.96, w: 5.72, h: 2.24,
    barDir: "bar",
    barGrouping: "percentStacked",
    barGapWidthPct: 80,
    chartColors: [R2, R4],
    showTitle: true,
    title: "Share of transactions and value",
    titleFontFace: B, titleFontSize: 12, titleColor: INK,
    showValue: true,
    dataLabelPosition: "ctr",
    dataLabelFontFace: B, dataLabelFontSize: 10, dataLabelColor: PAPER,
    dataLabelFormatCode: '0.0"%"',
    showLegend: true, legendPos: "b",
    legendFontFace: B, legendFontSize: 10, legendColor: INK,
    catAxisLabelFontFace: B, catAxisLabelFontSize: 10.5, catAxisLabelColor: INK,
    valAxisHidden: true,
    valGridLine: { style: "none" },
    catGridLine: { style: "none" },
  }
);

/* bottom band ------------------------------------------------------------ */
s.addShape("roundRect", {
  x: 0.6, y: 5.46, w: 12.1, h: 1.3, rectRadius: 0.09,
  fill: { color: NAVY }, line: { type: "none" },
});
s.addText("WHAT SEPARATES THE TWO", {
  x: 1.0, y: 5.64, w: 11.3, h: 0.26, margin: 0,
  fontFace: B, fontSize: 10.5, bold: true, charSpacing: 1.8, color: MINT,
});
s.addText(
  [
    { text: `Average ticket is ${(amtA / cntA).toFixed(0)} at Amazon against ${(amtO / cntO).toFixed(0)} at Oil`, options: { color: MINT, bold: true } },
    {
      text: ` — so near-equal transaction counts split value ${p1(amtA, amtT).replace(".0%", "%")} / ${p1(amtO, amtT).replace(".0%", "%")}. Amazon users transact more often (${(cntA / amzUsers).toFixed(1)} each against ${(cntO / oilUsers).toFixed(1)}) but spend ${n0(amtA / amzUsers)} against ${n0(amtO / oilUsers)}.`,
      options: { color: "C3D0E2" },
    },
    {
      text: `\nCross-shopping runs one way: ${p1(both, amzUsers)} of Amazon users also buy oil, but only ${p1(both, oilUsers)} of Oil users visit Amazon.`,
      options: { color: "C3D0E2" },
    },
  ],
  {
    x: 1.0, y: 5.92, w: 11.3, h: 0.76, margin: 0, valign: "top",
    lineSpacing: 16, fontFace: B, fontSize: 11,
  }
);

s.addText(
  "Source: figures as supplied, PTT OR, data as of Apr 2026; currency not stated. 'mtu amazon' and 'mtu oil' are read as single-brand users and 'mtu both' as the overlap — on that reading the three sum to " +
    `${n0(oilOnly + both + amzOnly)} against the stated ${n0(mtu)}, a gap of ${n0(oilOnly + both + amzOnly - mtu)} worth confirming. ` +
    "Per-user figures divide each brand's transactions by its total users, single-brand plus both. All percentages are computed from the supplied counts.",
  { x: 0.6, y: 6.88, w: 12.1, h: 0.5, margin: 0, fontFace: B, fontSize: 8.5, italic: true, color: MUTED }
);

s.addNotes(
  "One-page read of the PTT OR figures as of Apr 2026. " +
    `Reach is thin: ${n0(mtu)} monthly transacting users is ${p1(mtu, base)} of the ${n0(base)} base. ` +
    `Oil is the larger footprint - ${n0(oilUsers)} users against Amazon's ${n0(amzUsers)} - and dominates value at ${p1(amtO, amtT)} on ${p1(cntO, cntT)} of transactions. ` +
    `Average ticket ${(amtA / cntA).toFixed(2)} vs ${(amtO / cntO).toFixed(2)}, a ${(amtO / cntO / (amtA / cntA)).toFixed(1)}x gap. ` +
    "The cross-shop asymmetry is the actionable piece: Amazon users are far more likely to also buy fuel than the reverse, so Oil is the harder brand to pull customers into."
);

const out = path.join(__dirname, "PTT_OR_Summary.pptx");
pres.writeFile({ fileName: out }).then(() => console.log("wrote", out));
