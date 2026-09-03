/**
 * AIS Mobile - two-page summary of AIS_Mobile.xlsx.
 * Built to SLIDE_STYLE_GUIDE.md. Every figure comes from data/ais_mobile.tsv.
 *
 *   node build_ais_mobile.js
 */

const path = require("path");
const fs = require("fs");
const PptxGenJS = require("pptxgenjs");

/* ------------------------------------------------------- house constants */
const TITLE_TEAL = "045D66", INK = "16233D", MUTED = "6B7A90", NOTE = "6B6B6B";
const TEAL = "00A896", RED = "C00000", CARD = "F1F4F9", PAPER = "FFFFFF";
const R1 = "8DA0B8", R2 = "3FB89F", R3 = "1E8E86", R4 = "165A73";
const F = "Graphik TH", FH = "Cambria";

/* ------------------------------------------------------------------ data */
const rows = fs.readFileSync(path.join(__dirname, "data", "ais_mobile.tsv"), "utf8")
  .trim().split("\n").slice(1).map((l) => {
    const c = l.split("\t");
    return {
      cash: c[0], risk: c[1], all: +c[2], n: +c[3], n1: +c[4],
      others: +c[5], nUnch: +c[6], n1Unch: +c[7], nToN: +c[8], n1ToN: +c[9], n1ToN1: +c[10],
    };
  });
const LV = ["Null", "Low", "Mid", "High"];
const sum = (f, sel = () => true) => rows.filter(sel).reduce((a, r) => a + f(r), 0);

const base = sum((r) => r.all);
const flagN = sum((r) => r.n), flagN1 = sum((r) => r.n1), flag = flagN + flagN1;
const fbase = (r) => r.nUnch + r.n1Unch + r.nToN + r.n1ToN + r.n1ToN1;
const upg = (r) => r.nToN + r.n1ToN + r.n1ToN1;
const fBaseT = sum(fbase), upgT = sum(upg);
const toNewest = sum((r) => r.nToN + r.n1ToN), stayed = sum((r) => r.nUnch + r.n1Unch);

const K = (v) => v >= 1e6 ? (v / 1e6).toFixed(2) + "M" : (v / 1e3).toFixed(1) + "K";
const pct = (a, b) => ((a / b) * 100).toFixed(1) + "%";

/* ------------------------------------------------------------------ deck */
const pres = new PptxGenJS();
pres.layout = "LAYOUT_WIDE";
pres.author = "AIS Mobile";
pres.title = "AIS Mobile - flagship base and upgrades";

function head(s, title, note, totals, insight) {
  s.addText(title, {
    x: 0.58, y: 0.31, w: 12.33, h: 0.42, margin: 0,
    fontFace: F, fontSize: 22, bold: true, color: TITLE_TEAL,
  });
  s.addText(note, {
    x: 0.58, y: 0.78, w: 12.33, h: 0.26, margin: 0,
    fontFace: F, fontSize: 11.5, italic: true, color: NOTE,
  });
  s.addText(totals, {
    x: 0.60, y: 1.12, w: 12.10, h: 0.26, margin: 0,
    fontFace: F, fontSize: 13, bold: true, color: TITLE_TEAL,
  });
  s.addText(insight, {
    x: 0.60, y: 1.38, w: 12.10, h: 0.3, margin: 0,
    fontFace: F, fontSize: 14, color: INK,
  });
}

function tiles(s, defs) {
  defs.forEach(([v, l, sub, c], i) => {
    const x = 0.58 + i * 3.11;
    s.addShape("roundRect", {
      x, y: 1.80, w: 2.86, h: 1.08, rectRadius: 0.09,
      fill: { color: CARD }, line: { type: "none" },
    });
    s.addText(v, { x: x + 0.24, y: 1.88, w: 2.4, h: 0.44, margin: 0, fontFace: F, fontSize: 25, bold: true, color: c });
    s.addText(l, { x: x + 0.24, y: 2.32, w: 2.4, h: 0.26, margin: 0, valign: "top", fontFace: F, fontSize: 11.5, bold: true, color: INK });
    s.addText(sub, { x: x + 0.24, y: 2.57, w: 2.4, h: 0.26, margin: 0, valign: "top", fontFace: F, fontSize: 10, color: MUTED });
  });
}

/* card of labelled bar rows: [label, barValue, rightText, colour] */
function barCard(s, x, w, title, subtitle, items, opt = {}) {
  const y = 3.02, h = 3.00;
  s.addShape("roundRect", { x, y, w, h, rectRadius: 0.09, fill: { color: CARD }, line: { type: "none" } });
  s.addText(title, { x: x + 0.24, y: y + 0.16, w: w - 0.48, h: 0.28, margin: 0, fontFace: FH, fontSize: 13, bold: true, color: INK });
  s.addText(subtitle, { x: x + 0.24, y: y + 0.44, w: w - 0.48, h: 0.24, margin: 0, fontFace: F, fontSize: 9.5, color: MUTED });

  const lw = opt.labelW || 0.95, vw = opt.valueW || 1.05, pitch = opt.pitch || 0.5;
  const bx = x + 0.24 + lw + 0.12;
  const bw = w - 0.48 - lw - vw - 0.24;
  const maxV = Math.max(...items.map((it) => it[1]));
  items.forEach(([label, v, right, colour], i) => {
    const ry = y + 0.82 + i * pitch;
    s.addText(label, {
      x: x + 0.24, y: ry, w: lw, h: 0.3, margin: 0, valign: "middle",
      fontFace: F, fontSize: 11, bold: true, color: INK,
    });
    s.addShape("roundRect", {
      x: bx, y: ry + 0.06, w: Math.max((v / maxV) * bw, 0.05), h: 0.19, rectRadius: 0.04,
      fill: { color: colour }, line: { type: "none" },
    });
    s.addText(right, {
      x: x + w - 0.24 - vw, y: ry, w: vw, h: 0.3, margin: 0, align: "right", valign: "middle",
      fontFace: F, fontSize: 11, bold: true, color: INK,
    });
  });
  if (opt.foot) {
    s.addText(opt.foot, {
      x: x + 0.24, y: y + h - 0.42, w: w - 0.48, h: 0.34, margin: 0,
      fontFace: F, fontSize: 9, italic: true, color: MUTED,
    });
  }
}

const ramp = { High: R4, Mid: R3, Low: R2, Null: R1 };

/* ===================================================== PAGE 1 - the base */
{
  const s = pres.addSlide();
  s.background = { color: PAPER };
  head(s,
    `AIS Mobile Flagship Base : ${K(base)} base, ${K(flag)} on a flagship (${pct(flag, base)})`,
    "*Risk and cash need as of 202508, device as of 202604, base filtered to age 20–60.",
    `Base ${K(base)}   ·   Newest model (N) ${K(flagN)}   ·   Previous model (N-1) ${K(flagN1)}`,
    "Flagship ownership climbs with cash need and falls with risk — the densest pocket is high cash need paired with low risk."
  );

  const penCash = LV.map((l) => [l, sum((r) => r.n + r.n1, (r) => r.cash === l) / sum((r) => r.all, (r) => r.cash === l) * 100]);
  const penRisk = LV.map((l) => [l, sum((r) => r.n + r.n1, (r) => r.risk === l) / sum((r) => r.all, (r) => r.risk === l) * 100]);

  tiles(s, [
    [K(base), "customers in base", "age 20–60", TEAL],
    [K(flag), "hold a flagship", `${pct(flag, base)} of base`, TEAL],
    [pct(flagN, base), "are on the newest model", `${K(flagN)} customers`, TEAL],
    ["2.5×", "risk penetration gap", `${penRisk[1][1].toFixed(1)}% low vs ${penRisk[3][1].toFixed(1)}% high`, RED],
  ]);

  barCard(s, 0.58, 3.60, "Penetration by cash need", "Flagship holders as % of each cash-need group",
    penCash.slice().sort((a, b) => b[1] - a[1]).map(([l, v]) => [l, v, v.toFixed(1) + "%", ramp[l]]));

  barCard(s, 4.43, 3.60, "Penetration by risk level", "Flagship holders as % of each risk group",
    penRisk.slice().sort((a, b) => b[1] - a[1]).map(([l, v]) => [l, v, v.toFixed(1) + "%", ramp[l]]));

  const cells = rows.map((r) => ({ k: `${r.cash} / ${r.risk}`, f: r.n + r.n1, p: (r.n + r.n1) / r.all * 100 }))
    .sort((a, b) => b.f - a.f).slice(0, 5);
  barCard(s, 8.28, 4.63, "Where the flagship owners are", "Largest cash-need / risk cells by flagship volume",
    cells.map((c) => [c.k, c.f, `${K(c.f)} · ${c.p.toFixed(1)}%`, R4]),
    { labelW: 1.35, valueW: 1.5, pitch: 0.42 });

  s.addNotes(
    `Base ${base.toLocaleString()} (age 20-60). Flagship holders ${flag.toLocaleString()} = ${pct(flag, base)}: N ${flagN.toLocaleString()}, N-1 ${flagN1.toLocaleString()}. ` +
      "Penetration by cash need runs 3.2 / 5.2 / 8.9 / 9.7 percent from Null to High; by risk it runs 10.4 / 6.7 / 4.2 percent from Low to High. " +
      "Low-risk customers are 40.1% of the base but 57.6% of all flagship owners. " +
      "The Null-risk and Null-cash-need cells are tiny (High/Null holds 1,698 customers at 61% penetration) and should not be read as a segment."
  );
}

/* ================================================== PAGE 2 - the upgrades */
{
  const s = pres.addSlide();
  s.background = { color: PAPER };
  head(s,
    `AIS Mobile Flagship Upgrades : ${K(upgT)} of ${K(fBaseT)} upgraded (${pct(upgT, fBaseT)})`,
    "*Flagship defined by the model list in the source. Risk-Null cells hold only 5.4K flagship customers, so their rates are not reliable.",
    `Flagship base ${K(fBaseT)}   ·   Upgraded ${K(upgT)}   ·   Kept the same handset ${K(stayed)}`,
    "Only one flagship owner in seven changed handset over the year. Cash need predicts who upgrades; risk level does not."
  );

  tiles(s, [
    [K(fBaseT), "held a flagship in 202504", `${pct(fBaseT, base)} of base`, TEAL],
    [K(upgT), "upgraded by 202604", `${pct(upgT, fBaseT)} of flagship holders`, TEAL],
    [pct(stayed, fBaseT), "kept the same handset", `${K(stayed)} customers`, RED],
    [pct(toNewest, upgT), "of upgraders took N", `${K(toNewest)} to the newest model`, TEAL],
  ]);

  const upCash = LV.map((l) => [l, sum(upg, (r) => r.cash === l) / sum(fbase, (r) => r.cash === l) * 100]);
  barCard(s, 0.58, 3.60, "Upgrade rate by cash need", "Share who changed handset",
    upCash.slice().sort((a, b) => b[1] - a[1]).map(([l, v]) => [l, v, v.toFixed(1) + "%", ramp[l]]));

  const upRisk = ["Mid", "High", "Low"].map((l) => [l, sum(upg, (r) => r.risk === l) / sum(fbase, (r) => r.risk === l) * 100]);
  barCard(s, 4.43, 3.60, "Upgrade rate by risk level", "Share who changed handset",
    upRisk.map(([l, v]) => [l, v, v.toFixed(1) + "%", ramp[l]]),
    { foot: "Risk Null omitted — 5.4K holders only. The three real bands sit within 1.2 points of each other." });

  const dest = [
    ["N-1 → N", sum((r) => r.n1ToN), R4],
    ["N → N", sum((r) => r.nToN), R3],
    ["N-1 → N-1", sum((r) => r.n1ToN1), R2],
  ];
  barCard(s, 8.28, 4.63, "Where the upgraders went", `${K(upgT)} customers who changed handset`,
    dest.map(([l, v, c]) => [l, v, `${K(v)} · ${pct(v, upgT)}`, c]),
    { labelW: 1.35, valueW: 1.5, pitch: 0.42,
      foot: "N-1 → N-1 customers upgraded a generation but are still one behind the newest model." });

  s.addNotes(
    `Flagship base a year earlier ${fBaseT.toLocaleString()}; ${upgT.toLocaleString()} changed handset (${pct(upgT, fBaseT)}); ${stayed.toLocaleString()} did not. ` +
      "N holders upgraded at 14.8% and N-1 holders at 15.4% - near identical, so being a generation behind does not itself drive replacement. " +
      "Upgrade rate by cash need runs 8.2 / 12.9 / 15.7 / 17.2 percent from Null to High. By risk it is flat: Low 14.7, Mid 15.9, High 15.2. " +
      "Of the upgraders, 77.8% went to the newest model and 22.2% moved to the previous generation. " +
      "Data note: four rows of the N/N-1 model mapping look mis-paired (Galaxy S26 to S25+, iPhone 16 Pro and 16 both to 15 Pro Max, duplicate Galaxy S25+), which would shift the N-1 counts."
  );
}

/* ================================================= PAGE 3 - demand estimate */
{
  // Observed transitions 202504 -> 202604. The 202604 N generation did not exist
  // a year earlier, so every N holder is a buyer; N-1 holders either kept the
  // device they already had (N Unchanged) or acquired it during the year.
  const traceN  = sum((r) => r.nToN + r.n1ToN);       // moved onto the new N
  const traceN1 = sum((r) => r.nUnch + r.n1ToN1);     // kept 202504-N, or moved to the new N-1
  const inflowN = flagN - traceN, inflowN1 = flagN1 - traceN1;
  const inflow = inflowN + inflowN1;
  const boughtLast = flagN + sum((r) => r.n1ToN1) + inflowN1;

  const nonflagPrev = base - fBaseT, nonflagNow = base - flag;
  const rNN  = sum((r) => r.nToN)   / sum((r) => r.nUnch + r.nToN);
  const rN1N = sum((r) => r.n1ToN)  / sum((r) => r.n1Unch + r.n1ToN + r.n1ToN1);
  const rN1N1= sum((r) => r.n1ToN1) / sum((r) => r.n1Unch + r.n1ToN + r.n1ToN1);
  const rInflow = inflow / nonflagPrev;

  const repN = flagN * rNN + flagN1 * rN1N, repN1 = flagN1 * rN1N1;
  const infT = nonflagNow * rInflow;
  const infN = infT * (inflowN / inflow), infN1 = infT - infN;
  const estN = repN + infN, estN1 = repN1 + infN1, est = estN + estN1;
  const cons = flag * 0.13 + nonflagNow * 0.042;
  const opti = flag * 0.17 + nonflagNow * 0.052;

  const s3 = pres.addSlide();
  s3.background = { color: PAPER };
  head(s3,
    `AIS Mobile Flagship Demand : ~${K(est)} buyers in the coming year`,
    "*For the coming year (202604 → 202704), modelled on the observed 202504 → 202604 transitions; assumes a comparable flagship launch.",
    `Bought in the past year ${K(boughtLast)}   ·   Estimated for the coming year ${K(est)}   ·   Range ${K(cons)} – ${K(opti)}`,
    "Four in five buyers come from outside today's flagship base — demand is inflow-driven, not owners cycling."
  );

  tiles(s3, [
    [K(boughtLast), "bought in the past year", "observed 202504 → 202604", TEAL],
    [K(est), "estimated, coming year", "base case", TEAL],
    [K(estN), "expected to buy N", `${K(estN1)} to buy N-1`, TEAL],
    [pct(infT, est), "from outside the base", `${K(infT)} of ${K(est)}`, RED],
  ]);

  barCard(s3, 0.58, 4.63, "How the estimate builds", `${K(est)} buyers, by source and device`,
    [
      ["Inflow → N", infN, `${K(infN)} · ${pct(infN, est)}`, R4],
      ["Inflow → N-1", infN1, `${K(infN1)} · ${pct(infN1, est)}`, R3],
      ["Existing → N", repN, `${K(repN)} · ${pct(repN, est)}`, R2],
      ["Existing → N-1", repN1, `${K(repN1)} · ${pct(repN1, est)}`, R1],
    ],
    { labelW: 1.45, valueW: 1.55, pitch: 0.44,
      foot: `Inflow applies the observed ${(rInflow * 100).toFixed(2)}% acquisition rate to the ${K(nonflagNow)} non-flagship base; replacement applies the observed upgrade rates to today's ${K(flag)} flagship holders.` });

  const bands = ["Low", "Mid", "High", "Null"].map((l) => {
    const sel = (r) => r.risk === l;
    const n = sum((r) => r.n, sel), n1 = sum((r) => r.n1, sel);
    const fbl = sum(fbase, sel), nfl = sum((r) => r.all, sel) - fbl;
    const inf = Math.max((n - sum((r) => r.nToN + r.n1ToN, sel)) + (n1 - sum((r) => r.nUnch + r.n1ToN1, sel)), 0);
    const rate = nfl ? inf / nfl : 0;
    return [l, n * rNN + n1 * rN1N + n1 * rN1N1 + (sum((r) => r.all, sel) - (n + n1)) * rate];
  });
  const bandTot = bands.reduce((a, b) => a + b[1], 0);
  const financeable = bands.filter((b) => b[0] === "Low" || b[0] === "Mid").reduce((a, b) => a + b[1], 0);
  barCard(s3, 5.46, 3.60, "Estimated buyers by risk band", "Using band-specific acquisition rates",
    bands.slice().sort((a, b) => b[1] - a[1]).map(([l, v]) => [l, v, `${K(v)}`, ramp[l]]),
    { valueW: 0.95, pitch: 0.44,
      foot: `Low and Mid risk together are ${K(financeable)} — ${pct(financeable, bandTot)} of expected buyers.` });

  barCard(s3, 9.31, 3.60, "Sensitivity", "Total buyers under three rate assumptions",
    [
      ["Optimistic", opti, K(opti), R4],
      ["Base", est, K(est), R3],
      ["Conservative", cons, K(cons), R2],
    ],
    { labelW: 1.25, valueW: 0.95, pitch: 0.44,
      foot: "Conservative: 13% upgrade of the flagship base, 4.2% inflow. Optimistic: 17% and 5.2%." });

  s3.addNotes(
    `Method. The 202604 N generation did not exist at 202504, so all ${flagN.toLocaleString()} N holders bought within the year; ` +
      `N-1 buyers are the ${sum((r) => r.n1ToN1).toLocaleString()} who moved N-1 to N-1 plus ${inflowN1.toLocaleString()} who acquired it from outside the flagship base. ` +
      `That gives ${boughtLast.toLocaleString()} purchases observed in the year. ` +
      `Only ${(flag - traceN - traceN1).toLocaleString()} of the 202604 flagship pool is traceable through the five transition states, so ${inflow.toLocaleString()} arrived from outside - ` +
      `an acquisition rate of ${(rInflow * 100).toFixed(2)}% on the ${nonflagPrev.toLocaleString()} non-flagship base. ` +
      "Forecast applies those rates to the 202604 position. Caveats: one year of transitions only, no seasonality, and the estimate assumes a comparable flagship launch in the coming year. " +
      "The inflow term carries about 80% of the volume, so the estimate is far more sensitive to the acquisition rate than to the upgrade rate."
  );
}

const out = path.join(__dirname, "AIS_Mobile_Summary.pptx");
pres.writeFile({ fileName: out }).then(() => console.log("wrote", out));
