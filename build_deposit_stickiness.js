/**
 * Deposit Stickiness Framework - rebuilt to SLIDE_STYLE_GUIDE.md.
 * Content is faithful to Deposit_Stickiness_Framework_Updated_v11; every object
 * is a native PowerPoint shape, text box or table, so the deck stays editable.
 * Matrices come from data/deposit_stickiness.tsv.
 *
 *   node build_deposit_stickiness.js
 */

const path = require("path");
const fs = require("fs");
const PptxGenJS = require("pptxgenjs");

/* ------------------------------------------------------- house constants */
const TITLE_TEAL = "045D66", INK = "16233D", MUTED = "6B7A90", NOTE = "6B6B6B";
const TEAL = "00A896", RED = "C00000", AMBER = "E9A03B", CARD = "F1F4F9";
const PAPER = "FFFFFF", RULE = "D8DFE9", WARM = "FBF0E6";
const R1 = "8DA0B8", R2 = "3FB89F", R3 = "1E8E86", R4 = "165A73";
const F = "Graphik TH", FH = "Cambria";

const BANDS = [
  ["80–100", "Very High", R4],
  ["60–79", "High", R3],
  ["40–59", "Medium", R2],
  ["20–39", "Low", AMBER],
  ["<20", "Very Low", R1],
];

/* ------------------------------------------------------------------ data */
const M = {};
fs.readFileSync(path.join(__dirname, "data", "deposit_stickiness.tsv"), "utf8")
  .trim().split("\n").slice(1).forEach((l) => {
    const c = l.split("\t");
    (M[c[0]] = M[c[0]] || []).push({ s: c[1], v: c.slice(2).map(Number) });
  });
const RS = ["Very Low", "Low", "Medium", "High", "Very High"];
const gTot = (m) => M[m].reduce((a, r) => a + r.v.reduce((x, y) => x + y, 0), 0);
const gHiRS = (m) => M[m].reduce((a, r) => a + r.v[3] + r.v[4], 0);

const smCust = gTot("savemax_count"), smHiRS = gHiRS("savemax_count");
const smBal = gTot("savemax_balance"), smAtRisk = gHiRS("savemax_balance");
const smSticky = M.savemax_balance.filter((r) => r.s === "High" || r.s === "Very High")
  .reduce((a, r) => a + r.v[3] + r.v[4], 0);
const smRunoff = smAtRisk - smSticky;
const allCust = gTot("all_count"), allHiRS = gHiRS("all_count");
const allLowStick = M.all_count.filter((r) => r.s === "Low" || r.s === "Very Low")
  .reduce((a, r) => a + r.v.reduce((x, y) => x + y, 0), 0);

const k = (v) => (v / 1e3).toFixed(1) + "k";
const m = (v) => (v / 1e6).toFixed(0) + "m";
const mm = (v) => (v / 1e6).toFixed(1) + "m";
const pc = (a, b) => Math.round((a / b) * 100) + "%";

/* --------------------------------------------------------------- helpers */
const pres = new PptxGenJS();
pres.layout = "LAYOUT_WIDE";
pres.author = "Deposit Analytics";
pres.title = "Deposit Stickiness Framework";

function head(s, title, note, totals, insight) {
  s.addText(title, { x: 0.58, y: 0.31, w: 12.33, h: 0.42, margin: 0, fontFace: F, fontSize: 22, bold: true, color: TITLE_TEAL });
  if (note) s.addText(note, { x: 0.58, y: 0.78, w: 12.33, h: 0.26, margin: 0, fontFace: F, fontSize: 11.5, italic: true, color: NOTE });
  if (totals) s.addText(totals, { x: 0.60, y: 1.09, w: 12.10, h: 0.26, margin: 0, fontFace: F, fontSize: 13, bold: true, color: TITLE_TEAL });
  if (insight) s.addText(insight, { x: 0.60, y: 1.40, w: 12.10, h: 0.3, margin: 0, fontFace: F, fontSize: 14, color: INK });
}

function card(s, x, y, w, h, fill) {
  s.addShape("roundRect", { x, y, w, h, rectRadius: 0.09, fill: { color: fill || CARD }, line: { type: "none" } });
}
function cardHead(s, x, y, w, title, sub) {
  s.addText(title, { x: x + 0.28, y: y + 0.16, w: w - 0.56, h: 0.3, margin: 0, fontFace: FH, fontSize: 14, bold: true, color: INK });
  if (sub) s.addText(sub, { x: x + 0.28, y: y + 0.46, w: w - 0.56, h: 0.26, margin: 0, fontFace: F, fontSize: 9.5, color: MUTED });
}
function tiles(s, defs, y) {
  const gap = 0.25, w = (12.33 - gap * (defs.length - 1)) / defs.length;
  defs.forEach(([v, l, sub, c], i) => {
    const x = 0.58 + i * (w + gap);
    card(s, x, y, w, 1.08);
    s.addText(v, { x: x + 0.24, y: y + 0.08, w: w - 0.48, h: 0.44, margin: 0, fontFace: F, fontSize: (defs[i][4] || 24), bold: true, color: c || TEAL });
    s.addText(l, { x: x + 0.24, y: y + 0.52, w: w - 0.48, h: 0.26, margin: 0, valign: "top", fontFace: F, fontSize: 11.5, bold: true, color: INK });
    if (sub) s.addText(sub, { x: x + 0.24, y: y + 0.77, w: w - 0.48, h: 0.26, margin: 0, valign: "top", fontFace: F, fontSize: 9.5, color: MUTED });
  });
}
/* label / value rows inside a card */
function rows(s, x, y, w, items, opt = {}) {
  const lw = opt.labelW || 2.0, pitch = opt.pitch || 0.36;
  items.forEach(([l, v], i) => {
    s.addText(l, { x: x + 0.28, y: y + i * pitch, w: lw, h: 0.3, margin: 0, valign: "middle", fontFace: F, fontSize: opt.size || 10.5, bold: true, color: TITLE_TEAL });
    s.addText(v, { x: x + 0.28 + lw, y: y + i * pitch, w: w - 0.56 - lw, h: 0.3, margin: 0, valign: "middle", fontFace: F, fontSize: opt.size || 10.5, color: INK });
  });
}
/* the five score bands as swatch + range + label */
function bandRow(s, x, y, w, size = 9.5) {
  const cw = w / BANDS.length;
  BANDS.forEach(([range, name, col], i) => {
    const bx = x + i * cw;
    s.addShape("roundRect", { x: bx, y, w: 0.34, h: 0.19, rectRadius: 0.04, fill: { color: col }, line: { type: "none" } });
    s.addText(range, { x: bx + 0.42, y: y - 0.06, w: cw - 0.46, h: 0.22, margin: 0, valign: "middle", fontFace: F, fontSize: size, bold: true, color: INK });
    s.addText(name, { x: bx + 0.42, y: y + 0.13, w: cw - 0.46, h: 0.22, margin: 0, valign: "middle", fontFace: F, fontSize: size - 1, color: MUTED });
  });
}
/* white -> colour blend, used for the heat matrices */
function blend(hex, t) {
  const c = [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16));
  return c.map((v) => Math.round(255 + (v - 255) * t).toString(16).padStart(2, "0")).join("").toUpperCase();
}
/* 5x5 stickiness x rate-sensitivity matrix as a native table. The High and
   Very High rate-sensitivity columns are the at-risk block, tinted in the same
   red as the at-risk summary figure; nothing else is colour-coded. */
const AT_RISK_TINT = blend(RED, 0.13);
function matrix(s, x, y, w, key, fmt) {
  const labelW = w * 0.22, cellW = (w - labelW) / 5;
  const atRisk = (i) => i >= 3;
  const hdr = [
    { text: "Stickiness", options: { fill: { color: TITLE_TEAL }, color: PAPER, bold: true, fontFace: F, fontSize: 9, align: "left" } },
    ...RS.map((n) => ({
      text: n,
      options: { fill: { color: TITLE_TEAL }, color: PAPER, bold: true, fontFace: F, fontSize: 9, align: "center" },
    })),
  ];
  const body = M[key].map((r) => [
    { text: r.s, options: { fill: { color: CARD }, color: INK, bold: true, fontFace: F, fontSize: 9.5, align: "left" } },
    ...r.v.map((v, i) => ({
      text: fmt(v),
      options: {
        fill: { color: atRisk(i) ? AT_RISK_TINT : PAPER },
        color: atRisk(i) ? RED : INK,
        bold: atRisk(i),
        fontFace: F, fontSize: 9.5, align: "right",
      },
    })),
  ]);
  s.addTable([hdr, ...body], {
    x, y, w, colW: [labelW, ...Array(5).fill(cellW)], rowH: 0.3,
    border: { type: "solid", color: RULE, pt: 0.5 }, valign: "middle", margin: [0.02, 0.08, 0.02, 0.08],
  });
}

/* ============================================================== SLIDE 1 */
{
  const s = pres.addSlide(); s.background = { color: PAPER };
  head(s, "Deposit Stickiness : customer-level framework",
    "*Pre-maturity view · behavioural stickiness score, independent of deposit size.",
    "Deposit Stickiness Score  =  40% BP  +  25% FC  +  20% RI  +  15% DE",
    "Each component scores 0–100 and unavailable components are re-weighted, so every customer gets a comparable score.");

  tiles(s, [
    ["40%", "Balance persistence", "Avg balance: latest ÷ prior period", TEAL],
    ["25%", "Funding continuity", "% eligible days with balance ≥ THB100", TEAL],
    ["20%", "Recurring inflow", "% eligible months with inflow ≥ THB100", TEAL],
    ["15%", "Daily engagement", "Active txn days ÷ 20-day / 90-day target*", AMBER],
  ], 1.80);

  card(s, 0.58, 3.06, 7.35, 3.30);
  cardHead(s, 0.58, 3.06, 7.35, "How the score works", "Component weights, and how balance persistence adapts to tenure");
  let wx = 0.86;
  [["40%", R4, 2.83], ["25%", R3, 1.77], ["20%", R2, 1.42], ["15%", AMBER, 1.06]].forEach(([lab, col, ww]) => {
    s.addShape("roundRect", { x: wx, y: 3.80, w: ww, h: 0.40, rectRadius: 0.05, fill: { color: col }, line: { type: "none" } });
    s.addText(lab, { x: wx, y: 3.80, w: ww, h: 0.40, margin: 0, align: "center", valign: "middle", fontFace: F, fontSize: 11, bold: true, color: PAPER });
    wx += ww + 0.04;
  });
  s.addText("Balance persistence adapts to tenure", { x: 0.86, y: 4.40, w: 6.8, h: 0.28, margin: 0, fontFace: F, fontSize: 11, bold: true, color: INK });
  [["≥ 60 days", "30D vs 30D"], ["30–59 days", "15D vs 15D"], ["14–29 days", "7D vs 7D"]].forEach(([a, b], i) => {
    const bx = 0.86 + i * 2.35;
    card(s, bx, 4.74, 2.20, 0.72, PAPER);
    s.addText(a, { x: bx + 0.18, y: 4.82, w: 1.9, h: 0.26, margin: 0, fontFace: F, fontSize: 11, bold: true, color: TITLE_TEAL });
    s.addText(b, { x: bx + 0.18, y: 5.06, w: 1.9, h: 0.26, margin: 0, fontFace: F, fontSize: 10, color: MUTED });
  });
  s.addText("BP = balance persistence · FC = funding continuity · RI = recurring inflow · DE = daily engagement", {
    x: 0.86, y: 5.62, w: 6.8, h: 0.5, margin: 0, fontFace: F, fontSize: 9.5, italic: true, color: MUTED,
  });

  card(s, 8.18, 3.06, 4.73, 3.30);
  cardHead(s, 8.18, 3.06, 4.73, "Score cut-offs", "Initial score bands (0–100)");
  BANDS.forEach(([range, name, col], i) => {
    const y = 3.78 + i * 0.42;
    s.addShape("roundRect", { x: 8.46, y: y + 0.04, w: 0.42, h: 0.22, rectRadius: 0.04, fill: { color: col }, line: { type: "none" } });
    s.addText(range, { x: 9.02, y, w: 1.1, h: 0.3, margin: 0, valign: "middle", fontFace: F, fontSize: 11, bold: true, color: INK });
    s.addText(name, { x: 10.16, y, w: 2.4, h: 0.3, margin: 0, valign: "middle", fontFace: F, fontSize: 11, color: INK });
  });
  card(s, 8.46, 5.92, 4.17, 0.34, WARM);
  s.addText("Rate sensitivity stays outside this score until validated against post-maturity runoff.", {
    x: 8.56, y: 5.92, w: 3.97, h: 0.34, margin: 0, valign: "middle", fontFace: F, fontSize: 9, italic: true, color: NOTE,
  });
  s.addNotes("Stickiness is a behavioural score built from four components and is deliberately independent of deposit size. Unavailable components are re-weighted so scores stay comparable. Rate sensitivity is kept as a separate score until post-maturity runoff validates it.");
}

/* ============================================================== SLIDE 2 */
{
  const s = pres.addSlide(); s.background = { color: PAPER };
  head(s, "Promo behaviour reveals rate sensitivity",
    "*Balance placement around the promotional caps is a behavioural signal, not a statement of customer intent.",
    "Rate Sensitivity  =  70% Cap Precision  +  30% (1 − Above-Cap Balance Ratio)",
    "Customers who stop near a promotional cap look rate-optimising; balance above the cap suggests lower rate sensitivity.");

  tiles(s, [
    ["THB 20k", "Save Max 4% cap", "Only the first 20k earns the high rate", AMBER],
    ["THB 1m", "Save More 2% cap", "Only the first 1m earns the high rate", TEAL],
    ["Promo dependency", "Promo-exposed ÷ total deposit", "Measures exposure, not intent", TITLE_TEAL, 16],
  ], 1.80);

  card(s, 0.58, 3.06, 7.35, 3.30);
  cardHead(s, 0.58, 3.06, 7.35, "Three behaviour patterns", "Orange = balance within the promotional cap · teal = balance above the cap");
  [
    ["Cap optimizer", "~20k in Save Max + ~1m in Save More", 1.00, "VERY HIGH", RED],
    ["Mixed saver", "Uses promo pockets but also holds balance above cap", 0.62, "MEDIUM", AMBER],
    ["Core depositor", "Material balance above cap or in Main Pocket", 0.34, "LOW", R3],
  ].forEach(([name, desc, share, sig, col], i) => {
    const y = 3.86 + i * 0.78;
    s.addText(name, { x: 0.86, y, w: 1.45, h: 0.28, margin: 0, fontFace: F, fontSize: 11.5, bold: true, color: INK });
    s.addText(desc, { x: 0.86, y: y + 0.26, w: 3.05, h: 0.44, margin: 0, valign: "top", fontFace: F, fontSize: 9.5, color: MUTED });
    const bw = 2.35;
    s.addShape("roundRect", { x: 4.05, y: y + 0.06, w: bw * share, h: 0.26, rectRadius: 0.04, fill: { color: AMBER }, line: { type: "none" } });
    if (share < 1) s.addShape("roundRect", { x: 4.05 + bw * share, y: y + 0.06, w: bw * (1 - share), h: 0.26, rectRadius: 0.04, fill: { color: R3 }, line: { type: "none" } });
    card(s, 6.55, y + 0.02, 1.10, 0.34, PAPER);
    s.addText(sig, { x: 6.55, y: y + 0.02, w: 1.10, h: 0.34, margin: 0, align: "center", valign: "middle", fontFace: F, fontSize: 9, bold: true, color: col });
  });
  s.addText("Cap precision = near-cap promo products ÷ active promo products, where near-cap is within ±5% of the cap.", {
    x: 0.86, y: 6.02, w: 6.8, h: 0.28, margin: 0, fontFace: F, fontSize: 9, italic: true, color: MUTED,
  });

  card(s, 8.18, 3.06, 4.73, 3.30);
  cardHead(s, 8.18, 3.06, 4.73, "Decision view", "Stickiness against rate sensitivity");
  const quad = [
    ["Core sticky", "Protect & grow", "E8F6F3", TITLE_TEAL],
    ["Sticky today", "Maturity risk", WARM, AMBER],
    ["Weak relation", "Build usage", PAPER, MUTED],
    ["Rate driven", "Runoff risk", WARM, RED],
  ];
  quad.forEach(([t, sub, fill, col], i) => {
    const qx = 9.32 + (i % 2) * 1.72, qy = 3.82 + Math.floor(i / 2) * 1.02;
    card(s, qx, qy, 1.62, 0.90, fill);
    s.addText(t, { x: qx, y: qy + 0.12, w: 1.62, h: 0.28, margin: 0, align: "center", fontFace: F, fontSize: 11, bold: true, color: col });
    s.addText(sub, { x: qx, y: qy + 0.42, w: 1.62, h: 0.26, margin: 0, align: "center", fontFace: F, fontSize: 9.5, color: MUTED });
  });
  s.addText("HIGH\nstickiness", { x: 8.42, y: 3.82, w: 0.85, h: 0.90, margin: 0, align: "center", valign: "middle", fontFace: F, fontSize: 9, bold: true, color: MUTED });
  s.addText("LOW\nstickiness", { x: 8.42, y: 4.84, w: 0.85, h: 0.90, margin: 0, align: "center", valign: "middle", fontFace: F, fontSize: 9, bold: true, color: MUTED });
  s.addText("Low sensitivity (<60)", { x: 9.32, y: 5.80, w: 1.62, h: 0.24, margin: 0, align: "center", fontFace: F, fontSize: 8.5, color: MUTED });
  s.addText("High sensitivity (≥60)", { x: 11.04, y: 5.80, w: 1.62, h: 0.24, margin: 0, align: "center", fontFace: F, fontSize: 8.5, color: MUTED });
  s.addText("Initial cut-offs are heuristic; maturity will recalibrate them.", {
    x: 8.46, y: 6.02, w: 4.17, h: 0.28, margin: 0, fontFace: F, fontSize: 8.5, italic: true, color: NOTE,
  });
  s.addNotes("Rate sensitivity is a psychographic proxy inferred from where customers park balance relative to the promotional caps. Promo dependency is an exposure overlay, not a driver of the score.");
}

/* ============================================================== SLIDE 3 */
{
  const s = pres.addSlide(); s.background = { color: PAPER };
  head(s, "Rate-sensitive balances sit in the promotional pockets",
    "*Balance at risk = customers scoring High or Very High rate sensitivity (≥60) · pre-maturity view.",
    `Save Max at risk THB ${m(smAtRisk)}   ·   Save More at risk THB 3.27bn   ·   Highest-runoff subset THB ${Math.round(smRunoff / 1e6)}m + THB 92m`,
    "Almost all rate-sensitive balance sits with customers who look sticky today — which is why maturity is the real test.");

  [
    ["Save Max 4% — maturity exposure", `THB ${m(smAtRisk)}`, `${pc(smAtRisk, smBal)} of pocket balance`, `Total balance THB ${m(smBal)}`,
      [["Very High RS", 374, R4], ["High RS", 104, AMBER], ["RS < 60", 207, R1]], 0.58],
    ["Save More 2% — promotional exposure", "THB 3.27bn", "78% of pocket balance", "Total balance THB 4.16bn",
      [["Very High RS", 2210, R4], ["High RS", 1060, AMBER], ["RS < 60", 890, R1]], 6.87],
  ].forEach(([title, big, sub, total, segs, x]) => {
    card(s, x, 1.80, 6.04, 1.98);
    cardHead(s, x, 1.80, 6.04, title, total);
    s.addText(big, { x: x + 0.28, y: 2.42, w: 2.6, h: 0.46, margin: 0, fontFace: F, fontSize: 24, bold: true, color: AMBER });
    s.addText(sub, { x: x + 2.95, y: 2.50, w: 2.8, h: 0.3, margin: 0, valign: "middle", fontFace: F, fontSize: 10, color: MUTED });
    const tot = segs.reduce((a, g) => a + g[1], 0);
    let bx = x + 0.28;
    segs.forEach(([lab, v, col]) => {
      const ww = (v / tot) * 5.48;
      s.addShape("roundRect", { x: bx, y: 2.98, w: ww, h: 0.26, rectRadius: 0.04, fill: { color: col }, line: { type: "none" } });
      bx += ww;
    });
    s.addText(segs.map((g) => g[0]).join("        "), { x: x + 0.28, y: 3.30, w: 5.48, h: 0.26, margin: 0, fontFace: F, fontSize: 8.5, color: MUTED });
  });

  card(s, 0.58, 3.96, 7.35, 2.42);
  cardHead(s, 0.58, 3.96, 7.35, "What is actually at risk?", "Most rate-sensitive balance still looks sticky today");
  [
    `Save Max: THB ${m(smSticky)} (${pc(smSticky, smAtRisk)}) of rate-sensitive balance sits with High / Very High stickiness customers.`,
    "Save More: THB 3.18bn (97%) of rate-sensitive balance sits with High / Very High stickiness customers.",
  ].forEach((t, i) => s.addText(t, { x: 0.86, y: 4.66 + i * 0.42, w: 6.8, h: 0.38, margin: 0, valign: "top", fontFace: F, fontSize: 11, color: INK }));
  card(s, 0.86, 5.56, 6.79, 0.62, WARM);
  s.addText("Current stickiness may be propped up by promotional pricing — maturity is the real retention test.", {
    x: 1.06, y: 5.56, w: 6.4, h: 0.62, margin: 0, valign: "middle", fontFace: F, fontSize: 11, bold: true, color: RED,
  });

  card(s, 8.18, 3.96, 4.73, 2.42, WARM);
  cardHead(s, 8.18, 3.96, 4.73, "Highest runoff-risk subset", "Stickiness < 60 and rate sensitivity ≥ 60");
  [[`THB ${Math.round(smRunoff / 1e6)}m`, "Save Max"], ["THB 92m", "Save More"]].forEach(([v, l], i) => {
    s.addText(v, { x: 8.46, y: 4.66 + i * 0.52, w: 1.8, h: 0.42, margin: 0, valign: "middle", fontFace: F, fontSize: 19, bold: true, color: RED });
    s.addText(l, { x: 10.3, y: 4.66 + i * 0.52, w: 2.3, h: 0.42, margin: 0, valign: "middle", fontFace: F, fontSize: 11, color: INK });
  });
  s.addText("Prioritise for maturity and repricing retention actions.", {
    x: 8.46, y: 5.78, w: 4.17, h: 0.3, margin: 0, fontFace: F, fontSize: 10.5, bold: true, color: TITLE_TEAL,
  });
  s.addNotes("At-risk is a heuristic definition before maturity. Validate against actual post-maturity runoff and recalibrate the cut-offs.");
}

/* ============================================================== SLIDE 4 */
{
  const s = pres.addSlide(); s.background = { color: PAPER };
  head(s, "The two scores, side by side",
    "*20-day engagement target is prorated by tenure. Rate-sensitivity cut-offs are initial heuristics, to be validated after maturity.",
    null,
    "Stickiness measures retention behaviour; rate sensitivity measures behaviour around the promotional rate caps.");

  [
    ["1)  Deposit stickiness", "Score = 40% BP + 25% FC + 20% RI + 15% DE",
      [["BP · Balance persistence", "Latest avg balance ÷ prior avg balance"],
       ["FC · Funding continuity", "% eligible days with balance ≥ THB100"],
       ["RI · Recurring inflow", "% eligible months with transfer-in ≥ THB100"],
       ["DE · Daily engagement", "Active txn days ÷ 20-day / 90-day target*"]],
      "Higher score = money stays longer, the account stays funded, fresh money keeps arriving, and the customer uses the account regularly.", 0.58],
    ["2)  Rate sensitivity", "Score = 70% Cap Precision + 30% (1 − Above-Cap Ratio)",
      [["Cap precision", "Near-cap promo products ÷ active promo products"],
       ["Above-cap balance ratio", "Balance above caps ÷ total promo-pocket balance"],
       ["Promo caps", "Save Max 4%: THB 20k · Save More 2%: THB 1m"],
       ["Near-cap", "Within ±5% of the promotional cap"]],
      "Higher score = the customer behaves like a rate optimiser, keeping balances close to the promo caps and little money above them.", 6.87],
  ].forEach(([title, formula, defs, layman, x]) => {
    card(s, x, 1.80, 6.04, 4.58);
    s.addText(title, { x: x + 0.28, y: 1.94, w: 5.48, h: 0.34, margin: 0, fontFace: FH, fontSize: 15, bold: true, color: TITLE_TEAL });
    card(s, x + 0.28, 2.34, 5.48, 0.46, "E8F6F3");
    s.addText(formula, { x: x + 0.28, y: 2.34, w: 5.48, h: 0.46, margin: 0, align: "center", valign: "middle", fontFace: F, fontSize: 11, bold: true, color: INK });
    rows(s, x, 2.96, 6.04, defs, { labelW: 2.15, pitch: 0.38, size: 10 });
    s.addShape("rect", { x: x + 0.28, y: 4.54, w: 5.48, h: 0.012, fill: { color: RULE }, line: { type: "none" } });
    s.addText("In plain terms", { x: x + 0.28, y: 4.66, w: 5.48, h: 0.28, margin: 0, fontFace: F, fontSize: 11.5, bold: true, color: TITLE_TEAL });
    s.addText(layman, { x: x + 0.28, y: 4.94, w: 5.48, h: 0.62, margin: 0, valign: "top", fontFace: F, fontSize: 10.5, color: INK });
    s.addText("Score cut-offs", { x: x + 0.28, y: 5.62, w: 5.48, h: 0.26, margin: 0, fontFace: F, fontSize: 11, bold: true, color: TITLE_TEAL });
    bandRow(s, x + 0.28, 5.96, 5.48, 8.5);
  });
  s.addNotes("Both scores run 0-100 and share the same five cut-off bands, so they can be crossed directly in the decision view.");
}

/* ============================================================== SLIDE 5 */
/* Two worked examples. Every score below is computed from the pocket
   balances shown, using the two formulas on the previous slide, so the
   examples stay internally consistent if the inputs are edited. */
{
  const s = pres.addSlide(); s.background = { color: PAPER };
  const CAP_MAX = 20e3, CAP_MORE = 1e6;
  const bandCol = (v) => (v >= 80 ? R4 : v >= 60 ? R3 : v >= 40 ? R2 : v >= 20 ? AMBER : R1);
  const bandName = (v) => (v >= 80 ? "Very High" : v >= 60 ? "High" : v >= 40 ? "Medium" : v >= 20 ? "Low" : "Very Low");
  const chip = (s2, x, y, w, h, text, fill, size) => {
    s2.addShape("roundRect", { x, y, w, h, rectRadius: 0.05, fill: { color: fill }, line: { type: "none" } });
    s2.addText(text, { x, y, w, h, margin: 0, align: "center", valign: "middle", fontFace: F, fontSize: size, bold: true, color: PAPER });
  };
  const thb = (v) => "THB " + (v >= 1e6 ? (v / 1e6).toFixed(1) + "m" : Math.round(v / 1e3) + "k");

  head(s, "Two customers, equally sticky today — very different repricing risk",
    "*Illustrative examples. Stickiness measures observed retention behaviour; rate sensitivity measures how balances sit around the promotional caps.",
    null,
    "Both examples score Very High stickiness — they differ only in where balance sits relative to the caps.");

  [
    {
      x: 0.58, title: "Example 1 · Sticky but rate-sensitive", flag: "Maturity risk", flagCol: AMBER,
      sub: "Balances sit exactly at the promotional caps.",
      pockets: [["Main Pocket", 50e3, INK], ["Save Max", 20e3, AMBER, true], ["Save More", 1.0e6, AMBER, true]],
      comp: [95, 100, 75, 50], nearCap: 2,
      fill: WARM,
      layman: "Layman view: the money has stayed so far, but the customer behaves like a rate optimiser — retest retention when the promo rate reprices.",
    },
    {
      x: 6.87, title: "Example 2 · Core sticky depositor", flag: "Lower risk", flagCol: R3,
      sub: "Holds meaningful balance above both promotional caps.",
      pockets: [["Main Pocket", 300e3, INK], ["Save Max", 100e3, R3, true], ["Save More", 1.5e6, R3, true]],
      comp: [90, 95, 100, 75], nearCap: 0,
      fill: "E8F6F3",
      layman: "Layman view: the money stays and the customer holds well above the high-rate limits — a sign of structural, not promo-driven, deposits.",
    },
  ].forEach((e) => {
    const X = e.x, IN = X + 0.28, IW = 5.48;
    card(s, X, 1.80, 6.04, 4.56);
    s.addText(e.title, { x: IN, y: 1.92, w: 4.05, h: 0.30, margin: 0, valign: "middle", fontFace: FH, fontSize: 14, bold: true, color: INK });
    chip(s, X + 4.41, 1.90, 1.35, 0.34, e.flag, e.flagCol, 9.5);
    s.addText(e.sub, { x: IN, y: 2.24, w: IW, h: 0.24, margin: 0, fontFace: F, fontSize: 9.5, color: MUTED });

    /* pocket balances */
    const tw = (IW - 0.32) / 3;
    /* the promotional pockets carry a holding flag: Y when the customer has
       balance in that pocket, N when they do not. Main Pocket is not
       promotional, so it has no flag. */
    e.pockets.forEach(([lab, v, col, promo], i) => {
      const tx = IN + i * (tw + 0.16);
      card(s, tx, 2.48, tw, 0.66, PAPER);
      s.addText(lab, { x: tx + 0.16, y: 2.55, w: tw - (promo ? 0.72 : 0.32), h: 0.20, margin: 0, fontFace: F, fontSize: 9, bold: true, color: MUTED });
      if (promo) chip(s, tx + tw - 0.68, 2.545, 0.52, 0.21, "FLAG " + (v > 0 ? "Y" : "N"), v > 0 ? R3 : R1, 7);
      s.addText(thb(v), { x: tx + 0.16, y: 2.75, w: tw - 0.32, h: 0.32, margin: 0, valign: "middle", fontFace: F, fontSize: 15, bold: true, color: col });
    });

    /* the four stickiness components */
    s.addText("Deposit stickiness components", { x: IN, y: 3.24, w: IW, h: 0.26, margin: 0, fontFace: F, fontSize: 11, bold: true, color: TITLE_TEAL });
    const COMP = [
      ["Balance persistence", "latest ÷ prior avg. balance", 0.40],
      ["Funding continuity", "% eligible days funded", 0.25],
      ["Recurring inflow", "% months with inflow", 0.20],
      ["Daily engagement", "active days vs target", 0.15],
    ];
    COMP.forEach(([name, desc, w], i) => {
      const y = 3.52 + i * 0.32, v = e.comp[i];
      s.addText(name, { x: IN, y, w: 1.70, h: 0.28, margin: 0, valign: "middle", fontFace: F, fontSize: 10, bold: true, color: INK });
      s.addText(`${desc}   ·   ${Math.round(w * 100)}%`, { x: X + 2.02, y, w: 2.90, h: 0.28, margin: 0, valign: "middle", fontFace: F, fontSize: 9, color: MUTED });
      chip(s, X + 4.96, y + 0.01, 0.80, 0.26, String(v), bandCol(v), 9.5);
    });

    s.addShape("rect", { x: IN, y: 4.86, w: IW, h: 0.012, fill: { color: RULE }, line: { type: "none" } });

    /* the two scores, computed from the inputs above */
    const st = COMP.reduce((a, c, i) => a + c[2] * e.comp[i], 0);
    const promoBal = e.pockets[1][1] + e.pockets[2][1];
    const aboveCap = Math.max(e.pockets[1][1] - CAP_MAX, 0) + Math.max(e.pockets[2][1] - CAP_MORE, 0);
    const aboveRatio = aboveCap / promoBal;
    const rs = 0.7 * (e.nearCap / 2) * 100 + 0.3 * (1 - aboveRatio) * 100;
    [
      ["Stickiness score", "40/25/20/15 weighted", st, bandCol(st)],
      ["Rate sensitivity", `${e.nearCap}/2 near cap  ·  ${(aboveRatio * 100).toFixed(0)}% above cap`, rs, rs >= 60 ? RED : R3],
    ].forEach(([lab, detail, v, col], i) => {
      const y = 4.98 + i * 0.36;
      s.addText(lab, { x: IN, y, w: 1.70, h: 0.30, margin: 0, valign: "middle", fontFace: F, fontSize: 10.5, bold: true, color: INK });
      s.addText(detail, { x: X + 2.02, y, w: 2.14, h: 0.30, margin: 0, valign: "middle", fontFace: F, fontSize: 9, color: MUTED });
      chip(s, X + 4.22, y + 0.01, 1.54, 0.28, `${Math.round(v)}  ·  ${bandName(v)}`, col, 9.5);
    });

    card(s, IN, 5.74, IW, 0.56, e.fill);
    s.addText(e.layman, { x: IN + 0.16, y: 5.74, w: IW - 0.32, h: 0.56, margin: 0, valign: "middle", fontFace: F, fontSize: 9.5, color: INK });
  });

  card(s, 0.58, 6.50, 12.33, 0.58, "E8F6F3");
  s.addText([
    { text: "Stickiness and rate sensitivity answer different questions", options: { bold: true, color: TITLE_TEAL } },
    { text: "  —  both examples score Very High stickiness, but only Example 2 keeps material balance above the caps. A customer can look sticky today and still be exposed when the promotion reprices.", options: { color: INK } },
  ], { x: 0.88, y: 6.50, w: 11.73, h: 0.58, margin: 0, valign: "middle", fontFace: F, fontSize: 11 });

  s.addNotes("Worked examples, illustrative. Example 1 parks exactly at both caps: 20k in Save Max and 1.0m in Save More, so cap precision is 2/2 and nothing sits above the caps - rate sensitivity 100, Very High. " +
    "Example 2 holds 100k in Save Max and 1.5m in Save More, so neither pocket is near its cap and 580k of the 1.6m promotional balance is above cap (36%) - rate sensitivity 19, Very Low. " +
    "Stickiness scores 86 and 91 respectively, both Very High, so the stickiness score alone would not separate these two customers. " +
    "This is the argument for keeping rate sensitivity as a second, independent lens rather than folding it into the stickiness score.");
}

/* ============================================================== SLIDE 6 */
{
  const s = pres.addSlide(); s.background = { color: PAPER };
  head(s, `Save Max : ${k(smCust)} customers, THB ${m(smBal)} — THB ${m(smAtRisk)} rate-sensitive`,
    "*Pre-maturity view · customers with Save Max balance > 0, by deposit stickiness × rate sensitivity.",
    `Customers ${k(smCust)}   ·   High rate sensitivity ${k(smHiRS)}   ·   Balance THB ${m(smBal)}   ·   At risk THB ${m(smAtRisk)}`,
    "Balances concentrate in the High and Very High stickiness tiers, yet most of that balance is also highly rate-sensitive.");

  tiles(s, [
    [k(smCust), "Save Max customers", "distinct ccd_id, balance > 0", TEAL],
    [k(smHiRS), "high-rate-sensitivity", `${pc(smHiRS, smCust)} of the cohort`, TEAL],
    [`THB ${m(smBal)}`, "total Save Max balance", "current pocket balance", TEAL],
    [`THB ${m(smAtRisk)}`, "balance at risk", `${pc(smAtRisk, smBal)} · High + Very High RS`, RED],
  ], 1.80);

  card(s, 0.58, 3.06, 6.04, 2.54);
  cardHead(s, 0.58, 3.06, 6.04, "Customer count", "Rows: stickiness · columns: rate sensitivity");
  matrix(s, 0.82, 3.72, 5.56, "savemax_count", (v) => v.toLocaleString("en-US"));

  card(s, 6.87, 3.06, 6.04, 2.54);
  cardHead(s, 6.87, 3.06, 6.04, "Balance (THB m)", "Rows: stickiness · columns: rate sensitivity");
  matrix(s, 7.11, 3.72, 5.56, "savemax_balance", (v) => (v / 1e6).toFixed(1));

  card(s, 0.58, 5.78, 12.33, 0.86, WARM);
  s.addText([
    { text: `THB ${m(smSticky)} (${pc(smSticky, smAtRisk)}) of the at-risk balance sits with High / Very High stickiness customers`, options: { bold: true, color: RED } },
    { text: `  —  so today's stickiness may be promotional. The genuinely fragile subset is stickiness < 60 with rate sensitivity ≥ 60: THB ${Math.round(smRunoff / 1e6)}m. Monitor it for early runoff signals.`, options: { color: INK } },
  ], { x: 0.88, y: 5.78, w: 11.73, h: 0.86, margin: 0, valign: "middle", fontFace: F, fontSize: 11 });
  s.addNotes(`Cohort totals: ${smCust.toLocaleString()} customers and THB ${smBal.toLocaleString()} of balance; ${smAtRisk.toLocaleString()} at risk. Amber column headers mark the High and Very High rate-sensitivity segments. Balance matrix is shown in THB millions.`);
}

/* ============================================================== SLIDE 7 */
{
  const s = pres.addSlide(); s.background = { color: PAPER };
  const balCust = gTot("all_bal_count");
  const balLowStick = M.all_bal_count.filter((r) => r.s === "Low" || r.s === "Very Low")
    .reduce((a, r) => a + r.v.reduce((x, y) => x + y, 0), 0);
  const balHiRS = gHiRS("all_bal_count");
  const vlAll = M.all_count.find((r) => r.s === "Very Low").v.reduce((a, b) => a + b, 0);
  const vlBal = M.all_bal_count.find((r) => r.s === "Very Low").v.reduce((a, b) => a + b, 0);
  const vhAll = M.all_count.find((r) => r.s === "Very High").v.reduce((a, b) => a + b, 0);
  const vhBal = M.all_bal_count.find((r) => r.s === "Very High").v.reduce((a, b) => a + b, 0);

  head(s, `All customers : ${(allCust / 1e6).toFixed(2)}m — only ${k(balCust)} hold a balance`,
    "*Full customer base · distinct ccd_id by deposit stickiness × rate sensitivity.",
    `Total ${(allCust / 1e6).toFixed(2)}m   ·   With balance > 0  ${k(balCust)} (${pc(balCust, allCust)})   ·   High rate sensitivity ${k(allHiRS)}`,
    `The funded base is far stickier than the account base — low stickiness falls from ${pc(allLowStick, allCust)} to ${pc(balLowStick, balCust)} once you require a balance.`);

  tiles(s, [
    [`${(allCust / 1e6).toFixed(2)}m`, "total customers", "distinct ccd_id", TEAL],
    [k(balCust), "customers with balance", `${pc(balCust, allCust)} of the base`, TEAL],
    [pc(balLowStick, balCust), "low stickiness, funded", `against ${pc(allLowStick, allCust)} of the full base`, AMBER],
    [k(balHiRS), "high rate sensitivity", "every one of them holds a balance", RED],
  ], 1.80);

  card(s, 0.58, 3.06, 6.04, 2.54);
  cardHead(s, 0.58, 3.06, 6.04, `All customers  ${(allCust / 1e6).toFixed(2)}m`, "Rows: stickiness · columns: rate sensitivity");
  matrix(s, 0.82, 3.72, 5.56, "all_count", (v) => v.toLocaleString("en-US"));

  card(s, 6.87, 3.06, 6.04, 2.54);
  cardHead(s, 6.87, 3.06, 6.04, `Customers with balance > 0  ${k(balCust)}`, "Rows: stickiness · columns: rate sensitivity");
  matrix(s, 7.11, 3.72, 5.56, "all_bal_count", (v) => v.toLocaleString("en-US"));

  card(s, 0.58, 5.78, 12.33, 0.86, "E8F6F3");
  s.addText([
    { text: `Very Low stickiness is largely an unfunded population — only ${pc(vlBal, vlAll)} of them hold a balance, against ${pc(vhBal, vhAll)} of Very High`, options: { bold: true, color: TITLE_TEAL } },
    { text: "  —  so use the funded view for maturity and runoff risk, and the full base for activation and funding strategy.", options: { color: INK } },
  ], { x: 0.88, y: 5.78, w: 11.73, h: 0.86, margin: 0, valign: "middle", fontFace: F, fontSize: 11 });
  s.addNotes(`Full base ${allCust.toLocaleString()}; ${balCust.toLocaleString()} hold a balance (${pc(balCust, allCust)}). ` +
    `Low + Very Low stickiness is ${allLowStick.toLocaleString()} of the full base (${pc(allLowStick, allCust)}) but only ${balLowStick.toLocaleString()} of the funded base (${pc(balLowStick, balCust)}). ` +
    `Funded share by tier: Very High ${pc(vhBal, vhAll)}, Very Low ${pc(vlBal, vlAll)}. ` +
    `High + Very High rate sensitivity is ${allHiRS.toLocaleString()} in both grids - every rate-sensitive customer holds a balance, which follows from the score being computed off balance placement. ` +
    "The balance>0 grid fills the TBD left open in the source deck.");
}

/* ============================================================== SLIDE 8 */
{
  const s = pres.addSlide(); s.background = { color: PAPER };
  const B = "all_bal_balance", C = "all_bal_count";
  const balTot = gTot(B), balCust = gTot(C);
  const rsBal = gHiRS(B), rsCust = gHiRS(C);
  const stickyBal = M[B].filter((r) => r.s === "High" || r.s === "Very High")
    .reduce((a, r) => a + r.v.reduce((x, y) => x + y, 0), 0);
  const rsSticky = M[B].filter((r) => r.s === "High" || r.s === "Very High")
    .reduce((a, r) => a + r.v[3] + r.v[4], 0);
  const fragile = rsBal - rsSticky;
  const rowTot = (mx, name) => M[mx].find((r) => r.s === name).v.reduce((a, b) => a + b, 0);
  const bn = (v) => "THB " + (v / 1e9).toFixed(2) + "bn";

  head(s, `Funded balance : ${bn(balTot)} — ${pc(rsBal, balTot)} of it is rate-sensitive`,
    "*Customers with balance > 0. Only 4 customers score Medium rate sensitivity, so that band is effectively empty — the score is close to binary.",
    `Total ${bn(balTot)}   ·   Rate-sensitive ${bn(rsBal)} (${pc(rsBal, balTot)})   ·   Save Max ${bn(smBal)} — ${pc(smBal, balTot)} of the book`,
    "Seven baht in ten across the deposit book is rate-sensitive — the exposure runs far wider than Save Max.");

  tiles(s, [
    [bn(balTot), "total deposit balance", `${k(balCust)} funded customers`, TEAL],
    [bn(rsBal), "rate-sensitive balance", `${pc(rsBal, balTot)} of the book · RS ≥ 60`, RED],
    [pc(stickyBal, balTot), "of balance in sticky tiers", "High + Very High stickiness", TEAL],
    [`THB ${Math.round(fragile / 1e6)}m`, "fragile subset", "stickiness < 60 and RS ≥ 60", RED],
  ], 1.80);

  card(s, 0.58, 3.06, 7.35, 2.54);
  cardHead(s, 0.58, 3.06, 7.35, "Balance (THB m)", "Rows: stickiness · columns: rate sensitivity");
  matrix(s, 0.82, 3.72, 6.87, B, (v) => (v / 1e6).toLocaleString("en-US", { minimumFractionDigits: 1, maximumFractionDigits: 1 }));

  card(s, 8.18, 3.06, 4.73, 2.54);
  cardHead(s, 8.18, 3.06, 4.73, "Average balance per customer", "Funded customers, by stickiness tier");
  const tiers = ["Very High", "High", "Medium", "Low", "Very Low"];
  const avgs = tiers.map((t) => [t, rowTot(B, t) / rowTot(C, t)]);
  const maxAvg = Math.max(...avgs.map((a) => a[1]));
  const tierCol = { "Very High": R4, High: R3, Medium: R2, Low: AMBER, "Very Low": R1 };
  avgs.forEach(([t, v], i) => {
    const y = 3.80 + i * 0.34;
    s.addText(t, { x: 8.46, y, w: 1.15, h: 0.28, margin: 0, valign: "middle", fontFace: F, fontSize: 10.5, bold: true, color: INK });
    s.addShape("roundRect", { x: 9.68, y: y + 0.05, w: Math.max((v / maxAvg) * 1.55, 0.04), h: 0.18, rectRadius: 0.03, fill: { color: tierCol[t] }, line: { type: "none" } });
    s.addText("THB " + Math.round(v).toLocaleString("en-US"), {
      x: 11.32, y, w: 1.31, h: 0.28, margin: 0, align: "right", valign: "middle",
      fontFace: F, fontSize: 10.5, bold: true, color: INK,
    });
  });
  s.addText("Very High stickiness holds ~346× the balance of Very Low.", {
    x: 8.46, y: 5.56, w: 4.17, h: 0.3, margin: 0, fontFace: F, fontSize: 9, italic: true, color: MUTED,
  });

  card(s, 0.58, 5.78, 12.33, 0.86, WARM);
  s.addText([
    { text: `${pc(rsSticky, rsBal)} of the rate-sensitive balance sits with High / Very High stickiness customers`, options: { bold: true, color: RED } },
    { text: `  —  the same pattern as Save Max, at eight times the size. The genuinely fragile slice is THB ${Math.round(fragile / 1e6)}m: customers who are neither sticky nor indifferent to price.`, options: { color: INK } },
  ], { x: 0.88, y: 5.78, w: 11.73, h: 0.86, margin: 0, valign: "middle", fontFace: F, fontSize: 11 });

  s.addNotes(`Total funded balance THB ${balTot.toLocaleString()} across ${balCust.toLocaleString()} customers, average THB ${Math.round(balTot / balCust).toLocaleString()}. ` +
    `Rate-sensitive balance THB ${rsBal.toLocaleString()} (${pc(rsBal, balTot)}) held by only ${rsCust.toLocaleString()} customers (${pc(rsCust, balCust)} of the funded base). ` +
    `Save Max is THB ${smBal.toLocaleString()}, just ${pc(smBal, balTot)} of the book, so rate-sensitivity exposure extends well beyond the maturing product. ` +
    `Balance concentrates hard: High and Very High stickiness hold ${pc(stickyBal, balTot)} of balance on ${pc(rowTot(C, "High") + rowTot(C, "Very High"), balCust)} of funded customers. ` +
    "Only 4 customers land in the Medium rate-sensitivity band, so the score is effectively binary - worth revisiting when the cut-offs are recalibrated after maturity.");
}

/* ============================================================== SLIDE 9 */
{
  const s = pres.addSlide(); s.background = { color: PAPER };
  const B = "all_bal_balance", C = "all_bal_count";
  const sum = (a) => a.reduce((x, y) => x + y, 0);
  const low = (mx) => M[mx].filter((r) => r.s === "Low" || r.s === "Very Low").reduce((a, r) => a + sum(r.v), 0);
  const sub60 = (mx) => M[mx].filter((r) => ["Medium", "Low", "Very Low"].includes(r.s)).reduce((a, r) => a + sum(r.v), 0);
  const sub60Hi = (mx) => M[mx].filter((r) => ["Medium", "Low", "Very Low"].includes(r.s)).reduce((a, r) => a + r.v[3] + r.v[4], 0);

  const cTot = gTot(C), bTot = gTot(B);
  const lowC = low(C), lowB = low(B);
  const rsC = gHiRS(C), rsB = gHiRS(B);
  const fragC = sub60Hi(C), fragB = sub60Hi(B);
  const rsSticky = rsB - fragB;
  const bn = (v) => "THB " + (v / 1e9).toFixed(2) + "bn";
  const p1 = (a, b) => ((a / b) * 100).toFixed(1) + "%";

  head(s, "Two risk lenses : low stickiness and high rate sensitivity",
    "*Funded base only (balance > 0). Low stickiness = score < 40 (Low + Very Low bands) · high rate sensitivity = score ≥ 60 (High + Very High bands).",
    `Funded ${k(cTot)} · ${bn(bTot)}   ·   Low stickiness ${k(lowC)} → THB ${mm(lowB)}   ·   High rate sensitivity ${k(rsC)} → ${bn(rsB)}`,
    "Low-stickiness customers are numerous but barely funded; rate-sensitive customers are few but hold seven baht in ten.");

  tiles(s, [
    [k(lowC), "low-stickiness customers", `${p1(lowC, cTot)} of the funded base`, AMBER],
    [`THB ${mm(lowB)}`, "balance they hold", `${p1(lowB, bTot)} of the deposit book`, AMBER],
    [k(rsC), "rate-sensitive customers", `${p1(rsC, cTot)} of the funded base · RS ≥ 60`, RED],
    [bn(rsB), "balance at risk", `${pc(rsB, bTot)} of the deposit book`, RED],
  ], 1.80);

  /* left: share of customers vs share of balance, for each lens ---------- */
  card(s, 0.58, 3.06, 7.35, 2.54);
  cardHead(s, 0.58, 3.06, 7.35, "Share of customers against share of balance",
    "Each bar is a percentage of the funded base · full track = 100%");
  const TRACK_X = 2.92, TRACK_W = 3.45, VAL_X = 6.48;
  const drawBar = (y, label, share, col) => {
    s.addText(label, { x: 0.86, y, w: 2.0, h: 0.28, margin: 0, valign: "middle", fontFace: F, fontSize: 10, color: INK });
    s.addShape("roundRect", { x: TRACK_X, y: y + 0.04, w: TRACK_W, h: 0.20, rectRadius: 0.03, fill: { color: "E4E9F0" }, line: { type: "none" } });
    s.addShape("roundRect", { x: TRACK_X, y: y + 0.04, w: Math.max(TRACK_W * share, 0.05), h: 0.20, rectRadius: 0.03, fill: { color: col }, line: { type: "none" } });
    s.addText((share * 100).toFixed(1) + "%", {
      x: VAL_X, y, w: 1.19, h: 0.28, margin: 0, align: "right", valign: "middle",
      fontFace: F, fontSize: 10.5, bold: true, color: col,
    });
  };
  [
    ["Low stickiness  ·  score < 40", AMBER, [["Share of customers", lowC / cTot], ["Share of balance", lowB / bTot]]],
    ["High rate sensitivity  ·  score ≥ 60", RED, [["Share of customers", rsC / cTot], ["Share of balance", rsB / bTot]]],
  ].forEach(([grp, col, bars], gi) => {
    const gy = 3.72 + gi * 0.94;
    s.addText(grp, { x: 0.86, y: gy, w: 6.8, h: 0.24, margin: 0, fontFace: F, fontSize: 10.5, bold: true, color: col });
    bars.forEach(([lab, share], i) => drawBar(gy + 0.26 + i * 0.32, lab, share, col));
  });

  /* right: the overlap of the two lenses --------------------------------- */
  card(s, 8.18, 3.06, 4.73, 2.54);
  cardHead(s, 8.18, 3.06, 4.73, "Where the two lenses overlap", "Stickiness < 60 and rate sensitivity ≥ 60");
  [[fragC.toLocaleString("en-US"), "customers"], [`THB ${Math.round(fragB / 1e6)}m`, "balance at risk"]].forEach(([v, l], i) => {
    const y = 3.76 + i * 0.50;
    s.addText(v, { x: 8.46, y, w: 2.20, h: 0.44, margin: 0, valign: "middle", fontFace: F, fontSize: 22, bold: true, color: RED });
    s.addText(l, { x: 10.72, y, w: 1.91, h: 0.44, margin: 0, valign: "middle", fontFace: F, fontSize: 11, color: INK });
  });
  card(s, 8.46, 4.80, 4.17, 0.72, WARM);
  s.addText(`The other ${pc(rsSticky, rsB)} of rate-sensitive balance — ${bn(rsSticky)} — sits with High / Very High stickiness customers.`, {
    x: 8.62, y: 4.80, w: 3.85, h: 0.72, margin: 0, valign: "middle", fontFace: F, fontSize: 9.5, color: INK,
  });

  card(s, 0.58, 5.78, 12.33, 0.86, WARM);
  s.addText([
    { text: "The balance at risk is not with the disengaged — it is with the rate-optimisers", options: { bold: true, color: RED } },
    { text: `  —  ${k(lowC)} low-stickiness customers hold only THB ${mm(lowB)}, while ${k(rsC)} rate-sensitive customers hold ${bn(rsB)}. Retention spend belongs with the second group; activation and funding with the first.`, options: { color: INK } },
  ], { x: 0.88, y: 5.78, w: 11.73, h: 0.86, margin: 0, valign: "middle", fontFace: F, fontSize: 11 });

  s.addNotes(`Funded base ${cTot.toLocaleString()} customers holding THB ${bTot.toLocaleString()}. ` +
    `Low stickiness (Low + Very Low) is ${lowC.toLocaleString()} customers, ${p1(lowC, cTot)} of the funded base, but only THB ${Math.round(lowB).toLocaleString()} - ${p1(lowB, bTot)} of the book. ` +
    `High rate sensitivity (High + Very High) is ${rsC.toLocaleString()} customers, ${p1(rsC, cTot)}, holding THB ${Math.round(rsB).toLocaleString()} - ${pc(rsB, bTot)} of the book. ` +
    `Widening the stickiness lens to score < 60 adds the Medium band: ${sub60(C).toLocaleString()} customers (${p1(sub60(C), cTot)}) and THB ${Math.round(sub60(B)).toLocaleString()} (${p1(sub60(B), bTot)}). ` +
    `The overlap of the two lenses - stickiness < 60 and rate sensitivity >= 60 - is ${fragC.toLocaleString()} customers and THB ${Math.round(fragB).toLocaleString()}, the priority watchlist. ` +
    `The remaining ${pc(rsSticky, rsB)} of rate-sensitive balance sits with customers who currently score High or Very High stickiness, which is exactly the stickiness maturity will test.`);
}

/* ===================================================== SLIDES 10 and 11 */
/* Post-maturity validation. Rows are the Save Max cohort crossed by rate
   sensitivity, Save More holding and whether the customer emptied Save Max.
   Every figure below is aggregated from data/savemax_maturity.tsv. */
{
  const MAT = fs.readFileSync(path.join(__dirname, "data", "savemax_maturity.tsv"), "utf8")
    .trim().split("\n").slice(1).map((l) => {
      const c = l.split("\t");
      return { rs: c[0], sm: +c[1], mv: c[2], n: +c[3], p: +c[4], l: +c[5], ps: +c[6], ls: +c[7] };
    });
  const S = (f, g = () => true) => MAT.filter(g).reduce((a, r) => a + f(r), 0);
  const moved = (r) => r.mv === "Move";
  const cell = (mv, sm) => (r) => (r.mv === "Move") === mv && r.sm === sm;

  const nAll = S((r) => r.n);
  const pAll = S((r) => r.p), lAll = S((r) => r.l);
  const psAll = S((r) => r.ps), lsAll = S((r) => r.ls);
  const nMv = S((r) => r.n, moved);
  const bn2 = (v) => "THB " + (v / 1e9).toFixed(2) + "bn";
  const mR = (v) => "THB " + Math.round(v / 1e6).toLocaleString("en-US") + "m";
  const sg = (v) => (v < 0 ? "−" : "+") + "THB " + Math.round(Math.abs(v) / 1e6) + "m";
  const sp = (v) => (v < 0 ? "−" : "+") + Math.abs(v * 100).toFixed(1) + "%";
  const n0 = (v) => Math.round(v).toLocaleString("en-US");

  /* horizontal bar row: label | track | value */
  function bar(s, y, lx, lw, tx, tw, vx, vw, label, frac, col, value, opt = {}) {
    s.addText(label, { x: lx, y, w: lw, h: 0.28, margin: 0, valign: "middle", fontFace: F, fontSize: opt.size || 10, bold: !!opt.bold, color: opt.labelCol || INK });
    s.addShape("roundRect", { x: tx, y: y + 0.05, w: tw, h: 0.18, rectRadius: 0.03, fill: { color: "E4E9F0" }, line: { type: "none" } });
    s.addShape("roundRect", { x: tx, y: y + 0.05, w: Math.max(tw * frac, 0.04), h: 0.18, rectRadius: 0.03, fill: { color: col }, line: { type: "none" } });
    s.addText(value, { x: vx, y, w: vw, h: 0.28, margin: 0, align: "right", valign: "middle", fontFace: F, fontSize: opt.size || 10, bold: true, color: col });
  }

  /* ---------------------------------------------------------- SLIDE 10 */
  {
    const s = pres.addSlide(); s.background = { color: PAPER };
    head(s, `Maturity outcome : Save Max ${sg(lsAll - psAll)}, total balance ${sg(lAll - pAll)}`,
      "*Save Max holders with balance > 0. 'Move' = Save Max now zero. prev / latest balance is the customer's total deposit balance, not Save Max alone.",
      `Customers ${n0(nAll)}  ·  Moved ${n0(nMv)} (${(nMv / nAll * 100).toFixed(1)}%)  ·  Save Max ${mR(psAll)} → ${Math.round(lsAll / 1e6)}m  ·  Total ${bn2(pAll)} → ${(lAll / 1e9).toFixed(2)}bn`,
      `The money left the pocket, not the bank — Save Max fell ${Math.abs((lsAll / psAll - 1) * 100).toFixed(0)}% while total deposit balance rose ${((lAll / pAll - 1) * 100).toFixed(0)}%.`);

    tiles(s, [
      [k(nAll), "customers in cohort", "held Save Max last period", TEAL],
      [n0(nMv), "moved out of Save Max", `${(nMv / nAll * 100).toFixed(1)}% · Save Max now zero`, AMBER],
      [sg(lsAll - psAll), "Save Max balance", `${sp(lsAll / psAll - 1)} period on period`, RED],
      [sg(lAll - pAll), "total deposit balance", `${sp(lAll / pAll - 1)} period on period`, TEAL],
    ], 1.80);

    /* left: the two balance measures, prior vs latest, on one scale */
    card(s, 0.58, 3.06, 6.04, 2.54);
    cardHead(s, 0.58, 3.06, 6.04, "Prior period against latest", "Both measures on one scale — Save Max is a pocket inside the total");
    const scale = Math.max(pAll, lAll);
    [
      ["Save Max pocket", [["Prior", psAll, R1], ["Latest", lsAll, RED]]],
      ["Total deposit balance", [["Prior", pAll, R1], ["Latest", lAll, R3]]],
    ].forEach(([grp, bars], gi) => {
      const gy = 3.72 + gi * 0.92;
      s.addText(grp, { x: 0.86, y: gy, w: 5.48, h: 0.22, margin: 0, fontFace: F, fontSize: 10.5, bold: true, color: TITLE_TEAL });
      bars.forEach(([lab, v, col], i) =>
        bar(s, gy + 0.24 + i * 0.32, 0.86, 1.35, 2.33, 2.55, 4.98, 1.36, lab, v / scale, col, mR(v)));
    });

    /* right: move rate by rate-sensitivity segment */
    card(s, 6.87, 3.06, 6.04, 2.54);
    cardHead(s, 6.87, 3.06, 6.04, "Who emptied Save Max", "Share of each rate-sensitivity segment that moved");
    const segs = ["Very High", "High", "Medium", "Low", "Very Low"];
    const segCol = { "Very High": R4, High: RED, Medium: R2, Low: AMBER, "Very Low": R1 };
    segs.forEach((rs, i) => {
      const tot = S((r) => r.n, (r) => r.rs === rs), mvd = S((r) => r.n, (r) => r.rs === rs && moved(r));
      const y = 3.76 + i * 0.34;
      bar(s, y, 7.15, 1.15, 8.40, 2.00, 10.50, 0.85, rs, mvd / tot, segCol[rs],
        (mvd / tot * 100).toFixed(1) + "%", { bold: true });
      s.addText(`${n0(mvd)} of ${n0(tot)}`, { x: 11.45, y, w: 1.40, h: 0.28, margin: 0, align: "right", valign: "middle", fontFace: F, fontSize: 9, color: MUTED });
    });
    s.addText("High is the mover band — Very High moves at barely a quarter of that rate.", {
      x: 7.15, y: 5.42, w: 5.48, h: 0.18, margin: 0, fontFace: F, fontSize: 8.5, italic: true, color: MUTED,
    });

    card(s, 0.58, 5.78, 12.33, 0.86, "E8F6F3");
    s.addText([
      { text: `Save Max lost ${mR(psAll - lsAll)} but the book gained ${mR(lAll - pAll)}`, options: { bold: true, color: TITLE_TEAL } },
      { text: `  —  the ${n0(nMv)} customers who emptied Save Max took ${mR(S((r) => r.ps, moved))} out of the pocket, yet total deposits rose. Runoff from the promotional pocket is not the same as runoff from the bank.`, options: { color: INK } },
    ], { x: 0.88, y: 5.78, w: 11.73, h: 0.86, margin: 0, valign: "middle", fontFace: F, fontSize: 11 });

    s.addNotes(`Cohort: ${n0(nAll)} customers holding Save Max in the prior period. ${n0(nMv)} of them (${(nMv / nAll * 100).toFixed(1)}%) now show zero Save Max balance. ` +
      `Save Max fell from THB ${n0(psAll)} to THB ${n0(lsAll)}, down ${Math.abs((lsAll / psAll - 1) * 100).toFixed(1)}%, of which THB ${n0(S((r) => r.ps, moved))} came from movers zeroing out and THB ${n0(S((r) => r.ps, (r) => !moved(r)) - lsAll)} from non-movers trimming. ` +
      `Total deposit balance nevertheless rose from THB ${n0(pAll)} to THB ${n0(lAll)}, up ${((lAll / pAll - 1) * 100).toFixed(1)}%. Save Max fell from ${(psAll / pAll * 100).toFixed(1)}% to ${(lsAll / lAll * 100).toFixed(1)}% of the book. ` +
      "Move rate is not monotonic in rate sensitivity: the High band moves at 49.5% against 12.8% for Very High, so the current cut-offs do not rank runoff risk correctly and should be recalibrated on this outcome.");
  }

  /* ---------------------------------------------------------- SLIDE 11 */
  {
    const s = pres.addSlide(); s.background = { color: PAPER };
    const grp = (mv, sm) => {
      const g = cell(mv, sm);
      const n = S((r) => r.n, g), p = S((r) => r.p, g), l = S((r) => r.l, g);
      return { n, p, l, ret: l / p };
    };
    const mY = grp(true, 1), mN = grp(true, 0), sY = grp(false, 1), sN = grp(false, 0);

    head(s, "The Save More flag decides whether the money stays",
      "*Balance retained = latest total deposit balance ÷ prior total deposit balance, for the same customers.",
      `Movers ${n0(nMv)}   ·   With Save More ${n0(mY.n)} → ${sg(mY.l - mY.p)} (${sp(mY.ret - 1)})   ·   Without Save More ${n0(mN.n)} → ${sg(mN.l - mN.p)} (${sp(mN.ret - 1)})`,
      "Among customers who emptied Save Max, holding Save More is the difference between growth and near-total runoff.");

    tiles(s, [
      [n0(mY.n), "movers holding Save More", `${(mY.n / nMv * 100).toFixed(0)}% of all movers`, TEAL],
      [(mY.ret * 100).toFixed(0) + "%", "balance retained", `${mR(mY.p)} → ${bn2(mY.l)}`, TEAL],
      [n0(mN.n), "movers without Save More", `${(mN.n / nMv * 100).toFixed(0)}% of all movers`, RED],
      [(mN.ret * 100).toFixed(0) + "%", "balance retained", `${mR(mN.p)} → ${mR(mN.l)}`, RED],
    ], 1.80);

    /* left: retention for all four groups, with a 100% reference line */
    card(s, 0.58, 3.06, 6.04, 2.54);
    cardHead(s, 0.58, 3.06, 6.04, "Balance retained after maturity", "Latest ÷ prior total deposit balance · the rule marks 100%");
    const TX = 2.98, TW = 2.15, SCALE = 1.10;
    s.addShape("rect", { x: TX + TW * (1 / SCALE), y: 3.74, w: 0.012, h: 1.64, fill: { color: MUTED }, line: { type: "none" } });
    [
      ["Moved  ·  holds Save More", mY, R3],
      ["Moved  ·  no Save More", mN, RED],
      ["Stayed  ·  holds Save More", sY, R3],
      ["Stayed  ·  no Save More", sN, R1],
    ].forEach(([lab, g, col], i) =>
      bar(s, 3.80 + i * 0.42, 0.86, 2.00, TX, TW, 5.23, 1.11, lab, g.ret / SCALE, col,
        (g.ret * 100).toFixed(1) + "%", { size: 10 }));
    s.addText("100%", { x: TX + TW / SCALE - 0.30, y: 5.40, w: 0.60, h: 0.18, margin: 0, align: "center", fontFace: F, fontSize: 8, color: MUTED });

    /* right: the full 2x2 */
    card(s, 6.87, 3.06, 6.04, 2.54);
    cardHead(s, 6.87, 3.06, 6.04, "Customers and balance change", "Rows: Save Max behaviour · columns: Save More holding");
    const th = (t) => ({ text: t, options: { fill: { color: TITLE_TEAL }, color: PAPER, bold: true, fontFace: F, fontSize: 9, align: "center" } });
    const cl = (g, bad) => ({
      text: `${n0(g.n)}\n${sp(g.ret - 1)}   ${sg(g.l - g.p)}`,
      options: { fill: { color: bad ? AT_RISK_TINT : PAPER }, color: bad ? RED : INK, bold: bad, fontFace: F, fontSize: 9.5, align: "center" },
    });
    const rh = (t) => ({ text: t, options: { fill: { color: CARD }, color: INK, bold: true, fontFace: F, fontSize: 9.5, align: "left" } });
    s.addTable([
      [{ text: "Save Max", options: { fill: { color: TITLE_TEAL }, color: PAPER, bold: true, fontFace: F, fontSize: 9, align: "left" } }, th("Holds Save More"), th("No Save More")],
      [rh("Moved out"), cl(mY, false), cl(mN, true)],
      [rh("Stayed"), cl(sY, false), cl(sN, false)],
    ], {
      x: 7.11, y: 3.76, w: 5.56, colW: [1.60, 1.98, 1.98], rowH: 0.40,
      border: { type: "solid", color: RULE, pt: 0.5 }, valign: "middle", margin: [0.04, 0.08, 0.04, 0.08],
    });
    s.addText(`Save More holders move ${(S((r) => r.n, (r) => r.sm === 1 && moved(r)) / S((r) => r.n, (r) => r.sm === 1) * 100).toFixed(1)}% of the time against ${(S((r) => r.n, (r) => r.sm === 0 && moved(r)) / S((r) => r.n, (r) => r.sm === 0) * 100).toFixed(1)}% — they move far more often, but into the next pocket rather than out of the bank.`, {
      x: 7.15, y: 5.06, w: 5.48, h: 0.44, margin: 0, valign: "top", fontFace: F, fontSize: 9.5, color: INK,
    });

    card(s, 0.58, 5.78, 12.33, 0.86, WARM);
    s.addText([
      { text: `The ${n0(mN.n)} movers with no Save More kept only ${(mN.ret * 100).toFixed(0)}% of their balance`, options: { bold: true, color: RED } },
      { text: `  —  ${sg(mN.l - mN.p)} gone. The same behaviour with a Save More holding grew ${sp(mY.ret - 1)}. Cross-holding, not the stickiness score, is what converts a maturing promotion into retained deposits.`, options: { color: INK } },
    ], { x: 0.88, y: 5.78, w: 11.73, h: 0.86, margin: 0, valign: "middle", fontFace: F, fontSize: 11 });

    s.addNotes(`Movers split sharply on the Save More flag. With Save More: ${n0(mY.n)} customers, THB ${n0(mY.p)} → THB ${n0(mY.l)}, ${(mY.ret * 100).toFixed(1)}% retained. ` +
      `Without: ${n0(mN.n)} customers, THB ${n0(mN.p)} → THB ${n0(mN.l)}, only ${(mN.ret * 100).toFixed(1)}% retained. ` +
      `Non-movers are stable either way: ${(sY.ret * 100).toFixed(1)}% with Save More and ${(sN.ret * 100).toFixed(1)}% without. ` +
      "Caveat on causality: Save More holders are also much larger balances, so part of the gap is cohort composition rather than the product itself. " +
      "The actionable read is still that a second promotional pocket gives maturing Save Max money somewhere to go inside the bank - worth testing as a deliberate retention offer before the next maturity wave.");
  }
}

/* ============================================================= SLIDE 12 */
/* Move rate across both scores. Rows are deposit stickiness, columns are
   rate sensitivity, so the grid matches the matrices earlier in the deck.
   Built from data/savemax_maturity_stickiness.tsv. */
{
  const s = pres.addSlide(); s.background = { color: PAPER };
  const D = fs.readFileSync(path.join(__dirname, "data", "savemax_maturity_stickiness.tsv"), "utf8")
    .trim().split("\n").slice(1).map((l) => {
      const c = l.split("\t");
      return { rs: c[0], st: c[1], sm: +c[2], mv: c[3] === "Move", n: +c[4], ps: +c[7] };
    });
  const S = (f, g = () => true) => D.filter(g).reduce((a, r) => a + f(r), 0);
  const N = S((r) => r.n), MOV = S((r) => (r.mv ? r.n : 0)), baseR = MOV / N;
  const ST = ["Very High", "High", "Medium", "Low", "Very Low"];
  const rate = (g) => { const n = S((r) => r.n, g); return n ? { n, r: S((x) => (x.mv ? x.n : 0), g) / n } : null; };
  const hiB = (x) => x === "Very High" || x === "High";
  const walked = S((r) => (r.mv ? r.ps : 0));
  const walkedSticky = S((r) => (r.mv ? r.ps : 0), (r) => hiB(r.st));
  const fragN = S((r) => r.n, (r) => !hiB(r.st) && hiB(r.rs));
  const fragB = S((r) => r.ps, (r) => !hiB(r.st) && hiB(r.rs));
  const lowN = S((r) => r.n, (r) => !hiB(r.st)), lowB = S((r) => r.ps, (r) => !hiB(r.st));
  const n0 = (v) => Math.round(v).toLocaleString("en-US");
  const best = { st: "High", rs: "High" };
  const bestC = rate((r) => r.st === best.st && r.rs === best.rs);

  head(s, "Move rate : deposit stickiness against rate sensitivity",
    "*Save Max holders with balance > 0. 'Move' = Save Max balance now zero. Each cell shows the move rate with its customer count beneath.",
    `Cohort ${n0(N)}   ·   Moved ${n0(MOV)} (${(baseR * 100).toFixed(1)}%)   ·   Hottest cell: High stickiness × High sensitivity ${(bestC.r * 100).toFixed(1)}% on ${n0(bestC.n)} customers`,
    "Move rate climbs with both scores — the runoff sits in the sticky, rate-sensitive core, not in the low-stickiness corner.");

  tiles(s, [
    [(baseR * 100).toFixed(1) + "%", "cohort move rate", `${n0(MOV)} of ${n0(N)} emptied Save Max`, TEAL],
    [(bestC.r * 100).toFixed(1) + "%", "hottest cell", "High stickiness × High sensitivity", RED],
    [(walkedSticky / walked * 100).toFixed(1) + "%", "of walked balance", "came from stickiness ≥ 60", RED],
    [n0(fragN), "in the 'fragile' corner", "stickiness < 60 and RS ≥ 60", MUTED],
  ], 1.80);

  /* the grid ------------------------------------------------------------ */
  card(s, 0.58, 3.06, 7.35, 2.54);
  cardHead(s, 0.58, 3.06, 7.35, "Move rate by cell", "Rows: deposit stickiness · columns: rate sensitivity · shaded by move rate");
  const maxR = Math.max(...ST.flatMap((st) => RS.map((rs) => (rate((r) => r.st === st && r.rs === rs) || { r: 0 }).r)));
  const hdr = [
    { text: "Stickiness", options: { fill: { color: TITLE_TEAL }, color: PAPER, bold: true, fontFace: F, fontSize: 9, align: "left" } },
    ...RS.map((x) => ({ text: x, options: { fill: { color: TITLE_TEAL }, color: PAPER, bold: true, fontFace: F, fontSize: 9, align: "center" } })),
  ];
  const body = ST.map((st) => [
    { text: st, options: { fill: { color: CARD }, color: INK, bold: true, fontFace: F, fontSize: 9.5, align: "left" } },
    ...RS.map((rs) => {
      const c = rate((r) => r.st === st && r.rs === rs);
      if (!c) return { text: "—", options: { fill: { color: PAPER }, color: MUTED, fontFace: F, fontSize: 9, align: "center" } };
      if (c.n < 50) return {
        text: [
          { text: (c.r * 100).toFixed(1) + "%", options: { fontFace: F, fontSize: 9.5, color: MUTED } },
          { text: "\n" + n0(c.n), options: { fontFace: F, fontSize: 7, color: MUTED, breakLine: false } },
        ],
        options: { fill: { color: PAPER }, align: "center" },
      };
      const t = c.r / maxR;
      return {
        text: [
          { text: (c.r * 100).toFixed(1) + "%", options: { fontFace: F, fontSize: 9.5, bold: true, color: t > 0.55 ? PAPER : INK } },
          { text: "\n" + n0(c.n), options: { fontFace: F, fontSize: 7, color: t > 0.55 ? "F0D6D6" : MUTED, breakLine: false } },
        ],
        options: { fill: { color: blend(RED, t) }, align: "center" },
      };
    }),
  ]);
  s.addTable([hdr, ...body], {
    x: 0.82, y: 3.80, w: 6.87, colW: [1.27, 1.12, 1.12, 1.12, 1.12, 1.12],
    rowH: [0.26, 0.30, 0.30, 0.30, 0.30, 0.30],
    border: { type: "solid", color: RULE, pt: 0.5 }, valign: "middle", margin: [0.01, 0.04, 0.01, 0.04],
  });

  /* marginal: the new dimension on its own ------------------------------ */
  card(s, 8.18, 3.06, 4.73, 2.54);
  cardHead(s, 8.18, 3.06, 4.73, "Move rate by stickiness", "The new dimension, on its own");
  const stCol = { "Very High": R4, High: R3, Medium: R2, Low: AMBER, "Very Low": R1 };
  const marg = ST.map((st) => [st, rate((r) => r.st === st)]);
  const mMax = Math.max(...marg.map(([, c]) => c.r));
  marg.forEach(([st, c], i) => {
    const y = 3.76 + i * 0.30;
    s.addText(st, { x: 8.46, y, w: 1.15, h: 0.26, margin: 0, valign: "middle", fontFace: F, fontSize: 10, bold: true, color: INK });
    s.addShape("roundRect", { x: 9.66, y: y + 0.04, w: Math.max((c.r / mMax) * 1.55, 0.04), h: 0.18, rectRadius: 0.03, fill: { color: stCol[st] }, line: { type: "none" } });
    s.addText((c.r * 100).toFixed(1) + "%", { x: 11.26, y, w: 0.72, h: 0.26, margin: 0, align: "right", valign: "middle", fontFace: F, fontSize: 10, bold: true, color: stCol[st] });
    s.addText(n0(c.n), { x: 12.02, y, w: 0.61, h: 0.26, margin: 0, align: "right", valign: "middle", fontFace: F, fontSize: 8.5, color: MUTED });
  });
  card(s, 8.46, 5.26, 4.17, 0.30, WARM);
  s.addText(`Stickiness < 60: ${(lowN / N * 100).toFixed(1)}% of customers, ${(lowB / S((r) => r.ps) * 100).toFixed(1)}% of pocket.`, {
    x: 8.62, y: 5.26, w: 3.85, h: 0.30, margin: 0, valign: "middle", fontFace: F, fontSize: 9.5, color: INK,
  });

  card(s, 0.58, 5.78, 12.33, 0.86, WARM);
  s.addText([
    { text: `${(walkedSticky / walked * 100).toFixed(1)}% of the Save Max balance that walked came from customers scoring stickiness ≥ 60`, options: { bold: true, color: RED } },
    { text: `  —  the low-stickiness corner the decision view flags as fragile holds ${n0(fragN)} customers and THB ${(fragB / 1e6).toFixed(1)}m. Low stickiness here means dormant, not disloyal.`, options: { color: INK } },
  ], { x: 0.88, y: 5.78, w: 11.73, h: 0.86, margin: 0, valign: "middle", fontFace: F, fontSize: 11 });

  s.addNotes(`Move rate rises with stickiness, which is the opposite of what the name suggests: Very High ${(rate((r) => r.st === "Very High").r * 100).toFixed(1)}% against Very Low ${(rate((r) => r.st === "Very Low").r * 100).toFixed(1)}%. ` +
    "The reason is dormancy, not loyalty - average Save Max balance falls from THB 18,104 in the Very High tier to THB 18 in the Very Low tier, so the low tiers have nothing to move and nobody managing the account. " +
    `Read the columns, not the rows: the High rate-sensitivity column is hot at every stickiness level, peaking at ${(bestC.r * 100).toFixed(1)}% for High stickiness. ` +
    `The fragile quadrant - stickiness < 60 with rate sensitivity >= 60 - is ${n0(fragN)} customers holding THB ${n0(fragB)}, so the decision view on slide 2 points at an almost empty cell. ` +
    "On predictive value, stickiness adds little once rate sensitivity and the Save More flag are in hand (AUC 0.766 to 0.779), so its practical use is as an eligibility filter to suppress dormant micro-balance accounts from retention campaigns, not as a targeting dimension.");
}

/* ============================================================= SLIDE 13 */
/* Balance change by deposit stickiness. Absolute change and percentage
   change are shown for every tier; the prior-balance column is kept
   alongside so the large percentages on the small tiers can be read
   against the base they come from. */
{
  const s = pres.addSlide(); s.background = { color: PAPER };
  const D = fs.readFileSync(path.join(__dirname, "data", "savemax_maturity_stickiness.tsv"), "utf8")
    .trim().split("\n").slice(1).map((l) => {
      const c = l.split("\t");
      return { st: c[1], n: +c[4], p: +c[5], l: +c[6] };
    });
  const S = (f, g = () => true) => D.filter(g).reduce((a, r) => a + f(r), 0);
  const TIERS = ["Very High", "High", "Medium", "Low", "Very Low"];
  const n0 = (v) => Math.round(v).toLocaleString("en-US");
  const mAdapt = (v) => (v / 1e6 >= 10 ? (v / 1e6).toFixed(1) : (v / 1e6).toFixed(2));
  const dm = (v) => (v >= 0 ? "+" : "−") + Math.abs(v / 1e6).toFixed(1);
  const dpct = (p, l) => (l >= p ? "+" : "−") + Math.abs((l / p - 1) * 100).toFixed(1) + "%";
  const R = TIERS.map((t) => {
    const g = (r) => r.st === t;
    return { t, n: S((r) => r.n, g), p: S((r) => r.p, g), l: S((r) => r.l, g) };
  });
  const tot = { n: S((r) => r.n), p: S((r) => r.p), l: S((r) => r.l) };
  const gain = tot.l - tot.p;
  const sticky = R.filter((r) => r.t === "Very High" || r.t === "High").reduce((a, r) => a + (r.l - r.p), 0);
  const vl = R.find((r) => r.t === "Very Low"), hi = R.find((r) => r.t === "High");

  head(s, "Balance change by deposit stickiness",
    "*Change in total deposit balance, prior to latest. Percentage change is shown for every tier — read it against the prior-balance column.",
    `Cohort ${n0(tot.n)}   ·   Prior THB ${mAdapt(tot.p)}m → latest THB ${mAdapt(tot.l)}m   ·   Change ${dm(gain)}m (${dpct(tot.p, tot.l)})`,
    "Every tier gained. The percentages reverse the ranking, but the low tiers grow from near-zero balances.");

  tiles(s, [
    [dm(gain) + "m", "total balance change", `${dpct(tot.p, tot.l)} across the cohort`, TEAL],
    [dm(hi.l - hi.p) + "m", "High stickiness tier", `${dpct(hi.p, hi.l)} · largest contributor`, TEAL],
    [dpct(vl.p, vl.l), "Very Low stickiness", `on a THB ${mAdapt(vl.p)}m prior balance`, AMBER, 20],
    [(sticky / gain * 100).toFixed(1) + "%", "of the gain", "came from stickiness ≥ 60", TEAL],
  ], 1.80);

  /* the table ----------------------------------------------------------- */
  card(s, 0.58, 3.06, 7.35, 2.54);
  cardHead(s, 0.58, 3.06, 7.35, "Total deposit balance by stickiness tier", "THB m · absolute change and percentage change for every tier");
  const th = (t, al) => ({ text: t, options: { fill: { color: TITLE_TEAL }, color: PAPER, bold: true, fontFace: F, fontSize: 8.5, align: al || "right" } });
  const td = (t, al, col, bold) => ({ text: t, options: { fill: { color: PAPER }, color: col || INK, bold: !!bold, fontFace: F, fontSize: 9, align: al || "right" } });
  const body = R.map((r) => {
    const d = r.l - r.p, up = d >= 0;
    return [
      { text: r.t, options: { fill: { color: CARD }, color: INK, bold: true, fontFace: F, fontSize: 9, align: "left" } },
      td(n0(r.n)),
      td(mAdapt(r.p)),
      td(mAdapt(r.l)),
      td(dm(d), "right", up ? R3 : RED, true),
      td(dpct(r.p, r.l), "right", up ? R3 : RED, true),
    ];
  });
  const totRow = [
    { text: "All tiers", options: { fill: { color: CARD }, color: INK, bold: true, fontFace: F, fontSize: 9, align: "left" } },
    ...[n0(tot.n), mAdapt(tot.p), mAdapt(tot.l), dm(gain), dpct(tot.p, tot.l)].map((t, i) => ({
      text: t, options: { fill: { color: CARD }, color: i >= 3 ? R3 : INK, bold: true, fontFace: F, fontSize: 9, align: "right" },
    })),
  ];
  s.addTable([
    [th("Stickiness", "left"), th("Customers"), th("Prior"), th("Latest"), th("Change"), th("Change %")],
    ...body, totRow,
  ], {
    x: 0.82, y: 3.80, w: 6.87, colW: [1.22, 1.06, 1.12, 1.12, 1.12, 1.23],
    rowH: [0.25, 0.25, 0.25, 0.25, 0.25, 0.25, 0.25],
    border: { type: "solid", color: RULE, pt: 0.5 }, valign: "middle", margin: [0.01, 0.06, 0.01, 0.06],
  });

  /* contribution to the total gain -------------------------------------- */
  card(s, 8.18, 3.06, 4.73, 2.54);
  cardHead(s, 8.18, 3.06, 4.73, "Share of the total gain", `Each tier's contribution to ${dm(gain)}m`);
  const stCol = { "Very High": R4, High: R3, Medium: R2, Low: AMBER, "Very Low": R1 };
  const shMax = Math.max(...R.map((r) => Math.abs(r.l - r.p) / gain));
  R.forEach((r, i) => {
    const sh = (r.l - r.p) / gain, y = 3.76 + i * 0.30;
    s.addText(r.t, { x: 8.46, y, w: 1.15, h: 0.26, margin: 0, valign: "middle", fontFace: F, fontSize: 10, bold: true, color: INK });
    s.addShape("roundRect", { x: 9.66, y: y + 0.04, w: Math.max((Math.abs(sh) / shMax) * 1.55, 0.03), h: 0.18, rectRadius: 0.03, fill: { color: stCol[r.t] }, line: { type: "none" } });
    s.addText((sh * 100).toFixed(1) + "%", { x: 11.26, y, w: 0.72, h: 0.26, margin: 0, align: "right", valign: "middle", fontFace: F, fontSize: 10, bold: true, color: stCol[r.t] });
    s.addText(dm(r.l - r.p), { x: 12.00, y, w: 0.63, h: 0.26, margin: 0, align: "right", valign: "middle", fontFace: F, fontSize: 8.5, color: MUTED });
  });
  card(s, 8.46, 5.30, 4.17, 0.30, WARM);
  s.addText(`Stickiness ≥ 60 holds ${(S((r) => r.p, (r) => r.st === "Very High" || r.st === "High") / tot.p * 100).toFixed(1)}% of the prior balance.`, {
    x: 8.62, y: 5.30, w: 3.85, h: 0.30, margin: 0, valign: "middle", fontFace: F, fontSize: 9.5, color: INK,
  });

  card(s, 0.58, 5.78, 12.33, 0.86, "E8F6F3");
  s.addText([
    { text: `Stickiness ≥ 60 contributed ${dm(sticky)}m of the ${dm(gain)}m gain`, options: { bold: true, color: TITLE_TEAL } },
    { text: `  —  the Low and Very Low tiers post ${dpct(R[3].p, R[3].l)} and ${dpct(vl.p, vl.l)}, but on prior balances of THB ${mAdapt(R[3].p)}m and THB ${mAdapt(vl.p)}m. Rank the tiers on absolute change; read the percentages as a signal that dormant accounts are being funded again.`, options: { color: INK } },
  ], { x: 0.88, y: 5.78, w: 11.73, h: 0.86, margin: 0, valign: "middle", fontFace: F, fontSize: 11 });

  s.addNotes(`Total deposit balance rose from THB ${n0(tot.p)} to THB ${n0(tot.l)}, ${dm(gain)}m or ${dpct(tot.p, tot.l)}. ` +
    `Every stickiness tier gained. In absolute terms the gain is concentrated: High ${dm(hi.l - hi.p)}m and Very High ${dm(R[0].l - R[0].p)}m together are ${(sticky / gain * 100).toFixed(1)}% of it, which follows from those two tiers holding ${(S((r) => r.p, (r) => r.st === "Very High" || r.st === "High") / tot.p * 100).toFixed(1)}% of prior balance. ` +
    `In percentage terms the ranking reverses: Very High ${dpct(R[0].p, R[0].l)}, High ${dpct(hi.p, hi.l)}, Medium ${dpct(R[2].p, R[2].l)}, Low ${dpct(R[3].p, R[3].l)}, Very Low ${dpct(vl.p, vl.l)}. ` +
    `Those last two are computed on prior balances of THB ${n0(R[3].p)} and THB ${n0(vl.p)} across ${n0(R[3].n)} and ${n0(vl.n)} customers - average prior balances of THB ${n0(R[3].p / R[3].n)} and THB ${n0(vl.p / vl.n)} - so they represent dormant accounts receiving small deposits rather than material growth. ` +
    "Both readings are kept on the slide deliberately: the absolute column is the one to rank on, and the percentage column is a useful early indicator that previously unfunded accounts are becoming funded.");
}

/* ============================================================= SLIDE 14 */
/* The two scores side by side on both measures, to show that the
   move-rate / balance-change inversion belongs to rate sensitivity and
   does not repeat for deposit stickiness. */
{
  const s = pres.addSlide(); s.background = { color: PAPER };
  const D = fs.readFileSync(path.join(__dirname, "data", "savemax_maturity_stickiness.tsv"), "utf8")
    .trim().split("\n").slice(1).map((l) => {
      const c = l.split("\t");
      return { rs: c[0], st: c[1], mv: c[3] === "Move", n: +c[4], p: +c[5], l: +c[6] };
    });
  const S = (f, g = () => true) => D.filter(g).reduce((a, r) => a + f(r), 0);
  const TIERS = ["Very High", "High", "Medium", "Low", "Very Low"];
  const PMIN = 5e6, NMIN = 50;
  const n0 = (v) => Math.round(v).toLocaleString("en-US");
  const dm = (v) => (v >= 0 ? "+" : "−") + Math.abs(v / 1e6).toFixed(1);
  const row = (key, t) => {
    const g = (r) => r[key] === t, n = S((r) => r.n, g);
    if (!n) return null;
    const p = S((r) => r.p, g), l = S((r) => r.l, g);
    return { t, n, p, d: l - p, mv: S((r) => (r.mv ? r.n : 0), g) / n, pct: p >= PMIN ? (l / p - 1) * 100 : null };
  };
  const stR = TIERS.map((t) => row("st", t)).filter(Boolean);
  const rsR = TIERS.map((t) => row("rs", t)).filter(Boolean);
  const gainSticky = S((r) => r.l - r.p, (r) => r.st === "Very High" || r.st === "High");
  const gainAll = S((r) => r.l - r.p);
  const hiRS = rsR.find((r) => r.t === "High"), vhRS = rsR.find((r) => r.t === "Very High");
  const vhST = stR.find((r) => r.t === "Very High"), vlST = stR.find((r) => r.t === "Very Low");

  head(s, "The inversion belongs to rate sensitivity, not stickiness",
    `*Move rate = share of tier that emptied Save Max. Percentages suppressed below THB ${PMIN / 1e6}m prior balance; tiers under ${NMIN} customers greyed.`,
    `Cohort ${n0(S((r) => r.n))}   ·   Move rate ${(S((r) => (r.mv ? r.n : 0)) / S((r) => r.n) * 100).toFixed(1)}%   ·   Balance ${dm(gainAll)}m   ·   Stickiness ≥ 60 contributed ${dm(gainSticky)}m of it`,
    "Stickiness ranks both measures the same way; rate sensitivity ranks them in opposite directions.");

  tiles(s, [
    [`${(vlST.mv * 100).toFixed(1)}% → ${(vhST.mv * 100).toFixed(1)}%`, "move rate by stickiness", "Very Low tier to Very High tier", TEAL, 20],
    [dm(gainSticky) + "m", "gained by stickiness ≥ 60", `${(gainSticky / gainAll * 100).toFixed(0)}% of the cohort's total gain`, TEAL],
    [`${(hiRS.mv * 100).toFixed(1)}% / ${(vhRS.mv * 100).toFixed(1)}%`, "move rate, High / V.High", "the score's top band moves least", RED, 18],
    [`${dm(hiRS.d)} / ${dm(vhRS.d)}`, "balance change, THB m", "same pair, opposite directions", RED, 18],
  ], 1.80);

  /* one card per score: move-rate bar, then diverging balance bar -------- */
  const panel = (X, title, sub, rows) => {
    card(s, X, 3.06, 6.04, 2.54);
    cardHead(s, X, 3.06, 6.04, title, sub);
    const mvMax = Math.max(...rows.map((r) => r.mv));
    const dMax = Math.max(...rows.map((r) => Math.abs(r.d)));
    const ZX = X + 3.92, HALF = 0.52;
    s.addText("Move rate", { x: X + 1.34, y: 3.70, w: 1.62, h: 0.20, margin: 0, fontFace: F, fontSize: 8.5, bold: true, color: MUTED });
    s.addText("Balance change, THB m", { x: X + 3.30, y: 3.70, w: 2.46, h: 0.20, margin: 0, align: "center", fontFace: F, fontSize: 8.5, bold: true, color: MUTED });
    s.addShape("rect", { x: ZX, y: 3.92, w: 0.012, h: 1.58, fill: { color: MUTED }, line: { type: "none" } });
    rows.forEach((r, i) => {
      const y = 3.96 + i * 0.32, faint = r.n < NMIN;
      const mvCol = faint ? RULE : R4, dCol = faint ? RULE : r.d >= 0 ? R3 : RED;
      s.addText(r.t, { x: X + 0.28, y, w: 1.00, h: 0.28, margin: 0, valign: "middle", fontFace: F, fontSize: 9.5, bold: true, color: faint ? MUTED : INK });
      s.addShape("roundRect", { x: X + 1.34, y: y + 0.05, w: Math.max((r.mv / mvMax) * 1.10, 0.03), h: 0.18, rectRadius: 0.03, fill: { color: mvCol }, line: { type: "none" } });
      s.addText((r.mv * 100).toFixed(1) + "%", { x: X + 2.48, y, w: 0.56, h: 0.28, margin: 0, align: "right", valign: "middle", fontFace: F, fontSize: 9.5, bold: true, color: faint ? MUTED : R4 });
      const w = Math.max((Math.abs(r.d) / dMax) * HALF, 0.03);
      s.addShape("roundRect", { x: r.d >= 0 ? ZX : ZX - w, y: y + 0.05, w, h: 0.18, rectRadius: 0.03, fill: { color: dCol }, line: { type: "none" } });
      s.addText(dm(r.d) + (r.pct === null ? "" : ` (${r.pct >= 0 ? "+" : "−"}${Math.abs(r.pct).toFixed(1)}%)`), {
        x: X + 4.48, y, w: 1.28, h: 0.28, margin: 0, align: "right", valign: "middle",
        fontFace: F, fontSize: 8.5, bold: true, color: faint ? MUTED : dCol,
      });
    });
  };
  panel(0.58, "Deposit stickiness", "Both measures rise together — no inversion", stR);
  panel(6.87, "Rate sensitivity", "The two measures point in opposite directions", rsR);

  card(s, 0.58, 5.78, 12.33, 0.86, WARM);
  s.addText([
    { text: "Only rate sensitivity needs re-fitting", options: { bold: true, color: RED } },
    { text: `  —  stickiness is directionally consistent, though its alignment is partly a size effect: the two top tiers hold 98.8% of the balance. Rate sensitivity genuinely diverges — High gained ${dm(hiRS.d)}m while Very High lost ${Math.abs(vhRS.d / 1e6).toFixed(1)}m.`, options: { color: INK } },
  ], { x: 0.88, y: 5.78, w: 11.73, h: 0.86, margin: 0, valign: "middle", fontFace: F, fontSize: 11 });

  s.addNotes("Deposit stickiness ranks the two measures consistently: move rate runs 1.9% to 19.2% from Very Low to Very High, and balance change is positive in every tier. " +
    "The caveat is that this consistency is partly mechanical - the High and Very High tiers hold 98.8% of prior balance and contributed 95.5% of the cohort's gain, so the ranking largely restates where the money is. " +
    "In percentage terms the picture is flatter and mildly reversed: Very High +3.2%, High +4.0%, Medium +8.8%, with the Low and Very Low tiers unusable because their prior balances are near zero. " +
    "Rate sensitivity diverges for real: High moves at 49.5% and gained THB 144.7m, while Very High moves at 12.9% and lost THB 36.0m. Very Low also lost, THB 11.0m on a 9.6% move rate. " +
    "Medium is 4 customers in both scores and is greyed out rather than read. " +
    "Conclusion for the framework: stickiness can stay as an activity and eligibility filter, but rate-sensitivity cut-offs should be re-fitted on observed outcomes, and the two outcomes - emptying the pocket and losing balance - should be scored separately.");
}

const out = path.join(__dirname, "Deposit_Stickiness_Framework.pptx");
pres.writeFile({ fileName: out }).then(() => console.log("wrote", out));
