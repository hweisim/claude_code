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
/* 5x5 stickiness x rate-sensitivity heat matrix as a native table */
function matrix(s, x, y, w, key, fmt) {
  const data = M[key];
  const max = Math.max(...data.flatMap((r) => r.v));
  const labelW = w * 0.22, cellW = (w - labelW) / 5;
  const hdr = [
    { text: "Stickiness", options: { fill: { color: TITLE_TEAL }, color: PAPER, bold: true, fontFace: F, fontSize: 9, align: "left" } },
    ...RS.map((n, i) => ({
      text: n,
      options: { fill: { color: i >= 3 ? AMBER : TITLE_TEAL }, color: PAPER, bold: true, fontFace: F, fontSize: 9, align: "center" },
    })),
  ];
  const body = data.map((r) => [
    { text: r.s, options: { fill: { color: CARD }, color: INK, bold: true, fontFace: F, fontSize: 9.5, align: "left" } },
    ...r.v.map((v) => {
      const t = max ? Math.sqrt(v / max) : 0;
      return {
        text: fmt(v),
        options: {
          fill: { color: blend(R4, t) },
          color: t > 0.55 ? PAPER : INK,
          fontFace: F, fontSize: 9.5, align: "right",
        },
      };
    }),
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

/* ============================================================== SLIDE 6 */
{
  const s = pres.addSlide(); s.background = { color: PAPER };
  head(s, `All customers : ${(allCust / 1e6).toFixed(2)}m — a very different shape`,
    "*Full customer base · distinct ccd_id by deposit stickiness × rate sensitivity.",
    `Total customers ${(allCust / 1e6).toFixed(2)}m   ·   Low stickiness ${(allLowStick / 1e6).toFixed(2)}m   ·   High rate sensitivity ${k(allHiRS)}`,
    "The full base is dominated by Very Low stickiness; the funded Save Max cohort concentrates in High and Very High.");

  tiles(s, [
    [`${(allCust / 1e6).toFixed(2)}m`, "total customers", "distinct ccd_id", TEAL],
    ["TBD", "customers with balance", "populate from current_bal_amt > 0", AMBER],
    [`${(allLowStick / 1e6).toFixed(2)}m`, "low deposit stickiness", `${pc(allLowStick, allCust)} of the base`, RED],
    [k(allHiRS), "high rate sensitivity", `${(allHiRS / allCust * 100).toFixed(1)}% of the base`, TEAL],
  ], 1.80);

  card(s, 0.58, 3.06, 12.33, 2.54);
  cardHead(s, 0.58, 3.06, 12.33, "All-customer count", "Rows: stickiness · columns: rate sensitivity · distinct ccd_id");
  matrix(s, 0.82, 3.72, 11.85, "all_count", (v) => v.toLocaleString("en-US"));

  card(s, 0.58, 5.78, 12.33, 0.86, "E8F6F3");
  s.addText([
    { text: "Use the Save Max cohort for maturity and runoff risk; use the all-customer view for activation and funding strategy", options: { bold: true, color: TITLE_TEAL } },
    { text: `  —  ${pc(allLowStick, allCust)} of the base scores Low or Very Low stickiness, but only ${k(allHiRS)} customers are highly rate-sensitive, so the broad opportunity is engagement rather than repricing.`, options: { color: INK } },
  ], { x: 0.88, y: 5.78, w: 11.73, h: 0.86, margin: 0, valign: "middle", fontFace: F, fontSize: 11 });
  s.addNotes(`Full base ${allCust.toLocaleString()} customers. Low + Very Low stickiness ${allLowStick.toLocaleString()}. High + Very High rate sensitivity ${allHiRS.toLocaleString()}. "Customers with balance > 0" is still TBD in the source and needs populating from current_bal_amt > 0.`);
}

const out = path.join(__dirname, "Deposit_Stickiness_Framework.pptx");
pres.writeFile({ fileName: out }).then(() => console.log("wrote", out));
