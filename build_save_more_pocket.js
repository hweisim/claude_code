/**
 * Save More gap - why customers without a second savings pocket take
 * their money out of the bank, and what to do about it.
 * Built to SLIDE_STYLE_GUIDE.md. Every figure comes from
 * data/save_more_quality.tsv exactly as supplied.
 *
 *   node build_save_more_pocket.js
 */

const path = require("path");
const fs = require("fs");
const PptxGenJS = require("pptxgenjs");

/* ------------------------------------------------------- house constants */
const TITLE_TEAL = "045D66", INK = "16233D", MUTED = "6B7A90", NOTE = "6B6B6B";
const TEAL = "00A896", RED = "C00000", AMBER = "E9A03B", CARD = "F1F4F9";
const PAPER = "FFFFFF", RULE = "D8DFE9", WARM = "FBF0E6", MINT = "E8F6F3";
const R1 = "8DA0B8", R2 = "3FB89F", R3 = "1E8E86", R4 = "165A73";
const F = "Graphik TH", FH = "Cambria";

/* ------------------------------------------------------------------ data */
const D = fs.readFileSync(path.join(__dirname, "data", "save_more_quality.tsv"), "utf8")
  .trim().split("\n").slice(1).map((l) => {
    const c = l.split("\t");
    return { mv: c[0] === "Move", sm: +c[1], q: +c[2], n: +c[3], p: +c[4], l: +c[5] };
  });
const S = (f, g = () => true) => D.filter(g).reduce((a, r) => a + f(r), 0);
const agg = (g) => {
  const n = S((r) => r.n, g), p = S((r) => r.p, g), l = S((r) => r.l, g);
  return { n, p, l, d: l - p, pct: (l / p - 1) * 100, avg: p / n };
};

const ALL = agg(() => true);
const noSM = agg((r) => r.sm === 0), hasSM = agg((r) => r.sm === 1);
const noSMmv = agg((r) => r.sm === 0 && r.mv), noSMst = agg((r) => r.sm === 0 && !r.mv);
const hasSMmv = agg((r) => r.sm === 1 && r.mv), hasSMst = agg((r) => r.sm === 1 && !r.mv);
const qual = agg((r) => r.q === 1), nonq = agg((r) => r.q === 0);
const TARGET = agg((r) => r.q === 1 && r.sm === 0);
const TGTmv = agg((r) => r.q === 1 && r.sm === 0 && r.mv);
const TGTst = agg((r) => r.q === 1 && r.sm === 0 && !r.mv);
const qHas = agg((r) => r.q === 1 && r.sm === 1);
const nqNo = agg((r) => r.q === 0 && r.sm === 0);
const nqHas = agg((r) => r.q === 0 && r.sm === 1);
const mvRateNo = S((r) => r.n, (r) => r.sm === 0 && r.mv) / noSM.n * 100;
const mvRateHas = S((r) => r.n, (r) => r.sm === 1 && r.mv) / hasSM.n * 100;

const n0 = (v) => Math.round(v).toLocaleString("en-US");
const mm = (v) => (v / 1e6).toFixed(1);
const dm = (v) => (v >= 0 ? "+" : "−") + "THB " + Math.abs(v / 1e6).toFixed(1) + "m";
const dp = (v) => (v >= 0 ? "+" : "−") + Math.abs(v).toFixed(1) + "%";
const p1 = (a, b) => (a / b * 100).toFixed(1) + "%";

/* --------------------------------------------------------------- helpers */
const pres = new PptxGenJS();
pres.layout = "LAYOUT_WIDE";
pres.author = "Deposit Analytics";
pres.title = "Save More gap and the savings-pocket opportunity";

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
  defs.forEach(([v, l, sub, c, size], i) => {
    const x = 0.58 + i * (w + gap);
    card(s, x, y, w, 1.08);
    s.addText(v, { x: x + 0.24, y: y + 0.08, w: w - 0.48, h: 0.44, margin: 0, fontFace: F, fontSize: size || 24, bold: true, color: c || TEAL });
    s.addText(l, { x: x + 0.24, y: y + 0.52, w: w - 0.48, h: 0.26, margin: 0, valign: "top", fontFace: F, fontSize: 11.5, bold: true, color: INK });
    if (sub) s.addText(sub, { x: x + 0.24, y: y + 0.77, w: w - 0.48, h: 0.26, margin: 0, valign: "top", fontFace: F, fontSize: 9.5, color: MUTED });
  });
}
/* label | proportional bar | value, with an optional reference line */
function barRow(s, y, lx, lw, tx, tw, vx, vw, label, frac, col, value, size = 10) {
  s.addText(label, { x: lx, y, w: lw, h: 0.28, margin: 0, valign: "middle", fontFace: F, fontSize: size, color: INK });
  s.addShape("roundRect", { x: tx, y: y + 0.05, w: tw, h: 0.18, rectRadius: 0.03, fill: { color: "E4E9F0" }, line: { type: "none" } });
  s.addShape("roundRect", { x: tx, y: y + 0.05, w: Math.max(tw * frac, 0.03), h: 0.18, rectRadius: 0.03, fill: { color: col }, line: { type: "none" } });
  s.addText(value, { x: vx, y, w: vw, h: 0.28, margin: 0, align: "right", valign: "middle", fontFace: F, fontSize: size, bold: true, color: col });
}
const th = (t, al) => ({ text: t, options: { fill: { color: TITLE_TEAL }, color: PAPER, bold: true, fontFace: F, fontSize: 8.5, align: al || "right" } });
const td = (t, al, col, bold) => ({ text: t, options: { fill: { color: PAPER }, color: col || INK, bold: !!bold, fontFace: F, fontSize: 9, align: al || "right" } });
const rh = (t) => ({ text: t, options: { fill: { color: CARD }, color: INK, bold: true, fontFace: F, fontSize: 9, align: "left" } });

/* ============================================================== SLIDE 1 */
{
  const s = pres.addSlide(); s.background = { color: PAPER };

  head(s, "Without Save More, the money leaves the bank",
    "*Save More flag read as held (1) or not held (null in source). Balances are total deposit balance, prior against latest period.",
    `Base ${n0(ALL.n)}   ·   Without Save More ${n0(noSM.n)} (${dm(noSM.d)}, ${dp(noSM.pct)})   ·   With Save More ${n0(hasSM.n)} (${dm(hasSM.d)}, ${dp(hasSM.pct)})`,
    "The two groups move in opposite directions: without Save More the balance goes, with it the balance stays.");

  tiles(s, [
    [n0(noSM.n), "without Save More", `${p1(noSM.n, ALL.n)} of base · avg THB ${n0(noSM.avg)}`, RED],
    [dm(noSM.d), "balance lost", `${dp(noSM.pct)} period on period`, RED],
    [n0(hasSM.n), "with Save More", `${p1(hasSM.n, ALL.n)} of base · avg THB ${n0(hasSM.avg)}`, TEAL],
    [dm(hasSM.d), "balance gained", `${dp(hasSM.pct)} period on period`, TEAL],
  ], 1.80);

  card(s, 0.58, 3.06, 7.35, 2.54);
  cardHead(s, 0.58, 3.06, 7.35, "Balance by Save More holding and Save Max behaviour", "Balances in THB m · 'Move' means the Save Max pocket is now empty");
  const rows = [
    ["No Save More · moved", noSMmv, true], ["No Save More · stayed", noSMst, true],
    ["Has Save More · moved", hasSMmv, false], ["Has Save More · stayed", hasSMst, false],
  ];
  s.addTable([
    [th("Group", "left"), th("Customers"), th("Prior"), th("Latest"), th("Change"), th("Change %")],
    ...rows.map(([lab, g, bad]) => [
      rh(lab), td(n0(g.n)), td(mm(g.p)), td(mm(g.l)),
      td(dm(g.d).replace("THB ", ""), "right", g.d >= 0 ? R3 : RED, true),
      td(dp(g.pct), "right", g.d >= 0 ? R3 : RED, true),
    ]),
    [rh("All"), ...[n0(ALL.n), mm(ALL.p), mm(ALL.l), dm(ALL.d).replace("THB ", ""), dp(ALL.pct)].map((t, i) => ({
      text: t, options: { fill: { color: CARD }, color: i >= 3 ? RED : INK, bold: true, fontFace: F, fontSize: 9, align: "right" },
    }))],
  ], {
    x: 0.82, y: 3.80, w: 6.87, colW: [1.92, 1.02, 0.92, 0.92, 1.02, 1.07],
    rowH: [0.26, 0.28, 0.28, 0.28, 0.28, 0.28],
    border: { type: "solid", color: RULE, pt: 0.5 }, valign: "middle", margin: [0.01, 0.06, 0.01, 0.06],
  });

  card(s, 8.18, 3.06, 4.73, 2.54);
  cardHead(s, 8.18, 3.06, 4.73, "Balance retained", "Latest ÷ prior · the rule marks 100%");
  const SCALE = 1.10, TX = 9.86, TW = 1.55;
  s.addShape("rect", { x: TX + TW / SCALE, y: 3.78, w: 0.012, h: 1.42, fill: { color: MUTED }, line: { type: "none" } });
  [["No SM · moved", noSMmv, RED], ["No SM · stayed", noSMst, RED],
   ["Has SM · moved", hasSMmv, R3], ["Has SM · stayed", hasSMst, R3]].forEach(([lab, g, col], i) => {
    const ret = g.l / g.p;
    barRow(s, 3.82 + i * 0.34, 8.46, 1.36, TX, TW, 11.56, 1.07, lab, ret / SCALE, col, (ret * 100).toFixed(1) + "%", 9.5);
  });
  s.addText("100%", { x: TX + TW / SCALE - 0.30, y: 5.22, w: 0.60, h: 0.18, margin: 0, align: "center", fontFace: F, fontSize: 8, color: MUTED });

  card(s, 0.58, 5.78, 12.33, 0.86, WARM);
  s.addText([
    { text: `Customers who emptied Save Max without a Save More holding kept only ${(noSMmv.l / noSMmv.p * 100).toFixed(0)}% of their balance`, options: { bold: true, color: RED } },
    { text: `  —  ${dm(noSMmv.d)} gone. The same behaviour with a Save More holding retained ${(hasSMmv.l / hasSMmv.p * 100).toFixed(0)}%. Without a second pocket to move into, the money moves out of the bank.`, options: { color: INK } },
  ], { x: 0.88, y: 5.78, w: 11.73, h: 0.86, margin: 0, valign: "middle", fontFace: F, fontSize: 11 });

  s.addNotes(`Base of ${n0(ALL.n)} customers holding THB ${n0(ALL.p)} prior, THB ${n0(ALL.l)} latest, ${dm(ALL.d)} or ${dp(ALL.pct)}. ` +
    `The split on the Save More flag is the whole story: ${n0(noSM.n)} customers without it lost ${dm(noSM.d)} (${dp(noSM.pct)}), while ${n0(hasSM.n)} with it gained ${dm(hasSM.d)} (${dp(hasSM.pct)}). ` +
    `Sharpest of all is the group that emptied Save Max without a Save More holding: ${n0(noSMmv.n)} customers retaining only ${(noSMmv.l / noSMmv.p * 100).toFixed(1)}% of balance. ` +
    `Note the direction of the move rate: Save More holders move out of Save Max far more often (${mvRateHas.toFixed(1)}% against ${mvRateNo.toFixed(1)}%) but the money stays in the bank, so moving the pocket and leaving the bank are different events. ` +
    "Caveat on composition: Save More holders are much larger balances - average THB " + n0(hasSM.avg) + " against THB " + n0(noSM.avg) + " - so part of the gap is who holds the product rather than the product itself.");
}

/* ============================================================== SLIDE 2 */
{
  const s = pres.addSlide(); s.background = { color: PAPER };

  head(s, "The whole loss sits with quality customers who have no Save More",
    "*Quality customer = makes transactions. Balances are total deposit balance, prior against latest period.",
    `Quality ${n0(qual.n)} (${p1(qual.n, ALL.n)}) ${dm(qual.d)}   ·   Non-quality ${n0(nonq.n)} ${dm(nonq.d)}   ·   Target ${n0(TARGET.n)} ${dm(TARGET.d)}`,
    "Quality customers holding Save More kept their balance flat. Those without lost 17.2%.");

  tiles(s, [
    [p1(qual.n, ALL.n), "are quality customers", `${n0(qual.n)} of ${n0(ALL.n)} · they transact`, TEAL],
    [n0(TARGET.n), "quality, no Save More", "the group losing money", RED],
    [dm(TARGET.d), "their balance loss", `${dp(TARGET.pct)} period on period`, RED],
    [dp(qHas.pct), "quality with Save More", `THB ${mm(qHas.p)}m held flat`, TEAL],
  ], 1.80);

  card(s, 0.58, 3.06, 7.35, 2.54);
  cardHead(s, 0.58, 3.06, 7.35, "Quality against Save More holding", "Balances in THB m · the loss is confined to one cell");
  const grid = [
    ["Quality · no Save More", TARGET, true], ["Quality · has Save More", qHas, false],
    ["Non-quality · no Save More", nqNo, false], ["Non-quality · has Save More", nqHas, false],
  ];
  s.addTable([
    [th("Group", "left"), th("Customers"), th("Prior"), th("Latest"), th("Change"), th("Change %")],
    ...grid.map(([lab, g, bad]) => [
      { text: lab, options: { fill: { color: bad ? "F6E2E2" : CARD }, color: bad ? RED : INK, bold: true, fontFace: F, fontSize: 9, align: "left" } },
      ...[n0(g.n), mm(g.p), mm(g.l), dm(g.d).replace("THB ", ""), dp(g.pct)].map((t, i) => ({
        text: t,
        options: { fill: { color: bad ? "F6E2E2" : PAPER }, color: i >= 3 ? (g.d >= 0 ? R3 : RED) : INK, bold: i >= 3 || bad, fontFace: F, fontSize: 9, align: "right" },
      })),
    ]),
  ], {
    x: 0.82, y: 3.80, w: 6.87, colW: [2.14, 0.98, 0.88, 0.88, 0.97, 1.02],
    rowH: [0.26, 0.32, 0.32, 0.32, 0.32],
    border: { type: "solid", color: RULE, pt: 0.5 }, valign: "middle", margin: [0.01, 0.06, 0.01, 0.06],
  });
  s.addText(`Quality customers are ${p1(qual.p, ALL.p)} of prior balance — this is not a small-value population.`, {
    x: 0.86, y: 5.38, w: 6.8, h: 0.20, margin: 0, fontFace: F, fontSize: 9, italic: true, color: MUTED,
  });

  card(s, 8.18, 3.06, 4.73, 2.54);
  cardHead(s, 8.18, 3.06, 4.73, "Inside the target group", "Quality customers with no Save More");
  [[n0(TGTmv.n), "emptied Save Max", `${dp(TGTmv.pct)} · ${dm(TGTmv.d)}`, RED],
   [n0(TGTst.n), "stayed in Save Max", `${dp(TGTst.pct)} · ${dm(TGTst.d)}`, AMBER]].forEach(([v, l, sub, col], i) => {
    const y = 3.80 + i * 0.72;
    s.addText(v, { x: 8.46, y, w: 1.55, h: 0.40, margin: 0, valign: "middle", fontFace: F, fontSize: 21, bold: true, color: col });
    s.addText(l, { x: 10.06, y, w: 2.57, h: 0.24, margin: 0, valign: "top", fontFace: F, fontSize: 10.5, bold: true, color: INK });
    s.addText(sub, { x: 10.06, y: y + 0.22, w: 2.57, h: 0.22, margin: 0, valign: "top", fontFace: F, fontSize: 9.5, color: col });
  });
  card(s, 8.46, 5.10, 4.17, 0.40, WARM);
  s.addText(`Even the ones who stayed put lost ${dp(TGTst.pct)} of balance.`, {
    x: 8.62, y: 5.10, w: 3.85, h: 0.40, margin: 0, valign: "middle", fontFace: F, fontSize: 9.5, color: INK,
  });

  card(s, 0.58, 5.78, 12.33, 0.86, WARM);
  s.addText([
    { text: `Quality customers with Save More held ${dp(qHas.pct)}; without it they lost ${dp(TARGET.pct)}`, options: { bold: true, color: RED } },
    { text: `  —  and non-quality customers gained in both columns. The loss is not about balance size or Save Max behaviour, it is about whether an engaged customer had somewhere else to keep the money.`, options: { color: INK } },
  ], { x: 0.88, y: 5.78, w: 11.73, h: 0.86, margin: 0, valign: "middle", fontFace: F, fontSize: 11 });

  s.addNotes(`Quality customers are ${n0(qual.n)}, ${p1(qual.n, ALL.n)} of the base and ${p1(qual.p, ALL.p)} of prior balance. They lost ${dm(qual.d)} (${dp(qual.pct)}) while non-quality customers gained ${dm(nonq.d)} (${dp(nonq.pct)}). ` +
    `Crossing quality with the Save More flag isolates it completely: quality customers without Save More lost ${dm(TARGET.d)}, which is the entire quality-customer loss, while quality customers with Save More moved from THB ${n0(qHas.p)} to THB ${n0(qHas.l)}, effectively flat. ` +
    `Both non-quality cells gained ${dp(nqNo.pct)} and ${dp(nqHas.pct)}. ` +
    `Within the target group the damage is worst among those who emptied Save Max - ${n0(TGTmv.n)} customers down ${dp(TGTmv.pct)} - but the ${n0(TGTst.n)} who stayed still lost ${dp(TGTst.pct)}, so this is not only a maturity event. ` +
    "Caveat: quality is a binary supplied with the data and its definition should be confirmed before it is used for targeting.");
}

/* ============================================================== SLIDE 3 */
{
  const s = pres.addSlide(); s.background = { color: PAPER };
  const retainIfFlat = -TARGET.d;

  head(s, "Action plan : give these customers a second savings pocket",
    "*Sizing uses the observed target group. The retention figure is what the period would have looked like had the group held flat, not a forecast.",
    `Target ${n0(TARGET.n)} quality customers with no Save More   ·   Balance today THB ${mm(TARGET.l)}m   ·   Lost last period ${dm(TARGET.d)}`,
    "They transact, they hold balance, and they have nowhere in the bank to move it when the promotion ends.");

  tiles(s, [
    [n0(TARGET.n), "customers to target", `${p1(TARGET.n, ALL.n)} of the base · they transact`, TEAL],
    [`THB ${mm(TARGET.l)}m`, "balance still with us", `down from THB ${mm(TARGET.p)}m`, AMBER],
    [dm(retainIfFlat).replace("+", ""), "at stake per period", "if the decline continues", RED],
    [`THB ${n0(TARGET.avg)}`, "average balance", "small individually, large in total", TEAL],
  ], 1.80);

  card(s, 0.58, 3.06, 6.04, 2.54);
  cardHead(s, 0.58, 3.06, 6.04, "What to offer", "A pocket sized for this group, not for the Save More base");
  [
    ["Second savings pocket", "A capped, higher-rate pocket they can open in-app, sized for balances around THB " + n0(TARGET.avg) + " rather than the THB 1m Save More cap."],
    ["Open it before maturity", "Offer at the point the Save Max promotion is due to end, not after the balance has already moved."],
    ["Keep the transaction link", "Position it beside the account they already transact on, so the pocket adds to the relationship instead of replacing it."],
  ].forEach(([t, d], i) => {
    const y = 3.80 + i * 0.60;
    s.addShape("roundRect", { x: 0.86, y: y + 0.02, w: 0.26, h: 0.26, rectRadius: 0.05, fill: { color: TITLE_TEAL }, line: { type: "none" } });
    s.addText(String(i + 1), { x: 0.86, y: y + 0.02, w: 0.26, h: 0.26, margin: 0, align: "center", valign: "middle", fontFace: F, fontSize: 10, bold: true, color: PAPER });
    s.addText(t, { x: 1.24, y, w: 4.4, h: 0.24, margin: 0, fontFace: F, fontSize: 11, bold: true, color: INK });
    s.addText(d, { x: 1.24, y: y + 0.22, w: 4.42, h: 0.38, margin: 0, valign: "top", fontFace: F, fontSize: 9, color: MUTED });
  });

  card(s, 6.87, 3.06, 6.04, 2.54);
  cardHead(s, 6.87, 3.06, 6.04, "Why it should hold", "What the same customers did when they had a second pocket");
  const bars = [
    ["Quality · has Save More", qHas.l / qHas.p, R3],
    ["Quality · no Save More", TARGET.l / TARGET.p, RED],
    ["Non-quality · has Save More", nqHas.l / nqHas.p, R1],
    ["Non-quality · no Save More", nqNo.l / nqNo.p, R1],
  ];
  const SC = 1.10, TX2 = 9.36, TW2 = 1.50;
  s.addShape("rect", { x: TX2 + TW2 / SC, y: 3.78, w: 0.012, h: 1.42, fill: { color: MUTED }, line: { type: "none" } });
  bars.forEach(([lab, ret, col], i) =>
    barRow(s, 3.82 + i * 0.34, 7.15, 2.15, TX2, TW2, 11.06, 1.07, lab, ret / SC, col, (ret * 100).toFixed(1) + "%", 9.5));
  s.addText(`Balance retained. Engaged customers with a second pocket held flat; the same customers without one lost ${dp(TARGET.pct)}.`, {
    x: 7.15, y: 5.22, w: 5.48, h: 0.32, margin: 0, valign: "top", fontFace: F, fontSize: 9, italic: true, color: MUTED,
  });

  card(s, 0.58, 5.78, 12.33, 0.86, MINT);
  s.addText([
    { text: `${n0(TARGET.n)} customers who transact, hold THB ${mm(TARGET.l)}m, and are leaving at ${dp(TARGET.pct)} a period`, options: { bold: true, color: TITLE_TEAL } },
    { text: `  —  they are the customers worth keeping. Give them a pocket sized for a THB ${n0(TARGET.avg)} balance and the money has a reason to stay. Measure the next maturity wave against a held-out control before scaling.`, options: { color: INK } },
  ], { x: 0.88, y: 5.78, w: 11.73, h: 0.86, margin: 0, valign: "middle", fontFace: F, fontSize: 11 });

  s.addNotes(`Target group: ${n0(TARGET.n)} quality customers with no Save More, holding THB ${n0(TARGET.l)} today against THB ${n0(TARGET.p)} last period, ${dm(TARGET.d)} or ${dp(TARGET.pct)}. Average balance THB ${n0(TARGET.avg)}. ` +
    "The proposal is a second savings pocket sized for this group. Save More's THB 1m cap is built for the existing holder base, which averages THB " + n0(hasSM.avg) + "; a product aimed at a THB " + n0(TARGET.avg) + " balance is a different design. " +
    "Timing matters more than rate: the worst cell is the one that emptied Save Max, so the offer has to land before maturity rather than after the balance has already moved. " +
    `On sizing the prize: ${dm(retainIfFlat)} is simply the observed loss, shown as what the period would have looked like had the group held flat like quality customers with Save More. It is not a forecast, and the two groups differ sharply in balance - THB ${n0(TARGET.avg)} against THB ${n0(qHas.avg)} - so the comparison is indicative only. ` +
    "Recommend running the next maturity wave with a held-out control group so the effect can be measured rather than assumed.");
}

const out = path.join(__dirname, "Save_More_Pocket_Opportunity.pptx");
pres.writeFile({ fileName: out }).then(() => console.log("wrote", out));
