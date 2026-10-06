/**
 * CLICX captive opportunity - summary of the four source views:
 * overall captive, CLICX-from-captive segments, lending opportunity,
 * active-but-no-CLICX reach, and the TrueMoney Wallet overlap.
 * Built to SLIDE_STYLE_GUIDE.md. Every figure comes from data/clicx_*.tsv.
 *
 *   node build_clicx_captive.js
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
const rd = (f) => {
  const L = fs.readFileSync(path.join(__dirname, "data", f), "utf8").trim().split("\n");
  const h = L[0].split("\t");
  return L.slice(1).map((l) => {
    const c = l.split("\t"), o = {};
    h.forEach((k, i) => { o[k] = (c[i] !== "" && !isNaN(+c[i])) ? +c[i] : c[i]; });
    return o;
  });
};
const SEG = rd("clicx_segments.tsv");
const LEND = rd("clicx_lending.tsv");
const CASH = rd("clicx_cashneed.tsv");
const TMW = rd("clicx_tmw.tsv");
const REACH = rd("clicx_reach.tsv");
const sum = (A, k) => A.reduce((a, r) => a + r[k], 0);

const CAPTIVE = 48.0e6;                       /* as stated on the source */
const onClicx = sum(SEG, "have_clicx");
const noClicx = sum(SEG, "no_clicx");
const aisReach = sum(REACH, "ais_active"), ktbReach = sum(REACH, "ktb_active");
const hasOR = (r) => /OR/.test(r.segment);
const orBase = sum(SEG.filter(hasOR), "no_clicx") + sum(SEG.filter(hasOR), "have_clicx");
const orClicx = sum(SEG.filter(hasOR), "have_clicx");
const nonBase = sum(SEG.filter((r) => !hasOR(r)), "no_clicx") + sum(SEG.filter((r) => !hasOR(r)), "have_clicx");
const nonClicxOR = sum(SEG.filter((r) => !hasOR(r)), "have_clicx");
const penOR = orClicx / orBase, penNo = nonClicxOR / nonBase;
const appin = sum(LEND, "appin"), booked = sum(LEND, "booked"), drawn = sum(LEND, "drawdown");
const lOR = LEND.filter(hasOR), lNo = LEND.filter((r) => !hasOR(r));
const bookOR = sum(lOR, "booked") / sum(lOR, "appin"), bookNo = sum(lNo, "booked") / sum(lNo, "appin");
const mh = sum(CASH, "mh_cash"), lowc = sum(CASH, "low_cash");
const clicxT = TMW.filter((r) => r.scope === "clicx"), baseT = TMW.filter((r) => r.scope === "base");
const dep = clicxT.find((r) => r.segment === "Deposit"), pay = clicxT.find((r) => r.segment === "Payment");
const CLICX_TMW = 755000, CLICX_TOTAL = 1.7e6, BASE_TMW = 10e6, BASE_TOTAL = 45e6;

const M = (v) => (v >= 1e6 ? (v / 1e6).toFixed(1) + "M" : Math.round(v / 1e3) + "K");
const p1 = (a, b) => (a / b * 100).toFixed(1) + "%";
const p0 = (a, b) => Math.round(a / b * 100) + "%";

/* --------------------------------------------------------------- helpers */
const pres = new PptxGenJS();
pres.layout = "LAYOUT_WIDE";
pres.author = "CLICX Analytics";
pres.title = "CLICX captive opportunity";

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
function barRow(s, y, lx, lw, tx, tw, vx, vw, label, frac, col, value, size = 9.5) {
  s.addText(label, { x: lx, y, w: lw, h: 0.26, margin: 0, valign: "middle", fontFace: F, fontSize: size, color: INK });
  s.addShape("roundRect", { x: tx, y: y + 0.045, w: tw, h: 0.17, rectRadius: 0.03, fill: { color: "E4E9F0" }, line: { type: "none" } });
  s.addShape("roundRect", { x: tx, y: y + 0.045, w: Math.max(tw * frac, 0.03), h: 0.17, rectRadius: 0.03, fill: { color: col }, line: { type: "none" } });
  s.addText(value, { x: vx, y, w: vw, h: 0.26, margin: 0, align: "right", valign: "middle", fontFace: F, fontSize: size, bold: true, color: col });
}
const th = (t, al) => ({ text: t, options: { fill: { color: TITLE_TEAL }, color: PAPER, bold: true, fontFace: F, fontSize: 8, align: al || "right" } });
const td = (t, al, col, bold) => ({ text: t, options: { fill: { color: PAPER }, color: col || INK, bold: !!bold, fontFace: F, fontSize: 8.5, align: al || "right" } });
const rh = (t, hi) => ({ text: t, options: { fill: { color: hi ? MINT : CARD }, color: hi ? TITLE_TEAL : INK, bold: true, fontFace: F, fontSize: 8.5, align: "left" } });

/* ============================================================== SLIDE 1 */
{
  const s = pres.addSlide(); s.background = { color: PAPER };
  const prio = SEG.filter((r) => r.segment === "AIS & KTB" || r.segment === "AIS KTB OR");
  const prioClicx = sum(prio, "have_clicx");
  const prioBase = sum(prio, "no_clicx") + sum(prio, "have_clicx");

  head(s, "48M captive customers, 1.7M on CLICX",
    "*Captive = AIS × KTB × OR consortium base. Figures as supplied and rounded; the segment rows sum to 47.3M against a stated 48.0M.",
    `Captive ${M(CAPTIVE)}   ·   On CLICX ${M(onClicx)} (${p1(onClicx, noClicx + onClicx)})   ·   Reachable at AIS ${M(aisReach)}   ·   Reachable at KTB ${M(ktbReach)}`,
    `96% of the consortium base has never opened CLICX — ${M(aisReach)} of them are active at AIS and ${M(ktbReach)} at KTB today.`);

  tiles(s, [
    [M(CAPTIVE), "captive customers", "AIS × KTB × OR consortium", TEAL],
    [M(onClicx), "on CLICX", `${p1(onClicx, noClicx + onClicx)} of the captive base`, TEAL],
    [M(aisReach), "reachable at AIS", "no CLICX, active today", AMBER],
    [M(ktbReach), "reachable at KTB", "no CLICX, active today", AMBER],
  ], 1.80);

  card(s, 0.58, 3.06, 7.35, 2.54);
  cardHead(s, 0.58, 3.06, 7.35, "Where the captive base sits", "Penetration = CLICX customers ÷ segment base · priority segments highlighted");
  const hi = (n) => n === "AIS & KTB" || n === "AIS KTB OR";
  s.addTable([
    [th("Segment", "left"), th("No CLICX"), th("On CLICX"), th("Penetration"), th("App-in"), th("Deposit")],
    ...SEG.map((r) => {
      const base = r.no_clicx + r.have_clicx, h = hi(r.segment);
      return [rh(r.segment, h),
        td(M(r.no_clicx), "right", h ? TITLE_TEAL : INK, h),
        td(M(r.have_clicx), "right", h ? TITLE_TEAL : INK, h),
        td(p1(r.have_clicx, base), "right", r.have_clicx / base > 0.05 ? R3 : INK, true),
        td(M(r.appin), "right", h ? TITLE_TEAL : INK, h),
        td(M(r.deposit), "right", h ? TITLE_TEAL : INK, h)];
    }),
    [rh("Total"), ...[M(noClicx), M(onClicx), p1(onClicx, noClicx + onClicx), M(sum(SEG, "appin")), M(sum(SEG, "deposit"))].map((t) => ({
      text: t, options: { fill: { color: CARD }, color: INK, bold: true, fontFace: F, fontSize: 8.5, align: "right" },
    }))],
  ], {
    x: 0.82, y: 3.80, w: 6.87, colW: [1.42, 1.07, 1.07, 1.17, 1.07, 1.07],
    rowH: [0.24, 0.21, 0.21, 0.21, 0.21, 0.21, 0.21, 0.21, 0.23],
    border: { type: "solid", color: RULE, pt: 0.5 }, valign: "middle", margin: [0.01, 0.06, 0.01, 0.06],
  });

  card(s, 8.18, 3.06, 4.73, 2.54);
  cardHead(s, 8.18, 3.06, 4.73, "Where to find the 48M", "Not on CLICX, active at a consortium partner");
  const rMax = Math.max(...REACH.map((r) => Math.max(r.ais_active, r.ktb_active)));
  REACH.forEach((r, i) => {
    const y = 3.76 + i * 0.27;
    s.addText(r.segment, { x: 8.46, y, w: 1.30, h: 0.26, margin: 0, valign: "middle", fontFace: F, fontSize: 8.5, bold: true, color: INK });
    [["ais_active", R4, 9.82], ["ktb_active", R2, 11.28]].forEach(([k, col, bx]) => {
      if (!r[k]) return;
      s.addShape("roundRect", { x: bx, y: y + 0.045, w: Math.max((r[k] / rMax) * 0.90, 0.03), h: 0.17, rectRadius: 0.03, fill: { color: col }, line: { type: "none" } });
      s.addText(M(r[k]), { x: bx + 0.94, y, w: 0.50, h: 0.26, margin: 0, align: "right", valign: "middle", fontFace: F, fontSize: 8, bold: true, color: col });
    });
  });
  s.addText(`Active at AIS  ${M(aisReach)}`, { x: 9.82, y: 5.40, w: 1.44, h: 0.20, margin: 0, fontFace: F, fontSize: 8.5, bold: true, color: R4 });
  s.addText(`Active at KTB  ${M(ktbReach)}`, { x: 11.28, y: 5.40, w: 1.44, h: 0.20, margin: 0, fontFace: F, fontSize: 8.5, bold: true, color: R2 });

  card(s, 0.58, 5.78, 12.33, 0.86, MINT);
  s.addText([
    { text: `AIS & KTB and AIS KTB OR hold ${p0(prioClicx, onClicx)} of CLICX customers from ${p0(prioBase, noClicx + onClicx)} of the base`, options: { bold: true, color: TITLE_TEAL } },
    { text: `  —  they over-index on adoption, on app-in and on deposit, and they are still active at AIS or KTB. These two segments are where acquisition should start.`, options: { color: INK } },
  ], { x: 0.88, y: 5.78, w: 11.73, h: 0.86, margin: 0, valign: "middle", fontFace: F, fontSize: 11 });

  s.addNotes(`Captive base of ${M(CAPTIVE)} as stated; the seven segment rows sum to ${M(noClicx)} without CLICX plus ${M(onClicx)} with, so there is roughly 0.7M of rounding or residual between the two. ` +
    `CLICX reaches ${p1(onClicx, noClicx + onClicx)} of the consortium base. 97% of CLICX customers are captive, and 46K sit outside it. ` +
    `AIS & KTB and AIS KTB OR together are ${p0(prioClicx, onClicx)} of CLICX customers, ${p0(prioBase, noClicx + onClicx)} of the base, and the same two segments lead on app-in and deposit. ` +
    `Reach is the practical point: ${M(aisReach)} non-CLICX customers are active at AIS and ${M(ktbReach)} at KTB, and the AIS & KTB and AIS KTB OR segments appear in both columns, so they can be approached through either partner. ` +
    "One cell to check at source: OR Only shows 11K CLICX customers but 16K deposit customers, which cannot both be right as a subset. Small cell, likely rounding, but worth confirming.");
}

/* ============================================================== SLIDE 2 */
{
  const s = pres.addSlide(); s.background = { color: PAPER };
  const pen = (n) => { const r = SEG.find((x) => x.segment === n); return r.have_clicx / (r.no_clicx + r.have_clicx); };
  /* segment names differ between the three source tables, so map explicitly
     rather than deriving them - a silent miss here reads as NaN on the slide */
  const pairs = [
    { a: "AIS only", b: "AIS & OR", la: "AIS", lb: "AIS & OR", ca: "AIS", cb: "AIS & OR" },
    { a: "KTB only", b: "KTB OR", la: "KTB", lb: "KTB & OR", ca: "KTB", cb: "KTB & OR" },
    { a: "AIS & KTB", b: "AIS KTB OR", la: "AIS & KTB", lb: "AIS & KTB & OR", ca: "AIS & KTB", cb: "AIS & KTB x OR" },
  ];
  const pick = (A, n) => { const r = A.find((x) => x.segment === n); if (!r) throw new Error("no row for " + n); return r; };

  head(s, "An OR relationship multiplies CLICX adoption",
    "*Penetration = CLICX customers ÷ segment base. Cash-need intensity is the medium-high share of the non-CLICX segment.",
    `With OR ${M(orBase)} base → ${M(orClicx)} on CLICX (${p1(orClicx, orBase)})   ·   Without OR ${M(nonBase)} → ${M(nonClicxOR)} (${p1(nonClicxOR, nonBase)})   ·   ${(penOR / penNo).toFixed(1)}× lift`,
    "The same lift shows up three times over: adoption, loan booking, and cash-need intensity.");

  tiles(s, [
    [p1(orClicx, orBase), "penetration with OR", `${M(orClicx)} of ${M(orBase)}`, TEAL],
    [p1(nonClicxOR, nonBase), "penetration without OR", `${M(nonClicxOR)} of ${M(nonBase)}`, R1],
    [(pen("AIS & OR") / pen("AIS only")).toFixed(1) + "×", "AIS only → AIS & OR", `${(pen("AIS only") * 100).toFixed(2)}% → ${(pen("AIS & OR") * 100).toFixed(2)}%`, TEAL],
    [p1(bookOR * 100, 100), "loan booking with OR", `against ${p1(bookNo * 100, 100)} without`, TEAL],
  ], 1.80);

  card(s, 0.58, 3.06, 6.04, 2.54);
  cardHead(s, 0.58, 3.06, 6.04, "CLICX penetration by segment", "Segments carrying an OR relationship in teal");
  const pMax = Math.max(...SEG.map((r) => r.have_clicx / (r.no_clicx + r.have_clicx)));
  SEG.forEach((r, i) => {
    const v = r.have_clicx / (r.no_clicx + r.have_clicx);
    barRow(s, 3.78 + i * 0.25, 0.86, 1.45, 2.40, 2.35, 4.86, 0.90, r.segment, v / pMax, hasOR(r) ? R3 : R1, (v * 100).toFixed(2) + "%");
  });


  card(s, 6.87, 3.06, 6.04, 2.54);
  cardHead(s, 6.87, 3.06, 6.04, "The same lift, three ways", "Adding OR to a segment, measured on three outcomes");
  s.addTable([
    [th("Pair", "left"), th("Penetration"), th("Lift"), th("Booking rate"), th("M-H cash need")],
    ...pairs.map(({ a, b, la: lan, lb: lbn, ca: can, cb: cbn }) => {
      const la = pick(LEND, lan), lb = pick(LEND, lbn);
      const ca = pick(CASH, can), cb = pick(CASH, cbn);
      return [
        rh(a + " → " + b),
        td(`${(pen(a) * 100).toFixed(1)}% → ${(pen(b) * 100).toFixed(1)}%`, "right", R3, true),
        td((pen(b) / pen(a)).toFixed(1) + "×", "right", R3, true),
        td(`${p0(la.booked, la.appin)} → ${p0(lb.booked, lb.appin)}`, "right", R3, true),
        td(`${p0(ca.mh_cash, ca.mh_cash + ca.low_cash)} → ${p0(cb.mh_cash, cb.mh_cash + cb.low_cash)}`, "right", R3, true),
      ];
    }),
  ], {
    x: 7.11, y: 3.84, w: 5.56, colW: [1.70, 1.22, 0.60, 1.00, 1.04],
    rowH: [0.26, 0.34, 0.34, 0.34],
    border: { type: "solid", color: RULE, pt: 0.5 }, valign: "middle", margin: [0.02, 0.06, 0.02, 0.06],
  });
  card(s, 7.15, 5.18, 5.48, 0.34, MINT);
  s.addText("OR presence moves all three in the same direction, every time.", {
    x: 7.31, y: 5.18, w: 5.16, h: 0.34, margin: 0, valign: "middle", fontFace: F, fontSize: 9.5, bold: true, color: TITLE_TEAL,
  });

  card(s, 0.58, 5.78, 12.33, 0.86, MINT);
  s.addText([
    { text: `OR-linked customers adopt CLICX at ${p1(orClicx, orBase)} against ${p1(nonClicxOR, nonBase)}, and book loans at ${p1(bookOR * 100, 100)} against ${p1(bookNo * 100, 100)}`, options: { bold: true, color: TITLE_TEAL } },
    { text: "  —  the OR relationship is the strongest single predictor in this data. Lead acquisition through the OR channel and prioritise OR-linked segments in every campaign.", options: { color: INK } },
  ], { x: 0.88, y: 5.78, w: 11.73, h: 0.86, margin: 0, valign: "middle", fontFace: F, fontSize: 11 });

  s.addNotes(`Penetration with an OR relationship is ${p1(orClicx, orBase)} against ${p1(nonClicxOR, nonBase)} without, a ${(penOR / penNo).toFixed(1)}x lift on a base of ${M(orBase)} against ${M(nonBase)}. ` +
    "Pairwise the lift is 6.6x for AIS only to AIS & OR, 2.6x for KTB only to KTB OR, and 2.0x for AIS & KTB to AIS KTB OR. " +
    `The same direction holds on loan booking - ${p1(bookOR * 100, 100)} with OR against ${p1(bookNo * 100, 100)} without - and on cash-need intensity among non-CLICX customers, where the medium-high share rises in every pair. ` +
    "Caveat on causality: this is association, not proof. OR customers may simply be a more digitally engaged population, so the lift would partly survive without any OR intervention. The practical read is that OR is an excellent targeting signal regardless of whether it is causal.");
}

/* ============================================================== SLIDE 3 */
{
  const s = pres.addSlide(); s.background = { color: PAPER };
  const worst = LEND.reduce((a, r) => (r.booked / r.appin < a.booked / a.appin ? r : a));
  const best = LEND.reduce((a, r) => (r.booked / r.appin > a.booked / a.appin ? r : a));
  const aisKtb = CASH.find((r) => r.segment === "AIS & KTB");

  head(s, `Lending : ${M(appin)} app-ins convert to ${M(booked)} booked loans`,
    "*App-in, booked and drawdown as supplied, rounded to thousands. Cash need is derived from AIS and KTB data for customers not on CLICX.",
    `App-in ${M(appin)}   ·   Booked ${M(booked)} (${p1(booked, appin)})   ·   Drawn ${M(drawn)} (${p1(drawn, booked)})   ·   Cash-need pool outside CLICX ${M(mh)}`,
    `Drawdown is uniformly high at ${p1(drawn, booked)}. The leak is at booking, and it is a segment problem, not a product one.`);

  tiles(s, [
    [M(appin), "loan app-ins", `${p1(appin, onClicx)} of CLICX customers`, TEAL],
    [M(booked), "booked", `${p1(booked, appin)} of app-ins`, AMBER],
    [M(drawn), "drawn down", `${p1(drawn, booked)} of booked`, TEAL],
    [M(mh), "medium-high cash need", "outside CLICX · the pool", RED],
  ], 1.80);

  card(s, 0.58, 3.06, 6.04, 2.54);
  cardHead(s, 0.58, 3.06, 6.04, "Booking rate by segment", "Booked ÷ app-in · OR-linked segments in teal");
  const bMax = Math.max(...LEND.map((r) => r.booked / r.appin));
  LEND.forEach((r, i) => {
    const v = r.booked / r.appin;
    barRow(s, 3.78 + i * 0.25, 0.86, 1.70, 2.66, 2.00, 4.86, 0.90, r.segment, v / bMax, hasOR(r) ? R3 : AMBER,
      (v * 100).toFixed(1) + "%");
  });


  card(s, 6.87, 3.06, 6.04, 2.54);
  cardHead(s, 6.87, 3.06, 6.04, "Cash need outside CLICX", "Customers not on CLICX, THB-free sizing in customers");
  s.addTable([
    [th("Segment", "left"), th("M-H cash need"), th("Low"), th("M-H share")],
    ...CASH.map((r) => {
      const t = r.mh_cash + r.low_cash, h = r.segment === "AIS & KTB";
      return [rh(r.segment, h),
        td(M(r.mh_cash), "right", h ? TITLE_TEAL : INK, true),
        td(M(r.low_cash), "right", MUTED),
        td(p0(r.mh_cash, t), "right", r.mh_cash / t > 0.4 ? R3 : INK, true)];
    }),
    [rh("Total"), ...[M(mh), M(lowc), p0(mh, mh + lowc)].map((t) => ({
      text: t, options: { fill: { color: CARD }, color: INK, bold: true, fontFace: F, fontSize: 8.5, align: "right" },
    }))],
  ], {
    x: 7.11, y: 3.80, w: 5.56, colW: [1.86, 1.35, 1.10, 1.25],
    rowH: [0.24, 0.21, 0.21, 0.21, 0.21, 0.21, 0.21, 0.21, 0.23],
    border: { type: "solid", color: RULE, pt: 0.5 }, valign: "middle", margin: [0.01, 0.06, 0.01, 0.06],
  });

  card(s, 0.58, 5.78, 12.33, 0.86, WARM);
  s.addText([
    { text: `${M(aisKtb.mh_cash)} AIS & KTB customers outside CLICX have medium-high cash need`, options: { bold: true, color: RED } },
    { text: `  —  the largest single lending pool in the base, and the same segment that already supplies ${p0(LEND.find((r) => r.segment === "AIS & KTB").appin, appin)} of app-ins. Fixing ${worst.segment} booking at ${p1(worst.booked, worst.appin)} is the other half of the prize.`, options: { color: INK } },
  ], { x: 0.88, y: 5.78, w: 11.73, h: 0.86, margin: 0, valign: "middle", fontFace: F, fontSize: 11 });

  s.addNotes(`Funnel: ${M(appin)} app-ins, ${M(booked)} booked (${p1(booked, appin)}), ${M(drawn)} drawn down (${p1(drawn, booked)} of booked). ` +
    "Drawdown holds between 83% and 100% across every segment, so once a loan is booked it is almost always taken. The variance is all at booking. " +
    `Booking runs from ${p1(worst.booked, worst.appin)} for ${worst.segment} to ${p1(best.booked, best.appin)} for ${best.segment}, and every OR-linked segment books above 22% while KTB-only books at 8.8%. ` +
    `Outside CLICX, ${M(mh)} customers show medium-high cash need against ${M(lowc)} low, and ${M(aisKtb.mh_cash)} of the medium-high sit in AIS & KTB alone. AIS & KTB x OR has the highest intensity at 60%. ` +
    "Two actions follow: target AIS & KTB for app-in volume, and investigate why KTB-only applicants fail at booking - credit policy, data quality or journey - since the drawdown rate shows demand is real.");
}

/* ============================================================== SLIDE 4 */
{
  const s = pres.addSlide(); s.background = { color: PAPER };
  const penClicx = CLICX_TMW / CLICX_TOTAL, penBase = BASE_TMW / BASE_TOTAL;
  const ktbOnly = baseT.find((r) => r.segment === "KTB only");

  head(s, "755K CLICX customers already hold a TrueMoney wallet",
    "*TMW = TrueMoney Wallet holding, from AIS and KTB data. TrueMoney VB = the TrueMoney virtual bank.",
    `CLICX with TMW ${M(CLICX_TMW)} (${p0(CLICX_TMW, CLICX_TOTAL)})   ·   Deposit at risk ${M(dep.tmw)}   ·   Payment at risk ${M(pay.tmw)}   ·   Outside CLICX ${M(BASE_TMW)} of ${M(BASE_TOTAL)}`,
    "CLICX's own base is twice as wallet-penetrated as the base it has not reached — the installed base is the exposed one.");

  tiles(s, [
    [M(dep.tmw), "deposit customers at risk", `${p0(dep.tmw, dep.total)} of the CLICX deposit base`, RED],
    [M(pay.tmw), "payment customers at risk", `${p0(pay.tmw, pay.total)} — most exposed product`, RED],
    [p0(CLICX_TMW, CLICX_TOTAL), "of CLICX hold TMW", `${M(CLICX_TMW)} of ${M(CLICX_TOTAL)}`, RED],
    [p0(BASE_TMW, BASE_TOTAL), "outside CLICX hold TMW", `${M(BASE_TMW)} of ${M(BASE_TOTAL)}`, AMBER],
  ], 1.80);

  card(s, 0.58, 3.06, 6.04, 2.54);
  cardHead(s, 0.58, 3.06, 6.04, "Exposure inside CLICX", "By product · share of each product base holding a wallet");
  clicxT.forEach((r, i) => {
    const y = 3.86 + i * 0.44;
    s.addText(r.segment, { x: 0.86, y, w: 1.70, h: 0.30, margin: 0, valign: "middle", fontFace: F, fontSize: 10.5, bold: true, color: INK });
    s.addText(M(r.total), { x: 2.60, y, w: 0.70, h: 0.30, margin: 0, align: "right", valign: "middle", fontFace: F, fontSize: 10, color: MUTED });
    s.addShape("roundRect", { x: 3.44, y: y + 0.06, w: 1.70, h: 0.18, rectRadius: 0.03, fill: { color: "E4E9F0" }, line: { type: "none" } });
    s.addShape("roundRect", { x: 3.44, y: y + 0.06, w: 1.70 * (r.tmw / r.total), h: 0.18, rectRadius: 0.03, fill: { color: RED }, line: { type: "none" } });
    s.addText(`${M(r.tmw)}  ${p0(r.tmw, r.total)}`, { x: 5.22, y, w: 1.12, h: 0.30, margin: 0, align: "right", valign: "middle", fontFace: F, fontSize: 10, bold: true, color: RED });
  });
  card(s, 0.86, 5.14, 5.48, 0.36, WARM);
  s.addText(`Payment is the most exposed at ${p0(pay.tmw, pay.total)}; deposit is the largest at ${M(dep.tmw)} customers.`, {
    x: 1.02, y: 5.14, w: 5.16, h: 0.36, margin: 0, valign: "middle", fontFace: F, fontSize: 9.5, color: INK,
  });

  card(s, 6.87, 3.06, 6.04, 2.54);
  cardHead(s, 6.87, 3.06, 6.04, "Wallet penetration outside CLICX", "By mobile relationship · the acquisition battleground");
  const tMax = Math.max(...baseT.map((r) => r.tmw / r.total));
  baseT.forEach((r, i) => {
    const v = r.tmw / r.total;
    barRow(s, 3.84 + i * 0.32, 7.15, 1.70, 9.00, 2.10, 11.20, 1.43, r.segment, v / tMax, v > 0.25 ? RED : R3,
      `${M(r.tmw)} of ${M(r.total)}  ${p0(r.tmw, r.total)}`, 9);
  });
  s.addText("Wallet penetration tracks the mobile relationship, not the bank one.", {
    x: 7.15, y: 5.38, w: 5.48, h: 0.20, margin: 0, fontFace: F, fontSize: 9, italic: true, color: MUTED,
  });

  card(s, 0.58, 5.78, 12.33, 0.86, MINT);
  s.addText([
    { text: `KTB-only customers are ${100 - Math.round(ktbOnly.tmw / ktbOnly.total * 100)}% wallet-free — ${M(ktbOnly.no_tmw)} people TrueMoney has not reached either`, options: { bold: true, color: TITLE_TEAL } },
    { text: `  —  the cleanest acquisition pool in the base. Defend the ${M(dep.tmw)} deposit customers who already hold a wallet, and win the ones where no wallet relationship exists yet.`, options: { color: INK } },
  ], { x: 0.88, y: 5.78, w: 11.73, h: 0.86, margin: 0, valign: "middle", fontFace: F, fontSize: 11 });

  s.addNotes(`${M(CLICX_TMW)} of ${M(CLICX_TOTAL)} CLICX customers, ${p0(CLICX_TMW, CLICX_TOTAL)}, already hold a TrueMoney wallet. By product: deposit ${M(dep.tmw)} (${p0(dep.tmw, dep.total)}), lending ${M(clicxT[1].tmw)} (${p0(clicxT[1].tmw, clicxT[1].total)}), payment ${M(pay.tmw)} (${p0(pay.tmw, pay.total)}). ` +
    "Payment is proportionally the most exposed, deposit the largest in absolute terms and the one carrying balance. " +
    `Outside CLICX the rate is ${p0(BASE_TMW, BASE_TOTAL)}, so CLICX's own customers are roughly twice as wallet-penetrated as the population it has not acquired. That is uncomfortable: the installed base is the exposed one. ` +
    "Penetration outside CLICX tracks the mobile relationship rather than the bank one - 55.6% for customers with both pre and post paid, 43.8% post-paid, 27.1% pre-paid, against 6.8% for KTB-only and 4.0% non-mobile. " +
    "The strategic read is two-sided: defend the dual-holders before the virtual bank launches, and prioritise KTB-only and non-mobile customers for acquisition because no wallet relationship competes there yet.");
}

/* ============================================================== SLIDE 5 */
{
  const s = pres.addSlide(); s.background = { color: PAPER };
  const aisKtb = CASH.find((r) => r.segment === "AIS & KTB");
  const ktbOnly = baseT.find((r) => r.segment === "KTB only");
  const worst = LEND.reduce((a, r) => (r.booked / r.appin < a.booked / a.appin ? r : a));

  head(s, "Where to play",
    "*Sizing uses the segments as supplied. Each priority is drawn from a different one of the four source views.",
    `Defend ${M(CLICX_TMW)}   ·   Acquire through OR ${M(orBase)} at ${p1(orClicx, orBase)} penetration   ·   Convert ${M(aisKtb.mh_cash)} AIS & KTB with cash need`,
    "One defensive priority and two offensive ones, in the order the data supports.");

  const PANELS = [
    {
      x: 0.58, num: "1", kicker: "DEFEND", col: RED, title: "The dual-holders",
      big: M(CLICX_TMW), bigSub: `CLICX customers who already hold a TrueMoney wallet (${p0(CLICX_TMW, CLICX_TOTAL)})`,
      rows: [
        ["Deposit at risk", `${M(dep.tmw)} customers, ${p0(dep.tmw, dep.total)} of the deposit base`],
        ["Payment at risk", `${M(pay.tmw)} customers, ${p0(pay.tmw, pay.total)} — the most exposed product`],
        ["Act before launch", "Deepen the deposit relationship while the wallet is still a wallet, not a bank."],
      ],
    },
    {
      x: 4.77, num: "2", kicker: "ACQUIRE", col: TITLE_TEAL, title: "OR-linked segments",
      big: p1(orClicx, orBase), bigSub: `CLICX penetration with an OR relationship, against ${p1(nonClicxOR, nonBase)} without`,
      rows: [
        ["Strongest signal", "OR presence lifts adoption 2× to 7×, booking to 24.5%, and cash-need intensity."],
        ["Start here", `AIS KTB OR converts at 8.2% — the highest of any segment, on a ${M(5.2e6)} base.`],
        ["Channel", "Lead through the OR touchpoint; it reaches customers AIS and KTB alone do not."],
      ],
    },
    {
      x: 8.96, num: "3", kicker: "CONVERT", col: AMBER, title: "Cash need in AIS & KTB",
      big: M(aisKtb.mh_cash), bigSub: "non-CLICX AIS & KTB customers with medium-high cash need",
      rows: [
        ["Largest pool", `${p0(aisKtb.mh_cash, mh)} of all medium-high cash need outside CLICX.`],
        ["Reachable", `Part of the ${M(aisReach)} active at AIS and ${M(ktbReach)} active at KTB.`],
        ["Fix booking", `${worst.segment} books at ${p1(worst.booked, worst.appin)} against 26% for KTB & OR — demand is real, the funnel is not.`],
      ],
    },
  ];

  PANELS.forEach((pn) => {
    const X = pn.x, IN = X + 0.26, W = 3.42;
    card(s, X, 1.80, 3.94, 4.44);
    s.addShape("roundRect", { x: IN, y: 1.94, w: 0.30, h: 0.28, rectRadius: 0.05, fill: { color: pn.col }, line: { type: "none" } });
    s.addText(pn.num, { x: IN, y: 1.94, w: 0.30, h: 0.28, margin: 0, align: "center", valign: "middle", fontFace: F, fontSize: 11, bold: true, color: PAPER });
    s.addText(pn.kicker, { x: IN + 0.42, y: 1.94, w: W - 0.42, h: 0.28, margin: 0, valign: "middle", fontFace: F, fontSize: 9.5, bold: true, charSpacing: 1.6, color: pn.col });
    s.addText(pn.title, { x: IN, y: 2.30, w: W, h: 0.30, margin: 0, fontFace: FH, fontSize: 14, bold: true, color: INK });
    s.addText(pn.big, { x: IN, y: 2.66, w: W, h: 0.50, margin: 0, valign: "middle", fontFace: F, fontSize: 28, bold: true, color: pn.col });
    s.addText(pn.bigSub, { x: IN, y: 3.18, w: W, h: 0.42, margin: 0, valign: "top", fontFace: F, fontSize: 9, color: MUTED });
    s.addShape("rect", { x: IN, y: 3.68, w: W, h: 0.012, fill: { color: RULE }, line: { type: "none" } });
    pn.rows.forEach(([t, d], i) => {
      const y = 3.80 + i * 0.78;
      s.addText(t, { x: IN, y, w: W, h: 0.24, margin: 0, fontFace: F, fontSize: 10.5, bold: true, color: pn.col });
      s.addText(d, { x: IN, y: y + 0.22, w: W, h: 0.54, margin: 0, valign: "top", fontFace: F, fontSize: 9, color: INK });
    });
  });

  card(s, 0.58, 6.38, 12.33, 0.56, MINT);
  s.addText([
    { text: "Defend 755K, then acquire where OR already works", options: { bold: true, color: TITLE_TEAL } },
    { text: `  —  CLICX reaches ${p1(onClicx, noClicx + onClicx)} of its own consortium base, and the part it has reached is the part most exposed to TrueMoney. The growth and the risk sit in the same place.`, options: { color: INK } },
  ], { x: 0.88, y: 6.38, w: 11.73, h: 0.56, margin: 0, valign: "middle", fontFace: F, fontSize: 10.5 });

  s.addNotes("Three priorities, each drawn from a different source view, ordered by urgency rather than size. " +
    `Defend first: ${M(CLICX_TMW)} CLICX customers already hold a TrueMoney wallet, ${M(dep.tmw)} of them on deposit. That exposure is live the day the virtual bank launches, and it sits with customers already acquired, so the cost of losing them is sunk acquisition plus balance. ` +
    `Acquire second: the OR relationship is the strongest single predictor in the data, lifting penetration from ${p1(nonClicxOR, nonBase)} to ${p1(orClicx, orBase)}, booking from 16.0% to 24.5%, and cash-need intensity in every pairwise comparison. ` +
    `Convert third: ${M(aisKtb.mh_cash)} AIS & KTB customers outside CLICX carry medium-high cash need, ${p0(aisKtb.mh_cash, mh)} of the total pool, and they are reachable through either partner. ` +
    "The uncomfortable summary is that growth and risk overlap: the 3.5% of the consortium base CLICX has reached is twice as wallet-penetrated as the 96.5% it has not.");
}

const out = path.join(__dirname, "CLICX_Captive_Opportunity.pptx");
pres.writeFile({ fileName: out }).then(() => console.log("wrote", out));
