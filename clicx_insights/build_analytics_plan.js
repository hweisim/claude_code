// CLICX board storyline v3 — situation, complication, resolution.
// Data: "DRAFT 20261105 - Data for Strategy 2027", Partner Payment Story v1, product segment extract (31 Oct 2026).
const pptxgen = require("pptxgenjs");
const { applyTheme } = require(process.env.SKILL + "/scripts/apply_theme.js");

const OUT = process.argv[2] || "CLICX_Analytics_Plan.pptx";

const THEME = {
  name: "CLICX Blue Grey Teal",
  headFontFace: "Arial",
  bodyFontFace: "Arial",
  colors: {
    dk1: "1C2733", lt1: "FFFFFF", dk2: "1F3A5F", lt2: "EEF2F6",
    accent1: "2E6DB4", accent2: "1AA39A", accent3: "1F3A5F",
    accent4: "8A97A6", accent5: "5B6B7C", accent6: "9FD8D3",
    hlink: "2E6DB4", folHlink: "5B6B7C",
  },
};
const NAVY = "1F3A5F", BLUE = "2E6DB4", CYAN = "1AA39A", GREY = "8A97A6", LGREY = "D3D9E0", LCYAN = "9FD8D3",
  TXT = "1C2733", MUTED = "5B6B7C", PANELHEX = "EEF2F6", TINT = "E5F1F0", SOFT = "8FB3DA", ORANGE = "E3A15F",
  TRUEC = "D9822B", WHITE = "FFFFFF";
const MOVE = { ACQUIRE: NAVY, ACTIVATE: CYAN, GROW: BLUE, DEFEND: TRUEC };

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "CLICX analytics plan";
pres.company = "CLICX";
const C = pres.SchemeColor;
const IC = (f) => __dirname + "/icons/" + f;

// ---------- Layouts ----------
pres.defineSlideMaster({
  title: "TITLE",
  background: { color: C.text2 },
  objects: [
    { placeholder: { options: { name: "kicker", type: "body", x: 0.8, y: 2.0, w: 11.5, h: 0.4, fontSize: 14, bold: true, color: C.accent2, charSpacing: 2 }, text: "" } },
    { placeholder: { options: { name: "title", type: "title", x: 0.8, y: 2.45, w: 11.5, h: 1.4, fontSize: 40, bold: true, color: C.background1, align: "left", valign: "top" }, text: "" } },
    { placeholder: { options: { name: "sub", type: "body", x: 0.8, y: 3.95, w: 10.5, h: 1.0, fontSize: 18, color: C.accent6, valign: "top" }, text: "" } },
    { placeholder: { options: { name: "date", type: "body", x: 0.8, y: 6.3, w: 8, h: 0.4, fontSize: 12, color: C.background1 }, text: "" } },
    { image: { path: __dirname + "/logo_white.png", x: 11.13, y: 0.5, w: 1.7, h: 0.426, altText: "CLICX logo" } },
  ],
});
pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: C.background1 },
  margin: [0.5, 0.5, 0.6, 0.5],
  objects: [
    { placeholder: { options: { name: "tracker", type: "body", x: 0.5, y: 0.3, w: 5.6, h: 0.3, fontSize: 11, bold: true, color: C.accent1, charSpacing: 1, margin: 0 }, text: "" } },
    { placeholder: { options: { name: "title", type: "title", x: 0.5, y: 0.62, w: 12.33, h: 0.9, fontSize: 24, bold: true, color: C.text2, align: "left", valign: "top", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "foot", type: "body", x: 0.5, y: 6.9, w: 11.2, h: 0.32, fontSize: 9.5, color: "5B6B7C", valign: "top", margin: 0 }, text: "" } },
    { image: { path: __dirname + "/logo.png", x: 11.63, y: 0.2, w: 1.2, h: 0.3, altText: "CLICX logo" } },
  ],
  slideNumber: { x: 12.33, y: 6.95, w: 0.5, h: 0.3, fontSize: 10, color: "5B6B7C", align: "right" },
});

// ---------- Helpers ----------
const T = (s, t, o) => s.addText(t, { margin: 0, isTextBox: true, ...o });
const R = (s, x, y, w, h, fill, name) => s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: fill }, line: { type: "none" }, objectName: name });
const RR = (s, x, y, w, h, fill, name, r = 0.08) => s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: r, fill: { color: fill }, line: { type: "none" }, objectName: name });
const badge = (s, x, y, n, col = NAVY, d = 0.36, fs = 12) => {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: col }, line: { type: "none" }, objectName: `Badge ${n} bg` });
  T(s, String(n), { x, y, w: d, h: d, fontSize: fs, bold: true, color: WHITE, align: "center", valign: "middle", objectName: `Badge ${n}` });
};
const circIcon = (s, x, y, d, icon, col, name) => {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: col }, line: { type: "none" }, objectName: name + " bg" });
  s.addImage({ path: IC(icon), x: x + d * 0.25, y: y + d * 0.25, w: d * 0.5, h: d * 0.5, altText: name, objectName: name });
};
const head = (s, x, y, w, n, title, sub, col = NAVY) => {
  if (n) badge(s, x + 0.18, y + 0.16, n, col);
  const hx = n ? x + 0.66 : x + 0.2;
  T(s, title, { x: hx, y: y + 0.11, w: w - (hx - x) - 0.15, h: 0.46, fontSize: 13.5, bold: true, color: NAVY, valign: "middle", objectName: `Head ${title.slice(0, 14)}` });
  if (sub) T(s, sub, { x: x + 0.2, y: y + 0.62, w: w - 0.4, h: 0.24, fontSize: 9.5, color: MUTED, objectName: `Sub ${title.slice(0, 14)}` });
};
const soWhat = (s, text, y = 6.3) => {
  RR(s, 0.5, y, 12.33, 0.46, NAVY, "So what");
  T(s, [{ text: "SO WHAT  ", options: { bold: true, fontSize: 9.5, color: LCYAN, charSpacing: 1 } }, { text, options: { fontSize: 11.5, color: WHITE } }],
    { x: 0.7, y, w: 11.93, h: 0.46, valign: "middle", objectName: "So what text" });
};
const legend = (s, x, y, items, fs = 9.5) => {
  let cx = x;
  items.forEach((it, k) => {
    R(s, cx, y + 0.05, 0.15, 0.15, it[1], `Legend key ${it[0].slice(0, 10)}`);
    const w = it[0].length * fs * 0.0078 + 0.12;
    T(s, it[0], { x: cx + 0.2, y, w, h: 0.25, fontSize: fs, color: MUTED, valign: "middle", objectName: `Legend ${it[0].slice(0, 10)}` });
    cx += 0.2 + w + 0.15;
  });
};
// big number + sentence
const numRow = (s, x, y, w, h, num, text, o = {}) => {
  const nw = o.nw || 1.3;
  T(s, num, { x, y, w: nw, h, fontSize: o.nfs || 20, bold: true, color: o.col || BLUE, valign: "middle", objectName: `Num ${num}` });
  T(s, text, { x: x + nw + 0.1, y, w: w - nw - 0.1, h, fontSize: o.fs || 10.5, color: TXT, valign: "middle", objectName: `Num text ${num}` });
};
// drawn horizontal bars; rows = [label, value, valueText, color, note]
function hbars(s, x, y, w, rows, o = {}) {
  const lw = o.lw || 1.8, vw = o.vw || 0.55, rh = o.rh || 0.4, max = o.max || 1, bw = w - lw - vw, fs = o.fs || 10.5, nm = o.name || "Bar";
  if (o.bench) {
    const bx = x + lw + bw * o.bench.v / max;
    s.addShape(pres.shapes.LINE, { x: bx, y: y - 0.04, w: 0, h: rows.length * rh + 0.08, line: { color: MUTED, width: 1, dashType: "dash" }, objectName: `${nm} bench` });
    T(s, o.bench.label, { x: bx - 1.0, y: y + rows.length * rh + 0.05, w: 2.0, h: 0.2, fontSize: 8.5, color: MUTED, align: "center", objectName: `${nm} bench label` });
  }
  rows.forEach((r, k) => {
    const yy = y + k * rh, ww = Math.max(0.03, bw * r[1] / max);
    T(s, r[0], { x, y: yy, w: lw - 0.1, h: rh, fontSize: fs, color: TXT, valign: "middle", objectName: `${nm} label ${k}` });
    R(s, x + lw, yy + rh * 0.2, ww, rh * 0.6, r[3], `${nm} ${k}`);
    const vo = { x: x + lw + ww + 0.05, y: yy + rh * 0.16, w: o.valW || 0.5, h: rh * 0.68, fontSize: fs, bold: true, color: NAVY, valign: "middle", objectName: `${nm} value ${k}` };
    if (o.labelFill) vo.fill = { color: o.labelFill };
    T(s, r[2], vo);
    if (r[4]) T(s, r[4], { x: x + lw + ww + 0.05 + (o.valW || 0.5), y: yy, w: 1.9, h: rh, fontSize: 8.5, italic: true, color: MUTED, valign: "middle", objectName: `${nm} note ${k}` });
  });
}
// drawn clustered columns; series = [{ name, color, vals }]
function ccols(s, x, y, w, h, cats, series, o = {}) {
  const max = o.max || 1, gw = w / cats.length, bwid = gw * 0.78 / series.length, fs = o.fs || 8.5, nm = o.name || "Col";
  s.addShape(pres.shapes.LINE, { x, y: y + h, w, h: 0, line: { color: LGREY, width: 1 }, objectName: `${nm} axis` });
  cats.forEach((c, i) => {
    series.forEach((se, j) => {
      const v = se.vals[i], bh = Math.max(0.012, h * v / max), bx = x + i * gw + gw * 0.11 + j * bwid;
      R(s, bx, y + h - bh, bwid - 0.02, bh, se.color, `${nm} ${i}-${j}`);
      T(s, o.fmt ? o.fmt(v) : String(v), { x: bx - 0.15, y: y + h - bh - 0.2, w: bwid + 0.28, h: 0.19, fontSize: fs, color: TXT, align: "center", valign: "bottom", objectName: `${nm} val ${i}-${j}` });
    });
    T(s, c, { x: x + i * gw, y: y + h + 0.03, w: gw, h: 0.22, fontSize: o.cfs || 9.5, color: MUTED, align: "center", objectName: `${nm} cat ${i}` });
  });
}
// chapter navigator + slide factory
const CH = ["WHAT WE BUILT", "WHAT HOLDS US BACK", "THE THREAT", "THE PLAN"], CHW = [1.22, 1.6, 0.95, 0.85];
function mk(ch, tracker, title, foot, section = "Storyline") {
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: section });
  s.addText(tracker, { placeholder: "tracker" });
  s.addText(title, { placeholder: "title" });
  if (foot) s.addText(foot, { placeholder: "foot" });
  if (ch !== null) {
    let x = 11.4 - (CHW.reduce((a, b) => a + b, 0) + 3 * 0.07);
    CH.forEach((c, k) => {
      const on = k === ch;
      s.addText(c, { x, y: 0.3, w: CHW[k], h: 0.25, shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.12, fill: { color: on ? NAVY : PANELHEX }, line: { type: "none" },
        fontSize: 7.5, bold: on, color: on ? WHITE : GREY, align: "center", valign: "middle", margin: 0, charSpacing: 0.5, objectName: `Chapter ${k + 1}` });
      x += CHW[k] + 0.07;
    });
  }
  return s;
}
// =====================================================================
// Analytics plan for three management asks: saving pocket, AIS top-up / bill pay, mobile financing
const SRCP = "Figures from the CLICX segment extract (1 Oct 2026), Partner Payment Story v1 and Data for Business Strategy (draft, 5 Nov 2026). Timings are a proposal and depend on the AIS data requests";
pres.addSection({ title: "Analytics plan" });

const WS = [
  { n: 1, tag: "URGENT · BEFORE TRUE'S LAUNCH", col: TRUEC, name: "AIS saving pocket", sub: "Free AIS package above a balance threshold",
    q: ["Who would it reach, and what deposits would it win or protect?", "Which threshold pays for itself?", "Which non-AIS customers could convert to AIS?"],
    know: [["฿1.98B", "63% of deposits sit with 1.26M True and TrueMoney customers"], ["1.9%", "of customers hold over ฿20K; 76% hold ฿0"]],
    a: ["Threshold scenarios (฿5K–50K): who qualifies, deposits gained, package cost vs margin",
      "Size three groups: AIS + TrueMoney (retain), non-AIS (convert), AIS only (protect)",
      "Rank AIS + TrueMoney customers by deposits at risk",
      "Pilot design with a control group and a 4–6 week read"],
    out: "Go/no-go threshold table and a ranked offer list", when: "Weeks 1–2" },
  { n: 2, tag: "HIGH", col: CYAN, name: "AIS top-up and bill pay", sub: "From 20th in market share to top of mind",
    q: ["What is our share of AIS top-up and bill pay?", "Where do our customers top up instead?", "Which segments to target, with what offer?"],
    know: [["8.4% vs 12.8%", "of AIS prepaid vs postpaid customers pay with us"], ["91K → 59K", "AIS payers after the Aug–Sep promotion ended"]],
    a: ["Market share baseline by channel and bank, and our share of our own AIS customers' top-ups",
      "Segment by plan, top-up habit, current channel, age, job, region and CLICX status",
      "Size targets: ~470K prepaid non-payers, ~32K lapsed AIS payers, borrowers on AIS, TrueMoney top-up users",
      "Find what keeps people paying, then test offers against control groups"],
    out: "Segment map, target lists and a campaign test plan", when: "Weeks 2–4" },
  { n: 3, tag: "MEDIUM", col: BLUE, name: "Mobile financing", sub: "Close the gap to Studio7",
    q: ["How big is device financing in our partner bases?", "Do we lose on reach, approval, price or channel?", "Which customers and shops to focus on?"],
    know: [["5–11%", "of loan applicants get a loan today: approval is the likely bottleneck"], ["362K", "AIS Serenade customers, a ready upgrade pool"]],
    a: ["Size upgraders: AIS customers with a handset 24+ months old or a contract ending",
      "Funnel by channel and device: apply, approve, book, first payment, reject reasons",
      "Benchmark Studio7 and others on rate, down payment, tenor, approval time",
      "Risk by segment and device price, and financing share by AIS shop"],
    out: "Market size, funnel diagnosis and terms benchmark", when: "Weeks 3–6" },
];

// ---------- Slide 1: the plan on a page
{
  const s = mk(null, "ANALYTICS PLAN  |  THREE MANAGEMENT ASKS", "One analytics plan for three asks: size the saving pocket before True launches, then target AIS top-up, then diagnose mobile financing", SRCP, "Analytics plan");
  const cw = (12.33 - 0.3) / 3, y = 1.6, H = 5.2;
  WS.forEach((w, k) => {
    const x = 0.5 + k * (cw + 0.15);
    RR(s, x, y, cw, H, C.background2, `Plan card ${k}`);
    R(s, x, y, cw, 0.08, w.col, `Plan strip ${k}`);
    badge(s, x + 0.18, y + 0.24, w.n, w.col, 0.4, 13);
    T(s, w.tag, { x: x + 0.7, y: y + 0.2, w: cw - 0.85, h: 0.2, fontSize: 8, bold: true, color: w.col, charSpacing: 1, objectName: `Plan tag ${k}` });
    T(s, [{ text: w.name, options: { bold: true, fontSize: 14, color: NAVY, breakLine: true } }, { text: w.sub, options: { fontSize: 9.5, color: MUTED } }], { x: x + 0.7, y: y + 0.4, w: cw - 0.85, h: 0.5, valign: "top", objectName: `Plan name ${k}` });
    // questions
    T(s, "QUESTIONS", { x: x + 0.2, y: y + 0.98, w: 2, h: 0.18, fontSize: 8, bold: true, color: MUTED, charSpacing: 1, objectName: `Plan q head ${k}` });
    T(s, w.q.map((t, i) => ({ text: t, options: { bullet: { indent: 10 }, breakLine: i < w.q.length - 1 } })), { x: x + 0.2, y: y + 1.17, w: cw - 0.4, h: 0.72, fontSize: 9.5, color: TXT, valign: "top", paraSpaceAfter: 1, objectName: `Plan q ${k}` });
    // what we know
    RR(s, x + 0.15, y + 1.92, cw - 0.3, 0.86, TINT, `Plan know box ${k}`, 0.05);
    w.know.forEach((f, i) => T(s, [{ text: f[0] + "  ", options: { bold: true, fontSize: 11, color: k === 0 ? TRUEC : BLUE } }, { text: f[1], options: { fontSize: 9, color: TXT } }],
      { x: x + 0.27, y: y + 1.96 + i * 0.4, w: cw - 0.54, h: 0.38, valign: "middle", objectName: `Plan know ${k}-${i}` }));
    // analyses
    T(s, "ANALYSES", { x: x + 0.2, y: y + 2.88, w: 2, h: 0.18, fontSize: 8, bold: true, color: MUTED, charSpacing: 1, objectName: `Plan a head ${k}` });
    w.a.forEach((t, i) => {
      const yy = y + 3.08 + i * 0.4;
      T(s, String(i + 1), { x: x + 0.2, y: yy + 0.03, w: 0.2, h: 0.2, fontSize: 8.5, bold: true, color: w.col, objectName: `Plan a num ${k}-${i}` });
      T(s, t, { x: x + 0.42, y: yy, w: cw - 0.6, h: 0.38, fontSize: 9, color: TXT, valign: "top", objectName: `Plan a ${k}-${i}` });
    });
    // output
    RR(s, x + 0.15, y + 4.7, cw - 0.3, 0.38, WHITE, `Plan out box ${k}`, 0.05);
    T(s, [{ text: w.when + "  ", options: { bold: true, fontSize: 9.5, color: w.col } }, { text: w.out, options: { fontSize: 9.5, color: NAVY, bold: true } }], { x: x + 0.27, y: y + 4.7, w: cw - 0.54, h: 0.38, valign: "middle", objectName: `Plan out ${k}` });
  });
}

// ---------- Slide 2: data, schedule, measurement
{
  const s = mk(null, "ANALYTICS PLAN  |  WHAT IT TAKES", "To start, we need one data request to AIS; with it, all three first outputs land within six weeks", SRCP, "Analytics plan");
  const y = 1.6, LW = 6.15, RX = 0.5 + LW + 0.15, RW = 12.83 - RX;
  // data requests
  RR(s, 0.5, y, LW, 3.3, C.background2, "Data panel");
  head(s, 0.5, y, LW, null, "Data we need", "Dots show which ask each item feeds");
  const cols = [["1", TRUEC], ["2", CYAN], ["3", BLUE]];
  cols.forEach((c, j) => badge(s, 0.5 + LW - 1.15 + j * 0.34, y + 0.62, c[0], c[1], 0.26, 9));
  const D = [["AIS", "TrueMoney usage per customer: frequency and amount", [1, 0, 0]], ["AIS", "Plan type, value and contract end date", [1, 1, 1]],
    ["AIS", "Top-up and bill-pay transactions by channel", [0, 1, 0]], ["AIS", "Handset model and age; shop sales by financing provider", [0, 0, 1]],
    ["CLICX", "Daily balance flows by customer", [1, 0, 0]], ["CLICX", "Biller-level payments (which biller was paid)", [0, 1, 0]],
    ["CLICX", "Device-loan applications, approvals and reject reasons by shop", [0, 0, 1]], ["Finance", "Package cost, cost of funds and deposit margin", [1, 0, 0]]];
  D.forEach((d, i) => {
    const yy = y + 0.95 + i * 0.29;
    T(s, d[0], { x: 0.7, y: yy, w: 0.75, h: 0.27, fontSize: 8.5, bold: true, color: d[0] === "AIS" ? NAVY : (d[0] === "CLICX" ? CYAN : MUTED), valign: "middle", objectName: `Data src ${i}` });
    T(s, d[1], { x: 1.45, y: yy, w: LW - 2.75, h: 0.27, fontSize: 9.5, color: TXT, valign: "middle", objectName: `Data item ${i}` });
    d[2].forEach((on, j) => s.addShape(pres.shapes.OVAL, { x: 0.5 + LW - 1.15 + j * 0.34 + 0.07, y: yy + 0.07, w: 0.13, h: 0.13, fill: { color: on ? cols[j][1] : WHITE }, line: { color: on ? cols[j][1] : LGREY, width: 0.75 }, objectName: `Data dot ${i}-${j}` }));
  });
  // schedule
  RR(s, RX, y, RW, 3.3, C.background2, "Schedule panel");
  head(s, RX, y, RW, null, "Six-week schedule", "Proposed; starts when the AIS data arrives");
  const gx = RX + 2.35, gw = RW - 2.55, wk = gw / 6;
  for (let i = 0; i < 6; i++) T(s, "W" + (i + 1), { x: gx + i * wk, y: y + 0.9, w: wk, h: 0.2, fontSize: 8.5, bold: true, color: MUTED, align: "center", objectName: `Week ${i + 1}` });
  const G = [["Data request to AIS", 0, 1, GREY], ["1  Saving pocket sizing", 0, 2, TRUEC], ["1  Pilot launch and read", 2, 4, TRUEC],
    ["2  Top-up baseline, segments", 1, 3, CYAN], ["2  Campaign test", 3, 3, CYAN], ["3  Financing market, funnel", 2, 3, BLUE], ["3  Terms benchmark", 3, 3, BLUE]];
  G.forEach((g, i) => {
    const yy = y + 1.13 + i * 0.29;
    T(s, g[0], { x: RX + 0.2, y: yy, w: 2.15, h: 0.28, fontSize: 9.5, color: TXT, valign: "middle", objectName: `Gantt label ${i}` });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: gx + g[1] * wk + 0.03, y: yy + 0.05, w: g[2] * wk - 0.06, h: 0.18, rectRadius: 0.05, fill: { color: g[3] }, line: { type: "none" }, objectName: `Gantt bar ${i}` });
  });
  // measurement
  RR(s, 0.5, y + 3.45, 12.33, 1.7, C.background2, "Measure panel");
  head(s, 0.5, y + 3.45, 12.33, null, "How each pilot will be judged", "Every pilot runs against a control group, so results show what the offer added");
  const M = [["1  Saving pocket", TRUEC, "Net deposit inflow and balances kept vs control · AIS sign-ups from non-AIS savers · package cost per baht retained"],
    ["2  AIS top-up and bill pay", CYAN, "Share of our AIS customers' top-ups through CLICX · repeat rate after 60 days · cost per retained payer"],
    ["3  Mobile financing", BLUE, "Approval and booking rate by channel · financing share in AIS shops · early default by segment"]];
  const mw = (12.33 - 0.4 - 0.3) / 3;
  M.forEach((m, k) => {
    const x = 0.7 + k * (mw + 0.15);
    RR(s, x, y + 4.35, mw, 0.68, WHITE, `Measure box ${k}`, 0.05);
    T(s, [{ text: m[0], options: { bold: true, fontSize: 10, color: m[1], breakLine: true } }, { text: m[2], options: { fontSize: 9, color: TXT } }], { x: x + 0.12, y: y + 4.35, w: mw - 0.24, h: 0.68, valign: "middle", objectName: `Measure ${k}` });
  });
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
