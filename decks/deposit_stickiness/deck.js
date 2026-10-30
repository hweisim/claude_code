const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");
const { applyTheme } = require("/root/.claude/skills/synced/ab465d10-90d7-41e2-9a3b-b6dc85394be5_b50fba5f-f80d-4696-8ec5-49df53e26a55/pptx/scripts/apply_theme.js");

const OUT = process.argv[2] || "Deposit_Stickiness_Revamp.pptx";

// ---------- Palette: blue / grey / teal ----------
const HEX = {
  navy: "0B1F3A", blue: "1F4E8C", sky: "4A7FC1", teal: "0E8A85", tealLt: "A8DCD9",
  slate: "6B7A8F", grey: "C9D1DB", ink: "1E2A38", bgLt: "F2F5F8", white: "FFFFFF",
  muted: "5F6B7A", tintTeal: "E3F3F2", tintBlue: "E4ECF7", tintGrey: "EEF1F4", tintNavy: "DCE2EB",
};
const THEME = {
  name: "Deposit Stickiness",
  headFontFace: "Georgia",
  bodyFontFace: "Arial",
  colors: {
    dk1: HEX.ink, lt1: HEX.white, dk2: HEX.navy, lt2: HEX.bgLt,
    accent1: HEX.blue, accent2: HEX.teal, accent3: HEX.slate, accent4: HEX.sky,
    accent5: HEX.grey, accent6: HEX.tealLt, hlink: HEX.blue, folHlink: HEX.slate,
  },
};
// Outcome colours: Increase = teal, Stable = slate, Decrease/depleted = navy
const OUTC = { inc: HEX.teal, stb: HEX.slate, dec: HEX.navy };

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "Deposit stickiness segmentation";
const C = pres.SchemeColor;
const S = {
  ink: C.text1, navy: C.text2, white: C.background1, bgLt: C.background2,
  blue: C.accent1, teal: C.accent2, slate: C.accent3, sky: C.accent4, grey: C.accent5, tealLt: C.accent6,
};

// ---------- Layouts ----------
pres.defineSlideMaster({
  title: "TITLE_DARK",
  background: { color: S.navy },
  objects: [
    { placeholder: { options: { name: "kicker", type: "body", x: 0.8, y: 1.7, w: 11, h: 0.4, fontSize: 13, bold: true, color: S.tealLt, charSpacing: 2 }, text: "" } },
    { placeholder: { options: { name: "title", type: "title", x: 0.8, y: 2.2, w: 11.2, h: 2.2, fontSize: 36, bold: true, color: S.white, valign: "top", align: "left", fontFace: "Georgia" }, text: "" } },
    { placeholder: { options: { name: "subtitle", type: "body", x: 0.8, y: 4.65, w: 10.5, h: 0.9, fontSize: 18, color: S.grey, valign: "top" }, text: "" } },
    { placeholder: { options: { name: "date", type: "body", x: 0.8, y: 6.4, w: 8, h: 0.4, fontSize: 12, color: S.grey }, text: "" } },
  ],
});
pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: S.white },
  margin: [0.5, 0.6, 0.6, 0.6],
  slideNumber: { x: 12.3, y: 7.0, w: 0.5, h: 0.3, fontSize: 9, color: S.slate, align: "right", fontFace: "Arial" },
  objects: [
    { placeholder: { options: { name: "tracker", type: "body", x: 0.6, y: 0.3, w: 10, h: 0.3, fontSize: 10, bold: true, color: S.teal, charSpacing: 1, margin: 0 }, text: "" } },
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.62, w: 12.1, h: 0.95, fontSize: 24, bold: true, color: S.navy, valign: "top", align: "left", margin: 0, fontFace: "Georgia" }, text: "" } },
    { placeholder: { options: { name: "footnote", type: "body", x: 0.6, y: 6.9, w: 11.2, h: 0.45, fontSize: 9, color: S.slate, valign: "top", margin: 0 }, text: "" } },
  ],
});

// ---------- Helpers ----------
async function iconPng(Icon, color = "#FFFFFF", size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Icon, { color, size: String(size) }));
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}
// Sparkline for a pair of moves (-1 / 0 / +1)
async function spark(m1, m2, color) {
  const W = 300, H = 140, pad = 24, step = 34;
  const ys = [0, m1, m1 + m2].map((v) => H / 2 - v * step);
  const xs = [pad, W / 2, W - pad];
  const pts = xs.map((x, i) => `${x},${ys[i]}`).join(" ");
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}">
    <line x1="${pad}" y1="${H / 2}" x2="${W - pad}" y2="${H / 2}" stroke="#9AA6B5" stroke-width="2" stroke-dasharray="6,6"/>
    <polyline points="${pts}" fill="none" stroke="#${color}" stroke-width="7" stroke-linejoin="round" stroke-linecap="round"/>
    ${xs.map((x, i) => `<circle cx="${x}" cy="${ys[i]}" r="11" fill="#${color}" stroke="#FFFFFF" stroke-width="3"/>`).join("")}
  </svg>`;
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}
function content(tracker, title, footnote, section) {
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: section });
  s.addText(tracker.toUpperCase(), { placeholder: "tracker" });
  s.addText(title, { placeholder: "title" });
  if (footnote) s.addText(footnote, { placeholder: "footnote" });
  return s;
}
function tb(s, text, o) { s.addText(text, Object.assign({ isTextBox: true, margin: 0, fontSize: 14, color: S.ink, valign: "top" }, o)); }
function box(s, o) { s.addShape(pres.shapes.RECTANGLE, Object.assign({ line: { color: S.white, width: 0 } }, o)); }
function numCircle(s, n, x, y, d, fill, name) {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { color: fill, width: 0 }, objectName: name });
  tb(s, String(n), { x, y, w: d, h: d, align: "center", valign: "middle", bold: true, color: S.white, fontSize: Math.round(d * 30) });
}
function kpi(s, x, y, w, h, big, label, sub, color, name) {
  box(s, { x, y, w, h, fill: { color: S.bgLt }, objectName: name });
  tb(s, big, { x: x + 0.25, y: y + 0.15, w: w - 0.4, h: 0.6, fontSize: 30, bold: true, color, fontFace: "Georgia", valign: "middle" });
  tb(s, label, { x: x + 0.25, y: y + 0.78, w: w - 0.4, h: 0.3, fontSize: 13, bold: true, color: S.ink });
  tb(s, sub, { x: x + 0.25, y: y + 1.07, w: w - 0.4, h: 0.3, fontSize: 11, color: S.slate });
}
const chartBase = () => ({
  catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt", dataLabelFontFace: "+mn-lt", legendFontFace: "+mn-lt",
  catAxisLabelColor: HEX.ink, valAxisLabelColor: HEX.slate, catAxisLabelFontSize: 12, valAxisLabelFontSize: 10,
  dataLabelFontSize: 11, legendFontSize: 11, catAxisLineShow: false, valAxisHidden: true,
  valGridLine: { style: "none" }, catGridLine: { style: "none" },
});

(async () => {
  const icons = {
    observe: await iconPng(fa.FaChartLine), assign: await iconPng(fa.FaLayerGroup), validate: await iconPng(fa.FaCheckDouble),
    shield: await iconPng(fa.FaShieldAlt), gem: await iconPng(fa.FaGem), bolt: await iconPng(fa.FaBolt), sliders: await iconPng(fa.FaSlidersH),
  };

  // ===== 1. Title =====
  pres.addSection({ title: "Introduction" });
  {
    const s = pres.addSlide({ masterName: "TITLE_DARK", sectionTitle: "Introduction" });
    s.addText("DEPOSIT STICKINESS", { placeholder: "kicker" });
    s.addText("Three months of balance behavior tells us which deposits will stay — and which will run off", { placeholder: "title" });
    s.addText("Segmentation concept (Jul–Sep) and October validation across 1.45m CLICX customers", { placeholder: "subtitle" });
    s.addText("Conceptual framework · thresholds subject to recalibration", { placeholder: "date" });
  }

  // ===== 2. Executive summary =====
  {
    const s = content("Executive summary", "Behavior-based segments separate sticky deposits from runoff risk; most customers, however, still need activation", null, "Introduction");
    const rows = [
      ["The approach", "Each customer is assigned one segment at end-Sep from Jul–Sep monthly average balances, then frozen and tested against end-Oct balances."],
      ["Where customers sit", "96% of customers fall into priority segments (New, Zero, Low); the 3.8% with a full funded history hold 51% of balances (THB 3.0bn of 5.85bn)."],
      ["It predicts runoff", "Mixed (35.7%) and Decrease (33.5%) lose balances at ~4.5× the rate of Increase (7.6%) — the segmentation ranks risk correctly."],
      ["Stable and Increase are the core", "72% of Stable customers stayed within band; 85% of Increase-segment balances stayed or grew in October."],
      ["Activation is the bigger gap", "95% of New customers remain at ≤THB 100 — a funding-activation issue, not a retention signal."],
    ];
    const y0 = 1.85, rh = 0.95;
    rows.forEach((r, i) => {
      const y = y0 + i * rh;
      numCircle(s, i + 1, 0.6, y + 0.12, 0.45, i === 2 ? S.teal : S.navy, `exec-num-${i + 1}`);
      tb(s, r[0], { x: 1.3, y: y + 0.1, w: 3.4, h: 0.55, bold: true, fontSize: 15, color: S.navy, valign: "middle" });
      tb(s, r[1], { x: 4.85, y: y + 0.05, w: 7.85, h: 0.7, fontSize: 14, valign: "middle" });
      if (i < rows.length - 1) s.addShape(pres.shapes.LINE, { x: 1.3, y: y + rh - 0.08, w: 11.4, h: 0, line: { color: S.grey, width: 0.75 }, objectName: `exec-sep-${i + 1}` });
    });
  }

  // ===== 3. Approach on a page =====
  pres.addSection({ title: "Segmentation concept" });
  {
    const s = content("Concept | Approach on a page", "Stickiness is inferred from three months of behavior, frozen in September, then checked against October", "Monthly average balance is used for segmentation; M3 (Sep) end balance is the baseline for end-Oct validation.", "Segmentation concept");
    // Timeline
    const ty = 2.05;
    tb(s, "Observation window — builds the segment", { x: 0.6, y: ty, w: 8.6, h: 0.3, fontSize: 12, bold: true, color: S.blue });
    tb(s, "Validation window", { x: 9.55, y: ty, w: 3.2, h: 0.3, fontSize: 12, bold: true, color: S.teal });
    const months = [["M1", "July"], ["M2", "August"], ["M3", "September"]];
    months.forEach((m, i) => {
      const x = 0.6 + i * 2.95;
      box(s, { x, y: ty + 0.4, w: 2.85, h: 0.85, fill: { color: i === 2 ? S.blue : HEX.tintBlue }, objectName: `tl-${m[0]}` });
      tb(s, [{ text: m[0] + "  ", options: { bold: true, fontSize: 18 } }, { text: m[1], options: { fontSize: 13 } }],
        { x: x + 0.2, y: ty + 0.4, w: 2.5, h: 0.85, valign: "middle", color: i === 2 ? S.white : S.navy });
    });
    box(s, { x: 9.55, y: ty + 0.4, w: 3.18, h: 0.85, fill: { color: S.teal }, objectName: "tl-oct" });
    tb(s, [{ text: "October  ", options: { bold: true, fontSize: 18 } }, { text: "end balance", options: { fontSize: 13 } }],
      { x: 9.75, y: ty + 0.4, w: 2.9, h: 0.85, valign: "middle", color: S.white });
    // freeze marker
    s.addShape(pres.shapes.DIAMOND, { x: 9.27, y: ty + 0.62, w: 0.4, h: 0.4, fill: { color: S.navy }, line: { color: S.white, width: 2 }, objectName: "freeze-marker" });
    tb(s, "Segment frozen at 30 Sep — never recalculated with October data", { x: 6.2, y: ty + 1.35, w: 6.5, h: 0.3, fontSize: 11, italic: true, color: S.navy, align: "right" });

    const cards = [
      [icons.observe, S.blue, "1  Observe", ["Track total CLICX balance for Jul, Aug, Sep (M1–M3)", "Use monthly averages to remove payday and one-day transaction noise"]],
      [icons.assign, S.navy, "2  Assign one segment", ["Apply 7 segments in priority order", "History first, then current balance, then 3-month direction"]],
      [icons.validate, S.teal, "3  Validate", ["Compare Sep end balance with Oct end balance", "Success = Stable / Increase retain better; Decrease / Zero run off more"]],
    ];
    cards.forEach((c, i) => {
      const x = 0.6 + i * 4.085, y = 3.95, w = 3.95, h = 2.7;
      box(s, { x, y, w, h, fill: { color: S.bgLt }, objectName: `step-card-${i + 1}` });
      s.addShape(pres.shapes.OVAL, { x: x + 0.3, y: y + 0.3, w: 0.6, h: 0.6, fill: { color: c[1] }, line: { color: c[1], width: 0 }, objectName: `step-icon-bg-${i + 1}` });
      s.addImage({ data: c[0], x: x + 0.45, y: y + 0.45, w: 0.3, h: 0.3, objectName: `step-icon-${i + 1}` });
      tb(s, c[2], { x: x + 1.1, y: y + 0.3, w: w - 1.3, h: 0.6, fontSize: 17, bold: true, color: S.navy, valign: "middle" });
      tb(s, c[3].map((t, j) => ({ text: t, options: { bullet: true, breakLine: j < c[3].length - 1 } })),
        { x: x + 0.3, y: y + 1.1, w: w - 0.55, h: 1.5, fontSize: 14, paraSpaceAfter: 8 });
    });
  }

  // ===== 4. Threshold + worked example =====
  {
    const s = content("Concept | What counts as a real move", "A balance only “moves” if it changes by more than 10% or THB 1,000, whichever is larger",
      "Illustrative customer; figures are monthly average balances in THB. Threshold is applied to each month-on-month change separately.", "Segmentation concept");
    // Left: rule
    box(s, { x: 0.6, y: 1.85, w: 5.4, h: 1.55, fill: { color: S.navy }, objectName: "rule-box" });
    tb(s, "THE RULE", { x: 0.85, y: 2.0, w: 5, h: 0.3, fontSize: 11, bold: true, color: S.tealLt, charSpacing: 1 });
    tb(s, "A move counts if the change is greater than max (10% × prior-month balance, THB 1,000)", { x: 0.85, y: 2.35, w: 5.0, h: 0.9, fontSize: 16, bold: true, color: S.white, fontFace: "Georgia" });
    tb(s, "Threshold by prior-month balance", { x: 0.6, y: 3.65, w: 5.4, h: 0.3, fontSize: 13, bold: true, color: S.navy });
    const th = [["Prior-month balance", "Threshold", "Why"], ["THB 5,000", "± 1,000", "Floor applies"], ["THB 10,000", "± 1,000", "10% = floor"], ["THB 50,000", "± 5,000", "10% applies"], ["THB 100,000", "± 10,000", "10% applies"]];
    s.addTable(th.map((r, i) => r.map((c, j) => ({
      text: c, options: {
        bold: i === 0 || j === 1, color: i === 0 ? HEX.white : (j === 1 ? HEX.teal : HEX.ink), fill: { color: i === 0 ? HEX.slate : (i % 2 ? HEX.white : HEX.bgLt) },
        align: j === 0 ? "left" : "center", fontSize: 13, valign: "middle",
      },
    }))), { x: 0.6, y: 4.05, w: 5.4, colW: [2.2, 1.5, 1.7], rowH: 0.42, border: { type: "solid", pt: 0.5, color: HEX.grey }, margin: [0.03, 0.1, 0.03, 0.1] });

    // Right: worked example chart
    tb(s, "Worked example: how one customer gets a segment", { x: 6.6, y: 1.85, w: 6.1, h: 0.3, fontSize: 13, bold: true, color: S.navy });
    s.addChart(pres.charts.BAR, [{ name: "Avg balance (THB k)", labels: ["Jul (M1)", "Aug (M2)", "Sep (M3)"], values: [50, 52, 40] }], Object.assign(chartBase(), {
      x: 6.6, y: 2.2, w: 6.1, h: 2.2, barDir: "col", barGapWidthPct: 90, chartColors: [HEX.sky, HEX.sky, HEX.navy],
      showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: '"THB "0"k"', dataLabelColor: HEX.ink, dataLabelFontBold: true,
      showLegend: false, valAxisMinVal: 0, valAxisMaxVal: 65, objectName: "example-chart",
    }));
    const steps = [
      ["Jul → Aug", "+2.0k", "within ± 5.0k", "Stable", S.slate],
      ["Aug → Sep", "–12.0k", "beyond ± 5.2k", "Decrease", S.navy],
    ];
    steps.forEach((st, i) => {
      const x = 6.6 + i * 3.1, y = 4.55;
      box(s, { x, y, w: 2.95, h: 1.05, fill: { color: S.bgLt }, objectName: `example-step-${i + 1}` });
      tb(s, st[0], { x: x + 0.2, y: y + 0.1, w: 2.6, h: 0.3, fontSize: 12, bold: true, color: S.slate });
      tb(s, [{ text: st[1] + "  ", options: { bold: true, fontSize: 16, color: S.ink } }, { text: st[2], options: { fontSize: 12, color: S.slate } }], { x: x + 0.2, y: y + 0.4, w: 2.65, h: 0.3 });
      tb(s, "→ " + st[3], { x: x + 0.2, y: y + 0.7, w: 2.6, h: 0.3, fontSize: 13, bold: true, color: st[4] });
    });
    box(s, { x: 6.6, y: 5.75, w: 6.05, h: 0.85, fill: { color: S.teal }, objectName: "example-result" });
    tb(s, [{ text: "Result:  ", options: { bold: true } }, { text: "Stable → Decrease  =  Segment 5 “Decrease”" }], { x: 6.85, y: 5.75, w: 5.7, h: 0.85, valign: "middle", fontSize: 15, color: S.white });
  }

  // ===== 5. Priority ladder with distribution =====
  {
    const s = content("Concept | Seven segments, assigned in priority order", "Priority checks route 96% of customers to activation segments; the 4% with full history hold half of balances",
      "Customer counts and balances as of 30 Sep. Behavioral segments require an account opened on or before 1 Jul and Sep end balance > THB 100.", "Segmentation concept");
    const segs = [
      ["1", "New customer", "Opened after 1 Jul (<3 months history)", 1301.7, 2853, HEX.slate],
      ["2", "Zero balance", "Sep end balance = THB 0", 51.9, 0, HEX.slate],
      ["3", "Low balance", "Sep end balance > 0 and ≤ THB 100", 40.8, 0.6, HEX.slate],
      ["4", "Increase", "Rising across M1→M2→M3, never falling", 15.0, 2573, HEX.teal],
      ["5", "Stable", "Both monthly moves within threshold", 29.8, 297, HEX.slate],
      ["6", "Mixed", "Direction reverses (up then down, or down then up)", 4.3, 86, HEX.sky],
      ["7", "Decrease", "Falling across M1→M2→M3, never rising", 5.6, 43, HEX.navy],
    ];
    // NB: original numbering — 4 Increase, 5 Decrease, 6 Stable, 7 Mixed. Keep the original IDs.
    const ids = { Increase: "4", Stable: "6", Mixed: "7", Decrease: "5" };
    const cx = { rule: 1.95, cust: 7.05, bal: 10.0 };
    tb(s, "Segment and rule", { x: 1.95, y: 1.8, w: 4.5, h: 0.3, fontSize: 12, bold: true, color: S.slate });
    tb(s, "Customers, k", { x: cx.cust, y: 1.8, w: 2.6, h: 0.3, fontSize: 12, bold: true, color: S.slate });
    tb(s, "Sep balance, THB m", { x: cx.bal, y: 1.8, w: 2.7, h: 0.3, fontSize: 12, bold: true, color: S.slate });
    const y0 = 2.15, rh = 0.55, gap = 0.2;
    segs.forEach((g, i) => {
      const y = y0 + i * rh + (i >= 3 ? gap : 0);
      const id = ids[g[1]] || g[0];
      const isBeh = i >= 3;
      numCircle(s, id, 1.4, y + 0.1, 0.38, isBeh ? g[5] : HEX.slate, `ladder-num-${id}`);
      tb(s, g[1], { x: cx.rule, y: y + 0.02, w: 4.9, h: 0.28, fontSize: 14, bold: true, color: S.navy });
      tb(s, g[2], { x: cx.rule, y: y + 0.3, w: 4.9, h: 0.25, fontSize: 11, color: S.slate });
      const cw = 1.9 * Math.max(g[3] / 1301.7, 0.012);
      box(s, { x: cx.cust, y: y + 0.12, w: cw, h: 0.34, fill: { color: isBeh ? g[5] : HEX.grey }, objectName: `cust-bar-${id}` });
      tb(s, g[3].toLocaleString("en-US", { minimumFractionDigits: 1 }), { x: cx.cust + cw + 0.08, y: y + 0.12, w: 0.85, h: 0.34, fontSize: 12, bold: true, valign: "middle" });
      const bw = 2.1 * Math.max(g[4] / 2853, 0.012);
      if (g[4] > 0) box(s, { x: cx.bal, y: y + 0.12, w: bw, h: 0.34, fill: { color: isBeh ? g[5] : HEX.grey }, objectName: `bal-bar-${id}` });
      tb(s, g[4] === 0 ? "0" : g[4] < 1 ? "0.6" : g[4].toLocaleString("en-US"), { x: cx.bal + (g[4] > 0 ? bw + 0.08 : 0), y: y + 0.12, w: 0.8, h: 0.34, fontSize: 12, bold: true, valign: "middle" });
    });
    // group brackets (labels on the left)
    box(s, { x: 0.6, y: y0, w: 0.6, h: 3 * rh - 0.05, fill: { color: HEX.tintGrey }, objectName: "group-priority" });
    tb(s, "Priority checks", { x: 0.6, y: y0, w: 0.6, h: 3 * rh - 0.05, fontSize: 10, bold: true, color: S.slate, align: "center", valign: "middle", vert: "vert270" });
    box(s, { x: 0.6, y: y0 + 3 * rh + gap, w: 0.6, h: 4 * rh - 0.05, fill: { color: HEX.tintTeal }, objectName: "group-behavior" });
    tb(s, "3-month behavior", { x: 0.6, y: y0 + 3 * rh + gap, w: 0.6, h: 4 * rh - 0.05, fontSize: 10, bold: true, color: S.teal, align: "center", valign: "middle", vert: "vert270" });
    // totals strip
    const ty = y0 + 7 * rh + gap + 0.08;
    box(s, { x: 1.4, y: ty, w: 11.3, h: 0.42, fill: { color: S.bgLt }, objectName: "totals-strip" });
    tb(s, [{ text: "Priority 1–3: ", options: { bold: true } }, { text: "1,394k customers (96%) · THB 2.85bn (49%)     " }, { text: "Behavioral 4–7: ", options: { bold: true, color: S.teal } }, { text: "55k customers (4%) · THB 3.00bn (51%)" }],
      { x: 1.6, y: ty, w: 11, h: 0.42, fontSize: 13, valign: "middle", color: S.navy });
  }

  // ===== 6. Matrix =====
  {
    const s = content("Concept | Behavioral segments 4–7", "For funded customers, the two monthly moves combine into four behavior patterns",
      "Move = change in monthly average balance beyond ± max(10% × prior-month balance, THB 1,000). Original segment numbers retained.", "Segmentation concept");
    const segMeta = {
      Increase: { id: 4, color: HEX.teal, tint: HEX.tintTeal },
      Stable: { id: 6, color: HEX.slate, tint: HEX.tintGrey },
      Mixed: { id: 7, color: HEX.sky, tint: HEX.tintBlue },
      Decrease: { id: 5, color: HEX.navy, tint: HEX.tintNavy },
    };
    const moves = [["Decrease", -1], ["Stable", 0], ["Increase", 1]];
    const grid = [ // [m1][m2]
      ["Decrease", "Decrease", "Mixed"],
      ["Decrease", "Stable", "Increase"],
      ["Mixed", "Increase", "Increase"],
    ];
    const gx = 2.35, gy = 2.3, cw = 2.55, ch = 1.38;
    tb(s, "Second move: Aug → Sep (M2 → M3)", { x: gx, y: 1.75, w: 3 * cw, h: 0.3, fontSize: 12, bold: true, color: S.slate, align: "center" });
    tb(s, "First move: Jul → Aug (M1 → M2)", { x: 0.6, y: gy, w: 0.4, h: 3 * ch, fontSize: 12, bold: true, color: S.slate, align: "center", valign: "middle", vert: "vert270" });
    moves.forEach((m, j) => tb(s, m[0], { x: gx + j * cw, y: gy - 0.28, w: cw, h: 0.25, fontSize: 12, bold: true, color: S.navy, align: "center" }));
    moves.forEach((m, i) => tb(s, m[0], { x: 1.1, y: gy + i * ch, w: 1.15, h: ch, fontSize: 12, bold: true, color: S.navy, valign: "middle", align: "right" }));
    for (let i = 0; i < 3; i++) for (let j = 0; j < 3; j++) {
      const seg = grid[i][j], meta = segMeta[seg];
      const x = gx + j * cw + 0.05, y = gy + i * ch + 0.05, w = cw - 0.1, h = ch - 0.1;
      box(s, { x, y, w, h, fill: { color: meta.tint }, objectName: `cell-${i}-${j}` });
      s.addImage({ data: await spark(moves[i][1], moves[j][1], meta.color), x: x + 0.15, y: y + 0.12, w: 1.2, h: 0.56, objectName: `spark-${i}-${j}` });
      tb(s, `${meta.id}  ${seg}`, { x: x + 0.15, y: y + 0.78, w: w - 0.3, h: 0.4, fontSize: 15, bold: true, color: meta.color === HEX.sky ? HEX.blue : meta.color });
    }
    // Right legend
    const lx = 10.2;
    tb(s, "How to read it", { x: lx, y: 1.75, w: 2.5, h: 0.3, fontSize: 13, bold: true, color: S.navy });
    const legend = [
      ["Increase", "Grew at least once and never fell"],
      ["Stable", "No meaningful move in either month"],
      ["Mixed", "Direction reversed — up then down, or down then up"],
      ["Decrease", "Fell at least once and never rose"],
    ];
    legend.forEach((l, i) => {
      const y = 2.25 + i * 1.05, meta = segMeta[l[0]];
      s.addShape(pres.shapes.OVAL, { x: lx, y: y + 0.06, w: 0.2, h: 0.2, fill: { color: meta.color }, line: { color: meta.color, width: 0 }, objectName: `legend-dot-${i}` });
      tb(s, l[0], { x: lx + 0.3, y, w: 2.2, h: 0.3, fontSize: 14, bold: true, color: S.navy });
      tb(s, l[1], { x: lx + 0.3, y: y + 0.32, w: 2.25, h: 0.65, fontSize: 12, color: S.ink });
    });
  }

  // ===== 7. Runoff by segment =====
  pres.addSection({ title: "October validation" });
  {
    const s = content("Validation | Gross runoff", "Mixed and Decrease segments ran off ~4.5x faster than Increase — the segmentation ranks risk correctly",
      "Gross runoff = sum of customer-level balance declines Sep → Oct ÷ 30 Sep segment balance. Zero balance excluded (Sep base = 0). October data as supplied.", "October validation");
    tb(s, "Gross runoff rate, Sep → Oct, % of 30 Sep segment balance", { x: 0.6, y: 1.8, w: 7.6, h: 0.3, fontSize: 13, bold: true, color: S.navy });
    const cats = ["Increase", "Stable", "Mixed", "Decrease", "New", "Low"];
    s.addChart(pres.charts.BAR, [{ name: "Gross runoff", labels: cats, values: [0.076, 0.223, 0.357, 0.335, 0.095, 0.171] }], Object.assign(chartBase(), {
      x: 0.6, y: 2.15, w: 7.6, h: 4.0, barDir: "bar", barGapWidthPct: 55, catAxisOrientation: "maxMin",
      chartColors: [HEX.teal, HEX.slate, HEX.navy, HEX.navy, HEX.grey, HEX.grey],
      showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0.0%", dataLabelColor: HEX.ink, dataLabelFontBold: true, dataLabelFontSize: 12,
      showLegend: false, valAxisMinVal: 0, valAxisMaxVal: 0.42, objectName: "runoff-chart",
    }));
    // Right callouts
    box(s, { x: 9.6, y: 1.85, w: 3.1, h: 2.25, fill: { color: S.navy }, objectName: "callout-multiple" });
    tb(s, "4.7×", { x: 9.85, y: 2.0, w: 2.7, h: 0.9, fontSize: 48, bold: true, color: S.white, fontFace: "Georgia" });
    tb(s, "Mixed runoff (35.7%) vs Increase (7.6%); Decrease is 4.4×", { x: 9.85, y: 2.95, w: 2.7, h: 1.0, fontSize: 13, color: S.white });
    box(s, { x: 9.6, y: 4.3, w: 3.1, h: 1.85, fill: { color: S.bgLt }, objectName: "callout-thb" });
    tb(s, "THB runoff in October", { x: 9.85, y: 4.42, w: 2.7, h: 0.3, fontSize: 12, bold: true, color: S.navy });
    tb(s, [
      { text: "Increase  THB 195.6m", options: { breakLine: true } },
      { text: "Stable  THB 66.3m", options: { breakLine: true } },
      { text: "Mixed  THB 30.6m", options: { breakLine: true } },
      { text: "Decrease  THB 14.4m" },
    ], { x: 9.85, y: 4.75, w: 2.7, h: 1.3, fontSize: 12, color: S.ink, paraSpaceAfter: 4 });
    tb(s, "Largest absolute runoff sits in the large Increase book, even at the lowest rate.", { x: 0.6, y: 6.3, w: 8.8, h: 0.4, fontSize: 12, italic: true, color: S.slate });
  }

  // ===== 8. Behavioral outcome mix =====
  {
    const s = content("Validation | Funded customers (segments 4–7)", "Among 54.6k funded customers, Mixed deteriorated most — 49% of customers and 55% of balances declined",
      "Outcome based on October end balance vs September end balance. Decreased / depleted includes October Decrease, Zero and Low balance outcomes.", "October validation");
    const k = [
      ["10.2%", "Increased", "+THB 221m added", HEX.teal],
      ["60.1%", "Stable", "32,804 customers within band", HEX.slate],
      ["29.7%", "Decreased / depleted", "–THB 294m lost", HEX.navy],
      ["–THB 73m", "Net balance change", "–2.4% of Sep THB 3.00bn base", HEX.navy],
    ];
    k.forEach((t, i) => kpi(s, 0.6 + i * 3.08, 1.75, 2.88, 1.45, t[0], t[1], t[2], t[3], `kpi-beh-${i + 1}`));
    const cats = ["Increase", "Stable", "Mixed", "Decrease"];
    const byCust = [[0.18, 0.04, 0.21, 0.13], [0.48, 0.72, 0.30, 0.51], [0.34, 0.24, 0.49, 0.36]];
    const byBal = [[0.11, 0.03, 0.15, 0.10], [0.74, 0.71, 0.30, 0.39], [0.15, 0.26, 0.55, 0.51]];
    const sub = [["Share of customers", "15.0k · 29.8k · 4.3k · 5.6k customers"], ["Share of Sep balance", "THB 2,573m · 297m · 86m · 43m"]];
    [byCust, byBal].forEach((d, i) => {
      const x = 0.6 + i * 6.25;
      tb(s, sub[i][0], { x, y: 3.45, w: 5.9, h: 0.3, fontSize: 13, bold: true, color: S.navy });
      tb(s, sub[i][1], { x, y: 3.73, w: 5.9, h: 0.25, fontSize: 10, color: S.slate });
      s.addChart(pres.charts.BAR, [
        { name: "Increase", labels: cats, values: d[0] }, { name: "Stable", labels: cats, values: d[1] }, { name: "Decrease / depleted", labels: cats, values: d[2] },
      ], Object.assign(chartBase(), {
        x, y: 4.0, w: 5.9, h: 2.75, barDir: "bar", barGrouping: "percentStacked", barGapWidthPct: 45, catAxisOrientation: "maxMin",
        chartColors: [OUTC.inc, OUTC.stb, OUTC.dec], showValue: true, dataLabelPosition: "ctr", dataLabelFormatCode: '[<0.05]"";0%', dataLabelColor: HEX.white, dataLabelFontBold: true,
        showLegend: true, legendPos: "b", legendColor: HEX.ink, objectName: `mix-chart-${i + 1}`,
      }));
    });
  }

  // ===== 9. Full base / activation =====
  {
    const s = content("Validation | All 1.45m customers", "Across the full base, 93% ended October decreased or depleted — driven by New customers who never funded",
      "Outcome based on October end balance. Decreased / depleted includes October Zero, Low balance and Decrease outcomes. Net change as supplied.", "October validation");
    const cats = ["New (1,301.7k)", "Zero (51.9k)", "Low (40.8k)", "Increase (15.0k)", "Stable (29.8k)", "Mixed (4.3k)", "Decrease (5.6k)"];
    s.addChart(pres.charts.BAR, [
      { name: "Increase", labels: cats, values: [0.01, 0.01, 0.04, 0.18, 0.04, 0.21, 0.13] },
      { name: "Stable", labels: cats, values: [0.03, 0.01, 0.05, 0.48, 0.72, 0.30, 0.51] },
      { name: "Decrease / depleted", labels: cats, values: [0.96, 0.98, 0.91, 0.34, 0.24, 0.49, 0.36] },
    ], Object.assign(chartBase(), {
      x: 0.6, y: 2.1, w: 7.9, h: 4.6, barDir: "bar", barGrouping: "percentStacked", barGapWidthPct: 40, catAxisOrientation: "maxMin",
      chartColors: [OUTC.inc, OUTC.stb, OUTC.dec], showValue: true, dataLabelPosition: "ctr", dataLabelFormatCode: '[<0.05]"";0%', dataLabelColor: HEX.white, dataLabelFontBold: true,
      showLegend: true, legendPos: "b", legendColor: HEX.ink, objectName: "full-mix-chart",
    }));
    tb(s, "Oct outcome, % of customers in each Sep segment", { x: 0.6, y: 1.75, w: 7.9, h: 0.3, fontSize: 13, bold: true, color: S.navy });
    const k = [
      ["1.5% | 5.4% | 93.0%", "Increased | Stable | Decreased", "Net –THB 6m, –0.1% of Sep THB 5.85bn base"],
      ["94.9%", "of New customers stay ≤ THB 100", "1.30m customers need a first funding trigger"],
      ["75%", "of New-customer balances held stable", "The few who fund behave like core customers"],
    ];
    k.forEach((t, i) => {
      const y = 1.85 + i * 1.6;
      box(s, { x: 8.9, y, w: 3.8, h: 1.45, fill: { color: i === 1 ? S.navy : S.bgLt }, objectName: `full-kpi-${i + 1}` });
      tb(s, t[0], { x: 9.1, y: y + 0.12, w: 3.5, h: 0.55, fontSize: i === 0 ? 20 : 28, bold: true, color: i === 1 ? S.white : S.navy, fontFace: "Georgia", valign: "middle" });
      tb(s, t[1], { x: 9.1, y: y + 0.7, w: 3.5, h: 0.3, fontSize: 12, bold: true, color: i === 1 ? S.white : S.ink });
      tb(s, t[2], { x: 9.1, y: y + 1.0, w: 3.5, h: 0.3, fontSize: 11, color: i === 1 ? S.grey : S.slate });
    });
  }

  // ===== 10. Implications =====
  pres.addSection({ title: "Implications" });
  {
    const s = content("Implications | Proposed next steps", "Each segment group calls for a different action: retain, protect or activate",
      "Proposed actions for discussion. Customer counts and balances as of 30 Sep.", "Implications");
    const cols = [
      [icons.shield, HEX.navy, "Retain", "Mixed + Decrease", "9.8k customers · THB 129m", "34–36% gross runoff", ["Early-warning list from the Sep segment", "Targeted retention offers before balances fall further"]],
      [icons.gem, HEX.teal, "Protect", "Increase + Stable", "44.8k customers · THB 2.87bn", "8–22% gross runoff", ["Keep the core book sticky; watch large Increase balances", "Monitor Stable customers drifting to Decrease"]],
      [icons.bolt, HEX.slate, "Activate", "New + Zero + Low", "1.39m customers · THB 2.85bn", "91–98% unfunded or depleted", ["Funding-activation journeys for new accounts", "Re-score once 3 months of history exist"]],
    ];
    cols.forEach((c, i) => {
      const x = 0.6 + i * 4.085, y = 1.85, w = 3.95, h = 3.95;
      box(s, { x, y, w, h, fill: { color: S.bgLt }, objectName: `impl-card-${i + 1}` });
      s.addShape(pres.shapes.OVAL, { x: x + 0.3, y: y + 0.3, w: 0.65, h: 0.65, fill: { color: c[1] }, line: { color: c[1], width: 0 }, objectName: `impl-icon-bg-${i + 1}` });
      s.addImage({ data: c[0], x: x + 0.46, y: y + 0.46, w: 0.33, h: 0.33, objectName: `impl-icon-${i + 1}` });
      tb(s, c[2], { x: x + 1.15, y: y + 0.27, w: 2.6, h: 0.4, fontSize: 20, bold: true, color: S.navy, fontFace: "Georgia" });
      tb(s, c[3], { x: x + 1.15, y: y + 0.66, w: 2.6, h: 0.3, fontSize: 12, bold: true, color: c[1] === HEX.slate ? S.slate : c[1] });
      tb(s, c[4], { x: x + 0.3, y: y + 1.2, w: w - 0.5, h: 0.3, fontSize: 13, bold: true, color: S.ink });
      tb(s, c[5], { x: x + 0.3, y: y + 1.5, w: w - 0.5, h: 0.3, fontSize: 12, color: S.slate });
      tb(s, c[6].map((t, j) => ({ text: t, options: { bullet: true, breakLine: j < c[6].length - 1 } })), { x: x + 0.3, y: y + 2.0, w: w - 0.5, h: 1.8, fontSize: 14, paraSpaceAfter: 10 });
    });
    box(s, { x: 0.6, y: 6.0, w: 12.12, h: 0.75, fill: { color: S.white }, line: { color: S.grey, width: 1 }, objectName: "refine-strip" });
    s.addShape(pres.shapes.OVAL, { x: 0.85, y: 6.1, w: 0.55, h: 0.55, fill: { color: S.blue }, line: { color: S.blue, width: 0 }, objectName: "refine-icon-bg" });
    s.addImage({ data: icons.sliders, x: 0.985, y: 6.235, w: 0.28, h: 0.28, objectName: "refine-icon" });
    tb(s, [{ text: "Refine the model:  ", options: { bold: true, color: S.navy } }, { text: "recalibrate the 10% / THB 1,000 threshold after reviewing segment distribution, and repeat validation over further months before operational use." }],
      { x: 1.7, y: 6.0, w: 10.8, h: 0.75, fontSize: 14, valign: "middle" });
  }

  // ===== 11. Appendix fact base =====
  pres.addSection({ title: "Appendix" });
  {
    const s = content("Appendix | Segment fact base", "All seven segments at a glance",
      "* Low: Sep end balance THB 0.62m → Oct THB 15.61m. Zero end-balance % undefined (Sep = 0). Outcome mix for New, Zero and Low derived from bar lengths in source; may not sum exactly. October data as supplied.", "Appendix");
    const H = (t) => ({ text: t, options: { bold: true, color: HEX.white, fill: { color: HEX.navy }, align: "center", fontSize: 11, valign: "middle" } });
    const hdr = ["Sep segment", "Customers, k", "Sep balance, THB m", "Gross runoff", "Runoff, THB m", "Avg bal Δ", "End bal Δ", "Oct: Increase", "Oct: Stable", "Oct: Decr./depl."].map(H);
    const data = [
      ["1 New", "1,301.7", "2,853", "9.5%", "270.0", "+5.5%", "+1.5%", "1%", "3%", "96%"],
      ["2 Zero", "51.9", "0", "n/a", "–", "–79.1%", "n/a", "1%", "1%", "98%"],
      ["3 Low", "40.8", "0.6", "17.1%", "0.1", "–36.1%", "+2,404%*", "4%", "5%", "91%"],
      ["4 Increase", "15.0", "2,573", "7.6%", "195.6", "+3.7%", "–0.7%", "18%", "48%", "34%"],
      ["6 Stable", "29.8", "297", "22.3%", "66.3", "–10.0%", "–13.9%", "4%", "72%", "24%"],
      ["7 Mixed", "4.3", "86", "35.7%", "30.6", "–4.0%", "–9.1%", "21%", "30%", "49%"],
      ["5 Decrease", "5.6", "43", "33.5%", "14.4", "–26.4%", "–11.4%", "13%", "51%", "36%"],
      ["Total", "1,448.9", "5,850", "", "", "", "", "1.5%", "5.4%", "93.0%"],
    ];
    const rows = data.map((r, i) => r.map((c, j) => {
      const neg = /^–/.test(c), pos = /^\+/.test(c);
      const hot = (j === 3 && (i === 5 || i === 6));
      return { text: c, options: {
        bold: j === 0 || i === 7 || hot, align: j === 0 ? "left" : "center", fontSize: 12, valign: "middle",
        color: hot ? HEX.white : (j >= 5 && j <= 6 ? (neg ? HEX.navy : pos ? HEX.teal : HEX.ink) : HEX.ink),
        fill: { color: hot ? HEX.navy : (i === 7 ? HEX.tintGrey : (i >= 3 ? HEX.white : HEX.bgLt)) },
      } };
    }));
    s.addTable([hdr, ...rows], { x: 0.6, y: 1.8, w: 12.12, colW: [1.6, 1.15, 1.35, 1.1, 1.15, 1.05, 1.15, 1.17, 1.17, 1.23], rowH: [0.55, 0.47, 0.47, 0.47, 0.47, 0.47, 0.47, 0.47, 0.47], border: { type: "solid", pt: 0.5, color: HEX.grey }, margin: [0.03, 0.06, 0.03, 0.06] });
    tb(s, "Shaded rows = priority segments (activation focus). Highlighted = highest runoff.", { x: 0.6, y: 6.35, w: 12, h: 0.3, fontSize: 11, italic: true, color: S.slate });
  }

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
