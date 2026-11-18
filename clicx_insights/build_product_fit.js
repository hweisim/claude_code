// CLICX board storyline v3 — situation, complication, resolution.
// Data: "DRAFT 20261105 - Data for Strategy 2027", Partner Payment Story v1, product segment extract (31 Oct 2026).
const pptxgen = require("pptxgenjs");
const { applyTheme } = require(process.env.SKILL + "/scripts/apply_theme.js");

const OUT = process.argv[2] || "CLICX_Product_Fit_Segments.pptx";

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
pres.title = "CLICX product-fit segments";
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
// Product-fit segments: 15 rule-based segments from age x occupation x SME x address
const FIT = require(__dirname + "/../fit/segments15.json");
const SRCF = "Segment extract, 1,507,418 customers: balances as of 1 Oct 2026, payments in Oct. Applied, no loan = applied in the app and never booked (rejected or cancelled). Over ฿20K stands in for high-yield saving until the Save Max/More flag is added. Bangkok area = Bangkok, Nonthaburi, Pathum Thani, Samut Prakan";
pres.addSection({ title: "Product-fit segments" });

const SEG = {
  1: ["Students and young non-earners", "Under 20, or under 30 with no job"],
  2: ["Seniors", "Aged 60+"],
  3: ["SME owners, Bangkok area", "Business owners and family businesses (offline)"],
  4: ["SME owners, up-country", "Business owners and family businesses (offline)"],
  5: ["Online sellers", "Online merchants, any region"],
  6: ["Farm households", "Agriculture, owners and workers"],
  7: ["Government and state staff", "Officials, employees, state enterprise"],
  8: ["Young company staff", "Company employees, 20–29"],
  9: ["Prime company staff, Bangkok area", "Company employees, 30–59"],
  10: ["Prime company staff, up-country", "Company employees, 30–59"],
  11: ["Blue-collar workers", "Factory and manual employees"],
  12: ["Riders and drivers", "Delivery, ride-hailing, transport"],
  13: ["Freelancers", "Self-employed and temporary workers"],
  14: ["Wage earners", "Daily and hourly paid"],
  15: ["Non-working adults", "Aged 30–59, no job on file"],
};
const GROUPS = [["LIFE STAGE", [1, 2], GREY], ["BUSINESS OWNERS", [3, 4, 5, 6], CYAN], ["SALARIED", [7, 8, 9, 10], NAVY], ["MANUAL AND GIG", [11, 12, 13, 14], BLUE], ["OTHER", [15], GREY]];

// ---------- Slide 1: the 15 segments and their product footprint
{
  const s = mk(null, "PRODUCT-FIT SEGMENTS  |  WHO THEY ARE", "15 segments built on age, job, business ownership and address: government staff borrow most; seniors and Bangkok customers save most", SRCF, "Product-fit segments");
  const A = FIT.all;
  const pct = (v) => (v * 100).toFixed(v < 0.1 ? 1 : 0) + "%";
  const hd = ["#", "Segment", "Who", "Customers", "Active", "Applied, no loan", "Booked a loan", "Over ฿20K", "Paid in Oct", "Bangkok area", "AIS / OR use", "Credit quality"];
  const head = hd.map((h, j) => ({ text: h, options: { bold: true, fontSize: 8.5, color: WHITE, fill: { color: j >= 10 ? GREY : NAVY }, align: j < 3 ? "left" : "center", valign: "middle" } }));
  const metric = (v, base, hi) => {
    const r = v / base;
    const fill = r >= 1.25 ? (hi ? "CDEAE7" : "F4DCDC") : (r <= 0.75 ? (hi ? "F4DCDC" : "CDEAE7") : null);
    return { v, fill };
  };
  const body = [];
  GROUPS.forEach(([gname, ids, gcol]) => {
    ids.forEach((id, gi) => {
      const r = FIT.rows.find((x) => x.id === id);
      const zebra = body.length % 2 ? WHITE : "F7F8FA";
      const cells = [
        { text: String(id), options: { bold: true, color: WHITE, fill: { color: gcol }, align: "center" } },
        { text: SEG[id][0], options: { bold: true, color: NAVY } },
        { text: SEG[id][1], options: { color: MUTED, fontSize: 8 } },
        { text: `${(r.customers / 1000).toFixed(r.customers < 100000 ? 1 : 0)}K · ${pct(r.share)}`, options: { align: "right" } },
      ];
      [["active", true], ["rejected", false], ["booked", true], ["core", true], ["pay", true], ["bkk", null]].forEach(([k, hi]) => {
        const m = hi === null ? { v: r[k], fill: null } : metric(r[k], A[k], hi);
        cells.push({ text: pct(m.v), options: { align: "center", fill: m.fill ? { color: m.fill } : undefined } });
      });
      cells.push({ text: "to add", options: { align: "center", italic: true, color: GREY } }, { text: "to add", options: { align: "center", italic: true, color: GREY } });
      body.push(cells.map((c) => ({ text: c.text, options: { fontSize: 9, color: TXT, valign: "middle", fill: { color: zebra }, ...c.options } })));
    });
  });
  const all = ["", "All customers", "", "1.51M · 100%", pct(A.active), pct(A.rejected), pct(A.booked), pct(A.core), pct(A.pay), pct(A.bkk), "", ""].map((t, j) => ({ text: t, options: { fontSize: 9, bold: true, color: NAVY, fill: { color: WHITE }, align: j < 3 ? "left" : (j === 3 ? "right" : "center"), valign: "middle", border: [{ pt: 1.25, color: NAVY }, { type: "none" }, { type: "none" }, { type: "none" }] } }));
  s.addTable([head, ...body, all], { x: 0.5, y: 1.58, w: 12.33, colW: [0.3, 2.35, 2.45, 1.0, 0.65, 0.78, 0.72, 0.7, 0.68, 0.72, 0.64, 0.74], rowH: [0.42, ...body.map(() => 0.283), 0.3], fontFace: "Arial", border: { type: "solid", pt: 0.5, color: "E3E8EE" }, margin: [0.02, 0.05, 0.02, 0.05], objectName: "Segment table" });
  T(s, "Teal = 25%+ better than all customers, rose = 25%+ worse (for applied with no loan, lower is better). Grey columns are the two requested dimensions not in this extract", { x: 0.5, y: 6.6, w: 12.33, h: 0.22, fontSize: 8.5, italic: true, color: MUTED, objectName: "Table key" });
}

// ---------- Slide 2: where each product fits
{
  const s = mk(null, "PRODUCT-FIT SEGMENTS  |  WHERE PRODUCTS FIT", "Each product has a natural home: SME loans for owners and sellers, cards for salaried staff, high-yield saving for seniors and Bangkok customers", "Lead product = strongest fit on the segment's data (what it already holds, applies for and keeps); next = a second step. Judgement on the data, to validate with AIS/OR use and credit quality before targeting. " + SRCF, "Product-fit segments");
  const P = ["High-yield saving / AIS pocket", "Investing", "Credit card", "Salary-bridged or personal loan", "SME loan", "Mobile financing", "Insurance", "Top-up and bill pay"];
  // 2 = lead, 1 = next
  const M = {
    1: [1, 1, 0, 0, 0, 0, 0, 2], 2: [2, 1, 0, 0, 0, 0, 2, 0], 3: [2, 0, 1, 0, 2, 0, 0, 1], 4: [1, 0, 0, 0, 2, 0, 1, 0], 5: [0, 0, 1, 0, 2, 0, 0, 2],
    6: [1, 0, 0, 0, 1, 0, 2, 0], 7: [1, 0, 2, 2, 0, 0, 1, 0], 8: [1, 0, 2, 0, 0, 2, 0, 1], 9: [2, 2, 2, 0, 0, 1, 0, 0], 10: [1, 0, 2, 2, 0, 0, 0, 0],
    11: [1, 0, 0, 1, 0, 2, 1, 0], 12: [0, 0, 0, 1, 0, 1, 2, 1], 13: [1, 0, 0, 1, 0, 0, 1, 2], 14: [1, 0, 0, 1, 0, 0, 0, 2], 15: [2, 1, 0, 0, 0, 0, 1, 0],
  };
  const WHY = {
    1: "Pay most (12%); thin balances", 2: "12% over ฿20K; no loans after 59", 3: "17% save, most of any job; 70% rejected", 4: "77% applied without a loan; few save",
    5: "10% pay in-app; 77% rejected", 6: "Least active (14%); seasonal income", 7: "21% booked a loan, the most; 31% active", 8: "81% hold ฿0; first credit years",
    9: "6% over ฿20K; 29% active", 10: "14% booked; few save", 11: "Largest segment; 83% hold ฿0", 12: "86% rejected; least active", 13: "80% rejected; pay over borrow",
    14: "14% active; daily income", 15: "25% save; 11% over ฿20K",
  };
  const x0 = 0.5, nameW = 2.8, whyW = 2.4, cellW = (12.33 - nameW - whyW) / P.length, y0 = 1.62, hh = 0.62, rh = 0.272;
  P.forEach((p, j) => T(s, p, { x: x0 + nameW + j * cellW, y: y0, w: cellW, h: hh, fontSize: 8.5, bold: true, color: NAVY, align: "center", valign: "bottom", objectName: `Product head ${j}` }));
  T(s, "Why", { x: x0 + nameW + P.length * cellW + 0.1, y: y0, w: whyW, h: hh, fontSize: 8.5, bold: true, color: NAVY, valign: "bottom", objectName: "Why head" });
  s.addShape(pres.shapes.LINE, { x: x0, y: y0 + hh + 0.03, w: 12.33, h: 0, line: { color: NAVY, width: 1.25 }, objectName: "Matrix rule" });
  let row = 0;
  GROUPS.forEach(([gname, ids, gcol]) => {
    ids.forEach((id) => {
      const y = y0 + hh + 0.08 + row * rh;
      if (row % 2 === 0) R(s, x0, y, 12.33, rh, "F7F8FA", `Matrix band ${id}`);
      const r = FIT.rows.find((q) => q.id === id);
      T(s, [{ text: id + "  ", options: { bold: true, color: gcol } }, { text: SEG[id][0], options: { color: TXT } }, { text: "  " + (r.customers / 1000).toFixed(0) + "K", options: { color: MUTED, fontSize: 8 } }],
        { x: x0 + 0.05, y, w: nameW - 0.05, h: rh, fontSize: 9, valign: "middle", objectName: `Matrix seg ${id}` });
      M[id].forEach((v, j) => {
        if (!v) return;
        const cx = x0 + nameW + j * cellW + cellW / 2, d = 0.15;
        s.addShape(pres.shapes.OVAL, { x: cx - d / 2, y: y + (rh - d) / 2, w: d, h: d, fill: v === 2 ? { color: CYAN } : { type: "none" }, line: { color: CYAN, width: 1.25 }, objectName: `Fit ${id}-${j}` });
      });
      T(s, WHY[id], { x: x0 + nameW + P.length * cellW + 0.1, y, w: whyW - 0.1, h: rh, fontSize: 8.5, color: MUTED, valign: "middle", objectName: `Matrix why ${id}` });
      row++;
    });
  });
  const ly = y0 + hh + 0.12 + row * rh;
  s.addShape(pres.shapes.OVAL, { x: x0 + 0.05, y: ly + 0.04, w: 0.15, h: 0.15, fill: { color: CYAN }, line: { color: CYAN, width: 1.25 }, objectName: "Key lead" });
  T(s, "Lead product", { x: x0 + 0.25, y: ly, w: 1.2, h: 0.23, fontSize: 9, color: MUTED, valign: "middle", objectName: "Key lead text" });
  s.addShape(pres.shapes.OVAL, { x: x0 + 1.45, y: ly + 0.04, w: 0.15, h: 0.15, fill: { type: "none" }, line: { color: CYAN, width: 1.25 }, objectName: "Key next" });
  T(s, "Next product", { x: x0 + 1.65, y: ly, w: 1.2, h: 0.23, fontSize: 9, color: MUTED, valign: "middle", objectName: "Key next text" });
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
