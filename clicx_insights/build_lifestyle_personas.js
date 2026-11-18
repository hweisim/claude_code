// CLICX board storyline v3 — situation, complication, resolution.
// Data: "DRAFT 20261105 - Data for Strategy 2027", Partner Payment Story v1, product segment extract (31 Oct 2026).
const pptxgen = require("pptxgenjs");
const { applyTheme } = require(process.env.SKILL + "/scripts/apply_theme.js");

const OUT = process.argv[2] || "CLICX_Lifestyle_Personas.pptx";

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
pres.title = "CLICX lifestyle personas";
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
// Lifestyle personas: CLICX data plus partner lifestyle signals (AIS, OR, KTB)
pres.addSection({ title: "Lifestyle personas" });
const LIFE = "5B6B7C";
const GR = [
  ["DIGITAL LIFESTYLE", BLUE, [
    ["Metro digital professionals", "Salaried, 20–34, Bangkok area; app-first", "136K", "salaried 20–34 in the Bangkok area; 14% have booked a loan", "Credit card · investing", "CLICX"],
    ["Up-country tech savvy", "Aged 25–44; heavy banking, wallet, crypto apps", "379K", "tech-savvy customers aged 25–44; up-country split to add", "Fractional shares · digital card", "AIS·KTB"],
    ["Young wealth aware", "Under 25, tech savvy, curious about investing", "63K", "tech savvy under 25; 8.7K already hold over ฿20K or Save Max", "Learn-and-earn investing · saving pocket", "AIS·KTB"],
    ["Crypto and gold traders", "Already send money to crypto, gold, stock apps", "฿31M", "paid to crypto, gold and stock apps in Oct; customer count to size", "Investing platform · high-yield saving", "CLICX"],
  ]],
  ["EVERYDAY LIFESTYLE", CYAN, [
    ["Café and commute regulars", "Café Amazon coffee plus PTT fuel, every week", "294K", "use both Café Amazon and PTT; 12% pay with us vs 8% for PTT only", "Rewards card · fuel instalments", "OR"],
    ["Streaming fans", "Subscribe to Western or Asian streaming", "489K", "of 1.1M aged 20–44 earning over ฿15K: 352K Western, 137K Asian", "Subscription-perk card · AIS bundle", "AIS"],
    ["Online shoppers", "Buy on Shopee, TikTok Shop and Lazada", "฿15.5M", "paid to the three shopping apps in Oct; customer count to size", "Buy now, pay later · cashback card", "CLICX"],
    ["Prepaid top-up loyalists", "AIS prepaid; top up little and often", "516K", "AIS prepaid customers; 8.4% pay with us vs 12.8% on postpaid", "Top-up-first offers · saving pocket", "AIS"],
  ]],
  ["LIVELIHOOD", NAVY, [
    ["Social sellers", "Sell on social media and marketplace apps", "92K", "online merchants; 77% applied without getting a loan", "SME loan on sales data · QR", "CLICX"],
    ["Neighbourhood shopkeepers", "Up-country shop and stall owners", "108K", "offline business owners up-country; 77% applied without a loan", "QR acquiring · working capital", "CLICX"],
    ["Gig hustlers", "Riders, drivers and freelancers paid daily", "268K", "57K riders and drivers, 211K freelancers; 80%+ got no loan", "Daily-pay micro-loan · accident cover", "CLICX"],
    ["Factory shift workers", "Blue-collar staff on shifts, paid monthly", "223K", "blue-collar employees, our largest job group; 83% hold ฿0", "Payday advance · mobile financing", "CLICX"],
  ]],
  ["LIFE STAGE", LIFE, [
    ["Steady government earners", "Officials and state-enterprise staff", "176K", "21% have booked a loan, the most of any job", "Salary-bridged loan · credit card", "CLICX"],
    ["Up-country breadwinners", "Aged 35–59, salaried or gig, raising a family", "213K", "up-country, aged 35–59; marital status and children to add", "Family insurance · education saving", "CLICX"],
    ["Credit rebuilders", "Applied, turned down, no money with us", "956K", "applied without a loan and hold ฿0; 450K KTB Jai-pump rejects also applied", "Secured loan · credit-builder saving", "CLICX·KTB"],
    ["Silver savers, 50+", "Aged 50+, protecting savings and health", "67K", "42K salaried aged 50–59 and 25K aged 60+; 7–12% hold over ฿20K", "High-yield saving · health insurance", "CLICX"],
  ]],
];
const SRCL = "Bases differ: CLICX counts from the segment extract (1.51M customers, balances 1 Oct 2026, payments Oct); tech savvy (AIS and KTB alternative data), streaming and AIS plans from Data for Strategy 2027 (draft); café and fuel from OR data (31 Oct 2026). ฿ = in-app spend in Oct. Personas overlap, so sizes do not add up";

// ---------- Slide 1: 16 lifestyle personas
{
  const s = mk(null, "LIFESTYLE PERSONAS  |  WHO THEY ARE", "Lifestyle widens the target: streaming, prepaid top-up and café-and-fuel groups are each bigger than all our borrowers today (165K)", SRCL, "Lifestyle personas");
  const cw = (12.33 - 0.45) / 4, y0 = 1.55, cy = 1.9, gap = 0.1, ch = (6.8 - cy - 3 * gap) / 4;
  let n = 0;
  GR.forEach(([g, col, items], k) => {
    const x = 0.5 + k * (cw + 0.15);
    T(s, g, { x, y: y0, w: cw, h: 0.26, fontSize: 9.5, bold: true, color: col, charSpacing: 1, valign: "middle", objectName: `Group ${k}` });
    s.addShape(pres.shapes.LINE, { x, y: y0 + 0.28, w: cw, h: 0, line: { color: col, width: 1.5 }, objectName: `Group rule ${k}` });
    items.forEach(([name, who, num, txt, lead, src], i) => {
      n++;
      const y = cy + i * (ch + gap);
      RR(s, x, y, cw, ch, PANELHEX, `Card ${n}`, 0.05);
      R(s, x, y + 0.06, 0.05, ch - 0.12, col, `Card strip ${n}`);
      const cx = x + 0.16, iw = cw - 0.28;
      T(s, [{ text: n + "  ", options: { color: col } }, { text: name, options: { color: NAVY } }], { x: cx, y: y + 0.07, w: iw - 0.52, h: 0.22, fontSize: 9.5, bold: true, valign: "middle", objectName: `Card name ${n}` });
      T(s, src, { x: x + cw - 0.12 - (src.length > 7 ? 0.66 : 0.5), y: y + 0.09, w: src.length > 7 ? 0.66 : 0.5, margin: 0, h: 0.17, fontSize: 6.5, bold: true, color: src.startsWith("CLICX") && src === "CLICX" ? CYAN : BLUE, align: "center", valign: "middle", shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.08, fill: { color: WHITE }, line: { color: src === "CLICX" ? CYAN : BLUE, width: 0.75 }, objectName: `Card source ${n}` });
      T(s, who, { x: cx, y: y + 0.29, w: iw, h: 0.19, fontSize: 8.5, color: MUTED, valign: "middle", objectName: `Card who ${n}` });
      T(s, num, { x: cx, y: y + 0.5, w: 0.74, h: 0.34, fontSize: 14, bold: true, color: col, valign: "middle", objectName: `Card size ${n}` });
      T(s, txt, { x: cx + 0.76, y: y + 0.49, w: iw - 0.76, h: 0.36, fontSize: 7.5, color: TXT, valign: "middle", objectName: `Card evidence ${n}` });
      T(s, [{ text: "LEAD  ", options: { bold: true, fontSize: 7, color: CYAN, charSpacing: 0.5 } }, { text: lead, options: { bold: true, fontSize: 8.5, color: NAVY } }], { x: cx, y: y + ch - 0.25, w: iw, h: 0.19, valign: "middle", objectName: `Card lead ${n}` });
    });
  });
}

// ---------- Slide 2: signals behind the personas and how to use them
{
  const s = mk(null, "LIFESTYLE PERSONAS  |  HOW TO BUILD THEM", "Most lifestyle signals already exist as partner totals: one customer-level data request to AIS, OR and KTB turns these personas into target lists", "Persona numbers refer to the previous slide. Partner totals = counts the partners have shared for CLICX customers as a group, not a flag per customer. Sources as on the previous slide", "Lifestyle personas");
  const ST = { have: ["In hand", CYAN], part: ["Partner totals", BLUE], req: ["To request", TRUEC], none: ["Not held", GREY] };
  const D = [
    ["Tech-savvy app use", "Comfort with digital money: banking, wallet, crypto apps", "AIS, KTB", "1–3", "part"],
    ["Streaming services", "Entertainment taste and subscription spend", "AIS", "6", "part"],
    ["Plan type and Serenade tier", "Spend level, loyalty and top-up habit", "AIS", "8", "part"],
    ["Café Amazon, PTT fuel, Blue+", "Daily routine and commuting", "OR", "5", "part"],
    ["Jai-pump reject, Paotang, NEXT", "Credit history and state-scheme use", "KTB", "15", "part"],
    ["TrueMoney use", "Wallet habit and exposure to True", "AIS", "6–8", "req"],
    ["Merchants paid in the app", "Shopping, investing, food and ride habits", "CLICX", "4, 7, 11", "have"],
    ["Balance, loans, applications", "Saving and credit need", "CLICX", "All", "have"],
    ["Age, job, address", "Life stage, place and way of earning", "CLICX", "All", "have"],
    ["Marital status, children", "Family stage", "None yet", "14, 16", "none"],
  ];
  const LW = 7.75, y = 1.6;
  RR(s, 0.5, y, LW, 5.15, PANELHEX, "Signal panel");
  head(s, 0.5, y, LW, null, "Lifestyle signals", "What each signal tells us, where it lives and which personas it feeds");
  const cx = [0.7, 2.75, 5.6, 6.35, 7.0], hdr = ["Signal", "What it tells us", "Source", "Personas", "Status"];
  const hy = y + 0.9;
  hdr.forEach((h, j) => T(s, h, { x: cx[j], y: hy, w: (cx[j + 1] || 8.1) - cx[j], h: 0.22, fontSize: 8.5, bold: true, color: MUTED, objectName: `Signal head ${j}` }));
  s.addShape(pres.shapes.LINE, { x: 0.7, y: hy + 0.25, w: LW - 0.4, h: 0, line: { color: NAVY, width: 1 }, objectName: "Signal rule" });
  const rh = 0.385;
  D.forEach((d, i) => {
    const yy = hy + 0.3 + i * rh;
    if (i % 2 === 0) R(s, 0.6, yy, LW - 0.2, rh, WHITE, `Signal band ${i}`);
    T(s, d[0], { x: cx[0], y: yy, w: cx[1] - cx[0] - 0.1, h: rh, fontSize: 9, bold: true, color: NAVY, valign: "middle", objectName: `Signal ${i}` });
    T(s, d[1], { x: cx[1], y: yy, w: cx[2] - cx[1] - 0.1, h: rh, fontSize: 8.5, color: TXT, valign: "middle", objectName: `Signal what ${i}` });
    T(s, d[2], { x: cx[2], y: yy, w: cx[3] - cx[2] - 0.05, h: rh, fontSize: 8.5, bold: true, color: d[2] === "CLICX" ? CYAN : (d[2] === "None yet" ? GREY : BLUE), valign: "middle", objectName: `Signal source ${i}` });
    T(s, d[3], { x: cx[3], y: yy, w: cx[4] - cx[3] - 0.05, h: rh, fontSize: 8.5, color: TXT, valign: "middle", objectName: `Signal personas ${i}` });
    const [lab, col] = ST[d[4]];
    T(s, lab, { x: cx[4], y: yy + 0.08, w: 1.05, h: rh - 0.16, fontSize: 7.5, bold: true, color: WHITE, align: "center", valign: "middle", shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.1, fill: { color: col }, line: { type: "none" }, objectName: `Signal status ${i}` });
  });
  // right: from personas to target lists
  const RX = 0.5 + LW + 0.15, RW = 12.83 - RX;
  RR(s, RX, y, RW, 5.15, PANELHEX, "Steps panel");
  head(s, RX, y, RW, null, "From personas to target lists", "Four steps, starting with the data request");
  const STEPS = [
    ["Ask partners for a flag per customer", "Turn AIS, OR and KTB totals into a yes/no flag on each CLICX customer, in one data request"],
    ["Tag every persona a customer matches", "Personas overlap: a café regular can also stream and top up on prepaid"],
    ["Give each customer one main persona", "Livelihood first, then life stage, then everyday and digital, so the money need leads the offer"],
    ["Test three first", null],
  ];
  STEPS.forEach(([t, d], i) => {
    const yy = y + 0.95 + i * 0.78;
    badge(s, RX + 0.2, yy + 0.02, i + 1, NAVY, 0.34, 11);
    T(s, t, { x: RX + 0.68, y: yy, w: RW - 0.85, h: 0.24, fontSize: 10.5, bold: true, color: NAVY, valign: "middle", objectName: `Step ${i + 1}` });
    if (d) T(s, d, { x: RX + 0.68, y: yy + 0.25, w: RW - 0.85, h: 0.44, fontSize: 9, color: TXT, valign: "top", objectName: `Step text ${i + 1}` });
  });
  const P3 = [["5", "Café and commute regulars", "Rewards card on OR fuel and coffee", CYAN], ["8", "Prepaid top-up loyalists", "Saving pocket with an AIS perk", CYAN], ["10", "Neighbourhood shopkeepers", "QR acquiring plus working capital", NAVY]];
  P3.forEach(([id, nm, off, col], i) => {
    const yy = y + 0.95 + 3 * 0.78 + 0.3 + i * 0.5;
    RR(s, RX + 0.68, yy, RW - 0.85, 0.44, WHITE, `Test box ${i}`, 0.05);
    T(s, [{ text: id + "  ", options: { bold: true, color: col } }, { text: nm, options: { bold: true, color: NAVY, breakLine: true } }, { text: off, options: { color: TXT, fontSize: 8.5 } }], { x: RX + 0.8, y: yy, w: RW - 1.05, h: 0.44, fontSize: 9.5, valign: "middle", objectName: `Test ${i}` });
  });
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
