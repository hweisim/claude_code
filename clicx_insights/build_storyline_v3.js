// CLICX board storyline v3 — situation, complication, resolution.
// Data: "DRAFT 20261105 - Data for Strategy 2027", Partner Payment Story v1, product segment extract (31 Oct 2026).
const pptxgen = require("pptxgenjs");
const { applyTheme } = require(process.env.SKILL + "/scripts/apply_theme.js");

const OUT = process.argv[2] || "CLICX_Storyline_v3.pptx";

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
pres.title = "CLICX board storyline v3";
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
const SRC = "Source: Data for Business Strategy (draft, 5 Nov 2026)";
pres.addSection({ title: "Storyline" });

// ---------- Cover
{
  const s = pres.addSlide({ masterName: "TITLE", sectionTitle: "Storyline" });
  s.addText("BOARD STORYLINE V3  ·  DATA FOR STRATEGY 2027", { placeholder: "kicker" });
  s.addText("From sign-ups to customers", { placeholder: "title" });
  s.addText("How CLICX activates, grows and defends its 2.1M base in 2027", { placeholder: "sub" });
  s.addText("Board discussion  ·  November 2026", { placeholder: "date" });
}

// ---------- 1 Executive summary
{
  const s = mk(-1, "1 / 11  |  EXECUTIVE SUMMARY", "We have won the customers we aimed for; 2027 is about getting them to use CLICX, and holding them before True arrives",
    SRC + ". CLICX data as of Sep–Oct 2026; AIS, KTB, OR as of Apr 2026. Active = monthly transacting customers");
  const cw = (12.33 - 0.3) / 3, y = 1.62, H = 4.48;
  const cols = [
    ["WHAT WE BUILT", NAVY, [["2.1M", "customers in a year; 96% came through AIS, KTB or OR"],
      ["82%", "of loan applicants are underserved, and tech-savvy customers join at twice the average: the customers we aimed for"],
      ["64%", "of the customers we share with KTB were inactive there: we add to KTB rather than take from it"]]],
    ["WHAT HOLDS US BACK", GREY, [["3 in 4", "came for a loan, and only 1 in 10 of them got one"],
      ["1.3M", "left with no loan and no money with us; only about 310K customers are active"],
      ["19%", "of partner payers stopped after the August promotion: payments lean on deals"],
      ["฿1.98B", "of deposits (63%) sit with True and TrueMoney customers as True's virtual bank launches"]]],
    ["WHAT WE WILL DO IN 2027", CYAN, [["ACQUIRE", "Make KTB's 1.9M yearly loan rejects a standing channel; reach 313K TrueMoney users first"],
      ["ACTIVATE", "A first deposit or first loan for the 1.3M; top-up offers for 516K AIS prepaid customers"],
      ["GROW", "Savings for 103K borrowers; SME loan, credit card and investing for pools we already hold"],
      ["DEFEND", "Hold ฿1.98B of deposits with savings rewards and a card for under-40s"]]],
  ];
  cols.forEach((c, k) => {
    const x = 0.5 + k * (cw + 0.15);
    RR(s, x, y, cw, H, C.background2, `Summary card ${k}`);
    R(s, x, y, cw, 0.42, c[1], `Summary band ${k}`);
    T(s, c[0], { x: x + 0.2, y, w: cw - 0.4, h: 0.42, fontSize: 11, bold: true, color: WHITE, charSpacing: 1, valign: "middle", objectName: `Summary head ${k}` });
    const n = c[2].length, rh = (H - 0.52) / n;
    c[2].forEach((r, i) => {
      const yy = y + 0.47 + i * rh;
      if (i) s.addShape(pres.shapes.LINE, { x: x + 0.2, y: yy, w: cw - 0.4, h: 0, line: { color: LGREY, width: 0.75 }, objectName: `Summary rule ${k}-${i}` });
      if (k < 2) {
        T(s, r[0], { x: x + 0.2, y: yy, w: 1.3, h: rh, fontSize: 21, bold: true, color: k === 0 ? BLUE : (i === 3 ? TRUEC : NAVY), valign: "middle", objectName: `Summary num ${k}-${i}` });
        T(s, r[1], { x: x + 1.55, y: yy, w: cw - 1.75, h: rh, fontSize: 10.5, color: TXT, valign: "middle", objectName: `Summary text ${k}-${i}` });
      } else {
        s.addText(r[0], { x: x + 0.2, y: yy + rh / 2 - 0.14, w: 1.05, h: 0.28, shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.06, fill: { color: MOVE[r[0]] }, line: { type: "none" },
          fontSize: 8.5, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0, charSpacing: 1, objectName: `Move tag ${i}` });
        T(s, r[1], { x: x + 1.4, y: yy, w: cw - 1.6, h: rh, fontSize: 10.5, color: TXT, valign: "middle", objectName: `Move text ${i}` });
      }
    });
  });
  RR(s, 0.5, 6.22, 12.33, 0.56, NAVY, "Decisions");
  T(s, [{ text: "DECISIONS WE NEED  ", options: { bold: true, fontSize: 9.5, color: LCYAN, charSpacing: 1 } },
    { text: "Risk appetite for alternative-data first loans · partner budget moved to top-up and repeat-use rewards · order and funding of SME loan, card and investing · unit economics in the data room", options: { fontSize: 10.5, color: WHITE } }],
    { x: 0.7, y: 6.22, w: 11.93, h: 0.56, valign: "middle", objectName: "Decisions text" });
}

// ---------- 2 Scale and fit
{
  const s = mk(0, "2 / 11  |  SCALE AND FIT", "Partners gave us 2.1M customers in a year, and they are the ones we aimed for: tech-savvy, young and underserved",
    "Join rate = CLICX customers ÷ the group's size in the 49.5M consortium base. Tech-savvy = uses Make, Kept, Dime or crypto. Target groups = early adopters (Gen Z, tech-savvy, multi-bank) and people without a steady income. CLICX Sep 2026; partners Apr 2026. " + SRC);
  const y = 1.62, H = 4.55, LW = 4.2, RX = 4.85, RW = 12.83 - 4.85;
  RR(s, 0.5, y, LW, H, C.background2, "Partner panel");
  head(s, 0.5, y, LW, 1, "96% came through our partners", "CLICX customers, Sep 2026");
  RR(s, 0.7, y + 0.95, LW - 0.4, 0.85, TINT, "Base stat");
  T(s, "2.1M", { x: 0.85, y: y + 0.95, w: 1.05, h: 0.85, fontSize: 24, bold: true, color: BLUE, valign: "middle", objectName: "Base stat num" });
  T(s, "customers; 1.86M gave partner consent, and 96% of those are AIS, KTB or OR customers", { x: 1.9, y: y + 0.95, w: LW - 1.75, h: 0.85, fontSize: 10, color: TXT, valign: "middle", objectName: "Base stat text" });
  T(s, "ON CLICX", { x: 1.6, y: y + 1.95, w: 0.9, h: 0.2, fontSize: 8, bold: true, color: MUTED, charSpacing: 0.5, objectName: "Partner col 1" });
  T(s, "SHARE OF ITS BASE THAT JOINED", { x: 2.5, y: y + 1.95, w: 2.1, h: 0.2, fontSize: 8, bold: true, color: MUTED, charSpacing: 0.5, objectName: "Partner col 2" });
  [["logo_ais.png", 2.74, "1.2M", 3.7], ["logo_ktb.jpeg", 2.58, "1.6M", 4.7], ["logo_or.png", 1.93, "725K", 7.4]].forEach((p, k) => {
    const yy = y + 2.2 + k * 0.52;
    s.addImage({ path: IC(p[0]), x: 0.7, y: yy + 0.11, h: 0.26, w: 0.26 * p[1], altText: "Partner logo", objectName: `Partner logo ${k}` });
    T(s, p[2], { x: 1.6, y: yy, w: 0.85, h: 0.48, fontSize: 13, bold: true, color: NAVY, valign: "middle", objectName: `Partner size ${k}` });
    R(s, 2.5, yy + 0.14, 1.3 * p[3] / 8, 0.2, CYAN, `Partner bar ${k}`);
    T(s, p[3].toFixed(1) + "%", { x: 2.5 + 1.3 * p[3] / 8 + 0.06, y: yy, w: 0.6, h: 0.48, fontSize: 11, bold: true, color: NAVY, valign: "middle", objectName: `Partner rate ${k}` });
  });
  RR(s, 0.7, y + 3.85, LW - 0.4, 0.52, WHITE, "Headroom");
  T(s, [{ text: "47M  ", options: { bold: true, fontSize: 16, color: NAVY } }, { text: "partner customers are not on CLICX yet", options: { fontSize: 10.5, color: TXT } }],
    { x: 0.85, y: y + 3.85, w: LW - 0.7, h: 0.52, valign: "middle", objectName: "Headroom text" });

  RR(s, RX, y, RW, H, C.background2, "Join panel");
  head(s, RX, y, RW, 2, "The groups we targeted join fastest", "% of each partner-customer group now on CLICX");
  legend(s, RX + RW - 2.75, y + 0.6, [["Target group", CYAN], ["Other", SOFT]]);
  T(s, "BY DIGITAL PROFILE AND AGE", { x: RX + 0.2, y: y + 1.0, w: 3.7, h: 0.22, fontSize: 8.5, bold: true, color: MUTED, charSpacing: 0.5, objectName: "Join head A" });
  hbars(s, RX + 0.2, y + 1.27, 3.85, [["Tech-savvy", 8.5, "8.5%", CYAN], ["Working Gen Z, 25–29", 5.6, "5.6%", CYAN], ["Working, 30–59", 4.1, "4.1%", SOFT],
    ["KTB Next only", 3.9, "3.9%", SOFT], ["Multi-bank users", 2.4, "2.4%", CYAN, "below average"], ["Aged 60+", 0.2, "0.2%", SOFT]],
    { lw: 1.75, rh: 0.36, max: 10, fs: 10, name: "Profile", labelFill: PANELHEX, bench: { v: 3.8, label: "All groups 3.8%" } });
  T(s, "BY JOB", { x: RX + 4.3, y: y + 1.0, w: 3.4, h: 0.22, fontSize: 8.5, bold: true, color: MUTED, charSpacing: 0.5, objectName: "Join head B" });
  hbars(s, RX + 4.3, y + 1.27, 3.5, [["Freelancer", 40, "40%", CYAN], ["Rider, driver", 22, "22%", CYAN], ["Blue collar", 16, "16%", SOFT], ["Merchant", 9, "9%", CYAN], ["White collar", 9, "9%", SOFT]],
    { lw: 1.2, rh: 0.36, max: 45, fs: 10, name: "Job", labelFill: PANELHEX, bench: { v: 3.8, label: "All 3.8%" } });
  const chips = [["71%", "aged 25–44"], ["70%", "live outside Bangkok"], ["82%", "of applicants underserved"]], cw = (RW - 0.4 - 0.2) / 3;
  chips.forEach((c, k) => {
    const x = RX + 0.2 + k * (cw + 0.1);
    RR(s, x, y + 3.85, cw, 0.52, WHITE, `Chip ${k}`);
    T(s, [{ text: c[0] + "  ", options: { bold: true, fontSize: 16, color: BLUE } }, { text: c[1], options: { fontSize: 10, color: TXT } }], { x: x + 0.12, y: y + 3.85, w: cw - 0.2, h: 0.52, valign: "middle", objectName: `Chip text ${k}` });
  });
  soWhat(s, "Growth is still inside the partner bases: aim campaigns at the groups that join at up to 10 times the average, and give multi-bank users a reason to switch");
}

// ---------- 3 KTB: additive
{
  const s = mk(0, "3 / 11  |  VALUE TO KTB", "We did it without taking from KTB: 64% of shared customers were inactive there, and 450K of its loan rejects applied with us",
    "KTB-inactive = no KTB activity before joining CLICX. Jai-pump rejects Jun 2025–May 2026. Early default (FPDX) is measured on borrowers who reached their first bill. CLICX data as of Sep 2026. " + SRC);
  const y = 1.62, H = 4.55, LW = 5.0, RX = 5.65, RW = 12.83 - 5.65;
  RR(s, 0.5, y, LW, H, C.background2, "KTB panel");
  head(s, 0.5, y, LW, 1, "1.1M shared customers were inactive at KTB", "KTB customers on CLICX");
  [["1.6M", "KTB customers on CLICX", 1, NAVY], ["1.1M · 64%", "were inactive at KTB", 0.64, BLUE]].forEach((f, k) => {
    const yy = y + 0.98 + k * 0.48, ww = (LW - 0.4) * f[2];
    R(s, 0.7, yy, ww, 0.4, f[3], `KTB bar ${k}`);
    T(s, [{ text: f[0] + "  ", options: { bold: true } }, { text: f[1] }], { x: 0.82, y: yy, w: ww - 0.2, h: 0.4, fontSize: 11, color: WHITE, valign: "middle", objectName: `KTB bar text ${k}` });
  });
  RR(s, 0.7, y + 2.08, LW - 0.4, 0.85, TINT, "KTB stat");
  numRow(s, 0.85, y + 2.08, LW - 0.7, 0.85, "84%", "of them came to CLICX for a loan: they borrowed or applied", { nw: 1.0, nfs: 24 });
  [["515K", "bank mainly with other banks"], ["548K", "were dormant or used only Paotang"], ["37%", "open the CLICX app; 4% use Save Max/More"]].forEach((f, k) =>
    numRow(s, 0.7, y + 3.05 + k * 0.47, LW - 0.4, 0.44, f[0], f[1], { nw: 0.75, nfs: 14, col: NAVY, fs: 10.5 }));
  RR(s, RX, y, RW, H, C.background2, "Reject panel");
  head(s, RX, y, RW, 2, "KTB's loan rejects become our customers", "Jai-pump rejects, last 12 months");
  const fn = [["1.9M", "Jai-pump rejects"], ["450K", "applied with CLICX: 24% of rejects"], ["51K", "got a loan: 11% of applicants"], ["20.2%", "early default (FPDX)"]];
  const fw = (RW - 0.4 - 0.6) / 4;
  fn.forEach((f, k) => {
    const x = RX + 0.2 + k * (fw + 0.2);
    RR(s, x, y + 0.98, fw, 1.3, k === 3 ? WHITE : TINT, `Funnel box ${k}`);
    T(s, [{ text: f[0], options: { bold: true, fontSize: 22, color: k === 3 ? NAVY : BLUE, breakLine: true } }, { text: f[1], options: { fontSize: 9.5, color: TXT } }], { x: x + 0.12, y: y + 1.0, w: fw - 0.2, h: 1.26, valign: "middle", objectName: `Funnel step ${k}` });
    if (k < 3) T(s, "›", { x: x + fw, y: y + 1.43, w: 0.2, h: 0.4, fontSize: 20, bold: true, color: BLUE, align: "center", valign: "middle", objectName: `Funnel arrow ${k}` });
  });
  const ins = [["Paotang converts better", "Rejects referred by Paotang apply at 30% vs 23% via NEXT, and default less (18% vs 21%)"],
    ["Riders apply most, get a loan least", "64% of rider rejects apply, but only 7% of those applicants get a loan"],
    ["Core age is 25–39", "61% of applicants; the apply rate peaks at 29% for ages 25–34"]];
  const iw = (RW - 0.4 - 0.3) / 3;
  ins.forEach((t, k) => {
    const x = RX + 0.2 + k * (iw + 0.15);
    RR(s, x, y + 2.45, iw, 1.9, WHITE, `Reject insight ${k}`);
    badge(s, x + 0.14, y + 2.62, "abc"[k], BLUE, 0.3, 10);
    T(s, t[0], { x: x + 0.54, y: y + 2.55, w: iw - 0.66, h: 0.44, fontSize: 11, bold: true, color: NAVY, valign: "middle", objectName: `Reject head ${k}` });
    T(s, t[1], { x: x + 0.14, y: y + 3.12, w: iw - 0.28, h: 1.15, fontSize: 10.5, color: TXT, valign: "top", objectName: `Reject text ${k}` });
  });
  soWhat(s, "Make the Paotang reject list a standing acquisition channel: under-40s and freelancers first, with loans priced for a 20% early-default rate");
}

// ---------- 4 The base map: came for cash, left with nothing
{
  const s = mk(1, "4 / 11  |  WHY CUSTOMERS GO QUIET", "But 3 in 4 customers came for a loan and only 1 in 10 of them got one: 1.3M left with nothing, and they stay away",
    "2.14M customers, Sep 2026. Active = paid, transferred in, used ATM, pocket or coupon, booked or repaid a loan, or joined a mission in the month. Group sizes reconciled across the draft's two activeness tables: see data notes, A5. " + SRC);
  const MX = 0.5, RHW = 1.85, CW = 2.08, G = 0.08, y0 = 1.62, HH = 0.62, r0 = y0 + HH + G, RH = 1.08;
  const cx = [0, 1, 2].map((i) => MX + RHW + G + i * (CW + G));
  const ry = [0, 1, 2].map((i) => r0 + i * (RH + G));
  [["Got a loan", "158K · 7% of customers"], ["Applied, no loan", "1.46M · 68%"], ["Never applied", "528K · 25%"]].forEach((h, k) => {
    RR(s, cx[k], y0, CW, HH, k === 0 ? NAVY : (k === 1 ? BLUE : GREY), `Map col ${k}`, 0.05);
    T(s, [{ text: h[0], options: { bold: true, fontSize: 12, color: WHITE, breakLine: true } }, { text: h[1], options: { fontSize: 9.5, color: WHITE } }], { x: cx[k] + 0.12, y: y0, w: CW - 0.24, h: HH, valign: "middle", objectName: `Map col text ${k}` });
  });
  [["Keep money with us", "balance over ฿0"], ["Have an account,\nno money in it", "deposit account at ฿0"], ["No account yet", "signed up only"]].forEach((h, k) => {
    T(s, [{ text: h[0], options: { bold: true, fontSize: 11, color: NAVY, breakLine: true } }, { text: h[1], options: { fontSize: 9, color: MUTED } }], { x: MX, y: ry[k], w: RHW - 0.1, h: RH, valign: "middle", objectName: `Map row ${k}` });
  });
  const cell = (x, y, w, h, n, a, fill, dark, name, extra) => {
    RR(s, x, y, w, h, fill, name, 0.05);
    const col = dark ? WHITE : NAVY;
    T(s, [{ text: n, options: { bold: true, fontSize: 24, color: col, breakLine: true } }, { text: a, options: { bold: true, fontSize: 11.5, color: dark ? WHITE : TXT, breakLine: !!extra } }].concat(extra ? [{ text: extra, options: { fontSize: 9, color: dark ? "DCE6EF" : MUTED } }] : []),
      { x: x + 0.14, y, w: w - 0.28, h, valign: "middle", objectName: name + " text" });
  };
  cell(cx[0], ry[0], CW, 3 * RH + 2 * G, "158K", "88% active", CYAN, true, "Cell loan", "includes repaying the loan");
  cell(cx[1], ry[0], CW, RH, "161K", "50% active", "8FD0C9", false, "Cell applied money");
  cell(cx[1], ry[1], CW, RH, "924K", "4% active", "DDE3EA", false, "Cell applied no money");
  cell(cx[1], ry[2], CW, RH, "371K", "not active", "EEF2F6", false, "Cell applied no account");
  cell(cx[2], ry[0], CW, RH, "92K", "52% active", "8FD0C9", false, "Cell never money");
  cell(cx[2], ry[1], CW, RH, "113K", "7% active", "DDE3EA", false, "Cell never no money");
  cell(cx[2], ry[2], CW, RH, "323K", "not active", "EEF2F6", false, "Cell never no account");
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: cx[1] - 0.05, y: ry[1] - 0.05, w: CW + 0.1, h: 2 * RH + G + 0.1, rectRadius: 0.07, fill: { type: "none" }, line: { color: TRUEC, width: 2.25 }, objectName: "Left with nothing outline" });
  T(s, [{ text: "1.3M · 60% of customers  ", options: { bold: true, color: TRUEC } }, { text: "came for cash and left with nothing", options: { color: TXT } }],
    { x: cx[0], y: ry[2] + RH + 0.08, w: 3 * CW + 2 * G, h: 0.26, fontSize: 10.5, align: "center", valign: "middle", objectName: "Left with nothing label" });
  const RX = 9.05, RW = 12.83 - 9.05;
  RR(s, RX, y0, RW, 4.42, C.background2, "Map side");
  T(s, "WHAT THE MAP TELLS US", { x: RX + 0.2, y: y0 + 0.15, w: RW - 0.4, h: 0.24, fontSize: 9.5, bold: true, color: MUTED, charSpacing: 1, objectName: "Map side head" });
  [["312K", "are active each month, 15% of sign-ups: 139K borrowers, 128K with money in and 45K others", BLUE],
    ["50% vs 4%", "active with money in vs without, among those who applied and got no loan", CYAN],
    ["1.3M", "applied, got no loan and keep no money with us: the largest group in the base", TRUEC]].forEach((r, k) => {
    const yy = y0 + 0.5 + k * 1.28;
    if (k) s.addShape(pres.shapes.LINE, { x: RX + 0.2, y: yy - 0.04, w: RW - 0.4, h: 0, line: { color: LGREY, width: 0.75 }, objectName: `Map side rule ${k}` });
    T(s, [{ text: r[0], options: { bold: true, fontSize: 22, color: r[2], breakLine: true } }, { text: r[1], options: { fontSize: 10.5, color: TXT } }], { x: RX + 0.2, y: yy, w: RW - 0.4, h: 1.2, valign: "middle", objectName: `Map side ${k}` });
  });
  soWhat(s, "Two levers follow: get more applicants a first loan safely, and get money in first. A balance lifts activity from 4% to 50% even without a loan");
}

// ---------- 5 Why we cannot simply lend more
{
  const s = mk(1, "5 / 11  |  THE LENDING CONSTRAINT", "We cannot simply lend more: the customers we targeted are the ones we approve least, and early default is already 20%",
    "Applied = ever applied for a loan in the app; got a loan = loan booked; % of each job group on CLICX, Sep 2026. Early default (FPDX) on loans to KTB Jai-pump rejects, Jun 2025–May 2026. Underserved = NCB classification, 27 Oct 2026. " + SRC);
  const y = 1.62, H = 4.55, LW = 6.6, RX = 7.25, RW = 12.83 - 7.25;
  RR(s, 0.5, y, LW, H, C.background2, "Apply panel");
  head(s, 0.5, y, LW, 1, "Almost everyone applies; few get a loan", "% of each job group on CLICX; customers who got a loan in brackets");
  legend(s, 0.72, y + 0.95, [["Applied for a loan", SOFT], ["Got a loan", NAVY]]);
  const lj = [["White collar", 85, 11, "79K"], ["Blue collar", 92, 7, "30K"], ["Merchant", 84, 7, "20K"], ["Freelancer", 88, 6, "13K"], ["Rider, driver", 91, 5, "3K"]];
  const lx = 0.72, lw = 1.35, bw = LW - 0.44 - lw - 1.1;
  lj.forEach((r, k) => {
    const yy = y + 1.32 + k * 0.6;
    T(s, r[0], { x: lx, y: yy, w: lw, h: 0.52, fontSize: 10.5, bold: k > 1, color: k > 1 ? CYAN : TXT, valign: "middle", objectName: `Job ${k}` });
    R(s, lx + lw, yy + 0.04, bw * r[1] / 100, 0.2, SOFT, `Applied ${k}`);
    T(s, r[1] + "%", { x: lx + lw + bw * r[1] / 100 + 0.06, y: yy, w: 0.6, h: 0.28, fontSize: 9.5, color: MUTED, valign: "middle", objectName: `Applied value ${k}` });
    R(s, lx + lw, yy + 0.27, bw * r[2] / 100, 0.2, NAVY, `Booked ${k}`);
    T(s, [{ text: r[2] + "%", options: { bold: true, color: NAVY } }, { text: "  (" + r[3] + ")", options: { color: MUTED } }], { x: lx + lw + bw * r[2] / 100 + 0.06, y: yy + 0.23, w: 1.2, h: 0.28, fontSize: 10, valign: "middle", objectName: `Booked value ${k}` });
  });
  T(s, "Teal = target groups without a steady income. They apply as much as anyone and get a loan least", { x: 0.7, y: y + H - 0.32, w: LW - 0.4, h: 0.24, fontSize: 9.5, italic: true, color: MUTED, valign: "middle", objectName: "Apply note" });
  RR(s, RX, y, RW, H, C.background2, "Risk panel");
  head(s, RX, y, RW, 2, "Lending more is a question of risk, not demand");
  [["20.2%", "early default on loans to KTB's Jai-pump rejects", TRUEC],
    ["18% vs 21%", "early default for rejects referred by Paotang vs NEXT: where an applicant comes from matters", NAVY],
    ["7%", "of rider applicants from the reject list get a loan, though 64% of rider rejects apply", NAVY],
    ["82%", "of applicants are underserved (NCB): little credit history to approve on", NAVY]].forEach((r, k) => {
    const yy = y + 0.78 + k * 0.93;
    if (k) s.addShape(pres.shapes.LINE, { x: RX + 0.2, y: yy - 0.02, w: RW - 0.4, h: 0, line: { color: LGREY, width: 0.75 }, objectName: `Risk rule ${k}` });
    numRow(s, RX + 0.2, yy, RW - 0.4, 0.88, r[0], r[1], { nw: 1.75, nfs: 20, col: r[2] });
  });
  soWhat(s, "Approving more takes better selection, not looser rules: alternative data, smaller first loans, and applicants sourced through Paotang");
}

// ---------- 6 Who is active (treemap, payers outlined)
{
  const s = mk(1, "6 / 11  |  WHO IS ACTIVE", "Only 1 in 5 customers really uses CLICX; 127K of them pay in-app, spread across every segment",
    "Box size = customers · orange outline = paid in-app in Oct · active = balance over ฿100, payment or loan · 1.51M customers in the product data, 31 Oct 2026 (see data notes, A5)");
  const TOTAL = 1507.4, ZERO = 1061.3, MONEY = 136.7, ACT = 309.4;
  const X0 = 0.5, W = 8.95;
  const sy = 1.62, sh = 0.46, zw = W * ZERO / TOTAL, mw = W * MONEY / TOTAL, aw = W - zw - mw;
  T(s, [{ text: "No product, ฿0 balance ", options: { bold: true } }, { text: "1.06M · 70%" }], { x: X0, y: sy, w: zw, h: sh, fontSize: 11, color: WHITE, fill: { color: GREY }, line: { color: WHITE, width: 1.5 }, margin: 7, valign: "middle", objectName: "Strip zero" });
  T(s, [{ text: "137K · 9%", options: { bold: true, fontSize: 9, breakLine: true } }, { text: "bal >฿0 & ≤฿100", options: { fontSize: 7.5 } }], { x: X0 + zw, y: sy, w: mw, h: sh, fontSize: 9, color: TXT, align: "center", fill: { color: "C8CFD6" }, line: { color: WHITE, width: 1.5 }, valign: "middle", objectName: "Strip money" });
  T(s, [{ text: "309K ", options: { bold: true } }, { text: "20.5%" }], { x: X0 + zw + mw, y: sy, w: aw, h: sh, fontSize: 10.5, color: WHITE, align: "center", fill: { color: "3A8A80" }, line: { color: WHITE, width: 1.5 }, valign: "middle", objectName: "Strip active" });
  const ty = 2.4, th = 3.28;
  const fam = [["Borrowers", 164.478, NAVY], ["Savers", 105.4, CYAN], ["Payers only", 39.5, MUTED]];
  const cwF = fam.map((f) => W * f[1] / ACT);
  const cx = [X0, X0 + cwF[0], X0 + cwF[0] + cwF[1]];
  fam.forEach((f, a) => {
    const lab = a === 2 ? [{ text: f[0], options: { bold: true } }] : [{ text: f[0], options: { bold: true } }, { text: "  " + Math.round(f[1]) + "K · " + Math.round(100 * f[1] / ACT) + "%" }];
    T(s, lab, { x: cx[a] + 0.02, y: ty - 0.27, w: cwF[a], h: 0.24, fontSize: 10, color: f[2], valign: "bottom", objectName: `Family ${a}` });
  });
  const box = (x, y, w, h, fill, txtCol, runs, pay, obj, opts = {}) =>
    s.addText(runs, { x, y, w, h, color: txtCol, fill: { color: fill }, line: { color: pay ? ORANGE : WHITE, width: pay ? 2.25 : 2 }, margin: opts.margin ?? 8, valign: opts.valign || "top", align: opts.align || "left", isTextBox: true, objectName: obj });
  const Hh = (v, f) => th * v / f;
  let callB, callS;
  { const x = cx[0], w = cwF[0], f = 164.478; let y = ty;
    const h1 = Hh(102.7, f), h2 = Hh(40.5, f), h3 = Hh(21.3, f);
    box(x, y, w, h1, NAVY, WHITE, [{ text: "Borrower only", options: { bold: true, fontSize: 14, breakLine: true } }, { text: "103K", options: { bold: true, fontSize: 30 } }, { text: " 33%", options: { fontSize: 11, breakLine: true } }, { text: "Mostly salaried, aged 20–39; 6 in 10 keep ฿0", options: { fontSize: 11 } }], false, "Tree Borrower only");
    y += h1;
    box(x, y, w, h2, "3A5A85", WHITE, [{ text: "Borrower + transactor", options: { bold: true, fontSize: 12, breakLine: true } }, { text: "41K", options: { bold: true, fontSize: 18 } }, { text: " 13% · salaried, 30–39", options: { fontSize: 10.5 } }], true, "Tree Borrower + transactor", { valign: "middle" });
    y += h2;
    const w1 = w * 11.0 / 21.3, w2 = w - w1;
    box(x, y, w1, h3, "5374A8", WHITE, [{ text: "11.0K · 4%", options: { bold: true, fontSize: 12 } }], false, "Tree Saver + borrower", { align: "center", valign: "middle", margin: 2 });
    box(x + w1, y, w2, h3, "5374A8", WHITE, [{ text: "10.3K · 3%", options: { bold: true, fontSize: 12 } }], true, "Tree Saver + borrower who pay", { align: "center", valign: "middle", margin: 2 });
    callB = [[x + w1 / 2, "Saver + borrower", "Salaried, aged 30–49"], [x + w1 + w2 / 2, "Saver + borrower who pay", "Salaried, aged 30–39"]];
  }
  { const x = cx[1], w = cwF[1], f = 105.4; let y = ty;
    const h1 = Hh(50.3, f), h2 = Hh(31.9, f), h3 = Hh(23.3, f);
    box(x, y, w, h1, "4BA394", WHITE, [{ text: "Emerging saver only", options: { bold: true, fontSize: 14, breakLine: true } }, { text: "50K", options: { bold: true, fontSize: 30 } }, { text: " 16%", options: { fontSize: 11, breakLine: true } }, { text: "Salaried and merchants, 40–59", options: { fontSize: 11 } }], false, "Tree Emerging saver only");
    y += h1;
    box(x, y, w, h2, "66B3A6", WHITE, [{ text: "Emerging saver + transactor", options: { bold: true, fontSize: 12, breakLine: true } }, { text: "32K", options: { bold: true, fontSize: 18 } }, { text: " 10%", options: { fontSize: 10.5, breakLine: true } }, { text: "Salaried and merchants, 30–49", options: { fontSize: 10.5 } }], true, "Tree Emerging saver + transactor");
    y += h2;
    const w1 = w * 18.6 / 23.3, w2 = w - w1;
    box(x, y, w1, h3, "9FCFC7", TXT, [{ text: "Core saver", options: { bold: true, fontSize: 12, breakLine: true } }, { text: "18.6K", options: { bold: true, fontSize: 18 } }, { text: " 6% · aged 40+", options: { fontSize: 10.5 } }], false, "Tree Core saver", { valign: "middle" });
    box(x + w1, y, w2, h3, "9FCFC7", TXT, [{ text: "4.7K", options: { bold: true, fontSize: 11.5, breakLine: true } }, { text: "2%", options: { fontSize: 10 } }], true, "Tree Core savers who pay", { align: "center", valign: "middle", margin: 2 });
    callS = [[x + w1 + w2 / 2, "Core savers who pay", "Salaried, aged 30–49"]];
  }
  box(cx[2], ty, cwF[2], th, "8395AB", WHITE, [{ text: "Pays only", options: { bold: true, fontSize: 12, breakLine: true } }, { text: "39K", options: { bold: true, fontSize: 18 } }, { text: " 13%", options: { fontSize: 10.5, breakLine: true } }, { text: "Wage, gig and online sellers", options: { fontSize: 10.5 } }], true, "Tree Pays only");
  const cy = ty + th;
  [...callB, ...callS].forEach((c, k) => {
    s.addShape(pres.shapes.LINE, { x: c[0], y: cy, w: 0, h: 0.08, line: { color: TXT, width: 1 }, objectName: `Callout tick ${k}` });
    const lw = 2.2, lx = k === 2 ? c[0] - lw + 0.05 : c[0] - lw / 2;
    T(s, [{ text: c[1], options: { bold: true, breakLine: true } }, { text: c[2], options: { color: MUTED } }], { x: lx, y: cy + 0.09, w: lw, h: 0.38, fontSize: 9.5, color: TXT, align: k === 2 ? "right" : "center", valign: "top", objectName: `Callout ${k}` });
  });
  const RX = 9.7, RW = 3.13;
  RR(s, RX, sy, RW, 0.9, TINT, "Active stat");
  T(s, [{ text: "309K", options: { bold: true, fontSize: 28, color: NAVY, breakLine: true } }, { text: "active: 20% of 1.51M", options: { fontSize: 11, color: TXT } }], { x: RX + 0.15, y: sy + 0.04, w: RW - 0.3, h: 0.82, valign: "middle", objectName: "Active stat text" });
  [["127K pay in-app (41%)", "88K also borrow or save", ORANGE], ["164K borrowers (53%)", "6 in 10 keep ฿0: sell savings", NAVY],
    ["105K savers (34%)", "1 in 5 core savers pays in-app", CYAN], ["137K hold ฿1–100 only", "No loan, no payment", "C8CFD6"]].forEach((t, k) => {
    const y = 2.72 + k * 0.7;
    R(s, RX, y, 0.05, 0.56, t[2], `Takeaway bar ${k}`);
    T(s, [{ text: t[0], options: { bold: true, fontSize: 12, color: NAVY, breakLine: true } }, { text: t[1], options: { fontSize: 10.5, color: TXT } }], { x: RX + 0.17, y, w: RW - 0.2, h: 0.56, valign: "middle", objectName: `Takeaway ${k}` });
  });
  soWhat(s, "Most active customers do one thing with us. The next step is a second product: savings for borrowers, and a deposit for those who only pay");
}

// ---------- 7 Payments
{
  const s = mk(1, "7 / 11  |  PAYMENTS", "3 in 4 payers come through AIS or OR; 19% of them stopped after the August deal, but those who stay pay 39% more",
    "Partner payers = ever paid AIS or PTT OR. 19% = drop in partner payers Aug → Oct (114.3K → 93.2K); 39% = spend per partner payer, ฿2.1K → ฿3.0K; +22K assumes prepaid pays at the postpaid rate. Points 1–2: Partner Payment Story (8.4% of 1.51M paid in Oct); point 3: Oct segment extract (7.8% of 1.50M); gig and wage = freelance, rider, wage earner");
  const cw = (12.33 - 0.3) / 3, top = 1.62, ch = 5.15;
  const cols = [
    { head: "Partners start payments, but promotions don't keep them", big: "73%", lbl: "of October payers came through AIS or OR, making 67% of spend", so: "Keep AIS and OR promotions as the way in; reward the second and third month" },
    { head: "AIS prepaid is the biggest payment gap", big: "516.5K", lbl: "AIS prepaid customers, only 8.4% pay", so: "Top-up-first offers for prepaid: up to +22K payers a month", note: "The gap is mostly bill pay and top-up: 4.6% vs 7.8%" },
    { head: "Teens, merchants and 30–49s pay most; blue-collar pay least", big: "13.5%", lbl: "of 10–19s paid in Oct; merchants 9.1–9.6%, vs 7.8% of all customers", so: "Aim payment offers at merchants, salaried 30–49s and students; give blue-collar and gig workers an easy first payment", note: "Salaried and merchants aged 30–49 are where the volume is" },
  ];
  cols.forEach((c, i) => {
    const x = 0.5 + i * (cw + 0.15);
    RR(s, x, top, cw, ch, C.background2, `Pay col ${i}`);
    badge(s, x + 0.16, top + 0.16, i + 1, NAVY, 0.4, 12.5);
    T(s, c.head, { x: x + 0.66, y: top + 0.1, w: cw - 0.8, h: 0.58, fontSize: 12.5, bold: true, color: NAVY, valign: "middle", objectName: `Pay head ${i}` });
    RR(s, x + 0.16, top + 0.8, cw - 0.32, 0.72, TINT, `Pay stat box ${i}`, 0.06);
    T(s, [{ text: c.big, options: { bold: true, fontSize: 20, color: BLUE, breakLine: true } }, { text: c.lbl, options: { fontSize: 10, color: TXT } }], { x: x + 0.28, y: top + 0.8, w: cw - 0.56, h: 0.72, valign: "middle", objectName: `Pay stat ${i}` });
    if (c.note) T(s, c.note, { x: x + 0.18, y: top + 4.02, w: cw - 0.36, h: 0.38, fontSize: 9.5, italic: true, color: MUTED, valign: "middle", objectName: `Pay note ${i}` });
    RR(s, x + 0.16, top + 4.45, cw - 0.32, 0.58, WHITE, `Pay so box ${i}`, 0.06);
    T(s, [{ text: "SO  ", options: { bold: true, color: CYAN, fontSize: 9.5, charSpacing: 1 } }, { text: c.so, options: { color: TXT, fontSize: 10 } }], { x: x + 0.28, y: top + 4.45, w: cw - 0.56, h: 0.58, valign: "middle", objectName: `Pay so ${i}` });
  });
  // chart 1: partner share of payers and spend
  { const x = 0.5;
    T(s, "Partner customers' share of payers and of spend", { x: x + 0.2, y: top + 1.6, w: cw - 0.4, h: 0.24, fontSize: 10.5, bold: true, color: NAVY, align: "center", objectName: "Chart 1 title" });
    ccols(s, x + 0.3, top + 2.08, cw - 0.6, 1.1, ["Jun", "Jul", "Aug", "Sep", "Oct"], [{ name: "Share of payers", color: NAVY, vals: [57, 67, 70, 73, 73] }, { name: "Share of spend", color: CYAN, vals: [51, 62, 66, 67, 67] }], { max: 80, fmt: (v) => v + "%", name: "Share" });
    legend(s, x + 0.95, top + 3.46, [["Share of payers", NAVY], ["Share of spend", CYAN]], 9);
    const hd = ["", "Jun", "Jul", "Aug", "Sep", "Oct"].map((t) => ({ text: t, options: { bold: true, color: MUTED, align: "right" } }));
    const rows = [["Partner payers", "12.8K", "31.6K", "114.3K", "95.0K", "93.2K"], ["Their spend", "฿20M", "฿88M", "฿244M", "฿262M", "฿275M"]].map((r) => r.map((t, j) => ({ text: t, options: { bold: j === 0, color: j === 0 ? NAVY : TXT, align: j === 0 ? "left" : "right" } })));
    const fw = 1.0, ow = (cw - 0.36 - fw) / 5;
    s.addTable([hd, ...rows], { x: x + 0.18, y: top + 3.78, w: cw - 0.36, colW: [fw, ow, ow, ow, ow, ow], rowH: 0.18, fontSize: 8.5, fontFace: "Arial", margin: [0, 0.04, 0, 0.04], valign: "middle", border: { type: "none" }, objectName: "Pay table" });
  }
  // chart 2: who paid, by partner plan
  { const x = 0.5 + cw + 0.15;
    T(s, "% who paid in October", { x: x + 0.2, y: top + 1.6, w: cw - 0.4, h: 0.24, fontSize: 10.5, bold: true, color: NAVY, align: "center", objectName: "Chart 2 title" });
    hbars(s, x + 0.25, top + 1.92, cw - 0.5, [["AIS postpaid", 12.8, "12.8%", NAVY], ["AIS prepaid", 8.4, "8.4%", CYAN], ["PTT + Amazon", 12.0, "12.0%", BLUE], ["Amazon only", 10.7, "10.7%", "5A86BE"], ["PTT only", 8.3, "8.3%", SOFT]],
      { lw: 1.3, rh: 0.36, max: 15, fs: 10, name: "Plan", labelFill: PANELHEX, valW: 0.52, bench: { v: 8.4, label: "All customers 8.4%" } });
  }
  // chart 3: who paid, by age and job
  { const x = 0.5 + 2 * (cw + 0.15);
    T(s, "% who paid in October, by age and job", { x: x + 0.2, y: top + 1.6, w: cw - 0.4, h: 0.24, fontSize: 10.5, bold: true, color: NAVY, align: "center", objectName: "Chart 3 title" });
    const G2 = "B8C2CC";
    hbars(s, x + 0.25, top + 1.9, cw - 0.5, [["Age 10–19", 13.5, "13.5%", CYAN], ["Age 20–29", 7.0, "7.0%", G2], ["Age 30–49", 8.1, "8.1%", CYAN], ["Age 50–69", 7.6, "7.6%", G2], ["Online merchant", 9.6, "9.6%", CYAN],
      ["Offline merchant", 9.1, "9.1%", CYAN], ["Salaried", 8.6, "8.6%", CYAN], ["Gig and wage", 6.8, "6.8%", G2], ["Blue collar", 5.9, "5.9%", G2]],
      { lw: 1.35, rh: 0.2, max: 16, fs: 9, name: "AgeJob", labelFill: PANELHEX, valW: 0.48, bench: { v: 7.8, label: "All customers 7.8%" } });
  }
}

// ---------- 8 True
{
  const s = mk(2, "8 / 11  |  DEFEND AGAINST TRUE", "True's new virtual bank puts ฿1.98B (63%) of our deposits at risk; defend first where KTB customers already use TrueMoney",
    "AIS no TrueMoney = AIS customers who don't use TrueMoney Wallet; True = True and TrueMoney customers (561,177 and 1,257,522 customers; 1.82M in all). Tech-savvy = uses Make, Kept, Dime or crypto. TrueMoney use and Jai-pump rejects from KTB data. " + SRC);
  const y = 1.62, H = 4.55, PW = (12.33 - 0.15) / 2;
  RR(s, 0.5, y, PW, H, C.background2, "Deposit panel");
  head(s, 0.5, y, PW, 1, "True and TrueMoney customers hold almost two-thirds of our deposits", "CLICX customers and balances", TRUEC);
  const bx = 1.85, bw = PW - 1.75;
  [["Customers", 0.3086, "561K · 31%", "1.26M · 69%"], ["Deposits", 0.37, "฿1.17B · 37%", "฿1.98B · 63%"]].forEach((r, k) => {
    const yy = y + 1.0 + k * 0.7;
    T(s, r[0], { x: 0.7, y: yy, w: 1.1, h: 0.56, fontSize: 11, bold: true, color: TXT, valign: "middle", objectName: `Split label ${k}` });
    s.addShape(pres.shapes.RECTANGLE, { x: bx, y: yy, w: bw * r[1], h: 0.56, fill: { color: NAVY }, line: { color: WHITE, width: 1.5 }, objectName: `AIS part ${k}` });
    s.addShape(pres.shapes.RECTANGLE, { x: bx + bw * r[1], y: yy, w: bw * (1 - r[1]), h: 0.56, fill: { color: TRUEC }, line: { color: WHITE, width: 1.5 }, objectName: `True part ${k}` });
    T(s, r[2], { x: bx + 0.1, y: yy, w: bw * r[1] - 0.15, h: 0.56, fontSize: 10.5, bold: true, color: WHITE, valign: "middle", objectName: `AIS value ${k}` });
    T(s, r[3], { x: bx + bw * r[1], y: yy, w: bw * (1 - r[1]), h: 0.56, fontSize: 10.5, bold: true, color: WHITE, align: "center", valign: "middle", objectName: `True value ${k}` });
  });
  legend(s, bx, y + 2.45, [["AIS, no TrueMoney", NAVY], ["True and TrueMoney", TRUEC]]);
  [["67% vs 58%", "of True vs AIS customers are under 40: True's base is younger"], ["33% vs 20%", "are tech-savvy: more digital, so easier to lose"]].forEach((f, k) => {
    RR(s, 0.7, y + 2.9 + k * 0.76, PW - 0.4, 0.66, WHITE, `True fact box ${k}`);
    numRow(s, 0.85, y + 2.9 + k * 0.76, PW - 0.7, 0.66, f[0], f[1], { nw: 1.5, nfs: 15, col: TRUEC });
  });
  const X2 = 0.5 + PW + 0.15, fw2 = PW - 0.4;
  RR(s, X2, y, PW, H, C.background2, "Overlap panel");
  head(s, X2, y, PW, 2, "Act first where KTB and True overlap", "KTB customers in the data room", TRUEC);
  RR(s, X2 + 0.2, y + 0.98, fw2, 0.5, NAVY, "KTB base");
  T(s, [{ text: "35M  ", options: { bold: true, fontSize: 14 } }, { text: "KTB customers  →  ", options: { fontSize: 11 } }, { text: "2.5M (7%) ", options: { bold: true, fontSize: 14 } }, { text: "use TrueMoney", options: { fontSize: 11 } }],
    { x: X2 + 0.35, y: y + 0.98, w: fw2 - 0.3, h: 0.5, color: WHITE, valign: "middle", objectName: "KTB base text" });
  const hw = (fw2 - 0.15) / 2;
  [["Already on CLICX", "225K", "144K KTB-active · 81K not", TINT], ["Not on CLICX", "2.31M", "1.65M KTB-active · 660K not", WHITE]].forEach((b, k) => {
    const x = X2 + 0.2 + k * (hw + 0.15);
    RR(s, x, y + 1.6, hw, 0.95, b[3], `Overlap box ${k}`);
    T(s, [{ text: b[0], options: { fontSize: 9.5, color: MUTED, breakLine: true } }, { text: b[1], options: { bold: true, fontSize: 18, color: NAVY, breakLine: true } }, { text: b[2], options: { fontSize: 9, color: TXT } }], { x: x + 0.12, y: y + 1.6, w: hw - 0.24, h: 0.95, valign: "middle", objectName: `Overlap text ${k}` });
  });
  T(s, "TURNED DOWN BY JAI-PUMP, NOT ON CLICX", { x: X2 + 0.2, y: y + 2.66, w: fw2, h: 0.22, fontSize: 9, bold: true, color: TRUEC, charSpacing: 1, objectName: "Priority head" });
  [["Priority 1", "99K", "KTB-inactive: easy to win, no KTB relationship to lose"], ["Priority 2", "214K", "KTB-active: protect them from True"]].forEach((p, k) => {
    const yy = y + 2.92 + k * 0.76;
    RR(s, X2 + 0.2, yy, fw2, 0.66, WHITE, `Priority box ${k}`);
    T(s, [{ text: p[0] + "   ", options: { bold: true, fontSize: 10, color: TRUEC } }, { text: p[1] + "  ", options: { bold: true, fontSize: 16, color: NAVY } }, { text: p[2], options: { fontSize: 10, color: TXT } }], { x: X2 + 0.35, y: yy, w: fw2 - 0.3, h: 0.66, valign: "middle", objectName: `Priority text ${k}` });
  });
  soWhat(s, "Hold deposits with savings rewards and a credit card for young True customers; offer loans to the 313K TrueMoney users KTB turned down");
}

// ---------- 9 The plan: plays sized by customers in reach
const PLAYS = [
  ["ACQUIRE", 1, "Make KTB's Paotang reject list a standing channel", 1900, "1.9M a year", "24% of rejects already apply; Paotang referrals default less (18% vs 21%)"],
  ["ACQUIRE", 2, "Reach TrueMoney users KTB turned down", 313, "313K", "Not on CLICX yet; 99K are KTB-inactive, so KTB has nothing to lose"],
  ["ACTIVATE", 3, "A first deposit or first loan for those who left with nothing", 1295, "1.3M", "With money in, 50% are active, against 4% without"],
  ["ACTIVATE", 4, "Top-up-first offers for AIS prepaid", 516.5, "516K", "Prepaid pay at 8.4% vs 12.8% for postpaid: up to +22K payers a month"],
  ["ACTIVATE", 5, "Win back payers who stopped", 114, "114K", "53K paid in Sep but not Oct; 61K paid only in the August promotion"],
  ["GROW", 6, "Savings for customers who only borrow", 103, "103K", "6 in 10 keep ฿0 with us; only 21K both borrow and save"],
  ["GROW", 7, "SME loan for merchants", 106, "106K", "Applied without getting a loan and earn over ฿30K; merchants open the app most"],
  ["GROW", 8, "Digital credit card", 1100, "1.1M", "Earn over ฿15K, aged 20–44; 85% have a high need for cash"],
  ["GROW", 9, "Investing and wealth", 528, "528K", "Tech-savvy customers; 110K already hold over ฿20K or Save Max/More"],
  ["DEFEND", 10, "Hold the deposits of True and TrueMoney customers", 1258, "1.26M · ฿1.98B", "63% of our deposits; 67% are under 40 and 33% tech-savvy"],
];
{
  const s = mk(3, "9 / 11  |  THE PLAN: FOUR MOVES", "Four moves for 2027, sized by customers in reach: the biggest pool is already ours, the 1.3M who left with nothing",
    "Customers in reach = size of the group each play addresses, not a forecast; pools overlap. Evidence from slides 2–8. " + SRC);
  const y0 = 1.62, rh = 0.435, xs = [0.5, 1.75, 6.25, 9.25];
  ["MOVE", "PLAY", "CUSTOMERS IN REACH", "WHY WE BELIEVE IT"].forEach((h, k) => T(s, h, { x: xs[k] + (k === 1 ? 0.38 : 0), y: y0, w: 3, h: 0.24, fontSize: 9, bold: true, color: MUTED, charSpacing: 1, objectName: `Plan head ${k}` }));
  const BW = 1.75;
  let prev = null, gStart = 0;
  const groups = {};
  PLAYS.forEach((p, k) => { (groups[p[0]] = groups[p[0]] || []).push(k); });
  Object.entries(groups).forEach(([m, idx]) => {
    const ya = y0 + 0.3 + idx[0] * rh, hh = idx.length * rh - 0.06;
    s.addText(m, { x: 0.5, y: ya, w: 1.1, h: hh, shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.06, fill: { color: MOVE[m] }, line: { type: "none" }, fontSize: 9.5, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0, charSpacing: 1, objectName: `Plan move ${m}` });
    if (idx[0]) s.addShape(pres.shapes.LINE, { x: 1.75, y: ya - 0.03, w: 11.08, h: 0, line: { color: LGREY, width: 1 }, objectName: `Plan rule ${m}` });
  });
  PLAYS.forEach((p, k) => {
    const yy = y0 + 0.3 + k * rh, hh = rh - 0.06, col = MOVE[p[0]];
    badge(s, xs[1], yy + hh / 2 - 0.13, p[1], col, 0.26, 9);
    T(s, p[2], { x: xs[1] + 0.38, y: yy, w: 4.0, h: hh, fontSize: 10.5, bold: true, color: NAVY, valign: "middle", objectName: `Play ${p[1]}` });
    const ww = Math.max(0.05, BW * p[3] / 1900);
    R(s, xs[2], yy + hh / 2 - 0.1, ww, 0.2, col, `Play bar ${p[1]}`);
    T(s, p[4], { x: xs[2] + ww + 0.07, y: yy, w: 1.5, h: hh, fontSize: 10.5, bold: true, color: NAVY, valign: "middle", objectName: `Play size ${p[1]}` });
    T(s, p[5], { x: xs[3], y: yy, w: 12.83 - xs[3], h: hh, fontSize: 9.5, color: TXT, valign: "middle", objectName: `Play why ${p[1]}` });
  });
  soWhat(s, "Start with the plays that use what we already have: the reject list, top-up, win-back and savings. New credit and new products follow as they are approved", 6.32);
}

// ---------- 10 Next products
{
  const s = mk(3, "10 / 11  |  THE PLAN: NEXT PRODUCTS", "Customers already point to our next products: a credit card for 1.1M, investing for 528K and SME loans for 106K merchants",
    "Pools overlap. Tech-savvy = uses Make, Kept, Dime or crypto (AIS and KTB data). Income from CLICX profile. Sep 2026. " + SRC);
  const cards = [
    ["card.png", NAVY, "Digital credit card", "1.1M", "earn over ฿15K a month, aged 20–44: 482K salaried, 637K without a steady income",
      ["57% have no steady income", "85% have a high need for cash", "41% (489K) pay for streaming", "66% are AIS and 38% OR customers"],
      "Alternative-data approval · 0% micro-instalments for bills, PTT fuel and daily spend · streaming subscription perks"],
    ["invest.png", CYAN, "Investing and wealth", "528K", "tech-savvy customers: 63K under 24, 379K aged 25–44, 86K aged 45–59",
      ["110K already hold over ฿20K or Save Max/More", "71–73% are AIS, KTB or OR customers", "Under-24s: high cash need, low wealth, pay more"],
      "Under 24: small-ticket, low-risk investing, learn-and-earn · 25–44: zero-fee fractional shares in ADVANC, OR, KTB, GULF · 45–59: pre-retirement plans with insurance"],
    ["sme.png", BLUE, "SME loan for merchants", "288K", "merchants: 102K online, 186K offline; half earn over ฿30K a month",
      ["222K applied without getting a loan in all", "106K applied for a loan without getting one and earn over ฿30K: first wave", "More active: 44% open the app vs 36% of others"],
      "Loan based on sales and partner data · start in Isan (21%), Greater Bangkok (16%) and Bangkok (15%)"],
  ];
  const cw = (12.33 - 0.3) / 3, y = 1.62, H = 4.55;
  cards.forEach((c, k) => {
    const x = 0.5 + k * (cw + 0.15);
    RR(s, x, y, cw, H, C.background2, `Product card ${k}`);
    circIcon(s, x + 0.2, y + 0.15, 0.48, c[0], c[1], `Product icon ${k}`);
    T(s, c[2], { x: x + 0.8, y: y + 0.15, w: cw - 1.0, h: 0.48, fontSize: 15, bold: true, color: NAVY, valign: "middle", objectName: `Product name ${k}` });
    RR(s, x + 0.2, y + 0.76, cw - 0.4, 0.95, TINT, `Focus box ${k}`);
    T(s, [{ text: c[3], options: { bold: true, fontSize: 24, color: BLUE, breakLine: true } }, { text: c[4], options: { fontSize: 9.5, color: TXT } }], { x: x + 0.32, y: y + 0.78, w: cw - 0.64, h: 0.91, valign: "middle", objectName: `Focus ${k}` });
    T(s, "WHAT WE KNOW", { x: x + 0.2, y: y + 1.82, w: 2, h: 0.2, fontSize: 9, bold: true, color: MUTED, charSpacing: 1, objectName: `Know head ${k}` });
    T(s, c[5].map((t, i) => ({ text: t, options: { bullet: { indent: 12 }, breakLine: i < c[5].length - 1 } })), { x: x + 0.2, y: y + 2.05, w: cw - 0.4, h: 1.18, fontSize: 10.5, color: TXT, valign: "top", paraSpaceAfter: 3, objectName: `Know ${k}` });
    T(s, "LEAD WITH", { x: x + 0.2, y: y + 3.28, w: 2, h: 0.2, fontSize: 9, bold: true, color: CYAN, charSpacing: 1, objectName: `Lead head ${k}` });
    RR(s, x + 0.2, y + 3.5, cw - 0.4, 0.9, WHITE, `Lead box ${k}`);
    T(s, c[6], { x: x + 0.32, y: y + 3.5, w: cw - 0.64, h: 0.9, fontSize: 10, bold: true, color: NAVY, valign: "middle", objectName: `Lead ${k}` });
  });
  soWhat(s, "Proposed order, by readiness: the SME loan first, since those merchants are identified and already applying; then the credit card; then investing");
}

// ---------- 11 Roadmap, scorecard, decisions
{
  const s = mk(3, "11 / 11  |  THE PLAN: PHASING AND SCORECARD", "We propose to start with what we already have, add new credit next, and track six numbers with the board",
    "Phasing is a proposal for discussion, not an approved plan. Numbers in circles refer to the plays on slide 9. Baselines: Sep–Oct 2026; early default on loans to KTB Jai-pump rejects. " + SRC);
  const y = 1.62, H = 4.5, LW = 7.65, pw = (LW - 0.3) / 3;
  const needs = ["Partner budget for top-up and repeat-use rewards", "Risk appetite for alternative-data loans; KTB data", "Order and funding of the new products"];
  const colOf = (n) => MOVE[PLAYS[n - 1][0]];
  const phases = [
    ["NOW", "First 90 days · uses what we already have", NAVY, [[1, "Paotang reject-list campaign with KTB"], [3, "First-deposit offer for the 1.3M who left with nothing"], [4, "Top-up-first offer for AIS prepaid"], [5, "Win-back for 114K lapsed payers"], [6, "Savings nudge for 103K borrowers"], [10, "Savings rewards for True and TrueMoney customers"]]],
    ["NEXT", "By mid-2027 · needs risk approval or KTB data", BLUE, [[3, "Smaller first loans approved on alternative data"], [7, "SME loan pilot with the 106K merchants"], [2, "Campaign to TrueMoney users KTB turned down"]]],
    ["LATER", "Second half of 2027 · new products", GREY, [[8, "Digital credit card"], [9, "Investing and wealth"]]],
  ];
  phases.forEach((p, k) => {
    const x = 0.5 + k * (pw + 0.15);
    RR(s, x, y, pw, H, C.background2, `Phase ${k}`);
    R(s, x, y, pw, 0.62, p[2], `Phase band ${k}`);
    T(s, [{ text: p[0], options: { bold: true, fontSize: 13, color: WHITE, charSpacing: 1, breakLine: true } }, { text: p[1], options: { fontSize: 8.5, color: WHITE } }], { x: x + 0.15, y, w: pw - 0.3, h: 0.62, valign: "middle", objectName: `Phase head ${k}` });
    p[3].forEach((it, i) => {
      const yy = y + 0.72 + i * 0.52;
      RR(s, x + 0.12, yy, pw - 0.24, 0.46, WHITE, `Phase item box ${k}-${i}`, 0.05);
      badge(s, x + 0.2, yy + 0.1, it[0], colOf(it[0]), 0.26, 9);
      T(s, it[1], { x: x + 0.54, y: yy, w: pw - 0.72, h: 0.46, fontSize: 9.5, color: TXT, valign: "middle", objectName: `Phase item ${k}-${i}` });
    });
    RR(s, x + 0.12, y + H - 0.6, pw - 0.24, 0.5, TINT, `Phase needs box ${k}`, 0.05);
    T(s, [{ text: "NEEDS  ", options: { bold: true, fontSize: 8.5, color: CYAN, charSpacing: 1 } }, { text: needs[k], options: { fontSize: 9, color: TXT } }], { x: x + 0.22, y: y + H - 0.6, w: pw - 0.44, h: 0.5, valign: "middle", objectName: `Phase needs ${k}` });
  });
  { const lx = 0.5 + 2 * (pw + 0.15) + 0.15;
    T(s, "CIRCLE COLOURS", { x: lx, y: y + 2.05, w: 2, h: 0.2, fontSize: 8, bold: true, color: MUTED, charSpacing: 0.5, objectName: "Move key head" });
    [["Acquire", NAVY], ["Activate", CYAN], ["Grow", BLUE], ["Defend", TRUEC]].forEach((m, i) => {
      s.addShape(pres.shapes.OVAL, { x: lx, y: y + 2.32 + i * 0.3, w: 0.18, h: 0.18, fill: { color: m[1] }, line: { type: "none" }, objectName: `Move key ${i}` });
      T(s, m[0], { x: lx + 0.26, y: y + 2.28 + i * 0.3, w: 1.5, h: 0.26, fontSize: 9.5, color: MUTED, valign: "middle", objectName: `Move key label ${i}` });
    }); }
  const RX = 8.3, RW = 12.83 - 8.3;
  RR(s, RX, y, RW, H, C.background2, "Scorecard");
  T(s, "Six numbers to track with the board", { x: RX + 0.2, y: y + 0.1, w: RW - 0.4, h: 0.44, fontSize: 13.5, bold: true, color: NAVY, valign: "middle", objectName: "Scorecard head" });
  T(s, "TODAY", { x: RX + 2.55, y: y + 0.58, w: 0.9, h: 0.2, fontSize: 8, bold: true, color: MUTED, align: "right", charSpacing: 0.5, objectName: "Scorecard col 1" });
  T(s, "2027 TARGET", { x: RX + 3.5, y: y + 0.58, w: 0.85, h: 0.2, fontSize: 8, bold: true, color: MUTED, align: "right", charSpacing: 0.5, objectName: "Scorecard col 2" });
  [["Active customers each month", "~310K"], ["Applicants who get a loan", "10%"], ["Early default on new loans", "20.2%"], ["Customers paying in-app each month", "127K"],
    ["Payers who pay again the next month", "60%"], ["Deposits held by True and TrueMoney customers", "63%"]].forEach((r, k) => {
    const yy = y + 0.82 + k * 0.6;
    s.addShape(pres.shapes.LINE, { x: RX + 0.2, y: yy, w: RW - 0.4, h: 0, line: { color: LGREY, width: 0.75 }, objectName: `Scorecard rule ${k}` });
    T(s, r[0], { x: RX + 0.2, y: yy, w: 2.35, h: 0.6, fontSize: 10, color: TXT, valign: "middle", objectName: `KPI ${k}` });
    T(s, r[1], { x: RX + 2.55, y: yy, w: 0.9, h: 0.6, fontSize: 14, bold: true, color: NAVY, align: "right", valign: "middle", objectName: `KPI today ${k}` });
    T(s, "to set", { x: RX + 3.5, y: yy, w: 0.85, h: 0.6, fontSize: 9.5, italic: true, color: GREY, align: "right", valign: "middle", objectName: `KPI target ${k}` });
  });
  RR(s, 0.5, 6.25, 12.33, 0.54, NAVY, "Decisions");
  T(s, [{ text: "BOARD DECISIONS  ", options: { bold: true, fontSize: 9.5, color: LCYAN, charSpacing: 1 } },
    { text: "1 Risk appetite for alternative-data first loans  ·  2 Partner budget for top-up and repeat-use rewards  ·  3 Order and funding of new products  ·  4 Unit economics in the data room", options: { fontSize: 10.5, color: WHITE } }],
    { x: 0.7, y: 6.25, w: 11.93, h: 0.54, valign: "middle", objectName: "Decisions text" });
}

// =====================================================================
pres.addSection({ title: "Appendix" });

// ---------- A1 Six segments
{
  const s = mk(null, "APPENDIX A1  |  SIX CUSTOMER SEGMENTS BY AGE", "30–44s are half the base; balances and saving rise with age, while loan demand peaks between 20 and 44",
    "Active = monthly transacting customer. Balance = average (median in brackets). Tech-savvy from AIS and KTB data. Segments sum to 2.0M; the rest have no age on file. Sep 2026. " + SRC, "Appendix");
  const hd = ["Segment", "Customers", "Active monthly", "Paying", "Save Max/ More", "Avg balance (median)", "Applied for a loan", "Tech-savvy", "What they look like"];
  const rows = [
    ["Student (under 20)", "14K", "21%", "18%", "9%", "฿1,019 (฿18)", "45%", "32%", "Wealth-aware minors: nurture as a future main bank"],
    ["Gen Z first jobber (20–24)", "169K", "13%", "10%", "4%", "฿612 (฿2)", "85%", "36%", "High cash need, tech-savvy: first credit card, first investing"],
    ["Gen Z career builder (25–29)", "350K", "15%", "12%", "5%", "฿1,419 (฿4)", "87%", "31%", "Wealth builders with a high need for cash"],
    ["Experienced worker (30–44)", "1,051K", "16%", "13%", "5%", "฿3,394 (฿10)", "83%", "28%", "The core of the base"],
    ["Working senior (45–59)", "369K", "17%", "13%", "7%", "฿7,902 (฿94)", "74%", "24%", "Affluent caregivers"],
    ["Senior (60+)", "39K", "14%", "11%", "12%", "฿35,202 (฿1,000)", "41%", "2%", "Capital preservers; high cost of funds"],
  ];
  const Hc = (t, al) => ({ text: t, options: { bold: true, fontSize: 10, color: WHITE, fill: { color: NAVY }, align: al, valign: "middle" } });
  const hi = { 2: [0], 4: [5], 5: [5], 6: [2], 7: [1] };
  const tbl = [hd.map((h, j) => Hc(h, j === 0 || j === 8 ? "left" : "center")), ...rows.map((r, i) => r.map((v, j) => ({ text: v, options: {
    fontSize: 10.5, color: TXT, bold: j === 0, align: j === 0 || j === 8 ? "left" : "center", valign: "middle",
    fill: { color: (hi[j] || []).includes(i) ? "CDEAE7" : (i % 2 ? "FFFFFF" : "F4F6F9") } } })))];
  s.addTable(tbl, { x: 0.5, y: 1.65, w: 12.33, colW: [2.05, 1.0, 0.95, 0.8, 0.9, 1.45, 0.95, 0.9, 3.33], rowH: [0.55, ...rows.map(() => 0.62)], fontFace: "Arial", border: { type: "solid", pt: 0.5, color: "E3E8EE" }, margin: [0.04, 0.08, 0.04, 0.08], objectName: "Segments table" });
  T(s, "Teal = highest in each column", { x: 0.5, y: 6.0, w: 4, h: 0.24, fontSize: 9.5, italic: true, color: MUTED, objectName: "Table key" });
}

// ---------- A2 KTB inactive vs all customers
{
  const s = mk(null, "APPENDIX A2  |  KTB-INACTIVE CUSTOMERS ON CLICX", "84% of the 1.1M KTB-inactive customers came to CLICX for a loan, against 73% of all customers",
    "Share of each base. Have balance = balance over ฿0. Loan applicants = applied but no loan booked. KTB-inactive base 1.06M; all customers 2.14M. CLICX data as of Sep 2026; KTB as of Apr 2026. " + SRC, "Appendix");
  const rows = [
    ["", "Borrower", 74, 7, 158, 7], ["", "Have balance", 26, 2, 92, 4], ["", "No balance", 54, 5, 111, 5], ["", "No deposit", 84, 8, 375, 17],
    ["Loan applicants (not booked)", "Have balance", 74, 7, 161, 8], ["", "No balance", 534, 50, 874, 41], ["", "No deposit", 216, 20, 371, 17]];
  const y0 = 1.95, rh = 0.6, lx = 0.5, gx = 2.6, bx = 4.4, bw = 4.8;
  legend(s, bx, 1.65, [["KTB-inactive (1.1M)", CYAN], ["All customers (2.1M)", "B8C2CC"]], 10);
  rows.forEach((r, k) => {
    const yy = y0 + k * rh + (k >= 4 ? 0.15 : 0);
    if (r[0]) T(s, r[0], { x: lx, y: yy, w: 1.9, h: rh * 3, fontSize: 11, bold: true, color: NAVY, valign: "middle", objectName: `A2 group ${k}` });
    T(s, r[1], { x: gx, y: yy, w: 1.7, h: rh, fontSize: 11, color: TXT, valign: "middle", objectName: `A2 label ${k}` });
    const w1 = bw * r[3] / 55, w2 = bw * r[5] / 55;
    R(s, bx, yy + 0.06, Math.max(0.03, w1), 0.22, k === 5 ? "3A8A80" : CYAN, `A2 bar ${k}`);
    T(s, `${r[2]}K (${r[3]}%)`, { x: bx + w1 + 0.06, y: yy + 0.02, w: 1.6, h: 0.3, fontSize: 10, bold: true, color: NAVY, valign: "middle", objectName: `A2 value ${k}` });
    R(s, bx, yy + 0.31, Math.max(0.03, w2), 0.18, "B8C2CC", `A2 base bar ${k}`);
    T(s, `${r[4]}K (${r[5]}%)`, { x: bx + w2 + 0.06, y: yy + 0.26, w: 1.6, h: 0.26, fontSize: 9, color: MUTED, valign: "middle", objectName: `A2 base value ${k}` });
  });
  s.addShape(pres.shapes.LINE, { x: 0.5, y: y0 + 4 * rh + 0.07, w: 9.9, h: 0, line: { color: LGREY, width: 1 }, objectName: "A2 divider" });
  RR(s, 10.55, 1.95, 2.28, 4.4, C.background2, "A2 side");
  T(s, [{ text: "84%", options: { bold: true, fontSize: 26, color: BLUE, breakLine: true } }, { text: "of KTB-inactive customers borrowed or applied, vs 73% overall", options: { fontSize: 10.5, color: TXT, breakLine: true } },
    { text: " ", options: { fontSize: 8, breakLine: true } }, { text: "50%", options: { bold: true, fontSize: 26, color: NAVY, breakLine: true } }, { text: "applied, didn't get a loan and keep no balance: the biggest group to re-engage", options: { fontSize: 10.5, color: TXT } }],
    { x: 10.7, y: 2.05, w: 2.0, h: 4.2, valign: "top", objectName: "A2 side text" });
}

// ---------- A3 Payment segments
{
  const s = mk(null, "APPENDIX A3  |  PAYMENT SEGMENTS, AUG–OCT 2026", "Win back the 114K who paid recently and stopped before chasing the 1.27M who never paid",
    "Each customer sits in one segment by the months they paid in-app (Aug, Sep, Oct). This extract counts 117K Oct payers; the partner analysis on slide 7 counts 127K from a different extract. Next steps are suggestions to test. " + SRC, "Appendix");
  const rows = [["Habitual", "Aug, Sep, Oct", 57652, "Offer the payment type they don't use yet", CYAN], ["Building", "Sep, Oct", 20220, "Reward a third month in a row", CYAN],
    ["New", "Oct only", 24502, "Cashback on a second payment within 30 days", CYAN], ["Reactivated", "Aug, Oct", 14672, "Autopay or recurring bills", CYAN],
    ["At risk", "Sep, not Oct", 52561, "Win back now: nudge before the bill date", ORANGE], ["Lapsed", "Aug only", 61323, "Low-cost win-back offer and survey", ORANGE],
    ["Never paid", "none", 1270212, "First payment via loan repayment (75K) or deposit (240K); 955K need a first deposit", GREY]];
  const y0 = 1.95, rh = 0.56, bx = 3.6, bw = 3.2;
  ["SEGMENT", "PAID IN", "CUSTOMERS", "NEXT STEP"].forEach((h, k) => T(s, h, { x: [0.5, 2.0, bx, 7.9][k], y: 1.65, w: 2, h: 0.24, fontSize: 9, bold: true, color: MUTED, charSpacing: 1, objectName: `A3 head ${k}` }));
  rows.forEach((r, k) => {
    const yy = y0 + k * rh + (k >= 4 ? 0.5 : 0) + (k >= 6 ? 0.1 : 0);
    if (k === 4 || k === 6 || k === 0) s.addShape(pres.shapes.LINE, { x: 0.5, y: yy - 0.04, w: 12.33, h: 0, line: { color: LGREY, width: 1 }, objectName: `A3 rule ${k}` });
    if (k >= 4 && k < 6) RR(s, 0.45, yy + 0.02, 12.43, rh - 0.04, "FBEFE2", `A3 risk ${k}`);
    T(s, r[0], { x: 0.55, y: yy, w: 1.45, h: rh, fontSize: 12, bold: true, color: NAVY, valign: "middle", objectName: `A3 seg ${k}` });
    T(s, r[1], { x: 2.0, y: yy, w: 1.5, h: rh, fontSize: 10.5, color: MUTED, valign: "middle", objectName: `A3 months ${k}` });
    const w = k === 6 ? bw : bw * r[2] / 61323;
    R(s, bx, yy + 0.15, w, rh - 0.3, r[4], `A3 bar ${k}`);
    T(s, r[2].toLocaleString("en-US") + (k === 6 ? "  (bar off scale)" : ""), { x: k === 6 ? bx + 0.1 : bx + w + 0.08, y: yy, w: 2.2, h: rh, fontSize: 11, bold: true, color: k === 6 ? WHITE : NAVY, valign: "middle", objectName: `A3 value ${k}` });
    T(s, r[3], { x: 7.9, y: yy, w: 4.9, h: rh, fontSize: 10.5, color: TXT, valign: "middle", objectName: `A3 next ${k}` });
  });
  T(s, [{ text: "117,046 paid in Oct", options: { bold: true, color: NAVY } }, { text: "  ·  59.7% of Sep payers paid again in Oct", options: { color: MUTED } }], { x: 2.0, y: y0 + 4 * rh + 0.05, w: 8, h: 0.4, fontSize: 11, valign: "middle", objectName: "A3 subtotal" });
}

// ---------- A4 Products by age and job
{
  const s = mk(null, "APPENDIX A4  |  PRODUCT HOLDING BY AGE AND JOB", "Age and job shape what people use: people in their 30s and 40s borrow, those over 60 save; salaried workers borrow, business owners save",
    "Share of each group's customers; a customer can hold several. By age: product segments, 31 Oct 2026 (borrow = any loan; save = balance over ฿100 or Save Max/More; pay = in-app payments). By job: Sep 2026 (save = balance over ฿0). * under 11K customers", "Appendix");
  const y = 1.62, H = 5.15, PW = (12.33 - 0.15) / 2;
  const ser = (b, sv, p) => [{ name: "Borrow", color: NAVY, vals: b }, { name: "Save", color: CYAN, vals: sv }, { name: "Pay", color: BLUE, vals: p }];
  [{ n: 1, head: "By age: teens pay, 30–49s borrow, and saving rises to 28% at 60–69", sub: "% of each age group holding each product", cats: ["10–19*", "20–29", "30–39", "40–49", "50–59", "60–69"],
    data: ser([0, 11, 12, 11, 9, 0], [11, 4, 7, 11, 17, 28], [14, 7, 8, 7, 7, 7]), max: 30, so: "Payments for the young, credit for 20–59s, savings and then investing for over-50s. No loans under 20 or over 59 (age limits)" },
  { n: 2, head: "By job: salaried borrow most, merchants save most; blue-collar and riders use least", sub: "% of each job group holding each product", cats: ["Salaried", "Merchant", "Freelance", "Rider", "Blue collar"],
    data: ser([11, 7, 6, 5, 7], [18, 20, 16, 13, 12], [14, 14, 12, 11, 10]), max: 22, so: "Cards and loans for salaried customers; savings and SME credit for merchants; savings first for blue-collar and riders" }].forEach((p, i) => {
    const x = 0.5 + i * (PW + 0.15);
    RR(s, x, y, PW, H, C.background2, `A4 panel ${i}`);
    head(s, x, y, PW, p.n, p.head, p.sub);
    legend(s, x + PW - 2.6, y + 0.6, p.data.map((d) => [d.name, d.color]));
    ccols(s, x + 0.3, y + 1.35, PW - 0.6, 2.35, p.cats, p.data, { max: p.max, fmt: (v) => v + "%", name: `A4 chart ${i}`, cfs: 10 });
    RR(s, x + 0.2, y + 4.25, PW - 0.4, 0.7, WHITE, `A4 so box ${i}`);
    T(s, [{ text: "SO  ", options: { bold: true, color: CYAN, fontSize: 9.5, charSpacing: 1 } }, { text: p.so, options: { color: TXT, fontSize: 10.5 } }], { x: x + 0.34, y: y + 4.25, w: PW - 0.68, h: 0.7, valign: "middle", objectName: `A4 so ${i}` });
  });
}

// ---------- A5 Data notes
{
  const s = mk(null, "APPENDIX A5  |  DATA NOTES", "The numbers behind this deck: five customer bases, and seven points to confirm before the board",
    SRC + "; Partner Payment Story v1; product segment extract, 31 Oct 2026", "Appendix");
  const y = 1.62, H = 5.15, LW = 5.2, RX = 5.85, RW = 12.83 - 5.85;
  RR(s, 0.5, y, LW, H, C.background2, "Bases panel");
  head(s, 0.5, y, LW, null, "Customer bases used", "Each slide's footnote says which one it uses");
  [["49.5M", "Consortium base: AIS, KTB and OR customers, Apr 2026"], ["2.14M", "CLICX sign-ups, Sep 2026: used for the customer map, loans and activity"],
    ["1.86M", "Gave partner consent (data room): used for join rates"], ["1.82M", "AIS and True split: 561,177 + 1,257,522 customers"],
    ["1.51M", "Product segment data, 31 Oct 2026: used for who is active and for payments"], ["~310K", "Active customers: 312K monthly active in Sep (15% of sign-ups); 309K in Oct with a balance over ฿100, a payment or a loan"]].forEach((r, k) => {
    const yy = y + 0.95 + k * 0.68;
    if (k) s.addShape(pres.shapes.LINE, { x: 0.7, y: yy - 0.02, w: LW - 0.4, h: 0, line: { color: LGREY, width: 0.75 }, objectName: `Base rule ${k}` });
    numRow(s, 0.7, yy, LW - 0.4, 0.64, r[0], r[1], { nw: 0.95, nfs: 15, col: NAVY, fs: 10 });
  });
  RR(s, RX, y, RW, H, C.background2, "Confirm panel");
  head(s, RX, y, RW, null, "Points to confirm in the draft data", "How this deck handled each one");
  [["Why product data covers 1.51M of 2.14M sign-ups", "Not confirmed. Both give about 310K active customers, so the deck leads with that number"],
    ["Activeness table: 80K and 48K shown as group sizes", "They are the active customers in groups of 161K and 92K. With those sizes the base adds to 2.14M and active to 312K; used here"],
    ["Join-rate table: under-24s and single-bank AIS", "The rates shown (1.5%, 0.3%) don't match their counts (3.3%, 0.7%); both left out"],
    ["October payers: 127K or 117K", "Partner file vs segment extract. Slide 7 states which point uses which; A3 uses 117K"],
    ["Share who came for a loan: 72%, 73% or 75%", "The draft says 72%; its two activeness tables give 73% (used in A2) and 75%; the deck says 3 in 4"],
    ["SME loan sizing: ฿106M or ฿1,064M", "Two different figures in the draft; neither is used"],
    ["Merchant activeness table", "Borrower percentages don't match their counts; not used"]].forEach((r, k) => {
    const yy = y + 0.95 + k * 0.585;
    badge(s, RX + 0.2, yy + 0.12, k + 1, GREY, 0.28, 9.5);
    T(s, [{ text: r[0], options: { bold: true, color: NAVY, breakLine: true } }, { text: r[1], options: { color: TXT } }], { x: RX + 0.6, y: yy, w: RW - 0.8, h: 0.56, fontSize: 9.5, valign: "middle", objectName: `Confirm ${k}` });
  });
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
