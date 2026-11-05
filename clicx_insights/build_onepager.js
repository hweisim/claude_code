// CLICX customer insights deck — question-led storyline, data from
// dtr_strategy2027_v4.pptx and CLICX_5-Year_Business_Plan_V1-Oct_2027.pptx
const pptxgen = require("pptxgenjs");
const { applyTheme } = require(process.env.SKILL + "/scripts/apply_theme.js");

const OUT = process.argv[2] || "CLICX_Customer_Insights.pptx";

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
// Hex mirrors for hex-only chart options
const NAVY = "1F3A5F", BLUE = "2E6DB4", CYAN = "1AA39A", GREY = "8A97A6",
  LGREY = "D3D9E0", RED = "5B6B7C", LCYAN = "9FD8D3", TXT = "1C2733", MUTED = "5B6B7C", TINT = "EEF2F6";

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "CLICX 2027 one-page summary";
pres.company = "CLICX";
const C = pres.SchemeColor;

// ---------- Layouts ----------
pres.defineSlideMaster({
  title: "TITLE",
  background: { color: C.text2 },
  objects: [
    { placeholder: { options: { name: "kicker", type: "body", x: 0.8, y: 2.0, w: 11.5, h: 0.4, fontSize: 14, bold: true, color: C.accent2, charSpacing: 2 }, text: "" } },
    { placeholder: { options: { name: "title", type: "title", x: 0.8, y: 2.45, w: 11.5, h: 1.4, fontSize: 40, bold: true, color: C.background1, valign: "top" }, text: "" } },
    { placeholder: { options: { name: "sub", type: "body", x: 0.8, y: 3.95, w: 10.5, h: 1.0, fontSize: 18, color: C.accent6, valign: "top" }, text: "" } },
    { placeholder: { options: { name: "date", type: "body", x: 0.8, y: 6.3, w: 8, h: 0.4, fontSize: 12, color: C.background1 }, text: "" } },
    { image: { path: __dirname + "/logo_white.png", x: 11.13, y: 0.5, w: 1.7, h: 0.426, altText: "CLICX logo" } },
  ],
});

pres.defineSlideMaster({
  title: "QUESTION",
  background: { color: C.background1 },
  margin: [0.5, 0.5, 0.6, 0.5],
  objects: [
    { placeholder: { options: { name: "tracker", type: "body", x: 0.5, y: 0.3, w: 9, h: 0.3, fontSize: 11, bold: true, color: C.accent1, charSpacing: 1, margin: 0 }, text: "" } },
    { placeholder: { options: { name: "title", type: "title", x: 0.5, y: 0.62, w: 12.33, h: 0.9, fontSize: 24, bold: true, color: C.text2, align: "left", valign: "top", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "answer", type: "body", x: 0.5, y: 1.62, w: 12.33, h: 0.65, fontSize: 16, color: C.text1, valign: "top", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "foot", type: "body", x: 0.5, y: 6.92, w: 11.2, h: 0.3, fontSize: 10, color: "5B6B7C", valign: "top", margin: 0 }, text: "" } },
    { image: { path: __dirname + "/logo.png", x: 11.63, y: 0.2, w: 1.2, h: 0.3, altText: "CLICX logo" } },
  ],
  slideNumber: { x: 12.33, y: 6.95, w: 0.5, h: 0.3, fontSize: 10, color: "5B6B7C", align: "right" },
});

// ---------- Helpers ----------
const AX = { catAxisLabelColor: MUTED, valAxisLabelColor: MUTED, catAxisLabelFontSize: 11, valAxisLabelFontSize: 10,
  catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt", dataLabelFontFace: "+mn-lt", legendFontFace: "+mn-lt",
  titleFontFace: "+mn-lt" };
const chartTitle = (t) => ({ showTitle: true, title: t, titleFontSize: 13, titleColor: NAVY, titleBold: true, titleAlign: "left" });
const quiet = () => ({ valGridLine: { style: "none" }, catGridLine: { style: "none" }, valAxisHidden: true,
  catAxisLineShow: true, catAxisLineColor: LGREY });

function qSlide(section, tracker, question, answer, foot) {
  const s = pres.addSlide({ masterName: "QUESTION", sectionTitle: section });
  s.addText(tracker, { placeholder: "tracker" });
  s.addText(question, { placeholder: "title" });
  s.addText(answer, { placeholder: "answer" });
  s.addText(foot, { placeholder: "foot" });
  return s;
}

// Insight panel: tinted box with numbered points
function insights(s, items, x = 9.05, y = 2.35, w = 3.78, h = 4.35, head = "So what") {
  s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: C.background2 }, line: { type: "none" }, objectName: "Insight panel" });
  s.addText(head.toUpperCase(), { x: x + 0.25, y: y + 0.18, w: w - 0.5, h: 0.3, fontSize: 11, bold: true, color: C.accent1, charSpacing: 1, margin: 0, isTextBox: true, objectName: "Insight header" });
  const step = (h - 0.65) / items.length;
  items.forEach((it, i) => {
    const yy = y + 0.6 + i * step;
    s.addShape(pres.shapes.OVAL, { x: x + 0.25, y: yy + 0.02, w: 0.32, h: 0.32, fill: { color: C.text2 }, line: { type: "none" }, objectName: `Badge ${i + 1}` });
    s.addText(String(i + 1), { x: x + 0.25, y: yy + 0.02, w: 0.32, h: 0.32, fontSize: 11, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: `Badge num ${i + 1}` });
    s.addText([
      { text: it[0], options: { bold: true, fontSize: 14, color: C.text2, breakLine: true } },
      { text: it[1], options: { fontSize: 12, color: C.text1 } },
    ], { x: x + 0.7, y: yy - 0.03, w: w - 0.9, h: step - 0.08, valign: "top", margin: 0, paraSpaceAfter: 3, isTextBox: true, objectName: `Insight ${i + 1}` });
  });
}

// Big stat callout
function stat(s, x, y, w, big, label, color = C.accent1) {
  s.addText(big, { x, y, w, h: 0.7, fontSize: 40, bold: true, color, margin: 0, isTextBox: true, objectName: `Stat ${big}` });
  s.addText(label, { x, y: y + 0.72, w, h: 0.55, fontSize: 12, color: C.text1, margin: 0, valign: "top", isTextBox: true, objectName: `Stat label ${big}` });
}

// =====================================================================
// Two-slide summary: customer landscape, then strategy to product
const SRC = "Source: Data for Strategy 2027 (Nov 2026); CLICX payment segments (1 Oct 2026); CLICX data as of Sep 2026; AIS, KTB, OR data as of Apr 2026";
pres.addSection({ title: "Summary" });
{
  const s = pres.addSlide({ masterName: "QUESTION", sectionTitle: "Summary" });
  s.addText("1 / 2  |  CURRENT CUSTOMER LANDSCAPE", { placeholder: "tracker" });
  s.addText("What does our current customer landscape look like?", { placeholder: "title" });
  s.addText("Source: Data Room; Data for Strategy 2027 (Nov 2026). CLICX data as of Sep 2026; AIS, KTB, OR data as of Apr 2026", { placeholder: "foot" });
  const STAT_BG = "E5F1F0";
  const qTitle = (n, t, x, y, w) => s.addText([{ text: "Q" + n + "  ", options: { color: CYAN, bold: true } }, { text: t, options: { color: NAVY, bold: true } }],
    { x, y, w, h: 0.32, fontSize: 13.5, margin: 0, valign: "middle", isTextBox: true, objectName: "Q" + n + " title" });
  const statBox = (x, y, w, h, big, label, n) => {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.12, fill: { color: STAT_BG }, line: { type: "none" }, objectName: "Q" + n + " stat box" });
    s.addText(big, { x: x + 0.18, y: y + 0.08, w: w - 0.3, h: 0.62, fontSize: 34, bold: true, color: BLUE, margin: 0, valign: "middle", isTextBox: true, objectName: "Q" + n + " stat" });
    s.addText(label, { x: x + 0.18, y: y + 0.7, w: w - 0.3, h: h - 0.78, fontSize: 11, color: TXT, margin: 0, valign: "top", isTextBox: true, objectName: "Q" + n + " stat label" });
  };
  // horizontal bar rows: [label, fraction 0-1, value text, color]
  const bars = (x, y, w, rows, n, opt = {}) => {
    const lw = opt.lw || 1.45, vw = opt.vw || 0.95, bw = w - lw - vw - 0.1, rh = opt.rh || 0.3;
    rows.forEach((r, k) => {
      const yy = y + k * rh;
      s.addText(r[0], { x, y: yy, w: lw, h: rh, fontSize: 10.5, color: TXT, margin: 0, valign: "middle", isTextBox: true, objectName: `Q${n} bar label ${k}` });
      s.addShape(pres.shapes.RECTANGLE, { x: x + lw, y: yy + rh * 0.3, w: bw, h: rh * 0.4, fill: { color: "E3E8EE" }, line: { type: "none" }, objectName: `Q${n} track ${k}` });
      s.addShape(pres.shapes.RECTANGLE, { x: x + lw, y: yy + rh * 0.3, w: Math.max(0.03, bw * r[1]), h: rh * 0.4, fill: { color: r[3] }, line: { type: "none" }, objectName: `Q${n} bar ${k}` });
      s.addText(r[2], { x: x + lw + bw + 0.1, y: yy, w: vw, h: rh, fontSize: 10.5, bold: true, color: r[3] === GREY ? MUTED : NAVY, margin: 0, valign: "middle", isTextBox: true, objectName: `Q${n} bar value ${k}` });
    });
  };
  const sub = (t, x, y, w, n) => s.addText(t, { x, y, w, h: 0.24, fontSize: 10, bold: true, color: MUTED, margin: 0, isTextBox: true, objectName: `Q${n} subhead` });

  // Q1 full width
  qTitle(1, "Do all current CLICX customers come from captive consortium customers?", 0.5, 1.42, 12.3);
  statBox(0.5, 1.8, 3.4, 1.05, "96%", "of customers come from AIS, KTB or OR", 1);
  const p1 = [["Krungthai (KTB)", 0.88, "88%", "1.6M", BLUE, "logo_ktb.jpeg", 2.58], ["AIS", 0.67, "67%", "1.2M", CYAN, "logo_ais.png", 2.74], ["OR", 0.40, "40%", "725K", NAVY, "logo_or.png", 1.93], ["Non-consortium", 0.04, "4%", "51K", GREY]];
  p1.forEach((p, k) => {
    const x = 4.15 + k * 2.18, w = 1.98;
    if (p[5]) s.addImage({ path: __dirname + "/icons/" + p[5], x, y: 1.83, h: 0.28, w: 0.28 * p[6], altText: p[0] + " logo", objectName: `Q1 logo ${k}` });
    else s.addText(p[0], { x, y: 1.85, w, h: 0.26, fontSize: 11, color: MUTED, margin: 0, isTextBox: true, objectName: `Q1 partner ${k}` });
    s.addText([{ text: p[2], options: { fontSize: 22, bold: true, color: p[4] === GREY ? MUTED : p[4] } }, { text: "  " + p[3], options: { fontSize: 12, color: TXT } }],
      { x, y: 2.12, w, h: 0.4, margin: 0, valign: "middle", isTextBox: true, objectName: `Q1 partner value ${k}` });
    s.addShape(pres.shapes.RECTANGLE, { x, y: 2.62, w, h: 0.1, fill: { color: "E3E8EE" }, line: { type: "none" }, objectName: `Q1 track ${k}` });
    s.addShape(pres.shapes.RECTANGLE, { x, y: 2.62, w: Math.max(0.04, w * p[1]), h: 0.1, fill: { color: p[4] }, line: { type: "none" }, objectName: `Q1 bar ${k}` });
  });

  // Row 2: Q2 and Q3
  const L = 0.5, R = 6.83, PW = 6.0, r2 = 3.05, r3 = 5.0;
  qTitle(2, "Are the personas who come to CLICX the ones we planned?", L, r2, PW);
  statBox(L, r2 + 0.4, 2.2, 1.4, "82%", "of loan applicants are underserved, our planned target", 2);
  sub("% of each group who joined CLICX", L + 2.4, r2 + 0.4, 3.6, 2);
  bars(L + 2.4, r2 + 0.65, 3.6, [["Tech-savvy", 0.085 / 0.1, "8.5%", CYAN], ["Gen Z, 23–29", 0.066 / 0.1, "6.6%", CYAN], ["Average", 0.038 / 0.1, "3.8%", GREY], ["Multi-bank", 0.024 / 0.1, "2.4%", NAVY]], 2, { lw: 1.15, vw: 0.5 });

  qTitle(3, "Do different personas come for the same or different products?", R, r2, PW);
  statBox(R, r2 + 0.4, 2.2, 1.4, "84–92%", "of every job type applies for a loan: the same entry point", 3);
  s.addText([{ text: "■ ", options: { color: NAVY } }, { text: "got a loan   ", options: { color: MUTED } }, { text: "■ ", options: { color: CYAN } }, { text: "keeps money in", options: { color: MUTED } }],
    { x: R + 2.4, y: r2 + 0.4, w: 3.6, h: 0.24, fontSize: 10, bold: true, margin: 0, isTextBox: true, objectName: "Q3 legend" });
  [["Office worker", 0.11, 0.18], ["Merchant", 0.07, 0.20], ["Blue-collar", 0.07, 0.12], ["Rider / driver", 0.05, 0.13]].forEach((r, k) => {
    const y = r3 - 1.3 + k * 0.3, x = R + 2.4, lw = 1.1, bw = 1.6;
    s.addText(r[0], { x, y, w: lw, h: 0.3, fontSize: 10.5, color: TXT, margin: 0, valign: "middle", isTextBox: true, objectName: `Q3 label ${k}` });
    s.addShape(pres.shapes.RECTANGLE, { x: x + lw, y: y + 0.05, w: bw * r[1] / 0.22, h: 0.09, fill: { color: NAVY }, line: { type: "none" }, objectName: `Q3 loan ${k}` });
    s.addShape(pres.shapes.RECTANGLE, { x: x + lw, y: y + 0.16, w: bw * r[2] / 0.22, h: 0.09, fill: { color: CYAN }, line: { type: "none" }, objectName: `Q3 save ${k}` });
    s.addText(Math.round(r[1] * 100) + "% · " + Math.round(r[2] * 100) + "%", { x: x + lw + bw + 0.08, y, w: 0.82, h: 0.3, fontSize: 10, bold: true, color: NAVY, margin: 0, valign: "middle", isTextBox: true, objectName: `Q3 value ${k}` });
  });

  // Row 3: Q4 and Q5
  qTitle(4, "Is the persona different for active vs inactive customers?", L, r3, PW);
  statBox(L, r3 + 0.4, 2.2, 1.4, "88%", "of borrowers are active each month. Age and job barely change it", 4);
  sub("% active each month, by what they hold", L + 2.4, r3 + 0.4, 3.6, 4);
  bars(L + 2.4, r3 + 0.65, 3.6, [["Borrowers", 0.88, "88%", CYAN], ["Money in, no loan", 0.51, "51%", BLUE], ["No money in", 0.04, "4%", GREY]], 4, { lw: 1.35, vw: 0.5 });
  s.addText("1.8M idle: 694K with no product, 1.1M not using theirs", { x: L + 2.4, y: r3 + 1.55, w: 3.6, h: 0.3, fontSize: 9.5, color: MUTED, margin: 0, isTextBox: true, objectName: "Q4 note" });

  qTitle(5, "Does CLICX cannibalize KTB customers or products?", R, r3, PW);
  statBox(R, r3 + 0.4, 2.2, 1.4, "64%", "of the 1.6M customers we share with KTB were inactive with KTB (1.1M)", 5);
  sub("CLICX borrowers and KTB's Jai-pump loan", R + 2.4, r3 + 0.4, 3.6, 5);
  bars(R + 2.4, r3 + 0.65, 3.6, [["All CLICX borrowers", 1, "164K", GREY], ["No Jai-pump loan ever", 0.81, "133K (81%)", BLUE], ["Rejected by Jai-pump", 0.14, "23K (14%)", NAVY]], 5, { lw: 1.55, vw: 0.85 });
  s.addText("450K KTB rejects applied to us; 51K got our loan", { x: R + 2.4, y: r3 + 1.55, w: 3.6, h: 0.3, fontSize: 9.5, color: MUTED, margin: 0, isTextBox: true, objectName: "Q5 note" });
}
{
  const s = pres.addSlide({ masterName: "QUESTION", sectionTitle: "Summary" });
  s.addText("2 / 2  |  STRATEGY TO PRODUCT", { placeholder: "tracker" });
  s.addText("What should we build in 2027 based on these answers?", { placeholder: "title" });
  s.addText("Customer pools from Data for Strategy 2027 and the Data Room. Revenue and conversion targets are still to be set. Icons: Font Awesome (open licence)", { placeholder: "foot" });
  const cards = [
    { icon: "partners", from: "Q1", title: "Grow beyond KTB through AIS and OR", why: "96% come from partners, but only 3.7% of AIS's base has joined",
      product: "AIS bill pay and top-up in app; partner rewards", big: "0.66M", pool: "not AIS customers yet", logos: [["logo_ais.png", 2.74], ["logo_or.png", 1.93]] },
    { icon: "lend", from: "Q2 · Q3", title: "Lend smarter to the underserved", why: "82% of applicants are underserved; 90%+ of gig workers apply, 5–7% approved",
      product: "Proxy-income loans; deposit-backed starter loan", big: "1.45M", pool: "underserved loan applicants" },
    { icon: "sme", from: "Q3", title: "Turn merchants into SME borrowers", why: "Merchants save most (20%) and spend; 84% already applied for a loan",
      product: "SME cash-flow loans", big: "293K", pool: "merchants; 245K already applied" },
    { icon: "card", from: "Q3", title: "Give office workers a first card", why: "Office workers borrow most (11%) and hold 3.2B THB; 239K are AIS Serenade members",
      product: "Digital credit card with AIS and PTT perks", big: "700K", pool: "office workers", logos: [["logo_ais.png", 2.74], ["logo_or.png", 1.93]] },
    { icon: "wealth", from: "Q3 · Q4", title: "Give savers a wealth home", why: "Working seniors and seniors hold most of our 5.5B THB deposits",
      product: "Savings pockets, then mutual funds", big: "643K", pool: "working seniors and seniors" },
    { icon: "activate", from: "Q4 · Q5", title: "Activate idle and KTB-rejected customers", why: "88% of borrowers are active vs 4% with no money in; 450K KTB rejects applied",
      product: "First-deposit bonus; Paotang reject referral", big: "1.8M", pool: "idle customers; 1.9M KTB rejects a year", logos: [["logo_ktb.jpeg", 2.58]] },
  ];
  const cw = 3.99, ch = 2.6, gx = 0.18, gy = 0.15, x0 = 0.5, y0 = 1.45;
  cards.forEach((c, i) => {
    const x = x0 + (i % 3) * (cw + gx), y = y0 + Math.floor(i / 3) * (ch + gy), n = i + 1;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: cw, h: ch, rectRadius: 0.1, fill: { color: C.background2 }, line: { type: "none" }, objectName: `Card ${n}` });
    s.addShape(pres.shapes.OVAL, { x: x + 0.2, y: y + 0.2, w: 0.62, h: 0.62, fill: { color: i % 2 ? BLUE : CYAN }, line: { type: "none" }, objectName: `Icon circle ${n}` });
    s.addImage({ path: __dirname + "/icons/" + c.icon + ".png", x: x + 0.34, y: y + 0.34, w: 0.34, h: 0.34, altText: c.title + " icon", objectName: `Icon ${n}` });
    s.addText([{ text: String(n) + "  ", options: { color: CYAN, bold: true } }, { text: c.title, options: { color: NAVY, bold: true } }],
      { x: x + 0.95, y: y + 0.18, w: cw - 1.75, h: 0.66, fontSize: 13, valign: "middle", margin: 0, isTextBox: true, objectName: `Title ${n}` });
    s.addText(c.from, { x: x + cw - 0.8, y: y + 0.2, w: 0.65, h: 0.26, fontSize: 9, bold: true, color: BLUE, align: "right", margin: 0, isTextBox: true, objectName: `From ${n}` });
    s.addText(c.why, { x: x + 0.2, y: y + 0.92, w: cw - 0.4, h: 0.42, fontSize: 10, color: MUTED, valign: "top", margin: 0, isTextBox: true, objectName: `Why ${n}` });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + 0.2, y: y + 1.4, w: cw - 0.4, h: 0.48, rectRadius: 0.06, fill: { color: "FFFFFF" }, line: { type: "none" }, objectName: `Product box ${n}` });
    s.addText([{ text: "2027 PRODUCT  ", options: { fontSize: 8.5, bold: true, color: MUTED, charSpacing: 1 } }, { text: c.product, options: { fontSize: 11, bold: true, color: CYAN } }],
      { x: x + 0.32, y: y + 1.4, w: cw - 0.64, h: 0.48, valign: "middle", margin: 0, isTextBox: true, objectName: `Product ${n}` });
    s.addText(c.big, { x: x + 0.2, y: y + 1.95, w: 1.2, h: 0.55, fontSize: 26, bold: true, color: BLUE, valign: "middle", margin: 0, isTextBox: true, objectName: `Big ${n}` });
    const lw = c.logos ? c.logos.reduce((a, [f, ar]) => a + 0.24 * ar + 0.08, 0) + 0.1 : 0;
    s.addText(c.pool, { x: x + 1.42, y: y + 1.95, w: cw - 1.62 - lw, h: 0.55, fontSize: 10.5, color: TXT, valign: "middle", margin: 0, isTextBox: true, objectName: `Pool ${n}` });
    if (c.logos) {
      let lx = x + cw - 0.2;
      c.logos.slice().reverse().forEach(([f, ar], k) => {
        const h = 0.24, w = h * ar; lx -= w;
        s.addImage({ path: __dirname + "/icons/" + f, x: lx, y: y + 2.1, w, h, altText: f.replace("logo_", "").split(".")[0].toUpperCase() + " logo", objectName: `Logo ${n}-${k}` });
        lx -= 0.08;
      });
    }
  });
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
