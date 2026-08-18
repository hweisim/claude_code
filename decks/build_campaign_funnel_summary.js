const pptxgen = require("pptxgenjs");

// ---------- palette ----------
const INK      = "0C2B33"; // deep teal-black
const INK_CARD = "143A45";
const TEAL     = "028090";
const SEAFOAM  = "00A896";
const MINT     = "02C39A";
const SLATE    = "7E9AA1"; // organic / unreached
const TINT     = "EEF5F5";
const BODY     = "33484F";
const MUTED    = "6E878E";
const ICE      = "A9C9CE";
const WHITE    = "FFFFFF";

const HEAD = "Cambria";
const SANS = "Calibri";

const sh = () => ({ type: "outer", color: "0C2B33", blur: 10, offset: 2, angle: 90, opacity: 0.10 });

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";           // 13.3 x 7.5
pres.author = "Analytics";
pres.title  = "Campaign Funnel Summary";

const W = 13.3, H = 7.5;

// ---------- helpers ----------
function titleBlock(slide, kicker, title, sub, dark) {
  slide.addText(kicker, {
    x: 0.6, y: 0.42, w: 9.0, h: 0.28, margin: 0,
    fontFace: SANS, fontSize: 11.5, bold: true, charSpacing: 2,
    color: dark ? SEAFOAM : TEAL,
  });
  slide.addText(title, {
    x: 0.6, y: 0.76, w: 11.6, h: 0.72, margin: 0,
    fontFace: HEAD, fontSize: 32, bold: true, color: dark ? WHITE : INK,
  });
  if (sub) {
    slide.addText(sub, {
      x: 0.6, y: 1.46, w: 11.6, h: 0.34, margin: 0,
      fontFace: SANS, fontSize: 14, color: dark ? ICE : MUTED,
    });
  }
}

function card(slide, o) {
  slide.addShape(pres.ShapeType.roundRect, {
    x: o.x, y: o.y, w: o.w, h: o.h, rectRadius: 0.10,
    fill: { color: o.fill }, line: { color: o.line || o.fill, width: 1 },
    shadow: o.noShadow ? undefined : sh(),
  });
}

// ==========================================================
// SLIDE 1 — title
// ==========================================================
{
  const s = pres.addSlide();
  s.background = { color: INK };

  s.addText("CAMPAIGN EXPOSURE ANALYSIS", {
    x: 0.8, y: 1.55, w: 6.4, h: 0.3, margin: 0,
    fontFace: SANS, fontSize: 12, bold: true, charSpacing: 2.4, color: SEAFOAM,
  });
  s.addText("Organic, View, Click", {
    x: 0.8, y: 1.98, w: 6.6, h: 1.5, margin: 0,
    fontFace: HEAD, fontSize: 38, bold: true, color: WHITE, lineSpacing: 42,
  });
  s.addText("How 13,332 customers split across the funnel — and where each segment falls out.", {
    x: 0.8, y: 3.52, w: 6.2, h: 0.9, margin: 0,
    fontFace: SANS, fontSize: 15, color: ICE, lineSpacing: 22,
  });

  s.addShape(pres.ShapeType.roundRect, {
    x: 0.8, y: 4.72, w: 3.0, h: 0.52, rectRadius: 0.26,
    fill: { color: MINT }, line: { color: MINT },
  });
  s.addText("13,332 customers", {
    x: 0.8, y: 4.72, w: 3.0, h: 0.52, margin: 0, align: "center", valign: "middle",
    fontFace: SANS, fontSize: 14, bold: true, color: INK,
  });

  const rows = [
    { pct: "16.1%", n: "2,148",  lbl: "Organic — never saw it",  c: SLATE },
    { pct: "49.8%", n: "6,637",  lbl: "Viewed, did not click",   c: TEAL  },
    { pct: "34.1%", n: "4,547",  lbl: "Viewed and clicked",      c: MINT  },
  ];
  rows.forEach((r, i) => {
    const y = 1.62 + i * 1.5;
    card(s, { x: 7.6, y, w: 4.9, h: 1.25, fill: INK_CARD, line: "1E4A56", noShadow: true });
    s.addShape(pres.ShapeType.ellipse, { x: 7.9, y: y + 0.5, w: 0.24, h: 0.24, fill: { color: r.c }, line: { color: r.c } });
    s.addText(r.pct, {
      x: 8.28, y: y + 0.18, w: 1.55, h: 0.9, margin: 0, valign: "middle",
      fontFace: HEAD, fontSize: 30, bold: true, color: WHITE,
    });
    s.addText(r.lbl, {
      x: 9.92, y: y + 0.24, w: 2.4, h: 0.34, margin: 0,
      fontFace: SANS, fontSize: 12.5, bold: true, color: WHITE,
    });
    s.addText(r.n + " customers", {
      x: 9.92, y: y + 0.62, w: 2.4, h: 0.3, margin: 0,
      fontFace: SANS, fontSize: 11.5, color: ICE,
    });
  });

  s.addText("Segments: Existing with Save Max · Existing · New", {
    x: 0.8, y: 6.55, w: 7.0, h: 0.3, margin: 0,
    fontFace: SANS, fontSize: 11, color: "6E8B93",
  });
  s.addNotes("Base of 13,332 customers, classified by view_f / click_f. Three mutually exclusive buckets: organic (no view, no click), viewed-only, and viewed-and-clicked. No customer clicked without a view, so the funnel is strictly sequential.");
}

// ==========================================================
// SLIDE 2 — the three buckets
// ==========================================================
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  titleBlock(s, "THE SPLIT", "The base divides three ways",
    "Every customer falls into exactly one bucket — no one clicked without first viewing.", false);

  const cards = [
    { n: "2,148", p: "16.1%", t: "Organic",            d: "Never exposed to the campaign — no view, no click.", c: SLATE },
    { n: "6,637", p: "49.8%", t: "View, no click",     d: "Saw it and stopped there. The largest single group.", c: TEAL  },
    { n: "4,547", p: "34.1%", t: "Click",              d: "Viewed and clicked through — the converted third.",   c: MINT  },
  ];
  cards.forEach((c, i) => {
    const x = 0.6 + i * 4.15;
    card(s, { x, y: 1.95, w: 3.8, h: 2.15, fill: TINT, line: "DDE9E9" });
    s.addShape(pres.ShapeType.ellipse, { x: x + 0.32, y: 2.25, w: 0.26, h: 0.26, fill: { color: c.c }, line: { color: c.c } });
    s.addText(c.t, {
      x: x + 0.68, y: 2.19, w: 2.8, h: 0.38, margin: 0, valign: "middle",
      fontFace: SANS, fontSize: 14, bold: true, color: INK,
    });
    s.addText([
      { text: c.n, options: { fontFace: HEAD, fontSize: 38, bold: true, color: c.c } },
      { text: "   " + c.p, options: { fontFace: SANS, fontSize: 15, bold: true, color: MUTED } },
    ], { x: x + 0.32, y: 2.66, w: 3.2, h: 0.72, margin: 0, valign: "middle" });
    s.addText(c.d, {
      x: x + 0.32, y: 3.42, w: 3.16, h: 0.6, margin: 0,
      fontFace: SANS, fontSize: 11.5, color: BODY, lineSpacing: 15,
    });
  });

  s.addChart(pres.ChartType.bar, [{
    name: "Customers",
    labels: ["Click", "View, no click", "Organic"],
    values: [4547, 6637, 2148],
  }], {
    x: 0.6, y: 4.4, w: 8.3, h: 2.5,
    barDir: "bar", barGapWidthPct: 55,
    chartColors: [MINT, TEAL, SLATE],
    showValue: true, dataLabelPosition: "outEnd", dataLabelColor: BODY,
    dataLabelFontFace: SANS, dataLabelFontSize: 12, dataLabelFontBold: true,
    dataLabelFormatCode: "#,##0",
    showLegend: false, showTitle: false,
    catAxisLabelColor: INK, catAxisLabelFontFace: SANS, catAxisLabelFontSize: 12,
    catAxisLineShow: false, catGridLine: { style: "none" },
    valAxisHidden: true, valGridLine: { style: "none" }, valAxisMaxVal: 7600,
  });

  card(s, { x: 9.2, y: 4.4, w: 3.5, h: 2.5, fill: INK, line: INK });
  s.addText("Read it as a funnel", {
    x: 9.5, y: 4.66, w: 2.9, h: 0.32, margin: 0,
    fontFace: SANS, fontSize: 12.5, bold: true, color: SEAFOAM,
  });
  s.addText([
    { text: "83.9%", options: { fontFace: HEAD, fontSize: 30, bold: true, color: WHITE, breakLine: true } },
    { text: "of the base was reached (11,184 viewed).", options: { fontFace: SANS, fontSize: 12, color: ICE } },
  ], { x: 9.5, y: 5.02, w: 2.9, h: 1.0, margin: 0 });
  s.addText([
    { text: "40.7%", options: { fontFace: HEAD, fontSize: 22, bold: true, color: MINT, breakLine: true } },
    { text: "of those reached clicked.", options: { fontFace: SANS, fontSize: 12, color: ICE } },
  ], { x: 9.5, y: 6.02, w: 2.9, h: 0.72, margin: 0 });

  s.addNotes("Organic 2,148 (16.1%); view-only 6,637 (49.8%); click 4,547 (34.1%). Reach = 11,184 of 13,332 = 83.9%. Click rate among those reached = 4,547 / 11,184 = 40.7%.");
}

// ==========================================================
// SLIDE 3 — funnel
// ==========================================================
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  titleBlock(s, "FUNNEL", "Reach is not the problem — the click is",
    "Bar width is proportional to customer count.", false);

  const stages = [
    { w: 10.0, c: INK,  label: "All customers",  n: "13,332", note: "" },
    { w: 8.39, c: TEAL, label: "Viewed",         n: "11,184", note: "83.9% of the base was reached" },
    { w: 3.41, c: MINT, label: "Clicked",        n: "4,547",  note: "40.7% of viewers clicked — the single biggest drop" },
  ];
  stages.forEach((st, i) => {
    const y = 2.1 + i * 1.35;
    s.addShape(pres.ShapeType.roundRect, {
      x: 0.9, y, w: st.w, h: 1.05, rectRadius: 0.08,
      fill: { color: st.c }, line: { color: st.c }, shadow: sh(),
    });
    s.addText([
      { text: st.n, options: { fontFace: HEAD, fontSize: 24, bold: true, color: WHITE } },
      { text: "    " + st.label, options: { fontFace: SANS, fontSize: 14, color: i === 0 ? ICE : "DCF2EE" } },
    ], { x: 1.2, y, w: st.w - 0.5, h: 1.05, margin: 0, valign: "middle" });
    if (st.note) {
      s.addText(st.note, {
        x: 0.9 + st.w + 0.28, y, w: 12.7 - (0.9 + st.w + 0.28), h: 1.05, margin: 0, valign: "middle",
        fontFace: SANS, fontSize: 12.5, color: BODY, lineSpacing: 17,
      });
    }
  });

  card(s, { x: 0.9, y: 6.15, w: 11.8, h: 0.88, fill: TINT, line: "DDE9E9", noShadow: true });
  s.addShape(pres.ShapeType.ellipse, { x: 1.2, y: 6.46, w: 0.26, h: 0.26, fill: { color: MINT }, line: { color: MINT } });
  s.addText("6,637 customers saw the campaign and did nothing — 3.1x the number who never saw it at all. The opportunity is in conversion, not in wider distribution.", {
    x: 1.6, y: 6.15, w: 10.9, h: 0.88, margin: 0, valign: "middle",
    fontFace: SANS, fontSize: 13, color: INK,
  });

  s.addNotes("Strictly sequential funnel: 13,332 -> 11,184 viewed (83.9%) -> 4,547 clicked (40.7% of viewers, 34.1% of base). The view-to-click step loses 6,637 customers, far more than the 2,148 never reached.");
}

// ==========================================================
// SLIDE 4 — by segment, counts
// ==========================================================
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  titleBlock(s, "BY SEGMENT", "Where the volume actually sits",
    "Customer counts in each bucket, by segment.", false);

  s.addChart(pres.ChartType.bar, [
    { name: "Click",          labels: ["New", "Existing", "Existing with Save Max"], values: [128, 2533, 1886] },
    { name: "View, no click", labels: ["New", "Existing", "Existing with Save Max"], values: [153, 2444, 4040] },
    { name: "Organic",        labels: ["New", "Existing", "Existing with Save Max"], values: [1159, 150, 839] },
  ], {
    x: 0.6, y: 1.9, w: 7.9, h: 4.55,
    barDir: "bar", barGrouping: "clustered", barGapWidthPct: 40,
    chartColors: [MINT, TEAL, SLATE],
    showValue: true, dataLabelPosition: "outEnd", dataLabelColor: BODY,
    dataLabelFontFace: SANS, dataLabelFontSize: 10.5, dataLabelFontBold: true,
    dataLabelFormatCode: "#,##0",
    showLegend: true, legendPos: "b", legendFontFace: SANS, legendFontSize: 11.5, legendColor: BODY,
    showTitle: false,
    catAxisLabelColor: INK, catAxisLabelFontFace: SANS, catAxisLabelFontSize: 11.5,
    catAxisLineShow: false, catGridLine: { style: "none" },
    valAxisHidden: true, valGridLine: { style: "none" }, valAxisMaxVal: 4700,
  });

  const segs = [
    { t: "Existing with Save Max", n: "6,765", d: "50.7% of the base · 1,886 clicks", c: INK },
    { t: "Existing",               n: "5,127", d: "38.5% of the base · 2,533 clicks", c: INK },
    { t: "New",                    n: "1,440", d: "10.8% of the base · 128 clicks",   c: INK },
  ];
  segs.forEach((g, i) => {
    const y = 1.95 + i * 1.55;
    card(s, { x: 8.8, y, w: 3.9, h: 1.32, fill: TINT, line: "DDE9E9" });
    s.addShape(pres.ShapeType.ellipse, { x: 9.1, y: y + 0.27, w: 0.22, h: 0.22, fill: { color: g.c }, line: { color: g.c } });
    s.addText(g.t, {
      x: 9.44, y: y + 0.16, w: 3.1, h: 0.44, margin: 0, valign: "middle",
      fontFace: SANS, fontSize: 12.5, bold: true, color: INK,
    });
    s.addText([
      { text: g.n, options: { fontFace: HEAD, fontSize: 22, bold: true, color: INK } },
    ], { x: 9.1, y: y + 0.6, w: 1.6, h: 0.42, margin: 0, valign: "middle" });
    s.addText(g.d, {
      x: 9.1, y: y + 1.0, w: 3.5, h: 0.26, margin: 0,
      fontFace: SANS, fontSize: 10.5, color: MUTED,
    });
  });

  s.addText("Existing customers deliver 2,533 of all 4,547 clicks (55.7%) from only 38.5% of the base. New customers contribute 2.8%.", {
    x: 0.6, y: 6.62, w: 12.1, h: 0.4, margin: 0,
    fontFace: SANS, fontSize: 12, color: BODY,
  });

  s.addNotes("Existing with Save Max is the biggest segment but converts least. Existing punches above its weight: 38.5% of the base, 55.7% of clicks. New is small and almost entirely unreached.");
}

// ==========================================================
// SLIDE 5 — reach vs click quality
// ==========================================================
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  titleBlock(s, "QUALITY OF EXPOSURE", "Reach and click quality pull apart",
    "Reach = share of segment that viewed. Click rate = share of viewers who clicked.", false);

  s.addChart(pres.ChartType.bar, [
    { name: "Reached (viewed)",     labels: ["Existing with\nSave Max", "Existing", "New"], values: [87.6, 97.1, 19.5] },
    { name: "Click rate of viewers", labels: ["Existing with\nSave Max", "Existing", "New"], values: [31.8, 50.9, 45.6] },
  ], {
    x: 0.6, y: 1.95, w: 6.5, h: 4.35,
    barDir: "col", barGrouping: "clustered", barGapWidthPct: 55,
    chartColors: [TEAL, MINT],
    showValue: true, dataLabelPosition: "outEnd", dataLabelColor: BODY,
    dataLabelFontFace: SANS, dataLabelFontSize: 11, dataLabelFontBold: true,
    dataLabelFormatCode: '0.0"%"',
    showLegend: true, legendPos: "b", legendFontFace: SANS, legendFontSize: 11.5, legendColor: BODY,
    showTitle: false,
    catAxisLabelColor: INK, catAxisLabelFontFace: SANS, catAxisLabelFontSize: 11,
    catAxisLineShow: false, catGridLine: { style: "none" },
    valAxisHidden: true, valGridLine: { style: "none" }, valAxisMaxVal: 118,
  });

  const notes = [
    { c: INK,  h: "Existing with Save Max — seen, ignored",
      d: "87.6% reached, but only 31.8% of those viewers clicked. The weakest conversion of the three, on the largest audience." },
    { c: INK,  h: "Existing — the engine",
      d: "97.1% reached and a 50.9% click rate. Nearly every customer saw it and half acted on it." },
    { c: INK, h: "New — barely reached",
      d: "Only 19.5% ever saw the campaign, yet 45.6% of those who did clicked. Distribution, not appeal, is the constraint." },
  ];
  notes.forEach((n, i) => {
    const y = 1.95 + i * 1.5;
    card(s, { x: 7.4, y, w: 5.3, h: 1.3, fill: TINT, line: "DDE9E9" });
    s.addShape(pres.ShapeType.ellipse, { x: 7.7, y: y + 0.28, w: 0.22, h: 0.22, fill: { color: n.c }, line: { color: n.c } });
    s.addText(n.h, {
      x: 8.04, y: y + 0.16, w: 4.4, h: 0.42, margin: 0, valign: "middle",
      fontFace: SANS, fontSize: 12.5, bold: true, color: INK,
    });
    s.addText(n.d, {
      x: 7.7, y: y + 0.58, w: 4.8, h: 0.62, margin: 0,
      fontFace: SANS, fontSize: 11, color: BODY, lineSpacing: 14,
    });
  });

  s.addText("Overall: 83.9% reached · 40.7% click rate among those reached.", {
    x: 0.6, y: 6.55, w: 12.1, h: 0.34, margin: 0,
    fontFace: SANS, fontSize: 12, bold: true, color: MUTED,
  });

  s.addNotes("Reach: Save Max 87.6%, Existing 97.1%, New 19.5%. Click rate among viewers: 31.8%, 50.9%, 45.6%. New customers convert well when reached — the gap is delivery.");
}

// ==========================================================
// SLIDE 6 — channel mix: banner vs push
// ==========================================================
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  titleBlock(s, "CHANNEL MIX", "Two channels, opposite profiles",
    "Banner carried the reach. Push notification carried the conversion.", false);

  const chans = [
    { name: "Banner", c: TEAL,
      stats: [["11,183", "viewers"], ["4,490", "clicks"], ["40.2%", "click rate"]],
      note: "Reached 83.9% of the base and delivered 98.7% of every click." },
    { name: "Push notification", c: MINT,
      stats: [["78", "viewers"], ["68", "clicks"], ["87.2%", "click rate"]],
      note: "Reached 0.6% of the base — but almost everyone who got it clicked." },
  ];
  chans.forEach((ch, i) => {
    const x = 0.6 + i * 6.2;
    card(s, { x, y: 1.9, w: 5.9, h: 2.2, fill: TINT, line: "DDE9E9" });
    s.addShape(pres.ShapeType.ellipse, { x: x + 0.32, y: 2.2, w: 0.26, h: 0.26, fill: { color: ch.c }, line: { color: ch.c } });
    s.addText(ch.name, {
      x: x + 0.68, y: 2.14, w: 4.9, h: 0.38, margin: 0, valign: "middle",
      fontFace: SANS, fontSize: 15, bold: true, color: INK,
    });
    ch.stats.forEach((st, j) => {
      const sx = x + 0.32 + j * 1.85;
      s.addText(st[0], {
        x: sx, y: 2.62, w: 1.8, h: 0.6, margin: 0, valign: "middle",
        fontFace: HEAD, fontSize: 26, bold: true, color: j === 2 ? ch.c : INK,
      });
      s.addText(st[1], {
        x: sx, y: 3.2, w: 1.8, h: 0.28, margin: 0,
        fontFace: SANS, fontSize: 11, color: MUTED,
      });
    });
    s.addText(ch.note, {
      x: x + 0.32, y: 3.56, w: 5.26, h: 0.44, margin: 0,
      fontFace: SANS, fontSize: 11.5, color: BODY,
    });
  });

  s.addChart(pres.ChartType.bar, [{
    name: "Click rate",
    labels: ["Banner", "Push notification", "All channels"],
    values: [40.2, 87.2, 40.7],
  }], {
    x: 0.6, y: 4.3, w: 5.9, h: 2.3,
    barDir: "col", barGapWidthPct: 70,
    chartColors: [TEAL, MINT, SLATE],
    showValue: true, dataLabelPosition: "outEnd", dataLabelColor: BODY,
    dataLabelFontFace: SANS, dataLabelFontSize: 12, dataLabelFontBold: true,
    dataLabelFormatCode: '0.0"%"',
    showLegend: false, showTitle: true, title: "Click rate among those reached",
    titleColor: INK, titleFontFace: SANS, titleFontSize: 12.5,
    catAxisLabelColor: INK, catAxisLabelFontFace: SANS, catAxisLabelFontSize: 11.5,
    catAxisLineShow: false, catGridLine: { style: "none" },
    valAxisHidden: true, valGridLine: { style: "none" }, valAxisMaxVal: 105,
  });

  card(s, { x: 6.8, y: 4.3, w: 5.9, h: 2.3, fill: INK, line: INK });
  s.addText("Scale is the gap, not quality", {
    x: 7.15, y: 4.56, w: 5.2, h: 0.32, margin: 0,
    fontFace: SANS, fontSize: 13, bold: true, color: SEAFOAM,
  });
  s.addText([
    { text: "143 : 1", options: { fontFace: HEAD, fontSize: 32, bold: true, color: WHITE, breakLine: true } },
    { text: "banner viewers for every push recipient — while push converts at 2.2x the banner's rate.", options: { fontFace: SANS, fontSize: 12.5, color: ICE } },
  ], { x: 7.15, y: 4.94, w: 5.2, h: 1.1, margin: 0, lineSpacing: 17 });
  s.addText("The overall 40.7% click rate is, in practice, the banner's number.", {
    x: 7.15, y: 6.06, w: 5.2, h: 0.3, margin: 0,
    fontFace: SANS, fontSize: 11.5, italic: true, color: "8FB4BA",
  });

  s.addText("77 customers saw both channels and 11 clicked both, so channel figures sum to more than the 11,184 total viewers and 4,547 total clicks.", {
    x: 0.6, y: 6.72, w: 12.1, h: 0.34, margin: 0,
    fontFace: SANS, fontSize: 10.5, color: MUTED,
  });

  s.addNotes("Banner: 11,183 viewers, 4,490 clicks, 40.2%. Push: 78 viewers, 68 clicks, 87.2%. Overlap is 77 viewers / 11 clickers, so the two channels do not sum to the totals. Push is 0.6% of the base but converts 2.2x better.");
}

// ==========================================================
// SLIDE 7 — push in detail: overlap and incrementality
// ==========================================================
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  titleBlock(s, "PUSH IN DETAIL", "Push rescued the clicks the banner lost",
    "Push went only to Existing with Save Max — 78 of that segment's 6,765 customers.", false);

  const flow = [
    { n: "6,765", l: "Save Max base",     c: INK   },
    { n: "5,925", l: "saw the banner",    c: TEAL  },
    { n: "78",    l: "also got push",     c: SEAFOAM },
    { n: "68",    l: "clicked the push",  c: MINT  },
  ];
  flow.forEach((f, i) => {
    const x = 0.6 + i * 3.15;
    s.addShape(pres.ShapeType.roundRect, {
      x, y: 1.95, w: 2.65, h: 1.15, rectRadius: 0.09,
      fill: { color: f.c }, line: { color: f.c },
    });
    s.addText(f.n, {
      x: x + 0.2, y: 2.06, w: 2.25, h: 0.52, margin: 0, valign: "middle",
      fontFace: HEAD, fontSize: 25, bold: true, color: i < 2 ? WHITE : INK,
    });
    s.addText(f.l, {
      x: x + 0.2, y: 2.56, w: 2.25, h: 0.32, margin: 0,
      fontFace: SANS, fontSize: 11.5, color: i < 2 ? ICE : "0A3B42",
    });
    if (i < 3) {
      s.addText("→", {
        x: x + 2.65, y: 1.95, w: 0.5, h: 1.15, margin: 0, align: "center", valign: "middle",
        fontFace: SANS, fontSize: 20, bold: true, color: SLATE,
      });
    }
  });

  const big = [
    { k: "89.1%", h: "of eligible recipients were rescued",
      d: "64 push recipients had already seen the banner and not clicked it. 57 of them clicked the push — clicks the banner had failed to convert." },
    { k: "2.8x", h: "the response, same customers",
      d: "Within Existing with Save Max: 30.9% click rate on banner, 87.2% on push. The audience is identical, so channel explains the gap." },
  ];
  big.forEach((b, i) => {
    const x = 0.6 + i * 6.2;
    card(s, { x, y: 3.35, w: 5.9, h: 1.9, fill: TINT, line: "DDE9E9" });
    s.addText(b.k, {
      x: x + 0.32, y: 3.55, w: 1.75, h: 0.62, margin: 0, valign: "middle",
      fontFace: HEAD, fontSize: 30, bold: true, color: TEAL,
    });
    s.addText(b.h, {
      x: x + 2.05, y: 3.55, w: 3.53, h: 0.62, margin: 0, valign: "middle",
      fontFace: SANS, fontSize: 12.5, bold: true, color: INK,
    });
    s.addText(b.d, {
      x: x + 0.32, y: 4.25, w: 5.26, h: 0.85, margin: 0,
      fontFace: SANS, fontSize: 11.5, color: BODY, lineSpacing: 15,
    });
  });

  card(s, { x: 0.6, y: 5.45, w: 12.1, h: 0.95, fill: INK, line: INK });
  s.addText([
    { text: "Banner only:  ", options: { fontFace: SANS, fontSize: 13, color: ICE } },
    { text: "40.3% clicked", options: { fontFace: SANS, fontSize: 13, bold: true, color: WHITE } },
    { text: "  (11,106 viewers)          ", options: { fontFace: SANS, fontSize: 12, color: "8FB4BA" } },
    { text: "Saw both channels:  ", options: { fontFace: SANS, fontSize: 13, color: ICE } },
    { text: "90.9% clicked", options: { fontFace: SANS, fontSize: 13, bold: true, color: MINT } },
    { text: "  (77 viewers)", options: { fontFace: SANS, fontSize: 12, color: "8FB4BA" } },
  ], { x: 1.0, y: 5.45, w: 11.3, h: 0.95, margin: 0, valign: "middle" });

  s.addText("Directional only: push reached 78 people, recipients were not randomly selected and there was no holdout, so part of the gap may be targeting rather than channel.", {
    x: 0.6, y: 6.6, w: 12.1, h: 0.4, margin: 0,
    fontFace: SANS, fontSize: 10.5, italic: true, color: MUTED,
  });

  s.addNotes("64 push recipients had seen the banner without clicking; 57 clicked the push (89.1%). Within Save Max, push converts at 87.2% vs 30.9% on banner. Dual-channel viewers click at 90.9% vs 40.3% for banner-only. Caveat: n=78, non-random selection, no holdout.");
}

// ==========================================================
// SLIDE 8 — campaign performance table
// ==========================================================
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  titleBlock(s, "CAMPAIGN PERFORMANCE", "Eight campaigns, two very different engines",
    "CLICX SaveMore, live 17 Aug 2026 — banner placements run to 30 Sep, push sends to 31 Aug.", false);

  const kpis = [
    { n: "8", l: "campaigns live", sub: "7 delivering" },
    { n: "138,637", l: "views", sub: "99.9% from banner" },
    { n: "11,278", l: "clicks", sub: "99.3% from banner" },
    { n: "11,485", l: "conversions", sub: "definition unconfirmed" },
  ];
  kpis.forEach((k, i) => {
    const x = 0.6 + i * 3.08;
    card(s, { x, y: 1.9, w: 2.86, h: 1.05, fill: TINT, line: "DDE9E9" });
    s.addText(k.n, {
      x: x + 0.26, y: 1.98, w: 2.34, h: 0.5, margin: 0, valign: "middle",
      fontFace: HEAD, fontSize: 24, bold: true, color: i === 0 ? INK : TEAL,
    });
    s.addText([
      { text: k.l, options: { fontFace: SANS, fontSize: 11.5, bold: true, color: INK } },
      { text: "   " + k.sub, options: { fontFace: SANS, fontSize: 10, color: MUTED } },
    ], { x: x + 0.26, y: 2.48, w: 2.34, h: 0.3, margin: 0 });
  });

  const hdr = ["Campaign", "Leads", "Views", "Clicks", "CTR", "Conversions"];
  const body = [
    ["Banner — Home",           "1,891,142", "127,981", "10,766", "8.4%",  "11,183"],
    ["Banner — Living Room",    "1,891,142", "10,552",  "430",    "4.1%",  "223"],
    ["Banner — Pocket Landing", "1,891,142", "0",       "0",      "—",     "0"],
    ["Push — ID89",             "11,499",    "28",      "21",     "75.0%", "21"],
    ["Push — ID23",             "11,339",    "26",      "19",     "73.1%", "21"],
    ["Push — ID01",             "11,469",    "19",      "16",     "84.2%", "12"],
    ["Push — ID45",             "11,481",    "16",      "14",     "87.5%", "15"],
    ["Push — ID67",             "11,367",    "15",      "12",     "80.0%", "10"],
  ];
  const subs = [
    ["All banner (3 placements)", "1,891,142*", "138,533", "11,196", "8.1%",  "11,406"],
    ["All push (5 sends)",        "57,155",     "104",     "82",     "78.8%", "79"],
  ];

  const rows = [];
  rows.push(hdr.map((h, j) => ({
    text: h,
    options: { bold: true, color: WHITE, fill: { color: INK }, align: j === 0 ? "left" : "right", fontSize: 11 },
  })));
  body.forEach((r, i) => {
    rows.push(r.map((cell, j) => ({
      text: cell,
      options: {
        color: cell === "0" || cell === "—" ? MUTED : BODY,
        fill: { color: i % 2 ? "F7FAFA" : WHITE },
        align: j === 0 ? "left" : "right",
        bold: j === 0,
        fontSize: 10.5,
      },
    })));
  });
  subs.forEach((r) => {
    rows.push(r.map((cell, j) => ({
      text: cell,
      options: { color: INK, fill: { color: "E3EEEE" }, align: j === 0 ? "left" : "right", bold: true, fontSize: 10.5 },
    })));
  });

  s.addTable(rows, {
    x: 0.6, y: 3.15, w: 12.1,
    colW: [3.4, 1.9, 1.8, 1.6, 1.2, 2.2],
    rowH: 0.285, valign: "middle",
    fontFace: SANS,
    border: { type: "solid", color: "DDE9E9", pt: 1 },
    margin: [0.04, 0.09, 0.04, 0.09],
  });

  s.addText("*The three banner placements share one audience pool of 1,891,142, so their leads are not additive. Pocket Landing has recorded no delivery at all. No applications, approvals or bookings have been recorded on any campaign yet.", {
    x: 0.6, y: 6.58, w: 12.1, h: 0.5, margin: 0,
    fontFace: SANS, fontSize: 10.5, color: MUTED, lineSpacing: 14,
  });

  s.addNotes("Banner Home is carrying the campaign: 127,981 of 138,637 views and 10,766 of 11,278 clicks. Living Room is a distant second and Pocket Landing has delivered nothing. The five push sends are near-identical cells of ~11.4k leads each but have produced only 104 views between them.");
}

// ==========================================================
// SLIDE 9 — lead to conversion funnel by channel
// ==========================================================
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  titleBlock(s, "LEAD TO CONVERSION", "Lead → View → Click → Conversion",
    "Same four stages, two channels that fail and succeed at opposite ends of the funnel.", false);

  const cols = [
    { x: 0.6, w: 6.0, name: "Banner", sub: "3 placements, one shared pool", c: TEAL,
      steps: [
        ["Leads assigned", "1,891,142", "shared pool"],
        ["Views",          "138,533",   "7.3% of leads"],
        ["Clicks",         "11,196",    "8.1% of views"],
        ["Conversions",    "11,406",    "101.9% of clicks"],
      ] },
    { x: 6.9, w: 5.8, name: "Push notification", sub: "5 sends, ~11.4k leads each", c: MINT,
      steps: [
        ["Leads assigned", "57,155", "5 cells combined"],
        ["Views",          "104",    "0.2% of leads"],
        ["Clicks",         "82",     "78.8% of views"],
        ["Conversions",    "79",     "96.3% of clicks"],
      ] },
  ];
  cols.forEach((col) => {
    s.addShape(pres.ShapeType.ellipse, { x: col.x, y: 1.98, w: 0.24, h: 0.24, fill: { color: col.c }, line: { color: col.c } });
    s.addText(col.name, {
      x: col.x + 0.36, y: 1.92, w: 3.0, h: 0.36, margin: 0, valign: "middle",
      fontFace: SANS, fontSize: 15, bold: true, color: INK,
    });
    s.addText(col.sub, {
      x: col.x + col.w - 2.6, y: 1.92, w: 2.6, h: 0.36, margin: 0, valign: "middle", align: "right",
      fontFace: SANS, fontSize: 11, color: MUTED,
    });
    col.steps.forEach((st, i) => {
      const y = 2.42 + i * 0.74;
      const last = i === col.steps.length - 1;
      card(s, { x: col.x, y, w: col.w, h: 0.64, fill: i === 0 ? "E3EEEE" : TINT, line: "DDE9E9", noShadow: true });
      s.addText(st[0], {
        x: col.x + 0.28, y, w: 1.85, h: 0.64, margin: 0, valign: "middle",
        fontFace: SANS, fontSize: 11.5, color: BODY,
      });
      s.addText(st[1], {
        x: col.x + 2.1, y, w: 1.85, h: 0.64, margin: 0, valign: "middle", align: "right",
        fontFace: HEAD, fontSize: 19, bold: true, color: i === 0 ? MUTED : INK,
      });
      s.addText(st[2], {
        x: col.x + 4.05, y, w: col.w - 4.33, h: 0.64, margin: 0, valign: "middle", align: "right",
        fontFace: SANS, fontSize: 11, bold: !last && i > 0, color: last ? "B4552F" : (i === 0 ? MUTED : col.c),
      });
    });
  });

  card(s, { x: 0.6, y: 5.5, w: 6.0, h: 1.15, fill: INK, line: INK });
  s.addText("Only push tracks the deep funnel", {
    x: 0.92, y: 5.66, w: 5.4, h: 0.28, margin: 0,
    fontFace: SANS, fontSize: 12.5, bold: true, color: SEAFOAM,
  });
  s.addText("82 clicks → 51 detail-page views (62.2%) → 29 through the door (56.9%). The banner placements record none of these events, so nothing below the click can be compared.", {
    x: 0.92, y: 5.96, w: 5.4, h: 0.6, margin: 0,
    fontFace: SANS, fontSize: 11, color: ICE, lineSpacing: 14,
  });

  card(s, { x: 6.9, y: 5.5, w: 5.8, h: 1.15, fill: "FBF0EA", line: "F0D9CB" });
  s.addText("Conversion needs a definition", {
    x: 7.22, y: 5.66, w: 5.2, h: 0.28, margin: 0,
    fontFace: SANS, fontSize: 12.5, bold: true, color: "B4552F",
  });
  s.addText("It exceeds clicks on Home (103.9%) and on three of five push sends, but reaches only 51.9% on Living Room. Settle what it counts before the number is reported anywhere.", {
    x: 7.22, y: 5.96, w: 5.2, h: 0.6, margin: 0,
    fontFace: SANS, fontSize: 11, color: "6B4636", lineSpacing: 14,
  });

  s.addText("Campaign-system event counts. These are a different unit from the customer-level pivot earlier in this deck, which counts unique customers — the two sets of figures should not be added together or reconciled line for line.", {
    x: 0.6, y: 6.78, w: 12.1, h: 0.42, margin: 0,
    fontFace: SANS, fontSize: 10.5, color: MUTED, lineSpacing: 14,
  });

  s.addNotes("Banner loses people at the click (8.1% of views) but reaches 7.3% of its pool. Push is the mirror image: 78.8% click rate but only 0.2% of assigned leads ever saw it. Two flags: the deep funnel is instrumented on push only, and the conversion metric is inconsistent across placements.");
}

// ==========================================================
// SLIDE 10 — takeaways
// ==========================================================
{
  const s = pres.addSlide();
  s.background = { color: INK };
  titleBlock(s, "WHAT TO DO NEXT", "Three moves the numbers point to", null, true);

  const acts = [
    { k: "01", h: "Close the reach gap on New",
      d: "1,159 of 1,440 never saw the campaign, and none received a push. They click at 45.6% when reached — the cheapest incremental volume available.",
      c: MINT },
    { k: "02", h: "Scale the push test on Save Max",
      d: "4,040 saw the banner and did nothing. Push already converts this exact segment at 87.2% versus 30.9% — but only 78 have ever received one.",
      c: SEAFOAM },
    { k: "03", h: "Protect what works for Existing",
      d: "97.1% reach and a 50.9% click rate on banner alone, with no push at all. Treat it as the benchmark and the holdout, not a target for more spend.",
      c: ICE },
  ];
  acts.forEach((a, i) => {
    const x = 0.6 + i * 4.15;
    card(s, { x, y: 2.0, w: 3.8, h: 3.5, fill: INK_CARD, line: "1E4A56", noShadow: true });
    s.addShape(pres.ShapeType.ellipse, { x: x + 0.34, y: 2.32, w: 0.62, h: 0.62, fill: { color: a.c }, line: { color: a.c } });
    s.addText(a.k, {
      x: x + 0.34, y: 2.32, w: 0.62, h: 0.62, margin: 0, align: "center", valign: "middle",
      fontFace: HEAD, fontSize: 16, bold: true, color: INK,
    });
    s.addText(a.h, {
      x: x + 0.34, y: 3.06, w: 3.14, h: 0.78, margin: 0,
      fontFace: HEAD, fontSize: 17, bold: true, color: WHITE, lineSpacing: 21,
    });
    s.addText(a.d, {
      x: x + 0.34, y: 3.88, w: 3.14, h: 1.3, margin: 0,
      fontFace: SANS, fontSize: 11.5, color: ICE, lineSpacing: 16,
    });
  });

  s.addShape(pres.ShapeType.roundRect, {
    x: 0.6, y: 5.9, w: 12.1, h: 0.82, rectRadius: 0.10,
    fill: { color: "0F343E" }, line: { color: "1E4A56" },
  });
  s.addText("Conversion is the constraint, not coverage — and the best channel has reached only 78 people.", {
    x: 1.0, y: 5.9, w: 11.3, h: 0.82, margin: 0, valign: "middle",
    fontFace: SANS, fontSize: 13, bold: true, color: WHITE,
  });

  s.addNotes("Priorities ranked by addressable volume: Save Max non-clickers (4,040) are the largest pool, New unreached (1,159) the cheapest lift, Existing the benchmark to protect.");
}

pres.writeFile({ fileName: "/tmp/claude-0/-home-user-claude-code/510dbb5b-ebd5-5640-9449-f84271416119/scratchpad/deck/campaign_funnel_summary.pptx" })
  .then(f => console.log("wrote", f));
