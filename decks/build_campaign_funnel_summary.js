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
// SLIDE 6 — takeaways
// ==========================================================
{
  const s = pres.addSlide();
  s.background = { color: INK };
  titleBlock(s, "WHAT TO DO NEXT", "Three moves the numbers point to", null, true);

  const acts = [
    { k: "01", h: "Close the reach gap on New",
      d: "1,159 of 1,440 new customers never saw the campaign. They click at 45.6% when they do — this is the cheapest incremental volume available.",
      c: MINT },
    { k: "02", h: "Fix the creative for Save Max holders",
      d: "4,040 saw it and did nothing. A 10pt lift in their 31.8% click rate is roughly 590 extra clicks — more than New can deliver in total.",
      c: SEAFOAM },
    { k: "03", h: "Protect what works for Existing",
      d: "97.1% reach and a 50.9% click rate already. Treat this segment as the benchmark and the control, not a target for further push.",
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
  s.addText("The binding constraint is conversion, not coverage: 6,637 reached non-clickers versus 2,148 never reached.", {
    x: 1.0, y: 5.9, w: 11.3, h: 0.82, margin: 0, valign: "middle",
    fontFace: SANS, fontSize: 13, bold: true, color: WHITE,
  });

  s.addNotes("Priorities ranked by addressable volume: Save Max non-clickers (4,040) are the largest pool, New unreached (1,159) the cheapest lift, Existing the benchmark to protect.");
}

pres.writeFile({ fileName: "/tmp/claude-0/-home-user-claude-code/510dbb5b-ebd5-5640-9449-f84271416119/scratchpad/deck/campaign_funnel_summary.pptx" })
  .then(f => console.log("wrote", f));
