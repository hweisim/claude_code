/**
 * CLICX Camp Campaign Performance - Key Insights deck.
 *
 * Source: CLICX Camp Campaign Performance Dashboard (campaign-level),
 * data as of 2 Aug 2026, campaign start window 17 Jun 26 - 1 Aug 26.
 *
 *   node build_deck.js
 */

const path = require("path");
const PptxGenJS = require(path.join(
  "/tmp/claude-0/-home-user-claude-code/a5758311-6270-5365-a679-5836c8e934f0/scratchpad",
  "node_modules",
  "pptxgenjs"
));

/* ---------------------------------------------------------------- palette */

const NAVY = "0E1E3C";
const NAVY_MID = "1B3159";
const INK = "16233D";
const TEAL = "00A896";
const MINT = "5FD6BE";
const AMBER = "E9A03B";
const CORAL = "D9503F";
const PAPER = "FFFFFF";
const TINT = "F1F4F9";
const TINT_WARM = "FBF0E6";
const TINT_COOL = "E8F6F3";
const MUTED = "6B7A90";
const MUTED_LT = "A8B4C6";
const RULE = "D8DFE9";

const H = "Cambria"; // headings + numerals
const B = "Calibri"; // body

/* ------------------------------------------------------------ dashboard data */

// Volumes in thousands, exactly as printed on the dashboard.
const F = {
  lead: 9174.3,
  view: 1601.9,
  click: 127.0,
  dpv: 31.7,
  ttd1: 0.3,
  ttd2: 0.0,
  conv: 38.7,
  appin: 38.7,
  approve: 13.0,
  book: 9.7,
};
const RATE = {
  view: 17.5,
  click: 7.9,
  dpv: 25.0,
  ttd1: 0.2,
  ttd2: 0.0,
  conv: 30.5,
  appin: 30.5,
  approve: 33.6,
  book: 75.0,
};
const AVG_APPROVE = 11.3; // K per approved customer
const AVG_SETCREDIT = 10.8; // K per booked customer

const booksPerClickPct = (F.book / F.click) * 100; // 7.64%

/* ------------------------------------------------------------------ helpers */

const shadow = (o = {}) =>
  Object.assign(
    { type: "outer", angle: 90, blur: 10, offset: 2, color: "8C99AD", opacity: 0.22 },
    o
  );

function card(slide, x, y, w, h, fill, opts = {}) {
  slide.addShape("roundRect", {
    x,
    y,
    w,
    h,
    rectRadius: opts.radius || 0.09,
    fill: { color: fill },
    line: opts.line || { type: "none" },
    shadow: opts.shadow === false ? undefined : shadow(opts.shadowOpts || {}),
  });
}

function badge(slide, x, y, d, fill, label, labelColor) {
  slide.addShape("ellipse", { x, y, w: d, h: d, fill: { color: fill }, line: { type: "none" } });
  slide.addText(label, {
    x,
    y,
    w: d,
    h: d,
    align: "center",
    valign: "middle",
    margin: 0,
    fontFace: H,
    fontSize: 15,
    bold: true,
    color: labelColor || PAPER,
  });
}

function title(slide, text, kicker, dark) {
  if (kicker) {
    slide.addText(kicker.toUpperCase(), {
      x: 0.6,
      y: 0.36,
      w: 12.1,
      h: 0.26,
      margin: 0,
      fontFace: B,
      fontSize: 11.5,
      bold: true,
      charSpacing: 2.2,
      color: dark ? MINT : TEAL,
    });
  }
  slide.addText(text, {
    x: 0.6,
    y: kicker ? 0.66 : 0.5,
    w: 12.1,
    h: 0.72,
    margin: 0,
    valign: "top",
    fontFace: H,
    fontSize: 32,
    bold: true,
    color: dark ? PAPER : INK,
  });
}

function footnote(slide, text, dark) {
  slide.addText(text, {
    x: 0.6,
    y: 6.94,
    w: 12.1,
    h: 0.32,
    margin: 0,
    fontFace: B,
    fontSize: 9.5,
    italic: true,
    color: dark ? MUTED_LT : MUTED,
  });
}

/* -------------------------------------------------------------------- deck */

const pres = new PptxGenJS();
pres.layout = "LAYOUT_WIDE"; // 13.3 x 7.5
pres.author = "Campaign Analytics";
pres.title = "CLICX Camp - Campaign Performance Key Insights";

/* === 1. Title ============================================================ */
{
  const s = pres.addSlide();
  s.background = { color: NAVY };

  // motif: oversized translucent funnel rings, top-right
  [[3.1, TEAL], [2.35, TEAL], [1.6, MINT]].forEach(([d, c], i) => {
    s.addShape("ellipse", {
      x: 10.05 + (3.1 - d) / 2,
      y: 0.75 + (3.1 - d) / 2,
      w: d,
      h: d,
      fill: { color: c, transparency: 88 - i * 4 },
      line: { type: "none" },
    });
  });

  s.addText("CLICX CAMP  ·  CAMPAIGN-LEVEL READ-OUT", {
    x: 0.75,
    y: 1.28,
    w: 9.0,
    h: 0.3,
    margin: 0,
    fontFace: B,
    fontSize: 12,
    bold: true,
    charSpacing: 2.6,
    color: MINT,
  });
  s.addText("Nine million leads,\nnine thousand bookings", {
    x: 0.75,
    y: 1.72,
    w: 9.1,
    h: 1.85,
    margin: 0,
    lineSpacing: 47,
    fontFace: H,
    fontSize: 42,
    bold: true,
    color: PAPER,
  });
  s.addText(
    "Where the campaign funnel creates value — and where it loses 98.7% of it.",
    {
      x: 0.75,
      y: 3.66,
      w: 8.6,
      h: 0.45,
      margin: 0,
      fontFace: B,
      fontSize: 16,
      color: "C3D0E2",
    }
  );

  const chips = [
    ["52", "campaigns"],
    ["46", "days observed"],
    ["9.17M", "assigned leads"],
    ["9.7K", "bookings"],
    ["0.11%", "lead → book"],
  ];
  chips.forEach(([v, l], i) => {
    const x = 0.75 + i * 2.42;
    card(s, x, 4.62, 2.16, 1.24, "17294B", { shadow: false });
    s.addText(v, {
      x,
      y: 4.76,
      w: 2.16,
      h: 0.55,
      margin: 0,
      align: "center",
      fontFace: H,
      fontSize: 25,
      bold: true,
      color: i === 4 ? AMBER : MINT,
    });
    s.addText(l, {
      x,
      y: 5.34,
      w: 2.16,
      h: 0.3,
      margin: 0,
      align: "center",
      fontFace: B,
      fontSize: 11,
      color: MUTED_LT,
    });
  });

  s.addText(
    "Campaign start window 17 Jun 2026 – 1 Aug 2026   |   Data as of 2 Aug 2026 (refreshed daily, T-1)",
    {
      x: 0.75,
      y: 6.42,
      w: 11.8,
      h: 0.3,
      margin: 0,
      fontFace: B,
      fontSize: 11.5,
      color: MUTED_LT,
    }
  );
  s.addNotes(
    "Read-out of the CLICX Camp campaign performance dashboard at campaign level. " +
      "52 campaigns launched across a 46-day window. The headline: acquisition volume is abundant, " +
      "conversion of that volume is not."
  );
}

/* === 2. Executive summary ================================================ */
{
  const s = pres.addSlide();
  s.background = { color: PAPER };
  title(s, "Five things the dashboard is telling us", "Executive summary");

  const rows = [
    [
      "Volume is not the constraint",
      "9.17M assigned leads produced 9.7K bookings — an end-to-end rate of 0.11%. There is no shortage of reach.",
      TEAL,
    ],
    [
      "98.7% of all loss sits at the top",
      "Lead→View and View→Click alone shed 9.05M of the 9.16M people the funnel loses in total.",
      CORAL,
    ],
    [
      "The credit funnel is healthy",
      "After a click: 30.5% apply, 33.6% of those are approved, 75.0% of approvals book. 7.6% of clicks become a booking.",
      TEAL,
    ],
    [
      "~0.1B of credit set, 9.7K customers",
      "Approved amount averages 11.3K per customer and set credit 10.8K — a ~4% step-down between approval and booking.",
      TEAL,
    ],
    [
      "Read the rates with care",
      "Conversions (38.7K) exceed detail-page views (31.7K) and the app-in amount is empty — Loose Criteria plus capture gaps.",
      AMBER,
    ],
  ];

  rows.forEach(([head, body, accent], i) => {
    const y = 1.62 + i * 1.06;
    card(s, 0.6, y, 12.1, 0.92, TINT, { shadow: false });
    badge(s, 0.92, y + 0.19, 0.54, accent, String(i + 1));
    s.addText(head, {
      x: 1.66,
      y,
      w: 3.7,
      h: 0.92,
      margin: 0,
      valign: "middle",
      fontFace: H,
      fontSize: 15,
      bold: true,
      color: INK,
    });
    s.addText(body, {
      x: 5.5,
      y,
      w: 6.95,
      h: 0.92,
      margin: 0,
      valign: "middle",
      fontFace: B,
      fontSize: 12.5,
      color: "44536B",
    });
  });

  footnote(
    s,
    "All figures as printed on the dashboard; derived rates computed from those printed values and subject to their rounding."
  );
  s.addNotes(
    "Order matters: abundance at the top, catastrophic top-funnel loss, a healthy bottom funnel, real value delivered, " +
      "and a data-quality caveat that qualifies everything above it."
  );
}

/* === 3. The funnel end to end ============================================ */
{
  const s = pres.addSlide();
  s.background = { color: PAPER };
  title(s, "72 leads buy one click. 946 buy a booking.", "The funnel, end to end");

  const stages = [
    ["Assigned Lead", F.lead, "9,174.3K", "baseline", "0E2A4E"],
    ["View", F.view, "1,601.9K", "17.5% of leads", "174569"],
    ["Click", F.click, "127.0K", "7.9% of views", "1E6180"],
    ["App-in / Conversion", F.appin, "38.7K", "30.5% of clicks", "2A8291"],
    ["Approve", F.approve, "13.0K", "33.6% of app-ins", "3EA79B"],
    ["Book", F.book, "9.7K", "75.0% of approvals", "5FD6BE"],
  ];

  const lgMin = Math.log10(F.book);
  const lgMax = Math.log10(F.lead);
  const wMin = 1.5;
  const wMax = 7.0;

  stages.forEach(([name, vol, volLabel, rate, color], i) => {
    const y = 1.5 + i * 0.63;
    const w = wMin + ((Math.log10(vol) - lgMin) / (lgMax - lgMin)) * (wMax - wMin);

    s.addText(name, {
      x: 0.6,
      y,
      w: 2.25,
      h: 0.5,
      margin: 0,
      align: "right",
      valign: "middle",
      fontFace: B,
      fontSize: 12.5,
      bold: true,
      color: INK,
    });
    s.addShape("roundRect", {
      x: 3.0,
      y,
      w,
      h: 0.5,
      rectRadius: 0.06,
      fill: { color },
      line: { type: "none" },
    });
    s.addText(volLabel, {
      x: 3.0,
      y,
      w: w - 0.18,
      h: 0.5,
      margin: 0,
      align: "right",
      valign: "middle",
      fontFace: H,
      fontSize: 14,
      bold: true,
      color: i >= 4 ? INK : PAPER,
    });
    s.addText(rate, {
      x: 10.25,
      y,
      w: 2.45,
      h: 0.5,
      margin: 0,
      valign: "middle",
      fontFace: B,
      fontSize: 12,
      color: i === 0 ? MUTED : "44536B",
    });
  });

  s.addText("Bar length is log-scaled — a linear scale would render every stage below View invisible.", {
    x: 3.0,
    y: 5.36,
    w: 9.7,
    h: 0.28,
    margin: 0,
    fontFace: B,
    fontSize: 10,
    italic: true,
    color: MUTED,
  });

  const stats = [
    ["1 in 72", "leads ever clicks", CORAL],
    ["1 in 946", "leads ever books", CORAL],
    ["1 in 13", "clicks ends in a booking", TEAL],
  ];
  stats.forEach(([v, l, c], i) => {
    const x = 0.6 + i * 4.07;
    card(s, x, 5.72, 3.86, 1.1, TINT, { shadow: false });
    s.addText(v, {
      x: x + 0.25,
      y: 5.85,
      w: 3.4,
      h: 0.5,
      margin: 0,
      fontFace: H,
      fontSize: 24,
      bold: true,
      color: c,
    });
    s.addText(l, {
      x: x + 0.25,
      y: 6.36,
      w: 3.4,
      h: 0.32,
      margin: 0,
      fontFace: B,
      fontSize: 12,
      color: "44536B",
    });
  });

  footnote(
    s,
    "Detail Page View 31.7K (25.0% of clicks) and Thru-The-Door 1/2 at 0.3K / 0.0K are omitted here — see the appendix."
  );
  s.addNotes(
    "72 = 9,174.3 / 127.0 leads per click. 946 = 9,174.3 / 9.7 leads per booking. 13 = 127.0 / 9.7 clicks per booking."
  );
}

/* === 4. Where the volume leaks =========================================== */
{
  const s = pres.addSlide();
  s.background = { color: PAPER };
  title(s, "Two stages destroy 9.05M of the 9.16M lost", "Where the volume leaks");

  s.addChart(
    pres.charts.BAR,
    [
      {
        name: "People lost (K)",
        labels: [
          "Lead → View",
          "View → Click",
          "Click → App-in",
          "App-in → Approve",
          "Approve → Book",
        ],
        values: [7572.4, 1474.9, 88.3, 25.7, 3.3],
      },
    ],
    {
      x: 0.5,
      y: 1.52,
      w: 7.75,
      h: 4.6,
      barDir: "bar",
      barGapWidthPct: 45,
      chartColors: [CORAL, CORAL, NAVY_MID, NAVY_MID, NAVY_MID],
      varyColors: true,
      showTitle: true,
      title: "People lost at each hand-off (thousands)",
      titleFontFace: B,
      titleFontSize: 13,
      titleColor: INK,
      showValue: true,
      dataLabelPosition: "outEnd",
      dataLabelFontFace: B,
      dataLabelFontSize: 11,
      dataLabelColor: INK,
      dataLabelFormatCode: "#,##0.0",
      showLegend: false,
      catAxisLabelFontFace: B,
      catAxisLabelFontSize: 11.5,
      catAxisLabelColor: INK,
      valAxisLabelFontFace: B,
      valAxisLabelFontSize: 10,
      valAxisLabelColor: MUTED,
      valAxisMinVal: 0,
      valAxisMaxVal: 8600,
      valGridLine: { color: "EDF1F6", size: 1 },
      catGridLine: { style: "none" },
    }
  );

  const notes = [
    [
      "82.5%",
      "of assigned leads never view",
      "7,572.4K people are targeted and never register a view. Deliverability, channel mix and contact policy are the first place to look.",
      CORAL,
    ],
    [
      "92.1%",
      "of viewers never click",
      "1,474.9K see the message and do nothing. Creative, offer relevance and placement own this step.",
      AMBER,
    ],
    [
      "1.3%",
      "of loss is everything else",
      "Click→App-in, Approve and Book together lose 117.3K — a rounding error next to the top two.",
      TEAL,
    ],
  ];
  notes.forEach(([big, lbl, body, c], i) => {
    const y = 1.62 + i * 1.62;
    card(s, 8.55, y, 4.15, 1.42, i === 2 ? TINT_COOL : TINT_WARM, { shadow: false });
    s.addText(big, {
      x: 8.82,
      y: y + 0.1,
      w: 1.5,
      h: 0.45,
      margin: 0,
      fontFace: H,
      fontSize: 22,
      bold: true,
      color: c,
    });
    s.addText(lbl, {
      x: 10.3,
      y: y + 0.13,
      w: 2.2,
      h: 0.42,
      margin: 0,
      valign: "middle",
      fontFace: B,
      fontSize: 11.5,
      bold: true,
      color: INK,
    });
    s.addText(body, {
      x: 8.82,
      y: y + 0.6,
      w: 3.65,
      h: 0.74,
      margin: 0,
      fontFace: B,
      fontSize: 10.5,
      color: "44536B",
    });
  });

  footnote(s, "Loss figures are the arithmetic difference between consecutive stage volumes as printed on the dashboard.");
  s.addNotes(
    "The chart is deliberately on a linear scale: the visual dominance of the first two bars is the finding."
  );
}

/* === 5. Two funnels in one =============================================== */
{
  const s = pres.addSlide();
  s.background = { color: PAPER };
  title(s, "A broken funnel bolted onto a working one", "Diagnosis");

  // left - broken
  card(s, 0.6, 1.62, 5.85, 3.9, TINT_WARM, { shadow: false });
  s.addText("ACQUISITION  ·  LEAD → CLICK", {
    x: 1.0,
    y: 1.92,
    w: 5.05,
    h: 0.3,
    margin: 0,
    fontFace: B,
    fontSize: 11,
    bold: true,
    charSpacing: 1.8,
    color: CORAL,
  });
  s.addText("1.38%", {
    x: 1.0,
    y: 2.28,
    w: 5.05,
    h: 1.0,
    margin: 0,
    fontFace: H,
    fontSize: 62,
    bold: true,
    color: CORAL,
  });
  s.addText("of assigned leads reach a click", {
    x: 1.0,
    y: 3.3,
    w: 5.05,
    h: 0.32,
    margin: 0,
    fontFace: B,
    fontSize: 14,
    color: INK,
  });
  [
    ["9,174.3K in", "1,601.9K viewed  ·  127.0K clicked"],
    ["9,047.3K lost", "before anyone ever reaches a product page"],
    ["Owned by", "targeting, deliverability, creative, placement"],
  ].forEach(([k, v], i) => {
    const y = 3.86 + i * 0.52;
    s.addText(k, {
      x: 1.0,
      y,
      w: 1.85,
      h: 0.38,
      margin: 0,
      valign: "middle",
      fontFace: B,
      fontSize: 12,
      bold: true,
      color: INK,
    });
    s.addText(v, {
      x: 2.9,
      y,
      w: 3.35,
      h: 0.38,
      margin: 0,
      valign: "middle",
      fontFace: B,
      fontSize: 11,
      color: "5A6B82",
    });
  });

  // right - working
  card(s, 6.85, 1.62, 5.85, 3.9, TINT_COOL, { shadow: false });
  s.addText("CREDIT  ·  CLICK → BOOK", {
    x: 7.25,
    y: 1.92,
    w: 5.05,
    h: 0.3,
    margin: 0,
    fontFace: B,
    fontSize: 11,
    bold: true,
    charSpacing: 1.8,
    color: TEAL,
  });
  s.addText("7.64%", {
    x: 7.25,
    y: 2.28,
    w: 5.05,
    h: 1.0,
    margin: 0,
    fontFace: H,
    fontSize: 62,
    bold: true,
    color: TEAL,
  });
  s.addText("of clicks end in a booked account", {
    x: 7.25,
    y: 3.3,
    w: 5.05,
    h: 0.32,
    margin: 0,
    fontFace: B,
    fontSize: 14,
    color: INK,
  });
  [
    ["127.0K in", "38.7K applied  ·  13.0K approved  ·  9.7K booked"],
    ["117.3K lost", "across three stages — the smallest leak in the chain"],
    ["Owned by", "product, credit policy, fulfilment"],
  ].forEach(([k, v], i) => {
    const y = 3.86 + i * 0.52;
    s.addText(k, {
      x: 7.25,
      y,
      w: 1.85,
      h: 0.38,
      margin: 0,
      valign: "middle",
      fontFace: B,
      fontSize: 12,
      bold: true,
      color: INK,
    });
    s.addText(v, {
      x: 9.15,
      y,
      w: 3.35,
      h: 0.38,
      margin: 0,
      valign: "middle",
      fontFace: B,
      fontSize: 11,
      color: "5A6B82",
    });
  });

  card(s, 0.6, 5.72, 12.1, 0.92, NAVY, { shadow: false });
  s.addText(
    [
      { text: "So every extra 1,000 clicks is worth about ", options: { color: "C3D0E2" } },
      { text: "76 bookings", options: { color: MINT, bold: true } },
      { text: " and ", options: { color: "C3D0E2" } },
      { text: "0.8M of set credit", options: { color: MINT, bold: true } },
      { text: " — with no change to credit policy.", options: { color: "C3D0E2" } },
    ],
    {
      x: 1.0,
      y: 5.72,
      w: 11.3,
      h: 0.92,
      margin: 0,
      valign: "middle",
      fontFace: B,
      fontSize: 15,
    }
  );

  footnote(s, "1,000 clicks × 7.64% book rate = 76 bookings × 10.8K average set credit ≈ 0.8M.");
  s.addNotes(
    "This is the central argument of the deck: the money is already being left at the top of the funnel, " +
      "not in credit decisioning."
  );
}

/* === 6. Value delivered ================================================== */
{
  const s = pres.addSlide();
  s.background = { color: NAVY };
  title(s, "What the 46 days actually produced", "Value delivered", true);

  const stats = [
    ["13.0K", "approvals", "33.6% of app-ins", MINT],
    ["9.7K", "bookings", "75.0% of approvals", MINT],
    ["11.3K", "avg approved amount", "per approved customer", PAPER],
    ["10.8K", "avg set credit", "per booked customer", AMBER],
  ];
  stats.forEach(([v, l, sub, c], i) => {
    const y = 1.66 + i * 1.15;
    card(s, 0.6, y, 5.55, 1.0, "17294B", { shadow: false });
    s.addText(v, {
      x: 0.95,
      y,
      w: 1.85,
      h: 1.0,
      margin: 0,
      valign: "middle",
      fontFace: H,
      fontSize: 30,
      bold: true,
      color: c,
    });
    s.addText(l, {
      x: 2.9,
      y: y + 0.18,
      w: 3.05,
      h: 0.34,
      margin: 0,
      fontFace: B,
      fontSize: 14,
      bold: true,
      color: PAPER,
    });
    s.addText(sub, {
      x: 2.9,
      y: y + 0.52,
      w: 3.05,
      h: 0.3,
      margin: 0,
      fontFace: B,
      fontSize: 11,
      color: MUTED_LT,
    });
  });

  s.addChart(
    pres.charts.BAR,
    [
      {
        name: "Customers (K)",
        labels: ["App-in", "Approve", "Book"],
        values: [38.7, 13.0, 9.7],
      },
    ],
    {
      x: 6.5,
      y: 1.6,
      w: 6.25,
      h: 3.3,
      barDir: "col",
      barGapWidthPct: 60,
      chartColors: ["2A8291", "3EA79B", MINT],
      varyColors: true,
      showTitle: true,
      title: "The money end of the funnel (thousands of customers)",
      titleFontFace: B,
      titleFontSize: 12.5,
      titleColor: PAPER,
      showValue: true,
      dataLabelPosition: "outEnd",
      dataLabelFontFace: B,
      dataLabelFontSize: 12,
      dataLabelColor: PAPER,
      dataLabelFormatCode: "#,##0.0",
      showLegend: false,
      catAxisLabelFontFace: B,
      catAxisLabelFontSize: 12,
      catAxisLabelColor: "C3D0E2",
      valAxisLabelFontFace: B,
      valAxisLabelFontSize: 10,
      valAxisLabelColor: MUTED_LT,
      valAxisMinVal: 0,
      valAxisMaxVal: 45,
      valGridLine: { color: "24395F", size: 1 },
      catGridLine: { style: "none" },
    }
  );

  card(s, 6.5, 5.08, 6.25, 1.5, "17294B", { shadow: false });
  s.addText("Set credit lands ~4% below approved value", {
    x: 6.85,
    y: 5.24,
    w: 5.6,
    h: 0.34,
    margin: 0,
    fontFace: H,
    fontSize: 15,
    bold: true,
    color: AMBER,
  });
  s.addText(
    "10.8K set per booked customer against 11.3K approved. Small per account, but ~5M of headroom across the 9.7K book — worth confirming whether it is customer choice or a fulfilment default.",
    {
      x: 6.85,
      y: 5.62,
      w: 5.6,
      h: 0.82,
      margin: 0,
      fontFace: B,
      fontSize: 11,
      color: "C3D0E2",
    }
  );

  footnote(
    s,
    "Dashboard prints approve amount and set credit amount both as 0.1B; the per-customer averages above give ~0.147B approved and ~0.105B set.",
    true
  );
  s.addNotes(
    "Amounts are shown in dashboard units, no currency stated on the source. The 0.1B rounding on both amount tiles " +
      "is why the per-customer averages are the more useful figures."
  );
}

/* === 7. Data integrity flags ============================================= */
{
  const s = pres.addSlide();
  s.background = { color: PAPER };
  title(s, "Fix these five before acting on the numbers", "Data integrity");

  const flags = [
    [
      "Conversions exceed detail-page views",
      "38.7K conversions against 31.7K detail-page views. In a strict click-path funnel that cannot happen — a Loose Criteria attribution artefact.",
      CORAL,
    ],
    [
      "App-in is identical to conversion",
      "Both read 38.7K at exactly 30.5% of clicks. The two columns look like one event under two names; confirm the definitions.",
      CORAL,
    ],
    [
      "App-in amount is empty",
      "Reads 0.0B while approve amount reads 0.1B. Approved value cannot exceed applied value — amount is not being captured at application.",
      CORAL,
    ],
    [
      "Thru-The-Door is effectively unmeasured",
      "TTD1 at 0.3K (0.2% of clicks) and TTD2 at 0.0K. Either the offline/branch journey is not instrumented, or the columns should be retired.",
      AMBER,
    ],
    [
      "Loose vs Strict criteria not pinned",
      "The dashboard itself warns that Loose Criteria can push response rates above 100%. Every rate in this deck should be re-baselined on Strict View-and-Click.",
      AMBER,
    ],
  ];

  flags.slice(0, 3).forEach(([head, body, c], i) => {
    const x = 0.6 + i * 4.07;
    card(s, x, 1.6, 3.86, 2.32, TINT, { shadow: false });
    badge(s, x + 0.3, 1.86, 0.5, c, "!");
    s.addText(head, {
      x: x + 0.3,
      y: 2.5,
      w: 3.26,
      h: 0.62,
      margin: 0,
      valign: "top",
      fontFace: H,
      fontSize: 14,
      bold: true,
      color: INK,
    });
    s.addText(body, {
      x: x + 0.3,
      y: 3.1,
      w: 3.26,
      h: 0.72,
      margin: 0,
      fontFace: B,
      fontSize: 10.5,
      color: "44536B",
    });
  });

  flags.slice(3).forEach(([head, body, c], i) => {
    const x = 0.6 + i * 6.19;
    card(s, x, 4.12, 5.98, 1.92, TINT_WARM, { shadow: false });
    badge(s, x + 0.32, 4.38, 0.5, c, "!");
    s.addText(head, {
      x: x + 1.0,
      y: 4.38,
      w: 4.66,
      h: 0.5,
      margin: 0,
      valign: "middle",
      fontFace: H,
      fontSize: 14,
      bold: true,
      color: INK,
    });
    s.addText(body, {
      x: x + 0.32,
      y: 4.98,
      w: 5.34,
      h: 1.22,
      margin: 0,
      valign: "top",
      fontFace: B,
      fontSize: 11,
      color: "44536B",
    });
  });

  footnote(
    s,
    "None of these invalidate the direction of the findings — the top-funnel gap is far too large to be a measurement artefact — but they do move the exact rates."
  );
  s.addNotes(
    "Frame these as prerequisites for decisions, not as a reason to stall: the 98.7% top-funnel loss survives any plausible re-baselining."
  );
}

/* === 8. Quantified opportunity =========================================== */
{
  const s = pres.addSlide();
  s.background = { color: PAPER };
  title(s, "Three levers, 6.3K incremental bookings", "Sizing the prize");

  s.addChart(
    pres.charts.BAR,
    [
      {
        name: "Incremental bookings (K)",
        labels: [
          "A · Click rate 7.9% → 10.0%",
          "B · View rate 17.5% → 20.0%",
          "C · Book rate 75% → 85%",
          "All three together",
        ],
        values: [2.57, 1.38, 1.3, 6.28],
      },
    ],
    {
      x: 0.5,
      y: 1.52,
      w: 7.75,
      h: 4.55,
      barDir: "bar",
      barGapWidthPct: 50,
      chartColors: ["2A8291", "3EA79B", "5FD6BE", AMBER],
      varyColors: true,
      showTitle: true,
      title: "Incremental bookings per lever (thousands)",
      titleFontFace: B,
      titleFontSize: 13,
      titleColor: INK,
      showValue: true,
      dataLabelPosition: "outEnd",
      dataLabelFontFace: B,
      dataLabelFontSize: 11.5,
      dataLabelColor: INK,
      dataLabelFormatCode: "#,##0.00",
      showLegend: false,
      catAxisLabelFontFace: B,
      catAxisLabelFontSize: 11,
      catAxisLabelColor: INK,
      valAxisLabelFontFace: B,
      valAxisLabelFontSize: 10,
      valAxisLabelColor: MUTED,
      valAxisMinVal: 0,
      valAxisMaxVal: 7.2,
      valGridLine: { color: "EDF1F6", size: 1 },
      catGridLine: { style: "none" },
    }
  );

  const levers = [
    ["A", "Click rate", "+2.57K books  ·  ~28M set credit", "Highest leverage. 33.6K extra clicks off the existing 1.60M views — creative and offer work, no new reach.", TEAL],
    ["B", "View rate", "+1.38K books  ·  ~15M set credit", "229K extra views off leads already assigned. Deliverability, channel mix, frequency capping.", TEAL],
    ["C", "Book rate", "+1.30K books  ·  ~14M set credit", "Cheapest of the three: no new traffic at all, just closing 1.3K already-approved customers.", TEAL],
  ];
  levers.forEach(([tag, name, gain, body, c], i) => {
    const y = 1.6 + i * 1.55;
    card(s, 8.55, y, 4.15, 1.36, TINT, { shadow: false });
    badge(s, 8.82, y + 0.16, 0.46, c, tag);
    s.addText(name, {
      x: 9.42,
      y: y + 0.14,
      w: 3.05,
      h: 0.32,
      margin: 0,
      fontFace: H,
      fontSize: 14,
      bold: true,
      color: INK,
    });
    s.addText(gain, {
      x: 9.42,
      y: y + 0.45,
      w: 3.05,
      h: 0.28,
      margin: 0,
      fontFace: B,
      fontSize: 10.5,
      bold: true,
      color: TEAL,
    });
    s.addText(body, {
      x: 8.82,
      y: y + 0.76,
      w: 3.65,
      h: 0.52,
      margin: 0,
      fontFace: B,
      fontSize: 10,
      color: "44536B",
    });
  });

  card(s, 8.55, 6.25, 4.15, 0.62, NAVY, { shadow: false });
  s.addText("Together: 9.7K → 16.0K bookings", {
    x: 8.55,
    y: 6.25,
    w: 4.15,
    h: 0.62,
    margin: 0,
    align: "center",
    valign: "middle",
    fontFace: H,
    fontSize: 14,
    bold: true,
    color: MINT,
  });

  footnote(
    s,
    "Each lever holds all downstream rates constant at their observed values; set-credit uplift priced at the observed 10.8K average per booked customer."
  );
  s.addNotes(
    "Combined case compounds: 9,174.3K × 20% view × 10% click × 30.5% app-in × 33.6% approve × 85% book = 15.98K bookings."
  );
}

/* === 9. Recommendations ================================================== */
{
  const s = pres.addSlide();
  s.background = { color: NAVY };
  title(s, "What to do next", "Recommendations", true);

  const recs = [
    ["Re-baseline on Strict criteria", "Rerun the funnel under Strict View-and-Click before any budget or targeting decision is taken off these numbers.", "Analytics", "This week"],
    ["Fix amount capture at app-in", "0.0B applied against 0.1B approved is not a rounding issue. Reconcile the App-in and Conversion definitions at the same time.", "Data eng.", "This week"],
    ["Attack the click rate first", "7.9% of viewers click. This is the single highest-leverage number on the dashboard — worth ~2.6K bookings for 2.1 points.", "Marketing", "Next cycle"],
    ["Audit the 82.5% who never view", "7.57M assigned leads register no view at all. Separate 'not delivered' from 'delivered, not opened' before spending on creative.", "Campaign ops", "Next cycle"],
    ["Close the approved-not-booked gap", "25% of approvals never book, and set credit lands 4% under approved. No new traffic needed — ~14M of set credit.", "Fulfilment", "Next cycle"],
    ["Instrument or retire Thru-The-Door", "0.3K and 0.0K across 52 campaigns means the columns currently mislead more than they inform.", "Analytics", "Next quarter"],
  ];

  recs.forEach(([head, body, owner, when], i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = 0.6 + col * 6.19;
    const y = 1.66 + row * 1.72;
    card(s, x, y, 5.98, 1.52, "17294B", { shadow: false });
    badge(s, x + 0.32, y + 0.26, 0.5, TEAL, String(i + 1));
    s.addText(head, {
      x: x + 1.0,
      y: y + 0.26,
      w: 4.66,
      h: 0.5,
      margin: 0,
      valign: "middle",
      fontFace: H,
      fontSize: 14.5,
      bold: true,
      color: PAPER,
    });
    s.addText(body, {
      x: x + 0.32,
      y: y + 0.8,
      w: 5.34,
      h: 0.42,
      margin: 0,
      fontFace: B,
      fontSize: 10.5,
      color: "C3D0E2",
    });
    s.addText(`${owner}   ·   ${when}`, {
      x: x + 0.32,
      y: y + 1.2,
      w: 5.34,
      h: 0.26,
      margin: 0,
      fontFace: B,
      fontSize: 10,
      bold: true,
      charSpacing: 0.8,
      color: MINT,
    });
  });

  footnote(s, "Owners are indicative — to be confirmed with the campaign steering group.", true);
  s.addNotes(
    "Sequencing matters: items 1 and 2 are prerequisites, 3 to 5 are the value levers, 6 is housekeeping."
  );
}

/* === 10. Appendix ======================================================== */
{
  const s = pres.addSlide();
  s.background = { color: PAPER };
  title(s, "Full funnel reference", "Appendix");

  const head = ["Stage", "Volume", "Stage rate", "% of leads", "Note"];
  const body = [
    ["Assigned Lead", "9,174.3K", "—", "100%", "Universe of targeted customers"],
    ["View", "1,601.9K", "17.5% of leads", "17.46%", "7,572.4K never registered a view"],
    ["Click", "127.0K", "7.9% of views", "1.38%", "1,474.9K viewed and did not click"],
    ["Detail Page View", "31.7K", "25.0% of clicks", "0.35%", "Product detail page reached"],
    ["Thru The Door 1", "0.3K", "0.2% of clicks", "0.003%", "Offline / branch — effectively nil"],
    ["Thru The Door 2", "0.0K", "0.0% of clicks", "0.00%", "No recorded volume"],
    ["Conversion", "38.7K", "30.5% of clicks", "0.42%", "Exceeds detail-page views — see flags"],
    ["App-in", "38.7K", "30.5% of clicks", "0.42%", "Identical to conversion count"],
    ["Approve", "13.0K", "33.6% of app-ins", "0.14%", "Approved amount 0.1B (11.3K avg)"],
    ["Book", "9.7K", "75.0% of approvals", "0.11%", "Set credit 0.1B (10.8K avg)"],
  ];

  const rows = [
    head.map((t) => ({
      text: t,
      options: { bold: true, color: PAPER, fill: { color: NAVY }, fontFace: B, fontSize: 11.5 },
    })),
    ...body.map((r, i) =>
      r.map((t, j) => ({
        text: t,
        options: {
          color: j === 4 ? "5A6B82" : INK,
          bold: j === 0,
          fill: { color: i % 2 ? TINT : PAPER },
          fontFace: j === 1 || j === 3 ? H : B,
          fontSize: 11,
        },
      }))
    ),
  ];

  s.addTable(rows, {
    x: 0.6,
    y: 1.58,
    w: 12.1,
    colW: [2.55, 1.6, 2.15, 1.35, 4.45],
    rowH: 0.42,
    border: { type: "solid", color: RULE, pt: 0.5 },
    align: "left",
    valign: "middle",
    margin: [0.04, 0.1, 0.04, 0.1],
  });

  s.addText(
    "Amounts: total 6.1M (0.2K avg)  ·  app-in 0.0B (not captured)  ·  approve 0.1B  ·  set credit 0.1B",
    {
      x: 0.6,
      y: 6.34,
      w: 12.1,
      h: 0.3,
      margin: 0,
      fontFace: B,
      fontSize: 11,
      bold: true,
      color: INK,
    }
  );
  footnote(
    s,
    "Source: CLICX Camp Campaign Performance Dashboard (campaign-level), Loose Criteria view. 52 campaigns, start window 17 Jun – 1 Aug 2026, data as of 2 Aug 2026."
  );
  s.addNotes("Reference table — volumes and stage rates exactly as printed; % of leads computed.");
}

const out = path.join(__dirname, "CLICX_Camp_Key_Insights.pptx");
pres.writeFile({ fileName: out }).then(() => console.log("wrote", out));
