# Slide Style Guide

Derived from the reviewed final versions of `Lending_2026.08.07.pptx` (6 slides) and the
PTT OR one-pager. These are the house conventions — follow them by default on every deck
unless told otherwise.

---

## 1. Canvas

| Property | Value |
|---|---|
| Slide size | 13.333 × 7.5 in (`LAYOUT_WIDE`) |
| Left / right margin | 0.58 in (content width 12.33) |
| Background | White |
| Card fill | `F1F4F9`, rounded corners, **no border, no shadow** |

Content starts at y ≈ 1.5 and ends by y ≈ 6.5. Leave the bottom strip empty — do not
stretch panels to fill it.

---

## 2. Typography

**Body/UI font: `Graphik TH`.** Used for titles, insight lines, tile text, axis and data
labels. `Cambria` survives only for the small bold panel headers inside cards
("Cash need", "Charge type").

| Element | Size | Weight | Colour |
|---|---|---|---|
| Slide title | 22 pt | bold | `045D66` |
| Asterisk note (under title) | 11.5 pt | italic | `6B6B6B` |
| Totals sub-header | 12–14 pt | bold | `045D66` |
| Insight line | 14 pt | regular | near-black |
| Panel header (in card) | 12 pt | bold | `16233D` (Cambria) |
| Chart title | 11–12 pt | regular | `16233D` |
| Measure sub-header | 10 pt | regular | `00A896` |
| Stat tile value | 25 pt | bold | `00A896` or `C00000` |
| Stat tile label | 11.5 pt | bold | `16233D` |
| Stat tile sub | 10 pt | regular | `6B7A90` |
| Axis / data labels | 7.5–10 pt | regular | ink or white on fill |

> **QA caveat:** Graphik TH is not installed in the render environment, so LibreOffice
> substitutes it and text-fit checks are approximate. Leave ~10% slack on any container
> where wrapping would break the layout.

---

## 3. Palette

```
045D66  title teal          — slide titles, totals sub-headers
16233D  ink                 — labels, panel headers
6B7A90  muted               — tile sub-lines, captions
6B6B6B  footnote grey       — asterisk notes
00A896  teal accent         — good/neutral hero numbers, measure sub-headers
C00000  alert red           — risk figures, highlight outlines
5B7B9A  badge slate         — numbered callout circles
F1F4F9  card fill
```

**Sequential ramp** (low → high, dark enough for white labels at every step):

```
8DA0B8  R1   null / none
3FB89F  R2   low
1E8E86  R3   mid
165A73  R4   high
```

Use the ramp for every ordinal series (Null/Low/Mid/High, tenure bands, charge type).
For binary flags use `R1` (N) and `R4` (Y).

**Red is reserved** for figures the reader should worry about (high risk, high cash need)
and for highlight outlines. Never decorative.

---

## 4. Slide anatomy

```
┌───────────────────────────────────────────────────────────┐
│ Factual title with the key number inline          22pt    │  y 0.31
│ *optional asterisk caveat                       11.5pt    │  y 0.82
│                                                            │
│ CL:801M   OS: 476M                    bold teal  12-14pt  │  y ~1.21
│ Plain-English insight sentence.                    14pt   │  y ~1.35
│                                                            │
│ ┌─── card ────┐ ┌─── card ────┐                           │  y 1.5+
│ │  chart      │ │  chart      │                           │
│ └─────────────┘ └─────────────┘                           │
└───────────────────────────────────────────────────────────┘
```

**Title** is factual and carries the totals inline — not an editorial headline:

- ✅ `AIS BASE Information : 70.1K - DL 64.8K – PL 6.3K`
- ✅ `Total Loan Booked as of 5 Aug 2026: 94,371`
- ✅ `Appendix: Cash Need & Risk Level - Credit Limit and OS Balance`
- ❌ `Weaker segments draw more of their line`
- ❌ Separate small-caps kicker line above the title

**Insight line** sits directly under the title as plain dark text at 14pt. One or two
sentences, no italics, no coloured band. This is where the "so what" lives.

**Two-tone titles** are fine when a brand leads: `PTT OR` in bright cyan, the rest in ink.

---

## 5. Components

### Stat tile
Card + 25pt value + 11.5pt bold label + 10pt muted sub-line. Teal by default, red when the
number is a risk. A tile may carry a second smaller value block beneath (e.g. "3.01 Million
/ Ever have Blue+ Application").

### Labelled bar rows
For a small set of named quantities (consortium penetration, segment split): row label
left, proportional rounded bar centre, `value · percent` right-aligned. Bars scaled to the
largest value in the set, not to 100.

### Chart panel
Card, optional Cambria panel header, then the chart. For a measure split, stack two charts
inside one card with `CREDIT LIMIT` / `OUTSTANDING` sub-headers at 10pt teal, legend on the
lower chart only.

### Numbered callout
A `5B7B9A` circle with a white bold number beside the insight sentence, and the same number
on a `C00000` outlined rectangle drawn over the exact bars that support it. Use when a
statement points at a specific region of a specific chart.

### Venn
For overlapping populations, hand-build circles with segment counts in near-black 12.5pt
bold and set labels in the ramp colours. Better than a table when the overlap is the point.

---

## 6. Chart conventions

- **Percent-stacked horizontal bars** are the default form.
- Value axis hidden; no gridlines on either axis.
- Legend at the bottom, 9–10pt.
- **Category labels carry the bar's total**: `PL 226.3M`, `2-Aug 325.1M`.
- Data labels centred inside segments, white on fill.
- **Suppress labels too narrow to fit** with a conditional format — `[<6]"";0"%"` — never
  let a label clip into a wrong number (`1.6%` rendering as `6%`).
- Percentages as integers by default; one decimal only where the precision carries the
  argument.
- **PL before DL**, everywhere, charts and tables alike. Horizontal bars plot the first
  label at the bottom, so pass `["DL", "PL"]` to get PL on top.
- A 100% stacked chart normalises each category independently, so rows may be pre-scaled to
  different units (counts in millions, amounts in billions) to keep labels readable.

---

## 7. Numbers and wording

- Titles and sub-headers state totals: `CL:801M   OS: 476M`.
- Compact units in charts and tables (`70.1K`, `226.3M`, `1092M`); **spell out "Million"**
  on hero tiles (`9.85 Million`, `฿ 18,850 Million`).
- Currency symbol on the value when known (`฿`), unit spelled out in tile sub-lines
  (`867 THB Per Transaction`, `4 Transactions per month`).
- Round per-user and per-ticket figures to whole numbers on tiles.
- House terminology: **Credit Limit** (not credit line), **OS / Outstanding**,
  **Book Date**, **MTU** for monthly transacting users.
- Write plain English in insight lines. No "profile identically", "part company",
  "punches above its headcount".

---

## 8. Deck structure

- **Lead with 1–2 story slides**, then push everything reference-grade behind an
  `Appendix:` title prefix.
- Keep decks short. The reviewed lending deck went from 12 slides to 6.
- **No raw data tables.** If the numbers matter, put them on the chart labels.
- Order: headline → supporting detail → appendix.

---

## 9. Do not

- ❌ Dark navy narrative bands across the bottom of a slide
- ❌ Slide-bottom source/method footnotes (use a short `*asterisk` note under the title)
- ❌ Small-caps kicker lines above the title
- ❌ Editorial or punning headlines
- ❌ Card borders, drop shadows, or accent stripes
- ❌ Data tables as slides
- ❌ Labels that clip or overflow their container

---

## 10. Build and QA

1. Generate with `pptxgenjs`; keep source data in `data/*.tsv` so charts and text can't
   disagree.
2. Render: `soffice --headless --convert-to pdf`, then rasterise with PyMuPDF at 110–130 dpi.
3. **Inspect every slide image.** Check first for clipped or overflowing text, then
   overlaps, uneven gaps, and low-contrast labels.
4. Re-render after every fix — the PDF must be regenerated before the images reflect changes.
5. Cross-check every derived figure against the source totals and state any gap on the slide.

### Reference constants

```js
const TITLE_TEAL = "045D66", INK = "16233D", MUTED = "6B7A90", NOTE = "6B6B6B";
const TEAL = "00A896", RED = "C00000", BADGE = "5B7B9A", CARD = "F1F4F9";
const R1 = "8DA0B8", R2 = "3FB89F", R3 = "1E8E86", R4 = "165A73";
const F = "Graphik TH", FH = "Cambria";   // body / panel headers

// title block
title:   { x: 0.58, y: 0.31, w: 12.33, size: 22,   bold: true, color: TITLE_TEAL }
note:    { x: 0.58, y: 0.82, w: 12.33, size: 11.5, italic: true, color: NOTE }
totals:  { x: 0.60, y: 1.21, w: 12.10, size: 13,   bold: true, color: TITLE_TEAL }
insight: { x: 0.60, y: 1.35, w: 12.10, size: 14,   color: INK }
```
