// Product holding by region: same layout as the age-and-job appendix slide.
// Data: segment_v2.xlsx, 1,507,418 customers (1 Oct 2026 snapshot, Oct payments), 8 largest provinces grouped into regions.
module.exports = function regionSlide(pres, o) {
  const NAVY = "1F3A5F", BLUE = "2E6DB4", CYAN = "1AA39A", GREY = "8A97A6", LGREY = "D3D9E0", TXT = "1C2733", MUTED = "5B6B7C",
    PANEL = "EEF2F6", WHITE = "FFFFFF", TINT = "CDEAE7", ROSE = "F4DCDC";
  const s = pres.addSlide({ masterName: o.master, sectionTitle: "Appendix" });
  s.addText(o.tracker, { placeholder: "tracker" });
  s.addText("Region shapes saving, not borrowing: Bangkok customers save twice as often as those in other provinces, while about 11% borrow everywhere", { placeholder: "title" });
  s.addText("1,507,418 customers by current address; balances as of 1 Oct 2026, payments in Oct. Borrow = ever booked a loan; save = balance over ฿100; pay = paid in-app in Oct. The file names only the 8 largest provinces, grouped here into regions; Northeast and North are one province each. Source: segment_v2 extract", { placeholder: "foot" });
  const T = (t, x) => s.addText(t, { margin: 0, isTextBox: true, ...x });
  const RR = (x, y, w, h, fill, name) => s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.08, fill: { color: fill }, line: { type: "none" }, objectName: name });
  const badge = (x, y, n) => {
    s.addShape(pres.shapes.OVAL, { x, y, w: 0.42, h: 0.42, fill: { color: NAVY }, line: { type: "none" }, objectName: `Region badge ${n} bg` });
    T(String(n), { x, y, w: 0.42, h: 0.42, fontSize: 13, bold: true, color: WHITE, align: "center", valign: "middle", objectName: `Region badge ${n}` });
  };
  // [region, provinces, customers, borrow, save, pay, applied, zero balance, over 20K]
  const R = [
    ["Bangkok", "", 267926, 10.8, 14.0, 9.3, 80.0, 69.7, 4.8],
    ["Greater Bangkok", "Nonthaburi, Pathum Thani, Samut Prakan", 191923, 10.7, 10.2, 8.7, 85.2, 73.7, 2.6],
    ["East", "Chon Buri, Rayong", 113530, 10.9, 6.9, 8.2, 89.3, 77.7, 1.2],
    ["Northeast*", "Nakhon Ratchasima", 52858, 11.8, 6.3, 8.3, 88.8, 78.4, 0.9],
    ["North*", "Chiang Mai", 39007, 13.1, 9.8, 10.6, 85.3, 72.4, 1.8],
    ["Other provinces", "", 842174, 10.9, 6.5, 8.0, 87.7, 78.0, 1.0],
  ];
  const ALL = [1507418, 11.0, 8.4, 8.4, 86.1, 75.8, 1.9];
  const y = 1.62, H = 5.15, LW = 7.75, RX = 0.5 + LW + 0.15, RW = 12.83 - RX;
  // ---- panel 1: grouped columns
  RR(0.5, y, LW, H, PANEL, "Region chart panel");
  badge(0.68, y + 0.18, 1);
  T("By region: borrowing is flat, saving drops outside Bangkok, payments vary little", { x: 1.22, y: y + 0.12, w: LW - 0.9, h: 0.55, fontSize: 13.5, bold: true, color: NAVY, valign: "middle", objectName: "Region head 1" });
  T("% of each region holding each product", { x: 0.7, y: y + 0.78, w: 4, h: 0.26, fontSize: 11, bold: true, color: NAVY, objectName: "Region chart title" });
  let lx = 5.0;
  [["Borrow", NAVY], ["Save", CYAN], ["Pay", BLUE]].forEach((l) => {
    s.addShape(pres.shapes.RECTANGLE, { x: lx, y: y + 0.84, w: 0.15, h: 0.15, fill: { color: l[1] }, line: { type: "none" }, objectName: `Region legend key ${l[0]}` });
    T(l[0], { x: lx + 0.2, y: y + 0.78, w: 0.65, h: 0.26, fontSize: 10, color: MUTED, valign: "middle", objectName: `Region legend ${l[0]}` });
    lx += 0.85;
  });
  const gx = 0.75, gw = LW - 0.5, gy = y + 1.4, gh = 2.3, max = 16, cw = gw / R.length, bw = cw * 0.25;
  s.addShape(pres.shapes.LINE, { x: gx, y: gy + gh, w: gw, h: 0, line: { color: LGREY, width: 1 }, objectName: "Region axis" });
  R.forEach((r, i) => {
    const x0 = gx + i * cw + cw * 0.125;
    [[NAVY, 3, "Borrow"], [CYAN, 4, "Save"], [BLUE, 5, "Pay"]].forEach((p, j) => {
      const v = r[p[1]], bh = gh * v / max, bx = x0 + j * bw;
      s.addShape(pres.shapes.RECTANGLE, { x: bx + 0.015, y: gy + gh - bh, w: bw - 0.03, h: bh, fill: { color: p[0] }, line: { type: "none" }, objectName: `Region ${p[2]} ${i}` });
      T(v.toFixed(1), { x: bx - 0.04, y: gy + gh - bh - 0.21, w: bw + 0.08, h: 0.19, fontSize: 8.5, bold: j === 1 && v >= 10, color: TXT, align: "center", valign: "bottom", objectName: `Region ${p[2]} value ${i}` });
    });
    T([{ text: r[0], options: { bold: true, color: TXT, breakLine: !!r[1] } }].concat(r[1] ? [{ text: r[1], options: { fontSize: 7.5, color: MUTED } }] : []),
      { x: gx + i * cw, y: gy + gh + 0.06, w: cw, h: 0.5, fontSize: 9.5, align: "center", valign: "top", objectName: `Region label ${i}` });
  });
  RR(0.7, y + 4.3, LW - 0.4, 0.68, WHITE, "Region so box");
  T([{ text: "SO  ", options: { bold: true, color: CYAN, fontSize: 9.5, charSpacing: 1 } }, { text: "Lead with savings and investing in Bangkok and Nonthaburi; outside Bangkok, lead with a first deposit alongside the loan most customers apply for", options: { color: TXT, fontSize: 10.5 } }],
    { x: 0.85, y: y + 4.3, w: LW - 0.7, h: 0.68, valign: "middle", objectName: "Region so text" });
  // ---- panel 2: outside Bangkok, more apply and fewer keep money
  RR(RX, y, RW, H, PANEL, "Region table panel");
  badge(RX + 0.18, y + 0.18, 2);
  T("Outside Bangkok, more apply for a loan and fewer keep money with us", { x: RX + 0.72, y: y + 0.12, w: RW - 0.85, h: 0.55, fontSize: 13, bold: true, color: NAVY, valign: "middle", objectName: "Region head 2" });
  const hd = ["Region", "Customers", "Applied for a loan", "฿0 balance", "Over ฿20K"].map((t, j) => ({ text: t, options: { bold: true, fontSize: 9, color: WHITE, fill: { color: NAVY }, align: j ? "right" : "left", valign: "middle" } }));
  const fmtK = (n) => (n >= 100000 ? Math.round(n / 1000) + "K" : (n / 1000).toFixed(1) + "K");
  const shade = (v, base, higherIsGood) => (v >= base * 1.15 ? (higherIsGood ? TINT : ROSE) : (v <= base * 0.85 ? (higherIsGood ? ROSE : TINT) : null));
  const rows = R.map((r, i) => {
    const cells = [r[0], fmtK(r[2]), r[6].toFixed(0) + "%", r[7].toFixed(0) + "%", r[8].toFixed(1) + "%"];
    const fills = [null, null, null, null, shade(r[8], ALL[6], true)];
    return cells.map((t, j) => ({ text: t, options: { fontSize: 10, color: TXT, bold: j === 0, align: j ? "right" : "left", valign: "middle", fill: { color: fills[j] || (i % 2 ? WHITE : "F7F8FA") } } }));
  });
  rows.push(["All customers", "1.51M", "86%", "76%", "1.9%"].map((t, j) => ({ text: t, options: { fontSize: 10, bold: true, color: NAVY, align: j ? "right" : "left", valign: "middle", fill: { color: WHITE }, border: [{ pt: 1.25, color: NAVY }, { type: "none" }, { type: "none" }, { type: "none" }] } })));
  const tw = RW - 0.36;
  s.addTable([hd, ...rows], { x: RX + 0.18, y: y + 0.85, w: tw, colW: [1.25, 0.8, 0.75, 0.62, 0.62].map((c) => c * tw / 4.04), rowH: [0.46, ...rows.map(() => 0.35)], fontFace: "Arial", border: { type: "solid", pt: 0.5, color: "E3E8EE" }, margin: [0.03, 0.05, 0.03, 0.05], objectName: "Region table" });
  T([{ text: "4.8% vs 1.0%  ", options: { bold: true, color: CYAN, fontSize: 13 } }, { text: "of Bangkok vs other-province customers hold over ฿20K. Shading: 15%+ above (teal) or below (rose) the all-customer rate. * One province each; 56% of customers live in other provinces", options: { fontSize: 9.5, color: TXT } }],
    { x: RX + 0.2, y: y + 3.95, w: RW - 0.4, h: 1.05, valign: "middle", objectName: "Region coverage note" });
};
