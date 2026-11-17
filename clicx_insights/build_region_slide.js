// Product holding by region: same layout as the age-and-job appendix slide.
// Pay is real (Oct 2026 segment extract, 8 largest provinces grouped into regions);
// Borrow and Save have no regional data yet, so they are drawn as dashed placeholders.
module.exports = function regionSlide(pres, o) {
  const NAVY = "1F3A5F", BLUE = "2E6DB4", CYAN = "1AA39A", GREY = "8A97A6", LGREY = "D3D9E0", TXT = "1C2733", MUTED = "5B6B7C",
    PANEL = "EEF2F6", WHITE = "FFFFFF", TINT = "E5F1F0";
  const s = pres.addSlide({ masterName: o.master, sectionTitle: "Appendix" });
  s.addText(o.tracker, { placeholder: "tracker" });
  s.addText("By region, payments sit close to the 7.8% average and only Chiang Mai stands out at 10%; borrowing and saving by region still need data", { placeholder: "title" });
  s.addText("Pay = paid in-app in Oct 2026; segment extract, 1,501,142 customers by current address. The file names only the 8 largest provinces, grouped here into regions; Northeast and North are one province each, and 56% of customers sit in other provinces. Region rates are customer-weighted. Borrow and save by region are not in any file shared so far", { placeholder: "foot" });
  const T = (t, x) => s.addText(t, { margin: 0, isTextBox: true, ...x });
  const RR = (x, y, w, h, fill, name) => s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.08, fill: { color: fill }, line: { type: "none" }, objectName: name });
  const badge = (x, y, n) => {
    s.addShape(pres.shapes.OVAL, { x, y, w: 0.42, h: 0.42, fill: { color: NAVY }, line: { type: "none" }, objectName: `Region badge ${n} bg` });
    T(String(n), { x, y, w: 0.42, h: 0.42, fontSize: 13, bold: true, color: WHITE, align: "center", valign: "middle", objectName: `Region badge ${n}` });
  };
  const R = [
    ["Bangkok", "", 267071, 8.6, 16.9],
    ["Greater Bangkok", "Nonthaburi, Pathum Thani, Samut Prakan", 191274, 8.0, 15.8],
    ["East", "Chon Buri, Rayong", 113077, 7.6, 15.1],
    ["Northeast*", "Nakhon Ratchasima", 52621, 7.7, 15.1],
    ["North*", "Chiang Mai", 38857, 10.0, 18.5],
    ["Other provinces", "", 838242, 7.4, 14.7],
  ];
  const y = 1.62, H = 5.15, LW = 7.75, RX = 0.5 + LW + 0.15, RW = 12.83 - RX;
  // ---- panel 1: grouped columns
  RR(0.5, y, LW, H, PANEL, "Region chart panel");
  badge(0.68, y + 0.18, 1);
  T("By region: payment rates barely move; the North (Chiang Mai) leads at 10%", { x: 1.22, y: y + 0.12, w: LW - 0.9, h: 0.55, fontSize: 13.5, bold: true, color: NAVY, valign: "middle", objectName: "Region head 1" });
  T("% of each region holding each product", { x: 0.7, y: y + 0.78, w: 4, h: 0.26, fontSize: 11, bold: true, color: NAVY, objectName: "Region chart title" });
  // legend
  let lx = 4.35;
  [["Borrow", NAVY, true], ["Save", CYAN, true], ["Pay", BLUE, false]].forEach((l) => {
    s.addShape(pres.shapes.RECTANGLE, { x: lx, y: y + 0.84, w: 0.15, h: 0.15, fill: l[2] ? { type: "none" } : { color: l[1] }, line: l[2] ? { color: l[1], width: 1, dashType: "dash" } : { type: "none" }, objectName: `Region legend key ${l[0]}` });
    T(l[0], { x: lx + 0.2, y: y + 0.78, w: 0.6, h: 0.26, fontSize: 10, color: MUTED, valign: "middle", objectName: `Region legend ${l[0]}` });
    lx += 0.8;
  });
  T("dashed = data needed", { x: lx, y: y + 0.78, w: 1.4, h: 0.26, fontSize: 9, italic: true, color: MUTED, valign: "middle", objectName: "Region legend note" });
  const gx = 0.75, gw = LW - 0.5, gy = y + 1.35, gh = 2.35, max = 12, cw = gw / R.length, bw = cw * 0.24;
  // average line
  const avgY = gy + gh - gh * 7.8 / max;
  s.addShape(pres.shapes.LINE, { x: gx, y: avgY, w: gw, h: 0, line: { color: GREY, width: 1, dashType: "dash" }, objectName: "Region pay average" });
  T("All 7.8%", { x: gx, y: avgY + 0.03, w: 0.7, h: 0.2, fontSize: 8.5, color: MUTED, objectName: "Region pay average label" });
  s.addShape(pres.shapes.LINE, { x: gx, y: gy + gh, w: gw, h: 0, line: { color: LGREY, width: 1 }, objectName: "Region axis" });
  R.forEach((r, i) => {
    const x0 = gx + i * cw + cw * 0.14;
    [[NAVY, "Borrow"], [CYAN, "Save"]].forEach((p, j) => {
      const bx = x0 + j * bw, ph = 0.42;
      s.addShape(pres.shapes.RECTANGLE, { x: bx + 0.02, y: gy + gh - ph, w: bw - 0.04, h: ph, fill: { type: "none" }, line: { color: p[0], width: 1, dashType: "dash" }, objectName: `Region ${p[1]} placeholder ${i}` });
      T("?", { x: bx, y: gy + gh - ph, w: bw, h: ph, fontSize: 10, bold: true, color: p[0], align: "center", valign: "middle", objectName: `Region ${p[1]} mark ${i}` });
    });
    const bx = x0 + 2 * bw, bh = gh * r[3] / max, hi = r[3] >= 9;
    s.addShape(pres.shapes.RECTANGLE, { x: bx + 0.02, y: gy + gh - bh, w: bw - 0.04, h: bh, fill: { color: BLUE }, line: { type: "none" }, objectName: `Region pay bar ${i}` });
    T(r[3].toFixed(1) + "%", { x: bx - 0.15, y: gy + gh - bh - 0.22, w: bw + 0.3, h: 0.2, fontSize: 9.5, bold: hi, color: TXT, align: "center", objectName: `Region pay value ${i}` });
    T([{ text: r[0], options: { bold: true, color: TXT, breakLine: !!r[1] } }].concat(r[1] ? [{ text: r[1], options: { fontSize: 7.5, color: MUTED } }] : []),
      { x: gx + i * cw, y: gy + gh + 0.06, w: cw, h: 0.5, fontSize: 9.5, align: "center", valign: "top", objectName: `Region label ${i}` });
  });
  RR(0.7, y + 4.3, LW - 0.4, 0.68, WHITE, "Region so box");
  T([{ text: "SO  ", options: { bold: true, color: CYAN, fontSize: 9.5, charSpacing: 1 } }, { text: "Region adds little beyond age and job for payments. To plan regional products, we need borrow and save by province for every region, not just the 8 largest provinces", options: { color: TXT, fontSize: 10.5 } }],
    { x: 0.85, y: y + 4.3, w: LW - 0.7, h: 0.68, valign: "middle", objectName: "Region so text" });
  // ---- panel 2: where customers live
  RR(RX, y, RW, H, PANEL, "Region table panel");
  badge(RX + 0.18, y + 0.18, 2);
  T("Where customers live, and who pays", { x: RX + 0.72, y: y + 0.12, w: RW - 0.85, h: 0.55, fontSize: 13.5, bold: true, color: NAVY, valign: "middle", objectName: "Region head 2" });
  const hd = ["Region", "Customers", "Paid Oct", "Paid Aug–Oct"].map((t, j) => ({ text: t, options: { bold: true, fontSize: 9.5, color: WHITE, fill: { color: NAVY }, align: j ? "right" : "left", valign: "middle" } }));
  const fmtK = (n) => (n >= 100000 ? Math.round(n / 1000) + "K" : (n / 1000).toFixed(1) + "K");
  const rows = R.map((r, i) => [r[0], fmtK(r[2]) + "  " + Math.round(100 * r[2] / 1501142) + "%", r[3].toFixed(1) + "%", r[4].toFixed(1) + "%"].map((t, j) => ({ text: t, options: {
    fontSize: 10, color: TXT, bold: j === 0 || (j > 1 && r[3] >= 9), align: j ? "right" : "left", valign: "middle", fill: { color: (j > 1 && r[3] >= 9) ? TINT : (i % 2 ? WHITE : "F7F8FA") } } })));
  rows.push(["All customers", "1.50M  100%", "7.8%", "15.4%"].map((t, j) => ({ text: t, options: { fontSize: 10, bold: true, color: NAVY, align: j ? "right" : "left", valign: "middle", fill: { color: WHITE }, border: [{ pt: 1.25, color: NAVY }, { type: "none" }, { type: "none" }, { type: "none" }] } })));
  s.addTable([hd, ...rows], { x: RX + 0.18, y: y + 0.85, w: RW - 0.36, colW: [1.3, 1.07, 0.66, 0.73].map((c) => c * (RW - 0.36) / 3.76), rowH: [0.42, ...rows.map(() => 0.36)], fontFace: "Arial", border: { type: "solid", pt: 0.5, color: "E3E8EE" }, margin: [0.03, 0.06, 0.03, 0.06], objectName: "Region table" });
  T([{ text: "56% ", options: { bold: true, color: NAVY, fontSize: 14 } }, { text: "of customers live outside the 8 largest provinces, so the file can't show Isan, the North or the South in full. * One province each", options: { fontSize: 10, color: TXT } }],
    { x: RX + 0.2, y: y + 4.0, w: RW - 0.4, h: 0.98, valign: "middle", objectName: "Region coverage note" });
};
