const React = require("react"), { renderToStaticMarkup } = require("react-dom/server"), sharp = require("sharp");
const fa = require("react-icons/fa");
const list = { salary: "FaMoneyCheckAlt", insure: "FaShieldAlt", invest: "FaChartLine", enabler: "FaSlidersH",
  p_corp: "FaUserTie", p_sme: "FaStore", p_helper: "FaHandsHelping", p_earner: "FaMoneyBillWave", p_owner: "FaBriefcase",
  p_rider: "FaMotorcycle", p_blue: "FaHardHat", p_white: "FaLaptop", p_wealth: "FaSeedling", p_tech: "FaMobileAlt" };
(async () => {
  for (const [k, n] of Object.entries(list)) {
    if (!fa[n]) { console.log("missing", n); continue; }
    const svg = renderToStaticMarkup(React.createElement(fa[n], { size: 256, color: "#FFFFFF" }));
    await sharp(Buffer.from(svg)).resize(256, 256, { fit: "contain", background: { r: 0, g: 0, b: 0, alpha: 0 } }).png().toFile(`icons/${k}.png`);
  }
  console.log("done");
})();
