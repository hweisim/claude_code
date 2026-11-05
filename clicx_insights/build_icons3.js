const React = require("react"), { renderToStaticMarkup } = require("react-dom/server"), sharp = require("sharp");
const fa = require("react-icons/fa");
const list = { target: "FaBullseye", deposit: "FaPiggyBank", habit: "FaSyncAlt", idle: "FaUserClock", acquire: "FaUserPlus" };
(async () => {
  for (const [k, n] of Object.entries(list)) {
    const svg = renderToStaticMarkup(React.createElement(fa[n], { size: 256, color: "#FFFFFF" }));
    await sharp(Buffer.from(svg)).resize(256, 256, { fit: "contain", background: { r: 0, g: 0, b: 0, alpha: 0 } }).png().toFile(`icons/${k}.png`);
  }
  console.log("done");
})();
