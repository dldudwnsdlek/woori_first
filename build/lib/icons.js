// react-icons(Lucide) → PNG(base64). 빌드 시작 시 한 번에 래스터화해 캐시합니다.
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const lu = require("react-icons/lu");
const { HEX } = require("./theme");

const cache = new Map();

async function prepare(list) {
  for (const [name, colorKey] of list) {
    const key = `${name}:${colorKey}`;
    if (cache.has(key)) continue;
    const Icon = lu[name];
    if (!Icon) throw new Error(`아이콘 없음: ${name}`);
    const color = "#" + (HEX[colorKey] || colorKey);
    const svg = ReactDOMServer.renderToStaticMarkup(
      React.createElement(Icon, { size: 384, color, strokeWidth: 2 })
    );
    const png = await sharp(Buffer.from(svg)).resize(384, 384).png().toBuffer();
    cache.set(key, "image/png;base64," + png.toString("base64"));
  }
}

function get(name, colorKey) {
  const key = `${name}:${colorKey}`;
  if (!cache.has(key)) throw new Error(`아이콘 미준비: ${key} — ICONS 목록에 추가하세요`);
  return cache.get(key);
}

module.exports = { prepare, get };
