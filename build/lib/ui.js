// 도형·텍스트 헬퍼. 모든 좌표는 inch.
// group(name, fn) 안에서 만든 도형은 objectName 이 `name#n` 이 되고,
// postprocess.py 가 같은 이름끼리 PowerPoint 그룹으로 묶어 애니메이션 대상으로 씁니다.
const { C, FONT, G } = require("./theme");
const icons = require("./icons");

let pres = null;
let cur = { slide: null, group: null, n: 0, label: "" };
const warnings = [];

function init(p) {
  pres = p;
}
function use(slide, label) {
  cur = { slide, group: null, n: 0, label };
}
function group(name, fn) {
  const prev = cur.group;
  cur.group = name;
  fn();
  cur.group = prev;
}
function oname() {
  cur.n += 1;
  return cur.group ? `${cur.group}#${cur.n}` : `el${cur.n}`;
}

// ── 글자 폭 추정 (Pretendard 기준, em 단위) — 넘침 경고용
function emWidth(str, bold) {
  let w = 0;
  for (const ch of str) {
    const c = ch.codePointAt(0);
    if (c >= 0xac00 && c <= 0xd7a3) w += 0.864;
    else if (c >= 0x3130 && c <= 0x318f) w += 0.864;
    else if (ch === " ") w += bold ? 0.23 : 0.25;
    else if (/[0-9]/.test(ch)) w += bold ? 0.64 : 0.6;
    else if (/[A-Z]/.test(ch)) w += bold ? 0.72 : 0.68;
    else if (/[a-z]/.test(ch)) w += bold ? 0.58 : 0.55;
    else if (/[.,·:;'‘’!|]/.test(ch)) w += 0.29;
    else if (/[“”"()\-−~/]/.test(ch)) w += 0.42;
    else if (/[①-⑳○●✓→←›‹×+=※㎡]/.test(ch)) w += 0.9;
    else w += 0.864;
  }
  return w;
}
function checkFit(lines, size, bold, w, h, ls, tag) {
  const em = size / 72;
  let total = 0;
  for (const ln of lines) {
    const lw = emWidth(ln, bold) * em;
    total += Math.max(1, Math.ceil((lw - 0.001) / w));
    if (lw > w * 1.0) {
      warnings.push(`[${cur.label}] 줄바꿈 발생: "${ln}" ${lw.toFixed(2)}in > ${w.toFixed(2)}in (${size}pt) ${tag || ""}`);
    }
  }
  const need = total * em * 1.2 * (ls || 1);
  if (need > h + 0.02) {
    warnings.push(`[${cur.label}] 높이 초과: "${lines[0]}" ${need.toFixed(2)}in > ${h.toFixed(2)}in`);
  }
}

function shadowOf(kind) {
  if (!kind) return undefined;
  if (kind === "lg") return { type: "outer", color: "0A2540", opacity: 0.22, blur: 28, offset: 10, angle: 90 };
  return { type: "outer", color: "0A2540", opacity: 0.1, blur: 10, offset: 3, angle: 90 };
}

// 사각형 / 둥근 사각형
function rect(x, y, w, h, o = {}) {
  const opts = {
    x, y, w, h,
    objectName: oname(),
    fill: o.fill ? { color: o.fill, transparency: o.alpha || 0 } : { type: "none" },
    line: o.line ? { color: o.line, width: o.lw || 1, dashType: o.dash || "solid" } : { type: "none" },
  };
  if (o.shadow) opts.shadow = shadowOf(o.shadow);
  if (o.r) {
    opts.rectRadius = Math.min(o.r, Math.min(w, h) / 2);
    cur.slide.addShape(pres.ShapeType.roundRect, opts);
  } else {
    cur.slide.addShape(pres.ShapeType.rect, opts);
  }
}

function circle(x, y, d, o = {}) {
  cur.slide.addShape(pres.ShapeType.ellipse, {
    x, y, w: d, h: d,
    objectName: oname(),
    fill: o.fill ? { color: o.fill, transparency: o.alpha || 0 } : { type: "none" },
    line: o.line ? { color: o.line, width: o.lw || 1 } : { type: "none" },
  });
}

function line(x1, y1, x2, y2, o = {}) {
  const opts = {
    x: Math.min(x1, x2), y: Math.min(y1, y2),
    w: Math.abs(x2 - x1), h: Math.abs(y2 - y1),
    objectName: oname(),
    line: {
      color: o.color || C.line, width: o.w || 1, dashType: o.dash || "solid",
      endArrowType: o.arrow ? "triangle" : undefined,
    },
  };
  if (x2 < x1) opts.flipH = true;
  if (y2 < y1) opts.flipV = true;
  cur.slide.addShape(pres.ShapeType.line, opts);
}

// 텍스트. str: 문자열 또는 run 배열 [{t, bold, color, size, br, strike}]
function text(str, x, y, w, h, o = {}) {
  const size = o.size || 16;
  const base = {
    x, y, w, h,
    objectName: oname(),
    isTextBox: true,
    margin: 0,
    fontFace: FONT,
    fontSize: size,
    bold: !!o.bold,
    color: o.color || C.navy,
    align: o.align || "left",
    valign: o.valign || "top",
    lineSpacingMultiple: o.ls || 1.0,
    charSpacing: o.cs || 0,
    wrap: true,
    fit: "none",
  };
  if (o.alpha) base.transparency = o.alpha;
  let payload;
  if (Array.isArray(str)) {
    payload = str.map((r) => ({
      text: r.t,
      options: {
        fontFace: FONT,
        bold: r.bold === undefined ? !!o.bold : r.bold,
        color: r.color || o.color || C.navy,
        fontSize: r.size || size,
        breakLine: !!r.br,
        strike: r.strike ? "sngStrike" : undefined,
      },
    }));
    // 줄 단위로 합쳐서 폭 검사
    const lines = [""];
    for (const r of str) {
      const parts = r.t.split("\n");
      parts.forEach((p, i) => {
        if (i > 0) lines.push("");
        lines[lines.length - 1] += p;
      });
      if (r.br) lines.push("");
    }
    if (!o.nocheck) checkFit(lines.filter((l) => l.length), size, !!o.bold, w, h, o.ls, "runs");
  } else {
    payload = str;
    if (!o.nocheck) checkFit(String(str).split("\n"), size, !!o.bold, w, h, o.ls);
  }
  cur.slide.addText(payload, base);
}

// 텍스트가 들어간 알약/버튼 (도형 하나)
function pill(str, x, y, w, h, o = {}) {
  const size = o.size || 12;
  if (!o.nocheck) checkFit([String(str)], size, o.bold !== false, w - 0.1, h + 0.2, 1, "pill");
  cur.slide.addText(str, {
    x, y, w, h,
    objectName: oname(),
    shape: pres.ShapeType.roundRect,
    rectRadius: Math.min(o.r === undefined ? h / 2 : o.r, h / 2),
    fill: o.fill ? { color: o.fill } : { type: "none" },
    line: o.line ? { color: o.line, width: o.lw || 1, dashType: o.dash || "solid" } : { type: "none" },
    fontFace: FONT,
    fontSize: size,
    bold: o.bold !== false,
    color: o.color || C.white,
    align: o.align || "center",
    valign: "middle",
    margin: 0,
    shadow: o.shadow ? shadowOf(o.shadow) : undefined,
  });
}

function icon(name, colorKey, x, y, s) {
  cur.slide.addImage({ data: icons.get(name, colorKey), x, y, w: s, h: s, objectName: oname(), altText: "" });
}

// 원 안의 아이콘
function iconCircle(name, x, y, d, o = {}) {
  circle(x, y, d, { fill: o.fill || C.light, line: o.line });
  const s = d * (o.scale || 0.52);
  icon(name, o.color || "blue", x + (d - s) / 2, y + (d - s) / 2, s);
}

// 숫자 배지
function badge(n, x, y, d, o = {}) {
  cur.slide.addText(String(n), {
    x, y, w: d, h: d,
    objectName: oname(),
    shape: pres.ShapeType.ellipse,
    fill: { color: o.fill || C.blue },
    line: { type: "none" },
    fontFace: FONT, fontSize: o.size || 14, bold: true,
    color: o.color || C.white, align: "center", valign: "middle", margin: 0,
  });
}

// ── 스마트폰 목업 (도형 기반, 편집 가능)
// 반환: 화면 안쪽 콘텐츠 좌표
function phone(o = {}) {
  const x = o.x === undefined ? G.phoneX : o.x;
  const y = o.y === undefined ? G.phoneY : o.y;
  const w = G.phoneW, h = G.phoneH;
  const bz = 0.075;
  group("phone-frame", () => {
    rect(x, y, w, h, { fill: C.device, r: 0.38, shadow: "lg" });
    rect(x + bz, y + bz, w - bz * 2, h - bz * 2, { fill: o.bg || C.bg, r: 0.31 });
    // 다이내믹 아일랜드 · 상태바
    rect(x + w / 2 - 0.4, y + 0.17, 0.8, 0.2, { fill: C.device, r: 0.1 });
    text(o.time || "9:41", x + 0.27, y + 0.17, 0.5, 0.2, { size: 8, bold: true, valign: "middle", nocheck: true });
    rect(x + w - 0.52, y + 0.215, 0.24, 0.11, { line: C.navy, lw: 0.75, r: 0.03 });
    rect(x + w - 0.505, y + 0.23, 0.16, 0.08, { fill: C.navy, r: 0.015 });
    // 홈 인디케이터
    rect(x + w / 2 - 0.45, y + h - 0.2, 0.9, 0.045, { fill: C.navy, r: 0.022 });
    if (o.title) {
      icon("LuChevronLeft", "navy", x + 0.2, y + 0.555, 0.2);
      text(o.title, x + 0.5, y + 0.5, w - 1.0, 0.32, { size: 10.5, bold: true, align: "center", valign: "middle", nocheck: true });
    }
  });
  return {
    x, y, w, h,
    cx: x + 0.22, // 콘텐츠 좌측
    cw: w - 0.44, // 콘텐츠 폭
    top: y + 0.92, // 헤더 아래 시작
    bottom: y + h - 0.3,
  };
}

// 폰 아래 캡션
function phoneCaption(str) {
  text(str, G.phoneX - 0.5, G.phoneY + G.phoneH + 0.1, G.phoneW + 1.0, 0.2, {
    size: 10, color: C.gray, align: "center", valign: "middle", nocheck: true,
  });
}

// 상단 소제목
function eyebrow(str) {
  text(str, G.mx, G.eyebrowY, 7, 0.28, { size: 13, bold: true, color: C.blue, valign: "middle", cs: 0.5, nocheck: true });
}

// 번호 + 제목 + 설명 콜아웃 (좌측 열)
function callout(n, head, sub, y, o = {}) {
  const x = o.x === undefined ? G.mx : o.x;
  const w = o.w || G.leftW;
  badge(n, x, y + 0.03, 0.46, { size: 16, fill: o.badge || C.blue });
  text(head, x + 0.68, y, w - 0.68, 0.46, { size: 24, bold: true, valign: "middle" });
  if (sub) text(sub, x + 0.68, y + 0.5, w - 0.68, 0.36, { size: 17, color: C.gray, valign: "middle" });
}

module.exports = {
  init, use, group, rect, circle, line, text, pill, icon, iconCircle, badge,
  phone, phoneCaption, eyebrow, callout, emWidth, warnings,
};
