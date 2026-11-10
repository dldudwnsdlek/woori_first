// WON 첫집 메이트 — 본선 발표자료 빌드 (pptxgenjs)
// 실행: node build_deck.js && python3 postprocess.py
const path = require("path");
const fs = require("fs");
const pptxgen = require("pptxgenjs");
const { W, H, FONT, THEME, C, G } = require("./lib/theme");
const icons = require("./lib/icons");
const ui = require("./lib/ui");
const { rect, circle, line, text, pill, icon, iconCircle, badge, group, callout, eyebrow } = ui;

let SCRIPT = [];
try {
  SCRIPT = require("./content/script").SLIDES;
} catch (e) {
  console.warn("발표 대본(content/script.js)이 없어 발표자 노트를 비워 둡니다.");
}

const OUT = path.join(__dirname, "out");
fs.mkdirSync(OUT, { recursive: true });
const LOGO = path.join(__dirname, "assets", "won_logo.png");

// [아이콘, 색] — 사용 전 래스터화
const ICONS = [
  ["LuChevronLeft", "navy"], ["LuHouse", "white"], ["LuHouse", "blue"], ["LuPlus", "gray"],
  ["LuFileText", "blue"], ["LuCalendarDays", "blue"], ["LuWallet", "blue"], ["LuUnlink", "grayLight"],
  ["LuFileSearch", "blue"], ["LuListChecks", "blue"], ["LuSend", "blue"], ["LuRepeat", "blue"],
  ["LuArrowRight", "gray"], ["LuArrowRight", "blue"], ["LuArrowRight", "white"], ["LuChevronRight", "blue"], ["LuChevronRight", "grayLight"],
  ["LuShieldCheck", "blue"], ["LuShieldCheck", "white"], ["LuShieldCheck", "gray"],
  ["LuCheck", "white"], ["LuCheck", "green"], ["LuCheck", "blue"],
  ["LuTriangleAlert", "orange"], ["LuCircleAlert", "orange"], ["LuCircleAlert", "white"],
  ["LuLock", "white"], ["LuLock", "grayLight"], ["LuLock", "blue"], ["LuScanFace", "white"], ["LuScanFace", "blue"],
  ["LuCirclePause", "orange"], ["LuBell", "navy"], ["LuMenu", "navy"],
  ["LuLandmark", "blue"], ["LuBanknote", "blue"], ["LuPiggyBank", "blue"], ["LuKeyRound", "blue"],
  ["LuUser", "blue"], ["LuHeartHandshake", "blue"], ["LuUsers", "blue"],
  ["LuCamera", "blue"], ["LuScanText", "blue"], ["LuRoute", "blue"], ["LuRoute", "white"], ["LuCalendarCheck", "blue"],
  ["LuHandCoins", "blue"], ["LuBuilding2", "blue"], ["LuFileCheck", "blue"], ["LuUserCheck", "blue"],
  ["LuArrowLeftRight", "blue"], ["LuImage", "grayLight"], ["LuMic", "grayLight"],
];

const anims = {}; // slideNo → [[{g, fx, delay}]]
const A = (g, fx = "rise", delay = 0) => ({ g, fx, delay });

async function main() {
  await icons.prepare(ICONS);

  const pres = new pptxgen();
  pres.defineLayout({ name: "WIDE", width: W, height: H });
  pres.layout = "WIDE";
  pres.title = "WON 첫집 메이트 — 본선 발표";
  pres.author = "팀 WEN";
  pres.company = "WOORI × SSAFY 우리WON뱅킹 아이디어 경진대회";
  pres.subject = "주제 ③ 우리WON뱅킹 신규 서비스 제안";
  pres.theme = { headFontFace: FONT, bodyFontFace: FONT };
  ui.init(pres);

  // ───────────────────────── 레이아웃
  const footer = (numColor) => [
    { image: { x: G.mx, y: 6.9, w: 0.3, h: 0.302, path: LOGO } },
    { text: { text: "WON 첫집 메이트", options: { x: G.mx + 0.42, y: 6.9, w: 3, h: 0.3, fontFace: FONT, fontSize: 10.5, bold: true, color: C.gray, valign: "middle", margin: 0 } } },
  ];
  const num = (color) => ({ x: 12.03, y: 6.9, w: 0.5, h: 0.3, fontFace: FONT, fontSize: 10.5, color, align: "right", valign: "middle", margin: 0 });
  const titlePh = (w, h, size, y) => ({
    placeholder: {
      options: { name: "title", type: "title", x: G.mx, y: y === undefined ? G.titleY : y, w, h, fontFace: FONT, fontSize: size, bold: true, color: C.navy, align: "left", valign: "top", margin: 0, lineSpacing: Math.round(size * 1.2) },
      text: "",
    },
  });

  pres.defineSlideMaster({
    title: "COVER", background: { color: "FFFFFF" },
    objects: [{ rect: { x: G.panelX, y: 0, w: G.panelW, h: H, fill: { color: C.light } } }, titlePh(7.4, 1.15, 56, 2.78)],
  });
  pres.defineSlideMaster({
    title: "FULL", background: { color: "FFFFFF" },
    objects: [...footer(), titlePh(G.fullW, 0.7, 34)],
    slideNumber: num(C.gray),
  });
  pres.defineSlideMaster({
    title: "PHONE", background: { color: "FFFFFF" },
    objects: [{ rect: { x: G.panelX, y: 0, w: G.panelW, h: H, fill: { color: C.light } } }, ...footer(), titlePh(7.4, 1.3, 34)],
    slideNumber: num(C.gray),
  });
  pres.defineSlideMaster({
    title: "CLOSE", background: { color: "FFFFFF" },
    objects: [{ rect: { x: G.panelX, y: 0, w: G.panelW, h: H, fill: { color: C.blue } } }, ...footer(), titlePh(7.4, 1.3, 34)],
    slideNumber: num(C.white),
  });

  const SECTIONS = ["도입", "문제 상황", "서비스 기획", "서비스 UI", "기대효과 · 실현 가능성"];
  SECTIONS.forEach((t) => pres.addSection({ title: t }));
  const add = (no, masterName, sectionTitle, title) => {
    const s = pres.addSlide({ masterName, sectionTitle });
    ui.use(s, `S${no}`);
    s.addText(title, { placeholder: "title" });
    const sc = SCRIPT[no - 1];
    if (sc) s.addNotes(sc.notes);
    return s;
  };

  // 폰 내부 카드
  const pcard = (x, y, w, h, o = {}) => rect(x, y, w, h, { fill: o.fill || C.white, r: 0.12, line: o.line, lw: o.lw, shadow: o.shadow });

  // ═════════════════════════ 1. 도입
  {
    add(1, "COVER", "도입", "WON 첫집 메이트");
    // 좌측
    text("WOORI × SSAFY 우리WON뱅킹 아이디어 경진대회", G.mx, 1.62, 7, 0.3, { size: 15, color: C.gray, valign: "middle" });
    pill("주제 ③ 우리WON뱅킹 신규 서비스 제안", G.mx, 2.04, 3.5, 0.4, { fill: C.light, color: C.blue, size: 13 });
    text("“저 12월 20일에 이사 가요.”", G.mx, 4.12, G.leftW, 0.58, { size: 30, bold: true, color: C.blue, valign: "middle" });
    text("한마디면, 첫 집의 모든 순서를 AI가 챙깁니다.", G.mx, 4.76, G.leftW, 0.46, { size: 24, valign: "middle" });
    text([{ t: "팀 WEN", bold: true, color: C.navy }, { t: "     유지민 · 김승현 · 이영준", color: C.gray }], G.mx, 6.5, 6, 0.3, { size: 14, valign: "middle" });

    // 우측 폰 — 첫 대화
    const P = ui.phone({ title: "첫집 메이트", time: "10:22" });
    rect(P.x + 0.075, P.top - 0.06, P.w - 0.15, 0.26, { fill: C.light });
    text("확인·계산·예약을 도와요. 송금은 직접 승인해요.", P.x + 0.1, P.top - 0.06, P.w - 0.2, 0.26, { size: 7, color: C.gray, align: "center", valign: "middle", nocheck: true });
    text("10월 23일 금요일", P.cx, P.top + 0.3, P.cw, 0.18, { size: 7, color: C.grayLight, align: "center", valign: "middle" });

    group("s1-user", () => {
      pill("저 12월 20일에 이사 가요", P.cx + P.cw - 2.1, P.top + 0.58, 2.1, 0.38, { fill: C.blue, color: C.white, size: 10.5, r: 0.14 });
    });
    group("s1-ai", () => {
      const ax = P.cx, ay = P.top + 1.14;
      iconCircle("LuHouse", ax, ay, 0.3, { fill: C.blue, color: "white", scale: 0.55 });
      text("첫집 메이트", ax + 0.38, ay - 0.02, 1.5, 0.18, { size: 7.5, color: C.gray, valign: "middle" });
      const bx = ax + 0.38, bw = P.cw - 0.38;
      pcard(bx, ay + 0.2, bw, 0.6);
      text("축하해요! 이사일에 맞춰\n순서대로 챙길게요.", bx + 0.14, ay + 0.2, bw - 0.28, 0.6, { size: 9.5, valign: "middle", ls: 1.1 });
      const cy = ay + 0.9;
      pcard(bx, cy, bw, 1.98);
      const rows = ["계약 위험점검", "할 일 일정표", "잔금 설계 · 송금 준비", "입주 후 자동이체"];
      rows.forEach((t, i) => {
        const ry = cy + 0.12 + i * 0.34;
        badge(i + 1, bx + 0.14, ry + 0.04, 0.22, { size: 8, fill: i === 2 ? C.orange : C.blue });
        text(t, bx + 0.46, ry, bw - 0.6, 0.3, { size: 10, bold: true, valign: "middle" });
      });
      pill("계약서 사진 올리기", bx + 0.14, cy + 1.54, bw - 0.28, 0.32, { line: C.blue, color: C.blue, size: 9.5, r: 0.1 });
    });
    // 입력창
    circle(P.cx, P.bottom - 0.44, 0.32, { fill: C.white });
    icon("LuPlus", "gray", P.cx + 0.08, P.bottom - 0.36, 0.16);
    rect(P.cx + 0.4, P.bottom - 0.44, P.cw - 0.4, 0.32, { fill: C.white, r: 0.16 });
    text("메이트에게 물어보세요", P.cx + 0.54, P.bottom - 0.44, 1.6, 0.32, { size: 8, color: C.grayLight, valign: "middle" });
    ui.phoneCaption("컨셉 화면 · 첫 대화");
    anims[1] = [[A("s1-user")], [A("s1-ai")]];
  }
  // 표지 로고 (레이아웃이 아닌 슬라이드에 직접)
  pres.slides[0].addImage({ path: LOGO, x: G.mx, y: 0.7, w: 0.62, h: 0.624, altText: "우리WON뱅킹 로고" });

  // ═════════════════════════ 2. 문제 상황
  {
    add(2, "FULL", "문제 상황", "첫 집을 준비할 때, 확인할 것은 많고 순서는 복잡합니다");
    eyebrow("문제 상황");
    const cw = 3.51, gap = 0.6, y0 = 2.1, ch = 3.0;
    const cards = [
      ["LuFileSearch", "서류 확인", "서류 6종 · 발급처 5곳", "등기부등본, 건축물대장…\n기관마다 따로 확인합니다"],
      ["LuCalendarDays", "일정 관리", "마감이 제각각인 할 일", "대출 심사, 이사, 전입신고…\n하나를 놓치면 뒤가 밀립니다"],
      ["LuWallet", "잔금 준비", "하루에 겹치는 여러 자금", "보증금, 대출, 예금…\n순서가 어긋나면 위험합니다"],
    ];
    cards.forEach(([ic, label, fact, desc], i) => {
      const x = G.mx + i * (cw + gap);
      group(`s2-c${i + 1}`, () => {
        rect(x, y0, cw, ch, { fill: C.bg, r: G.r });
        iconCircle(ic, x + 0.35, y0 + 0.35, 0.72, { fill: C.white });
        text(label, x + 0.35, y0 + 1.25, cw - 0.7, 0.5, { size: 26, bold: true, valign: "middle" });
        text(fact, x + 0.35, y0 + 1.8, cw - 0.7, 0.4, { size: 19, bold: true, color: C.blue, valign: "middle" });
        text(desc, x + 0.3, y0 + 2.24, cw - 0.5, 0.62, { size: 15, color: C.gray, ls: 1.08 });
        if (i > 0) {
          // 끊어진 연결 표시
          const gx = x - gap / 2 - 0.2;
          circle(gx, y0 + ch / 2 - 0.2, 0.4, { fill: C.white, line: C.line, lw: 1 });
          icon("LuUnlink", "grayLight", gx + 0.1, y0 + ch / 2 - 0.1, 0.2);
        }
      });
    });
    group("s2-stat", () => {
      text("3.6년", G.mx, 5.48, 1.95, 0.8, { size: 44, bold: true, color: C.blue, valign: "middle" });
      text("전월세 가구의 평균 거주 기간", G.mx + 2.05, 5.5, 4.3, 0.36, { size: 18, bold: true, valign: "middle" });
      text("몇 년에 한 번이라, 이사는 매번 처음입니다", G.mx + 2.05, 5.88, 4.3, 0.32, { size: 16, color: C.gray, valign: "middle" });
      text("출처: 국토교통부 2024년도 주거실태조사", G.mx + 2.05, 6.26, 4.3, 0.22, { size: 11, color: C.grayLight, valign: "middle" });
    });
    group("s2-bridge", () => {
      pill("문제는 정보를 찾는 것에서 끝나지 않습니다", 7.35, 5.52, 5.183, 0.9, { fill: C.navy, color: C.white, size: 19, r: G.r });
    });
    anims[2] = [[A("s2-c1")], [A("s2-c2")], [A("s2-c3")], [A("s2-stat"), A("s2-bridge", "rise", 600)]];
  }

  // ═════════════════════════ 3. 기획 배경
  {
    add(3, "FULL", "서비스 기획", "정보 안내는 있지만, 실행까지 잇는 경험은 비어 있습니다");
    eyebrow("서비스 기획 · 배경");
    const x0 = 3.5, cw = 2.1, gap = 0.21;
    const cx = (i) => x0 + i * (cw + gap);
    const stages = [["LuFileSearch", "확인"], ["LuListChecks", "계획"], ["LuSend", "실행"], ["LuRepeat", "사후관리"]];
    stages.forEach(([ic, t], i) => {
      iconCircle(ic, cx(i) + cw / 2 - 0.27, 2.1, 0.54, { fill: C.light });
      text(t, cx(i), 2.7, cw, 0.4, { size: 20, bold: true, align: "center", valign: "middle" });
    });
    // 기존
    const yA = 3.35, hA = 0.9;
    group("s3-rowA", () => {
      text("기존 서비스", G.mx, yA + 0.08, 2.5, 0.42, { size: 20, bold: true, color: C.ink2, valign: "middle" });
      text("점검 · 체크리스트 안내", G.mx, yA + 0.5, 2.5, 0.28, { size: 13, color: C.gray, valign: "middle" });
      pill("위험 점검", cx(0), yA, cw, hA, { fill: C.blueSoft, color: C.blue, size: 18, r: G.r });
      pill("체크리스트 안내", cx(1), yA, cw, hA, { fill: C.grayBg, color: C.ink2, size: 18, r: G.r });
    });
    group("s3-gap", () => {
      pill("실행은 고객이 직접, 기관마다 따로", cx(2), yA, cw * 2 + gap, hA, { line: C.orange, lw: 1.5, dash: "dash", color: C.orangeDeep, size: 18, r: G.r });
    });
    // 제안
    const yB = 4.78, hB = 1.0;
    group("s3-rowB", () => {
      text("WON 첫집 메이트", G.mx, yB + 0.12, 2.5, 0.42, { size: 20, bold: true, color: C.blue, valign: "middle" });
      text("제안", G.mx, yB + 0.56, 2.5, 0.28, { size: 13, color: C.gray, valign: "middle" });
      const labels = ["위험점검", "맞춤 일정", "잔금 설계 ·\n조건부 이체", "입주 후 자동화"];
      labels.forEach((t, i) => {
        pill(t, cx(i), yB, cw, hB, { fill: i === 2 ? C.blueDeep : C.blue, color: C.white, size: 17, r: G.r, nocheck: true });
        if (i > 0) icon("LuChevronRight", "blue", cx(i) - gap / 2 - 0.1, yB + hB / 2 - 0.1, 0.2);
      });
      pill("은행이 연결하는 실행", cx(2) + 0.2, yB - 0.38, cw - 0.4, 0.3, { fill: C.orange, color: C.white, size: 12 });
    });
    group("s3-line", () => {
      text(
        [{ t: "차이는 " }, { t: "‘실행’", bold: true, color: C.blue }, { t: "입니다. 돈이 실제로 움직이는 단계는 은행이 연결할 수 있습니다." }],
        G.mx, 6.08, G.fullW, 0.46, { size: 22, valign: "middle" }
      );
    });
    anims[3] = [[A("s3-rowA"), A("s3-gap", "fade", 600)], [A("s3-rowB")], [A("s3-line")]];
  }

  // ═════════════════════════ 4. 서비스 기획 — 네 단계
  {
    add(4, "FULL", "서비스 기획", "한마디면, AI가 첫 집의 과정을 네 단계로 연결합니다");
    eyebrow("서비스 기획 · 네 단계 여정");
    pill("“저 12월 20일에 이사 가요.”", G.mx, 2.02, 4.5, 0.64, { fill: C.blue, color: C.white, size: 22, r: 0.2 });
    icon("LuArrowRight", "gray", G.mx + 4.72, 2.2, 0.28);
    text("AI가 할 일을 나누고, 은행 기능으로 연결합니다", G.mx + 5.16, 2.02, 6.4, 0.64, { size: 18, color: C.gray, valign: "middle" });

    const w = 2.55, w3 = 3.18, gap = 0.3, yN = 3.3, hN = 2.9, y3 = 3.1, h3 = 3.3;
    const xs = [G.mx, G.mx + w + gap, G.mx + 2 * (w + gap), G.mx + 2 * (w + gap) + w3 + gap];
    const normal = [
      [0, "LuFileSearch", "계약 전후", "집 위험점검", "계약서 사진으로\n위험 신호 확인"],
      [1, "LuCalendarCheck", "D-60 ~ D-day", "맞춤 일정 관리", "이사일에서 거꾸로\n할 일을 정리"],
      [3, "LuRepeat", "D+1 ~", "입주 후 자동화", "자동이체·마감 알림\n한 번에 연결"],
    ];
    normal.forEach(([i, ic, phase, title, desc]) => {
      const x = xs[i];
      group(`s4-${i + 1}`, () => {
        if (i > 0) icon("LuChevronRight", "grayLight", x - gap / 2 - 0.12, 4.63, 0.24);
        rect(x, yN, w, hN, { fill: C.bg, r: G.r });
        text(phase, x + 0.3, yN + 0.28, w - 1.1, 0.26, { size: 13, bold: true, color: C.blue, valign: "middle" });
        text(`0${i + 1}`, x + w - 0.8, yN + 0.28, 0.5, 0.26, { size: 13, bold: true, color: C.grayLight, align: "right", valign: "middle" });
        iconCircle(ic, x + 0.3, yN + 0.72, 0.64, { fill: C.white });
        text(title, x + 0.3, yN + 1.56, w - 0.5, 0.46, { size: 23, bold: true, valign: "middle" });
        text(desc, x + 0.3, yN + 2.08, w - 0.5, 0.62, { size: 16, color: C.gray, ls: 1.08 });
      });
    });
    {
      const x = xs[2];
      group("s4-3", () => {
        icon("LuChevronRight", "grayLight", x - gap / 2 - 0.12, 4.63, 0.24);
        rect(x, y3, w3, h3, { fill: C.blue, r: G.r, shadow: "sm" });
        text("D-14 ~ D-day", x + 0.32, y3 + 0.3, 1.6, 0.26, { size: 13, bold: true, color: C.blueSoft, valign: "middle" });
        pill("핵심", x + w3 - 1.02, y3 + 0.28, 0.7, 0.3, { fill: C.orange, color: C.white, size: 12 });
        iconCircle("LuRoute", x + 0.32, y3 + 0.76, 0.68, { fill: C.white });
        text("잔금 설계 ·\n안전한 이체", x + 0.32, y3 + 1.62, w3 - 0.64, 0.98, { size: 25, bold: true, color: C.white, ls: 1.08 });
        text("잔금 지도로 자금 점검,\n조건 확인 후 고객 승인", x + 0.32, y3 + 2.56, w3 - 0.5, 0.62, { size: 16, color: C.blueSoft, ls: 1.08 });
      });
    }
    anims[4] = [[A("s4-1"), A("s4-2", "rise", 400)], [A("s4-3")], [A("s4-4")]];
  }

  // ═════════════════════════ 5. UI ① 위험점검
  {
    add(5, "PHONE", "서비스 UI", "계약서 사진 한 장에서\n시작합니다");
    eyebrow("서비스 UI ①  위험점검");
    const cs = [
      ["사진 한 장으로 핵심 정보 추출", "매매가 · 잔금 · 잔금일을 읽고, 고객이 확인합니다"],
      ["위험 신호는 신호등으로", "주의 1건, 나머지 4개 항목은 안전"],
      ["주의 항목은 다음 단계에 반영", "매도인 근저당 → 잔금일에 ‘상환·말소 확인’ 추가"],
    ];
    cs.forEach(([h, sub], i) =>
      group(`s5-c${i + 1}`, () => {
        callout(i + 1, h, sub, 2.72 + i * 1.08, { badge: i === 2 ? C.orange : C.blue });
        if (i === 2) {
          text("※ 등기 · 체납 · 시세 조회는 기존 진단 기능 확장과 공공 데이터 연계를 전제로 한\n    구축 제안입니다. 최종 판단은 고객과 전문가 확인을 거칩니다.", G.mx, 6.04, G.leftW, 0.54, { size: 13, color: C.gray, ls: 1.15 });
        }
      })
    );

    const P = ui.phone({ title: "우리집 위험점검", time: "10:24" });
    let y = P.top + 0.02;
    // 올린 계약서
    pcard(P.cx, y, P.cw, 0.46);
    rect(P.cx + 0.12, y + 0.08, 0.26, 0.3, { fill: C.grayBg, r: 0.03 });
    line(P.cx + 0.17, y + 0.17, P.cx + 0.33, y + 0.17, { color: C.grayLight, w: 0.75 });
    line(P.cx + 0.17, y + 0.23, P.cx + 0.33, y + 0.23, { color: C.grayLight, w: 0.75 });
    line(P.cx + 0.17, y + 0.29, P.cx + 0.28, y + 0.29, { color: C.grayLight, w: 0.75 });
    text("부동산(아파트) 매매계약서", P.cx + 0.5, y + 0.06, 1.7, 0.18, { size: 8.5, bold: true, valign: "middle" });
    text("사진 1장 · 방금 올림", P.cx + 0.5, y + 0.24, 1.7, 0.16, { size: 7, color: C.gray, valign: "middle" });
    y += 0.54;
    // 읽은 내용
    group("s5-p1", () => {
      pcard(P.cx, y, P.cw, 1.3);
      icon("LuScanText", "blue", P.cx + 0.14, y + 0.11, 0.17);
      text("계약서에서 읽은 내용", P.cx + 0.38, y + 0.08, 1.8, 0.22, { size: 9, bold: true, valign: "middle" });
      const rows = [["매매가", "5억 원", C.navy], ["계약금", "5,000만 원 · 지급 완료", C.navy], ["잔금", "4억 5,000만 원", C.navy], ["잔금일", "12월 20일 (일)", C.orange]];
      rows.forEach(([k, v, col], i) => {
        const ry = y + 0.37 + i * 0.22;
        text(k, P.cx + 0.14, ry, 0.6, 0.2, { size: 8, color: C.gray, valign: "middle" });
        text(v, P.cx + 0.7, ry, P.cw - 0.84, 0.2, { size: 9, bold: true, color: col, align: "right", valign: "middle" });
      });
    });
    y += 1.38;
    // 신호등 + 목록
    group("s5-p2", () => {
      pcard(P.cx, y, P.cw, 0.6);
      rect(P.cx + 0.14, y + 0.18, 0.62, 0.24, { fill: C.device, r: 0.12 });
      circle(P.cx + 0.19, y + 0.23, 0.14, { fill: "3A4556" });
      circle(P.cx + 0.38, y + 0.23, 0.14, { fill: C.yellow });
      circle(P.cx + 0.57, y + 0.23, 0.14, { fill: "3A4556" });
      text([{ t: "주의 1건", color: C.orangeDeep, bold: true, br: true }, { t: "나머지 4개는 안전해요", bold: true }], P.cx + 0.9, y + 0.07, P.cw - 1.0, 0.46, { size: 9.5, valign: "middle", ls: 1.1 });
      const ly = y + 0.68;
      pcard(P.cx, ly, P.cw, 1.98);
      // 근저당 (주의)
      circle(P.cx + 0.14, ly + 0.165, 0.09, { fill: C.orange });
      text("근저당", P.cx + 0.32, ly + 0.1, 1.2, 0.22, { size: 9.5, bold: true, valign: "middle" });
      pill("주의", P.cx + P.cw - 0.56, ly + 0.11, 0.42, 0.2, { fill: C.orangeSoft, color: C.orangeDeep, size: 7.5 });
      rect(P.cx + 0.14, ly + 0.38, P.cw - 0.28, 0.5, { fill: C.orangeSoft, r: 0.06 });
      text("매도인 대출이 남아 있어요.\n잔금일 순서에 ‘상환·말소 확인’을 넣었어요.", P.cx + 0.22, ly + 0.38, P.cw - 0.44, 0.5, { size: 7.2, color: C.ink2, valign: "middle", ls: 1.12 });
      const safe = [["압류 · 가압류", "안전"], ["매도인 세금 체납", "안전"], ["시세 대비 가격", "적정"], ["건축물대장", "안전"]];
      safe.forEach(([k, v], i) => {
        const ry = ly + 0.97 + i * 0.245;
        circle(P.cx + 0.14, ry + 0.065, 0.09, { fill: C.green });
        text(k, P.cx + 0.32, ry, 1.5, 0.22, { size: 9, valign: "middle" });
        pill(v, P.cx + P.cw - 0.56, ry + 0.01, 0.42, 0.2, { fill: C.greenSoft, color: C.greenDeep, size: 7.5 });
      });
    });
    group("s5-p3", () => {
      const ly = y + 0.68;
      rect(P.cx - 0.03, ly + 0.03, P.cw + 0.06, 0.92, { line: C.orange, lw: 2, r: 0.12 });
    });
    ui.phoneCaption("컨셉 화면 · 예시 데이터");
    anims[5] = [[A("s5-c1"), A("s5-p1", "fade")], [A("s5-c2"), A("s5-p2", "fade")], [A("s5-c3"), A("s5-p3", "fade")]];
  }

  // ═════════════════════════ 6. UI ② 일정 관리
  {
    add(6, "PHONE", "서비스 UI", "이사 날짜만 알면,\n해야 할 일을 역산합니다");
    eyebrow("서비스 UI ②  맞춤형 일정 관리");
    group("s6-c1", () => {
      callout(1, "휴일이면, 영업일로 먼저 조정", null, 2.6);
      const tx = G.mx + 0.68, ty = 3.16, tw = 2.45, th = 1.0;
      rect(tx, ty, tw, th, { fill: C.grayBg, r: G.r });
      text("12.20", tx + 0.28, ty + 0.08, tw - 0.4, 0.54, { size: 32, bold: true, color: C.grayLight, valign: "middle" });
      text("일요일 · 은행 휴무", tx + 0.28, ty + 0.62, tw - 0.4, 0.28, { size: 14, color: C.gray, valign: "middle" });
      icon("LuArrowRight", "blue", tx + tw + 0.16, ty + th / 2 - 0.15, 0.3);
      const t2 = tx + tw + 0.62;
      rect(t2, ty, tw, th, { fill: C.blue, r: G.r });
      text("12.18", t2 + 0.28, ty + 0.08, tw - 0.4, 0.54, { size: 32, bold: true, color: C.white, valign: "middle" });
      text("금요일 · 잔금일로 조정", t2 + 0.28, ty + 0.62, tw - 0.4, 0.28, { size: 14, color: C.white, valign: "middle" });
    });
    group("s6-c2", () => callout(2, "D-day에서 거꾸로 만든 체크리스트", "내 상황에 맞는 일만, 마감 전에 알려줍니다", 4.42));
    group("s6-c3", () => {
      callout(3, "지금 할 일 하나만 강조", "끝난 일은 체크, 남은 일은 D-45처럼 날짜로 표시", 5.4);
      text("※ 일정은 예시이며, 실제 계약 · 대출 조건에 따라 달라집니다.", G.mx, 6.46, G.leftW, 0.26, { size: 13, color: C.gray, valign: "middle" });
    });

    const P = ui.phone({ title: "할 일 일정표", time: "10:26" });
    let y = P.top + 0.02;
    rect(P.cx, y, P.cw, 1.3, { fill: C.blue, r: 0.12 });
    text("잔금 · 이사까지", P.cx + 0.16, y + 0.1, 1.5, 0.18, { size: 8, color: C.blueSoft, valign: "middle" });
    text("D-56", P.cx + 0.16, y + 0.27, 1.8, 0.5, { size: 30, bold: true, color: C.white, valign: "middle" });
    text("2026년 12월 18일 금요일", P.cx + 0.16, y + 0.76, 2.0, 0.2, { size: 9, bold: true, color: C.white, valign: "middle" });
    group("s6-p1", () => {
      pill("20일(일) → 18일(금) 조정 · 매도인 동의 완료", P.cx + 0.12, y + 1.0, P.cw - 0.24, 0.22, { fill: C.blueDeep, color: C.white, size: 7.2, r: 0.06, nocheck: true });
    });
    y += 1.38;
    group("s6-p2", () => {
      text([{ t: "할 일 11개 중 " }, { t: "3개", color: C.blue }, { t: " 끝났어요" }], P.cx + 0.04, y, 1.8, 0.22, { size: 8.5, bold: true, valign: "middle" });
      rect(P.cx + 0.04, y + 0.25, P.cw - 0.08, 0.05, { fill: C.line, r: 0.025 });
      rect(P.cx + 0.04, y + 0.25, (P.cw - 0.08) * (3 / 11), 0.05, { fill: C.blue, r: 0.025 });
      const ly = y + 0.4;
      pcard(P.cx, ly, P.cw, 2.66);
      const done = ["계약서 등록 · 위험점검", "잔금일 18일로 조정", "생애최초 대상 확인"];
      done.forEach((t, i) => {
        const ry = ly + 0.1 + i * 0.27;
        circle(P.cx + 0.14, ry + 0.03, 0.17, { fill: C.blue });
        icon("LuCheck", "white", P.cx + 0.175, ry + 0.065, 0.1);
        text(t, P.cx + 0.42, ry, 1.5, 0.23, { size: 8.5, color: C.gray, valign: "middle" });
        text("10.23", P.cx + P.cw - 0.6, ry, 0.46, 0.23, { size: 7.5, color: C.grayLight, align: "right", valign: "middle" });
      });
      // 다음 할 일
      const ny = ly + 0.95;
      circle(P.cx + 0.14, ny + 0.115, 0.17, { line: C.blue, lw: 1.25, fill: C.white });
      text("디딤돌대출 신청 서류 준비", P.cx + 0.42, ny + 0.03, 1.55, 0.2, { size: 9, bold: true, valign: "middle", nocheck: true });
      text("서류 4개 · 마감 3일 전 알림", P.cx + 0.42, ny + 0.22, 1.55, 0.16, { size: 7, color: C.gray, valign: "middle" });
      text([{ t: "D-45", color: C.orange, bold: true, br: true }, { t: "11.3 (화)", color: C.grayLight, bold: false, size: 7 }], P.cx + P.cw - 0.7, ny + 0.02, 0.56, 0.38, { size: 9, align: "right", valign: "middle" });
      const todo = [["이사업체 견적 받기", "D-30"], ["잔금 지도 확정", "D-14"], ["보증금 반환 시간 정하기", "D-7"], ["잔금 · 이사", "D-day"]];
      todo.forEach(([t, d], i) => {
        const ry = ly + 1.46 + i * 0.285;
        circle(P.cx + 0.14, ry + 0.03, 0.17, { line: C.grayLight, lw: 1, fill: C.white });
        text(t, P.cx + 0.42, ry, 1.5, 0.23, { size: 8.5, valign: "middle" });
        text(d, P.cx + P.cw - 0.6, ry, 0.46, 0.23, { size: 8, bold: true, color: C.gray, align: "right", valign: "middle" });
      });
    });
    group("s6-p3", () => {
      const ny = y + 0.4 + 0.95;
      rect(P.cx + 0.05, ny - 0.03, P.cw - 0.1, 0.46, { line: C.blue, lw: 1.75, r: 0.08 });
    });
    ui.phoneCaption("컨셉 화면 · 예시 일정");
    anims[6] = [[A("s6-c1"), A("s6-p1", "fade")], [A("s6-c2"), A("s6-p2", "fade")], [A("s6-c3"), A("s6-p3", "fade")]];
  }

  // ═════════════════════════ 7. UI ③ 잔금 지도
  {
    add(7, "PHONE", "서비스 UI", "잔금일에 돈이 어디서 오고\n어디로 가는지, 한눈에 보여줍니다");
    eyebrow("서비스 UI ③  잔금 지도");
    const sx = G.mx, sw = 2.95, sh = 0.8, sg = 0.14, sy = 2.62;
    const src = [["LuHandCoins", "돌려받을 보증금", "2억 원"], ["LuLandmark", "디딤돌대출 (생애최초)", "2억 4,000만 원"], ["LuPiggyBank", "내 예금", "1,000만 원"]];
    group("s7-src", () => {
      src.forEach(([ic, k, v], i) => {
        const y = sy + i * (sh + sg);
        rect(sx, y, sw, sh, { fill: C.bg, r: G.r });
        iconCircle(ic, sx + 0.18, y + 0.16, 0.48, { fill: C.white });
        text(k, sx + 0.82, y + 0.1, sw - 0.95, 0.26, { size: 13, color: C.gray, valign: "middle" });
        text(v, sx + 0.8, y + 0.34, sw - 0.85, 0.38, { size: 19, bold: true, valign: "middle" });
      });
    });
    const yEnd = sy + 3 * sh + 2 * sg; // 5.3
    const bx = sx + sw + 0.38, mid = (sy + yEnd) / 2;
    const rx = 4.62, rw = G.mx + G.leftW - rx;
    group("s7-sum", () => {
      src.forEach((_, i) => {
        const y = sy + i * (sh + sg) + sh / 2;
        line(sx + sw + 0.06, y, bx, y, { color: C.blue, w: 1.5 });
      });
      line(bx, sy + sh / 2, bx, yEnd - sh / 2, { color: C.blue, w: 1.5 });
      line(bx, mid, rx - 0.06, mid, { color: C.blue, w: 1.5, arrow: true });
      rect(rx, sy, rw, yEnd - sy, { fill: C.blue, r: G.r, shadow: "sm" });
      text("매도인에게 보낼 잔금", rx + 0.3, sy + 0.34, rw - 0.6, 0.3, { size: 15, bold: true, color: C.blueSoft, valign: "middle" });
      text("4억 5,000만 원", rx + 0.3, sy + 0.72, rw - 0.4, 0.6, { size: 25, bold: true, color: C.white, valign: "middle" });
      text("매매가 5억 - 계약금 5,000만", rx + 0.3, sy + 1.36, rw - 0.5, 0.28, { size: 13, color: C.blueSoft, valign: "middle" });
      rect(rx + 0.3, sy + 1.9, rw - 0.6, 0.46, { fill: C.white, r: 0.23 });
      circle(rx + 0.42, sy + 2.01, 0.24, { fill: C.green });
      icon("LuCheck", "white", rx + 0.47, sy + 2.06, 0.14);
      text("준비 자금과 일치", rx + 0.76, sy + 1.9, rw - 1.2, 0.46, { size: 14, bold: true, color: C.greenDeep, valign: "middle" });
    });
    const by = 5.56, bh = 1.02;
    group("s7-short", () => {
      rect(sx, by, G.leftW, bh, { fill: C.orangeSoft, r: G.r });
      iconCircle("LuCircleAlert", sx + 0.24, by + 0.27, 0.48, { fill: C.orange, color: "white", scale: 0.56 });
      text("부대비용 620만 원이 모자라요", sx + 0.92, by + 0.14, 4.0, 0.42, { size: 21, bold: true, color: C.orangeDeep, valign: "middle" });
      text("취득세 등 350만 + 중개보수 200만 + 법무사 70만", sx + 0.92, by + 0.58, 4.1, 0.28, { size: 12, color: C.ink2, valign: "middle" });
    });
    group("s7-fix", () => {
      const fx = sx + G.leftW - 1.92;
      text("2주 전 미리 알림", fx, by + 0.14, 1.72, 0.26, { size: 12.5, bold: true, color: C.orangeDeep, align: "center", valign: "middle" });
      pill("채우고  준비 완료", fx, by + 0.46, 1.72, 0.4, { fill: C.green, color: C.white, size: 13.5 });
    });

    const P = ui.phone({ title: "잔금 지도", time: "9:12" });
    let y = P.top + 0.02;
    pcard(P.cx, y, P.cw, 2.02);
    text("12월 18일까지 필요한 돈", P.cx + 0.16, y + 0.1, 2.0, 0.18, { size: 8, color: C.gray, valign: "middle" });
    text("4억 5,620만 원", P.cx + 0.16, y + 0.28, 2.2, 0.38, { size: 20, bold: true, valign: "middle" });
    text("잔금 4억 5,000만 + 부대비용 620만", P.cx + 0.16, y + 0.66, 2.2, 0.16, { size: 7.5, color: C.gray, valign: "middle" });
    const bw = P.cw - 0.32, bx0 = P.cx + 0.16, byy = y + 0.9;
    rect(bx0, byy, bw * 0.43, 0.11, { fill: C.blue, r: 0.03 });
    rect(bx0 + bw * 0.43 + 0.015, byy, bw * 0.5, 0.11, { fill: C.cyan });
    rect(bx0 + bw * 0.93 + 0.03, byy, bw * 0.035, 0.11, { fill: C.navy });
    rect(bx0 + bw * 0.965 + 0.045, byy, bw * 0.035 - 0.045, 0.11, { fill: C.orange, r: 0.03 });
    const lg = [[C.blue, "돌려받을 보증금", "2억"], [C.cyan, "디딤돌대출 (생애최초)", "2억 4,000만"], [C.navy, "내 예금", "1,000만"]];
    lg.forEach(([col, k, v], i) => {
      const ry = y + 1.12 + i * 0.28;
      circle(P.cx + 0.17, ry + 0.075, 0.09, { fill: col });
      text(k, P.cx + 0.36, ry, 1.4, 0.24, { size: 8.5, valign: "middle" });
      text(v, P.cx + P.cw - 1.2, ry, 1.04, 0.24, { size: 9.5, bold: true, align: "right", valign: "middle" });
    });
    y += 2.1;
    group("s7-palert", () => {
      rect(P.cx, y, P.cw, 1.02, { fill: C.orangeSoft, r: 0.12, line: C.orange, lw: 1 });
      iconCircle("LuCircleAlert", P.cx + 0.14, y + 0.13, 0.24, { fill: C.orange, color: "white", scale: 0.6 });
      text("620만 원이 모자라요", P.cx + 0.46, y + 0.1, 1.9, 0.3, { size: 12, bold: true, color: C.orangeDeep, valign: "middle" });
      text("부대비용까지 넣어 계산했어요.\n잔금일까지 2주 남았어요.", P.cx + 0.46, y + 0.42, 1.9, 0.3, { size: 7.5, color: C.ink2, ls: 1.1 });
      text("채우는 방법 보기 ›", P.cx + 0.46, y + 0.76, 1.6, 0.18, { size: 8, bold: true, color: C.orangeDeep, valign: "middle" });
    });
    y += 1.1;
    pcard(P.cx, y, P.cw, 1.16);
    text("따로 드는 돈", P.cx + 0.16, y + 0.09, 1.2, 0.22, { size: 8.5, bold: true, valign: "middle" });
    text("620만 원", P.cx + P.cw - 1.0, y + 0.09, 0.84, 0.22, { size: 8.5, bold: true, color: C.gray, align: "right", valign: "middle" });
    [["취득세 등", "350만"], ["중개보수", "200만"], ["법무사 비용", "70만"]].forEach(([k, v], i) => {
      const ry = y + 0.38 + i * 0.24;
      text(k, P.cx + 0.16, ry, 1.2, 0.22, { size: 8.5, color: C.ink2, valign: "middle" });
      text(v, P.cx + P.cw - 1.0, ry, 0.84, 0.22, { size: 9, bold: true, align: "right", valign: "middle" });
    });
    ui.phoneCaption("컨셉 화면 · 모든 금액은 예시");
    anims[7] = [[A("s7-src")], [A("s7-sum")], [A("s7-short"), A("s7-palert", "fade")], [A("s7-fix", "fade")]];
  }

  // ═════════════════════════ 8. UI ④ 잔금일 돈 길 내비게이션
  {
    add(8, "PHONE", "서비스 UI", "조건이 확인돼야 이체가 열리고,\n마지막 송금은 고객이 승인합니다");
    eyebrow("서비스 UI ④  잔금일 돈 길 내비게이션");
    const steps = [
      ["08:30", "등기부 최종 확인", "새 근저당 · 압류가 없는지 조회", "자동 조회"],
      ["09:00", "기존 보증금 수령", "2억 원 → 안심금고 입금 확인", "입금 확인"],
      ["10:00", "대출 실행", "2억 4,000만 원 · 매도인 대출 상환 확인", "은행 실행"],
      ["10:30", "잔금 송금", "안심금고 2억 1,000만 원 → 매도인", "고객 승인"],
    ];
    const ty = 2.58, rh = 0.76, nx = 2.12, nd = 0.3;
    const drawStep = (i, on) => {
      const [tm, title, desc, chip] = steps[i];
      const y = ty + i * rh;
      const last = i === 3;
      const main = last ? C.orange : C.blue;
      text(tm, G.mx, y, 1.2, 0.42, { size: 24, bold: true, color: on ? main : C.grayLight, valign: "middle" });
      if (i < 3) line(nx + nd / 2, y + 0.06 + nd, nx + nd / 2, y + rh + 0.06, { color: on ? C.blue : C.line, w: 2 });
      if (on) {
        circle(nx, y + 0.06, nd, { fill: main });
        icon(last ? "LuScanFace" : "LuCheck", "white", nx + 0.07, y + 0.13, 0.16);
      } else {
        circle(nx, y + 0.06, nd, { fill: C.white, line: C.grayLight, lw: 1.5 });
      }
      text(title, 2.66, y, 3.8, 0.42, { size: 22, bold: true, color: on ? C.navy : C.grayLight, valign: "middle" });
      text(desc, 2.66, y + 0.41, 3.9, 0.28, { size: 15, color: on ? C.ink2 : C.grayLight, valign: "middle" });
      if (on) pill(chip, 6.6, y + 0.03, 1.25, 0.38, { fill: last ? C.orange : C.light, color: last ? C.white : C.blue, size: 13 });
      else pill(chip, 6.6, y + 0.03, 1.25, 0.38, { line: C.line, color: C.grayLight, size: 13 });
    };
    for (let i = 0; i < 4; i++) group(`s8-L${i + 1}`, () => drawStep(i, true));

    group("s8-safe", () => {
      icon("LuCirclePause", "orange", G.mx, 5.725, 0.3);
      text("하나라도 어긋나면 모든 이체를 멈추고 고객에게 알립니다", G.mx + 0.42, 5.7, 6.6, 0.35, { size: 17, bold: true, color: C.orangeDeep, valign: "middle" });
    });
    group("s8-feat", () => {
      const fw = 2.2, fg = (G.leftW - fw * 3) / 2;
      ["잔금 안심금고", "조건부 이체 예약", "원터치 승인"].forEach((t, i) => {
        const x = G.mx + i * (fw + fg);
        rect(x, 6.2, fw, 0.46, { fill: C.light, r: 0.23 });
        badge(i + 1, x + 0.08, 6.28, 0.3, { size: 11 });
        text(t, x + 0.46, 6.2, fw - 0.5, 0.46, { size: 15, bold: true, color: C.blue, valign: "middle" });
      });
    });

    // 폰
    const P = ui.phone({ title: "잔금일 돈 길", time: "10:31" });
    let y = P.top;
    text("12월 18일 금요일 · 잔금일", P.cx + 0.04, y, 2.2, 0.18, { size: 8, bold: true, color: C.blue, valign: "middle" });
    text("조건이 확인되면 다음이 열려요", P.cx + 0.04, y + 0.19, 2.5, 0.24, { size: 11, bold: true, valign: "middle" });
    const segW = (P.cw - 0.08 - 0.09) / 4;
    // 안심금고
    const vy = y + 0.64;
    pcard(P.cx, vy, P.cw, 0.56, { line: C.blueSoft, lw: 1 });
    rect(P.cx + 0.12, vy + 0.12, 0.32, 0.32, { fill: C.blue, r: 0.08 });
    icon("LuLock", "white", P.cx + 0.2, vy + 0.2, 0.16);
    text("잔금 안심금고", P.cx + 0.54, vy + 0.1, 1.1, 0.2, { size: 9.5, bold: true, valign: "middle" });
    text("송금 전까지 보관", P.cx + 0.54, vy + 0.3, 1.0, 0.16, { size: 6.8, color: C.gray, valign: "middle", nocheck: true });
    // 단계 카드
    const ly = vy + 0.66, lh = 3.58;
    pcard(P.cx, ly, P.cw, lh);
    const prow = [
      ["등기부 최종 확인", "새 근저당 · 압류 없음", "08:31"],
      ["보증금 2억 원 입금", "지금 집주인 → 안심금고", "09:07"],
      ["대출 2억 4,000만 원 실행", "매도인 대출 2억 상환 · 4,000만 지급", "10:04"],
    ];
    const prh = 0.44, px = P.cx + 0.14, pnd = 0.18;
    const drawRow = (i, on) => {
      const [t, d, tm] = prow[i];
      const ry = ly + 0.1 + i * prh;
      line(px + pnd / 2, ry + 0.04 + pnd, px + pnd / 2, ry + prh + 0.04, { color: on ? C.green : C.line, w: 1.25 });
      if (on) {
        circle(px, ry + 0.04, pnd, { fill: C.green });
        icon("LuCheck", "white", px + 0.04, ry + 0.08, 0.1);
      } else circle(px, ry + 0.04, pnd, { fill: C.white, line: C.grayLight, lw: 1 });
      text(t, P.cx + 0.42, ry, 1.6, 0.21, { size: 9, bold: true, color: on ? C.navy : C.grayLight, valign: "middle" });
      text(d, P.cx + 0.42, ry + 0.2, 1.9, 0.16, { size: 6.8, color: on ? C.gray : C.grayLight, valign: "middle", nocheck: true });
      text(on ? tm : "대기", P.cx + P.cw - 0.58, ry, 0.46, 0.21, { size: 7.5, bold: true, color: on ? C.greenDeep : C.grayLight, align: "right", valign: "middle" });
    };
    const sy4 = ly + 0.1 + 3 * prh;
    const ay = sy4 + 0.3, ah = 1.5;
    // 하단 안내
    icon("LuShieldCheck", "gray", P.cx + 0.2, ly + lh - 0.235, 0.13);
    text("조건이 하나라도 어긋나면 모든 이체가 멈춰요", P.cx + 0.38, ly + lh - 0.27, P.cw - 0.5, 0.2, { size: 6.8, color: C.gray, valign: "middle", nocheck: true });

    for (let i = 0; i < 3; i++) {
      group(`s8-P${i + 1}`, () => {
        drawRow(i, true);
        rect(P.cx + 0.04 + i * (segW + 0.03), y + 0.5, segW, 0.05, { fill: C.green, r: 0.025 });
        if (i === 1) {
          text("2억 1,000만", P.cx + P.cw - 1.15, vy + 0.1, 1.03, 0.36, { size: 11, bold: true, color: C.blue, align: "right", valign: "middle" });
        }
      });
    }
    group("s8-P4", () => {
      rect(P.cx + 0.04 + 3 * (segW + 0.03), y + 0.5, segW, 0.05, { fill: C.orange, r: 0.025 });
      circle(px, sy4 + 0.04, pnd, { fill: C.orange });
      text("잔금 보내기", P.cx + 0.42, sy4, 1.4, 0.21, { size: 9, bold: true, valign: "middle" });
      text("승인 대기", P.cx + P.cw - 0.7, sy4, 0.58, 0.21, { size: 7.5, bold: true, color: C.orangeDeep, align: "right", valign: "middle" });
      rect(P.cx + 0.14, ay, P.cw - 0.28, ah, { fill: C.white, line: C.orange, lw: 1.5, r: 0.1, shadow: "sm" });
      icon("LuCheck", "green", P.cx + 0.26, ay + 0.125, 0.13);
      text("앞 단계 3개 모두 확인됨", P.cx + 0.44, ay + 0.09, 1.7, 0.2, { size: 7.5, bold: true, color: C.greenDeep, valign: "middle" });
      text("받는 분", P.cx + 0.26, ay + 0.34, 0.5, 0.18, { size: 7.2, color: C.gray, valign: "middle" });
      text("박○○ (매도인) · 우리 ···7310", P.cx + 0.76, ay + 0.34, P.cw - 1.02, 0.18, { size: 7.2, bold: true, align: "right", valign: "middle", nocheck: true });
      text("2억 1,000만 원", P.cx + 0.26, ay + 0.54, P.cw - 0.52, 0.36, { size: 17, bold: true, valign: "middle" });
      rect(P.cx + 0.26, ay + 0.94, P.cw - 0.52, 0.36, { fill: C.orange, r: 0.09 });
      icon("LuScanFace", "white", P.cx + 0.5, ay + 1.04, 0.16);
      text("얼굴 인증하고 보내기", P.cx + 0.26, ay + 0.94, P.cw - 0.52, 0.36, { size: 10, bold: true, color: C.white, align: "center", valign: "middle" });
      text("잠시 멈추기", P.cx + 0.26, ay + 1.32, P.cw - 0.52, 0.16, { size: 7, color: C.gray, align: "center", valign: "middle" });
    });
    group("s8-badges", () => {
      const bxx = P.x - 0.42;
      badge(1, bxx, vy + 0.13, 0.3, { size: 11 });
      badge(2, bxx, ly + 0.5, 0.3, { size: 11 });
      badge(3, bxx, ay + 0.97, 0.3, { size: 11 });
    });
    ui.phoneCaption("컨셉 화면 · 예시 금액 · 실제 거래 아님");
    anims[8] = [
      [A("s8-L1", "fade"), A("s8-P1", "fade")],
      [A("s8-L2", "fade"), A("s8-P2", "fade")],
      [A("s8-L3", "fade"), A("s8-P3", "fade")],
      [A("s8-L4", "fade"), A("s8-P4", "fade", 300)],
      [A("s8-safe")],
      [A("s8-feat"), A("s8-badges", "fade")],
    ];
  }

  // ═════════════════════════ 9. 입주 후 + 기대효과
  {
    add(9, "PHONE", "기대효과 · 실현 가능성", "첫 집의 첫날이, 우리은행과의\n새로운 금융 관계가 됩니다");
    eyebrow("입주 후 경험과 기대효과");
    const eff = [
      ["LuUser", "고객", "복잡한 절차를 순서대로", "처음이어도 놓치는 일 없이 잔금일까지 준비"],
      ["LuLandmark", "은행", "주거래 관계의 시작점", "대출 · 주거래 계좌 · 자동이체를 함께 연결할 기회"],
      ["LuHeartHandshake", "사회", "주거 금융 사고 위험을 줄이는 데 기여", "경험이 적은 청년층의 실수 가능성을 낮춥니다"],
    ];
    eff.forEach(([ic, tag, h, sub], i) => {
      const y = 2.62 + i * 1.26;
      group(`s9-e${i + 1}`, () => {
        rect(G.mx, y, G.leftW, 1.12, { fill: C.bg, r: G.r });
        iconCircle(ic, G.mx + 0.26, y + 0.28, 0.56, { fill: C.white });
        text(tag, G.mx + 1.02, y + 0.12, 0.8, 0.22, { size: 13, bold: true, color: C.blue, valign: "middle" });
        text(h, G.mx + 1.02, y + 0.34, G.leftW - 1.3, 0.42, { size: 22, bold: true, valign: "middle" });
        text(sub, G.mx + 1.02, y + 0.75, G.leftW - 1.3, 0.28, { size: 16, color: C.gray, valign: "middle" });
        if (i === 2) {
          text("※ 기대효과는 정성적 방향이며, 정량 효과는 시범 운영으로 검증이 필요합니다.", G.mx, 6.46, G.leftW, 0.26, { size: 13, color: C.gray, valign: "middle" });
        }
      });
    });

    const P = ui.phone({ time: "8:40" });
    text("우리WON뱅킹", P.cx + 0.02, P.y + 0.52, 1.6, 0.3, { size: 12, bold: true, color: C.blue, valign: "middle" });
    icon("LuBell", "navy", P.cx + P.cw - 0.5, P.y + 0.58, 0.18);
    icon("LuMenu", "navy", P.cx + P.cw - 0.2, P.y + 0.58, 0.18);
    let y = P.top + 0.02;
    pcard(P.cx, y, P.cw, 0.98);
    text("우리 WON통장  1002-***-**4821", P.cx + 0.16, y + 0.09, 2.2, 0.18, { size: 7.5, color: C.gray, valign: "middle" });
    text("3,240,500원", P.cx + 0.16, y + 0.27, 2.0, 0.32, { size: 15, bold: true, valign: "middle" });
    pill("내역", P.cx + 0.16, y + 0.64, (P.cw - 0.4) / 2, 0.24, { fill: C.grayBg, color: C.ink2, size: 8, r: 0.06 });
    pill("이체", P.cx + 0.24 + (P.cw - 0.4) / 2, y + 0.64, (P.cw - 0.4) / 2, 0.24, { fill: C.blue, color: C.white, size: 8, r: 0.06 });
    y += 1.08;
    // 첫집 메이트 카드
    const mh = 2.86;
    pcard(P.cx, y, P.cw, mh, { line: C.blue, lw: 1.5, shadow: "sm" });
    icon("LuHouse", "blue", P.cx + 0.16, y + 0.13, 0.16);
    text("첫집 메이트", P.cx + 0.38, y + 0.1, 1.2, 0.22, { size: 9, bold: true, color: C.blue, valign: "middle" });
    pill("입주 D+1", P.cx + P.cw - 0.78, y + 0.1, 0.64, 0.22, { fill: C.blue, color: C.white, size: 7.5 });
    text("새 집 자동이체 3건을\n연결했어요", P.cx + 0.16, y + 0.4, P.cw - 0.32, 0.5, { size: 12.5, bold: true, ls: 1.1 });
    [["디딤돌대출 원리금", "매월 18일"], ["아파트 관리비", "매월 25일"], ["전기 · 도시가스", "매월 말일"]].forEach(([k, v], i) => {
      const ry = y + 0.98 + i * 0.27;
      circle(P.cx + 0.16, ry + 0.035, 0.16, { fill: C.green });
      icon("LuCheck", "white", P.cx + 0.195, ry + 0.07, 0.09);
      text(k, P.cx + 0.42, ry, 1.3, 0.23, { size: 9, bold: true, valign: "middle" });
      text(v, P.cx + P.cw - 0.9, ry, 0.74, 0.23, { size: 8, color: C.gray, align: "right", valign: "middle" });
    });
    line(P.cx + 0.16, y + 1.88, P.cx + P.cw - 0.16, y + 1.88, { color: C.line, w: 0.75 });
    text("다음 할 일", P.cx + 0.16, y + 1.95, 1.2, 0.2, { size: 7.5, bold: true, color: C.gray, valign: "middle" });
    [["전입신고", "오늘", C.orange], ["취득세 신고 · 2027.2.16까지", "D-59", C.gray]].forEach(([k, v, col], i) => {
      const ry = y + 2.2 + i * 0.28;
      circle(P.cx + 0.16, ry + 0.035, 0.16, { line: C.grayLight, lw: 1, fill: C.white });
      text(k, P.cx + 0.42, ry, 1.6, 0.23, { size: 8.5, valign: "middle" });
      text(v, P.cx + P.cw - 0.66, ry, 0.5, 0.23, { size: 8, bold: true, color: col, align: "right", valign: "middle" });
    });
    y += mh + 0.1;
    pcard(P.cx, y, P.cw, 0.84);
    const q = [["LuShieldCheck", "전세지킴이"], ["LuBuilding2", "주택대출"], ["LuBell", "등기변동"], ["LuArrowLeftRight", "자동이체"]];
    q.forEach(([ic, t], i) => {
      const qw = P.cw / 4, qx = P.cx + i * qw;
      circle(qx + qw / 2 - 0.17, y + 0.12, 0.34, { fill: C.light });
      icon(ic === "LuBell" ? "LuKeyRound" : ic, "blue", qx + qw / 2 - 0.085, y + 0.205, 0.17);
      text(t, qx, y + 0.52, qw, 0.2, { size: 7, color: C.ink2, align: "center", valign: "middle" });
    });
    ui.phoneCaption("컨셉 화면 · 우리WON뱅킹 홈");
    anims[9] = [[A("s9-e1")], [A("s9-e2")], [A("s9-e3")]];
  }

  // ═════════════════════════ 10. 실현 가능성 · 안전장치 · 마무리
  {
    add(10, "CLOSE", "기대효과 · 실현 가능성", "전부 새로 만들지 않고,\n기존 기능을 안전하게 연결합니다");
    eyebrow("실현 가능성 · 안전장치");
    const c0 = G.mx, w0 = 1.25, c1 = c0 + w0, w1 = 3.0, c2 = c1 + w1, w2 = G.leftW - w0 - w1;
    const hy = 2.6, hh = 0.44, rh = 0.58;
    const rows = [
      ["위험점검", "전세지킴이 · 등기변동알리미", "계약서 정보 추출(OCR)"],
      ["자금 설계", "주택담보 · 디딤돌대출 취급", "잔금 지도"],
      ["잔금 실행", "예약이체 · 오픈뱅킹 조회", "안심금고 · 조건부 이체"],
      ["입주 후", "자동이체 · 예적금 · 보험", "기존 기능 연결"],
    ];
    text("단계", c0, hy, w0, hh, { size: 13, bold: true, color: C.gray, valign: "middle" });
    rows.forEach((r, i) => {
      const y = hy + hh + i * rh;
      line(c0, y, c0 + G.leftW, y, { color: C.line, w: 1 });
      text(r[0], c0, y, w0, rh, { size: 16, bold: true, valign: "middle" });
    });
    line(c0, hy + hh + 4 * rh, c0 + G.leftW, hy + hh + 4 * rh, { color: C.line, w: 1 });
    group("s10-old", () => {
      rect(c1 - 0.12, hy, w1 - 0.06, hh - 0.06, { fill: C.grayBg, r: 0.08 });
      text("기존 기반", c1 + 0.04, hy, w1 - 0.3, hh - 0.06, { size: 13, bold: true, color: C.ink2, valign: "middle" });
      rows.forEach((r, i) => text(r[1], c1 + 0.04, hy + hh + i * rh, w1 - 0.1, rh, { size: 16, color: C.ink2, valign: "middle" }));
    });
    group("s10-new", () => {
      rect(c2 - 0.12, hy, w2 + 0.12, hh - 0.06, { fill: C.light, r: 0.08 });
      text("신규 개발 · 연계 검토", c2 + 0.04, hy, w2 - 0.1, hh - 0.06, { size: 13, bold: true, color: C.blue, valign: "middle" });
      rows.forEach((r, i) => text(r[2], c2 + 0.04, hy + hh + i * rh, w2 - 0.04, rh, { size: 16, bold: true, color: i === 3 ? C.gray : C.blue, valign: "middle" }));
      text("※ 신규 기능은 내부 시스템 · 외부 기관 연계와 규제 검토가 필요한 제안입니다.", c0, hy + hh + 4 * rh + 0.1, G.leftW, 0.26, { size: 13, color: C.gray, valign: "middle" });
    });
    group("s10-safe", () => {
      const sy = 5.98, sw = (G.leftW - 0.4) / 3, sh = 0.64;
      const items = [["LuScanFace", "고객 최종 승인"], ["LuCirclePause", "조건 미충족 시 중단"], ["LuShieldCheck", "동의 · 전문가 확인"]];
      items.forEach(([ic, t], i) => {
        const x = c0 + i * (sw + 0.2);
        rect(x, sy, sw, sh, { fill: C.bg, r: G.r });
        icon(ic, ic === "LuCirclePause" ? "orange" : "blue", x + 0.16, sy + 0.17, 0.3);
        text(t, x + 0.54, sy, sw - 0.58, sh, { size: 14, bold: true, valign: "middle" });
      });
    });
    // 우측 블루 패널 — 마무리
    const px = G.panelX + 0.62, pw = G.panelW - 1.1;
    rect(px, 0.62, 0.62, 0.62, { fill: C.white, r: 0.14 });
    text("WON 첫집 메이트", px + 0.8, 0.62, 3, 0.62, { size: 15, bold: true, color: C.white, valign: "middle" });
    group("s10-close", () => {
      text("첫 집의\n첫날부터,\n우리은행이\n끝까지 함께.", px, 1.85, pw, 3.55, { size: 46, bold: true, color: C.white, ls: 1.12, valign: "middle" });
      text("감사합니다", px, 5.62, pw, 0.36, { size: 20, bold: true, color: C.white, valign: "middle" });
      text("팀 WEN  ·  유지민 · 김승현 · 이영준", px, 6.06, pw, 0.3, { size: 13, color: C.blueSoft, valign: "middle" });
    });
    anims[10] = [[A("s10-old")], [A("s10-new")], [A("s10-safe")], [A("s10-close", "fade")]];
  }
  pres.slides[9].addImage({ path: LOGO, x: G.panelX + 0.62 + 0.085, y: 0.62 + 0.084, w: 0.45, h: 0.453, altText: "우리WON뱅킹 로고" });

  await pres.writeFile({ fileName: path.join(OUT, "deck_raw.pptx") });
  // 애니메이션은 사용하지 않음 (그룹 이름만 편집 편의를 위해 유지)
  fs.writeFileSync(path.join(OUT, "anim.json"), JSON.stringify({ anims: {}, theme: THEME }, null, 2));
  if (ui.warnings.length) {
    console.log(`\n⚠ 텍스트 맞춤 경고 ${ui.warnings.length}건`);
    ui.warnings.forEach((w) => console.log("  " + w));
  } else console.log("텍스트 맞춤 경고 없음");
  console.log("→ out/deck_raw.pptx");
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
