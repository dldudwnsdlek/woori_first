// Q&A 부록 덱 (발표 10분에 포함하지 않는 백업 슬라이드)
const path = require("path");
const fs = require("fs");
const pptxgen = require("pptxgenjs");
const { W, H, FONT, THEME, C, G } = require("./lib/theme");
const icons = require("./lib/icons");
const ui = require("./lib/ui");
const { rect, line, text, pill, icon, iconCircle, eyebrow } = ui;

const OUT = path.join(__dirname, "out");
fs.mkdirSync(OUT, { recursive: true });
const LOGO = path.join(__dirname, "assets", "won_logo.png");

async function main() {
  await icons.prepare([
    ["LuChevronLeft", "navy"], ["LuCheck", "white"], ["LuArrowRight", "blue"],
    ["LuSparkles", "blue"], ["LuLandmark", "blue"], ["LuUserCheck", "blue"], ["LuCircleAlert", "white"],
  ]);
  const pres = new pptxgen();
  pres.defineLayout({ name: "WIDE", width: W, height: H });
  pres.layout = "WIDE";
  pres.title = "WON 첫집 메이트 — Q&A 부록";
  pres.author = "팀 WEN";
  pres.theme = { headFontFace: FONT, bodyFontFace: FONT };
  ui.init(pres);
  pres.defineSlideMaster({
    title: "FULL", background: { color: "FFFFFF" },
    objects: [
      { image: { x: G.mx, y: 6.9, w: 0.3, h: 0.302, path: LOGO } },
      { text: { text: "WON 첫집 메이트 · Q&A 부록", options: { x: G.mx + 0.42, y: 6.9, w: 4, h: 0.3, fontFace: FONT, fontSize: 10.5, bold: true, color: C.gray, valign: "middle", margin: 0 } } },
      { placeholder: { options: { name: "title", type: "title", x: G.mx, y: G.titleY, w: G.fullW, h: 0.7, fontFace: FONT, fontSize: 34, bold: true, color: C.navy, align: "left", valign: "top", margin: 0 }, text: "" } },
    ],
    slideNumber: { x: 12.03, y: 6.9, w: 0.5, h: 0.3, fontFace: FONT, fontSize: 10.5, color: C.gray, align: "right", valign: "middle", margin: 0 },
  });
  const add = (label, title, notes) => {
    const s = pres.addSlide({ masterName: "FULL" });
    ui.use(s, label);
    s.addText(title, { placeholder: "title" });
    if (notes) s.addNotes(notes);
    eyebrow(label);
    return s;
  };
  const col3 = (i) => G.mx + i * ((G.fullW - 0.6) / 3 + 0.3);
  const cw = (G.fullW - 0.6) / 3;

  // ── 부록 1. 금액 검증
  {
    add("부록 1 · 예시 금액", "예시 금액은 조달부터 지급까지 서로 일치합니다",
      "잔금 4억 5,000만 원은 ‘어디서 오는가’와 ‘어떻게 나가는가’ 양쪽에서 같은 합계가 됩니다. 고객이 직접 승인하는 송금은 안심금고의 2억 1,000만 원이고, 대출 2억 4,000만 원은 은행이 실행합니다. 부대비용 620만 원은 잔금과 별도입니다.");
    const blocks = [
      ["잔금은 얼마인가", [["매매가", "5억 원"], ["계약금 (지급 완료)", "− 5,000만 원"]], "잔금", "4억 5,000만 원"],
      ["어디서 오는가", [["기존 집 보증금", "2억 원"], ["디딤돌대출", "2억 4,000만 원"], ["내 예금", "1,000만 원"]], "합계", "4억 5,000만 원"],
      ["어떻게 나가는가", [["대출 → 매도인 대출 상환", "2억 원"], ["대출 → 매도인 지급", "4,000만 원"], ["안심금고 (고객 승인)", "2억 1,000만 원"]], "합계", "4억 5,000만 원"],
    ];
    const y0 = 2.05, bh = 3.15;
    blocks.forEach(([h, rows, sumK, sumV], i) => {
      const x = col3(i);
      rect(x, y0, cw, bh, { fill: C.bg, r: G.r });
      text(h, x + 0.3, y0 + 0.22, cw - 0.6, 0.4, { size: 20, bold: true, color: C.blue, valign: "middle" });
      rows.forEach(([k, v], j) => {
        const ry = y0 + 0.8 + j * 0.48;
        text(k, x + 0.3, ry, cw - 0.6, 0.4, { size: 14, color: C.ink2, valign: "middle" });
        text(v, x + 0.3, ry, cw - 0.6, 0.4, { size: 16, bold: true, align: "right", valign: "middle", nocheck: true });
      });
      line(x + 0.3, y0 + 2.36, x + cw - 0.3, y0 + 2.36, { color: C.line, w: 1 });
      text(sumK, x + 0.3, y0 + 2.5, 1.2, 0.44, { size: 15, bold: true, valign: "middle" });
      text(sumV, x + 0.3, y0 + 2.5, cw - 0.6, 0.44, { size: 22, bold: true, color: C.blue, align: "right", valign: "middle", nocheck: true });
    });
    rect(G.mx, 5.45, G.fullW, 1.05, { fill: C.orangeSoft, r: G.r });
    iconCircle("LuCircleAlert", G.mx + 0.28, 5.73, 0.5, { fill: C.orange, color: "white", scale: 0.56 });
    text("잔금과 별도로 드는 부대비용 620만 원", G.mx + 1.0, 5.56, 7, 0.42, { size: 20, bold: true, color: C.orangeDeep, valign: "middle" });
    text("취득세 등 350만 + 중개보수 200만(5억 × 상한 0.4%) + 법무사 70만 · 잔금 지도가 2주 전에 부족분으로 알리는 금액", G.mx + 1.0, 6.0, G.fullW - 1.3, 0.3, { size: 14, color: C.ink2, valign: "middle" });
  }

  // ── 부록 2. 역할 구분
  {
    add("부록 2 · 역할과 책임", "AI는 안내하고, 은행 시스템이 실행하며, 고객이 승인합니다",
      "AI가 독자적으로 판단하거나 송금하지 않습니다. 정보 정리와 순서 안내는 AI, 조건 확인과 금융거래는 은행 시스템과 관계 기관, 최종 결정은 고객과 전문가의 몫입니다.");
    const cols = [
      ["LuSparkles", "AI (첫집 메이트)", "정리하고 안내합니다", ["계약서 정보 정리", "이사일 기준 일정 역산 · 알림", "필요 자금 계산 · 부족분 안내", "잔금일 순서 안내"], "판단 · 송금은 하지 않음"],
      ["LuLandmark", "은행 시스템 · 관계 기관", "확인하고 실행합니다", ["등기 변동 · 입금 여부 확인", "대출 심사 · 실행", "안심금고 보관 · 조건부 이체 처리", "본인 인증"], "조건 미충족 시 거래 중단"],
      ["LuUserCheck", "고객 · 전문가", "결정하고 승인합니다", ["읽어 낸 계약 정보 확인 · 수정", "정보 조회 · 이용 동의", "최종 송금 승인", "법무사 · 공인중개사 확인"], "돈이 나가는 순간의 결정권"],
    ];
    const y0 = 2.05, bh = 4.45;
    cols.forEach(([ic, h, sub, items, foot], i) => {
      const x = col3(i);
      rect(x, y0, cw, bh, { fill: i === 1 ? C.light : C.bg, r: G.r });
      iconCircle(ic, x + 0.3, y0 + 0.3, 0.6, { fill: C.white });
      text(h, x + 1.05, y0 + 0.26, cw - 1.3, 0.38, { size: 19, bold: true, valign: "middle" });
      text(sub, x + 1.05, y0 + 0.64, cw - 1.3, 0.3, { size: 14, color: C.blue, bold: true, valign: "middle" });
      items.forEach((t, j) => {
        const ry = y0 + 1.3 + j * 0.5;
        rect(x + 0.34, ry + 0.16, 0.1, 0.1, { fill: C.blue, r: 0.05 });
        text(t, x + 0.6, ry, cw - 0.85, 0.42, { size: 16, color: C.ink2, valign: "middle" });
      });
      pill(foot, x + 0.3, y0 + bh - 0.72, cw - 0.6, 0.44, { fill: C.white, color: C.navy, size: 14, r: 0.1 });
    });
  }

  // ── 부록 3. 검토 과제와 단계적 도입
  {
    add("부록 3 · 실현 과제", "새 기능은 단계적으로 검증하며 도입하는 방식을 제안합니다",
      "1·2단계는 기존 전세지킴이와 등기변동알리미, 일정·알림 기능을 재사용해 먼저 만들 수 있습니다. 안심금고와 조건부 이체는 상품 설계와 규제 검토, 외부 기관 연계가 필요하므로 시범 운영으로 검증하는 순서를 제안합니다.");
    const steps = [
      ["1단계", "점검 · 일정", "기존 진단 · 알림 기능을 재사용해 매매 계약으로 확장"],
      ["2단계", "잔금 지도", "대출 · 계좌 정보와 부대비용 계산을 묶은 자금 설계 화면"],
      ["3단계", "안심금고 · 조건부 이체", "상품 · 약관 설계와 규제 검토 후 시범 운영으로 검증"],
    ];
    steps.forEach(([n, h, d], i) => {
      const x = col3(i);
      rect(x, 2.05, cw, 1.7, { fill: i === 2 ? C.blue : C.bg, r: G.r });
      text(n, x + 0.3, 2.22, 2, 0.3, { size: 14, bold: true, color: i === 2 ? C.blueSoft : C.blue, valign: "middle" });
      text(h, x + 0.3, 2.55, cw - 0.6, 0.42, { size: 21, bold: true, color: i === 2 ? C.white : C.navy, valign: "middle" });
      text(d, x + 0.3, 3.02, cw - 0.6, 0.6, { size: 14, color: i === 2 ? C.blueSoft : C.gray, ls: 1.1 });
      if (i > 0) icon("LuArrowRight", "blue", x - 0.27, 2.78, 0.24);
    });
    text("도입 전에 검토가 필요한 과제", G.mx, 4.1, 6, 0.4, { size: 18, bold: true, valign: "middle" });
    const tasks = [
      ["규제 · 상품 설계", "조건부 이체와 안심금고의 약관, 전자금융거래 관련 법규 검토"],
      ["외부 데이터 연계", "등기 · 체납 등 공공 데이터 조회와 제휴 범위, 조회 주기"],
      ["타 기관 자금 확인", "보증금 입금 · 매도인 대출 상환 여부를 확인하는 방식"],
      ["동의와 정보보호", "매도인 정보 조회 동의, 계약 문서 보관 · 삭제 기준"],
      ["정확도와 책임", "계약서 인식 오류에 대비한 고객 확인 절차와 책임 범위"],
    ];
    tasks.forEach(([k, v], i) => {
      const y = 4.62 + i * 0.42;
      line(G.mx, y, G.mx + G.fullW, y, { color: C.line, w: 1 });
      text(k, G.mx, y, 2.6, 0.42, { size: 15, bold: true, valign: "middle" });
      text(v, G.mx + 2.7, y, G.fullW - 2.7, 0.42, { size: 15, color: C.ink2, valign: "middle" });
    });
  }

  // ── 부록 4. 근거와 출처
  {
    add("부록 4 · 근거", "발표에 사용한 수치와 제도의 근거입니다",
      "제도와 한도는 바뀔 수 있어 실제 서비스에서는 최신 공식 기준을 연동해야 합니다. 취득세는 생애최초 감면 적용을 가정한 예시입니다.");
    const rows = [
      ["평균 거주 기간 3.6년", "국토교통부 「2024년도 주거실태조사」(2025.11 발표) — 임차 가구 3.6년, 자가 가구 11.5년, 전체 8.4년"],
      ["전세지킴이 · 등기변동알리미", "우리은행 × 세이프홈즈 제휴, 우리WON뱅킹 제공 (2025.6 발표) — 주소 · 보증금 입력 기반 전세 진단"],
      ["체크리스트형 서비스", "KB부동산 ‘전세안심 Check’ (2026.8 출시) — 매물 검색부터 이사까지 단계별 확인 사항 안내"],
      ["디딤돌대출 2억 4,000만 원", "생애최초 일반 가구 한도 (2025.6.27 대책 이후 기준). 실제 한도는 소득 · 주택가격 · LTV에 따라 달라짐"],
      ["중개보수 200만 원", "매매 2억 원 이상 9억 원 미만 상한요율 0.4% (2021.10 개편) × 5억 원. 상한 안에서 협의"],
      ["취득세 등 350만 원", "5억 원 주택 취득세 · 지방교육세에 생애최초 감면 적용을 가정한 예시. 요건 · 기한은 법령 확인 필요"],
      ["전세사기 피해 현황", "국토교통부 발표 누계 4만 1,559건, 40세 미만 75.9% (2026.10 발표). 제안서의 4만 936건은 2026.9 발표 기준"],
    ];
    rows.forEach(([k, v], i) => {
      const y = 2.0 + i * 0.64;
      line(G.mx, y, G.mx + G.fullW, y, { color: C.line, w: 1 });
      text(k, G.mx, y, 3.3, 0.64, { size: 15, bold: true, valign: "middle" });
      text(v, G.mx + 3.4, y, G.fullW - 3.4, 0.64, { size: 13.5, color: C.ink2, valign: "middle", ls: 1.1 });
    });
    line(G.mx, 2.0 + 7 * 0.64, G.mx + G.fullW, 2.0 + 7 * 0.64, { color: C.line, w: 1 });
  }

  await pres.writeFile({ fileName: path.join(OUT, "appendix_raw.pptx") });
  fs.writeFileSync(path.join(OUT, "appendix_anim.json"), JSON.stringify({ anims: {}, theme: THEME }, null, 2));
  if (ui.warnings.length) {
    console.log(`⚠ 텍스트 맞춤 경고 ${ui.warnings.length}건`);
    ui.warnings.forEach((w) => console.log("  " + w));
  } else console.log("텍스트 맞춤 경고 없음");
  console.log("→ out/appendix_raw.pptx");
}
main().catch((e) => { console.error(e); process.exit(1); });
