// 디자인 시스템 — 색상, 규격, 폰트
// 테마 색(스킴 컬러)은 postprocess.py 가 ppt/theme/theme1.xml 에 기록합니다.

const W = 13.333;
const H = 7.5;

const FONT = "Pretendard";

// PowerPoint 테마 색상표 (디자인 > 색에서 한 번에 바꿀 수 있는 핵심 팔레트)
const THEME = {
  name: "WON 첫집 메이트",
  font: FONT,
  colors: {
    dk1: "10233E", // Deep Navy — 본문
    lt1: "FFFFFF",
    dk2: "5B6B7F", // 보조 텍스트
    lt2: "EAF5FC", // Light Blue — 패널
    accent1: "006BB6", // Primary Blue
    accent2: "1BA5D7", // Accent Cyan
    accent3: "F47B20", // Accent Orange — 승인·경고에만
    accent4: "12A870", // 확인 완료
    accent5: "33475E", // 본문 2
    accent6: "D9E2EC", // 구분선
    hlink: "006BB6",
    folHlink: "5B6B7F",
  },
};

// 스킴 컬러 이름 (핵심) + 헥스 (틴트)
const C = {
  navy: "tx1",
  white: "bg1",
  gray: "tx2",
  light: "bg2",
  blue: "accent1",
  cyan: "accent2",
  orange: "accent3",
  green: "accent4",
  ink2: "accent5",
  line: "accent6",
  // tints (hex)
  bg: "F4F7FB",
  blueDeep: "00508A",
  blueSoft: "D5EAF7",
  grayLight: "A3AFBD",
  grayBg: "EEF2F6",
  orangeSoft: "FEF1E6",
  orangeDeep: "C2570A",
  greenSoft: "E3F6EE",
  greenDeep: "0B7F54",
  yellow: "F5B301",
  device: "0E1726",
};

// 아이콘 래스터화용 실제 헥스
const HEX = {
  navy: "10233E",
  white: "FFFFFF",
  gray: "5B6B7F",
  blue: "006BB6",
  cyan: "1BA5D7",
  orange: "F47B20",
  green: "12A870",
  ink2: "33475E",
  grayLight: "A3AFBD",
  orangeDeep: "C2570A",
  greenDeep: "0B7F54",
};

// 레이아웃 그리드 (inch)
const G = {
  mx: 0.8, // 좌우 여백
  fullW: 13.333 - 1.6, // 전체 폭 콘텐츠
  leftW: 7.05, // 폰 슬라이드의 좌측 콘텐츠 폭
  panelX: 8.45, // 우측 패널 시작
  panelW: 13.333 - 8.45,
  eyebrowY: 0.6,
  titleY: 0.98,
  footY: 6.98,
  // 폰 목업
  phoneW: 3.0,
  phoneH: 6.16,
  phoneX: 8.45 + (13.333 - 8.45 - 3.0) / 2,
  phoneY: 0.56,
  r: 0.16, // 카드 모서리 반경 (통일)
};

module.exports = { W, H, FONT, THEME, C, HEX, G };
