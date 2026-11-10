// content/script.js → docs/발표대본.md, docs/스토리보드.md 생성 + 발표 시간 검증
const fs = require("fs");
const path = require("path");
const { SLIDES } = require("./content/script");

const PACE = 5.0; // 초당 음절 (분당 300음절 — 또박또박 말하는 발표 속도)
const CLICK = 0.6; // 화면을 가리키며 잠깐 멈추는 시간(초)
const TURN = 1.0; // 슬라이드 전환 여유(초)

// ── 읽는 음절 수 추정: 숫자는 한자어 읽기로 바꿔 센다
const DIG = ["", "일", "이", "삼", "사", "오", "육", "칠", "팔", "구"];
function sino(n) {
  if (n === 0) return "영";
  const unit = ["", "십", "백", "천"];
  const big = ["", "만", "억", "조"];
  let out = "";
  const groups = [];
  while (n > 0) {
    groups.push(n % 10000);
    n = Math.floor(n / 10000);
  }
  for (let g = groups.length - 1; g >= 0; g--) {
    let v = groups[g];
    if (!v) continue;
    let s = "";
    const ds = String(v).padStart(4, "0").split("").map(Number);
    ds.forEach((d, i) => {
      if (!d) return;
      const u = unit[3 - i];
      s += (d === 1 && u ? "" : DIG[d]) + u;
    });
    if (g === 1 && s === "일") s = ""; // 만 → '일만' 아님
    out += s + big[g];
  }
  return out;
}
function spoken(text) {
  let t = text.replace(/\([^)]*\)/g, (m) => m); // 괄호 문장도 읽는 것으로 계산
  t = t.replace(/WON/g, "원").replace(/WEN/g, "웬").replace(/AI/g, "에이아이");
  t = t.replace(/(\d+)\.(\d+)/g, (_, a, b) => sino(+a) + "점" + b.split("").map((d) => DIG[+d] || "영").join(""));
  t = t.replace(/\d[\d,]*/g, (m) => sino(parseInt(m.replace(/,/g, ""), 10)));
  return t;
}
const syll = (text) => (spoken(text).match(/[가-힣]/g) || []).length;
const mmss = (s) => {
  const t = Math.round(s);
  return `${Math.floor(t / 60)}:${String(t % 60).padStart(2, "0")}`;
};

let total = 0, totalPlan = 0, totalSyl = 0, totalClicks = 0, totalOpt = 0;
const rows = [];
for (const s of SLIDES) {
  const body = s.segments.map((x) => x.text).join(" ");
  const tr = s.no < 10 ? s.transition : "";
  const n = syll(body) + syll(tr);
  const opt = s.segments.filter((x) => x.text.startsWith("(")).reduce((a, x) => a + syll(x.text), 0);
  const clicks = s.segments.filter((x) => x.cue).length;
  const est = n / PACE + clicks * CLICK + TURN;
  rows.push({ s, n, clicks, est, opt });
  total += est; totalPlan += s.sec; totalSyl += n; totalClicks += clicks; totalOpt += opt;
}

// ── 발표대본.md
let md = `# WON 첫집 메이트 — 10분 발표 대본\n\n`;
md += `> WOORI × SSAFY 우리WON뱅킹 아이디어 경진대회 본선 · 팀 WEN (유지민 · 김승현 · 이영준)\n>\n`;
md += `> 이 문서는 \`build/content/script.js\` 에서 자동 생성됩니다. 같은 내용이 PPTX의 발표자 노트에도 들어 있습니다.\n\n`;
md += `## 읽는 법\n\n`;
md += `- 슬라이드에 애니메이션은 없습니다. **◆ 가리킬 곳 n** 표시는 그 문장을 말할 때 레이저 포인터나 손으로 짚을 화면 요소입니다.\n`;
md += `- *(괄호로 묶인 문장)* 은 시간이 부족하면 생략해도 흐름이 끊기지 않는 문장입니다.\n`;
md += `- **[전환]** 은 다음 슬라이드로 넘기면서 하는 말입니다. 문장을 말한 뒤 넘기세요.\n`;
md += `- 리모컨은 슬라이드를 넘길 때만 누릅니다(총 9회).\n\n`;
md += `## 발표 시간 검증\n\n`;
md += `말하기 속도는 분당 300음절(초당 ${PACE}음절)로 잡았고, 숫자는 실제로 읽는 방식(예: 4억 5,000만 → 사억 오천만)으로 음절을 셌습니다. 화면을 가리킬 때마다 ${CLICK}초, 슬라이드 전환마다 ${TURN}초의 여유를 더했습니다.\n\n`;
md += `| 장 | 슬라이드 | 배정 | 음절 | 가리킴 | 예상 |\n|---|---|---|---|---|---|\n`;
for (const r of rows) md += `| ${r.s.no} | ${r.s.title} | ${mmss(r.s.sec)} | ${r.n} | ${r.clicks} | ${mmss(r.est)} |\n`;
md += `| **합계** | | **${mmss(totalPlan)}** | **${totalSyl}** | **${totalClicks}** | **${mmss(total)}** |\n\n`;
const fast = (totalSyl / 5.4) + totalClicks * CLICK + 10 * TURN;
const slow = (totalSyl / 4.7) + totalClicks * CLICK + 10 * TURN;
const cut = ((totalSyl - totalOpt) / PACE) + totalClicks * CLICK + 10 * TURN;
md += `- 기준 속도(분당 300음절): **${mmss(total)}**\n`;
md += `- 조금 빠르게(분당 약 325음절): ${mmss(fast)} · 조금 느리게(분당 약 280음절): ${mmss(slow)}\n`;
md += `- 괄호 문장을 모두 생략하면: ${mmss(cut)} (약 ${Math.round(totalOpt / PACE)}초 단축)\n`;
md += `- 리허설에서 7번 슬라이드를 마쳤을 때 **5분 50초 안팎**, 8번을 마쳤을 때 **7분 35초 안팎**이면 정상 페이스입니다.\n\n---\n\n`;
let acc = 0;
for (const r of rows) {
  const s = r.s;
  md += `## ${s.no}. ${s.title}\n\n`;
  md += `- **배정 시간**: ${mmss(s.sec)} (누적 ${mmss(acc)} → ${mmss(acc + s.sec)})\n`;
  md += `- **핵심 메시지**: ${s.board.message}\n\n`;
  md += `**발표 대본**\n\n`;
  let click = 0;
  for (const seg of s.segments) {
    if (seg.cue) {
      click += 1;
      md += `> **◆ 가리킬 곳 ${click}** — ${seg.cue}\n\n`;
    }
    md += seg.text.startsWith("(") ? `*${seg.text}*\n\n` : `${seg.text}\n\n`;
  }
  md += `**[전환]** ${s.transition}\n\n`;
  md += `**시선 유도 순서**: ${s.board.animation}\n\n---\n\n`;
  acc += s.sec;
}
fs.writeFileSync(path.join(__dirname, "..", "docs", "발표대본.md"), md);

// ── 스토리보드.md
let sb = `# WON 첫집 메이트 — 슬라이드별 스토리보드\n\n`;
sb += `> 흐름: **문제 상황(2) → 서비스 기획(3–4) → 서비스 UI(5–8) → 기대효과(9)** · 실현 가능성과 안전장치(10)는 이를 뒷받침합니다.\n>\n`;
sb += `> 한 고객의 한마디 “저 12월 20일에 이사 가요.”에서 시작해, 같은 고객·같은 금액으로 끝까지 이어집니다.\n\n`;
sb += `| 장 | 시간 | 핵심 메시지 |\n|---|---|---|\n`;
for (const s of SLIDES) sb += `| ${s.no} | ${mmss(s.sec)} | ${s.board.message} |\n`;
sb += `\n---\n\n`;
for (const s of SLIDES) {
  sb += `## ${s.no}. ${s.title} (${mmss(s.sec)})\n\n`;
  sb += `| 항목 | 내용 |\n|---|---|\n`;
  sb += `| 핵심 메시지 | ${s.board.message} |\n`;
  sb += `| 화면 배치 의도 | ${s.board.layout} |\n`;
  sb += `| 사용한 시각 자료 | ${s.board.visuals} |\n`;
  sb += `| 시선 유도 순서 | ${s.board.animation} |\n`;
  sb += `| 강조할 문장 | “${s.board.emphasis}” |\n\n`;
}
fs.writeFileSync(path.join(__dirname, "..", "docs", "스토리보드.md"), sb);

console.log("장  배정   음절  가리킴  예상");
for (const r of rows) console.log(String(r.s.no).padStart(2), mmss(r.s.sec).padStart(5), String(r.n).padStart(5), String(r.clicks).padStart(4), mmss(r.est).padStart(6));
console.log(`합계 ${mmss(totalPlan)}  ${totalSyl}음절  가리킴 ${totalClicks}  예상 ${mmss(total)} (빠르게 ${mmss(fast)} / 느리게 ${mmss(slow)} / 생략 시 ${mmss(cut)})`);
