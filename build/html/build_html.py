#!/usr/bin/env python3
"""WON 첫집 메이트 HTML 발표자료 빌드.

deck_src.html → ../../deck/WON첫집메이트_본선발표.html (인터넷 없이 열리는 단일 파일)
  - Pretendard Variable을 덱에 쓰인 글자만 남겨 woff2로 내장
  - WON 로고 PNG를 data URI로 내장
  - 발표자 노트에서 대본 문서(docs/본선발표_HTML_대본.md)와 시간표를 생성

필요: Python 3, fonttools, brotli (pip install fonttools brotli)
폰트 원본: build/node_modules/pretendard (npm install) 또는 환경변수 PRETENDARD_WOFF2
"""
import base64, html, os, re, subprocess, sys, tempfile
from html.parser import HTMLParser
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent.parent
SRC = HERE / "deck_src.html"
OUT = ROOT / "deck" / "WON첫집메이트_본선발표.html"
DOC = ROOT / "docs" / "본선발표_HTML_대본.md"
LOGO = ROOT / "build" / "assets" / "won_logo.png"
FONT = Path(os.environ.get("PRETENDARD_WOFF2",
            ROOT / "build" / "node_modules" / "pretendard" / "dist" / "web" / "variable" / "woff2" / "PretendardVariable.woff2"))

SB_ICONS = ('<svg width="20" height="13" viewBox="0 0 20 13"><rect x="0" y="9" width="3.4" height="4" rx="1" fill="#141a23"/>'
            '<rect x="5.4" y="6" width="3.4" height="7" rx="1" fill="#141a23"/><rect x="10.8" y="3" width="3.4" height="10" rx="1" fill="#141a23"/>'
            '<rect x="16.2" y="0" width="3.4" height="13" rx="1" fill="#141a23"/></svg>'
            '<svg width="18" height="13" viewBox="0 0 18 13"><path d="M1 4.6a11.5 11.5 0 0116 0M3.8 7.6a7.4 7.4 0 0110.4 0M6.6 10.4a3.4 3.4 0 014.8 0" '
            'fill="none" stroke="#141a23" stroke-width="2" stroke-linecap="round"/></svg>'
            '<svg width="28" height="13" viewBox="0 0 28 13"><rect x=".8" y=".8" width="23" height="11.4" rx="3.4" fill="none" stroke="#141a23" '
            'stroke-width="1.4" opacity=".45"/><rect x="2.6" y="2.6" width="17" height="7.8" rx="2" fill="#141a23"/>'
            '<rect x="25.4" y="4.3" width="2" height="4.4" rx="1" fill="#141a23" opacity=".45"/></svg>')


# ── 낭독 음절 수(대략): 한글 음절 + 숫자 읽기 + 영문 약어 ──
EN = {"AI": 3, "WON": 1, "WEN": 1, "D": 1, "SSAFY": 2, "CASS": 2, "OTP": 3, "MVP": 4, "PDF": 4}


def num_syl(s):
    s = s.replace(",", "")
    if not s.isdigit():
        return len(s)
    n = int(s)
    if n == 0:
        return 1
    tot, groups = 0, []
    while n > 0:
        groups.append(n % 10000)
        n //= 10000
    for gi, g in enumerate(groups):
        if g == 0:
            continue
        if gi > 0:
            tot += 1
        for p, ch in enumerate(str(g)[::-1]):
            c = int(ch)
            if c == 0:
                continue
            tot += 1 if p == 0 else (1 if c == 1 else 2)
    return tot


def syllables(text):
    tot = 0
    for m in re.finditer(r"[가-힣]+|\d[\d,]*(?:\.\d+)?|[A-Za-z]+|%", text):
        w = m.group(0)
        if re.match(r"[가-힣]", w):
            tot += len(w)
        elif w == "%":
            tot += 3
        elif w[0].isdigit():
            if "." in w:
                a, b = w.split(".")
                tot += num_syl(a) + 1 + len(b)
            else:
                tot += num_syl(w)
        else:
            tot += EN.get(w.upper(), len(w))
    return tot


class FrameParser(HTMLParser):
    """본문 장(section.frame)의 속성과 노트 문단을 모은다."""

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.frames, self.cur, self.in_notes, self.p, self.depth = [], None, False, None, 0

    def handle_starttag(self, tag, attrs):
        a = dict(attrs)
        if tag == "section" and "frame" in (a.get("class") or ""):
            self.cur = {"id": a.get("id"), "name": a.get("data-name", ""), "who": a.get("data-who"),
                        "time": int(a.get("data-time") or 0), "appx": a.get("data-appx"), "paras": []}
            self.frames.append(self.cur)
        elif tag == "aside" and "notes" in (a.get("class") or "") and self.cur is not None:
            self.in_notes = True
        elif self.in_notes and tag == "p":
            self.p = {"cls": a.get("class") or "", "text": ""}
        elif self.in_notes and self.p is not None and tag == "i" and "cue" in (a.get("class") or ""):
            self.p["text"] += " "

    def handle_endtag(self, tag):
        if tag == "aside" and self.in_notes:
            self.in_notes = False
        elif tag == "p" and self.p is not None:
            self.p["text"] = re.sub(r"\s+", " ", self.p["text"]).strip()
            self.cur["paras"].append(self.p)
            self.p = None

    def handle_data(self, data):
        if self.p is not None:
            self.p["text"] += data


def mmss(s):
    return f"{s // 60}:{s % 60:02d}"


def build_doc(src):
    fp = FrameParser()
    fp.feed(src)
    main = [f for f in fp.frames if not f["appx"]]
    speakers = dict(re.findall(r"([ABC]):'([^']+)'", re.search(r"const SPEAKERS=\{([^}]*)\}", src).group(1)))
    lines = ["# WON 첫집 메이트 — 본선 HTML 발표 대본·진행표", "",
             "> `deck/WON첫집메이트_본선발표.html`의 발표자 노트에서 자동 생성했습니다. 대본을 고칠 때는 `build/html/deck_src.html`의 노트를 고치고 다시 빌드하세요.",
             "> ▶n은 그 지점에서 키를 눌러 n단계를 띄운다는 뜻입니다. 낭독 시간은 분당 280음절(여유 있는 속도) 기준 추정입니다.", "",
             "## 진행표", "", "| # | 슬라이드 | 발표 | 권장 | 누적 | 음절 | 280음절/분 추정 |", "|---|---|---|---|---|---|---|"]
    acc = tot_syl = 0
    rows = []
    for k, f in enumerate(main, 1):
        speech = " ".join(p["text"] for p in f["paras"] if "tip" not in p["cls"]).replace("→ 인수인계:", "")
        syl = syllables(speech)
        acc += f["time"]
        tot_syl += syl
        est = round(syl * 60 / 280)
        rows.append((k, f, syl))
        lines.append(f"| {k} | {f['name']} | {f['who']} {speakers.get(f['who'], '')} | {mmss(f['time'])} | {mmss(acc)} | {syl} | {mmss(est)} |")
    lines += ["", f"- 권장 시간 합계 **{mmss(acc)}**, 대본 {tot_syl}음절 → 분당 280음절이면 약 **{mmss(round(tot_syl * 60 / 280))}**, 분당 300음절이면 약 {mmss(round(tot_syl / 5))}.",
              "- 장 전환·시연 조작 여유를 더해 10분 안에 끝나도록 잡았습니다(전환 1~2초 × 13장, 시연 키 조작 포함 약 30초).", ""]
    by = {}
    for k, f, syl in rows:
        by.setdefault(f["who"], 0)
        by[f["who"]] += f["time"]
    lines.append("- 발표 분담: " + " · ".join(f"{w} {speakers.get(w, '')} {mmss(t)}" for w, t in sorted(by.items())))
    lines += ["", "## 슬라이드별 대본", ""]
    for k, f, syl in rows:
        lines.append(f"### {k}. {f['name']} — {f['who']} {speakers.get(f['who'], '')} · 권장 {mmss(f['time'])} · {syl}음절")
        lines.append("")
        for p in f["paras"]:
            t = p["text"]
            if "tip" in p["cls"]:
                lines.append(f"> 연출 · {t}")
            elif "hand" in p["cls"]:
                lines.append(f"**{t}**")
            else:
                lines.append(t)
            lines.append("")
    return "\n".join(lines).replace("  ", " ")


def cue_fix(src):
    # 노트의 ▶n 표시는 대본 문서에서 그대로 보이도록 둔다
    return src


def subset_font(text):
    if not FONT.exists():
        sys.exit(f"폰트 원본이 없습니다: {FONT}\n  build 폴더에서 npm install 을 하거나 PRETENDARD_WOFF2 환경변수로 경로를 지정하세요.")
    chars = set(text) | set(" 0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ.,:;!?()[]{}<>+-=*/%·…‘’“”'\"~@#&|_→←▶◀›‹×○−")
    chars = {c for c in chars if ord(c) >= 32}
    with tempfile.TemporaryDirectory() as td:
        tf = Path(td) / "chars.txt"
        tf.write_text("".join(sorted(chars)), encoding="utf-8")
        out = Path(td) / "sub.woff2"
        cmd = [sys.executable, "-m", "fontTools.subset", str(FONT), f"--text-file={tf}", f"--output-file={out}",
               "--flavor=woff2", "--layout-features=*", "--no-hinting", "--desubroutinize"]
        subprocess.run(cmd, check=True)
        return base64.b64encode(out.read_bytes()).decode()


def main():
    src = SRC.read_text(encoding="utf-8")
    logo = "data:image/png;base64," + base64.b64encode(LOGO.read_bytes()).decode()
    page = src.replace("{{SB_ICONS}}", SB_ICONS)
    # 글자 목록은 로고/폰트를 넣기 전 원본에서 뽑는다(스크립트 문자열 포함)
    b64 = subset_font(page)
    face = ('@font-face{font-family:"Pretendard";font-style:normal;font-weight:45 920;font-display:block;'
            f'src:url(data:font/woff2;base64,{b64}) format("woff2")}}')
    page = page.replace("/*FONTS*/", face).replace("{{LOGO}}", logo)
    leftover = re.findall(r"\{\{[A-Z_]+\}\}", page)
    if leftover:
        sys.exit(f"치환되지 않은 자리표시자: {set(leftover)}")
    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text(page, encoding="utf-8")
    DOC.write_text(build_doc(src), encoding="utf-8")
    print(f"HTML  {OUT.relative_to(ROOT)}  {OUT.stat().st_size / 1024:.0f} KB (폰트 {len(b64) * 3 / 4 / 1024:.0f} KB)")
    print(f"대본  {DOC.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
