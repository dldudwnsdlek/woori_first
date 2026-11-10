#!/usr/bin/env python3
"""PPTX → PDF (+ 슬라이드 이미지) 렌더링.

LibreOffice 는 한글과 영문·숫자 사이에 자동 간격을 넣어 PowerPoint 와 줄바꿈이 달라집니다.
그래서 PPTX → ODP 로 바꾼 뒤 자동 간격을 끄고 PDF 로 내보냅니다.
(PowerPoint 가 있다면 '파일 > 내보내기 > PDF' 가 가장 정확합니다.)

사용: python3 render.py 입력.pptx 출력.pdf [--png 폴더 --dpi 110]
"""
import argparse
import os
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path


def run_soffice(args):
    helper = os.environ.get("SOFFICE_HELPER_DIR")
    if helper:
        sys.path.insert(0, helper)
        from office.soffice import run_soffice as _run  # type: ignore

        return _run(args, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    with tempfile.TemporaryDirectory(prefix="lo_profile_") as prof:
        cmd = ["soffice", "-env:UserInstallation=" + Path(prof).as_uri()] + list(args)
        return subprocess.run(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)


def convert(src, fmt, outdir):
    run_soffice(["--headless", "--convert-to", fmt, "--outdir", str(outdir), str(src)])
    out = Path(outdir) / (Path(src).stem + "." + fmt)
    if not out.exists():
        raise SystemExit("변환 실패: %s → %s" % (src, fmt))
    return out


def disable_autospace(odp, dst):
    with zipfile.ZipFile(odp) as zin, zipfile.ZipFile(dst, "w") as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename in ("styles.xml", "content.xml"):
                data = data.replace(b'style:text-autospace="ideograph-alpha"', b'style:text-autospace="none"')
            comp = zipfile.ZIP_STORED if item.filename == "mimetype" else zipfile.ZIP_DEFLATED
            zout.writestr(item, data, compress_type=comp)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx")
    ap.add_argument("pdf")
    ap.add_argument("--png", help="슬라이드 이미지를 저장할 폴더")
    ap.add_argument("--dpi", type=int, default=110)
    a = ap.parse_args()
    with tempfile.TemporaryDirectory(prefix="render_") as tmp:
        tmp = Path(tmp)
        work = tmp / "deck.pptx"
        shutil.copy(a.pptx, work)
        odp = convert(work, "odp", tmp)
        (tmp / "fixed").mkdir()
        fixed = tmp / "fixed" / "deck.odp"
        disable_autospace(odp, fixed)
        pdf = convert(fixed, "pdf", tmp / "fixed")
        os.makedirs(os.path.dirname(os.path.abspath(a.pdf)), exist_ok=True)
        shutil.copy(pdf, a.pdf)
    print("PDF →", a.pdf)
    if a.png:
        os.makedirs(a.png, exist_ok=True)
        for f in Path(a.png).glob("s-*.jpg"):
            f.unlink()
        subprocess.run(["pdftoppm", "-jpeg", "-r", str(a.dpi), a.pdf, str(Path(a.png) / "s")], check=True)
        print("이미지 →", a.png)


if __name__ == "__main__":
    main()
