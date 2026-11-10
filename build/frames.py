#!/usr/bin/env python3
"""애니메이션 단계별 상태를 검수하기 위한 '프레임 덱' 생성.

클릭 k 까지 진행된 화면을 만들기 위해, 그 뒤 클릭에서 나타날 그룹을 지운 PPTX 를 만든다.
사용: python3 frames.py 입력.pptx anim.json 출력폴더   → 출력폴더/state{k}.pptx
"""
import json
import os
import re
import sys
import zipfile

from lxml import etree

NS = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main"}


def main():
    src, cfg, out = sys.argv[1], sys.argv[2], sys.argv[3]
    anims = {int(k): v for k, v in json.load(open(cfg, encoding="utf-8"))["anims"].items()}
    os.makedirs(out, exist_ok=True)
    max_clicks = max(len(v) for v in anims.values())
    for k in range(max_clicks):  # 마지막 상태(전부 표시)는 원본과 같으므로 생략
        zin = zipfile.ZipFile(src)
        zout = zipfile.ZipFile(os.path.join(out, "state%d.pptx" % k), "w", zipfile.ZIP_DEFLATED)
        for item in zin.infolist():
            data = zin.read(item.filename)
            m = re.match(r"ppt/slides/slide(\d+)\.xml$", item.filename)
            if m:
                clicks = anims.get(int(m.group(1)), [])
                hidden = {e["g"] for c in clicks[k:] for e in c}
                root = etree.fromstring(data)
                for g in root.findall(".//p:grpSp", NS):
                    name = g.find("p:nvGrpSpPr/p:cNvPr", NS).get("name")
                    if name in hidden:
                        g.getparent().remove(g)
                for t in root.findall("p:timing", NS):
                    root.remove(t)
                data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
            zout.writestr(item, data)
        zout.close()
    print("프레임 덱 %d개 → %s" % (max_clicks, out))


if __name__ == "__main__":
    main()
