#!/usr/bin/env python3
"""pptxgenjs 결과물 후처리.

1) 같은 이름(`name#n`)의 도형을 PowerPoint 그룹으로 묶음 → 편집·애니메이션 단위
2) 슬라이드별 클릭 애니메이션(<p:timing>) 삽입 — 페이드 / 살짝 올라오며 페이드
3) 모든 슬라이드에 동일한 페이드 화면 전환 삽입
4) 테마 색·글꼴(한글 포함) 기록

사용: python3 postprocess.py [입력.pptx] [anim.json] [출력.pptx]
"""
import copy
import json
import os
import re
import sys
import zipfile
from collections import OrderedDict

from lxml import etree

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, "out", "deck_raw.pptx")
CFG = sys.argv[2] if len(sys.argv) > 2 else os.path.join(HERE, "out", "anim.json")
DST = sys.argv[3] if len(sys.argv) > 3 else os.path.join(HERE, "..", "deck", "WON첫집메이트_본선발표.pptx")

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
P = "{%s}" % NS["p"]
A = "{%s}" % NS["a"]

DUR = 500  # ms — 기본 등장 시간
RISE = "0.025"  # 슬라이드 높이 대비 이동량(약 18px @1080p)


def q(tag):
    pfx, name = tag.split(":")
    return "{%s}%s" % (NS[pfx], name)


def cnvpr(el):
    for path in ("p:nvSpPr/p:cNvPr", "p:nvPicPr/p:cNvPr", "p:nvCxnSpPr/p:cNvPr", "p:nvGrpSpPr/p:cNvPr"):
        r = el.find(path, NS)
        if r is not None:
            return r
    return None


def xfrm_box(el):
    x = el.find("p:spPr/a:xfrm", NS)
    if x is None:
        return None
    off, ext = x.find("a:off", NS), x.find("a:ext", NS)
    ox, oy, cx, cy = int(off.get("x")), int(off.get("y")), int(ext.get("cx")), int(ext.get("cy"))
    return ox, oy, ox + cx, oy + cy


def group_shapes(root):
    """spTree 의 `name#n` 도형을 이름별 그룹으로 묶고 {그룹이름: id} 를 돌려준다."""
    tree = root.find("p:cSld/p:spTree", NS)
    children = [c for c in tree if c.tag in (q("p:sp"), q("p:pic"), q("p:cxnSp"))]
    max_id = 1
    for c in tree.iter(q("p:cNvPr")):
        max_id = max(max_id, int(c.get("id")))
    groups = OrderedDict()
    for c in children:
        pr = cnvpr(c)
        name = pr.get("name", "") if pr is not None else ""
        m = re.match(r"^(.+)#(\d+)$", name)
        if m:
            groups.setdefault(m.group(1), []).append(c)
    ids = {}
    for gname, members in groups.items():
        # 연속성 확인 (z-순서가 바뀌지 않도록)
        idxs = [tree.index(m) for m in members]
        if idxs != list(range(idxs[0], idxs[0] + len(idxs))):
            print("  ! 그룹 '%s' 의 도형이 연속하지 않습니다 — z-순서를 확인하세요" % gname)
        boxes = [b for b in (xfrm_box(m) for m in members) if b]
        x0, y0 = min(b[0] for b in boxes), min(b[1] for b in boxes)
        x1, y1 = max(b[2] for b in boxes), max(b[3] for b in boxes)
        max_id += 1
        gid = max_id
        grp = etree.Element(q("p:grpSp"))
        nv = etree.SubElement(grp, q("p:nvGrpSpPr"))
        etree.SubElement(nv, q("p:cNvPr"), id=str(gid), name=gname)
        etree.SubElement(nv, q("p:cNvGrpSpPr"))
        etree.SubElement(nv, q("p:nvPr"))
        gpr = etree.SubElement(grp, q("p:grpSpPr"))
        xf = etree.SubElement(gpr, q("a:xfrm"))
        etree.SubElement(xf, q("a:off"), x=str(x0), y=str(y0))
        etree.SubElement(xf, q("a:ext"), cx=str(x1 - x0), cy=str(y1 - y0))
        etree.SubElement(xf, q("a:chOff"), x=str(x0), y=str(y0))
        etree.SubElement(xf, q("a:chExt"), cx=str(x1 - x0), cy=str(y1 - y0))
        tree.insert(idxs[0], grp)
        for i, m in enumerate(members):
            tree.remove(m)
            grp.append(m)
            pr = cnvpr(m)
            pr.set("name", "%s %d" % (gname, i + 1))
        ids[gname] = gid
    return ids


class Ids:
    def __init__(self):
        self.n = 0

    def next(self):
        self.n += 1
        return str(self.n)


def effect_xml(ids, spid, fx, node_type, delay):
    """등장 효과 하나. fx: fade | rise"""
    preset = "10" if fx == "fade" else "47"
    tgt = '<p:tgtEl><p:spTgt spid="%s"/></p:tgtEl>' % spid
    parts = [
        '<p:par><p:cTn id="%s" presetID="%s" presetClass="entr" presetSubtype="0" fill="hold" nodeType="%s">'
        % (ids.next(), preset, node_type),
        '<p:stCondLst><p:cond delay="%d"/></p:stCondLst><p:childTnLst>' % delay,
        '<p:set><p:cBhvr><p:cTn id="%s" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>%s'
        '<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr>'
        '<p:to><p:strVal val="visible"/></p:to></p:set>' % (ids.next(), tgt),
        '<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="%s" dur="%d"/>%s</p:cBhvr></p:animEffect>'
        % (ids.next(), DUR, tgt),
    ]
    if fx == "rise":
        for attr, frm in (("ppt_x", "#ppt_x"), ("ppt_y", "#ppt_y+" + RISE)):
            parts.append(
                '<p:anim calcmode="lin" valueType="num"><p:cBhvr additive="base">'
                '<p:cTn id="%s" dur="%d" decel="100000" fill="hold"/>%s'
                '<p:attrNameLst><p:attrName>%s</p:attrName></p:attrNameLst></p:cBhvr>'
                '<p:tavLst><p:tav tm="0"><p:val><p:strVal val="%s"/></p:val></p:tav>'
                '<p:tav tm="100000"><p:val><p:strVal val="#%s"/></p:val></p:tav></p:tavLst></p:anim>'
                % (ids.next(), DUR, tgt, attr, frm, attr)
            )
    parts.append("</p:childTnLst></p:cTn></p:par>")
    return "".join(parts)


def timing_xml(clicks, gids):
    ids = Ids()
    root_id, seq_id = ids.next(), ids.next()
    body = []
    for click in clicks:
        outer, inner = ids.next(), ids.next()
        effs = []
        for i, e in enumerate(click):
            if e["g"] not in gids:
                raise SystemExit("애니메이션 대상 그룹 없음: %s" % e["g"])
            effs.append(effect_xml(ids, gids[e["g"]], e.get("fx", "rise"), "clickEffect" if i == 0 else "withEffect", int(e.get("delay", 0))))
        body.append(
            '<p:par><p:cTn id="%s" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>'
            '<p:par><p:cTn id="%s" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>%s'
            "</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>" % (outer, inner, "".join(effs))
        )
    xml = (
        '<p:timing xmlns:p="%s"><p:tnLst><p:par><p:cTn id="%s" dur="indefinite" restart="never" nodeType="tmRoot">'
        '<p:childTnLst><p:seq concurrent="1" nextAc="seek"><p:cTn id="%s" dur="indefinite" nodeType="mainSeq">'
        "<p:childTnLst>%s</p:childTnLst></p:cTn>"
        '<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
        '<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>'
        "</p:seq></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>" % (NS["p"], root_id, seq_id, "".join(body))
    )
    return etree.fromstring(xml)


def set_tail(root, transition, timing):
    """<p:sld> 자식 순서: cSld, clrMapOvr, transition, timing, extLst"""
    for tag in ("p:transition", "p:timing"):
        for old in root.findall(tag, NS):
            root.remove(old)
    anchor = root.find("p:clrMapOvr", NS)
    if anchor is None:
        anchor = root.find("p:cSld", NS)
    pos = root.index(anchor) + 1
    root.insert(pos, transition)
    if timing is not None:
        root.insert(pos + 1, timing)


def patch_theme(xml, theme):
    root = etree.fromstring(xml)
    scheme = root.find(".//a:clrScheme", NS)
    scheme.set("name", theme["name"])
    for key, hexv in theme["colors"].items():
        el = scheme.find("a:%s" % key, NS)
        if el is None:
            continue
        for ch in list(el):
            el.remove(ch)
        etree.SubElement(el, q("a:srgbClr"), val=hexv)
    fs = root.find(".//a:fontScheme", NS)
    fs.set("name", theme["name"])
    for kind in ("a:majorFont", "a:minorFont"):
        f = fs.find(kind, NS)
        for tag in ("a:latin", "a:ea"):
            el = f.find(tag, NS)
            if el is not None:
                el.set("typeface", theme["font"])
        # 한글 스크립트 지정 글꼴
        for el in f.findall("a:font", NS):
            if el.get("script") == "Hang":
                el.set("typeface", theme["font"])
    root.find(".//a:themeElements/..", NS)
    root.set("name", theme["name"])
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def main():
    cfg = json.load(open(CFG, encoding="utf-8"))
    anims = {int(k): v for k, v in cfg["anims"].items()}
    zin = zipfile.ZipFile(SRC)
    os.makedirs(os.path.dirname(os.path.abspath(DST)), exist_ok=True)
    zout = zipfile.ZipFile(DST, "w", zipfile.ZIP_DEFLATED)
    total_clicks = 0
    for item in zin.infolist():
        if item.filename.endswith("/"):
            continue  # 폴더 항목은 일부 클라우드 서비스에서 열기 오류를 일으킴
        data = zin.read(item.filename)
        m = re.match(r"ppt/slides/slide(\d+)\.xml$", item.filename)
        if m:
            no = int(m.group(1))
            root = etree.fromstring(data)
            gids = group_shapes(root)
            clicks = anims.get(no, [])
            timing = timing_xml(clicks, gids) if clicks else None
            transition = etree.fromstring('<p:transition xmlns:p="%s" spd="fast"><p:fade/></p:transition>' % NS["p"])
            set_tail(root, transition, timing)
            total_clicks += len(clicks)
            print("슬라이드 %2d: 그룹 %2d개 · 클릭 %d회" % (no, len(gids), len(clicks)))
            data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
        elif re.match(r"ppt/theme/theme\d+\.xml$", item.filename):
            data = patch_theme(data, cfg["theme"])
        zout.writestr(item, data)
    zout.close()
    print("총 클릭 %d회 → %s" % (total_clicks, os.path.normpath(DST)))


if __name__ == "__main__":
    main()
