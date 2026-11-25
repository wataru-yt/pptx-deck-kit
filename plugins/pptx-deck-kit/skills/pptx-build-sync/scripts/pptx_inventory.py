#!/usr/bin/env python3
"""pptx の全図形を、座標・書式つきの短いテキストに書き出す（build.js と pptx の突き合わせ用）。

使い方:
  python3 pptx_inventory.py deck.pptx                 # 全スライド
  python3 pptx_inventory.py deck.pptx --slides 3,7    # 指定ページだけ
  python3 pptx_inventory.py deck.pptx --names         # 図形名も出す

出力の1行 = 1図形（<sp> 図形・<pic> 画像・<graphicFrame> 表/グラフ）。座標はインチ。
文字は ¶ 行に [サイズ,B(太字),色] 文字 の形で出る。表は列幅・行高・セルの文字、
グラフは系列名・値・色・点ごとの透明度、ノートとハイパーリンク先も出す。
Python 3.9 以上、標準ライブラリ + lxml のみ。
"""
import argparse
import posixpath
import re
import sys
import zipfile

from lxml import etree

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}
EMU = 914400.0
R_ID = "{%s}id" % NS["r"]
R_EMBED = "{%s}embed" % NS["r"]


def inch(v):
    return round(int(v) / EMU, 3)


def color_of(el):
    if el is None:
        return None
    fill = el.find("a:solidFill", NS)
    if fill is None:
        return "none" if el.find("a:noFill", NS) is not None else None
    c = fill[0]
    out = c.get("val")
    alpha = c.find("a:alpha", NS)
    if alpha is not None:
        out += "@a%g" % (int(alpha.get("val")) / 1000)
    return out


def line_info(ln):
    if ln is None:
        return ""
    s = "ln=%s" % color_of(ln)
    if ln.get("w"):
        s += "/%gpt" % (int(ln.get("w")) / 12700)
    dash = ln.find("a:prstDash", NS)
    if dash is not None:
        s += "/" + dash.get("val")
    for tag in ("headEnd", "tailEnd"):
        e = ln.find("a:" + tag, NS)
        if e is not None and e.get("type") != "none":
            s += "/%s:%s" % (tag, e.get("type"))
    return s


def text_info(tx, links):
    if tx is None:
        return ""
    body = tx.find("a:bodyPr", NS)
    attrs = []
    if body is not None:
        for k in ("anchor", "lIns", "tIns", "rIns", "bIns", "wrap", "vert"):
            if body.get(k) is not None:
                attrs.append("%s=%s" % (k, body.get(k)))
    out = ""
    if attrs:
        out += " body{" + " ".join(attrs) + "}"
    for p in tx.findall("a:p", NS):
        ppr = p.find("a:pPr", NS)
        pa = []
        if ppr is not None:
            if ppr.get("algn"):
                pa.append("al=" + ppr.get("algn"))
            ls = ppr.find("a:lnSpc/a:spcPct", NS)
            if ls is not None:
                pa.append("ls=%g" % (int(ls.get("val")) / 1000))
            sa = ppr.find("a:spcAft/a:spcPts", NS)
            if sa is not None:
                pa.append("sa=%g" % (int(sa.get("val")) / 100))
            if ppr.find("a:buChar", NS) is not None:
                pa.append("bullet")
        runs = []
        for r in p:
            tag = etree.QName(r).localname
            if tag == "br":
                runs.append("⏎")
                continue
            if tag not in ("r", "fld"):
                continue
            rpr = r.find("a:rPr", NS)
            t = r.find("a:t", NS).text or ""
            a = []
            if rpr is not None:
                if rpr.get("sz"):
                    a.append("%g" % (int(rpr.get("sz")) / 100))
                if rpr.get("b") == "1":
                    a.append("B")
                if rpr.get("i") == "1":
                    a.append("I")
                if rpr.get("u"):
                    a.append("U")
                c = color_of(rpr)
                if c:
                    a.append(c)
                if rpr.get("spc"):
                    a.append("spc" + rpr.get("spc"))
                h = rpr.find("a:hlinkClick", NS)
                if h is not None:
                    a.append("LINK:" + links.get(h.get(R_ID), "?"))
            runs.append("[%s]%s" % (",".join(a), t))
        out += "\n      ¶%s %s" % ("(" + " ".join(pa) + ")" if pa else "", "".join(runs))
    return out


def chart_summary(zf, path):
    root = etree.fromstring(zf.read(path))
    lines = []
    for bc in root.iter("{%s}barChart" % NS["c"]):
        d = bc.find("c:barDir", NS).get("val")
        g = bc.find("c:grouping", NS).get("val")
        lines.append("   chart bar dir=%s grouping=%s" % (d, g))
        for s in bc.findall("c:ser", NS):
            name = "".join(s.xpath("c:tx//c:v/text()", namespaces=NS))
            vals = [v.text for v in s.findall("c:val//c:v", NS)]
            sp = s.find("c:spPr", NS)
            col = color_of(sp) if sp is not None else None
            fades = []
            for dp in s.findall("c:dPt", NS):
                a = dp.find(".//a:alpha", NS)
                if a is not None:
                    fades.append("%s@a%g" % (dp.find("c:idx", NS).get("val"), int(a.get("val")) / 1000))
            lines.append("     ser '%s' color=%s vals=%s%s" % (name, col, vals, " dPt=" + ",".join(fades) if fades else ""))
        cats = [v.text for v in bc.find("c:ser", NS).findall("c:cat//c:v", NS)]
        lines.append("     cats=%s" % cats)
    for ax in root.iter("{%s}valAx" % NS["c"]):
        mx = ax.find("c:scaling/c:max", NS)
        mn = ax.find("c:scaling/c:min", NS)
        lines.append("     valAx min=%s max=%s" % (mn.get("val") if mn is not None else None, mx.get("val") if mx is not None else None))
    title = "".join(root.xpath("//c:chart/c:title//a:t/text()", namespaces=NS))
    lines.append("     title=%s" % (title or None))
    return lines


def shape_lines(sp, zf, rels, links, base_dir, depth=0, names=False):
    tag = etree.QName(sp).localname
    pre = "  " * depth
    nv = sp.find(".//p:cNvPr", NS)
    name = nv.get("name") if nv is not None else ""
    sppr = sp.find("p:spPr", NS) if tag != "grpSp" else sp.find("p:grpSpPr", NS)
    xf = sp.find("p:xfrm", NS) if tag == "graphicFrame" else (sppr.find("a:xfrm", NS) if sppr is not None else None)
    geo = ""
    if xf is not None:
        off, ext = xf.find("a:off", NS), xf.find("a:ext", NS)
        geo = "x=%s y=%s w=%s h=%s" % (inch(off.get("x")), inch(off.get("y")), inch(ext.get("cx")), inch(ext.get("cy")))
        if xf.get("rot"):
            geo += " rot=%g" % (int(xf.get("rot")) / 60000)
        if xf.get("flipH") == "1":
            geo += " flipH"
        if xf.get("flipV") == "1":
            geo += " flipV"
    desc = ""
    if sppr is not None:
        pg = sppr.find("a:prstGeom", NS)
        if pg is not None:
            desc += " " + pg.get("prst")
        f = color_of(sppr)
        if f:
            desc += " fill=" + f
        ln = sppr.find("a:ln", NS)
        if ln is not None:
            desc += " " + line_info(ln)
        if sppr.find("a:effectLst/a:outerShdw", NS) is not None:
            desc += " shadow"
    out = ["%s<%s> %s%s%s%s" % (pre, tag, ("'%s' " % name) if names else "", geo, desc, text_info(sp.find("p:txBody", NS), links))]
    if tag == "pic":
        blip = sp.find(".//a:blip", NS)
        out.append(pre + "   img " + (posixpath.basename(rels.get(blip.get(R_EMBED), "")) if blip is not None else ""))
    if tag == "graphicFrame":
        tbl = sp.find(".//a:tbl", NS)
        if tbl is not None:
            out.append(pre + "   tbl cols %s" % [inch(g.get("w")) for g in tbl.findall("a:tblGrid/a:gridCol", NS)])
            for tr in tbl.findall("a:tr", NS):
                out.append(pre + "   row h=%s" % inch(tr.get("h")))
                for tc in tr.findall("a:tc", NS):
                    tp = tc.find("a:tcPr", NS)
                    ex = ""
                    if tp is not None:
                        ex = " " + (color_of(tp) or "")
                        bl = []
                        for b in ("lnL", "lnR", "lnT", "lnB"):
                            e = tp.find("a:" + b, NS)
                            if e is not None:
                                bl.append(b + ":" + line_info(e).replace("ln=", ""))
                        ex += " " + " ".join(bl)
                    out.append(pre + "     cell" + ex + text_info(tc.find("a:txBody", NS), links))
        ch = sp.find(".//c:chart", NS)
        if ch is not None:
            target = rels.get(ch.get(R_ID))
            if target:
                out.extend(chart_summary(zf, target))
    if tag == "grpSp":
        for c in sp:
            if etree.QName(c).localname in ("sp", "pic", "grpSp", "graphicFrame", "cxnSp"):
                out.extend(shape_lines(c, zf, rels, links, base_dir, depth + 1, names))
    return out


def read_rels(zf, part):
    d, f = posixpath.split(part)
    rp = posixpath.join(d, "_rels", f + ".rels")
    rels, ext = {}, {}
    if rp in zf.namelist():
        for r in etree.fromstring(zf.read(rp)):
            t = r.get("Target")
            if r.get("TargetMode") == "External":
                ext[r.get("Id")] = t
            else:
                rels[r.get("Id")] = posixpath.normpath(posixpath.join(d, t)) if not t.startswith("/") else t.lstrip("/")
    return rels, ext


def slide_order(zf):
    pres = etree.fromstring(zf.read("ppt/presentation.xml"))
    rels, _ = read_rels(zf, "ppt/presentation.xml")
    return [rels[s.get(R_ID)] for s in pres.find("p:sldIdLst", NS)]


def inventory(path, names=False):
    """[(スライド番号, [行...]), ...] を返す。"""
    zf = zipfile.ZipFile(path)
    result = []
    for i, part in enumerate(slide_order(zf), 1):
        rels, ext = read_rels(zf, part)
        links = dict(ext)
        root = etree.fromstring(zf.read(part))
        lines = []
        for c in root.find(".//p:spTree", NS):
            if etree.QName(c).localname in ("sp", "pic", "grpSp", "graphicFrame", "cxnSp"):
                for x in shape_lines(c, zf, rels, links, posixpath.dirname(part), 0, names):
                    lines.extend(x.split("\n"))
        for k, v in rels.items():
            if "/notesSlides/" in v:
                notes = etree.fromstring(zf.read(v))
                txt = "\n".join("".join(t.text or "" for t in p.iter("{%s}t" % NS["a"])) for p in notes.iter("{%s}p" % NS["a"]))
                txt = txt.strip()
                if txt:
                    lines.append("NOTES " + txt.replace("\n", " / "))
        result.append((i, lines))
    return result


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("pptx")
    ap.add_argument("--slides", help="例: 3,7,10")
    ap.add_argument("--names", action="store_true")
    a = ap.parse_args()
    only = set(int(x) for x in a.slides.split(",")) if a.slides else None
    for n, lines in inventory(a.pptx, a.names):
        if only and n not in only:
            continue
        print("===== SLIDE %d =====" % n)
        print("\n".join(lines))


if __name__ == "__main__":
    sys.exit(main())
