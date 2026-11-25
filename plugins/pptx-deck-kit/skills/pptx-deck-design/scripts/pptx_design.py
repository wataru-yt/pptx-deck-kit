#!/usr/bin/env python3
"""会社の pptx の見た目（レイアウトの背景と飾り、題などの置き場所）を読む。theme_from_pptx.py から呼ぶ。

考え方: 飾りは図形の種類が多く（グラデーション、自由な形、写真、模様）、1つずつ写すと崩れる。
そこで、表紙・章扉・本文のレイアウトを「文字を入れていない1枚」にして画像に書き出し、背景として使う。
題・リードなどの文字は、部品が本物の文字で書く。置き場所・大きさ・色・そろえ方は、XML から読む。
画像にするのに PowerPoint か LibreOffice、PDF を PNG にするのに PyMuPDF（pip install pymupdf）か pdftoppm が要る。
"""
import os
import posixpath
import re
import shutil
import subprocess
import tempfile
import zipfile
import xml.etree.ElementTree as ET

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
}
for k, v in NS.items():
    if k in ("a", "p", "r"):
        ET.register_namespace(k, v)
R_ID = "{%s}id" % NS["r"]
W_IN, H_IN, M_IN = 13.333, 7.5, 0.6
BODY_TOP, BODY_BOTTOM = 2.2, 6.6
KINDS = ("cover", "chapter", "content")
KIND_JA = {"cover": "表紙", "chapter": "章扉", "content": "本文"}
TYPE_OF = {"cover": ("title",), "chapter": ("secHead",), "content": ("obj", "titleOnly", "tx")}
NAME_OF = {
    "cover": r"表紙|タイトル ?スライド|title ?slide|^cover",
    "chapter": r"章|中扉|セクション|区切り|section|chapter|divider",
    "content": r"本文|コンテンツ|タイトルと|タイトルのみ|title and content|title only|content|body",
}


def q(tag):
    p, t = tag.split(":")
    return "{%s}%s" % (NS[p], t)


def rels_path(part):
    d, b = posixpath.split(part)
    return posixpath.join(d, "_rels", b + ".rels")


def rels(z, part):
    rp = rels_path(part)
    if rp not in z.namelist():
        return {}
    d = posixpath.dirname(part)
    out = {}
    for rel in ET.fromstring(z.read(rp)).findall("rel:Relationship", NS):
        if rel.get("TargetMode") == "External":
            continue
        t = rel.get("Target")
        out[rel.get("Id")] = (t.lstrip("/") if t.startswith("/") else posixpath.normpath(posixpath.join(d, t)), rel.get("Type", "").rsplit("/", 1)[-1])
    return out


def related(z, part, kind):
    for full, typ in rels(z, part).values():
        if typ == kind:
            return full
    return None


class Deck:
    def __init__(self, path):
        self.path = path
        self.z = zipfile.ZipFile(path)
        self.pres = "ppt/presentation.xml"
        root = ET.fromstring(self.z.read(self.pres))
        sz = root.find("p:sldSz", NS)
        self.cx, self.cy = int(sz.get("cx")), int(sz.get("cy"))
        self.scale = W_IN / (self.cx / 914400)  # 元のスライドの大きさから 13.333in の幅へ
        prels = rels(self.z, self.pres)
        self.masters = [prels[m.get(R_ID)][0] for m in root.findall("p:sldMasterIdLst/p:sldMasterId", NS)]
        self.slides = [prels[s.get(R_ID)][0] for s in root.findall("p:sldIdLst/p:sldId", NS)]
        self.layouts = []
        for m in self.masters:
            mroot = ET.fromstring(self.z.read(m))
            mrels = rels(self.z, m)
            for el in mroot.findall("p:sldLayoutIdLst/p:sldLayoutId", NS):
                part = mrels[el.get(R_ID)][0]
                lroot = ET.fromstring(self.z.read(part))
                csld = lroot.find("p:cSld", NS)
                self.layouts.append({"part": part, "master": m, "type": lroot.get("type"), "name": (csld.get("name") if csld is not None else "") or ""})

    def pick(self, overrides):
        """表紙・章扉・本文に当てるレイアウトを選ぶ"""
        used = [related(self.z, s, "slideLayout") for s in self.slides]
        out = {}
        for kind in KINDS:
            want = overrides.get(kind)
            cand = None
            if want:
                cand = next((l for l in self.layouts if l["name"] == want), None)
                if not cand:
                    raise SystemExit("レイアウト「%s」が無い。使える名前: %s" % (want, "、".join(l["name"] for l in self.layouts)))
            if not cand:
                cand = next((l for t in TYPE_OF[kind] for l in self.layouts if l["type"] == t), None)
            if not cand:
                cand = next((l for l in self.layouts if re.search(NAME_OF[kind], l["name"], re.I)), None)
            if not cand and kind == "content":
                # 題と、大きな本文の置き場所を持つレイアウト（表紙に使うものは除く）
                for l in self.layouts:
                    if out.get("cover") and l["part"] == out["cover"]["part"]:
                        continue
                    phs = self.placeholders(l)
                    if any(b["type"] in ("title", "ctrTitle") and b["y"] < 1.2 for b in phs) and any(b["type"] in ("body", "obj") and b["h"] >= 3.0 for b in phs):
                        cand = l
                        break
            if not cand:
                idx = {"cover": 0, "chapter": 1, "content": 2}[kind]
                if idx < len(used):
                    cand = next((l for l in self.layouts if l["part"] == used[idx]), None)
            if cand:
                out[kind] = cand
        return out

    # ---- スライドに直接描いたデザイン ----
    def slide_source(self, n, kind):
        """レイアウトの代わりに、n 枚目のスライドを飾りと置き場所の元にする"""
        if not 1 <= n <= len(self.slides):
            raise SystemExit("スライド %d は無い（1〜%d）" % (n, len(self.slides)))
        part = self.slides[n - 1]
        lay = related(self.z, part, "slideLayout")
        base = next(l for l in self.layouts if l["part"] == lay)
        return dict(base, slide=part, kind=kind, name="スライド %d" % n)

    def bbox(self, el):
        """図形の位置と大きさ（13.333in の幅にそろえたインチ）。指定が無ければ None"""
        for path in ("p:spPr/a:xfrm", "p:grpSpPr/a:xfrm", "p:xfrm"):
            xf = el.find(path, NS)
            if xf is not None:
                off, ext = xf.find("a:off", NS), xf.find("a:ext", NS)
                if off is not None and ext is not None:
                    k = self.scale / 914400
                    return int(off.get("x")) * k, int(off.get("y")) * k, int(ext.get("cx")) * k, int(ext.get("cy")) * k
        return None

    def walk(self, el, to=lambda b: b):
        """図形を、グループの中まで、スライドの上の位置と一緒にたどる"""
        for child in el:
            tag = child.tag.split("}")[1]
            if tag == "sp":
                b = self.bbox(child)
                yield child, None if b is None else to(b)
            elif tag == "grpSp":
                xf = child.find("p:grpSpPr/a:xfrm", NS)
                inner = to
                if xf is not None:
                    v = {n: xf.find("a:" + n, NS) for n in ("off", "ext", "chOff", "chExt")}
                    if all(e is not None for e in v.values()):
                        k = self.scale / 914400
                        ox, oy = int(v["off"].get("x")) * k, int(v["off"].get("y")) * k
                        cx, cy = int(v["chOff"].get("x")) * k, int(v["chOff"].get("y")) * k
                        sx = int(v["ext"].get("cx")) / max(1, int(v["chExt"].get("cx")))
                        sy = int(v["ext"].get("cy")) / max(1, int(v["chExt"].get("cy")))
                        inner = lambda b, ox=ox, oy=oy, cx=cx, cy=cy, sx=sx, sy=sy, to=to: to((ox + (b[0] - cx) * sx, oy + (b[1] - cy) * sy, b[2] * sx, b[3] * sy))
                yield from self.walk(child, inner)

    def slide_decor(self, lay):
        """飾りだけにしたスライドの XML。文字・表・グラフは外し、本文は上下左右の端にある図形だけ残す"""
        root = ET.fromstring(self.z.read(lay["slide"]))
        tree = root.find("p:cSld/p:spTree", NS)
        H = self.cy / 914400 * self.scale

        def edge(b):
            x, y, w, h = b
            return (y + h <= 1.75 or y >= H - 0.9 or x + w <= 1.0 or x >= W_IN - 1.0
                    or (w >= W_IN * 0.95 and h >= H * 0.95))

        for el in list(tree):
            tag = el.tag.split("}")[1]
            if tag in ("nvGrpSpPr", "grpSpPr"):
                continue
            b = self.bbox(el)
            drop = tag not in ("sp", "pic", "grpSp", "cxnSp") or el.find(".//p:nvPr/p:ph", NS) is not None
            if not drop and lay["kind"] == "content":
                drop = b is None or not edge(b)
            if drop:
                tree.remove(el)
        for body in root.iter(q("p:txBody")):
            for p in body.findall("a:p", NS):
                for child in list(p):
                    if child.tag in (q("a:r"), q("a:fld"), q("a:br")):
                        p.remove(child)
        timing = root.find("p:timing", NS)
        if timing is not None:
            root.remove(timing)
        for prefix, uri in list(NS.items()):
            if prefix != "rel":
                ET.register_namespace(prefix, uri)
        return ET.tostring(root, encoding="utf-8", xml_declaration=True)

    def slide_placeholders(self, lay):
        """スライドの文字の箱から、題・副題・日付・章番号の置き場所を決める"""
        root = ET.fromstring(self.z.read(lay["slide"]))
        master = lay["master"]
        k = self.scale / 914400
        boxes = []
        for sp, b in self.walk(root.find("p:cSld/p:spTree", NS)):
            text = "".join((t.text or "") for t in sp.iter(q("a:t"))).strip()
            if not text:
                continue
            ph = sp.find("p:nvSpPr/p:nvPr/p:ph", NS)
            if b is None and ph is not None:
                src = self.find_ph(lay["part"], ph.get("type") or "body", ph.get("idx")) or self.find_ph(master, ph.get("type") or "body", ph.get("idx"), True)
                b = self.bbox(src) if src is not None else None
            if b is None:
                continue
            rp = sp.find(".//a:r/a:rPr", NS)
            pp = sp.find(".//a:p/a:pPr", NS)
            bp = sp.find("p:txBody/a:bodyPr", NS)
            size = int(rp.get("sz")) / 100 if rp is not None and rp.get("sz") else (32 if ph is not None and ph.get("type") in ("title", "ctrTitle") else 18)
            box = {"x": b[0], "y": b[1], "w": b[2], "h": b[3], "text": text, "size": round(size * self.scale),
                   "bold": rp is not None and rp.get("b") in ("1", "true"),
                   "color": self.color(rp.find("a:solidFill", NS), master) if rp is not None else None,
                   "align": {"ctr": "center", "r": "right"}.get(pp.get("algn") if pp is not None else None, "left"),
                   "valign": {"ctr": "middle", "b": "bottom"}.get(bp.get("anchor") if bp is not None else None, "top"),
                   "is_title": ph is not None and ph.get("type") in ("title", "ctrTitle")}
            boxes.append(box)
        if not boxes:
            return []
        H = self.cy / 914400 * self.scale
        pool = [b for b in boxes if b["y"] < 1.75] if lay["kind"] == "content" else [b for b in boxes if not re.match(r"^[\d０-９.\s]+$", b["text"])]
        title = next((b for b in boxes if b["is_title"]), None) or max(pool or boxes, key=lambda b: (b["size"], -b["y"]))
        out = [dict(title, type="title", idx=None)]
        if lay["kind"] == "cover":
            rest = sorted([b for b in boxes if b is not title and b["y"] > title["y"] and b["y"] < H - 0.3], key=lambda b: b["y"])
            out += [dict(b, type="body", idx=None) for b in rest[:2]]
        elif lay["kind"] == "chapter":
            num = next((b for b in boxes if b is not title and re.match(r"^[\d０-９]{1,2}\.?$", b["text"])), None)
            if num:
                out.append(dict(num, type="body", idx=None))
        return out

    def has_slide_design(self):
        """スライドに、置き場所でない図形や画像がいくつもあるか"""
        count = 0
        for part in self.slides[:8]:
            root = ET.fromstring(self.z.read(part))
            tree = root.find("p:cSld/p:spTree", NS)
            count += sum(1 for el in tree if el.tag.split("}")[1] in ("sp", "pic", "grpSp", "cxnSp") and el.find(".//p:nvPr/p:ph", NS) is None)
        return count >= 6

    # ---- 文字の置き場所 ----
    def theme_of(self, master):
        t = related(self.z, master, "theme")
        colors, fonts = {}, {}
        if t:
            root = ET.fromstring(self.z.read(t))
            cs = root.find(".//a:clrScheme", NS)
            for el in list(cs) if cs is not None else []:
                s, y = el.find("a:srgbClr", NS), el.find("a:sysClr", NS)
                v = s.get("val") if s is not None else (y.get("lastClr") if y is not None else None)
                if v:
                    colors[el.tag.rsplit("}", 1)[-1]] = v.upper()
            for kind in ("major", "minor"):
                f = root.find(".//a:fontScheme/a:%sFont" % kind, NS)
                if f is None:
                    continue
                ea, lat = f.find("a:ea", NS), f.find("a:latin", NS)
                jp = next((x for x in f.findall("a:font", NS) if x.get("script") == "Jpan"), None)
                fonts[kind] = (ea.get("typeface") if ea is not None else "") or (jp.get("typeface") if jp is not None else "") or (lat.get("typeface") if lat is not None else "")
        cmap = ET.fromstring(self.z.read(master)).find("p:clrMap", NS)
        return colors, fonts, dict(cmap.attrib) if cmap is not None else {}

    def color(self, fill, master):
        """a:solidFill → 6桁（schemeClr は clrMap とテーマで解く。lumMod・lumOff を当てる）"""
        if fill is None:
            return None
        colors, _, cmap = self.theme_of(master)
        el = fill.find("a:srgbClr", NS)
        v = el.get("val").upper() if el is not None else None
        if v is None:
            el = fill.find("a:schemeClr", NS)
            if el is None:
                el = fill.find("a:sysClr", NS)
                return el.get("lastClr") if el is not None else None
            name = el.get("val")
            v = colors.get(cmap.get(name, name))
        if v is None:
            return None
        mod, off = el.find("a:lumMod", NS), el.find("a:lumOff", NS)
        if mod is not None or off is not None:
            import colorsys
            h, l, s = colorsys.rgb_to_hls(*[int(v[i:i + 2], 16) / 255 for i in (0, 2, 4)])
            l = min(1, max(0, l * (int(mod.get("val")) / 1e5 if mod is not None else 1) + (int(off.get("val")) / 1e5 if off is not None else 0)))
            v = "".join("%02X" % round(c * 255) for c in colorsys.hls_to_rgb(h, l, s))
        return v

    def shapes(self, part):
        root = ET.fromstring(self.z.read(part))
        out = []
        for sp in root.iter(q("p:sp")):
            ph = sp.find("p:nvSpPr/p:nvPr/p:ph", NS)
            if ph is not None:
                out.append((ph.get("type") or "body", ph.get("idx"), sp))
        return out

    def find_ph(self, part, typ, idx, by_type_only=False):
        if not part:
            return None
        items = self.shapes(part)
        if idx is not None and not by_type_only:
            for t, i, sp in items:
                if i == idx:
                    return sp
        alias = {"ctrTitle": "title", "subTitle": "body", "obj": "body"}
        for t, i, sp in items:
            if alias.get(t, t) == alias.get(typ, typ):
                return sp
        return None

    def placeholders(self, layout):
        """レイアウトの置き場所を、レイアウト → マスターの順に埋める。レイアウトを使うスライドがあれば、その指定を優先する"""
        if layout.get("slide"):
            return self.slide_placeholders(layout)
        master = layout["master"]
        mroot = ET.fromstring(self.z.read(master))
        _, fonts, _ = self.theme_of(master)
        slide = next((s for s in self.slides if related(self.z, s, "slideLayout") == layout["part"]), None)
        out = []
        for typ, idx, sp in self.shapes(layout["part"]):
            msp = self.find_ph(master, typ, idx, by_type_only=True)
            ssp = self.find_ph(slide, typ, idx) if slide else None
            chain = [x for x in (ssp, sp, msp) if x is not None]
            style_name = "titleStyle" if typ in ("title", "ctrTitle") else ("otherStyle" if typ in ("sldNum", "dt", "ftr") else "bodyStyle")
            st = mroot.find("p:txStyles/p:%s/a:lvl1pPr" % style_name, NS)
            box = {"type": typ, "idx": idx}
            for x in chain:
                off, ext = x.find("p:spPr/a:xfrm/a:off", NS), x.find("p:spPr/a:xfrm/a:ext", NS)
                if off is not None and ext is not None:
                    k = self.scale / 914400
                    box.update(x=int(off.get("x")) * k, y=int(off.get("y")) * k, w=int(ext.get("cx")) * k, h=int(ext.get("cy")) * k)
                    break
            if "x" not in box:
                continue
            props = []  # 優先の高い順の (pPr, rPr)
            for x in chain:
                body = x.find("p:txBody", NS)
                if body is None:
                    continue
                run_p = body.find("a:p/a:pPr", NS)
                run_r = body.find("a:p/a:r/a:rPr", NS)
                if run_r is None:
                    run_r = body.find("a:p/a:fld/a:rPr", NS)
                if x is ssp:
                    props.append((run_p, run_r))
                lvl = body.find("a:lstStyle/a:lvl1pPr", NS)
                if lvl is not None:
                    props.append((lvl, lvl.find("a:defRPr", NS)))
                if x is not ssp:
                    props.append((run_p, run_r))
            if st is not None:
                props.append((st, st.find("a:defRPr", NS)))

            def first(get):
                for pp, rp in props:
                    v = get(pp, rp)
                    if v is not None:
                        return v
                return None

            size = first(lambda pp, rp: rp.get("sz") if rp is not None else None)
            if size:
                box["size"] = round(int(size) / 100 * self.scale)
            bold = first(lambda pp, rp: rp.get("b") if rp is not None else None)
            box["bold"] = bold in ("1", "true")
            algn = first(lambda pp, rp: pp.get("algn") if pp is not None else None)
            box["align"] = {"ctr": "center", "r": "right"}.get(algn, "left")
            col = first(lambda pp, rp: self.color(rp.find("a:solidFill", NS), master) if rp is not None else None)
            if col:
                box["color"] = col

            def face(pp, rp):
                if rp is None:
                    return None
                for tag in ("a:ea", "a:latin"):
                    el = rp.find(tag, NS)
                    if el is not None and el.get("typeface"):
                        t = el.get("typeface")
                        if t.startswith("+mj"):
                            return fonts.get("major")
                        if t.startswith("+mn"):
                            return fonts.get("minor")
                        return t
                return None

            f = first(face)
            if f:
                box["font"] = f
            anchor = None
            for x in chain:
                bp = x.find("p:txBody/a:bodyPr", NS)
                if bp is not None and bp.get("anchor"):
                    anchor = bp.get("anchor")
                    break
            box["valign"] = {"ctr": "middle", "b": "bottom"}.get(anchor, "top")
            if typ == "sldNum":
                txt = "".join((t.text or "") if el.tag == q("a:r") else "{n}"
                              for p in (sp.findall("p:txBody/a:p", NS)) for el in p if el.tag in (q("a:r"), q("a:fld"))
                              for t in [el.find("a:t", NS)] if t is not None)
                box["format"] = txt if "{n}" in txt else "{n}"
            out.append(box)
        return out

    def static_texts(self, layout):
        """レイアウトとマスターに書いてある固定の文字（背景の画像に焼き込まれる）"""
        out = []
        for part in ([] if layout.get("slide") else [layout["master"], layout["part"]]):
            for sp in ET.fromstring(self.z.read(part)).iter(q("p:sp")):
                if sp.find("p:nvSpPr/p:nvPr/p:ph", NS) is not None:
                    continue
                t = "".join(x.text or "" for x in sp.iter(q("a:t"))).strip()
                if t:
                    out.append(t)
        return out

    # ---- 固定の図形：章ナビと、ページ番号の前後の文字 ----
    def static_shapes(self, layout):
        """レイアウトとマスターの固定の図形（置き場所でないもの）。位置・文字・色"""
        out = []
        master = layout["master"]
        k = self.scale / 914400
        for part in ([] if layout.get("slide") else [master, layout["part"]]):
            for sp in ET.fromstring(self.z.read(part)).iter(q("p:sp")):
                if sp.find("p:nvSpPr/p:nvPr/p:ph", NS) is not None:
                    continue
                off, ext = sp.find("p:spPr/a:xfrm/a:off", NS), sp.find("p:spPr/a:xfrm/a:ext", NS)
                if off is None or ext is None:
                    continue
                rp = sp.find(".//a:r/a:rPr", NS)
                size = rp.get("sz") if rp is not None else None
                out.append({
                    "part": part, "id": (sp.find("p:nvSpPr/p:cNvPr", NS).get("id")),
                    "x": int(off.get("x")) * k, "y": int(off.get("y")) * k, "w": int(ext.get("cx")) * k, "h": int(ext.get("cy")) * k,
                    "text": "".join(x.text or "" for x in sp.iter(q("a:t"))).strip(),
                    "fill": self.color(sp.find("p:spPr/a:solidFill", NS), master),
                    "color": self.color(rp.find("a:solidFill", NS), master) if rp is not None else None,
                    "bold": rp is not None and rp.get("b") in ("1", "true"),
                    "size": round(int(size) / 100 * self.scale) if size else None,
                })
        return out

    def find_nav(self, layout):
        """同じ高さに短い文字が3つ以上並んでいれば章ナビとみなす。(nav の設定, 外す図形, 項目名) を返す"""
        shapes = self.static_shapes(layout)
        texts = [s for s in shapes if s["text"] and len(s["text"]) <= 16 and s["h"] <= 0.7 and s["y"] < 1.4]
        rows = {}
        for s in texts:
            rows.setdefault(round(s["y"] * 10), []).append(s)
        row = max(rows.values(), key=len, default=[])
        if len(row) < 3:
            return None, [], []
        row.sort(key=lambda s: s["x"])
        x0, x1 = row[0]["x"], row[-1]["x"] + row[-1]["w"]
        y0, h = row[0]["y"], row[0]["h"]

        def odd(key):
            vals = [s[key] for s in row]
            common = max(set(vals), key=vals.count)
            rare = next((v for v in vals if v != common), None)
            return common, rare

        near = [s for s in shapes if s not in row and not s["text"] and s["h"] <= 0.2 and y0 - 0.1 <= s["y"] <= y0 + h + 0.25 and s["x"] >= x0 - 0.1 and s["x"] + s["w"] <= x1 + 0.1]
        idle_color, act_color = odd("color")
        idle_fill, act_fill = odd("fill")
        lines = sorted(near, key=lambda s: s["x"])
        nav = {"x": r2(x0), "y": r2(y0), "w": r2(x1 - x0), "h": r2(h), "align": "justify", "size": min(18, max(11, row[0]["size"] or 12))}
        if any(s["fill"] for s in row):
            nav["style"] = "tabs"
            if idle_fill:
                nav["fill"] = idle_fill
            if act_fill:
                nav["activeFill"] = act_fill
        else:
            nav["style"] = "underline"
            fills = [s["fill"] for s in lines if s["fill"]]
            if fills:
                common = max(set(fills), key=fills.count)
                nav["fill"] = common
                nav["activeFill"] = next((f for f in fills if f != common), act_color or common)
        if idle_color:
            nav["color"] = idle_color
        if act_color:
            nav["activeColor"] = act_color
        if all(re.match(r"^\d", s["text"]) for s in row):
            nav["numbers"] = True
        return nav, row + near, [re.sub(r"^\d+[\s　.．]*", "", s["text"]) for s in row]

    def number_affixes(self, layout, num):
        """ページ番号の置き場所の前後にある固定の文字（− ・ P. ・ / 12）。(外す図形, 番号の形) を返す"""
        found = []
        for s in self.static_shapes(layout):
            t = s["text"]
            if not t or not re.match(r"^(?:[−–—―-]|[Pp]\.?|/\s*\d+)$", t):
                continue
            if min(s["y"] + s["h"], num["y"] + num["h"]) - max(s["y"], num["y"]) <= 0.05:
                continue
            if max(s["x"], num["x"]) - min(s["x"] + s["w"], num["x"] + num["w"]) > 0.6:
                continue
            found.append(s)
        if not found:
            return [], num.get("format", "{n}")
        center = num["x"] + num["w"] / 2
        before = "".join(s["text"] for s in sorted(found, key=lambda s: s["x"]) if s["x"] + s["w"] / 2 < center)
        after = "".join(s["text"] for s in sorted(found, key=lambda s: s["x"]) if s["x"] + s["w"] / 2 >= center)
        fmt = num.get("format", "{n}")
        if before:
            fmt = (before + " " if re.match(r"^[−–—―-]$|\.$", before) else before) + fmt
        if after:
            fmt = fmt + (" / {total}" if re.match(r"^/\s*\d+$", after) else " " + after)
        return found, fmt

    # ---- レイアウトだけの1枚を画像にする ----
    def render_backgrounds(self, picked, outdir, drop=(), render_script=None, remove=()):
        tmp = tempfile.mkdtemp(prefix="pptx-design-")
        try:
            stripped = os.path.join(tmp, "layouts.pptx")
            self.write_blank(picked, stripped, drop, remove)
            subprocess.run(["zsh" if shutil.which("zsh") else "bash", render_script, stripped, tmp, "70"], stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
            pdf = os.path.join(tmp, "layouts.pdf")
            if not os.path.exists(pdf):
                raise SystemExit("画像にできない（PowerPoint か LibreOffice が要る）。--no-design で、色・書体・ロゴだけ読み取れる")
            pages = rasterize(pdf, tmp)
            out = {}
            for i, kind in enumerate([k for k in KINDS if k in picked]):
                dst = os.path.join(outdir, "bg-%s.png" % kind)
                shutil.copyfile(pages[i], dst)
                out[kind] = dst
            return out
        finally:
            shutil.rmtree(tmp, ignore_errors=True)

    def write_blank(self, picked, dst, drop, remove=()):
        z = self.z
        kinds = [k for k in KINDS if k in picked]
        parts = {k: "ppt/slides/designblank%d.xml" % i for i, k in enumerate(kinds, 1)}
        blank = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld xmlns:a="%s" xmlns:r="%s" xmlns:p="%s"><p:cSld><p:spTree>'
                 '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld></p:sld>') % (NS["a"], NS["r"], NS["p"])
        pres = z.read(self.pres).decode("utf-8")
        ids = "".join('<p:sldId id="%d" r:id="rIdDesign%d"/>' % (5000 + i, i) for i in range(1, len(kinds) + 1))
        if re.search(r"<p:sldIdLst\s*/>", pres):
            pres = re.sub(r"<p:sldIdLst\s*/>", "<p:sldIdLst>%s</p:sldIdLst>" % ids, pres)
        elif "<p:sldIdLst" in pres:
            pres = re.sub(r"<p:sldIdLst>.*?</p:sldIdLst>", "<p:sldIdLst>%s</p:sldIdLst>" % ids, pres, flags=re.S)
        else:
            pres = pres.replace("<p:sldSz", "<p:sldIdLst>%s</p:sldIdLst><p:sldSz" % ids)
        prels = z.read(rels_path(self.pres)).decode("utf-8")
        prels = re.sub(r"<Relationship\b[^>]*relationships/slide\"[^>]*/>", "", prels)
        prels = prels.replace("</Relationships>", "".join(
            '<Relationship Id="rIdDesign%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/designblank%d.xml"/>' % (i, i)
            for i in range(1, len(kinds) + 1)) + "</Relationships>")
        ct = z.read("[Content_Types].xml").decode("utf-8")
        ct = re.sub(r'<Override\b[^>]*PartName="/ppt/slides/slide\d+\.xml"[^>]*/>', "", ct)
        # .potx はテンプレートの種類で書かれているので、プレゼンテーションの種類にする
        ct = ct.replace("presentationml.template.main+xml", "presentationml.presentation.main+xml")
        ct = ct.replace("</Types>", "".join('<Override PartName="/%s" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>' % p for p in parts.values()) + "</Types>")
        old = set(self.slides)
        old |= {rels_path(s) for s in self.slides}
        masters_layouts = set(self.masters) | {l["part"] for l in self.layouts}
        with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as out:
            for it in z.infolist():
                if it.filename in old:
                    continue
                data = z.read(it.filename)
                if it.filename == self.pres:
                    data = pres.encode("utf-8")
                elif it.filename == rels_path(self.pres):
                    data = prels.encode("utf-8")
                elif it.filename == "[Content_Types].xml":
                    data = ct.encode("utf-8")
                elif it.filename in masters_layouts:
                    ids = {i for p, i in remove if p == it.filename}
                    if drop or ids:
                        data = drop_shapes(data, drop, ids)
                out.writestr(it.filename, data)
            for kind in kinds:
                p = parts[kind]
                src = picked[kind].get("slide")
                if src:
                    # 飾りだけにした元のスライドを、画像などの関係ごと写す（ノートの関係は外す）
                    out.writestr(p, self.slide_decor(picked[kind]))
                    rels_xml = z.read(rels_path(src)).decode("utf-8")
                    out.writestr(rels_path(p), re.sub(r"<Relationship\b[^>]*relationships/notesSlide\"[^>]*/>", "", rels_xml))
                    continue
                out.writestr(p, blank)
                target = posixpath.relpath(picked[kind]["part"], "ppt/slides")
                out.writestr(rels_path(p), '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="%s"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="%s"/></Relationships>' % (NS["rel"], target))


def drop_shapes(data, drop, ids=()):
    """固定の文字が drop のどれかと同じ図形と、ids の図形を、マスター・レイアウトから外す（章ナビなど、部品が自分で描くもの）"""
    xml = data.decode("utf-8")

    def keep(m):
        if "<p:ph" in m.group(0):
            return m.group(0)
        t = "".join(re.findall(r"<a:t>([^<]*)</a:t>", m.group(0))).strip()
        i = re.search(r'<p:cNvPr[^>]*?\bid="(\d+)"', m.group(0))
        return "" if (t and t in drop) or (i and i.group(1) in ids) else m.group(0)

    return re.sub(r"<p:sp>.*?</p:sp>", keep, xml, flags=re.S).encode("utf-8")


def rasterize(pdf, outdir, width=1920):
    pages = []
    try:
        import fitz
        doc = fitz.open(pdf)
        for i, page in enumerate(doc):
            m = width / page.rect.width
            p = os.path.join(outdir, "page%02d.png" % i)
            page.get_pixmap(matrix=fitz.Matrix(m, m), alpha=False).save(p)
            pages.append(p)
        return pages
    except ImportError:
        if not shutil.which("pdftoppm"):
            raise SystemExit("PDF を PNG にできない。pip install pymupdf か、poppler（pdftoppm）を入れる")
        subprocess.run(["pdftoppm", "-png", "-scale-to-x", str(width), "-scale-to-y", "-1", pdf, os.path.join(outdir, "page")], check=True)
        return sorted(os.path.join(outdir, f) for f in os.listdir(outdir) if re.match(r"page-\d+\.png$", f))


# ---- 画像から読む（PIL があれば） ----
def lum(rgb):
    def ch(c):
        c /= 255
        return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4
    return 0.2126 * ch(rgb[0]) + 0.7152 * ch(rgb[1]) + 0.0722 * ch(rgb[2])


def analyze(path):
    """本文の地の明るさと、下の飾りが始まる高さ"""
    from PIL import Image
    im = Image.open(path).convert("RGB").resize((666, 375))
    px = im.load()
    k = 666 / W_IN
    area = [px[x, y] for y in range(int(BODY_TOP * k), int(5.6 * k), 3) for x in range(int(M_IN * k), int((W_IN - M_IN) * k), 3)]
    base = sorted(area, key=lum)[len(area) // 2]
    base_l = lum(base)
    bottom = None
    for y in range(int(5.5 * k), int(7.2 * k)):
        row = [px[x, y] for x in range(int(M_IN * k), int((W_IN - M_IN) * k), 2)]
        odd = sum(1 for p in row if (max(lum(p), base_l) + 0.05) / (min(lum(p), base_l) + 0.05) > 1.12) / len(row)
        if odd > 0.04:
            bottom = y / k
            break
    busy = sum(1 for p in area if (max(lum(p), base_l) + 0.05) / (min(lum(p), base_l) + 0.05) > 1.25) / len(area)
    head = None  # 上の帯の下端：幅の大半が地と違う行の、いちばん下
    for y in range(0, int(2.0 * k)):
        row = [px[x, y] for x in range(int(M_IN * k), int((W_IN - M_IN) * k), 2)]
        if sum(1 for p in row if (max(lum(p), base_l) + 0.05) / (min(lum(p), base_l) + 0.05) > 1.5) / len(row) > 0.6:
            head = (y + 1) / k
    return {"base": "%02X%02X%02X" % base, "dark": base_l < 0.6, "decor_from": bottom, "busy": busy, "head_to": head}


def is_blank(path):
    """画像がほぼ無地の白か"""
    from PIL import Image
    px = list(Image.open(path).convert("L").resize((160, 90)).getdata())
    return sum(1 for v in px if v < 245) / len(px) < 0.005


def contrast_on(path, box, color):
    """置き場所の左上寄りの地と、文字の色の明暗比（中央値）"""
    from PIL import Image
    im = Image.open(path).convert("RGB").resize((666, 375))
    px = im.load()
    k = 666 / W_IN
    w = min(box["w"], 4.0)
    x0 = box["x"] + {"center": (box["w"] - w) / 2, "right": box["w"] - w}.get(box.get("align"), 0)
    h = box.get("h", 0.5)
    ys = box["y"] + {"middle": h / 2, "bottom": max(0, h - 0.4)}.get(box.get("valign"), min(h, 0.4) / 2)
    pts = [px[min(665, max(0, int((x0 + w * (i + 0.5) / 12) * k))), min(374, max(0, int((ys + d) * k)))] for i in range(12) for d in (-0.12, 0, 0.12)]
    c = lum(tuple(int(color[i:i + 2], 16) for i in (0, 2, 4)))
    vals = sorted((max(lum(p), c) + 0.05) / (min(lum(p), c) + 0.05) for p in pts)
    return vals[len(vals) // 2]


def analyze_sides(path):
    """本文の高さで、左右の端から飾りがどこまで張り出しているか（インチ。左右の大きい方。端から 1.2in まで）"""
    from PIL import Image
    im = Image.open(path).convert("RGB").resize((666, 375))
    px = im.load()
    k = 666 / W_IN
    ys = range(int(BODY_TOP * k), int(5.6 * k), 2)
    mid = sorted((px[333, y] for y in ys), key=lum)[len(ys) // 2]
    base_l = lum(mid)
    far = 0
    for side in (0, 1):
        for i in range(int(1.2 * k)):
            x = i if side == 0 else 665 - i
            odd = sum(1 for y in ys if (max(lum(px[x, y]), base_l) + 0.05) / (min(lum(px[x, y]), base_l) + 0.05) > 1.3) / len(ys)
            if odd > 0.5:  # 縦に長く続く帯だけを飾りとみなす（細い枠線は余白の中なら問題にならない）
                far = max(far, (i + 1) / k)
    return far


def dominant_color(paths):
    """背景の画像で、いちばん広く使われているはっきりした色（白の上で読める濃さ）。テーマの色が Office の既定のときに使う"""
    from PIL import Image
    import colorsys
    count = {}
    for p in paths:
        im = Image.open(p).convert("RGB").resize((320, 180))
        colors = im.getcolors(320 * 180)
        base = max(colors)[1]  # いちばん広い色（地）
        for n, c in colors:
            _, l, s = colorsys.rgb_to_hls(*[v / 255 for v in c])
            if s < 0.25 or (1.05 / (lum(c) + 0.05)) < 3.0 or (lum(base) < 0.05 and sum(abs(a - b) for a, b in zip(base, c)) < 90):
                continue
            key = tuple(v // 24 for v in c)
            e = count.setdefault(key, [0, c, 0])
            e[0] += n
            if n > e[2]:
                e[1], e[2] = c, n
    if not count:
        return None
    return "%02X%02X%02X" % max(count.values(), key=lambda e: e[0])[1]


def r2(v):
    return round(v + 1e-9, 2)


def build_master(deck, picked, outdir, rel_from, render_script, drop=()):
    """theme.json の master・footer・layout を作る。(辞書, 知らせる行) を返す"""
    notes = []
    # 章ナビと、ページ番号の前後の文字は、部品が描くので背景から外す
    remove, navs, formats = set(), {}, {}
    for kind, lay in picked.items():
        nav, shapes, names = deck.find_nav(lay)
        if nav:
            navs[kind] = nav
            remove |= {(s["part"], s["id"]) for s in shapes}
            notes.append("%sの章ナビ（%s）を背景から外し、nav の設定にした。章の名前は build.js の createDeck の nav に書く" % (KIND_JA[kind], "・".join(names)))
        num = next((b for b in deck.placeholders(lay) if b["type"] == "sldNum"), None)
        if num:
            found, fmt = deck.number_affixes(lay, num)
            formats[kind] = fmt
            remove |= {(s["part"], s["id"]) for s in found}
    bgs = deck.render_backgrounds(picked, outdir, drop=drop, render_script=render_script, remove=remove)
    master, extra = {}, {}
    for kind in KINDS:
        if kind not in picked:
            notes.append("%sのレイアウトが見つからない（既定の見た目になる）" % KIND_JA[kind])
            continue
        lay = picked[kind]
        notes.append("%s = レイアウト「%s」" % (KIND_JA[kind], lay["name"]))
        removed = {s["text"] for s in deck.static_shapes(lay) if (s["part"], s["id"]) in remove}
        st = [t for t in deck.static_texts(lay) if t not in removed and t not in drop]
        if st:
            notes.append("  背景の画像に焼き込んだ固定の文字: %s（章ナビや「/ 12」のように部品が描くものは --drop-text で外す）" % "、".join(st[:12]))
        phs = deck.placeholders(lay)
        titles = [b for b in phs if b["type"] in ("title", "ctrTitle")]
        others = sorted([b for b in phs if b["type"] in ("body", "subTitle", "obj")], key=lambda b: (b["y"], b["x"]))
        nums = [b for b in phs if b["type"] == "sldNum"]
        page = {"background": os.path.relpath(bgs[kind], rel_from), "boxes": {}}
        info = analyze(bgs[kind])

        def put(name, b, **over):
            d = {"x": r2(b["x"]), "y": r2(b["y"]), "w": r2(b["w"]), "h": r2(b["h"]), "align": b["align"], "valign": b["valign"]}
            for key in ("size", "color", "font"):
                if b.get(key):
                    d[key] = b[key]
            d["bold"] = b["bold"]
            d.update(over)
            d["w"] = r2(min(d["w"], W_IN - 0.4))
            d["x"] = r2(max(0.2, min(d["x"], W_IN - 0.2 - d["w"])))
            page["boxes"][name] = d

        if kind == "content":
            if info["dark"] or info["busy"] > 0.06:
                why = "地が濃いか色が強い（%s）" % info["base"] if info["dark"] else "図を置く所に絵や模様がある（約 %d%%）" % round(info["busy"] * 100)
                notes.append("  本文は%s。部品の図は薄い無地が前提なので、本文のページの背景は使わない（表紙・章扉は使う）。別のレイアウトを使うなら --content <名前>" % why)
                continue
            if titles:
                t = titles[0]
                size = min(32, max(20, t.get("size", 26)))
                line = size * 1.2 / 72
                y = t["y"] + {"middle": (t["h"] - line) / 2, "bottom": t["h"] - line}.get(t["valign"], 0)
                low = y + line > 1.55  # 題が低く、リードと図の場所に食い込む
                if low:
                    notes.append("  本文の題の置き場所が低い（y %.2f）ので、部品の標準の位置にする" % y)
                    page["boxes"]["title"] = {k: v for k, v in {"size": size, "color": t.get("color"), "bold": t["bold"]}.items() if v is not None}
                else:
                    put("title", t, y=r2(max(0.05, y)), h=r2(max(0.5, line + 0.05)), size=size, valign="top")
                below = [b for b in others if b["y"] >= t["y"] + t["h"] * 0.5 and b["h"] <= 1.1 and b["w"] >= 6]
                if low:
                    pass
                elif below:
                    b = below[0]
                    size = min(26, max(18, b.get("size", 20)))
                    line = size * 1.2 / 72
                    y = b["y"] + {"middle": (b["h"] - line) / 2, "bottom": b["h"] - line}.get(b["valign"], 0)
                    y = min(1.75, max(y, 0.8))
                    col = {"color": b["color"]} if b.get("color") else {}
                    page["boxes"]["lead"] = dict({"x": r2(b["x"]), "y": r2(y), "w": r2(b["w"]), "h": r2(max(0.45, min(0.8, 2.3 - y))), "size": size, "align": b["align"]}, **col)
                else:
                    y = min(1.45, max(1.15, t["y"] + t["h"] + 0.1))
                    if t["w"] >= 10:
                        page["boxes"]["lead"] = {"x": r2(t["x"]), "y": r2(y), "w": r2(t["w"]), "align": t["align"]}
                    else:  # 狭い題の置き場所に合わせない
                        x = t["x"] if M_IN <= t["x"] <= 1.2 else M_IN
                        page["boxes"]["lead"] = {"x": r2(x), "y": r2(y), "w": r2(W_IN - x * 2), "align": "left"}
            head = info["head_to"]
            if head and head > 1.0:
                lead = page["boxes"].setdefault("lead", {})
                if lead.get("y", 1.2) < head + 0.08:
                    y = r2(min(1.75, head + 0.12))
                    lead.update(y=y, h=r2(max(0.4, 2.18 - y)), size=min(20, lead.get("size", 20)))
                    lead.pop("color", None)
                    notes.append("  上の帯が y %.2f まであるので、リードはその下（y %.2f）に置く" % (head, y))
            if info["decor_from"] is not None and info["decor_from"] - 0.12 < BODY_BOTTOM:
                bb = r2(max(5.6, info["decor_from"] - 0.12))
                extra.setdefault("layout", {})["bodyBottom"] = bb
                notes.append("  下の飾りが y %.2f から始まるので、本文の下端を %.2f にする" % (info["decor_from"], bb))
        elif kind == "cover":
            if titles:
                put("title", titles[0], size=min(60, titles[0].get("size", 40)))
            if others:
                clamp = lambda b, lo, hi, default: min(hi, max(lo, b.get("size", default)))

                def wide(b):  # 副題と日付が折り返さない幅にする
                    if b["w"] >= 6:
                        return b
                    cx = b["x"] + b["w"] / 2
                    return dict(b, w=6.0, x={"center": cx - 3.0, "right": b["x"] + b["w"] - 6.0}.get(b["align"], b["x"]))
                if len(others) >= 2:
                    put("sub", wide(others[0]), size=clamp(others[0], 16, 28, 18))
                    put("date", wide(others[-1]), size=clamp(others[-1], 12, 20, 14))
                else:
                    put("date", wide(others[0]), size=clamp(others[0], 12, 20, 14))
        else:  # chapter
            if titles:
                t = titles[0]
                put("title", t)
                if others:
                    big = max(others, key=lambda b: b.get("size", 0))
                    if (big["h"] <= 2.0 or big["w"] <= 5.0) and (big.get("size", 0) >= 28 or big["y"] + big["h"] <= t["y"] + 0.1):
                        put("num", big)
        for name, b in page["boxes"].items():
            if "color" in b and "x" in b and contrast_on(bgs[kind], b, b["color"]) < 1.8:
                notes.append("  %sの %s の色（%s）は地に溶けるので使わない（部品が読める色を選ぶ）" % (KIND_JA[kind], name, b["color"]))
                del b["color"]
        if nums:
            n = nums[0]
            pn = {"x": r2(n["x"]), "y": r2(n["y"]), "w": r2(max(n["w"], 0.8) if n["align"] != "right" else n["w"]), "h": r2(n["h"]), "align": n["align"], "size": n.get("size", 12), "format": formats.get(kind, "{n}"), "bold": n["bold"]}
            if pn["format"] != "{n}":  # 前後の文字の分、置き場所を広げる
                extra_w = max(0, len(pn["format"]) * pn["size"] / 72 * 0.7 - pn["w"])
                pn["x"] = r2(pn["x"] - extra_w * {"center": 0.5, "right": 1}.get(n["align"], 0))
                pn["w"] = r2(pn["w"] + extra_w)
            if n.get("color"):
                pn["color"] = n["color"]
            page["pageNumber"] = pn
        else:
            page["pageNumber"] = False
        if kind in navs:
            page["nav"] = navs[kind]
        master[kind] = page
    # 飾りが左右に張り出していれば、本文の余白を広げる。題が内側にそろえてあれば、本文もそれにそろえる
    if "content" in master:
        side = analyze_sides(bgs["content"])
        t = master["content"]["boxes"].get("title")
        t = t if t and "x" in t else None
        want = t["x"] if t and t.get("align") == "left" and 0.75 <= t["x"] <= 1.6 else 0.6
        margin = min(1.0, max(0.6, side + 0.15 if side > 0.45 else 0.6, min(want, 0.9)))
        if margin > 0.62:
            extra.setdefault("layout", {})["margin"] = r2(margin)
            notes.append("左右の飾りと題の位置に合わせて、本文の左右の余白を %.2f にする" % margin)
    blank = [k for k, p in bgs.items() if not picked[k].get("slide") and is_blank(p)]
    if blank and len(blank) == len(bgs) and deck.slides and deck.has_slide_design():
        notes.append("レイアウトに飾りが無い。デザインはスライドに直接描かれているようなので、元の pptx を画像にして、表紙・章扉・本文にあたるスライドの番号を "
                     "--cover-slide・--chapter-slide・--content-slide で指定して読み直す（全 %d 枚）" % len(deck.slides))
    return master, extra, bgs, notes
