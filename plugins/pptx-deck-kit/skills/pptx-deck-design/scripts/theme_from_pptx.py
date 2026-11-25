#!/usr/bin/env python3
"""会社の pptx（または .potx・.thmx）から、pptx-deck-design のテーマの JSON を作る。

使い方:
  python3 theme_from_pptx.py <pptx|potx|thmx> <出力の theme.json> [--logo-out <フォルダ>]

読むもの（標準ライブラリの zipfile と xml だけで読む。元のファイルは書き換えない）:
  - スライドマスターが使うテーマ（ppt/theme/themeN.xml）の色：dk1・lt1・dk2・lt2・accent1〜6・hlink
  - 同じテーマの書体：major（見出し）と minor（本文）の latin・ea（無ければ日本語の script="Jpan"）
  - スライドマスターにある画像から、ロゴらしいもの（名前に logo を含むもの。無ければ、背景ほど大きくない画像のうち一番大きいもの）
  - スライドマスター・レイアウト・スライドのフッターの欄（ftr）の文字
読み取れなかった項目は書かない（部品の既定に任せる）。何を読み取り、どの色を強調（emph）に当てたかを表示する。
置き場所: プロジェクトなら <プロジェクト>/.claude/pptx-deck/theme.json、自分の既定なら ~/.claude/pptx-deck/theme.json
"""
import argparse
import colorsys
import json
import os
import posixpath
import re
import shutil
import sys
import zipfile
import xml.etree.ElementTree as ET

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}
R_EMBED = "{%s}embed" % NS["r"]


def rels_of(z, part):
    """part の関係（Id → 実体のパス）を返す"""
    d, b = posixpath.split(part)
    rp = posixpath.join(d, "_rels", b + ".rels")
    if rp not in z.namelist():
        return {}
    out = {}
    root = ET.fromstring(z.read(rp))
    for rel in root.findall("rel:Relationship", NS):
        if rel.get("TargetMode") == "External":
            continue
        t = rel.get("Target")
        full = posixpath.normpath(posixpath.join(d, t)) if not t.startswith("/") else t.lstrip("/")
        out[rel.get("Id")] = (full, rel.get("Type", "").rsplit("/", 1)[-1])
    return out


def find_parts(z, kind):
    pat = re.compile(r"(^|/)%ss/%s\d+\.xml$" % (kind, kind))
    names = [n for n in z.namelist() if pat.search(n) and not n.startswith("themeVariants/")]
    return sorted(names, key=lambda n: int(re.search(r"(\d+)\.xml$", n).group(1)))


def color_of(el):
    """a:dk1 などの中の色（srgbClr か sysClr の lastClr）"""
    if el is None:
        return None
    s = el.find("a:srgbClr", NS)
    if s is not None and s.get("val"):
        return s.get("val").upper()
    s = el.find("a:sysClr", NS)
    if s is not None and s.get("lastClr"):
        return s.get("lastClr").upper()
    return None


def rgb(h):
    return tuple(int(h[i:i + 2], 16) / 255 for i in (0, 2, 4))


def luminance(h):
    def ch(c):
        return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4
    r, g, b = (ch(c) for c in rgb(h))
    return 0.2126 * r + 0.7152 * g + 0.0722 * b


def contrast_white(h):
    return 1.05 / (luminance(h) + 0.05)


def is_reddish(h):
    hh, l, s = colorsys.rgb_to_hls(*rgb(h))
    deg = hh * 360
    return s > 0.35 and (deg < 15 or deg > 345)


def hue_dist(a, b):
    ha = colorsys.rgb_to_hls(*rgb(a))[0] * 360
    hb = colorsys.rgb_to_hls(*rgb(b))[0] * 360
    d = abs(ha - hb) % 360
    return min(d, 360 - d)


# 日本語の書体名は、PowerPoint・部品の既定と同じ英語の名前にそろえる（どちらでも PowerPoint は読める）
FONT_NAMES = {"游ゴシック": "Yu Gothic", "游ゴシック Light": "Yu Gothic Light", "游ゴシック Medium": "Yu Gothic Medium", "游明朝": "Yu Mincho",
              "メイリオ": "Meiryo", "ＭＳ ゴシック": "MS Gothic", "ＭＳ Ｐゴシック": "MS PGothic", "ＭＳ 明朝": "MS Mincho", "ＭＳ Ｐ明朝": "MS PMincho",
              "BIZ UDPゴシック": "BIZ UDPGothic", "BIZ UDゴシック": "BIZ UDGothic"}


def is_grayish(h):
    _, _, s = colorsys.rgb_to_hls(*rgb(h))
    return s < 0.12


def read_theme(z):
    masters = find_parts(z, "slideMaster")
    theme_part = None
    if masters:
        for full, typ in rels_of(z, masters[0]).values():
            if typ == "theme":
                theme_part = full
                break
    if not theme_part:
        cands = [n for n in z.namelist() if re.search(r"(^|/)theme/theme\d+\.xml$", n) and not n.startswith("themeVariants/")]
        theme_part = sorted(cands)[0] if cands else None
    if not theme_part or theme_part not in z.namelist():
        return None, None, None
    root = ET.fromstring(z.read(theme_part))
    cs = root.find(".//a:clrScheme", NS)
    colors = {}
    if cs is not None:
        for k in ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink"]:
            c = color_of(cs.find("a:" + k, NS))
            if c:
                colors[k] = c
    fonts = {}
    fs_ = root.find(".//a:fontScheme", NS)
    if fs_ is not None:
        for kind in ("major", "minor"):
            f = fs_.find("a:%sFont" % kind, NS)
            if f is None:
                continue
            lat = f.find("a:latin", NS)
            ea = f.find("a:ea", NS)
            jp = [x for x in f.findall("a:font", NS) if x.get("script") == "Jpan"]
            fonts[kind] = {
                "latin": (lat.get("typeface") if lat is not None else "") or "",
                "ea": (ea.get("typeface") if ea is not None else "") or "",
                "jpan": (jp[0].get("typeface") if jp else "") or "",
            }
    return theme_part, colors, fonts


def slide_size(z):
    for n in z.namelist():
        if re.search(r"(^|/)presentation\.xml$", n):
            root = ET.fromstring(z.read(n))
            sz = root.find("p:sldSz", NS)
            if sz is not None:
                return int(sz.get("cx")), int(sz.get("cy"))
    return None


def pics_in(z, part):
    """part の中の画像：(名前, 説明, 面積 EMU², 実体のパス)"""
    rels = rels_of(z, part)
    root = ET.fromstring(z.read(part))
    out = []
    for pic in root.iter("{%s}pic" % NS["p"]):
        nv = pic.find(".//p:cNvPr", NS)
        blip = pic.find(".//a:blip", NS)
        ext = pic.find(".//a:xfrm/a:ext", NS)
        if blip is None or blip.get(R_EMBED) not in rels:
            continue
        target = rels[blip.get(R_EMBED)][0]
        area = int(ext.get("cx")) * int(ext.get("cy")) if ext is not None else 0
        out.append(((nv.get("name") if nv is not None else "") or "", (nv.get("descr") if nv is not None else "") or "", area, target))
    return out


def find_logo(z, size):
    masters = find_parts(z, "slideMaster")
    layouts = find_parts(z, "slideLayout")
    slide_area = size[0] * size[1] if size else 12192000 * 6858000
    for where, parts in (("スライドマスター", masters), ("レイアウト", layouts)):
        pics = []
        for part in parts:
            pics += [(n, d, a, t, part) for (n, d, a, t) in pics_in(z, part)]
        pics = [p for p in pics if p[3] in z.namelist() and re.search(r"\.(png|jpe?g|gif)$", p[3], re.I)]
        if not pics:
            continue
        named = [p for p in pics if re.search(r"logo|ロゴ", p[0] + " " + p[1] + " " + p[3], re.I)]
        if named:
            return max(named, key=lambda p: p[2]), where, "名前に logo を含む"
        small = [p for p in pics if p[2] < slide_area * 0.25]  # 背景・全面の写真は外す
        if small:
            return max(small, key=lambda p: p[2]), where, "背景ほど大きくない画像のうち一番大きい"
    return None, None, None


def para_text(sp):
    paras = []
    for p in sp.iter("{%s}p" % NS["a"]):
        t = "".join(x.text or "" for x in p.iter("{%s}t" % NS["a"]))
        if t.strip():
            paras.append(t.strip())
    return " ".join(paras)


def footer_texts(z):
    out = []
    parts = find_parts(z, "slideMaster") + find_parts(z, "slideLayout") + find_parts(z, "slide")
    for part in parts:
        root = ET.fromstring(z.read(part))
        for sp in root.iter("{%s}sp" % NS["p"]):
            ph = sp.find(".//p:nvPr/p:ph", NS)
            if ph is None or ph.get("type") != "ftr":
                continue
            t = para_text(sp)
            if t and not re.fullmatch(r"(フッター|Footer|フッター テキスト|Click to edit.*)", t, re.I) and t not in out:
                out.append(t)
    return out


def pick_font(f):
    if not f:
        return None
    for k in ("ea", "jpan", "latin"):
        v = f.get(k, "")
        if v and not v.startswith("+"):
            return FONT_NAMES.get(v, v)
    return None


def main():
    ap = argparse.ArgumentParser(description="会社の pptx・potx・thmx から、pptx-deck-design のテーマの JSON を作る")
    ap.add_argument("src", help="読む pptx・potx・thmx（書き換えない）")
    ap.add_argument("out", help="書き出す theme.json")
    ap.add_argument("--logo-out", help="ロゴの画像を書き出すフォルダ（既定は theme.json と同じフォルダ）")
    a = ap.parse_args()

    z = zipfile.ZipFile(a.src)
    theme = {"name": os.path.splitext(os.path.basename(a.src))[0], "description": "%s から theme_from_pptx.py で作った" % os.path.basename(a.src)}
    got, miss = [], []

    size = slide_size(z)
    if size and abs(size[0] / size[1] - 16 / 9) > 0.02:
        print("注意: 元のスライドは %.2f:1（16:9 ではない）。部品は 16:9 固定で描く（4:3 には対応しない）" % (size[0] / size[1]))

    theme_part, colors, fonts = read_theme(z)
    if not theme_part:
        print("テーマ（theme1.xml）が見つからない。色と書体は既定のまま", file=sys.stderr)
        colors, fonts = {}, {}
    else:
        got.append("テーマ: %s" % theme_part)

    # ---- 色 ----
    tc = {}
    if colors:
        print("テーマの色: " + "  ".join("%s=%s" % (k, v) for k, v in colors.items()))
        cands = [k for k in ["accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "dk2"] if k in colors]
        emph_key = None
        for k in cands:
            c = colors[k]
            if contrast_white(c) >= 3.0 and not is_reddish(c) and not is_grayish(c):
                emph_key = k
                break
        if emph_key is None:
            for k in cands:
                if contrast_white(colors[k]) >= 3.0 and not is_reddish(colors[k]):
                    emph_key = k
                    break
        if emph_key:
            tc["emph"] = colors[emph_key]
            got.append("強調の色（emph）= %s（%s）" % (colors[emph_key], emph_key))
            skipped = [k for k in cands[:cands.index(emph_key)]]
            why = "。前の %s は、白の上で薄すぎるか・赤（問題の色）か・灰なので外した" % "・".join(skipped) if skipped else ""
            print("強調の色（emph）に %s の %s を当てた%s。別の色にするなら theme.json の colors.emph を書き換える" % (emph_key, colors[emph_key], why))
        else:
            miss.append("強調の色（白の上で読める濃さの accent が無い）")
        ink = colors.get("dk1")
        if ink and ink not in ("000000",) and contrast_white(ink) >= 7:
            tc["ink"] = ink
            got.append("文字の色（ink）= %s（dk1）" % ink)
        elif ink:
            print("文字の色: dk1 が %s（真っ黒）なので書かない。部品の既定（黒に近い色）を使う" % ink)
        # 問題の色（accent）：テーマにはっきりした赤があり、強調の色と色相が離れているときだけ当てる
        red = [colors[k] for k in cands if k.startswith("accent") and is_reddish(colors[k]) and contrast_white(colors[k]) >= 3.0
               and (not emph_key or hue_dist(colors[k], colors[emph_key]) >= 40)]
        if red:
            tc["accent"] = red[0]
            got.append("問題の色（accent）= %s（テーマの赤）" % red[0])
        if colors.get("hlink"):
            tc["link"] = colors["hlink"]
    if tc:
        theme["colors"] = tc
    else:
        miss.append("色")

    # ---- 書体 ----
    face = pick_font(fonts.get("minor"))
    head = pick_font(fonts.get("major"))
    if fonts:
        print("テーマの書体: " + "  ".join("%s(latin=%s ea=%s Jpan=%s)" % (k, v["latin"] or "-", v["ea"] or "-", v["jpan"] or "-") for k, v in fonts.items()))
    if face:
        theme["font"] = {"face": face}
        got.append("本文の書体 = %s" % face)
        if head and head != face and re.search(r"Light|Thin|ExtraLight|細", head, re.I):
            print("見出しの書体 %s は細い字の書体なので書かない（部品の題は太字で描く）。本文と同じ %s を使う" % (head, face))
        elif head and head != face:
            theme["font"]["heading"] = head
            got.append("見出しの書体 = %s" % head)
    else:
        miss.append("書体")

    # ---- ロゴ ----
    out_dir = os.path.dirname(os.path.abspath(a.out))
    logo, where, why = find_logo(z, size)
    if logo:
        name, descr, area, target, part = logo
        ld = os.path.abspath(a.logo_out) if a.logo_out else out_dir
        os.makedirs(ld, exist_ok=True)
        dst = os.path.join(ld, "logo" + os.path.splitext(target)[1].lower())
        with open(dst, "wb") as f:
            f.write(z.read(target))
        rel = os.path.relpath(dst, out_dir)
        logo_path = rel if not rel.startswith("..") else dst
        theme["footer"] = {"logo": logo_path}
        theme["cover"] = {"logo": logo_path, "logoPos": "top-left"}
        got.append("ロゴ = %s（%s の「%s」、%s）→ %s" % (target, where, name or "名前なし", why, dst))
    else:
        miss.append("ロゴ（スライドマスターに画像が無い）")

    # ---- フッターの文字 ----
    fts = footer_texts(z)
    if fts:
        ft = theme.setdefault("footer", {})
        for t in fts:
            if re.search(r"社外秘|極秘|部外秘|取扱注意|confidential|internal use", t, re.I) and "confidential" not in ft:
                ft["confidential"] = t
                got.append("機密の表記 = %s" % t)
            elif re.search(r"©|\(c\)|copyright|all rights reserved", t, re.I) and "copyright" not in ft:
                ft["copyright"] = t
                got.append("著作権の表記 = %s（帯のフッター style: band のときに出る）" % t)
            elif "text" not in ft and "copyright" not in ft:
                ft["text"] = t
                got.append("フッターの文字 = %s" % t)
        if "copyright" in ft:
            ft.setdefault("style", "band")
    else:
        miss.append("フッターの文字")

    with open(a.out, "w", encoding="utf-8") as f:
        json.dump(theme, f, ensure_ascii=False, indent=2)
        f.write("\n")
    print()
    print("書き出した: %s" % os.path.abspath(a.out))
    for g in got:
        print("  読み取った: " + g)
    for m in miss:
        print("  読み取れなかった（既定のまま）: " + m)
    print("置き場所: プロジェクトは <プロジェクト>/.claude/pptx-deck/theme.json、自分の既定は ~/.claude/pptx-deck/theme.json")


if __name__ == "__main__":
    main()
