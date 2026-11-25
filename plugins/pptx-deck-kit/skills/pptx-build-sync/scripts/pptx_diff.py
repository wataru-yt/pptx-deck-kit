#!/usr/bin/env python3
"""2つの pptx を、スライドごとに図形単位で比べる（build.js の出力 vs 手修正された pptx など）。

使い方:
  python3 pptx_diff.py built.pptx current.pptx
  python3 pptx_diff.py built.pptx current.pptx --map 9:10,10:11   # ページがずれているとき（左:右）
  python3 pptx_diff.py built.pptx current.pptx --summary           # ページごとの一致/差分の件数だけ

比較は「図形の集合」で行う（重なり順は見ない）。次の揺れは同じものとして扱う:
  座標の0.005インチ未満の差、空の段落、画像ファイル名、表の外枠の大きさ（PowerPointが再計算する）、
  既定値と同じ属性（anchor=t、marL=0 など）。
差分は「左のみ」「右のみ」の図形として出る。文字・数値・色・表・グラフ・ノート・リンク先が対象。
"""
import argparse
import os
import re
import sys
from collections import Counter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from pptx_inventory import inventory  # noqa: E402


def r2(m):
    return "%s=%g" % (m.group(1), int(float(m.group(2)) * 100 + 0.5) / 100)


def normalize(block):
    first, _, rest = block.partition("\n")
    if first.startswith("<graphicFrame>") and "tbl cols" in rest:
        first = re.sub(r" w=[\d.]+ h=[\d.]+", "", first)
    else:
        first = re.sub(r"\b([xywh])=(-?[\d.]+)", r2, first)
    b = first + "\n" + rest
    b = re.sub(r"img \S+", "img IMG", b)
    b = b.replace(" anchor=t", "").replace("(marL=0.0)", "").replace(" marL=0.0", "").replace(" wrap=square", "")
    b = b.replace("ln=none", "ln=None")
    b = re.sub(r"\(al=l( |\))", lambda m: "(" if m.group(1) == " " else "()", b).replace("()", "")
    b = re.sub(r"\n\s*¶\s*(?=\n|$)", "", b)
    b = re.sub(r"\n\s*¶ *$", "", b)
    b = re.sub(r"(?m)^(\s*¶)\(\s+", r"\1(", b)
    return b


def blocks(lines):
    out, cur = [], []
    for ln in lines:
        if ln.strip() == "¶":
            continue
        if re.match(r"^\s*<(sp|pic|graphicFrame|grpSp|cxnSp)>", ln) or ln.startswith("NOTES"):
            if cur:
                out.append("\n".join(cur))
            cur = [ln]
        else:
            cur.append(ln)
    if cur:
        out.append("\n".join(cur))
    return [normalize(b) for b in out]


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("left")
    ap.add_argument("right")
    ap.add_argument("--map", help="左ページ:右ページ をカンマ区切りで（既定は同じ番号）")
    ap.add_argument("--summary", action="store_true")
    ap.add_argument("--width", type=int, default=240)
    a = ap.parse_args()
    L = dict(inventory(a.left))
    R = dict(inventory(a.right))
    mapping = {}
    if a.map:
        for kv in a.map.split(","):
            k, v = kv.split(":")
            mapping[int(k)] = int(v)
    pairs = [(n, mapping.get(n, n)) for n in sorted(L)]
    for rn in sorted(R):
        if rn not in [r for _, r in pairs]:
            pairs.append((None, rn))
    total = 0
    for ln_, rn in pairs:
        if ln_ is None or rn not in R:
            print("#### 左 %s / 右 %s: 片側にしか存在しない" % (ln_, rn))
            total += 1
            continue
        bl, br = blocks(L[ln_]), blocks(R[rn])
        cl, cr = Counter(bl), Counter(br)
        only_l, only_r = cl - cr, cr - cl
        label = "左 %d / 右 %d" % (ln_, rn) if ln_ != rn else "スライド %d" % ln_
        if not only_l and not only_r:
            if not a.summary:
                print("#### %s: 一致" % label)
            continue
        total += 1
        print("#### %s: 左のみ %d / 右のみ %d" % (label, sum(only_l.values()), sum(only_r.values())))
        if a.summary:
            continue
        for title, only, src in (("--- 左のみ", only_l, bl), ("+++ 右のみ", only_r, br)):
            print(title)
            for b in src:
                if only.get(b, 0) > 0:
                    print("\n".join(x[: a.width] for x in b.split("\n")))
                    only[b] -= 1
    print("差分のあるスライド: %d" % total)
    return 0


if __name__ == "__main__":
    sys.exit(main())
