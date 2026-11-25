#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""線の絵（1色のアイコン）を、資料の色の組に合わせて塗り直す（pptx-deck-design）。

使い方:
  python3 tint_icons.py <icons のフォルダ> --ink <文字の色> [--emph <強調の色>] [--accent <問題の色>]
例:
  python3 tint_icons.py icons --ink 1A1F2B --emph 0E7C86 --accent C5221F

すること:
- 線の絵（下の GENERIC の名前と、pic-*.png）を、形はそのままに、文字の色（--ink）で塗り直す
- 色ちがいを作る: <名前>-w.png（白。色を塗った丸の上に置く）、<名前>-e.png（強調の色）、<名前>-r.png（問題の色）
- 外部の製品・会社の公式のアイコン（slack.png など）は触らない
- 何色で塗ったかを icons/.tint.json に残す（deck-parts.js が、資料の色の組と合っているかを確かめる）
色は '#' なしの6桁。PIL（Pillow）が要る。
"""
import json, os, sys
from PIL import Image

# 1色の線の絵（make-icons.js の GENERIC と同じ名前）。ここに無い名前は、公式のアイコンとして触らない
GENERIC = """user users laptop desktop smartphone tablet internet cloud server compute function container database storage bucket
folder file files spreadsheet api gateway webhook plug queue network auth shield lock monitoring logs search vector-search ai brain bot
workflow settings chart code terminal git-branch pull-request mail chat building""".split()
SUFFIX = ("-w", "-e", "-r", "-b")

def rgb(h): return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))

def main():
    a = sys.argv[1:]
    if not a or "--ink" not in a: sys.exit(__doc__)
    d = a[0]
    opt = lambda k: a[a.index(k) + 1].lstrip("#").upper() if k in a else None
    ink, emph, accent = opt("--ink"), opt("--emph"), opt("--accent") or "C5221F"
    emph = emph or ink
    names = [n for n in GENERIC if os.path.exists(os.path.join(d, n + ".png"))]
    names += sorted(f[:-4] for f in os.listdir(d) if f.startswith("pic-") and f.endswith(".png") and not f[:-4].endswith(SUFFIX))
    tones = {"": ink, "-w": "FFFFFF", "-e": emph, "-r": accent}
    for n in names:
        alpha = Image.open(os.path.join(d, n + ".png")).convert("RGBA").getchannel("A")
        for suf, col in tones.items():
            im = Image.new("RGBA", alpha.size, rgb(col) + (0,)); im.putalpha(alpha)
            im.save(os.path.join(d, n + suf + ".png"))
    json.dump({"ink": ink, "emph": emph, "accent": accent, "icons": len(names)}, open(os.path.join(d, ".tint.json"), "w"))
    print(f"tinted {len(names)} icons x {len(tones)} tones in {d} (ink {ink}, emph {emph}, accent {accent})")

if __name__ == "__main__":
    main()
