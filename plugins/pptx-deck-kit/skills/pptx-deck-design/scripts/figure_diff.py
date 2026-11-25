"""重ね技の確認：2枚のページ画像を、図の領域（タイトルとリードより下）だけで比べる。

同じ図を背面に敷いたページ（echo）が、元のページと1px も動いていないかを確かめる。
確認用のページは build.js で echo(..., { check: true }) にして描き（覆いと前面を外す）、画像にしてから比べる。

使い方:
  python3 figure_diff.py 元のページ.jpg 確認用のページ.jpg [--top 2.1] [--bottom 6.9] [--threshold 40]

--top / --bottom は図の領域の上端・下端（インチ、スライドの高さ 7.5 に対する位置）。
差のある画素が 0 なら図は動いていない。0 でなければ、差分画像（<確認用>.diff.png）を見る。
"""
import argparse
import os

from PIL import Image, ImageChops

SLIDE_H = 7.5


def main():
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("base")
    p.add_argument("check")
    p.add_argument("--top", type=float, default=2.1)
    p.add_argument("--bottom", type=float, default=6.9)
    p.add_argument("--threshold", type=int, default=40)
    a = p.parse_args()

    im_a = Image.open(a.base).convert("L")
    im_b = Image.open(a.check).convert("L")
    if im_a.size != im_b.size:
        raise SystemExit(f"画像の大きさが違う: {im_a.size} と {im_b.size}（同じ解像度で画像にする）")
    w, h = im_a.size
    box = (0, int(h * a.top / SLIDE_H), w, int(h * a.bottom / SLIDE_H))
    diff = ImageChops.difference(im_a.crop(box), im_b.crop(box))
    n = sum(1 for v in diff.getdata() if v > a.threshold)
    print(f"図の領域（y {a.top}〜{a.bottom}in）で差のある画素: {n}")
    if n:
        out = os.path.splitext(a.check)[0] + ".diff.png"
        diff.point(lambda v: 255 if v > a.threshold else 0).save(out)
        print(f"差分画像: {out}")


if __name__ == "__main__":
    main()
