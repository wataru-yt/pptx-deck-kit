#!/usr/bin/env python3
"""render_pptx.sh で作った2つの画像フォルダをページごとに比べ、差のあるページを並べた画像にする。

使い方:
  python3 render_compare.py dirA dirB out_dir [--threshold 40] [--only 6,9]
出力: out_dir/cmp-N.png（上段=A と B を横に並べる、下段左=差分が白で出る）。標準出力にページごとの差分画素数。
目安: 数百画素以下は字形の差（PowerPointの再保存・フォント指定の違い）。数千以上は図形のずれや欠落を疑う。
Python 3.9 以上、Pillow が必要。
"""
import argparse
import glob
import os

from PIL import Image, ImageChops


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("a")
    ap.add_argument("b")
    ap.add_argument("out")
    ap.add_argument("--threshold", type=int, default=40)
    ap.add_argument("--only")
    ap.add_argument("--min-pixels", type=int, default=500, help="これ以上のページだけ画像を作る")
    args = ap.parse_args()
    fa = sorted(glob.glob(os.path.join(args.a, "s-*.jpg")))
    fb = sorted(glob.glob(os.path.join(args.b, "s-*.jpg")))
    if len(fa) != len(fb):
        print("ページ数が違う: %d vs %d" % (len(fa), len(fb)))
    only = set(int(x) for x in args.only.split(",")) if args.only else None
    os.makedirs(args.out, exist_ok=True)
    for i, (x, y) in enumerate(zip(fa, fb), 1):
        A, B = Image.open(x).convert("RGB"), Image.open(y).convert("RGB")
        d = ImageChops.difference(A.convert("L"), B.convert("L")).point(lambda v: 255 if v > args.threshold else 0)
        n = sum(1 for v in d.getdata() if v)
        print("%2d: %d" % (i, n))
        if (only and i in only) or (not only and n >= args.min_pixels):
            w, h = A.size
            s = Image.new("RGB", (w * 2, h * 2), "white")
            s.paste(A, (0, 0))
            s.paste(B, (w, 0))
            s.paste(Image.merge("RGB", (d, d, d)), (0, h))
            s.save(os.path.join(args.out, "cmp-%d.png" % i))


if __name__ == "__main__":
    main()
