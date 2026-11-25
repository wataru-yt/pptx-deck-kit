# -*- coding: utf-8 -*-
# PDF を1ページずつ PNG にし、確認用の一覧（4ページずつの縦横2×2）も作る。pdftoppm が無い Mac 用
# 使い方: python3 pdf2png.py out/render/deck.pdf out/render
import sys, fitz
from PIL import Image
pdf, outdir = sys.argv[1], sys.argv[2]
doc = fitz.open(pdf)
paths = []
for i, page in enumerate(doc, 1):
    p = f"{outdir}/{i:02d}.png"
    page.get_pixmap(dpi=96).save(p)
    paths.append(p)
for g in range(0, len(paths), 4):
    ims = [Image.open(p) for p in paths[g:g + 4]]
    w, h = ims[0].size
    sheet = Image.new("RGB", (w * 2, h * 2), "white")
    for k, im in enumerate(ims):
        sheet.paste(im, ((k % 2) * w, (k // 2) * h))
    sheet.save(f"{outdir}/sheet{g // 4 + 1}.png")
print(len(paths))
