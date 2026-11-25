#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""外部の製品・会社の公式のアイコンを取得して、資料の icons/ に PNG で置く（pptx-deck-design）。

使い方:
  python3 fetch_icon.py <名前> <公式の URL か手元のファイル> [出力先のフォルダ（既定 ./icons）] [--note "出どころの説明"]
例:
  python3 fetch_icon.py teams https://example.com/brand/Teams-Icon.svg icons --note "公式サイトの製品ページの画像"
  python3 fetch_icon.py slack /Applications/Slack.app icons          # 手元に入っているアプリのアイコン（macOS）

決まり:
- 取得する先は、その会社の公式のサイト・公式のブランド素材・手元に入っているアプリだけ。似た形を自分で描かない
- <名前>.png（512px の正方形。余白は透明）で保存し、出どころと日付を icons/SOURCES.md に1行足す
- SVG は、macOS では sips で PNG にする（ImageMagick は、グラデーションのある SVG を黒くつぶすことがある）。できた PNG は画像で確かめる
- 外に配る資料では、各社のロゴの使い方の決まりを確かめる
"""
import datetime, os, shutil, subprocess, sys, tempfile, urllib.request

def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    note = sys.argv[sys.argv.index("--note") + 1] if "--note" in sys.argv else ""
    if "--note" in sys.argv and note in args: args.remove(note)
    if len(args) < 2: sys.exit(__doc__)
    name, src, out = args[0], args[1], (args[2] if len(args) > 2 else "icons")
    os.makedirs(out, exist_ok=True)
    work = tempfile.mkdtemp(prefix="fetch-icon-")
    try:
        if src.startswith(("http://", "https://")):
            ext = os.path.splitext(src.split("?")[0])[1].lower() or ".png"
            raw = os.path.join(work, "raw" + ext)
            req = urllib.request.Request(src, headers={"User-Agent": "deck-icon-fetch"})  # 利用者の情報は送らない
            with urllib.request.urlopen(req, timeout=30) as r, open(raw, "wb") as f: f.write(r.read())
        elif src.endswith(".app"):
            # macOS のアプリ：アイコンのファイル（.icns）を探す
            res = os.path.join(src, "Contents", "Resources")
            icns = sorted([f for f in os.listdir(res) if f.endswith(".icns")], key=lambda f: -os.path.getsize(os.path.join(res, f)))
            if not icns: sys.exit(f"アイコンが見つからない: {src}")
            raw = os.path.join(res, icns[0])
        else:
            raw = src
        if not os.path.exists(raw): sys.exit(f"ファイルが無い: {raw}")
        dst = os.path.join(out, name + ".png")
        if shutil.which("sips"):
            subprocess.run(["sips", "-s", "format", "png", "-Z", "512", raw, "--out", dst], check=True, capture_output=True)
        else:
            from PIL import Image  # sips の無い環境（PNG・JPEG だけ）
            im = Image.open(raw).convert("RGBA"); im.thumbnail((512, 512)); im.save(dst)
        with open(os.path.join(out, "SOURCES.md"), "a", encoding="utf-8") as f:
            if f.tell() == 0: f.write("# アイコンの出どころ\n\n| ファイル | 出どころ | 取得した日 | 説明 |\n|---|---|---|---|\n")
            f.write(f"| {name}.png | {src} | {datetime.date.today().isoformat()} | {note} |\n")
        print("saved", dst)
    finally:
        shutil.rmtree(work, ignore_errors=True)

if __name__ == "__main__":
    main()
