#!/bin/zsh
# pptx を PDF とページ画像（JPEG）に書き出す。
# 使い方: render_pptx.sh deck.pptx out_dir [dpi=70]
#   macOS: Microsoft PowerPoint があればそれで書き出す（実機と同じ描画）。なければ LibreOffice (soffice)。RENDER_WITH=libreoffice で LibreOffice を使う。
#   画像化に pdftoppm（poppler）を使う。
# PowerPoint はサンドボックスのため、任意のフォルダのファイルを開くと「ファイル アクセスを許可」のダイアログが出て止まることがある。
# これを避けるため、PowerPoint のコンテナ内の作業フォルダへコピーしてから開く（ダイアログは出ない）。
# 開いた文書の名前を確認してから書き出す（別の文書を書き出す事故を防ぐ）。元ファイルは変更しない。
set -e
IN=$(cd "$(dirname "$1")" && pwd)/$(basename "$1")
OUT=$2
DPI=${3:-70}
mkdir -p "$OUT"
OUT=$(cd "$OUT" && pwd)
BASE=$(basename "${IN%.*}")
PDF="$OUT/$BASE.pdf"
find "$OUT" -maxdepth 1 -name 's-*.jpg' -delete
rm -f "$PDF"
# RENDER_WITH=libreoffice で、PowerPoint があっても LibreOffice で書き出す（PowerPoint が使えないとき）
# PowerPoint で書き出せなければ（サインイン切れ・ダイアログなど）、自動で LibreOffice に切り替える
PPT_OK=0
if [ "${RENDER_WITH:-}" != "libreoffice" ] && [ -d "/Applications/Microsoft PowerPoint.app" ]; then
(
set -e
  WORK="$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Documents/pptx-build-sync-tmp"
  mkdir -p "$WORK"
  TMPNAME="render-$$-$BASE.pptx"
  cp "$IN" "$WORK/$TMPNAME"
  rm -f "$WORK/render-$$-$BASE.pdf"
  osascript <<OSA
tell application "Microsoft PowerPoint"
	-- 開くのも、やり直す（コピーした直後に -9074 で失敗することがある）
	set opened to false
	repeat 15 times
		try
			open POSIX file "$WORK/$TMPNAME"
			set opened to true
			exit repeat
		on error
			delay 2
		end try
	end repeat
	if not opened then error "PowerPoint で開けない（-9074 が続く）。PowerPoint を開いて、サインイン（サブスクリプション）の表示やダイアログが無いかを確かめる"
	set found to false
	repeat 30 times
		repeat with i from 1 to (count of presentations)
			if (name of presentation i) is "$TMPNAME" then
				set found to true
				exit repeat
			end if
		end repeat
		if found then exit repeat
		delay 1
	end repeat
	if not found then error "開いた文書が見つからない: $TMPNAME"
	set pr to presentation "$TMPNAME"
	-- 開いた直後は読み込みが終わっておらず、保存が -9074 で失敗することがある。少し待って、やり直す
	set saved to false
	repeat 15 times
		try
			save pr in POSIX file "$WORK/render-$$-$BASE.pdf" as save as PDF
			set saved to true
			exit repeat
		on error
			delay 2
		end try
	end repeat
	if not saved then
		close pr saving no
		error "PDF に書き出せない（保存が失敗し続ける）。PowerPoint がサインインしていない・サブスクリプションが切れていると、保存できない（読み取り専用になる）"
	end if
	close pr saving no
end tell
OSA
  mv -f "$WORK/render-$$-$BASE.pdf" "$PDF"
  rm -f "$WORK/$TMPNAME"
) || true
# 成否は PDF ができたかで決める（( … ) || … の中では set -e が効かず、途中の失敗を拾えないため）
if [ -s "$PDF" ]; then PPT_OK=1; else
  echo "[render] PowerPoint で書き出せなかった（サインインかサブスクリプションを確かめる）。LibreOffice で書き出す" >&2
  rm -f "$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Documents/pptx-build-sync-tmp/render-$$-$BASE.pptx"
fi
fi
if [ "$PPT_OK" != 1 ]; then
  SOFFICE=$(command -v soffice || true)
  [ -z "$SOFFICE" ] && [ -x "/Applications/LibreOffice.app/Contents/MacOS/soffice" ] && SOFFICE="/Applications/LibreOffice.app/Contents/MacOS/soffice"
  if [ -z "$SOFFICE" ]; then
    echo "PowerPoint も LibreOffice（soffice）も見つからない。どちらかを入れる" >&2
    exit 1
  fi
  # LibreOffice を画面なし（headless）で動かすと、Mac では OS の日本語の書体が見えず、文字が四角になる。
  # 専用のプロファイル（利用者の LibreOffice の設定は触らない）の fonts/ に書体をつなぐ。
  # 游ゴシック（Yu Gothic）は Mac では Office の中にあるので、あればそれをつなぐ（PowerPoint と同じ書体になる）
  PROFILE="${XDG_CACHE_HOME:-$HOME/.cache}/pptx-deck-kit/lo-profile"
  mkdir -p "$PROFILE/user/fonts"
  FONT=""
  YUDIR="/Applications/Microsoft PowerPoint.app/Contents/Resources/DFonts"
  if ls "$YUDIR"/YuGoth*.ttc >/dev/null 2>&1; then
    for f in "$YUDIR"/YuGoth*.ttc; do ln -sf "$f" "$PROFILE/user/fonts/"; done
    # テーマで書体を替えたとき（メイリオ・MS ゴシック・游明朝）も同じ書体で描けるよう、Office にあればつなぐ
    for f in "$YUDIR"/meiryo*.ttc "$YUDIR"/msgothic.ttc "$YUDIR"/yumin*.ttf; do [ -f "$f" ] && ln -sf "$f" "$PROFILE/user/fonts/"; done
  elif [ "$(uname)" = "Darwin" ]; then
    for f in /System/Library/Fonts/ヒラギノ角ゴシック\ W3.ttc /System/Library/Fonts/ヒラギノ角ゴシック\ W6.ttc; do [ -f "$f" ] && ln -sf "$f" "$PROFILE/user/fonts/"; done
    FONT="${RENDER_FONT:-Hiragino Sans}"
  else
    FONT="${RENDER_FONT:-Noto Sans CJK JP}"
  fi
  SRC="$IN"
  if [ -n "$FONT" ]; then
    # 游ゴシックが無いときは、書体の名前を入っている日本語の書体に置き換えた一時の写しを書き出す
    SRC="$OUT/lo-$$-$BASE.pptx"
    python3 - "$IN" "$SRC" "$FONT" <<'PY'
import sys, zipfile
src, dst, font = sys.argv[1:4]
zin = zipfile.ZipFile(src); zout = zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED)
for it in zin.infolist():
    d = zin.read(it.filename)
    if it.filename.endswith(".xml"):
        for name in ("Yu Gothic", "游ゴシック", "Yu Gothic UI", "YuGothic"):
            d = d.replace(('typeface="%s"' % name).encode("utf-8"), ('typeface="%s"' % font).encode("utf-8"))
    zout.writestr(it, d)
zout.close()
PY
  fi
  "$SOFFICE" -env:UserInstallation="file://$PROFILE" --headless --convert-to pdf --outdir "$OUT" "$SRC" >/dev/null
  if [ "$SRC" != "$IN" ]; then mv -f "$OUT/lo-$$-$BASE.pdf" "$PDF"; rm -f "$SRC"; fi
fi
# 画像にする：pdftoppm（poppler）があればそれで、無ければ pdf2png.py（PyMuPDF・Pillow）で 01.png… と4ページずつの一覧 sheetN.png を作る
if command -v pdftoppm >/dev/null 2>&1; then
  pdftoppm -jpeg -r "$DPI" "$PDF" "$OUT/s"
  find "$OUT" -maxdepth 1 -name 's-*.jpg' | wc -l | xargs echo "pages:"
else
  python3 "$(dirname "$0")/pdf2png.py" "$PDF" "$OUT"
fi
