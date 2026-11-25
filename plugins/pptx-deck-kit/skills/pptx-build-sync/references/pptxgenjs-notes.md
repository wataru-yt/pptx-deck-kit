# pptxgenjs の癖と、この方式で詰まりやすい点

pptxgenjs 4.0.1 で確認した内容。

## 書き方

- 色は `"002B5C"`（`#` なし・6桁）。透明度は色に混ぜず、`transparency`（0〜100。大きいほど薄い）で渡す。線は `line.transparency`、塗りは `fill.transparency`、文字は文字オプションの `transparency`。
- 文字を座標どおりに置くには、`margin: 0`・`valign: "top"`・`isTextBox: true` を既定にする。折り返しなしは `wrap: false`。段落は文字列中の `\n` で分かれ、`lineSpacingMultiple: 1.2` が行間120%になる。
- 図形に文字を入れる（丸に「！」など）は `slide.addText("！", { shape: pres.shapes.OVAL, x, y, w, h, fill, line, align: "center", valign: "middle", margin: 0 })`。
- `rectRadius` は `ROUNDED_RECTANGLE` だけに付ける。`ROUND_2_SAME_RECTANGLE` などに付けると `<a:gd name="adj">` が書かれ、PowerPoint で開くと修復ダイアログが出る（書き出しの AppleScript もエラー -9074 で止まる）。
- 2点を結ぶ線は `LINE` の `x,y,w,h` で表す。右上がりは `flipV: true`。縦線は `w: 0`。矢印は `line.beginArrowType` / `endArrowType: "triangle"`。
- 表のセル余白 `margin` と `rowH`・`colW` はインチ。`margin` は `[上, 右, 下, 左]`。罫線はセルの `border: [上, 右, 下, 左]` に `{ type: "none" }` か `{ type: "solid", pt, color }`。表の外枠の高さは PowerPoint が開いたときに再計算されるので、比較では無視する。
- オプションのオブジェクトは pptxgenjs が書き換える。使い回さず、呼び出しごとに新しく作る。
- ハイパーリンクは `hyperlink: { url: encodeURI("../資料/x.xlsx") }`。pptx からの相対パスで書く。
- 注記は `slide.addNotes("...")`。
- **日本語の禁則が効かない。** pptxgenjs は文字に `lang="en-US"` を書くため、PowerPoint で行頭に「、」「）」が来る。書き出し後に全スライドの XML で `lang="en-US"` を `lang="ja-JP"` に置き換える（`assets/build-template.js` の書き出し処理に実装済み）。

## グラフ

- `addChart(pres.charts.BAR, series, opts)`。積み上げは `barGrouping: "stacked"`。軸の上限は `valAxisMaxVal` / `valAxisMinVal` / `valAxisMajorUnit`。
- **棒ごとの透明度は書けない**。書き出し後に zip を開き、`ppt/charts/chartN.xml` の各 `<c:ser>` に `<c:dPt>`（`idx` と `alpha` 付きの `spPr`）を足す。`assets/build-template.js` の `fadeBars` が実装例。`<c:dPt>` は `<c:invertIfNegative>` の直後に入れる。対象のグラフはタイトルの語で見分ける。
- 図の上に重ねる注記（点・線・ラベル）はグラフの座標と一致させる。グラフの内側の位置は自動配置なので、重ねる図形は「グラフを置いた後に実測した座標」で固定する。

## pptx を比べるとき

- 現行 pptx が PowerPoint で保存し直されていると、図形名・言語タグ・東アジア用フォント指定・空の段落・表の枠の大きさが変わる。`pptx_diff.py` はこれらを無視する。
- 文字の途中で色や書式が変わる箇所は、現行 pptx で複数の run に分かれる。見た目が同じなら build.js 側は1つの run でよい。
- ページ数が違うときは `--map 左:右` で対応させる。削除・追加したページは「片側にしか存在しない」と出る。
- 灰色などの色が「意図しない残り」に見えても、現行 pptx に合わせて再現し、その旨をユーザーに伝える。直すかどうかはユーザーが決める。

## 画像化（macOS）

- LibreOffice が入っていない環境では、PowerPoint の AppleScript で PDF にする（`scripts/render_pptx.sh`）。初回は数十秒〜数分かかる。
- AppleScript の変数名に `out` を使わない（PowerPoint の語と衝突する）。真偽値は `as text` で文字列にしてから連結する。
- Python 3.9 では、Claude の pptx スキルの `soffice.py` は動かない（`TemporaryDirectory` の引数の違い）。
