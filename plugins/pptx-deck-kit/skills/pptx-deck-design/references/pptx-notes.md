# pptxgenjs で崩れやすい点

一般的な癖は `pptx-build-sync` の `references/pptxgenjs-notes.md` を先に読む。ここはデザインの再現で効くものだけ。

- 禁則: 書き出した pptx の `lang="en-US"` を `ja-JP` に置き換えないと、行頭の句読点や閉じ括弧が残る。`finalize` が行う
- リードの改行: `leadBreak` は、1行に収まれば1行、2行なら「。」で改行し、1文が収まらないときだけ「、」の直後で折る。それでも3行になるときは文字を縮めず、言葉を縮める
- 文字幅の見積もり: 全角 1em、半角 0.6em。9.5pt で半角 0.08in・全角 0.15in。`deck-parts.js` の警告（タイトルが1行に収まらない、リードが3行、本文の下端超え、1〜2文字だけの行）は目安で、数字の途中の改行（「約7／割」）は検出できない。語の切れ目で `\n` を入れる
- 1行のラベルは `wrap: false`。付箋は `fit: "none"`
- `ROUND_2_SAME_RECTANGLE` に `rectRadius` を書くとファイルが壊れる
- 矢羽（CHEVRON）の中の文字枠は狭い。文字は別のテキストボックスで重ねる
- 線は負の幅にしない（`segLine` は flipV で向きを表す）。負の幅でファイルが壊れた例がある
- 動画の細い枠は動画の後ろに置く（前に置くと再生を妨げる）
- 棒ごとの強調・薄めは pptxgenjs では書けない。書き出し後に XML の dPt を足す（`barChart` の `highlight` と `finalize`）
- 描画の確認: `bash <skills>/pptx-build-sync/scripts/render_pptx.sh <pptx> <出力dir> 70`（`<skills>` は Claude Code が `~/.claude/skills`、Codex が `~/.codex/skills`） で PDF とページ画像にする。PDF で正しくても PowerPoint の編集画面で崩れることがあるので、仕上げでは編集画面でも開く
