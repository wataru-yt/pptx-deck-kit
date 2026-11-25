---
name: pptx-build-sync
description: pptxgenjsのbuild.jsでPowerPoint資料（.pptx）を生成・管理し、build.jsと実際のpptxを一致させたまま修正を重ねる。新しい資料をbuild.js方式で作る、PowerPointやCodexで手修正されたpptxの内容をbuild.jsへ戻す、build.jsの出力が現行pptxと一致するか検証してから本番ファイルを再生成する、といった場面で使う。「build.jsを更新して」「pptxとbuild.jsを揃えて」「スクリプトで資料を再生成できるようにして」と言われたときに使用する。
---

# PPTX Build Sync

資料の中身を `build.js`（pptxgenjs）に持たせ、pptx はその出力として扱う。文言・表・図の座標・グラフの値がコードに残るので、直しが差分で追え、同じ資料を何度でも再生成できる。

## 守ること

- **pptx と build.js は一致している状態を保つ。** pptx を PowerPoint や別ツールで直接直したら、同じ作業の中で build.js に反映する。反映しないまま build.js を実行すると、直接の修正が消える。
- **build.js を実行する前に、出力先（`OUT`）を確認する。** 本番ファイルを上書きするのは、ユーザーが指示したときだけ。それまでは一時フォルダへ出力する（`node build.js 出力先.pptx`、またはコピーの `OUT` を書き換える）。
- **本番を上書きする前に、現行ファイルをバックアップする。** pptx が git 管理外のことが多く、上書きすると戻せない。一時フォルダに `backup-日付.pptx` として残す。
- **PowerPoint で開いている間は上書きしない。** 同じフォルダに `~$ファイル名.pptx` があれば開いている。閉じてもらうか、保存されていない変更がないか確認する。
- **一致の確認を先に行う。** 下の「一致の確認」を通ってから本番へ出力する。

## 環境

- Node: `pptxgenjs`（`require` できる場所が必要）。プラグイン（pptx-deck-kit）として入れたときは自動で入っているので、`NODE_PATH="${CLAUDE_PLUGIN_ROOT}/node_modules"` を付けて実行する。スキルを直接置いたときは、作業フォルダで `npm install pptxgenjs@4.0.1 jszip@3.10.2` する
- Python 3.9 以上: `lxml`、`Pillow`
- 画像化: `pdftoppm`（poppler）と、Microsoft PowerPoint（macOS）または LibreOffice（`soffice`）

## 手順

### A. 新しく作る

1. 見せ方は `pptx-deck-design` スキルの `assets/deck-parts.js` と、題材に近い見本（`assets/sample-*.js`）を作業フォルダ（例: `.claude/scripts/<資料名>/`）へコピーして始める。
2. `OUT`、`LINK`（資料内リンク）、`MINI`・`AGENDA`（ミニ目次・目次）、スライドを書く。ページ番号は `PAGE.キー = page` で記録し、目次は最後に描く（ページ追加で番号がずれない）。見せ方の部品は `deck-parts.js` の `createDeck` を使う（`pptx-deck-design` の SKILL.md）。
   - `pptx-deck-design` の部品（`deck-parts.js`）で作る資料では、章立ては `createDeck` の `nav` に1回だけ書き、アジェンダは `agendaPage`、ページの名前は `base` の `key` で付ける。`MINI`・`AGENDA`・`PAGE.キー` は書かない
3. 一時フォルダへ出力して確認し、ユーザーの指示があれば本番へ出す。

### B. pptx が先に直されている（build.js を追いつかせる）

1. 現行 pptx を一時フォルダへコピーし、build.js も一時出力に向けて実行する。
2. 比べる:
   ```bash
   python3 scripts/pptx_diff.py 一時出力.pptx 現行.pptx            # 図形単位の差分
   python3 scripts/pptx_diff.py 一時出力.pptx 現行.pptx --summary  # ページごとの件数だけ
   python3 scripts/pptx_diff.py 一時出力.pptx 現行.pptx --map 9:10 # ページ数がずれているとき（左:右）
   ```
   差分は「左のみ（build.js の出力）」「右のみ（現行 pptx）」で出る。右のみの図形が build.js に足すべきもの。
3. 差分の種類ごとに build.js を直す。
   - 文言・数値・色・座標が違う → 該当スライドのコードを直す。
   - ページが増えた・消えた → スライドのブロックを足す・消す。ページ番号やミニ目次は自動で変わる。
   - 図形が現行にだけある → 同じ座標・書式で足す。繰り返しの形は関数にする（図の座標系、ラベルなど）。
   - グラフの値・色・点ごとの透明度が違う → `pptx_inventory.py` の `chart` 行を見て揃える。点ごとの透明度は書き出し後に XML へ足す（`references/pptxgenjs-notes.md`）。
4. 差分が「座標の0.005インチ未満」「空の段落」「図形名」だけになるまで繰り返す。これらは PowerPoint の再保存で出る揺れで、直さなくてよい。
5. 「一致の確認」へ進む。

### C. 通常の修正

1. build.js を直す。
2. 一時フォルダへ出力し、`pptx_inventory.py 一時出力.pptx --slides N` で変更したページを確認する。
3. 直前の本番ファイルとの `pptx_diff.py` で、意図した変更だけが出ているか見る。
4. 描画を比べる（下記）。問題がなければ、ユーザーの指示で本番へ出力する。

## 一致の確認

```bash
python3 scripts/pptx_diff.py 一時出力.pptx 現行.pptx --summary   # 差分が既知の揺れだけか
scripts/render_pptx.sh 一時出力.pptx r_new 70                     # PowerPoint（なければ soffice）で画像化
scripts/render_pptx.sh 現行コピー.pptx r_cur 70
python3 scripts/render_compare.py r_cur r_new cmp                 # ページごとの差分画素数と比較画像
```

- 差分画素数は、数百以下なら字形の差。数千以上は図形のずれ・欠落を疑い、`cmp/cmp-N.png`（下段左が差分）を見る。
- 現行 pptx を直接 PowerPoint で開いて書き出さない。コピーを渡す。
- 比べていないもの: 重なり順、アニメーション・画面切り替え、スライドの非表示、PowerPoint が付ける属性（言語タグ・東アジア用フォント指定・図形名）。重なりが見た目に出るページは、画像比較で確認する。

## ファイル

- `scripts/pptx_inventory.py` — pptx の全図形・表・グラフ・ノート・リンク先を、座標と書式つきの短いテキストに書き出す
- `scripts/pptx_diff.py` — 2つの pptx をスライドごとに図形単位で比べる
- `scripts/render_pptx.sh` — pptx を PDF とページ画像にする
- `scripts/render_compare.py` — 画像フォルダ2つをページごとに比べる
- `references/pptxgenjs-notes.md` — pptxgenjs の癖と、この方式で詰まりやすい点
