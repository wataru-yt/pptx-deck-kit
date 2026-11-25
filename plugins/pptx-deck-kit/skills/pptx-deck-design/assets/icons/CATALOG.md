# アイコンの一覧

`_catalog.png` に、線の絵の見た目がある。

## 線の絵（1色。`scripts/tint_icons.py` で、資料の色の組に合わせて塗り直す）

色ちがい: `<名前>-w.png`（白）、`<名前>-e.png`（強調の色）、`<名前>-r.png`（問題の色）。部品では `d.icon(s, 名前, cx, cy, 大きさ, "w"|"e"|"r")`。

| 名前 | 意味 |
|---|---|
| `user` | 人（1人） |
| `users` | 人（複数） |
| `laptop` | パソコン |
| `desktop` | 画面 |
| `smartphone` | スマホ |
| `tablet` | タブレット |
| `internet` | インターネット |
| `cloud` | クラウド |
| `server` | サーバー |
| `compute` | 計算 |
| `function` | 関数 |
| `container` | コンテナ |
| `database` | データベース |
| `storage` | ストレージ |
| `bucket` | 置き場 |
| `folder` | フォルダ |
| `file` | 文書 |
| `files` | 複数の文書 |
| `spreadsheet` | 表計算 |
| `api` | API |
| `gateway` | 入口 |
| `webhook` | Webhook |
| `plug` | 連携 |
| `queue` | 待ち行列 |
| `network` | ネットワーク |
| `auth` | 認証 |
| `shield` | 守り |
| `lock` | 鍵 |
| `monitoring` | 監視 |
| `logs` | ログ |
| `search` | 検索 |
| `vector-search` | 意味の検索 |
| `ai` | 生成AI |
| `brain` | AIのモデル |
| `bot` | ボット |
| `workflow` | 処理の流れ |
| `settings` | 設定 |
| `chart` | 集計 |
| `code` | コード |
| `terminal` | コマンド |
| `git-branch` | ブランチ |
| `pull-request` | プルリクエスト |
| `mail` | メール |
| `chat` | チャット |
| `building` | 社屋 |
| `pic-people` | 人々 |
| `pic-building` | 会社・ビル |
| `pic-laptop` | パソコン |
| `pic-search` | 検証・確認 |
| `pic-bubble` | 指摘・発言 |
| `pic-box` | 製品・公開物 |
| `pic-contract` | 契約・合意 |
| `pic-calendar` | 日付・期日 |
| `pic-yen` | お金 |
| `pic-star` | 評価・レビュー |
| `pic-pie` | 割合 |
| `pic-play` | 動画 |
| `pic-launch` | 公開・発売 |
| `pic-umbrella` | 予備・備え |
| `pic-book` | 台帳・記録 |
| `pic-factory` | 工場 |
| `pic-fax` | FAX・印刷 |
| `pic-phone` | 電話 |
| `pic-merge` | 合流・マージ |
| `pic-check` | 完了・合格 |

## 他社の製品・会社のロゴ（同梱しない）

Slack・Excel・GitHub などのロゴは、各社の商標なので配布物に入れていない。資料に使うときは、各社の公式サイト（ブランドの素材のページ）から取ってきて、作業フォルダの `icons/` に置く。

```bash
python3 <skills>/pptx-deck-design/scripts/fetch_icon.py slack <公式の SVG か PNG の URL> icons --note "取ってきた場所"
```

ロゴは塗り直さない（`tint_icons.py` は、下の線の絵だけを塗る）。各社の使い方の決まり（大きさ・余白・色）に従う。

無い絵は、線の太さをそろえた SVG（48×48 の枠、線の太さ 3.4、角は丸）から作り、`pic-<名前>.png` で置く。そのあと `tint_icons.py` を実行する。
