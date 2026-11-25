// make-icons.js : 構成図（archDiagram）で使うアイコンを PNG にする（pptx-deck-design、2026-11-08）
// 使い方: NODE_PATH=<react・react-dom・react-icons の入った node_modules>:<sharp の入った node_modules> node make-icons.js [出力先]
//   例: NODE_PATH=<react-icons・react・react-dom の入った node_modules>:<sharp の入った node_modules> node make-icons.js
//   出力先の既定はこのファイルと同じフォルダ（deck-parts.js の隣の icons/）。ファイル名がそのまま archDiagram の icon の名前になる
// 方針:
//   汎用のアイコンは Lucide（線の太さをそろえるため、足りないものだけ同じ線の作りの Tabler）を ink（002B5C）1色で塗る
const React = require("react");
const { renderToStaticMarkup } = require("react-dom/server");
const sharp = require("sharp");
const path = require("path");
const fs = require("fs");
const lu = require("react-icons/lu");
const tb = require("react-icons/tb");

const OUT = path.resolve(process.argv[2] || __dirname);
const PX = 512;
const INK = "#002B5C";

// アイコンの SVG の中身（viewBox と子要素）を取り出す
const svgOf = (Icon, attrs = {}) => renderToStaticMarkup(React.createElement(Icon, { size: PX, ...attrs }));
const pathD = Icon => { const m = svgOf(Icon).match(/ d="([^"]+)"/); if (!m) throw new Error("path が無い"); return m[1]; };
const viewBox = Icon => svgOf(Icon).match(/viewBox="([^"]+)"/)[1];
// 閉じた部分パス（…z m…）を、絶対座標の M で始まる部分に分ける（相対の m は直前の部分の始点からの差）
const splitSub = d => {
  const parts = d.replace(/z(?=\s*[Mm])/g, "z|").split("|").map(p => p.trim());
  let start = [0, 0];
  return parts.map(p => {
    const m = p.match(/^([Mm])\s*(-?[\d.]+)[ ,]?(-?[\d.]+)(.*)$/s);
    const [x, y] = m[1] === "M" ? [+m[2], +m[3]] : [start[0] + +m[2], start[1] + +m[3]];
    start = [x, y];
    // 相対の m の後に続く座標の組は相対の lineto。M に直したら l を明示する
    const rest = m[1] === "m" && /^\s*[-\d.]/.test(m[4]) ? `l${m[4]}` : m[4];
    return `M${x} ${y}${rest}`;
  });
};
const wrap = (vb, body, defs = "") => `<svg xmlns="http://www.w3.org/2000/svg" viewBox="${vb}" width="${PX}" height="${PX}">${defs}${body}</svg>`;

// 他社の製品・会社のロゴは配布しない（各社の商標）。使うときは scripts/fetch_icon.py で公式のものを取ってくる

const GENERIC = [
  ["user", lu.LuUser, "ユーザー（1人）"],
  ["users", lu.LuUsers, "ユーザー（複数）・チーム"],
  ["laptop", lu.LuLaptop, "パソコン"],
  ["desktop", lu.LuMonitor, "デスクトップ・画面"],
  ["smartphone", lu.LuSmartphone, "スマートフォン"],
  ["tablet", lu.LuTablet, "タブレット"],
  ["internet", lu.LuGlobe, "インターネット（地球）"],
  ["cloud", lu.LuCloud, "クラウド"],
  ["server", lu.LuServer, "サーバー"],
  ["compute", lu.LuCpu, "計算の資源（EC2 などの仮想サーバーに相当）"],
  ["function", tb.TbLambda, "関数の実行（Lambda に相当）"],
  ["container", lu.LuContainer, "コンテナ（ECS などに相当）"],
  ["database", lu.LuDatabase, "データベース（RDS・DynamoDB に相当）"],
  ["storage", lu.LuHardDrive, "ストレージ"],
  ["bucket", tb.TbBucket, "オブジェクトの置き場（S3 に相当）"],
  ["folder", lu.LuFolder, "フォルダ・共有ドライブ"],
  ["file", lu.LuFileText, "ファイル・文書"],
  ["files", lu.LuFiles, "複数のファイル"],
  ["spreadsheet", lu.LuFileSpreadsheet, "表計算のファイル"],
  ["api", tb.TbApi, "API"],
  ["gateway", lu.LuWaypoints, "API の入口（API Gateway に相当）"],
  ["webhook", lu.LuWebhook, "Webhook"],
  ["plug", lu.LuPlug, "連携・接続"],
  ["queue", lu.LuInbox, "待ち行列（SQS に相当）"],
  ["network", lu.LuNetwork, "社内のネットワーク（VPC に相当）"],
  ["auth", lu.LuKeyRound, "認証（Cognito・IAM に相当）"],
  ["shield", lu.LuShieldCheck, "守り・権限の確認"],
  ["lock", lu.LuLock, "秘密の情報・暗号化"],
  ["monitoring", lu.LuActivity, "監視（CloudWatch に相当）"],
  ["logs", lu.LuLogs, "記録・ログ"],
  ["search", lu.LuSearch, "検索"],
  ["vector-search", lu.LuScanSearch, "意味の検索（ベクトル検索）"],
  ["ai", lu.LuSparkles, "生成AI（Bedrock などのモデルに相当）"],
  ["brain", lu.LuBrainCircuit, "AI のモデル・推論"],
  ["bot", lu.LuBot, "ボット・エージェント"],
  ["workflow", lu.LuWorkflow, "処理の流れ"],
  ["settings", lu.LuSettings, "設定"],
  ["chart", lu.LuChartColumn, "集計・ダッシュボード"],
  ["code", lu.LuCode, "コード"],
  ["terminal", lu.LuSquareTerminal, "コマンドの実行"],
  ["git-branch", lu.LuGitBranch, "ブランチ"],
  ["pull-request", lu.LuGitPullRequest, "プルリクエスト・レビュー"],
  ["mail", lu.LuMail, "メール"],
  ["chat", lu.LuMessageSquare, "チャット・メッセージ"],
  ["building", lu.LuBuilding2, "社内・拠点"],
];

(async () => {
  fs.mkdirSync(OUT, { recursive: true });
  const toPng = (svg, name) => sharp(Buffer.from(svg)).resize(PX, PX, { fit: "contain", background: { r: 0, g: 0, b: 0, alpha: 0 } }).png().toFile(path.join(OUT, `${name}.png`));
  for (const [name, Icon] of GENERIC) await toPng(svgOf(Icon, { color: INK, strokeWidth: 1.75 }), name);
  console.log(`${GENERIC.length} 個のアイコンを ${OUT} に書き出した`);
})();
