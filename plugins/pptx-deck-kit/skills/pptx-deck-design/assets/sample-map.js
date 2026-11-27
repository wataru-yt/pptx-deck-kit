// sample-map.js : 2軸の図（ポジショニングマップ）、ファネル、アイコンの構成図の見本（matrix2x2・funnel・archDiagram。内容はすべて架空。会社名は A社〜G社）
// 実行: NODE_PATH=<pptxgenjs と jszip の入った node_modules> node sample-map.js <出力先.pptx>
//
// 論旨：すぐ使えて業務まで任せられる社内AIは、まだ無い。当社の試作版は使い始めた人の半分以上が2週目に離れ、
//       理由は回答の根拠が見えないこと。1つの入口と社内文書の検索で、根拠を示して答える。
//
// 色の組：blue（強調 2F6BFF）。社内のAI活用（IT・ソフトウェア）の提案のため。線の絵（icons/）も blue の色で塗ってある。
//         問題（離れる段）は赤（accent）。
//
// 形の割り当て（1つの形は1つの意味にだけ使う）
//   matrix2x2   … 製品を2軸で位置づける（P.3 図だけ → P.4 同じ図を敷いて左上の3社を言う）
//   funnel      … 利用者が段階ごとに絞られる（P.5 図だけ → P.6 同じ図を敷いて離れる理由を言う）
//   archDiagram … 基盤の仕組み（P.7 図だけ → P.8 同じ図を敷いて文書の取り込みを言う）
//
// 見本の約束
// - 題はアジェンダの項目名だけ（nav に1回書き、base の題は null にして章から取る）。主張はリード
// - 全ページで見る所を1つ宣言し（focus）、強調はそこにだけ置く。2軸とファネルはグラフの扱いなので、見る所の値を focusValue（28pt）で1つ書く
// - 1つの図で2つ言うときは、series で同じ図を敷いた続きのページに分ける（2枚目は spot か only で見る所だけ残し、前面に箱を1つ）
// - 補足の sub は使わない。図の外に読み取りの欄（matrix2x2 の side）を置かず、言うことは続きのページの前面の箱で言う
// - 他社の製品のロゴは同梱しないので、構成図は汎用の線の絵（icons/CATALOG.md）で描く。実在の製品名を図に出すなら、公式のアイコンを取ってくる（決まり5）
const path = require("path");
const pptxgen = require("pptxgenjs");
const { createDeck } = require("./deck-parts");

const OUT = path.resolve(process.argv[2] || path.join(__dirname, "sample-map.pptx"));
const CHECK = process.env.ECHO_CHECK === "1"; // 1 にすると続きのページを前面なしで描く（figure_diff.py で図が動いていないかを確かめる用）
const pres = new pptxgen();
const d = createDeck(pres, {
  audience: "insider", // 聞き手の前提（決まり24）。見本は部品の使い方を見せるので、前提を共有済みとして描く
  title: "社内のAI活用の進め方",
  theme: "blue",
  footerText: "社内のAI活用の進め方",
  nav: [
    { name: "製品の位置づけ", ch: 1 },
    { name: "利用の定着", ch: 2 },
    { name: "基盤の構成", ch: 3 },
  ],
});

d.cover({ title: "社内のAI活用の進め方", date: "2027年2月" });
d.agendaPage();

// ===== 1. 製品の位置づけ（位置づける）：バブル（円の面積＝利用企業数）、自社だけ強調の色、空白地帯 =====
// 点の名前は社名だけ（区分の説明は象限の名前で言う）。読み取りは右の欄に置かず、続きのページの箱で言う
d.setChapter(1);
const MAP = {
  y: 2.55,
  xAxis: { name: "導入の手軽さ", low: "設定に数か月", high: "その日に使える" },
  yAxis: { name: "業務への入り込み", low: "汎用の対話だけ", high: "業務の手順まで任せる" },
  quads: ["組み込みの基盤", "", "", "汎用の対話"],
  size: { unit: "社", max: 2500 },
  points: [
    { name: "A社", x: 0.21, y: 0.75, size: 900 }, // 左上の象限の名前の下に置く
    { name: "B社", x: 0.8, y: 0.22, size: 2500 },
    { name: "C社", x: 0.68, y: 0.3, size: 700 },
    { name: "D社", x: 0.3, y: 0.68, size: 450 },
    { name: "E社", x: 0.57, y: 0.24, size: 1100 },
    { name: "F社", x: 0.2, y: 0.36, size: 300 },
    { name: "G社", x: 0.37, y: 0.82, size: 520 },
    { name: "当社", x: 0.6, y: 0.68, size: 12, self: true, pos: "l" },
  ],
  zone: { x: 0.83, y: 0.8, w: 0.27, h: 0.28, label: "空白地帯" },
  move: { from: 7 },
};
// 円の大きさの意味は、図の左上の見出しに書く（右に凡例の欄を作らない）
const pos = d.series((s, x) => { d.figHead(s, "円の大きさ＝利用企業数"); return d.matrix2x2(s, { ...MAP, ...x }); }, { check: CHECK });
{
  const { s, r } = pos.page({ lead: "試作版の当社（12社）は、右上の空白地帯を狙う。", focus: "当社の点（12社）と右上への矢印" });
  const [px, py] = r.pt(7);
  d.focusValue(s, "12社", { x: px - 1.0, y: py - 0.68, w: 1.4, align: "center" });
}
{
  const { s, r } = pos.page({
    lead: "業務まで任せられる3社は、導入に3か月以上かかる。", focus: "左上の象限の3社（A社・D社・G社）",
    spot: r => r.quad(0),
    // 箱は左下の象限の空き（F社の右）に置き、引き出し線は垂直に D社の円の下を指す
    box: r => { const [dx, dy] = r.pt(3), dr = r.d(3) / 2, y = r.py(0.42); return { x: dx - 0.4, y, w: 3.0, h: 1.1, head: "手順の登録が要る", body: "業務ごとに登録する。", from: [dx, y], to: [dx, dy + dr + 0.04] }; },
  });
  const q = r.quad(0);
  d.focusValue(s, "3か月以上", { x: q.x + 0.15, y: r.py(0.57) - 0.25, w: 2.2 });
}

// ===== 2. 利用の定着（絞り込む）：段は一定の比率で狭め、値は右の数字で。離れる段だけ赤 =====
// 説明の列は、段の定義を短い札で1行だけ（文を積まない）。離れる人数を見る所の値にし、図を描く関数に入れて続きのページでも敷く
d.setChapter(2);
const FUN = {
  x: 1.25, unit: "人", drop: 2,
  stages: [{ name: "対象の社員", value: 800 }, { name: "登録", value: 620 }, { name: "初めての質問", value: 480 }, { name: "2週目利用", value: 210 }, { name: "毎週使う", value: 150 }],
  cols: [{ head: "段階の定義", cells: ["現場の作業者を除く全社員", "チャットにボットを追加", "質問を1回以上送った", "2週目にも質問を送った", "4週続けて週1回以上"] }],
};
const drawFunnel = (s, x) => {
  const r = d.funnel(s, { ...FUN, ...x });
  const b = r.band(3);
  d.focusValue(s, "−270人", { x: 0.6, y: b.y - 0.3, w: b.x - 0.7, align: "right" }, { problem: true });
  return r;
};
const fun = d.series(drawFunnel, { check: CHECK });
fun.page({ lead: "最も離れるのは2週目で、480人のうち270人が使わなくなる。", focus: "2週目の手前で離れる人数（−270人）と率（44%）" });
fun.page({
  lead: "離れた理由は、回答の根拠が示されないことだ。", focus: "2週目の手前で離れる人数（−270人）",
  // 見る所：離れる段の前後（段・人数・率）と、見る所の値だけを元の濃さで残す
  spot: r => { const a = r.band(2), b = r.band(3), c = r.cell(0, 2); return { x: 0.6, y: a.y - 0.14, w: c.x - 0.25 - 0.6, h: b.y + b.h - a.y + 0.28 }; },
  // 箱は、離れる段の定義のセル2つ（薄くしたもの）をちょうど覆い、引き出し線は水平に率の赤枠を指す
  box: r => { const c2 = r.cell(0, 2), c3 = r.cell(0, 3), [cx, cy] = r.conv(2); return { x: c2.x, y: c2.y, w: c2.w, h: c3.y + c3.h - c2.y, head: "回答の根拠が分からない", body: "出典が出ず、確認し直す手間がかかる。", from: [c2.x, cy], to: [cx + 0.62, cy] }; },
});

// ===== 3. 基盤の構成（仕組み）：アイコンのノード、まとまりの領域、鉤の矢印 =====
// 格子は 5列 × 3行。左から 利用者 → 社内で使う道具 → 社内のAI基盤 → 外部のAI。下に箱1つ分の空きを残す（h）
// 名前が大きいので、札を小さくし（tile）、行を 0.24・1.19・2.14 に置く
d.setChapter(3);
const ARCH = {
  cols: 5, rows: 3, tile: 0.48, h: 3.75,
  groups: [
    { id: "tools", name: "社内の道具", c: [1, 1], r: [-0.08, 2.18], style: "dash", inset: 0.18 }, // 枠の左右を内へ寄せて、利用者からの線のラベルが入る長さを取る（上下は r をずらして他のまとまりとそろえる）
    { id: "base", name: "社内のAI基盤", icon: "cloud", c: [2, 3], r: [0, 2.1], style: "fill" },
    { id: "ai", name: "外部のAI", c: [4, 4], r: [0, 2.1], style: "dash" },
  ],
  nodes: [
    { id: "emp", c: 0, r: 0.24, icon: "users", name: "社員 800人" },
    { id: "dev", c: 0, r: 1.19, icon: "laptop", name: "開発者 60人" },
    { id: "chat", c: 1, r: 0.24, icon: "chat", name: "社内チャット" },
    { id: "git", c: 1, r: 1.19, icon: "git-branch", name: "コード管理" },
    { id: "folder", c: 1, r: 2.14, icon: "folder", name: "共有フォルダ" },
    { id: "gw", c: 2, r: 0.24, icon: "gateway", name: "AIの入口", strong: true },
    { id: "route", c: 3, r: 0.24, icon: "workflow", name: "AIの振り分け" },
    { id: "log", c: 2, r: 1.19, icon: "database", name: "利用の記録" },
    { id: "search", c: 3, r: 1.19, icon: "vector-search", name: "文書の検索" },
    { id: "ingest", c: 3, r: 2.14, icon: "function", name: "文書の取り込み" },
    { id: "aiTalk", c: 4, r: 0.24, icon: "ai", name: "AI（日常の質問）" },
    { id: "aiCode", c: 4, r: 1.19, icon: "code", name: "AI（コード・長文）" },
    { id: "aiTable", c: 4, r: 2.14, icon: "spreadsheet", name: "AI（表・画像）" },
  ],
  edges: [
    { from: "emp", to: "chat", label: "質問" },
    { from: "dev", to: "git", label: "依頼" },
    { from: "chat", to: "gw" },
    { from: "git", to: "gw", route: "hvh", mid: 1.5 },
    { from: "gw", to: "route" },
    { from: "gw", to: "log" },
    { from: "route", to: "search", both: true }, // 縦の短い線にはラベルを置かない（ノードの名前で分かる）
    { from: "route", to: "aiTalk" },
    { from: "route", to: "aiCode", route: "hvh", mid: 3.5 },
    { from: "route", to: "aiTable", route: "hvh", mid: 3.5 },
    { from: "folder", to: "ingest", dashed: true, label: "毎晩" },
    { from: "ingest", to: "search" },
  ],
};
const arch = d.series((s, x) => d.archDiagram(s, { ...ARCH, ...x }), { check: CHECK });
arch.page({ lead: "社員と開発者は、1つの入口を通して3つのAIを使う。", focus: "AIの入口のノード（強調の枠）" });
arch.page({
  lead: "共有フォルダの文書は毎晩取り込み、回答の根拠にする。", focus: "文書の取り込みのノード",
  only: ["ingest"],
  // 箱は図の下の空きに置き、引き出し線は垂直に取り込みのノードの名前の下を指す
  box: r => { const n = r.node("ingest"), y = n.bottom + 0.1; return { x: n.cx - 2.5, y, w: 5.0, h: 0.55, head: "更新分を権限つきで取り込む", from: [n.cx, y], to: [n.cx, n.bottom + 0.03] }; },
});

d.finalize(OUT);
