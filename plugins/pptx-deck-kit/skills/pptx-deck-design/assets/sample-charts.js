// sample-charts.js : 王道のグラフを1種類1ページで描く見本（31種・32ページ。内容はすべて架空）。一覧と選び方は references/charts.md
//   ページ：P.2〜7 構成比、P.8〜14 比較、P.15〜16 増減の分解、P.17〜23 推移、P.24〜30 分布・関係、P.31〜32 市場規模
//   スモールマルチプルと KPI の数値タイルは、見本から外した（同じデータを区分ごとに小さく並べる見せ方・カードの並びは、意味をなさないため）
// 実行: NODE_PATH=<pptxgenjs と jszip の入った node_modules> node sample-charts.js <出力先.pptx>
//
// 論旨：架空の家電の会社の上期の業績を、構成比・比較・増減・推移・分布・市場規模の順に、1ページ1つのグラフで見せる。
//
// 色の組：blue（強調 2F6BFF）。見本は題材を選ばないので、灰の中で見る所が最も目立つ青にした（実際の資料では、題材で選び直す）。問題は赤（accent）。
//
// 形の割り当て：この見本は、部品の使い方を見せるために、同じ部品を別の中身で使う（colBars を対の棒・正負の棒・ヒストグラムに、など）。
//   実際の資料では、1つの形を1つの意味にだけ使い、ここに「どの形を何に使うか」を書く（SKILL.md「形・ページ分け・見る所・アイコンの決まり」の1）。
//
// 見本の約束（SKILL.md の決まり 7・16・22・23）
// - 題はアジェンダの項目名だけ（createDeck の nav から取る。base の題は null）。主張はリードの1文に書き、見る所の名前か値を名指しする
// - 全ページで見る所（focus）を1つ宣言する。強調の色は見る所の棒・区画・点・線だけに付け、他はすべて灰（比べる相手は濃い灰 C.faint、残りは薄い灰）
// - 見る所の値は focusValue（36pt・太字・強調の色、問題なら赤）で1つだけ書く。置く所は、その要素の普通の値の位置（棒の上・点の上・区画の中・行の値の列）。離れた所に置かない
//   主役の値を描く口のある部品（colBars の focus、hBars・slopeChart・bridge・waterfall の要素の focusValue、lineChart・heatmap の focusValue、tamSamSom の strong）は、
//   部品が 36pt の値を普通の位置に描き、隣の棒・値との重なりもよけるか警告する。手で focusValue を置くのは、口の無い部品（円・積み上げ・メッコなど）だけ
//   手で置くときは、区画の中（強調の色の塗りの上）は onFill で白にし、部品も同じ値を描くなら、書式の関数でその値だけ空の文字にする
// - 本文の領域（y 2.2〜6.6）を使い切る。部品の描く範囲（top・bottom・base・area・y0）は、既定のまま使う
// - 部品が区画ごとの強調を持たない図（積み上げ・メッコ・帯でつなぐ積み上げ・レーダーなど）は、戻り値の座標で、見る所の区画だけを強調の色で上から塗る
// - 補足の小さな文字（sub）は使わない。何の数値か・単位・時点は figHead に1行
const path = require("path");
const pptxgen = require("pptxgenjs");
const { createDeck } = require("./deck-parts");

const OUT = path.resolve(process.argv[2] || path.join(__dirname, "sample-charts.pptx"));
const pres = new pptxgen();
const d = createDeck(pres, {
  audience: "insider", // 聞き手の前提（決まり24）。見本は部品の使い方を見せるので、前提を共有済みとして描く
  title: "王道のグラフの見本",
  theme: "blue",
  footerText: "王道のグラフの見本",
  nav: [{ name: "構成比", ch: 1 }, { name: "比較", ch: 2 }, { name: "増減の分解", ch: 3 }, { name: "推移", ch: 4 }, { name: "分布・関係", ch: 5 }, { name: "市場規模", ch: 6 }],
});
const { C, M, W, BW, FS } = d;
// 灰の4段（濃い→薄い）。比べる相手は G1、残りは G2〜G4
const G1 = C.faint, G2 = C.bar, G3 = "DADDE1", G4 = "E9EBEF";
// 散布図の点を置く範囲。既定（y 2.75）だと縦軸の名前が figHead の行に重なるので、軸の名前の分だけ下げ、下は本文の下端近くまで使う
const SC_AREA = { x: M + 0.95, y: 2.95, w: BW - 1.6, h: 3.3 };
// colBars の棒の下端。既定（6.15）だと区分名が本文の下端 6.6 を超えて警告が出るので、区分名が 6.6 に収まる 6.05 にする
const COL_BOTTOM = 6.05;
const MONTHS = ["4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月", "1月", "2月", "3月"];
// 見る所の区画を、強調の色で上から塗る（区画の境の白い線も部品と同じにする）
const focusSeg = (s, x, y, w, h) => d.rect(s, x, y, w, h, C.emph, { color: C.white, width: 0.75 });
const dot = (s, cx, cy, dd, color) => s.addShape(pres.shapes.OVAL, { x: cx - dd / 2, y: cy - dd / 2, w: dd, h: dd, fill: { color }, line: { color: C.white, width: 1 } });

d.cover({ title: "王道のグラフの見本", date: "2026年11月" });
d.agendaPage();

// ===== 1. 構成比 =====
d.setChapter(1);

// 円：区分5つまで、大きい順に12時から時計回り、主役だけ強調の色。名前・割合・前年差を区分の外に
// 主役の割合は部品のラベルを半角の空白にし（fmt。空の文字だと行が詰まる）、空いた行に focusValue で 28pt
{
  const s = d.base(null, "上期の売上120億円のうち、法人向けが42%を占めた。", { focus: "法人向けの区分と42%" });
  d.figHead(s, "区分別の売上の構成比（2026年度上期、合計120億円）");
  const tot = 120, p = d.pieChart(s, {
    labelW: 3.3, // 1行の注（金額と前年差）が折れない幅
    fmt: it => (it.strong ? " " : Math.round(it.value / tot * 100) + "%"),
    // 注（金額と前年差）は主役の区分だけに付ける（4つの区分に3行ずつ積むと、左の列で名前と注が詰まって区切りが読めない）
    // 主役の注の頭に改行を入れ、36pt の割合を置く行の高さを空ける
    items: [
      { name: "法人向け", value: 50.4, note: "\n50.4億円・前年差+6pt", strong: true },
      { name: "個人・店頭", value: 26.4 },
      { name: "個人・通販", value: 21.6 },
      { name: "保守・修理", value: 14.4 },
      { name: "その他", value: 7.2, other: true },
    ],
  });
  // 主役のラベルは区分の中央の角度の外（右の列）。名前と前年差の間の空けた行に、割合を大きく書く
  const m = p.mid(0) * Math.PI / 180, ly = p.cy + Math.sin(m) * (p.r + 0.2);
  d.focusValue(s, "42%", { x: p.cx + p.r + 0.43, y: ly - 0.21, w: 2.4 });
}

// ドーナツ：中央に合計。区分の外に割合と1行の値
{
  const s = d.base(null, "月8,400件の問い合わせのうち、50%が電話で届く。", { focus: "電話の区分と50%" });
  d.figHead(s, "経路別の問い合わせの構成比（2026年9月）");
  const tot = 8400, p = d.pieChart(s, {
    donut: true, center: { value: "8,400件", label: "月の問い合わせ" },
    fmt: it => (it.strong ? " " : Math.round(it.value / tot * 100) + "%"),
    items: [ // 注は主役だけ（円と同じ）
      { name: "電話", value: 4200, note: "\n平均の待ち 6分", strong: true },
      { name: "Webのフォーム", value: 2016 },
      { name: "チャット", value: 1344 },
      { name: "メール", value: 840 },
    ],
  });
  const m = p.mid(0) * Math.PI / 180, ly = p.cy + Math.sin(m) * (p.r + 0.2);
  d.focusValue(s, "50%", { x: p.cx + p.r + 0.43, y: ly - 0.21, w: 2.4 });
}

// 100%積み上げ縦棒：構成比を区分ごとに比べる。区分の中に割合、区分名は右端の棒の横。区分は灰の濃淡、見る所の1区画だけ強調の色、その割合を区画の中に 28pt（白）
{
  const s = d.base(null, "20代は購入の72%が通販で、年代が上がるほど店頭が増える。", { focus: "20代の通販の区画と72%" });
  d.figHead(s, "年代別の購入経路の構成比（2026年度上期の購入 3.2万件）");
  const r = d.stackBars(s, {
    pct: true,
    pctFmt: v => (Math.round(v * 100) === 72 ? "" : Math.round(v * 100) + "%"),
    cats: ["20代", "30代", "40代", "50代", "60代", "70代以上"],
    series: [
      { name: "店頭", color: G3, values: [18, 24, 31, 40, 52, 61] },
      { name: "訪問・電話", color: G4, values: [10, 12, 14, 16, 18, 22] },
      { name: "通販", color: G1, values: [72, 64, 55, 44, 30, 17], main: true },
    ],
  });
  const [t, b] = r.seg[0][2];
  focusSeg(s, r.bx(0), t, r.bw, b - t);
  d.focusValue(s, "72%", { x: r.bx(0), y: (t + b) / 2 - 0.25, w: r.bw, align: "center" }, { onFill: true });
}

// 横の100%積み上げ：アンケートの回答。区分名は一番上の棒の上、右端に「満足＋やや満足」の割合
{
  const s = d.base(null, "価格の「満足」は23%で、6項目の中で最も低い。", { focus: "価格の「満足」の区画と23%" });
  d.figHead(s, "項目別の満足度（2026年9月の利用者調査、回答 1,240人）");
  const rows = [
    { label: "品質", values: [48, 32, 13, 7] },
    { label: "使いやすさ", values: [45, 31, 15, 9] },
    { label: "サポートの対応", values: [42, 30, 17, 11] },
    { label: "納期", values: [36, 33, 19, 12] },
    { label: "品ぞろえ", values: [31, 35, 22, 12] },
    { label: "価格", values: [23, 29, 24, 24] },
  ];
  const r = d.hStackBars(s, {
    series: [{ name: "満足", color: G1, main: true }, { name: "やや満足", color: G2 }, { name: "どちらでもない", color: G3 }, { name: "不満", color: G4 }],
    rows, pctFmt: v => (Math.round(v * 100) === 23 ? "" : Math.round(v * 100) + "%"),
    // 右端の見出しは「不満」の区分名（細い区画の上）と重ならない長さに
    right: { head: "やや満足以上", value: row => (row.values[0] + row.values[1]) + "%" },
  });
  const i = rows.length - 1, y = r.rowY(i), w = rows[i].values[0] / 100 * (r.x1 - r.x0);
  focusSeg(s, r.x0, y, w, r.bh);
  d.focusValue(s, "23%", { x: r.x0, y: y + r.bh / 2 - 0.25, w, align: "center", fontSize: 24 }, { onFill: true }); // 棒の中の主役の値は、棒の太さに入る大きさにする
}

// マリメッコ：横幅＝市場の大きさ、高さ＝区分の中のシェア。プレーヤー名は右端の列の横
{
  const s = d.base(null, "最大の区分の法人・大口で、自社はシェア38%の首位にある。", { focus: "法人・大口の自社の区画と38%" });
  d.figHead(s, "区分別の市場の大きさ（億円）と区分の中のシェア（2025年度）。列の上は市場全体に占める割合");
  const r = d.mekko(s, {
    share: true,
    cols: [
      { name: "法人・大口", size: 420, values: [38, 27, 20, 15] },
      { name: "法人・中小", size: 310, values: [24, 30, 26, 20] },
      { name: "個人・店頭", size: 260, values: [14, 33, 31, 22] },
      { name: "個人・通販", size: 180, values: [11, 22, 37, 30] },
      { name: "公共", size: 90, values: [29, 41, 12, 18] },
    ],
    players: [{ name: "自社", color: G1, main: true }, { name: "A社", color: G2 }, { name: "B社", color: G3 }, { name: "その他", color: G4 }],
    sizeFmt: v => `${v}億円`, pctFmt: v => (Math.round(v * 100) === 38 ? "" : Math.round(v * 100) + "%"),
  });
  const c = r.col(0), [t, b] = c.seg[0];
  focusSeg(s, c.x, t, c.w, b - t);
  d.focusValue(s, "38%", { x: c.x, y: (t + b) / 2 - 0.25, w: c.w, align: "center" }, { onFill: true });
}

// ツリーマップ：区分が6つ以上の構成。大きい順に敷き詰め、名前・値・割合を区画の中に。名前が入らない小さい区分は「その他」にまとめる
{
  const s = d.base(null, "販管費84億円のうち、人件費が35%で最大の費目だ。", { focus: "人件費の区画と35%" });
  d.figHead(s, "費目別の販管費（2025年度、億円）");
  const r = d.treemap(s, {
    fmt: v => `${v.toFixed(1)}億円`, pctFmt: v => (Math.round(v * 100) === 35 ? "" : Math.round(v * 100) + "%"),
    items: [
      { name: "人件費", value: 29.4, strong: true }, { name: "外注費", value: 12.6 }, { name: "広告宣伝費", value: 9.2 },
      { name: "物流費", value: 8.4 }, { name: "賃借料", value: 7.1 }, { name: "システム費", value: 6.3 },
      { name: "販売手数料", value: 4.6 }, { name: "その他", value: 6.4 },
    ],
  });
  const c = r.cell("人件費");
  // 区画の中の「名前・金額」（18pt・16pt の2行）の下に、割合を大きく
  d.focusValue(s, "35%", { x: c.x + 0.1, y: c.y + c.h - 0.7, w: c.w - 0.2 }, { onFill: true }); // 区画の下に（上は名前と金額。文字の大きさは区画に合わせて変わる）
}

// ===== 2. 比較 =====
d.setChapter(2);

// 縦棒（単系列）：図形の棒（colBars）。値は棒の外の上、縦軸と目盛り線は置かない。見る所の1本だけ強調の色、その値は棒の上に 28pt（focus）
// ネイティブの棒（barChart）は1本だけ値を大きくできないので、見る所のあるページでは使わない（期が多く、PowerPoint 上で値を差し替える資料だけ）
{
  const s = d.base(null, "受注は3月が1,840件で最も多く、4月の1.6倍ある。", { focus: "3月の棒と1,840" });
  d.figHead(s, "月別の受注件数（2025年度、件）");
  d.colBars(s, {
    cats: MONTHS, series: [{ name: "受注", color: G2, values: [1150, 1080, 1210, 1260, 1190, 1320, 1380, 1290, 1450, 1360, 1520, 1840] }],
    focus: [[0, 11]], bottom: COL_BOTTOM, fmt: v => d.fmtNum(v),
  });
}

// 対の棒：2系列（計画と実績）。比べる相手（実績）は濃い灰、残り（計画）は薄い灰、見る所の1本だけ強調の色と 28pt の値（colBars の focus）
{
  const s = d.base(null, "計画を上回ったのは関東だけで、実績は42.2億円だった。", { focus: "関東の実績の棒と42.2" });
  // 凡例は見出しの行に置かない。系列の名前は、部品が棒のそば（最後の区分の右）に書く
  d.figHead(s, "地域別の売上の計画と実績（2026年度上期、億円）");
  d.colBars(s, {
    cats: ["関東", "近畿", "中部", "九州", "北海道・東北"],
    series: [{ name: "計画", color: G2, values: [38.0, 24.5, 18.0, 11.5, 9.0] }, { name: "実績", color: G1, values: [42.2, 21.4, 16.9, 10.8, 8.1] }],
    focus: [[1, 0]], bottom: COL_BOTTOM, fmt: v => v.toFixed(1),
  });
}

// 積み上げ縦棒：量と内訳。合計は棒の上、区分の値は区分の中、区分名は右端の棒の横。見る所の1区画だけ強調の色
{
  const s = d.base(null, "法人向けは5年で2.4倍の121億円になり、売上の伸びの大半を占める。", { focus: "2025年度の法人向けの区画と121" });
  d.figHead(s, "事業別の売上（億円）");
  const r = d.stackBars(s, {
    fmt: v => (v === 121 ? "" : d.fmtNum(v)),
    cats: ["2021年度", "2022年度", "2023年度", "2024年度", "2025年度"],
    series: [
      { name: "個人向け", color: G3, values: [82, 84, 86, 88, 90] },
      { name: "保守・修理", color: G4, values: [16, 18, 20, 22, 25] },
      { name: "法人向け", color: G1, values: [50, 66, 84, 101, 121], main: true },
    ],
  });
  const [t, b] = r.seg[4][2];
  focusSeg(s, r.bx(4), t, r.bw, b - t);
  d.focusValue(s, "121", { x: r.bx(4), y: (t + b) / 2 - 0.25, w: r.bw, align: "center" }, { onFill: true });
}

// 横棒（順位）：項目名が長いとき、順位を読ませるとき。見る所の行だけ強調の色
{
  const s = d.base(null, "駅前店の売上は6.8億円で、10店の中で最も多い。", { focus: "駅前店の棒と6.8" });
  d.figHead(s, "店舗別の売上（2026年度上期、億円）");
  const rows = [
    { label: "駅前店", value: 6.8, focusValue: true }, { label: "本店", value: 6.1 }, { label: "港南店", value: 5.4 },
    { label: "北口店", value: 4.2 }, { label: "空港店", value: 3.9 }, { label: "大学前店", value: 3.6 }, { label: "西口店", value: 3.0 },
    { label: "郊外店", value: 2.1 }, { label: "工業団地店", value: 1.9 }, { label: "山手店", value: 1.6 },
  ];
  d.hBars(s, { fmt: v => v.toFixed(1), rows }); // 主役の行の focusValue で、6.8 を部品が 36pt で棒の先に描く
}

// 正負の棒：0の線から上下に。増えた区分と減った区分を1枚で。見る所の1本は colBars の focus（28pt の値も部品が描く）
{
  const s = d.base(null, "前年比で最も伸びたのは空調で、+12.4%だった。", { focus: "空調の棒と+12.4%" });
  d.figHead(s, "商品別の売上の前年比（2026年度上期、%）");
  d.colBars(s, {
    cats: ["空調", "冷蔵", "調理", "照明", "掃除", "洗濯"],
    series: [{ name: "前年比", color: G2, values: [12.4, 8.1, -0.9, -2.5, -4.2, -6.8] }],
    focus: [[0, 0]], bottom: COL_BOTTOM, fmt: v => d.fmtNum(v, { plus: true, digits: 1 }),
  });
}

// ブレット：KPI の実績と目標。灰の帯が段階、縦の太線が目標。届かない行（問題）を1つだけ problem にし、実績を赤の 28pt で
{
  const s = d.base(null, "新規の顧客数は820社で、目標の1,100社に届かない。", { focus: "新規の顧客数の実績 820社（赤）" });
  d.figHead(s, "上期のKPIの実績と目標。濃い帯は「不足」の域。解約は少ないほど良い"); // 行の名前を短くするため、解約の向きはここで言う
  const unitFmt = (v, r) => (r.unit === "億円" || r.unit === "%" ? v.toFixed(1) : v.toLocaleString()) + r.unit;
  const rows = [
    { label: "受注件数", value: 15050, target: 14000, bands: [11000], unit: "件" },
    { label: "売上", value: 121.3, target: 120, bands: [100], unit: "億円" },
    { label: "粗利率", value: 34.2, target: 33, bands: [28], unit: "%" },
    { label: "継続率", value: 91.5, target: 90, bands: [85], unit: "%" },
    { label: "新規の顧客数", value: 820, target: 1100, bands: [900], unit: "社", problem: true },
    { label: "解約の件数", value: 48, target: 50, bands: [60], unit: "件", inverse: true },
  ];
  const r = d.bulletChart(s, { rows, fmt: (v, row) => (row.problem && v === row.value ? "" : unitFmt(v, row)) });
  // 実績の列は x1＋0.35 から幅 1.15。行の高さ（帯の太さ）は部品と同じ式で求める
  const bh = Math.min(0.44, (r.rowY(1) - r.rowY(0)) * 0.66);
  d.focusValue(s, "820社", { x: r.x1 + 0.35 - 0.6, y: r.rowY(4) + bh / 2 - 0.25, w: 1.15 + 0.6, align: "right" }, { problem: true });
}

// ===== 3. 増減の分解 =====
d.setChapter(3);

// ブリッジ：前年から今年への増減を要因ごとに。増は上へ、減は下へ浮かせる。見る所の要因だけ focusValue（強調の色の塗りと、36pt の値を部品が描く）
{
  const s = d.base(null, "値上げの+6.8億円で、原材料の値上がりの−7.5億円をほぼ補った。", { focus: "値上げの棒と+6.8" });
  d.figHead(s, "営業利益の増減の要因（2024年度→2025年度、億円）");
  d.bridge(s, {
    fmt: v => v.toFixed(1), stepFmt: v => d.fmtNum(v, { plus: true, digits: 1 }),
    start: { label: "2024年度", value: 32.0 },
    steps: [
      { label: "販売数量", value: 4.2 }, { label: "値上げ", value: 6.8, focusValue: true }, { label: "原材料", value: -7.5 },
      { label: "物流費", value: -1.6 }, { label: "人件費", value: -2.1 }, { label: "為替", value: 1.3 },
    ],
    end: { label: "2025年度" },
  });
}

// 累計の積み上げ（waterfall）：期ごとの増分を積み、右端の累計だけ内訳で。内訳は灰の濃淡、見る所の内訳だけ強調の色
// 見る所の内訳に focusValue を付けると、部品がその区分を強調の色で塗り、右のラベルの値を 36pt で描く
{
  const s = d.base(null, "3年の投資9.0億円のうち、4割の3.6億円を工場の設備に充てた。", { focus: "累計の工場の設備の区画と3.6" });
  d.figHead(s, "年度ごとの投資額と累計（億円）");
  d.waterfall(s, {
    fmt: v => v.toFixed(1), top: 2.9, base: 6.15, labelW: 2.55,
    steps: [{ label: "2023年度", value: 2.1 }, { label: "2024年度", value: 3.4 }, { label: "2025年度", value: 3.5 }],
    total: { label: "累計", parts: [{ name: "工場の設備", value: 3.6, color: G1, focusValue: true }, { name: "システム", value: 2.9, color: G2 }, { name: "研究開発", value: 2.5, color: G3 }] },
  });
}

// ===== 4. 推移 =====
d.setChapter(4);

// 折れ線（単系列）：実績は実線、見込みは点線、目標は太い破線と真上のラベル。値は言う点だけ（labelAt）、見る所の値は 28pt で点の上
{
  const s = d.base(null, "会員は2026年度4Qに50.1万人となり、目標の50万人に届く見込み。", { focus: "26年度4Qの点と50.1" });
  d.figHead(s, "会員数（四半期末、万人）。2026年度の2Q以降は見込み");
  d.lineChart(s, {
    cats: ["24年度1Q", "2Q", "3Q", "4Q", "25年度1Q", "2Q", "3Q", "4Q", "26年度1Q", "2Q", "3Q", "4Q"],
    series: [{ name: "会員数", values: [28, 29.6, 31, 32.8, 34.1, 35.5, 37.2, 38.6, 40.0, 43.2, 46.5, 50.1], main: true, from: 8, below: true, labelAt: [0, 8] }],
    targets: [{ value: 50, label: "目標 50万人" }], min: 20, max: 56, fmt: v => v.toFixed(1),
    focusValue: [0, 11], // 主役の点（系列 0 の 11 番目）。部品が 36pt で点の上に描き、右端の名前から値を外す
  });
}

// 複数系列の折れ線：見る所の系列（自社）を強調の色の 3pt、他社は灰の 1.25pt。系列名と最後の値は線の右端に。折れ線は3本まで（決まり）
{
  const s = d.base(null, "自社のシェアは2025年に26%となり、3年続けて伸びた。", { focus: "自社の線と26%" });
  d.figHead(s, "国内の出荷台数のシェア（%）");
  d.lineChart(s, {
    cats: ["2019年", "2020年", "2021年", "2022年", "2023年", "2024年", "2025年"],
    series: [
      { name: "A社", values: [31, 32, 31, 30, 28, 27, 25] },
      { name: "B社", values: [24, 23, 24, 23, 21, 20, 19] },
      { name: "自社", values: [17, 17, 18, 18, 21, 24, 26], main: true, labelAt: [0, 3] },
    ],
    min: 10, max: 36, fmt: v => `${v}%`, focusValue: [2, 6],
  });
}

// 積み上げ面：点の多い時系列の総量と内訳。区分名と最後の値は右端の面の横。見る所の区分（帯）だけ強調の色
{
  const s = d.base(null, "通販は2年で月300件から920件に伸び、受注の伸びの大半を占める。", { focus: "通販の面と920件" });
  d.figHead(s, "チャネル別の受注件数（件／月）");
  const n = 24, lerp = (a, b) => Array.from({ length: n }, (_, i) => (i === 0 || i === n - 1 ? (i ? b : a) : Math.round(a + (b - a) * i / (n - 1) + ((i * 37) % 7 - 3) * 6)));
  const r = d.areaChart(s, {
    cats: Array.from({ length: n }, (_, i) => (i % 12 === 0 ? `${2024 + Math.floor(i / 12)}年4月` : `${((i + 3) % 12) + 1}月`)), catEvery: 3,
    fmt: v => (v === 920 ? "" : d.fmtNum(v)),
    series: [
      { name: "店頭", color: G3, values: lerp(560, 640) },
      { name: "訪問・電話", color: G4, values: lerp(240, 290) },
      { name: "通販", color: C.emph, values: lerp(300, 920), main: true },
    ],
  });
  const ym = (r.yOf(r.cum[n - 1][2][0]) + r.yOf(r.cum[n - 1][2][1])) / 2;
  d.focusValue(s, "920件", { x: r.xOf(n - 1) + 0.15 + d.textW("通販　", 20) + 0.08, y: ym - 0.25, w: 1.6 });
}

// 帯でつなぐ積み上げ（ribbonBars）：構成の推移。棒の上は見る所の区分の割合。見る所の1区画（最新の年の大口）だけ強調の色
{
  const s = d.base(null, "大口の比率は5年で14%から29%に上がった。", { focus: "2025年度の大口の区画と29%" });
  d.figHead(s, "規模別の取引先の数（社）。棒の上は大口の比率");
  const MIX = {
    cats: ["2021年度", "2022年度", "2023年度", "2024年度", "2025年度"],
    series: [
      { name: "個人事業", color: G3, values: [820, 800, 770, 740, 700] },
      { name: "中小", color: G2, values: [814, 860, 900, 930, 933] },
      { name: "大口", color: G1, values: [266, 360, 480, 580, 667], main: true },
    ],
  };
  const share = i => Math.round(MIX.series[2].values[i] / MIX.series.reduce((a, sr) => a + sr.values[i], 0) * 100) + "%";
  const r = d.ribbonBars(s, { ...MIX, topLabel: i => (i === 4 ? "" : share(i)) });
  const [t, b] = r.seg[4][2];
  focusSeg(s, r.bx(4), t, r.bw, b - t);
  d.focusValue(s, share(4), { x: r.bx(4) - 0.5, y: t - 0.54, w: r.bw + 1.0, align: "center" });
}

// スロープ：2時点の変化を1本の線で。左右に名前と値。見る所の項目だけ強調の色の太線
{
  const s = d.base(null, "保守だけが、利益率を9.0%から14.0%へ5ポイント上げた。", { focus: "保守の線と14.0%" });
  d.figHead(s, "事業部別の営業利益率（%）");
  d.slopeChart(s, {
    left: "2023年度", right: "2025年度", fmt: v => `${v.toFixed(1)}%`,
    rows: [
      { name: "空調", a: 12.1, b: 12.8 }, { name: "冷蔵", a: 10.4, b: 11.2 }, { name: "保守", a: 9.0, b: 14.0, focusValue: true },
      { name: "照明", a: 7.6, b: 6.2 }, { name: "調理", a: 6.8, b: 7.9 }, { name: "掃除", a: 5.2, b: 4.1 }, { name: "洗濯", a: 4.0, b: 3.1 },
    ],
  });
  // 主役の行の focusValue で、右端の 14.0% を部品が 36pt で描く（ほかの行の名前は、その高さの分よける）
}

// 棒と折れ線の複合：2つの縦軸を重ねず、上の段に率、下の段に量。見る所は棒の focus だけ（線は soft の色）
{
  const s = d.base(null, "商談は月620件まで増えたが、成約率は24%から17%に下がった。", { focus: "25年度4Qの棒と620" });
  d.figHead(s, "四半期ごとの商談の件数（件／月）と成約率（%）");
  const r = d.comboChart(s, {
    cats: ["24年度1Q", "2Q", "3Q", "4Q", "25年度1Q", "2Q", "3Q", "4Q"],
    line: { name: "成約率", values: [24, 23.5, 22, 21.4, 20, 19.2, 18.1, 17], fmt: v => `${v}%` },
    bars: { name: "商談の件数", values: [410, 430, 455, 480, 520, 560, 590, 620], focus: [7], fmt: v => (v === 620 ? "" : d.fmtNum(v)) },
  });
  d.focusValue(s, "620", { x: r.cx(7) - 0.8, y: r.barTop(7) - 0.54, w: 1.6, align: "center" });
}

// 面積棒（areaBars）：横幅＝期間、高さ＝月額、面積＝総額。費目は灰の濃淡。見る所の段階を強調の色の枠で囲み、総額を 28pt
// 審査の線（gates）は部品が赤で描く（赤は問題だけ）ので、この見本では使わない
{
  const s = d.base(null, "総額は約1.3億円で、4割強を全社への展開で使う。", { focus: "全社への展開の面と5,600万円" });
  d.figHead(s, "段階ごとの費用（万円）");
  const r = d.areaBars(s, {
    fmt: v => (v === 5600 ? "" : `${v.toLocaleString()}万円`), unitFmt: v => `月${v.toLocaleString()}万円`, howTo: "高さ：月額\n横幅：期間",
    cats: [["人件費", G1], ["外部委託", G2], ["広告", G3], ["その他", G4]].map(([k, c]) => ({ key: k, name: k, color: c })),
    stages: [
      { name: "聞き取り", months: 2, parts: { 人件費: 400, その他: 50 } },
      { name: "仕組みの整備", months: 3, parts: { 人件費: 600, 外部委託: 300, その他: 50 } },
      { name: "部署での試行", months: 3, parts: { 人件費: 700, 外部委託: 300, その他: 100 } },
      { name: "全社への展開", months: 4, parts: { 人件費: 800, 外部委託: 300, 広告: 200, その他: 100 } },
    ],
    rows: [{ label: "人数", h: 0.4, cells: ["1名", "2名", "2名", "3名"] }],
  });
  d.rect(s, r.sx(3) + 0.03, r.yT, r.sw(3) - 0.06, r.yB - r.yT, null, { color: C.emph, width: 2.25 });
  d.focusValue(s, "5,600万円", { x: r.sx(3), y: r.yT - 0.82, w: r.sw(3), align: "center" }); // 下の「月1,400万円」に重ならない高さ
}

// ===== 5. 分布・関係 =====
d.setChapter(5);

// ヒストグラム：区間ごとの件数（隙間の小さい棒）。見る所の区間（1時間超え）だけ強調の色、割合を区間の上に 28pt。中央値は灰の基準の線
{
  const s = d.base(null, "対応時間が1時間を超える案件が、全体の19%ある。", { focus: "60分以上の棒と19%" });
  d.figHead(s, "1件あたりの対応時間の分布（2026年9月の 2,860件、件）");
  const bins = [96, 312, 548, 604, 452, 310, 188, 132, 92, 62, 40, 24], n = bins.length, x0 = M + 0.4, x1 = W - M - 0.4, slot = (x1 - x0) / n;
  const r = d.colBars(s, {
    // 区間は上端だけ書く（「100〜110」は 16pt で隣の区間名と重なる）
    cats: bins.map((_, i) => `〜${i * 10 + 10}`), series: [{ name: "件数", color: G2, values: bins }],
    focus: bins.map((_, i) => [0, i]).filter(([, i]) => i >= 6), focusFs: FS.body,
    x0, x1, bw: slot - 0.05, top: 2.85, bottom: COL_BOTTOM, max: 700, fmt: v => d.fmtNum(v),
  });
  const xAt = m => x0 + m / 10 * slot;
  d.refLine(s, { x: xAt(38), y0: 2.85, y1: r.base, label: "中央値 38分", kind: "avg" });
  d.refLine(s, { x: xAt(60), y0: 2.85, y1: r.base, label: "1時間", kind: "target" });
  // 見る所は6本の棒のまとまり。その割合を、まとまりの真上（先頭の棒の値の上）に
  d.focusValue(s, "19%", { x: xAt(60), y: r.barTop(0, 6) - 0.95, w: xAt(120) - xAt(60), align: "center" });
  d.T(s, "分", { x: x1 + 0.05, y: r.base + 0.3, w: 0.4, h: 0.3, fontSize: FS.body, color: C.soft });
}

// 散布図：点に名前を直接。平均の破線で4つに分け、象限の名前を角に。見る所の点だけ強調の色、値は点の左に 28pt
{
  const s = d.base(null, "駅前店の売上は6.8億円で、同じ広さの南店の2.8倍ある。", { focus: "駅前店の点と6.8億円" });
  d.figHead(s, "店舗ごとの売場の面積と売上（2026年度上期、14店）"); // 16pt の名前が重ならないよう、20店から14店に絞った
  const P = [
    ["駅前店", 620, 6.8, true], ["本店", 980, 6.1], ["港南店", 840, 5.4], ["北口店", 560, 4.2], ["空港店", 700, 3.9], ["大学前店", 480, 3.3],
    ["西口店", 650, 3.0], ["郊外店", 1100, 2.1], ["工業団地店", 900, 1.9], ["山手店", 420, 1.6], ["中央店", 760, 4.6], ["東店", 530, 2.8],
    ["南店", 610, 2.4], ["丘の上店", 1020, 3.1],
  ].map(([name, x, y, strong]) => ({ name, x, y, strong }));
  const r = d.scatter(s, {
    area: SC_AREA, points: P, xName: "売場の面積（㎡）", yName: "売上（億円）", xMin: 300, xMax: 1300, yMax: 8, xFmt: v => d.fmtNum(v), yFmt: v => v.toFixed(0),
    refX: { value: 720, label: "平均 720㎡" }, refY: { value: 3.7, label: "平均 3.7億円" }, quads: ["広くて売れる", "狭くて売れる", "狭くて売れない", "広いが売れない"],
  });
  const [px, py] = r.pt(0);
  d.focusValue(s, "6.8億円", { x: px - 0.18 - 2.0, y: py - 0.25, w: 2.0, align: "right" });
}

// バブル：横＝市場の成長率、縦＝自社のシェア、大きさ＝市場の大きさ。見る所の円だけ強調の色、値は円の上に 28pt
{
  const s = d.base(null, "最も伸びる海外市場で、自社のシェアは4%にとどまる。", { focus: "海外の円と4%" });
  d.figHead(s, "市場ごとの成長率と自社のシェア（2025年度）。円の大きさは市場の大きさ");
  const pts = [
    { name: "法人・大口　420億円", x: 0.5, y: 41, size: 420 }, { name: "法人・中小　310億円", x: 4.8, y: 24, size: 310 },
    { name: "個人・店頭　260億円", x: -2.1, y: 14, size: 260 }, { name: "個人・通販　180億円", x: 12.4, y: 9, size: 180 },
    { name: "公共　90億円", x: 1.5, y: 29, size: 90 }, { name: "海外　150億円", x: 14.1, y: 4, size: 150, strong: true }, { name: "保守　120億円", x: 6.3, y: 33, size: 120 },
  ];
  const maxD = 1.15, r = d.scatter(s, {
    // 右端の海外の円の右に名前が入るよう、横軸を 22% まで
    xName: "市場の成長率（年率、%）", yName: "自社のシェア（%）", xMin: -4, xMax: 22, area: SC_AREA, yMax: 50, xFmt: v => d.fmtNum(v, { unit: "%" }), yFmt: v => `${v}%`, maxD,
    refX: { value: 5, label: "全体の成長率 5%" }, points: pts,
  });
  const [px, py] = r.pt(5), dd = maxD * Math.sqrt(150 / 420);
  // 名前は円の右に置かれるので、値は円の真上に
  d.focusValue(s, "4%", { x: px - 0.6, y: py - dd / 2 - 0.52, w: 1.2, align: "center" });
}

// ヒートマップ：行×列の濃淡。値はセルの中、セルの間は白い溝。見る所のセルを強調の色の枠で囲み、値をそのセルの中に 28pt（濃い塗りの上なので白）
{
  const s = d.base(null, "問い合わせは月曜の10時台に集中し、1時間に140件届く。", { focus: "月曜10時のセルと140" });
  d.figHead(s, "曜日×時間帯ごとの問い合わせ（2026年7〜9月の平均、件／時）");
  const base = [40, 62, 58, 44, 50, 55, 48, 42, 36, 30, 22, 14], fac = [2.3, 1.4, 1.2, 1.15, 1.25, 0.7, 0.55], wk = [0, 0, 0, 0, 0, 1, 1];
  const vals = fac.map((f, r) => base.map((b, c) => Math.round(b * f * (wk[r] && c > 3 && c < 9 ? 1.6 : 1) * (r === 0 && c === 1 ? 0.98 : 1))));
  // 列は営業時間の9〜17時の9列に絞る（12列だとセルが 36pt の値より狭い）
  const V = vals.map(row => row.slice(0, 9));
  d.heatmap(s, {
    rows: ["月曜", "火曜", "水曜", "木曜", "金曜", "土曜", "日曜"],
    cols: ["9時", "10時", "11時", "12時", "13時", "14時", "15時", "16時", "17時"],
    values: V, fmt: v => d.fmtNum(v), focusValue: [0, 1], mark: [[0, 1]], markColor: C.emph, total: { head: "9〜17時の計", value: i => d.fmtNum(V[i].reduce((a, b) => a + b, 0)) },
  });
}

// レーダー：5〜8観点の形。外周（満点）と軸の線だけ。自社は濃い灰、比べる相手は薄い灰、見る所の1点（問題）だけ赤の点と 28pt
{
  const s = d.base(null, "価格だけが3.1点で、A社の3.9点を下回る。", { focus: "価格の自社の点と3.1（赤）" });
  const fh = "観点別の評価（5点満点、2026年9月の調査）"; // 凡例は文と分けて置く
  d.figHead(s, fh); d.figHead(s, "", [["自社", C.soft], ["A社", G2]], { x: W - M - 2.2 }); // 凡例は右端に（中央に置くと上の観点「品質」に重なる）
  const axes = ["品質", "機能", "デザイン", "価格", "サポート", "納期"], cy = 4.58, rr = 1.5;
  const r = d.radarChart(s, {
    axes, max: 5, fmt: v => (v === 3.1 ? "" : v.toFixed(1)), cy, r: rr,
    series: [{ name: "自社", values: [4.4, 4.2, 4.0, 3.1, 4.3, 4.1], main: true, color: C.soft }, { name: "A社", values: [3.9, 3.8, 3.6, 3.9, 3.4, 3.7], color: G2 }],
  });
  const [px, py] = r.pt(3, 3.1);
  dot(s, px, py, 0.2, C.accent);
  // 価格の観点の名前は下。2行目は「（自社の値）／A社の値」なので、自社の値を「／3.9」の左に大きく書く
  const y2 = cy + rr + 0.16 + 0.43;
  d.focusValue(s, "3.1", { x: W / 2 - 0.4 - 1.2, y: y2 - 0.25, w: 1.2, align: "right" }, { problem: true });
}

// 人口ピラミッド：2群を左右に。項目名は中央、値は棒の外。見る所の1本（30代の女性）だけ強調の色
{
  const s = d.base(null, "最も多いのは30代の女性で、6.0万人いる。", { focus: "30代の女性の棒と6.0" });
  d.figHead(s, "年代別・性別の会員数（2026年9月末、万人）");
  const rows = [
    { label: "70代以上", left: 6, right: 9 }, { label: "60代", left: 12, right: 17 }, { label: "50代", left: 22, right: 30 },
    { label: "40代", left: 44, right: 58 }, { label: "30代", left: 46, right: 60 }, { label: "20代", left: 30, right: 38 }, { label: "10代", left: 12, right: 16 },
  ];
  const r = d.butterfly(s, { names: ["男性", "女性"], fmt: v => (v === 60 ? "" : (v / 10).toFixed(1)), rows });
  const x = r.cx + r.mw / 2, w = 60 * r.k, y = r.rowY(4);
  focusSeg(s, x, y, w, r.bh);
  d.focusValue(s, "6.0", { x: x + w + 0.08, y: y + r.bh / 2 - 0.25, w: 1.0 });
}

// ファネル：段階ごとの件数と歩留まり。落ちる所（問題）の率を赤の 28pt で
// funnel の drop（赤枠）は率を 15pt で描き、率の書式も変えられないので、drop を使わず、その率を白で覆って focusValue を置く
{
  const s = d.base(null, "試用から見積りへ進むのは25%で、最も落ちる。", { focus: "試用→見積りの率 25%（赤）" });
  const r = d.funnel(s, {
    stages: [{ name: "問い合わせ", value: 420 }, { name: "商談", value: 231 }, { name: "試用", value: 120 }, { name: "見積り", value: 30 }, { name: "契約", value: 18 }],
  });
  const [ccx, cy] = r.conv(2);
  d.rect(s, ccx - 0.62, cy - 0.2, 1.24, 0.4, C.white);
  d.focusValue(s, "25%", { x: ccx - 0.62, y: cy - 0.25, w: 1.24, align: "center" }, { problem: true });
}

// ===== 6. 市場規模 =====
d.setChapter(6);

// 市場の絞り込み：入れ子の円（大きさは金額に比例しない）と、金額・対象・根拠の3行。狙う1行だけ strong（部品が金額を 36pt・強調の色で描く）
// 見本の値なので（仮）は付けない。実在の資料で確かめていない値は、金額の文字に（仮）を付ける
{
  const s = d.base(null, "業務記録ツールの国内市場1,650億円のうち、3年目に約27億円を狙う。", { focus: "SOM の行と約27億円" });
  d.tamSamSom(s, {
    // 円の中の一言は見出しの大きさ。内側の円ほど入る字数が減る（内側は全角4字まで）。絞り込みの条件（国内・クラウド・3年目）は対象の列で言う
    // 列の幅：金額の列は主役の「約27億円」が入る幅、残りを対象と根拠に。文は意味の切れ目で改行し、短い行を残さない
    colW: [2.35, 1.9, 2.4],
    items: [
      { name: "TAM", what: "業務記録ツール", value: "約1,650億円", scope: "国内企業の\n支払総額\n（2025年）", basis: "架空の業界推計\nを業種で集計" },
      { name: "SAM", what: "中小企業向け", value: "約480億円", scope: "300人未満で\nクラウド利用", basis: "TAM\n×クラウド40%\n×300人未満73%" },
      { name: "SOM", what: "有料顧客", value: "約27億円", scope: "3年目の\n有料1,500社", basis: "1,500社\n×年180万円", strong: true },
    ],
  });
}

// 市場の成長（年平均成長率）：実績の市場規模を、0 から始まる縦棒で年ごとに並べ、最初と最後の棒を矢印でつないで年平均成長率を書く（colBars の cagr）
// 率は部品が値から計算する（手で書かない）。対数軸・推計の幅の帯・出どころの違う推計の線は重ねない。見る所は年平均の値
{
  const s = d.base(null, "国内市場は6年で約4.3倍の41億円に伸び、年平均+27.6%で成長した。", { focus: "年平均の成長率（+27.6%）" });
  d.figHead(s, "国内の業務記録ツールの市場規模（億円、各年の実績）");
  d.colBars(s, {
    cats: ["2019年", "2020年", "2021年", "2022年", "2023年", "2024年", "2025年"],
    series: [{ name: "実績", color: C.bar, values: [9.5, 12, 15.5, 19, 25, 32, 41] }],
    fmt: v => d.fmtNum(v, { digits: 1 }), bottom: COL_BOTTOM, cagr: { from: 0, to: 6 },
  });
}

d.finalize(OUT);
