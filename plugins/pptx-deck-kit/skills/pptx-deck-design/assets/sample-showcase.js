// sample-showcase.js : 高密度の見本（多段の樹形図・指標ごとの棒・評価の丸の表・段階の道・プロジェクトのガント。内容はすべて架空）
// 実行: NODE_PATH=<pptxgenjs と jszip の入った node_modules> node sample-showcase.js <出力先.pptx>
//
// 論旨：窓口への問い合わせが減らない根本は、回答データの規程が古いこと。4案のうち問い合わせAIが件数を最も減らし、改定にも即日で追従する。
//       正答率90%を検証してから、6月までに全社へ広げる。
//
// 色の組：blue（強調 2F6BFF）。社内のIT施策（問い合わせAI）なので、IT の色にした。問題（手を打つべき根本の原因）は赤（accent）。
//
// 形の割り当て（1つの形は1つの意味にだけ使う）
//   kpiTree      … 件数を掛け算・足し算で分解し、どこが増えたか（P.3）
//   logicTree    … 原因の分解（P.4）
//   metricBars   … 単位の違う数値で案を比べる（P.4）
//   compareTable … 満たすかどうかで案を比べる（P.5。評価の丸だけ。セルに文字を入れない）
//   roadmap      … 段階と、次へ進む条件（P.6）
//   projectGantt … 作業の日程（P.7）
//
// 見本の約束：密度は、図の中の数値と注釈で作る（ノードの数値、棒の値、駅の条件の値、作業と担当）。マス目に文を積まない。
//   題はアジェンダの項目名だけ（createDeck の nav から取る）。主張はリード。全ページで見る所（focus）を1つ宣言し、強調はそこだけ。補足の sub は使わない
const path = require("path");
const pptxgen = require("pptxgenjs");
const { createDeck } = require("./deck-parts");

const OUT = path.resolve(process.argv[2] || path.join(__dirname, "sample-showcase.pptx"));
const pres = new pptxgen();
const d = createDeck(pres, {
  audience: "insider", // 聞き手の前提（決まり24）。見本は部品の使い方を見せるので、前提を共有済みとして描く
  title: "社内問い合わせAIの全社展開計画",
  theme: "blue",
  footerText: "社内問い合わせAIの全社展開計画",
  nav: [{ name: "問い合わせが減らない原因", ch: 1 }, { name: "対策案の比較", ch: 2 }, { name: "進め方", ch: 3 }, { name: "全社展開の計画", ch: 4 }],
});
const { M, T, FS, C } = d;

d.cover({ title: "社内問い合わせAIの全社展開計画", date: "2027年1月" });
d.agendaPage();

// ===== 1. 減らない原因（1）：KPIツリー（件数を掛け算・足し算で分解し、各ノードに値と前年比） =====
// 親＝子の和（＋）か積（×）になるよう、数値をつなげる。増えた所1つだけ赤枠で、前年比を主役の値にする
d.setChapter(1);
{
  const s = d.base(null, "窓口の件数が減らないのは、窓口への聞き直しが月260件に増えたことが原因。", { ask: "全体は何件？ → 図の月1,100件", focus: "赤枠の「窓口への聞き直し」の +18%" });
  d.kpiTree(s, {
    root: { name: "窓口の件数（月）", value: "1,100件", chg: "−1.8%", op: "＋", children: [
      { name: "初めての問い合わせ", value: "840件", chg: "−6.7%", op: "×", children: [
        { name: "問い合わせた社員", value: "600人", chg: "−4.8%" },
        { name: "1人あたりの件数", value: "1.4件", chg: "−2.0%" },
      ] },
      { name: "窓口への聞き直し", value: "260件", chg: "+18%", op: "×", problem: true, children: [
        { name: "AIへの問い合わせ", value: "1,440件", chg: "+20.0%" },
        { name: "聞き直す割合", value: "18%", chg: "−0.3pt" },
      ] },
    ] },
  });
}

// ===== 1. 減らない原因（2）：多段の樹形図（問題 → 直接の原因 → 根本の原因） =====
// 1つのノードから3本ずつ広がる（1 → 3 → 9）。ノードの名前に数値を書き込む（下に補足の小さな文字を足さない）。手を打つべき根本の原因1つだけ赤枠
{
  const s = d.base(null, "手を打つべき根本の原因は、回答データの規程が古いこと。", { ask: "どれくらい古い？ → 図の旧版23件", focus: "赤枠の根本の原因（回答データの規程が古い）" });
  d.logicTree(s, {
    heads: ["問題", "直接の原因", "根本の原因"],
    root: {
      text: "窓口への問い合わせが\n月1,100件から減らない",
      children: [
        { text: "AIの回答が間違う\n誤答率 18%", children: [
          { text: "回答データの規程が古い（旧版23件）", problem: true },
          { text: "部署ごとの例外が未登録（40件）" },
          { text: "社内用語を読み違える（辞書なし）" },
        ] },
        { text: "社員がAIを使わない\n利用した社員 31%", children: [
          { text: "AIの存在を知らない（案内は1か所）" },
          { text: "電話が早いと思う（入口まで3回の操作）" },
          { text: "前の誤答で信用しない（訂正が不明）" },
        ] },
        { text: "窓口へ聞き直す\n月260件", children: [
          { text: "回答に根拠の規程が出ない" },
          { text: "答えられない質問を窓口へ渡せない" },
          { text: "窓口が聞き直しを断らない" },
        ] },
      ],
    },
  });
}

// ===== 2. 4案の比較（1）：測れる軸は、指標ごとの棒 =====
// 行＝案、列＝指標。単位や桁の違う数値を表に並べず、指標ごとに尺度をそろえた棒で長さを見せる。決め手の値1つを大きく（key）
// 推奨案の弱点（導入の期間が最も長い）も、他の値と同じ濃さで見せる
const OPTS = ["現状", "FAQの整備", "問い合わせAI", "窓口の外部委託"];
const OPT_ICON = ["pic-phone", "file", "bot", "pic-building"];
d.setChapter(2);
{
  const s = d.base(null, "窓口の件数が最も減る案は問い合わせAIで、月660件に減少。", { ask: "今は何件で、費用は？ → 図の現状1,100と年間の費用", focus: "問い合わせAIの窓口の件数（660件）" });
  const IX = 0.55;
  const r = d.metricBars(s, {
    x: M + IX, w: d.BW - IX, items: OPTS, focus: 2, key: [2, 0],
    metrics: [
      { label: "窓口の件数", unit: "件／月", values: [1100, 990, 660, 1100] },
      { label: "年間の費用", unit: "万円", values: [0, 80, 600, 1800] },
      { label: "導入の期間", unit: "か月", values: [0, 2, 6, 3] },
    ],
  });
  OPT_ICON.forEach((nm, i) => { const rw = r.row(i); d.icon(s, nm, M + 0.22, rw.y + rw.h / 2, 0.36); });
}

// ===== 2. 4案の比較（2）：語で比べる軸は、満たすかどうかに言い直して評価の丸の表 =====
// 同じ章の続きなので、題は変えない（何のページかはリードで言う）。セルは評価の丸だけ。決め手のセル1つを focus。結論は表の横（note）
{
  const s = d.base(null, "規程の改定を即日で反映できるのは、問い合わせAIだけ。", { ask: "どうやって即日にする？ → 図の右の「差し替えで即日」", focus: "問い合わせAIの「改定を即日で反映」のセル" });
  d.compareTable(s, {
    legend: true, colGap: 0.2, // 行は部品が本文の下端まで広げる（行を足して埋めない）
    note: { head: "差し替えで即日" }, // 結論は短い見出しだけ（補足は置かない）
    // 列の幅は見出しが1行に入る幅だけにし、行の名前を短くして（全角8字まで）、右に結論（note）の幅を残す
    cols: [{ head: "現状", base: true, w: 0.9 }, { head: "FAQの整備", w: 1.5 }, { head: "問い合わせAI", w: 1.8 }, { head: "窓口の外部委託", w: 1.9 }],
    rows: [
      { label: "改定を即日で反映", cells: [{ ball: 2 }, { ball: 1 }, { ball: 4, focus: true }, { ball: 1 }] },
      { label: "正答率90%以上", cells: [{ ball: 3 }, { ball: 3 }, { ball: 3 }, { ball: 1 }] },
      { label: "夜間・休日に回答", cells: [{ ball: 0 }, { ball: 4 }, { ball: 4 }, { ball: 2 }] },
      { label: "部署の例外に回答", cells: [{ ball: 4 }, { ball: 1 }, { ball: 2 }, { ball: 2 }] },
    ],
  });
}

// ===== 3. 進め方：段階の道（段階×観点の表 journeyMap の代わり） =====
// 段階ごとに置くのは、名前・時期・絵・次へ進む条件の値（3つまで。箱は値と何の値かを合わせて全角7字ほど）だけ。段階は5つまでに絞る（6つだと駅の幅が狭く、条件の文字が箱からはみ出す）。下の帯は費用。見る所の段階（審査）を1つだけ strong
d.setChapter(3);
{
  const s = d.base(null, "正答率90%の検証後に、問い合わせAIを営業本部と全社へ展開。", { ask: "いつ広げる？ → 図の4〜5月と5〜6月", focus: "審査の駅（正答率90%以上）" });
  d.roadmap(s, {
    stages: [
      { name: "準備", when: "2027年1〜2月", icon: "pic-book", kpis: [["上位10件", "対象"], ["120件", "修正"], ["2人", "規程の担当"]] },
      { name: "試行", when: "2〜3月", icon: "pic-laptop", kpis: [["2部署", "人事経理"], ["60%以上", "利用"], ["全件", "聞き直し"]] },
      { name: "審査", when: "4月", icon: "pic-check", strong: true, kpis: [["90%以上", "正答", { key: true }], ["0件", "重大な誤答"]] }, // 条件は2つに絞る（聞き直しの上限は箱に入らない）
      { name: "営業本部へ展開", when: "4〜5月", icon: "pic-building", kpis: [["800人", "社員"], ["月880件", "窓口"], ["当日", "訂正"]] },
      { name: "全社へ展開", when: "5〜6月", icon: "pic-people", kpis: [["3,000人", "社員"], ["月660件", "窓口"], ["3か所", "案内"]] },
    ],
    bands: [{ from: 0, to: 1, label: "準備と試行　420万円" }, { from: 2, to: 4, label: "展開　600万円" }],
  });
}

// ===== 4. 展開の計画：プロジェクトのガント（計画だけ。グループと作業の2段、日単位） =====
// 帯は同じ太さ。リードが言う作業の帯だけ強調の色（task.focus）、他は灰。遅れや進み具合は、別のページで同じ図を背面に敷いて前面で言う
// 4グループなら作業は8つまで（それ以上は行が文字より低くなる）。各グループで、リードと次の段階につながる作業だけ残す
// 見る所の帯のすぐ右に、条件の数値を注釈で書く（図の外に説明欄を作らない）
d.setChapter(4);
{
  const s = d.base(null, "正答率の測定は3月の4週間で、4月の審査に間に合う日程。", { ask: "合格の線は？ → 図の正答率90%以上", focus: "「正答率の測定」の帯（3月の4週間）" });
  const r = d.projectGantt(s, {
    from: "2027.01", to: "2027.06", labelW: 2.6, // 作業名（全角9字）が1行に入る幅
    groups: [
      { name: "1. 準備", owner: "情シス", tasks: [
        { name: "回答データの棚卸し", owner: "情シス", bar: ["2027.01.13", "2027.02.14"] },
        { name: "差し替え担当の決定", owner: "総務", bar: ["2027.01.20", "2027.01.31"] },
      ] },
      { name: "2. 試行", owner: "人事", tasks: [
        { name: "人事・経理での試行", owner: "人事", bar: ["2027.02.17", "2027.03.21"] },
        { name: "正答率の測定", owner: "情シス", bar: ["2027.03.01", "2027.03.28"], focus: true },
      ] },
      { name: "3. 審査", owner: "PM", tasks: [
        { name: "審査資料の作成", owner: "PM", bar: ["2027.03.29", "2027.04.11"] },
        { name: "窓口の対応の見直し", owner: "窓口", bar: ["2027.04.01", "2027.04.18"] },
      ] },
      { name: "4. 展開", owner: "PM", tasks: [
        { name: "営業本部への展開", owner: "営業", bar: ["2027.04.19", "2027.05.16"] },
        { name: "全社への展開", owner: "PM", bar: ["2027.05.24", "2027.06.27"] },
      ] },
    ],
  });
  const fr = r.rows.find(q => q.t && q.t.focus);
  T(s, "4週間　合格は正答率90%以上", { x: r.xOf("2027.03.28", true) + 0.12, y: fr.y, w: 3.4, h: fr.h, fontSize: FS.body, bold: true, color: C.emph, valign: "middle", wrap: false });
}

d.finalize(OUT);
