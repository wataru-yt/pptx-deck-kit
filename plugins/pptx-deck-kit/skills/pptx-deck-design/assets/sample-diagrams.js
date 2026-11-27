// sample-diagrams.js : 時期の点・状態のマップ・段階の道と決裁の見本（eventTimeline・stateMap・roadmap＋series。内容はすべて架空）
//   流れ図・スイムレーン・循環・体制図は sample-flow.js、2軸・ファネル・構成図は sample-map.js
// 実行: NODE_PATH=<pptxgenjs と jszip の入った node_modules> node sample-diagrams.js <出力先.pptx>
//
// 論旨：購買の電子化は同業で7社目。試行で身についた力を使い、4月に全社へ広げる。その費用480万円の承認をお願いしたい。
//
// 色の組：blue（強調 2F6BFF）。社内の業務システム（購買の仕組み）の見直しなので、IT の色にした。問題は赤（accent）。
//
// 形の割り当て（1つの形は1つの意味にだけ使う）
//   eventTimeline … 同業の会社が電子化した時期（P.3）
//   stateMap      … 項目ごとの、身についたかどうかの状態（P.4）
//   roadmap       … 段階と、次へ進む条件（P.5。P.6 の決裁で、同じ図を series で背面に敷く）
//
// 見本の約束：題はアジェンダの項目名だけ（createDeck の nav から取る）。主張はリード。全ページで見る所（focus）を1つ宣言し、強調はそこだけ。
//   補足の sub は使わない。決裁は、根拠の図を背面に敷き、前面の messageBox で1件だけ言い、決める額を focusValue で大きく書く（decisionPage は使わない）
const path = require("path");
const pptxgen = require("pptxgenjs");
const { createDeck } = require("./deck-parts");

const OUT = path.resolve(process.argv[2] || path.join(__dirname, "sample-diagrams.pptx"));
const CHECK = process.env.ECHO_CHECK === "1"; // 1 にすると続きのページの覆いと前面を外して描く（figure_diff.py で図が動いていないかを確かめる用）
const pres = new pptxgen();
const d = createDeck(pres, {
  audience: "insider", // 聞き手の前提（決まり24）。見本は部品の使い方を見せるので、前提を共有済みとして描く
  title: "購買の仕組みの見直し",
  theme: "blue",
  footerText: "購買の仕組みの見直し",
  nav: [{ name: "同業他社の動向", ch: 1 }, { name: "習得したスキル", ch: 2 }, { name: "展開計画", ch: 3 }, { name: "決裁のお願い", ch: 4 }],
});

d.cover({ title: "購買の仕組みの見直し", date: "2027年2月" });
d.agendaPage();

// コンセプト：新しい仕組みを初めて見せるページ。属性を並べず、価値を言う一文を中央に（決まり24・25）
// 伝えたいこと：承認が当日に終わる／分かればOK：紙の稟議だと3日、新しい仕組みなら当日
{
  const s = d.base("見直しのコンセプト", null, { focus: "新しい仕組みの「当日」" });
  d.concept(s, {
    statement: "稟議の承認が、当日。",
    key: "当日",
    contrast: {
      ticks: [{ value: 0, label: "申請" }, { value: 1, label: "1日" }, { value: 2, label: "2日" }, { value: 3, label: "3日" }],
      rows: [
        { name: "今の紙の稟議", value: 3 },
        { name: "新しい仕組み", value: 0, big: "当日に承認", strong: true }, // 当日＝0日後（帯は最小の幅）
      ],
    },
  });
}

// ===== 1. 同業他社の動向：時期の点の図 =====
// 名前は点の真下、近いものは段をずらして細い引出線。自社だけ濃紺・太字（見る所）。会社と制度の変更は行を分ける
d.setChapter(1);
{
  const s = d.base(null, "当社の購買の電子化は同業で7社目で、大手の3社より7年遅い。", { ask: "大手はいつ始めた？ → 図の2019年", focus: "当社の点（2026年10月）" });
  const r = d.eventTimeline(s, {
    from: "2019.01", to: "2026.12",
    levels: 2,
    rows: [
      { label: "同業の大手", events: [
        { at: "2019.04", name: "A社" }, { at: "2019.07", name: "B社" }, { at: "2019.11", name: "C社" },
      ] },
      { label: "同業の中堅", events: [
        { at: "2021.06", name: "D社" }, { at: "2021.09", name: "E社" }, { at: "2023.02", name: "F社" },
        { at: "2026.10", name: "当社", self: true },
      ] },
      { label: "制度の変更", events: [
        { at: "2020.04", name: "記録の保存の指針" }, { at: "2022.10", name: "請求書の電子化の推奨" }, { at: "2025.04", name: "保存期間の延長" },
      ] },
    ],
  });
  // リードが言う数値（7社目）を、見る所の点のすぐ隣に大きく書く（決まり16）
  const [px, py] = r.dot(1, 3);
  d.focusValue(s, "7社目", { x: px - 2.0, y: py - 0.62, w: 1.85, align: "right" });
}

// ===== 2. 習得したスキル：状態を描き分けるマップ =====
// 塗りの色を増やさず、枠・破線・文字の濃さで3つの状態を描き分ける。凡例は右上に細い枠の見本だけ
d.setChapter(2);
{
  const s = d.base(null, "企画から試行までに、6つの分野すべてで新しいスキルを身につけた。", { ask: "何を身につけた？ → 図の濃い枠の箱（14項目）", focus: "濃い枠の箱（今回身についた14項目）" });
  const N = t => ({ text: t, state: "new" }), E = t => ({ text: t, state: "tried" }), B = t => ({ text: t, state: "base" });
  d.stateMap(s, {
    itemH: 0.46, itemGap: 0.07,
    states: [{ key: "new", label: "今回身についた（14項目）", style: "frame" }, { key: "tried", label: "経験した（再現性はまだない）", style: "dashed" }, { key: "base", label: "以前から", style: "plain" }],
    // 6列だと箱の中は全角6字まで（箱は1行の高さ）。項目名は6字までの熟語にする
    groups: [
      { name: "企画", items: [N("業務の調査"), N("効果の試算"), E("予算の申請")] },
      { name: "業務の設計", items: [N("承認の流れ"), N("例外の整理"), N("手引きの作成"), E("社内への説明")] },
      { name: "システム", items: [B("要件定義"), N("製品の比較"), B("設計"), N("委託先の管理"), B("テスト"), E("認証の連携")] },
      { name: "調整", items: [N("取引先の調整"), N("経理部の合意"), N("部署の説明会")] },
      { name: "契約・法務", items: [N("委託の契約"), E("取引先の審査"), N("個人情報確認")] },
      { name: "管理", items: [N("日程の管理"), E("予算の管理"), N("試行の評価")] },
    ],
  });
}

// ===== 3. 展開計画：段階の道（P.6 の決裁で、同じ図を背面に敷く） =====
// 段階ごとに置くのは、名前・時期・絵・次へ進む条件の値だけ（マス目に文を積まない）。下の帯は費用。見る所の段階を1つだけ strong
// 条件の箱は 16pt で、値と「何の値か」を合わせて全角8字ほどまで（例「約1日　待ち」）。何の値かは熟語にする
const ROAD = {
  stages: [
    { name: "試行", when: "2026年10〜12月", icon: "pic-search", kpis: [["2部署", "人事と経理"], ["約1日", "待ち"], ["1割", "差し戻し"]] },
    { name: "全社展開", when: "2027年4月〜", icon: "pic-building", strong: true, kpis: [["約1日", "全社の待ち", { key: true }], ["12部署", "使う部署"], ["5人", "購買の担当"]] },
    { name: "定着", when: "2027年10月〜", icon: "pic-check", kpis: [["月400件", "申請"], ["9割", "入力率"], ["半年", "手引き更新"]] },
    { name: "取引先と連携", when: "2028年4月〜", icon: "pic-contract", kpis: [["3社", "電子の発注"], ["約2日", "見積り"], ["0件", "転記"]] },
  ],
  bands: [{ from: 0, to: 0, label: "試行　120万円" }, { from: 1, to: 1, label: "展開　480万円" }, { from: 2, to: 3, label: "運用　年180万円" }],
};
const road = d.series(s => d.roadmap(s, ROAD), { check: CHECK });
d.setChapter(3);
road.page({ lead: "4月に新しい仕組みを12部署へ広げ、確認の待ちを全社で約1日にする。", ask: "本当に約1日にできる？ → 図の試行の実績（約1日）", focus: "全社展開の駅（2027年4月〜）" });

// ===== 4. 決裁のお願い：同じ段階の道を背面に敷き、決めてもらう段階だけ残す =====
// 前面の messageBox で決めていただくことを1件だけ言い、決める額を focusValue で大きく書く。2件以上なら、同じ図でページを分ける
d.setChapter(4);
{
  const box = r => {
    // 定着と取引先との連携の条件の箱を、ちょうど覆う位置と大きさにする（薄くした文字に半端に重ねない）
    const x0 = r.X(2) - r.CW / 2 + 0.08, x1 = r.X(3) + r.CW / 2 - 0.08;
    return { x: x0, y: r.K0, w: x1 - x0, h: 3 * r.KP - 0.07, head: "全社展開の承認", body: "4月に12部署で始め、担当を1人増やす。" };
  };
  const { s, r } = road.page({ lead: "4月からの全社展開と、費用480万円の承認をお願いしたい。", ask: "何に使う？ → 吹き出し（12部署で開始、担当を1人増員）", focus: "決める額（480万円）", spot: rr => rr.col(1), box });
  if (!CHECK) { const b = box(r); d.focusValue(s, "480万円", { x: b.x + b.w - 2.3, y: b.y + 0.06, w: 2.2, align: "right" }); }
}

d.finalize(OUT);
