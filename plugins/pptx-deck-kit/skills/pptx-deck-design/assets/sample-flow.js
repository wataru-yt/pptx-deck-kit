// sample-flow.js : 工程の流れ・受け渡し・繰り返し・体制の図の見本（processFlow・swimlane・cycleDiagram・orgChart。内容はすべて架空）
// 実行: NODE_PATH=<pptxgenjs と jszip の入った node_modules> node sample-flow.js <出力先.pptx>
//
// 論旨：購買の申請は、支払いまで約13日かかる。最も長いのは購買部の確認で、その多くは差し戻しと待ち。
//       購買部が主導する体制で、申請の書式と確認の手順を見直す。
//
// 色の組：navy（強調 002B5C）。社内の業務の見直しを報告する、落ち着いた報告書のため。
//         問題（詰まる工程・差し戻し）は赤（accent）。
//
// 形の割り当て（1つの形は1つの意味にだけ使う）
//   processFlow  … 工程が順に進み、どこで日数がかかるか（P.3 図だけ → P.4 同じ図を敷いて待ちを言う）
//   swimlane     … 誰から誰へ受け渡すか（P.5）
//   cycleDiagram … 見積りの取り直しが回る（P.6 図だけ → P.7 同じ図を敷いて止まる区画を言う）
//   orgChart     … 見直しの体制（P.8）
//   raciTable    … 作業ごとに、誰が実行し誰が決めるか（P.9）
//
// 見本の約束
// - 題はアジェンダの項目名だけ（nav に1回書き、base の題は null にして章から取る）。主張はリードの1文
// - 全ページで見る所を1つ宣言し（focus）、強調はそこにだけ置く
// - 1つの図で2つ言うときは、series で同じ図を敷いた続きのページに分ける（2枚目は spot か only で見る所だけ残し、前面に箱を1つ）
// - 図の中の文字は、名前・値・短い札だけ（全角12字まで・1行）。補足の sub は使わない。文はリードと前面の箱に書く
const path = require("path");
const pptxgen = require("pptxgenjs");
const { createDeck } = require("./deck-parts");

const OUT = path.resolve(process.argv[2] || path.join(__dirname, "sample-flow.pptx"));
const CHECK = process.env.ECHO_CHECK === "1"; // 1 にすると続きのページを前面なしで描く（figure_diff.py で図が動いていないかを確かめる用）
const pres = new pptxgen();
const d = createDeck(pres, {
  audience: "insider", // 聞き手の前提（決まり24）。見本は部品の使い方を見せるので、前提を共有済みとして描く
  title: "購買の見直し",
  theme: "navy",
  footerText: "購買の見直し",
  nav: [
    { name: "業務の流れ", ch: 1 },
    { name: "受け渡し", ch: 2 },
    { name: "見積りの取り直し", ch: 3 },
    { name: "体制", ch: 4 },
    { name: "役割の分担", ch: 5 },
  ],
});

d.cover({ title: "購買の見直し", date: "2027年2月" });
d.agendaPage();

// ===== 1. 業務の流れ（順に進む）：工程の矢羽、所要（作業と待ち）、短い札の行を2行まで =====
// 詰まる工程の列だけ赤枠。工程ごとの説明は、続きのページの前面の箱で言う。下に箱1つ分の空きを残す（bottom）
d.setChapter(1);
const FLOW = {
  bottom: 5.3,
  headH: 0.75, // 下に置く箱がフッターにかからないように、矢羽を低くする
  labelW: 1.1, // 行の名前（2字）に合わせて詰め、工程の列を広げる
  steps: [
    { name: "起票", work: 0.5, wait: 0.5 },
    { name: "承認", work: 0.2, wait: 0.8 },
    { name: "確認", work: 0.5, wait: 4.5, problem: true },
    { name: "見積り", work: 1, wait: 2 },
    { name: "発注", work: 0.3, wait: 0.7 },
    { name: "検収・支払", work: 1, wait: 1 },
  ],
  rows: [
    { label: "担当", cells: ["申請者", "上長", "購買部", "バイヤー", "購買部", "経理部"] },
    { label: "課題", cells: ["内訳が抜ける", "上限が不統一", "!差し戻し3割", "書式が不統一", "手で転記", "照合が手作業"] },
  ],
  backs: [{ from: 2, to: 0, label: "差し戻し" }],
};
const flow = d.series((s, x) => d.processFlow(s, { ...FLOW, ...x }), { check: CHECK });
flow.page({ lead: "6工程で約13日かかり、最長は購買部の確認の約5日。", focus: "3 確認の列（赤枠・約5日）" });
flow.page({
  lead: "確認の約5日のうち、作業は半日で、残りは待ち。", focus: "3 確認の列（待ち 4.5日）",
  only: [2],
  // 箱は図の下の空きに置き、引き出し線は垂直に列の下端を指す
  box: r => { const c = r.col(2), cx = c.x + c.w / 2; return { x: c.x, y: c.y + c.h + 0.3, w: 5.2, h: 1.05, head: "待ちは4.5日、作業は半日", body: "2人が、申請を週2回まとめて確認する。", from: [cx, c.y + c.h + 0.3], to: [cx, c.y + c.h + 0.05] }; },
});

// ===== 2. 受け渡し（受け渡す）：4者とシステムのレーン、判断のひし形、帳票、システムの円柱、差し戻しは赤 =====
// 5レーン×9列に絞る（受け渡しの要点：確認から申請への差し戻し、取引先との見積り・発注、検収と支払いの登録）
// 工程の名前は1行に全角5字まで（列の幅から決まる）。2行にするときは、語の切れ目で "\n" を入れる。判断のひし形とシステムの円柱は1行（ひし形は全角3字まで）
d.setChapter(2);
{
  const s = d.base(null, "購買部の確認で、3件に1件を申請部署へ差し戻す。", { focus: "「内訳？」から申請へ戻る赤の線（3件に1件）" });
  d.swimlane(s, {
    lanes: ["申請部署", "購買部", "取引先", "経理部", "システム"],
    phases: [{ name: "申請と承認", from: 0, to: 1 }, { name: "確認", from: 2, to: 3 }, { name: "見積りと発注", from: 4, to: 5 }, { name: "納品と支払い", from: 6, to: 8 }],
    nodes: [
      { id: "apply", lane: 0, col: 0, name: "申請を\n起票する" },
      { id: "regA", lane: 4, col: 0, kind: "system", name: "申請登録" },
      { id: "ok", lane: 0, col: 1, name: "上長が\n承認する" },
      { id: "check", lane: 1, col: 2, name: "内訳を\n確認する", problem: true },
      { id: "q", lane: 1, col: 3, kind: "decision", name: "内訳？" },
      { id: "bid", lane: 2, col: 4, name: "見積りを\n出す" },
      { id: "po", lane: 1, col: 5, name: "選んで\n発注する" },
      { id: "poDoc", lane: 2, col: 5, kind: "doc", name: "注文書" },
      { id: "ship", lane: 2, col: 6, name: "納品して\n請求する" },
      { id: "inv", lane: 3, col: 6, kind: "doc", name: "請求書" },
      { id: "recv", lane: 0, col: 7, name: "検収する" },
      { id: "regR", lane: 4, col: 7, kind: "system", name: "検収登録" },
      { id: "pay", lane: 3, col: 8, name: "支払う" },
      { id: "regP", lane: 4, col: 8, kind: "system", name: "支払登録" },
    ],
    edges: [
      { from: "apply", to: "regA", dashed: true },
      { from: "apply", to: "ok" },
      { from: "ok", to: "check" },
      { from: "check", to: "q" },
      { from: "q", to: "bid", exit: "b", label: "あり" },
      { from: "q", to: "apply", label: "内訳なし　3件に1件" },
      { from: "bid", to: "po" },
      { from: "po", to: "poDoc", dashed: true },
      { from: "po", to: "ship" },
      { from: "ship", to: "inv", dashed: true },
      { from: "ship", to: "recv" },
      { from: "recv", to: "regR", dashed: true },
      { from: "recv", to: "pay" },
      { from: "pay", to: "regP", dashed: true },
    ],
    times: ["約1日", "約1日", "約5日", null, "約3日", "約1日", null, "約1日", "約1日"],
  });
}

// ===== 3. 見積りの取り直し（回る）：環を4区画に分け、中心に回すもの・回数・1周の日数 =====
// 区画の外には名前と、担当・所要の値だけ。止まる区画は、続きのページで only で残して前面の箱で言う
d.setChapter(3);
const CYCLE = {
  steps: [
    { name: "仕様を決める", meta: [["担当", "申請者"], ["所要", "約2日"]] },
    { name: "見積りを依頼する", meta: [["担当", "購買部"], ["所要", "約3日"]] },
    { name: "購買部の確認", meta: [["担当", "購買部"], ["所要", "約5日"]] },
    { name: "比べて直す", meta: [["担当", "申請者"], ["所要", "約1日"]] },
  ],
  center: { label: "見積りの取り直し", value: "5回以上", note: "1周 約11日" },
};
const cycle = d.series((s, x) => d.cycleDiagram(s, { ...CYCLE, ...x }), { check: CHECK });
cycle.page({ lead: "新しい機器は、仕様が固まるまで見積りを5回以上取り直す。", focus: "環の中心の回数（5回以上）" });
cycle.page({
  lead: "1周の約11日のうち、約5日は購買部の確認で止まる。", focus: "3 購買部の確認の区画",
  only: [2],
  // 箱は左の列の空き（区画3と区画4の名前の間）に置き、引き出し線は垂直に区画3の名前を指す
  box: r => { const l = r.label(2), x = 0.6; return { x, y: l.y - 1.25, w: 3.9, h: 1.05, head: "依頼のたびに確認する", body: "同じ取引先でも毎回見直す。", from: [l.x + 0.4, l.y - 0.3], to: [l.x + 0.4, l.y - 0.04] }; },
});

// ===== 4. 体制（分かれる）：責任者 → 部署 → チーム（リーダー）→ メンバー。主導の部署だけ強調 =====
// 部署には人数だけを添える（役目の説明を足さない）。メンバーの役割は全角5字まで（兼務は「兼務」を含めて5字）。部署とチームの名前は全角5字まで
d.setChapter(4);
{
  const s = d.base(null, "購買部が3チームで主導し、システム部と経理部が支える。", { focus: "主導の購買部の箱（15名）" });
  d.orgChart(s, {
    head: { role: "管理本部長" },
    staff: [{ tag: "事務局", role: "購買改革のPMO（2名）" }],
    staffW: 3.2,
    depts: [
      { name: "購買部", count: "15名", lead: true, teams: [
        { name: "企画担当", leader: "課長", members: [{ role: "制度設計", n: 2 }, { role: "支出分析", n: 1 }] },
        { name: "購買担当", leader: "主任", members: [{ role: "バイヤー", n: 4 }, { role: "発注事務", n: 2 }] },
        { name: "取引先", leader: "主任", members: [{ role: "審査", n: 2 }, { role: "契約", n: 1 }] },
      ] },
      { name: "システム部", count: "4名", teams: [
        { name: "システム", leader: "課長代理", members: [{ role: "開発", n: 2 }, { role: "運用", n: 1 }] },
      ] },
      { name: "経理部", count: "3名", teams: [
        { name: "支払い", leader: "課長", dual: true, members: [{ role: "支払い", n: 1 }, { role: "仕訳", n: 1 }] },
      ] },
      { name: "委託先", count: "5名", ext: true, teams: [
        { name: "開発担当", leader: "PM", members: [{ role: "開発", n: 3 }, { role: "デザイン", n: 1 }] },
      ] },
    ],
  });
}

// ===== 5. 役割の分担：作業×担当に R・A・C・I の印だけ。決める人（A）がいない作業を赤枠で =====
// セルに文字を入れない。A は1つの作業に1人。A がいない行は problem を付ける（部品が「A がいない」を赤で書く）
d.setChapter(5);
{
  const s = d.base(null, "規程の改定だけ、決める人（A）がいない。", { focus: "赤枠の「規程の改定」の行" });
  d.raciTable(s, {
    roles: ["管理本部長", "購買部", "システム部", "経理部", "申請部署"],
    rows: [
      { name: "方針の決定", marks: ["A", "R", "C", "C", "I"] },
      { name: "申請書式の改定", marks: ["I", "A", "C", "", "C"] },
      { name: "確認の手順", marks: ["", "A", "C", "", "I"] },
      { name: "規程の改定", marks: ["", "R", "", "R", "C"], problem: true },
      { name: "システムの設定", marks: ["", "C", "A", "", "I"] },
      { name: "支払いの照合", marks: ["", "C", "", "A", ""] },
    ],
  });
}

d.finalize(OUT);
