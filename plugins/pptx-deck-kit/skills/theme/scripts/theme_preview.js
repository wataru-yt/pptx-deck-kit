// theme_preview.js : テーマ（theme.json）の見た目を、4ページの見本で確かめる（表紙・章扉・棒グラフ・工程の流れ。内容は架空）
// 使い方: node theme_preview.js <theme.json のパス | 既成のテーマの名前 | auto> <出力.pptx>
//   保存する前の theme.json を、そのファイルだけで描く（個人・プロジェクトの theme.json は読まない）
//   pptxgenjs はプラグインの node_modules から読む（プラグインの中で動かせば自動で見つかる）
const path = require("path");
const pptxgen = require("pptxgenjs");
const { createDeck } = require(path.join(__dirname, "..", "..", "pptx-deck-design", "assets", "deck-parts.js"));
const [themeArg, outArg] = process.argv.slice(2);
if (!themeArg || !outArg) { console.error("使い方: node theme_preview.js <theme.json | 名前 | auto> <出力.pptx>"); process.exit(1); }
const theme = themeArg === "auto" ? "blue" : (/\.json$/i.test(themeArg) ? path.resolve(themeArg) : themeArg);
const pres = new pptxgen();
const d = createDeck(pres, { theme, themeFiles: false, footerText: "見た目の確認", nav: [{ name: "売上の推移", ch: 1 }, { name: "業務の流れ", ch: 2 }, { name: "今後の計画", ch: 3 }] }); // 章ナビは、テーマに nav があるときだけ出る
d.cover({ title: "見た目の確認", sub: "表紙の副題の例", date: "2026年11月" });
d.chapter("01", "売上の推移");
{
  const s = d.base("売上の推移", "売上は3年で1.4倍になり、2026年度は120億円の見込み。", { focus: "2026年度の棒" });
  d.figHead(s, "売上（億円、各年度）");
  d.colBars(s, { cats: ["2023年度", "2024年度", "2025年度", "2026年度"], series: [{ name: "売上", color: d.C.bar, values: [86, 97, 108, 120] }], focus: [[0, 3]], fmt: v => String(v) });
}
d.setChapter(2);
{
  const s = d.base("業務の流れ", "4工程で約9日かかり、最長は確認の約5日。", { focus: "3 確認の列" });
  d.processFlow(s, {
    steps: [{ name: "申請", work: 0.5, wait: 0.5 }, { name: "承認", work: 0.3, wait: 0.7 }, { name: "確認", work: 0.5, wait: 4.5, problem: true }, { name: "発注", work: 1, wait: 1 }],
    rows: [{ label: "担当", cells: ["申請者", "上長", "購買部", "購買部"] }],
  });
}
d.finalize(path.resolve(outArg));
