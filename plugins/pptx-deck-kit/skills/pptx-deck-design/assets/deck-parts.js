"use strict";
// deck-parts.js : pptx の資料を pptxgenjs で組む部品ライブラリ（pptx-deck-design）
// 資料ごとに作業フォルダへコピーして使う。スキルの置き場所から require しない
// （部品を後で直したとき、過去の資料の再生成結果が黙って変わり、build.js と pptx の一致が崩れるのを防ぐ）。
//
// 使い方:
//   const pptxgen = require("pptxgenjs");
//   const { createDeck } = require("./deck-parts");
//   const pres = new pptxgen();
//   const d = createDeck(pres, { nav: [{ name: "概要", ch: 1 }, { name: "反省", ch: 2 }] });
//   d.cover({ title: "資料名", date: "2026年11月" });
//   d.chapter("1.", "章題");
//   const s = d.base("テーマ", "このページの結論を言い切る。", { focus: "見る所", ask: "聴講者の疑問 → 答えの場所" });
//   d.blockTable(s, { cols: [...], rows: [...] });
//   d.finalize("out.pptx");
//
// 単位はインチ、文字は pt。座標の既定値は振り返り報告書の build.js の実測値。
// 部品は「1ページ1メッセージ、強調は1か所、罫線を使わず塗りブロックと白い溝」を既定動作にしている。
// 文字があふれそうなとき・リードが3行になるときは console.warn で知らせる（描画の確認は別途、画像で行う）。

const path = require("path");
const fs = require("fs");
const os = require("os");

// ===== トークン =====

// 色（5.1）。役割を固定する：ink＝主・強調、accent＝問題・ゲート・視線誘導、gray＝ラベル・退かせる
const C = {
  ink: "002B5C",    // タイトル・見出し・主文字
  emph: "002B5C",   // 強調の枠・線 2〜2.25pt、データの印（主役の棒・面・点）、番号丸。既定は ink と同じ。テーマで別の色にできる
  dark: "002B5C",   // 区切りページの地。既定は ink と同じ
  soft: "4A5568",   // 説明文、表の値（読ませる本文の最も薄い色）
  gray: "8A8F99",   // 列見出し、ラベル、接頭、補助の矢印、章名（読ませる本文には使わない）
  pale: "B3B8C1",   // ミニ目次の非選択、「—」
  line: "D9DDE3",   // 細線 0.75pt、空枠の破線、画像・動画の枠
  faint: "979DA8",  // 目次の「P.n」、比較対象の線
  tint: "F4F6F9",   // 標準の塗りブロック
  strong: "E3EAF4", // 強調ブロック（原則1ページ1か所）
  accent: "B3261E", // 問題・注意、赤枠、ゲートの破線、番号丸、マイルストーン
  link: "0066CC",   // リンク
  white: "FFFFFF",
  grid: "EEEEEE",   // グラフの目盛線
  bar: "C9CED6",    // 単系列の棒の標準色（主役の1本だけ ink で塗る。濃紺の境界は tokens.md）
};
let TINT = C.tint, TINT_STRONG = C.strong;
// 色の組を、強調の色1つから作る（題材に合わせて選ぶ。2026-11-24）。文字は黒に近い色、強調の塗りは強調の色を白で薄めた色
//   themeFrom("2F6BFF") → { ink, emph, dark, soft, gray, pale, line, faint, tint, strong, accent }。o で個別に上書きできる
const mixHex = (a, b, t) => [0, 2, 4].map(i => Math.round(parseInt(a.slice(i, i + 2), 16) * (1 - t) + parseInt(b.slice(i, i + 2), 16) * t).toString(16).padStart(2, "0")).join("").toUpperCase();
const themeFrom = (emph, o = {}) => ({ ink: "1A1F2B", emph, dark: mixHex(emph, "0B0E14", 0.72), soft: "5A6372", gray: "8D95A3", pale: "C4C9D1", line: "E1E4EA", faint: "9AA1AD", tint: "F5F6F8", strong: mixHex(emph, "FFFFFF", 0.86), accent: "C5221F", ...o });
// 色の組の見本（強調の色）。題材に合わせて選ぶ。赤は「問題」の色なので、強調には使わない
const PALETTES = {
  navy: { ink: "002B5C", emph: "002B5C", dark: "002B5C" },   // 落ち着いた報告書（これまでの既定）
  blue: themeFrom("2F6BFF"),     // IT・ソフトウェア・スタートアップ
  teal: themeFrom("0E7C86"),     // 医療・ヘルスケア・公共
  green: themeFrom("1F8A5B"),    // 環境・農業・食品・サステナビリティ
  orange: themeFrom("D9730D"),   // 小売・消費財・飲食・教育
  purple: themeFrom("6B4FD8"),   // クリエイティブ・エンタメ・AI
  slate: themeFrom("3D4A5C"),    // 金融・法務・製造（控えめに）
}; // テーマで差し替わる（createDeck の opts.theme）
// Before/After の面（4.2）。中のブロックは白にする
const BA_BG = { before: "EEF0F3", after: "E4EDF8" };
// 区分色（ガント・積み上げ棒・内訳で同じ名前・同じ色・同じ順にする。5.1）。002B5C は白文字になり強く見えるので、主役の区分だけに使う
const CAT = ["AFC6A6", "E5C47F", "D3A39C", "DADDE1", "002B5C"];
// 別系統の区分（費目など。面積棒 p38 の値）
const CAT_ALT = ["5B8DB8", "6BAA8C", "E0A030", "9A86C2"];
// 付箋（4.7）
const STICKY = { yellow: "FFF1A6", pink: "FFD9E2", blue: "CFE6FF", text: "1D2433", off: "6B7280" };

// 文字の段（5.2）。本文に使うのはこの6段だけ。太さは bold と regular の2段
// 小さい文字の段は、2026-11-24 に上げた（本文 10.5・注 9.5 → 12 → 案B の 16）。映して見せる資料で読める大きさ
// 案B、2026-11-24：映して見せて読める大きさ。本文 16・名前 18・やや大きい文字 20（12pt の版は、横 13.33in の版では小さすぎた）
const FS = { title: 26, lead: 20, head: 20, block: 20, body: 19, note: 19 };
const FOOT_FS = 16; // フッター・ページ番号・ミニ目次は、本文より小さくしておく
// リードの行の大きさ（2026-11-24）：映して見せる資料で、主張の1文が本文と同じくらいに見えないよう、題（26pt）に次ぐ大きさにする。FS.lead（20pt）は図の中のやや大きい文字として残す
let LEAD_FS = 22; // テーマの sizes.lead で替わる（下限 18）
// グラフの見る所（2026-11-24）：見る所の値は 36pt（FOCUS_FS）・強調の色（問題なら赤）、見る所の線は 3pt。他の線は 1.25pt、棒は灰
let FOCUS_FS = 36; const FOCUS_LW = 3; // FOCUS_FS はテーマの sizes.focus で替わる（下限 24）
// 表紙・章扉・目次・数値の強調だけで使う大きさ（5.2）
const FS_DISPLAY = { cover: 48, chapter: 30, chapterLong: 26, chapterNum: 20, agendaNo: 20, agendaHead: 19, num: 22 };

// 溝の段（5.3）：密な表／標準／行を強く分ける／列の間（Before/After・工程の横並び）
const GAP = { dense: 0.06, std: 0.12, para: 0.2, col: 0.5 };
// ブロック内の文字の余白（5.3）
const PAD = { x: 0.18, y: 0.12 };

// グリッドと縦の固定位置（2.1・5.4）
const W = 13.333, H = 7.5, BODY = 2.2;
let M = 0.6, BW = W - 2 * M; // 左右の余白と本文の幅。テーマの layout.margin で替わる（広げるだけ。左右の飾りがあるテンプレート向け）
let BOTTOM = 6.6, FH = 0.305; // 本文の下端とフッターの帯の高さ。テーマの footer.height・layout.bodyBottom で替わる（下端は上げるだけ）
const Y = { section: 0.38, title: 0.66, lead: 1.32, body: BODY, bottom: BOTTOM, source: H - 0.66, footer: H - FH };

// 書体。テーマの font で替わる：face＝本文（F）、heading＝題・表紙・章扉（FHEAD。既定は本文と同じ）、number＝番号丸などの数字（FNUM）
//   既定は BIZ UDPゴシック（Windows 10 以降に標準。Mac は Microsoft 365 のクラウドフォントか Google Fonts）。数字の書体も本文にそろえる
let F = "BIZ UDPGothic", FHEAD = F, FNUM = F;

// ===== 文字幅の見積もり =====
// EMW：英数字の幅の倍率。テーマの font.widthScale で替わる（漢字・かなの幅はどの書体も 1em なので替えない）
//   BIZ UDPゴシックは英数字が広く、約物が狭いので、実測の幅を使う（BIZ_LATIN）。ほかの書体は英数字 0.6em（游ゴシック・Noto Sans JP）にメイリオなどの倍率を掛ける
//   部品のあふれの見積もりと、面の中の文字の大きさの調整に効く
const isBizP = face => /^BIZ ?UDP/i.test(String(face));
const bizEm = c => (c === " " ? 0.34 : /[0-9]/.test(c) ? 0.75 : /[MWmw%@]/.test(c) ? 0.98 : /[A-Z]/.test(c) ? 0.77 : /[iIjl.,:;'!|()\[\]]/.test(c) ? 0.36 : /[a-z]/.test(c) ? 0.63 : c.charCodeAt(0) < 0x2000 ? 0.6 : /[、。「」（）・：]/.test(c) ? 0.6 : 1);
let EMW = 1, BIZ_W = true;
const em = t => [...String(t)].reduce((a, c) => a + (BIZ_W ? bizEm(c) * (c.charCodeAt(0) < 0x2000 ? EMW : 1) : c.charCodeAt(0) < 0x2000 ? 0.6 * EMW : 1), 0);
const textW = (t, pt) => em(t) * pt / 72;
// 幅 w に入れたときの行数（明示の改行を含む）
const lineCount = (t, pt, w) => String(t).split("\n").reduce((a, l) => a + Math.max(1, Math.ceil(textW(l, pt) / w - 0.02)), 0);
// 折り返した最後の行の長さ（em）。1〜2文字だけの行（9章「1語だけの行」）を見つけるのに使う
const tails = (t, pt, w) => { const cap = w * 72 / pt; return String(t).split("\n").map(l => em(l)).filter(e => e > cap).map(e => e - Math.floor(e / cap) * cap); };
const hasOrphan = (t, pt, w) => !!t && tails(t, pt, w).some(r => r > 0 && r <= 2.2);
// 1行の高さ（PowerPoint は文字サイズ×1.2×行間）
const lineH = (pt, ls = 1.25) => pt * 1.2 * ls / 72;

// リードの改行：1行に収まれば1行、2行なら「。」で改行、1文が収まらないときだけ「、」の直後で折る
const leadBreak = (text, fs = LEAD_FS, width = BW) => {
  const cap = width * 0.97 * 72 / fs; // 1行に入る em 数
  const flat = String(text).replace(/\n/g, "");
  if (em(flat) <= cap) return flat;
  const pack = parts => { const out = []; let cur = ""; parts.forEach(p => { if (cur && em(cur + p) > cap) { out.push(cur); cur = p; } else cur += p; }); if (cur) out.push(cur); return out; };
  const lines = [];
  (flat.match(/[^。]+。?/g) || [flat]).forEach(sen => {
    if (em(sen) <= cap) lines.push(sen);
    else lines.push(...pack(sen.match(/[^、]+、?/g) || [sen]));
  });
  const merged = pack(lines);
  return merged.some(l => em(l) > cap) ? flat : merged.join("\n");
};

// 素材（アイコン）の場所の候補。deck-parts.js は資料ごとに作業フォルダへコピーして使うので、
// コピー先 → 作業フォルダ → スキルの置き場所（Claude Code・Codex。CLAUDE_CONFIG_DIR・CODEX_HOME も見る）の順に探す
const assetDirs = sub => {
  const homes = [process.env.CLAUDE_CONFIG_DIR, path.join(os.homedir(), ".claude"), process.env.CODEX_HOME, path.join(os.homedir(), ".codex")].filter(Boolean);
  return [path.join(__dirname, sub), path.join(__dirname, "assets", sub), path.join(process.cwd(), sub), path.join(process.cwd(), "assets", sub),
    ...homes.map(h => path.join(h, "skills", "pptx-deck-design", "assets", sub))];
};

// 色の明るさ（帯の上の文字を白にするかの判定）
const isDark = hex => { const n = parseInt(hex, 16), r = n >> 16, g = (n >> 8) & 255, b = n & 255; return (0.299 * r + 0.587 * g + 0.114 * b) < 110; };

// 明るさと読みやすさ（WCAG の相対輝度と明暗比）
const relLum = hex => { const ch = [0, 2, 4].map(i => parseInt(hex.slice(i, i + 2), 16) / 255).map(v => (v <= 0.03928 ? v / 12.92 : ((v + 0.055) / 1.055) ** 2.4)); return 0.2126 * ch[0] + 0.7152 * ch[1] + 0.0722 * ch[2]; };
const contrast = (a, b) => { const x = relLum(a), y = relLum(b); return (Math.max(x, y) + 0.05) / (Math.min(x, y) + 0.05); };
// PNG の画素を読む（8bit・非インターレースだけ。背景の絵の上で文字が読めるかを確かめるのに使う）。読めなければ null
const PNG_CACHE = {};
const pngReader = file => {
  if (file in PNG_CACHE) return PNG_CACHE[file];
  let out = null;
  try {
    const zlib = require("zlib"), b = fs.readFileSync(file);
    if (b.slice(1, 4).toString() === "PNG") {
      let w = 0, h = 0, depth = 0, type = 0, lace = 0, pal = null, trns = null; const idat = [];
      for (let i = 8; i < b.length;) {
        const len = b.readUInt32BE(i), tag = b.slice(i + 4, i + 8).toString(), d = b.slice(i + 8, i + 8 + len);
        if (tag === "IHDR") { w = d.readUInt32BE(0); h = d.readUInt32BE(4); depth = d[8]; type = d[9]; lace = d[12]; }
        else if (tag === "PLTE") pal = d; else if (tag === "tRNS") trns = d; else if (tag === "IDAT") idat.push(d);
        i += 12 + len;
      }
      const ch = { 0: 1, 2: 3, 3: 1, 4: 2, 6: 4 }[type];
      if (depth === 8 && !lace && ch) {
        const raw = zlib.inflateSync(Buffer.concat(idat)), stride = w * ch, px = Buffer.alloc(stride * h);
        for (let y = 0; y < h; y++) {
          const f = raw[y * (stride + 1)], line = raw.slice(y * (stride + 1) + 1, (y + 1) * (stride + 1));
          for (let x = 0; x < stride; x++) {
            const a = x >= ch ? px[y * stride + x - ch] : 0, up = y ? px[(y - 1) * stride + x] : 0, c = y && x >= ch ? px[(y - 1) * stride + x - ch] : 0;
            let v = line[x];
            if (f === 1) v += a; else if (f === 2) v += up; else if (f === 3) v += (a + up) >> 1;
            else if (f === 4) { const p = a + up - c, pa = Math.abs(p - a), pb = Math.abs(p - up), pc = Math.abs(p - c); v += pa <= pb && pa <= pc ? a : pb <= pc ? up : c; }
            px[y * stride + x] = v & 255;
          }
        }
        // (fx, fy) は 0〜1。[r, g, b, 不透明度 0〜1] を返す
        out = (fx, fy) => {
          const i = Math.min(h - 1, Math.max(0, Math.floor(fy * h))) * stride + Math.min(w - 1, Math.max(0, Math.floor(fx * w))) * ch;
          if (type === 2) return [px[i], px[i + 1], px[i + 2], 1];
          if (type === 6) return [px[i], px[i + 1], px[i + 2], px[i + 3] / 255];
          if (type === 0) return [px[i], px[i], px[i], 1];
          if (type === 4) return [px[i], px[i], px[i], px[i + 1] / 255];
          const k = px[i]; return [pal[k * 3], pal[k * 3 + 1], pal[k * 3 + 2], trns && k < trns.length ? trns[k] / 255 : 1];
        };
      }
    }
  } catch (e) { /* 読めない */ }
  return (PNG_CACHE[file] = out);
};
const toHex = rgb => rgb.map(v => Math.round(v).toString(16).padStart(2, "0")).join("").toUpperCase();
const blend = (under, over, alpha) => toHex([0, 2, 4].map((i, k) => parseInt(under.slice(i, i + 2), 16) * (1 - alpha) + (Array.isArray(over) ? over[k] : parseInt(over.slice(i, i + 2), 16)) * alpha));

// ===== テーマ（トンマナとフォーマット。2026-11-24） =====
// 各自の見た目（色・書体・文字の段・フッター・表紙・題とリードの位置・自分のアイコン）を、プラグインの外の JSON に置いて読み込む。
// プラグインは更新のたびに丸ごと入れ替わるので、各自の設定はプラグインの中に置かない。
// 読み込みの順（上ほど優先。項目ごとに上の層が勝つ）：
//   1. build.js の createDeck の opts（footer・footerText・copyright・logo・font・sizes・layout・cover・pageNumber・confidential・iconDir）
//   2. opts.theme（色の組の名前 blue など／既成のテーマの名前 standard・corporate・pitch／JSON のパス／オブジェクト）
//   3. プロジェクトの .claude/pptx-deck/theme.json（作業フォルダから上へ探す）
//   4. 個人の ~/.claude/pptx-deck/theme.json
//   5. 今の既定（下の DEFAULT_TOKENS）
// 何も置かなければ、今とまったく同じ出力になる。opts.themeFiles: false か、環境変数 PPTX_DECK_THEME_FILES=0 で 3・4 を読まない。
const DEFAULT_TOKENS = { FS: { ...FS }, LEAD_FS, FOCUS_FS, F, FNUM, Y: { ...Y }, BOTTOM, FH, M, BW };
const FONT_WIDTH = { "Meiryo": 1.12, "メイリオ": 1.12 }; // 書体ごとの幅の倍率の既定（游ゴシックより幅の広い書体）
const THEME_SECTIONS = ["colors", "font", "sizes", "footer", "cover", "layout", "icons", "master", "nav"];
const THEME_KEYS = new Set(["$schema", "name", "description", "extends", ...THEME_SECTIONS, "mode", "note"]);
// 文字の段：テーマの名前 → 部品の中の名前。下限を下回る値は使わず、下限にして警告する
//   title＝題（FS.title）、lead＝リード（LEAD_FS）、large＝やや大きい文字（FS.lead）、name＝名前・見出し（FS.head・FS.block）、body＝本文（FS.body）、note＝注（FS.note）、focus＝主役の値（FOCUS_FS）
const SIZE_MIN = { title: 20, lead: 18, large: 14, name: 14, body: 14, note: 14, focus: 24 };
const SIZE_MAX = { title: 32, lead: 26, large: 24, name: 24, body: 22, note: 22, focus: 48 }; // 超えても使うが、部品の箱からあふれやすいので知らせる
// 題・リード・章名の縦の位置と、本文の下端（インチ）。本文の上端（y 2.2）より上と、下端より下だけで動かせる
const LAYOUT_RANGE = { sectionY: [0.05, 0.5], titleY: [0.12, 0.85], leadY: [1.15, 1.45], bodyBottom: [5.6, 6.6], margin: [0.6, 1.0] };
const FOOTER_KEYS = new Set(["style", "pageNumber", "text", "copyright", "logo", "confidential", "bandColor", "height", "color"]);
const NAV_KEYS = new Set(["style", "align", "x", "y", "w", "h", "size", "color", "activeColor", "fill", "activeFill", "numbers"]);
// 飾りの層（master）：ページの種類ごとに、背景・飾り（objects）・文字の置き場所（boxes）を持つ
const PAGE_TYPES = { content: ["section", "title", "lead"], cover: ["title", "sub", "kicker", "date", "affil", "logo"], chapter: ["num", "title", "sub"], divider: ["title", "sub"] };
const MASTER_KEYS = new Set(["background", "objects", "boxes", "footer", "pageNumber", "nav"]);
const OBJ_KEYS = new Set(["type", "x", "y", "w", "h", "color", "line", "transparency", "rotate", "flipH", "flipV", "path", "text", "size", "bold", "align", "valign", "font", "width", "dash"]);
const BOX_KEYS = new Set(["x", "y", "w", "h", "color", "size", "align", "valign", "bold", "font"]);
const COVER_KEYS = new Set(["style", "logo", "logoPos", "affil"]);
const FONT_KEYS = new Set(["face", "heading", "number", "widthScale"]);
const themeWarn = (src, msg) => console.warn(`[deck-parts] テーマ（${src}）: ${msg}`);
const homeDir = () => process.env.CLAUDE_CONFIG_DIR || path.join(os.homedir(), ".claude");
const expandHome = p => (typeof p === "string" && /^~(\/|$)/.test(p) ? path.join(os.homedir(), p.slice(1)) : p);
// 既成のテーマ（themes/<名前>.json）の置き場所の候補。deck-parts.js のコピー先 → 作業フォルダ → スキル・プラグインの置き場所
const pluginAssetDirs = sub => {
  const out = [];
  if (process.env.CLAUDE_PLUGIN_ROOT) out.push(path.join(process.env.CLAUDE_PLUGIN_ROOT, "skills", "pptx-deck-design", "assets", sub));
  const cache = path.join(homeDir(), "plugins", "cache");
  try {
    for (const mk of fs.readdirSync(cache)) {
      const base = path.join(cache, mk, "pptx-deck-kit");
      if (!fs.existsSync(base)) continue;
      fs.readdirSync(base).sort().reverse().forEach(v => out.push(path.join(base, v, "skills", "pptx-deck-design", "assets", sub)));
    }
  } catch (e) { /* キャッシュが無い */ }
  return out;
};
const themeFileOf = name => [...assetDirs("themes"), ...pluginAssetDirs("themes")].map(d => path.join(d, `${name}.json`)).find(p => fs.existsSync(p));
const builtinThemes = () => { const set = new Set(); [...assetDirs("themes"), ...pluginAssetDirs("themes")].forEach(d => { try { fs.readdirSync(d).filter(f => /\.json$/.test(f)).forEach(f => set.add(f.replace(/\.json$/, ""))); } catch (e) { /* 無い */ } }); return [...set]; };
// 色の組の展開（今の決め方と同じ）：名前なら PALETTES、強調の色だけなら themeFrom で残りを作る、ink だけなら emph・dark も同じ色
const expandColors = (c, src, themeForm = false) => {
  let t0;
  if (typeof c === "string") { if (!PALETTES[c]) { themeWarn(src, `colors の名前が無い（${c}）。${Object.keys(PALETTES).join("・")} のどれか`); return {}; } t0 = { ...PALETTES[c] }; }
  else if (c && typeof c === "object") {
    t0 = {};
    for (const [k, v0] of Object.entries(c)) {
      if (!(k in C)) { themeWarn(src, `colors に知らない色の名前がある（${k}）。使える名前: ${Object.keys(C).join("・")}`); continue; }
      const v = String(v0).replace(/^#/, "");
      if (!/^[0-9A-Fa-f]{6}$/.test(v)) { themeWarn(src, `colors.${k} が6桁の16進の色でない（${v0}）。使わない`); continue; }
      t0[k] = v;
    }
  } else return {};
  const th = t0.emph && !t0.strong && !t0.ink ? themeFrom(t0.emph, t0) : { ...t0 };
  // テーマの JSON の形では、文字の色（ink）と強調の色（emph）を両方書いて strong・dark を書かなければ、強調の色から作る（今の { emph } と同じ作り方）
  if (themeForm && t0.ink && t0.emph) { if (!t0.strong) th.strong = mixHex(t0.emph, "FFFFFF", 0.86); if (!t0.dark) th.dark = mixHex(t0.emph, "0B0E14", 0.72); }
  if (th.ink && !th.emph) th.emph = th.ink;
  if (th.ink && !th.dark) th.dark = th.ink;
  return th;
};
const checkKeys = (obj, keys, where, src) => Object.keys(obj).forEach(k => { if (!keys.has(k)) themeWarn(src, `${where} に知らない項目がある（${k}）。使える項目: ${[...keys].join("・")}`); });
// テーマの1つの層を、決まった形にそろえる（パスは、その JSON の場所から解く）
const normLayer = (raw, src, dir) => {
  const L = { src, dir };
  Object.keys(raw).forEach(k => { if (!THEME_KEYS.has(k)) themeWarn(src, `知らない項目がある（${k}）。使える項目: ${[...THEME_KEYS].join("・")}`); });
  const abs = p => (p ? path.resolve(dir, expandHome(String(p))) : p);
  if (raw.colors !== undefined) { L.colors = expandColors(raw.colors, src, true); L.colorSrc = src; }
  if (raw.font !== undefined) { const f = typeof raw.font === "string" ? { face: raw.font } : { ...raw.font }; checkKeys(f, FONT_KEYS, "font", src); L.font = f; }
  if (raw.sizes !== undefined) L.sizes = { ...raw.sizes };
  if (raw.footer !== undefined) { const f = typeof raw.footer === "string" ? { style: raw.footer } : { ...raw.footer }; checkKeys(f, FOOTER_KEYS, "footer", src); if (f.logo) f.logo = abs(f.logo); L.footer = f; }
  if (raw.cover !== undefined) { const c = { ...raw.cover }; checkKeys(c, COVER_KEYS, "cover", src); if (c.logo) c.logo = abs(c.logo); L.cover = c; }
  if (raw.layout !== undefined) L.layout = { ...raw.layout };
  if (raw.icons !== undefined) L.icons = [].concat(raw.icons).filter(Boolean).map(abs);
  if (raw.nav !== undefined) { const n = typeof raw.nav === "string" || raw.nav === false ? { style: raw.nav } : { ...raw.nav }; checkKeys(n, NAV_KEYS, "nav", src); L.nav = n; }
  if (raw.master !== undefined) {
    L.master = {};
    Object.entries(raw.master || {}).forEach(([type, p]) => {
      if (!PAGE_TYPES[type]) { themeWarn(src, `master に知らないページの種類がある（${type}）。使える種類: ${Object.keys(PAGE_TYPES).join("・")}`); return; }
      checkKeys(p, MASTER_KEYS, `master.${type}`, src);
      const out = { objects: [], boxes: {} };
      if (p.footer === false) out.footer = false;
      if (p.pageNumber !== undefined) out.pageNumber = p.pageNumber;
      if (p.nav !== undefined) { if (p.nav && typeof p.nav === "object") checkKeys(p.nav, NAV_KEYS, `master.${type}.nav`, src); out.nav = p.nav; }
      if (p.background !== undefined) {
        const bg = typeof p.background === "string" ? (/\.(png|jpe?g|gif)$/i.test(p.background) ? { image: p.background } : { color: p.background }) : { ...p.background };
        if (bg.image) bg.image = abs(bg.image);
        out.background = bg;
      }
      (p.objects || []).forEach((o, i) => { checkKeys(o, OBJ_KEYS, `master.${type}.objects[${i}]`, src); out.objects.push({ ...o, ...(o.path ? { path: abs(o.path) } : {}) }); });
      Object.entries(p.boxes || {}).forEach(([name, b]) => {
        if (!PAGE_TYPES[type].includes(name)) { themeWarn(src, `master.${type}.boxes に知らない名前がある（${name}）。使える名前: ${PAGE_TYPES[type].join("・")}`); return; }
        checkKeys(b, BOX_KEYS, `master.${type}.boxes.${name}`, src); out.boxes[name] = { ...b };
      });
      L.master[type] = out;
    });
  }
  return L;
};
// テーマの指定（名前・パス・オブジェクト）を、層の並び（下から上）にする。extends で別のテーマを下に敷ける
const themeLayers = (ref, baseDir, src, depth = 0) => {
  if (depth > 5) throw new Error(`createDeck：テーマの extends が深すぎる（${src}）`);
  if (ref == null || ref === false) return [];
  if (typeof ref === "string") {
    if (PALETTES[ref]) return [{ src: `${src}: ${ref}`, dir: baseDir, colors: expandColors(ref, src), colorSrc: `${src}: ${ref}` }];
    let file = null;
    if (/[\\/]|\.json$/i.test(ref) || ref.startsWith("~")) { file = path.resolve(baseDir, expandHome(ref)); if (!fs.existsSync(file)) throw new Error(`createDeck：テーマの JSON が見つからない（${file}）`); }
    else file = themeFileOf(ref);
    if (!file) throw new Error(`createDeck：theme の名前が無い（${ref}）。色の組 ${Object.keys(PALETTES).join("・")}、既成のテーマ ${builtinThemes().join("・") || "standard・corporate・pitch（themes/ が見つからない。deck-parts.js と同じ所に themes/ をコピーする）"} のどれか、または JSON のパスか色を渡す`);
    return themeFileLayers(file, depth);
  }
  if (typeof ref === "object") {
    const isTheme = Object.keys(ref).some(k => THEME_KEYS.has(k));
    if (!isTheme) return [{ src, dir: baseDir, colors: expandColors(ref, src), colorSrc: src }]; // 今の書き方（{ emph } や色の全部）
    const below = ref.extends ? themeLayers(ref.extends, baseDir, `${src} の extends`, depth + 1) : [];
    return [...below, normLayer(ref, src, baseDir)];
  }
  throw new Error(`createDeck：theme は名前・JSON のパス・オブジェクトのどれかにする（${typeof ref}）`);
};
const themeFileLayers = (file, depth = 0) => {
  let raw;
  try { raw = JSON.parse(fs.readFileSync(file, "utf8")); } catch (e) { throw new Error(`createDeck：テーマの JSON が読めない（${file}）: ${e.message}`); }
  const dir = path.dirname(file);
  const below = raw.extends ? themeLayers(raw.extends, dir, `${file} の extends`, depth + 1) : [];
  return [...below, normLayer(raw, file, dir)];
};
// プロジェクトの .claude/pptx-deck/theme.json（作業フォルダから上へ）と、個人の ~/.claude/pptx-deck/theme.json
const findThemeFiles = () => {
  const personal = path.join(homeDir(), "pptx-deck", "theme.json");
  let project = null;
  for (let d = process.cwd(); ; d = path.dirname(d)) {
    const p = path.join(d, ".claude", "pptx-deck", "theme.json");
    if (p !== personal && fs.existsSync(p)) { project = p; break; }
    if (path.dirname(d) === d) break;
  }
  return { personal: fs.existsSync(personal) ? personal : null, project };
};
// 画像の縦横（PNG・JPEG・GIF）。ロゴを縦横比のまま置くのに使う。読めなければ null
const imageSize = p => {
  try {
    const b = fs.readFileSync(p);
    if (b.slice(1, 4).toString() === "PNG") return { w: b.readUInt32BE(16), h: b.readUInt32BE(20) };
    if (b.slice(0, 3).toString() === "GIF") return { w: b.readUInt16LE(6), h: b.readUInt16LE(8) };
    if (b[0] === 0xff && b[1] === 0xd8) {
      for (let i = 2; i < b.length - 9;) {
        if (b[i] !== 0xff) { i++; continue; }
        const m = b[i + 1], len = b.readUInt16BE(i + 2);
        if (m >= 0xc0 && m <= 0xcf && m !== 0xc4 && m !== 0xc8 && m !== 0xcc) return { w: b.readUInt16BE(i + 7), h: b.readUInt16BE(i + 5) };
        i += 2 + len;
      }
    }
  } catch (e) { /* 読めない */ }
  return null;
};
// 箱（maxW × maxH）に縦横比のまま収める
const fitBox = (p, maxW, maxH) => { const sz = imageSize(p); if (!sz || !sz.w || !sz.h) return { w: maxW, h: maxH, unknown: true }; const r = sz.w / sz.h; return r > maxW / maxH ? { w: maxW, h: maxW / r } : { w: maxH * r, h: maxH }; };

// ===== デッキ =====
// createDeck：状態（ページ番号・現在の章・ページの記録・後描き・グラフの後処理）を内部に持ち、部品関数を返す
//   opts.nav      ミニ目次の項目 [{ name, ch, subs?: [ch, ...] }]。null（既定）ならミニ目次を描かない
//   opts.copyright フッターの帯の左の表記（footer が "band" のときだけ。既定は空）
//   opts.logo     フッターの帯の右に置くロゴの画像のパス（footer が "band" のときだけ。既定はなし）
//   opts.theme    色の組。色のトークンを差し替える { ink（文字）, emph（強調の枠・線・データの印）, dark（区切りページの地）, soft, gray, pale, line, faint, tint, strong, accent（問題の赤）, bar }。'#' なしの6桁
//   opts.title    pptx の文書名
//   opts.titleStyle "action" で、題に結論の1文を書く形にする（リードは置かない。20pt・2行まで）。既定は項目名の題＋リード
//   opts.wordy    リードの点検に足す言い回し [[正規表現, 直し方], ...]。false で言い回しの点検をしない
//   opts.footerText フッターの左に出す文字（資料名など。footer が "page" のとき。細い線の下に置く）
//   opts.miniNav  true で、右上にミニ目次（現在地）を出す。既定は出さない
//   opts.footer   "page"（既定。ページ番号だけを右下に置く）| "band"（下に濃紺の帯。左に opts.copyright、中央にページ番号、右に opts.logo）
//   ── テーマ（見た目の設定。上の footer・footerText・copyright・logo も、テーマの footer の項目を上書きする）──
//   opts.theme    色の組の名前（blue など）・既成のテーマの名前（standard・corporate・pitch）・テーマの JSON のパス・オブジェクトのどれか。
//                 オブジェクトは、今の書き方（{ emph } や色の全部）か、テーマの形（{ colors, font, sizes, footer, cover, layout, icons, extends }）
//   opts.font     書体（"Meiryo" か { face, heading, number }）
//   opts.sizes    文字の段 { title, lead, large, name, body, note, focus }（下限：本文・注 14、リード 18、主役の値 24。下回ると下限を使って警告）
//   opts.layout   題・リード・章名の縦の位置 { titleY, leadY, sectionY }（動かせる範囲だけ。ほかの位置は固定）
//   opts.cover    表紙 { style: "auto"|"simple"|"full", logo, logoPos: "top-left"|"top-right"|"center", affil }
//   opts.pageNumber false でページ番号を出さない　opts.confidential 「社外秘」などの表記（全ページの右下）
//   opts.themeFiles false で、プロジェクトと個人の theme.json を読まない（環境変数 PPTX_DECK_THEME_FILES=0 でも同じ）
function createDeck(pres, opts = {}) {
  pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5（16:9 固定。4:3 には対応しない）
  if (opts.title) pres.title = opts.title;
  // ---- テーマ：層を下から積む（個人 → プロジェクト → opts.theme → opts）。項目ごとに上の層が勝つ ----
  // 文字の段・書体・位置は、前の createDeck の値を持ち越さないよう、既定に戻してから当てる（色の C は今まで通り、当てた分だけ替わる）
  Object.assign(FS, DEFAULT_TOKENS.FS); Object.assign(Y, DEFAULT_TOKENS.Y);
  LEAD_FS = DEFAULT_TOKENS.LEAD_FS; FOCUS_FS = DEFAULT_TOKENS.FOCUS_FS; F = DEFAULT_TOKENS.F; FHEAD = F; FNUM = DEFAULT_TOKENS.FNUM; EMW = 1; BIZ_W = true;
  BOTTOM = DEFAULT_TOKENS.BOTTOM; FH = DEFAULT_TOKENS.FH; M = DEFAULT_TOKENS.M; BW = DEFAULT_TOKENS.BW;
  const layers = [];
  const useFiles = opts.themeFiles !== false && process.env.PPTX_DECK_THEME_FILES !== "0";
  const files = useFiles ? findThemeFiles() : {};
  if (useFiles && !files.personal && !files.project) console.log("[deck-parts] 見た目の設定（theme.json）が無い。質問のツール（AskUserQuestion）が使えるなら、見た目（会社の色・ロゴ／おまかせ）を1回だけ聞く（/pptx-deck-kit:theme）。使えなければ聞かずにおまかせで進む");
  if (files.personal) layers.push(...themeFileLayers(files.personal));
  if (files.project) layers.push(...themeFileLayers(files.project));
  if (opts.theme) layers.push(...themeLayers(opts.theme, process.cwd(), "opts.theme").map(L => ({ ...L, fromOpts: true })));
  {
    // build.js の opts（いちばん上の層）
    const top = { src: "build.js の opts", dir: process.cwd() }, ft = {};
    if (opts.footer !== undefined) Object.assign(ft, typeof opts.footer === "string" ? { style: opts.footer } : opts.footer);
    if (opts.footerText !== undefined) ft.text = opts.footerText;
    if (opts.copyright !== undefined) ft.copyright = opts.copyright;
    if (opts.logo !== undefined) { ft.logo = opts.logo; ft.legacyLogo = true; }
    if (opts.pageNumber !== undefined) ft.pageNumber = opts.pageNumber;
    if (opts.confidential !== undefined) ft.confidential = opts.confidential;
    if (Object.keys(ft).length) top.footer = ft;
    const raw = {};
    ["font", "sizes", "layout", "cover"].forEach(k => { if (opts[k] !== undefined) raw[k] = opts[k]; });
    const L = Object.keys(raw).length ? normLayer(raw, top.src, top.dir) : top;
    if (top.footer) L.footer = top.footer;
    layers.push(L);
  }
  const TH = { colors: {}, font: {}, sizes: {}, footer: {}, cover: {}, layout: {}, icons: [], master: {}, nav: {} };
  let colorSrc = null;
  const KEY_SRC = {}; // "sizes.body" → どの層から来たか（警告に出す）
  const short = p => String(p).replace(os.homedir(), "~");
  // 会社の色を優先する（2026-11-24）：個人・プロジェクトの theme.json に色があれば、build.js の opts.theme の色は使わない
  //   build.js の色は、多くは AI が題材に合わせて選んだもの。会社の色が決まっているなら、そちらに合わせる
  const fileColors = layers.some(L => !L.fromOpts && L.colors && Object.keys(L.colors).length);
  let skippedOptsColor = false;
  layers.forEach(L => {
    if (L.fromOpts && fileColors && L.colors && Object.keys(L.colors).length) skippedOptsColor = true;
    else if (L.colors && Object.keys(L.colors).length) { Object.assign(TH.colors, L.colors); colorSrc = L.colorSrc; }
    ["font", "sizes", "footer", "cover", "layout", "master", "nav"].forEach(k => { if (L[k]) { Object.assign(TH[k], L[k]); Object.keys(L[k]).forEach(kk => { KEY_SRC[`${k}.${kk}`] = short(L.src); }); } });
    if (L.icons) TH.icons = [...L.icons, ...TH.icons]; // 上の層のフォルダを先に探す
  });
  const fileSrcs = layers.map(L => L.src).filter((s0, i, a) => /\.json$/i.test(s0) && a.indexOf(s0) === i);
  if (fileSrcs.length) console.log(`[deck-parts] テーマ: ${fileSrcs.map(short).join(" ← ")}（右ほど優先）${colorSrc ? `。色は ${short(colorSrc)}` : ""}`);
  if (skippedOptsColor) console.log(`[deck-parts] 会社の色（${short(colorSrc)}）を使い、build.js の theme の色は使わなかった`);
  // 色の組（テーマ）：色のトークン C を差し替える。ink だけを替えたときは、emph と dark も同じ色にする（1色の資料のまま）
  // 色の組は、題材に合わせて決める。名前（PALETTES の見本）・強調の色だけ（{ emph }）・全部の指定、のどれでも渡せる
  if (colorSrc) {
    Object.assign(C, TH.colors);
    TINT = C.tint; TINT_STRONG = C.strong;
    CAT[4] = C.emph; // 主役の区分の色
  }
  else console.warn("[deck-parts] 色の組が既定のまま。題材に合わせて theme を決める（名前で選ぶ: " + Object.keys(PALETTES).join("・") + "。または { emph: \"強調の色\" }）。決めた理由を build.js の先頭に書く");
  // 書体
  if (TH.font.face) F = String(TH.font.face);
  FHEAD = TH.font.heading ? String(TH.font.heading) : F;
  FNUM = TH.font.number ? String(TH.font.number) : F; // 数字の書体は、指定が無ければ本文と同じ
  BIZ_W = isBizP(F);
  {
    const ws = TH.font.widthScale ?? FONT_WIDTH[F] ?? 1;
    if (typeof ws !== "number" || ws < 0.8 || ws > 1.4) themeWarn("font", `font.widthScale は 0.8〜1.4 の数にする（${ws}）。1 を使う`);
    else EMW = ws;
  }
  // 文字の段（下限を下回る値は使わない）
  {
    const sz = {};
    Object.entries(TH.sizes).forEach(([k, v]) => {
      const src = KEY_SRC[`sizes.${k}`] || "sizes";
      if (!(k in SIZE_MIN)) { themeWarn(src, `知らない文字の段（${k}）。使える名前: ${Object.keys(SIZE_MIN).join("・")}`); return; }
      if (typeof v !== "number" || !Number.isFinite(v)) { themeWarn(src, `sizes.${k} が数でない（${v}）。使わない`); return; }
      if (v < SIZE_MIN[k]) { themeWarn(src, `sizes.${k} が ${v}pt で、下限 ${SIZE_MIN[k]}pt を下回る。${SIZE_MIN[k]}pt を使う（映して読める大きさ）`); v = SIZE_MIN[k]; }
      else if (v > SIZE_MAX[k]) themeWarn(src, `sizes.${k} が ${v}pt で大きい（目安 ${SIZE_MAX[k]}pt まで）。部品の箱から文字があふれやすいので、部品の警告と画像で確かめる`);
      sz[k] = v;
    });
    if (sz.body !== undefined && sz.note === undefined) sz.note = sz.body; // 注は本文と同じ大きさ（今の既定と同じ考え方）
    if (sz.title !== undefined) FS.title = sz.title;
    if (sz.lead !== undefined) LEAD_FS = sz.lead;
    if (sz.large !== undefined) FS.lead = sz.large;
    if (sz.name !== undefined) { FS.head = sz.name; FS.block = sz.name; }
    if (sz.body !== undefined) FS.body = sz.body;
    if (sz.note !== undefined) FS.note = sz.note;
    if (sz.focus !== undefined) FOCUS_FS = sz.focus;
    if (Object.keys(sz).length && LEAD_FS >= FS.title + 2) themeWarn("sizes", `リード（${LEAD_FS}pt）が題（${FS.title}pt）より大きい。題 ≧ リード ＞ 本文 の順にする`);
  }
  // 題・リード・章名の位置（動かせる範囲だけ）
  Object.entries(TH.layout).forEach(([k, v]) => {
    const r = LAYOUT_RANGE[k], src = KEY_SRC[`layout.${k}`] || "layout";
    if (!r) { themeWarn(src, `動かせない位置（${k}）。動かせるのは ${Object.keys(LAYOUT_RANGE).join("・")} だけ（本文の上端 y 2.2・スライドの大きさは固定）`); return; }
    if (typeof v !== "number" || v < r[0] || v > r[1]) { themeWarn(src, `layout.${k} は ${r[0]}〜${r[1]} の数にする（${v}）。既定の ${k === "bodyBottom" ? DEFAULT_TOKENS.BOTTOM : k === "margin" ? DEFAULT_TOKENS.M : DEFAULT_TOKENS.Y[k.replace(/Y$/, "")]} を使う`); return; }
    if (k === "margin") { M = v; BW = W - 2 * M; } else if (k !== "bodyBottom") Y[k.replace(/Y$/, "")] = v;
  });
  if (Y.lead - Y.title < 0.5) themeWarn("layout", `リード（${Y.lead}）が題（${Y.title}）に近すぎる。0.5 以上離す`);
  const SH = pres.shapes;
  const FT = TH.footer;
  const FOOTER = FT.style ?? "page";
  if (FOOTER !== "band" && FOOTER !== "page") throw new Error(`createDeck：opts.footer は "page" か "band" にする（${FOOTER}）`);
  const NAV = opts.nav || null;
  const COPY = FT.copyright || "";
  const FOOT_TEXT = FT.text || "";               // フッターの左の文字（page のとき）
  const PAGE_NO = FT.pageNumber !== false;        // ページ番号を出すか
  const CONF = FT.confidential || "";             // 「社外秘」などの表記
  const BAND = FT.bandColor ? String(FT.bandColor).replace(/^#/, "") : null; // 帯の色（既定は強調の色）
  const LOGO = FT.logo && fs.existsSync(FT.logo) ? FT.logo : null;
  if (FT.logo && !LOGO) themeWarn("footer", `ロゴの画像が見つからない（${FT.logo}）`);
  // フッターの帯の高さ：高くした分だけ、本文の下端を上げる（部品は下端に合わせて図の大きさを決める）
  if (FT.height !== undefined) {
    if (FOOTER !== "band") themeWarn("footer", "footer.height は style: \"band\" のときだけ効く。下の飾りを master で描くときは layout.bodyBottom で本文の下端を上げる");
    else if (typeof FT.height !== "number" || FT.height < 0.2 || FT.height > 0.9) themeWarn("footer", `footer.height は 0.2〜0.9 の数にする（${FT.height}）。既定の ${DEFAULT_TOKENS.FH} を使う`);
    else { FH = FT.height; BOTTOM = Math.min(BOTTOM, Math.round((6.905 - FH) * 100) / 100); }
  }
  if (typeof TH.layout.bodyBottom === "number" && TH.layout.bodyBottom >= 5.6 && TH.layout.bodyBottom <= 6.6) BOTTOM = Math.min(BOTTOM, TH.layout.bodyBottom);
  Y.bottom = BOTTOM; Y.footer = H - FH; Y.source = BOTTOM + 0.24;
  if (BOTTOM < DEFAULT_TOKENS.BOTTOM) console.log(`[deck-parts] テーマのフッター・飾りに合わせて、本文の下端を ${DEFAULT_TOKENS.BOTTOM} → ${BOTTOM} にした（図はその分小さくなる）`);
  const logoMaxH = FOOTER === "band" ? (FH <= 0.31 ? 0.2 : Math.min(0.55, FH * 0.62)) : 0.26;
  const LOGO_BOX = LOGO ? (FT.legacyLogo ? { w: 0.8, h: 0.167 } : fitBox(LOGO, FOOTER === "band" ? Math.max(1.6, logoMaxH * 6) : 1.4, logoMaxH)) : null; // opts.logo は今の大きさのまま
  const COVER = TH.cover;
  if (COVER.style && !["auto", "simple", "full"].includes(COVER.style)) themeWarn("cover", `cover.style は auto・simple・full のどれか（${COVER.style}）`);
  if (COVER.logoPos && !["top-left", "top-right", "center"].includes(COVER.logoPos)) themeWarn("cover", `cover.logoPos は top-left・top-right・center のどれか（${COVER.logoPos}）`);
  if (COVER.logo && !fs.existsSync(COVER.logo)) themeWarn("cover", `表紙のロゴの画像が見つからない（${COVER.logo}）`);
  const THEME_ICON_DIRS = TH.icons.filter(d => { if (fs.existsSync(d)) return true; themeWarn("icons", `アイコンのフォルダが見つからない（${d}）`); return false; });

  // ---- 飾りの層（テーマの master）：背景と飾りはスライドマスターに描く。本文の部品の下に敷かれ、PowerPoint でも編集できる ----
  const hexOf = (v, fallback) => {
    if (v === undefined || v === null) return fallback;
    if (typeof v === "string" && v in C) return C[v]; // 色のトークンの名前（emph・ink など）
    const h = String(v).replace(/^#/, "");
    if (/^[0-9A-Fa-f]{6}$/.test(h)) return h.toUpperCase();
    themeWarn("master", `色が6桁の16進でも色の名前でもない（${v}）`); return fallback;
  };
  const MASTER = {};
  Object.entries(TH.master).forEach(([type, p]) => {
    const out = { footer: p.footer, pageNumber: p.pageNumber, nav: p.nav, boxes: {}, objects: [], bgColor: type === "divider" ? C.dark : C.white, bgImage: null };
    if (p.background) {
      if (p.background.color !== undefined) out.bgColor = hexOf(p.background.color, out.bgColor);
      if (p.background.image) { if (fs.existsSync(p.background.image)) out.bgImage = p.background.image; else themeWarn("master", `背景の画像が見つからない（${p.background.image}）`); }
    }
    (p.objects || []).forEach((o, i) => {
      const where = `master.${type}.objects[${i}]`, t = o.type || (o.path ? "image" : o.text !== undefined ? "text" : "rect");
      const ob = { ...o, type: t, x: o.x ?? 0, y: o.y ?? 0 };
      if (t === "image") {
        if (!o.path || !fs.existsSync(o.path)) { themeWarn("master", `${where} の画像が見つからない（${o.path}）`); return; }
        const sz = imageSize(o.path), r = sz && sz.h ? sz.w / sz.h : 1;
        ob.w = o.w ?? (o.h !== undefined ? o.h * r : 1); ob.h = o.h ?? ob.w / r;
      } else if (t === "line") { ob.w = o.w ?? 0; ob.h = o.h ?? 0; }
      else if (t !== "text" && !pres.ShapeType[t]) { themeWarn("master", `${where} の type が使えない（${t}）。rect・roundRect・ellipse・triangle・rtTriangle・parallelogram・trapezoid・line・image・text など`); return; }
      if (![ob.x, ob.y, ob.w, ob.h].every(v => typeof v === "number" && Number.isFinite(v))) { themeWarn("master", `${where} は x・y・w・h を数（インチ）で書く`); return; }
      if (t !== "image") ob.color = hexOf(o.color, t === "text" ? C.ink : C.emph);
      out.objects.push(ob);
    });
    Object.entries(p.boxes || {}).forEach(([name, b]) => { out.boxes[name] = { ...b, ...(b.color !== undefined ? { color: hexOf(b.color, undefined) } : {}) }; });
    out.drawn = !!(out.bgImage || (p.background && p.background.color !== undefined) || out.objects.length);
    MASTER[type] = out;
    if (!out.drawn) return;
    const objects = out.objects.map(o => {
      const pos = { x: o.x, y: o.y, w: o.w, h: o.h }, tr = o.transparency ?? 0;
      if (o.type === "image") return { image: { path: o.path, altText: "飾り", ...pos, ...(tr ? { transparency: tr } : {}), ...(o.rotate ? { rotate: o.rotate } : {}) } };
      if (o.type === "line") return { line: { ...pos, h: o.h || 0, ...(o.flipV ? { flipV: true } : {}), line: { color: o.color, width: o.width ?? 1, ...(o.dash ? { dashType: o.dash } : {}), ...(tr ? { transparency: tr } : {}) } } };
      if (o.type === "text") return { text: { text: String(o.text ?? ""), options: { ...pos, margin: 0, color: o.color, fontSize: o.size ?? 12, bold: !!o.bold, align: o.align || "left", valign: o.valign || "middle", fontFace: o.font || F, ...(o.rotate ? { rotate: o.rotate } : {}), ...(tr ? { transparency: tr } : {}) } } };
      const line = o.line ? { color: hexOf(o.line.color, C.line), width: o.line.width ?? 1 } : { type: "none" };
      return { text: { text: "", options: { shape: pres.ShapeType[o.type], ...pos, fill: o.color && o.fill !== false ? { color: o.color, transparency: tr } : { type: "none" }, line, ...(o.rotate ? { rotate: o.rotate } : {}), ...(o.flipH ? { flipH: true } : {}), ...(o.flipV ? { flipV: true } : {}) } } };
    });
    pres.defineSlideMaster({ title: `DK_${type}`, background: out.bgImage ? { path: out.bgImage } : { color: out.bgColor }, objects });
  });
  // 地の色：(x, y) の位置で、文字の下に何色があるか。背景 → 飾りの順に重ねる
  const SOLID_SHAPES = new Set(["rect", "roundRect", "ellipse", "parallelogram", "trapezoid", "chevron", "homePlate"]);
  const colorAt = (type, x, y) => {
    const m = MASTER[type];
    let c = m ? m.bgColor : type === "divider" ? C.dark : C.white;
    if (!m) return c;
    if (m.bgImage) { const rd = pngReader(m.bgImage); if (rd) { const p = rd(x / W, y / H); c = blend(c, p, p[3]); } }
    m.objects.forEach(o => {
      if (x < o.x || x > o.x + o.w || y < o.y || y > o.y + o.h) return;
      const a = 1 - (o.transparency ?? 0) / 100;
      if (o.type === "image") { const rd = pngReader(o.path); if (rd) { const p = rd((x - o.x) / o.w, (y - o.y) / o.h); c = blend(c, p, p[3] * a); } }
      else if (SOLID_SHAPES.has(o.type) && o.fill !== false) {
        if (o.type === "ellipse" && ((x - o.x - o.w / 2) / (o.w / 2)) ** 2 + ((y - o.y - o.h / 2) / (o.h / 2)) ** 2 > 1) return;
        c = blend(c, o.color, a);
      }
    });
    return c;
  };
  // 矩形の中の9×3点で、文字の色がいちばん読みにくい点の明暗比
  const worstContrast = (type, box, color) => { let w = 21; for (let i = 0; i < 9; i++) for (let j = 0; j < 3; j++) w = Math.min(w, contrast(color, colorAt(type, box.x + box.w * (i + 0.5) / 9, box.y + box.h * (j + 0.5) / 3))); return w; };
  const midY = b => ({ x: b.x, y: b.y + b.h * 0.15, w: b.w, h: b.h * 0.7 });
  // 文字の置き場所：既定 → テーマの boxes の順に重ねる。色を指定していなければ、地が濃いときは白にする
  const boxOf = (type, name, def) => {
    const m = MASTER[type], u = (m && m.boxes[name]) || {};
    const b = { ...def, ...u };
    if (m && u.color === undefined && def.color !== undefined) {
      const shown = unGray(def.color), probe = midY({ ...b, w: Math.min(b.w, 4), x: b.align === "right" ? b.x + b.w - Math.min(b.w, 4) : b.align === "center" ? b.x + (b.w - Math.min(b.w, 4)) / 2 : b.x });
      const now = worstContrast(type, probe, shown);
      const white = worstContrast(type, probe, C.white);
      if ((now < 3 && white > now) || (now < 4.5 && white > now * 1.15)) b.color = C.white;
    }
    return b;
  };
  // 読みにくさの点検（同じ内容は1回だけ）
  const THEME_SEEN = new Set();
  const warnTheme = (key, msg) => { if (THEME_SEEN.has(key)) return; THEME_SEEN.add(key); themeWarn("master", msg); };
  const checkReadable = (type, name, box, color, text, size) => {
    if (!MASTER[type] || !text) return;
    const w = Math.min(box.w, textW(text, size)), x = box.align === "right" ? box.x + box.w - w : box.align === "center" ? box.x + (box.w - w) / 2 : box.x;
    const c = worstContrast(type, midY({ x, y: box.y, w, h: Math.min(box.h, lineH(size, 1.2)) }), unGray(color));
    const need = size >= 20 ? 3 : 4.5;
    if (c < need) warnTheme(`${type}.${name}`, `${type} の ${name} の文字が、地の色・絵の上で読みにくい（明暗比 ${c.toFixed(1)}、目安 ${need} 以上）。boxes.${name}.color を替えるか、飾りをずらす`);
    // 飾りが文字に半分かかる（帯のように文字をすっぽり含むものは除く）
    const tb = { x, y: box.y, w, h: Math.min(box.h, lineH(size, 1.2)) };
    (MASTER[type].objects || []).forEach((o, i) => {
      if ((o.transparency ?? 0) >= 80) return;
      const ox = Math.min(tb.x + tb.w, o.x + (o.w || 0.02)) - Math.max(tb.x, o.x), oy = Math.min(tb.y + tb.h, o.y + (o.h || 0.02)) - Math.max(tb.y, o.y);
      const covers = o.x <= tb.x + 0.01 && o.y <= tb.y + 0.05 && o.x + o.w >= tb.x + tb.w - 0.01 && o.y + o.h >= tb.y + tb.h - 0.05;
      if (ox > 0.05 && oy > 0.08 && !covers) warnTheme(`${type}.${name}.obj${i}`, `${type} の ${name} の文字に、飾り（objects[${i}]）がかかる。飾りか boxes.${name} の位置・幅を替える`);
    });
  };
  // 本文のページの点検：本文の領域は薄い地で、飾りを入れない（図の色と文字は、薄い地を前提にしている）
  if (MASTER.content && MASTER.content.drawn) {
    const m = MASTER.content, area = { x: M, y: BODY, w: BW, h: BOTTOM - BODY };
    const samples = [];
    for (let i = 0; i < 48; i++) for (let j = 0; j < 20; j++) samples.push(colorAt("content", area.x + area.w * (i + 0.5) / 48, area.y + area.h * (j + 0.5) / 20));
    const dark = samples.filter(c => contrast(C.ink, c) < 7 || contrast(C.soft, c) < 4.5).length / samples.length;
    if (dark > 0.5) warnTheme("content.dark", "本文のページの地が濃く、図の文字が読めない。本文のページは白か薄い色にする（濃い地は表紙・章扉・区切りだけ）");
    else if (dark > 0.005) warnTheme("content.busy", `本文の領域（y ${BODY}〜${BOTTOM}）に、濃い絵や模様がかかっている（約 ${Math.round(dark * 100)}%）。薄くするか、領域の外へ寄せる`);
    m.objects.forEach((o, i) => {
      if ((o.transparency ?? 0) >= 80) return;
      const ox = Math.min(area.x + area.w, o.x + o.w) - Math.max(area.x, o.x), oy = Math.min(area.y + area.h, o.y + o.h) - Math.max(area.y, o.y);
      if (ox > 0.05 && oy > 0.05) warnTheme(`content.in${i}`, `本文のページの飾り objects[${i}]（${o.type}）が、本文の領域（x ${M}〜${(W - M).toFixed(2)}、y ${BODY}〜${BOTTOM}）に入る。外へ寄せるか、layout.bodyBottom で本文の下端を上げる`);
    });
    if (!m.bgImage && contrast(C.ink, m.bgColor) < 7) warnTheme("content.bg", `本文のページの地（${m.bgColor}）が濃い。白か薄い色にする`);
  }
  // 前面を薄くする白（echo・overlay）は、本文の地の色に合わせる
  const PAGE_BG = () => (MASTER.content && MASTER.content.drawn && !MASTER.content.bgImage ? MASTER.content.bgColor : C.white);

  let page = 0, chapterNow = 0;
  const SLIDES = [];      // { s, ch, nav, charts }
  const PAGE = {};        // 名前 → ページ番号（目次・参照の「P.n」に使う）
  const LATER = [];       // 全ページを作った後に描くもの（目次のページ番号、参照）
  const CHART_PATCH = []; // 書き出し後に XML を直すグラフ
  const warn = msg => console.warn(`[deck-parts] P.${page} ${msg}`);
  // 改行の点検：1〜2文字だけの行ができそうなら知らせる（語の切れ目で \n を入れるか、言葉を縮める）
  // 2026-11-24 の決まりの記録（形の使い回し・見る所・外部の製品のアイコン）
  const FORM_USE = {};   // 形（部品の名前）→ 最初に使った中身とページ。1つの形は、1つの資料の中で1つの意味にだけ使う
  const FOCUS = [];      // ページごとの「見る所」の宣言。点検で、画像を見て最初に目が行った所と引き比べる
  const BRANDS_ON = {};  // ページ → 図の中に名前が出た外部の製品・会社
  const ICONS_ON = {};   // ページ → 置いたアイコン
  // 外部の製品・会社の名前（アイコンのファイル名 → 名前の形）。opts.brands で足す・替える。false で点検しない
  const BRAND = opts.brands === false ? {} : { slack: /Slack/, excel: /Excel/, powerpoint: /PowerPoint/, teams: /Teams/, notion: /Notion/, claude: /Claude/, openai: /ChatGPT|Codex|OpenAI/, gemini: /Gemini/, grok: /Grok/, cursor: /Cursor/, github: /GitHub/, zoom: /Zoom/, salesforce: /Salesforce/, hubspot: /HubSpot/, jira: /Jira/, aws: /AWS/, dropbox: /Dropbox/, zapier: /Zapier/, ...(opts.brands || {}) };
  let wordHook = null; // 図の中の文の言い回しの点検（WORDY の定義の後で入れる）
  const noteBrands = text => {
    if (wordHook) wordHook(text);
    if (!page) return;
    const t = typeof text === "string" ? text : Array.isArray(text) ? text.map(r => (r && r.text) || "").join("") : "";
    for (const k of Object.keys(BRAND)) if (BRAND[k].test(t)) (BRANDS_ON[page] = BRANDS_ON[page] || {})[k] = 1;
  };
  const checkWrap = (t, pt, w) => { if (hasOrphan(t, pt, w)) warn(`1〜2文字だけの行ができる（「${String(t).replace(/\n/g, "／").slice(0, 16)}…」）`); };

  // ---- 基本の描画 ----
  // 文字：余白0・左上揃えを既定にする（座標がそのまま見た目の位置になる）。書体は Yu Gothic 1つ（5.2）
  // 灰の小さな文（補足）を見つける：14pt 未満・灰か薄い色・8字を超える文。フッター（inFooter）は除く
  let inFooter = false; const FAINT_SEEN = new Set();
  const checkFaint = (text, o) => {
    if (inFooter || !page || !o) return;
    const runs = Array.isArray(text) ? text.map(r => ({ t: (r && r.text) || "", fs: (r && r.options && r.options.fontSize) || o.fontSize, c: (r && r.options && r.options.color) || o.color })) : [{ t: String(text), fs: o.fontSize, c: o.color }];
    if (process.env.FAINT_DIAG) runs.forEach(r => { const g = r.c === C.gray || r.c === C.pale || r.c === C.faint; const sz = r.fs ?? 12; if ((g || sz < 12) && r.c !== C.white && r.c !== "FFFFFF" && r.t.trim()) console.log(`DIAG\tP.${page}\t${sz}\t${g ? "gray" : r.c === C.soft ? "soft" : "ink"}\t${new Error().stack.split("\n")[5].trim().replace(/^at /, "").replace(/ \(.*deck-parts\.js:/, ":").replace(/\)$/, "")}\t${r.t.replace(/\n/g, " ").slice(0, 16)}`); });
    runs.forEach(r => {
      const faint = r.c === C.gray || r.c === C.pale || r.c === C.faint;
      if (faint && (r.fs ?? 19) < minFig() && em(r.t) > 8 && !FAINT_SEEN.has(r.t)) { FAINT_SEEN.add(r.t); warn(`薄く小さい補足の文がある（「${r.t.replace(/\n/g, "").slice(0, 14)}…」）。要るなら読める大きさと色で書き、要らないなら消す`); }
    });
  };
  // 文字の灰・薄い色は、本文の色（soft）で描く（2026-11-24）。灰・薄い色は線と図形の色にだけ使う。フッターは除く
  const GRAYS = () => [C.gray, C.pale, C.faint];
  const unGray = c => (!inFooter && GRAYS().includes(c) ? C.soft : c);
  const unGrayText = text => (Array.isArray(text) ? text.map(r => (r && r.options && r.options.color ? { ...r, options: { ...r.options, color: unGray(r.options.color) } } : r)) : text);
  const BIGF = {}, CHARTP = {};
  const noteBig = (text, o) => { if (inFooter || inChrome || !page || !o) return; const runs = Array.isArray(text) ? text.map(r => ({ fs: (r && r.options && r.options.fontSize) || o.fontSize, c: (r && r.options && r.options.color) || o.color, t: (r && r.text) || "" })) : [{ fs: o.fontSize, c: o.color, t: String(text) }]; if (runs.some(r => (r.fs ?? 0) >= 24 && (r.c === C.emph || r.c === C.accent || r.c === C.white || r.c === "FFFFFF") && String(r.t).trim())) BIGF[page] = (BIGF[page] || 0) + 1; };
  const chartPage = () => { if (page) CHARTP[page] = true; };
  // 見る所の値：36pt（FOCUS_FS）・太字・強調の色（problem なら赤）。グラフの見る所のすぐ隣に置く
  // 見る所の値は、その要素の普通の値の位置に書く（棒なら棒の上、区画なら区画の中）。塗りの上に置くときは onFill で白にする
  // 範囲しか分からない値（「100g未満」「120g以上」「1〜3日」）は、長さ・位置・角度にしない（2026-11-27）
  const isRange = t => /未満|以上|以下|以内|\d[^\d\s]{0,3}\s*[〜～~]\s*\d/.test(String(t ?? ""));
  const rangeBar = (part, texts) => { const t = texts.find(isRange); if (t !== undefined) warn(`${part}：値が範囲（${t}）。範囲しか分からない値は棒にしない（原稿にある1つの値で比べる形に替えるか、文字で言う）`); };
  const focusValue = (s, text, box, o = {}) => T(s, text, { fontSize: FOCUS_FS, bold: true, color: o.onFill ? C.white : o.problem ? C.accent : C.emph, valign: "middle", wrap: false, h: 0.5, ...box });
  // 部品が主役の値を自分で描くときの補助（2026-11-24）。呼ぶ側で値を空にして focusValue を手で置くと、隣の棒・値との重なりを部品が知らないので重なる
  //   focusTH：主役の値の1行の高さ。boxHit：2つの矩形 { x, y, w, h } が重なるか（pad だけ離れていないと重なりとみなす）
  //   liftClear：文字の範囲 r が、ほかの棒・値（obs）に重なる（横は 0.08、縦は pad より近い）なら、dir（-1＝上へ、1＝下へ）へ重ならない所までずらす（r.y を書き換える）。ずらした量を返す
  const focusTH = () => FOCUS_FS * 1.2 / 72;
  const boxHit = (a, b, pad = 0) => a.x < b.x + b.w + pad && b.x < a.x + a.w + pad && a.y < b.y + b.h + pad && b.y < a.y + a.h + pad;
  const liftClear = (r, obs, dir, pad = 0.03) => {
    const y0 = r.y;
    for (let n = 0; n < 60; n++) {
      const hit = obs.find(b => b.w > 0 && b.h > 0 && r.x < b.x + b.w + 0.08 && b.x < r.x + r.w + 0.08 && r.y < b.y + b.h + pad && b.y < r.y + r.h + pad);
      if (!hit) break;
      r.y = dir < 0 ? hit.y - pad - r.h - 0.001 : hit.y + hit.h + pad + 0.001;
    }
    return Math.abs(r.y - y0);
  };
  // 面の中の文字の大きさ：面の高さから決め（行の高さの約 1/1.6）、幅に入るまで下げる。下限 16pt・上限 28pt（構成比の区画、ツリーマップ）
  const areaFs = (text, w, h, lines = 1, lo = FS.body, hi = 28) => { let f = Math.max(lo, Math.min(hi, Math.floor(h * 72 / (lines * 1.6)))); const t = String(text).split("\n").reduce((a, b) => (em(b) > em(a) ? b : a), ""); while (f > lo && textW(t, f) > w - 0.16) f -= 1; return f; };
  // 文字の大きさの下限（2026-11-27 に一段上げた）：図の中の文字は 18pt 以上（本文は 19pt）。16pt は密な図（スイムレーンの工程名）だけ。テーマで本文を下げたときは、その大きさまで。フッターと表紙の外は除く
  let denseFig = false;
  const minFig = () => Math.min(denseFig ? 16 : 18, FS.body);
  const SMALL_SEEN = new Set(); let smallOK = false, inChrome = false; // inChrome：題・リード・ナビなど、図ではない文字を描く間。smallOK：付箋のボード（ブレストの量を優先する例外）を描く間だけ点検しない
  const checkSmall = (text, o) => {
    if (inFooter || inChrome || smallOK || !page || !o) return;
    const runs = Array.isArray(text) ? text.map(r => ({ t: (r && r.text) || "", fs: (r && r.options && r.options.fontSize) || o.fontSize })) : [{ t: String(text), fs: o.fontSize }];
    const r = runs.find(x => x.fs != null && x.fs < minFig() && x.t.trim());
    if (r && !SMALL_SEEN.has(page)) { SMALL_SEEN.add(page); warn(`図の中に ${r.fs}pt の文字がある（「${r.t.replace(/\n/g, "").slice(0, 14)}」）。図の中の文字は 18pt 以上（密な図だけ 16pt）。入らないなら言葉を縮めるか、項目を減らす`); }
  };
  const T = (s, text, o) => (noteBrands(text), checkFaint(text, o), checkSmall(text, o), noteBig(text, o), s.addText(unGrayText(text), { fontFace: F, color: C.ink, margin: 0, isTextBox: true, valign: "top", ...o, ...(o && o.color ? { color: unGray(o.color) } : {}) }));
  // 細線 0.75pt line 色（5.5）。罫線はこれだけ
  const hline = (s, x, y, w, o = {}) => s.addShape(SH.LINE, { x, y, w, h: 0, line: { color: C.line, width: 0.75, ...o } });
  // 2点を結ぶ線（右上がり・右下がりを flipV で表す。負の幅を作らない：9章「負の幅の線でファイル破損」）
  const segLine = (s, x1, y1, x2, y2, o) => s.addShape(SH.LINE, { x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.max(Math.abs(x2 - x1), 0.001), h: Math.max(Math.abs(y2 - y1), 0.001), flipV: (y2 < y1) !== (x2 < x1), line: o });
  // 矩形（枠なしの塗りが既定）
  const rect = (s, x, y, w, h, fill, line) => s.addShape(SH.RECTANGLE, { x, y, w, h, fill: fill ? { color: fill } : { type: "none" }, line: line || { type: "none" } });
  // リンクの文字片（詳細資料へ飛ばす1行の出典に使う。）
  const link = (text, url, o = {}) => ({ text, options: { hyperlink: { url: encodeURI(url), tooltip: url }, color: C.link, underline: { style: "sng" }, ...o } });

  // ---- ページの記録 ----
  let chapterMark = 0; // setChapter の直後のページに、章の名前（ch<番号>）を付ける
  let plainFoot = false; // 表紙・区切り（nav: false のページ）は、フッターの資料名と線を出さない
  const track = (s, o = {}, type = "content") => { plainFoot = o.nav === false; page++; SLIDES.push({ s, type, ch: chapterNow, nav: o.nav !== false, charts: 0 }); if (o.key) PAGE[o.key] = page; if (chapterMark) { if (PAGE[`ch${chapterMark}`] === undefined) PAGE[`ch${chapterMark}`] = page; chapterMark = 0; } return s; };
  // 現在の章（ミニ目次で濃くする章）を切り替える。0＝どれも濃くしない
  const setChapter = n => { chapterNow = n; chapterMark = n; };
  // 今のページに名前を付ける（目次・参照の「P.n」に使う。）
  const mark = key => { PAGE[key] = page; return page; };
  // 全ページを作った後に描く（ページ番号が確定してから描くもの）
  const later = fn => LATER.push(fn);
  let AGENDA_DONE = false; // アジェンダのページを置いたか（finalize で確かめる。2026-11-25）
  // 聞き手の前提（2026-11-25）："first-time"＝その事業・製品を初めて聞く（社外、投資家、新しい提案を初めて聞く役員など）／"insider"＝前提を共有済みの内部の人
  //   初見なら、新しい物は「一言で何か」を中央に出すページ（concept）から入り、仕組み・比べる図は次のページから
  const AUD = opts.audience;
  if (AUD !== undefined && AUD !== "first-time" && AUD !== "insider") console.warn(`[deck-parts] audience は "first-time"（初見）か "insider"（前提を共有済み）にする（${AUD}）`);
  let CONCEPT_DONE = false;
  const needConcept = name => { if (AUD === "first-time" && !CONCEPT_DONE) warn(`初見の聞き手（audience: "first-time"）に、コンセプトのページより先に仕組みの図（${name}）を出している。先に「一言で何か」を中央に出すページ（concept）を置き、仕組みは次のページから見せる`); };
  // 参照の「P.n」を後で描く（離れた場所の対応を番号で読ませるためではなく、詳細の所在を示すため。10.13）
  const ref = (s, key, o) => { const { prefix = "", ...pos } = o; later(() => T(s, `${prefix}P.${PAGE[key] ?? "?"}`, { fontSize: FS.note, color: C.gray, align: "right", valign: "middle", wrap: false, h: 0.26, ...pos })); };

  // ---- 骨格 ----
  // フッター：既定はページ番号だけを右下に（テーマの footer で替わる）。opts.footer が "band" なら、濃紺の帯に、左に表記・中央にページ番号・右にロゴ
  const footer = (s, type = "content") => { if (MASTER[type] && MASTER[type].footer === false) return; inFooter = true; try { footer0(s, type); } finally { inFooter = false; } };
  // フッターの文字の色：footer.color。指定がなくて地が濃ければ白
  const FOOT_COLOR = FT.color !== undefined ? hexOf(FT.color, undefined) : undefined;
  const footColor = (type, box, def) => FOOT_COLOR ?? (MASTER[type] && worstContrast(type, box, def) < 3 && worstContrast(type, box, C.white) > worstContrast(type, box, def) ? C.white : def);
  // ページ番号を好きな場所に描く（footer.pageNumber が { x, y, w, h, align, color, size, bold, format } のとき）。format の {n} は番号、{total} は総ページ数
  const pnOf = type => (MASTER[type] && MASTER[type].pageNumber !== undefined ? MASTER[type].pageNumber : FT.pageNumber);
  const pageNumberBox = (s, type, PNB) => {
    const n = page, size = PNB.size ?? Math.min(FS.note, FOOT_FS), box = { x: PNB.x ?? W - M - 1.2, y: PNB.y ?? H - 0.5, w: PNB.w ?? 1.2, h: PNB.h ?? 0.3 };
    const color = PNB.color !== undefined ? hexOf(PNB.color, C.soft) : footColor(type, box, C.soft), fmt = String(PNB.format ?? "{n}");
    const draw = () => T(s, fmt.replace(/\{n\}/g, String(n)).replace(/\{total\}/g, String(page)), { ...box, fontSize: size, bold: !!PNB.bold, color, align: PNB.align || "right", valign: "middle", wrap: false });
    if (fmt.includes("{total}")) later(() => { const k = inFooter; inFooter = true; try { draw(); } finally { inFooter = k; } });
    else draw();
  };
  const footer0 = (s, type) => {
    const pn = pnOf(type), PNB = pn && typeof pn === "object" ? pn : null;
    if (PNB) pageNumberBox(s, type, PNB);
    const NUM_ON = pn !== false && !PNB;
    if (FOOTER === "page") {
      // フッターの文字（opts.footerText・テーマの footer.text）があれば、細い線の下の左に資料名、右にページ番号（本文の下端 6.6 との間に余白を残す。線は y 6.98）
      let lined = !!FOOT_TEXT && !plainFoot;
      // 会社の背景・飾りがある地では、資料名を置く所に絵やロゴがあれば出さない
      const MF = MASTER[type];
      if (lined && MF && MF.drawn) {
        const x0 = M + (LOGO ? LOGO_BOX.w + 0.2 : 0), w = Math.min(textW(FOOT_TEXT, Math.min(FS.note, FOOT_FS)), 5);
        const at = [];
        for (let i = 0; i < 40; i++) for (let j = 0; j < 5; j++) at.push(colorAt(type, x0 + w * (i + 0.5) / 40, H - 0.44 + 0.26 * (j + 0.5) / 5));
        if (at.some(c => contrast(c, at[0]) > 1.15) || contrast(footColor(type, { x: x0, y: H - 0.44, w, h: 0.26 }, C.gray), at[0]) < 2) { lined = false; warnTheme(`footer.${type}`, `${type} の下の左に絵やロゴがあるので、資料名（footerText）は出さない。出すなら飾りをずらすか、footer.pageNumber と master の text で置き場所を決める`); }
      }
      const rowY = lined ? H - 0.44 : H - 0.5;
      const col = def => footColor(type, { x: M, y: rowY, w: BW, h: 0.26 }, def);
      // テーマのロゴ（page のとき）：左下。表紙・区切りのページには出さない
      let tx = M;
      if (LOGO && !plainFoot) { s.addImage({ path: LOGO, altText: "ロゴ", x: M, y: rowY + 0.13 - LOGO_BOX.h / 2, w: LOGO_BOX.w, h: LOGO_BOX.h }); tx = M + LOGO_BOX.w + 0.2; }
      if (lined) { if (!MASTER[type] || !MASTER[type].drawn) hline(s, M, H - 0.52, BW); T(s, FOOT_TEXT, { x: tx, y: H - 0.44, w: BW - 1.2 - (tx - M), h: 0.26, fontSize: Math.min(FS.note, FOOT_FS), color: col(C.gray), valign: "middle", wrap: false }); }
      if (NUM_ON) T(s, String(page), { x: W - M - 1, y: rowY, w: 1, h: 0.26, fontSize: Math.min(FS.note, FOOT_FS), color: col(C.gray), align: "right", valign: "middle", wrap: false });
      if (CONF) { const cw = textW(CONF, Math.min(FS.note, FOOT_FS)) + 0.1, right = W - M - (NUM_ON ? textW(String(page), Math.min(FS.note, FOOT_FS)) + 0.3 : 0); T(s, CONF, { x: right - cw, y: rowY, w: cw, h: 0.26, fontSize: Math.min(FS.note, FOOT_FS), bold: true, color: col(C.ink), align: "right", valign: "middle", wrap: false }); }
      return;
    }
    const band = BAND ? hexOf(BAND, C.emph) : C.emph;
    // 帯の上の文字：帯が明るい色なら文字の色（ink）、暗い色なら白（明るい帯に白い文字は読めない）
    const onBand = FOOT_COLOR ?? (relLum(band) > 0.35 ? C.ink : C.white);
    rect(s, 0, H - FH, W, FH, band);
    if (COPY) T(s, COPY, { x: 0.28, y: H - FH, w: 5, h: FH, fontSize: Math.min(FS.note, FOOT_FS), color: onBand, valign: "middle", wrap: false });
    if (NUM_ON) T(s, String(page), { x: W / 2 - 0.5, y: H - FH, w: 1, h: FH, fontSize: 11, color: onBand, align: "center", valign: "middle" });
    if (LOGO) s.addImage({ path: LOGO, altText: "ロゴ", x: W - 0.28 - LOGO_BOX.w, y: H - FH + (FH - LOGO_BOX.h) / 2, w: LOGO_BOX.w, h: LOGO_BOX.h });
    if (CONF) { const cw = textW(CONF, Math.min(FS.note, FOOT_FS)) + 0.1, right = W - 0.28 - (LOGO ? LOGO_BOX.w + 0.25 : 0); T(s, CONF, { x: right - cw, y: H - FH, w: cw, h: FH, fontSize: Math.min(FS.note, FOOT_FS), bold: true, color: onBand, align: "right", valign: "middle", wrap: false }); }
  };

  // タイトルの点検：タイトルは、いまどの話かの項目名だけ（長い命題文・文末の「。」・「：」でつないだ要点は、リードへ移す）（2.1、既定4）
  // 題の形：既定は「いまどの話か」の項目名（主張はリード）。titleStyle: "action" は、結論の1文を題に書き、リードを置かない（送って読ませる報告書）
  const ACTION = opts.titleStyle === "action";
  if (ACTION) warn("題は項目名にする。結論の1文はリードに書く（titleStyle: \"action\" は使わない。報告書も同じ）。題に結論を書かない");
  // 補足の sub：使うと知らせる（部品ごとに1回）。要る中身は名前に書き、要らないなら消す
  const SUB_WARNED = new Set();
  const warnSub = where => { if (SUB_WARNED.has(where + page)) return; SUB_WARNED.add(where + page); warn(`${where}に補足の小さな文字（sub）がある。補足は足さない。要る中身は名前に書く（単位は「（万円）」のように名前の後ろに）`); };
  const checkTitle = (t, w = BW, fs0 = FS.title) => {
    if (ACTION) { if (lineCount(t, 20, BW) > 2) warn(`題が2行を超える（${t}）。結論を1文に縮める`); return; }
    if (textW(t, fs0) > w) warn(`タイトルが1行に収まらない（${t}）。冗長な語尾を削る`);
    if (/。$/.test(t)) warn(`タイトルが文になっている（${t}）。結論の文はリードに置く`);
    // タイトルは「いまどの話か」の項目名だけ。主張はリードに書く（同じことを、タイトルとリードで2回言わない）
    if (/[：:、]/.test(t)) warn(`タイトルに要点が入っている（${t}）。タイトルは項目名だけにし、主張はリードに書く`);
    else if (em(t) > 14) warn(`タイトルが長い（${t}）。項目名は全角14字まで`);
  };

  // 新しいページ：テーマに飾りがある種類は、そのスライドマスターを使う。無ければ今まで通り（白。区切りは濃い地）
  const newSlide = type => {
    if (MASTER[type] && MASTER[type].drawn) return pres.addSlide({ masterName: `DK_${type}` });
    const s = pres.addSlide();
    s.background = { color: type === "divider" ? C.dark : C.white };
    return s;
  };

  // 標準ページ：タイトル・リード・フッター（2.1）。lead は結論（1ページ1メッセージ、2行・70字まで）。o.key＝ページの名前、o.nav:false でミニ目次を外す
  // リードの言い回しの点検：冗長な文を、短く言い切る文に直す
  //   熟語で縮められる言い回し（[見つける形, 直し方]）。opts.wordy で足す・替える。false で点検しない
  const WORDY = opts.wordy === false ? [] : [
    [/自分のお金で/, "自費で"], [/することができ/, "でき"], [/ことができ/, "でき（「〜できる」）"], [/を行[うっい]/, "「〜する」"], [/という(こと|もの|形)/, "「という」を削る"],
    [/たくさん/, "多数・多く"], [/もう一度/, "再度"], [/そのまま使/, "流用"], [/使っている人/, "利用者・使う人"], [/していくこと/, "「〜すること」"],
    [/のうちに/, "以内に・で"], [/確かめ[るたて]/, "検証する・確認する"], [/お金/, "料金・費用・資金（「お金をとらない」は「無料」）"], [/一緒に/, "共同で"], [/いちばん/, "最も・最大の"], [/(できる|する)ようにする/, "「〜可能にする」「〜する」"], [/になっている/, "「〜だ」「〜である」か言い切り"], [/必要(が|と)(ある|なる)/, "「〜が要る」「必須」"], [/可能(です|である|だ)/, "「〜できる」"],
    ...(opts.wordy || []),
  ];
  const OPENERS = /^(対象|ターゲット|課題|問題|解決策|目的|概要|結論|ポイント|ここで|このページで)は、?/;
  // リードは結論を言い切る文（2026-11-27）：予告、ぼかし、数字にできる言葉、体言止めを知らせる。長さは全体で LEAD_MAX 字まで（上限であって目標ではない）
  const LEAD_MAX = 70;
  const PREVIEW = /について(ご)?(説明|紹介|解説|述べ|お話)|を(ご)?(説明|紹介|解説)(します|いたします)|見ていき|見ていく|見てみ(ます|ましょう)/;
  const HEDGE = /と思われ|かもしれ|可能性があ|と考えられ|ではないか|ではないでしょうか/;
  const VAGUE = /大幅|多くの|一定の|さまざま|様々/;
  // 比喩・抽象語（例。ここに無い比喩は、聴講者の疑問の欄で見つける）。「必要」「重要」「要る」「要否」には当てない
  const METAPHOR = /足場|土台|武器|勝ち筋|突破口|追い風|向かい風|伸びしろ|起爆剤|切り札|カギ(?!括弧|かっこ)|鍵を握|(成功|成長|勝負|勝ち|解決|普及|拡大|定着|差別化|改善|競争|攻略)の鍵|が鍵(?=だ|です|となる|になる|。)|かなめ(?=だ|です|となる|になる|。)|[のが]要(?=だ|です|となる|になる|。)/;
  // 否定で結ぶ文：何が言えるかが、聴講者に残らない
  const NEGATIVE = /(ではない|でない|じゃない|でもない|ではなかった|ではありません|でもありません|とは(言|い)え(ない|ません)|とは限(らない|りません))$/;
  // 数字の取り出し：全角と半角、桁区切りのカンマ、末尾の .0 をそろえる
  const numsOf = t => (String(t).normalize("NFKC").replace(/(\d),(?=\d{3})/g, "$1").match(/\d+(\.\d+)?/g) || []).map(v => String(Number(v)));
  // 図に描いた文字（リードの数字と照合する）：addText・addTable・addChart の引数から取り出す
  const figTexts = (method, a) => {
    const join = t => Array.isArray(t) ? t.map(r => (r && typeof r === "object" ? join(r.text ?? "") + (r.options && r.options.breakLine ? "\n" : "") : String(r ?? ""))).join("") : String(t ?? "");
    if (method === "addText") return [join(a[0])];
    if (method === "addTable") return (a[0] || []).flatMap(row => (row || []).map(c => join(c && typeof c === "object" && !Array.isArray(c) ? c.text ?? "" : c)));
    const data = Array.isArray(a[0]) ? a[0].flatMap(x => (x && x.data) || []) : a[1] || [];
    return data.flatMap(d => [d.name, ...(d.labels || []).flat(), ...(d.values || []).flatMap(v => (Math.abs(v) <= 1 ? [v, Math.round(v * 1000) / 10] : [v]))]).map(v => String(v ?? ""));
  };
  const LEADS = []; // { page, title, lead, ask, fig }。一覧（.leads.md）と、数字の照合に使う
  // 図の中の文・メッセージボックス・表の文：リードと同じ一覧で、縮められる言い回しを知らせる（同じ文は1回だけ。題とリードは checkLead が見る）
  const HEAD_TEXT = new Set(), WORD_SEEN = new Set();
  wordHook = text => {
    const t = typeof text === "string" ? text : Array.isArray(text) ? text.map(r => (r && r.text) || "").join("") : "";
    if (!t || !page || HEAD_TEXT.has(t)) return;
    WORDY.forEach(([re, to]) => {
      const w = t.match(re); if (!w) return;
      const k = `${w[0]}|${t}`; if (WORD_SEEN.has(k)) return;
      WORD_SEEN.add(k); warn(`図の中の言い回しを縮める（「${w[0]}」→ ${to}）：「${t.replace(/\n/g, "").slice(0, 18)}…」`);
    });
  };
  const checkLead = (title, lead) => {
    const t = String(lead);
    const m = t.match(OPENERS);
    if (m) warn(`リードの書き出し（${m[0]}）は、タイトルで分かるので要らない。中身から書き出す`);
    else if (title && t.startsWith(String(title))) warn(`リードが、タイトルと同じ言葉（${title}）で始まっている。タイトルで分かることを言い直さない`);
    WORDY.forEach(([re, to]) => { const w = t.match(re); if (w) warn(`リードの言い回しを縮める（「${w[0]}」→ ${to}）`); });
    const flat = t.replace(/\n/g, "").trim(), len = [...flat].length;
    if (len > LEAD_MAX) warn(`リードが長い（${len}字。2行・${LEAD_MAX}字まで）：「${flat.slice(0, 18)}…」。言うことを1つに絞る（短く言えるなら、短いほどよい）`);
    if (!flat.endsWith("。")) warn(`リードが「。」で終わっていない（「…${flat.slice(-12)}」）。結論を文で言い切る`);
    flat.split("。").forEach(x => {
      if (NEGATIVE.test(x.replace(/（[^（）]*）|\([^()]*\)/g, "").replace(/[」』】）)"”]+$/, ""))) warn(`リードが否定で終わる（「…${x.slice(-14)}」）。何が言えるかを肯定で書く（例：高すぎる値段ではない → 450円の価格受容性は裏づけられた）`);
    });
    flat.split("。").slice(0, -1).forEach(x => {
      const tail = x.replace(/（[^（）]*）|\([^()]*\)/g, "").replace(/[」』】）)"”]+$/, "").slice(-1);
      if (/[\p{Script=Han}\p{Script=Katakana}々ー]/u.test(tail)) warn(`リードが体言止め（「…${x.slice(-12)}。」）。「〜だ」「〜する」まで書いて文にする（図の中の札は体言止めでよい）`);
    });
    let w;
    if ((w = flat.match(PREVIEW))) warn(`リードが予告になっている（「${w[0]}」）。何を話すかではなく、結論を書く`);
    if ((w = flat.match(HEDGE))) warn(`リードがぼかしている（「${w[0]}」）。言い切る（言い切れないなら、何が分かれば言えるかを書く）`);
    if ((w = flat.match(VAGUE))) warn(`リードの「${w[0]}」は、数字で書く`);
    if ((w = flat.match(METAPHOR))) warn(`リードの「${w[0]}」は比喩・抽象語。何を・どうするで書く（例：足場が要る → 個食への参入が必要だ）`);
  };
  const base = (title, lead, o = {}) => {
    const navNow = NAV && chapterNow ? NAV.find(n => n.ch === chapterNow || (n.subs || []).includes(chapterNow)) : null;
    if (title == null && navNow && !ACTION) title = navNow.name; // 章立て（nav）を1か所に書けば、タイトルはそこから取る
    // 題の上に小さな章名は出さない（2026-11-27）。今どの話かは、アジェンダと題で分かる
    if (o.section) warn(`題の上の章名（section: ${o.section}）は出さない。今どの話かは、アジェンダと題で分かる。base の section を消す`);
    const s = newSlide("content");
    track(s, o);
    FOCUS.push({ page, title, focus: o.focus });
    const fig = [];
    LEADS.push({ page, title, lead: lead ? String(lead) : "", ask: [].concat(o.ask || []).join(" ／ "), fig });
    // 題・リード・フッター・ナビを除いた、図の文字を控える
    ["addText", "addTable", "addChart"].forEach(m => { const f = s[m].bind(s); s[m] = (...a) => { if (!inChrome && !inFooter) fig.push(...figTexts(m, a)); return f(...a); }; });
    if (title && o.focus === undefined) warn("見る所の宣言がない（base の o.focus に、持ち帰る1つ（最初に目が行くべき要素）を書く。false は、リードが主張を言わないページだけ）");
    if (o.focus === false && lead && !/付録|参考|一覧|目次|アジェンダ/.test(String(title || ""))) warn("主張を言うリードがあるのに、見る所がない（focus: false）。持ち帰る1つ（語か数値）を決めて、それがいちばん大きく見える形にする");
    if (title) HEAD_TEXT.add(String(title));
    if (o.sub) HEAD_TEXT.add(String(o.sub));
    // 文字の置き場所（テーマの master.content.boxes で動かせる）
    const tb = boxOf("content", "title", { x: M, y: Y.title, w: BW, h: 0.6, size: FS.title, color: C.ink, bold: true, font: FHEAD });
    const lb = boxOf("content", "lead", { x: M, y: Y.lead - 0.04, w: BW, h: 0.8, size: LEAD_FS, color: C.ink, bold: true });
    if (tb.y + Math.min(tb.h, 0.6) > lb.y + 0.25 && title && lead) warnTheme("content.overlap", "題の置き場所が、リードと重なる。boxes.title と boxes.lead の y を離す");
    if (lb.y + lb.h > BODY + 0.12) warnTheme("content.leadLow", `リードの置き場所の下端が、本文の領域（y ${BODY}）に入る`);
    const boxText = (b, color) => ({ x: b.x, y: b.y, w: b.w, h: b.h, fontSize: b.size, color: color ?? b.color, ...(b.bold ? { bold: true } : {}), ...(b.align ? { align: b.align } : {}), ...(b.valign ? { valign: b.valign } : {}), ...(b.font && b.font !== F ? { fontFace: b.font } : {}) });
    if (lead) { HEAD_TEXT.add(leadBreak(lead, lb.size, lb.w)); checkLead(title, lead); }
    if (lead && (String(lead).match(/。/g) || []).length >= 3) warn("リードが3文以上ある。主張は1つにし、残りは同じ図を敷いた次のページに分ける（series）");
    inChrome = true;
    if (title) {
      checkTitle(title, tb.w, tb.size);
      checkReadable("content", "title", tb, tb.color, title, tb.size);
      if (ACTION) T(s, title, { ...boxText(tb), y: tb.y - 0.12, h: 0.95, fontSize: 20, valign: "middle", lineSpacingMultiple: 1.2 });
      else T(s, title, boxText(tb));
      // 題に副題を付けない（2026-11-24。題は項目名だけ・薄い補足の文字を足さない）。同じ章で別の図が続くときも題は同じにし、何のページかはリードで言う
      if (o.sub && !ACTION) warn(`題に副題を付けない（${o.sub}）。題はアジェンダの項目名だけにし、何のページかはリードで言う`);
    }
    if (ACTION && lead) warn("結論を題に書く資料（titleStyle: \"action\"）では、リードを置かない（主張を書く場所は1か所）");
    if (lead) {
      const t = leadBreak(lead, lb.size, lb.w);
      const n = lineCount(t, lb.size, lb.w * 0.99);
      if (n > 2) warn(`リードが${n}行になる（2行まで）。図を動かさず文を縮める（2.1）`);
      checkReadable("content", "lead", lb, lb.color, t.split("\n")[0], lb.size);
      T(s, t, { ...boxText(lb), lineSpacingMultiple: 1.2 });
    }
    inChrome = false;
    BRANDS_ON[page] = {}; // 題とリードに出た製品名は数えない（アイコンが要るのは、図の中の要素として出したとき）
    footer(s);
    return s;
  };

  // 章扉・区切りの文字を、テーマの置き場所に書く
  const placed = (s, type, name, text, def, extra = {}) => {
    if (text === undefined || text === null || text === "") return;
    const b = boxOf(type, name, def);
    checkReadable(type, name, b, b.color, String(text), b.size);
    inChrome = true;
    T(s, text, { x: b.x, y: b.y, w: b.w, h: b.h, fontSize: b.size, color: b.color, ...(b.bold ? { bold: true } : {}), ...(b.align ? { align: b.align } : {}), ...(b.valign ? { valign: b.valign } : {}), ...(b.font && b.font !== F ? { fontFace: b.font } : {}), ...extra });
    inChrome = false;
  };

  // 表紙：既定はタイトルと日付だけ（既定7）。経営会議用の報告書だけ kicker（上の小見出し）・sub（副題）・affil（所属）を足す（4.12）
  // テーマの cover：style（auto＝今の決め方／simple＝題と日付／full＝小見出し・題・短い線・副題・日付と所属）、affil（所属の既定）、logo と logoPos（ロゴの位置）
  const cover = (o) => {
    const s = newSlide("cover");
    track(s, { nav: false, key: o.key }, "cover");
    const affil = o.affil ?? COVER.affil;
    const full = COVER.style === "full" ? true : COVER.style === "simple" ? false : (o.kicker || o.sub || affil);
    const custom = !!(MASTER.cover && Object.keys(MASTER.cover.boxes).length); // テーマで表紙の置き場所を決めている
    const align = (custom && MASTER.cover.boxes.title && MASTER.cover.boxes.title.align) || "center";
    const put = (name, text, def, extra = {}) => {
      if (!text) return;
      const b = boxOf("cover", name, { align, color: C.ink, valign: "middle", ...def });
      if (custom) used.forEach(u => {
        const h1 = Math.min(b.h, lineH(b.size, 1.1)), y1 = b.valign === "top" ? b.y : b.y + (b.h - h1) / 2;
        if (Math.min(y1 + h1, u.y + u.h) - Math.max(y1, u.y) > 0.05 && Math.min(b.x + b.w, u.x + u.w) - Math.max(b.x, u.x) > 0.1) warnTheme(`cover.${name}.${u.name}`, `表紙の ${name} と ${u.name} が重なる。boxes.${name} か boxes.${u.name} の y を離す`);
      });
      { const h1 = Math.min(b.h, lineH(b.size, 1.1)); used.push({ name, x: b.x, w: b.w, h: h1, y: b.valign === "top" ? b.y : b.y + (b.h - h1) / 2 }); }
      checkReadable("cover", name, b, b.color, text, b.size);
      inChrome = true;
      T(s, text, { x: b.x, y: b.y, w: b.w, h: b.h, fontSize: b.size, color: b.color, align: b.align, valign: b.valign, ...(b.bold ? { bold: true } : {}), ...(b.font && b.font !== F ? { fontFace: b.font } : {}), ...extra });
      inChrome = false;
    };
    // 題が置き場所に1行で収まらないときは、2行分の高さがあれば2行、無ければ下げて1行に保つ
    const tdef = boxOf("cover", "title", { x: 0, y: full ? 2.6 : 2.75, w: W, h: 1.0, size: FS_DISPLAY.cover });
    let coverFs = tdef.size;
    {
      const wMax = Math.min(tdef.w, BW) * 0.98, was = coverFs;
      if (!custom) {
        if (textW(o.title, coverFs) > wMax) coverFs = Math.round(coverFs * 40 / 48);
      } else if (!String(o.title).includes("\n") && textW(o.title, coverFs) > wMax) {
        const one = Math.floor(wMax * 72 / em(o.title));
        const twoOk = tdef.h >= 2 * lineH(coverFs, 1.0) && textW(o.title, coverFs) <= wMax * 1.8 && !hasOrphan(o.title, coverFs, wMax);
        if (one >= coverFs * 0.7) coverFs = one;
        else if (!twoOk) coverFs = Math.max(Math.round(coverFs * 0.7), Math.floor(wMax * 1.7 * 72 / em(o.title)));
      }
      if (coverFs !== was) warn(`表紙の題が ${was}pt で1行に収まらないため ${coverFs}pt にした（${o.title}）`);
    }
    let coverTitle = o.title;
    if (custom && !String(o.title).includes("\n") && textW(o.title, coverFs) > Math.min(tdef.w, BW) * 0.98) {
      const t = String(o.title), wMax = Math.min(tdef.w, BW) * 0.98;
      let best = null;
      for (let i = 2; i < t.length - 1; i++) {
        if (!"のとをにへでは、。・：　 ".includes(t[i - 1])) continue;
        const a = t.slice(0, i).trim(), b = t.slice(i).trim(), longer = Math.max(textW(a, coverFs), textW(b, coverFs));
        if (longer <= wMax && (!best || longer < best.longer)) best = { text: `${a}\n${b}`, longer };
      }
      if (best) coverTitle = best.text;
    }
    const left = custom ? { x: tdef.x, w: tdef.w } : { x: 0, w: W };
    // 置き場所を指定した表紙は、指定のない行を、題の置き場所の上下に順に並べる
    const used = [], own = (MASTER.cover && MASTER.cover.boxes) || {};
    const under = tdef.y + tdef.h;
    const noSubSlot = custom && !!o.sub && !own.sub && own.date?.y !== undefined; // 副題の置き場所が無く、日付の置き場所はある
    const tight = noSubSlot && own.date.y + 0.9 > H - 0.05; // 日付の置き場所が下端に近い
    const subY = own.sub?.y ?? (noSubSlot ? own.date.y - (tight ? 0.2 : 0) : custom ? under + 0.1 : 4.1);
    const dateY = def => (custom ? (full && o.sub ? subY + (own.sub?.h ?? 0.5) + 0.15 : under + 0.15) : def);
    if (!full) {
      put("title", coverTitle, { x: 0, y: 2.75, w: W, h: 1.0, size: coverFs, bold: true, font: FHEAD }, { fontSize: coverFs });
      const dl = [o.date, affil].filter(Boolean).join("　"); // simple でも所属があれば日付の横に
      put("date", dl, { ...left, y: dateY(4.05), h: 0.4, size: FS.lead, color: C.soft });
    } else {
      put("kicker", o.kicker, { ...left, y: custom ? tdef.y - 0.5 : 2.05, h: 0.4, size: 14, bold: true, valign: "top" }, { charSpacing: 4 });
      put("title", coverTitle, { x: 0, y: 2.6, w: W, h: 1.0, size: coverFs, bold: true, font: FHEAD, valign: "top" }, { fontSize: coverFs });
      if (!custom) s.addShape(SH.LINE, { x: W / 2 - 0.6, y: 3.85, w: 1.2, h: 0, line: { color: "E6F0FA", width: 2.25 } });
      put("sub", o.sub, { ...left, ...(noSubSlot ? { x: own.date.x, w: own.date.w, align: own.date.align || align } : {}), y: subY, h: tight ? 0.34 : 0.5, size: tight ? 16 : 18, color: MASTER.cover ? C.soft : "666666", valign: "top" });
      const ownAffil = custom && MASTER.cover.boxes.affil;
      const ownDate = own.date;
      if (noSubSlot) MASTER.cover.boxes.date = { ...ownDate, ...(tight ? { y: subY + 0.4, h: 0.3, size: 12, valign: "top" } : { y: ownDate.y + 0.55 }) };
      put("date", [o.date, ownAffil ? null : affil].filter(Boolean).join("　"), { ...left, y: dateY(5.1), h: 0.4, size: 14, valign: "top" });
      if (noSubSlot) MASTER.cover.boxes.date = ownDate; // 同じ設定で2枚目の表紙を作るときのために戻す
      if (ownAffil) put("affil", affil, { ...left, y: dateY(5.1) + 0.45, h: 0.4, size: 14, valign: "top" });
    }
    if (COVER.logo && fs.existsSync(COVER.logo)) {
      const lbx = MASTER.cover && MASTER.cover.boxes.logo;
      if (lbx) {
        const b = fitBox(COVER.logo, lbx.w ?? 2.4, lbx.h ?? 0.6), bw = lbx.w ?? b.w;
        const x = (lbx.x ?? M) + (lbx.align === "right" ? bw - b.w : lbx.align === "center" ? (bw - b.w) / 2 : 0);
        s.addImage({ path: COVER.logo, altText: "ロゴ", x, y: lbx.y ?? 0.5, w: b.w, h: b.h });
      } else {
        const pos = COVER.logoPos || "top-left", b = fitBox(COVER.logo, 2.4, 0.6);
        const x = pos === "top-right" ? W - M - b.w : pos === "center" ? (W - b.w) / 2 : M;
        const y = pos === "center" ? (full ? (o.kicker ? 1.95 : 2.45) : 2.6) - b.h - 0.1 : 0.5;
        s.addImage({ path: COVER.logo, altText: "ロゴ", x, y, w: b.w, h: b.h });
      }
    }
    footer(s, "cover");
    return s;
  };

  // 章扉：番号と章題だけ（説明・副題は既定で書かない。4.12、10.18）。num の数字を章番号として現在の章に切り替える
  const chapter = (num, title, o = {}) => {
    const n = o.ch ?? parseInt(num, 10);
    if (!Number.isNaN(n)) chapterNow = n;
    const s = newSlide("chapter");
    track(s, { key: o.key || `ch${String(num).replace(/\.$/, "")}`, nav: o.nav }, "chapter");
    // 題の置き場所だけを指定したテーマでは、番号は題の上（上に空きが無ければ下）に置く
    const cb = (MASTER.chapter && MASTER.chapter.boxes) || {}, ct = cb.title;
    const numAt = ct && ct.y !== undefined && !cb.num
      ? { x: ct.x ?? M + 0.4, w: ct.w ?? 3, align: ct.align, y: ct.y >= 1.2 ? ct.y - 0.55 : ct.y + (ct.h ?? 0.9) + 0.7 }
      : { x: M + 0.4, y: 2.75, w: 3 };
    placed(s, "chapter", "num", num, { ...numAt, h: 0.5, size: FS_DISPLAY.chapterNum, color: C.gray });
    placed(s, "chapter", "title", title, { x: M + 0.4, y: 3.25, w: 12.0, h: 0.9, size: em(title) > 26 ? FS_DISPLAY.chapterLong : FS_DISPLAY.chapter, color: C.ink, bold: true, font: FHEAD });
    placed(s, "chapter", "sub", o.sub, { x: M + 0.4, y: 4.25, w: 12.0, h: 0.8, size: 18, color: C.gray }, { lineSpacingMultiple: 1.3 });
    footer(s, "chapter");
    return s;
  };

  // 区切りページ：紺地に白の題（一覧表示でも切り替わりが分かるように塗る。4.12）。濃紺の全面塗りはここだけ（10.2）
  const divider = (title, sub, o = {}) => {
    if (o.ch !== undefined) chapterNow = o.ch;
    const s = newSlide("divider");
    track(s, { key: o.key, nav: false }, "divider");
    // 区切りの地を明るくしたテーマでは、文字を濃い色にする
    const light = MASTER.divider && MASTER.divider.drawn && contrast(C.white, colorAt("divider", M + 2, 3.45)) < 3;
    placed(s, "divider", "title", title, { x: M + 0.4, y: 3.0, w: 12.0, h: 0.9, size: FS_DISPLAY.chapter, color: light ? C.ink : C.white, bold: true, font: FHEAD });
    placed(s, "divider", "sub", sub, { x: M + 0.4, y: 4.0, w: 12.0, h: 0.9, size: FS.lead, color: light ? C.soft : "C9D3E0" }, { lineSpacingMultiple: 1.3 });
    footer(s, "divider");
    return s;
  };

  // 目次：章（大項目）だけを載せ、ページ番号は最後に確定して描く（4.12 drawAgenda）
  //   cols: [{ no, head, ref?, note?, items: [{ no, title, key }] }]（最大2列。ref＝参考の列は線・色を一段弱く）
  const drawAgenda = (s, cols, o = {}) => {
    const two = cols.length > 1, xs = two ? [M + 0.1, 7.35] : [M + 0.1], ws = two ? [6.4, W - M - 7.35] : [BW - 0.2];
    cols.forEach((col, ci) => {
      const x = xs[ci], w = ws[ci], ref = !!col.ref;
      let y = o.y0 ?? 1.9;
      T(s, [{ text: `${col.no}　`, options: { fontSize: FS_DISPLAY.agendaNo, bold: true, color: ref ? C.gray : C.ink } }, { text: col.head, options: { fontSize: FS_DISPLAY.agendaHead, bold: true, color: ref ? C.soft : C.ink } }], { x, y, w, h: 0.45, valign: "middle", wrap: false });
      y += 0.47;
      s.addShape(SH.LINE, { x, y, w, h: 0, line: { color: ref ? C.line : C.ink, width: ref ? 0.75 : 1.25 } });
      y += 0.14;
      if (col.note) { T(s, col.note, { x, y, w, h: 0.5, fontSize: FS.body, color: C.gray, lineSpacingMultiple: 1.15 }); y += 0.55; }
      col.items.forEach(it => {
        T(s, it.no, { x: x + 0.25, y, w: 0.45, h: 0.4, fontSize: FS.block, bold: true, color: C.gray, valign: "middle", wrap: false });
        T(s, [{ text: it.title }, { text: `　P.${PAGE[it.key] ?? "?"}`, options: { fontSize: FS.note, color: C.faint } }], { x: x + 0.75, y, w: w - 0.75, h: 0.4, fontSize: FS.block, color: ref ? C.soft : C.ink, valign: "middle", wrap: false });
        y += 0.4;
      });
    });
  };
  // 目次のページを作る（中身は finalize で描く）
  const agenda = (cols, o = {}) => { AGENDA_DONE = true; const s = base(o.title || "目次", null, { nav: o.nav, key: o.key }); later(() => drawAgenda(s, cols, o)); return s; };

  // アジェンダ：話す順の項目名だけを縦に並べる（createDeck の nav から作る）。主張や数字は書かない（それは要約の役目で、アジェンダとは別のページ）
  //   各ページのタイトルは、この項目名と同じ言葉になる（base がタイトルを nav から取る）。右上のミニ目次が、現在地を示す
  //   o.title（"アジェンダ"）, o.active（章の番号。章の変わり目でもう一度出すとき、その章だけ濃くする）, o.pages（false でページ番号を出さない）, o.key
  //   項目は 3〜7 個が目安。10 個を超えると警告（章をまとめる）
  const agendaPage = (o = {}) => {
    if (!NAV) throw new Error("agendaPage：createDeck の nav に章立てを書く（[{ name, ch }]）");
    AGENDA_DONE = true;
    const keep = chapterNow; chapterNow = 0; // アジェンダのページ自体は、どの章にも入れない
    const s = base(o.title || "アジェンダ", null, { key: o.key, focus: false });
    chapterNow = keep;
    if (NAV.length > 10) warn(`アジェンダの項目が ${NAV.length} 個ある。章をまとめて 7 個前後にする`);
    const n = NAV.length, y0 = 1.75, pitch = Math.min(0.7, (BOTTOM - y0) / n), fsz = pitch < 0.55 ? 18 : 22;
    later(() => NAV.forEach((nv, i) => {
      const y = y0 + i * pitch, on = o.active === undefined || o.active === nv.ch, col = on ? C.ink : C.pale;
      T(s, String(i + 1), { x: M + 0.1, y, w: 0.5, h: pitch, fontFace: FNUM, fontSize: fsz, bold: true, color: on ? C.gray : C.pale, valign: "middle", wrap: false });
      T(s, nv.name, { x: M + 0.7, y, w: 6, h: pitch, fontSize: fsz, bold: true, color: col, valign: "middle", wrap: false });
      if (o.pages !== false && PAGE[`ch${nv.ch}`] !== undefined) T(s, `P.${PAGE[`ch${nv.ch}`]}`, { x: M + 6.8, y, w: 1, h: pitch, fontSize: FS.note, color: C.faint, valign: "middle", wrap: false });
    }));
    return s;
  };

  // ミニ目次：右上 y0.38、FS.note（16pt）右揃え。全ページで同じ項目・同じ位置、現在の章だけ ink bold（4.12 miniAgenda）。finalize で全ページに一括で描く
  // ナビゲーションの形：テーマの nav.style（text＝文字だけ／underline＝今の章に下線／tabs＝箱を並べる／bar＝全幅の帯）。build.js の miniNav が勝つ
  let NAVT = TH.nav || {};
  const NAV_STYLES = ["text", "underline", "tabs", "bar"];
  let NAV_STYLE = opts.miniNav === false || NAVT.style === false ? null : NAVT.style ?? (opts.miniNav === true || Object.keys(NAVT).length ? "text" : null);
  if (NAV_STYLE && !NAV_STYLES.includes(NAV_STYLE)) { themeWarn("nav", `nav.style は ${NAV_STYLES.join("・")} か false にする（${NAV_STYLE}）。text を使う`); NAV_STYLE = "text"; }
  const navLayout = () => {
    const NUM = "①②③④⑤⑥⑦⑧⑨";
    let size = NAVT.size ?? (NAV_STYLE === "text" ? Math.min(FS.note, FOOT_FS) : 14);
    if (size < 11 || size > 18) { themeWarn("nav", `nav.size は 11〜18 にする（${size}）`); size = Math.min(18, Math.max(11, size)); }
    const bar = NAV_STYLE === "bar", tabs = NAV_STYLE === "tabs";
    const h = NAVT.h ?? (bar ? 0.42 : tabs ? 0.34 : 0.28), y = NAVT.y ?? (bar ? 0 : Y.section);
    const items = [];
    NAV.forEach((nv, i) => {
      items.push({ t: NAVT.numbers ? `${i + 1}　${nv.name}` : nv.name, on: a => a === nv.ch, near: a => (nv.subs || []).includes(a), gap: i ? (tabs ? 0.06 : 0.3) : 0 });
      if (!bar && NAVT.align !== "justify") (nv.subs || []).forEach((c, k) => items.push({ t: NUM[k], on: a => a === c, gap: k ? 0.04 : 0.08 }));
    });
    const align = NAVT.align === "left" ? "left" : "right";
    if (bar || NAVT.align === "justify") {
      const x = NAVT.x ?? (bar ? 0 : M), w = NAVT.w ?? (bar ? W : W - 2 * M), gap = bar ? 0 : 0.1, cw = (w - gap * (items.length - 1)) / items.length;
      items.forEach((o, i) => { o.x = x + i * (cw + gap); o.w = cw; if (textW(o.t, size) > cw - 0.1) warnTheme("nav.long", `ナビゲーションの項目が帯の区画に収まらない（${o.t}）。項目名を縮めるか nav.size を下げる`); });
      return { items, size, y, h, x, w, bar, tabs, justify: !bar };
    }
    items.forEach(o => { o.w = textW(o.t, size) * 1.04 + (tabs ? 0.28 : 0.02); });
    const total = items.reduce((a, o) => a + o.w + o.gap, 0);
    let x = align === "left" ? NAVT.x ?? M : (NAVT.x !== undefined ? NAVT.x + (NAVT.w ?? total) : W - M) - total;
    if (total > W - 2 * M + 0.01) warnTheme("nav.wide", "ナビゲーションの項目が、横幅に収まらない。章を減らすか、項目名を縮める");
    items.forEach(o => { x += o.gap; o.x = x; x += o.w; });
    return { items, size, y, h, x: items[0].x, w: total, bar, tabs };
  };
  const drawNav = (s, active, L, type) => {
    const color = (v, def) => (v !== undefined ? hexOf(v, def) : def);
    const dark = !L.bar && MASTER[type] && MASTER[type].drawn && worstContrast(type, { x: L.x, y: L.y, w: L.w, h: L.h }, C.ink) < 3;
    ((MASTER[type] && MASTER[type].objects) || []).forEach((ob, i) => {
      if ((ob.transparency ?? 0) >= 80 || (ob.x <= L.x + 0.01 && ob.y <= L.y + 0.02 && ob.x + ob.w >= L.x + L.w - 0.01 && ob.y + ob.h >= L.y + L.h - 0.02)) return;
      if (Math.min(L.x + L.w, ob.x + ob.w) - Math.max(L.x, ob.x) > 0.05 && Math.min(L.y + L.h, ob.y + (ob.h || 0.02)) - Math.max(L.y, ob.y) > 0.05) warnTheme(`nav.${type}.obj${i}`, `ナビゲーションが、${type} の飾り objects[${i}]（${ob.type}）に重なる。nav.x・nav.y か飾りの位置を替える`);
    });
    const idle = color(NAVT.color, dark ? "D5DBE5" : C.soft);
    const onFill = color(NAVT.activeFill, dark ? C.white : C.emph);
    const on = color(NAVT.activeColor, L.tabs || L.bar ? (relLum(onFill) > 0.35 ? C.ink : C.white) : dark ? C.white : C.ink);
    const base0 = color(NAVT.fill, C.tint);
    if (L.bar) rect(s, L.x, L.y, L.w, L.h, base0);
    L.items.forEach(o => {
      const act = o.on(active) || (o.near && o.near(active));
      if (L.tabs) rect(s, o.x, L.y, o.w, L.h, act ? onFill : base0);
      else if (L.bar && act) rect(s, o.x, L.y, o.w, L.h, onFill);
      const tabIdle = color(NAVT.color, C.soft);
      if (NAV_STYLE === "underline" && !act && NAVT.fill !== undefined) rect(s, o.x, L.y + L.h + 0.01, o.w, 0.04, base0);
      T(s, o.t, { x: o.x, y: L.y, w: o.w, h: L.h, fontSize: L.size, align: L.tabs || L.bar ? "center" : L.justify ? "left" : "right", valign: "middle", wrap: false, color: act ? on : L.tabs || L.bar ? tabIdle : idle, bold: !!act });
      if (NAV_STYLE === "underline" && act) rect(s, o.x, L.y + L.h + 0.01, o.w, 0.04, color(NAVT.activeFill, dark ? C.white : C.emph));
    });
  };

  // 出典の1行：既定では付けない（既定6）。詳細へ飛ばしたい数値だけ、リンク付きで1行置く（10.9）
  const source = (s, runs) => T(s, typeof runs === "string" ? runs : runs, { x: M, y: FOOT_TEXT ? Math.min(H - 0.86, Y.source) : Y.source, w: BW, h: 0.3, fontSize: FS.body, color: C.soft, valign: "middle" }); // 出典は要る情報なので、読める大きさと色で // フッターに資料名があるときは、その線の上に置く

  // ---- 小部品 ----
  // 塗りの右矢印（既定 gray、枠なし）。文字の「→」を混ぜない
  const arrowR = (s, x, y, w = 0.45, h = 0.32, color = C.gray) => s.addShape(SH.RIGHT_ARROW, { x, y, w, h, fill: { color }, line: { type: "none" } });
  // 手順図の流れの小三角（0.1×0.14 の灰。dir は r/l/u/d。4.13 PT）
  const PT = (s, dir, x, y, w = 0.1, h = 0.14, color = C.gray) => s.addShape(SH.ISOSCELES_TRIANGLE, { x, y, w, h, rotate: { r: 90, l: 270, d: 180, u: 0 }[dir], fill: { color }, line: { type: "none" } });
  // 紺の番号丸（手順の頭など、番号を付けたもの自体がその場にあるときに使う。4.13・10.13）
  const numDot = (s, n, x, y, d = 0.3) => {
    s.addShape(SH.OVAL, { x, y, w: d, h: d, fill: { color: C.emph }, line: { type: "none" } });
    T(s, String(n), { x, y, w: d, h: d, fontSize: d < 0.28 ? 9.5 : 11, bold: true, color: C.white, align: "center", valign: "middle", fontFace: FNUM });
  };
  // 赤の番号丸（選んだ付箋・吹き出し・図中の位置。白い縁取り。cx, cy は中心。4.13）
  const redDot = (s, n, cx, cy, d = 0.3) => s.addText(String(n), { shape: SH.OVAL, x: cx - d / 2, y: cy - d / 2, w: d, h: d, fill: { color: C.accent }, line: { color: C.white, width: 1.5 }, fontFace: FNUM, fontSize: d >= 0.32 ? 13 : 10, bold: true, color: C.white, align: "center", valign: "middle", margin: 0 });
  // 人のピクトグラム（頭＋肩。編集できる図形。round2SameRect に rectRadius を書くと壊れるので付けない。4.14）
  const person = (s, cx, top, d = 0.46, color = C.ink) => {
    s.addShape(SH.OVAL, { x: cx - d / 2, y: top, w: d, h: d, fill: { color }, line: { type: "none" } });
    s.addShape(SH.ROUND_2_SAME_RECTANGLE, { x: cx - d * 0.9, y: top + d * 1.12, w: d * 1.8, h: d * 1.05, fill: { color }, line: { type: "none" } });
  };

  // ---- 塗りブロックの表 ----
  // セル1つ：tint の矩形＋「見出し bold ink ＋ 6pt の空行 ＋ 説明 regular soft」。強調・問題・該当なし・例示を描き分ける
  //   cell: 文字列 | null（該当なし＝破線の空枠） | { head, body, strong, problem, frame, tag, num, unit, span, align, empty }
  const cellRuns = (c, o = {}) => {
    const runs = [], up = c.strong || c.frame ? 1 : 0, headColor = c.problem ? C.accent : C.ink;
    if (c.num !== undefined) {
      runs.push({ text: String(c.num), options: { fontSize: FS_DISPLAY.num, bold: true, color: headColor } });
      if (c.unit) runs.push({ text: c.unit, options: { fontSize: FS.block, bold: true, color: headColor } });
      if (c.head || c.body) runs[runs.length - 1].options.breakLine = true;
    }
    if (c.head) runs.push({ text: c.head, options: { fontSize: (o.headFs || FS.block) + up, bold: true, color: headColor, breakLine: !!c.body } });
    if (c.head && c.body) runs.push({ text: " ", options: { fontSize: 6, breakLine: true } });
    if (c.body) runs.push({ text: c.body, options: { fontSize: o.bodyFs || FS.body, color: c.head || c.num !== undefined ? C.soft : (o.bodyColor || C.ink) } });
    return runs;
  };
  // セルに必要な高さの見積もり
  const cellNeed = (c, w, o = {}) => {
    if (!c || typeof c !== "object") return c ? lineCount(c, o.bodyFs || FS.body, w - 2 * PAD.x) * lineH(o.bodyFs || FS.body) + 2 * PAD.y : 0;
    if (c.mark) w -= o.markW ?? 0.42;
    let h = 2 * PAD.y;
    if (c.num !== undefined) h += lineH(FS_DISPLAY.num, 1.1);
    if (c.head) h += lineCount(c.head, (o.headFs || FS.block) + (c.strong ? 1 : 0), w - 2 * PAD.x) * lineH(o.headFs || FS.block);
    if (c.head && c.body) h += lineH(6);
    if (c.body) h += lineCount(c.body, o.bodyFs || FS.body, w - 2 * PAD.x) * lineH(o.bodyFs || FS.body);
    if (c.tag) h += 0.2;
    return h;
  };
  const drawCell = (s, x, y, w, h, c, o = {}) => {
    if (c === undefined) return;
    if (c === null || (typeof c === "object" && c.empty)) {
      // 該当なしは塗らず line 色 0.75pt の破線の空枠。「未設定」「要確認」と書かない（4.1-11）
      rect(s, x, y, w, h, null, { color: C.line, width: 0.75, dashType: "dash" });
      if (c && c.empty && c.text) T(s, c.text, { x, y, w, h, fontSize: FS.body, color: C.pale, align: "center", valign: "middle" });
      return;
    }
    const cell = typeof c === "string" ? { body: c } : c;
    const fill = cell.problem || o.example ? C.white : (cell.strong || o.strong) ? TINT_STRONG : (o.onFace ? C.white : TINT);
    const line = cell.problem ? { color: C.accent, width: 2 } : cell.frame ? { color: C.emph, width: 2 } : o.example ? { color: C.line, width: 0.75 } : null;
    rect(s, x, y, w, h, fill, line);
    if (cell.tag) T(s, cell.tag, { x: x + PAD.x, y: y + 0.08, w: w - 2 * PAD.x, h: 0.22, fontSize: FS.note, bold: true, color: cell.problem ? C.accent : C.ink, wrap: false });
    // 評価の記号（◎○△×）：セルの左に置き、コメントは右に続ける。強い印（◎○）は ink の太字、弱い印（△×）は灰
    const w0 = w; // 高さの見積もり（cellNeed）は記号の幅を自分で引くので、元の幅を渡す
    if (cell.mark) {
      const weak = /[△×－-]/.test(cell.mark), mw = o.markW ?? 0.42;
      T(s, cell.mark, { x: x + PAD.x - 0.04, y, w: mw, h, fontSize: o.markFs || 19, bold: !weak, color: weak ? C.pale : C.ink, align: "center", valign: "middle", wrap: false });
      x += mw; w -= mw;
    }
    if (cell.head) checkWrap(cell.head, (o.headFs || FS.block) + (cell.strong ? 1 : 0), w - 2 * PAD.x);
    if (cell.body) checkWrap(cell.body, o.bodyFs || FS.body, w - 2 * PAD.x);
    const need = cellNeed(cell, w0, o);
    if (need > h + 0.02) warn(`ブロックの文字があふれる（「${String(cell.head || cell.body).slice(0, 12)}…」 必要 ${need.toFixed(2)} > 高さ ${h.toFixed(2)}）`);
    // 小見出し（tag）があるときは、小見出しのすぐ下から上詰めで置く（上下中央だと間が空く）
    T(s, cellRuns(cell, o), { x: x + PAD.x, y: cell.tag ? y + 0.32 : y, w: w - 2 * PAD.x, h: cell.tag ? h - 0.4 : h, valign: cell.tag ? "top" : "middle", align: cell.align, lineSpacingMultiple: 1.25 });
  };
  // 塗りブロックの表：列見出しと行見出しは塗らず、セルは独立した矩形、間は白い溝（4.1）。addTable と罫線を使わない（既定1）
  //   o.cols: [{ head, w?, strong?, example?, ref? }]（w を省いた列は等分）
  //   o.rows: [{ label?, no?, h?, strong?, cells: [cell, ...] }]（cell.span で複数列にまたがる1本の帯。strong＝選ぶ1行を濃い塗り＋太字＋行全体の ink 枠で示す）
  //   o.x, o.w, o.y（最初の行の上端。既定 2.75）, o.headY, o.labelW, o.gap（溝。既定 GAP.std）, o.rowGap, o.rowH（数か配列。省くと一番高い行に全行をそろえる）, o.onFace（面の上ではセルを白）
  const blockTable = (s, o) => {
    warn("文字を入れる表（blockTable）は使わない。全部読まないと分からない表は、表の意味をなさない。項目ごとの箇条書き（points）か、図にする");
    // 溝：列の間は狭く、行の間は広く（縦横同じ幅だと格子に見え、読む向きが決まらない。調査 1.0）
    const gap = o.gap ?? GAP.dense, rg = o.rowGap ?? 0.16;
    const x0 = o.x ?? M, wAll = o.w ?? (W - M - x0);
    const hasLabel = o.rows.some(r => r.label);
    const lw = hasLabel ? (o.labelW ?? 1.7) : 0, lg = hasLabel ? 0.15 : 0;
    const cx0 = x0 + lw + lg, cwAll = wAll - lw - lg;
    const fixed = o.cols.reduce((a, c) => a + (c.w || 0), 0) + gap * (o.cols.length - 1);
    const flexN = o.cols.filter(c => !c.w).length, fw = flexN ? (cwAll - fixed) / flexN : 0;
    const xs = []; let x = cx0;
    o.cols.forEach(c => { const w = c.w || fw; xs.push([x, w]); x += w + gap; });
    const y0 = o.y ?? 2.75, hy = o.headY ?? y0 - 0.38;
    const wOf = (j, span = 1) => xs[j + span - 1][0] + xs[j + span - 1][1] - xs[j][0];
    // 行の高さ：指定がなければ、全行を一番高い行に合わせる（4.1-7）。本文の下端まで空くときは行を広げて面を埋める（6.2「空いた面」。広げるのは 1.4 まで。o.compact で止める）
    const need = o.rows.map(r => { let j = 0; return Math.max(0.7, ...r.cells.map(c => { const span = (c && c.span) || 1, h = cellNeed(c, wOf(j, span), o); j += span; return h; }), r.label ? lineCount(r.label, FS.head, lw) * lineH(FS.head, 1.2) : 0); });
    const fillH = o.compact ? 0 : Math.min(1.4, ((o.bottom ?? BOTTOM) - y0 - rg * (o.rows.length - 1)) / o.rows.length);
    const rhs = o.rows.map((r, i) => r.h ?? (Array.isArray(o.rowH) ? o.rowH[i] : o.rowH) ?? Math.max(fillH, ...need));
    // 列見出し：塗らず、下線を引かず、bold gray を最初の行の上に置く（4.1-3）
    // 列見出し：塗らず、下線を引かない。比べる対象の名前なので bold ink（灰だと表の中で一番弱い文字になる。調査 1.0）
    o.cols.forEach((c, j) => c.head && T(s, c.head, { x: xs[j][0], y: hy, w: xs[j][1], h: 0.3, fontSize: FS.block, bold: true, color: c.ref ? C.soft : C.ink, valign: "bottom", wrap: false }));
    let y = y0;
    const rowY = [];
    o.rows.forEach((r, i) => {
      const rh = rhs[i];
      rowY.push(y);
      // 行見出し：塗らず bold ink、番号は gray＋全角スペース（4.1-4）
      if (r.label) T(s, r.no !== undefined ? [{ text: `${r.no}　`, options: { color: C.gray } }, { text: r.label }] : r.label, { x: x0, y, w: lw, h: rh, fontSize: FS.head, bold: true, valign: "middle", lineSpacingMultiple: 1.2 });
      let j = 0;
      r.cells.forEach(c => {
        const span = (c && c.span) || 1;
        // 行の強調（選択肢が行に並ぶ表で、選ぶ1行を示す）：セルを濃い塗り・太字・1段大きい文字にする
        const cc = r.strong && c ? (typeof c === "string" ? { head: c, strong: true } : { ...c, strong: true }) : c;
        drawCell(s, xs[j][0], y, wOf(j, span), rh, cc, { ...o.cols[j], onFace: o.onFace, headFs: r.strong ? (o.headFs ?? FS.block) + 2 : o.headFs, bodyFs: o.bodyFs });
        j += span;
      });
      // 行全体を ink 2pt の枠1本で囲む（セルごとに枠を付けない）
      if (r.strong) rect(s, cx0 - 0.02, y - 0.02, cwAll + 0.04, rh + 0.04, null, { color: C.emph, width: 2 });
      y += rh + rg;
    });
    const bottom = y - rg;
    if (bottom > BOTTOM + 0.05) warn(`表の下端 ${bottom.toFixed(2)} が本文の下端 ${BOTTOM} を超える`);
    return { xs, rowY, rowH: rhs, bottom, headY: hy };
  };

  // 箇条書き（1列）：事実の列挙・活動報告・総括・知見は図にせず、これで書く（3章「事実の列挙」、既定5）
  //   items: [文字列 | { head, items: [文字列, ...] }]（head 付きは小見出し＋その下の箇条書き）
  //   o: { x, y（既定 BODY）, w（既定 BW）, h（既定は本文の下端まで）, fs（既定 15）, headFs（既定 17）, numbered（番号を付ける）, gap（項目の間 pt、既定 fs×0.9） }
  //   項目が少なく下に空きが出るときは、文字を上げるより gap を広げ、h を本文の下端までにして上下の重心をそろえる
  const bullets = (s, items, o = {}) => {
    const fs = o.fs ?? 20, hfs = o.headFs ?? 22, gap = o.gap ?? fs * 0.9; // 案B
    const runs = []; let n = 0;
    const item = t => { n++; runs.push({ text: t, options: { fontSize: fs, color: C.ink, bullet: o.numbered ? { type: "number", indent: fs * 1.4 } : { indent: fs * 1.1 }, paraSpaceAfter: gap, breakLine: true } }); };
    items.forEach(it => {
      if (typeof it === "string") return item(it);
      runs.push({ text: it.head, options: { fontSize: hfs, bold: true, color: C.ink, paraSpaceBefore: runs.length ? gap : 0, paraSpaceAfter: gap * 0.5, breakLine: true } });
      (it.items || []).forEach(item);
    });
    const x = o.x ?? M, y = o.y ?? BODY, w = o.w ?? BW, h = o.h ?? BOTTOM - y;
    items.flatMap(it => typeof it === "string" ? [it] : it.items || []).forEach(t => checkWrap(t, fs, w - 0.4));
    T(s, runs, { x, y, w, h, valign: o.valign ?? "top", lineSpacingMultiple: 1.3 });
    return n;
  };

  // ---- 時系列 ----
  // 月の座標：from〜to（"2025.04" の形）を軸 x0〜x1 に等分する。mx＝月の左端、mxEnd＝右端、span＝期間、mid＝中央（元データ7 mx/span/mid）
  const timeline = (from, to, o = {}) => {
    const months = [];
    let [y, m] = from.split(".").map(Number);
    const [ty, tm] = to.split(".").map(Number);
    while (y < ty || (y === ty && m <= tm)) { months.push(`${y}.${String(m).padStart(2, "0")}`); m++; if (m > 12) { m = 1; y++; } }
    const x0 = o.x0 ?? M + 1.55, x1 = o.x1 ?? W - M, mw = (x1 - x0) / months.length;
    const idx = p => { const i = months.indexOf(p); if (i < 0) throw new Error(`timeline: ${p} は ${from}〜${to} の外`); return i; };
    const mx = p => x0 + idx(p) * mw;
    return { months, x0, x1, mw, fyStart: o.fyStart ?? 4, mx, mxEnd: p => mx(p) + mw, span: (a, b) => ({ x: mx(a), w: (idx(b) - idx(a) + 1) * mw }), mid: p => mx(p) + mw / 2 };
  };
  // ガント：行＝区分（区分色の帯、枠なし）、列＝月。FY と Q の目盛り。節目は赤の縦破線＋ひし形＋右揃えの赤字で、グリッド線上・帯の背面に描く（4.5 drawGantt）
  //   o.rows: [{ label, color, bars: [[from, to, 文字], ...] }], o.milestones: [{ at, edge: "start"|"end", label, side: "l"|"r" }]
  //   同じ図を別ページで重ねるときは、同じ tl・同じ o で呼ぶ（4.11）。行が少ないと行の間を広げて本文の下端まで使う。重ねる吹き出しの場所を残すときは o.compact
  const drawGantt = (s, tl, o) => {
    const y0 = o.y0 ?? BODY + 0.1, n = o.rows.length, fz = o.fs ?? FS.body;
    // 行の高さ：予定と実績を重ねる行は帯を細くしすぎないよう高めにする。行が少なく下が空くときは、行の間を広げて本文の下端まで使う（o.compact で止める）
    const LH = o.LH ?? (o.rows.some(r => r.plan) ? 0.55 : 0.4);
    const fillRG = n > 1 ? (BOTTOM - 0.1 - (y0 + 0.6) - n * LH) / (n - 1) : 0.3;
    const RG = o.RG ?? (o.compact ? 0.3 : Math.min(0.9, Math.max(0.3, fillRG)));
    const rowsTop = y0 + 0.6, rowsBottom = rowsTop + n * LH + (n - 1) * RG, gridB = rowsBottom + 0.1;
    hline(s, tl.x0, y0 + 0.42, tl.x1 - tl.x0);
    const fyDone = new Set();
    // o.axis === "month"：年度・四半期の代わりに暦年と月の目盛り（構成案が暦の年月で書かれ、月単位の遅れを読ませるページ用）
    if (o.axis === "month") tl.months.forEach((p, i) => {
      const [yy, mm] = p.split(".").map(Number), x = tl.mx(p), yearHead = i === 0 || mm === 1;
      if (i > 0) s.addShape(SH.LINE, { x, y: mm === 1 ? y0 : y0 + 0.22, w: 0, h: gridB - (mm === 1 ? y0 : y0 + 0.22), line: { color: mm === 1 ? C.gray : C.line, width: 0.75 } });
      T(s, `${mm}月`, { x, y: y0 + 0.2, w: tl.mw, h: 0.2, fontSize: FS.note, color: C.gray, align: "center" });
      if (yearHead) T(s, `${yy}年`, { x: x + 0.06, y: y0 - 0.02, w: 1, h: 0.22, fontSize: FS.note, bold: true, color: C.soft, wrap: false });
    });
    else tl.months.forEach((p, i) => {
      const [yy, mm] = p.split(".").map(Number), q = (mm - tl.fyStart + 12) % 12;
      if (q % 3 !== 0) return;
      const fyHead = q === 0, fy = mm >= tl.fyStart ? yy : yy - 1, x = tl.mx(p);
      if (i > 0 || fyHead) s.addShape(SH.LINE, { x, y: fyHead ? y0 : y0 + 0.22, w: 0, h: gridB - (fyHead ? y0 : y0 + 0.22), line: { color: fyHead ? C.gray : C.line, width: 0.75 } });
      T(s, `${q / 3 + 1}Q`, { x, y: y0 + 0.2, w: Math.min(3 * tl.mw, tl.x1 - x), h: 0.2, fontSize: FS.note, color: C.gray, align: "center" });
      if (!fyDone.has(fy)) { fyDone.add(fy); T(s, `FY${String(fy).slice(2)}`, { x: x + 0.06, y: y0 - 0.02, w: 1, h: 0.22, fontSize: FS.note, bold: true, color: C.soft, wrap: false }); }
    });
    // 節目の線は帯の背面に描く（帯の文字に重ねない）
    const ms = (o.milestones || []).map(m => ({ ...m, x: m.edge === "end" ? tl.mxEnd(m.at) : tl.mx(m.at) }));
    ms.forEach(m => s.addShape(SH.LINE, { x: m.x, y: y0 + 0.42, w: 0, h: rowsBottom - y0 - 0.42, line: { color: C.accent, width: m.width || 1, dashType: "dash" } }));
    const rowY = [];
    let y = rowsTop;
    // 行に plan（予定）があるときは、行の上側に予定の細い帯（soft 1pt の破線の空枠）、下側に実績の帯を描く。全行で同じ段組みにする
    const planMode = o.rows.some(r => r.plan), pH = planMode ? (o.planH ?? 0.16) : 0, pG = planMode ? 0.05 : 0, bH = LH - pH - pG;
    o.rows.forEach(r => {
      T(s, r.label, { x: M, y, w: tl.x0 - M - 0.1, h: LH, fontSize: FS.body, bold: true, color: C.soft, valign: "middle", lineSpacingMultiple: 1.0 });
      const dark = isDark(r.color);
      (r.plan || []).forEach(([a, b]) => { const { x, w } = tl.span(a, b); rect(s, x + 0.02, y, w - 0.04, pH, null, { color: C.soft, width: 1, dashType: "dash" }); });
      r.bars.forEach(([a, b, t]) => {
        const { x, w } = tl.span(a, b), by = y + pH + pG;
        rect(s, x + 0.02, by, w - 0.04, bH, r.color);
        if (t) {
          T(s, t, { x: x + 0.07, y: by, w: w - 0.09, h: bH, fontSize: fz, bold: dark, color: dark ? C.white : C.ink, valign: "middle", lineSpacingMultiple: 0.95 });
          if (lineCount(t, fz, w - 0.09) > Math.floor(bH / lineH(fz, 0.95))) warn(`ガントの帯の文字があふれる（${t}）`);
        }
      });
      // 帯の外に置く短い注記（r.notes: [[月, 文字]]）。指す帯の隣に置く
      (r.notes || []).forEach(([a, t]) => T(s, t, { x: tl.mx(a) + 0.06, y: y + pH + pG, w: textW(t, FS.body) * 1.05 + 0.1, h: bH, fontSize: FS.body, bold: true, color: C.ink, valign: "middle", wrap: false }));
      rowY.push(y);
      y += LH + RG;
    });
    ms.forEach(m => {
      s.addShape(SH.DIAMOND, { x: m.x - 0.08, y: y0 + 0.42 - 0.08, w: 0.16, h: 0.16, fill: { color: C.accent }, line: { type: "none" } });
      const tw = textW(m.label, FS.note) * 1.05 + 0.04;
      // ラベルは既定で線の左。左に入らない（軸の左端や年の表示にかかる）ときは右へ置く
      const side = m.side ?? (m.x - 0.1 - tw < tl.x0 + 0.7 ? "r" : "l");
      T(s, m.label, { x: side === "r" ? m.x + 0.1 : m.x - 0.1 - tw, y: y0 - 0.02, w: tw, h: 0.22, fontSize: FS.note, bold: true, color: C.accent, align: side === "r" ? "left" : "right", wrap: false });
    });
    // 予定と実績の凡例。行ラベル列の目盛りの高さに縦に置く（目盛りの文字と重ねない）。破線の線種の意味を図の中で示す
    if (planMode) {
      const lx = M, ly = y0 + 0.02;
      rect(s, lx, ly + 0.04, 0.3, 0.1, null, { color: C.soft, width: 1, dashType: "dash" });
      T(s, "予定", { x: lx + 0.38, y: ly, w: 0.6, h: 0.18, fontSize: FS.note, color: C.soft, valign: "middle", wrap: false });
      rect(s, lx, ly + 0.24, 0.3, 0.14, "DADDE1");
      T(s, "実績", { x: lx + 0.38, y: ly + 0.22, w: 0.6, h: 0.18, fontSize: FS.note, color: C.soft, valign: "middle", wrap: false });
    }
    return { y0, rowY, LH, RG, rowMid: i => rowY[i] + LH / 2, bottom: rowsBottom };
  };

  // ---- 重ね技 ----
  // 前ページと同じ図を同じ関数・同じ座標で描いた上に、白の半透明を重ねて薄くする（transparency 18〜30）。前面に吹き出し・メッセージを置く
  const overlay = (s, x = 0, y = 2.15, w = W, h = BOTTOM - 2.15, transparency = 30) => s.addShape(SH.RECTANGLE, { x, y, w, h, fill: { color: PAGE_BG(), transparency }, line: { type: "none" } });
  // 前面の吹き出し：白箱に赤 1.25 枠＋弱い影、赤の番号丸、見出し bold＋起きたこと soft。引き出し線で図中の位置（赤の番号丸）を指す（4.8）
  //   o: { x, y, w, h, no, head, body, to: [x, y]（指す位置）, from: [x, y]（線の出どころ。既定は箱の下辺中央）, refKey（右下に「P.n」） , refPrefix }
  const callout = (s, o) => {
    const h = o.h ?? 1.0, from = o.from || [o.x + o.w / 2, o.to && o.to[1] < o.y ? o.y : o.y + h];
    if (o.to) segLine(s, from[0], from[1], o.to[0], o.to[1], { color: C.accent, width: 1.25 });
    s.addShape(SH.RECTANGLE, { x: o.x, y: o.y, w: o.w, h, fill: { color: C.white }, line: { color: C.accent, width: 1.25 }, shadow: { type: "outer", color: "000000", opacity: 0.15, blur: 6, offset: 2, angle: 90 } });
    if (o.no !== undefined) redDot(s, o.no, o.x + 0.3, o.y + 0.3, 0.3);
    const tx = o.no !== undefined ? o.x + 0.55 : o.x + PAD.x, tw = o.x + o.w - tx - 0.15;
    T(s, o.head, { x: tx, y: o.y + 0.1, w: tw, h: 0.4, fontSize: FS.block, bold: true, valign: "middle", wrap: false });
    if (textW(o.head, FS.block) > tw) warn(`吹き出しの見出しが1行に収まらない（${o.head}）`);
    if (o.body) checkWrap(o.body, FS.body, tw);
    if (o.body) T(s, o.body, { x: tx, y: o.y + 0.48, w: tw, h: h - 0.55, fontSize: FS.body, color: C.soft, lineSpacingMultiple: 1.2 });
    if (o.to && o.no !== undefined) redDot(s, o.no, o.to[0], o.to[1], 0.34);
    if (o.refKey) ref(s, o.refKey, { x: o.x + o.w - 1.3, y: o.y + h - 0.32, w: 1.2, prefix: o.refPrefix });
  };

  // 前面のメッセージボックス：同じ図を背面に敷いたページで、分かったことを図の上に重ねて示す（4.9・4.11）
  //   白箱に ink 1pt の枠＋弱い影、見出し FS.head（18pt）bold＋説明 FS.body（16pt）soft。o.to で指す位置へ赤の線（視線誘導）を引き、先端に小さな丸
  //   o: { x, y, w, h, head, body?, to?: [x, y] | toRange?: [[x1, y], [x2, y]]（範囲を指す）, num?（左上に紺の番号丸）, refKey?（右下に「P.n」、根拠のページ）, refPrefix? }
  const messageBox = (s, o) => {
    if (o.head && em(o.head) > 20) warn(`メッセージボックスの見出しが長い（全角${Math.round(em(o.head))}字。20字まで）：${o.head}`);
    if (o.body) String(o.body).replace(/\n/g, "").split("。").filter(x => x.trim()).forEach(x => { if (em(x) > 40) warn(`メッセージボックスの1文が長い（全角${Math.round(em(x))}字。40字まで）：「${x.slice(0, 18)}…」`); });
    // 高さを指定しないときは、文字の行数から決める
    const mbw = o.w - (o.num !== undefined ? 0.52 : PAD.x) - PAD.x;
    const need = lineCount(o.head, FS.head, mbw) * lineH(FS.head) + (o.body ? lineH(6) + lineCount(o.body, FS.body, mbw) * lineH(FS.body) : 0);
    const h = o.h ?? Math.max(o.body ? 1.0 : 0.6, Math.ceil((need + 0.1) * 20) / 20);
    // 範囲を指す（toRange: [[x1, y], [x2, y]]）：2点を赤の線で結び、両端に短い縦線を付けて、その中点を指す
    if (o.toRange) {
      const [[ax, ay], [bx, by]] = o.toRange, my = Math.min(ay, by);
      segLine(s, ax, my, bx, my, { color: C.accent, width: 1.25 });
      [ax, bx].forEach(x => segLine(s, x, my, x, my + 0.12, { color: C.accent, width: 1.25 }));
      o = { ...o, to: [(ax + bx) / 2, my], dot: false };
    }
    if (o.to) {
      const from = o.from || [o.to[0] < o.x ? o.x : o.to[0] > o.x + o.w ? o.x + o.w : o.x + o.w / 2, o.to[1] < o.y ? o.y : o.y + h];
      segLine(s, from[0], from[1], o.to[0], o.to[1], { color: C.accent, width: 1.25 });
      if (o.dot !== false) s.addShape(SH.OVAL, { x: o.to[0] - 0.05, y: o.to[1] - 0.05, w: 0.1, h: 0.1, fill: { color: C.accent }, line: { type: "none" } });
    }
    s.addShape(SH.RECTANGLE, { x: o.x, y: o.y, w: o.w, h, fill: { color: C.white }, line: { color: C.emph, width: 1 }, shadow: { type: "outer", color: "000000", opacity: 0.15, blur: 6, offset: 2, angle: 90 } });
    if (o.num !== undefined) numDot(s, o.num, o.x + 0.15, o.y + 0.15, 0.26);
    const tx = o.num !== undefined ? o.x + 0.52 : o.x + PAD.x, tw = o.x + o.w - tx - PAD.x;
    checkWrap(o.head, FS.head, tw); if (o.body) checkWrap(o.body, FS.body, tw);
    // 文字が大きいと箱も高くなる。本文の下端より下は、フッターにかかる
    if (o.y + h > BOTTOM + 0.35) warn(`メッセージボックスの下端 ${(o.y + h).toFixed(2)} がフッターにかかる（${(BOTTOM + 0.35).toFixed(2)} まで。箱を上げるか、図の下端（bottom）を上げる）`);
    { // 見出しと本文の行数が箱の高さに入るか（行間 1.25）
      if (need > h + 0.02) warn(`メッセージボックスの文字が箱の高さに入らない（見出し${lineCount(o.head, FS.head, tw)}行＋本文${o.body ? lineCount(o.body, FS.body, tw) : 0}行で ${need.toFixed(2)} ＞ 高さ ${h.toFixed(2)}）。h を ${need.toFixed(2)} 以上にするか、文を縮める`);
    }
    T(s, [{ text: o.head, options: { bold: true, fontSize: FS.head, breakLine: !!o.body } }, ...(o.body ? [{ text: " ", options: { fontSize: 6, breakLine: true } }, { text: o.body, options: { fontSize: FS.body, color: C.soft } }] : [])],
      { x: tx, y: o.y, w: tw, h, valign: "middle", lineSpacingMultiple: 1.25 });
    if (o.refKey) ref(s, o.refKey, { x: o.x + o.w - 1.3, y: o.y + h - 0.3, w: 1.2, prefix: o.refPrefix });
  };
  // 同じ図を背面に敷く：drawFig(s) で前のページと同じ図を同じ座標で描き、白の半透明で薄くする。前面の要素はこの後に足す（4.11）
  //   area: overlay の範囲 { x, y, w, h }（既定は本文の領域）。transparency は 18〜30
  //   keep: (s, r) => {}  覆いの上に、前面で指す要素（主役の棒、値のラベルなど）だけを元の濃さで描き直す。r は drawFig の戻り値
  //   check: true         覆いと keep を描かない（元のページと画像を引き比べ、図が動いていないことを確かめる用。確かめたら外す）
  const echo = (s, drawFig, o = {}) => {
    const r = drawFig(s);
    if (o.check) return r;
    const a = o.area || {};
    // spot：見る所の矩形（{x, y, w, h} か、図の戻り値から矩形を返す関数）。そこだけ覆わず、元の濃さで残す。only を持たない部品でも使える。見る所は1つなので、矩形も1つ
    const hole = typeof o.spot === "function" ? o.spot(r) : o.spot;
    if (hole) {
      const ax = a.x ?? 0, ay = a.y ?? 2.15, aw = a.w ?? W, ah = a.h ?? BOTTOM - 2.15, t = o.transparency ?? 25, pd = 0.06;
      const hx = Math.max(ax, hole.x - pd), hy = Math.max(ay, hole.y - pd), hx2 = Math.min(ax + aw, hole.x + hole.w + pd), hy2 = Math.min(ay + ah, hole.y + hole.h + pd);
      [[ax, ay, aw, hy - ay], [ax, hy2, aw, ay + ah - hy2], [ax, hy, hx - ax, hy2 - hy], [hx2, hy, ax + aw - hx2, hy2 - hy]].forEach(([x, y, w, h]) => { if (w > 0.005 && h > 0.005) overlay(s, x, y, w, h, t); });
    } else overlay(s, a.x, a.y, a.w, a.h, o.transparency ?? 25);
    if (o.keep) o.keep(s, r);
    return r;
  };
  // 同じ図の続きページ：1つの図で2つ以上を言うときは、図を1回だけ書き、言うことの数だけページを出す（1ページ1メッセージ）
  //   drawFig: (s, extra) => 部品の戻り値。extra は { only } など、部品の引数に足すもの（例: (s, x) => d.processFlow(s, { ...FLOW, ...x })）
  //   so: { area（覆う範囲。本文の下端を越える図は { y, h } を渡す）, transparency, check（true で覆いと前面を外す。図が動いていないかの確かめ用） }
  //   戻り値: { page(p), pages([p, ...]) }。p: { title, lead, focus（見る所。必須）, key?, section?,
  //             plain?（図だけのページ。1枚目の既定）, spot?（見る所の矩形か関数）, only?（部品の only に渡す値）, keep?（(s, r) => 前面に描き直す）, box?（messageBox の引数か、その配列か、図の戻り値から返す関数） }
  //   続きのページは、連続していなくてよい（ロードマップの図を、後ろの資金調達のページでもう一度敷く、など）
  const series = (drawFig, so = {}) => {
    let n = 0;
    const one = p => {
      const s = base(p.title, p.lead, { key: p.key, section: p.section, focus: p.focus, ask: p.ask });
      const first = n++ === 0, front = p.spot || p.only || p.keep || p.box;
      if (p.plain || (first && !front)) return { s, r: drawFig(s, {}) };
      const keep = p.keep ? (s2, r2) => p.keep(s2, r2) : p.only ? s2 => drawFig(s2, { only: p.only }) : undefined;
      const r = echo(s, s2 => drawFig(s2, {}), { area: so.area, transparency: so.transparency, check: so.check, spot: p.spot, keep });
      if (!so.check && p.box) [].concat(typeof p.box === "function" ? p.box(r) : p.box).forEach(b => messageBox(s, b));
      return { s, r };
    };
    return { page: one, pages: ps => ps.map(one) };
  };

  // ---- 付箋ボード ----
  // ブレスト専用。帯（区分）を上から、帯の中のまとまりを左から並べる。選んだ付箋だけ不透明＋赤枠＋赤の番号丸、他は透過45・灰字。傾けない
  //   o.bands: [{ key, name, color, rows }], o.items: [{ band, group, text, pick? }], o.y0, o.sh（付箋の高さ）, o.fs（本文の文字。既定 5.5 の例外）
  const stickyBoard = (s, o) => { smallOK = true; try { return stickyBoard0(s, o); } finally { smallOK = false; } };
  const stickyBoard0 = (s, o) => {
    const items = o.items, bands = o.bands, key = items.some(d => d.pick);
    const lw = o.labelW ?? 1.0, x0 = M + lw + 0.1, cgap = 0.16, sg = 0.05, sh = o.sh ?? 0.5, rg = 0.045, th = 0.22, bg = o.bandGap ?? 0.12, fz = o.fs ?? 5.5;
    const groupsOf = b => [...new Set(items.filter(d => d.band === b.key).map(d => d.group))];
    const colsOf = b => groupsOf(b).map(g => Math.ceil(items.filter(d => d.band === b.key && d.group === g).length / b.rows));
    // 付箋の幅：列数が一番多い帯で決める
    const widest = bands.map(b => { const cs = colsOf(b); return { nc: cs.reduce((a, c) => a + c, 0), ng: cs.length }; }).sort((a, b) => (b.nc * 1000 + b.ng) - (a.nc * 1000 + a.ng))[0];
    const sw = (W - M - x0 - cgap * (widest.ng - 1) - sg * (widest.nc - widest.ng)) / widest.nc;
    let y = o.y0 ?? 2.2;
    const geo = { bands: {}, items: [] };
    bands.forEach(b => {
      const its = items.filter(d => d.band === b.key), R = b.rows, bh = th + R * sh + (R - 1) * rg + 0.04;
      rect(s, M, y + th + 0.02, 0.12, bh - th - 0.04, b.color);
      T(s, [{ text: b.name, options: { breakLine: true } }, { text: `${its.length}件`, options: { fontSize: FS.note, color: C.gray, bold: false } }], { x: M + 0.2, y: y + th, w: lw - 0.2, h: bh - th, fontSize: FS.note, bold: true, valign: "middle", lineSpacingMultiple: 1.1 });
      let x = x0;
      groupsOf(b).forEach(g => {
        const gi = its.filter(d => d.group === g), nC = Math.ceil(gi.length / R), cw = nC * sw + (nC - 1) * sg;
        T(s, g, { x, y, w: Math.max(cw, 1.2), h: th, fontSize: FS.body, bold: true, color: C.soft, valign: "bottom", wrap: false });
        gi.forEach((d, i) => {
          const col = Math.floor(i / R), row = i % R, sx = x + col * (sw + sg), sy = y + th + 0.04 + row * (sh + rg), on = d.pick || !key;
          // 選んだ付箋は左上に番号丸を置くため、本文の左余白を広げて丸と重ねない
          s.addText(d.text, { shape: SH.RECTANGLE, x: sx, y: sy, w: sw, h: sh, fill: { color: b.color, transparency: on ? 0 : 45 }, line: d.pick ? { color: C.accent, width: 2.25 } : { type: "none" }, shadow: { type: "outer", color: "000000", opacity: 0.12, blur: 2, offset: 1, angle: 90 }, fontFace: F, fontSize: fz, color: on ? STICKY.text : STICKY.off, bold: !!d.pick, valign: "top", margin: d.pick ? [2.5, 2.5, 2, 10] : [2.5, 2.5, 2, 2.5], lineSpacingMultiple: 1.0, fit: "none" });
          if (lineCount(d.text, fz, sw - (d.pick ? 0.18 : 0.07)) * lineH(fz, 1.0) > sh - 0.04) warn(`付箋の文字があふれる（${d.text.slice(0, 12)}…）`);
          checkWrap(d.text, fz, sw - (d.pick ? 0.18 : 0.07));
          if (d.pick) redDot(s, d.pick, sx - 0.02, sy - 0.02, 0.26);
          geo.items.push({ ...d, x: sx, y: sy, w: sw, h: sh });
        });
        x += cw + cgap;
      });
      geo.bands[b.key] = { y, h: bh };
      y += bh + bg;
    });
    if (y - bg > BOTTOM + 0.05) warn(`付箋ボードの下端 ${(y - bg).toFixed(2)} が本文の下端を超える`);
    return geo;
  };

  // ---- 数値 ----
  // ウォーターフォール：期ごとの増分は gray の棒＋「+値」bold、右端の累計の棒だけ区分色で積み、右に「区分 soft＋値 bold」（4.6）
  //   o.steps: [{ label, value, sub? }], o.total: { label, parts: [{ name, value, color }] }, o.before: { label, value, sub }（累計に含めない額。破線の棒）
  //   o.base（基線 y、既定 6.25）, o.top（最大の高さの上端、既定 2.95）, o.labelW（累計の右のラベル幅、既定 2.55）, o.fmt
  //   主役の値: steps・total.parts の1つ、または total に focusValue: true。その値を主役の値 36pt で普通の位置（増分は棒の上、内訳は右のラベルの名前の後ろ、累計は棒の上）に部品が描く
  //     内訳を主役にしたときは、その区分を強調の色（C.emph）で塗る。呼ぶ側で値を空にして focusValue を手で置かない
  const waterfall = (s, o) => {
    const nFV = o.steps.filter(x => x.focusValue).length + o.total.parts.filter(x => x.focusValue).length + (o.total.focusValue ? 1 : 0);
    if (nFV > 1) warn("ウォーターフォール：主役の値（focusValue）は1つだけにする");
    const fvT = (x, y, w, text, valign = "bottom") => T(s, text, { x, y, w, h: 0.62, fontSize: FOCUS_FS, bold: true, color: C.emph, align: "center", valign, wrap: false });
    const fmt = o.fmt || (v => v.toLocaleString()), base0 = o.base ?? BOTTOM - 0.35, top = o.top ?? 2.95, lblW = o.labelW ?? 2.55;
    const x0 = o.x ?? M, wAll = o.w ?? BW, pre = o.before ? 1 : 0, n = o.steps.length + pre + 1;
    const total = o.total.parts.reduce((a, p) => a + p.value, 0), stepSum = o.steps.reduce((a, p) => a + p.value, 0);
    if (Math.abs(total - stepSum) > 0.5) warn(`ウォーターフォールの累計（${total}）と増分の合計（${stepSum}）が一致しない`);
    const unit = (base0 - top) / Math.max(total, o.before ? o.before.value : 0), slot = (wAll - lblW) / n, bw = Math.min(slot * 0.62, 1.3);
    const colX = i => x0 + i * slot + (slot - bw) / 2;
    const xLabel = (i, label, sub, bold) => T(s, sub ? [{ text: label, options: { breakLine: true, bold } }, { text: sub, options: { fontSize: FS.note, color: C.gray } }] : [{ text: label, options: { bold } }], { x: x0 + i * slot, y: base0 + 0.06, w: slot, h: 0.44, fontSize: FS.body, color: bold ? C.ink : C.soft, align: "center" });
    if (o.before) {
      const v = o.before.value;
      rect(s, colX(0), base0 - v * unit, bw, v * unit, null, { color: C.soft, width: 1, dashType: "dash" });
      T(s, fmt(v), { x: colX(0) - 0.2, y: base0 - v * unit - 0.34, w: bw + 0.4, h: 0.32, fontSize: FS.body, bold: true, color: C.soft, align: "center" });
      xLabel(0, o.before.label, o.before.sub);
    }
    let cum = 0;
    o.steps.forEach((st, j) => {
      const i = j + pre, y1 = base0 - (cum + st.value) * unit;
      rect(s, colX(i), y1, bw, st.value * unit, st.focusValue ? C.emph : C.gray);
      if (st.focusValue) { const t = `+${fmt(st.value)}`, ww = Math.max(bw + 0.4, textW(t, FOCUS_FS) + 0.2); fvT(colX(i) + bw / 2 - ww / 2, y1 - 0.66, ww, t); if (textW(t, FOCUS_FS) > slot + 0.3) warn(`ウォーターフォール：主役の値（${t}）が棒の間に入らない`); }
      else T(s, `+${fmt(st.value)}`, { x: colX(i) - 0.2, y: y1 - 0.34, w: bw + 0.4, h: 0.32, fontSize: FS.body, bold: true, align: "center" });
      cum += st.value;
      hline(s, colX(i) + bw, base0 - cum * unit, slot - bw);
      xLabel(i, st.label, st.sub);
    });
    // 累計の棒（区分色で積む）と、右のラベル（細い区分のラベルが重ならないよう、下から順に間隔を空ける）
    let acc = 0;
    const lab = [];
    o.total.parts.forEach(p => {
      const y = base0 - (acc + p.value) * unit, h = p.value * unit;
      rect(s, colX(n - 1), y, bw, h, p.focusValue ? C.emph : p.color, { color: C.white, width: 0.5 });
      lab.push({ p, cy: y + h / 2, hh: p.focusValue ? FOCUS_FS / 144 + 0.02 : 0.12 });
      acc += p.value;
    });
    for (let k = 1; k < lab.length; k++) { const g = lab[k - 1].hh + lab[k].hh; if (lab[k - 1].cy - lab[k].cy < g) lab[k].cy = lab[k - 1].cy - g; }
    const labW = lblW + slot - bw - 0.2;
    lab.forEach(({ p, cy }) => {
      if (!p.focusValue) return T(s, [{ text: p.name, options: { color: C.soft } }, { text: `　${fmt(p.value)}`, options: { bold: true } }], { x: colX(n - 1) + bw + 0.12, y: cy - 0.11, w: labW, h: 0.22, fontSize: FS.note, valign: "middle", wrap: false });
      T(s, [{ text: p.name, options: { fontSize: FS.note, color: C.soft } }, { text: "　", options: { fontSize: FS.note } }, { text: fmt(p.value), options: { fontSize: FOCUS_FS, bold: true, color: C.emph } }], { x: colX(n - 1) + bw + 0.12, y: cy - 0.31, w: labW, h: 0.62, valign: "middle", wrap: false });
      if (textW(`${p.name}　`, FS.note) + textW(fmt(p.value), FOCUS_FS) > labW + 0.05) warn(`ウォーターフォール：主役の値（${p.name}　${fmt(p.value)}）が右のラベルの幅に入らない（labelW を広げる）`);
    });
    const topLab = lab.length ? lab[lab.length - 1] : null;
    if (topLab && topLab.cy - topLab.hh < top - 0.45) warn("ウォーターフォール：累計の右のラベルが図の上端を超える（区分を減らすか、top を下げる）");
    if (o.total.focusValue) fvT(colX(n - 1) + bw / 2 - Math.max(bw + 0.6, textW(fmt(total), FOCUS_FS) + 0.2) / 2, base0 - total * unit - 0.66, Math.max(bw + 0.6, textW(fmt(total), FOCUS_FS) + 0.2), fmt(total));
    else T(s, fmt(total), { x: colX(n - 1) - 0.3, y: base0 - total * unit - 0.37, w: bw + 0.6, h: 0.34, fontSize: FS.head, bold: true, align: "center" });
    xLabel(n - 1, o.total.label || "累計", null, true);
    hline(s, x0, base0, wAll);
  };

  // 面積棒：横幅＝期間、高さ＝月あたりの額、面積が段階の総額。段階の境に審査の赤破線、上に段階名、下に内訳の行（4.6 面積棒・3章「費用・期間」）
  //   o.stages: [{ name, months, parts: { 区分キー: 月額 } }], o.cats: [{ key, name, color }], o.gates: [{ at: 段階の番号（その段階の左端）, label, strong? }]
  //   o.rows: [{ label, cells: [段階ごとの文字] }]（内訳）, o.onFace（After の面の上ならブロックを白）, o.headY, o.fmt, o.unitFmt
  const areaBars = (s, o) => {
    const fmt = o.fmt || (v => v.toLocaleString()), lw = o.labelW ?? 1.35, x0 = M + lw + 0.08, x1 = o.x1 ?? W - M;
    const st = o.stages, months = st.map(a => a.months), total = months.reduce((a, b) => a + b, 0), um = (x1 - x0) / total;
    const sx = j => x0 + months.slice(0, j).reduce((a, b) => a + b, 0) * um, sw = j => months[j] * um;
    const fill = o.onFace ? C.white : TINT, hy = o.headY ?? 2.7, hh = 0.5, yT = hy + hh + 0.88, yB = yT + 1.45;
    st.forEach((a, j) => { rect(s, sx(j) + 0.03, hy, sw(j) - 0.06, hh, fill); T(s, `${j + 1}. ${a.name}`, { x: sx(j) + 0.06, y: hy, w: sw(j) - 0.12, h: hh, fontSize: FS.body, bold: true, align: "center", valign: "middle", lineSpacingMultiple: 1.05 }); });
    const mon = a => Object.values(a.parts).reduce((x, y) => x + y, 0), max = Math.max(...st.map(mon)), k = (yB - yT) / max;
    s.addShape(SH.LINE, { x: x0, y: yB, w: x1 - x0, h: 0, line: { color: C.gray, width: 0.75 } });
    st.forEach((a, j) => {
      let yy = yB;
      o.cats.forEach(c => { const v = a.parts[c.key] || 0; if (!v) return; const h = v * k; rect(s, sx(j) + 0.03, yy - h, sw(j) - 0.06, h, c.color); yy -= h; });
      const vh = lineH(FS.body, 1.05) + lineH(FS.note, 1.05);
      T(s, [{ text: fmt(mon(a) * a.months), options: { fontSize: FS.body, bold: true, breakLine: true } }, { text: (o.unitFmt || (v => `月${fmt(v)}`))(mon(a)), options: { fontSize: FS.note, color: C.soft } }], { x: sx(j), y: yy - vh - 0.03, w: sw(j), h: vh, align: "center", valign: "bottom", lineSpacingMultiple: 1.05 });
    });
    // 読み方と凡例は左のラベル列に置く。凡例の上下の並びを積みの並びと同じにする
    const legP = Math.max(0.27, lineH(FS.note, 1.0)), howH = 2 * lineH(FS.note, 1.1);
    T(s, o.howTo || "高さ：月あたりの費用\n横幅：期間", { x: M, y: yB - legP * o.cats.length - howH - 0.08, w: Math.max(lw, 2.4), h: howH, fontSize: FS.note, color: C.gray, lineSpacingMultiple: 1.1 });
    o.cats.forEach((c, i) => {
      const ly = yB - legP * (i + 1) + 0.03;
      rect(s, M, ly + 0.04, 0.16, 0.16, c.color);
      T(s, c.name, { x: M + 0.24, y: ly, w: lw - 0.24, h: 0.24, fontSize: FS.note, color: C.soft, valign: "middle", wrap: false });
    });
    let y = yB + 0.14;
    (o.rows || []).forEach(r => {
      const h = Math.max(r.h ?? 0.8, Math.max(...r.cells.map((t, j) => lineCount(t, FS.note, sw(j) - 0.24))) * lineH(FS.note, 1.2) + 0.08); // 行は文字が入る高さより低くしない
      T(s, r.label, { x: M, y, w: lw, h, fontSize: FS.body, bold: true, valign: "middle" });
      r.cells.forEach((t, j) => { rect(s, sx(j) + 0.03, y, sw(j) - 0.06, h, fill); T(s, t, { x: sx(j) + 0.12, y, w: sw(j) - 0.24, h, fontSize: FS.note, color: C.soft, valign: "middle", lineSpacingMultiple: 1.2 }); checkWrap(t, FS.note, sw(j) - 0.24); if (lineCount(t, FS.note, sw(j) - 0.24) * lineH(FS.note, 1.2) > h - 0.06) warn(`面積棒の内訳があふれる（${t.slice(0, 12)}…）`); });
      y += h + 0.06;
    });
    const bottom = y - 0.06;
    (o.gates || []).forEach(g => {
      const gx = sx(g.at), ly = hy - 0.4;
      s.addShape(SH.LINE, { x: gx, y: ly + 0.26, w: 0, h: bottom - ly - 0.26, line: { color: C.accent, width: g.strong ? 3 : 1.5, dashType: "dash" } });
      T(s, g.label, { x: gx - 1.0, y: ly - 0.12, w: 2.0, h: 0.36, fontSize: FS.note, bold: true, color: C.accent, align: "center", valign: "bottom", wrap: false });
    });
    return { sx, sw, yT, yB, bottom };
  };

  // 積み上げ棒（ネイティブのグラフ。目盛り EEEEEE、文字 gray）。o.highlight で棒ごとの強調を書き出し後の XML 処理で入れる（4.6・元データ7 focusChart）
  //   series: [{ name, values, color }], o: { x, y, w, h, labels, max, major, legend, grouping, highlight: { series, points: [番号], mode: "frame"|"fade", alpha } }
  //   強調の既定は「対象の棒に赤の太枠 2.5pt、他の色は元のまま」（全要素を読ませる図では他を薄くしない。2.3・10.3）
//   mode: "focus" … 問題ではない主役（最新の年、最多の区分）の棒だけ濃紺（C.ink）で塗る。他の棒は灰（C.bar）にしておく
//   mode: "frame" … 問題を示す棒に赤の太枠。mode: "fade" … 多数から選んだ盤面に限り、他を薄くする
  const barChart = (s, series, o) => {
    chartPage();
    const rec = SLIDES[SLIDES.length - 1], nth = rec.charts++;
    const natural = (o.w - 0.6) / (o.labels.length * ((o.grouping === "clustered" ? series.length : 1) + (o.gapPct ?? 60) / 100));
    if (natural > 1.3) warn(`棒グラフ（ネイティブ）：棒が太くなる（およそ ${natural.toFixed(1)}）。w を狭める（区分が少ないグラフは、幅いっぱいにしない）`);
    s.addChart(pres.charts.BAR, series.map(sr => ({ name: sr.name, labels: o.labels, values: sr.values })), {
      x: o.x, y: o.y, w: o.w, h: o.h, barDir: "col", barGrouping: o.grouping || "stacked", barGapWidthPct: o.gapPct ?? 60,
      chartColors: series.map(sr => sr.color),
      showLegend: o.legend !== false, legendPos: "b", legendFontSize: FS.note, legendFontFace: F, legendColor: C.soft,
      catAxisLabelFontSize: FS.note, valAxisLabelFontSize: FS.note, catAxisLabelFontFace: F, valAxisLabelFontFace: F,
      catAxisLabelColor: C.gray, valAxisLabelColor: C.gray, catAxisLineShow: true, catAxisLineColor: C.line, valAxisLineShow: false,
      ...(o.max ? { valAxisMaxVal: o.max, valAxisMinVal: 0 } : {}), ...(o.major ? { valAxisMajorUnit: o.major } : {}),
      valGridLine: o.values ? { style: "none" } : { color: C.grid, size: 0.5 }, catGridLine: { style: "none" }, showTitle: false,
      // o.values：棒の外（上端）に値を出す { format, fontSize, color, bold }。値を出したら縦軸と目盛り線は消す（同じ情報を二重に置かない）
      //   棒の中には入れない。積み上げでは上端の外に置けないので、単系列（grouping: "clustered"）だけに使う
      ...(o.values ? { valAxisHidden: true, showValue: true, dataLabelPosition: "outEnd", dataLabelFontFace: F, dataLabelFontSize: o.values.fontSize ?? FS.block, dataLabelColor: o.values.color ?? C.soft, dataLabelFontBold: !!o.values.bold, dataLabelFormatCode: o.values.format ?? "#,##0" } : {}),
    });
    if (o.highlight) CHART_PATCH.push({ slide: SLIDES.length, nth, ...o.highlight });
  };
  // 棒ごとの強調の XML（pptxgenjs では書けない dPt を足す。4.6）
  const patchChartXml = (xml, h) => xml.replace(/<c:ser>[\s\S]*?<\/c:ser>/g, ser => {
    const color = (ser.match(/<c:spPr>[\s\S]*?<a:srgbClr val="([0-9A-Fa-f]{6})"/) || [])[1];
    const count = +((ser.match(/<c:val>[\s\S]*?<c:ptCount val="(\d+)"/) || [])[1] || 0);
    const isTarget = new RegExp(`<c:tx>[\\s\\S]*?<c:v>${h.series.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}</c:v>`).test(ser);
    const dPt = (i, fill, ln) => `<c:dPt><c:idx val="${i}"/><c:invertIfNegative val="0"/><c:bubble3D val="0"/><c:spPr>${fill}${ln}</c:spPr></c:dPt>`;
    let dpts = "";
    if ((h.mode || "frame") === "frame") {
      if (isTarget) dpts = h.points.map(i => dPt(i, `<a:solidFill><a:srgbClr val="${color}"/></a:solidFill>`, `<a:ln w="31750"><a:solidFill><a:srgbClr val="${C.accent}"/></a:solidFill></a:ln>`)).join("");
    } else if (h.mode === "focus") {
      // 主役の棒だけ濃紺で塗る（データの印は濃紺で塗ってよい。他の棒は系列の色のまま）。問題を示す棒は frame（赤枠）を使う
      if (isTarget) dpts = h.points.map(i => dPt(i, `<a:solidFill><a:srgbClr val="${h.color || C.emph}"/></a:solidFill>`, "<a:ln><a:noFill/></a:ln>")).join("");
    } else {
      const keep = new Set(isTarget ? h.points : []);
      dpts = Array.from({ length: count }, (_, i) => i).filter(i => !keep.has(i)).map(i => dPt(i, `<a:solidFill><a:srgbClr val="${color}"><a:alpha val="${(h.alpha ?? 22) * 1000}"/></a:srgbClr></a:solidFill>`, "<a:ln><a:noFill/></a:ln>")).join("");
    }
    return dpts ? ser.replace(/(<c:invertIfNegative val="0"\/>)/, `$1${dpts}`) : ser;
  });

  // ---- 図形で描く数値の図 ----
  // 数の書式：負は全角のマイナス、plus で正に「+」を付ける
  const fmtNum = (v, o = {}) => { const a = Math.abs(v), t = o.digits != null ? a.toFixed(o.digits) : a.toLocaleString(); return (v < 0 ? "−" : o.plus && v > 0 ? "+" : "") + t + (o.unit || ""); };
  // 図の左上の見出し：何の数値か（単位・時点）を FS.body（16pt）soft で、全ページ同じ位置（x M、y 本文の上端）に置く。凡例はその右に同じ行で並べる
  //   legend: [[名前, 色], ...]（棒・帯と同じ並び）。区分名を図の中（右端の棒の横など）に置けるなら凡例は置かない
  // ---- 棒の作法（太さ・間・名前の位置。出典は references/charts.md の「棒の作法」） ----
  // 項目名の欄の幅：中身から決める（o.labelW を渡したらその幅）。名前は欄の右にそろえ、軸のすぐ横に置く
  const labelCol = (name, labels, given, pt = FS.body, max = 4.2) => {
    const ls = labels.map(t => (Array.isArray(t) ? t : [t, pt]));
    const lw = given ?? Math.max(0.6, Math.min(max, Math.max(0, ...ls.map(([t, p]) => textW(t, p))) + 0.24));
    const over = ls.find(([t, p]) => textW(t, p) > lw - 0.1);
    if (over) warn(`${name}：項目名が欄に入らない（${over[0]}）。名前を縮めるか、labelW を広げる`);
    return lw;
  };
  // 横に伸びる棒の太さと行の間隔：文字の大きさから決める（太さは本文の高さの1.15〜1.6倍、間は太さの半分ほど）
  const barRows = (name, room, o) => {
    const t = FS.body / 72;
    const bh = o.bh ?? (o.pitch != null ? Math.min(o.pitch / 1.55, t * 1.6) : Math.max(t * 1.15, Math.min(t * 1.6, room / 1.55)));
    const pitch = o.pitch ?? Math.max(bh * 1.25, Math.min(room, bh * 1.55));
    if (o.bh != null && o.bh > t * 1.65) warn(`${name}：棒が太い（bh ${o.bh}）。指定（bh・pitch）を外すと、文字の大きさから決まる（余白は残してよい）`);
    else if (o.bh != null || o.pitch != null) {
      const r = (pitch - bh) / bh;
      if (r < 0.2 || r > 1.05) warn(`${name}：棒の間（${(pitch - bh).toFixed(2)}）が、棒の太さ（${bh.toFixed(2)}）の${r < 0.2 ? "4分の1より狭い" : "幅より広い"}。指定（bh・pitch）を外すと、文字の大きさから決まる`);
    }
    return { bh, pitch };
  };
  // 縦の棒の太さと区分の幅：区分が少ないときは図を狭めて中央に置き、区分の間を棒より広げない。x0 と x1 の両方を渡したときは、その幅を使う
  const colSlot = (name, { x0, x1, n, nS = 1, gap = 0, bw, needW = 0, fill = false }) => {
    const avail = (x1 - x0) / n;
    const w = bw ?? Math.max(0.2, Math.min(nS > 1 ? 0.65 : 1.0, (avail * 0.65 - gap * (nS - 1)) / nS));
    const gw = w * nS + gap * (nS - 1);
    const slot = fill ? avail : Math.min(avail, Math.max(gw * (nS > 1 ? 1.5 : 1.6), needW));
    if (bw != null && bw > 1.05) warn(`${name}：棒が太い（bw ${bw}）。指定を外すと、区分の数から決まる`);
    const xs = x0 + ((x1 - x0) - slot * n) / 2;
    return { bw: w, slot, x0: xs, x1: xs + slot * n };
  };

  const figHead = (s, text, legend, o = {}) => {
    if (legend && legend.length) s._figLegend = true;
    const y = o.y ?? Y.body, x = o.x ?? M;
    // 何の数値か・単位・時点は、要る情報なので、読める大きさと色で書く（2026-11-24。9.5pt の灰をやめた）
    T(s, text, { x, y: y - 0.04, w: 9, h: 0.32, fontSize: FS.body, color: C.soft, valign: "middle", wrap: false });
    let cx = x + textW(text, FS.body) + 0.45;
    (legend || []).forEach(([name, color]) => {
      rect(s, cx, y + 0.06, 0.14, 0.14, color);
      T(s, name, { x: cx + 0.2, y, w: 3, h: 0.26, fontSize: FS.body, color: C.soft, valign: "middle", wrap: false });
      cx += 0.2 + textW(name, FS.body) + 0.35;
    });
  };
  // 縦棒（図形）：正負の値、1つの区分に複数の系列（対の棒）、棒の外の値、主役の塗りを1つの関数で描く
  //   o.cats: [区分名], o.series: [{ name, color, values, main? }]（main＝主役の系列。値を太字・一段大きく）
  //   o.focus: [[系列の番号, 区分の番号], ...]（その棒だけ濃紺にし、値を主役の値 36pt で棒の真上に描く。単系列でも対の棒でも使える。
  //     対の棒で文字が隣の棒・値に重なるときは、重ならない所まで上げる。呼ぶ側で値を空にして focusValue を手で置かない）
  //   o.x0, o.x1（既定 M+0.7〜W−M−0.6）, o.top, o.bottom（棒の領域。既定 2.85〜5.85）, o.min, o.max（既定はデータから）, o.bw, o.gap, o.fmt
  //   o.cagr: { from, to, years?, label?, j? }（年平均成長率。from・to の棒を上で矢印につなぎ、「年平均 +x%（from→to の n 年）」を書く。率は値から計算する。cagrBars）
  //   戻り値: { barX(j, i), barTop(j, i), cx(i), base, k } … 重ね技で前面から指す位置に使う
  const colBars = (s, o) => {
    chartPage();
    const ser = o.series, nC = o.cats.length, nS = ser.length;
    const all = ser.flatMap(sr => sr.values);
    const G = o.cagr, min = o.min ?? Math.min(0, ...all) * 1.25;
    let max = o.max ?? Math.max(0, ...all) * (G ? 1.55 : 1.12);
    if (G && min < 0) warn("年平均成長率（cagr）は、0 から始まる正の値の棒で見せる");
    const top = o.top ?? 2.85;
    const bottom = BOTTOM < 6.6 ? Math.min(o.bottom ?? BOTTOM - 0.6, BOTTOM - 0.6) : o.bottom ?? BOTTOM - 0.6;
    let k = (bottom - top) / (max - min), base = top + max * k;
    const fmt = o.fmt || (v => fmtNum(v, { plus: min < 0 }));
    // 系列が2つ以上のときは、系列の名前を棒のそば（最後の区分の右、軸の近く）に部品が書く。見出しの行（figHead）に凡例を置いたときは書かない
    const keys = nS > 1 && o.legend !== false && !s._figLegend && ser.every(sr => sr.name);
    if (nS > 1 && s._figLegend) warn("縦棒：凡例が見出しの行にあって棒から遠い。figHead の凡例を消すと、系列の名前を部品が棒のそばに書く");
    const keyW = keys ? Math.max(...ser.map(sr => textW(sr.name, FS.body))) + 0.55 : 0;
    const gap = o.gap ?? 0.08;
    const labW = Math.max(...o.cats.map(c => textW(c, FS.body))) + 0.25;
    const valW = o.values === false ? 0 : Math.max(...all.map(v => textW(fmt(v), FS.head))) + 0.15;
    const { bw, slot, x0, x1 } = colSlot("縦棒", { x0: o.x0 ?? M + 0.7, x1: (o.x1 ?? W - M - 0.6) - keyW, n: nC, nS, gap, bw: o.bw,
      needW: Math.max(labW, nS === 1 ? valW : 0, o.focus ? textW(fmt(all.reduce((a, b) => Math.abs(b) > Math.abs(a) ? b : a, 0)), o.focusFs ?? FOCUS_FS) + 0.15 : 0), fill: o.x0 != null && o.x1 != null });
    const gw = bw * nS + gap * (nS - 1);
    rangeBar("縦棒", all.map(v => fmt(v)));
    const focus = new Set((o.focus || []).map(([j, i]) => `${j}:${i}`));
    const barX = (j, i) => x0 + i * slot + (slot - gw) / 2 + j * (bw + gap);
    const barTop = (j, i) => { const v = ser[j].values[i]; return v >= 0 ? base - v * k : base; };
    // 棒と値の文字の範囲（文字の幅と1行の高さ）。主役の値（focus）の位置を決めるのに使う
    let BAR = {}, LAB = {}, lifted = [];
    const LIM = Math.max(BODY + 0.3, top - 0.45); // 主役の値の字面の上端の限り（図の見出しの行に入らない）
    const layout = () => {
    k = (bottom - top) / (max - min); base = top + max * k; BAR = {}; LAB = {}; lifted = [];
    ser.forEach((sr, j) => sr.values.forEach((v, i) => {
      const h = Math.max(Math.abs(v) * k, 0.02), y = v >= 0 ? base - h : base, f = focus.has(`${j}:${i}`), fs = f ? (o.focusFs ?? FOCUS_FS) : sr.main ? FS.head : FS.body;
      // 文字の範囲は字面（大きさ×1.0）。1行の高さ（×1.2）との差 off を上下に残す
      const tw = textW(fmt(v), fs), th = fs * 1.2 / 72, gh = fs / 72, off = (th - gh) / 2, cx = barX(j, i) + bw / 2;
      BAR[`${j}:${i}`] = { x: barX(j, i), y, w: bw, h };
      LAB[`${j}:${i}`] = { x: cx - tw / 2, y: (v >= 0 ? y - 0.04 - th : y + h + 0.04) + off, w: tw, h: gh, off, fs, f, v };
    }));
    // 主役の値（36pt）は、その棒の真上（負なら真下）に部品が描く。呼ぶ側で値を空にして focusValue を手で置かない
    //   文字が棒より広く、隣の棒・隣の値に重なるとき（対の棒で起きる）は、重ならない所まで上（負なら下）へずらす
    if (o.values !== false) Object.keys(LAB).filter(key => LAB[key].f).forEach(key => {
      const L = LAB[key], obs = [...Object.keys(BAR).filter(q => q !== key).map(q => BAR[q]), ...Object.keys(LAB).filter(q => q !== key).map(q => LAB[q])];
      const y0L = L.y;
      if (liftClear(L, obs, L.v >= 0 ? -1 : 1) > 0.005) lifted.push({ L, y0L });
    });
    };
    layout();
    // 上げた主役の値が図の見出しの行に入るときは、o.max を渡していなければ、上の余白を広げて描き直す（棒を少し低くする）
    for (let n = 0; n < 12 && o.max == null && lifted.some(q => q.L.v >= 0 && q.L.y < LIM); n++) { max *= 1.06; layout(); }
    lifted.forEach(({ L, y0L }) => {
      {
        if (L.y < LIM || L.y + L.h > BOTTOM) warn(`縦棒の主役の値（${fmt(L.v)}）を、隣の棒・値に重ならない所まで${L.v >= 0 ? "上" : "下"}げたら、図の${L.v >= 0 ? "上端" : "下端"}を超えた（o.max${L.v >= 0 ? "" : "・o.min"} を広げるか、bw を細くして隣と離す）`);
        else if (Math.abs(L.y - y0L) > 0.9) warn(`縦棒の主役の値（${fmt(L.v)}）が、隣の値に重ならないよう棒から ${Math.abs(L.y - y0L).toFixed(2)} 離れた。棒との対応が分かりにくいので、画像で確かめる`);
      }
    });
    ser.forEach((sr, j) => sr.values.forEach((v, i) => {
      const key = `${j}:${i}`, { y, h } = BAR[key], L = LAB[key], f = L.f, strong = sr.main || f;
      rect(s, barX(j, i), y, bw, h, f ? C.emph : (focus.size && sr.color === C.emph ? C.gray : sr.color));
      // 主役の棒（focus）の値は、色だけでなく大きさで示す（既定 36pt。o.focusFs）
      if (o.values !== false) {
        const bxW = Math.max(bw + 0.8, L.w + 0.2), bxH = f ? 0.62 : 0.36;
        T(s, fmt(v), { x: barX(j, i) + bw / 2 - bxW / 2, y: v >= 0 ? L.y + L.h + L.off - bxH : L.y - L.off, w: bxW, h: bxH, fontSize: L.fs, bold: strong, color: f ? C.emph : strong ? C.ink : C.soft, align: "center", valign: v >= 0 ? "bottom" : "top", wrap: false });
      }
    }));
    o.cats.forEach((c, i) => T(s, c, { x: x0 + i * slot, y: bottom + 0.3, w: slot, h: 0.3, fontSize: FS.body, color: C.soft, align: "center", wrap: false }));
    hline(s, x0, base, x1 - x0, { color: C.gray });
    if (keys) ser.forEach((sr, j) => {
      const y = base - 0.42 * (nS - j) - 0.08;
      rect(s, x1 + 0.2, y + 0.12, 0.16, 0.16, focus.size && sr.color === C.emph ? C.gray : sr.color);
      T(s, sr.name, { x: x1 + 0.42, y, w: keyW - 0.4, h: 0.4, fontSize: FS.body, color: sr.main ? C.ink : C.soft, bold: !!sr.main, valign: "middle", wrap: false });
    });
    if (bottom + 0.6 > BOTTOM + 0.05) warn(`縦棒の区分名が本文の下端を超える（bottom ${bottom}）`);
    // 年平均成長率：最初と最後の棒を、値のラベルの上で矢印につなぐ（cagrBars）。率は値から計算し、期間と一緒に書く
    if (G) {
      const j = G.j ?? 0, v0 = ser[j].values[G.from], v1 = ser[j].values[G.to], n = G.years ?? (G.to - G.from);
      if (!(v0 > 0 && v1 > 0 && n > 0)) warn("年平均成長率（cagr）：最初と最後の値は正、期間は 1 年以上にする");
      else {
        const r = Math.pow(v1 / v0, 1 / n) - 1, rate = `${r >= 0 ? "+" : "−"}${(Math.abs(r) * 100).toFixed(1)}%`;
        const xa = barX(j, G.from) + bw / 2, xb = barX(j, G.to) + bw / 2, ytop = Math.min(barTop(j, G.from), barTop(j, G.to)), yl = ytop - 0.75;
        polyArrow(s, [[xa, barTop(j, G.from) - 0.45], [xa, yl], [xb, yl], [xb, barTop(j, G.to) - 0.45]], { color: C.emph, width: 2 });
        const lab = G.label ?? `${o.cats[G.from]}→${o.cats[G.to]}の${n}年`;
        const tw = textW(`年平均 ${rate}`, FOCUS_FS) + textW(`（${lab}）`, FS.body) + 0.3;
        T(s, [{ text: "年平均 ", options: { fontSize: FS.block, bold: true, color: C.emph } }, { text: rate, options: { fontSize: FOCUS_FS, bold: true, color: C.emph } }, { text: `（${lab}）`, options: { fontSize: FS.body, color: C.soft } }],
          { x: (xa + xb) / 2 - tw / 2, y: yl - 0.66, w: tw, h: 0.62, align: "center", valign: "bottom", fill: { color: C.white }, wrap: false });
        if (yl - 0.66 < top - 0.3) warn("年平均成長率（cagr）の文字が図の上端を超える。o.max を上げる");
      }
    }
    return { barX, barTop, cx: i => x0 + i * slot + slot / 2, base, k, bw };
  };
  // 横棒（図形）：0の線から左右に伸ばす。項目の名前が長い、正負が混ざる、全体と内訳を並べるときに使う
  //   o.rows: [{ label, value, strong?, focusValue?, problem? } | null（行の間を空ける）], o.groups: [{ from, to, text }]（連続した行をまとめた注記。空いた側に縦線と文字）
  //     focusValue: true＝主役の行（strong と同じ塗り）。その値を主役の値 36pt で棒の先に部品が描く（problem も付けると赤）。値の列はその幅だけ広げる
  //   o.y0（既定 2.6）, o.fmt。項目名は欄の右にそろえて軸のすぐ横に置く。欄の幅・棒の太さ・間は部品が決める（o.labelW・o.bh・o.pitch は、別のページと図をそろえるときだけ）
  const hBars = (s, o) => {
    { // 割合（シェア・構成比）の合計を、内訳と同じ棒に並べたら知らせる（2026-11-25）。予算の「合計と内訳」のような金額は対象外
      const rws = (o.rows || []).filter(Boolean), isTot = r => /合計|全体|上位\s*\d|計$/.test(String(r.label || "")), parts = rws.filter(r => !isTot(r));
      if (rws.some(isTot) && parts.length > 1 && parts.every(r => Math.abs(r.value) <= 100) && parts.reduce((a, r) => a + Math.abs(r.value), 0) <= 100.5) warn("割合の合計（上位の合計・全体）を、内訳と同じ棒に並べない。全体に占める割合は円（pieChart）で描き、合計は円の中に書く");
    }
    chartPage();
    const rows = o.rows, vals = rows.filter(Boolean).map(r => r.value), fmt = o.fmt || (v => fmtNum(v, { plus: true }));
    rangeBar("横棒", vals.map(v => fmt(v)));
    if (rows.filter(r => r && r.focusValue).length > 1) warn("横棒：主役の値（focusValue）は1行だけにする");
    // 主役の値（36pt）が入るよう、その側の値の列を広げる
    const fvW = side => Math.max(0, ...rows.filter(r => r && r.focusValue && (r.value >= 0) === side).map(r => textW(fmt(r.value), FOCUS_FS) + 0.14));
    const maxP = Math.max(0, ...vals), maxN = Math.max(0, ...vals.map(v => -v)), valW = Math.max(1.05, ...vals.map(v => textW(fmt(v), FS.head) + 0.12)); // 値の欄は中身が入る幅
    const lw = labelCol("横棒", rows.filter(Boolean).map(r => [r.label, r.strong || r.focusValue ? FS.head : FS.body]), o.labelW);
    const left = M + lw + (maxN ? Math.max(valW, fvW(false)) : 0), right = W - M - Math.max(valW, fvW(true));
    const k = (right - left) / (maxP + maxN || 1), Z = left + maxN * k;
    const ys = []; let y = o.y0 ?? 2.6;
    let { bh, pitch } = barRows("横棒", (BOTTOM - y) / rows.reduce((a, r) => a + (r ? 1 : 0.5), 0), o);
    rows.forEach(r => { ys.push(y); y += r ? pitch : pitch / 2; });
    const OBS = []; // 棒と値の文字の範囲（主役の値の重なりの点検）
    rows.forEach((r, i) => {
      if (!r) return;
      const w = Math.abs(r.value) * k, yy = ys[i], fv = !!r.focusValue, st = !!r.strong || fv;
      rect(s, r.value >= 0 ? Z : Z - w, yy, w, bh, st ? C.emph : (r.color || C.bar));
      T(s, r.label, { x: M, y: yy, w: lw - 0.12, h: bh, fontSize: st ? FS.head : FS.body, bold: st, color: st ? C.ink : C.soft, align: "right", valign: "middle", wrap: false });
      const vt = fmt(r.value), pos = r.value >= 0;
      if (fv) {
        const tw = textW(vt, FOCUS_FS), cy = yy + bh / 2, tr = { x: pos ? Z + w + 0.1 : Z - w - 0.1 - tw, y: cy - FOCUS_FS / 144, w: tw, h: FOCUS_FS / 72, i, fv: true };
        focusValue(s, vt, { x: pos ? tr.x : tr.x - 0.1, y: cy - 0.25, w: tw + 0.1, align: pos ? "left" : "right" }, { problem: r.problem });
        OBS.push(tr, { x: r.value >= 0 ? Z : Z - w, y: yy, w, h: bh, i });
      } else {
        T(s, vt, { x: pos ? Z + w + 0.08 : Z - w - 0.08 - valW, y: yy, w: valW, h: bh, fontSize: st ? FS.head : FS.body, bold: st, color: st ? C.ink : C.soft, align: pos ? "left" : "right", valign: "middle", wrap: false });
        const fs = st ? FS.head : FS.body, tw = textW(vt, fs);
        OBS.push({ x: pos ? Z + w + 0.08 : Z - w - 0.08 - tw, y: yy + bh / 2 - fs / 144, w: tw, h: fs / 72, i }, { x: pos ? Z : Z - w, y: yy, w, h: bh, i });
      }
    });
    OBS.filter(b => b.fv).forEach(b => {
      if (OBS.some(q => q.i !== b.i && boxHit(b, q))) warn(`横棒の主役の値（${fmt(rows[b.i].value)}）が、上下の行の棒・値に重なる（pitch を広げるか、主役の行を棒の長い順の先頭に置く）`);
      if (b.x + b.w > W - M + 0.05 || b.x < M + lw) warn(`横棒の主役の値（${fmt(rows[b.i].value)}）が図の幅に入らない`);
    });
    const last = ys[ys.length - 1] + bh;
    s.addShape(SH.LINE, { x: Z, y: ys[0] - 0.1, w: 0, h: last + 0.1 - (ys[0] - 0.1), line: { color: C.gray, width: 0.75 } });
    // 群のまとめの注記：負の群は 0 の線の右の空きに、正の群は群のいちばん長い棒（と値）の右に置く（左に置くと項目名に重なり、ページの外へ出る）
    (o.groups || []).forEach(g => {
      const gy = ys[g.from], gh = ys[g.to] + bh - gy, neg = rows[g.from].value < 0;
      const ends = rows.slice(g.from, g.to + 1).filter(Boolean).map(r => r.value >= 0 ? Math.abs(r.value) * k + 0.1 + textW(fmt(r.value), r.focusValue ? FOCUS_FS : r.strong ? FS.head : FS.body) : 0);
      const gx = neg ? Z + 0.2 : Z + Math.max(0, ...ends) + 0.25, tw = W - M - (gx + 0.15);
      s.addShape(SH.LINE, { x: gx, y: gy, w: 0, h: gh, line: { color: C.soft, width: 1 } });
      T(s, g.text, { x: gx + 0.15, y: gy, w: Math.max(tw, 0.5), h: gh, fontSize: FS.block, color: C.soft, align: "left", valign: "middle" });
      if (lineCount(g.text, FS.block, Math.max(tw, 0.5)) * lineH(FS.block) > gh + 0.1) warn(`横棒の群の注記（${String(g.text).slice(0, 12)}…）が、群の右の空きに入らない（注記を縮めるか、群の行を増やす）`);
    });
    if (last > BOTTOM + 0.05) warn(`横棒の下端 ${last.toFixed(2)} が本文の下端を超える`);
    return { ys, Z, k, bh };
  };
  // 帯でつなぐ積み上げ棒（配分の推移）：棒の高さは総量に比例、同じ区分を隣の期と帯でつなぐ。区分名は右端の棒の横に置き、凡例は置かない
  //   o.cats: [期], o.series: [{ name, color, values, main? }]（下から積む順。main の区分は名前を太字に）
  //   o.topLabel: i => 文字（棒の上に置く値。合計や主役の区分の割合など、リードで言う数値）, o.x0, o.x1, o.top, o.base, o.max, o.bw
  const ribbonBars = (s, o) => {
    chartPage();
    const n = o.cats.length, x0 = o.x0 ?? M + 0.6, x1 = o.x1 ?? W - M - 1.9, top = o.top ?? 2.6, base = o.base ?? BOTTOM - 0.4;
    const tot = i => o.series.reduce((a, sr) => a + sr.values[i], 0);
    const max = o.max ?? Math.max(...o.cats.map((_, i) => tot(i))) * 1.12, k = (base - top) / max, slot = (x1 - x0) / n, bw = o.bw ?? Math.min(0.62, slot * 0.55);
    const bx = i => x0 + i * slot + (slot - bw) / 2;
    const seg = o.cats.map((_, i) => { let y = base; return o.series.map(sr => { const h = sr.values[i] * k; const r = [y - h, y]; y -= h; return r; }); });
    for (let i = 0; i < n - 1; i++) {
      const xa = bx(i) + bw, xb = bx(i + 1), w = xb - xa;
      o.series.forEach((sr, j) => {
        const [at, ab] = seg[i][j], [bt, bb] = seg[i + 1][j], y = Math.min(at, bt), h = Math.max(ab, bb) - y;
        s.addShape(SH.CUSTOM_GEOMETRY, { x: xa, y, w, h, fill: { color: sr.color, transparency: sr.main ? 78 : 50 }, line: { type: "none" },
          points: [{ x: 0, y: at - y }, { x: w, y: bt - y }, { x: w, y: bb - y }, { x: 0, y: ab - y }, { close: true }] });
      });
    }
    o.cats.forEach((c, i) => {
      o.series.forEach((sr, j) => { const [t, b] = seg[i][j]; rect(s, bx(i), t, bw, b - t, sr.color, { color: C.white, width: 0.75 }); });
      if (o.topLabel) T(s, o.topLabel(i), { x: bx(i) - 0.3, y: seg[i][o.series.length - 1][0] - 0.32, w: bw + 0.6, h: 0.28, fontSize: FS.block, bold: true, align: "center", valign: "bottom", wrap: false });
      T(s, c, { x: x0 + i * slot, y: base + 0.1, w: slot, h: 0.3, fontSize: FS.body, color: C.soft, align: "center", wrap: false });
    });
    o.series.forEach((sr, j) => { const [t, b] = seg[n - 1][j]; T(s, sr.name, { x: bx(n - 1) + bw + 0.15, y: (t + b) / 2 - 0.14, w: 1.6, h: 0.28, fontSize: FS.body, color: sr.main ? C.ink : C.soft, bold: !!sr.main, valign: "middle", wrap: false }); });
    return { bx, seg, slot, bw, k, base };
  };
  // 要点と説明の2段：示唆・提言・申し送りのように、1項目に「言い切りの要点」と「その根拠の1文」があるときに使う（事実だけの列挙は bullets）
  //   items: [{ head, body, refKey? }]（refKey で右端に「根拠 P.n」）, o: { y0（2.45）, pitch（本文の下端までを等分、最大 1.1）, headFs（17）, bodyFs（15）, refPrefix }
  const points = (s, items, o = {}) => {
    const curTitle = (FOCUS[FOCUS.length - 1] || {}).title || "";
    if (/要約|サマリー/.test(curTitle)) warn("要約を points（見出しと補足の文の羅列）で描かない。summary を使う：決めてほしいこと（結論）を上に1つ、根拠を下に3つまで並べて線でつなぐ（2026-11-25）");
    const y0 = o.y0 ?? 2.45, pitch = o.pitch ?? Math.min(1.25, (BOTTOM - y0) / items.length), hf = o.headFs ?? 22, bf = o.bodyFs ?? FS.lead; // 見出し 22pt・根拠 20pt（案B）
    const w = BW - (items.some(it => it.refKey) ? 1.4 : 0);
    items.forEach((it, i) => {
      const y = y0 + i * pitch;
      if (textW(it.head, hf) > w) warn(`要点が1行に収まらない（${it.head}）`);
      T(s, it.head, { x: M, y, w, h: 0.42, fontSize: hf, bold: true, valign: "middle", wrap: false });
      if (it.body) { checkWrap(it.body, bf, w); T(s, it.body, { x: M, y: y + 0.48, w, h: 0.38 * Math.max(1, lineCount(it.body, bf, w)), fontSize: bf, color: C.soft, valign: "top", lineSpacingMultiple: 1.2 }); }
      if (it.refKey) ref(s, it.refKey, { x: W - M - 1.2, y: y + 0.05, w: 1.2, prefix: o.refPrefix ?? "根拠 " });
    });
  };

  // コンセプト（2026-11-25）：新しい事業・製品・サービスを初めて見せるページ。属性（誰に・何を・いくらで）を並べる場ではなく、価値を一目で見せる
  //   作る前に「このページで一番伝えたいこと」と「見た人に分かってもらえば OK なこと」を1文ずつ決め、そこから逆算する（決まり25）
  //   初見の聞き手には、仕組み（bizModel・processFlow・archDiagram）の前に、このページを置く
  //   o.name（名前。上に小さく）, o.statement（一言のコンセプト。価値を言う。全角30字まで）, o.key（statement の中で強調の色にする語。1つ）
  //   o.contrast（任意）：価値の差を、1本の物差しで比べる図。何と比べて、どれだけ違うかを一目で見せる
  //     { min?（物差しの左端の値。既定 0）, max?（右端の値。既定は rows と ticks の最大）, ticks: [{ value, label }]（物差しの目盛り）, resultHead?, note?,
  //       rows: [{ name, value（原稿にある値。帯の長さは部品が計算する）, end?（帯の端の短い札）, result?（右端の値）, strong?（自社の行。1つだけ）, big?（自社の帯の横に大きく書く一語。例「当日に冷凍」） }] }
  //     原稿に無い値を value にしない。値が無い行・範囲しか分からない行（「100g未満」「120g以上」）は帯を描かない。そうなるなら contrast を渡さず、一文だけにする
  //   見る所は自社の行（strong）か statement。base の focus に書く
  const concept = (s, o) => {
    CONCEPT_DONE = true;
    const st = String(o.statement || ""), key = o.key;
    if (!st) warn("コンセプト：一言のコンセプト（statement）を書く。属性ではなく、価値（何が、何と比べてどう違うか）を言う");
    if (st.length > 30) warn(`コンセプト：一言が長い（${st.length}字）。全角30字までに縮める。一言で言えないなら、何が価値なのかを決め直す`);
    if (o.name || (o.contrast && (o.contrast.note || (o.contrast.rows || []).some(r => r.end)))) warn("コンセプト：説明の小さな文（name・note・end）は置かない（2026-11-25）。名前は題か行の名前で、物差しは目盛りで分かるようにする。描かずに進める");
    if (o.points) warn("コンセプト：要点の列（points）は使わない（2026-11-25）。誰に・何を・いくらでのような属性を並べても、価値は伝わらない。価値の差は contrast で1本の物差しに並べて見せる");
    const runs = key && st.includes(key) ? st.split(key).flatMap((t, i, a) => (i < a.length - 1 ? [{ text: t }, { text: key, options: { color: C.emph } }] : [{ text: t }])).filter(r => r.text) : [{ text: st }];
    if (key && !st.includes(key)) warn(`コンセプト：強調する語（${key}）が一言の中に無い`);
    const ct = o.contrast;
    if (!ct) {
      // 一文だけを中央に（「、」があれば、そこで改行して2行に）
      const ny = 2.95;
      // 説明の小さな文（名前など）は置かない（2026-11-25）。名前は題か一文の中で言う
      const two = runs.flatMap(r => (String(r.text).includes("、") && st.length > 20 ? [{ text: r.text.split("、")[0] + "、", options: { ...(r.options || {}), breakLine: true } }, { text: r.text.split("、").slice(1).join("、"), options: r.options || {} }] : [r])).filter(r => r.text);
      T(s, two.map(r => ({ text: r.text, options: { fontSize: 40, bold: true, color: C.ink, ...(r.options || {}) } })), { x: M, y: ny + 0.55, w: BW, h: 1.6, align: "center", valign: "middle", lineSpacingMultiple: 1.15 });
      s.addShape(SH.LINE, { x: W / 2 - 1.2, y: ny + 2.4, w: 2.4, h: 0, line: { color: C.emph, width: 2 } });
      return {};
    }
    // 上に名前と一文、下に価値の対比
    // 一文はマジックワード（短い言い切り）。短ければ大きく出す
    const SFS = st.length <= 16 ? 44 : 30;
    T(s, runs.map(r => ({ text: r.text, options: { fontSize: SFS, bold: true, color: C.ink, ...(r.options || {}) } })), { x: M, y: 1.7, w: BW, h: SFS > 40 ? 1.0 : 0.75, align: "center", valign: "middle" });
    const rows = ct.rows || [], n = rows.length, x0 = M + 2.9, x1 = W - M - 2.0, y0 = SFS > 40 ? 3.9 : 3.55, pitch = Math.min(0.95, 2.7 / Math.max(n, 1)), bh = 0.42, X = t => x0 + (x1 - x0) * t;
    if (rows.filter(r => r.strong).length !== 1) warn("コンセプトの対比：自社の行（strong）を1つだけ決める");
    if (n !== 2) warn(`コンセプトの対比：比べるのは、相手1つと自社の2行だけ（${n} 行）。コンセプトは一点突破。相手を増やすと、何が突き抜けているかが薄まる`);
    if (rows.some(r => r.result !== undefined)) warn("コンセプトの対比：右端の結果の値（result）を足さない。2つ目の価値になる。伝える価値は1つだけ（物差しの差そのもの）");
    // 物差しの位置は、値から部品が計算する（割合 to・at を手で決めると、原稿に無い値が長さになる）
    const isNum = v => typeof v === "number" && Number.isFinite(v);
    const lo = ct.min ?? 0, hi = ct.max ?? Math.max(...rows.map(r => r.value).filter(isNum), ...(ct.ticks || []).map(t => t.value).filter(isNum));
    const posOf = v => (v - lo) / (hi - lo);
    const byValue = v => isNum(v) && isNum(hi) && hi > lo;
    if (rows.some(r => r.to !== undefined)) warn("コンセプトの対比：帯の長さを to（割合）で書かない。原稿の値を value に書き、長さは部品に計算させる");
    if ((ct.ticks || []).some(t => t.value === undefined)) warn("コンセプトの対比：目盛りの位置を at（割合）で書かない。目盛りの値を value に書く");
    const lenOf = r => {
      if ([r.value, r.big].some(isRange)) { rangeBar("コンセプトの対比", [r.value, r.big]); return null; }
      if (byValue(r.value)) {
        if (r.value < lo || r.value > hi) warn(`コンセプトの対比：「${r.name}」の値（${r.value}）が物差しの範囲（${lo}〜${hi}）の外`);
        return Math.min(1, Math.max(0, posOf(r.value)));
      }
      if (isNum(r.to)) return r.to; // 前の書き方（上で知らせる）
      warn(`コンセプトの対比：「${r.name}」の値（value）が無いので、帯を描かない。値の無い行があるなら contrast を渡さず、一文だけにする`);
      return null;
    };
    (ct.ticks || []).map(tk => ({ ...tk, at: byValue(tk.value) ? posOf(tk.value) : tk.at })).filter(tk => isNum(tk.at)).forEach(tk => {
      s.addShape(SH.LINE, { x: X(tk.at), y: y0 - 0.35, w: 0, h: pitch * n - 0.1, line: { color: C.line, width: 1 } });
      T(s, tk.label, { x: X(tk.at) - 0.6, y: y0 - 0.72, w: 1.2, h: 0.3, fontSize: FS.note, color: C.soft, align: "center", valign: "middle", wrap: false });
    });
    if (ct.resultHead) T(s, ct.resultHead, { x: x1 + 0.35, y: y0 - 0.72, w: 1.6, h: 0.3, fontSize: FS.note, color: C.soft, valign: "middle", wrap: false });
    rows.forEach((r, i) => {
      const y = y0 + i * pitch, on = !!r.strong;
      T(s, r.name, { x: M, y, w: x0 - M - 0.2, h: bh, fontSize: on ? FS.head : FS.body, bold: on, color: on ? C.ink : C.soft, align: "right", valign: "middle", wrap: false });
      const len = lenOf(r);
      if (len !== null) rect(s, x0, y, Math.max(0.12, X(len) - x0), bh, on ? C.emph : C.bar);
      if (on && r.big) focusValue(s, r.big, { x: len === null ? x0 : Math.max(X(len), x0 + 0.12) + 0.15, y: y - 0.1, w: 4, h: 0.62 });
      if (r.result !== undefined) T(s, String(r.result), { x: x1 + 0.35, y: y - 0.1, w: 1.6, h: 0.62, fontSize: on ? 36 : 22, bold: true, color: on ? C.emph : C.gray, valign: "middle", wrap: false });
    });
    return { rowY: i => y0 + i * pitch, X };
  };

  // 要約（2026-11-25）：結論（決めてほしいこと）を上に1つ、根拠を下に 2〜4 個並べ、線でつなぐ（ピラミッド）。上から下へ読めば「何を決めるか → なぜか」が分かる
  //   要約を points（見出しと補足の文の羅列）で描かない。根拠は「見出し1つ＋値1つ」で、補足の文を付けない
  //   o.decision: { label?（"決めていただくこと"）, text（何を決めるか。全角20字まで）, value（決める額・数。主役の値 36pt で強調の色）, ask?（最後の1語。例 "を、今日決めていただきたい"）}
  //   o.reasons: [{ head（根拠の言い切り。全角24字まで・2行まで）, value（根拠の数字。28pt 太字）, unit?（値の横の短い語。全角10字まで）, refKey? }]
  //   o.link（つなぎの線の上の語。既定 "根拠"）, o.y0（2.35）
  //   見る所は決めること（decision.value）。base の focus にその値を書く
  const summary = (s, o) => {
    const dc = o.decision || {}, rs = o.reasons || [];
    if (!dc.text || dc.value === undefined) warn("要約：決めてほしいこと（decision の text と value）を書く。要約の主役は、何を決めるか");
    if (rs.length < 2 || rs.length > 4) warn(`要約：根拠は 2〜4 個にする（${rs.length} 個）`);
    const y0 = o.y0 ?? 2.35, bh = 1.15, VAL_FS = 28, HEAD_FS = FS.lead;
    // 上：決めてほしいこと（強調の枠。この1か所だけ）
    rect(s, M, y0, BW, bh, TINT, { color: C.emph, width: 2 });
    T(s, dc.label || "決めていただくこと", { x: M + 0.3, y: y0 + 0.14, w: 6, h: 0.3, fontSize: FS.note, color: C.soft, valign: "middle", wrap: false });
    const vt = String(dc.value ?? ""), vw = textW(vt, FOCUS_FS) + 0.15;
    if (textW(dc.text || "", 24) > BW - vw - 0.9) warn(`要約：決めることの文が長い（${dc.text}）。全角20字までにする`);
    T(s, dc.text || "", { x: M + 0.3, y: y0 + 0.48, w: BW - vw - 0.9, h: 0.55, fontSize: 24, bold: true, color: C.ink, valign: "middle", wrap: false });
    focusValue(s, vt, { x: W - M - 0.3 - vw, y: y0 + 0.48, w: vw, h: 0.55, align: "right" });
    // つなぎ：上の箱の下の真ん中から、根拠の列の真ん中へ、直角に折れる線
    const n = Math.max(rs.length, 1), gap = 0.35, cw = (BW - gap * (n - 1)) / n, cx = i => M + i * (cw + gap) + cw / 2;
    const ya = y0 + bh, yb = ya + 0.42, yc = yb + 0.3, ln = { color: C.faint, width: 1.5 };
    s.addShape(SH.LINE, { x: W / 2, y: ya, w: 0, h: yb - ya, line: ln });
    if (n > 1) s.addShape(SH.LINE, { x: cx(0), y: yb, w: cx(n - 1) - cx(0), h: 0, line: ln });
    rs.forEach((r, i) => s.addShape(SH.LINE, { x: cx(i), y: yb, w: 0, h: yc - yb, line: ln }));
    T(s, o.link ?? "根拠", { x: W / 2 + 0.12, y: ya + 0.06, w: 2, h: 0.28, fontSize: FS.note, color: C.soft, valign: "middle", wrap: false });
    // 下：根拠の列（見出し → 値。補足の文は付けない）
    rs.forEach((r, i) => {
      const x = M + i * (cw + gap), y = yc + 0.2;
      if (r.body) warn(`要約：根拠に補足の文（body）を付けない（${r.head}）。言いたいことは head に1文で書き、数字は value に`);
      const nl = lineCount(r.head, HEAD_FS, cw);
      if (nl > 2) warn(`要約：根拠の見出しが ${nl} 行になる（${r.head}）。全角24字までに縮める`);
      checkWrap(r.head, HEAD_FS, cw);
      T(s, r.head, { x, y, w: cw, h: 0.8, fontSize: HEAD_FS, bold: true, color: C.ink, valign: "top", lineSpacingMultiple: 1.15 });
      const runs = [{ text: String(r.value ?? ""), options: { fontSize: VAL_FS, bold: true, color: C.ink, fontFace: FNUM } }];
      if (r.unit) runs.push({ text: "　" + r.unit, options: { fontSize: FS.note, color: C.soft } });
      if (r.unit && String(r.unit).length > 10) warn(`要約：値の横の語が長い（${r.unit}）。全角10字まで`);
      T(s, runs, { x, y: y + 1.0, w: cw, h: 0.6, valign: "middle", wrap: false });
      if (r.refKey) ref(s, r.refKey, { x, y: y + 1.68, w: cw, align: "left", prefix: "根拠 " });
    });
    const bottom = yc + 0.2 + 1.68 + 0.26;
    if (bottom > BOTTOM + 0.05) warn(`要約の下端 ${bottom.toFixed(2)} が本文の下端を超える（o.y0 を上げる）`);
    return { decision: { x: M, y: y0, w: BW, h: bh }, col: i => ({ x: M + i * (cw + gap), y: yc, w: cw, h: bottom - yc }) };
  };

  // ---- 詳細な表・多段の樹形図・プロジェクトのガント ----
  // 段階×観点の全体像（ジャーニーの表）：上に段階の矢羽、下に観点の行。1枚で高密度に見せる（FIT JOURNEY 並みの密度）
  //   o.stages: [段階名]（"\n" で改行位置を指定）, o.tag: [小, 大]（左上の見出し。例 ["After", "あるべき姿"]）
  //   o.rows: [{ label, h, cells: [cell, ...], align?: "center" }]
  //     cell: 文字列（"\n" 区切りの行。先頭が "!" の行は赤字＝今回やってしまったこと等の注意）, null（破線の空枠）,
  //           { span: n, text, frame?: true }（n 段階にまたがる1本の帯。会議の線の上に重ねる）, { people: [[役割, 社外?], ...] }（人のピクトを人数分）
  //   o.gates: [{ at: j（段階 j の前の境。段階の数なら最後の後）, label, kind: "gate"|"meeting", strong? }]（審査ゲート＝赤の破線、既存の会議＝灰の破線）
  //   o.face: "after" | "before" | null（面を敷く。反省の After と同じ面）, o.y0（会議名の行の上端、既定 2.1）, o.labelW（1.35）, o.fs（セルの文字 8.5）
  //   1枚で全体像を見せる高密度のページなので、本文の下端（6.6）を越えて 7.0 まで使ってよい
  const journeyMap = (s, o) => {
    warn("段階×観点の表（journeyMap）は使わない。マス目に文を積む形は、全部読まないと分からない。段階の道（roadmap：駅と、次へ進む条件の値）で描く");
    const n = o.stages.length, lw = o.labelW ?? 1.35, x0 = M + lw + 0.08, ag = o.gap ?? 0.06;
    const cw = (W - M - x0 - ag * (n - 1)) / n, cx = j => x0 + j * (cw + ag), fs = o.fs ?? 8.5;
    const y0 = o.y0 ?? 2.1, hy = y0 + 0.46, hh = o.headH ?? 0.5;
    const total = o.rows.reduce((a, r) => a + r.h + 0.06, 0) - 0.06, bottom = hy + hh + 0.1 + total;
    if (o.face) rect(s, M - 0.12, y0 - 0.1, W - 2 * M + 0.24, bottom - y0 + 0.22, BA_BG[o.face]);
    const cellFill = o.face ? C.white : TINT;
    if (o.tag) T(s, [{ text: o.tag[0], options: { color: C.gray, fontSize: FS.body, breakLine: true } }, { text: o.tag[1], options: { bold: true, fontSize: FS.head } }], { x: M, y: hy, w: lw, h: hh, valign: "middle", lineSpacingMultiple: 1.1 });
    o.stages.forEach((nm, j) => {
      s.addShape(j === 0 ? SH.PENTAGON : SH.CHEVRON, { x: cx(j), y: hy, w: j === n - 1 ? cw : cw + 0.12, h: hh, fill: { color: o.face ? C.white : TINT_STRONG }, line: { type: "none" } });
      // 矢羽の中の文字枠は狭いため、文字は別の箱に置く（pptx-notes）
      T(s, `${j + 1}. ${nm}`, { x: cx(j) + 0.06, y: hy, w: cw, h: hh, fontSize: FS.body, bold: true, align: "center", valign: "middle", lineSpacingMultiple: 1.0 });
    });
    let y = hy + hh + 0.1;
    const rowY = [];
    o.rows.forEach(r => {
      rowY.push(y);
      T(s, r.label, { x: M, y, w: lw, h: r.h, fontSize: FS.body, bold: true, valign: "middle", lineSpacingMultiple: 1.1 });
      let j = 0;
      r.cells.forEach(c => {
        const span = (c && typeof c === "object" && c.span) || 1, x = cx(j), w = cx(j + span - 1) + cw - x;
        j += span;
        if (c === null) { rect(s, x, y, w, r.h, null, { color: C.line, width: 0.75, dashType: "dash" }); return; }
        if (typeof c === "object" && c.people) {
          rect(s, x, y, w, r.h, cellFill);
          const k = c.people.length, sw = Math.min(0.42, (w - 0.06) / k), sx = x + (w - sw * k) / 2;
          c.people.forEach(([role, ext], q) => {
            person(s, sx + q * sw + sw / 2, y + 0.05, 0.1, ext ? C.pale : C.ink);
            T(s, role, { x: sx + q * sw - 0.03, y: y + 0.29, w: sw + 0.06, h: 0.18, fontSize: k >= 6 ? 5.5 : 6.5, color: C.soft, align: "center", valign: "middle", wrap: false });
          });
          return;
        }
        const text = typeof c === "string" ? c : c.text;
        rect(s, x, y, w, r.h, cellFill, typeof c === "object" && c.frame ? { color: C.emph, width: 0.75 } : null);
        const ls = String(text).split("\n");
        T(s, ls.map((l, q) => ({ text: l.replace(/^!/, ""), options: { color: l.startsWith("!") ? C.accent : (typeof c === "object" && c.frame ? C.ink : C.soft), breakLine: q < ls.length - 1 } })),
          { x: x + 0.06, y, w: w - 0.12, h: r.h, fontSize: fs, valign: "middle", align: (typeof c === "object" && c.span) || r.align === "center" ? "center" : "left", lineSpacingMultiple: 1.15 });
        if (lineCount(ls.join("\n"), fs, w - 0.12) * lineH(fs, 1.15) > r.h + 0.02) warn(`ジャーニーの「${r.label}」のセルの文字があふれる（${ls[0].slice(0, 10)}…）`);
      });
      y += r.h + 0.06;
    });
    (o.gates || []).forEach(g => {
      const gx = g.at >= n ? cx(n - 1) + cw + ag / 2 : cx(g.at) - ag / 2, meet = g.kind === "meeting";
      s.addShape(SH.LINE, { x: gx, y: y0 + 0.36, w: 0, h: bottom - y0 - 0.36, line: { color: meet ? C.gray : C.accent, width: g.strong ? 3 : meet ? 1 : 1.75, dashType: "dash" } });
      T(s, g.label, { x: gx - 0.95, y: y0 - 0.06, w: 1.9, h: 0.42, fontSize: FS.body, bold: true, color: meet ? C.gray : C.accent, align: "center", valign: "bottom", lineSpacingMultiple: 1.05 });
    });
    if (bottom > BOTTOM + (FOOT_TEXT ? 0.3 : 0.4)) warn(`ジャーニーの表の下端 ${bottom.toFixed(2)} がフッターに近い（行の高さを詰める）`);
    return { cx, cw, rowY, bottom };
  };

  // 多段の樹形図（ロジックツリー）：左の1つのノードから右へ枝分かれし、1つのノードから2〜3本ずつ広がる（4.3 を多段にしたもの）
  //   o.root: { text, sub?, strong?, problem?, children: [ノード, ...] }（深さ 3〜4 まで。葉の数は 12 前後まで）
  //   o.heads: [各列の見出し]（例 ["問題", "直接の原因", "根本の原因"]）, o.colW: [各列の幅]（省くと葉の列を広く自動配分）
  //   o.y0（2.55）, o.y1（6.6）, o.colGap（0.36）
  //   強調: strong（濃い塗り＋ink 2pt 枠。結論・対策の起点）、problem（白地＋赤枠＋赤字。手を打つべき根本の原因）。1ページで同じ意味の強調だけを使う
  const logicTree = (s, o) => {
    const depthOf = nd => 1 + Math.max(0, ...(nd.children || []).map(depthOf));
    const D = depthOf(o.root), gap = o.colGap ?? 0.36, avail = W - 2 * M - gap * (D - 1);
    // 列の幅：根（問題）は1行に収まる幅、中間は短い見出し＋数値、葉（根本の原因）は見出し＋説明が1行ずつ入る幅
    const wts = o.colW ? null : Array.from({ length: D }, (_, i) => i === D - 1 ? 1.75 : i === 0 ? 1.15 : 1.0);
    const colW = o.colW || wts.map(v => avail * v / wts.reduce((a, b) => a + b, 0));
    const colX = colW.map((_, i) => M + colW.slice(0, i).reduce((a, b) => a + b, 0) + gap * i);
    const leaves = []; const walk = (nd, d) => { nd._d = d; if (!nd.children || !nd.children.length) leaves.push(nd); else nd.children.forEach(c => walk(c, d + 1)); }; walk(o.root, 0);
    const y0 = o.y0 ?? 2.45, y1 = o.y1 ?? BOTTOM + 0.15, lg = o.leafGap ?? 0.06;
    const lh = Math.min(0.72, (y1 - y0 - lg * (leaves.length - 1)) / leaves.length);
    leaves.forEach((lf, i) => { lf._y = y0 + i * (lh + lg); lf._h = lh; });
    const place = nd => { if (nd.children && nd.children.length) { nd.children.forEach(place); const a = nd.children[0], b = nd.children[nd.children.length - 1]; const top = a._y, bot = b._y + b._h; const h = Math.min(bot - top, Math.max(lh, o.nodeH ?? 0.9)); nd._h = h; nd._y = (top + bot) / 2 - h / 2; } }; place(o.root);
    if (o.heads) o.heads.forEach((t, i) => i < D && T(s, t, { x: colX[i], y: y0 - 0.38, w: colW[i], h: 0.3, fontSize: FS.block, bold: true, color: C.gray, valign: "bottom", wrap: false }));
    const fsOf = d => d === 0 ? FS.head : d === D - 1 ? FS.body : FS.block;
    const draw = nd => {
      const d = nd._d, x = colX[d], w = colW[d], leaf = !nd.children || !nd.children.length;
      const fill = nd.problem ? C.white : nd.strong || d === 0 ? TINT_STRONG : TINT;
      const line = nd.problem ? { color: C.accent, width: 2 } : nd.strong ? { color: C.emph, width: 2 } : null;
      rect(s, x, nd._y, w, nd._h, fill, line);
      const color = nd.problem ? C.accent : C.ink, f = fsOf(d);
      const runs = [{ text: nd.text, options: { fontSize: f, bold: !leaf || nd.problem || nd.strong, color: leaf && !nd.problem && !nd.strong ? C.ink : color, breakLine: !!nd.sub } }];
      if (nd.sub && (warnSub("図の要素"), true)) runs.push({ text: nd.sub, options: { fontSize: Math.min(f, FS.body) - (leaf ? 1 : 0), color: nd.problem ? C.accent : C.soft } });
      T(s, runs, { x: x + 0.12, y: nd._y, w: w - 0.24, h: nd._h, valign: "middle", lineSpacingMultiple: 1.15 });
      const sf = Math.min(f, FS.body) - (leaf ? 1 : 0);
      const need = lineCount(nd.text, f, w - 0.24) * lineH(f, 1.15) + (nd.sub ? lineCount(nd.sub, sf, w - 0.24) * lineH(sf, 1.15) : 0) + 0.02;
      if (need > nd._h + 0.02) warn(`樹形図のノードの文字があふれる（${nd.text.slice(0, 12)}…）。葉を減らすか列を広げる`);
      if (!leaf) {
        const sx = x + w + gap / 2, my = nd._y + nd._h / 2, ln = { color: C.gray, width: 1 };
        segLine(s, x + w, my, sx, my, ln);
        const cys = nd.children.map(c => c._y + c._h / 2);
        if (cys.length > 1) segLine(s, sx, Math.min(...cys), sx, Math.max(...cys), ln);
        cys.forEach(cy => segLine(s, sx, cy, colX[d + 1], cy, ln));
        nd.children.forEach(draw);
      }
    };
    draw(o.root);
    return { colX, colW, leaves };
  };

  // KPIツリー（指標の分解）：左の指標を、掛け算・足し算で右へ分解する樹形図。各ノードに名前・値・増減を置き、分かれ目に演算子（× ＋）の丸
  //   原因の分解（logicTree）は文の樹形図。KPIツリーは、数値がつながる（親が子の積や和になる）ときに使う
  //   o.root: { name, value, chg?（増減。"−11.5%" など）, op?（子の結び方 "×" "＋" "−" "÷"）, children?,
  //             problem?（赤枠。増減を主役の値 36pt の赤で）, strong?（強調の色の枠。増減を主役の値 36pt の強調の色で） }
  //   見る所（problem か strong）は1つだけ。葉は6個まで、段は4段まで
  //   o.y0（2.4）, o.y1（本文の下端）, o.colGap（0.85。演算子の丸が入る幅）
  //   戻り値: { nodes: [{ name, x, y, w, h }] } … 前面から指す位置
  const kpiTree = (s, o) => {
    chartPage();
    const depthOf = nd => 1 + Math.max(0, ...(nd.children || []).map(depthOf));
    const D = depthOf(o.root), gap = o.colGap ?? 0.85, bw = (BW - gap * (D - 1)) / D;
    const leaves = [], all = [];
    const walk = (nd, d) => { nd._d = d; all.push(nd); if (nd.children && nd.children.length) nd.children.forEach(c => walk(c, d + 1)); else leaves.push(nd); };
    walk(o.root, 0);
    const foc = all.filter(n => n.problem || n.strong);
    if (!foc.length) warn("KPIツリーに見る所が無い。リードで言う指標の1つに problem（悪い所）か strong（効く所）を付ける");
    if (foc.length > 1) warn(`KPIツリーの見る所が ${foc.length} つある。1つにする`);
    if (leaves.length > 6) warn(`KPIツリーの葉が ${leaves.length} 個ある。6個までにする（リードに効く枝だけを分解する）`);
    if (D > 4) warn(`KPIツリーが ${D} 段ある。4段までにする`);
    const y0 = o.y0 ?? 2.4, y1 = o.y1 ?? BOTTOM + 0.1, n = leaves.length;
    const bh = Math.min(1.05, (y1 - y0 - 0.14 * (n - 1)) / n);
    if (bh < 0.9) warn(`KPIツリーの箱が低い（${bh.toFixed(2)}）。値が入らないので、葉を減らす`);
    const lg = n > 1 ? Math.min(0.45, (y1 - y0 - n * bh) / (n - 1)) : 0, used = n * bh + (n - 1) * lg, top = y0 + (y1 - y0 - used) / 2;
    leaves.forEach((lf, i) => { lf._y = top + i * (bh + lg); });
    const place = nd => { if (nd.children && nd.children.length) { nd.children.forEach(place); const a = nd.children[0], b = nd.children[nd.children.length - 1]; nd._y = (a._y + b._y) / 2; } };
    place(o.root);
    const colX = d => M + d * (bw + gap), ln = { color: C.gray, width: 1.25 };
    const nodes = [];
    const draw = nd => {
      const x = colX(nd._d), y = nd._y, big = !!(nd.problem || nd.strong);
      rect(s, x, y, bw, bh, nd.problem ? C.white : big ? TINT_STRONG : TINT, nd.problem ? { color: C.accent, width: 2 } : nd.strong ? { color: C.emph, width: 2 } : null);
      T(s, nd.name, { x: x + 0.16, y: y + 0.06, w: bw - 0.3, h: 0.34, fontSize: FS.body, bold: true, color: nd.problem ? C.accent : C.soft, valign: "middle", wrap: false });
      const vfs = 24, vt = String(nd.value ?? ""), ct = nd.chg == null ? "" : String(nd.chg);
      T(s, vt, { x: x + 0.16, y: y + bh - 0.6, w: bw - 0.3, h: 0.5, fontSize: vfs, bold: true, color: C.ink, valign: "middle", wrap: false });
      if (ct) {
        if (big) focusValue(s, ct, { x: x + 0.1, y: y + bh - 0.62, w: bw - 0.24, align: "right" }, { problem: nd.problem });
        else T(s, ct, { x: x + 0.1, y: y + bh - 0.6, w: bw - 0.26, h: 0.5, fontSize: FS.lead, bold: true, color: C.soft, align: "right", valign: "middle", wrap: false });
      }
      const need = textW(vt, vfs) + (ct ? textW(ct, big ? FOCUS_FS : FS.lead) + 0.25 : 0) + 0.3;
      if (need > bw) warn(`KPIツリーの「${nd.name}」の値と増減が箱の幅に入らない（必要 ${need.toFixed(2)} > ${bw.toFixed(2)}）。段を減らすか、値を短くする（単位を「万」「億」に）`);
      if (textW(nd.name, FS.body) > bw - 0.3) warn(`KPIツリーの名前が1行に入らない（${nd.name}）`);
      nodes.push({ name: nd.name, x, y, w: bw, h: bh });
      if (!nd.children || !nd.children.length) return;
      const my = y + bh / 2, xm = x + bw + gap / 2, cys = nd.children.map(c => c._y + bh / 2);
      segLine(s, x + bw, my, xm, my, ln);
      if (cys.length > 1) segLine(s, xm, Math.min(...cys), xm, Math.max(...cys), ln);
      cys.forEach(cy => segLine(s, xm, cy, xm + gap / 2, cy, ln));
      if (nd.op) {
        const d0 = 0.44, oy = (Math.min(...cys) + Math.max(...cys)) / 2;
        s.addShape(SH.OVAL, { x: xm - d0 / 2, y: oy - d0 / 2, w: d0, h: d0, fill: { color: C.white }, line: ln });
        T(s, nd.op, { x: xm - d0 / 2, y: oy - d0 / 2, w: d0, h: d0, fontSize: FS.block, bold: true, color: C.ink, align: "center", valign: "middle", wrap: false });
      } else if (nd.children.length > 1) warn(`KPIツリーの「${nd.name}」に演算子（op）が無い。子との関係（× か ＋）を書く`);
      nd.children.forEach(draw);
    };
    draw(o.root);
    return { nodes };
  };

  // RACI（役割の分担）：作業（行）×担当（列）のセルに、R・A・C・I の印だけを置く。セルに文字は入れない
  //   R＝実行する（灰の濃い塗り）、A＝決める（強調の色。1つの作業に1人）、C＝相談を受ける（白抜き・濃い線）、I＝報告を受ける（白抜き・薄い線）
  //   o.roles: [担当], o.rows: [{ name, marks: ["A", "R", "", …]（担当の順）, problem?（赤枠。A がいない・2人いる行）, note?（赤の短い札。省くと部品が「A がいない」などを書く） }]
  //   o.legend（既定 true。印の意味を表の上に1行）, o.y0（表の本文の上端 3.05）, o.labelW（作業名の幅。省くと最長に合わせる）
  //   A が無い・2つ以上ある行に problem を付けないと警告する（見せたい問題でないなら、表を直す）
  //   戻り値: { rowY, rh, colX(j), x0 } … 前面から指す位置
  const raciTable = (s, o) => {
    const roles = o.roles, rows = o.rows, nr = rows.length;
    if (roles.length > 6) warn(`RACI の担当が ${roles.length} 列ある。6列までにする`);
    if (nr > 8) warn(`RACI の作業が ${nr} 行ある。8行までにする（リードに効く作業だけ残す）`);
    const lfs = FS.block, mfs = FS.block;
    const lw = o.labelW ?? Math.max(2.4, ...rows.map(r => textW(r.name, lfs) + (r.problem ? textW(r.note || "A がいない", FS.body) + 0.3 : 0))) + 0.3;
    const x0 = M + lw, cw = (W - M - x0) / roles.length, colX = j => x0 + j * cw + cw / 2;
    const y0 = o.y0 ?? 3.05, rh = Math.min(0.72, (BOTTOM + 0.1 - y0) / nr), dd = Math.min(0.46, rh - 0.16);
    // R は灰の濃い塗り、A は強調の色。強調の色が文字の色と同じテーマ（navy）でも、A と R を見分けられるようにする
    const ST = { R: [C.soft, C.white, C.soft, 1.25], A: [C.emph, C.white, C.emph, 1.25], C: [C.white, C.ink, C.ink, 1.5], I: [C.white, C.soft, C.gray, 1.25] };
    const mark = (cx, cy, k, d0) => {
      if (!k) return;
      const st = ST[k];
      if (!st) { warn(`RACI の印は R・A・C・I のどれか（${k}）`); return; }
      s.addShape(SH.OVAL, { x: cx - d0 / 2, y: cy - d0 / 2, w: d0, h: d0, fill: { color: st[0] }, line: { color: st[2], width: st[3] } });
      T(s, k, { x: cx - d0 / 2, y: cy - d0 / 2, w: d0, h: d0, fontFace: FNUM, fontSize: d0 >= 0.4 ? mfs : FS.body, bold: true, color: st[1], align: "center", valign: "middle", wrap: false });
    };
    if (o.legend !== false) {
      let lx = M;
      [["R", "実行する"], ["A", "決める（1人）"], ["C", "相談を受ける"], ["I", "報告を受ける"]].forEach(([k, t]) => {
        mark(lx + 0.17, y0 - 0.82, k, 0.34);
        T(s, t, { x: lx + 0.42, y: y0 - 0.99, w: textW(t, FS.body) + 0.1, h: 0.34, fontSize: FS.body, color: C.soft, valign: "middle", wrap: false });
        lx += 0.42 + textW(t, FS.body) + 0.45;
      });
    }
    roles.forEach((r, j) => {
      T(s, r, { x: x0 + j * cw, y: y0 - 0.5, w: cw, h: 0.42, fontSize: FS.body, bold: true, color: C.ink, align: "center", valign: "bottom", wrap: false });
      if (textW(r, FS.body) > cw - 0.08) warn(`RACI の担当の名前が列の幅に入らない（${r}）。担当を減らすか、名前を縮める`);
    });
    hline(s, M, y0 - 0.03, BW, { color: C.gray, width: 1 });
    const rowY = [];
    rows.forEach((r, i) => {
      const y = y0 + i * rh, nA = r.marks.filter(k => k === "A").length;
      rowY.push(y);
      if (r.marks.length !== roles.length) warn(`RACI の「${r.name}」の印の数（${r.marks.length}）が担当の数（${roles.length}）と違う`);
      if (nA !== 1 && !r.problem) warn(`RACI の「${r.name}」に A が ${nA} つある。A は1つの作業に1人。問題として見せるなら problem を付ける`);
      if (i) hline(s, M, y, BW, { color: C.line, width: 0.75 });
      T(s, r.name, { x: M + 0.12, y, w: lw - 0.2, h: rh, fontSize: lfs, bold: true, color: r.problem ? C.accent : C.ink, valign: "middle", wrap: false });
      r.marks.forEach((k, j) => mark(colX(j), y + rh / 2, k, dd));
      if (r.problem) {
        rect(s, M - 0.05, y + 0.04, BW + 0.1, rh - 0.08, null, { color: C.accent, width: 2 });
        const nt = r.note || (nA === 0 ? "A がいない" : `A が${nA}人`);
        T(s, nt, { x: M + 0.12 + textW(r.name, lfs) + 0.22, y, w: textW(nt, FS.body) + 0.1, h: rh, fontSize: FS.body, bold: true, color: C.accent, valign: "middle", wrap: false });
      }
    });
    if (rows.filter(r => r.problem).length > 1) warn("RACI の problem の行が2つ以上ある。見る所は1つにする");
    return { rowY, rh, colX, x0 };
  };

  // プロジェクトのガント：計画だけを、グループ（大項目）と作業（小項目）の2段で並べる。10〜20行、日単位の帯
  //   計画を示す図なので、見た目は1種類にそろえる。作業の帯はすべて同じ色・同じ太さ。遅れ・予定と実績・今日の線・ひし形・会議の線は描かない
  //   （遅れや進み具合を示すときは、別のページで同じ図を echo で背面に敷き、前面の messageBox で言う）
  //   o.from, o.to（"2027.01" の形。列は月）, o.groups: [{ name, owner?, tasks: [{ name, owner?, bar: [開始, 終了] }] }]
  //     日付は "2027.01.15" か "2027.01"（月の初め／終わり）。bar の終了は「その日まで」
  //   o.labelW（2.3）, o.ownerW（0.85。担当が無ければ 0）, o.y0（2.25）, o.rowH / o.groupH（省くと本文の下端までで自動）, o.color（帯の色、既定 ink）
  const projectGantt0 = (s, o) => {
    const tfs = o.fs ?? 16, gfs = tfs + 2; // 10〜20行を並べる密な図（作業名 16pt、グループ名 18pt）
    const months = []; { let [y, m] = o.from.split(".").map(Number); const [ty, tm] = o.to.split(".").map(Number); while (y < ty || (y === ty && m <= tm)) { months.push([y, m]); m++; if (m > 12) { m = 1; y++; } } }
    const hasOwner = o.groups.some(g => g.owner || g.tasks.some(t => t.owner));
    const lw = o.labelW ?? Math.max(1.4, Math.min(3.4, Math.max(...o.groups.map(g => textW(g.name, gfs) + 0.24), ...o.groups.flatMap(g => g.tasks.map(t => textW(t.name, tfs) + 0.55))))), ow = hasOwner ? (o.ownerW ?? 0.85) : 0, x0 = M + lw + ow + 0.1, x1 = W - M, mw = (x1 - x0) / months.length;
    const dim = (y, m) => new Date(y, m, 0).getDate();
    const xOf = (p, end) => {
      const [y, m, d] = p.split(".").map(Number), i = months.findIndex(([yy, mm]) => yy === y && mm === m);
      if (i < 0) throw new Error(`projectGantt: ${p} は ${o.from}〜${o.to} の外`);
      const days = dim(y, m), frac = d ? (end ? d : d - 1) / days : (end ? 1 : 0);
      return x0 + (i + frac) * mw;
    };
    const nG = o.groups.length, nT = o.groups.reduce((a, g) => a + g.tasks.length, 0);
    const y0 = o.y0 ?? 2.25, headH = 0.6, top = y0 + headH + 0.04, avail = BOTTOM + 0.15 - top;
    const k = avail / (nT + nG * 1.15 + (nG - 1) * 0.35);
    const rh = o.rowH ?? Math.min(0.34, k), gh = o.groupH ?? Math.min(0.38, k * 1.15), gsep = Math.min(0.12, k * 0.35);
    const color = o.color || C.ink;
    // 行の高さが文字（作業名 FS.body・グループ名 FS.block）の1行より低いと、上下の行の文字が重なる
    if (rh < lineH(tfs, 1.0) - 0.005) warn(`ガントの作業の行の高さ ${rh.toFixed(2)} が、文字（${tfs}pt、${lineH(tfs, 1.0).toFixed(2)}）より低い。作業を減らすか、グループごとに2枚に分ける`);
    if (gh < lineH(gfs, 1.0) - 0.005) warn(`ガントのグループの行の高さ ${gh.toFixed(2)} が、文字（${gfs}pt、${lineH(gfs, 1.0).toFixed(2)}）より低い。作業を減らすか、グループごとに2枚に分ける`);
    // 見出し：上段に年、下段に月
    months.forEach(([y, m], i) => {
      const x = x0 + i * mw;
      if (i === 0 || m === 1) T(s, `${y}年`, { x: x + 0.04, y: y0, w: 1.2, h: 0.28, fontSize: tfs, bold: true, color: C.soft, wrap: false });
      T(s, `${m}月`, { x, y: y0 + 0.3, w: mw, h: 0.28, fontSize: tfs, color: C.gray, align: "center", wrap: false });
    });
    if (hasOwner) T(s, "担当", { x: M + lw, y: y0 + 0.3, w: ow, h: 0.28, fontSize: tfs, color: C.gray, wrap: false });
    let y = top; const rows = [];
    o.groups.forEach((g, gi) => {
      if (gi > 0) y += gsep;
      rows.push({ g, y, h: gh }); y += gh;
      g.tasks.forEach(t => { rows.push({ t, g, y, h: rh }); y += rh; });
    });
    const bottom = y;
    // グループの行は全幅に薄い塗りの帯（罫線の代わりにまとまりを示す）。月の区切りは細い縦線（背面）
    rows.forEach(r => { if (!r.t) rect(s, M, r.y + 0.02, W - 2 * M, r.h - 0.04, TINT); });
    months.forEach((_, i) => i > 0 && s.addShape(SH.LINE, { x: x0 + i * mw, y: y0 + (months[i][1] === 1 ? 0 : 0.3), w: 0, h: bottom - y0 - (months[i][1] === 1 ? 0 : 0.3), line: { color: months[i][1] === 1 ? C.gray : C.line, width: 0.5 } }));
    rows.forEach(r => {
      if (!r.t) {
        T(s, r.g.name, { x: M + 0.08, y: r.y, w: lw - 0.08, h: r.h, fontSize: gfs, bold: true, valign: "middle", wrap: false });
        if (r.g.owner) T(s, r.g.owner, { x: M + lw, y: r.y, w: ow, h: r.h, fontSize: tfs, bold: true, color: C.soft, valign: "middle", wrap: false });
        return;
      }
      const t = r.t, bH = r.h * 0.5, by = r.y + (r.h - bH) / 2;
      // 作業名は時間の軸の側（欄の右）にそろえる。グループ名は左のまま（階層が分かる）
      T(s, t.name, { x: M + 0.3, y: r.y, w: lw - 0.3, h: r.h, fontSize: tfs, color: C.soft, align: hasOwner ? "left" : "right", valign: "middle", wrap: false });
      if (textW(t.name, tfs) > lw - 0.3) warn(`ガントの作業名が列に収まらない（${t.name}）`);
      if (t.owner) T(s, t.owner, { x: M + lw, y: r.y, w: ow, h: r.h, fontSize: tfs, color: C.soft, valign: "middle", wrap: false });
      const a = xOf(t.bar[0]), b = xOf(t.bar[1], true);
      rect(s, a, by, Math.max(b - a, 0.04), bH, t.focus ? C.emph : rows.some(q => q.t && q.t.focus) ? C.bar : color);
    });
    if (bottom > BOTTOM + 0.2) warn(`ガントの下端 ${bottom.toFixed(2)} が本文の下端を超える（作業を減らすか、グループごとに2枚に分ける）`);
    return { xOf, rows, bottom };
  };
  const projectGantt = (s, o) => { denseFig = true; try { return projectGantt0(s, o); } finally { denseFig = false; } };

  // ---- 比較の表（数値・評価・選択の表の既定。罫線を使わず、ブロックと余白で組む） ----
  // 行ごとに1本の薄い塗りの帯（列は帯の中の余白で分ける）、行の間は白い溝。罫線・縦線・外枠は使わない
  // 調査（戦略コンサルの表）から取り入れたこと: 選択肢名は bold ink、単位は見出しに一度、数値は右揃え、判断の行で理由を書く、強調は1列か1行だけ
  // ハーベイボール：v は 0〜4（0%・25%・50%・75%・100%）。判断の軸だけに使い、測れる軸は数値で書く
  const harvey = (s, cx, cy, v, d = 0.26) => {
    s.addShape(SH.OVAL, { x: cx - d / 2, y: cy - d / 2, w: d, h: d, fill: { color: v >= 4 ? C.ink : C.white }, line: { color: C.ink, width: 0.75 } });
    if (v > 0 && v < 4) s.addShape(SH.PIE, { x: cx - d / 2, y: cy - d / 2, w: d, h: d, angleRange: [270, (270 + 90 * v) % 360], fill: { color: C.ink }, line: { type: "none" } }); // 丸の塗りは1色（強調の色は見る所の枠だけ）
  };
  // 比較の表
  //   o.cols: [{ head, sub?（使わない。警告が出る。単位は見出しの名前に「（万円）」のように書く）, w?, align?: "left"|"right", base?（比べる基準＝現状の列。見出しと値を soft）, strong?（推奨・選ぶ列。列全体を ink の枠で囲む）, tag?（見出しの上の札。例 "推奨"） }]
  //   o.rows: [{ label, sub?, group?（この行から始まる群の名前。例 "効果"）, strong?（選ぶ1行。帯を濃くし ink の枠）, cells: [cell, ...] }]
  //   o.judge: { label, cells }（最後の「判断」の行。各案を選ぶ・選ばない理由を1行で。結論は表の中に置く）
  //   cell: 文字列・数値 | null（"—"） | { text, bold?, problem?（赤字） } | { ball: 0〜4, text }（ハーベイボール＋根拠） | { bar: 値, max, text, main? }（割合などの横棒＋値）
  //   o.x（M）, o.w（BW）, o.y（最初の行の上端。既定 2.95）, o.labelW（2.4）, o.colGap（0.24）, o.fs（10.5。11行以上は 9.5 にしてよい）, o.legend（true でハーベイボールの凡例を右上に1行）
  //   強調は1ページ1か所: 列の strong か、行の strong のどちらか一方。太字にするのは選択肢名と決め手の値（cell.bold）だけ
  const compareTable = (s, o) => {
    // 表の組み方（2026-11-24）：(1) 行を詰める。行の高さは文字の約2倍で、本文の領域に合わせて引き伸ばさない（間が広いと、行にも列にも見えず、語が散って見える）
    //   (2) 値を主役にする。値は濃い 18pt、行の名前と列の見出しは控えめな 16pt（枠組みの文字が値より目立つと、値が後回しになる）
    //   (3) 1つの表には、同じ種類の値だけを入れる（数値・語・評価の丸を混ぜると、列を縦に読めない）。表の幅は中身に合わせる
    const x0 = o.x ?? M, wAll = o.w ?? BW, cg = o.colGap ?? 0.4, fs = o.fs ?? 20, hfs = fs - 1, lfs = fs - 1, pad = 0.1; // 値 20pt、行の名前と列の見出し 19pt（案B）
    const textOf = c => c == null ? "—" : typeof c === "object" ? String(c.text ?? (c.stars != null ? `★${c.stars}` : c.ball != null ? `●${c.ball}` : "")) : String(c);
    const widest = (t, f) => Math.max(0, ...String(t).split("\n").map(l => textW(l, f)));
    if (o.rows.length > 5 || o.rows.length * o.cols.length > 20) warn(`表が大きい（${o.rows.length}行×${o.cols.length}列）。主張に効く行だけ残す（5行×4列まで）。違いを言いたいだけなら、表をやめて図にする`);
    // 数値の表で、行ごとに桁が違う（最大の値が10倍以上離れている）と、縦に並んだ数字を見比べてしまい、読みにくい。指標ごとの棒（metricBars）にする
    { const mx = o.rows.filter(r => r.cells.every(c => /^[-−+±]?[\d,.]+%?$/.test(textOf(c)))).map(r => Math.max(...r.cells.map(c => Math.abs(parseFloat(textOf(c).replace(/[,+±%]/g, "").replace("−", "-"))) || 0))).filter(v => v > 0);
      if (mx.length >= 2 && Math.max(...mx) / Math.min(...mx) >= 10) warn("表の行ごとに、数値の桁が違う。数値を表に並べず、指標ごとに尺度をそろえた棒で比べる（metricBars）"); }
    // 行見出しの幅：いちばん長い行見出しに合わせる（o.labelW で固定もできる）
    const lw = o.labelW ?? Math.min(3.4, Math.max(1.2, ...o.rows.map(r => widest(r.label, lfs)), o.judge ? widest(o.judge.label || "判断", lfs) : 0) + pad + 0.3);
    const fixed = o.cols.reduce((a, c) => a + (c.w || 0), 0), flexN = o.cols.filter(c => !c.w).length;
    const fwMax = flexN ? (wAll - lw - fixed - cg * o.cols.length - pad) / flexN : 0;
    // 列の幅：中身のいちばん長い行に合わせる（上限は、幅いっぱいに等分したときの幅と 3.4 の小さいほう。長い文は折り返す）。o.fit: false で、幅いっぱいに等分する
    const natW = j => Math.max(widest(o.cols[j].head, hfs), widest(o.cols[j].sub || "", FS.note), ...o.rows.map(r => { const c = r.cells[j]; return c && typeof c === "object" && c.stars != null ? (c.max ?? 5) * 0.25 : widest(textOf(c), c && typeof c === "object" && c.focus ? fs + 5 : fs) + (c && typeof c === "object" && c.ball != null ? 0.32 : 0); }), o.judge ? widest(textOf(o.judge.cells[j]), fs) : 0) + 0.2 + (o.cols[j].icon ? 0.4 : 0);
    const fw = j => o.fit === false ? fwMax : Math.min(fwMax, 3.4, Math.max(1.3, natW(j)));
    const xs = []; let cx = x0 + lw + cg;
    o.cols.forEach((c, j) => { const w = c.w || fw(j); xs.push([cx, w]); cx += w + cg; });
    const right = Math.min(x0 + wAll, cx - cg + pad), y0 = o.y ?? 2.95, hasTag = o.cols.some(c => c.tag), rg = o.rowGap ?? 0.05, gsp = 0.14;
    const numLike = c => typeof c === "number" || (c != null && (typeof c !== "object" || (c.ball == null && c.bar == null)) && /^[-−+±]?[\d,.]+%?$/.test(textOf(c)));
    const colAlign = o.cols.map((c, j) => c.align || (o.rows.length && o.rows.every(r => numLike(r.cells[j])) ? "right" : "left"));
    // 表は、一目で比べるためのもの。セルに文字を入れない（全部読まないと分からない表は、表の意味をなさない）
    //   入れてよいのは、数値・ゲージ（{ bar, max }）・評価の丸（{ ball: 0〜4 }）・星（{ stars: n, max? }）と、全角4字までの短い札（「なし」「Mac」）だけ
    { const isMark = c => c && typeof c === "object" && (c.ball != null || c.bar != null || c.stars != null);
      const bad = [];
      o.rows.forEach(r => r.cells.forEach(c => {
        if (c == null || numLike(c)) return;
        if (isMark(c)) { if (c.ball != null && c.text) bad.push(c.text); return; }
        if (em(textOf(c)) > 4) bad.push(textOf(c));
      }));
      if (bad.length) warn(`表のセルに文字が入っている（${bad.length}か所。「${String(bad[0]).slice(0, 12)}」ほか）。表のセルは、数値・ゲージ・評価の丸・星だけにする。文字で説明が要る内容は、表にしない（箇条書き・図にする）`);
      const kindOf = r => r.cells.every(c => c == null || numLike(c)) ? "数値" : r.cells.some(isMark) ? "評価の印" : "語";
      const kinds = [...new Set(o.rows.map(kindOf))];
      if (kinds.length > 1) warn(`表の行ごとに、中身の種類が違う（${kinds.join("・")}）。1つの表には同じ種類の値だけを入れる`);
      if (o.judge) warn("表に「判断」の行がある。文字の行なので、結論はリードか、表の横（o.note）に書く"); }
    // 決め手のセル（focus）の文字の大きさ：数値は 20pt、短い語（全角12字まで）は本文＋3pt、長い文は本文のまま
    const focusFsOf = c => !(c && typeof c === "object" && c.focus) ? fs : numLike(c) ? 22 : em(textOf(c)) <= 12 ? fs + 3 : fs;
    // 行の高さ：中身の行数から決める（1行 0.44、ハーベイボール＋根拠 0.48、2行 0.66 前後）。空きが出ても引き伸ばすのは 0.62 まで
    const need = r => Math.max(0.4, ...r.cells.map((c, j) => {
      const w = xs[j][1] - (c && typeof c === "object" && c.ball != null ? 0.32 : c && typeof c === "object" && c.bar != null ? xs[j][1] * 0.55 : 0);
      const cf = focusFsOf(c);
      return lineCount(textOf(c), cf, Math.max(0.3, w - (c && typeof c === "object" && c.focus ? 0.16 : 0.05))) * lineH(Math.min(cf, fs + 3), 1.15) + 0.2; // 18pt の数値1行は、ふつうの行の高さに収まる
    }), lineCount(r.label, lfs, lw - 0.25) * lineH(lfs, 1.15) + (r.sub ? lineH(9.5, 1.1) : 0) + 0.2);
    const rows = o.rows.map(r => ({ r, h: need(r) }));
    const judgeH = o.judge ? Math.max(0.5, need({ label: o.judge.label, cells: o.judge.cells })) : 0;
    const gCount = o.rows.filter((r, i) => r.group && i > 0).length, gLab = o.rows.filter(r => r.group).length;
    const fixedH = rows.length * rg + gCount * gsp + gLab * 0.26 + (o.judge ? 0.16 + judgeH : 0);
    const natural = rows.reduce((a, x) => a + x.h, 0) + fixedH;
    const stretch = 0; // 行は引き伸ばさない。表は詰めて置き、余った場所は空けておく（結論を横に置くなら o.note）
    rows.forEach(x => { x.h += stretch; });
    const jH = judgeH + (o.judge ? stretch : 0);
    // 見出し：選択肢名は 16pt bold（推奨の列は ink、他は soft）。単位は2行目に 9.5pt gray。札は見出しの上
    const headY = y0 - 0.66;
    o.cols.forEach((c, j) => {
      const [x0c, w0c] = xs[j], al = colAlign[j];
      // 外部の製品・会社が列の見出しなら、名前の左に公式のアイコンを置く（c.icon。名前か配列）
      const hic = [].concat(c.icon || []), hiw = hic.length * 0.36;
      hic.forEach((nm, q) => icon(s, nm, x0c + 0.15 + q * 0.36, headY + 0.56 - 0.17, 0.3));
      const x = x0c + hiw + (hiw ? 0.04 : 0), w = w0c - hiw - (hiw ? 0.04 : 0);
      if (c.sub) warnSub("表の列の見出し");
      T(s, [{ text: c.head, options: { fontSize: hfs, bold: true, color: c.strong ? C.ink : C.soft, breakLine: !!c.sub } }, ...(c.sub ? [{ text: c.sub, options: { fontSize: FS.note, color: C.gray } }] : [])],
        { x, y: headY, w, h: 0.56, align: al, valign: "bottom", lineSpacingMultiple: 1.1 });
      if (widest(c.head, hfs) > w + 0.02) warn(`表の列見出しが1行に収まらない（${c.head}）`);
      if (c.tag) T(s, c.tag, { x, y: headY - 0.26, w, h: 0.24, fontSize: FS.note, bold: true, color: C.ink, align: al, valign: "middle", wrap: false });
    });
    if (o.labelHead) T(s, o.labelHead, { x: x0, y: y0 - 0.4, w: Math.max(lw, textW(o.labelHead, FS.body) + 0.1), h: 0.3, fontSize: FS.body, color: C.soft, valign: "bottom", wrap: false });
    if (o.legend) {
      const items = [[4, "満たす"], [2, "一部"], [0, "満たさない"]];
      let lx = x0;
      // 凡例は札と同じ行の左端（行見出しの列の上）に置き、表の上に段を増やさない。推奨の列の枠と重ねない
      items.forEach(([v, t]) => { harvey(s, lx + 0.09, headY - 0.14, v, 0.16); T(s, t, { x: lx + 0.22, y: headY - 0.28, w: 1.4, h: 0.28, fontSize: FS.body, color: C.soft, valign: "middle", wrap: false }); lx += 0.28 + textW(t, FS.body) + 0.24; });
    }
    // same＝その行の中で、ほかの列と同じ値（薄くする）。違う値だけが濃く残り、違いが形に出る（o.dimSame: false でやめる）
    const drawCellR = (c, j, y, h, rowStrong, same) => {
      const [x, w] = xs[j], col = o.cols[j], soft = col.base, al = colAlign[j];
      const isNum = typeof c === "number" || (typeof c === "string" && /^[-−+]?[\d,.]+$/.test(c));
      const color = c && typeof c === "object" && c.problem ? C.accent : same ? C.gray : soft ? C.soft : C.ink;
      // 決め手のセル（c.focus）：そのセルだけ濃い塗りと ink の枠。列や行の丸ごとではなく、見る所を1つに絞るときに使う
      const cellFocus = !!(c && typeof c === "object" && c.focus);
      if (cellFocus) rect(s, x - cg / 2 + 0.03, y + 0.03, w + cg - 0.06, h - 0.06, TINT_STRONG, { color: C.emph, width: 1.75 });
      const bold = !!(c && typeof c === "object" && (c.bold || c.focus));
      if (c && typeof c === "object" && c.ball != null) {
        harvey(s, x + 0.1, y + h / 2, c.ball);
        if (c.text) T(s, c.text, { x: x + 0.3, y, w: w - 0.3, h, fontSize: fs, color: bold ? C.ink : C.soft, bold, valign: "middle", lineSpacingMultiple: 1.15 });
        return;
      }
      if (c && typeof c === "object" && c.stars != null) {
        // 星：max（既定5）個のうち n 個を塗る。図形で描く（記号の文字は、書体によって出ないことがある）
        const mx = c.max ?? 5, d0 = 0.2, gp = 0.05;
        for (let q = 0; q < mx; q++) s.addShape(SH.STAR_5_POINT, { x: x + q * (d0 + gp), y: y + h / 2 - d0 / 2, w: d0, h: d0, fill: { color: q < c.stars ? (same ? C.gray : C.emph) : C.line }, line: { type: "none" } });
        return;
      }
      if (c && typeof c === "object" && c.bar != null) {
        const bw = w * 0.55 * Math.min(1, c.bar / (c.max || 1)), bh = 0.14;
        rect(s, x, y + h / 2 - bh / 2, Math.max(bw, 0.02), bh, c.main || rowStrong ? C.emph : C.bar);
        T(s, c.text ?? String(c.bar), { x: x + bw + 0.08, y, w: w - bw - 0.08, h, fontSize: fs, color: c.main || rowStrong ? C.ink : C.soft, bold: c.main || rowStrong, valign: "middle", wrap: false });
        return;
      }
      checkWrap(textOf(c), fs, w - 0.05);
      T(s, textOf(c), { x: cellFocus ? x + 0.08 : x, y, w: w - (cellFocus ? 0.16 : 0.05), h, fontSize: focusFsOf(c), // 決め手のセルは、塗りと枠に加えて、数値と短い語は文字も大きくする
        color: c == null ? C.pale : (bold ? C.ink : color), bold, align: al, valign: "middle", lineSpacingMultiple: 1.15 });
    };
    // 行：見出しの下に線1本、行の間は薄い細線。群の名前は行の上に FS.note（16pt）gray bold、群の間は広い余白で分ける
    let y = y0; const rowY = [], bandX = x0, bandW = right - x0;
    // 表の形（2026-11-24）：行の全幅に背景色を敷かない。見出しと本文を線1本で分け、行の間は薄い細線と余白で分ける。縦線・外枠・縞模様は使わない
    hline(s, bandX, y0 - 0.05, bandW, { color: C.gray, width: 1 });
    rows.forEach(({ r, h }, i) => {
      if (r.group) { if (i > 0) y += gsp; T(s, r.group, { x: x0, y, w: lw, h: 0.24, fontSize: FS.note, bold: true, color: C.gray, valign: "middle", wrap: false }); y += 0.26; }
      if (r.strong) rect(s, bandX, y, bandW, h, TINT_STRONG, { color: C.emph, width: 1.75 });
      else if (i < rows.length - 1 && !rows[i + 1].r.strong) hline(s, bandX, y + h + rg / 2, bandW, { color: C.line, width: 0.5 });
      rowY.push(y);
      if (r.sub) warnSub("表の行の名前");
      T(s, [{ text: r.label, options: { fontSize: lfs, bold: !!r.strong, color: r.strong ? C.ink : C.soft, breakLine: !!r.sub } }, ...(r.sub ? [{ text: r.sub, options: { fontSize: FS.note, color: C.gray } }] : [])], { x: x0 + pad, y, w: lw - pad - 0.05, h, valign: "middle", lineSpacingMultiple: 1.15 });
      const cnt = {}; r.cells.forEach(c => { const t = textOf(c); cnt[t] = (cnt[t] || 0) + 1; });
      r.cells.forEach((c, j) => drawCellR(c, j, y, h, r.strong, o.dimSame !== false && r.cells.length > 2 && cnt[textOf(c)] > 1 && cnt[textOf(c)] < r.cells.length));
      y += h + rg;
    });
    let judgeY = null;
    if (o.judge) {
      y += 0.16 - rg; judgeY = y;
      hline(s, bandX, y - 0.08, bandW, { color: C.gray, width: 1 }); // 判断の行は、線1本で本文と分ける
      T(s, o.judge.label || "判断", { x: x0 + pad, y, w: lw - pad, h: jH, fontSize: lfs, color: C.soft, valign: "middle" });
      o.judge.cells.forEach((c, j) => {
        const strong = o.cols[j].strong;
        T(s, textOf(c), { x: xs[j][0], y, w: xs[j][1] - 0.05, h: jH, fontSize: fs, color: strong ? C.ink : C.soft, bold: !!strong, align: colAlign[j], valign: "middle", lineSpacingMultiple: 1.15 }); // 判断の行も、列と同じ側にそろえる
      });
      y += jH + rg;
    }
    const bottom = y - rg;
    // 読み取った結論（o.note: { head, body? }）：表の右の空きに、枠なしの文字で置く。表を引き伸ばして空きを埋める代わりに使う
    if (o.note) {
      const nx = right + 0.4, nw = x0 + wAll - nx;
      const noteW = Math.max(...String(o.note.head).split("\n").map(l => textW(l, FS.block))); // 改行した結論は、いちばん長い行で見る
      if (nw < 2.0) warn("表の横に、結論を置く幅が無いので、結論を描いていない（列を減らすか、結論はリードに書く）");
      else if (noteW > nw - 0.05) warn(`表の横の結論が折り返す（1行は全角${Math.floor((nw - 0.05) * 72 / FS.block)}字まで。言葉を縮めるか "\\n" で折る）`);
      if (nw >= 2.0) T(s, [{ text: o.note.head, options: { fontSize: FS.block, bold: true, color: C.ink, breakLine: !!o.note.body } }, ...(o.note.body ? [{ text: " ", options: { fontSize: 6, breakLine: true } }, { text: o.note.body, options: { fontSize: FS.body, color: C.soft } }] : [])],
        { x: nx, y: headY, w: nw, h: bottom - headY, valign: "middle", lineSpacingMultiple: 1.25 });
    }
    // 推奨・選ぶ列：列全体（札〜最後の行）を ink の枠1本で囲む。帯の上に重ねる
    o.cols.forEach((c, j) => { if (c.strong) { const top = headY - (c.tag ? 0.32 : 0.06); rect(s, xs[j][0] - cg / 2, top, xs[j][1] + cg, bottom - top + 0.06, null, { color: C.emph, width: 1.75 }); } });
    if (bottom > BOTTOM + 0.05) warn(`比較の表の下端 ${bottom.toFixed(2)} が本文の下端を超える（主張に効く行だけ残す。文字は小さくしない）`);
    return { xs, rowY, bottom, judgeY };
  };

  // ---- 構造の図 ----
  // 9つの部品は、どれも「薄い塗りの箱と白い溝、強調は1種類、赤は問題と視線誘導だけ」を既定にする。戻り値は echo の keep と前面の messageBox / callout から指す座標
  // 矢じりつきの線：(x1, y1) から (x2, y2) へ引き、終わりの側に矢じり。向きは flipH / flipV で表す（負の幅を作らない。pptx-notes）
  const arrowSeg = (s, x1, y1, x2, y2, line, head = true) => s.addShape(SH.LINE, { x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.max(Math.abs(x2 - x1), 0.001), h: Math.max(Math.abs(y2 - y1), 0.001), flipH: x2 < x1, flipV: y2 < y1, line: head ? { ...line, endArrowType: "triangle" } : line });
  // 鉤の折れ線：点を順に結び、最後の区間の先だけに矢じり（線は直線と鉤線だけ。4.13）
  const polyArrow = (s, pts, line, head = true) => { for (let i = 1; i < pts.length; i++) arrowSeg(s, pts[i - 1][0], pts[i - 1][1], pts[i][0], pts[i][1], line, head && i === pts.length - 1); };
  // 戻りの赤い矢印：(fx, fy) から下の通り道 cy へ下り、横へ戻って (tx, ty) へ上がる。線の中ほどに白地の赤字で注記（4.13「繰り返しの矢印」、p33）
  const backArrow = (s, fx, fy, tx, ty, cy, label, o = {}) => {
    polyArrow(s, [[fx, fy], [fx, cy], [tx, cy], [tx, ty]], { color: C.accent, width: o.width ?? 2 });
    if (!label) return;
    const fz = o.fs ?? FS.block, tw = textW(label, fz) + 0.3;
    if (tw > Math.abs(fx - tx) - 0.2) warn(`戻りの矢印の注記が線より長い（${label}）`);
    T(s, label, { x: (fx + tx) / 2 - tw / 2, y: cy - 0.17, w: tw, h: 0.34, fontSize: fz, bold: true, color: C.accent, align: "center", valign: "middle", fill: { color: C.white }, wrap: false });
  };
  // 箱の中の2段の文字（見出し bold ＋ 6pt の空行 ＋ 説明 soft。4.1-5）と、その高さの見積もり
  const boxRuns = (head, body, o = {}) => {
    const r = [{ text: head, options: { fontSize: o.headFs ?? FS.head, bold: o.bold ?? true, color: o.color ?? C.ink, breakLine: !!body } }];
    if (body) r.push({ text: " ", options: { fontSize: 6, breakLine: true } }, { text: body, options: { fontSize: o.bodyFs ?? FS.body, color: o.bodyColor ?? C.soft } });
    return r;
  };
  const boxNeed = (head, body, w, o = {}) => lineCount(head, o.headFs ?? FS.head, w) * lineH(o.headFs ?? FS.head) + (body ? lineH(6) + lineCount(body, o.bodyFs ?? FS.body, w) * lineH(o.bodyFs ?? FS.body) : 0);
  // 強調の箱：problem＝白地＋赤 2pt 枠（赤字）、strong＝濃い塗り＋ink 2pt 枠、それ以外は base の塗り（4.1-8・9）
  const emFill = (it, base) => it.problem ? C.white : it.strong ? TINT_STRONG : base;
  const emLine = it => it.problem ? { color: C.accent, width: 2 } : it.strong ? { color: C.emph, width: 2 } : null;
  const checkOneEm = (list, what) => { if (list.some(x => x && x.problem) && list.some(x => x && x.strong)) warn(`${what}に problem と strong が混ざる。強調は1ページ1種類にする`); };
  const checkBox = (head, body, w, h, o, what) => {
    checkWrap(head, o.headFs ?? FS.head, w); if (body) checkWrap(body, o.bodyFs ?? FS.body, w);
    const need = boxNeed(head, body, w, o);
    if (need > h - 0.08) warn(`${what}の文字があふれる（「${String(head).slice(0, 12)}…」 必要 ${need.toFixed(2)} > 高さ ${h.toFixed(2)}）`);
  };
  // 月の並び（"2025.04"〜"2026.03"）
  const monthList = (from, to) => { const ms = []; let [y, m] = from.split(".").map(Number); const [ty, tm] = to.split(".").map(Number); while (y < ty || (y === ty && m <= tm)) { ms.push([y, m]); m++; if (m > 12) { m = 1; y++; } } return ms; };

  // 工程の流れ（業務フロー分析）：上に工程の矢羽を左→右に並べ、その下に工程ごとの詳細（担当・所要・入力・作業・出力・システム・課題など）を同じ列幅のブロックの行で積む
  //箱の中に箇条書きを入れて横に並べた形は、本文の下半分が空いた。
  //         戦略コンサルの業務フロー分析は、工程の列ごとに観点の行を積んで1枚で工程の全体を読ませる（journeyMap と同じ組み方。罫線を使わず塗りブロックと白い溝）。行の高さは文字の量で決め、余りを配って本文の下端まで使い切る
  //   担当ごとの受け渡しが主題なら swimlane（このページの lanes は廃止）
  //   o.steps: [{ name, work?, wait?, time?, problem?, strong? }]
  //     name＝矢羽の工程名（"\n" で改行）。work / wait＝作業と待ちの日数（数値）。所要の行に合計と待ちの日数、作業（濃紺）と待ち（灰）の積み上げ棒を同じ尺度で描く。time＝合計の表記を変えるとき（"約5日"）
  //     problem＝詰まる工程（列全体を赤 2pt の枠で囲み、工程名と所要を赤字）、strong＝問題ではない主役の工程（列全体を ink 2pt の枠）。1ページでどちらか1種類だけ
  //   o.rows: [{ label, cells: [文字列 | null], h? }]（詳細の行。セルは "\n" 区切りの行。先頭が "!" の行は赤字＝問題。null は「—」）
  //   o.backs: [{ from, to, label }]（差し戻し。矢羽の上の通り道を赤で戻る。中ほどに白地の赤字で注記）
  //   o.timeLabel（所要の行の名前 "所要"）, o.unit（"日"）, o.x, o.w, o.y0（BODY）, o.bottom（BOTTOM）, o.labelW（左の名前の欄。既定は中身から、1.3 以上）, o.gap（列の間 0.08）, o.headH（0.5）, o.timeH（0.52）, o.fs（セルの文字 10.5）
  //   o.only: [工程の番号]（その列だけを描く。echo の keep で、前面から指す列を元の濃さに戻すときに使う）
  //   戻り値: { col(i), head(i), time(i), cell(r, i) → {x, y, w, h}, cx(i), rowY, rowH, bottom, box(i)（= head(i)） } … 重ね技で前面から列・セルを指すときに使う
  const processFlow = (s, o) => {
    needConcept("processFlow");
    if (o.lanes) warn("工程の流れ図：担当ごとの受け渡しは swimlane で描く（processFlow の lanes は廃止）");
    const st = o.steps, n = st.length, rows = o.rows || [], fs = o.fs ?? FS.body, unit = o.unit ?? "日";
    checkOneEm(st, "工程の流れ図");
    const hasTime = st.some(it => it.work != null || it.wait != null || it.time);
    // 左の名前の欄は、いちばん長い名前（所要の見出し・行の名前）が1行で入る幅にする
    const lw = Math.max(o.labelW ?? 1.3, labelCol("工程の流れ図", [...(hasTime ? [o.timeLabel ?? "所要", "■ 作業"] : []), ...rows.map(r => r.label || "")], o.labelW, FS.block, 2.4));
    const x0 = o.x ?? M, wAll = o.w ?? (W - M - x0), cg = o.gap ?? 0.08;
    const cx0 = x0 + lw + 0.1, cw = (x0 + wAll - cx0 - cg * (n - 1)) / n, colX = i => cx0 + i * (cw + cg);
    const backs = o.backs || [], chanH = backs.length ? 0.36 + (backs.length - 1) * 0.3 : 0;
    const y0 = o.y0 ?? BODY, bottom = o.bottom ?? BOTTOM, rg = GAP.dense;
    // マス目をやめ、矢羽を大きく濃く（灰の矢羽、見る所の工程だけ強調の色か赤）、所要は大きな数字と太い棒、詳細の行は線と余白だけ（2026-11-24）
    const total = it => (it.work ?? 0) + (it.wait ?? 0);
    const fmt = v => String(Math.round(v * 10) / 10);
    const cellText = c => c == null ? "—" : typeof c === "object" ? String(c.text ?? "") : String(c); // セルは文字列か { icons: [アイコンの名前], text }
    const plain = c => cellText(c).split("\n").map(l => l.replace(/^!/, "")).join("\n");
    { const long = rows.flatMap(r => r.cells).filter(c => c != null && (/\n/.test(cellText(c)) || em(plain(c)) > 12)).length;
      if (!o.only && (rows.length > 2 || long)) warn(`工程の流れ図の行に、文が入っている（${rows.length}行、長い札 ${long} か所）。工程ごとに置くのは、名前・所要・短い札（全角12字まで・1行、2行まで）。説明は続きのページのメッセージボックスで言う`); }
    // 縦の割り付け：基本の高さを置き、余った高さを矢羽・所要・行に配って本文の下端近くまで使う
    let hh = o.headH ?? 0.85, th = hasTime ? (o.timeH ?? 0.95) : 0, rwh = 0.5;
    const gapA = 0.14, base = chanH + hh + (hasTime ? gapA + th : 0) + (rows.length ? gapA + rows.length * rwh : 0);
    let extra = Math.max(0, bottom - y0 - base - 0.1);
    if (o.headH == null) { const d = Math.min(extra * 0.25, 0.35); hh += d; extra -= d; }
    if (hasTime && o.timeH == null) { const d = Math.min(extra * 0.4, 0.5); th += d; extra -= d; }
    if (rows.length) rwh += Math.min(extra / rows.length, 0.35);
    // 所要の段は、左の欄の見出しと凡例2行が入る高さを下限にする（下の行に重ねない）
    if (hasTime && o.timeH == null) th = Math.max(th, (FS.block + 2 * FS.note) * 1.2 / 72);
    const hy = y0 + chanH, ty = hy + hh + gapA, ry0 = (hasTime ? ty + th : hy + hh) + gapA;
    const rh = rows.map(() => rwh), rowY = rows.map((_, k) => ry0 + k * rwh);
    const figBottom = rows.length ? ry0 + rows.length * rwh : (hasTime ? ty + th : hy + hh);
    const only = o.only ? new Set(o.only) : null, show = i => !only || only.has(i);
    if (!only) {
      if (hasTime) T(s, [
        { text: o.timeLabel ?? "所要", options: { fontSize: FS.block, bold: true, color: C.ink, breakLine: true } },
        { text: "■", options: { fontSize: FS.note, color: C.ink } }, { text: " 作業", options: { fontSize: FS.note, color: C.soft, breakLine: true } },
        { text: "■", options: { fontSize: FS.note, color: C.bar } }, { text: " 待ち", options: { fontSize: FS.note, color: C.soft } },
      ], { x: x0, y: ty, w: lw, h: th, valign: "middle", lineSpacingMultiple: 1.15 });
      rows.forEach((r, k) => {
        hline(s, x0, rowY[k], cx0 + n * cw + (n - 1) * cg - x0, { color: C.line, width: 0.75 });
        T(s, r.label, { x: x0, y: rowY[k], w: lw, h: rh[k], fontSize: FS.block, bold: true, valign: "middle" }); checkWrap(r.label, FS.block, lw);
      });
    }
    const maxT = Math.max(...st.map(total), 0.001), barW = cw - 0.2;
    st.forEach((it, i) => {
      if (!show(i)) return;
      const x = colX(i), red = !!it.problem, strong = !!it.strong, on = red || strong, fg = on ? C.white : C.ink;
      // 矢羽：灰。見る所の工程だけ、強調の色（主役）か赤（問題）で塗り、文字を白に
      flowChevron(s, x, hy, cw, hh, i === 0, red ? C.accent : strong ? C.emph : C.bar, 0.18);
      T(s, [{ text: `${i + 1}`, options: { fontFace: FNUM, fontSize: FS.block, bold: true, color: fg, breakLine: true } }, { text: it.name, options: { fontSize: FS.block, bold: true, color: fg } }],
        { x: x + 0.22, y: hy, w: cw - 0.4, h: hh, align: "center", valign: "middle", lineSpacingMultiple: 1.0 });
      if (textW(String(it.name).split("\n")[0], FS.block) > cw - 0.4) warn(`工程の流れ図の工程名が矢羽に入らない（${it.name}。言葉を縮めるか "\\n" で折る）`);
      if (hasTime) {
        const tot = it.time ?? (total(it) ? `約${fmt(total(it) >= 1 ? Math.round(total(it)) : total(it))}${unit}` : "—");
        const vfs = on ? FOCUS_FS : 22, vcol = red ? C.accent : strong ? C.emph : C.ink;
        T(s, tot, { x: x + 0.1, y: ty, w: cw - 0.2, h: th * 0.55, fontSize: vfs, bold: true, color: vcol, align: "center", valign: "bottom", wrap: false });
        const k = barW / maxT, by = ty + th * 0.62, bh2 = 0.14;
        rect(s, x + 0.1, by, barW, bh2, C.tint);
        if (it.work) rect(s, x + 0.1, by, it.work * k, bh2, on ? vcol : C.ink);
        if (it.wait) rect(s, x + 0.1 + (it.work || 0) * k, by, it.wait * k, bh2, C.bar);
        if (it.wait != null) T(s, `待ち ${fmt(it.wait)}${unit}`, { x: x + 0.1, y: by + bh2 + 0.04, w: barW, h: 0.28, fontSize: FS.body, color: on ? vcol : C.soft, bold: on, align: "center", valign: "top", wrap: false });
      }
      rows.forEach((r, k) => {
        const c = r.cells[i], y = rowY[k], ls = cellText(c).split("\n");
        const cic = c && typeof c === "object" && c.icons ? c.icons : [], ciw = cic.length ? cic.length * 0.34 + 0.06 : 0;
        cic.forEach((nm, q) => icon(s, nm, x + 0.1 + 0.15 + q * 0.34, y + rh[k] / 2, 0.3));
        T(s, ls.map((l, q) => ({ text: l.replace(/^!/, ""), options: { color: c == null ? C.soft : l.startsWith("!") ? C.accent : on ? C.ink : C.soft, bold: on || l.startsWith("!"), breakLine: q < ls.length - 1 } })),
          { x: x + (ciw ? 0.1 : 0.02) + ciw, y, w: cw - (ciw ? 0.2 : 0.04) - ciw, h: rh[k], fontSize: FS.body, align: ciw ? "left" : "center", valign: "middle", lineSpacingMultiple: 1.1 });
        checkWrap(plain(c), FS.body, cw - (ciw ? 0.2 : 0.04) - ciw);
      });
      // 見る所の工程は、列全体を1つの枠で囲む（矢羽・所要・行を同じ1か所として示す）
      if (on) rect(s, x - 0.05, hy - 0.05, cw + 0.1, figBottom - hy + 0.1, null, { color: red ? C.accent : C.emph, width: 2 });
    });
    // 差し戻し：矢羽の上の通り道を赤で戻る。短い戻りほど矢羽の近くを通す（線が交わらない）
    if (!only) backs.map((b, k) => ({ ...b, k })).sort((a, b) => Math.abs(a.from - a.to) - Math.abs(b.from - b.to)).forEach((bk, q) => {
      const fx = colX(bk.from) + cw / 2, tx = colX(bk.to) + cw / 2, cy = hy - 0.2 - q * 0.3, top = hy - (st[bk.from].problem || st[bk.from].strong ? 0.05 : 0.01);
      polyArrow(s, [[fx, top], [fx, cy], [tx, cy], [tx, hy - 0.01]], { color: C.accent, width: 2 });
      if (bk.label) {
        const lw2 = textW(bk.label, FS.block) + 0.3;
        if (lw2 > Math.abs(fx - tx) - 0.2) warn(`工程の流れ図の差し戻しの注記が線より長い（${bk.label}）`);
        T(s, bk.label, { x: (fx + tx) / 2 - lw2 / 2, y: cy - 0.16, w: lw2, h: 0.32, fontSize: FS.block, bold: true, color: C.accent, align: "center", valign: "middle", fill: { color: C.white }, wrap: false });
      }
    });
    if (figBottom > BOTTOM + 0.05) warn(`工程の流れ図の下端 ${figBottom.toFixed(2)} が本文の下端を超える`);
    const head = i => ({ x: colX(i), y: hy, w: cw, h: hh });
    return {
      col: i => ({ x: colX(i), y: hy, w: cw, h: figBottom - hy }), head, box: head,
      time: i => ({ x: colX(i), y: ty, w: cw, h: th }), cell: (r, i) => ({ x: colX(i), y: rowY[r], w: cw, h: rh[r] }),
      cx: i => colX(i) + cw / 2, rowY, rowH: rh, bottom: figBottom,
    };
  };



  // 2軸の図：ポジショニングマップと、項目を入れる2×2を1つの関数で描く
  //   チープに見えた原因と対策:
  //     点が小さく全部同じ（0.16in の灰の点）     → 点を第3の量（売上・利用者数）の大きさのバブルにする。量が無くても 0.34in 以上
  //     名前が薄く小さい（灰の 10.5pt）           → 名前は ink 16pt（FS.body）、自社は 18pt bold（FS.block）。補足の2段目（sub）は使わない
  //     名前が点・軸・他の名前と重なる            → 8方向の候補から、重なりの一番少ない位置を自動で選ぶ（pos で固定もできる）
  //     象限が空の箱で、軸の意味が角の小さな灰字   → 軸は中央で交わる両矢印、両端に意味を bold で書き、高い側に軸の名前を添える
  //     空白地帯が細い破線の楕円と小さな灰字       → 地の上に白く抜いた楕円（＝誰もいない所）に、見出しと説明を ink／soft で
  //     右側が空く                                → side に読み取り（示唆）のブロックとバブルの凡例を置き、本文の領域を使い切る
  //   2つの形（quads に items か policy があれば2×2、無ければポジショニングマップ）
  //   ■ ポジショニングマップ（o.points）
  //     地は tint 1枚、軸は中央で交わる灰の両矢印。強調は1種類：自社の点（濃紺・太字）か、狙う象限（target：濃い塗り＋ink 枠）のどちらか
  //     o.xAxis / o.yAxis: { name（何を測るか）, low, high（両端の意味。右・上ほど high） }
  //     o.points: [{ name, sub?（区分・量などの補足1行）, x, y（0〜1）, size?（第3の量）, self?, pos?: "r"|"l"|"t"|"b"|"tr"|"br"|"tl"|"bl" }]
  //     o.size: { label（凡例の見出し「円の大きさ＝利用企業数」）, unit（"社"）, ref: [凡例の値, …], max?（最大の円に当たる値。既定は点の最大） }
  //       円の面積を量に比例させる（直径 ∝ √量）。最大の直径 o.dMax（0.95）、最小 0.2
  //     o.quads: [左上, 右上, 左下, 右下]（象限の名前。文字列か { name, body }。外側の角に置く）, o.target（狙う象限 0〜3）
  //     o.zone: { x, y（中心）, w, h（0〜1 の幅。rx, ry でも可）, label, body? }（空白地帯。白く抜いた破線の楕円）
  //     o.move: { from: 点の番号, to?: [x, y]（既定は zone の縁） }（自社が向かう先。紺の破線の矢印）
  //     o.side: [{ head, body }]（右の読み取り。上にバブルの凡例、その下にブロックを本文の下端まで）, o.sideW（3.3）
  //   ■ 2×2（o.quads: [{ name, policy?, items: [文字列 | { name, sub }] }]）
  //     象限は塗りブロックと白い溝、軸は左と下の灰の矢印。各象限に名前（狙う象限は大きく）、方針の1〜2行、項目の白い札（2列）
  //     項目の位置は象限の中で意味を持たせない（点の細かな位置で読ませない）。o.itemCols（2）, o.labelW（左の軸の列 1.0）
  //   共通: o.x（M）, o.y（2.2）, o.w（本文幅）, o.bottom（6.6）
  //   戻り値: マップ { px(v), py(v), pt(i) → [x, y], d(i), quad(i) → {x, y, w, h} } / 2×2 { quad(i), item(i, k) → {x, y, w, h} }
  const matrix2x2 = (s, o) => {
    chartPage();
    const Q = (o.quads || []).map(q => q == null ? null : typeof q === "string" ? { name: q } : { ...q });
    (o.quadBody || []).forEach((b, i) => { if (Q[i] && b && !Q[i].body) Q[i].body = b; });
    return Q.some(q => q && (q.items || q.policy)) ? mxGrid(s, o, Q) : mxMap(s, o, Q);
  };

  // ファネル：中心をそろえた台形の段を上から下へ重ね、右に件数・率・段の間の歩留まりをそろえて置く。その右に段ごとの説明の列
  //   前の版は棒の高さを率に比例させ、420件→18件で棒が線になり、帯の傾きも段ごとにばらばらだった
  //   形の判断：段の幅は値に比例させない。上端の幅から、段ごとに一定の比率 k で狭める（最後の段の下端＝上端 × minRatio、k＝minRatio^(1/段数)）。
  //     値に比例させると、4%の段が 0.2in の線になり名前が入らない。率は形ではなく右の数字で読ませる（形は「段を経て減る」ことだけを表す）
  //   段は灰（C.bar）で塗り、段名は段の中に bold ink。濃紺で塗らない（文字を載せる面のため）。落ちる所（o.drop）だけ歩留まりを赤枠＋赤字にする。
  //   次の案件で使う目安なら rate だけを渡し、drop を付けない（今回の実績を赤で強調すると「今回が悪かった」印象になる）
  //   o.stages: [{ name, value?, rate? }]（rate は 0〜1。省くと value ÷ 最初の value）, o.unit（"件"）, o.fmt
  //   o.drop（落ちる所：段 i と i+1 の間なら i）
  //   o.cols: [{ head, cells: [文字列 | { head, body, strong?, problem? }] }]（右の説明の列。段と同じ高さのブロック。cells は段の数だけ）
  //   o.labels: { value（既定は単位＋「数」（件数・人数）か「比率（目安）」）, conv（"次へ進む率"） }
  //   o.x（M）, o.y（2.2）, o.w（本文幅）, o.bottom（6.6）, o.funnelW（説明の列があれば 4.0、無ければ 6.0）, o.minRatio（0.32）, o.gap（段の間 0.06）
  //   戻り値: { band(i) → {x, y, w, h, top, bot}（top/bot は段の上端・下端の幅）, conv(i) → [x, y], cell(j, i) → {x, y, w, h} }
  const funnel = (s, o) => {
    chartPage();
    const st = o.stages, n = st.length, cols = o.cols || [];
    const x0 = o.x ?? M, right = x0 + (o.w ?? (W - M - x0)), top = o.y ?? BODY, bottom = o.bottom ?? BOTTOM;
    const headH = 0.32, by0 = top + headH + 0.12, gap = o.gap ?? 0.06, bh = (bottom - by0 - gap * (n - 1)) / n;
    // 率は1つに絞る（2026-11-24）：既定は「次へ進む率」（どこで落ちるか）。o.rate: "fromTop" なら「最初からの率」だけ。件数の無い目安のファネルは比率の列を出す
    const rateMode = o.rate ?? "next", noVal = st[0].value == null, showVal = rateMode === "fromTop" || noVal, showConv = rateMode !== "fromTop";
    if (rateMode === "fromTop" && o.drop !== undefined) warn("ファネル：最初からの率（rate: fromTop）では、落ちる所（drop）の印を付けられない。次へ進む率にする");
    const fw = o.funnelW ?? (cols.length ? 4.0 : 6.0), vw = showVal ? (o.valueW ?? 1.3) : 0, cvw = showConv ? (o.convW ?? 1.25) : 0;
    const vx = x0 + fw + 0.2, cvx = showVal ? vx + vw + 0.1 : vx, dx0 = (showConv ? cvx + cvw : vx + vw) + 0.3;
    const cg = GAP.std, dcw = cols.length ? (right - dx0 - cg * (cols.length - 1)) / cols.length : 0;
    const minR = o.minRatio ?? 0.32, k = Math.pow(minR, 1 / n), wAt = i => fw * Math.pow(k, i), fcx = x0 + fw / 2;
    const bandY = i => by0 + i * (bh + gap);
    const v0 = st[0].value, rate = it => it.rate ?? (it.value != null && v0 ? it.value / v0 : 1);
    const fmt = o.fmt || (v => v.toLocaleString()), unit = o.unit ?? "件", pct = r => `${Math.round(r * 100)}%`;
    const cellsAll = cols.flatMap(c => c.cells || []).filter(c => c && typeof c === "object");
    checkOneEm(cellsAll, "ファネルの説明の列");
    if (o.drop !== undefined && cellsAll.some(c => c.strong)) warn("ファネル：落ちる所（赤）と説明の列の strong（紺）を同じページで使っている。強調は1種類にする");
    // 見出しの行
    const L = o.labels || {}, hy = top;
    if (showVal) T(s, L.value ?? (v0 != null ? "最初からの率" : "比率（目安）"), { x: vx, y: hy, w: vw, h: headH, fontSize: FS.block, bold: true, align: "right", valign: "bottom", wrap: false });
    if (showConv) T(s, L.conv ?? "次へ進む率", { x: cvx, y: hy, w: cvw, h: headH, fontSize: FS.block, bold: true, align: "center", valign: "bottom", wrap: false });
    cols.forEach((c, j) => c.head && T(s, c.head, { x: dx0 + j * (dcw + cg), y: hy, w: dcw, h: headH, fontSize: FS.block, bold: true, valign: "bottom", wrap: false }));
    const band = [];
    st.forEach((it, i) => {
      const y = bandY(i), wt = wAt(i), wb = wAt(i + 1);
      s.addShape(SH.CUSTOM_GEOMETRY, { x: x0, y, w: fw, h: bh, fill: { color: o.color ?? C.bar }, line: { type: "none" },
        points: [{ x: (fw - wt) / 2, y: 0 }, { x: (fw + wt) / 2, y: 0 }, { x: (fw + wb) / 2, y: bh }, { x: (fw - wb) / 2, y: bh }, { close: true }] });
      // 段の中に、段名と件数（2026-11-24。件数は段の外に離さない）
      const inVal = it.value != null ? `${fmt(it.value)}${unit}` : "";
      T(s, [{ text: it.name, options: { fontSize: FS.block, bold: true, color: C.ink, breakLine: !!inVal } }, ...(inVal ? [{ text: inVal, options: { fontSize: FS.lead, bold: true, color: C.ink } }] : [])],
        { x: fcx - wb / 2, y, w: wb, h: bh, align: "center", valign: "middle", wrap: false, lineSpacingMultiple: 1.05 });
      if (Math.max(textW(it.name, FS.block), textW(inVal, FS.lead)) > wb - 0.1) warn(`ファネルの段名が段に収まらない（${it.name}）。minRatio を上げるか funnelW を広げる`);
      band.push({ x: fcx - wt / 2, y, w: wt, h: bh, top: wt, bot: wb });
      // 件数と率（右揃え）：件数を主に、最初の段からの率を下に小さく
      const runs = [{ text: pct(rate(it)), options: { fontSize: FS.block, bold: !i, color: C.ink } }]; // 右の列は、最初の段からの率
      if (showVal) T(s, runs, { x: vx, y, w: vw, h: bh, align: "right", valign: "middle", lineSpacingMultiple: 1.1 });
    });
    // 段の間の歩留まり：段の境目の高さに、下向きの小さな三角と率。落ちる所だけ赤枠＋赤字
    const conv = [];
    for (let i = 0; showConv && i < n - 1; i++) {
      const cy = bandY(i + 1) - gap / 2, r = rate(st[i + 1]) / rate(st[i]), t = pct(r), drop = o.drop === i, ccx = cvx + cvw / 2;
      const fz = drop ? FS.lead : FS.block, tw = textW(t, fz) + 0.06, tri = 0.14, tot = tri + 0.08 + tw, lx = ccx - tot / 2;
      if (drop) rect(s, lx - 0.14, cy - 0.21, tot + 0.28, 0.42, C.white, { color: C.accent, width: 2 });
      PT(s, "d", lx, cy - 0.05, tri, 0.1, drop ? C.accent : C.gray);
      T(s, t, { x: lx + tri + 0.08, y: cy - 0.17, w: tw, h: 0.34, fontSize: fz, bold: drop, color: drop ? C.accent : C.ink, valign: "middle", wrap: false });
      conv.push([ccx, cy]);
    }
    // 説明の列：段と同じ高さの塗りブロック
    const cell = (j, i) => ({ x: dx0 + j * (dcw + cg), y: bandY(i), w: dcw, h: bh });
    cols.forEach((c, j) => (c.cells || []).forEach((v, i) => {
      if (v == null || i >= n) return;
      const r = cell(j, i), it = typeof v === "string" ? { body: v } : v, iw = r.w - 2 * PAD.x;
      rect(s, r.x, r.y, r.w, r.h, emFill(it, TINT), emLine(it));
      const runs = it.head ? boxRuns(it.head, it.body, { headFs: FS.block, color: it.problem ? C.accent : C.ink }) : [{ text: it.body, options: { fontSize: FS.body, color: C.ink } }];
      T(s, runs, { x: r.x + PAD.x, y: r.y, w: iw, h: r.h, valign: "middle", lineSpacingMultiple: 1.15 });
      // 高さの見積もり（行間 1.15 の実測に合わせ、1行＝文字の大きさ×1.25）
      const need = (it.head ? lineCount(it.head, FS.block, iw) * FS.block * 1.25 / 72 + 6 * 1.25 / 72 : 0) + (it.body ? lineCount(it.body, FS.body, iw) * FS.body * 1.25 / 72 : 0);
      if (need > r.h - 0.06) warn(`ファネルの説明の文字があふれる（「${String(it.head || it.body).slice(0, 12)}…」 必要 ${need.toFixed(2)} > 高さ ${r.h.toFixed(2)}）`);
      if (it.head) checkWrap(it.head, FS.block, iw); if (it.body) checkWrap(it.body, FS.body, iw);
    }));
    if (!cols.length && vx + vw + cvw > right) warn("ファネル：数字の列が本文の幅を超える");
    return { band: i => band[i], conv: i => conv[i], cell };
  };

  // アーキテクチャ図：AWS の構成図のように、アイコンのノード・まとまりの領域・鉤の矢印で仕組みを描く
  //
  //   ノード＝アイコンの札（正方形）＋名前 FS.block（18pt）bold＋短い説明 FS.body（16pt）soft。アイコンは opts.iconDir の PNG（icons/make-icons.js が作る。ブランドは実物の色、汎用は ink）
  //   まとまり（社内・AWS・外部のサービスなど）は、薄い塗りの領域（style "fill"）か破線の境界（style "dash"）で囲み、左上に印と名前。入れ子にしてよい（AWS の中の社内の網など）
  //   線は鉤に折った矢印（soft 1.25pt）。社外・非同期のやり取りは灰の破線。ラベルは線の一番長い区間の脇に FS.body（16pt）soft で短く
  //   強調は1種類：主役のノード（strong＝濃い塗り＋ink 枠）か、問題のノード（problem＝赤枠＋赤字）か、線（strong＝ink 2pt／problem＝赤 2pt）
  //   示唆は前面の番号つきの箱（o.notes、または同じ図を echo で敷いて戻り値の note で重ねる）。本筋でない区画（計測・運用など）は描かない
  //   座標は格子：本文の領域（x, y, w, h）を cols 列 × rows 行に割り、ノードは列 c・行 r（0 始まり。0.5 刻みなど小数も可）の升の中央に置く
  //   o.cols, o.rows（既定はノードの c・r の最大 + 1）, o.x（M）, o.y（2.2）, o.w（本文幅）, o.h（6.6 − y）
  //   o.groups: [{ id?, name, sub?, icon?, c: [c0, c1], r: [r0, r1]（升の範囲。端を含む。小数も可）, style?: "fill"|"dash", inset?（升の縁からの内側 0.08） }]
  //     描く順は配列の順（外側のまとまりを先に）。名前の高さ 0.4 を上に取るので、最上段のノードは r を 0.2〜0.4 ほど下げる
  //   o.nodes: [{ id, c, r, icon, name, sub?, strong?, problem? }]（sub は "\n" で2行まで）
  //   o.edges: [{ from, to, label?, dashed?, both?（両向き）, strong?, problem?, route?, mid?, via?, fromOff?, toOff?, labelPos?, labelSeg? }]
  //     route: "h"（横一直線）| "v"（縦一直線）| "hvh"（横→縦→横。縦の区間は列 mid、既定は2つの間）| "vhv"（縦→横→縦。横の区間は行 mid）| "hv" | "vh"（L字）
  //       省くと、同じ行なら h、同じ列なら v、それ以外は hvh。via: [[c, r], …] で通る点を直接指定（最初と最後の点はノードと縦か横にそろえる）
  //     fromOff / toOff: 札の辺に沿って線の出入り口をずらす（in。平行な線を分ける）。labelPos: "a"（上）| "b"（下）| "l" | "r"。labelSeg: ラベルを置く区間の番号
  //   o.notes: [{ num, head, body?, x, y, w, h?, to: id }]（前面の示唆。白地に ink 枠の箱から、to のノードへ赤の線。箱が下なら名前の下端、上・横なら札の縁を指す）
  //   o.tile（札の一辺 0.6）, o.only: [id]（そのノードだけを描く。echo の keep で、前面から指すノードを元の濃さに戻すときに使う）
  //   戻り値: { node(id) → {x, y, w, h, cx, cy, bottom}（札の矩形、bottom は名前の下端）, group(id) → {x, y, w, h}, cellX(c), cellY(r), note(s, 示唆) }
  const archDiagram = (s, o) => {
    needConcept("archDiagram");
    if (o.layers) throw new Error("archDiagram: 層の帯（layers・kind）の形は 2026-11-08 に廃止。groups と nodes の c・r・icon で描く（sample-map.js）");
    const nodes = o.nodes || [], x0 = o.x ?? M, y0 = o.y ?? BODY, w = o.w ?? (W - M - x0), h = o.h ?? (BOTTOM - y0);
    const nc = o.cols ?? Math.ceil(Math.max(...nodes.map(n => n.c), ...(o.groups || []).map(g => g.c[1])) + 1);
    const nr = o.rows ?? Math.ceil(Math.max(...nodes.map(n => n.r), ...(o.groups || []).map(g => g.r[1])) + 1);
    const cw = w / nc, rh = h / nr, cellX = c => x0 + (c + 0.5) * cw, cellY = r => y0 + (r + 0.5) * rh;
    const ts = o.tile ?? 0.6, shift = o.shift ?? 0.2, only = o.only ? new Set(o.only) : null;
    checkOneEm([...nodes, ...(o.edges || [])], "構成図");
    // まとまり
    const G = {}, fills = [];
    (o.groups || []).forEach((g, gi) => {
      const ins = g.inset ?? 0.08, gx = x0 + g.c[0] * cw + ins, gy = y0 + g.r[0] * rh + ins, gw = (g.c[1] - g.c[0] + 1) * cw - 2 * ins, gh = (g.r[1] - g.r[0] + 1) * rh - 2 * ins;
      const r = { x: gx, y: gy, w: gw, h: gh, style: g.style || "fill" };
      G[g.id ?? gi] = r;
      if (r.style === "fill") fills.push(r);
      if (only) return;
      if (r.style === "fill") rect(s, gx, gy, gw, gh, TINT);
      else s.addShape(SH.RECTANGLE, { x: gx, y: gy, w: gw, h: gh, fill: { type: "none" }, line: { color: C.soft, width: 1, dashType: "dash" } });
      let lx = gx + 0.14;
      if (g.icon && icon(s, g.icon, lx + 0.13, gy + 0.22, 0.26)) lx += 0.34;
      const runs = [{ text: g.name, options: { fontSize: FS.block, bold: true, color: C.ink } }];
      if (g.sub) runs.push({ text: `　${g.sub}`, options: { fontSize: FS.body, color: C.soft } });
      const lw = textW(g.name, FS.block) + (g.sub ? textW(`　${g.sub}`, FS.body) : 0) + 0.06;
      T(s, runs, { x: lx, y: gy + 0.08, w: Math.min(lw, gx + gw - lx - 0.05), h: 0.28, valign: "middle", wrap: false });
      if (lx + lw > gx + gw) warn(`構成図：まとまりの名前が枠に入らない（${g.name}）`);
      r.label = [gx, gy, lx + lw - gx, 0.38];
    });
    // ノードの位置：札の中心は升の中央から shift だけ上（名前を下に置くため）。同じ行の札は同じ高さになる
    const N = {};
    const inFill = (px, py) => fills.some(f => px > f.x && px < f.x + f.w && py > f.y && py < f.y + f.h);
    nodes.forEach(nd => {
      const cx = cellX(nd.c), cy = cellY(nd.r) - shift, sub = nd.sub ? (warnSub("構成図のノード"), String(nd.sub)) : "";
      const nameH = lineH(FS.block, 1.1), subH = sub ? sub.split("\n").length * lineH(FS.body, 1.1) : 0;
      N[nd.id] = { x: cx - ts / 2, y: cy - ts / 2, w: ts, h: ts, cx, cy, bottom: cy + ts / 2 + 0.06 + nameH + subH, nameH, subH, nd };
    });
    const node = id => { const r = N[id]; if (!r) throw new Error(`archDiagram: ノード ${id} が無い`); return r; };
    // 点検：名前が升の幅に収まるか、札・名前がまとまりの名前に重ならないか
    if (!only) nodes.forEach(nd => {
      const r = N[nd.id], lw = Math.max(textW(nd.name, FS.block), ...(nd.sub ? String(nd.sub).split("\n").map(l => textW(l, FS.body)) : [0]));
      if (lw > cw - 0.08) warn(`構成図：ノードの名前が升の幅（${cw.toFixed(2)}）に入らない（${nd.name}）。"\\n" で折るか cols を減らす`);
      const box = [r.cx - Math.max(ts, lw) / 2, r.y, Math.max(ts, lw), r.bottom - r.y];
      Object.values(G).forEach(g => { if (g.label && mxRectHit(box, g.label) > 0.001) warn(`構成図：「${nd.name}」がまとまりの名前に重なる。r を下げる`); });
      if (r.bottom > y0 + h + 0.02) warn(`構成図：「${nd.name}」の名前が図の下端を超える`);
    });
    // 線の端：横に出入りするときは札の左右の辺（札の中心の高さ＋off）、縦は上は札の上辺・下は名前の下端
    const sideP = (r, toward, off = 0) => [toward >= r.cx ? r.x + r.w : r.x, r.cy + off];
    const vertP = (r, toward, off = 0) => [r.cx + off, toward >= r.cy ? r.bottom + 0.04 : r.y];
    const routeOf = e => {
      const a = node(e.from), b = node(e.to), fo = e.fromOff || 0, to = e.toOff || 0;
      if (e.via) {
        const V = e.via.map(([c, r]) => [cellX(c), cellY(r) - shift]);
        const f = V[0], l = V[V.length - 1];
        const p0 = Math.abs(f[1] - a.cy) < 0.03 ? sideP(a, f[0], fo) : vertP(a, f[1], fo);
        const p1 = Math.abs(l[1] - b.cy) < 0.03 ? sideP(b, l[0], to) : vertP(b, l[1], to);
        if (Math.abs(f[1] - a.cy) < 0.03) V[0] = [f[0], p0[1]]; else V[0] = [p0[0], f[1]];
        if (Math.abs(l[1] - b.cy) < 0.03) V[V.length - 1] = [l[0], p1[1]]; else V[V.length - 1] = [p1[0], l[1]];
        return [p0, ...V, p1];
      }
      const rt = e.route || (Math.abs(a.cy - b.cy) < 0.03 ? "h" : Math.abs(a.cx - b.cx) < 0.03 ? "v" : "hvh");
      if (rt === "h") { const p = sideP(a, b.cx, fo); return [p, [sideP(b, a.cx)[0], p[1]]]; }
      if (rt === "v") { const p = vertP(a, b.cy, fo); return [p, [p[0], vertP(b, a.cy)[1]]]; }
      if (rt === "hv") { const p = sideP(a, b.cx, fo), q = vertP(b, a.cy, to); return [p, [q[0], p[1]], q]; }
      if (rt === "vh") { const p = vertP(a, b.cy, fo), q = sideP(b, a.cx, to); return [p, [p[0], q[1]], q]; }
      if (rt === "vhv") { const p = vertP(a, b.cy, fo), q = vertP(b, a.cy, to), ym = e.mid != null ? cellY(e.mid) : (p[1] + q[1]) / 2; return [p, [p[0], ym], [q[0], ym], q]; }
      const p = sideP(a, b.cx, fo), q = sideP(b, a.cx, to), xm = e.mid != null ? cellX(e.mid) : (p[0] + q[0]) / 2;
      return [p, [xm, p[1]], [xm, q[1]], q];
    };
    // ノードの占める矩形（札＋名前）。線とラベルがほかのノードを横切らないかの点検に使う
    // 札の矩形と、その下の名前の矩形の2つ
    const nboxes = r => { const lw = Math.max(textW(r.nd.name, FS.block), ...(r.nd.sub ? String(r.nd.sub).split("\n").map(l => textW(l, FS.body)) : [0])); return [[r.x, r.y, r.w, r.h], [r.cx - lw / 2, r.y + r.h, lw, r.bottom - r.y - r.h]]; };
    const hitNode = (r, b) => nboxes(r).some(nb => mxRectHit(nb, b) > 0.0005);
    const segBox = (p, q) => [Math.min(p[0], q[0]) - 0.01, Math.min(p[1], q[1]) - 0.01, Math.abs(q[0] - p[0]) + 0.02, Math.abs(q[1] - p[1]) + 0.02];
    const EP = only ? [] : (o.edges || []).filter(e => {
      if (!N[e.from] || !N[e.to]) { warn(`構成図：線の端のノードが無い（${e.from} → ${e.to}）`); return false; }
      return true;
    }).map(e => ({ e, pts: routeOf(e).filter((p, i, a) => i === 0 || Math.hypot(p[0] - a[i - 1][0], p[1] - a[i - 1][1]) > 0.005) }));
    EP.forEach(({ e, pts }) => {
      for (let i = 1; i < pts.length; i++) Object.values(N).forEach(r => {
        if (r.nd.id === e.from || r.nd.id === e.to) return;
        if (hitNode(r, segBox(pts[i - 1], pts[i]))) warn(`構成図：線 ${e.from} → ${e.to} が「${r.nd.name}」を横切る。route・mid・via を変える`);
      });
    });
    EP.forEach(({ e, pts }) => {
      const line = e.problem ? { color: C.accent, width: 2 } : e.strong ? { color: C.emph, width: 2 } : e.dashed ? { color: C.gray, width: 1.25, dashType: "dash" } : { color: C.soft, width: 1.25 };
      for (let i = 1; i < pts.length; i++) {
        const [x1, y1] = pts[i - 1], [x2, y2] = pts[i], last = i === pts.length - 1, first = i === 1;
        s.addShape(SH.LINE, { x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.max(Math.abs(x2 - x1), 0.001), h: Math.max(Math.abs(y2 - y1), 0.001), flipH: x2 < x1, flipV: y2 < y1,
          line: { ...line, ...(last ? { endArrowType: "triangle" } : {}), ...(first && e.both ? { beginArrowType: "triangle" } : {}) } });
      }
      if (e.label) {
        // ラベルの置き場：区間を、まとまりの境界とほかの線が横切る所で切り、一番長い切れ目の中央に置く（labelSeg で区間を固定できる）
        const cutsOf = (p, q) => {
          const horiz = Math.abs(p[1] - q[1]) < 0.01, a = horiz ? Math.min(p[0], q[0]) : Math.min(p[1], q[1]), b = horiz ? Math.max(p[0], q[0]) : Math.max(p[1], q[1]), c = horiz ? p[1] : p[0];
          const cuts = [a, b];
          Object.values(G).forEach(g => { if (horiz ? c > g.y && c < g.y + g.h : c > g.x && c < g.x + g.w) (horiz ? [g.x, g.x + g.w] : [g.y, g.y + g.h]).forEach(v => { if (v > a && v < b) cuts.push(v); }); });
          EP.forEach(qq => { if (qq.e === e) return; qq.pts.forEach((pp, i) => { if (!i) return; const p0 = qq.pts[i - 1]; const crossV = horiz ? Math.abs(p0[0] - pp[0]) < 0.01 && c >= Math.min(p0[1], pp[1]) - 0.01 && c <= Math.max(p0[1], pp[1]) + 0.01 : Math.abs(p0[1] - pp[1]) < 0.01 && c >= Math.min(p0[0], pp[0]) - 0.01 && c <= Math.max(p0[0], pp[0]) + 0.01; if (crossV) { const v = horiz ? p0[0] : p0[1]; if (v > a && v < b) cuts.push(v); } }); });
          cuts.sort((u, v) => u - v);
          let best = [a, b, -1]; for (let i = 1; i < cuts.length; i++) if (cuts[i] - cuts[i - 1] > best[2]) best = [cuts[i - 1], cuts[i], cuts[i] - cuts[i - 1]];
          return { horiz, c, mid: (best[0] + best[1]) / 2, len: best[2] };
        };
        const segs = e.labelSeg != null ? [e.labelSeg + 1] : pts.map((_, i) => i).slice(1);
        const cand = segs.map(i => cutsOf(pts[i - 1], pts[i])).sort((u, v) => v.len - u.len)[0];
        const horiz = cand.horiz, tw = textW(e.label, FS.body) + 0.08, mx = horiz ? cand.mid : cand.c, my = horiz ? cand.c : cand.mid;
        const pos = e.labelPos || (horiz ? "a" : "r"), col = e.problem ? C.accent : C.soft;
        if (horiz && tw > cand.len - 0.12) warn(`構成図：線のラベルが線より長い（${e.label}）。境界やほかの線で切れた区間に入らない`);
        const box = pos === "a" ? [mx - tw / 2, my - 0.29, tw, 0.26] : pos === "b" ? [mx - tw / 2, my + 0.04, tw, 0.26] : pos === "l" ? [mx - 0.08 - tw, my - 0.13, tw, 0.26] : [mx + 0.08, my - 0.13, tw, 0.26];
        const hitLine = EP.some(q => q.e !== e && q.pts.some((pp, i) => i > 0 && mxRectHit(segBox(q.pts[i - 1], pp), box) > 0.0005));
        if (hitLine || Object.values(N).some(r => hitNode(r, box))) warn(`構成図：線のラベル「${e.label}」がほかの線かノードに重なる。labelPos・labelSeg を変える`);
        T(s, e.label, { x: box[0], y: box[1], w: box[2], h: box[3], fontSize: FS.body, color: col, align: pos === "l" ? "right" : pos === "r" ? "left" : "center", valign: pos === "a" ? "bottom" : pos === "b" ? "top" : "middle", wrap: false });
      }
    });
    // ノード：札（塗りの領域の中は白、外は tint）にアイコン、下に名前と説明
    Object.values(N).forEach(r => {
      const nd = r.nd;
      if (only && !only.has(nd.id)) return;
      const base = inFill(r.cx, r.cy) ? C.white : TINT;
      rect(s, r.x, r.y, r.w, r.h, emFill(nd, base), emLine(nd));
      icon(s, nd.icon, r.cx, r.cy, ts * (o.iconScale ?? 0.6));
      const runs = [{ text: nd.name, options: { fontSize: FS.block, bold: true, color: nd.problem ? C.accent : C.ink, breakLine: !!nd.sub } }];
      if (nd.sub && (warnSub("図の要素"), true)) runs.push({ text: String(nd.sub), options: { fontSize: FS.body, color: C.soft } });
      T(s, runs, { x: r.cx - cw / 2 + 0.03, y: r.y + r.h + 0.06, w: cw - 0.06, h: r.nameH + r.subH + 0.04, align: "center", valign: "top", lineSpacingMultiple: 1.1 });
    });
    // 前面の示唆：箱から、指すノードの札の一番近い所へ赤の線（messageBox の to）
    const note = (s2, nt) => {
      const r = node(nt.to), nh = nt.h ?? (nt.body ? 1.0 : 0.6), bcx = nt.x + nt.w / 2, bcy = nt.y + nh / 2;
      // 箱がノードより下にあれば名前の下端を、上・横にあれば札の縁を指す（線が名前を横切らないように）
      const to = nt.y > r.bottom ? [r.cx, r.bottom + 0.03] : [Math.min(Math.max(bcx, r.x), r.x + r.w), Math.min(Math.max(bcy, r.y), r.y + r.h)];
      const side = to[0] > nt.x + nt.w ? "r" : to[0] < nt.x ? "l" : to[1] < nt.y ? "t" : "b";
      const from = side === "r" || side === "l" ? [side === "r" ? nt.x + nt.w : nt.x, Math.min(Math.max(to[1], nt.y + 0.15), nt.y + nh - 0.15)] : [Math.min(Math.max(to[0], nt.x + 0.15), nt.x + nt.w - 0.15), side === "t" ? nt.y : nt.y + nh];
      messageBox(s2, { ...nt, h: nh, to, from });
    };
    if (!only) (o.notes || []).forEach(nt => note(s, nt));
    return { node, group: id => G[id], cellX, cellY, note };
  };

  // 時期の点の図：年月の軸の上に出来事を点で置き、名前は点の真下。近いものは段をずらして細い引出線でつなぐ。自社だけ濃紺・太字」）
  //   番号つきの丸と下の一覧で対応づけない（直感的に分からない）。行を分けるときは、出来事・判断・対象を同じ行に混ぜない
  //   o.from, o.to（"2019.01" の形。列は月）, o.rows: [{ label?, events: [{ at, name, self? }] }]（1行なら o.events だけでよい）
  //   o.y0（年の目盛りの上端 2.3）, o.levels（名前の段の数 3）, o.labelW（行の名前の列。既定は中身から、1.45 以上）, o.bottom
  //   戻り値: { xOf(at), axisY(i), dot(ri, ei) → [x, y] }
  const eventTimeline = (s, o) => {
    const months = monthList(o.from, o.to), rows = o.rows || [{ events: o.events }], hasLabel = rows.some(r => r.label);
    const lw = hasLabel ? Math.max(o.labelW ?? 1.45, labelCol("時期の点の図", rows.map(r => r.label || ""), o.labelW, FS.head, 2.4)) : 0, x0 = M + lw + (hasLabel ? 0.1 : 0), x1 = W - M, mw = (x1 - x0) / months.length;
    const xOf = p => { const [y, m] = p.split(".").map(Number), i = months.findIndex(([yy, mm]) => yy === y && mm === m); if (i < 0) throw new Error(`eventTimeline: ${p} は ${o.from}〜${o.to} の外`); return x0 + (i + 0.5) * mw; };
    const y0 = o.y0 ?? 2.3, levels = o.levels ?? 3, step = 0.36, need = 0.2 + levels * step + 0.15;
    const rowsTop = y0 + 0.4, avail = (o.bottom ?? BOTTOM) - rowsTop, rowH = Math.max(need, avail / rows.length);
    if (need * rows.length > avail + 0.05) warn("時期の点の図：行が本文の下端を超える（段の数か行を減らす）");
    const axisY = i => rowsTop + i * rowH + (rowH - need) / 2 + 0.12, last = rowsTop + rows.length * rowH;
    // 年の目盛り：年の始まりに細い縦線、上に年。隣の年と詰まるときは「年」を省き、それでも詰まれば何年かおきに書く
    const yrW = 12 * mw, withNen = textW("0000年", FS.note) + 0.12 <= yrW, yw = textW(withNen ? "0000年" : "0000", FS.note) + 0.12, every = Math.max(1, Math.ceil(yw / yrW));
    const jan = months.findIndex(([, m], i) => i > 0 && m === 1);
    months.forEach(([y, m], i) => {
      if (i !== 0 && m !== 1) return;
      const x = x0 + i * mw;
      s.addShape(SH.LINE, { x, y: y0 + 0.26, w: 0, h: last - (y0 + 0.26), line: { color: C.line, width: 0.75 } });
      if (i === 0 && jan > 0 && jan * mw < yw) return; // 途中の月から始まる最初の年が、次の年に重なる
      if (i > 0 && (y - months[0][0]) % every !== 0) return;
      T(s, withNen ? `${y}年` : `${y}`, { x: x + 0.05, y: y0, w: yw, h: 0.24, fontSize: FS.note, bold: true, color: C.soft, wrap: false });
    });
    const dots = [];
    rows.forEach((r, ri) => {
      const ay = axisY(ri);
      if (r.label) T(s, r.label, { x: M, y: ay - 0.25, w: lw, h: 0.5, fontSize: FS.head, bold: true, valign: "middle", lineSpacingMultiple: 1.1 });
      s.addShape(SH.LINE, { x: x0, y: ay, w: x1 - x0, h: 0, line: { color: C.gray, width: 1 } });
      // 段の割り当て：左から順に、次の3つを満たす一番上の段へ置く（名前は点の真下が既定。満たせないときだけ左右へずらす。名前は必ず自分の点の下にかかる）
      //   ・同じ段の名前と重ならない　・それより深い段へ下りる引出線を覆わない　・自分の引出線が浅い段の名前を通らない（通るなら、その名前を左へずらす）
      const ev = r.events.map((e, ei) => ({ ...e, ei, x: xOf(e.at), fz: e.self ? FS.head : FS.body })).sort((a, b) => a.x - b.x);
      const right = Array(levels).fill(-Infinity), placed = [];
      const reRight = () => { right.fill(-Infinity); placed.forEach(f => { right[f.L] = Math.max(right[f.L], f.l + f.w); }); };
      ev.forEach(e => {
        e.w = textW(e.name, e.fz) + 0.1;
        let ok = false;
        for (let L = 0; L < levels && !ok; L++) {
          const deeper = placed.filter(f => f.L > L).reduce((a, f) => Math.max(a, f.x + 0.08), -Infinity);
          const lo = Math.max(right[L] + 0.1, deeper, x0 - 0.05, e.x - e.w + 0.06), hi = Math.min(e.x - 0.06, x1 - e.w);
          if (lo > hi) continue;
          const cover = placed.filter(f => f.L < L && f.l + f.w > e.x - 0.08 && f.l < e.x + 0.08);
          if (cover.some(f => e.x - 0.08 - f.w < f.minL || e.x - 0.08 < f.x + 0.06)) continue;
          cover.forEach(f => { f.l = e.x - 0.08 - f.w; });
          e.L = L; e.minL = lo; e.l = Math.min(Math.max(e.x - e.w / 2, lo), hi); placed.push(e); reRight(); ok = true;
        }
        if (!ok) {
          const L = right.indexOf(Math.min(...right));
          e.L = L; e.minL = -Infinity; e.l = Math.min(Math.max(e.x - e.w / 2, x0 - 0.05), x1 - e.w); placed.push(e); reRight();
          warn(`時期の点の図：名前か引出線が重なる（${e.name}）。段を増やすか名前を縮める`);
        }
      });
      placed.forEach(e => {
        const ly = ay + 0.16 + e.L * step;
        if (e.L > 0) s.addShape(SH.LINE, { x: e.x, y: ay + 0.08, w: 0, h: ly - ay - 0.06, line: { color: C.gray, width: 0.75 } });
        T(s, e.name, { x: e.l, y: ly, w: e.w, h: e.self ? 0.3 : 0.26, fontSize: e.fz, bold: !!e.self, color: e.self ? C.ink : C.soft, align: "center", valign: "top", fill: { color: C.white }, wrap: false });
      });
      placed.forEach(e => { const dd = e.self ? 0.2 : 0.13; s.addShape(SH.OVAL, { x: e.x - dd / 2, y: ay - dd / 2, w: dd, h: dd, fill: { color: e.self ? C.emph : C.faint }, line: { color: C.white, width: 1 } }); });
      dots[ri] = []; placed.forEach(e => { dots[ri][e.ei] = [e.x, ay]; });
    });
    return { xOf, axisY, dot: (ri, ei) => dots[ri][ei] };
  };

  // 状態を描き分けるマップ：区分ごとの枠（tint の列）に項目の白箱を縦に並べ、状態を塗り・枠・文字の描き方で描き分ける」「整理済みの能力・要件」、p41）
  //   状態を3つ以上分けるときも塗りの色を増やさない。凡例は右上に細い枠の見本と名前だけ（図の下に注記を足さない）。付箋（ブレスト）には使わない
  //   o.groups: [{ name, items: [文字列 | { text, state }] }]（文字列は最初の状態）
  //   o.states: [{ key, label, style }]（style: "frame"＝白＋ink 1.25 枠＋太字、"dashed"＝塗りなし＋灰の破線＋灰字、"plain"＝白・枠なし・灰字、"fill"＝濃い塗り＋太字）
  //     既定は new（今回身についた・frame）／tried（経験した・dashed）／base（以前から・plain）
  //   o.x, o.w, o.y0（列の上端 2.6）, o.bottom, o.itemH（0.5）, o.itemGap（0.08）, o.colGap（0.12）, o.legend（false で凡例を置かない）
  //   戻り値: { col(i) → {x, y, w, h}, item(gi, ii) → {x, y, w, h}, legendY }
  const stateMap = (s, o) => {
    const G = o.groups, n = G.length, x0 = o.x ?? M, wAll = o.w ?? (W - M - x0), cg = o.colGap ?? GAP.std;
    const cw = (wAll - cg * (n - 1)) / n, colX = i => x0 + i * (cw + cg);
    const states = o.states || [{ key: "new", label: "今回身についた", style: "frame" }, { key: "tried", label: "経験した（再現性はまだない）", style: "dashed" }, { key: "base", label: "以前から", style: "plain" }];
    const stOf = k => states.find(t => t.key === k) || (warn(`状態のマップ：状態 ${k} が states に無い`), states[0]);
    const y0 = o.y0 ?? 2.6, y1 = o.bottom ?? BOTTOM, ih = o.itemH ?? 0.5, ig = o.itemGap ?? 0.08, pad = 0.08, inset = 0.05, fz = o.fs ?? FS.body;
    const look = style => style === "frame" ? { fill: C.white, line: { color: C.emph, width: 1.25 }, color: C.ink, bold: true }
      : style === "dashed" ? { fill: null, line: { color: C.gray, width: 1, dashType: "dash" }, color: C.gray, bold: false }
      : style === "fill" ? { fill: TINT_STRONG, line: null, color: C.ink, bold: true }
      : { fill: C.white, line: null, color: C.gray, bold: false };
    let legendY = null;
    if (o.legend !== false) {
      legendY = y0 - 0.38; let lx = x0 + wAll;
      [...states].reverse().forEach(t => {
        const tw = textW(t.label, FS.note) + 0.06, lk = look(t.style);
        lx -= tw; T(s, t.label, { x: lx, y: legendY, w: tw, h: 0.24, fontSize: FS.note, color: C.soft, valign: "middle", wrap: false });
        lx -= 0.38; rect(s, lx, legendY + 0.03, 0.3, 0.18, lk.fill === C.white || !lk.fill ? null : lk.fill, lk.line || { color: C.line, width: 0.75 });
        lx -= 0.3;
      });
    }
    const items = [];
    G.forEach((g, gi) => {
      const x = colX(gi);
      rect(s, x, y0, cw, y1 - y0, TINT);
      T(s, g.name, { x: x + pad, y: y0 + 0.1, w: cw - 2 * pad, h: 0.3, fontSize: FS.block, bold: true, valign: "middle", wrap: false });
      if (textW(g.name, FS.block) > cw - 2 * pad) warn(`状態のマップの区分名が列に収まらない（${g.name}）`);
      items[gi] = [];
      g.items.forEach((it, ii) => {
        const c = typeof it === "string" ? { text: it, state: states[0].key } : it, lk = look(stOf(c.state).style);
        const r = { x: x + pad, y: y0 + 0.5 + ii * (ih + ig), w: cw - 2 * pad, h: ih };
        rect(s, r.x, r.y, r.w, r.h, lk.fill, lk.line);
        T(s, c.text, { x: r.x + inset, y: r.y, w: r.w - 2 * inset, h: r.h, fontSize: fz, bold: lk.bold, color: lk.color, valign: "middle", lineSpacingMultiple: 1.05 });
        checkWrap(c.text, fz, r.w - 2 * inset);
        if (lineCount(c.text, fz, r.w - 2 * inset) * lineH(fz, 1.05) > ih - 0.04) warn(`状態のマップの項目があふれる（${c.text}）`);
        if (r.y + r.h > y1 - 0.06) warn(`状態のマップの「${g.name}」の項目が列の下端を超える`);
        items[gi].push(r);
      });
    });
    return { col: i => ({ x: colX(i), y: y0, w: cw, h: y1 - y0 }), item: (gi, ii) => items[gi][ii], legendY };
  };

  // 決裁の依頼：会議で決めてほしい事項を1枚に並べ、番号・決める内容・選択肢や条件・根拠のページ・決めた後の流れを、左から1行で読ませる
  //   頷くだけで済むように、選択肢には推奨を1つずつ決めて示す（濃い塗り＋ink 枠＋「推奨」。同じ意味の強調が行ごとに並ぶのは可）。詳しい根拠は書かず、ページ（P.n）で示す
  //   o.items: [{ head, body?, options?: [{ text, rec? }], cond?, basis?, refKeys?: [ページの名前], when?, after? }]
  //     cond＝条件の1行（選択肢の下）、basis＝根拠の名前（「試行の結果」）、refKeys＝根拠のページ（mark / base の key。P.n は最後に自動で入る）、when＝決めた後の最初の時期、after＝その後の流れ
  //   o.heads（列見出し。既定 ["決めていただくこと", "選択肢・条件", "根拠", "決めた後"]）, o.colW（4列の幅。既定 [3.3, 4.4, 1.6, 残り]）
  //   o.y0（行の上端 2.65）, o.rowGap（行の間 0.14）, o.rowH（省くと本文の下端までを等分。最大 1.3）
  //   戻り値: { rowY, rowH, colX, colW }
  const decisionPage = (s, o) => {
    warn("決裁の依頼（decisionPage）は、文のマス目。決裁のページは、根拠の図（計画・費用）を echo で背面に敷き、前面の messageBox で「決めていただくこと」を1つ言い、決める額を focusValue で大きく書く（references/components.md「決裁のページ」）");
    const items = o.items, n = items.length, heads = o.heads || ["決めていただくこと", "選択肢・条件", "根拠", "決めた後"];
    const nd = 0.36, x0 = M + nd + 0.2, wAll = W - M - x0, g = GAP.dense, rg = o.rowGap ?? 0.14;
    const cw0 = o.colW || [3.3, 4.4, 1.6]; const colW = cw0.length === 4 ? cw0 : [...cw0, wAll - cw0.reduce((a, b) => a + b, 0) - g * 3];
    const colX = colW.map((_, i) => x0 + colW.slice(0, i).reduce((a, b) => a + b, 0) + g * i);
    const y0 = o.y0 ?? 2.65, rowH = o.rowH ?? Math.min(1.3, (BOTTOM - y0 - rg * (n - 1)) / n), pad = PAD.x;
    heads.forEach((h, j) => T(s, h, { x: colX[j], y: y0 - 0.38, w: colW[j], h: 0.3, fontSize: FS.block, bold: true, color: C.gray, valign: "bottom", wrap: false }));
    const rowY = [];
    items.forEach((it, i) => {
      const y = y0 + i * (rowH + rg); rowY.push(y);
      numDot(s, i + 1, M, y + rowH / 2 - nd / 2, nd);
      colW.forEach((w, j) => rect(s, colX[j], y, w, rowH, TINT));
      // 決める内容
      const w0 = colW[0] - 2 * pad, ro = { headFs: FS.lead };
      T(s, boxRuns(it.head, it.body, ro), { x: colX[0] + pad, y, w: w0, h: rowH, valign: "middle", lineSpacingMultiple: 1.2 });
      checkBox(it.head, it.body, w0, rowH, ro, "決裁の依頼の「決めていただくこと」");
      // 選択肢：白い札を横に並べ、推奨の1つだけ濃い塗り＋ink 枠＋「推奨」。条件はその下に1行
      const ops = it.options || [], x1 = colX[1] + pad, w1 = colW[1] - 2 * pad, condH = it.cond ? 0.3 : 0;
      if (ops.filter(p => p.rec).length > 1) warn(`決裁の依頼：推奨（rec）が2つある（${it.head}）`);
      if (ops.length) {
        const og = 0.1, ow = (w1 - og * (ops.length - 1)) / ops.length, oh = Math.min(0.62, rowH - 0.24 - condH), oy = y + (rowH - oh - condH) / 2;
        ops.forEach((p, k) => {
          const ox = x1 + k * (ow + og);
          rect(s, ox, oy, ow, oh, p.rec ? TINT_STRONG : C.white, p.rec ? { color: C.emph, width: 1.5 } : null);
          const runs = p.rec ? [{ text: "推奨", options: { fontSize: FS.note, bold: true, color: C.ink, breakLine: true } }, { text: p.text, options: { fontSize: FS.block, bold: true, color: C.ink } }] : [{ text: p.text, options: { fontSize: FS.block, color: C.soft } }];
          T(s, runs, { x: ox + 0.1, y: oy, w: ow - 0.2, h: oh, align: "center", valign: "middle", lineSpacingMultiple: 1.1 });
          checkWrap(p.text, FS.block, ow - 0.2);
          if (lineCount(p.text, FS.block, ow - 0.2) * lineH(FS.block, 1.1) + (p.rec ? 0.2 : 0) > oh) warn(`決裁の依頼の選択肢があふれる（${p.text}）`);
        });
        if (it.cond) { T(s, it.cond, { x: x1, y: oy + oh + 0.04, w: w1, h: condH, fontSize: FS.body, color: C.soft, valign: "middle", wrap: false }); if (textW(it.cond, FS.body) > w1) warn(`決裁の依頼の条件が1行に収まらない（${it.cond}）`); }
      } else if (it.cond) { T(s, it.cond, { x: x1, y, w: w1, h: rowH, fontSize: FS.body, color: C.soft, valign: "middle" }); checkWrap(it.cond, FS.body, w1); }
      // 根拠：名前と P.n（ページ番号は全ページを作った後に入れる）
      const x2 = colX[2] + pad, w2 = colW[2] - 2 * pad;
      later(() => {
        const runs = [];
        if (it.basis) runs.push({ text: it.basis, options: { fontSize: FS.body, color: C.soft, breakLine: !!(it.refKeys && it.refKeys.length) } });
        if (it.refKeys && it.refKeys.length) runs.push({ text: it.refKeys.map(k => `P.${PAGE[k] ?? "?"}`).join("・"), options: { fontSize: FS.block, bold: true, color: C.ink } });
        if (runs.length) T(s, runs, { x: x2, y, w: w2, h: rowH, valign: "middle", lineSpacingMultiple: 1.25 });
      });
      if (it.basis) checkWrap(it.basis, FS.body, w2);
      // 決めた後：最初の時期（bold）＋その後の流れ
      const x3 = colX[3] + pad, w3 = colW[3] - 2 * pad;
      if (it.when) {
        T(s, boxRuns(it.when, it.after, { headFs: FS.block }), { x: x3, y, w: w3, h: rowH, valign: "middle", lineSpacingMultiple: 1.2 });
        checkBox(it.when, it.after, w3, rowH, { headFs: FS.block }, "決裁の依頼の「決めた後」");
      } else if (it.after) { T(s, it.after, { x: x3, y, w: w3, h: rowH, fontSize: FS.body, color: C.soft, valign: "middle", lineSpacingMultiple: 1.2 }); checkWrap(it.after, FS.body, w3); }
    });
    const bottom = y0 + n * rowH + (n - 1) * rg;
    if (bottom > BOTTOM + 0.05) warn(`決裁の依頼の下端 ${bottom.toFixed(2)} が本文の下端を超える（事項を減らすか2枚に分ける）`);
    return { rowY, rowH, colX, colW };
  };

  // 折れ線で囲んだ面（custGeom）：pts はスライドの座標 [[x, y], ...]。外接の矩形を図形の枠にし、点はその中の相対位置で渡す
  //   根拠: 循環図のドーナツの区画（先端が矢じり・後ろが切り欠き）と、切り欠きの深さを変えた矢羽は、既製の図形（blockArc・chevron）では形を決められないため
  const flowPoly = (s, pts, fill, line) => {
    const xs = pts.map(p => p[0]), ys = pts.map(p => p[1]), x = Math.min(...xs), y = Math.min(...ys);
    s.addShape(SH.CUSTOM_GEOMETRY, {
      x, y, w: Math.max(Math.max(...xs) - x, 0.001), h: Math.max(Math.max(...ys) - y, 0.001),
      fill: fill ? { color: fill } : { type: "none" }, line: line || { type: "none" },
      points: [...pts.map(([px, py]) => ({ x: px - x, y: py - y })), { close: true }],
    });
  };

  // 段階の矢羽（切り欠き 0.12 の浅い矢羽。先頭は切り欠きなしの五角形）。工程の流れ図の工程名と、スイムレーンの段階の見出しで共通に使う
  //   根拠: 既製の CHEVRON は切り欠きが高さの半分で深く、矢羽の間の白い溝が下のブロックの溝（0.06〜0.08）の3倍に開いた
  const flowChevron = (s, x, y, w, h, first, fill, d = 0.12) => {
    const pts = [[x, y], [x + w - d, y], [x + w, y + h / 2], [x + w - d, y + h], [x, y + h]];
    if (!first) pts.push([x + d, y + h / 2]);
    flowPoly(s, pts, fill);
  };

  // 詳細なスイムレーン：担当（5〜6レーン）×時間の列（10前後）の格子に工程を置き、受け渡しを鉤に折った矢印でつなぐ。判断はひし形、帳票は書類の形、システムへの登録は円柱と破線
  //4レーン×6工程の箱を線でつないだだけでは、どこで誰が判断し、何を登録し、どこで待つかが読めなかった。
  //         戦略コンサルの業務分析のスイムレーンは、判断の分岐・帳票・システム・待ち時間を1枚に描く。線は鉤に折り（4.13）、交差を避ける（交差・箱への重なりは console.warn で知らせる）
  //   o.lanes: [名前 | { name, sub? }]（上から。sub は人数などの小さな灰の語）
  //   o.nodes: [{ id, lane（番号か名前）, col, name, kind?, sub?, problem?, strong? }]
  //     kind: "task"（既定。白の箱）| "decision"（判断のひし形）| "doc"（帳票。下が波の書類）| "system"（システム。円柱）| "end"（始まり・終わり。角の丸い箱）
  //     problem＝詰まる工程（赤 2pt 枠＋赤字。その列の所要も赤）、strong＝問題ではない主役（濃い塗り＋ink 枠）。1ページでどちらか1種類だけ
  //   o.edges: [{ from, to, label?, back?, dashed?, exit?: "r"|"t"|"b", via?: "src"|"dst" }]
  //     既定は紺の線。to の列が from より前（または back）なら赤の戻り（差し戻し）で、先頭のレーンへ戻るときはレーンの上の白い帯を通る。dashed＝システムへの登録・帳票の受け渡し（灰の破線）
  //     レーンをまたぐ線は、from の右の溝（via "src"）か to の左の溝（via "dst"）で縦に折る。既定は箱に重ならない方。exit で判断の分岐の出口（上・下）を選ぶ
  //     label は、戻りの線なら通り道の中ほどに白地の赤字で、上・下へ出る分岐なら出口の横に置く（右へ出る線は溝が狭く置けないので、条件はもう一方の分岐か戻りの線に書く）
  //   o.phases: [{ name, from, to }]（列の範囲ごとの段階。上に浅い矢羽で示す）
  //   o.times: [文字列 | null]（列ごとの所要・待ち。下の行に列にそろえて置く。problem の工程がある列は赤字）, o.timeLabel（"所要"）
  //   o.x, o.w, o.y0（BODY）, o.bottom（BOTTOM）, o.labelW（レーンの名前の欄。既定は中身から、1.3 以上）, o.cols（列の数。既定は工程から）, o.nodeW, o.nodeH, o.fs（工程名の文字 10。10〜13列の密な図のため本文より 0.5pt 小さい。判断のひし形の中は 9.5）
  //   戻り値: { node(id) → {x, y, w, h, cx, cy}, laneY(i), laneH, colCx(c), pitch, bottom }
  const swimlane = (s, o) => { denseFig = true; try { return swimlane0(s, o); } finally { denseFig = false; } };
  const swimlane0 = (s, o) => {
    const lanes = o.lanes.map(l => typeof l === "string" ? { name: l } : l), L = lanes.length, nodes = o.nodes, edges = o.edges || [];
    checkOneEm(nodes, "スイムレーン");
    const laneIdx = v => typeof v === "number" ? v : lanes.findIndex(l => l.name === v);
    const byId = {}; nodes.forEach(nd => { byId[nd.id] = nd; });
    const nc = o.cols ?? Math.max(...nodes.map(nd => nd.col)) + 1;
    const lw = Math.max(o.labelW ?? 1.3, labelCol("スイムレーン", lanes.flatMap(l => [[String(l.name), FS.block], ...(l.sub ? [[String(l.sub), FS.note]] : [])]), o.labelW, FS.block, 2.2));
    const x0 = o.x ?? M, wAll = o.w ?? (W - M - x0), bx = x0 + lw + 0.1, bwAll = x0 + wAll - bx;
    const inPad = 0.06, pitch = (bwAll - 2 * inPad) / nc, colCx = c => bx + inPad + (c + 0.5) * pitch;
    const y0 = o.y0 ?? BODY, bottom = o.bottom ?? BOTTOM, fs = o.fs ?? 16; // 工程の名前は 16pt（5〜6レーン×10前後の密な図。本文 16pt では2行が箱に入らない）
    const isBack = e => byId[e.from] && byId[e.to] && (e.back || byId[e.to].col < byId[e.from].col);
    const topBacks = edges.filter(e => isBack(e) && laneIdx(byId[e.to].lane) === 0);
    const phH = o.phases ? 0.4 : 0, chanH = topBacks.length ? 0.34 + (topBacks.length - 1) * 0.28 : (o.phases ? 0.08 : 0);
    const hasTime = !!o.times, tH = hasTime ? 0.34 : 0, tG = hasTime ? 0.08 : 0, lg = GAP.dense;
    const ly0 = y0 + phH + chanH, ly1 = bottom - tH - tG;
    const laneH = (ly1 - ly0 - lg * (L - 1)) / L, laneY = i => ly0 + i * (laneH + lg), laneMid = i => laneY(i) + laneH / 2;
    const nw = o.nodeW ?? Math.min(1.4, pitch - 0.12), nh = o.nodeH ?? Math.min(0.9, laneH - 0.04);
    // レーン：名前は塗らず bold ink、帯は tint、その上の工程は白
    lanes.forEach((l, i) => {
      if (l.sub) warnSub("スイムレーンのレーン");
      T(s, [{ text: l.name, options: { fontSize: FS.block, bold: true, color: C.ink, breakLine: !!l.sub } }, ...(l.sub ? [{ text: l.sub, options: { fontSize: FS.note, color: C.gray } }] : [])],
        { x: x0, y: laneY(i), w: lw, h: laneH, valign: "middle", lineSpacingMultiple: 1.1 });
      checkWrap(l.name, FS.block, lw);
      // 白地に線でレーンを分ける（帯を塗ると、白い箱との差が無くなって全体がぼやけた。2026-11-24）
      hline(s, x0, laneY(i) - lg / 2, bwAll + lw + 0.1, { color: C.line, width: 0.75 });
      if (i === L - 1) hline(s, x0, laneY(i) + laneH + lg / 2, bwAll + lw + 0.1, { color: C.line, width: 0.75 });
    });
    // 段階：列の範囲の上に浅い矢羽
    (o.phases || []).forEach((p, k) => {
      const xa = colCx(p.from) - pitch / 2 + 0.02, xb = colCx(p.to) + pitch / 2 - 0.02;
      flowChevron(s, xa, y0, xb - xa, phH, k === 0, C.bar, 0.12);
      T(s, p.name, { x: xa + 0.14, y: y0, w: xb - xa - 0.26, h: phH, fontSize: FS.block, bold: true, align: "center", valign: "middle", wrap: false });
      if (textW(p.name, FS.block) > xb - xa - 0.26) warn(`スイムレーンの段階の名前が矢羽に入らない（${p.name}）`);
    });
    // 工程の位置
    const R = {};
    nodes.forEach(nd => {
      const li = laneIdx(nd.lane), k = nd.kind || "task";
      if (li < 0 || li >= L) { warn(`スイムレーン：レーンが合わない工程（${nd.name}）`); return; }
      let w = nd.w ?? nw, h = nh;
      if (k === "decision") { w = pitch - 0.02; h = laneH - 0.04; }
      if (k === "system") h = laneH - 0.06;
      const cx = colCx(nd.col), cy = laneMid(li);
      R[nd.id] = { x: cx - w / 2, y: cy - h / 2, w, h, cx, cy, lane: li, col: nd.col, nd };
    });
    const anchor = (r, side) => side === "r" ? [r.x + r.w, r.cy] : side === "l" ? [r.x, r.cy] : side === "t" ? [r.cx, r.y] : [r.cx, r.y + r.h];
    // 線と箱の重なり・線どうしの交差の判定（線は水平と垂直だけ）
    const hitRect = (p, q, r) => Math.min(p[0], q[0]) < r.x + r.w - 0.02 && Math.max(p[0], q[0]) > r.x + 0.02 && Math.min(p[1], q[1]) < r.y + r.h - 0.02 && Math.max(p[1], q[1]) > r.y + 0.02;
    const hitAny = (pts, skip) => { for (let k = 1; k < pts.length; k++) for (const id in R) if (!skip.includes(id) && hitRect(pts[k - 1], pts[k], R[id])) return id; return null; };
    let topK = 0;
    const routes = edges.map(e => {
      const a = R[e.from], b = R[e.to];
      if (!a || !b) { warn(`スイムレーン：線の端の工程が無い（${e.from} → ${e.to}）`); return null; }
      const skip = [e.from, e.to];
      if (isBack(e)) {
        const up = b.lane <= a.lane, start = anchor(a, e.exit ?? (up ? "t" : "b")), end = anchor(b, up ? "t" : "b");
        const cy = up && b.lane === 0 ? ly0 - 0.2 - (topK++) * 0.28 : up ? laneY(b.lane) - lg / 2 : laneY(b.lane) + laneH + lg / 2;
        return { e, kind: "back", pts: [start, [start[0], cy], [end[0], cy], end] };
      }
      const kind = e.dashed ? "data" : "flow";
      if (a.col === b.col) return { e, kind, pts: b.cy > a.cy ? [anchor(a, "b"), anchor(b, "t")] : [anchor(a, "t"), anchor(b, "b")] };
      const ex = e.exit ?? "r";
      if (ex !== "r") { const st0 = anchor(a, ex); return { e, kind, pts: [st0, [st0[0], b.cy], anchor(b, "l")] }; }
      const st0 = anchor(a, "r"), en = anchor(b, "l");
      if (Math.abs(a.cy - b.cy) < 0.01) return { e, kind, pts: [st0, en] };
      const cand = v => { const xm = colCx(v === "dst" ? b.col - 1 : a.col) + pitch / 2; return [st0, [xm, a.cy], [xm, b.cy], en]; };
      const order = e.via ? [e.via] : ["src", "dst"];
      const pick = order.find(v => !hitAny(cand(v), skip)) || order[0];
      return { e, kind, pts: cand(pick) };
    }).filter(Boolean);
    // 同じ溝を縦に通る線は少しずらす（重ねない）
    const gapUse = {};
    routes.forEach(rt => {
      if (rt.kind === "back" || rt.pts.length !== 4 || Math.abs(rt.pts[1][0] - rt.pts[2][0]) > 0.001) return;
      const key = rt.pts[1][0].toFixed(3), k = gapUse[key] = (gapUse[key] ?? -1) + 1, d = [0, 0.05, -0.05][k % 3];
      rt.pts[1][0] += d; rt.pts[2][0] += d;
    });
    // 点検：線が他の工程に重なる、線どうしが交わる
    routes.forEach(rt => { const id = hitAny(rt.pts, [rt.e.from, rt.e.to]); if (id) warn(`スイムレーン：線（${rt.e.from} → ${rt.e.to}）が「${R[id].nd.name}」に重なる（via か exit を変える、列をずらす）`); });
    const segsOf = rt => rt.pts.slice(1).map((q, k) => [rt.pts[k], q]);
    for (let i = 0; i < routes.length; i++) for (let j = i + 1; j < routes.length; j++) {
      const A = routes[i], B = routes[j];
      const cross = segsOf(A).some(([p, q]) => segsOf(B).some(([u, v]) => {
        const h1 = Math.abs(p[1] - q[1]) < 0.001, h2 = Math.abs(u[1] - v[1]) < 0.001;
        if (h1 === h2) return false;
        const [H, V] = h1 ? [[p, q], [u, v]] : [[u, v], [p, q]];
        const x = V[0][0], y = H[0][1];
        return x > Math.min(H[0][0], H[1][0]) + 0.01 && x < Math.max(H[0][0], H[1][0]) - 0.01 && y > Math.min(V[0][1], V[1][1]) + 0.01 && y < Math.max(V[0][1], V[1][1]) - 0.01;
      }));
      if (cross) warn(`スイムレーン：線が交差する（${A.e.from} → ${A.e.to} と ${B.e.from} → ${B.e.to}）`);
    }
    // 線：進む＝紺、戻り＝赤、登録・帳票＝灰の破線。工程より先に描く
    const style = { flow: { color: C.soft, width: 1.75 }, back: { color: C.accent, width: 2 }, data: { color: C.gray, width: 1.25, dashType: "dash" } };
    routes.forEach(rt => polyArrow(s, rt.pts, style[rt.kind]));
    // 工程
    Object.values(R).forEach(r => {
      const nd = r.nd, k = nd.kind || "task", red = !!nd.problem;
      const fill = nd.strong ? C.emph : emFill(nd, k === "system" || k === "decision" ? C.white : C.tint), line = emLine(nd) || { color: k === "system" || k === "decision" ? C.soft : C.bar, width: 1 };
      const shape = { decision: SH.DIAMOND, doc: SH.FLOWCHART_DOCUMENT, system: SH.CAN, end: SH.FLOWCHART_TERMINATOR }[k] || SH.RECTANGLE;
      s.addShape(shape, { x: r.x, y: r.y, w: r.w, h: r.h, fill: { color: fill }, line: line || { type: "none" } });
      const nfs = fs, iw = k === "decision" ? r.w * 0.64 : r.w - 0.08, ty = k === "system" ? r.y + r.h * 0.16 : r.y, thh = k === "system" ? r.h * 0.84 : k === "doc" ? r.h * 0.86 : r.h;
      const runs = [{ text: nd.name, options: { fontSize: nfs, bold: true, color: red ? C.accent : nd.strong ? C.white : C.ink, breakLine: !!nd.sub } }];
      if (nd.sub && (warnSub("図の要素"), true)) runs.push({ text: nd.sub, options: { fontSize: FS.note, color: C.soft } });
      T(s, runs, { x: r.cx - iw / 2, y: ty, w: iw, h: thh, align: "center", valign: "middle", lineSpacingMultiple: 1.0 });
      checkWrap(nd.name, nfs, iw);
      const needH = lineCount(nd.name, nfs, iw) * lineH(nfs, 1.0) + (nd.sub ? lineH(FS.note, 1.05) : 0);
      if (needH > thh * (k === "decision" ? 0.85 : 1) + 0.02) warn(`スイムレーンの工程の文字があふれる（${String(nd.name).replace(/\n/g, "")}。この列の幅では1行が全角${Math.floor(iw * 72 / nfs)}字まで。言葉を縮めるか、列を減らす）`);
    });
    // 戻りの注記：通り道の横の区間の中ほどに白地で
    routes.filter(rt => rt.kind === "back" && rt.e.label).forEach(rt => {
      const [p, q] = [rt.pts[1], rt.pts[2]], lw2 = textW(rt.e.label, FS.note) + 0.2;
      if (lw2 > Math.abs(q[0] - p[0]) - 0.15) warn(`スイムレーンの戻りの注記が線より長い（${rt.e.label}）`);
      T(s, rt.e.label, { x: (p[0] + q[0]) / 2 - lw2 / 2, y: p[1] - 0.13, w: lw2, h: 0.26, fontSize: FS.note, bold: true, color: C.accent, align: "center", valign: "middle", fill: { color: C.white }, wrap: false });
    });
    // 判断の分岐の注記：上・下へ出る線は、出口のすぐ横（線の右）に置く。右へ出る線は溝が狭く置けないので、もう一方の分岐か戻りの線に条件を書く
    routes.filter(rt => rt.kind !== "back" && rt.e.label).forEach(rt => {
      const [p, q] = rt.pts, vert = Math.abs(p[0] - q[0]) < 0.001;
      if (!vert) { warn(`スイムレーン：右へ出る線のラベル（${rt.e.label}）は置けない。条件はもう一方の分岐か戻りの線に書く`); return; }
      const lw2 = textW(rt.e.label, FS.note) + 0.06, down = q[1] > p[1], lr = { x: p[0] + 0.05, y: down ? p[1] + 0.01 : p[1] - 0.23, w: lw2, h: 0.22 };
      const hit = Object.values(R).find(r => lr.x < r.x + r.w && lr.x + lr.w > r.x && lr.y < r.y + r.h && lr.y + lr.h > r.y);
      if (hit) warn(`スイムレーン：分岐の注記（${rt.e.label}）が「${hit.nd.name}」に重なる`);
      T(s, rt.e.label, { ...lr, fontSize: FS.note, color: C.soft, valign: "middle", wrap: false });
    });
    // 所要：列にそろえて下の行に。詰まる工程の列だけ赤字
    if (hasTime) {
      const yT = ly1 + tG;
      T(s, o.timeLabel ?? "所要", { x: x0, y: yT, w: lw, h: tH, fontSize: FS.block, bold: true, valign: "middle", wrap: false });
      o.times.forEach((t, c) => {
        if (t == null) return;
        const red = nodes.some(nd => nd.col === c && nd.problem);
        T(s, t, { x: colCx(c) - pitch / 2, y: yT, w: pitch, h: tH, fontSize: red ? FS.block : FS.body, bold: red, color: red ? C.accent : C.soft, align: "center", valign: "middle", wrap: false });
        if (textW(t, red ? FS.block : FS.body) > pitch) warn(`スイムレーンの所要が列の幅に入らない（${t}）`);
      });
    }
    return { node: id => R[id], laneY, laneH, colCx, pitch, bottom: hasTime ? ly1 + tG + tH : ly1 };
  };

  // 循環図：ドーナツ状の環を工程の数（3〜6）に分け、各区画の先端を矢じり・後ろを切り欠きにして、時計回りに回る向きを形で示す。工程の名前と説明は環の外側に引出線で、回すもの・回数・主張は環の中心に置く
  //
  //         箱を横に並べて赤い戻り矢印でつないだ形は、繰り返しが「一方通行＋例外の戻り」に見えた。円を区画に分ければ、どこからでも次へ同じ向きに進むことが形で分かる。下の別のメッセージ欄は作らない
  //   o.steps: [{ name, body?, meta?: [[ラベル, 値], ...], problem?, strong? }]（3〜6。上の始点から時計回り）
  //     problem＝繰り返しのたびに詰まる区画（白地＋赤 2pt 枠、番号と名前を赤）、strong＝問題ではない主役の区画（濃い塗り＋ink 枠）。1ページでどちらか1種類だけ
  //   o.center: { label?, value?, note? }（label＝回すもの FS.head（18pt）bold、value＝回数 22pt bold、note＝主張・前提 FS.body（16pt）。環の内側に入る長さにする）
  //   o.cx, o.cy, o.R（外の半径。既定は本文の高さの半分）, o.thick（環の太さ 0.62）, o.start（最初の区画の始まりの角度。既定は名前が環の真上・真下に来ない角度）, o.y0, o.bottom
  //   o.only: [工程の番号]（その区画と名前だけを描く。echo の keep で使う）
  //   戻り値: { seg(i) → { a0, a1, mid, at: [x, y]（区画の中心） }, label(i) → {x, y, w, h}, anchor(i) → [x, y], center → {x, y, w, h}, cx, cy, R, r }
  const cycleDiagram = (s, o) => {
    const st = o.steps, n = st.length;
    if (n < 3 || n > 6) warn(`循環図：工程は3〜6つにする（${n}）`);
    checkOneEm(st, "循環図");
    const y0 = o.y0 ?? BODY, bottom = o.bottom ?? BOTTOM;
    const R = o.R ?? Math.min(2.1, (bottom - y0) / 2 - 0.04), thick = o.thick ?? 0.62, r = R - thick, rm = R - thick / 2;
    const cx = o.cx ?? W / 2, cy = o.cy ?? (y0 + bottom) / 2;
    const span = 360 / n, deg = v => v * 180 / Math.PI, gapD = deg(0.08 / rm), tipD = deg((o.tip ?? 0.26) / rm);
    const midOf = (off, i) => off + i * span + (span - gapD + tipD) / 2;
    // 始まりの角度：区画の中心が環の真上・真下（名前の置き場のない所）から最も離れる角度。同点なら真上に近い方
    const score = off => Math.min(...st.map((_, i) => { const m = ((midOf(off, i) % 180) + 180) % 180; return Math.min(m, 180 - m); }));
    let off = o.start;
    if (off == null) { off = 0; let best = -1; for (let t = -Math.floor(span / 2); t <= Math.floor(span / 2); t++) { const sc = score(t); if (sc > best + 0.5 || (Math.abs(sc - best) <= 0.5 && Math.abs(t) < Math.abs(off))) { best = Math.max(best, sc); off = t; } } }
    const P = (rho, a) => { const t = a * Math.PI / 180; return [cx + rho * Math.sin(t), cy - rho * Math.cos(t)]; };
    const arc = (rho, a, b) => { const k = Math.max(2, Math.ceil(Math.abs(b - a) / 3)); return Array.from({ length: k + 1 }, (_, q) => P(rho, a + (b - a) * q / k)); };
    const segs = st.map((_, i) => { const a0 = off + i * span, a1 = a0 + span - gapD, mid = midOf(off, i); return { a0, a1, mid, at: P(rm, mid) }; });
    const only = o.only ? new Set(o.only) : null, show = i => !only || only.has(i);
    // 区画：外の弧 → 先端の矢じり → 内の弧 → 後ろの切り欠き。切り欠きに前の区画の矢じりが入り、白い溝で回る向きを示す
    st.forEach((it, i) => {
      if (!show(i)) return;
      const g = segs[i];
      flowPoly(s, [...arc(R, g.a0, g.a1), P(rm, g.a1 + tipD), ...arc(r, g.a1, g.a0), P(rm, g.a0 + tipD)], emFill(it, o.fill ?? C.bar), emLine(it));
      T(s, String(i + 1), { x: g.at[0] - 0.25, y: g.at[1] - 0.18, w: 0.5, h: 0.36, fontFace: FNUM, fontSize: FS.head, bold: true, color: it.problem ? C.accent : C.ink, align: "center", valign: "middle", wrap: false });
    });
    // 名前と説明：区画の中心の角度で環の外へ引出線を出し、右半分は右の列、左半分は左の列に置く。重なるときは縦にずらす
    const gutter = 0.42, colR = cx + R + gutter, wR = W - M - colR, wL = cx - R - gutter - M;
    const headL = lineH(FS.head, 1.15);
    const items = st.map((it, i) => {
      const g = segs[i], side = Math.sin(g.mid * Math.PI / 180) >= 0 ? 1 : -1, maxW = side > 0 ? wR : wL;
      const meta = (it.meta || []).map(([k, v]) => `${k} ${v}`).join("　");
      const wOf = t => t ? Math.max(...String(t).split("\n").map(l => textW(l, l === t ? FS.body : FS.body))) : 0;
      const bw = Math.min(maxW, Math.max(textW(String(it.name).split("\n")[0], FS.head), meta ? textW(meta, FS.body) : 0, it.body ? wOf(it.body) : 0) + 0.06);
      const h = lineCount(it.name, FS.head, bw) * headL + (meta ? lineH(FS.body, 1.2) + 0.03 : 0) + (it.body ? 0.05 + lineCount(it.body, FS.body, bw) * lineH(FS.body, 1.2) : 0);
      const anc = P(R + 0.04, g.mid);
      return { i, it, side, bw, h, anc, top: anc[1] - headL / 2 };
    });
    [1, -1].forEach(sd => {
      const list = items.filter(t => t.side === sd).sort((a, b) => a.anc[1] - b.anc[1]);
      list.forEach((t, k) => { t.top = Math.max(t.top, y0, k ? list[k - 1].top + list[k - 1].h + 0.22 : -9); });
      for (let k = list.length - 1; k >= 0; k--) { const lim = k === list.length - 1 ? bottom : list[k + 1].top - 0.22; if (list[k].top + list[k].h > lim) list[k].top = lim - list[k].h; }
      if (list.length && list[0].top < y0 - 0.01) warn("循環図：片側の名前と説明が本文の高さに入らない（説明を縮める）");
    });
    const labels = [];
    items.forEach(t => {
      const x = t.side > 0 ? colR : cx - R - gutter - t.bw, hcy = t.top + headL / 2, ex = t.side > 0 ? x - 0.08 : x + t.bw + 0.08;
      labels[t.i] = { x, y: t.top, w: t.bw, h: t.h };
      // 名前を縦にずらしたときは、引出線の根元を区画の中で同じ高さまで滑らせ、線を水平に保つ
      if (Math.abs(hcy - t.anc[1]) >= 0.02) {
        const c = (cy - hcy) / (R + 0.04);
        if (Math.abs(c) <= 1) {
          let th = Math.acos(c) * 180 / Math.PI; if (t.side < 0) th = 360 - th;
          const g = segs[t.i], norm = a => ((a - g.a0) % 360 + 360) % 360, lo = 3, hi = g.a1 - g.a0 - 3;
          if (norm(th) >= lo && norm(th) <= hi) t.anc = P(R + 0.04, th);
        }
      }
      if (!show(t.i)) return;
      const ln = { color: C.gray, width: 0.75 };
      if (Math.abs(hcy - t.anc[1]) < 0.02) segLine(s, t.anc[0], t.anc[1], ex, t.anc[1], ln);
      else { const bxp = t.anc[0] + t.side * 0.22; segLine(s, t.anc[0], t.anc[1], bxp, hcy, ln); segLine(s, bxp, hcy, ex, hcy, ln); }
      const red = !!t.it.problem, meta = t.it.meta || [];
      const runs = [{ text: t.it.name, options: { fontSize: FS.head, bold: true, color: red ? C.accent : C.ink, breakLine: !!(meta.length || t.it.body) } }];
      meta.forEach(([k, v], q) => runs.push({ text: `${q ? "　" : ""}${k} `, options: { fontSize: FS.note, color: C.gray } }, { text: v, options: { fontSize: FS.body, color: C.soft, breakLine: q === meta.length - 1 && !!t.it.body } }));
      if (t.it.body) runs.push({ text: t.it.body, options: { fontSize: FS.body, color: C.soft } });
      T(s, runs, { x, y: t.top, w: t.bw + 0.04, h: t.h + 0.04, valign: "top", lineSpacingMultiple: 1.15 });
      checkWrap(t.it.name, FS.head, t.bw); if (t.it.body) checkWrap(t.it.body, FS.body, t.bw);
    });
    // 中心：回すもの・回数・主張
    const c0 = o.center || {}, ciw = r * 1.62, cih = r * 1.5, center = { x: cx - ciw / 2, y: cy - cih / 2, w: ciw, h: cih };
    if (!only && (c0.label || c0.value || c0.note)) {
      const runs = [];
      if (c0.label) runs.push({ text: c0.label, options: { fontSize: FS.head, bold: true, color: C.ink, breakLine: true } });
      if (c0.value) runs.push({ text: c0.value, options: { fontSize: FS_DISPLAY.num, bold: true, color: C.ink, breakLine: !!c0.note } });
      if (c0.note) runs.push({ text: c0.note, options: { fontSize: FS.body, color: C.soft } });
      T(s, runs, { ...center, align: "center", valign: "middle", lineSpacingMultiple: 1.2 });
      [[c0.label, FS.head], [c0.note, FS.body]].forEach(([t, f]) => { if (t && String(t).split("\n").some(l => textW(l, f) > ciw)) warn(`循環図の中心の文字が環の内側に入らない（${t}）`); });
    }
    return { seg: i => segs[i], label: i => labels[i], anchor: i => items[i].anc, center, cx, cy, R, r };
  };

  // 体制図（樹形）：責任者 → 部署 → チーム（チームリーダー）→ メンバー の3〜4階層を、分岐の線でつなぐ。部署とチームに人数、メンバーに役割と人数、兼務、社外の協力先を書く。主導する部署だけ強調
  //
  //         部署の箱に人のピクトを並べただけの形は、部署の中の指揮の筋（誰がどのチームを率い、何人がどの役割か）が見えなかった。戦略コンサルの体制図は、樹形で階層と人数・役割・兼務を1枚に書く
  //   線は所属の線だけ（灰 1pt の鉤の分岐）。社外の協力先は右端に、破線の枠と破線のつなぎで置く。ピクトは責任者とチームリーダーにだけ（大きな箱に人を並べない）
  //   o.head: { role, body?, tag?（既定 "責任者"） }
  //   o.staff: [{ role, body?, tag? }]（責任者の直属。責任者の右に横の線でつなぐ。事務局・PMO など）
  //   o.depts: [{ name, count?, body?, lead?, ext?, teams: [{ name, leader?, dual?, members: [{ role, n, dual? }] }] }]
  //     count＝部署の人数（"14名"。名前の下の行に役目 body と並べる）、lead＝主導する部署（濃い塗り＋ink 2pt 枠＋「主導」。1つだけ）、ext＝社外（右端にまとめる）
  //     teams の leader＝チームリーダーの役職（"課長" など。2〜4字。小さな人のピクトを添え、左上の凡例で「チームのリーダー」と一度だけ言う）、dual＝兼務。members の n＝人数、dual＝兼務（役割の後に「兼務」）
  //   o.x, o.w, o.y0（BODY）, o.bottom（BOTTOM）, o.headW（3.2）, o.teamGap（同じ部署のチームの間 0.1）, o.deptGap（部署の間 0.26）
  //   戻り値: { head, staff(k), dept(i), team(i, j), member(i, j, k) → {x, y, w, h} } … 重ね技で前面から指すときに使う
  const orgChart = (s, o) => {
    const D = o.depts, x0 = o.x ?? M, wAll = o.w ?? (W - M - x0), y0 = o.y0 ?? BODY, bottom = o.bottom ?? BOTTOM;
    if (D.filter(d => d.lead).length > 1) warn("体制図：主導する部署（lead）は1つだけにする");
    const firstExt = D.findIndex(d => d.ext), lastInt = firstExt < 0 ? D.length - 1 : firstExt - 1;
    if (firstExt >= 0 && D.slice(firstExt).some(d => !d.ext)) warn("体制図：社外の協力先は右端にまとめる");
    const teamsOf = d => d.teams && d.teams.length ? d.teams : [];
    const ig = o.teamGap ?? 0.1, dg = o.deptGap ?? 0.26, eg = firstExt > 0 ? 0.26 : 0;
    const nLeaf = D.reduce((a, d) => a + Math.max(1, teamsOf(d).length), 0);
    const lfW = (wAll - D.reduce((a, d) => a + (Math.max(1, teamsOf(d).length) - 1) * ig, 0) - dg * (D.length - 1) - eg) / nLeaf;
    const teamX = [], deptR = []; { let x = x0; D.forEach((d, i) => { if (i) x += dg + (i === firstExt ? eg : 0); const xs = []; for (let j = 0; j < Math.max(1, teamsOf(d).length); j++) { if (j) x += ig; xs.push(x); x += lfW; } teamX.push(xs); deptR.push({ x: xs[0], w: x - xs[0] }); }); }
    const pad = 0.08, ln = { color: C.gray, width: 1 }, dl = { color: C.gray, width: 1, dashType: "dash" };
    // 高さ：部署とチームの箱は文字の量でそろえ、残りをメンバーの行に使う
    // 部署の箱：1行目に名前、2行目に「主導」・人数・役目（幅の狭い部署でも名前を折らない）
    const deptSub = d => `${d.lead ? "主導　" : ""}${d.count ? d.count + "　" : ""}${d.body || ""}`.trim();
    const deptRuns = d => {
      const r = [{ text: d.name, options: { fontSize: FS.head, bold: true, color: d.ext ? C.soft : C.ink, breakLine: !!deptSub(d) } }];
      if (d.lead) r.push({ text: "主導　", options: { fontSize: FS.note, bold: true, color: C.ink } });
      if (d.count) r.push({ text: `${d.count}${d.body ? "　" : ""}`, options: { fontSize: FS.body, bold: true, color: d.ext ? C.soft : C.ink } });
      if (d.body) r.push({ text: d.body, options: { fontSize: FS.body, color: C.soft } });
      return r;
    };
    const dNeed = Math.max(...D.map((d, i) => lineCount(d.name, FS.head, deptR[i].w - 2 * pad) * lineH(FS.head, 1.15) + (deptSub(d) ? lineCount(deptSub(d), FS.body, deptR[i].w - 2 * pad) * lineH(FS.body, 1.15) : 0)));
    const allTeams = D.flatMap(d => teamsOf(d));
    const tNeed = Math.max(0.3, ...allTeams.map(t => lineCount(t.name, FS.block, lfW - 2 * pad) * lineH(FS.block, 1.1) + (t.leader ? lineH(FS.body, 1.2) + 0.04 : 0)));
    // 縦の配分：最小の間と行の高さで組み、余りを階層の間とメンバーの行に配る（上に詰めて下を空けない）
    const hH = o.headH ?? 0.8, dH = dNeed + 0.2, tH = tNeed + 0.18;
    const maxRows = Math.max(1, ...allTeams.map(t => (t.members || []).length)), row0 = 0.3;
    const spare = Math.max(0, bottom - (y0 + hH + 0.34 + dH + 0.3 + tH + 0.12 + maxRows * row0));
    const rowH = Math.min(0.42, row0 + spare * 0.5 / maxRows), rest = spare - (rowH - row0) * maxRows;
    const g1 = 0.34 + rest / 2, g2 = 0.3 + rest / 2;
    const dy = y0 + hH + g1, ty = dy + dH + g2, my = ty + tH + 0.12;
    if (my + maxRows * 0.26 > bottom + 0.02) warn("体制図：メンバーの行が本文の下端に入らない（行を減らすか、チームの名前を短くする）");
    // 責任者：社内の部署の中央の上。人のピクト＋役割
    const midX = (deptR[0].x + deptR[lastInt].x + deptR[lastInt].w) / 2, hw = o.headW ?? 3.2, head = { x: midX - hw / 2, y: y0, w: hw, h: hH };
    const H0 = o.head || {};
    // 凡例：チームの箱の小さなピクトはチームのリーダー（各チームで「リーダー」と書かず、凡例で一度だけ言う）
    if (allTeams.some(t => t.leader) && head.x > x0 + 2.2) {
      person(s, x0 + 0.07, y0 + 0.06, 0.07, C.emph);
      T(s, "チームのリーダー", { x: x0 + 0.22, y: y0, w: 1.9, h: 0.26, fontSize: FS.note, color: C.soft, valign: "middle", wrap: false });
    }
    rect(s, head.x, head.y, hw, hH, TINT);
    person(s, head.x + 0.36, y0 + (hH - 0.17 * 2.17) / 2, 0.17, C.emph);
    T(s, [{ text: H0.tag ?? "責任者", options: { fontSize: FS.note, bold: true, color: C.gray, breakLine: true } }, { text: H0.role ?? "", options: { fontSize: FS.head, bold: true, color: C.ink, breakLine: !!H0.body } }, ...(H0.body ? [{ text: H0.body, options: { fontSize: FS.body, color: C.soft } }] : [])],
      { x: head.x + 0.68, y: y0, w: hw - 0.8, h: hH, valign: "middle", lineSpacingMultiple: 1.1 });
    // 直属（事務局など）：責任者の右に横の線でつなぐ
    const staff = [];
    (o.staff || []).forEach((sf, k) => {
      const sw = o.staffW ?? 2.6, sx = head.x + hw + 0.45 + k * (sw + 0.2), sh = hH - 0.12, sy = y0 + 0.06;
      segLine(s, k ? staff[k - 1].x + staff[k - 1].w : head.x + hw, y0 + hH / 2, sx, y0 + hH / 2, ln);
      rect(s, sx, sy, sw, sh, TINT);
      T(s, [{ text: sf.tag ?? "直属", options: { fontSize: FS.note, bold: true, color: C.gray, breakLine: true } }, { text: sf.role, options: { fontSize: FS.block, bold: true, color: C.ink, breakLine: !!sf.body } }, ...(sf.body ? [{ text: sf.body, options: { fontSize: FS.body, color: C.soft } }] : [])],
        { x: sx + pad, y: sy, w: sw - 2 * pad, h: sh, valign: "middle", lineSpacingMultiple: 1.05 });
      if (sf.body) checkWrap(sf.body, FS.body, sw - 2 * pad);
      // 役割（role）は1行で入れる。タグ・役割・本文の行が箱の高さに入るか、箱が右の余白に入るかも確かめる
      if (textW(sf.role, FS.block) > sw - 2 * pad) warn(`体制図：直属の箱の役割（${sf.role}）が幅に入らない（${FS.block}pt で全角${Math.floor((sw - 2 * pad) * 72 / FS.block)}字まで。staffW を広げるか、役割を縮める）`);
      // 1行の高さ（×1.2×1.05）には字面の上下の余白が含まれるので、箱の高さを 0.08 まで超えても見た目ではみ出さない（タグ＋役割の2行で確かめた）
      const sNeed = lineH(FS.note, 1.05) + lineCount(sf.role, FS.block, sw - 2 * pad) * lineH(FS.block, 1.05) + (sf.body ? lineCount(sf.body, FS.body, sw - 2 * pad) * lineH(FS.body, 1.05) : 0);
      if (sNeed > sh + 0.08) warn(`体制図：直属の箱（${sf.role}）の文字が箱の高さに入らない（本文を縮めるか、headH を広げる）`);
      if (sx + sw > W - M + 0.02) warn(`体制図：直属の箱（${sf.role}）が右の余白にはみ出す（staffW・headW を狭める）`);
      staff.push({ x: sx, y: sy, w: sw, h: sh });
    });
    // 責任者 → 部署：幹の線と分岐。社外は破線
    const bus1 = dy - g1 / 2, dcx = i => deptR[i].x + deptR[i].w / 2;
    segLine(s, midX, y0 + hH, midX, bus1, ln);
    if (lastInt >= 0) { segLine(s, Math.min(dcx(0), midX), bus1, Math.max(dcx(lastInt), midX), bus1, ln); for (let i = 0; i <= lastInt; i++) segLine(s, dcx(i), bus1, dcx(i), dy, ln); }
    if (firstExt >= 0) { segLine(s, Math.max(dcx(lastInt), midX), bus1, dcx(D.length - 1), bus1, dl); for (let i = firstExt; i < D.length; i++) segLine(s, dcx(i), bus1, dcx(i), dy, dl); }
    const team = [], member = [];
    D.forEach((d, i) => {
      const r0 = deptR[i], ext = !!d.ext;
      // 部署の箱：名前・人数・役目。主導は濃い塗り＋ink 2pt 枠＋「主導」
      rect(s, r0.x, dy, r0.w, dH, d.lead ? TINT_STRONG : ext ? C.white : TINT, d.lead ? { color: C.emph, width: 2 } : ext ? dl : null);
      T(s, deptRuns(d), { x: r0.x + pad, y: dy, w: r0.w - 2 * pad, h: dH, valign: "middle", lineSpacingMultiple: 1.15 });
      checkWrap(d.name, FS.head, r0.w - 2 * pad); if (deptSub(d)) checkWrap(deptSub(d), FS.body, r0.w - 2 * pad);
      // 部署 → チーム
      const ts = teamsOf(d), bus2 = ty - g2 / 2, line = ext ? dl : ln, tcx = j => teamX[i][j] + lfW / 2;
      if (ts.length > 1) { segLine(s, dcx(i), dy + dH, dcx(i), bus2, line); segLine(s, tcx(0), bus2, tcx(ts.length - 1), bus2, line); ts.forEach((_, j) => segLine(s, tcx(j), bus2, tcx(j), ty, line)); }
      else if (ts.length === 1) segLine(s, dcx(i), dy + dH, dcx(i), ty, line);
      team.push([]); member.push([]);
      ts.forEach((t, j) => {
        const tx = teamX[i][j], tr = { x: tx, y: ty, w: lfW, h: tH };
        team[i].push(tr);
        rect(s, tx, ty, lfW, tH, ext ? C.white : TINT, ext ? dl : null);
        const runs = [{ text: t.name, options: { fontSize: FS.block, bold: true, color: ext ? C.soft : C.ink, breakLine: !!t.leader } }];
        T(s, runs, { x: tx + pad, y: ty + 0.09, w: lfW - 2 * pad, h: tH - 0.18 - (t.leader ? lineH(FS.body, 1.2) + 0.04 : 0), valign: "middle", lineSpacingMultiple: 1.1 });
        checkWrap(t.name, FS.block, lfW - 2 * pad);
        // チームリーダー：小さなピクト＋役職（兼務なら「兼務」）
        if (t.leader) {
          const ly = ty + tH - 0.09 - lineH(FS.body, 1.2);
          person(s, tx + pad + 0.06, ly + 0.03, 0.07, ext ? C.pale : C.ink);
          T(s, [{ text: t.leader, options: { fontSize: FS.body, color: C.soft } }, ...(t.dual ? [{ text: "　兼務", options: { fontSize: FS.note, bold: true, color: C.gray } }] : [])],
            { x: tx + pad + 0.2, y: ly, w: lfW - 2 * pad - 0.2, h: lineH(FS.body, 1.2), valign: "middle", wrap: false });
          if (textW(`${t.leader}${t.dual ? "　兼務" : ""}`, FS.body) > lfW - 2 * pad - 0.2) warn(`体制図：「${t.name}」のリーダーの行がチームの幅に入らない`);
        }
        // メンバー：チームの左の縦の線から枝を出し、役割と人数を1行ずつ
        const ms = t.members || [];
        if (ms.length) {
          const spx = tx + 0.12, lastY = my + (ms.length - 0.5) * rowH;
          segLine(s, spx, ty + tH, spx, lastY, line);
          ms.forEach((m, k) => {
            const yy = my + k * rowH, cyy = yy + rowH / 2, mr = { x: spx + 0.14, y: yy, w: tx + lfW - spx - 0.14, h: rowH }, nw = textW(`${m.n}名`, FS.body) + 0.04;
            member[i][j] = member[i][j] || []; member[i][j].push(mr);
            segLine(s, spx, cyy, spx + 0.1, cyy, line);
            T(s, [{ text: m.role, options: { fontSize: FS.body, color: C.soft } }, ...(m.dual ? [{ text: "　兼務", options: { fontSize: FS.note, bold: true, color: C.gray } }] : [])],
              { x: mr.x, y: yy, w: mr.w - nw, h: rowH, valign: "middle", wrap: false });
            T(s, `${m.n}名`, { x: mr.x + mr.w - nw, y: yy, w: nw, h: rowH, fontSize: FS.body, bold: true, color: ext ? C.soft : C.ink, align: "right", valign: "middle", wrap: false });
            if (textW(`${m.role}${m.dual ? "　兼務" : ""}`, FS.body) > mr.w - nw + 0.01) warn(`体制図：「${t.name}」の役割（${m.role}）が幅に入らない`);
          });
        }
      });
    });
    return { head, staff: k => staff[k], dept: i => ({ x: deptR[i].x, y: dy, w: deptR[i].w, h: dH }), team: (i, j) => team[i][j], member: (i, j, k) => (member[i][j] || [])[k] };
  };

  // アイコンの PNG の場所（構成図のノード・まとまりの印）。opts.iconDir（既定は deck-parts.js と同じ階層の icons/）にある <名前>.png
  //   PNG は icons/make-icons.js が react-icons から作る（ブランドは実物の色、汎用は ink 1色）。名前の一覧は make-icons.js の BRAND・MULTI・GENERIC
  //   名前の代わりに .png のパスを渡してもよい（相対パスは作業フォルダから）。無いアイコンは警告して描かない（別の印で代用しない）
  const ICON_DIR = opts.iconDir || assetDirs("icons").find(d => fs.existsSync(path.join(d, "user.png"))) || path.join(__dirname, "icons");
  // 探す順：opts.iconDir → テーマの icons（自分のアイコンのフォルダ。同じ名前なら同梱の絵より先に使う）→ 同梱の icons/
  const ICON_DIRS = opts.iconDir ? [ICON_DIR, ...THEME_ICON_DIRS] : [...THEME_ICON_DIRS, ICON_DIR];
  const findIcon = nm => { for (const d of ICON_DIRS) { const p = path.join(d, `${nm}.png`); if (fs.existsSync(p)) return p; } return null; };
  const isUserIcon = p => !!p && THEME_ICON_DIRS.some(d => path.resolve(path.dirname(p)) === path.resolve(d)) && path.resolve(path.dirname(p)) !== path.resolve(ICON_DIR);
  // 別名：ChatGPT・Codex は OpenAI の印、パソコン・地球は汎用の名前へ
  const ICON_ALIAS = { chatgpt: "openai", codex: "openai", pc: "laptop", globe: "internet" };
  const ICON_MISS = new Set();
  const iconPath = name => {
    if (!name) return null;
    if (/\.png$/i.test(name)) { const p = path.resolve(name); if (fs.existsSync(p)) return p; }
    else { const p = findIcon(ICON_ALIAS[name] || name); if (p) return p; }
    if (!ICON_MISS.has(name)) { ICON_MISS.add(name); warn(`アイコン「${name}」が見つからない（${ICON_DIRS.join("・")}）。opts.iconDir か名前を確かめる`); }
    return null;
  };

  // アイコンを1つ置く（正方形の PNG。cx, cy は中心、d は一辺）。置けたら true
  // 線の絵（1色のアイコン）の名前。色は scripts/tint_icons.py で、資料の色の組に合わせて塗り直す（公式のアイコンは塗らない）
  const LINE_ICONS = new Set("user users laptop desktop smartphone tablet internet cloud server compute function container database storage bucket folder file files spreadsheet api gateway webhook plug queue network auth shield lock monitoring logs search vector-search ai brain bot workflow settings chart code terminal git-branch pull-request mail chat building".split(" "));
  const isLineIcon = nm => LINE_ICONS.has(String(nm).replace(/-[wer]$/, "")) || /^pic-/.test(String(nm));
  const ICON_TINT = (() => { try { return JSON.parse(fs.readFileSync(path.join(ICON_DIR, ".tint.json"), "utf8")); } catch (e) { return null; } })();
  let lineIconUsed = false;
  // アイコンを置く。tone で色ちがいを選ぶ: "w"＝白（色を塗った丸の上）、"e"＝強調の色、"r"＝問題の色。その色の画像（<名前>-<tone>.png）が無ければ、元の色で置く
  const icon = (s, name0, cx, cy, d = 0.4, tone) => { const name = tone && findIcon(`${ICON_ALIAS[name0] || name0}-${tone}`) ? `${ICON_ALIAS[name0] || name0}-${tone}` : name0; const p = iconPath(name); if (isLineIcon(name) && !isUserIcon(p)) lineIconUsed = true; if (p) { s.addImage({ path: p, altText: `アイコン：${path.basename(String(name0), ".png")}`, x: cx - d / 2, y: cy - d / 2, w: d, h: d }); (ICONS_ON[page] = ICONS_ON[page] || {})[ICON_ALIAS[name] || name] = 1; } return !!p; };

  // 2軸の図の名前の置き場を決める補助：矩形どうしの重なりの面積（矩形は [x, y, w, h]）
  const mxRectHit = (a, b) => Math.max(0, Math.min(a[0] + a[2], b[0] + b[2]) - Math.max(a[0], b[0])) * Math.max(0, Math.min(a[1] + a[3], b[1] + b[3]) - Math.max(a[1], b[1]));

  // matrix2x2 のポジショニングマップ（引数は matrix2x2 の説明）
  const mxMap = (s, o, Q) => {
    const x0 = o.x ?? M, top = o.y ?? BODY, right = x0 + (o.w ?? (W - M - x0)), bottom = o.bottom ?? BOTTOM;
    const side = o.side || null, sw = side ? (o.sideW ?? 3.3) : 0, sg = side ? 0.4 : 0;
    const lab = 0.36, px0 = x0, px1 = right - sw - sg, py0 = top + lab, py1 = bottom - lab, pw = px1 - px0, ph = py1 - py0;
    const X = v => px0 + v * pw, Yv = v => py0 + (1 - v) * ph, cxm = X(0.5), cym = Yv(0.5);
    const pts = o.points || [], tgt = o.target, hasSelf = pts.some(p => p.self);
    if (hasSelf && tgt !== undefined) warn("2軸の図：自社の点と狙う象限の両方を強調している。強調は1ページ1種類にする");
    const quad = i => ({ x: i % 2 ? cxm : px0, y: i >= 2 ? cym : py0, w: pw / 2, h: ph / 2 });
    // 地と狙う象限
    rect(s, px0, py0, pw, ph, TINT);
    if (tgt !== undefined) { const q = quad(tgt); rect(s, q.x, q.y, q.w, q.h, TINT_STRONG, { color: C.emph, width: 2 }); }
    // 空白地帯：地の上に白く抜いた破線の楕円
    const obst = []; // 名前を置かない所 [x, y, w, h]
    let zone = null;
    if (o.zone) {
      const z = o.zone, rx = (z.rx ?? (z.w ?? 0.24) / 2) * pw, ry = (z.ry ?? (z.h ?? 0.24) / 2) * ph, zx = X(z.x), zy = Yv(z.y);
      zone = { cx: zx, cy: zy, rx, ry };
      s.addShape(SH.OVAL, { x: zx - rx, y: zy - ry, w: 2 * rx, h: 2 * ry, fill: { color: C.white }, line: { color: C.soft, width: 1, dashType: "dash" } });
      const tw = rx * 2 * 0.8, runs = [{ text: z.label, options: { fontSize: FS.block, bold: true, color: C.ink, breakLine: !!z.body } }];
      if (z.body) runs.push({ text: z.body, options: { fontSize: FS.body, color: C.soft } });
      T(s, runs, { x: zx - tw / 2, y: zy - ry * 0.8, w: tw, h: ry * 1.6, align: "center", valign: "middle", lineSpacingMultiple: 1.15 });
      checkWrap(z.label, FS.block, tw); if (z.body) checkWrap(z.body, FS.body, tw);
      obst.push([zx - rx * 0.92, zy - ry * 0.92, rx * 1.84, ry * 1.84]);
    }
    // 軸：中央で交わる灰の両矢印。両端の意味は bold ink、高い側に軸の名前を soft で添える
    const axl = { color: C.gray, width: 1.25, beginArrowType: "triangle", endArrowType: "triangle" };
    s.addShape(SH.LINE, { x: px0 + 0.02, y: cym, w: pw - 0.04, h: 0, line: axl });
    s.addShape(SH.LINE, { x: cxm, y: py0 + 0.02, w: 0, h: ph - 0.04, line: axl });
    const XA = o.xAxis || {}, YA = o.yAxis || {};
    const endRuns = (word, name) => [{ text: word || "", options: { bold: true, color: C.ink, fontSize: FS.block } }, ...(name ? [{ text: `（${String(name).replace(/\n/g, "")}）`, options: { color: C.soft, fontSize: FS.body } }] : [])];
    const endW = (word, name) => textW(word || "", FS.block) + (name ? textW(`（${String(name).replace(/\n/g, "")}）`, FS.body) : 0) + 0.06;
    if (XA.low) { const w = endW(XA.low); T(s, endRuns(XA.low), { x: px0 + 0.12, y: cym + 0.07, w, h: 0.28, valign: "middle", wrap: false }); obst.push([px0 + 0.12, cym + 0.07, w, 0.28]); }
    if (XA.high) { const w = endW(XA.high, XA.name); T(s, endRuns(XA.high, XA.name), { x: px1 - 0.12 - w, y: cym + 0.07, w, h: 0.28, align: "right", valign: "middle", wrap: false }); obst.push([px1 - 0.12 - w, cym + 0.07, w, 0.28]); }
    if (YA.high) { const w = endW(YA.high, YA.name); T(s, endRuns(YA.high, YA.name), { x: cxm - w / 2, y: top, w, h: lab - 0.06, align: "center", valign: "bottom", wrap: false }); if (w > pw) warn(`2軸の図：縦軸の上の言葉が長い（${YA.high}）`); }
    if (YA.low) { const w = endW(YA.low); T(s, endRuns(YA.low), { x: cxm - w / 2, y: py1 + 0.06, w, h: lab - 0.06, align: "center", valign: "top", wrap: false }); }
    // 象限の名前：外側の角。狙う象限だけ ink、他は gray（見出しなので gray でよい）
    Q.forEach((q, i) => {
      if (!q || !q.name) return;
      const r = quad(i), rt = i % 2 === 1, low = i >= 2, on = i === tgt, w = r.w * 0.46;
      const runs = [{ text: q.name, options: { fontSize: FS.block, bold: true, color: on ? C.ink : C.gray, breakLine: !!q.body } }];
      if (q.body) runs.push({ text: q.body, options: { fontSize: FS.body, color: C.soft } });
      const hh = lineH(FS.block, 1.15) + (q.body ? lineCount(q.body, FS.body, w) * lineH(FS.body, 1.15) : 0) + 0.04;
      const bx = rt ? r.x + r.w - 0.14 - w : r.x + 0.14, by = low ? r.y + r.h - 0.12 - hh : r.y + 0.12;
      T(s, runs, { x: bx, y: by, w, h: hh, align: rt ? "right" : "left", valign: low ? "bottom" : "top", lineSpacingMultiple: 1.15 });
      const nw = Math.max(textW(q.name, FS.block), q.body ? Math.min(w, textW(q.body, FS.body)) : 0) + 0.04;
      obst.push([rt ? bx + w - nw : bx, by, nw, hh]);
    });
    // バブルの大きさ：面積を量に比例（直径 ∝ √量）
    const sz = o.size || null, vmax = sz ? (sz.max ?? Math.max(...pts.map(p => p.size || 0), 1)) : 1, dMax = o.dMax ?? 0.95;
    const dOf = v => Math.max(0.2, dMax * Math.sqrt(Math.max(v, 0) / vmax));
    const D = pts.map(p => sz && p.size != null ? Math.max(p.self ? 0.26 : 0.2, dOf(p.size)) : p.self ? 0.42 : 0.34);
    const P = pts.map(p => [X(p.x), Yv(p.y)]);
    // 円が軸の言葉・象限の名前・空白地帯に重なるなら知らせる（点の位置の値を少し動かす）
    P.forEach(([cx, cy], i) => { const q = D[i] / 2 * 0.85; if (obst.some(b => mxRectHit([cx - q, cy - q, 2 * q, 2 * q], b) > 0.004)) warn(`2軸の図：「${pts[i].name}」の円が軸の言葉・象限の名前・空白地帯に重なる`); });
    // 大きい円から描く（小さい円を上に）
    pts.map((p, i) => i).sort((a, b) => D[b] - D[a]).forEach(i => {
      const [cx, cy] = P[i], d = D[i], self = !!pts[i].self;
      s.addShape(SH.OVAL, { x: cx - d / 2, y: cy - d / 2, w: d, h: d, fill: { color: self ? C.emph : sz ? C.bar : C.faint }, line: { color: C.white, width: 1.5 } });
    });
    // 自社の動き：円の縁から空白地帯の縁（または to）へ、紺の破線の矢印
    if (o.move && P[o.move.from]) {
      const [fx, fy] = P[o.move.from], fd = D[o.move.from];
      let tx, ty;
      if (o.move.to) { tx = X(o.move.to[0]); ty = Yv(o.move.to[1]); }
      else if (zone) {
        const dx = zone.cx - fx, dy = zone.cy - fy, a = (dx / zone.rx) ** 2 + (dy / zone.ry) ** 2, b = 2 * (((fx - zone.cx) * dx) / zone.rx ** 2 + ((fy - zone.cy) * dy) / zone.ry ** 2), c = ((fx - zone.cx) / zone.rx) ** 2 + ((fy - zone.cy) / zone.ry) ** 2 - 1;
        const t = (-b - Math.sqrt(Math.max(0, b * b - 4 * a * c))) / (2 * a); tx = fx + dx * t; ty = fy + dy * t;
      }
      if (tx !== undefined) {
        const L = Math.hypot(tx - fx, ty - fy) || 1, r0 = fd / 2 + 0.05;
        arrowSeg(s, fx + (tx - fx) * r0 / L, fy + (ty - fy) * r0 / L, tx - (tx - fx) * 0.04 / L, ty - (ty - fy) * 0.04 / L, { color: C.emph, width: 1.75, dashType: "dash" });
        // 矢印の上に名前を置かない（矢印に沿って小さな矩形を並べて避ける）
        for (let t = r0; t < L; t += 0.08) obst.push([fx + (tx - fx) * t / L - 0.04, fy + (ty - fy) * t / L - 0.04, 0.08, 0.08]);
      }
    }
    // 名前：8方向の候補から、他の円・名前・軸の言葉・空白地帯と重ならず、地からはみ出さない位置を選ぶ
    const placed = [];
    const ORDER = ["r", "l", "t", "b", "tr", "br", "tl", "bl"];
    pts.forEach((p, i) => {
      const [cx, cy] = P[i], rr = D[i] / 2, self = !!p.self, fz = self ? FS.head : FS.block;
      const w = Math.max(textW(p.name, fz), p.sub ? textW(p.sub, FS.body) : 0) + 0.06, h = lineH(fz, 1.1) + (p.sub ? lineH(FS.body, 1.1) : 0) + 0.02;
      const at = pos => {
        const k = rr * 0.72, g = 0.06;
        return { r: [cx + rr + g, cy - h / 2], l: [cx - rr - g - w, cy - h / 2], t: [cx - w / 2, cy - rr - 0.03 - h], b: [cx - w / 2, cy + rr + 0.03],
          tr: [cx + k + 0.02, cy - k - h + 0.04], br: [cx + k + 0.02, cy + k - 0.04], tl: [cx - k - 0.02 - w, cy - k - h + 0.04], bl: [cx - k - 0.02 - w, cy + k - 0.04] }[pos];
      };
      const score = (pos, idx) => {
        const [lx, ly] = at(pos), R = [lx, ly, w, h];
        let sc = idx * 0.002;
        P.forEach(([qx, qy], j) => { if (j !== i) { const q = D[j] / 2 * 0.9; sc += mxRectHit(R, [qx - q, qy - q, 2 * q, 2 * q]) * 12; } });
        placed.forEach(b => { sc += mxRectHit(R, b) * 20; });
        obst.forEach(b => { sc += mxRectHit(R, b) * 15; });
        sc += (w * h - mxRectHit(R, [px0 + 0.05, py0 + 0.05, pw - 0.1, ph - 0.1])) * 30;
        // 軸の線をまたぐ置き方は避ける（名前が線で切れて読みにくい）
        let soft = 0;
        if (lx < cxm && lx + w > cxm) soft += 0.3 * h; if (ly < cym && ly + h > cym) soft += 0.3 * w;
        return [sc + soft, sc];
      };
      const cand = p.pos ? [p.pos] : ORDER;
      const best = cand.map((pos, idx) => [pos, ...score(pos, idx)]).sort((a, b) => a[1] - b[1])[0];
      if (!p.pos && best[2] > 0.15) warn(`2軸の図：「${p.name}」の名前の置き場が見つからない（重なる）。点を動かすか pos を指定する`);
      const pos = best[0], [lx, ly] = at(pos), al = pos === "r" || pos === "tr" || pos === "br" ? "left" : pos === "l" || pos === "tl" || pos === "bl" ? "right" : "center";
      const runs = [{ text: p.name, options: { fontSize: fz, bold: self, color: C.ink, breakLine: !!p.sub } }];
      if (p.sub) { warnSub("2軸の点"); runs.push({ text: p.sub, options: { fontSize: FS.body, color: C.soft } }); }
      T(s, runs, { x: lx, y: ly, w, h, align: al, valign: "middle", lineSpacingMultiple: 1.1, wrap: false });
      placed.push([lx, ly, w, h]);
    });
    // 右の列：バブルの凡例と読み取りのブロック
    if (side || sz) {
      const sx = px1 + sg; let sy = top;
      if (sz && sz.ref) {
        const refs = sz.ref, ds = refs.map(dOf), gx = 0.18, maxD = Math.max(...ds);
        T(s, sz.label || "円の大きさ", { x: sx, y: sy, w: sw, h: 0.28, fontSize: FS.body, color: C.soft, valign: "middle", wrap: false });
        let lx = sx; const base = sy + 0.3 + maxD;
        refs.forEach((v, k) => {
          const d = ds[k], cw = Math.max(d, textW(fmtNum(v, { unit: sz.unit || "" }), FS.body) + 0.04);
          s.addShape(SH.OVAL, { x: lx + (cw - d) / 2, y: base - d, w: d, h: d, fill: { color: C.bar }, line: { color: C.white, width: 1 } });
          T(s, fmtNum(v, { unit: sz.unit || "" }), { x: lx - 0.1, y: base + 0.04, w: cw + 0.2, h: 0.26, fontSize: FS.body, color: C.soft, align: "center", wrap: false });
          lx += cw + gx;
        });
        if (lx - gx > sx + sw + 0.01) warn("2軸の図：バブルの凡例が右の列に入らない。ref を減らす");
        sy = base + 0.4;
      }
      if (side) {
        const gap = GAP.std, n = side.length, bh = (bottom - sy - gap * (n - 1)) / n;
        side.forEach((b, k) => {
          const by = sy + k * (bh + gap);
          rect(s, sx, by, sw, bh, TINT);
          T(s, boxRuns(b.head, b.body), { x: sx + PAD.x, y: by, w: sw - 2 * PAD.x, h: bh, valign: "middle", lineSpacingMultiple: 1.2 });
          const iw = sw - 2 * PAD.x, need = lineCount(b.head, FS.head, iw) * FS.head * 1.25 / 72 + (b.body ? (6 + lineCount(b.body, FS.body, iw) * FS.body) * 1.25 / 72 : 0);
          if (need > bh - 0.1) warn(`2軸の図の読み取りの文字があふれる（「${String(b.head).slice(0, 12)}…」）`);
          checkWrap(b.head, FS.head, iw); if (b.body) checkWrap(b.body, FS.body, iw);
        });
      }
    }
    return { px: X, py: Yv, pt: i => P[i], d: i => D[i], quad };
  };

  // matrix2x2 の2×2（象限に名前・方針・項目。引数は matrix2x2 の説明）
  const mxGrid = (s, o, Q) => {
    const x0 = o.x ?? M, top = o.y ?? BODY, right = x0 + (o.w ?? (W - M - x0)), bottom = o.bottom ?? BOTTOM;
    const YN = (o.yAxis || {}).name, lw = o.labelW ?? Math.max(1.0, ...(YN ? String(YN).split("\n").map(l => textW(l, FS.block) + 0.12) : [0])), gx0 = x0 + lw + 0.22, gy0 = top, gy1 = bottom - 0.4, g = GAP.std + 0.02;
    const qw = (right - gx0 - g) / 2, qh = (gy1 - gy0 - g) / 2, tgt = o.target, ic = o.itemCols ?? 2;
    const quad = i => ({ x: gx0 + (i % 2) * (qw + g), y: gy0 + Math.floor(i / 2) * (qh + g), w: qw, h: qh });
    // 軸：左と下の灰の矢印。軸の名前は bold ink、両端は soft
    const axl = { color: C.gray, width: 1.25 }, ax = gx0 - 0.14, ay = gy1 + 0.12;
    arrowSeg(s, ax, ay, ax, gy0, axl);
    arrowSeg(s, ax, ay, right, ay, axl);
    const XA = o.xAxis || {}, YA = o.yAxis || {}, yw = ax - 0.1 - x0;
    if (XA.low) T(s, XA.low, { x: gx0, y: ay + 0.05, w: 3, h: 0.26, fontSize: FS.body, color: C.soft, wrap: false });
    if (XA.high) T(s, XA.high, { x: right - 3, y: ay + 0.05, w: 3, h: 0.26, fontSize: FS.body, color: C.soft, align: "right", wrap: false });
    if (XA.name) T(s, XA.name, { x: gx0 + (right - gx0) / 2 - 2, y: ay + 0.04, w: 4, h: 0.28, fontSize: FS.block, bold: true, align: "center", wrap: false });
    if (YA.high) T(s, YA.high, { x: x0, y: gy0, w: yw, h: 0.26, fontSize: FS.body, color: C.soft, align: "right", wrap: false });
    if (YA.low) T(s, YA.low, { x: x0, y: gy1 - 0.26, w: yw, h: 0.26, fontSize: FS.body, color: C.soft, align: "right", wrap: false });
    if (YA.name) T(s, YA.name, { x: x0, y: gy0 + (gy1 - gy0) / 2 - 0.4, w: yw, h: 0.8, fontSize: FS.block, bold: true, align: "right", valign: "middle", lineSpacingMultiple: 1.15 });
    if (YA.name) String(YA.name).split("\n").forEach(l => { if (textW(l, FS.block) > yw) warn(`2×2：縦軸の名前が左の列に入らない（${l}）。"\\n" で折る`); });
    // 項目の札：全象限で同じ大きさ。一番狭い象限の空きと一番多い行数から、高さを決める（1行の札 0.44、補足つき 0.6 まで）
    const lists = Q.map(q => ((q && q.items) || []).map(it => typeof it === "string" ? { name: it } : it)), hasSub = lists.some(l => l.some(it => it.sub));
    const headOf = (q, i) => !q ? 0 : 0.14 + lineH(i === tgt ? FS.lead : FS.head, 1.1) + 0.04 + ((q.policy || q.body) ? lineCount(q.policy || q.body, FS.body, qw - 2 * PAD.x) * lineH(FS.body, 1.2) : 0) + 0.14;
    const rowsMax = Math.max(1, ...lists.map(l => Math.ceil(l.length / ic))), availMin = Math.min(...Q.map((q, i) => qh - headOf(q, i) - 0.14));
    const rg = 0.08, ch = Math.min(hasSub ? 0.6 : 0.44, (availMin - rg * (rowsMax - 1)) / rowsMax);
    if (ch < (hasSub ? 0.42 : 0.3)) warn(`2×2：項目の札が低すぎる（${ch.toFixed(2)}）。項目を減らすか itemCols を増やす`);
    const items = [];
    Q.forEach((q, i) => {
      const r = quad(i), on = i === tgt, px = PAD.x, iw = r.w - 2 * px;
      rect(s, r.x, r.y, r.w, r.h, on ? TINT_STRONG : TINT, on ? { color: C.emph, width: 2 } : null);
      if (!q) { items.push([]); return; }
      const nfs = on ? FS.lead : FS.head;
      let y = r.y + 0.14;
      T(s, q.name, { x: r.x + px, y, w: iw, h: lineH(nfs, 1.1), fontSize: nfs, bold: true, valign: "middle", wrap: false });
      if (textW(q.name, nfs) > iw) warn(`2×2：象限の名前が長い（${q.name}）`);
      y += lineH(nfs, 1.1) + 0.04;
      const pol = q.policy || q.body;
      if (pol) { const n = lineCount(pol, FS.body, iw); T(s, pol, { x: r.x + px, y, w: iw, h: n * lineH(FS.body, 1.2), fontSize: FS.body, color: C.soft, lineSpacingMultiple: 1.2 }); checkWrap(pol, FS.body, iw); y += n * lineH(FS.body, 1.2); }
      y += 0.14;
      const list = lists[i], cg = 0.1, cw = (iw - cg * (ic - 1)) / ic;
      const box = [];
      list.forEach((it, k) => {
        const bx = r.x + px + (k % ic) * (cw + cg), by = y + Math.floor(k / ic) * (ch + rg);
        rect(s, bx, by, cw, ch, C.white);
        const runs = [{ text: it.name, options: { fontSize: FS.block, color: C.ink, breakLine: !!it.sub } }];
        if (it.sub) runs.push({ text: it.sub, options: { fontSize: FS.body, color: C.soft } });
        T(s, runs, { x: bx + 0.12, y: by, w: cw - 0.2, h: ch, valign: "middle", lineSpacingMultiple: 1.1 });
        if (textW(it.name, FS.block) > cw - 0.2) warn(`2×2：項目の名前が札に入らない（${it.name}）`);
        if (it.sub && textW(it.sub, FS.body) > cw - 0.2) warn(`2×2：項目の補足が札に入らない（${it.sub}）`);
        box.push({ x: bx, y: by, w: cw, h: ch });
      });
      items.push(box);
    });
    return { quad, item: (i, k) => items[i] && items[i][k] };
  };

  // ---- 王道のグラフ（図形で描く） ----
  // 共通の規律: 主役の系列・区分だけ濃紺（ink）、他は灰（C.bar／C.faint）か区分色（CAT）。目盛り線は引かず、値と名前を図の中に直接書く（凡例の代わり）。
  // ネイティブのグラフ（pptxgenjs の pie・line・scatter など）は、凡例が要る・ラベルの位置を PowerPoint が決める・1点だけの強調ができない・目盛り線が残る、のでここでは使わない。
  // 補助: 切りのよい上限（1・2・2.5・5 × 10^n）。軸の目盛りを置く図（散布図）で使う
  const chNice = v => { if (!(v > 0)) return 1; const p = Math.pow(10, Math.floor(Math.log10(v))), f = v / p; return (f <= 1 ? 1 : f <= 2 ? 2 : f <= 2.5 ? 2.5 : f <= 5 ? 5 : 10) * p; };

  // 補助: ラベルの縦位置を重ならないようにずらす（items の y＝中心を書き換える）。gap＝中心どうしの最小の間、lo〜hi に収める
  const chSpread = (items, gap, lo = -Infinity, hi = Infinity) => {
    const a = [...items].sort((p, q) => p.y - q.y);
    a.forEach((it, i) => { it.y = Math.max(it.y, lo); if (i && it.y - a[i - 1].y < gap) it.y = a[i - 1].y + gap; });
    for (let i = a.length - 1; i >= 0; i--) { if (a[i].y > hi) a[i].y = hi; if (i < a.length - 1 && a[i + 1].y - a[i].y < gap) a[i].y = a[i + 1].y - gap; }
    return items;
  };

  // 補助: chSpread の、項目ごとに高さの違う版。items の hh＝半分の高さ。隣どうしの中心の間を hh の和以上にする（主役の値 36pt の行を広く取る）
  const chSpreadH = (items, lo = -Infinity, hi = Infinity) => {
    const a = [...items].sort((p, q) => p.y - q.y), gap = (p, q) => p.hh + q.hh;
    a.forEach((it, i) => { it.y = Math.max(it.y, lo); if (i && it.y - a[i - 1].y < gap(a[i - 1], it)) it.y = a[i - 1].y + gap(a[i - 1], it); });
    for (let i = a.length - 1; i >= 0; i--) { if (a[i].y > hi) a[i].y = hi; if (i < a.length - 1 && a[i + 1].y - a[i].y < gap(a[i], a[i + 1])) a[i].y = a[i + 1].y - gap(a[i], a[i + 1]); }
    return items;
  };

  // 補助: 折れ線・多角形を1つの図形で描く（custGeom）。pts は絶対座標 [[x, y], ...]。o.close で閉じる、o.fill / o.transparency / o.line
  //   線を区間ごとの図形に分けないので、PowerPoint で1本の線として選べる
  const chPath = (s, pts, o = {}) => {
    const xs = pts.map(p => p[0]), ys = pts.map(p => p[1]);
    const x = Math.min(...xs), y = Math.min(...ys), w = Math.max(Math.max(...xs) - x, 0.01), h = Math.max(Math.max(...ys) - y, 0.01);
    const points = pts.map(([px, py]) => ({ x: px - x, y: py - y }));
    if (o.close) points.push({ close: true });
    s.addShape(SH.CUSTOM_GEOMETRY, { x, y, w, h, points, fill: o.fill ? { color: o.fill, transparency: o.transparency || 0 } : { type: "none" }, line: o.line || { type: "none" } });
  };

  // 補助: 点（丸）。fill を null にすると白抜き（line で輪郭）
  const chDot = (s, cx, cy, d, fill, line) => s.addShape(SH.OVAL, { x: cx - d / 2, y: cy - d / 2, w: d, h: d, fill: { color: fill || C.white }, line: line || { type: "none" } });

  // 補助: 線の書式（dashType を渡したときだけ付ける。undefined を XML に書かない）
  const chLine = (color, width, dash) => (dash ? { color, width, dashType: dash } : { color, width });

  // 基準の線：目標・平均・中央値・閾値を、棒・折れ線・ヒストグラムに重ねる。横（y を渡す）か縦（x を渡す）
  //   o: { y, x0, x1 } か { x, y0, y1 }, o.label（線の真上に置く）, o.at（横線のラベルの端 "end"|"start"）, o.kind（"target"＝目標の太い破線 soft／"avg"＝平均・中央値の細い破線 gray）
  //   目標は赤にしない（赤は問題だけ）。目標に届かない値を問題として示すときは、値の側（棒・点）を赤にする
  const refLine = (s, o) => {
    const target = (o.kind || "target") === "target", line = chLine(target ? C.soft : C.gray, target ? 1.75 : 1, "dash");
    if (o.y != null) {
      s.addShape(SH.LINE, { x: o.x0, y: o.y, w: o.x1 - o.x0, h: 0, line });
      if (o.label) { const end = (o.at || "end") === "end"; T(s, o.label, { x: end ? o.x1 - 4 : o.x0, y: o.y - 0.28, w: 4, h: 0.25, fontSize: FS.body, color: C.soft, bold: target, align: end ? "right" : "left", valign: "bottom", wrap: false }); }
    } else {
      s.addShape(SH.LINE, { x: o.x, y: o.y0, w: 0, h: o.y1 - o.y0, line });
      if (o.label) T(s, o.label, { x: o.x - 2, y: o.y0 - 0.3, w: 4, h: 0.26, fontSize: FS.body, color: C.soft, bold: target, align: "center", valign: "bottom", wrap: false });
    }
  };

  // 折れ線（図形）：推移。主役の系列は濃紺の太線と塗りの点、他は灰の細線と白抜きの点、見込みは点線、目標は太い破線と真上のラベル（components.md「推移の線」）
  //   系列名と最後の値は線の右端に直接書き、凡例を置かない。縦軸と目盛り線は置かず、主役の値を点の上に書く
  //   ネイティブの折れ線は、系列名が凡例になる・主役の点だけに値を書けない・見込みの点線と目標線を描き分けられないので使わない
  //   o.cats: [期], o.series: [{ name, values（null で欠け）, main?, color?, from?（この番号から先は見込みの点線。最初の値の番号を渡すと全区間が点線＝推計の系列）, below?（値を点の下に）, labelAt?: [番号]（値を書く点を指定。他の線と重なるとき）, cagr?（文字列。例 "年+18.6%"。右端の「最後の値　系列名」の後ろに続ける） }]
  //   o.values: "main"（既定。主役の全点、点が14以上なら最初と最後）| "ends"（全系列の最初と最後）| "none"
  //   o.targets: [{ value, label, at? }]（目標の線）, o.marks: [{ at: 区分の番号, label }]（出来事の縦の破線）
  //   o.x0, o.x1, o.top, o.bottom（既定 M+0.5〜右端の名前の手前、2.85〜6.05）, o.min（既定 0。負の値があれば、その少し下から始め、0 の線を引く（o.zero: false で引かない）。折れ線は棒と違い 0 から始めなくてよい）, o.max, o.fmt, o.labelW（右端の名前の幅 1.9）, o.catEvery（区分名を n 個おき）
  //   o.log（true で底10の対数軸。桁の違う系列の伸びを比べるとき。傾きが年平均成長率を示す。o.min は必須で 0 より大きい値、データも 0 より大きい値だけ。満たさなければ throw）
  //   o.focusValue: [系列の番号, 点の番号]（主役の点。その値を主役の値 36pt で点の上に部品が描き、その点の小さな値は描かない。最後の点なら右端の名前から値を外す）
  //     o.focusBelow（true で点の下に）, o.focusProblem（true で赤）。呼ぶ側で値を空にして focusValue を手で置かない
  //   戻り値: { xOf(i), yOf(v), pt(j, i) } … 前面から指す位置
  const lineChart = (s, o) => {
    chartPage();
    const ser = o.series, n = o.cats.length, fmt = o.fmt || (v => fmtNum(v));
    const FV = o.focusValue, isFV = (j, i) => !!FV && FV[0] === j && FV[1] === i;
    if (FV && (!ser[FV[0]] || ser[FV[0]].values[FV[1]] == null)) warn(`折れ線：主役の値（focusValue: [${FV}]）の点が無い`);
    const lw = o.labelW ?? 1.9, x0 = o.x0 ?? M + 0.5, x1 = o.x1 ?? W - M - lw - 0.2, top = o.top ?? 2.85, bottom = o.bottom ?? BOTTOM - 0.4;
    // 折れ線は3本まで。推計（予測）が何本もあるときは、線を引かず、推計の幅（range）の帯で見せる（2026-11-24）
    if (ser.length > 3) warn(`折れ線が ${ser.length} 本ある。3本までにする。推計がいくつもあるときは、o.range（推計の幅の帯）で見せる`);
    const R = o.range, all = ser.flatMap(sr => sr.values.filter(v => v != null)).concat((o.targets || []).map(t => t.value)).concat(R ? [...R.low, ...R.high].filter(v => v != null) : []);
    if (o.log) warn("折れ線の対数軸は、読み手が大きさの差を見誤る。市場の大きさ・成長は、0 から始まる縦棒と年平均成長率（colBars の cagr）で見せる");
    if (o.log && !(o.min > 0)) throw new Error(`lineChart：対数軸（log）には 0 より大きい o.min が必要（${o.min}）`);
    if (o.log && all.some(v => v <= 0)) throw new Error("lineChart：対数軸（log）に 0 以下の値は描けない");
    const lo = Math.min(...all), min = o.min ?? (lo < 0 ? lo * 1.08 : 0), max = o.max ?? Math.max(...all) * 1.08, step = n > 1 ? (x1 - x0) / (n - 1) : 0;
    if (o.min != null && lo < o.min) warn(`折れ線の値（${fmt(lo)}）が o.min（${fmt(o.min)}）より小さく、図の下にはみ出す`);
    const scale = o.log ? Math.log10 : (v => v);
    const xOf = i => x0 + i * step, yOf = v => bottom - (scale(v) - scale(min)) / (scale(max) - scale(min)) * (bottom - top);
    hline(s, x0 - 0.25, bottom, x1 - x0 + 0.5, { color: C.gray });
    // 負の値があるときは、0 の線を引き、左端に「0」と書く（増減の境が分かるように）
    if (!o.log && o.zero !== false && min < 0 && max > 0 && !(o.targets || []).some(t => t.value === 0)) { // 0 の線を自分で引く（refLine など）なら o.zero: false
       hline(s, x0 - 0.25, yOf(0), x1 - x0 + 0.5, { color: C.gray, width: 1 }); T(s, "0", { x: x0 - 0.85, y: yOf(0) - 0.15, w: 0.5, h: 0.3, fontSize: FS.body, color: C.soft, align: "right", valign: "middle", wrap: false }); }
    const catW = Math.max(...o.cats.map(c => textW(c, FS.body))) + 0.12; // 区分名が隣に重ならない間隔で間引く
    const every = o.catEvery ?? Math.max(Math.ceil(n / 12), Math.ceil(catW / Math.max(step, 0.01)));
    o.cats.forEach((c, i) => { if (i !== n - 1 && (i % every || (n - 1 - i) < every)) return; T(s, c, { x: xOf(i) - Math.max(1.2, catW) / 2, y: bottom + 0.1, w: Math.max(1.2, catW), h: 0.32, fontSize: FS.body, color: C.soft, align: "center", wrap: false }); });
    (o.marks || []).forEach(m => { s.addShape(SH.LINE, { x: xOf(m.at), y: top - 0.05, w: 0, h: bottom - top + 0.05, line: chLine(C.gray, 1, "dash") }); T(s, m.label, { x: xOf(m.at) - 1.5, y: top - 0.36, w: 3, h: 0.26, fontSize: FS.note, color: C.gray, align: "center", valign: "bottom", wrap: false }); });
    (o.targets || []).forEach(t => refLine(s, { y: yOf(t.value), x0: x0 - 0.25, x1: x1 + 0.05, label: t.label, at: t.at || "start" }));
    // 推計の幅（range）：o.range = { low: [値], high: [値], label（"推計の幅"） }。下端と上端の間を薄い帯で塗り、右端に幅を書く
    if (R) {
      const idx = o.cats.map((_, i) => i).filter(i => R.low[i] != null && R.high[i] != null);
      if (idx.length > 1) {
        chPath(s, idx.map(i => [xOf(i), yOf(R.high[i])]).concat(idx.slice().reverse().map(i => [xOf(i), yOf(R.low[i])])), { close: true, fill: R.color ?? TINT_STRONG });
        const li = idx[idx.length - 1], ym = (yOf(R.high[li]) + yOf(R.low[li])) / 2;
        T(s, [{ text: `${fmt(R.low[li])}〜${fmt(R.high[li])}`, options: { fontSize: FS.block, bold: true, color: C.ink, breakLine: true } }, { text: R.label ?? "推計の幅", options: { fontSize: FS.body, color: C.soft } }],
          { x: x1 + 0.15, y: ym - 0.3, w: lw, h: 0.6, valign: "middle", wrap: false, lineSpacingMultiple: 1.05 });
      }
    }
    const order = ser.map((_, j) => j).sort((a, b) => (ser[a].main ? 1 : 0) - (ser[b].main ? 1 : 0));
    const mode = o.values || "main", ends = [];
    order.forEach(j => {
      const sr = ser[j], main = !!sr.main, col = sr.color || (main ? C.emph : C.faint), wid = main ? FOCUS_LW : 1.25;
      const pts = sr.values.map((v, i) => (v == null ? null : [xOf(i), yOf(v)]));
      const seg = (a, b) => pts.slice(a, b + 1).filter(Boolean);
      const sEnd = sr.from ?? n - 1;
      if (seg(0, sEnd).length > 1) chPath(s, seg(0, sEnd), { line: chLine(col, wid) });
      if (sr.from != null && seg(sr.from, n - 1).length > 1) chPath(s, seg(sr.from, n - 1), { line: chLine(col, wid, "dash") });
      const last = pts.map((p, i) => (p ? i : -1)).filter(i => i >= 0).pop(), first = pts.findIndex(Boolean);
      pts.forEach((p, i) => {
        if (!p || (n > 24 && i !== first && i !== last)) return;
        const fc = sr.from != null && i > sr.from;
        chDot(s, p[0], p[1], main ? 0.12 : 0.1, main && !fc ? col : C.white, main && !fc ? null : chLine(col, 1.25));
        const show = !isFV(j, i) && (sr.labelAt ? sr.labelAt.includes(i) && i !== last : mode === "none" ? false : mode === "ends" ? (i === first && i !== last) : main && i !== last && (n <= 13 || i === first));
        if (show) T(s, fmt(sr.values[i]), { x: p[0] - 0.5, y: sr.below ? p[1] + 0.1 : p[1] - 0.36, w: 1, h: 0.26, fontSize: FS.body, color: main ? C.ink : C.soft, align: "center", valign: sr.below ? "top" : "bottom", wrap: false });
      });
      if (last != null) ends.push({ sr, main, y: pts[last][1], y0: pts[last][1], x: pts[last][0], v: sr.values[last], noV: isFV(j, last) });
    });
    // 右端の名前と最後の値（重ならないよう縦にずらす）
    chSpread(ends, 0.3, top - 0.1, bottom);
    ends.forEach(e => { if (!e.sr.cagr) return; const need = textW(fmt(e.v), e.main ? FS.head : FS.body) + textW(`　${e.sr.name}　${e.sr.cagr}`, FS.body), room = Math.min(lw + (x1 - e.x) + 0.1, W - M + 0.1 - (e.x + 0.16)); if (need > room) warn(`折れ線の右端のラベルが幅に入らない（${e.sr.name}　${e.sr.cagr}。必要 ${need.toFixed(2)} > ${room.toFixed(2)}。labelW を広げる）`); });
    ends.forEach(e => T(s, [{ text: e.noV ? "" : fmt(e.v), options: { bold: true, fontSize: e.main ? FS.head : FS.body, color: e.main ? C.ink : C.soft } }, { text: `　${e.sr.name}`, options: { bold: e.main, color: e.main ? C.ink : C.soft } }, ...(e.sr.cagr ? [{ text: `　${e.sr.cagr}`, options: { bold: e.main, color: e.main ? C.ink : C.soft } }] : [])],
      { x: e.x + 0.16, y: e.y - 0.15, w: lw + (x1 - e.x) + 0.1, h: 0.3, fontSize: FS.body, valign: "middle", wrap: false }));
    // 主役の値（36pt）：点の上（focusBelow で下）。右端の名前・目標のラベル・図の上端に重ならないかを確かめる
    if (FV && ser[FV[0]] && ser[FV[0]].values[FV[1]] != null) {
      const v = ser[FV[0]].values[FV[1]], vt = fmt(v), tw = textW(vt, FOCUS_FS), bw0 = Math.max(1.6, tw + 0.2), px = xOf(FV[1]), py = yOf(v);
      const by = o.focusBelow ? py + 0.12 : py - 0.62;
      // 最後の点の主役の値は、右端の名前に掛からないよう、点の左側へ寄せる
      const tx = FV[1] === o.cats.length - 1 ? px + 0.1 - tw : px - tw / 2;
      focusValue(s, vt, { x: tx + tw / 2 - bw0 / 2, y: by, w: bw0, align: "center" }, { problem: o.focusProblem });
      const tr = { x: tx, y: by + 0.25 - FOCUS_FS / 144, w: tw, h: FOCUS_FS / 72 };
      const others = [
        ...ends.map(e => ({ x: e.x + 0.16, y: e.y - FS.body / 144, w: (e.noV ? 0 : textW(fmt(e.v), e.main ? FS.head : FS.body)) + textW(`　${e.sr.name}${e.sr.cagr ? "　" + e.sr.cagr : ""}`, FS.body), h: FS.body / 72 })),
        ...(o.targets || []).filter(t => t.label).map(t => { const tw2 = textW(t.label, FS.body), end = (t.at || "start") === "end"; return { x: end ? x1 + 0.05 - tw2 : x0 - 0.25, y: yOf(t.value) - 0.28, w: tw2, h: 0.25 }; }),
      ];
      if (others.some(b => boxHit(tr, b))) warn(`折れ線の主役の値（${vt}）が、右端の名前か目標のラベルに重なる（focusBelow で点の下にするか、o.max・labelW を変える）`);
      if (tr.y < BODY + 0.35 || tr.y + tr.h > bottom + 0.4) warn(`折れ線の主役の値（${vt}）が図の${tr.y < BODY + 0.35 ? "上端" : "下端"}を超える（o.max を上げるか、focusBelow を変える）`);
      if (tr.x + tr.w > W - M + 0.05) warn(`折れ線の主役の値（${vt}）が右の余白にはみ出す`);
    }
    return { xOf, yOf, pt: (j, i) => [xOf(i), yOf(ser[j].values[i])] };
  };

  // 積み上げ面（図形）：総量と内訳の推移で、点が多い（8期以上の）連続した時系列。区分名と最後の値は右端の面の横に直接書く
  //   期が7つ以下なら、棒で区切った stackBars か、帯でつなぐ ribbonBars のほうが値を読みやすい。1系列だけなら lineChart
  //   o.cats, o.series: [{ name, values, color?, main? }]（下から積む順。主役の区分は濃紺、他は区分色）, o.total（"ends"＝最初と最後の合計を上に。既定／"all"／"none"）
  //   o.x0, o.x1, o.top, o.bottom, o.max, o.fmt, o.labelW（1.9）, o.catEvery
  const areaChart = (s, o) => {
    const ser = o.series, n = o.cats.length, fmt = o.fmt || (v => fmtNum(v));
    const lw = o.labelW ?? 1.9, x0 = o.x0 ?? M + 0.5, x1 = o.x1 ?? W - M - lw - 0.2, top = o.top ?? 2.85, bottom = o.bottom ?? BOTTOM - 0.4;
    const tot = i => ser.reduce((a, sr) => a + sr.values[i], 0);
    const max = o.max ?? Math.max(...o.cats.map((_, i) => tot(i))) * 1.1, step = (x1 - x0) / (n - 1);
    const xOf = i => x0 + i * step, yOf = v => bottom - v / max * (bottom - top);
    const cum = o.cats.map((_, i) => { let a = 0; return ser.map(sr => { const lo = a; a += sr.values[i]; return [lo, a]; }); });
    ser.forEach((sr, j) => {
      const up = o.cats.map((_, i) => [xOf(i), yOf(cum[i][j][1])]), lo = o.cats.map((_, i) => [xOf(i), yOf(cum[i][j][0])]).reverse();
      chPath(s, up.concat(lo), { close: true, fill: sr.color || (sr.main ? C.emph : C.bar), line: chLine(C.white, 0.75) });
    });
    hline(s, x0, bottom, x1 - x0, { color: C.gray });
    const every = o.catEvery ?? Math.ceil(n / 12), tm = o.total || "ends";
    o.cats.forEach((c, i) => {
      if (!(i % every) || i === n - 1) T(s, c, { x: xOf(i) - 0.6, y: bottom + 0.1, w: 1.2, h: 0.28, fontSize: FS.body, color: C.soft, align: "center", wrap: false });
      if (tm === "all" || (tm === "ends" && (i === 0 || i === n - 1))) T(s, fmt(tot(i)), { x: xOf(i) - 0.6, y: yOf(tot(i)) - 0.34, w: 1.2, h: 0.28, fontSize: i === n - 1 ? FS.block : FS.body, bold: i === n - 1, color: C.ink, align: "center", valign: "bottom", wrap: false });
    });
    const ends = ser.map((sr, j) => ({ sr, y: (yOf(cum[n - 1][j][0]) + yOf(cum[n - 1][j][1])) / 2 }));
    chSpread(ends, 0.28, top, bottom - 0.1);
    ends.forEach(e => T(s, [{ text: e.sr.name, options: { bold: !!e.sr.main, color: e.sr.main ? C.ink : C.soft } }, { text: `　${fmt(e.sr.values[n - 1])}`, options: { bold: true, color: e.sr.main ? C.ink : C.soft } }],
      { x: x1 + 0.15, y: e.y - 0.14, w: lw, h: 0.28, fontSize: FS.body, valign: "middle", wrap: false }));
    return { xOf, yOf, cum };
  };

  // 積み上げ縦棒（図形）：区分ごとの量と内訳（既定）、または構成比の比較（pct: true で高さを100%にそろえる）。区分名は右端の棒の横に直接書き、凡例は置かない
  //   積み上げは上端の外に区分の値を置けないので、区分の値は区分の中に書く（高さが足りない区分は省く）。合計は棒の上（pct では置かない）
  //   ネイティブの積み上げ（barChart）は凡例が下に付き、区分の中の値の書式を区分ごとに変えられない。期が多く PowerPoint 上で値を差し替える資料だけ barChart
  //   o.cats, o.series: [{ name, values, color?, main? }]（下から積む）, o.pct, o.inner（区分の中に値。既定 true）, o.total（棒の上に合計。既定 true）
  //   o.x0, o.x1, o.top, o.base, o.max, o.bw, o.fmt, o.pctFmt, o.labelW（1.8）
  //   戻り値: { bx(i), seg[i][j]=[上, 下], bw, slot, base }
  const stackBars = (s, o) => {
    const n = o.cats.length, ser = o.series, pct = !!o.pct, lw = o.labelW ?? 1.8;
    const top = o.top ?? 2.85, base = o.base ?? BOTTOM - 0.4;
    const tot = i => ser.reduce((a, sr) => a + sr.values[i], 0);
    const max = pct ? 1 : (o.max ?? Math.max(...o.cats.map((_, i) => tot(i))) * 1.08);
    const { bw, slot, x0 } = colSlot("積み上げ縦棒", { x0: o.x0 ?? M + 0.3, x1: o.x1 ?? W - M - lw, n, bw: o.bw, needW: Math.max(...o.cats.map(c => textW(c, FS.body))) + 0.25, fill: o.x0 != null && o.x1 != null });
    const bx = i => x0 + i * slot + (slot - bw) / 2;
    const fmt = o.fmt || (v => fmtNum(v)), pf = o.pctFmt || (r => Math.round(r * 100) + "%");
    const seg = o.cats.map((_, i) => { let y = base; const t = tot(i); return ser.map(sr => { const v = pct ? sr.values[i] / t : sr.values[i], h = v / max * (base - top), r = [y - h, y]; y -= h; return r; }); });
    o.cats.forEach((c, i) => {
      ser.forEach((sr, j) => {
        const [t, b] = seg[i][j], col = sr.color || (sr.main ? C.emph : C.bar);
        if (b - t < 0.002) return;
        rect(s, bx(i), t, bw, b - t, col, chLine(C.white, 0.75));
        if (o.inner !== false && b - t >= 0.3 && textW(pct ? pf(sr.values[i] / tot(i)) : fmt(sr.values[i]), minFig()) <= bw - 0.04) T(s, pct ? pf(sr.values[i] / tot(i)) : fmt(sr.values[i]), { x: bx(i), y: t, w: bw, h: b - t, fontSize: areaFs(pct ? pf(sr.values[i] / tot(i)) : fmt(sr.values[i]), bw, b - t, 1, minFig()), bold: !!sr.main, color: isDark(col) ? C.white : C.ink, align: "center", valign: "middle", wrap: false });
      });
      if (!pct && o.total !== false) T(s, fmt(tot(i)), { x: bx(i) - 0.3, y: seg[i][ser.length - 1][0] - 0.32, w: bw + 0.6, h: 0.28, fontSize: FS.block, bold: true, color: C.ink, align: "center", valign: "bottom", wrap: false });
      T(s, c, { x: x0 + i * slot, y: base + 0.1, w: slot, h: 0.3, fontSize: FS.body, color: C.soft, align: "center", wrap: false });
    });
    const labs = ser.map((sr, j) => { const [t, b] = seg[n - 1][j]; return { sr, y: (t + b) / 2 }; });
    chSpread(labs, 0.27, top, base - 0.1);
    labs.forEach(({ sr, y }) => T(s, sr.name, { x: bx(n - 1) + bw + 0.15, y: y - 0.14, w: lw, h: 0.28, fontSize: FS.body, color: sr.main ? C.ink : C.soft, bold: !!sr.main, valign: "middle", wrap: false }));
    hline(s, x0, base, n * slot, { color: C.gray });
    return { bx, seg, bw, slot, base };
  };

  // 横の100%積み上げ（図形）：アンケートの回答の構成、項目ごとの構成比の比較。項目名は左、区分名は一番上の棒の上に直接書き、凡例は置かない
  //   o.rows: [{ label, values, strong? }], o.series: [{ name, color?, main? }]（左から並べる順。肯定→否定）, o.right: { head, value: row => 文字 }（右端の列。例「満足の計」）
  //   o.y0（最初の棒の上端 2.95）, o.pctFmt。項目名の欄の幅・棒の太さ・間は部品が決める（o.labelW・o.bh・o.pitch は、別のページと図をそろえるときだけ）
  //   強調: 主役の区分（main、濃紺）か、主役の行（strong、名前と右端の値を太字）のどちらか1つ
  const hStackBars = (s, o) => {
    const rows = o.rows, ser = o.series, rw = o.right ? Math.max(1.0, textW(o.right.head, FS.note) + 0.1, ...o.rows.map(r => textW(o.right.value(r), FS.head) + 0.1)) : 0; // 右端の列は中身の幅
    const lw = labelCol("横の積み上げ", rows.map(r => [r.label, r.strong ? FS.block : FS.body]), o.labelW);
    const x0 = M + lw, x1 = W - M - rw - (o.right ? 0.25 : 0), y0 = o.y0 ?? 2.95;
    const { bh, pitch } = barRows("横の積み上げ", (BOTTOM - y0) / rows.length, o);
    const pf = o.pctFmt || (r => Math.round(r * 100) + "%"), heads = [];
    rows.forEach((r, i) => {
      const y = y0 + i * pitch, t = r.values.reduce((a, b) => a + b, 0); let x = x0;
      T(s, r.label, { x: M, y, w: lw - 0.12, h: bh, fontSize: r.strong ? FS.block : FS.body, bold: !!r.strong, color: r.strong ? C.ink : C.soft, align: "right", valign: "middle", wrap: false });
      ser.forEach((sr, j) => {
        const w = r.values[j] / t * (x1 - x0), col = sr.color || (sr.main ? C.emph : C.bar);
        if (w > 0.002) rect(s, x, y, w, bh, col, chLine(C.white, 0.75));
        if (w >= textW(pf(r.values[j] / t), minFig()) + 0.06) T(s, pf(r.values[j] / t), { x, y, w, h: bh, fontSize: areaFs(pf(r.values[j] / t), w, bh, 1, minFig()), bold: !!sr.main, color: isDark(col) ? C.white : C.ink, align: "center", valign: "middle", wrap: false });
        if (i === 0) heads[j] = { x, w };
        x += w;
      });
      if (o.right) T(s, o.right.value(r), { x: x1 + 0.25, y, w: rw, h: bh, fontSize: r.strong ? FS.head : FS.body, bold: true, color: C.ink, align: "right", valign: "middle", wrap: false });
    });
    // 区分名：一番上の棒の区分の上に。狭い区分の名前が隣と重なるときは、左右にずらす
    const hs = heads.map((h, j) => ({ j, w: textW(ser[j].name, FS.note) + 0.1, c: h.x + h.w / 2 }));
    for (let k = 1; k < hs.length; k++) { const need = hs[k - 1].c + hs[k - 1].w / 2 + 0.12 + hs[k].w / 2; if (hs[k].c < need) hs[k].c = need; }
    const over = hs.length ? hs[hs.length - 1].c + hs[hs.length - 1].w / 2 - x1 : 0;
    if (over > 0) for (let k = hs.length - 1; k >= 0; k--) { const lim = k === hs.length - 1 ? x1 - hs[k].w / 2 : hs[k + 1].c - hs[k + 1].w / 2 - 0.12 - hs[k].w / 2; if (hs[k].c > lim) hs[k].c = lim; }
    hs.forEach(h => T(s, ser[h.j].name, { x: h.c - h.w / 2 - 0.05, y: y0 - 0.34, w: h.w + 0.1, h: 0.3, fontSize: FS.note, color: ser[h.j].main ? C.ink : C.soft, bold: !!ser[h.j].main, align: "center", valign: "bottom", wrap: false }));
    if (o.right) T(s, o.right.head, { x: x1 + 0.25, y: y0 - 0.34, w: rw, h: 0.3, fontSize: FS.note, color: C.gray, align: "right", valign: "bottom", wrap: false });
    const bottom = y0 + (rows.length - 1) * pitch + bh;
    if (bottom > BOTTOM + 0.05) warn(`横の積み上げの下端 ${bottom.toFixed(2)} が本文の下端を超える`);
    return { x0, x1, rowY: i => y0 + i * pitch, bh, bottom };
  };

  // マリメッコ（メッコ、図形）：市場の区分ごとの大きさ（横幅）と、その中の構成比（高さ）を1枚で見せる。プレーヤー名は右端の列の横に直接書く
  //   列の下に区分名と大きさ、セルの中にシェア（狭いセルは省く）。主役のプレーヤーだけ濃紺、他は灰の濃淡か区分色
  //   o.cols: [{ name, size, values: [プレーヤーごとの量] }]（左から大きい順）, o.players: [{ name, color?, main? }]（下から積む順）
  //   o.x0, o.x1, o.top, o.base, o.gap（列の間 0.05）, o.sizeFmt（列の大きさ）, o.pctFmt, o.labelW（1.8）, o.share（true で列の上に市場全体に占める割合）
  //   戻り値: { col(i) => { x, w, seg[j]=[上, 下] } }
  const mekko = (s, o) => {
    const cols = o.cols, pl = o.players, lw = o.labelW ?? 1.8;
    const x0 = o.x0 ?? M + 0.1, x1 = o.x1 ?? W - M - lw, top = o.top ?? (o.share ? 2.95 : 2.75), base = o.base ?? BOTTOM - 0.45, gap = o.gap ?? 0.05;
    const total = cols.reduce((a, c) => a + c.size, 0), uw = (x1 - x0 - gap * (cols.length - 1)) / total;
    const sf = o.sizeFmt || (v => fmtNum(v)), pf = o.pctFmt || (r => Math.round(r * 100) + "%");
    let x = x0; const out = [];
    cols.forEach(c => {
      const w = c.size * uw, t = c.values.reduce((a, b) => a + b, 0), col = { x, w, seg: [] }; let y = base;
      pl.forEach((p, j) => {
        const h = c.values[j] / t * (base - top), fill = p.color || (p.main ? C.emph : C.bar);
        if (h > 0.002) rect(s, x, y - h, w, h, fill, chLine(C.white, 0.75));
        if (w >= textW(pf(c.values[j] / t), minFig()) + 0.08 && h >= 0.3) T(s, pf(c.values[j] / t), { x, y: y - h, w, h, fontSize: areaFs(pf(c.values[j] / t), w, h, 1, minFig()), bold: !!p.main, color: isDark(fill) ? C.white : C.ink, align: "center", valign: "middle", wrap: false });
        col.seg.push([y - h, y]); y -= h;
      });
      if (o.share) T(s, pf(c.size / total), { x: x - 0.1, y: top - 0.3, w: w + 0.2, h: 0.26, fontSize: FS.body, color: C.soft, align: "center", valign: "bottom", wrap: false });
      const lbw = Math.max(w + 0.08, textW(c.name, FS.body) + 0.1, textW(sf(c.size), FS.body) + 0.1); // 細い区分でも、名前と量を折り返さない
      T(s, [{ text: c.name, options: { breakLine: true, color: C.soft } }, { text: sf(c.size), options: { bold: true, color: C.ink } }], { x: x + w / 2 - lbw / 2, y: base + 0.08, w: lbw, h: 0.62, fontSize: FS.body, align: "center", lineSpacingMultiple: 1.1, wrap: false });
      if (textW(c.name, FS.body) > w + 0.08) warn(`メッコの列の名前が列の幅に入らない（${c.name}）。小さい列は「その他」にまとめる`);
      out.push(col); x += w + gap;
    });
    const last = out[out.length - 1], labs = pl.map((p, j) => ({ p, y: (last.seg[j][0] + last.seg[j][1]) / 2 }));
    chSpread(labs, 0.27, top, base - 0.1);
    labs.forEach(({ p, y }) => T(s, p.name, { x: last.x + last.w + 0.15, y: y - 0.14, w: lw - 0.15, h: 0.28, fontSize: FS.body, bold: !!p.main, color: p.main ? C.ink : C.soft, valign: "middle", wrap: false }));
    return { col: i => out[i] };
  };

  // ツリーマップ（図形）：区分が6つ以上ある構成（費目、商品、顧客）を面積で。大きい順に左上から敷き詰め（squarify）、名前・値・割合を区画の中に書く
  //   区分が5つまでなら pieChart、順位を読ませたいなら hBars。小さすぎて名前が入らない区画は「その他」にまとめる（警告が出る）
  //   o.items: [{ name, value, strong?, color? }], o.x, o.y（2.6）, o.w, o.h（本文の下端まで）, o.gap（白い溝 0.05）, o.fmt, o.pctFmt
  //   戻り値: { cell(name) => { x, y, w, h } }
  const treemap = (s, o) => {
    const x = o.x ?? M, y = o.y ?? 2.6, w = o.w ?? BW, h = o.h ?? BOTTOM - y, gap = o.gap ?? 0.05;
    const items = o.items.slice().sort((a, b) => b.value - a.value), total = items.reduce((a, b) => a + b.value, 0);
    const fmt = o.fmt || (v => fmtNum(v)), pf = o.pctFmt || (r => Math.round(r * 100) + "%");
    const worst = (row, side, sc) => { const ar = row.map(r => r.value * sc), sum = ar.reduce((a, b) => a + b, 0); return Math.max(...ar.map(a => Math.max(side * side * a / (sum * sum), sum * sum / (side * side * a)))); };
    const cells = []; let box = { x, y, w, h }, rest = items;
    while (rest.length) {
      const remain = rest.reduce((a, b) => a + b.value, 0), sc = box.w * box.h / remain, side = Math.min(box.w, box.h);
      const row = [rest[0]];
      while (row.length < rest.length && worst([...row, rest[row.length]], side, sc) <= worst(row, side, sc)) row.push(rest[row.length]);
      const thick = row.reduce((a, b) => a + b.value, 0) * sc / side, vert = box.w >= box.h; let off = 0;
      row.forEach(r => { const len = r.value * sc / thick; cells.push(vert ? { it: r, x: box.x, y: box.y + off, w: thick, h: len } : { it: r, x: box.x + off, y: box.y, w: len, h: thick }); off += len; });
      box = vert ? { x: box.x + thick, y: box.y, w: box.w - thick, h: box.h } : { x: box.x, y: box.y + thick, w: box.w, h: box.h - thick };
      rest = rest.slice(row.length);
    }
    const map = {};
    cells.forEach(c => {
      const it = c.it, fill = it.strong ? C.emph : (it.color || C.bar), cx = c.x + gap / 2, cy = c.y + gap / 2, cw = c.w - gap, chh = c.h - gap;
      rect(s, cx, cy, cw, chh, fill);
      map[it.name] = { x: cx, y: cy, w: cw, h: chh };
      const tc = isDark(fill) ? C.white : C.ink, vline = `${fmt(it.value)}　${pf(it.value / total)}`, nameFs = areaFs(`${it.name}\n${vline}`, cw, Math.min(chh, 1.6), 2, it.strong ? FS.head : FS.body);
      const fits = textW(it.name, nameFs) <= cw - 0.2 && chh >= 0.62;
      if (fits) T(s, [{ text: it.name, options: { bold: true, fontSize: nameFs, breakLine: true } }, { text: vline, options: { fontSize: nameFs, bold: !!it.strong } }],
        { x: cx + 0.1, y: cy + 0.08, w: cw - 0.2, h: chh - 0.16, color: tc, valign: "top", lineSpacingMultiple: 1.15 });
      else if (textW(it.name, FS.note) <= cw - 0.12 && chh >= 0.3) T(s, it.name, { x: cx + 0.06, y: cy + 0.04, w: cw - 0.12, h: 0.24, fontSize: FS.note, color: tc, bold: true, wrap: false });
      else warn(`ツリーマップの区画「${it.name}」に名前が入らない。小さい区分は「その他」にまとめる`);
    });
    return { cell: name => map[name] };
  };

  // 円・ドーナツ（図形）：1時点の構成比。区分は5つまで、大きい順に12時から時計回り（「その他」は最後）、主役の区分だけ濃紺、他は灰。名前と値は区分の外に直接書き、凡例を置かない
  //   時系列（年ごとの円）や、2つの円の比較には使わない（stackBars の pct か ribbonBars）。区分が6つ以上なら hBars か treemap
  //   ネイティブの円は、主役のラベルだけ太字にできない・ラベルの位置を PowerPoint が決める・ドーナツの中央に合計を置けないので使わない
  //   o.items: [{ name, value, note?（値の下の1行）, strong?, other?（「その他」。並びの点検から外す） }], o.donut, o.center: { value, label }（ドーナツの中央）
  //   o.cx（W/2）, o.cy（4.6）, o.r（1.95）, o.thickness（ドーナツの太さ。半径に対する比 0.4）, o.fmt（it => 値の文字。既定は割合）, o.colors（主役以外の色。既定 C.bar）, o.labelW（2.9）
  //   戻り値: { cx, cy, r, mid(i)（区分の中央の角度） }
  const pieChart = (s, o) => {
    const items = o.items, tot = items.reduce((a, b) => a + b.value, 0), r = o.r ?? 1.95, cx = o.cx ?? W / 2, cy = o.cy ?? 4.6, lw = o.labelW ?? 2.9;
    if (items.length > 5) warn(`円グラフの区分が${items.length}つある。5つまでにする（小さい区分は「その他」にまとめるか、hBars・treemap にする）`);
    if (items.some((it, i) => i && !it.other && it.value > items[i - 1].value)) warn("円グラフの区分が大きい順に並んでいない");
    // 見る所の区分を strong で強調の色に（2026-11-25）。リードが合計（上位4社の69%など）を言うときは、その合計に入る区分をすべて strong にし、残りは灰
    if (!items.some(it => it.strong)) warn("円グラフの見る所の区分が無い。リードが言う区分（合計を言うなら、その合計に入る区分すべて）に strong を付け、強調の色で塗る。ほかは灰（色を手で全部同じにしない）");
    if (o.colors && items.some(it => it.strong)) warn("円グラフに colors と strong を両方渡している。見る所は strong に任せ、colors は渡さない");
    const pf = o.fmt || (it => Math.round(it.value / tot * 100) + "%");
    let a = 270; const mids = [];
    items.forEach((it, i) => {
      const sw = it.value / tot * 360, fill = it.strong ? C.emph : (o.colors ? o.colors[i] : C.bar);
      s.addShape(o.donut ? SH.BLOCK_ARC : SH.PIE, { x: cx - r, y: cy - r, w: 2 * r, h: 2 * r, angleRange: [a % 360, (a + sw) % 360], ...(o.donut ? { arcThicknessRatio: o.thickness ?? 0.4 } : {}), fill: { color: fill }, line: chLine(C.white, 2) });
      mids.push(a + sw / 2); a += sw;
    });
    // ラベル：区分の中央の角度の外側。左右の列に分け、縦に重ならないようにずらし、細い引出線でつなぐ
    // ラベルの高さは行数から出す（名前 1 行＋割合 1 行＋注の行。注は明示の改行と、幅 lw での折り返しを数える）
    const labH = it => { const st = !!it.strong; return lineH(st ? FS.block : FS.body, 1.05) + lineH(st ? FS_DISPLAY.num : FS.head, 1.05) + (it.note ? lineCount(it.note, FS.body, lw) * lineH(FS.body, 1.05) : 0) + 0.04; };
    const labs = items.map((it, i) => { const m = mids[i] * Math.PI / 180, c = Math.cos(m), sn = Math.sin(m), h = Math.max(it.note ? 0.82 : 0.58, labH(it)); return { it, i, c, sn, right: c >= -0.02, y: cy + sn * (r + 0.2), h, hh: h / 2 + 0.03 }; });
    [true, false].forEach(side => {
      const L = labs.filter(l => l.right === side);
      const hMax = Math.max(0.58, ...L.map(l => l.h));
      const lo = 2.55 + (L.length ? L.reduce((a, l) => (l.y < a.y ? l : a)).h / 2 : 0), hi = BOTTOM - (L.length ? L.reduce((a, l) => (l.y > a.y ? l : a)).h / 2 : 0);
      if (L.every(l => Math.abs(l.h - hMax) < 0.005)) chSpread(L, hMax + 0.06, 2.55 + hMax / 2, BOTTOM - hMax / 2); // 高さがそろうときは今まで通り
      else chSpreadH(L, lo, hi);
      const tot = L.reduce((a, l) => a + l.h + 0.06, 0) - 0.06;
      if (tot > BOTTOM - 2.55 + 0.05) warn(`円グラフの${side ? "右" : "左"}のラベル（注を含めて ${tot.toFixed(2)}）が本文の高さに入らない。注を短くするか、注を主役の区分だけにする`);
      L.forEach(l => { if (l.y - l.h / 2 < 2.5 || l.y + l.h / 2 > BOTTOM + 0.05) warn(`円グラフのラベル（${l.it.name}）が本文の領域からはみ出す`); });
      L.forEach(l => {
        // 引出線：区分の縁 → ラベルと同じ高さで円の外の点 → ラベル（円の中を横切らない）
        const ex = cx + l.c * (r + 0.04), ey = cy + l.sn * (r + 0.04), R2 = r + 0.16, dy = l.y - cy;
        const kx = Math.abs(dy) < R2 ? cx + (side ? 1 : -1) * Math.sqrt(R2 * R2 - dy * dy) : ex;
        const lx = side ? Math.max(cx + r + 0.35, kx + 0.12) : Math.min(cx - r - 0.35, kx - 0.12);
        chPath(s, [[ex, ey], [kx, l.y], [lx, l.y]], { line: chLine(C.gray, 0.75) });
        const st = !!l.it.strong, runs = [
          { text: l.it.name, options: { fontSize: st ? FS.block : FS.body, bold: st, color: st ? C.ink : C.soft, breakLine: true } },
          { text: pf(l.it), options: { fontSize: st ? FS_DISPLAY.num : FS.head, bold: true, color: st ? C.ink : C.soft, breakLine: !!l.it.note } },
        ];
        if (l.it.note) runs.push({ text: l.it.note, options: { fontSize: FS.body, color: C.soft } });
        T(s, runs, { x: side ? lx + 0.08 : lx - 0.08 - lw, y: l.y - l.h / 2, w: lw, h: l.h, align: side ? "left" : "right", valign: "middle", lineSpacingMultiple: 1.05 });
        if (l.it.note) checkWrap(l.it.note, FS.body, lw);
      });
    });
    if (o.donut && o.center) T(s, [{ text: o.center.value, options: { fontSize: FS_DISPLAY.num, bold: true, color: C.ink, breakLine: true } }, { text: o.center.label, options: { fontSize: FS.body, color: C.soft } }],
      { x: cx - r * 0.6, y: cy - 0.42, w: r * 1.2, h: 0.84, align: "center", valign: "middle", lineSpacingMultiple: 1.05 });
    return { cx, cy, r, mid: i => mids[i] };
  };

  // ダンベル（2時点・2条件の差を、2つの点と線で示す図）は、2026-11-25 に部品から削除した（何を言いたいのかが読み取りにくいため）。
  //   2つの条件を比べるときは、対の棒（colBars の series）か、横棒（hBars）で描く
  const dumbbell = () => { throw new Error("dumbbell は削除した（2026-11-25）。2つの条件の比較は、対の棒（colBars の series）か横棒（hBars）で描く"); };

  // ブレット（図形）：KPI の実績と目標を項目ごとに。灰の帯（段階）の上に実績の棒、目標は縦の太線。右に実績・目標・達成率の列
  //   単位の違う KPI を1枚に並べられる（各行は自分の目盛り：0〜max(実績, 目標, 帯)×1.1）。ゲージやメーターの代わりに使う
  //   o.rows: [{ label, value, target, bands?: [段階の境の値, ...]（低い順。1つなら不足／それ以上、2つなら不足／標準／良好）, unit?, inverse?（少ないほど良い指標。達成率を 目標÷実績 にする）, strong?, problem? }]
  //   o.heads（右の列の見出し。既定 ["実績", "目標", "達成率"]）, o.fmt（(v, row) => 文字）, o.y0（2.95）, o.pitch, o.labelW（2.6）
  //   強調: 主役の行は strong（濃紺の棒・太字）、目標に届かない問題の行は problem（赤の値）。1ページでどちらか1種類
  //   戻り値: { rowY(i), x0, x1 }
  const bulletChart = (s, o) => {
    checkOneEm(o.rows, "ブレット");
    const rows = o.rows, lw = labelCol("ブレット", o.rows.map(r => r.label), o.labelW, FS.block), heads = o.heads || ["実績", "目標", "達成率"];
    const fmt = o.fmt || ((v, r) => fmtNum(v) + (r.unit || ""));
    const cw = Math.max(1.15, ...o.rows.flatMap(r => [textW(fmt(r.value, r), FS.head) + 0.2, textW(fmt(r.target, r), FS.body) + 0.2]), ...heads.map(h => textW(h, FS.note) + 0.15)), x0 = M + lw, x1 = W - M - cw * heads.length - 0.35, y0 = o.y0 ?? 2.95;
    const pitch = o.pitch ?? Math.min(0.72, (BOTTOM - y0) / rows.length), bh = Math.min(0.44, pitch * 0.66), rowY = i => y0 + i * pitch;
    heads.forEach((h, k) => T(s, h, { x: x1 + 0.35 + k * cw, y: y0 - 0.34, w: cw, h: 0.26, fontSize: FS.note, color: C.gray, align: "right", valign: "bottom", wrap: false }));
    rows.forEach((r, i) => {
      const y = rowY(i), max = Math.max(r.value, r.target, ...(r.bands || [])) * 1.1, k = (x1 - x0) / max, st = !!r.strong, pb = !!r.problem;
      T(s, r.label, { x: M, y, w: lw - 0.12, h: bh, fontSize: st ? FS.block : FS.body, bold: st, color: st || pb ? C.ink : C.soft, align: "right", valign: "middle", wrap: false });
      const edges = [0, ...(r.bands || []), max];
      // 段階の帯：不足の域が一番濃い。少ないほど良い指標（inverse）は右端が不足の域
      const z = edges.length - 1, pal0 = z === 1 ? [TINT] : z === 2 ? ["DADDE1", TINT] : ["DADDE1", BA_BG.before, TINT], pal = r.inverse ? pal0.slice().reverse() : pal0;
      edges.slice(0, -1).forEach((e, b) => rect(s, x0 + e * k, y, (edges[b + 1] - e) * k, bh, pal[Math.min(b, pal.length - 1)]));
      rect(s, x0, y + bh * 0.3, r.value * k, bh * 0.4, st ? C.ink : C.gray);
      s.addShape(SH.LINE, { x: x0 + r.target * k, y: y + bh * 0.1, w: 0, h: bh * 0.8, line: chLine(C.emph, 2.25) });
      const cells = [fmt(r.value, r), fmt(r.target, r), Math.round((r.inverse ? r.target / r.value : r.value / r.target) * 100) + "%"];
      cells.forEach((t, c) => T(s, t, { x: x1 + 0.35 + c * cw, y, w: cw, h: bh, fontSize: c !== 1 && (st || pb) ? FS.block : FS.body, bold: c !== 1, color: c !== 1 && pb ? C.accent : c === 1 ? C.soft : C.ink, align: "right", valign: "middle", wrap: false }));
    });
    // 目標の線の読み方を一番上の行の目標の上に1回だけ
    const r0 = rows[0], k0 = (x1 - x0) / (Math.max(r0.value, r0.target, ...(r0.bands || [])) * 1.1);
    T(s, "目標", { x: x0 + r0.target * k0 - 0.5, y: y0 - 0.32, w: 1, h: 0.26, fontSize: FS.note, color: C.gray, align: "center", valign: "bottom", wrap: false });
    const bottom = rowY(rows.length - 1) + bh;
    if (bottom > BOTTOM + 0.05) warn(`ブレットの下端 ${bottom.toFixed(2)} が本文の下端を超える`);
    return { rowY, x0, x1 };
  };

  // 増減の分解（ブリッジ、図形）：始点の値から要因ごとの増減を経て終点の値へ（前年→今年の利益、計画→実績の差）。増は上へ、減は下へ浮かせた棒
  //   既存の waterfall は「期ごとの増分を積み、累計を内訳で見せる」図で、減る要因を描けない。増減の要因を分けるときはこちら
  //   o.start: { label, value }, o.steps: [{ label, value（正負）, focus?, problem? }], o.subtotals: [{ after: 要因の番号, label }], o.end: { label }（値は自動で計算）
  //   o.x0, o.x1, o.top, o.base, o.min（既定 0。0 以外にすると始点・終点の棒の下に省略の印）, o.max, o.fmt, o.stepFmt（増減の棒の値の書式関数。既定は fmtNum で正に「+」、桁は o.digits）
  //   色: 始点・小計・終点は濃い灰（C.faint）、増減は灰（C.bar）。リードが名指しした要因だけ focus で濃紺、問題の要因は problem で赤枠（どちらか1種類）
  //   主役の値: steps の1つ（または start・end）に focusValue: true。その値を主役の値 36pt で普通の位置（増は棒の上、減は棒の下）に部品が描く（problem の要因なら赤）
  //     要因の focusValue は focus（強調の色の塗り）も兼ねる。文字が隣の棒・値に重なるなら、重ならない所まで上（減なら下）へずらす。呼ぶ側で値を空にして focusValue を手で置かない
  //   戻り値: { bx(i), bw, yOf(v), items }（i は始点を 0 とする並び）
  const bridge = (s, o) => {
    chartPage();
    if (o.steps.filter(st => st.focusValue).length + (o.start.focusValue ? 1 : 0) + (o.end.focusValue ? 1 : 0) > 1) warn("ブリッジ：主役の値（focusValue）は1つだけにする");
    o = { ...o, steps: o.steps.map(st => (st.focusValue && !st.problem ? { ...st, focus: true } : st)) };
    checkOneEm(o.steps.map(st => ({ ...st, strong: st.focus })), "ブリッジ");
    const fmt = o.fmt || (v => fmtNum(v)), stepFmt = o.stepFmt || (v => fmtNum(v, { plus: true, digits: o.digits })), items = [{ kind: "total", label: o.start.label, value: o.start.value }];
    let run = o.start.value;
    o.steps.forEach((st, i) => {
      items.push({ kind: "step", ...st, from: run, to: run + st.value }); run += st.value;
      (o.subtotals || []).filter(t => t.after === i).forEach(t => items.push({ kind: "total", label: t.label, value: run, sub: true }));
    });
    items.push({ kind: "total", label: o.end.label, value: run, end: true, focusValue: o.end.focusValue });
    if (o.start.focusValue) items[0].focusValue = true;
    const n = items.length, x0 = o.x0 ?? M + 0.1, x1 = o.x1 ?? W - M - 0.1, top = o.top ?? 2.85, base = o.base ?? BOTTOM - 0.45, min = o.min ?? 0;
    const hiV = Math.max(...items.map(it => (it.kind === "total" ? it.value : Math.max(it.from, it.to)))), max = o.max ?? min + (hiV - min) * 1.06;
    const yOf = v => base - (v - min) / (max - min) * (base - top), slot = (x1 - x0) / n, bw = Math.min(0.9, slot * 0.62), bx = i => x0 + i * slot + (slot - bw) / 2;
    // 棒と値の文字の範囲（主役の値の位置を決める）
    const BR = [], LB = [];
    items.forEach((it, i) => {
      const x = bx(i), tot = it.kind === "total", yT = tot ? yOf(it.value) : yOf(Math.max(it.from, it.to)), yB = tot ? base : yOf(Math.min(it.from, it.to)), up = tot || it.value >= 0;
      const t = tot ? fmt(it.value) : stepFmt(it.value), fs = it.focusValue ? FOCUS_FS : tot ? (it.end ? FS.head : FS.block) : (it.focus || it.problem) ? FS.head : FS.body;
      const tw = textW(t, fs), th = fs * 1.2 / 72, gh = fs / 72, off = (th - gh) / 2; // 文字の範囲は字面（大きさ×1.0）
      BR.push({ x, y: yT, w: bw, h: Math.max(yB - yT, 0.02) });
      LB.push({ x: x + bw / 2 - tw / 2, y: (up ? yT - 0.04 - th : yB + 0.04) + off, w: tw, h: gh, off, t, fs, up });
    });
    items.forEach((it, i) => {
      if (!it.focusValue) return;
      const L = LB[i], y0L = L.y;
      if (liftClear(L, [...BR.filter((_, q) => q !== i), ...LB.filter((_, q) => q !== i)], L.up ? -1 : 1) > 0.005 && (L.y < BODY + 0.3 || L.y + L.h > BOTTOM)) warn(`ブリッジの主役の値（${L.t}）を、隣の棒・値に重ならない所までずらしたら、図の外に出た（o.max・o.min を広げる）`);
      if (Math.abs(L.y - y0L) > 0.9) warn(`ブリッジの主役の値（${L.t}）が、隣に重ならないよう棒から ${Math.abs(L.y - y0L).toFixed(2)} 離れた。画像で確かめる`);
    });
    items.forEach((it, i) => {
      const x = bx(i), L = LB[i], fv = !!it.focusValue, bxW = Math.max(bw + 0.8, L.w + 0.2);
      if (it.kind === "total") {
        const y = yOf(it.value);
        rect(s, x, y, bw, base - y, C.faint);
        if (min > 0) [0.16, 0.24].forEach(dy => s.addShape(SH.LINE, { x: x - 0.04, y: base - dy, w: bw + 0.08, h: 0.06, flipV: true, line: chLine(C.white, 1.5) }));
        if (fv) T(s, L.t, { x: x + bw / 2 - bxW / 2, y: L.y + L.h + L.off - 0.62, w: bxW, h: 0.62, fontSize: FOCUS_FS, bold: true, color: C.emph, align: "center", valign: "bottom", wrap: false });
        else T(s, fmt(it.value), { x: x - 0.4, y: y - 0.34, w: bw + 0.8, h: 0.3, fontSize: it.end ? FS.head : FS.block, bold: true, color: C.ink, align: "center", valign: "bottom", wrap: false });
      } else {
        const yT = yOf(Math.max(it.from, it.to)), yB = yOf(Math.min(it.from, it.to)), h = Math.max(yB - yT, 0.02);
        rect(s, x, yT, bw, h, it.focus ? C.emph : C.bar, it.problem ? chLine(C.accent, 2.25) : null);
        const up = it.value >= 0, em = it.focus || it.problem;
        if (fv) T(s, L.t, { x: x + bw / 2 - bxW / 2, y: up ? L.y + L.h + L.off - 0.62 : L.y - L.off, w: bxW, h: 0.62, fontSize: FOCUS_FS, bold: true, color: it.problem ? C.accent : C.emph, align: "center", valign: up ? "bottom" : "top", wrap: false });
        else T(s, stepFmt(it.value), { x: x - 0.4, y: up ? yT - 0.32 : yB + 0.04, w: bw + 0.8, h: 0.28, fontSize: em ? FS.head : FS.body, bold: true, color: it.problem ? C.accent : em ? C.ink : C.soft, align: "center", valign: up ? "bottom" : "top", wrap: false });
      }
      // 次の棒への連結線（今の水準）
      if (i < n - 1) { const lv = it.kind === "total" ? it.value : it.to; hline(s, x + bw, yOf(lv), slot - bw, { color: C.gray }); }
      const tot = it.kind === "total";
      T(s, it.label, { x: x0 + i * slot + 0.02, y: base + 0.08, w: slot - 0.04, h: 0.5, fontSize: FS.body, bold: tot, color: tot ? C.ink : C.soft, align: "center", valign: "top", lineSpacingMultiple: 1.05 });
      checkWrap(it.label, FS.body, slot - 0.04);
      if (lineCount(it.label, FS.body, slot - 0.04) > 2) warn(`ブリッジの要因名が3行になる（${it.label}）`);
    });
    hline(s, x0, base, x1 - x0, { color: C.gray });
    return { bx, bw, yOf, items };
  };

  // 棒と折れ線の複合（図形）：量（棒）と率（折れ線）を同じ期で（売上と利益率、件数と成約率）。2つの縦軸を重ねず、上の段に率の線、下の段に量の棒を置き、どちらも値を直接書く
  //   ネイティブの複合グラフは左右に2つの縦軸ができ、どちらの軸で読むか迷わせる。線と棒の交わりに意味があるように見えてしまうので使わない
  //   o.cats, o.bars: { name, values, focus?: [番号], fmt? }, o.line: { name, values, fmt?, min?, max? }, o.lineH（線の段の高さ 1.0）
  //   o.x0, o.x1, o.top（2.75）, o.base（6.05）, o.labelW（左の段の名前の幅 1.5）, o.max（棒の上限）
  //   強調: 棒の focus（濃紺）だけ。線は soft の色で、最後の値だけ太字
  //   戻り値: { cx(i), barTop(i), lineY(i) }
  const comboChart = (s, o) => {
    const n = o.cats.length, lw = o.labelW ?? 1.5, x0 = o.x0 ?? M + lw, x1 = o.x1 ?? W - M, top = o.top ?? 2.75, base = o.base ?? BOTTOM - 0.4, lh = o.lineH ?? 1.0;
    const slot = (x1 - x0) / n, cx = i => x0 + slot * (i + 0.5), bw = Math.min(0.7, slot * 0.56);
    const L = o.line, lmin = L.min ?? Math.min(...L.values), lmax = L.max ?? Math.max(...L.values), lTop = top + 0.32, lBot = top + lh;
    const lineY = i => lBot - (L.values[i] - lmin) / (lmax - lmin || 1) * (lBot - lTop), lf = L.fmt || (v => fmtNum(v));
    const bTop = lBot + 0.75, B = o.bars, bmax = o.max ?? Math.max(...B.values) * 1.05, bf = B.fmt || (v => fmtNum(v)), focus = new Set(B.focus || []);
    const barTop = i => base - B.values[i] / bmax * (base - bTop);
    T(s, L.name, { x: M, y: (lTop + lBot) / 2 - 0.2, w: lw - 0.15, h: 0.4, fontSize: FS.body, bold: true, color: C.soft, valign: "middle" });
    T(s, B.name, { x: M, y: (bTop + base) / 2 - 0.2, w: lw - 0.15, h: 0.4, fontSize: FS.body, bold: true, color: C.soft, valign: "middle" });
    chPath(s, L.values.map((_, i) => [cx(i), lineY(i)]), { line: chLine(C.soft, 2) });
    L.values.forEach((v, i) => {
      chDot(s, cx(i), lineY(i), 0.1, C.soft);
      T(s, lf(v), { x: cx(i) - 0.5, y: lineY(i) - 0.34, w: 1, h: 0.26, fontSize: FS.body, bold: i === n - 1, color: C.soft, align: "center", valign: "bottom", wrap: false });
    });
    B.values.forEach((v, i) => {
      const f = focus.has(i), y = barTop(i);
      rect(s, cx(i) - bw / 2, y, bw, base - y, f ? C.emph : C.bar);
      T(s, bf(v), { x: cx(i) - 0.6, y: y - 0.32, w: 1.2, h: 0.28, fontSize: f ? FS.head : FS.body, bold: f, color: f ? C.ink : C.soft, align: "center", valign: "bottom", wrap: false });
      T(s, o.cats[i], { x: cx(i) - slot / 2, y: base + 0.1, w: slot, h: 0.28, fontSize: FS.body, color: C.soft, align: "center", wrap: false });
    });
    hline(s, x0, base, x1 - x0, { color: C.gray });
    return { cx, barTop, lineY };
  };

  // スロープ（図形）：2時点の値（順位）の変化を、項目ごとの1本の線で。左右の端に名前と値を直接書く。主役の項目だけ濃紺の太線、他は灰の細線
  //   3時点以上なら lineChart。項目が10を超えると線が絡むので、主役と比べる相手だけに絞る
  //   o.left, o.right（時点の名前）, o.rows: [{ name, a, b, strong?, focusValue?, problem? }], o.min, o.max, o.fmt, o.xa, o.xb（線の両端）, o.top（2.95）, o.bottom（6.3）
  //     focusValue: true＝主役の行（strong と同じ強調）。右の端の値（b）を主役の値 36pt で、普通の位置（点の右、名前の前）に部品が描く（problem も付けると赤）
  //     ほかの行の名前は、主役の値の高さの分だけ上下へよける。呼ぶ側で値を空にして focusValue を手で置かない
  //   戻り値: { yOf(v), xa, xb }
  const slopeChart = (s, o) => {
    const rows = o.rows, fmt = o.fmt || (v => fmtNum(v)), xa = o.xa ?? M + 3.9, xb = o.xb ?? W - M - 3.9, top = o.top ?? 3.05, bottom = o.bottom ?? BOTTOM - 0.3;
    const vals = rows.flatMap(r => [r.a, r.b]), min = o.min ?? Math.min(...vals), max = o.max ?? Math.max(...vals);
    const yOf = v => bottom - (v - min) / (max - min || 1) * (bottom - top);
    [[o.left, xa], [o.right, xb]].forEach(([t, x]) => T(s, t, { x: x - 1.2, y: top - 0.55, w: 2.4, h: 0.3, fontSize: FS.block, bold: true, color: C.ink, align: "center", valign: "bottom", wrap: false }));
    [xa, xb].forEach(x => s.addShape(SH.LINE, { x, y: top - 0.22, w: 0, h: bottom - top + 0.37, line: chLine(C.line, 0.75) }));
    if (rows.filter(r => r.focusValue).length > 1) warn("スロープ：主役の値（focusValue）は1行だけにする");
    const stOf = r => !!(r.strong || r.focusValue);
    rows.map((r, i) => i).sort((p, q) => (stOf(rows[p]) ? 1 : 0) - (stOf(rows[q]) ? 1 : 0)).forEach(i => {
      const r = rows[i], st = stOf(r), col = st ? C.emph : C.bar;
      segLine(s, xa, yOf(r.a), xb, yOf(r.b), chLine(col, st ? 3 : 1.5));
      chDot(s, xa, yOf(r.a), st ? 0.13 : 0.1, st ? C.ink : C.faint); chDot(s, xb, yOf(r.b), st ? 0.13 : 0.1, st ? C.ink : C.faint);
    });
    [["a", true], ["b", false]].forEach(([key, left]) => {
      const L = rows.map(r => ({ r, y: yOf(r[key]), hh: !left && r.focusValue ? FOCUS_FS / 144 + 0.02 : 0.145 }));
      chSpreadH(L, top - 0.1, bottom + 0.1);
      L.forEach(({ r, y }) => {
        const st = stOf(r), fv = !left && r.focusValue;
        const val = { text: fmt(r[key]), options: { bold: true, color: fv ? (r.problem ? C.accent : C.emph) : st ? C.ink : C.soft, fontSize: fv ? FOCUS_FS : st ? FS.head : FS.body } }, nm = { text: r.name, options: { bold: st, color: st ? C.ink : C.soft, fontSize: st ? FS.block : FS.body } }, sp = fv ? { text: "　", options: { fontSize: FS.body } } : { text: "　" };
        T(s, left ? [nm, sp, val] : [val, sp, nm], { x: left ? xa - 0.15 - 3.3 : xb + 0.15, y: y - (fv ? 0.31 : 0.15), w: 3.3, h: fv ? 0.62 : 0.3, align: left ? "right" : "left", valign: "middle", wrap: false });
        if (fv && textW(fmt(r[key]), FOCUS_FS) + textW(`　${r.name}`, FS.block) > W - M + 0.1 - (xb + 0.15)) warn(`スロープ：主役の値と名前（${fmt(r[key])}　${r.name}）が右の余白に入らない（o.xb を左へ）`);
      });
    });
    return { yOf, xa, xb };
  };

  // 散布図・バブル（図形）：2つの量の関係（成長率×シェア、費用×効果）。点に名前を直接書く。主役の点だけ濃紺、他は灰。size を渡すとバブル（面積が量に比例）
  //   軸は L 字の線と、最小・中ほど・最大の目盛りの値だけ（目盛り線は引かない）。平均などの基準は refX / refY の破線で、象限の名前は quads で角に
  //   ネイティブの散布図は点ごとに名前を書けず（凡例か番号になる）、主役の1点だけの色も変えられないので使わない
  //   o.points: [{ name, x, y, size?, strong?, label?（false で名前を省く） }], o.xName, o.yName（軸の名前と単位）, o.xFmt, o.yFmt
  //   o.xMin, o.xMax, o.yMin, o.yMax（既定 0〜切りのよい上限）, o.ticks（目盛りの数 3）, o.refX / o.refY: { value, label }, o.quads: [右上, 左上, 左下, 右下]
  //   o.area: { x, y, w, h }（点を置く範囲。既定 本文いっぱい）, o.maxD（バブルの最大の直径 0.95）, o.sizeLabel（バブルの大きさの意味。図の見出しに書くなら省く）
  //   戻り値: { px(v), py(v), pt(i) => [x, y] }
  const scatter = (s, o) => {
    chartPage();
    const P = o.points, ar = o.area || { x: M + 0.95, y: 2.95, w: BW - 1.6, h: 3.3 }, xf = o.xFmt || (v => fmtNum(v)), yf = o.yFmt || (v => fmtNum(v));
    const xMin = o.xMin ?? 0, yMin = o.yMin ?? 0, xMax = o.xMax ?? chNice(Math.max(...P.map(p => p.x)) * 1.05), yMax = o.yMax ?? chNice(Math.max(...P.map(p => p.y)) * 1.05);
    const px = v => ar.x + (v - xMin) / (xMax - xMin) * ar.w, py = v => ar.y + ar.h - (v - yMin) / (yMax - yMin) * ar.h, bot = ar.y + ar.h;
    const nt = o.ticks ?? 3, bubble = P.some(p => p.size != null), smax = Math.max(...P.map(p => p.size || 0)), maxD = o.maxD ?? 0.95;
    // 軸（L 字）と目盛りの値
    s.addShape(SH.LINE, { x: ar.x, y: ar.y - 0.1, w: 0, h: ar.h + 0.1, line: chLine(C.gray, 0.75) });
    hline(s, ar.x, bot, ar.w + 0.1, { color: C.gray });
    for (let k = 0; k < nt; k++) {
      const vx = xMin + (xMax - xMin) * k / (nt - 1), vy = yMin + (yMax - yMin) * k / (nt - 1);
      T(s, xf(vx), { x: px(vx) - 0.6, y: bot + 0.06, w: 1.2, h: 0.24, fontSize: FS.note, color: C.gray, align: "center", wrap: false });
      T(s, yf(vy), { x: ar.x - 0.95, y: py(vy) - 0.12, w: 0.85, h: 0.24, fontSize: FS.note, color: C.gray, align: "right", valign: "middle", wrap: false });
    }
    T(s, o.yName, { x: ar.x - 0.95, y: ar.y - 0.42, w: 6, h: 0.26, fontSize: FS.note, color: C.gray, valign: "bottom", wrap: false });
    T(s, o.xName, { x: ar.x + ar.w - 6 + 0.1, y: bot + 0.3, w: 6, h: 0.26, fontSize: FS.note, color: C.gray, align: "right", wrap: false });
    (o.quads || []).forEach((q, k) => { if (!q) return; const right = k === 0 || k === 3, upper = k < 2; T(s, q, { x: right ? ar.x + ar.w - 3.1 : ar.x + 0.12, y: upper ? ar.y - 0.05 : bot - 0.32, w: 3, h: 0.26, fontSize: FS.body, bold: true, color: C.gray, align: right ? "right" : "left", valign: "middle", wrap: false }); });
    if (o.refX) { refLine(s, { x: px(o.refX.value), y0: ar.y, y1: bot, kind: "avg" }); T(s, o.refX.label, { x: px(o.refX.value) + 0.06, y: ar.y - 0.05, w: 2.6, h: 0.26, fontSize: FS.note, color: C.gray, valign: "middle", wrap: false }); }
    if (o.refY) { refLine(s, { y: py(o.refY.value), x0: ar.x, x1: ar.x + ar.w, kind: "avg" }); T(s, o.refY.label, { x: ar.x + ar.w - 2.6, y: py(o.refY.value) - 0.27, w: 2.6, h: 0.25, fontSize: FS.note, color: C.gray, align: "right", valign: "bottom", wrap: false }); }
    const dOf = p => (bubble ? Math.max(0.12, maxD * Math.sqrt((p.size || 0) / smax)) : p.strong ? 0.2 : 0.16);
    const order = P.map((_, i) => i).sort((a, b) => (P[a].strong ? 1 : 0) - (P[b].strong ? 1 : 0) || (bubble ? (P[b].size || 0) - (P[a].size || 0) : 0));
    order.forEach(i => { const p = P[i]; chDot(s, px(p.x), py(p.y), dOf(p), p.strong ? C.emph : bubble ? C.bar : C.faint, bubble ? chLine(C.white, 1) : null); });
    // 名前：右→左→上→下の順に、他の名前・点と重ならない位置を探す
    const boxes = P.map(p => { const d = dOf(p); return { x: px(p.x) - d / 2, y: py(p.y) - d / 2, w: d, h: d }; }), hit = (a, b) => a.x < b.x + b.w && b.x < a.x + a.w && a.y < b.y + b.h && b.y < a.y + a.h;
    const placed = [];
    // 基準の線と象限の名前も、点の名前が避ける対象にする
    if (o.refX) placed.push({ x: px(o.refX.value) + 0.06, y: ar.y - 0.05, w: textW(o.refX.label, FS.note) + 0.1, h: 0.28 });
    if (o.refY) { const w = textW(o.refY.label, FS.note) + 0.1; placed.push({ x: ar.x + ar.w - w, y: py(o.refY.value) - 0.3, w, h: 0.28 }); }
    (o.quads || []).forEach((q, k) => { if (!q) return; const w = textW(q, FS.body) + 0.1, right = k === 0 || k === 3; placed.push({ x: right ? ar.x + ar.w - w : ar.x + 0.12, y: k < 2 ? ar.y - 0.05 : bot - 0.32, w, h: 0.28 }); });
    // 点は円なので、重なりは四角でなく円で見る（大きい円の角の外には名前を置ける）
    const onDot = (q, b) => Math.hypot(Math.max(b.x - (q.x + q.w / 2), 0, q.x + q.w / 2 - b.x - b.w), Math.max(b.y - (q.y + q.h / 2), 0, q.y + q.h / 2 - b.y - b.h)) < q.w / 2;
    order.slice().reverse().forEach(i => {
      const p = P[i]; if (p.label === false) return;
      const st = !!p.strong, fs = st ? FS.block : FS.body, tw = textW(p.name, fs) + 0.06, th = lineH(fs, 1.0), d = dOf(p), cx = px(p.x), cy = py(p.y); // 名前の箱の高さは文字の大きさから
      const cand = [{ x: cx + d / 2 + 0.06, y: cy - th / 2, a: "left" }, { x: cx - d / 2 - 0.06 - tw, y: cy - th / 2, a: "right" }, { x: cx - tw / 2, y: cy - d / 2 - th - 0.02, a: "center" }, { x: cx - tw / 2, y: cy + d / 2 + 0.02, a: "center" }];
      // 4方向に入らなければ、斜め（右上・右下・左上・左下）
      [-1, 1].forEach(v => { const y = cy + v * d * 0.36 - (v < 0 ? th : 0); cand.push({ x: cx + d * 0.36 + 0.03, y, a: "left" }, { x: cx - d * 0.36 - 0.03 - tw, y, a: "right" }); });
      const ok = c => { const b = { x: c.x, y: c.y, w: tw, h: th }; return b.x >= ar.x && b.x + b.w <= ar.x + ar.w + 0.6 && b.y + b.h <= bot + 0.01 && !placed.some(q => hit(q, b)) && !boxes.some((q, k) => k !== i && onDot(q, b)); };
      const found = cand.find(ok), c = found || cand[0];
      if (!found) warn(`散布図：名前「${p.name}」を置く場所がなく、点や軸に重なる。軸の範囲（xMin・xMax）を広げるか、名前を縮める`);
      placed.push({ x: c.x, y: c.y, w: tw, h: th });
      T(s, p.name, { x: c.x, y: c.y, w: tw, h: th, fontSize: fs, bold: st, color: st ? C.ink : C.soft, align: c.a, valign: "middle", wrap: false });
    });
    return { px, py, pt: i => [px(P[i].x), py(P[i].y)] };
  };

  // ヒートマップ（図形）：行×列の値の濃淡（地域×月、商品×顧客層の利用率）。値はセルの中に書き、セルの間は白い溝（罫線を使わない）。濃いほど大きい
  //   1色（tint→ink）の濃淡だけで描き、色相を増やさない。濃いセルの文字は白
  //   o.rows: [行名], o.cols: [列名], o.values: [[...]], o.min, o.max（既定はデータから）, o.fmt, o.mark: [[行, 列], ...]（問題のセルに赤枠。o.markColor で ink 枠）
  //   o.x0, o.y0（セルの上端 2.95）, o.labelW（1.9）, o.cellH（既定は本文の下端までを等分、最大 0.52）, o.total: { head, value: 行の番号 => 文字 }（右端の列）
  //   o.focusValue: [行, 列]（主役のセル。その値を主役の値 36pt でセルの中に部品が描く。濃いセルは白、薄いセルは強調の色。o.focusProblem で薄いセルの値を赤に。問題のセルは mark の赤枠でも示す）
  //     36pt の値がセルの幅より広いと警告（列を減らすか、labelW を狭める）。呼ぶ側で値を空にして focusValue を手で置かない
  //   戻り値: { cell(r, c) => { x, y, w, h } }
  const heatmap = (s, o) => {
    const nr = o.rows.length, nc = o.cols.length, g = 0.04, fmt = o.fmt || (v => fmtNum(v));
    const lw = o.labelW ?? Math.max(1.2, ...o.rows.map(t => textW(t, FS.body) + 0.3)), tw = o.total ? Math.max(1.1, textW(o.total.head, FS.note) - 0.1) : 0;
    const x0 = (o.x0 ?? M) + lw, x1 = W - M - tw - (o.total ? 0.15 : 0), y0 = o.y0 ?? 2.95, ch = o.cellH ?? Math.min(0.52, (BOTTOM - y0) / nr), cw = (x1 - x0) / nc;
    const all = o.values.flat(), min = o.min ?? Math.min(...all), max = o.max ?? Math.max(...all);
    const lo = [0xF4, 0xF6, 0xF9], hi = [0x00, 0x2B, 0x5C];
    const colOf = v => { const t = Math.max(0, Math.min(1, (v - min) / (max - min || 1))); return lo.map((c, k) => Math.round(c + (hi[k] - c) * t).toString(16).padStart(2, "0")).join("").toUpperCase(); };
    const cell = (r, c) => ({ x: x0 + c * cw + g / 2, y: y0 + r * ch + g / 2, w: cw - g, h: ch - g });
    o.cols.forEach((c, j) => T(s, c, { x: x0 + j * cw, y: y0 - 0.32, w: cw, h: 0.26, fontSize: FS.body, color: C.soft, align: "center", valign: "bottom", wrap: false }));
    if (o.total) T(s, o.total.head, { x: x1 + 0.15 - 0.4, y: y0 - 0.32, w: tw + 0.4, h: 0.26, fontSize: FS.note, color: C.gray, align: "right", valign: "bottom", wrap: false });
    o.rows.forEach((rn, i) => {
      T(s, rn, { x: o.x0 ?? M, y: y0 + i * ch, w: lw - 0.15, h: ch, fontSize: FS.body, color: C.soft, valign: "middle", wrap: false });
      o.values[i].forEach((v, j) => {
        const b = cell(i, j), f = colOf(v), fv = o.focusValue && o.focusValue[0] === i && o.focusValue[1] === j;
        rect(s, b.x, b.y, b.w, b.h, f);
        if (fv) {
          // セルに入る大きさ（28pt より下げない）
          const fs = Math.max(28, Math.min(FOCUS_FS, Math.floor((b.w - 0.04) * 72 / em(fmt(v))))), fw = textW(fmt(v), fs);
          focusValue(s, fmt(v), { x: b.x + b.w / 2 - Math.max(b.w, fw + 0.1) / 2, y: b.y + b.h / 2 - 0.25, w: Math.max(b.w, fw + 0.1), align: "center", fontSize: fs }, { onFill: isDark(f), problem: o.focusProblem });
          if (fw > b.w) warn(`ヒートマップ：主役の値（${fmt(v)}、${fs}pt で幅 ${fw.toFixed(2)}）がセルの幅 ${b.w.toFixed(2)} より広い（列を減らすか、labelW・total を狭める）`);
        } else T(s, fmt(v), { ...b, fontSize: FS.body, color: isDark(f) ? C.white : C.ink, align: "center", valign: "middle", wrap: false });
      });
      if (o.total) T(s, o.total.value(i), { x: x1 + 0.15, y: y0 + i * ch, w: tw, h: ch, fontSize: FS.body, bold: true, color: C.ink, align: "right", valign: "middle", wrap: false });
    });
    (o.mark || []).forEach(([r, c]) => { const b = cell(r, c); rect(s, b.x - 0.01, b.y - 0.01, b.w + 0.02, b.h + 0.02, null, chLine(o.markColor || C.accent, 2.25)); });
    if (y0 + nr * ch > BOTTOM + 0.05) warn(`ヒートマップの下端が本文の下端を超える`);
    return { cell };
  };

  // レーダー（図形）：5〜8観点の評価の形を2〜3者で比べる（能力の評価、製品の比較）。外周の多角形（満点）と軸の線だけを引き、内側の目盛りの輪は引かない
  //   主役の系列は濃紺の線と薄い面、他は灰の線と白抜きの点。観点の名前の下に、各系列の値を主役から順に書く（目盛りの代わり）
  //   観点が4つ以下なら colBars（対の棒）、順位を言うなら hBars。軸の順で形が変わるので、比べるページで軸の順をそろえる
  //   o.axes: [観点], o.series: [{ name, values, main? }], o.max（満点）, o.cx, o.cy（4.5）, o.r（1.75）, o.fmt
  //   系列名は figHead の凡例に置く（線の横に置くと観点の名前と重なるため）
  //   戻り値: { pt(i, v) => [x, y] }
  const radarChart = (s, o) => {
    const n = o.axes.length, r = o.r ?? 1.75, cx = o.cx ?? W / 2, cy = o.cy ?? 4.5, max = o.max ?? Math.max(...o.series.flatMap(x => x.values)), fmt = o.fmt || (v => String(v));
    if (n < 5) warn(`レーダーの観点が${n}つ。4つ以下なら対の棒（colBars）にする`);
    const ang = i => -Math.PI / 2 + i * 2 * Math.PI / n, pt = (i, v) => [cx + Math.cos(ang(i)) * r * v / max, cy + Math.sin(ang(i)) * r * v / max];
    chPath(s, o.axes.map((_, i) => pt(i, max)), { close: true, line: chLine(C.line, 1) });
    o.axes.forEach((_, i) => { const [x, y] = pt(i, max); segLine(s, cx, cy, x, y, chLine(C.line, 0.75)); });
    const sorted = o.series.slice().sort((a, b) => (a.main ? 1 : 0) - (b.main ? 1 : 0));
    sorted.forEach(sr => {
      const pts = sr.values.map((v, i) => pt(i, v)), col = sr.color || (sr.main ? C.emph : C.faint);
      chPath(s, pts, { close: true, fill: sr.main ? col : null, transparency: 86, line: chLine(col, sr.main ? 2.25 : 1.5) });
      pts.forEach(([x, y]) => chDot(s, x, y, sr.main ? 0.12 : 0.1, sr.main ? col : C.white, sr.main ? null : chLine(col, 1.25)));
    });
    const main = o.series.find(x => x.main) || o.series[0], others = o.series.filter(x => x !== main);
    o.axes.forEach((a, i) => {
      const c = Math.cos(ang(i)), sn = Math.sin(ang(i)), lx = cx + c * (r + 0.2), ly = cy + sn * (r + 0.16), w = 2.6, h = lineH(FS.body, 1.05) + lineH(FS.head, 1.05); // 観点の名前と値の2行
      const align = Math.abs(c) < 0.25 ? "center" : c > 0 ? "left" : "right";
      const x = align === "center" ? lx - w / 2 : align === "left" ? lx : lx - w, y = sn < -0.25 ? ly - h : sn > 0.25 ? ly : ly - h / 2;
      const runs = [{ text: a, options: { fontSize: FS.body, color: C.soft, breakLine: true } }, { text: fmt(main.values[i]), options: { fontSize: FS.head, bold: true, color: C.ink } }];
      others.forEach(sr => runs.push({ text: `／${fmt(sr.values[i])}`, options: { fontSize: FS.body, color: C.soft } }));
      T(s, runs, { x, y, w, h, align, valign: sn < -0.25 ? "bottom" : sn > 0.25 ? "top" : "middle", lineSpacingMultiple: 1.05 });
    });
    return { pt };
  };

  // スモールマルチプル（図形）：同じ形の小さな図を格子に並べ、項目ごとの推移を比べる（地域別・商品別・店舗別）。縦の目盛りは全パネルで共通、各パネルに最初と最後の値
  //   1枚の折れ線に5本以上の線を重ねると絡むときに使う。主役のパネルだけ濃紺の線と太字の名前、他は灰
  //   o.panels: [{ name, values, strong? }], o.cats（各パネルの下に最初と最後の区分名）, o.cols（4）, o.kind（"line" | "bar"）, o.max, o.min（共通の目盛り。既定 0〜データの最大）, o.index（true で各パネルの最初の値を100にそろえて描く。大きさの違う項目の伸びを比べるとき。書く値は元の値）, o.fmt
  //   o.change（パネルの右上に置く変化。p => 文字。既定は最初→最後の増減率）, o.x0, o.y0（2.55）, o.w, o.h（本文の下端まで）, o.gapX（0.45）, o.gapY（0.3）
  //   戻り値: { panel(i) => { x, y, w, h } }
  const smallMultiples = (s, o) => {
    chartPage();
    const P = o.panels, nc = o.cols ?? 4, nr = Math.ceil(P.length / nc), x0 = o.x0 ?? M, y0 = o.y0 ?? 2.55, w = o.w ?? BW, h = o.h ?? BOTTOM - y0, gx = o.gapX ?? 0.45, gy = o.gapY ?? 0.3;
    const pv = p => (o.index ? p.values.map(v => v / p.values[0] * 100) : p.values), allV = P.flatMap(pv), lo = Math.min(...allV), hi = Math.max(...allV);
    const pw = (w - gx * (nc - 1)) / nc, ph = (h - gy * (nr - 1)) / nr, fmt = o.fmt || (v => fmtNum(v));
    const min = o.min ?? (o.index ? lo - (hi - lo) * 0.14 : 0), max = o.max ?? (o.index ? hi + (hi - lo) * 0.05 : hi * 1.05);
    const chg = o.change || (p => fmtNum(Math.round((p.values[p.values.length - 1] / p.values[0] - 1) * 100), { plus: true, unit: "%" }));
    const out = [];
    P.forEach((p, i) => {
      const px = x0 + (i % nc) * (pw + gx), py = y0 + Math.floor(i / nc) * (ph + gy), st = !!p.strong, col = st ? C.emph : C.faint, n = p.values.length;
      out.push({ x: px, y: py, w: pw, h: ph });
      T(s, p.name, { x: px, y: py, w: pw * 0.62, h: 0.3, fontSize: FS.block, bold: true, color: st ? C.ink : C.soft, valign: "middle", wrap: false });
      T(s, chg(p), { x: px + pw * 0.5, y: py, w: pw * 0.5, h: 0.3, fontSize: st ? FS.head : FS.block, bold: true, color: st ? C.ink : C.soft, align: "right", valign: "middle", wrap: false });
      const vt = py + 0.62, vb = py + ph - 0.3, vx0 = px + 0.55, vx1 = px + pw - 0.55, yv = pv(p), yK = k => vb - (yv[k] - min) / (max - min) * (vb - vt);
      if (o.kind === "bar") {
        const sl = (vx1 - vx0) / n, bw = sl * 0.64;
        p.values.forEach((v, k) => rect(s, vx0 + k * sl + (sl - bw) / 2, yK(k), bw, vb - yK(k), st ? C.emph : C.bar));
      } else {
        const stp = (vx1 - vx0) / (n - 1);
        chPath(s, p.values.map((v, k) => [vx0 + k * stp, yK(k)]), { line: chLine(col, st ? FOCUS_LW : 1.25) });
        [0, n - 1].forEach(k => chDot(s, vx0 + k * stp, yK(k), 0.09, col));
      }
      hline(s, px, vb, pw, { color: C.line });
      const fx = o.kind === "bar" ? vx0 + (vx1 - vx0) / n / 2 : vx0, lx = o.kind === "bar" ? vx1 - (vx1 - vx0) / n / 2 : vx1;
      T(s, fmt(p.values[0]), { x: px - 0.05, y: yK(0) - 0.13, w: fx - px - 0.02, h: 0.26, fontSize: FS.body, color: C.soft, align: "right", valign: "middle", wrap: false });
      T(s, fmt(p.values[n - 1]), { x: lx + 0.08, y: yK(n - 1) - 0.13, w: px + pw - lx, h: 0.26, fontSize: FS.body, bold: true, color: st ? C.ink : C.soft, align: "left", valign: "middle", wrap: false });
      if (o.cats) { T(s, o.cats[0], { x: fx - 0.6, y: vb + 0.04, w: 1.2, h: 0.22, fontSize: FS.note, color: C.gray, align: "center", wrap: false }); T(s, o.cats[o.cats.length - 1], { x: lx - 0.6, y: vb + 0.04, w: 1.2, h: 0.22, fontSize: FS.note, color: C.gray, align: "center", wrap: false }); }
    });
    return { panel: i => out[i] };
  };

  // KPI の数値タイル（図形）：主要な指標を数字で並べる（経営の概況、月次の報告の1枚目）。1枚に名前・値（22pt）・前期との差・小さな推移の線（スパークライン）
  //   タイルは枠なしの薄い塗り。スパークラインは灰の線で、最後の点だけ濃紺。差が悪い向きのタイルは bad で差の文字を赤に（問題だけ赤）
  //   o.tiles: [{ label, value, unit?, delta?（文字。例「前年比 +4.2%」）, bad?, sub?（目標・注記の1行）, spark?: [値], strong? }]
  //   o.cols（既定 4）, o.x0, o.y0（2.6）, o.w, o.h（本文の下端まで）, o.gap（0.15）
  //   強調: 主役のタイルは strong（濃い塗り＋ink 枠）。bad の赤字と同じページで重ねない
  //   戻り値: { tile(i) => { x, y, w, h } }
  const kpiTiles = (s, o) => {
    warn("KPI のタイル（数値のカードの並び）は使わない。言いたい指標を1つ選び、推移か目標との差の図にする");
    const ts = o.tiles, nc = o.cols ?? 4, nr = Math.ceil(ts.length / nc), x0 = o.x0 ?? M, y0 = o.y0 ?? 2.6, w = o.w ?? BW, h = o.h ?? BOTTOM - y0, g = o.gap ?? 0.15;
    if (ts.some(t => t.strong) && ts.some(t => t.bad)) warn("KPI のタイルに strong と bad が混ざる。強調は1ページ1種類にする");
    const tw = (w - g * (nc - 1)) / nc, th = (h - g * (nr - 1)) / nr, out = [];
    ts.forEach((t, i) => {
      const x = x0 + (i % nc) * (tw + g), y = y0 + Math.floor(i / nc) * (th + g), st = !!t.strong, px = 0.2;
      rect(s, x, y, tw, th, st ? TINT_STRONG : TINT, st ? chLine(C.emph, 2) : null);
      out.push({ x, y, w: tw, h: th });
      T(s, t.label, { x: x + px, y: y + 0.16, w: tw - 2 * px, h: 0.3, fontSize: FS.block, bold: true, color: C.ink, valign: "middle", wrap: false });
      T(s, [{ text: t.value, options: { fontSize: FS_DISPLAY.num + (st ? 4 : 0), bold: true, color: C.ink } }, ...(t.unit ? [{ text: ` ${t.unit}`, options: { fontSize: FS.block, color: C.soft } }] : [])], { x: x + px, y: y + 0.5, w: tw - 2 * px, h: 0.5, valign: "middle", wrap: false });
      if (t.delta) T(s, t.delta, { x: x + px, y: y + 1.02, w: tw - 2 * px, h: 0.26, fontSize: FS.body, bold: !!t.bad, color: t.bad ? C.accent : C.soft, valign: "middle", wrap: false });
      if (t.sub) T(s, t.sub, { x: x + px, y: y + 1.28, w: tw - 2 * px, h: 0.26, fontSize: FS.body, color: C.soft, valign: "middle", wrap: false });
      if (t.spark && th > 1.9) {
        const sx0 = x + px, sx1 = x + tw - px - 0.1, sb = y + th - 0.18, stp = t.spark.length > 1 ? (sx1 - sx0) / (t.spark.length - 1) : 0;
        const sTop = Math.max(y + (t.sub ? 1.62 : 1.38), sb - 0.6), smin = Math.min(...t.spark), smax = Math.max(...t.spark), yOf = v => sb - (v - smin) / (smax - smin || 1) * (sb - sTop);
        chPath(s, t.spark.map((v, k) => [sx0 + k * stp, yOf(v)]), { line: chLine(C.faint, 1.25) });
        chDot(s, sx1, yOf(t.spark[t.spark.length - 1]), 0.09, C.emph);
      }
      [t.delta, t.sub].forEach(tx => { if (tx && textW(tx, FS.body) > tw - 2 * px) warn(`KPI のタイルの文字が入らない（${tx}）`); });
    });
    return { tile: i => out[i] };
  };

  // 左右対称の横棒（人口ピラミッド、図形）：2群（男性と女性、自社と市場）を同じ項目で左右に比べる。項目名は中央の列、群の名前は上、値は棒の外
  //   o.rows: [{ label, left, right, strong? }]（上から並べる順。年齢なら高い年齢を上に）, o.names: [左の名前, 右の名前], o.max, o.fmt
  //   o.y0（最初の棒の上端 2.85）, o.pitch（既定は本文の下端までを等分）, o.midW（中央の列 1.1）, o.cx（中央 W/2）
  //   強調: リードが名指しした行だけ strong（両側の棒を濃紺、値を太字）
  //   戻り値: { rowY(i), bh, k, mw（項目名の欄の幅）, cx }
  const butterfly = (s, o) => {
    const rows = o.rows, cx = o.cx ?? W / 2, y0 = o.y0 ?? 2.85, fmt = o.fmt || (v => fmtNum(v));
    const mw = o.midW ?? Math.max(1.1, ...rows.map(r => textW(r.label, FS.body) + 0.24)), vw = Math.max(0.85, ...rows.flatMap(r => [textW(fmt(r.left), FS.block) + 0.1, textW(fmt(r.right), FS.block) + 0.1]));
    const pitch = o.pitch ?? (BOTTOM - y0) / rows.length, bh = Math.min(0.4, pitch * 0.72), rowY = i => y0 + i * pitch + (pitch - bh) / 2;
    const max = o.max ?? Math.max(...rows.flatMap(r => [r.left, r.right])), span = cx - mw / 2 - (M + vw + 0.05), k = span / max;
    const fs = pitch < 0.3 ? FS.note : FS.body;
    T(s, o.names[0], { x: cx - mw / 2 - 3, y: y0 - 0.34, w: 3, h: 0.28, fontSize: FS.block, bold: true, color: C.soft, align: "right", valign: "bottom", wrap: false });
    T(s, o.names[1], { x: cx + mw / 2, y: y0 - 0.34, w: 3, h: 0.28, fontSize: FS.block, bold: true, color: C.soft, valign: "bottom", wrap: false });
    rows.forEach((r, i) => {
      const y = rowY(i), st = !!r.strong, col = st ? C.emph : C.bar, wl = r.left * k, wr = r.right * k;
      rect(s, cx - mw / 2 - wl, y, wl, bh, col); rect(s, cx + mw / 2, y, wr, bh, col);
      T(s, r.label, { x: cx - mw / 2, y, w: mw, h: bh, fontSize: fs, bold: st, color: st ? C.ink : C.soft, align: "center", valign: "middle", wrap: false });
      T(s, fmt(r.left), { x: cx - mw / 2 - wl - 0.06 - vw, y, w: vw, h: bh, fontSize: st ? FS.block : fs, bold: st, color: st ? C.ink : C.soft, align: "right", valign: "middle", wrap: false });
      T(s, fmt(r.right), { x: cx + mw / 2 + wr + 0.06, y, w: vw, h: bh, fontSize: st ? FS.block : fs, bold: st, color: st ? C.ink : C.soft, valign: "middle", wrap: false });
    });
    return { rowY, bh, k, mw, cx };
  };

  // 市場の絞り込み（TAM・SAM・SOM、図形）：市場をどこまで絞り、狙う金額はいくらかを1つの図で。左に下端で接する入れ子の円、右に同じ高さの行を3段
  //   円の大きさは金額に比例させない（桁が違いすぎて比例では読めないため。金額は行の数字で読ませる）。1金額ごとに対象と算定の根拠を対で書く。仮値は金額の文字に（仮）を付ける
  //   o.items: [{ name（"TAM" など）, what（何の市場かの一言。円の中に出す。その円に入る字数を超えると警告。無いと警告）, scope（対象）, basis（算定の根拠）, value（金額の文字。例 "約11億ドル"）, sub?（金額の下の小さな文字。例 "2024年・年+18.6%"）, strong? }] の3件（外側から順。3件以外は throw）
  //   o.y0（既定 Y.body＋0.05。figHead を置くなら 2.55）, o.bottom（既定 BOTTOM）, o.circleW（円の列の幅＝円の直径。既定は本文の高さ）, o.colW（[区分と金額, 対象, 算定の根拠] の幅。既定は帯の幅から）
  //   強調: strong は1件だけ（帯を TINT_STRONG＋強調の色の 2pt の枠にし、金額を主役の値 36pt・強調の色で部品が描く。呼ぶ側で金額を空にして focusValue を手で置かない）
  //     金額を空（空白だけ）にした古い呼び方は、今まで通り 22pt で描く（呼ぶ側が focusValue を置く）
  //   戻り値: { row(i) => { x, y, w, h }, circle(i) => { cx, cy, r } }
  const tamSamSom = (s, o) => {
    const items = o.items || [];
    if (items.length !== 3) throw new Error(`tamSamSom：items は3件にする（${items.length}件）`);
    if (items.filter(it => it.strong).length > 1) warn("市場の絞り込み：strong が2件以上ある。強調は1ページ1か所にする");
    // 略語（TAM・SAM・SOM）と金額だけでは、何の市場か分からない。各段に、何を指すかを一言で書く（what。全角12字まで）
    items.forEach(it => { if (!it.what) warn(`市場の絞り込み：${it.name} が何の市場かを、一言で書く（items の what。例「国内の中小企業」）`); }); // 一言の長さは、下で円に入る字数で確かめる
    const y0 = o.y0 ?? Y.body + 0.05, bottom = o.bottom ?? BOTTOM, areaH = bottom - y0;
    if (y0 < BODY - 0.005 || bottom > BOTTOM + 0.005) warn(`市場の絞り込み：図（${y0.toFixed(2)}〜${bottom.toFixed(2)}）が本文の領域 ${BODY}〜${BOTTOM} を超える`);
    // 円：3つとも下端で接する。半径の比は金額と無関係の固定値（区分名を置く余白が各円の上に残る比）
    const cw = o.circleW ?? areaH, D = Math.min(cw, areaH), cx = M + cw / 2;
    const circles = [1, 0.69, 0.38].map(k => { const r = D / 2 * k; return { cx, cy: bottom - r, r }; });
    // 行：compareTable と同じ薄い塗りの帯（行の間は白い溝、罫線なし）。列見出しは最初の行の上に1回だけ
    const headH = 0.3, rg = 0.07, padL = 0.18, padR = 0.14, cg = 0.18, rowX = M + cw + 0.3, rowW = W - M - rowX, rowTop = y0 + headH;
    const rowH = (bottom - rowTop - 2 * rg) / 3, rowY = i => rowTop + i * (rowH + rg);
    const inner = rowW - padL - padR - 2 * cg, col1 = 2.5, colW = o.colW ?? [col1, (inner - col1) * 0.44, (inner - col1) * 0.56];
    if (colW[0] + colW[1] + colW[2] > inner + 0.01) warn(`市場の絞り込み：列の幅の合計 ${(colW[0] + colW[1] + colW[2]).toFixed(2)} が帯の中の幅 ${inner.toFixed(2)} を超える`);
    const colX = [rowX + padL, rowX + padL + colW[0] + cg, rowX + padL + colW[0] + colW[1] + 2 * cg];
    circles.forEach((c, i) => chDot(s, c.cx, c.cy, 2 * c.r, [TINT, TINT_STRONG, C.emph][i]));
    // 結ぶ線：行の中央の高さで円の右側の縁から出す（既定の寸法では水平）。円の中の区分名は中央の列にあるので、線は文字を横切らない
    circles.forEach((c, i) => {
      const ym = rowY(i) + rowH / 2, yS = Math.max(c.cy - c.r * 0.8, Math.min(c.cy + c.r * 0.8, ym));
      segLine(s, c.cx + Math.sqrt(c.r * c.r - (yS - c.cy) * (yS - c.cy)), yS, rowX, ym, chLine(C.line, 0.75));
    });
    // 区分名：各円の上の余白（次の円の上端までの中央。一番内側は円の中央より少し上）
    circles.forEach((c, i) => {
      const top = c.cy - c.r, mid = i < 2 ? (top + circles[i + 1].cy - circles[i + 1].r) / 2 : c.cy, lh = items[i].what ? 0.62 : 0.3, ly = mid - lh / 2; // 円の中に、区分名と何の市場かの一言（2026-11-24。右の行は金額と根拠）く。同じ言葉を2か所に置かない）
      // 円に入る幅：各行の上端と下端のうち、円の中心から遠いほうの高さの弦（狭いほう）で測る
      const chordAt = yy => 2 * Math.sqrt(Math.max(0, c.r * c.r - (c.cy - yy) * (c.cy - yy)));
      const lineW = (yTop, h) => Math.min(chordAt(yTop), chordAt(yTop + h)) - 0.12;
      const fg = i === 2 ? C.white : C.ink, wt = items[i].what || "";
      const nameH = lineH(FS.body, 1.15), wW = wt ? lineW(ly + nameH, lineH(FS.block, 1.0)) : 0, nW = lineW(ly, lineH(FS.body, 1.0));
      if (textW(items[i].name, FS.body) > nW) warn(`市場の絞り込み：区分名「${items[i].name}」が円に入らない`);
      if (wt && textW(wt, FS.block) > wW) warn(`市場の絞り込み：「${wt}」が円に入らない（全角${Math.round(em(wt))}字。この円に入るのは全角${Math.max(0, Math.floor(wW * 72 / FS.block))}字まで）。言葉を縮める`);
      T(s, [{ text: items[i].name, options: { fontSize: FS.body, bold: true, color: fg, breakLine: !!wt } }, ...(wt ? [{ text: wt, options: { fontSize: FS.block, bold: true, color: fg } }] : [])],
        { x: c.cx - c.r, y: ly, w: 2 * c.r, h: lh, align: "center", valign: "middle", wrap: false, lineSpacingMultiple: 1.15 });
    });
    (o.heads || ["区分と金額", "対象", "算定の根拠"]).forEach((t, j) => T(s, t, { x: colX[j], y: y0, w: colW[j], h: headH - 0.04, fontSize: FS.note, bold: true, color: C.gray, valign: "bottom", wrap: false }));
    items.forEach((it, i) => {
      const y = rowY(i), st = !!it.strong, legacy = st && !String(it.value ?? "").trim(), vf = st ? (legacy ? FS_DISPLAY.num : FOCUS_FS) : FS.lead;
      if (st) rect(s, rowX, y, rowW, rowH, TINT_STRONG, chLine(C.emph, 2));
      else if (i < items.length - 1 && !items[i + 1].strong) hline(s, rowX, y + rowH + rg / 2, rowW, { color: C.line, width: 0.5 });
      if (i === 0) hline(s, rowX, rowTop - 0.03, rowW, { color: C.gray, width: 1 });
      T(s, [
        // 区分名の横に、何の市場かの一言（略語は小さく灰、一言を太く）
        { text: it.name, options: { fontSize: FS.block, bold: true, color: C.ink, breakLine: true } }, // 何の市場かは円の中に書く
        { text: it.value, options: { fontSize: vf, bold: true, color: st && !legacy ? C.emph : C.ink, breakLine: !!it.sub } },
        ...(it.sub ? (warnSub("市場の図の行"), [{ text: it.sub, options: { fontSize: FS.note, color: C.soft } }]) : []),
      ], { x: colX[0], y, w: colW[0], h: rowH, valign: "middle", lineSpacingMultiple: 1.1 });
      if (String(it.value).split("\n").some(l => textW(l, vf) > colW[0])) warn(`市場の絞り込み：金額が列に収まらない（${it.value}${st && !legacy ? "。主役の金額は 36pt" : ""}）。colW を広げるか言葉を縮める`);
      const need = lineH(FS.block, 1.1) + String(it.value).split("\n").length * lineH(vf, 1.1) + (it.sub ? lineCount(it.sub, FS.note, colW[0]) * lineH(FS.note, 1.1) : 0);
      if (need > rowH - 0.08) warn(`市場の絞り込み：区分と金額の文字があふれる（${it.name}）`);
      [[it.scope, 1], [it.basis, 2]].forEach(([t, j]) => {
        T(s, t, { x: colX[j], y, w: colW[j], h: rowH, fontSize: FS.body, color: C.soft, valign: "middle", lineSpacingMultiple: 1.15 });
        checkWrap(t, FS.body, colW[j]);
        if (lineCount(t, FS.body, colW[j]) * lineH(FS.body, 1.15) > rowH - 0.1) warn(`市場の絞り込み：文字があふれる（${String(t).slice(0, 12)}…）`);
      });
    });
    return { row: i => ({ x: rowX, y: rowY(i), w: rowW, h: rowH }), circle: i => ({ ...circles[i] }) };
  };

  // ---- 書き出し ----
  // 後描き（目次・参照・ミニ目次）→ 書き出し → lang="ja-JP" 置換（禁則を効かせる）→ グラフの棒ごとの強調 → 保存（5.2・4.6・11.3）
  // ビジネスモデル図：主体（人・会社・製品・場）を格子に置き、その間を流れるもの（お金・物やサービス・情報や権利）を矢印で描く。誰が誰に何を渡し、どこでお金が入るかを1枚で読ませる
  //   使う問い：「誰が、誰に、何を渡すか」。比べる（compareTable）・順に進む（processFlow）・回る（cycleDiagram）とは別の形
  //   o.cols（3）, o.rows（2）：格子。o.nodes: [{ id, col, row, name, sub?, icon?（アイコンの名前か配列）, strong? }]
  //   o.flows: [{ from, to, label, kind?: "money" | "goods"（既定）| "info", lane?（同じ2者の間に2本あるとき、線を横へずらす量。1 か -1）, at?: "t"|"b"|"l"|"r"（名前を印のどちら側に置くか）, labelW?, strong? }]
  //     印：お金＝円の中に ¥、物やサービス＝塗りの四角、情報や権利＝白抜きの四角。強調（strong）は1本だけ：ink の太い線と太字
  //   線は、水平・垂直と鉤線だけ。斜め線は使わない（主体が同じ行・同じ列に無ければ、直角に1回折る。flows の via: "h" | "v" で、先に横か先に縦かを選ぶ）
  //   o.only: { nodes: [id], flows: [番号] }（その要素だけを描く。echo の keep・series の only で使う）
  //   o.x, o.w, o.y0（BODY＋0.1）, o.bottom（BOTTOM−0.3。下に印の凡例を1行置く）, o.nodeW（既定は主体の名前から、2.5〜3.0）, o.nodeH（1.1）, o.labelW（縦の線の名前の幅 2.0）
  //   戻り値: { node(id) → {x, y, w, h, cx, cy}, flow(i) → { mx, my, p1, p2 } }
  const bizModel = (s, o) => {
    needConcept("bizModel");
    const nc = o.cols ?? 3, nr = o.rows ?? 2, x0 = o.x ?? M, wAll = o.w ?? BW, y0 = o.y0 ?? BODY + 0.1, bottom = o.bottom ?? BOTTOM - 0.3;
    const needW = Math.max(...o.nodes.map(nd => { const k = [].concat(nd.icon || []).length; return Math.max(...String(nd.name).split("\n").map(t => textW(t, FS.head))) + (k ? k * 0.48 + 0.04 : 0) + 0.3; }));
    const nw = o.nodeW ?? Math.min(3.0, Math.max(2.5, needW)), nh = o.nodeH ?? 1.1, lw0 = o.labelW ?? 2.0;
    const cxOf = c => x0 + nw / 2 + (nc > 1 ? c * (wAll - nw) / (nc - 1) : (wAll - nw) / 2);
    const cyOf = r => y0 + nh / 2 + (nr > 1 ? r * (bottom - y0 - nh) / (nr - 1) : (bottom - y0 - nh) / 2);
    const only = o.only || null, showN = id => !only || (only.nodes || []).includes(id), showF = i => !only || (only.flows || []).includes(i);
    const N = {}; o.nodes.forEach(nd => { const cx = cxOf(nd.col), cy = cyOf(nd.row); N[nd.id] = { ...nd, cx, cy, x: cx - nw / 2, y: cy - nh / 2, w: nw, h: nh }; });
    if (o.flows.filter(f => f.strong).length > 1) warn("ビジネスモデル図の強調（strong）は1本だけにする");
    // 線：水平・垂直と鉤線だけ。斜め線は使わない（主体が同じ行・同じ列に無いときは、直角に1回折る）
    //   折れる角は「上の行の主体の真下 × 下の行の主体の真横」。下の主体の横から出て、上の主体の下へ入る（via: "h" で先に横、"v" で先に縦、を選べる）
    //   同じ2者の間に2本あるときは lane で平行にずらす（向きが逆の矢印は、同じ lane でも内側と外側に分かれる）
    const G = 0.1, LANE = 0.14;
    const F = o.flows.map((f, i) => {
      const a = N[f.from], b = N[f.to];
      if (!a || !b) { warn(`ビジネスモデル図の矢印の相手が無い（${f.from} → ${f.to}）`); return null; }
      const sameRow = Math.abs(a.cy - b.cy) < 0.01, sameCol = Math.abs(a.cx - b.cx) < 0.01, d0 = (f.lane || 0) * LANE;
      let pts;
      if (sameRow) { const dir = b.cx > a.cx ? 1 : -1, y = a.cy + d0 * dir; pts = [[a.cx + dir * (nw / 2 + G), y], [b.cx - dir * (nw / 2 + G), y]]; }
      else if (sameCol) { const dir = b.cy > a.cy ? 1 : -1, x = a.cx - d0 * dir; pts = [[x, a.cy + dir * (nh / 2 + G)], [x, b.cy - dir * (nh / 2 + G)]]; }
      else {
        // 上の主体 U と下の主体 L。角は (U の真下, L の真横)。外側の線ほど、横の線は下へ、縦の線は L から遠い側へずらす
        const fromLower = a.cy > b.cy, U = fromLower ? b : a, L = fromLower ? a : b, side = U.cx > L.cx ? 1 : -1, dl = (fromLower ? 1 : -1) * d0;
        const via = f.via || "h";
        let lp, corner, up;
        if (via === "h") { lp = [L.cx + side * (nw / 2 + G), L.cy + dl]; corner = [U.cx + side * dl, L.cy + dl]; up = [U.cx + side * dl, U.cy + nh / 2 + G]; }
        else { lp = [L.cx - side * dl, L.cy - nh / 2 - G]; corner = [L.cx - side * dl, U.cy - dl]; up = [U.cx - side * (nw / 2 + G), U.cy - dl]; }
        pts = fromLower ? [lp, corner, up] : [up, corner, lp];
      }
      // 斜めの線が1本でもあれば止める（決まり）。ほかの主体の箱を横切る線は読めないので知らせる
      for (let k = 1; k < pts.length; k++) {
        const [x1, y1] = pts[k - 1], [x2, y2] = pts[k];
        if (Math.abs(x1 - x2) > 0.001 && Math.abs(y1 - y2) > 0.001) throw new Error(`ビジネスモデル図：斜めの線になる（${f.from} → ${f.to}）`);
        Object.values(N).forEach(nd => { if (nd.id === f.from || nd.id === f.to) return; for (let q = 1; q < 10; q++) { const x = x1 + (x2 - x1) * q / 10, y = y1 + (y2 - y1) * q / 10; if (x > nd.x && x < nd.x + nd.w && y > nd.y && y < nd.y + nd.h) { warn(`ビジネスモデル図の矢印（${f.from} → ${f.to}）が、「${nd.name}」の箱を横切る（主体の置き場所か via を変える）`); break; } } });
      }
      // 印と名前は、いちばん長い線の真ん中に置く
      let best = 1, len = 0;
      for (let k = 1; k < pts.length; k++) { const l = Math.abs(pts[k][0] - pts[k - 1][0]) + Math.abs(pts[k][1] - pts[k - 1][1]); if (l > len) { len = l; best = k; } }
      const mx = (pts[best][0] + pts[best - 1][0]) / 2, my = (pts[best][1] + pts[best - 1][1]) / 2, horiz = Math.abs(pts[best][1] - pts[best - 1][1]) < 0.001;
      return { f, i, pts, p1: pts[0], p2: pts[pts.length - 1], mx, my, horiz, len, outer: d0 * (a.cy > b.cy ? 1 : -1) > 0 };
    });
    F.forEach(fl => {
      if (!fl || !showF(fl.i)) return;
      const { f, pts, mx, my } = fl, st = !!f.strong, col = st ? C.emph : C.gray, d = 0.26;
      polyArrow(s, pts, { color: col, width: st ? 2.25 : 1.25 });
      if ((f.kind || "goods") === "money") s.addText("¥", { shape: SH.OVAL, x: mx - d / 2, y: my - d / 2, w: d, h: d, fill: { color: st ? C.emph : C.white }, line: { color: col, width: 1.25 }, fontFace: FNUM, fontSize: 10, bold: true, color: st ? C.white : C.gray, align: "center", valign: "middle", margin: 0 });
      else rect(s, mx - 0.08, my - 0.08, 0.16, 0.16, f.kind === "info" ? C.white : col, { color: col, width: 1.25 });
      // 名前：横の線なら上（外側にずらした線は下）、縦の線なら右。線に重ねない
      const at = f.at || (fl.horiz ? (fl.outer ? "b" : "t") : "r"), fsz = st ? FS.block : FS.body;
      const lw = f.labelW ?? (fl.horiz ? Math.max(1.2, Math.min(fl.len + 2 * G - 0.04, 3.4)) : lw0), lh = lineCount(f.label, fsz, lw) * fsz / 72 * 1.3 + 0.06;
      const bx = at === "t" ? { x: mx - lw / 2, y: my - 0.2 - lh, align: "center", valign: "bottom" } : at === "b" ? { x: mx - lw / 2, y: my + 0.2, align: "center", valign: "top" }
        : at === "l" ? { x: mx - 0.24 - lw, y: my - lh / 2, align: "right", valign: "middle" } : { x: mx + 0.24, y: my - lh / 2, align: "left", valign: "middle" };
      T(s, f.label, { ...bx, w: lw, h: lh, fontSize: fsz, bold: st, color: st ? C.ink : C.soft, lineSpacingMultiple: 1.15 });
    });
    Object.values(N).forEach(nd => {
      if (!showN(nd.id)) return;
      rect(s, nd.x, nd.y, nd.w, nd.h, nd.strong ? TINT_STRONG : TINT, nd.strong ? { color: C.emph, width: 2 } : null);
      const ics = [].concat(nd.icon || []), d = 0.42, iw = ics.length ? ics.length * (d + 0.06) + 0.04 : 0;
      ics.forEach((nm, q) => icon(s, nm, nd.x + 0.18 + d / 2 + q * (d + 0.06), nd.cy, d));
      const tx = nd.x + 0.18 + iw, tw = nd.x + nd.w - 0.1 - tx;
      checkWrap(nd.name, FS.head, tw);
      if (nd.sub) warnSub("ビジネスモデル図の主体");
      T(s, [{ text: nd.name, options: { fontSize: FS.head, bold: true, color: C.ink, breakLine: !!nd.sub } }, ...(nd.sub ? [{ text: nd.sub, options: { fontSize: FS.note, color: C.soft } }] : [])], { x: tx, y: nd.y, w: tw, h: nd.h, valign: "middle", lineSpacingMultiple: 1.15 });
    });
    // 印の意味（図の左下に1行）。既定では出さない。矢印の名前で読めるようにする。出すなら o.legend: true
    if (!only && o.legend === true) {
      const ly = bottom + 0.12; let lx = x0;
      [["money", "対価"], ["goods", "物・サービス"], ["info", "情報・権利"]].forEach(([k, t]) => {
        if (!o.flows.some(f => (f.kind || "goods") === k)) return;
        if (k === "money") s.addText("¥", { shape: SH.OVAL, x: lx, y: ly, w: 0.2, h: 0.2, fill: { color: C.white }, line: { color: C.gray, width: 1 }, fontFace: FNUM, fontSize: 8, bold: true, color: C.gray, align: "center", valign: "middle", margin: 0 });
        else rect(s, lx + 0.03, ly + 0.03, 0.14, 0.14, k === "info" ? C.white : C.gray, { color: C.gray, width: 1 });
        T(s, t, { x: lx + 0.27, y: ly - 0.04, w: 1.6, h: 0.28, fontSize: FS.body, color: C.soft, valign: "middle", wrap: false }); // 凡例は読み方に要るので、読める大きさで
        lx += 0.27 + textW(t, FS.body) + 0.34;
      });
    }
    return { node: id => N[id], flow: i => F[i] };
  };

  // 指標ごとの棒：単位や桁の違う数値を、表に並べずに比べる。行＝比べる対象、列＝指標。列ごとに尺度をそろえた横棒と値（同じ列の棒は、縦に長さを見比べられる）
  //   数値を表に並べると、「2」と「840」のように桁の違う値が縦横に並び、どこを比べればよいのか分からなくなる。比べるのは同じ指標の中だけなので、指標ごとに棒にする
  //   o.items: [対象の名前]（行。3〜5）, o.focus: 主役の行の番号（棒を強調の色、名前を太字）
  //   o.metrics: [{ label, unit?, values: [数値], fmt? }]（列。2〜4）, o.key: [行, 列]（決め手の値1つ。大きくする）
  //   o.x, o.w, o.y0（BODY＋0.2）, o.bottom（BOTTOM）, o.labelW（行の名前の幅。既定は中身から）, o.pitch（行の高さ。既定は 0.95 まで）
  //   戻り値: { cell(i, j) → {x, y, w, h}, row(i) → {x, y, w, h}, bottom }
  const metricBars = (s, o) => {
    chartPage();
    const items = o.items, ms = o.metrics, n = ms.length, x0 = o.x ?? M, wAll = o.w ?? BW, y0 = o.y0 ?? BODY + 0.2, bottom = o.bottom ?? BOTTOM;
    if (items.length > 5 || n > 4) warn(`指標ごとの棒が大きい（${items.length}行×${n}列）。5行×4列まで`);
    const lw = labelCol("指標ごとの棒", items, o.labelW, FS.block, 3.6), cg = 0.4, lg = 0.2;
    const cw = (wAll - lw - lg - cg * (n - 1)) / n, colX = j => x0 + lw + lg + j * (cw + cg);
    const headH = 0.62, pitch = o.pitch ?? Math.min(0.85, (bottom - y0 - headH) / items.length), bh = Math.min(pitch * 0.55, FS.body / 72 * 1.5);
    const key = o.key ? `${o.key[0]}:${o.key[1]}` : null;
    ms.forEach((m, j) => {
      const x = colX(j), max = Math.max(...m.values, 1e-9), fmt = m.fmt || (v => fmtNum(v));
      // 単位は、名前の後ろに同じ行で書く（下の小さな灰の行にしない）
      T(s, [{ text: m.label, options: { fontSize: FS.block, bold: true, color: C.ink } }, ...(m.unit ? [{ text: `（${m.unit}）`, options: { fontSize: FS.block, color: C.soft } }] : [])], { x, y: y0, w: cw, h: headH - 0.1, valign: "bottom", lineSpacingMultiple: 1.1 });
      const valW = Math.max(...m.values.map((v, i) => textW(fmt(v), key === `${i}:${j}` ? FOCUS_FS : FS.block))) + 0.2, barMax = cw - valW;
      m.values.forEach((v, i) => {
        const y = y0 + headH + i * pitch, cy = y + pitch / 2, on = i === o.focus, k = key === `${i}:${j}`, bwid = Math.max(0.03, v / max * barMax);
        rect(s, x, cy - bh / 2, bwid, bh, k || (on && !key) ? C.emph : on ? C.gray : C.bar);
        T(s, fmt(v), { x: x + bwid + 0.12, y: cy - 0.25, w: valW, h: 0.5, fontSize: k ? FOCUS_FS : FS.block, bold: on || k, color: k ? C.emph : on ? C.ink : C.soft, valign: "middle", wrap: false });
      });
    });
    hline(s, x0, y0 + headH - 0.04, wAll, { color: C.gray, width: 1 });
    items.forEach((t, i) => {
      const y = y0 + headH + i * pitch, on = i === o.focus;
      T(s, t, { x: x0, y, w: lw + 0.06, h: pitch, fontSize: FS.block, bold: true, color: on ? C.ink : C.soft, align: "right", valign: "middle", wrap: false });
      if (i < items.length - 1) hline(s, x0, y + pitch, wAll, { color: C.line, width: 0.5 });
    });
    const figBottom = y0 + headH + items.length * pitch;
    ms.forEach((_, j) => s.addShape(SH.LINE, { x: colX(j), y: y0 + headH - 0.04, w: 0, h: figBottom - y0 - headH, line: { color: C.gray, width: 0.75 } })); // 列ごとの 0 の線
    if (figBottom > BOTTOM + 0.05) warn(`指標ごとの棒の下端 ${figBottom.toFixed(2)} が本文の下端を超える`);
    return { cell: (i, j) => ({ x: colX(j), y: y0 + headH + i * pitch, w: cw, h: pitch }), row: i => ({ x: x0, y: y0 + headH + i * pitch, w: wAll, h: pitch }), bottom: figBottom };
  };

  // 段階の道（ロードマップ）：横の道に段階の駅を並べ、駅の下に「次へ進む条件」を値で置く。いちばん下に、資金や体制の帯
  //   段階×観点の表（journeyMap）の代わり。マス目に文を積まず、段階ごとに、名前・時期・絵・条件の値（最大4つ）だけを置く
  //   o.stages: [{ name, when?, icon?（丸の中の絵。icons/ の名前）, kpis?: [[値, 何の値か, 絵?], ...]（最大4つ。「値　何の値か」が1行で箱の幅に入らないと警告。4段階なら全角10字ほど）, strong? }]（3〜6段階。strong＝見る所の段階を強調の色で塗る。1つだけ）
  //   o.bands: [{ from, to, label, strong? }]（下の帯。資金・体制など）, o.y0（BODY）
  //   戻り値: { col(i) → {x, y, w, h}, X(i), CW, K0, KP, right, bottom }
  const roadmap = (s, o) => {
    const st = o.stages, n = st.length, y0 = o.y0 ?? BODY;
    if (n < 3 || n > 6) warn(`段階の道：段階は3〜6つにする（${n}）`);
    if (st.filter(x => x.strong).length > 1) warn("段階の道：強調（strong）は1つだけにする");
    const CW = Math.min(2.9, (BW - 0.1) / n - 0.08), X0 = M + CW / 2 + 0.05, STEP = (W - 2 * M - CW - 0.1) / (n - 1), X = i => X0 + i * STEP;
    const WY = y0, NY = y0 + 0.82, R = 0.42, NAME_Y = y0 + 1.36, K0 = y0 + 1.88, KP = 0.47, BH = 0.36;
    const maxK = Math.max(0, ...st.map(x => (x.kpis || []).length)), BY = K0 + Math.max(maxK, 1) * KP + 0.18;
    if (maxK > 4) warn("段階の道：条件は、段階ごとに4つまで");
    const has = nm => !!findIcon(nm);
    segLine(s, X(0), NY, X(n - 1), NY, { color: C.line, width: 4 });
    st.forEach((x, i) => {
      const cx = X(i), on = !!x.strong;
      if (x.when && textW(x.when, FS.body) > CW) warn(`段階の道：時期の文字が段階の幅に入らない（${x.when}。全角${Math.floor(CW * 72 / FS.body)}字まで）`);
      if (x.when) T(s, x.when, { x: cx - CW / 2, y: WY - 0.04, w: CW, h: 0.34, fontSize: FS.body, bold: true, color: on ? C.emph : C.soft, align: "center", valign: "middle", wrap: false });
      s.addShape(SH.OVAL, { x: cx - R, y: NY - R, w: 2 * R, h: 2 * R, fill: { color: on ? C.emph : C.white }, line: on ? { type: "none" } : { color: C.line, width: 2.25 } });
      if (x.icon) icon(s, x.icon, cx, NY, 0.46, on ? "w" : undefined);
      else T(s, String(i + 1), { x: cx - R, y: NY - R, w: 2 * R, h: 2 * R, fontFace: FNUM, fontSize: 18, bold: true, color: on ? C.white : C.gray, align: "center", valign: "middle" });
      if (textW(x.name, FS.block) > CW - 0.04) warn(`段階の道：段階の名前が長い（${x.name}。${FS.block}pt で全角${Math.floor((CW - 0.04) * 72 / FS.block)}字まで）`);
      T(s, x.name, { x: cx - CW / 2, y: NAME_Y, w: CW, h: 0.34, fontSize: FS.block, bold: true, color: on ? C.emph : C.ink, align: "center", valign: "middle", wrap: false });
      (x.kpis || []).forEach(([big, small, pic], k) => {
        const y = K0 + k * KP, ix = pic ? 0.4 : 0;
        // 「値　何の値か」を1行で書く（折り返さない）。今の文字の大きさ（本文 FS.body）で、箱の幅に入るかを測る
        const kw = CW - 0.28 - ix, need = textW(big, FS.body) + (small ? textW(`　${small}`, FS.body) : 0);
        if (need > kw) warn(`段階の道：条件「${big}${small ? "　" + small : ""}」が箱の幅に入らない（全角${Math.ceil(em(`${big}${small ? "　" + small : ""}`))}字分。この幅で入るのは全角${Math.floor(kw * 72 / FS.body)}字まで）。言葉を縮める`);
        rect(s, cx - CW / 2 + 0.08, y, CW - 0.16, KP - 0.07, TINT);
        if (pic) icon(s, pic, cx - CW / 2 + 0.3, y + (KP - 0.07) / 2, 0.24);
        T(s, [{ text: big, options: { fontSize: FS.body, bold: true, color: C.ink } }, ...(small ? [{ text: `　${small}`, options: { fontSize: FS.body, color: C.soft } }] : [])], { x: cx - CW / 2 + 0.16 + ix, y, w: CW - 0.28 - ix, h: KP - 0.07, valign: "middle", wrap: false });
      });
    });
    (o.bands || []).forEach(b => {
      const x = X(b.from) - CW / 2 + 0.08, w = X(b.to) - X(b.from) + CW - 0.16, on = !!b.strong;
      s.addShape(SH.ROUNDED_RECTANGLE, { x, y: BY, w, h: BH, rectRadius: BH / 2, fill: { color: on ? C.emph : C.line }, line: { type: "none" } });
      T(s, b.label, { x, y: BY, w, h: BH, fontSize: FS.body, bold: true, color: on ? C.white : C.ink, align: "center", valign: "middle", wrap: false });
      if (textW(b.label, FS.body) > w - 0.2) warn(`段階の道：下の帯の文字が帯の幅に入らない（${b.label}）`);
    });
    const bottom = (o.bands || []).length ? BY + BH : BY - 0.18;
    if (bottom > BOTTOM + 0.05) warn(`段階の道の下端 ${bottom.toFixed(2)} が本文の下端を超える（条件を減らす）`);
    return { col: i => ({ x: X(i) - CW / 2, y: WY - 0.04, w: CW, h: bottom - WY + 0.1 }), X, CW, K0, KP, right: X(n - 1) + CW / 2, bottom };
  };

  // 1つの形は、1つの資料の中で1つの意味にだけ使う：同じ部品を、別の中身でもう一度呼んだら知らせる
  //   同じ図をもう一度出す（echo・series。only だけが違う）のは数えない。月ごとの同じ表のように、同じ意味で中身だけ違うときは、引数の fig に同じ名前を書く
  const formGuard = (name, fn) => (...args) => {
    const o = args.find((a, i) => i > 0 && a && typeof a === "object" && !Array.isArray(a)) || {};
    const key = o.fig ?? JSON.stringify(o, (k, v) => (k === "only" || k === "fig" ? undefined : v));
    const u = FORM_USE[name];
    if (!u) FORM_USE[name] = { key, page, told: {} };
    else if (u.key !== key && !u.told[key]) { u.told[key] = 1; warn(`同じ形（${name}）を、P.${u.page} とは別の中身で使っている。別の目的なら別の形にする。同じ図の続きなら同じ引数で描く（series）。同じ意味で中身だけ違うなら、両方の引数に同じ fig を書く`); }
    return fn(...args);
  };

  // 画像の代替テキスト（descr）がファイルの場所になっていないか。pptxgenjs は altText が無いと画像のパス（ユーザー名を含む）を入れる
  const isFilePath = v => /^(\/|~[\/\\]|[A-Za-z]:[\/\\]|\\\\|file:)/.test(v) || (/[\/\\]/.test(v) && /\.(png|jpe?g|gif|svg|bmp|webp|tiff?|emf|wmf)$/i.test(v));
  const scrubDescr = async zip => {
    const hits = {};
    for (const n of Object.keys(zip.files).filter(n => /^ppt\/(slides|slideLayouts|slideMasters|notesSlides)\/[^/]+\.xml$/.test(n))) {
      const xml = await zip.file(n).async("string");
      const out = xml.replace(/ descr="([^"]*)"/g, (m, v) => { if (!isFilePath(v)) return m; hits[n] = (hits[n] || 0) + 1; return ' descr=""'; });
      if (out !== xml) zip.file(n, out);
    }
    Object.entries(hits).forEach(([n, k]) => {
      const slide = n.match(/^ppt\/slides\/slide(\d+)\.xml$/);
      console.warn(`[deck-parts] ${slide ? `P.${slide[1]}` : n} の画像の代替テキストが、ファイルの場所になっていた（${k} 件）。空にして書き出した。build.js の addImage に、画像が何かの短い説明を altText で書く`);
    });
  };
  const finalize = async outPath => {
    LATER.forEach(fn => fn());
    // 聞き手の前提：宣言していなければ知らせる（2026-11-25）
    if (AUD === undefined) console.warn(`[deck-parts] 聞き手の前提（createDeck の audience）を決めていない。"first-time"（その事業・製品を初めて聞く）か "insider"（前提を共有済み）を書く。初見なら、新しい物は一言のコンセプト（concept）から入る`);
    // アジェンダ：8 枚以上の資料は、表紙（と要約）の後にアジェンダを置く（2026-11-25）
    if (!AGENDA_DONE && page >= 8) console.warn(`[deck-parts] アジェンダのページが無い（${page} 枚）。8 枚以上の資料は、表紙（と要約）の後にアジェンダを置く：createDeck の nav に章立てを書き、agendaPage() を呼ぶ。各ページのタイトルは章の項目名になる。「15 枚」のような枚数の指定は、表紙とアジェンダを数えない本編の数として扱い、アジェンダを省く理由にしない`);
    // グラフのページ：見る所の値（FOCUS_FS・強調の色）が1つだけあるか
    FOCUS.forEach(f => { if (!CHARTP[f.page] || f.focus === false) return; const n = BIGF[f.page] || 0; if (!n) console.warn(`[deck-parts] P.${f.page} グラフの見る所の値が無い。見る所の値を ${FOCUS_FS}pt・強調の色で、見る所の隣に書く（focusValue か、部品の focusValue の口）`); else if (n > 1) console.warn(`[deck-parts] P.${f.page} 見る所の大きな値が ${n} つある。強調の色の大きな値は1つにする`); });
    // リードの数字は、同じページの図で確かめられるようにする。10 までの数え方（「2つ」「5年で」「3社」）は除き、「4割」は 40 でもよい。図の文字が無いページは見ない
    LEADS.forEach(l => {
      if (!l.lead || !l.fig.length) return;
      const inFig = new Set(l.fig.flatMap(numsOf));
      const text = l.lead.normalize("NFKC").replace(/\n/g, "").replace(/(?<![\d.,])([1-9]|10)(?=つ|年(?!度)|か月|週間|時間|項目|工程|段階|社|店|案|種類)/g, "");
      const tens = new Set((text.match(/\d+(\.\d+)?(?=割)/g) || []).map(v => String(Number(v))));
      const missing = [...new Set(numsOf(text))].filter(v => !inFig.has(v) && !(tens.has(v) && inFig.has(String(Number(v) * 10))));
      if (missing.length) console.warn(`[deck-parts] P.${l.page} リードの数字（${missing.join("、")}）が、このページの図に無い。図に同じ値を同じ書き方で示すか、リードから外す`);
    });
    // 聴講者が持ちうる疑問：リードのある全ページで、問いと答えの場所を書いてから点検する（画面には出さない）
    const noAsk = LEADS.filter(l => l.lead && !l.ask).map(l => l.page);
    if (noAsk.length) console.warn(`[deck-parts] 聴講者が持ちうる疑問（base の o.ask）が空のページ：${noAsk.map(p => `P.${p}`).join("、")}。聴講者として「何を？どれ？どうする？どれくらい？」と問い、答えがリードか同じページの図にあるかを書く（無ければリードを直す）`);
    // 外部の製品・会社：図の中に名前があるのに、そのページにアイコンが無ければ知らせる
    Object.keys(BRANDS_ON).forEach(pg => Object.keys(BRANDS_ON[pg]).forEach(k => { if (!(ICONS_ON[pg] || {})[k]) console.warn(`[deck-parts] P.${pg} 外部の製品・会社（${k}）を図に出しているが、アイコンが無い。公式のアイコンを取得して icons/ に置き、名前の横に置く（scripts/fetch_icon.py）`); }));
    // 線の絵の色：資料の色の組（文字の色・強調の色）で塗ってあるかを確かめる
    if (lineIconUsed && (!ICON_TINT || String(ICON_TINT.ink).toUpperCase() !== String(C.ink).toUpperCase() || String(ICON_TINT.emph).toUpperCase() !== String(C.emph).toUpperCase()))
      console.warn(`[deck-parts] 線の絵（アイコン）の色が、資料の色の組と合っていない（アイコン: ${ICON_TINT ? ICON_TINT.ink + "・" + ICON_TINT.emph : "塗っていない"}、資料: ${C.ink}・${C.emph}）。作業フォルダで塗り直す: python3 <skills>/pptx-deck-design/scripts/tint_icons.py icons --ink ${C.ink} --emph ${C.emph} --accent ${C.accent}`);
    // 見る所の一覧：点検で、画像を見て最初に目が行った所と引き比べる（review-checklist）
    if (opts.focusList !== false) fs.writeFileSync(String(outPath).replace(/\.pptx$/i, "") + ".focus.md", "# 見る所の一覧\n\n画像だけを見て「最初に目が行く所」を答え、下の宣言と合うかを確かめる。合わないページは、強調を直す。\n\n| ページ | タイトル | 見る所（宣言） |\n|---|---|---|\n" + FOCUS.filter(f => f.title).map(f => `| ${f.page} | ${f.title} | ${f.focus === false ? "（列挙。見る所なし）" : f.focus ?? "（宣言なし）"} |`).join("\n") + "\n");
    // リードの一覧：リードだけを通して読み、話が通るかを確かめる（review-checklist）
    if (opts.leadList !== false) { const cell = v => String(v || "—").replace(/\n/g, "").replace(/\|/g, "｜"); fs.writeFileSync(String(outPath).replace(/\.pptx$/i, "") + ".leads.md", "# リードの一覧\n\nリードだけを上から順に読み、話が通るかを確かめる。2文目があるページは、足したのが判断か次の行動の1つだけで、蛇足でないかを確かめる。疑問の欄は、聴講者として「何を？どれ？どうする？どれくらい？」と問い、答えがリードか同じページの図にあるかを確かめる。無ければリードを直す。\n\n| ページ | タイトル | リード | 聴講者が持ちうる疑問 |\n|---|---|---|---|\n" + LEADS.filter(l => l.title || l.lead).map(l => `| ${l.page} | ${cell(l.title)} | ${cell(l.lead)} | ${l.lead ? l.ask ? cell(l.ask) : "（未記入）" : "—"} |`).join("\n") + "\n"); }
    // ナビゲーションは、テーマの nav か opts.miniNav: true のときだけ描く（章立て nav は、アジェンダとタイトルに使う。現在地はタイトルで分かる）
    if (NAV && opts.miniNav !== false) {
      const baseT = TH.nav || {}, baseStyle = NAV_STYLE, layouts = {};
      inFooter = true;
      try {
        SLIDES.forEach(r => {
          if (!r.nav) return;
          const own = MASTER[r.type] && MASTER[r.type].nav;
          if (own === false) return;
          const style = (own && own.style) || baseStyle || (own ? "text" : null);
          if (!style) return;
          NAVT = { ...baseT, ...(own || {}) }; NAV_STYLE = style;
          if (!layouts[r.type]) layouts[r.type] = navLayout();
          drawNav(r.s, r.ch, layouts[r.type], r.type);
        });
      } finally { inFooter = false; NAVT = baseT; NAV_STYLE = baseStyle; }
    }
    const JSZip = require("jszip"); // pptxgenjs の依存
    const zip = await JSZip.loadAsync(await pres.write({ outputType: "nodebuffer" }));
    const names = Object.keys(zip.files);
    for (const n of names.filter(n => /^ppt\/(slides\/slide|charts\/chart)\d+\.xml$/.test(n))) zip.file(n, (await zip.file(n).async("string")).replace(/lang="en-US"/g, 'lang="ja-JP"'));
    for (const p of CHART_PATCH) {
      const rels = await zip.file(`ppt/slides/_rels/slide${p.slide}.xml.rels`).async("string");
      const charts = [...rels.matchAll(/Target="(?:\.\.|\/ppt)\/charts\/(chart(\d+)\.xml)"/g)].sort((a, b) => a[2] - b[2]).map(m => m[1]);
      const target = charts[p.nth];
      if (!target) { console.warn(`[deck-parts] P.${p.slide} のグラフが見つからない（強調を入れられない）`); continue; }
      zip.file(`ppt/charts/${target}`, patchChartXml(await zip.file(`ppt/charts/${target}`).async("string"), p));
    }
    await scrubDescr(zip);
    fs.writeFileSync(outPath, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
    console.log("written", outPath, `(${page} pages)`);
    return outPath;
  };

  return {
    // トークン
    C, TINT, TINT_STRONG, BA_BG, CAT, CAT_ALT, STICKY, FS, FS_DISPLAY, GAP, PAD, W, H, M, BW, BODY, BOTTOM, FH, Y, F,
    FHEAD, FNUM, LEAD_FS, theme: TH, // テーマ（読み込んだ設定。colors・font・sizes・footer・cover・layout・icons）
    // 状態
    get page() { return page; }, PAGE, setChapter, mark, later, ref,
    // 文字・線
    T, hline, segLine, rect, link, leadBreak, textW, lineCount, focusValue, FOCUS_FS, FOCUS_LW,
    // 骨格
    base, cover, chapter, divider, agenda, agendaPage, source, bullets, points, summary, concept,
    // 表・比較
    blockTable, compareTable, raciTable: formGuard("raciTable", raciTable), metricBars, roadmap: formGuard("roadmap", roadmap), harvey, journeyMap: formGuard("journeyMap", journeyMap),
    // 因果
    logicTree: formGuard("logicTree", logicTree), kpiTree: formGuard("kpiTree", kpiTree),
    // 時系列
    timeline, drawGantt, projectGantt: formGuard("projectGantt", projectGantt),
    // 重ね技
    overlay, callout, messageBox, echo, series,
    // 小部品
    arrowR, PT, numDot, redDot, person, stickyBoard,
    // 数値
    waterfall, areaBars, barChart, colBars, hBars, ribbonBars, figHead, fmtNum,
    // 構造の図
    processFlow: formGuard("processFlow", processFlow), matrix2x2: formGuard("matrix2x2", matrix2x2), funnel: formGuard("funnel", funnel), archDiagram: formGuard("archDiagram", archDiagram),
    eventTimeline: formGuard("eventTimeline", eventTimeline), stateMap: formGuard("stateMap", stateMap), bizModel: formGuard("bizModel", bizModel), decisionPage, arrowSeg, polyArrow,
    // work-flow
    swimlane: formGuard("swimlane", swimlane), cycleDiagram: formGuard("cycleDiagram", cycleDiagram), orgChart: formGuard("orgChart", orgChart),
    // work-map
    icon, iconPath,
    // work-chart
    refLine, lineChart, areaChart, stackBars, hStackBars, mekko, treemap, pieChart, dumbbell, bulletChart, bridge, comboChart, slopeChart, scatter, heatmap, radarChart, smallMultiples, kpiTiles, butterfly, tamSamSom: formGuard("tamSamSom", tamSamSom),
    // 書き出し
    finalize,
  };
}

module.exports = { createDeck, themeFrom, PALETTES, themeLayers, findThemeFiles, C, TINT, TINT_STRONG, BA_BG, CAT, CAT_ALT, STICKY, FS, FS_DISPLAY, GAP, PAD, W, H, M, BW, BODY, BOTTOM, leadBreak, textW, lineCount };
