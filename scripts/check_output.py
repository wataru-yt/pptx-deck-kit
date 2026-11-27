#!/usr/bin/env python3
"""公開の前の確認: 直す前と後で、見本から作る pptx が変わっていないかを比べる。

使い方:
  python3 scripts/check_output.py                 # origin/main（公開中の版）と、手元の作業中を比べる
  python3 scripts/check_output.py --base HEAD     # 直前のコミットと比べる
  python3 scripts/check_output.py --keep          # 作った pptx を残す（画像にして確かめるとき）

準備: plugins/pptx-deck-kit で npm install（pptxgenjs と jszip）。

比べるもの: 同梱の見本（sample-*.js）。手元に trials/（公開しない試しの資料）があれば、それも。
個人の設定（~/.claude/pptx-deck など）は読まずに作る。
違いが出たら、意図した変更かを画像で確かめる（trials/render.sh か、pptx-build-sync の render_pptx.sh）。
終了コード: 0 = 同じ / 1 = 違いがある / 2 = 作れない・検査に通らない
"""
import argparse
import io
import json
import os
import re
import shutil
import subprocess
import sys
import tarfile
import tempfile
import zipfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[1]
PLUGIN = Path("plugins/pptx-deck-kit")
ASSETS = PLUGIN / "skills/pptx-deck-design/assets"
EMU_TOLERANCE = 2  # 図形の位置と大きさの丸めの差は、違いに数えない
POSITION = re.compile(r'\b(x|y|cx|cy|w|h)="(-?\d+)"')


def git(*args):
    return subprocess.run(["git", "-C", str(REPO), *args], check=True, capture_output=True).stdout


def build(script, cwd, out, env):
    result = subprocess.run(["node", str(script), str(out)], cwd=cwd, env=env, capture_output=True, text=True)
    log = result.stdout + result.stderr
    ok = result.returncode == 0 and out.exists()
    return ok, sum(1 for line in log.splitlines() if re.match(r"\[deck-parts\] P\.\d+", line)), log


def normalize(data):
    text = data.decode("utf-8", "replace")
    text = re.sub(r' descr="[^"]*"', "", text)  # 画像の元のパス（作業の場所で変わる）
    return re.sub(r"\d{4}-\d\d-\d\dT[\d:.]+Z", "T", text)


def same_part(name, a, b):
    if not name.endswith((".xml", ".rels")):
        return a == b
    a, b = normalize(a), normalize(b)
    if a == b:
        return True
    if POSITION.sub(r"\1", a) != POSITION.sub(r"\1", b):
        return False
    return all(abs(int(p[1]) - int(q[1])) <= EMU_TOLERANCE for p, q in zip(POSITION.findall(a), POSITION.findall(b)))


def path_descr(pptx):
    """画像の代替テキストに、作った PC のファイルの場所が入っていないか（比べるときは descr を無視するので、別に見る）。"""
    with zipfile.ZipFile(pptx) as z:
        return sorted(n for n in z.namelist() if n.endswith(".xml") and re.search(rb' descr="(/|[A-Za-z]:\\|~/|file:)', z.read(n)))


def differences(before, after):
    with zipfile.ZipFile(before) as za, zipfile.ZipFile(after) as zb:
        na, nb = set(za.namelist()), set(zb.namelist())
        # グラフに埋め込む xlsx は作るたびに中の日時が変わる。値はグラフの XML で比べる
        changed = [n for n in na & nb if "/embeddings/" not in n and not same_part(n, za.read(n), zb.read(n))]
        return sorted(changed + list(na ^ nb))


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--base", default="origin/main", help="比べる元のコミット（既定 origin/main）")
    parser.add_argument("--keep", action="store_true", help="作ったファイルを残す")
    args = parser.parse_args()

    modules = REPO / PLUGIN / "node_modules"
    if not (modules / "pptxgenjs").is_dir():
        print(f"pptxgenjs が無い。先に: (cd {PLUGIN} && npm install)", file=sys.stderr)
        return 2

    tmp = Path(tempfile.mkdtemp(prefix="deck-check-"))
    sides = {"before": tmp / "before", "after": tmp / "after"}
    archive = git("archive", args.base, str(ASSETS))
    with tarfile.open(fileobj=io.BytesIO(archive)) as tar:
        tar.extractall(tmp / "base-tree")
    shutil.copytree(tmp / "base-tree" / ASSETS, sides["before"] / "assets")
    shutil.copytree(REPO / ASSETS, sides["after"] / "assets", ignore=shutil.ignore_patterns("node_modules", "*.pptx"))

    (tmp / "home").mkdir()
    env = {k: v for k, v in os.environ.items() if k not in ("CLAUDE_CONFIG_DIR", "CODEX_HOME", "CLAUDE_PLUGIN_ROOT")}
    env.update(HOME=str(tmp / "home"), NODE_PATH=str(modules), PPTX_DECK_THEME_FILES="0")

    jobs = [(p.stem, None, p.name) for p in sorted((REPO / ASSETS).glob("sample-*.js"))]
    trials = sorted(p.parent for p in (REPO / "trials").glob("*/build.js"))
    jobs += [(f"trial-{t.name}", t, "build.js") for t in trials]

    for side, root in sides.items():
        (root / "out").mkdir()
        for name, trial, _ in jobs:
            if trial is None:
                continue
            work = root / name
            shutil.copytree(trial, work, ignore=shutil.ignore_patterns("out", "deck-parts.js", "*.pptx"))
            shutil.copy(root / "assets/deck-parts.js", work)
            for folder in ("icons", "themes"):
                if not (work / folder).exists() and (root / "assets" / folder).is_dir():
                    shutil.copytree(root / "assets" / folder, work / folder)

    changed, broken = [], []
    for name, trial, script in jobs:
        results = {}
        for side, root in sides.items():
            cwd = root / "assets" if trial is None else root / name
            if not (cwd / script).exists():
                results[side] = None  # 元のコミットに無い見本
                continue
            results[side] = build(cwd / script, cwd, root / "out" / f"{name}.pptx", env)
        if not results["after"][0]:
            broken.append(name)
            print(f"× {name}: 作れない\n{results['after'][2][-1500:]}")
            continue
        leaked = path_descr(sides["after"] / "out" / f"{name}.pptx")
        if leaked:
            broken.append(f"{name}: 代替テキストにファイルの場所")
            print(f"× {name}: 代替テキストにファイルの場所が入っている（{', '.join(leaked[:4])}）")
        if results["before"] is None or not results["before"][0]:
            changed.append(name)
            print(f"? {name}: 元の版では作れないので比べられない（警告 {results['after'][1]}件）")
            continue
        diff = differences(sides["before"] / "out" / f"{name}.pptx", sides["after"] / "out" / f"{name}.pptx")
        warnings = f"警告 {results['before'][1]}→{results['after'][1]}件"
        if diff or results["before"][1] != results["after"][1]:
            changed.append(name)
            slides = sorted((int(m[1]) for n in diff if (m := re.fullmatch(r"ppt/slides/slide(\d+)\.xml", n))))
            where = f"ページ {', '.join(map(str, slides))}" if slides else ", ".join(diff[:4]) or "警告の数だけ"
            print(f"≠ {name}: 違う（{where}）{warnings}")
        else:
            print(f"= {name}: 同じ（{warnings}）")

    if shutil.which("claude"):
        for target in (REPO, REPO / PLUGIN):
            result = subprocess.run(["claude", "plugin", "validate", str(target)], capture_output=True, text=True)
            label = target.relative_to(REPO.parent)
            if result.returncode:
                broken.append(f"validate {label}")
                print(f"× 検査に通らない: {label}\n{result.stdout}{result.stderr}")
            else:
                print(f"= 検査に通る: {label}")
    else:
        print("claude が無いので、プラグインの検査は飛ばした")

    manifest = PLUGIN / ".claude-plugin/plugin.json"
    old_version = json.loads(git("show", f"{args.base}:{manifest}"))["version"]
    new_version = json.loads((REPO / manifest).read_text(encoding="utf-8"))["version"]
    edited = git("diff", "--name-only", args.base, "--", str(PLUGIN)).strip()
    edited = edited or git("ls-files", "--others", "--exclude-standard", "--", str(PLUGIN)).strip()
    if edited and old_version == new_version:
        print(f"注意: プラグインを直したが version は {new_version} のまま。上げないと、入れた人に更新が届かない")

    if args.keep:
        print(f"作ったファイル: {tmp}")
    else:
        shutil.rmtree(tmp)
    print(f"結果: {len(jobs)}件中、違い {len(changed)}件、失敗 {len(broken)}件（比べた元: {args.base}）")
    return 2 if broken else 1 if changed else 0


if __name__ == "__main__":
    sys.exit(main())
