# -*- coding: utf-8 -*-
"""CI と同じ素の環境を、手元で再現して走らせる。

★ なぜ在るか: 「手元に在って CI に無いもの」で CI が落ちる事故が、この repo で
4 度起きている ── lxml（requirements-dev.txt に顛末が書いてある）・ollama・
LibreOffice・Pillow（2026-08-24）。どれも**手元では緑**だった。
「居るから見えない」── 在るものは、依存していることを教えてくれない。

★ 処方は requirements-dev.txt に既に書かれていた:
  (a) 手元に合わせて CI に入れる → 差が消えるのではなく**見えなくなる**
  (b) 素の環境が常に被覆されるようにする → こちらを採る
この script は (b) を**押す前に**やる。

やること: requirements-dev.txt に書いた依存（とその依存）だけを import 可能にし、
それ以外の import を ImportError にした上で `pytest -m "not local"` を走らせる。
"""
from __future__ import annotations

import re
import subprocess
import sys
from importlib import metadata
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
NL = chr(10)


def _declared() -> set:
    names = {"pytest", "pip", "setuptools", "wheel", "_pytest"}
    # ★ repo 自身のモジュール（src/ 配下）は当然 CI にも在る ── 常に許可する。
    #   初版はこれを忘れて ailine_core を止めた（遮断器が働きすぎた）。
    #   tests/ 配下も同じ（検体同士が import し合う: golden / _run_argv 等）。
    for base in (ROOT / "src", ROOT / "tests"):
        if not base.is_dir():
            continue
        for child in base.iterdir():
            if child.is_dir() and not child.name.startswith("__"):
                names.add(child.name.lower())
            elif child.suffix == ".py":
                names.add(child.stem.lower())
    req = ROOT / "requirements-dev.txt"
    if req.exists():
        for line in req.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if line and not line.startswith("#"):
                names.add(re.split(r"[<>=\[;]", line)[0].strip().lower())
    return names


def _closure(seed: set) -> set:
    """宣言した依存の、依存の、依存…まで辿って許可集合を作る。"""
    allowed, todo = set(seed), list(seed)
    while todo:
        name = todo.pop()
        try:
            requires = metadata.requires(name) or []
        except metadata.PackageNotFoundError:
            continue
        for spec in requires:
            if "extra ==" in spec:          # optional extras は CI も入れない
                continue
            dep = re.split(r"[<>=\[;(\s]", spec)[0].strip().lower()
            if dep and dep not in allowed:
                allowed.add(dep)
                todo.append(dep)
    return allowed


def _top_levels(allowed: set) -> set:
    """配布名 → import 名（top_level.txt / モジュール名）に直す。"""
    # ★ 配布名の表記ゆれ（et-xmlfile / et_xmlfile）で取りこぼした実測あり ──
    #   比較は両側を正規化してから（- と _ を同一視する）。
    norm = lambda n: (n or "").lower().replace("-", "_")
    allowed = {norm(a) for a in allowed}
    mods = set(allowed)
    for dist in metadata.distributions():
        name = norm(dist.metadata["Name"])
        if name not in allowed:
            continue
        try:
            text = dist.read_text("top_level.txt") or ""
        except Exception:
            text = ""
        mods.update(t.strip() for t in text.splitlines() if t.strip())
        mods.add(name)
    return mods


def main() -> int:
    allowed = _top_levels(_closure(_declared()))
    runner = ROOT / 'scripts' / '_ci_parity_entry.py'
    runner.write_text(
        'import sys' + NL +
        'sys.path.insert(0, ' + repr(str(ROOT / 'scripts')) + ')' + NL +
        'import _ci_parity_blocker as b' + NL +
        'sys.exit(b.run(' + repr(sorted(allowed)) + '))' + NL,
        encoding='utf-8')
    try:
        import os
        env = dict(os.environ)
        env['PYTHONPATH'] = str(ROOT / 'src')
        # ★ CI は openpyxl+pytest しか入れないので、第三者 pytest プラグインは
        #   存在しない。手元では site-packages から自動読み込みされて anyio 等を
        #   引くため、遮断器が正しく怒る ── 自動読み込みを切って条件を揃える。
        env['PYTEST_DISABLE_PLUGIN_AUTOLOAD'] = '1'
        # ★★ 2026-08-28（CI を 2 度赤くして気づいた・「居るから見えない」の 6 度目）:
        #   ここが遮断していたのは**宣言外のパッケージ**だけだった。
        #   ・俺のホームに溜まった状態（~/.ailine の初回告知の印）
        #   ・手元で動いている ollama
        #   はどちらも素通りで、それに寄りかかった検体が手元で緑・CI で赤になった。
        #   ★ CI に無い物は**全部**無いことにする。ここに在る物に寄りかかった検体は、
        #     `-m local`（実物が要る側）へ置くのが この repo の作法。
        import tempfile
        env['OLLAMA_HOST'] = 'http://127.0.0.1:9'      # 使われないポート（必ず届かない）
        with tempfile.TemporaryDirectory(prefix='ailine_ci_home_') as home:
            env['AILINE_HOME'] = home
            return subprocess.run([sys.executable, str(runner)], cwd=str(ROOT),
                                   env=env).returncode
    finally:
        runner.unlink(missing_ok=True)


if __name__ == "__main__":
    sys.exit(main())
