# SQL との比較が「測ったこと」に留まっていること（2026-09-04）。
#
# ★★ なぜ在るか: `docs/なぜこの形か.md` は当初「SQL はほぼ全部の軸でこれより上」と書き、
#   検証の難しさを「表計算にスキーマが無いから」で説明していた。★ **軸の置き方が
#   結論を決めていた** ── 集合演算・トランザクション・規模はどれも DB の得意分野で、
#   その軸を置いた時点で SQL が勝つのは自明。
#   実測したら、**スキーマが在っても同じ穴が開く**ことが分かった（CHECK 制約を全部
#   通りながら、金額 ≠ 件数 × 単価 の表が残る）。
#
# ★ 主張を強めた以上、**その主張が測ったことのままである**ことを機械が守る:
#   ① 実験が実際に走ること（走らない主張は主張ではない）
#   ② 文書に書いた数字が、実験の出力と一致すること
#   ③ ★ 誇張していないこと ── DB が上である 4 つの軸を、文書が今も認めていること

import json
import re
import subprocess
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
DEMO = REPO / "bench" / "sql_invariant_demo.py"
DOC = REPO / "docs" / "なぜこの形か.md"


def _run() -> dict:
    r = subprocess.run([sys.executable, str(DEMO), "--json"], cwd=str(REPO),
                        capture_output=True, text=True, encoding="utf-8")
    assert r.returncode == 0, r.stderr[-600:]
    return json.loads(r.stdout)


def test_the_demonstration_actually_runs():
    """① 実験が走ること。★ 走らない主張は主張ではない。

    ★ 依存は標準ライブラリの sqlite3 だけ（宣言外の依存を足していない）。
    """
    d = _run()
    assert d["sql_said"] == "2 行更新しました"
    assert d["constraints_passed"] is True, "CHECK を破ってしまっている（実験が成立しない）"


def test_the_schema_did_not_catch_it():
    """★ 核心 ── 制約を全部通りながら、表が矛盾していること。

    ★ ここが崩れたら主張の土台が消える。「DB でも同じ穴が開く」と書けなくなる。
    """
    d = _run()
    assert d["rows_inconsistent"] == ["丸和物流", "みどり建設"], d["rows_inconsistent"]
    assert d["identities_found"], "操作前の等式を 1 つも見つけられていない（★ 陽性対照）"
    assert d["identities_broken"], "崩れた等式を検出できていない"
    assert "成り立たなくなりました" in (d["ailine_said"] or "")


def test_the_document_matches_the_demonstration():
    """② 文書の数字が実験と一致すること。★ 人が書いた数は必ず古くなる。"""
    t = DOC.read_bytes().decode("utf-8")
    d = _run()
    assert f"「{d['sql_said']}」" in t, "SQL の応答が文書とずれている"
    for name in d["rows_inconsistent"]:
        assert name in t, f"矛盾した行 {name} が文書に出ていない"
    assert "bench/sql_invariant_demo.py" in t, "実験への導線が文書から消えた"


def test_the_document_still_concedes_what_db_is_better_at():
    """③ ★ 誇張していないこと。

    ★ 「SQL にも同じ層が要る」は言えるが、「SQL より上」は言えない。
      DB が上である軸を文書が認めなくなったら、それは誇張になっている。
    """
    t = DOC.read_bytes().decode("utf-8")
    for axis in ("集合演算", "トランザクション", "同時実行", "規模"):
        assert axis in t, f"DB が上である軸『{axis}』を文書が認めなくなった"
    assert "SQL より上" not in t.replace("「SQL より上」という話ではありません", ""), (
        "『SQL より上』と読める主張が入っている")
    assert "SQL にも同じ層が要る" in t, "主張の芯（SQL にも同じ層が要る）が消えた"
