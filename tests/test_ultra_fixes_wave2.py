# 完成度レビュー第二波 ── 基盤 1（実 home 汚染の構造遮断）+ 本家 3 件 + 中 2。
# 実装より先に凍結した赤い検体。出典: SEALED-20260823-jisaku-ultra.md + 本家 ultra 所見。
#
# 契約:
#   ① AILINE_HOME 環境変数で全ホームファイル（history/vocab/misclass/aliases/run.lock/backups）
#      の親を差し替えられる ── subprocess 起動のテストにも env 継承で届く構造の隔離
#      （monkeypatch は同一プロセスにしか効かない・14 ファイルの subprocess テストが素通りしていた）
#   ② _COLUMN_ARG_KEYS に CHART の category_col（本家: 依存つき連鎖 fallback の 4 本目の配線）
#   ③ CSV 制御文字の二重報告の解消（1 件は 1 件・⚠ 行も 1 行）
#   ④ residue の resolved_args を長さ降順で消費（本家: 断片漏れ→偽の残差注記）
#   ⑤ ops 表に英字 op 名の列（alias add の誘導先が実在する答えを持つ）

import json
import os
import subprocess
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
import ailine  # noqa: E402

sys.path.insert(0, str(Path(__file__).resolve().parent))
from test_golden_transcripts import _book, _isolate, _run_main  # noqa: E402

needs_impl = pytest.mark.xfail(
    "AILINE_HOME" not in (ailine.__doc__ or "") and not hasattr(ailine, "resolve_home_dir"),
    reason="第二波 未実装（契約は凍結済み・実装が来たら自動実測化）",
    strict=True,
)


# --- ① AILINE_HOME による構造の隔離 ------------------------------------------------

@needs_impl
def test_subprocess_respects_ailine_home(tmp_path):
    """subprocess 起動でも AILINE_HOME 配下に書く ── 実 ~/.ailine に 1 バイトも触らない。"""
    home = tmp_path / "home"
    env = dict(os.environ)
    env["AILINE_HOME"] = str(home)
    # ★ 恒真回避: 読むだけの history でなく、確実に書く経路（csv 変換 = history 追記+
    #   run.lock 取得）で測る。実 home のファイル内容ハッシュの前後一致が契約。
    csvp = tmp_path / "in.csv"
    csvp.write_bytes(b"a,b\n1,2\n")
    real = Path.home() / ".ailine"
    import hashlib
    def _state():
        if not real.exists():
            return {}
        return {p.name: hashlib.sha256(p.read_bytes()).hexdigest()
                 for p in real.glob("*") if p.is_file()}
    before = _state()
    r = subprocess.run([sys.executable, "-m", "ailine", "csv", str(csvp)],
                        capture_output=True, text=True, env=env, cwd=str(tmp_path),
                        encoding="utf-8", errors="replace", timeout=120)
    assert r.returncode == 0, r.stdout + r.stderr
    after = _state()
    assert before == after, f"subprocess が実 home に触った: { {k for k in set(before)|set(after) if before.get(k)!=after.get(k)} }"
    assert (home / "history.jsonl").exists(), "AILINE_HOME 側に history が書かれていない"


@needs_impl
def test_run_lock_lives_under_ailine_home(tmp_path, monkeypatch):
    """run.lock も AILINE_HOME 配下（並行 pytest の相互妨害の根治）。"""
    home = tmp_path / "home"
    monkeypatch.setenv("AILINE_HOME", str(home))
    resolved = ailine.resolve_home_dir()
    assert str(resolved).startswith(str(home)), f"AILINE_HOME が効いていない: {resolved}"


# --- ② _COLUMN_ARG_KEYS の category_col（本家 bug_005）----------------------------

@needs_impl
def test_column_arg_keys_covers_chart_category_col():
    assert "category_col" in ailine._COLUMN_ARG_KEYS.get("CHART", ()), \
        "依存つき連鎖 fallback の 4 本目の配線（category_col）が無い"


# --- ③ 制御文字の二重報告（本家 bug_001）------------------------------------------

@needs_impl
def test_control_char_reported_exactly_once(tmp_path, monkeypatch, capsys):
    """制御文字 1 件 ── ⚠ は 1 行・「⚠ N 件」の N も 1。"""
    _isolate(monkeypatch, tmp_path)
    p = tmp_path / "in.csv"
    p.write_bytes("メモ\nok\na\x01b\n".encode("utf-8"))
    rc, out = _run_main(["csv", str(p)], capsys)
    warn_lines = [ln for ln in out.splitlines() if "制御文字" in ln and "⚠" in ln]
    assert len(warn_lines) == 1, f"制御文字 1 件が {len(warn_lines)} 行になっている: {warn_lines}"
    assert "⚠ 2 件" not in out, f"1 件を 2 件と数えている: {out}"


# --- ④ residue の長さ順消費（本家 bug_008）----------------------------------------

@needs_impl
def test_residue_consumes_longer_args_first():
    """args の値が重なる形（商品 ⊂ 商品コード）── key の順序に関わらず断片を残さない。"""
    from ailine_core import residue
    task = "商品コードで単価表と突合して単価を入れて"
    pool = ("転記", "引っ張ってくる", "突合", "単価")
    for args in ({"key_col": "商品", "target_col": "商品コード", "source_sheet": "単価表"},
                  {"target_col": "商品コード", "key_col": "商品", "source_sheet": "単価表"}):
        left = residue.find_unconsumed_words(task, resolved_args=args, pool_phrases=pool)
        assert "コード" not in (left or []), \
            f"順序 {list(args)} で断片『コード』が残差に漏れた（偽の残差注記）: {left}"


# --- ⑤ ops 表の英字 op 名（誘導先の実在・自作 dim4）--------------------------------

@needs_impl
def test_ops_table_shows_op_codes(tmp_path, monkeypatch, capsys):
    """ailine ops の出力に SORT 等の英字 op 名が現れる（alias add の誘導が解決する）。"""
    _isolate(monkeypatch, tmp_path)
    rc, out = _run_main(["ops"], capsys)
    assert rc == 0
    assert "SORT" in out and "DEDUP" in out, f"英字 op 名が表に無い: {out[:500]}"
