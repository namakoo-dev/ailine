# 終了コードの分離 ── 実装より先に凍結した赤い検体。
# 出典: ABSORB-20260823-microsoft.md の副産物（実測）:
#   「sys.exit("...")（openpyxl 欠落・basrun 不在・ollama 不通・LO 起動失敗）と
#    検証 ⚠ の return 1 が全部 exit 1 に潰れている。CI や自動化から
#    『⚠ が出た』と『道具が壊れた』が区別できない」
#
# 契約:
#   ① 実行の前提が満たされていない（依存の欠落・外部プログラムに繋がらない・入力が無い）
#      → **exit 9**。既存の 1(失敗) 2(argparse) 3(CLARIFY) 4(忠実度) 5(verify) 6(ロック)
#      7(上書き関所) 8(自由生成の関所) とは別の意味を持たせる
#   ② 検証の失敗（適用したが事後条件を満たさない等）は従来どおり **exit 1**（意味を変えない）
#   ③ README に終了コードの表が 1 つある（k1LoW 台帳 ①-5・散らばった記述を 1 箇所へ）

import re
import sys
from pathlib import Path

import openpyxl
import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
import ailine  # noqa: E402

sys.path.insert(0, str(Path(__file__).resolve().parent))
from test_golden_transcripts import _isolate, _run_main  # noqa: E402

def _run_exit(argv, capsys):
    """★ 治具: 実挙動は SystemExit(コード) で落ちる（CLI として正しい）。OS が見る終了
       コードと同じものを取るため、ここで拾って番号を返す。assert 側は不変。"""
    try:
        rc = ailine.main(argv)
    except SystemExit as e:
        rc = e.code if isinstance(e.code, int) else 1
    return rc, capsys.readouterr().out


needs_impl = pytest.mark.xfail(
    not hasattr(ailine, "EXIT_ENVIRONMENT"),
    reason="終了コードの分離 未実装（契約は凍結済み）",
    strict=True,
)


@needs_impl
def test_missing_input_file_is_environment_not_generic_failure(tmp_path, monkeypatch, capsys):
    """存在しない文書を渡した ── 前提の不備なので 9（1 に潰さない）。"""
    _isolate(monkeypatch, tmp_path)
    rc, out = _run_exit(["run", str(tmp_path / "無い.xlsx"), "並べ替えて"], capsys)
    assert rc == ailine.EXIT_ENVIRONMENT == 9, f"exit={rc}: {out}"


@needs_impl
def test_basrun_missing_is_environment(tmp_path, monkeypatch, capsys):
    """basrun.py が見つからない ── 道具が無い。9。"""
    _isolate(monkeypatch, tmp_path)
    monkeypatch.setattr(ailine, "_find_basrun_path", lambda: None)
    # ★ 治具: 見出し検出が確信を持てる表にする（1 行だけだと基準に届かず CLARIFY 3 で
    #   basrun まで到達しない ── 測りたい経路の手前で止まっては検体にならない）
    p = tmp_path / "b.xlsx"
    wb = openpyxl.Workbook(); ws = wb.active
    ws.append(["商品", "金額"]); ws.append(["b", 50]); ws.append(["a", 100])
    wb.save(p)
    monkeypatch.setattr(
        ailine, "translate_task",
        lambda model, task, book_meta, temperature=0.1:
        {"op": "SORT", "args": {"col": "金額", "order": "desc"}})
    rc, out = _run_exit(["run", str(p), "金額で降順に並べ替えて", "--copy"], capsys)
    assert rc == 9, f"exit={rc}: {out}"


def test_verification_failure_stays_exit_1(tmp_path, monkeypatch, capsys):
    # xfail 対象外: 現状も緑（分離を入れても 1 の意味が保たれることの番人）
    """★ 意味を変えない側の番人: 適用したが事後条件を満たさない ── 従来どおり 1。"""
    _isolate(monkeypatch, tmp_path)
    p = tmp_path / "b.xlsx"
    wb = openpyxl.Workbook(); ws = wb.active
    ws.append(["商品", "金額"]); ws.append(["b", 50]); ws.append(["a", 100])
    wb.save(p)
    monkeypatch.setattr(
        ailine, "translate_task",
        lambda model, task, book_meta, temperature=0.1:
        {"op": "SORT", "args": {"col": "金額", "order": "desc"}})
    monkeypatch.setattr(ailine, "basrun_apply",
                         lambda *a, **k: (True, None, "ok"))   # 適用したが何も変えない
    rc, out = _run_main(["run", str(p), "金額で降順に並べ替えて", "--copy"], capsys)
    assert rc == 1, f"検証の失敗が 1 でなくなった（意味を壊した）: exit={rc}\n{out}"


@needs_impl
def test_readme_has_one_exit_code_table():
    """③: 終了コードの意味が docs/ENGINEERING.md の 1 つの表に集まっている。
       ★ 2026-08-27 に提出用 README を新設した際、表ごとこちらへ移した。"""
    t = (REPO / "docs" / "ENGINEERING.md").read_text(encoding="utf-8")
    for code in ("1", "3", "4", "5", "6", "7", "8", "9"):
        assert re.search(rf"\|\s*{code}\s*\|", t), f"終了コード表に {code} の行が無い"
