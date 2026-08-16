"""C1-F9（本命）: `ailine.main(argv)` を実際に叩き、標準出力の全文を承認テストで固定する。

★ 現在 ailine.main( を通るテストが test_ailine.py に 0 本だった（CLI の契約=main(argv)
そのものが一度も検証されていなかった）。ここでは translate_task / basrun_apply を
「録画済み」相当の固定戻り値・固定副作用の関数に差し替え、それ以外は本物のパイプライン
（verify_dsl_args → codegen_dsl → build_advisories → run_postcondition → ...）を通す。

シナリオ20+本（brief の目安どおり）:
  dsl: pass / warn / fail / 実行時エラー
  plan: 全ok / 混在warn / 途中fail / CLARIFY含み
  freeform: 関所y / 関所N / 非対話 / 総なめ検出
  破壊の関所: y / N / 非対話
  忠実度ゲート(exit 4) / Excelロック(exit 5) / runロック(exit 6)
  header-row 指定
  --dry × 3経路(dsl/plan/freeform)
= 4+4+4+3+1+1+1+1+3 = 22本。

ゴールデンは tests/golden/f9_transcripts/<name>.txt（標準出力そのもの）。
更新の作法は tests/golden/_harness.py 参照。
"""
import argparse
import json
import re
import sys
from pathlib import Path

import openpyxl
import pytest
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Font

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import ailine  # noqa: E402

from golden._harness import GOLDEN_ROOT, assert_golden_text  # noqa: E402

F9_DIR = GOLDEN_ROOT / "f9_transcripts"


def _book(tmp_path, rows, name="b.xlsx"):
    p = tmp_path / name
    wb = openpyxl.Workbook()
    ws = wb.active
    for r in rows:
        ws.append(r)
    wb.save(p)
    return p


def _isolate(monkeypatch, tmp_path):
    """本番のユーザーディレクトリ(~/.ailine 等)に一切触れないようにする（実行するたび
    実ファイルへ書き込む既存の危険な既定を、テストでは全部 tmp_path に寄せる）。"""
    monkeypatch.setattr(ailine, "HISTORY_FILE", tmp_path / "history.jsonl")
    monkeypatch.setattr(ailine, "BACKUP_DIR", tmp_path / "backups")
    monkeypatch.setattr(ailine, "VOCAB_FILE", tmp_path / "vocab.json")
    monkeypatch.setattr(ailine, "RUN_LOCK_FILE", tmp_path / "run.lock")
    monkeypatch.setattr(ailine, "normalize_book", lambda book, workdir, timeout=None: book)


def _run_main(argv, capsys) -> tuple:
    rc = ailine.main(argv)
    out = capsys.readouterr().out
    return rc, out


# ===========================================================================
# dsl: pass / warn / fail / 実行時エラー
# ===========================================================================

def _t_dsl_pass(tmp_path, monkeypatch, capsys):
    book = _book(tmp_path, [["商品", "金額"], ["a", 200], ["b", 300], ["c", 100]])
    monkeypatch.setattr(ailine, "translate_task",
                         lambda model, task, book_meta, temperature=0.1:
                         {"op": "SORT", "args": {"col": "金額", "order": "desc"}})

    def fake_apply(out_book, code, workdir, helper_files=(), timeout=None):
        wb = openpyxl.load_workbook(out_book)
        ws = wb.active
        for i, (name, val) in enumerate([("b", 300), ("a", 200), ("c", 100)], start=2):
            ws.cell(row=i, column=1, value=name)
            ws.cell(row=i, column=2, value=val)
        wb.save(out_book)
        return True, None, "ok"

    monkeypatch.setattr(ailine, "basrun_apply", fake_apply)
    return _run_main(["run", str(book), "金額で降順に並べ替えて", "--copy"], capsys)


def _t_dsl_warn(tmp_path, monkeypatch, capsys):
    book = _book(tmp_path, [["商品", "金額"], ["a", 100]])   # データ1行のみ→順序の意味なし
    monkeypatch.setattr(ailine, "translate_task",
                         lambda model, task, book_meta, temperature=0.1:
                         {"op": "SORT", "args": {"col": "金額", "order": "desc"}})
    monkeypatch.setattr(ailine, "basrun_apply",
                         lambda out_book, code, workdir, helper_files=(), timeout=None:
                         (True, None, "ok"))
    return _run_main(["run", str(book), "金額で降順に並べ替えて", "--copy"], capsys)


def _t_dsl_fail(tmp_path, monkeypatch, capsys):
    book = _book(tmp_path, [["商品", "金額"], ["a", 100], ["b", 300], ["c", 200]])   # 未整列のまま
    monkeypatch.setattr(ailine, "translate_task",
                         lambda model, task, book_meta, temperature=0.1:
                         {"op": "SORT", "args": {"col": "金額", "order": "desc"}})
    # ★ basrun_apply は「成功はしたが実際には並べ替わっていない」状態を模する
    #   （LibreOffice+LLM がもっともらしく成功報告しつつ何もしないケースの再現）。
    monkeypatch.setattr(ailine, "basrun_apply",
                         lambda out_book, code, workdir, helper_files=(), timeout=None:
                         (True, None, "ok"))
    return _run_main(["run", str(book), "金額で降順に並べ替えて", "--copy"], capsys)


def _t_dsl_runtime_error(tmp_path, monkeypatch, capsys):
    book = _book(tmp_path, [["商品", "金額"], ["a", 100], ["b", 300]])
    monkeypatch.setattr(ailine, "translate_task",
                         lambda model, task, book_meta, temperature=0.1:
                         {"op": "SORT", "args": {"col": "金額", "order": "desc"}})
    monkeypatch.setattr(
        ailine, "basrun_apply",
        lambda out_book, code, workdir, helper_files=(), timeout=None:
        (False, "BasicError: Object variable not set (line 12)", "raw-uno-output"))
    return _run_main(["run", str(book), "金額で降順に並べ替えて", "--copy"], capsys)


# ===========================================================================
# plan: 全ok / 混在warn / 途中fail / CLARIFY含み
# ===========================================================================

def _plan_book_presorted_bold(tmp_path):
    book = _book(tmp_path, [["商品", "金額"], ["a", 300], ["b", 200], ["c", 100]])
    wb = openpyxl.load_workbook(book)
    ws = wb.active
    for c in (1, 2):
        ws.cell(row=1, column=c).font = Font(bold=True)
    wb.save(book)
    return book


def _t_plan_all_ok(tmp_path, monkeypatch, capsys):
    book = _plan_book_presorted_bold(tmp_path)
    monkeypatch.setattr(ailine, "translate_task",
                         lambda model, task, book_meta, temperature=0.1:
                         {"plan": [{"op": "SORT", "args": {"col": "金額", "order": "desc"}},
                                   {"op": "BOLD", "args": {"target": "row:1"}}]})
    monkeypatch.setattr(ailine, "basrun_apply",
                         lambda out_book, code, workdir, helper_files=(), timeout=None:
                         (True, None, "ok"))
    return _run_main(["run", str(book), "金額で降順に並べ替えて見出しを太字に", "--copy"], capsys)


def _t_plan_mixed_warn(tmp_path, monkeypatch, capsys):
    book = _plan_book_presorted_bold(tmp_path)
    monkeypatch.setattr(ailine, "translate_task",
                         lambda model, task, book_meta, temperature=0.1:
                         {"plan": [{"op": "SORT", "args": {"col": "金額", "order": "desc"}},
                                   {"op": "OUT_OF_VOCAB", "about": "条件付き書式"}]})

    def fake_apply(out_book, code, workdir, helper_files=(), timeout=None):
        # ★ SORT 段は既に整列済みなのでファイルを変えなくても postcondition は通る。
        #   語彙外(自由生成)段は「変化した」ことそのものが唯一の機械確認(no-op ガード)
        #   なので、こちらだけ必ず何かしら変更する（code の中身で段を見分ける — 同じ
        #   basrun_apply 差し替えを両方の段が共有するため、SORT 側まで毎回書き換えると
        #   語彙外段の前後比較が『既に書き換え済み』で差分無し=no-op に化ける）。
        if "SortByColumn" not in code:
            wb = openpyxl.load_workbook(out_book)
            wb.active.cell(row=1, column=5, value="cf-applied")
            wb.save(out_book)
        return True, None, "ok"
    monkeypatch.setattr(ailine, "basrun_apply", fake_apply)

    def fake_ollama(model, msgs, temperature=0.2):
        return "Sub Run(oDoc As Object)\nEnd Sub"
    monkeypatch.setattr(ailine, "ollama_generate", fake_ollama)
    return _run_main(["run", str(book), "金額で降順に並べ替えて条件付き書式もつけて", "--copy",
                       "--allow-freeform"], capsys)


def _t_plan_partial_fail(tmp_path, monkeypatch, capsys):
    book = _book(tmp_path, [["商品", "金額"], ["a", 300], ["b", 200], ["c", 100]])   # 見出しは太字でない
    monkeypatch.setattr(ailine, "translate_task",
                         lambda model, task, book_meta, temperature=0.1:
                         {"plan": [{"op": "SORT", "args": {"col": "金額", "order": "desc"}},
                                   {"op": "BOLD", "args": {"target": "row:1"}}]})
    monkeypatch.setattr(ailine, "basrun_apply",
                         lambda out_book, code, workdir, helper_files=(), timeout=None:
                         (True, None, "ok"))   # BOLD 側は何もしない→太字にならず fail
    return _run_main(["run", str(book), "金額で降順に並べ替えて見出しを太字に", "--copy"], capsys)


def _t_plan_with_clarify(tmp_path, monkeypatch, capsys):
    book = _plan_book_presorted_bold(tmp_path)
    monkeypatch.setattr(ailine, "translate_task",
                         lambda model, task, book_meta, temperature=0.1:
                         {"plan": [{"op": "SORT", "args": {"col": "金額", "order": "desc"}},
                                   {"op": "CLARIFY", "question": "何色にしますか？"}]})
    monkeypatch.setattr(ailine, "basrun_apply",
                         lambda out_book, code, workdir, helper_files=(), timeout=None:
                         (True, None, "ok"))
    return _run_main(["run", str(book), "金額で降順に並べ替えて色も変えて", "--copy"], capsys)


# ===========================================================================
# freeform: 関所y / 関所N / 非対話 / 総なめ検出
# ===========================================================================

def _freeform_setup(monkeypatch, code="Sub Run(oDoc As Object)\nEnd Sub"):
    monkeypatch.setattr(ailine, "translate_task",
                         lambda model, task, book_meta, temperature=0.1:
                         {"op": "FREEFORM", "args": {}})
    monkeypatch.setattr(ailine, "ollama_generate",
                         lambda model, msgs, temperature=0.2: code)


def _t_freeform_gate_yes(tmp_path, monkeypatch, capsys):
    book = _book(tmp_path, [["商品", "金額"], ["a", 100]])
    _freeform_setup(monkeypatch)

    def fake_apply(out_book, code, workdir, helper_files=(), timeout=None):
        wb = openpyxl.load_workbook(out_book)
        wb.active.cell(row=1, column=3, value="備考")
        wb.save(out_book)
        return True, None, "ok"
    monkeypatch.setattr(ailine, "basrun_apply", fake_apply)
    monkeypatch.setattr("builtins.input", lambda prompt="": "y")
    return _run_main(["run", str(book), "何か列を足して", "--copy"], capsys)


def _t_freeform_gate_no(tmp_path, monkeypatch, capsys):
    book = _book(tmp_path, [["商品", "金額"], ["a", 100]])
    _freeform_setup(monkeypatch)
    monkeypatch.setattr("builtins.input", lambda prompt="": "n")
    return _run_main(["run", str(book), "何か列を足して", "--copy"], capsys)


def _t_freeform_noninteractive(tmp_path, monkeypatch, capsys):
    book = _book(tmp_path, [["商品", "金額"], ["a", 100]])
    _freeform_setup(monkeypatch)

    def _raise_eof(prompt=""):
        raise EOFError()
    monkeypatch.setattr("builtins.input", _raise_eof)
    return _run_main(["run", str(book), "何か列を足して", "--copy"], capsys)


def _t_freeform_helper_sweep_detected(tmp_path, monkeypatch, capsys):
    book = _book(tmp_path, [["商品", "金額"], ["a", 100]])
    # ★ detect_helper_sweep の閾値(4種以上)を満たすよう、実在ヘルパ名を5つ Call する。
    code = ("Sub Run(oDoc As Object)\n"
            "    Call AutoFitColumns(oDoc)\n"
            "    Call AlignCenter(oDoc, 0, 1)\n"
            "    Call FormatThousands(oDoc, 0, 1)\n"
            "    Call DrawTableBorders(oDoc)\n"
            "    Call StyleBold(oDoc, 0, 0, 1, 0)\n"
            "End Sub\n")
    _freeform_setup(monkeypatch, code=code)
    monkeypatch.setattr("builtins.input", lambda prompt="": "y")

    def fake_apply(out_book, code, workdir, helper_files=(), timeout=None):
        wb = openpyxl.load_workbook(out_book)
        wb.active.cell(row=1, column=3, value="x")
        wb.save(out_book)
        return True, None, "ok"
    monkeypatch.setattr(ailine, "basrun_apply", fake_apply)
    return _run_main(["run", str(book), "備考の列を足して", "--copy"], capsys)


# ===========================================================================
# 破壊の関所: y / N / 非対話
# ===========================================================================

def _overwrite_gate_setup(tmp_path):
    return _book(tmp_path, [["商品", "数量", "単価", "金額"],
                             ["a", 2, 100, 999], ["b", 3, 150, 999]])   # 金額に既存値あり


def _t_overwrite_gate_yes(tmp_path, monkeypatch, capsys):
    book = _overwrite_gate_setup(tmp_path)
    monkeypatch.setattr(ailine, "translate_task",
                         lambda model, task, book_meta, temperature=0.1:
                         {"op": "COMPUTE_COLUMN",
                          "args": {"operands": ["数量", "単価"], "operator": "*", "target": "金額"}})

    def fake_apply(out_book, code, workdir, helper_files=(), timeout=None):
        wb = openpyxl.load_workbook(out_book)
        ws = wb.active
        ws.cell(row=2, column=4, value=200)
        ws.cell(row=3, column=4, value=450)
        wb.save(out_book)
        return True, None, "ok"
    monkeypatch.setattr(ailine, "basrun_apply", fake_apply)
    monkeypatch.setattr("builtins.input", lambda prompt="": "y")
    # ★ 破壊の関所は --copy 時は走らない（原本に触れないため確認不要・
    #   test_confirm_overwrite_or_gate_none_when_copy_mode の実装どおり）。
    #   関所そのものを見るシナリオなので --copy は付けない（既定=原本直接反映）。
    return _run_main(["run", str(book), "金額を数量×単価で上書きして", "--values"], capsys)


def _t_overwrite_gate_no(tmp_path, monkeypatch, capsys):
    book = _overwrite_gate_setup(tmp_path)
    monkeypatch.setattr(ailine, "translate_task",
                         lambda model, task, book_meta, temperature=0.1:
                         {"op": "COMPUTE_COLUMN",
                          "args": {"operands": ["数量", "単価"], "operator": "*", "target": "金額"}})
    monkeypatch.setattr("builtins.input", lambda prompt="": "n")
    return _run_main(["run", str(book), "金額を数量×単価で上書きして", "--values"], capsys)


def _t_overwrite_gate_noninteractive(tmp_path, monkeypatch, capsys):
    book = _overwrite_gate_setup(tmp_path)
    monkeypatch.setattr(ailine, "translate_task",
                         lambda model, task, book_meta, temperature=0.1:
                         {"op": "COMPUTE_COLUMN",
                          "args": {"operands": ["数量", "単価"], "operator": "*", "target": "金額"}})

    def _raise_eof(prompt=""):
        raise EOFError()
    monkeypatch.setattr("builtins.input", _raise_eof)
    return _run_main(["run", str(book), "金額を数量×単価で上書きして", "--values"], capsys)


# ===========================================================================
# 忠実度ゲート(exit 4) / Excel ロック(exit 5) / run ロック(exit 6)
# ===========================================================================

def _t_fidelity_gate(tmp_path, monkeypatch, capsys):
    book = _book(tmp_path, [["商品", "金額"], ["a", 100], ["b", 200]])
    ws = openpyxl.load_workbook(book).active
    wb = ws.parent
    ws.conditional_formatting.add("B2:B3", CellIsRule(operator="greaterThan", formula=["150"]))
    wb.save(book)

    def fake_normalize(b, workdir, timeout=None):
        norm = workdir / ("normalized" + b.suffix)
        w = openpyxl.Workbook()
        s = w.active
        s.append(["商品", "金額"])
        s.append(["a", 100])
        s.append(["b", 200])   # 正規化後は CF が消えている(往復忠実度ロス)
        w.save(norm)
        return norm
    monkeypatch.setattr(ailine, "normalize_book", fake_normalize)
    monkeypatch.setattr(ailine, "translate_task",
                         lambda model, task, book_meta, temperature=0.1:
                         {"op": "SORT", "args": {"col": "金額", "order": "asc"}})
    return _run_main(["run", str(book), "金額で並べ替えて"], capsys)   # --copy なし=既定(原本直接)


def _t_excel_lock(tmp_path, monkeypatch, capsys):
    book = _book(tmp_path, [["商品", "金額"], ["a", 1]])
    (tmp_path / f"~${book.name}").write_bytes(b"lock")
    return _run_main(["run", str(book), "何か変更して", "--copy"], capsys)


def _t_run_lock_busy(tmp_path, monkeypatch, capsys):
    book = _book(tmp_path, [["商品", "金額"], ["a", 1]])
    lock_path = tmp_path / "run.lock"
    other_pid = 999999
    # ★ ts は「30分以内」でないと _lock_is_stale が奪取可能と判定してしまう
    #   （固定の過去日時だと実測で stale 扱いされ、ロック無視で先へ進んでしまった実測
    #   バグをここで踏んだ）。実行時刻を使いつつ、golden 側では <TS> に正規化して
    #   時刻依存性を消す（下の test_transcript_golden の redaction 参照）。
    ts = ailine.datetime.now(ailine.timezone.utc).isoformat(timespec="seconds")
    lock_path.write_text(json.dumps({"pid": other_pid, "ts": ts}), encoding="utf-8")
    monkeypatch.setattr(ailine, "_pid_alive", lambda pid: pid == other_pid)
    return _run_main(["run", str(book), "何か変更して", "--copy"], capsys)


# ===========================================================================
# header-row 指定
# ===========================================================================

def _t_header_row_explicit(tmp_path, monkeypatch, capsys):
    book = _book(tmp_path, [["x", 1], ["y", 2], ["商品", "金額"], ["a", 300], ["b", 100]])
    ambiguous = {"sheets": {"Sheet": {"rows": {
        1: {"nonempty": 2, "str": 2, "bold": 0}, 2: {"nonempty": 2, "str": 1, "bold": 0},
        3: {"nonempty": 2, "str": 2, "bold": 0}, 4: {"nonempty": 2, "str": 1, "bold": 0},
    }}}}
    monkeypatch.setattr(ailine, "build_struct_dump", lambda b, workdir: ambiguous)
    monkeypatch.setattr(ailine, "translate_task",
                         lambda model, task, book_meta, temperature=0.1:
                         {"op": "SORT", "args": {"col": "金額", "order": "desc"}})
    monkeypatch.setattr(ailine, "basrun_apply",
                         lambda out_book, code, workdir, helper_files=(), timeout=None:
                         (True, None, "ok"))
    return _run_main(["run", str(book), "金額で降順に並べ替えて", "--copy", "--header-row", "3"],
                      capsys)


# ===========================================================================
# --dry × 3経路
# ===========================================================================

def _t_dry_dsl(tmp_path, monkeypatch, capsys):
    book = _book(tmp_path, [["商品", "金額"], ["a", 100], ["b", 200]])
    monkeypatch.setattr(ailine, "translate_task",
                         lambda model, task, book_meta, temperature=0.1:
                         {"op": "SORT", "args": {"col": "金額", "order": "desc"}})
    return _run_main(["run", str(book), "金額で降順に並べ替えて", "--dry"], capsys)


def _t_dry_plan(tmp_path, monkeypatch, capsys):
    book = _book(tmp_path, [["商品", "金額"], ["a", 100], ["b", 200]])
    monkeypatch.setattr(ailine, "translate_task",
                         lambda model, task, book_meta, temperature=0.1:
                         {"plan": [{"op": "SORT", "args": {"col": "金額", "order": "desc"}},
                                   {"op": "BOLD", "args": {"target": "row:1"}}]})
    return _run_main(["run", str(book), "金額で降順に並べ替えて見出しを太字に", "--dry"], capsys)


def _t_dry_freeform(tmp_path, monkeypatch, capsys):
    book = _book(tmp_path, [["商品", "金額"], ["a", 100], ["b", 200]])
    _freeform_setup(monkeypatch)
    return _run_main(["run", str(book), "何か列を足して", "--dry"], capsys)


CASES = {
    "dsl_pass": _t_dsl_pass,
    "dsl_warn": _t_dsl_warn,
    "dsl_fail": _t_dsl_fail,
    "dsl_runtime_error": _t_dsl_runtime_error,
    "plan_all_ok": _t_plan_all_ok,
    "plan_mixed_warn": _t_plan_mixed_warn,
    "plan_partial_fail": _t_plan_partial_fail,
    "plan_with_clarify": _t_plan_with_clarify,
    "freeform_gate_yes": _t_freeform_gate_yes,
    "freeform_gate_no": _t_freeform_gate_no,
    "freeform_noninteractive": _t_freeform_noninteractive,
    "freeform_helper_sweep_detected": _t_freeform_helper_sweep_detected,
    "overwrite_gate_yes": _t_overwrite_gate_yes,
    "overwrite_gate_no": _t_overwrite_gate_no,
    "overwrite_gate_noninteractive": _t_overwrite_gate_noninteractive,
    "fidelity_gate_exit4": _t_fidelity_gate,
    "excel_lock_exit5": _t_excel_lock,
    "run_lock_exit6": _t_run_lock_busy,
    "header_row_explicit": _t_header_row_explicit,
    "dry_dsl": _t_dry_dsl,
    "dry_plan": _t_dry_plan,
    "dry_freeform": _t_dry_freeform,
}


_ISO_TS_RE = re.compile(r"\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\+\d{2}:\d{2}")


@pytest.mark.parametrize("name", sorted(CASES.keys()))
def test_transcript_golden(tmp_path, monkeypatch, capsys, name):
    _isolate(monkeypatch, tmp_path)
    rc, out = CASES[name](tmp_path, monkeypatch, capsys)
    # ★ run_lock_exit6 だけ「実行中ロックの発行時刻」が実行のたび変わる現在時刻を含む
    #   （_lock_is_stale が30分超で奪取可能と判定するため、固定の過去日時は使えない）。
    #   golden の再現性のため ISO タイムスタンプを一律 <TS> に正規化してから比較する。
    out = _ISO_TS_RE.sub("<TS>", out)
    # rc をトランスクリプトの先頭に埋め込む（終了コードの回帰もこの1ファイルで拾える）。
    text = f"[exit code: {rc}]\n{out}"
    assert_golden_text(F9_DIR / f"{name}.txt", text, label=name)
