"""dsl_step — C7: 単発 DSL 経路(ailine.py cmd_run_dsl)と複合計画の DSL 段
   (ailine.py cmd_run_plan/_run_dsl_plan_step)が共有する実行エンジン。

★ なぜ（C7 ブリーフより）: 「単発は動くが複合が壊れる」という階級は、器官（接地検証・
破壊の関所・助言・切り詰め注記…）が経路ごとに別々に実装されていたことから生まれた
（scan_rate_literals が複合段にだけ無かった W10f、_truncation_notice が複合の DSL 段に
だけ無い未修正の穴、等）。「単発 = 1 段の計画」— 確認行の印字から事後条件チェックまでの
一連の流れを、単発/複合のどちらから呼ばれても同じコードを通すことで、この階級を構造的に
無くす。

★ ここに移した範囲（と、あえて移さなかった範囲）: ③確認行〜破壊の関所(print_dsl_confirmation)
と⑤適用〜⑥事後条件(apply_dsl_step)── どちらも単発/複合計画で完全に同じ手順を踏む部分。
②検証(verify_dsl_args・複合計画は依存つき連鎖の新規列フォールバックが加わる)と、助言の
組み立て(build_advisories と _structural_advisories+unrequested_new_sheet_advisory は
中身が違う)は、呼び出し側(ailine.py)にそのまま残した ── 単発は元々1呼び出しで済み、
複合計画側の追加ロジックも数行で、共有関数に切り出すより ailine.py 側に残すほうが
シンプルだった（line budget（ailine.py 5466 行）を尊重する観点でも、抽象化の層を
増やすより既存の関数を直接呼ぶほうが正味の行数が少ない）。

★★ 出力は1バイトも変えない（純リファクタ）: 単発と複合計画は、印字の分量・タイミング・
`--json` の型など複数の観測可能な違いを既に持っている（例: 複合計画の DSL 段は生成した
.bas を印字しない・変更点の行を印字しない・--ask を効かせない・_truncation_notice を
呼ばない）。これらは (c) 挙動変更であり、この回では**選ばず・報告する**（C7 ブリーフの
最重要の縛り）。このモジュールはその違いを**明示的な引数**として運ぶ ── 「どちらかに
揃える」のではなく「今の違いをそのままパラメータ化する」。将来どちらかを直したくなったら
呼び出し側の1引数を変えるだけで済む、というのがこの統合の実利（詳細は
docs/behavior-corpus/nodes/dsl-pipeline.md 参照）。

★ 置き場所: C4/C5/C6 に倣い ailine_core/ に置く。ailine.py 側の関数（format_confirmation_line
/ basrun_apply / run_postcondition 等）は import せず、呼び出し側が DslStepDeps に**呼び出し
時点で**（モジュール読み込み時点ではなく）詰めて渡す ── モジュール読み込み時点で束ねると、
テストの monkeypatch（`monkeypatch.setattr(ailine, "basrun_apply", ...)`）より先に関数
参照を固定してしまい、差し替えが効かなくなる（golden/transcript テストが軒並み壊れる）。
呼び出し側(ailine.py の cmd_run_dsl / _run_dsl_plan_step)が実行のたびに DslStepDeps を
組み立てることで、その時点の（monkeypatch 後の）ailine.py モジュール属性を拾う。
"""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Callable


@dataclass
class DslStepDeps:
    """ailine.py 側の関数参照の束。呼び出し側が呼び出しのたびに組み立てる（上記docstring参照）。"""
    format_confirmation_line: Callable
    maybe_warn_header_col_mismatch: Callable
    maybe_warn_target_overwrite: Callable
    interpretation_summary_line: Callable
    confirm_overwrite_or_gate: Callable
    basrun_apply: Callable
    snapshot: Callable
    diff_snapshots: Callable
    run_postcondition: Callable
    progress_start: Callable
    progress_end: Callable
    pivot_caveat: str
    verify_dsl_args: Callable
    apply_new_column_fallback: Callable
    build_advisories: Callable
    structural_advisories: Callable
    unrequested_new_sheet_advisory: Callable
    neutralize_new_column_ghost_warning: Callable
    neutralize_declared_new_sheet_warning: Callable
    neutralize_declared_sheet_replace_warning: Callable


@dataclass
class DslGroundResult:
    """resolve_dsl_step_args の戻り値。ok が False なら resolved 以下は無効。"""
    ok: bool
    resolved: dict | None
    inferred: set | None
    err: str | None
    new_cols: list


def resolve_dsl_step_args(op: str, raw_args: dict, task: str, meta: dict, vocab: dict, *,
                           original_headers: dict | None = None, first_sheet: str | None = None,
                           deps: DslStepDeps) -> DslGroundResult:
    """②検証（grounding）。単発(cmd_run_dsl)は original_headers/first_sheet=None のまま
       呼ぶ（new_cols は常に空になり、_apply_new_column_fallback は無条件でスキップされる
       ── 新規列の概念自体が単発には無いため、この呼び方で単発の従来挙動と完全一致する）。
       複合計画は直前までの段の original_headers/current_meta/first_sheet を渡し、
       『直前段が作った新規列』への依存つき連鎖フォールバックを1回だけ試みる。"""
    new_cols: list = []
    if first_sheet and original_headers is not None:
        new_cols = [c for c in meta["headers"].get(first_sheet, [])
                    if c not in original_headers.get(first_sheet, [])]
    ok, resolved, inferred, err = deps.verify_dsl_args(op, raw_args, meta, task=task, vocab=vocab)
    if not ok and new_cols and first_sheet:
        patched = deps.apply_new_column_fallback(
            op, raw_args, meta["headers"].get(first_sheet, []), new_cols)
        if patched != raw_args:
            ok2, resolved2, inferred2, err2 = deps.verify_dsl_args(
                op, patched, meta, task=task, vocab=vocab)
            if ok2:
                ok, resolved, inferred, err = ok2, resolved2, inferred2, err2
    return DslGroundResult(ok=ok, resolved=resolved, inferred=inferred, err=err, new_cols=new_cols)


def compose_dsl_step_advisories(mode: str, op: str, resolved: dict, meta: dict, task: str,
                                 before: dict, after: dict, *, exclude_sheets: set | None = None,
                                 deps: DslStepDeps) -> list:
    """⑤適用後の助言。mode="flat"（単発）は build_advisories(exclude_sheets 込み) を丸ごと、
       mode="structural"（複合計画の段）は _structural_advisories + unrequested_new_sheet_advisory
       だけ（依頼文言との重なり④は計画全体に対して1回だけ評価するため、段ごとのここには
       含めない ── 呼び出し側(cmd_run_plan)のコメント参照）。3つの中立化(neutralize)は
       両モード共通。"""
    if mode == "flat":
        advisories = deps.build_advisories(task, before, after, exclude_sheets=exclude_sheets)
    else:
        advisories = list(deps.structural_advisories(before, after))
        advisories.extend(deps.unrequested_new_sheet_advisory(task, before, after))
    advisories = deps.neutralize_new_column_ghost_warning(advisories, op, resolved, meta)
    advisories = deps.neutralize_declared_new_sheet_warning(advisories, op, before, after)
    advisories = deps.neutralize_declared_sheet_replace_warning(advisories, op, before, after)
    return advisories


@dataclass
class DslConfirmResult:
    """print_dsl_confirmation の戻り値。gate_exit が None でなければ、呼び出し側は
       それをそのまま return すべき exit code（破壊の関所で拒否/非対話終了）。"""
    line: str                      # "解釈: ..." の全文
    label: str                     # line から "解釈: " を除いたもの
    warn_overwrite: str | None
    mismatch_warning: str | None
    gate_exit: int | None


def print_dsl_confirmation(op: str, resolved: dict, inferred: set, task: str, *,
                            meta: dict, warn_book: Path, new_cols: list | None,
                            a, deps: DslStepDeps, step_prefix: str = "") -> DslConfirmResult:
    """②検証済み(resolved/inferred)を受け取り、③確認行〜破壊の関所までを印字・実行する。
       単発(cmd_run_dsl)は step_prefix=""・new_cols=None（新規列の概念が無い＝
       常に mismatch_warning は None になる）で呼ぶ。複合計画は段番号つき step_prefix と、
       直前までの段が作った新規列の一覧を渡す。
       ★ warn_book/meta: 単発は (book_meta, book)（まだ何も反映されていない原本）、
       複合計画は (current_meta, out_book)（直前までの段を反映済みの作業コピー）を渡す
       ── 上書き検知の対象が「今の実体」であることは共通で、その実体をどちらから見るかが
       経路によって違う（既存の非対称・今回はそのまま踏襲する）。"""
    line = deps.format_confirmation_line(op, resolved, inferred)
    label = line[len("解釈: "):]
    print(f"{step_prefix}{line}")
    if op == "PIVOT":   # ★ W9 項目4
        print(f"{step_prefix}（{deps.pivot_caveat}）")
    mismatch_warning = deps.maybe_warn_header_col_mismatch(op, resolved, new_cols or [], task)
    if mismatch_warning:
        print(f"{step_prefix}{mismatch_warning}")
    warn_overwrite = deps.maybe_warn_target_overwrite(op, resolved, meta, warn_book)
    if warn_overwrite:
        summary = deps.interpretation_summary_line(resolved, inferred)   # ★ W10a 項目3
        if summary:
            print(f"{step_prefix}{summary}")
        print(f"{step_prefix}{warn_overwrite}")
    for w in resolved.get("_warnings", []):   # ★ A': LLM由来の値と機械抽出の食い違い
        print(f"{step_prefix}⚠ {w}")
    gate_exit = deps.confirm_overwrite_or_gate(a, warn_overwrite, step_prefix=step_prefix)
    return DslConfirmResult(line=line, label=label, warn_overwrite=warn_overwrite,
                             mismatch_warning=mismatch_warning, gate_exit=gate_exit)


@dataclass
class DslApplyResult:
    """apply_dsl_step の戻り値。runtime_error が None でなければ適用そのものが失敗
       （postcondition_status 以下は評価されず None のまま）。"""
    runtime_error: str | None
    after: dict | None
    changes: list | None
    changed: bool
    postcondition_status: str | None    # "pass"/"warn"/"fail"/"error"/None
    postcondition_reason: str | None


def apply_dsl_step(op: str, resolved: dict, code: str, *, apply_target: Path, before: dict,
                    before_charts: int, workdir: Path, helper_files, apply_timeout,
                    header_row: int, use_formula: bool, source_book: Path | None,
                    deps: DslStepDeps, apply_progress_label: str,
                    print_changes: bool) -> DslApplyResult:
    """④codegen 済みの code を⑤適用し⑥事後条件を見る。
       ★ print_changes: 単発(cmd_run_dsl)は True（「変更点:」+差分行を常に印字）・
       複合計画は False（段ごとの差分行は印字せず、after を助言計算にだけ使う ── 既存の
       非対称・今回はそのまま踏襲する）。
       ★★ _truncation_notice はここに含めない（あえて）: 単発は常に呼び・複合計画の
       DSL 段は一度も呼ばない（C6 が発見した未修正の穴）という非対称を、呼び出し側の
       ailine.py 側に「呼ぶ/呼ばない」という**直接の呼び出しの有無**として残すため
       （もしここに bool 引数として畳み込むと、単発と複合計画の両方が同じこの関数を
       呼ぶようになり、_truncation_notice を実際に呼ぶかどうかが関数内部の分岐に隠れて
       しまう ── stage_organs.py の AST 反射番人は「関数の中身」までは追わない設計
       （同モジュールの docstring 参照）なので、それでは単発/複合計画のどちらが
       呼んでいるかを機械で見分けられなくなる。呼び出し側に残すことで、「単発は直接呼ぶ・
       複合計画は呼ばない」という違いが今までどおり AST から見える形のまま保たれる）。
       呼び出し側(cmd_run_dsl)がこの関数の戻り値(after)を使って自分で呼ぶ。"""
    t0 = deps.progress_start(apply_progress_label)
    okrun, err_apply, _raw = deps.basrun_apply(apply_target, code, workdir, helper_files,
                                                timeout=apply_timeout)
    deps.progress_end(t0)
    if not okrun:
        return DslApplyResult(runtime_error=err_apply, after=None, changes=None, changed=False,
                               postcondition_status=None, postcondition_reason=None)

    after = deps.snapshot(apply_target)
    changed, lines = deps.diff_snapshots(before, after)
    if print_changes:
        print("\n変更点:" if changed else "\n（文書に変化は検出されなかった）")
        for ln in lines:
            print(ln)

    status, reason = deps.run_postcondition(
        op, apply_target, resolved, before_charts=before_charts,
        header_row=header_row, use_formula=use_formula, source_book=source_book)
    return DslApplyResult(runtime_error=None, after=after, changes=lines, changed=changed,
                           postcondition_status=status, postcondition_reason=reason)
