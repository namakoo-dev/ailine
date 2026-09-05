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

from ailine_core.projection import render_projection_notice

import shutil
from dataclasses import dataclass
from pathlib import Path
from typing import Callable

from ailine_core.subject import CONTRADICTED, contradiction_lines, unspoken_subjects   # ★ 単位E

# ★ 単位B: 「対象の列がこの計画の直前の段で新規作成された列である」という事実を言う1句。
#   手書きの if（ailine.py の _maybe_warn_header_col_mismatch）の文と、③の ⚠ に畳み込む
#   注記の**両方がここを読む** ―― 同じ事実を2箇所に書くと、文面が黙って食い違うため。
NEW_COLUMN_ORIGIN = "この計画の直前の段で新規作成された列"

# ★ 摩擦⑥: LO の一時的な不調の凍結マーカー。normalize_book(M2c) が監査2回で2回再現した
#   既知の摩擦（RuntimeException: Could not create system bitmap!・ailine.py 参照）に、
#   operator8③の真因だった DisposedException を加えた2種。この1箇所だけに置き、適用側の
#   2経路（apply_dsl_step・ailine.py の run_freeform_plan_step）と、正規化側の由来コメントが
#   ここを指す ── 表を2箇所に書くと片方だけ更新されて食い違うため。それ以外のエラーは
#   普通の実行時エラーとして即時失敗させる（盲目リトライをしない・検体③）。
TRANSIENT_LO_MARKERS = ("DisposedException", "Could not create system bitmap")

# ★ 摩擦⑥: 再試行した事実の開示（1行）。「再試行」の3文字を含む（検体が機械検査）。
TRANSIENT_LO_RETRY_NOTICE = "LibreOffice の一時的な不調のため、再起動して再試行しました。"


def is_transient_lo_error(err: str | None) -> bool:
    """err が TRANSIENT_LO_MARKERS のいずれかに部分一致すれば True。"""
    if not err:
        return False
    return any(marker in err for marker in TRANSIENT_LO_MARKERS)


@dataclass
class DslStepDeps:
    """ailine.py 側の関数参照の束。呼び出し側が呼び出しのたびに組み立てる（上記docstring参照）。"""
    format_confirmation_line: Callable
    maybe_warn_header_col_mismatch: Callable
    maybe_warn_target_overwrite: Callable
    interpretation_summary_line: Callable
    confirm_overwrite_or_gate: Callable
    basrun_apply: Callable
    # ★ 摩擦⑥: LO の一時不調から復元する再試行が使う（ailine.py の _stop_office）。
    stop_office: Callable
    snapshot: Callable
    # ★ 2026-08-25（復元の中10）: 成果物が Excel として開けるかを見る（開けるなら None）。
    #   ailine_core は ailine を import しない規律なので、本体から注入する。
    why_output_is_unusable: Callable
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
    # ★ 単位E: 対象スロットの出所を①②③に仕分ける（ailine.py の classify_subject_provenance）。
    #   None なら仕分けない（既定・後方互換 ―― 直接この関数を呼ぶ既存テストを壊さない）。
    classify_subject_provenance: Callable | None = None
    # ★ 挙動変更#3: シート名の衝突で既定(1枚目)へ後退した時に「解釈:」行の直後で3択を
    #   聞く関門（ailine.py の _sheet_conflict_gate）。None なら聞かない（既定）。
    #   ★ ここに置く理由: 3択は**解釈行のすぐ後**でなければならない（操作が確定してから
    #   具体的な日本語で選ばせるという設計判断・ailine_core/target_sheet.py 参照）。
    #   print_dsl_confirmation の**後**に呼ぶと、破壊の関所（上書きしますか？）の方が
    #   先に聞いてしまい、どのシートを触るか未確定のまま上書き可否を尋ねることになる。
    sheet_conflict_gate: Callable | None = None


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
    """②検証（grounding）。単発(cmd_run_dsl)は original_headers=None のまま呼ぶ（new_cols は
       常に空になり、_apply_new_column_fallback は無条件でスキップされる ── 新規列の概念自体
       が単発には無いため、この呼び方で単発の従来挙動と完全一致する）。★ 挙動変更#2:
       first_sheet は今は単発でも渡す（resolve_target_sheet が一箇所で決めた対象シート）。
       original_headers=None のガードにより new_cols への影響は無い（下記参照）。
       複合計画は直前までの段の original_headers/current_meta/first_sheet を渡し、
       『直前段が作った新規列』への依存つき連鎖フォールバックを1回だけ試みる。"""
    new_cols: list = []
    if first_sheet and original_headers is not None:
        new_cols = [c for c in meta["headers"].get(first_sheet, [])
                    if c not in original_headers.get(first_sheet, [])]
    ok, resolved, inferred, err = deps.verify_dsl_args(op, raw_args, meta, task=task, vocab=vocab,
                                                         target_sheet=first_sheet)
    if not ok and new_cols and first_sheet:
        patched = deps.apply_new_column_fallback(
            op, raw_args, meta["headers"].get(first_sheet, []), new_cols)
        if patched != raw_args:
            ok2, resolved2, inferred2, err2 = deps.verify_dsl_args(
                op, patched, meta, task=task, vocab=vocab, target_sheet=first_sheet)
            if ok2:
                ok, resolved, inferred, err = ok2, resolved2, inferred2, err2
    return DslGroundResult(ok=ok, resolved=resolved, inferred=inferred, err=err, new_cols=new_cols)


def compose_dsl_step_advisories(mode: str, op: str, resolved: dict, meta: dict, task: str,
                                 before: dict, after: dict, *, exclude_sheets: set | None = None,
                                 sheet_conflict=None, precondition_broken: str | None = None,
                                 after_path=None, deps: DslStepDeps) -> list:
    """⑤適用後の助言。mode="flat"（単発）は build_advisories(exclude_sheets 込み) を丸ごと、
       mode="structural"（複合計画の段）は _structural_advisories + unrequested_new_sheet_advisory
       だけ（依頼文言との重なり④は計画全体に対して1回だけ評価するため、段ごとのここには
       含めない ── 呼び出し側(cmd_run_plan)のコメント参照）。
       ★ C9: op/resolved/meta（今回の段の宣言済み効果）を build_advisories/
       structural_advisories/unrequested_new_sheet_advisory へそのまま渡す ── 宣言済み効果と
       一致する行の中立化は、以前はここで3つの neutralize_* を後処理として適用していたが、
       各生成関数自身が発生源で判定するようになったため、この後処理は不要になった
       （ailine.py の detect_ghost_data/unrequested_new_sheet_advisory/
       existing_sheet_replaced_advisory 参照）。
       ★ 誤爆#3: sheet_conflict は resolve_target_sheet が「この語は列名とも一致したので
       曖昧だから既定へ後退した」と決めた結果（SheetNameConflict）。助言側が同じ判定を
       やり直さないよう、そのまま build_advisories へ運ぶだけにする（mode="structural" の
       段は依頼文言との重なり④を評価しないので受け取っても使い道が無く、渡さない）。
       ★★ 単位G: precondition_broken は「単位F の前提検査で破れた種類」（破れていなければ
       None）。中立化は前提が成立していた時だけ行われるべきなので、判定する側（助言の
       発生源）まで運ぶ。★ ここでも同じ検査をやり直さない ── 検査は呼び出し側で 1 度だけ。
       ★ operator10 ⑤: after_path（適用後の実ファイル・省略可）は build_advisories/
       structural_advisories へそのまま横流しするだけ（数式セルの偽アラーム判定専用）。"""
    if mode == "flat":
        advisories = deps.build_advisories(task, before, after, exclude_sheets=exclude_sheets,
                                            op=op, resolved=resolved, meta=meta,
                                            sheet_conflict=sheet_conflict,
                                            precondition_broken=precondition_broken,
                                            after_path=after_path)
    else:
        advisories = list(deps.structural_advisories(before, after, op=op, resolved=resolved, meta=meta,
                                                     precondition_broken=precondition_broken,
                                                     after_path=after_path))
        advisories.extend(deps.unrequested_new_sheet_advisory(task, before, after, op=op))
    # ★ 2026-08-25（塊①）: 事後条件の checker が「検証できなかった行」を機械の値として
    #   resolved["_unverified"] に残す。ここは単発・複合計画の**両方が通る唯一の合流点**
    #   なので、1 箇所で助言に載せる（呼び出し側 4 箇所に書き写さない ── 今日までに
    #   片配線を 6 回踏んでいる）。⚠ 始まりなので決裁③が数えて ✓ を △ に降ろす。
    from ailine_core.claim import render_unverified_advisories
    advisories.extend(render_unverified_advisories((resolved or {}).get("_unverified")))
    # ★★ 2026-08-26: 削除は**画面の差分に何も出ない**操作なので、何を消したかを言わなければ
    #   人は取り返しがつくかを判断できない（「消えたものは差分に出ない」の家系）。
    #   ★ ⚠ は付けない ── 頼まれたとおりに消えたことは事実で、✓ を降ろす理由にはならない。
    #     ただし**必ず読める場所に出す**（undo で戻せることも同時に言う）。
    deleted = (resolved or {}).get("_deleted") or []
    if deleted:
        advisories.append(f"消した中身（{len(deleted)} 行）── 戻すなら ailine undo:")
        for row in deleted[:10]:
            shown = "／".join("" if v is None else str(v) for v in row)
            advisories.append(f"  ・{shown}")
        if len(deleted) > 10:
            advisories.append(f"  ・ほか {len(deleted) - 10} 行")
    return advisories


@dataclass
class DslConfirmResult:
    """print_dsl_confirmation の戻り値。gate_exit が None でなければ、呼び出し側は
       それをそのまま return すべき exit code（破壊の関所で拒否/非対話終了）。"""
    line: str                      # "解釈: ..." の全文
    label: str                     # line から "解釈: " を除いたもの
    warn_overwrite: str | None
    # ★ 単位B: ③の ⚠ へ畳み込んだ場合は None を返す（呼び出し側が助言へ再掲するのを止める
    #   ―― 畳み込み後の1本は既に段の位置で出ており、助言に同じ事実をもう一度は載せない）。
    mismatch_warning: str | None
    gate_exit: int | None
    # ★ 単位E: 対象スロットの出所。subject_warnings は③（依頼文の語と矛盾する対象・
    #   適用前に印字済み・✓ を出さない理由）、unspoken は②（無言なので機械決定した対象・
    #   ✓ の直後の1文の材料）。仕分けをしない呼び出し（deps 未設定）では両方とも空。
    subject_warnings: tuple = ()
    unspoken: tuple = ()
    # ★ 段1: 対象スロットの判定結果そのもの（SubjectVerdict のリスト）。呼び出し側が
    #   `interpretation`（--json の機械可読な解釈）を組む材料として使う。仕分けをしない
    #   呼び出し（deps 未設定）では空 ―― ここでも二重に classify_subject_provenance を
    #   呼ばない（判定は1回・消費の台帳(Consumed)を余計に進めないため）。
    verdicts: tuple = ()


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
    # ★ 挙動変更#3: 複数シートのブックでは「解釈:」行の先頭に対象シートを載せる
    #   （1枚だけのブックは meta["sheets"] が1件なので従来どおり沈黙＝出力は不変）。
    line = deps.format_confirmation_line(op, resolved, inferred, sheets=meta.get("sheets"),
                                          target_sheet=resolved.get("_target_sheet"))
    label = line[len("解釈: "):]
    print(f"{step_prefix}{line}")
    # ★★ 2026-09-05（投影法）: 表を別の形へ写す op は、**何を保存しないか**を
    #   書く前に言う。地図が「メルカトルです」と明記するのと同じ ── 面積が歪むことを
    #   認めているから書ける。
    #   ★ ここは単発も複合計画も通る**唯一の合流点**（呼び出し側 2 箇所に配らない）。
    #   ★ 挙動は 1 ビットも変えない ── 見せるだけ。
    for _proj_line in render_projection_notice(op):
        print(f"{step_prefix}{_proj_line}")
    # ★ 挙動変更#3: 衝突で既定へ後退していたら、ここ（解釈行の直後・まだ原本に触れる前）で
    #   3択を聞く。単発(step_prefix=="")だけを対象にする ── 複合計画の途中の段で対象シート
    #   を選び直すと、直前までの段を適用済みの作業コピーの上で計画をやり直すことになり、
    #   「原本にはまだ触れていないので安全」という前提が崩れるため（ASSUMED・報告参照）。
    if not step_prefix and deps.sheet_conflict_gate is not None:
        gate = deps.sheet_conflict_gate(a, op)
        if gate is not None:
            return DslConfirmResult(line=line, label=label, warn_overwrite=None,
                                     mismatch_warning=None, gate_exit=gate)
    if op == "PIVOT":   # ★ W9 項目4
        print(f"{step_prefix}（{deps.pivot_caveat}）")
    mismatch_warning = deps.maybe_warn_header_col_mismatch(op, resolved, new_cols or [], task)
    # ★ 単位E: 対象スロットの出所を①②③に仕分ける。③（依頼文の語と矛盾する対象）はここで
    #   ⚠ を出し、下の関所へ「止めて確認する理由」として渡す ―― 適用前に言わなければ
    #   確認にならない。②は印字せず持ち帰り、✓ の直後の1文（範囲を狭める）に使う。
    # ★ 単位B: 手書きの if（mismatch_warning）と一般則（③）が**同じスロット**について同時に
    #   鳴ったら、⚠ を2行並べずに1本へ畳む ―― 手書きの if が持つ固有の事実（この列は直前の段が
    #   作った）を③の行の注記として運び、判断の材料（依頼文が指している語）はそのまま残す。
    #   ★ 手書きの if を消すのではない: 一般則が沈黙する反例が実測で在る（前段が作った新規列の
    #   名前が依頼文にそのまま出る場合・tests/test_subject_provenance.py の
    #   TestGeneralRuleVsHandWrittenIf）。その時は今までどおり単独で鳴り、助言にも再掲される。
    subject_warnings: tuple = ()
    unspoken: tuple = ()
    verdicts: tuple = ()
    folded = False
    if deps.classify_subject_provenance is not None:
        verdicts = tuple(deps.classify_subject_provenance(op, resolved, meta, task, a))
        target = str(resolved.get("target") or "")
        folded = bool(mismatch_warning) and any(
            v.tier == CONTRADICTED and str(v.slot.value) == target for v in verdicts)
        subject_warnings = tuple(contradiction_lines(
            verdicts, notes={target: NEW_COLUMN_ORIGIN} if folded else None))
        unspoken = tuple(unspoken_subjects(verdicts))
    if mismatch_warning and not folded:
        print(f"{step_prefix}{mismatch_warning}")
    for w in subject_warnings:
        print(f"{step_prefix}{w}")
    warn_overwrite = deps.maybe_warn_target_overwrite(op, resolved, meta, warn_book)
    if warn_overwrite:
        summary = deps.interpretation_summary_line(resolved, inferred)   # ★ W10a 項目3
        if summary:
            print(f"{step_prefix}{summary}")
        print(f"{step_prefix}{warn_overwrite}")
    for w in resolved.get("_warnings", []):   # ★ A': LLM由来の値と機械抽出の食い違い
        print(f"{step_prefix}⚠ {w}")
    gate_exit = deps.confirm_overwrite_or_gate(a, warn_overwrite, step_prefix=step_prefix,
                                                subject_mismatch=bool(subject_warnings))
    return DslConfirmResult(line=line, label=label, warn_overwrite=warn_overwrite,
                             mismatch_warning=None if folded else mismatch_warning,
                             gate_exit=gate_exit,
                             subject_warnings=subject_warnings, unspoken=unspoken,
                             verdicts=verdicts)


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
                    print_changes: bool, step_prefix: str = "",
                    before_chart_paths: frozenset | None = None) -> DslApplyResult:
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
       呼び出し側(cmd_run_dsl)がこの関数の戻り値(after)を使って自分で呼ぶ。
       ★ 摩擦⑥: LO の一時不調（TRANSIENT_LO_MARKERS）で①回目の適用が失敗したら、
       stop_office() → source_book（無垢の原本）から apply_target を作り直し → 1回だけ
       再試行する（正規化側の normalize_book/M2c と同型）。半適用の残骸の上に再実行しない
       （②契約）。source_book が無い(None)呼び出しは復元先が無いので再試行しない。
       2回目も失敗すれば従来どおり err_apply をそのまま返す（③・正直な失敗）。
       step_prefix: 複合計画の段番号表示（"  1段目: " 等）。単発は既定の "" のまま。
       ★ 致命④(2026-08-23レビュー): before_chart_paths（snapshot()["chart_paths"]）は
       run_postcondition の CHART 判定へそのまま運ぶだけ（今回増えた1個の同定・
       chart_check.check_chart_series 参照）。None なら従来どおり（後方互換）。"""
    t0 = deps.progress_start(apply_progress_label)
    okrun, err_apply, _raw = deps.basrun_apply(apply_target, code, workdir, helper_files,
                                                timeout=apply_timeout)
    deps.progress_end(t0)
    if not okrun and source_book is not None and is_transient_lo_error(err_apply):
        deps.stop_office()
        shutil.copy2(source_book, apply_target)   # 半適用の残骸を消し、無垢の原本から作り直す
        print(f"{step_prefix}{TRANSIENT_LO_RETRY_NOTICE}")
        t0 = deps.progress_start(apply_progress_label)
        okrun, err_apply, _raw = deps.basrun_apply(apply_target, code, workdir, helper_files,
                                                    timeout=apply_timeout)
        deps.progress_end(t0)
    if not okrun:
        return DslApplyResult(runtime_error=err_apply, after=None, changes=None, changed=False,
                               postcondition_status=None, postcondition_reason=None)

    # ★★ 2026-08-25（復元の中10・盲検）: 適用直後・読み戻しの**前**に「そもそも開けるか」
    #   を見る。旧版は関門が無く、壊れた成果物（zip として読めない等）を原本へ被せてから
    #   「読み戻して確認できませんでした」と言っていた ── **確認は原本を潰した後**だった。
    #   ★ ここは全 op が通る唯一の合流点で、原本はまだ無傷。止められるのはここだけ。
    #   ★ 中身の正しさは見ない（それは事後条件の仕事）── 見るのは「開けるか」だけ。
    broken = deps.why_output_is_unusable(apply_target)
    if broken:
        return DslApplyResult(runtime_error=f"作った結果が壊れています（{broken}）",
                               after=None, changes=None, changed=False,
                               postcondition_status=None, postcondition_reason=None)

    after = deps.snapshot(apply_target)
    changed, lines = deps.diff_snapshots(before, after)
    if print_changes:
        print("\n変更点:" if changed else "\n（文書に変化は検出されなかった）")
        for ln in lines:
            print(ln)

    status, reason = deps.run_postcondition(
        op, apply_target, resolved, before_charts=before_charts,
        header_row=header_row, use_formula=use_formula, source_book=source_book,
        before_chart_paths=before_chart_paths)
    return DslApplyResult(runtime_error=None, after=after, changes=lines, changed=changed,
                           postcondition_status=status, postcondition_reason=reason)
