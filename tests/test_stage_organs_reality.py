"""C6b: 段種 × 器官の宣言表（stage_organs.STAGE_ORGANS）を ailine.py の実装と突き合わせる番人。

★ 背景: C6 で入った missing_cells() は表の**内的整合**（全マス目にキーがあるか）しか
見ない。実測したところ、この表を*読むコード*が ailine.py に0箇所ある ―― つまり宣言が
現実と照合されていない。誰かが scan_rate_literals の呼び出しを複合段から消しても、表は
True のまま・missing_cells は緑のまま通る。これはこの repo が繰り返し踏んできた欠陥
（宣言はあるが実体と突き合わせる機械が無い）そのもの。

★ やること: ailine.py を AST で読み、「True と宣言したマス目の器官が、その段の実際の
経路（stage_organs.STAGE_ENTRY_FUNCTIONS が宣言する関数の本体）で呼ばれているか」を
検査する。逆向き（None と宣言したのに実際には呼ばれている）も見る
（stage_organs.reality_mismatches が両方向を判定する純ロジック本体・ここでは
ailine.py を読んで AST から「呼ばれている関数名の集合」を作るところだけを担当する）。

★ ailine.py は読むだけ（ast.parse のみ・import しない）。BASRUN 等の実行時依存を
一切引かない（DoD6: BASRUN を存在しないパスに向けても緑）。
★ C6b の時点では ailine.py には1行も触れていなかった（このテストの新規追加のみ。
stage_organs.py への追加は STAGE_ENTRY_FUNCTIONS/ORGAN_FUNCTION_CANDIDATES/
reality_mismacthes の3点＝ブリーフが明示的に許可した「段↔関数の対応表」）。
★★ C7（三経路の統合）でこの前提は変わった: ailine.py 自体を書き換える回だったため、
STAGE_ENTRY_FUNCTIONS/ORGAN_FUNCTION_CANDIDATES の宣言をその変更に合わせて更新している
（dsl_plan_step の代表関数を cmd_run_plan → _run_dsl_plan_step へ、clarify_plan_step を
unverifiable から卒業、destructive_gate の候補に print_dsl_confirmation を追加）。
このファイル自身（AST を読むだけの番人のロジック）は変えていない。
"""
import ast
from pathlib import Path

import pytest

from ailine_core import stage_organs
from ailine_core.stage_organs import (
    ORGANS,
    ORGAN_FUNCTION_CANDIDATES,
    STAGE_ENTRY_FUNCTIONS,
    STAGE_ORGANS,
    STAGES,
    reality_mismatches,
)

REPO_ROOT = Path(__file__).resolve().parent.parent
AILINE_PY = REPO_ROOT / "ailine.py"


def _ailine_ast() -> ast.Module:
    """ailine.py を AST として読む（import はしない＝BASRUN 等の実行時依存を引かない）。"""
    src = AILINE_PY.read_text(encoding="utf-8")
    return ast.parse(src, filename=str(AILINE_PY))


def _toplevel_functions(tree: ast.Module) -> dict:
    """モジュール直下（ネスト無し）の関数定義だけを name → node で拾う。
       ★ STAGE_ENTRY_FUNCTIONS が指す5関数はいずれもトップレベル定義（ailine.py に
       class は _FreeformGateAbort の1つだけで、器官/段の対象関数はどれもその外）。"""
    return {node.name: node
            for node in tree.body
            if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef))}


def _called_names(func_node: ast.AST) -> set:
    """func_node（関数定義）のサブツリー内で直接呼ばれている呼び出し先の名前の集合。
       ast.Name（例: verify_dsl_args(...)）と ast.Attribute（例: obj.method(...)）の
       両方を拾う。★ ここで拾うのは func_node 自身の本体だけ ―― 呼び出し先の関数の中に
       何があるかは追わない（間接呼び出しは構造的に見えない＝ヘルパ経由の呼び出しが
       あれば見逃す。それは stage_organs.STAGE_ENTRY_FUNCTIONS の設計で承知の上の限界
       ―― 誤検知よりは見逃しを選ぶ）。"""
    names = set()
    for node in ast.walk(func_node):
        if isinstance(node, ast.Call):
            f = node.func
            if isinstance(f, ast.Name):
                names.add(f.id)
            elif isinstance(f, ast.Attribute):
                names.add(f.attr)
    return names


def _found_names_by_stage() -> dict:
    """STAGE_ENTRY_FUNCTIONS が宣言する各段種の代表関数を ailine.py の AST から実際に
       引き当て、その本体で呼ばれている名前の集合を段種ごとに返す。
       ★★ 宣言した関数名が ailine.py に見当たらない場合は誤魔化さずに即失敗する
       （関数が改名/削除されて宣言が腐った状態を「呼び出しゼロ」と誤読しない ―― 手で
       書いた対応表がコードの変化に取り残される、というこの repo の実測パターンへの
       対策そのもの）。"""
    tree = _ailine_ast()
    funcs = _toplevel_functions(tree)
    result = {}
    for stage, entry_funcs in STAGE_ENTRY_FUNCTIONS.items():
        found: set = set()
        for fname in entry_funcs:
            node = funcs.get(fname)
            assert node is not None, (
                f"stage_organs.STAGE_ENTRY_FUNCTIONS[{stage!r}] が指す関数 "
                f"{fname!r} が ailine.py に見当たらない（改名/削除されて宣言が"
                "腐っている可能性。対応表を先に直す）"
            )
            found |= _called_names(node)
        result[stage] = found
    return result


# --- 対応表自体の網羅（missing_cells と同じ役割の入口ガード） -------------------------

def test_stage_entry_functions_declares_every_stage():
    """STAGE_ENTRY_FUNCTIONS のキー集合が STAGES と完全一致する（段種を1つ足したのに
       ここへの宣言（unverifiable の明示的な空タプルも含む）を書き忘れる事故を防ぐ）。"""
    assert set(STAGE_ENTRY_FUNCTIONS.keys()) == set(STAGES)


def test_organ_function_candidates_declares_every_organ():
    """ORGAN_FUNCTION_CANDIDATES のキー集合が ORGANS と完全一致する（器官を1つ足したのに
       ここへの候補関数の宣言を書き忘れる事故を防ぐ）。"""
    assert set(ORGAN_FUNCTION_CANDIDATES.keys()) == set(ORGANS)


def test_unverifiable_stages_are_exactly_the_declared_one():
    """★ DoD4/C7 DoD7: 追えなかった（unverifiable に逃がした）段種の一覧を固定する。
       ★★ C7 で空集合に更新: clarify_plan_step が卒業した。C7（三経路の統合）で
       cmd_run_plan の DSL 段の実体を _run_dsl_plan_step という別関数に切り出し、
       --dry プレビュー本体も _preview_dsl_plan に切り出したことで、cmd_run_plan
       自身の関数本体はもう CLARIFY 分岐(inline 3行)＋委譲呼び出し＋ループ外集計
       （どの ORGAN_FUNCTION_CANDIDATES にも該当しない）だけになった ── dsl_plan_step の
       器官呼び出しを誤って拾う心配が無くなったので、cmd_run_plan を clarify_plan_step の
       代表関数として安全に使えるようになった（stage_organs.STAGE_ENTRY_FUNCTIONS の
       コメント参照）。
       ★ ここが増えたら（新しい段種が追えなくなったら）このテストが教えてくれる
       （unverifiable が静かに広がっていくのを防ぐ）。"""
    unverifiable = {stage for stage, funcs in STAGE_ENTRY_FUNCTIONS.items() if not funcs}
    assert unverifiable == set()


# --- ★ 番人の本体: 実際の ailine.py に対して両方向を検査する ---------------------------

def test_declared_cells_match_ailine_py_reality():
    """★ DoD の核心。stage_organs.STAGE_ORGANS が宣言する True/None を、ailine.py の
       実際の呼び出しと突き合わせる。両方向:
       - True 宣言なのに呼び出しが見つからない（宣言だけ残って実装が消えた）
       - None 宣言なのに呼び出しが見つかる（無いと言ったのに実は在る）
       のどちらも mismatches に出る。今の ailine.py（1行も変更していない）に対しては
       空でなければならない。"""
    found = _found_names_by_stage()
    mismatches = reality_mismatches(found)
    assert mismatches == [], (
        f"宣言表 (STAGE_ORGANS) と ailine.py の実装がずれているマス目: {mismatches}"
    )


def test_dsl_plan_step_truncation_notice_gap_stays_green():
    """★ DoD5: C6 の報告で見つかった穴（dsl_plan_step × truncation_notice）が、この番人でも
       赤くならないことを固定する。この穴は「本来あるべきなのに無い」欠陥だが、
       stage_organs.STAGE_ORGANS の宣言自体が None（無いと確認した宣言）になっている
       ので、宣言と現実は一致している ―― この番人は宣言と現実のズレしか検査しない
       （あるべき理想との比較はしない、C6 の設計方針をそのまま継承）。緑のままが正しい。"""
    assert STAGE_ORGANS["dsl_plan_step"]["truncation_notice"] is None
    found = _found_names_by_stage()
    candidates = set(ORGAN_FUNCTION_CANDIDATES["truncation_notice"])
    assert not (found["dsl_plan_step"] & candidates), (
        "dsl_plan_step の関数本体 (cmd_run_plan) で _truncation_notice が呼ばれている"
        "ことが検出された ―― 穴が塞がれたのであれば良い変化。ただしその場合は"
        "stage_organs.STAGE_ORGANS['dsl_plan_step']['truncation_notice'] を True に更新し、"
        "このテストの前提コメントも直すこと（None のままでは reality チェックの方が赤くなる）"
    )


# --- ★ DoD3: 番人の発火実証（両方向・実ファイルは書き換えず in-memory で行う） -----------

def test_guard_fires_when_a_true_declared_call_is_erased_in_memory():
    """(a) 「True と宣言した器官の呼び出しを ailine.py からメモリ上で消す」の実演。
       実際に見つかった found_names_by_stage のコピーから、本当に True で本当に見つかって
       いる1マス（dsl_single × grounding = resolve_dsl_step_args。★ C7: cmd_run_dsl は
       ailine_core.dsl_step の共有 grounding 関数を経由するようになった）の候補名を
       取り除き、『誰かがその呼び出しを消した』状態を再現する。"""
    found = _found_names_by_stage()
    assert STAGE_ORGANS["dsl_single"]["grounding"] is True
    assert "resolve_dsl_step_args" in found["dsl_single"], "前提が崩れている（現状 True なのに見つからない）"

    poisoned = {stage: set(names) for stage, names in found.items()}
    poisoned["dsl_single"].discard("resolve_dsl_step_args")   # ← 呼び出しをメモリ上で消す

    mismatches = reality_mismatches(poisoned)
    assert ("dsl_single", "grounding", "declared_true_not_found") in mismatches, (
        "番人が発火しなかった（消した呼び出しを検出できていない）"
    )
    # 戻す: poisoned は found のコピーなので、本物の found/STAGE_ORGANS は無傷のまま。
    restored = reality_mismatches(found)
    assert restored == [], "戻した後も赤いまま（実ファイルの状態が既にずれている）"


def test_guard_fires_when_a_none_cell_is_promoted_to_true_without_a_real_call():
    """(a') 上の代替実演: 「宣言を True に増やす」。dsl_single × rate_scan は実際には
       構造的に対象が無く None（かつ現実にも呼ばれていない）。ここを誤って True と
       書いてしまった場合を再現する（宣言側だけを動かす・found はそのまま実ファイル）。"""
    found = _found_names_by_stage()
    assert STAGE_ORGANS["dsl_single"]["rate_scan"] is None
    assert "scan_rate_literals" not in found["dsl_single"], "前提が崩れている（現状 None なのに見つかっている）"

    poisoned_table = {stage: dict(row) for stage, row in STAGE_ORGANS.items()}
    poisoned_table["dsl_single"]["rate_scan"] = True   # ← 宣言だけ True に増やす

    mismatches = reality_mismatches(found, table=poisoned_table)
    assert ("dsl_single", "rate_scan", "declared_true_not_found") in mismatches, (
        "番人が発火しなかった（実装の無い True 宣言を検出できていない）"
    )
    restored = reality_mismatches(found)
    assert restored == [], "戻した後も赤いまま（実ファイルの状態が既にずれている）"


def test_guard_fires_when_a_true_cell_is_downgraded_to_none_while_the_call_still_exists():
    """(b) ★ 逆向き実演: 「None と宣言したマス目の器官が、実際には呼ばれている」ケース。
       dsl_single × grounding は実際に resolve_dsl_step_args を呼んでいる（True・確認済み）。
       これを None に書き換えると『無いと言ったのに実は在る』状態になる ―― reality
       チェックがこれも赤にすることを実演する。"""
    found = _found_names_by_stage()
    assert STAGE_ORGANS["dsl_single"]["grounding"] is True
    assert "resolve_dsl_step_args" in found["dsl_single"], "前提が崩れている（現状 True なのに見つからない）"

    poisoned_table = {stage: dict(row) for stage, row in STAGE_ORGANS.items()}
    poisoned_table["dsl_single"]["grounding"] = None   # ← 宣言を「無い」に書き換える

    mismatches = reality_mismatches(found, table=poisoned_table)
    assert ("dsl_single", "grounding", "declared_none_but_found") in mismatches, (
        "番人が発火しなかった（None 宣言なのに実在する呼び出しを検出できていない＝逆向きチェックの穴）"
    )
    restored = reality_mismatches(found)
    assert restored == [], "戻した後も赤いまま（実ファイルの状態が既にずれている）"


# --- CI 相当（BASRUN 非依存の実証） -------------------------------------------------

def test_reality_check_does_not_touch_basrun_env(monkeypatch):
    """★ DoD6: BASRUN を存在しないパスに向けてもこの番人は緑のまま
       （ailine.py を AST で読むだけで import も実行もしないため、basrun.py の所在に
       一切依存しない）。"""
    monkeypatch.setenv("BASRUN", str(REPO_ROOT / "does" / "not" / "exist" / "basrun.py"))
    found = _found_names_by_stage()
    assert reality_mismatches(found) == []


if __name__ == "__main__":
    raise SystemExit(pytest.main([__file__, "-q"]))
