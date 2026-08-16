"""C6: 段種 × 器官の宣言表（ailine_core/stage_organs.py）の網羅を検査する番人。

★ 背景: 独立監査がこの repo の真因を「器官が『呼び出し規約』でなく『記憶』で適用されて
いる」と名指しした（同じ知見が、見つかった場所にだけ適用されて他所へ伝播しない）。
唯一「増やしても壊れない」と実証された形が OP_WRITE_TARGET の宣言駆動＋番人
（test_op_write_target_declares_all_ops）だったので、それを段種 × 器官へ一般化した。

★ この番人が検査するのは表の**内的整合**（全マス目にキーがあるか）だけ。表の値が
実装と実際に一致しているか（「在り」と書いたマス目が本当に走っているか）は、番人ではなく
人手の照合表（PR/報告に載せる）で担保する ―― OP_WRITE_TARGET も同じ役割分担
（test_op_write_target_declares_all_ops は「宣言漏れ」だけを機械的に防ぎ、宣言の中身の
正しさは実行系テスト側が担う）。

★ 純ロジックのみ・LibreOffice/basrun 不要（CI で走る）。ailine_core.stage_organs は
ailine.py に依存しない（import 依存ゼロ）ので、ここでは ailine を import しない。
"""
from ailine_core import stage_organs
from ailine_core.stage_organs import ORGANS, STAGE_ORGANS, STAGES, missing_cells


def test_import_has_no_ailine_dependency():
    """★ DoD7: 循環 import になっていないことの実証。ailine_core.stage_organs は
       import 時に ailine.py（や他の ailine_core モジュール）へ一切触れない
       （モジュール属性が str/tuple/dict/関数だけであることで間接的に確認する）。"""
    import types
    for name, value in vars(stage_organs).items():
        if name.startswith("__") or name == "annotations":   # `from __future__ import annotations` の束縛
            continue
        assert isinstance(value, (str, tuple, dict, types.FunctionType)), (
            f"stage_organs.{name} が予期しない型 {type(value)} ―― "
            "他モジュール（特に ailine.py）への依存が紛れ込んでいないか確認する"
        )


def test_stages_and_organs_are_nonempty_and_unique():
    assert len(STAGES) >= 1 and len(set(STAGES)) == len(STAGES)
    assert len(ORGANS) >= 1 and len(set(ORGANS)) == len(ORGANS)


def test_stage_organs_declares_all_cells():
    """★ 番人の本体: 全段種 × 全器官のマス目に明示的な値がある（＝missing_cells が空）。
       段種を1つ足した人／器官を1つ足した人が、マス目を埋め忘れたらここが赤くなる
       （test_op_write_target_declares_all_ops の一般化）。"""
    missing = missing_cells()
    assert missing == [], (
        f"STAGE_ORGANS に埋め忘れたマス目がある（段種, 器官）: {missing}"
    )


def test_stage_organs_has_no_stray_rows_or_columns():
    """STAGE_ORGANS のキー集合が STAGES/ORGANS と完全一致する（宣言済みだが登録簿から
       落ちた段種・器官が残り続ける事故を防ぐ ―― missing_cells は「無い」側しか見ないため
       別テストで「余分」側を見る）。"""
    assert set(STAGE_ORGANS.keys()) == set(STAGES), (
        f"STAGE_ORGANS のキーが STAGES と食い違っている: "
        f"余分={set(STAGE_ORGANS.keys()) - set(STAGES)} "
        f"不足={set(STAGES) - set(STAGE_ORGANS.keys())}"
    )
    for stage, row in STAGE_ORGANS.items():
        assert set(row.keys()) == set(ORGANS), (
            f"STAGE_ORGANS[{stage!r}] の器官キーが ORGANS と食い違っている: "
            f"余分={set(row.keys()) - set(ORGANS)} 不足={set(ORGANS) - set(row.keys())}"
        )


def test_cell_values_are_true_or_none_only():
    """★ 値は True（在り）/ None（無いと確認した宣言）の2値のみ。False 等の別の
       『無い』表現を紛れ込ませない（意味の統一）。"""
    offenders = [(stage, organ, value)
                 for stage, row in STAGE_ORGANS.items()
                 for organ, value in row.items()
                 if value is not True and value is not None]
    assert offenders == [], f"True/None 以外の値が入っているマス目: {offenders}"


# --- ★ DoD3: 番人の発火実証（マス目を1つ削る→赤→戻す） ------------------------------

def test_missing_cells_detects_a_dropped_organ_column():
    """既存の表から器官1マスを削ると missing_cells が検出することを実証する
       （実ファイルは書き換えない・in-memory のコピーで自己検証する）。"""
    poisoned = {stage: dict(row) for stage, row in STAGE_ORGANS.items()}
    poisoned["dsl_single"].pop("rate_scan")   # ← 1マス削る（赤にする）
    missing = missing_cells(table=poisoned)
    assert ("dsl_single", "rate_scan") in missing, "番人が発火しなかった（削ったマスを検出できていない）"
    # 戻す: poisoned は STAGE_ORGANS のコピーなので、原本は無傷のまま
    # （このアサーションが「戻せば緑に戻る」ことの実証を兼ねる）。
    restored = missing_cells(table=STAGE_ORGANS)
    assert restored == [], "戻した後の本物の STAGE_ORGANS が赤いまま（元の表が既に壊れている）"


def test_missing_cells_detects_a_dropped_stage_row():
    """段種1行を丸ごと削ると、その行の全器官が missing として返ることを実証する。"""
    poisoned = {stage: dict(row) for stage, row in STAGE_ORGANS.items()}
    del poisoned["freeform_plan_step"]   # ← 1行削る（赤にする）
    missing = missing_cells(table=poisoned)
    assert set(missing) == {("freeform_plan_step", organ) for organ in ORGANS}, (
        "行削除の検出が全器官分そろっていない（一部だけ検出/過検出している）"
    )
    restored = missing_cells(table=STAGE_ORGANS)
    assert restored == [], "戻した後の本物の STAGE_ORGANS が赤いまま（元の表が既に壊れている）"


def test_missing_cells_detects_a_newly_added_stage_not_yet_declared():
    """★ 実運用のシナリオそのもの: STAGES に新しい段種を1つ足したが STAGE_ORGANS 側の
       宣言を書き忘れた場合を再現する（段種を足した人がマス目を埋め忘れる、が
       このモジュールの想定する再発パターン）。"""
    new_stages = STAGES + ("dsl_batch_step",)   # まだ STAGE_ORGANS に無い新段種
    missing = missing_cells(stages=new_stages)
    assert set(missing) == {("dsl_batch_step", organ) for organ in ORGANS}


def test_missing_cells_detects_a_newly_added_organ_not_yet_declared():
    """★ 同上・器官側: ORGANS に新しい器官を1つ足したが、全段種の宣言を書き忘れた場合。"""
    new_organs = ORGANS + ("semantic_sanitize",)   # まだどの段種にも無い新器官
    missing = missing_cells(organs=new_organs)
    assert set(missing) == {(stage, "semantic_sanitize") for stage in STAGES}
