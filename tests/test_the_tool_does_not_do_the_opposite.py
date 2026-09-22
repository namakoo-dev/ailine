"""依頼が「その操作を**打ち消せ**」と言っているのに、その操作そのものをしない（2026-09-22）。

★★ 事故（盲検 6 体目 ①・**致命**・こちらで 2/2 再現）:

    > python -m ailine run 経費精算8月.xlsx "B1からF1の結合を解除して"
    解釈: 操作:セル結合 範囲:B1:F1
    （文書に変化は検出されなかった）
    検算しました（セル結合）: B1:F1 の結合を確認
    ✓ 経費精算8月.xlsx は機械検証済みの内容です

  **依頼と真逆の操作に ✓ が付いた。** 買い手（税理士事務所の担当者）の言葉:
    「私が 1 時間で『できていないのに ✓ と言った』場面を引き当てました」── 値付け **0 円**。

★★ なぜ事後条件で捕まらないか（三項）:
    依頼 = 解除 / 宣言 = セル結合 / 実体 = 結合されている
  宣言と実体は**一致している**ので検算は通る。破れているのは**依頼↔宣言**で、
  そこは「人が『解釈:』行を読む」という設計になっている ── 原理的に事後条件の外。

★ 一般に依頼↔宣言は機械で確かめられない。★ だが**この形だけ**は安く捕まる ──
  依頼文が「その op の概念語」と「打ち消しの語」を**両方**含む時。
  概念語は宣言（OP_LABELS・31 op）から引く。手書きの名簿を作らない。

★ 08-29 に**同じ族**を踏んでいる（「合計を金額表示にして」が合計追加になり、既にある
  合計をもう一度書いて ✓）。その時の処置は**その op 専用**だったので族は開いたままだった。

★ 配線は `resolve_dsl_step_args`（単発・帳票・様式写像の 3 経路が共通で通る）1 箇所。
  呼び出し側に配るとまた片配線になる。
"""
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

import ailine  # noqa: E402


@pytest.mark.parametrize("task, op", [
    ("B1からF1の結合を解除して", "MERGE"),
    ("結合を外して", "MERGE"),
    ("セル結合を取り消して", "MERGE"),
    ("太字を解除して", "BOLD"),
    ("並べ替えを取り消して", "SORT"),
    ("けい線を消して", "DRAW_BORDERS"),
])
def test_asking_to_undo_an_op_is_refused(task, op):
    """★★ 本体 ── 打ち消しを頼まれて、その操作そのものをしない。"""
    why = ailine.task_asks_to_undo_this_op(task, op)
    assert why, f"打ち消しの依頼が素通りする: {op} 「{task}」"
    assert ailine.OP_LABELS[op] in why, why


@pytest.mark.parametrize("task, op", [
    ("B1からF1を結合して", "MERGE"),
    ("見出しを太字にして", "BOLD"),
    ("金額で降順に並べ替えて", "SORT"),
    ("けい線を引いて", "DRAW_BORDERS"),
    ("重複している行を削除して", "DEDUP_DELETE"),
    ("3行目を削除して", "DELETE_ROWS"),
    ("備考の列を全部「確認済」に書き換えて", "SET_COLUMN_VALUE"),
])
def test_ordinary_requests_are_not_touched(task, op):
    """★ 陰性対照 ── 普通の依頼を止めない（断りは安全側だが、止めすぎたら道具でない）。

    ★ 削除系を入れてあるのが要点: 「削除して」は打ち消しの語ではない。
      DELETE_ROWS は**削除を宣言している** op なので、依頼と宣言は一致している。
    """
    assert ailine.task_asks_to_undo_this_op(task, op) is None, (op, task)


def test_the_concept_word_comes_from_the_declaration_not_a_list():
    """★ 概念語を試験に書き写さない ── 宣言（OP_LABELS）から引いていること。

    ★ 字面の名簿にすると、ラベルを言い換えた日に**空振りで緑**になる（今日 3 本見つけた形）。
    """
    for op, label in ailine.OP_LABELS.items():
        if len(label) < 2:
            continue
        why = ailine.task_asks_to_undo_this_op(f"{label}を解除して", op)
        assert why, f"{op}（{label}）の打ち消しが素通りする"


def test_an_unknown_op_or_empty_task_says_nothing():
    """★ 材料が無ければ黙る（初回・別 PC で挙動が変わらない）。"""
    assert ailine.task_asks_to_undo_this_op("", "MERGE") is None
    assert ailine.task_asks_to_undo_this_op("結合を解除して", None) is None
    assert ailine.task_asks_to_undo_this_op("結合を解除して", "NOPE") is None


def test_the_guard_is_wired_where_all_three_paths_pass():
    """★★ 器官を置いて配線しない、をさせない ── 3 経路の共通点で呼ばれていること。

    ★ この repo が何度も踏んだ形。呼び出し側に配ると M 箇所だけ直る。
    """
    src = (Path(ailine.__file__).resolve().parent.parent
           / "ailine_core" / "dsl_step.py").read_text(encoding="utf-8")
    assert "task_asks_to_undo_this_op" in src, (
        "resolve_dsl_step_args に配線されていません（呼び出し側に配らないこと）")
    assert src.count("deps.task_asks_to_undo_this_op(") == 1, (
        "打ち消しの判定が 2 箇所以上にあります ── 1 箇所に畳んでください")
