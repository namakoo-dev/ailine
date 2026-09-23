# -*- coding: utf-8 -*-
"""第二段翻訳の注記（_OP_SCHEMA_NOTES）を、**渡す op 全部**について決める（2026-09-17）。

★ なぜ在るか（仕分け③）: 配線盤に「★ 未調査」で出ていた列。
  「slot 名だけでは意味が伝わらない op にだけ 1 行足す」とは書いてあるが、
  **どの op がそれなのかを誰がどう決めたのか**が書かれていなかった。

★★ 分母は 30 op ではない。注記が読まれるのは `_op_schema_doc(op)` ただ 1 箇所で、
  それを使うのは **op を固定した第二段翻訳**だけ。op を文字どおり書いて渡している
  呼び出しを実装から数えると 5 op ── これを分母にする（手書きしない）。
  ★ ただし op が**変数**の呼び出しも 3 箇所在り、そこはどの op でも来うる。
    だから「5 op が全部」ではなく「5 op は必ず決める」という縛りにする。

★★ 4 本の注記を読むと、全部が同じことを言っている（2026-09-17 に実測）──
  **「これは入れない。機械が決める」**。
    SWAP             行番号・列番号・A1 の座標は入れない
    EXTRACT_COLUMNS  列番号や A1 の座標は入れない
    SET_WHERE        書き込む値と閾値の数字は入れない（機械が依頼文から取る）
    ADD_COLUMN       位置（右/左/末尾）は入れない（機械が実表の見出しから決める）
  つまり注記は飾りの説明ではなく、**分担の宣言**である。この形を番人が縛る。

★ OPS_DOC は増やさない ── 16 行足したら op 一致が 98.1% → 94.2% に落ちた実測が在る。
  注記が第二段にだけ出るのは、op が既に決まっていて語彙を濁さないから。
"""
import ast
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
import ailine  # noqa: E402

from _product_source import product_files  # noqa: E402 ── ★ 番人は本体決め打ちでなく製品コード全体を読む


def _trees():
    """★ 2026-09-23: 本体 1 冊でなく製品全体の AST（呼び出しが ailine_core へ移っても数える）。"""
    return [ast.parse(p.read_bytes().decode("utf-8")) for p in product_files()]


def ops_passed_literally() -> set:
    """`translate_task_fixed_op(model, "OP", ...)` と**文字どおり**書かれている op。

    ★ 実装から数える。ここを手書きにすると、呼び出しを足した日に誰も気づかない。
    """
    out = set()
    for node in (n for tree in _trees() for n in ast.walk(tree)):
        if (isinstance(node, ast.Call) and isinstance(node.func, ast.Name)
                and node.func.id == "translate_task_fixed_op" and len(node.args) > 1
                and isinstance(node.args[1], ast.Constant)
                and isinstance(node.args[1].value, str)):
            out.add(node.args[1].value)
    return out


def ops_passed_as_a_variable() -> int:
    """op が**変数**の呼び出しの数（どの op でも来うる口）。"""
    n = 0
    for node in (x for tree in _trees() for x in ast.walk(tree)):
        if (isinstance(node, ast.Call) and isinstance(node.func, ast.Name)
                and node.func.id == "translate_task_fixed_op" and len(node.args) > 1
                and not isinstance(node.args[1], ast.Constant)):
            n += 1
    return n


#: 文字どおり渡す op 1 つずつの決めと理由。★ 「足さない」も決めであって空白ではない。
DECIDED = {
    "SWAP": (True,
             "slot が a / b ── 名前から意味が出てこない。さらに行か列かは機械が実表を"
             "見て決める分担なので、座標を入れさせない 1 行が要る"),
    "SET_WHERE": (True,
                  "col（印を書く列）と cond_col（条件を見る列）が両方 col 系で取り違える。"
                  "書き込む値と閾値はスキーマに無く、機械が依頼文から取る"),
    "ADD_COLUMN": (True,
                   "args が空 ── 名前が言われていなければ空 {} にする、という指示は"
                   "スキーマの形からは出てこない。位置も機械が決める"),
    "SET_CELL_VALUE": (False,
                       "★ 足していない。2026-08-30 に『col の欄に書き込む値を入れる』"
                       "誤訳を実測しているが、直しは機械側（依頼文から実表でセルを解く）で"
                       "入っており、主経路では LLM の col は**使われない**（実機 3/3 で確認）。"
                       "注記の効果は第二段の層では測れた（18/30 → 30/30・陰性対照 20/20 不変）が、"
                       "col がそのまま使われる退避経路（args を丸ごと採る側）へは"
                       "8 回試して**到達できなかった** ── 未確認であって無害ではない。"
                       "到達できた日に足す"),
    "ADD_ROW": (False,
                "★ 足していない。at / values は名前が読めるうえ、values は別の関所が"
                "ふるいに掛けている。誤訳の実測が無いので、当て推量で 1 行足さない"),
}

#: 文字どおりの呼び出しは無いが注記を持つ op（変数の口から来る）。
DECLARED_REACHED_DYNAMICALLY = {
    "EXTRACT_COLUMNS": "文字どおり渡す呼び出しは無い。頷き・別名ヒットの読み直しが"
                       "op を変数で渡す口から来る（cols に座標を入れさせない 1 行）",
}


def test_the_denominator_comes_from_the_implementation():
    """★ 分母は実装から数える ── 固定 op の呼び出しを足した日に、ここへ来させる。"""
    lit = ops_passed_literally()
    assert lit, "★ 分母が空（下の検査が全部素通りする）"
    assert lit == set(DECIDED), (
        f"文字どおり渡す op と決めた op が食い違う: 決めていない {sorted(lit - set(DECIDED))} / "
        f"もう渡していないのに決めている {sorted(set(DECIDED) - lit)}")


def test_the_variable_call_sites_are_not_forgotten():
    """★ op が変数の口は「どの op でも来うる」── 在ることを忘れないための等号。

    ★ ここが 0 になったら、注記の分母は文字どおりの 5 op で閉じる（台帳を書き直す合図）。
    """
    assert ops_passed_as_a_variable() == 3, (
        f"op を変数で渡す呼び出しが {ops_passed_as_a_variable()} 箇所になった ── "
        "注記の分母の考え方が変わる。この試験の docstring と台帳を見直すこと")


def test_every_decision_has_a_reason():
    """★ 「足さない」にも理由を書かせる ── 空白は『まだ見ていない』と読めてしまう。"""
    assert DECIDED
    for op, (_has, why) in DECIDED.items():
        assert str(why).strip(), f"{op} の理由が空"


def test_the_ledger_matches_the_product():
    """★ 台帳と製品の注記が一致すること（片方だけ直さない）。"""
    want = {op for op, (has, _w) in DECIDED.items() if has} | set(DECLARED_REACHED_DYNAMICALLY)
    got = set(ailine._OP_SCHEMA_NOTES)
    assert got == want, f"注記 {sorted(got)} ≠ 台帳 {sorted(want)}"


def test_every_note_declares_what_the_llm_must_not_fill_in():
    """★★ 注記は説明ではなく**分担の宣言** ── 「入れない」を必ず含むこと。

    4 本すべてがその形だった（2026-09-17）。説明だけの 1 行を足し始めると、
    第二段の prompt が「読み物」になり、OPS_DOC を太らせた時と同じ道を辿る。
    """
    assert ailine._OP_SCHEMA_NOTES, "★ 注記が 1 本も無い"
    for op, note in ailine._OP_SCHEMA_NOTES.items():
        assert "入れない" in note, (
            f"{op} の注記が「何を入れないか」を言っていない ── "
            "第二段の注記は分担の宣言であって、slot の説明ではありません")


def test_a_note_only_exists_for_an_op_that_has_a_schema():
    """★ 効かない注記を置かない（存在しない op の注記は誰にも読まれない）。"""
    for op in ailine._OP_SCHEMA_NOTES:
        assert op in ailine.OP_SCHEMA, f"知らない op の注記: {op}"


def test_the_note_actually_reaches_the_second_stage_prompt():
    """★★ 宣言どうしで閉じない ── 注記が**本当に prompt に載る**ことを実物で見る。

    ★ 注記の在否を dict で見るだけだと、_op_schema_doc が注記を読まなくなった日に
      気づけない（「在っても鳴らない」の形）。文面が組み上がる側から確かめる。
    """
    for op, note in ailine._OP_SCHEMA_NOTES.items():
        doc = ailine._op_schema_doc(op)
        assert note in doc, f"{op} の注記が第二段のスキーマ文に載っていない"
    plain = [o for o in ailine.OP_SCHEMA if o not in ailine._OP_SCHEMA_NOTES]
    assert plain, "★ 注記の無い op が 1 つも無い（陰性対照が取れない）"
    assert ailine._op_schema_doc(plain[0]).count("\n") == 0, (
        f"注記の無い op（{plain[0]}）のスキーマ文が 2 行以上ある ── "
        "注記以外の何かが混ざっている")
