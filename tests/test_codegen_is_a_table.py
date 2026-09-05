"""Basic を組む所は、分岐の積層でなく**名簿**であること（2026-09-05）。

★★ 出所（盲検の査定・所見⑨の続き）: 査定者は `_translate_and_dispatch`（681 行）と
  並べて `codegen_dsl`（600 行）を「手つかず」と名指しした。実際に測ると:

      607 行のうち **536 行が `op == "…"` の 29 分岐**

  ★ `verify_dsl_args` を 641 件凍結して 1,735 → 227 行に割ったのと**同じ形**だった。
  ★ ただし読み直しの層（15 塊）と違い、ここは**順序に意味が無い**（op で振り分ける
    だけ・先に当たったものが残りを黙らせる、が無い）。だから切り出しで止めず
    **表まで畳める** ── 「op を足す」が名簿に 1 行になる。

★ 割る前に確かめたこと（順番を守る・README「番人を作ってから割った」）:
  ① ゴールデンが既に在るか → `tests/golden/f1_codegen/` に **78 本**、29 op すべて
  ② どの分岐も必ず `return` で抜けるか（落ちる枝が 1 つでもあれば表にできない）→ 0 件
  ③ 分岐の間に実行文が挟まっていないか → コメント 6 行のみ
  ★ 本文は **1 行も書き換えていない**（名前をすべて同じまま引数で受ける）。
    ゴールデン 78 本の md5 が割る前後で一致することを確かめた。

★ この番人が守るのは「畳んだ形が戻らないこと」── 次に op を足す人が
  `if op == "NEW":` を書き足せば、また 600 行へ向かって伸びていく。
"""
import ast
import inspect
import re
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))

import ailine  # noqa: E402

SRC = inspect.getsource(ailine)
TREE = ast.parse(SRC)


def _fn(name):
    return next(n for n in TREE.body if isinstance(n, ast.FunctionDef) and n.name == name)


def test_the_dispatcher_stays_small():
    """★ 本体は名簿を引くだけ ── ここが伸び始めたら、また分岐が生えている。"""
    n = _fn("codegen_dsl").end_lineno - _fn("codegen_dsl").lineno + 1
    assert n <= 60, f"codegen_dsl が {n} 行（名簿を引くだけのはず ── 分岐が戻っていないか）"


def test_no_op_branch_remains_in_the_dispatcher():
    """★ 畳んだ形が戻らないこと（次に op を足す人が if を書き足せる形にしない）。"""
    body = ast.get_source_segment(SRC, _fn("codegen_dsl"))
    found = re.findall(r'op == "([A-Z_]+)"', body or "")
    assert not found, (
        f"codegen_dsl に op ごとの分岐が戻っている: {found} ── "
        "新しい op は CODEGEN_BY_OP に 1 行足すこと")


def test_every_op_in_the_register_has_a_generator():
    """★ 名簿と実体が食い違わないこと（両向き）。"""
    table = ailine.CODEGEN_BY_OP
    assert len(table) >= 29, f"名簿が {len(table)} 件しかない"
    for op, gen in table.items():
        assert callable(gen), (op, gen)
        assert gen.__name__ == f"_codegen_{op.lower()}", (op, gen.__name__)


def test_the_register_covers_the_ops_that_need_code():
    """★ Basic を生成する op が名簿から漏れていないこと。

    ★ 「生成しない op」（フォルダ経路など）は対象外 ── 実際に codegen_dsl を
      通る op だけを分母にする（分母を入力側から作る）。
    """
    missing = [op for op in ailine.OP_SCHEMA
               if op not in ailine.CODEGEN_BY_OP and op not in ("CLARIFY", "FREEFORM")
               and not ailine.OP_META.get(op, {}).get("folder")]
    assert not missing, f"Basic を組む関数が無い op: {missing}"


def test_an_unknown_op_still_raises():
    """★ 名簿に無い op は、黙って None を返さず落ちること。"""
    import pytest
    with pytest.raises(ValueError):
        ailine.codegen_dsl("ぬるぽ", {}, {"headers": {"S": []}, "sheets": ["S"]})


def test_the_generators_do_not_reach_for_globals_that_used_to_be_locals():
    """★ 移動が「移動のみ」であること ── 各関数は必要な物を**引数で**受ける。

    ★ ここが緩むと、切り出した関数がモジュール変数を掴んで、呼ぶ場所で結果が変わる
      （読み直しの層で `_sheet_h` を中に閉じ込めたのと同じ規律）。
    """
    want = {"op", "resolved_args", "book_meta", "use_formula",
            "headers", "first_sheet", "header_row", "hr0", "wrap"}
    for op, gen in ailine.CODEGEN_BY_OP.items():
        got = set(inspect.signature(gen).parameters)
        assert got == want, f"{op}: 引数が {sorted(got)}（{sorted(want)} のはず）"
        for p in inspect.signature(gen).parameters.values():
            assert p.kind is inspect.Parameter.KEYWORD_ONLY, f"{op}: {p.name} が位置引数"
