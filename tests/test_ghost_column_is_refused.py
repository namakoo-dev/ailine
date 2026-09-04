"""依頼文に無い列名を、実在する別の列にすり替えて実行しないこと（2026-09-04）。

★★ 実測した事故: 在庫表（品名・棚・数量・備考）に「**原価**の列を削除して」と頼むと、
  7B が args を col="備考" に差し替え、機械は実在列なので通し、**✓ 機械検証済み**を
  出して備考を消した。**4/4 で再現**した。⚠ ですらなく ✓ が出る、静かなデータ喪失。

★ 関所（fabricated_subject_refusal）は在ったのに鳴らなかった。絞り込みが
  **「〜で」で名指しされた語**だけを掴む形だったため（「原価**の**列を」は素通り）。
  ★ そしてその真上に、当時の判断がこう書いてあった ──
    「正直に残す穴: 『金額を並べ替えて』（を）は掴めない ── **同じ事故が『を』で
      再来したら測り直す**」。**『の』で再来した。**

★ 当時『を』で絞れなかった理由は「売上高の列を作って」のような**新しい列名**まで
  掴むからだった。だがそれは**助詞の問題ではない** ── 分かれ目は
  「この op が列を新しく作るか」で、それは既に**宣言**（OP_WRITE_TARGET）に在る。
  助詞で当てずに宣言で分ける。新しい表は増やさない。

★ ここは純ロジックで測る（LLM も LibreOffice も要らない）。
"""
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

import ailine  # noqa: E402

_META = {
    "sheets": ["在庫"],
    "headers": {"在庫": ["品名", "棚", "数量", "備考"]},
    "header_rows": {"在庫": 1},
}


def _refusal(op, resolved, task):
    return ailine.fabricated_subject_refusal(op, resolved, _META, task, "在庫")


@pytest.mark.parametrize("task, resolved", [
    ("原価の列を削除して", {"col": "備考"}),
    ("原価列を削除して", {"col": "備考"}),
])
def test_a_column_the_request_never_named_is_refused(task, resolved):
    """★ 本命 ── 依頼文が名指しした列が無いのに、別の実在列で実行しようとしている。"""
    why = _refusal("DELETE_COLUMN", resolved, task)
    assert why, f"捏造を通した: {task}"
    assert "原価" in why and "備考" in why, why


@pytest.mark.parametrize("task, resolved", [
    ("備考の列を削除して", {"col": "備考"}),
    ("棚の列を削除して", {"col": "棚"}),
    ("列を削除して", {"col": "備考"}),          # 無指定 ── 従来どおり通す
])
def test_a_real_request_still_passes(task, resolved):
    assert _refusal("DELETE_COLUMN", resolved, task) is None, task


@pytest.mark.parametrize("op, task, resolved", [
    ("ADD_COLUMN", "売上高の列を追加して", {"name": "売上高"}),
    ("COMPUTE_COLUMN", "数量と単価をかけた金額の列を作って",
     {"operands": ["数量", "単価"], "target": "金額"}),
    ("SPLIT_CELL", "備考の列を読点で分けて", {"col": "備考"}),
])
def test_ops_that_create_a_column_may_name_one_that_does_not_exist(op, task, resolved):
    """★ 列を**新しく作る** op は、無い名前を言うのが正しい。ここを断ってはいけない。"""
    assert _refusal(op, resolved, task) is None, f"{op}: {task}"


def test_the_old_de_form_still_fires():
    """★ 退行の番人 ── 2026-08-24 に直した「〜で」の形が、今も掴めること。"""
    meta = {"sheets": ["S"], "headers": {"S": ["日付", "取引先", "商品", "数量", "単価"]},
            "header_rows": {"S": 1}}
    why = ailine.fabricated_subject_refusal(
        "SORT", {"col": "数量"}, meta, "金額で降順に並べ替えて", "S")
    assert why and "金額" in why, why


def test_the_split_is_driven_by_the_declaration_not_by_a_list():
    """★ 「列を作る op」を手で列挙していないこと（列挙は漏れる ── この repo が 3 度踏んだ形）。

    宣言（OP_WRITE_TARGET）から引いていれば、新しい op が増えても自動で正しい側に入る。
    """
    creates = {op for op, wt in ailine.OP_WRITE_TARGET.items()
               if ailine.WRITE_NEW_COLUMN in (getattr(wt, "writes", None) or ())}
    assert creates == {"ADD_COLUMN", "COMPUTE_COLUMN", "LOOKUP_FILL", "SPLIT_CELL"}, creates
    src = (Path(__file__).resolve().parent.parent
           / "src" / "ailine" / "__init__.py").read_text(encoding="utf-8")
    i = src.index("def fabricated_subject_refusal")
    body = src[i:src.index("\ndef ", i + 10)]
    assert "OP_WRITE_TARGET" in body, "宣言から引いていない（手書きの一覧になっている）"
