# -*- coding: utf-8 -*-
"""消えるものが在る削除は、消す前に聞く（2026-09-17・盲検 3 体目）。

★★ 起きたこと（製造業の購買担当・初見・盲検）:

    $ ailine run 発注台帳.xlsx "検収日の列を消して"
    消した中身（15 行）── 戻すなら ailine undo: …
    ✓ 発注台帳.xlsx は機械検証済みの内容です
    EXIT=0

  値の入った検収日が**確認なしに**消えて exit 0 の ✓。買い手の言葉:
  **「上書きより削除の方が怖いのに、厳しい方が緩い」**。

★★ 器官は在るが配線が無い、の形だった:
  ・破壊の関所（_confirm_overwrite_or_gate）は在った
  ・削除用の聞き文「削除しますか？」も、逃げ道の文（「削除を承知して続行する」）も
    **既に書かれていた**
  ・配線されていたのは **DELETE_ROWS で「名前が複数行に当たった」1 ケースだけ**
  ・列の削除も、行番号で指した行削除も素通り（実測:「3行目を削除して」で値 11 個が消え exit 0）
  ★ 2026-09-07 の決裁は「削除は取り返しがつかないので**必ず聞く**」。目の前に在った
    1 ケースにだけ配線された ── 直しは**宣言（WRITE_REMOVE）の家系ぜんぶ**に。

★ 鳴る条件は上書きと**対称**にする ── 消えるものが在る時だけ。空の列・空の行は黙って消す
  （関所が毎回鳴ると読まなくなる。1 体目が「★ が毎回出るので読まなくなった」と言っている）。
★ 数えるのは**値の個数**であって行数ではない（事後の表示は空行まで数えて「15 行」と出していた）。
"""
import sys
from pathlib import Path

import openpyxl
import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
import ailine  # noqa: E402


def _book(tmp_path, rows, name="b.xlsx"):
    p = tmp_path / name
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "台帳"
    for r in rows:
        ws.append(r)
    wb.save(p)
    return p


FULL = [["取引先", "金額", "担当"],
        ["丸和物流", 1000, "高橋"],
        ["みどり建設", 2000, "田中"]]


def _meta(path):
    return {"path": str(path), "sheets": ["台帳"], "header_rows": {"台帳": 1},
            "headers": {"台帳": ["取引先", "金額", "担当"]}}


def test_every_op_that_removes_is_in_this_test():
    """★★ 分母を宣言から引く ── 「消す」と宣言した op が増えた日に、ここへ来させる。

    ★ ここが無いと、3 つ目の削除 op が黙って関所の外に生まれる（今回がまさにそれ）。
    """
    removing = {op for op in ailine.OP_SCHEMA
                if ailine.WRITE_REMOVE in (
                    getattr(ailine.OP_WRITE_TARGET.get(op), "writes", ()) or ())}
    assert removing == {"DELETE_COLUMN", "DELETE_ROWS"}, (
        f"『消す』と宣言した op が変わった: {sorted(removing)} ── "
        "新しい op が関所に載っているか、この検体で確かめること")


def test_deleting_a_column_with_data_asks_first(tmp_path):
    """★ 事故そのもの: 値の入った列の削除は、聞かずに進まない。"""
    p = _book(tmp_path, FULL)
    resolved = {"col": "担当", "_target_sheet": "台帳"}
    out = ailine._verify_delete_column(resolved, set(), _meta(p), ["台帳"], {})
    assert out is None, f"断られた: {out}"
    assert resolved.get("_confirm_delete"), "値の入った列を、聞かずに消そうとしている"
    assert "2 件" in resolved["_confirm_delete"], resolved["_confirm_delete"]


def test_deleting_an_empty_column_stays_quiet(tmp_path):
    """★★ 陰性対照: 消えるものが無いなら黙る。

    ★ ここが無いと「削除は全部聞く」でも上の試験が通る ── それは関所が毎回鳴る形で、
      1 体目の買い手が「★ が毎回出るので読まなくなった」と言った失敗をもう一度やる。
    """
    p = _book(tmp_path, [["取引先", "金額", "備考"],
                         ["丸和物流", 1000, None],
                         ["みどり建設", 2000, None]])
    meta = {"path": str(p), "sheets": ["台帳"], "header_rows": {"台帳": 1},
            "headers": {"台帳": ["取引先", "金額", "備考"]}}
    resolved = {"col": "備考", "_target_sheet": "台帳"}
    out = ailine._verify_delete_column(resolved, set(), meta, ["台帳"], {})
    assert out is None, f"断られた: {out}"
    assert not resolved.get("_confirm_delete"), (
        f"空の列を消すのに関所が鳴った: {resolved.get('_confirm_delete')}")


def test_counting_the_rows_counts_values_not_rows(tmp_path):
    """★ 数えるのは値の個数（事後表示の『15 行』は空行まで数えていた）。"""
    p = _book(tmp_path, FULL)
    assert ailine._rows_existing_value_count(p, "台帳", 2, 1) == 3
    assert ailine._rows_existing_value_count(p, "台帳", 2, 2) == 6
    assert ailine._rows_existing_value_count(p, "台帳", 99, 1) == 0, "表の外は 0"
    assert ailine._rows_existing_value_count(None, "台帳", 2, 1) == 0, "冊が無ければ 0"
    assert ailine._rows_existing_value_count(p, "無いシート", 2, 1) == 0


def test_deleting_a_row_with_data_asks_first(tmp_path):
    """★ 同じ家系の片割れ ── 行番号で指した削除も聞く（実測で素通りしていた）。"""
    p = _book(tmp_path, FULL)
    resolved = {"at": "2", "_target_sheet": "台帳"}
    out = ailine._verify_add_row(resolved, set(), _meta(p), "2行目を削除して",
                                 ["台帳"], {"台帳": FULL[0]}, "DELETE_ROWS")
    assert out is None, f"断られた: {out}"
    assert resolved.get("_confirm_delete"), "値の入った行を、聞かずに消そうとしている"
    assert "3 件" in resolved["_confirm_delete"], resolved["_confirm_delete"]


def test_deleting_an_empty_row_stays_quiet(tmp_path):
    """★ 陰性対照: 空の行は黙って消す。"""
    p = _book(tmp_path, FULL + [[None, None, None]])
    resolved = {"at": "4", "_target_sheet": "台帳"}
    out = ailine._verify_add_row(resolved, set(), _meta(p), "4行目を削除して",
                                 ["台帳"], {"台帳": FULL[0]}, "DELETE_ROWS")
    assert out is None, f"断られた: {out}"
    assert not resolved.get("_confirm_delete"), (
        f"空の行を消すのに関所が鳴った: {resolved.get('_confirm_delete')}")


def test_the_named_multi_row_delete_still_speaks_for_itself(tmp_path):
    """★★ 2026-09-07 に足した「名前が複数行に当たった」回の文言を、上書きしていないこと。

    ★ 新しい配線が古い配線を**黙って置き換える**のが、この repo の 2 番目に多い壊れ方。
    """
    src = (REPO / "src" / "ailine" / "__init__.py").read_bytes().decode("utf-8")
    assert "に当てはまる {len(_rows0)} 行" in src or "に当てはまる" in src, (
        "名前で複数行に当たった時の文言が消えている")
    assert 'if not resolved.get("_confirm_delete"):' in src, (
        "行番号の関所が、名前の関所を上書きしないための条件が消えている")


@pytest.mark.parametrize("op", ["DELETE_COLUMN", "DELETE_ROWS"])
def test_the_gate_knows_how_to_say_delete(op):
    """★ 関所が『削除しますか？』と聞ける口を持っていること（文言は既に在った）。"""
    src = (REPO / "src" / "ailine" / "__init__.py").read_bytes().decode("utf-8")
    assert '削除しますか？' in src, "削除の聞き文が消えている"
    assert "削除を承知して続行する" in src, "削除時の逃げ道の文が消えている"
    assert op in ailine.OP_SCHEMA
