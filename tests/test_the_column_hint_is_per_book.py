"""`--column` の指定は**冊ごとの手がかり**であること（2026-09-22）。

★★ 事故（盲検 6 体目 ⑥・**致命**・こちらで 2/2 再現）:
  道具が自分で「`--column 借方勘定科目=科目` を付けてもう一度実行してください」と案内し、
  **その通りに打つと過去の仕訳の側が全部落ちた**。
    今回 = 自社の精算書（見出し『科目』）
    過去 = 会計ソフトの書き出し（見出し『借方科目』）
  ★ **導線が嘘**になっていた ── この repo の線では「導線が嘘なら無い方がまし」。

★★ 2 つの原因が重なっていた:
  ① 指定が**全冊に同時に掛かり**、当たらない冊は読めずに落ちた。
     旧コメントは「実際の冊は同じ書き出しなので、見出しも同じ」と書いていたが、
     これは**出所が違う 2 種を突き合わせる道具**なので、揃う方が珍しい。
     買い手:「うちの精算書と会計ソフト出力の見出しが一致することはまずありません」
  ② 同じ役割を 2 回教えると**後勝ち**で上書きされ、どちらを教えても片方が落ちた。

★ 直し: 役割ごとに**候補の並び**を持ち、冊ごとに当たったものを使う。
  当たらなければ自動照合へ落ちる ── ★ ただし**黙って別の列を読まない**。
  落ちたことは `JournalBook.notes` で名指しして画面に出す。

★ 別名は**足していない**。設計文書（DESIGN-20260913）が
  「実物で確かめるまで列名を看板にしない」と書いており、『借方科目』を実物で
  確かめた資料が無いため。★ 語彙でなく**仕組み**で解いた。
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from ailine_core import accounts_core  # noqa: E402


def _rows(headers, *data):
    return [(1, list(headers))] + [(i + 2, list(r)) for i, r in enumerate(data)]


def test_the_same_role_can_be_taught_more_than_one_header():
    """★ 後勝ちで上書きしない ── 2 冊ぶんの見出しを同時に持てること。"""
    got, why = accounts_core.parse_column_overrides(
        ["借方勘定科目=科目", "借方勘定科目=借方科目"])
    assert why is None, why
    assert got["借方勘定科目"] == ["科目", "借方科目"], got


def test_each_book_uses_the_hint_that_actually_lands():
    """★★ 本体 ── 同じ指定で、見出しの違う 2 冊がどちらも読めること。"""
    overrides, _ = accounts_core.parse_column_overrides(
        ["借方勘定科目=科目", "借方勘定科目=借方科目"])

    # ★ 必須列（借方勘定科目・借方金額）は両方そろえる ── 欠けると解決部は断り、
    #   header_map は空で返る（推測で先へ進まない設計）。測りたいのはそこではない。
    a_notes: list = []
    _hr, _hd, a_map, a_ref = accounts_core.resolve_accounts_columns(
        _rows(["社員", "日付", "科目", "借方金額"],
              ["佐藤", "2026-08-02", "旅費交通費", 1280]),
        overrides, notes=a_notes)
    b_notes: list = []
    _hr, _hd, b_map, b_ref = accounts_core.resolve_accounts_columns(
        _rows(["日付", "摘要", "借方科目", "借方金額"],
              ["2026-08-03", "用紙", "消耗品費", 3200]),
        overrides, notes=b_notes)

    assert a_ref is None or "借方勘定科目" not in a_ref, a_ref
    assert b_ref is None or "借方勘定科目" not in b_ref, b_ref
    assert a_map.get(accounts_core.DEBIT_ACCOUNT) == 3, (a_map, a_ref)
    assert b_map.get(accounts_core.DEBIT_ACCOUNT) == 3, (b_map, b_ref)
    assert not a_notes and not b_notes, (a_notes, b_notes)


def test_a_hint_that_does_not_land_says_so_and_falls_back():
    """★ 黙って別の列を読まない ── 落ちたことを名指しする。"""
    overrides, _ = accounts_core.parse_column_overrides(["借方勘定科目=存在しない見出し"])
    notes: list = []
    accounts_core.resolve_accounts_columns(
        _rows(["日付", "摘要", "借方勘定科目", "借方金額"],
              ["2026-08-03", "用紙", "消耗品費", 3200]),
        overrides, notes=notes)
    assert notes, "当たらなかったことを黙っている"
    assert "存在しない見出し" in notes[0] and "自動で探しました" in notes[0], notes


def test_two_columns_with_the_same_name_still_refuse():
    """★ 緩めていない ── 同じ見出しが 2 列ある冊は、今までどおり決めない。"""
    overrides, _ = accounts_core.parse_column_overrides(["借方勘定科目=科目"])
    _hr, _hd, got, refusal = accounts_core.resolve_accounts_columns(
        _rows(["科目", "科目", "金額"], ["a", "b", 1]), overrides, notes=[])
    assert not got and refusal and "2 列あります" in refusal, (got, refusal)


def test_the_roles_still_come_from_the_declaration():
    """★ 役割の顔ぶれは COLUMN_ALIASES が唯一の出どころ（手で並べない）。"""
    _got, why = accounts_core.parse_column_overrides(["そんな役割=科目"])
    assert why and "役割はありません" in why, why
    for role in accounts_core.COLUMN_ALIASES:
        assert role in why, f"{role} が案内に出ていない"


# ---------------------------------------------------------------------------
# ★★ 2026-09-23（盲検 6 体目 ⑥ の残り・導線の台帳を歩いて発見・2 回とも再現）
#   09-22 の直しで「冊ごとに違う見出しを教えられる」仕組みは入ったが、**案内がそれを教えて
#   いなかった**。1 回目は今回の冊の分だけ「`--column 借方勘定科目=科目` を付けてもう一度」と
#   言い、従うと過去の冊（『借方科目』）で落ち、2 回目の画面に次に打てる形が無かった。
# ---------------------------------------------------------------------------

import re  # noqa: E402

TODAY = [["社員", "日付", "科目", "借方金額", "摘要"],
         ["佐藤", "2026-08-02", "旅費交通費", 1280, "電車代"]]
PAST = [["日付", "摘要", "借方科目", "借方金額"],
        ["2026-07-03", "電車代", "旅費交通費", 980]]


def _accounts(tmp_path, capsys, today, past, extra=()):
    import openpyxl
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    import ailine
    for name, rows in (("今回.xlsx", today), ("過去.xlsx", past)):
        wb = openpyxl.Workbook()
        for r in rows:
            wb.active.append(r)
        wb.save(tmp_path / name)
    rc = ailine.main(["accounts", str(tmp_path / "今回.xlsx"), "--past", str(tmp_path / "過去.xlsx"),
                      "--out", str(tmp_path / f"候補{len(extra)}.xlsx"), *extra])
    return rc, capsys.readouterr().out


def _flags(screen):
    return [x for f in re.findall(r"`--column ([^`]+)`", screen) for x in ("--column", f)]


def test_the_first_screen_names_every_book_in_one_line(tmp_path, capsys):
    """★★ 本体: 1 回目の画面で、全部の冊の分を 1 行の `--column` にまとめて出す。"""
    rc, out = _accounts(tmp_path, capsys, TODAY, PAST)
    assert rc == 4, out
    assert "全部の冊を読むには" in out, out
    assert "`--column 借方勘定科目=科目`" in out and "`--column 借方勘定科目=借方科目`" in out, out
    assert "その列でよければ" not in out, "1 冊ぶんの「もう一度」がまだ出ている（従うと別の冊で落ちる）"


def test_the_one_line_typed_as_is_reaches(tmp_path, capsys):
    """★ その 1 行を**そのまま**打つと通る（案内が嘘でない）。"""
    _rc, out = _accounts(tmp_path, capsys, TODAY, PAST)
    rc2, out2 = _accounts(tmp_path, capsys, TODAY, PAST, _flags(out.split("全部の冊を読むには")[1]))
    assert rc2 == 0, out2


def test_a_book_that_cannot_be_selected_is_named_not_looped(tmp_path, capsys):
    """★ 同じ見出しが 2 列ある冊は `--column` では選べない ── 正直に言い、「全部」と言わない。
       旧版は「外せ ⇄ 付けろ」の往復になった（accounts_core.py:345 の家系）。"""
    past_dup = [["日付", "科目", "借方金額", "科目"], ["2026-07-03", "旅費交通費", 980, "現金"]]
    rc, out = _accounts(tmp_path, capsys, TODAY, past_dup)
    assert rc == 4
    assert "全部の冊を読むには" not in out, "読めない冊が残るのに『全部』と言った"
    assert "全部ではありません" in out and "写しを渡してください" in out, out
    rc2, out2 = _accounts(tmp_path, capsys, TODAY, past_dup, _flags(out))
    assert "写しを渡してください" in out2, f"従った後の画面に次の手が無い（往復の入口）\n{out2}"


def test_a_software_export_header_counts_as_near():
    """『借方科目』は『借方勘定科目』の近い見出し（部分列）。『借方金額』はならない。"""
    near = {h for _i, h in accounts_core.near_headers(["借方科目", "借方金額", "摘要"], "借方勘定科目")}
    assert near == {"借方科目"}, near
