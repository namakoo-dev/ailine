# 「4行目の下に丸山工業の行を作って」── 2026-08-29。Namakoo「行の追加が出来なくなってる」
#
# ★★ 実測（デモ材料そのもの・画面で）:
#     4 行目に**空行**が挿さり、丸山工業 はどこにも書かれなかった。
#   ★ ところが **同じ意味を名前で言えば動いていた**:
#       「ヤマノ食品の下に丸山工業の行を作って」→ 5行目・値も入る（✓）
#       「4行目の下に丸山工業の行を作って」    → 機械が黙る → LLM の 4 が通る（✗）
#   ★ **指し方が名前か番号かで結果が変わる** ── 行と列の非対称と同じ形の非対称。
#
# 真因は 2 つ、どちらも「機械が黙った」ことから来ていた:
#   ① `resolve_row_anchor` が「行番号は名前ではない」を「**何も決めない**」と書いていた。
#      表に訊く必要すら無い（番号と向きが揃っている＝引き算で出る）のに、諦めていた。
#   ② 位置が出ないので `insert_rows_should_have_been_add_row` の証拠①が立たず、
#      値つきの行（ADD_ROW）へ回らなかった ── 空行に落ちる道が開きっぱなしだった。
#
# ★★ さらに①を直した直後、今度は審査の側が誤爆した:
#     「⚠ 対象『5』は依頼文の語と機械照合できません（依頼文が指しているのは: 4行目）」
#   ★ **精密に言うほど怒られる**形。三項（依頼／宣言／実体）は壊れていない ── 依頼は
#     『4行目』で確かに在り、機械がしたのは引き算だけで、それは解釈行に出ている。
#   ★ 黙らせるのではなく、**照合する相手を導出元に変える**（画面のシート選択を
#     「語の一致より強い証拠」と認めたのと同じ線）。
#
# ★ 片配線を作らないために、引き算は `row_number_anchor` **1 つだけ**が持つ。
#   位置を決める側（resolve_row_anchor）と、その位置を審査する側（_subject_slots）が
#   同じ関数に同じ問いを出す。下の test_one_source_of_the_arithmetic が変異で縛る。

import sys
from pathlib import Path

import openpyxl
import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
import ailine  # noqa: E402
from ailine_core import subject  # noqa: E402

HEADERS = ["取引先", "項目", "件数"]
ROWS = [["丸和物流", "配送業務一式", 12], ["近江スチール", "鋼材加工", 5],
        ["ヤマノ食品", "食品仕入", 28], ["北斗精機", "精密部品", 3]]


@pytest.fixture()
def meta(tmp_path):
    p = tmp_path / "b.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "請求"
    ws.append(HEADERS)
    for r in ROWS:
        ws.append(r)
    wb.save(p)
    return {"sheets": ["請求"], "headers": {"請求": list(HEADERS)},
            "header_rows": {"請求": 1}, "path": str(p)}


# --- ① 引き算そのもの ------------------------------------------------------------------

@pytest.mark.parametrize("task, at, src", [
    ("4行目の下に丸山工業の行を作って", 5, 4),
    ("4行目の後に丸山工業の行を作って", 5, 4),
    ("4行目の上に丸山工業の行を作って", 4, 4),
    ("2行目の前に1行挿入して", 2, 2),
    ("４行目の下に丸山工業の行を作って", 5, 4),      # 全角
    # ★★ 2026-08-29（Namakoo）:「4行目と5行目は両方ともヤマノ食品。取引先で指定は
    #   出来ない」── 中身で指せない表では、人は番号でしか言えない。
    ("4行目と5行目の間に丸山工業の行を作って", 5, 4),
])
def test_a_row_number_and_a_direction_are_enough(task, at, src):
    """★ 「4行目の下」は 5 行目 ── 表を読まずに出る。ここで諦めていたのが真因①。"""
    got_at, got_src, note = ailine.row_number_anchor(task)
    assert (got_at, got_src) == (at, src), (got_at, got_src, note)
    assert str(at) in note and str(src) in note, note


@pytest.mark.parametrize("task", [
    "ヤマノ食品の下に丸山工業の行を作って",     # 名前で指している（表に訊く側の仕事）
    "4行目の行を削除して",                     # 向きが無い
    "7行目にヤマノ食品を追加して",             # 相対ではない
    "4行目と6行目の間に入れて",                # 隣り合っていない（決めない）
    "丸和物流と近江スチールの間に北斗精機を作って",   # 名前で指している
])
def test_it_stays_quiet_when_it_is_not_arithmetic(task):
    assert ailine.row_number_anchor(task) == (None, None, "")


# --- ② 名前で言っても番号で言っても同じ場所 ---------------------------------------------

def test_the_same_place_whichever_way_you_say_it(meta):
    """★ これが今回の芯 ── 指し方で結果が変わってはいけない。"""
    by_name, _n1 = ailine.resolve_row_anchor("ヤマノ食品の下に丸山工業の行を作って", meta, "請求")
    by_number, _n2 = ailine.resolve_row_anchor("4行目の下に丸山工業の行を作って", meta, "請求")
    assert by_name == by_number == 5, (by_name, by_number)


def test_the_basis_is_shown(meta):
    _at, note = ailine.resolve_row_anchor("4行目の下に丸山工業の行を作って", meta, "請求")
    assert note and "4行目" in note and "5行目" in note, note


# --- ③ 値つきの行へ回る（空行に落ちない）------------------------------------------------

def test_it_routes_to_a_row_with_values(meta):
    """★ 真因② ── 位置が出なかったので空行の挿入のままだった。"""
    why = ailine.insert_rows_should_have_been_add_row(
        "4行目の下に丸山工業の行を作って", {}, meta, "請求")
    assert why and "5行目" in why, why


def test_an_empty_row_on_purpose_is_still_an_empty_row(meta):
    """★ 黙りすぎていないこと: 空行が欲しい依頼を record に化けさせない。"""
    assert ailine.insert_rows_should_have_been_add_row(
        "2行目の前に1行挿入して", {}, meta, "請求") is None


# --- ④ 導出した位置を審査が誤爆しない ---------------------------------------------------

def _verdicts(task, at):
    slots = ailine._subject_slots("ADD_ROW", {"at": at, "_target_sheet": "請求"}, ["請求"], task)
    return subject.classify_slots(slots, task=task, columns=HEADERS, header_row=1,
                                   sheets=["請求"])


def test_a_derived_row_matches_through_its_origin():
    """★★ 「⚠ 対象『5』は…（依頼文が指しているのは: 4行目）」が出ていた形。
       依頼は『4行目』で在る ── 機械の引き算は解釈行に出ている。①（照合できた）でよい。"""
    v = _verdicts("4行目の下に丸山工業の行を作って", 5)
    assert [x.tier for x in v] == [subject.MATCHED], [(x.slot.value, x.tier) for x in v]
    assert subject.contradiction_lines(v) == []


def test_the_origin_is_consumed_so_it_cannot_become_counter_evidence():
    """★ 導出元を消費しないと、『4行目』が「誰も拾わなかった語」として残り、
       同じ計画の他のスロットの反証に化ける（段またぎの台帳の性質）。"""
    slots = ailine._subject_slots("ADD_ROW", {"at": 5, "_target_sheet": "請求"}, ["請求"],
                                   "4行目の下に丸山工業の行を作って")
    c = subject.Consumed()
    subject.classify_slots(slots, task="4行目の下に丸山工業の行を作って", columns=HEADERS,
                            header_row=1, sheets=["請求"], consumed=c)
    assert 4 in c.rows, c.rows


def test_a_row_the_request_never_mentioned_is_still_challenged():
    """★★ ここが緩んだら意味が無い ── 導出元と**一致しない**位置は今までどおり ⚠。"""
    v = _verdicts("4行目の下に丸山工業の行を作って", 9)
    assert [x.tier for x in v] == [subject.CONTRADICTED], [(x.slot.value, x.tier) for x in v]
    assert "4行目" in subject.contradiction_lines(v)[0]


# --- ⑤ 引き算の産地は 1 つ -------------------------------------------------------------

def test_one_source_of_the_arithmetic():
    """★ 位置を決める側と審査する側が**同じ関数**を読んでいること。
       片方だけ賢くすると、また「名前なら通るが番号だと止まる」が戻ってくる。"""
    src = (REPO / "src" / "ailine" / "__init__.py").read_text(encoding="utf-8")
    assert src.count("def row_number_anchor(") == 1
    assert src.count("row_number_anchor(task)") == 2, "決める側／審査する側の 2 箇所のはず"
    # 引き算（+1）がこの関数の外に写し取られていないこと
    body = src[src.index("def row_number_anchor("):]
    body = body[:body.index(chr(10) + "def ", 10)]
    assert "n + 1 if after else n" in body


# --- ⑥ 中身で指せない表では、断りが番号の道を示す -----------------------------------

def test_when_the_name_is_ambiguous_it_points_at_the_row_numbers(tmp_path):
    """★★ 2026-08-29（Namakoo）:「どうしても中身でさせない場面が出てくる。例えば
       4行目と5行目は両方ともヤマノ食品。取引先で指定は出来ない」
    ★ 同じ名前が 2 行あるなら中身では指せない ── **断って終わる場所ではなく、
      行番号の道へ渡す場所**。候補の行番号は機械がもう知っている（人に数え直させない）。"""
    p = tmp_path / "dup.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "請求"
    ws.append(HEADERS)
    for r in ROWS[:2] + [["ヤマノ食品", "食品仕入", 28], ["ヤマノ食品", "冷蔵配送", 6]]:
        ws.append(r)
    wb.save(p)
    m = {"sheets": ["請求"], "headers": {"請求": list(HEADERS)},
         "header_rows": {"請求": 1}, "path": str(p)}
    at, note = ailine.resolve_row_anchor("ヤマノ食品の下に丸山工業の行を作って", m, "請求")
    assert at is None
    assert "4、5行目" in note, note
    assert "行番号で指してください" in note, note
    # ★ 示した道が実際に通ること（案内だけして行けない、をここで塞ぐ）
    assert ailine.resolve_row_anchor("4行目の下に丸山工業の行を作って", m, "請求")[0] == 5


# --- ⑦ 値が決まらない回は、空行を挿して ✓ を出さない -----------------------------------

def test_the_switch_is_wired_to_refuse_when_values_cannot_be_decided():
    """★★ 2026-08-29（効果検体で実測・「3行目と4行目の間に新品を作って」）:
       ここまで来た時点で機械は「**値を入れる行が欲しい依頼だ**」と分かっている
       （それが _why）。なのに値が決まらなかった回は、そのまま**空行を挿して ✓ を
       出して**いた ── 宣言（空行を挿す）と実体は一致するので検算は通る。
       だが依頼とは違う ── 三項のうち「依頼」を捨てた形の再演。
    ★ 壊す前に止まる。断りっぱなしにせず、通る言い方を必ず添える。
    """
    src = (REPO / "src" / "ailine" / "__init__.py").read_text(encoding="utf-8")
    i = src.index("_clean = add_row_values_from_request(")
    seg = src[i:i + 2400]
    assert "if _fixed and _clean:" in seg
    assert "？ 入れる値を依頼文から決められません" in seg, "値が無い回に断りが無い"
    assert "空の行が欲しいなら" in seg, "空行が欲しい人への出口が無い"
    # ★ 断りの直後に必ず抜けること（黙って先へ進まない）
    j = seg.index("？ 入れる値を依頼文から決められません")
    assert "return 3" in seg[j:j + 700], seg[j:j + 700]
