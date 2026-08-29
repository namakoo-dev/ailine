# 1 セルへの書き込みを、機械が実表から解く ── 2026-08-29。
# Namakoo「丸山重工の行を作ってから『丸山重工の右にPCパーツ』が動作しなかった」
#
# ★★ 両方のモデルで実測した（同じ表・同じ指示・各 3 回）:
#
#   指示                          qwen2.5-coder:7b        gemma4:e4b-it-qat
#   丸山重工の右にPCパーツ        SPLIT_CELL 3/3          ADD_ROW at:1 2/3・CLARIFY 1/3
#   丸山重工の項目をPCパーツに    OUT_OF_VOCAB 3/3        SET_COLUMN_VALUE 2/3（列を全部潰す）
#   8行目の項目にPCパーツと入れて ADD_ROW 3/3（行を挿す） ADD_ROW 3/3（同じく）
#
#   ★ **どちらのモデルも直せない**。3 つとも「1 セルに書く」だけの依頼で、
#     機械は既に答えを持っている（丸山重工は 8 行目・項目は 2 列目）。
#     誰も表に訊いていなかっただけ ── モデルを替えても、画像を見せても直らない。
#
# 契約:
#   ① 行: 行番号、または依頼文に literal で在る**実在の値**（1 行に限る）
#   ② 列: 見出しの名前 / A1 の列名 / **行の名前のセルからの相対**（右・隣・左）
#   ③ 値: 第二段が出さない回があるので、機械が依頼文から**引き算で**切り出す
#   ④ 引用符を要求しない（引用符は道具の都合であって、人の書き方の問題ではない）
#   ★★ 恒真殺し: 正当な依頼を横取りしないこと（実測で 2 種類やらかした）

import sys
from pathlib import Path

import openpyxl
import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
import ailine  # noqa: E402

HEADERS = ["取引先", "項目", "件数", "金額"]
ROWS = [["丸和物流", "配送業務一式", 12, 57600],
         ["みどり建設", "内装工事", 9, 64800],
         ["丸山重工", None, None, None]]


def _book(tmp_path, name="b.xlsx", headers=None, rows=None):
    p = tmp_path / name
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "請求"
    ws.append(headers or HEADERS)
    for r in (rows or ROWS):
        ws.append(r)
    wb.save(p)
    return p


def _meta(path, headers=None):
    return {"sheets": ["請求"], "headers": {"請求": list(headers or HEADERS)},
            "header_rows": {"請求": 1}, "path": str(path)}


# --- ①② 行と列を実表から解く ----------------------------------------------------------

@pytest.mark.parametrize("task,want_row,want_col", [
    ("丸山重工の右にPCパーツ", 4, 2),        # 名前のセルの 1 つ右
    ("丸山重工の隣にPCパーツ", 4, 2),
    ("丸山重工のとなりにPCパーツ", 4, 2),
    ("丸山重工の項目をPCパーツにして", 4, 2),  # 見出しの名前
    ("4行目の項目にPCパーツと入れて", 4, 2),   # 行番号
    ("丸山重工のB列にPCパーツ", 4, 2),        # A1 の列名
    ("内装工事の左にみどり工業", 3, 1),        # 左（名前のセルの 1 つ左）
])
def test_the_cell_is_resolved_from_the_real_table(tmp_path, task, want_row, want_col):
    p = _book(tmp_path)
    got = ailine.resolve_cell_target_from_task(task, _meta(p), "請求")
    assert got is not None, task
    assert (got[0], got[1]) == (want_row, want_col), got


@pytest.mark.parametrize("task", [
    "担当を全部佐藤にして",              # 行を指していない
    "金額で降順に並べ替えて",            # そもそも 1 セルの話ではない
    "丸和物流の件数と金額を入れ替えて",  # 見出しが 2 つ ── 決めない
])
def test_it_refuses_to_decide_when_the_request_is_not_one_cell(tmp_path, task):
    p = _book(tmp_path)
    assert ailine.resolve_cell_target_from_task(task, _meta(p), "請求") is None, task


# --- ★★ 恒真殺し: 正当な依頼を横取りしない（実測で 2 種類やらかした）-------------------

def test_a_word_that_describes_the_operation_is_not_a_row_name(tmp_path):
    """★★ 実測（既存の検体が捕まえた）: 「金額の**合計**を一番下に出して」の『合計』を
       行の名前と読み、値に『一番下に出』という**切れ端**を書こうとした。
    ★ 歯止め: 人がセルを指すときは「**〜の**」と言う。『合計を』は操作の説明。"""
    p = _book(tmp_path, headers=["商品", "金額"],
               rows=[["りんご", 100], ["みかん", 200], ["合計", 300]])
    got = ailine.resolve_cell_target_from_task(
        "金額の合計を一番下に出して", _meta(p, ["商品", "金額"]), "請求")
    assert got is None, got


def test_a_number_in_the_request_is_not_a_row_name(tmp_path):
    """★★ 実測: 「数量を10にして」の『10』が、10 という値を持つ行に当たっていた。
       依頼文に数字は普通に出る ── 数字は行の名前にしない。"""
    p = _book(tmp_path, headers=["商品", "数量"],
               rows=[["りんご", 10], ["みかん", 20]])
    got = ailine.resolve_cell_target_from_task(
        "数量を10にして", _meta(p, ["商品", "数量"]), "請求")
    assert got is None, got


def test_the_brakes_are_both_present():
    src = (REPO / "src" / "ailine" / "__init__.py").read_text(encoding="utf-8")
    i = src.index("def resolve_cell_target_from_task(")
    seg = src[i:i + 3000]
    assert "_is_number_like(name)" in seg, "数字を行の名前にしない歯止めが無い"
    assert 'f"{name}の" not in' in seg, "『〜の』の歯止めが無い"


# --- ③ 値は機械が引き算で切り出す ------------------------------------------------------

@pytest.mark.parametrize("task,want", [
    ("丸山重工の右にPCパーツ", "PCパーツ"),
    ("丸山重工の項目をPCパーツにして", "PCパーツ"),
    ("丸山重工の隣にPCパーツ", "PCパーツ"),
    ("4行目の項目にPCパーツと入れて", "PCパーツ"),
    ("丸山重工のB列にPCパーツ", "PCパーツ"),
])
def test_the_value_is_carved_out_of_the_request(task, want):
    """★ 実測: 第二段は row/col だけ返して **value を返さない**回がある（両モデルとも）。
       LLM が出さないなら機械が出す ── 既に分かっている物を依頼文から引く。"""
    got = ailine.bare_value_from_task(task, "丸山重工", "項目", HEADERS)
    assert got == want, got


def test_a_value_that_is_not_in_the_request_is_never_written():
    """★ A' 原則: 依頼文に無い値は作らない。"""
    assert ailine.value_written_in_task("丸山重工の右にPCパーツ", "未定", HEADERS) is None
    assert ailine.value_written_in_task("丸山重工の右にPCパーツ", "PCパーツ", HEADERS) == "PCパーツ"
    # 見出しの語をそのまま値にしない
    assert ailine.value_written_in_task("丸山重工の項目を…", "項目", HEADERS) is None


def test_a_carved_value_must_be_one_continuous_piece():
    """★ 切れ端を継ぎ足した幽霊の値を作らない（連続していなければ決めない）。"""
    assert ailine.bare_value_from_task("丸山重工の項目を", "丸山重工", "項目", HEADERS) is None


# --- ④ 引用符を要求しない --------------------------------------------------------------

def test_an_unquoted_value_is_accepted_when_it_is_in_the_request(tmp_path):
    p = _book(tmp_path)
    ok, r, _i, err = ailine.verify_dsl_args(
        "SET_CELL_VALUE",
        {"row": "丸山重工", "col": "項目", "value": "PCパーツ", "row_number": 4},
        _meta(p), task="丸山重工の右にPCパーツ")
    assert ok, err
    assert r["value"] == "PCパーツ" and r["_row_index"] == 4


def test_a_value_the_request_never_mentions_is_still_refused(tmp_path):
    """★ 緩めた分の歯止め: 依頼文に無い値は、引用符が無くても通さない。"""
    p = _book(tmp_path)
    ok, _r, _i, err = ailine.verify_dsl_args(
        "SET_CELL_VALUE",
        {"row": "丸山重工", "col": "項目", "value": "でっちあげ", "row_number": 4},
        _meta(p), task="丸山重工の右にPCパーツ")
    assert not ok and "読み取れません" in err, err


# --- 門は計画の長さで閉じない ----------------------------------------------------------

def test_the_gate_does_not_close_on_plan_length():
    """★★ 実測: 同じ依頼で 1 段と 2 段の計画が返り分かれ、**2 段の回だけ**素通りして
       いた（4 回中 2 回が別々の結果）。長さは依頼の性質ではなくモデルの気分。"""
    src = (REPO / "src" / "ailine" / "__init__.py").read_text(encoding="utf-8")
    i = src.index("_cell = resolve_cell_target_from_task(")
    seg = src[max(0, i - 900):i]
    assert "any(_already_writes_one_cell(st) for st in plan)" in seg, seg[-400:]
    assert "len(plan) == 1" not in seg, "計画の長さで門を閉じている"


# --- 途中まで入力した行があっても、見出しの検出が止まらない ----------------------------

def test_a_half_filled_row_does_not_break_header_detection():
    """★★ 2026-08-29（Namakoo が実測・基本操作が丸ごと止まった）:
       「丸山工業／PCパーツ」だけ埋めた行を作ったら、そのシートで**何も**できなくなった
       （？ 見出しが何行目か分かりません）。
    ★ その行は「非空セルが全部文字列」で下の行に数字がある ── 見出しの条件を満たす。
    ★ 一般則で切る: **本物の見出しは表の幅いっぱいに並ぶ**。
       途中まで入力した行は 1〜2 セルしか埋まっていないので幅が違う。"""
    rows = {1: {"nonempty": 7, "str": 7, "bold": 1}}          # 見出し（7 列）
    for r in range(2, 7):
        rows[r] = {"nonempty": 7, "str": 3, "bold": 0}        # データ
    rows[7] = {"nonempty": 2, "str": 2, "bold": 0}            # 途中まで入力した行
    rows[8] = {"nonempty": 7, "str": 3, "bold": 0}
    got, confident = ailine.detect_header_row({"rows": rows})
    assert (got, confident) == (1, True), (got, confident)


def test_two_equally_wide_candidates_are_still_ambiguous():
    """★ 黙りすぎない側の対: 同じ幅の候補が並ぶ（見出し 2 段・表が縦に 2 つ）なら
       今までどおり**決めない**。幅で切れるのは、幅が違うときだけ。"""
    rows = {1: {"nonempty": 2, "str": 2, "bold": 0},
             2: {"nonempty": 2, "str": 1, "bold": 0},
             3: {"nonempty": 2, "str": 2, "bold": 0},
             4: {"nonempty": 2, "str": 1, "bold": 0}}
    got, confident = ailine.detect_header_row({"rows": rows})
    assert (got, confident) == (None, False), (got, confident)
