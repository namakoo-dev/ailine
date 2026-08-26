# 表の基本操作 3 種（2026-08-26）── ADD_ROW / DELETE_ROWS / DELETE_COLUMN。
#
# ★ きっかけは Namakoo が GUI を触った実測:
#     「5行目に商品として梨を追加して。売上は600　原価は300」
#   → 行挿入 + 一括書換×3 に分解され、4 段とも別々の理由で落ちた。
#   21 op のどれにも「データを 1 行足す」「行/列を消す」が無かった。
#
# ★ Namakoo の指摘（設計の芯）:
#   「最後尾に追加するだけとは限らない。途中に行を追加する必要もあるし、
#     削除した場合はそこを詰めないといけない」
#   → insertByIndex / removeByIndex を使う（clearContents だと空行が残る）。
#     そして**そう書いたから正しい、では済まない** ── 事後条件で証明する。
#
# 契約:
#   ① 追加: 行数 +1・その行が宣言どおり・**at より下が 1 行ずれてそのまま**（押し下げた証拠）
#   ② 追加: at より上が 1 セルも変わらない
#   ③ 削除: 残りが「消した分を抜いた並び」と**順序ごと連続で**一致（詰めた証拠）
#   ④ 列削除: 他の列が 1 セルも変わらない
#   ⑤ 消した中身を機械の値として持ち、画面に出す（削除は差分に何も出ない操作）
#   ⑥ 実在しない列名・見出し行より上は断る（幻覚と骨格破壊の封鎖）
#   ⑦ 数値は数値のまま入る（文字列化すると下流の SUM が静かに壊れる）

import sys
from pathlib import Path

import openpyxl
import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
sys.path.insert(0, str(Path(__file__).resolve().parent))
import ailine  # noqa: E402

META = {"sheets": ["売上"], "headers": {"売上": ["商品", "売上", "原価"]},
         "header_rows": {"売上": 1}}
ROWS = [["商品", "売上", "原価"], ["りんご", 1200, 700],
         ["みかん", 800, 300], ["ぶどう", 1500, 900]]


def _book(tmp_path, rows=None, name="b.xlsx"):
    p = tmp_path / name
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "売上"
    for r in (rows or ROWS):
        ws.append(r)
    wb.save(p)
    return p


def _grid(p):
    ws = openpyxl.load_workbook(p)["売上"]
    return [[c.value for c in row] for row in ws.iter_rows()]


# --- ⑥ 関所 ---------------------------------------------------------------------------

@pytest.mark.parametrize("op,args,expect", [
    ("ADD_ROW", {"at": 5, "values": {"存在しない列": 1}}, "がこの表にありません"),
    ("ADD_ROW", {"at": 1, "values": {"商品": "梨"}}, "見出し行"),
    ("ADD_ROW", {"at": 5, "values": {}}, "読み取れません"),
    ("DELETE_ROWS", {"at": 1}, "見出し行"),
    ("DELETE_COLUMN", {"col": "利益"}, "がこの表にありません"),
])
def test_refusals(op, args, expect):
    ok, _resolved, _inf, err = ailine.verify_dsl_args(op, dict(args), META, task="t")
    assert not ok, f"{op} {args} を通した"
    assert expect in err, err


def test_a_real_request_passes():
    """誤爆防止: 実在する列・見出しより下なら通る。"""
    ok, resolved, _inf, err = ailine.verify_dsl_args(
        "ADD_ROW", {"at": 3, "values": {"商品": "梨", "売上": 600}}, META, task="t")
    assert ok, err
    assert resolved["_values_label"] == "商品=梨／売上=600"


# --- ①②⑦ 追加（途中に挿す）------------------------------------------------------------

def test_add_row_in_the_middle_pushes_the_rest_down(tmp_path):
    before = _book(tmp_path, name="before.xlsx")
    after = _book(tmp_path, [["商品", "売上", "原価"], ["りんご", 1200, 700],
                              ["梨", 600, 300],           # ← 3 行目に挿さった
                              ["みかん", 800, 300], ["ぶどう", 1500, 900]])
    args = {"at": 3, "values": {"商品": "梨", "売上": 600, "原価": 300}}
    status, reason = ailine.check_add_row(after, args, source_book=before)
    assert status == "pass", reason
    assert "元のまま" in reason


def test_add_row_catches_an_overwrite(tmp_path):
    """★ 恒真殺し: 押し下げずに**上書き**したら落ちること。"""
    before = _book(tmp_path, name="before.xlsx")
    after = _book(tmp_path, [["商品", "売上", "原価"], ["りんご", 1200, 700],
                              ["梨", 600, 300],           # みかんを潰した
                              ["ぶどう", 1500, 900]])
    args = {"at": 3, "values": {"商品": "梨", "売上": 600, "原価": 300}}
    status, reason = ailine.check_add_row(after, args, source_book=before)
    assert status == "fail", reason


def test_add_row_catches_a_wrong_value(tmp_path):
    before = _book(tmp_path, name="before.xlsx")
    after = _book(tmp_path, [["商品", "売上", "原価"], ["りんご", 1200, 700],
                              ["梨", 999, 300],
                              ["みかん", 800, 300], ["ぶどう", 1500, 900]])
    args = {"at": 3, "values": {"商品": "梨", "売上": 600, "原価": 300}}
    status, reason = ailine.check_add_row(after, args, source_book=before)
    assert status == "fail" and "売上" in reason, reason


def test_added_numbers_stay_numbers(tmp_path):
    """⑦ 数値が文字列で入ると下流の SUM が静かに壊れる。"""
    after = _book(tmp_path, [["商品", "売上", "原価"], ["梨", 600, 300]])
    args = {"at": 2, "values": {"商品": "梨", "売上": 600, "原価": 300}}
    status, _ = ailine.check_add_row(after, args, source_book=None)
    assert status == "warn"          # before 無しなら断定しない
    after2 = _book(tmp_path, [["商品", "売上", "原価"], ["梨", "600", 300]], name="s.xlsx")
    status2, reason2 = ailine.check_add_row(after2, args, source_book=None)
    assert status2 == "fail", f"文字列の 600 を通した: {reason2}"


# --- ③⑤ 行削除（詰める）----------------------------------------------------------------

def test_delete_rows_closes_the_gap(tmp_path):
    before = _book(tmp_path, name="before.xlsx")
    after = _book(tmp_path, [["商品", "売上", "原価"], ["りんご", 1200, 700],
                              ["ぶどう", 1500, 900]])       # みかんが消えて詰まった
    args = {"at": 3, "count": 1}
    status, reason = ailine.check_delete_rows(after, args, source_book=before)
    assert status == "pass", reason
    assert args["_deleted"] == [["みかん", 800, 300]], args.get("_deleted")


def test_delete_rows_catches_a_blank_left_behind(tmp_path):
    """★ 恒真殺し: 詰めずに空行を残したら落ちること（clearContents 実装への番人）。"""
    before = _book(tmp_path, name="before.xlsx")
    after = _book(tmp_path, [["商品", "売上", "原価"], ["りんご", 1200, 700],
                              [None, None, None], ["ぶどう", 1500, 900]])
    status, reason = ailine.check_delete_rows(after, {"at": 3, "count": 1}, source_book=before)
    assert status == "fail", reason


def test_delete_rows_catches_deleting_the_wrong_row(tmp_path):
    before = _book(tmp_path, name="before.xlsx")
    after = _book(tmp_path, [["商品", "売上", "原価"], ["みかん", 800, 300],
                              ["ぶどう", 1500, 900]])       # りんごを消した
    status, reason = ailine.check_delete_rows(after, {"at": 3, "count": 1}, source_book=before)
    assert status == "fail", reason


# --- ④⑤ 列削除 -------------------------------------------------------------------------

def test_delete_column_leaves_the_others_untouched(tmp_path):
    before = _book(tmp_path, name="before.xlsx")
    after = _book(tmp_path, [["商品", "売上"], ["りんご", 1200],
                              ["みかん", 800], ["ぶどう", 1500]])
    args = {"col": "原価"}
    status, reason = ailine.check_delete_column(after, args, source_book=before)
    assert status == "pass", reason
    assert args["_deleted"] == [[700], [300], [900]], args.get("_deleted")


def test_delete_column_catches_taking_a_neighbour(tmp_path):
    """★ 恒真殺し: 隣の列を巻き込んだら落ちること。"""
    before = _book(tmp_path, name="before.xlsx")
    after = _book(tmp_path, [["商品", "原価"], ["りんご", 700],
                              ["みかん", 300], ["ぶどう", 900]])   # 売上を消してしまった
    status, reason = ailine.check_delete_column(after, {"col": "原価"}, source_book=before)
    assert status == "fail", reason


# --- ⑤ 消した中身が画面に出る -----------------------------------------------------------

def test_deleted_content_reaches_the_report():
    from ailine_core import dsl_step
    # 本番の合流点が組み立てる助言に、消した中身が載ること
    src = Path(REPO / "src" / "ailine_core" / "dsl_step.py").read_text(encoding="utf-8")
    assert '"_deleted"' in src, "本番の合流点が消した中身を読んでいない"
    assert "戻すなら ailine undo" in src, "取り返しがつくことを言っていない"


# --- ⑧ 位置は相対で言われる（Namakoo が実測）--------------------------------------------
#
# 「みかんの下に梨を追加して」「みかんとぶどうの間に梨を追加して」がどちらも動かなかった。
# ★ 根: ADD_ROW は位置を**行番号**でしか受け取れないのに、人は相対で言う。
#   LLM に数えさせると外し、空行だけの INSERT_ROWS に落ちていた。
# ★ 分担を変えた: LLM は「誰の隣か」を言うだけ／**行番号は機械が実表を数えて決める**
#   （列名の解決を機械 3 段でやっているのと同じ形）。

def _anchor_meta(tmp_path):
    p = _book(tmp_path, name="anchor.xlsx")
    return p, {"sheets": ["売上"], "headers": {"売上": ["商品", "売上", "原価"]},
                "header_rows": {"売上": 1}, "path": str(p)}


@pytest.mark.parametrize("task,expect_at", [
    ("みかんの下に梨を追加して", 4),
    ("みかんの後に梨を追加して", 4),
    ("みかんとぶどうの間に梨を追加して", 4),
    ("りんごの上に梨を追加して", 2),
    ("ぶどうの前に梨を追加して", 4),
])
def test_relative_positions_resolve_to_a_row_number(tmp_path, task, expect_at):
    _p, meta = _anchor_meta(tmp_path)
    at, note = ailine.resolve_row_anchor(task, meta, "売上")
    assert at == expect_at, f"{task} → {at} / {note}"
    assert note and "行目" in note, note


def test_an_unknown_anchor_is_refused_by_name(tmp_path):
    """★ 推測で行を挿さない（静かに別の場所へ入るのが一番こわい）。"""
    _p, meta = _anchor_meta(tmp_path)
    at, note = ailine.resolve_row_anchor("すいかの下に梨を追加して", meta, "売上")
    assert at is None and "すいか" in note, note


def test_an_ambiguous_anchor_is_refused(tmp_path):
    """同じ名前が 2 行あるなら決めない。"""
    p = _book(tmp_path, [["商品", "売上", "原価"], ["みかん", 1, 1], ["みかん", 2, 2]],
               name="dup.xlsx")
    meta = {"sheets": ["売上"], "headers": {"売上": ["商品", "売上", "原価"]},
             "header_rows": {"売上": 1}, "path": str(p)}
    at, note = ailine.resolve_row_anchor("みかんの下に足して", meta, "売上")
    assert at is None and "2 行" in note, note


def test_the_machine_position_beats_the_model(tmp_path):
    """★ 実表を見た側が正しい ── LLM が別の行番号を言っても機械が上書きし、根拠を残す。"""
    _p, meta = _anchor_meta(tmp_path)
    ok, resolved, _inf, err = ailine.verify_dsl_args(
        "ADD_ROW", {"at": 99, "values": {"商品": "梨"}}, meta,
        task="みかんの下に梨を追加して")
    assert ok, err
    assert resolved["at"] == 4, resolved
    assert "みかん" in resolved["_at_basis"], resolved


def test_positional_values_are_named_by_the_machine(tmp_path):
    """★ LLM は values を**並び**で返すことがある。列名は機械が付ける（決めた対応は出す）。"""
    _p, meta = _anchor_meta(tmp_path)
    ok, resolved, _inf, err = ailine.verify_dsl_args(
        "ADD_ROW", {"at": 3, "values": ["梨", 600, 300]}, meta, task="3行目に足して")
    assert ok, err
    assert resolved["values"] == {"商品": "梨", "売上": 600, "原価": 300}, resolved
    assert resolved["_values_label"] == "商品=梨／売上=600／原価=300"


def test_too_many_positional_values_are_refused(tmp_path):
    _p, meta = _anchor_meta(tmp_path)
    ok, _r, _i, err = ailine.verify_dsl_args(
        "ADD_ROW", {"at": 3, "values": [1, 2, 3, 4]}, meta, task="3行目に足して")
    assert not ok and "4 個" in err, err


def test_insert_rows_with_values_in_the_task_is_re_read():
    """★ 「値の指定が在る」という証拠がある時だけ、行追加として読み直す（黙って変えない）。"""
    assert ailine.insert_rows_should_have_been_add_row("みかんの下に梨を追加して。売上は600", {})
    assert ailine.insert_rows_should_have_been_add_row("3行目に1行挿入して", {}) is None
