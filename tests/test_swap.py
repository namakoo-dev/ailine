# 入れ替え（SWAP）── 2026-08-27。Namakoo「行や列の入れ替えを頼む。付随する売上や原価も
# 同様に動かなければならない」。
#
# ★★ 設計は実測が決めた（bench/swap_formula_spike_RESULTS.md）:
#   値を文字として交換すると **式が壊れる** ── みかんの行の `=B3*C3` がりんごの金額を
#   出すようになる。見た目は正しく並んでいるので人は気づけない。この repo が最も嫌う
#   「静かに壊れる」形。だから moveRange（LibreOffice が参照を付け替える）で実装し、
#   事後条件は**式の文字ではなく計算結果**で突き合わせる。
#
# 契約:
#   ① 行か列かは **機械が実表を見て**決める（LLM に軸を当てさせない）
#   ② 両方に当たる／どちらにも当たらない時は**決めない**（断る）
#   ③ 依頼文が「行を」「列を」とはっきり書いていれば、その語で曖昧さを解く
#   ④ 生成する Basic には**名前で**渡す（Python が数えた番号を渡さない＝独立な 2 実装）
#   ⑤ 事後条件は「入れ替わった」だけでなく「中身が自分の値のまま移った」を証明する
#   ⑥ 入れ替えで計算結果が変わったら **fail**（挿入と違い、変わる正当な理由が無い）

import sys
from pathlib import Path

import openpyxl
import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
import ailine  # noqa: E402

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


def _meta(path, headers=("商品", "売上", "原価")):
    return {"sheets": ["売上"], "headers": {"売上": list(headers)},
            "header_rows": {"売上": 1}, "path": str(path)}


# --- ①③ 軸の決定 -------------------------------------------------------------------

def test_two_row_names_resolve_as_rows(tmp_path):
    ok, resolved, _i, err = ailine.verify_dsl_args(
        "SWAP", {"a": "みかん", "b": "ぶどう"}, _meta(_book(tmp_path)),
        task="みかんとぶどうを入れ替えて")
    assert ok, err
    assert resolved["_axis"] == "row"
    assert (resolved["_a_pos"], resolved["_b_pos"]) == (3, 4)


def test_two_column_names_resolve_as_columns(tmp_path):
    ok, resolved, _i, err = ailine.verify_dsl_args(
        "SWAP", {"a": "売上", "b": "原価"}, _meta(_book(tmp_path)),
        task="売上と原価を入れ替えて")
    assert ok, err
    assert resolved["_axis"] == "column"
    assert (resolved["_a_pos"], resolved["_b_pos"]) == (2, 3)


def test_a_name_that_is_both_a_column_and_a_row_is_refused(tmp_path):
    """② 両方に当たったら決めない ── 推測して別のものを動かすのが一番こわい。"""
    rows = [["商品", "売上", "原価"], ["売上", 1, 2], ["原価", 3, 4]]
    ok, _r, _i, err = ailine.verify_dsl_args(
        "SWAP", {"a": "売上", "b": "原価"}, _meta(_book(tmp_path, rows)),
        task="売上と原価を入れ替えて")
    assert not ok
    assert "両方あります" in err, err


def test_an_explicit_axis_word_breaks_the_tie(tmp_path):
    """③ 依頼文が『列を』と書いていれば、それで解ける（人が言ったことは使う）。"""
    rows = [["商品", "売上", "原価"], ["売上", 1, 2], ["原価", 3, 4]]
    ok, resolved, _i, err = ailine.verify_dsl_args(
        "SWAP", {"a": "売上", "b": "原価"}, _meta(_book(tmp_path, rows)),
        task="売上と原価の列を入れ替えて")
    assert ok, err
    assert resolved["_axis"] == "column"


@pytest.mark.parametrize("args,task,expect", [
    ({"a": "みかん", "b": "すいか"}, "みかんとすいかを入れ替えて", "決められません"),
    ({"a": "みかん", "b": "みかん"}, "みかんとみかんを入れ替えて", "同じもの"),
    ({"a": "", "b": "ぶどう"}, "入れ替えて", "取り出せませんでした"),
    ({"a": "みかん", "b": "ぶどう"}, "みかんとぶどうの列を入れ替えて", "列として決められません"),
])
def test_refusals(tmp_path, args, task, expect):
    ok, _r, _i, err = ailine.verify_dsl_args("SWAP", dict(args), _meta(_book(tmp_path)), task=task)
    assert not ok, f"{args} を通した"
    assert expect in err, err


def test_the_header_row_is_never_swapped(tmp_path):
    """見出しを巻き込む入れ替えは受け付けない（表の骨格を壊す）。"""
    rows = [["商品", "売上", "原価"], ["みかん", 800, 300]]
    ok, _r, _i, err = ailine.verify_dsl_args(
        "SWAP", {"a": "商品", "b": "みかん"}, _meta(_book(tmp_path, rows)),
        task="商品とみかんの行を入れ替えて")
    assert not ok
    assert "行として決められません" in err or "見出し" in err, err


# --- 依頼文の見分け -------------------------------------------------------------------

@pytest.mark.parametrize("task,asks,hint", [
    ("みかんとぶどうを入れ替えて", True, None),
    ("みかんとぶどうの行を入れ替えて", True, "row"),
    ("原価と売上の列を入れ替えて", True, "column"),
    ("売上と原価を交換して", True, None),
    ("みかんとぶどうの順番を逆にして", True, None),
    ("金額で降順に並べ替えて", False, None),
])
def test_the_request_is_recognised(task, asks, hint):
    assert ailine.task_asks_for_a_swap(task) is asks
    assert ailine._swap_axis_hint(task) == hint


# --- ④ 生成 --------------------------------------------------------------------------

def test_codegen_passes_names_not_numbers(tmp_path):
    """④ Basic には**名前**を渡す ── Basic 自身が実文書を走査して位置を見つける。
       Python が数えた番号を渡すと、突き合わせる相手が自分になる（恒真）。"""
    ok, resolved, _i, err = ailine.verify_dsl_args(
        "SWAP", {"a": "みかん", "b": "ぶどう"}, _meta(_book(tmp_path)),
        task="みかんとぶどうを入れ替えて")
    assert ok, err
    code = ailine.codegen_dsl("SWAP", resolved, _meta(_book(tmp_path)))
    assert 'Call SwapRowsByName(oDoc, "みかん", "ぶどう", 0, 0)' in code, code
    assert "3" not in code.split("SwapRowsByName")[1].split(")")[0], "行番号を渡している"


def test_codegen_column_variant(tmp_path):
    ok, resolved, _i, err = ailine.verify_dsl_args(
        "SWAP", {"a": "売上", "b": "原価"}, _meta(_book(tmp_path)),
        task="売上と原価の列を入れ替えて")
    assert ok, err
    assert 'Call SwapColumnsByName(oDoc, "売上", "原価", 0)' in ailine.codegen_dsl("SWAP", resolved, _meta(_book(tmp_path)))


def test_the_helpers_exist_in_the_bas(tmp_path):
    """★ 在っても鳴らない対策: 生成した Call の相手が実体として在ること。"""
    bas = (REPO / "src" / "ailine" / "helpers" / "AiLineHelpers.bas").read_text(encoding="utf-8")
    for name in ("Sub SwapRowsByName(", "Sub SwapColumnsByName(", "Function FindColByNameAt("):
        assert name in bas, f"{name} が helpers に無い"
    assert "moveRange" in bas


# --- ⑤⑥ 事後条件 ---------------------------------------------------------------------

_F = [["商品", "売上", "原価", "利益"],
       ["りんご", 1200, 700, "=B2-C2"],
       ["みかん", 800, 300, "=B3-C3"],
       ["ぶどう", 1500, 900, "=B4-C4"]]


def _args(axis="row", a="みかん", b="ぶどう", ap=3, bp=4):
    return {"a": a, "b": b, "_axis": axis, "_a_pos": ap, "_b_pos": bp}


def test_a_correct_row_swap_passes(tmp_path):
    before = _book(tmp_path, _F, name="before.xlsx")
    after = _book(tmp_path, [["商品", "売上", "原価", "利益"],
                              ["りんご", 1200, 700, "=B2-C2"],
                              ["ぶどう", 1500, 900, "=B3-C3"],     # 式は自分の行を指す
                              ["みかん", 800, 300, "=B4-C4"]])
    status, reason = ailine.check_swap(after, _args(), source_book=before)
    assert status == "pass", reason
    assert "自分の値のまま" in reason


def test_a_swap_that_did_not_happen_fails(tmp_path):
    before = _book(tmp_path, _F, name="before.xlsx")
    after = _book(tmp_path, list(_F))          # 何も動いていない
    status, reason = ailine.check_swap(after, _args(), source_book=before)
    assert status == "fail", reason


def test_a_value_swap_that_breaks_formulas_fails(tmp_path):
    """★★ これが実測した壊れ方そのもの（スパイク Q1）: 名前と数値だけを交換して
       式の文字を置いていくと、各行の計算結果が**他の行の値**になる。
       並びは正しく見えるので、人の目では気づけない。"""
    before = _book(tmp_path, [["商品", "売上", "原価", "利益"],
                               ["りんご", 1200, 700, 500],
                               ["みかん", 800, 300, 500],
                               ["ぶどう", 1500, 900, 600]], name="before.xlsx")
    after = _book(tmp_path, [["商品", "売上", "原価", "利益"],
                              ["りんご", 1200, 700, 500],
                              ["ぶどう", 1500, 900, 500],     # ← 利益がみかんのまま残った
                              ["みかん", 800, 300, 600]])
    status, reason = ailine.check_swap(after, _args(), source_book=before)
    assert status == "fail", f"値が入れ替わっていないのに通した: {reason}"


def test_a_swap_that_disturbs_another_row_fails(tmp_path):
    """★ 恒真殺し: 入れ替えた 2 行以外が 1 セルでも変わったら落ちる。"""
    before = _book(tmp_path, ROWS, name="before.xlsx")
    after = _book(tmp_path, [["商品", "売上", "原価"],
                              ["りんご", 9999, 700],          # ← 巻き込んだ
                              ["ぶどう", 1500, 900],
                              ["みかん", 800, 300]])
    status, reason = ailine.check_swap(after, _args(), source_book=before)
    assert status == "fail", reason


def test_a_correct_column_swap_passes(tmp_path):
    before = _book(tmp_path, ROWS, name="before.xlsx")
    after = _book(tmp_path, [["商品", "原価", "売上"],
                              ["りんご", 700, 1200],
                              ["みかん", 300, 800],
                              ["ぶどう", 900, 1500]])
    status, reason = ailine.check_swap(after, _args(axis="column", a="売上", b="原価"),
                                        source_book=before)
    assert status == "pass", reason


def test_a_column_swap_that_moved_only_the_headers_fails(tmp_path):
    """★ 見出しだけ入れ替えて中身を置いていく ── 数字の意味が入れ替わる最悪の形。"""
    before = _book(tmp_path, ROWS, name="before.xlsx")
    after = _book(tmp_path, [["商品", "原価", "売上"],
                              ["りんご", 1200, 700],
                              ["みかん", 800, 300],
                              ["ぶどう", 1500, 900]])
    status, reason = ailine.check_swap(after, _args(axis="column", a="売上", b="原価"),
                                        source_book=before)
    assert status == "fail", f"見出しだけの入れ替えを通した: {reason}"


def test_without_the_before_file_it_does_not_claim(tmp_path):
    """適用前が無ければ断定しない（warn ＝ ✓ は出ない）。"""
    after = _book(tmp_path, ROWS)
    status, _reason = ailine.check_swap(after, _args(), source_book=None)
    assert status == "warn"


# --- 入れ替えの門を、op 名でなく証拠で作る（2026-08-29・Namakoo が実測）---------------

def test_the_swap_gate_is_not_a_list_of_op_names():
    """★★ 実測: 「税込み金額列と金額列を入れ替えて」が **COMPUTE_COLUMN**（計算列）に
       読まれ、金額列を掛け算で潰しかけた（関所が止めた）。門が
       ("CLARIFY","FREEFORM","OUT_OF_VOCAB","SORT") という **op 名の列挙**だったので、
       それ以外を返した回は素通りしていた ── 今日 4 度目の同じ形。"""
    src = (REPO / "src" / "ailine" / "__init__.py").read_text(encoding="utf-8")
    i = src.index("task_asks_for_a_swap(a.task)")
    seg = src[max(0, i - 600):i + 900]
    assert '("CLARIFY", "FREEFORM", "OUT_OF_VOCAB", "SORT")' not in seg, "op 名の列挙が残っている"
    assert 'get("op") == "SWAP"' in seg, "既に入れ替えで読めている回を除いていない"


def test_both_targets_must_be_written_in_the_request():
    """★★ 門を広げた途端、元の狭い門が守っていた物が壊れた:
       「税込み金額の**順番を逆にして**」で、第二段が**相手をでっち上げて**
       入れ替えに化けた（正当な並べ替えを壊す）。
    ★ A' 原則をここにも通す ── 入れ替える 2 つは、どちらも依頼文に在ること。
      片方しか書かれていない依頼は、入れ替えの依頼ではない。"""
    src = (REPO / "src" / "ailine" / "__init__.py").read_text(encoding="utf-8")
    i = src.index("_sa, _sb = str(_sw_args.get(")
    seg = src[i:i + 900]
    assert '_sa in (a.task or "")' in seg and '_sb in (a.task or "")' in seg, seg[:400]


def test_a_nested_column_name_still_resolves(tmp_path):
    """★ 片方の名前がもう片方を含む（税込み金額 ⊃ 金額）── 部分文字列の穴の再演を止める。"""
    p = tmp_path / "n.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "請求"
    ws.append(["取引先", "金額", "税込み金額"])
    ws.append(["丸和物流", 100, 110])
    wb.save(p)
    meta = {"sheets": ["請求"], "headers": {"請求": ["取引先", "金額", "税込み金額"]},
             "header_rows": {"請求": 1}, "path": str(p)}
    assert ailine._swap_pair_resolves(meta, "請求", "税込み金額", "金額")
    ok, r, _i, err = ailine.verify_dsl_args(
        "SWAP", {"a": "税込み金額", "b": "金額"}, meta, task="税込み金額列と金額列を入れ替えて")
    assert ok, err
    assert r.get("_axis") == "column", r          # 軸は機械が決める（見出しで一致）
    assert (r["_a_pos"], r["_b_pos"]) == (3, 2), r
