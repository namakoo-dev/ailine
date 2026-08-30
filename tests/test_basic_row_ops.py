# 表の基本操作は、どの表でも・どの言い方でも ── 2026-08-28。
# Namakoo「正直この基本操作があらゆる表に対して絶対にできないと使い物にはならない」
#
# ★★ 実測した 3 つの壊れ（全部**言い回しの列挙**が原因だった）:
#   ① 「丸和物流と近江スチールの間に北斗精機を**作って**」→ 空行が挿さった。
#      読み直しの条件が動詞の列挙（追加/足し/入れ）で、「作って」が漏れていた。
#   ② 同じ依頼で回ごとに INSERT_ROWS / CLARIFY / EXTRACT に化けた。
#      読み直しの門が **op 名の列挙**だったので、EXTRACT の回は素通りした。
#   ③ 「ナット**を**削除して」→『1行目は見出し行です』。位置解決が「〜の行」という
#      言い回しを要求していた。
#
# ★ 直しの線: **列挙をやめて、実表に訊く**。
#   ・位置: 相対の言い回しが当たらなければ、依頼文に literal で現れる**実在の値**が
#     ちょうど 1 行にしか無いかを表に訊く（2 行に当たったら決めない）
#   ・門: op 名でなく**宣言**（OP_WRITE_TARGET）で作る
#   ・値: LLM に作らせず、依頼文と実表から機械が決める（A' 原則）
#
# ★★ この試験が守る核心: **表の中身が変わっても同じように効くこと**。
#   だから検体は「請求書の表」ではなく、語彙も列名も違う表を複数使う。

import sys
from pathlib import Path

import openpyxl
import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
import ailine  # noqa: E402

# ★ 3 つとも別世界の表（在庫・名簿・献立）。同じ規則が効くことを見る。
TABLES = {
    "在庫": (["品名", "棚", "数量"],
              [["ボルト", "A-1", 120], ["ナット", "A-2", 80], ["ワッシャー", "B-1", 300]]),
    "名簿": (["氏名", "所属", "内線"],
              [["山田", "営業", 101], ["鈴木", "経理", 202], ["高橋", "総務", 303]]),
    "献立": (["料理", "主材料", "分量"],
              [["カレー", "牛肉", 4], ["味噌汁", "豆腐", 2], ["サラダ", "レタス", 3]]),
}


def _book(tmp_path, key, name=None):
    headers, rows = TABLES[key]
    p = tmp_path / (name or f"{key}.xlsx")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = key
    ws.append(headers)
    for r in rows:
        ws.append(r)
    wb.save(p)
    return p


def _meta(path, key):
    headers, _rows = TABLES[key]
    return {"sheets": [key], "headers": {key: list(headers)},
            "header_rows": {key: 1}, "path": str(path)}


# --- ① 位置は「実表に訊く」（言い回しの列挙ではない）------------------------------------

@pytest.mark.parametrize("key,task,want_row,want_name", [
    ("在庫", "ボルトとナットの間にスプリングを作って", 3, "ボルト"),
    ("在庫", "ナットの下にピンを追加して", 4, "ナット"),
    ("在庫", "ワッシャーの上にリベットを入れて", 4, "ワッシャー"),
    ("名簿", "山田と鈴木の間に佐藤を作って", 3, "山田"),
    ("名簿", "鈴木を削除して", 3, "鈴木"),          # ★「〜の行」と言わなくても解ける
    ("献立", "味噌汁を消して", 3, "味噌汁"),
    ("献立", "カレーと味噌汁の間にコロッケを入れて", 3, "カレー"),
])
def test_the_position_is_resolved_from_the_real_table(tmp_path, key, task, want_row, want_name):
    p = _book(tmp_path, key)
    at, note = ailine.resolve_row_anchor(task, _meta(p, key), key)
    assert at == want_row, (at, note)
    assert want_name in (note or ""), note


def test_a_condition_is_not_mistaken_for_a_row_name(tmp_path):
    """★ 恒真殺し: 「数量が100未満の行」を行の名前として当てにいかない
       （当てにいくと、たまたま似た値の行を消す）。"""
    p = _book(tmp_path, "在庫")
    at, note = ailine.resolve_row_anchor("数量が100未満の行を削除して", _meta(p, "在庫"), "在庫")
    assert at is None and note and "見つかりません" in note, (at, note)


def test_a_name_on_two_rows_is_never_guessed(tmp_path):
    """★ 2 行に当たったら決めない ── 推測で別の行を消すのが一番こわい。"""
    p = tmp_path / "dup.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "在庫"
    ws.append(["品名", "棚", "数量"])
    for r in [["ボルト", "A-1", 120], ["ナット", "A-2", 80], ["ナット", "B-9", 5]]:
        ws.append(r)
    wb.save(p)
    at, _note = ailine.resolve_row_anchor("ナットを削除して", _meta(p, "在庫"), "在庫")
    assert at is None, at


def test_a_header_word_is_not_a_row_name(tmp_path):
    """★ 列名を行の名前と読み違えない（「棚を削除して」は列の話）。"""
    p = _book(tmp_path, "在庫")
    got = ailine._row_named_anywhere_in_task(
        "棚を削除して",
        {2: ["ボルト", "A-1", "120"], 3: ["ナット", "A-2", "80"]}, ["品名", "棚", "数量"])
    assert got is None, got


# --- ② 新しい行の値は、機械が依頼文と実表から決める（A' 原則）--------------------------

@pytest.mark.parametrize("key,task,want", [
    ("在庫", "ボルトとナットの間にスプリングを作って", {"品名": "スプリング"}),
    ("名簿", "山田と鈴木の間に佐藤を作って", {"氏名": "佐藤"}),
    ("献立", "カレーと味噌汁の間にコロッケを入れて", {"料理": "コロッケ"}),
])
def test_the_payload_lands_in_the_column_where_the_anchors_live(tmp_path, key, task, want):
    """★★ 実測: 第二段は『取引先=丸和物流（＝目印そのもの）／項目=北斗精機（別の列）』を
       返した。目印は位置であって値ではないし、置く物は目印と**同じ列**の住人。"""
    p = _book(tmp_path, key)
    # LLM が返しがちな形をそのまま渡す（目印・別列・でっち上げが混ざったもの）
    anchors = ailine.row_anchor_names(task)
    h0, h1, h2 = TABLES[key][0]
    # 目印を 1 列目に・置く物を**別の列**に・でっち上げを 3 列目に（実測どおりの形）
    llm = {h0: anchors[0], h1: list(want.values())[0], h2: "未定"}
    got = ailine.add_row_values_from_request(task, _meta(p, key), key, llm)
    assert got == want, got


def test_invented_values_are_dropped(tmp_path):
    """★ 依頼文に literal で無い値は入れない（未定・未設定・件は実測で出た）。"""
    p = _book(tmp_path, "在庫")
    got = ailine.add_row_values_from_request(
        "ボルトとナットの間にスプリングを作って", _meta(p, "在庫"), "在庫",
        {"品名": "スプリング", "棚": "未定", "数量": "未設定"})
    assert got == {"品名": "スプリング"}, got


def test_an_explicitly_named_column_keeps_its_value(tmp_path):
    """★ 人が列を名指しした値は、その列へ（「数量は50」）。"""
    p = _book(tmp_path, "在庫")
    got = ailine.add_row_values_from_request(
        "ボルトとナットの間にスプリングを作って。数量は50", _meta(p, "在庫"), "在庫",
        {"品名": "スプリング", "数量": 50})
    assert got == {"品名": "スプリング", "数量": 50}, got


def test_the_anchor_names_are_never_written_as_values(tmp_path):
    p = _book(tmp_path, "在庫")
    assert ailine.row_anchor_names("ボルトとナットの間にスプリングを作って") == ["ボルト", "ナット"]
    got = ailine.add_row_values_from_request(
        "ボルトとナットの間にスプリングを作って", _meta(p, "在庫"), "在庫",
        {"品名": "ボルト"})
    # ★ 目印（ボルト）は値にならない ── そこが この試験の芯で、変わっていない。
    # ★★ 2026-08-30: 篩が空になった後、**機械の引き算が値を取り戻す**ようになった
    #   （旧: {} ＝ 何も書かない）。目印でない・依頼文に在る・助詞を含まない、の 3 つを
    #   満たすものだけが残る。
    assert got == {"品名": "スプリング"}, got
    assert got.get("品名") not in ("ボルト", "ナット")


# --- ③ 読み直しの門は「宣言」で作る（op 名の列挙ではない）------------------------------

def test_the_gate_is_built_from_declarations_not_op_names():
    """★★ 実測: 同じ依頼文で INSERT_ROWS / CLARIFY / EXTRACT を返し分ける。
       op 名を数え上げても必ず漏れる ── 宣言で門を作る。"""
    src = (REPO / "src" / "ailine" / "__init__.py").read_text(encoding="utf-8")
    assert '_reread_ops = (' not in src, "op 名の列挙が残っている"
    i = src.index("def _already_places_a_row(st):")
    seg = src[i:i + 1600]
    # ★ 2026-08-29: 条件を「行をずらす**かつ**末尾に置く」から「**新しい行に中身を置く**」
    #   の 1 点に絞った ── 合計行のようにずらさずに置く op（APPEND_TOTAL）が素通りして
    #   いたため（Namakoo の通しで実測: 「件数の合計も合計行に入れて」が行追加に化けた）。
    assert "WRITE_NEW_ROW_AT_END" in seg, seg[:300]
    assert "WRITE_FORMAT_ONLY" in seg and "WRITE_REMOVE" in seg, seg[:900]


def test_the_verb_list_is_gone():
    """★ 『追加/足し/入れ』の列挙が戻ったら、また別の動詞で漏れる。"""
    src = (REPO / "src" / "ailine" / "__init__.py").read_text(encoding="utf-8")
    i = src.index("def insert_rows_should_have_been_add_row(")
    seg = src[i:i + 2600]
    assert '"追加" in text or "足し" in text' not in seg, "動詞の列挙が戻っている"


@pytest.mark.parametrize("op,expect", [
    ("ADD_ROW", True), ("INSERT_ROWS", False), ("EXTRACT", False),
    ("CLARIFY", False), ("FILL_COLOR", False), ("DELETE_ROWS", False),
])
def test_only_add_row_declares_that_it_places_a_row_with_values(op, expect):
    """★ 門の土台。ここがずれると、読み直しが要る回に読み直さない。"""
    got = (ailine._op_writes(op, ailine.WRITE_ROW_SHIFT)
            and ailine._op_writes(op, ailine.WRITE_NEW_ROW_AT_END))
    assert got is expect, op


def test_a_formatting_or_removing_plan_is_left_alone():
    """★ 黙りすぎない側の対: 見た目・削除・並べ替えは『別の仕事』── 横取りしない。"""
    for op in ("FILL_COLOR", "BOLD", "CENTER_ALIGN", "DELETE_ROWS", "DELETE_COLUMN", "SORT"):
        assert any(ailine._op_writes(op, k) for k in
                    (ailine.WRITE_FORMAT_ONLY, ailine.WRITE_REMOVE, ailine.WRITE_REORDER)), op


# --- ④ 途中への挿入で「末尾に足すはず」の助言が誤爆しない ------------------------------

def test_inserting_in_the_middle_does_not_trip_the_append_precondition():
    """★★ 2026-08-28 の朝、ADD_ROW に『末尾に足す』を宣言したせいで、**途中に挿した回**が
       「既存の行の値を 29 件書き換えました」で △ に落ちた（俺が同じ日に開けた穴）。
       ★ 番号で突き合わせる前提は、位置がずれる op では使えない。"""
    from ailine_core import write_precondition as wp
    assert "new_row_at_end" in wp.POSITION_BASED
    before = {"cells": {"S!2,1": ("A",), "S!3,1": ("B",)}}
    after = {"cells": {"S!2,1": ("A",), "S!3,1": ("X",), "S!4,1": ("B",)}}
    kw = dict(cell_ref=lambda r, c: f"R{r}C{c}", fmt_value=repr)
    # 宣言に row_shift があれば黙る（挿入で下がずれるのは意図どおり）
    assert wp.check_write_preconditions_detail(
        ("row_shift", "new_row_at_end"), before, after, **kw) is None
    # row_shift の宣言が無ければ、今も捕まる（黙りすぎていない）
    got = wp.check_write_preconditions_detail(("new_row_at_end",), before, after, **kw)
    assert got and got[0] == "new_row_at_end", got


# --- 1 回の依頼に、位置を作るのは 1 回だけ（2026-08-29・Namakoo の設計判断）------------

def test_two_placements_on_the_same_axis_are_refused():
    """★★ 実測: 「味噌汁の上に新品を入れて」で一段目が
         [INSERT_ROWS at:2（空行）, ADD_ROW at:2（値つき）]
       を返し、両方走って**行が 2 本**増えた。同じ仕事の二重宣言。
    ★ Namakoo の設計判断:「行や列を 2 つ以上増やす操作はもともと無いから縛っていい。
       複数必要なら順次増やせばいい」── これは**座標の法則の形と一致する**
       （1 つの操作 = 1 つの写像 π。写像は合成しない）。
    ★ 数えるのは op 名でなく**宣言**（新しい op が増えても勝手に数に入る）。"""
    assert ailine.placements_in_plan([{"op": "INSERT_ROWS"}, {"op": "ADD_ROW"}]) == \
        {"row": 2, "col": 0}
    assert ailine.too_many_placements([{"op": "INSERT_ROWS"}, {"op": "ADD_ROW"}])
    assert ailine.too_many_placements([{"op": "ADD_COLUMN"}, {"op": "ADD_COLUMN"}])


def test_a_genuine_compound_is_not_refused():
    """★ 黙りすぎない側の対: 「足してから並べ替えて」のような**別種の合成**は縛らない。
       縛るのは『同じ軸に 2 回 place する』形だけ。"""
    assert ailine.too_many_placements([{"op": "ADD_ROW"}, {"op": "SORT"}]) is None
    assert ailine.too_many_placements([{"op": "ADD_ROW"}]) is None
    assert ailine.too_many_placements([{"op": "ADD_ROW"}, {"op": "ADD_COLUMN"}]) is None


def test_the_gate_runs_before_anything_is_applied():
    """★ 当て物でなく**関所**であること（畳めなかった回に、壊す前に止まる）。"""
    src = (REPO / "src" / "ailine" / "__init__.py").read_text(encoding="utf-8")
    i = src.index("if (_dup := too_many_placements(plan)):")
    j = src.index("if len(plan) == 1:", i)
    assert j > i and "return 3" in src[i:j], src[i:i + 300]


# --- 値に助詞は入らない（文法の線で弾く）-----------------------------------------------

@pytest.mark.parametrize("task,anchors,want", [
    ("味噌汁の上に新品を入れて", ["味噌汁"], "新品"),
    ("丸山重工の右にPCパーツ", ["丸山重工"], "PCパーツ"),
    # ★★ 2026-08-30 に**取れるようになった**（旧: 語尾が落ちきらず None ＝ 決めない）。
    #   「を追加して」は語尾の一覧に在り「を作って」は無い ── 列挙の穴そのものだった。
    #   語ではなく**形**（「を」＋短い語＋て/た/る…）で落とすようにしたので、
    #   「つくって」「作る」など未知の言い方にも効く。
    ("ボルトとナットの間にスプリングを作って", ["ボルト", "ナット"], "スプリング"),
    # ★ 黙りすぎていないこと: 行そのものを足す依頼は値ではない（『足』を値にしない）
    ("ボルトとナットの間に1行足して", ["ボルト", "ナット"], None),
])
def test_a_value_never_contains_a_particle(task, anchors, want):
    """★★ 2026-08-29（3 度目の『列挙は漏れる』）: 動詞の語尾を数え上げても漏れる。
       ★ 文法の線で弾く ── セルに書く値の中に助詞は入らない。
       ★★ 2026-08-30: 落とす側も**語の一覧をやめて形にした**。
         それまでは『スプリングを作って』を丸ごと値にしないために **決めない**
         側へ倒していたが、1B の検体 6 件・7B でも同じ形で落ちていた
         （『AとBの間にXを作って』が空行の挿入になる）。"""
    assert ailine.bare_value_from_task(task, anchors, "品名", ["品名", "棚", "数量"]) == want


def test_the_longer_of_two_task_derived_values_wins(tmp_path):
    """★ 実測: 第二段が『新』を返し、篩が『依頼文に在る』だけを見て通した
       ── 短い部分文字列は必ず通る。両方とも依頼文由来なら**長い方**を採る。"""
    p = _book(tmp_path, "献立")
    got = ailine.add_row_values_from_request(
        "味噌汁の上に新品を入れて", _meta(p, "献立"), "献立", {"料理": "新"})
    assert got == {"料理": "新品"}, got


# --- 「除く」の 2 つの読み（2026-08-29・84 件の効果検体で最後まで残った穴）--------------

@pytest.mark.parametrize("task,want_row", [
    ("ナットの行を除いて", 3), ("ナットの行を削除して", 3), ("ナットを消して", 3),
    ("ナットの行はいらない", 3),
])
def test_a_removal_reading_resolves_the_row(tmp_path, task, want_row):
    """★ 一段目は 3 表で EXTRACT / OUT_OF_VOCAB / 条件付き抽出 と返り分けた。
       ★ ここは**語の列挙で正しい** ── 判定しているのが「表のどこか」ではなく
         「人がどの動作を言ったか」だから（動作は言葉でしか分からない）。
       ★ 漏れた時の壊れ方が違うのが肝: 語が無ければ**何も起きない**（今までどおり）。
         黙って別のことはしないし、読みは書く前に画面に出る。"""
    p = _book(tmp_path, "在庫")
    got = ailine.removal_reading(task, _meta(p, "在庫"), "在庫")
    assert got and got[0] == want_row, got


@pytest.mark.parametrize("task", [
    "ナット以外を抜き出して", "ナットを除いた行を抜き出して",
])
def test_the_except_reading_is_refused_not_flipped(tmp_path, task):
    """★★ 実測（自分で開けた片配線）: 「味噌汁**以外**を抜き出して」が
       味噌汁**だけ**を抜き出して ✓ になった ── **逆のことをして合格**。
    ★ 『以外』（残す側を選ぶ）は cmp『〜でない』が要るが、述語は 3 箇所が独立に持つ
      凍結表なので、いまは無い。**無いものは名指しで断る** ──
      ここで削除に化けさせると、残したい行を消す取り返しのつかない事故になる。"""
    p = _book(tmp_path, "在庫")
    assert ailine.removal_reading(task, _meta(p, "在庫"), "在庫") is None
    assert ailine.unsupported_except_reading(task), task


def test_a_plain_extract_is_not_refused():
    """★ 黙りすぎない側の対: 普通の抽出は通す。"""
    assert ailine.unsupported_except_reading("数量が100以上の行を抜き出して") is None
    assert ailine.unsupported_except_reading("ナットの行だけ抜き出して") is None


def test_the_refusal_sits_outside_the_reread_gate():
    """★★ 片配線の再演を止める: 断りを読み直しの門の**内側**に置いていたので、
       先に別の読み直しが印を立てた回に素通りした。断りは門の外に置く。"""
    src = (REPO / "src" / "ailine" / "__init__.py").read_text(encoding="utf-8")
    i = src.index("if (_exc := unsupported_except_reading(a.task)):")
    j = src.index("_reread_done = bool(getattr(a,")
    k = src.index("def _already_places_a_row(st):")
    assert j < i < k, "『以外』の断りが、読み直しの層の中に入っている"


def test_the_anchor_is_never_written_even_when_the_llm_offers_nothing_else(tmp_path):
    """★★ 2026-08-29（84 件の効果検体で最後に残った 1 件・また片配線）:
       「鈴木**の上に**新品を入れて」で **氏名=鈴木**（＝位置の目印そのもの）が
       新しい行に書かれた。
    ★ 2 段構えで壊れていた:
       ① 値の篩を「読み直した経路」にだけ入れていた（一段目が最初から ADD_ROW を
          返した回は素通り）── 処方は「両方に入れる」でなく**必ず同じ関数を通す**
       ② 第二段が目印だけを返した回、篩で全部落ちて空になり、呼び出し側が
          「置き換え無し」と見て**悪い値のまま**通した ── 篩が空の回にも機械の
          引き算を使う（LLM が何も出さない回と同じ扱い）
    """
    p = _book(tmp_path, "名簿")
    got = ailine.add_row_values_from_request(
        "鈴木の上に新品を入れて", _meta(p, "名簿"), "名簿", {"氏名": "鈴木"})
    assert got == {"氏名": "新品"}, got
    # 値をまったく返さなかった回も同じ結果になること
    assert ailine.add_row_values_from_request(
        "鈴木の上に新品を入れて", _meta(p, "名簿"), "名簿", {}) == {"氏名": "新品"}


def test_the_sieve_is_on_the_path_every_add_row_takes():
    """★ 経路が増えても篩が外れない形になっていること（片配線の再演を止める）。"""
    src = (REPO / "src" / "ailine" / "__init__.py").read_text(encoding="utf-8")
    i = src.index('if (_st or {}).get("op") != "ADD_ROW":')
    seg = src[max(0, i - 700):i + 500]
    assert "for _st in plan:" in seg, seg[:300]
    assert "add_row_values_from_request(" in seg, seg[-300:]


def test_a_total_row_request_is_not_stolen_by_the_row_placement_reread():
    """★★ 2026-08-29（Namakoo の通しで実測）: 「件数の合計も合計行に入れて」が
       **行追加**に化けた。一段目は 3/3 とも正しく APPEND_TOTAL を返していたのに、
       読み直しの門が「行を**ずらす**」op だけを『もう置けている』と数えていて、
       合計行のように**ずらさずに末尾へ置く** op が素通りしていた。
    ★ 見るべきは「新しい行に中身を置く」と宣言しているか、の 1 点だけ。"""
    for op in ("ADD_ROW", "APPEND_TOTAL"):
        assert ailine._op_writes(op, ailine.WRITE_NEW_ROW_AT_END), op
    for op in ("INSERT_ROWS", "EXTRACT", "FILL_COLOR"):
        assert not ailine._op_writes(op, ailine.WRITE_NEW_ROW_AT_END), op
