# -*- coding: utf-8 -*-
"""「機械が依頼文から取り直す引数」の名簿を、**届く op 全部**について決める（2026-09-17）。

★ なぜ在るか（仕分け②）: 配線盤に「★ 未調査」で出ていた列。EXTRACT を入れた理由
  （2026-08-30 の実測事故）は書かれていたが、**他の op を入れていない理由が 1 行も無い**。
  理由が無い名簿は「調べた結果こうなっている」のか「まだ見ていない」のか読めない。

★★ 分母は 30 op ではない。この名簿を読むのは fold_identical_steps ただ 1 箇所で、
  そこを通るのは **WRITE_NEW_SHEET を書くと宣言した段だけ** ── 7 op。
  分母を 30 と書くと穴が 29 に見え、**本当の穴 2 件がその中に埋もれる**。
  だから分母は宣言から導き、7 op すべてに理由を書かせる。

★★ 線（2026-09-17 に実測して引いた）: 機械の取り直しは**段ごとでなく依頼文ごと**。
  compare_words.read(task) も EXTRACT_COLUMNS の _asked も task 全体を読むので、
  同じ op の 2 段は解決後に**必ず同じ値**になる。
  ⇒ 取り直す引数は比べる前に外すのが正しく、**外し漏れだけ**が事故になる。
    外しすぎは起きない ── 本当に別の仕事なら、取り直さない引数（col など）が違う。

★ 実測で見つけた穴（どちらも「畳めるはずの 2 段が畳まれない」＝ 2026-08-30 の再演）:
    ① EXTRACT の cmp ── compare_words が LLM の比較に勝つのに名簿に無かった
    ② EXTRACT_COLUMNS の cols ── LLM の答えが一切効かないのに列ごと無かった
  さらに名簿に在った "values" は **EXTRACT の引数ですらなかった**（ADD_ROW の引数）。
  在るだけで「値まわりは押さえてある」と読ませる、いちばん静かな嘘。
"""
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
import ailine  # noqa: E402


def reachable_ops() -> set:
    """名簿が**効く**op ── fold_identical_steps がこの条件で絞っている。

    ★ ここは手書きしない。宣言（OP_WRITE_TARGET）から導く ── op が増えた日に
      名簿を見に来させるための分母。
    """
    return {op for op in ailine.OP_SCHEMA
            if ailine._op_writes(op, ailine.WRITE_NEW_SHEET)}


#: 届く op 1 つずつの**決め**と、その理由。★ 「入れない」も決めであって、空白ではない。
DECIDED = {
    "EXTRACT": (
        ("value", "cmp"),
        "value は依頼文が名指しする実在の値（task_names_real_values）で機械が勝つ。"
        "cmp も compare_words.read(task) が非 None なら LLM の比較を捨てて勝つ ── "
        "どちらも task 全体から取るので、2 段は解決後に必ず同じになる"),
    "EXTRACT_COLUMNS": (
        ("cols",),
        "残す列は**依頼文に現れる実在の列名**を出現順に機械が拾う（_asked）。"
        "LLM の cols は実在照合を通った候補としてしか使われず、依頼文が列名を"
        "含む限り LLM の答えは結果に効かない（2026-09-17 に 4 通りで実測）"),
    "DEDUP": (
        (),
        "keys は resolve_col_ref が**LLM の書いた参照**を正規化するだけで、依頼文から"
        "取り直してはいない ── LLM の答えが結果を決めるので外せない。"
        "★ 残る穴: 『取引先』と『A』のような別表記は解決後に同じ列になるが畳まれない。"
        "これは A' の取り直しでなく**参照の正規化**で、畳む側は headers を持っていない"),
    "AGGREGATE": (
        (),
        "group_col / value_col は検算関数が resolved へ書き戻さない ── LLM の値がそのまま"
        "使われる（2026-09-17 に AST で全代入を数えて確認）"),
    "PIVOT": (
        (),
        "AGGREGATE と同じ ── 検算関数は resolved の公開引数を 1 つも書き換えない"),
    "FORMAT_MAP": (
        (),
        "template_sheet は書き換えない。検算が積むのは _ 始まりの導出値だけで、"
        "それらは**畳んだ後の検算**で付くので比較には現れない"),
    "REPORT_PER_ROW": (
        (),
        "template_sheet / name_col は書き換えない。FORMAT_MAP と同じ形"),
}


def test_the_denominator_comes_from_the_declaration():
    """★ 分母は宣言から導く ── 新しく「シートを作る」op が増えたら、ここで決めさせる。"""
    reach = reachable_ops()
    assert reach, "★ 分母が空（下の検査が全部素通りする）"
    assert reach == set(DECIDED), (
        f"届く op と決めた op が食い違う: 決めていない {sorted(reach - set(DECIDED))} / "
        f"届かないのに決めている {sorted(set(DECIDED) - reach)}")


def test_every_reachable_op_has_a_written_reason():
    """★ 「入れない」にも理由を書かせる ── 空白は『まだ見ていない』と読めてしまう。"""
    assert DECIDED, "★ 台帳が空"
    for op, (_keys, why) in DECIDED.items():
        assert str(why).strip(), f"{op} の理由が空"


def test_the_roster_matches_what_was_decided():
    """★ 台帳と製品の名簿が一致すること（片方だけ直さない）。"""
    want = {op: tuple(keys) for op, (keys, _w) in DECIDED.items() if keys}
    got = {op: tuple(keys) for op, keys in ailine.MACHINE_DERIVED_ARGS.items()}
    assert got == want, f"名簿 {got} ≠ 台帳 {want}"


def test_no_key_in_the_roster_is_a_phantom():
    """★★ 名簿の鍵は、その op が**本当に持つ引数**であること。

    2026-09-17 まで EXTRACT に "values" が入っていたが、それは ADD_ROW の引数で、
    EXTRACT のスキーマには無い。落としても何も起きない鍵が、名簿を厚く見せていた。
    """
    assert ailine.MACHINE_DERIVED_ARGS, "★ 名簿が空"
    for op, keys in ailine.MACHINE_DERIVED_ARGS.items():
        schema = set(ailine.OP_SCHEMA[op])
        assert keys, f"{op}: 空の組を名簿に置かない（入れないなら行ごと無い）"
        for k in keys:
            assert k in schema, (
                f"{op} に引数『{k}』は無い（ある: {sorted(schema)}）── "
                "効かない鍵は『押さえてある』という嘘になる")


def test_the_roster_only_lists_ops_it_can_reach():
    """★ 届かない op を名簿に置かない（読む側が『効いている』と誤読する）。"""
    reach = reachable_ops()
    assert reach
    stray = sorted(set(ailine.MACHINE_DERIVED_ARGS) - reach)
    assert not stray, f"fold_identical_steps が見ない op が名簿に在る: {stray}"


# --- 挙動の側から縛る（宣言どうしの突き合わせは恒真になる） -------------------------------

def test_the_llm_answer_does_not_reach_the_result_for_extract_columns():
    """★★ 別実装でなく**実物の検算関数**に 3 通り食わせて、結果が動かないことを見る。

    ★ 名簿は「機械が取り直す」と主張している。その主張を、名簿を読まない側
      （検算関数）から確かめる ── 同じ辞書を読み合うと恒真になる。
    """
    task = "取引先と金額だけ残して"
    headers = {"データ": ["取引先", "金額", "担当", "日付"]}
    book_meta = {"header_rows": {"データ": 1}, "sheets": ["データ"]}
    seen = set()
    for llm in (["取引先"], ["金額"], ["担当"], ["取引先", "金額"]):
        resolved = {"cols": list(llm)}
        out = ailine._verify_extract_columns(
            resolved, set(), "データ", book_meta, task, headers)
        assert out is None, f"断られた: {out}"
        seen.add(tuple(resolved["cols"]))
    assert seen == {("取引先", "金額")}, (
        f"LLM の答えが結果に効いている: {sorted(seen)} ── "
        "効くなら cols は『機械が取り直す引数』ではない（名簿から外すこと）")
    assert "cols" in ailine.MACHINE_DERIVED_ARGS["EXTRACT_COLUMNS"]


def test_two_column_extractions_from_one_request_are_one_job():
    """★ 畳めること ── 畳み損なうと 2 段目が連鎖の規則で 1 段目の出力を食う
       （EXTRACT_COLUMNS では『全部の列が指定されています』という的外れな断りになる）。"""
    plan = [{"op": "EXTRACT_COLUMNS", "args": {"cols": ["取引先"]}},
            {"op": "EXTRACT_COLUMNS", "args": {"cols": ["金額"]}}]
    folded, dropped = ailine.fold_identical_steps(plan)
    assert dropped == 1 and len(folded) == 1, folded


def test_two_extracts_that_differ_only_in_the_comparison_are_one_job():
    """★ cmp も機械が勝つ ── LLM が段ごとに違う比較を書いても、解決後は同じになる。"""
    plan = [{"op": "EXTRACT", "args": {"col": "取引先", "cmp": "eq", "value": "丸和物流"}},
            {"op": "EXTRACT", "args": {"col": "取引先", "cmp": "contains", "value": "みどり建設"}}]
    folded, dropped = ailine.fold_identical_steps(plan)
    assert dropped == 1 and len(folded) == 1, folded


def test_two_comparisons_in_one_request_are_refused_not_folded():
    """★★ 畳みすぎの心配への答え（実測）。

    「3000以上の行と3000未満の行」のように**比較が 2 つ**ある依頼は、辞書が
    ambiguous を返して**断られる** ── 段が畳まれても畳まれなくても実行されない。
    つまり cmp を外したことで「別の条件が黙って消える」形は作れない。
    """
    from ailine_core import compare_words
    r = compare_words.read("金額が3000以上の行と、3000未満の行を抜き出して")
    assert r.ambiguous, "比較が 2 つある依頼が断られない ── 畳む側の前提が崩れている"


def test_different_jobs_are_still_two_steps():
    """★ 陰性対照: 取り直さない引数が違えば畳まない（黙りすぎの検出）。"""
    plan = [{"op": "AGGREGATE", "args": {"group_col": "部署", "value_col": "金額"}},
            {"op": "AGGREGATE", "args": {"group_col": "担当", "value_col": "金額"}}]
    _folded, dropped = ailine.fold_identical_steps(plan)
    assert dropped == 0
    plan = [{"op": "EXTRACT", "args": {"col": "取引先", "cmp": "eq", "value": "丸和物流"}},
            {"op": "EXTRACT", "args": {"col": "金額", "cmp": "gte", "value": 40000}}]
    _folded, dropped = ailine.fold_identical_steps(plan)
    assert dropped == 0
