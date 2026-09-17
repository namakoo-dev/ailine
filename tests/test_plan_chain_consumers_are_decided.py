# -*- coding: utf-8 -*-
"""連鎖の受け手（PLAN_CHAIN_CONSUMER_OPS）を、規則 1 本＋**例外を名指し**で決める（2026-09-17）。

★ なぜ在るか（仕分け⑤）: 配線盤に「★ 未調査」で出ていた最後の列。
  名簿の頭には「表の行を読む op（連鎖の対象）。書式だけを触る op は元表に掛けたい場合も
  多いので外す」と 1 行だけ書いてある。ところが宣言と突き合わせると、その 1 行では
  説明のつかない op が 3 つ在った ── **どれも理由がどこにも書かれていない**。

★★ ここは「機械に決まらない」側の列である。連鎖に入れるかどうかは
  「**絞り込んだ結果に掛けたいのか、元の表に掛けたいのか**」という人の意図の問題で、
  実装からは導けない。だから導けるのは**既定**までで、残りは名指しの例外にする。

★ 既定（宣言から引く）: 「新しい表・新しい行・新しい列を作る」と宣言していて、
  かつ「行をずらす」（＝構造をいじる）と宣言していない op は受け手にする。
  ── 行をずらす op（ADD_ROW / ADD_COLUMN / INSERT_ROWS）は、絞り込んだ写しでなく
     元の表に足したいことがほとんどなので、既定で外れる。

★ 例外は 3 つ。1 つずつ理由を書く（下の EXCEPTIONS）。数が増える日が来たら、
  それは「既定の引き方が実態に合わなくなった」合図なので、規則の側を見直すこと。

★★ EXTRACT_COLUMNS について（2026-09-17 に実測して据え置いた）:
  規則は「受け手にせよ」と言うのに外れている。入れるべきかを実機で測ろうとしたが、
  **2 つの形でどちらも先に断られ、害に到達できなかった** ──
    ・依頼文に 1 段目の列名が在る → 列抽出が「全部の列が指定されています」で断る
    ・依頼文から外す               → 三項の番人が ⚠ を出して止める（exit 7）
  到達できない＝未確認であって無害ではない。だから**現状のまま**据え置き、
  到達できた日に入れる（当て推量で連鎖の規則を広げない）。
  ★ 併せて分かったこと: 残す列は依頼文**全体**から機械が拾うので、複数段の計画では
    他の段が言った列名まで拾う。上の 1 つ目の断りはそれが原因。
"""
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
import ailine  # noqa: E402

#: 「新しいものを作る」と読める宣言。
MAKES_SOMETHING_NEW = {"new_sheet", "new_row_at_end", "new_column"}
#: 構造をいじる宣言（元の表に掛けたい側）。
SHIFTS_THE_TABLE = {"row_shift"}


def consumers_by_rule() -> set:
    """既定の受け手 ── 宣言から引く（手書きしない）。"""
    out = set()
    for op in ailine.OP_SCHEMA:
        writes = set(getattr(ailine.OP_WRITE_TARGET.get(op), "writes", ()) or ())
        if writes & MAKES_SOMETHING_NEW and not (writes & SHIFTS_THE_TABLE):
            out.add(op)
    return out


#: 既定から外れる op と、その理由。★ ここが名簿の**本体**（規則は足場に過ぎない）。
EXCEPTIONS = {
    "CHART": "既定は外すと言う（書式しか書かない）が**入れている** ── グラフは"
             "行を読んで描くもので、絞り込んだ結果をグラフにするのが自然。"
             "元の表のグラフが欲しい時は段を分けて頼める",
    "SORT": "既定は外すと言う（並べ替えは位置の操作）が**入れている** ── 並べ替えは"
            "行を読んで順を決める。同じ『並べ替え』でも SWAP / MOVE_COLUMN は"
            "位置の指定だけで行を読まないので、そちらは外れたままで正しい",
    "EXTRACT_COLUMNS": "既定は入れると言う（新しいシートを作る）が**外している** ── "
                       "理由が書かれていなかった。2026-09-17 に入れるべきかを実機で"
                       "測ろうとしたが、2 つの形でどちらも先に断られ**害に到達できなかった**"
                       "（依頼文に 1 段目の列名が在れば『全部の列が指定されています』、"
                       "無ければ三項の番人が ⚠ で止める）。未確認であって無害ではないので"
                       "据え置き、到達できた日に入れる",
}


def test_the_rule_is_not_empty():
    """★ 足場が崩れたら下の検査は全部素通りする ── 分母を先に確かめる。"""
    by_rule = consumers_by_rule()
    assert len(by_rule) >= 8, f"既定の受け手が少なすぎる: {sorted(by_rule)}"
    assert ailine.PLAN_CHAIN_CONSUMER_OPS, "★ 名簿が空"


def test_every_departure_from_the_rule_is_named_with_a_reason():
    """★★ 既定と名簿の差は、**ちょうど例外の集合**であること。

    ★ ここが等号でないと、4 つ目の例外が黙って入る。盤の食い違いは「列が食い違っている」
      までしか言わないので、**どの op が食い違っているか**はここで縛る。
    """
    declared = set(ailine.PLAN_CHAIN_CONSUMER_OPS)
    diff = consumers_by_rule() ^ declared
    assert diff == set(EXCEPTIONS), (
        f"既定から外れる op が例外の台帳と食い違う: "
        f"理由が無い {sorted(diff - set(EXCEPTIONS))} / "
        f"もう外れていないのに残っている {sorted(set(EXCEPTIONS) - diff)}")


def test_every_exception_has_a_reason():
    """★ 「例外である」だけ書いて理由を書かない、を許さない。"""
    assert EXCEPTIONS, "★ 例外の台帳が空（上の等号が空集合どうしで通ってしまう）"
    for op, why in EXCEPTIONS.items():
        assert op in ailine.OP_SCHEMA, f"知らない op の例外: {op}"
        assert len(str(why).strip()) >= 20, f"{op} の理由が短すぎる"


def test_the_roster_is_actually_read_by_the_chain():
    """★★ 宣言どうしで閉じない ── 名簿が**本当に連鎖の判定を変える**ことを実物で見る。

    ★ 名簿の中身を見るだけの試験は、chain_target_sheet が名簿を読まなくなった日に
      気づけない（「在っても鳴らない」の形）。判定の側から確かめる。
    """
    derived = [{"step": 0, "op": "EXTRACT", "sheet": "金額が5000以上"}]
    sheets = ["売上", "金額が5000以上"]
    headers = {"売上": ["取引先", "金額"], "金額が5000以上": ["取引先", "金額"]}
    task = "金額が5000以上の行を抜き出して、合計を出して"
    inside = ailine.chain_target_sheet("AGGREGATE", task, derived, sheets, headers)
    outside = ailine.chain_target_sheet("BOLD", task, derived, sheets, headers)
    assert inside == "金額が5000以上", f"受け手が直前の出力を見ていない: {inside!r}"
    assert outside is None, f"受け手でない op が連鎖に乗った: {outside!r}"
