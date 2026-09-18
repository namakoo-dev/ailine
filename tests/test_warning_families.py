# -*- coding: utf-8 -*-
"""⚠ を出す家系を数える（2026-09-18・Namakoo「三項全ては成り立っているか？ 漏れはない？」）。

★★ 調べて分かったこと: この repo には台帳が並んでいる ── 事後条件・読み直し・断り・
  配線・記法。ところが **⚠ を出す家系だけ台帳が無かった**。個別の試験は在るが
  「家系がいくつ在るか」を持つ物が無く、**増えても減っても誰も気づかない**形だった。

★★ 置いた瞬間に手勘定が 1 つ直った:
    人が `warning_count += 1` を grep して数えた   → 6
    実装から AST で導いた                          → 7
  落ちていたのは `warning_count += count_suspicious_advisories([msg])` の形
  （出力の忠実性の警告）。★ 名簿は手書きせず**宣言から導く**（棚の線）。

★★ 第 2 の列（Namakoo「漏れは塞ぐけど実際に欠陥とは限らないのか」）:
  7 家系はどれも三項のうち **2 項しか持たない**。だがそれ自体は欠陥ではない ──
  棚の処方「運べない項があるなら主張の範囲を狭める」に従い、どれも ✓ を名乗らず
  ⚠ だけを出す。**欠陥かどうかは盲点を 1 つずつ検体で確かめて決める**。
  ★ だから台帳は盲点ごとに measured（見えないと実測した）/ assumed（そう思っている
    だけ）を持つ。assumed は「安全」ではなく「未調査」。
"""
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "tests"))

import warning_families_core as board  # noqa: E402


def test_the_register_matches_what_the_product_actually_does():
    """★★ 契約①: 実装から導いた家系と台帳が一致すること。

    ★ 家系が増えれば「台帳に無い」で赤、消えれば「実装から消えた」で赤。
      どちらも**黙って起きてはいけない**変化。
    """
    bad = board.mismatches(board.survey())
    assert not bad, "台帳と実装がずれている:\n" + "\n".join(
        f"  {b['key']}  {b['kind']}  {b['where']}" for b in bad)


def test_the_families_are_derived_not_copied():
    """★★ 恒真殺し: 導出が台帳を読んでいないこと（名簿を写して名簿と比べない）。

    ★ 記法の盤（2026-09-18・同日）で実際に踏んだ形。導出の中で名簿を読むと、
      何を書いても緑になる。
    """
    import inspect
    body = inspect.getsource(board.derive)
    assert "load_register" not in body and "REGISTER" not in body, (
        "★ 導出が台帳を読んでいる（恒真）")
    assert "ast" in body or "_Finder" in body, "★ 導出が実装を見ていない"


def test_every_family_declares_which_terms_it_has():
    """★ 三項のどれを持つかを、家系ごとに**書かせる**（空欄を許さない）。

    ★ ここが空だと「2 項しかない」ことが見えなくなり、⚠ を ✓ と同じ強さだと
      読む人が出てくる。
    """
    reg = board.load_register()
    for key, rec in reg["families"].items():
        t = rec.get("terms") or {}
        for term in ("request", "declaration", "artifact"):
            assert term in t, f"{key} の terms に {term} が無い"
            assert t[term] in (True, False, "partial"), f"{key}.{term}: {t[term]}"
        assert rec.get("label"), f"{key} に人が読む名前が無い"


def test_no_family_claims_all_three_terms_without_saying_so():
    """★★ 三項そろったと名乗る家系が出たら、それは**大きな変化**なので手で見る。

    ★ いまは 7 家系すべてが 2 項。もし 3 項そろった家系が入ったら、それは ⚠ でなく
      ✓ を名乗れる強さになったということ ── 黙って通していい変更ではない。
    """
    reg = board.load_register()
    full = [k for k, r in reg["families"].items()
            if all((r.get("terms") or {}).get(t) is True
                   for t in ("request", "declaration", "artifact"))]
    assert not full, (
        f"★ 三項そろった家系が現れた: {full} ── "
        "⚠ でなく ✓ を名乗れる強さかを人が判断すること（台帳の not_a_defect を読む）")


def test_every_family_names_at_least_one_blind_spot():
    """★★ 「見えない形」を書かせる ── 空欄は「穴が無い」ではなく「調べていない」。

    ★ 2 項しか持たない関所に盲点が 0 件ということはありえない。空なら未記入。
    """
    reg = board.load_register()
    for key, rec in reg["families"].items():
        assert rec.get("blind_spots"), f"{key} に盲点が 1 つも書かれていない"
        for b in rec["blind_spots"]:
            assert b.get("status") in ("measured", "assumed"), f"{key}: {b}"
            assert b.get("what"), f"{key} の盲点に中身が無い"


def test_a_measured_blind_spot_points_at_evidence():
    """★★ 「測った」と名乗るなら、**どこで測ったか**を指すこと。

    ★ ここが無いと measured が「たぶん大丈夫」の言い換えになる ── この repo が
      何度も踏んだ「調べた≠確かめた」。
    """
    reg = board.load_register()
    thin = []
    for key, rec in reg["families"].items():
        for b in rec["blind_spots"]:
            if b.get("status") == "measured" and not (b.get("specimen") or b.get("note")):
                thin.append(f"{key}: {b.get('what')}")
    assert not thin, "★ 測ったと書いてあるのに出所が無い:\n  " + "\n  ".join(thin)


def test_the_unstudied_ones_stay_visible():
    """★★ assumed を「片付いた」と同じ色にしない（配線盤が 2026-09-17 に学んだ線）。

    ★★ 初版は「★ 未調査 が出力に在ること」を縛っていて、**未調査を全部調べ終えた
      瞬間に赤くなった**（2026-09-18・実際に踏んだ）。縛るべきは在庫の有無ではなく
      **在るなら名指しすること**。だから合成の盤を通して描画の契約だけを見る
      ── 本物の台帳の中身に依存させない（依存させると、調べるたびに番人が壊れる）。
    """
    fake = {"commit": "-", "declared_only": [], "families": [{
        "key": "x", "label": "作り物の家系", "sites": ["x.py:1"], "declared": True,
        "terms": {"request": True, "declaration": False, "artifact": True},
        "blind_spots": [{"what": "まだ見ていない形", "status": "assumed"},
                         {"what": "測った形", "status": "measured", "note": "出所"}]}]}
    out = board.render(fake)
    assert "まだ調べていない盲点（assumed）: 1" in out, f"未調査の件数が出ていない: {out}"
    assert "★ 未調査" in out and "まだ見ていない形" in out, (
        f"未調査の中身が名指しされていない: {out}")
    # ★ 実物の盤でも件数の行は必ず出ること（0 件でも「0」と言う ── 黙らない）。
    assert "まだ調べていない盲点" in board.render(board.survey())
