# -*- coding: utf-8 -*-
"""配線盤の番人（2026-09-16）。

★ 守るのは 4 つ:
   ① 盤面が痩せていないこと（op も列も落ちていない ── 空回りの検出）
   ② 実測の導き方を持つ列が**減っていない**こと
      （減らすのは「見えなくする」向きの変更。増やすのがこの道具の使い道）
   ③ 宣言と実測の食い違いは、台帳に**理由つきで宣言**されていること
   ④ ★ 未調査・無防備の列を、盤が**目立たせている**こと
      （理由欄を埋めただけで「片付いた」と読ませない）

★★ なぜ番人が要るか: 依存の図は 13 日間、生成器と番人が**同じ盲点**を共有していたせいで
  48% 間違ったまま緑だった。盤も同じ轍を踏みうる。だからここでは
  「生成器の出力と一致するか」を見ない ── **実体から導いた集合**と突き合わせる。

★ import は tests 側の中身（wiring_board_core）から。scripts/ を import すると
  素の環境の番人に「宣言外の依存」として止められる（2026-09-16 に refresh_records で実測）。
"""
import json
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
REGISTER = Path(__file__).resolve().parent / "wiring_board_register.json"

sys.path.insert(0, str(REPO / "tests"))
sys.path.insert(0, str(REPO / "src"))
import ailine  # noqa: E402
from wiring_board_core import mismatches, survey  # noqa: E402

STANCES = {"derived", "watched", "explained", "unstudied", "bare"}


def _register() -> dict:
    return json.loads(REGISTER.read_bytes().decode("utf-8"))


def test_the_board_is_not_empty():
    """① 空回りの検出 ── 盤面が痩せたら、下の試験は全部素通りする。"""
    d = survey()
    assert len(d["ops"]) == len(ailine.OP_SCHEMA), "op が盤面から落ちている"
    assert len(d["columns"]) >= 15, f"判断の列が少なすぎる: {len(d['columns'])}"


def test_every_partial_roster_is_a_column():
    """① 部分の名簿は全部、盤面の列になること（新しい名簿が黙って盤外に出ない）。"""
    from test_op_completeness import discover_op_rosters
    n = len(ailine.OP_SCHEMA)
    want = {k.split(":")[-1] for k, v in discover_op_rosters().items() if len(v["ops"]) < n}
    got = {c["name"] for c in survey()["columns"]}
    assert want <= got, f"盤面に出ていない名簿: {sorted(want - got)}"


def test_every_column_has_a_known_stance():
    """④ 立場は 5 つのどれか（新しい名前を黙って増やさない）。"""
    bad = {c["name"]: c["stance"] for c in survey()["columns"]
           if c["stance"] not in STANCES}
    assert not bad, f"知らない立場: {bad}（許される: {sorted(STANCES)}）"


def test_the_number_of_verified_columns_does_not_shrink():
    """② 実測の導き方は**減らさない**。

    ★ 減らせば盤面は静かになるが、見えなくなるだけ。増やすためにこの数を縛る。
    ★ 増やしたら、この数を上げて commit する（上げ忘れは赤にならない ── そこは人の仕事）。
    """
    floor = int(_register()["verified_columns_at_least"])
    got = [c["name"] for c in survey()["columns"] if c["stance"] == "derived"]
    assert len(got) >= floor, (
        f"実測の導き方を持つ列が {len(got)} 本に減った（下限 {floor}）: {got}。"
        " 導き方を消すのは『見えなくする』向きの変更です")


def test_every_mismatch_is_declared_with_a_reason():
    """③ 宣言と実測のずれは、台帳に理由つきで在ること（黙って放置しない）。"""
    declared = _register()["known_mismatches"]
    for m in mismatches(survey()):
        col = m["column"]
        assert col in declared, (
            f"宣言と実測が食い違っているのに台帳に無い: {col} / "
            f"宣言だけ {m['declared_only']} / 実測だけ {m['derived_only']}。"
            " tests/wiring_board_register.json に理由を書くこと（★ 直すのはここではない）")
        assert str(declared[col].get("why", "")).strip(), f"{col} の理由が空"


def test_no_declaration_outlives_its_mismatch():
    """③ 腐り防止 ── 揃ったのに台帳に残っていたら赤（台帳は縮めるためのもの）。"""
    now = {m["column"] for m in mismatches(survey())}
    stale = sorted(set(_register()["known_mismatches"]) - now)
    assert not stale, f"もう食い違っていないのに台帳に残っている: {stale}"


def test_unstudied_columns_are_not_hidden_among_the_explained():
    """④ ★★ 「未調査」を「導けない」に混ぜないこと。

    ★ 2026-09-17 に踏みかけた: 第 3 の色（導けない理由）を入れた瞬間、盤の
      「無防備」が 0 になった。だが理由欄に**自分で「★ 未調査」と書いた**列が
      3 本あり、調べ終えた列と同じ色で並んでいた ── 理由を書いただけで
      片付いたように見える、いちばん静かな嘘。
    """
    from wiring_board_core import NO_DERIVATION_REASON
    cols = {c["name"]: c for c in survey()["columns"]}
    for name, why in NO_DERIVATION_REASON.items():
        if name not in cols or cols[name]["stance"] == "derived":
            continue
        if "未調査" in why:
            assert cols[name]["stance"] == "unstudied", (
                f"{name} の理由は『未調査』なのに立場が {cols[name]['stance']} ── "
                "調べ終えた列と同じ色にしない")


def test_the_reason_is_written_for_every_column_without_a_derivation():
    """④ 導出も番人も無い列は、理由が書いてあること（無ければ ★ 無防備で出る）。"""
    bare = [c["name"] for c in survey()["columns"] if c["stance"] == "bare"]
    assert not bare, (
        f"導出も番人も理由も無い列: {bare} ── "
        "tests/wiring_board_core.py の NO_DERIVATION_REASON に、"
        "調べた結果か『★ 未調査』かを書くこと")


def test_watchers_are_real_test_functions():
    """★ 「番人あり」の根拠が、実在する試験を指していること（数だけ上げない）。

    ★★ 2026-09-17: 初版は for の中だけで assert していて、「番人あり」の列が
      0 本なら**1 回も回らずに緑**だった（tests/test_guard_ledger.py が捕まえた）。
      分母を先に確かめる ── 回らない番人は、在っても鳴らない。
    """
    watched = [c for c in survey()["columns"] if c["stance"] == "watched"]
    assert watched, "『番人あり』の列が 1 本も無い（下の検査が空回りする）"
    for c in watched:
        assert c["watchers"], f"{c['name']}: 番人ありなのに名前が空"
        for w in c["watchers"]:
            fname, _, tname = w.partition(":")
            p = REPO / "tests" / fname
            assert p.is_file(), f"{c['name']}: 実在しない試験ファイル {fname}"
            assert f"def {tname}(" in p.read_text(encoding="utf-8", errors="replace"), (
                f"{c['name']}: {fname} に {tname} が無い")
