# -*- coding: utf-8 -*-
"""断りが示した道を、機械が**実際に歩く**（2026-09-19・導通試験）。

★★ 出所（Namakoo「到達率を極大にするのは大事だけど、**漏らした依頼を適切に判断し
  ユーザを正解に導いてやること**までやるから」）: 合格線の 3 条目「通る道を示す」は、
  機械が escape/example（文言の性質）しか見ていなかった。**その道が通るか**は
  誰も確かめていない ── 宣言 vs 実体の隙間。

★★ 判定 5 つ（Namakoo 承認）: walked / by_design / vague / no_path / path_fails。
  ★ path_fails は no_path より重い ── 道が無いのは不足、通らない道を示すのは誤情報。
  ★ 道具を作る前に 1 件を手で歩いた時点で path_fails が 1 件出た
    （--op ADD_ROW → 「その列のデータ行を全部書き換えます」は ADD_ROW には嘘）。

★ ここが縛るのは **A 群（LLM を回さない 7 件）**。翻訳を固定して引き金を引くので
  素の環境でも毎回走る。B 群（実機が要る 10 件）は -m local 側の仕事。
"""
from __future__ import annotations

import sys
import tempfile
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "tests"))

import walk_refusals_core as walk  # noqa: E402


def test_no_refusal_shows_a_path_that_does_not_work():
    """★★ いちばん重い契約: **通らない道を示していない**こと。

    ★ ここが赤くなったら、買い手に嘘の案内をしている（返金の話に直結する形）。
    """
    bad = [r for r in walk.survey() if r["verdict"] == "path_fails"]
    assert not bad, "★ 通らない道を示している:\n" + "\n".join(
        f"  {r['key']}: {r['detail']}\n    画面: {r.get('screen','')[:120]}" for r in bad)


def test_the_tool_can_actually_report_a_broken_path():
    """★★ 陽性対照: わざと通らない道を渡して path_fails が出ること。

    ★ これが無いと「path_fails 0 件」が**守れている**のか**見られていない**のか
      分からない（この repo が何度も踏んだ「在っても鳴らない」）。
    """
    spec = {"kind": "folder",
            "plan": [{"op": "EXTRACT", "args": {"col": "金額", "cmp": "まんなか", "value": 100}}],
            "task": "金額がまんなかの行を抜き出して",
            "path": {"kind": "example", "task": "この依頼は通らないはずの文です"}}
    with tempfile.TemporaryDirectory() as td:
        got = walk.walk_one("★対照", {"walk": spec}, Path(td))
    assert got["verdict"] == "path_fails", got


def test_a_specimen_that_stopped_triggering_is_not_counted_as_walked():
    """★★ 恒真殺し: 断りが出なくなった検体を、黙って walked に数えないこと。

    ★ 引き金が古くなると「道を歩いた」ではなく「そもそも断られなかった」になる。
      そこを walked と数えると、盤は直った顔のまま何も見ていない。
    """
    spec = {"kind": "folder",
            "plan": [{"op": "EXTRACT", "args": {"col": "金額", "cmp": "gte", "value": 100}}],
            "task": "金額が100以上の行を抜き出して",
            "path": {"kind": "example", "task": "金額が100以上の行を抜き出して"}}
    with tempfile.TemporaryDirectory() as td:
        got = walk.walk_one("★対照", {"walk": spec}, Path(td))
    assert got["verdict"] == "引き金が引けない", got


def test_the_verdict_is_never_written_in_the_register():
    """★★ 恒真殺し: 台帳に verdict を書かない ── 歩いた結果から決める。"""
    reg = walk.load_register()
    wrote = [k for k, v in reg["refusals"].items()
             if "verdict" in (v.get("walk") or {})]
    assert not wrote, f"★ 台帳に verdict が書いてある（名簿を写して名簿と比べる形）: {wrote}"


def test_every_walkable_refusal_keeps_its_trigger_fresh():
    """★ A 群の検体は**引き金が引けている**こと（断りが実際に出る）。

    ★ 製品が変わって断りが出なくなったら、その検体は古い ── 直すのは検体の側。
    """
    stale = [r for r in walk.survey() if r["verdict"] == "引き金が引けない"]
    assert not stale, "★ 断りが出なくなった検体がある（検体を直すこと）:\n" + "\n".join(
        f"  {r['key']}: {r['detail']}" for r in stale)


def test_the_inventory_is_visible():
    """★ 盤が在庫（vague / no_path / 未記入）を**名指しで**出すこと。

    ★ 件数だけにすると、何が残っているか分からなくなる（数で書くと腐る）。
    """
    out = walk.render(walk.survey())
    assert "導通の盤" in out
    assert "walked" in out, out[:200]
