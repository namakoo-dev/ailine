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
    bad = walk.broken_paths()
    assert not bad, "★ 通らない道を示している（2 回とも）:\n" + "\n".join(
        f"  {r['key']}: {r['detail']}\n    画面: {r.get('screen','')[:120]}" for r in bad)


def test_a_one_off_failure_is_not_called_a_broken_path():
    """★★ 再現しない失敗を赤にしないこと（前夜の偽の赤の直し）。

    ★ いま path_fails は 0 件なので、「1 回で赤／2 回で赤」の違いが**実データでは出ない**
      ── 変異試験がそこを緑で通した（2026-09-20）。だから偽の結果を注入して直接縛る。
    """
    calls = []

    def flaky():
        calls.append(1)
        v = "path_fails" if len(calls) == 1 else "walked"
        return [{"key": "★対照", "verdict": v, "detail": "偽", "screen": ""}]

    assert walk.broken_paths(flaky) == [], "★ 1 回だけの失敗を『通らない道』と断定している"
    assert len(calls) == 2, f"★ 2 回歩いていない: {len(calls)}"


def test_a_repeated_failure_is_called_a_broken_path():
    """★ 逆側: 2 回とも出たら、ちゃんと赤にすること（緩めすぎていない）。"""
    always = lambda: [{"key": "★対照", "verdict": "path_fails", "detail": "偽", "screen": ""}]
    assert walk.broken_paths(always), "★ 再現した失敗を見逃している"


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


def test_a_busy_machine_is_not_called_a_broken_path():
    """★★ 測れなかった回を「通らない道」と言わないこと（2026-09-20）。

    ★ 出所: 別の走行が実行ロックを握っている間に盤を回したら、**10 件が一度に**
      `path_fails` になった。画面は全部「別の ailine が実行中です」＝ exit 6。
      道の良し悪しは一切測れていないのに、盤は「通らない道を示した」と主張していた。
    ★ 2 回歩く形では防げない（ロックは数分握られたまま）。読むのは製品の終了コード。
    ★ pytest 側は conftest が AILINE_HOME を差し替えるので**緑のまま**だった ──
      盤が**走らせ方によって別の答えを出す**、より悪い形。だからここで直接縛る。

    ★ 対照は **6 を直に書く** ── `walk.BUSY` を使うと、定数を変える変異に対照が
      ついて回って必ず緑になる（恒真。変異試験が実際にそれを指した）。
      6 は製品が宣言している番号なので、下の試験で宣言と突き合わせる。
    ★ 引き金の側と道の側を**別々に**測る ── 片方ずつ塞ぐ。まとめて塞ぐと、
      残った片方が拾ってしまい、どちらを消しても緑になる（これも変異試験が指した）。
    """
    spec = {"kind": "book", "plan": [{"op": "CLARIFY", "question": "？"}],
            "task": "いい感じにして", "path": {"kind": "example", "task": "けい線を引いて"}}
    real = walk._run
    for where, codes in (("引き金", [6, 0]), ("道", [3, 6])):
        seq = list(codes)
        walk._run = lambda *a, **k: (seq.pop(0), "× 別の ailine が実行中です（pid=1）")
        try:
            with tempfile.TemporaryDirectory() as td:
                got = walk.walk_one("★対照", {"walk": spec}, Path(td))
        finally:
            walk._run = real
        assert got["verdict"] == "歩けなかった", (where, got)


def test_the_busy_code_is_the_one_the_product_declares():
    """★ 盤が見ている番号が、製品の宣言（終了コードの表）と同じであること。

    ★ 数を 2 箇所に持つと片方だけ動く ── だから**宣言の側**と突き合わせる。
    """
    assert walk.BUSY == 6, walk.BUSY
    doc = (REPO / "docs" / "ENGINEERING.md").read_bytes().decode("utf-8")
    assert "| 6 | 並行実行の拒否 |" in doc, "★ 終了コードの宣言が動いた（盤の BUSY を合わせること）"


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


def test_the_pass_line_and_the_board_do_not_drift():
    """★★ 合格線の 3 条目が「歩いて確かめる」と言う以上、盤が無ければ**文書が嘘**になる。

    ★ 2026-09-19 に 3 条目を書き換えた（Namakoo GO）。条が指す道具と、その道具が
      実際に出す verdict の語が、文書と機械でずれないことを縛る。
    ★ 導通率は**合格条件ではない**（実測で決めた: 買い手が導入を断る理由として名指しした
      5 件はすべて「嘘の到達」で、導通は 1 件も入っていない）。本質は path_fails が 0。
    """
    doc = (REPO / "docs" / "FROZEN-20260917-盲検の合格線.md").read_bytes().decode("utf-8")
    assert "その道が通ることを機械で歩いて確かめてある" in doc, "★ 3 条目が書き換わっていない"
    # ★ 道具を指す箇所は**条の表**と**説明**の 2 つ。片方だけ消えても気づくよう数で縛る
    #   （1 箇所しか見ていない番人は、もう片方を消す変異を緑で通した・2026-09-19）。
    assert doc.count("scripts/walk_refusals.py") == 2, (
        f"★ 条が指す道具の記載が 2 箇所でない: {doc.count('scripts/walk_refusals.py')}")
    assert "path_fails" in doc, "★ 本質（通らない道を示さない）が文書に出ていない"
    # ★★ verdict の顔ぶれは**両方向で**縛る（2026-09-19・変異試験が 2 度穴を指した）:
    #   ① 初版は文書側の 5 語しか見ておらず、「未調査」を道具から落とす変異が緑で通った
    #   ② 次は walk.ORDER を回したが、**道具から語を減らすと検査そのものが緩む**
    #      （空集合を回して常に真＝恒真の形）── だから顔ぶれを凍結して突き合わせる。
    #   ③ 2026-09-20 に `歩けなかった` が増えた（機械が塞がっていた回を path_fails と
    #     読んでいた ── 測れなかったのに「通らない」と主張していた形）。
    VERDICTS = {"path_fails", "vague", "no_path", "未調査", "未記入",
                "引き金が引けない", "walked", "by_design", "歩けなかった"}
    assert set(walk.ORDER) == VERDICTS, (
        f"★ 道具の verdict の顔ぶれが変わった: {sorted(set(walk.ORDER) ^ VERDICTS)} "
        "── 文書（合格線）と一緒に動かすこと")
    for word in VERDICTS:
        assert word in doc, f"★ verdict が文書に無い: {word}"


def test_the_rate_is_recorded_but_not_a_threshold():
    """★ 導通率で合否を決めていないこと ── 落ちるのは path_fails が在る回だけ。

    ★ ここが閾値判定に変わると、買い手が一度も不満を言っていない理由で不合格になる。
    """
    import inspect
    body = inspect.getsource(walk.main)
    assert "broken_paths(" in body, "★ 合否の根拠が『再現した path_fails』でない"
    judge = inspect.getsource(walk.broken_paths)
    assert '"path_fails"' in judge, "★ 判定が path_fails を見ていない"
    assert "walked" not in judge.split("return")[-1], "★ 件数や率で合否を決めている疑い"


def test_the_inventory_is_visible():
    """★ 盤が在庫（vague / no_path / 未記入）を**名指しで**出すこと。

    ★ 件数だけにすると、何が残っているか分からなくなる（数で書くと腐る）。
    """
    out = walk.render(walk.survey())
    assert "導通の盤" in out
    assert "walked" in out, out[:200]
