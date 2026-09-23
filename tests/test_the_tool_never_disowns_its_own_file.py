# -*- coding: utf-8 -*-
"""道具は**自分が作ったファイルを他人のものだと言わない**（2026-09-20・盲検 5 体目 ①）。

★★ 出所（買い手の言葉）:

> 私は置いていません。そしてその `商品リスト.out.xlsx` を開いたら、**私が頼んだ
> 並べ替えがちゃんと済んだ状態で入っていました。**… 素人にファイルを消させるのは
> 怖いです。**作業はできていたのに、それを教えてくれない**のが一番困りました。

  関所（exit 8）で止まった run が `.out.xlsx` を黙って残し、次の run が
  「この道具が書いた記録がありません（人が置いたファイルか…）」と塞いだ。
  `run` には上書き許可のフラグが無いので、人がファイルを消すまで行き止まり。

★★ **同じ家系の 3 件目**。1・2 件目は直してあり、docstring にそう書いてある:
    2026-08-26 `_finish_failed_apply`（反映に**失敗**した経路）
    2026-09-16 `_finish_gated`（**上書き**の関所）
  それでも 3 件目が出た ── 直し方が「見つけた出口を 1 つずつ塞ぐ」だったから。

★★ だからここでは**出口を数えない**。守るのは結果の側の不変:

      どの出口から出ても、`.out` が残っているなら
        ① 画面がそれを**言う**
        ② 履歴に**自分が作ったと記録**する（次の run が作り直せる）

★ 引き金は**翻訳を固定して**引く（2026-09-20・初版は実機の翻訳に任せて skip した）。
  空回りする番人は「在っても鳴らない」── 関所に入る計画を直に渡して、毎回引く。
"""
from __future__ import annotations

import contextlib
import io
import json
import sys
from pathlib import Path

import openpyxl
import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))

import ailine  # noqa: E402

#: ★ 1 段目は語彙内（適用されて `.out` が出来る）／2 段目は語彙外（関所が止める）。
#:   ★ これが「作業はできていたのに黙って残る」形そのもの。
GATED_PLAN = [
    {"op": "SORT", "args": {"col": "品番", "order": "asc"}},
    {"op": "FREEFORM", "about": "この道具が持っていない操作"},
]


def _book(d: Path) -> Path:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "表"
    ws.append(["品番", "品名", "数量"])
    for r in [["0012", "ガラス花瓶", 3], ["0003", "タンブラー", 7]]:
        ws.append(r)
    p = d / "商品リスト.xlsx"
    wb.save(p)
    wb.close()
    return p


def _run(argv: list, plan=None) -> tuple:
    """製品を 1 回走らせ、(exit, 画面) を返す。plan を渡せば翻訳をそれに固定する。

    ★ 関所は**非対話**（`input()` が `EOFError`）で exit 8 に落ちる ── 買い手の環境も
      そうだった。ここでもそれを起こす（対話だと `input()` で止まってしまう）。
    """
    import builtins
    buf = io.StringIO()
    real, real_input = ailine.translate_task, builtins.input
    if plan is not None:
        ailine.translate_task = lambda *a, **k: {"plan": plan}

    def _no_tty(*_a, **_k):
        raise EOFError

    builtins.input = _no_tty
    try:
        with contextlib.redirect_stdout(buf), contextlib.redirect_stderr(buf):
            try:
                rc = ailine.main(argv)
            except SystemExit as e:
                rc = e.code
    finally:
        ailine.translate_task, builtins.input = real, real_input
    return rc, buf.getvalue()


def _history(home: Path) -> list:
    p = home / "history.jsonl"
    if not p.exists():
        return []
    return [json.loads(l) for l in p.read_text(encoding="utf-8").splitlines() if l.strip()]


@pytest.mark.local
def test_a_gated_run_does_not_disown_the_file_it_wrote(tmp_path, monkeypatch):
    """★★ 事故そのもの ── 関所で止まった回が残した `.out` を、次の run が拒まないこと。

    ★ 1 段目の適用に LibreOffice が要るので `-m local`。翻訳は固定するので**毎回引ける**。
    """
    home = tmp_path / "home"
    home.mkdir()
    # ★★ 2026-09-20（初版が 1 度外した）: `HISTORY_FILE` は **import 時に固定**される
    #   （`resolve_home_dir()` の docstring は「呼び出しのたび読む」と言うが、実体は定数）。
    #   だから同一プロセスで走らせる番人は、環境変数だけでは履歴の行き先を動かせない
    #   ── conftest が既定を差し替えているのと同じやり方で、明示的に向ける。
    monkeypatch.setenv("AILINE_HOME", str(home))
    monkeypatch.setattr(ailine, "HISTORY_FILE", home / "history.jsonl")
    book = _book(tmp_path)
    out = book.with_name(book.stem + ".out.xlsx")

    rc, screen = _run(["run", str(book), "品番の小さい順に並べ替えて、できないことをして",
                       "--copy"], plan=GATED_PLAN)
    assert rc != 0, f"★ 関所が止めていない（exit {rc}）:\n{screen[-600:]}"
    assert out.exists(), (
        "★ 引き金が引けていない ── 関所で止まったのに `.out` が残っていない:\n"
        + screen[-600:])

    # ① 残したことを画面が言う（人が次に何を見ればいいか分かる）
    assert out.name in screen, "★ `.out` を残したのに画面が黙っている:\n" + screen[-700:]
    # ② 履歴に自分が作ったと記録する
    outs = [r.get("out") for r in _history(home) if r.get("out")]
    assert any(out.name in str(o) for o in outs), (
        f"★ 自分が作った `.out` を履歴に残していない: {outs}")

    # ★★ 三項の実体 ── 次の run が**塞がれない**こと（買い手が踏んだ行き止まり）
    rc2, screen2 = _run(["run", str(book), "品番の小さい順に並べ替えて", "--copy"],
                        plan=[GATED_PLAN[0]])
    assert "書いた記録がありません" not in screen2, (
        "★ 1 分前に自分が作ったファイルを『人が置いた物』と言っている:\n" + screen2[-700:])


def test_every_gate_exit_goes_through_the_finisher():
    """★★ 出口を**数えない**形で縛る ── 4 件目を出さないために。

    ★ 1・2 件目は「見つけた出口を塞ぐ」で直した。3 件目が出たのは、その直し方が
      **出口の一覧を持っていた**から（持てば必ず漏れる ── この repo の系譜どおり）。
    ★ ここは AST で、関所から抜ける `except` が後始末を通っているかだけを見る。
      出口を足した日に自動で縛られる（一覧を持たない）。
    """
    import ast

    from _product_source import product_files
    # ★ 2026-09-23: 本体 1 冊でなく製品全体の try を見る（出口が ailine_core へ移っても縛る）。
    seen, bad = 0, []
    for path in product_files():
        src = path.read_bytes().decode("utf-8")
        for node in ast.walk(ast.parse(src)):
            if not isinstance(node, ast.Try):
                continue
            for handler in node.handlers:
                text = ast.get_source_segment(src, handler) or ""
                if "GateAbort" not in text:
                    continue
                seen += 1
                if "_finish_" not in text:
                    bad.append(f"{path.name}:{handler.lineno}")
    assert seen, "★ 関所から抜ける出口が 1 つも見つからない（この検査が空回りしている）"
    assert not bad, (
        "★ 関所で抜ける出口が後始末を通っていない（`.out` が孤児になる）: 行 "
        + ", ".join(bad))
