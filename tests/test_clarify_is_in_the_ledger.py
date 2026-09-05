"""聞き返し（CLARIFY）も台帳に残ること（2026-09-05）。

★★ 出所: 属性の登録がどれだけ効くかを測ろうとして、履歴を数えた ──

    語彙外        234 件  ← 台帳に在る
    CLARIFY       0 件    ← 台帳に無い（`_finish_run` を通らず return 3 していた）

  ★ その日の実機 3 発は**全部 CLARIFY に落ちた**。つまり「いちばん多く踏んだ経路」が
    台帳から見えないまま、「この機能は全 run の 1% に効く」と言おうとしていた。
    「出ないことは信号でない」の形そのもの ── 分母が見えていなかった。

★ 処方は記録を足すことだけでなく、**記録する処理を 1 本に畳む**こと。
  2 箇所に書けば片方だけ直る（この repo の系譜）。`_finish_run` と CLARIFY が
  同じ `_record_history` を通る。
★ ただし `_finish_run` は --json を印字する ── CLARIFY で呼ぶと --json の出力が
  変わってしまう。畳んだ側は**履歴だけ**を書く（印字しない）ことを機械で縛る。
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

import ailine  # noqa: E402

NL = chr(10)


def test_the_history_writer_is_a_single_function():
    """★ 履歴を書く実装が 2 つ無いこと（append_history の呼び出しは 1 箇所）。"""
    src = Path(ailine.__file__).read_text(encoding="utf-8")
    body = src.split("def _record_history")[1].split(NL + "def ")[0]
    assert "append_history(" in body, "畳んだ関数が履歴を書いていない"
    assert src.count("append_history(build_history_entry(") == 1, (
        "履歴を書く実装が 2 箇所ある（片方だけ直る）")


def test_finish_run_goes_through_it():
    src = Path(ailine.__file__).read_text(encoding="utf-8")
    body = src.split("def _finish_run")[1].split(NL + "def ")[0]
    assert "_record_history(" in body, "_finish_run が畳んだ関数を通っていない"


def test_the_clarify_dead_end_is_recorded():
    src = Path(ailine.__file__).read_text(encoding="utf-8")
    body = src.split("def _translate_and_dispatch")[1].split(NL + "def ")[0]
    clarify = body.split('if op == "CLARIFY":')[1][:1400]
    assert "_record_history(" in clarify, "聞き返しが台帳に残らない"
    assert '"clarify"' in clarify, "理由のラベルが付いていない"


def test_recording_does_not_print(capsys, monkeypatch, tmp_path):
    """★ --json の出力を変えないこと（印字は _finish_run の仕事のまま）。"""
    import argparse
    seen = {}
    monkeypatch.setattr(ailine, "append_history", lambda e: seen.update(e or {}))
    a = argparse.Namespace(task="これは何ですか", model="m", json=True, dry=False)
    ailine._record_history(a, tmp_path / "b.xlsx",
                           {"ok": False, "attempts": 0, "task": a.task, "model": "m",
                            "path": "clarify", "command": None, "postcondition": None,
                            "changes": [], "out": str(tmp_path / "b.xlsx")},
                           "clarify")
    assert capsys.readouterr().out == "", "履歴を書くだけのはずが印字している"
    assert seen.get("failure_kind") == "clarify", seen


def test_a_broken_history_does_not_break_the_run(capsys, monkeypatch, tmp_path):
    """★ 台帳の失敗で行き止まりの表示を壊さない（警告は stderr へ）。"""
    import argparse

    def boom(_e):
        raise OSError("ディスクが無い")

    monkeypatch.setattr(ailine, "append_history", boom)
    a = argparse.Namespace(task="これは何ですか", model="m", json=False, dry=False)
    ailine._record_history(a, tmp_path / "b.xlsx", {"path": "clarify"}, "clarify")
    got = capsys.readouterr()
    assert got.out == "" and "履歴の記録に失敗" in got.err
