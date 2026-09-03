# 画面から呼んだ子プロセスに、**端末を継がせない**番人（2026-09-02）。
#
# ★★ 実測した事故（2026-09-01・実演の練習中）: 画面が「下書きで実行中…」のまま固まった。
#   4 分動いて CPU 0.5 秒 ── 計算しておらず、**入力を待っていた**。
#   真因: `subprocess.run(..., input=None)` は「stdin を渡さない」ではなく
#   **「この画面サーバの端末を子にそのまま継がせる」**。子の `input("[y/N]: ")` は
#   人に見えない端末で待ち続ける。
#
# ★ 関所（既存データのある列へ書く時の「上書きしますか？」）の側は正しい ──
#   EOFError を拾って逃げ道を出し 7 で抜ける。**呼ぶ側が stdin を閉じていなかっただけ。**
#
# ★ `_stdin_isatty()` では直せない: この時 stdin は**本物の端末**なので True になる。
#   塞ぐべきは「端末を継ぐこと」そのもの。
#
# 契約:
#   ① `_ailine` は **必ず** `input=` を渡す（None を渡さない＝端末を継がない）
#   ② 答えがある回は、その答えが子に届く
#   ③ 関所は 1 ミリも緩めない（この試験は関所の中身を触らない）

import importlib.util
import sys
from pathlib import Path
from _product_source import product_text  # noqa: E402 ── ★ 番人は本体決め打ちでなく製品コード全体を読む

REPO = Path(__file__).resolve().parent.parent
SERVER_PY = REPO / "gui" / "server.py"


def _server_module():
    spec = importlib.util.spec_from_file_location("_ailine_gui_server_stdin", SERVER_PY)
    mod = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = mod
    spec.loader.exec_module(mod)
    return mod


class _FakeProc:
    returncode = 0
    stdout = ""
    stderr = ""


def _capture(mod, monkeypatch):
    seen = {}

    def fake_run(cmd, **kw):
        seen.update(kw)
        return _FakeProc()

    monkeypatch.setattr(mod.subprocess, "run", fake_run)
    return seen


def test_the_child_never_inherits_a_terminal(monkeypatch):
    """① 答えが無い回でも input= を渡す（None だと端末を継いで固まる）。"""
    mod = _server_module()
    seen = _capture(mod, monkeypatch)
    mod._ailine(["ops"])
    assert "input" in seen, "input= を渡していない（端末を継ぐ）"
    assert seen["input"] is not None, (
        "input=None は『渡さない』ではなく『端末を継ぐ』── 画面が固まる")


def test_an_answer_still_reaches_the_child(monkeypatch):
    """② 逃げ道を塞いでいないこと ── 人が答えた回は、その答えが届く。

    ★ 陰性対照。①だけなら「常に空文字」でも通るが、それでは関所に答えられない。
    """
    mod = _server_module()
    seen = _capture(mod, monkeypatch)
    mod._ailine(["run", "x.xlsx", "何か"], answer="y")
    assert seen["input"].startswith("y"), seen["input"]


def test_the_gate_itself_is_untouched():
    """③ 番人が関所を緩めていないこと ── EOFError の逃げ道が本体に在る。

    ★ 「画面が固まらない」を、関所を外すことで達成していないか見る。
    """
    assert "except EOFError:" in product_text()
    assert "上書きしますか？" in product_text()
