# `ailine demo` ── 実装より先に凍結した赤い検体（2026-08-24）。
#
# ★ 出所（盲検の査定・2 回目）: 「README の最初のコマンドが落ちる。買い手の最初の 90 秒がこれ」
#     $ ailine run demo/sample.xlsx "..."
#     文書が無い: C:\Dev\ailine\demo\sample.xlsx   EXIT=9
#   `demo/` は repo に無い（実体は `src/ailine/demo/`）。しかも **install した人は
#   どちらのパスも持っていない**ので、README を直すだけでは install 経路が救われない。
#   査定者の結論: 「値段を止めているのは製品の能力ではなく、**能力に到達するまでの距離**」。
#
# 契約:
#   ① `ailine demo` で同梱サンプルが**手元に出る**（repo でも install 後でも同じに動く）
#   ② 出したあと、**次に打つコマンドをそのまま見せる**（コピペで最初の成功に届く）
#   ③ 既にあるファイルを黙って上書きしない
#   ④ 書き込めない場所でも落ちない（理由を言って exit する）

import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
import ailine  # noqa: E402

#: doctor が出す python の検査名（★ 定数から作る ── 手書きは下限を上げた日に腐る）
_PY_CHECK = f"python {ailine.MIN_PYTHON[0]}.{ailine.MIN_PYTHON[1]}+"

sys.path.insert(0, str(Path(__file__).resolve().parent))
from test_golden_transcripts import _isolate, _run_main  # noqa: E402

needs_impl = pytest.mark.xfail(
    not hasattr(ailine, "cmd_demo"),
    reason="demo コマンド 未実装（契約は凍結済み）", strict=True)


@needs_impl
def test_demo_copies_the_bundled_sample_here(tmp_path, monkeypatch, capsys):
    _isolate(monkeypatch, tmp_path)
    monkeypatch.chdir(tmp_path)
    rc, out = _run_main(["demo"], capsys)
    assert rc == 0, out
    made = sorted(p.name for p in tmp_path.glob("*.xlsx"))
    assert "sample.xlsx" in made, f"サンプルが出ていない: {made}"


@needs_impl
def test_demo_shows_the_next_command_to_type(tmp_path, monkeypatch, capsys):
    """★ 出すだけでは「距離」は縮まらない。**次に打つ行**を見せる。

    ★ 「居るから見えない」の実演（2026-08-24・CI が赤で教えた）: 初版は前提を差し替えずに
      書いたので、**開発機（ollama も LibreOffice も在る）では緑・CI では赤**になった。
      CI には ollama が無いので、正しく「先に足りないものがあります」に分岐していた。
      ── 検体が俺の環境に依存していた。**居ない側を既定にする**（明示的に揃った状態を作る）。
    """
    monkeypatch.setattr(ailine, "doctor_checks",
                         lambda model="qwen2.5-coder:7b": [[_PY_CHECK, True, ""]])
    _isolate(monkeypatch, tmp_path)
    monkeypatch.chdir(tmp_path)
    rc, out = _run_main(["demo"], capsys)
    assert "ailine run" in out and "sample.xlsx" in out, f"次の一手が書かれていない: {out}"


@needs_impl
def test_demo_does_not_overwrite_silently(tmp_path, monkeypatch, capsys):
    """③ 既にあるファイルを黙って潰さない。"""
    _isolate(monkeypatch, tmp_path)
    monkeypatch.chdir(tmp_path)
    mine = tmp_path / "sample.xlsx"
    mine.write_bytes(b"human's file")
    rc, out = _run_main(["demo"], capsys)
    assert mine.read_bytes() == b"human's file", "人のファイルを潰した"
    assert rc != 0 or "既に" in out, f"上書きを黙ってやったか、断りが無い: {out}"


@needs_impl
def test_demo_works_from_the_installed_package(tmp_path):
    """① ★ install した人は repo を持っていない。**パッケージ内の同梱**から出せること。"""
    src = ailine.bundled_demo_dir()
    assert src.exists(), f"同梱の demo が見つからない: {src}"
    assert (src / "sample.xlsx").exists()
    assert "ailine" in str(src), f"repo のパスを見ている（install 経路で壊れる）: {src}"


# --- ★ 「距離」を縮める本体: 落ちるコマンドを勧めない（2026-08-24）--------------------
#
# ★ 実測: install した人の環境で `ailine demo` はサンプルを置き、次の一手を見せたが、
#   その一手が **basrun.py が見つからない** で落ちた。
#   置いただけでは距離は縮まらない ── **前提が欠けているなら、それを先に言う。**

@needs_impl
def test_demo_names_what_is_missing_instead_of_suggesting_a_failing_command(
        tmp_path, monkeypatch, capsys):
    _isolate(monkeypatch, tmp_path)
    monkeypatch.chdir(tmp_path)
    monkeypatch.setattr(ailine, "doctor_checks",
                         lambda model="qwen2.5-coder:7b": [
                             [_PY_CHECK, True, ""],
                             ["LibreOffice", True, "C:/LO"],
                             ["basrun.py", False, "環境変数 BASRUN に…"],
                         ])
    rc, out = _run_main(["demo"], capsys)
    assert "basrun.py" in out, f"欠けている前提を名指ししていない: {out}"
    assert "環境変数 BASRUN" in out, f"直し方を見せていない: {out}"
    assert "次にこれを打って" not in out, \
        f"前提が欠けているのに、落ちるコマンドを勧めた: {out}"
    assert (tmp_path / "sample.xlsx").exists(), "サンプル自体は置いてよい"


@needs_impl
def test_demo_suggests_the_run_when_everything_is_ready(tmp_path, monkeypatch, capsys):
    """誤爆防止: 前提が揃っていれば今までどおり次の一手を見せる。"""
    _isolate(monkeypatch, tmp_path)
    monkeypatch.chdir(tmp_path)
    monkeypatch.setattr(ailine, "doctor_checks",
                         lambda model="qwen2.5-coder:7b": [[_PY_CHECK, True, ""]])
    rc, out = _run_main(["demo"], capsys)
    assert rc == 0 and "次にこれを打って" in out, out
