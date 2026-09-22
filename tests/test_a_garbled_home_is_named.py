"""置き場の名が壊れていたら、黙って書かずに人へ言う（2026-09-22）。

★★ 事故（盲検 6 体目 ⑨）: 買い手の画面に台帳のパスが
  `...\\bench\\blind\\6菴鍋岼\\home\\history.jsonl` と出た。同じ画面の下の行は
  `6体目` と正しく出ており、買い手は「**文字化け**」と報告した。

★★ こちらで実体を確かめたら、**表示の問題ではなかった**:
  道具はその名前で**実際にフォルダを作り**、履歴もバックアップも **undo の世代も**
  そちらへ書いていた（正しい側の `home/` は空だった）。
  ★ `ailine undo` の退避先ごと別の場所になる ── 買い手の報告より重い。
  ★ 私はその化けたフォルダを **commit にまで入れてしまった**（すぐ取り消した）。

★ 見分け方（実測で確定）: `6体目` を UTF-8 で符号化したバイト列を cp932 として
  解釈すると、ちょうど `6菴鍋岼` になる。**逆に戻して意味のある日本語になるなら、
  渡され方が壊れている** ── 偶然そうなる確率は極めて低い。

★ 直さない（勝手に読み替えない）── どちらが本物かは人しか決められない。言うだけ。
★ 配線は `main()` 1 箇所（全コマンドの共通の入口）。呼び出し側に配らない。
"""
import os
import subprocess
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

import ailine  # noqa: E402


def test_the_real_corruption_is_recognised():
    """★★ 実物 ── 買い手の画面に出た名前を、元の読みに戻せること。"""
    assert ailine.mojibake_reading("6菴鍋岼") == "6体目"


def test_ordinary_names_are_left_alone():
    """★ 陰性対照 ── 普通の名前に難癖をつけない（毎回鳴る警告は読まれなくなる）。"""
    for name in ("6体目", "home", ".ailine", "C--Windows-system32", "ailine",
                 "日本語のフォルダ", "請求一覧", "2026年8月", "Temp"):
        assert ailine.mojibake_reading(name) is None, name


def test_a_name_that_becomes_symbols_is_not_reported():
    """★ 戻した先が日本語として読めない時は言わない（記号の羅列は手がかりでない）。"""
    assert ailine.mojibake_reading("abc") is None
    assert ailine.mojibake_reading("") is None


def test_the_warning_says_where_it_is_writing(tmp_path, monkeypatch):
    """★ 「表示だけの問題ではない」ことが画面で分かること。"""
    monkeypatch.setenv("AILINE_HOME", str(tmp_path / "6菴鍋岼"))
    said = ailine.home_dir_looks_garbled()
    assert said and "6体目" in said, said
    assert "いま書いている先" in said, said
    assert "undo" in said, "★ undo の世代も行くことを言っていない: " + said


def test_a_healthy_home_says_nothing(tmp_path, monkeypatch):
    monkeypatch.setenv("AILINE_HOME", str(tmp_path / "6体目"))
    assert ailine.home_dir_looks_garbled() is None


def test_every_command_gets_the_warning(tmp_path):
    """★★ 配線 ── `history` のような読むだけのコマンドでも出ること。

    ★ 実機で確かめる（関数が返すだけでは「画面に出る」ことの証明にならない）。
    """
    repo = Path(ailine.__file__).resolve().parents[2]
    r = subprocess.run([sys.executable, "-m", "ailine", "history"],
                       capture_output=True, text=True, encoding="utf-8",
                       errors="replace", cwd=str(repo),
                       env={**os.environ, "PYTHONPATH": str(repo / "src"),
                            "AILINE_HOME": str(tmp_path / "6菴鍋岼")})
    both = (r.stdout or "") + (r.stderr or "")
    assert "置き場の名前が壊れて見えます" in both, both[-400:]
    assert "6体目" in both, both[-400:]


def test_the_check_is_wired_in_one_place():
    """★ 器官を置いて配線しない、をさせない ── 呼び出し側に配っていないこと。"""
    from _product_source import code_only_text
    lines = code_only_text().splitlines()
    called = sum(1 for ln in lines if "home_dir_looks_garbled()" in ln)
    assert called == 2, (
        f"home_dir_looks_garbled の呼び出しが {called} 箇所 ── "
        "定義 1 + main 1 のはず（呼び出し側に配らない）")
