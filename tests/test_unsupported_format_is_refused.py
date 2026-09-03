# 扱えない形式は、触る前に**説明して**断る番人（2026-09-02）。
#
# ★★ 実測（README の「未実装」に自分で書いていた）: `.ods` を渡すと**生の traceback**が出た。
#     openpyxl.utils.exceptions.InvalidFileException: openpyxl does not support .ods ...
#   `--help` は 3 箇所で「.xlsx / .ods」と約束していたのに、`build_book_meta` は
#   openpyxl なので読めない ── **約束だけが先行していた。**
#
# ★★ 置き場所が肝（設計判断）: この関所を `refuse_if_locked` に相乗りさせない。
#   あれは **undo も通る**。この repo は 2026-08 の盲検で出た
#   「`.ods` の拒否を全形式に広げよう」を**却下している** ── 断る範囲を広げると
#   **命綱（undo・バックアップ）に届く前に止まる経路**ができるため。
#   ★ だから関所は **run の入口 1 箇所だけ**。ここを試験で縛る（また広げられないように）。
#
# 契約:
#   ① 扱えない拡張子は、読む前に断る（traceback を出さない）
#   ② 断りは行き止まりにしない ── 直し方を言う
#   ③ 扱える拡張子は素通しする（断りを広げすぎない＝陰性対照）
#   ④ **undo には掛かっていない**（命綱を塞がない）

import re
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))

import ailine                                        # noqa: E402
from ailine_core.filetypes import RUN_SUPPORTED_SUFFIXES   # noqa: E402


def test_an_unsupported_suffix_is_refused_before_reading(tmp_path, capsys):
    """① 中身を読まずに断る ── 実体は .xlsx でも、名前が .ods なら止まる。

    ★ 「読む前」であることを、**中身を xlsx にして**確かめる:
      もし読んでから判断していたら、これは通ってしまう。
    """
    p = tmp_path / "t.ods"
    import openpyxl
    openpyxl.Workbook().save(p)                       # 中身は正しい xlsx
    rc = ailine.refuse_if_run_cannot_handle(p)
    assert rc == ailine.EXIT_ENVIRONMENT, rc
    out = capsys.readouterr().out
    assert "操作できる形式ではありません" in out
    assert "Traceback" not in out


def test_the_refusal_says_how_to_fix_it(tmp_path, capsys):
    """② 断りは行き止まりにしない（この repo の作法）。"""
    p = tmp_path / "t.ods"
    p.write_bytes(b"")
    ailine.refuse_if_run_cannot_handle(p)
    out = capsys.readouterr().out
    assert ".xlsx" in out, "直し方に、どの形式なら扱えるかを書く"
    assert "保存し直す" in out, "具体的な直し方を書く"


@pytest.mark.parametrize("suffix", sorted(RUN_SUPPORTED_SUFFIXES))
def test_supported_suffixes_pass_through(tmp_path, suffix):
    """③ 陰性対照 ── 扱える形式まで止めていないこと。

    ★ ①②だけなら「常に断る」でも通る。断りを広げすぎるのが、この repo で
      一度**却下された**方向（命綱に届く前に止まる）。
    """
    p = tmp_path / f"t{suffix}"
    p.write_bytes(b"")
    assert ailine.refuse_if_run_cannot_handle(p) is None


def test_the_gate_is_not_wired_into_undo():
    """④ 命綱を塞いでいないこと ── undo の経路にこの関所が入っていない。

    ★ ここが本当の契約。「.ods を断る」を `refuse_if_locked` に入れると
      **undo も通る場所**なので、壊れた形式のブックが復元できなくなる。
      2026-08 の盲検所見「全形式に広げよう」を却下した理由そのもの。
    """
    src = (REPO / "src" / "ailine" / "__init__.py").read_text(encoding="utf-8")
    i = src.index("def _cmd_undo_body(")
    j = src.index("\ndef ", i + 10)
    assert "refuse_if_run_cannot_handle" not in src[i:j], (
        "undo に形式の関所が掛かっている ── 命綱が塞がる")
    # ★ 呼ばれているのは 1 箇所だけ（定義を除く）
    calls = [m for m in re.finditer(r"refuse_if_run_cannot_handle\(", src)]
    assert len(calls) == 2, f"定義 1 + 呼び出し 1 のはず（実際 {len(calls)}）"


def test_the_help_no_longer_promises_ods():
    """★ 約束と実体を合わせる ── help が .ods を約束したままだと、また嘘になる。"""
    src = (REPO / "src" / "ailine" / "__init__.py").read_text(encoding="utf-8")
    hits = list(re.finditer(r'help="対象の文書 \(([^)]*)\)', src))
    # ★ 分母を先に確かめる（2026-09-03）: マッチ 0 件だとループが 1 回も回らず、
    #   help の文言が変わっただけで黙って通る ──「回らないループ」の形。
    assert hits, "help の文言が見つからない（番人が空振りしている）"
    for m in hits:
        assert ".ods" not in m.group(1), f"help がまだ .ods を約束している: {m.group(0)}"
