# `ailine` から見える名前の顔ぶれを凍結する番人（2026-09-02）。
#
# ★★ なぜ**分割の前に**置くか:
#   README は「単一ファイルを割るべきだが、**挙動を変えずに割ったことを確かめる番人**を
#   用意できていない」と書いている。その番人の 1 本目がこれ。
#   ★ 分割で一番起きやすいのは **import の綱渡り**（循環・遅延読み込み・名前の消失）で、
#     それは効果の検体（160 件・25 分）やゴールデンに出る**前に**捕まえたい。
#   ★ そして**消えたものは差分に出ない** ── 名前が 1 つ消えても、
#     それを呼ばないテストは全部緑のまま通る。だから**顔ぶれを数え上げる側**を置く。
#
# 何を凍結するか: `import ailine` した後にモジュールから見える**全ての名前**
#   （`_` 始まりも含む ── 内部ヘルパこそ分割で行方不明になる）と、呼べるものの**署名**。
#   ★ 「公開」だけに絞らない: この repo のテストや GUI は `_extract_predicate` や
#     `_CONFIRM_FIELDS` のような内部名にも触っている。線を引くと、線の外が黙って壊れる。
#
# 更新の仕方（意図した変更なら記録を書き換える）:
#   AILINE_REGEN_SURFACE=1 python -m pytest tests/test_public_surface_is_frozen.py
#   ★ 生成したら **git diff で中身を読むこと**。増えたぶんは意図した追加か、
#     減ったぶんは意図した削除かを、人が見てから commit する。

import inspect
import os
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))

import ailine  # noqa: E402

RECORD = Path(__file__).resolve().parent / "ailine_public_surface.txt"
REGEN = os.environ.get("AILINE_REGEN_SURFACE") == "1"


def _surface() -> list:
    """いま `ailine` から見える名前と署名（並びは決定論的）。

    ★ import したモジュールそのもの（os, json …）は数えない ── それは顔ぶれでなく依存。
    ★ 署名は「呼び出し側との契約」なので取る。取れないもの（C 実装等）は名前だけ。
    """
    out = []
    for name in sorted(vars(ailine)):
        if name.startswith("__"):
            continue
        obj = getattr(ailine, name)
        if inspect.ismodule(obj):
            continue
        kind = "class" if inspect.isclass(obj) else (
            "def" if callable(obj) else "value")
        if kind in ("def", "class"):
            try:
                sig = str(inspect.signature(obj))
            except (TypeError, ValueError):
                sig = "(?)"
            out.append(f"{kind} {name}{sig}")
        else:
            out.append(f"{kind} {name}")
    return out


def test_no_name_disappears_or_changes_shape():
    """★ 顔ぶれと署名が記録どおりであること。

    ★ 差分の読み方:
      **減った** = 分割で名前が行方不明になった（一番こわい・呼ぶ側が黙って壊れる）
      **署名が変わった** = 呼び出し側との契約が変わった
      **増えた** = 新しく足した（意図したものなら記録を更新する）
    """
    now = _surface()
    if REGEN:
        RECORD.write_text(chr(10).join(now) + chr(10), encoding="utf-8", newline=chr(10))
        return
    assert RECORD.exists(), (
        f"記録が無い: {RECORD}\n"
        "AILINE_REGEN_SURFACE=1 で生成してから、git diff で中身を確認すること")
    want = RECORD.read_text(encoding="utf-8").splitlines()
    gone = [x for x in want if x not in now]
    added = [x for x in now if x not in want]
    assert not gone, (
        f"名前または署名が消えた（{len(gone)} 件）: {gone[:8]}\n"
        "★ 分割で行方不明になっていないか確かめること。"
        "意図した削除なら AILINE_REGEN_SURFACE=1 で記録を更新する")
    assert not added, (
        f"記録に無い名前が増えた（{len(added)} 件）: {added[:8]}\n"
        "★ 意図した追加なら AILINE_REGEN_SURFACE=1 で記録を更新し、diff を読んでから commit")


def test_the_record_is_not_trivially_small():
    """★ 陽性対照 ── 記録が空/極小なら、上の比較は何も守っていない。

    ★ 今日 2 度踏んだ形（空集合どうしの比較が通る）を、ここでも先に塞ぐ。
    """
    assert RECORD.exists(), "記録が無い"
    n = len(RECORD.read_text(encoding="utf-8").splitlines())
    assert n >= 400, f"記録が小さすぎる（{n} 行）── 顔ぶれを取れていない疑い"


def test_the_core_package_is_not_imported_lazily():
    """★ 分割の綱渡りを直接見る ── 純ロジック層が import 時に解決できること。

    ★ 循環を避けるために関数の中へ import を隠すと、**呼ばれるまで壊れに気づかない**。
      いま `ailine_core` の各モジュールが素直に import できることを凍結しておく
      （分割でここが崩れたら赤くする）。
    """
    import importlib
    for m in ("cellmap", "row_identity", "filetypes", "total_row",
               "subject", "xml_readback"):
        importlib.import_module(f"ailine_core.{m}")
