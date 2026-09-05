"""読み直しの層が**どこに住んでいるか**を、テスト側で 1 箇所だけが知る。

★★ なぜ在るか（2026-09-05）: 読み直しの 15 塊を `_translate_and_dispatch`（681 行）から
  `_reread_the_plan` へ切り出したとき、**5 つの番人が同時に落ちた** ──
  4 つの試験がそれぞれ `def _translate_and_dispatch(` を手書きしていたからだ。

  ★ 落ちたこと自体は正しい（黙って分母がゼロになるより遥かに良い）。
    問題は**直す場所が 4 つあった**こと ── この repo の系譜「二重化した経路は
    片配線が既定で起きる」が、製品コードでなく**試験側**で起きていた。

★ もう 1 つ同じ日に学んだこと: 3 つの番人が「断りの直後に抜ける」を `"return 3"` という
  **字面**で守っていた。切り出しで `return plan, 3` になった瞬間に全部外れた。
  ★ 不変（＝抜けること）は保たれているのに、字面が変わっただけで鳴る番人は
    リファクタのたびに緩められる。**意味で見る**（exits_with_three）。
"""
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from _product_source import product_text  # noqa: E402

#: 層が住んでいる関数の定義行（★ 移したらここだけ直す）
HOME = "def _reread_the_plan("

#: 「終了コード 3 で抜ける」の書き方（層の中は (plan, rc) を返す形）
_EXIT_RE = re.compile(r"return\s+(?:plan\s*,\s*)?3\b")


def segment() -> str:
    """読み直しの層の本文（次の def まで）。"""
    text = product_text()
    i = text.index(HOME)
    j = text.index("\ndef ", i + 10)
    return text[i:j]


def window_after(marker: str, after: int = 900) -> str:
    """層の中で marker が現れる所から後ろを切り出す。"""
    seg = segment()
    return seg[seg.index(marker):][:after]


def exits_with_three(text: str) -> bool:
    """★ 字面でなく意味で見る ── `return 3` でも `return plan, 3` でも真。"""
    return _EXIT_RE.search(text) is not None
