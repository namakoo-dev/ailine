# 比べ方（gte / lte / …）の**コードが 3 箇所でずれない**ための番人（2026-09-02）。
#
# ★★ なぜ要るか（README の「作らなかったこと」に自分で書いていた）:
#   「〜以外」の抽出をやらなかった理由は「この述語は Python・Basic・凍結した真理値表の
#   **3 箇所が独立に持つ**ので、締切前に触ると 3 つの同期がずれる」だった。
#   ★ 締切は過ぎた。だが**触る前に、ずれたら気づく形にしておく** ──
#     でないと次に足す誰か（俺を含む）が同じ理由で手を出せない。
#
# ★ 繋いでいるのは**整数のコード**だけで、名前でも型でもない。
#   Python が `ne: 7` を足して Basic に `Case 7` を書き忘れると、
#   **Basic は黙って False を返す**（Select Case の Case Else）── 条件に合う行が
#   1 行も無かったのか、比べ方を知らなかったのかが、画面では区別できない。
#   出ないことは信号でない、そのもの。
#
# 契約:
#   ① Python が持つコードを、Basic の RowMatches が**全部**扱う
#   ② Basic が扱うコードは、Python が**全部**知っている（勝手な拡張を残さない）
#   ③ 表示名（以上・以下…）も同じ顔ぶれ ── 画面に出ない比べ方を作らない

import re
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))

import ailine  # noqa: E402

BAS = REPO / "src" / "ailine" / "helpers" / "AiLineHelpers.bas"


def _basic_row_matches_codes() -> set:
    """Basic の RowMatches が `Select Case` で扱っているコードの集合。

    ★ 窓は構造で切る（語で切らない）── `Function RowMatches` から `End Function` まで。
    """
    src = BAS.read_text(encoding="utf-8")
    i = src.index("Function RowMatches")
    j = src.index("End Function", i)
    body = src[i:j]
    codes = set()
    for m in re.finditer(r"^\s*Case\s+(.+)$", body, re.M):
        # ★ Basic の行末コメント（`Case 0   ' 以上`）を落としてから読む。
        #   初版はこれを忘れて**1 つも読めず**、①② が空集合どうしで通っていた
        #   ── 陽性対照（読めていることを先に確かめる）が捕まえた。
        arg = m.group(1).split("'")[0].strip()
        if arg.lower().startswith("else"):
            continue
        for part in arg.split(","):
            part = part.strip()
            if part.isdigit():
                codes.add(int(part))
    return codes


def test_every_python_code_is_handled_by_basic():
    """① Python が知っている比べ方を、Basic が全部扱うこと。"""
    py = set(ailine._EXTRACT_CMP_CODE.values())
    bas = _basic_row_matches_codes()
    missing = sorted(py - bas)
    assert not missing, (
        f"Basic の RowMatches が扱っていないコード: {missing}"
        f"（Python: {sorted(py)} / Basic: {sorted(bas)}）── "
        "Case が無いと Basic は黙って False を返し、"
        "『条件に合う行が無い』と区別が付かない")


def test_basic_has_no_codes_python_does_not_know():
    """② 逆向き ── Basic だけが知っている比べ方を残さない。

    ★ 片側だけ見る番人は、片側だけの追加を通してしまう（今日 2 度踏んだ打ち消し合いと同じ）。
    """
    py = set(ailine._EXTRACT_CMP_CODE.values())
    bas = _basic_row_matches_codes()
    extra = sorted(bas - py)
    assert not extra, f"Python が知らないコードが Basic に在る: {extra}"


def test_every_code_has_a_japanese_label():
    """③ 画面に出ない比べ方を作らない（解釈行・断りで名前が要る）。"""
    codes = set(ailine._EXTRACT_CMP_CODE)
    labels = set(ailine._EXTRACT_CMP_LABELS)
    assert codes <= labels, f"表示名が無い比べ方: {sorted(codes - labels)}"


def test_every_code_is_implemented_by_the_python_predicate():
    """★★ 述語の**3 箇所目**（2026-09-02 に気づいた・初版はここを見ていなかった）。

      `_extract_predicate` は事後条件のための**独立実装**（Basic とは別に同じ勘定を
      書いて一致を見る）。ここに実装を足し忘れると、Basic は正しく抜き出すのに
      **事後条件が全部 False** になり、「0 行しか抜き出せていない」という
      **嘘の失敗**が出る ── 実装が無いことと、条件に合う行が無いことが区別できない。
    ★ 構造で見る: 関数の中に、その比べ方の名前が現れること。
      意味そのものは凍結した真理値表（tests/test_predicate_truth_table.py）が守る。
    """
    import inspect
    src = inspect.getsource(ailine._extract_predicate)
    missing = [c for c in ailine._EXTRACT_CMP_CODE if f'"{c}"' not in src]
    assert not missing, (
        f"_extract_predicate が扱っていない比べ方: {missing} ── "
        "Basic 側だけ実装すると、事後条件が嘘の失敗を出す")


def test_the_guard_can_actually_read_the_basic_side():
    """★ 陽性対照 ── 読めていないのに「ずれていない」と言っていないか。

    ★ 空集合どうしを比べれば ①② は必ず通る。**読めていること**を先に確かめる。
    """
    bas = _basic_row_matches_codes()
    assert len(bas) >= 5, f"Basic 側を読めていない（{bas}）"
