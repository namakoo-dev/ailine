# 断片ガードの二重実装の同値性番人（完成度レビュー 次元⑤ 所見2・2026-08-23）
# ailine.py の _raw_target_not_embedded_in_task と ailine_core/alias_store.py の
# phrase_is_standalone_in_task は依存方向の制約（ailine_core は ailine を import しない）
# による意図的な写経 ── どちらかだけ例外が足されて判定がずれる将来を、この番人が塞ぐ。
# ずれたら: どちらが正か決めて**両方**を直し、この検体に境界例を足す。

import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
import ailine  # noqa: E402
from ailine_core import alias_store  # noqa: E402

CASES = [
    ("金額", "税込金額の列を作って"),        # 断片（漢字の内部）→ False
    ("金額", "金額で並べて"),                 # 独立の語 → True
    ("金額", "合計金額と金額で並べて"),       # 断片+独立の混在 → True
    ("並べ替え", "並べ替えという名前の列"),   # ひらがな境界 → True
    ("単価", "単価表と突合して"),             # 漢字連続の内部 → False
    ("単価", ""),                             # 空 task → False
    ("", "金額で並べて"),                     # 空 phrase → False
    ("A", "A列で並べて"),                     # ASCII 1 字
    ("金額", "金額"),                         # 完全一致
    ("コード", "商品コードで突合"),           # カタカナ末尾の境界
]


@pytest.mark.parametrize("phrase,task", CASES)
def test_two_implementations_agree(phrase, task):
    a = ailine._raw_target_not_embedded_in_task(phrase, task)
    b = alias_store.phrase_is_standalone_in_task(phrase, task)
    assert a == b, f"二重実装がずれた: ailine={a} alias_store={b} ({phrase!r}, {task!r})"
