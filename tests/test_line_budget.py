"""C4: ailine.py の行数が単調減少することを機械で見張る番人。

★ 分割の一歩目（BookView 導入）の要件そのもの。「新しい単位は ailine.py に足さず
ailine_core/ に置く」という約束は、宣言だけでは次のリファクタで簡単に破られる。
tests/ailine_py_line_budget.txt に記録した行数を「これ以上は増やさない上限」として
凍結し、実際の行数がそれ以下であることをここで検査する。

運用:
- ailine.py が縮んだら、tests/ailine_py_line_budget.txt の数字を新しい行数まで
  **下げて** commit する（このテスト自身が「縮んだ」ことの証跡になる）。
- ★★ 記録値を**増やす方向の更新は禁止**。ailine.py に新しいコードを足したくなったら、
  それは ailine_core/ などの別モジュールに書く場所であって、記録値を緩める理由にしない。
  （どうしても増やす必要が生じた場合は、この禁止をコード上で外すのではなく、
  なぜ例外的に必要かを commit メッセージと docs/behavior-corpus に書いた上で
  人が判断して記録値を更新する。）

★ このテストが実際に赤くなることは
`AILINE_LINE_BUDGET_SELFTEST=1 pytest tests/test_line_budget.py` で自己検証できる
（記録値を一時的に1下げて赤を確認するデバッグ用フック。通常運用では使わない）。
"""
import os
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
AILINE_PY = REPO_ROOT / "ailine.py"
BUDGET_FILE = Path(__file__).resolve().parent / "ailine_py_line_budget.txt"


def _current_line_count() -> int:
    return len(AILINE_PY.read_text(encoding="utf-8").splitlines())


def _recorded_budget() -> int:
    first_line = BUDGET_FILE.read_text(encoding="utf-8").splitlines()[0]
    return int(first_line.strip())


def test_ailine_py_does_not_exceed_recorded_line_budget():
    """ailine.py の実際の行数が、記録された上限（縮んだら下げる・上げるのは禁止）を
       超えていないことを検査する。"""
    current = _current_line_count()
    budget = _recorded_budget()
    # ★ デバッグ用の自己検証: 記録値を一時的に1下げて本当に赤くなるかを実証できる
    #   （DoD 3「番人テストが発火することの実証」用のフック。通常運用では未設定）。
    if os.environ.get("AILINE_LINE_BUDGET_SELFTEST") == "1":
        budget -= 1
    assert current <= budget, (
        f"ailine.py が {current} 行で、記録された上限 {budget} 行を超えている。"
        f"新規コードを ailine.py 本体に足していないか確認する（新しい単位は "
        f"ailine_core/ 等の別モジュールに置く）。意図的に縮めた場合は "
        f"tests/ailine_py_line_budget.txt の数字を {current} 以下に下げて commit する"
    )
