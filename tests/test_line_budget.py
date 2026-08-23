"""分割の番人。★2026-08-16 に測る性質を組み替えた。

## なぜ組み替えたか（オーナーの意図の訂正）

当初この番人は「`ailine.py` の行数の**上限**」を守っていた。だが実際に運用したら、
**規則が目的を裏切った**。型破壊の安全網を入れたとき、新しいロジックは全部
`ailine_core/formula_health.py` に置いたにもかかわらず、4 箇所の配線に数行必要で、
上限に収めるために**行を詰めて空行を削る**という整形が起きた。100 行以上を外へ出した
うえでの数行なのに、字面が整形を強いた。

オーナー（Namakoo）の言:「分割の目的は本体を太らせないこともあるが、**後の保守管理の
しやすさや問題発生時の切り分けをシステマティックに出来るようにしたいのが主眼**。
モジュールに分けておけば**別プロジェクトに移植**する必要が出たとき当該部分だけを
持ち出せる。だから**太ること自体は問題ない。それが筋肉なら良い。贅肉で太るのは困る**」。

**行数は筋肉と贅肉を区別しない指標だった。** そこで測る性質を 2 つに置き換えた:

1. **移植可能性**（下記 `test_ailine_core_modules_are_portable`）— これが主眼の機械化。
   `ailine_core/*` が `ailine` を import していたら、その部分だけを別プロジェクトへ
   持ち出せない。実測（2026-08-16）では全 6 モジュールが逆流ゼロで、依存は標準ライブラリと
   openpyxl のみ。**成立しているが何も守っていなかった**ので番人を付けた。
2. **行数は上限でなく「記録との一致」**（下記）— 増減どちらでも記録の更新を強制するので、
   `ailine.py` の増減が必ず diff に現れ、commit メッセージで説明されることになる。
   **増えること自体は禁止しない**（筋肉なら良い）。**見えないまま増えることを禁止する。**

★ 今日の語彙で言えば、行数は**境界**でなく**指標**だった。境界のように扱ったので歪んだ。
"""
import ast
import os
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
AILINE_PY = REPO_ROOT / "src" / "ailine" / "__init__.py"
CORE_DIR = REPO_ROOT / "src" / "ailine_core"
BUDGET_FILE = Path(__file__).resolve().parent / "ailine_py_line_budget.txt"


def _current_line_count() -> int:
    return len(AILINE_PY.read_text(encoding="utf-8").splitlines())


def _recorded_budget() -> int:
    first_line = BUDGET_FILE.read_text(encoding="utf-8").splitlines()[0]
    return int(first_line.strip())


def _imported_roots(path: Path) -> set:
    """そのファイルが import しているトップレベルのモジュール名の集合。"""
    tree = ast.parse(path.read_text(encoding="utf-8"))
    roots = set()
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            roots.update(a.name.split(".")[0] for a in node.names)
        elif isinstance(node, ast.ImportFrom):
            if node.level == 0 and node.module:
                roots.add(node.module.split(".")[0])
    return roots


def test_ailine_core_modules_are_portable():
    """★ 主眼: `ailine_core/*` が `ailine` を import していないこと。

    import していたら、そのモジュールだけを別プロジェクトへ持ち出せない
    （オーナーが挙げた分割の目的そのもの）。依存の向きは
    `ailine.py -> ailine_core` の一方通行に保つ。
    """
    offenders = {}
    for py in sorted(CORE_DIR.glob("*.py")):
        back = {r for r in _imported_roots(py) if r == "ailine"}
        if back:
            offenders[py.name] = sorted(back)
    assert not offenders, (
        f"ailine_core のモジュールが ailine を import している: {offenders}。"
        f"逆流があるとその部分だけを別プロジェクトに持ち出せない。"
        f"必要な値は引数で渡すか、依存を ailine_core 側の下位モジュールへ寄せる"
    )


def test_ailine_py_line_count_matches_the_record():
    """`ailine.py` の行数が記録と**一致**すること（上限でなく一致）。

    増減どちらでも `tests/ailine_py_line_budget.txt` の更新を強制するので、
    本体の増減が必ず diff に現れ、commit メッセージで説明されることになる。
    ★ 増えること自体は禁止しない（配線や本質的な追加なら正当）。
      **見えないまま増えること**を禁止する。
    """
    current = _current_line_count()
    recorded = _recorded_budget()
    # ★ デバッグ用の自己検証: 記録値を一時的にずらして本当に赤くなるかを実証できる
    if os.environ.get("AILINE_LINE_BUDGET_SELFTEST") == "1":
        recorded -= 1
    assert current == recorded, (
        f"ailine.py が {current} 行で、記録は {recorded} 行。"
        f"tests/ailine_py_line_budget.txt を {current} に更新し、"
        f"**なぜ増減したかを commit メッセージに書く**こと。"
        f"増えた場合は「新しい単位は ailine_core/ に置いたか」を確認する"
        f"（配線のための数行なら正当・新しいロジック本体なら置き場所が違う）"
    )
