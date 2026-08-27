# 実機 LibreOffice が要る試験の**本数**を、文書と実測で突き合わせる番人。
#
# ★ 実測した嘘（2026-08-27）: README（当時）は 3 箇所で「1 本」と書いていた。実際は 14 本。
#   version 0.1.1 のころに 1 本だったものが増え続け、**誰も直さなかった**。
#   数は人の記憶で守れない ── 「数字は機械で守る」を、この文書にも適用する。
#
# 契約（判定には三項が要る: 宣言・実体・そして**別実装**の分母）:
#   ① 宣言 = docs/ENGINEERING.md の <!-- LOCAL_TESTS -->N<!-- /LOCAL_TESTS -->
#   ② 実体 = pytest 自身に数えさせた `-m local` の収集件数（この試験と別プロセス）
#   ③ 宣言は**この repo に 1 箇所しかない**（2 箇所に書くと片方が古くなる = 片配線）
import re
import subprocess
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
ENGINEERING = REPO / "docs" / "ENGINEERING.md"
MARK = re.compile(r"<!--\s*LOCAL_TESTS\s*-->\s*(\d+)\s*<!--\s*/LOCAL_TESTS\s*-->")


def _declared_all() -> list:
    """repo 内の文書すべてから宣言を拾う（1 箇所しか無いことを確かめるため全部見る）。"""
    found = []
    for md in sorted(REPO.glob("*.md")) + sorted((REPO / "docs").glob("*.md")):
        for m in MARK.finditer(md.read_text(encoding="utf-8")):
            found.append((md.relative_to(REPO).as_posix(), int(m.group(1))))
    return found


def _collected() -> tuple:
    """★ 分母を別実装から取る: 印を grep で数えず、**pytest 自身**に選ばせて数えさせる
       （parametrize や skipif で本数は変わる ── 印の個数と収集件数は同じ物ではない）。
       返り値: (local の本数, 全体の本数) ── 1 回の収集から両方を読む。"""
    r = subprocess.run(
        [sys.executable, "-m", "pytest", "-m", "local", "--collect-only", "-q",
         "-p", "no:cacheprovider", "tests"],
        cwd=str(REPO), capture_output=True, text=True, encoding="utf-8", errors="replace")
    out = (r.stdout or "") + (r.stderr or "")
    m = re.search(r"(\d+)\s*/\s*(\d+)\s+tests collected", out)
    assert m, f"収集件数を読み取れない（pytest の出力形式が変わった）:\n{out[-800:]}"
    return int(m.group(1)), int(m.group(2))


def test_the_number_is_declared_exactly_once():
    """③ 宣言は 1 箇所だけ ── 同じ数を 2 箇所に書いた時点で、片方は必ず古くなる。"""
    found = _declared_all()
    assert len(found) == 1, (
        "`-m local` の本数の宣言が 1 箇所ではない（0 なら番人が守るものが無い・"
        f"2 以上なら片方が必ず古くなる）: {found}")
    assert found[0][0] == "docs/ENGINEERING.md", f"宣言の置き場所が変わった: {found}"


def test_the_declared_number_matches_what_pytest_actually_collects():
    """①②: 書いてある本数と、実際に選ばれる本数が一致すること。"""
    declared = _declared_all()[0][1]
    actual, _total = _collected()
    assert declared == actual, (
        f"docs/ENGINEERING.md は `pytest -m local` を {declared} 本と書いているが、"
        f"実測は {actual} 本 ── 文書を直すこと（この数は手で守れないから機械が見ている）")

# ── 主ファイルの行数（README が「割らなかった」理由として名指しする数） ──────────
LINES_MARK = re.compile(r"<!--\s*MAIN_FILE_LINES\s*-->\s*(\d+)\s*<!--\s*/MAIN_FILE_LINES\s*-->")
BUDGET = REPO / "tests" / "ailine_py_line_budget.txt"


def test_the_main_file_line_count_in_the_readme_is_not_stale():
    """★ 同じ形の嘘を先回りして塞ぐ: 行数は毎回変わるのに、README には固定で書いてある。
       分母は README 自身でも人の記憶でもなく、`test_line_budget` が縛っている
       tests/ailine_py_line_budget.txt から取る（そちらは実ファイルと一致を強制される）。"""
    readme = (REPO / "README.md").read_text(encoding="utf-8")
    found = LINES_MARK.findall(readme)
    assert len(found) == 1, f"主ファイルの行数の宣言が 1 箇所ではない: {found}"
    pinned = int(BUDGET.read_text(encoding="utf-8").splitlines()[0])
    assert int(found[0]) == pinned, (
        f"README は主ファイルを {found[0]} 行と書いているが、縛られている値は {pinned} 行")


# ── 試験の総数（README が評価者に「これだけ緑になります」と見せる数）────────────────
TOTAL_MARK = re.compile(r"<!--\s*TOTAL_TESTS\s*-->\s*(\d+)\s*<!--\s*/TOTAL_TESTS\s*-->")


def test_the_total_in_the_readme_matches_what_pytest_collects():
    """★ README は「この 2 つの数は機械が突き合わせている」と書いている ── 書いた以上、
       **両方**見ていないと、その一文自体が嘘になる。
       ★ 実際、初版は local の本数しか見ていないのに「2 つとも見ている」と書いていた
       （在るのに、その事故の形では鳴らない番人）。"""
    readme = (REPO / "README.md").read_text(encoding="utf-8")
    found = TOTAL_MARK.findall(readme)
    assert len(found) == 1, f"試験の総数の宣言が 1 箇所ではない: {found}"
    _local, total = _collected()
    assert int(found[0]) == total, (
        f"README は試験を {found[0]} 本と書いているが、実測は {total} 本")
