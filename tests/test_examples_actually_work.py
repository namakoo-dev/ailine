"""導線に出す例は、**そのまま打てば通る**こと（2026-09-05）。

★★ 実測した事故: 「整えて」と頼むと道具はこう返していた ──

    「整える」とは具体的に何をしますか（例: けい線を引く／列幅を合わせる／**太字にする**）

  そのまま「太字にして」と打つと **？ 対象『all』は 太字 では未対応です**。
  ★ **道具が自分で示した例を、自分で断っていた。**
  Namakoo:「間違った方向に誘導されると体験を損なう」── 導線が嘘なら、導線が無いより悪い。

★ 真因は 2 段:
  ① 例文が **few-shot に書かれた作文**で、LLM がそれを写して返していた
  ② 語彙表の `synonyms`（太字・ボールド・強調）も**そのままでは通らない** ──
     実測 10 件中 4 件が断られた。分かれ目は「対象が要る op かどうか」。
     synonyms は**呼び名**であって依頼文ではなかった。

★★ この試験が本体: **例を足した瞬間から縛られる。**
  ここが緑でなければ、その例は導線に出してはいけない。
  ★ 過去の判断（render_refusal の「『こう言えば通る』と書くと嘘になるので弱める」）を、
    弱めるのではなく**本当にする**側で解いた。
"""
import sys
from pathlib import Path

import openpyxl
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

import ailine  # noqa: E402
from ailine_core.examples import (  # noqa: E402
    DEFAULT_SUGGESTIONS, EXAMPLE_TASKS, example_task_for,
    render_example_line, render_examples_for, replace_examples_in_question)


def test_examples_exist_for_the_ops_we_suggest_by_default():
    """★ 曖昧な依頼で真っ先に見せる op は、必ず例を持つこと。"""
    for op in DEFAULT_SUGGESTIONS:
        assert example_task_for(op), f"{op} に例が無いのに提案の先頭に置いている"


def test_every_example_names_a_real_op():
    """★ 存在しない op に例を書かない。"""
    for op in EXAMPLE_TASKS:
        assert op in ailine.OP_SCHEMA, f"知らない op: {op}"


def test_an_op_without_an_example_stays_silent():
    """★ 無い例を発明しない（黙る）。"""
    assert example_task_for("CHART") is None or "CHART" in EXAMPLE_TASKS
    assert example_task_for("NOPE") is None
    assert render_example_line("NOPE") is None


def test_the_question_keeps_its_wording_and_only_swaps_the_examples():
    """★ 聞き返しの主文は LLM の方が場面に合う ── **例の括弧だけ**を差し替える。"""
    q = "「整える」とは具体的に何をしますか（例: けい線を引く／列幅を合わせる／太字にする）"
    got = replace_examples_in_question(q)
    assert got.startswith("「整える」とは具体的に何をしますか"), got
    assert "太字にする）" not in got, got
    assert "見出しを太字にして" in got, got


def test_a_question_without_examples_is_untouched():
    got = replace_examples_in_question("どの列が同じなら重複とみなしますか")
    assert got == "どの列が同じなら重複とみなしますか"


def test_when_no_example_exists_the_parenthesis_is_dropped():
    """★ 通らない例を残すより、例が無い方がまし。"""
    got = replace_examples_in_question("何をしますか（例: なにか）", ops=["NOPE"])
    assert "（例" not in got, got
    assert got == "何をしますか"


def test_examples_are_not_re_listed_in_the_caller():
    """★ 文面は 1 箇所に置く（断り・聞き返し・提案で同じ言い方にする）。"""
    src = Path(ailine.__file__).read_text(encoding="utf-8")
    assert "けい線を引いて" not in src, "呼び出し側に例文を書き写している"
    assert "render_example_line(" in src and "replace_examples_in_question(" in src


@pytest.mark.local
@pytest.mark.parametrize("op, task", sorted(EXAMPLE_TASKS.items()))
def test_the_example_actually_runs(tmp_path, op, task):
    """★★ 本体 ── 例をそのまま打って**通る**こと。実機で毎回確かめる。

    ★ 例を足したら、この試験が自動で回る（列挙を手で増やさない）。
    """
    import os, subprocess
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "表"
    ws.append(["品名", "部門", "金額", "備考"])
    for r in [["机", "営業", 12000, None], ["椅子", "経理", 800, None],
              ["棚", "営業", 15000, None]]:
        ws.append(r)
    src = tmp_path / "in.xlsx"; wb.save(src)
    repo = Path(ailine.__file__).resolve().parents[2]
    p = subprocess.run([sys.executable, "-m", "ailine", "run", str(src), task,
                        "--copy", "--sheet", "表", "--timeout", "150"],
                       capture_output=True, text=True, encoding="utf-8", errors="replace",
                       cwd=str(repo), env={**os.environ, "PYTHONPATH": str(repo / "src")})
    assert p.returncode == 0, (
        f"導線に出している例が通らない: {op} 「{task}」\n" + p.stdout[-500:])


@pytest.mark.local
def test_the_examples_shown_for_a_vague_request_can_be_typed_back(tmp_path):
    """★ 曖昧な依頼 → 示された例 → そのまま打つ、の一往復が成立すること。"""
    import os, re, subprocess
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "在庫"
    ws.append(["品名", "棚", "数量", "備考"])
    ws.append(["ボルト", "A-1", 120, None])
    src = tmp_path / "in.xlsx"; wb.save(src)
    repo = Path(ailine.__file__).resolve().parents[2]
    env = {**os.environ, "PYTHONPATH": str(repo / "src")}

    def _run(task, book):
        return subprocess.run([sys.executable, "-m", "ailine", "run", str(book), task,
                               "--copy", "--sheet", "在庫", "--timeout", "150"],
                              capture_output=True, text=True, encoding="utf-8",
                              errors="replace", cwd=str(repo), env=env)

    first = _run("整えて", src)
    assert first.returncode == 3, first.stdout[-400:]
    shown = re.findall(r"「([^」]{5,})」", first.stdout)
    assert shown, f"例が 1 つも出ていない: {first.stdout[-400:]}"
    for example in shown[:3]:
        again = _run(example, src)
        assert again.returncode == 0, (
            f"示した例が通らない: 「{example}」\n" + again.stdout[-400:])
