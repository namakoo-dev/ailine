# -*- coding: utf-8 -*-
"""③の到達 ── 出力を作った画面から検算に着けること（2026-09-13）。

## なぜ在るか

`ailine verify` は stack / extract / forms / split の出力を独立に検算できるのに、
**出力を作った画面には呼び方が一言も出ていなかった**（実測: 3 つの報告に `ailine verify` が 0 件）。
★ ① と同じ形で、しかも独立の検算を足した**当日**に同じ穴を開けた ──
**知られない検算は、無い検算と同じ**（到達 > 検出）。

契約:
  - 案内の文言は `cli_render.verify_hint` **1 本**（報告の側で書き写さない）
  - 案内を出す種類の宣言（`VERIFY_HINT_KINDS`）は、書き手の印の登録簿と**等号**で縛る
    （★ csv だけは独立の検算が無いので対象外と宣言する ── 案内すると嘘になる）
  - ★ **案内した呼び方が本当に通る**（そのまま実行して exit 0）── 文字列だけの案内を作らない
"""
from __future__ import annotations

import re
import subprocess
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
sys.path.insert(0, str(Path(__file__).resolve().parent))

import ailine   # noqa: E402
from ailine_core import cli_render, stack as stack_core   # noqa: E402
from test_forms_e2e import _invoice   # noqa: E402
from test_split_e2e import _book   # noqa: E402
# ★ 同名の器を 2 つ作らない（台帳の番人が居る）── 既にある方を借りる。
from test_folder_routes_hygiene import _registered_subcommands   # noqa: E402

#: 今回の器では案内しない印（★ 宣言 ── 増えたら等号の番人が鳴る）。**理由は別**なので分けて書く:
#:   ailine csv    独立の検算がそもそも無い（案内すると嘘になる）
#:   ailine match  検算は在るが**元が 2 冊**で引数の形が違う（この器は 元 1 つ）── 射程外
HINT_OUT_OF_SCOPE = {"ailine csv", "ailine match"}
_HINT = re.compile(r"あとから確かめる: ailine verify (\"[^\"]+\"|\S+) (\"[^\"]+\"|\S+)")


def _run(args, timeout: int = 300):
    return subprocess.run([sys.executable, "-m", "ailine", *args], capture_output=True,
                          text=True, timeout=timeout, encoding="utf-8", errors="replace",
                          cwd=str(REPO))


def _hint_args(stdout: str) -> list:
    """案内の行から 2 つのパスを取り出す（引用は外す）。"""
    m = _HINT.search(stdout)
    assert m, f"★ 案内の行が無い: {stdout[-600:]}"
    return [g.strip('"') for g in m.groups()]


def test_every_verifiable_kind_declares_a_hint():
    """★ 等号 ── 新しい出力の種類を足したら、案内を書くか「対象外」と宣言するまで赤くなる。"""
    declared = set(cli_render.VERIFY_HINT_KINDS) | HINT_OUT_OF_SCOPE
    assert declared == set(stack_core.CREATOR_MARKS), (
        f"宣言も対象外も無い印: {sorted(set(stack_core.CREATOR_MARKS) - declared)} / "
        f"登録簿に無い印: {sorted(declared - set(stack_core.CREATOR_MARKS))}")


def test_a_kind_without_an_independent_check_gets_no_hint():
    """★ 嘘の案内を作らない ── csv には独立の検算が無いので 1 行も出さない。"""
    assert cli_render.verify_hint("ailine csv", "out.xlsx", "元.csv") == []
    assert cli_render.verify_hint("ailine forms", "out.xlsx", "元") != []


def test_a_path_with_a_space_is_quoted():
    """★ そのまま貼って動く形 ── 空白を含むパスは引用する。"""
    line = cli_render.verify_hint("ailine forms", "C:/a b/一覧.xlsx", "C:/a b/元")[0]
    assert '"C:/a b/一覧.xlsx"' in line and '"C:/a b/元"' in line, line


def test_the_forms_report_offers_a_check_that_actually_runs(tmp_path):
    """★★ 案内をそのまま実行して exit 0（文字列だけの案内を作らない）。"""
    folder = tmp_path / "受け取った"
    folder.mkdir()
    _invoice(folder / "請求書1.xlsx", "取引先1株式会社", 11000)
    out = tmp_path / "一覧.xlsx"
    r = _run(["forms", str(folder), "--out", str(out)])
    assert r.returncode == 0, f"{r.stdout}\n{r.stderr}"
    args = _hint_args(r.stdout)
    assert _run(["verify", *args]).returncode == 0, args


def test_the_split_report_offers_a_check_that_actually_runs(tmp_path):
    """★ split は「出力はフォルダ・元は冊」だが並びは同じ ── 器を 1 本にできた理由。"""
    book = _book(tmp_path / "元" / "売上一覧.xlsx")
    out = tmp_path / "配る"
    r = _run(["split", str(book), "--by", "担当者", "--out", str(out)])
    assert r.returncode == 0, f"{r.stdout}\n{r.stderr}"
    args = _hint_args(r.stdout)
    assert _run(["verify", *args]).returncode == 0, args


def test_the_stack_report_offers_a_check_that_actually_runs(tmp_path):
    folder = tmp_path / "束"
    folder.mkdir()
    for i in (1, 2):
        _book(folder / f"売上{i}.xlsx")
    out = tmp_path / "縦積み.xlsx"
    r = _run(["stack", str(folder), "--out", str(out)])
    assert r.returncode == 0, f"{r.stdout}\n{r.stderr}"
    args = _hint_args(r.stdout)
    assert _run(["verify", *args]).returncode == 0, args


@pytest.mark.parametrize("kind", ["ailine forms", "ailine split", "ailine stack"])
def test_the_wording_lives_in_one_place(kind):
    """★ 報告の側に文言を書き写していないこと ── 書き写すと種類が増えた時に片配線になる。"""
    text = (REPO / "src" / "ailine_core" / "cli_render.py").read_text(encoding="utf-8")
    assert text.count("あとから確かめる: ailine verify") == 1, \
        "★ 案内の文言が 2 箇所以上にある（1 本に畳むこと）"
    assert cli_render.verify_hint(kind, "o", "s"), kind


def test_the_hint_carries_the_condition_the_person_typed(tmp_path):
    """★★ 勧める検算が、本人が実行した検算より**弱くならない**こと。

    `--amount 金額` を付けて分けたのに案内が付けないと、そのまま貼った人は
    「金額は測っていません」で終わる ── 勧める側が黙って手を抜くことになる。
    """
    book = _book(tmp_path / "元" / "売上一覧.xlsx")
    out = tmp_path / "配る"
    r = _run(["split", str(book), "--by", "担当者", "--out", str(out), "--amount", "金額"])
    assert r.returncode == 0, f"{r.stdout}\n{r.stderr}"
    assert "--amount" in r.stdout, f"★ 渡した条件が案内に出ていない: {r.stdout[-400:]}"
    m = _HINT.search(r.stdout)
    args = [g.strip('"') for g in m.groups()]
    tail = r.stdout[m.end():].splitlines()[0].strip()
    extra = [x.strip('"') for x in tail.split(" ", 1)] if tail else []
    got = _run(["verify", *args, *extra])
    assert got.returncode == 0, f"{got.stdout}\n{got.stderr}"
    assert "Σ元" in got.stdout, f"★ 金額を測っていない: {got.stdout}"


# --- 実機が無い PC でも「今できること」に着ける（2026-09-13・買い手の初回体験）------------
#
# ★★ なぜ在るか（買い手役を冷たい状態から通した実測）:
#   `ailine doctor` が LibreOffice と basrun に × を出して exit 1 を返し、その直後に
#   `ailine demo` が**動かない `ailine run`** を勧めていた。実際には forms / verify / scan は
#   読むだけなので LO も AI も要らないのに、**それを言う文が 1 つも無かった**。
#   買い手は「この PC では動かない」と結論して離脱する ── 最初の 5 分で一番大きい穴。
# ★ 名指しする入口は宣言 1 箇所（`NEEDS_MACHINE`）から導く。今朝の「手書きの白名簿」の轍を踏まない。


def test_every_subcommand_declares_whether_it_needs_the_machine():
    """★ 等号 ── 新しいサブコマンドは、実機が要るかを分類するまで赤くなる。"""
    declared, registered = set(ailine.NEEDS_MACHINE), _registered_subcommands()
    assert declared == registered, (
        f"分類されていない: {sorted(registered - declared)} / "
        f"登録簿に無い: {sorted(declared - registered)}")


def test_the_routes_named_when_the_machine_is_missing_really_work_without_it():
    """★ 嘘の案内を作らない ── 名指しする入口は、実機不要と**宣言されている**ものだけ。"""
    free = set(ailine.machine_free_routes())
    named = set(ailine.MACHINE_FREE_SHOWN)
    assert named <= free, f"実機が要る入口を『要らない』と案内している: {sorted(named - free)}"
    assert "run" not in free and "stop" not in free, "★ 実機が要る入口を要らない側に分類している"


def test_doctor_says_what_still_works_when_the_machine_is_missing():
    """★★ 「動かない」だけを言って去られない ── 何なら動くかを最後に必ず言う。"""
    broken = [("python 3.12+", True, ""), ("openpyxl", True, ""),
              ("LibreOffice", False, "見つかりません"), ("basrun.py", False, "ありません")]
    text, ok = ailine.format_doctor_report(broken)
    assert ok is False
    assert ailine.MACHINE_FREE_NOTE in text, text
    for name in ("forms", "verify"):
        assert f"ailine {name}" in text, f"{name} が案内に無い: {text}"


def test_doctor_stays_quiet_when_the_machine_is_there():
    """★ 陰性対照 ── 実機が在る PC に余計な 1 行を足さない。"""
    text, ok = ailine.format_doctor_report([("LibreOffice", True, ""), ("basrun.py", True, "")])
    assert ok is True
    assert ailine.MACHINE_FREE_NOTE not in text, text


def test_a_non_machine_failure_does_not_trigger_the_note():
    """★ 陰性対照 ── openpyxl が無いような**本当に動かない**回に「使えます」と言わない。"""
    text, _ok = ailine.format_doctor_report([("openpyxl", False, "入っていません")])
    assert ailine.MACHINE_FREE_NOTE not in text, text


def test_the_guidance_shows_every_required_option():
    """★★ 案内した形が**そのまま打てる**こと（2026-09-13・買い手の初回体験）。

    ★ 実測: README と `ailine ops` は `ailine forms <フォルダ>` と案内していたが、
      `--out` が必須なので**そのまま打つと exit 2 で落ちた**。stack / split / accounts も同じ。
      案内が動かないのは、断るより悪い。
    ★ 必須かどうかは argparse に言わせる（手で並べない）。
    """
    import argparse
    parser = ailine.build_parser()
    subs = [ac for ac in parser._actions if isinstance(ac, argparse._SubParsersAction)]
    out = _run(["ops"])
    assert out.returncode == 0, out.stdout
    for name in sorted(ailine.multi_file_routes()):
        sp = subs[0].choices.get(name)
        required = [ac.option_strings[-1] for ac in (sp._actions if sp else [])
                    if ac.option_strings and getattr(ac, "required", False)]
        for flag in required:
            assert flag in out.stdout, f"★ ops の案内に {name} の必須 {flag} が無い"


def test_the_readme_shows_every_required_option_too():
    """★ 買い手が最初に読むのは README ── そこに書いた形も打てること。"""
    import argparse
    text = (REPO / "README.md").read_text(encoding="utf-8")
    parser = ailine.build_parser()
    subs = [ac for ac in parser._actions if isinstance(ac, argparse._SubParsersAction)]
    checked = []
    for name in sorted(ailine.multi_file_routes()):
        sp = subs[0].choices.get(name)
        required = [ac.option_strings[-1] for ac in (sp._actions if sp else [])
                    if ac.option_strings and getattr(ac, "required", False)]
        if not required:
            continue
        line = next((ln for ln in text.splitlines() if f"`ailine {name} " in ln), None)
        assert line, f"★ README に {name} の行が無い"
        for flag in required:
            assert flag in line, f"★ README の {name} の行に必須 {flag} が無い: {line.strip()}"
        checked.append(name)
    # ★ 分母を先に確かめる ── 必須の指定を持つ入口が 0 件なら、上のループは
    #   **1 回も回らずに通る**（番人の番人に教わった・2026-09-13）。
    assert len(checked) >= 3, f"必須の指定を持つ入口が {checked} しか見つかっていない"
