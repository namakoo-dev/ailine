# 翻訳精度の数字を、文書と実測で縛る ── 2026-08-29。
#
# ★★ 見つかった食い違い: README は「op 一致 51/52 = 98.1%」と書いていたが、
#   同じ凍結検体を今日 2 回走らせると **49/52 = 94.2%** だった。
#   OPS_DOC に 16 行足した回の低下（98.1%→94.2%・当時 実測済み）が、
#   **文書に反映されないまま残っていた**。
#   ★ これは面接で口に出す数字だ。手で守れない数字は機械が守る
#     （試験の本数・主ファイルの行数と同じ扱いにする）。
#
# 二段構え:
#   ① 非 local: 文書の数字と**記録**（tests/battery_recorded.json）が一致すること
#   ② local:    記録と**実測**が一致すること（実物の LLM が要るので実機側）

import json
import re
import subprocess
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
RECORD = Path(__file__).resolve().parent / "battery_recorded.json"


def _record():
    return json.loads(RECORD.read_text(encoding="utf-8"))


def test_the_readme_matches_the_record():
    """① 文書 vs 記録。★ 数字を直すときは記録も直す（片方だけ動かせない形にする）。"""
    rec = _record()
    want = f"{rec['op_correct']}/{rec['op_total']} = {rec['op_correct'] / rec['op_total'] * 100:.1f}%"
    text = (REPO / "README.md").read_text(encoding="utf-8")
    m = re.search(r"<!-- BATTERY_OP -->(.*?)<!-- /BATTERY_OP -->", text)
    assert m, "README に BATTERY_OP の印が無い"
    assert m.group(1) == want, f"README は {m.group(1)!r}・記録は {want!r}"


def test_the_runbook_does_not_quote_a_stale_number():
    """★ 手順書は当日そのまま読み上げる紙 ── 古い数字が残っていたら赤くする。"""
    rec = _record()
    pct = f"{rec['op_correct'] / rec['op_total'] * 100:.1f}%"
    text = (REPO / "demo" / "手順.md").read_text(encoding="utf-8")
    nums = set(re.findall(r"op 一致 ([0-9]+\.[0-9]%)", text))
    assert nums <= {pct}, f"手順書に古い数字が残っている: {nums - {pct}}（今は {pct}）"


def test_the_record_names_how_it_was_measured():
    """★ 数字だけを残さない ── いつ・どのモデルで・何回で出たかを一緒に置く
       （後から誰が見ても、同じ条件で測り直せる形にしておく）。"""
    rec = _record()
    for key in ("measured_on", "model", "runs", "misassert", "misassert_total"):
        assert rec.get(key) not in (None, ""), key
    assert rec["runs"] >= 2, "1 回の測定を記録にしない（LLM は揺れる）"


@pytest.mark.local
def test_the_record_still_matches_the_machine():
    """② 記録 vs 実測。実物の ollama が要るので実機側。
       ★ ここが赤くなったら、直すのは**記録と文書**（実測が正）。"""
    r = subprocess.run(
        [sys.executable, str(REPO / "bench" / "translation_dsl_battery_run.py"),
         _record()["model"]],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
        timeout=2400, cwd=str(REPO / "bench"),
        env={**__import__("os").environ, "PYTHONPATH": str(REPO / "src")})
    assert r.returncode == 0, r.stdout[-600:]
    m = re.search(r"op 分類: (\d+)/(\d+)", r.stdout)
    assert m, r.stdout[-600:]
    rec = _record()
    got, total = int(m.group(1)), int(m.group(2))
    assert total == rec["op_total"], f"検体の数が変わった（凍結のはず）: {total}"
    # ★ LLM は揺れる ── 1 件のぶれは許すが、それ以上ずれたら記録を測り直す
    assert abs(got - rec["op_correct"]) <= 1, (
        f"実測 {got}/{total}・記録 {rec['op_correct']}/{total} ── "
        "記録と文書を測り直して直すこと（実測が正）")
