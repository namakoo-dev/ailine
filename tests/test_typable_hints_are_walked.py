"""「こう打てば進める」と言った所は、その道が実際に通る（2026-09-23・導線の台帳を歩く）。

★★ なぜ在るか（盲検 6 体目 ⑥・致命）: 道具が `--column …` を案内し、**その通りに打つと落ちた**。
  断りの道を歩く道具（walk_refusals_core）は既に在ったのに、導線の台帳（40 件）には歩き方が
  1 件も無く、`walked` は人が手で書く欄だった（38 件が `未調査`）。
★ ここは台帳の `walked` を**歩いた結果で縛る**。判定は台帳に書かず、歩いて決める。
"""
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent))
import walk_refusals_core as walk  # noqa: E402

#: 台帳の walked 欄に書いてよい値のうち、機械が歩いて決めるもの
MACHINE_DECIDES = {"walked", "path_fails", "by_design"}


@pytest.fixture(scope="module")
def rows():
    """★ 1 回だけ歩く（40 件 × 引き金と道）── 3 本の試験で書き写さない。"""
    return {r["key"]: r for r in walk.survey_hints()}


def test_no_hint_shows_a_path_that_does_not_work(rows):
    """★★ 本体: 案内どおりに打って落ちる導線が無い（2 回歩いて再現したものだけ数える）。"""
    first = [r for r in rows.values() if r["verdict"] == "path_fails"]
    if not first:
        return
    again = {r["key"] for r in walk.survey_hints() if r["verdict"] == "path_fails"}
    bad = [r for r in first if r["key"] in again]
    assert not bad, "案内どおりに打つと落ちる導線:\n" + "\n".join(
        f"  {r['key']}  {r['detail']}\n      画面: {r.get('screen', '')[:160]}" for r in bad)


def test_every_written_walk_actually_fires(rows):
    """歩き方を書いた項目は、どれも案内を実際に出せる（出せないなら歩き方が古い）。"""
    stale = [r for r in rows.values() if r["verdict"] in ("引き金が引けない", "vague")]
    assert not stale, "歩き方が案内を出せていない:\n" + "\n".join(
        f"  {r['key']}  {r['detail']}" for r in stale)


def test_the_register_says_what_the_machine_saw(rows):
    """★ 台帳の walked は歩いた結果と一致する。歩き方の無い項目は `未調査` 以外を名乗らない。"""
    wrong = []
    for h in walk.load_hints():
        key = walk.hint_key(h)
        seen = rows[key]["verdict"]
        if not h.get("walk"):
            if h.get("walked") != "未調査":
                wrong.append(f"  {key}  台帳は {h.get('walked')}・歩き方が無いのに（未調査のはず）")
        elif seen in MACHINE_DECIDES and h.get("walked") != seen:
            wrong.append(f"  {key}  台帳は {h.get('walked')}・歩いたら {seen}")
    assert not wrong, "台帳と歩いた結果が食い違う:\n" + "\n".join(wrong)


@pytest.mark.local
def test_typable_hints_walk_on_the_machine(monkeypatch):
    """★ 機械の状態に左右される歩き（machine: true）を、実機で実際に歩く（素の環境では歩かない分）。
       ── 素の環境でも実機でも見ていない、という穴を作らない。"""
    monkeypatch.setenv("AILINE_WALK_ON_MACHINE", "1")
    hints = [h for h in walk.load_hints() if (h.get("walk") or {}).get("machine")]
    assert hints, "machine の歩きが 0 件 ── 印が外れていないか"
    import tempfile
    bad = []
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
        for i, h in enumerate(hints):
            r = walk.walk_hint(h, Path(td) / f"m{i:02d}")
            if r["verdict"] != h.get("walked"):
                bad.append(f"  {r['key'][:70]}  台帳は {h.get('walked')}・歩いたら {r['verdict']}（{r['detail'][:80]}）")
    assert not bad, "実機で歩いた結果が台帳と食い違う:\n" + "\n".join(bad)
