"""盲検の記録は、**数えられる形**で書かれていること（2026-09-22）。

★★ なぜ在るか: 1〜5 体目を数え直したら、回ごとに様式が違って**比べられなかった**。
  4 体目が 8/8 再現、5 体目が 5/5 に見えるのは、様式が変わったせいかもしれないし
  実際に減ったのかもしれない ── **区別がつかない**。
  ★ 決めずに走らせると、比べられない記録がもう 1 つ増える。だから先に固定した。

★ この番人が縛るのは **docs/盲検の記録様式.md を決めた日以降**の記録だけ。
  ★ 過去 4 回は遡って直さない ── 記録を書き換えたら、それは記録でなくなる。

★ いまは対象が 0 件（6 体目がまだ）。**0 件でも緑になる番人**は「在っても鳴らない」
  形なので、下に「様式そのものが在ること」を縛る試験を置いてある。
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from blind_rate_core import (KINDS, REPRO, SCHEMA_FROM, parse_findings,  # noqa: E402
                             records_under_the_schema, rate, required_fields)

SCHEMA_DOC = Path(__file__).resolve().parent.parent / "docs" / "盲検の記録様式.md"


def test_the_schema_document_exists_and_names_the_three_kinds():
    """★ 様式そのものが在ること（対象が 0 件でも、この番人は何かを守る）。"""
    assert SCHEMA_DOC.exists(), f"{SCHEMA_DOC.name} が無い"
    text = SCHEMA_DOC.read_text(encoding="utf-8")
    for k in KINDS:
        assert f"`{k}`" in text, f"様式に区分『{k}』の説明が無い"
    for r in REPRO:
        assert f"`{r}`" in text, f"様式に再現の値『{r}』の説明が無い"
    assert "役柄" in text, "★ 役柄を固定するか比べないかの線が様式に無い"


def test_every_new_record_parses_into_the_fixed_columns():
    """★★ 本体 ── 様式より後の記録は、5 欄の表として読めること。"""
    for p in records_under_the_schema():
        findings = parse_findings(p)
        assert findings, (
            f"{p.name} が様式どおりに読めません ── "
            "docs/盲検の記録様式.md の 5 欄（# / 区分 / 所見 / 再現 / 致命）で書いてください")
        for f in findings:
            assert f["kind"] in KINDS, (p.name, f)
            assert f["repro"] in REPRO, (p.name, f)
            assert f["fatal"] in ("致命", "─", "-", ""), (p.name, f)


def test_every_new_record_states_its_denominator_and_role():
    """★ 数字だけ残さない ── 分母（緑だった試験数）と役柄と渡したものを一緒に書く。

    ★ 役柄が毎回違えば、数字は製品でなく**役柄**を測ってしまう。
    """
    for p in records_under_the_schema():
        got = required_fields(p)
        missing = [k for k, ok in got.items() if not ok]
        assert not missing, f"{p.name} に {missing} が書かれていません"


def test_the_rate_keeps_the_false_report_rate_beside_it():
    """★ 誤報率を一緒に返すこと ── 外の目も測定器で、持たないと恒真になる。

    ★ 実測（2026-09-22・過去 4 回）: 所見 31 件のうち 7 件（23%）が再現しなかった。
    """
    sample = [{"mark": "①", "kind": "欠陥", "text": "x", "repro": "再現", "fatal": "致命"},
              {"mark": "②", "kind": "欠陥", "text": "y", "repro": "未再現", "fatal": "─"},
              {"mark": "③", "kind": "摩擦", "text": "z", "repro": "引用で明白", "fatal": "─"}]
    r = rate(sample)
    assert r["欠陥"] == 2 and r["再現した欠陥"] == 1
    assert r["誤報率"] is not None and abs(r["誤報率"] - 1 / 3) < 1e-9
    assert r["致命"] == 1


def test_the_schema_date_is_not_moved_backwards():
    """★ 固定した日を過去へ動かすと、書き換えたくない記録まで縛られる。"""
    assert SCHEMA_FROM == "20260922", (
        "様式を固定した日を動かしています ── 過去の記録を遡って様式に合わせないこと")
