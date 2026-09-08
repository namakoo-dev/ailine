# -*- coding: utf-8 -*-
"""決めたこと・保留していることの文書が、索引から**指されている**か（2026-09-09）。

★★ なぜ在るか: `docs/DECIDED-20260907-value-without-quotes.md` は、置かれてから
  2 日間**どこからも指されていなかった**（grep で 0 件）。置いただけの文書は
  未来の誰も拾わない ── この repo が何度も踏んだ「番人不在の契約」と同じ形。

★ 両向きに縛る（片側だけだと恒真になる）:
    実体に在って索引に無い  → 増やしたのに載せ忘れた
    索引に在って実体が無い  → 消したのに索引が残った
"""
from __future__ import annotations

import re
from pathlib import Path

DOCS = Path(__file__).resolve().parent.parent / "docs"
INDEX = DOCS / "ENGINEERING.md"


def _indexed() -> set:
    text = INDEX.read_text(encoding="utf-8")
    return set(re.findall(r"\[([A-Z]+-\d{8}[^\]]*\.md)\]", text))


def _on_disk() -> set:
    return {p.name for p in DOCS.glob("*.md")
            if re.match(r"^(DECIDED|PENDING)-\d{8}", p.name)}


def test_every_decision_doc_is_pointed_at():
    disk, idx = _on_disk(), _indexed()
    assert disk, "決裁/保留の文書が 1 つも無い（番人が守る対象を失っている）"
    assert not (disk - idx), (
        f"docs/ に在るのに索引から指されていない: {sorted(disk - idx)} ── "
        "docs/ENGINEERING.md の『決めたこと・保留していることの索引』に足すこと")


def test_the_index_does_not_point_at_a_ghost():
    disk, idx = _on_disk(), _indexed()
    ghosts = {n for n in idx if n.startswith(("DECIDED-", "PENDING-"))} - disk
    assert not ghosts, f"索引が指す先が docs/ に無い: {sorted(ghosts)}"
