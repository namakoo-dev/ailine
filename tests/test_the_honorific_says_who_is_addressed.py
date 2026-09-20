# -*- coding: utf-8 -*-
"""「御中」は**組織にしか付かない** ── 屋号の宛先を発行元と取り違えない（2026-09-20）。

★★ 出所（盲検 4 体目 ②・記帳代行）: 請求書 5 通で**発行元が宛先の欄に入り、請求元が空**。

    元ファイル             | 請求元 | 宛先               | 請求額(税込)
    請求書_大東金属.xlsx   |       | 大東金属 株式会社   | 141900

  買い手:「請求元が全部空・宛先に仕入先名という表は、**そのままでは紙より悪い**
  （間違った情報が入っているので）。**この機能は無いのと同じ**です」

★★ 根（読んで確かめた）: 宛先のセルは `みどり商事 御中` で、法人格（株式会社…）が無い。
  `_looks_like_org` は法人格しか見ないので組織と認めず、「敬称のセルは部署か担当者 ──
  名前はその上」の規則が働いて**直上の発行元**を宛先にしていた。
  ★ **1 つの取り違えが 2 つの欄を壊す**（請求元は「宛先と同じ名前」として消される）。

★★ 直し: 「御中」は日本語の商習慣で**組織にしか付かない**（人には付かない）。だから
  御中 のセルは「そのセルが組織を指している」という証拠そのもの ── 法人格が書かれて
  いなくても屋号なら名前になる。ただし**下部組織**（総務課・経理部）は除く。

★★ 過去 2 回の失敗を繰り返さないための線（`docs/PENDING-20260918-…` に記録）:
  ・案 A「敬称が付けば名前」→ `総務課 御中` を宛先にして **225 冊が悪化**
  ・案 B「上へ辿ると請求元が残らないなら辿りを信じない」→ **16 冊が悪化**
  ★ どちらも **305 冊の基準線**が commit 前に止めた。この試験はその線を機械にする。
"""
from __future__ import annotations

import sys
import tempfile
from pathlib import Path

import openpyxl
import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
sys.path.insert(0, str(REPO / "tests"))

import forms_baseline_core as B  # noqa: E402
from ailine_core import form_read  # noqa: E402


def _invoice(d: Path, cells: dict, name: str = "請求書.xlsx") -> Path:
    """請求書らしい冊を 1 つ作る（印が 2 種類以上ないと読んでもらえない）。"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "請求書"
    base = {"A1": "請求書", "A7": "下記のとおりご請求申し上げます",
            "A9": "小計", "B9": 129000, "A10": "消費税", "B10": 12900,
            "A11": "合計", "B11": 141900,
            "A13": "請求日", "B13": "2026-09-18",
            "A14": "請求番号", "B14": "A-1024"}
    base.update(cells)
    for at, v in base.items():
        ws[at] = v
    p = d / name
    wb.save(p)
    wb.close()
    return p


def _read(cells: dict) -> dict:
    with tempfile.TemporaryDirectory() as td:
        return B.read_one(_invoice(Path(td), cells))


def test_the_buyers_invoice_is_read_correctly_now():
    """★★ 事故そのもの ── 屋号の宛先と、その上の発行元を取り違えないこと。

    ★ 1 つの取り違えが 2 つの欄を壊していたので、**両方**が戻ることまで見る。
    """
    got = _read({"A3": "大東金属 株式会社", "A5": "みどり商事 御中"})
    assert got["宛先"]["値"] == "みどり商事", got["宛先"]
    assert got["請求元"]["値"] == "大東金属 株式会社", got["請求元"]


def test_a_department_still_looks_above_for_the_name():
    """★★ 案 A が壊した所（225 冊）── 部署に「御中」が付いていても、名前はその上。

    ★ ここが緩むと『総務課』を宛先として出し、巻き添えで請求元まで消える。
    """
    got = _read({"A3": "ナギ商会株式会社", "A5": "総務課 御中"})
    assert got["宛先"]["値"] == "ナギ商会株式会社", got["宛先"]


@pytest.mark.parametrize("subunit", ["経理部", "購買部", "総務課", "第一営業所",
                                     "品質管理室", "東京支店", "配送センター"])
def test_every_kind_of_subunit_still_looks_above(subunit):
    """★ 下部組織の語尾は 1 つ残らず「上を見る」側に居ること。"""
    got = _read({"A3": "ナギ商会株式会社", "A5": f"{subunit} 御中"})
    assert got["宛先"]["値"] == "ナギ商会株式会社", (subunit, got["宛先"])


def test_a_person_with_sama_is_not_taken_as_the_organisation():
    """★★ 「様」は人に付く ── 新しい規則が走らないこと。

    ★ ここが走ると『佐藤 花子』が宛先の組織になる（旧版が正しくやっていた所）。
    """
    got = _read({"A4": "ナギ商会株式会社", "A6": "佐藤 花子 様"})
    assert got["宛先"]["値"] == "ナギ商会株式会社", got["宛先"]


def test_a_template_placeholder_is_still_refused():
    """★ 雛形のまま（〇〇 御中）は今も名前にしないこと（空欄は誤値より安い）。"""
    got = _read({"A3": "ナギ商会株式会社", "A5": "〇〇〇〇 御中"})
    assert got["宛先"]["値"] == "ナギ商会株式会社", got["宛先"]


def test_the_issuer_side_does_not_use_the_new_rule():
    """★★ 名前を作る道は 1 本のまま ── 請求元の側でこの規則が走らないこと。

    ★ 走ると、敬称の付いた宛先ブロックが請求元の候補になりうる。
    ★ 既定は「走らない」── 呼び出し側が明示した時だけ（型で守る）。
    """
    assert form_read.clean_org_name("みどり商事 御中")[0] == ""
    assert form_read.clean_org_name("みどり商事 御中", honorific_says_org=True)[0] == "みどり商事"


def test_the_frozen_baseline_of_305_books_does_not_move():
    """★★ 305 冊の基準線から**1 冊も動かない**こと。

    ★★ なぜこれを機械にするか: 2026-09-18 に 2 案を当てて、225 冊と 16 冊の退行を
      どちらもこの基準線が commit 前に止めた。止めたのは人ではなく**数**だった。
      その基準線は scratchpad に在って消えた ── だから repo の物として持ち直す。
    ★ 記録は人が明示的に取り直す（`python scripts/forms_baseline.py --write`）。
      走るたびに自動で揃えると、記録は常に一致して**二度と警告しなくなる**（恒真）。
    """
    before = B.load()
    assert before, ("★ 基準線が無い（`python scripts/forms_baseline.py --write` で取る）"
                    " ── 無ければこの検査は何も見ていない")
    assert len(before) >= 300, f"★ 基準線の冊が少なすぎる: {len(before)}"
    after = B.survey()
    moved = B.diff(before, after)
    assert not moved, "★ 読みが動いた冊がある:\n" + B.render(before, after)
