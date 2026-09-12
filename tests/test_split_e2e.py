# -*- coding: utf-8 -*-
"""`ailine split <book> --by <見出し> --out <フォルダ>` の検体 ── 需要⑤の出口（2026-09-12）。

契約（設計 docs/DESIGN-20260912-担当者別に分けて配る.md）:
  - 値ごとに 1 冊（見出し行 ＋ その人の行 ＋ 出所列）＋ `_検分.xlsx`
  - 書き手の印（creator）と条件（description）が焼かれている ── 次に自分の出力と分かる
  - 合計・小計の行はどの冊にも入らない／空欄・表記ゆれ・複数担当は名指しされる
  - 担当者の見出しが 2 列に当たる冊は分けない（exit 4・1 冊も作らない）
  - 配る先に人のファイルがあれば exit 7（stack / forms と同じ関所）
  - ★★ 証明（D5）は**書いた冊を読み戻して**数える ── 読み戻しを差し替えたら exit 5 になり、
    1 冊も置かれない（＝自分の記憶で数えていないことの実証）
  - 原本は 1 バイトも変えない

★ LLM も LibreOffice も要らない（読むだけ・openpyxl のみ）。
"""
from __future__ import annotations

import argparse
import hashlib
import json
import subprocess
import sys
from pathlib import Path

import openpyxl
import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))

from ailine_core import split_people   # noqa: E402
from ailine_core import stack as stack_core   # noqa: E402

_HEADERS = ["日付", "案件", "顧客", "担当者", "金額"]
_ROWS = [
    ["2026-06-01", "広告", "甲社", "山田", 12000],
    ["2026-06-02", "印刷", "乙社", "佐藤", 8000],
    ["2026-06-03", "保守", "丙社", "山田", 25000],
    ["2026-06-04", "広告", "丁社", "山田　", 18000],   # ★ 表記ゆれ（全角空白）
    ["2026-06-05", "保守", "戊社", None, 7000],        # ★ 空欄
    ["2026-06-06", "印刷", "己社", "山田/佐藤", 5000],  # ★ 複数担当
    ["合計", None, None, None, 75000],                  # ★ 分けない行
]


def _book(path: Path, rows=None, headers=None, title=("2026年6月度 売上一覧",)):
    """見出しが 1 行目でない一覧表（実物の形）。★ 式は使わない（LO を起動しない）。"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "売上"
    for text in title:
        ws.append([text])
    ws.append(list(headers or _HEADERS))
    for row in (rows if rows is not None else _ROWS):
        ws.append(list(row))
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    wb.close()
    return path


def _split(book: Path, out: Path, *extra, by="担当者"):
    return subprocess.run(
        [sys.executable, "-m", "ailine", "split", str(book), "--by", by,
         "--out", str(out), *extra],
        capture_output=True, text=True, timeout=180, encoding="utf-8",
        errors="replace", cwd=str(REPO))


def _values(path: Path) -> list:
    wb = openpyxl.load_workbook(path)
    got = [list(r) for r in wb.worksheets[0].iter_rows(values_only=True)]
    wb.close()
    return got


def _marks(path: Path) -> tuple:
    wb = openpyxl.load_workbook(path)
    got = (wb.properties.creator, wb.properties.description)
    wb.close()
    return got


@pytest.fixture()
def book(tmp_path):
    return _book(tmp_path / "元" / "売上一覧.xlsx")


def test_one_book_per_person_with_the_provenance_columns(book, tmp_path):
    """★ 値ごとに 1 冊。見出し ＋ その人の行 ＋ 出所列（元ファイル・元行）。"""
    out = tmp_path / "配る"
    r = _split(book, out, "--amount", "金額")
    assert r.returncode == 0, f"落ちた: {r.stdout}\n{r.stderr}"
    made = sorted(p.name for p in out.glob("*.xlsx"))
    assert made == sorted(["山田.xlsx", "山田_.xlsx", "佐藤.xlsx", "_検分.xlsx"]), made
    head, *rows = _values(out / "山田.xlsx")
    assert head == _HEADERS + list(stack_core.PROVENANCE_HEADERS), head
    assert [row[-2:] for row in rows] == [["売上一覧.xlsx", 3], ["売上一覧.xlsx", 5]], rows
    assert [row[3] for row in rows] == ["山田", "山田"]


def test_the_writer_mark_and_the_condition_are_burned_in(book, tmp_path):
    """★ 印（creator）と条件（description）── 次に「自分の出力か」を機械で言えるように。"""
    out = tmp_path / "配る"
    assert _split(book, out).returncode == 0
    creator, description = _marks(out / "山田.xlsx")
    assert creator == split_people.CREATOR_MARK
    cond = json.loads(description)
    assert cond == {"tool": "ailine", "kind": "split", "by": "担当者", "value": "山田"}
    assert stack_core.own_output_mark(out / "山田.xlsx") == split_people.CREATOR_MARK
    assert stack_core.own_output_mark(out / "_検分.xlsx") == split_people.CREATOR_MARK, \
        "★ 検分の冊が自分の出力と認識されない（次回の関所が人のファイルと誤認する）"


def test_the_original_is_not_touched(book, tmp_path):
    before = hashlib.sha256(book.read_bytes()).hexdigest()
    _split(book, tmp_path / "配る", "--amount", "金額")
    assert hashlib.sha256(book.read_bytes()).hexdigest() == before, "★ 原本が変わった"


def test_a_total_row_never_lands_in_anyones_book(book, tmp_path):
    """★★ 最悪の混入（D2）── 合計行が誰かの冊に入っていないこと。"""
    out = tmp_path / "配る"
    assert _split(book, out, "--amount", "金額").returncode == 0
    for path in out.glob("*.xlsx"):
        if path.name == "_検分.xlsx":
            continue
        for row in _values(path)[1:]:
            assert "合計" not in [str(v) for v in row], f"{path.name} に合計行が混入: {row}"
    assert not (out / "合計.xlsx").exists()


def test_the_report_names_the_blank_the_lookalike_and_the_two_people(book, tmp_path):
    r = _split(book, tmp_path / "配る", "--amount", "金額")
    assert "空欄" in r.stdout and "表記ゆれ" in r.stdout and "複数担当" in r.stdout, r.stdout
    assert "✓ 部分の和 ＝ 全体" in r.stdout, f"証明を出していない: {r.stdout}"


def test_json_carries_the_whole_contract(book, tmp_path):
    """★ 機械可読の側にも同じ事実（採点器と自動化がここを読む）。"""
    r = _split(book, tmp_path / "配る", "--amount", "金額", "--json")
    payload = json.loads(r.stdout.strip().splitlines()[-1])
    for key in ("parts", "blank", "multi", "excluded", "lookalike", "refused",
                "proof", "unparsed"):
        assert key in payload, f"--json に {key} が無い: {sorted(payload)}"
    assert payload["refused"] is None
    assert payload["blank"] == [7] and [m[0] for m in payload["multi"]] == [8]
    assert payload["excluded"] == [9]
    assert sorted(sorted(p) for p in payload["lookalike"]) == [sorted(["山田", "山田　"])]
    rows = payload["proof"]["rows"]
    assert rows["whole"] == rows["parts"] + rows["blank"] + rows["excluded"] + rows["multi"]
    assert payload["proof"]["ok"] is True
    placed = sorted(r0 for part in payload["parts"].values() for r0 in part["rows"])
    assert placed == [3, 4, 5, 6], placed
    assert payload["parts"]["山田"]["amount"] == 37000


def test_the_inspection_book_carries_the_proof(book, tmp_path):
    out = tmp_path / "配る"
    assert _split(book, out, "--amount", "金額").returncode == 0
    head, *rows = _values(out / "_検分.xlsx")
    assert head == list(split_people.REPORT_HEADERS), head
    kinds = {str(row[0]) for row in rows}
    for kind in ("配った", "空欄", "複数担当", "分けない行", "表記ゆれ", "証明（行）",
                 "証明（金額）"):
        assert kind in kinds, f"検分に『{kind}』が無い: {sorted(kinds)}"


def test_two_matching_headers_refuse_and_write_nothing(tmp_path):
    """★★ 変異「見出しが 2 列でも分ける」── exit 4 で 1 冊も作らない（設計 D1）。"""
    book = _book(tmp_path / "元" / "実績.xlsx",
                 headers=["日付", "顧客", "担当", "営業担当", "金額"],
                 rows=[["2026-06-02", "甲社", "内藤", "山田", 30000],
                       ["2026-06-03", "乙社", "内藤", "佐藤", 18000]])
    out = tmp_path / "配る"
    r = _split(book, out, by="担当")
    assert r.returncode == 4, f"分けてしまった: {r.stdout}"
    assert "担当" in r.stdout and "営業担当" in r.stdout, r.stdout
    assert not out.exists() or not list(out.glob("*.xlsx")), "1 冊も作らないはず"


def test_a_missing_header_refuses_by_name(book, tmp_path):
    r = _split(book, tmp_path / "配る", by="記入者")
    assert r.returncode == 4
    assert "記入者" in r.stdout and "担当者" in r.stdout, r.stdout


def test_an_existing_human_file_stops_the_write(book, tmp_path):
    """★ 配る先に人のファイルがあれば exit 7（stack / forms と同じ関所）。"""
    out = tmp_path / "配る"
    out.mkdir()
    human = out / "大事な表.xlsx"
    wb = openpyxl.Workbook()
    wb.active["A1"] = "人が書いた大事な表"
    wb.save(human)
    wb.close()
    before = human.read_bytes()
    r = _split(book, out)
    assert r.returncode == 7, f"人のファイルの隣に黙って配った: {r.stdout}"
    assert human.read_bytes() == before
    assert "--overwrite" in r.stdout, "次の手を言っていない"
    assert not (out / "山田.xlsx").exists(), "止めたのに配っている"
    assert _split(book, out, "--overwrite").returncode == 0, "承知の上なら通ること"


def test_its_own_previous_output_is_not_a_gate(book, tmp_path):
    """★ 前回配った冊の上へは、関所なしで配り直せる（自分の出力だと分かるから）。"""
    out = tmp_path / "配る"
    assert _split(book, out).returncode == 0
    again = _split(book, out)
    assert again.returncode == 0, f"自分の前回出力を人のファイルと誤認した: {again.stdout}"


def test_the_other_sheets_are_disclosed(tmp_path):
    """★ 見たのは 1 枚目だけ ── 同じ冊に隠れている古い表を黙って無視しない。"""
    book = _book(tmp_path / "元" / "記録.xlsx")
    wb = openpyxl.load_workbook(book)
    ws2 = wb.create_sheet("売上_旧")
    ws2.append(_HEADERS)
    ws2.append(["2026-05-01", "広告", "甲社", "山田", 19000])
    ws2.sheet_state = "hidden"
    wb.save(book)
    wb.close()
    r = _split(book, tmp_path / "配る")
    assert "売上_旧" in r.stdout, f"他のシートを開示していない: {r.stdout}"


def test_a_named_sheet_that_is_missing_is_refused_not_replaced(book, tmp_path):
    r = _split(book, tmp_path / "配る", "--sheet", "存在しない")
    assert r.returncode == 4 and "売上" in r.stdout, r.stdout


def test_a_clean_table_raises_no_findings(tmp_path):
    """★ 陰性対照: 何も仕込まない表で所見 0（偽の疑いを立てたら道具の負け）。"""
    book = _book(tmp_path / "元" / "集計.xlsx",
                 rows=[["2026-05-01", "広告", "甲社", "山田", 20000],
                       ["2026-05-02", "印刷", "乙社", "佐藤", 9000],
                       ["2026-05-05", "保守", "丙社", "鈴木", 8500]])
    r = _split(book, tmp_path / "配る", "--amount", "金額", "--json")
    payload = json.loads(r.stdout.strip().splitlines()[-1])
    assert payload["blank"] == [] and payload["lookalike"] == []
    assert payload["multi"] == [] and payload["excluded"] == [] and payload["unparsed"] == []
    assert payload["proof"]["ok"] is True and len(payload["parts"]) == 3


def test_verify_does_not_pretend_the_mark_is_missing(book, tmp_path):
    """★ `ailine verify` は split の出力に「印がありません」と言わない（正直に断る）。

    ★ 2026-09-12（③）: 独立の検算が入ったので、断り文は「まだ実装していません」から
      **通る形の名指し**に変わった ── 1 冊では和が閉じないのでフォルダで受ける。
      詳しい検算そのものの契約は `tests/test_verify_split.py`。
    """
    out = tmp_path / "配る"
    assert _split(book, out).returncode == 0
    r = subprocess.run([sys.executable, "-m", "ailine", "verify", str(out / "山田.xlsx"),
                        str(book.parent)], capture_output=True, text=True, timeout=120,
                       encoding="utf-8", errors="replace", cwd=str(REPO))
    assert "印がありません" not in r.stdout, r.stdout
    assert "ailine verify <出力フォルダ> <元の冊>" in r.stdout, r.stdout


def test_the_proof_is_counted_from_the_written_books_not_from_memory(book, tmp_path,
                                                                     monkeypatch):
    """★★ 変異「証明を出力でなく自分の記憶で数える」がここで赤くなる。

    読み戻し（`xml_readback.read_grid`）の答えを 1 行だけ削って返すよう差し替える。
    記憶で数えているなら差し替えは効かず exit 0 のまま通る ── 出力から数えているなら
    等式が閉じず exit 5 になり、**1 冊も置かれない**。
    """
    import ailine

    real = ailine.xml_readback.read_grid

    def lying_read(path, sheet_name=None):
        data = real(path, sheet_name)
        rows = [r for (r, _c) in data["grid"] if r > 1]
        if rows:
            drop = max(rows)
            data["grid"] = {k: v for k, v in data["grid"].items() if k[0] != drop}
        return data

    monkeypatch.setattr(ailine.xml_readback, "read_grid", lying_read)
    out = tmp_path / "配る"
    args = argparse.Namespace(book=str(book), by="担当者", out=str(out), amount="金額",
                              sheet=None, overwrite=False, json=False)
    assert ailine.cmd_split(args) == 5, "★ 読み戻しが嘘をついても通った（記憶で数えている）"
    assert not out.exists() or not list(out.glob("*.xlsx")), \
        "★ 証明できていないのに配った"
