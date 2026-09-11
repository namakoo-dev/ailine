# -*- coding: utf-8 -*-
"""`ailine forms <folder> --out <path>` の検体 ── 需要①の出口（2026-09-11）。

契約:
  - 帳票の山を読み、**1 冊 1 行**の一覧を新しいブックに書く（原本は読むだけ）
  - 確/単 は値を書き、**割/無 は空欄**（2026-09-11 に凍結した判断）
  - 空欄は**必ず**検分シートに理由つきで 1 行出る（空欄の数 == 理由の数）
  - 請求書でないものからは値を作らない（官公庁の実表・空の雛形）
  - 出力先に人のファイルがあれば exit 7 で止まる（stack と同じ関所）
  - 自分の前回出力は入力に数えない（二重計上を防ぐ）

★ LLM も LibreOffice も要らない（読むだけ・openpyxl のみ）。
"""
from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

import openpyxl
import pytest

REPO = Path(__file__).resolve().parent.parent


def _invoice(path: Path, issuer: str, total: int, *, addressee: str = "ナギ商会株式会社"):
    """最小の請求書（帳票の形）。見出しの位置も帯も実物の形に合わせる。"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "請求書"
    ws["B2"] = "請求書"
    ws["B3"] = addressee
    ws["B4"] = "経理部　御中"
    ws["G5"] = issuer
    ws["G6"] = "〒000-0000"
    ws["G3"] = "請求日："
    ws["H3"] = "2026/8/31"
    ws["G4"] = "請求番号："
    ws["H4"] = f"INV-{total}"
    ws["B11"] = " ご請求金額　"
    ws["C11"] = total
    ws["B15"] = "品番・品名"
    ws["E15"] = "数量"
    ws["G15"] = "単価"
    ws["H15"] = "金額"
    ws["B16"] = "用紙代"
    ws["E16"] = 1
    ws["G16"] = total // 11 * 10
    ws["H16"] = total // 11 * 10
    ws["E37"] = "小計"
    ws["H37"] = total // 11 * 10
    ws["E38"] = "消費税"
    ws["G38"] = 0.1
    ws["H38"] = total - total // 11 * 10
    ws["E39"] = "合計金額"
    ws["H39"] = total
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    wb.close()


def _forms(folder: Path, out: Path, *extra):
    return subprocess.run(
        [sys.executable, "-m", "ailine", "forms", str(folder), "--out", str(out), *extra],
        capture_output=True, text=True, timeout=180, encoding="utf-8",
        errors="replace", cwd=str(REPO))


def _sheets(out: Path) -> dict:
    wb = openpyxl.load_workbook(out)
    got = {ws.title: [list(r) for r in ws.iter_rows(values_only=True)] for ws in wb.worksheets}
    got["_creator"] = wb.properties.creator
    wb.close()
    return got


@pytest.fixture()
def folder(tmp_path):
    d = tmp_path / "受領"
    _invoice(d / "a.xlsx", "あかね商事株式会社", 3300)
    _invoice(d / "b.xlsx", "いろは工業株式会社", 5500)
    return d


def test_one_row_per_invoice_with_the_fields_filled(folder, tmp_path):
    """★ 帳票の山 → 1 冊 1 行の一覧。"""
    out = tmp_path / "一覧.xlsx"
    r = _forms(folder, out)
    assert r.returncode == 0, f"落ちた: {r.stdout}\n{r.stderr}"
    got = _sheets(out)
    assert got["_creator"] == "ailine forms", "★ 書き手の印が無い（次回 自分の出力と分からない）"
    head, *rows = got["一覧"]
    assert head == ["元ファイル", "請求元", "宛先", "請求額(税込)", "請求日", "請求番号"], head
    assert len(rows) == 2, rows
    by_file = {r0[0]: r0 for r0 in rows}
    assert by_file["a.xlsx"][1] == "あかね商事株式会社"
    assert by_file["a.xlsx"][2] == "ナギ商会株式会社"
    assert by_file["a.xlsx"][3] == 3300
    assert by_file["b.xlsx"][3] == 5500


def test_the_originals_are_not_touched(folder, tmp_path):
    """★ 原本は 1 バイトも変わらない（読むだけの契約）。"""
    import hashlib
    before = {p.name: hashlib.sha256(p.read_bytes()).hexdigest() for p in folder.glob("*.xlsx")}
    _forms(folder, tmp_path / "一覧.xlsx")
    after = {p.name: hashlib.sha256(p.read_bytes()).hexdigest() for p in folder.glob("*.xlsx")}
    assert before == after, "★ 原本が変わった"


def test_a_book_that_is_not_an_invoice_yields_no_values(folder, tmp_path):
    """★★ 請求書でないものから値を作らない（G2' の必達『混入で沈黙』）。

    ★ 官公庁の統計表を混ぜても、その行は 3 項目とも空欄になること。
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "家計調査"
    ws["A1"] = "家計調査"
    ws["A3"] = "世帯数"
    ws["B3"] = 1234
    ws["A4"] = "金額"
    ws["B4"] = 5678
    wb.save(folder / "統計.xlsx")
    wb.close()

    out = tmp_path / "一覧.xlsx"
    assert _forms(folder, out).returncode == 0
    got = _sheets(out)
    row = next(r for r in got["一覧"][1:] if r[0] == "統計.xlsx")
    assert row[1:] == [None] * 5, f"★ 請求書でない冊から値を作った: {row}"


def test_every_blank_has_a_reason_in_the_inspection_sheet(folder, tmp_path):
    """★★ 空欄の数 == 理由の数（凍結した判断・型でも守っているが、ここでも数える）。

    ★「型が守っているはず」は検算ではない ── 出力を読んで数える。
    """
    wb = openpyxl.Workbook()
    wb.active["A1"] = "何かのメモ"
    wb.save(folder / "メモ.xlsx")
    wb.close()

    out = tmp_path / "一覧.xlsx"
    r = _forms(folder, out)
    assert r.returncode == 0, r.stdout
    got = _sheets(out)
    blanks = sum(1 for row in got["一覧"][1:] for v in row[1:] if v is None)
    inspect = {(x[0], x[1]) for x in got["検分"][1:]}
    for row in got["一覧"][1:]:
        for i, v in enumerate(row[1:]):
            if v is None:
                field = ["請求元", "宛先", "請求額", "請求日", "請求番号"][i]
                assert (row[0], field) in inspect, f"★ 空欄に理由が無い: {row[0]}/{field}"
    assert blanks >= 3, f"★ 空欄が出ていない（検体が弱い）: {blanks}"
    reasons = [x[3] for x in got["検分"][1:] if (x[0], x[1]) in inspect]
    assert all(str(t).strip() for t in reasons), "★ 理由が空の行がある"


def test_the_report_names_the_denominator_and_the_blanks(folder, tmp_path):
    """★ 買い手の信用条件: 分母・区分の内訳・空欄の数を人に見せる。"""
    r = _forms(folder, tmp_path / "一覧.xlsx")
    assert "2 ファイル中 2 冊を読みました" in r.stdout, r.stdout
    assert "項目の区分" in r.stdout, r.stdout


def test_an_existing_human_file_stops_the_write(folder, tmp_path):
    """★ 出力先に人のファイルがあれば exit 7（stack と同じ関所）。"""
    out = tmp_path / "一覧.xlsx"
    wb = openpyxl.Workbook()
    wb.active["A1"] = "人が書いた大事な表"
    wb.save(out)
    wb.close()
    before = out.read_bytes()
    r = _forms(folder, out)
    assert r.returncode == 7, f"★ 人のファイルを黙って上書きした: {r.stdout}"
    assert out.read_bytes() == before, "★ 原本が変わった"
    assert "--overwrite" in r.stdout, "★ 次の手を言っていない"


def test_its_own_previous_output_is_not_counted_as_an_input(folder, tmp_path):
    """★ 前回の一覧を同じフォルダに置いても、入力に数えない（二重計上を防ぐ）。"""
    out = folder / "一覧.xlsx"
    assert _forms(folder, out).returncode == 0
    r = _forms(folder, out)
    assert r.returncode == 0, r.stdout
    assert "2 ファイル中 2 冊を読みました" in r.stdout, r.stdout
    assert "自分の出力" in r.stdout, r.stdout


def test_json_carries_the_denominator_and_the_blank_accounting(folder, tmp_path):
    """★ 機械可読の側にも分母と空欄の会計を載せる（人向けと同じ事実）。"""
    r = _forms(folder, tmp_path / "一覧.xlsx", "--json")
    payload = json.loads(r.stdout.strip().splitlines()[-1])
    assert payload["denominator"] == 2 and payload["collected"] == 2
    assert payload["blanks"] == payload["blanks_with_reason"], payload
    assert payload["grades"], payload


def test_the_inspection_sheet_does_not_cry_wolf(folder, tmp_path):
    """★★ 検分に載るのは**人が手を動かすものだけ**（空欄）── 普通の状態を並べない。

    ★ 実測（検体 87 冊）: 初版は「単（裏が取れていない）」も 1 行ずつ出していて、
      単 201 行が 割 12 + 無 15 の **27 行を埋めていた**。
      単 は口が 1 つしかない普通の状態で、それを所見として並べるのは
      「オオカミ少年防止を謳う道具が自分でオオカミ少年になる」形。
    """
    out = tmp_path / "一覧.xlsx"
    assert _forms(folder, out).returncode == 0
    got = _sheets(out)
    grades = {x[2] for x in got["検分"][1:]}
    assert "単" not in grades, f"★ 普通の状態（単）を検分に並べている: {got['検分'][1:3]}"
    assert "確" not in grades, "★ 裏が取れたものまで検分に並べている"


def test_the_grades_are_still_reachable_for_automation(folder, tmp_path):
    """★ 検分から外した『単/確』は **消えていない** ── --json に 1 冊 1 項目で残る。

    ★ 画面から消すことと、機械可読から消すことは別 ── 出す先を分けただけにする。
    """
    r = _forms(folder, tmp_path / "一覧.xlsx", "--json")
    payload = json.loads(r.stdout.strip().splitlines()[-1])
    fg = payload["field_grades"]
    assert len(fg) == 2 * 5, fg          # 2 冊 × 5 項目
    assert {x["field"] for x in fg} == {"請求元", "宛先", "請求額", "請求日", "請求番号"}
    assert all(x["grade"] in ("確", "単", "割", "無") for x in fg), fg


def test_a_missing_record_counts_as_a_blank_without_a_reason():
    """★ 負の被覆: 記録そのものが無い項目は『理由の無い空欄』として数えられること。

    ★ 2026-09-11 に実際に生まれた形 ── read_book の一枝が自前の 3 つ組を持っていて、
      請求日・請求番号を足したとき記録が無い空欄ができた。型（Record）は記録が無ければ
      働かない。初版の blanks_have_reasons はこれを continue で飛ばしていて数えなかった。
    """
    sys.path.insert(0, str(REPO / "src"))
    from ailine_core import forms_collect
    blanks, with_reason, bad = forms_collect.blanks_have_reasons([("x.xlsx", {})])
    assert blanks == len(forms_collect.FIELDS) and with_reason == 0
    assert bad and all("記録が無い" in b for b in bad), bad


def test_verify_tells_the_truth_about_its_own_forms_output(folder, tmp_path):
    """★ 穴 A（2026-09-11）: `ailine verify` は forms の出力に自分の印を見つけながら
    「印がありません」と言っていた。独立の検算が無いのは事実なので、csv と同じく
    unsupported（対応外）で正直に返す ── unmarked（他人のファイル）に混ぜない。"""
    out = tmp_path / "一覧.xlsx"
    assert _forms(folder, out).returncode == 0
    from ailine_core import verify as multifile_verify
    got = multifile_verify.verify_output(out, folder)
    assert not got.get("unmarked"), "★ 自分の印を『印が無い』と言っている"
    assert "実装していません" in (got.get("unsupported") or ""), got
    r = subprocess.run([sys.executable, "-m", "ailine", "verify", str(out), str(folder)],
                       capture_output=True, text=True, timeout=180, encoding="utf-8",
                       errors="replace", cwd=str(REPO))
    assert r.returncode == 4, r.stdout
    assert "印がありません" not in r.stdout, r.stdout
