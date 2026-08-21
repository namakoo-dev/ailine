"""M2 `ailine run <フォルダ> "<依頼>"`（抽出集約）の検体。
   ★ 実装前に凍結した赤い検体（DESIGN-20260821-multifile M2 節・E3 が筆頭）。

   ★ 7B を使わない: 翻訳は translate_task の monkeypatch（f9 transcripts と同じ作法・
   製品コードにテスト用の口を彫らない）。測るのはフォルダ分岐の配管と事後条件であって、
   翻訳の質は battery の仕事。"""
import json
import sys
from pathlib import Path

import openpyxl
import pytest

# ★ M2 実装前の凍結検体: 実装が通ったら strict xfail が XPASS で赤くなり、
# この行を外す変更が同じ diff に必ず現れる（黙って通過できない）。
pytestmark = pytest.mark.xfail(strict=True, reason="M2（run のフォルダ分岐）実装前")

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
import ailine  # noqa: E402

HDRS = ["注文ID", "取引先", "金額"]


def _book(path, headers, rows=()):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(list(headers))
    for r in rows:
        ws.append(list(r) + [None] * (len(headers) - len(r)))
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def _mock_translation(monkeypatch, plan):
    monkeypatch.setattr(ailine, "translate_task",
                        lambda model, task, book_meta, temperature=0.1: {"plan": plan})


_EXTRACT_40000 = [{"op": "EXTRACT", "args": {"column": "金額", "cmp": "gte", "value": 40000}}]


def _run(folder, task, *extra, capsys=None):
    rc = ailine.main(["run", str(folder), task, *extra])
    out = capsys.readouterr().out if capsys else ""
    return rc, out


def test_e3_total_rows_do_not_leak_into_condition_matches(tmp_path, monkeypatch, capsys):
    """★ E3（検体の筆頭）: 『金額 40000 以上』の抽出では合計行が必ず条件を満たす。
       単位L の除外が条件適用の 前 に回っていないと、合計行が混ざって二重計上。"""
    _mock_translation(monkeypatch, _EXTRACT_40000)
    folder = tmp_path / "src"
    _book(folder / "a.xlsx", HDRS,
          [("J-1", "甲", 50000), ("J-2", "乙", 30000), ("合計", None, 80000)])
    _book(folder / "b.xlsx", HDRS, [("J-3", "丙", 45000)])
    rc, out = _run(folder, "金額が40000以上の行を抜き出して", capsys=capsys)
    assert rc == 0, out
    out_files = list(tmp_path.glob("*.xlsx"))
    assert out_files, f"出力ブックがフォルダの親に無い:\n{out}"
    ws = openpyxl.load_workbook(out_files[0]).active
    labels = [r[0].value for r in ws.iter_rows(min_row=2)]
    assert "合計" not in labels, f"★ 合計行 80000 が条件を通って混ざった: {labels}"
    assert sorted(x for x in labels if x) == ["J-1", "J-3"], labels


def test_folder_run_never_reports_excel_lock_lie(tmp_path, monkeypatch, capsys):
    """★ architect 致命4 の実バグの凍結: 今日の実装は run <フォルダ> に
       「Excel で開かれています」という嘘の診断を返す（フォルダの open(r+b) を
       ロックと誤読）。フォルダ分岐後はこの文言がフォルダに対して出ないこと。"""
    _mock_translation(monkeypatch, _EXTRACT_40000)
    folder = tmp_path / "src"
    _book(folder / "a.xlsx", HDRS, [("J-1", "甲", 50000)])
    rc, out = _run(folder, "金額が40000以上の行を抜き出して", capsys=capsys)
    assert "Excel で開かれています" not in out, f"フォルダに Excel ロックの嘘の診断:\n{out}"


def test_unsupported_op_on_folder_is_refused_by_name(tmp_path, monkeypatch, capsys):
    """★ E11: フォルダに未対応 op（並べ替え）→ 名指しの断り + 次の手 + exit 3。
       ★ 黙って 1 冊目に適用が最悪の形（原本無変更を機械で確認）。"""
    _mock_translation(monkeypatch, [{"op": "SORT", "args": {"column": "金額", "order": "desc"}}])
    folder = tmp_path / "src"
    _book(folder / "a.xlsx", HDRS, [("J-1", "甲", 50000)])
    before = (folder / "a.xlsx").read_bytes()
    rc, out = _run(folder, "金額で並べ替えて", capsys=capsys)
    assert rc == 3, f"exit={rc}\n{out}"
    assert "並べ替え" in out and ("抽出" in out or "1 冊" in out), f"名指しの断りと次の手が無い:\n{out}"
    assert (folder / "a.xlsx").read_bytes() == before, "断ったのに原本に触った"


def test_row_accounting_identity_is_reported(tmp_path, monkeypatch, capsys):
    """★ 憲法（行の完全会計・⑨）: 読めた各ファイルで データ行 = 一致 + 不一致 + 除外 の
       恒等が --json の multifile 節に出る（run の既存キーは壊さず入れ子で足す）。"""
    _mock_translation(monkeypatch, _EXTRACT_40000)
    folder = tmp_path / "src"
    _book(folder / "a.xlsx", HDRS,
          [("J-1", "甲", 50000), ("J-2", "乙", 30000), ("合計", None, 80000)])
    rc, out = _run(folder, "金額が40000以上の行を抜き出して", "--json", capsys=capsys)
    assert rc == 0, out
    data = json.loads(out)
    mf = data.get("multifile")
    assert mf, f"multifile 節が --json に無い: {list(data)}"
    f = {x["name"]: x for x in mf["files"]}["a.xlsx"]
    assert f["rows_matched"] == 1 and f["rows_unmatched"] == 1 and f["total_rows_excluded"] == 1
    assert f["rows_matched"] + f["rows_unmatched"] + f["total_rows_excluded"] == 3
