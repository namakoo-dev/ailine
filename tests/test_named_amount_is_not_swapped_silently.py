# -*- coding: utf-8 -*-
"""依頼が名指しした列は、他の列に黙って替えない（2026-10-01・盲検 7 体目）。

★★ 起きたこと（請求と入金の突き合わせ・本体が再現済み）:

    $ ailine run 請求一覧.xlsx 入金明細.xlsx "請求番号をキーに、請求の税込金額と入金額を突き合わせて、…"
    キー: 請求番号(A) / 請求番号(B)　金額: 税抜金額(A) / 入金額(B)
    ⚠ INV-0901: −15000（A 150000 / B 165000）   … 正しく払われた請求まで全部「差額あり」・exit 0

  『税込金額』は式（=E5+F5）のままで計算結果が無く、照合の読みでは空の列に見える。
  数値の列に数えられず、残った数値の列『税抜金額』が resolve_role の②（型が合う列が 1 本なら採る）で
  **断りなく**選ばれた。依頼の項（金額＝税込金額）が欠けた嘘。
  ★ 式の列に値を入れる器官（read_with_values_filled_in・e58d9fa）は在ったが、
    **列が決まらなかった回にしか**呼ばれない ── 黙って決まってしまったので器官まで届かなかった。

★ 直し: 依頼が名指しした列が使えない時は②へ落ちず「決まらない」を返す（match.named_but_unusable）。
  決まらなければ既存の器官が値を入れて読み直し、入れられなければ理由つきで断る。
"""
import argparse
import contextlib
import io
import shutil
import sys
from pathlib import Path

import openpyxl

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
import ailine  # noqa: E402
from ailine_core import match  # noqa: E402

TASK = "請求番号をキーに、請求の税込金額と入金額を突き合わせて、入金が足りないものを教えて"
INVOICES = [("INV-0901", 150000, "入金済"), ("INV-0902", 48000, "未入金"),
            ("INV-0903", 220000, "入金済"), ("INV-0906", 12500, "一部入金")]
PAID = [("INV-0901", 165000), ("INV-0903", 242000), ("INV-0906", 10000)]


def _invoices(p: Path, *, cached: bool) -> Path:
    """盲検 7 体目の形: タイトル行つき・見出し 4 行目・消費税と税込金額は式（cached=False で値なし）。"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "2026年9月"
    ws.append(["請求一覧（2026年9月分）"])
    ws.append(["作成：2026/9/30"])
    ws.append([])
    ws.append(["請求番号", "取引先名", "税抜金額", "消費税", "税込金額", "入金状況"])
    for i, (no, net, state) in enumerate(INVOICES, start=5):
        tax = net // 10
        ws.append([no, "取引先", net, tax if cached else f"=ROUNDDOWN(C{i}*0.1,0)",
                   net + tax if cached else f"=C{i}+D{i}", state])
    wb.save(p)
    wb.close()
    return p


def _payments(p: Path) -> Path:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "入金"
    ws.append(["普通預金 入金明細（9月）"])
    ws.append([])
    ws.append(["入金日", "請求番号", "振込名義", "入金額"])
    for no, amt in PAID:
        ws.append(["2026/09/29", no, "ﾒｲｷﾞ", amt])
    wb.save(p)
    wb.close()
    return p


def _read(p: Path):
    _hr, headers, rows = ailine._peek_match_book(p)
    return headers, rows


def _run_match(a: Path, b: Path, monkeypatch, tmp_path) -> tuple:
    monkeypatch.setenv("AILINE_HOME", str(tmp_path / "home"))
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf):
        rc = ailine.cmd_run_match(argparse.Namespace(json=False), a, b, TASK)
    return rc, buf.getvalue()


# --- 解決そのもの -----------------------------------------------------------------------

def test_the_named_column_without_values_is_not_replaced(tmp_path):
    """★ 本体: 『税込金額』に値が無い時、金額を『税抜金額』に黙って決めない。"""
    ha, ra = _read(_invoices(tmp_path / "A.xlsx", cached=False))
    hb, rb = _read(_payments(tmp_path / "B.xlsx"))
    res = match.resolve_columns(TASK, ha, ra, hb, rb)
    assert res.amount_a != "税抜金額", "★ 依頼の『税込金額』を黙って『税抜金額』に替えている"
    assert res.amount_a is None and not res.ok
    assert ("A", "amount", "税込金額") in {(s, r, h) for s, r, h, _w in res.unusable}
    # 他の役割は今までどおり決まる（直しすぎない）
    assert (res.key_a, res.key_b, res.amount_b) == ("請求番号", "請求番号", "入金額")


def test_positive_control_with_values_the_named_column_is_used(tmp_path):
    """陽性対照: 同じ表に値が入っていれば、名指しどおり『税込金額』が採られる。"""
    ha, ra = _read(_invoices(tmp_path / "A.xlsx", cached=True))
    hb, rb = _read(_payments(tmp_path / "B.xlsx"))
    res = match.resolve_columns(TASK, ha, ra, hb, rb)
    assert res.ok and res.amount_a == "税込金額" and not res.unusable


def test_the_hint_does_not_offer_the_other_column(tmp_path):
    """★ 断りの案内が『税抜金額を金額に』を勧めない（打てば通るが、答えが違う）。"""
    ha, ra = _read(_invoices(tmp_path / "A.xlsx", cached=False))
    hb, rb = _read(_payments(tmp_path / "B.xlsx"))
    res = match.resolve_columns(TASK, ha, ra, hb, rb)
    assert match.rewording_that_resolves(TASK, ha, ra, hb, rb, res.unresolved) is None
    line = match.rewording_line(TASK, ha, ra, hb, rb, res.unresolved, {"A": "A.xlsx", "B": "B.xlsx"})
    assert "税抜金額を金額に" not in line
    assert "『税込金額』" in line and "黙って照合はしません" in line, line


def test_a_column_named_for_the_other_role_does_not_block(tmp_path):
    """直しすぎない: 「備考をキーに」の空の『備考』は、金額の役割を止めない。"""
    headers = ["備考", "品名", "金額"]
    rows = [(2, [None, "ボルト", 100], []), (3, [None, "ナット", 50], [])]
    assert match.named_but_unusable("備考をキーに突き合わせて", headers, {"金額"},
                                    {"備考"}, "amount") == []
    got, _c = match.resolve_role("備考をキーに突き合わせて", headers, {"金額"}, "amount",
                                 headers, match.empty_columns(headers, rows))
    assert got == "金額"


def test_a_role_marked_column_of_the_wrong_type_is_not_replaced():
    """同じ型の別の形: 「品名を金額に」と言われて品名が文字の列なら、唯一の数値列へ黙って替えない。"""
    headers = ["請求番号", "品名", "数量"]
    got, _c = match.resolve_role("品名を金額に突き合わせて", headers, {"数量"}, "amount", headers, set())
    assert got is None
    # 鍵の側: 数値だけの列を「キーに」と言われたら、唯一の文字の列へ黙って替えない
    got, _c = match.resolve_role("伝票番号をキーに", ["伝票番号", "取引先"], {"伝票番号"}, "key",
                                 ["伝票番号", "取引先"], set())
    assert got is None


# --- 入口ごと（cmd_run_match） ----------------------------------------------------------

def test_the_route_fills_values_and_matches_on_the_named_column(tmp_path, monkeypatch):
    """★ 配線: 決まらない → 既存の器官で値を入れて読み直す → 名指しどおり税込で照合する。

    ★ LibreOffice の代わりに、値の入った同じ表を返す（器官の中身は local の試験が持つ）。"""
    a = _invoices(tmp_path / "請求一覧.xlsx", cached=False)
    b = _payments(tmp_path / "入金明細.xlsx")
    filled = _invoices(tmp_path / "filled_src.xlsx", cached=True)

    def fake_normalize(book, workdir, timeout=None):
        dst = Path(workdir) / Path(book).name
        shutil.copy(filled, dst)
        return dst

    monkeypatch.setattr(ailine, "normalize_book", fake_normalize)
    rc, out = _run_match(a, b, monkeypatch, tmp_path)
    assert rc == 0, out[-800:]
    assert "金額: 税込金額(A) / 入金額(B)" in out, out[-800:]
    assert "LibreOffice で開いて計算させました" in out
    assert "INV-0901" not in out and "INV-0903" not in out, "★ 正しく払われた請求に ⚠ が出ている:\n" + out
    assert "INV-0906: +3750" in out, out[-800:]
    assert "INV-0902: A のみ" in out, out[-800:]


def test_the_route_refuses_when_values_cannot_be_filled(tmp_path, monkeypatch):
    """★ 値を入れられなければ断る ── 『税抜金額』で黙って照合しない（exit 0 にしない）。"""
    def boom(book, workdir, timeout=None):
        raise SystemExit(9)

    monkeypatch.setattr(ailine, "normalize_book", boom)
    a = _invoices(tmp_path / "請求一覧.xlsx", cached=False)
    b = _payments(tmp_path / "入金明細.xlsx")
    rc, out = _run_match(a, b, monkeypatch, tmp_path)
    assert rc == 3, f"exit {rc}:\n{out[-800:]}"
    assert "税抜金額(A)" not in out and "INV-0901: −" not in out, out[-800:]
    assert "『税込金額』は式のままで計算結果が入っていない" in out, out[-800:]
    assert "税抜金額を金額に" not in out, "★ 別の列を勧めている:\n" + out[-800:]
