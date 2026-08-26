# データの出入口の盲検（2026-08-26）── export-pdf の読み戻しが「載っている」を
# **PDF 全文への素の部分文字列の包含**で判定していた（位置も出現回数も見ない）。
#
# 致命1 の実測:
#   ・隠し行・印刷範囲で 6 行中 3 行が PDF に無いのに `✓ 欠落 0`
#     （消えた行の値が他の行にも在ると「載っている」と数える）
#   ・隠した品番 `12345` は可視の `123456` の部分文字列なので検出されない
#
# 高5 の実測: 時刻・パーセント・真偽値のある表は**どうやっても ✓ が出ず常に exit 3**
#   （renderings に分岐が無く `09:00:00` / `True` を探していた。PDF は `9:00` / `TRUE`）
# 高6 の実測: 列幅で折り返されたセルが × になる（空白しか除いていなかった）
#
# 契約:
#   ① 同じ値が n 個あるなら、PDF にも n 回無ければ欠落
#   ② 数字は前後が数字でないこと（部分一致で拾わない）
#   ③ 時刻・真偽値は PDF の見た目の表記を候補に持つ（誤爆させない）
#   ④ 折り返し（改行）で不一致にしない
#   ⑤ 空文字の罠: 先頭・末尾に在る値を取りこぼさない

import datetime
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
from ailine_core import pdf_export as pe  # noqa: E402


# --- ①②⑤ 数え方 ---------------------------------------------------------------------

@pytest.mark.parametrize("flat,cand,want", [
    ("123456 7", "12345", 0),      # ② 部分一致で拾わない
    ("10001", "1000", 0),
    ("1000 1000", "1000", 2),      # ① 2 回在る
    ("1000", "1000", 1),           # ⑤ 先頭かつ末尾（空文字の罠）
    ("x1000", "1000", 1),          # ⑤ 末尾
    ("鉛筆5鉛筆5", "鉛筆", 2),
])
def test_occurrence_counting(flat, cand, want):
    assert pe._count_occurrences(flat, cand) == want


def test_a_vanished_duplicate_row_is_missing(monkeypatch, tmp_path):
    """★ 実測の形: 同じ金額の行が 2 つあり、片方が PDF から消えている。

    旧実装は「載っている」と数えて ✓ 欠落 0 を出した。
    """
    monkeypatch.setattr(pe, "readback_available", lambda: True)
    monkeypatch.setattr(pe, "read_pdf_text", lambda p: "品名 数量 単価\n鉛筆 5 80")
    # シートには鉛筆の行が 2 つある（1 つが隠し行で PDF に出ていない）
    r = pe.verify_values_in_pdf(tmp_path / "x.pdf", ["鉛筆", 5, 80, "鉛筆", 5, 80])
    assert r.missing, "同じ値の行が消えているのに欠落 0 と言った"


def test_no_false_alarm_when_all_copies_are_there(monkeypatch, tmp_path):
    """④⑤ 誤爆しない: 必要な回数だけ載っていれば通る。"""
    monkeypatch.setattr(pe, "readback_available", lambda: True)
    monkeypatch.setattr(pe, "read_pdf_text", lambda p: "鉛筆 5 80\n鉛筆 5 80")
    r = pe.verify_values_in_pdf(tmp_path / "x.pdf", ["鉛筆", 5, 80, "鉛筆", 5, 80])
    assert not r.missing, r.missing


def test_a_substring_number_does_not_count_as_present(monkeypatch, tmp_path):
    """② 実測: 隠した品番 12345 が可視の 123456 に含まれて素通りしていた。"""
    monkeypatch.setattr(pe, "readback_available", lambda: True)
    monkeypatch.setattr(pe, "read_pdf_text", lambda p: "品番 数量\n123456 7")
    r = pe.verify_values_in_pdf(tmp_path / "x.pdf", [12345])
    assert r.missing == ["12345"], r.missing


# --- ③④ 誤爆を止める -----------------------------------------------------------------

def test_time_and_bool_use_the_pdf_spelling():
    """③ 勤怠表・真偽値のある表が『どうやっても ✓ を出せない』のを止める。"""
    assert "9:00" in pe.renderings(datetime.time(9, 0))
    assert "TRUE" in pe.renderings(True)
    assert "FALSE" in pe.renderings(False)


@pytest.mark.xfail(strict=True, reason=(
    "★ 未処置（高6）: 列幅で折り返されたセルは、値の文字が PDF 上で**離れた場所**に"
    "分かれる（`1月` → 見出しの `月` と本文の `1`）。空白・改行を除いても隣り合わないので、"
    "テキスト層の文字列照合では拾えない。直すには pdfplumber の語の座標を使う必要があり、"
    "今週の枠では番人ごと用意できない。誤 × であって誤 ✓ ではないので順位を下げた。"
    "★ この xfail が XPASS になったら、直したのに印を消し忘れている。"))
def test_wrapped_cell_is_still_found(monkeypatch, tmp_path):
    """④ 列幅で折り返された `1月` を欠落にしない ── **まだ守れていない契約**。"""
    monkeypatch.setattr(pe, "readback_available", lambda: True)
    monkeypatch.setattr(pe, "read_pdf_text", lambda p: "月 売上\n1 100\n月\n2 200")
    r = pe.verify_values_in_pdf(tmp_path / "x.pdf", ["1月"])
    assert not r.missing, r.missing
