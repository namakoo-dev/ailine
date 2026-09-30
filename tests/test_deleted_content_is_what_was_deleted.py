# 「消した中身」は実際に消した行を言う（2026-09-30・盲検 7 体目の致命 ②）。
#
# ★ 実測: 見出し 4 行目・データ 5〜12 行目の請求一覧で「入金状況が入金済の行を削除して」
#   → 解釈は 5・7・11 行目、ファイルも 5・7・11 を正しく消した。ところが画面の
#   「消した中身（3 行）」は 5・6・7 行目を出した（残っている INV-0902 を消したと言った）。
# ★ 原因は check_delete_rows の中の片配線: 残りの突き合わせは宣言された一覧（_delete_rows）
#   から作り、note_deleted だけが at から連続で拾っていた。
#   直しは「消した添字の集合」を 1 回だけ作り、両方がそれを読む形。
#
# 契約:
#   ① 非連続の行を消したとき、_deleted は**実際に消した行**の中身（昇順）
#   ② 見出しが 1 行目でない表でも同じ（添字の起点がずれない）
#   ③ _delete_rows が無い従来の連続削除は従来どおり

import sys
from pathlib import Path

import openpyxl

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
import ailine  # noqa: E402

HEAD = ["請求番号", "取引先", "金額", "入金状況"]
DATA = [
    ["INV-0901", "A社", 1000, "入金済"],   # 5 行目
    ["INV-0902", "B社", 2000, "未入金"],   # 6
    ["INV-0903", "C社", 3000, "入金済"],   # 7
    ["INV-0904", "D社", 4000, "未入金"],   # 8
    ["INV-0905", "E社", 5000, "未入金"],   # 9
    ["INV-0906", "F社", 6000, "未入金"],   # 10
    ["INV-0907", "G社", 7000, "入金済"],   # 11
    ["INV-0908", "H社", 8000, "未入金"],   # 12
]


def _book(tmp_path, data, name):
    """見出しを 4 行目に置く（1〜3 行目は表題と空行）。"""
    p = tmp_path / name
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "請求"
    ws.append(["請求一覧"])
    ws.append([])
    ws.append([])
    ws.append(HEAD)
    for r in data:
        ws.append(r)
    wb.save(p)
    return p


def test_non_contiguous_delete_reports_the_rows_actually_deleted(tmp_path):
    before = _book(tmp_path, DATA, "before.xlsx")
    kept = [r for r in DATA if r[3] != "入金済"]
    after = _book(tmp_path, kept, "after.xlsx")
    args = {"at": 5, "count": 3, "_delete_rows": [5, 7, 11]}
    status, reason = ailine.check_delete_rows(after, args, header_row=4,
                                              source_book=before)
    assert status == "pass", reason
    assert args["_deleted"] == [list(r) for r in DATA if r[3] == "入金済"], \
        args.get("_deleted")


def test_declared_rows_out_of_order_are_reported_ascending(tmp_path):
    before = _book(tmp_path, DATA, "before.xlsx")
    kept = [r for i, r in enumerate(DATA) if i not in (0, 2, 6)]
    after = _book(tmp_path, kept, "after.xlsx")
    args = {"at": 5, "count": 3, "_delete_rows": [11, 5, 7]}
    status, reason = ailine.check_delete_rows(after, args, header_row=4,
                                              source_book=before)
    assert status == "pass", reason
    assert [r[0] for r in args["_deleted"]] == ["INV-0901", "INV-0903", "INV-0907"]


def test_contiguous_delete_without_a_list_is_unchanged(tmp_path):
    before = _book(tmp_path, DATA, "before.xlsx")
    after = _book(tmp_path, DATA[:1] + DATA[3:], "after.xlsx")   # 6・7 行目を消した
    args = {"at": 6, "count": 2}
    status, reason = ailine.check_delete_rows(after, args, header_row=4,
                                              source_book=before)
    assert status == "pass", reason
    assert [r[0] for r in args["_deleted"]] == ["INV-0902", "INV-0903"]
