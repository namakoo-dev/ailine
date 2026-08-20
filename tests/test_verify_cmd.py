"""`ailine verify <out.xlsx> <srcfolder>`（検算の単独再実行）の検体。
   ★ 実装前に書いた赤い検体。信用の条件⑥「検算側を単独で再実行できること」──
   需要ノートの言葉で「信じる対象が道具から検算に移る」。

   契約の要点:
   - stack の出力ブックと元フォルダだけを引数に、検算（行数照合・数値列 Σ 照合）を再実行
   - 出所列（元ファイル/元行）を使って出力の各行を原本に引き当てる
   - 合格: 両側の数字を並べて exit 0 ／ 不一致: どの列がいくつ対いくつ、まで名指しで exit 5
   - どちらのファイルも変更しない（読むだけ）"""
import hashlib
import subprocess
import sys
from pathlib import Path

import openpyxl

REPO = Path(__file__).resolve().parent.parent
HDRS = ["注文ID", "取引先", "金額"]


def _book(path, headers, rows=()):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(list(headers))
    for r in rows:
        ws.append(list(r) + [None] * (len(headers) - len(r)))
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def _run(cmd, *args):
    return subprocess.run(
        [sys.executable, str(REPO / "ailine.py"), cmd, *map(str, args)],
        capture_output=True, text=True, timeout=180, encoding="utf-8")


def _made(tmp_path):
    folder = tmp_path / "src"
    _book(folder / "a.xlsx", HDRS, [("J-1", "甲", 100), ("J-2", "乙", 200)])
    _book(folder / "b.xlsx", HDRS, [("J-3", "丙", 300)])
    out = tmp_path / "out.xlsx"
    p = _run("stack", folder, "--out", out)
    assert p.returncode == 0, "前提の stack が失敗:\n" + p.stderr[-500:]
    return folder, out


def test_verify_passes_on_untouched_output_with_both_sides(tmp_path):
    folder, out = _made(tmp_path)
    p = _run("verify", out, folder)
    assert p.returncode == 0, p.stdout + p.stderr[-500:]
    assert p.stdout.count("600") >= 2, f"Σ の両側表示が無い:\n{p.stdout}"


def test_verify_names_a_tampered_value_and_exits_5(tmp_path):
    """出力の金額 1 セルを 100→999 に改竄 → 列名と両側の数字つきで不合格・exit 5。"""
    folder, out = _made(tmp_path)
    wb = openpyxl.load_workbook(out)
    ws = wb.active
    for row in ws.iter_rows(min_row=2):
        if row[2].value == 100:
            row[2].value = 999
            break
    wb.save(out)
    p = _run("verify", out, folder)
    assert p.returncode == 5, f"exit={p.returncode}\n{p.stdout}"
    assert "金額" in p.stdout
    assert "600" in p.stdout and "1499" in p.stdout, f"両側の数字が無い:\n{p.stdout}"


def test_verify_names_a_deleted_row_and_exits_5(tmp_path):
    """出力からデータ行を 1 行削除 → 行数の不一致が両側の数字つきで出る・exit 5。"""
    folder, out = _made(tmp_path)
    wb = openpyxl.load_workbook(out)
    ws = wb.active
    ws.delete_rows(2)
    wb.save(out)
    p = _run("verify", out, folder)
    assert p.returncode == 5
    assert "3" in p.stdout and "2" in p.stdout, f"行数の両側が無い:\n{p.stdout}"


def test_verify_is_read_only_on_both_sides(tmp_path):
    folder, out = _made(tmp_path)
    before = {p.name: hashlib.sha256(p.read_bytes()).hexdigest()
              for p in [*folder.glob("*.xlsx"), out]}
    assert _run("verify", out, folder).returncode == 0
    after = {p.name: hashlib.sha256(p.read_bytes()).hexdigest()
             for p in [*folder.glob("*.xlsx"), out]}
    assert before == after
