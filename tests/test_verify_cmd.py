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


def test_verify_passes_on_output_whose_totals_were_excluded(tmp_path):
    """★ 実機の敵対検分（2026-08-21 06:1x）で踏んだ穴の凍結: 合計行を正しく除外した
       正当な出力に対して、verify が除外の意味論を再現できず 元7/出力4 で誤 FAIL した。
       verify は stack と同じ除外規則（total_row）を XML 直読みの側でも再現すること。"""
    folder = tmp_path / "src"
    _book(folder / "a.xlsx", HDRS,
          [("J-1", "甲", 600), ("J-2", "乙", 400),
           ("小計", None, 1000), ("消費税", None, 100), ("合計", None, 1100)])
    out = tmp_path / "out.xlsx"
    p = _run("stack", folder, "--out", out)
    assert p.returncode == 0, p.stdout + p.stderr[-500:]
    v = _run("verify", out, folder)
    assert v.returncode == 0, f"正当な出力で verify が落ちた:\n{v.stdout}"
    assert v.stdout.count("1100") >= 2, f"Σ の両側（税込 1100）が無い:\n{v.stdout}"


def test_verify_agrees_with_stack_when_dates_and_reorder_present(tmp_path):
    """★ 予測①の的中を凍結（2026-08-21 06:1x 実機 probe の再現）: XML 直読みは日付の
       シリアル値を数値と見るため、数値列の引き当てが stack(openpyxl=datetime は非数値)と
       食い違い、除外規則が別の列で走って 元7/出力4 の誤 FAIL になった。
       両側の読み手は『日付書式のセルは数値列の候補にしない』で一致すること。"""
    import datetime
    folder = tmp_path / "src"
    _book(folder / "inv1.xlsx", ["注文ID", "受注日", "金額"],
          [("J-1", datetime.date(2026, 8, 1), 600), ("J-2", datetime.date(2026, 8, 3), 400),
           ("小計", None, 1000), ("消費税", None, 100), ("合計", None, 1100)])
    _book(folder / "inv2.xlsx", ["金額", "注文ID", "受注日"],
          [(300, "J-3", datetime.date(2026, 8, 5)), (999, "合計", None)])
    out = tmp_path / "out.xlsx"
    p = _run("stack", folder, "--out", out)
    assert p.returncode == 0, p.stdout + p.stderr[-500:]
    assert "Σ受注日" not in p.stdout, "日付列を合計している"
    v = _run("verify", out, folder)
    assert v.returncode == 0, f"正当な出力で verify が落ちた:\n{v.stdout}"
    assert v.stdout.count("1400") >= 2, f"Σ金額の両側（1400）が無い:\n{v.stdout}"
    assert "Σ受注日" not in v.stdout
