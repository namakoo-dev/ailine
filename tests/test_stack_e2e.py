"""M1書き `ailine stack <folder> --out <path>`（縦積み・UNION ALL）の検体。
   ★ 実装前に書いた赤い検体（DESIGN-20260821-multifile v2/v2.1・specimen-first）。

   契約の要点:
   - 基準ファイル方式で照合できたファイルだけを、基準の列順に揃えて新ブックへ縦積み
   - 出所列（『元ファイル』『元行』）を右端に付ける ── 抜き打ち検証の装置 + 自分の出力の署名
   - 単位L: 合計行は積まない（除外は開示・閉じる検査の不一致は両側の数字つき ⚠）
   - 事後条件: Σ行数一致・数値列ごとの Σ 両側表示・原本無変更・分母つき報告
   - 出力先が既存の人のファイルなら止まる（exit 7・--overwrite で通す）。
     自分の前回の縦積み出力なら作り直して開示（単位H の作法）
   - 決定論: 同一入力で 2 回走らせて、セル内容が完全一致（zip バイトではなく中身）"""
import hashlib
import json
import subprocess
import sys
from pathlib import Path

import openpyxl
import pytest

REPO = Path(__file__).resolve().parent.parent

HDRS = ["注文ID", "取引先", "金額"]


def _book(path, headers, rows=()):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(list(headers))
    for r in rows:
        padded = list(r) + [None] * (len(headers) - len(r))
        ws.append(padded[: len(headers)])
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def _stack(folder, out, *extra):
    return subprocess.run(
        [sys.executable, str(REPO / "ailine.py"), "stack", str(folder), "--out", str(out), *extra],
        capture_output=True, text=True, timeout=180, encoding="utf-8")


def _cells(path):
    """出力の中身（シート名・セル値・型）を決定論比較用に正規化して返す。"""
    wb = openpyxl.load_workbook(path)
    out = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if c.value is not None:
                    out.append((ws.title, c.coordinate, type(c.value).__name__, str(c.value)))
    wb.close()
    return out


def _folder3(tmp_path):
    folder = tmp_path / "src"
    _book(folder / "a.xlsx", HDRS, [("J-1", "甲", 100), ("J-2", "乙", 200)])
    _book(folder / "b.xlsx", ["金額", "注文ID", "取引先"], [(300, "J-3", "丙")])   # 並びだけ違う
    _book(folder / "c.xlsx", HDRS, [("J-4", "丁", 50), ("合計", None, 350)])       # ★ 合計行入り
    return folder


def test_stack_unions_rows_in_base_column_order_with_provenance(tmp_path):
    """本命: 3 冊 → 1 冊。基準（a.xlsx）の列順に揃い、出所列が右端に付く。"""
    folder = _folder3(tmp_path)
    out = tmp_path / "out.xlsx"
    p = _stack(folder, out)
    assert p.returncode == 0, p.stderr[-800:] + p.stdout[-800:]
    ws = openpyxl.load_workbook(out).active
    headers = [c.value for c in ws[1]]
    assert headers == HDRS + ["元ファイル", "元行"], headers
    rows = [[c.value for c in r] for r in ws.iter_rows(min_row=2)]
    assert ["J-3", "丙", 300, "b.xlsx", 2] in rows, "並べ替え縦積みが崩れている"
    assert len(rows) == 4, f"データ行 4 のはず（合計行は積まない）: {len(rows)}"


def test_total_row_is_not_stacked_and_exclusion_is_disclosed(tmp_path):
    """単位L の配線: c.xlsx の合計行は積まれず、その事実がファイル名つきで開示される。"""
    folder = _folder3(tmp_path)
    out = tmp_path / "out.xlsx"
    p = _stack(folder, out)
    assert p.returncode == 0
    assert "合計行" in p.stdout and "c.xlsx" in p.stdout, f"除外の開示が無い:\n{p.stdout}"
    ws = openpyxl.load_workbook(out).active
    labels = [r[0].value for r in ws.iter_rows(min_row=2)]
    assert "合計" not in labels


def test_report_shows_both_sides_of_row_and_sum_reconciliation(tmp_path):
    """事後条件①②: 行数と数値列の合計を 両側の数字 で並べる（「一致」だけの報告は感想）。"""
    folder = _folder3(tmp_path)
    out = tmp_path / "out.xlsx"
    p = _stack(folder, out)
    assert p.returncode == 0
    assert "3 ファイル中 3 積んだ" in p.stdout, f"分母つき報告が無い:\n{p.stdout}"
    # 金額の総和 = 100+200+300+50 = 650 が「元」「出力」の両側で出る
    assert p.stdout.count("650") >= 2, f"Σ金額の両側表示が無い:\n{p.stdout}"


def test_unmatched_file_is_skipped_named_and_counted(tmp_path):
    """列が欠けたファイルは積まず、名指し + 分母に現れる。exit は 0（黙る失敗だけが罪）。"""
    folder = _folder3(tmp_path)
    _book(folder / "d.xlsx", ["注文ID", "取引先"], [("J-9", "戊")])
    out = tmp_path / "out.xlsx"
    p = _stack(folder, out)
    assert p.returncode == 0
    assert "4 ファイル中 3 積んだ" in p.stdout
    assert "d.xlsx" in p.stdout and "金額" in p.stdout


def test_originals_unchanged_and_output_outside_folder(tmp_path):
    """事後条件④: 原本は 1 バイトも変わらない（sha256）。"""
    folder = _folder3(tmp_path)
    before = {p.name: hashlib.sha256(p.read_bytes()).hexdigest() for p in folder.glob("*.xlsx")}
    out = tmp_path / "out.xlsx"
    r = _stack(folder, out)
    assert r.returncode == 0
    after = {p.name: hashlib.sha256(p.read_bytes()).hexdigest() for p in folder.glob("*.xlsx")}
    assert before == after


def test_deterministic_content_across_two_runs(tmp_path):
    """事後条件⑥: 同一入力で 2 回 → セル内容（シート名・座標・型・値）が完全一致。
       ★ zip バイト一致は openpyxl の保存時刻で原理的に不能（実測）── 中身で比べる。"""
    folder = _folder3(tmp_path)
    out1, out2 = tmp_path / "o1.xlsx", tmp_path / "o2.xlsx"
    assert _stack(folder, out1).returncode == 0
    assert _stack(folder, out2).returncode == 0
    assert _cells(out1) == _cells(out2)


def test_existing_foreign_output_is_a_gate_not_a_casualty(tmp_path):
    """writes=(new_book,) の関所: 出力先に人のファイルが居たら止まる（exit 7・中身無傷）。
       --overwrite で通す（既存の overwrite 作法と同じ）。"""
    folder = _folder3(tmp_path)
    out = tmp_path / "mine.xlsx"
    _book(out, ["大事な列"], [("消えては困る",)])
    sha = hashlib.sha256(out.read_bytes()).hexdigest()
    p = _stack(folder, out)
    assert p.returncode == 7, f"exit={p.returncode}\n{p.stdout}\n{p.stderr}"
    assert hashlib.sha256(out.read_bytes()).hexdigest() == sha, "関所で止まったのに書き換えた"
    p2 = _stack(folder, out, "--overwrite")
    assert p2.returncode == 0


def test_own_previous_output_is_rebuilt_with_disclosure(tmp_path):
    """出力先が自分の前回の縦積み（署名 = 右端の出所列 2 本）なら作り直し + 開示（単位H の作法）。"""
    folder = _folder3(tmp_path)
    out = tmp_path / "out.xlsx"
    assert _stack(folder, out).returncode == 0
    p = _stack(folder, out)   # 2 回目: 自分の出力の上に
    assert p.returncode == 0, f"自分の前回出力で関所が閉まった:\n{p.stdout}\n{p.stderr}"
    assert "前回" in p.stdout or "作り直" in p.stdout, "作り直しの開示が無い"


def test_own_output_inside_input_folder_is_excluded_from_sources(tmp_path):
    """★ 敵対検体（V6・自己参照）: 出力を入力フォルダ内に置いて 2 回走らせても
       行数が増えない（自分の出力を入力から署名で除外 + 開示）。"""
    folder = _folder3(tmp_path)
    out = folder / "まとめ.xlsx"
    assert _stack(folder, out).returncode == 0
    n1 = openpyxl.load_workbook(out).active.max_row
    p = _stack(folder, out)
    assert p.returncode == 0
    n2 = openpyxl.load_workbook(out).active.max_row
    assert n1 == n2, f"2 週目に自分の出力を食って {n1}→{n2} 行"
    assert "除外" in p.stdout or "自分の出力" in p.stdout, "自己参照除外の開示が無い"


def test_mismatched_total_row_warns_with_both_numbers(tmp_path):
    """単位L の閉じる検査: 合計行の値が明細の和と合わないファイルは 両側の数字つき ⚠。"""
    folder = tmp_path / "src"
    _book(folder / "a.xlsx", HDRS, [("J-1", "甲", 600), ("J-2", "乙", 400), ("合計", None, 1200)])
    out = tmp_path / "out.xlsx"
    p = _stack(folder, out)
    assert p.returncode == 0
    assert "1200" in p.stdout and "1000" in p.stdout, f"両側の数字が無い:\n{p.stdout}"


def test_json_carries_denominator_per_file_and_totals(tmp_path):
    """--json: 分母・ファイルごとの積んだ行数・除外合計行数・Σ の両側が機械可読で出る。"""
    folder = _folder3(tmp_path)
    out = tmp_path / "out.xlsx"
    p = _stack(folder, out, "--json")
    assert p.returncode == 0, p.stderr[-500:]
    data = json.loads(p.stdout)
    assert data["denominator"] == 3 and data["stacked_files"] == 3
    assert data["rows_written"] == 4
    per = {f["name"]: f for f in data["files"]}
    assert per["c.xlsx"]["rows_stacked"] == 1 and per["c.xlsx"]["total_rows_excluded"] == 1
    assert data["sums"]["金額"]["source"] == 650 and data["sums"]["金額"]["output"] == 650


def test_data_row_with_empty_numeric_cell_is_still_stacked(tmp_path):
    """★ 地雷の先回り（total_row 検分 2026-08-21）: adopted_rows は数値行だけを数える
       （閉じる検査の分母としては正しい）。だが積む対象は『除外行以外の全データ行』──
       金額が空欄のデータ行を adopted_rows 基準で黙って落とさないこと。"""
    folder = tmp_path / "src"
    _book(folder / "a.xlsx", HDRS,
          [("J-1", "甲", 100), ("J-2", "乙", None), ("合計", None, 100)])
    out = tmp_path / "out.xlsx"
    p = _stack(folder, out)
    assert p.returncode == 0, p.stderr[-500:]
    ws = openpyxl.load_workbook(out).active
    rows = [[c.value for c in r] for r in ws.iter_rows(min_row=2)]
    assert any(r[0] == "J-2" for r in rows), "金額が空のデータ行が黙って消えた"
    assert len(rows) == 2, f"データ行 2（J-1, J-2）のはず: {len(rows)}"


def test_suffixed_provenance_output_is_still_recognized_as_own(tmp_path):
    """★ 凍結予測③が的中した穴（2026-08-21 06:1x 実機で確認）: 基準が『元ファイル』列を
       持つと出所列は 元ファイル_2 になるが、①その事実の開示が無く ②自分の前回出力の
       署名判定が素の列名しか見ず exit 7 で誤って閉まった。署名はサフィックス形も自分と
       認識し、衝突の開示も 1 行出すこと。"""
    folder = tmp_path / "src"
    _book(folder / "a.xlsx", ["注文ID", "元ファイル", "金額"], [("J-1", "memo1", 100)])
    _book(folder / "b.xlsx", ["注文ID", "元ファイル", "金額"], [("J-2", "memo2", 200)])
    out = tmp_path / "out.xlsx"
    p1 = _stack(folder, out)
    assert p1.returncode == 0
    assert "元ファイル_2" in p1.stdout, f"サフィックスの開示が無い:\n{p1.stdout}"
    p2 = _stack(folder, out)   # 2 周目: 自分のサフィックスつき出力の上に
    assert p2.returncode == 0, f"自分の出力を他人と誤認して閉まった (exit={p2.returncode})"
    assert "前回" in p2.stdout or "作り直" in p2.stdout


# ---- jisaku-review 確定 6 件の凍結（2026-08-21 06:4x・81b9aa9..bdfc4a8 への盲検レビュー）----


def test_foreign_file_with_coincident_provenance_names_is_still_gated(tmp_path):
    """★ review#1 critical: 末尾 2 列がたまたま『元ファイル』『元行』という名前の
       人のファイルを、名前だけの署名で自分の前回出力と誤認して無警告上書きした。
       署名は名前の一致だけで成立してはならない ── 人のデータは exit 7 で守る。"""
    folder = tmp_path / "src"
    _book(folder / "a.xlsx", HDRS, [("J-1", "甲", 100)])
    out = tmp_path / "mine.xlsx"
    _book(out, HDRS + ["元ファイル", "元行"],
          [("X-1", "人の大事なデータ", 999, "note", "n/a")])
    sha = hashlib.sha256(out.read_bytes()).hexdigest()
    p = _stack(folder, out)
    assert p.returncode == 7, f"人のファイルを署名誤認で通した (exit={p.returncode})\n{p.stdout}"
    assert hashlib.sha256(out.read_bytes()).hexdigest() == sha, "人のデータが消えた"


def test_gate_refusal_discloses_reason_and_next_step(tmp_path):
    """★ review#5: exit 7 の関所が無言で閉まっていた（報告が成果物、の自己矛盾）。
       止まる時こそ 理由 + 次の手（--overwrite）を言う。"""
    folder = tmp_path / "src"
    _book(folder / "a.xlsx", HDRS, [("J-1", "甲", 100)])
    out = tmp_path / "mine.xlsx"
    _book(out, ["大事な列"], [("消えては困る",)])
    p = _stack(folder, out)
    assert p.returncode == 7
    combined = p.stdout + p.stderr
    assert "--overwrite" in combined, f"次の手の開示が無い:\n{combined}"
    assert "mine.xlsx" in combined, "何が邪魔なのかの名指しが無い"


def test_multiple_subtotal_groups_do_not_false_alarm(tmp_path):
    """★ review#2 major: 閉じる検査が先頭からの累積和としか比べず、2 個目の小計で
       正しい表に偽 ⚠ を出した。区間和（直前の除外行から）か累積和のどちらかが
       合えば閉じる、が正しい形（総計は累積で・各小計は区間で閉じる）。"""
    folder = tmp_path / "src"
    _book(folder / "a.xlsx", HDRS,
          [("部署A-1", "x", 100), ("部署A-2", "x", 200), ("小計", None, 300),
           ("部署B-1", "y", 50), ("部署B-2", "y", 150), ("小計", None, 200),
           ("総計", None, 500)])
    out = tmp_path / "out.xlsx"
    p = _stack(folder, out)
    assert p.returncode == 0
    assert "≠" not in p.stdout, f"正しい表に偽の不一致 ⚠:\n{p.stdout}"
    assert p.stdout.count("500") >= 2, "Σ の両側（500）が無い"
    ws = openpyxl.load_workbook(out).active
    assert ws.max_row - 1 == 4, "積む行数が違う（明細 4 行のはず）"


def test_all_numeric_columns_are_reconciled_not_just_the_first(tmp_path):
    """★ review#3/#6 major: stack の Σ 照合・報告が最初の数値列 1 本だけだった。
       契約（このファイル冒頭）は『数値列ごとの Σ 両側表示』── 数量と金額の両方が出ること。"""
    folder = tmp_path / "src"
    _book(folder / "a.xlsx", ["注文ID", "数量", "金額"],
          [("J-1", 2, 100), ("J-2", 3, 200)])
    out = tmp_path / "out.xlsx"
    p = _stack(folder, out, "--json")
    assert p.returncode == 0, p.stderr[-500:]
    data = json.loads(p.stdout)
    assert "数量" in data["sums"] and "金額" in data["sums"], f"数値列ごとの Σ が無い: {list(data['sums'])}"
    assert data["sums"]["数量"] == {"source": 5, "output": 5}
    assert data["sums"]["金額"] == {"source": 300, "output": 300}
    p2 = _stack(folder, tmp_path / "out2.xlsx")
    assert "数量" in p2.stdout and "金額" in p2.stdout, "人間向けにも両列の Σ が要る"


def test_json_carries_closure_mismatches_for_automation(tmp_path):
    """★ review#4 major (conv=3): --json に閉じる検査の不一致が載らず、自動化経路だけ
       データ不整合が見えなかった。テキストの ⚠ と同じ情報を機械可読でも出す。"""
    folder = tmp_path / "src"
    _book(folder / "a.xlsx", HDRS,
          [("J-1", "甲", 600), ("J-2", "乙", 400), ("合計", None, 1200)])
    out = tmp_path / "out.xlsx"
    p = _stack(folder, out, "--json")
    assert p.returncode == 0, p.stderr[-500:]
    data = json.loads(p.stdout)
    mm = data.get("mismatches")
    assert mm, f"mismatches が JSON に無い: {list(data)}"
    assert any(m.get("excluded_value") == 1200 and m.get("adopted_sum") == 1000 for m in mm), mm


# ---- P1: 書き手の印の集合化（M2 前の契約改修・architect 致命2 の凍結）----


def _mark_creator(path, mark):
    wb = openpyxl.load_workbook(path)
    wb.properties.creator = mark
    wb.save(path)


def test_other_ailine_kind_output_is_gated_with_named_reason(tmp_path):
    """★ 致命2: 将来の ailine extract 等『別コマンドの出力』に stack --out を向けたら、
       無警告の作り直しではなく exit 7 + 名指し（ailine の別コマンドの出力です）。
       作り直してよいのは印が完全一致（ailine stack 自身の前回出力）の時だけ。"""
    folder = tmp_path / "src"
    _book(folder / "a.xlsx", HDRS, [("J-1", "甲", 100)])
    out = tmp_path / "extracted.xlsx"
    _book(out, HDRS + ["元ファイル", "元行"], [("X-1", "抽出結果", 999, "b.xlsx", 2)])
    _mark_creator(out, "ailine extract")
    sha = hashlib.sha256(out.read_bytes()).hexdigest()
    p = _stack(folder, out)
    assert p.returncode == 7, f"別コマンドの出力を黙って作り直した (exit={p.returncode})\n{p.stdout}"
    assert hashlib.sha256(out.read_bytes()).hexdigest() == sha
    assert "別のコマンド" in p.stdout or "ailine extract" in p.stdout, \
        f"別コマンドの出力である旨の名指しが無い:\n{p.stdout}"


def test_other_ailine_kind_output_inside_folder_is_excluded_from_inputs(tmp_path):
    """★ 致命2 の裏面: 入力フォルダ内の『ailine の別コマンドの出力』は、二重計上を防ぐため
       入力から除外 + 開示（is_own_output の marks 集合判定が V6 除外にも効くこと）。"""
    folder = tmp_path / "src"
    _book(folder / "a.xlsx", HDRS, [("J-1", "甲", 100)])
    other = folder / "extracted.xlsx"
    _book(other, HDRS + ["元ファイル", "元行"], [("X-1", "抽出結果", 999, "b.xlsx", 2)])
    _mark_creator(other, "ailine extract")
    out = tmp_path / "out.xlsx"
    p = _stack(folder, out)
    assert p.returncode == 0, p.stdout + p.stderr[-300:]
    ws = openpyxl.load_workbook(out).active
    assert ws.max_row - 1 == 1, "別コマンドの出力を入力に食って二重計上した"
    assert "除外" in p.stdout, "除外の開示が無い"


# ---- P2: シート引き当ての整合と開示（architect 致命5 前段・stack 側の開示）----


def test_sheet_fallback_is_disclosed_in_text_and_json(tmp_path):
    """★ 致命5 前段: 基準名のシートが無いソースは1枚目へ落ちる（従来どおり無警告では
       止まらない）が、その事実は人間向け報告に1行 + --json（files 配列内・sheet_fallbacks
       両方）で開示すること。"""
    folder = tmp_path / "src"
    base = folder / "a.xlsx"
    _book(base, HDRS, [("J-1", "甲", 100)])
    wb = openpyxl.load_workbook(base)
    wb.active.title = "明細"
    wb.save(base)
    # b.xlsx には『明細』シートが無い（1枚目『集計』しかない）── 落ちて照合を試みる。
    b = folder / "b.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "集計"
    ws.append(HDRS)
    ws.append(["J-2", "乙", 200])
    wb.save(b)
    out = tmp_path / "out.xlsx"
    p = _stack(folder, out)
    assert p.returncode == 0, p.stdout
    assert "明細" in p.stdout and "集計" in p.stdout and "b.xlsx" in p.stdout, \
        f"シート引き当てのフォールバック開示が無い:\n{p.stdout}"

    out2 = tmp_path / "out2.xlsx"
    p2 = _stack(folder, out2, "--json")
    assert p2.returncode == 0, p2.stdout
    data = json.loads(p2.stdout)
    fb = data.get("sheet_fallbacks")
    assert fb and any(f["name"] == "b.xlsx" and f["wanted"] == "明細" and f["used"] == "集計"
                       for f in fb), f"sheet_fallbacks が JSON に無い/不正: {fb}"
