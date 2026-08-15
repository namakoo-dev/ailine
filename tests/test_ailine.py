"""ailine の純粋ロジックの単体テスト（ollama / LibreOffice を要さない部分）。
   生成・適用の統合は実機（basrun_spike）で検証済み。ここは回帰用の土台。
"""
import argparse
import json
import os
import subprocess
import sys
import urllib.error
from pathlib import Path

import openpyxl
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import ailine


# --- コード抽出 / 署名 -----------------------------------------------------

def test_extract_bas_strips_markdown_fence():
    raw = "```basic\nOption Explicit\nSub Run(oDoc As Object)\nEnd Sub\n```"
    assert ailine.extract_bas(raw).startswith("Option Explicit")
    assert "```" not in ailine.extract_bas(raw)

def test_extract_bas_passthrough_without_fence():
    raw = "Sub Run(oDoc As Object)\nEnd Sub"
    assert ailine.extract_bas(raw) == raw

@pytest.mark.parametrize("code,ok", [
    ("Sub Run(oDoc As Object)\nEnd Sub", True),
    ("sub run( oDoc as object )", True),          # 大文字小文字・空白ゆらぎ
    ("Sub Run()\nEnd Sub", False),                # 引数なし
    ("Sub Other(oDoc As Object)", False),         # 別名
    ("' コメントだけ", False),
])
def test_valid_signature(code, ok):
    assert ailine.valid_signature(code) is ok


# --- 参照ライブラリ --------------------------------------------------------

def test_load_refs_bundles_examples():
    text = ailine.load_refs(ailine.DEFAULT_REFS)
    assert "Sub Run(oDoc As Object)" in text
    assert "参考" in text

def test_load_refs_missing_dir_is_empty(tmp_path):
    assert ailine.load_refs(tmp_path / "nope") == ""


# --- ヘルパ・ライブラリ（呼ぶだけ） ----------------------------------------

def test_load_helpers_catalog_and_files():
    catalog, files = ailine.load_helpers(ailine.DEFAULT_HELPERS)
    assert any(f.name.endswith(".bas") for f in files)
    assert "SortByColumn" in catalog
    assert "InsertBarChart" in catalog
    assert "Call" in catalog          # Call 形式で呼ばせる指示が入っている

def test_load_helpers_missing_dir(tmp_path):
    catalog, files = ailine.load_helpers(tmp_path / "nope")
    assert catalog == "" and files == []


# ★ 太字は native（StyleBold ヘルパが Basic で CharWeight+CharWeightAsian を当てる）。
#   openpyxl 後付けは撤去した。日本語太字は CharWeightAsian が要る点が要（実測）。
#   ここは Basic 側の実挙動なので純ロジック test では検証せず、通し試験＋描画で確認する。


# --- snapshot / 差分（no-op ガードの核） -----------------------------------

def _book(tmp_path, rows):
    p = tmp_path / "b.xlsx"
    wb = openpyxl.Workbook(); ws = wb.active
    for r in rows:
        ws.append(r)
    wb.save(p)
    return p

def test_diff_detects_value_change(tmp_path):
    p = _book(tmp_path, [["a", 1], ["b", 2]])
    before = ailine.snapshot(p)
    wb = openpyxl.load_workbook(p); wb.active.cell(1, 3, "new"); wb.save(p)
    after = ailine.snapshot(p)
    changed, lines = ailine.diff_snapshots(before, after)
    assert changed is True
    assert any("new" in ln for ln in lines)

def test_diff_noop_when_unchanged(tmp_path):
    p = _book(tmp_path, [["a", 1]])
    before = ailine.snapshot(p)
    after = ailine.snapshot(p)   # 何も変えない
    changed, lines = ailine.diff_snapshots(before, after)
    assert changed is False       # ← no-op を正しく no-op と判定
    assert lines == []

def test_diff_detects_new_sheet(tmp_path):
    p = _book(tmp_path, [["a", 1]])
    before = ailine.snapshot(p)
    wb = openpyxl.load_workbook(p); wb.create_sheet("集計"); wb.save(p)
    after = ailine.snapshot(p)
    changed, lines = ailine.diff_snapshots(before, after)
    assert changed is True
    assert any("集計" in ln for ln in lines)

def test_diff_detects_fill_only_change(tmp_path):
    # 値でなく背景色だけ変えても『変化した』と見えること（no-op 誤検出を防ぐ）
    from openpyxl.styles import PatternFill
    p = _book(tmp_path, [["a", 1]])
    before = ailine.snapshot(p)
    wb = openpyxl.load_workbook(p)
    wb.active.cell(1, 1).fill = PatternFill("solid", fgColor="FFCCCC")
    wb.save(p)
    after = ailine.snapshot(p)
    changed, _ = ailine.diff_snapshots(before, after)
    assert changed is True

def test_diff_detects_border_only_change(tmp_path):
    # 罫線だけの変更も検出すること（罫線ヘルパが no-op 誤判定されないため）
    from openpyxl.styles import Border, Side
    p = _book(tmp_path, [["a", 1]])
    before = ailine.snapshot(p)
    wb = openpyxl.load_workbook(p)
    thin = Side(style="thin")
    wb.active.cell(1, 1).border = Border(left=thin, right=thin, top=thin, bottom=thin)
    wb.save(p)
    after = ailine.snapshot(p)
    changed, _ = ailine.diff_snapshots(before, after)
    assert changed is True

def test_diff_detects_merge(tmp_path):
    p = _book(tmp_path, [["a", "b"]])
    before = ailine.snapshot(p)
    wb = openpyxl.load_workbook(p); wb.active.merge_cells("A1:B1"); wb.save(p)
    after = ailine.snapshot(p)
    changed, lines = ailine.diff_snapshots(before, after)
    assert changed is True
    assert any("結合" in ln for ln in lines)

def test_diff_detects_colwidth(tmp_path):
    p = _book(tmp_path, [["a", 1]])
    before = ailine.snapshot(p)
    wb = openpyxl.load_workbook(p); wb.active.column_dimensions["A"].width = 30; wb.save(p)
    after = ailine.snapshot(p)
    changed, lines = ailine.diff_snapshots(before, after)
    assert changed is True
    assert any("列幅" in ln for ln in lines)

def test_diff_detects_align_only_change(tmp_path):
    # 中央揃えだけの変更も検出すること（AlignCenter ヘルパが no-op 誤判定されないため）
    from openpyxl.styles import Alignment
    p = _book(tmp_path, [["a", 1]])
    before = ailine.snapshot(p)
    wb = openpyxl.load_workbook(p)
    wb.active.cell(1, 1).alignment = Alignment(horizontal="center")
    wb.save(p)
    after = ailine.snapshot(p)
    changed, _ = ailine.diff_snapshots(before, after)
    assert changed is True


# --- 差分見出し（P1: セル値変更が無見出しで続いていた不整合の修正） -----------

def test_diff_cell_change_has_own_heading(tmp_path):
    p = _book(tmp_path, [["a", 1], ["b", 2]])
    before = ailine.snapshot(p)
    wb = openpyxl.load_workbook(p); wb.active.cell(1, 3, "new"); wb.save(p)
    after = ailine.snapshot(p)
    changed, lines = ailine.diff_snapshots(before, after)
    assert changed is True
    heading_idx = next(i for i, ln in enumerate(lines) if ln.startswith("＊セル値変更:"))
    detail_idx = next(i for i, ln in enumerate(lines) if "new" in ln)
    assert heading_idx < detail_idx   # 見出し → 明細の順

def test_diff_cell_and_rowheight_each_get_own_heading(tmp_path):
    # 行高変更とセル値変更が両方あるとき、セル値変更が行高見出しの下に
    # 無見出しでぶら下がらず、自前の見出しを持つこと（修正前の不整合の再現）
    p = _book(tmp_path, [["a", 1], ["b", 2]])
    before = ailine.snapshot(p)
    wb = openpyxl.load_workbook(p)
    wb.active.row_dimensions[1].height = 30
    wb.active.cell(1, 3, "new")
    wb.save(p)
    after = ailine.snapshot(p)
    changed, lines = ailine.diff_snapshots(before, after)
    assert changed is True
    assert any(ln.startswith("＊行高変更:") for ln in lines)
    assert any(ln.startswith("＊セル値変更:") for ln in lines)


# --- ★ メッセージの条件（P1: 失敗/--dry でも無条件に出ていた不整合の修正） -----

def test_success_message_on_real_success():
    msg = ailine.success_message({"ok": True, "attempts": 1})
    assert msg is not None
    assert "no-op ガードは正しさを保証しない" in msg

def test_success_message_none_on_dry():
    assert ailine.success_message({"ok": True, "dry": True}) is None

def test_success_message_none_on_failure():
    assert ailine.success_message({"ok": False}) is None


# --- ollama エラー分類（P1: 404 なのに ollama serve を疑わせる誤ヒントの修正） ---

def test_ollama_generate_404_suggests_pull(monkeypatch):
    def fake_urlopen(req, timeout=300):
        raise urllib.error.HTTPError(url="http://x", code=404, msg="Not Found", hdrs=None, fp=None)
    monkeypatch.setattr(ailine.urllib.request, "urlopen", fake_urlopen)
    with pytest.raises(SystemExit) as exc:
        ailine.ollama_generate("qwen2.5-coder:7b", [{"role": "user", "content": "hi"}])
    msg = str(exc.value)
    assert "qwen2.5-coder:7b" in msg
    assert "pull" in msg
    assert "ollama serve" not in msg   # 接続不能の案内と混同しない

def test_ollama_generate_connection_refused_suggests_serve(monkeypatch):
    def fake_urlopen(req, timeout=300):
        raise urllib.error.URLError("connection refused")
    monkeypatch.setattr(ailine.urllib.request, "urlopen", fake_urlopen)
    with pytest.raises(SystemExit) as exc:
        ailine.ollama_generate("qwen2.5-coder:7b", [{"role": "user", "content": "hi"}])
    msg = str(exc.value)
    assert "ollama serve" in msg
    assert "pull" not in msg   # 404 の案内と混同しない

def test_ollama_generate_other_http_error_is_distinct(monkeypatch):
    # 404/接続不能のどちらの定型文にも紐付けない（誤誘導しない）
    def fake_urlopen(req, timeout=300):
        raise urllib.error.HTTPError(url="http://x", code=500, msg="Internal Error", hdrs=None, fp=None)
    monkeypatch.setattr(ailine.urllib.request, "urlopen", fake_urlopen)
    with pytest.raises(SystemExit) as exc:
        ailine.ollama_generate("qwen2.5-coder:7b", [{"role": "user", "content": "hi"}])
    msg = str(exc.value)
    assert "500" in msg
    assert "pull" not in msg
    assert "ollama serve" not in msg


# --- 文書の説明 ------------------------------------------------------------

def test_describe_book_lists_headers(tmp_path):
    p = _book(tmp_path, [["商品", "金額", "在庫"], ["りんご", 1200, 8]])
    desc = ailine.describe_book(p)
    assert "列0=商品" in desc
    assert "列1=金額" in desc
    assert "シート一覧" in desc


# --- ★ M1: 進捗表示（生成中の完全沈黙の解消） --------------------------------

def test_fmt_elapsed_formats_one_decimal():
    assert ailine._fmt_elapsed(12.34) == "(12.3s)"
    assert ailine._fmt_elapsed(0) == "(0.0s)"

def test_progress_start_writes_without_newline_to_stderr(capsys):
    ailine.progress_start("⏳ テスト中…")
    captured = capsys.readouterr()
    assert captured.err == "⏳ テスト中…"
    assert captured.out == ""    # ★ stdout は汚さない（--json 互換の要）

def test_progress_end_appends_elapsed_and_newline(capsys):
    t0 = ailine.progress_start("⏳ テスト中…")
    ailine.progress_end(t0)
    captured = capsys.readouterr()
    assert captured.err.startswith("⏳ テスト中…")
    assert captured.err.rstrip("\n").endswith("s)")
    assert captured.err.endswith("\n")


# --- ★ M1: 適用タイムアウト（暴走マクロ対策） --------------------------------

class _FakeProcTimeoutOnce:
    """1回目の communicate() は TimeoutExpired、2回目(kill 後の回収)は正常終了。"""
    pid = 4242
    returncode = None
    def __init__(self):
        self.calls = 0
    def communicate(self, timeout=None):
        self.calls += 1
        if self.calls == 1:
            raise subprocess.TimeoutExpired(cmd="x", timeout=timeout)
        return ("", "")

class _FakeProcReturns:
    """指定した (stdout, stderr, returncode) をそのまま返す固定 Popen 偽物。"""
    def __init__(self, out="", err="", returncode=0, pid=1):
        self._out, self._err, self.returncode, self.pid = out, err, returncode, pid
    def communicate(self, timeout=None):
        return (self._out, self._err)

def _fake_book(tmp_path):
    return tmp_path / "book.xlsx"

def test_basrun_apply_outer_timeout_kills_and_classifies_as_runtime_error(tmp_path, monkeypatch):
    fake = _FakeProcTimeoutOnce()
    monkeypatch.setattr(ailine.subprocess, "Popen", lambda *a, **k: fake)
    monkeypatch.setattr(ailine, "basrun_path", lambda: Path("dummy_basrun.py"))
    killed = {}
    monkeypatch.setattr(ailine, "_kill_process_tree", lambda pid: killed.setdefault("pid", pid))

    ok, err, raw = ailine.basrun_apply(_fake_book(tmp_path), "code", tmp_path, timeout=5.0)

    assert ok is False
    assert err == "実行時エラー: マクロが 5 秒で終了しない（無限ループの可能性）"
    assert killed["pid"] == 4242    # ★ PID 指定で kill（taskkill /IM の一括killはしない）

def test_basrun_apply_detects_basruns_own_internal_timeout_message(tmp_path, monkeypatch):
    # basrun.py 自身の内部タイムアウト（--timeout 転送先）が先に発火して非ゼロで返る場合も
    # 同じ「実行時エラー: マクロが N 秒で終了しない」に正規化されること。
    msg = "apply が 5 秒応答しなかった (BASRUN_APPLY_TIMEOUT/--timeout)。接続先の LibreOffice を終了させて中止した。"
    monkeypatch.setattr(ailine.subprocess, "Popen",
                        lambda *a, **k: _FakeProcReturns(err=msg, returncode=1))
    monkeypatch.setattr(ailine, "basrun_path", lambda: Path("dummy_basrun.py"))

    ok, err, raw = ailine.basrun_apply(_fake_book(tmp_path), "code", tmp_path, timeout=5.0)

    assert ok is False
    assert err == ailine._timeout_error_message(5.0)

def test_basrun_apply_generic_runtime_error_is_not_misclassified_as_timeout(tmp_path, monkeypatch):
    monkeypatch.setattr(ailine.subprocess, "Popen",
                        lambda *a, **k: _FakeProcReturns(err="何かの実行時エラー", returncode=1))
    monkeypatch.setattr(ailine, "basrun_path", lambda: Path("dummy_basrun.py"))

    ok, err, raw = ailine.basrun_apply(_fake_book(tmp_path), "code", tmp_path, timeout=5.0)

    assert ok is False
    assert "無限ループ" not in err
    assert "何かの実行時エラー" in err

def test_basrun_apply_success_path_unaffected(tmp_path, monkeypatch):
    monkeypatch.setattr(ailine.subprocess, "Popen",
                        lambda *a, **k: _FakeProcReturns(out="applied", returncode=0))
    monkeypatch.setattr(ailine, "basrun_path", lambda: Path("dummy_basrun.py"))

    ok, err, raw = ailine.basrun_apply(_fake_book(tmp_path), "code", tmp_path, timeout=5.0)

    assert ok is True
    assert err is None

def test_basrun_apply_includes_timeout_flag_when_enabled(tmp_path, monkeypatch):
    captured = {}
    def fake_popen(cmd, **kw):
        captured["cmd"] = cmd
        return _FakeProcReturns(returncode=0)
    monkeypatch.setattr(ailine.subprocess, "Popen", fake_popen)
    monkeypatch.setattr(ailine, "basrun_path", lambda: Path("dummy_basrun.py"))

    ailine.basrun_apply(_fake_book(tmp_path), "code", tmp_path, timeout=42.0)

    assert "--timeout" in captured["cmd"]
    assert "42.0" in captured["cmd"]

def test_basrun_apply_omits_timeout_flag_when_disabled(tmp_path, monkeypatch):
    # --timeout 0 相当（呼び出し側が None を渡す＝旧挙動の無制限）
    captured = {}
    def fake_popen(cmd, **kw):
        captured["cmd"] = cmd
        return _FakeProcReturns(returncode=0)
    monkeypatch.setattr(ailine.subprocess, "Popen", fake_popen)
    monkeypatch.setattr(ailine, "basrun_path", lambda: Path("dummy_basrun.py"))

    ailine.basrun_apply(_fake_book(tmp_path), "code", tmp_path, timeout=None)

    assert "--timeout" not in captured["cmd"]


# --- ★ M1: ailine doctor（セットアップ診断） ---------------------------------

def test_check_python_version_ok_on_current_interpreter():
    ok, detail = ailine._check_python_version()
    assert ok is True
    assert detail == ""

def test_check_python_version_fails_on_old(monkeypatch):
    monkeypatch.setattr(ailine.sys, "version_info", (3, 9, 0))
    ok, detail = ailine._check_python_version()
    assert ok is False
    assert "3.10" in detail

def test_check_openpyxl_ok():
    ok, detail = ailine._check_openpyxl()
    assert ok is True and detail == ""

class _FakeUrlopenCtx:
    def __init__(self, payload):
        self._payload = payload
    def __enter__(self):
        return self
    def __exit__(self, *a):
        return False

def test_check_ollama_reachable_ok(monkeypatch):
    monkeypatch.setattr(ailine.urllib.request, "urlopen",
                        lambda req, timeout=3.0: _FakeUrlopenCtx({}))
    monkeypatch.setattr(ailine.json, "load", lambda f: {})
    ok, detail = ailine._check_ollama_reachable()
    assert ok is True and detail == ""

def test_check_ollama_reachable_fails_suggests_serve(monkeypatch):
    def raiser(req, timeout=3.0):
        raise OSError("接続を拒否されました")
    monkeypatch.setattr(ailine.urllib.request, "urlopen", raiser)
    ok, detail = ailine._check_ollama_reachable()
    assert ok is False
    assert "ollama serve" in detail

def test_check_model_available_found(monkeypatch):
    monkeypatch.setattr(ailine.urllib.request, "urlopen",
                        lambda req, timeout=3.0: _FakeUrlopenCtx(None))
    monkeypatch.setattr(ailine.json, "load", lambda f: {"models": [{"name": "qwen2.5-coder:7b"}]})
    ok, detail = ailine._check_model_available("qwen2.5-coder:7b")
    assert ok is True and detail == ""

def test_check_model_available_missing_suggests_pull(monkeypatch):
    monkeypatch.setattr(ailine.urllib.request, "urlopen",
                        lambda req, timeout=3.0: _FakeUrlopenCtx(None))
    monkeypatch.setattr(ailine.json, "load", lambda f: {"models": []})
    ok, detail = ailine._check_model_available("missing:1b")
    assert ok is False
    assert "ollama pull missing:1b" in detail

def test_check_model_available_unreachable(monkeypatch):
    def raiser(req, timeout=3.0):
        raise OSError("x")
    monkeypatch.setattr(ailine.urllib.request, "urlopen", raiser)
    ok, detail = ailine._check_model_available("m")
    assert ok is False

def test_check_libreoffice_missing_basrun(monkeypatch):
    monkeypatch.setattr(ailine, "_find_basrun_path", lambda: None)
    ok, detail = ailine._check_libreoffice()
    assert ok is False
    assert "basrun.py" in detail

def test_check_libreoffice_ok(monkeypatch, tmp_path):
    fake_path = tmp_path / "basrun.py"
    fake_path.write_text("# dummy", encoding="utf-8")
    monkeypatch.setattr(ailine, "_find_basrun_path", lambda: fake_path)
    class _FakeMod:
        @staticmethod
        def office_dir():
            return Path("C:/Program Files/LibreOffice/program")
    monkeypatch.setattr(ailine, "_load_module_from_path", lambda p, name: _FakeMod())
    ok, detail = ailine._check_libreoffice()
    assert ok is True
    assert "LibreOffice" in detail

def test_check_libreoffice_not_found(monkeypatch, tmp_path):
    fake_path = tmp_path / "basrun.py"
    fake_path.write_text("# dummy", encoding="utf-8")
    monkeypatch.setattr(ailine, "_find_basrun_path", lambda: fake_path)
    class _FakeMod:
        @staticmethod
        def office_dir():
            raise SystemExit("LibreOffice が見つからない")
    monkeypatch.setattr(ailine, "_load_module_from_path", lambda p, name: _FakeMod())
    ok, detail = ailine._check_libreoffice()
    assert ok is False

def test_check_basrun_found(monkeypatch, tmp_path):
    p = tmp_path / "basrun.py"
    monkeypatch.setattr(ailine, "_find_basrun_path", lambda: p)
    ok, detail = ailine._check_basrun()
    assert ok is True

def test_check_basrun_missing_gives_clone_hint(monkeypatch):
    monkeypatch.setattr(ailine, "_find_basrun_path", lambda: None)
    ok, detail = ailine._check_basrun()
    assert ok is False
    assert "clone" in detail

def test_check_demo_dir_ok_on_real_repo():
    ok, detail = ailine._check_demo_dir()
    assert ok is True   # このリポジトリの demo/ には .xlsx サンプルが同梱されている

def test_check_demo_dir_missing(monkeypatch, tmp_path):
    monkeypatch.setattr(ailine, "HERE", tmp_path)
    ok, detail = ailine._check_demo_dir()
    assert ok is False

def test_format_doctor_report_all_ok():
    text, all_ok = ailine.format_doctor_report([("a", True, ""), ("b", True, "detail")])
    assert all_ok is True
    assert "✓ a" in text
    assert "✓ b (detail)" in text

def test_format_doctor_report_has_failure_and_fix_hint():
    text, all_ok = ailine.format_doctor_report([("a", True, ""), ("b", False, "直せ")])
    assert all_ok is False
    assert "× b — 直せ" in text

def test_doctor_checks_returns_seven_items():
    # ①python ②openpyxl ③ollama ④モデル ⑤LibreOffice ⑥basrun.py ⑦demo/
    results = ailine.doctor_checks()
    assert len(results) == 7
    for name, ok, detail in results:
        assert isinstance(name, str)
        assert isinstance(ok, bool)


# --- ★ W8a 項目5: doctor の事務向け一行説明（7項目全部・翻訳表示） -----------------

def test_doctor_business_notes_cover_all_seven_real_check_names():
    names = [
        "python 3.10+", "openpyxl", f"ollama 到達 ({ailine.OLLAMA})",
        f"モデル '{ailine.DEFAULT_MODEL}'", "LibreOffice", "basrun.py", "demo/",
    ]
    for name in names:
        assert ailine._doctor_business_note(name) is not None

def test_format_doctor_report_shows_business_note_for_ollama():
    text, _ = ailine.format_doctor_report([(f"ollama 到達 ({ailine.OLLAMA})", True, "")])
    assert "AI エンジン (ollama) に接続できています" in text

def test_format_doctor_report_no_business_note_for_unknown_dummy_name():
    # ★ 既存テスト(test_format_doctor_report_all_ok 等)のダミー名 "a"/"b" は
    #   どの実プレフィックスにも一致しないため、従来どおり内部名のみ表示する。
    text, _ = ailine.format_doctor_report([("a", True, "")])
    assert text == "✓ a"


# --- ★ M1: 実行履歴（最小版） ------------------------------------------------

def test_build_history_entry_shape_and_truncation():
    result = {"ok": True, "attempts": 2, "changes": ["a", "b", "c", "d"], "out": "x.xlsx"}
    e = ailine.build_history_entry(result, Path("book.xlsx"), "タスク", "モデル", "none")
    assert e["ok"] is True
    assert e["attempts"] == 2
    assert e["changes"] == ["a", "b", "c"]      # 先頭3件のみ
    assert e["failure_kind"] == "none"
    assert e["book"] == str(Path("book.xlsx"))
    assert e["task"] == "タスク"
    assert e["model"] == "モデル"
    assert "ts" in e and "T" in e["ts"]          # ISO 形式

def test_append_and_read_history_roundtrip_newest_first(tmp_path):
    p = tmp_path / "history.jsonl"
    e1 = ailine.build_history_entry({"ok": True, "attempts": 1}, Path("a.xlsx"), "t1", "m", "none")
    e2 = ailine.build_history_entry({"ok": False, "attempts": 3}, Path("b.xlsx"), "t2", "m", "noop")
    ailine.append_history(e1, path=p)
    ailine.append_history(e2, path=p)
    entries = ailine.read_history(path=p, max_n=10)
    assert len(entries) == 2
    assert entries[0]["task"] == "t2"   # 新しい順
    assert entries[1]["task"] == "t1"

def test_read_history_missing_file_returns_empty(tmp_path):
    assert ailine.read_history(path=tmp_path / "nope.jsonl") == []

def test_read_history_skips_broken_lines(tmp_path):
    p = tmp_path / "history.jsonl"
    p.write_text('{"task": "ok1"}\nnot json\n{"task": "ok2"}\n', encoding="utf-8")
    entries = ailine.read_history(path=p, max_n=10)
    assert [e["task"] for e in entries] == ["ok2", "ok1"]

def test_read_history_respects_max_n(tmp_path):
    p = tmp_path / "history.jsonl"
    p.write_text("\n".join(f'{{"task": "t{i}"}}' for i in range(5)) + "\n", encoding="utf-8")
    entries = ailine.read_history(path=p, max_n=2)
    assert len(entries) == 2
    assert entries[0]["task"] == "t4"

def test_append_history_raises_when_path_unwritable(tmp_path):
    # 親のつもりの場所が実はファイル → mkdir が失敗する。ここは例外を投げてよい
    # （run 本体を落とさない責務は呼び出し側 cmd_run の try/except が持つ設計）。
    blocker = tmp_path / "blocker"
    blocker.write_text("x", encoding="utf-8")
    bad_path = blocker / "sub" / "history.jsonl"
    with pytest.raises(Exception):
        ailine.append_history({"a": 1}, path=bad_path)

def test_format_history_table_empty_says_none_yet():
    assert ailine.format_history_table([]) == "履歴はまだ無い"

def test_format_history_table_shows_failure_kind_tag():
    entries = [{"ts": "2026-01-01T00:00:00+00:00", "ok": False, "attempts": 3,
                "model": "m", "book": "b.xlsx", "task": "t", "failure_kind": "noop"}]
    text = ailine.format_history_table(entries)
    assert "[noop]" in text
    assert "×" in text

def test_format_history_table_success_has_no_kind_tag():
    entries = [{"ts": "2026-01-01T00:00:00+00:00", "ok": True, "attempts": 1,
                "model": "m", "book": "b.xlsx", "task": "t", "failure_kind": "none"}]
    text = ailine.format_history_table(entries)
    assert "[none]" not in text
    assert "✓" in text


# --- ★ W8a 項目1: dry(下見) と実適用の履歴区別 ---------------------------------

def test_build_history_entry_dry_field_defaults_false():
    e = ailine.build_history_entry({"ok": True, "attempts": 1}, Path("a.xlsx"), "t", "m", "none")
    assert e["dry"] is False

def test_build_history_entry_dry_field_true_when_result_dry():
    e = ailine.build_history_entry({"ok": True, "dry": True}, Path("a.xlsx"), "t", "m", "none")
    assert e["dry"] is True

def test_format_history_table_marks_dry_rows_with_kensaku_label():
    # 「下見」= 事務の言葉。「dry-run」は表示に出さない。
    entries = [{"ts": "2026-01-01T00:00:00+00:00", "ok": True, "attempts": 1, "dry": True,
                "model": "m", "book": "b.xlsx", "task": "t", "failure_kind": "none"}]
    text = ailine.format_history_table(entries)
    assert "(下見)" in text
    assert "dry" not in text.lower()

def test_format_history_table_old_rows_without_dry_key_read_as_applied():
    # ★ 後方互換: 旧 history.jsonl の行（"dry" キーが無い）は実適用扱いのまま読める
    #   （dict.get("dry", False) で False にフォールバック）。
    entries = [{"ts": "2026-01-01T00:00:00+00:00", "ok": True, "attempts": 1,
                "model": "m", "book": "b.xlsx", "task": "t", "failure_kind": "none"}]
    text = ailine.format_history_table(entries)
    assert "(下見)" not in text

def test_cmd_run_dry_survives_history_write_failure(tmp_path, monkeypatch, capsys):
    # ★ 履歴の書き込み失敗で run 本体を落とさない（try で包み WARN のみ）ことを
    #   --dry（ollama 生成だけで LibreOffice を要さない）経路で確認する。
    book = _book(tmp_path, [["a", 1]])
    # ★ M2b: run はまず translate_task を呼ぶ。ここは自由生成経路の回帰なので FREEFORM に固定する。
    monkeypatch.setattr(ailine, "translate_task",
                        lambda model, task, book_meta, temperature=0.1: {"op": "FREEFORM", "args": {}})
    monkeypatch.setattr(ailine, "ollama_generate",
                        lambda model, msgs, temperature=0.2:
                        "Sub Run(oDoc As Object)\nEnd Sub")
    def boom(entry, path=None):
        raise OSError("書き込み失敗（テスト用）")
    monkeypatch.setattr(ailine, "append_history", boom)

    ns = argparse.Namespace(
        book=str(book), task="テスト", model="qwen2.5-coder:7b",
        refs=None, helpers=None, repair=0, temperature=0.2,
        dry=True, inplace=False, json=False, timeout=180.0, ask=False)

    rc = ailine.cmd_run(ns)
    captured = capsys.readouterr()
    assert rc == 0   # ★ 履歴書き込みが例外を投げても run 自体は成功のまま
    assert "WARN" in captured.err


# --- ★ M1: 差分の人間化（生 tuple 漏れの解消） -------------------------------

def test_describe_cell_change_value_only_shows_only_value_label():
    before = None   # snapshot() が完全既定セルを辞書に持たない状態を模す
    after = ("りんご", "General", None, False, None, None)
    desc = ailine.describe_cell_change(before, after)
    assert desc == "値 (空)→'りんご'"
    assert "数値書式" not in desc   # ★ 既定 tuple を正しく使わないと偽の差分が出る罠の回帰

def test_describe_cell_change_multiple_fields_listed():
    before = ("りんご", "General", None, False, None, None)
    after = ("リンゴ", "General", None, True, None, None)
    desc = ailine.describe_cell_change(before, after)
    assert "値 'りんご'→'リンゴ'" in desc
    assert "太字 False→True" in desc

def test_describe_cell_change_no_diff_returns_placeholder():
    t = ("a", "General", None, False, None, None)
    assert ailine.describe_cell_change(t, t) == "(差分なし)"

def test_cell_ref_uses_a1_notation():
    assert ailine._cell_ref(1, 1) == "A1"
    assert ailine._cell_ref(2, 2) == "B2"

def test_diff_snapshots_new_value_cell_uses_a1_ref_and_no_raw_tuple(tmp_path):
    p = _book(tmp_path, [["a", 1], ["b", 2]])
    before = ailine.snapshot(p)
    wb = openpyxl.load_workbook(p); wb.active.cell(1, 3, "new"); wb.save(p)
    after = ailine.snapshot(p)
    changed, lines = ailine.diff_snapshots(before, after)
    detail = next(ln for ln in lines if "new" in ln)
    assert detail.strip().startswith("C1:")     # row=1,col=3 → A1形式で C1
    assert "値" in detail
    assert "None" not in detail                 # 生 tuple の名残（Noneの生表示）が無い
    assert "数値書式" not in detail              # 値だけの変更で偽の書式差分が出ない


# --- ★ M2a: 疑わしい変化の機械検出 ------------------------------------------

def test_detect_ghost_data_single_cell_outside_original_range(tmp_path):
    p = _book(tmp_path, [["a", 1], ["b", 2]])   # 使用範囲は A1:B2
    before = ailine.snapshot(p)
    wb = openpyxl.load_workbook(p)
    wb.active.cell(row=2, column=26, value="ghost")   # Z2（範囲外）
    wb.save(p)
    after = ailine.snapshot(p)
    msg = ailine.detect_ghost_data(before, after)
    assert msg is not None
    assert "Z2" in msg
    assert "★ 疑わしい" in msg

def test_detect_ghost_data_range_of_outside_cells(tmp_path):
    p = _book(tmp_path, [["a", 1], ["b", 2]])
    before = ailine.snapshot(p)
    wb = openpyxl.load_workbook(p)
    for r in (2, 3, 4):
        wb.active.cell(row=r, column=26, value=0)
    wb.save(p)
    after = ailine.snapshot(p)
    msg = ailine.detect_ghost_data(before, after)
    assert msg is not None
    assert "Z2:Z4" in msg

def test_detect_ghost_data_none_when_any_change_is_in_range(tmp_path):
    p = _book(tmp_path, [["a", 1], ["b", 2]])
    before = ailine.snapshot(p)
    wb = openpyxl.load_workbook(p)
    wb.active.cell(row=2, column=26, value="ghost")   # 範囲外
    wb.active.cell(row=1, column=1, value="変更")      # 範囲内も混ざる
    wb.save(p)
    after = ailine.snapshot(p)
    assert ailine.detect_ghost_data(before, after) is None

def test_detect_ghost_data_none_when_no_change():
    snap = {"cells": {}}
    assert ailine.detect_ghost_data(snap, snap) is None


def test_detect_uniform_fill_flags_same_value_into_blank_cells(tmp_path):
    p = _book(tmp_path, [["a", 1], ["b", 2]])
    before = ailine.snapshot(p)
    wb = openpyxl.load_workbook(p)
    wb.active.cell(row=1, column=3, value=0)   # C列は元々空欄
    wb.active.cell(row=2, column=3, value=0)
    wb.save(p)
    after = ailine.snapshot(p)
    msg = ailine.detect_uniform_fill(before, after)
    assert msg is not None
    assert "値 0 × 2 セル" in msg
    assert "★ 疑わしい" in msg

def test_detect_uniform_fill_none_when_overwriting_existing_value(tmp_path):
    p = _book(tmp_path, [["a", 1], ["b", 2]])
    before = ailine.snapshot(p)
    wb = openpyxl.load_workbook(p)
    wb.active.cell(row=1, column=1, value=0)   # 既存値の上書き（空欄からではない）
    wb.active.cell(row=1, column=3, value=0)   # こちらは空欄から
    wb.save(p)
    after = ailine.snapshot(p)
    assert ailine.detect_uniform_fill(before, after) is None

def test_detect_uniform_fill_none_when_values_differ(tmp_path):
    p = _book(tmp_path, [["a", 1], ["b", 2]])
    before = ailine.snapshot(p)
    wb = openpyxl.load_workbook(p)
    wb.active.cell(row=1, column=3, value=0)
    wb.active.cell(row=2, column=3, value=1)   # 値がバラバラ
    wb.save(p)
    after = ailine.snapshot(p)
    assert ailine.detect_uniform_fill(before, after) is None

def test_detect_uniform_fill_fires_when_mixed_with_border_and_center_align(tmp_path):
    # ★ M2c 回帰: 冷間再監査3回目の実測ケース。罫線+中央揃え(値は不変)と、空欄への
    #   3セル0埋め(値が変わる)が混在すると、旧実装は書式のみの変更セルまで一様性判定に
    #   巻き込んで見逃していた（値変更の部分集合だけで評価するよう修正）。
    from openpyxl.styles import Border, Side, Alignment
    p = _book(tmp_path, [["商品", "金額", "備考"], ["a", 100, None], ["b", 200, None], ["c", 300, None]])
    before = ailine.snapshot(p)
    wb = openpyxl.load_workbook(p)
    ws = wb.active
    thin = Border(left=Side(style="thin"), right=Side(style="thin"),
                   top=Side(style="thin"), bottom=Side(style="thin"))
    center = Alignment(horizontal="center")
    # 罫線+中央揃え(値は変えない・書式のみ) — A1:B4 全体
    for r in range(1, 5):
        for c in (1, 2):
            cell = ws.cell(row=r, column=c)
            cell.border = thin
            cell.alignment = center
    # 空欄だった備考列(C列)に0を一様書き込み(値が変わる)
    ws.cell(row=2, column=3, value=0)
    ws.cell(row=3, column=3, value=0)
    ws.cell(row=4, column=3, value=0)
    wb.save(p)
    after = ailine.snapshot(p)
    msg = ailine.detect_uniform_fill(before, after)
    assert msg is not None
    assert "値 0 × 3 セル" in msg

def test_detect_ghost_data_ignores_format_only_changes(tmp_path):
    # 値は変わらず書式だけ変わったセルは幽霊データ判定の対象外（誤検知を増やさない）。
    from openpyxl.styles import Alignment
    p = _book(tmp_path, [["a", 1], ["b", 2]])
    before = ailine.snapshot(p)
    wb = openpyxl.load_workbook(p)
    wb.active.cell(row=1, column=1).alignment = Alignment(horizontal="center")   # 範囲内・書式のみ
    wb.save(p)
    after = ailine.snapshot(p)
    assert ailine.detect_ghost_data(before, after) is None


# --- ★ M2a: 件数の突き合わせ -------------------------------------------------

def test_count_reconciliation_reports_data_vs_changed_rows(tmp_path):
    p = _book(tmp_path, [["商品", "金額"], ["りんご", 100], ["バナナ", 200], ["みかん", 300]])
    before = ailine.snapshot(p)
    wb = openpyxl.load_workbook(p)
    wb.active.cell(row=2, column=2, value=999)   # りんご行のみ変更
    wb.active.cell(row=3, column=2, value=999)   # バナナ行のみ変更（みかん行は無変更）
    wb.save(p)
    after = ailine.snapshot(p)
    msg = ailine.count_reconciliation(before, after)
    assert msg == "列 B: データ 3 行のうち 2 行を変更（1 行は未変更）"

def test_count_reconciliation_none_when_multiple_columns_changed(tmp_path):
    p = _book(tmp_path, [["商品", "金額"], ["りんご", 100], ["バナナ", 200]])
    before = ailine.snapshot(p)
    wb = openpyxl.load_workbook(p)
    wb.active.cell(row=2, column=1, value="リンゴ")
    wb.active.cell(row=2, column=2, value=999)
    wb.save(p)
    after = ailine.snapshot(p)
    assert ailine.count_reconciliation(before, after) is None

# --- ★ W8a 項目2: 件数突合の算数バグ（分子に見出し行が混入していた） ------------------

def test_count_reconciliation_excludes_header_row_from_numerator(tmp_path):
    # ★ 実測(e2e_work/w3_e2e3_log.txt): 新規列を作り見出し(F1)にもデータ全5行(F2-F6)にも
    #   書く COMPUTE_COLUMN で「データ5行のうち6行を変更」という算数が壊れた表示になって
    #   いた。見出し行は分子から除外し、別語「＋見出し行」で添える。
    p = _book(tmp_path, [["商品", "金額"], ["a", 100], ["b", 200], ["c", 300], ["d", 400], ["e", 500]])
    before = ailine.snapshot(p)
    wb = openpyxl.load_workbook(p)
    ws = wb.active
    ws.cell(row=1, column=3, value="金額(税込)")   # 新規列の見出し
    for r in range(2, 7):
        ws.cell(row=r, column=3, value=f"=B{r}*1.1")
    wb.save(p)
    after = ailine.snapshot(p)
    msg = ailine.count_reconciliation(before, after)
    assert msg == "列 C: データ 5 行のうち 5 行を変更（0 行は未変更）＋見出し行"

def test_count_reconciliation_no_header_suffix_when_header_untouched(tmp_path):
    # 見出し行に触れていない既存の回帰(りんご欠落型)には「＋見出し行」を付けない。
    p = _book(tmp_path, [["商品", "金額"], ["りんご", 100], ["バナナ", 200], ["みかん", 300]])
    before = ailine.snapshot(p)
    wb = openpyxl.load_workbook(p)
    wb.active.cell(row=2, column=2, value=999)
    wb.save(p)
    after = ailine.snapshot(p)
    msg = ailine.count_reconciliation(before, after)
    assert "＋見出し行" not in msg

def test_count_reconciliation_invariant_denominator_never_less_than_numerator(tmp_path):
    # ★ 不変条件: 「データ N 行のうち M 行を変更」で常に N >= M（分子=データ行のみ）。
    p = _book(tmp_path, [["商品", "金額"], ["a", 100], ["b", 200], ["c", 300]])
    before = ailine.snapshot(p)
    wb = openpyxl.load_workbook(p)
    ws = wb.active
    ws.cell(row=1, column=2, value="金額(税込)")   # 見出しも変更
    ws.cell(row=2, column=2, value=999)
    wb.save(p)
    after = ailine.snapshot(p)
    msg = ailine.count_reconciliation(before, after)
    import re as _re
    m = _re.search(r"データ (\d+) 行のうち (\d+) 行を変更", msg)
    data_rows, changed_rows = int(m.group(1)), int(m.group(2))
    assert data_rows >= changed_rows
    assert "＋見出し行" in msg


# --- ★ M2a: 依頼文と変更範囲の重なりチェック ----------------------------------

def test_extract_task_mentions_column_letter_and_number_and_row():
    m = ailine.extract_task_mentions("列Zの値を全部2倍にする", ["Sheet"])
    assert m["cols"] == {26}
    m2 = ailine.extract_task_mentions("Z列の値を確認", ["Sheet"])
    assert m2["cols"] == {26}
    m3 = ailine.extract_task_mentions("列3を直して行5も見て", ["Sheet"])
    assert m3["cols"] == set()          # 数字表記は曖昧なので cols に断定で入れない
    assert m3["digit_cols"] == {3}      # 生の値のまま保持し照合側で両解釈
    assert m3["rows"] == {5}

def test_extract_task_mentions_only_real_sheet_names_match():
    m = ailine.extract_task_mentions("集計シートを見て", ["Sheet", "集計"])
    assert m["sheets"] == {"集計"}
    m2 = ailine.extract_task_mentions("集計シートを見て", ["Sheet"])   # 実在しない
    assert m2["sheets"] == set()

def test_extract_task_mentions_empty_when_no_mentions():
    m = ailine.extract_task_mentions("いい感じにして", ["Sheet"])
    assert m == {"cols": set(), "digit_cols": set(), "rows": set(), "sheets": set()}


def test_mention_digit_col_accepts_both_zero_and_one_based(tmp_path):
    """再監査 2 回目の誤警報の再発防止: 『在庫(列2)』= 0 起点で C 列(1 起点 3) を
    正しく変更したのに B 列不在の★が出た。数字表記は両解釈のどちらかが触られて
    いれば沈黙する。"""
    p = _book(tmp_path, [["商品", "金額", "在庫"], ["りんご", 100, 8]])
    before = ailine.snapshot(p)
    wb = openpyxl.load_workbook(p)
    wb.active.cell(row=2, column=3, value=16)   # C 列 = 0 起点の列2 を変更
    wb.save(p)
    after = ailine.snapshot(p)
    mentions = ailine.extract_task_mentions("在庫(列2)の値を2倍にする", before["sheets"])
    assert ailine.mention_overlap_advisory(mentions, before, after) == []   # 誤警報なし

    # 両解釈とも外れている場合だけ警告し、警告文は数字のまま
    mentions9 = ailine.extract_task_mentions("列9を2倍にする", before["sheets"])
    lines = ailine.mention_overlap_advisory(mentions9, before, after)
    assert any("『列9』" in ln for ln in lines)

def test_mention_overlap_flags_unmatched_column(tmp_path):
    p = _book(tmp_path, [["商品", "金額"], ["りんご", 100], ["バナナ", 200]])
    before = ailine.snapshot(p)
    wb = openpyxl.load_workbook(p)
    wb.active.cell(row=2, column=2, value=999)   # B列だけを変更
    wb.save(p)
    after = ailine.snapshot(p)
    mentions = ailine.extract_task_mentions("列Zを2倍にする", before["sheets"])
    lines = ailine.mention_overlap_advisory(mentions, before, after)
    assert any("列Z" in ln and "★" in ln for ln in lines)

def test_mention_overlap_silent_when_mentioned_column_was_changed(tmp_path):
    p = _book(tmp_path, [["商品", "金額"], ["りんご", 100]])
    before = ailine.snapshot(p)
    wb = openpyxl.load_workbook(p)
    wb.active.cell(row=2, column=2, value=999)
    wb.save(p)
    after = ailine.snapshot(p)
    mentions = ailine.extract_task_mentions("B列を確認", before["sheets"])
    assert ailine.mention_overlap_advisory(mentions, before, after) == []

def test_mention_overlap_flags_untouched_real_sheet(tmp_path):
    p = tmp_path / "b.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    for row in [["商品", "金額"], ["りんご", 100]]:
        ws.append(row)
    wb.create_sheet("集計")
    wb.save(p)
    before = ailine.snapshot(p)
    wb2 = openpyxl.load_workbook(p)
    wb2.active.cell(row=2, column=2, value=999)   # Sheet 側だけ変更、集計シートは無変更
    wb2.save(p)
    after = ailine.snapshot(p)
    mentions = ailine.extract_task_mentions("集計シートも直して", before["sheets"])
    lines = ailine.mention_overlap_advisory(mentions, before, after)
    assert any("集計" in ln for ln in lines)

def test_mention_overlap_empty_when_no_mentions(tmp_path):
    p = _book(tmp_path, [["a", 1]])
    snap = ailine.snapshot(p)
    assert ailine.mention_overlap_advisory({"cols": set(), "rows": set(), "sheets": set()}, snap, snap) == []


# --- ★ W8a 項目4: 率リテラルの機械スキャン（判断棚から昇格） -----------------------

def test_scan_rate_literals_fires_when_rate_not_in_task():
    code = "Sub Run(oDoc As Object)\n  x = 100 * 1.08\nEnd Sub"
    out = ailine.scan_rate_literals(code, "税込み合計を出して")
    assert len(out) == 1
    assert "1.08" in out[0]
    assert "検算してください" in out[0]

def test_scan_rate_literals_silent_when_task_mentions_percentage():
    code = "Sub Run(oDoc As Object)\n  x = 100 * 1.08\nEnd Sub"
    out = ailine.scan_rate_literals(code, "消費税8%込みの合計を出して")
    assert out == []

def test_scan_rate_literals_silent_when_vocab_has_matching_rate():
    code = "Sub Run(oDoc As Object)\n  x = 100 * 1.1\nEnd Sub"
    out = ailine.scan_rate_literals(code, "税込み合計を出して", vocab={"消費税": 1.1})
    assert out == []

def test_scan_rate_literals_ignores_comment_lines():
    code = "Sub Run(oDoc As Object)\n  ' 参考値 1.08 は例\n  x = 1\nEnd Sub"
    out = ailine.scan_rate_literals(code, "合計を出して")
    assert out == []

def test_scan_rate_literals_ignores_integers():
    code = "Sub Run(oDoc As Object)\n  x = 100 * 5\nEnd Sub"
    out = ailine.scan_rate_literals(code, "合計を出して")
    assert out == []

def test_scan_rate_literals_ignores_out_of_range_decimals():
    code = "Sub Run(oDoc As Object)\n  x = 3.14\nEnd Sub"
    out = ailine.scan_rate_literals(code, "合計を出して")
    assert out == []

def test_scan_rate_literals_dedupes_repeated_literal():
    code = "Sub Run(oDoc As Object)\n  x = 100 * 1.1\n  y = 200 * 1.1\nEnd Sub"
    out = ailine.scan_rate_literals(code, "合計を出して")
    assert len(out) == 1

def test_scan_rate_literals_boundary_values_inclusive():
    code = "Sub Run(oDoc As Object)\n  x = 100 * 0.05\n  y = 200 * 1.2\nEnd Sub"
    out = ailine.scan_rate_literals(code, "合計を出して")
    assert len(out) == 2


# --- ★ M2a: 生成コードの切断検出 ---------------------------------------------

@pytest.mark.parametrize("code,truncated", [
    ("Sub Run(oDoc As Object)\nEnd Sub", False),
    ("Sub Run(oDoc As Object)\n  oDoc.Sheets.getByIndex(0)\nEnd Sub", False),
    ("", True),
    ("   \n  ", True),
    ("Sub Run(oDoc As Object)\n  oDoc.Sheets.getByIndex(0", True),          # 開き括弧で途切れ
    ("Sub Run(oDoc As Object)\n  Dim oSheet As Object\n  oSheet", True),    # 識別子で途切れ
    ("Sub Run(oDoc As Object)\nEnd Sub\nDim leftover", True),              # End Sub の後に続きがある
])
def test_is_truncated_code(code, truncated):
    assert ailine.is_truncated_code(code) is truncated


# --- ★ M2a: 実行時エラー表示の整形 --------------------------------------------

def test_short_error_summary_returns_last_line_of_traceback():
    tb = ("Traceback (most recent call last):\n"
          '  File "basrun.py", line 10, in <module>\n'
          "    raise ValueError('bad state')\n"
          "ValueError: bad state")
    assert ailine.short_error_summary(tb) == "ValueError: bad state"

def test_short_error_summary_single_line():
    assert ailine.short_error_summary("何かの実行時エラー") == "何かの実行時エラー"

def test_short_error_summary_empty():
    assert ailine.short_error_summary("") == "(詳細不明)"
    assert ailine.short_error_summary(None) == "(詳細不明)"

def test_cmd_run_shows_short_error_and_records_full_detail_in_history(tmp_path, monkeypatch, capsys):
    # ★ M2a: 端末には最終行だけ、履歴 jsonl には全文（トレースバックをそのまま出さない）。
    # ★ W3: 見出し検出には「見出し行(複数の非空文字列)+データ行(型の混在)」が要るため、
    #   単一行 [["a", 1]] でなく見出し+データの2行にする（意図は runtime_error 経路の確認）。
    book = _book(tmp_path, [["商品", "金額"], ["a", 1]])
    monkeypatch.setattr(ailine, "translate_task",
                        lambda model, task, book_meta, temperature=0.1: {"op": "FREEFORM", "args": {}})
    monkeypatch.setattr(ailine, "ollama_generate",
                        lambda model, msgs, temperature=0.2:
                        "Sub Run(oDoc As Object)\nEnd Sub")
    monkeypatch.setattr(ailine, "normalize_book", lambda book, workdir, timeout=None: book)
    traceback_text = ("Traceback (most recent call last):\n"
                       '  File "basrun.py", line 10, in <module>\n'
                       "    raise ValueError('bad state')\n"
                       "ValueError: bad state")
    monkeypatch.setattr(ailine, "basrun_apply",
                        lambda out_book, code, workdir, helper_files=(), timeout=None:
                        (False, traceback_text, traceback_text))
    recorded = {}
    monkeypatch.setattr(ailine, "append_history", lambda entry, path=None: recorded.update(entry))

    ns = argparse.Namespace(
        book=str(book), task="テスト", model="qwen2.5-coder:7b",
        refs=None, helpers=None, repair=0, temperature=0.2,
        dry=False, inplace=False, json=False, timeout=180.0, ask=False)

    rc = ailine.cmd_run(ns)
    captured = capsys.readouterr()
    assert rc == 1
    assert "Traceback (most recent call last):" not in captured.out
    assert "ValueError: bad state" in captured.out
    assert recorded["failure_kind"] == "runtime_error"
    assert recorded["error_detail"] == traceback_text


# --- ★ W8a 項目4/5: 単段 FREEFORM の正直な⚠枠・率スキャン・語彙翻訳 ----------------

def test_cmd_run_freeform_success_shows_honest_warning_not_checkmark(tmp_path, monkeypatch, capsys):
    # ★ W8a 項目4: 単段 FREEFORM/OUT_OF_VOCAB は成功しても『✓』でなく『⚠ AI が直接
    #   作成した処理です（機械保証なし）』の正直な枠で表示する
    #   （実測: 8%仮定やラベル貼りが「✓できました」で素通りしていた）。
    book = _book(tmp_path, [["商品", "金額"], ["a", 100]])
    monkeypatch.setattr(ailine, "HISTORY_FILE", tmp_path / "history.jsonl")
    monkeypatch.setattr(ailine, "VOCAB_FILE", tmp_path / "vocab.json")
    monkeypatch.setattr(ailine, "translate_task",
                        lambda model, task, book_meta, temperature=0.1: {"op": "FREEFORM", "args": {}})
    monkeypatch.setattr(ailine, "ollama_generate",
                        lambda model, msgs, temperature=0.2: "Sub Run(oDoc As Object)\nEnd Sub")
    monkeypatch.setattr(ailine, "normalize_book", lambda book, workdir, timeout=None: book)

    def fake_apply(out_book, code, workdir, helper_files=(), timeout=None):
        wb2 = openpyxl.load_workbook(out_book)
        wb2.active.cell(row=1, column=3, value="税込合計")
        wb2.save(out_book)
        return True, None, "ok"
    monkeypatch.setattr(ailine, "basrun_apply", fake_apply)

    ns = argparse.Namespace(
        book=str(book), task="税込み合計を出して", model="qwen2.5-coder:7b",
        refs=None, helpers=None, repair=0, temperature=0.2,
        dry=False, inplace=False, json=False, timeout=180.0, ask=False)
    rc = ailine.cmd_run(ns)
    captured = capsys.readouterr()
    assert rc == 0
    assert "⚠ AI が直接作成した処理です（機械保証なし）— 確認してください" in captured.out
    assert "✓ 適用され" not in captured.out

def test_cmd_run_freeform_banner_uses_ai_direct_wording_not_jargon(tmp_path, monkeypatch, capsys):
    # ★ W8a 項目5: 「自由生成経路」→「AI が直接作成（機械保証なし）」（operator の語彙翻訳）。
    book = _book(tmp_path, [["a", 1]])
    monkeypatch.setattr(ailine, "translate_task",
                        lambda model, task, book_meta, temperature=0.1: {"op": "FREEFORM", "args": {}})
    monkeypatch.setattr(ailine, "ollama_generate",
                        lambda model, msgs, temperature=0.2: "Sub Run(oDoc As Object)\nEnd Sub")
    ns = argparse.Namespace(
        book=str(book), task="何かして", model="qwen2.5-coder:7b",
        refs=None, helpers=None, repair=0, temperature=0.2,
        dry=True, inplace=False, json=False, timeout=180.0, ask=False)
    ailine.cmd_run(ns)
    captured = capsys.readouterr()
    assert "自由生成経路" not in captured.out
    assert "AI が直接作成" in captured.out

def test_cmd_run_freeform_rate_literal_scan_fires_when_task_silent_on_rate(tmp_path, monkeypatch, capsys):
    # ★ W8a 項目4: 依頼文にも用語集にも率の出典が無いのに生成コードに率らしい数値が
    #   あれば、検算を促す助言が変更点の後ろに出る。
    book = _book(tmp_path, [["商品", "金額"], ["a", 100]])
    monkeypatch.setattr(ailine, "HISTORY_FILE", tmp_path / "history.jsonl")
    monkeypatch.setattr(ailine, "VOCAB_FILE", tmp_path / "vocab.json")
    monkeypatch.setattr(ailine, "translate_task",
                        lambda model, task, book_meta, temperature=0.1: {"op": "FREEFORM", "args": {}})
    code = ("Sub Run(oDoc As Object)\n"
            "  oDoc.Sheets.getByIndex(0).getCellByPosition(2, 0).setValue(100 * 1.08)\n"
            "End Sub")
    monkeypatch.setattr(ailine, "ollama_generate", lambda model, msgs, temperature=0.2: code)
    monkeypatch.setattr(ailine, "normalize_book", lambda book, workdir, timeout=None: book)

    def fake_apply(out_book, c, workdir, helper_files=(), timeout=None):
        wb2 = openpyxl.load_workbook(out_book)
        wb2.active.cell(row=1, column=3, value=108)
        wb2.save(out_book)
        return True, None, "ok"
    monkeypatch.setattr(ailine, "basrun_apply", fake_apply)

    ns = argparse.Namespace(
        book=str(book), task="税込み合計を出して", model="qwen2.5-coder:7b",
        refs=None, helpers=None, repair=0, temperature=0.2,
        dry=False, inplace=False, json=False, timeout=180.0, ask=False)
    rc = ailine.cmd_run(ns)
    captured = capsys.readouterr()
    assert rc == 0
    assert "★ 率らしい数値 (1.08) が依頼に無いのに使われています — 検算してください" in captured.out

def test_cmd_run_freeform_rate_literal_scan_silent_when_task_states_rate(tmp_path, monkeypatch, capsys):
    # 対照: 依頼文に率(消費税8%)の出典があれば発火しない。
    book = _book(tmp_path, [["商品", "金額"], ["a", 100]])
    monkeypatch.setattr(ailine, "HISTORY_FILE", tmp_path / "history.jsonl")
    monkeypatch.setattr(ailine, "VOCAB_FILE", tmp_path / "vocab.json")
    monkeypatch.setattr(ailine, "translate_task",
                        lambda model, task, book_meta, temperature=0.1: {"op": "FREEFORM", "args": {}})
    code = ("Sub Run(oDoc As Object)\n"
            "  oDoc.Sheets.getByIndex(0).getCellByPosition(2, 0).setValue(100 * 1.08)\n"
            "End Sub")
    monkeypatch.setattr(ailine, "ollama_generate", lambda model, msgs, temperature=0.2: code)
    monkeypatch.setattr(ailine, "normalize_book", lambda book, workdir, timeout=None: book)

    def fake_apply(out_book, c, workdir, helper_files=(), timeout=None):
        wb2 = openpyxl.load_workbook(out_book)
        wb2.active.cell(row=1, column=3, value=108)
        wb2.save(out_book)
        return True, None, "ok"
    monkeypatch.setattr(ailine, "basrun_apply", fake_apply)

    ns = argparse.Namespace(
        book=str(book), task="消費税8%込みの合計を出して", model="qwen2.5-coder:7b",
        refs=None, helpers=None, repair=0, temperature=0.2,
        dry=False, inplace=False, json=False, timeout=180.0, ask=False)
    rc = ailine.cmd_run(ns)
    captured = capsys.readouterr()
    assert rc == 0
    assert "率らしい数値" not in captured.out

def test_cmd_run_dsl_dry_banner_says_rule_conversion_not_deterministic(tmp_path, monkeypatch, capsys):
    # ★ W8a 項目5: 「決定論」はユーザー向け文字列から排除（内部名・コメントは不変）。
    book = _book(tmp_path, [["商品", "金額"], ["a", 100], ["b", 200]])
    monkeypatch.setattr(ailine, "HISTORY_FILE", tmp_path / "history.jsonl")
    monkeypatch.setattr(ailine, "translate_task",
                        lambda model, task, book_meta, temperature=0.1:
                        {"op": "SORT", "args": {"col": "金額", "order": "desc"}})
    ns = argparse.Namespace(
        book=str(book), task="金額で降順に並べ替えて", model="qwen2.5-coder:7b",
        refs=None, helpers=None, repair=0, temperature=0.2,
        dry=True, inplace=False, json=False, timeout=180.0, ask=False, values=False)
    rc = ailine.cmd_run(ns)
    captured = capsys.readouterr()
    assert rc == 0
    assert "決定論" not in captured.out
    assert "ルール変換" in captured.out


# --- ★ M2a: --inplace バックアップ + restore --------------------------------

def test_utc_ts_format():
    # ★ W8b: マイクロ秒まで含む（秒精度だけの同名衝突を避けるための拡張・後方互換は
    #   _BACKUP_TS_RE/_ts_sort_key が旧6桁形式も受け付けることで担保）。
    import re as _re
    assert _re.match(r"^\d{8}T\d{12}Z$", ailine._utc_ts())

def test_ts_sort_key_orders_old_and_new_format_correctly():
    # 旧(秒精度)・新(マイクロ秒精度)が混在しても実時刻順に並ぶこと。
    old = "20260101T120000Z"
    new_earlier = "20260101T115959500000Z"
    new_later = "20260101T120000500000Z"
    keys = sorted([old, new_earlier, new_later], key=ailine._ts_sort_key)
    assert keys == [new_earlier, old, new_later]

def test_ts_sort_key_broken_format_sorts_as_oldest():
    assert ailine._ts_sort_key("not-a-timestamp") == (ailine.datetime.min, 0)

def test_ts_sort_key_seq_suffix_breaks_ties_within_same_timestamp():
    # ★ W8b: make_backup の "-N" 衝突回避連番は、同一 ts 内でも新しい順に並ぶこと。
    base = "20260101T120000000000Z"
    keys = sorted([f"{base}-2", base, f"{base}-3"], key=ailine._ts_sort_key)
    assert keys == [base, f"{base}-2", f"{base}-3"]

def test_backup_path_for_uses_stem_ts_suffix(tmp_path, monkeypatch):
    # ★ W8b 項目3: バックアップは book の親フォルダごとの名前空間ディレクトリの下に置く
    #   （同名ファイルが別フォルダにあっても取り違えない・undo 混線の根治）。
    monkeypatch.setattr(ailine, "BACKUP_DIR", tmp_path / "backups")
    book = Path("demo/sample.xlsx")
    ns = ailine._backup_namespace(book)
    p = ailine.backup_path_for(book, ts="20260814T120000Z")
    assert p == tmp_path / "backups" / ns / "sample.20260814T120000Z.xlsx"

def test_make_backup_copies_content(tmp_path, monkeypatch):
    monkeypatch.setattr(ailine, "BACKUP_DIR", tmp_path / "backups")
    book = tmp_path / "book.xlsx"
    book.write_bytes(b"HELLO")
    dst = ailine.make_backup(book)
    assert dst.exists()
    assert dst.read_bytes() == b"HELLO"
    # ★ W8b 項目3: 新規バックアップは名前空間ディレクトリの下（フラット直下ではない）。
    assert dst.parent == tmp_path / "backups" / ailine._backup_namespace(book)

def test_list_backups_sorted_newest_first_and_filters_by_stem_suffix(tmp_path, monkeypatch):
    backups = tmp_path / "backups"
    backups.mkdir(parents=True)
    monkeypatch.setattr(ailine, "BACKUP_DIR", backups)
    (backups / "book.20200101T000000Z.xlsx").write_bytes(b"old")
    (backups / "book.20260101T000000Z.xlsx").write_bytes(b"new")
    (backups / "other.20260101T000000Z.xlsx").write_bytes(b"unrelated")   # 別 book
    (backups / "junk.txt").write_bytes(b"not a backup")                  # 形が違う
    result = ailine.list_backups(tmp_path / "book.xlsx")
    assert [p.name for p in result] == ["book.20260101T000000Z.xlsx", "book.20200101T000000Z.xlsx"]

def test_list_backups_empty_when_dir_missing(tmp_path, monkeypatch):
    monkeypatch.setattr(ailine, "BACKUP_DIR", tmp_path / "nope")
    assert ailine.list_backups(tmp_path / "book.xlsx") == []

def test_restore_backup_restores_and_stays_reversible(tmp_path, monkeypatch):
    backups = tmp_path / "backups"
    backups.mkdir(parents=True)
    monkeypatch.setattr(ailine, "BACKUP_DIR", backups)
    book = tmp_path / "book.xlsx"
    (backups / "book.20200101T000000Z.xlsx").write_bytes(b"BACKED_UP_CONTENT")
    book.write_bytes(b"CURRENT")   # --inplace で上書きされた後の現状を模す

    used = ailine.restore_backup(book)

    assert used.name == "book.20200101T000000Z.xlsx"
    assert book.read_bytes() == b"BACKED_UP_CONTENT"
    remaining = ailine.list_backups(book)
    assert len(remaining) == 2   # 復元前の CURRENT も退避されている＝復元自体も可逆
    contents = {p.read_bytes() for p in remaining}
    assert b"CURRENT" in contents
    assert b"BACKED_UP_CONTENT" in contents

def test_restore_backup_raises_when_none_exist(tmp_path, monkeypatch):
    monkeypatch.setattr(ailine, "BACKUP_DIR", tmp_path / "backups")
    with pytest.raises(FileNotFoundError):
        ailine.restore_backup(tmp_path / "book.xlsx")


# --- ★ W8b 項目3: バックアップの名前空間化（同名別フォルダの undo 混線の根治） ------

def test_backup_namespace_differs_by_parent_folder(tmp_path):
    a = tmp_path / "A" / "見積.xlsx"
    b = tmp_path / "B" / "見積.xlsx"
    assert ailine._backup_namespace(a) != ailine._backup_namespace(b)

def test_backup_namespace_stable_for_same_folder(tmp_path):
    p1 = tmp_path / "見積.xlsx"
    p2 = tmp_path / "見積.xlsx"
    assert ailine._backup_namespace(p1) == ailine._backup_namespace(p2)

def test_backup_and_restore_two_same_named_files_in_different_folders_do_not_cross(tmp_path, monkeypatch):
    # ★ 回帰テスト（architect 指摘）: A\見積.xlsx と B\見積.xlsx を交互に backup→restore
    #   しても混線しない（名前空間分離前は同名+フラット領域のため取り違えが起きえた）。
    monkeypatch.setattr(ailine, "BACKUP_DIR", tmp_path / "backups")
    dir_a = tmp_path / "A"; dir_a.mkdir()
    dir_b = tmp_path / "B"; dir_b.mkdir()
    book_a = dir_a / "見積.xlsx"
    book_b = dir_b / "見積.xlsx"
    book_a.write_bytes(b"A-ORIGINAL")
    book_b.write_bytes(b"B-ORIGINAL")

    ailine.make_backup(book_a)
    ailine.make_backup(book_b)
    book_a.write_bytes(b"A-EDITED")
    book_b.write_bytes(b"B-EDITED")

    ailine.restore_backup(book_a)
    ailine.restore_backup(book_b)

    assert book_a.read_bytes() == b"A-ORIGINAL"
    assert book_b.read_bytes() == b"B-ORIGINAL"
    # A のバックアップ一覧に B の内容(B-ORIGINAL 等)が紛れ込んでいないこと
    a_contents = {p.read_bytes() for p in ailine.list_backups(book_a)}
    assert b"B-ORIGINAL" not in a_contents and b"B-EDITED" not in a_contents

def test_list_backups_reads_legacy_flat_area_read_only(tmp_path, monkeypatch):
    # ★ 旧フラット領域（名前空間分離前の名残）は読み取りのみ互換。
    backups = tmp_path / "backups"
    backups.mkdir(parents=True)
    monkeypatch.setattr(ailine, "BACKUP_DIR", backups)
    (backups / "legacy.20200101T000000Z.xlsx").write_bytes(b"old-flat")
    book = tmp_path / "legacy.xlsx"
    found = ailine.list_backups(book)
    assert [p.name for p in found] == ["legacy.20200101T000000Z.xlsx"]

def test_make_backup_never_writes_to_legacy_flat_area(tmp_path, monkeypatch):
    backups = tmp_path / "backups"
    monkeypatch.setattr(ailine, "BACKUP_DIR", backups)
    book = tmp_path / "book.xlsx"
    book.write_bytes(b"x")
    ailine.make_backup(book)
    flat_files = [p for p in backups.iterdir() if p.is_file()] if backups.is_dir() else []
    assert flat_files == []   # 新規バックアップは名前空間ディレクトリの下だけ


# --- ★ W8b 項目5: ailine undo（restore の昇格） --------------------------------

def test_cmd_undo_restores_and_shows_remaining_count(tmp_path, monkeypatch, capsys):
    backups = tmp_path / "backups"
    monkeypatch.setattr(ailine, "BACKUP_DIR", backups)
    book = tmp_path / "book.xlsx"
    book.write_bytes(b"CURRENT")
    ailine.make_backup(book)   # 1世代目
    book.write_bytes(b"EDITED")

    rc = ailine.cmd_undo(argparse.Namespace(book=str(book), list=False))
    captured = capsys.readouterr()
    assert rc == 0
    assert book.read_bytes() == b"CURRENT"
    assert "あと" in captured.out and "回戻せます" in captured.out

def test_cmd_undo_list_shows_backups_without_restoring(tmp_path, monkeypatch, capsys):
    backups = tmp_path / "backups"
    monkeypatch.setattr(ailine, "BACKUP_DIR", backups)
    book = tmp_path / "book.xlsx"
    book.write_bytes(b"CURRENT")
    ailine.make_backup(book)
    rc = ailine.cmd_undo(argparse.Namespace(book=str(book), list=True))
    captured = capsys.readouterr()
    assert rc == 0
    assert book.read_bytes() == b"CURRENT"   # --list は復元しない
    assert "世代" in captured.out

def test_cmd_undo_fails_honestly_when_no_backup(tmp_path, monkeypatch, capsys):
    monkeypatch.setattr(ailine, "BACKUP_DIR", tmp_path / "backups")
    book = tmp_path / "book.xlsx"
    book.write_bytes(b"x")
    rc = ailine.cmd_undo(argparse.Namespace(book=str(book), list=False))
    captured = capsys.readouterr()
    assert rc == 1
    assert "×" in captured.out


# --- ★ M2c: バックアップのプルーニング -----------------------------------------

def test_prune_backups_keeps_only_newest_n(tmp_path, monkeypatch):
    backups = tmp_path / "backups"
    backups.mkdir(parents=True)
    monkeypatch.setattr(ailine, "BACKUP_DIR", backups)
    book = tmp_path / "book.xlsx"
    for i in range(5):
        (backups / f"book.2026010{i+1}T000000Z.xlsx").write_bytes(b"x")
    deleted = ailine.prune_backups(book, keep=3)
    remaining = ailine.list_backups(book)
    assert len(remaining) == 3
    assert len(deleted) == 2
    assert [p.name for p in remaining] == [
        "book.20260105T000000Z.xlsx", "book.20260104T000000Z.xlsx", "book.20260103T000000Z.xlsx"]

def test_prune_backups_negative_keep_means_unlimited(tmp_path, monkeypatch):
    backups = tmp_path / "backups"
    backups.mkdir(parents=True)
    monkeypatch.setattr(ailine, "BACKUP_DIR", backups)
    book = tmp_path / "book.xlsx"
    for i in range(3):
        (backups / f"book.2026010{i+1}T000000Z.xlsx").write_bytes(b"x")
    deleted = ailine.prune_backups(book, keep=-1)
    assert deleted == []
    assert len(ailine.list_backups(book)) == 3

def test_make_backup_prunes_beyond_keep(tmp_path, monkeypatch):
    monkeypatch.setattr(ailine, "BACKUP_DIR", tmp_path / "backups")
    book = tmp_path / "book.xlsx"
    book.write_bytes(b"content")
    for i in range(3):
        monkeypatch.setattr(ailine, "_utc_ts", lambda i=i: f"2026010{i+1}T000000Z")
        ailine.make_backup(book, keep=2)
    remaining = ailine.list_backups(book)
    assert len(remaining) == 2
    assert remaining[0].name == "book.20260103T000000Z.xlsx"   # 最新が残る

def test_make_backup_default_keep_is_ten(tmp_path, monkeypatch):
    monkeypatch.setattr(ailine, "BACKUP_DIR", tmp_path / "backups")
    book = tmp_path / "book.xlsx"
    book.write_bytes(b"content")
    for i in range(12):
        monkeypatch.setattr(ailine, "_utc_ts", lambda i=i: f"202601{i+1:02d}T000000Z")
        ailine.make_backup(book)   # keep 省略 = 既定 DEFAULT_KEEP_BACKUPS(=10)
    assert len(ailine.list_backups(book)) == 10

def test_cmd_restore_list_shows_generation_count(tmp_path, monkeypatch, capsys):
    backups = tmp_path / "backups"
    backups.mkdir(parents=True)
    monkeypatch.setattr(ailine, "BACKUP_DIR", backups)
    (backups / "book.20200101T000000Z.xlsx").write_bytes(b"a")
    (backups / "book.20200102T000000Z.xlsx").write_bytes(b"b")
    ns = argparse.Namespace(book=str(tmp_path / "book.xlsx"), list=True)
    rc = ailine.cmd_restore(ns)
    captured = capsys.readouterr()
    assert rc == 0
    assert "2 世代" in captured.out


# ===========================================================================
# ★ W8b 項目2: Excel ロックの事前検出
# ===========================================================================

def test_check_excel_lock_none_when_free(tmp_path):
    book = tmp_path / "book.xlsx"
    book.write_bytes(b"x")
    assert ailine.check_excel_lock(book) is None

def test_check_excel_lock_detects_tilde_dollar_lock_file(tmp_path):
    book = tmp_path / "book.xlsx"
    book.write_bytes(b"x")
    (tmp_path / "~$book.xlsx").write_bytes(b"lock")
    reason = ailine.check_excel_lock(book)
    assert reason is not None
    assert "ロックファイル" in reason

def test_check_excel_lock_detects_exclusively_opened_file(tmp_path):
    # ★ 実際の Excel の排他 open を模す。Python の open() 二重呼び出しは Windows
    #   既定の共有モードでは失敗しないため、CreateFileW を dwShareMode=0 で直接呼ぶ
    #   （Excel 相当の排他アクセス）。
    import ctypes
    from ctypes import wintypes
    book = tmp_path / "book.xlsx"
    book.write_bytes(b"x")
    GENERIC_READ = 0x80000000
    GENERIC_WRITE = 0x40000000
    OPEN_EXISTING = 3
    CreateFileW = ctypes.windll.kernel32.CreateFileW
    CreateFileW.restype = wintypes.HANDLE
    CreateFileW.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD, wintypes.LPVOID,
                             wintypes.DWORD, wintypes.DWORD, wintypes.HANDLE]
    handle = CreateFileW(str(book), GENERIC_READ | GENERIC_WRITE, 0, None, OPEN_EXISTING, 0, None)
    assert handle not in (0, -1)
    try:
        reason = ailine.check_excel_lock(book)
        assert reason is not None
        assert "ロック" in reason
    finally:
        ctypes.windll.kernel32.CloseHandle(handle)

def test_cmd_run_exits_5_when_excel_lock_detected(tmp_path, monkeypatch, capsys):
    book = _book(tmp_path, [["商品", "金額"], ["a", 1]])
    (tmp_path / f"~${book.name}").write_bytes(b"lock")
    called = {"n": 0}
    def boom(*a, **k):
        called["n"] += 1
        return {"op": "FREEFORM", "args": {}}
    monkeypatch.setattr(ailine, "translate_task", boom)
    ns = argparse.Namespace(
        book=str(book), task="何かして", model="qwen2.5-coder:7b",
        refs=None, helpers=None, repair=0, temperature=0.2,
        dry=False, inplace=False, json=False, timeout=180.0, ask=False)
    rc = ailine.cmd_run(ns)
    captured = capsys.readouterr()
    assert rc == 5
    assert "Excel で開かれています" in captured.out
    assert called["n"] == 0   # LO 起動・翻訳より前に止まっている


# ===========================================================================
# ★ W8b 項目1: 往復忠実度ゲート
# ===========================================================================

def _make_fake_zip(tmp_path, name, members):
    import zipfile as _zf
    p = tmp_path / name
    with _zf.ZipFile(p, "w") as z:
        for m in members:
            z.writestr(m, b"x")
    return p

def test_check_zip_fidelity_loss_detects_categories(tmp_path):
    original = _make_fake_zip(tmp_path, "orig.xlsx", [
        "xl/workbook.xml", "xl/worksheets/sheet1.xml",
        "xl/drawings/drawing1.xml", "xl/media/image1.png", "xl/media/image2.png",
        "xl/vbaProject.bin", "xl/pivotCache/pivotCacheDefinition1.xml",
        "xl/pivotTables/pivotTable1.xml", "xl/worksheets/_rels/sheet1.xml.rels",
    ])
    normalized = _make_fake_zip(tmp_path, "norm.xlsx", ["xl/workbook.xml", "xl/worksheets/sheet1.xml"])
    result = dict(ailine.check_zip_fidelity_loss(original, normalized))
    assert result["図形/描画"] == 1
    assert result["画像"] == 2
    assert result["VBA マクロ"] == 1
    assert result["ピボットテーブル"] == 2   # pivotCache + pivotTables
    assert result["リンク情報(_rels)"] == 1

def test_check_zip_fidelity_loss_empty_when_nothing_lost(tmp_path):
    members = ["xl/workbook.xml", "xl/worksheets/sheet1.xml", "xl/media/image1.png"]
    original = _make_fake_zip(tmp_path, "orig.xlsx", members)
    normalized = _make_fake_zip(tmp_path, "norm.xlsx", members)
    assert ailine.check_zip_fidelity_loss(original, normalized) == []

def test_check_zip_fidelity_loss_ignores_unrelated_removed_members(tmp_path):
    original = _make_fake_zip(tmp_path, "orig.xlsx", ["xl/workbook.xml", "docProps/core.xml"])
    normalized = _make_fake_zip(tmp_path, "norm.xlsx", ["xl/workbook.xml"])
    assert ailine.check_zip_fidelity_loss(original, normalized) == []

def test_check_zip_fidelity_loss_empty_when_original_not_a_zip(tmp_path):
    p = tmp_path / "not_a_zip.xlsx"
    p.write_bytes(b"not a zip file")
    normalized = _make_fake_zip(tmp_path, "norm.xlsx", ["xl/workbook.xml"])
    assert ailine.check_zip_fidelity_loss(p, normalized) == []

def _cf_dv_book(tmp_path, name, add_cf=True, add_dv=True):
    # ★ 見出し(1行目)+データ(2行以降・型混在)を持たせる。単一セルだけの本だと
    # detect_header_row が確信を持てず、cmd_run 統合テストが CLARIFY に落ちてしまう。
    from openpyxl.formatting.rule import CellIsRule
    from openpyxl.worksheet.datavalidation import DataValidation
    p = tmp_path / name
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["商品", "金額"])
    ws.append(["りんご", 100])
    ws.append(["バナナ", 200])
    if add_cf:
        ws.conditional_formatting.add("B2:B3", CellIsRule(operator="greaterThan", formula=["150"]))
    if add_dv:
        dv = DataValidation(type="list", formula1='"a,b,c"')
        ws.add_data_validation(dv)
        dv.add("C2")
    wb.save(p)
    return p

def test_check_openpyxl_fidelity_loss_detects_lost_cf_and_dv(tmp_path):
    original = _cf_dv_book(tmp_path, "orig.xlsx", add_cf=True, add_dv=True)
    normalized = _cf_dv_book(tmp_path, "norm.xlsx", add_cf=False, add_dv=False)
    result = ailine.check_openpyxl_fidelity_loss(original, normalized)
    labels = {r[0] for r in result}
    assert labels == {"条件付き書式", "入力規則"}

def test_check_openpyxl_fidelity_loss_empty_when_unchanged(tmp_path):
    original = _cf_dv_book(tmp_path, "orig.xlsx")
    normalized = _cf_dv_book(tmp_path, "norm.xlsx")
    assert ailine.check_openpyxl_fidelity_loss(original, normalized) == []

def test_check_openpyxl_fidelity_loss_empty_when_unreadable(tmp_path):
    p = tmp_path / "bad.xlsx"
    p.write_bytes(b"not xlsx")
    normalized = _cf_dv_book(tmp_path, "norm.xlsx")
    assert ailine.check_openpyxl_fidelity_loss(p, normalized) == []

def test_check_round_trip_fidelity_combines_both_checks(tmp_path):
    original = _cf_dv_book(tmp_path, "orig.xlsx", add_cf=True, add_dv=False)
    normalized = _cf_dv_book(tmp_path, "norm.xlsx", add_cf=False, add_dv=False)
    fidelity = ailine.check_round_trip_fidelity(original, normalized)
    assert fidelity["lost"] is True
    assert any(it["label"] == "条件付き書式" and it["count"] == 1 for it in fidelity["items"])

def test_check_round_trip_fidelity_silent_when_nothing_lost(tmp_path):
    original = _cf_dv_book(tmp_path, "orig.xlsx")
    normalized = _cf_dv_book(tmp_path, "norm.xlsx")
    assert ailine.check_round_trip_fidelity(original, normalized) == {"lost": False, "items": []}

def test_format_fidelity_warning_lists_categories_and_counts():
    fidelity = {"lost": True, "items": [{"label": "条件付き書式", "count": 3}]}
    msg = ailine.format_fidelity_warning(fidelity)
    assert msg == "⚠ このファイルには、処理すると失われる飾りがあります（条件付き書式 3 件）"


def _fidelity_gate_ns(book, **overrides):
    base = dict(
        book=str(book), task="何かして", model="qwen2.5-coder:7b",
        refs=None, helpers=None, repair=0, temperature=0.2,
        dry=False, inplace=True, json=False, timeout=180.0, ask=False,
        accept_loss=False, copy=False)
    base.update(overrides)
    return argparse.Namespace(**base)

def _patch_lossy_normalize(monkeypatch):
    """normalize_book を『CF が消える正規化』に差し替える（実際の LO を使わない）。
       値・見出しは原本と同じ形に保つ（header 検出が CLARIFY に落ちないように）。"""
    def fake_normalize(book, workdir, timeout=None):
        norm = workdir / ("normalized" + book.suffix)
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["商品", "金額"])
        ws.append(["りんご", 100])
        ws.append(["バナナ", 200])   # CF/DV なし
        wb.save(norm)
        return norm
    monkeypatch.setattr(ailine, "normalize_book", fake_normalize)

def test_cmd_run_inplace_fidelity_gate_blocks_without_flag(tmp_path, monkeypatch, capsys):
    book = _cf_dv_book(tmp_path, "book.xlsx", add_cf=True, add_dv=False)
    monkeypatch.setattr(ailine, "HISTORY_FILE", tmp_path / "history.jsonl")
    _patch_lossy_normalize(monkeypatch)
    called = {"n": 0}
    def boom(*a, **k):
        called["n"] += 1
        return {"op": "FREEFORM", "args": {}}
    monkeypatch.setattr(ailine, "translate_task", boom)
    rc = ailine.cmd_run(_fidelity_gate_ns(book))
    captured = capsys.readouterr()
    assert rc == 4
    assert "失われる飾りがあります" in captured.out
    assert "--accept-loss" in captured.out
    assert "--copy" in captured.out
    assert called["n"] == 0   # ゲートで止まる＝翻訳より前

def test_cmd_run_inplace_fidelity_gate_accept_loss_continues(tmp_path, monkeypatch, capsys):
    book = _cf_dv_book(tmp_path, "book.xlsx", add_cf=True, add_dv=False)
    monkeypatch.setattr(ailine, "HISTORY_FILE", tmp_path / "history.jsonl")
    monkeypatch.setattr(ailine, "VOCAB_FILE", tmp_path / "vocab.json")
    monkeypatch.setattr(ailine, "BACKUP_DIR", tmp_path / "backups")
    _patch_lossy_normalize(monkeypatch)
    monkeypatch.setattr(ailine, "translate_task",
                        lambda model, task, book_meta, temperature=0.1: {"op": "FREEFORM", "args": {}})
    monkeypatch.setattr(ailine, "ollama_generate",
                        lambda model, msgs, temperature=0.2: "Sub Run(oDoc As Object)\nEnd Sub")
    def fake_apply(out_book, code, workdir, helper_files=(), timeout=None):
        wb2 = openpyxl.load_workbook(out_book)
        wb2.active.cell(row=1, column=2, value="changed")
        wb2.save(out_book)
        return True, None, "ok"
    monkeypatch.setattr(ailine, "basrun_apply", fake_apply)
    rc = ailine.cmd_run(_fidelity_gate_ns(book, accept_loss=True))
    captured = capsys.readouterr()
    assert "続行します" in captured.out
    assert f"適用先: {book.name}" in captured.out   # 原本へ実際に適用された

def test_cmd_run_inplace_fidelity_gate_copy_flag_downgrades_to_out(tmp_path, monkeypatch, capsys):
    book = _cf_dv_book(tmp_path, "book.xlsx", add_cf=True, add_dv=False)
    original_bytes = book.read_bytes()
    monkeypatch.setattr(ailine, "HISTORY_FILE", tmp_path / "history.jsonl")
    monkeypatch.setattr(ailine, "VOCAB_FILE", tmp_path / "vocab.json")
    _patch_lossy_normalize(monkeypatch)
    monkeypatch.setattr(ailine, "translate_task",
                        lambda model, task, book_meta, temperature=0.1: {"op": "FREEFORM", "args": {}})
    monkeypatch.setattr(ailine, "ollama_generate",
                        lambda model, msgs, temperature=0.2: "Sub Run(oDoc As Object)\nEnd Sub")
    def fake_apply(out_book, code, workdir, helper_files=(), timeout=None):
        wb2 = openpyxl.load_workbook(out_book)
        wb2.active.cell(row=1, column=2, value="changed")
        wb2.save(out_book)
        return True, None, "ok"
    monkeypatch.setattr(ailine, "basrun_apply", fake_apply)
    rc = ailine.cmd_run(_fidelity_gate_ns(book, copy=True))
    captured = capsys.readouterr()
    assert "--copy 指定のため .out へ切り替え" in captured.out
    assert book.read_bytes() == original_bytes   # 原本は無変更
    out_book = book.with_name(book.stem + ".out" + book.suffix)
    assert out_book.exists()

def test_cmd_run_inplace_fidelity_gate_silent_when_no_loss(tmp_path, monkeypatch, capsys):
    # 喪失ゼロなら --inplace でも無言で通る（体験は不変）。ゲート自体は --dry では
    # 走らない（LibreOffice に一切触れない設計不変条件）ので、ここは非 dry で確認する。
    book = _book(tmp_path, [["商品", "金額"], ["a", 1]])
    monkeypatch.setattr(ailine, "HISTORY_FILE", tmp_path / "history.jsonl")
    monkeypatch.setattr(ailine, "BACKUP_DIR", tmp_path / "backups")
    monkeypatch.setattr(ailine, "normalize_book", lambda b, workdir, timeout=None: b)   # 変化なし＝喪失ゼロ
    monkeypatch.setattr(ailine, "translate_task",
                        lambda model, task, book_meta, temperature=0.1: {"op": "FREEFORM", "args": {}})
    monkeypatch.setattr(ailine, "ollama_generate",
                        lambda model, msgs, temperature=0.2: "Sub Run(oDoc As Object)\nEnd Sub")
    def fake_apply(out_book, code, workdir, helper_files=(), timeout=None):
        wb2 = openpyxl.load_workbook(out_book)
        wb2.active.cell(row=1, column=3, value="new")
        wb2.save(out_book)
        return True, None, "ok"
    monkeypatch.setattr(ailine, "basrun_apply", fake_apply)
    ns = _fidelity_gate_ns(book, dry=False)
    ailine.cmd_run(ns)
    captured = capsys.readouterr()
    assert "失われる飾り" not in captured.out

def test_cmd_run_inplace_fidelity_records_history_field(tmp_path, monkeypatch, capsys):
    book = _cf_dv_book(tmp_path, "book.xlsx", add_cf=True, add_dv=False)
    monkeypatch.setattr(ailine, "VOCAB_FILE", tmp_path / "vocab.json")
    monkeypatch.setattr(ailine, "BACKUP_DIR", tmp_path / "backups")
    _patch_lossy_normalize(monkeypatch)
    monkeypatch.setattr(ailine, "translate_task",
                        lambda model, task, book_meta, temperature=0.1: {"op": "FREEFORM", "args": {}})
    monkeypatch.setattr(ailine, "ollama_generate",
                        lambda model, msgs, temperature=0.2: "Sub Run(oDoc As Object)\nEnd Sub")
    def fake_apply(out_book, code, workdir, helper_files=(), timeout=None):
        wb2 = openpyxl.load_workbook(out_book)
        wb2.active.cell(row=1, column=2, value="changed")
        wb2.save(out_book)
        return True, None, "ok"
    monkeypatch.setattr(ailine, "basrun_apply", fake_apply)
    recorded = {}
    monkeypatch.setattr(ailine, "append_history", lambda entry, path=None: recorded.update(entry))
    ailine.cmd_run(_fidelity_gate_ns(book, accept_loss=True))
    assert recorded["fidelity"]["lost"] is True
    assert any(it["label"] == "条件付き書式" for it in recorded["fidelity"]["items"])

def test_cmd_run_inplace_no_fidelity_field_when_gate_not_run(tmp_path, monkeypatch, capsys):
    # --inplace すら要求していない run では、ゲートは走らず fidelity は None のまま。
    book = _book(tmp_path, [["a", 1], ["b", 2]])
    monkeypatch.setattr(ailine, "translate_task",
                        lambda model, task, book_meta, temperature=0.1: {"op": "FREEFORM", "args": {}})
    monkeypatch.setattr(ailine, "ollama_generate",
                        lambda model, msgs, temperature=0.2: "Sub Run(oDoc As Object)\nEnd Sub")
    recorded = {}
    monkeypatch.setattr(ailine, "append_history", lambda entry, path=None: recorded.update(entry))
    ns = _fidelity_gate_ns(book, inplace=False, dry=True)
    ailine.cmd_run(ns)
    assert recorded["fidelity"] is None


# ===========================================================================
# ★ W8b 項目4: アトミック置換
# ===========================================================================

def test_atomic_replace_inplace_success_replaces_book_and_cleans_up(tmp_path, monkeypatch):
    monkeypatch.setattr(ailine, "BACKUP_DIR", tmp_path / "backups")
    book = tmp_path / "book.xlsx"
    book.write_bytes(b"ORIGINAL")
    out_book = tmp_path / "book.out.xlsx"
    out_book.write_bytes(b"NEW-CONTENT")
    workdir = tmp_path / ".ailine_book"
    workdir.mkdir()
    ok, err = ailine.atomic_replace_inplace(book, out_book, workdir)
    assert ok is True
    assert err is None
    assert book.read_bytes() == b"NEW-CONTENT"
    assert not out_book.exists()          # 置換後は .out を残さない（旧 shutil.move と同じ終状態）
    assert not (workdir / f"staged{book.suffix}").exists()   # staging は片付く
    backups = ailine.list_backups(book)
    assert len(backups) == 1
    assert backups[0].read_bytes() == b"ORIGINAL"   # 置換前の原本がバックアップされている

def test_atomic_replace_inplace_backup_failure_aborts_without_touching_book(tmp_path, monkeypatch):
    book = tmp_path / "book.xlsx"
    book.write_bytes(b"ORIGINAL")
    out_book = tmp_path / "book.out.xlsx"
    out_book.write_bytes(b"NEW-CONTENT")
    workdir = tmp_path / ".ailine_book"
    workdir.mkdir()
    monkeypatch.setattr(ailine, "make_backup", lambda *a, **k: (_ for _ in ()).throw(OSError("disk full")))
    ok, err = ailine.atomic_replace_inplace(book, out_book, workdir)
    assert ok is False
    assert "バックアップに失敗" in err
    assert book.read_bytes() == b"ORIGINAL"   # 原本は無変更
    assert out_book.exists()                   # .out はそのまま残る

def test_atomic_replace_inplace_falls_back_to_copy2_when_os_replace_fails(tmp_path, monkeypatch):
    monkeypatch.setattr(ailine, "BACKUP_DIR", tmp_path / "backups")
    book = tmp_path / "book.xlsx"
    book.write_bytes(b"ORIGINAL")
    out_book = tmp_path / "book.out.xlsx"
    out_book.write_bytes(b"NEW-CONTENT")
    workdir = tmp_path / ".ailine_book"
    workdir.mkdir()
    monkeypatch.setattr(ailine.os, "replace", lambda *a, **k: (_ for _ in ()).throw(OSError("locked")))
    ok, err = ailine.atomic_replace_inplace(book, out_book, workdir)
    assert ok is True
    assert err is None
    assert book.read_bytes() == b"NEW-CONTENT"   # copy2 フォールバックで反映されている
    backups = ailine.list_backups(book)
    assert len(backups) == 1   # バックアップは os.replace 失敗の前に既に確保済み

def test_atomic_replace_inplace_both_replace_and_fallback_fail_reports_honestly(tmp_path, monkeypatch):
    import shutil as _shutil
    monkeypatch.setattr(ailine, "BACKUP_DIR", tmp_path / "backups")
    book = tmp_path / "book.xlsx"
    book.write_bytes(b"ORIGINAL")
    out_book = tmp_path / "book.out.xlsx"
    out_book.write_bytes(b"NEW-CONTENT")
    workdir = tmp_path / ".ailine_book"
    workdir.mkdir()
    monkeypatch.setattr(ailine.os, "replace", lambda *a, **k: (_ for _ in ()).throw(OSError("locked")))
    real_copy2 = _shutil.copy2
    calls = {"n": 0}
    def fake_copy2(src, dst):
        calls["n"] += 1
        # 1回目=バックアップ作成の内部コピー・2回目=staging へのコピー は成功させ、
        # 3回目=os.replace 失敗後のフォールバック copy2(out_book, book) だけ失敗させる。
        if calls["n"] <= 2:
            return real_copy2(src, dst)
        raise OSError("also locked")
    monkeypatch.setattr(ailine.shutil, "copy2", fake_copy2)
    ok, err = ailine.atomic_replace_inplace(book, out_book, workdir)
    assert ok is False
    assert "置換に失敗した" in err


# ===========================================================================
# ★ W8b 項目6: グローバル run ロック
# ===========================================================================

def test_acquire_and_release_run_lock_roundtrip(tmp_path):
    lock = tmp_path / "run.lock"
    acquired, msg = ailine.acquire_run_lock(lock)
    assert acquired is True
    assert msg is None
    assert lock.exists()
    ailine.release_run_lock(lock)
    assert not lock.exists()

def test_acquire_run_lock_fails_when_held_by_live_other_process(tmp_path, monkeypatch):
    lock = tmp_path / "run.lock"
    other_pid = 999999   # 実在しないふりをする pid（_pid_alive を差し替えて『生きている』にする）
    lock.write_text(json.dumps({"pid": other_pid, "ts": ailine.datetime.now(ailine.timezone.utc)
                                 .isoformat(timespec="seconds")}), encoding="utf-8")
    monkeypatch.setattr(ailine, "_pid_alive", lambda pid: pid == other_pid)
    acquired, msg = ailine.acquire_run_lock(lock)
    assert acquired is False
    assert "別の ailine が実行中です" in msg
    assert lock.exists()   # 奪取していない

def test_acquire_run_lock_reclaims_when_pid_is_dead(tmp_path, monkeypatch):
    lock = tmp_path / "run.lock"
    lock.write_text(json.dumps({"pid": 999999, "ts": ailine.datetime.now(ailine.timezone.utc)
                                 .isoformat(timespec="seconds")}), encoding="utf-8")
    monkeypatch.setattr(ailine, "_pid_alive", lambda pid: False)
    acquired, msg = ailine.acquire_run_lock(lock)
    assert acquired is True
    info = json.loads(lock.read_text(encoding="utf-8"))
    assert info["pid"] == os.getpid()

def test_acquire_run_lock_reclaims_when_stale_by_age(tmp_path, monkeypatch):
    import datetime as _dt
    lock = tmp_path / "run.lock"
    old_ts = (ailine.datetime.now(ailine.timezone.utc) - _dt.timedelta(hours=1)).isoformat(timespec="seconds")
    lock.write_text(json.dumps({"pid": 999999, "ts": old_ts}), encoding="utf-8")
    monkeypatch.setattr(ailine, "_pid_alive", lambda pid: True)   # pid 自体は生きている
    acquired, msg = ailine.acquire_run_lock(lock)
    assert acquired is True   # 30分超の age で stale 判定・奪取できる

def test_acquire_run_lock_reclaims_own_leftover_lock_same_pid(tmp_path):
    # ★ 同一プロセス内で前回の呼び出しが解放し損ねた場合も自己修復する
    #   （テスト実行のようにシーケンシャルに同じ pid で cmd_run が繰り返し走る状況の安全網）。
    lock = tmp_path / "run.lock"
    lock.write_text(json.dumps({"pid": os.getpid(), "ts": ailine.datetime.now(ailine.timezone.utc)
                                 .isoformat(timespec="seconds")}), encoding="utf-8")
    acquired, msg = ailine.acquire_run_lock(lock)
    assert acquired is True

def test_acquire_run_lock_broken_file_is_stale(tmp_path):
    lock = tmp_path / "run.lock"
    lock.write_text("not json", encoding="utf-8")
    acquired, msg = ailine.acquire_run_lock(lock)
    assert acquired is True

def test_cmd_run_exits_6_when_run_lock_busy(tmp_path, monkeypatch, capsys):
    book = _book(tmp_path, [["a", 1], ["b", 2]])
    lock_path = tmp_path / "run.lock"
    monkeypatch.setattr(ailine, "RUN_LOCK_FILE", lock_path)
    other_pid = 999999
    lock_path.write_text(json.dumps({"pid": other_pid, "ts": ailine.datetime.now(ailine.timezone.utc)
                                     .isoformat(timespec="seconds")}), encoding="utf-8")
    monkeypatch.setattr(ailine, "_pid_alive", lambda pid: pid == other_pid)
    called = {"n": 0}
    monkeypatch.setattr(ailine, "check_excel_lock", lambda b: called.__setitem__("n", called["n"] + 1) or None)
    ns = argparse.Namespace(
        book=str(book), task="何かして", model="qwen2.5-coder:7b",
        refs=None, helpers=None, repair=0, temperature=0.2,
        dry=True, inplace=False, json=False, timeout=180.0, ask=False)
    rc = ailine.cmd_run(ns)
    captured = capsys.readouterr()
    assert rc == 6
    assert "別の ailine が実行中です" in captured.out
    assert called["n"] == 0   # ロックで止まった＝本体は一切呼ばれていない

def test_cmd_run_releases_lock_even_on_early_sys_exit(tmp_path, monkeypatch):
    # book が無い場合は sys.exit() する経路（SystemExit）。それでも lock は解放されること。
    lock_path = tmp_path / "run.lock"
    monkeypatch.setattr(ailine, "RUN_LOCK_FILE", lock_path)
    ns = argparse.Namespace(
        book=str(tmp_path / "nope.xlsx"), task="何かして", model="qwen2.5-coder:7b",
        refs=None, helpers=None, repair=0, temperature=0.2,
        dry=True, inplace=False, json=False, timeout=180.0, ask=False)
    with pytest.raises(SystemExit):
        ailine.cmd_run(ns)
    assert not lock_path.exists()


# ===========================================================================
# ★ W8b 項目4: run 終了時に自分の workdir を掃除する
# ===========================================================================

def test_cmd_run_cleans_up_workdir_after_success(tmp_path, monkeypatch):
    book = _book(tmp_path, [["a", 1], ["b", 2]])
    monkeypatch.setattr(ailine, "translate_task",
                        lambda model, task, book_meta, temperature=0.1: {"op": "FREEFORM", "args": {}})
    monkeypatch.setattr(ailine, "ollama_generate",
                        lambda model, msgs, temperature=0.2: "Sub Run(oDoc As Object)\nEnd Sub")
    ns = argparse.Namespace(
        book=str(book), task="何かして", model="qwen2.5-coder:7b",
        refs=None, helpers=None, repair=0, temperature=0.2,
        dry=True, inplace=False, json=False, timeout=180.0, ask=False)
    ailine.cmd_run(ns)
    assert not (book.parent / f".ailine_{book.stem}").exists()

def test_cmd_run_cleans_up_workdir_after_clarify_exit(tmp_path, monkeypatch):
    book = _book(tmp_path, [["a", 1], ["b", 2]])
    ambiguous = {"sheets": {"Sheet": {"rows": {
        1: {"nonempty": 2, "str": 2, "bold": 0},
        2: {"nonempty": 2, "str": 1, "bold": 0},
        3: {"nonempty": 2, "str": 2, "bold": 0},
        4: {"nonempty": 2, "str": 1, "bold": 0},
    }}}}
    monkeypatch.setattr(ailine, "normalize_book", lambda book, workdir, timeout=None: book)
    monkeypatch.setattr(ailine, "build_struct_dump", lambda book, workdir: ambiguous)
    ns = argparse.Namespace(
        book=str(book), task="いい感じにして", model="qwen2.5-coder:7b",
        refs=None, helpers=None, repair=0, temperature=0.2,
        dry=False, inplace=False, json=False, timeout=180.0, ask=False)
    rc = ailine.cmd_run(ns)
    assert rc == 3
    assert not (book.parent / f".ailine_{book.stem}").exists()

# --- ★ M2c: 正規化パス失敗時の1回だけ自動リトライ ------------------------------

def test_normalize_book_retries_once_after_stop_and_succeeds(tmp_path, monkeypatch):
    book = tmp_path / "book.xlsx"
    book.write_bytes(b"x")
    calls = {"apply": 0, "stop": 0}
    def fake_apply(normalized, code, workdir, timeout=None):
        calls["apply"] += 1
        if calls["apply"] == 1:
            return False, "RuntimeException: Could not create system bitmap!", "raw"
        return True, None, "raw"
    monkeypatch.setattr(ailine, "basrun_apply", fake_apply)
    monkeypatch.setattr(ailine, "_stop_office", lambda: calls.__setitem__("stop", calls["stop"] + 1))
    result = ailine.normalize_book(book, tmp_path)
    assert calls["apply"] == 2
    assert calls["stop"] == 1
    assert result.exists()

def test_normalize_book_gives_up_after_second_failure(tmp_path, monkeypatch):
    book = tmp_path / "book.xlsx"
    book.write_bytes(b"x")
    monkeypatch.setattr(ailine, "basrun_apply", lambda *a, **k: (False, "boom", "raw"))
    monkeypatch.setattr(ailine, "_stop_office", lambda: None)
    with pytest.raises(SystemExit):
        ailine.normalize_book(book, tmp_path)

def test_normalize_book_succeeds_first_try_without_retry(tmp_path, monkeypatch):
    book = tmp_path / "book.xlsx"
    book.write_bytes(b"x")
    calls = {"apply": 0, "stop": 0}
    def fake_apply(normalized, code, workdir, timeout=None):
        calls["apply"] += 1
        return True, None, "raw"
    monkeypatch.setattr(ailine, "basrun_apply", fake_apply)
    monkeypatch.setattr(ailine, "_stop_office", lambda: calls.__setitem__("stop", calls["stop"] + 1))
    ailine.normalize_book(book, tmp_path)
    assert calls["apply"] == 1
    assert calls["stop"] == 0

def test_cmd_restore_list_shows_backups(tmp_path, monkeypatch, capsys):
    backups = tmp_path / "backups"
    backups.mkdir(parents=True)
    monkeypatch.setattr(ailine, "BACKUP_DIR", backups)
    (backups / "book.20200101T000000Z.xlsx").write_bytes(b"a")
    ns = argparse.Namespace(book=str(tmp_path / "book.xlsx"), list=True)
    rc = ailine.cmd_restore(ns)
    captured = capsys.readouterr()
    assert rc == 0
    assert "book.20200101T000000Z.xlsx" in captured.out

def test_cmd_restore_list_says_none_when_no_backups(tmp_path, monkeypatch, capsys):
    monkeypatch.setattr(ailine, "BACKUP_DIR", tmp_path / "backups")
    ns = argparse.Namespace(book=str(tmp_path / "book.xlsx"), list=True)
    rc = ailine.cmd_restore(ns)
    captured = capsys.readouterr()
    assert rc == 0
    assert "無い" in captured.out

def test_cmd_restore_restores_and_reports_success(tmp_path, monkeypatch, capsys):
    backups = tmp_path / "backups"
    backups.mkdir(parents=True)
    monkeypatch.setattr(ailine, "BACKUP_DIR", backups)
    book = tmp_path / "book.xlsx"
    book.write_bytes(b"current")
    (backups / "book.20200101T000000Z.xlsx").write_bytes(b"restored")
    ns = argparse.Namespace(book=str(book), list=False)
    rc = ailine.cmd_restore(ns)
    captured = capsys.readouterr()
    assert rc == 0
    assert "✓" in captured.out
    assert book.read_bytes() == b"restored"

def test_cmd_restore_fails_when_no_backups(tmp_path, monkeypatch, capsys):
    monkeypatch.setattr(ailine, "BACKUP_DIR", tmp_path / "backups")
    ns = argparse.Namespace(book=str(tmp_path / "book.xlsx"), list=False)
    rc = ailine.cmd_restore(ns)
    captured = capsys.readouterr()
    assert rc == 1
    assert "×" in captured.out


# ---------------------------------------------------------------------------
# ★ M2b: 中間命令言語（DSL）パイプライン
# ---------------------------------------------------------------------------

# --- ① 翻訳（json 退避） -----------------------------------------------------

_SAMPLE_META = {"sheets": ["Sheet"], "headers": {"Sheet": ["商品", "金額", "在庫", "売上", "原価"]}}


def test_translate_task_valid_json_nested_args(monkeypatch):
    # ★ M2c: translate_task は常に {"plan": [...]} を返す。後方互換で単一 op は長さ1の計画。
    monkeypatch.setattr(ailine, "ollama_generate_json",
                        lambda model, msgs, temperature=0.1, num_predict=300:
                        '{"op": "SORT", "args": {"col": "金額", "order": "desc"}}')
    got = ailine.translate_task("qwen2.5-coder:7b", "金額で降順に並べ替えて", _SAMPLE_META)
    assert got == {"plan": [{"op": "SORT", "args": {"col": "金額", "order": "desc"}}]}

def test_translate_task_flat_args_are_rescued(monkeypatch):
    # モデルが args で包まず op と slot をフラットに返した場合も救済する。
    monkeypatch.setattr(ailine, "ollama_generate_json",
                        lambda model, msgs, temperature=0.1, num_predict=300:
                        '{"op": "SORT", "col": "金額", "order": "desc"}')
    got = ailine.translate_task("qwen2.5-coder:7b", "金額で降順に並べ替えて", _SAMPLE_META)
    assert got["plan"][0]["op"] == "SORT"
    assert got["plan"][0]["args"] == {"col": "金額", "order": "desc"}

def test_translate_task_clarify_passthrough(monkeypatch):
    monkeypatch.setattr(ailine, "ollama_generate_json",
                        lambda model, msgs, temperature=0.1, num_predict=300:
                        '{"op": "CLARIFY", "question": "どの列ですか？"}')
    got = ailine.translate_task("qwen2.5-coder:7b", "並べ替えて", _SAMPLE_META)
    assert got["plan"][0]["op"] == "CLARIFY"
    assert got["plan"][0]["question"] == "どの列ですか？"

def test_translate_task_invalid_json_falls_back_to_freeform(monkeypatch):
    monkeypatch.setattr(ailine, "ollama_generate_json",
                        lambda model, msgs, temperature=0.1, num_predict=300:
                        "これは JSON ではない")
    got = ailine.translate_task("qwen2.5-coder:7b", "いい感じにして", _SAMPLE_META)
    assert got["plan"][0]["op"] == "FREEFORM"

def test_translate_task_missing_required_slot_falls_back_to_freeform(monkeypatch):
    # order 欠落 → 必須 slot 不足 → FREEFORM に退避（クラッシュしない）。
    monkeypatch.setattr(ailine, "ollama_generate_json",
                        lambda model, msgs, temperature=0.1, num_predict=300:
                        '{"op": "SORT", "args": {"col": "金額"}}')
    got = ailine.translate_task("qwen2.5-coder:7b", "並べ替えて", _SAMPLE_META)
    assert got["plan"][0]["op"] == "FREEFORM"

def test_translate_task_unknown_op_falls_back_to_freeform(monkeypatch):
    monkeypatch.setattr(ailine, "ollama_generate_json",
                        lambda model, msgs, temperature=0.1, num_predict=300:
                        '{"op": "DELETE_ROW", "args": {}}')
    got = ailine.translate_task("qwen2.5-coder:7b", "行を消して", _SAMPLE_META)
    assert got["plan"][0]["op"] == "FREEFORM"

def test_translate_task_transport_failure_falls_back_to_freeform(monkeypatch):
    def boom(*a, **k):
        raise OSError("ollama 不通（テスト用）")
    monkeypatch.setattr(ailine, "ollama_generate_json", boom)
    got = ailine.translate_task("qwen2.5-coder:7b", "何かして", _SAMPLE_META)
    assert got["plan"][0]["op"] == "FREEFORM"

# --- ★ M2c: 複合計画パース (plan 配列・OUT_OF_VOCAB 保持) -----------------------

def test_translate_task_parses_multi_step_plan(monkeypatch):
    monkeypatch.setattr(ailine, "ollama_generate_json",
                        lambda model, msgs, temperature=0.1, num_predict=300:
                        '{"plan": [{"op": "SORT", "args": {"col": "金額", "order": "desc"}}, '
                        '{"op": "BOLD", "args": {"target": "row:1"}}]}')
    got = ailine.translate_task("qwen2.5-coder:7b", "金額で降順に並べ替えて見出しを太字に", _SAMPLE_META)
    assert len(got["plan"]) == 2
    assert got["plan"][0]["op"] == "SORT"
    assert got["plan"][1] == {"op": "BOLD", "args": {"target": "row:1"}}

def test_translate_task_keeps_out_of_vocab_step_with_about_not_silently_dropped(monkeypatch):
    # ★ 黙落禁止: 語彙外の段は about 付きの OUT_OF_VOCAB として必ず計画に残る。
    monkeypatch.setattr(ailine, "ollama_generate_json",
                        lambda model, msgs, temperature=0.1, num_predict=300:
                        '{"plan": [{"op": "COMPUTE_COLUMN", "args": {"operands": ["数量", "単価"], '
                        '"operator": "*", "target": "小計"}}, '
                        '{"op": "OUT_OF_VOCAB", "about": "税込み合計"}]}')
    got = ailine.translate_task("qwen2.5-coder:7b", "小計に数量×単価を入れて税込み合計も出して", _SAMPLE_META)
    assert len(got["plan"]) == 2
    assert got["plan"][0]["args"]["target"] == "小計"
    assert got["plan"][1] == {"op": "OUT_OF_VOCAB", "about": "税込み合計", "args": {}}

def test_translate_task_out_of_vocab_missing_about_gets_placeholder(monkeypatch):
    monkeypatch.setattr(ailine, "ollama_generate_json",
                        lambda model, msgs, temperature=0.1, num_predict=300:
                        '{"plan": [{"op": "OUT_OF_VOCAB"}]}')
    got = ailine.translate_task("qwen2.5-coder:7b", "いい感じにして", _SAMPLE_META)
    assert got["plan"][0]["op"] == "OUT_OF_VOCAB"
    assert got["plan"][0]["about"]   # 空文字ではない

def test_translate_task_empty_plan_array_falls_back_to_freeform(monkeypatch):
    monkeypatch.setattr(ailine, "ollama_generate_json",
                        lambda model, msgs, temperature=0.1, num_predict=300: '{"plan": []}')
    got = ailine.translate_task("qwen2.5-coder:7b", "何かして", _SAMPLE_META)
    assert got["plan"][0]["op"] == "FREEFORM"


# --- ② 検証（接地・数字表記の両解釈） -----------------------------------------

def test_resolve_col_ref_exact_name():
    v, inferred, err = ailine.resolve_col_ref("金額", ["商品", "金額", "在庫"])
    assert (v, inferred, err) == ("金額", False, None)

def test_resolve_col_ref_digit_unique_candidate_is_inferred():
    # "2" は 0起点なら在庫、1起点なら金額 → どちらも実在するが同じ列名になる場合は一意
    v, inferred, err = ailine.resolve_col_ref("0", ["商品", "金額", "在庫"])
    assert v == "商品"
    assert inferred is True
    assert err is None

def test_resolve_col_ref_digit_ambiguous_two_distinct_candidates():
    v, inferred, err = ailine.resolve_col_ref("1", ["商品", "金額", "在庫"])
    # 0起点=金額, 1起点=商品 → 二通りの実在列名に分かれるので一意に決まらない
    assert v is None
    assert "複数の解釈" in err

def test_resolve_col_ref_unknown_lists_known_columns():
    v, inferred, err = ailine.resolve_col_ref("存在しない列", ["商品", "金額"])
    assert v is None
    assert "商品" in err and "金額" in err

def test_verify_dsl_args_sort_ok():
    ok, resolved, inferred, err = ailine.verify_dsl_args("SORT", {"col": "金額", "order": "desc"}, _SAMPLE_META)
    assert ok is True
    assert resolved == {"col": "金額", "order": "desc"}
    assert inferred == set()
    assert err is None

def test_verify_dsl_args_sort_unknown_column_is_clarify_error():
    ok, resolved, inferred, err = ailine.verify_dsl_args("SORT", {"col": "存在しない", "order": "desc"}, _SAMPLE_META)
    assert ok is False
    assert "がありません" in err

def test_verify_dsl_args_sort_bad_order():
    ok, resolved, inferred, err = ailine.verify_dsl_args("SORT", {"col": "金額", "order": "up"}, _SAMPLE_META)
    assert ok is False
    assert "asc/desc" in err

def test_verify_dsl_args_compute_column_resolves_operands():
    ok, resolved, inferred, err = ailine.verify_dsl_args(
        "COMPUTE_COLUMN", {"operands": ["売上", "原価"], "operator": "-"}, _SAMPLE_META)
    assert ok is True
    assert resolved["operands"] == ["売上", "原価"]

def test_verify_dsl_args_compute_column_bad_operator():
    ok, resolved, inferred, err = ailine.verify_dsl_args(
        "COMPUTE_COLUMN", {"operands": ["売上", "原価"], "operator": "%"}, _SAMPLE_META)
    assert ok is False
    assert "演算子" in err

# --- ★ M2c: COMPUTE_COLUMN の target(名指し列への書き込み) ----------------------

def test_verify_dsl_args_compute_column_target_resolves_existing_column():
    ok, resolved, inferred, err = ailine.verify_dsl_args(
        "COMPUTE_COLUMN",
        {"operands": ["売上", "原価"], "operator": "-", "target": "金額"}, _SAMPLE_META)
    assert ok is True
    assert resolved["target"] == "金額"

def test_verify_dsl_args_compute_column_target_unknown_falls_back_to_new_column():
    # ★ W3: 実測で qwen2.5-coder:7b が「利益列を作って」の『利益』(新規列名) を target に
    #   誤って埋める頻度が高いと判明（E2E③『売上から原価を引いた利益列を作って』）。
    #   target が実在しない場合は一意性の曖昧さ(digit候補の複数一致)とは別なので、
    #   CLARIFY で止めず target 無指定＝新規列作成にフォールバックする。
    ok, resolved, inferred, err = ailine.verify_dsl_args(
        "COMPUTE_COLUMN",
        {"operands": ["売上", "原価"], "operator": "-", "target": "存在しない列"}, _SAMPLE_META)
    assert ok is True
    assert "target" not in resolved
    assert err is None

def test_verify_dsl_args_compute_column_target_ambiguous_digit_still_errors():
    # ★ W3: 実在しない場合と違い、複数解釈が可能な曖昧なケースは引き続き CLARIFY で止める
    #   （推測で断定しない原則は真に曖昧なケースにだけ残す）。
    meta = {"sheets": ["Sheet"], "headers": {"Sheet": ["列0", "列1", "列2", "列3"]}}
    ok, resolved, inferred, err = ailine.verify_dsl_args(
        "COMPUTE_COLUMN",
        {"operands": ["列0", "列1"], "operator": "-", "target": "2"}, meta)
    assert ok is False
    assert "一意に決まりません" in err

def test_verify_dsl_args_compute_column_no_target_still_ok():
    # target 無指定は従来どおり合格（新規列パス）。
    ok, resolved, inferred, err = ailine.verify_dsl_args(
        "COMPUTE_COLUMN", {"operands": ["売上", "原価"], "operator": "-"}, _SAMPLE_META)
    assert ok is True
    assert "target" not in resolved

def test_format_confirmation_line_compute_column_without_target_omits_field():
    line = ailine.format_confirmation_line(
        "COMPUTE_COLUMN", {"operands": ["売上", "原価"], "operator": "-"}, set())
    assert "対象列" not in line

def test_format_confirmation_line_compute_column_with_target_shows_it():
    line = ailine.format_confirmation_line(
        "COMPUTE_COLUMN", {"operands": ["売上", "原価"], "operator": "-", "target": "金額"}, set())
    assert "対象列:金額" in line

def test_codegen_dsl_compute_column_with_target_writes_into_existing_column():
    # ★ W3 Part3: 既定は式（setFormula）。値ベタ書きは use_formula=False（--values）で確認する。
    code = ailine.codegen_dsl(
        "COMPUTE_COLUMN",
        {"operands": ["売上", "原価"], "operator": "-", "target": "金額"}, _SAMPLE_META)
    # 金額は _SAMPLE_META の列1（0起点）。新規列(列5)には書かず、既存の列1に書く。
    assert "getCellByPosition(1, i).setFormula" in code
    assert '"D" & (i + 1) & "-" & "E" & (i + 1)' in code   # 売上=D列(3+1) 原価=E列(4+1)
    assert 'setString("売上-原価")' not in code   # 見出しは上書きしない(既存のまま)
    assert ailine.valid_signature(code)

def test_codegen_dsl_compute_column_values_mode_writes_static_getvalue():
    # ★ W3 Part3: --values（use_formula=False）は旧来の値ベタ書きのまま。
    code = ailine.codegen_dsl(
        "COMPUTE_COLUMN",
        {"operands": ["売上", "原価"], "operator": "-", "target": "金額"}, _SAMPLE_META,
        use_formula=False)
    assert "getCellByPosition(1, i).setValue" in code
    assert "setFormula" not in code

def test_verify_dsl_args_lookup_fill_ok():
    meta = {"sheets": ["明細", "単価表"],
            "headers": {"明細": ["商品", "数量", "単価"], "単価表": ["商品", "単価"]}}
    ok, resolved, inferred, err = ailine.verify_dsl_args(
        "LOOKUP_FILL",
        {"target_sheet": "明細", "target_col": "単価", "source_sheet": "単価表", "key_col": "商品"},
        meta)
    assert ok is True
    assert resolved["target_col"] == "単価"

def test_verify_dsl_args_lookup_fill_rejects_non_first_sheet_target():
    meta = {"sheets": ["明細", "単価表"],
            "headers": {"明細": ["商品", "数量", "単価"], "単価表": ["商品", "単価"]}}
    ok, resolved, inferred, err = ailine.verify_dsl_args(
        "LOOKUP_FILL",
        {"target_sheet": "単価表", "target_col": "単価", "source_sheet": "明細", "key_col": "商品"},
        meta)
    assert ok is False
    assert "1枚目" in err

def test_verify_dsl_args_fill_color_unknown_color():
    ok, resolved, inferred, err = ailine.verify_dsl_args(
        "FILL_COLOR", {"target": "col:金額", "color": "虹色"}, _SAMPLE_META)
    assert ok is False
    assert "未対応" in err

def test_verify_dsl_args_center_align_all_rejected_for_bold():
    ok, resolved, inferred, err = ailine.verify_dsl_args("BOLD", {"target": "all"}, _SAMPLE_META)
    assert ok is False

def test_verify_dsl_args_merge_bad_range_format():
    ok, resolved, inferred, err = ailine.verify_dsl_args("MERGE", {"range": "not-a-range"}, _SAMPLE_META)
    assert ok is False
    assert "形式" in err

def test_verify_dsl_args_bold_col_marks_digit_resolution_as_inferred():
    ok, resolved, inferred, err = ailine.verify_dsl_args("BOLD", {"target": "col:0"}, _SAMPLE_META)
    assert ok is True
    assert resolved["target"] == "col:商品"
    assert "target" in inferred


# --- ③ 確認行 ----------------------------------------------------------------

def test_format_confirmation_line_sort():
    line = ailine.format_confirmation_line("SORT", {"col": "金額", "order": "desc"}, set())
    assert line == "解釈: 操作:並べ替え 対象:金額 順:降順"

def test_format_confirmation_line_marks_inferred_arg():
    line = ailine.format_confirmation_line("SORT", {"col": "金額", "order": "desc"}, {"col"})
    assert "対象:金額(推定)" in line

def test_format_confirmation_line_lookup_fill():
    line = ailine.format_confirmation_line(
        "LOOKUP_FILL",
        {"target_sheet": "明細", "target_col": "単価", "source_sheet": "単価表", "key_col": "商品"}, set())
    assert line == "解釈: 操作:転記 対象シート:明細 対象列:単価 参照シート:単価表 キー列:商品"


# --- ④ 決定論 codegen ---------------------------------------------------------

def test_codegen_dsl_sort_calls_helper_with_zero_based_index():
    # ★ W3: SortByColumn は headerRow(0起点) を新たな第2引数に取る。_SAMPLE_META は
    #   header_rows を持たない旧テスト値＝既定1行目(hr0=0)になる。
    code = ailine.codegen_dsl("SORT", {"col": "金額", "order": "desc"}, _SAMPLE_META)
    assert "Call SortByColumn(oDoc, 0, 4, 1, False)" in code   # lastCol=4 (_SAMPLE_META は5列)
    assert ailine.valid_signature(code)
    assert not ailine.is_truncated_code(code)

def test_codegen_dsl_sort_ascending():
    code = ailine.codegen_dsl("SORT", {"col": "在庫", "order": "asc"}, _SAMPLE_META)
    assert "Call SortByColumn(oDoc, 0, 4, 2, True)" in code

def test_codegen_dsl_lookup_fill_calls_helper():
    meta = {"sheets": ["明細", "単価表"],
            "headers": {"明細": ["商品", "数量", "単価"], "単価表": ["商品", "単価"]}}
    code = ailine.codegen_dsl(
        "LOOKUP_FILL",
        {"target_sheet": "明細", "target_col": "単価", "source_sheet": "単価表", "key_col": "商品"}, meta)
    assert 'Call VLookupFromTable(oDoc, 0, 0, 2, "単価表")' in code

def test_codegen_dsl_aggregate_calls_helper():
    code = ailine.codegen_dsl("AGGREGATE", {"group_col": "商品", "value_col": "売上"}, _SAMPLE_META)
    assert "Call SummaryTable(oDoc, 0, 0, 3)" in code

def test_codegen_dsl_number_format_calls_helper():
    code = ailine.codegen_dsl("NUMBER_FORMAT", {"col": "金額", "style": "thousands"}, _SAMPLE_META)
    assert "Call FormatThousands(oDoc, 0, 1)" in code

def test_codegen_dsl_merge_converts_a1_range_to_zero_based():
    code = ailine.codegen_dsl("MERGE", {"range": "A1:E1"}, _SAMPLE_META)
    assert "Call MergeCells(oDoc, 0, 0, 4, 0)" in code

def test_codegen_dsl_chart_calls_helper():
    code = ailine.codegen_dsl("CHART", {"value_col": "金額"}, _SAMPLE_META)
    assert "Call InsertBarChart(oDoc, 0, 1)" in code

def test_codegen_dsl_center_align_all_calls_helper():
    code = ailine.codegen_dsl("CENTER_ALIGN", {"target": "all"}, _SAMPLE_META)
    assert "Call AlignCenter(oDoc, 0, 4)" in code   # lastCol=4 (_SAMPLE_META は5列)

def test_codegen_dsl_center_align_col_writes_template():
    code = ailine.codegen_dsl("CENTER_ALIGN", {"target": "col:在庫"}, _SAMPLE_META)
    assert "CellHoriJustify.CENTER" in code
    assert "getCellRangeByPosition(2, 0, 2, lastRow)" in code
    assert ailine.valid_signature(code)

def test_codegen_dsl_bold_row_calls_helper_with_scanned_range():
    code = ailine.codegen_dsl("BOLD", {"target": "row:1"}, _SAMPLE_META)
    assert "Call StyleBold(oDoc, 0, 0, 4, 0)" in code   # ★ W3: lastCol は走査でなく接地済み列数から決定論的に決まる

def test_codegen_dsl_bold_col_calls_helper_with_scanned_range():
    code = ailine.codegen_dsl("BOLD", {"target": "col:商品"}, _SAMPLE_META)
    assert "Call StyleBold(oDoc, 0, 0, 0, lastRow)" in code

def test_codegen_dsl_fill_color_row_writes_hex_literal():
    code = ailine.codegen_dsl("FILL_COLOR", {"target": "row:1", "color": "yellow"}, _SAMPLE_META)
    assert "&HFFFF00&" in code

def test_codegen_dsl_fill_color_col_writes_hex_literal():
    code = ailine.codegen_dsl("FILL_COLOR", {"target": "col:在庫", "color": "red"}, _SAMPLE_META)
    assert "&HFF0000&" in code
    assert "getCellByPosition(2, r)" in code

def test_codegen_dsl_compute_column_writes_new_column_at_end():
    # ★ W3 Part3: 既定は式。値ベタ書きの回帰は use_formula=False 側で確認する。
    code = ailine.codegen_dsl("COMPUTE_COLUMN", {"operands": ["売上", "原価"], "operator": "-"}, _SAMPLE_META)
    assert 'setString("売上-原価")' in code
    assert "getCellByPosition(5, 0)" in code   # 既存5列(0..4)の次=列5
    assert 'getCellByPosition(5, i).setFormula("=" & "D" & (i + 1) & "-" & "E" & (i + 1))' in code
    assert ailine.valid_signature(code)

def test_codegen_dsl_compute_column_values_mode_new_column_uses_getvalue():
    code = ailine.codegen_dsl("COMPUTE_COLUMN", {"operands": ["売上", "原価"], "operator": "-"},
                               _SAMPLE_META, use_formula=False)
    assert 'setString("売上-原価")' in code
    assert "getCellByPosition(5, 0)" in code
    assert "getCellByPosition(3, i).getValue() - oSheet.getCellByPosition(4, i).getValue()" in code
    assert ailine.valid_signature(code)
    # ★ is_truncated_code() は未検証: "Exit Sub" 直後の改行+識別子("...Exit Sub\n    oSheet")を
    #   \s+ が跨いで新しい Sub 開始と誤認する既存の罠がある（DSL 経路では is_truncated_code を
    #   呼んでいないため実害は無いが、ここでは対象外として素通りする）。


# --- ⑥ op 別事後条件（達成の機械検証） ----------------------------------------
# ★ 止血1/2（bench/realworld/BASELINE.md の D/C②検体の根治）: 各 check_* /
#   run_postcondition の戻り値を (ok: bool, reason: str) から (status: str, reason: str)
#   に変更した（status ∈ {"pass","warn","fail","error"}）。検証対象が0件/意味を
#   持たない少数のケースを「合格(True)」で素通りさせないための3値化。
#   以下の既存テストは ok is True/False を status == "pass"/"fail" に機械的に置換
#   （通常ケース＝検証対象が十分ある場合は判定結果そのものは変わらない）。

def _wb_save(tmp_path, build_fn, name="p.xlsx"):
    p = tmp_path / name
    wb = openpyxl.Workbook()
    build_fn(wb)
    wb.save(p)
    return p

def test_check_sort_passes_when_sorted_desc(tmp_path):
    p = _book(tmp_path, [["商品", "金額"], ["a", 300], ["b", 200], ["c", 100]])
    status, reason = ailine.check_sort(p, {"col": "金額", "order": "desc"})
    assert status == "pass"

def test_check_sort_fails_when_not_sorted(tmp_path):
    p = _book(tmp_path, [["商品", "金額"], ["a", 100], ["b", 300], ["c", 200]])
    status, reason = ailine.check_sort(p, {"col": "金額", "order": "desc"})
    assert status == "fail"
    assert "並んでいない" in reason

def test_check_compute_column_passes_when_values_match(tmp_path):
    p = _book(tmp_path, [["売上", "原価", "売上-原価"], [500, 300, 200], [900, 400, 500]])
    status, reason = ailine.check_compute_column(p, {"operands": ["売上", "原価"], "operator": "-"})
    assert status == "pass"

def test_check_compute_column_fails_when_value_wrong(tmp_path):
    p = _book(tmp_path, [["売上", "原価", "売上-原価"], [500, 300, 999]])
    status, reason = ailine.check_compute_column(p, {"operands": ["売上", "原価"], "operator": "-"})
    assert status == "fail"

def test_check_compute_column_target_checks_existing_column(tmp_path):
    # ★ M2c: target 指定時は新規列名でなく target 列そのものを検証する。
    p = _book(tmp_path, [["数量", "単価", "小計"], [2, 100, 200], [3, 150, 450]])
    status, reason = ailine.check_compute_column(
        p, {"operands": ["数量", "単価"], "operator": "*", "target": "小計"})
    assert status == "pass"

def test_check_compute_column_target_fails_when_target_value_wrong(tmp_path):
    p = _book(tmp_path, [["数量", "単価", "小計"], [2, 100, 999]])
    status, reason = ailine.check_compute_column(
        p, {"operands": ["数量", "単価"], "operator": "*", "target": "小計"})
    assert status == "fail"

def test_check_lookup_fill_passes_when_all_transcribed(tmp_path):
    wb = openpyxl.Workbook()
    ws1 = wb.active; ws1.title = "明細"
    for row in [["商品", "数量", "単価"], ["りんご", 2, 100], ["バナナ", 3, 200]]:
        ws1.append(row)
    ws2 = wb.create_sheet("単価表")
    for row in [["商品", "単価"], ["りんご", 100], ["バナナ", 200]]:
        ws2.append(row)
    p = tmp_path / "lookup.xlsx"
    wb.save(p)
    args = {"target_sheet": "明細", "target_col": "単価", "source_sheet": "単価表", "key_col": "商品"}
    status, reason = ailine.check_lookup_fill(p, args)
    assert status == "pass"

def test_check_lookup_fill_fails_when_value_missing_a_row(tmp_path):
    wb = openpyxl.Workbook()
    ws1 = wb.active; ws1.title = "明細"
    for row in [["商品", "数量", "単価"], ["りんご", 2, 100], ["バナナ", 3, None]]:
        ws1.append(row)
    ws2 = wb.create_sheet("単価表")
    for row in [["商品", "単価"], ["りんご", 100], ["バナナ", 200]]:
        ws2.append(row)
    p = tmp_path / "lookup.xlsx"
    wb.save(p)
    args = {"target_sheet": "明細", "target_col": "単価", "source_sheet": "単価表", "key_col": "商品"}
    status, reason = ailine.check_lookup_fill(p, args)
    assert status == "fail"

def test_check_aggregate_passes_when_sums_correct(tmp_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    for row in [["部門", "金額"], ["営業", 100], ["営業", 200], ["経理", 50]]:
        ws.append(row)
    out = wb.create_sheet("集計")
    out.append(["部門", "合計 - 金額"])
    out.append(["営業", 300])
    out.append(["経理", 50])
    out.append(["合計", 350])
    p = tmp_path / "agg.xlsx"
    wb.save(p)
    status, reason = ailine.check_aggregate(p, {"group_col": "部門", "value_col": "金額"})
    assert status == "pass"

def test_check_aggregate_fails_when_sum_wrong(tmp_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    for row in [["部門", "金額"], ["営業", 100], ["営業", 200]]:
        ws.append(row)
    out = wb.create_sheet("集計")
    out.append(["部門", "合計 - 金額"])
    out.append(["営業", 999])
    p = tmp_path / "agg.xlsx"
    wb.save(p)
    status, reason = ailine.check_aggregate(p, {"group_col": "部門", "value_col": "金額"})
    assert status == "fail"

def test_check_aggregate_fails_when_sheet_missing(tmp_path):
    p = _book(tmp_path, [["部門", "金額"], ["営業", 100]])
    status, reason = ailine.check_aggregate(p, {"group_col": "部門", "value_col": "金額"})
    assert status == "fail"
    assert "集計" in reason

def test_check_bold_row_passes_when_all_bold(tmp_path):
    from openpyxl.styles import Font
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["商品", "金額"])
    ws.append(["a", 1])
    ws["A1"].font = Font(bold=True)
    ws["B1"].font = Font(bold=True)
    p = tmp_path / "b.xlsx"
    wb.save(p)
    status, reason = ailine.check_bold(p, {"target": "row:1"})
    assert status == "pass"

def test_check_bold_row_fails_when_partial(tmp_path):
    from openpyxl.styles import Font
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["商品", "金額"])
    ws["A1"].font = Font(bold=True)   # B1 は太字にしない
    p = tmp_path / "b.xlsx"
    wb.save(p)
    status, reason = ailine.check_bold(p, {"target": "row:1"})
    assert status == "fail"

def test_check_bold_col_passes_when_all_bold(tmp_path):
    from openpyxl.styles import Font
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["商品", "金額"])
    ws.append(["a", 1])
    ws.append(["b", 2])
    for r in (1, 2, 3):
        ws.cell(row=r, column=1).font = Font(bold=True)
    p = tmp_path / "b.xlsx"
    wb.save(p)
    status, reason = ailine.check_bold(p, {"target": "col:商品"})
    assert status == "pass"

def test_check_fill_color_passes_when_matching(tmp_path):
    from openpyxl.styles import PatternFill
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["商品", "在庫"])
    ws.append(["a", 3])
    for r in (1, 2):
        ws.cell(row=r, column=2).fill = PatternFill("solid", fgColor="FFFF00")
    p = tmp_path / "b.xlsx"
    wb.save(p)
    status, reason = ailine.check_fill_color(p, {"target": "col:在庫", "color": "yellow"})
    assert status == "pass"

def test_check_fill_color_fails_when_not_colored(tmp_path):
    p = _book(tmp_path, [["商品", "在庫"], ["a", 3]])
    status, reason = ailine.check_fill_color(p, {"target": "col:在庫", "color": "yellow"})
    assert status == "fail"

def test_check_number_format_passes_when_thousands_applied(tmp_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["商品", "金額"])
    ws.append(["a", 1000])
    ws["B2"].number_format = "#,##0"
    p = tmp_path / "b.xlsx"
    wb.save(p)
    status, reason = ailine.check_number_format(p, {"col": "金額", "style": "thousands"})
    assert status == "pass"

def test_check_number_format_fails_when_not_applied(tmp_path):
    p = _book(tmp_path, [["商品", "金額"], ["a", 1000]])
    status, reason = ailine.check_number_format(p, {"col": "金額", "style": "thousands"})
    assert status == "fail"

def test_check_merge_passes_when_range_merged(tmp_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["a", "b", "c"])
    ws.merge_cells("A1:C1")
    p = tmp_path / "b.xlsx"
    wb.save(p)
    status, reason = ailine.check_merge(p, {"range": "A1:C1"})
    assert status == "pass"

def test_check_merge_fails_when_not_merged(tmp_path):
    p = _book(tmp_path, [["a", "b", "c"]])
    status, reason = ailine.check_merge(p, {"range": "A1:C1"})
    assert status == "fail"

def test_check_chart_passes_when_count_plus_one(monkeypatch, tmp_path):
    p = tmp_path / "c.xlsx"
    p.write_bytes(b"dummy")
    monkeypatch.setattr(ailine, "_charts_count", lambda path: 1)
    status, reason = ailine.check_chart(p, 0)
    assert status == "pass"

def test_check_chart_fails_when_count_unchanged(monkeypatch, tmp_path):
    p = tmp_path / "c.xlsx"
    p.write_bytes(b"dummy")
    monkeypatch.setattr(ailine, "_charts_count", lambda path: 0)
    status, reason = ailine.check_chart(p, 0)
    assert status == "fail"

def test_check_center_align_all_passes(tmp_path):
    from openpyxl.styles import Alignment
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["a", "b"])
    ws.append([1, 2])
    for r in (1, 2):
        for c in (1, 2):
            ws.cell(row=r, column=c).alignment = Alignment(horizontal="center")
    p = tmp_path / "b.xlsx"
    wb.save(p)
    status, reason = ailine.check_center_align(p, {"target": "all"})
    assert status == "pass"

def test_check_center_align_col_fails_when_not_centered(tmp_path):
    p = _book(tmp_path, [["a", "b"], [1, 2]])
    status, reason = ailine.check_center_align(p, {"target": "col:a"})
    assert status == "fail"

def test_run_postcondition_dispatches_by_op(tmp_path):
    p = _book(tmp_path, [["商品", "金額"], ["a", 300], ["b", 200]])
    status, reason = ailine.run_postcondition("SORT", p, {"col": "金額", "order": "desc"})
    assert status == "pass"


# ---------------------------------------------------------------------------
# ★ 止血: 空虚な検証合格の禁止（対象0件/少数を「合格」にしない）・None安全・
#   MAX_ROWS 切り詰め表示の正直な出し分け（bench/realworld/BASELINE.md の D/C②/B 検体）
# ---------------------------------------------------------------------------

# --- check_sort: 0件/1件は「合格」にしない・非数値は除外して件数を表示 -----------

def test_check_sort_fails_when_zero_data_rows(tmp_path):
    # ★ D検体の根治対象: 見出しだけで実データ行が無い(=検証対象0件)を「行数が少なく
    #   比較不要」で素通ししていた旧挙動を止める。
    p = _book(tmp_path, [["商品", "金額"]])
    status, reason = ailine.check_sort(p, {"col": "金額", "order": "desc"})
    assert status == "fail"
    assert "検証対象が0件" in reason

def test_check_sort_warns_when_only_one_row(tmp_path):
    p = _book(tmp_path, [["商品", "金額"], ["a", 100]])
    status, reason = ailine.check_sort(p, {"col": "金額", "order": "desc"})
    assert status == "warn"
    assert "1行のみ" in reason

def test_check_sort_excludes_non_numeric_rows_and_still_passes(tmp_path):
    # ★ C②検体の根治対象: 合計行のように key_col(A列)は埋まっているが対象列が空欄
    #   (None)の行は比較から除外し、除外件数を表示する。None>=int でクラッシュしない。
    p = _book(tmp_path, [["商品", "数量"], ["a", 30], ["b", 20], ["c", 10], ["合計", None]])
    status, reason = ailine.check_sort(p, {"col": "数量", "order": "desc"})
    assert status == "pass"
    assert "3 行を検証" in reason
    assert "数値でない 1 行は対象外" in reason

def test_check_sort_fails_when_all_rows_non_numeric(tmp_path):
    p = _book(tmp_path, [["商品", "備考"], ["a", "x"], ["b", "y"]])
    status, reason = ailine.check_sort(p, {"col": "備考", "order": "desc"})
    assert status == "fail"
    assert "検証対象が0件" in reason
    assert "数値でない 2 行は対象外" in reason

# --- check_compute_column: 0件は合格にしない・演算対象が空欄の行は対象外 --------

def test_check_compute_column_fails_when_zero_data_rows(tmp_path):
    p = _book(tmp_path, [["売上", "原価", "売上-原価"]])
    status, reason = ailine.check_compute_column(p, {"operands": ["売上", "原価"], "operator": "-"})
    assert status == "fail"
    assert "検証対象が0件" in reason

def test_check_compute_column_excludes_blank_total_row_and_still_passes(tmp_path):
    p = _book(tmp_path, [["区分", "売上", "原価", "売上-原価"],
                          ["a", 500, 300, 200], ["b", 900, 400, 500],
                          ["合計", None, None, 700]])
    status, reason = ailine.check_compute_column(
        p, {"operands": ["売上", "原価"], "operator": "-"})
    assert status == "pass"
    assert "2 行を検証" in reason
    assert "数値でない 1 行は対象外" in reason

def test_check_compute_column_fails_when_all_rows_blank(tmp_path):
    p = _book(tmp_path, [["区分", "売上", "原価", "売上-原価"], ["合計", None, None, None]])
    status, reason = ailine.check_compute_column(
        p, {"operands": ["売上", "原価"], "operator": "-"})
    assert status == "fail"
    assert "検証対象が0件" in reason

# --- check_lookup_fill: 対象シートに行が0件 --------------------------------------

def test_check_lookup_fill_fails_when_target_sheet_has_no_rows(tmp_path):
    wb = openpyxl.Workbook()
    ws1 = wb.active; ws1.title = "明細"
    ws1.append(["商品", "数量", "単価"])
    ws2 = wb.create_sheet("単価表")
    for row in [["商品", "単価"], ["りんご", 100]]:
        ws2.append(row)
    p = tmp_path / "lookup.xlsx"
    wb.save(p)
    args = {"target_sheet": "明細", "target_col": "単価", "source_sheet": "単価表", "key_col": "商品"}
    status, reason = ailine.check_lookup_fill(p, args)
    assert status == "fail"
    assert "検証対象が0件" in reason

# --- check_aggregate: 元データが0件 ----------------------------------------------

def test_check_aggregate_fails_when_source_has_no_rows(tmp_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["部門", "金額"])
    out = wb.create_sheet("集計")
    out.append(["部門", "合計 - 金額"])
    p = tmp_path / "agg.xlsx"
    wb.save(p)
    status, reason = ailine.check_aggregate(p, {"group_col": "部門", "value_col": "金額"})
    assert status == "fail"
    assert "検証対象が0件" in reason

# --- check_bold/check_center_align: 検証対象0件(見出しすら無い空シート) ----------

def test_check_bold_fails_when_sheet_completely_empty(tmp_path):
    wb = openpyxl.Workbook()
    p = tmp_path / "empty.xlsx"
    wb.save(p)
    status, reason = ailine.check_bold(p, {"target": "row:1"})
    assert status == "fail"
    assert "検証対象が0件" in reason

def test_check_center_align_fails_when_sheet_completely_empty(tmp_path):
    wb = openpyxl.Workbook()
    p = tmp_path / "empty.xlsx"
    wb.save(p)
    status, reason = ailine.check_center_align(p, {"target": "all"})
    assert status == "fail"
    assert "検証対象が0件" in reason

# --- check_number_format: データ行0件 --------------------------------------------

def test_check_number_format_fails_when_zero_data_rows(tmp_path):
    p = _book(tmp_path, [["商品", "金額"]])
    status, reason = ailine.check_number_format(p, {"col": "金額", "style": "thousands"})
    assert status == "fail"
    assert "検証対象が0件" in reason

# --- run_postcondition: チェッカー内例外は生トレースバックを出さず"error"に変換 --

def test_run_postcondition_catches_checker_exception(tmp_path, monkeypatch):
    p = _book(tmp_path, [["商品", "金額"], ["a", 1]])

    def boom(path, args):
        raise TypeError("boom")

    monkeypatch.setitem(ailine.POSTCONDITIONS, "SORT", boom)
    status, reason = ailine.run_postcondition("SORT", p, {"col": "金額", "order": "desc"})
    assert status == "error"
    assert "事後条件の検証に失敗" in reason
    assert "TypeError" in reason

def test_cmd_run_dsl_postcondition_warn_does_not_claim_verified(tmp_path, monkeypatch, capsys):
    # ★ 止血1: 検証対象が1行のみの SORT は「機械検証済み」と名乗らず ⚠ で報告する。
    book = _book(tmp_path, [["商品", "金額"], ["a", 100]])
    monkeypatch.setattr(ailine, "HISTORY_FILE", tmp_path / "history.jsonl")
    monkeypatch.setattr(ailine, "translate_task",
                        lambda model, task, book_meta, temperature=0.1:
                        {"op": "SORT", "args": {"col": "金額", "order": "desc"}})
    monkeypatch.setattr(ailine, "normalize_book", lambda book, workdir, timeout=None: book)
    monkeypatch.setattr(ailine, "basrun_apply",
                        lambda out_book, code, workdir, helper_files=(), timeout=None:
                        (True, None, "ok"))
    ns = argparse.Namespace(
        book=str(book), task="金額で降順に並べ替えて", model="qwen2.5-coder:7b",
        refs=None, helpers=None, repair=0, temperature=0.2,
        dry=False, inplace=False, json=False, timeout=180.0, ask=False)
    rc = ailine.cmd_run(ns)
    captured = capsys.readouterr()
    assert rc == 0
    assert "⚠" in captured.out
    assert "機械検証済み" not in captured.out

# --- MAX_ROWS 切り詰め表示の正直な出し分け（B検体所見の根治） --------------------

def test_truncation_notice_is_none_when_not_truncated():
    assert ailine._truncation_notice({"truncated": False}, {"truncated": False}, True) is None

def test_truncation_notice_dsl_path_says_verification_covers_all_rows():
    msg = ailine._truncation_notice({"truncated": False}, {"truncated": True},
                                     exhaustive_postcondition=True)
    assert "検証・適用は全行に対して実施" in msg

def test_truncation_notice_freeform_path_says_verification_also_truncated():
    msg = ailine._truncation_notice({"truncated": True}, {"truncated": False},
                                     exhaustive_postcondition=False)
    assert "検証も先頭" in msg
    assert "行のみ" in msg

def test_snapshot_marks_truncated_when_rows_exceed_max_rows(tmp_path, monkeypatch):
    monkeypatch.setattr(ailine, "MAX_ROWS", 3)
    p = _book(tmp_path, [["a"], [1], [2], [3], [4], [5]])
    snap = ailine.snapshot(p)
    assert snap["truncated"] is True

def test_snapshot_not_truncated_when_within_max_rows(tmp_path):
    p = _book(tmp_path, [["a"], [1], [2]])
    snap = ailine.snapshot(p)
    assert snap["truncated"] is False

def test_cmd_run_dsl_prints_truncation_notice_when_snapshot_truncated(tmp_path, monkeypatch, capsys):
    book = _book(tmp_path, [["商品", "金額"], ["a", 300], ["b", 200], ["c", 100]])
    monkeypatch.setattr(ailine, "MAX_ROWS", 2)   # 3行あるデータを2行に切り詰めさせる
    monkeypatch.setattr(ailine, "HISTORY_FILE", tmp_path / "history.jsonl")
    monkeypatch.setattr(ailine, "translate_task",
                        lambda model, task, book_meta, temperature=0.1:
                        {"op": "SORT", "args": {"col": "金額", "order": "desc"}})
    monkeypatch.setattr(ailine, "normalize_book", lambda book, workdir, timeout=None: book)
    monkeypatch.setattr(ailine, "basrun_apply",
                        lambda out_book, code, workdir, helper_files=(), timeout=None:
                        (True, None, "ok"))
    ns = argparse.Namespace(
        book=str(book), task="金額で降順に並べ替えて", model="qwen2.5-coder:7b",
        refs=None, helpers=None, repair=0, temperature=0.2,
        dry=False, inplace=False, json=False, timeout=180.0, ask=False)
    rc = ailine.cmd_run(ns)
    captured = capsys.readouterr()
    assert rc == 0
    assert "表示は先頭 2 行の変化のみ。検証・適用は全行に対して実施" in captured.out


# --- run コマンド: 翻訳の分岐（CLARIFY exit 3 / DSL 経路の事後条件不合格） -------

def test_cmd_run_clarify_prints_question_and_exits_3(tmp_path, monkeypatch, capsys):
    book = _book(tmp_path, [["商品", "金額"]])
    # ★ W3: 正規化パス(StructDump)は翻訳より前に走るので、CLARIFY 系の単体テストも
    #   normalize_book を差し替えて LibreOffice を要さないようにする。
    monkeypatch.setattr(ailine, "normalize_book", lambda book, workdir, timeout=None: book)
    monkeypatch.setattr(ailine, "translate_task",
                        lambda model, task, book_meta, temperature=0.1:
                        {"op": "CLARIFY", "question": "どの列を並べ替えますか？", "args": {}})
    ns = argparse.Namespace(
        book=str(book), task="いい感じにして", model="qwen2.5-coder:7b",
        refs=None, helpers=None, repair=0, temperature=0.2,
        dry=False, inplace=False, json=False, timeout=180.0, ask=False)
    rc = ailine.cmd_run(ns)
    captured = capsys.readouterr()
    assert rc == 3
    assert "どの列を並べ替えますか？" in captured.out

def test_cmd_run_dsl_verification_failure_falls_back_to_clarify_exit_3(tmp_path, monkeypatch, capsys):
    book = _book(tmp_path, [["商品", "金額"]])
    monkeypatch.setattr(ailine, "normalize_book", lambda book, workdir, timeout=None: book)
    monkeypatch.setattr(ailine, "translate_task",
                        lambda model, task, book_meta, temperature=0.1:
                        {"op": "SORT", "args": {"col": "存在しない列", "order": "desc"}})
    ns = argparse.Namespace(
        book=str(book), task="存在しない列で並べ替えて", model="qwen2.5-coder:7b",
        refs=None, helpers=None, repair=0, temperature=0.2,
        dry=False, inplace=False, json=False, timeout=180.0, ask=False)
    rc = ailine.cmd_run(ns)
    captured = capsys.readouterr()
    assert rc == 3
    assert "がありません" in captured.out

def test_cmd_run_dsl_dry_shows_confirmation_and_code_without_applying(tmp_path, monkeypatch, capsys):
    book = _book(tmp_path, [["商品", "金額"]])
    monkeypatch.setattr(ailine, "HISTORY_FILE", tmp_path / "history.jsonl")   # 実 history を汚さない
    monkeypatch.setattr(ailine, "translate_task",
                        lambda model, task, book_meta, temperature=0.1:
                        {"op": "SORT", "args": {"col": "金額", "order": "desc"}})
    ns = argparse.Namespace(
        book=str(book), task="金額で降順に並べ替えて", model="qwen2.5-coder:7b",
        refs=None, helpers=None, repair=0, temperature=0.2,
        dry=True, inplace=False, json=True, timeout=180.0, ask=False)
    rc = ailine.cmd_run(ns)
    captured = capsys.readouterr()
    assert rc == 0
    assert "解釈: 操作:並べ替え 対象:金額 順:降順" in captured.out
    assert "Call SortByColumn" in captured.out
    assert '"path": "dsl"' in captured.out

def test_cmd_run_dsl_postcondition_failure_returns_1(tmp_path, monkeypatch, capsys):
    # basrun_apply/snapshot を差し替え、事後条件が満たされない場合に exit 1 になることを確認。
    book = _book(tmp_path, [["商品", "金額"], ["a", 100], ["b", 200]])
    monkeypatch.setattr(ailine, "HISTORY_FILE", tmp_path / "history.jsonl")   # 実 history を汚さない
    monkeypatch.setattr(ailine, "translate_task",
                        lambda model, task, book_meta, temperature=0.1:
                        {"op": "SORT", "args": {"col": "金額", "order": "desc"}})
    monkeypatch.setattr(ailine, "normalize_book", lambda book, workdir, timeout=None: book)
    monkeypatch.setattr(ailine, "basrun_apply",
                        lambda out_book, code, workdir, helper_files=(), timeout=None:
                        (True, None, "ok"))
    # basrun_apply は成功したことにするが、実際には out_book の中身は昇順のまま(=事後条件不合格)。
    ns = argparse.Namespace(
        book=str(book), task="金額で降順に並べ替えて", model="qwen2.5-coder:7b",
        refs=None, helpers=None, repair=0, temperature=0.2,
        dry=False, inplace=False, json=False, timeout=180.0, ask=False)
    rc = ailine.cmd_run(ns)
    captured = capsys.readouterr()
    assert rc == 1
    assert "事後条件を満たさない" in captured.out


# ---------------------------------------------------------------------------
# ★ M2c: 複合依頼の計画実行と正直な範囲表示
# ---------------------------------------------------------------------------

# --- 項目別報告の整形 ---------------------------------------------------------

def test_format_plan_report_ok_warn_fail_lines():
    items = [
        (1, "操作:計算列 対象列:小計", "ok", "3 行を検証"),
        (2, "税込み合計", "warn", None),
        (3, "操作:並べ替え", "fail", "列『在庫』がありません"),
    ]
    lines = ailine.format_plan_report(items)
    assert lines[0] == "1. 操作:計算列 対象列:小計 → ✓ 機械検証済み（3 行を検証）"
    # ★ W8a 項目5: 表示文言「自由生成」→「AI が直接作成（機械保証なし）」に追従。
    assert lines[1] == "2. 税込み合計 → ⚠ 語彙外のため AI が直接作成（機械保証なし）で実行（確認してください）"
    assert lines[2] == "3. 操作:並べ替え → × 未対応: 列『在庫』がありません"

def test_format_plan_report_ok_without_detail_omits_parens():
    lines = ailine.format_plan_report([(1, "操作:太字", "ok", None)])
    assert lines[0] == "1. 操作:太字 → ✓ 機械検証済み"

# --- 総合判定規則 -------------------------------------------------------------

def test_overall_verdict_all_ok():
    line, v = ailine.overall_verdict([(1, "x", "ok", "r")])
    assert v == "ok"
    assert "すべて機械検証済み" in line

def test_overall_verdict_warn_without_fail():
    line, v = ailine.overall_verdict([(1, "x", "ok", "r"), (2, "y", "warn", None)])
    assert v == "warn"
    assert "確認が必要" in line

def test_overall_verdict_fail_dominates_over_warn():
    line, v = ailine.overall_verdict(
        [(1, "x", "ok", "r"), (2, "y", "warn", None), (3, "z", "fail", "reason")])
    assert v == "fail"

# --- 依存つき連鎖(#107 型)の新規列フォールバック --------------------------------

def test_apply_new_column_fallback_substitutes_sole_new_column():
    args = ailine._apply_new_column_fallback(
        "SORT", {"col": "利益", "order": "desc"},
        ["商品", "売上", "原価", "売上-原価"], ["売上-原価"])
    assert args["col"] == "売上-原価"

def test_apply_new_column_fallback_noop_when_multiple_candidates():
    args = ailine._apply_new_column_fallback(
        "SORT", {"col": "利益", "order": "desc"}, ["商品", "a", "b"], ["a", "b"])
    assert args["col"] == "利益"   # 候補2つ → 書き換えない(保守的)

def test_apply_new_column_fallback_leaves_existing_reference_untouched():
    args = ailine._apply_new_column_fallback(
        "SORT", {"col": "商品", "order": "desc"}, ["商品", "売上-原価"], ["売上-原価"])
    assert args["col"] == "商品"   # 既に実在する参照は書き換えない

def test_apply_new_column_fallback_noop_when_no_new_columns():
    args = ailine._apply_new_column_fallback(
        "SORT", {"col": "利益", "order": "desc"}, ["商品", "売上"], [])
    assert args["col"] == "利益"

# --- cmd_run_plan: 複合計画の実行と honest な報告 -------------------------------

def _plan_book(tmp_path, rows, name="plan_b.xlsx"):
    p = tmp_path / name
    wb = openpyxl.Workbook(); ws = wb.active
    for r in rows:
        ws.append(r)
    wb.save(p)
    return p

def test_cmd_run_plan_all_dsl_steps_pass_gives_full_verdict(tmp_path, monkeypatch, capsys):
    from openpyxl.styles import Font
    p = _plan_book(tmp_path, [["商品", "金額"], ["a", 300], ["b", 200], ["c", 100]])
    wb = openpyxl.load_workbook(p)
    ws = wb.active
    for c in (1, 2):
        ws.cell(row=1, column=c).font = Font(bold=True)   # 見出し行を先に太字にしておく(pre-condition)
    wb.save(p)

    monkeypatch.setattr(ailine, "HISTORY_FILE", tmp_path / "history.jsonl")
    monkeypatch.setattr(ailine, "translate_task",
                        lambda model, task, book_meta, temperature=0.1:
                        {"plan": [{"op": "SORT", "args": {"col": "金額", "order": "desc"}},
                                  {"op": "BOLD", "args": {"target": "row:1"}}]})
    monkeypatch.setattr(ailine, "normalize_book", lambda book, workdir, timeout=None: book)
    monkeypatch.setattr(ailine, "basrun_apply",
                        lambda out_book, code, workdir, helper_files=(), timeout=None:
                        (True, None, "ok"))
    ns = argparse.Namespace(
        book=str(p), task="金額で降順に並べ替えて見出しを太字に", model="qwen2.5-coder:7b",
        refs=None, helpers=None, repair=0, temperature=0.2,
        dry=False, inplace=False, json=True, timeout=180.0, ask=False)
    rc = ailine.cmd_run(ns)
    captured = capsys.readouterr()
    assert rc == 0
    assert "1. " in captured.out and "✓ 機械検証済み" in captured.out
    assert "2. " in captured.out
    assert "✓ すべて機械検証済み" in captured.out
    assert '"path": "plan"' in captured.out
    assert '"status": "ok"' in captured.out

def test_cmd_run_plan_mixes_dsl_success_and_freeform_warns(tmp_path, monkeypatch, capsys):
    p = _plan_book(tmp_path, [["商品", "金額"], ["a", 300], ["b", 200], ["c", 100]])

    monkeypatch.setattr(ailine, "HISTORY_FILE", tmp_path / "history.jsonl")
    monkeypatch.setattr(ailine, "translate_task",
                        lambda model, task, book_meta, temperature=0.1:
                        {"plan": [{"op": "SORT", "args": {"col": "金額", "order": "desc"}},
                                  {"op": "OUT_OF_VOCAB", "about": "条件付き書式"}]})
    monkeypatch.setattr(ailine, "normalize_book", lambda book, workdir, timeout=None: book)
    monkeypatch.setattr(ailine, "ollama_generate",
                        lambda model, msgs, temperature=0.2: "Sub Run(oDoc As Object)\nEnd Sub")

    def fake_apply(out_book, code, workdir, helper_files=(), timeout=None):
        wb2 = openpyxl.load_workbook(out_book)
        ws2 = wb2.active
        cell = ws2.cell(row=1, column=10)   # postcondition が見ない列にダミーの変化を残す
        cell.value = (cell.value or 0) + 1
        wb2.save(out_book)
        return True, None, "ok"
    monkeypatch.setattr(ailine, "basrun_apply", fake_apply)

    ns = argparse.Namespace(
        book=str(p), task="金額で降順に並べ替えて条件付き書式もつけて", model="qwen2.5-coder:7b",
        refs=None, helpers=None, repair=0, temperature=0.2,
        dry=False, inplace=False, json=False, timeout=180.0, ask=False)
    rc = ailine.cmd_run(ns)
    captured = capsys.readouterr()
    assert rc == 0   # ⚠ は失敗ではない
    assert "✓ 機械検証済み" in captured.out
    assert "条件付き書式" in captured.out
    # ★ W8a 項目5: 表示文言「自由生成」→「AI が直接作成（機械保証なし）」に追従。
    assert "⚠ 語彙外のため AI が直接作成（機械保証なし）で実行（確認してください）" in captured.out
    assert "⚠ 一部は確認が必要です" in captured.out
    assert "すべて機械検証済み" not in captured.out

def test_cmd_run_plan_all_steps_fail_grounding_gives_overall_failure(tmp_path, monkeypatch, capsys):
    p = _plan_book(tmp_path, [["商品", "金額"], ["a", 300]])
    monkeypatch.setattr(ailine, "HISTORY_FILE", tmp_path / "history.jsonl")
    monkeypatch.setattr(ailine, "translate_task",
                        lambda model, task, book_meta, temperature=0.1:
                        {"plan": [{"op": "SORT", "args": {"col": "存在しない列1", "order": "desc"}},
                                  {"op": "BOLD", "args": {"target": "col:存在しない列2"}}]})
    monkeypatch.setattr(ailine, "normalize_book", lambda book, workdir, timeout=None: book)
    ns = argparse.Namespace(
        book=str(p), task="存在しない列で並べ替えて太字にして", model="qwen2.5-coder:7b",
        refs=None, helpers=None, repair=0, temperature=0.2,
        dry=False, inplace=False, json=False, timeout=180.0, ask=False)
    rc = ailine.cmd_run(ns)
    captured = capsys.readouterr()
    assert rc == 1
    assert captured.out.count("× 未対応") == 2
    assert "達成できませんでした" in captured.out

def test_cmd_run_plan_dry_previews_without_applying(tmp_path, monkeypatch, capsys):
    p = _plan_book(tmp_path, [["商品", "金額"], ["a", 300], ["b", 200]])
    monkeypatch.setattr(ailine, "HISTORY_FILE", tmp_path / "history.jsonl")
    monkeypatch.setattr(ailine, "translate_task",
                        lambda model, task, book_meta, temperature=0.1:
                        {"plan": [{"op": "SORT", "args": {"col": "金額", "order": "desc"}},
                                  {"op": "OUT_OF_VOCAB", "about": "条件付き書式"}]})
    ns = argparse.Namespace(
        book=str(p), task="金額で降順に並べ替えて条件付き書式もつけて", model="qwen2.5-coder:7b",
        refs=None, helpers=None, repair=0, temperature=0.2,
        dry=True, inplace=False, json=True, timeout=180.0, ask=False)
    rc = ailine.cmd_run(ns)
    captured = capsys.readouterr()
    assert rc == 0
    assert "1. " in captured.out
    assert "条件付き書式" in captured.out
    out_book = p.with_name(p.stem + ".out" + p.suffix)
    assert not out_book.exists()   # --dry は適用しない
    assert '"path": "plan"' in captured.out
    assert '"dry": true' in captured.out

def test_cmd_run_plan_dependent_chaining_resolves_new_column_reference(tmp_path, monkeypatch, capsys):
    # ★ battery v2 #107 型: 「利益列を作って、利益で降順に並べ替えて」。2段目の "利益" は
    #   実在せず、1段目が作る自動命名列(売上-原価)を指す＝直前段の適用後の列構成で解決する。
    p = _plan_book(tmp_path, [["商品", "売上", "原価"], ["a", 500, 300], ["b", 900, 400]])
    monkeypatch.setattr(ailine, "HISTORY_FILE", tmp_path / "history.jsonl")
    monkeypatch.setattr(ailine, "translate_task",
                        lambda model, task, book_meta, temperature=0.1:
                        {"plan": [
                            {"op": "COMPUTE_COLUMN", "args": {"operands": ["売上", "原価"], "operator": "-"}},
                            {"op": "SORT", "args": {"col": "利益", "order": "desc"}}]})
    monkeypatch.setattr(ailine, "normalize_book", lambda book, workdir, timeout=None: book)

    def fake_apply(out_book, code, workdir, helper_files=(), timeout=None):
        wb2 = openpyxl.load_workbook(out_book)
        ws2 = wb2.active
        if "SortByColumn" in code:
            ws2.cell(row=1, column=1, value="商品"); ws2.cell(row=1, column=2, value="売上")
            ws2.cell(row=1, column=3, value="原価"); ws2.cell(row=1, column=4, value="売上-原価")
            ws2.cell(row=2, column=1, value="b"); ws2.cell(row=2, column=2, value=900)
            ws2.cell(row=2, column=3, value=400); ws2.cell(row=2, column=4, value=500)
            ws2.cell(row=3, column=1, value="a"); ws2.cell(row=3, column=2, value=500)
            ws2.cell(row=3, column=3, value=300); ws2.cell(row=3, column=4, value=200)
        else:
            ws2.cell(row=1, column=4, value="売上-原価")
            ws2.cell(row=2, column=4, value=200)   # a: 500-300
            ws2.cell(row=3, column=4, value=500)   # b: 900-400
        wb2.save(out_book)
        return True, None, "ok"
    monkeypatch.setattr(ailine, "basrun_apply", fake_apply)

    ns = argparse.Namespace(
        book=str(p), task="売上から原価を引いた利益列を作って、利益で降順に並べ替えて",
        model="qwen2.5-coder:7b", refs=None, helpers=None, repair=0, temperature=0.2,
        dry=False, inplace=False, json=False, timeout=180.0, ask=False,
        # ★ W3 Part3: fake_apply は静的な値を直接書き込む(式は書かない)ので、この
        #   テストは --values（値ベタ書き）経路として実行する。式検証(二層)は
        #   check_compute_column の専用ユニットテストで別途カバーする。
        values=True)
    rc = ailine.cmd_run(ns)
    captured = capsys.readouterr()
    assert rc == 0
    assert "列『利益』がありません" not in captured.out
    assert "✓ すべて機械検証済み" in captured.out


# ---------------------------------------------------------------------------
# ★ W3 Part1/2: StructDump（LibreOffice の目で構造を読む）+ 見出し行推定
#   ★ LO 依存部（実使用範囲の取得）は fixture の dict/テキストで代用する
#   （normalize_book/basrun_apply は差し替え、StructDump のパーサ・ヒューリスティクス
#   だけを純ロジックとして検証する）。実機での動作は E2E ログで別途確認する。
# ---------------------------------------------------------------------------

def test_parse_structdump_raw_parses_tab_delimited_lines():
    text = "SHEET\tSheet\t0\t0\t4\t9\t0\t1\t0\nSHEET\t単価表\t0\t0\t1\t3\t0\t0\t0\n"
    got = ailine.parse_structdump_raw(text)
    assert got["Sheet"]["used_range"] == {"start_col": 0, "start_row": 0, "end_col": 4, "end_row": 9}
    assert got["Sheet"]["charts"] == 1
    assert got["単価表"]["shapes"] == 0

def test_parse_structdump_raw_ignores_malformed_lines():
    text = "not a sheet line\nSHEET\ttoo\tfew\tcolumns\n"
    assert ailine.parse_structdump_raw(text) == {}

def test_parse_structdump_raw_empty_text_gives_empty_dict():
    assert ailine.parse_structdump_raw("") == {}

def test_row_char_stats_counts_nonempty_str_and_bold(tmp_path):
    from openpyxl.styles import Font
    p = tmp_path / "b.xlsx"
    wb = openpyxl.Workbook(); ws = wb.active
    ws.append(["商品", "金額", "在庫"])
    ws.append(["りんご", 100, 5])
    ws["A1"].font = Font(bold=True)
    ws["B1"].font = Font(bold=True)
    wb.save(p)
    wb2 = openpyxl.load_workbook(p)
    stats = ailine._row_char_stats(wb2.active, 1, 2, 1, 3)
    assert stats[1] == {"nonempty": 3, "str": 3, "bold": 2}
    assert stats[2] == {"nonempty": 3, "str": 1, "bold": 0}   # りんご=文字列, 100/5=数値

def test_build_struct_dump_falls_back_to_openpyxl_when_no_raw_dump(tmp_path, monkeypatch):
    # ★ normalize_book が差し替えられて structdump.txt が書かれない場合(単体テストの通常経路)、
    #   openpyxl の max_row/max_column から used_range を推定する（CLARIFY を誤って出さないため）。
    p = _book(tmp_path, [["商品", "金額"], ["a", 100], ["b", 200]])
    workdir = tmp_path / "wd"
    workdir.mkdir()
    dump = ailine.build_struct_dump(p, workdir)
    sheet = dump["sheets"]["Sheet"]
    assert sheet["used_range"]["start_row"] == 1
    assert sheet["rows"][1] == {"nonempty": 2, "str": 2, "bold": 0}
    assert sheet["rows"][2]["nonempty"] == 2 and sheet["rows"][2]["str"] == 1

def test_detect_header_row_simple_single_level_header():
    # 普通の帳票（demo/sales.xlsx 型）: 行1=見出し、行2以降=数値混在データ。
    sheet_struct = {"rows": {
        1: {"nonempty": 2, "str": 2, "bold": 0},
        2: {"nonempty": 2, "str": 1, "bold": 0},
        3: {"nonempty": 2, "str": 1, "bold": 0},
    }}
    row, confident = ailine.detect_header_row(sheet_struct)
    assert (row, confident) == (1, True)

def test_detect_header_row_title_row_then_header_at_row3():
    # A検体型: 行1=結合タイトル(str=1・閾値未満で候補外)、行2=単一文字列、行3=見出し(str=5)、
    # 行4以降=型混在データ。
    sheet_struct = {"rows": {
        1: {"nonempty": 1, "str": 1, "bold": 1},
        2: {"nonempty": 1, "str": 1, "bold": 0},
        3: {"nonempty": 5, "str": 5, "bold": 0},
        4: {"nonempty": 5, "str": 1, "bold": 0},
    }}
    row, confident = ailine.detect_header_row(sheet_struct)
    assert (row, confident) == (3, True)

def test_detect_header_row_two_level_header_picks_child_row():
    # D検体型: 行1=親見出し(結合・str=3)、行2=子見出し(str=4・型混在なし直下）、
    # 行3=型混在データ → 子見出し行(2)を採用する。
    sheet_struct = {"rows": {
        1: {"nonempty": 3, "str": 3, "bold": 0},
        2: {"nonempty": 4, "str": 4, "bold": 0},
        3: {"nonempty": 5, "str": 1, "bold": 0},
    }}
    row, confident = ailine.detect_header_row(sheet_struct)
    assert (row, confident) == (2, True)

def test_detect_header_row_ambiguous_two_equally_valid_candidates_is_not_confident():
    # 型混在の直下が2つとも存在する（曖昧）場合は推測しない。
    sheet_struct = {"rows": {
        1: {"nonempty": 2, "str": 2, "bold": 0},
        2: {"nonempty": 2, "str": 1, "bold": 0},
        3: {"nonempty": 2, "str": 2, "bold": 0},
        4: {"nonempty": 2, "str": 1, "bold": 0},
    }}
    row, confident = ailine.detect_header_row(sheet_struct)
    assert confident is False
    assert row is None

def test_detect_header_row_empty_sheet_is_not_confident():
    assert ailine.detect_header_row({"rows": {}}) == (None, False)

def test_resolve_header_rows_confident_detection_no_clarify():
    struct_dump = {"sheets": {"Sheet": {"rows": {
        1: {"nonempty": 2, "str": 2, "bold": 0},
        2: {"nonempty": 2, "str": 1, "bold": 0},
    }}}}
    header_rows, clarify = ailine.resolve_header_rows(struct_dump, ["Sheet"])
    assert header_rows == {"Sheet": 1}
    assert clarify is None

def test_resolve_header_rows_ambiguous_asks_clarify_with_exact_wording():
    struct_dump = {"sheets": {"Sheet": {"rows": {
        1: {"nonempty": 2, "str": 2, "bold": 0},
        2: {"nonempty": 2, "str": 1, "bold": 0},
        3: {"nonempty": 2, "str": 2, "bold": 0},
        4: {"nonempty": 2, "str": 1, "bold": 0},
    }}}}
    header_rows, clarify = ailine.resolve_header_rows(struct_dump, ["Sheet"])
    # ★ W8a 項目3: 旧文言は答え方が無い行き止まりだった。--header-row の使い方まで添える。
    assert clarify == ("見出しが何行目か分かりません。"
                        "`--header-row 3` のように指定して再実行してください")
    assert header_rows == {"Sheet": 1}   # 既定のまま(呼び出し側は CLARIFY で止まるので使われない)

def test_resolve_header_rows_no_struct_dump_defaults_to_row1_no_clarify():
    header_rows, clarify = ailine.resolve_header_rows({}, ["Sheet", "単価表"])
    assert header_rows == {"Sheet": 1, "単価表": 1}
    assert clarify is None

def test_resolve_header_rows_empty_sheets_list_is_noop():
    assert ailine.resolve_header_rows({"sheets": {}}, []) == ({}, None)

def test_cmd_run_clarify_on_ambiguous_header_before_translation(tmp_path, monkeypatch, capsys):
    # ★ W3: 見出し推定が曖昧なら、翻訳(translate_task)を呼ぶ前に CLARIFY して exit 3 になる
    #   （『三層全部が同じ見出し推定を使う』の前提＝そもそも翻訳に渡す接地が無い）。
    book = _book(tmp_path, [["a", 1], ["b", 2]])
    monkeypatch.setattr(ailine, "HISTORY_FILE", tmp_path / "history.jsonl")
    monkeypatch.setattr(ailine, "normalize_book", lambda book, workdir, timeout=None: book)
    ambiguous = {"sheets": {"Sheet": {"rows": {
        1: {"nonempty": 2, "str": 2, "bold": 0},
        2: {"nonempty": 2, "str": 1, "bold": 0},
        3: {"nonempty": 2, "str": 2, "bold": 0},
        4: {"nonempty": 2, "str": 1, "bold": 0},
    }}}}
    monkeypatch.setattr(ailine, "build_struct_dump", lambda book, workdir: ambiguous)
    called = {"n": 0}
    def boom(*a, **k):
        called["n"] += 1
        return {"plan": [{"op": "FREEFORM", "args": {}}]}
    monkeypatch.setattr(ailine, "translate_task", boom)
    ns = argparse.Namespace(
        book=str(book), task="いい感じにして", model="qwen2.5-coder:7b",
        refs=None, helpers=None, repair=0, temperature=0.2,
        dry=False, inplace=False, json=False, timeout=180.0, ask=False)
    rc = ailine.cmd_run(ns)
    captured = capsys.readouterr()
    assert rc == 3
    # ★ W8a 項目3: CLARIFY 文言更新（--header-row の使い方まで添える）に追従。
    assert "見出しが何行目か分かりません" in captured.out
    assert "--header-row" in captured.out
    assert called["n"] == 0   # 翻訳は一度も呼ばれていない

def test_cmd_run_header_row_flag_bypasses_ambiguous_detection_no_clarify(tmp_path, monkeypatch, capsys):
    # ★ W8a 項目3: 見出し検出が曖昧(CLARIFY相当)な帳票でも、--header-row を指定すれば
    #   検出を丸ごとスキップしてその行を採用し、CLARIFY の行き止まりに落ちない。
    #   接地(book_meta)経由で codegen/事後条件にも同じ header_row が伝わる（三層貫通）。
    book = _book(tmp_path, [["x", 1], ["y", 2], ["商品", "金額"], ["a", 300], ["b", 100]])
    monkeypatch.setattr(ailine, "HISTORY_FILE", tmp_path / "history.jsonl")
    monkeypatch.setattr(ailine, "normalize_book", lambda book, workdir, timeout=None: book)
    ambiguous = {"sheets": {"Sheet": {"rows": {
        1: {"nonempty": 2, "str": 2, "bold": 0},
        2: {"nonempty": 2, "str": 1, "bold": 0},
        3: {"nonempty": 2, "str": 2, "bold": 0},
        4: {"nonempty": 2, "str": 1, "bold": 0},
    }}}}
    monkeypatch.setattr(ailine, "build_struct_dump", lambda book, workdir: ambiguous)
    monkeypatch.setattr(ailine, "translate_task",
                        lambda model, task, book_meta, temperature=0.1:
                        {"op": "SORT", "args": {"col": "金額", "order": "desc"}})
    monkeypatch.setattr(ailine, "basrun_apply",
                        lambda out_book, code, workdir, helper_files=(), timeout=None:
                        (True, None, "ok"))
    ns = argparse.Namespace(
        book=str(book), task="金額で降順に並べ替えて", model="qwen2.5-coder:7b",
        refs=None, helpers=None, repair=0, temperature=0.2,
        dry=False, inplace=False, json=False, timeout=180.0, ask=False, values=False,
        header_row=3)
    ailine.cmd_run(ns)
    captured = capsys.readouterr()
    assert "？" not in captured.out    # CLARIFY に落ちていない
    assert "対象:金額" in captured.out   # 3行目を見出しとして『金額』列を解決できている

def test_cmd_run_without_header_row_flag_still_clarifies_on_ambiguous(tmp_path, monkeypatch, capsys):
    # 対照: --header-row を指定しなければ従来どおり曖昧な場合は CLARIFY のまま。
    book = _book(tmp_path, [["x", 1], ["y", 2], ["商品", "金額"], ["a", 300], ["b", 100]])
    monkeypatch.setattr(ailine, "HISTORY_FILE", tmp_path / "history.jsonl")
    monkeypatch.setattr(ailine, "normalize_book", lambda book, workdir, timeout=None: book)
    ambiguous = {"sheets": {"Sheet": {"rows": {
        1: {"nonempty": 2, "str": 2, "bold": 0},
        2: {"nonempty": 2, "str": 1, "bold": 0},
        3: {"nonempty": 2, "str": 2, "bold": 0},
        4: {"nonempty": 2, "str": 1, "bold": 0},
    }}}}
    monkeypatch.setattr(ailine, "build_struct_dump", lambda book, workdir: ambiguous)
    called = {"n": 0}
    def boom(*a, **k):
        called["n"] += 1
        return {"op": "SORT", "args": {"col": "金額", "order": "desc"}}
    monkeypatch.setattr(ailine, "translate_task", boom)
    ns = argparse.Namespace(
        book=str(book), task="金額で降順に並べ替えて", model="qwen2.5-coder:7b",
        refs=None, helpers=None, repair=0, temperature=0.2,
        dry=False, inplace=False, json=False, timeout=180.0, ask=False, values=False,
        header_row=None)
    rc = ailine.cmd_run(ns)
    captured = capsys.readouterr()
    assert rc == 3
    assert "--header-row" in captured.out
    assert called["n"] == 0

def test_cmd_run_dry_skips_structdump_and_uses_physical_row1(tmp_path, monkeypatch, capsys):
    # ★ --dry は LibreOffice に触れない（既存の設計不変条件）。normalize_book が
    #   呼ばれていないことを確認する。
    book = _book(tmp_path, [["商品", "金額"]])
    monkeypatch.setattr(ailine, "HISTORY_FILE", tmp_path / "history.jsonl")
    called = {"n": 0}
    def boom(*a, **k):
        called["n"] += 1
        return a[0]
    monkeypatch.setattr(ailine, "normalize_book", boom)
    monkeypatch.setattr(ailine, "translate_task",
                        lambda model, task, book_meta, temperature=0.1:
                        {"op": "SORT", "args": {"col": "金額", "order": "desc"}})
    ns = argparse.Namespace(
        book=str(book), task="金額で降順に並べ替えて", model="qwen2.5-coder:7b",
        refs=None, helpers=None, repair=0, temperature=0.2,
        dry=True, inplace=False, json=False, timeout=180.0, ask=False)
    rc = ailine.cmd_run(ns)
    assert rc == 0
    assert called["n"] == 0


# --- codegen: header_row(hr0) が三層で一貫して使われる（接地→codegen） -------------

def test_codegen_dsl_sort_uses_detected_header_row_from_book_meta():
    meta = {"sheets": ["Sheet"], "headers": {"Sheet": ["商品", "金額"]},
            "header_rows": {"Sheet": 3}}
    code = ailine.codegen_dsl("SORT", {"col": "金額", "order": "desc"}, meta)
    assert "Call SortByColumn(oDoc, 2, 1, 1, False)" in code   # hr0 = 3-1 = 2, lastCol=1(列2つ)

def test_codegen_dsl_compute_column_formula_uses_detected_header_row():
    meta = {"sheets": ["Sheet"], "headers": {"Sheet": ["商品", "売上", "原価"]},
            "header_rows": {"Sheet": 3}}
    code = ailine.codegen_dsl("COMPUTE_COLUMN", {"operands": ["売上", "原価"], "operator": "-"}, meta)
    assert "For i = 3 To lastRow" in code   # hr0+1 = 2+1 = 3
    assert 'setString("売上-原価")' in code
    assert 'getCellByPosition(3, 2).setString' in code   # 新規列は列3, 見出し行は hr0=2


# --- 事後条件: header_row が接地・codegen と同じ行を使う ---------------------------

def test_check_sort_with_header_row_three_matches_a_specimen_layout(tmp_path):
    p = tmp_path / "a_like.xlsx"
    wb = openpyxl.Workbook(); ws = wb.active
    ws["A1"] = "タイトル"
    ws["A2"] = "作成日"
    for c, h in enumerate(["商品", "金額"], start=1):
        ws.cell(row=3, column=c, value=h)
    ws.append(["a", 300]); ws.append(["b", 200]); ws.append(["c", 100])
    # ↑ append は末尾行に追記するため、3行目の直後(4行目)から入る
    wb.save(p)
    status, reason = ailine.check_sort(p, {"col": "金額", "order": "desc"}, header_row=3)
    assert status == "pass"

def test_check_sort_with_wrong_header_row_fails_to_find_column(tmp_path):
    p = tmp_path / "a_like.xlsx"
    wb = openpyxl.Workbook(); ws = wb.active
    ws["A1"] = "タイトル"
    ws["A2"] = "作成日"
    for c, h in enumerate(["商品", "金額"], start=1):
        ws.cell(row=3, column=c, value=h)
    ws.append(["a", 300])
    wb.save(p)
    status, reason = ailine.check_sort(p, {"col": "金額", "order": "desc"}, header_row=1)
    assert status == "fail"
    assert "がありません" not in reason or "見つからない" in reason


def _inject_formula_cache(path, sheet_filename: str, addr_to_value: dict) -> None:
    """テスト専用: xlsx の数式セルへキャッシュ値(<v>)を直接注入する（openpyxl は数式を
       計算しないため、LO を使わずに二層事後条件(式+キャッシュ値)を検証するための小道具）。"""
    import re
    import zipfile
    tmp = path.with_suffix(".tmp.xlsx")
    with zipfile.ZipFile(path, "r") as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == sheet_filename:
                text = data.decode("utf-8")
                for addr, value in addr_to_value.items():
                    # openpyxl は数式セルに空の <v></v> を既に書いていることがある（無い場合もある）。
                    # どちらでも対応できるよう <v>...</v> の有無を任意にして丸ごと置き換える。
                    pattern = re.compile(rf'(<c r="{addr}"[^>]*>.*?<f>.*?</f>)(?:<v>.*?</v>)?(</c>)')
                    text = pattern.sub(rf'\1<v>{value}</v>\2', text, count=1)
                data = text.encode("utf-8")
            zout.writestr(item, data)
    tmp.replace(path)


def _formula_book(tmp_path, header_row_values, data_rows, formula_col_letter, operator_col_letters,
                   operator, name="f.xlsx"):
    """演算対象2列+式列を持つブックを作る。式列には setFormula 相当の文字列を直接書く
       （キャッシュ値は別途 _inject_formula_cache で注ぐ）。"""
    p = tmp_path / name
    wb = openpyxl.Workbook(); ws = wb.active
    ws.append(header_row_values)
    for i, row in enumerate(data_rows, start=2):
        ws[f"A{i}"] = f"item{i}"   # ★ _scan_last_row は A列(key_col=1既定)で行の有無を判定する
        for col_letter, val in row.items():
            ws[f"{col_letter}{i}"] = val
        c1, c2 = operator_col_letters
        ws[f"{formula_col_letter}{i}"] = f"={c1}{i}{operator}{c2}{i}"
    wb.save(p)
    return p

def test_check_compute_column_formula_mode_passes_when_formula_and_cache_both_match(tmp_path):
    p = _formula_book(tmp_path, ["商品", "数量", "単価", "小計"],
                       [{"B": 3, "C": 120}, {"B": 5, "C": 80}], "D", ("B", "C"), "*")
    _inject_formula_cache(p, "xl/worksheets/sheet1.xml", {"D2": 360, "D3": 400})
    status, reason = ailine.check_compute_column(
        p, {"operands": ["数量", "単価"], "operator": "*", "target": "小計"}, use_formula=True)
    assert status == "pass"
    assert "式・キャッシュ値とも一致" in reason

def test_check_compute_column_formula_mode_fails_when_formula_string_wrong(tmp_path):
    p = tmp_path / "f.xlsx"
    wb = openpyxl.Workbook(); ws = wb.active
    ws.append(["商品", "数量", "単価", "小計"])
    ws.append(["りんご", 3, 120, None])
    ws["D2"] = "=B2+C2"   # ★ 期待は * なのに + を書いてしまった想定
    wb.save(p)
    _inject_formula_cache(p, "xl/worksheets/sheet1.xml", {"D2": 123})
    status, reason = ailine.check_compute_column(
        p, {"operands": ["数量", "単価"], "operator": "*", "target": "小計"}, use_formula=True)
    assert status == "fail"
    assert "式が期待形でない" in reason

def test_check_compute_column_formula_mode_fails_when_cache_value_wrong(tmp_path):
    p = _formula_book(tmp_path, ["商品", "数量", "単価", "小計"],
                       [{"B": 3, "C": 120}], "D", ("B", "C"), "*")
    _inject_formula_cache(p, "xl/worksheets/sheet1.xml", {"D2": 999})   # ★ 期待360なのに999
    status, reason = ailine.check_compute_column(
        p, {"operands": ["数量", "単価"], "operator": "*", "target": "小計"}, use_formula=True)
    assert status == "fail"
    assert "キャッシュ値が不一致" in reason

def test_check_compute_column_values_mode_unaffected_by_formula_flag_default(tmp_path):
    # use_formula 省略時は既定 False（旧テスト・旧挙動と同一）。
    p = _book(tmp_path, [["売上", "原価", "売上-原価"], [500, 300, 200]])
    status, reason = ailine.check_compute_column(p, {"operands": ["売上", "原価"], "operator": "-"})
    assert status == "pass"
    assert "キャッシュ値" not in reason


# --- codegen: COMPUTE_COLUMN の式化は LO 方言(;/.) を要さない（formula_spike の実測どおり） --

def test_codegen_dsl_compute_column_formula_has_no_semicolon_or_sheet_dot():
    # ★ bench/formula_spike_RESULTS.md: setFormula は多引数(;)・シート参照(.)の時だけ
    #   LO 方言が要る。COMPUTE_COLUMN は単純な行内二項演算(=B2*C2 型)でどちらも使わない。
    meta = {"sheets": ["Sheet"], "headers": {"Sheet": ["数量", "単価", "小計"]}}
    code = ailine.codegen_dsl("COMPUTE_COLUMN",
                               {"operands": ["数量", "単価"], "operator": "*", "target": "小計"}, meta)
    assert ";" not in code
    assert '"A" & (i + 1) & "*" & "B" & (i + 1)' in code   # 数量=A列(0+1) 単価=B列(1+1)


# --- _scan_last_row_basic: header_row 対応の走査開始/ガード閾値 -------------------

def test_scan_last_row_basic_default_matches_legacy_output():
    assert ailine._scan_last_row_basic() == (
        "    lastRow = 1\n"
        "    Do While oSheet.getCellByPosition(0, lastRow).getString() <> \"\"\n"
        "        lastRow = lastRow + 1\n"
        "    Loop\n"
        "    lastRow = lastRow - 1\n"
        "    If lastRow < 1 Then Exit Sub\n")

def test_scan_last_row_basic_custom_start_row_and_min_ok():
    out = ailine._scan_last_row_basic(start_row="3", min_ok="2")
    assert "lastRow = 3\n" in out
    assert "If lastRow < 2 Then Exit Sub\n" in out


# ===========================================================================
# W6: APPEND_TOTAL 語彙昇格（監査3回連続失敗の実測による昇格経済学）
# ===========================================================================

def test_ops_doc_and_fewshot_mention_append_total():
    assert "APPEND_TOTAL" in ailine.OPS_DOC
    assert any("APPEND_TOTAL" in assistant_ex for _u, assistant_ex in ailine.TRANSLATION_FEWSHOT)

def test_translate_task_parses_append_total_with_label_and_factor(monkeypatch):
    # ★ A': プロンプトはもう factor を求めないが、translate_task 自体は仮に LLM が
    #   余計な factor を返しても素通しする（黙って落とさない・verify_dsl_args 側で
    #   factor は無視され機械抽出に置き換わる。ここは翻訳層のJSON解析の頑健性テスト）。
    monkeypatch.setattr(
        ailine, "ollama_generate_json",
        lambda model, msgs, temperature=0.1, num_predict=300:
        '{"plan": [{"op": "APPEND_TOTAL", "args": '
        '{"col": "小計", "label": "税込み合計", "factor": 1.1}}]}')
    meta = {"sheets": ["Sheet"], "headers": {"Sheet": ["品目", "数量", "単価", "小計"]}}
    result = ailine.translate_task("m", "税込み合計を一番下に出して（消費税10%）", meta)
    step = result["plan"][0]
    assert step["op"] == "APPEND_TOTAL"
    assert step["args"] == {"col": "小計", "label": "税込み合計", "factor": 1.1}

def test_ops_doc_does_not_ask_llm_for_factor():
    # ★ A': factor は machine-determined。プロンプトが LLM に数値化を求めないことを固定する。
    assert "factor" not in ailine.OPS_DOC
    for _u, assistant_ex in ailine.TRANSLATION_FEWSHOT:
        if "APPEND_TOTAL" in assistant_ex:
            assert '"factor"' not in assistant_ex

def test_translate_task_parses_append_total_without_optional_args(monkeypatch):
    monkeypatch.setattr(
        ailine, "ollama_generate_json",
        lambda model, msgs, temperature=0.1, num_predict=300:
        '{"plan": [{"op": "APPEND_TOTAL", "args": {"col": "金額"}}]}')
    meta = {"sheets": ["Sheet"], "headers": {"Sheet": ["部門", "金額"]}}
    result = ailine.translate_task("m", "金額の合計を最後に", meta)
    step = result["plan"][0]
    assert step["op"] == "APPEND_TOTAL"
    assert step["args"] == {"col": "金額"}   # label/factor の既定は verify_dsl_args 側が確定する


# --- ② 検証（接地） ----------------------------------------------------------

def test_verify_dsl_args_append_total_defaults_label_and_factor():
    meta = {"sheets": ["Sheet"], "headers": {"Sheet": ["部門", "金額"]}}
    ok, resolved, inferred, err = ailine.verify_dsl_args("APPEND_TOTAL", {"col": "金額"}, meta)
    assert ok
    assert resolved["col"] == "金額"
    assert resolved["label"] == "合計"
    assert resolved["factor"] == 1.0

def test_verify_dsl_args_append_total_resolves_factor_from_task_text_percent():
    # ★ A': factor は LLM でなく、依頼文の明示率を machine が抽出する。
    meta = {"sheets": ["Sheet"], "headers": {"Sheet": ["品目", "数量", "単価", "小計"]}}
    ok, resolved, inferred, err = ailine.verify_dsl_args(
        "APPEND_TOTAL", {"col": "小計", "label": "税込み合計"}, meta,
        task="税込み合計を一番下に出して（消費税10%）")
    assert ok
    assert resolved["label"] == "税込み合計"
    assert resolved["factor"] == 1.1
    assert resolved["_sources"]["factor"] == "依頼文: 10%"

def test_verify_dsl_args_append_total_resolves_factor_from_vocab():
    # 用語集の語が依頼文に部分一致で含まれる場合だけ引く（無関係な語を勝手に当てない）。
    meta = {"sheets": ["Sheet"], "headers": {"Sheet": ["品目", "数量", "単価", "小計"]}}
    ok, resolved, inferred, err = ailine.verify_dsl_args(
        "APPEND_TOTAL", {"col": "小計", "label": "税込み合計"}, meta,
        task="消費税込みの合計を出して", vocab={"消費税": 1.1})
    assert ok
    assert resolved["factor"] == 1.1
    assert resolved["_sources"]["factor"] == "用語集: 消費税"

def test_verify_dsl_args_append_total_task_text_wins_over_vocab():
    # 依頼文に明示率があれば、用語集より依頼文を優先する。
    meta = {"sheets": ["Sheet"], "headers": {"Sheet": ["品目", "数量", "単価", "小計"]}}
    ok, resolved, inferred, err = ailine.verify_dsl_args(
        "APPEND_TOTAL", {"col": "小計", "label": "税込み合計"}, meta,
        task="税込み合計を出して（消費税8%で）", vocab={"消費税": 1.1})
    assert ok
    assert resolved["factor"] == 1.08
    assert "依頼文" in resolved["_sources"]["factor"]

def test_verify_dsl_args_append_total_llm_factor_ignored_but_warns_on_mismatch():
    # LLM が(旧仕様の名残や幻覚で)factor を返しても機械抽出が常に勝つ。食い違いは WARN。
    meta = {"sheets": ["Sheet"], "headers": {"Sheet": ["品目", "数量", "単価", "小計"]}}
    ok, resolved, inferred, err = ailine.verify_dsl_args(
        "APPEND_TOTAL", {"col": "小計", "label": "税込み合計", "factor": 1.08}, meta,
        task="税込み合計を一番下に出して（消費税10%）")
    assert ok
    assert resolved["factor"] == 1.1
    assert resolved["_warnings"]
    assert "1.08" in resolved["_warnings"][0] and "1.1" in resolved["_warnings"][0]

def test_verify_dsl_args_append_total_non_tax_label_defaults_without_clarify():
    # label が税/込を含まなければ、倍率がどこにも無くても既定1.0で通す
    # （恒真式の番人は「税/込ラベルなのに倍率不明」の場合だけに絞る）。
    meta = {"sheets": ["Sheet"], "headers": {"Sheet": ["部門", "金額"]}}
    ok, resolved, inferred, err = ailine.verify_dsl_args(
        "APPEND_TOTAL", {"col": "金額"}, meta, task="金額の合計を最後に")
    assert ok
    assert resolved["factor"] == 1.0
    assert "_sources" not in resolved

def test_verify_dsl_args_append_total_clarifies_when_label_implies_tax_but_no_rate():
    # ★ 恒真式の番人（最優先）: label が税/込を含むのに倍率がどこにも無いと CLARIFY へ倒す
    #   （税抜き金額に「税込み」ラベルが付く恒真の誤りを機械で止める）。
    meta = {"sheets": ["Sheet"], "headers": {"Sheet": ["品目", "数量", "単価", "小計"]}}
    ok, resolved, inferred, err = ailine.verify_dsl_args(
        "APPEND_TOTAL", {"col": "小計", "label": "税込み合計"}, meta, task="税込み合計を出して")
    assert not ok
    assert "倍率が分かりません" in err
    assert "ailine vocab add" in err

def test_verify_dsl_args_append_total_ignores_non_numeric_llm_factor():
    # ★ A': factor はもう「ユーザー入力のエラー」ではない。LLM が壊れた値を返しても
    #   単に無視して機械確定（既定1.0）にフォールバックする（float() できない値は
    #   食い違いWARNの対象にもしない＝比較不能なので黙って無視）。
    meta = {"sheets": ["Sheet"], "headers": {"Sheet": ["部門", "金額"]}}
    ok, resolved, inferred, err = ailine.verify_dsl_args(
        "APPEND_TOTAL", {"col": "金額", "factor": "abc"}, meta, task="金額の合計を最後に")
    assert ok
    assert resolved["factor"] == 1.0
    assert "_warnings" not in resolved

def test_verify_dsl_args_append_total_rejects_zero_factor_from_text_extraction():
    meta = {"sheets": ["Sheet"], "headers": {"Sheet": ["部門", "金額"]}}
    ok, resolved, inferred, err = ailine.verify_dsl_args(
        "APPEND_TOTAL", {"col": "金額"}, meta, task="0倍の合計を出して")
    assert not ok
    assert "正の数" in err

def test_verify_dsl_args_append_total_unknown_column_is_clarify_error():
    meta = {"sheets": ["Sheet"], "headers": {"Sheet": ["部門", "金額"]}}
    ok, resolved, inferred, err = ailine.verify_dsl_args(
        "APPEND_TOTAL", {"col": "存在しない"}, meta)
    assert not ok
    assert "がありません" in err

def test_format_confirmation_line_append_total_shows_col_label_factor():
    # ★ W8a 項目5: 表示ラベルのみ「倍率」→「率」（内部キー"factor"は不変）。
    line = ailine.format_confirmation_line(
        "APPEND_TOTAL", {"col": "小計", "label": "税込み合計", "factor": 1.1}, set())
    assert "対象列:小計" in line
    assert "ラベル:税込み合計" in line
    assert "率:1.1" in line

def test_format_confirmation_line_append_total_shows_factor_source():
    # ★ A': 来歴の可視化。率の出典を確認行に添える。★ W8a 項目5: 表示は「率」。
    line = ailine.format_confirmation_line(
        "APPEND_TOTAL",
        {"col": "小計", "label": "税込み合計", "factor": 1.1,
         "_sources": {"factor": "依頼文: 10%"}},
        set())
    assert "率:1.1（依頼文: 10%）" in line


# --- ④ codegen（決定論） ------------------------------------------------------

def test_codegen_dsl_append_total_places_label_left_of_value_and_formula_with_factor():
    # ★ B: 挿入耐性式（SUM(D2:INDEX(D:D;ROW()-1))型・LO方言のセミコロンで setFormula する。
    #   保存後はカンマ形に自動変換される・formula spike で実測済み）。
    meta = {"sheets": ["Sheet"], "headers": {"Sheet": ["品目", "数量", "単価", "小計"]}}
    code = ailine.codegen_dsl(
        "APPEND_TOTAL", {"col": "小計", "label": "税込み合計", "factor": 1.1}, meta)
    # 小計=列3(0起点)。ラベルはその左隣=列2に置く（既存構造を壊さない置き方）。
    assert 'getCellByPosition(2, totalRow).setString("税込み合計")' in code
    assert 'getCellByPosition(3, totalRow).setFormula(' in code
    assert ('"=SUM(" & "D" & 2 & ":INDEX(" & "D" & ":" & "D" & ";ROW()-1))" & "*1.1"') in code

def test_codegen_dsl_append_total_omits_factor_tail_when_factor_is_one():
    meta = {"sheets": ["Sheet"], "headers": {"Sheet": ["部門", "金額"]}}
    code = ailine.codegen_dsl("APPEND_TOTAL", {"col": "金額", "label": "合計", "factor": 1}, meta)
    assert ('"=SUM(" & "B" & 2 & ":INDEX(" & "B" & ":" & "B" & ";ROW()-1))" & ""') in code

def test_codegen_dsl_append_total_leftmost_column_skips_label():
    meta = {"sheets": ["Sheet"], "headers": {"Sheet": ["金額", "部門"]}}
    code = ailine.codegen_dsl("APPEND_TOTAL", {"col": "金額", "label": "合計", "factor": 1}, meta)
    assert "setString(" not in code   # 列0は左隣が無いためラベルを省略する


# --- ⑥ 事後条件（二層: 式文字列 + キャッシュ値、ラベル一致） -------------------

def test_check_append_total_passes_with_factor_and_label(tmp_path):
    # ★ B: 保存後のカンマ形(=SUM(D2:INDEX(D:D,ROW()-1))*1.1)と照合する（codegen が
    #   LO 方言(;)で書いても、basrun 保存後はこの形に変換される・formula spike で実測済み）。
    p = tmp_path / "inv.xlsx"
    wb = openpyxl.Workbook(); ws = wb.active
    ws.append(["品目", "数量", "単価", "小計"])
    ws.append(["a", 3, 50000, 150000])
    ws.append(["b", 1, 120000, 120000])
    ws.append(["c", 12, 8000, 96000])
    ws["C5"] = "税込み合計"
    ws["D5"] = "=SUM(D2:INDEX(D:D,ROW()-1))*1.1"
    wb.save(p)
    _inject_formula_cache(p, "xl/worksheets/sheet1.xml", {"D5": 402600})
    status, reason = ailine.check_append_total(
        p, {"col": "小計", "label": "税込み合計", "factor": 1.1})
    assert status == "pass"
    assert "式・キャッシュ値・ラベルとも一致" in reason

def test_check_append_total_fails_when_formula_not_sum(tmp_path):
    p = tmp_path / "inv.xlsx"
    wb = openpyxl.Workbook(); ws = wb.active
    ws.append(["品目", "数量", "単価", "小計"])
    ws.append(["a", 3, 50000, 150000])
    ws["D3"] = 999999   # 合計行のはずが数式でなく値
    wb.save(p)
    status, reason = ailine.check_append_total(p, {"col": "小計", "label": "合計", "factor": 1})
    assert status == "fail"
    assert "式が期待形" in reason

def test_check_append_total_fails_when_cache_value_wrong(tmp_path):
    p = tmp_path / "inv.xlsx"
    wb = openpyxl.Workbook(); ws = wb.active
    ws.append(["部門", "金額"])
    ws.append(["a", 100])
    ws.append(["b", 200])
    ws["A4"] = "合計"
    ws["B4"] = "=SUM(B2:INDEX(B:B,ROW()-1))"
    wb.save(p)
    _inject_formula_cache(p, "xl/worksheets/sheet1.xml", {"B4": 999})
    status, reason = ailine.check_append_total(p, {"col": "金額", "label": "合計", "factor": 1})
    assert status == "fail"
    assert "キャッシュ値が不一致" in reason

def test_check_append_total_fails_when_label_wrong(tmp_path):
    p = tmp_path / "inv.xlsx"
    wb = openpyxl.Workbook(); ws = wb.active
    ws.append(["部門", "金額"])
    ws.append(["a", 100])
    ws.append(["b", 200])
    ws["A4"] = "TOTAL"   # 期待ラベルは「合計」
    ws["B4"] = "=SUM(B2:INDEX(B:B,ROW()-1))"
    wb.save(p)
    _inject_formula_cache(p, "xl/worksheets/sheet1.xml", {"B4": 300})
    status, reason = ailine.check_append_total(p, {"col": "金額", "label": "合計", "factor": 1})
    assert status == "fail"
    assert "ラベルが期待" in reason

def test_check_append_total_fails_when_zero_data_rows(tmp_path):
    p = tmp_path / "inv.xlsx"
    wb = openpyxl.Workbook(); ws = wb.active
    ws.append(["部門", "金額"])
    wb.save(p)
    status, reason = ailine.check_append_total(p, {"col": "金額", "label": "合計", "factor": 1})
    assert status == "fail"
    assert reason == ailine._ZERO_TARGET_REASON

def test_check_append_total_leftmost_column_skips_label_check(tmp_path):
    p = tmp_path / "inv.xlsx"
    wb = openpyxl.Workbook(); ws = wb.active
    ws.append(["金額", "備考"])
    ws.append([100, "x"])
    ws.append([200, "y"])
    ws["A4"] = "=SUM(A2:INDEX(A:A,ROW()-1))"
    wb.save(p)
    _inject_formula_cache(p, "xl/worksheets/sheet1.xml", {"A4": 300})
    status, reason = ailine.check_append_total(p, {"col": "金額", "label": "合計", "factor": 1})
    assert status == "pass"

def test_check_append_total_fails_on_old_static_range_formula(tmp_path):
    # ★ B: 静的な "=SUM(B2:B3)" 型（挿入耐性が無い旧形）はもう合格させない。
    p = tmp_path / "inv.xlsx"
    wb = openpyxl.Workbook(); ws = wb.active
    ws.append(["部門", "金額"])
    ws.append(["a", 100])
    ws.append(["b", 200])
    ws["A4"] = "合計"
    ws["B4"] = "=SUM(B2:B3)"
    wb.save(p)
    _inject_formula_cache(p, "xl/worksheets/sheet1.xml", {"B4": 300})
    status, reason = ailine.check_append_total(p, {"col": "金額", "label": "合計", "factor": 1})
    assert status == "fail"
    assert "期待形" in reason

def test_run_postcondition_dispatches_append_total(tmp_path):
    p = tmp_path / "inv.xlsx"
    wb = openpyxl.Workbook(); ws = wb.active
    ws.append(["部門", "金額"])
    ws.append(["a", 100])
    ws["A3"] = "合計"
    ws["B3"] = "=SUM(B2:INDEX(B:B,ROW()-1))"
    wb.save(p)
    _inject_formula_cache(p, "xl/worksheets/sheet1.xml", {"B3": 100})
    status, reason = ailine.run_postcondition(
        "APPEND_TOTAL", p, {"col": "金額", "label": "合計", "factor": 1})
    assert status == "pass"


# ===========================================================================
# ★ A': factor の機械抽出（依頼文の明示率・regex のみ・LLM不使用）
# ===========================================================================

def test_extract_rate_factor_percent():
    assert ailine.extract_rate_factor("税込み合計を出して（消費税10%）") == (1.1, "10%")

def test_extract_rate_factor_percent_with_space_and_zenkaku():
    assert ailine.extract_rate_factor("8 ％引きで") == (1.08, "8 ％")

def test_extract_rate_factor_bai_suffix():
    assert ailine.extract_rate_factor("1.1倍にして合計を出して") == (1.1, "1.1倍")

def test_extract_rate_factor_bare_decimal_near_tax_keyword():
    # 税/倍率の語の近傍だけにある裸の小数（0-1未満は 1+n として解釈）
    assert ailine.extract_rate_factor("税率が0.1です") == (1.1, "0.1")

def test_extract_rate_factor_bare_decimal_far_from_keyword_is_ignored():
    # 無関係な数値の誤爆を避ける（税/倍率の語から離れた裸の小数は拾わない）
    factor, snippet = ailine.extract_rate_factor("0.1個くらい余分に買って集計して")
    assert factor is None and snippet is None

def test_extract_rate_factor_no_mention_returns_none():
    assert ailine.extract_rate_factor("金額の合計を最後に") == (None, None)

def test_extract_rate_factor_conflicting_values_returns_none():
    # 複数の異なる値が出たら断定しない（CLARIFY に委ねる）
    factor, snippet = ailine.extract_rate_factor("消費税10%か8%のどちらかで")
    assert factor is None and snippet is None

def test_extract_rate_factor_same_value_repeated_is_not_conflicting():
    factor, _ = ailine.extract_rate_factor("消費税10%(10%)込みで")
    assert factor == 1.1

def test_extract_rate_factor_empty_text():
    assert ailine.extract_rate_factor("") == (None, None)


# ===========================================================================
# ★ A': 用語集（vocab）
# ===========================================================================

def test_load_vocab_missing_file_returns_empty_dict(tmp_path):
    assert ailine.load_vocab(tmp_path / "nope.json") == {}

def test_save_and_load_vocab_roundtrip(tmp_path):
    p = tmp_path / "vocab.json"
    ailine.save_vocab({"消費税": 1.1, "軽減税率": 1.08}, path=p)
    assert ailine.load_vocab(p) == {"消費税": 1.1, "軽減税率": 1.08}

def test_load_vocab_corrupt_json_returns_empty_dict_not_crash(tmp_path):
    p = tmp_path / "vocab.json"
    p.write_text("{not valid json", encoding="utf-8")
    assert ailine.load_vocab(p) == {}

def test_load_vocab_non_dict_json_returns_empty_dict(tmp_path):
    p = tmp_path / "vocab.json"
    p.write_text("[1, 2, 3]", encoding="utf-8")
    assert ailine.load_vocab(p) == {}

def test_load_vocab_skips_non_numeric_values(tmp_path):
    p = tmp_path / "vocab.json"
    p.write_text('{"消費税": 1.1, "壊れた語": "abc", "null値": null}', encoding="utf-8")
    assert ailine.load_vocab(p) == {"消費税": 1.1}

def test_load_vocab_skips_control_character_terms(tmp_path):
    # ★ codegen へ渡る経路の防御。改行を含む語は読み捨てる。
    p = tmp_path / "vocab.json"
    p.write_text(json.dumps({"改行\n入り": 1.1, "消費税": 1.1}, ensure_ascii=False),
                encoding="utf-8")
    assert ailine.load_vocab(p) == {"消費税": 1.1}

def test_load_vocab_skips_overlong_terms(tmp_path):
    p = tmp_path / "vocab.json"
    long_term = "あ" * (ailine.DEFAULT_VOCAB_MAX_TERM_LEN + 1)
    p.write_text(json.dumps({long_term: 1.1, "消費税": 1.1}, ensure_ascii=False),
                encoding="utf-8")
    assert ailine.load_vocab(p) == {"消費税": 1.1}

def test_load_vocab_caps_entry_count(tmp_path):
    p = tmp_path / "vocab.json"
    many = {f"語{i}": 1.0 + i / 1000 for i in range(ailine.DEFAULT_VOCAB_MAX_ENTRIES + 20)}
    p.write_text(json.dumps(many, ensure_ascii=False), encoding="utf-8")
    assert len(ailine.load_vocab(p)) == ailine.DEFAULT_VOCAB_MAX_ENTRIES

def test_vocab_add_new_term(tmp_path):
    p = tmp_path / "vocab.json"
    ok, msg = ailine.vocab_add("消費税", "1.1", path=p)
    assert ok
    assert ailine.load_vocab(p) == {"消費税": 1.1}

def test_vocab_add_updates_existing_term(tmp_path):
    p = tmp_path / "vocab.json"
    ailine.vocab_add("消費税", 1.08, path=p)
    ailine.vocab_add("消費税", 1.1, path=p)
    assert ailine.load_vocab(p) == {"消費税": 1.1}

def test_vocab_add_rejects_non_numeric_value(tmp_path):
    p = tmp_path / "vocab.json"
    ok, msg = ailine.vocab_add("消費税", "abc", path=p)
    assert not ok
    assert "数値ではありません" in msg
    assert ailine.load_vocab(p) == {}

def test_vocab_add_rejects_control_character_term(tmp_path):
    p = tmp_path / "vocab.json"
    ok, msg = ailine.vocab_add("改行\n入り", 1.1, path=p)
    assert not ok

def test_vocab_add_rejects_empty_term(tmp_path):
    p = tmp_path / "vocab.json"
    ok, msg = ailine.vocab_add("   ", 1.1, path=p)
    assert not ok

def test_vocab_add_rejects_when_entry_cap_reached_for_new_term(tmp_path):
    p = tmp_path / "vocab.json"
    many = {f"語{i}": 1.0 for i in range(ailine.DEFAULT_VOCAB_MAX_ENTRIES)}
    ailine.save_vocab(many, path=p)
    ok, msg = ailine.vocab_add("新語", 1.1, path=p)
    assert not ok
    assert "上限" in msg
    # 既存語の更新は上限に関係なく可能
    ok2, _ = ailine.vocab_add("語0", 2.0, path=p)
    assert ok2

def test_lookup_vocab_factor_matches_substring():
    assert ailine.lookup_vocab_factor("消費税込みの合計を出して", {"消費税": 1.1}) == (1.1, "消費税")

def test_lookup_vocab_factor_no_match_returns_none():
    assert ailine.lookup_vocab_factor("金額の合計を出して", {"消費税": 1.1}) == (None, None)

def test_lookup_vocab_factor_conflicting_terms_returns_none():
    factor, term = ailine.lookup_vocab_factor(
        "消費税と軽減税率どちらか", {"消費税": 1.1, "軽減税率": 1.08})
    assert factor is None and term is None

def test_lookup_vocab_factor_empty_vocab_returns_none():
    assert ailine.lookup_vocab_factor("消費税込み", {}) == (None, None)


# ===========================================================================
# ★ A': vocab CLI (`ailine vocab add` / `ailine vocab list`)
# ===========================================================================

def test_cmd_vocab_add_success(tmp_path, monkeypatch, capsys):
    monkeypatch.setattr(ailine, "VOCAB_FILE", tmp_path / "vocab.json")
    ns = argparse.Namespace(vocab_cmd="add", term="消費税", value="1.1")
    rc = ailine.cmd_vocab(ns)
    captured = capsys.readouterr()
    assert rc == 0
    assert "登録" in captured.out
    assert ailine.load_vocab(tmp_path / "vocab.json") == {"消費税": 1.1}

def test_cmd_vocab_add_failure_returns_1(tmp_path, monkeypatch, capsys):
    monkeypatch.setattr(ailine, "VOCAB_FILE", tmp_path / "vocab.json")
    ns = argparse.Namespace(vocab_cmd="add", term="消費税", value="abc")
    rc = ailine.cmd_vocab(ns)
    captured = capsys.readouterr()
    assert rc == 1
    assert "×" in captured.out

def test_cmd_vocab_list_empty(tmp_path, monkeypatch, capsys):
    monkeypatch.setattr(ailine, "VOCAB_FILE", tmp_path / "vocab.json")
    rc = ailine.cmd_vocab(argparse.Namespace(vocab_cmd="list"))
    captured = capsys.readouterr()
    assert rc == 0
    assert "空" in captured.out

def test_cmd_vocab_list_shows_entries(tmp_path, monkeypatch, capsys):
    p = tmp_path / "vocab.json"
    monkeypatch.setattr(ailine, "VOCAB_FILE", p)
    ailine.save_vocab({"消費税": 1.1}, path=p)
    rc = ailine.cmd_vocab(argparse.Namespace(vocab_cmd="list"))
    captured = capsys.readouterr()
    assert rc == 0
    assert "消費税 = 1.1" in captured.out


# ===========================================================================
# ★ A': cmd_run_dsl 統合（APPEND_TOTAL の倍率解決を通しで確認）
# ===========================================================================

def test_cmd_run_dsl_append_total_shows_factor_source_from_task_text(tmp_path, monkeypatch, capsys):
    book = _book(tmp_path, [["品目", "数量", "単価", "小計"], ["a", 3, 50000, 150000]])
    monkeypatch.setattr(ailine, "HISTORY_FILE", tmp_path / "history.jsonl")
    monkeypatch.setattr(ailine, "VOCAB_FILE", tmp_path / "vocab.json")   # 実 vocab を汚さない
    monkeypatch.setattr(ailine, "translate_task",
                        lambda model, task, book_meta, temperature=0.1:
                        {"op": "APPEND_TOTAL", "args": {"col": "小計", "label": "税込み合計"}})
    ns = argparse.Namespace(
        book=str(book), task="税込み合計を一番下に出して（消費税10%）", model="qwen2.5-coder:7b",
        refs=None, helpers=None, repair=0, temperature=0.2,
        dry=True, inplace=False, json=False, timeout=180.0, ask=False, values=False)
    rc = ailine.cmd_run(ns)
    captured = capsys.readouterr()
    assert rc == 0
    # ★ W8a 項目5: 表示ラベル「倍率」→「率」。
    assert "率:1.1（依頼文: 10%）" in captured.out

def test_cmd_run_dsl_append_total_clarify_hint_has_copy_paste_command(tmp_path, monkeypatch, capsys):
    book = _book(tmp_path, [["品目", "数量", "単価", "小計"], ["a", 3, 50000, 150000]])
    monkeypatch.setattr(ailine, "HISTORY_FILE", tmp_path / "history.jsonl")
    monkeypatch.setattr(ailine, "VOCAB_FILE", tmp_path / "vocab.json")   # 空の用語集
    monkeypatch.setattr(ailine, "translate_task",
                        lambda model, task, book_meta, temperature=0.1:
                        {"op": "APPEND_TOTAL", "args": {"col": "小計", "label": "税込み合計"}})
    ns = argparse.Namespace(
        book=str(book), task="税込み合計を出して", model="qwen2.5-coder:7b",
        refs=None, helpers=None, repair=0, temperature=0.2,
        dry=True, inplace=False, json=False, timeout=180.0, ask=False, values=False)
    rc = ailine.cmd_run(ns)
    captured = capsys.readouterr()
    assert rc == 3
    assert "ailine vocab add" in captured.out

def test_cmd_run_dsl_append_total_uses_registered_vocab(tmp_path, monkeypatch, capsys):
    book = _book(tmp_path, [["品目", "数量", "単価", "小計"], ["a", 3, 50000, 150000]])
    monkeypatch.setattr(ailine, "HISTORY_FILE", tmp_path / "history.jsonl")
    vocab_path = tmp_path / "vocab.json"
    monkeypatch.setattr(ailine, "VOCAB_FILE", vocab_path)
    ailine.vocab_add("消費税", 1.1, path=vocab_path)
    monkeypatch.setattr(ailine, "translate_task",
                        lambda model, task, book_meta, temperature=0.1:
                        {"op": "APPEND_TOTAL", "args": {"col": "小計", "label": "税込み合計"}})
    ns = argparse.Namespace(
        book=str(book), task="消費税込みの合計を出して", model="qwen2.5-coder:7b",
        refs=None, helpers=None, repair=0, temperature=0.2,
        dry=True, inplace=False, json=False, timeout=180.0, ask=False, values=False)
    rc = ailine.cmd_run(ns)
    captured = capsys.readouterr()
    assert rc == 0
    # ★ W8a 項目5: 表示ラベル「倍率」→「率」。
    assert "率:1.1（用語集: 消費税）" in captured.out

def test_cmd_run_dsl_append_total_warns_on_llm_factor_mismatch(tmp_path, monkeypatch, capsys):
    book = _book(tmp_path, [["品目", "数量", "単価", "小計"], ["a", 3, 50000, 150000]])
    monkeypatch.setattr(ailine, "HISTORY_FILE", tmp_path / "history.jsonl")
    monkeypatch.setattr(ailine, "VOCAB_FILE", tmp_path / "vocab.json")
    monkeypatch.setattr(ailine, "translate_task",
                        lambda model, task, book_meta, temperature=0.1:
                        {"op": "APPEND_TOTAL",
                         "args": {"col": "小計", "label": "税込み合計", "factor": 1.08}})
    ns = argparse.Namespace(
        book=str(book), task="税込み合計を一番下に出して（消費税10%）", model="qwen2.5-coder:7b",
        refs=None, helpers=None, repair=0, temperature=0.2,
        dry=True, inplace=False, json=False, timeout=180.0, ask=False, values=False)
    rc = ailine.cmd_run(ns)
    captured = capsys.readouterr()
    assert rc == 0
    assert "⚠" in captured.out
    assert "1.08" in captured.out and "1.1" in captured.out


# ===========================================================================
# ★ A': history.jsonl の provenance
# ===========================================================================

def test_build_history_entry_includes_provenance():
    result = {"ok": True, "attempts": 1, "provenance": {"factor": "依頼文: 10%"}}
    e = ailine.build_history_entry(result, Path("book.xlsx"), "タスク", "モデル", "none")
    assert e["provenance"] == {"factor": "依頼文: 10%"}

def test_build_history_entry_provenance_defaults_to_none():
    e = ailine.build_history_entry({"ok": True}, Path("book.xlsx"), "タスク", "モデル", "none")
    assert e["provenance"] is None


# --- CONTRACT: 自由生成に「新規シートを作らない」誘導が入っている ----------------

def test_contract_prompt_discourages_new_sheet_creation():
    assert "新しいシートを作らず" in ailine.CONTRACT


# ===========================================================================
# W6: 助言網を新規シートへ拡張（監査所見: 新規『集計』シートの全0埋めが★素通り）
# ===========================================================================

def test_new_sheet_advisories_flags_entirely_zero_filled_new_sheet(tmp_path):
    p = _book(tmp_path, [["商品", "金額"], ["a", 100], ["b", 200]])
    before = ailine.snapshot(p)
    wb = openpyxl.load_workbook(p)
    out = wb.create_sheet("集計")
    for r in range(1, 4):
        for c in range(1, 3):
            out.cell(row=r, column=c, value=0)
    wb.save(p)
    after = ailine.snapshot(p)
    lines = ailine.new_sheet_advisories(before, after)
    assert any("新規シート『集計』の" in ln and "★ 疑わしい" in ln and "値 0" in ln for ln in lines)

def test_new_sheet_advisories_empty_when_new_sheet_has_normal_varied_content(tmp_path):
    p = _book(tmp_path, [["商品", "金額"], ["a", 100], ["b", 200]])
    before = ailine.snapshot(p)
    wb = openpyxl.load_workbook(p)
    out = wb.create_sheet("集計")
    out.append(["部門", "合計"])
    out.append(["a", 100])
    out.append(["b", 200])
    wb.save(p)
    after = ailine.snapshot(p)
    assert ailine.new_sheet_advisories(before, after) == []

def test_new_sheet_advisories_empty_when_no_new_sheet(tmp_path):
    p = _book(tmp_path, [["a", 1], ["b", 2]])
    before = ailine.snapshot(p)
    wb = openpyxl.load_workbook(p)
    wb.active.cell(row=1, column=1, value="変更")
    wb.save(p)
    after = ailine.snapshot(p)
    assert ailine.new_sheet_advisories(before, after) == []

def test_build_advisories_still_flags_new_sheet_zero_fill_when_mixed_with_unrelated_change(tmp_path):
    # ★ 監査実測の再現: COMPUTE_COLUMN が小計列へ正しい(非一様な)値を書き込みつつ、
    #   同じ実行で新規『集計』シートが全0埋め(バグ)で作られるケース。
    #   旧実装は detect_uniform_fill が『変更セル全部が同一値』を diff 全体に要求するため、
    #   小計列の非一様な正しい値に引きずられて新規シートの異常が丸ごと素通りしていた。
    p = _book(tmp_path, [["品目", "小計"], ["a", None], ["b", None]])
    before = ailine.snapshot(p)
    wb = openpyxl.load_workbook(p)
    ws = wb.active
    ws.cell(row=2, column=2, value=150000)   # 小計へ正しい値（非一様）
    ws.cell(row=3, column=2, value=96000)
    out = wb.create_sheet("集計")
    for r in range(1, 3):
        for c in range(1, 3):
            out.cell(row=r, column=c, value=0)   # ★バグ想定: 新規シートが全0埋め
    wb.save(p)
    after = ailine.snapshot(p)
    lines = ailine.build_advisories("小計を計算して集計もして", before, after)
    assert any("新規シート『集計』の" in ln and "★ 疑わしい" in ln for ln in lines)

def test_detect_ghost_data_still_fires_for_existing_sheet_when_new_sheet_also_created(tmp_path):
    # ★ W6: 以前は新規シートが混ざるだけで rect=None → 関数全体が None を返し、
    #   既存シートの本当のゴーストデータまで丸ごと素通りしていた（回帰）。
    p = _book(tmp_path, [["a", 1], ["b", 2]])   # 使用範囲 A1:B2
    before = ailine.snapshot(p)
    wb = openpyxl.load_workbook(p)
    wb.active.cell(row=2, column=26, value="ghost")   # Z2、既存シートの範囲外
    wb.create_sheet("集計").cell(row=1, column=1, value="新規")   # 同時に新規シートも作る
    wb.save(p)
    after = ailine.snapshot(p)
    msg = ailine.detect_ghost_data(before, after)
    assert msg is not None
    assert "Z2" in msg

def test_detect_ghost_data_none_when_all_changes_are_in_a_new_sheet(tmp_path):
    p = _book(tmp_path, [["a", 1], ["b", 2]])
    before = ailine.snapshot(p)
    wb = openpyxl.load_workbook(p)
    wb.create_sheet("集計").cell(row=1, column=1, value="新規")
    wb.save(p)
    after = ailine.snapshot(p)
    assert ailine.detect_ghost_data(before, after) is None


# ===========================================================================
# W6: 依頼にないシート新設の申告
# ===========================================================================

def test_unrequested_new_sheet_advisory_fires_when_task_does_not_mention_sheet():
    before = {"sheets": ["Sheet"]}
    after = {"sheets": ["Sheet", "集計"]}
    lines = ailine.unrequested_new_sheet_advisory("部門ごとに金額を集計して", before, after)
    assert lines == ["★ 依頼にない新しいシートが作成されました（集計）"]

@pytest.mark.parametrize("phrase", ["集計シートを作って", "ピボットにして", "別に表を作って"])
def test_unrequested_new_sheet_advisory_silent_when_task_mentions_keyword(phrase):
    before = {"sheets": ["Sheet"]}
    after = {"sheets": ["Sheet", "集計"]}
    assert ailine.unrequested_new_sheet_advisory(phrase, before, after) == []

def test_unrequested_new_sheet_advisory_empty_when_no_new_sheet():
    before = {"sheets": ["Sheet"]}
    after = {"sheets": ["Sheet"]}
    assert ailine.unrequested_new_sheet_advisory("金額を並べ替えて", before, after) == []

def test_cmd_run_dsl_aggregate_flags_unrequested_new_sheet(tmp_path, monkeypatch, capsys):
    book = _book(tmp_path, [["部門", "金額"], ["a", 100], ["a", 200], ["b", 300]])
    monkeypatch.setattr(ailine, "HISTORY_FILE", tmp_path / "history.jsonl")
    monkeypatch.setattr(ailine, "translate_task",
                        lambda model, task, book_meta, temperature=0.1:
                        {"op": "AGGREGATE", "args": {"group_col": "部門", "value_col": "金額"}})
    monkeypatch.setattr(ailine, "normalize_book", lambda book, workdir, timeout=None: book)

    def fake_apply(out_book, code, workdir, helper_files=(), timeout=None):
        wb2 = openpyxl.load_workbook(out_book)
        out = wb2.create_sheet("集計")
        out.append(["部門", "合計 - 金額"])
        out.append(["a", 300])
        out.append(["b", 300])
        out.append(["合計", 600])
        wb2.save(out_book)
        return True, None, "ok"
    monkeypatch.setattr(ailine, "basrun_apply", fake_apply)

    ns = argparse.Namespace(
        book=str(book), task="部門ごとに金額を集計して", model="qwen2.5-coder:7b",
        refs=None, helpers=None, repair=0, temperature=0.2,
        dry=False, inplace=False, json=False, timeout=180.0, ask=False)
    rc = ailine.cmd_run(ns)
    captured = capsys.readouterr()
    assert rc == 0
    assert "★ 依頼にない新しいシートが作成されました（集計）" in captured.out

def test_cmd_run_dsl_aggregate_silent_when_task_mentions_sheet_keyword(tmp_path, monkeypatch, capsys):
    book = _book(tmp_path, [["部門", "金額"], ["a", 100], ["b", 300]])
    monkeypatch.setattr(ailine, "HISTORY_FILE", tmp_path / "history.jsonl")
    monkeypatch.setattr(ailine, "translate_task",
                        lambda model, task, book_meta, temperature=0.1:
                        {"op": "AGGREGATE", "args": {"group_col": "部門", "value_col": "金額"}})
    monkeypatch.setattr(ailine, "normalize_book", lambda book, workdir, timeout=None: book)

    def fake_apply(out_book, code, workdir, helper_files=(), timeout=None):
        wb2 = openpyxl.load_workbook(out_book)
        out = wb2.create_sheet("集計")
        out.append(["部門", "合計 - 金額"])
        out.append(["a", 100])
        out.append(["b", 300])
        out.append(["合計", 400])
        wb2.save(out_book)
        return True, None, "ok"
    monkeypatch.setattr(ailine, "basrun_apply", fake_apply)

    ns = argparse.Namespace(
        book=str(book), task="部門ごとに金額を集計シートにまとめて", model="qwen2.5-coder:7b",
        refs=None, helpers=None, repair=0, temperature=0.2,
        dry=False, inplace=False, json=False, timeout=180.0, ask=False)
    rc = ailine.cmd_run(ns)
    captured = capsys.readouterr()
    assert rc == 0
    assert "依頼にない新しいシート" not in captured.out
