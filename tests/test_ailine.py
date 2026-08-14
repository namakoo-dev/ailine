"""ailine の純粋ロジックの単体テスト（ollama / LibreOffice を要さない部分）。
   生成・適用の統合は実機（basrun_spike）で検証済み。ここは回帰用の土台。
"""
import argparse
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

def test_cmd_run_dry_survives_history_write_failure(tmp_path, monkeypatch, capsys):
    # ★ 履歴の書き込み失敗で run 本体を落とさない（try で包み WARN のみ）ことを
    #   --dry（ollama 生成だけで LibreOffice を要さない）経路で確認する。
    book = _book(tmp_path, [["a", 1]])
    monkeypatch.setattr(ailine, "ollama_generate",
                        lambda model, msgs, temperature=0.2:
                        "Sub Run(oDoc As Object)\nEnd Sub")
    def boom(entry, path=None):
        raise OSError("書き込み失敗（テスト用）")
    monkeypatch.setattr(ailine, "append_history", boom)

    ns = argparse.Namespace(
        book=str(book), task="テスト", model="qwen2.5-coder:7b",
        refs=None, helpers=None, repair=0, temperature=0.2,
        dry=True, inplace=False, json=False, timeout=180.0)

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
