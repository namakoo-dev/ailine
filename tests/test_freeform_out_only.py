"""freeform の書き込み面の制御（Namakoo 決裁 2026-08-21 18:48「原本への書き込みはご法度」）。
   ★ 実装前に凍結した赤い検体（operator 盲検 7 度目・摩擦②由来）。

   契約: 機械保証の無い経路（freeform）は、--allow-freeform があっても原本に書けない ──
   常に .out へ。原本不可侵を「機械保証なし」の側にだけ機械で強制する。
   + history は自由生成と機械検証を区分して記録する（摩擦⑦・需要センサの土台）。"""
import sys
from pathlib import Path

import openpyxl
import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
import ailine  # noqa: E402

pytestmark = pytest.mark.xfail(strict=True, reason="freeform .out 強制 + history 区分 実装前")


def _book(tmp_path):
    p = tmp_path / "b.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["商品", "金額"])
    ws.append(["a", 100])
    wb.save(p)
    return p


def _freeform_mocks(monkeypatch):
    monkeypatch.setattr(ailine, "translate_task",
                        lambda model, task, book_meta, temperature=0.1:
                        {"plan": [{"op": "FREEFORM"}]})
    monkeypatch.setattr(ailine, "ollama_generate",
                        lambda *a, **k: 'Sub Run(oDoc As Object)\n'
                        '    oDoc.Sheets.getByIndex(0).getCellByPosition(2, 0).setString("x")\n'
                        'End Sub', raising=False)

    def fake_apply(out_book, code, workdir, helper_files=(), timeout=None):
        wb = openpyxl.load_workbook(out_book)
        wb.active["C1"] = "x"
        wb.save(out_book)
        return True, None, "ok"
    monkeypatch.setattr(ailine, "basrun_apply", fake_apply)


def test_freeform_never_writes_to_original_even_with_allow_flag(tmp_path, monkeypatch, capsys):
    """★ 本命: --allow-freeform を付けても原本は 1 バイトも変わらず、結果は .out へ。
       その旨（機械保証が無いため原本には書きません）が報告に出る。"""
    import hashlib
    _freeform_mocks(monkeypatch)
    book = _book(tmp_path)
    sha = hashlib.sha256(book.read_bytes()).hexdigest()
    rc = ailine.main(["run", str(book), "何か語彙に無いことをして", "--allow-freeform"])
    out = capsys.readouterr().out
    assert hashlib.sha256(book.read_bytes()).hexdigest() == sha, "★ 原本が書き換わった（ご法度）"
    assert (tmp_path / "b.out.xlsx").exists(), f".out が無い:\n{out}"
    assert "原本には書きません" in out or "原本には触りません" in out or ".out" in out, \
        f".out 強制の開示が無い:\n{out}"


def test_history_distinguishes_freeform_from_verified(tmp_path, monkeypatch, capsys):
    """★ 摩擦⑦: history の一覧で自由生成（機械保証なし）と機械検証済みが区別できる。
       これが「freeform に落ちた依頼の一覧 = 需要センサ」の土台になる。"""
    _freeform_mocks(monkeypatch)
    monkeypatch.setenv("AILINE_HISTORY_DIR", str(tmp_path / "hist"))
    book = _book(tmp_path)
    rc = ailine.main(["run", str(book), "何か語彙に無いことをして", "--allow-freeform"])
    capsys.readouterr()
    rc = ailine.main(["history"])
    out = capsys.readouterr().out
    assert "自由生成" in out or "機械保証なし" in out or "freeform" in out, \
        f"history に自由生成の区分が無い:\n{out}"


def _freeform_with_code(monkeypatch, code):
    monkeypatch.setattr(ailine, "translate_task",
                        lambda model, task, book_meta, temperature=0.1:
                        {"plan": [{"op": "FREEFORM"}]})
    monkeypatch.setattr(ailine, "ollama_generate", lambda *a, **k: code, raising=False)
    def fake_apply(out_book, code_, workdir, helper_files=(), timeout=None):
        raise AssertionError("★ 適用まで到達してはならない（関所の手前で拒否するはず）")
    monkeypatch.setattr(ailine, "basrun_apply", fake_apply)


def test_freeform_code_with_path_literal_is_refused_before_apply(tmp_path, monkeypatch, capsys):
    """★ Namakoo の指摘（2026-08-21 18:51）: 文書ハンドルはコピーでも、生成 Basic は
       loadComponentFromURL 等で任意のパスに自力で触れる。operator ⑤はパスの捏造まで実演。
       契約: パス風リテラルを含む freeform コードは --allow-freeform があっても適用拒否
       （正当な freeform コードにパスリテラルの用途は無い ── oDoc しか要らない）。"""
    _freeform_with_code(monkeypatch,
        'Sub Run(oDoc As Object)\n'
        '    d = createUnoService("com.sun.star.frame.Desktop")\n'
        '    o = d.loadComponentFromURL("file:///C:/path/to/lookup.xlsx", "_blank", 0, Array())\n'
        'End Sub')
    book = _book(tmp_path)
    rc = ailine.main(["run", str(book), "照合して", "--allow-freeform"])
    out = capsys.readouterr().out
    assert rc != 0, f"パスリテラル入りのコードが適用まで届いた:\n{out}"
    assert "lookup.xlsx" in out or "パス" in out or "別のファイル" in out, \
        f"何を拒否したかの名指しが無い:\n{out}"


def test_freeform_code_with_dangerous_api_is_refused(tmp_path, monkeypatch, capsys):
    """危険 API（Shell/Kill/FileCopy 等）を含む生成コードは適用拒否（名指しつき）。"""
    _freeform_with_code(monkeypatch,
        'Sub Run(oDoc As Object)\n'
        '    Shell("cmd /c echo x", 1, "", True)\n'
        'End Sub')
    book = _book(tmp_path)
    rc = ailine.main(["run", str(book), "何かして", "--allow-freeform"])
    out = capsys.readouterr().out
    assert rc != 0, f"Shell 入りのコードが適用まで届いた:\n{out}"
    assert "Shell" in out, f"どの語を拒否したかの名指しが無い:\n{out}"
