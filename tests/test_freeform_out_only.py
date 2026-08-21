"""freeform 廃止（Namakoo 最終決裁 2026-08-21 19:37）の契約検体。
   ★ 実装前に凍結した赤い検体。「不完全な機能が見えたままだと信頼感が失われる」。

   契約: 語彙外の依頼は 生成しない ── 0 秒で 断り + vocab_miss 記録 + 次の手。
   保証できることしかしない、が全面で成立する（K-1 の完成形）。"""
import sys
from pathlib import Path

import openpyxl
import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
import ailine  # noqa: E402

def _book(tmp_path):
    p = tmp_path / "b.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["商品", "金額"])
    ws.append(["a", 100])
    wb.save(p)
    return p


def _vocab_miss(monkeypatch):
    monkeypatch.setattr(ailine, "translate_task",
                        lambda model, task, book_meta, temperature=0.1:
                        {"plan": [{"op": "FREEFORM"}]})

    def _must_not_generate(*a, **k):
        raise AssertionError("★ 廃止後に生成が呼ばれた（生成しない、が契約）")
    monkeypatch.setattr(ailine, "ollama_generate", _must_not_generate, raising=False)
    monkeypatch.setattr(ailine, "ollama_generate_json", _must_not_generate, raising=False)
    monkeypatch.setattr(ailine, "basrun_apply",
                        lambda *a, **k: (_ for _ in ()).throw(AssertionError("適用も禁止")))


def test_vocab_miss_refuses_instantly_records_and_guides(tmp_path, monkeypatch, capsys):
    """★ 本命: 生成ゼロ（呼んだら検体が落ちる）・断りの理由・要望として記録した旨・
       次の手（ops への導線）が出る。原本は無変更。"""
    import hashlib
    _vocab_miss(monkeypatch)
    book = _book(tmp_path)
    sha = hashlib.sha256(book.read_bytes()).hexdigest()
    rc = ailine.main(["run", str(book), "取引先が同じ行を重複として削除して"])
    out = capsys.readouterr().out
    assert rc != 0
    assert hashlib.sha256(book.read_bytes()).hexdigest() == sha
    assert "照合できませんでした" in out or "一覧にありません" in out, f"断りの理由が無い: {out}"
    assert "要望" in out and "記録" in out, f"vocab_miss 記録の開示が無い: {out}"
    assert "ops" in out, f"次の手（ops 導線）が無い: {out}"


def test_allow_freeform_flag_gets_sunset_notice_not_generation(tmp_path, monkeypatch, capsys):
    """--allow-freeform は互換で受けて廃止告知 1 行 ── 生成はしない。"""
    _vocab_miss(monkeypatch)
    book = _book(tmp_path)
    rc = ailine.main(["run", str(book), "何か語彙に無いこと", "--allow-freeform"])
    out = capsys.readouterr().out
    assert rc != 0
    assert "廃止" in out or "提供を終了" in out, f"廃止告知が無い: {out}"


def test_vocab_miss_is_recorded_and_listable(tmp_path, monkeypatch, capsys):
    """★ 需要センサ: vocab_miss が記録され、後から一覧できる（頻度×原始性の開発キューの土台）。"""
    _vocab_miss(monkeypatch)
    # ★ 既存の history 機構は AILINE_HISTORY_DIR という環境変数を読まない（ailine.HISTORY_FILE
    #   というモジュール定数を直接 monkeypatch する既存の作法 ── 大量の既存テストが使っている
    #   形にそろえる。契約のアサーション行は変えていない）。
    monkeypatch.setattr(ailine, "HISTORY_FILE", tmp_path / "hist" / "history.jsonl")
    book = _book(tmp_path)
    ailine.main(["run", str(book), "取引先が同じ行を重複として削除して"])
    capsys.readouterr()
    rc = ailine.main(["history"])
    out = capsys.readouterr().out
    assert "重複として削除" in out, f"vocab_miss が履歴に残らない: {out}"
    assert "語彙外" in out or "要望" in out or "未対応" in out,         f"履歴で vocab_miss と分かる区分が無い: {out}"
