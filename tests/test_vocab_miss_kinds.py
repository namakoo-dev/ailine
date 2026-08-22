"""W10 前提工事①（architect レビュー致命5-2）: 単発の語彙外（FREEFORM/OUT_OF_VOCAB）が
   落ちた理由を history の failure_kind に持ち込む検体。

   直す前: _normalize_plan_step（未知op/必須slot欠落/非dict要素）と translate_task の
   except（ollama不通/JSON不正/空応答）が全部 op="FREEFORM" 一段に合流し、
   cmd_refuse_vocab_miss は failure_kind="語彙外" 一色で記録していた（3つの病気の合流）。

   直した後: failure_kind は「語彙外/<理由>」（out_of_vocab / slot_missing /
   translate_error）に下位区分される。上位ラベル「語彙外」は接頭辞として残る
   （format_history_table の表示互換のため・表示は畳んでよいが記録は区分を保持する）。
   ★ translate_error だけは「頼める操作の一覧に照合できませんでした」と言うと嘘になる
   （照合を試みてすらいない）ので、断りの文言自体も差し替える。
"""
import sys
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import ailine  # noqa: E402

from _run_argv import run_argv  # noqa: E402  — C2: cmd_run 直呼び用 Namespace → main(argv) 変換


def _book(tmp_path):
    p = tmp_path / "b.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["商品", "金額"])
    ws.append(["a", 100])
    ws.append(["b", 200])
    wb.save(p)
    return p


def _run(tmp_path, monkeypatch, task="何かして"):
    book = _book(tmp_path)
    argv = run_argv(
        book=str(book), task=task, model="qwen2.5-coder:7b",
        refs=None, helpers=None, repair=0, temperature=0.2,
        dry=False, copy=False, json=False, timeout=180.0, ask=False)
    rc = ailine.main(argv)
    entries = ailine.read_history(max_n=10)
    return rc, entries


def test_unknown_op_is_recorded_as_out_of_vocab(tmp_path, monkeypatch, capsys):
    """未知の op（DELETE_ROW は OP_SCHEMA に無い）→ 語彙外/out_of_vocab。"""
    monkeypatch.setattr(ailine, "ollama_generate_json",
                         lambda model, msgs, temperature=0.1, num_predict=300:
                         '{"op": "DELETE_ROW", "args": {}}')
    rc, entries = _run(tmp_path, monkeypatch, task="行を消して")
    capsys.readouterr()
    assert rc == 3
    assert entries, "history に記録されていない"
    assert entries[0]["failure_kind"] == "語彙外/out_of_vocab", entries[0]


def test_missing_required_slot_is_recorded_as_slot_missing(tmp_path, monkeypatch, capsys):
    """SORT の必須 slot 'order' が欠落 → 語彙外/slot_missing。断りの文言は従来どおり
       （slot_missing は文言を変える指示が無い）。"""
    monkeypatch.setattr(ailine, "ollama_generate_json",
                         lambda model, msgs, temperature=0.1, num_predict=300:
                         '{"op": "SORT", "args": {"col": "金額"}}')
    rc, entries = _run(tmp_path, monkeypatch, task="並べ替えて")
    out = capsys.readouterr().out
    assert rc == 3
    assert entries[0]["failure_kind"] == "語彙外/slot_missing", entries[0]
    assert "照合できませんでした" in out


def test_translate_task_transport_failure_is_recorded_as_translate_error(tmp_path, monkeypatch, capsys):
    """ollama_generate_json が例外を投げる（不通/タイムアウト等）→ translate_task 自体が
       退避（この関数を経由しない翻訳失敗）→ 語彙外/translate_error。"""
    def boom(*a, **k):
        raise OSError("ollama 不通（テスト用）")
    monkeypatch.setattr(ailine, "ollama_generate_json", boom)
    rc, entries = _run(tmp_path, monkeypatch, task="何かして")
    capsys.readouterr()
    assert rc == 3
    assert entries[0]["failure_kind"] == "語彙外/translate_error", entries[0]


def test_translate_error_refusal_does_not_claim_vocab_mismatch(tmp_path, monkeypatch, capsys):
    """★ 本命: translate_error の断り文言に「語彙外」も「照合できませんでした」も
       出ない（照合を試みてすらいないので、それを言うと嘘になる）。代わりに
       翻訳そのものが失敗したことを言う。"""
    def boom(*a, **k):
        raise OSError("ollama 不通（テスト用）")
    monkeypatch.setattr(ailine, "ollama_generate_json", boom)
    rc, _entries = _run(tmp_path, monkeypatch, task="何かして")
    out = capsys.readouterr().out
    assert rc == 3
    assert "語彙外" not in out, f"translate_error なのに語彙外と表示した: {out}"
    assert "照合できませんでした" not in out, f"照合を試みていないのに照合失敗と言った: {out}"
    assert "翻訳に失敗" in out, f"翻訳失敗の理由が出ていない: {out}"
