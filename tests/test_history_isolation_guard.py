"""W10 前提工事②（architect レビュー致命5-3）: tests/conftest.py の
   `_guard_real_home_writes` 番人が実際に効いていることの検体。

   ★ このテストは意図的に HISTORY_FILE/VOCAB_FILE/MISCLASS_FILE を monkeypatch しない
   （_isolate も呼ばない）── 「個々のテストが適用を忘れた」状況をそのまま再現し、
   conftest の autouse 番人だけで実ホームへの書き込みを防げているかを確かめる。
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import ailine  # noqa: E402


def test_guard_forces_history_vocab_misclass_off_real_home(tmp_path):
    real_history = Path.home() / ".ailine" / "history.jsonl"
    real_vocab = Path.home() / ".ailine" / "vocab.json"
    real_misclass = Path.home() / ".ailine" / "misclass.jsonl"

    # ★ 番人が既定を書き換えていること自体の確認（実ホームを指したままなら番人が効いていない）。
    assert ailine.HISTORY_FILE != real_history
    assert ailine.VOCAB_FILE != real_vocab
    assert ailine.MISCLASS_FILE != real_misclass
    assert str(tmp_path) in str(ailine.HISTORY_FILE)
    assert str(tmp_path) in str(ailine.VOCAB_FILE)
    assert str(tmp_path) in str(ailine.MISCLASS_FILE)

    before_history = real_history.read_bytes() if real_history.exists() else None
    before_vocab = real_vocab.read_bytes() if real_vocab.exists() else None
    before_misclass = real_misclass.read_bytes() if real_misclass.exists() else None

    # 実際に書き込み経路を叩く（_isolate 相当の monkeypatch は一切していない）。
    ailine.append_history({"marker": "guard-test-should-not-reach-real-home"})
    ailine.vocab_add("guard_test_term", 1.0)
    ailine.append_misclass({"marker": "guard-test-should-not-reach-real-home"})

    # 書き込み自体は起きている（guard 先の tmp ファイルに）。
    assert "guard-test-should-not-reach-real-home" in ailine.HISTORY_FILE.read_text(encoding="utf-8")
    assert "guard_test_term" in ailine.VOCAB_FILE.read_text(encoding="utf-8")
    assert "guard-test-should-not-reach-real-home" in ailine.MISCLASS_FILE.read_text(encoding="utf-8")

    # 実ホームは無傷。
    after_history = real_history.read_bytes() if real_history.exists() else None
    after_vocab = real_vocab.read_bytes() if real_vocab.exists() else None
    after_misclass = real_misclass.read_bytes() if real_misclass.exists() else None
    assert before_history == after_history, "実 ~/.ailine/history.jsonl に書き込んでしまった"
    assert before_vocab == after_vocab, "実 ~/.ailine/vocab.json に書き込んでしまった"
    assert before_misclass == after_misclass, "実 ~/.ailine/misclass.jsonl に書き込んでしまった"
