# W10 便A: 別名ストア（言い回し → op 名）── 実装より先に凍結した赤い検体。
# 出典: REVIEW-20260822-w10-architect.md（3-2/3-3/6-4）+ Namakoo 決裁（二段目翻訳・文字マッチ開始）。
#
# 契約:
#   ① aliases.json は vocab.json と別ファイル（load_vocab は float 以外を黙って捨てる=同居不能）
#   ② 検疫: op 名は OP_META に実在するものだけ・言い回しは _sanitize_vocab_term 同等の検疫・
#      件数上限あり（超えたら古いものから拒否でなく登録を断る）
#   ③ 照合は**双方向の包含判定**（「金額」⊂「税込金額」の断片問題の 3 度目を踏まない ──
#      片方向 in では 短い別名が長い依頼文の断片に誤ヒットする）
#   ④ remove / undo（直近の登録の取り消し）が効く ── 機械が書く層には取り消しが要る
#   ⑤ 別名は op を決めるだけ（slot は決めない）── lookup の戻り値は op 名のみ

import json
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
import ailine  # noqa: E402

sys.path.insert(0, str(Path(__file__).resolve().parent))
from test_golden_transcripts import _isolate, _run_main  # noqa: E402

needs_impl = pytest.mark.xfail(
    not hasattr(ailine, "ALIASES_FILE"),
    reason="別名ストア 未実装（契約は凍結済み・実装が来たら自動で実測に切替）",
    strict=True,
)


@needs_impl
def test_alias_add_list_remove_roundtrip(tmp_path, monkeypatch, capsys):
    """④: add → list に出る → remove → list から消える。vocab.json は無傷。"""
    _isolate(monkeypatch, tmp_path)
    monkeypatch.setattr(ailine, "ALIASES_FILE", tmp_path / "aliases.json")
    rc, out = _run_main(["alias", "add", "大きい順にして", "SORT"], capsys)
    assert rc == 0, out
    rc, out = _run_main(["alias", "list"], capsys)
    assert "大きい順にして" in out and "SORT" in out
    rc, out = _run_main(["alias", "remove", "大きい順にして"], capsys)
    assert rc == 0, out
    rc, out = _run_main(["alias", "list"], capsys)
    assert "大きい順にして" not in out


@needs_impl
def test_alias_undo_removes_most_recent(tmp_path, monkeypatch, capsys):
    """④: undo は直近の登録だけを取り消す。"""
    _isolate(monkeypatch, tmp_path)
    monkeypatch.setattr(ailine, "ALIASES_FILE", tmp_path / "aliases.json")
    _run_main(["alias", "add", "一番下に総額", "APPEND_TOTAL"], capsys)
    _run_main(["alias", "add", "大きい順にして", "SORT"], capsys)
    rc, out = _run_main(["alias", "undo"], capsys)
    assert rc == 0, out
    rc, out = _run_main(["alias", "list"], capsys)
    assert "大きい順にして" not in out, "直近でない方が消えた/残った"
    assert "一番下に総額" in out


@needs_impl
def test_alias_add_rejects_unknown_op(tmp_path, monkeypatch, capsys):
    """②: 実在しない op は登録できない（幻覚 op の封鎖）。"""
    _isolate(monkeypatch, tmp_path)
    monkeypatch.setattr(ailine, "ALIASES_FILE", tmp_path / "aliases.json")
    rc, out = _run_main(["alias", "add", "何かして", "MAGIC_OP"], capsys)
    assert rc != 0
    assert "MAGIC_OP" in out


@needs_impl
def test_alias_lookup_requires_containment_both_ways_guard(tmp_path, monkeypatch):
    """③: 断片問題 ── 別名「金額」が依頼文「税込金額の列を作って」に誤ヒットしない。
       ヒットの条件は『依頼文（の正規化形）が別名を **語として** 含む』では不十分で、
       別名が他の語の断片になっている場合は当てない。"""
    _isolate(monkeypatch, tmp_path)
    monkeypatch.setattr(ailine, "ALIASES_FILE", tmp_path / "aliases.json")
    ailine.save_alias("金額", "SORT")          # 意地悪な短い別名
    assert ailine.lookup_alias("税込金額の列を作って") is None, \
        "断片（金額 ⊂ 税込金額）に誤ヒット ── 3 度目の断片問題"
    assert ailine.lookup_alias("金額で並べて") == "SORT", "語としての一致は当ててよい"


@needs_impl
def test_alias_lookup_returns_op_name_only(tmp_path, monkeypatch):
    """⑤: lookup の戻り値は op 名（str）だけ。slot を運ばない（第二の DSL を作らない）。"""
    _isolate(monkeypatch, tmp_path)
    monkeypatch.setattr(ailine, "ALIASES_FILE", tmp_path / "aliases.json")
    ailine.save_alias("大きい順にして", "SORT")
    hit = ailine.lookup_alias("大きい順にして")
    assert hit == "SORT"
    assert isinstance(hit, str)


@needs_impl
def test_alias_store_capped(tmp_path, monkeypatch, capsys):
    """②: 上限（DEFAULT_VOCAB_MAX_ENTRIES と同数）を超える登録は断る。"""
    _isolate(monkeypatch, tmp_path)
    monkeypatch.setattr(ailine, "ALIASES_FILE", tmp_path / "aliases.json")
    cap = ailine.DEFAULT_VOCAB_MAX_ENTRIES
    for i in range(cap):
        assert ailine.save_alias(f"言い回し{i}", "SORT") is True
    assert ailine.save_alias("あふれた言い回し", "SORT") is False
