# W10 便B: 二段目翻訳（op 固定で args だけ埋めさせる）── 実装より先に凍結した赤い検体。
# 出典: REVIEW-20260822-w10-architect.md 1-1 + Namakoo 決裁「二段目翻訳」。
# 頷きの対象（本物の解釈行）と、別名ヒット後の翻訳の両方が使う心臓。
#
# 契約:
#   ① op は機械が固定する ── LLM が別の op を返しても、返り値の op は固定した op のまま
#      （毒の第一防壁: 頷いた op と違う操作が走る経路を構造的に塞ぐ）
#   ② LLM の応答が壊れていたら（JSON 不正・args 欠落）正直に None を返す ── 幻覚 args で
#      進まない。呼び出し側が CLARIFY に倒せる形
#   ③ プロンプトは第 4 の凍結定数 ── test_prompt_freeze の SHA 番人の対象に入る
#      （W9 実測: few-shot 1 例で誤断定 27.3% ── prompt は動くと壊れる部品）
#
# ★ 検体の訂正（2026-08-22 夕・測定器の修正 2 度目）: 初版は mock 対象を ollama_generate に
#   凍結していた ── それに合わせた実装が素の呼び出しを選び、7B の ```json 柵で正解応答を
#   捨てていた（実機 2/5）。本流 translate_task の窒息点は ollama_generate_json（format=json
#   強制）── mock 対象をそちらに訂正。期待値は不変。検体は窒息点の指定まで含めて仮説。

import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
import ailine  # noqa: E402

needs_impl = pytest.mark.xfail(
    not hasattr(ailine, "translate_task_fixed_op"),
    reason="二段目翻訳 未実装（契約は凍結済み・実装が来たら自動で実測に切替）",
    strict=True,
)

_META = {"sheets": ["Sheet"], "headers": {"Sheet": ["商品", "金額"]},
          "header_rows": {"Sheet": 1}}


@needs_impl
def test_fixed_op_is_forced_even_if_llm_disobeys(monkeypatch):
    """①: LLM が MERGE を返しても、固定した SORT が勝つ（args は採用してよいが op は不動）。"""
    monkeypatch.setattr(
        ailine, "ollama_generate_json",
        lambda model, msgs, temperature=0.2, num_predict=None:
        '{"op": "MERGE", "args": {"col": "金額", "order": "desc"}}')
    plan = ailine.translate_task_fixed_op("qwen2.5-coder:7b", "SORT",
                                            "金額の大きい順にして", _META)
    assert plan is not None
    assert plan["op"] == "SORT", f"固定した op が LLM に負けた: {plan}"
    assert plan["args"].get("col") == "金額"


@needs_impl
def test_fixed_op_normal_fill(monkeypatch):
    """正常系: 固定 op の args が埋まって返る。"""
    monkeypatch.setattr(
        ailine, "ollama_generate_json",
        lambda model, msgs, temperature=0.2, num_predict=None:
        '{"op": "SORT", "args": {"col": "金額", "order": "desc"}}')
    plan = ailine.translate_task_fixed_op("qwen2.5-coder:7b", "SORT",
                                            "金額の大きい順にして", _META)
    assert plan == {"op": "SORT", "args": {"col": "金額", "order": "desc"}}


@needs_impl
def test_fixed_op_broken_response_returns_none(monkeypatch):
    """②: JSON 不正は None（幻覚 args で進まない・呼び出し側が CLARIFY に倒す）。"""
    monkeypatch.setattr(
        ailine, "ollama_generate_json",
        lambda model, msgs, temperature=0.2, num_predict=None: "ごめんなさい、わかりません")
    plan = ailine.translate_task_fixed_op("qwen2.5-coder:7b", "SORT",
                                            "金額の大きい順にして", _META)
    assert plan is None


@needs_impl
def test_fixed_op_prompt_is_frozen_constant():
    """③: 第 4 の凍結定数が存在し、test_prompt_freeze の番人対象に入っている
       （SHA の値自体は prompt_freeze 側が持つ ── ここでは対象化だけを凍結）。"""
    assert hasattr(ailine, "TRANSLATION_FIXED_OP_SYSTEM"), "第 4 定数が無い"
    src = (Path(__file__).parent / "test_prompt_freeze.py").read_text(encoding="utf-8")
    assert "TRANSLATION_FIXED_OP_SYSTEM" in src, \
        "第 4 定数が prompt_freeze の番人対象に入っていない（W9 の 27.3% を繰り返す穴）"
