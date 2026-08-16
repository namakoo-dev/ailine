"""C5: Claim 型の不変条件、および ailine_core/claim.py のレンダラ関数の単体テスト。

   ★ 純ロジックのみ・LibreOffice/basrun 不要（CI で走る）。
   ★ 統合側（format_plan_report/overall_verdict の出力文言そのもの・golden 相当の厳密一致）
   は tests/test_ailine.py に既存の test_format_plan_report_* / test_overall_verdict_* /
   test_cmd_run_dsl_success_prints_scope_note 等が引き続き見る（C5 はそこを1バイトも
   変えない純リファクタ）。ここでは Claim 単体の型としての不変条件だけを検査する。
"""
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from ailine_core.claim import Claim, render_single_op_claim


# --- 不変条件: basis の値域 ----------------------------------------------------

def test_claim_accepts_known_basis_values():
    for basis, scope in (("declaration", "対象:金額"), ("request", ""), ("diff_only", "")):
        c = Claim(verified=True, basis=basis, scope=scope, evidence="", observation_complete=True)
        assert c.basis == basis

def test_claim_rejects_unknown_basis():
    with pytest.raises(ValueError):
        Claim(verified=True, basis="vibes", scope="対象:金額", evidence="",
              observation_complete=True)


# --- ★ 不変条件本体: basis="declaration" は scope 必須 ----------------------------

def test_claim_declaration_basis_requires_nonempty_scope():
    """basis='declaration' で scope が空だと構築時に落ちる（型で強制・文言の書き忘れで
       なくテストが落ちる形にする、という DoD の要求そのもの）。"""
    with pytest.raises(ValueError):
        Claim(verified=True, basis="declaration", scope="", evidence="3 行を検証",
              observation_complete=True)

def test_claim_declaration_basis_with_scope_constructs_fine():
    c = Claim(verified=True, basis="declaration", scope="操作:並べ替え 対象:金額 順:降順",
               evidence="3 行を検証（降順）", observation_complete=True)
    assert c.scope == "操作:並べ替え 対象:金額 順:降順"

def test_claim_non_declaration_basis_allows_empty_scope():
    """basis="request"/"diff_only" は scope 必須の対象外（現状これらは実装されていないが、
       将来 basis を増やしたときに declaration だけの制約であることを固定する）。"""
    c = Claim(verified=False, basis="diff_only", scope="", evidence="",
               observation_complete=False)
    assert c.scope == ""


# --- 型として frozen（不変）であること -----------------------------------------

def test_claim_is_frozen():
    c = Claim(verified=True, basis="declaration", scope="対象:金額", evidence="",
               observation_complete=True)
    with pytest.raises(Exception):   # dataclasses.FrozenInstanceError（AttributeError のサブクラス）
        c.verified = False


# --- render_single_op_claim: verified=False を渡すのは呼び出し側のバグとして落ちる ------

def test_render_single_op_claim_rejects_unverified_claim():
    c = Claim(verified=False, basis="declaration", scope="対象:金額", evidence="",
               observation_complete=True)
    with pytest.raises(AssertionError):
        render_single_op_claim(c, "並べ替え")

def test_render_single_op_claim_matches_existing_banner_shape():
    """render_single_op_claim の戻り値が、既存の cmd_run_dsl 実装が出していたのと同じ
       2行（✓ バナー＋範囲注記）であることを確認する（byte 一致は golden 側が担保・
       ここでは形だけ）。"""
    c = Claim(verified=True, basis="declaration", scope="操作:並べ替え 対象:金額 順:降順",
               evidence="3 行を検証（降順）", observation_complete=True)
    lines = render_single_op_claim(c, "並べ替え")
    assert lines[0] == "\n✓ 達成を機械検証済み（操作:並べ替え）: 3 行を検証（降順）"
    assert lines[1].startswith("★「機械検証済み」は、上の「解釈:")
