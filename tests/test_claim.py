"""C5/C9: Claim 型の不変条件、および ailine_core/claim.py のレンダラ関数の単体テスト。

   ★ 純ロジックのみ・LibreOffice/basrun 不要（CI で走る）。
   ★ 統合側（format_plan_report/overall_verdict の出力文言そのもの・golden 相当の厳密一致）
   は tests/test_ailine.py と tests/golden/f9_transcripts が見る。ここでは Claim 単体の
   型としての不変条件だけを検査する。
   ★★ C9 で加わった中核の不変条件: **verified=True の Claim は「原本(--copy なら .out)が
   確定した後に読み戻した」ものでなければ構築できない**（observed_after_apply/observed_on）。
   「反映前に決まった ✓」「未実行なのに ✓」という状態を、文言でなく型が禁止する。
"""
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from ailine_core.claim import (
    Claim, format_plan_preview, format_plan_report, overall_verdict,
    count_suspicious_advisories, render_applied_claim_demoted,
    render_applied_claim, render_applied_unobservable, render_applied_unverified,
)


def _observed(**overrides):
    """反映後に読み戻した（＝構築が許される）verified=True の Claim。"""
    kwargs = dict(verified=True, basis="declaration", scope="操作:並べ替え 対象:金額 順:降順",
                   evidence="Sheet: 4行×2列・値のあるセル 8", observation_complete=True,
                   observed_on=r"C:\tmp\b.xlsx", observed_after_apply=True)
    kwargs.update(overrides)
    return Claim(**kwargs)


# --- 不変条件: basis の値域 ----------------------------------------------------

def test_claim_accepts_known_basis_values():
    for basis, scope in (("declaration", "対象:金額"), ("request", ""), ("diff_only", "")):
        c = _observed(basis=basis, scope=scope)
        assert c.basis == basis

def test_claim_rejects_unknown_basis():
    with pytest.raises(ValueError):
        _observed(basis="vibes")


# --- 不変条件: basis="declaration" は scope 必須 ----------------------------------

def test_claim_declaration_basis_requires_nonempty_scope():
    """basis='declaration' で scope が空だと構築時に落ちる（型で強制・文言の書き忘れで
       なくテストが落ちる形にする、という DoD の要求そのもの）。"""
    with pytest.raises(ValueError):
        _observed(scope="")

def test_claim_declaration_basis_with_scope_constructs_fine():
    c = _observed()
    assert c.scope == "操作:並べ替え 対象:金額 順:降順"

def test_claim_non_declaration_basis_allows_empty_scope():
    """basis="request"/"diff_only" は scope 必須の対象外（現状これらは実装されていないが、
       将来 basis を増やしたときに declaration だけの制約であることを固定する）。"""
    c = Claim(verified=False, basis="diff_only", scope="", evidence="",
               observation_complete=False)
    assert c.scope == ""


# --- ★ C9 の中核: 反映後に読み戻していない ✓ は構築できない --------------------------

def test_claim_verified_requires_observation_after_apply():
    """★ 査定2本が名指しした「反映前に確定した ✓」を型で塞ぐ。既定値のまま
       （observed_after_apply=False）verified=True を作ろうとすると落ちる。"""
    with pytest.raises(ValueError) as e:
        Claim(verified=True, basis="declaration", scope="操作:並べ替え", evidence="ev",
              observation_complete=True)
    assert "observed_after_apply" in str(e.value)

def test_claim_verified_requires_observed_on_path():
    """どのファイルを読み戻して言っているのかを言えない ✓ は出せない。"""
    with pytest.raises(ValueError) as e:
        Claim(verified=True, basis="declaration", scope="操作:並べ替え", evidence="ev",
              observation_complete=True, observed_on="", observed_after_apply=True)
    assert "observed_on" in str(e.value)

def test_claim_unverified_may_be_unobserved():
    """verified=False（＝何も主張していない）側には読み戻しの義務を課さない
       ―― 制約は『✓ と名乗るとき』にだけ掛ける。"""
    c = Claim(verified=False, basis="declaration", scope="操作:並べ替え", evidence="",
               observation_complete=False)
    assert c.observed_after_apply is False and c.observed_on == ""


# --- 型として frozen（不変）であること -----------------------------------------

def test_claim_is_frozen():
    c = _observed()
    with pytest.raises(Exception):   # dataclasses.FrozenInstanceError（AttributeError のサブクラス）
        c.verified = False


# --- render_applied_claim: ✓ を出せる唯一の関数 ----------------------------------

def test_render_applied_claim_prints_file_and_readback_evidence():
    lines = render_applied_claim(_observed(), "b.xlsx")
    assert lines == ["\n✓ b.xlsx は機械検証済みの内容です"
                      "（適用後に読み戻して確認: Sheet: 4行×2列・値のあるセル 8）"]

def test_render_applied_claim_rejects_unverified_claim():
    c = Claim(verified=False, basis="declaration", scope="対象:金額", evidence="",
               observation_complete=True)
    with pytest.raises(AssertionError):
        render_applied_claim(c, "b.xlsx")

def test_render_applied_claim_adds_caveat_when_observation_incomplete():
    lines = render_applied_claim(_observed(observation_complete=False), "b.xlsx")
    assert len(lines) == 2 and "一部しか見ていません" in lines[1]

def test_render_applied_unverified_never_says_verified():
    line = render_applied_unverified("b.xlsx", "Sheet: 2行×2列・値のあるセル 4")[0]
    assert "✓" not in line and "機械保証はありません" in line

def test_render_applied_unobservable_never_says_verified():
    line = render_applied_unobservable("b.xlsx", "BadZipFile: not a zip")[0]
    assert "✓" not in line and "読み戻して確認できませんでした" in line


# --- ★ 決裁③(2026-08-22): ⚠ による ✓ の降格 -------------------------------------

def test_count_suspicious_advisories_counts_only_star_prefixed_lines():
    lines = [
        "★ 疑わしい: 変更が元データの範囲外です（Z2:Z6）",
        "（新規列の追加は意図どおりです）",   # 中立表示（★ 無し）は数えない
        "列 C: データ 3 行のうち 2 行を変更（1 行は未変更）",   # count_reconciliation の素の報告
        "★ 依頼で言及された『列Z』は存在しません/変更されていません",
        None,
        "",
    ]
    assert count_suspicious_advisories(lines) == 2

def test_count_suspicious_advisories_empty_is_zero():
    assert count_suspicious_advisories([]) == 0

def test_render_applied_claim_demoted_says_triangle_not_checkmark():
    lines = render_applied_claim_demoted(_observed(), "b.xlsx", 2)
    assert "✓" not in "".join(lines)
    assert lines[0].startswith("\n△ b.xlsx は宣言どおりの変化を確認しました")
    assert "⚠ 2 件を先に確認してください" in lines[0]

def test_render_applied_claim_demoted_rejects_unverified_claim():
    c = Claim(verified=False, basis="declaration", scope="対象:金額", evidence="",
               observation_complete=True)
    with pytest.raises(AssertionError):
        render_applied_claim_demoted(c, "b.xlsx", 1)

def test_render_applied_claim_demoted_rejects_zero_warning_count():
    """warning_count=0 で呼ぶのは呼び出し側の誤り（0件なら render_applied_claim を使う）。"""
    with pytest.raises(AssertionError):
        render_applied_claim_demoted(_observed(), "b.xlsx", 0)

def test_render_applied_claim_demoted_adds_caveat_when_observation_incomplete():
    lines = render_applied_claim_demoted(_observed(observation_complete=False), "b.xlsx", 1)
    assert len(lines) == 2 and "一部しか見ていません" in lines[1]


# --- 段別報告 / プレビュー: ✓ はもうどこにも出ない --------------------------------

def test_format_plan_report_ok_step_states_evidence_without_a_check_mark():
    lines = format_plan_report([(1, "操作:計算列 対象列:小計", "ok", "3 行を検証")])
    assert lines[0] == "1. 操作:計算列 対象列:小計 → 実行: 3 行を検証"
    assert "✓" not in lines[0]

def test_format_plan_report_ok_without_detail_still_reports_success():
    lines = format_plan_report([(1, "操作:太字", "ok", None)])
    assert lines[0] == "1. 操作:太字 → 実行"

def test_format_plan_preview_never_claims_verification():
    lines = format_plan_preview([(1, "操作:並べ替え", "ok", None),
                                  (2, "条件付き書式", "warn", None),
                                  (3, "操作:集計", "fail", "列がありません")])
    assert lines[0] == "1. 操作:並べ替え → 実行予定（未実行）"
    assert "✓" not in "".join(lines)
    assert "未実行" in lines[1] and lines[2].startswith("3. 操作:集計 → ×")


# --- 総合判定: 全段 ok では判定文を出さない（✓ は反映後の1行に一本化） --------------

def test_overall_verdict_all_ok_emits_no_line():
    line, v = overall_verdict([(1, "x", "ok", "r")])
    assert v == "ok" and line is None

def test_overall_verdict_warn_without_fail():
    line, v = overall_verdict([(1, "x", "ok", "r"), (2, "y", "warn", None)])
    assert v == "warn"
    assert "確認が必要" in line

def test_overall_verdict_fail_dominates_over_warn():
    line, v = overall_verdict(
        [(1, "x", "ok", "r"), (2, "y", "warn", None), (3, "z", "fail", "reason")])
    assert v == "fail"


def test_count_suspicious_advisories_counts_warn_prefixed_lines():
    """★ 片配線の追補(2026-08-22): 複合計画の見出し警告は「⚠ 」前置で step_advisories に
       入る ── ★ だけを数える初版の規則から漏れて ⚠ と ✓ が同居できた。印は ★/⚠ の両方。"""
    lines = ["⚠ 対象の列『税込』は直前の段で新規作成された列です。"
             "依頼に「見出し」とあるため、見出し行（行全体）を意図していないか確認してください",
             "（新規列の追加は意図どおりです）",
             "★ 疑わしい: 適用後にエラー値のセルが増えました（計1件）: A1"]
    assert count_suspicious_advisories(lines) == 2
