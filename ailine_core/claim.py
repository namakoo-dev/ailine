"""Claim — 「✓ 機械検証済み」相当の表示を1箇所に閉じ込める。

★ C5（再設計 分割の続き）: ブラインド査定で、複合依頼「小計に数量×単価を入れて、
見出しを太字にして」の2段目(BOLD)が、実は前段(COMPUTE_COLUMN)が新規作成した列
（`"数量*単価"`）に適用されたのに『✓ 機械検証済み』が出た（2ファイルで100%再現）。
原因は構造で、事後条件は `args["target"]`＝**計画が宣言した対象**を検証しており、
機械は正しく動いていた（docs/behavior-corpus/nodes/verification-scope-honesty.md
参照）。検証していた対象が「依頼」でなく「計画」だっただけ、というズレ。

W10e（commit 414fdc3）で範囲注記（`_VERIFY_SCOPE_NOTE`）を文字列として足して応急処置
したが、文字列で謝っているだけで型が無かった。「✓ 機械検証済み」の文字列リテラルは
複数箇所（cmd_run_dsl 単発 / format_plan_report 段別 / overall_verdict 総合）に散って
おり、次に経路が増えたときに同じ穴が開く。

Claim は「検証は何と（basis）・どこまで（scope）・何を根拠に（evidence）通ったか」を
型で運ぶ。★「✓ 機械検証済み」相当の文字列はこのモジュールの render_* 関数からしか
出さない — 番人は tests/test_claim_render_guard.py（このモジュール以外に該当リテラルが
現れたら赤）。

★★ 今回は純リファクタ（C5 のスコープ制限）: 出力される文字列は1バイトも変えていない。
3箇所（cmd_run_dsl 単発の ✓ バナー・format_plan_report の段別 ✓ 行・overall_verdict の
「✓ すべて機械検証済み」）は互いに文言の型が違う（byte 一致ではない）ため統合はせず、
「この1モジュールからしか出せない」という置き場所だけを一元化した。
★ 矛盾警告（差分番人が「変化なし」と言っているのに verified=True が立つ場合の警告）は
今回は入れていない。挙動変更にあたるため、別コミットで宣言つきにやる。
"""
from __future__ import annotations

from dataclasses import dataclass

# ★ Claim.basis の許容値。"declaration"(計画が宣言した対象と照合・現状の全経路がこれ)
#   / "request"(依頼そのものと照合・現状は未実装) / "diff_only"(差分の有無のみ・意味は見ない)。
_BASES = frozenset({"declaration", "request", "diff_only"})


@dataclass(frozen=True)
class Claim:
    """『検証済み』表示が主張する範囲を型で運ぶ。verified=True を「✓ 機械検証済み」相当の
       文字列として描画してよいのは、このモジュールの render_* 関数だけ。

       verified: 機械検証（事後条件チェッカー）が通ったか。
       basis: 何と照合したか。"declaration"（計画が宣言した対象と照合）が現状唯一の実装
              （事後条件チェッカーは常に args 由来の対象を検証する＝
              verification-scope-honesty.md の「機械は正しく動いているが検証対象は
              依頼でなく計画」という限界そのもの）。
       scope: 照合した宣言そのもの（例:「操作:計算列 対象列:小計」「金額 = 数量*単価」）。
              basis="declaration" のときは必須 — 空だと構築時に落ちる（★下記
              __post_init__。「計画が宣言した対象」を空のまま『検証済み』と名乗ることは
              できない、という不変条件を型で強制する）。
       evidence: 検証の中身（例:「12 行を検証（式・キャッシュ値とも一致）」）。無ければ
              空文字列（None ではなく ""）。
       observation_complete: 検証が snapshot の MAX_ROWS 切り詰めの影響を受けていないか。
              ★ 現状の事後条件チェッカー(check_*)は openpyxl で対象ブックを直接・全行
              (ws.max_row まで)走査する。MAX_ROWS の影響を受けるのは「変更点:」表示用の
              snapshot() 側だけで、事後条件自体は既に exhaustive（`_truncation_notice` の
              `exhaustive_postcondition` 引数が指す区別と同じ）。そのため verified=True の
              Claim は現状すべて observation_complete=True になる — 将来 MAX_ROWS 以内に
              チェッカー自体を制限する経路が増えたら、そこだけ False を渡す。"""
    verified: bool
    basis: str
    scope: str
    evidence: str
    observation_complete: bool

    def __post_init__(self) -> None:
        if self.basis not in _BASES:
            raise ValueError(
                f"Claim.basis は {sorted(_BASES)} のいずれかである必要がある: {self.basis!r}")
        if self.basis == "declaration" and not self.scope:
            raise ValueError(
                "Claim(basis='declaration') は scope が必須 — 「計画が宣言した対象」を"
                "空にしたまま『検証済み』と名乗ることはできない"
                "（docs/behavior-corpus/nodes/verification-scope-honesty.md 参照）")


# --- ①③: 単発 DSL 経路（cmd_run_dsl）の ✓ バナー ----------------------------------

# ★ 致命1(W10e): 「機械検証済み」が保証する範囲を正直に言う一文。バナー自体の語は
#   既存テストが厳密一致で見ている（format_plan_report/overall_verdict）ため変えず、
#   この注記を別行として1回だけ添える方針にした（『計画どおり』と『依頼どおり』の
#   同一視をやめる・出しすぎない＝行数を増やさない）。
_VERIFY_SCOPE_NOTE = ("★「機械検証済み」は、上の「解釈:」行どおりに実行されたことの検証です。"
                       "その解釈が依頼の意図と合っているかまでは含みません — 「解釈:」行を確認してください。")
_VERIFY_SCOPE_NOTE_PLAN = ("★「機械検証済み」は各段の「◯段目: 解釈:」行どおりに実行されたことの検証です。"
                            "各段の解釈が依頼の意図と合っているかまでは含みません"
                            " — 各段の「解釈:」行を確認してください。")


def render_single_op_claim(claim: Claim, op_label: str) -> list:
    """単発 DSL 経路（cmd_run_dsl）で事後条件が pass した時の表示行（✓ バナー＋範囲注記）。
       ★ claim.verified=True の Claim だけを受け取る前提（呼び出し側で warn/fail は
       別文言に分岐済み・ここに verified=False を渡すのは呼び出し側のバグ）。
       op_label は OP_LABELS.get(op, op) の結果（ailine.py 側のカタログを持ち込まない
       ため、解決済み文字列として受け取る）。"""
    assert claim.verified, "render_single_op_claim は verified=True の Claim だけを受け取る"
    return [
        f"\n✓ 達成を機械検証済み（操作:{op_label}）: {claim.evidence}",
        _VERIFY_SCOPE_NOTE,
    ]


# --- ②③: 複合計画（M2c）の段別報告・総合判定 --------------------------------------

_ITEM_STATUS_MARK = {"ok": "✓", "warn": "⚠", "fail": "×"}


def _render_verified_fragment(claim: Claim) -> str:
    """Claim(verified=True) から format_plan_report の段別行に載せる『機械検証済み』
       文字列片を組む。★ ここが『機械検証済み』の文字列を生成する唯一の場所
       （呼び出し元はこの戻り値を使い、raw literal を自分では書かない）。"""
    assert claim.verified, "_render_verified_fragment は verified=True の Claim だけを受け取る"
    suffix = f"（{claim.evidence}）" if claim.evidence else ""
    return f"機械検証済み{suffix}"


def format_plan_report(items: list) -> list:
    """複合計画の項目別報告を行のリストにする。items: [(idx, label, status, detail), ...]
       status は 'ok'/'warn'/'fail'。★ FREEFORM 段の成功は『機械検証済み』とは絶対に言わない
       （✓ 適用され文書が変化した級に留める＝warn 表示の固定文言で担保）。
       ★ 止血1: 'warn' は2種類の由来を持つ — 語彙外(FREEFORM)段は detail=None の
       固定文言、DSL 段の事後条件が「検証対象が少なすぎる」場合は detail に理由が
       入るのでそちらを見せる（どちらも『機械検証済み』とは言わない点は共通）。"""
    lines = []
    for idx, label, status, detail in items:
        mark = _ITEM_STATUS_MARK[status]
        if status == "ok":
            # label は format_confirmation_line が返す「解釈: ...」の宣言テキストそのもの
            # （ailine.py 側で line[len("解釈: "):] を渡す）＝計画が宣言した対象。
            claim = Claim(verified=True, basis="declaration", scope=label,
                          evidence=detail or "", observation_complete=True)
            lines.append(f"{idx}. {label} → {mark} {_render_verified_fragment(claim)}")
        elif status == "warn":
            if detail:
                lines.append(f"{idx}. {label} → {mark} 機械検証できませんでした: {detail}")
            else:
                # ★ W8a 項目5: 「自由生成」→「AI が直接作成（機械保証なし）」（operator の語彙翻訳）。
                lines.append(f"{idx}. {label} → {mark} 語彙外のため AI が直接作成（機械保証なし）で実行（確認してください）")
        else:
            lines.append(f"{idx}. {label} → {mark} 未対応: {detail}")
    return lines


def overall_verdict(items: list) -> tuple:
    """(判定文, 総合status)。★ 総合判定は最弱の段に従う:
       全段 ok → 「✓ すべて機械検証済み」/ fail 無しで warn を含む → 「⚠ 一部は確認が必要です」/
       fail を含む → 失敗。『達成を機械検証済み』の語は機械検証が実際に通った段にだけ付ける
       （ここでは全段が ok の時だけそう言う）。"""
    statuses = {it[2] for it in items}
    if "fail" in statuses:
        return "× 一部の操作が未対応/失敗のため、達成できませんでした", "fail"
    if "warn" in statuses:
        # ★ 止血1: warn の由来は2種類（語彙外の自由生成／DSL段の検証対象不足）ある
        #   ため、どちらにも当てはまる言い方にする。
        # ★ W8a 項目5: 表示文言のみ「自由生成」→「AI が直接作成」（operator の語彙翻訳）。
        return ("⚠ 一部は確認が必要です（語彙外で AI が直接作成した段、または検証対象不足の段があり、"
                "機械検証はしていません）", "warn")
    # 全段 ok。scope は全段の宣言（label）を連ねた総覧 — 個々の evidence は段別報告
    # (format_plan_report) 側で既に見せているため、総合判定の表示自体には出さない
    # （現状の文言「✓ すべて機械検証済み」に evidence の suffix は元から付かない）。
    claim = Claim(verified=True, basis="declaration",
                  scope="; ".join(label for _idx, label, _status, _detail in items),
                  evidence="", observation_complete=True)
    return f"✓ すべて{_render_verified_fragment(claim)}", "ok"
