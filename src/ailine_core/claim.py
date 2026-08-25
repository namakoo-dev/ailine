"""Claim — 「✓」の発生点を「原本が確定した後の1箇所」に閉じ込める。

★ C9（この回）: ブラインド査定2本が独立に同じことを書いた ―― `✓ 機械検証済み` が
**3つの違う意味**で出ていた。
  ①原本に入った ②入っていない（複合計画の別の段が失敗して全段破棄。それでも成功した段に
  ✓ が出て、しかも「原本は変更していません」の行が無い）③そもそも実行していない
  （`--dry` が `✓ 機械検証済み（未実行・プレビューのみ）`）。
査定者の言:「毎回 openpyxl で開くまでファイルがどうなったか分からなかった。この道具の
一番の売りは検証なのに。」

真因は位置だった: 段別 ✓ は `format_plan_report`、`--dry` の ✓ は `_preview_dsl_plan` で、
どちらも**原本反映の可否より前**に確定していた ―― ✓ は最終ファイルを一度も読み戻して
いなかった。

**この回で `✓` の意味を1種類にする:「あなたのファイルは今こうなっている（機械が読み戻して
確かめた）」。** そのため:
  1. 段別の行から `✓` を外した（段は evidence だけ述べる。成功は伝えるが ✓ とは呼ばない）
  2. `--dry` は Claim を構築しない（`format_plan_preview` というプレビュー専用レンダラに分けた。
     status を差し替えて同じレンダラを通す形は、未実行を『検証済み』と呼べる経路を残すため採らない）
  3. 総合の `✓` は `ailine.py` の `_finish_apply`（原本 or `.out` が確定した直後）で1回だけ
  4. `observed_after_apply=False` の Claim は**構築できない**（下記 `__post_init__`。既存の
     「basis='declaration' なら scope 必須」と同じ手口の再利用 ―― 新しい概念は増やさない）

★★ 採らなかった設計（明示的に否定された）: 「最終ファイルで全段の事後条件を再実行する」。
不健全 ―― `APPEND_TOTAL` の後に `SORT` が来れば合計行の位置は正当に動き
（`check_append_total` は `"=SUM("` の初出行で合計行を探す）、**正しい run が偽 fail になる**。
再実行ではなく「反映が成功したことを読み戻すだけ」でよい。

★ ここは純ロジック（ファイルを開かない）。読み戻しの実行は `ailine.py` 側
（`observe_book_state`）にあり、Claim はその結果を受け取るだけ。
★ 『機械検証済み』相当の文字列はこのモジュールの render_* 関数からしか出さない ―― 番人は
tests/test_claim_render_guard.py（このモジュール以外に該当リテラルが現れたら赤）。
"""
from __future__ import annotations

from dataclasses import dataclass

# ★ Claim.basis の許容値。"declaration"(計画が宣言した対象と照合・現状の全経路がこれ)
#   / "request"(依頼そのものと照合・現状は未実装) / "diff_only"(差分の有無のみ・意味は見ない)。
_BASES = frozenset({"declaration", "request", "diff_only"})


@dataclass(frozen=True)
class Claim:
    """『検証済み』表示が主張する範囲を型で運ぶ。verified=True を「✓」相当の文字列として
       描画してよいのは、このモジュールの render_* 関数だけ。

       verified: 機械検証（事後条件チェッカー）が全段通ったか。
       basis: 何と照合したか。"declaration"（計画が宣言した対象と照合）が現状唯一の実装
              （事後条件チェッカーは常に args 由来の対象を検証する＝
              verification-scope-honesty.md の「機械は正しく動いているが検証対象は
              依頼でなく計画」という限界そのもの）。
       scope: 照合した宣言そのもの（例:「操作:計算列 対象列:小計」「金額 = 数量*単価」）。
              basis="declaration" のときは必須 — 空だと構築時に落ちる。
       evidence: ★ C9 で意味が変わった。「読み戻して観測した最終ファイルの状態」
              （例:「Sheet: 4行×2列・値のあるセル 8」）。段ごとの事後条件の理由
              （「3 行を検証（降順）」等）は段別報告の側が述べる ―― ✓ の evidence は
              **最終ファイルだけから独立に再導出できる事実**に限る（事後条件の再実行は
              モジュール冒頭の理由で採らない）。
       observation_complete: 読み戻しが最終ファイル全体を見たか。★ observe_book_state は
              全シート・全行を走査するので現状すべて True（snapshot() の MAX_ROWS 切り詰めは
              『変更点:』表示側だけの話で、ここには効かない）。将来、読み戻しを一部に
              制限する経路が増えたら、そこだけ False を渡す。
       observed_on: 読み戻したファイルのパス（原本、--copy なら .out）。verified=True なら必須。
       observed_after_apply: 原本（or .out）への反映が確定した**後**に読み戻したか。
              ★ verified=True の Claim はこれが True でなければ構築できない（下記）。
              これが C9 の中核 ―― 「反映前に確定した ✓」という状態を型が禁止する。"""
    verified: bool
    basis: str
    scope: str
    evidence: str
    observation_complete: bool
    observed_on: str = ""
    observed_after_apply: bool = False

    def __post_init__(self) -> None:
        if self.basis not in _BASES:
            raise ValueError(
                f"Claim.basis は {sorted(_BASES)} のいずれかである必要がある: {self.basis!r}")
        if self.basis == "declaration" and not self.scope:
            raise ValueError(
                "Claim(basis='declaration') は scope が必須 — 「計画が宣言した対象」を"
                "空にしたまま『検証済み』と名乗ることはできない"
                "（docs/behavior-corpus/nodes/verification-scope-honesty.md 参照）")
        if self.verified and not self.observed_after_apply:
            raise ValueError(
                "Claim(verified=True) は observed_after_apply=True が必須 — 原本(--copy なら"
                " .out)が確定する前に決まった結果を『✓』と名乗ることはできない"
                "（docs/behavior-corpus/nodes/verified-means-readback.md 参照）")
        if self.verified and not self.observed_on:
            raise ValueError(
                "Claim(verified=True) は observed_on（読み戻したファイルのパス）が必須 — "
                "どのファイルを見て言っているのかを言えない ✓ は出せない")


# --- ★ ✓ を出せる唯一の場所（原本 or .out が確定した後・run につき1回） ------------------

# ★★ 単位E: 常時表示の範囲注記（旧 _VERIFY_SCOPE_NOTE / _VERIFY_SCOPE_NOTE_PLAN）を廃止した。
#   あれは ✓ が出る**全 run で必ず**出ていた ―― 発火率 100% ＝ 情報量ゼロ。誤爆する警告を
#   批判している設計側が同じ病気にかかっていた（オオカミ少年化は「毎回出る」でも起きる）。
#   ★ 消したのは「常時」だけ。**範囲を明示するという役割そのものは残す** ―― その run で
#   実際に依頼文と照合できなかった対象だけを名指しする1文（下記 render_scope_notes）に
#   置き換えた。照合できた run では何も出ない＝出たときに意味がある。
#   （3段階の仕分けは ailine_core/subject.py。①照合できた=満額 ②無言=この1文 ③矛盾=✓ を出さない）


def render_scope_notes(unspoken_subjects: list) -> list:
    """② の run 固有の1文。unspoken_subjects は「対象『X』」等の表示句のリスト（空なら何も出さない）。
       ★ ✓ の行の直後に1回だけ添える。『計画どおり』と『依頼どおり』の同一視をやめる、という
       旧注記の役割はここが引き継ぐ ―― ただし**その run で実際に照合できなかった対象**だけを言う。"""
    if not unspoken_subjects:
        return []
    return [f"★ ただし{'・'.join(unspoken_subjects)}は依頼文の語と機械照合していません"
            f"（ブックの実体・既定から機械決定しました） — 「解釈:」行を確認してください。"]


def count_suspicious_advisories(lines) -> int:
    """★ 決裁③(2026-08-22): 「疑わしい系の ⚠」の件数を、advisory の**結果オブジェクト**
       （render 済みの最終出力ではなく、build_advisories/check_write_preconditions_detail が
       返す文字列そのもの）から機械的に数える。
       ★ 判別規則: ★ または ⚠ で始まる行を数える。中立表示（「（新規列の追加は意図どおり
       です）」「（表示は先頭 N 行の変化のみ…）」等・parens のみで印を持たない）や
       count_reconciliation の素の件数報告（「列 C: データ 3 行のうち…」）は疑わしい系では
       ないので数えない ―― advisory 生成関数群（ailine.py の _structural_advisories 系・
       write_precondition.py の _check_* 系・_maybe_warn_header_col_mismatch）は、疑わしいと
       言いたい行にだけ既に ★/⚠ を付けている（この関数はその既存の合図を読むだけで、
       新しい判定基準を作らない）。⚠ を含めるのは片配線の追補(2026-08-22 検分):
       複合計画の見出し警告（⚠ 前置で step_advisories に入る）が ★ だけの規則から漏れて
       ⚠ と ✓ が同居できた。
       lines: 文字列のイテラブル（None・空文字は無視）。"""
    return sum(1 for ln in lines
               if isinstance(ln, str) and ln.lstrip().startswith(("★", "⚠")))


def render_applied_claim_demoted(claim: Claim, display_name: str, warning_count: int) -> list:
    """★ 決裁③(2026-08-22): 疑わしい ⚠ が1件でも出た run は「✓ 機械検証済み」を名乗らない
       ―― ✓ の絶対性の適用（✓ が出た run では買い手は差分を読まなくなる、という実測を
       踏まえ、⚠ と ✓ の同居そのものを無くす）。
       ★ verified=True の Claim をそのまま受け取る（宣言どおりの照合自体は成立している ──
       嘘ではない。文字だけ変える）: 「△ 宣言どおりの変化は確認しました」で検証が通った
       事実と、「⚠ N 件を先に確認してください」で ⚠ の存在を分けて言う。
       ★ warning_count は呼び出し側が数えた「疑わしい ⚠」の総数（0 で呼んではいけない ──
       0 件なら render_applied_claim を使う。呼び出し側の choke point は ailine.py
       の _finish_apply 1箇所）。"""
    assert claim.verified and claim.observed_after_apply, (
        "render_applied_claim_demoted も反映後に読み戻した verified=True の Claim だけを受け取る")
    assert warning_count > 0, "render_applied_claim_demoted は warning_count > 0 の時だけ呼ぶ"
    line = (f"\n△ {display_name} は宣言どおりの変化を確認しました"
            f"（適用後に読み戻して確認: {claim.evidence}）"
            f" ── ただし ⚠ {warning_count} 件を先に確認してください")
    if not claim.observation_complete:
        return [line, "★ ただし読み戻しは最終ファイルの一部しか見ていません（全体は未確認）。"]
    return [line]


def render_applied_claim(claim: Claim, display_name: str) -> list:
    """★ 『✓』を出せる唯一の関数。原本（--copy なら .out）が確定した後、その最終ファイルを
       読み戻して確かめた結果だけを述べる。
       display_name は表示用のファイル名（Claim.observed_on はフルパスを運ぶ）。
       ★ verified=False / observed_after_apply=False の Claim はそもそも構築できないので、
       ここに来る Claim は定義上『反映後に読み戻し済み』。"""
    assert claim.verified and claim.observed_after_apply, (
        "render_applied_claim は反映後に読み戻した verified=True の Claim だけを受け取る")
    line = f"\n✓ {display_name} は機械検証済みの内容です（適用後に読み戻して確認: {claim.evidence}）"
    if not claim.observation_complete:
        return [line, "★ ただし読み戻しは最終ファイルの一部しか見ていません（全体は未確認）。"]
    return [line]


def render_unverified_advisories(unverified) -> list:
    """検証できなかった行を、人へ見せる ⚠ 行にする。★ ⚠ で始めるので決裁③が数えて ✓ を △ に降ろす。

    ★ 2026-08-25（塊②）: ここが**文言の唯一の実装**。以前は ailine 本体に同名関数が在り、
      本番（dsl_step）は同じ文を**その場に書き写して**いた ── 名前つきの方は誰も呼ばず、
      試験だけがそれを守っていた（番人が本番でない方を見張る片配線）。
      ailine_core に置いたのは、本番の合流点である dsl_step から直接呼べる側だから。
    """
    return [f"⚠ {u['rows']} 行は検証できていません（{u['why']}）"
            " ── この行については「宣言どおり」と言えません"
            for u in (unverified or [])]


def render_applied_unverified(display_name: str, observed: str) -> list:
    """機械保証が無い経路（自由生成・検証対象不足の段を含む計画）の反映後の行。✓ は使わない。
       読み戻し自体は行うので「今このファイルはこうなっている」だけは同じ強度で言える。"""
    return [f"\n⚠ {display_name} に適用しましたが、機械保証はありません"
            f"（適用後に読み戻して確認: {observed}）"]


def render_applied_unobservable(display_name: str, error: str) -> list:
    """反映はできたが読み戻せなかった（ファイルが壊れている/開けない）ときの行。
       ★ ここで ✓ を出さないことがこの設計の要 ―― 読み戻せていないなら何も保証しない。"""
    return [f"\n⚠ {display_name} に適用しましたが、読み戻して確認できませんでした（{error}）"
            f" — ファイルを開いて中身を確かめてください"]


# --- 複合計画（M2c）の段別報告・プレビュー・総合判定 --------------------------------

# ★ C9: "ok" の記号を廃止した（✓ は最終の1行だけ・記号は増やさず減らす）。
_ITEM_STATUS_MARK = {"warn": "⚠", "fail": "×"}


def format_plan_report(items: list) -> list:
    """複合計画の項目別報告を行のリストにする。items: [(idx, label, status, detail), ...]
       status は 'ok'/'warn'/'fail'。
       ★ C9: 'ok' の段は evidence（事後条件が実際に何を見たか）だけを述べ、✓ とは呼ばない
       ―― その段が成功したことと、あなたのファイルが今どうなっているかは別の主張であり、
       後者は原本が確定した後の1行（render_applied_claim）だけが言う。
       ★ FREEFORM 段の成功は元から『機械検証済み』とは言わない（warn 表示の固定文言）。
       ★ 'warn' は2種類の由来を持つ — 語彙外(FREEFORM)段は detail=None の固定文言、
       DSL 段の事後条件が「検証対象が少なすぎる」場合は detail に理由が入る。"""
    lines = []
    for idx, label, status, detail in items:
        if status == "ok":
            lines.append(f"{idx}. {label} → 実行: {detail}" if detail
                          else f"{idx}. {label} → 実行")
            continue
        mark = _ITEM_STATUS_MARK[status]
        if status == "warn":
            if detail:
                lines.append(f"{idx}. {label} → {mark} 機械検証できませんでした: {detail}")
            else:
                # ★ W8a 項目5: 「自由生成」→「AI が直接作成（機械保証なし）」（operator の語彙翻訳）。
                lines.append(f"{idx}. {label} → {mark} 語彙外のため AI が直接作成（機械保証なし）で実行（確認してください）")
        else:
            lines.append(f"{idx}. {label} → {mark} 未対応: {detail}")
    return lines


def format_plan_preview(items: list) -> list:
    """★ C9: `--dry` 専用のレンダラ。**Claim を構築しない**。
       査定が指摘した3つ目の意味（そもそも実行していないのに ✓）は、実行経路と同じ
       レンダラへ status='ok' を流し込んでいたことが原因だった。プレビューは
       「まだ何もしていない」としか言えないので、言える文だけを持つ別の関数に分ける。"""
    lines = []
    for idx, label, status, detail in items:
        if status == "ok":
            lines.append(f"{idx}. {label} → 実行予定（未実行）")
        elif status == "warn":
            lines.append(f"{idx}. {label} → 実行時に AI が直接作成（機械保証なし）で対応（未実行）")
        else:
            lines.append(f"{idx}. {label} → {_ITEM_STATUS_MARK['fail']} 未対応: {detail}")
    return lines


def overall_verdict(items: list) -> tuple:
    """(判定文 or None, 総合status)。★ 総合判定は最弱の段に従う:
       fail を含む → 失敗 / fail 無しで warn を含む → 「⚠ 一部は確認が必要です」/
       全段 ok → **判定文は出さない**（None）。
       ★ C9: 旧「✓ すべて機械検証済み」はここで出していたが、この時点では原本へ反映
       できるかどうかがまだ分かっていない（--copy か・置換が成功するか）。全段 ok の
       事実は呼び出し側が machine_verified として持ち回り、原本が確定した後の1行が言う。"""
    statuses = {it[2] for it in items}
    if "fail" in statuses:
        return "× 一部の操作が未対応/失敗のため、達成できませんでした", "fail"
    if "warn" in statuses:
        # ★ warn の由来は2種類（語彙外の自由生成／DSL段の検証対象不足）あるため、
        #   どちらにも当てはまる言い方にする。
        return ("⚠ 一部は確認が必要です（語彙外で AI が直接作成した段、または検証対象不足の段があり、"
                "機械検証はしていません）", "warn")
    return None, "ok"
