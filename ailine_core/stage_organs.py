"""stage_organs — 段種 × 器官の宣言表と網羅の番人（C6）。

★ 背景（独立監査が名指しした真因）: この repo の欠陥の共通の根は「器官が『呼び出し規約』
でなく『記憶』で適用されている」こと ―― 同じ知見が、見つかった場所にだけ適用されて
他所へ伝播しない。実例2つ:
  - scan_rate_literals（A' 原則を守る唯一の機械）は単発の自由生成(cmd_run_freeform)には
    あったが、複合計画の自由生成段(run_freeform_plan_step)には無かった（W10f で発見・修正。
    査定5回が偶然当たらなかっただけで穴は開いていた）。
  - 破壊の関所が `if op != "COMPUTE_COLUMN"` という op ごとの白名簿だった（W10c で
    OP_WRITE_TARGET という宣言駆動へ一般化して解決 — ailine.py 参照。この解法だけが
    再発していない）。

OP_WRITE_TARGET が「op ごとに書き込み先列があるか」を宣言駆動で一般化して以来
再発していない、という唯一実証された形（宣言表 + 網羅の番人）を、段種（実行経路の単位）
× 器官（同じ知見が繰り返し実装される道具）の全体へ一般化したのがこのモジュール。

★★ この表は現実を写す。理想を書かない: 埋めていく過程で「ここには器官があるべきなのに
無い」マス目が見つかっても、器官を足さずに None（無い）と書いて report 側に列挙する
（器官を足すのは挙動変更＝ゴールデンが赤くなる。この回は純リファクタ）。
★ None は「忘れた」でなく「無いと確認した」という宣言。番人（missing_cells）が検査する
のは「キーが存在するか」であって、値の True/None ではない ―― 値が None であること自体は
何も壊さない。キーが無いことだけが「埋め忘れ」として赤くなる。

★ 置き場所: ailine_core/ 側にデータのみを置く（C4/C5 に倣う）。器官の実体（関数）は
ailine.py 側にあり、循環 import を避けるためここでは import しない・関数参照も持たない
（True/None の在/無だけの宣言表）。この表自体は ailine.py のどの経路からも読まれない
（=「在り」の値が実装と一致しているかは番人ではなく report 側の手動照合で担保する。
ailine.py の line budget（tests/ailine_py_line_budget.txt・上限に余地ゼロ）を尊重して
ailine.py 側は一切変更していない）。単体で `python -c "import ailine_core.stage_organs"`
が通る（import 依存ゼロ・標準ライブラリのみ）。
"""
from __future__ import annotations

# --- 段種 -------------------------------------------------------------------
# 実行経路の単位。cmd_run の分岐 + cmd_run_plan（複合計画）の段種別実行を実測して確定した。
# ★ 「dsl / freeform / clarify」の3種だけでは、この repo で実際に起きた欠陥
#   （単発と複合計画の段で器官の適用が食い違う）を表現できない ―― W10f/今回発見の穴は
#   どちらも「同じ段種の *single* 版にはあるが *plan_step* 版に無い」形だったため、
#   single/plan_step を別の段種として分けることが本命（3種に丸めると欠陥そのものが
#   表から消える）。
STAGES = (
    "dsl_single",          # 単発 DSL 経路（ailine.py cmd_run_dsl）
    "dsl_plan_step",        # 複合計画(cmd_run_plan)の DSL 語彙段
    "freeform_single",      # 単発自由生成（ailine.py cmd_run_freeform）
    "freeform_plan_step",   # 複合計画(cmd_run_plan)の語彙外(FREEFORM/OUT_OF_VOCAB)段
                             # （ailine.py run_freeform_plan_step。cmd_run_plan は両方を同じ経路で扱う）
    "clarify_single",       # 単発の計画が CLARIFY 1段だけになった場合（_cmd_run_dispatch）
    "clarify_plan_step",    # 複合計画の中の CLARIFY 段（cmd_run_plan）
)

# --- 器官 -------------------------------------------------------------------
# 「同じ知見を全段種へ機械的に伝播させたい道具」として実測された6種
# （ブリーフが例示した接地検証・破壊の関所・率リテラル走査・ヘルパ総なめ検出・助言・
# 切り詰め注記、そのままの粒度）。
ORGANS = (
    "grounding",            # 接地検証: verify_dsl_args による DSL 引数のブック実体との照合
    "destructive_gate",     # 破壊の関所: OP_WRITE_TARGET 宣言 → _maybe_warn_target_overwrite
                             # → _confirm_overwrite_or_gate（既存列上書きの検知+確認）
    "rate_scan",             # 率リテラル走査: scan_rate_literals（A' 原則＝LLM に率を確定
                             # させないことの機械監査）
    "helper_sweep_detect",   # ヘルパ総なめ検出: detect_helper_sweep + それを踏まえた
                             # _confirm_freeform_apply（自由生成の関所）
    "advisories",             # 助言: build_advisories 系（幽霊データ/一様埋め/件数突合/
                             # 新規シート申告/依頼文言との重なり）
    "truncation_notice",      # 切り詰め注記: _truncation_notice（snapshot() の MAX_ROWS
                             # 切り詰めを無言で切らず申告する）
)

# --- 宣言表 -------------------------------------------------------------------
# 値は True（在り）/ None（無いと確認した宣言）のみ。False は使わない
# （「無い」は None 一種類に統一 ―― 二重の否定表現を持たせない）。
STAGE_ORGANS = {
    "dsl_single": {
        "grounding": True,             # ailine.py cmd_run_dsl: verify_dsl_args 呼び出し
        "destructive_gate": True,      # ailine.py cmd_run_dsl: _maybe_warn_target_overwrite/_confirm_overwrite_or_gate
        # ★ DSL はルールベース codegen（codegen_dsl）であり LLM の自由生成コードを一切
        #   経由しない。「率らしい数値リテラルが紛れ込む」「無関係なヘルパを総なめする」は
        #   自由生成特有のリスクで、DSL には構造的に対象が無い。
        "rate_scan": None,
        "helper_sweep_detect": None,
        "advisories": True,            # ailine.py cmd_run_dsl: build_advisories 呼び出し
        "truncation_notice": True,     # ailine.py cmd_run_dsl: _truncation_notice 呼び出し
    },
    "dsl_plan_step": {
        "grounding": True,             # ailine.py cmd_run_plan（DSL 段）: verify_dsl_args 呼び出し
        "destructive_gate": True,      # ailine.py cmd_run_plan（DSL 段）: _maybe_warn_target_overwrite/_confirm_overwrite_or_gate
        "rate_scan": None,             # dsl_single と同じ理由（構造的に対象が無い）
        "helper_sweep_detect": None,   # 同上
        "advisories": True,            # ailine.py cmd_run_plan（DSL 段）: _structural_advisories 等（W10d で追加）
        # ★★ 発見（この回の本命）: _truncation_notice はこの段種では一度も呼ばれていない。
        #   _truncation_notice 自身の docstring（ailine.py）は「exhaustive_postcondition=True
        #   （DSL経路・cmd_run_dsl / cmd_run_plan の DSL 段）」と書いており、複合計画の DSL 段
        #   にも適用されている前提で書かれているが、実装(cmd_run_plan の DSL 分岐)はこの関数を
        #   一度も呼んでいない ―― コメントの主張と実装が食い違っている。DSL 経路は事後条件
        #   チェッカーが全行を検証するため安全性への影響は無いが、「表示が MAX_ROWS で
        #   切り詰められている」ことを利用者に伝える1行が、この段種だけ出ない。
        "truncation_notice": None,
    },
    "freeform_single": {
        # 自由生成は DSL args という構造化された概念を持たない（verify_dsl_args を呼ぶ
        # 対象が無い）。
        "grounding": None,
        # OP_WRITE_TARGET が守る「書き込み先列」という構造化された対象が無い。代わりに
        # 別の関所（_confirm_freeform_apply＝「機械検証できません。適用しますか？」）が
        # あるが、これは helper_sweep_detect 器官の一部として扱う（下記）。
        "destructive_gate": None,
        "rate_scan": True,             # ailine.py cmd_run_freeform: scan_rate_literals 呼び出し
        "helper_sweep_detect": True,   # ailine.py cmd_run_freeform: detect_helper_sweep + _confirm_freeform_apply
        "advisories": True,            # ailine.py cmd_run_freeform: build_advisories 呼び出し
        "truncation_notice": True,     # ailine.py cmd_run_freeform: _truncation_notice 呼び出し
    },
    "freeform_plan_step": {
        "grounding": None,             # freeform_single と同じ理由
        "destructive_gate": None,      # 同上
        "rate_scan": True,             # ailine.py run_freeform_plan_step: scan_rate_literals 呼び出し（W10f で追加）
        "helper_sweep_detect": True,   # ailine.py run_freeform_plan_step: detect_helper_sweep + _confirm_freeform_apply
        "advisories": True,            # ailine.py run_freeform_plan_step: build_advisories 呼び出し
        "truncation_notice": True,     # ailine.py run_freeform_plan_step: _truncation_notice 呼び出し
    },
    "clarify_single": {
        # ★ CLARIFY はコード生成も適用も一切発生しない段（_cmd_run_dispatch は質問を
        #   印字して exit 3 で終わる）。器官はどれも「適用されたものを検証/助言する」道具
        #   なので、適用そのものが起きないこの段種には構造的に対象が無い。
        #   ★ これは「見つかった穴」ではなく「対象が存在しないので宣言不要」という意味の
        #   None ―― dsl_plan_step の truncation_notice（対象は存在するのに無い）とは
        #   性質が違う。report 側で両者を区別して報告する。
        "grounding": None,
        "destructive_gate": None,
        "rate_scan": None,
        "helper_sweep_detect": None,
        "advisories": None,
        "truncation_notice": None,
    },
    "clarify_plan_step": {
        # 同上（cmd_run_plan 内の CLARIFY 分岐も適用を一切行わず、その段を fail 項目として
        # 記録するだけ）。
        "grounding": None,
        "destructive_gate": None,
        "rate_scan": None,
        "helper_sweep_detect": None,
        "advisories": None,
        "truncation_notice": None,
    },
}


def missing_cells(table: dict = STAGE_ORGANS, stages=STAGES, organs=ORGANS) -> list:
    """(段種, 器官) のうち、table にキーとして明示的な値が無いものを列挙する。

       ★ 番人の本体: 値が None であること自体は問題ない（＝「無い」という宣言）。
       キーそのものが無い（＝埋め忘れ）ことだけを検査する。段種を1つ足した人／器官を
       1つ足した人が、その行/列のどこか1マスでも埋め忘れたらここが空でなくなる
       （test_op_write_target_declares_all_ops の一般化）。
       戻り値: [(stage, organ), ...]（無ければ空リスト）。"""
    missing = []
    for stage in stages:
        row = table.get(stage)
        if row is None:
            missing.extend((stage, organ) for organ in organs)
            continue
        for organ in organs:
            if organ not in row:
                missing.append((stage, organ))
    return missing
