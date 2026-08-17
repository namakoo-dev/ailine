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
        # ★★ C6 で発見された穴（この段種だけ _truncation_notice が一度も呼ばれておらず、
        #   _truncation_notice 自身の docstring が主張する適用範囲と実装が食い違っていた）は
        #   ★ C9 で塞いだ（ailine.py _run_dsl_plan_step が step_before/step_after で呼ぶ）。
        #   塞いだ理由: ✓ の意味を「最終ファイルを読み戻して確かめた」に一本化する以上、
        #   「表示は先頭 MAX_ROWS 行しか見ていない」は ✓ の主張範囲に直接効くため。
        "truncation_notice": True,     # ailine.py _run_dsl_plan_step: _truncation_notice 呼び出し（C9 で追加）
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


# --- ★ C6b: 宣言を現実（ailine.py の実装）と突き合わせる番人 ---------------------------
# C6 の missing_cells() は表の**内的整合**（全マス目が埋まっているか）だけを見る。
# ここから下は、その先の欠陥（W10f の再演: 表が True のまま裏の呼び出しだけ消えても
# missing_cells は気づけない）を防ぐための、**表 × ailine.py の実 AST の突合せ**。
#
# ★ 段↔関数の対応（STAGE_ENTRY_FUNCTIONS）は「その段の実行経路を代表する ailine.py の
#   トップレベル関数」を宣言する。値は関数名のタプル（AST 走査はこの関数の本体
#   *だけ* を見る＝他関数の呼び出し先までは追わない・呼び出しヘルパ経由の間接呼び出しは
#   構造的に見えない）。
#   ★★ 空タプル `()` は「この段種は関数単位では他の段種の分岐と分離できない」という
#   明示的な unverifiable 宣言（missing ではない ―― キーは必ず埋める。C6 の None と同じ
#   思想: 「追えない」も「忘れた」と区別して機械で読めるようにする）。
#
# ★★★ C7（三経路の統合）で更新: dsl_single(cmd_run_dsl) と dsl_plan_step の DSL 段は
#   ailine_core.dsl_step の共有エンジン（print_dsl_confirmation/apply_dsl_step）を通る
#   ようになった。dsl_plan_step の実体（依存つき連鎖の新規列フォールバック含む）は
#   cmd_run_plan の for ループから独立した関数 _run_dsl_plan_step に切り出したため、
#   dsl_plan_step の代表関数を cmd_run_plan から _run_dsl_plan_step へ更新する。
#   ★ 副産物（DoD7）: この切り出しにより cmd_run_plan 自身の関数本体はもう DSL 段の
#   器官（verify_dsl_args/_maybe_warn_target_overwrite/_structural_advisories 等）を
#   直接呼ばなくなった。--dry プレビュー本体も同じ理由で _preview_dsl_plan へ切り出した
#   （cmd_run_plan がそこでも verify_dsl_args/format_confirmation_line を直接呼んでいた
#   ── これも解消しないと clarify_plan_step の誤検知の種になる）。結果、cmd_run_plan
#   自身の本体は CLARIFY 分岐(inline 3行) + run_freeform_plan_step/_run_dsl_plan_step/
#   _preview_dsl_plan への委譲 + ループ外の集計（mention_overlap_advisory 等・どの
#   ORGAN_FUNCTION_CANDIDATES にも含まれない）だけになり、以前のように dsl_plan_step の
#   呼び出しを誤って拾う心配が無くなった ―― clarify_plan_step を unverifiable から
#   卒業させ、cmd_run_plan を代表関数にできる（旧コメントが警告していた誤検知の種は
#   dsl_plan_step 分岐が別関数へ分かれたことで消えた）。
STAGE_ENTRY_FUNCTIONS = {
    "dsl_single": ("cmd_run_dsl",),
    "dsl_plan_step": ("_run_dsl_plan_step",),
    "freeform_single": ("cmd_run_freeform",),
    "freeform_plan_step": ("run_freeform_plan_step",),
    # ★ 単発の CLARIFY 分岐（plan が1段だけで CLARIFY の場合）は print+return の3行だけで、
    #   器官の呼び出しは一切無い。かつ同じ関数内の他の分岐（cmd_run_dsl/cmd_run_freeform/
    #   cmd_run_plan の呼び出し）は別関数への委譲であって AST には inline されていないため、
    #   関数単位の走査でも分岐混線が起きない（clarify_plan_step と違い、安全に検証できる）。
    #   ★★ 挙動変更#3 で更新: この分岐は _cmd_run_dispatch から _translate_and_dispatch へ
    #   移った（シート名の衝突の3択②で「翻訳からやり直す」ために、対象シート決定より後を
    #   別関数へ切り出したため）。代表関数を実体のある側へ追随させる ── 表は現実を写す。
    "clarify_single": ("_translate_and_dispatch",),
    # ★★ C7 で unverifiable を卒業（旧: 空タプル）。cmd_run_plan の CLARIFY 分岐（3行）
    #   と DSL 段の実体(_run_dsl_plan_step)が別関数に分かれたことで、cmd_run_plan 自身の
    #   AST を安全に代表関数として使えるようになった（上のコメント参照）。
    "clarify_plan_step": ("cmd_run_plan",),
}

# ★ 器官 → 「その器官が満たされたとみなす ailine.py 側の呼び出し先関数名」の候補集合。
#   1器官が複数関数のどれかで満たされる場合がある ―― どの候補が実際に呼ばれていても
#   「いずれかが呼ばれていれば満たす」で判定する。
#   ★ C7: grounding/destructive_gate/advisories に ailine_core.dsl_step の共有関数
#   （resolve_dsl_step_args/print_dsl_confirmation/compose_dsl_step_advisories）を追加。
#   cmd_run_dsl と _run_dsl_plan_step はどちらも同じ関数群を経由してこれらの器官を満たす。
#   dsl_single と dsl_plan_step はこの3器官をどちらも True と宣言しているため、同じ候補名を
#   共有しても True/None の判定はぶれない（両側とも「見つかれば True 相当」で一致する）。
#   ★★ truncation_notice には共有関数(apply_dsl_step)を候補に**加えていない**（意図的）:
#   apply_dsl_step は単発/複合計画の両方から呼ばれるが、_truncation_notice を実際に
#   呼ぶかどうかは呼び出し側(ailine.py)に残した（ailine_core/dsl_step.py の
#   apply_dsl_step docstring 参照）。候補を「呼び出し側に実在する関数名」1つに絞って
#   おくことで、この器官の在/無は各段の関数本体を見れば必ず判定できる
#   （AST の名前ベース走査は「関数の中身」までは追えないという、この番人の設計そのものの
#   限界 ―― モジュール先頭のコメント参照）。★ C9 で dsl_plan_step 側も True になった。
ORGAN_FUNCTION_CANDIDATES = {
    "grounding": ("verify_dsl_args", "resolve_dsl_step_args"),
    "destructive_gate": ("_maybe_warn_target_overwrite", "_confirm_overwrite_or_gate",
                          "print_dsl_confirmation"),
    "rate_scan": ("scan_rate_literals",),
    "helper_sweep_detect": ("detect_helper_sweep", "_confirm_freeform_apply"),
    "advisories": ("build_advisories", "_structural_advisories", "compose_dsl_step_advisories"),
    "truncation_notice": ("_truncation_notice",),
}


def reality_mismatches(found_names_by_stage: dict,
                        table: dict = STAGE_ORGANS,
                        entry_functions: dict = STAGE_ENTRY_FUNCTIONS,
                        organ_candidates: dict = ORGAN_FUNCTION_CANDIDATES,
                        organs=ORGANS) -> list:
    """★ 番人の本体（C6b）: 宣言表 (table) と「実際に呼ばれている器官の集合」
       (found_names_by_stage: {段種: その段の関数本体で見つかった呼び出し先関数名の集合}) を
       突き合わせ、食い違うマス目を列挙する。純ロジック（ailine.py を読む/AST を作るのは
       呼び出し側の責務 ―― C4/C5 に倣い、ここは ailine.py に一切依存しない）。

       食い違いは両方向:
       - declared_true_not_found: True と宣言したのに、その段の関数本体に器官の呼び出しが
         見当たらない（＝誰かが呼び出しを消した、または宣言だけ先に True にした）。
       - declared_none_but_found: None（無いという宣言）なのに、実際にはその段の関数本体で
         器官の呼び出しが見つかった（＝宣言と現実がずれている。無いと言ったのに在る）。

       entry_functions の値が空タプルの段種（unverifiable）はここでは判定しない
       （found_names_by_stage に何を渡しても無視される＝呼び出し側が空集合を渡しても
       誤検知しない）。
       戻り値: [(stage, organ, kind), ...]（kind は上の2種、無ければ空リスト）。"""
    mismatches = []
    for stage, funcs in entry_functions.items():
        if not funcs:   # unverifiable
            continue
        found = found_names_by_stage.get(stage, set())
        row = table.get(stage, {})
        for organ in organs:
            declared = row.get(organ)
            candidates = set(organ_candidates.get(organ, ()))
            is_found = bool(found & candidates)
            if declared is True and not is_found:
                mismatches.append((stage, organ, "declared_true_not_found"))
            elif declared is not True and is_found:
                mismatches.append((stage, organ, "declared_none_but_found"))
    return mismatches


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
