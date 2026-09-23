"""op 完全性の番人（宣言駆動・免除簿つき）。

★ なぜ在るか: op を1つ足すとは1機能ではなく「翻訳(battery)×生成×検証(事後条件)×
安全(writes)×表示(8表)」の束を足すことである。実測(2026-08-19〜20)で
SET_COLUMN_VALUE は battery 検体0件・APPEND_TOTAL は1件（どちらも看板機能なのに翻訳
精度がほとんど測られていない）、CENTER_ALIGN は1 op 内で3項が食い違っていた ──
「足したのに束のどれかが無い」を、足した瞬間に赤にする。

免除は tests/op_completeness_register.json（免除簿）に宣言する。免除には reason（正直な
理由。分からなければ「未調査」）と unlock（解除条件）が必須。免除簿が空になったとき、
この番人は今の実装の穴の分だけ機械的に赤くなる（＝これが番人の発火試験そのもの）。

3つの仕事:
① 目録の同期（AST・discover_op_tables）: ailine.py / ailine_core/*.py を静的走査し、
   op 名を3つ以上キーに持つ dict リテラル代入を全部見つける。見つかった表が免除簿の
   registered_op_tables に無ければ赤（新しい表が出来たのに気づかれていない）。逆に
   registered にあるのに見つからなければ赤（表がリネーム/削除された＝目録の腐り）。
② op の完全性: OP_SCHEMA の全 op について、
   - 5つの全域表（OP_META/OP_LABELS/OP_WRITE_TARGET/OP_SUBJECT_SLOTS/_CONFIRM_FIELDS）に
     在ること（免除不可）
   - POSTCONDITIONS に在ること、または免除簿に {op, "postcondition"} の宣言
   - ailine_core.target_sheet._OP_VERBS に在ること、または {op, "op_verbs"} の宣言
   - bench/translation_battery.json にその op を expected とする検体が2件以上、
     または {op, "battery>=2"} の宣言
③ 免除の腐り防止（test_stale_exemptions_are_removed 他）: 免除簿の各エントリについて、
   実体が既に要求を満たしていれば赤（「免除を消せ」）。存在しない op への免除も赤。
   reason/unlock が空のエントリも赤。
"""
import ast
import inspect
import json
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))

import ailine  # noqa: E402
from ailine_core import target_sheet  # noqa: E402

REGISTER_PATH = Path(__file__).resolve().parent / "op_completeness_register.json"
BATTERY_PATH = REPO / "bench" / "translation_battery.json"

OP_SET = set(ailine.OP_SCHEMA.keys())
REQUIREMENT_KINDS = {"postcondition", "op_verbs", "battery>=2"}

# ★ 免除不可の全域表。OP_LABELS は OP_META から機械生成される派生物（dict 内包表記）
#   なので discover_op_tables の AST 走査（キーが文字列リテラルの dict リテラルのみ対象）
#   では見つからない。それでも「全 op を持つべき表」であることは変わらないため、
#   ①（目録の同期）とは別にここへ明示的に列挙して②（完全性）の対象にする。
FULL_DOMAIN_TABLES = {
    "OP_META": lambda: ailine.OP_META,
    "OP_LABELS": lambda: ailine.OP_LABELS,
    "OP_WRITE_TARGET": lambda: ailine.OP_WRITE_TARGET,
    "OP_SUBJECT_SLOTS": lambda: ailine.OP_SUBJECT_SLOTS,
    "_CONFIRM_FIELDS": lambda: ailine._CONFIRM_FIELDS,
}


def _load_register() -> dict:
    return json.loads(REGISTER_PATH.read_text(encoding="utf-8"))


def _exemptions_by_requirement(register: dict, requirement: str) -> set:
    return {e["op"] for e in register["exemptions"] if e.get("requirement") == requirement}


def _scan_target_files() -> list:
    """★ 2026-09-23: 手で並べず芯から引く（再帰する ── 初版の `glob("*.py")` は
    ailine_core/postconditions/ を見ていなかった）。視野は元と同じ src の下。"""
    from _product_source import src_files
    return src_files()


def discover_op_rosters(min_hits: int = 1) -> dict:
    """ailine.py / ailine_core/*.py を ast で走査し、**op 名を並べた名簿**を全部見つける。
    戻り値: {"相対パス:変数名": {"kind": 種別, "ops": ヒットした op 名の集合}}。

    ★★ 2026-09-16 の掃き出しでここを広げた。それまでは「dict リテラルで op 3 つ以上」
      だけを見ていて、**23 件の名簿のうち 9 件が不可視**だった:
        frozenset PLAN_CHAIN_CONSUMER_OPS(12) / _OPS_THAT_SKIP_NON_DATA_ROWS(3) /
        ROW_REDUCING_OPS(3)、tuple DEFAULT_SUGGESTIONS(3) / PLAN_CHAIN_WARNING_OPS(2) /
        KEEP_FOR_COLUMN_REQUEST(2)、dict OP_DECLARED_SHEET_NAME(2) /
        PLAN_CHAIN_UNNAMED_PRODUCERS(2) / MACHINE_DERIVED_ARGS(1)。
      ★ 見えない名簿は「部分なのは意図か」を誰にも問われない ── そこに片配線が溜まる。
        実際 _OPS_THAT_SKIP_NON_DATA_ROWS は AGGREGATE を落としていて、
        「合計行を除いて集計して」が利用者の合計行を**実際に消していた**。
      ★ 種別（dict/set/frozenset/tuple/list）と個数の下限（1）を広げただけで、
        判定条件そのものは前と同じ ── キーは文字列リテラルのみ・代入先は単純な Name のみ。

    ★ 判定条件（誤検知対策・実測どおりに絞ってある）:
      - 対象は ast.Assign。値は Dict / Set / List / Tuple / frozenset(...) / set(...)。
        dict 内包表記(DictComp)は対象外 ―― キーが式のため静的に op 名を取れない
        （例: `OP_LABELS = {op: meta["label"] ...}`。OP_LABELS は FULL_DOMAIN_TABLES 側で見る）。
      - 要素/キーは ast.Constant(str) のみ数える（f-string・変数キー等は無視）。
      - 変数名が "OP_SCHEMA" のものは対象外（比較の基準そのもの）。
    """
    found = {}
    for path in _scan_target_files():
        tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
        rel = path.relative_to(REPO).as_posix()
        for node in ast.walk(tree):
            if not isinstance(node, ast.Assign):
                continue
            names = [t.id for t in node.targets if isinstance(t, ast.Name)]
            if not names:
                continue
            value, kind, items = node.value, None, []
            if isinstance(value, ast.Dict):
                kind, items = "dict", list(value.keys)
            elif isinstance(value, (ast.Set, ast.List, ast.Tuple)):
                kind, items = type(value).__name__.lower(), list(value.elts)
            elif (isinstance(value, ast.Call) and isinstance(value.func, ast.Name)
                    and value.func.id in ("frozenset", "set")):
                kind = value.func.id
                for arg in value.args:
                    if isinstance(arg, (ast.Set, ast.List, ast.Tuple)):
                        items += list(arg.elts)
            if kind is None:
                continue
            keys = {n.value for n in items
                    if isinstance(n, ast.Constant) and isinstance(n.value, str)}
            hit = keys & OP_SET
            if len(hit) < min_hits:
                continue
            for name in names:
                if name == "OP_SCHEMA":
                    continue
                found[f"{rel}:{name}"] = {"kind": kind, "ops": hit}
    return found


def discover_op_tables(min_hits: int = 3) -> dict:
    """★ 旧名（dict・op 3 つ以上だけ）。①の目録が守ってきた範囲をそのまま残す。"""
    return {k: v["ops"] for k, v in discover_op_rosters(min_hits=min_hits).items()
            if v["kind"] == "dict" and len(v["ops"]) >= min_hits}


def battery_op_counts() -> dict:
    """bench/translation_battery.json の items* リスト群から、op ごとの検体数を数える。

    ★ 数え方（構造で読む。substring 数えは過大/過小になるため使わない）:
      - 検体の "expect" が dict で "op" キーを持てば、その op を1件と数える
        （tests/... の items/items_v4/items_v5 など単発検体の形）。
      - 検体が "expect_plan"（複合計画・items_v2）を持てば、中の各 step の "op" を
        1件ずつ数える（複合計画の1段も、その op の翻訳が実際に測られている証拠になる。
        例: items_v2 id=102 の [COMPUTE_COLUMN, APPEND_TOTAL] は両方に+1）。
      - "expect" が文字列（"clarify"/"freeform"。items の一部・items_v3 の "clarify" 系）
        の検体、または上記どちらの形にも op を含まない検体は、どの op にもカウントしない
        （CLARIFY/FREEFORM 経路の検体であり、特定 op の翻訳精度は測っていないため）。
    """
    data = json.loads(BATTERY_PATH.read_text(encoding="utf-8"))
    counts = {op: 0 for op in OP_SET}
    for value in data.values():
        if not isinstance(value, list):
            continue
        for item in value:
            if not isinstance(item, dict):
                continue
            plan = item.get("expect_plan")
            if isinstance(plan, list):
                for step in plan:
                    if isinstance(step, dict) and step.get("op") in counts:
                        counts[step["op"]] += 1
                continue
            expect = item.get("expect")
            if isinstance(expect, dict) and expect.get("op") in counts:
                counts[expect["op"]] += 1
    return counts


# ---------------------------------------------------------------------------
# ① 目録の同期
# ---------------------------------------------------------------------------

def test_discovered_op_tables_are_registered():
    register = _load_register()
    registered = set(register["registered_op_tables"])
    discovered = set(discover_op_tables().keys())
    unregistered = discovered - registered
    assert not unregistered, (
        "AST が見つけた op 表が免除簿の registered_op_tables に無い"
        "（新しい表が出来たのに気づかれていない）: "
        f"{sorted(unregistered)}。tests/op_completeness_register.json の"
        " registered_op_tables に追記すること。"
    )


def test_registered_op_tables_still_exist():
    register = _load_register()
    registered = set(register["registered_op_tables"])
    discovered = set(discover_op_tables().keys())
    ghosts = registered - discovered
    assert not ghosts, (
        "免除簿の registered_op_tables に載っているが AST で見つからなくなった表がある"
        f"（目録が腐っている＝表がリネーム/削除された）: {sorted(ghosts)}。"
        " tests/op_completeness_register.json を実体に合わせて更新すること。"
    )


# ---------------------------------------------------------------------------
# ② op の完全性
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("table_name", sorted(FULL_DOMAIN_TABLES))
def test_full_domain_table_covers_every_op(table_name):
    """免除不可。5つの全域表は OP_SCHEMA の全 op を持つこと。"""
    table = FULL_DOMAIN_TABLES[table_name]()
    missing = sorted(op for op in OP_SET if op not in table)
    assert not missing, f"{table_name} に無い op（免除不可・必ず埋めること）: {missing}"


def test_every_op_has_postcondition_or_declared_exemption():
    register = _load_register()
    exempt = _exemptions_by_requirement(register, "postcondition")
    missing = sorted(op for op in OP_SET
                      if op not in ailine.POSTCONDITIONS and op not in exempt)
    assert not missing, (
        "POSTCONDITIONS に無く、免除簿にも {requirement: postcondition} の宣言が無い op"
        f"（事後条件が検証されないまま足された可能性）: {missing}"
    )


def test_every_op_has_op_verb_or_declared_exemption():
    register = _load_register()
    exempt = _exemptions_by_requirement(register, "op_verbs")
    missing = sorted(op for op in OP_SET
                      if op not in target_sheet._OP_VERBS and op not in exempt)
    assert not missing, (
        "ailine_core.target_sheet._OP_VERBS に無く、免除簿にも"
        " {requirement: op_verbs} の宣言が無い op: "
        f"{missing}（自動フォールバック(ラベル+する/した)が自然か未確認のまま）"
    )


def test_every_op_has_battery_coverage_or_declared_exemption():
    register = _load_register()
    exempt = _exemptions_by_requirement(register, "battery>=2")
    counts = battery_op_counts()
    missing = sorted(op for op in OP_SET if counts.get(op, 0) < 2 and op not in exempt)
    assert not missing, (
        "battery の検体が2件未満で、免除簿にも {requirement: battery>=2} の宣言が無い op: "
        + ", ".join(f"{op}({counts.get(op, 0)}件)" for op in missing)
    )


# ---------------------------------------------------------------------------
# ③ 免除の腐り防止
# ---------------------------------------------------------------------------

def test_exemptions_reference_real_ops():
    register = _load_register()
    unknown = sorted({e["op"] for e in register["exemptions"] if e["op"] not in OP_SET})
    assert not unknown, f"OP_SCHEMA に存在しない op への免除がある: {unknown}"


def test_exemptions_have_known_requirement_kind():
    register = _load_register()
    bad = sorted({e.get("requirement") for e in register["exemptions"]
                  if e.get("requirement") not in REQUIREMENT_KINDS})
    assert not bad, f"未知の requirement 種別を持つ免除がある: {bad}（許される種別: {sorted(REQUIREMENT_KINDS)}）"


def test_exemptions_have_reason_and_unlock():
    register = _load_register()
    bad = [f"{e.get('op')}/{e.get('requirement')}" for e in register["exemptions"]
           if not str(e.get("reason", "")).strip() or not str(e.get("unlock", "")).strip()]
    assert not bad, f"reason または unlock が空の免除エントリがある: {bad}"


def test_stale_exemptions_are_removed():
    """免除の理由が既に解消され、実体が要求を満たしているのに免除簿に残っているものを検出する。"""
    register = _load_register()
    counts = battery_op_counts()
    stale = []
    for e in register["exemptions"]:
        op, req = e.get("op"), e.get("requirement")
        if op not in OP_SET:
            continue   # ← test_exemptions_reference_real_ops が別途落とす
        if req == "postcondition" and op in ailine.POSTCONDITIONS:
            stale.append(e)
        elif req == "op_verbs" and op in target_sheet._OP_VERBS:
            stale.append(e)
        elif req == "battery>=2" and counts.get(op, 0) >= 2:
            stale.append(e)
    assert not stale, (
        "実体が既に要求を満たしているのに免除簿に残っている免除がある（免除を消せ）: "
        + ", ".join(f"{e['op']}/{e['requirement']}" for e in stale)
    )


# ---------------------------------------------------------------------------
# ④ 被覆の宣言（2026-09-16 の掃き出しで新設）
#
# ★ なぜ在るか: op の名簿が**部分**であること自体は正しい場合が多い。悪いのは
#   「部分なのは意図か、それとも足し忘れか」を誰も問わないまま置かれることである。
#   実測: 23 件の名簿のうち 9 件は番人にそもそも見えておらず、その中の
#   _OPS_THAT_SKIP_NON_DATA_ROWS は AGGREGATE を落としていた（合計行が消える実害）。
#   ★ だから「全域」か「部分＋理由＋解除条件」を**全名簿に宣言させる**。
#     新しい名簿を足した人は、必ずどちらかを書く ── 書かなければ赤くなる。
# ---------------------------------------------------------------------------

def _load_coverage(register: dict) -> dict:
    return register.get("op_roster_coverage", {})


def test_every_op_roster_declares_its_coverage():
    """見つけた名簿は全部、被覆の宣言を持つこと（新しい名簿を足した瞬間に赤くなる）。"""
    register = _load_register()
    declared = set(_load_coverage(register))
    discovered = set(discover_op_rosters())
    missing = sorted(discovered - declared)
    assert not missing, (
        "op の名簿が見つかったのに op_roster_coverage に宣言が無い"
        f"（部分被覆が意図か足し忘れかを誰も問うていない）: {missing}。"
        " tests/op_completeness_register.json の op_roster_coverage に"
        ' {"coverage": "full"} か {"coverage": "partial", "reason": ..., "unlock": ...} を書くこと。'
    )


def test_coverage_declarations_are_not_stale():
    """宣言に在るのに実体が見つからない名簿（リネーム/削除）を赤にする。"""
    register = _load_register()
    ghosts = sorted(set(_load_coverage(register)) - set(discover_op_rosters()))
    assert not ghosts, f"op_roster_coverage に在るが実体が見つからない名簿: {ghosts}"


def test_full_coverage_declarations_are_true():
    """「全域」と宣言した名簿は、本当に全 op を持つこと。"""
    rosters = discover_op_rosters()
    bad = {}
    for name, decl in _load_coverage(_load_register()).items():
        if decl.get("coverage") != "full" or name not in rosters:
            continue
        miss = sorted(OP_SET - rosters[name]["ops"])
        if miss:
            bad[name] = miss
    assert not bad, f"「全域」と宣言したのに欠けている op がある: {bad}"


def test_partial_coverage_declarations_have_reason_and_unlock():
    """「部分」と宣言した名簿は、正直な理由と解除条件を書くこと（免除簿と同じ作法）。"""
    bad = [name for name, d in _load_coverage(_load_register()).items()
           if d.get("coverage") == "partial"
           and (not str(d.get("reason", "")).strip() or not str(d.get("unlock", "")).strip())]
    assert not bad, f"「部分」の宣言に reason または unlock が無い: {sorted(bad)}"


def test_coverage_kind_is_known():
    known = {"full", "partial"}
    bad = sorted({d.get("coverage") for d in _load_coverage(_load_register()).values()
                  if d.get("coverage") not in known})
    assert not bad, f"未知の coverage 種別: {bad}（許される: {sorted(known)}）"


def test_a_partial_roster_that_became_full_must_be_redeclared():
    """「部分」と宣言した名簿が全 op を覆うようになったら、宣言を直させる（腐り防止）。"""
    rosters = discover_op_rosters()
    stale = [name for name, d in _load_coverage(_load_register()).items()
             if d.get("coverage") == "partial" and name in rosters
             and not (OP_SET - rosters[name]["ops"])]
    assert not stale, f"「部分」と宣言したが実体は全 op を覆っている（宣言を full に直せ）: {sorted(stale)}"


# ---------------------------------------------------------------------------
# ⑤ 「自分で外す op」の名簿を**実装から導いて等号で縛る**
#
# ★ 2026-09-16: 手書きの名簿 _OPS_THAT_SKIP_NON_DATA_ROWS が AGGREGATE を落としており、
#   「合計行を除いて部署別に集計して」が、先頭の行削除の段を落とせずに
#   **利用者の台帳から合計行を実際に消していた**（同じ言い方でも並べ替え・抽出では消えない）。
#   ★ 名簿を目で読んでも気づけない ── 実装と突き合わせて初めて出た。だから機械に縛らせる。
# ★ 「自分で外す」の実装は 2 通りある。どちらも数える（片方だけ見ると SORT を見落とす）:
#     _skip_rows を生成関数に渡す … AGGREGATE / EXTRACT / SET_WHERE
#     _sort_end_row で末尾を切る  … SORT（別の腕 SortByColumnUpTo を呼ぶ）
# ---------------------------------------------------------------------------

SKIP_MARKERS = ("_skip_rows", "_sort_end_row")


def ops_that_skip_non_data_rows_derived_from_codegen() -> set:
    """生成関数の本体を**名前から**引いて、自分で非データ行を外す op を導く。

    tests/test_guard_ledger.py の台帳: 本体を**場所で決め打ちしない**
    （ファイルを分割すると空振りする）。inspect.getsource で関数から引く。
    """
    out = set()
    for op, fn in ailine.CODEGEN_BY_OP.items():
        try:
            body = inspect.getsource(fn)
        except (OSError, TypeError):
            continue
        if any(m in body for m in SKIP_MARKERS):
            out.add(op)
    return out


def test_skip_roster_equals_what_the_codegen_actually_does():
    declared = set(ailine._OPS_THAT_SKIP_NON_DATA_ROWS)
    derived = ops_that_skip_non_data_rows_derived_from_codegen()
    assert declared == derived, (
        "名簿 _OPS_THAT_SKIP_NON_DATA_ROWS が実装と食い違っている。"
        f"名簿に無いのに実際は外している: {sorted(derived - declared)} / "
        f"名簿に在るのに実装が外していない: {sorted(declared - derived)}。"
        "★ 名簿に無いと「合計行を除いて…」の行削除の段が落ちず、合計行が実際に消える。"
    )
