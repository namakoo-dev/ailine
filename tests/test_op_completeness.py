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
import json
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))

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
    return [REPO / "ailine.py"] + sorted((REPO / "ailine_core").glob("*.py"))


def discover_op_tables(min_hits: int = 3) -> dict:
    """ailine.py / ailine_core/*.py を ast で走査し、「dict リテラルを値に持つ単純代入
    （`X = {...}`）で、キーの文字列集合が OP_SCHEMA の op 名を min_hits 件以上含むもの」を
    全部見つける。戻り値: {"相対パス:変数名": {ヒットした op 名の集合}}。

    ★ 判定条件（誤検知対策・実測 2026-08-20 で確認した通りに絞ってある）:
      - 対象は ast.Assign かつ value が ast.Dict のみ。dict 内包表記(DictComp)は対象外
        ―― キーが文字列リテラルでなく式のため機械的に op 名の集合を取り出せない
        （例: `OP_LABELS = {op: meta["label"] for op, meta in OP_META.items()}` は
        キーが変数 op であり静的には拾えない。OP_LABELS は OP_META の派生物として
        FULL_DOMAIN_TABLES 側で個別に完全性を見る）。
      - キーは ast.Constant(str) のみ数える（f-string・変数キー等は無視）。
      - 代入先が単純な Name（`X = {...}`）のものだけを対象にする。
      - 変数名が "OP_SCHEMA" のものは対象外（op 名そのものを列挙する基底表であり、
        比較の基準そのものなので自分自身は目録に含めない）。
      - ast.walk でモジュール内の全ノード（関数内のローカル変数も含む）を走査するが、
        実測（2026-08-20・ailine.py + ailine_core/*.py 全ファイル）ではモジュール直下の
        8件（OP_SCHEMA を含む）以外に該当は無く、誤検知は出なかった。将来ローカル変数の
        辞書が誤検知したら、ここに絞り込み条件（例: 関数内を除外）を追記すること。
    """
    found = {}
    for path in _scan_target_files():
        tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
        rel = path.relative_to(REPO).as_posix()
        for node in ast.walk(tree):
            if not (isinstance(node, ast.Assign) and isinstance(node.value, ast.Dict)):
                continue
            names = [t.id for t in node.targets if isinstance(t, ast.Name)]
            if not names:
                continue
            keys = {k.value for k in node.value.keys
                    if isinstance(k, ast.Constant) and isinstance(k.value, str)}
            hit = keys & OP_SET
            if len(hit) < min_hits:
                continue
            for name in names:
                if name == "OP_SCHEMA":
                    continue
                found[f"{rel}:{name}"] = hit
    return found


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
