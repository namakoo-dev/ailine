"""README のヘルパ表 / ailine.py が LLM に見せるカタログの Call 例 / helpers/*.bas の実体、
   三者の呼び出し規約（引数の個数）が食い違っていないかを検査する回帰テスト。

   ★ 背景（全域監査 M1/M2 で発覚）: W3 で headerRow を全ヘルパ関数に通す変更をした際、
   ailine.py 側のカタログは追随したが README の表だけ古いまま残った。実測で
   SortByColumn / InsertBarChart / AlignCenter / FormatThousands / VLookupFromTable /
   SummaryTable の6件が該当（監査の「6箇所」と一致）。

   ★ 純ロジックのみ・LibreOffice/basrun 不要（CI で走る）。
   ★ 誤検知対策: README・ailine.py・helpers/*.bas それぞれの実際の書式に素直な緩い
   正規表現で抽出し、引数「個数」の一致だけを見る（README/カタログ側は `c1, r1` のような
   短縮した引数名を使うため、名前までは揃えない）。"""
import re
from pathlib import Path

import pytest

REPO_ROOT = Path(__file__).resolve().parent.parent
HELPERS_BAS = REPO_ROOT / "src" / "ailine" / "helpers" / "AiLineHelpers.bas"
README = REPO_ROOT / "README.md"

import sys
sys.path.insert(0, str(REPO_ROOT / "src"))
import ailine


def _arg_count(args: str) -> int:
    """引数リストの文字列（カンマ区切り）から個数を数える。空文字列は0引数。"""
    return 0 if args.strip() == "" else len(args.split(","))


def _parse_bas_signatures(text: str) -> dict:
    """`Sub 名(引数...)` を実体の正とみなし、名前→引数個数の辞書にする。"""
    sigs = {}
    for m in re.finditer(r"^Sub\s+([A-Za-z_]\w*)\s*\(([^)]*)\)", text, re.MULTILINE):
        sigs[m.group(1)] = _arg_count(m.group(2))
    return sigs


def _parse_readme_table(text: str) -> dict:
    """README のヘルパ表の1列目 `` `名前(引数...)` `` から名前→引数個数を拾う。"""
    sigs = {}
    for m in re.finditer(r"^\|\s*`([A-Za-z_]\w*)\(([^)]*)\)`\s*\|", text, re.MULTILINE):
        sigs[m.group(1)] = _arg_count(m.group(2))
    return sigs


def _parse_catalog_call_examples(catalog_text: str) -> dict:
    """ailine.py の load_helpers() が組む文字列のうち、手書きの説明文中にある
       `Call 名前(引数...)` の例から名前→引数個数を拾う。
       ★ カタログ後半は helpers/*.bas の原文をそのまま埋め込んだもの（コピペそのもので
       ズレようがない）なので、`--- 定義済み` マーカーより前の手書き部分だけを見る
       （そこを含めると内部専用の BoldRange 呼び出し等まで「公開ヘルパ」扱いしてしまう）。"""
    hand_written = catalog_text.split("--- 定義済み", 1)[0]
    sigs = {}
    for m in re.finditer(r"Call\s+([A-Za-z_]\w*)\(([^)]*)\)", hand_written):
        sigs.setdefault(m.group(1), _arg_count(m.group(2)))
    return sigs


def test_helper_signatures_stay_in_sync_across_readme_ailine_and_bas():
    """README のヘルパ表 / ailine.py カタログの Call 例 / helpers/*.bas の実体、三者の
       引数個数が一致することを検査する（helpers/*.bas の Sub 宣言を実体の正とする）。"""
    bas_sigs = _parse_bas_signatures(HELPERS_BAS.read_text(encoding="utf-8"))
    assert bas_sigs, "helpers/AiLineHelpers.bas から Sub 宣言が1件もパースできなかった"

    readme_sigs = _parse_readme_table(README.read_text(encoding="utf-8"))
    assert readme_sigs, "README のヘルパ表が1件もパースできなかった（表の書式が変わった？）"

    catalog, _ = ailine.load_helpers(REPO_ROOT / "src" / "ailine" / "helpers")
    catalog_sigs = _parse_catalog_call_examples(catalog)
    assert catalog_sigs, "ailine.py のカタログから Call 例が1件もパースできなかった"

    mismatches = []
    for name, count in readme_sigs.items():
        bas_count = bas_sigs.get(name)
        if bas_count is None:
            mismatches.append(f"README: `{name}` は helpers/*.bas に実体が無い")
        elif bas_count != count:
            mismatches.append(f"README: {name}({count}引数) != 実体({bas_count}引数)")
    for name, count in catalog_sigs.items():
        bas_count = bas_sigs.get(name)
        if bas_count is None:
            mismatches.append(f"ailine.py カタログ: `{name}` は helpers/*.bas に実体が無い")
        elif bas_count != count:
            mismatches.append(f"ailine.py カタログ: {name}({count}引数) != 実体({bas_count}引数)")

    assert not mismatches, (
        "README/カタログ と ヘルパ実体の引数個数が食い違っている:\n" + "\n".join(mismatches)
    )


def _split_bas_subs(bas_text: str) -> dict:
    """`Sub 名(...) ... End Sub` の中身を名前→本文（文字列）の辞書にする。"""
    subs = {}
    for m in re.finditer(r"Sub\s+([A-Za-z_]\w*)\s*\([^)]*\).*?End Sub", bas_text, re.S):
        name = re.match(r"Sub\s+([A-Za-z_]\w*)", m.group(0)).group(1)
        subs[name] = m.group(0)
    return subs


def test_bold_helpers_set_charweight_and_charweightasian_together():
    """太字を当てるヘルパ（セルに `.CharWeight = ...BOLD` を設定する Sub）は、必ず
       `.CharWeightAsian` も同じ Sub の中で設定していることを静的に検査する。

       ★ 日本語は CharWeight だけでは太字が効かず CharWeightAsian が必須という過去の
       苦労した経緯があり、「環境的に太字は無理」と一度誤断した実例もある。この行の
       無言の退行は重大インシデント級。
       ★ 対象はセルの太字（check_bold が見る対象）に限る。グラフタイトルの
       `oChart.Title.CharWeight`（InsertBarChart）はセルと別の見出しオブジェクトで
       check_bold の対象外のため、"Title" を含む行は除外する（この試験のスコープ外）。
       ★ 限界（正直に）: これは「両方セットする行が書いてある」ことしか保証しない。
       実際に xlsx へ効いているかどうかまでは見ない（openpyxl による静的な検体では
       CharWeightAsian を経由しないため測れない）。実際に basrun で当てて openpyxl で
       読み戻す確認は tests/test_bold_local.py（要 LibreOffice・@pytest.mark.local）が担う。"""
    bas_text = HELPERS_BAS.read_text(encoding="utf-8")
    subs = _split_bas_subs(bas_text)
    assert subs, "helpers/AiLineHelpers.bas から Sub が1件もパースできなかった"

    offenders = []
    for name, body in subs.items():
        lines = body.splitlines()
        sets_weight = any(re.search(r"\.CharWeight\s*=", ln) and "Title" not in ln
                           for ln in lines)
        sets_asian = any(re.search(r"\.CharWeightAsian\s*=", ln) for ln in lines)
        if sets_weight and not sets_asian:
            offenders.append(name)

    assert not offenders, (
        "セルの CharWeight は設定しているが CharWeightAsian を設定していないヘルパがある"
        f"（日本語で太字が効かない）: {offenders}"
    )
