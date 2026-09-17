# -*- coding: utf-8 -*-
"""配線盤の中身 ── 経路 × 判断のマス目を実体から作る（2026-09-16）。

★★ なぜ tests/ に在るか: 素の環境の番人（scripts/_ci_parity_blocker.py）は
  「requirements-dev.txt に無い import」を全部止める。scripts/ に置いて番人から
  import すると、**自分の repo の道具なのに弾かれる** ── 2026-09-16 に
  refresh_records で実際に踏み、push が落ちた。だから中身はここ、
  scripts/wiring_board.py は薄い入口だけにする。

★ なぜ行列で、節と辺の図でないか（Namakoo「最適な図はある？」）:
  この repo の欠陥はほぼ 1 つの形をしている ──
  **同じ判断が、或る経路には在って或る経路には無い**。
  依存の図は「何が何に依存するか」を見るのに向くが、「どこに無いか」は映らない。

★★ マスの値を「宣言されているか」だけにしない。**三項**にする ──
  宣言（名簿に在るか）／実測（実装から導いた集合に在るか）／その食い違い。
  2026-09-16 のバグ 3 件はどれも「宣言どうし」ではなく**宣言と実装**のずれで、
  宣言しか見ない盤面では 1 件も映らなかった（試作の陽性対照が 0/3 で落ちた）。

★★ 列ごとに「何に守られているか」を 5 つに分ける（2026-09-17）:
  導出あり / 番人あり / 導けない（理由つき）/ ★ 未調査 / ★ 無防備。
  それまでは「導出が在るか」の 2 値で、**別の番人が縛っている列まで黄色**にしていた ──
  18 列中 14 列を実態より悲観的に見せ、読む人に同じ調査を繰り返させる形だった。
  ★ 第 3 の色を入れた直後に「無防備 0」と出たが、それは理由欄を埋めただけ。
    「未調査」を調べ終えた列と同じ色にするのは嘘なので、別の色に割った。
"""
from __future__ import annotations

import argparse
import inspect
import json
import re
import subprocess
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
SRC = REPO / "src"
sys.path.insert(0, str(SRC))
sys.path.insert(0, str(REPO / "tests"))

import ailine  # noqa: E402
from test_op_completeness import discover_op_rosters  # noqa: E402
from test_column_arg_keys_is_derived import derived as _column_arg_keys_derived  # noqa: E402

BAS_PATH = SRC / "ailine" / "helpers" / "AiLineHelpers.bas"
SKIP_MARKERS = ("_skip_rows", "_sort_end_row")


def _codegen_source(op: str) -> str:
    try:
        return inspect.getsource(ailine.CODEGEN_BY_OP.get(op))
    except (OSError, TypeError):
        return ""


def derivations(ops, codegen) -> dict:
    """実装から名簿を**導き直す**方法を持つ列だけを返す（名前 → (集合, 導き方の説明)）。

    ★ ここに 1 行足すたび、盤面の黄色い列が 1 本減る。増やすのがこの道具の使い道。
    ★ 導いた集合は「そうであるはず」の**候補**であって判定ではない ──
      食い違いは「要確認」と出す（勝手にバグと呼ばない）。
    """
    return {
        "_OPS_THAT_SKIP_NON_DATA_ROWS": (
            {op for op in ops if any(m in codegen[op] for m in SKIP_MARKERS)},
            "生成関数が _skip_rows / _sort_end_row を渡すか"),
        "PROJECTIONS": (
            {op for op in ops if ailine._op_writes(op, ailine.WRITE_NEW_SHEET)} | {"LOOKUP_FILL"},
            "OP_WRITE_TARGET が「新しいシートを作る」と宣言するか（＋その場で列を埋める LOOKUP_FILL）"),
        "STRUCTURAL_EFFECT": (
            {op for op in ops
             if any(ailine._op_writes(op, w) for w in (ailine.WRITE_NEW_COLUMN, ailine.WRITE_REMOVE,
                                                        ailine.WRITE_NEW_ROW_AT_END,
                                                        ailine.WRITE_ROW_SHIFT))},
            "OP_WRITE_TARGET の writes が行や列を増減させるか"),
        "_COLUMN_ARG_KEYS": (
            # ★ 正本は tests/test_column_arg_keys_is_derived.py の derived()。
            #   ここに書き写すと導出が 2 つになる ── 2026-09-16 に実際そうなり、
            #   盤が「食い違い 1 列」と嘘を表示した。**呼ぶ**こと。
            set(_column_arg_keys_derived()),
            "列のスロットのうち書き込み先を外し、削除する op も外す（＋CHART の例外）"),
    }


#: ★ 導出が作れない列の、**理由**（作らないのか、作れないのか）。
#:   ここに書けない列は「まだ調べていない」として盤に出る ── 黙って白にしない。
NO_DERIVATION_REASON = {
    "EXAMPLE_TASKS": "例が通るかは**実機を回さないと決まらない**"
                     "（番人 test_the_example_actually_runs が 1 例につき実走行 1 本を打つ）",
    "DEFAULT_SUGGESTIONS": "曖昧な依頼に見せる**表示用の並び**。上限 3 で、"
                           "実装から導ける性質のものではない",
    "PLAN_CHAIN_WARNING_OPS": "★ 読み手が消えた**死んだ名前**。導出ではなく削除が答え",
    "_OP_VERBS": "日本語の**活用形**の手書き。無い op はラベル+する/した に落ちる"
                 "（免除簿に 5 op 宣言済み）",
    "POSTCONDITIONS": "op → チェッカーの辞書そのもの。導出元が存在しない"
                      "（全 op か免除簿か、を番人が縛る）",
    "AXES": "「その op の軸に**まだ無い兄弟**が在るか」── 実装からは"
            "「無い」ことを導けない（実需が来たら足す閉集合）",
    "REDUCING_EXAMPLES": "ROW_REDUCING_OPS の 3 op に対する文例。等号の番人が縛る",
    "ROW_REDUCING_OPS": "盲検 9/9 の実測で**狭くした**線。実装からは導けない",
    "KEEP_FOR_COLUMN_REQUEST": "「中身のある列を作る」op の allowlist。"
                               "除外側を数えると op が増えるたび穴が開くため手書き",
    "MACHINE_DERIVED_ARGS": "★ 未調査 ── EXTRACT を入れた理由は書かれているが、"
                            "他の op を入れない理由が書かれていない",
    "OP_DECLARED_SHEET_NAME": "★ 未調査 ── helpers が固定名で書き出す op の対応表。"
                              "helpers 側から導ける可能性は測っていない",
    "_OP_SCHEMA_NOTES": "★ 未調査 ── 第二段翻訳に 1 行足す op の選び方が書かれていない",
    "PLAN_CHAIN_CONSUMER_OPS": "★ 未調査 ── 「絞り込んだ結果に掛けたいか元表か」は"
                               "機械に決まらない、という理由だけが書かれている",
}


def watchers_of(name: str) -> list:
    """その名簿を **assert の中で**名指ししている試験（ファイル:関数名）。

    ★ 言及だけ（コメント・docstring）は数えない ── 見ているとは言えないから。
    ★ これが在る列は「導出は無いが**無防備ではない**」。盤はそれを別の色で出す
      （2026-09-17: それまで一律に黄色で、18 列中 14 列を実態より悲観的に見せていた）。
    """
    import ast as _ast
    import re as _re
    out = []
    for p in sorted((REPO / "tests").rglob("*.py")):
        try:
            src = p.read_text(encoding="utf-8", errors="replace")
            tree = _ast.parse(src)
        except SyntaxError:
            continue
        if name not in src:
            continue
        for n in _ast.walk(tree):
            if not (isinstance(n, _ast.FunctionDef) and n.name.startswith("test_")):
                continue
            body = _ast.get_source_segment(src, n) or ""
            if name not in body:
                continue
            if any(_re.match(r"\s*assert\b", ln) for ln in body.splitlines()):
                out.append(f"{p.name}:{n.name}")
    return sorted(set(out))


def survey() -> dict:
    ops = sorted(ailine.OP_SCHEMA)
    n = len(ops)
    codegen = {op: _codegen_source(op) for op in ops}
    bas = BAS_PATH.read_text(encoding="utf-8", errors="replace")
    arms = set(re.findall(r"^\s*(?:Sub|Function)\s+(\w+)", bas, re.M))

    rosters = {k.split(":")[-1]: sorted(v["ops"])
               for k, v in discover_op_rosters().items() if len(v["ops"]) < n}
    der = derivations(ops, codegen)

    columns = []
    for name in sorted(rosters):
        d = der.get(name)
        w = watchers_of(name)
        unstudied = "未調査" in NO_DERIVATION_REASON.get(name, "")
        if d:
            stance = "derived"          # 実装から導き直せる
        elif unstudied:
            # ★★ 2026-09-17: 番人の有無より**理由の側**を優先する。
            #   MACHINE_DERIVED_ARGS は試験が 1 本在るが、中身は
            #   「EXTRACT に value が在る」の 1 行だけで、名簿が正しいかは見ていない。
            #   数だけ見て「番人あり」に吸うと、★ 未調査が盤から消える（実際そうなった）。
            #   ★ 番人が在ることと、その番人がその列を守っていることは別。
            stance = "unstudied"
        elif w:
            stance = "watched"          # 導出は無いが、別の番人が縛っている
        elif name in NO_DERIVATION_REASON:
            # ★★ 理由欄を埋めただけで安全になったことにしない。
            #   「★ 未調査」と書いた列は**調べ終えた列と別の色**にする ──
            #   でないと、盤を見た人が「全部片付いている」と読む（2026-09-17 に踏みかけた）。
            stance = ("unstudied" if "未調査" in NO_DERIVATION_REASON[name]
                      else "explained")
        else:
            stance = "bare"             # ★ 理由すら書かれていない
        columns.append({
            "name": name,
            "declared": rosters[name],
            "derived": sorted(d[0]) if d else None,
            "how": d[1] if d else "",
            "stance": stance,
            "watchers": w,
            "why_no_derivation": NO_DERIVATION_REASON.get(name, ""),
        })

    from ailine_core import stage_organs as so
    stage_panel = {
        "stages": list(so.STAGES),
        "organs": list(so.ORGANS),
        "cells": {s: {o: bool(so.STAGE_ORGANS[s].get(o)) for o in so.ORGANS} for s in so.STAGES},
    }

    reg = REPO / "tests" / "unwired_register.json"
    unwired = json.loads(reg.read_bytes().decode("utf-8"))["declared"] if reg.exists() else {}

    return {
        "commit": subprocess.run(["git", "log", "-1", "--format=%h %ad %s", "--date=short"],
                                  cwd=str(REPO), capture_output=True, text=True,
                                  encoding="utf-8").stdout.strip(),
        "ops": ops,
        "columns": columns,
        "codegen": {op: {"fn": getattr(ailine.CODEGEN_BY_OP.get(op), "__name__", ""),
                          "arms": sorted({m for m in re.findall(r"Call\s+(\w+)\s*\(", codegen[op])
                                          if m in arms}),
                          "post": getattr(ailine.POSTCONDITIONS.get(op), "__name__", "")}
                     for op in ops},
        "stage_panel": stage_panel,
        "unwired": unwired,
    }


def mismatches(data: dict) -> list:
    """宣言と実測がずれている列（名前 → 宣言だけ / 実測だけ）。"""
    out = []
    for c in data["columns"]:
        if c["derived"] is None:
            continue
        d, r = set(c["declared"]), set(c["derived"])
        if d != r:
            out.append({"column": c["name"], "declared_only": sorted(d - r),
                         "derived_only": sorted(r - d)})
    return out
