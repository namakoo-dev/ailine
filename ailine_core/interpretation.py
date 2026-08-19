"""interpretation — architect の上流レビュー「段1」: 解釈を機械可読で外に出す。

★ 症状: `--json` も history.jsonl も、解釈をレンダリング済みの日本語1文字列
（`command`）でしか吐いていない。`resolved`（解決値）と `inferred`（推定フラグ）が
外に出ていないため、「一度正しく読めた解釈を凍結して使い回す」ができない。

★ ここでやるのは「束ねて出す」だけ（新しい語彙も新しい判定も無い）:
  - スロットの形は ailine_core/subject.py の Slot がすでに持っている
  - 出所の3段階(matched/unspoken/contradicted)は SubjectVerdict.tier がすでに持っている
  - kind の種別(column/region/row/sheet/label)は subject.py がすでに持っている
  - inferred/_sources（vocab・task_literal の材料）は ailine.py が呼び出し時に渡す

★ `provenance` はここから導く派生ビュー（単位C の教訓: 出所を運ぶ場所が2つあると
片方だけ更新されて食い違う）。`resolved.get("_sources")` を読むのはこの関数の中
1箇所だけにし、ailine.py 側の2箇所（単発 cmd_run_dsl / 複合計画 _run_dsl_plan_step）
はどちらもこの戻り値を使う。★ 出力される値・型は今までと完全に同じ
（`resolved.get("_sources")` をそのまま返すだけ・golden の `provenance: null` は変わらない）。

★ 呼ぶのは DSL 経路（単発・複合計画の段）だけ。自由生成経路には `resolved` が
1箇所も存在しないため、この関数を呼ばない＝`interpretation` キーごと省略される
（`"interpretation": null` は「空の解釈がある」と読めるので出さない）。
★ ここは純ロジック（ファイルを開かない・ailine を import しない）。
"""
from __future__ import annotations

_VOCAB_PREFIX = "用語集:"
_TASK_LITERAL_PREFIX = "依頼文:"


def _slot_origin(key: str, tier_by_key: dict, inferred, sources: dict) -> str:
    """★ 語の優先順位（決定事項③・新しい語を発明しない）:
       対象スロットの出所(matched/unspoken/contradicted) → 推定(inferred) →
       出典(_sources の "用語集: X"/"依頼文: X") → それ以外は default。
       対象スロットとして判定された(①②③)キーは、常にその判定を優先する
       （例: SORT の col が推定でもあり①でもあれば①を名乗る ―― 依頼文の語と
       機械照合できたという事実の方が強い）。"""
    if key in tier_by_key:
        return tier_by_key[key]
    if key in inferred:
        return "inferred"
    src = sources.get(key)
    if isinstance(src, str):
        if src.startswith(_VOCAB_PREFIX):
            return "vocab"
        if src.startswith(_TASK_LITERAL_PREFIX):
            return "task_literal"
    return "default"


def build_interpretation(op: str, resolved: dict | None, inferred, verdicts, books) -> tuple:
    """resolved/inferred/verdicts を束ねて `interpretation` を組む。
       戻り値: (interpretation: dict, provenance: dict | None)。

       op:       DSL 語彙のオペレーション名（"SORT" 等）。
       resolved: 検証済みの解決値（verify_dsl_args 等の戻り値）。"_" 始まりの内部キー
                 （_target_sheet/_sources/_warnings 等）は slots から除く（_target_sheet
                 だけは "sheet" として別に載せる）。
       inferred: resolved の中で機械推定されたキーの集合。
       verdicts: ailine.classify_subject_provenance() が返す SubjectVerdict のリスト
                 （対象スロットだけが対象・INPUT 種別や限定語の無い LABEL は含まれない）。
       books:    このスロットが対象にしたブック名のイテラブル（★ 決定事項②: 値が1件でも配列）。"""
    resolved = resolved or {}
    inferred = inferred or set()
    sources = resolved.get("_sources") or {}
    tier_by_key = {v.slot.key: v.tier for v in (verdicts or ())}
    kind_by_key = {v.slot.key: v.slot.kind for v in (verdicts or ())}

    slots = []
    for key, value in resolved.items():
        if key.startswith("_"):
            continue
        slots.append({
            "key": key,
            "value": value,
            "kind": kind_by_key.get(key),
            "origin": _slot_origin(key, tier_by_key, inferred, sources),
        })

    interpretation = {
        "op": op,
        "books": list(books),
        "sheet": resolved.get("_target_sheet"),
        "slots": slots,
    }
    provenance = resolved.get("_sources")
    return interpretation, provenance
