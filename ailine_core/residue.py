"""residue — W10 便C2 S5: もしかして提案の残差検出（部分対応の罠への機械の防壁）。

★ なぜ在るか（test_suggest_flow.py 冒頭の決裁参照）: 判定器(judge_ops_via_llm)の
自己申告に頼ると「対応外の部分を黙ったまま提案する」害が起きる。7B に「一部だけ対応で
残りは対応外」と自己申告させる実験は 5/6 で素通りした（部分対応の罠に自分では気づかない）。
指示は意図、保証は機械 ── 提案する側が「この操作に反映される部分」を機械で確定し、
反映されない残りを名指しする。この判定は LLM を一切使わない（+0ms/+0依存）。

★ 判定方法（形態素解析はしない・置換による span 除去）: 依頼文から
  ①数字 ②解決済み args の文字列値（列名・ラベル・引用値など） ③op の照合語彙
  (pool_phrases・label/synonyms/match_phrases) を文字列として取り除き、残った文字列から
  「内容語らしい連続語」（漢字/カタカナ/半角英数字の2文字以上の連続）を拾う。

★ ひらがなは対象にしない ── ailine.py の `_raw_target_not_embedded_in_task`
（単位B・列名照合の断片ガード）が「ひらがな/カタカナは日本語の語境界」として扱うのと
同じ観察を裏返しに使う: 助詞・活用語尾はひらがなで書かれるため、内容語の候補にそもそも
含めなければ、専用の助詞リストを持たなくても「消費」したのと同じ効果になる。

★ オオカミ少年回避が最優先（きれいな依頼に誤って残差行を出すと信頼を失う ──
Namakoo 決裁「迷ったら出さない側に倒す」）: 呼び出し側(ailine.py)は pool_phrases に
op の label/synonyms/match_phrases 全部（suggest_ops の照合プールと同じもの）を渡すこと。
広く消費させるほど、残差行の誤発火（きれいな依頼への誤爆）は減る。

★ 置き場所: ailine_core/（sum_identity.py と同じ理由）。ailine を import しない
（移植可能性の番人 test_line_budget.py が機械で守る）。
"""
from __future__ import annotations

import re

# 内容語らしい連続語: 漢字(2文字以上) / カタカナ(長音符込み・2文字以上) / 半角英数字(2文字以上)。
# ひらがなは含めない（docstring 参照 ── 助詞・活用語尾は拾わないことで「消費」を兼ねる）。
_CONTENT_RUN_RE = re.compile(
    r"[㐀-䶿一-鿿豈-﫿]{2,}|[ァ-ヶー]{2,}|[A-Za-z]{2,}")
_NUMBER_RE = re.compile(r"\d+(?:\.\d+)?")


def find_unconsumed_words(task: str, resolved_args: dict, pool_phrases) -> list:
    """依頼文 task のうち、resolved_args の文字列値・pool_phrases・数字のどれにも
       消費されなかった内容語を、出現順・重複除去で返す（無ければ空リスト）。
       ★ span 除去方式（文字列としてこの文中に現れるかだけを見る・形態素解析はしない）。
       消費の候補が広いほど安全（オオカミ少年を避ける側に倒れる）ので、resolved_args は
       検証済みの解決値（列名・ラベル等）を、pool_phrases は op の照合語彙を広く渡すこと。"""
    if not task:
        return []
    remaining = _NUMBER_RE.sub(" ", task)
    # ★ 第二波 ④（本家 bug_008）: dict の反復順（＝呼び出し側が args を組んだ key の順）で
    #   はなく、値の**長さ降順**で消費する。「商品」「商品コード」のように片方がもう片方を
    #   部分文字列として含む場合、短い方を先に消費すると長い方の残骸（「コード」）が
    #   偽の残差として漏れる（pool_phrases 側は元々この順でやっていた・args 側に同じ規律を
    #   足すだけ＝pool と対称）。
    for v in sorted({v for v in (resolved_args or {}).values() if isinstance(v, str) and v},
                    key=len, reverse=True):
        if v in remaining:
            remaining = remaining.replace(v, " ")
    for phrase in sorted({p for p in (pool_phrases or ()) if p}, key=len, reverse=True):
        if phrase in remaining:
            remaining = remaining.replace(phrase, " ")
    seen = []
    for m in _CONTENT_RUN_RE.finditer(remaining):
        w = m.group(0)
        if w not in seen:
            seen.append(w)
    return seen


def render_residue_note(words: list) -> str | None:
    """残差語があれば注記1行、無ければ None（呼び出し側は None なら何も印字しない
       ── きれいな依頼に沈黙させるのが既定・S5 の対照検体）。"""
    if not words:
        return None
    joined = "・".join(words)
    return f"（『{joined}』などの部分はこの操作に反映されません）"
