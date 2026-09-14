# -*- coding: utf-8 -*-
"""compare_words — 比較の語（以上／以下／未満／超え…）の**辞書と、それを正確に引く 1 つの関数**。

★★ なぜ在るか（2026-09-14 夜・Namakoo「辞書登録後は正確に引いてこれる仕組みも必要だ」）:
  言い回し 120 件の盲検で辞書に口語を足したが、引き方そのものに欠陥が 4 つあり、
  **辞書を増やすほど悪化する**形だった（変える前に実測・凍結 ── `docs/DESIGN-20260915-辞書を正確に引く.md`）:

    「残業時間が20時間を超えない人を抜き出して」  → gt    ★ 逆（超えない＝以下）
    「在庫数が10個を下回らない品番を抜き出して」  → lt    ★ 逆（下回らない＝以上）
    「金額が3000以上5000未満の行を抜き出して」    → gte   ★ 『未満』を黙って捨てる（最初の一致で採る）
    「在庫数が10を切っていない品番」              → None  ★ 引けない回は LLM の不等号が黙って通る

    ① 否定が読めない（語に「ない」が付いても同じ答え）
    ② 最初の一致で採る（最長一致でない・2 つ目の比較語を捨てる）
    ③ 断片ガードが語ごとの場当たり（2 系統の集合に分かれていた）
    ④ 引けなかった回に黙って LLM に負ける（呼び出し側の規則 ── ここでは Reading(None) を返す）

★ 形: **辞書は表（1 語 1 レコード）**。語だけの tuple をやめ、出所（誰が言ったか）と
  陽性の検体を語と一緒に持つ ── 番人が全件を機械で回す（tests/test_compare_words.py）:
    1. 1 語 1 検体（検体の無い語は入らない）  2. 到達できない語は入らない（自分の検体で勝てない）
    3. 語彙の共食いの凍結（bench/compare_words_freeze.json・408 句の読みが動いたら赤）
    4. 出所の無い語は入らない
★ 引くのは `read(task)` だけ。全位置を走査し、重なる一致は**長い方が勝つ**。
  否定（直後に ない／ません…）は**反転させず聞き返す**（黙って反転させるのは、いま直している
  事故と同じ形 ── 実測の検体が無い語を機械が決めない）。違う比較が 2 つ在れば**曖昧**として
  返す（黙って片方を捨てない）。
★ ここは純関数（ファイルも LLM も触らない）。`ailine` を import しない（可搬性の番人）。
"""
from __future__ import annotations

import re
from dataclasses import dataclass

#: 数値の比較（境目の数が要る比較）── threshold.NUMERIC_CMPS と同じ集合。
NUMERIC_CMPS = frozenset({"gt", "lt", "gte", "lte"})

#: 人に見せる比較の名（★ ailine 本体の _EXTRACT_CMP_LABELS と同じ語 ── 断り文が本体の表示と揃う）。
LABELS = {"gte": "以上", "lte": "以下", "gt": "超", "lt": "未満",
          "eq": "等しい", "contains": "を含む"}

#: 否定したときの比較（数値の 4 つだけ定義できる）。★ 初版は**反転させない** ── 聞き返す文に使う。
NEGATED_OF = {"gt": "lte", "lt": "gte", "gte": "lt", "lte": "gt"}

# ── 断片ガード（3 種だけ。語ごとの場当たりを畳んだ）──────────────────────────────
NONE = "none"                     #: ガード無し（文末定型と衝突しない長い語）
NUM_BEFORE = "num_before"         #: 直前 10 文字（同じ文の中）に数字が在るときだけ比較語
COUNTER_BEFORE = "counter_before" #: 直前 1 文字が数か数え語のときだけ（「締め切って」「区切って」の断片）
GUARDS = (NONE, NUM_BEFORE, COUNTER_BEFORE)

_NUM_WINDOW = 10
_NUM_RE = re.compile(r"[0-9０-９]")
_COUNTER_RE = re.compile(r"[0-9０-９個件人円時間日点本枚台%万千]$")
#: 一致の**直後**に付く否定の形（「超え|ない」「下回|らない」「切っ|ていない」「以上|ではない」）。
_NEGATION_AFTER = re.compile(r"^(?:てい|て|ら|では|じゃ|で|は)?(?:ない|ません|なかった|なけれ|ぬ)")


@dataclass(frozen=True)
class Word:
    """辞書の 1 レコード。

    text     … 依頼文に現れる字面（活用の共通部分。「超え」は 超える／超えて／超えてる に当たる）
    cmp      … 読む比較
    guard    … 断片ガード（NONE / NUM_BEFORE / COUNTER_BEFORE）
    after    … 直後の形の条件（regex・無ければ None）── 2 つの意味を持つ語を後ろで見分ける
    source   … ★ 誰が言ったか（必須・番人が空を赤にする）。頭から思いついた語を足さない
    since    … 入れた日
    example  … ★ 陽性の検体 1 件（必須・番人が全件回し、この語が勝つことを確かめる）
    """
    text: str
    cmp: str
    guard: str = NONE
    after: str | None = None
    source: str = ""
    since: str = ""
    example: str = ""


#: ★ 出所の略記: OP9 = operator 盲検 9 回目（PRED-20260823-operator9.md）で「意味から広めた」列挙。
#:   実文の出所を持たない語はそう書く（出所を偽らない）。
_OP9 = "operator9（2026-08-23）意味から広めた列挙・実文なし"
_BLIND = "言い回し 120 件の盲検（2026-09-14・sonnet の役 agent）"

WORDS: tuple[Word, ...] = (
    # gt ── より大きい
    Word("より大きい", "gt", source=_OP9, since="2026-08-23", example="金額が5000より大きい行を抜き出して"),
    Word("より大きく", "gt", source=_OP9, since="2026-08-23", example="金額が5000より大きくなった行"),
    Word("を超える", "gt", source=_OP9, since="2026-08-23", example="金額が5000を超える行を抜き出して"),
    Word("を超えて", "gt", source=_OP9, since="2026-08-23", example="金額が5000を超えている行"),
    Word("より多い", "gt", source=_OP9, since="2026-08-23", example="数量が10より多い行"),
    Word("より多く", "gt", source=_OP9, since="2026-08-23", example="数量が10より多くある行"),
    Word("より高い", "gt", source=_OP9, since="2026-08-23", example="単価が1000より高い行"),
    Word("より高く", "gt", source=_OP9, since="2026-08-23", example="単価が1000より高くなった行"),
    Word("超え", "gt", NUM_BEFORE, source=_BLIND + " #56「残業時間が20時間超えてる人」",
         since="2026-09-14", example="残業時間が20時間超えてる人"),
    Word("上回", "gt", NUM_BEFORE, source="需要センサ misclass.jsonl「発注点を上回る」（設計書 §3）",
         since="2026-09-14", example="金額が5000を上回る行"),
    Word("過ぎ", "gt", NUM_BEFORE, source="Qwen3-8B の掃き 167 件「退勤時間が22時を過ぎる分」",
         since="2026-09-14", example="退勤時間が22時を過ぎる分だけ抽出しといて"),
    # lt ── より小さい
    Word("未満", "lt", source=_OP9, since="2026-08-23", example="金額が5000未満の行を抜き出して"),
    Word("より小さい", "lt", source=_OP9, since="2026-08-23", example="金額が5000より小さい行を抜き出して"),
    Word("より小さく", "lt", source=_OP9, since="2026-08-23", example="金額が5000より小さくなった行"),
    Word("より少ない", "lt", source=_OP9, since="2026-08-23", example="数量が10より少ない行"),
    Word("より少なく", "lt", source=_OP9, since="2026-08-23", example="数量が10より少なくなった行"),
    Word("より安い", "lt", source=_OP9, since="2026-08-23", example="単価が1000より安い行"),
    Word("より安く", "lt", source=_OP9, since="2026-08-23", example="単価が1000より安くなった行"),
    Word("切っ", "lt", COUNTER_BEFORE, source=_BLIND + " #61「在庫数が10個切ってるの教えて」",
         since="2026-09-14", example="在庫数が10個切ってるの教えて"),
    Word("下回", "lt", NUM_BEFORE, source="需要センサ misclass.jsonl「発注点を下回る」（設計書 §3）",
         since="2026-09-14", example="在庫が発注点の5を下回る行"),
    Word("に満たない", "lt", NUM_BEFORE, source="設計書 §3（口語の列挙・2026-09-14）",
         since="2026-09-14", example="金額が5000に満たない行"),
    Word("マイナス", "lt", after=r"^(?:になっ|になる|の行|のもの|の品|だけ)",
         source=_BLIND + " #21「在庫数がマイナスになってる行を教えて」・#74",
         since="2026-09-14", example="在庫数がマイナスになってる行を教えて"),
    # gte / lte ── 文末定型（「以上です」）の断片になるので数字が近くに要る
    Word("以上", "gte", NUM_BEFORE, source=_OP9, since="2026-08-23", example="金額が5000以上の行を抜き出して"),
    Word("以下", "lte", NUM_BEFORE, source=_OP9, since="2026-08-23", example="金額が5000以下の行を抜き出して"),
    # contains / eq
    Word("を含む", "contains", source=_OP9, since="2026-08-23", example="備考に東京を含む行を抜き出して"),
    Word("を含んで", "contains", source=_OP9, since="2026-08-23", example="備考に東京を含んでいる行"),
    Word("が含まれる", "contains", source=_OP9, since="2026-08-23", example="備考に東京が含まれる行"),
    Word("を含める", "contains", source=_OP9, since="2026-08-23", example="備考に東京を含める行"),
    Word("と等しい", "eq", source=_OP9, since="2026-08-23", example="状態が完了と等しい行を抜き出して"),
    Word("に等しい", "eq", source=_OP9, since="2026-08-23", example="状態が完了に等しい行"),
    Word("と同じ", "eq", source=_OP9, since="2026-08-23", example="状態が完了と同じ行"),
)


@dataclass(frozen=True)
class Reading:
    """`read()` の結果。

    cmp        … 機械が引けた比較（1 つに決まった時だけ・それ以外は None）
    word       … 勝った語（cmp が在る時）
    hit        … 比較の語が**在った**か（否定・曖昧でも True ── 「比較の依頼である」ことの印）
    ambiguous  … 空でなければ**聞き返す／断る**文（否定・2 つの比較）。呼び出し側は cmp を使わない
    negated    … 否定が付いていた
    """
    cmp: str | None = None
    word: str | None = None
    hit: bool = False
    ambiguous: str = ""
    negated: bool = False


def _passes_guard(w: Word, text: str, idx: int) -> bool:
    if w.guard == NUM_BEFORE:
        window = text[max(0, idx - _NUM_WINDOW):idx]
        if "。" in window or not _NUM_RE.search(window):
            return False
    elif w.guard == COUNTER_BEFORE:
        if not _COUNTER_RE.search(text[:idx]):
            return False
    if w.after is not None and not re.match(w.after, text[idx + len(w.text):]):
        return False
    return True


def matches(task: str | None) -> list:
    """ガードを通った一致を、重なりを畳んで（同じ場所は長い語が勝つ）出現順に返す。
    要素は (start, end, Word, negation) ── negation は直後に付いた否定の字面（無ければ ""）。"""
    text = task or ""
    if not text:
        return []
    cands = []
    for w in WORDS:
        idx = text.find(w.text)
        while idx >= 0:
            if _passes_guard(w, text, idx):
                cands.append((idx, idx + len(w.text), w))
            idx = text.find(w.text, idx + 1)
    # ★ 最長一致: 開始位置順・同じ開始なら長い方を先に。前に採った一致と重なる候補は捨てる。
    cands.sort(key=lambda c: (c[0], -(c[1] - c[0])))
    taken, last_end = [], -1
    for s, e, w in cands:
        if s < last_end:
            continue
        m = _NEGATION_AFTER.match(text[e:])
        taken.append((s, e, w, m.group(0) if m else ""))
        last_end = e
    return taken


def read(task: str | None) -> Reading:
    """依頼文から比較を 1 つ引く。引けない時は決めない（Reading.cmp=None）。

    - 一致なし                → Reading()（hit=False）。★ 呼び出し側は LLM の不等号を採らず聞き返す
    - 否定が付いた            → 聞き返す（反転させない）
    - 違う比較が 2 つ以上     → 曖昧（範囲の抽出はこの道具にない）── 黙って片方を捨てない
    - 同じ比較が何度出ても    → その比較（「5000以上の行を、以上で」など）
    """
    found = matches(task)
    if not found:
        return Reading()
    negs = [(w, w.text + n) for _s, _e, w, n in found if n]
    if negs:
        w, shown = negs[0]
        inv = NEGATED_OF.get(w.cmp)
        hint = (f"『{LABELS[inv]}』のことですか？ " if inv else "")
        return Reading(None, w.text, True, negated=True, ambiguous=(
            f"依頼文の『{shown}』は否定の比較です ── {hint}"
            "否定の比較は機械で確かめられないので実行しません"
            + (f"（例: 「金額が5000{LABELS[inv]}の行」のように書き直してください）" if inv
               else "（列に在る値を名指しして「〜以外」と書いてください）")))
    cmps = list(dict.fromkeys(w.cmp for _s, _e, w, _n in found))
    numeric = [c for c in cmps if c in NUMERIC_CMPS]
    if len(numeric) > 1:
        # ★ 範囲（『以上』と『未満』）── 測った誤配の形。黙って片方を捨てない。
        shown = "』『".join(dict.fromkeys(w.text for _s, _e, w, _n in found if w.cmp in NUMERIC_CMPS))
        return Reading(None, None, True, ambiguous=(
            f"比較の語が 2 つあります（『{shown}』）── 範囲の抽出はこの道具にありません。"
            "片方で頼んでください（例: 「金額が5000以上の行を抜き出して」）"))
    # ★ 数値の比較と 等しい／含む が並ぶ回（「金額が1000以上で部門が営業と同じ行」）は **2 条件の
    #   依頼** ── 比較は数値の側。2 組目は呼び出し側が実表の値で読む（2026-09-06 の AND の道）。
    # ★ 数値でない比較が 2 つ（「を含む」と「と同じ」）は実文が無く測っていない ── 旧来どおり
    #   先に出た語（保留・発火条件: 実文が 1 件出た日に、範囲と同じ線で断るかを測る）。
    chosen = next((w for _s, _e, w, _n in found if w.cmp in NUMERIC_CMPS), found[0][2])
    return Reading(chosen.cmp, chosen.text, True)


def unconfirmed(llm_cmp: str, *, example: str) -> str:
    """辞書に当たらないのに LLM が**数値の比較**を返した回の断り文（呼び出し側の規則 ④）。

    ★ 「引けなかった回は LLM に黙って負ける」を閉じる ── 依頼／宣言／実体の三項のうち
      依頼の側に比較が見つからないなら、宣言（LLM の不等号）だけで通さない。
      依頼文が無い経路（DSL 直渡し）は呼び出し側がここへ来ない。
    """
    return ("比較の語（以上／以下／未満／超え…）が依頼文に見つかりません ── "
            f"LLM が返した比較({LABELS.get(llm_cmp, llm_cmp)})を機械で確かめられないので実行しません"
            f"（例: 「{example}」のように比較の語を書いてください）")
