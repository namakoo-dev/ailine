"""subject — 単位E: A' 原則を「値」から「対象スロット」へ広げる。

★ 症状（ブラインド査定2本が独立に致命の筆頭に置いた）: 依頼「数量と単価をかけた金額列を
作って、見出しを太字にして」に対し、2段目の対象が `col:数量*単価`（前段が作った新規列）に
解決され、**見出し行は太字にならないまま**「✓ 機械検証済み」が出た。
**依頼文には「見出し」という語がある。解決された対象は `col:数量*単価`。この2つは矛盾して
いるのに、機械はそれを一度も突き合わせていなかった。**

★ 足場は既にあった: `_sources`（倍率・値の出典）と「LLM が返した値 vs 機械抽出の値の
食い違い警告」は、どちらも**値スロット**に対する同じ形の仕掛けだった。ここはそれを
**対象スロット**（どこを操作するか）へ広げる。

## 出所の3段階（★ 「証拠が無い」と「反証がある」を同じ扱いにしない）

| 段階 | 意味 | 振る舞い |
|---|---|---|
| ① MATCHED | 依頼文の語と機械照合できた | ✓ 満額（何も足さない） |
| ② UNSPOKEN | 依頼文はそのカテゴリについて無言（ブック実体・既定から機械決定） | ✓ ＋ その run 固有の1文 |
| ③ CONTRADICTED | 依頼文にそのカテゴリの語があり、解決値と一致しない | ✓ を出さない・⚠ ＋ 確認 |

**② は主張の範囲を狭めるだけ・③ だけが止める。** 証拠が無いこと（②）を反証がある
こと（③）と同じに扱うと、まっとうな run が軒並み止まる。

## ★ 「依頼文にそのカテゴリの語がある」の意味 ―― **誰も拾わなかった語だけが反証**

素朴に「依頼文にどれか対象の語がある」を条件にすると、複合依頼で必ず誤爆する（実測で
掴んだ: 依頼「金額で降順に並べ替えて」＋計画 [SORT 金額 / BOLD row:1] の 2 段目が③に落ち、
f7 のゴールデンが動いた）。依頼文の「金額」は 1 段目が正しく拾っており、2 段目の対象に
ついて依頼文は**無言**だからだ。

そこで反証の条件を**消費**で定義する: 依頼文が指した対象のうち、**ここまでのどのスロットも
照合しなかった語が残っている**のに、このスロットの解決値が依頼文と照合できない場合だけ③。
残りが無ければ②（無言）。★ 「見出しを太字に」→`col:数量*単価` が③になるのは、
「見出し」を誰も拾っていないから。★ 「金額で降順に」→`row:1` が②なのは、「金額」を
1 段目が拾い、残りが無いから。

## ★ 凍結した照合の定義（先に決め、数字を見てから動かさない）

解決値 `V`（列名・シート名）が依頼文と「機械照合できた」とは:

  (i) `V` が依頼文に部分文字列として現れる（＝この repo が既に3箇所で使っている慣行
      ―― `resolve_target_sheet` の実在シート名照合・`extract_task_mentions`・
      LOOKUP_FILL の `raw_str in task`）。**ただしその出現が、他の実在名の一部としてしか
      説明できないなら証拠にしない**（(ii) と同じ曖昧性の排除・下記）。
  (ii) `V` の長さ2以上の連続部分文字列 `s` で、`s` が依頼文に現れ、かつ **`s` が他のどの
      実在名（同じブックの他の列名/シート名）にも現れない**ものが存在する。

**なぜ (ii) を足したか（誤爆＝バー2の側の要求）**: 人は実在名を短く言う（実列名
『小計金額』に対して依頼文は「小計の列を太くして」）。完全一致だけだと、この正しい run が
②へ落ちて注記が付き、他の列名が依頼文に混ざる変種では③へ落ちて **✓ が消える**。

**なぜ無条件の双方向部分一致にしなかったか（真陽性＝バー1の側の要求）**: `数量*単価` は
`数量` を含むので、無条件の部分一致だと「数量と単価をかけた…見出しを太字に」の依頼文と
照合してしまい、**症状そのものが①になって素通りする**。断片が他の実在列も指しうるなら、
その断片は照合の証拠にならない ―― この曖昧性の排除だけが、両方のバーを同時に満たす線。

**★ (i) にも同じ排除を当てた（実測 2026-08-17 の穴・単位B の残り）**: 列が『商品/金額/
税込金額』のブックで「税込金額で並べ替えて」→ 解決値『金額』が (i) の素朴な部分文字列だけで
①になり、**税抜きの列が並べ替えられたまま ✓ が出ていた**。『金額』の依頼文中の出現は
『税込金額』の一部としてしか説明できない ―― 依頼者が言ったのは『税込金額』の方だ。
そこで (i) も「他の実在名に飲み込まれていない出現が1つ以上あること」を要求する
（`_standalone_occurrence`）。★ 逆向き（依頼文「金額で」＋解決値『税込金額』）は元から③で、
そちらの判定は動かさない ―― 片方向だけ抜けていた穴を塞ぐ変更。

★ A' 原則はここでも同じ: 照合の材料は**実在物だけ**（実在列名・実在シート名・行番号・
見出し行）。依頼文から自由に名前を切り出すことはしない（LLM も使わない）。
★ ここは純ロジック（ファイルを開かない・ailine を import しない）。どのスロットが
「対象」かの宣言（OP_SUBJECT_SLOTS）と、実在列名/見出し行の取得は ailine.py 側にある。
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field

# --- 出所の3段階 -------------------------------------------------------------
MATCHED = "matched"              # ① 依頼文の語と機械照合できた
UNSPOKEN = "unspoken"            # ② 依頼文はそのカテゴリについて無言
CONTRADICTED = "contradicted"    # ③ 依頼文にそのカテゴリの語があり、解決値と一致しない

# --- 対象スロットの種別 -------------------------------------------------------
COLUMN = "column"   # 列名そのもの（SORT の col・AGGREGATE の group_col 等）
REGION = "region"   # col:列名 / row:N / all の複合形（BOLD/FILL_COLOR/CENTER_ALIGN の target）
ROW = "row"         # 行番号（INSERT_ROWS の at）
SHEET = "sheet"     # シート名（resolved["_target_sheet"]・複数シートのブックだけ）
LABEL = "label"     # 書き込むラベル（★ 金額の性質を限定する語が依頼文にある時だけ問う。下記）
# ★ INPUT: 「対象」ではないが計画が実際に使った実在列（COMPUTE_COLUMN の演算対象など）。
#   判定はしない（①②③ を付けない）が、**依頼文の語を消費する**。これが無いと
#   「売上から原価を引いた利益列を作って、利益で降順に」のような正しい連鎖で、
#   誰にも拾われないままの『売上』『原価』が 2 段目の反証に化ける（実測で掴んだ誤爆）。
INPUT = "input"
# ★ operator8 ①: LOOKUP_FILL の source_sheet のように「対象」ではないが依頼文のシート言及を
#   消費する実在シート名。INPUT の列版と同じ考え方（「言及は参照側で消費された」）だが、
#   consumed の家系が列/行/全体とは別（sheets）なので INPUT とは別種別にする
#   （_match_slot で others=sheets・token family="sheet" を使うため）。
SHEET_INPUT = "sheet_input"

_KINDS = frozenset({COLUMN, REGION, ROW, SHEET, LABEL, INPUT, SHEET_INPUT})

# 「見出し行」を指す語。★ 実在物への接地: 見出し行はブックの実体（book_meta["header_rows"]）
# として機械が知っている行番号なので、この語は行番号への機械照合が可能。
_HEADER_WORD_RE = re.compile(r"見出し|ヘッダー|ヘッダ")
# 「表全体」を指す語（target="all" と照合する）。
_WHOLE_WORD_RE = re.compile(r"全部|全体|すべて|全て|全セル")
_ROW_ORDINAL_RE = re.compile(r"(\d+)\s*行目")
_ROW_PLAIN_RE = re.compile(r"行\s*(\d+)")
# ★ 金額の性質を限定する語（税込み/税抜き）。解釈のどこにもこの限定が現れないなら、
#   その ✓ は「依頼どおり」ではない（査定Bの致命1:「税込み合計」→ ラベル『合計』）。
_QUALIFIER_CHARS = ("税", "込", "抜")

_MIN_FRAGMENT = 2   # 照合の定義 (ii) の最小長（1文字の漢字は偶然一致しすぎる）


@dataclass(frozen=True)
class Slot:
    """判定の入力1件。

    key:     resolved args のキー（表示はしない・報告とデバッグ用）
    value:   解決値（文字列化済み。INSERT_ROWS の at のような整数も文字列で受ける）
    kind:    上の種別のどれか
    context: LABEL 種別だけが使う「同じ解釈行の他のフィールド」（対象列名など）。
             限定語がそちらに現れていれば、ラベルに無くても限定は解釈に現れている。
    """
    key: str
    value: str
    kind: str
    context: str = ""


@dataclass(frozen=True)
class TaskDesignators:
    """依頼文が**機械照合可能な形で**指している対象の一覧（実在物のみ）。

    columns/sheets は実在名そのもの、rows は1起点の行番号、whole は「表全体」の語の有無。
    *_words は表示用（「依頼文が指しているのは: …」の中身）。"""
    columns: tuple = ()
    rows: tuple = ()          # ★ row_words と**同じ長さ・同じ順序**（_remaining が zip で対にする）
    row_words: tuple = ()
    whole_word: str = ""
    sheets: tuple = ()

    @property
    def whole(self) -> bool:
        return bool(self.whole_word)


@dataclass(frozen=True)
class SubjectVerdict:
    """スロット1件の判定結果。designators は③の説明に使う「依頼文が指していた語（未消費）」。"""
    slot: Slot
    tier: str
    designators: tuple = field(default=())


@dataclass
class Consumed:
    """★ 「誰も拾わなかった語だけが反証」を成立させるための、run をまたぐ消費の記録。
       ①になったスロットが実際に照合した実在名/行番号/全体の語をここへ積む。
       複合計画では段をまたいで持ち回る（1段目が拾った語で2段目が誤爆しないようにする）。
       ★ 可変オブジェクト（frozen にしない） ―― 段をまたいで足していく台帳そのものだから。"""
    columns: set = field(default_factory=set)
    rows: set = field(default_factory=set)
    whole: bool = False
    sheets: set = field(default_factory=set)


# --- ★ 凍結した照合の定義（モジュール冒頭の docstring 参照） ---------------------

def _fragment_credit(name: str, task: str, others) -> bool:
    """定義 (ii): name の長さ2以上の連続部分文字列で、依頼文に現れ、かつ他のどの実在名にも
       現れないものがあるか。★ 他の実在名も指しうる断片は証拠として採らない（曖昧性の排除）。"""
    n = len(name)
    for i in range(n):
        for j in range(i + _MIN_FRAGMENT, n + 1):
            s = name[i:j]
            if s in task and not any(s in o for o in others):
                return True
    return False


def _standalone_occurrence(name: str, task: str, others) -> bool:
    """定義 (i): name が依頼文に現れる ―― ただし **他の実在名の一部としてしか説明できない
       出現は証拠にしない**（(ii) と同じ曖昧性の排除を (i) にも当てる）。

       ★ なぜ（実測 2026-08-17 の穴）: 列が『商品/金額/税込金額』のブックで「税込金額で
       並べ替えて」→ 解決値『金額』が ① になり ✓ が出ていた。『金額』は依頼文に確かに現れる
       が、その出現は『税込金額』の一部としてしか現れていない ―― 依頼者が言ったのは
       『税込金額』であって『金額』ではない。**税込のつもりで頼んで税抜きの列が並べ替えられ、
       しかも ✓ が出る**という実害のある型なので、この出現は証拠として採らない。
       ★ 逆向き（依頼文「金額で」＋解決値『税込金額』）は元々③（(i) も (ii) も成立しない）
       ―― こちらの判定は変えない。"""
    covers = [o for o in others if name in o and o != name]
    if not covers:
        return name in task
    spans = []
    for o in covers:   # 他の実在名が依頼文のどこを占めているか
        at = task.find(o)
        while at >= 0:
            spans.append((at, at + len(o)))
            at = task.find(o, at + 1)
    at = task.find(name)
    while at >= 0:
        if not any(s <= at and at + len(name) <= e for s, e in spans):
            return True   # 他のどの実在名にも飲み込まれていない出現がある＝依頼者はこの名を言った
        at = task.find(name, at + 1)
    return False


def name_matches_task(name, task: str, others=()) -> bool:
    """実在名 name が依頼文と機械照合できたか（定義 (i) or (ii)）。
       others は「同じブックの他の実在名」（name 自身は除いて渡してよい・ここでも除く）。"""
    name, task = str(name or ""), str(task or "")
    if not name or not task:
        return False
    others = [str(o) for o in others if o and str(o) != name]
    if _standalone_occurrence(name, task, others):
        return True
    return _fragment_credit(name, task, others)


def task_designators(task: str, columns=(), header_row: int = 1, sheets=()) -> TaskDesignators:
    """依頼文が指している対象を実在物との照合だけで拾う（②と③を分ける材料）。
       ★ ここは完全一致（部分文字列）だけを使う ―― 「依頼文がそのカテゴリについて何か
       言ったか」は保守的に判定する（断片一致まで designator と認めると、②であるべき run が
       ③に落ちて ✓ が消える）。
       ★ ただし (i) と同じ曖昧性の排除は当てる（_standalone_occurrence）: 『税込金額』としか
       現れていない『金額』を「依頼文が指した対象」に数えると、⚠ が「照合できません（依頼文が
       指しているのは: 金額・税込金額）」と自己矛盾して読め、さらに『税込金額』を拾った次の段が
       残った『金額』で誤って③に落ちる。"""
    task = str(task or "")
    names = [str(c) for c in columns if c]
    cols = tuple(c for c in columns if c and _standalone_occurrence(
        str(c), task, [o for o in names if o != str(c)]))
    rows: list = []
    words: list = []
    for pat in (_ROW_ORDINAL_RE, _ROW_PLAIN_RE):
        for m in pat.finditer(task):
            n = int(m.group(1))
            if n >= 1 and n not in rows:
                rows.append(n)
                words.append(m.group(0))
    hm = _HEADER_WORD_RE.search(task)
    if hm and header_row and header_row >= 1 and header_row not in rows:
        # ★ 見出し行はブックの実体（header_rows）として機械が知っている行番号なので、
        #   「見出し」という語はここで初めて実在物への照合対象になる。
        rows.append(header_row)
        words.append(hm.group(0))
    wm = _WHOLE_WORD_RE.search(task)
    snames = [str(s) for s in sheets if s]
    return TaskDesignators(columns=cols, rows=tuple(rows), row_words=tuple(words),
                            whole_word=wm.group(0) if wm else "",
                            sheets=tuple(s for s in sheets if s and _standalone_occurrence(
                                str(s), task, [o for o in snames if o != str(s)])))


def _row_value(raw: str):
    s = str(raw).strip()
    return int(s) if s.isdigit() else None


def _match_slot(slot: Slot, task: str, columns, d: TaskDesignators, sheets,
                 qualifier_signal: bool):
    """1件目のパス: そのスロットが依頼文と照合できたか（matched, 消費した語）を返す。
       戻り値: (matched: bool | None, token or None)。matched=None は「この種別は判定の
       対象そのものにしない」（INPUT・限定語の無い LABEL）。**token は照合できた時だけ
       返す**（消費されるのは実際に拾われた語だけ）。"""
    raw = str(slot.value)
    if slot.kind == INPUT:   # 判定はしない（None）が、照合できたなら語を消費する
        matched = name_matches_task(raw, task, others=columns)
        return None, (("column", raw) if matched else None)
    if slot.kind == SHEET_INPUT:   # ★ operator8 ①: シート版の INPUT（判定はしない・消費のみ）
        matched = name_matches_task(raw, task, others=sheets)
        return None, (("sheet", raw) if matched else None)
    if slot.kind == SHEET:
        if name_matches_task(raw, task, others=sheets):
            return True, ("sheet", raw)
        return False, None
    if slot.kind == LABEL:
        if not qualifier_signal:
            return None, None   # 限定語が無い依頼では、ラベルは問わない（毎回の注記＝情報量ゼロ）
        return any(ch in (raw + slot.context) for ch in _QUALIFIER_CHARS), None
    if slot.kind == ROW:
        n = _row_value(raw)
        hit = n is not None and n in d.rows
        return hit, (("row", n) if hit else None)
    if slot.kind == REGION:
        if raw == "all":
            return d.whole, (("whole", True) if d.whole else None)
        if raw.startswith("row:"):
            n = _row_value(raw[4:])
            hit = n is not None and n in d.rows
            return hit, (("row", n) if hit else None)
        if raw.startswith("col:"):
            name = raw[4:]
            hit = name_matches_task(name, task, others=columns)
            return hit, (("column", name) if hit else None)
        return None, None   # 未知の形は判定しない（保守的）
    hit = name_matches_task(raw, task, others=columns)
    return hit, (("column", raw) if hit else None)


def _remaining(d: TaskDesignators, consumed: Consumed, kind: str) -> tuple:
    """まだどのスロットも拾っていない依頼文の語（③の反証そのもの）。
       シートは別の家系（列/行/全体とは互いに反証にならない）。"""
    if kind == SHEET:
        return tuple(s for s in d.sheets if s not in consumed.sheets)
    left = tuple(c for c in d.columns if c not in consumed.columns)
    left += tuple(w for n, w in zip(d.rows, d.row_words) if n not in consumed.rows)
    if d.whole and not consumed.whole:
        left += (d.whole_word,)
    return left


def classify_slots(slots, *, task: str, columns=(), header_row: int = 1, sheets=(),
                    qualifier_signal: bool = False, consumed: Consumed | None = None) -> list:
    """対象スロット群を①②③に仕分ける（純関数）。slots は Slot のリスト。
       columns/sheets は**対象シートの実在列名・ブックの実在シート名**（照合の材料）。
       consumed を渡すと、そこに積まれた「既に拾われた語」を反証から除き、この呼び出しで
       ①になった語を積み足す（複合計画が段をまたいで持ち回るための台帳・上記 Consumed 参照）。"""
    task = str(task or "")
    d = task_designators(task, columns, header_row, sheets)
    consumed = consumed if consumed is not None else Consumed()
    targets = [s for s in slots if s.kind in _KINDS and str(s.value)]

    results = []
    for slot in targets:
        matched, token = _match_slot(slot, task, columns, d, sheets, qualifier_signal)
        results.append((slot, matched))
        if token:   # 実際に拾われた語は、他のスロットの反証にはならない
            kind, value = token
            if kind == "column":
                consumed.columns.add(value)
            elif kind == "row":
                consumed.rows.add(value)
            elif kind == "whole":
                consumed.whole = True
            elif kind == "sheet":
                consumed.sheets.add(value)

    verdicts = []
    for slot, matched in results:
        if matched is None:
            continue            # 判定対象にしない種別（上記）
        if matched:
            verdicts.append(SubjectVerdict(slot, MATCHED, ()))
            continue
        if slot.kind == LABEL:  # 限定語は「残り物」の概念を持たない（依頼文に在るか否かだけ）
            verdicts.append(SubjectVerdict(slot, CONTRADICTED, ("税込み/税抜き等の限定",)))
            continue
        left = _remaining(d, consumed, slot.kind)
        verdicts.append(SubjectVerdict(slot, CONTRADICTED if left else UNSPOKEN, left))
    return verdicts


# --- 表示（③の⚠行・②の素材） -------------------------------------------------

def _subject_phrase(slot: Slot, note: str = "") -> str:
    head = "対象シート" if slot.kind == SHEET else "対象"
    return f"{head}『{slot.value}』（{note}）" if note else f"{head}『{slot.value}』"


def contradiction_lines(verdicts, notes=None) -> list:
    """③ のスロットごとの ⚠ 行（適用前に出す・確認の関所へ渡す理由そのもの）。

       notes: 解決値 → **その対象の出所を言う1句**（例:「この計画の直前の段で新規作成された
       列」）。★ 呼び出し側が同じスロットについて別の ⚠ 行を持っている場合、その事実を
       ここに畳み込んで**1本の ⚠ にする**ためにある ―― 意味の違う2文でも、同じスロットに
       ついて2行並べば読み手には冗長で、どちらを見て判断すればよいか分からなくなる
       （事実は落とさない・言うのは1度）。"""
    notes = notes or {}
    lines = []
    for v in verdicts:
        if v.tier != CONTRADICTED:
            continue
        if v.slot.kind == LABEL:
            lines.append(
                f"⚠ 依頼文は金額の性質を限定していますが（税込み/税抜き等）、解釈にその限定が"
                f"現れていません（ラベル『{v.slot.value}』）。意図した金額か確認してください")
            continue
        pointed = "・".join(str(w) for w in v.designators) or "（不明）"
        lines.append(
            f"⚠ {_subject_phrase(v.slot, notes.get(str(v.slot.value), ''))}は依頼文の語と機械照合"
            f"できません（依頼文が指しているのは: {pointed}）。意図した対象か確認してください")
    return lines


def unspoken_subjects(verdicts) -> list:
    """② のスロットの表示句（「対象『X』」「対象シート『Y』」）。✓ の直後の1文の材料。"""
    return [_subject_phrase(v.slot) for v in verdicts if v.tier == UNSPOKEN]
