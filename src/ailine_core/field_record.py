"""field_record — 集めた 1 項目の記録と、その**区分の導出**。

★★ この層だけが区分（確 / 単 / 割 / 無）を作る。
  `tests/test_a_grade_has_exactly_one_source.py` が機械で縛っている ──
  出口（画面・出力ブック・検分シート・--json）が自分で文字列を組み立てると、
  導出の規則を直しても出口の 1 つが古い判断のまま残る（片配線・docs/開発手法.md §13）。

★★ 区分の意味（2026-09-11 に凍結・Namakoo 決裁）:

    確   独立した根拠が 2 つ以上あり、値が一致した
    単   根拠が 1 つだけ（矛盾はしていない・裏が取れていないだけ）→ 値を出す＋印
    割   根拠が 2 つ以上あり、食い違った                        → ★ 値を出さない
    無   根拠が無い                                             → 値を出さない

  ★ 「食い違ったら空欄」の根拠:
      ① 空欄は誤値より安い（経理の表で、間違った金額より空いている方が直せる）
      ② $0 条件『黙って失敗する』から最も遠い
      ③ 単 と 割 を分ける意味がここで立つ ── 単は「まだ裏が無い」、
        割は「矛盾している」。矛盾を値で出すと、人はどちらが正しいか分からないまま片方を受け取る

★★ 空欄には**必ず理由を添える**（同日・Namakoo 追加）:
  空欄だけだと「取れなかった」と「矛盾していた」が同じ顔になり、
  買い手には「この道具は空欄を返してくる」の 1 印象に潰れる。区分を分けた意味が画面で消える。
  → **理由を持たずに空欄を作れない形**にする（下の `Record` は理由を必須引数にしている）。
    後から足す形にすると、片方の経路だけ理由なしで空欄を書く事故が必ず生える。

★★ 独立した根拠を数えるときの落とし穴（2026-09-10 に実測で踏んだ）:
  結合セルを展開したまま数えると、**同じ値セルが結合の幅ぶん重複**する。
  それを 2 つ目の根拠として数えると、**自分の写しと一致して「裏が取れた」**ことになる。
  実測: 畳まないと 40 冊すべて「確」／アンカーで畳むと 確 14・単 26。
  → だから `Evidence` は `at`（アンカーの番地）を持ち、`grade()` は**番地で畳んでから**数える。

★ ailine を import しない（ailine_core の作法）。
"""
from __future__ import annotations

from dataclasses import dataclass, field as _field

#: 区分の語。★ ここが唯一の出どころ（番人が src 全体を AST で走査して縛る）。
CONFIRMED = "確"
SINGLE = "単"
SPLIT = "割"
NONE_FOUND = "無"

#: 値を出してよい区分。★ 割 と 無 は空欄（凍結した判断）。
GRADES_WITH_VALUE = (CONFIRMED, SINGLE)


@dataclass(frozen=True)
class Evidence:
    """1 つの根拠 ──「どこから・どう読んで・何が出たか」。

    ★ `at` はアンカーの番地。**同一性はこれで判定する**（結合の写しを数えないため）。
    ★ `how` は人が追える 1 行（検分シートにそのまま出る）。
    """
    rule: str          #: 規則の名前（"上部ラベルの右" など）
    value: object      #: 読めた値
    at: str            #: "H39"（★ アンカーの番地・必須）
    how: str           #: 読み方の 1 行


@dataclass(frozen=True)
class Record:
    """1 項目の記録。★ 区分と値は**持たない** ── `grade()` / `value()` が導く。

    ★ `blank_reason` は空欄になったときの理由。**必須**にしてあるのは、
      「理由を持たずに空欄を作れない」を型で守るため（上の docstring 参照）。
      根拠が在って値が出る場合は空文字でよい。
    """
    field: str                    #: "請求元" / "請求額(税込)" …
    evidences: tuple = ()         #: Evidence の並び（0 個なら 無）
    rivals: tuple = ()            #: 採らなかった候補（番地つき・検分に出す）
    excluded: tuple = ()          #: 捨てた候補と、捨てた理由
    blank_reason: str = ""        #: ★ 空欄のときの理由（下の検査が空を許さない）
    swept: bool = False           #: ★ 食い違う口が無いことを掃き出して確かめたか
    swept_how: str = ""           #: 何を掃き出したかの 1 行（人に見せる）
    conflict: bool = False        #: ★ 値そのものではない所で食い違いを見つけたか
    conflict_why: str = ""        #: その食い違いの 1 行

    def __post_init__(self):
        if grade(self) not in GRADES_WITH_VALUE and not self.blank_reason:
            raise ValueError(
                f"{self.field}: 空欄なのに理由が無い ── "
                "空欄には必ず理由を添える（2026-09-11 に凍結した判断）")


def _distinct_sources(evidences) -> dict:
    """根拠を**アンカーの番地で畳む**。★ 結合の写しを 2 つ目に数えないため。

    戻り値: 番地 → その番地から出た値（同じ番地に別の値は在り得ない）
    """
    out = {}
    for e in evidences:
        out.setdefault(e.at, e.value)
    return out


def grade_of(evidences, swept: bool = False, conflict: bool = False) -> str:
    """根拠の並びから区分を導く。★ **区分を作るのはここだけ。**

    ★ 数えるのは「根拠の個数」ではなく「**別々の出所の個数**」。
      同じセル（アンカー）から 2 回読んでも 1 つ。

    ★★ `swept`（掃き出したか）── 2026-09-11、未見 51 冊の実測から入れた条件。
      初版は「開いた 2 つの口が一致した」だけで `確` と言っていた。
      だが実測 12 件すべてで、**開いていない第三の口**が食い違っていた
      （源泉徴収の振込金額／備考の再掲／明細の合計／2 枚目のシート／通貨／
        そもそも納品書か見積書か）。
      ★ 一致は「他の口が黙っている」ことを意味しない ── **開いていないだけ**だ。
      → だから `確` は「一致した」ではなく
        「**一致した ＋ 食い違う口が無いことを見た**」を意味する。
        掃き出していないなら、言えるのは `単` まで（裏が取れていない）。
      ★ これは「到達できず＝未確認であって安全でない」と同じ線。

    ★ `Record` を組む前に区分を知りたい呼び出し側のために、並びを直接受ける口を出す
      （これが無いと、上の層が `Record` を偽造して区分を先読みし、
       「区分は 1 箇所からしか作れない」が骨抜きになる）。

    ★★ `conflict`（値そのものではない所で食い違いを見つけた）── 同日、同じ実測から。
      明細の合計と小計が合わない冊で、**食い違いを見つけているのに区分は「単」**
      という記録が出ていた。理由の文字列にだけ書いて、導出は `evidences` しか
      見ていなかった（片配線）。
      → 食い違いは散文ではなく**導出への入力**にする。ここに立てば必ず `割` になる。
    """
    sources = _distinct_sources(evidences)
    if conflict:
        # ★ 上部と帯が一致していても、明細と合わないなら値は出せない。
        #   「一致した 2 つ」を根拠に値を出すと、壊れた冊を黙って通す。
        return SPLIT if sources else NONE_FOUND
    if not sources:
        return NONE_FOUND
    values = {_key(v) for v in sources.values()}
    if len(sources) == 1:
        return SINGLE
    if len(values) != 1:
        return SPLIT
    return CONFIRMED if swept else SINGLE


def grade(rec: Record) -> str:
    """記録の区分。★ 導出は `grade_of` 一本（ここでは何も判断しない）。"""
    return grade_of(rec.evidences, rec.swept, rec.conflict)


def value(rec: Record):
    """出す値。★ 割 と 無 は **None**（空欄・凍結した判断）。"""
    g = grade(rec)
    if g not in GRADES_WITH_VALUE:
        return None
    sources = _distinct_sources(rec.evidences)
    return next(iter(sources.values()))


def _key(v):
    """値の同一性。★ 数は丸め、文字は前後の空白を落として比べる。"""
    if isinstance(v, bool):
        return v
    if isinstance(v, (int, float)):
        return round(float(v), 2)
    return str(v).strip()


def both_sides(rec: Record) -> tuple:
    """割 のときに人へ見せる「両側の数字」。★ 買い手の信用条件④。

    戻り値: ((番地, 値, 読み方), …) ── 出所ごとに 1 件
    """
    if grade(rec) != SPLIT:
        return ()
    seen, out = set(), []
    for e in rec.evidences:
        if e.at in seen:
            continue
        seen.add(e.at)
        out.append((e.at, e.value, e.how))
    return tuple(out)


def describe(rec: Record) -> str:
    """人に見せる 1 行。★ 区分ごとに形が決まる（出口ごとに書き分けない）。"""
    g = grade(rec)
    if g == CONFIRMED:
        srcs = _distinct_sources(rec.evidences)
        tail = f"・{rec.swept_how}" if rec.swept_how else ""
        return f"{len(srcs)} つの根拠が一致（{'／'.join(sorted(srcs))}）{tail}"
    if g == SINGLE:
        srcs = _distinct_sources(rec.evidences)
        if len(srcs) > 1:
            # ★ 一致はしているが掃き出していない ── そう言う（黙って 確 にしない）
            return (f"{len(srcs)} つの根拠が一致（{'／'.join(sorted(srcs))}）── ただし"
                    "他に食い違う数字が無いかは確かめられていません")
        e = rec.evidences[0]
        return f"根拠は 1 つだけ（{e.at}・{e.how}）── 裏が取れていません"
    if g == SPLIT:
        parts = "／".join(f"{at}={v}" for at, v, _ in both_sides(rec))
        return f"根拠が食い違いました（{parts}）── 値は出しません: {rec.blank_reason}"
    return f"根拠が見つかりません: {rec.blank_reason}"
